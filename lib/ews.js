var when = require('when');
var ntlm = require('httpntlm');
var soap = require('soap');
var _ = require('lodash');

var path = require('path');
var tmp = require('tmp');
var fs = require('fs');

var NtlmSecurity = require('./ntlmSecurity');
var HttpClient = require('./http');

function EWS(username, password, host, options) {

  if(!username || !password || !host) {
    throw new Error('missing required parameters');
  }

  options = options || {};

  // ews url
  this.urlApi = host + '/EWS/Exchange.asmx';

  // ews soap urls
  this.urlServices = host + '/ews/services.wsdl';
  this.urlMessages = host + '/ews/messages.xsd';
  this.urlTypes = host + '/ews/types.xsd';

  // temp dir
  this.tempDir;

  // ntlm auth
  this.ntlm = {
    wsdlOptions: { httpClient: HttpClient },
    authProfile: new NtlmSecurity(username, password, options),
    getFile: function(url, filePath) {
      var ntlmOptions = { username: username, password: password };
      ntlmOptions = _.merge(ntlmOptions, _.clone(options));
      ntlmOptions.url = url;

      return when.promise((resolve, reject) => {
        ntlm.get(ntlmOptions, function (err, res) {
          if(err) reject(err);
          else if(res.statusCode == 401) reject(new Error('NTLM StatusCode 401: Unauthorized.'));
          else fs.writeFile(filePath, res.body, function(err) {
            if(err) reject(err);
            else resolve(filePath);
          });
        });
      });
    }
  };

  // set default auth
  this.auth = 'ntlm';
}

// check for pre-exisiting ews soap wsdl and xsd files
EWS.prototype.init = function() {

  var urlServices = this.urlServices;
  var urlMessages = this.urlMessages;
  var urlTypes = this.urlTypes;
  var urlApi = this.urlApi;
  var getFile = this[this.auth].getFile;

  // if file exists...
  function ifFileExists(filePath) {
    if(!filePath) when.reject(new Error('File not specified'));

    else return when.promise((resolve, reject) => {
      fs.stat(filePath, (err, stats) => {
        if(err) reject(err);
        else if(stats.isFile()) resolve(filePath);
        else reject(new Error('Invalid File: ' + filePath));
      })
    });
  }

  // if dir exists...
  function ifDirExists(dirPath) {
    if(!dirPath) return when.reject(new Error('Directory not specified'));

    return when(dirPath)
      .then(dirPath => when.promise((resolve, reject) => {
        fs.stat(dirPath, (err, stats) => {
          if(err) reject(err);
          else if(stats.isDirectory()) resolve(dirPath);
          else reject(new Error('Invalid Directory: ' + dirPath));
        })
      }));
  }

  // apply fix to wsdl at file path
  function fixWsdl(filePath) {
    // wsdl fix
    var wsdlFix = '\n<wsdl:service name="ExchangeServices">\n<wsdl:port name="ExchangeServicePort" binding="tns:ExchangeServiceBinding">\n<soap:address location="' + urlApi + '"/>\n</wsdl:port>\n</wsdl:service>\n</wsdl:definitions>';

    // fix ms wdsl...
    // https://msdn.microsoft.com/en-us/library/office/aa580675.aspx
    // The EWS WSDL file, services.wsdl, does not fully conform to the WSDL standard
    // because it does not include a WSDL service definition. This is because EWS is
    // not designed to be hosted on a computer that has a predefined address.
    return when.promise((resolve, reject) => {
      fs.readFile(filePath, 'utf8', function(err, wsdl) {
        if(err) reject(err);

        // error on malformed wsdl file
        else if(wsdl.search('<wsdl:definitions') == -1) reject(new Error('Invalid or malformed wsdl file: ' + file));

        else {
          // break file into an array of lines
          var wsdlLines = wsdl.split('\n');

          // remove last 2 elements of array
          wsdlLines.splice(-2,2);

          // join the array back into a single string and apply fix
          var newWsdl = wsdlLines.join('\n') + wsdlFix;

          // write file back to disk
          fs.writeFile(filePath, newWsdl, function(err) {
            if(err) reject(err);
            else resolve(filePath);
          });
        }
      });
    });
  }

  // check for existing files in tempDir, download if not found
  function initFiles(tempDir) {

    // normalize tempDir
    if(!path.isAbsolute(tempDir)) {
      tempDir = path.normalize(path.join(__dirname, '../../..', tempDir));
    } else {
      tempDir = path.normalize(tempDir);
    }

    return when(tempDir)

      // verify messages.wsd
      .then(() => {
        var filePath = path.join(tempDir, path.basename(urlMessages));
        return ifFileExists(filePath)
          .catch(() => getFile(urlMessages, filePath));
      })

      // verify types.wsd
      .then(() => {
        var filePath = path.join(tempDir, path.basename(urlTypes));
        return ifFileExists(filePath)
          .catch(() => getFile(urlTypes, filePath));
      })

      // verify services.wsdl
      .then(() => {
        var filePath = path.join(tempDir, path.basename(urlServices));
        return ifFileExists(filePath)
          .catch(() => {
            return getFile(urlServices, filePath)
              .then(wsdlFilePath => fixWsdl(wsdlFilePath));
          });
      });
  }

  // establish a temporary directory
  function getTempDir() {
    return when.promise((resolve, reject) => {
      tmp.dir({unsafeCleanup: false}, (err, temp) => {
        if(err) {
          reject(err);
        } else {
          resolve(temp);
        }
      });
    });
  }

  return ifDirExists(this.tempDir)
    .then(tempDir => initFiles(tempDir))
    .catch(() => {
      return getTempDir()
        .then(temp => {
          this.tempDir = temp;
          return initFiles(temp)
        });
    });
}

// run ews function with args
EWS.prototype.run = function(ewsFunction, ewsArgs) {

  var authProfile = this[this.auth].authProfile;
  var wsdlOptions = this[this.auth].wsdlOptions;

  if(typeof ewsFunction !== 'string' || typeof ewsArgs !== 'object') {
    throw new Error('missing required parameters');
  }

  return this.init()
    .then(wsdlFilePath => {
      return when.promise((resolve, reject) => {

      //create soap client
      soap.createClient(wsdlFilePath, wsdlOptions, (err, client) => {
        if(err) reject(err);
        else {
          // validate ewsFunction
          var validEwsFunctions = _.keys(client.describe().ExchangeServices.ExchangeServicePort);
          if(!_.includes(validEwsFunctions, ewsFunction)){
            reject(new Error('ewsFunction not found in WSDL: ' + ewsFunction));
            return;
          }

          // define security plugin
          client.setSecurity(authProfile);

          // run ews soap function
          client[ewsFunction](ewsArgs, (err, result) => {
            if(err) reject(err);
            else resolve(result);
          });
        }
      });

    })
  });
};

module.exports = EWS;
