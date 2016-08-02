var when = require('when');
var ntlm = require('httpntlm');
var soap = require('soap');
var path = require('path');
var tmp = require('tmp');
var fs = require('fs');
var _ = require('lodash');

var NtlmSecurity = require('./ntlmSecurity');
var HttpClient = require('./http');

function EWS(username, password, host, options) {

  if(!username || !password || !host) {
    throw new Error('missing required parameters');
  }

  this.username = username;
  this.password = password;
  this.host = host;
  this.options = options || {};

  // ews url
  this.urlApi = this.host + '/EWS/Exchange.asmx';

  // ews soap urls
  this.urlServices = this.host + '/ews/services.wsdl';
  this.urlMessages = this.host + '/ews/messages.xsd';
  this.urlTypes = this.host + '/ews/types.xsd';

  // temp dir
  this.tempDir;
}

// check for pre-exisiting ews soap wsdl and xsd files
EWS.prototype.init = function() {

  // wsdl fix
  var wsdlFix = '\n<wsdl:service name="ExchangeServices">\n<wsdl:port name="ExchangeServicePort" binding="tns:ExchangeServiceBinding">\n<soap:address location="' + this.urlApi + '"/>\n</wsdl:port>\n</wsdl:service>\n</wsdl:definitions>';

  // ntlm options
  var ntlmOptions = {
    username: this.username,
    password: this.password
  };
  _.merge(ntlmOptions, this.options);


  // get file via ntlm
  function getFile(url, file) {
    ntlmOptions.url = url;
    return when.promise((resolve, reject) => {
      ntlm.get(ntlmOptions, function (err, res) {
        if(err) reject(err);
        else {
          fs.writeFile(file, res.body, function(err) {
            if(err) reject(err);
            else resolve(file);
          });
        }
      });
    });
  }

  // get temp directory
  var temp = when.promise((resolve, reject) => {
    tmp.dir({unsafeCleanup: true}, (err, temp) => {
      if(err) {
        reject(err);
      } else {
        this.tempDir = temp;
        resolve(temp);
      }
    });
  });


  return when(temp)
    .then(() => getFile(this.urlMessages, path.join(this.tempDir, path.basename(this.urlMessages))))
    .then(() => getFile(this.urlTypes, path.join(this.tempDir, path.basename(this.urlTypes))))
    .then(() => getFile(this.urlServices, path.join(this.tempDir, path.basename(this.urlServices))))
    .then(file => {

      // fix ms wdsl...
      // https://msdn.microsoft.com/en-us/library/office/aa580675.aspx
      // The EWS WSDL file, services.wsdl, does not fully conform to the WSDL standard
      // because it does not include a WSDL service definition. This is because EWS is
      // not designed to be hosted on a computer that has a predefined address.
      return when.promise((resolve, reject) => {
        fs.readFile(file, 'utf8', function(err, wsdl) {
          if(err) reject(err);
          else {
            // break file into an array of lines
            var wsdlLines = wsdl.split('\n');
            // remove last 2 elements of array
            wsdlLines.splice(-2,2);
            // join the array back into a single string and apply fix
            var newWsdl = wsdlLines.join('\n') + wsdlFix;
            // write file back to disk
            fs.writeFile(file, newWsdl, function(err) {
              if(err) reject(err);
              else resolve(file);
            });
          }
        });
      })

    });
}

// run ews function with args
EWS.prototype.run = function(ewsFunction, ewsArgs) {

  if(typeof ewsFunction !== 'string' || typeof ewsArgs !== 'object') {
    throw new Error('missing required parameters');
  }

  return this.init()
    .then(wsdl => when.promise((resolve, reject) => {

      //create soap client
      soap.createClient(wsdl, { httpClient: HttpClient }, (err, client) => {
        if(err) reject(err);
        else {
          // define security plugin
          client.setSecurity(new NtlmSecurity(this.username, this.password, this.options));

          // run ews soap function
          client[ewsFunction](ewsArgs, (err, result) => {
            if(err) reject(err);
            else {
              resolve(result);
            }
          });
        }
      });

    }));
};

module.exports = EWS;
