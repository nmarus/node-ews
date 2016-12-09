"use strict";

var request = require('request');
var when = require('when');
var ntlm = require('httpntlm');
var soap = require('soap');
var _ = require('lodash');

var path = require('path');
var tmp = require('tmp');
var fs = require('fs');

var NtlmSecurity = require('./ntlmSecurity');
var HttpClient = require('./http');

function EWS(config, options) {

  // config valid?
  var configIsValid = (config && config.username && config.password && config.host);

  // validate options
  options = typeof options === 'object' ? options : {};

  // if required config found...
  if(configIsValid) {

    // ntlm auth
    var ntlmAuth = {
      wsdlOptions: { httpClient: HttpClient },
      authProfile: new NtlmSecurity(config.username, config.password, options),
      getFile: function(url, filePath) {
        var ntlmOptions = { 'username': config.username, 'password': config.password };
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

    // basic auth
    var basicAuth = {
      wsdlOptions: {},
      authProfile: new soap.BasicAuthSecurity(config.username, config.password, options),
      getFile: function(url, filePath) {
        // request options
        var requestOptions = { 'auth': { 'user': config.username, 'pass': config.password, 'sendImmediately': false } };
        requestOptions = _.merge(requestOptions, _.clone(options));
        requestOptions.url = url;

        return when.promise((resolve, reject) => {
          request(requestOptions, (err, res, body) => {
            if(err) reject(err);
            else if(res.statusCode == 401) reject(new Error('Basic Auth StatusCode 401: Unauthorized.'));
            else fs.writeFile(filePath, body, function(err) {
              if(err) reject(err);
              else resolve(filePath);
            });
          });
        });
      }
    };

    // if auth defined in options
    if(config && typeof config.auth === 'string' && _.toLower(config.auth) === 'basic') {
      this.auth = basicAuth;
    } else {
      this.auth = ntlmAuth;
    }

    // if temp is defined in options
    if(config && typeof config.temp === 'string') {
      this.tempDir = config.temp;
    } else {
      this.tempDir;
    }

    // build ews url
    this.urlApi = config.host + '/EWS/Exchange.asmx';

    // build ews soap urls
    this.urlServices = config.host + '/ews/services.wsdl';
    this.urlMessages = config.host + '/ews/messages.xsd';
    this.urlTypes = config.host + '/ews/types.xsd';
  }

  // else, required config missing so throw error
  else {
    throw new Error('missing required config parameters');
  }
}

// check for pre-exisiting ews soap wsdl and xsd files
EWS.prototype.init = function() {

  var urlServices = this.urlServices;
  var urlMessages = this.urlMessages;
  var urlTypes = this.urlTypes;
  var urlApi = this.urlApi;
  var getFile = this.auth.getFile;

  // apply fix to wsdl at file path
  function fixWsdl(filePath) {
    // wsdl fix
    var wsdlFix = '\n<wsdl:service name="ExchangeServices">\n<wsdl:port name="ExchangeServicePort" binding="tns:ExchangeServiceBinding">\n<soap:address location="' + urlApi + '"/>\n</wsdl:port>\n</wsdl:service>\n</wsdl:definitions>';
    var alreadyFixed = '<wsdl:service name="ExchangeServices">'
    // fix ms wdsl...
    // https://msdn.microsoft.com/en-us/library/office/aa580675.aspx
    // The EWS WSDL file, services.wsdl, does not fully conform to the WSDL standard
    // because it does not include a WSDL service definition. This is because EWS is
    // not designed to be hosted on a computer that has a predefined address.
    return when.promise((resolve, reject) => {
      fs.readFile(filePath, 'utf8', function(err, wsdl) {
        if(err) reject(err);

        // error on malformed wsdl file
        else if(wsdl.search('<wsdl:definitions') == -1) reject(new Error('Invalid or malformed wsdl file: ' + filePath));

        else {
          if (wsdl.search(alreadyFixed)==-1) { 
          // Remove </wsdl:definitions> and replace with our fix and
          // write file back to disk
          fs.writeFile(filePath, wsdl.replace("</wsdl:definitions>",wsdlFix), function(err) {
            if(err) reject(err);
            else resolve(filePath);
          });
          } else {
             resolve(filePath);
          }   
        }
      });
    });
  }

  // check for existing files in tempDir, download if not found
  function initFiles(tempDir) {

    // if file exists...
    function ifFileExists(filePath) {
      if(!filePath) return when.reject(new Error('File not specified'));

      else return when.promise((resolve, reject) => {
        fs.stat(filePath, (err, stats) => {
          if(err) reject(new Error('File does not exist.'));
          else if(stats.isFile()) resolve(filePath);
          else reject(new Error('Invalid file at' + filePath));
        })
      });
    }

    // if dir exists
    function ifDirExists(dirPath) {
      if(!dirPath) return when.reject(new Error('Directory not specified'));

      else return when.promise((resolve, reject) => {
        fs.stat(dirPath, (err, stats) => {
          if(err) reject(new Error('Directory does not exist.'));
          else if(stats.isDirectory()) resolve(dirPath);
          else reject(new Error('Invalid directory at ' + dirPath));
        })
      });
    }

    return ifDirExists(tempDir)

      // verify messages.wsd
      .then(() => {
        var messageswsdPath = path.join(tempDir, path.basename(urlMessages));
        return ifFileExists(messageswsdPath)
          .catch(() => getFile(urlMessages, messageswsdPath));
      })

      // verify types.wsd
      .then(() => {
        var typeswsdPath = path.join(tempDir, path.basename(urlTypes));
        return ifFileExists(typeswsdPath)
          .catch(() => getFile(urlTypes, typeswsdPath));
      })

      // verify services.wsdl
      .then(() => {
        var serviceswsdlPath = path.join(tempDir, path.basename(urlServices));
        return ifFileExists(serviceswsdlPath)
          .catch(() => getFile(urlServices, serviceswsdlPath));
      })

      // fix wsdl
      .then(serviceswsdlPath => fixWsdl(serviceswsdlPath));
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

  return initFiles(this.tempDir)
    .catch(() => {
      console.log('EWS temp directory not sepcified or does not exist. Using system temp directory instead.');

      return getTempDir()
        .then(temp => {
          this.tempDir = temp;
          return initFiles(temp);
        });
    });
}

// run ews function with args
EWS.prototype.run = function(ewsFunction, ewsArgs, soapHeader) {

  var authProfile = this.auth.authProfile;
  var wsdlOptions = this.auth.wsdlOptions;

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

          // add optional soap header
          if(typeof soapHeader === 'object') {
            client.addSoapHeader(soapHeader);
          }

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
