'use strict';

const debug = require('debug')('ews');
const when = require('when');
const soap = require('soap');
const _ = require('lodash');

const path = require('path');
const tmp = require('tmp');
const fs = require('fs');

const BasicAuth = require('./auth/basic');
const BearerAuth = require('./auth/bearer');
const NTLMAuth = require('./auth/ntlm');

const EWS = function(config, options) {
  let ews = this;

  // validate options
  ews.options = typeof options === 'object' ? options : {};

  // validate required config keys
  if(typeof config === 'object') {
    ews.config = config;

    // validate credentials
    if(_.has(ews.config, 'auth')) {
      // basic
      if(_.toLower(ews.config.auth) === 'basic') {
        ews.auth = new BasicAuth(ews.config, ews.options);
      }

      // bearer
      else if(_.toLower(ews.config.auth) === 'bearer') {
        ews.auth = new BearerAuth(ews.config, ews.options);
      }

      // else assume ntlm
      else {
        ews.auth = new NTLMAuth(ews.config, ews.options);
      }
    }

    // else assume ntlm
    else {
      ews.auth = new NTLMAuth(ews.config, ews.options);
    }
  }

  // else required config missing
  else {
    throw new Error('missing required config parameters');
  }

  // if temp is defined in config
  if(_.has(ews.config, 'temp') && typeof ews.config.temp === 'string') {
    ews.tempDir = ews.config.temp;
  } else {
    ews.tempDir = false;
  }

  // build ews url
  ews.urlApi = ews.config.host + '/EWS/Exchange.asmx';

  // build ews soap urls
  ews.urlServices = ews.config.host + '/ews/services.wsdl';
  ews.urlMessages = ews.config.host + '/ews/messages.xsd';
  ews.urlTypes = ews.config.host + '/ews/types.xsd';
};

// check for pre-exisiting ews soap wsdl and xsd files
EWS.prototype.init = function() {

  let ews = this;

  // if file exists, return fullfilled promise with filePath, else reject
  let ifFileExists = function(filePath) {
    if(!filePath) return when.reject(new Error('File not specified'));

    else return when.promise((resolve, reject) => {
      fs.stat(filePath, function(err, stats) {
        if(err) reject(new Error('File does not exist.'));
        else if(stats.isFile()) resolve(filePath);
        else reject(new Error('Invalid file at' + filePath));
      })
    });
  };

  // if dir exists, return fullfilled promise with dirPath, else reject
  let ifDirExists = function(dirPath) {
    if(!dirPath) return when.reject(new Error('Directory not specified'));

    else return when.promise((resolve, reject) => {
      fs.stat(dirPath, function(err, stats) {
        if(err) reject(new Error('Directory does not exist.'));
        else if(stats.isDirectory()) resolve(dirPath);
        else reject(new Error('Invalid directory at ' + dirPath));
      })
    });
  };

  // apply fix to wsdl at file path
  let fixWsdl = function(filePath) {

    // wsdl service definition
    let wsdlService = '<wsdl:service name="ExchangeServices"><wsdl:port name="ExchangeServicePort" binding="tns:ExchangeServiceBinding"><soap:address location="' + ews.urlApi + '"/></wsdl:port></wsdl:service>';

    // fix ms wdsl...
    // https://msdn.microsoft.com/en-us/library/office/aa580675.aspx
    // The EWS WSDL file, services.wsdl, does not fully conform to the WSDL standard
    // because it does not include a WSDL service definition. This is because EWS is
    // not designed to be hosted on a computer that has a predefined address.
    return when.promise((resolve, reject) => {
      fs.readFile(filePath, 'utf8', function(err, wsdl) {
        if(err) {
          reject(err);
        } else {
          // wsdl malformed
          let isMalformed = (wsdl.search('<wsdl:definitions') == -1);

          // wsdl already fixed
          let isFixed = (wsdl.search('<wsdl:service name="ExchangeServices">') >= 0);

          // if error on malformed wsdl file
          if(isMalformed) {
            reject(new Error('Invalid or malformed wsdl file: ' + filePath));
          }

          // else, if wsdl fix is already in place
          else if(isFixed) {
            resolve(filePath);
          }

          // else, insert wsdl service
          else {
            debug('inserting service definition into services.wsdl...');
            fs.writeFile(filePath, wsdl.replace('</wsdl:definitions>', '\n' + wsdlService + '\n</wsdl:definitions>'), function(err) {
              if(err) reject(err);
              else resolve(filePath);
            });
          }
        }
      });
    });
  };

  // check for existing files in tempDir, download if not found
  let initFiles = function(tempDir) {

    return ifDirExists(tempDir)

      // verify messages.xsd
      .then(() => {
        debug('validating messages.xsd...');
        let messageswsdPath = path.join(tempDir, path.basename(ews.urlMessages));
        return ifFileExists(messageswsdPath)
          .catch(() => ews.auth.getUrl(ews.urlMessages, messageswsdPath));
      })

      // verify types.xsd
      .then(() => {
        debug('validating types.xsd...');
        let typeswsdPath = path.join(tempDir, path.basename(ews.urlTypes));
        return ifFileExists(typeswsdPath)
          .catch(() => ews.auth.getUrl(ews.urlTypes, typeswsdPath));
      })

      // verify services.wsdl
      .then(() => {
        debug('validating services.wsdl...');
        let serviceswsdlPath = path.join(tempDir, path.basename(ews.urlServices));
        return ifFileExists(serviceswsdlPath)
          .catch(() => ews.auth.getUrl(ews.urlServices, serviceswsdlPath));
      })

      // fix wsdl
      .then(serviceswsdlPath => fixWsdl(serviceswsdlPath));
  };

  // establish a temporary directory
  let getTempDir = function() {
    return when.promise((resolve, reject) => {
      tmp.dir({unsafeCleanup: false}, function(err, temp) {
        if(err) {
          reject(err);
        } else {
          debug('created temp dir at "%s"', temp);
          resolve(temp);
        }
      });
    });
  };

  return initFiles(ews.tempDir)
    .catch(() => {
      debug('temp directory not sepcified or does not exist');

      return getTempDir()
        .then(temp => {
          ews.tempDir = temp;
          return initFiles(ews.tempDir);
        });
    });
}

// run ews function with args
EWS.prototype.run = function(ewsFunction, ewsArgs, soapHeader) {

  let ews = this;

  let authProfile = ews.auth.authProfile;
  let wsdlOptions = ews.auth.wsdlOptions;

  if(typeof ewsFunction !== 'string' || typeof ewsArgs !== 'object') {
    throw new Error('missing required parameters');
  }

  return ews.init()
    .then(wsdlFilePath => {
      return when.promise((resolve, reject) => {

      //create soap client
      soap.createClient(wsdlFilePath, wsdlOptions, (err, client) => {
        if(err) reject(err);
        else {
          // validate ewsFunction
          let validEwsFunctions = _.keys(client.describe().ExchangeServices.ExchangeServicePort);
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
          client[ewsFunction](ewsArgs, function(err, result) {
            if(err) reject(err);
            else resolve(result);
          });
        }
      });

    })
  });
};

// create EWS Notification Service listener
EWS.prototype.notificationService = function(options, callback) {

  let ews = this;

  if(typeof options !== 'object' || typeof callback !== 'function') {
    throw new Error('missing required parameter(s)');
  }

  if(_.has(options, 'xml') && options.xml !== '') {
    // merge user specified options with defaults
    options = _.merge(ews[ews.auth].wsdlOptions, { port: 8000, path:'/notification' }, options);
  } else {
    throw new Error('options.xml is missing');
  }

  // define the services
  options.services = {
    NotificationServices: {
      NotificationServicePort: {
        SendNotification: callback
      }
    }
  };

  return ews.init()
    .then(wsdlFilePath => {
      return when.promise((resolve, reject) => {
        // for the includes specified in the XML, the uri must be specified and set to the same location as the main EWS services.wsdl
        options.uri = wsdlFilePath;
        // replace the default blank location with the location of the EWS server uri
        options.xml = options.xml.replace('soap:address location=""','soap:address location="'+ ews.urlApi + '"');
        // create the basic http server
        var server =  require('http').createServer(function(request, response) {
          response.end('404: Not Found: ' + request.url);
        });
        // start the server
        server.listen(options.port);
        // resolve the service with soap.listen
        resolve(soap.listen(server, options));
      })
    });
};

module.exports = EWS;
