"use strict";

const request = require('request');
const debug = require('debug')('ews');
const when = require('when');
const ntlm = require('httpntlm');
const soap = require('soap');
const _ = require('lodash');

const path = require('path');
const tmp = require('tmp');
const fs = require('fs');

const NtlmSecurity = require('./ntlmSecurity');
const HttpClient = require('./http');

function EWS(config, options) {
  // validate options
  options = typeof options === 'object' ? options : {};

  // define class vars
  this.auth;
  this.tempDir;

  function configIsValid(config) {
    if (!config) {
      return false;
    }

    if (config.auth === 'bearer') {
      return config.username && config.host && config.token;
    } else {
      // nltm and basic auth
      return config.username && config.password && config.host;
    }
  }

  // if required config found...
  if(configIsValid(config)) {

    // ntlm auth
    var ntlmAuth = {
      wsdlOptions: { httpClient: HttpClient },
      authProfile: new NtlmSecurity(config.username, config.password, options),
      getUrl: function(url, filePath) {
        var ntlmOptions = { 'username': config.username, 'password': config.password };
        ntlmOptions = _.merge(ntlmOptions, _.clone(options));
        ntlmOptions.url = url;

        return when.promise((resolve, reject) => {
          ntlm.get(ntlmOptions, function(err, res) {
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
      getUrl: function(url, filePath) {
        // request options
        var requestOptions = { 'auth': { 'user': config.username, 'pass': config.password, 'sendImmediately': false } };
        requestOptions = _.merge(requestOptions, _.clone(options));
        requestOptions.url = url;

        return when.promise((resolve, reject) => {
          request(requestOptions, function(err, res, body) {
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

    // bearer auth
    var bearerAuth = {
      wsdlOptions: {},
      authProfile: new soap.BearerSecurity(config.token, options),
      getUrl: function(url, filePath) {
        // request options
        var requestOptions = {
            auth: {
                bearer: config.token
            }
        };
        requestOptions = _.merge(requestOptions, _.clone(options));
        requestOptions.url = url;

        return when.promise((resolve, reject) => {
          request(requestOptions, function(err, res, body) {
            if(err) reject(err);
            else if(res.statusCode == 401) reject(new Error('Bearer Auth StatusCode 401: Unauthorized.'));
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
    } else if(config && typeof config.auth === 'string' && _.toLower(config.auth) === 'bearer') {
      this.auth = bearerAuth;
    } else {
      this.auth = ntlmAuth;
    }

    // if temp is defined in options
    if(config && typeof config.temp === 'string') {
      this.tempDir = config.temp;
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

  var ews = this;

  // if file exists, return fullfilled promise with filePath, else reject
  function ifFileExists(filePath) {
    if(!filePath) return when.reject(new Error('File not specified'));

    else return when.promise((resolve, reject) => {
      fs.stat(filePath, function(err, stats) {
        if(err) reject(new Error('File does not exist.'));
        else if(stats.isFile()) resolve(filePath);
        else reject(new Error('Invalid file at' + filePath));
      })
    });
  }

  // if dir exists, return fullfilled promise with dirPath, else reject
  function ifDirExists(dirPath) {
    if(!dirPath) return when.reject(new Error('Directory not specified'));

    else return when.promise((resolve, reject) => {
      fs.stat(dirPath, function(err, stats) {
        if(err) reject(new Error('Directory does not exist.'));
        else if(stats.isDirectory()) resolve(dirPath);
        else reject(new Error('Invalid directory at ' + dirPath));
      })
    });
  }

  // apply fix to wsdl at file path
  function fixWsdl(filePath) {

    // wsdl service definition
    var wsdlService = '<wsdl:service name="ExchangeServices"><wsdl:port name="ExchangeServicePort" binding="tns:ExchangeServiceBinding"><soap:address location="' + ews.urlApi + '"/></wsdl:port></wsdl:service>';

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
          var isMalformed = (wsdl.search('<wsdl:definitions') == -1);

          // wsdl already fixed
          var isFixed = (wsdl.search('<wsdl:service name="ExchangeServices">') >= 0);

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
  }

  // check for existing files in tempDir, download if not found
  function initFiles(tempDir) {

    return ifDirExists(tempDir)

      // verify messages.xsd
      .then(() => {
        debug('validating messages.xsd...');
        var messageswsdPath = path.join(tempDir, path.basename(ews.urlMessages));
        return ifFileExists(messageswsdPath)
          .catch(() => ews.auth.getUrl(ews.urlMessages, messageswsdPath));
      })

      // verify types.xsd
      .then(() => {
        debug('validating types.xsd...');
        var typeswsdPath = path.join(tempDir, path.basename(ews.urlTypes));
        return ifFileExists(typeswsdPath)
          .catch(() => ews.auth.getUrl(ews.urlTypes, typeswsdPath));
      })

      // verify services.wsdl
      .then(() => {
        debug('validating services.wsdl...');
        var serviceswsdlPath = path.join(tempDir, path.basename(ews.urlServices));
        return ifFileExists(serviceswsdlPath)
          .catch(() => ews.auth.getUrl(ews.urlServices, serviceswsdlPath));
      })

      // fix wsdl
      .then(serviceswsdlPath => fixWsdl(serviceswsdlPath));
  }

  // establish a temporary directory
  function getTempDir() {
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
  }

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
EWS.prototype.notificationService = function(options,callback) {
	
	if(typeof options == "undefined" || typeof callback == "undefined") throw new Error('missing required parameter(s)');
	if(typeof options.xml == "undefined" || options.xml == '') throw new Error('options.xml is missing');
	
	// merge user specified options with defaults
	options = _.merge(this[this.auth].wsdlOptions, {port:8000, path:"/notification"}, options)
	
	// define the services
	options.services = {
		NotificationServices:{
			NotificationServicePort:{
				SendNotification:callback
			}
		}
	};

	
	return this.init()
	.then(wsdlFilePath => {
		return when.promise((resolve, reject) => {
			// for the includes specified in the XML, the uri must be specified and set to the same location as the main EWS services.wsdl
			options.uri = wsdlFilePath;
			// replace the default blank location with the location of the EWS server uri
			options.xml = options.xml.replace('soap:address location=""','soap:address location="'+ this.urlApi + '"');
			// create the basic http server
			var server =  require('http').createServer(function(request,response) {
				response.end("404: Not Found: " + request.url);
			});
			// start the server
			server.listen(options.port);
			// resolve the service with soap.listen
			resolve(soap.listen(server, options));
		})
	});
};

module.exports = EWS;
