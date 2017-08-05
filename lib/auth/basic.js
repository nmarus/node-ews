'use strict';

const request = require('request');
const when = require('when');
const soap = require('soap');
const _ = require('lodash');

const fs = require('fs');

// define basic auth
const BasicAuth = function(config, options) {
  if(typeof config === 'object'
    && _.has(config, 'host')
    && _.has(config, 'username')
    && _.has(config, 'password')
  ) {
    return {
      wsdlOptions: {},
      authProfile: new soap.BasicAuthSecurity(config.username, config.password, options),
      getUrl: function(url, filePath) {
        // request options
        let requestOptions = { 'auth': { 'user': config.username, 'pass': config.password, 'sendImmediately': false } };
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
  } else {
    throw new Error('missing required config parameters');
  }
};

module.exports = BasicAuth;
