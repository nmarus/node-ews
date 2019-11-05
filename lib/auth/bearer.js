'use strict';

const request = require('request');
const when = require('when');
const soap = require('soap');
const _ = require('lodash');

const fs = require('fs');

// define bearer auth
const BearerAuth = function(config, options) {
  if(typeof config === 'object'
    && _.has(config, 'host')
    && _.has(config, 'username')
    && _.has(config, 'token')
  ) {
    return {
      wsdlOptions: {},
      authProfile: new soap.BearerSecurity(config.token, options),
      getUrl: function(url, filePath) {
        // request options
        let requestOptions = {
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
  } else {
    throw new Error('missing required config parameters');
  }
};

module.exports = BearerAuth;
