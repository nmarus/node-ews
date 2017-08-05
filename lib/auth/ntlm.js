'use strict';

const when = require('when');
const soap = require('soap');
const _ = require('lodash');

const NtlmSecurity = require('./ntlm/ntlmSecurity');
const HttpClient = require('./ntlm/http');

// define ntlm auth
const NTLMAuth = function(config, options) {
  if(typeof config === 'object'
    && _.has(config, 'host')
    && _.has(config, 'username')
    && _.has(config, 'password')
  ) {
    return {
      wsdlOptions: { httpClient: HttpClient },
      authProfile: new NtlmSecurity(config.username, config.password, options),
      getUrl: function(url, filePath) {
        let ntlmOptions = { 'username': config.username, 'password': ews.config.password };
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
  } else {
    throw new Error('missing required config parameters');
  }
};

module.exports = NTLMAuth;
