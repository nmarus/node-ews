'use strict';

const when = require('when');
const ntlm = require('httpntlm');
const _ = require('lodash');

const fs = require('fs');

const NtlmSecurity = require('./ntlm/ntlmSecurity');
const HttpClient = require('./ntlm/http');

// define ntlm auth
const NTLMAuth = function(config, options) {
  const passwordIsPlainText = _.has(config, 'password');
  const passwordIsEncrypted = _.has(config, 'nt_password') && _.has(config, 'lm_password');

  if(typeof config === 'object'
    && _.has(config, 'host')
    && _.has(config, 'username')
    && (passwordIsPlainText || passwordIsEncrypted)
  ) {
    return {
      wsdlOptions: { httpClient: HttpClient },
      authProfile: new NtlmSecurity(config, options),
      getUrl: function(url, filePath) {
        let ntlmOptions = { 'username': config.username };
        if (passwordIsPlainText) {
          ntlmOptions.password = config.password;
        }
        else {
          ntlmOptions.nt_password = config.nt_password;
          ntlmOptions.lm_password = config.lm_password;
        }
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
