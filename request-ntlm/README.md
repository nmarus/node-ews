# Request-NTLM

Module for authenticating with NTLM; An ntlm authentication wrapper for the Request module.

## Install with NPM

```
$ npm install --save-dev request-ntlm-continued
```

## Usage

```javascript
var ntlm = require('request-ntlm-continued');

var opts = {
  username: 'username',
  password: 'password',
  ntlm_domain: 'yourdomain',
  workstation: 'workstation',
  url: 'http://example.com/path/to/resource'
};
var json = {
  // whatever object you want to submit
};
ntlm.post(opts, json, function(err, response) {
  // do something
});
```

## Changes from original:

* don't assume the post body is an object and should be made into json
* options.domain is in use by request. Use ntlm_domain instead
* ability to set custom headers
* ability to use http and not only https
* gracefully complete the request if the server doesn't actually require NTLM.
  Fail only if `options.ntlm.strict` is set to `true` (default=`false`).
