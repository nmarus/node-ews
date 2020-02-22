'use strict';

var ntlm = require('ntlm-client');
var _ = require('lodash');

module.exports = {

  request: function(rurl, data, callback, exheaders, exoptions) {
    // parse args
    var args = Array.prototype.slice.call(arguments);
    rurl = args.shift();
    data = typeof args[0] === "string" ? args.shift() : null;
    callback = args.shift();
    // optional args
    exheaders = args.length > 0 ? args.shift() : null;
    exoptions = args.length > 0 ? args.shift() : null;

    var method = data ? "post" : "get";
    var headers = {};
    if (typeof data === 'string') {
      headers["Content-Type"] = "text/xml;charset=UTF-8";
    }

    var options = {
      url: rurl,
      headers: headers,
      allowRedirects: true
    };

    options.body = data ? data : null;

    _.merge(options, exoptions);
    _.merge(options.headers, exheaders);

    const promise = ntlm.request({
      uri: options.url,
      method: method,
      username: options.username,
      password: options.password,
      request: options
    });

    promise.then((res) => {
      var body = res.body;
      if (typeof body === "string") {
        // Remove any extra characters that appear before or after the SOAP
        // envelope.
        var match = body.match(/(?:<\?[^?]*\?>[\s]*)?<([^:]*):Envelope([\S\s]*)<\/\1:Envelope>/i);
        if (match) {
          body = match[0];
        }
      }
      callback(null, res, body);
    }, (err) => {
      callback(err);
    })

    // build request and return to node-soap
    var req = {};
    req.headers = options.headers;
    req.body = options.body;
    return req;
  }

};