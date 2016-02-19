"use strict";

var httpntlm = require('httpntlm');
var _ = require('lodash');

module.exports = {

  request: function(rurl, data, callback, exheaders, exoptions) {

    var method = data ? "post" : "get";
    var headers = {};
    if (typeof data === 'string') {
      headers["Content-Type"] = "text/xml;charset=UTF-8";
    }

    var options = {
      url: rurl,
      method: method,
      headers: headers
    };
    if(data){
      options.body = data;
    }

    _.merge(options, exoptions);
    _.merge(options.headers, exheaders);


    httpntlm[method](options, function (error, res) {
      if (error) {
        callback(error);
      } else {
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
      }
    });

    var req = {};
    req.headers = headers;
    return req;

  }

};