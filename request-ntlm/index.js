var _       = require('lodash');
var async   = require('async');
var request = require('request');
var ntlm    = require('./lib/ntlm');

var makeRequest = function(method, options, params, callback) {
  // parse args
  var args = Array.prototype.slice.call(arguments);
  method = args.shift();
  options = args.shift();
  callback = args.pop();
  // optional args
  params = args.length > 0 ? args.shift() : null;

  var KeepAlive = require('agentkeepalive');
  if (options.url.toLowerCase().indexOf('https://') === 0) {
    KeepAlive = KeepAlive.HttpsAgent;
  }

  var keepaliveAgent = new KeepAlive();

  if (!options.workstation) options.workstation = '';
  if (!options.ntlm_domain) options.ntlm_domain = '';
  if (!options.headers    ) options.headers     = {};

  options.ntlm = options.ntlm || {};
  options.ntlm.strict = Boolean(options.ntlm.strict);

  function startAuth($) {
    var type1msg = ntlm.createType1Message(options);
    options.method = method;
    _.extend(options.headers, {
      'Connection': 'keep-alive',
      'Authorization': type1msg
    });
    options.agent = keepaliveAgent;
    request(options, $);
  }

  function requestComplete(res, body, $) {
    if (!res.headers['www-authenticate']) {
      return options.ntlm.strict
          ? $(new Error('www-authenticate not found on response of second request'))
          : $(null, res, body);
    }

    var type2msg = ntlm.parseType2Message(res.headers['www-authenticate']);
    var type3msg = ntlm.createType3Message(type2msg, options);
    options.method = method;
    _.extend(options.headers, {
      'Connection': 'keep-alive',
      'Authorization': type3msg
    });

    options.agent = keepaliveAgent;

    if(params && typeof params == "string") {
      options.body = params;
    } else if(params) { 
      options.json = params;
    }

    request(options, $);
  }

  async.waterfall([startAuth, requestComplete], callback);
};

exports.get   = _.partial(makeRequest, 'get');
exports.post  = _.partial(makeRequest, 'post');
exports.put   = _.partial(makeRequest, 'put');
exports.delete= _.partial(makeRequest, 'delete');
