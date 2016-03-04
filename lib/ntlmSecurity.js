"use strict";

var _ = require('lodash');

function ntlm(username, password, ignoreSSL) {
    this.defaults = {
        username: username,
        password: password
    };

    // ignore SSL?
    if(ignoreSSL) {
      _.merge(this.defaults, GLOBAL.IGNORESSLOPTS);
    }
}

ntlm.prototype.addHeaders = function(headers) {

};

ntlm.prototype.toXML = function() {
    return '';
};

ntlm.prototype.addOptions = function(options) {
    _.merge(options, this.defaults);
};

module.exports = ntlm;
