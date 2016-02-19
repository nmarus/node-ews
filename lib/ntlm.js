"use strict";

var _ = require('lodash');


function ntlm(username, password) {
    this.defaults = {
        username: username,
        password: password
    };
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
