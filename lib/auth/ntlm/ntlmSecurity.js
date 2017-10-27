'use strict';

var _ = require('lodash');

function ntlm(config, defaults) {
    this.defaults = {
      username: config.username
    };
    if (config.password) {
        this.defaults.password = config.password;
    }
    else {
        this.defaults.nt_password = config.nt_password;
        this.defaults.lm_password = config.lm_password;
    }
    _.merge(this.defaults, defaults);
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
