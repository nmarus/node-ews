
var should = require('should');
var request = require('../index.js');

describe('request-ntlm-continued', function(){

    function simpleGetRequest(options) {
        return function(callback) {
            request.get(options, undefined, function (err, res, body) {
                should.not.exist(err);
                should.exist(res);
                should.exist(body);
                callback();
            });
        }
    }

    this.timeout(10000);

    var executeRequestAgainsNtlmServer = undefined;
    try {
        var options = require(__dirname + '/ntlm-options');
        executeRequestAgainsNtlmServer = simpleGetRequest(options);
    }
    catch(err) {
    }

    it('successfully execute a request against an NTLM server', executeRequestAgainsNtlmServer);

    describe("non-NTLM requests", function(){

        var options;
        beforeEach(function(){
            options = {
                username    : 'username',
                password    : 'password',
                ntlm_domain : 'yourdomain',
                workstation : 'workstation',
                url         : 'https://www.google.com/search?q=ntlm',
                strictSSL   : false
            };
        });

        it('successfully execute a request against a non-NTLM server', function(done){
            simpleGetRequest(options)(done);
        });

        it('fails when requesting non-NTLM resource with `options.ntlm.strict`', function(done){
            options.ntlm = { strict:true };
            request.get(options, undefined, function (err) {
                should.exist(err);
                done();
            });
        });
    });
});
