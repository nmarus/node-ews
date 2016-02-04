var parseString = require('xml2js').parseString;
var processors = require('xml2js').processors;
var httpntlm = require('httpntlm');
var debug = require('debug')('node-ews');
var async = require('async');
var soap = require('../node-soap');
var path = require('path');
var tmp = require('tmp');
var fs = require('fs');
var _ = require('lodash');

function EWS() {

}

// ntlm authorization
EWS.prototype.auth = function(username, password, host) {

  var $this = this;

  if(!username || !password || !host) {
    throw new Error('missing required parameters');
  }

  $this.username = username;
  $this.password = password;
  $this.host = host;

  // ews url
  $this.ewsApi = 'https://' + this.host + '/EWS/Exchange.asmx';

  // ews soap urls
  $this.ewsSoap = [
    { url: 'https://' + $this.host + '/ews/services.wsdl' },
    { url: 'https://' + $this.host + '/ews/messages.xsd' },
    { url: 'https://' + $this.host + '/ews/types.xsd' }
  ];

  // temp dir
  $this.tempDir;

};

// check for pre-exisiting ews soap wsdl and xsd files
EWS.prototype._validate = function(callback) {

  var $this = this;

  // wsdl fix
  var wsdlFix = '\n<wsdl:service name="ExchangeServices">\n<wsdl:port name="ExchangeServicePort" binding="tns:ExchangeServiceBinding">\n<soap:address location="' + $this.ewsApi + '"/>\n</wsdl:port>\n</wsdl:service>\n</wsdl:definitions>';

  // get url with auth and save
  function getFile(url, file, callback) {
    httpntlm.get({ url: url, username: $this.username, password: $this.password },
    function(err, res) {
      if(err) {
        callback(err);
      } else {
        fs.writeFile(file, res.body, function(err) {
          callback(err);
        });
      }
    });
  }

  // check tempdir, wsdl, and xsd files
  if($this.tempDir && $this.ewsSoap[0].path && $this.ewsSoap[1].path && $this.ewsSoap[2].path) {
    return callback(null);
  } else {
    async.series([
      function(cb) {
        //
        // create temp directory
        //
        tmp.dir({unsafeCleanup: true}, function(err, temp) {
          if(err) {
            cb(err);
          } else {
            $this.tempDir = temp;
            cb(err);
          }
        });
      },
      function(cb) {
        //
        // download wdsl and xsd file(s)
        //
        async.each($this.ewsSoap, function(ewsFile, eaCb) {
          ewsFile.path = path.join($this.tempDir,path.basename(ewsFile.url));
          getFile(ewsFile.url, ewsFile.path, function(err) {
            eaCb(err);
          });
        },
        function(err) {
          cb(err);
        });
      },
      function(cb) {
        //
        // fix ms wdsl...
        //
        // https://msdn.microsoft.com/en-us/library/office/aa580675.aspx
        // The EWS WSDL file, services.wsdl, does not fully conform to the WSDL standard
        // because it does not include a WSDL service definition. This is because EWS is
        // not designed to be hosted on a computer that has a predefined address.
        fs.readFile($this.ewsSoap[0].path, 'utf8', function(err, wsdl) {
          if(err) {
            cb(err);
          } else {
            // break file into an array of lines
            var wsdlLines = wsdl.split('\n');
            // remove last 2 elements of array
            wsdlLines.splice(-2,2);
            // join the array back into a single string and apply fix
            var newWsdl = wsdlLines.join('\n') + wsdlFix;
            // write file back to disk
            fs.writeFile($this.ewsSoap[0].path, newWsdl, function(err) {
              cb(err);
            });
          }
        });
      },
    ],function(err) {
      return callback(err);
    });

  }
}

// run ews function with args
EWS.prototype.run = function(ewsFunction, ewsArgs, callback) {

  var $this = this;

  if(!$this.username || !$this.password || !$this.host || !$this.ewsApi || !$this.ewsSoap) {
    throw new Error('missing required parameters');
  }

  //
  // check for presence of wsdl and xsd files
  //
  $this._validate(function(err) {
    if(err) {
      return callback(err, null);
    } else {
      //
      // create soap client, query, and parse results to json
      //
      soap.createClient($this.ewsSoap[0].path, function(err, client) {
        if(err) {
          return callback(err, null);
        } else {
          // define security plugin
          client.setSecurity(new soap.NtlmSecurity($this.username, $this.password));

          // run ews soap function
          client[ewsFunction](ewsArgs, function(err, result, body) {
            if(err) {
              return callback(err, null);
            } else {
              // parse body xml
              parseString(body, {
                tagNameProcessors: [processors.stripPrefix],
                attrNameProcessors: [processors.stripPrefix],
                valueProcessors: [processors.stripPrefix],
                attrValueProcessors: [processors.stripPrefix]
              }, function(err, result){
                if(err) {
                  return callback(err, null);
                } else {
                  return callback(null, result);
                }
              });
            }
          });
        }
      });
    }
  });

};

module.exports = new EWS;
