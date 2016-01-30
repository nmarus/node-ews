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

// constructor
function EWS() {

}

// ntlm authorization
EWS.prototype.auth = function(username, password, ewsHost) {

  var $this = this;

  if(!username || !password || !ewsHost) {
    throw new Error('missing required parameters');
  }

  $this.username = username;
  $this.password = password;
  $this.ewsHost = ewsHost;

  // ews url
  $this.ews = 'https://' + this.ewsHost + '/EWS/Exchange.asmx';

  // ews soap urls
  $this.ewsSoap = [
    { url: 'https://' + $this.ewsHost + '/ews/services.wsdl' },
    { url: 'https://' + $this.ewsHost + '/ews/messages.xsd' },
    { url: 'https://' + $this.ewsHost + '/ews/types.xsd' }
  ];
};

EWS.prototype.get = function(ewsFunction, ewsArgs, callback) {

  var $this = this;

  if(!$this.username || !$this.password || !$this.ewsHost || !$this.ews || !$this.ewsSoap) {
    throw new Error('missing required parameters');
  }

  // temp dir
  var tempDir;

  // wsdl fix
  var wsdlFix = '\n<wsdl:service name="ExchangeServices">\n<wsdl:port name="ExchangeServicePort" binding="tns:ExchangeServiceBinding">\n<soap:address location="' + $this.ews + '"/>\n</wsdl:port>\n</wsdl:service>\n</wsdl:definitions>';

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

  async.series([
    function(cb) {
      //
      // create temp directory
      //
      tmp.dir({unsafeCleanup: true}, function(err, temp) {
        if(err) {
          cb(err);
        } else {
          tempDir = temp;
          cb(err);
        }
      });
    },
    function(cb) {
      //
      // download wdsl and xsd file(s)
      //
      async.each($this.ewsSoap, function(ewsFile, eaCb) {
        ewsFile.path = path.join(tempDir,path.basename(ewsFile.url));
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
      // fix non ms wdsl...
      //
      // https://msdn.microsoft.com/en-us/library/office/aa580675(v=exchg.150).aspx
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
    function(cb) {
      //
      // create soap client
      //
      soap.createClient($this.ewsSoap[0].path, function(err, client) {
        if(err) {
          cb(err);
        } else {
          // define security plugin
          client.setSecurity(new soap.NtlmSecurity($this.username, $this.password));

          // run ews soap function
          client[ewsFunction](ewsArgs, function(err, result, body) {
            if(err) {
              cb(err);
            } else {
              // parse body xml
              parseString(body, {
                tagNameProcessors: [processors.stripPrefix],
                attrNameProcessors: [processors.stripPrefix],
                valueProcessors: [processors.stripPrefix],
                attrValueProcessors: [processors.stripPrefix]
              }, function(err, result){
                if(err) {
                  cb(err, null);
                } else {
                  // var requestResult = result['Envelope']['Body'][0]['ExpandDLResponse'][0]['ResponseMessages'][0]['ExpandDLResponseMessage'][0]['DLExpansion'][0]['Mailbox'];
                  // console.log(requestResult);
                  cb(err, result);
                }
              });
            }
          });

        }
      });
    }
  ],function(err, result) {
    tmp.setGracefulCleanup();
    if(err) {
      callback(err, null);
    } else {
      callback(null, result[3]);
    }
  });

};

module.exports = new EWS;
