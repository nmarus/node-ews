var xml2js = require('xml2js');
var when = require('when');
var _ = require('lodash');

var util = require('util');

function convert(xml) {

  var attrkey = 'attributes';
  var charkey = '$value'; // added this to get the correct key > value that exchange expects

  var parser = new xml2js.Parser({
    attrkey: attrkey,
    charkey: charkey,
    trim: true,
    ignoreAttrs: false,
    explicitRoot: false,
    explicitCharkey: false,
    explicitArray: false,
    explicitChildren: false,
    tagNameProcessors: [
      function(tag) {
        return tag.replace('t:', '').replace('m:',''); // and this to cleanup some extra tags,
                                                       // not specifically used in this example but it is needed
      }
    ]
  });

  return when.promise((resolve, reject) => {
    parser.parseString(xml, (err, result) => {
      if(err) reject(err);
      else {
        var ewsFunction = _.keys(result['soap:Body'])[0];
        var parsed = result['soap:Body'][ewsFunction];
        parsed[attrkey] = _.omit(parsed[attrkey], ['xmlns', 'xmlns:t']);
        if(_.isEmpty(parsed[attrkey])) parsed = _.omit(parsed, [attrkey]);
        resolve(parsed);
      }
    });
  });

}

var xml = '<?xml version="1.0" encoding="utf-8"?>' +
  '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
                 'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '<soap:Body>' +
      '<CreateItem MessageDisposition="SendAndSaveCopy" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
      '<SavedItemFolderId>' +
        '<t:DistinguishedFolderId Id="sentitems" />' +
      '</SavedItemFolderId>' +
      '<Items>' +
        '<t:Message>' +
          '<t:ItemClass>IPM.Note</t:ItemClass>' +
          '<t:Subject>EWS node email</t:Subject>' +
          '<t:Body BodyType="HTML"><![CDATA[This is <b>bold</b> - <b>Priority</b> - Update specification]]></t:Body>' +
          '<t:ToRecipients>' +
            '<t:Mailbox>' +
              '<t:EmailAddress>me@example.com</t:EmailAddress>' +
            '</t:Mailbox>' +
          '</t:ToRecipients>' +
          '<t:IsRead>false</t:IsRead>' +
        '</t:Message>' +
      '</Items>' +
    '</CreateItem>' +
    '</soap:Body>' +
  '</soap:Envelope>';

convert(xml).then(json => {
  console.log('ewsArgs = ' + util.inspect(json, false, null));
});
