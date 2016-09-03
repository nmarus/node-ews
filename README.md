# node-ews
###### A simple JSON wrapper for the Exchange Web Services (EWS) SOAP API

```
npm install node-ews
```

#### Updates in version 2.0.0

- Removed xml2js dependancy and using XML parser built into soap library
- Removed async dependancy and library now returns promises
- Changed constructor setup
- Added optional "options" to constructor that is passed to soap library
- Host config now allows specifying http or https urls
- Replaced customized ntlm library with httpntlm

#### Updates in version 2.1.0

- Added the ability to specify the directory where XSD and WSDL files are stored
- Added authentication error handling
- Added WSDL validation error handling
- Added logic to determine if XSD and WDSL files have already been downloaded rather than attempting download on each run

#### About
A extension of node-soap with httpntlm to make queries to Microsoft's Exchange Web Service API work.

##### Features:
- Assumes NTLM Authentication over HTTPs (basic auth not currently supported)
- Connects to configured EWS Host and downloads it's WSDL file so it might be concluded that this is "fairly" version agnostic
- After downloading the  WSDL file, the wrapper dynamically exposes all EWS SOAP functions
- Attempts to standardize Microsoft's  WSDL by modifying the file to include missing service name, port, and bindings
- This DOES NOT work with anything Microsoft Documents as using the EWS Managed API.

#### Example 1: Get Exchange Distribution List Members Using ExpandDL
###### https://msdn.microsoft.com/EN-US/library/office/aa564755.aspx
```js
var EWS = require('node-ews');

// exchange server connection info
var username = 'myuser@domain.com';
var password = 'mypassword';
var host = 'https://ews.domain.com';

// initialize node-ews
var ews = new EWS(username, password, host);

var ewsFunction = 'ExpandDL';
var ewsArgs = {
  'Mailbox': {
    'EmailAddress':'publiclist@domain.com'
  }
};

// query ews, print resulting JSON to console
ews.run(ewsFunction, ewsArgs)
  .then(result => {
    console.log(JSON.stringify(result));
  })
  .catch(err => {
    console.log(err.message);
  });

```

#### Example 2: Setting OOO Using SetUserOofSettings
###### https://msdn.microsoft.com/en-us/library/office/aa580294.aspx
```js
var EWS = require('node-ews');

// exchange server connection info
var username = 'myuser@domain.com';
var password = 'mypassword';
var host = 'https://ews.domain.com';

// initialize node-ews
var ews = new EWS(username, password, host);

var ewsFunction = 'SetUserOofSettings';
var ewsArgs = {
  'Mailbox': {
    'Address':'email@somedomain.com'
  },
  'UserOofSettings': {
    'OofState':'Enabled',
    'ExternalAudience':'All',
    'Duration': {
      'StartTime':'2016-08-22T00:00:00',
      'EndTime':'2016-08-23T00:00:00'
    },
    'InternalReply': {
      'Message':'I am out of office.  This is my internal reply.'
    },
    'ExternalReply': {
      'Message':'I am out of office. This is my external reply.'
    }
  }
};

// query ews, print resulting JSON to console
ews.run(ewsFunction, ewsArgs)
  .then(result => {
    console.log(JSON.stringify(result));
  })
  .catch(err => {
    console.log(err.stack);
  });
```

#### Example 3: Getting OOO Using GetUserOofSettings
###### https://msdn.microsoft.com/en-us/library/office/aa563465.aspx
```js
var EWS = require('node-ews');

// exchange server connection info
var username = 'myuser@domain.com';
var password = 'mypassword';
var host = 'https://ews.domain.com';

// initialize node-ews
var ews = new EWS(username, password, host);

var ewsFunction = 'GetUserOofSettings';
var ewsArgs = {
  'Mailbox': {
    'Address':'email@somedomain.com'
  }
};

// query ews, print resulting JSON to console
ews.run(ewsFunction, ewsArgs)
  .then(result => {
    console.log(JSON.stringify(result));
  })
  .catch(err => {
    console.log(err.stack);
  });
```


### Advanced Options

#### Disable SSL verification:

To disable SSL authentication modify the above examples with the following:

```js
var options = {
 rejectUnauthorized: false,
 strictSSL: false
};
process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';

var ews = new EWS(username, password, host, options);
```

#### Specify Temp Directory:

By default, node-ews creates a temp directory to store the xsd and wsdl file retrieved from the Exchange Web Service Server.

To override this behavior and use a persistent folder, add the following before you you execute `ews.run()`

```js
ews.tempDir = '/path/to/temp/folder';
```

# Constructing the ewsArgs JSON Object

Take the example of "[FindItem](https://msdn.microsoft.com/en-us/library/office/aa566107.aspx)" that is referenced in the Microsoft EWS API Docs as follows:

```xml
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <FindItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
              Traversal="Shallow">
      <ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
      </ItemShape>
      <ParentFolderIds>
        <t:DistinguishedFolderId Id="deleteditems"/>
      </ParentFolderIds>
    </FindItem>
  </soap:Body>
</soap:Envelope>
```

The equivalent JSON Request for node-ews (ewsArgs) would be:

```js
{
  'attributes': {
    'Traversal': 'Shallow'
  },
  'ItemShape': {
    'BaseShape': 'IdOnly'
  },
  'ParentFolderIds' : {
    'DistinguishedFolderId': {
      'attributes': {
        'Id': 'deleteditems'
      }
    }
  }
}
```

It is important to note the structure of the request. Items such as "BaseShape" are children of their parent element. In this case it is "ItemShape". In regards "BaseShape" it is enclosed between an opening and closing tag so it is defined as a direct clid of "ItemShape".

However, "DistinguishedFolderId" has no closing tag and you must specify an ID. Rather than a direct child object, you must use the JSON child object "attributes".

**Alternatively you can create something like this to convert the EWS Soap Query to a SOAP JSON query:**

```js
var xml2js = require('xml2js');
var when = require('when');
var _ = require('lodash');

var util = require('util');

function convert(xml) {

  var attrkey = 'attributes';

  var parser = new xml2js.Parser({
    attrkey: attrkey,
    trim: true,
    ignoreAttrs: false,
    explicitRoot: false,
    explicitCharkey: false,
    explicitArray: false,
    explicitChildren: false,
    tagNameProcessors: [
      function(tag) {
        return tag.replace('t:', '');
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
      '<FindItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
                'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
                'Traversal="Shallow">' +
        '<ItemShape>' +
          '<t:BaseShape>IdOnly</t:BaseShape>' +
        '</ItemShape>' +
        '<ParentFolderIds>' +
          '<t:DistinguishedFolderId Id="deleteditems"/>' +
        '</ParentFolderIds>' +
      '</FindItem>' +
    '</soap:Body>' +
  '</soap:Envelope>';

convert(xml).then(json => {
  console.log('ewsArgs = ' + util.inspect(json, false, null));
});

// console output ready for ewsArgs

// ewsArgs = { attributes: { Traversal: 'Shallow' },
//   ItemShape: { BaseShape: 'IdOnly' },
//   ParentFolderIds: { DistinguishedFolderId: { attributes: { Id: 'deleteditems' } } } }
```
