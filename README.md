# node-ews
###### A simple JSON wrapper for the Exchange Web Services (EWS) SOAP API

```
npm install node-ews
```

#### About
A extension of node-soap with httpntlm to make queries to Microsoft's Exchange Web Service API work.

##### Features:
- Assumes NTLM Authentication over HTTPs
- Connects to configured EWS Host and downloads it's wsdl file so it might be concluded that this is "fairly" version agnostic
- After downloading the wsdl file, the wrapper dynamically exposes all EWS SOAP functions
- Attempts to standardize Microsoft's wsdl by modifying the file to include missing service name, port, and bindings

#### Example 1: Get Exchange Distribution List Members Using ExpandDL
###### https://msdn.microsoft.com/EN-US/library/office/aa564755.aspx
```js
var ews = require('node-ews');

// exchange server connection info
var username = 'myuser@domain.com';
var password = 'mypassword';
var host = 'ews.domain.com';

// exchange ews query on Public Distribution List by email
var ewsFunction = 'ExpandDL';
var ewsArgs = {
  'Mailbox': {
    'EmailAddress':'publiclist@domain.com'
  }
};

// ignore ssl verification (optional)
ews.ignoreSSL = true;

// setup authentication
ews.auth(username, password, host);

// query ews, print resulting JSON to console
ews.run(ewsFunction, ewsArgs, function(err, result) {
  if(!err) console.log(JSON.stringify(result));
});
````

#### Custom Soap Headers
```js
ews.run('FindItem', {
  body: {
    attributes: {
      Traversal: 'Shallow'
    },
    ItemShape: {
      BaseShape: 'IdOnly',
      AdditionalProperties: {
        FieldURI: {
          attributes: {
            FieldURI: 'item:Subject'
          }
        }
      }
    },
    CalendarView: {
      attributes: {
        MaxEntriesReturned: 10,
        StartDate: '2016-01-01T00:00:00Z',
        EndDate: '2016-12-31T00:00:00Z'
      }
    },
    ParentFolderIds: {
      DistinguishedFolderId: {
        attributes: {
          Id: 'calendar'
        }
      }
    }
  },
  headers: {
    'http://schemas.microsoft.com/exchange/services/2006/types': [{
      RequestServerVersion: {
        attributes: {
          Version: 'Exchange2010'
        }
      },
      ExchangeImpersonation: {
        ConnectingSID: {
          SmtpAddress: '...'
        }
      }
    }]
  }
}, function(err, result) {});
```

#### Specifying XML processors
You can specify the XML processors you want to use when parsing the response. If you are seeing what should be Date/Times having values like '00Z', disable the stripPrefix processor that is run by default on the XML value by specifying:
```
ews.processors: {
  valueProcessors: null
}
```
Alternatively you can pass in an array of functions which will be merged with the default stripPrefix processor. Details can be found at the [node-xml2js](https://github.com/Leonidas-from-XIV/node-xml2js#processing-attribute-tag-names-and-values)

#### Known Issues / Limits / TODOs:
- Returned json requires a lot of parsing. Probably can be optimized to remove common parent elements to the EWS responses or dynamically filter based on query.
- Outside of the example above, nothing has been tested (aka "It's production ready!")
- Temp file cleanup logic needs to be validated to ensure file cleanup after process exit or object destruction
