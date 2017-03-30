# node-ews
###### A simple JSON wrapper for the Exchange Web Services (EWS) SOAP API

```
npm install node-ews
```

#### Updates in patch 3.1.0 (new)
- Merged PR #47 to add support for bearer auth.

#### Updates in patch 3.0.6
- Addressed issues in PR #38, cleaned up code.
- Merged PR for issues #37
- Merged PR for issues #36 to fix typo in 3.0.2
- Merged PR for issues #34 to fix typo in 3.0.1
- Applied temporary fix for Issue #17 by pointing node-soap reference in package.json to modified fork.

#### Updates in version 3.0.0

- Resolved issue #31 (courtesy of @eino-makitalo)
- Resolved issue #29
- Updated docs to reflect issue discovered in #26
- Resolved issue #20 #15 with added basic-auth functionality
- Updated README to include Office 365 connection example and notes on basic-auth
- The EWS constructor now requires authentication (username, password, host) inside config object. See updated examples below. **Note: This is a breaking change from the method used in version 2.x!**
- The temp file is now specified in the EWS config object. See updated examples below. **Note: This is a breaking change from the method used in version 2.x!**

#### About
A extension of node-soap with httpntlm to make queries to Microsoft's Exchange Web Service API work.

##### Features:
- Assumes NTLM Authentication over HTTPs (basic and bearer auth **now** supported)
- Connects to configured EWS Host and downloads it's WSDL file so it might be concluded that this is "fairly" version agnostic
- After downloading the  WSDL file, the wrapper dynamically exposes all EWS SOAP functions
- Attempts to standardize Microsoft's  WSDL by modifying the file to include missing service name, port, and bindings
- This DOES NOT work with anything Microsoft Documents as using the EWS Managed API.

#### Example 1: Get Exchange Distribution List Members Using ExpandDL
###### https://msdn.microsoft.com/EN-US/library/office/aa564755.aspx
```js
var EWS = require('node-ews');

// exchange server connection info
var ewsConfig = {
  username: 'myuser@domain.com',
  password: 'mypassword',
  host: 'https://ews.domain.com'
};

// initialize node-ews
var ews = new EWS(ewsConfig);

// define ews api function
var ewsFunction = 'ExpandDL';

// define ews api function args
var ewsArgs = {
  'Mailbox': {
    'EmailAddress':'publiclist@domain.com'
  }
};

// query EWS and print resulting JSON to console
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
var ewsConfig = {
  username: 'myuser@domain.com',
  password: 'mypassword',
  host: 'https://ews.domain.com'
};

// initialize node-ews
var ews = new EWS(ewsConfig);

// define ews api function
var ewsFunction = 'SetUserOofSettings';

// define ews api function args
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
var ewsConfig = {
  username: 'myuser@domain.com',
  password: 'mypassword',
  host: 'https://ews.domain.com'
};

// initialize node-ews
var ews = new EWS(ewsConfig);

// define ews api function
var ewsFunction = 'GetUserOofSettings';

// define ews api function args
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

#### Example 4: Sending Email (version 3.0.1 required to avoid issue #17)
##### https://msdn.microsoft.com/en-us/library/office/aa566468

```js
var EWS = require('node-ews');

// exchange server connection info
var ewsConfig = {
  username: 'myuser@domain.com',
  password: 'mypassword',
  host: 'https://ews.domain.com'
};

// initialize node-ews
var ews = new EWS(ewsConfig);

// define ews api function
var ewsFunction = 'CreateItem';

// define ews api function args
var ewsArgs = {
  "attributes" : {
    "MessageDisposition" : "SendAndSaveCopy"
  },
  "SavedItemFolderId": {
    "DistinguishedFolderId": {
      "attributes": {
        "Id": "sentitems"
      }
    }
  },
  "Items" : {
    "Message" : {
      "ItemClass": "IPM.Note",
      "Subject" : "Test EWS Email",
      "Body" : {
        "attributes": {
          "BodyType" : "Text"
        },
        "$value": "This is a test email"
      },
      "ToRecipients" : {
        "Mailbox" : {
          "EmailAddress" : "someone@gmail.com"
        }
      },
      "IsRead": "false"
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

### Office 365

Below is a template that works with Office 365.

```js
var EWS = require('node-ews');

// exchange server connection info
var ewsConfig = {
  username: 'myuser@domain.com',
  password: 'mypassword',
  host: 'https://outlook.office365.com',
  auth: 'basic'
};

// initialize node-ews
var ews = new EWS(ewsConfig);

// define ews api function
var ewsFunction = 'ExpandDL';

// define ews api function args
var ewsArgs = {
  'Mailbox': {
    'EmailAddress':'publiclist@domain.com'
  }
};

// query EWS and print resulting JSON to console
ews.run(ewsFunction, ewsArgs)
  .then(result => {
    console.log(JSON.stringify(result));
  })
  .catch(err => {
    console.log(err.message);
  });
```

### Advanced Options

#### Adding Optional Soap Headers

To add an optional soap header to the Exchange Web Services request, you can pass an optional 3rd variable to the ews.run() function as demonstrated by the following:

```js
var EWS = require('node-ews');

// exchange server connection info
var ewsConfig = {
  username: 'myuser@domain.com',
  password: 'mypassword',
  host: 'https://ews.domain.com'
};

// initialize node-ews
var ews = new EWS(ewsConfig);

// define ews api function
var ewsFunction = 'GetUserOofSettings';

// define ews api function args
var ewsArgs = {
  'Mailbox': {
    'Address':'email@somedomain.com'
  }
};

// define custom soap header
var ewsSoapHeader = {
  't:RequestServerVersion': {
    attributes: {
      Version: "Exchange2013"
    }
  }
};

// query ews, print resulting JSON to console
ews.run(ewsFunction, ewsArgs, ewsSoapHeader)
  .then(result => {
    console.log(JSON.stringify(result));
  })
  .catch(err => {
    console.log(err.stack);
  });

```

#### Enable Basic Auth instead of NTLM:

```js
// exchange server connection info
var ewsConfig = {
  username: 'myuser@domain.com',
  password: 'mypassword',
  host: 'https://ews.domain.com',
  auth: 'basic'
};

// initialize node-ews
var ews = new EWS(ewsConfig);
```

#### Enable Bearer Auth instead of NTLM:

```js
// exchange server connection info
var ewsConfig = {
  username: 'myuser@domain.com',
  token: 'oauth_token...',
  host: 'https://ews.domain.com',
  auth: 'bearer'
};

// initialize node-ews
var ews = new EWS(ewsConfig);
```

#### Disable SSL verification:

To disable SSL authentication modify the above examples with the following:

**Basic and Bearer Auth**

```js
var options = {
 rejectUnauthorized: false,
 strictSSL: false
};
process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';

var ews = new EWS(config, options);
```

**NTLM**

```js
var options = {
 strictSSL: false
};
process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';

var ews = new EWS(config, options);
```

#### Specify Temp Directory:

By default, node-ews creates a temp directory to store the xsd and wsdl file retrieved from the Exchange Web Service Server.

To override this behavior and use a persistent folder add the following to your config object.

```js
// exchange server connection info
var ewsConfig = {
  username: 'myuser@domain.com',
  password: 'mypassword',
  host: 'https://ews.domain.com',
  temp: '/path/to/temp/folder'
};

// initialize node-ews
var ews = new EWS(ewsConfig);
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

## License

MIT License Copyright (c) 2016 Nicholas Marus

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
