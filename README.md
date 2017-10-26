# node-ews
###### A simple JSON wrapper for the Exchange Web Services (EWS) SOAP API

```
npm install node-ews
```

#### Updates in 3.2.0 (new)
- Reverted back to official soap library to address workaround in issue #17
- Added Notifications from PR #50 (See example 5)
- Code cleanup

#### About
A extension of node-soap with httpntlm to make queries to Microsoft's Exchange
Web Service API work. Utilize node-soap for json to xml query processing and
returns responses as json objects.

##### Features:
- Supports NTLM, Basic, or Bearer Authentication.
- Connects to configured EWS Host and downloads it's WSDL file so it might be concluded that this is "fairly" version agnostic
- After downloading the WSDL file, the wrapper dynamically exposes all EWS SOAP functions
- Attempts to standardize Microsoft's WSDL by modifying the file to include missing service name, port, and bindings
- This DOES NOT work with anything Microsoft Documents as using the EWS Managed API.

#### Example 1: Get Exchange Distribution List Members Using ExpandDL
###### https://msdn.microsoft.com/EN-US/library/office/aa564755.aspx
```js
const EWS = require('node-ews');

// exchange server connection info
const ewsConfig = {
  username: 'myuser@domain.com',
  password: 'mypassword',
  host: 'https://ews.domain.com'
};

// initialize node-ews
const ews = new EWS(ewsConfig);

// define ews api function
const ewsFunction = 'ExpandDL';

// define ews api function args
const ewsArgs = {
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
const EWS = require('node-ews');

// exchange server connection info
const ewsConfig = {
  username: 'myuser@domain.com',
  password: 'mypassword',
  host: 'https://ews.domain.com'
};

// initialize node-ews
const ews = new EWS(ewsConfig);

// define ews api function
const ewsFunction = 'SetUserOofSettings';

// define ews api function args
const ewsArgs = {
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
const EWS = require('node-ews');

// exchange server connection info
const ewsConfig = {
  username: 'myuser@domain.com',
  password: 'mypassword',
  host: 'https://ews.domain.com'
};

// initialize node-ews
const ews = new EWS(ewsConfig);

// define ews api function
const ewsFunction = 'GetUserOofSettings';

// define ews api function args
const ewsArgs = {
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
const EWS = require('node-ews');

// exchange server connection info
const ewsConfig = {
  username: 'myuser@domain.com',
  password: 'mypassword',
  host: 'https://ews.domain.com'
};

// initialize node-ews
const ews = new EWS(ewsConfig);

// define ews api function
const ewsFunction = 'CreateItem';

// define ews api function args
const ewsArgs = {
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

#### Example 5: Creating a Push Notification Service Listener

```js
// specify listener service options
const serviceOptions = {
  port: 8080, // defaults to port 8000
  path: '/', // defaults to '/notification'
  // If you do not have NotificationService.wsdl it can be found via a quick Google search
  xml:fs.readFileSync('NotificationService.wsdl', 'utf8') // the xml field is required
};

// create the listener service
ews.notificationService(serviceOptions, function(response) {
  console.log(new Date().toISOString(), '| Received EWS Push Notification');
  console.log(new Date().toISOString(), '| Response:', JSON.stringify(response));
  // Do something with response
  return {SendNotificationResult: { SubscriptionStatus: 'OK' } }; // respond with 'OK' to keep subscription alive
  // return {SendNotificationResult: { SubscriptionStatus: 'UNSUBSCRIBE' } }; // respond with 'UNSUBSCRIBE' to unsubscribe
});

// The soap.listen object is passed through the promise so you can optionally use the .log() functionality
// https://github.com/vpulim/node-soap#server-logging
.then(server => {
  server.log = function(type, data) {
    console.log(new Date().toISOString(), '| ', type, ':', data);
  };
});

// create a push notification subscription
// https://msdn.microsoft.com/en-us/library/office/aa566188
const ewsConfig = {
  PushSubscriptionRequest: {
    FolderIds: {
      DistinguishedFolderId: {
        attributes: {
          Id: 'inbox'
        }
      }
    },
    EventTypes: {
      EventType: ['CreatedEvent']
    },
    StatusFrequency: 1,
    // subscription notifications will be sent to our listener service
    URL: 'http://' + require('os').hostname() + ':' + serviceOptions.port + serviceOptions.path
  }
};
ews.run('Subscribe', ewsConfig);
```

### Office 365

Below is a template that works with Office 365.

```js
const EWS = require('node-ews');

// exchange server connection info
const ewsConfig = {
  username: 'myuser@domain.com',
  password: 'mypassword',
  host: 'https://outlook.office365.com',
  auth: 'basic'
};

// initialize node-ews
const ews = new EWS(ewsConfig);

// define ews api function
const ewsFunction = 'ExpandDL';

// define ews api function args
const ewsArgs = {
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
const EWS = require('node-ews');

// exchange server connection info
const ewsConfig = {
  username: 'myuser@domain.com',
  password: 'mypassword',
  host: 'https://ews.domain.com'
};

// initialize node-ews
const ews = new EWS(ewsConfig);

// define ews api function
const ewsFunction = 'GetUserOofSettings';

// define ews api function args
const ewsArgs = {
  'Mailbox': {
    'Address':'email@somedomain.com'
  }
};

// define custom soap header
const ewsSoapHeader = {
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

#### Use Encrypted Credentials for NTLM:
This allows you to persist their password as separate hashes instead of as plain text.
This utilizes the [options](https://github.com/SamDecrock/node-http-ntlm#options) available to the underlying NTLM lib.
[Here](https://github.com/SamDecrock/node-http-ntlm#pre-encrypt-the-password) is an example from its README.
Below is an example for this lib:
```js
const NTLMAuth = require('httpntlm').ntlm;
const passwordPlainText = 'mypassword';

// store the ntHashedPassword and lmHashedPassword to reuse later for reconnecting
const ntHashedPassword = NTLMAuth.create_NT_hashed_password(passwordPlainText);
const lmHashedPassword = NTLMAuth.create_LM_hashed_password(passwordPlainText);

// exchange server connection info
const ewsConfig = {
    username: 'myuser@domain.com',
    nt_password: ntHashedPassword,
    lm_password: lmHashedPassword,
    host: 'https://ews.domain.com'
};

// initialize node-ews
const ews = new EWS(ewsConfig);
```

#### Enable Basic Auth instead of NTLM:

```js
// exchange server connection info
const ewsConfig = {
  username: 'myuser@domain.com',
  password: 'mypassword',
  host: 'https://ews.domain.com',
  auth: 'basic'
};

// initialize node-ews
const ews = new EWS(ewsConfig);
```

#### Enable Bearer Auth instead of NTLM:

```js
// exchange server connection info
const ewsConfig = {
  username: 'myuser@domain.com',
  token: 'oauth_token...',
  host: 'https://ews.domain.com',
  auth: 'bearer'
};

// initialize node-ews
const ews = new EWS(ewsConfig);
```

#### Disable SSL verification:

To disable SSL authentication modify the above examples with the following:

**Basic and Bearer Auth**

```js
const options = {
 rejectUnauthorized: false,
 strictSSL: false
};
process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';

const ews = new EWS(config, options);
```

**NTLM**

```js
const options = {
 strictSSL: false
};
process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';

const ews = new EWS(config, options);
```

#### Specify Temp Directory:

By default, node-ews creates a temp directory to store the xsd and wsdl file retrieved from the Exchange Web Service Server.

To override this behavior and use a persistent folder add the following to your config object.

```js
// exchange server connection info
const ewsConfig = {
  username: 'myuser@domain.com',
  password: 'mypassword',
  host: 'https://ews.domain.com',
  temp: '/path/to/temp/folder'
};

// initialize node-ews
const ews = new EWS(ewsConfig);
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
