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
var EWS = require('node-ews');

// exchange server connection info
var username = 'myuser@domain.com';
var password = 'mypassword';
var host = 'https://ews.domain.com';

// disable ssl verification
var options = {
//  rejectUnauthorized: false,
//  strictSSL: false
};

// initialize node-ews
var ews = new EWS(username, password, host, options);

// define a ews query on Public Distribution List by email
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

````
