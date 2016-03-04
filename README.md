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

#### Known Issues / Limits / TODOs:
- Returned json requires a lot of parsing. Probably can be optimized to remove common parent elements to the EWS responses or dynamically filter based on query.
- Outside of the example above, nothing has been tested (aka "It's production ready!")
- Temp file cleanup logic needs to be validated to ensure file cleanup after process exit or object destruction
