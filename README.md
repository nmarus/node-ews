# node-ews
###### A simple JSON wrapper for the Exchange Web Services (EWS) SOAP API

#### About
A collobberation of node-soap and httpntlm wrapped up with some modifications to make queries to Microsoft's Exchange Web Service API work. I'm actually surprised this works...

##### Features:
- Assumes NTLM Authentication over HTTPs
- Connects to configured EWS Host and downloads appropriate wsdl file so it might be concluded that this is "fairly" version agnostic
- Dynamically exposes all EWS SOAP functions in the downloaded wsdl file
- Attempts to standardize Microsoft's wsdl by modifying the file to include missing components

##### Known Issues / Limits / TODOs:
- Depends on a modified fork of the soap-ntlm fork of node-soap. This is located in this repository.
- Downloads to temporary location the wsdl and xsd files on each call of ews.run()
- This is not published as a NPM
- Returned json requires a lot of parsing. Probably can be optimized to remove common parent elements to the EWS responses or dynamically filter based on query.
- Outside of the example below, nothing has been tested (aka "production ready!")

#### Example 1: Get Exchange Distribution List Members Using ExpandDL
###### https://msdn.microsoft.com/EN-US/library/office/aa564755.aspx
```js
var ews = require('./node-ews');
var _ = require('lodash');

// exchange server connection info
var username = 'myuser@domain.com';
var password = 'mypassword';
var ewsHost = 'ews.domain.com';

// exchange ews query on Public Distribution List by email
var ewsFunction = 'ExpandDL';
var ewsArgs = {
  'Mailbox': {
    'EmailAddress':'publiclist@domain.com'
  }
};

// setup authentication
ews.auth(username, password, ewsHost);

// query ews
ews.run(ewsFunction, ewsArgs, function(err, result) {
  if(err) throw err;

  // parse JSON results to dump email addresses for DL to console
  var requestResult = result['Envelope']['Body'][0]['ExpandDLResponse'][0]['ResponseMessages'][0]['ExpandDLResponseMessage'][0]['DLExpansion'][0]['Mailbox'];
  _.map(requestResult, 'EmailAddress').forEach(function(x) {
    console.log(x[0]);
  });
});
````
