# node-ews
#### A simple JSON wrapper for the Exchange Web Services (EWS) SOAP API


##### Example 1: Get Exchange Distribution List Members Using ExpandDL
###### https://msdn.microsoft.com/EN-US/library/office/aa564755.aspx
```js
var ews = require('./node-ews');
var _ = require('lodash');

// exchange server connection info
var username = 'myuser@domain.com';
var password = 'mypassword';
var ewsHost = 'ews.domain.com';

// exchange ews query
var ewsFunction = 'ExpandDL';
var ewsArgs = {
  'Mailbox': {
    'EmailAddress':'publiclist@domain.com'
  }
};

// setup authentication
ews.auth(username, password, ewsHost);

// query ews
ews.get(ewsFunction, ewsArgs, function(err, result) {
  if(err) throw err;

  // parse JSON results to dump email addresses for DL to console
  var requestResult = result['Envelope']['Body'][0]['ExpandDLResponse'][0]['ResponseMessages'][0]['ExpandDLResponseMessage'][0]['DLExpansion'][0]['Mailbox'];
  _.map(requestResult, 'EmailAddress').forEach(function(x) {
    console.log(x[0]);
  });
});
````
