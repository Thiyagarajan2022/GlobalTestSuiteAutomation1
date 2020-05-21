/* This simple app uses the '/translate' resource to translate text from
one language to another. */

/* This template relies on the request module, a simplified and user friendly
way to make HTTP requests. */
const request = require('request');
const uuidv4 = require('uuid/v4');
var EventEmitter = require("events").EventEmitter;
var body = new EventEmitter();
var key_var = "c50aa997c50f4c3c95fa1f1c9b54b102";
/*if (!process.env[key_var]) {
    throw new Error('Please set/export the following environment variable: ' + key_var);
}*/
var subscriptionKey = key_var;
var endpoint_var = "https://api.cognitive.microsofttranslator.com/translate?api-version=3.0";
/*if (!process.env[endpoint_var]) {
    throw new Error('Please set/export the following environment variable: ' + endpoint_var);
}*/
var endpoint = endpoint_var;

const args = process.argv;
var path = args[1];
var lang = args[2];
var  convtword = args[3];

/* If you encounter any issues with the base_url or path, make sure that you are
using the latest endpoint: https://docs.microsoft.com/azure/cognitive-services/translator/reference/v3-0-translate */
function translateText(){
var myObj, i, j, x = "";
    let options = {
        method: 'POST',
        baseUrl: endpoint,
        url: 'translate',
        qs: {
          'api-version': '3.0',
          'to': lang
        },
        headers: {
          'Ocp-Apim-Subscription-Key': subscriptionKey,
          'Ocp-Apim-Subscription-Region':'centralindia'  ,
          'Content-type': 'application/json',
          'X-ClientTraceId': uuidv4().toString()
        },
        body: [{
              'text': convtword
        }],
        json: true,
    };

    request(options, function(err, res, data){
    body.data = data;
    body.emit('update');
    });

body.on('update', function () { 
myObj = (body.data);
  console.log(myObj[0].translations[0].text) ;
 

});

};

// Call the function to translate text.
translateText(path,lang,convtword);