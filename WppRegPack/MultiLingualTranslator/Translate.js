const translate = require('google-translate-api');
translate('God', {to: 'en'}).then(res => {
  console.log(res.text);
}).catch(err => {
   console.error(err);
});
