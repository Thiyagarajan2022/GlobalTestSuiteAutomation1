const translate = require('google-translate-api');
translate('Purchase Order No', {to: 'en'}).then(res => {
  console.log(res.text);
}).catch(err => {
   console.error(err);
});
