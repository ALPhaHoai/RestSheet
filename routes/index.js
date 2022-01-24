var express = require('express');
var googlesheet = require('./googlesheet');
var router = express.Router();

router.get('/', function (req, res, next) {
  const limit = Math.max(parseInt(req.limit) || 100, 100)
  const table = req.table
  res.json();
});

router.get('/google', async function (req, res, next) {
  let authUrl = await googlesheet.generateAuthUrl();
  res.redirect(authUrl)
});

router.post('/google', async function (req, res, next) {
  const {code} = req.body
  if (code) {
    const result = await googlesheet.updateCode(code)
    res.send({
      status: result
    })
  } else {
    res.redirect("/")
  }
});

module.exports = router;
