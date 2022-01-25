const express = require('express');
const router = express.Router();
var googlesheet = require('../core/googlesheet');

router.get('/', function (req, res, next) {
    const limit = Math.max(parseInt(req.limit) || 100, 100)
    const table = req.table
    res.json();
});

router.get('/getAuthurl', async function (req, res, next) {
    let authUrl = await googlesheet.generateAuthUrl();
    res.json({
        url: authUrl
    })
});

router.get('/generateAuthUrl', async function (req, res, next) {
    let authUrl = await googlesheet.generateAuthUrl();
    res.redirect(authUrl)
});

router.post('/updateCode', updateCode);
router.get('/updateCode', updateCode);

async function updateCode(req, res, next) {
    const {code} = req.method === "GET" ? req.query : req.body
    if (code) {
        const result = await googlesheet.updateCode(code)
        res.send({
            status: result
        })
    } else {
        res.redirect("/")
    }
}

module.exports = router