var express = require('express');
var googlesheet = require('../core/googlesheet');
var config = require('../config');
var router = express.Router();

/*
[
    {
        "properties": {
            "sheetId": 621183740,
            "title": "Students",
            "index": 1,
            "sheetType": "GRID",
            "gridProperties": {
                "rowCount": 1000,
                "columnCount": 26
            }
        },
        "data": {
            "range": "Students!A1:Z1000",
            "majorDimension": "ROWS",
            "values": [
                [
                    "x",
                    "y"
                ]
            ]
        }
    }
]
*/
router.get('/', async function (req, res, next) {
    const limit = Math.max(parseInt(req.query.limit) || 100, 100)
    const table = req.query.table
    const data = await googlesheet.getSpreadData(config.spreadsheets[0].spreadsheetId);
    res.json(data);
});


module.exports = router