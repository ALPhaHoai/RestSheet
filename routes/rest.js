const express = require('express');
const sheet_table = require('../core/sheet_table');
const router = express.Router();

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
    const data = await sheet_table.getTableData();
    res.json(data);
});

router.post('/createTable', async function (req, res, next) {
    const {name, columns} = req.body
    const data = await sheet_table.createTable(name, columns);
    res.json(data);
});

router.post('/getTableData', async function (req, res, next) {
    const {name} = req.body
    const data = await sheet_table.getTableData(name);
    res.json(data);
});
router.post('/insertRows', async function (req, res, next) {
    const {name, rows} = req.body
    const data = await sheet_table.insertRows(name, rows);
    res.json(data);
});


module.exports = router