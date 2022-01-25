const {spreadsheets} = require("../config")
const googlesheet = require("./googlesheet");
const spreadsheetId = spreadsheets[0].spreadsheetId
const {v4: uuidv4} = require('uuid')

module.exports = {
    async createTable(name, columns) {
        if (!columns?.includes('Id')) {
            return {
                success: false,
                msg: 'Missing Id column'
            }
        }

        //move Id column to first
        columns.sort((c1, c2) => c1 === 'Id' ? -1 : (c2 === 'Id' ? 1 : 0))

        if (!await googlesheet.isHasSheet(spreadsheetId, name)) {
            if (!await googlesheet.createSheet(spreadsheetId, name)) {
                return {
                    success: false,
                    msg: `Can not create table ${name}`
                }
            }
            const result = await googlesheet.fillRow(spreadsheetId, name, 1, columns)
            if (result) {
                return {
                    success: true,
                }
            } else {
                return {
                    success: false,
                    msg: `Can not create table ${name}`
                }
            }
        } else {
            return {
                success: false,
                msg: `Table ${name} already exists`
            }
        }
    },
    async getTableData(tableName) {
        const data = await googlesheet.getSpreadData(spreadsheetId, tableName)
        if (!Array.isArray(data) || data.length < 2) {
            return []
        }
        const _data = []
        const columns = data[0]
        for (let i = 1; i < data.length; i++) {
            const rowArr = data[i]
            const rowObj = {}
            for (let j = 0; j < columns.length; j++) {
                rowObj[columns[j]] = typeof rowArr[j] === "undefined" ? '' : rowArr[j]
            }
            _data.push(rowObj)
        }
        return _data
    },
    async insertRows(tableName, rows) {
        rows.forEach(r => r.Id = uuidv4())
        const result = await googlesheet.insertRows(spreadsheetId, tableName, rows)
        return result ? {
            success: true,
            data: rows
        } : {
            success: false,
            msg: `Can not insert ${rows.length} row(s) in table ${tableName}`
        }
    }
}