const {spreadsheets, tableConfig} = require("../config")
const googlesheet = require("./googlesheet");
const spreadsheetId = spreadsheets[0].spreadsheetId
const {v4: uuidv4} = require('uuid')
const {DateTime} = require("luxon");


const sheet_table = {
    async createTable(name, columns) {
        columns = getValidateColumns(columns)
        if (!columns?.includes(tableConfig.idColumn)) {
            return {
                success: false,
                msg: `Missing ${tableConfig.idColumn} column`
            }
        }

        //move Id column to first
        columns.sort((c1, c2) => c1 === tableConfig.idColumn ? -1 : (c2 === tableConfig.idColumn ? 1 : 0))

        //add createdAt, updatedAt columns
        if (!columns?.includes(tableConfig.createdAtColumn)) {
            columns?.push(tableConfig.createdAtColumn)
        }
        if (!columns?.includes(tableConfig.updatedAtColumn)) {
            columns?.push(tableConfig.updatedAtColumn)
        }
        columns.sort((c1, c2) => c1 === tableConfig.createdAtColumn ? 1 : (c2 === tableConfig.createdAtColumn ? -1 : 0))
        columns.sort((c1, c2) => c1 === tableConfig.updatedAtColumn ? 1 : (c2 === tableConfig.updatedAtColumn ? -1 : 0))

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
        if (!Array.isArray(rows)) {
            return {
                success: false,
                msg: `Rows input invalid`
            }
        }
        rows.forEach(function (r) {
            r[tableConfig.idColumn] = uuidv4()
            r[tableConfig.updatedAtColumn] = r[tableConfig.createdAtColumn] = DateTime.now().toString()
        })
        const result = await googlesheet.insertRows(spreadsheetId, tableName, rows)
        return result ? {
            success: true,
            data: rows
        } : {
            success: false,
            msg: `Can not insert ${rows.length} row(s) in table ${tableName}`
        }
    },
    async updateRows(tableName, row, conditions) {
        if (!row) {
            return {
                success: false,
                msg: `Rows input invalid`
            }
        }
        row[tableConfig.updatedAtColumn] = DateTime.now().toString()

        const data = await sheet_table.find(tableName, conditions, undefined, true)
        if (!Array.isArray(data)) return false
        const result = await googlesheet.updateRows(spreadsheetId, tableName, row, data.map(r => r[tableConfig.indexColumn]))
        if (result?.success) {
            result.row = await sheet_table.find(tableName, conditions)
        }
        return result
    },
    async find(tableName, conditions, limit = undefined, withIndex = false) {
        const data = await googlesheet.getSpreadData(spreadsheetId, tableName)
        if (!Array.isArray(data)) return []
        const columns = data[0]
        const results = []
        for (let i = 1; i < data.length; i++) {
            let match = true
            for (let j = 0; j < columns.length; j++) {
                const column = columns[j]
                if (conditions[column] !== undefined) {
                    if (conditions[column] !== data[i][j]) {
                        match = false
                        break
                    }
                }
            }
            if (match) {
                const obj = {}
                for (let j = 0; j < columns.length; j++) {
                    const column = columns[j]
                    obj[column] = data[i][j]
                }
                if (withIndex) {
                    obj[tableConfig.indexColumn] = i + 1
                }
                results.push(obj)
                if (typeof limit === "number" && limit <= results.length) {
                    break
                }
            }
        }
        return results
    },
    async delete(tableName, conditions) {
        const data = await sheet_table.find(tableName, conditions, undefined, true)
        if (!Array.isArray(data)) return false
        const result = await googlesheet.deleteRows(spreadsheetId, tableName, data.map(r => r[tableConfig.indexColumn]))
        if (result) {
            return {
                success: true,
                deletedRow: data.map(r => ({[tableConfig.idColumn]: r[tableConfig.idColumn]}))
            }
        } else return {
            success: false,
        }
    }
}

module.exports = sheet_table