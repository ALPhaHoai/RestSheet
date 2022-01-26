const express = require('express');
const config = require('../config');
const fs = require('fs');
const readline = require('readline');
const {google} = require('googleapis');
const {tableConfig} = require("../config");
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
const TOKEN_PATH = 'token.json';
const maxRows = 10000
const maxColumns = 50

async function getSheetApi() {
    const auth = await getAuth()
    const sheets = google.sheets({version: 'v4', auth});
    return sheets
}

const getAuth = (function () {
    let auth

    return async function () {
        if (auth) return auth
        return new Promise((resolve, reject) => {
            authorize(function (oAuth2Client) {
                auth = oAuth2Client
                resolve(auth)
            })
        })
    }
})()

function authorize(resolve) {
    fs.readFile('credentials.json', (err, content) => {
        if (err) {
            console.log('Error loading client secret file:', err)
            resolve()
        } else {
            const credentials = JSON.parse(content)

            const {client_secret, client_id, redirect_uris} = credentials.installed || credentials.web;
            const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

            fs.readFile(TOKEN_PATH, (err, token) => {
                if (err) {
                    // getNewToken(oAuth2Client, resolve)
                    resolve()
                } else {
                    oAuth2Client.setCredentials(JSON.parse(token));
                    resolve(oAuth2Client);
                }
            });
        }
    });

}

async function updateCode(code) {
    return new Promise(resolve => {

        fs.readFile('credentials.json', (err, content) => {
            if (err) {
                resolve(false)
            } else {

                const credentials = JSON.parse(content)

                const {client_secret, client_id, redirect_uris} = credentials.installed || credentials.web;
                const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
                oAuth2Client.getToken(code, (err, token) => {
                    if (err) {
                        resolve(false)
                    } else {
                        oAuth2Client.setCredentials(token);
                        fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
                            if (err) {
                                resolve(false)
                            }
                            resolve(true)
                        });
                    }
                });
            }
        });
    })
}


async function generateAuthUrl() {
    return new Promise((resolve) => {
        fs.readFile('credentials.json', (err, content) => {
            if (err) {
                console.log('Error loading client secret file:', err)
                resolve()
            } else {

                const credentials = JSON.parse(content)

                const {client_secret, client_id, redirect_uris} = credentials.installed || credentials.web;
                const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
                const authUrl = oAuth2Client.generateAuthUrl({
                    access_type: 'offline',
                    scope: SCOPES,
                });
                resolve(authUrl)
            }
        });
    })
}

function getNewToken(oAuth2Client, callback) {
    const authUrl = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: SCOPES,
    });
    console.log('Authorize this app by visiting this url:', authUrl);
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
    });
    rl.question('Enter the code from that page here: ', (code) => {
        rl.close();
        oAuth2Client.getToken(code, (err, token) => {
            if (err) return console.error('Error while trying to retrieve access token', err);
            oAuth2Client.setCredentials(token);
            // Store the token to disk for later program executions
            fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
                if (err) return console.error(err);
                console.log('Token stored to', TOKEN_PATH);
            });
            callback(oAuth2Client);
        });
    });
}

//index >= 1
function getColumnLabelFromIndex(index) {
    const str = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    if (index <= str.length) {
        return str[index - 1]
    } else {
        const times = Math.floor(index / str.length)
        const sub = index - times * str.length + 1;
        return getColumnLabelFromIndex(times) + (sub > 0 ? getColumnLabelFromIndex(sub) : '')
    }
}


function getColumLabel(data, label) {
    if (!label) return
    if (Array.isArray(data)) {
        let row = data[0];
        if (row) {
            for (let j = 0; j < row.length; j++) {
                try {
                    if (row[j].trim().toLowerCase() == label.trim().toLowerCase()) {
                        return getColumnLabelFromIndex(j + 1)
                    }
                } catch (e) {
                }
            }
        }
    }
}

async function getSpreadData(spreadsheetId, sheetName) {
    if (!spreadsheetId) {
        return
    }

    let spreadInfo = await getSpreadInfo(spreadsheetId);
    if (!Array.isArray(spreadInfo.sheets)) return

    if (sheetName) {
        if (!Array.isArray(sheetName)) {
            sheetName = [sheetName]
        }
    } else {
        sheetName = spreadInfo.sheets.map(s => s.properties.title)
    }

    const ranges = sheetName.map(s => `'${s}'!1:${maxRows}`)

    const sheetApi = await getSheetApi()

    try {
        const result = await sheetApi.spreadsheets.values.batchGet({
            spreadsheetId: spreadsheetId,
            ranges,
        });
        return result?.data?.valueRanges?.[0]?.values
    } catch (e) {
        console.error(e);
        return null
    }
}

async function getSheetDataInfo(spreadsheetId, sheetName) {
    if (!spreadsheetId || !sheetName) {
        return
    }

    const sheetApi = await getSheetApi()

    try {
        const result = await sheetApi.spreadsheets.values.batchGet({
            spreadsheetId: spreadsheetId,
            ranges: [`'${sheetName}'!A1:A${maxRows}`, `'${sheetName}'!1:1`],
        });
        const idData = result?.data?.valueRanges?.[0]?.values
        const columns = result?.data?.valueRanges?.[1]?.values?.[0]
        return idData && columns ? {columns, idData} : null
    } catch (e) {
        console.error(e);
    }
}

async function getSheetIdData(spreadsheetId, sheetName) {
    if (!spreadsheetId || !sheetName) {
        return
    }

    const sheetApi = await getSheetApi()

    try {
        const result = await sheetApi.spreadsheets.values.batchGet({
            spreadsheetId: spreadsheetId,
            ranges: [`'${sheetName}'!A1:A${maxRows}`],
        });
        return result?.data?.valueRanges?.[0]?.values
    } catch (e) {
        console.error(e);
        return false
    }
}

async function createSheet(spreadsheetId, name) {
    const sheetApi = await getSheetApi()
    try {
        const result = await sheetApi.spreadsheets.batchUpdate({
            spreadsheetId: spreadsheetId,
            resource: {
                requests: [{
                    addSheet: {
                        properties: {
                            title: name,
                            gridProperties: {
                                rowCount: maxRows,
                                columnCount: maxColumns
                                // columnCount: columns.length + 1
                            },
                            tabColor: {
                                red: Math.random(),
                                green: Math.random(),
                                blue: Math.random()
                            }
                        }
                    }
                }]
            }
        });

        return true
    } catch (err) {
        console.error('Sheets API Error: ' + err);
        return false
    }
}


/*
*
* {
  spreadsheetId: "1Ktfyuaqvt6oiS014mSPoM7HmXJZFrExHk1h-OH5v1Hc",
  properties: {
    title: "facebbook product",
    locale: "vi_VN",
    autoRecalc: "ON_CHANGE",
    timeZone: "Asia/Saigon",
  },
  sheets: [
    {
      properties: {
        sheetId: 0,
        title: "Trang tính1",
        index: 0,
        sheetType: "GRID",
        gridProperties: {
          rowCount: 887,
          columnCount: 21
        }
      },
    }
  ],
  spreadsheetUrl: "https://docs.google.com/spreadsheets/d/1Ktfyuaqvt6oiS014mSPoM7HmXJZFrExHk1h-OH5v1Hc/edit"
}
* */
async function getSpreadInfo(spreadsheetId) {
    if (!spreadsheetId) {
        return
    }

    const sheetApi = await getSheetApi()

    try {
        const result = await sheetApi.spreadsheets.get({
            spreadsheetId: spreadsheetId,
        });
        return result?.data
    } catch (e) {
        console.error(e);
    }
}

async function isHasSheet(spreadsheetId, sheetName) {
    const spreadInfo = await getSpreadInfo(spreadsheetId)
    return spreadInfo?.sheets?.some(s => s?.properties?.title == sheetName)
}

async function fillRow(spreadsheetId, sheetName, rowIndex, values) {
    const params = {
        spreadsheetId,
        resource: {
            data: [{
                range: `'${sheetName}'!${rowIndex}:${rowIndex}`,
                values: [values]
            }],
            valueInputOption: "USER_ENTERED",
        },
    };

    const sheetApi = await getSheetApi()

    try {
        const result = await sheetApi.spreadsheets.values.batchUpdate(params);
        return result?.data
    } catch (e) {
        console.error(e);
    }

}

async function insertRows(spreadsheetId, sheetName, rows) {
    const sheetIdData = await getSheetDataInfo(spreadsheetId, sheetName)
    if (!sheetIdData?.columns) return false
    const {columns, idData} = sheetIdData
    let nextIndex = idData.length + 1

    const data = []
    for (let i = 0; i < rows.length; i++) {
        const item = rows[i]

        for (let columnIndex = 0; columnIndex < columns.length; columnIndex++) {
            const column = columns[columnIndex]
            let columnLabel = getColumnLabelFromIndex(columnIndex + 1)
            let value = item[column]
            if (value === null || value === undefined) {
                value = ""
            }
            data.push({
                range: `'${sheetName}'!${columnLabel}${nextIndex}:${columnLabel}${nextIndex}`,
                values: [[value]]
            })
        }

        nextIndex++
    }

    const params = {
        spreadsheetId,
        resource: {
            data,
            valueInputOption: "USER_ENTERED",
        },
    };

    const sheetApi = await getSheetApi()

    try {
        const result = await sheetApi.spreadsheets.values.batchUpdate(params);
        return result?.data
    } catch (e) {
        console.error(e);
    }

}

async function updateRows(spreadsheetId, sheetName, row, rowsIndex) {
    const sheetIdData = await getSheetDataInfo(spreadsheetId, sheetName)
    if (!sheetIdData?.columns) return false
    const {columns, idData} = sheetIdData

    const data = []
    for (let i = 0; i < rowsIndex.length; i++) {
        const index = rowsIndex[i]

        for (let columnIndex = 0; columnIndex < columns.length; columnIndex++) {
            const column = columns[columnIndex]
            let columnLabel = getColumnLabelFromIndex(columnIndex + 1)
            if (row.hasOwnProperty(column)) {
                let value = row[column]
                if (value === null || value === undefined) {
                    value = ""
                }
                data.push({
                    range: `'${sheetName}'!${columnLabel}${index}:${columnLabel}${index}`,
                    values: [[value]]
                })
            }
        }
    }

    if (!data.length) {
        return {
            success: false,
            msg: 'There is no row meets the condition'
        }
    }

    const params = {
        spreadsheetId,
        resource: {
            data,
            valueInputOption: "USER_ENTERED",
        },
    };

    const sheetApi = await getSheetApi()

    try {
        const result = await sheetApi.spreadsheets.values.batchUpdate(params);
        return result?.data ? {
            success: true,
            totalUpdatedRows: result?.data?.totalUpdatedRows,
        } : {
            success: false,
        }
    } catch (e) {
        console.error(e);
        return {
            success: false,
        }
    }

}

async function deleteRows(spreadsheetId, sheetName, rowsIndex) {
    const spreadInfo = await getSpreadInfo(spreadsheetId)
    if (!spreadInfo) return
    const sheetId = spreadInfo.sheets.find(s => s.properties.title === sheetName)?.properties?.sheetId
    if (!sheetId) return
    rowsIndex.sort()
    const sheetApi = await getSheetApi()

    try {
        let offsetIndex = 0
        const result = await sheetApi.spreadsheets.batchUpdate({
            spreadsheetId,
            resource: {
                requests: rowsIndex.map(function (index) {
                    const deleteDimension = {
                        deleteDimension: {
                            range: {
                                sheetId: sheetId,
                                dimension: "ROWS",
                                startIndex: index - 1 - offsetIndex,
                                endIndex: index - offsetIndex
                            }
                        }
                    };
                    offsetIndex++
                    return deleteDimension;
                })
            }
        });
        return result?.data
    } catch (e) {
        console.error(e);
    }

}


module.exports = {
    generateAuthUrl,
    updateCode,
    getSpreadData,
    createSheet,
    getSpreadInfo,
    isHasSheet,
    fillRow,
    insertRows,
    updateRows,
    deleteRows,
}