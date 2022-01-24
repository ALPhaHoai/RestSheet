var express = require('express');
var router = express.Router();
const fs = require('fs');
const readline = require('readline');
const {google} = require('googleapis');
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
const TOKEN_PATH = 'token.json';
const {spreadsheets} = require("../config")

let auth

async function getAuth() {
    if (auth) return auth
    return new Promise((resolve, reject) => {
        authorize(function (oAuth2Client) {
            auth = oAuth2Client
            resolve(auth)
        })
    })
}

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


async function getSheetData(spreadsheetId, sheetName) {
    const auth = await getAuth()
    const sheets = google.sheets({version: 'v4', auth});
    return new Promise(resolve => {
        sheets.spreadsheets.values.get({
            spreadsheetId: spreadsheetId,
            range: `'${sheetName}!'A1:ZZ1000`,
        }, (err, res) => {
            if (err) {
                resolve()
            } else {
                resolve(res.data.values)
            }
        });
    })
}

function getColumnLabelFromIndex(index) {
    try {
        const str = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        return str[index - 1]
    } catch (e) {
        return str[str.length]
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

async function updateProducts(productList) {
    const auth = await getAuth()
    const sheets = google.sheets({version: 'v4', auth});

    for (let i1 = 0; i1 < spreadsheets.length; i1++) {
        let {
            spreadsheetId,
            mainSheetName,
            basePriceColumn,
            promotePriceColumn,
            productStockColumn,
            inventoryStockColumn,
            productImageColumn,
            productNameColumn,
            tradeMarkNameColumn,
            availability_InstockName,
            availability_SoldoutName,
            productImageColumnWidth,
            productImageColumnMode,
            productImageColumnHeight,
            promotePriceFormat,
            productWebLinkColumn,
            productSpecificationColumn,
        } = spreadsheets[i1];

        const spreadData = await getSpreadData(spreadsheetId);
        if (!Array.isArray(spreadData) || !spreadData.length) {
            console.log("Cant get spreadData " + spreadsheetId);
            continue
        }

        const updateData = []
        for (let i = 0; i < spreadData.length; i++) {
            const sheetInfo = spreadData[i]
            let sheetData = sheetInfo.data.values
            const sheetName = sheetInfo.properties.title
            if (!Array.isArray(sheetData) || !sheetData.length) continue//sheet is empty
            const basePriceColumnLabel = getColumLabel(sheetData, basePriceColumn)
            const promotePriceColumnLabel = getColumLabel(sheetData, promotePriceColumn)
            const productStockColumnLabel = getColumLabel(sheetData, productStockColumn)
            const inventoryStockColumnLabel = getColumLabel(sheetData, inventoryStockColumn)
            const productImageColumnLabel = getColumLabel(sheetData, productImageColumn)
            const productNameColumnLabel = getColumLabel(sheetData, productNameColumn)
            const productWebLinkColumnLabel = getColumLabel(sheetData, productWebLinkColumn)
            const productSpecificationColumnLabel = getColumLabel(sheetData, productSpecificationColumn)
            const tradeMarkNameColumnLabel = getColumLabel(sheetData, tradeMarkNameColumn)

            let endRow = sheetData.length
            for (let i = 0; i < productList.length; i++) {
                try {
                    let {
                        product_code,
                        productName,
                        basePrice,
                        promote_price,
                        productStock,
                        kiotvietImages,
                        categoryName,
                        tradeMarkName,
                        ProductWebLink,
                        productSpecification,
                    } = productList[i]

                    if (mainSheetName) {
                        if (mainSheetName !== sheetName) continue
                    } else {
                        if (tradeMarkName !== sheetName) continue
                    }

                    let productRowIndex = getProductSheetRowIndex(sheetData, product_code)
                    const isInsertNewRow = typeof productRowIndex === "undefined"
                    if (isInsertNewRow) {
                        //insert new row
                        productRowIndex = ++endRow

                        updateData.push({
                            range: `'${sheetName}'!` + "A" + productRowIndex,
                            values: [[product_code]]
                        })
                    }

                    if (productName !== undefined && productNameColumnLabel) {
                        updateData.push({
                            range: `'${sheetName}'!` + productNameColumnLabel + productRowIndex,
                            values: [[productName]]
                        })
                    }

                    if (ProductWebLink !== undefined && productWebLinkColumnLabel) {
                        updateData.push({
                            range: `'${sheetName}'!` + productWebLinkColumnLabel + productRowIndex,
                            values: [[ProductWebLink]]
                        })
                    }

                    if (productSpecification !== undefined && productSpecificationColumnLabel) {
                        updateData.push({
                            range: `'${sheetName}'!` + productSpecificationColumnLabel + productRowIndex,
                            values: [[productSpecification]]
                        })
                    }

                    if (basePrice !== undefined && basePriceColumnLabel) {
                        if (basePrice === null || basePrice === "") basePrice = 0;
                        updateData.push({
                            range: `'${sheetName}'!` + basePriceColumnLabel + productRowIndex,
                            values: [[basePrice + " VND"]]
                        })
                    }
                    if (promote_price !== undefined && promotePriceColumnLabel) {
                        if (promote_price === null || promote_price === "") promote_price = 0;
                        let value = promote_price + " VND"

                        if (typeof promotePriceFormat === "function") {
                            value = promotePriceFormat(promote_price)
                        }

                        updateData.push({
                            range: `'${sheetName}'!` + promotePriceColumnLabel + productRowIndex,
                            values: [[value]]
                        })
                    }
                    if (tradeMarkName !== undefined && tradeMarkNameColumnLabel) {
                        updateData.push({
                            range: `'${sheetName}'!` + tradeMarkNameColumnLabel + productRowIndex,
                            values: [[tradeMarkName]]
                        })
                    }
                    if (Array.isArray(kiotvietImages) && productImageColumnLabel) {
                        let value = kiotvietImages[0] || ""
                        if (value) {
                            if (productImageColumnMode === 4) {
                                value = `=IMAGE("${value}"; 4; ${productImageColumnWidth}; ${productImageColumnHeight})`
                            } else {
                                value = `=IMAGE("${value}"; ${productImageColumnMode})`
                            }
                        }
                        updateData.push({
                            range: `'${sheetName}'!` + productImageColumnLabel + productRowIndex,
                            values: [[value]],
                        })
                    }
                    if (!isNaN(productStock)) {
                        if (productStockColumnLabel) {
                            const value = productStock > 0 ? availability_InstockName : availability_SoldoutName
                            updateData.push({
                                range: `'${sheetName}'!` + productStockColumnLabel + productRowIndex,
                                values: [[value]]
                            })
                        }
                        if (inventoryStockColumnLabel) {
                            updateData.push({
                                range: `'${sheetName}'!` + inventoryStockColumnLabel + productRowIndex,
                                values: [[productStock]]
                            })
                        }
                    }
                } catch (e) {
                    console.log(e);
                }
            }
        }
        const params = {
            spreadsheetId,
            resource: {
                data: updateData,
                valueInputOption: "USER_ENTERED",
            },
        };
        if (updateData.length) {
            await batchUpdateAsync(params, sheets)
        }
    }
}

function batchUpdateAsync(params, sheets, sleepTime = 1) {
    return new Promise(resolve => {
        batchUpdateCallback(params, sheets, function (err) {
            resolve(err)
        }, sleepTime)
    })
}

function batchUpdateCallback(params, sheets, callback, sleepTime = 1) {
    console.log("spreadsheets batchUpdate", sleepTime);
    sheets.spreadsheets.values.batchUpdate(params, (err, result) => {
        if (err) {
            console.log(err);
            if (sleepTime > 10) {
                callback(err)
            } else {
                setTimeout(function () {
                    batchUpdateCallback(params, sheets, callback, sleepTime * 2)
                }, sleepTime * 1000)
            }
        } else {
            callback()
        }
    });
}

function getProductSheetRowIndex(data, product_code) {
    for (let i = 0; i < data.length; i++) {
        const row = data[i]
        if (row[0] == product_code) {
            return (i + 1)
        }
    }
}

async function getSpreadData(spreadsheetId) {
    if (!spreadsheetId) {
        return
    }

    let spreadInfo = await getSpreadInfo(spreadsheetId);
    if (!Array.isArray(spreadInfo.sheets)) return
    const ranges = []

    for (let i = 0; i < spreadInfo.sheets.length; i++) {
        const {properties: {sheetId, title, index, sheetType}} = spreadInfo.sheets[i]
        ranges.push(`'${title}'!A1:ZZ1000`)
    }

    const auth = await getAuth()
    const sheets = google.sheets({version: 'v4', auth});

    return new Promise(resolve => {
        getSpreadDataCallback(sheets, spreadInfo, ranges, spreadsheetId, function (data) {
            resolve(data)
        })
    })
}

function getSpreadDataCallback(sheets, spreadInfo, ranges, spreadsheetId, callback, retry = 1) {
    if (!spreadsheetId) {
        callback()
        return
    }

    if (!spreadInfo || !Array.isArray(spreadInfo.sheets)) {
        callback()
        return
    }

    sheets.spreadsheets.values.batchGet({
        spreadsheetId: spreadsheetId,
        ranges,
    }, (err, res) => {
        if (err) {
            if (retry > 10) {
                callback()
            } else {
                setTimeout(function () {
                    getSpreadDataCallback(sheets, spreadInfo, ranges, spreadsheetId, callback, retry + 1)
                }, retry * 1000)
            }
        } else {
            for (let i = 0; i < spreadInfo.sheets.length; i++) {
                spreadInfo.sheets[i].data = res.data.valueRanges[i]
            }
            callback(spreadInfo.sheets)
        }
    });
}


/*
*
* {
  "spreadsheetId": "1Ktfyuaqvt6oiS014mSPoM7HmXJZFrExHk1h-OH5v1Hc",
  "properties": {
    "title": "facebbook product",
    "locale": "vi_VN",
    "autoRecalc": "ON_CHANGE",
    "timeZone": "Asia/Saigon",
  },
  "sheets": [
    {
      "properties": {
        "sheetId": 0,
        "title": "Trang tính1",
        "index": 0,
        "sheetType": "GRID",
        "gridProperties": {
          "rowCount": 887,
          "columnCount": 21
        }
      },
    }
  ],
  "spreadsheetUrl": "https://docs.google.com/spreadsheets/d/1Ktfyuaqvt6oiS014mSPoM7HmXJZFrExHk1h-OH5v1Hc/edit"
}
* */
async function getSpreadInfo(spreadsheetId) {
    if (!spreadsheetId) {
        return
    }

    const auth = await getAuth()
    const sheets = google.sheets({version: 'v4', auth});

    return new Promise(resolve => {
        getSpreadInfoCallback(sheets, spreadsheetId, function (data) {
            resolve(data)
        })
    })
}

function getSpreadInfoCallback(sheets, spreadsheetId, callback, retry = 1) {
    if (!spreadsheetId) {
        callback()
    } else {
        sheets.spreadsheets.get({
            spreadsheetId: spreadsheetId,
        }, (err, res) => {
            if (err) {
                if (retry > 10) {
                    callback()
                } else {
                    setTimeout(function () {
                        getSpreadInfoCallback(sheets, spreadsheetId, callback, retry + 1)
                    }, retry * 1000)
                }
            } else callback(res.data)
        });
    }
}

async function getProducstCode() {
    let data = []
    for (let i = 0; i < spreadsheets.length; i++) {
        let {spreadsheetId} = spreadsheets[i];
        const sheets = await getSpreadData(spreadsheetId)
        if (Array.isArray(sheets)) {
            for (let j = 0; j < sheets.length; j++) {
                const values = sheets[j].data.values
                for (let k = 0; k < values.length; k++) {
                    const value = values[k]
                    if (value && value[0] && value[0].toLocaleLowerCase() !== "id") {
                        data.push(value[0])
                    }
                }
            }
        }
    }
    return data
}

module.exports = {
    generateAuthUrl,
    updateCode,
    updateSpreadsheetId: function (_sheetId) {
        // spreadsheetId = _sheetId
    },
    getProducstCode,
    updateProducts,
    getSpreadData,
}