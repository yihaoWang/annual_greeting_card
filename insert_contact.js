/* eslint-disable no-console */
const XLSX = require('xlsx');

// Constant
const FILE_NAME = 'test.xlsx';
const VALID_SHEET_NAME = new Set(['捐款人']);
const HEADER_LOOKUP = {
    '中文全名': 'name',
    'nickname': 'nickname',
    '地址': 'address',
    'email': 'email',
    '紙本賀卡': 'paperCard',
    '年度報告': 'annualReport',
    '年度收據': 'annualReceipt',
};

function sheet2arr(sheet) {
    var result = [];
    var range = XLSX.utils.decode_range(sheet['!ref']);

    for (var rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
        const row = [];

        for (var colNum = range.s.c; colNum <= range.e.c; colNum++) {
            var nextCell = sheet[
                XLSX.utils.encode_cell({ r: rowNum, c: colNum })
            ];

            if (typeof nextCell === 'undefined') {
                row.push(undefined);
            } else {
                row.push(nextCell.w);
            }
        }

        const isValidRow = row.find((value) => {
            return (value !== undefined);
        });

        if (isValidRow) {
            result.push(row);
        }
    }

    return result;
}

function getIndexFromHeader(row) {
    let result = {};

    for (var key in HEADER_LOOKUP) {
        const index = row.indexOf(key);

        if (index === -1) {
            return false;
        }

        result[HEADER_LOOKUP[key]] = index;
    }

    return result;
}

function parseRow(row, indexLookup) {
    let result = {};

    for (var key in indexLookup) {
        const index = indexLookup[key];

        result[key] = row[index];
    }

    return result;
}

function parseSheet(ws) {
    const rows = sheet2arr(ws);
    const numRow = rows.length;
    let indexLookup,
        result = [];

    for (var i = 0; i < numRow; i++) {
        const row = rows[i];

        indexLookup = indexLookup || getIndexFromHeader(row);

        if (!indexLookup) {
            continue;
        }

        result.push(parseRow(row, indexLookup));
    }

    return result;
}

const wb = XLSX.readFile(FILE_NAME);

for (var sheetName in wb.Sheets) {
    if (!VALID_SHEET_NAME.has(sheetName)) {
        continue;
    }

    const result = parseSheet(wb.Sheets[sheetName]);

    console.log(result);
}
