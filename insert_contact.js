/* eslint-disable no-console */
const XLSX = require('xlsx');

// Constant
const FILE_NAME = 'test.xlsx';
const VALID_SHEET_NAME = new Set(['捐款人']);

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

const wb = XLSX.readFile(FILE_NAME);

for (var sheetName in wb.Sheets) {
    if (!VALID_SHEET_NAME.has(sheetName)) {
        continue;
    }

    const ws = sheet2arr(wb.Sheets[sheetName]);

    console.log(ws);
}
