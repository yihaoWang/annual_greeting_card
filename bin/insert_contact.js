/* eslint-disable no-console */
import XLSX from 'xlsx';
import Contact from '../models/Contact';

// Constant
const HEADER_LOOKUP = {
    '身份別': 'identity',
    '中文全名': 'name',
    'nickname': 'nickname',
    '單位名稱': 'unit',
    '部門名稱': 'department',
    'email': 'email',
    '地址': 'address',
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

    return Contact.fromObject(result);
}

function parseNormalSheet(ws) {
    const rows = sheet2arr(ws);
    const numRow = rows.length;
    let indexLookup,
        result = [];

    for (var i = 0; i < numRow; i++) {
        const row = rows[i];

        if (indexLookup) {
            try {
                result.push(parseRow(row, indexLookup));
            } catch(err) {
                console.log(`Invalid: ${JSON.stringify(row)}`);
            }
        }

        indexLookup = (indexLookup || getIndexFromHeader(row));
    }

    return result;
}

function mergeRows(rows) {
    const numRows = rows.length;
    let nameLookup = {},
        noNameRows = [];

    for (let i = 0; i < numRows; i++) {
        const row = rows[i];
        const name = row.name;
        let matchRows = nameLookup[name];

        if (!name) {
            noNameRows.push(row);

            continue;
        }
        if (!matchRows) {
            nameLookup[name] = [row];

            continue;
        }

        const email = row.email;
        const address = row.address;
        const numMatchRows = matchRows.length;

        let mergedIndex = numMatchRows;
        for (let j = 0; j < numMatchRows; j++) {
            const matchRow = matchRows[j];
            const matchRowEmail = matchRow.email;
            const matchRowAddress = matchRow.address;

            if (
                (email ===  matchRowEmail) &&
                (address === matchRowAddress)
            ) {
                mergedIndex = j;

                break;
            } else if (
                (email === matchRowEmail) &&
                (address !== matchRowAddress)
                // handle ER
            ) {
                mergedIndex = j;

                break;
            } else if (
                (!email || !matchRowEmail) &&
                (address === matchRowAddress)
            ) {
                mergedIndex = j;

                break;
            }
        }

        const mergedRow = matchRows[mergedIndex];

        matchRows[mergedIndex] = mergedRow ? mergedRow.merge(row) : row;
        nameLookup[name] = matchRows;
    }

    return noNameRows.concat.apply(noNameRows, Object.values(nameLookup));
}

function parseSheets(sheets) {
    let rawData = [];

    for (var sheetName in sheets) {
        rawData = [
            ...rawData,
            ...parseNormalSheet(wb.Sheets[sheetName])
        ];
    }

    return rawData;
}

const wb = XLSX.readFile(process.argv[2]);
const result = mergeRows(parseSheets(wb.Sheets));

const outputWS = XLSX.utils.json_to_sheet(
    result.map((row) => (
        row.toJSON()
    )),
    { header: [
        'identities',
        'name',
        'nicknames',
        'units',
        'departments',
        'email',
        'addresses',
        'paperCard',
        'annulReport',
        'annulReceipt'
    ]}
);
const outputWB = {
    SheetNames: ['all'],
    Sheets: {
        all: outputWS
    }
};

XLSX.writeFile(outputWB, './output.xlsx');
