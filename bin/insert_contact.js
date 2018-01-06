/* eslint-disable no-console */
import XLSX from 'xlsx';
import Contact from '../models/Contact';

// Constant
const FILE_NAME = 'test.xlsx';
const VALID_SHEET_NAME = new Set(['捐款人', 'MKT']);
const HEADER_LOOKUP = {
    'email': 'email',
    '中文全名': 'name',
    '地址': 'address',
    '身份別': 'identity',
    'nickname': 'nickname',
    '單位': 'unit',
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
            const contact = parseRow(row, indexLookup);

            // FIXME Filter invalid data
            if (contact.name !== '地址貼上的名字') {
                result.push(parseRow(row, indexLookup));
            }
        }

        indexLookup = (indexLookup || getIndexFromHeader(row));
    }

    return result;
}

function mergeByEmailAndName(rows) {
    const numRows = rows.length;
    let emailLookup = {},
        noEmailRows = [];

    for (var i = 0; i < numRows; i++) {
        const row = rows[i];
        const email = row.email;
        const name = row.name;

        if (!email && !name) {
            // Collect no email cells
            noEmailRows.push(row);

            continue;
        }

        const key = `${email}/${name}`;
        const previousRow = emailLookup[key];

        if (previousRow) {
            emailLookup[key] = previousRow.merge(row);
        } else {
            emailLookup[key] = row;
        }
    }

    return [
        ...noEmailRows,
        ...Object.values(emailLookup),
    ];
}

function mergeByNameAndAddress(rows) {
    const numRows = rows.length;
    let result = [];
    let nameAddressLookup = {};

    for (var i = 0; i < numRows; i++) {
        const row = rows[i];
        const nameAddressList = row.nameAddressList;
        const numNameAddress = nameAddressList.length;

        if (numNameAddress === 0) {
            result.push(row);

            continue;
        }

        let mergedRowIndex;
        for (let j = 0; j < numNameAddress; j++) {
            mergedRowIndex = nameAddressLookup[nameAddressList[j]];

            if (mergedRowIndex !== undefined) {
                // merge to previous row
                const previousRow = result[mergedRowIndex];

                if ((previousRow.email === null) || (row.email === null)) {
                    result[mergedRowIndex] = previousRow.merge(row);
                }

                break;
            }
        }

        if (mergedRowIndex === undefined) {
            // create new row
            const currentIndex = result.length;

            result.push(row);
            // update nameAddressLookup
            for (let j = 0; j < numNameAddress; j++) {
                nameAddressLookup[nameAddressList[j]] = currentIndex;
            }
        }
    }

    return result;
}

function parseSheets(sheets) {
    let rawData = [];

    for (var sheetName in sheets) {
        if (!VALID_SHEET_NAME.has(sheetName)) {
            continue;
        }

        rawData = [
            ...rawData,
            ...parseNormalSheet(wb.Sheets[sheetName])
        ];
    }

    return rawData;
}

const wb = XLSX.readFile(FILE_NAME);
const result = mergeByNameAndAddress(mergeByEmailAndName(parseSheets(wb.Sheets)));
const outputWS = XLSX.utils.json_to_sheet(
    result.map((row) => (
        row.toJSON()
    )),
    { header: [
        'email',
        'name',
        'addresses',
        'identities',
        'nicknames',
        'units',
        'departments',
        'paperCard',
        'annulReport',
        'annulReceipt'
    ]}
);
const outputWB = {
    SheetNames: ['Sheet1'],
    Sheets: {
        Sheet1: outputWS
    }
};

XLSX.writeFile(outputWB, './output.xlsx');
