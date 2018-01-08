/* eslint-disable no-console */
import fs from 'fs';
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

let logs = [];

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

function parseRow(row, indexLookup, isEr = false) {
    let result = {};

    for (var key in indexLookup) {
        const index = indexLookup[key];

        result[key] = row[index];
    }

    result.er = isEr;

    return Contact.fromObject(result);
}

function parseSheet(ws, isEr) {
    const rows = sheet2arr(ws);
    const numRow = rows.length;
    let indexLookup,
        result = [];

    for (var i = 0; i < numRow; i++) {
        const row = rows[i];

        if (indexLookup) {
            try {
                result.push(parseRow(row, indexLookup, isEr));
            } catch(err) {
                logs.push(`Invalid row: ${JSON.stringify(row)}`);
            }
        }

        indexLookup = (indexLookup || getIndexFromHeader(row));
    }

    return result;
}

function mergeRows(rows) {
    const numRows = rows.length;
    let nameLookup = {},
        noNameRows = [],
        mergeCount = 0;

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
        const address = (row.addresses.size === 0) ? null : [...row.addresses][0];
        const er = row.er;
        const numMatchRows = matchRows.length;

        let mergedIndex = numMatchRows,
            mergeReason;
        for (let j = 0; j < numMatchRows; j++) {
            const matchRow = matchRows[j];
            const matchRowEmail = matchRow.email;
            const matchRowAddresses = matchRow.addresses;
            const matchRowEr = matchRow.er;

            if (
                (email ===  matchRowEmail) &&
                (matchRowAddresses.has(address))
            ) {
                mergedIndex = j;
                mergeReason = 'Merge rows by name, eamil and address';

                break;
            } else if (
                (email === matchRowEmail) &&
                (!matchRowAddresses.has(address)) &&
                (!er && !matchRowEr)
            ) {
                mergedIndex = j;
                mergeReason = 'Merge rows by name and email';

                break;
            } else if (
                (!email || !matchRowEmail) &&
                (matchRowAddresses.has(address))
            ) {
                mergedIndex = j;
                mergeReason = 'Merge rows by name and address';

                break;
            }
        }

        const mergedRow = matchRows[mergedIndex];

        if (mergedRow) {
            mergeCount++;
            matchRows[mergedIndex] = mergedRow.merge(row);

            logs.push(`${mergeReason}: ${JSON.stringify(row.toJSON())} ${JSON.stringify(mergedRow.toJSON())}`);
        } else {
            matchRows[mergedIndex] = row;
        }

        nameLookup[name] = matchRows;
    }
    console.log(`${mergeCount} rows are merged`);

    return noNameRows.concat.apply(noNameRows, Object.values(nameLookup));
}

function parseSheets(sheets) {
    let result = [];

    for (var sheetName in sheets) {
        const rawData = parseSheet(wb.Sheets[sheetName], (sheetName === 'ER'));

        result = result.concat(rawData);
        console.log(`Find ${rawData.length} valid rows from ${sheetName}`);
    }


    return result;
}

function createSheet(rows) {
    return XLSX.utils.json_to_sheet(
        rows.map((row) => (row.toJSON())),
        { header: [
            'identities',
            'email',
            'name',
            'nicknames',
            'addresses',
            'units',
            'departments',
            'paperCard',
            'annulReport',
            'annulReceipt'
        ]}
    );
}

function generateOutput(rows, outputName) {
    const outputWB = {
        SheetNames: ['ER','Other'],
        Sheets: {
            ER: createSheet(rows.filter((row) => (row.er))),
            Other: createSheet(rows.filter((row) => (!row.er))),
        }
    };

    XLSX.writeFile(outputWB, outputName);
}

function writeLogs(logs) {
    fs.open('log.txt', 'w', (err, fd) => {
        fs.writeSync(fd, logs.join('\n'));
        fs.close(fd);
    });
}

const wb = XLSX.readFile(process.argv[2]);
const result = mergeRows(parseSheets(wb.Sheets));

generateOutput(
    result.sort((a, b) => ([...a.identities][0] > [...b.identities][0])),
    process.argv[3] || './output.xlsx'
);
writeLogs(logs);
