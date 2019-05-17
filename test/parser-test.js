"use strict";
const fs = require('fs');
const XlsxPopulate = require('../lib/XlsxPopulate');

const files = [];
const extensions = ['.xlsx', '.xlsm', '.xlsb'];

describe('Read test', () => {
    const fileNames = fs.readdirSync('./excels');
    fileNames.forEach(file => {
        const i = file.lastIndexOf('.');
        if (i < 0) return;
        if (extensions.includes(file.slice(i)) && file.indexOf('out') === -1) {
            files.push(file);
        }
    });
    files.forEach(file => {
        it(`should read ${file}`, done => {
            XlsxPopulate.fromFileAsync('./excels/' + file)
                .then(workbook => {
                    done();
                    return workbook;
                });
            // .then(workbook => {
            //     workbook.toFileAsync(`${file.slice(0, file.lastIndexOf('.'))}.out${file.slice(file.lastIndexOf('.'))}`);
            //     done();
            // });
        });
    });
});

// it('should read file', function (done) {
//     XlsxPopulate.fromFileAsync('./excels/' + '2017-18 Q4 Community LHIN Managed BLANK V1.xlsm')
//         .then(workbook => {
//             done();
//             return workbook;
//         });
// });
