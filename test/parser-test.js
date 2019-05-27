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
                    workbook.sheet(0).cell(1, 1).style('bold');
                    return workbook;
                })
                .then(workbook => {
                    workbook.toFileAsync(`./excels/out/${file.slice(0, file.lastIndexOf('.'))}.out${file.slice(file.lastIndexOf('.'))}`)
                        .then(() => {
                            done();
                        });
                });
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
