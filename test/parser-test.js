"use strict";
const fs = require('fs');
const XlsxPopulate = require('../lib/XlsxPopulate');

const files = [];
const extensions = ['.xlsx', '.xlsm', '.xlsb'];

describe('Read test', () => {
    const fileNames = fs.readdirSync('./');
    fileNames.forEach(file => {
        const i = file.lastIndexOf('.');
        if (i < 0) return;
        if (extensions.includes(file.slice(i)) && file.indexOf('out') === -1) {
            files.push(file);
        }
    });
    files.forEach(file => {
        it(`should read ${file}`, done => {
            XlsxPopulate.fromFileAsync(file)
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
