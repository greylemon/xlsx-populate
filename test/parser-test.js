"use strict";
const fs = require('fs');
const XlsxPopulate = require('../lib/XlsxPopulate');

let files = [];
const extensions = ['.xlsx', '.xlsm', '.xlsb'];

describe('Read test', function () {
    this.timeout(15000);
    const fileNames = fs.readdirSync('./excels/');
    fileNames.forEach(file => {
        const i = file.lastIndexOf('.');
        if (i < 0) return;
        if (extensions.includes(file.slice(i)) && file.indexOf('out') === -1) {
            files.push(file);
        }
    });
    // files = ['formula-test.xlsx', 'TF33674675.xlsx'];
    files.forEach(file => {
        it(`should read ${file}`, async () => {
            const workbook = await XlsxPopulate.fromFileAsync('./excels/' + file);
            for (let i = 1; i < 100; i++) {
                workbook.sheet(0).getCell(i, i).setStyle('bold', true)
            }
            const outputName1 = `./excels/out/${file.slice(0, file.lastIndexOf('.'))}.out${file.slice(file.lastIndexOf('.'))}`;
            const outputName2 = `./excels/out/${file.slice(0, file.lastIndexOf('.'))}.out2${file.slice(file.lastIndexOf('.'))}`;
            await workbook.toFileAsync(outputName1);
            await workbook.toFileAsync(outputName2);
            // const workbook2 = await XlsxPopulate.fromFileAsync(outputName1);
            // await workbook2.toFileAsync(outputName2);
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
