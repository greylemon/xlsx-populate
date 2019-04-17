"use strict";
const XlsxPopulate = require('../lib/XlsxPopulate');
const MAX_ROW = 1048576, MAX_COLUMN = 16384;

let t = Date.now();
let parser, depParser, wb, rt;

/**
 *
 * @param {Workbook} workbook
 */
function something(workbook) {
    wb = workbook;
    console.log(`open workbook uses ${Date.now() - t}ms`);
    t = Date.now();
    // console.log(workbook.sheet('Fin_Summary').cell('I11').formula());
    // console.log(workbook.sheet('Act_Summary').cell('H3176').formula());
    // workbook.sheet('Act_Summary').cell('H11').setValue(123);
    // console.log(workbook.sheet('Act_Summary').cell('I11').getValue());
    // console.log(workbook.theme().themeColor(1));
    // console.log(workbook.sheet(0).getCell(3, 3).getValue());
    const cell = workbook.sheet(0).cell('AD18');
    const cell2 = workbook.sheet(0).cell('A1');
    cell.setValue(2020);
    console.log(cell2.getValue())

    console.log(`process formulas uses ${Date.now() - t}ms, with ? formulas, query data uses ??ms`);
    t = Date.now();
    // get data
    // console.log(JSON.stringify(rt._data).length)

    console.log(`process formulas uses ${Date.now() - t}ms`);
}

setTimeout(() => {
    t = Date.now();
    XlsxPopulate.fromFileAsync("./test/test.xlsm").then(something);
}, 0);

