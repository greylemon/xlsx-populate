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
    //
    // console.log('...');
    //
    // const cell1 = workbook.sheet('Fin_CMHP1').cell('E10');
    // const cell2 = workbook.sheet('Fin_CMHP1').cell('F10');
    // console.log(cell1.setValue(123));
    // console.log(cell2.value());
    // console.log(cell2.setValue(12));
    // console.log(cell1.setValue(1234));
    // console.log(cell2.value()); // should be 12
    //
    // console.log('...');
    // const a1 = workbook.sheet(0).cell('A1');
    // const a2 = workbook.sheet(0).cell('A2');
    // const a3 = workbook.sheet(0).cell('A3');
    // a1.value(1);
    // a2.value(2);
    // a3.setFormula('A1+A2');
    // console.log(a3.value());
    // a2.value(3);
    // console.log(a3.value());

    // console.log('...');
    // const a4 = workbook.sheet(0).cell('A4');
    // const a5 = workbook.sheet(0).cell('A5');
    // a4.formula('1233');
    // a5.setFormula('A4');
    // console.log(a5.value());



    // const cell = workbook.sheet(0).cell('AD18');
    // const cell2 = workbook.sheet(0).cell('A10');
    // const cell3 = workbook.sheet(0).cell('L3');
    // cell.setValue(2009);
    // console.log(cell2.getValue())
    // console.log(cell3.getValue())

    // workbook.sheet('Main').hyperlink('')

    console.log(`process formulas uses ${Date.now() - t}ms, with ? formulas, query data uses ??ms`);
    t = Date.now();
    // get data
    // console.log(JSON.stringify(rt._data).length)

    console.log(`process formulas uses ${Date.now() - t}ms`);
}
const file = ['./TF33674675.xlsx', './test.xlsm']
setTimeout(() => {
    t = Date.now();
    XlsxPopulate.fromFileAsync(file[0])
        .then(something)
        .then(() => {
            wb.toFileAsync('./out.xlsx');
        });

}, 0);

