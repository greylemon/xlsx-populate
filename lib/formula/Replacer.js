/* eslint-disable */
"use strict";
const ac = require("../addressConverter");

// Prevent using lookbehind since only chrome supports it. (?<!)
// match, sheetMatch, sheet1, sheet2, isRange, absCol, colName, absRow, row, isFunction
const cellReference = /((?:([A-Za-z_.\d\u007F-\uFFFF]+)|'((?:(?![\\\/\[\]*?:]).)+?)')!)?(:)?(\$)?([A-Za-z]+)(\$)?([0-9]+)(\()?(?!:)/g;

// match, sheetMatch, sheet1, sheet2, absCol1, colName1, absRow1, row1, absCol2, colName2, absRow2, row2
const rangeReference = /((?:([A-Za-z_.\d\u007F-\uFFFF]+)|'((?:(?![\\\/\[\]*?:]).)+?)')!)?(\$)?([A-Za-z]+)(\$)?([0-9]+):(\$)?([A-Za-z]+)(\$)?([0-9]+)/g;

// match, sheetMatch, sheet1, sheet2, absRow1, row1, absRow2, row2
const RowReference = /((?:([A-Za-z_.\d\u007F-\uFFFF]+)|'((?:(?![\\\/\[\]*?:]).)+?)')!)?(\$)?([0-9]+):(\$)?([0-9]+)/g;

// match, sheetMatch, sheet1, sheet2, absCol1, colName1, absCol2, colName2
const ColumnReference = /((?:([A-Za-z_.\d\u007F-\uFFFF]+)|'((?:(?![\\\/\[\]*?:]).)+?)')!)?(\$)?([A-Za-z]+):(\$)?([A-Za-z]+)(\$)?/g;

module.exports = {
    /**
     * Remove/Replace corresponding part (sheetName, row, col) in the given formula.
     * For cell reference , replace it with #REF!;
     * For range reference,
     * @param {string} formula - the formula to replace
     * @param {string} currSheetName - the sheet that the formula in
     * @param {string} sheetName - cell reference
     * @param {number} row - cell reference
     * @param {number} col - cell reference
     */
    removeReference: (formula, currSheetName, sheetName, row, col) => {

    },

    replaceRowNumber: (formula, currSheetName, sheetName, oldRowNum, newRowNum) => {
        const diff = oldRowNum - newRowNum;
        formula = formula.replace(cellReference, (match, sheetMatch, sheet1, sheet2, isRange, absCol, colName, absRow, row, isFunction) => {
            if (isFunction || isRange) {
                return match;
            }
            const sheet = sheet1 ? sheet1 : (sheet2 ? sheet2 : currSheetName);
            if (sheet !== sheetName) {
                return match;
            }
            if (oldRowNum === +row) {
                row = newRowNum;
            }
            return (sheetMatch ? sheetMatch : '') + (absCol ? '$' : '') + colName + (absRow ? '$' : '') + row;
});

        formula = formula.replace(rangeReference, (match, sheetMatch, sheet1, sheet2, absCol1, colName1, absRow1, row1, absCol2, colName2, absRow2, row2) => {
            const sheet = sheet1 ? sheet1 : (sheet2 ? sheet2 : currSheetName);
            if (sheet !== sheetName) {
                return match;
            }
            row1 = Math.min(row1, row2);
            row2 = Math.max(row1, row2);
            if (row1 <= oldRowNum <= row2 && diff > 0 && Math.abs(row2 - row1) <= diff) return '#REF!';
            if (row1 <= oldRowNum <= row2) {
                row2 -= diff;
            }
            return (sheetMatch ? sheetMatch : '') + (absCol1 ? '$' : '') + colName1 + (absRow1 ? '$' : '') + row1 + ':'
            + (absCol2 ? '$' : '') + colName2 + (absRow2 ? '$' : '') + row2;
        });

        formula = formula.replace(RowReference, (match, sheetMatch, sheet1, sheet2, absRow1, row1, absRow2, row2) => {
            const sheet = sheet1 ? sheet1 : (sheet2 ? sheet2 : currSheetName);
            if (sheet !== sheetName) {
                return match;
            }
            if (oldRowNum === +row1) row1 = newRowNum;
            if (oldRowNum === +row2) row2 = newRowNum;
            return (sheetMatch ? sheetMatch : '') + (absRow1 ? '$' : '') + row1 + ':' + (absRow2 ? '$' : '') + row2;
        });
        return formula;
    }
};
