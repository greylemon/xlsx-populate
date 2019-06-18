/* eslint-disable */
"use strict";
const ac = require("../addressConverter");

// Prevent using lookbehind since only chrome supports it. (?<!)
// match, sheetMatch, sheet1, sheet2, absCol, colName, absRow, row, isFunction, isRange
const cellReference = /((?:([A-Za-z_.\d\u007F-\uFFFF]+)|'((?:(?![\\\/\[\]*?:]).)+?)')!)?(\$)?([A-Za-z]+)(\$)?([0-9]+)(\()?(:(?:\$)?[A-Za-z]+(?:\$)?[0-9]+)?/g;

// match, sheetMatch, sheet1, sheet2, absCol1, colName1, absRow1, row1, absCol2, colName2, absRow2, row2
const rangeReference = /((?:([A-Za-z_.\d\u007F-\uFFFF]+)|'((?:(?![\\\/\[\]*?:]).)+?)')!)?(\$)?([A-Za-z]+)(\$)?([0-9]+):(\$)?([A-Za-z]+)(\$)?([0-9]+)/g;

// match, sheetMatch, sheet1, sheet2, absRow1, row1, absRow2, row2
const RowReference = /((?:([A-Za-z_.\d\u007F-\uFFFF]+)|'((?:(?![\\\/\[\]*?:]).)+?)')!)?(\$)?([0-9]+):(\$)?([0-9]+)/g;

// match, sheetMatch, sheet1, sheet2, absCol1, colName1, absCol2, colName2
const ColumnReference = /((?:([A-Za-z_.\d\u007F-\uFFFF]+)|'((?:(?![\\\/\[\]*?:]).)+?)')!)?(\$)?([A-Za-z]+):(\$)?([A-Za-z]+)(\$)?/g;

/**
 * @typedef {{COL: 0, ROW: 1}} MODE
 */
const MODE = {
    ROW: 0,
    COL: 1
};

module.exports = {
    MODE,
    /**
     * Remove/Replace corresponding part (sheetName, row, col) in the given formula.
     * For cell reference , replace it with #REF!;
     * For range reference,
     * @param {string} formula - the formula to replace
     * @param {string} currSheetName - the sheet that the formula in
     * @param {ReferenceLiteral} ref - cell reference
     * @param {MODE} mode
     * @return {string|undefined}
     */
    removeReference: (formula, currSheetName, ref, mode) => {
        if (formula == null)
            return;
        // remove one cell, thus diff = 1;
        const diff = 1;
        formula = formula.replace(cellReference, (match, sheetMatch, sheet1, sheet2, absCol, colName, absRow, row, isFunction, isRange) => {
            if (isFunction || isRange) {
                return match;
            }
            const sheet = sheet1 ? sheet1 : (sheet2 ? sheet2 : currSheetName);
            if (sheet !== ref.sheet) {
                return match;
            }
            if (+row === ref.row && ac.columnNameToNumber(colName) === ref.col) {
                return '#REF!';
            }
            return (sheetMatch ? sheetMatch : '') + (absCol ? '$' : '') + colName + (absRow ? '$' : '') + row;
        });

        formula = formula.replace(rangeReference, (match, sheetMatch, sheet1, sheet2, absCol1, colName1, absRow1, row1, absCol2, colName2, absRow2, row2) => {
            const sheet = sheet1 ? sheet1 : (sheet2 ? sheet2 : currSheetName);
            if (sheet !== ref.sheet) {
                return match;
            }
            if (mode === MODE.ROW) {
                row1 = Math.min(row1, row2);
                row2 = Math.max(row1, row2);
                if (row1 <= ref.row && ref.row <= row2 && Math.abs(row2 - row1) <= diff) return '#REF!';
                if (row1 <= ref.row && ref.row <= row2) {
                    row2 -= diff;
                }
            } else if (mode === MODE.COL) {
                let col1 = ac.columnNameToNumber(colName1);
                let col2 = ac.columnNameToNumber(colName2);
                col1 = Math.min(col1, col2);
                col2 = Math.max(col1, col2);
                if (col1 <= ref.col && ref.col <= col2 && Math.abs(col2 - col1) <= diff) return '#REF!';
                if (col1 <= ref.col && ref.col <= col2) {
                    col2 -= diff;
                }
                colName1 = ac.columnNumberToName(col1);
                colName2 = ac.columnNumberToName(col2);
            }

            return (sheetMatch ? sheetMatch : '') + (absCol1 ? '$' : '') + colName1 + (absRow1 ? '$' : '') + row1 + ':'
                + (absCol2 ? '$' : '') + colName2 + (absRow2 ? '$' : '') + row2;
        });

        if (mode === MODE.ROW) {
            formula = formula.replace(RowReference, (match, sheetMatch, sheet1, sheet2, absRow1, row1, absRow2, row2) => {
                const sheet = sheet1 ? sheet1 : (sheet2 ? sheet2 : currSheetName);
                if (sheet !== ref.sheet) {
                    return match;
                }
                row1 = Math.min(row1, row2);
                row2 = Math.max(row1, row2);
                if (row1 <= ref.row && ref.row <= row2 && Math.abs(row2 - row1) <= diff) return '#REF!';
                if (row1 <= ref.row && ref.row <= row2) {
                    row2 -= diff;
                }
                return (sheetMatch ? sheetMatch : '') + (absRow1 ? '$' : '') + row1 + ':' + (absRow2 ? '$' : '') + row2;
            });
        } else if (mode === MODE.COL) {
            formula = formula.replace(ColumnReference, (match, sheetMatch, sheet1, sheet2, absCol1, colName1, absCol2, colName2) => {
                const sheet = sheet1 ? sheet1 : (sheet2 ? sheet2 : currSheetName);
                if (sheet !== ref.sheet) {
                    return match;
                }
                let col1 = ac.columnNameToNumber(colName1);
                let col2 = ac.columnNameToNumber(colName2);
                col1 = Math.min(col1, col2);
                col2 = Math.max(col1, col2);
                if (col1 <= ref.col && ref.col <= col2 && Math.abs(col2 - col1) <= diff) return '#REF!';
                if (col1 <= ref.col && ref.col <= col2) {
                    col2 -= diff;
                }
                return (sheetMatch ? sheetMatch : '') + (absCol1 ? '$' : '') + ac.columnNumberToName(col1) + ':'
                    + (absCol2 ? '$' : '') + ac.columnNumberToName(col2);
            });
        }

        return formula;
    },

    /**
     * Called when delete/add row(s) between `oldRowNum` and `newRowNum` inclusive.
     * if (oldRowNum - newRowNum) > 0, then called when remove row(s); delete (oldRowNum - newRowNum) row(s) before oldRowNum;
     * if (oldRowNum - newRowNum) < 0, then called when add row(s); add (oldRowNum - newRowNum) row(s) before oldRowNum;
     * @param {string} formula - the formula to replace
     * @param {string} currSheetName - the sheet that the formula in
     * @param {string} sheetName
     * @param {number} changedRowNum - the row number of added or deleted row.
     * @param {number} offset - positive if add a row, negative if remove a row.
     * @return {string|undefined}
     */
    replaceRowNumber: (formula, currSheetName, sheetName, changedRowNum, offset) => {
        if (formula == null)
            return;

        formula = formula.replace(cellReference, (match, sheetMatch, sheet1, sheet2, absCol, colName, absRow, row, isFunction, isRange) => {
            if (isFunction || isRange) {
                return match;
            }
            const sheet = sheet1 ? sheet1 : (sheet2 ? sheet2 : currSheetName);
            if (sheet !== sheetName) {
                return match;
            }
            row = +row;
            if (row >= changedRowNum) { // "=" sign is meaningless when delete a row.
                row += offset;
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
            if (offset < 0 && row1 <= changedRowNum && row2 <= changedRowNum - offset) return '#REF!';
            if (row1 >= changedRowNum) row1 += offset;
            if (row2 >= changedRowNum) row2 += offset;
            if (row1 <= 0) row1 = 1;
            if (row2 - row1 === 0) return '#REF!';

            return (sheetMatch ? sheetMatch : '') + (absCol1 ? '$' : '') + colName1 + (absRow1 ? '$' : '') + row1 + ':'
                + (absCol2 ? '$' : '') + colName2 + (absRow2 ? '$' : '') + row2;
        });

        formula = formula.replace(RowReference, (match, sheetMatch, sheet1, sheet2, absRow1, row1, absRow2, row2) => {
            const sheet = sheet1 ? sheet1 : (sheet2 ? sheet2 : currSheetName);
            if (sheet !== sheetName) {
                return match;
            }
            row1 = Math.min(row1, row2);
            row2 = Math.max(row1, row2);
            if (offset < 0 && row1 <= changedRowNum && row2 <= changedRowNum - offset) return '#REF!';
            if (row1 >= changedRowNum) row1 += offset;
            if (row2 >= changedRowNum) row2 += offset;
            if (row1 <= 0) row1 = 1;
            if (row2 - row1 === 0) return '#REF!';

            return (sheetMatch ? sheetMatch : '') + (absRow1 ? '$' : '') + row1 + ':' + (absRow2 ? '$' : '') + row2;
        });
        return formula;
    },

};
