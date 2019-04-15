"use strict";
const FormulaParser = require('fast-formula-parser');
const { DepParser } = FormulaParser;
const MAX_ROW = 1048576, MAX_COLUMN = 16384;


class Parser {
    constructor(workbook) {
        this._workbook = workbook;
        this._depParser = new DepParser();

        this._parser = new FormulaParser({
            onCell: ref => {
                let val = null;
                const sheet = this._workbook.sheet(ref.sheet);
                if (sheet.hasCell(ref.row, ref.col)) {
                    val = sheet.getCell(ref.row, ref.col).getValue();
                }

                // console.log(`Get cell ${val}`);
                return val == null ? undefined : val;
            },
            onRange: ref => {
                const arr = [];
                const sheet = this._workbook.sheet(ref.sheet);

                // whole column
                if (ref.to.row === MAX_ROW) {
                    sheet._rows.forEach((row, rowNumber) => {
                        const cellValue = row.cell(ref.from.row)._value;
                        arr[rowNumber] = [cellValue == null ? null : cellValue];
                    });
                } else if (ref.to.col === MAX_COLUMN) {
                    // whole row
                    arr.push([]);
                    sheet._rows.get(ref.from.row).forEach(cell => {
                        arr[0].push(cell._value == null ? null : cell._value);
                    });
                } else {
                    const sheet = this._workbook.sheet(ref.sheet);

                    for (let row = ref.from.row; row <= ref.to.row; row++) {
                        const innerArr = [];

                        // row exists
                        if (sheet._rows.has(row)) {
                            for (let col = ref.from.col; col <= ref.to.col; col++) {
                                const cell = sheet._rows.get(row)._cells.get(col);
                                if (cell != null) {
                                    innerArr[col - 1] = cell._value;
                                }
                            }
                        }
                        arr.push(innerArr);
                    }
                }
                return arr;
            }
        });
    }

    /**
     * Parse dependencies.
     * @param {Cell} cell - The Cell want to parse
     */
    parseDep(cell) {
        if (cell._formula == null) return [];
        const deps = this._depParser.parse(cell._formula, {
            sheet: cell.sheet().name(),
            row: cell.rowNumber(),
            col: cell.columnNumber()
        });
        return deps;
    }

    /**
     *
     * @param {Cell} cell
     * @return {*}
     */
    parse(cell) {
        if (cell._formula == null) return [];
        const result = this._parser.parse(cell._formula, {
            sheet: cell.sheet().name(),
            row: cell.rowNumber(),
            col: cell.columnNumber()
        });
        return result;
    }
}

module.exports = Parser;
