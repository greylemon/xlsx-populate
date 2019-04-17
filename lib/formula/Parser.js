"use strict";
const Cell = require('../Cell');
const Range = require('../Range');
const Row = require('../Row');
const Column = require('../Column');
const FormulaParser = require('fast-formula-parser');
const { DepParser } = FormulaParser;
const MAX_ROW = 1048576, MAX_COLUMN = 16384;

class Parser {
    constructor(workbook) {
        this._workbook = workbook;
        this.unfinishedStack = [];

        const config = {
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
            },
            onVariable: (name, sheet) => {
                // console.log(`Get Variable ${name} in ${sheet}`);
                // try sheet scoped first
                let range = this.currCell.sheet().definedName(name);
                if (!range) {
                    try {
                        range = this.currCell.workbook().definedName(name);
                    } catch (e) {
                        this.unfinishedStack.push(this.currCell);
                        return;
                    }
                }

                if (range instanceof Cell) {
                    return { row: range.rowNumber(), col: range.columnNumber(), sheet: range.sheet().name() };
                } else if (range instanceof Range) {
                    return {
                        sheet: range.startCell().sheet().name(),
                        from: { row: range.startCell().rowNumber(), col: range.startCell().columnNumber() },
                        to: { row: range.endCell().rowNumber(), col: range.endCell().columnNumber() }
                    };
                } else if (range instanceof Row) {
                    return {
                        sheet: range.sheet().name(),
                        from: { row: range.rowNumber(), col: 1 },
                        to: { row: range.rowNumber(), col: MAX_COLUMN }
                    };
                } else if (range instanceof Column) {
                    return {
                        sheet: range.sheet().name(),
                        from: { row: 1, col: range.columnNumber() },
                        to: { row: MAX_ROW, col: range.columnNumber() }
                    };
                }
            }
        };

        this._depParser = new DepParser(config);
        this._parser = new FormulaParser(config);
    }

    /**
     * Parse dependencies.
     * @param {Cell} cell - The Cell want to parse
     */
    parseDep(cell) {
        if (cell._formula == null) return [];
        this.currCell = cell;
        const deps = this._depParser.parse(cell._formula, cell.getRef());
        return deps;
    }

    /**
     *
     * @param {Cell} cell
     * @return {*}
     */
    parse(cell) {
        const formula = cell.getFormula();
        if (formula == null) return;
        this.currCell = cell;
        const result = this._parser.parse(formula, cell.getRef());
        return result;
    }

    decode(value) {
        return value;
        return value.replace(/(&amp;)|(&lt;)|(&gt;)/g, (val, g1, g2, g3) => {
            return g1 ? '&' : (g2 ? '<' : (g3 ? '>' : ''));
        });
    }
}

module.exports = Parser;
