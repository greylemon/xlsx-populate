"use strict";
const Cell = require('../worksheets/Cell');
const RichText = require('../worksheets/RichText');
const FormulaParser = require('fast-formula-parser');
const { DepParser } = FormulaParser;
const MAX_ROW = 1048576, MAX_COLUMN = 16384;

class Parser {
    constructor(workbook) {
        this._workbook = workbook;

        const config = {
            onCell: ref => {
                let val = null;
                const sheet = this._workbook.sheet(ref.sheet);
                if (sheet.hasCell(ref.row, ref.col)) {
                    val = sheet.getCell(ref.row, ref.col).getValue();
                    if (val instanceof RichText)
                        val = val.text();
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
                        let cellValue = row.cell(ref.from.row).getValue();
                        if (cellValue instanceof RichText)
                            cellValue = cellValue.text();
                        arr[rowNumber] = [cellValue == null ? null : cellValue];
                    });
                } else if (ref.to.col === MAX_COLUMN) {
                    // whole row
                    arr.push([]);
                    sheet._rows.get(ref.from.row).forEach(cell => {
                        let cellValue = cell.getValue();
                        if (cellValue instanceof RichText)
                            cellValue = cellValue.text();
                        arr[0].push(cellValue == null ? null : cellValue);
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
                                    let cellValue = cell.getValue();
                                    if (cellValue instanceof RichText)
                                        cellValue = cellValue.text();
                                    innerArr[col - 1] = cellValue;
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
                let ref = this.currCell.sheet().definedName(name);
                if (!ref) {
                    ref = this.currCell.workbook().definedName(name);
                }
                return ref.toObject();
            }
        };

        this._depParser = new DepParser(config);
        this._parser = new FormulaParser(config);
        this._refParser = new DepParser(config);
    }

    /**
     * Parse dependencies.
     * @param {Cell} cell - The Cell want to parse
     * @return {Array} The dependencies of the given formula.
     */
    parseDep(cell) {
        if (cell._formula == null) return [];
        this.currCell = cell;
        let deps;
        try {
            deps = this._depParser.parse(cell._formula, cell.getRef());
        } catch (e) {
            console.warn(`Error when parsing dependency: ${cell._formula}, skipped, position:${JSON.stringify(cell.getRef())}`);
            // console.error(e);
            return [];
        }
        return deps;
    }

    /**
     * Parse a cell's formula.
     * @param {Cell} cell - A cell that contains formula.
     * @return {*} - The result of the formula
     */
    parse(cell) {
        const formula = cell.getFormula();
        if (formula == null) return;
        this.currCell = cell;
        let result;
        try {
            result = this._parser.parse(formula, cell.getRef());
        } catch (e) {
            console.warn(`Error when parsing: ${formula}, skipped`);
        }
        return result;
    }

    /**
     * Parse a reference, i.e. D19, A1:C9, Sheet2!A1, A:C, 1:6...
     * @param {string} address - The reference string to parse.
     * @param {Sheet} [sheet] - The sheet where the address in.
     * @return {{sheet: string, row: number, col: number}} The reference object.
     */
    parseReference(address, sheet) {
        return this._refParser.parse(address, { sheet })[0];
    }
}

module.exports = Parser;
