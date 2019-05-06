"use strict";

const MAX_COL = 2 ** 14 + 1, MAX_ROW = 2 ** 20 + 1;

/**
 * A reference table:  refA -> refB (refB depends on refA)
 * Using 1-based index.
 */
class ReferenceTable {
    constructor() {
        //  refA -> refB (refB depends on refA)
        this._data = new Map();

        // store dependency tree to help remove data fast.
        //  refB -> refA (refB depends on refA)
        this._depData = new Map();
    }

    /**
     * refB depends on refA, which means changes on refA will trigger
     * re-calculation on refB.
     * refA -> refB
     * @param {{sheet, from: {row, col}, to: {row, col}} | {row, col}} refA - Dependency of refB.
     * @param {{row, col, sheet}} refB - A cell reference.
     * @return {undefined}
     */
    add(refA, refB) {
        const sheet = refA.sheet;
        if (typeof refA.sheet !== "string" || typeof refB.sheet !== "string")
            throw Error(`Sheet must be string`);

        let refASheet = this._data.get(sheet);
        if (!refASheet) {
            refASheet = new Map();
            this._data.set(sheet, refASheet);
        }

        let rowRange, colRange;

        if (refA.from) {
            // refA is a range reference
            rowRange = this.encodeRowRange(refA.from.row, refA.to.row);
            colRange = this.encodeColRange(refA.from.col, refA.to.col);
        } else {
            rowRange = this.encodeRowRange(refA.row, refA.row);
            colRange = this.encodeColRange(refA.col, refA.col);
        }

        // get row range
        let refARow = refASheet.get(rowRange);
        if (!refARow) {
            refARow = new Map();
            refASheet.set(rowRange, refARow);
        }

        // get col range
        let refACol = refARow.get(colRange);
        if (!refACol) {
            refACol = new Map();
            refARow.set(colRange, refACol);
        }

        // get refB sheet
        let refBSheet = refACol.get(refB.sheet);
        if (!refBSheet) {
            refBSheet = [];
            refACol.set(refB.sheet, refBSheet);
        }

        refBSheet.push(this.encodeCell(refB.row, refB.col));
    }

    /**
     * Called when a cell (refB)'s value is cleared.
     * Remember refA -> refB,
     * @param {{sheet, from: {row, col}, to: {row, col}} | {row, col}} refA - Dependency of refB.
     * @param {{row, col, sheet}} refB - A cell reference.
     * @return {undefined}
     */
    remove(refA, refB) {
        const sheet = refA.sheet;
        if (typeof refA.sheet !== "string" || typeof refB.sheet !== "string")
            throw Error(`Sheet must be string`);

        const refASheet = this._data.get(sheet);
        if (!refASheet) {
            return;
        }

        let rowRange, colRange;

        if (refA.from) {
            // refA is a range reference
            rowRange = this.encodeRowRange(refA.from.row, refA.to.row);
            colRange = this.encodeColRange(refA.from.col, refA.to.col);
        } else {
            rowRange = this.encodeRowRange(refA.row, refA.row);
            colRange = this.encodeColRange(refA.col, refA.col);
        }

        // get row range
        const refARow = refASheet.get(rowRange);
        if (!refARow) {
            return;
        }

        // get col range
        const refACol = refARow.get(colRange);
        if (!refACol) {
            return;
        }

        // get refB sheet
        const refBSheet = refACol.get(refB.sheet);
        if (!refBSheet) {
            return;
        }
        const valueToRemove = this.encodeCell(refB.row, refB.col);

        refBSheet.filter(element => {
            return element === valueToRemove;
        });
    }

    /**
     * Get all cells need to update in order given a ref is modified. Called when a cell is modified.
     * @param {{row, col, sheet}} ref - A cell reference
     * @param result
     */
    getCalculationOrder(ref, result) {
        if (!result) result = [];

        const refASheet = this._data.get(ref.sheet);
        if (!refASheet) return result;

        refASheet.forEach((refARow, rowRange) => {
            if (!this.isInRowRange(rowRange, ref.row)) return result;

            refARow.forEach((refACol, colRange) => {
                if (!this.isInColRange(colRange, ref.col)) return result;

                refACol.forEach((refBSheet, sheet) => {
                    refBSheet.forEach(cell => {
                        const refB = this.decodeCell(cell);

                        // skip self references
                        if (refB.row === ref.row && refB.col === ref.col)
                            return;
                        refB.sheet = sheet;
                        const idxExist = result.findIndex(element => {
                            // ensure calculation on a cell won't happen many times.
                            return element.row === refB.row && element.col === refB.col && element.sheet === refB.sheet;
                        });
                        if (idxExist !== -1) result.splice(idxExist, 1);
                        result.push(refB);
                        this.getCalculationOrder(refB, result);
                    });
                });
            });
        });
        return result;
    }

    encodeRowRange(from, to) {
        return from * MAX_ROW + to;
    }

    decodeRowRange(range) {
        const to = range % MAX_ROW;
        return { to, from: (range - to) / MAX_ROW };
    }

    encodeColRange(from, to) {
        return from * MAX_COL + to;
    }

    decodeColRange(range) {
        const to = range % MAX_COL;
        return { to, from: (range - to) / MAX_COL };
    }

    isInRowRange(range, rowNum) {
        const { from, to } = this.decodeRowRange(range);
        return from <= rowNum && rowNum <= to;
    }

    isInColRange(range, colNum) {
        const { from, to } = this.decodeColRange(range);
        return from <= colNum && colNum <= to;
    }

    encodeCell(row, col) {
        return row * MAX_COL + col;
    }

    decodeCell(cell) {
        const col = cell % MAX_COL;
        return { col, row: (cell - col) / MAX_COL };
    }
}

module.exports = ReferenceTable;
