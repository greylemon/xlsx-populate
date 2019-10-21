"use strict";

const {
    decodeCell, encodeCell, encodeColRange, encodeRowRange, isInColRange, isInRowRange
} = require('./Utils');

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
            rowRange = encodeRowRange(refA.from.row, refA.to.row);
            colRange = encodeColRange(refA.from.col, refA.to.col);
        } else {
            rowRange = encodeRowRange(refA.row, refA.row);
            colRange = encodeColRange(refA.col, refA.col);
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

        refBSheet.push(encodeCell(refB.row, refB.col));
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
            rowRange = encodeRowRange(refA.from.row, refA.to.row);
            colRange = encodeColRange(refA.from.col, refA.to.col);
        } else {
            rowRange = encodeRowRange(refA.row, refA.row);
            colRange = encodeColRange(refA.col, refA.col);
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
        const valueToRemove = encodeCell(refB.row, refB.col);

        refACol.set(refB.sheet, refBSheet.filter(element => {
            return element !== valueToRemove;
        }));
    }

    /**
     * Get all cells need to update in order given a ref is modified. Called when a cell is modified.
     * dep: include cell that already depends on other cell.
     * @param {{row, col, sheet}} ref - A cell reference
     * @param result
     */
    getCalculationOrder(ref, dep, result) {
        if (!result) result = [];
        if (!dep) dep = [];

        const refASheet = this._data.get(ref.sheet);
        if (!refASheet) return result;

        refASheet.forEach((refARow, rowRange) => {
            if (!isInRowRange(rowRange, ref.row)) return result;

            refARow.forEach((refACol, colRange) => {
                if (!isInColRange(colRange, ref.col)) return result;

                refACol.forEach((refBSheet, sheet) => {
                    refBSheet.forEach(cell => {
                        const refB = decodeCell(cell);

                        // skip self references
                        if (refB.row === ref.row && refB.col === ref.col)
                            return;
                        refB.sheet = sheet;

                        // avoid edit mutiple times
                        const idxExist1 = result.findIndex(element => {
                          // ensure calculation on a cell won't happen many times.
                          return element.row === refB.row && element.col === refB.col && element.sheet === refB.sheet;
                        });
                        if(idxExist1 === -1) {
                            const idxExist = dep.findIndex(element => {
                              // ensure calculation on a cell won't happen many times.
                              return element.row === refB.row && element.col === refB.col && element.sheet === refB.sheet;
                            });
                            if (idxExist !== -1) {
                              dep.splice(idxExist, 1);
                            }

                            else
                              result.push(refB);
                            dep.push(refB);
                            this.getCalculationOrder(refB, dep, result);
                        }
                    });
                });
            });
        });
        return result;
    }

    /**
     * Get CalculationOrder for the first layer.
     * @param {{row, col, sheet}} ref - A cell reference
     * @return {Array} The cells direct reference the given cell reference.
     */
    getDirectReferences(ref) {
        const result = [];

        const refASheet = this._data.get(ref.sheet);
        if (!refASheet) return result;

        refASheet.forEach((refARow, rowRange) => {
            if (!isInRowRange(rowRange, ref.row)) return result;

            refARow.forEach((refACol, colRange) => {
                if (!isInColRange(colRange, ref.col)) return result;

                refACol.forEach((refBSheet, sheet) => {
                    refBSheet.forEach(cell => {
                        const refB = decodeCell(cell);

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
                    });
                });
            });
        });
        return result;
    }
}

module.exports = ReferenceTable;
