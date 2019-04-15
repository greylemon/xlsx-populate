"use strict";

const MAX_COL = 2 ** 16, MAX_ROW = 2 ** 20;

/**
 * A reference table:  refA -> refB (refB depends on refA)
 * Using 1-based index.
 */
class ReferenceTable {
    constructor() {
        this._data = new Map();
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
            colRange = this.encodeRowRange(refA.from.col, refA.to.col);
        } else {
            rowRange = this.encodeRowRange(refA.row, refA.row);
            colRange = this.encodeRowRange(refA.col, refA.col);
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
     * Called when a cell's value is cleared.
     * Remember refA -> refB,
     * Ideally ref can be either refA or refB, but ref is treated as refA here. (Partial remove)
     * In future calculation, we will check if refB is a formula, if it's not, remove it.
     * @param ref
     */
    remove(ref) {

    }

    /**
     * Get all cells need to update given a ref is modified. Called when a cell is modified.
     * @param ref
     */
    get(ref, result) {
        if (!result) result = [];
        let sheet = ref.sheet;
        if (typeof ref.sheet === "string")
            sheet = this._sheetNames.indexOf(ref.sheet);
        let data = this._data.get(sheet);
        if (!data) return;
        data = this._data.get(ref.row * 100000 + ref.col);
        if (!sheet) return;
    }

    encodeRowRange(from, to) {
        return from * MAX_ROW + to; // left shift 20 bits
    }

    decodeRowRange(range) {
        const to = range % MAX_ROW;
        return { to, from: (range - to) / MAX_ROW };
    }

    encodeColRange(from, to) {
        return from * MAX_COL + to; // left shift 16 bits
    }

    decodeColRange(range) {
        const to = range % MAX_COL;
        return { to, from: (range - to) / MAX_COL };
    }

    isInRowRange(range, rowNum) {
        const { from, to } = this.decodeRowRange(range);
        return from <= rowNum && rowNum <= to;
    }

    isInCowRange(range, rowNum) {
        const { from, to } = this.decodeColRange(range);
        return from <= rowNum && rowNum <= to;
    }

    encodeCell(row, col) {
        return row * MAX_COL + col; // left shift 16 bits
    }

    decodeCell(cell) {
        const col = cell % MAX_COL;
        return { col, row: (cell - col) / MAX_COL };
    }
}

module.exports = ReferenceTable;
