"use strict";

const MAX_COL = 2 ** 14, MAX_ROW = 2 ** 20;

const Type = {
    CELL: 0,
    RANGE: 1,
    ROW: 2,
    COL: 3
};

class Reference {
    /**
     * @param {{row, col, sheet}|{from: {row, col}, to: {row, col}}} ref - The reference.
     * @param {Workbook} workbook - Workbook
     */
    constructor(ref, workbook) {
        this._workbook = workbook;
        this._sheet = ref.sheet;
        if (ref.from.row === ref.to.row && ref.from.col === ref.to.col || !ref.from) {
            // cell reference
            // row, col
            this._data = new Uint32Array([ref.row, ref.col]);
        } else {
            // range reference
            // fromRow, fromCol, toRow, toCol
            this._data = new Uint32Array([ref.from.row, ref.from.col, ref.to.row, ref.to.col]);
        }
    }

    get sheet() {
        return this._sheet;
    }

    get from() {
        if (this.type === Type.CELL) return;
        return { row: this._data[0], col: this._data[1] };
    }

    get to() {
        if (this.type === Type.CELL) return;
        return { row: this._data[1], col: this._data[3] };
    }

    get row() {
        if (this.type === Type.CELL)
            return this._data[0];
    }

    get col() {
        if (this.type === Type.CELL)
            return this._data[1];
    }

    get data() {
        return this._data;
    }

    get type() {
        if (this._data.length === 2)
            return Type.CELL;
        else if (this._data[0] !== this._data[2] && this._data[1] !== this._data[3])
            return Type.RANGE;
        else if (this._data[0] === this._data[2] && this._data[1] === 1 && this._data[3] === MAX_COL)
            return Type.ROW;
        else if (this._data[1] === this._data[3] && this._data[0] === 1 && this._data[2] === MAX_ROW)
            return Type.COL;
    }

    /**
     * Retrieve the referenced object. Can be Cell, Range, Row, Column.
     * @param {Workbook} [workbook] - The workbook uses to retrieve reference.
     * @return {Cell|Range|Row|Column} The referenced object.
     */
    retrieve(workbook) {
        if (workbook) this._workbook = workbook;
        const type = this.type;
        const sheet = this._workbook.sheet(this._sheet);
        if (type === Type.CELL)
            return sheet.getCell(this._data[0], this._data[1]);
        else if (type === Type.RANGE)
            return sheet.range(...this._data);
        else if (type === Type.ROW)
            return sheet.row(this._data[0]);
        else if (type === Type.COL)
            return sheet.col(this._data[2]);
    }

    toObject() {
        return this.type === Type.CELL ? {
            sheet: this._sheet,
            row: this._data[0],
            col: this._data[1]
        } : {
            sheet: this._sheet,
            from: {
                row: this._data[0],
                col: this._data[1]
            },
            to: {
                row: this._data[2],
                col: this._data[3]
            }
        };
    }
}

Reference.Type = Type;

module.exports = Reference;
