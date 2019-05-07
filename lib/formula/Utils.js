"use strict";

const MAX_COL = 2 ** 14 + 1, MAX_ROW = 2 ** 20 + 1;

const helpers = {
    encodeRowRange: (from, to) => {
        return from * MAX_ROW + to;
    },

    decodeRowRange: range => {
        const to = range % MAX_ROW;
        return { to, from: (range - to) / MAX_ROW };
    },

    encodeColRange: (from, to) => {
        return from * MAX_COL + to;
    },

    decodeColRange: range => {
        const to = range % MAX_COL;
        return { to, from: (range - to) / MAX_COL };
    },

    isInRowRange: (range, rowNum) => {
        const { from, to } = helpers.decodeRowRange(range);
        return from <= rowNum && rowNum <= to;
    },

    isInColRange: (range, colNum) => {
        const { from, to } = helpers.decodeColRange(range);
        return from <= colNum && colNum <= to;
    },

    encodeCell: (row, col) => {
        return row * MAX_COL + col;
    },

    decodeCell: cell => {
        const col = cell % MAX_COL;
        return { col, row: (cell - col) / MAX_COL };
    }
};

module.exports = helpers;
