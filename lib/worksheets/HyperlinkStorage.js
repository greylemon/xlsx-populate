"use strict";

const {
    decodeCell, encodeCell, encodeColRange, encodeRowRange, isInColRange, isInRowRange
} = require('../formula/Utils');

class HyperlinkStorage {
    constructor(sheet, hyperlinkNode) {
        this._sheet = sheet;
        this._hyperlinkNode = hyperlinkNode;

    }
}

module.exports = HyperlinkStorage;
