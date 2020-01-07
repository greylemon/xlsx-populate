"use strict";

const { cloneDeep } = require('../utils');
const Cell = require("./Cell");
const Row = require("./Row");
const Column = require("./Column");
const Range = require("./Range");
const Relationships = require("../workbooks/Relationships");
const xmlq = require("../xml/xmlq");
const regexify = require("../regexify");
const addressConverter = require("../addressConverter");
const ArgHandler = require("../ArgHandler");
const colorIndexes = require("../colorIndexes");
const PageBreaks = require("./PageBreaks");
const Hyperlinks = require("./Hyperlinks");
const DataValidations = require("./DataValidations");
const MergeCells = require("./MergeCells");
const Reference = require("../formula/Reference");
const { Extensions, ExtURI } = require("./Extensions");
const FormulaReplacer = require("../formula/Replacer");

// Order of the nodes as defined by the spec.
const nodeOrder = [
    "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
    "sheetCalcPr", "sheetProtection", "autoFilter", "protectedRanges", "scenarios", "autoFilter",
    "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr",
    "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions",
    "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
    "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
    "drawingHF", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts",
    "extLst"
];

/**
 * A worksheet.
 */
class Sheet {
    // /**
    //  * Creates a new instance of Sheet.
    //  * @param {Workbook} workbook - The parent workbook.
    //  * @param {{}} idNode - The sheet ID node (from the parent workbook).
    //  * @param {{}} node - The sheet node.
    //  * @param {{}} [relationshipsNode] - The optional sheet relationships node.
    //  */
    constructor(workbook, idNode, node, relationshipsNode) {
        this._workbook = workbook;
        this._idNode = idNode;
        this._node = node;
        this._relationshipsNode = relationshipsNode;
        // this.totalcol = this.row.size;
        // this.totalraw = this.ro;
    }

    /* PUBLIC */


    /* FAST GETTER AND SETTER */

    /**
     * Get cell using 1-based index.
     * @param {number} rowNum
     * @param {number} colNum
     * @return {Cell}
     */
    getCell(rowNum, colNum) {
        return this.row(rowNum).cell(colNum);
    }

    getCellByAddress(address) {
        const ref = addressConverter.fromAddress(address);
        if (ref.type !== 'cell') throw new Error('Sheet.cell: Invalid address.');
        return this.row(ref.rowNumber).cell(ref.columnNumber);
    }

    hasCell(rowNum, colNum) {
        return this._rows.has(rowNum) && this._rows.get(rowNum)._cells.has(colNum);
    }

    getId() {
        return this._idNode.attributes.sheetId;
    }

    /**
     * Gets a value indicating whether the sheet is the active sheet in the workbook.
     * @returns {boolean} True if active, false otherwise.
     *//**
     * Make the sheet the active sheet in the workkbok.
     * @param {boolean} active - Must be set to `true`. Deactivating directly is not supported. To deactivate, you should activate a different sheet instead.
     * @returns {Sheet} The sheet.
     */
    active() {
        return new ArgHandler('Sheet.active', arguments)
            .case(() => {
                return this.workbook().activeSheet() === this;
            })
            .case('boolean', active => {
                if (!active) throw new Error("Deactivating sheet directly not supported. Activate a different sheet instead.");
                this.workbook().activeSheet(this);
                return this;
            })
            .handle();
    }

    /**
     * Get the active cell in the sheet.
     * @returns {Cell|Range} The active cell.
     *//**
     * Set the active cell in the workbook.
     * @param {string|Cell|Range} cell - The cell or address of cell to activate.
     * @returns {Sheet} The sheet.
     *//**
     * Set the active cell in the workbook by row and column.
     * @param {number} rowNumber - The row number of the cell.
     * @param {string|number} columnNameOrNumber - The column name or number of the cell.
     * @returns {Sheet} The sheet.
     */
    activeCell() {
        const sheetViewNode = this._getOrCreateSheetViewNode();
        let selectionNode = xmlq.findChild(sheetViewNode, "selection", child => !!child.attributes.activeCell);
        const toCellOrRange = address => {
            const ref = addressConverter.fromAddress(address);
            if (ref.type === 'range') {
                return this.range(ref.startRowNumber, ref.startColumnNumber, ref.endRowNumber, ref.endColumnNumber);
            } else {
                return this.row(ref.rowNumber).cell(ref.columnNumber);
            }
        };
        return new ArgHandler('Sheet.activeCell', arguments)
            .case(() => {
                const cellAddress = selectionNode ? selectionNode.attributes.sqref : "A1";
                return toCellOrRange(cellAddress);
            })
            .case(['number', '*'], (rowNumber, columnNameOrNumber) => {
                const cell = this.cell(rowNumber, columnNameOrNumber);
                return this.activeCell(cell);
            })
            .case('*', cellOrRange => {
                if (!selectionNode) {
                    selectionNode = {
                        name: "selection",
                        attributes: {},
                        children: []
                    };

                    xmlq.appendChild(sheetViewNode, selectionNode);
                }
                if (typeof cellOrRange === 'string') cellOrRange = toCellOrRange(cellOrRange);

                selectionNode.attributes.sqref = cellOrRange.address();
                selectionNode.attributes.activeCell = cellOrRange instanceof Cell ? cellOrRange.address()
                    : cellOrRange.startCell().address();
                return this;
            })
            .handle();
    }

    /**
     * Gets the cell with the given address.
     * @param {string} address - The address of the cell.
     * @returns {Cell} The cell.
     *//**
     * Gets the cell with the given row and column numbers.
     * @param {number} rowNumber - The row number of the cell.
     * @param {string|number} columnNameOrNumber - The column name or number of the cell.
     * @returns {Cell} The cell.
     */
    cell() {
        return new ArgHandler('Sheet.cell', arguments)
            .case('string', address => {
                return this.getCellByAddress(address);
            })
            .case(['number', '*'], (rowNumber, columnNameOrNumber) => {
                return this.getCell(rowNumber, columnNameOrNumber);
            })
            .handle();
    }

    /**
     * Gets a column in the sheet.
     * @param {string|number} columnNameOrNumber - The name or number of the column.
     * @returns {Column} The column.
     */
    column(columnNameOrNumber) {
        const columnNumber = typeof columnNameOrNumber === "string" ? addressConverter.columnNameToNumber(columnNameOrNumber) : columnNameOrNumber;

        // If we're already created a column for this column number, return it.
        if (this._columns[columnNumber]) return this._columns[columnNumber];

        // We need to create a new column, which requires a backing col node. There may already exist a node whose min/max cover our column.
        // First, see if there is an existing col node.
        const existingColNode = this._colNodes[columnNumber];

        let colNode;
        if (existingColNode) {
            // If the existing node covered earlier columns than the new one, we need to have a col node to cover the min up to our new node.
            if (existingColNode.attributes.min < columnNumber) {
                // Clone the node and set the max to the column before our new col.
                const beforeColNode = cloneDeep(existingColNode);
                beforeColNode.attributes.max = columnNumber - 1;

                // Update the col nodes cache.
                for (let i = beforeColNode.attributes.min; i <= beforeColNode.attributes.max; i++) {
                    this._colNodes[i] = beforeColNode;
                }
            }

            // Make a clone for the new column. Set the min/max to the column number and cache it.
            colNode = cloneDeep(existingColNode);
            colNode.attributes.min = columnNumber;
            colNode.attributes.max = columnNumber;
            this._colNodes[columnNumber] = colNode;

            // If the max of the existing node is greater than the nre one, create a col node for that too.
            if (existingColNode.attributes.max > columnNumber) {
                const afterColNode = cloneDeep(existingColNode);
                afterColNode.attributes.min = columnNumber + 1;
                for (let i = afterColNode.attributes.min; i <= afterColNode.attributes.max; i++) {
                    this._colNodes[i] = afterColNode;
                }
            }
        } else {
            // The was no existing node so create a new one.
            colNode = {
                name: 'col',
                attributes: {
                    min: columnNumber,
                    max: columnNumber
                },
                children: []
            };

            this._colNodes[columnNumber] = colNode;
        }

        // Create the new column and cache it.
        const column = new Column(this, colNode);
        this._columns[columnNumber] = column;
        return column;
    }

    /**
     * Gets a defined name scoped to the sheet.
     * @param {string} name - The defined name.
     * @returns {undefined|string|Cell|Range|Row|Column} What the defined name refers to or undefined if not found. Will return the string formula if not a Row, Column, Cell, or Range.
     *//**
     * Set a defined name scoped to the sheet.
     * @param {string} name - The defined name.
     * @param {string|Cell|Range|Row|Column} refersTo - What the name refers to.
     * @returns {Workbook} The workbook.
     */
    definedName() {
        return new ArgHandler("Workbook.definedName", arguments)
            .case('string', name => {
                return this.workbook().scopedDefinedName(this, name);
            })
            .case(['string', '*'], (name, refersTo) => {
                this.workbook().scopedDefinedName(this, name, refersTo);
                return this;
            })
            .handle();
    }

    /**
     * Deletes the sheet and returns the parent workbook.
     * @returns {Workbook} The workbook.
     */
    delete() {
        this.workbook().deleteSheet(this);
        return this.workbook();
    }

    /**
     * Find the given pattern in the sheet and optionally replace it.
     * @param {string|RegExp} pattern - The pattern to look for. Providing a string will result in a case-insensitive substring search. Use a RegExp for more sophisticated searches.
     * @param {string|function} [replacement] - The text to replace or a String.replace callback function. If pattern is a string, all occurrences of the pattern in each cell will be replaced.
     * @returns {Array.<Cell>} The matching cells.
     */
    find(pattern, replacement) {
        pattern = regexify(pattern);

        let matches = [];
        this._rows.forEach(row => {
            if (!row) return;
            matches = matches.concat(row.find(pattern, replacement));
        });

        return matches;
    }

    /**
     * Gets a value indicating whether this sheet's grid lines are visible.
     * @returns {boolean} True if selected, false if not.
     *//**
     * Sets whether this sheet's grid lines are visible.
     * @param {boolean} selected - True to make visible, false to hide.
     * @returns {Sheet} The sheet.
     */
    gridLinesVisible() {
        const sheetViewNode = this._getOrCreateSheetViewNode();
        return new ArgHandler('Sheet.gridLinesVisible', arguments)
            .case(() => {
                return sheetViewNode.attributes.showGridLines === 1 || sheetViewNode.attributes.showGridLines === undefined;
            })
            .case('boolean', visible => {
                sheetViewNode.attributes.showGridLines = visible ? 1 : 0;
                return this;
            })
            .handle();
    }

    /**
     * Gets a value indicating if the sheet is hidden or not.
     * @returns {boolean|string} True if hidden, false if visible, and 'very' if very hidden.
     *//**
     * Set whether the sheet is hidden or not.
     * @param {boolean|string} hidden - True to hide, false to show, and 'very' to make very hidden.
     * @returns {Sheet} The sheet.
     */
    hidden() {
        return new ArgHandler('Sheet.hidden', arguments)
            .case(() => {
                if (this._idNode.attributes.state === 'hidden') return true;
                if (this._idNode.attributes.state === 'veryHidden') return "very";
                return false;
            })
            .case('*', hidden => {
                if (hidden) {
                    const visibleSheets = this.workbook().sheets().filter(sheet => !sheet.hidden());
                    if (visibleSheets.length === 1 && visibleSheets[0] === this) {
                        throw new Error("This sheet may not be hidden as a workbook must contain at least one visible sheet.");
                    }

                    // If activate, activate the first other visible sheet.
                    if (this.active()) {
                        const activeIndex = visibleSheets[0] === this ? 1 : 0;
                        visibleSheets[activeIndex].active(true);
                    }
                }

                if (hidden === 'very') this._idNode.attributes.state = 'veryHidden';
                else if (hidden) this._idNode.attributes.state = 'hidden';
                else delete this._idNode.attributes.state;
                return this;
            })
            .handle();
    }

    /**
     * Move the sheet.
     * @param {number|string|Sheet} [indexOrBeforeSheet] The index to move the sheet to or the sheet (or name of sheet) to move this sheet before. Omit this argument to move to the end of the workbook.
     * @returns {Sheet} The sheet.
     */
    move(indexOrBeforeSheet) {
        this.workbook().moveSheet(this, indexOrBeforeSheet);
        return this;
    }

    /**
     * Get or Set the name of the sheet.
     * *Note: this method does not rename references to the sheet so formulas, etc. can be broken. Use with caution!*
     * @param {string} [name] - The name to set to the sheet.
     * @return {Sheet|string} The sheet if set sheet name;
     *                      The sheet name if get sheet name.
     */
    name(name) {
        if (name === undefined) {
            return `${this._idNode.attributes.name}`;
        } else {
            this._idNode.attributes.name = name;
            return this;
        }
    }

    /**
     * Gets a range from the given range address.
     * @param {string} address - The range address (e.g. 'A1:B3').
     * @returns {Range} The range.
     *//**
     * Gets a range from the given cells or cell addresses.
     * @param {string|Cell} startCell - The starting cell or cell address (e.g. 'A1').
     * @param {string|Cell} endCell - The ending cell or cell address (e.g. 'B3').
     * @returns {Range} The range.
     *//**
     * Gets a range from the given row numbers and column names or numbers.
     * @param {number} startRowNumber - The starting cell row number.
     * @param {string|number} startColumnNameOrNumber - The starting cell column name or number.
     * @param {number} endRowNumber - The ending cell row number.
     * @param {string|number} endColumnNameOrNumber - The ending cell column name or number.
     * @returns {Range} The range.
     */
    range() {
        return new ArgHandler('Sheet.range', arguments)
            .case('string', address => {
                const ref = addressConverter.fromAddress(address);
                if (ref.type !== 'range') throw new Error('Sheet.range: Invalid address');
                return this.range(ref.startRowNumber, ref.startColumnNumber, ref.endRowNumber, ref.endColumnNumber);
            })
            .case(['*', '*'], (startCell, endCell) => {
                if (typeof startCell === "string") startCell = this.cell(startCell);
                if (typeof endCell === "string") endCell = this.cell(endCell);
                return new Range(startCell, endCell);
            })
            .case(['number', '*', 'number', '*'], (startRowNumber, startColumnNameOrNumber, endRowNumber, endColumnNameOrNumber) => {
                return this.range(this.cell(startRowNumber, startColumnNameOrNumber), this.cell(endRowNumber, endColumnNameOrNumber));
            })
            .handle();
    }

    /**
     * Unsets sheet autoFilter.
     * @returns {Sheet} This sheet.
     *//**
     * Sets sheet autoFilter to a Range.
     * @param {Range} range - The autoFilter range.
     * @returns {Sheet} This sheet.
     */
    autoFilter(range) {
        this._autoFilter = range;

        return this;
    }

    /**
     * Gets the row with the given number.
     * @param {number} rowNumber - The row number.
     * @returns {Row} The row with the given number.
     */
    row(rowNumber) {
        if (rowNumber < 1) throw new RangeError(`Invalid row number ${rowNumber}. Remember that spreadsheets use 1-based indexing.`);

        if (this._rows.get(rowNumber)) return this._rows.get(rowNumber);

        const rowNode = {
            name: 'row',
            attributes: {
                r: rowNumber
            },
            children: []
        };

        const row = new Row(this, rowNode);
        this._rows.set(rowNumber, row);
        return row;
    }

    /**
     * Get the tab color. (See style [Color](#color).)
     * @returns {undefined|Color} The color or undefined if not set.
     *//**
     * Sets the tab color. (See style [Color](#color).)
     * @returns {Color|string|number} color - Color of the tab. If string, will set an RGB color. If number, will set a theme color.
     */
    tabColor() {
        return new ArgHandler("Sheet.tabColor", arguments)
            .case(() => {
                const tabColorNode = xmlq.findChild(this._sheetPrNode, "tabColor");
                if (!tabColorNode) return;

                const color = {};
                if (tabColorNode.attributes.hasOwnProperty('rgb')) color.rgb = tabColorNode.attributes.rgb;
                else if (tabColorNode.attributes.hasOwnProperty('theme')) color.theme = tabColorNode.attributes.theme;
                else if (tabColorNode.attributes.hasOwnProperty('indexed')) color.rgb = colorIndexes[tabColorNode.attributes.indexed];

                if (tabColorNode.attributes.hasOwnProperty('tint')) color.tint = tabColorNode.attributes.tint;

                return color;
            })
            .case("string", rgb => this.tabColor({ rgb }))
            .case("integer", theme => this.tabColor({ theme }))
            .case("nil", () => {
                xmlq.removeChild(this._sheetPrNode, "tabColor");
                return this;
            })
            .case("object", color => {
                const tabColorNode = xmlq.appendChildIfNotFound(this._sheetPrNode, "tabColor");
                xmlq.setAttributes(tabColorNode, {
                    rgb: color.rgb && color.rgb.toUpperCase(),
                    indexed: null,
                    theme: color.theme,
                    tint: color.tint
                });

                return this;
            })
            .handle();
    }

    /**
     * Gets a value indicating whether this sheet is selected.
     * @returns {boolean} True if selected, false if not.
     *//**
     * Sets whether this sheet is selected.
     * @param {boolean} selected - True to select, false to deselected.
     * @returns {Sheet} The sheet.
     */
    tabSelected() {
        const sheetViewNode = this._getOrCreateSheetViewNode();
        return new ArgHandler('Sheet.tabSelected', arguments)
            .case(() => {
                return sheetViewNode.attributes.tabSelected === 1;
            })
            .case('boolean', selected => {
                if (selected) sheetViewNode.attributes.tabSelected = 1;
                else delete sheetViewNode.attributes.tabSelected;
                return this;
            })
            .handle();
    }

    /**
     * Get the range of cells in the sheet that have contained a value or style at any point. Useful for extracting the entire sheet contents.
     * @returns {Range|undefined} The used range or undefined if no cells in the sheet are used.
     */
    usedRange() {
        const minRowNumber = Math.min(...this._rows.keys());
        const maxRowNumber = Math.max(...this._rows.keys());

        let minColumnNumber = 0;
        let maxColumnNumber = 0;

        this._rows.forEach(row => {
            if (!row) return;

            const minUsedColumnNumber = row.minUsedColumnNumber();
            const maxUsedColumnNumber = row.maxUsedColumnNumber();
            if (minUsedColumnNumber > 0 && (!minColumnNumber || minUsedColumnNumber < minColumnNumber)) minColumnNumber = minUsedColumnNumber;
            if (maxUsedColumnNumber > 0 && (!maxColumnNumber || maxUsedColumnNumber > maxColumnNumber)) maxColumnNumber = maxUsedColumnNumber;
        });

        // Return undefined if nothing in the sheet is used.
        if (minRowNumber <= 0 || minColumnNumber <= 0 || maxRowNumber <= 0 || maxColumnNumber <= 0) return;

        return this.range(minRowNumber, minColumnNumber, maxRowNumber, maxColumnNumber);
    }

    /**
     * Gets the parent workbook.
     * @returns {Workbook} The parent workbook.
     */
    workbook() {
        return this._workbook;
    }

    /**
     * Gets all page breaks.
     * @returns {{}} the object holds both vertical and horizontal PageBreaks.
     */
    pageBreaks() {
        return this._pageBreaks;
    }

    /**
     * Gets the vertical page breaks.
     * @returns {PageBreaks} vertical PageBreaks.
     */
    verticalPageBreaks() {
        return this._pageBreaks.colBreaks;
    }

    /**
     * Gets the horizontal page breaks.
     * @returns {PageBreaks} horizontal PageBreaks.
     */
    horizontalPageBreaks() {
        return this._pageBreaks.rowBreaks;
    }

    /* INTERNAL */

    /**
     * Clear cells that are using a given shared formula ID.
     * @param {number} sharedFormulaId - The shared formula ID.
     * @returns {undefined}
     * @ignore
     */
    clearCellsUsingSharedFormula(sharedFormulaId) {
        this._rows.forEach(row => {
            if (!row) return;
            row.clearCellsUsingSharedFormula(sharedFormulaId);
        });
    }

    /**
     * Get an existing column style ID.
     * @param {number} columnNumber - The column number.
     * @returns {undefined|number} The style ID.
     * @ignore
     */
    existingColumnStyleId(columnNumber) {
        // This will work after setting Column.style because Column updates the attributes live.
        const colNode = this._colNodes[columnNumber];
        return colNode && colNode.attributes.style;
    }

    /**
     * Call a callback for each column number that has a node defined for it.
     * @param {Function} callback - The callback.
     * @returns {undefined}
     * @ignore
     */
    forEachExistingColumnNumber(callback) {
        this._colNodes.forEach((node, columnNumber) => {
            if (!node) return;
            callback(columnNumber);
        });
    }

    /**
     * Call a callback for each existing row.
     * @param {Function} callback - The callback.
     * @returns {undefined}
     * @ignore
     */
    forEachExistingRow(callback) {
        this._rows.forEach((row, rowNumber) => {
            if (row) callback(row, rowNumber);
        });

        return this;
    }

    /**
     * Get the hyperlink attached to the cell with the given address.
     * @param {string} address - The address of the hyperlinked cell.
     * @returns {string|undefined|Reference} The hyperlink or undefined if not set.
     *//**
     * Set the hyperlink on the cell with the given address.
     * @param {string} address - The address of the hyperlinked cell.
     * @param {string} hyperlink - The hyperlink to set or undefined to clear.
     * @param {boolean} [internal] - The flag to force hyperlink to be internal. If true, then autodetect is skipped.
     * @returns {Sheet} The sheet.
     *//**
     * Set the hyperlink on the cell with the given address. If opts is a Cell an internal hyperlink is added.
     * @param {string} address - The address of the hyperlinked cell.
     * @param {object|Cell} opts - Options.
     * @returns {Sheet} The sheet.
     * @ignore
     *//**
     * Set the hyperlink on the cell with the given address and options.
     * @param {string} address - The address of the hyperlinked cell.
     * @param {{}|Cell} opts - Options or Cell. If opts is a Cell then an internal hyperlink is added.
     * @param {string|Cell} [opts.hyperlink] - The hyperlink to set, can be a Cell or an internal/external string.
     * @param {string} [opts.tooltip] - Additional text to help the user understand more about the hyperlink.
     * @param {string} [opts.email] - Email address, ignored if opts.hyperlink is set.
     * @param {string} [opts.emailSubject] - Email subject, ignored if opts.hyperlink is set.
     * @returns {Sheet} The sheet.
     */
    hyperlink() {
        return new ArgHandler('Sheet.hyperlink', arguments)
            .case('string', address => {
                const hyperlinkNode = this._hyperlinks.get(address);
                if (!hyperlinkNode) return;

                // internal reference
                const refersTo = hyperlinkNode.attributes.location;
                if (refersTo) {
                    // Try to parse the address.
                    const ref = this._hyperlinks.parse(refersTo);
                    if (!ref)
                        throw Error(`Cannot parse reference: ${refersTo}`);
                    return new Reference(ref, this.workbook());
                } else {
                    // external reference
                    const relationship = this._relationships.findById(hyperlinkNode.attributes['r:id']);
                    return {
                        hyperlink: relationship.attributes.Target,
                        tooltip: hyperlinkNode.attributes.tooltip
                    };
                }
            })
            .case(['string', 'nil'], address => {
                // TODO: delete relationship, make sure the relationship is not used by other thing.
                delete this._hyperlinks.delete(address);
                return this;
            })
            .case(['string', 'string'], (address, hyperlink) => {
                return this.hyperlink(address, hyperlink, false);
            })
            .case(['string', 'string', 'boolean'], (address, hyperlink, internal) => {
                const isHyperlinkInternalAddress = internal || addressConverter.fromAddress(hyperlink);
                let nodeAttributes;
                if (isHyperlinkInternalAddress) {
                    nodeAttributes = {
                        ref: address,
                        location: hyperlink,
                        display: hyperlink
                    };
                } else {
                    const relationship = this._relationships.add("hyperlink", hyperlink, "External");
                    nodeAttributes = {
                        ref: address,
                        'r:id': relationship.attributes.Id
                    };
                }
                this._hyperlinks.set(address, {
                    name: 'hyperlink',
                    attributes: nodeAttributes,
                    children: []
                });
                return this;
            })
            .case(['string', 'object'], (address, opts) => {
                if (opts instanceof Cell) {
                    const cell = opts;
                    const hyperlink = cell.address({ includeSheetName: true });
                    this.hyperlink(address, hyperlink, true);
                } else if (opts.hyperlink) {
                    this.hyperlink(address, opts.hyperlink);
                } else if (opts.email) {
                    const email = opts.email;
                    const subject = opts.emailSubject || '';
                    this.hyperlink(address, encodeURI(`mailto:${email}?subject=${subject}`));
                }
                const hyperlinkNode = this._hyperlinks.get(address);
                if (hyperlinkNode && opts.tooltip) {
                    hyperlinkNode.attributes.tooltip = opts.tooltip;
                }
                return this;
            })
            .handle();
    }

    /**
     * Increment and return the max shared formula ID.
     * @returns {number} The new max shared formula ID.
     * @ignore
     */
    incrementMaxSharedFormulaId() {
        return ++this._maxSharedFormulaId;
    }

    /**
     * Get a value indicating whether the cells in the given address are merged.
     * @param {string} address - The address to check.
     * @returns {Reference} A reference if merged, false if not merged.
     * @ignore
     *//**
     * Merge/unmerge cells by adding/removing a mergeCell entry.
     * @param {string} address - The address to merge.
     * @param {boolean} merged - True to merge, false to unmerge.
     * @returns {Sheet} The sheet.
     * @ignore
     */
    merged() {
        return new ArgHandler('Sheet.merge', arguments)
            .case('string', address => {
                const ref = this._mergeCells.get(address);
                return ref ? new Reference(ref, this.workbook()) : false;
            })
            .case(['string', '*'], (address, merge) => {
                if (merge) {
                    this._mergeCells.set(address, {
                        name: 'mergeCell',
                        attributes: { ref: address },
                        children: []
                    });
                } else {
                    this._mergeCells.delete(address);
                }
                return this;
            })
            .handle();
    }


    /**
     * Gets a Object or undefined of the cells in the given address.
     * @param {string} address - The address to check.
     * @returns {object|boolean} Object or false if not set
     * @ignore
     *//**
     * Removes dataValidation at the given address
     * @param {string} address - The address to remove.
     * @param {boolean} obj - false to delete.
     * @returns {boolean} true if removed.
     * @ignore
     *//**
     * Add dataValidation to cells at the given address if object or string
     * @param {string} address - The address to set.
     * @param {object|string} obj - Object or String to set
     * @returns {Sheet} The sheet.
     * @ignore
     */
    dataValidation() {
        return new ArgHandler('Sheet.dataValidation', arguments)
            .case('string', address => {
                return this._dataValidations.getDataValidation(address);
            })
            .case(['string', 'boolean'], (address, obj) => {
                const node = this._dataValidations.get(address);
                if (node) {
                    if (!obj) return this._dataValidations.delete(address);
                } else {
                    return false;
                }
            })
            .case(['string', '*'], (address, obj) => {
                this._dataValidations.setDataValidation(address, obj);
                return this;
            })
            .handle();
    }

    /**
     * Convert the sheet to a collection of XML objects.
     * @returns {{}} The XML forms.
     * @ignore
     */
    toXmls() {
        // Shallow clone the node so we don't have to remove these children later if they don't belong.
        const node = {...this._node};
        node.children = node.children.slice();

        // Add the columns if needed.
        this._colsNode.children = this._colNodes.filter((colNode, i) => {
            // Columns should only be present if they have attributes other than min/max.
            return colNode && i === colNode.attributes.min && Object.keys(colNode.attributes).length > 2;
        });
        if (this._colsNode.children.length) {
            xmlq.insertInOrder(node, this._colsNode, nodeOrder);
        }

        // Add the hyperlinks if needed.
        const hyperlinksNode = this._hyperlinks.toXml();
        if (hyperlinksNode.children.length) {
            xmlq.insertInOrder(node, hyperlinksNode, nodeOrder);
        }

        // Add the printOptions if needed.
        if (this._printOptionsNode) {
            if (Object.keys(this._printOptionsNode.attributes).length) {
                xmlq.insertInOrder(node, this._printOptionsNode, nodeOrder);
            }
        }

        // Add the pageMargins if needed.
        if (this._pageMarginsNode && this._pageMarginsPresetName) {
            // Clone to preserve the current state of this sheet.
            const childNode = {...this._pageMarginsNode};
            if (Object.keys(this._pageMarginsNode.attributes).length) {
                // Fill in any missing attribute values with presets.
                childNode.attributes = Object.assign(
                    this._pageMarginsPresets[this._pageMarginsPresetName],
                    this._pageMarginsNode.attributes);
            } else {
                // No need to fill in, all attributes is currently empty, simply replace.
                childNode.attributes = this._pageMarginsPresets[this._pageMarginsPresetName];
            }
            xmlq.insertInOrder(node, childNode, nodeOrder);
        }

        // Add the merge cells if needed.
        const mergeCellsNode = this._mergeCells.toXml();
        if (mergeCellsNode.children.length) {
            xmlq.insertInOrder(node, mergeCellsNode, nodeOrder);
        }

        // Add datavalidations node if needed.
        const dataValidationsNodes = this._dataValidations.toXml();
        if (dataValidationsNodes.node.children.length) {
            xmlq.insertInOrder(node, dataValidationsNodes.node, nodeOrder);
        }
        this._extensions.set(ExtURI.dataValidations, dataValidationsNodes.x14Node);

        // Add extlst node if needed
        const extensions = this._extensions.toXml();
        if (extensions.children.length) {
            xmlq.insertInOrder(node, extensions, nodeOrder);
        }

        if (this._autoFilter) {
            xmlq.insertInOrder(node, {
                name: "autoFilter",
                children: [],
                attributes: {
                    ref: this._autoFilter.address()
                }
            }, nodeOrder);
        }

        // Add the PageBreaks nodes if needed.
        ['colBreaks', 'rowBreaks'].forEach(name => {
            const breaks = this[`_${name}Node`];
            if (breaks.attributes.count) {
                xmlq.insertInOrder(node, breaks, nodeOrder);
            }
        });

        return {
            id: this._idNode,
            sheet: node,
            relationships: this._relationships
        };
    }

    /**
     * Update the max shared formula ID to the given value if greater than current.
     * @param {number} sharedFormulaId - The new shared formula ID.
     * @returns {undefined}
     * @ignore
     */
    updateMaxSharedFormulaId(sharedFormulaId) {
        if (sharedFormulaId > this._maxSharedFormulaId) {
            this._maxSharedFormulaId = sharedFormulaId;
        }
    }

    /**
     * Gets the shared formula ID, return undefined if not exists.
     * @param {number} sharedFormulaId - The attribute 'si', shared formula ID
     * @return {Cell|undefined} The reference cell
     * @ignore
     */
    getSharedFormulaRefCell(sharedFormulaId) {
        return this._sharedFormulaRefCells.get(sharedFormulaId);
    }

    /**
     * Sets the shared formula ID with a reference cell
     * @param {number} sharedFormulaId - The 'si', shared formula ID
     * @param {Cell} cell - The cell that the shared formula ID references.
     * @return {undefined}
     * @ignore
     */
    setSharedFormulaRefCell(sharedFormulaId, cell) {
        this._sharedFormulaRefCells.set(sharedFormulaId, cell);
    }

    /**
     * Get the print option given a valid print option attribute.
     * @param {string} attributeName - Attribute name of the printOptions.
     *   gridLines - Used in conjunction with gridLinesSet. If both gridLines and gridlinesSet are true, then grid lines shall print. Otherwise, they shall not (i.e., one or both have false values).
     *   gridLinesSet - Used in conjunction with gridLines. If both gridLines and gridLinesSet are true, then grid lines shall print. Otherwise, they shall not (i.e., one or both have false values).
     *   headings - Print row and column headings.
     *   horizontalCentered - Center on page horizontally when printing.
     *   verticalCentered - Center on page vertically when printing.
     * @returns {boolean}
     *//**
     * Set the print option given a valid print option attribute and a value.
     * @param {string} attributeName - Attribute name of the printOptions. See get print option for list of valid attributes.
     * @param {undefined|boolean} attributeEnabled - If `undefined` or `false` then the attribute is removed, otherwise the print option is enabled.
     * @returns {Sheet} The sheet.
     */
    printOptions() {
        const supportedAttributeNames = [
            'gridLines', 'gridLinesSet', 'headings', 'horizontalCentered', 'verticalCentered'];
        const checkAttributeName = this._getCheckAttributeNameHelper('printOptions', supportedAttributeNames);
        return new ArgHandler('Sheet.printOptions', arguments)
            .case(['string'], attributeName => {
                checkAttributeName(attributeName);
                return this._printOptionsNode.attributes[attributeName] === 1;
            })
            .case(['string', 'nil'], attributeName => {
                checkAttributeName(attributeName);
                delete this._printOptionsNode.attributes[attributeName];
                return this;
            })
            .case(['string', 'boolean'], (attributeName, attributeEnabled) => {
                checkAttributeName(attributeName);
                if (attributeEnabled) {
                    this._printOptionsNode.attributes[attributeName] = 1;
                    return this;
                } else {
                    return this.printOptions(attributeName, undefined);
                }
            })
            .handle();
    }

    /**
     * Get the print option for the gridLines attribute value.
     * @returns {boolean}
     *//**
     * Set the print option for the gridLines attribute value.
     * @param {undefined|boolean} enabled - If `undefined` or `false` then attribute is removed, otherwise gridLines is enabled.
     * @returns {Sheet} The sheet.
     */
    printGridLines() {
        return new ArgHandler('Sheet.gridLines', arguments)
            .case(() => {
                return this.printOptions('gridLines') && this.printOptions('gridLinesSet');
            })
            .case(['nil'], () => {
                this.printOptions('gridLines', undefined);
                this.printOptions('gridLinesSet', undefined);
                return this;
            })
            .case(['boolean'], enabled => {
                this.printOptions('gridLines', enabled);
                this.printOptions('gridLinesSet', enabled);
                return this;
            })
            .handle();
    }

    /**
     * Get the page margin given a valid attribute name.
     * If the value is not yet defined, then it will return the current preset value.
     * @param {string} attributeName - Attribute name of the pageMargins.
     *     left - Left Page Margin in inches.
     *     right - Right page margin in inches.
     *     top - Top Page Margin in inches.
     *     buttom - Bottom Page Margin in inches.
     *     footer - Footer Page Margin in inches.
     *     header - Header Page Margin in inches.
     * @returns {number} the attribute value.
     *//**
     * Set the page margin (or override the preset) given an attribute name and a value.
     * @param {string} attributeName - Attribute name of the pageMargins. See get page margin for list of valid attributes.
     * @param {undefined|number|string} attributeStringValue - If `undefined` then set back to preset value, otherwise, set the given attribute value.
     * @returns {Sheet} The sheet.
     */
    pageMargins() {
        if (this.pageMarginsPreset() === undefined) {
            throw new Error('Sheet.pageMargins: preset is undefined.');
        }
        const supportedAttributeNames = [
            'left', 'right', 'top', 'bottom', 'header', 'footer'];
        const checkAttributeName = this._getCheckAttributeNameHelper('pageMargins', supportedAttributeNames);
        const checkRange = this._getCheckRangeHelper('pageMargins', 0, undefined);
        return new ArgHandler('Sheet.pageMargins', arguments)
            .case(['string'], attributeName => {
                checkAttributeName(attributeName);
                const attributeValue = this._pageMarginsNode.attributes[attributeName];
                if (attributeValue !== undefined) {
                    return parseFloat(attributeValue);
                } else if (this._pageMarginsPresetName) {
                    return parseFloat(this._pageMarginsPresets[this._pageMarginsPresetName][attributeName]);
                } else {
                    return undefined;
                }
            })
            .case(['string', 'nil'], attributeName => {
                checkAttributeName(attributeName);
                delete this._pageMarginsNode.attributes[attributeName];
                return this;
            })
            .case(['string', 'number'], (attributeName, attributeNumberValue) => {
                checkAttributeName(attributeName);
                checkRange(attributeNumberValue);
                this._pageMarginsNode.attributes[attributeName] = attributeNumberValue;
                return this;
            })
            .case(['string', 'string'], (attributeName, attributeStringValue) => {
                return this.pageMargins(attributeName, parseFloat(attributeStringValue));
            })
            .handle();
    }

    /**
     * Page margins preset is a set of page margins associated with a name.
     * The page margin preset acts as a fallback when not explicitly defined by `Sheet.pageMargins`.
     * If a sheet already contains page margins, it attempts to auto-detect, otherwise they are defined as the template preset.
     * If no page margins exist, then the preset is undefined and will not be included in the output of `Sheet.toXmls`.
     * Available presets include: normal, wide, narrow, template.
     *
     * Get the page margins preset name. The registered name of a predefined set of attributes.
     * @returns {string} The preset name.
     *//**
     * Set the page margins preset by name, clearing any existing/temporary attribute values.
     * @param {undefined|string} presetName - The preset name. If `undefined`, page margins will not be included in the output of `Sheet.toXmls`.
     * @returns {Sheet} The sheet.
     *//**
     * Set a new page margins preset by name and attributes object.
     * @param {string} presetName - The preset name.
     * @param {object} presetAttributes - The preset attributes.
     * @returns {Sheet} The sheet.
     */
    pageMarginsPreset() {
        return new ArgHandler('Sheet.pageMarginsPreset', arguments)
            .case(() => {
                return this._pageMarginsPresetName;
            })
            .case(['nil'], () => {
                // Remove all preset overrides and exclude from sheet
                this._pageMarginsPresetName = undefined;

                // Remove all preset overrides
                this._pageMarginsNode.attributes = {};
                return this;
            })
            .case(['string'], presetName => {
                const checkPresetName = this._getCheckAttributeNameHelper(
                    'pageMarginsPreset', Object.keys(this._pageMarginsPresets));
                checkPresetName(presetName);

                // Change to new preset
                this._pageMarginsPresetName = presetName;

                // Remove all preset overrides
                this._pageMarginsNode.attributes = {};
                return this;
            })
            .case(['string', 'object'], (presetName, presetAttributes) => {
                if (this._pageMarginsPresets.hasOwnProperty(presetName)) {
                    throw new Error(`Sheet.pageMarginsPreset: The preset ${presetName} already exists!`);
                }

                // Validate preset attribute keys.
                const pageMarginsAttributeNames = [
                    'left', 'right', 'top', 'bottom', 'header', 'footer'];
                const isValidPresetAttributeKeys =
                    Object.keys(presetAttributes).every(name => pageMarginsAttributeNames.includes(name))
                if (isValidPresetAttributeKeys === false) {
                    throw new Error(`Sheet.pageMarginsPreset: Invalid preset attributes for one or key(s)! - "${Object.keys(presetAttributes)}"`);
                }

                // Validate preset attribute values.
                Object.values(presetAttributes).forEach((attributeValue) => {
                    const attributeNumberValue = parseFloat(attributeValue);
                    if (isNaN(attributeNumberValue) || typeof attributeNumberValue !== 'number') {
                        throw new Error(`Sheet.pageMarginsPreset: Invalid preset attribute value! - "${attributeValue}"`);
                    }
                });

                // Change to new preset
                this._pageMarginsPresetName = presetName;

                // Remove all preset overrides
                this._pageMarginsNode.attributes = {};

                // Register the preset
                this._pageMarginsPresets[presetName] = presetAttributes;
                return this;
            })
            .handle();
    }

    /**
     * https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.pane?view=openxml-2.8.1
     * @typedef {Object} PaneOptions
     * @property {string} activePane=bottomRight Active Pane. The pane that is active.
     * @property {string} state Split State. Indicates whether the pane has horizontal / vertical splits,
     * and whether those splits are frozen.
     * @property {string} topLeftCell Top Left Visible Cell. Location of the top left visible cell in the bottom
     * right pane (when in Left-To-Right mode).
     * @property {number} xSplit (Horizontal Split Position) Horizontal position of the split, in 1/20th of a point;
     * 0 (zero) if none. If the pane is frozen, this value indicates the number of columns visible in the top pane.
     * @property {number} ySplit (Vertical Split Position) Vertical position of the split, in 1/20th of a point; 0
     * (zero) if none. If the pane is frozen, this value indicates the number of rows visible in the left pane.
     *//**
     * Gets sheet view pane options
     * @return {PaneOptions} sheet view pane options
     *//**
     * Sets sheet view pane options
     * @param {PaneOptions|null|undefined} paneOptions sheet view pane options
     * @return {Sheet} The sheet
     */
    panes() {
        const supportedStates = ['split', 'frozen', 'frozenSplit'];
        const supportedActivePanes = ['bottomLeft', 'bottomRight', 'topLeft', 'topRight'];
        const checkStateName = this._getCheckAttributeNameHelper('pane.state', supportedStates);
        const checkActivePane = this._getCheckAttributeNameHelper('pane.activePane', supportedActivePanes);
        const sheetViewNode = this._getOrCreateSheetViewNode();
        let paneNode = xmlq.findChild(sheetViewNode, 'pane');
        return new ArgHandler('Sheet.pane', arguments)
            .case(() => {
                if (paneNode) {
                    const result = cloneDeep(paneNode.attributes);
                    if (!result.state) result.state = 'split';
                    return result;
                }
            })
            .case(['nil'], () => {
                xmlq.removeChild(sheetViewNode, 'pane');
                return this;
            })
            .case(['object'], paneAttributes => {
                const attributes = Object.assign({ activePane: 'bottomRight' }, paneAttributes);
                checkStateName(attributes.state);
                checkActivePane(attributes.activePane);
                if (paneNode) {
                    paneNode.attributes = attributes;
                } else {
                    paneNode = {
                        name: "pane",
                        attributes,
                        children: []
                    };
                    xmlq.appendChild(sheetViewNode, paneNode);
                }
                return this;
            })
            .handle();
    }

    /**
     * Freezes Panes for this sheet.
     * @param {number} xSplit the number of columns visible in the top pane. 0 (zero) if none.
     * @param {number} ySplit the number of rows visible in the left pane. 0 (zero) if none.
     * @return {Sheet} The sheet
     *//**
     * freezes Panes for this sheet.
     * @param {string} topLeftCell Top Left Visible Cell. Location of the top left visible cell in the bottom
     * right pane (when in Left-To-Right mode).
     * @return {Sheet} The sheet
     */
    freezePanes() {
        return new ArgHandler('Sheet.feezePanes', arguments)
            .case(['integer', 'integer'], (xSplit, ySplit) => {
                const topLeftCell = addressConverter.columnNumberToName(xSplit + 1) + (ySplit + 1);
                let activePane = xSplit === 0 ? 'bottomLeft' : 'bottomRight';
                activePane = ySplit === 0 ? 'topRight' : activePane;
                return this.panes({ state: 'frozen', topLeftCell, xSplit, ySplit, activePane });
            })
            .case(['string'], topLeftCell => {
                const ref = addressConverter.fromAddress(topLeftCell);
                const xSplit = ref.columnNumber - 1, ySplit = ref.rowNumber - 1;
                let activePane = xSplit === 0 ? 'bottomLeft' : 'bottomRight';
                activePane = ySplit === 0 ? 'topRight' : activePane;
                return this.panes({ state: 'frozen', topLeftCell, xSplit, ySplit, activePane });
            })
            .handle();
    }

    /**
     * Splits Panes for this sheet.
     * @param {number} xSplit (Horizontal Split Position) Horizontal position of the split,
     * in 1/20th of a point; 0 (zero) if none.
     * @param {number} ySplit (Vertical Split Position) VVertical position of the split,
     * in 1/20th of a point; 0 (zero) if none.
     * @return {Sheet} The sheet
     */
    splitPanes(xSplit, ySplit) {
        return this.panes({ state: 'split', xSplit, ySplit });
    }

    /**
     * resets to default sheet view panes.
     * @return {Sheet} The sheet
     */
    resetPanes() {
        return this.panes(null);
    }

    // Add a category
    addRow(rowNumber) {
      const rowNode = {
        name: 'row',
        attributes: {
          r: rowNumber
        },
        children: [],
      };
      const keys = Array.from(this._rows.keys()).sort((n1, n2) => n1 - n2);
      const totalChanges = [];
      //TODO: can be optimize(REDUCE LOOP)
      for (let k = keys.length - 1; k >=0; k--){
        const i = keys[k];
        if (i >= rowNumber) {
          const row = this._rows.get(i);

          row._cells.forEach(cell => {

            const changes = this._workbook._refTable.getCalculationOrder(cell.getRef(), totalChanges);
            changes.forEach(change => {
              totalChanges.push(change);
              const cell = this._workbook.sheet(change.sheet).row(change.row).cell(change.col);
              if (cell.isSharedFormula()) cell.sharedFormulaToNormalFormula();
              const newFormula = FormulaReplacer.replaceRowNumber(cell.getFormula(), change.sheet,
                this.name(), rowNumber, 1);


              // only rebuild reference table, do not perform formula calculations
              cell.setFormula(newFormula, false);
            });
          });
        }
      }

      // re-assign index
      for (let k = keys.length - 1; k >=0; k--){
        const i = keys[k];

        //This code is kind of Hard code.
        if (i >= rowNumber) {
          const row = this._rows.get(i);
          row._cells.forEach(cell => {
            const formula = cell.getFormula();
            if(formula) {
              cell.clear();
            }
            cell._formula = formula;
          });
          row._node.attributes.r = i + 1;
          this._rows.set(i + 1, row);

          row._cells.forEach(cell => {
            const formula = cell._formula;
            if(formula) {
              cell.setFormula(formula,true);
            }
          });
          this._rows.delete(i);
        }
      }

      const newRow = new Row(this, rowNode);
      this._rows.set(rowNumber, newRow);
      const copyRow = this.row(rowNumber + 1);
      copyRow._cells.forEach(cell => {
        const newCell = newRow.cell(cell.columnNumber());
        newCell._style = cell._style;
        newCell._styleId = cell._styleId;
      });
      for (let i = 0; i < totalChanges.length; i++) {
        const change = totalChanges[i];
        if (change.row >= rowNumber) {
          const cell = this._workbook.sheet(change.sheet).row(change.row + 1).cell(change.col);
          cell.recalculate();
        }
        else {
          const cell = this._workbook.sheet(change.sheet).row(change.row).cell(change.col);
          cell.recalculate();
        }
        // perform calculations

      }
      return rowNode
    }

    deleteRow(rowNumber) {
        // TODO: update defined names, merged cells
        const deletedRow = this._rows.get(rowNumber);
        const totalChangesPart1 = [];

        // make the cell formulas that references the deletedCells #REF, or reduce range reference
        deletedRow._cells.forEach(cell => {
          if (typeof cell != "undefined") {
            const ref = cell.getRef();
            const changes = this._workbook._refTable.getDirectReferences(ref);
            changes.forEach(change => {
              totalChangesPart1.push(change);
              const cell = this._workbook.sheet(change.sheet).row(change.row).cell(change.col);
              // convert shared formulas to normal formulas
              if (cell.isSharedFormula()) cell.sharedFormulaToNormalFormula();
              const newFormula = FormulaReplacer.removeReference(cell.getFormula(), change.sheet,
                ref, FormulaReplacer.MODE.ROW);
              cell.setFormula(newFormula, false);
            });
          }
        });
        this._rows.delete(rowNumber);
        const keys = Array.from(this._rows.keys()).sort((n1, n2) => n1 - n2);
        const totalChangesPart2 = [];
        //TODO: can be optimize(REDUCE LOOP)
        for (const i of keys){
            if (i > rowNumber) {
                const row = this._rows.get(i);

                row._cells.forEach(cell => {
                    // if (typeof cell != "undefined") {
                    //   const ref = cell.getRef();
                    //   const changes = this._workbook._refTable.getDirectReferences(ref);
                    //   changes.forEach(change => {
                    //     // TODO: a more efficient query using Map
                    //     const idxExist = totalChanges.findIndex(element => {
                    //       // ensure calculation on a cell won't happen many times.
                    //       return element.row === change.row && element.col === change.col && element.sheet === change.sheet;
                    //     });
                    //     if (idxExist === -1) {
                    //       totalChanges.push(change);
                    //       const cell = this._workbook.sheet(change.sheet).row(change.row).cell(change.col);
                    //
                    //       // convert shared formulas to normal formulas
                    //       if (cell.isSharedFormula()) cell.sharedFormulaToNormalFormula();
                    //
                    //       // normal cell
                    //       const newFormula = FormulaReplacer.replaceRowNumber(cell.getFormula(), change.sheet,
                    //         this.name(), rowNumber, -1);
                    //       cell.setFormula(newFormula, false);
                    //     }
                    //   });
                    // }
                  const changes = this._workbook._refTable.getCalculationOrder(cell.getRef(), totalChangesPart2);
                  changes.forEach(change => {
                    totalChangesPart2.push(change);
                    const cell = this._workbook.sheet(change.sheet).row(change.row).cell(change.col);
                    if (cell.isSharedFormula()) cell.sharedFormulaToNormalFormula();
                    const newFormula = FormulaReplacer.replaceRowNumber(cell.getFormula(), change.sheet,
                      this.name(), rowNumber, -1);


                    // only rebuild reference table, do not perform formula calculations
                    cell.setFormula(newFormula, false);
                  });
                });
            }
        }

      // re-assign index
      for (const i of keys){

        //This code is kind of Hard code.
        if (i > rowNumber) {
          const row = this._rows.get(i);
          row._cells.forEach(cell => {
            const formula = cell.getFormula();
            if(formula) {
              cell.clear();
            }
            cell._formula = formula;
          });
          row._node.attributes.r = i - 1;
          this._rows.set(i - 1, row);

          row._cells.forEach(cell => {
            const formula = cell._formula;
            if(formula) {
              cell.setFormula(formula,true);
            }
          });
          this._rows.delete(i);
        }
      }

      // perform calculations
      for (let i = 0; i < totalChangesPart1.length; i++) {
        const change = totalChangesPart1[i];
        const cell = this._workbook.sheet(change.sheet).row(change.row).cell(change.col);
        cell.recalculate();
      }
      for (let i = 0; i < totalChangesPart2.length; i++) {
        const change = totalChangesPart2[i];
        if (change.row > rowNumber) {
          const cell = this._workbook.sheet(change.sheet).row(change.row - 1).cell(change.col);
          cell.recalculate();
        }
        else {
          const cell = this._workbook.sheet(change.sheet).row(change.row).cell(change.col);
          cell.recalculate();
        }
      }
    }

    addColumn(colNumber) {

      const totalChanges = [];
      const keys = Array.from(this._rows.keys()).sort((n1, n2) => n1 - n2);
      for(const i of keys) {
        const row = this._rows.get(i);
        const keys = Array.from(row._cells.keys())
          .sort((n1, n2) => n1 - n2);
        for (let k = keys.length - 1; k >= 0; k--) {
          const i = keys[k];
          if (i >= colNumber) {
            const cell = row._cells.get(i);


            const changes = this._workbook._refTable.getCalculationOrder(cell.getRef(), totalChanges);
            changes.forEach(change => {
              totalChanges.push(change);
              const cell = this._workbook.sheet(change.sheet).row(change.row).cell(change.col);
              if (cell.isSharedFormula()) cell.sharedFormulaToNormalFormula();

              const newFormula = FormulaReplacer.replaceColNumber(cell.getFormula(), change.sheet,
                this.name(), colNumber, 1);

              // only rebuild reference table, do not perform formula calculations
              cell.setFormula(newFormula, false);

            });
          }
        }
      }

      // re-assign index
      let maxCol = 0;
      for(const i of keys) {
        const row = this._rows.get(i);
        const keys = Array.from(row._cells.keys())
          .sort((n1, n2) => n1 - n2);
        for (let k = keys.length - 1; k >= 0; k--) {
          const i = keys[k];
          if(i > maxCol) {maxCol = i}
          if (i >= colNumber) {
            const cell = row._cells.get(i);
            const formula = cell.getFormula();
            if (formula) {
              cell.clear();
            }
            cell._columnNumber = cell._columnNumber + 1;
            row._cells.set(i + 1, cell);
            if(formula)
              cell.setFormula(formula,true);
            row._cells.delete(i);
          }
        }


        const newCell = new Cell(row, colNumber);
        const copyCell = newCell.row().cell(colNumber + 1);
        row._cells.set(colNumber, newCell);
        row._node.children = row._cells;
        newCell._style = copyCell._style;
        newCell._styleId = copyCell._styleId;

      }
      for (let i = maxCol; i >= colNumber; i= i-1) {
        this.column(i+1).width(this.column(i).width());
        this.column(i+1).hidden(this.column(i).hidden());
      }

      // perform calculations
      for (let i = 0; i < totalChanges.length; i++) {
          const change = totalChanges[i];
          if (change.col >= colNumber) {
            const cell = this._workbook.sheet(change.sheet).row(change.row).cell(change.col + 1);
            cell.recalculate();
          }
          else {
            const cell = this._workbook.sheet(change.sheet).row(change.row).cell(change.col);
            cell.recalculate();
          }
      }
    }


    // delete a column
    deleteColumn(colNumber) {
      const keys = Array.from(this._rows.keys())
        .sort((n1, n2) => n1 - n2);

      const totalChangesPart1 = [];
      for (const i of keys) {
        const row = this._rows.get(i);
        const cell = row._cells.get(colNumber);
        if (typeof cell != "undefined") {
          const ref = cell.getRef();
          const changes = this._workbook._refTable.getDirectReferences(ref);

          changes.forEach(change => {
            totalChangesPart1.push(change);
            const cell1 = this._workbook.sheet(change.sheet).row(change.row).cell(change.col);
            // convert shared formulas to normal formulas
            if (cell1.isSharedFormula()) cell1.sharedFormulaToNormalFormula();
            const newFormula = FormulaReplacer.removeReference(cell1.getFormula(), change.sheet,
              ref, FormulaReplacer.MODE.ROW);
            cell1.setFormula(newFormula, false);
          });
        }
      }

      const totalChangesPart2 = [];
      for (const i of keys) {
        const row = this._rows.get(i);
        const keys = Array.from(row._cells.keys())
          .sort((n1, n2) => n1 - n2);
        for (const i of keys) {
          if (i > colNumber) {
            const cell = row._cells.get(i);
            if (typeof cell != "undefined") {

              const changes = this._workbook._refTable.getCalculationOrder(cell.getRef(), totalChangesPart2);
              changes.forEach(change => {
                totalChangesPart2.push(change);
                const cell = this._workbook.sheet(change.sheet).row(change.row).cell(change.col);
                if (cell.isSharedFormula()) cell.sharedFormulaToNormalFormula();

                const newFormula = FormulaReplacer.replaceColNumber(cell.getFormula(), change.sheet,
                  this.name(), colNumber, -1);

                // only rebuild reference table, do not perform formula calculations
                cell.setFormula(newFormula, false);

              });
            }
          }
        }
      }

      // re-assign index
      let maxCol = 0;
      for(const i of keys) {
        const row = this._rows.get(i);
        const keys = Array.from(row._cells.keys())
          .sort((n1, n2) => n1 - n2);
        for (const i of keys) {
          if (i > colNumber) {
            if(i > maxCol) {maxCol = i}
            const cell = row._cells.get(i);
            const formula = cell.getFormula();
            if(formula) {
              cell.clear();
            }
            cell._columnNumber = cell._columnNumber - 1;
            row._cells.set(i - 1, cell);
            if(formula)
              cell.setFormula(formula,true);
            row._cells.delete(i);
          }
        }
      }

      for (let i = colNumber; i < maxCol; i++) {
        this.column(i).width(this.column(i+1).width());
        this.column(i).hidden(this.column(i+1).hidden());
      }
      // perform calculations
      for (let i = 0; i < totalChangesPart1.length; i++) {
        const change = totalChangesPart1[i];
        const cell = this._workbook.sheet(change.sheet).row(change.row).cell(change.col);
        cell.recalculate();
      }
      for (let i = 0; i < totalChangesPart2.length; i++) {
        const change = totalChangesPart2[i];
        if (change.col > colNumber) {
          const cell = this._workbook.sheet(change.sheet).row(change.row).cell(change.col - 1);
          cell.recalculate();
        }
        else {
          const cell = this._workbook.sheet(change.sheet).row(change.row).cell(change.col);
          cell.recalculate();
        }
      }
    }

    setColPermit(colNum, permit){
      const keys = Array.from(this._rows.keys())
        .sort((n1, n2) => n1 - n2);
      for (const i of keys) {
        const row = this._rows.get(i);
        const cell = row._cells.get(colNumber);
      }
    }

    /**
     * Get all the cells need to be prepopulate
     */
    getPrepopulateCell() {
      const keys = Array.from(this._rows.keys())
        .sort((n1, n2) => n1 - n2);
      let cells = [];
      for (const i of keys) {
        const row = this._rows.get(i);
        const catId = row.cell(1).getValue();
        if (catId !== undefined) {
          row.setCategoryId(catId);
          const keys = Array.from(row._cells.keys())
            .sort((n1, n2) => n1 - n2);
          for (const j of keys) {
            const cell = row._cells.get(j);
            const attCell = this.row(1).cell(j);
            if (attCell !== undefined)
            {
              cell.column().setAttributeId(attCell.getValue());
              if(cell.row().getCategoryId() !== undefined && cell.column().getAttributeId() !== undefined)
                cells.push(cell);
            }
          }
        }
      }

      return cells;
    }

    /* PRIVATE */

    /**
     * Get a helper function to check that the attribute name provided is supported.
     * @param {string} functionName - Name of the parent function.
     * @param {array} supportedAttributeNames - Array of supported attribute name strings.
     * @returns {function} The helper function, which takes an attribute name. If the array of supported attribute names does not contain the given attribute name, then an Error is thrown.
     * @ignore
     */
    _getCheckAttributeNameHelper(functionName, supportedAttributeNames) {
        return attributeName => {
            if (!supportedAttributeNames.includes(attributeName)) {
                throw new Error(`Sheet.${functionName}: "${attributeName}" is not supported.`);
            }
        };
    }

    /**
     * Get a helper function to check that the value is of the expected type.
     * @param {string} functionName - Name of the parent function.
     * @param {string} valueType - A string produced by typeof.
     * @returns {function} The helper function, which takes a value. If the value type is not expected, a TypeError is thrown.
     * @ignore
     */
    _getCheckTypeHelper(functionName, valueType) {
        return value => {
            if (typeof value !== valueType) {
                throw new TypeError(`Sheet.${functionName}: invalid type - value must be of type ${valueType}.`);
            }
        };
    }

    /**
     * Get a helper function to check that the value is within the expected range.
     * @param {string} functionName - Name of the parent function.
     * @param {undefined|number} valueMin - The minimum value of the range. This value is range-inclusive.
     * @param {undefined|number} valueMax - The maximum value of the range. This value is range-exclusive.
     * @returns {function} The helper function, which takes a value. If the value type is not 'number', a TypeError is thrown. If the value is not within the range, a RangeError is thrown.
     * @ignore
     */
    _getCheckRangeHelper(functionName, valueMin, valueMax) {
        const checkType = this._getCheckTypeHelper(functionName, 'number');
        return value => {
            checkType(value);
            if (valueMin !== undefined) {
                if (value < valueMin) {
                    throw new RangeError(`Sheet.${functionName}: value too small - value must be greater than or equal to ${valueMin}.`);
                }
            }
            if (valueMax !== undefined) {
                if (valueMax <= value) {
                    throw new RangeError(`Sheet.${functionName}: value too large - value must be less than ${valueMax}.`);
                }
            }
        };
    }

    /**
     * Get the sheet view node if it exists or create it if it doesn't.
     * @returns {{}} The sheet view node.
     * @private
     */
    _getOrCreateSheetViewNode() {
        let sheetViewsNode = xmlq.findChild(this._node, "sheetViews");
        if (!sheetViewsNode) {
            sheetViewsNode = {
                name: "sheetViews",
                attributes: {},
                children: [{
                    name: "sheetView",
                    attributes: {
                        workbookViewId: 0
                    },
                    children: []
                }]
            };

            xmlq.insertInOrder(this._node, sheetViewsNode, nodeOrder);
        }

        return xmlq.findChild(sheetViewsNode, "sheetView");
    }

    /**
     * Initializes the sheet.
     * @param {Workbook} workbook - The parent workbook.
     * @param {{}} idNode - The sheet ID node (from the parent workbook).
     * @param {{}} node - The sheet node.
     * @param {{}} [relationshipsNode] - The optional sheet relationships node.
     * @returns {Promise<any>}
     * @ignore
     */
    initAsync() {
        return Promise.resolve()
            .then(() => {
                const workbook = this._workbook;
                const idNode = this._idNode;
                let node = this._node;
                const relationshipsNode = this._relationshipsNode;
                if (!node) {
                    node = {
                        name: "worksheet",
                        attributes: {
                            xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                            'xmlns:r': "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                            'xmlns:mc': "http://schemas.openxmlformats.org/markup-compatibility/2006",
                            'mc:Ignorable': "x14ac",
                            'xmlns:x14ac': "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
                        },
                        children: [{
                            name: "sheetData",
                            attributes: {},
                            children: []
                        }]
                    };
                }

                this._workbook = workbook;
                this._idNode = idNode;
                this._node = node;
                this._maxSharedFormulaId = -1;
                this._sharedFormulaRefCells = new Map();
                this._mergeCells = {};
                this._dataValidations = {};
                this._hyperlinks = {};
                this._autoFilter = null;

                // Create the relationships.
                this._relationships = new Relationships(relationshipsNode);

                // Delete the optional dimension node
                xmlq.removeChild(this._node, "dimension");
            })
            .then(() => {
                // Create the rows.
                this._rows = new Map();
                this._sheetDataNode = xmlq.findChild(this._node, "sheetData");

                // Divide into separate tasks, to prevent block UI thread in browser
                const splitedRows = [];
                const numRows = this._sheetDataNode.children.length, divider = 70; // tweak this value, 1 -> 100
                for (let i = 0; i < Math.ceil(numRows / divider); i++) {
                    const endIndex = numRows < (i + 1) * divider ? numRows : (i + 1) * divider;
                    splitedRows.push(this._sheetDataNode.children.slice(i * divider, endIndex));
                }
                if (splitedRows.length === 1) {
                    // does not create a new task on small amount of rows
                    const rowNodes = splitedRows[0];
                    for (let j = 0; j < rowNodes.length; j++) {
                        const row = new Row(this, rowNodes[j]);
                        this._rows.set(row.rowNumber(), row);
                    }
                } else if (splitedRows.length > 1) {
                    return Promise.all(splitedRows.map(rowNodes => {
                        return new Promise(resolve => {
                            setTimeout(() => {
                                for (let j = 0; j < rowNodes.length; j++) {
                                    const row = new Row(this, rowNodes[j]);
                                    this._rows.set(row.rowNumber(), row);
                                }
                                resolve();
                            });
                        });
                    }));
                }
            }).then(() => {
                // store rows
                this._sheetDataNode.children = this._rows;

                // Create the columns node.
                this._columns = [];
                this._colsNode = xmlq.findChild(this._node, "cols");
                if (this._colsNode) {
                    xmlq.removeChild(this._node, this._colsNode);
                } else {
                    this._colsNode = { name: 'cols', attributes: {}, children: [] };
                }

                // Cache the col nodes.
                this._colNodes = [];
                this._colsNode.children.forEach(colNode => {
                    const min = colNode.attributes.min;
                    const max = colNode.attributes.max;
                    for (let i = min; i <= max; i++) {
                        this._colNodes[i] = colNode;
                    }
                });

                // Create the sheet properties node.
                this._sheetPrNode = xmlq.findChild(this._node, "sheetPr");
                if (!this._sheetPrNode) {
                    this._sheetPrNode = { name: 'sheetPr', attributes: {}, children: [] };
                    xmlq.insertInOrder(this._node, this._sheetPrNode, nodeOrder);
                }

                // Create the merge cells.
                const mergeCellsNode = xmlq.findChild(this._node, "mergeCells");
                if (mergeCellsNode) {
                    xmlq.removeChild(this._node, mergeCellsNode);
                }
                this._mergeCells = new MergeCells(this, mergeCellsNode);

                const extLst = xmlq.findChild(this._node, 'extLst');
                if (extLst) {
                    xmlq.removeChild(this._node, extLst);
                }
                this._extensions = new Extensions(extLst);

                // Create the DataValidations.
                const dataValidationsNode = xmlq.findChild(this._node, "dataValidations");
                if (dataValidationsNode) {
                    xmlq.removeChild(this._node, dataValidationsNode);
                }
                this._dataValidations = new DataValidations(this, dataValidationsNode,
                    this._extensions.get(ExtURI.dataValidations));

                // Create the hyperlinks.
                const hyperlinksNode = xmlq.findChild(this._node, "hyperlinks");
                if (hyperlinksNode) {
                    xmlq.removeChild(this._node, hyperlinksNode);
                }
                this._hyperlinks = new Hyperlinks(this, hyperlinksNode);


                // Create the printOptions.
                this._printOptionsNode = xmlq.findChild(this._node, "printOptions");
                if (this._printOptionsNode) {
                    xmlq.removeChild(this._node, this._printOptionsNode);
                } else {
                    this._printOptionsNode = { name: 'printOptions', attributes: {}, children: [] };
                }


                // Create the pageMargins.
                this._pageMarginsPresets = {
                    normal: {
                        left: 0.7,
                        right: 0.7,
                        top: 0.75,
                        bottom: 0.75,
                        header: 0.3,
                        footer: 0.3
                    },
                    wide: {
                        left: 1,
                        right: 1,
                        top: 1,
                        bottom: 1,
                        header: 0.5,
                        footer: 0.5
                    },
                    narrow: {
                        left: 0.25,
                        right: 0.25,
                        top: 0.75,
                        bottom: 0.75,
                        header: 0.3,
                        footer: 0.3
                    }
                };
                this._pageMarginsNode = xmlq.findChild(this._node, "pageMargins");
                if (this._pageMarginsNode) {
                    // Sheet has page margins, assume preset is template.
                    this._pageMarginsPresetName = 'template';

                    // Search for a preset that matches existing attributes.
                    for (const presetName in this._pageMarginsPresets) {
                        if (Object.is(this._pageMarginsNode.attributes, this._pageMarginsPresets[presetName])) {
                            this._pageMarginsPresetName = presetName;
                            break;
                        }
                    }

                    // If template preset, then register as template preset, and clear attributes.
                    if (this._pageMarginsPresetName === 'template') {
                        this._pageMarginsPresets.template = this._pageMarginsNode.attributes;
                        this._pageMarginsNode.attributes = {};
                    }

                    xmlq.removeChild(this._node, this._pageMarginsNode);
                } else {
                    // Sheet has no page margins, the preset assignment is therefore undefined.
                    this._pageMarginsPresetName = undefined;
                    this._pageMarginsNode = { name: 'pageMargins', attributes: {}, children: [] };
                }

                // Create the pageBreaks
                ['colBreaks', 'rowBreaks'].forEach(name => {
                    this[`_${name}Node`] = xmlq.findChild(this._node, name);
                    if (this[`_${name}Node`]) {
                        xmlq.removeChild(this._node, this[`_${name}Node`]);
                    } else {
                        this[`_${name}Node`] = {
                            name,
                            children: [],
                            attributes: {
                                count: 0,
                                manualBreakCount: 0
                            }
                        };
                    }
                });
                this._pageBreaks = {
                    colBreaks: new PageBreaks(this._colBreaksNode),
                    rowBreaks: new PageBreaks(this._rowBreaksNode)
                };
            });
      this.totalcol = this._columns.length;
      this.totalraw = this._rows.size;
    }
}

module.exports = Sheet;

/*
xl/workbook.xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet ...>
    ...

    <printOptions headings="1" gridLines="1" />
    <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />
    <pageSetup orientation="portrait" horizontalDpi="0" verticalDpi="0" />
</worksheet>
// */
