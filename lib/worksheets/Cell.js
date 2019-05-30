"use strict";

const _ = require("lodash");
const ArgHandler = require("../ArgHandler");
const addressConverter = require("../addressConverter");
const dateConverter = require("../dateConverter");
const regexify = require("../regexify");
const xmlq = require("../xml/xmlq");
const FormulaError = require("../FormulaError");
const Style = require("./Style");
const RichText = require("./RichText");

/**
 * A cell
 */
class Cell {
    // /**
    //  * Creates a new instance of cell.
    //  * @param {Row} row - The parent row.
    //  * @param {{}} node - The cell node.
    //  */
    constructor(row, node, styleId) {
        this._row = row;
        this._init(node, styleId);
    }

    /* PUBLIC */

    /**
     * Gets the value of the cell.
     * @returns {string|boolean|number|Date|undefined|RichText} The value of the cell.
     */
    getValue() {
        if (this._value instanceof RichText) {
            return this._value.getInstanceWithCellRef(this);
        }
        return this._value;
    }

    /**
     * Sets the value of the cell.
     * @param {RichText|string|boolean|number|null|undefined} value - The value to set.
     * @param {boolean} [clear] - If clear this cell.
     * @returns {Array.<{}>} The cells updated
     */
    setValue(value, clear = true) {
        if (clear && this._formulaType)
            this.clear();
        if (value instanceof RichText)
            this._value = value.copy(this);
        else
            this._value = value;

        // trigger updates on other cells
        const calculations = this.workbook()._refTable.getCalculationOrder(this.getRef());
        calculations.forEach(cal => {
            const cell = this.workbook().sheet(cal.sheet).getCell(cal.row, cal.col);
            cell._value = cell._evaluateFormula();
        });
        return calculations;
    }

    getFormula() {
        if (this._formulaType === "shared" && !this._formula) {
            return this.sharedFormula();
        }
        return this._formula;
    }

    /**
     * Sets the formula of this cell.
     * @param {string} formula - The formula to set.
     * @returns {Array.<{}>} The cells updated
     */
    setFormula(formula) {
        // skip if the formula does not change
        if (formula === this._formula)
            return [];

        this.clear();
        this._formulaType = "normal";
        this._formula = formula;
        const result = this._evaluateFormula();

        // update reference table
        this._dep = this.workbook()._parser.parseDep(this);
        this._dep.forEach(refA => {
            this.workbook()._refTable.add(refA, this.getRef());
        });

        // trigger updates on other cells
        return this.setValue(result, false);
    }

    /**
     * Evaluate the cell's formula.
     * Note: This method does not save the formula to this cell.
     * @return {*} The result of the formula
     */
    _evaluateFormula() {
        let result;
        try {
            result = this.workbook()._parser.parse(this);
            if (typeof result === 'object')
                result = result.result;
        } catch (e) {
            console.error(e);
        }
        return result;
    }

    getRef() {
        return {
            sheet: this.sheet().name(),
            row: this.rowNumber(),
            col: this.columnNumber()
        };
    }

    /**
     * check if style exists, if not create one.
     * @private
     * @return {undefined}
     */
    _checkStyle() {
        if (!this._style && !(arguments[0] instanceof Style)) {
            this._style = this.workbook().styleSheet().createStyle(this._styleId);
        }
    }

    /**
     * Get cell's style.
     * @param {string} name - Style name
     * @return {*} Style
     */
    getStyle(name) {
        this._checkStyle();
        return this._style.getStyle(name);
    }

    /**
     * Set cell style
     * @param {string} name - Style name
     * @param {*} value - Style
     * @return {Cell} The cell
     */
    setStyle(name, value) {
        this._checkStyle();
        this._style.setStyle(name, value);
        return this;
    }

    /**
     * Gets a value indicating whether the cell is the active cell in the sheet.
     * @returns {boolean} True if active, false otherwise.
     */
    /**
     * Make the cell the active cell in the sheet.
     * @param {boolean} active - Must be set to `true`. Deactivating directly is not supported. To deactivate, you should activate a different cell instead.
     * @returns {Cell} The cell.
     */
    active() {
        return new ArgHandler('Cell.active', arguments)
            .case(() => {
                return this.sheet().activeCell() === this;
            })
            .case('boolean', active => {
                if (!active) throw new Error("Deactivating cell directly not supported. Activate a different cell instead.");
                this.sheet().activeCell(this);
                return this;
            })
            .handle();
    }

    /**
     * Get the address of the column.
     * @param {{}} [opts] - Options
     * @param {boolean} [opts.includeSheetName] - Include the sheet name in the address.
     * @param {boolean} [opts.rowAnchored] - Anchor the row.
     * @param {boolean} [opts.columnAnchored] - Anchor the column.
     * @param {boolean} [opts.anchored] - Anchor both the row and the column.
     * @returns {string} The address
     */
    address(opts) {
        return addressConverter.toAddress({
            type: 'cell',
            rowNumber: this.rowNumber(),
            columnNumber: this.columnNumber(),
            sheetName: opts && opts.includeSheetName && this.sheet().name(),
            rowAnchored: opts && (opts.rowAnchored || opts.anchored),
            columnAnchored: opts && (opts.columnAnchored || opts.anchored)
        });
    }

    /**
     * Gets the parent column of the cell.
     * @returns {Column} The parent column.
     */
    column() {
        return this.sheet().column(this.columnNumber());
    }

    /**
     * Clears the contents from the cell.
     * @returns {Cell} The cell.
     */
    clear() {
        // update reference table;
        if (this._formula) {
            this._dep.forEach(refA => {
                this.workbook()._refTable.remove(refA, this.getRef());
            });
        }

        const hostSharedFormulaId = this._formulaRef && this._sharedFormulaId;

        this._value = undefined;
        this._formulaType = undefined;
        this._formula = undefined;
        this._sharedFormulaId = undefined;
        this._formulaRef = undefined;

        // TODO in future version: Move shared formula to some other cell. This would require parsing the formula...
        if (!_.isNil(hostSharedFormulaId)) this.sheet().clearCellsUsingSharedFormula(hostSharedFormulaId);

        return this;
    }

    /**
     * Gets the column name of the cell.
     * @returns {number} The column name.
     */
    columnName() {
        return addressConverter.columnNumberToName(this.columnNumber());
    }

    /**
     * Gets the column number of the cell (1-based).
     * @returns {number} The column number.
     */
    columnNumber() {
        return this._columnNumber;
    }

    /**
     * Find the given pattern in the cell and optionally replace it.
     * @param {string|RegExp} pattern - The pattern to look for. Providing a string will result in a case-insensitive substring search. Use a RegExp for more sophisticated searches.
     * @param {string|function} [replacement] - The text to replace or a String.replace callback function. If pattern is a string, all occurrences of the pattern in the cell will be replaced.
     * @returns {boolean} A flag indicating if the pattern was found.
     */
    find(pattern, replacement) {
        pattern = regexify(pattern);

        const value = this.value();
        if (typeof value !== 'string') return false;

        if (_.isNil(replacement)) {
            return pattern.test(value);
        } else {
            const replaced = value.replace(pattern, replacement);
            if (replaced === value) return false;
            this.value(replaced);
            return true;
        }
    }

    /**
     * Gets the formula in the cell. Note that if a formula was set as part of a range, the getter will return 'SHARED'. This is a limitation that may be addressed in a future release.
     * @returns {string} The formula in the cell.
     *//**
     * Sets the formula in the cell. The previous formula result will be removed.
     * @param {string|undefined|null} formula - The formula to set.
     * @returns {Cell} The cell.
     */
    /**
     *  @param {string} formula - The formula to set.
     *  @param {string|number|boolean|Date} result - The formula result.
     * @return {Cell} The cell
     */
    formula() {
        return new ArgHandler('Cell.formula', arguments)
            .case(() => {
                return this.getFormula();
            })
            .case('nil', () => {
                this.clear();
                return this;
            })
            .case('string', formula => {
                this.setFormula(formula);
                return this;
            })
            .case(['string', '*'], (formula, result) => {
                const supportedType = ['string', 'number', 'boolean'];
                if (!supportedType.includes(typeof result) && !(result instanceof Date))
                    throw new Error('Formula result can only be string, number, boolean or Date.');
                this.clear();
                this._formulaType = "normal";
                this._formula = formula;
                this._value = result;
                return this;
            })
            .handle();
    }

    /**
     * Gets the hyperlink attached to the cell.
     * @returns {string|undefined} The hyperlink or undefined if not set.
     *//**
     * Set or clear the hyperlink on the cell.
     * @param {string|Cell|undefined|null} hyperlink - The hyperlink to set or undefined to clear.
     * @returns {Cell} The cell.
     *//**
     * Set the internal/external hyperlink.
     * @param {string|Cell|undefined} hyperlink - The hyperlink to set or undefined to clear.
     * @param {boolean} internal - Is internal hyperlink.
     * @returns {Cell} The cell.
     */
    /**
     * Set the hyperlink options on the cell.
     * @param {{}|Cell} opts - Options or Cell. If opts is a Cell then an internal hyperlink is added.
     * @param {string|Cell} [opts.hyperlink] - The hyperlink to set, can be a Cell or an internal/external string.
     * @param {string} [opts.tooltip] - Additional text to help the user understand more about the hyperlink.
     * @param {string} [opts.email] - Email address, ignored if opts.hyperlink is set.
     * @param {string} [opts.emailSubject] - Email subject, ignored if opts.hyperlink is set.
     * @returns {Cell} The cell.
     */
    hyperlink() {
        return new ArgHandler('Cell.hyperlink', arguments)
            .case(() => {
                return this.sheet().hyperlink(this.address());
            })
            .case('nil', () => {
                this.sheet().hyperlink(this.address(), null);
                return this;
            })
            .case('string', hyperlink => {
                this.sheet().hyperlink(this.address(), hyperlink);
                return this;
            })
            .case(['string', 'boolean'], (hyperlink, internal) => {
                this.sheet().hyperlink(this.address(), hyperlink, internal);
                return this;
            })
            .case(['object'], opts => {
                this.sheet().hyperlink(this.address(), opts);
                return this;
            })
            .handle();
    }


    /**
     * Gets the data validation object attached to the cell.
     * @returns {object|undefined} The data validation or undefined if not set.
     */
    /**
     * Set or clear the data validation object of the cell.
     * @param {object|undefined} dataValidation - Object or null to clear.
     * @returns {Cell} The cell.
     */
    dataValidation() {
        return new ArgHandler('Cell.dataValidation', arguments)
            .case(() => {
                return this.sheet().dataValidation(this.address());
            })
            .case('boolean', obj => {
                return this.sheet().dataValidation(this.address(), obj);
            })
            .case('*', obj => {
                this.sheet().dataValidation(this.address(), obj);
                return this;
            })
            .handle();
    }

    /**
     * Gets a value indicating whether the cells in the range are merged.
     * @returns {boolean|Reference} If it is merged, return a reference indicating where it merges, otherwise
     *                              return false.
     */
    merged() {
        return this.sheet().merged(this.address());
    }

    /**
     * Gets if this cell is the master merged cell.
     * Node: If this cell is not part of a merged cell, return false.
     * @param {{}} [merged] - Result from Cell.merged(), supply this can speed up performance.
     * @return {boolean} - If the cell is master merged cell.
     */
    isMasterMergedCell(merged) {
        if (!merged) merged = this.merged();
        if (!merged) return false;
        return merged.from.row === this.rowNumber() && merged.from.col === this.columnNumber();
    }

    /**
     * Gets the master merged cell.
     * Note: Return undefined if this cell is not part of the merged cell.
     * @param {{}} [merged] - Result from Cell.merged(), supply this can speed up performance.
     * @return {Cell|undefined} - The master merged cell.
     */
    masterMergedCell(merged) {
        if (!merged) merged = this.merged();
        if (!merged) return;
        return this.sheet().getCell(merged.from.row, merged.from.col);
    }

    /**
     * Callback used by tap.
     * @callback Cell~tapCallback
     * @param {Cell} cell - The cell
     * @returns {undefined}
     */
    /**
     * Invoke a callback on the cell and return the cell. Useful for method chaining.
     * @param {Cell~tapCallback} callback - The callback function.
     * @returns {Cell} The cell.
     */
    tap(callback) {
        callback(this);
        return this;
    }

    /**
     * Callback used by thru.
     * @callback Cell~thruCallback
     * @param {Cell} cell - The cell
     * @returns {*} The value to return from thru.
     */
    /**
     * Invoke a callback on the cell and return the value provided by the callback. Useful for method chaining.
     * @param {Cell~thruCallback} callback - The callback function.
     * @returns {*} The return value of the callback.
     */
    thru(callback) {
        return callback(this);
    }

    /**
     * Create a range from this cell and another.
     * @param {Cell|string} cell - The other cell or cell address to range to.
     * @returns {Range} The range.
     */
    rangeTo(cell) {
        return this.sheet().range(this, cell);
    }

    /**
     * Returns a cell with a relative position given the offsets provided.
     * @param {number} rowOffset - The row offset (0 for the current row).
     * @param {number} columnOffset - The column offset (0 for the current column).
     * @returns {Cell} The relative cell.
     */
    relativeCell(rowOffset, columnOffset) {
        const row = rowOffset + this.rowNumber();
        const column = columnOffset + this.columnNumber();
        return this.sheet().cell(row, column);
    }

    /**
     * Gets the parent row of the cell.
     * @returns {Row} The parent row.
     */
    row() {
        return this._row;
    }

    /**
     * Gets the row number of the cell (1-based).
     * @returns {number} The row number.
     */
    rowNumber() {
        return this.row().rowNumber();
    }

    /**
     * Gets the parent sheet.
     * @returns {Sheet} The parent sheet.
     */
    sheet() {
        return this.row().sheet();
    }

    /**
     * Gets an individual style.
     * @param {string} name - The name of the style.
     * @returns {*} The style.
     *//**
     * Gets multiple styles.
     * @param {Array.<string>} names - The names of the style.
     * @returns {object.<string, *>} Object whose keys are the style names and values are the styles.
     *//**
     * Sets an individual style.
     * @param {string} name - The name of the style.
     * @param {*} value - The value to set.
     * @returns {Cell} The cell.
     *//**
     * Sets the styles in the range starting with the cell.
     * @param {string} name - The name of the style.
     * @param {Array.<Array.<*>>} - 2D array of values to set.
     * @returns {Range} The range that was set.
     *//**
     * Sets multiple styles.
     * @param {object.<string, *>} styles - Object whose keys are the style names and values are the styles to set.
     * @returns {Cell} The cell.
     */
    /**
     * Sets to a specific style
     * @param {Style} style - Style object given from stylesheet.createStyle
     * @returns {Cell} The cell.
     */
    style() {
        this._checkStyle();
        return new ArgHandler("Cell.style", arguments)
            .case('string', name => {
                // Get single value
                return this._style.style(name);
            })
            .case('array', names => {
                // Get list of values
                const values = {};
                names.forEach(name => {
                    values[name] = this.style(name);
                });

                return values;
            })
            .case(["string", "array"], (name, values) => {
                const numRows = values.length;
                const numCols = values[0].length;
                const range = this.rangeTo(this.relativeCell(numRows - 1, numCols - 1));
                return range.style(name, values);
            })
            .case(['string', '*'], (name, value) => {
                // Set a single value for all cells to a single value
                this._style.style(name, value);
                return this;
            })
            .case('object', nameValues => {
                // Object of key value pairs to set
                for (const name in nameValues) {
                    if (!nameValues.hasOwnProperty(name)) continue;
                    const value = nameValues[name];
                    this.style(name, value);
                }

                return this;
            })
            .case('Style', style => {
                this._style = style;
                this._styleId = style.id();

                return this;
            })
            .handle();
    }

    /**
     * Gets the value of the cell.
     * @returns {string|boolean|number|Date|RichText|undefined} The value of the cell.
     *//**
     * Sets the value of the cell.
     * @param {string|boolean|number|null|undefined|RichText} value - The value to set.
     * @returns {Cell} The cell.
     */
    /**
     * Sets the values in the range starting with the cell.
     * @param {Array.<Array.<string|boolean|number|null|undefined>>} - 2D array of values to set.
     * @returns {Range} The range that was set.
     */
    value() {
        return new ArgHandler('Cell.value', arguments)
            .case(() => {
                return this.getValue();
            })
            .case("array", values => {
                const numRows = values.length;
                const numCols = values[0].length;
                const range = this.rangeTo(this.relativeCell(numRows - 1, numCols - 1));
                return range.value(values);
            })
            .case('*', value => {
                this.setValue(value);
                return this;
            })
            .handle();
    }

    /**
     * Gets the parent workbook.
     * @returns {Workbook} The parent workbook.
     */
    workbook() {
        return this.row().workbook();
    }

    /**
     * Append horizontal page break after the cell.
     * @returns {Cell} the cell.
     */
    addHorizontalPageBreak() {
        this.row().addPageBreak();
        return this;
    }

    /**
     * Gets the translated shared formula.
     * Based on the approach from {@link https://github.com/dtjohnson/xlsx-populate/issues/129} but more efficient.
     * @return {string} Translated shared formula
     */
    sharedFormula() {
        if (this._sharedFormulaId == null)
            throw Error(`This cell ${this.address()} does not contain shared formulas.`);
        const refCell = this.sheet().getSharedFormulaRefCell(this._sharedFormulaId);
        const refCol = refCell.columnNumber();
        const refRow = refCell.rowNumber();
        const cellCol = this.columnNumber();
        const cellRow = this.rowNumber();

        const offsetCol = cellCol - refCol;
        const offsetRow = cellRow - refRow;

        const formula = refCell._formula
            .replace(/(\$)?([A-Za-z]+)(\$)?([0-9]+)(\()?/g, (match, absCol, colName, absRow, row, isFunction, index) => {
                if (isFunction) {
                    return match;
                }

                const col = +addressConverter.columnNameToNumber(colName);
                row = +row;

                const _col = absCol ? col : col + offsetCol;
                const _row = absRow ? row : row + offsetRow;

                const _colName = addressConverter.columnNumberToName(_col);
                return `${_colName}${_row}`;
            });
        return formula;
    }

    dependencies() {
        return this._dep;
    }

    /* INTERNAL */

    /**
     * Gets the formula if a shared formula ref cell.
     * @returns {string|undefined} The formula.
     * @ignore
     */
    getSharedRefFormula() {
        return this._formulaType === "shared" ? this._formulaRef && this._formula : undefined;
    }

    /**
     * Check if this cell uses a given shared a formula ID.
     * @param {number} id - The shared formula ID.
     * @returns {boolean} A flag indicating if shared.
     * @ignore
     */
    sharesFormula(id) {
        return this._formulaType === "shared" && this._sharedFormulaId === id;
    }

    /**
     * Set a shared formula on the cell.
     * @param {number} id - The shared formula index.
     * @param {string} [formula] - The formula (if the reference cell).
     * @param {string} [sharedRef] - The address of the shared range (if the reference cell).
     * @returns {undefined}
     * @ignore
     */
    setSharedFormula(id, formula, sharedRef) {
        this.clear();

        this._formulaType = "shared";
        this._sharedFormulaId = id;
        this._formula = formula;
        this._formulaRef = sharedRef;
    }

    /**
     * Convert the cell to an XML object.
     * @returns {{}} The XML form.
     * @ignore
     */
    toXml() {
        // Create a node.
        const node = {
            name: 'c',
            attributes: this._remainingAttributes || {}, // Start with any remaining attributes we don't current handle.
            children: []
        };

        // Set the address.
        node.attributes.r = this.address();

        if (!_.isNil(this._formulaType)) {
            // Add the formula.
            const fNode = {
                name: 'f',
                attributes: this._remainingFormulaAttributes || {}
            };

            if (this._formulaType === 'normal') {
                if (this._formula != null) fNode.children = [this._formula];
            } else {
                if (this._formulaType != null)
                    fNode.attributes.t = this._formulaType;
                if (this._sharedFormulaId != null)
                    fNode.attributes.si = this._sharedFormulaId;

                // main shared formula
                if (this._formulaRef != null) {
                    fNode.attributes.ref = this._formulaRef;
                    fNode.children = [this._formula];
                }
            }

            node.children.push(fNode);

            // save formula value
            if (!_.isNil(this._value)) {
                // the type attribute can be empty (means number or date), 'str' or 'b'
                let type, text;
                if (typeof this._value === "string") {
                    type = "str";
                    text = this._value;
                } else if (typeof this._value === "boolean") {
                    type = "b";
                    text = this._value ? 1 : 0;
                } else if (typeof this._value === "number") {
                    text = this._value;
                } else if (this._value instanceof Date) {
                    text = dateConverter.dateToNumber(this._value);
                }
                if (type) node.attributes.t = type;
                node.children.push({ name: 'v', children: [text] });
            }
        } else if (!_.isNil(this._value)) {
            // Add the value. Don't emit value if a formula is set as Excel will show this stale value.
            let type, text;
            if (typeof this._value === "string") {
                type = "s";
                text = this.workbook().sharedStrings().getIndexForString(this._value);
            } else if (typeof this._value === "boolean") {
                type = "b";
                text = this._value ? 1 : 0;
            } else if (typeof this._value === "number") {
                text = this._value;
            } else if (this._value instanceof Date) {
                text = dateConverter.dateToNumber(this._value);
            } else if (this._value instanceof RichText) {
                type = "s";
                text = this.workbook().sharedStrings().getIndexForString(this._value.toXml());
            }

            if (type) node.attributes.t = type;
            const vNode = { name: 'v', children: [text] };
            node.children.push(vNode);
        }

        // If the style is set, set the style ID.
        if (!_.isNil(this._style)) {
            node.attributes.s = this._style.id();
        } else if (!_.isNil(this._styleId)) {
            node.attributes.s = this._styleId;
        }

        // Add any remaining children that we don't currently handle.
        if (this._remainingChildren) {
            node.children = node.children.concat(this._remainingChildren);
        }

        return node;
    }

    /* PRIVATE */

    /**
     * Initialize the cell node.
     * @param {{}|number} nodeOrColumnNumber - The existing node or the column number of a new cell.
     * @param {number} [styleId] - The style ID for the new cell.
     * @returns {undefined}
     * @private
     */
    _init(nodeOrColumnNumber, styleId) {
        if (_.isObject(nodeOrColumnNumber)) {
            // Parse the existing node.
            this._parseNode(nodeOrColumnNumber);
        } else {
            // This is a new cell.
            this._columnNumber = nodeOrColumnNumber;
            if (!_.isNil(styleId)) this._styleId = styleId;
        }
    }

    /**
     * Parse the existing node.
     * @param {{}} node - The existing node.
     * @returns {undefined}
     * @private
     */
    _parseNode(node) {
        // Parse the column numbr out of the address.
        const ref = addressConverter.fromAddress(node.attributes.r);
        this._columnNumber = ref.columnNumber;

        // Store the style ID if present.
        if (!_.isNil(node.attributes.s)) this._styleId = node.attributes.s;

        // Parse the formula if present..
        const fNode = xmlq.findChild(node, 'f');
        if (fNode) {
            this._formulaType = fNode.attributes.t || "normal"; // or "shared"
            this._formulaRef = fNode.attributes.ref;
            this._sharedFormulaId = fNode.attributes.si;
            this._formula = fNode.children.length ? `${fNode.children[0]}` : undefined;

            // this is the main cell the shared formula references.
            if (this._formulaType === "shared" && this._formulaRef) {
                // store the shared formula to sheet
                this.sheet().setSharedFormulaRefCell(this._sharedFormulaId, this);

                // Update the sheet's max shared formula ID so we can set future IDs an index beyond this.
                this.sheet().updateMaxSharedFormulaId(this._sharedFormulaId);
            } else if (this._formulaType === 'shared' && !this._formula) {
                // evaluate shared formula
                this._formula = this.sharedFormula();
            }

            // parse dependencies
            this._dep = this.workbook()._parser.parseDep(this);
            this._dep.forEach(refA => {
                this.workbook()._refTable.add(refA, this.getRef());
            });

            // Delete the known attributes.
            fNode.attributes.t = undefined;
            fNode.attributes.ref = undefined;
            fNode.attributes.si = undefined;

            // If any unknown attributes are still present, store them for later output.
            this._remainingFormulaAttributes = xmlq.filterEmptyAttribute(fNode);
        }

        // Parse the value.
        const type = node.attributes.t;
        if (type === "s") {
            // String value.
            const vNode = xmlq.findChild(node, 'v');
            if (vNode) {
                const sharedIndex = vNode.children[0];
                this._value = this.workbook().sharedStrings().getStringByIndex(sharedIndex);

                // rich text
                if (_.isArray(this._value)) {
                    this._value = new RichText(this._value);
                }
            } else {
                this._value = '';
            }
        } else if (type === "str") {
            // Simple string value.
            const vNode = xmlq.findChild(node, 'v');
            this._value = vNode && vNode.children[0];
        } else if (type === "inlineStr") {
            // Inline string value: can be simple text or rich text.
            const isNode = xmlq.findChild(node, 'is');
            if (isNode.children[0].name === "t") {
                const tNode = isNode.children[0];
                this._value = tNode.children[0];
            } else {
                this._value = isNode.children;
            }
        } else if (type === "b") {
            // Boolean value.
            this._value = xmlq.findChild(node, 'v').children[0] === 1;
        } else if (type === "e") {
            // Error value.
            const error = xmlq.findChild(node, 'v').children[0];
            this._value = FormulaError.getError(error);
        } else {
            // Number value.
            const vNode = xmlq.findChild(node, 'v');
            this._value = vNode && Number(vNode.children[0]);
        }

        // Delete known attributes.
        node.attributes.r = undefined;
        node.attributes.s = undefined;
        node.attributes.t = undefined;

        // If any unknown attributes are still present, store them for later output.
        this._remainingAttributes = xmlq.filterEmptyAttribute(node);

        // Delete known children.
        xmlq.removeChild(node, 'f');
        xmlq.removeChild(node, 'v');
        xmlq.removeChild(node, 'is');

        // If any unknown children are still present, store them for later output.
        if (node.children.length !== 0) this._remainingChildren = node.children;
    }
}

module.exports = Cell;

/*
<c r="A6" s="1" t="s">
    <v>2</v>
</c>
*/

