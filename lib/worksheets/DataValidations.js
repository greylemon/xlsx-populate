"use strict";

const {
    decodeRowRange, decodeColRange, encodeColRange, encodeRowRange, isInColRange, isInRowRange
} = require('../formula/Utils');

const xm = 'http://schemas.microsoft.com/office/excel/2006/main';

// https://docs.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/89029dfc-1ca8-4ff9-afe0-46f9454d09c6
const defaultAttributes = {
    type: 'none',
    errorStyle: 'stop',
    imeMode: 'noControl',
    operator: 'between',
    allowBlank: false,
    showDropDown: false, // show dropdown combo box. false if show in-cell dropdown
    showInputMessage: false,
    showErrorMessage: false,
    errorTitle: undefined,
    error: undefined,
    promptTitle: undefined,
    prompt: undefined
};
const keys = Object.keys(defaultAttributes);
const booleanKeys = ['allowBlank', 'showDropDown', 'showInputMessage', 'showErrorMessage'];

/**
 * A class stores dataValidations.
 */
class DataValidations {
    /**
     * @param {Sheet} sheet - The sheet.
     * @param {{}} dataValidationsNode - The dataValidation node.
     * @param {{}} extNode - The extension data validation node.
     */
    constructor(sheet, dataValidationsNode, extNode) {
        this._sheet = sheet;

        // helps faster query.
        this._data = new Map();
        this._position = { sheet: sheet.name() };
        this._addressCache = {};

        if (dataValidationsNode) {
            dataValidationsNode.children.forEach(dataValidationNode => {
                this.set(dataValidationNode.attributes.sqref, dataValidationNode);
            });
        }
        if (extNode && extNode.children.length > 0) {
            extNode.children[0].children.forEach(x14DataValidationNode => {
                // xm:sqref
                this.set(x14DataValidationNode.children[x14DataValidationNode.children.length - 1].children[0],
                    x14DataValidationNode);
            });
        }
    }

    /**
     * Get a data validation on the given address.
     * @param {string} address - The address of a cell or range.
     * @param {boolean} [parseFormula=true] - Whether parsing the formulas. If true, the returned object contains
     *                                  formula1Result amd formula2Result.
     * @return {undefined|{allowBlank, formula1, showErrorMessage, imeMode, showDropDown, formula2, showInputMessage,
     *          type, error, operator, errorStyle, errorTitle, promptTitle, prompt, formula1Result, formula2Result}}
     *          - The data validation object.
     */
    getDataValidation(address, parseFormula = true) {
        const node = this.get(address);
        if (!node) return;

        const dataValidation = {
            formula1: undefined,
            formula2: undefined
        };
        keys.forEach(key => {
            dataValidation[key] = node.attributes[key] || defaultAttributes[key];

            // cast to boolean
            if (booleanKeys.includes(key))
                dataValidation[key] = Boolean(dataValidation[key]);
        });

        const isX14 = node.name.charAt(0) === 'x';
        const hasFormula1 = isX14 ? node.children[0].name === 'x14:formula1' : node.children.length > 0;
        const hasFormula2 = isX14 ? node.children[1].name === 'x14:formula2' : node.children.length > 1;
        if (hasFormula1) {
            dataValidation.formula1 = String(isX14 ? node.children[0].children[0].children[0] : node.children[0].children[0]);
            if (parseFormula) dataValidation.formula1Result = this.parseFormula(dataValidation.formula1);
        }
        if (hasFormula2) {
            dataValidation.formula2 = String(isX14 ? node.children[1].children[0].children[0] : node.children[1].children[0]);
            if (parseFormula) dataValidation.formula2Result = this.parseFormula(dataValidation.formula2);
        }

        // pre-process the list data validation
        if (dataValidation.type === 'list') {
            if (typeof dataValidation.formula1Result === "string") {
                dataValidation.formula1Result = dataValidation.formula1Result.replace(/\s[,;]\s/g, ',').split(',');
                for (let i = 0; i < dataValidation.formula1Result.length; i++) {
                    dataValidation.formula1Result[i] = dataValidation.formula1Result[i].trim();
                }
            } else if (Array.isArray(dataValidation.formula1Result)) {
                dataValidation.formula1Result = dataValidation.formula1Result.flat();
            } else {
                dataValidation.formula1Result = [dataValidation.formula1Result];
            }
        }
        return dataValidation;
    }

    /**
     * Set a data validation on the given address.
     * @param {string} address - The address of a cell or range.
     * @param {undefined|{allowBlank, formula1, showErrorMessage, imeMode, showDropDown, formula2, showInputMessage,
     *          type, error, operator, errorStyle, errorTitle, promptTitle, prompt}} dataValidation - The data validation object.
     * @return {undefined}
     */
    setDataValidation(address, dataValidation) {
        const formula1 = String(dataValidation.formula1);
        const formula2 = String(dataValidation.formula2);

        // if one of the formula contains sheet reference.
        const isX14 = formula1.formula1.replace(/"[^"]*"/g, '').indexOf('!') > 0
            || formula2.replace(/"[^"]*"/g, '').indexOf('!') > 0;

        const node = {
            name: 'dataValidation',
            attributes: {},
            children: []
        };
        keys.forEach(key => {
            if (dataValidation[key] && dataValidation[key] !== defaultAttributes[key])
                node.attributes[key] = dataValidation[key];
        });

        if (isX14) {
            node.name = 'x14:dataValidation';
            if (dataValidation.formula1 != null) {
                node.children.push({
                    name: 'x14:formula1',
                    attributes: {},
                    children: [
                        {
                            name: 'xm:f',
                            attributes: {},
                            children: [dataValidation.formula1]
                        }
                    ]
                });
                if (dataValidation.formula2) {
                    node.children.push({
                        name: 'x14:formula2',
                        attributes: {},
                        children: [
                            {
                                name: 'xm:f',
                                attributes: {},
                                children: [dataValidation.formula2]
                            }
                        ]
                    });
                }
            }
            node.children.push({
                name: 'xm:sqref',
                attributes: {},
                children: [address]
            });
        } else {
            // normal data validation (no other sheet reference)
            node.attributes.sqref = address;
            if (dataValidation.formula1 != null) {
                node.children.push({
                    name: 'formula1',
                    attributes: {},
                    children: [dataValidation.formula1]
                });
                if (dataValidation.formula2 != null) {
                    node.children.push({
                        name: 'formula1',
                        attributes: {},
                        children: [dataValidation.formula1]
                    });
                }
            }
        }
    }

    /**
     * Validate a input.
     * @param {Cell} cell - The cell to validate.
     * @param {string|boolean|number|undefined|null} input - The input to validate.
     * @param {boolean} [isFormula] - If true, the input is a formula.
     * @return {{result: boolean, dataValidation: {}}} An object that contains Whether the input is valid.
     * @example
     *      validate(cell, '1+1', true)
     *      validate(cell, '111')
     */
    validate(cell, input, isFormula) {
        const dataValidation = this.getDataValidation(cell.address());
        const res = { result: true, dataValidation };
        if (!dataValidation)
            return res;
        if (dataValidation.allowBlank && input == null)
            return res;
        if (isFormula)
            input = this._sheet.workbook().formulaParser._parser.parse(`${input}`, cell.getRef());
        if (Array.isArray(input))
            input = input.flat()[0];
        if (dataValidation.type === 'whole') { // whole number
            res.result = typeof input === "number" && Math.trunc(input) === input
                && this._testNumber(input, dataValidation);
        } else if (dataValidation.type === 'decimal') {
            res.result = typeof input === "number" && this._testNumber(input, dataValidation);
        } else if (dataValidation.type === 'list') {
            res.result = dataValidation.formula1Result.includes(input);
        } else if (dataValidation.type === 'date') {
            res.result = typeof input === "number" && this._testNumber(Math.trunc(input), dataValidation);
        } else if (dataValidation.type === 'time') {
            res.result = typeof input === "number" && this._testNumber(input % 1, dataValidation);
        } else if (dataValidation.type === 'textLength') {
            res.result = this._testNumber(String(input).length, dataValidation);
        } else if (dataValidation.type === 'custom') {
            res.result = Boolean(dataValidation.formula1Result);
        }
        return res;
    }

    _testNumber(number, dataValidation) {
        const { operator, formula1Result, formula2Result } = dataValidation;
        switch (operator) {
            case 'between':
                return number >= formula1Result && number <= formula2Result;
            case 'notBetween':
                return number < formula1Result && number > formula2Result;
            case 'equal':
                return number === formula1Result;
            case 'notEqual':
                return number !== formula1Result;
            case 'greaterThan':
                return number > formula1Result;
            case 'lessThan':
                return number < formula1Result;
            case 'greaterThanOrEqual':
                return number >= formula1Result;
            case 'lessThanOrEqual':
                return number <= formula1Result;
            default:
                throw Error(`DataValidations._testNumber: Unknown operator ${operator} in ${JSON.stringify(dataValidation)}`);
        }
    }

    /**
     * Parse a formula.
     * @param {string|number} formula - The formula to parse.
     * @return {*} The formula result
     */
    parseFormula(formula) {
        return this._sheet.workbook().formulaParser._parser.parse(`${formula}`,
            { sheet: this._sheet.name() }, true);
    }

    /**
     * Parse a reference, i.e. D19, A1:C9, Sheet2!A1, A:C, 1:6...
     * @param {string} address - The reference string to parse.
     * @return {{sheet: string, row: number, col: number}} The reference object.
     */
    parse(address) {
        let ref = this._addressCache[address];
        if (!ref) {
            ref = this._sheet.workbook()._parser.parseReference(address, this._position.sheet);
            this._addressCache[address] = ref;
        }
        return ref;
    }

    /**
     * Retrieve a dataValidation node.
     * @param {string} refString - The reference string to parse.
     * @return {{}|undefined} A dataValidation node.
     */
    get(refString) {
        if (this._data.size === 0) return;
        const ref = this.parse(refString);
        let curr;
        this._data.forEach((refRow, rowRange) => {
            const row = decodeRowRange(rowRange);
            if (!(row.from <= ref.row && ref.row <= row.to))
                return;
            refRow.forEach((dataValidation, colRange) => {
                const col = decodeColRange(colRange);
                if (col.from <= ref.col && ref.col <= col.to) {
                    // prefer range reference
                    if (row.from !== row.to || col.from !== col.to || !curr)
                        curr = dataValidation;
                }
            });
        });
        return curr;
    }

    /**
     * Add dataValidation. If exists, replace it.
     * @param {string} refString - The reference string to add.
     * @param {{name, attributes, children}} dataValidation - The inner dataValidation node.
     * @return {undefined}
     */
    set(refString, dataValidation) {
        const addresses = refString.trim().split(' ');
        if (addresses.length > 1) {
            for (let i = 0; i < addresses.length; i++) {
                this.set(addresses[i], dataValidation);
            }
        } else {
            const ref = this.parse(refString);

            let rowRange, colRange;
            if (ref.from) {
                // ref is a range reference
                rowRange = encodeRowRange(ref.from.row, ref.to.row);
                colRange = encodeColRange(ref.from.col, ref.to.col);
            } else {
                // cell reference
                rowRange = encodeRowRange(ref.row, ref.row);
                colRange = encodeColRange(ref.col, ref.col);
            }

            // get row range
            let refRow = this._data.get(rowRange);
            if (!refRow) {
                refRow = new Map();
                this._data.set(rowRange, refRow);
            }
            refRow.set(colRange, dataValidation);
        }
    }

    /**
     * Remove a dataValidation.
     * @param {string} refString - The reference you want to remove.
     * @return {undefined}
     */
    delete(refString) {
        const ref = this.parse(refString);
        this._data.forEach((refRow, rowRange) => {
            if (!isInRowRange(rowRange, ref.row))
                return;
            refRow.forEach((dataValidation, colRange) => {
                if (!isInColRange(colRange, ref.col))
                    return;
                rowRange.delete(colRange);
            });
            if (refRow.size === 0)
                this._data.delete(rowRange);
        });
    }

    forEach(fn) {
        this._data.forEach(refRow => {
            refRow.forEach(node => {
                fn(node);
            });
        });
    }

    /**
     * Return an xml form object.
     * @return {{node, x14Node}} two Xml object.
     */
    toXml() {
        const node = {
            name: 'dataValidations',
            attributes: {},
            children: []
        };
        const x14Node = {
            name: 'x14:dataValidations',
            attributes: { 'xmlns:xm': xm },
            children: []
        };
        this._data.forEach(refRow => {
            refRow.forEach(dataValidation => {
                if (dataValidation != null) {
                    if (dataValidation.name.charAt(0) === 'x') {
                        x14Node.children.push(dataValidation);
                    } else {
                        node.children.push(dataValidation);
                    }
                }
            });
        });
        node.children = [...new Set(node.children)];
        node.attributes.count = node.children.length;
        x14Node.children = [...new Set(x14Node.children)];
        x14Node.attributes.count = x14Node.children.length;
        return { node, x14Node };
    }
}

module.exports = DataValidations;
