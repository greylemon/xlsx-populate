"use strict";

const Workbook = require("./workbooks/Workbook");
const FormulaError = require("./FormulaError");
const dateConverter = require("./dateConverter");
const RichText = require("./worksheets/RichText");
const FormulaParser = require("fast-formula-parser");
const Range = require("./worksheets/Range");

/**
 * xlsx-poulate namespace.
 * @namespace
 */
class XlsxPopulate {
    /**
     * Convert a date to a number for Excel.
     * @param {Date} date - The date.
     * @returns {number} The number.
     */
    static dateToNumber(date) {
        return dateConverter.dateToNumber(date);
    }

    /**
     * Create a new blank workbook.
     * @returns {Promise.<Workbook>} The workbook.
     */
    static fromBlankAsync() {
        return Workbook.fromBlankAsync();
    }

    /**
     * Loads a workbook from a data object. (Supports any supported [JSZip data types]{@link https://stuk.github.io/jszip/documentation/api_jszip/load_async.html}.)
     * @param {string|Array.<number>|ArrayBuffer|Uint8Array|Buffer|Blob|Promise.<*>} data - The data to load.
     * @param {{}} [opts] - Options
     * @param {string} [opts.password] - The password to decrypt the workbook.
     * @returns {Promise.<Workbook>} The workbook.
     */
    static fromDataAsync(data, opts) {
        return Workbook.fromDataAsync(data, opts);
    }

    /**
     * Loads a workbook from file.
     * @param {string} path - The path to the workbook.
     * @param {{}} [opts] - Options
     * @param {string} [opts.password] - The password to decrypt the workbook.
     * @returns {Promise.<Workbook>} The workbook.
     */
    static fromFileAsync(path, opts) {
        return Workbook.fromFileAsync(path, opts);
    }

    /**
     * Convert an Excel number to a date.
     * @param {number} number - The number.
     * @returns {Date} The date.
     */
    static numberToDate(number) {
        return dateConverter.numberToDate(number);
    }
}

/**
 * The XLSX mime type.
 * @type {string}
 */
XlsxPopulate.MIME_TYPE = Workbook.MIME_TYPE;

/**
 * Formula error class.
 * @type {FormulaError}
 */
XlsxPopulate.FormulaError = FormulaError;

/**
 * RichTexts class
 * @type {RichText}
 */
XlsxPopulate.RichText = RichText;

XlsxPopulate.Range = Range;

/**
 * A formula parser library.
 * https://github.com/LesterLyu/fast-formula-parser
 */
XlsxPopulate.FormulaParser = FormulaParser;

module.exports = XlsxPopulate;

/**
 * @typedef {Object} ReferenceLiteral
 * @property {string} sheet - The sheet name.
 * @property {number} [row] - The row number. (1-based indexing)
 * @property {number} [col] - The column number. (1-based indexing)
 * @property {{row: number, col: number}} [from] - Range from.
 * @property {{row: number, col: number}} [to] - Range to.
 */
