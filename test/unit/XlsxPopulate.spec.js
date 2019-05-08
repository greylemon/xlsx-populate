"use strict";

const XlsxPopulate = require("../../lib/XlsxPopulate");

const expect = require('chai').expect;

describe("XlsxPopulate", () => {
    describe("dateToNumber", () => {
        it("should call dateConverter.dateToNumber", () => {
            expect(XlsxPopulate.dateToNumber).to.be.a('function');
        });
    });

    describe("fromBlankAsync", () => {
        it("should call Workbook.fromBlankAsync", () => {
            expect(XlsxPopulate.fromBlankAsync).to.be.a('function');
        });
    });

    describe("fromDataAsync", () => {
        it("should call Workbook.fromDataAsync", () => {
            expect(XlsxPopulate.fromDataAsync).to.be.a('function');
        });
    });

    describe("fromFileAsync", () => {
        it("should call Workbook.fromFileAsync", () => {
            expect(XlsxPopulate.fromFileAsync).to.be.a('function');
        });
    });

    describe("numberToDate", () => {
        it("should call dateConverter.numberToDate", () => {
            expect(XlsxPopulate.numberToDate).to.be.a('function');
        });
    });

    describe("statics", () => {
        it("should set the statics", () => {
            expect(XlsxPopulate.MIME_TYPE).to.be.a('string');
            expect(XlsxPopulate.FormulaError).to.eq(require('../../lib/FormulaError'));
        });
    });
});
