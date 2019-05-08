"use strict";

const FormulaError = require("../../lib/FormulaError");
const expect = require('chai').expect;

describe("FormulaError", () => {

    describe("error", () => {
        it("should return the error", () => {
            const formulaError = new FormulaError("foo");
            expect(formulaError.error()).to.eq("foo");
        });
    });

    describe("static", () => {
        it("should create the static instances", () => {
            expect(FormulaError.DIV0.error()).to.eq("#DIV/0!");
            expect(FormulaError.NA.error()).to.eq("#N/A");
            expect(FormulaError.NAME.error()).to.eq("#NAME?");
            expect(FormulaError.NULL.error()).to.eq("#NULL!");
            expect(FormulaError.NUM.error()).to.eq("#NUM!");
            expect(FormulaError.REF.error()).to.eq("#REF!");
            expect(FormulaError.VALUE.error()).to.eq("#VALUE!");
        });
    });

    describe("getError", () => {
        it("should get the matching error", () => {
            expect(FormulaError.getError("#VALUE!")).to.eq(FormulaError.VALUE);
        });

        it("should create a new instance for unknown errors", () => {
            expect(FormulaError.getError("foo").error()).to.eq("foo");
        });
    });
});
