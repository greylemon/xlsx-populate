"use strict";

const { replaceRowNumber } = require("../../../lib/formula/Replacer");
const expect = require('chai').expect;

describe("Replacer", () => {
    describe("Replacer.replaceRowNumber", () => {
        it("should replace cell reference", done => {
            const res = replaceRowNumber('Fin_CMHP1!F29', 'Fin_CMHP1', 'Fin_CMHP1', 29, 30);
            expect(res).to.eq('Fin_CMHP1!F30');
            done();
        });

        it("should replace range reference", done => {
            const res = replaceRowNumber('Fin_CMHP1!F29:H34', 'Fin_CMHP1', 'Fin_CMHP1', 30, 29);
            expect(res).to.eq('Fin_CMHP1!F29:H33');
            done();
        });
    });
});
