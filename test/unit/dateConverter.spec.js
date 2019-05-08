"use strict";

const dateConverter = require("../../lib/dateConverter");
const expect = require('chai').expect;

describe("dateConverter", () => {
    describe("dateToNumber", () => {
        it("should convert date to number", () => {
            expect(dateConverter.dateToNumber(new Date('01 Jan 1900 00:00:00'))).to.eq(1);
            expect(dateConverter.dateToNumber(new Date('28 Feb 1900 00:00:00'))).to.eq(59);
            expect(dateConverter.dateToNumber(new Date('01 Mar 1900 00:00:00'))).to.eq(61);
            expect(dateConverter.dateToNumber(new Date('07 Mar 2015 13:26:24'))).to.eq(42070.56);
            expect(dateConverter.dateToNumber(new Date('04 Apr 2017 20:00:00'))).to.closeTo(42829.8333333333, 0.0000000001);
        });
    });

    describe("numberToDate", () => {
        it("should convert number to date", () => {
            expect(dateConverter.numberToDate(1)).to.deep.eq(new Date('01 Jan 1900 00:00:00'));
            expect(dateConverter.numberToDate(59)).to.deep.eq(new Date('28 Feb 1900 00:00:00'));
            expect(dateConverter.numberToDate(61)).to.deep.eq(new Date('01 Mar 1900 00:00:00'));
            expect(dateConverter.numberToDate(42070.56)).to.deep.eq(new Date('07 Mar 2015 13:26:24'));
            expect(dateConverter.numberToDate(42829.8333333333)).to.deep.eq(new Date('04 Apr 2017 20:00:00'));
        });
    });
});
