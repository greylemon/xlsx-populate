"use strict";

const { replaceRowNumber, removeReference, MODE } = require("../../../lib/formula/Replacer");
const expect = require('chai').expect;

describe("Replacer", () => {
    describe("Replacer.replaceRowNumber", () => {
        describe("delete rows", () => {
            it("should replace cell reference", done => {
                const res = replaceRowNumber('Sheet1!F30', 'Sheet1', 'Sheet1', 30, 29);
                expect(res).to.eq('Sheet1!F29');
                done();
            });

            it("should replace range reference", done => {
                let res = replaceRowNumber('Sheet1!F29:H34', 'Sheet1', 'Sheet1', 30, 29);
                expect(res).to.eq('Sheet1!F29:H33');
                res = replaceRowNumber('Sheet1!F29:H34', 'Sheet1', 'Sheet1', 30, 28);
                expect(res).to.eq('Sheet1!F29:H32');
                res = replaceRowNumber('Sheet1!F29:H34', 'Sheet1', 'Sheet1', 34, 29);
                expect(res).to.eq('#REF!');
                res = replaceRowNumber('Sheet1!F29:H34', 'Sheet1', 'Sheet1', 35, 28);
                expect(res).to.eq('#REF!');
                done();
            });

            it("should replace row reference", done => {
                let res = replaceRowNumber('Sheet1!29:34', 'Sheet1', 'Sheet1', 30, 29);
                expect(res).to.eq('Sheet1!29:33');
                res = replaceRowNumber('Sheet1!29:34', 'Sheet1', 'Sheet1', 30, 28);
                expect(res).to.eq('Sheet1!29:32');
                res = replaceRowNumber('Sheet1!29:34+1', 'Sheet1', 'Sheet1', 34, 29);
                expect(res).to.eq('#REF!+1');
                res = replaceRowNumber('Sheet1!29:34', 'Sheet1', 'Sheet1', 35, 28);
                expect(res).to.eq('#REF!');
                done();
            });
        });

        describe("add rows", () => {
            it("should replace cell reference", done => {
                const res = replaceRowNumber('Sheet1!F29', 'Sheet1', 'Sheet1', 29, 30);
                expect(res).to.eq('Sheet1!F30');
                done();
            });

            it("should replace range reference", done => {
                let res = replaceRowNumber('Sheet1!F29:H34', 'Sheet1', 'Sheet1', 29, 30);
                expect(res).to.eq('Sheet1!F30:H35');
                res = replaceRowNumber('Sheet1!F29:H34', 'Sheet1', 'Sheet1', 34, 35);
                expect(res).to.eq('Sheet1!F29:H35');
                res = replaceRowNumber('Sheet1!F29:H34', 'Sheet1', 'Sheet1', 30, 32);
                expect(res).to.eq('Sheet1!F29:H36');
                done();
            });

            it("should replace row reference", done => {
                let res = replaceRowNumber('Sheet1!29:34', 'Sheet1', 'Sheet1', 29, 30);
                expect(res).to.eq('Sheet1!30:35');
                res = replaceRowNumber('Sheet1!29:34', 'Sheet1', 'Sheet1', 34, 35);
                expect(res).to.eq('Sheet1!29:35');
                res = replaceRowNumber('Sheet1!29:34', 'Sheet1', 'Sheet1', 30, 32);
                expect(res).to.eq('Sheet1!29:36');
                done();
            });
        });
    });
    describe("removeReference", () => {
        describe("remove row", () => {
            it('should remove cell reference', done => {
                const res = removeReference('Sheet1!F30+1', 'Sheet1', { sheet: 'Sheet1', row: 30, col: 6 }, MODE.ROW);
                expect(res).to.eq('#REF!+1');
                done();
            });

            it('should update range reference', done => {
                let res = removeReference('Sheet1!F29:H34', 'Sheet1', { sheet: 'Sheet1', row: 30, col: 6 }, MODE.ROW);
                expect(res).to.eq('Sheet1!F29:H33');
                res = removeReference('Sheet1!F29:H34', 'Sheet1', { sheet: 'Sheet1', row: 29, col: 6 }, MODE.ROW);
                expect(res).to.eq('Sheet1!F29:H33');
                res = removeReference('Sheet1!F29:H34', 'Sheet1', { sheet: 'Sheet1', row: 27, col: 6 }, MODE.ROW);
                expect(res).to.eq('Sheet1!F29:H34');
                res = removeReference('Sheet1!F29:F29', 'Sheet1', { sheet: 'Sheet1', row: 29, col: 6 }, MODE.ROW);
                expect(res).to.eq('#REF!');
                done();
            });

            it('should update row reference', done => {
                let res = removeReference('Sheet1!29:34', 'Sheet1', { sheet: 'Sheet1', row: 30, col: 6 }, MODE.ROW);
                expect(res).to.eq('Sheet1!29:33');
                res = removeReference('Sheet1!29:34', 'Sheet1', { sheet: 'Sheet1', row: 29, col: 6 }, MODE.ROW);
                expect(res).to.eq('Sheet1!29:33');
                res = removeReference('Sheet1!29:34', 'Sheet1', { sheet: 'Sheet1', row: 27, col: 6 }, MODE.ROW);
                expect(res).to.eq('Sheet1!29:34');
                res = removeReference('Sheet1!29:29', 'Sheet1', { sheet: 'Sheet1', row: 29, col: 6 }, MODE.ROW);
                expect(res).to.eq('#REF!');
                done();
            });
        });

        describe("remove col", () => {
            it('should remove cell reference', done => {
                const res = removeReference('Sheet1!F30+1', 'Sheet1', { sheet: 'Sheet1', row: 30, col: 6 }, MODE.COL);
                expect(res).to.eq('#REF!+1');
                done();
            });

            it('should update range reference', done => {
                let res = removeReference('Sheet1!F29:H34', 'Sheet1', { sheet: 'Sheet1', row: 30, col: 6 }, MODE.COL);
                expect(res).to.eq('Sheet1!F29:G34');
                res = removeReference('Sheet1!F29:H34', 'Sheet1', { sheet: 'Sheet1', row: 29, col: 7 }, MODE.COL);
                expect(res).to.eq('Sheet1!F29:G34');
                res = removeReference('Sheet1!F29:H34', 'Sheet1', { sheet: 'Sheet1', row: 29, col: 3 }, MODE.COL);
                expect(res).to.eq('Sheet1!F29:H34');
                res = removeReference('Sheet1!F29:F34', 'Sheet1', { sheet: 'Sheet1', row: 29, col: 6 }, MODE.COL);
                expect(res).to.eq('#REF!');
                done();
            });

            it('should update col reference', done => {
                let res = removeReference('Sheet1!F:H', 'Sheet1', { sheet: 'Sheet1', row: 30, col: 6 }, MODE.COL);
                expect(res).to.eq('Sheet1!F:G');
                res = removeReference('Sheet1!F:H', 'Sheet1', { sheet: 'Sheet1', row: 29, col: 7 }, MODE.COL);
                expect(res).to.eq('Sheet1!F:G');
                res = removeReference('Sheet1!F:H', 'Sheet1', { sheet: 'Sheet1', row: 29, col: 3 }, MODE.COL);
                expect(res).to.eq('Sheet1!F:H');
                res = removeReference('Sheet1!F:F', 'Sheet1', { sheet: 'Sheet1', row: 29, col: 6 }, MODE.COL);
                expect(res).to.eq('#REF!');
                done();
            });
        });
    });
});
