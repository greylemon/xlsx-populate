"use strict";

const ArgHandler = require("../../lib/ArgHandler");
const expect = require('chai').expect;

describe("ArgHandler", () => {
    let argHandler;

    beforeEach(() => {
        // handlers = {
        //     empty: jasmine.createSpy("empty").and.returnValue('empty'),
        //     nil: jasmine.createSpy("nil").and.returnValue("nil"),
        //     string: jasmine.createSpy("string").and.returnValue("string"),
        //     boolean: jasmine.createSpy("boolean").and.returnValue("boolean"),
        //     number: jasmine.createSpy("number").and.returnValue("number"),
        //     integer: jasmine.createSpy("integer").and.returnValue("integer"),
        //     function: jasmine.createSpy("function").and.returnValue("function"),
        //     array: jasmine.createSpy("array").and.returnValue("array"),
        //     date: jasmine.createSpy("date").and.returnValue("date"),
        //     object: jasmine.createSpy("object").and.returnValue("object"),
        //     Style: jasmine.createSpy("style").and.returnValue("Style"),
        //     '*': jasmine.createSpy("*").and.returnValue("*")
        // };

    });

    describe("handle", () => {
        it("should handle empty", done => {
            argHandler = new ArgHandler("METHOD", [])
                .case(() => done()).handle();
        });

        it("should handle nil: undefined", done => {
            argHandler = new ArgHandler("METHOD", [undefined])
                .case('nil', param => {
                    expect(param).to.be.undefined;
                    done();
                })
                .handle();
        });

        it("should handle nil: null", done => {
            argHandler = new ArgHandler("METHOD", [null])
                .case('nil', param => {
                    expect(param).to.be.null;
                    done();
                })
                .handle();
        });

        it("should handle string", done => {
            argHandler = new ArgHandler("METHOD", [''])
                .case('string', param => {
                    expect(param).to.eq('');
                    done();
                })
                .handle();
        });

        it("should handle boolean", done => {
            argHandler = new ArgHandler("METHOD", [true])
                .case('boolean', param => {
                    expect(param).to.eq(true);
                    done();
                })
                .handle();
        });

        it("should handle number", done => {
            argHandler = new ArgHandler("METHOD", [123])
                .case('number', param => {
                    expect(param).to.eq(123);
                    done();
                })
                .handle();
        });

        it("should handle integer", done => {
            expect(() => new ArgHandler("METHOD", [1.5]).case('integer', () => {
            }).handle()).to.throw();
            argHandler = new ArgHandler("METHOD", [123])
                .case('integer', param => {
                    expect(param).to.eq(123);
                    done();
                })
                .handle();
        });

        it("should handle function", done => {
            const func = () => {
            };
            argHandler = new ArgHandler("METHOD", [func])
                .case('function', param => {
                    expect(param).to.eq(func);
                    done();
                })
                .handle();
        });

        it("should handle array", done => {
            argHandler = new ArgHandler("METHOD", [[1, 2, 3]])
                .case('array', param => {
                    expect(param).to.deep.eq([1, 2, 3]);
                    done();
                })
                .handle();
        });

        it("should handle date", done => {
            const date = new Date();
            argHandler = new ArgHandler("METHOD", [date])
                .case('date', param => {
                    expect(param).to.eq(date);
                    done();
                })
                .handle();
        });

        it("should handle object", done => {
            argHandler = new ArgHandler("METHOD", [{ a: 1 }])
                .case('object', param => {
                    expect(param).to.deep.eq({ a: 1 });
                    done();
                })
                .handle();
        });

        it("should handle Styles", done => {
            const Style = class {};

            const style = new Style();
            argHandler = new ArgHandler("METHOD", [style])
                .case('Style', param => {
                    expect(param).to.eq(style);
                    done();
                })
                .handle();
        });

        it("should handle *", done => {
            argHandler = new ArgHandler("METHOD", [undefined, null, 1, true, "1", {}, []])
                .case(['*', '*', '*', '*', '*', '*', '*'], (...params) => {
                    expect(params).to.deep.eq([undefined, null, 1, true, "1", {}, []]);
                    done();
                })
                .handle();
        });
    });
});
