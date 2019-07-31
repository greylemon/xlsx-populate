"use strict";

const Cell = require("../../lib/worksheets/Cell");
const Column = require("../../lib/worksheets/Column");
const FormulaError = require('../../lib/FormulaError');
const RichTexts = require('../../lib/worksheets/RichText');
const XlsxPopulate = require('../../lib/XlsxPopulate');
const expect = require('chai').expect;

describe("Cell", () => {
    let cell, cellNode, row, sheet, workbook, sharedStrings, styleSheet, style;
    let b2, a1, c7, formulaCell, sharedFormulaCell, e3, e4, hyperlinkCell,
        dataValidationCell;

    beforeEach(async () => {
        workbook = await XlsxPopulate.fromFileAsync('../files/cell.xlsx');
        sheet = workbook.sheet(0);
        b2 = sheet.cell('B2');
        a1 = sheet.cell('A1');
        cell = c7 = sheet.cell('C7');
        formulaCell = sheet.cell('D2');
        sharedFormulaCell = sheet.cell('E2');
        e3 = sheet.cell('E3');
        e4 = sheet.cell('E4');
        hyperlinkCell = sheet.cell('F1');
        dataValidationCell = sheet.cell('G1');
    });

    describe("active", () => {
        it("should return true/false", () => {
            expect(a1.active()).to.eq(false);
            expect(b2.active()).to.eq(true);
        });

        it("should set the sheet active cell", () => {
            expect(a1.active(true)).to.eq(a1);
            expect(a1.active(true)).to.eq(a1);
        });

        it("should throw an error if attempting to deactivate", () => {
            expect(() => cell.active(false)).to.throw();
        });
    });

    describe("address", () => {
        it("should return the address", () => {
            expect(cell.address()).to.eq('C7');
            expect(cell.address({ rowAnchored: true })).to.eq('C$7');
            expect(cell.address({ columnAnchored: true })).to.eq('$C7');
            expect(cell.address({ includeSheetName: true })).to.eq("'Sheet1'!C7");
            expect(cell.address({
                includeSheetName: true,
                rowAnchored: true,
                columnAnchored: true
            })).to.eq("'Sheet1'!$C$7");
            expect(cell.address({ anchored: true })).to.eq("$C$7");
        });
    });

    describe("column", () => {
        it("should return the parent column", () => {
            expect(cell.column()).instanceOf(Column);
            expect(cell.column().columnNumber()).to.eq(3);
        });
    });

    describe("clear", () => {
        it("should clear the cell contents", () => {
            expect(formulaCell.clear()).to.eq(formulaCell);
            expect(formulaCell._value).to.eq(undefined);
            expect(formulaCell._formulaType).to.eq(undefined);
            expect(formulaCell._formula).to.eq(undefined);
            expect(formulaCell._sharedFormulaId).to.eq(undefined);

        });

        it("should clear the cell with shared ref", () => {
            expect(sharedFormulaCell.clear()).to.eq(sharedFormulaCell);
            expect(sharedFormulaCell._value).to.eq(undefined);
            expect(sharedFormulaCell._formulaType).to.eq(undefined);
            expect(sharedFormulaCell._formula).to.eq(undefined);
            expect(sharedFormulaCell._sharedFormulaId).to.eq(undefined);
            expect(sharedFormulaCell._formulaRef).to.eq(undefined);
        });
    });

    describe("columnName", () => {
        it("should return the column name", () => {
            expect(cell.columnName()).to.eq("C");
        });
    });

    describe("columnNumber", () => {
        it("should return the column number", () => {
            expect(cell.columnNumber()).to.eq(3);
        });
    });

    describe("formula", () => {
        it("should return undefined if formula not set", () => {
            expect(cell.formula()).to.eq(undefined);
        });

        it("should return the formula if set", () => {
            expect(formulaCell.formula()).to.eq("SUM(1+3)");
        });

        it("should return the shared formula if the ref cell", () => {
            expect(sharedFormulaCell.formula()).to.eq("SUM(A1:A2)");
            expect(e3.formula()).to.eq("SUM(A2:A3)");
            expect(e4.formula()).to.eq("SUM(A3:A4)");
        });

        it("should clear the formula", () => {
            cell.formula("SUM(1+1)");
            expect(cell.formula(undefined)).to.eq(cell);

            expect(cell._formula).to.eq(undefined);
            expect(cell._formulaType).to.eq(undefined);
        });

        it("should set the formula and clear the value", () => {
            cell.formula("SUM(1+1)");
            expect(cell.formula("SUM(1+2)")).to.eq(cell);

            expect(cell._value).to.eq(3);
            expect(cell._formula).to.eq("SUM(1+2)");
            expect(cell._formulaType).to.eq('normal');
        });
    });

    describe("hyperlink", () => {
        it("should get the hyperlink from the sheet", () => {
            expect(hyperlinkCell.hyperlink().retrieve()).to.eq(a1);
        });

        it("should set the hyperlink on the sheet", () => {
            cell.hyperlink(a1);
            expect(cell.hyperlink().retrieve()).to.eq(a1);
        });

        it("should set the hyperlink with tooltip on the sheet", () => {
            const opts = { hyperlink: "www.google.ca", tooltip: "TOOLTIP" };
            expect(cell.hyperlink(opts)).to.eq(cell);
            expect(cell.hyperlink()).to.deep.eq({
                hyperlink: "www.google.ca",
                tooltip: "TOOLTIP"
            });
        });

        it("should set the hyperlink with email", () => {
            const opts = { email: "example@mail.com", emailSubject: 'subject', tooltip: "TOOLTIP" };
            expect(cell.hyperlink(opts)).to.eq(cell);
            expect(cell.hyperlink()).to.deep.eq({
                hyperlink: "mailto:example@mail.com?subject=subject",
                tooltip: "TOOLTIP"
            });
        });

        it("should clear hyperlink", () => {
            const opts = { hyperlink: "www.google.ca", tooltip: "TOOLTIP" };
            cell.hyperlink(opts);
            cell.hyperlink(undefined);
            expect(cell.hyperlink()).to.eq(undefined);
        });
    });

    describe('dataValidation', () => {
        it('should return the cell', () => {
            const validation = {
                formula1: '"testing, testing2"',
                type: 'list'
            };
            expect(cell.dataValidation(validation)).to.eq(cell);
            expect(cell.dataValidation().formula1).to.deep.eq(validation.formula1);
            expect(cell.dataValidation().formula1Result).to.deep.eq(['testing', 'testing2']);
        });

        it('should return the cell', () => {
            const validations = {
                type: 'list',
                allowBlank: false,
                showInputMessage: false,
                prompt: 'hh',
                promptTitle: 'title',
                showErrorMessage: false,
                error: 'err',
                errorTitle: 'err title',
                formula1: '"test1, test2, test3"'
            };
            expect(cell.dataValidation(validations)).to.eq(cell);
            const result = cell.dataValidation();
            Object.keys(validations).forEach(key => {
                expect(result[key]).to.eq(validations[key], key);
            });
        });

        it("should get the dataValidation from the cell", () => {
            expect(dataValidationCell.dataValidation().formula1Result).to.deep.eq([
                1, 2, 3, 4, 5
            ]);
        });
    });

    describe("find", () => {
        it("should return true if substring found and false otherwise", () => {
            expect(cell.find('data')).to.eq(true);
            expect(cell.find('cell')).to.eq(true);
            expect(cell.find('goo')).to.eq(false);
        });

        it("should return true if regex matches and false otherwise", () => {
            cell.value("Foo bar baz");
            expect(cell.find(/\w{3}/)).to.eq(true);
            expect(cell.find(/\w{4}/)).to.eq(false);
            expect(cell.find(/Foo/)).to.eq(true);
        });

        it("should not replace if replacement is nil", () => {
            cell.value("Foo bar baz");
            expect(cell.find("foo", undefined)).to.eq(true);
            expect(cell.value()).to.eq('Foo bar baz');
            expect(cell.find("bar", null)).to.eq(true);
            expect(cell.value()).to.eq('Foo bar baz');
        });

        it("should replace all occurences of substring", () => {
            cell.value("Foo bar baz");
            expect(cell.find('foo', 'XXX')).to.eq(true);
            expect(cell.value()).to.eq('XXX bar baz');
            cell.value("Foo bar baz foo");
            expect(cell.find('foot', 'XXX')).to.eq(false);
        });

        it("should replace regex matches", () => {
            cell.value("Foo bar baz foo");
            expect(cell.find(/[a-z]{3}/, 'XXX')).to.eq(true);
            expect(cell.value()).to.eq('Foo XXX baz foo');
        });

        it("should replace regex matches", () => {
            cell.value("Foo bar baz foo");
            let times = 0;
            const replacer = (...args) => {
                if (times === 0) expect(args).to.deep.eq(['Foo', 'F', 'oo', 0, 'Foo bar baz foo']);
                else expect(args).to.deep.eq(['foo', 'f', 'oo', 12, 'Foo bar baz foo']);
                times++;
                return 'REPLACEMENT';
            };
            expect(cell.find(/(\w)(o{2})/g, replacer)).to.eq(true);
            expect(cell.value()).to.eq('REPLACEMENT bar baz REPLACEMENT');
        });
    });

    describe("rangeTo", () => {
        it("should create a range", () => {
            const range = cell.rangeTo(a1);
            expect(range.startCell()).to.eq(c7);
            expect(range.endCell()).to.eq(a1);
        });
    });

    describe("relativeCell", () => {
        it("should call sheet.cell with the appropriate row/column", () => {
            expect(cell.relativeCell(0, 0)).to.eq(c7);

            expect(cell.relativeCell(-6, -2)).to.eq(a1);

            expect(cell.relativeCell(1, 1)).to.eq(sheet.cell('D8'));
        });
    });

    describe("row", () => {
        it("should return the parent row", () => {
            expect(cell.row()).to.eq(sheet.row(7));
        });
    });

    describe("rowNumber", () => {
        it("should return the row number", () => {
            expect(cell.rowNumber()).to.eq(7);
        });
    });

    describe("sheet", () => {
        it("should return the parent sheet", () => {
            expect(cell.sheet()).to.eq(sheet);
        });
    });

    describe("style", () => {
        it("should create a new style with the set style ID", () => {
            expect(cell._style).to.eq(undefined);
            cell.style("bold", true);
            expect(cell._style).to.not.eq(undefined);
        });

        it("should not create a style if one already exists", () => {
            cell._style = style;
            cell.style("bold", true);
            const id = cell._style._id;
            cell.style("italic", true);
            expect(cell._style._id).to.eq(id);
        });

        it("should get a single style", () => {
            cell.style("bold", true);
            expect(cell.style('bold')).to.eq(true);
        });

        it("should get multiple styles", () => {
            cell.style("bold", true);
            cell.style("italic", true);
            cell.style("underline", true);
            expect(cell.style(["bold", "italic", "underline"])).to.deep.eq({
                bold: true, italic: true, underline: true
            });
        });

        it("should set the values in the range", () => {
            cell.style("bold", [[true, true], [false, false], [false, true]]);
            expect(cell.relativeCell(0, 0).style('bold')).to.eq(true);
            expect(cell.relativeCell(0, 1).style('bold')).to.eq(true);
            expect(cell.relativeCell(1, 0).style('bold')).to.eq(false);
            expect(cell.relativeCell(1, 1).style('bold')).to.eq(false);
            expect(cell.relativeCell(2, 0).style('bold')).to.eq(false);
            expect(cell.relativeCell(2, 1).style('bold')).to.eq(true);
        });

        it("should set multiple styles", () => {
            expect(cell.style({
                bold: true, italic: true, underline: true
            })).to.eq(cell);
            expect(cell.style('bold')).to.eq(true);
            expect(cell.style('italic')).to.eq(true);
            expect(cell.style('underline')).to.eq(true);
        });
    });

    describe("value", () => {
        beforeEach(() => {
            spyOn(cell, "clear");
        });

        it("should get the value", () => {
            expect(cell.value()).to.eq(undefined);
            cell._value = "foo";
            expect(cell.value()).to.eq('foo');
        });

        it("should clear the cell", () => {
            cell._value = "foo";
            cell.value(undefined);
            expect(cell._value).to.eq(undefined);
            expect(cell.clear).toHaveBeenCalledWith();
        });

        it("should clear the cell and set the value", () => {
            cell.value(5.6);
            expect(cell._value).to.eq(5.6);
            expect(cell.clear).toHaveBeenCalledWith();
        });

        it("should set the values in the range", () => {
            spyOn(cell, "relativeCell");
            cell.value([[1, 2], [3, 4]]);
            expect(cell.relativeCell).toHaveBeenCalledWith(1, 1);
            expect(range.value).toHaveBeenCalledWith([[1, 2], [3, 4]]);
        });
    });

    describe("workbook", () => {
        it("should return the workbook from the row", () => {
            expect(cell.workbook()).to.eq(workbook);
        });
    });

    describe('addHorizontalPageBreak', () => {
        it("should add a rowBreak and return the cell", () => {
            expect(cell.addHorizontalPageBreak()).to.eq(cell);
        });
    });

    /* INTERNAL */

    describe("getSharedRefFormula", () => {
        it("should return the shared ref formula", () => {
            cell._formulaType = 'shared';
            cell._formulaRef = 'REF';
            cell._formula = "FORMULA";
            expect(cell.getSharedRefFormula()).to.eq("FORMULA");
        });

        it("should return undefined if not a ref cell", () => {
            cell._formulaType = 'shared';
            cell._formula = "FORMULA";
            expect(cell.getSharedRefFormula()).to.eq(undefined);
        });

        it("should return undefined if not a shared cell", () => {
            cell._formulaType = 'array';
            cell._formulaRef = 'REF';
            cell._formula = "FORMULA";
            expect(cell.getSharedRefFormula()).to.eq(undefined);
        });
    });

    describe("sharesFormula", () => {
        it("should return true/false if shares a given formula or not", () => {
            cell._formulaType = "shared";
            cell._sharedFormulaId = 6;

            expect(cell.sharesFormula(6)).to.eq(true);
            expect(cell.sharesFormula(3)).to.eq(false);
        });

        it("should return false if it doesn't share any formula", () => {
            expect(cell.sharesFormula(6)).to.eq(false);
        });
    });

    describe("setSharedFormula", () => {
        it("should set a shared formula", () => {
            spyOn(cell, "clear");
            cell.setSharedFormula(3, "A1*A2", "A1:C1");
            expect(cell.clear).toHaveBeenCalledWith();
            expect(cell._formulaType).to.eq("shared");
            expect(cell._sharedFormulaId).to.eq(3);
            expect(cell._formula).to.eq("A1*A2");
            expect(cell._formulaRef).to.eq("A1:C1");
        });
    });

    describe("toXml", () => {
        beforeEach(() => {
            cell.clear();
        });

        it("should set the cell address", () => {
            expect(cell.toXml().attributes.r).to.eq("C7");
        });

        it("should set the formula", () => {
            cell._formulaType = "TYPE";
            cell._formula = "FORMULA";
            cell._formulaRef = "REF";
            cell._sharedFormulaId = "SHARED_ID";

            expect(cell.toXml().children).toEqualJson([{
                name: 'f',
                attributes: {
                    t: 'TYPE',
                    ref: 'REF',
                    si: 'SHARED_ID'
                },
                children: ['FORMULA']
            }]);
        });

        it("should set the formula with remaining attributes", () => {
            cell._formulaType = "normal";
            cell._formula = "FORMULA";
            cell._remainingFormulaAttributes = { foo: 'foo' };

            expect(cell.toXml().children).toEqualJson([{
                name: 'f',
                attributes: {
                    foo: 'foo'
                },
                children: ['FORMULA']
            }]);
        });

        it("should set a string value", () => {
            cell._value = "STRING";

            expect(cell.toXml()).toEqualJson({
                name: 'c',
                attributes: {
                    r: "C7",
                    t: 's'
                },
                children: [{
                    name: 'v',
                    children: [7]
                }]
            });
            expect(sharedStrings.getIndexForString).toHaveBeenCalledWith('STRING');
        });

        it("should set a rich text value", () => {
            const rt = new RichTexts();
            cell._value = rt;

            expect(cell.toXml()).toEqualJson({
                name: 'c',
                attributes: {
                    r: "C7",
                    t: 's'
                },
                children: [{
                    name: 'v',
                    children: [7]
                }]
            });
            expect(sharedStrings.getIndexForString).toHaveBeenCalledWith(rt.toXml());
        });

        it("should set a true bool value", () => {
            cell._value = true;

            expect(cell.toXml()).toEqualJson({
                name: 'c',
                attributes: {
                    r: "C7",
                    t: 'b'
                },
                children: [{
                    name: 'v',
                    children: [1]
                }]
            });
        });

        it("should set a false bool value", () => {
            cell._value = false;

            expect(cell.toXml()).toEqualJson({
                name: 'c',
                attributes: {
                    r: "C7",
                    t: 'b'
                },
                children: [{
                    name: 'v',
                    children: [0]
                }]
            });
        });

        it("should set a number value", () => {
            cell._value = -6.89;

            expect(cell.toXml()).toEqualJson({
                name: 'c',
                attributes: {
                    r: "C7"
                },
                children: [{
                    name: 'v',
                    children: [-6.89]
                }]
            });
        });

        it("should set a date value", () => {
            cell._value = new Date(2017, 0, 1);

            expect(cell.toXml()).toEqualJson({
                name: 'c',
                attributes: {
                    r: "C7"
                },
                children: [{
                    name: 'v',
                    children: [42736]
                }]
            });
        });

        it("should set the defined style id", () => {
            cell._styleId = "STYLE_ID";
            expect(cell.toXml().attributes.s).to.eq("STYLE_ID");
        });

        it("should set the id from the style", () => {
            cell._style = style;
            expect(cell.toXml().attributes.s).to.eq(4);
        });

        it("should handle an empty cell", () => {
            expect(cell.toXml()).toEqualJson({
                name: 'c',
                attributes: {
                    r: "C7"
                },
                children: []
            });
        });

        it("should preserve remaining attributes and children", () => {
            cell._value = -6.89;
            cell._remainingAttributes = { foo: 'foo', bar: 'bar' };
            cell._remainingChildren = [{ name: 'foo' }, { name: 'bar' }];

            expect(cell.toXml()).toEqualJson({
                name: 'c',
                attributes: {
                    r: "C7",
                    foo: 'foo',
                    bar: 'bar'
                },
                children: [{
                    name: 'v',
                    children: [-6.89]
                }, { name: 'foo' }, { name: 'bar' }]
            });
        });
    });

    /* PRIVATE */

    describe("_init", () => {
        beforeEach(() => {
            cell.clear();
            delete cell._columnNumber;
            // spyOn(cell, "_parseNode");
        });

        it("should parse the node", () => {
            const node = {};
            cell._init(node);
            expect(cell._columnNumber).to.eq(undefined);
            expect(cell._parseNode).toHaveBeenCalledWith(node);
        });

        it("should init a cell without a node", () => {
            cell._init(5, 3);
            expect(cell._columnNumber).to.eq(5);
            expect(cell._styleId).to.eq(3);
            expect(cell._parseNode).not.toHaveBeenCalled();
        });
    });

    describe("_parseNode", () => {
        let node;

        beforeEach(() => {
            node = {
                attributes: {
                    r: "D8"
                }
            };

            cell.clear();
            delete cell._columnNumber;
            sheet.updateMaxSharedFormulaId.calls.reset();
        });

        it("should parse the column number", () => {
            cell._parseNode(node);
            expect(cell._columnNumber).to.eq(4);
        });

        it("should store the style ID", () => {
            node.attributes.s = "STYLE_ID";
            cell._parseNode(node);
            expect(cell._styleId).to.eq("STYLE_ID");
        });

        it("should parse a normal formula", () => {
            node.children = [{
                name: 'f',
                attributes: {},
                children: ["FORMULA"]
            }];

            cell._parseNode(node);
            expect(cell._formulaType).to.eq("normal");
            expect(cell._formula).to.eq("FORMULA");
            expect(cell._formulaRef).to.eq(undefined);
            expect(cell._sharedFormulaId).to.eq(undefined);
            expect(cell._remainingFormulaAttributes).to.eq(undefined);
            expect(sheet.updateMaxSharedFormulaId).not.toHaveBeenCalled();
        });

        it("should parse a shared formula", () => {
            node.children = [{
                name: 'f',
                attributes: {
                    t: "shared",
                    ref: "REF",
                    si: "SHARED_INDEX"
                },
                children: ["FORMULA"]
            }];

            cell._parseNode(node);
            expect(cell._formulaType).to.eq("shared");
            expect(cell._formula).to.eq("FORMULA");
            expect(cell._formulaRef).to.eq("REF");
            expect(cell._sharedFormulaId).to.eq("SHARED_INDEX");
            expect(cell._remainingFormulaAttributes).to.eq(undefined);
            expect(sheet.updateMaxSharedFormulaId).toHaveBeenCalledWith("SHARED_INDEX");
        });

        it("should preserve unknown formula attributes", () => {
            node.children = [{
                name: 'f',
                attributes: {
                    t: "TYPE",
                    foo: "foo",
                    bar: "bar"
                },
                children: []
            }];

            cell._parseNode(node);
            expect(cell._formulaType).to.eq("TYPE");
            expect(cell._formula).to.eq(undefined);
            expect(cell._formulaRef).to.eq(undefined);
            expect(cell._sharedFormulaId).to.eq(undefined);
            expect(cell._remainingFormulaAttributes).toEqualJson({
                foo: "foo",
                bar: "bar"
            });
            expect(sheet.updateMaxSharedFormulaId).not.toHaveBeenCalled();
        });

        it("should parse string values", () => {
            node.attributes.t = "s";
            node.children = [{
                name: 'v',
                children: [5]
            }];

            cell._parseNode(node);
            expect(cell._value).to.eq("STRING");
            expect(sharedStrings.getStringByIndex).toHaveBeenCalledWith(5);
        });

        it("should parse string values with no shared string child", () => {
            node.attributes.t = "s";
            node.children = [];

            cell._parseNode(node);
            expect(cell._value).to.eq("");
        });

        it("should parse simple string values", () => {
            node.attributes.t = "str";
            node.children = [{
                name: 'v',
                children: ['SIMPLE STRING']
            }];

            cell._parseNode(node);
            expect(cell._value).to.eq("SIMPLE STRING");
        });

        it("should parse inline string values", () => {
            node.attributes.t = "inlineStr";
            node.children = [{
                name: 'is',
                children: [{
                    name: 't',
                    children: ["INLINE_STRING"]
                }]
            }];

            cell._parseNode(node);
            expect(cell._value).to.eq("INLINE_STRING");
        });

        it("should parse inline string rich text values", () => {
            node.attributes.t = "inlineStr";
            node.children = [{
                name: 'is',
                children: [{
                    name: 'r',
                    children: [{
                        name: 't',
                        children: "FOO"
                    }]
                }]
            }];

            cell._parseNode(node);
            expect(cell._value).toEqualJson([{
                name: 'r',
                children: [{
                    name: 't',
                    children: "FOO"
                }]
            }]);
        });

        it("should parse true values", () => {
            node.attributes.t = "b";
            node.children = [{
                name: 'v',
                children: [1]
            }];

            cell._parseNode(node);
            expect(cell._value).to.eq(true);
        });

        it("should parse false values", () => {
            node.attributes.t = "b";
            node.children = [{
                name: 'v',
                children: [0]
            }];

            cell._parseNode(node);
            expect(cell._value).to.eq(false);
        });

        it("should parse error values", () => {
            node.attributes.t = "e";
            node.children = [{
                name: 'v',
                children: ["#ERR"]
            }];

            cell._parseNode(node);
            expect(cell._value).to.eq("ERROR");
            expect(FormulaError.getError).toHaveBeenCalledWith("#ERR");
        });

        it("should parse number values", () => {
            node.children = [{
                name: 'v',
                children: [-1.67]
            }];

            cell._parseNode(node);
            expect(cell._value).to.eq(-1.67);
            expect(cell._remainingAttributes).to.eq(undefined);
            expect(cell._remainingChildren).to.eq(undefined);
        });

        it("should handle empty cells", () => {
            cell._parseNode(node);
            expect(cell._value).to.eq(undefined);
        });

        it("should preserve unknown attributes and children", () => {
            node.attributes.foo = "foo";
            node.attributes.bar = "bar";
            node.children = [{
                name: 'v',
                children: [0]
            }, {
                name: 'foo'
            }, {
                name: 'bar'
            }];

            cell._parseNode(node);
            expect(cell._value).to.eq(0);
            expect(cell._remainingAttributes).toEqualJson({
                foo: "foo",
                bar: "bar"
            });
            expect(cell._remainingChildren).toEqualJson([
                { name: 'foo' },
                { name: 'bar' }
            ]);
        });
    });
});
