"use strict";

const XlsxPoplate = require('../../lib/XlsxPopulate');
const RichText = require('../../lib/worksheets/RichText');
const RichTextFragment = require('../../lib/worksheets/RichTextFragment');
const expect = require('chai').expect;

describe("RichText", () => {
    let cell, workbook, cell2, cell3;

    beforeEach(done => {
        XlsxPoplate.fromBlankAsync()
            .then(wb => {
                workbook = wb;
                cell = workbook.sheet(0).cell(1, 1);
                cell2 = workbook.sheet(0).cell(1, 2);
                cell3 = workbook.sheet(0).cell(1, 3);
                done();
            });
    });

    it('global export', () => {
        expect(RichText === XlsxPoplate.RichText).to.eq(true);
    });

    describe("add/get", () => {
        it("should add and get normal text", () => {
            const rt = new RichText();
            cell.value(rt);
            expect(cell.value() instanceof RichText).to.eq(true);
            rt.add('hello');
            rt.add('hello2');
            expect(rt.length).to.eq(2);
            expect(rt.get(0).value()).to.eq('hello');
            expect(rt.get(1).value()).to.eq('hello2');
        });

        it("should transfer line separator to \r\n", () => {
            const rt = new RichText();
            rt.add('hello\n');
            rt.add('hel\r\nlo2');
            rt.add('hel\rlo2');
            cell.value(rt);
            expect(rt.get(0).value()).to.eq('hello\r\n');
            expect(rt.get(1).value()).to.eq('hel\r\nlo2');
            expect(rt.get(2).value()).to.eq('hel\r\nlo2');
        });

        it("should set wrapText to true", () => {
            const rt = new RichText();
            rt.add('hello\n');
            cell.value(rt);
            expect(cell.style('wrapText')).to.eq(true);
        });

        it("should set style", () => {
            const rt = new RichText();
            rt.add('hello\n', { bold: true });
            cell.value(rt);
            expect(rt.get(0).style('bold')).to.eq(true);
        });
    });

    it("should clear the rich text", () => {
        const rt = new RichText();
        rt.add('hello');
        rt.clear();
        expect(rt.text()).to.eq('');
    });

    it("should get concatenated text", () => {
        const rt = new RichText();
        rt.add('hello')
            .add(' I', { fontColor: 'FF0000FF' })
            .add("'m \n ")
            .add('lester');

        expect(rt.text()).to.eq("hello I'm \r\n lester");
    });

    describe("change related cell", () => {
        it("should set wrapText to true in the new cell", () => {
            const rt = new RichText();
            rt.add('hello\n');
            cell.value(rt);
            expect(cell.style('wrapText')).to.eq(true);
        });
    });

    describe('Cell.value', () => {
        it('should assign a deep copy of rich text instance', () => {
            const rt = new RichText();
            rt.add('string\n');
            cell.value(rt);
            cell2.value(rt);
            const value1 = cell.value(), value2 = cell2.value();
            expect(value1).not.to.eq(rt);
            expect(value2).not.to.eq(rt);
            expect(value1).not.to.eq(value2);

            value1.add('test');
            expect(cell.value().text()).to.eq('string\r\ntest');
            expect(cell2.value().text()).to.eq('string\r\n');
        });

        it('should get instance with cell reference', () => {
            const rt = new RichText();
            rt.add('string');
            cell.value(rt);
            expect(cell.value().cell).to.eq(cell);
        });

        it('should re-assign cell reference', () => {
            const rt = new RichText();
            rt.add('string');
            cell.value(rt);
            expect(cell.value().cell).to.eq(cell);
            const value = cell.value();
            cell2.value(value);
            expect(cell2.value().cell).to.eq(cell2);
        });
    });

    describe('Sheet.range', () => {
        it('should set range of rich texts', () => {
            const rt = new RichText();
            rt.add('string');
            workbook.sheet(0).range('A1:C3').value(rt);
            expect(cell.value().cell).to.eq(cell);
            expect(cell.value().text()).to.eq('string');
            expect(cell.value()).not.to.eq(rt);
            expect(cell.value()).not.to.eq(cell2.value());
        });
    });

    describe('styles', () => {
        let fontNode, fragment;

        beforeEach(() => {
            fragment = new RichTextFragment('text');
            fontNode = fragment._fontNode;
        });

        describe("bold", () => {
            it("should get/set bold", () => {
                expect(fragment.style("bold")).to.eq(false);
                fragment.style("bold", true);
                expect(fragment.style("bold")).to.eq(true);
                expect(fontNode.children).to.deep.eq([{ name: "b", attributes: {}, children: [] }]);
                fragment.style("bold", false);
                expect(fragment.style("bold")).to.eq(false);
                expect(fontNode.children).to.deep.eq([]);
            });
        });

        describe("italic", () => {
            it("should get/set italic", () => {
                expect(fragment.style("italic")).to.eq(false);
                fragment.style("italic", true);
                expect(fragment.style("italic")).to.eq(true);
                expect(fontNode.children).to.deep.eq([{ name: "i", attributes: {}, children: [] }]);
                fragment.style("italic", false);
                expect(fragment.style("italic")).to.eq(false);
                expect(fontNode.children).to.deep.eq([]);
            });
        });

        describe("underline", () => {
            it("should get/set underline", () => {
                expect(fragment.style("underline")).to.eq(false);
                fragment.style("underline", true);
                expect(fragment.style("underline")).to.eq(true);
                expect(fontNode.children).to.deep.eq([{ name: "u", attributes: {}, children: [] }]);
                fragment.style("underline", "double");
                expect(fragment.style("underline")).to.eq("double");
                expect(fontNode.children).to.deep.eq([{ name: "u", attributes: { val: "double" }, children: [] }]);
                fragment.style("underline", true);
                expect(fragment.style("underline")).to.eq(true);
                expect(fontNode.children).to.deep.eq([{ name: "u", attributes: {}, children: [] }]);
                fragment.style("underline", false);
                expect(fragment.style("underline")).to.eq(false);
                expect(fontNode.children).to.deep.eq([]);
            });
        });

        describe("strikethrough", () => {
            it("should get/set strikethrough", () => {
                expect(fragment.style("strikethrough")).to.eq(false);
                fragment.style("strikethrough", true);
                expect(fragment.style("strikethrough")).to.eq(true);
                expect(fontNode.children).to.deep.eq([{ name: 'strike', attributes: {}, children: [] }]);
                fragment.style("strikethrough", false);
                expect(fragment.style("strikethrough")).to.eq(false);
                expect(fontNode.children).to.deep.eq([]);
            });
        });

        describe("subscript", () => {
            it("should get/set subscript", () => {
                expect(fragment.style("subscript")).to.eq(false);
                fragment.style("subscript", true);
                expect(fragment.style("subscript")).to.eq(true);
                expect(fontNode.children).to.deep.eq([{
                    name: "vertAlign",
                    attributes: { val: "subscript" },
                    children: []
                }]);
                fragment.style("subscript", false);
                expect(fragment.style("subscript")).to.eq(false);
                expect(fontNode.children).to.deep.eq([]);
            });
        });

        describe("superscript", () => {
            it("should get/set superscript", () => {
                expect(fragment.style("superscript")).to.eq(false);
                fragment.style("superscript", true);
                expect(fragment.style("superscript")).to.eq(true);
                expect(fontNode.children).to.deep.eq([{
                    name: "vertAlign",
                    attributes: { val: "superscript" },
                    children: []
                }]);
                fragment.style("superscript", false);
                expect(fragment.style("superscript")).to.eq(false);
                expect(fontNode.children).to.deep.eq([]);
            });
        });

        describe("fontSize", () => {
            it("should get/set fontSize", () => {
                expect(fragment.style("fontSize")).to.eq(undefined);
                fragment.style("fontSize", 17);
                expect(fragment.style("fontSize")).to.eq(17);
                expect(fontNode.children).to.deep.eq([{ name: 'sz', attributes: { val: 17 }, children: [] }]);
                fragment.style("fontSize", undefined);
                expect(fragment.style("fontSize")).to.eq(undefined);
                expect(fontNode.children).to.deep.eq([]);
            });
        });

        describe("fontFamily", () => {
            it("should get/set fontFamily", () => {
                expect(fragment.style("fontFamily")).to.eq(undefined);
                fragment.style("fontFamily", "Comic Sans MS");
                expect(fragment.style("fontFamily")).to.eq("Comic Sans MS");
                expect(fontNode.children).to.deep.eq([{
                    name: 'rFont',
                    attributes: { val: "Comic Sans MS" },
                    children: []
                }]);
                fragment.style("fontFamily", undefined);
                expect(fragment.style("fontFamily")).to.eq(undefined);
                expect(fontNode.children).to.deep.eq([]);
            });
        });

        describe("fontGenericFamily", () => {
            it("should get/set fontGenericFamily", () => {
                expect(fragment.style("fontGenericFamily")).to.eq(undefined);
                fragment.style("fontGenericFamily", 1);
                expect(fragment.style("fontGenericFamily")).to.eq(1);
                expect(fontNode.children).to.deep.eq([{ name: 'family', attributes: { val: 1 }, children: [] }]);
                fragment.style("fontGenericFamily", undefined);
                expect(fragment.style("fontGenericFamily")).to.eq(undefined);
                expect(fontNode.children).to.deep.eq([]);
            });
        });

        describe("fontScheme", () => {
            it("should get/set fontScheme", () => {
                expect(fragment.style("fontScheme")).to.eq(undefined);
                fragment.style("fontScheme", 'minor');
                expect(fragment.style("fontScheme")).to.eq('minor');
                expect(fontNode.children).to.deep.eq([{ name: 'scheme', attributes: { val: 'minor' }, children: [] }]);
                fragment.style("fontScheme", undefined);
                expect(fragment.style("fontScheme")).to.eq(undefined);
                expect(fontNode.children).to.deep.eq([]);
            });
        });

        describe("fontColor", () => {
            it("should get/set fontColor", () => {
                expect(fragment.style("fontColor")).to.eq(undefined);

                fragment.style("fontColor", "ff0000");
                expect(fragment.style("fontColor")).to.deep.eq({ rgb: "FF0000" });
                expect(fontNode.children).to.deep.eq([{ name: 'color', attributes: { rgb: "FF0000" }, children: [] }]);

                fragment.style("fontColor", 5);
                expect(fragment.style("fontColor")).to.deep.eq({ theme: 5 });
                expect(fontNode.children).to.deep.eq([{ name: 'color', attributes: { theme: 5 }, children: [] }]);

                fragment.style("fontColor", { theme: 3, tint: -0.2 });
                expect(fragment.style("fontColor")).to.deep.eq({ theme: 3, tint: -0.2 });
                expect(fontNode.children).to.deep.eq([{
                    name: 'color',
                    attributes: { theme: 3, tint: -0.2 },
                    children: []
                }]);

                fragment.style("fontColor", undefined);
                expect(fragment.style("fontColor")).to.eq(undefined);
                expect(fontNode.children).to.deep.eq([]);

                fontNode.children = [{ name: 'color', attributes: { indexed: 7 }, children: [] }];
                expect(fragment.style("fontColor")).to.deep.eq({ rgb: "00FFFF" });
            });
        });
    });
});
