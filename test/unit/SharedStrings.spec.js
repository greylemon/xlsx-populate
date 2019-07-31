"use strict";

const SharedStrings = require('../../lib/workbooks/SharedStrings');
const expect = require('chai').expect;

describe("SharedStrings", () => {
    let sharedStrings, sharedStringsNode;

    beforeEach(() => {
        sharedStringsNode = {
            name: "sst",
            attributes: {
                xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                count: 3,
                uniqueCount: 7
            },
            children: [
                {
                    name: "si",
                    children: [
                        {
                            name: "t",
                            children: ["foo"]
                        }
                    ]
                }
            ]
        };

        sharedStrings = new SharedStrings(sharedStringsNode);
    });

    describe("getIndexForString", () => {
        beforeEach(() => {
            sharedStrings._stringArray = [
                "foo",
                "bar",
                [{ name: "r", children: [{}] }, { name: "r", children: [{}] }]
            ];

            sharedStrings._indexMap = {
                foo: 0,
                bar: 1,
                '[{"name":"r","children":[{}]},{"name":"r","children":[{}]}]': 2
            };
        });

        it("should return the index if the string already exists", () => {
            expect(sharedStrings.getIndexForString("foo")).to.eq(0);
            expect(sharedStrings.getIndexForString("bar")).to.eq(1);
            expect(sharedStrings.getIndexForString([{ name: "r", children: [{}] }, { name: "r", children: [{}] }])).to.eq(2);
        });

        it("should create a new entry if the string doesn't exist", () => {
            expect(sharedStrings.getIndexForString("baz")).to.eq(3);
            expect(sharedStrings._stringArray).to.deep.eq([
                "foo",
                "bar",
                [{ name: "r", children: [{}] }, { name: "r", children: [{}] }],
                "baz"
            ]);
            expect(sharedStrings._indexMap).to.deep.eq({
                foo: 0,
                bar: 1,
                '[{"name":"r","children":[{}]},{"name":"r","children":[{}]}]': 2,
                baz: 3
            });
            expect(sharedStringsNode.children[sharedStringsNode.children.length - 1]).to.deep.eq({
                name: "si",
                children: [
                    {
                        name: "t",
                        attributes: { 'xml:space': "preserve" },
                        children: ["baz"]
                    }
                ]
            });
        });

        it("should create a new array entry if the array doesn't exist", () => {
            expect(sharedStrings.getIndexForString([{ name: "r", children: [{}] }])).to.eq(3);
            expect(sharedStrings._stringArray).to.deep.eq([
                "foo",
                "bar",
                [{ name: "r", children: [{}] }, { name: "r", children: [{}] }],
                [{ name: "r", children: [{}] }]
            ]);
            expect(sharedStrings._indexMap).to.deep.eq({
                foo: 0,
                bar: 1,
                '[{"name":"r","children":[{}]},{"name":"r","children":[{}]}]': 2,
                '[{"name":"r","children":[{}]}]': 3
            });
            expect(sharedStringsNode.children[sharedStringsNode.children.length - 1]).to.deep.eq({
                name: "si",
                children: [{ name: "r", children: [{}] }]
            });
        });
    });

    describe("getStringByIndex", () => {
        it("should return the string at a given index", () => {
            sharedStrings._stringArray = ["foo", "bar", "baz"];
            expect(sharedStrings.getStringByIndex(0)).to.eq("foo");
            expect(sharedStrings.getStringByIndex(1)).to.eq("bar");
            expect(sharedStrings.getStringByIndex(2)).to.eq("baz");
            expect(sharedStrings.getStringByIndex(3)).to.eq(undefined);
        });
    });

    describe("toXml", () => {
        it("should return the node as is", () => {
            expect(sharedStrings.toXml()).to.deep.eq(sharedStringsNode);
        });
    });

    describe("_cacheExistingSharedStrings", () => {
        it("should cache the existing shared strings", () => {
            sharedStrings._node.children = [
                { name: "si", children: [{ name: "t", children: ["foo"] }] },
                { name: "si", children: [{ name: "t", children: ["bar"] }] },
                { name: "si", children: [{ name: "r", children: [{}] }, { name: "r", children: [{}] }] },
                { name: "si", children: [{ name: "t", children: ["baz"] }] }
            ];

            sharedStrings._stringArray = [];
            sharedStrings._indexMap = {};
            sharedStrings._cacheExistingSharedStrings();

            expect(sharedStrings._stringArray).to.deep.eq([
                "foo",
                "bar",
                [{ name: "r", children: [{}] }, { name: "r", children: [{}] }],
                "baz"
            ]);
            expect(sharedStrings._indexMap).to.deep.eq({
                foo: 0,
                bar: 1,
                '[{"name":"r","children":[{}]},{"name":"r","children":[{}]}]': 2,
                baz: 3
            });
        });
    });

    describe("_init", () => {
        it("should create the node if needed", () => {
            sharedStrings._init(null);
            expect(sharedStrings._node).to.deep.eq({
                name: "sst",
                attributes: {
                    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                },
                children: []
            });
        });

        it("should clear the counts", () => {
            expect(sharedStrings._node.attributes).to.deep.eq({
                xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            });
        });
    });
});
