"use strict";

const Relationships = require("../../lib/workbooks/Relationships");
const expect = require('chai').expect;

describe("Relationships", () => {
    let relationships, relationshipsNode;

    beforeEach(() => {
        relationshipsNode = {
            name: "Relationships",
            attributes: {
                xmlns: "http://schemas.openxmlformats.org/package/2006/relationships"
            },
            children: [
                {
                    name: "Relationship",
                    attributes: {
                        Id: "rId2",
                        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                        Target: "theme/theme1.xml"
                    }
                },
                {
                    name: "Relationship",
                    attributes: {
                        Id: "rId1",
                        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                        Target: "worksheets/sheet1.xml"
                    }
                }
            ]
        };

        relationships = new Relationships(relationshipsNode);
    });

    describe("add", () => {
        it("should add a new relationship", () => {
            relationships.add("TYPE", "TARGET");
            expect(relationshipsNode.children[2]).to.deep.eq({
                name: "Relationship",
                attributes: {
                    Id: "rId3",
                    Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/TYPE",
                    Target: "TARGET"
                }
            });
        });

        it("should add a new relationship with target mode", () => {
            relationships.add("TYPE", "TARGET", "TARGET_MODE");
            expect(relationshipsNode.children[2]).to.deep.eq({
                name: "Relationship",
                attributes: {
                    Id: "rId3",
                    Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/TYPE",
                    Target: "TARGET",
                    TargetMode: "TARGET_MODE"
                }
            });
        });
    });

    describe("findById", () => {
        it("should return the relationship if matched", () => {
            expect(relationships.findById("rId1")).to.eq(relationshipsNode.children[1]);
            expect(relationships.findById("rId2")).to.eq(relationshipsNode.children[0]);
        });

        it("should return undefined if not matched", () => {
            expect(relationships.findById("rId5")).to.be.undefined;
        });
    });

    describe("findByType", () => {
        it("should return the relationship if matched", () => {
            expect(relationships.findByType("worksheet")).to.eq(relationshipsNode.children[1]);
            expect(relationships.findByType("theme")).to.eq(relationshipsNode.children[0]);
        });

        it("should return undefined if not matched", () => {
            expect(relationships.findByType("foo")).to.be.undefined;
        });
    });

    describe("toXml", () => {
        it("should return the node as is", () => {
            expect(relationships.toXml()).to.eq(relationshipsNode);
        });

        it("should return undefined", () => {
            relationshipsNode.children.length = 0;
            expect(relationships.toXml()).to.be.undefined;
        });
    });

    describe("_getStartingId", () => {
        it("should set the next ID to 1 if no children", () => {
            relationships._node.children = [];
            relationships._getStartingId();
            expect(relationships._nextId).to.eq(1);
        });

        it("should set the next ID to last found ID + 1", () => {
            relationships._node.children = [
                { attributes: { Id: 'rId2' } },
                { attributes: { Id: 'rId1' } },
                { attributes: { Id: 'rId3' } }
            ];
            relationships._getStartingId();
            expect(relationships._nextId).to.eq(4);
        });
    });

    describe("_init", () => {
        it("should create the node if needed", () => {
            relationships._init(null);
            expect(relationships._node).to.deep.eq({
                name: "Relationships",
                attributes: {
                    xmlns: "http://schemas.openxmlformats.org/package/2006/relationships"
                },
                children: []
            });
        });
    });
});
