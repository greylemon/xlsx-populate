"use strict";

const xmlq = require("../../lib/xml/xmlq");
const expect = require('chai').expect;

describe("xmlq", () => {
    describe("appendChild", () => {
        it("should append the child", () => {
            const node = { name: 'parent', children: [{ name: 'existing' }] };
            const child = { name: 'new' };
            xmlq.appendChild(node, child);
            expect(node).to.deep.eq({ name: 'parent', children: [{ name: 'existing' }, { name: 'new' }] });
        });

        it("should create the children array if needed", () => {
            const node = { name: 'parent' };
            const child = { name: 'new' };
            xmlq.appendChild(node, child);
            expect(node).to.deep.eq({ name: 'parent', children: [{ name: 'new' }] });
        });
    });

    describe("appendChildIfNotFound", () => {
        it("should append the child", () => {
            const node = { name: 'parent', children: [{ name: 'existing' }] };
            expect(xmlq.appendChildIfNotFound(node, 'new')).to.deep.eq({ name: 'new', attributes: {}, children: [] });
            expect(node).to.deep.eq({ name: 'parent', children: [{ name: 'existing' }, { name: 'new', attributes: {}, children: [] }] });
        });

        it("should not append the child", () => {
            const node = { name: 'parent', children: [{ name: 'existing' }] };
            expect(xmlq.appendChildIfNotFound(node, 'existing')).to.deep.eq({ name: 'existing' });
            expect(node).to.deep.eq({ name: 'parent', children: [{ name: 'existing' }] });
        });
    });

    describe("findChild", () => {
        it("should return the child", () => {
            const node = { name: 'parent', children: [{ name: 'A' }, { name: 'B' }] };
            expect(xmlq.findChild(node, 'A')).to.eq(node.children[0]);
            expect(xmlq.findChild(node, 'B')).to.eq(node.children[1]);
            expect(xmlq.findChild(node, 'C')).to.be.undefined;
        });
    });

    describe("getChildAttribute", () => {
        it("should return the child attribute", () => {
            const node = { name: 'parent', children: [
                { name: 'A' },
                { name: 'B', attributes: {} },
                { name: 'C', attributes: { foo: "FOO" } }
            ] };

            expect(xmlq.getChildAttribute(node, 'A', 'foo')).to.be.undefined;
            expect(xmlq.getChildAttribute(node, 'B', 'foo')).to.be.undefined;
            expect(xmlq.getChildAttribute(node, 'C', 'foo')).to.eq('FOO');
        });
    });

    describe("hasChild", () => {
        it("should return true/false", () => {
            const node = { name: 'parent', children: [{ name: 'A' }, { name: 'B' }] };
            expect(xmlq.hasChild(node, 'A')).to.eq(true);
            expect(xmlq.hasChild(node, 'B')).to.eq(true);
            expect(xmlq.hasChild(node, 'C')).to.eq(false);
        });
    });

    describe("insertAfter", () => {
        it("should insert the child after the node", () => {
            const node = { name: 'parent', children: [{ name: 'A' }, { name: 'B' }] };
            xmlq.insertAfter(node, { name: 'new' }, node.children[0]);
            expect(node.children).to.deep.eq([{ name: 'A' }, { name: 'new' }, { name: 'B' }]);
        });
    });

    describe("insertBefore", () => {
        it("should insert the child before the node", () => {
            const node = { name: 'parent', children: [{ name: 'A' }, { name: 'B' }] };
            xmlq.insertBefore(node, { name: 'new' }, node.children[1]);
            expect(node.children).to.deep.eq([{ name: 'A' }, { name: 'new' }, { name: 'B' }]);
        });
    });

    describe("insertInOrder", () => {
        it("should insert in the middle", () => {
            const node = { name: 'parent', children: [{ name: 'A' }, { name: 'C' }] };
            xmlq.insertInOrder(node, { name: 'B' }, ['A', 'B', 'C']);
            expect(node.children).to.deep.eq([{ name: 'A' }, { name: 'B' }, { name: 'C' }]);
        });

        it("should insert at the beginning", () => {
            const node = { name: 'parent', children: [{ name: 'C' }] };
            xmlq.insertInOrder(node, { name: 'A' }, ['A', 'B', 'C']);
            expect(node.children).to.deep.eq([{ name: 'A' }, { name: 'C' }]);
        });

        it("insert at the end", () => {
            const node = { name: 'parent', children: [{ name: 'A' }] };
            xmlq.insertInOrder(node, { name: 'C' }, ['A', 'B', 'C']);
            expect(node.children).to.deep.eq([{ name: 'A' }, { name: 'C' }]);
        });

        it("append if node not expected in order", () => {
            const node = { name: 'parent', children: [{ name: 'A' }, { name: 'C' }] };
            xmlq.insertInOrder(node, { name: 'D' }, ['A', 'B', 'C']);
            expect(node.children).to.deep.eq([{ name: 'A' }, { name: 'C' }, { name: 'D' }]);
        });
    });

    describe("isEmpty", () => {
        it("should return true/false", () => {
            const nodes = [
                { name: 'A' },
                { name: 'B', attributes: {}, children: [] },
                { name: 'C', attributes: { foo: 1 }, children: [] },
                { name: 'D', attributes: {}, children: [{}] },
                { name: 'E', attributes: { foo: 0 }, children: [{}] }
            ];

            expect(xmlq.isEmpty(nodes[0])).to.eq(true);
            expect(xmlq.isEmpty(nodes[1])).to.eq(true);
            expect(xmlq.isEmpty(nodes[2])).to.eq(false);
            expect(xmlq.isEmpty(nodes[3])).to.eq(false);
            expect(xmlq.isEmpty(nodes[4])).to.eq(false);
        });
    });

    describe("removeChild", () => {
        it("should remove the children", () => {
            const node = { name: 'parent', children: [{ name: 'A' }, { name: 'B' }, { name: 'C' }] };
            xmlq.removeChild(node, node.children[1]);
            expect(node.children).to.deep.eq([{ name: 'A' }, { name: 'C' }]);
            xmlq.removeChild(node, 'A');
            expect(node.children).to.deep.eq([{ name: 'C' }]);
            xmlq.removeChild(node, 'foo');
            expect(node.children).to.deep.eq([{ name: 'C' }]);
        });
    });

    describe("setAttributes", () => {
        it("should set/unset the attributes", () => {
            const node = { attributes: { foo: 1, bar: 1, baz: 1 } };
            xmlq.setAttributes(node, {
                foo: undefined,
                bar: null,
                goo: 1,
                gar: 1
            });
            expect(node.attributes).to.deep.eq({
                baz: 1,
                goo: 1,
                gar: 1
            });
        });
    });

    describe("setChildAttributes", () => {
        it("should append the child with the attributes", () => {
            const node = { name: 'parent', children: [{ name: 'existing' }] };
            expect(xmlq.setChildAttributes(node, 'new', { foo: 1, bar: null })).to.deep.eq({ name: 'new', attributes: { foo: 1 }, children: [] });
            expect(node.children).to.deep.eq([{ name: 'existing' }, { name: 'new', attributes: { foo: 1 }, children: [] }]);
        });

        it("should not append the child but should set the attributes", () => {
            const node = { name: 'parent', children: [{ name: 'existing', attributes: { bar: 1 } }] };
            expect(xmlq.setChildAttributes(node, 'existing', { foo: 1, bar: null })).to.deep.eq({ name: 'existing', attributes: { foo: 1 } });
            expect(node).to.deep.eq({ name: 'parent', children: [{ name: 'existing', attributes: { foo: 1 } }] });
        });
    });

    describe("removeChildIfEmpty", () => {
        it("should remove the children", () => {
            const node = { name: 'parent', children: [{ name: 'A' }, { name: 'B' }, { name: 'C', attributes: { foo: 1 } }] };
            xmlq.removeChildIfEmpty(node, node.children[1]);
            expect(node.children).to.deep.eq([{ name: 'A' }, { name: 'C', attributes: { foo: 1 } }]);
            xmlq.removeChildIfEmpty(node, 'A');
            expect(node.children).to.deep.eq([{ name: 'C', attributes: { foo: 1 } }]);
            xmlq.removeChildIfEmpty(node, 'C');
            expect(node.children).to.deep.eq([{ name: 'C', attributes: { foo: 1 } }]);
            xmlq.removeChildIfEmpty(node, node.children[0]);
            expect(node.children).to.deep.eq([{ name: 'C', attributes: { foo: 1 } }]);
            xmlq.removeChildIfEmpty(node, 'foo');
            expect(node.children).to.deep.eq([{ name: 'C', attributes: { foo: 1 } }]);
        });
    });
});
