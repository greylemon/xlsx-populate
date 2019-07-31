"use strict";

const proxyquire = require("proxyquire");
const expect = require('chai').expect;

describe("CoreProperties", () => {
    let CoreProperties, coreProperties, corePropertiesNode;

    beforeEach(() => {
        CoreProperties = proxyquire("../../lib/workbooks/CoreProperties", {
            '@noCallThru': true
        });

        corePropertiesNode = {
            name: "Types",
            attributes: {
                xmlns: "http://schemas.openxmlformats.org/package/2006/content-types"
            },
            children: []
        };

        coreProperties = new CoreProperties(corePropertiesNode);
    });

    describe("set", () => {
        it("should set a property value", () => {
            coreProperties.set("Title", "A_TITLE");
            expect(coreProperties._properties.title).to.eq("A_TITLE");
        });

        it("should throw if not an allowed property name", () => {
            const invalidPropertyName = "invalid-property-name";
            expect(() => {
                coreProperties.set(invalidPropertyName, "SOME_VALUE");
            }).to.throw(`Unknown property name: "${invalidPropertyName}"`);
        });
    });

    describe("get", () => {
        it("should get a property value", () => {
            coreProperties.set("title", "A_TITLE");
            expect(coreProperties.get("title")).to.eq("A_TITLE");
        });

        it("should throw if not an allowed property name", () => {
            let invalidPropertyName = "invalid-property-name";
            expect(() => {
                coreProperties.get(invalidPropertyName);
            }).to.throw(`Unknown property name: "${invalidPropertyName}"`);
        });
    });

    describe("toXml", () => {
        it("should return the node as is", () => {
            coreProperties.set("Title", "A_TITLE");

            expect(coreProperties.toXml()).to.eq(corePropertiesNode);
        });
    });
});
