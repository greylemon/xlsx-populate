"use strict";
const parser = require('fast-xml-parser');
const util = require('fast-xml-parser/src/util');

function decode(value) {
    return value.replace(/(&amp;)|(&lt;)|(&gt;)/g, (val, g1, g2, g3) => {
        return g1 ? '&' : (g2 ? '<' : (g3 ? '>' : ''));
    });
}


const parseOptions = {
    attributeNamePrefix: "",
    attrNodeName: "attributes",
    textNodeName: "#text",
    ignoreAttributes: false,
    ignoreNameSpace: false,
    allowBooleanAttributes: true,
    parseNodeValue: true,
    parseAttributeValue: true,
    trimValues: true,
    localeRange: "", // To support non english character in tag/attribute values.
    parseTrueNumberOnly: true,
    tagValueProcessor: a => decode(a)
};

class XmlParser {
    /**
     * Parse the XML text into a JSON object.
     * @param {string} xmlText - The XML text.
     * @returns {Promise} The JSON object.
     */
    parseAsync(xmlText) {
        return new Promise(resolve => {
            const node = parser.getTraversalObj(xmlText, parseOptions);
            resolve(XmlParser._transform(node).children[0]);
        });
    }

    static _transform(node) {
        const jObj = { name: node.tagname, attributes: {}, children: [] };

        // when no child node or attr is present
        if ((!node.child || util.isEmptyObject(node.child)) && (!node.attrsMap || util.isEmptyObject(node.attrsMap))) {
            if (node.val != null && (typeof node.val === 'string' && node.val.length > 0)) {
                jObj.children.push(node.val);
            }
            return jObj;
        }
        if (node.attrsMap && node.attrsMap.attributes)
            jObj.attributes = node.attrsMap.attributes;

        const keys = Object.keys(node.child);
        for (let index = 0; index < keys.length; index++) {
            const tagName = keys[index];
            if (node.child[tagName] && node.child[tagName].length > 1) {
                for (const tag in node.child[tagName]) {
                    jObj.children.push(XmlParser._transform(node.child[tagName][tag]));
                }
            } else {
                jObj.children.push(XmlParser._transform(node.child[tagName][0]));
            }
        }
        return jObj;
    }
}

module.exports = XmlParser;
//
// const xmlText = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
//     '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><authors><author>McClellandM</author><author>Patrick Manhire</author></authors><commentList><comment ref="C9" authorId="0" shapeId="0" xr:uid="{00000000-0006-0000-0200-000001000000}"><text><r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t>This can be found on your payment notice</t></r></text></comment><comment ref="D12" authorId="1" shapeId="0" xr:uid="{00000000-0006-0000-0200-000002000000}"><text><r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t xml:space="preserve">Enter the name of the HSP used most commonly\n' +
//     '</t></r></text></comment><comment ref="D13" authorId="1" shapeId="0" xr:uid="{00000000-0006-0000-0200-000003000000}"><text><r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t xml:space="preserve">Enter legal name of HSP as it appears on your M-SAA\n' +
//     '</t></r></text></comment></commentList></comments>';
//
// const XmlParser2 = require('./XmlParser');
//
// const res = new XmlParser2().parseAsync(xmlText).then(data => {
//     console.log(JSON.stringify(data, null, 2));
// });
