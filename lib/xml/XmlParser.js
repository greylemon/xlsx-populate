"use strict";

const saxes = require("saxes/lib/saxes");
const parser = new saxes.SaxesParser({ xmlns: true });

// Regex to check if string is all whitespace.
const allWhitespaceRegex = /^\s+$/;

/**
 * XML _parser.
 * @private
 */
class XmlParser {
    /**
     * Parse the XML text into a JSON object.
     * @param {string} xmlText - The XML text.
     * @returns {{}} The JSON object.
     */
    parseAsync(xmlText) {
        return new Promise((resolve, reject) => {
            // Create the SAX _parser.

            // Parsed is the full parsed object. Current is the current node being parsed. Stack is the current stack of
            // nodes leading to the current one.
            let parsed, current;
            const stack = [];

            // On error: Reject the promise.
            parser.onerror = reject;

            // On text nodes: If it is all whitespace, do nothing. Otherwise, try to convert to a number and add as a child.
            parser.ontext = text => {
                if (allWhitespaceRegex.test(text)) {
                    if (current && current.attributes['xml:space'] === 'preserve') {
                        current.children.push(text);
                    }
                } else {
                    current.children.push(this._covertToNumberIfNumber(text));
                }
            };

            // On open tag start: Create a child element. If this is the root element, set it as parsed. Otherwise, add
            // it as a child to the current node.
            parser.onopentag = node => {
                const child = { name: node.name, attributes: {}, children: [] };
                if (current) {
                    current.children.push(child);
                } else {
                    parsed = child;
                }
                Object.keys(node.attributes).forEach(attrKey => {
                    child.attributes[attrKey] = this._covertToNumberIfNumber(node.attributes[attrKey].value);
                });

                stack.push(child);
                current = child;
            };

            // On close tag: Pop the stack.
            parser.onclosetag = node => {
                stack.pop();
                current = stack[stack.length - 1];
            };

            // On end: Resolve the promise.
            parser.onend = () => {
                resolve(parsed);
            };

            // Start parsing the text.
            parser.write(xmlText).close();
        });
    }

    /**
     * Convert the string to a number if it looks like one.
     * @param {string} str - The string to convert.
     * @returns {string|number} The number if converted or the string if not.
     * @private
     */
    _covertToNumberIfNumber(str) {
        const num = Number(str);
        return Number.isNaN(num) || str.substring(0, 2) === '00' ? str: num;
    }
}

module.exports = XmlParser;
