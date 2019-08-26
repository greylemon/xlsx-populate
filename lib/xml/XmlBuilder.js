"use strict";

const XML_DECLARATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;

/**
 * XML document builder.
 * @private
 */
class XmlBuilder {
    /**
     * Build an XML string from the JSON object.
     * @param {{}} node - The node.
     * @returns {string} The XML text.
     */
    async build(node) {
        this._i = 0;
        const xml = await this._build(node, '');
        if (xml === '') return;
        return XML_DECLARATION + xml;
    }

    /**
     * Build the XML string. (This is the internal recursive method.)
     * @param {{}} node - The node.
     * @param {string} xml - The initial XML doc string.
     * @returns {string} The generated XML element.
     * @private
     */
    async _build(node, xml) {
        // For CPU performance, JS engines don't truly concatenate strings; they create a tree of pointers to
        // the various concatenated strings. The downside of this is that it consumes a lot of memory, which
        // will cause problems with large workbooks. So periodically, we grab a character from the xml, which
        // causes the JS engine to flatten the tree into a single string. Do this too often and CPU takes a hit.
        // Too frequently and memory takes a hit. Every 100k nodes seems to be a good balance.
        if (this._i++ % 1000000 === 0) {
            this._c = xml[0];
        }

        // If the node has a toXml method, call it.
        if (node && typeof node.toXml === 'function') node = await node.toXml();

        if (node != null && typeof node === 'object') {
            // If the node is an object, then it maps to an element. Check if it has a name.
            if (!node.name) throw new Error(`XML node does not have name: ${JSON.stringify(node)}`);

            const attributes = node.attributes;
            const children = node.children;

            // Add the opening tag.
            xml += `<${node.name}`;

            // Add any node attributes
            // eslint-disable-next-line guard-for-in
            for (const name in attributes) {
                xml += ` ${name}="${this._escapeString(attributes[name], true)}"`;
            }

            if (!children || (Array.isArray(children) && (children.length === 0 || children.length === 1 && children[0] == null))
                || children instanceof Map && children.size === 0) {
                // Self-close the tag if no children.
                xml += "/>";
            } else {
                xml += ">";

                // Recursively add any children.
                if (Array.isArray(children)) {
                    // Array
                    for (let i = 0; i < children.length; i++) {
                        xml = await this._build(children[i], xml);
                    }
                } else {
                    // Map
                    const sorted = [...children.entries()].sort((entry1, entry2) => {
                        return entry1[0] - entry2[0];
                    });
                    const childrenXml = await Promise.all(sorted.map(([key, child]) => this._build(child, '')));
                    for (const childXml of childrenXml)
                        xml += childXml;
                }


                // Close the tag.
                xml += `</${node.name}>`;
            }
        } else if (node != null) {
            // It not an object, this should be a text node. Just add it.
            xml += this._escapeString(node);
        }

        // Return the updated XML element.
        return xml;
    }

    /**
     * Escape a string for use in XML by replacing &, ", ', <, and >.
     * @param {*} value - The value to escape.
     * @param {boolean} [isAttribute] - A flag indicating if this is an attribute.
     * @returns {string} The escaped string.
     * @private
     */
    _escapeString(value, isAttribute) {
        if (value == null)
            return value;
        value = value.toString()
            .replace(/&/g, "&amp;") // Escape '&' first as the other escapes add them.
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;");

        if (isAttribute) {
            value = value.replace(/"/g, "&quot;");
        }

        return value;
    }
}

module.exports = XmlBuilder;
