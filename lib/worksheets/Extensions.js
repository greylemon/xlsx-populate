"use strict";

// https://docs.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/07d607af-5618-4ca2-b683-6a78dc0d9627
const ExtURI = {
    conditionalFormattings: '{78C0D931-6437-407D-A8EE-F0AAD7539E65}',
    dataValidations: '{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}',
    sparklineGroups: '{05C60535-1F16-4FD2-B633-F4F36F0B64E0}',
    slicerList: '{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}',
    protectedRanges: '{FC87AEE6-9EDD-4A0A-B7FB-166176984837}',
    ignoredErrors: '{01252117-D84E-4E92-8308-4BE1C098FCBB}',
    webExtensions: '{F7C9EE02-42E1-4005-9D12-6889AFFD525C}',
    timelineRefs: '{7E03D99C-DC04-49d9-9315-930204A7B6E9}'
};
const x14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';

class Extensions {
    constructor(extLstNode) {
        if (extLstNode == null) {
            this._extLstNode = {
                name: 'extLst',
                attributes: {},
                children: []
            };
        } else {
            this._extLstNode = extLstNode;
        }
    }

    /**
     * Get an extension.
     * @param {string} uri - Example: {CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}
     * @return {{}|undefined} The extension
     */
    get(uri) {
        for (let i = 0; i < this._extLstNode.children.length; i++) {
            const extension = this._extLstNode.children[i];
            if (extension.attributes.uri === uri) {
                return extension;
            }
        }
    }

    /**
     * Set and extension, override if exists.
     * @param {string} uri - Example: {CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}
     * @param {{}} node - The node you want it assign to.
     * @return {undefined}
     */
    set(uri, node) {
        if (node.children.length === 0) return;
        for (let i = 0; i < this._extLstNode.children.length; i++) {
            const extension = this._extLstNode.children[i];
            if (extension.attributes.uri === uri) {
                extension.children[0] = [node];
            }
        }

        // add if not found
        this._extLstNode.children.push({
            name: 'ext',
            attributes: {
                'xmlns:x14': x14,
                uri
            },
            children: [node]
        });
    }

    delete(uri) {
        let i;
        for (i = 0; i < this._extLstNode.children.length; i++) {
            const extension = this._extLstNode.children[i];
            if (extension.attributes.uri === uri) {
                break;
            }
        }
        this._extLstNode.children.splice(i, 1);
    }

    toXml() {
        const extensions = [];
        for (let i = 0; i < this._extLstNode.children.length; i++) {
            const extension = this._extLstNode.children[i];
            if (extension.children.length > 0)
                extensions.push(extension);
        }
        this._extLstNode.children = extensions;
        return this._extLstNode;
    }
}

module.exports = { ExtURI, Extensions };
