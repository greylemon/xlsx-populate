"use strict";

const xmlq = require("../xml/xmlq");
const Style = require("../worksheets/Style");
const { hash } = require('../xml/hash');
const { cloneDeep } = require('../utils');

/**
 * Optimize fills, borders, fonts. numFmts nodes.
 * @param {string} OuterNodeName
 * @param {{}} oldOuterNode
 * @param {number} numOfReservedNodes
 * @param {string} attributeName
 * @returns {any[]}
 */
const optimizeNode = (OuterNodeName, oldOuterNode, numOfReservedNodes, attributeName) => {
    const map = new Map();
    const newOuterNode = {
        name: OuterNodeName,
        attributes: oldOuterNode.attributes,
        children: []
    };

    // add reserved nodes
    for (let i = 0; i < numOfReservedNodes; i++) {
        const oldInnerNode = oldOuterNode.children[i];
        newOuterNode.children.push(oldInnerNode);
        map.set(hash(oldInnerNode), i);
    }

    const updateXfNode = xf => {
        const oldInnerNode = oldOuterNode.children[xf.attributes[attributeName]];
        const key = hash(oldInnerNode);
        let idx = map.get(key);
        if (idx === undefined) {
            idx = newOuterNode.children.length;
            newOuterNode.children.push(oldInnerNode);
            map.set(key, idx);
        }
        xf.attributes[attributeName] = idx;
    };

    return [newOuterNode, updateXfNode];
};

/**
 * Standard number format codes
 * Taken from http://polymathprogrammer.com/2011/02/15/built-in-styles-for-excel-open-xml/
 * @private
 */
const STANDARD_CODES = {
    0: 'General',
    1: '0',
    2: '0.00',
    3: '#,##0',
    4: '#,##0.00',
    9: '0%',
    10: '0.00%',
    11: '0.00E+00',
    12: '# ?/?',
    13: '# ??/??',
    14: 'mm-dd-yy',
    15: 'd-mmm-yy',
    16: 'd-mmm',
    17: 'mmm-yy',
    18: 'h:mm AM/PM',
    19: 'h:mm:ss AM/PM',
    20: 'h:mm',
    21: 'h:mm:ss',
    22: 'm/d/yy h:mm',
    37: '#,##0 ;(#,##0)',
    38: '#,##0 ;[Red](#,##0)',
    39: '#,##0.00;(#,##0.00)',
    40: '#,##0.00;[Red](#,##0.00)',
    45: 'mm:ss',
    46: '[h]:mm:ss',
    47: 'mmss.0',
    48: '##0.0E+0',
    49: '@'
};

/**
 * The starting ID for custom number formats. The first 163 indexes are reserved.
 * @private
 */
const STARTING_CUSTOM_NUMBER_FORMAT_ID = 164;

/**
 * A style sheet.
 * Creates an instance of _StyleSheet.
 * @ignore
 */
class StyleSheet {
    /**
     * Creates an instance of _StyleSheet.
     * @param {{}} node - The style sheet node
     */
    constructor(node) {
        this._init(node);
        this._cacheNumberFormats();
        /**
         * @type {Map<string, number>}
         * @private
         */
        this._styleMap = new Map();
        /**
         * @type {Map<number, number>}
         * @private
         */
        this._oldStyleId2NewStyleId = new Map();
    }

    /**
     * Create a style.
     * @param {number} [sourceId] - The source style ID to copy, if provided.
     * @returns {Style} The style.
     */
    createStyle(sourceId) {
        let fontNode, fillNode, borderNode, xfNode;
        if (sourceId >= 0) {
            const sourceXfNode = this._cellXfsNode.children[sourceId];
            xfNode = cloneDeep(sourceXfNode);

            if (sourceXfNode.attributes.applyFont) {
                const fontId = sourceXfNode.attributes.fontId;
                fontNode = cloneDeep(this._fontsNode.children[fontId]);
            }

            if (sourceXfNode.attributes.applyFill) {
                const fillId = sourceXfNode.attributes.fillId;
                fillNode = cloneDeep(this._fillsNode.children[fillId]);
            }

            if (sourceXfNode.attributes.applyBorder) {
                const borderId = sourceXfNode.attributes.borderId;
                borderNode = cloneDeep(this._bordersNode.children[borderId]);
            }
        }

        if (!fontNode) fontNode = { name: "font", attributes: {}, children: [] };
        this._fontsNode.children.push(fontNode);

        if (!fillNode) fillNode = { name: "fill", attributes: {}, children: [] };
        this._fillsNode.children.push(fillNode);

        // The border sides must be in order
        if (!borderNode) borderNode = { name: "border", attributes: {}, children: [] };
        borderNode.children = [
            xmlq.findChild(borderNode, "left") || { name: "left", attributes: {}, children: [] },
            xmlq.findChild(borderNode, "right") || { name: "right", attributes: {}, children: [] },
            xmlq.findChild(borderNode, "top") || { name: "top", attributes: {}, children: [] },
            xmlq.findChild(borderNode, "bottom") || { name: "bottom", attributes: {}, children: [] },
            xmlq.findChild(borderNode, "diagonal") || { name: "diagonal", attributes: {}, children: [] }
        ];
        this._bordersNode.children.push(borderNode);

        if (!xfNode) xfNode = { name: "xf", attributes: {}, children: [] };
        Object.assign(xfNode.attributes, {
            fontId: this._fontsNode.children.length - 1,
            fillId: this._fillsNode.children.length - 1,
            borderId: this._bordersNode.children.length - 1,
            applyFont: 1,
            applyFill: 1,
            applyBorder: 1
        });

        this._cellXfsNode.children.push(xfNode);

        return new Style(this, this._cellXfsNode.children.length - 1, xfNode, fontNode, fillNode, borderNode);
    }

    /**
     * Get the number format code for a given ID.
     * @param {number} id - The number format ID.
     * @returns {string} The format code.
     */
    getNumberFormatCode(id) {
        return this._numberFormatCodesById[id];
    }

    /**
     * Get the nuumber format ID for a given code.
     * @param {string} code - The format code.
     * @returns {number} The number format ID.
     */
    getNumberFormatId(code) {
        let id = this._numberFormatIdsByCode[code];
        if (id === undefined) {
            id = this._nextNumFormatId++;
            this._numberFormatCodesById[id] = code;
            this._numberFormatIdsByCode[code] = id;

            this._numFmtsNode.children.push({
                name: 'numFmt',
                attributes: {
                    numFmtId: id,
                    formatCode: code
                }
            });
        }

        return id;
    }

    /**
     * Get the index of xf node in cellXfs.
     * Note: usually call this after optimize().
     * @param xfNode
     * @return {number|undefined}
     */
    getStyleId(xfNode) {
        return this._styleMap.get(hash(xfNode));
    }

    /**
     * Get the new style ID given the old style id.
     * This method is only for updating row and column style id.
     * @param id
     * @returns {number}
     */
    getRowColNewStyleId(id) {
        return this._oldStyleId2NewStyleId.get(id);
    }

    /**
     * Optimize the StyleSheet. Link duplicated nodes and remove useless nodes. Should be called before saving.
     * Call it before saving.
     * @returns {undefined}
     */
    optimize() {
        // optimize borders, fills, fonts node.
        const [newBordersNode, updateBorderId] = optimizeNode('borders', this._bordersNode, 1, 'borderId');
        const [newFillsNode, updateFillId] = optimizeNode('fills', this._fillsNode, 2, 'fillId');
        const [newFontsNode, updateFontId] = optimizeNode('fonts', this._fontsNode, 1, 'fontId');

        for (const xf of this._cellXfsNode.children) {
            updateBorderId(xf);
            updateFillId(xf);
            updateFontId(xf);
        }
        for (const xf of this._cellStyleXfsNode.children) {
            updateBorderId(xf);
            updateFillId(xf);
            updateFontId(xf);
        }

        // get current node order
        const nodeOrder = this._node.children.map(node => node.name);

        xmlq.removeChild(this._node, this._bordersNode);
        xmlq.insertInOrder(this._node, newBordersNode, nodeOrder);
        this._bordersNode = newBordersNode;

        xmlq.removeChild(this._node, this._fillsNode);
        xmlq.insertInOrder(this._node, newFillsNode, nodeOrder);
        this._fillsNode = newFillsNode;

        xmlq.removeChild(this._node, this._fontsNode);
        xmlq.insertInOrder(this._node, newFontsNode, nodeOrder);
        this._fontsNode = newFontsNode;

        this._numFmtsNode.attributes.count = this._numFmtsNode.children.length;
        this._fontsNode.attributes.count = this._fontsNode.children.length;
        this._fillsNode.attributes.count = this._fillsNode.children.length;
        this._bordersNode.attributes.count = this._bordersNode.children.length;
        this._cellStyleXfsNode.attributes.count = this._cellStyleXfsNode.children.length;

        // optimize cellXfs
        this._styleMap = new Map();
        const newCellXfsChildren = [];
        const isFirstTime = this._oldStyleId2NewStyleId.size === 0;
        const map = new Map();
        let i = 0;
        for (const xf of this._cellXfsNode.children) {
            const key = hash(xf);
            let newIndex = this._styleMap.get(key);
            if (newIndex == null) {
                newIndex = newCellXfsChildren.length;
                newCellXfsChildren.push(xf);
                this._styleMap.set(key, newIndex);
            }
            if (isFirstTime) {
                this._oldStyleId2NewStyleId.set(i, newIndex);
            } else {
                map.set(i, newIndex);
            }
            i++;
        }
        this._cellXfsNode.children = newCellXfsChildren;
        this._cellXfsNode.attributes.count = this._cellXfsNode.children.length;

        if (!isFirstTime) {
            for (const [key, value] of this._oldStyleId2NewStyleId) {
                const newValue = map.get(value);
                if (newValue != null)  this._oldStyleId2NewStyleId.set(key, newValue);
            }
        }
    }

    /**
     * Convert the style sheet to an XML object.
     * @returns {{}} The XML form.
     * @ignore
     */
    toXml() {
        return this._node;
    }

    /**
     * Cache the number format codes
     * @returns {undefined}
     * @private
     */
    _cacheNumberFormats() {
        // Load the standard number format codes into the caches.
        this._numberFormatCodesById = {};
        this._numberFormatIdsByCode = {};
        for (const id in STANDARD_CODES) {
            if (!STANDARD_CODES.hasOwnProperty(id)) continue;
            const code = STANDARD_CODES[id];
            this._numberFormatCodesById[id] = code;
            this._numberFormatIdsByCode[code] = parseInt(id);
        }

        // Set the next number format code.
        this._nextNumFormatId = STARTING_CUSTOM_NUMBER_FORMAT_ID;

        // If there are custom number formats, cache them all and update the next num as needed.
        this._numFmtsNode.children.forEach(node => {
            const id = node.attributes.numFmtId;
            const code = node.attributes.formatCode;
            this._numberFormatCodesById[id] = code;
            this._numberFormatIdsByCode[code] = id;
            if (id >= this._nextNumFormatId) this._nextNumFormatId = id + 1;
        });
    }

    /**
     * Initialize the style sheet node.
     * @param {{}} [node] - The node
     * @returns {undefined}
     * @private
     */
    _init(node) {
        this._node = node;

        // Cache the refs to the collections.
        this._numFmtsNode = xmlq.findChild(this._node, "numFmts");
        this._fontsNode = xmlq.findChild(this._node, "fonts");
        this._fillsNode = xmlq.findChild(this._node, "fills");
        this._bordersNode = xmlq.findChild(this._node, "borders");
        this._cellXfsNode = xmlq.findChild(this._node, "cellXfs");
        this._cellStyleXfsNode = xmlq.findChild(this._node, "cellStyleXfs");

        if (!this._numFmtsNode) {
            this._numFmtsNode = {
                name: "numFmts",
                attributes: {},
                children: []
            };

            // Number formats need to be before the others.
            xmlq.insertBefore(this._node, this._numFmtsNode, this._fontsNode);
        }

        // Remove the optional counts so we don't have to keep them up to date.
        delete this._numFmtsNode.attributes.count;
        delete this._fontsNode.attributes.count;
        delete this._fillsNode.attributes.count;
        delete this._bordersNode.attributes.count;
        delete this._cellXfsNode.attributes.count;
        delete this._cellStyleXfsNode.attributes.count;
    }
}

module.exports = StyleSheet;

/*
xl/styles.xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main">
    <numFmts count="1">
        <numFmt numFmtId="164" formatCode="#,##0_);[Red]\(#,##0\)\)"/>
    </numFmts>
    <fonts count="1" x14ac:knownFonts="1">
        <font>
            <sz val="11"/>
            <color theme="1"/>
            <name val="Calibri"/>
            <family val="2"/>
            <scheme val="minor"/>
        </font>
    </fonts>
    <fills count="11">
        <fill>
            <patternFill patternType="none"/>
        </fill>
        <fill>
            <patternFill patternType="gray125"/>
        </fill>
        <fill>
            <patternFill patternType="solid">
                <fgColor rgb="FFC00000"/>
                <bgColor indexed="64"/>
            </patternFill>
        </fill>
        <fill>
            <patternFill patternType="lightDown">
                <fgColor theme="4"/>
                <bgColor rgb="FFC00000"/>
            </patternFill>
        </fill>
        <fill>
            <gradientFill degree="90">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill>
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill degree="45">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill degree="135">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill type="path">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill type="path" left="0.5" right="0.5" top="0.5" bottom="0.5">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill degree="270">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
    </fills>
    <borders count="10">
        <border>
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal/>
        </border>
        <border diagonalDown="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="hair">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalDown="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="dotted">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalDown="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="dashDotDot">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalDown="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="dashDot">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalDown="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="dashed">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalUp="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="mediumDashDotDot">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalUp="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="slantDashDot">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalUp="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="mediumDashDot">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalUp="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="mediumDashed">
                <color auto="1"/>
            </diagonal>
        </border>
    </borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="19">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="3" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="4" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="5" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="6" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="7" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="8" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="9" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="7" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="8" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="9" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="10" borderId="0" xfId="0" applyFill="1"/>
    </cellXfs>
    <cellStyles count="1">
        <cellStyle name="Normal" xfId="0" builtinId="0"/>
    </cellStyles>
    <dxfs count="0"/>
    <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
    <extLst>
        <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
        </ext>
        <ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
            <x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/>
        </ext>
    </extLst>
</styleSheet>
*/
