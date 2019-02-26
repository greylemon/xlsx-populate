"use strict";

const xmlq = require("./xmlq");

/**
 * A readonly theme.
 */
class Theme {
    /**
     * Creates a new instance of Theme. (Read only)
     * @param {{}} node - The node.
     */
    constructor(node) {
        if (!node) {
            return;
        }
        this._node = node;
        const themeElements = xmlq.findChild(node, "a:themeElements");
        if (!themeElements) return;
        this._clrScheme = xmlq.findChild(themeElements, "a:clrScheme");
        this._colorKeys = ['a:lt1', 'a:dk1', 'a:lt2', 'a:dk2', 'a:accent1', 'a:accent2', 'a:accent3', 'a:accent4',
            'a:accent5', 'a:accent6', 'a:hlink', 'a:folHlink'];
    }

    /**
     * Get the rgb theme color
     * @param {number} index - theme color index
     * @param {number|undefined|null} tint - tint value
     * @return {string} color
     */
    themeColor(index, tint) {
        if (!this._node) throw new Error(`Theme.themeColor: This workbook does not have theme.`);
        index = parseInt(index);
        tint = parseFloat(tint);
        let color;
        let colorTag = xmlq.findChild(this._clrScheme, this._colorKeys[index]);
        colorTag = colorTag.children[0];
        if (index === 0 || index === 1) {
            color = colorTag.attributes[Object.keys(colorTag.attributes)[1]];
        } else {
            color = colorTag.attributes[Object.keys(colorTag.attributes)[0]];
        }
        if (tint === undefined || tint === null) {
            return color;
        }
        const rgb = Theme._hexToRgb(color);
        const hsl = Theme._rgbToHsl(rgb.r, rgb.g, rgb.b);

        // the following link tells how to calculate given a tint value.
        // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tabcolor.aspx
        if (tint < 0) {
            hsl.l *= 1 + tint;
            return Theme._rgbToHex(Theme._hslToRgb(hsl.h, hsl.s, hsl.l));
        } else if (tint > 0) {
            hsl.l = hsl.l * (1 - tint) + tint;
            return Theme._rgbToHex(Theme._hslToRgb(hsl.h, hsl.s, hsl.l));
        } else {
            // no change
            return color;
        }
    }


    /**
     *
     * @param {string} hex string
     * @return {{r: number, b: number, g: number}} rgb
     * @private
     */
    static _hexToRgb(hex) {
        const r = parseInt(hex.substring(0, 2), 16);
        const g = parseInt(hex.substring(2, 4), 16);
        const b = parseInt(hex.substring(4, 6), 16);
        return { r, g, b };
    }

    /**
     *
     * @param {{r: number, b: number, g: number}} rgb representation
     * @return {string} hex string
     * @private
     */
    static _rgbToHex(rgb) {
        return Math.round(rgb.r).toString(16) + Math.round(rgb.g).toString(16) + Math.round(rgb.b).toString(16);
    }

    /**
     * Converts an RGB color value to HSL. Conversion formula
     * adapted from http://en.wikipedia.org/wiki/HSL_color_space.
     * Assumes r, g, and b are contained in the set [0, 255] and
     * returns h, s, and l in the set [0, 1].
     *
     * @private
     * @param   {number}  r       The red color value
     * @param   {number}  g       The green color value
     * @param   {number}  b       The blue color value
     * @return  {{h: number, s: number, l: number}}   The HSL representation
     */
    static _rgbToHsl(r, g, b) {
        r /= 255;
        g /= 255;
        b /= 255;

        const max = Math.max(r, g, b), min = Math.min(r, g, b);
        let h, s;
        const l = (max + min) / 2;

        if (max === min) {
            h = s = 0; // achromatic
        } else {
            const d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);

            switch (max) {
                case r:
                    h = (g - b) / d + (g < b ? 6 : 0);
                    break;
                case g:
                    h = (b - r) / d + 2;
                    break;
                case b:
                    h = (r - g) / d + 4;
                    break;
                default:
                    break;
            }
            h /= 6;
        }
        return { h, s, l };
    }

    /**
     * Converts an HSL color value to RGB. Conversion formula
     * adapted from http://en.wikipedia.org/wiki/HSL_color_space.
     * Assumes h, s, and l are contained in the set [0, 1] and
     * returns r, g, and b in the set [0, 255].
     *
     * @private
     * @param   {number}  h       The hue
     * @param   {number}  s       The saturation
     * @param   {number}  l       The lightness
     * @return  {{r: number, b: number, g: number}}          The RGB representation
     */
    static _hslToRgb(h, s, l) {
        let r, g, b;

        if (s === 0) {
            r = g = b = l; // achromatic
        } else {
            const hue2rgb = (p, q, t) => {
                if (t < 0) t += 1;
                if (t > 1) t -= 1;
                if (t < 1 / 6) return p + (q - p) * 6 * t;
                if (t < 1 / 2) return q;
                if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
                return p;
            };

            const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            const p = 2 * l - q;

            r = hue2rgb(p, q, h + 1 / 3);
            g = hue2rgb(p, q, h);
            b = hue2rgb(p, q, h - 1 / 3);
        }
        return { r: r * 255, g: g * 255, b: b * 255 };
    }
}

module.exports = Theme;

/*
xl/theme/theme1.xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
    <a:themeElements>
        <a:clrScheme name="Office">
            <a:dk1>
                <a:sysClr val="windowText" lastClr="000000"/>
            </a:dk1>
            <a:lt1>
                <a:sysClr val="window" lastClr="FFFFFF"/>
            </a:lt1>
            <a:dk2>
                <a:srgbClr val="44546A"/>
            </a:dk2>
            <a:lt2>
                <a:srgbClr val="E7E6E6"/>
            </a:lt2>
            <a:accent1>
                <a:srgbClr val="5B9BD5"/>
            </a:accent1>
            <a:accent2>
                <a:srgbClr val="ED7D31"/>
            </a:accent2>
            <a:accent3>
                <a:srgbClr val="A5A5A5"/>
            </a:accent3>
            <a:accent4>
                <a:srgbClr val="FFC000"/>
            </a:accent4>
            <a:accent5>
                <a:srgbClr val="4472C4"/>
            </a:accent5>
            <a:accent6>
                <a:srgbClr val="70AD47"/>
            </a:accent6>
            <a:hlink>
                <a:srgbClr val="0563C1"/>
            </a:hlink>
            <a:folHlink>
                <a:srgbClr val="954F72"/>
            </a:folHlink>
        </a:clrScheme>
        <a:fontScheme name="Office">
            <a:majorFont>
                <a:latin typeface="Calibri Light" panose="020F0302020204030204"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
                <a:font script="Jpan" typeface="游ゴシック Light"/>
                <a:font script="Hang" typeface="맑은 고딕"/>
                <a:font script="Hans" typeface="等线 Light"/>
                <a:font script="Hant" typeface="新細明體"/>
                <a:font script="Arab" typeface="Times New Roman"/>
                <a:font script="Hebr" typeface="Times New Roman"/>
                <a:font script="Thai" typeface="Tahoma"/>
                <a:font script="Ethi" typeface="Nyala"/>
                <a:font script="Beng" typeface="Vrinda"/>
                <a:font script="Gujr" typeface="Shruti"/>
                <a:font script="Khmr" typeface="MoolBoran"/>
                <a:font script="Knda" typeface="Tunga"/>
                <a:font script="Guru" typeface="Raavi"/>
                <a:font script="Cans" typeface="Euphemia"/>
                <a:font script="Cher" typeface="Plantagenet Cherokee"/>
                <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
                <a:font script="Tibt" typeface="Microsoft Himalaya"/>
                <a:font script="Thaa" typeface="MV Boli"/>
                <a:font script="Deva" typeface="Mangal"/>
                <a:font script="Telu" typeface="Gautami"/>
                <a:font script="Taml" typeface="Latha"/>
                <a:font script="Syrc" typeface="Estrangelo Edessa"/>
                <a:font script="Orya" typeface="Kalinga"/>
                <a:font script="Mlym" typeface="Kartika"/>
                <a:font script="Laoo" typeface="DokChampa"/>
                <a:font script="Sinh" typeface="Iskoola Pota"/>
                <a:font script="Mong" typeface="Mongolian Baiti"/>
                <a:font script="Viet" typeface="Times New Roman"/>
                <a:font script="Uigh" typeface="Microsoft Uighur"/>
                <a:font script="Geor" typeface="Sylfaen"/>
            </a:majorFont>
            <a:minorFont>
                <a:latin typeface="Calibri" panose="020F0502020204030204"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
                <a:font script="Jpan" typeface="游ゴシック"/>
                <a:font script="Hang" typeface="맑은 고딕"/>
                <a:font script="Hans" typeface="等线"/>
                <a:font script="Hant" typeface="新細明體"/>
                <a:font script="Arab" typeface="Arial"/>
                <a:font script="Hebr" typeface="Arial"/>
                <a:font script="Thai" typeface="Tahoma"/>
                <a:font script="Ethi" typeface="Nyala"/>
                <a:font script="Beng" typeface="Vrinda"/>
                <a:font script="Gujr" typeface="Shruti"/>
                <a:font script="Khmr" typeface="DaunPenh"/>
                <a:font script="Knda" typeface="Tunga"/>
                <a:font script="Guru" typeface="Raavi"/>
                <a:font script="Cans" typeface="Euphemia"/>
                <a:font script="Cher" typeface="Plantagenet Cherokee"/>
                <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
                <a:font script="Tibt" typeface="Microsoft Himalaya"/>
                <a:font script="Thaa" typeface="MV Boli"/>
                <a:font script="Deva" typeface="Mangal"/>
                <a:font script="Telu" typeface="Gautami"/>
                <a:font script="Taml" typeface="Latha"/>
                <a:font script="Syrc" typeface="Estrangelo Edessa"/>
                <a:font script="Orya" typeface="Kalinga"/>
                <a:font script="Mlym" typeface="Kartika"/>
                <a:font script="Laoo" typeface="DokChampa"/>
                <a:font script="Sinh" typeface="Iskoola Pota"/>
                <a:font script="Mong" typeface="Mongolian Baiti"/>
                <a:font script="Viet" typeface="Arial"/>
                <a:font script="Uigh" typeface="Microsoft Uighur"/>
                <a:font script="Geor" typeface="Sylfaen"/>
            </a:minorFont>
        </a:fontScheme>
        <a:fmtScheme name="Office">
            <a:fillStyleLst>
                <a:solidFill>
                    <a:schemeClr val="phClr"/>
                </a:solidFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:lumMod val="110000"/>
                                <a:satMod val="105000"/>
                                <a:tint val="67000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="50000">
                            <a:schemeClr val="phClr">
                                <a:lumMod val="105000"/>
                                <a:satMod val="103000"/>
                                <a:tint val="73000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:lumMod val="105000"/>
                                <a:satMod val="109000"/>
                                <a:tint val="81000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:lin ang="5400000" scaled="0"/>
                </a:gradFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:satMod val="103000"/>
                                <a:lumMod val="102000"/>
                                <a:tint val="94000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="50000">
                            <a:schemeClr val="phClr">
                                <a:satMod val="110000"/>
                                <a:lumMod val="100000"/>
                                <a:shade val="100000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:lumMod val="99000"/>
                                <a:satMod val="120000"/>
                                <a:shade val="78000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:lin ang="5400000" scaled="0"/>
                </a:gradFill>
            </a:fillStyleLst>
            <a:lnStyleLst>
                <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr"/>
                    </a:solidFill>
                    <a:prstDash val="solid"/>
                    <a:miter lim="800000"/>
                </a:ln>
                <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr"/>
                    </a:solidFill>
                    <a:prstDash val="solid"/>
                    <a:miter lim="800000"/>
                </a:ln>
                <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr"/>
                    </a:solidFill>
                    <a:prstDash val="solid"/>
                    <a:miter lim="800000"/>
                </a:ln>
            </a:lnStyleLst>
            <a:effectStyleLst>
                <a:effectStyle>
                    <a:effectLst/>
                </a:effectStyle>
                <a:effectStyle>
                    <a:effectLst/>
                </a:effectStyle>
                <a:effectStyle>
                    <a:effectLst>
                        <a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
                            <a:srgbClr val="000000">
                                <a:alpha val="63000"/>
                            </a:srgbClr>
                        </a:outerShdw>
                    </a:effectLst>
                </a:effectStyle>
            </a:effectStyleLst>
            <a:bgFillStyleLst>
                <a:solidFill>
                    <a:schemeClr val="phClr"/>
                </a:solidFill>
                <a:solidFill>
                    <a:schemeClr val="phClr">
                        <a:tint val="95000"/>
                        <a:satMod val="170000"/>
                    </a:schemeClr>
                </a:solidFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:tint val="93000"/>
                                <a:satMod val="150000"/>
                                <a:shade val="98000"/>
                                <a:lumMod val="102000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="50000">
                            <a:schemeClr val="phClr">
                                <a:tint val="98000"/>
                                <a:satMod val="130000"/>
                                <a:shade val="90000"/>
                                <a:lumMod val="103000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:shade val="63000"/>
                                <a:satMod val="120000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:lin ang="5400000" scaled="0"/>
                </a:gradFill>
            </a:bgFillStyleLst>
        </a:fmtScheme>
    </a:themeElements>
    <a:objectDefaults/>
    <a:extraClrSchemeLst/>
    <a:extLst>
        <a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}">
            <thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme"
                               id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}"
                               vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/>
        </a:ext>
    </a:extLst>
</a:theme>

*/

