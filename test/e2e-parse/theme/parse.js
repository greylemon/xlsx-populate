"use strict";

module.exports = workbook => {
    const color1 = workbook.sheet(0).cell('A1').style('fontColor');
    const color2 = workbook.sheet(0).cell('A2').style('fontColor');
    return [
        workbook.theme().themeColor(color1.theme, color1.tint),
        workbook.theme().themeColor(color2.theme, color2.tint).toUpperCase()
    ];
};
