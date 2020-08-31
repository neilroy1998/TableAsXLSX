function rgb2argb(input) {
    let rgb = input;
    try {
        if (/^#[0-9A-F]{6}$/i.test(rgb)) return rgb;

        rgb = rgb.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*(\d+))?\)$/);

        function hex(x) {
            return ("0" + parseInt(x).toString(16)).slice(-2);
        }

        return "ff" + hex(rgb[1]) + hex(rgb[2]) + hex(rgb[3]);
    } catch (e) {
        return input;
    }
}

let excelize = function (tableID, propDetails) {

    try {
        propDetails.funcAtStart;
    } catch (e) {
        console.error("ERROR: Error at 'funcAtStart'");
        console.error(e);
    }

    let tableData = {};

    let head = [];
    let body = [];
    let foot = [];
    let headIndex = 0;
    let bodyIndex = 0;
    let footIndex = 0;

    /* todo
    *   just tr(s)
    *   header
    *   sheets
    *   comments
    *   about*/

    let totalBodyRows = $(tableID + " > tbody > tr").length - propDetails.rowExclude.length;
    let totalBodyCols = $(tableID + " > thead > tr > th").length - propDetails.colExclude.length;

    try {
        $(tableID + " > thead > tr > th").each(function () {
            if (!propDetails.colExclude.includes($(this).parent().children().index($(this)).toString().trim())) {
                t = {};
                t["index"] = headIndex++ - 1;
                t["val"] = $(this).text().trim().replace(/\s+/g, '');
                t["col"] = $(this).parent().children().index($(this));
                t["colIndex"] = ++t.index;
                t["excelIndex"] = String.fromCharCode(65 + t.colIndex) + "1";
                t["background-color"] = rgb2argb($(this).css('background-color'));
                t["color"] = rgb2argb($(this).css('color'));
                if (t["background-color"] === t["color"] && t["color"] === "ff000000") {
                    t["background-color"] = 'ffffffff';
                }
                t["style"] = {
                    "bold": false,
                    "italics": false,
                    "underline": false,
                    "size": 11,
                    "family": 'Calibri'
                };
                if (parseInt($(this).css('font-weight')) >= 700) t.style.bold = true;
                if ($(this).css('font-style').includes("italic")) t.style.italics = true;
                try {
                    t.style.size = $(this).css('font-size').split("px")[0];
                } catch(e) {t.style.size = 11;}
                if ($(this).css("text-decoration").split(" ").includes("underline")) t.style.underline = true;
                try {
                    t.style.family = $(this).css("font-family").split(',')[0].replace(/['"]/g,"");
                } catch(e) {
                    t.style.family = 'sans-serif';
                }
                head.push(t);
            }
        });
    } catch (e) {
        console.error("Error at thead reading");
        console.error(e);
    }

    try {
        $(tableID + " > tbody > tr > td").each(function () {
            if (!propDetails.colExclude.includes($(this).parent().children().index($(this)).toString().trim()) &&
                !propDetails.rowExclude.includes($(this).parent().parent().children().index($(this).parent()).toString().trim())) {
                t = {};
                t["index"] = bodyIndex++;
                t["val"] = $(this).text().trim().replace(/\s+/g, '');
                t["col"] = $(this).parent().children().index($(this));
                t["row"] = $(this).parent().parent().children().index($(this).parent());
                t["colIndex"] = String.fromCharCode(65 + t.index % totalBodyCols);
                t["rowIndex"] = 2 + Math.floor(t.index / totalBodyCols);
                t["excelIndex"] = t.colIndex + t.rowIndex;
                t["background-color"] = rgb2argb($(this).css('background-color'));
                t["color"] = rgb2argb($(this).css('color'));
                if (t["background-color"] === t["color"] && t["color"] === "ff000000") {
                    t["background-color"] = 'ffffffff';
                }
                t["style"] = {
                    "bold": false,
                    "italics": false,
                    "underline": false,
                    "size": 11,
                    "family": 'Calibri'
                };
                if (parseInt($(this).css('font-weight')) >= 700) t.style.bold = true;
                if ($(this).css('font-style').includes("italic")) t.style.italics = true;
                try {
                    t.style.size = $(this).css('font-size').split("px")[0];
                } catch(e) {t.style.size = 11;}
                if ($(this).css("text-decoration").split(" ").includes("underline")) t.style.underline = true;
                try {
                    t.style.family = $(this).css("font-family").split(',')[0].replace(/['"]/g,"");
                } catch(e) {
                    t.style.family = 'sans-serif';
                }
                body.push(t);
            }
        });
    } catch (e) {
        console.error("Error at tbody reading");
        console.error(e);
    }

    try {
        $(tableID + " > tfoot > tr > td").each(function () {
            if (!propDetails.colExclude.includes($(this).parent().children().index($(this)).toString().trim())) {
                t = {};
                t["index"] = footIndex++ - 1;
                t["val"] = $(this).text().trim().replace(/\s+/g, '');
                t["col"] = $(this).parent().children().index($(this));
                t["colIndex"] = ++t.index;
                t["excelIndex"] = String.fromCharCode(65 + t.colIndex) + (totalBodyRows + 2);
                t["background-color"] = rgb2argb($(this).css('background-color'));
                t["color"] = rgb2argb($(this).css('color'));
                if (t["background-color"] === t["color"] && t["color"] === "ff000000") {
                    t["background-color"] = 'ffffffff';
                }
                t["style"] = {
                    "bold": false,
                    "italics": false,
                    "underline": false,
                    "size": 11,
                    "family": 'Calibri'
                };
                if (parseInt($(this).css('font-weight')) >= 700) t.style.bold = true;
                if ($(this).css('font-style').includes("italic")) t.style.italics = true;
                try {
                    t.style.size = $(this).css('font-size').split("px")[0];
                } catch(e) {t.style.size = 11;}
                if ($(this).css("text-decoration").split(" ").includes("underline")) t.style.underline = true;
                try {
                    t.style.family = $(this).css("font-family").split(',')[0].replace(/['"]/g,"");
                } catch(e) {
                    t.style.family = 'sans-serif';
                }
                foot.push(t);
            }
        });
    } catch (e) {
        console.error("Error at thead reading");
        console.error(e);
    }

    tableData["head"] = head;
    tableData["body"] = body;
    tableData["foot"] = foot;
    tableData["totalBodyRows"] = totalBodyRows;
    tableData["totalBodyCols"] = totalBodyCols;

    if (propDetails.consoleLogIteration) console.log(JSON.stringify(tableData));

    headIndex = 0;
    bodyIndex = 0;
    footIndex = 0;

    excelCreateAndExport(tableData, propDetails);
}

let excelCreateAndExport = function (iteratedValue, propDetails) {

    let wb = new ExcelJS.Workbook();
    let ws = wb.addWorksheet("Sheet1");

    let headArray = iteratedValue.head;
    let bodyArray = iteratedValue.body;
    let footArray = iteratedValue.foot;

    let columnConfig = [];
    let colKeyIndex = 65;

    /* todo ask for col widths
        default: 8.11 or 9
        wrap*/

    if (propDetails.defaultWidth <= 0) propDetails.defaultWidth = 9;

    headArray.forEach(function (h) {
        temp = {
            header: h.val,
            key: String.fromCharCode(colKeyIndex++),
            width: propDetails.defaultWidth
        };
        columnConfig.push(temp);
    });

    ws.columns = columnConfig;

    bodyArray.forEach(function (v) {

        if (v.val === "") v.val = "";
        else if (!isNaN(v.val.replace(/,/g, ""))) ws.getCell(v.excelIndex).value = parseInt(v.val);
        else ws.getCell(v.excelIndex).value = v.val;

        if (v["background-color"] !== "ffffffff") {
            ws.getCell(v.excelIndex).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: {argb: v["background-color"]}
            };
        }
        ws.getCell(v.excelIndex).font = {
            color: {argb: v.color},
            name: v.style.family,
            family: 2,
            bold: v.style.bold,
            underline: v.style.underline,
            italic: v.style.italics
        };
        ws.getCell(v.excelIndex).alignment = {wrapText: true};
    });

    headArray.forEach(function (h) {
        if (h["background-color"] !== "ffffffff") {
            ws.getCell(h.excelIndex).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: {argb: h["background-color"]}
            };
        }
        ws.getCell(h.excelIndex).font = {
            color: {argb: h.color},
            name: h.style.family,
            family: 2,
            bold: h.style.bold,
            underline: h.style.underline,
            italic: h.style.italics
        };
        ws.getCell(h.excelIndex).alignment = {wrapText: true};
    });

    footArray.forEach(function (f) {

        if (f.val === "") f.val = "";
        else if (!isNaN(f.val.replace(/,/g, ""))) ws.getCell(f.excelIndex).value = parseInt(f.val);
        else ws.getCell(f.excelIndex).value = f.val;

        if (f["background-color"] !== "ffffffff") {
            ws.getCell(f.excelIndex).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: {argb: f["background-color"]}
            };
        }
        ws.getCell(f.excelIndex).font = {
            color: {argb: f.color},
            name: f.style.family,
            family: 2,
            bold: f.style.bold,
            underline: f.style.underline,
            italic: f.style.italics
        };
        ws.getCell(f.excelIndex).alignment = {wrapText: true};
    });


    headArray.concat(bodyArray).concat(footArray).forEach(function (i) {
        ws.getCell(i.excelIndex).border = {
            top: {style: 'thin', color: {argb: 'FFD4D4D4'}},
            left: {style: 'thin', color: {argb: 'FFD4D4D4'}},
            bottom: {style: 'thin', color: {argb: 'FFD4D4D4'}},
            right: {style: 'thin', color: {argb: 'FFD4D4D4'}}
        };
    });

// Using window.saveAs rather than fileSaver.saveAs
    wb.xlsx.writeBuffer()
        // todo ask for name
        .then(buffer => window.saveAs(new Blob([buffer]), `${propDetails.fileName}.xlsx`))
        .catch(err => console.error('Error writing excel export', err));

    try {
        propDetails.funcAtEnd;
    } catch (e) {
        console.error("ERROR: Error at 'funcAtEnd'");
        console.error(e);
    }
}

