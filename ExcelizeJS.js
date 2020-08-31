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

function timeConverter(UNIX_timestamp) {
    var a = new Date(UNIX_timestamp);
    var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    var year = a.getFullYear();
    var monthNumber = a.getMonth();
    var month = months[monthNumber];
    var date = a.getDate();
    var properDate = "";
    if (date < 10) properDate = month + " 0" + date + ", " + year;
    else properDate = month + " " + date + ", " + year;
    return properDate;
}

function letterCounter(num) {
    "use strict";
    var mod = num % 26,
        pow = num / 26 | 0,
        out = mod ? String.fromCharCode(64 + mod) : (--pow, 'Z');
    return pow ? letterCounter(pow) + out : out;
}

let excelize = function (initDetails, tableData) {
    try {
        initDetails.funcAtStart;
    } catch (e) {
        console.error("ERROR: Error at 'initFuncAtStart'");
        console.error(e);
    }

    initDetails = {
        "fileName": initDetails.fileName || "Excel Export on " + timeConverter(Date.now()),
        "initFuncAtStart": initDetails.initFuncAtStart || null,
        "initFuncAtEnd": initDetails.initFuncAtEnd || null
    }

    tableData = tableData || null;

    if (tableData) {

        let wb = new ExcelJS.Workbook();

        for (var i = 0; i < tableData.length; i++) {

            tableData[i] = {
                "tableID": tableData[i].tableID,
                "colExclude": tableData[i].colExclude || [],
                "rowExclude": tableData[i].rowExclude || [],
                "funcAtStart": tableData[i].funcAtStart || null,
                "funcAtEnd": tableData[i].funcAtEnd || null,
                "consoleLogIteration": tableData[i].consoleLogIteration || false,
                "defaultWidth": tableData[i].defaultWidth || 9,
                "sheetDetails": tableData[i].sheetDetails || {"sheetName": null, "sheetTabColor": null},
                "sheetTabColor": tableData[i].sheetTabColor || null,
                "customWidth": tableData[i].customWidth || [{col: null, width: null}]
            }

            tableIterator(wb, tableData[i], i);
        }

        // Using window.saveAs rather than fileSaver.saveAs
        wb.xlsx.writeBuffer()
            .then(buffer => window.saveAs(new Blob([buffer]), `${initDetails.fileName}.xlsx`))
            .catch(err => console.error('Error writing excel export', err));

    }

    try {
        initDetails.initFuncAtEnd;
    } catch (e) {
        console.error("ERROR: Error at 'initFuncAtEnd'");
        console.error(e);
    }
}

let tableIterator = function (wb, tableDetails, tableIndex) {

    try {
        tableDetails.funcAtStart;
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
    *   comments
    *   about*/

    let totalBodyRows = $(tableDetails.tableID + " > tbody > tr").length - tableDetails.rowExclude.length;
    let totalBodyCols = $(tableDetails.tableID + " > tbody > tr:eq(0) > td").length - tableDetails.colExclude.length;

    try {
        $(tableDetails.tableID + " > thead > tr > th").each(function () {
            if (!tableDetails.colExclude.includes($(this).parent().children().index($(this)).toString().trim())) {
                t = {};
                t["index"] = headIndex++ - 1;
                t["val"] = $(this).text().trim().replace(/\s+/g, '');
                t["col"] = $(this).parent().children().index($(this));
                t["colIndex"] = ++t.index;
                t["colWidth"] = tableDetails.defaultWidth;
                tableDetails.customWidth.forEach(function (cw) {
                    if (t.index === cw.col) t["colWidth"] = cw.width;
                })
                t["excelIndex"] = letterCounter(1 + t.colIndex) + "1";
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
                } catch (e) {
                    t.style.size = 11;
                }
                if ($(this).css("text-decoration").split(" ").includes("underline")) t.style.underline = true;
                try {
                    t.style.family = $(this).css("font-family").split(',')[0].replace(/['"]/g, "");
                } catch (e) {
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
        $(tableDetails.tableID + " > tbody > tr > td").each(function () {
            if (!tableDetails.colExclude.includes($(this).parent().children().index($(this)).toString().trim()) &&
                !tableDetails.rowExclude.includes($(this).parent().parent().children().index($(this).parent()).toString().trim())) {
                t = {};
                t["index"] = bodyIndex++;
                t["val"] = $(this).text().trim().replace(/\s+/g, '');
                t["col"] = $(this).parent().children().index($(this));
                t["row"] = $(this).parent().parent().children().index($(this).parent());
                t["colIndex"] = letterCounter(1 + (t.index % totalBodyCols));
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
                } catch (e) {
                    t.style.size = 11;
                }
                if ($(this).css("text-decoration").split(" ").includes("underline")) t.style.underline = true;
                try {
                    t.style.family = $(this).css("font-family").split(',')[0].replace(/['"]/g, "");
                } catch (e) {
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
        $(tableDetails.tableID + " > tfoot > tr > td").each(function () {
            if (!tableDetails.colExclude.includes($(this).parent().children().index($(this)).toString().trim())) {
                t = {};
                t["index"] = footIndex++ - 1;
                t["val"] = $(this).text().trim().replace(/\s+/g, '');
                t["col"] = $(this).parent().children().index($(this));
                t["colIndex"] = ++t.index;
                t["excelIndex"] = letterCounter(1 + t.colIndex) + (totalBodyRows + 2);
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
                } catch (e) {
                    t.style.size = 11;
                }
                if ($(this).css("text-decoration").split(" ").includes("underline")) t.style.underline = true;
                try {
                    t.style.family = $(this).css("font-family").split(',')[0].replace(/['"]/g, "");
                } catch (e) {
                    t.style.family = 'sans-serif';
                }
                foot.push(t);
            }
        });
    } catch (e) {
        console.error("Error at tfoot reading");
        console.error(e);
    }

    tableData["head"] = head;
    tableData["body"] = body;
    tableData["foot"] = foot;
    tableData["totalBodyRows"] = totalBodyRows;
    tableData["totalBodyCols"] = totalBodyCols;

    headIndex = 0;
    bodyIndex = 0;
    footIndex = 0;

    if (tableDetails.consoleLogIteration) console.log(JSON.stringify(tableData));

    excelCreateAndExport(wb, tableData, tableDetails, tableIndex);

}

let excelCreateAndExport = function (wb, iteratedValue, tableDetails, tableIndex) {

    let sheetName;
    if (tableDetails.sheetDetails.sheetName) sheetName = tableDetails.sheetDetails.sheetName;
    else sheetName = "Sheet" + (tableIndex + 1).toString();

    let sheetTabColor;
    if (tableDetails.sheetDetails.sheetTabColor) sheetTabColor = "FF" + tableDetails.sheetDetails.sheetTabColor.split("#")[1];
    else sheetTabColor = "FFFFFFFF";

    let ws = wb.addWorksheet(sheetName,
        {
            properties: {tabColor: {argb: sheetTabColor}},
            views: [{state: 'frozen', ySplit: 1}]
        });

    let headArray = iteratedValue.head;
    let bodyArray = iteratedValue.body;
    let footArray = iteratedValue.foot;

    let columnConfig = [];
    let colKeyIndex = 1;

    /* todo ask for col widths
        default: 8.11 or 9
        wrap*/

    if (tableDetails.defaultWidth <= 0) tableDetails.defaultWidth = 9;

    headArray.forEach(function (h) {
        temp = {
            header: h.val,
            key: letterCounter(colKeyIndex++),
            width: h.colWidth || 9
        };
        columnConfig.push(temp);
    });

    ws.columns = columnConfig;

    let addBackgroundColor = function (v) {
        return {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {argb: v["background-color"]}
        };
    }

    let addFontStyle = function (v) {
        return {
            color: {argb: v.color},
            name: v.style.family,
            family: 2,
            bold: v.style.bold,
            underline: v.style.underline,
            italic: v.style.italics
        };
    }

    let addStyles = function (ws, v) {
        if (v["background-color"] !== "ffffffff") ws.getCell(v.excelIndex).fill = addBackgroundColor(v);
        ws.getCell(v.excelIndex).font = addFontStyle(v);
        ws.getCell(v.excelIndex).alignment = {wrapText: true};
        return ws;
    }

    let addValues = function (ws, v) {
        if (v.val === "") v.val = "";
        else if (!isNaN(v.val.replace(/,/g, ""))) ws.getCell(v.excelIndex).value = parseInt(v.val);
        else ws.getCell(v.excelIndex).value = v.val;
        return ws;
    }

    bodyArray.forEach(function (v) {
        ws = addValues(ws, v);
        ws = addStyles(ws, v);
    });

    headArray.forEach(function (h) {
        ws = addStyles(ws, h);
    });

    footArray.forEach(function (f) {
        ws = addValues(ws, f);
        ws = addStyles(ws, f);
    });


    headArray.concat(bodyArray).concat(footArray).forEach(function (i) {
        ws.getCell(i.excelIndex).border = {
            top: {style: 'thin', color: {argb: 'FFD4D4D4'}},
            left: {style: 'thin', color: {argb: 'FFD4D4D4'}},
            bottom: {style: 'thin', color: {argb: 'FFD4D4D4'}},
            right: {style: 'thin', color: {argb: 'FFD4D4D4'}}
        };
    });

    try {
        tableDetails.funcAtEnd;
    } catch (e) {
        console.error("ERROR: Error at 'funcAtEnd'");
        console.error(e);
    }
}

