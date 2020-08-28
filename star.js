function rgb2hex(input) {
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

let tableIterate = function (tableID, propDetails) {

    try {
        propDetails.funcAtStart;
    } catch (e) {
        console.log("ERROR: Error at 'funcAtStart'");
        console.log(e);
    }

    let tableData = {};
    let defaultClasses = [];
    if (!propDetails.removeColors) defaultClasses = ["background-color", "color"];

    let head = [];
    let body = [];
    let headIndex = 0;
    let bodyIndex = 0;

    /* todo
    *   borders
    *   tfoot
    *   just tr(s)
    *   header
    *   sheets
    *   comments
    *   about*/

    let totalBodyRows = $("#table1 > tbody > tr").length - propDetails.rowExclude.length;
    let totalBodyCols = $("#table1 > thead > tr > th").length - propDetails.colExclude.length;

    try {
        $(tableID + " > thead > tr > th").each(function () {
            if (!propDetails.colExclude.includes($(this).parent().children().index($(this)).toString().trim())) {
                t = {};
                t["index"] = headIndex++ - 1;
                t["val"] = $(this).text().trim().replace(/\s+/g, '');
                t["col"] = $(this).parent().children().index($(this));
                t["colIndex"] = ++t.index;
                t["excelIndex"] = String.fromCharCode(65 + t.colIndex) + "1";
                t["background-color"] = rgb2hex($(this).css('background-color'));
                t["color"] = rgb2hex($(this).css('color'));
                if (t["background-color"] === t["color"] && t["color"] === "ff000000") {
                    t["background-color"] = 'ffffffff';
                }
                head.push(t);
            }
        });
    } catch (e) {
        console.log("Error at thead reading");
        console.log(e);
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
                t["colIndex"] = String.fromCharCode(65 + t.index%totalBodyCols);
                t["rowIndex"] = 2 + Math.floor(t.index/totalBodyCols);
                t["excelIndex"] = t.colIndex + t.rowIndex;
                t["background-color"] = rgb2hex($(this).css('background-color'));
                t["color"] = rgb2hex($(this).css('color'));
                if (t["background-color"] === t["color"] && t["color"] === "ff000000") {
                    t["background-color"] = 'ffffffff';
                }
                body.push(t);
            }
        });
    } catch (e) {
        console.log("Error at tbody reading");
        console.log(e);
    }

    tableData["head"] = head;
    tableData["body"] = body;
    tableData["totalBodyRows"] = totalBodyRows;
    tableData["totalBodyCols"] = totalBodyCols;

    if (propDetails.consoleLogIteration) console.log(JSON.stringify(tableData));

    headIndex = 0;
    bodyIndex = 0;

    try {
        propDetails.funcBeforeReturn;
    } catch (e) {
        console.log("ERROR: Error at 'funcBeforeReturn'");
        console.log(e);
    }

    return tableData;
}
