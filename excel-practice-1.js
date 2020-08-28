$(document).ready(function () {

    $('#btn2').click(function () {
        excelTest();
    });
});

let excelTest = function (iteratedValue) {

    let wb = new ExcelJS.Workbook();
    let ws = wb.addWorksheet("Sheet1");

    let headArray = iteratedValue.head;
    let bodyArray = iteratedValue.body;

    let columnConfig = [];
    let colKeyIndex = 65;

    /* todo ask for col widths
        default: 8.11 or 9
        wrap*/
    headArray.forEach(function (h) {
        temp = {
            header: h.val,
            key: String.fromCharCode(colKeyIndex++),
            width: 9
        };
        columnConfig.push(temp);
    });

    ws.columns = columnConfig;

    bodyArray.forEach(function (v) {

        if (v.val === "") v.val = "";
        else if(!isNaN(v.val.replace(/,/g, ""))) ws.getCell(v.excelIndex).value = parseInt(v.val);
        else ws.getCell(v.excelIndex).value = v.val;

        if (v["background-color"]!=="ffffffff") {
            ws.getCell(v.excelIndex).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: {argb: v["background-color"]}
            };
        }
        ws.getCell(v.excelIndex).font = {
            color: {argb: v.color}
        };
        ws.getCell(v.excelIndex).alignment = { wrapText: true };
    });

    headArray.forEach(function (h) {
        if (h["background-color"]!=="ffffffff") {
            ws.getCell(h.excelIndex).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: {argb: h["background-color"]}
            };
        }
        ws.getCell(h.excelIndex).font = {
            color: {argb: h.color}
        };
        ws.getCell(h.excelIndex).alignment = { wrapText: true };
    });

    headArray.concat(bodyArray).forEach(function (i) {
        ws.getCell(i.excelIndex).border = {
            top: {style:'thin', color: {argb:'FFD4D4D4'}},
            left: {style:'thin', color: {argb:'FFD4D4D4'}},
            bottom: {style:'thin', color: {argb:'FFD4D4D4'}},
            right: {style:'thin', color: {argb:'FFD4D4D4'}}
        };
    });

// Using window.saveAs rather than fileSaver.saveAs
    wb.xlsx.writeBuffer()
        .then(buffer => window.saveAs(new Blob([buffer]), `Testing_book.xlsx`))
        .catch(err => console.log('Error writing excel export', err));
}