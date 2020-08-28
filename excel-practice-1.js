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
    let completeArray = headArray.concat(bodyArray);

    let columnConfig = [];
    let colKeyIndex = 65;

    /* todo ask for col widths
        default: 8.11 or 9 */
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
        ws.getCell(v.excelIndex).value = v.val;
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
    });

// Using window.saveAs rather than fileSaver.saveAs
    wb.xlsx.writeBuffer()
        .then(buffer => window.saveAs(new Blob([buffer]), `Testing_book.xlsx`))
        .catch(err => console.log('Error writing excel export', err));
}