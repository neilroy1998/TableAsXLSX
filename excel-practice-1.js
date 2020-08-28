$(document).ready(function () {

    $('#btn2').click(function () {
        excelTest();
    });
});

let excelTest = function (iteratedValue) {

    let wb = new ExcelJS.Workbook();
    let ws = wb.addWorksheet("Sheet1");

    let totalBodyRows = iteratedValue.totalBodyRows;
    let totalBodyCols = iteratedValue.totalBodyCols;
    let totalRows = totalBodyRows + 1;

    let headArray = iteratedValue.head;
    let bodyArray = iteratedValue.body;
    let completeArray = headArray.concat(bodyArray);
    completeArray.forEach(v => ws.getCell(v.excelIndex).value = v.val);

    // Using window.saveAs rather than fileSaver.saveAs
    wb.xlsx.writeBuffer()
        .then(buffer => window.saveAs(new Blob([buffer]), `Testing_book.xlsx`))
        .catch(err => console.log('Error writing excel export', err));
}