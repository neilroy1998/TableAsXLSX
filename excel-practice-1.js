$(document).ready(function () {

    $('#btn2').click(function () {
        excelTest();
    });
});

let excelTest = function (iteratedValue) {

    let rowValues = [];
    let colValues = [];

    return;

    let wb = new ExcelJS.Workbook();
    let ws = wb.addWorksheet("Sheet1");

    let totalRows = 3;
    let totalCols = 3;
    let val = 1;
    for (var r = 65; r < 65+totalRows; r++) {
        for (var c = 1; c < 1+totalCols; c++) {
            ws.getCell(String.fromCharCode(r) + c).value = val++;
        }
    }

    // Using window.saveAs rather than fileSaver.saveAs
    wb.xlsx.writeBuffer()
        .then(buffer => window.saveAs(new Blob([buffer]), `Testing_book.xlsx`))
        .catch(err => console.log('Error writing excel export', err));
}