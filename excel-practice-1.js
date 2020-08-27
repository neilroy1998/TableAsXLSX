$(document).ready(function() {

    $('#btn2').click(function () {
        let wb = new ExcelJS.Workbook();
        let ws = wb.addWorksheet("Sheet1");

        ws.getCell('A1').value = 'Hello, World!';
        ws.getCell('A2').value = 7;

        // Using window.saveAs rather than fileSaver.saveAs
        wb.xlsx.writeBuffer()
            .then(buffer => window.saveAs(new Blob([buffer]), `Testing_book.xlsx`))
            .catch(err => console.log('Error writing excel export', err));
    });
});