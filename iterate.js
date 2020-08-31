$(function () {
    $("#btn1").click(function () {
        let tp = excelize(
            {
                "fileName": "TestBook",
                "initFuncAtStart": null,
                "initFuncAtEnd": null
            },
            [{
                "tableID": "#table1",
                "colExclude": ['1'],
                "rowExclude": ['1'],
                "funcAtStart": null,
                "funcAtEnd": null,
                "consoleLogIteration": false,
                "defaultWidth": 25,
                "sheetName": "Tab1",
                "sheetTabColor": "#FF0000"
            }, {
                "tableID": "#table2",
                "defaultWidth": 15
            }]);

    });
});