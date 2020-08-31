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
                "consoleLogIteration": true,
                "defaultWidth": 25,
                "sheetDetails": {
                    "sheetName": "Tab1",
                    "sheetTabColor": "#FF0000"
                },
                "customWidth": [{col: 0, width: 60}]
            },
                {
                    "tableID": "#table2",
                    "defaultWidth": 15,
                    "customWidth": [{
                        col: 1,
                        width: 50
                    }, {
                        col: 2,
                        width: 5
                    }]
                }]);

    });
});