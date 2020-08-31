$(function () {
    $("#btn1").click(function () {
        let tp = excelize("#table1", {
            "colExclude": ['1'],
            "rowExclude": ['1'],
            "funcAtStart": null,
            "funcAtEnd": null,
            "consoleLogIteration": false,
            "defaultWidth": 25,
            "fileName": "TestBook",
            "header": "TEST-header"
        });

    });
});