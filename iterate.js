$(function () {
    $("#btn1").click(function () {
        let tp = tableIterate("#table1", {
            "colExclude": ['1'],
            "styleInclude": [],
            "rowExclude": ['1'],
            "funcAtStart": null,
            "funcBeforeReturn": null,
            "consoleLogIteration": false,
            "removeColors": false
        });

        $("#result").val(
            "Head\n" +
            JSON.stringify(tp.head).replace(/},{/g, ",\n") +
            "\n\nBody\n" +
            JSON.stringify(tp.body).replace(/},{/g, ",\n")
        );

        excelTest(tp);
    });
});