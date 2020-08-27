$(function () {
    $("#btn1").click(function () {
        let tp = tableIterate("#table1", {
            "useHexColor": false, 
            "colExclude": [],
            "styleInclude": [],
            "rowExclude": [],
            "styleExclude": [],
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
    });
});