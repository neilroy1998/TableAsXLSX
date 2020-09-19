let t0 = 0;
let t1 = 0;
let ton = function () {
    t0 = performance.now();
}
let toff = function () {
    t1 = performance.now();
    console.log("TIME:: " + (t1 - t0) + "ms");
}

$(function () {
    $("#btn1").click(function () {
        let tp = web2xlsx(
            {
                "fileName": "TestBook",
                "initFuncAtStart": ton(),
                "initFuncAtEnd": toff()
            },
            [{
                "tableID": "#table1",
                "colExclude": ['1'],
                "rowExclude": ['1'],
                "funcAtStart": null,
                "funcAtEnd": null,
                "consoleLogIteration": false,
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
                },
                {
                    "tableID": ".exportDemandTable",
                    "defaultWidth": 15,
                    "customWidth": [{col: 25, width: 50}],
                    "sheetDetails": {
                        "sheetName": "MR",
                        "sheetTabColor": "#00FF00"
                    }
                }]);

    });
});
