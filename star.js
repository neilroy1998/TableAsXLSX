let tableIterate = function (tableID, propDetails) {

    propDetails.funcAtStart;

    let tableData = {};
    let defaultClasses = ["background-color", "color"];

    defaultClasses = defaultClasses.concat(propDetails.styleInclude);
    defaultClasses = defaultClasses.filter(function (value) {
        return (!propDetails.styleExclude.includes(value));
    });

    let head = [];
    let body = [];

    $(tableID + " > thead > tr > th").each(function () {
        t = {};
        t["val"] = $(this).text().trim().replace(/\s+/g, '');
        t["col"]= $(this).parent().children().index($(this));
        for (var i = 0; i < defaultClasses.length; i++) {
            t[defaultClasses[i]] = $(this).css(defaultClasses[i]);
        }
        if (!propDetails.colExclude.includes(t["col"].toString().trim())) head.push(t);
    });

    $(tableID + " > tbody > tr > td").each(function () {
        t = {};
        t["val"] = $(this).text().trim().replace(/\s+/g, '');
        t["col"]= $(this).parent().children().index($(this));
        t["row"] = $(this).parent().parent().children().index($(this).parent());
        for (var i = 0; i < defaultClasses.length; i++) {
            t[defaultClasses[i]] = $(this).css(defaultClasses[i]);
        }
        if (!propDetails.colExclude.includes(t["col"].toString().trim()) &&
            !propDetails.rowExclude.includes(t["row"].toString().trim())) body.push(t);
    });

    tableData["head"] = head;
    tableData["body"] = body;

    if (propDetails.consoleLogIteration) console.log(JSON.stringify(tableData));

    propDetails.funcBeforeReturn;

    return tableData;
}
