/**
 * Export any HTML table to .xlsx format
 * @author Rajit Roy
 * @updated 19th Sept, 2020
 * @licence MIT
 * Requires jQuery, FileSaver, ExcelJS (in that order: before this script)
 */

/**
 * @license
 * Copyright (c) 2020 Rajit Roy
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

// Start of Script

/**
 * @param elem - HTML Object to find it's background color
 * @returns {string|string} - the background color in rgb
 * Gets the background color of element. If transparent: gets the background color of it's parent recursively
 */
function realBackgroundColor(elem) {

	/**
	 * @type {string}
	 * the default value returned by a browser if no background-color property found
	 */
	var transparent = 'rgba(0, 0, 0, 0)';
	var transparentIE11 = 'transparent';
	if (!elem) return 'rgb(255, 255, 255)';
	
	var bg = getComputedStyle(elem).backgroundColor;
	
	/**
	 * if none, find background color of it's parent recursively
	 */
	if (bg === transparent || bg === transparentIE11) {
		if (elem.nodeName === "HTML") return 'rgb(255, 255, 255)';
		return realBackgroundColor(elem.parentElement);
	} else {
		return bg;
	}
}

/**
 * @param input - the rgb as string
 * @returns {string|*} - the argb value
 * Converts rgb to argb (without # prefix)
 */
function rgb2argb(input) {
	var rgb = input;
	try {
		
		/**
		 * check if in hex format
		 */
		if (/^#[0-9A-F]{6}$/i.test(rgb)) return rgb;
		
		/**
		 * should be supporting rgba
		 */
		rgb = rgb.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*(\d+))?\)$/);
		
		/**
		 * @param x
		 * @returns {string}
		 * Gets in hex string format
		 */
		function hex(x) {
			return ("0" + parseInt(x).toString(16)).slice(-2);
		}
		
		/**
		 * put ff in the front to make it argb
		 */
		return "ff" + hex(rgb[1]) + hex(rgb[2]) + hex(rgb[3]);
	} catch (e) {
		
		/**
		 * in case not a color value
		 */
		return input;
	}
}

/**
 * @param UNIX_timestamp - input UNIX datetime
 * @returns {string} - date as readable format
 * Converts UNIX datetime to readable DD MMM, YYYY format
 */
function timeConverter(UNIX_timestamp) {
	
	var a = new Date(UNIX_timestamp);
	var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
	
	var year = a.getFullYear();
	var monthNumber = a.getMonth();
	var month = months[monthNumber];
	var date = a.getDate();
	var properDate = "";
	
	if (date < 10) properDate = month + " 0" + date + ", " + year;
	else properDate = month + " " + date + ", " + year;
	
	return properDate;
}

/**
 * @param num - any integer from 1
 * @returns {*} - string count
 * Returns AA (like in Excel) after Z and so on
 */
function numberCounter(num) {
	"use strict";
	
	/**
	 * @type {number}
	 * Get mod and append suitable alphabet and return it
	 */
	var mod = num % 26,
		pow = num / 26 | 0,
		out = mod ? String.fromCharCode(64 + mod) : (--pow, 'Z');
	return pow ? numberCounter(pow) + out : out;
}

/**
 * @param initDetails - Excel Initialization details
 * @param tableData - Data of Each individual table
 * Initial function as called by user. Starts all the functions and then serves the final Excel file
 */
var web2xlsx = function (initDetails, tableData) {
	try {
		initDetails.funcAtStart;
	} catch (e) {
		console.error("ERROR: Error at 'initFuncAtStart'");
		console.error(e);
	}
	
	initDetails = {
		
		/**
		 * Name of Excel file
		 */
		"fileName": initDetails.fileName || "Excel Export on " + timeConverter(Date.now()),
		
		/**
		 * Function to run at beginning of it all
		 * Example: let foo = function (bar) {}... The pass variable foo to this
		 */
		"initFuncAtStart": initDetails.initFuncAtStart || null,
		
		/**
		 * Function to run at the end of it all
		 * Example: let foo = function (bar) {}... The pass variable foo to this
		 */
		"initFuncAtEnd": initDetails.initFuncAtEnd || null,
	}
	
	tableData = tableData || null;
	
	if (tableData) {
		
		/**
		 * @type {ExcelJS.Workbook}
		 * Create new workbook
		 */
		var wb = new ExcelJS.Workbook();
		
		/**
		 * Default Values for each table
		 */
		for (var i = 0; i < tableData.length; i++) {
			
			tableData[i] = {
				
				/**
				 * Table Element to be converted to Excel sheet
				 */
				"tableID": tableData[i].tableID,
				
				/**
				 * Exclude certain columns.
				 * Count from '0'. Example (3rd column or column C in Excel) ['2']
				 */
				"colExclude": tableData[i].colExclude || [],
				
				/**
				 * Exclude certain rows of body.
				 * Count from '0'. Example (3rd row) ['2']
				 */
				"rowExclude": tableData[i].rowExclude || [],
				
				/**
				 * Function to run at beginning of this table's operation
				 * Example: let foo = function (bar) {}... The pass variable foo to this
				 */
				"funcAtStart": tableData[i].funcAtStart || null,
				
				/**
				 * Function to run at end of this table's operation
				 * Example: let foo = function (bar) {}... The pass variable foo to this
				 */
				"funcAtEnd": tableData[i].funcAtEnd || null,
				
				/**
				 * Log in browser console each iterated value
				 */
				"consoleLogIteration": tableData[i].consoleLogIteration || false,
				
				/**
				 * Default Width of columns
				 */
				"defaultWidth": tableData[i].defaultWidth || 9,
				
				/**
				 * JSON Details of sheet
				 * sheetName: Name of the sheet
				 * sheetTabColor: Background Color of sheet tab (at bottom of Excel)
				 */
				"sheetDetails": tableData[i].sheetDetails || {"sheetName": null, "sheetTabColor": null},
				
				/**
				 * JSON Array to set custom width (count starts at 0)
				 * Example: Column 2 will have width 50 and Column 5 will have width 30
				 * So pass [{col: 1, width: 50}, {col: 4, width: 30}]
				 */
				"customWidth": tableData[i].customWidth || [{col: null, width: null}],
				
				/**
				 * Wrap Values in Excel boolean
				 */
				"wrap": tableData[i].wrap || false
			}
			
			tableIterator(wb, tableData[i], i);
		}
		
		/**
		 * Write entire workbook and all sheets within and save as an Excel file
		 */
		// Using window.saveAs rather than fileSaver.saveAs
		wb.xlsx.writeBuffer()
			.then(buffer => window.saveAs(new Blob([buffer]), `${initDetails.fileName}.xlsx`))
			.catch(err => console.error('Error writing excel export', err));
		
	}
	
	try {
		initDetails.initFuncAtEnd;
	} catch (e) {
		console.error("ERROR: Error at 'initFuncAtEnd'");
		console.error(e);
	}
}

/**
 * @param elem - $(this)
 * @param t - JSON with all values
 * @returns t
 * Get colors and formatting values of each cell in HTML table
 */
function getStyles(elem, t) {
	t["background-color"] = rgb2argb(realBackgroundColor(elem[0]));
	t["color"] = rgb2argb(elem.css('color'));
	
	/**
	 * @type {{size: number, underline: boolean, italics: boolean, bold: boolean, family: string}}
	 * Font styles
	 */
	t["style"] = {
		"bold": false,
		"italics": false,
		"underline": false,
		"size": 11,
		"family": 'Calibri'
	};
	
	/**
	 * Bold
	 */
	if (parseInt(elem.css('font-weight')) >= 700) t.style.bold = true;
	
	/**
	 * Italics
	 */
	if (elem.css('font-style').includes("italic")) t.style.italics = true;
	
	/**
	 * Font size
	 */
	try {
		t.style.size = elem.css('font-size').split("px")[0];
	} catch (e) {
		t.style.size = 11;
	}
	
	/**
	 * Underlined
	 */
	if (elem.css("text-decoration").split(" ").includes("underline")) t.style.underline = true;
	
	/**
	 * Get font style (family)
	 */
	try {
		t.style.family = elem.css("font-family").split(',')[0].replace(/['"]/g, "");
	} catch (e) {
		t.style.family = 'sans-serif';
	}
	
	return t;
}

/**
 * @param wb - the workbook created by web2xlsx
 * @param tableDetails - Data of a particular table as sent by web2xlsx
 * @param tableIndex - Index # of that table
 * Goes through given table and returns colors, font styles, values etc. for each cell
 */
function tableIterator(wb, tableDetails, tableIndex) {
	
	try {
		tableDetails.funcAtStart;
	} catch (e) {
		console.error("ERROR: Error at 'funcAtStart'");
		console.error(e);
	}
	
	var tableData = {};
	
	/**
	 * @type {*[]}
	 * For thead, tbody, tfoot
	 */
	var head = [];
	var body = [];
	var foot = [];
	
	var headIndex = 0;
	var bodyIndex = 0;
	var footIndex = 0;
	
	/* todo
	*   readme*/
	
	var totalBodyRows = $(tableDetails.tableID + " > tbody > tr").length - tableDetails.rowExclude.length;
	var totalBodyCols = $(tableDetails.tableID + " > tbody > tr:eq(0) > td").length - tableDetails.colExclude.length;
	
	try {
		
		/**
		 * Traverse through cells of thead
		 */
		$(tableDetails.tableID + " > thead > tr > th").each(function () {
			
			if (!tableDetails.colExclude.includes($(this).parent().children().index($(this)).toString().trim())) {
				
				var t = {};
				
				t["index"] = headIndex++ - 1;
				t["val"] = $(this).text().trim();
				
				/**
				 * @type {jQuery|number}
				 * Column number
				 */
				t["col"] = $(this).parent().children().index($(this));
				
				/**
				 * @type {number}
				 * Column Index for Excel
				 */
				t["colIndex"] = ++t.index;
				
				t["colWidth"] = tableDetails.defaultWidth;
				tableDetails.customWidth.forEach(function (cw) {
					if (t.index === cw.col) t["colWidth"] = cw.width;
				})
				
				/**
				 * @type {string}
				 * Position in Excel sheet (Example: B3 or AA2)
				 */
				t["excelIndex"] = numberCounter(1 + t.colIndex) + "1";
				
				/**
				 * Get colors and formatting options
				 */
				t = getStyles($(this), t);
				
				head.push(t);
			}
		});
	} catch (e) {
		console.error("Error at thead reading");
		console.error(e);
	}
	
	try {
		
		/**
		 * Traverse through cells of tbody
		 */
		$(tableDetails.tableID + " > tbody > tr > td").each(function () {
			
			if (!tableDetails.colExclude.includes($(this).parent().children().index($(this)).toString().trim()) &&
				!tableDetails.rowExclude.includes($(this).parent().parent().children().index($(this).parent()).toString().trim())) {
				
				var t = {};
				
				t["index"] = bodyIndex++;
				t["val"] = $(this).text().trim();
				
				/**
				 * @type {jQuery|number}
				 * Column number
				 */
				t["col"] = $(this).parent().children().index($(this));
				
				/**
				 * @type {jQuery|number}
				 * Row number
				 */
				t["row"] = $(this).parent().parent().children().index($(this).parent());
				
				/**
				 * @type {*}
				 * Column Index for Excel
				 */
				t["colIndex"] = numberCounter(1 + (t.index % totalBodyCols));
				
				/**
				 * @type {number}
				 * Row Index for Excel
				 */
				t["rowIndex"] = 2 + Math.floor(t.index / totalBodyCols);
				
				/**
				 * @type {string}
				 * Position in Excel sheet (Example: B3 or AA2)
				 */
				t["excelIndex"] = t.colIndex + t.rowIndex;
				
				/**
				 * Get colors and formatting options
				 */
				t = getStyles($(this), t);
				
				body.push(t);
			}
		});
	} catch (e) {
		console.error("Error at tbody reading");
		console.error(e);
	}
	
	try {
		
		/**
		 * Traverse through cells of tfoot
		 */
		$(tableDetails.tableID + " > tfoot > tr > td").each(function () {
			
			if (!tableDetails.colExclude.includes($(this).parent().children().index($(this)).toString().trim())) {
				
				var t = {};
				
				t["index"] = footIndex++ - 1;
				t["val"] = $(this).text().trim();
				
				/**
				 * @type {jQuery|number}
				 * Column number
				 */
				t["col"] = $(this).parent().children().index($(this));
				
				/**
				 * @type {number}
				 * Column Index for Excel
				 */
				t["colIndex"] = ++t.index;
				
				/**
				 * @type {string}
				 * Position in Excel sheet (Example: B3 or AA2)
				 */
				t["excelIndex"] = numberCounter(1 + t.colIndex) + (totalBodyRows + 2);
				
				/**
				 * Get colors and formatting options
				 */
				t = getStyles($(this), t);
				
				foot.push(t);
			}
		});
	} catch (e) {
		console.error("Error at tfoot reading");
		console.error(e);
	}
	
	tableData["head"] = head;
	tableData["body"] = body;
	tableData["foot"] = foot;
	tableData["totalBodyRows"] = totalBodyRows;
	tableData["totalBodyCols"] = totalBodyCols;
	
	headIndex = 0;
	bodyIndex = 0;
	footIndex = 0;
	
	if (tableDetails.consoleLogIteration) console.log(JSON.stringify(tableData));
	
	excelCreateAndExport(wb, tableData, tableDetails, tableIndex);
	
}

/**
 * @param wb - the workbook created by web2xlsx
 * @param iteratedValue - The data from function tableIterator for each table passed to it
 * @param tableDetails - Data of a particular table as sent by web2xlsx
 * @param tableIndex - Index # of that table
 * Creates Excel workbook and indivisual sheets and fills them according to result from function tableIterator
 */
function excelCreateAndExport(wb, iteratedValue, tableDetails, tableIndex) {
	
	var sheetName;
	if (tableDetails.sheetDetails.sheetName) sheetName = tableDetails.sheetDetails.sheetName;
	else sheetName = "Sheet" + (tableIndex + 1).toString();
	
	/**
	 * Get Sheet Tab Color in argb (at the bottom of Ms Excel)
	 */
	var sheetTabColor;
	if (tableDetails.sheetDetails.sheetTabColor) sheetTabColor = "FF" + tableDetails.sheetDetails.sheetTabColor.split("#")[1];
	else sheetTabColor = "FFFFFFFF";
	
	var ws = wb.addWorksheet(sheetName,
		{
			properties: {tabColor: {argb: sheetTabColor}},
			
			/**
			 * First Row (head) frozen
			 */
			views: [{state: 'frozen', ySplit: 1}]
		});
	
	var headArray = iteratedValue.head;
	var bodyArray = iteratedValue.body;
	var footArray = iteratedValue.foot;
	
	/**
	 * @type {*[]}
	 * Settings for each column
	 */
	var columnConfig = [];
	var colKeyIndex = 1;
	
	if (tableDetails.defaultWidth <= 0) tableDetails.defaultWidth = 9;
	
	/**
	 * Settings for each column to be set
	 */
	headArray.forEach(function (h) {
		var temp = {
			header: h.val,
			key: numberCounter(colKeyIndex++),
			width: h.colWidth || 9
		};
		columnConfig.push(temp);
	});
	
	/**
	 * @type {*[]}
	 * Push these column settings to the worksheet
	 */
	ws.columns = columnConfig;
	
	var addBackgroundColor = function (v) {
		return {
			type: 'pattern',
			pattern: 'solid',
			fgColor: {argb: v["background-color"]}
		};
	}
	
	var addFontStyle = function (v) {
		return {
			color: {argb: v.color},
			name: v.style.family,
			family: 2,
			bold: v.style.bold,
			underline: v.style.underline,
			italic: v.style.italics
		};
	}
	
	var addStyles = function (ws, v) {
		if (v["background-color"] !== "ffffffff") ws.getCell(v.excelIndex).fill = addBackgroundColor(v);
		ws.getCell(v.excelIndex).font = addFontStyle(v);
		if (tableDetails.wrap) ws.getCell(v.excelIndex).alignment = {vertical: 'middle', wrapText: true};
		else ws.getCell(v.excelIndex).alignment = {vertical: 'middle'};
		return ws;
	}
	
	var addValues = function (ws, v) {
		if (v.val === "") v.val = "";
		else if (!isNaN(v.val.replace(/,/g, ""))) ws.getCell(v.excelIndex).value = parseInt(v.val);
		else ws.getCell(v.excelIndex).value = v.val;
		return ws;
	}
	
	/**
	 * Formatting and Values for each body, head and foot cell
	 */
	bodyArray.forEach(function (v) {
		ws = addValues(ws, v);
		ws = addStyles(ws, v);
	});
	headArray.forEach(function (h) {
		ws = addStyles(ws, h);
	});
	footArray.forEach(function (f) {
		ws = addValues(ws, f);
		ws = addStyles(ws, f);
	});
	
	/**
	 * Set the default Excel thin gray border to each cell
	 * Increases Readability
	 */
	headArray.concat(bodyArray).concat(footArray).forEach(function (i) {
		ws.getCell(i.excelIndex).border = {
			top: {style: 'thin', color: {argb: 'FFD4D4D4'}},
			left: {style: 'thin', color: {argb: 'FFD4D4D4'}},
			bottom: {style: 'thin', color: {argb: 'FFD4D4D4'}},
			right: {style: 'thin', color: {argb: 'FFD4D4D4'}}
		};
	});
	
	try {
		tableDetails.funcAtEnd;
	} catch (e) {
		console.error("ERROR: Error at 'funcAtEnd'");
		console.error(e);
	}
}

// End of Script