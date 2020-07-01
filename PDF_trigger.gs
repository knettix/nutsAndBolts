
function finalSubmit(e) {
    Logger.log(e);
}
function exportCustomePDF() {
    var html = HtmlService.createHtmlOutputFromFile("PDF-model");
    html.setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(html, " ");
}
function getSheetNames() {
    return SpreadsheetApp.getActive()
        .getSheets()
        .map(function (sheet) {
        return sheet.getName();
    });
}
function exportAllSheets() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var pdfSetting = ss.getSheetByName(sheet_PdfSetting);
    if (!pdfSetting) {
        SpreadsheetApp.getUi().alert("No sheet with name :" + sheet_PdfSetting);
        return;
    }
    var configs = pdfSetting
        .getRange(4, 1, pdfSetting.getLastRow(), pdfSetting.getLastColumn())
        .getValues();
    Logger.log("configs");
    Logger.log(configs);
    var allSheets = ss.getSheets().map(function (e) {
        return e.getName();
    });
    for (var i = 0; i < configs.length; i++) {
        var sheetName = configs[i][0] + "";
        var finalSheetName = "";
        allSheets.forEach(function (name) {
            if (finalSheetName == "")
                if (name.toLowerCase().trim().includes(sheetName.toLowerCase().trim())) {
                    finalSheetName = name;
                }
        });
        if (!finalSheetName) {
            continue;
        }
        if (finalSheetName.toLowerCase().trim() ==
            sheet_PdfSetting.toLowerCase().trim())
            continue;
        Logger.log("finalSheetName:" + finalSheetName);
        var sheet = ss.getSheetByName(finalSheetName);
        if (!sheet)
            continue;
        sheet.activate();
        exportCurrentPDF(true, i == configs.length - 1);
    }
}
function ExportSheetFromJ2() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var pdfSetting = ss.getSheetByName(sheet_PdfSetting);
    if (!pdfSetting) {
        SpreadsheetApp.getUi().alert("No sheet with name :" + sheet_PdfSetting);
        return;
    }
    var configs = pdfSetting
        .getRange(4, 1, pdfSetting.getLastRow(), pdfSetting.getLastColumn())
        .getValues();
    Logger.log("configs");
    Logger.log(configs);
    var allSheets = ss.getSheets().map(function (e) {
        return e.getName();
    });
    Logger.log(allSheets);
    var listSheet = ss.getSheetByName(sheet_setup);
    var sheetToExport = listSheet.getRange("J3").getValues()[0][0];
    var finalSheetName = "";
    allSheets.forEach(function (name) {
        if (finalSheetName == "")
            if (name.toLowerCase().trim().includes(sheetToExport.toLowerCase().trim())) {
                finalSheetName = name;
            }
    });
    if (!finalSheetName) {
        return;
    }
    Logger.log("finalSheetName:" + finalSheetName);
    var sheet = ss.getSheetByName(finalSheetName);
    sheet.activate();
    exportCurrentPDF(false, false);
    listSheet.activate();
}
function exportCurrentPDF(exportAll, isLastSheet) {
    var activeSheet = SpreadsheetApp.getActiveSheet().getName();
    var actualName = activeSheet;
    if (activeSheet.includes(".")) {
        activeSheet = activeSheet.substring(activeSheet.indexOf("."));
        activeSheet = activeSheet.replace(".", "");
    }
    Logger.log("activeSheet: " + activeSheet);
    var pdfSetting = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_PdfSetting);
    if (!pdfSetting) {
        SpreadsheetApp.getUi().alert("No sheet with name :" + sheet_PdfSetting);
        return;
    }
    var configs = pdfSetting
        .getRange(4, 1, pdfSetting.getLastRow(), pdfSetting.getLastColumn())
        .getValues();
    var found = false;
    var settingRow = [];
    for (var i = 0; i < configs.length; i++) {
        var sheetName = configs[i][0].toString().trim();
        if (sheetName.includes(".")) {
            sheetName = sheetName.substring(sheetName.indexOf("."));
            sheetName = sheetName.replace(".", "");
        }
        if (activeSheet.trim() == sheetName) {
            found = true;
            settingRow = configs[i];
            break;
        }
    }
    if (!found) {
        return SpreadsheetApp.getUi().alert("No setting found for current Sheet in config Sheet : " + activeSheet);
    }
    Logger.log(settingRow);
    var range = "";
    if (settingRow.length > 1) {
        range = settingRow[1].toString().split(" ");
        range = range[0];
        if (range.includes("(")) {
            range = range.substring(0, range.indexOf("("));
        }
        if (range.includes(":")) {
            var splited = range.split(":");
            if (splited.length >= 2) {
                if (splited[1].trim() == "") {
                    range =
                        range +
                            "" +
                            columnToLetter(SpreadsheetApp.getActiveSheet().getLastColumn());
                }
            }
        }
    }
    var rangeToSkip = "";
    if (settingRow.length > 2) {
        rangeToSkip = settingRow[2];
    }
    var PageFormat = "";
    if (settingRow.length > 3) {
        PageFormat = settingRow[3];
    }
    var Margins = "";
    if (settingRow.length > 4) {
        Margins = settingRow[4];
    }
    var customHeader = false;
    if (settingRow.length > 5) {
        customHeader = settingRow[5].toString().toLowerCase() == "y" ? true : false;
    }
    var reapeatHeaders = false;
    if (settingRow.length > 6) {
        reapeatHeaders =
            settingRow[6].toString().toLowerCase() == "y" ? true : false;
    }
    var _scale = "";
    if (settingRow.length > 7) {
        _scale = settingRow[7];
    }
    var scale = 1;
    switch (_scale) {
        case "Scale 100%":
            scale = 1;
            break;
        case "Fit Width":
            scale = 2;
            break;
        case "Fit Height":
            scale = 3;
            break;
        case "Fit Page":
            scale = 4;
            break;
        default:
            scale: 1;
            break;
    }
    var align = "";
    if (settingRow.length > 8) {
        align = settingRow[8];
    }
    var _subFolder_path = "";
    if (settingRow.length > 9) {
        _subFolder_path = settingRow[9];
    }
    var fileName = "";
    if (settingRow.length > 10) {
        fileName = settingRow[10];
    }
    //   SpreadsheetApp.getUi().alert(`
    //     range:${range}\n
    //     range:${range}\n
    //   `);
    var _parentFolder = pdfSetting.getRange(cell_parentFolder).getValues()[0][0];
    var Ite_parentFolder = DriveApp.getFoldersByName(_parentFolder);
    var parent = null;
    if (Ite_parentFolder.hasNext()) {
        parent = Ite_parentFolder.next();
    }
    else {
        parent = DriveApp.createFolder(_parentFolder);
    }
    _subFolder_path = _subFolder_path.replace("/", "").trim();
    var Ite_parentFolder = DriveApp.getFoldersByName(_parentFolder);
    Ite_parentFolder = parent.getFoldersByName(_subFolder_path);
    var subFolder = null;
    if (Ite_parentFolder.hasNext()) {
        subFolder = Ite_parentFolder.next();
    }
    else {
        subFolder = parent.createFolder(_subFolder_path);
    }
    console.log(subFolder.getUrl());
    exportPdf(sheet_header, actualName, range, {
        range: range,
        rangeToSkip: rangeToSkip,
        PageFormat: PageFormat,
        Margins: Margins,
        fileName: fileName,
        customHeader: customHeader,
        reapeatHeaders: reapeatHeaders,
        scale: scale,
        align: align,
        subFolder: _subFolder_path,
        folderToSaveFile: subFolder.getId()
    }, exportAll, isLastSheet);
}
function addHeader() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var activeSheetName = ss.getActiveSheet().getName();
    var activeSheetName_whole = ss.getActiveSheet().getName();
    if (activeSheetName.includes(".")) {
        activeSheetName = activeSheetName.substring(activeSheetName.indexOf("."));
        activeSheetName = activeSheetName.replace(".", "");
    }
    var headerSheetName = "Header";
    if (activeSheetName.toLowerCase().trim().includes("auto") &&
        activeSheetName.toLowerCase().trim().includes("weights")) {
        headerSheetName = "Header2";
    }
    var sheet_header = ss.getSheetByName(headerSheetName);
    var activeSheet = ss.getSheetByName(activeSheetName_whole);
    var checkCol = 3;
    if (headerSheetName.toLowerCase().trim() == "header2") {
        checkCol = 4;
    }
    //add header to sheet
    //set column width as required
    var pdfSetting = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_PdfSetting);
    if (!pdfSetting) {
        SpreadsheetApp.getUi().alert("No sheet with name :" + sheet_PdfSetting);
        return;
    }
    var configs = pdfSetting
        .getRange(4, 1, pdfSetting.getLastRow(), pdfSetting.getLastColumn())
        .getValues();
    var found = false;
    var settingRow = [];
    for (var i = 0; i < configs.length; i++) {
        var sheetName = configs[i][0].toString().trim();
        if (sheetName.includes(".")) {
            sheetName = sheetName.substring(sheetName.indexOf("."));
            sheetName = sheetName.replace(".", "");
        }
        if (activeSheetName.trim() == sheetName) {
            found = true;
            settingRow = configs[i];
            break;
        }
    }
    if (!found) {
        return SpreadsheetApp.getUi().alert("No setting found for current Sheet in config Sheet : " + activeSheetName);
    }
    Logger.log(settingRow);
    var rangeToExport = "";
    if (settingRow.length > 1) {
        rangeToExport = settingRow[1].toString().split(" ");
        rangeToExport = rangeToExport[0];
        if (rangeToExport.includes("(")) {
            rangeToExport = rangeToExport.substring(0, rangeToExport.indexOf("("));
        }
        if (rangeToExport.includes(":")) {
            var splited = rangeToExport.split(":");
            if (splited.length >= 2) {
                if (splited[1].trim() == "") {
                    rangeToExport =
                        rangeToExport +
                            "" +
                            columnToLetter(SpreadsheetApp.getActiveSheet().getLastColumn());
                }
            }
        }
    }
    var totalCols = 0;
    var totalRows = 0;
    try {
        totalCols = activeSheet.getRange(rangeToExport).getValues()[0].length;
        totalRows = activeSheet.getRange(rangeToExport).getRow() - 1;
        if (totalRows < 1) {
            totalRows = 0;
        }
    }
    catch (_a) { }
    if (totalCols > checkCol) {
        totalCols = totalCols - checkCol;
        if (checkCol - 1 > 0 && totalCols > 0)
            sheet_header.insertColumns(checkCol, totalCols);
    }
    if (totalRows > 0)
        sheet_header.insertRows(1, totalRows);
    rangeToExport = rangeToExport.trim();
    var colsBeforerStart = letterToColumn(rangeToExport[0]) - 1;
    if (colsBeforerStart > 0)
        sheet_header.insertColumns(1, colsBeforerStart);
    SpreadsheetApp.flush();
    var range = sheet_header.getDataRange();
    var values = range.getDisplayValues();
    var _TextDirections = range.getTextDirections();
    var _FontFamilies = range.getFontFamilies();
    var _FontSizes = range.getFontSizes();
    var _FontColors = range.getFontColors();
    var _HorizontalAlignments = range.getHorizontalAlignments();
    var height = [];
    for (var i = 0; i < values.length; i++) {
        height.push(sheet_header.getRowHeight(i + 1));
    }
    var formula = range.getFormulas();
    var mergedRanges = range.getMergedRanges();
    Logger.log(values);
    if (values.length > 0) {
        activeSheet.insertRows(1, values.length);
        activeSheet
            .getRange(1, 1, values.length, activeSheet.getLastColumn())
            .clearContent()
            .clearFormat();
    }
    for (var i = 0; i < values.length; i++) {
        for (var j = 0; j < values[i].length; j++) {
            if (values[i][j]) {
                activeSheet.getRange(i + 1, j + 1, 1, 1).setValues([[values[i][j]]]);
            }
            else {
                activeSheet.getRange(i + 1, j + 1, 1, 1).setFormulas([[formula[i][j]]]);
            }
        }
    }
    var ____range = activeSheet.getRange(rangeToExport);
    var valuesAllTest = activeSheet
        .getRange(1, 1, ____range.getLastRow(), ____range.getLastColumn())
        .getValues();
    var values = ____range.getValues();
    var to = values.length;
    var checkFrom = values.length - 1;
    if (sheet_data.toLowerCase().indexOf("multi") >= 0 &&
        sheet_data.toLowerCase().indexOf("patch") >= 0) {
        Logger.log("multi touch");
        checkFrom = 900;
    }
    else {
        Logger.log("not multi touch");
    }
    Logger.log("checkFrom:" + checkFrom);
    for (var i = checkFrom; i >= 0; i--) {
        if (valuesAllTest[i][0] != "") {
            to = i;
            break;
        }
    }
    to++;
    if (sheet_data.includes("4.")) {
        to++;
    }
    var _rangeToExport = rangeToExport;
    if (rangeToExport.split(":").length == 2) {
        if (!isContainsNumber(rangeToExport.split(":")[1])) {
            _rangeToExport = rangeToExport + (to + range.getNumRows() + totalRows);
        }
        else {
            var tempRange = activeSheet.getRange(_rangeToExport);
            _rangeToExport = activeSheet
                .getRange(tempRange.getRow(), tempRange.getColumn(), tempRange.getNumRows() + totalRows + range.getNumRows(), tempRange.getNumColumns())
                .getA1Notation();
        }
    }
    for (var i = 0; i < mergedRanges.length; i++) {
        Logger.log(mergedRanges[i].getRow() + " ," + mergedRanges[i].getColumn() + " , " + mergedRanges[i].getLastRow() + " , " + mergedRanges[i].getLastColumn() + " ");
        activeSheet
            .getRange(mergedRanges[i].getRow(), mergedRanges[i].getColumn(), mergedRanges[i].getNumRows(), mergedRanges[i].getNumColumns())
            .merge();
    }
    for (var i = 0; i < height.length; i++) {
        activeSheet.setRowHeight(i + 1, height[i]);
    }
    var headerRange = activeSheet.getRange(range.getRow(), range.getColumn(), range.getNumRows(), range.getNumColumns());
    headerRange.setTextDirections(_TextDirections);
    headerRange.setFontFamilies(_FontFamilies);
    headerRange.setFontSizes(_FontSizes);
    headerRange.setFontColors(_FontColors);
    headerRange.setHorizontalAlignments(_HorizontalAlignments);
    if (totalRows > 0)
        sheet_header.deleteRows(1, totalRows);
    if (colsBeforerStart > 0)
        sheet_header.deleteColumns(1, colsBeforerStart);
    if (totalCols > 0) {
        if (checkCol > 0)
            sheet_header.deleteColumns(checkCol, totalCols);
    }
}
function deleteHeader() {
    var activeSheet = SpreadsheetApp.getActiveSheet().getName();
    SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(activeSheet)
        .deleteRows(1, 3);
}
function sendFolderAsMail() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var pdfSetting = ss.getSheetByName(sheet_PdfSetting);
    if (!pdfSetting) {
        SpreadsheetApp.getUi().alert("No sheet with name :" + sheet_PdfSetting);
        return;
    }
    var folderName = pdfSetting.getRange(cell_parentFolder).getValues()[0][0];
    var root = DriveApp.getFoldersByName(folderName);
    if (!root.hasNext()) {
        SpreadsheetApp.getUi().alert("No Folder founs with name: " + folderName);
        return;
    }
    var rootFolder = root.next();
    if (!rootFolder) {
        SpreadsheetApp.getUi().alert("No Folder founs with name: " + folderName);
        return;
    }
    var table = [];
    var subFoldersIt = rootFolder.getFolders();
    while (subFoldersIt.hasNext()) {
        var subFold = subFoldersIt.next();
        var filesIte = subFold.getFiles();
        table.push([subFold.getName(), subFold.getUrl()]);
        // while (filesIte.hasNext()) {
        //   var file = filesIte.next();
        // }
    }
    table = table.sort();
    table.push(["-", "-"]);
    table.push([rootFolder.getName(), rootFolder.getUrl()]);
    var emialListSheet = ss.getSheetByName(sheet_setup);
    var beforeText = emialListSheet.getRange(range_text_before).getValues()[0][0];
    var afterText = emialListSheet.getRange(range_text_after).getValues()[0][0];
    var emails = emialListSheet
        .getRange(range_Mails)
        .getValues()
        .map(function (e) {
        return e[0];
    });
    var htmlBody = makeHtmlBody(beforeText, afterText, table);
    Logger.log("emails");
    Logger.log(emails);
    emails.forEach(function (email) {
        try {
            if (email)
                __sendEails(email, htmlBody);
        }
        catch (e) { }
    });
}
function makeHtmlBody(before, after, data) {
    var html = "";
    html += before;
    html += "<br>";
    html += "<br>";
    html += "<table>";
    html += "<thead>";
    html += "<tr>";
    html += "<th>File Name</th>";
    html += "<th>Drive Link</th>";
    html += "</tr>";
    html += "</thead>";
    html += "<tbody>";
    data.forEach(function (ele) {
        html += "<tr>";
        for (var i = 0; i < ele.length; i++) {
            html += "<td>" + ele[i] + "</td>";
        }
        html += "</tr>";
    });
    html += "</thead>";
    html += "</table>";
    html += "<br>";
    html += after;
    return html;
}
function __sendEails(email, htmlBody) {
    GmailApp.sendEmail(email, "Files", htmlBody, {
        htmlBody: htmlBody
    });
}
