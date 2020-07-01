// Compiled using ts2gas 3.6.2 (TypeScript 3.9.3)
function getColFromWhereToStart(range) {
    var from = range.split(":")[0];
    return letterToColumn(from.replace(/[0-9]/g, "").trim());
}
function isContainsNumber(dataString) {
    for (var i = 0; i < dataString.length; i++) {
        if (!isNaN(Number(dataString[i]))) {
            return true;
        }
    }
    return false;
}
function exportPdf(sheet_header, sheet_data, rangeToExport, config, allSheetAtOnce, isLastSheet) {
    Logger.log("--args");
    Logger.log(sheet_header);
    Logger.log(sheet_data);
    Logger.log(rangeToExport);
    Logger.log(config);
    Logger.log("--args");
    var startColFrom = 1; // getColFromWhereToStart(rangeToExport);
    Logger.log("startColFrom:" + startColFrom);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getSheetByName(sheet_data);
    if (sheet_data.toLowerCase().trim().includes("6") &&
        sheet_data.toLowerCase().trim().includes("auto") &&
        sheet_data.toLowerCase().trim().includes("weights")) {
        sheet_header = "Header2";
    }
    var ____range = dataSheet.getRange(rangeToExport);
    var valuesAllTest = dataSheet
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
    var frozenRows = dataSheet.getFrozenRows();
    if (!config.customHeader) {
        // finalSheet = ss.insertSheet().setName("temp-" + new Date().getTime());
        if (config.rangeToSkip.trim() != "") {
            try {
                var hideRanges = config.rangeToSkip.split(",");
                hideRanges.forEach(function (__range) {
                    __range = __range.trim();
                    if (__range != "") {
                        try {
                            dataSheet.hideRow(dataSheet.getRange(__range));
                        }
                        catch (e) {
                            Logger.log(e);
                        }
                    }
                });
            }
            catch (e) {
                Logger.log(e);
            }
        }
        var _rangeToExport = rangeToExport;
        if (rangeToExport.split(":").length == 2) {
            if (!isContainsNumber(rangeToExport.split(":")[1]))
                _rangeToExport = rangeToExport + to;
        }
        var fileUrl = generatePDf(dataSheet, _rangeToExport, config.folderToSaveFile, config.fileName, config.reapeatHeaders, config.PageFormat.toLowerCase() == "portrait", config.scale, config.align, config.Margins);
        if (config.rangeToSkip.trim() != "") {
            try {
                var hideRanges = config.rangeToSkip.split(",");
                hideRanges.forEach(function (__range) {
                    __range = __range.trim();
                    if (__range != "") {
                        try {
                            dataSheet.unhideRow(dataSheet.getRange(__range));
                        }
                        catch (e) {
                            Logger.log(e);
                        }
                    }
                });
            }
            catch (e) {
                Logger.log(e);
            }
        }
        var template = HtmlService.createTemplateFromFile("PDF_complete.html");
        template.url = fileUrl;
        if (!allSheetAtOnce) {
            SpreadsheetApp.getUi().showModalDialog(template.evaluate(), "PDF Export Complete");
        }
        if (isLastSheet) {
            SpreadsheetApp.getUi().alert("All Sheet Exported");
        }
    }
    else {
        var checkCol = 3;
        if (sheet_header.toLowerCase().trim() == "header2") {
            checkCol = 4;
        }
        //add header to sheet
        //set column width as required
        var totalCols = 0;
        var totalRows = 0;
        try {
            totalCols = dataSheet.getRange(rangeToExport).getValues()[0].length;
            totalRows = dataSheet.getRange(rangeToExport).getRow() - 1;
            if (totalRows < 1) {
                totalRows = 0;
            }
        }
        catch (_a) { }
        var header = ss.getSheetByName(sheet_header);
        if (totalCols > checkCol) {
            totalCols = totalCols - checkCol;
            if (totalCols > 0 && checkCol > 0) {
                header.insertColumns(checkCol, totalCols);
            }
        }
        if (totalRows > 0) {
            header.insertRows(1, totalRows);
        }
        rangeToExport = rangeToExport.trim();
        var colsBeforerStart = letterToColumn(rangeToExport[0]) - 1;
        if (colsBeforerStart > 0) {
            header.insertColumns(1, colsBeforerStart);
        }
        SpreadsheetApp.flush();
        var range = header.getDataRange();
        var values = range.getDisplayValues();
        var _TextDirections = range.getTextDirections();
        var _FontFamilies = range.getFontFamilies();
        var _FontSizes = range.getFontSizes();
        var _FontColors = range.getFontColors();
        var _HorizontalAlignments = range.getHorizontalAlignments();
        var height = [];
        for (var i = 0; i < values.length; i++) {
            height.push(header.getRowHeight(i + 1));
        }
        var formula = range.getFormulas();
        var mergedRanges = range.getMergedRanges();
        Logger.log(values);
        if (values.length > 0) {
            dataSheet.insertRows(1, values.length);
        }
        dataSheet
            .getRange(1, 1, values.length, dataSheet.getLastColumn())
            .clearContent()
            .clearFormat();
        for (var i = 0; i < values.length; i++) {
            for (var j = 0; j < values[i].length; j++) {
                if (values[i][j]) {
                    dataSheet.getRange(i + 1, j + 1, 1, 1).setValues([[values[i][j]]]);
                }
                else {
                    dataSheet.getRange(i + 1, j + 1, 1, 1).setFormulas([[formula[i][j]]]);
                }
            }
        }
        var _rangeToExport = rangeToExport;
        if (rangeToExport.split(":").length == 2) {
            if (!isContainsNumber(rangeToExport.split(":")[1])) {
                _rangeToExport = rangeToExport + (to + range.getNumRows() + totalRows);
            }
            else {
                var tempRange = dataSheet.getRange(_rangeToExport);
                _rangeToExport = dataSheet
                    .getRange(tempRange.getRow(), tempRange.getColumn(), tempRange.getNumRows() + totalRows + range.getNumRows(), tempRange.getNumColumns())
                    .getA1Notation();
            }
        }
        for (var i = 0; i < mergedRanges.length; i++) {
            Logger.log(mergedRanges[i].getRow() + " ," + mergedRanges[i].getColumn() + " , " + mergedRanges[i].getLastRow() + " , " + mergedRanges[i].getLastColumn() + " ");
            dataSheet
                .getRange(mergedRanges[i].getRow(), mergedRanges[i].getColumn(), mergedRanges[i].getNumRows(), mergedRanges[i].getNumColumns())
                .merge();
        }
        for (var i = 0; i < height.length; i++) {
            dataSheet.setRowHeight(i + 1, height[i]);
        }
        var headerRange = dataSheet.getRange(range.getRow(), range.getColumn(), range.getNumRows(), range.getNumColumns());
        headerRange.setTextDirections(_TextDirections);
        headerRange.setFontFamilies(_FontFamilies);
        headerRange.setFontSizes(_FontSizes);
        headerRange.setFontColors(_FontColors);
        headerRange.setHorizontalAlignments(_HorizontalAlignments);
        if (config.reapeatHeaders) {
            if (sheet_data.toLowerCase().trim().includes("6") &&
                sheet_data.toLowerCase().trim().includes("auto") &&
                sheet_data.toLowerCase().trim().includes("weights")) {
                dataSheet.setFrozenRows(checkCol + totalRows + frozenRows - 1);
            }
            else {
                dataSheet.setFrozenRows(checkCol + totalRows + frozenRows);
            }
        }
        var fileUrl = generatePDf(dataSheet, _rangeToExport, config.folderToSaveFile, config.fileName, config.reapeatHeaders, config.PageFormat.toLowerCase() == "portrait", config.scale, config.align, config.Margins);
        if (config.rangeToSkip.trim() != "") {
            try {
                var hideRanges = config.rangeToSkip.split(",");
                hideRanges.forEach(function (__range) {
                    __range = __range.trim();
                    if (__range != "") {
                        try {
                            dataSheet.unhideRow(dataSheet.getRange(__range));
                        }
                        catch (e) {
                            Logger.log(e);
                        }
                    }
                });
            }
            catch (e) {
                Logger.log(e);
            }
        }
        if (totalRows > 0) {
            header.deleteRows(1, totalRows);
        }
        if (colsBeforerStart > 0) {
            header.deleteColumns(1, colsBeforerStart);
        }
        if (totalCols > 0) {
            header.deleteColumns(checkCol, totalCols);
        }
        dataSheet.deleteRows(range.getRow(), range.getNumRows());
        dataSheet.setFrozenRows(frozenRows);
        var template = HtmlService.createTemplateFromFile("PDF_complete.html");
        template.url = fileUrl;
        if (!allSheetAtOnce) {
            SpreadsheetApp.getUi().showModalDialog(template.evaluate(), "PDF Export Complete");
        }
        if (isLastSheet) {
            SpreadsheetApp.getUi().alert("All Sheet Exported");
        }
    }
}
function generatePDf(sheet_header, _rangeToExport, folderId, fileName, repeatHeaders, isPotrait, scale, align, margin) {
    Logger.log("_rangeToExport : " + _rangeToExport);
    var range = sheet_header.getRange(_rangeToExport);
    SpreadsheetApp.flush();
    //make pdf
    var rangeParam = "";
    var sheetParam = "";
    if (range) {
        rangeParam =
            "&r1=" +
                (range.getRow() - 1) +
                "&r2=" +
                range.getLastRow() +
                "&c1=" +
                (range.getColumn() - 1) +
                "&c2=" +
                range.getLastColumn();
    }
    if (sheet_header) {
        sheetParam = "&gid=" + sheet_header.getSheetId();
    }
    var url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    var mergin = margin.toLowerCase().trim() == "narrow" ? 0.15 : 0.2;
    var theurl = url.replace(/\/edit.*$/, "") +
        "/export?" +
        "exportFormat=pdf" +
        "&format=pdf" +
        "&size=LETTER" +
        "&portrait=" +
        (isPotrait ? "true" : "false") +
        "&scale=" +
        scale +
        (scale != 1
            ? "&top_margin=" +
                mergin +
                "&bottom_margin=" +
                mergin +
                "&left_margin=" +
                mergin +
                "&right_margin=" +
                mergin
            : "") +
        "&horizontal_alignment=" +
        (align.toLowerCase().trim() == "center" ? "CENTER" : "LEFT") +
        "&sheetnames=false" +
        "&printtitle=false" +
        "&pagenum=CENTER" +
        "&gridlines=false" +
        "&fzr=" +
        (repeatHeaders ? "true" : "false") +
        sheetParam +
        rangeParam;
    Logger.log("exportUrl=" + theurl);
    var token = ScriptApp.getOAuthToken();
    var docurl = UrlFetchApp.fetch(theurl, {
        headers: { Authorization: "Bearer " + token }
    });
    var masterFolder = DriveApp.getFolderById(folderId);
    var finalName = fileName + ".pdf";
    var iterator = masterFolder.getFilesByName(finalName);
    while (iterator.hasNext()) {
        iterator.next().setTrashed(true);
    }
    return masterFolder.createFile(docurl.getBlob()).setName(finalName).getUrl();
}
