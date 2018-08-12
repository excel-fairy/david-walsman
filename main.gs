var COLUMNS = {
    firstSaleColumn: 'A',
    transactionIdColumn: 'B',
    lastSaleColumn: 'Q',
};

var MOS_SCHOOL_PAYPAL_SALES_SPREADSHEET = {
    allPaypalSalesSheet: {
        sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ALL PAYPAL SALES'),
        firstSaleRow: 3
    }
};

var ZAPIER_IMPORT_SPREADSHEET = {
    importSheet: {
        name: 'Sheet1',
        firstSaleRow: 2
    }
};


/**
 * Time-based execution
 */
function importSales() {
    var lastImportedSaleRow = getAllSales(MOS_SCHOOL_PAYPAL_SALES_SPREADSHEET.allPaypalSalesSheet.sheet,
        MOS_SCHOOL_PAYPAL_SALES_SPREADSHEET.allPaypalSalesSheet.firstSaleRow).length + MOS_SCHOOL_PAYPAL_SALES_SPREADSHEET.allPaypalSalesSheet.firstSaleRow;
    var salesToImport = getUnimportedSales(getLastTransactionIdFromMOSSchoolPaypalSalesSpreadhseet());
    var destinationRange = MOS_SCHOOL_PAYPAL_SALES_SPREADSHEET.allPaypalSalesSheet.sheet.getRange(
        lastImportedSaleRow + 1,
        ColumnNames.letterToColumn(COLUMNS.firstSaleColumn),
        salesToImport.length,
        ColumnNames.letterToColumn(COLUMNS.lastSaleColumn) -
        ColumnNames.letterToColumn(COLUMNS.firstSaleColumn)
    );
    destinationRange.setValues(salesToImport);
}

function getAllSales(sheet, startRow) {
    return sheet.getRange(startRow,
        ColumnNames.letterToColumn(COLUMNS.firstSaleColumn),
        sheet.getLastRow()- startRow,
        ColumnNames.letterToColumn(COLUMNS.lastSaleColumn) -
        ColumnNames.letterToColumn(COLUMNS.firstSaleColumn)).getValues();
}

function getLastTransactionIdFromMOSSchoolPaypalSalesSpreadhseet() {
    var allSales = getAllSales(MOS_SCHOOL_PAYPAL_SALES_SPREADSHEET.allPaypalSalesSheet.sheet,
        MOS_SCHOOL_PAYPAL_SALES_SPREADSHEET.allPaypalSalesSheet.firstSaleRow);
    return allSales[allSales.length-1][ColumnNames.letterToColumnStart0(COLUMNS.transactionIdColumn)];
}

function getFirstUnimportedSaleRow(lastImportedTransactionId) {
    var allSales = getAllSales(getZapierImportSheet(), ZAPIER_IMPORT_SPREADSHEET.importSheet.firstSaleRow);
    for(var i=0; i < allSales.length; i++) {
         var transactionId = allSales[i][ColumnNames.letterToColumnStart0(COLUMNS.transactionIdColumn)];
         if(lastImportedTransactionId === transactionId)
             return i + 1 + ZAPIER_IMPORT_SPREADSHEET.importSheet.firstSaleRow;
    }
    return null;
}

function getUnimportedSales(lastImportedTransactionId) {
    var firstUnImportedSalesRow = getFirstUnimportedSaleRow(lastImportedTransactionId);
    var allSales = getAllSales(getZapierImportSheet(), ZAPIER_IMPORT_SPREADSHEET.importSheet.firstSaleRow);
    if(firstUnImportedSalesRow !== null || allSales.length <= firstUnImportedSalesRow) {
        var zapierImportSheet = SpreadsheetApp.openById(ZAPIER_IMPORT_SPREADSHEET_ID).getSheetByName(ZAPIER_IMPORT_SPREADSHEET.importSheet.name);
        return zapierImportSheet.getRange(firstUnImportedSalesRow + 1,
            ColumnNames.letterToColumn(COLUMNS.firstSaleColumn),
            zapierImportSheet.getLastRow() - firstUnImportedSalesRow + 2,
            ColumnNames.letterToColumn(COLUMNS.lastSaleColumn) -
            ColumnNames.letterToColumn(COLUMNS.firstSaleColumn)
        ).getValues();
    }
}

function getZapierImportSheet() {
    return SpreadsheetApp.openById(ZAPIER_IMPORT_SPREADSHEET_ID).getSheetByName(ZAPIER_IMPORT_SPREADSHEET.importSheet.name);
}
