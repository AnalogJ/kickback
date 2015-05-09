/**
 * @fileoverview Provides the custom functions DATEADD and DATESUBTRACT and
 * the helper functions that they use.
 */




/**
 * Runs when the add-on is installed.
 */
function onInstall() {
    onOpen();
}

/**
 * Runs when the document is opened, creating the add-on's menu. Custom function
 * add-ons need at least one menu item, since the add-on is only enabled in the
 * current spreadsheet when a function is run.
 */
function onOpen() {
    SpreadsheetApp.getUi().createAddonMenu()
        .addItem('Run Kickback Wizard', 'use')
        .addToUi();
}

function use(){
    //test assert
    //assert(false)

    //matches
    var rangeA = SpreadsheetApp.getActiveSheet().getRange("A1:C2");
    var rangeB = SpreadsheetApp.getActiveSheet().getRange("A1:C2");
    assert(_rangeIntersect(rangeA, rangeB))

    //contains
    rangeA = SpreadsheetApp.getActiveSheet().getRange("A1:C2");
    rangeB = SpreadsheetApp.getActiveSheet().getRange("B1");
    assert(_rangeIntersect(rangeA, rangeB))

    //inverse contains
    rangeA = SpreadsheetApp.getActiveSheet().getRange("B1");
    rangeB = SpreadsheetApp.getActiveSheet().getRange("A1:C2");
    assert(_rangeIntersect(rangeA, rangeB))

    //does not intesect
    rangeA = SpreadsheetApp.getActiveSheet().getRange("A1:C2");
    rangeB = SpreadsheetApp.getActiveSheet().getRange("D1:E2");
    assert(_rangeIntersect(rangeA, rangeB) == false)

    //intersects a column
    rangeA = SpreadsheetApp.getActiveSheet().getRange("A1:C2");
    rangeB = SpreadsheetApp.getActiveSheet().getRange("B:B");
    assert(_rangeIntersect(rangeA, rangeB))

}

function _rangeIntersect(rangeA, rangeB){
    var rangeA_row_begin = rangeA.getRow();
    var rangeA_row_end = rangeA.getLastRow();
    var rangeA_col_begin = rangeA.getColumn();
    var rangeA_col_end = rangeA.getLastColumn();

    var rangeB_row_begin = rangeB.getRow();
    var rangeB_row_end = rangeB.getLastRow();
    var rangeB_col_begin = rangeB.getColumn();
    var rangeB_col_end = rangeB.getLastColumn();
    return ((rangeA_row_begin <= rangeB_row_begin && rangeB_row_begin <= rangeA_row_end) || (rangeB_row_begin <= rangeA_row_begin && rangeA_row_begin <= rangeB_row_end)) &&
        ((rangeA_col_begin <= rangeB_col_begin && rangeB_col_begin <= rangeA_col_end) || (rangeB_col_begin <= rangeA_col_begin && rangeA_col_begin <= rangeB_col_end))

}

function assert(condition, message) {
    if (!condition) {
        throw message || "Assertion failed";
    }
}