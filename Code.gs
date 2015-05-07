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

/**
 * Enables the add-on on for the current spreadsheet (simply by running) and
 * shows a popup informing the user of the new functions that are available.
 */
function use() {
    var ui = SpreadsheetApp.getUi();
    //ui.alert(title, message, ui.ButtonSet.OK);

    if(_isSpreadsheetEmpty()){
        ui.alert('Welcome to the Kickback for Google Sheets wizard')
    }
    else{
        ui.alert('Unfortunately this workbook is not empty. To protect your data, we cannot run a wizard on a non empty workbook.')
    }

    //FOR TESTING HARDCODE USERS
    _addUser('darkmethodz@gmail.com', 'Jason1', 'Kulatunga', 'USD')
    _addUser('d.arkmethodz@gmail.com', 'Jason2', 'Kulatunga', 'USD')
    _addUser('da.rkmethodz@gmail.com', 'Jason3', 'Kulatunga', 'USD')
    _addUser('dar.kmethodz@gmail.com', 'Jason4', 'Kulatunga', 'USD')
    _addUser('dark.methodz@gmail.com', 'Jason5', 'Kulatunga', 'USD')


}

/**
 * This function will populate the google workbook with our designed sheets and base formulas
 * @private
 */
function _populateWorkbook(){
    var workbook = SpreadsheetApp.getActiveSpreadsheet();
    var transactionsSheet = workbook.insertSheet('Transactions', 0)
    var summarySheet = workbook.insertSheet('Summary', 1)

    //delete any other sheets.
    var sheets = workbook.getSheets()
    if(sheets.length > 2){
        for(var ndx =2; ndx<sheets.length; ndx++){
            workbook.deleteSheet(sheets[ndx])
        }
    }

    //populate the transactions sheet.
    //The transactions sheet has 3 distinct sections, entry information, payee information, payment information
    transactionsSheet.activate();
    var entryHeaderRange = transactionsSheet.getRange(1, 1, 5);
    entryHeaderRange.setValues(['Location','Item','Currency','Amount Paid', 'Amount Paid (USD)'])
    entryHeaderRange.setFontColor('#b1b2b1');
    entryHeaderRange.setFontFamily('Open Sans');
    entryHeaderRange.setFontSize(14)
    entryHeaderRange.setFontWeight("bold");
    entryHeaderRange.gsetHorizontalAlignment("center");
    
    //finishing up, make the transactions sheet the active sheet
    workbook.setActiveSheet(transactionsSheet)

}


function _addUser(email, first_name, last_name){
    var documentProperties = PropertiesService.getDocumentProperties();

    //send the user an invitation to the sheet.
    var workbook = SpreadsheetApp.getActiveSpreadsheet();
    workbook.addEditor(email)

    //save the user to the document properties.
    var users = documentProperties.getProperty('USERS')
    users[email] = {
        first_name: first_name,
        last_name: last_name,
        display_name: first_name + ' ' + last_name[0]
    }
    documentProperties.setProperty('USERS', users)

}

/**
 * Function will loop though all sheets in the spreadsheet and ensure that there is no content set.
 * @private
 */
function _isSpreadsheetEmpty(){
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
    for(var ndx in sheets){
        var sheet = sheets[ndx];
        if(sheet.getLastRow() != 0 || sheet.getLastColumn() != 0){
            return false
        }
    }
    return true
}
