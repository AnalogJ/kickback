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
    _addUser('darkmethodz@gmail.com', 'Jason1', 'Kulatunga', 'USD');
    _addUser('d.arkmethodz@gmail.com', 'Jason2', 'Kulatunga', 'USD');
    _addUser('da.rkmethodz@gmail.com', 'Jason3', 'Kulatunga', 'USD');
    _addUser('dar.kmethodz@gmail.com', 'Jason4', 'Kulatunga', 'USD');
    _addUser('dark.methodz@gmail.com', 'Jason5', 'Kulatunga', 'USD');

    _populateWorkbook()
}

//*************************************************************************************************
// Style/Design functions
//*************************************************************************************************

var HEADER_FONT_SIZE = 12;
var HEADER_FONT_FAMILY = 'Open Sans';
var HEADER_BACKGROUND_COLOR = '#b1b2b1';

var COLOR_SWATCHES = ['#468966', '#FFF0A5', '#FFB03B', '#B64926', '#8E2800','#0F2D40','#194759','#296B73','#3E8C84','#D8F2F0']

/**
 * This function will populate the google workbook with our designed sheets and base formulas
 * @private
 */
function _populateWorkbook(){
    var workbook = SpreadsheetApp.getActiveSpreadsheet();
    var transactionsSheet = workbook.insertSheet('Transactions', 0);
    //var summarySheet = workbook.insertSheet('Summary', 1)

    //delete any other sheets.
    var sheets = workbook.getSheets();
    if(sheets.length > 2){
        for(var ndx =2; ndx<sheets.length; ndx++){
            workbook.deleteSheet(sheets[ndx])
        }
    }

    var documentProperties = PropertiesService.getDocumentProperties();
    var users = _getUsers();
    //populate the transactions sheet.
    //The transactions sheet has 3 distinct sections, entry information, payee information, payment information
    transactionsSheet.activate();

    //getRange(row, column, numRows, numColumns)
    var ENTRY_HEADER_TEXT = [['Date Purchased','Location','Item','Currency','Amount Paid', 'Amount Paid (USD)', 'Who Paid'],
        ['','','','','','','']];
    var ENTRY_HEADER_LEFT = 1;
    var ENTRY_HEADER_LEFT_OFFSET = ENTRY_HEADER_TEXT[0].length;

    var entryHeaderRange = transactionsSheet.getRange(1,ENTRY_HEADER_LEFT,2,ENTRY_HEADER_LEFT_OFFSET);
    entryHeaderRange.mergeVertically();
    entryHeaderRange.setValues(ENTRY_HEADER_TEXT);
    entryHeaderRange.setBackgroundColor(HEADER_BACKGROUND_COLOR);
    entryHeaderRange.setFontFamily(HEADER_FONT_FAMILY);
    entryHeaderRange.setFontSize(HEADER_FONT_SIZE);
    entryHeaderRange.setFontWeight("bold");
    entryHeaderRange.setHorizontalAlignment("center");
    entryHeaderRange.setBorder(true, true, true, true, false, false);
    entryHeaderRange.setWrap(true);

    //getRange(row, column, numRows, numColumns)
    var PAYEE_HEADER_LEFT = ENTRY_HEADER_LEFT + ENTRY_HEADER_LEFT_OFFSET + 1;
    var PAYEE_HEADER_LEFT_OFFSET = users.length

    var payeeHeaderTopRange = transactionsSheet.getRange(1,PAYEE_HEADER_LEFT,1,PAYEE_HEADER_LEFT_OFFSET)
    payeeHeaderTopRange.mergeAcross();
    payeeHeaderTopRange.setValue('Paid For Who');
    payeeHeaderTopRange.setBackgroundColor(HEADER_BACKGROUND_COLOR);
    payeeHeaderTopRange.setFontFamily(HEADER_FONT_FAMILY);
    payeeHeaderTopRange.setFontSize(HEADER_FONT_SIZE);
    payeeHeaderTopRange.setFontWeight("bold");
    payeeHeaderTopRange.setHorizontalAlignment("center");
    payeeHeaderTopRange.setBorder(true, true, true, true, false, false);

    //getRange(row, column, numRows, numColumns)
    var payeeHeaderBottomRange = transactionsSheet.getRange(2,PAYEE_HEADER_LEFT,1,PAYEE_HEADER_LEFT_OFFSET);
    var names = [];
    for(var ndx in users){
        names.push(users[ndx].display_name)
    }
    payeeHeaderBottomRange.setValues([names]);
    payeeHeaderBottomRange.setBackgroundColor('#d0d0d0');
    payeeHeaderBottomRange.setFontFamily(HEADER_FONT_FAMILY);
    payeeHeaderBottomRange.setFontSize(9);
    payeeHeaderBottomRange.setHorizontalAlignment("center");
    payeeHeaderBottomRange.setBorder(true, true, true, true, false, false);

    //getRange(row, column, numRows, numColumns)
    var PAYMENT_HEADER_TEXT = [['Self Pay','Ind. Payment','Payer Collects'],['','','']];
    var PAYMENT_HEADER_LEFT = PAYEE_HEADER_LEFT + PAYEE_HEADER_LEFT_OFFSET + 1;
    var PAYMNET_HEADER_LEFT_OFFSET = PAYMENT_HEADER_TEXT[0].length;

    var paymentHeaderRange = transactionsSheet.getRange(1,PAYMENT_HEADER_LEFT, 2,PAYMNET_HEADER_LEFT_OFFSET);
    paymentHeaderRange.mergeVertically();
    paymentHeaderRange.setValues(PAYMENT_HEADER_TEXT);
    paymentHeaderRange.setBackgroundColor(HEADER_BACKGROUND_COLOR);
    paymentHeaderRange.setFontFamily(HEADER_FONT_FAMILY);
    paymentHeaderRange.setFontSize(HEADER_FONT_SIZE);
    paymentHeaderRange.setFontWeight("bold");
    paymentHeaderRange.setHorizontalAlignment("center");
    paymentHeaderRange.setBorder(true, true, true, true, false, false);
    entryHeaderRange.setWrap(true);

    //hide the first
    transactionsSheet.hideRows(3);
    transactionsSheet.setFrozenRows(2);


}

//*************************************************************************************************
// Document Storage functions
//*************************************************************************************************


function _getUsers(){
    var documentProperties = PropertiesService.getDocumentProperties();
    var users_str = documentProperties.getProperty('USERS') || '';
    return JSON.parse(users_str);
}

function _addUser(email, first_name, last_name){
    var documentProperties = PropertiesService.getDocumentProperties();

    //send the user an invitation to the sheet.
    var workbook = SpreadsheetApp.getActiveSpreadsheet();
    workbook.addEditor(email);

    //save the user to the document properties.
    var users_str = documentProperties.getProperty('USERS') || '';
    var users = [];
    if(users_str){
        users = JSON.parse(users_str);
    }

    var new_user = {
        first_name: first_name,
        last_name: last_name,
        display_name: first_name + ' ' + last_name[0]
    }

    if(!_arrayContains(users, new_user)){
        users.push(new_user);
        documentProperties.setProperty('USERS', JSON.stringify(users));
    }


}


//*************************************************************************************************
// Utility functions
//*************************************************************************************************

function _arrayContains(a, obj) {
    for (var i = 0; i < a.length; i++) {
        if (a[i] === obj) {
            return true;
        }
    }
    return false;
}

/**
 * Function will loop though all sheets in the spreadsheet and ensure that there is no content set.
 * @private
 */
function _isSpreadsheetEmpty(){
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for(var ndx in sheets){
        var sheet = sheets[ndx];
        if(sheet.getLastRow() != 0 || sheet.getLastColumn() != 0){
            return false
        }
    }
    return true
}
