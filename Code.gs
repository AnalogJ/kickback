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

    _setTripCurrency('USD');
    _setUserCurrency('CAD');

    //FOR TESTING HARDCODE USERS
    _addUser('darkmethodz@gmail.com', 'Jason1', 'Kulatunga');
    _addUser('d.arkmethodz@gmail.com', 'Jason2', 'Kulatunga');
    _addUser('da.rkmethodz@gmail.com', 'Jason3', 'Kulatunga');
    _addUser('dar.kmethodz@gmail.com', 'Jason4', 'Kulatunga');
    _addUser('dark.methodz@gmail.com', 'Jason5', 'Kulatunga');

    _populateWorkbook()
}

//*************************************************************************************************
// Style/Design functions
//*************************************************************************************************

var COLOR_SWATCHES = ['#468966', '#FFF0A5', '#FFB03B', '#B64926', '#8E2800','#0F2D40','#194759','#296B73','#3E8C84','#D8F2F0']

/**
 * This function will populate the google workbook with our designed sheets and base formulas
 * @private
 */
function _populateWorkbook() {
    var workbook = SpreadsheetApp.getActiveSpreadsheet();
    var transactionsSheet = workbook.insertSheet('Transactions', 0);
    var summarySheet = workbook.insertSheet('Summary', 1)

    //delete any other sheets.
    var sheets = workbook.getSheets();
    if (sheets.length > 2) {
        for (var ndx = 2; ndx < sheets.length; ndx++) {
            workbook.deleteSheet(sheets[ndx])
        }
    }

    _configureTransactionsSheet(transactionsSheet);
    _configureSummarySheet(summarySheet);
    transactionsSheet.activate();
}

function _configureTransactionsSheet(transactionsSheet){
    var BODY_TOP = 3;
    var BODY_TOP_OFFSET = 50;

    var users = _getUsers();
    //populate the transactions sheet.
    //The transactions sheet has 3 distinct sections, entry information, payee information, payment information
    transactionsSheet.activate();

    //getRange(row, column, numRows, numColumns)
    var ENTRY_HEADER_TEXT = [['Date Purchased','Location','Item','Currency','Amount Paid', 'Amount Paid ('+_getUserCurrency()+')', 'Who Paid'],
        ['','','','','','','']];
    var ENTRY_HEADER_LEFT = 1;
    var ENTRY_HEADER_LEFT_OFFSET = ENTRY_HEADER_TEXT[0].length;

    var entryHeaderRange = transactionsSheet.getRange(1,ENTRY_HEADER_LEFT,2,ENTRY_HEADER_LEFT_OFFSET);
    entryHeaderRange.mergeVertically();
    entryHeaderRange.setValues(ENTRY_HEADER_TEXT);
    _setHeaderStyle(entryHeaderRange);

    //getRange(row, column, numRows, numColumns)
    var PAYEE_HEADER_LEFT = ENTRY_HEADER_LEFT + ENTRY_HEADER_LEFT_OFFSET + 1;
    var PAYEE_HEADER_LEFT_OFFSET = users.length;

    var payeeHeaderTopRange = transactionsSheet.getRange(1,PAYEE_HEADER_LEFT,1,PAYEE_HEADER_LEFT_OFFSET)
    payeeHeaderTopRange.mergeAcross();
    payeeHeaderTopRange.setValue('Paid For Who');
    _setHeaderStyle(payeeHeaderTopRange);

    //getRange(row, column, numRows, numColumns)
    var payeeHeaderBottomRange = transactionsSheet.getRange(2,PAYEE_HEADER_LEFT,1,PAYEE_HEADER_LEFT_OFFSET);
    var names = [];
    for(var ndx in users){
        names.push(users[ndx].display_name)
    }
    payeeHeaderBottomRange.setValues([names]);
    _setSubHeaderStyle(payeeHeaderBottomRange);


    //getRange(row, column, numRows, numColumns)
    var PAYMENT_HEADER_TEXT = [['Self Pay','Ind. Payment','Payer Collects'],['','','']];
    var PAYMENT_HEADER_LEFT = PAYEE_HEADER_LEFT + PAYEE_HEADER_LEFT_OFFSET + 1;
    var PAYMENT_HEADER_LEFT_OFFSET = PAYMENT_HEADER_TEXT[0].length;

    var paymentHeaderRange = transactionsSheet.getRange(1,PAYMENT_HEADER_LEFT, 2,PAYMENT_HEADER_LEFT_OFFSET);
    paymentHeaderRange.mergeVertically();
    paymentHeaderRange.setValues(PAYMENT_HEADER_TEXT);
    _setHeaderStyle(paymentHeaderRange);

    //hide the first
    transactionsSheet.hideRows(3);
    transactionsSheet.setFrozenRows(2);


    //set the date purchased validation
    var datePurchasedColumn = ENTRY_HEADER_LEFT + ENTRY_HEADER_TEXT[0].indexOf('Date Purchased');
    var datePurchasedBodyRange = transactionsSheet.getRange(BODY_TOP,datePurchasedColumn, BODY_TOP_OFFSET,1);
    var datePurchasedRule = SpreadsheetApp.newDataValidation()
        .requireDate()
        .setAllowInvalid(true)
        .setHelpText('Date this item was purchased. This date is used to do any required currency conversion.')
        .build();
    datePurchasedBodyRange.setDataValidation(datePurchasedRule);
    datePurchasedBodyRange.setNumberFormat("dd-mm-yyyy");
    _setBodyStyle(datePurchasedBodyRange);

    //set the currency validation.
    //getRange(row, column, numRows, numColumns)
    var currencyColumn = ENTRY_HEADER_LEFT + ENTRY_HEADER_TEXT[0].indexOf('Currency');
    var currencyBodyRange = transactionsSheet.getRange(BODY_TOP,currencyColumn, BODY_TOP_OFFSET,1);

    var currencies = [_getTripCurrency(),_getUserCurrency()];
    var currencyRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(currencies, true)
        .setAllowInvalid(false)
        .setHelpText('Currency used to pay for this item.')
        .build();
    currencyBodyRange.setDataValidation(currencyRule);
    _setBodyStyle(currencyBodyRange);

    //set the amount paid validation
    var amountPaidColumn = ENTRY_HEADER_LEFT + ENTRY_HEADER_TEXT[0].indexOf('Amount Paid');
    var amountPaidBodyRange = transactionsSheet.getRange(BODY_TOP,amountPaidColumn, BODY_TOP_OFFSET,1);
    var amountPaidRule = SpreadsheetApp.newDataValidation()
        .requireNumberGreaterThan(0)
        .setAllowInvalid(false)
        .setHelpText('Amount this item was puchased for')
        .build();
    amountPaidBodyRange.setDataValidation(amountPaidRule);
    amountPaidBodyRange.setNumberFormat("$0.00");
    _setBodyStyle(amountPaidBodyRange);


    //set the amount paid currency conversion
    var amountPaidUserColumn = ENTRY_HEADER_LEFT + ENTRY_HEADER_TEXT[0].indexOf('Amount Paid ('+_getUserCurrency()+')');
    var amountPaidUserBodyRange = transactionsSheet.getRange(BODY_TOP,amountPaidUserColumn, BODY_TOP_OFFSET,1);
    var amountPaidUserRule = SpreadsheetApp.newDataValidation()
        .requireNumberGreaterThan(0)
        .setAllowInvalid(false)
        .setHelpText('Amount this item was puchased for')
        .build();
    amountPaidUserBodyRange.setDataValidation(amountPaidUserRule);
    amountPaidUserBodyRange.setNumberFormat("$0.00");
//TODO: look at the google Finanace method and lookup a specific date.
    amountPaidUserBodyRange.setFormulaR1C1("=GoogleFinance('Currency:USDCAD')*R[0]C[-1]");
    _setBodyStyle(amountPaidUserBodyRange);

    var whoPaidColumn =  ENTRY_HEADER_LEFT + ENTRY_HEADER_TEXT[0].indexOf('Who Paid');
    var whoPaidBodyRange = transactionsSheet.getRange(BODY_TOP,whoPaidColumn, BODY_TOP_OFFSET,1);
    var whoPaidRule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(payeeHeaderBottomRange, true)
        .setAllowInvalid(false)
        .setHelpText('The user who paid for this item.')
        .build();
    whoPaidBodyRange.setDataValidation(whoPaidRule);
    _setBodyStyle(whoPaidBodyRange);

    var paidForColumn = PAYEE_HEADER_LEFT
    var paidForBodyRange = transactionsSheet.getRange(BODY_TOP,paidForColumn, BODY_TOP_OFFSET,PAYMENT_HEADER_LEFT_OFFSET);
    var paidForRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Y','YS'], true)
        .setAllowInvalid(false)
        .setHelpText('When payer pays for themselves, use `YS`')
        .build();
    paidForBodyRange.setDataValidation(paidForRule);
    _setBodyStyle(paidForBodyRange);

    var selfPayColumn = PAYMENT_HEADER_TEXT + PAYMENT_HEADER_TEXT[0].indexOf('Self Pay');
    var selfPayBodyRange = transactionsSheet.getRange(BODY_TOP,selfPayColumn, BODY_TOP_OFFSET,1);
//TODO: fill this range with a formula
    _setBodyStyle(selfPayBodyRange);

    var indPaymentColumn = PAYMENT_HEADER_TEXT + PAYMENT_HEADER_TEXT[0].indexOf('Ind. Payment');
    var indPaymentBodyRange = transactionsSheet.getRange(BODY_TOP,indPaymentColumn, BODY_TOP_OFFSET,1);
//TODO: fill this range with a formula
    _setBodyStyle(indPaymentBodyRange);

    var payerCollectsColumn = PAYMENT_HEADER_TEXT + PAYMENT_HEADER_TEXT[0].indexOf('Payer Collects');
    var payerCollectsBodyRange = transactionsSheet.getRange(BODY_TOP,payerCollectsColumn, BODY_TOP_OFFSET,1);
//TODO: fill this range with a formula
    _setBodyStyle(payerCollectsBodyRange);
    payerCollectsBodyRange.setFontWeight("bold");
}

function _configureSummarySheet(summarySheet){
    var users = _getUsers();

    var BODY_TOP = 2;
    var BODY_TOP_OFFSET = users.length;

    //populate the transactions sheet.
    //The summary sheet has one block of information.
    summarySheet.activate();

    //getRange(row, column, numRows, numColumns)
    var SUMMARY_HEADER_TEXT = ['Name','Gets','Gives','Banker Collects','Rounded', '% Difference'];
    var SUMMARY_HEADER_LEFT = 1;
    var SUMMARY_HEADER_LEFT_OFFSET = SUMMARY_HEADER_TEXT[0].length;

    var summaryHeaderRange = summarySheet.getRange(1,SUMMARY_HEADER_LEFT,1,SUMMARY_HEADER_LEFT_OFFSET);
    _setSubHeaderStyle(summaryHeaderRange);
    summaryHeaderRange.setFontWeight("bold");
    summaryHeaderRange.setHorizontalAlignment("left");

    //set the date purchased validation
    var nameColumn = SUMMARY_HEADER_LEFT + SUMMARY_HEADER_TEXT.indexOf('Name');
    var nameBodyRange = summarySheet.getRange(BODY_TOP,nameColumn, BODY_TOP_OFFSET,1);
    var nameBodyRangeValues = [];
    var nameBodyRangeBackgrounds = [];
    for(var ndx in users){
        nameBodyRangeValues.push([users[ndx]])
        nameBodyRangeBackgrounds.push([COLOR_SWATCHES[ndx % COLOR_SWATCHES.length]])
    }
    _setHeaderStyle(nameBodyRange);
    nameBodyRange.setValues(nameBodyRangeValues);
    nameBodyRange.setBackgrounds(nameBodyRangeBackgrounds);

    var getsColumn = SUMMARY_HEADER_LEFT + SUMMARY_HEADER_TEXT.indexOf('Gets');
    var getsBodyRange = summarySheet.getRange(BODY_TOP,getsColumn, BODY_TOP_OFFSET,1);
//TODO: fill this range with a formula
    _setSummaryBodyStyle(getsBodyRange);

    var givesColumn = SUMMARY_HEADER_LEFT + SUMMARY_HEADER_TEXT.indexOf('Gives');
    var givesBodyRange = summarySheet.getRange(BODY_TOP,givesColumn, BODY_TOP_OFFSET,1);
//TODO: fill this range with a formula
    _setSummaryBodyStyle(givesBodyRange);

    var bankerCollectsColumn = SUMMARY_HEADER_LEFT + SUMMARY_HEADER_TEXT.indexOf('Banker Collects');
    var bankerCollectsBodyRange = summarySheet.getRange(BODY_TOP,bankerCollectsColumn, BODY_TOP_OFFSET,1);
//TODO: fill this range with a formula
    _setSummaryBodyStyle(bankerCollectsBodyRange);

    var roundedColumn = SUMMARY_HEADER_LEFT + SUMMARY_HEADER_TEXT.indexOf('Rounded');
    var roundedBodyRange = summarySheet.getRange(BODY_TOP, roundedColumn, BODY_TOP_OFFSET,1);
//TODO: fill this range with a formula
    _setSummaryBodyStyle(roundedBodyRange);

    var percentDiffColumn = SUMMARY_HEADER_LEFT + SUMMARY_HEADER_TEXT.indexOf('% Difference');
    var percentDiffBodyRange = summarySheet.getRange(BODY_TOP, percentDiffColumn, BODY_TOP_OFFSET,1);
//TODO: fill this range with a formula
    _setSummaryBodyStyle(percentDiffBodyRange);
}

var HEADER_FONT_SIZE = 12;
var HEADER_FONT_FAMILY = 'Open Sans';
var HEADER_BACKGROUND_COLOR = '#b1b2b1';

function _setHeaderStyle(range){
    range.setBackgroundColor(HEADER_BACKGROUND_COLOR);
    range.setFontFamily(HEADER_FONT_FAMILY);
    range.setFontSize(HEADER_FONT_SIZE);
    range.setFontWeight("bold");
    range.setHorizontalAlignment("center");
    range.setBorder(true, true, true, true, false, false);
    range.setWrap(true);
}
function _setSubHeaderStyle(range){
    range.setBackgroundColor('#d0d0d0');
    range.setFontFamily(HEADER_FONT_FAMILY);
    range.setFontSize(9);
    range.setHorizontalAlignment("center");
    range.setBorder(true, true, true, true, false, false);
}

function _setBodyStyle(range){
    range.setFontSize(7);
    range.setWrap(true);
}

function _setSummaryBodyStyle(range){
    range.setBackgroundColor('#d0d0d0');
    range.setFontFamily(HEADER_FONT_FAMILY);
    range.setFontSize(14);
}



//*************************************************************************************************
// Document Storage functions
//*************************************************************************************************


function _getUsers(){
    //var documentProperties = PropertiesService.getDocumentProperties();
    //var users_str = documentProperties.getProperty('USERS') || '';
    //return JSON.parse(users_str);
    return [
        {first_name:'Jas1',last_name:'K',display_name:'Jas1 K'},
        {first_name:'Jas2',last_name:'K',display_name:'Jas2 K'},
        {first_name:'Jas3',last_name:'K',display_name:'Jas3 K'},
        {first_name:'Jas4',last_name:'K',display_name:'Jas4 K'},
    ]
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

function  _getTripCurrency(){
    var documentProperties = PropertiesService.getDocumentProperties();
    return documentProperties.getProperty('TRIP_CURRENCY') || '';
}

function  _setTripCurrency(currency_code){
    var documentProperties = PropertiesService.getDocumentProperties();
    return documentProperties.setProperty('TRIP_CURRENCY',currency_code||'');
}

function  _getUserCurrency(){
    var documentProperties = PropertiesService.getDocumentProperties();
    return documentProperties.getProperty('USER_CURRENCY') || '';
}

function  _setUserCurrency(currency_code){
    var documentProperties = PropertiesService.getDocumentProperties();
    return documentProperties.setProperty('USER_CURRENCY',currency_code||'');
}
//*************************************************************************************************
// Utility functions
//*************************************************************************************************

function _arrayContains(a, obj) {
    for (var i = 0; i < a.length; i++) {
        if (a[i] == obj) {
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
