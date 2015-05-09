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
 * watch for changes in the following source columns:
 * Currency
 * Amount Paid
 * Who Paid
 * Paid for Who
 * @param e
 */
function onEdit(e){
    Logger.log(e)

    var currentSheet = e.range.getSheet();
    if(currentSheet.getName() != "Transactions"){
        //we dont care about changes to any sheet other than the Transactions sheet.
        return;
    }

    var workbook = SpreadsheetApp.getActiveSpreadsheet();
    var currencyRange = workbook.getRangeByName('TRANSACTIONS_BODY_CURRENCY');
    var amountPaidRange = workbook.getRangeByName('TRANSACTIONS_BODY_AMOUNT_PAID');
    var whoPaidRange = workbook.getRangeByName('TRANSACTIONS_BODY_WHO_PAID');
    var paidForRange = workbook.getRangeByName('TRANSACTIONS_BODY_PAID_FOR');

    if(!(_rangeIntersect(e.range,currencyRange) ||
        _rangeIntersect(e.range,amountPaidRange) ||
        _rangeIntersect(e.range,whoPaidRange) ||
        _rangeIntersect(e.range,paidForRange)
        )){
        //this edited range does not intesect with a watched range.
        return;
    }

    //Before processing rows, ensure that the rows we process match the row we care about.
    var first_row = Math.max(e.range.getRow(),currencyRange.getRow());
    var last_row = Math.min(e.range.getLastRow(),currencyRange.getLastRow());

    for(var row = first_row; row<=last_row; row++){
        //set the background for the currency col.
        var currencyCell = currentSheet.getRange(row, currencyRange.getColumn());
        if(currencyCell.getValue() == _getUserCurrency()){
            //if the currency of this item is the same as the user currency, the background color should be white.
            currencyCell.setBackground('white')
        }
        else{
            //currency is different, we should highlight that fact
            currencyCell.setBackground(COLOR_LIGHT_GREEN)
        }

        //set the background for the who paid column.
        var whoPaidCell = currentSheet.getRange(row, whoPaidRange.getColumn());
        whoPaidCell.setBackground(_getUserColor(whoPaidCell.getValue()));

        //paid for who columns.
        var paidForCells = currentSheet.getRange(row, paidForRange.getColumn(), 1, paidForRange.getLastColumn()-paidForRange.getColumn());
        var paidForCellsValues =  paidForCells.getValues();
        var paidForCellsBackgrounds = [];
        for(var ndx in paidForCellsValues[0]){
            var cellValue = paidForCellsValues[0][ndx];
            if(cellValue == "Y" || cellValue == "YS"){
                paidForCellsBackgrounds.push(COLOR_LIGHT_GREEN)
            }
            else{
                paidForCellsBackgrounds.push('white')
            }
        }
        paidForCells.setBackgrounds([paidForCellsBackgrounds]);
    }

    //update the background colors on the summary sheet.
    var summarySheet = workbook.getSheetByName('Summary');
    var bankerCollectsRange = workbook.getRangeByName('SUMMARY_BODY_BANKER_COLLECTS');

    var bankerCollectsCells = summarySheet.getRange(2, bankerCollectsRange.getColumn(), _getUsers().length,1);
    var bankerCollectsCellsValues =  bankerCollectsCells.getValues();
    var bankerCollectsCellsBackgrounds = [];
    for(var ndx in bankerCollectsCellsValues){
        var cellValue = bankerCollectsCellsValues[ndx][0];
        if(cellValue >= 0){
            paidForCellsBackgrounds.push([COLOR_LIGHT_GREEN])
        }
        else{
            paidForCellsBackgrounds.push([COLOR_LIGHT_RED])
        }
    }
    bankerCollectsCells.setBackgrounds(bankerCollectsCellsBackgrounds);
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
var COLOR_LIGHT_GREEN = '#bbedc3';
var COLOR_LIGHT_RED = '#feb8c3';
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

    var transactionsColumns = _configureTransactionsSheet(workbook,transactionsSheet);
    _configureSummarySheet(workbook,summarySheet,transactionsColumns);
    transactionsSheet.activate();
}

function _configureTransactionsSheet(workbook,transactionsSheet){
    var BODY_TOP = 3;
    var BODY_TOP_OFFSET = 50;

    var users = _getUsers();
    //populate the transactions sheet.
    //The transactions sheet has 3 distinct sections, entry information, payee information, payment information
    transactionsSheet.activate();

    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    // HEADER SETUP
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

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
    var PAYEE_HEADER_LEFT = ENTRY_HEADER_LEFT + ENTRY_HEADER_LEFT_OFFSET;
    var PAYEE_HEADER_LEFT_OFFSET = users.length;

    //getRange(row, column, numRows, numColumns)
    var payeeHeaderBottomRange = transactionsSheet.getRange(2,PAYEE_HEADER_LEFT,1,PAYEE_HEADER_LEFT_OFFSET);
    var names = [];
    for(var ndx in users){
        names.push(users[ndx].display_name)
    }
    payeeHeaderBottomRange.setValues([names]);
    _setSubHeaderStyle(payeeHeaderBottomRange);
    //resize the payee columns
    for(var ndx in users){
        transactionsSheet.autoResizeColumn(parseInt(PAYEE_HEADER_LEFT)+parseInt(ndx));
    }

    var payeeHeaderTopRange = transactionsSheet.getRange(1,PAYEE_HEADER_LEFT,1,PAYEE_HEADER_LEFT_OFFSET)
    payeeHeaderTopRange.mergeAcross();
    payeeHeaderTopRange.setValue('Paid For Who');
    _setHeaderStyle(payeeHeaderTopRange);


    //getRange(row, column, numRows, numColumns)
    var PAYMENT_HEADER_TEXT = [['Self Pay','Ind. Payment','Payer Collects'],['','','']];
    var PAYMENT_HEADER_LEFT = PAYEE_HEADER_LEFT + PAYEE_HEADER_LEFT_OFFSET;
    var PAYMENT_HEADER_LEFT_OFFSET = PAYMENT_HEADER_TEXT[0].length;

    var paymentHeaderRange = transactionsSheet.getRange(1,PAYMENT_HEADER_LEFT, 2,PAYMENT_HEADER_LEFT_OFFSET);
    paymentHeaderRange.mergeVertically();
    paymentHeaderRange.setValues(PAYMENT_HEADER_TEXT);
    _setHeaderStyle(paymentHeaderRange);

    //hide the first
    transactionsSheet.hideRows(3);
    transactionsSheet.setFrozenRows(2);

    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    // BODY SETUP
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

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

    var locationColumn = ENTRY_HEADER_LEFT + ENTRY_HEADER_TEXT[0].indexOf('Location');
    var locationBodyRange = transactionsSheet.getRange(BODY_TOP,locationColumn, BODY_TOP_OFFSET,1);
    _setBodyStyle(locationBodyRange);

    var itemColumn = ENTRY_HEADER_LEFT + ENTRY_HEADER_TEXT[0].indexOf('Item');
    var itemBodyRange = transactionsSheet.getRange(BODY_TOP,itemColumn, BODY_TOP_OFFSET,1);
    _setBodyStyle(itemBodyRange);

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
    workbook.setNamedRange('TRANSACTIONS_BODY_CURRENCY',currencyBodyRange);

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
    workbook.setNamedRange('TRANSACTIONS_BODY_AMOUNT_PAID',currencyBodyRange);


    //set the amount paid currency conversion
    var amountPaidUserColumn = ENTRY_HEADER_LEFT + ENTRY_HEADER_TEXT[0].indexOf('Amount Paid ('+_getUserCurrency()+')');
    var amountPaidUserBodyRange = transactionsSheet.getRange(BODY_TOP,amountPaidUserColumn, BODY_TOP_OFFSET,1);
    amountPaidUserBodyRange.setNumberFormat("$0.00");
//TODO: look at the google Finanace method and lookup a specific date.
    amountPaidUserBodyRange.setFormulaR1C1('=IF(OR(EQ("'+_getUserCurrency()+'",R[0]C[-2]),ISBLANK(R[0]C[-2])),R[0]C[-1],GOOGLEFINANCE(CONCATENATE("CURRENCY:",R[0]C[-2],"'+_getUserCurrency()+'"))*R[0]C[-1])');
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
    workbook.setNamedRange('TRANSACTIONS_BODY_WHO_PAID',whoPaidBodyRange);

    var paidForColumn = PAYEE_HEADER_LEFT;
    var paidForBodyRange = transactionsSheet.getRange(BODY_TOP,paidForColumn, BODY_TOP_OFFSET,PAYEE_HEADER_LEFT_OFFSET);
    var paidForRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Y','YS'], true)
        .setAllowInvalid(false)
        .setHelpText('When payer pays for themselves, use `YS`')
        .build();
    paidForBodyRange.setDataValidation(paidForRule);
    _setBodyStyle(paidForBodyRange);
    workbook.setNamedRange('TRANSACTIONS_BODY_PAID_FOR',paidForBodyRange);

    //R1C1 Formula helper to specify the current row payees
    var R1C1_CURRENT_ROW_PAYEES_RANGE = 'R[0]C'+PAYEE_HEADER_LEFT+':R[0]C'+(PAYEE_HEADER_LEFT+PAYEE_HEADER_LEFT_OFFSET-1)

//TODO: ask chi why this has to be so complicated, using a simpler version here.
    var selfPayColumn = PAYMENT_HEADER_LEFT + PAYMENT_HEADER_TEXT[0].indexOf('Self Pay');
    var selfPayBodyRange = transactionsSheet.getRange(BODY_TOP,selfPayColumn, BODY_TOP_OFFSET,1);
    selfPayBodyRange.setFormulaR1C1('IF(COUNTIF('+R1C1_CURRENT_ROW_PAYEES_RANGE+',"YS") > 0, "YS","")');
    _setBodyStyle(selfPayBodyRange);
    transactionsSheet.autoResizeColumn(selfPayColumn);
    workbook.setNamedRange('TRANSACTIONS_BODY_SELF_PAY',selfPayBodyRange);



    var indPaymentColumn = PAYMENT_HEADER_LEFT + PAYMENT_HEADER_TEXT[0].indexOf('Ind. Payment');
    var indPaymentBodyRange = transactionsSheet.getRange(BODY_TOP,indPaymentColumn, BODY_TOP_OFFSET,1);
    //=E3/(COUNTIF(G3:M3,"Y")+COUNTIF(G3:M3,"YS"))
    indPaymentBodyRange.setFormulaR1C1('=R[0]C'+amountPaidUserColumn+'/MAX((COUNTIF('+R1C1_CURRENT_ROW_PAYEES_RANGE+',"Y")+COUNTIF('+R1C1_CURRENT_ROW_PAYEES_RANGE+',"YS")),1)')
    _setBodyStyle(indPaymentBodyRange);
    indPaymentBodyRange.setNumberFormat("$0.00");
    workbook.setNamedRange('TRANSACTIONS_BODY_IND_PAYMENT',indPaymentBodyRange);


    var payerCollectsColumn = PAYMENT_HEADER_LEFT + PAYMENT_HEADER_TEXT[0].indexOf('Payer Collects');
    var payerCollectsBodyRange = transactionsSheet.getRange(BODY_TOP,payerCollectsColumn, BODY_TOP_OFFSET,1);
//TODO: fill this range with a formula
//TODO: ask chi why this has to be so complicated, using a simpler version here.
    //=IF(N5="YS",COUNTIF(G5:M5,"Y")*(E5/(COUNTIF(G5:M5,"Y")+1)),E5)
    payerCollectsBodyRange.setFormulaR1C1('=R[0]C[-1]*COUNTIF('+R1C1_CURRENT_ROW_PAYEES_RANGE+',"Y")')
    _setBodyStyle(payerCollectsBodyRange);
    payerCollectsBodyRange.setFontWeight("bold");
    payerCollectsBodyRange.setNumberFormat("$0.00");
    workbook.setNamedRange('TRANSACTIONS_BODY_PAYER_COLLECTS',payerCollectsBodyRange);


    return {
        whoPaidColumn: whoPaidColumn,
        paidForColumn: paidForColumn,
        indPaymentColumn: indPaymentColumn,
        payerCollectsColumn: payerCollectsColumn
    }
}

function _configureSummarySheet(workbook,summarySheet,transactionsColumns){
    var users = _getUsers();

    var BODY_TOP = 2;
    var BODY_TOP_OFFSET = users.length;

    //populate the transactions sheet.
    //The summary sheet has one block of information.
    summarySheet.activate();

    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    // HEADER SETUP
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    //getRange(row, column, numRows, numColumns)
    var SUMMARY_HEADER_TEXT = [['Name','Gets','Gives','Banker Collects','Rounded', '% Difference']];
    var SUMMARY_HEADER_LEFT = 1;
    var SUMMARY_HEADER_LEFT_OFFSET = SUMMARY_HEADER_TEXT[0].length;

    var summaryHeaderRange = summarySheet.getRange(1,SUMMARY_HEADER_LEFT,1,SUMMARY_HEADER_LEFT_OFFSET);
    summaryHeaderRange.setValues(SUMMARY_HEADER_TEXT);
    _setSubHeaderStyle(summaryHeaderRange);
    summaryHeaderRange.setFontWeight("bold");
    summaryHeaderRange.setHorizontalAlignment("left");

    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    // BODY SETUP
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    //set the date purchased validation
    var nameColumn = SUMMARY_HEADER_LEFT + SUMMARY_HEADER_TEXT[0].indexOf('Name');
    var nameBodyRange = summarySheet.getRange(BODY_TOP,nameColumn, BODY_TOP_OFFSET,1);
    var nameBodyRangeValues = [];
    var nameBodyRangeBackgrounds = [];
    for(var ndx in users){
        nameBodyRangeValues.push([users[ndx].display_name])
        nameBodyRangeBackgrounds.push([COLOR_SWATCHES[(ndx % COLOR_SWATCHES.length) -1]])
    }
    _setHeaderStyle(nameBodyRange);
    nameBodyRange.setValues(nameBodyRangeValues);
    nameBodyRange.setBackgrounds(nameBodyRangeBackgrounds);
    workbook.setNamedRange('SUMMARY_BODY_NAME',nameBodyRange);

    var getsColumn = SUMMARY_HEADER_LEFT + SUMMARY_HEADER_TEXT[0].indexOf('Gets');
    var getsBodyRange = summarySheet.getRange(BODY_TOP,getsColumn, BODY_TOP_OFFSET,1);
    //=SUMIF(Transactions!F:F,A2,Transactions!P:P)
    getsBodyRange.setFormulaR1C1('=SUMIF(Transactions!C'+transactionsColumns.whoPaidColumn+':C'+transactionsColumns.whoPaidColumn+', R[0]C[-1], Transactions!C'+transactionsColumns.payerCollectsColumn+':C'+transactionsColumns.payerCollectsColumn+')');
    _setSummaryBodyStyle(getsBodyRange);

    var givesColumn = SUMMARY_HEADER_LEFT + SUMMARY_HEADER_TEXT[0].indexOf('Gives');
    var givesBodyRange = summarySheet.getRange(BODY_TOP,givesColumn, BODY_TOP_OFFSET,1);
    var formulas = [];
    for(var ndx in users){
        formulas.push(['=SUMIF(Transactions!C'+(transactionsColumns.paidForColumn+ parseInt(ndx))+':C'+(transactionsColumns.paidForColumn+ parseInt(ndx))+', "Y", Transactions!C'+transactionsColumns.indPaymentColumn+':C'+transactionsColumns.indPaymentColumn+')'])
    }
    givesBodyRange.setFormulasR1C1(formulas);
    _setSummaryBodyStyle(givesBodyRange);

    var bankerCollectsColumn = SUMMARY_HEADER_LEFT + SUMMARY_HEADER_TEXT[0].indexOf('Banker Collects');
    var bankerCollectsBodyRange = summarySheet.getRange(BODY_TOP,bankerCollectsColumn, BODY_TOP_OFFSET,1);
    bankerCollectsBodyRange.setFormulaR1C1('=R[0]C[-2] - R[0]C[-1]');
    _setSummaryBodyStyle(bankerCollectsBodyRange);
    workbook.setNamedRange('SUMMARY_BODY_BANKER_COLLECTS',bankerCollectsBodyRange);

    var roundedColumn = SUMMARY_HEADER_LEFT + SUMMARY_HEADER_TEXT[0].indexOf('Rounded');
    var roundedBodyRange = summarySheet.getRange(BODY_TOP, roundedColumn, BODY_TOP_OFFSET,1);
    roundedBodyRange.setFormulaR1C1('=ROUND(R[0]C[-1]/5,0)*5');
    _setSummaryBodyStyle(roundedBodyRange);

    var percentDiffColumn = SUMMARY_HEADER_LEFT + SUMMARY_HEADER_TEXT[0].indexOf('% Difference');
    var percentDiffBodyRange = summarySheet.getRange(BODY_TOP, percentDiffColumn, BODY_TOP_OFFSET,1);
    percentDiffBodyRange.setFormulaR1C1('=((R[0]C[-1]/R[0]C[-2])-1)');
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
    range.setHorizontalAlignment("center");

}

function _setSummaryBodyStyle(range){
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
    var exists = false;
    for (var i = 0; i < a.length; i++) {
        for(var prop in obj){
            exists = (exists && (obj[prop] == a[i][prop]))
        }
        if(exists){
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

/**
 * Function will return true if the ranges intersect
 * @param rangeA
 * @param rangeB
 */
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

function _getUserColor(username){
    var user_ndx = _getUsers().indexOf(username);
    if(user_ndx == -1){
        return 'white'
    }
    else{
        return COLOR_SWATCHES[(user_ndx % COLOR_SWATCHES.length) -1]
    }
}