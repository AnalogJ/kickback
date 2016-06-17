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

    if(_getFlag('FLAG_WIZARD_INIT')){
        SpreadsheetApp.getUi().createAddonMenu()
            .addItem('Rerun Kickback Wizard', 'reset')
            .addItem('Add new traveller', 'add_traveller')
            .addItem('Add new trip currency', 'add_currency')
            .addToUi();
    }
    else{
        SpreadsheetApp.getUi().createAddonMenu()
            .addItem('Run Kickback Wizard', 'use')
            .addToUi();
    }

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
        Logger.log('Current Sheet is not Transactions. Skipping');
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
        Logger.log('The edited range does not intesect with a watched range. Skipping.');
        return;
    }

    //Before processing rows, ensure that the rows we process match the row we care about.
    var first_row = Math.max(e.range.getRow(),currencyRange.getRow());
    var last_row = Math.min(e.range.getLastRow(),currencyRange.getLastRow());
    Logger.log(first_row);
    Logger.log(last_row);

    for(var row = first_row; row<=last_row; row++){
        //set the background for the currency col.
        var currencyCell = currentSheet.getRange(row, currencyRange.getColumn());
        Logger.log('Currency cell value:' + currencyCell.getValue());
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
        var paidForCells = currentSheet.getRange(row, paidForRange.getColumn(), 1, paidForRange.getLastColumn()-paidForRange.getColumn()+1);
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
            bankerCollectsCellsBackgrounds.push([COLOR_LIGHT_GREEN])
        }
        else{
            bankerCollectsCellsBackgrounds.push([COLOR_LIGHT_RED])
        }
    }
    bankerCollectsCells.setBackgrounds(bankerCollectsCellsBackgrounds);
}


/**
 * Enables the add-on on for the current spreadsheet (simply by running) and
 * shows a popup informing the user of the new functions that are available.
 */
function use() {
    //ui.alert(title, message, ui.ButtonSet.OK);
    var ui = SpreadsheetApp.getUi();

    if(_isSpreadsheetEmpty()){
//        ui.alert('Welcome to the Kickback for Google Sheets wizard')

        var html = HtmlService.createHtmlOutputFromFile('view.wizard')
            .setSandboxMode(HtmlService.SandboxMode.IFRAME)
            .setWidth(500)
            .setHeight(500);
        ui.showModalDialog(html, 'Kickback Wizard');
        _setFlag('FLAG_WIZARD_INIT', true);
        ui.createAddonMenu()
            .addItem('Rerun Kickback Wizard', 'reset')
            .addItem('Add new traveller', 'add_traveller')
            .addItem('Add new trip currency', 'add_currency')
            .addToUi();

        return;
    }
    else{
        ui.alert('Unfortunately this workbook is not empty. To protect your data, we cannot run a wizard on a non empty workbook.')
        return;
    }
}

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

    _configureTransactionsSheet(workbook,transactionsSheet);
    _configureSummarySheet(workbook,summarySheet);
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
        names.push(users[ndx])
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

    //set the body defaults
    var bodyRange = transactionsSheet.getRange(BODY_TOP,1, BODY_TOP_OFFSET,ENTRY_HEADER_LEFT_OFFSET+PAYEE_HEADER_LEFT_OFFSET+PAYMENT_HEADER_LEFT_OFFSET);
    _setBodyStyle(bodyRange);
    bodyRange.setBorder(true,true,true,true,false,false);


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

    var locationColumn = ENTRY_HEADER_LEFT + ENTRY_HEADER_TEXT[0].indexOf('Location');
    var locationBodyRange = transactionsSheet.getRange(BODY_TOP,locationColumn, BODY_TOP_OFFSET,1);

    var itemColumn = ENTRY_HEADER_LEFT + ENTRY_HEADER_TEXT[0].indexOf('Item');
    var itemBodyRange = transactionsSheet.getRange(BODY_TOP,itemColumn, BODY_TOP_OFFSET,1);

    //set the currency validation.
    //getRange(row, column, numRows, numColumns)
    var currencyColumn = ENTRY_HEADER_LEFT + ENTRY_HEADER_TEXT[0].indexOf('Currency');
    var currencyBodyRange = transactionsSheet.getRange(BODY_TOP,currencyColumn, BODY_TOP_OFFSET,1);


    var currencies = [];
    currencies = currencies.concat(_getTripCurrencies());
    currencies.push(_getUserCurrency());
    currencies = _unique(currencies);
    var currencyRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(currencies, true)
        .setAllowInvalid(false)
        .setHelpText('Currency used to pay for this item.')
        .build();
    currencyBodyRange.setDataValidation(currencyRule);
    workbook.setNamedRange('TRANSACTIONS_BODY_CURRENCY',currencyBodyRange);

    //set the amount paid validation
    var amountPaidColumn = ENTRY_HEADER_LEFT + ENTRY_HEADER_TEXT[0].indexOf('Amount Paid');
    var amountPaidBodyRange = transactionsSheet.getRange(BODY_TOP,amountPaidColumn, BODY_TOP_OFFSET,1);
    workbook.setNamedRange('TRANSACTIONS_BODY_AMOUNT_PAID',amountPaidBodyRange);
    var amountPaidRule = SpreadsheetApp.newDataValidation()
        .requireNumberGreaterThan(0)
        .setAllowInvalid(false)
        .setHelpText('Amount this item was puchased for')
        .build();
    amountPaidBodyRange.setDataValidation(amountPaidRule);
    amountPaidBodyRange.setNumberFormat("$0.00");


    //set the amount paid currency conversion
    var amountPaidUserColumn = ENTRY_HEADER_LEFT + ENTRY_HEADER_TEXT[0].indexOf('Amount Paid ('+_getUserCurrency()+')');
    var amountPaidUserBodyRange = transactionsSheet.getRange(BODY_TOP,amountPaidUserColumn, BODY_TOP_OFFSET,1);
    workbook.setNamedRange('TRANSACTIONS_BODY_AMOUNT_PAID_USER',amountPaidUserBodyRange);
    amountPaidUserBodyRange.setNumberFormat("$0.00");
    _setCalculatedBodyStyle(amountPaidUserBodyRange);
//TODO: look at the google Finanace method and lookup a specific date.
    amountPaidUserBodyRange.setFormulaR1C1('=IF(OR(EQ("'+_getUserCurrency()+'",R[0]C[-2]),ISBLANK(R[0]C[-2])),R[0]C[-1],GOOGLEFINANCE(CONCATENATE("CURRENCY:",R[0]C[-2],"'+_getUserCurrency()+'"))*R[0]C[-1])');

    var whoPaidColumn =  ENTRY_HEADER_LEFT + ENTRY_HEADER_TEXT[0].indexOf('Who Paid');
    var whoPaidBodyRange = transactionsSheet.getRange(BODY_TOP,whoPaidColumn, BODY_TOP_OFFSET,1);
    workbook.setNamedRange('TRANSACTIONS_BODY_WHO_PAID',whoPaidBodyRange);
    var whoPaidRule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(payeeHeaderBottomRange, true)
        .setAllowInvalid(false)
        .setHelpText('The user who paid for this item.')
        .build();
    whoPaidBodyRange.setDataValidation(whoPaidRule);
    whoPaidBodyRange.setFontSize(9);
    whoPaidBodyRange.setFontWeight('bold');

    var paidForColumn = PAYEE_HEADER_LEFT;
    var paidForBodyRange = transactionsSheet.getRange(BODY_TOP,paidForColumn, BODY_TOP_OFFSET,PAYEE_HEADER_LEFT_OFFSET);
    workbook.setNamedRange('TRANSACTIONS_BODY_PAID_FOR',paidForBodyRange);
    var paidForRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Y','YS',''], true)
        .setAllowInvalid(false)
        .setHelpText('When payer pays for themselves, use `YS`')
        .build();
    paidForBodyRange.setDataValidation(paidForRule);


//TODO: ask chi why this has to be so complicated, using a simpler version here.
    var selfPayColumn = PAYMENT_HEADER_LEFT + PAYMENT_HEADER_TEXT[0].indexOf('Self Pay');
    var selfPayBodyRange = transactionsSheet.getRange(BODY_TOP,selfPayColumn, BODY_TOP_OFFSET,1);
    workbook.setNamedRange('TRANSACTIONS_BODY_SELF_PAY',selfPayBodyRange);
    selfPayBodyRange.setFormulaR1C1(_generateSelfPayFormulaR1C1());
    _setCalculatedBodyStyle(selfPayBodyRange);
    transactionsSheet.autoResizeColumn(selfPayColumn);

    var indPaymentColumn = PAYMENT_HEADER_LEFT + PAYMENT_HEADER_TEXT[0].indexOf('Ind. Payment');
    var indPaymentBodyRange = transactionsSheet.getRange(BODY_TOP,indPaymentColumn, BODY_TOP_OFFSET,1);
    workbook.setNamedRange('TRANSACTIONS_BODY_IND_PAYMENT',indPaymentBodyRange);
    //=E3/(COUNTIF(G3:M3,"Y")+COUNTIF(G3:M3,"YS"))
    indPaymentBodyRange.setFormulaR1C1(_generateIndPaymentFormulaR1C1());
    indPaymentBodyRange.setNumberFormat("$0.00");
    _setCalculatedBodyStyle(indPaymentBodyRange);


    var payerCollectsColumn = PAYMENT_HEADER_LEFT + PAYMENT_HEADER_TEXT[0].indexOf('Payer Collects');
    var payerCollectsBodyRange = transactionsSheet.getRange(BODY_TOP,payerCollectsColumn, BODY_TOP_OFFSET,1);
    workbook.setNamedRange('TRANSACTIONS_BODY_PAYER_COLLECTS',payerCollectsBodyRange);
    //=IF(N5="YS",COUNTIF(G5:M5,"Y")*(E5/(COUNTIF(G5:M5,"Y")+1)),E5)
    payerCollectsBodyRange.setFormulaR1C1(_generatePayerCollectsFormulaR1C1())
    payerCollectsBodyRange.setFontWeight("bold");
    payerCollectsBodyRange.setNumberFormat("$0.00");
    _setCalculatedBodyStyle(payerCollectsBodyRange);

}

function _configureSummarySheet(workbook,summarySheet){
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

    //set the body defaults
    var bodyRange = summarySheet.getRange(BODY_TOP,1, BODY_TOP_OFFSET,SUMMARY_HEADER_LEFT_OFFSET);
    _setSummaryBodyStyle(bodyRange);
    bodyRange.setBorder(true,true,true,true,false,false);

    //set the date purchased validation
    var nameColumn = SUMMARY_HEADER_LEFT + SUMMARY_HEADER_TEXT[0].indexOf('Name');
    var nameBodyRange = summarySheet.getRange(BODY_TOP,nameColumn, BODY_TOP_OFFSET,1);
    var nameBodyRangeValues = [];
    var nameBodyRangeBackgrounds = [];
    for(var ndx in users){
        nameBodyRangeValues.push([users[ndx]]);
        nameBodyRangeBackgrounds.push([COLOR_SWATCHES[ndx % COLOR_SWATCHES.length]])
    }
    _setHeaderStyle(nameBodyRange);
    nameBodyRange.setValues(nameBodyRangeValues);
    nameBodyRange.setBackgrounds(nameBodyRangeBackgrounds);
    workbook.setNamedRange('SUMMARY_BODY_NAME',nameBodyRange);

    var getsColumn = SUMMARY_HEADER_LEFT + SUMMARY_HEADER_TEXT[0].indexOf('Gets');
    var getsBodyRange = summarySheet.getRange(BODY_TOP,getsColumn, BODY_TOP_OFFSET,1);
    //=SUMIF(Transactions!F:F,A2,Transactions!P:P)
    getsBodyRange.setFormulaR1C1(_generateGetsFormulaR1C1());
    getsBodyRange.setNumberFormat("$0.00");
    _setCalculatedBodyStyle(getsBodyRange);

    var givesColumn = SUMMARY_HEADER_LEFT + SUMMARY_HEADER_TEXT[0].indexOf('Gives');
    var givesBodyRange = summarySheet.getRange(BODY_TOP,givesColumn, BODY_TOP_OFFSET,1);
    var formulas = [];
    for(var ndx in users){
        formulas.push([_generateGivesFormulaR1C1(ndx)])
    }
    givesBodyRange.setFormulasR1C1(formulas);
    givesBodyRange.setNumberFormat("$0.00");
    _setCalculatedBodyStyle(givesBodyRange);

    var bankerCollectsColumn = SUMMARY_HEADER_LEFT + SUMMARY_HEADER_TEXT[0].indexOf('Banker Collects');
    var bankerCollectsBodyRange = summarySheet.getRange(BODY_TOP,bankerCollectsColumn, BODY_TOP_OFFSET,1);
    bankerCollectsBodyRange.setFormulaR1C1(_generateBankerCollectsFormulaR1C1());
    workbook.setNamedRange('SUMMARY_BODY_BANKER_COLLECTS',bankerCollectsBodyRange);
    bankerCollectsBodyRange.setNumberFormat("$0.00");
    bankerCollectsBodyRange.setFontSize(14);
    bankerCollectsBodyRange.setFontWeight('bold');


    var roundedColumn = SUMMARY_HEADER_LEFT + SUMMARY_HEADER_TEXT[0].indexOf('Rounded');
    var roundedBodyRange = summarySheet.getRange(BODY_TOP, roundedColumn, BODY_TOP_OFFSET,1);
    roundedBodyRange.setFormulaR1C1(_generateRoundedFormulaR1C1());
    roundedBodyRange.setNumberFormat("$0.00");
    _setCalculatedBodyStyle(roundedBodyRange);

    var percentDiffColumn = SUMMARY_HEADER_LEFT + SUMMARY_HEADER_TEXT[0].indexOf('% Difference');
    var percentDiffBodyRange = summarySheet.getRange(BODY_TOP, percentDiffColumn, BODY_TOP_OFFSET,1);
    //=IF(NOT(D4=0),((E4/D4)-1),0)
    percentDiffBodyRange.setFormulaR1C1(_generateDifferenceFormulaR1C1());
    percentDiffBodyRange.setFontSize(9);
    percentDiffBodyRange.setNumberFormat("0.00%");
    _setCalculatedBodyStyle(percentDiffBodyRange);

}


//*************************************************************************************************
// Column Formula generators
//*************************************************************************************************

function _generateCurrentRowPayeesRangeFormulaR1C1(){
    //R1C1 Formula helper to specify the current row payees
    // var R1C1_CURRENT_ROW_PAYEES_RANGE = 'R[0]C'+PAYEE_HEADER_LEFT+':R[0]C'+(PAYEE_HEADER_LEFT+PAYEE_HEADER_LEFT_OFFSET-1)

    var workbook = SpreadsheetApp.getActiveSpreadsheet();
    var paidForRange = workbook.getRangeByName('TRANSACTIONS_BODY_PAID_FOR');
    return 'R[0]C'+paidForRange.getColumn()+':R[0]C'+ paidForRange.getLastColumn();
}

function _generateSelfPayFormulaR1C1(){
    return 'IF(COUNTIF('+_generateCurrentRowPayeesRangeFormulaR1C1()+',"YS") > 0, "YS","")'
}

function _generateIndPaymentFormulaR1C1(){
    var workbook = SpreadsheetApp.getActiveSpreadsheet();
    var amountPaidUserColumn = workbook.getRangeByName('TRANSACTIONS_BODY_AMOUNT_PAID_USER').getColumn();

    var currentRowPayeesRangeFormulaR1C1 =  _generateCurrentRowPayeesRangeFormulaR1C1();

    return '=R[0]C'+amountPaidUserColumn+'/MAX((COUNTIF('+currentRowPayeesRangeFormulaR1C1+',"Y")+COUNTIF('+currentRowPayeesRangeFormulaR1C1+',"YS")),1)'
}

function _generatePayerCollectsFormulaR1C1(){
    return '=R[0]C[-1]*COUNTIF('+_generateCurrentRowPayeesRangeFormulaR1C1()+',"Y")'
}


function _generateGetsFormulaR1C1(){
    var workbook = SpreadsheetApp.getActiveSpreadsheet();
    var whoPaidColumn = workbook.getRangeByName('TRANSACTIONS_BODY_WHO_PAID').getColumn();
    var payerCollectsColumn = workbook.getRangeByName('TRANSACTIONS_BODY_PAYER_COLLECTS').getColumn();
    return '=SUMIF(Transactions!C'+whoPaidColumn+':C'+whoPaidColumn+', R[0]C[-1], Transactions!C'+payerCollectsColumn+':C'+payerCollectsColumn+')'
}

function _generateGivesFormulaR1C1(ndx){
    var workbook = SpreadsheetApp.getActiveSpreadsheet();
    var paidForColumn = workbook.getRangeByName('TRANSACTIONS_BODY_PAID_FOR').getColumn();
    var indPaymentColumn = workbook.getRangeByName('TRANSACTIONS_BODY_IND_PAYMENT').getColumn();

    return '=SUMIF(Transactions!C'+(paidForColumn+ parseInt(ndx))+':C'+(paidForColumn+ parseInt(ndx))+', "Y", Transactions!C'+indPaymentColumn+':C'+indPaymentColumn+')'
}

function _generateBankerCollectsFormulaR1C1(){
    return '=R[0]C[-2] - R[0]C[-1]'
}

function _generateRoundedFormulaR1C1(){
    return '=ROUND(R[0]C[-1]/5,0)*5'
}

function _generateDifferenceFormulaR1C1(){
    return '=IF(NOT(R[0]C[-2]=0),((R[0]C[-1]/R[0]C[-2])-1),0)'
}


//*************************************************************************************************
// Popup/Modal Handlers.
//*************************************************************************************************
//TODO
function add_traveller(){
    var ui = SpreadsheetApp.getUi();
    var html = HtmlService.createHtmlOutputFromFile('view.add_traveller')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setWidth(500)
        .setHeight(200);
    ui.showModalDialog(html, 'Kickback Add Traveller');
}

//TODO
function add_traveller_submit(form_data){
    var data = JSON.parse(form_data);

    Logger.log(data);


    /*
    * We have to do the following changes in the Transactions Sheet
    * - add the user to the users settings
    * - modify the who paid column validation to include the new user
    * - add a new user under "Paid For Who" section with a new column + validation
    * - Update teh Self Pay formula with the new range.
    * - Update the Ind Paymnet forumen with the new range.
    * - Update the Payer Collections forumla with the new range.
    *
    * We have to make the following changes in the Summary Sheet
    * - Add a new row with correct formulas.
    * */


    //add user
    _addUser(data["traveller"])

    var workbook = SpreadsheetApp.getActiveSpreadsheet();

    //modify the who paid validation rules with new traveller
    var whoPaidRange = workbook.getRangeByName('TRANSACTIONS_BODY_WHO_PAID');
    var whoPaidRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(_getUsers(), true)
        .setAllowInvalid(false)
        .setHelpText('The user who paid for this item.')
        .build();
    whoPaidRange.setDataValidation(whoPaidRule);

    //modify the paid for who row with new column at the end.
    var transactionsSheet = workbook.getSheetByName("Transactions");
    transactionsSheet.activate();

    var paidForRange = workbook.getRangeByName('TRANSACTIONS_BODY_PAID_FOR');

    //adding a new column at the end of the current paidforrange (paid for range needs to be updated after htis)
    transactionsSheet.insertColumnAfter(paidForRange.getLastColumn());
    paidForRange = transactionsSheet.getRange(
        paidForRange.getRow(),
        paidForRange.getColumn(),
        paidForRange.getNumRows(),
        paidForRange.getNumColumns() + 1
    );
    workbook.setNamedRange('TRANSACTIONS_BODY_PAID_FOR',paidForRange);

    //set the value of the new header cell
    transactionsSheet.getRange(2, paidForRange.getLastColumn()).setValue(data["traveller"]);

    var payeeHeaderTopRange = transactionsSheet.getRange(1, paidForRange.getColumn(), 1, paidForRange.getNumColumns());
    payeeHeaderTopRange.mergeAcross();
    _setSubHeaderStyle(payeeHeaderTopRange);

    //update Self Pay formulas
    var selfPayBodyRange = workbook.getRangeByName('TRANSACTIONS_BODY_SELF_PAY');
    selfPayBodyRange.setFormulaR1C1(_generateSelfPayFormulaR1C1());

    //update Ind. Payment formulas
    var indPaymentBodyRange = workbook.getRangeByName('TRANSACTIONS_BODY_IND_PAYMENT');
    indPaymentBodyRange.setFormulaR1C1(_generateIndPaymentFormulaR1C1());

    //update Payer Collects formulas
    var payerCollectsBodyRange = workbook.getRangeByName('TRANSACTIONS_BODY_PAYER_COLLECTS');
    payerCollectsBodyRange.setFormulaR1C1(_generatePayerCollectsFormulaR1C1());

    //modify the summary name column with new traveller row
    var summarySheet = workbook.getSheetByName("Summary");
    _configureSummarySheet(workbook, summarySheet);
    transactionsSheet.activate();

}

//TODO
function add_currency(){
    var ui = SpreadsheetApp.getUi();
    var html = HtmlService.createHtmlOutputFromFile('view.add_currency')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setWidth(500)
        .setHeight(200);
    ui.showModalDialog(html, 'Kickback Add Currency');
}
//TODO
function add_currency_submit(form_data){
    var data = JSON.parse(form_data);

    //add currency
    _addTripCurrency(data['trip_currency']);

    //reconfigure the validation rules with the new sheet
    var workbook = SpreadsheetApp.getActiveSpreadsheet();
    var currencyRange = workbook.getRangeByName('TRANSACTIONS_BODY_CURRENCY');

    var currencies = [];
    currencies = currencies.concat(_getTripCurrencies())
    currencies.push(_getUserCurrency());
    currencies = _unique(currencies);
    var currencyRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(currencies, true)
        .setAllowInvalid(false)
        .setHelpText('Currency used to pay for this item.')
        .build();
    currencyRange.setDataValidation(currencyRule);
}


function reset(){
    var workbook = SpreadsheetApp.getActiveSpreadsheet();

    //delete all sheets.
    var sheets = workbook.getSheets();

    if (sheets.length) {
        sheets[0].clear();
        sheets[0].setName('DELETING');
        for (var ndx = 1; ndx < sheets.length; ndx++) {
            workbook.deleteSheet(sheets[ndx])
        }
    }

    //clear sheet settings
    _clearUserCurrency();
    _clearUsers();
    _clearTripCurrencies();

    var ui = SpreadsheetApp.getUi();
    var html = HtmlService.createHtmlOutputFromFile('view.wizard')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setWidth(500)
        .setHeight(500);
    ui.showModalDialog(html, 'Kickback Wizard');
}

function wizard_submit(form_data){
    Logger.log(form_data)

    var settings = JSON.parse(form_data);

    _setTripCurrencies(settings['trip_currencies[]']);
    _setUserCurrency(settings['trav_currency']);

    _clearUsers()
    for(var ndx in settings["traveller[]"]){
        _addUser(settings["traveller[]"][ndx]);
    }
    _populateWorkbook()
}