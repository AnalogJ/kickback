//*************************************************************************************************
// Style functions
//*************************************************************************************************
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
    range.setFontFamily(DEFAULT_FONT_FAMILY);
    range.setFontSize(7);
    range.setWrap(true);
    range.setHorizontalAlignment("center");

}

function _setSummaryBodyStyle(range){
    range.setFontFamily(DEFAULT_FONT_FAMILY);
    range.setFontSize(12);
}


//*************************************************************************************************
// Document Storage functions
//*************************************************************************************************
//USERS Storage
function _addUser(username){
    var users = _getUsers();
    users.push(username)
    _setUsers(users);
}

function _clearUsers(){
    var documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('USERS', JSON.stringify([]));
}

function _getUsers(){
    var documentProperties = PropertiesService.getDocumentProperties();
    var users_str = documentProperties.getProperty('USERS');
    Logger.log(users_str)

    if(users_str){
        return  JSON.parse(users_str)
    }
    return [];
}

function  _setUsers(users_array){
    var documentProperties = PropertiesService.getDocumentProperties();
    return documentProperties.setProperty('USERS', JSON.stringify(users_array ? _unique(users_array) : []));
}

//TRIP_CURRENCIES Storage
function _addTripCurrency(currency){
    var trip_currencies = _getTripCurrencies();
    trip_currencies.push(currency)
    _setTripCurrencies(trip_currencies);
}

function _clearTripCurrencies(){
    var documentProperties = PropertiesService.getDocumentProperties();
    return documentProperties.setProperty('TRIP_CURRENCIES', JSON.stringify([]));
}

function  _getTripCurrencies(){
    var documentProperties = PropertiesService.getDocumentProperties();
    var trip_currencies_str = documentProperties.getProperty('TRIP_CURRENCIES');
    Logger.log(trip_currencies_str)

    if(trip_currencies_str){
        return  JSON.parse(trip_currencies_str)
    }
    return [];
}

function  _setTripCurrencies(currency_array){
    var documentProperties = PropertiesService.getDocumentProperties();
    return documentProperties.setProperty('TRIP_CURRENCIES', JSON.stringify(currency_array ? _unique(currency_array) : []));
}

//USER_CURRENCY Storage
function _clearUserCurrency(){
    var documentProperties = PropertiesService.getDocumentProperties();
    return documentProperties.setProperty('USER_CURRENCY', '');
}

function  _getUserCurrency(){
    var documentProperties = PropertiesService.getDocumentProperties();
    return documentProperties.getProperty('USER_CURRENCY') || '';
}

function  _setUserCurrency(currency_code){
    var documentProperties = PropertiesService.getDocumentProperties();
    return documentProperties.setProperty('USER_CURRENCY',currency_code||'');
}

function _setTransactionRange(keyword, range){

}

function _setSummaryRange(keyword, range){

}

function _getTransactionRange(keyword){}
function _getSummaryRange(keyword){}

//*************************************************************************************************
// Utility functions
//*************************************************************************************************

function _unique(array){
    //make sure we only show unique currencies
    function onlyUnique(value, index, self) {
        return self.indexOf(value) === index;
    }
    return array.filter( onlyUnique );
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
        return COLOR_SWATCHES[user_ndx % COLOR_SWATCHES.length]
    }
}