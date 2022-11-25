// Valley - inforge.net //
function onEdit(e) {
  
  var range = e.range;
  var column = range.getColumn();
  if(column === 4 || column === 13){
    var spreadSheet = e.source;
    var sheetName = spreadSheet.getActiveSheet().getName();
    var row = range.getRow();
    var value = e.value;
    var returnValues = [];

    if(banks.includes(sheetName) && column === 4){

      var dataSheet = ss.getSheetByName("» Expenses: Cat./Subcat.");
      var mainSheet = ss.getSheetByName(sheetName);
      var lastRowData = dataSheet.getLastRow();   
      for(var i = 1; i <= lastRowData; i++){
        if(value == dataSheet.getRange(i, 1).getValue()){
          returnValues.push(dataSheet.getRange(i, 2).getValue());      
        }
      }
      mainSheet.getRange(row, column+1).clear();
      var dropdown = mainSheet.getRange(row, column+1);
      var rule = SpreadsheetApp.newDataValidation().requireValueInList(returnValues).build();
      dropdown.setDataValidation(rule);

    }else if(banks.includes(sheetName) && column === 13){

      var dataSheet = ss.getSheetByName("» Income: Cat./Subcat.");
      var mainSheet = ss.getSheetByName(sheetName);
      var lastRowData = dataSheet.getLastRow();    
      for(var i = 1; i <= lastRowData; i++){
        if(value == dataSheet.getRange(i, 1).getValue()){
          returnValues.push(dataSheet.getRange(i, 2).getValue());      
        }
      }
      mainSheet.getRange(row, column+1).clear();
      var dropdown = mainSheet.getRange(row, column+1);
      var rule = SpreadsheetApp.newDataValidation().requireValueInList(returnValues).build();
      dropdown.setDataValidation(rule);

    }
  }
}


function columnToLetter(column){
  var temp, letter = '';
  while (column > 0){
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}


function letterToColumn(letter){
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++){
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}


function updateMonthlyIncomeExpenses(){
  
  var type = ss.getSheetByName("Menu").getRange(3,2).getValue();
  
  if(type == "expenses"){
    var statsSheet = ss.getSheetByName("Monthly Expenses / Bank account");
    var settingsSheet = ss.getSheetByName("» Expenses: Cat./Subcat.");
  }else if(type == "income"){
    var statsSheet = ss.getSheetByName("Monthly Income / Bank account");
    var settingsSheet = ss.getSheetByName("» Income: Cat./Subcat.");
  }else{
    SpreadsheetApp.getUi().alert("Ops, there was a problem! I can't understand the 'type' of the request");
    return false;
  }
  // Cancella il vecchio contenuto del foglio.
  statsSheet.getDataRange().clear();
  // Scopre tutte le righe nascoste
  var rRows = statsSheet.getRange("A:A");
  statsSheet.unhideRow(rRows);
  for(k = 1; k <= banks.length; k++){
    if(k === 1){
      var firstCatStats = categoryStats(type, settingsSheet, statsSheet, categorySpace, banks[k-1], initialRowSpace, initialColumnSpace, firstLineCategory);
      var firstSubCatStats = subCategoryStats(type, settingsSheet, statsSheet, firstCatStats, banks[k-1], lastLineCategory, initialColumnSpace);
    }else if(k > 1){
      var nextCatStats = categoryStats(type, settingsSheet, statsSheet, (firstSubCatStats+2), banks[k-1], initialRowSpace, initialColumnSpace, firstLineCategory);
      subCategoryStats(type, settingsSheet, statsSheet, nextCatStats, banks[k-1], lastLineCategory, initialColumnSpace);
    }    
  };
}
