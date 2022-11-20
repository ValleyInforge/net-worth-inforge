function onEdit(e) {
  
  var range = e.range;
  var column = range.getColumn();
  if(column === 4 || column === 13){
    var spreadSheet = e.source;
    var sheetName = spreadSheet.getActiveSheet().getName();
    var row = range.getRow();
    var value = e.value;
    var returnValues = [];
    var banks = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu").getRange(3,5, 5).getValues();
    // CONVERTO L'ARRAY 2D A 1D
    banks = banks.reduce(function(prev, next) {
      return prev.concat(next);
    });
    banks = banks.filter(n => n);

    if(banks.includes(sheetName) && column === 4){

      var ss= SpreadsheetApp.getActiveSpreadsheet();
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

      var ss= SpreadsheetApp.getActiveSpreadsheet();
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
  // VERIFICARE CHE NON SERVA ED ELIMINARE
  /*else if(banks.includes(sheetName) && column === 5){
    var ss= SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getSheetByName("» Expenses: Cat./Subcat.");
    var lastRowData = dataSheet.getLastRow();
    var mainSheet = ss.getSheetByName(sheetName);
    for(var i = 1; i <= lastRowData; i++){
     if(dataSheet.getRange(i, 1).getValue() == mainSheet.getRange(row, column - 1).getValue() && dataSheet.getRange(i, 2).getValue() == mainSheet.getRange(row, column).getValue() && dataSheet.getRange(i, 3).getValue() === true){
       mainSheet.getRange(row, column + 2).setValue(1);
     }else if(dataSheet.getRange(i, 3).getValue() === false && mainSheet.getRange(row, column + 2).getValue() === 1){
       mainSheet.getRange(row, column + 2).setValue(0);
     }  
    }
  }else if(banks.includes(sheetName) && column === 14){
    var ss= SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getSheetByName("» Income: Cat./Subcat.");
    var lastRowData = dataSheet.getLastRow();
    var mainSheet = ss.getSheetByName(sheetName);
    for(var i = 1; i <= lastRowData; i++){
     if(dataSheet.getRange(i, 1).getValue() == mainSheet.getRange(row, column - 1).getValue() && dataSheet.getRange(i, 2).getValue() == mainSheet.getRange(row, column).getValue() && dataSheet.getRange(i, 3).getValue() === true){
       mainSheet.getRange(row, column - 5).setValue(1);
     }else if(dataSheet.getRange(i, 3).getValue() === false && mainSheet.getRange(row, column - 5).getValue() === 1){
       mainSheet.getRange(row, column - 5).setValue(0);
     }
    }
  }*/
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
  var banks = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu").getRange(3,5, 5).getValues();
  // CONVERTO L'ARRAY 2D A 1D
  banks = banks.reduce(function(prev, next) {
    return prev.concat(next);
  });
  banks = banks.filter(n => n);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  // Serve ad individuare la giusta posizione per la categoria successiva quando si genera la tabella. Impostato di partenza a 0.
  var categorySpace = 0;
  for(k = 1; k <= banks.length; k++){
    if(k === 1){
      SpreadsheetApp.getUi().alert("uno");
      var firstCatStats = categoryStats(type, settingsSheet, statsSheet, categorySpace, banks[k-1]);
      var firstSubCatStats = subCategoryStats(type, settingsSheet, statsSheet, firstCatStats, banks[k-1]);
    }else if(k > 1){
      SpreadsheetApp.getUi().alert(firstSubCatStats);
      var nextCatStats = categoryStats(type, settingsSheet, statsSheet, (firstSubCatStats+2), banks[k-1]);
      subCategoryStats(type, settingsSheet, statsSheet, nextCatStats, banks[k-1]);
    }    
  };
}
