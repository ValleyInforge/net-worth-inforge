function onEdit(e) {
  
  var range = e.range;
  var spreadSheet = e.source;
  var sheetName = spreadSheet.getActiveSheet().getName();
  var column = range.getColumn();
  var row = range.getRow();
  var value = e.value;
  var returnValues = [];
  var returnValuesTwo = [];
  
  if(('Hype - I/E' === sheetName || 'N26 - I/E' === sheetName || 'Satispay - I/E' === sheetName) && [4].includes(column)){

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

  }else if(('Hype - I/E' === sheetName || 'N26 - I/E' === sheetName || 'Satispay - I/E' === sheetName) && [13].includes(column)){

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

  }else if(('Hype - I/E' === sheetName || 'N26 - I/E' === sheetName || 'Satispay - I/E' === sheetName) && [5].includes(column)){
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
  }else if(('Hype - I/E' === sheetName || 'N26 - I/E' === sheetName || 'Satispay - I/E' === sheetName) && [14].includes(column)){
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
  }
  SpreadsheetApp.getActive().toast("Lo script ha finito le modifiche automatiche, prosegui pure :)", "Script terminato", 3);
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
  // Genero stats mensili per Hype
  var hypeCatStart = categoryStats(type, settingsSheet, statsSheet, categorySpace, 'Hype');
  var hypeSubCatStart = subCategoryStats(type, settingsSheet, statsSheet, hypeCatStart, 'Hype');
  // Genero stats mensili per N26
  var n26CatStart = categoryStats(type, settingsSheet, statsSheet, hypeSubCatStart+2, 'N26');
  var n26SubCatStart = subCategoryStats(type, settingsSheet, statsSheet, n26CatStart, 'N26');
}
