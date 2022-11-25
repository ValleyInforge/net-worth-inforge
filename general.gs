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
