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
  }
}
