// Funzione che organizza le operazioni di aggiornamento delle tabelle con le somme mensili per singolo conto. 
function updateMonthlyIncomeExpenses(){
  // Se nel menu è impostato di aggiornare le uscite..  
  if(type == "expenses"){
    var statsSheet = ss.getSheetByName("Monthly Expenses / Bank account");
    var settingsSheet = ss.getSheetByName("» Expenses: Cat./Subcat.");
  // ..altrimenti se è impostato di aggiornare le entrate..
  }else if(type == "income"){
    var statsSheet = ss.getSheetByName("Monthly Income / Bank account");
    var settingsSheet = ss.getSheetByName("» Income: Cat./Subcat.");
  // ..altrimenti c'è un problema.
  }else{
    SpreadsheetApp.getUi().alert("Ops, there was a problem! I can't understand the 'type' of the request");
    return false;
  }
  // Cancella il vecchio contenuto del foglio.
  statsSheet.getDataRange().clear();
  // Scopre tutte le righe nascoste
  var rRows = statsSheet.getRange("A:A");
  statsSheet.unhideRow(rRows);
  // Per ogni sheet dei conti..
  for(k = 1; k <= banks.length; k++){
    // ..se è il primo..
    if(k === 1){
      // Genera la tabella per categoria.
      var firstCatStats = categoryStats(type, settingsSheet, statsSheet, categorySpace, banks[k-1], initialRowSpace, initialColumnSpace, firstLineCategory);
      // Genera la tabella per sottocategoria
      var firstSubCatStats = subCategoryStats(type, settingsSheet, statsSheet, firstCatStats, banks[k-1], lastLineCategory, initialColumnSpace);
    // ..se non è il primo.
    }else if(k > 1){
      // Genera la tabella per categoria.
      var nextCatStats = categoryStats(type, settingsSheet, statsSheet, (firstSubCatStats+2), banks[k-1], initialRowSpace, initialColumnSpace, firstLineCategory);
      // Genera la tabella per sottocategoria
      subCategoryStats(type, settingsSheet, statsSheet, nextCatStats, banks[k-1], lastLineCategory, initialColumnSpace);
    }    
  };
}
