// Valley - inforge.net //
// Devo fare un array con settingsValue.length dal quale togliere il numero delle righe che non rispetta settingsValue.length[x][3] === false e poi con le restanti faccio il resto
function updateMonthlyIncomeExpensesTotal(){
  
  var type = ss.getSheetByName("Menu").getRange(3,2).getValue();
  var subCategoryCleaned = [];
  var z = 0;
  var rangeToModify = [];
  
  if(type == "expenses"){
    
    var statsSheet = ss.getSheetByName("Monthly Expenses - Total");
    var settingsSheet = ss.getSheetByName("» Expenses: Cat./Subcat.");
    var expensesCategoryList = ss.getSheetByName("» Expenses: Cat./Subcat.");
    var lastRowDataExpenses = expensesCategoryList.getLastRow();
  
    SpreadsheetApp.getActive().toast("Mi preparo per l'aggiornamento degli sheet dei conti..", "Script in esecuzione", -1);
    for(var i = 1; i <= lastRowDataExpenses; i++){
      if(expensesCategoryList.getRange(i, 3).getValue() === true){
        subCategoryCleaned[z] = { "category": expensesCategoryList.getRange(i, 1).getValue(), "subcategory": expensesCategoryList.getRange(i, 2).getValue() };
        z++;
      }
    }
    banks.forEach(function(item){
      SpreadsheetApp.getActive().toast("Procedo ad aggiornare lo sheet del conto "+ item +" con i valori delle uscite da non considerare nella conta totale..", "Script in esecuzione", -1);
      // Resetta i campi impostandoli tutti a zero
      ss.getSheetByName(item).getRange(4, 7, ss.getSheetByName(item).getRange("C4:C").getValues().filter(String).length).setValue(0);
      for(var m = 1; m <= (ss.getSheetByName(item).getRange("C4:C").getValues().filter(String).length); m++){
        subCategoryCleaned.some(element => {
          if (element.category === ss.getSheetByName(item).getRange(m + 3, 4).getValue() && element.subcategory === ss.getSheetByName(item).getRange(m + 3, 5).getValue()) {
            rangeToModify.push("G"+ (m + 3)+"");
          }
        });
      }
      if(rangeToModify.length !== 0){
        ss.getSheetByName(item).getRangeList(rangeToModify).setValue(1);
        rangeToModify = [];
      }
    });
  
  }else if(type == "income"){
    
    var statsSheet = ss.getSheetByName("Monthly Income - Total");
    var settingsSheet = ss.getSheetByName("» Income: Cat./Subcat.");
    var incomeCategoryList = ss.getSheetByName("» Income: Cat./Subcat.");
    var lastRowDataIncome = incomeCategoryList.getLastRow();
    
    SpreadsheetApp.getActive().toast("Mi preparo per l'aggiornamento degli sheet dei conti..", "Script in esecuzione", -1);
    for(var i = 1; i <= lastRowDataIncome; i++){
      if(incomeCategoryList.getRange(i, 3).getValue() === true){
        subCategoryCleaned[z] = { "category": incomeCategoryList.getRange(i, 1).getValue(), "subcategory": incomeCategoryList.getRange(i, 2).getValue() };
        z++;
      }
    }
    banks.forEach(function(item){
      SpreadsheetApp.getActive().toast("Procedo ad aggiornare lo sheet del conto "+ item +" con i valori delle uscite da non considerare nella conta totale..", "Script in esecuzione", -1);
      // Resetta i campi impostandoli tutti a zero
      ss.getSheetByName(item).getRange(4, 9, ss.getSheetByName(item).getRange("L4:L").getValues().filter(String).length).setValue(0);
      for(var m = 1; m <= (ss.getSheetByName(item).getRange("L4:L").getValues().filter(String).length); m++){
        subCategoryCleaned.some(element => {
          if (element.category === ss.getSheetByName(item).getRange(m + 3, 13).getValue() && element.subcategory === ss.getSheetByName(item).getRange(m + 3, 14).getValue()) {
            rangeToModify.push("I"+ (m + 3)+"");
          }
        });
      }
      if(rangeToModify.length !== 0){
        ss.getSheetByName(item).getRangeList(rangeToModify).setValue(1);
        rangeToModify = [];
      }
    });

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
  // Genero stats mensili totali
  SpreadsheetApp.getActive().toast("Genero la tabella delle statistiche mensili totali..", "Script in esecuzione", -1);
  var hypeCatStartTotal = categoryStatsTotal(type, settingsSheet, statsSheet, categorySpace, initialRowSpace, initialColumnSpace, firstLineCategory);
  subCategoryStatsTotal(type, settingsSheet, statsSheet, hypeCatStartTotal, lastLineCategory, initialColumnSpace);
  SpreadsheetApp.getActive().toast("Finito!", "Script terminato", 4);
}
