// Funzione che organizza le operazioni di aggiornamento delle tabelle con le somme mensili totali. 
function updateMonthlyIncomeExpensesTotal(){
  
  var subCategoryCleaned = [];
  var z = 0;
  var rangeToModify = [];
  
  // Se nel menu è impostato di aggiornare le uscite..  
  if(type == "expenses"){
    
    var statsSheet = ss.getSheetByName("Monthly Expenses - Total");
    var settingsSheet = ss.getSheetByName("» Expenses: Cat./Subcat.");
    var expensesCategoryList = ss.getSheetByName("» Expenses: Cat./Subcat.");
    var lastRowDataExpenses = expensesCategoryList.getLastRow();
    // Finestra che informa lo stato della funzione in corso
    SpreadsheetApp.getActive().toast("Mi preparo per l'aggiornamento degli sheet dei conti..", "Script in esecuzione", -1);
    // Per ogni sottocategoria...
    for(var i = 1; i <= lastRowDataExpenses; i++){
      // ...se la sottocategoria ha nella terza colonna "true" va considerata tra quelle da escludere nel conteggio totale, quindi viene messa dentro la var subCategoryCleaned.
      if(expensesCategoryList.getRange(i, 3).getValue() === true){
        subCategoryCleaned[z] = { "category": expensesCategoryList.getRange(i, 1).getValue(), "subcategory": expensesCategoryList.getRange(i, 2).getValue() };
        z++;
      }
    }
    // Per ogni conto..
    banks.forEach(function(item){
      // Finestra che informa lo stato della funzione in corso
      SpreadsheetApp.getActive().toast("Procedo ad aggiornare lo sheet del conto "+ item +" con i valori delle uscite da non considerare nella conta totale..", "Script in esecuzione", -1);
      // Resetta i campi impostandoli tutti a zero
      ss.getSheetByName(item).getRange(4, 7, ss.getSheetByName(item).getRange("C4:C").getValues().filter(String).length).setValue(0);
      // Per ogni record presente nello sheet del conto...
      for(var m = 1; m <= (ss.getSheetByName(item).getRange("C4:C").getValues().filter(String).length); m++){
        // ... se la categoria è presente tra quelle dentro subcategoryCleaned aggiunge la riga dello sheet del conto in questione nella variabile rangeToModify.
        subCategoryCleaned.some(element => {
          if (element.category === ss.getSheetByName(item).getRange(m + 3, 4).getValue() && element.subcategory === ss.getSheetByName(item).getRange(m + 3, 5).getValue()) {
            rangeToModify.push("G"+ (m + 3)+"");
          }
        });
      }
      // Se c'è almeno una riga nella variabile..
      if(rangeToModify.length !== 0){
        // ...imposta il valore "1" nei record presenti in rangeToModify.
        ss.getSheetByName(item).getRangeList(rangeToModify).setValue(1);
        // Svuota rangeToModify per essere utilizzato nello sheet del conto successivo.
        rangeToModify = [];
      }
    });
  // ..altrimenti se è impostato di aggiornare le entrate..
  }else if(type == "income"){
    
    var statsSheet = ss.getSheetByName("Monthly Income - Total");
    var settingsSheet = ss.getSheetByName("» Income: Cat./Subcat.");
    var incomeCategoryList = ss.getSheetByName("» Income: Cat./Subcat.");
    var lastRowDataIncome = incomeCategoryList.getLastRow();
    // Finestra che informa lo stato della funzione in corso
    SpreadsheetApp.getActive().toast("Mi preparo per l'aggiornamento degli sheet dei conti..", "Script in esecuzione", -1);
    // Per ogni sottocategoria...
    for(var i = 1; i <= lastRowDataIncome; i++){
      // ...se la sottocategoria ha nella terza colonna "true" va considerata tra quelle da escludere nel conteggio totale, quindi viene messa dentro la var subCategoryCleaned.
      if(incomeCategoryList.getRange(i, 3).getValue() === true){
        subCategoryCleaned[z] = { "category": incomeCategoryList.getRange(i, 1).getValue(), "subcategory": incomeCategoryList.getRange(i, 2).getValue() };
        z++;
      }
    }
    // Per ogni conto..
    banks.forEach(function(item){
      // Finestra che informa lo stato della funzione in corso
      SpreadsheetApp.getActive().toast("Procedo ad aggiornare lo sheet del conto "+ item +" con i valori delle uscite da non considerare nella conta totale..", "Script in esecuzione", -1);
      // Resetta i campi impostandoli tutti a zero
      ss.getSheetByName(item).getRange(4, 9, ss.getSheetByName(item).getRange("L4:L").getValues().filter(String).length).setValue(0);
      // Per ogni record presente nello sheet del conto...
      for(var m = 1; m <= (ss.getSheetByName(item).getRange("L4:L").getValues().filter(String).length); m++){
        // ... se la categoria è presente tra quelle dentro subcategoryCleaned aggiunge la riga dello sheet del conto in questione nella variabile rangeToModify.
        subCategoryCleaned.some(element => {
          if (element.category === ss.getSheetByName(item).getRange(m + 3, 13).getValue() && element.subcategory === ss.getSheetByName(item).getRange(m + 3, 14).getValue()) {
            rangeToModify.push("I"+ (m + 3)+"");
          }
        });
      }
      // Se c'è almeno una riga nella variabile..
      if(rangeToModify.length !== 0){
        // ...imposta il valore "1" nei record presenti in rangeToModify.
        ss.getSheetByName(item).getRangeList(rangeToModify).setValue(1);
        // Svuota rangeToModify per essere utilizzato nello sheet del conto successivo.
        rangeToModify = [];
      }
    });
  // ..altrimenti c'è un problema.
  }else{
    // Finestra che informa lo stato della funzione in corso
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
  // Finestra che informa lo stato della funzione in corso
  SpreadsheetApp.getActive().toast("Genero la tabella delle statistiche mensili totali..", "Script in esecuzione", -1);
  // Genero stats mensili totali
  var catStartTotal = categoryStatsTotal(type, settingsSheet, statsSheet, categorySpace, initialRowSpace, initialColumnSpace, firstLineCategory);
  subCategoryStatsTotal(type, settingsSheet, statsSheet, catStartTotal, lastLineCategory, initialColumnSpace);
  // Finestra che informa lo stato della funzione in corso
  SpreadsheetApp.getActive().toast("Finito!", "Script terminato", 4);
}
