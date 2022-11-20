// Devo fare un array con settingsValue.length dal quale togliere il numero delle righe che non rispetta settingsValue.length[x][3] === false e poi con le restanti faccio il resto
function updateMonthlyIncomeExpensesTotal(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var banks = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu").getRange(3,5, 5).getValues();
  // CONVERTO L'ARRAY 2D A 1D
  banks = banks.reduce(function(prev, next) {
    return prev.concat(next);
  });
  banks = banks.filter(n => n);
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
  var hypeCatStartTotal = categoryStatsTotal(type, settingsSheet, statsSheet, categorySpace);
  subCategoryStatsTotal(type, settingsSheet, statsSheet, hypeCatStartTotal);
  SpreadsheetApp.getActive().toast("Finito!", "Script terminato", 4);
}

/*
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
*/

function categoryStatsTotal(type, settingsSheet, statsSheet, categorySpace){
  
  var banks = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu").getRange(3,5, 5).getValues();
  // CONVERTO L'ARRAY 2D A 1D
  banks = banks.reduce(function(prev, next) {
    return prev.concat(next);
  });
  banks = banks.filter(n => n);
  var settingsValue = settingsSheet.getDataRange().getValues();
  // Indica la colonna con i valori delle categorie da usare per generare le tabelle per le uscite o le entrate.
  var categoryColumn = type == "expenses" ? 4 : type == "income" ? 13 : false;
  // Indica la colonna con i valori da usare per generare le tabelle per le uscite o le entrate.
  var moneyColumn = type == "expenses" ? 3 : type == "income" ? 12 : false;
  // Indica la colonna con i valori dei mesi da usare per generare le tabelle per le uscite o le entrate.
  var monthColumn = type == "expenses" ? 2 : type == "income" ? 11 : false;
  // Indica la colonna con i valori dei mesi da usare per generare le tabelle per le uscite o le entrate.
  var cleanColumns = type == "expenses" ? 7 : type == "income" ? 9 : false;
  // Indica lo spazio verticale (numero di righe) da tenere prima della lista delle categorie.
  var initialRowSpace = 5;
  var newInitialRowSpace = categorySpace;
  // Indica lo spazio orizzontale (numero di colonne) da tenere prima della lista delle categorie.
  var initialColumnSpace = 3;
  var months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

  // Crea l'intestazione per le categorie e le sotto-categorie.  
  statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 2, initialColumnSpace - 1).setValue("Categories");
  statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) -2, initialColumnSpace - 1, 1, 2).mergeAcross();
  //Crea l'intestazione dei mesi.
  for(i = 0; i < 12; i++){
    if(i === 11){
      statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 2, initialColumnSpace + 1 + i).setValue(months[i]);
      statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 1, initialColumnSpace + 1 + i).setValue(i + 1);
      statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 2, initialColumnSpace + 3 + i).setValue("Media");
      statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 2, initialColumnSpace + 4 + i).setValue("Ultimo mese risp media");
    }else{
      statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 2, initialColumnSpace + 1 + i).setValue(months[i]);
      statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 1, initialColumnSpace + 1 + i).setValue(i + 1);
    }
  }
  statsSheet.hideRows((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 1);
  var cleanedSettingValue = [];
  // Esclude dai record le sottocategorie da escludere perché non sono ne uscite ne entrate 
  for(currentlyCategoryLine = 0; currentlyCategoryLine < settingsValue.length; currentlyCategoryLine++){
    if(settingsValue[currentlyCategoryLine][2] === false){
      cleanedSettingValue.push(currentlyCategoryLine);
    }
  }
  var firstLineCategory = cleanedSettingValue[0];
  // Ricavo la lista dei mesi presenti nello sheet dei conti al fine di individuare l'ultimo mese inserito
  var monthList = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(banks[0]).getRange(columnToLetter(monthColumn) +"4:"+ columnToLetter(monthColumn)).getValues();
  monthList = monthList.reduce(function(prev, next) {
    return prev.concat(next);
  });
  monthList = monthList.filter(n => n);
  // Ricavo l'ultimo mese inserito
  var lastMonth = Math.max.apply(null, monthList);
  cleanedSettingValue.forEach(function (item, index){
    // Se le due categorie sono differenti o se si tratta dell'ultima sottocategoria.
    if(index === (cleanedSettingValue.length - 1) || settingsValue[item][0] !== settingsValue[firstLineCategory][0]){
      // Se si tratta dell'ultima sottocategoria.
      if(index === (cleanedSettingValue.length - 1)){
        // Scrive il nome della categoria attuale.
        statsSheet.getRange(initialRowSpace + categorySpace,initialColumnSpace - 1).setValue(settingsValue[item][0]);
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace - 1, 1, 2).merge();
        for(e = 1; e <= 12; e++){
          var string = "";
          banks.forEach(function(item, index, array){
            if (index === 0){ 
              string = string + "= \
                IFERROR(QUERY('"+ item +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                  "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + categorySpace) +"&\"' \n" +
                  "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                  "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
                "label sum("+ columnToLetter(moneyColumn) +") ''\");0) \n" +
                "+ \n";
            }else if(index === array.length - 1){
              string = string + "IFERROR(QUERY('"+ item +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                  "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + categorySpace) +"&\"' \n" +
                  "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                  "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
                "label sum("+ columnToLetter(moneyColumn) +") ''\");0)";
            }else{
              string = string + "IFERROR(QUERY('"+ item +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                  "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + categorySpace) +"&\"' \n" +
                  "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                  "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
                "label sum("+ columnToLetter(moneyColumn) +") ''\");0) \n" +
                "+ \n";
            }
          });
          statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace + e).setValue(string);
        }
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace + 14 /* 12 mesi + 1 colonna di spazio + la colonna in cui deve scrivere */).setValue("= \n" +
          "AVERAGE("+ columnToLetter(initialColumnSpace + 1) +""+ (initialRowSpace + categorySpace) +":"+ columnToLetter(initialColumnSpace + lastMonth)+""+ (initialRowSpace + categorySpace) +")");
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace + 15 /* 12 mesi + 2 colonna di spazio + la colonna in cui deve scrivere */).setValue("= \n" +
          "IFERROR(("+ columnToLetter(initialColumnSpace + lastMonth) +""+ (initialRowSpace + categorySpace) +"/Q"+ (initialRowSpace + categorySpace) +")-1;\"-\")");
        categorySpace++;
      // Se non si tratta dell'ultima sottocategoria.
      }else{
        // Scrive il nome della categoria precedente.
        statsSheet.getRange(initialRowSpace + categorySpace,initialColumnSpace).setValue(settingsValue[firstLineCategory][0]);
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace - 1, 1, 2).merge();
        for(e = 1; e <= 12; e++){
          var string = "";
          banks.forEach(function(item, index, array){
            if (index === 0){ 
              string = string + "= \
                IFERROR(QUERY('"+ item +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                  "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + categorySpace) +"&\"' \n" +
                  "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                  "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
                "label sum("+ columnToLetter(moneyColumn) +") ''\");0) \n" +
                "+ \n";
            }else if(index === array.length - 1){
              string = string + "IFERROR(QUERY('"+ item +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                  "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + categorySpace) +"&\"' \n" +
                  "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                  "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
                "label sum("+ columnToLetter(moneyColumn) +") ''\");0)";
            }else{
              string = string + "IFERROR(QUERY('"+ item +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                  "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + categorySpace) +"&\"' \n" +
                  "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                  "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
                "label sum("+ columnToLetter(moneyColumn) +") ''\");0) \n" +
                "+ \n";
            }
          });
          statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace + e).setValue(string);
        }
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace + 14 /* 12 mesi + 1 colonna di spazio + la colonna in cui deve scrivere */).setValue("= \n" +
          "AVERAGE("+ columnToLetter(initialColumnSpace + 1) +""+ (initialRowSpace + categorySpace) +":"+ columnToLetter(initialColumnSpace + lastMonth)+""+ (initialRowSpace + categorySpace) +")");
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace + 15 /* 12 mesi + 2 colonna di spazio + la colonna in cui deve scrivere */).setValue("= \n" +
          "IFERROR(("+ columnToLetter(initialColumnSpace + lastMonth) +""+ (initialRowSpace + categorySpace) +"/Q"+ (initialRowSpace + categorySpace) +")-1;\"-\")");
        categorySpace++;
      }
      // Aggiorna il valore di firstLineCategory al fine di individuare la giusta posizione per la categoria successiva.
      firstLineCategory = item;
    }
  });
  statsSheet
    .getRange(initialRowSpace + categorySpace, initialColumnSpace)
    .setValue("=SUM("+ columnToLetter(initialColumnSpace + 1) +""+ (initialRowSpace + categorySpace) +":"+ columnToLetter(initialColumnSpace + 12) +""+ (initialRowSpace + categorySpace) +")");
  for(e = 1; e <= 12; e++){
    statsSheet
      .getRange(initialRowSpace + categorySpace, initialColumnSpace + e)
      .setValue("=SUM("+ columnToLetter(initialColumnSpace + e) +""+((newInitialRowSpace === 0 ? initialRowSpace : newInitialRowSpace + initialRowSpace))+":"+ columnToLetter(initialColumnSpace + e) +""+(initialRowSpace + categorySpace-1)+")");
  }
  return categorySpace + initialRowSpace;
}

/*
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
*/

function subCategoryStatsTotal(type, settingsSheet, statsSheet, initialRowSpace){

  var banks = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu").getRange(3,5, 5).getValues();
  // CONVERTO L'ARRAY 2D A 1D
  banks = banks.reduce(function(prev, next) {
    return prev.concat(next);
  });
  banks = banks.filter(n => n);
  // Indica tutti i valori presenti nello sheet "Impostazioni".
  var settingsValue = settingsSheet.getDataRange().getValues();
  // Indica la colonna con i valori delle categorie da usare per generare le tabelle per le uscite o le entrate.
  var subCategoryColumn = type == "expenses" ? 5 : type == "income" ? 14 : false;
  // Indica la colonna con i valori delle categorie da usare per generare le tabelle per le uscite o le entrate.
  var categoryColumn = type == "expenses" ? 4 : type == "income" ? 13 : false;
  // Indica la colonna con i valori da usare per generare le tabelle per le uscite o le entrate.
  var moneyColumn = type == "expenses" ? 3 : type == "income" ? 12: false;
  // Indica la colonna con i valori dei mesi da usare per generare le tabelle per le uscite o le entrate.
  var monthColumn = type == "expenses" ? 2 : type == "income" ? 11 : false;
  // Indica la colonna con i valori dei mesi da usare per generare le tabelle per le uscite o le entrate.
  var cleanColumns = type == "expenses" ? 7 : type == "income" ? 9 : false;
  // Indica lo spazio orizzontale (numero di colonne) da tenere prima della lista delle categorie.
  var initialColumnSpace = 2;
  // Indica il numero di righe da unire perché appartenenti alla stessa categoria. Impostato di partenza a 0, al primo caso sale sempre a 1. Poi al reset tornerà sempre a 1.
  var lastLineCategory = 0;
  initialRowSpace += 5;
  var months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

  // Crea l'intestazione per le categorie e le sotto-categorie.
  statsSheet.getRange(initialRowSpace - 2, initialColumnSpace).setValue("Categories");
  statsSheet.getRange(initialRowSpace - 2, initialColumnSpace + 1).setValue("Subcategories");
  //Crea l'intestazione dei mesi.
  for(i = 0; i < 12; i++){
    if(i === 11){
      statsSheet.getRange(initialRowSpace - 2, initialColumnSpace + 2 + i).setValue(months[i]);
      statsSheet.getRange(initialRowSpace - 1, initialColumnSpace + 2 + i).setValue(i + 1);
      statsSheet.getRange(initialRowSpace - 2, initialColumnSpace + 4 + i).setValue("Media");
      statsSheet.getRange(initialRowSpace - 2, initialColumnSpace + 5 + i).setValue("Ultimo mese risp media");
    }else{
      statsSheet.getRange(initialRowSpace - 2, initialColumnSpace + 2 + i).setValue(months[i]);
      statsSheet.getRange(initialRowSpace - 1, initialColumnSpace + 2 + i).setValue(i + 1);
    }
  }
  statsSheet.hideRows(initialRowSpace - 1);
  var cleanedSettingValue = [];
  for(currentlyCategoryLine = 0; currentlyCategoryLine < settingsValue.length; currentlyCategoryLine++){
    if(settingsValue[currentlyCategoryLine][2] === false){
      cleanedSettingValue.push(currentlyCategoryLine);
    }
  }
  // Ricavo la lista dei mesi presenti nello sheet dei conti al fine di individuare l'ultimo mese inserito
  var monthList = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(banks[0]).getRange(columnToLetter(monthColumn) +"4:"+ columnToLetter(monthColumn)).getValues();
  monthList = monthList.reduce(function(prev, next) {
    return prev.concat(next);
  });
  monthList = monthList.filter(n => n);
  // Ricavo l'ultimo mese inserito
  var lastMonth = Math.max.apply(null, monthList);
  var previousLineCategory = cleanedSettingValue[0];
  var beginningFirstLineCategory = 0;
  var previousCategoryName;
  cleanedSettingValue.forEach(function (item, index){
    // SE SI TRATTA DELL'ULTIMA CATEGORIA
    if(index === (cleanedSettingValue.length - 1)){
      // SE LE CATEGORIE SONO DIVERSE
      if(settingsValue[item][0] !== settingsValue[previousLineCategory][0]){
        // SCRIVE IL NOME DELLA CATEGORIA PRECEDENTE
        statsSheet.getRange(initialRowSpace + beginningFirstLineCategory, initialColumnSpace).setValue(settingsValue[previousCategoryName][0]);
        // UNISCE LE CELLE APPARTENENTI ALLA STESSA CATEGORIA PRECEDENTE
        statsSheet.getRange(initialRowSpace + beginningFirstLineCategory, initialColumnSpace, lastLineCategory).mergeVertically();
        // SCRIVE IL NOME DELLA SOTTOCATEGORIA ATTUALE
        statsSheet.getRange(initialRowSpace + index, initialColumnSpace + 1).setValue(settingsValue[item][1]);
        // AGGIORNA IL VALORE DI previousLineCategory AL FINE DI INDIVIDUARE LA GIUSTA POSIZIONE PER LA CATEGORIA SUCCESSIVA E PER LE SOMME FINALI
        previousLineCategory = item;
        beginningFirstLineCategory = index;
        // AGGIUNGE I DATI MENSILI DELLA SOTTOCATEGORIA ATTUALE
        for(e = 1; e <= 12; e++){
          var string = "";
          banks.forEach(function(itemm, indexx, arrayy){
            if (indexx === 0){ 
              string = string + "= \
                IFERROR(QUERY('"+ itemm +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                  "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + beginningFirstLineCategory) +"&\"' \n" +
                  "and "+ columnToLetter(subCategoryColumn) +" = '\"&C"+ (initialRowSpace + index) +"&\"' \n" +
                  "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                  "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
                "label sum("+ columnToLetter(moneyColumn) +") ''\");0) \n" +
                "+ \n";
            }else if(indexx === arrayy.length - 1){
              string = string + "IFERROR(QUERY('"+ itemm +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                  "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + beginningFirstLineCategory) +"&\"' \n" +
                  "and "+ columnToLetter(subCategoryColumn) +" = '\"&C"+ (initialRowSpace + index) +"&\"' \n" +
                  "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                  "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
                "label sum("+ columnToLetter(moneyColumn) +") ''\");0)";
            }else{
              string = string + "IFERROR(QUERY('"+ itemm +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                  "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + beginningFirstLineCategory) +"&\"' \n" +
                  "and "+ columnToLetter(subCategoryColumn) +" = '\"&C"+ (initialRowSpace + index) +"&\"' \n" +
                  "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                  "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
                "label sum("+ columnToLetter(moneyColumn) +") ''\");0) \n" +
                "+ \n";
            }
          });
          statsSheet.getRange(initialRowSpace + index, initialColumnSpace + 1 + e).setValue(string);
        }
        statsSheet.getRange(initialRowSpace + index, initialColumnSpace + 15 /* 12 mesi + 2 colonna di spazio perché initialRow è 2 e non 3 come prima + la colonna in cui deve scrivere */).setValue("= \n" +
          "AVERAGE("+ columnToLetter(initialColumnSpace + 2) +""+ (initialRowSpace + index) +":"+ columnToLetter(initialColumnSpace + 1 + lastMonth)+""+ (initialRowSpace + index) +")");
        statsSheet.getRange(initialRowSpace + index, initialColumnSpace + 16 /* 12 mesi + 2 colonna di spazio + la colonna in cui deve scrivere */).setValue("= \n" +
          "IFERROR(("+ columnToLetter(initialColumnSpace + 1 + lastMonth) +""+ (initialRowSpace + index) +"/Q"+ (initialRowSpace + index) +")-1;\"-\")");
        // SCRIVE IL NOME DELLA CATEGORIA ATTUALE
        statsSheet.getRange(initialRowSpace + index, initialColumnSpace).setValue(settingsValue[item][0]);
      // SE LE CATEGORIE SONO UGUALI
      }else if (settingsValue[item][0] === settingsValue[previousLineCategory][0]){
        // SCRIVE IL NOME DELLA SOTTOCATEGORIA ATTUALE
        statsSheet.getRange(initialRowSpace + index, initialColumnSpace + 1).setValue(settingsValue[item][1]);
        // AGGIUNGE I DATI MENSILI DELLA SOTTOCATEGORIA ATTUALE
        for(e = 1; e <= 12; e++){
          var string = "";
          banks.forEach(function(itemm, indexx, arrayy){
            if (indexx === 0){ 
              string = string + "= \
                IFERROR(QUERY('"+ itemm +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                  "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + beginningFirstLineCategory) +"&\"' \n" +
                  "and "+ columnToLetter(subCategoryColumn) +" = '\"&C"+ (initialRowSpace + index) +"&\"' \n" +
                  "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                  "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
                "label sum("+ columnToLetter(moneyColumn) +") ''\");0) \n" +
                "+ \n";
            }else if(indexx === arrayy.length - 1){
              string = string + "IFERROR(QUERY('"+ itemm +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                  "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + beginningFirstLineCategory) +"&\"' \n" +
                  "and "+ columnToLetter(subCategoryColumn) +" = '\"&C"+ (initialRowSpace + index) +"&\"' \n" +
                  "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                  "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
                "label sum("+ columnToLetter(moneyColumn) +") ''\");0)";
            }else{
              string = string + "IFERROR(QUERY('"+ itemm +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                  "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + beginningFirstLineCategory) +"&\"' \n" +
                  "and "+ columnToLetter(subCategoryColumn) +" = '\"&C"+ (initialRowSpace + index) +"&\"' \n" +
                  "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                  "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
                "label sum("+ columnToLetter(moneyColumn) +") ''\");0) \n" +
                "+ \n";
            }
          });
          statsSheet.getRange(initialRowSpace + index, initialColumnSpace + 1 + e).setValue(string);
        }
        statsSheet.getRange(initialRowSpace + index, initialColumnSpace + 15 /* 12 mesi + 2 colonna di spazio perché initialRow è 2 e non 3 come prima + la colonna in cui deve scrivere */).setValue("= \n" +
          "AVERAGE("+ columnToLetter(initialColumnSpace + 2) +""+ (initialRowSpace + index) +":"+ columnToLetter(initialColumnSpace + 1 + lastMonth)+""+ (initialRowSpace + index) +")");
        statsSheet.getRange(initialRowSpace + index, initialColumnSpace + 16 /* 12 mesi + 2 colonna di spazio + la colonna in cui deve scrivere */).setValue("= \n" +
          "IFERROR(("+ columnToLetter(initialColumnSpace + 1 + lastMonth) +""+ (initialRowSpace + index) +"/Q"+ (initialRowSpace + index) +")-1;\"-\")");
        // INCREMENTA DI 1 IL NUMERO DI RIGHE DA UNIRE QUANDO LE CATEGORIE SONO LA STESSA
        lastLineCategory++;
        // SCRIVE IL NOME DELLA CATEGORIA ATTUALE
        statsSheet.getRange(initialRowSpace + beginningFirstLineCategory, initialColumnSpace).setValue(settingsValue[item][0]);
        // UNISCE LE CELLE APPARTENENTI ALLA STESSA CATEGORIA
        statsSheet.getRange(initialRowSpace + beginningFirstLineCategory, initialColumnSpace, lastLineCategory).mergeVertically();
        // AGGIORNA IL VALORE DI previousLineCategory AL FINE DI INDIVIDUARE LA GIUSTA POSIZIONE PER LE SOMME FINALI
        previousLineCategory = item;
        beginningFirstLineCategory = index;
      }
    // ALTRIMENTI SE NON È L'ULTIMA CATEGORIA MA SONO DIVERSE
    }else if(settingsValue[item][0] !== settingsValue[previousLineCategory][0]){
      // SCRIVE IL NOME DELLA CATEGORIA PRECEDENTE
      statsSheet.getRange(initialRowSpace + beginningFirstLineCategory, initialColumnSpace).setValue(settingsValue[previousCategoryName][0]);
      // UNISCE LE CELLE APPARTENENTI ALLA STESSA CATEGORIA PRECEDENTE
      statsSheet.getRange(initialRowSpace + beginningFirstLineCategory, initialColumnSpace, lastLineCategory).mergeVertically();
      // SCRIVE IL NOME DELLA SOTTOCATEGORIA ATTUALE
      statsSheet.getRange(initialRowSpace + index, initialColumnSpace + 1).setValue(settingsValue[item][1]);
      // AGGIORNA IL VALORE DI previousLineCategory AL FINE DI INDIVIDUARE LA GIUSTA POSIZIONE PER LA CATEGORIA SUCCESSIVA
      previousLineCategory = item;
      beginningFirstLineCategory = index;
      // AGGIUNGE I DATI MENSILI DELLA SOTTOCATEGORIA ATTUALE
      for(e = 1; e <= 12; e++){
        var string = "";
        banks.forEach(function(itemm, indexx, arrayy){
          if (indexx === 0){ 
            string = string + "= \
              IFERROR(QUERY('"+ itemm +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + beginningFirstLineCategory) +"&\"' \n" +
                "and "+ columnToLetter(subCategoryColumn) +" = '\"&C"+ (initialRowSpace + index) +"&\"' \n" +
                "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
              "label sum("+ columnToLetter(moneyColumn) +") ''\");0) \n" +
              "+ \n";
          }else if(indexx === arrayy.length - 1){
            string = string + "IFERROR(QUERY('"+ itemm +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + beginningFirstLineCategory) +"&\"' \n" +
                "and "+ columnToLetter(subCategoryColumn) +" = '\"&C"+ (initialRowSpace + index) +"&\"' \n" +
                "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
              "label sum("+ columnToLetter(moneyColumn) +") ''\");0)";
          }else{
            string = string + "IFERROR(QUERY('"+ itemm +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + beginningFirstLineCategory) +"&\"' \n" +
                "and "+ columnToLetter(subCategoryColumn) +" = '\"&C"+ (initialRowSpace + index) +"&\"' \n" +
                "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
              "label sum("+ columnToLetter(moneyColumn) +") ''\");0) \n" +
              "+ \n";
          }
        });
        statsSheet.getRange(initialRowSpace + index, initialColumnSpace + 1 + e).setValue(string);
      }
      statsSheet.getRange(initialRowSpace + index, initialColumnSpace + 15 /* 12 mesi + 2 colonna di spazio perché initialRow è 2 e non 3 come prima + la colonna in cui deve scrivere */).setValue("= \n" +
        "AVERAGE("+ columnToLetter(initialColumnSpace + 2) +""+ (initialRowSpace + index) +":"+ columnToLetter(initialColumnSpace + 1 + lastMonth)+""+ (initialRowSpace + index) +")");
      statsSheet.getRange(initialRowSpace + index, initialColumnSpace + 16 /* 12 mesi + 2 colonna di spazio + la colonna in cui deve scrivere */).setValue("= \n" +
        "IFERROR(("+ columnToLetter(initialColumnSpace + 1 + lastMonth) +""+ (initialRowSpace + index) +"/Q"+ (initialRowSpace + index) +")-1;\"-\")");
      // RESETTA IL VALORE DI lastLineCategory
      lastLineCategory = 1;
      previousCategoryName = item;
    // ALTRIMENTI SE NON È L'ULTIMA CATEGORIA MA SONO UGUALI
    }else if (settingsValue[item][0] === settingsValue[previousLineCategory][0]){
      // SCRIVE IL NOME DELLA SOTTOCATEGORIA ATTUALE
      statsSheet.getRange(initialRowSpace + index, initialColumnSpace + 1).setValue(settingsValue[item][1]);
      // AGGIUNGE I DATI MENSILI DELLA SOTTOCATEGORIA ATTUALE
      for(e = 1; e <= 12; e++){
        var string = "";
        banks.forEach(function(itemm, indexx, arrayy){
          if (indexx === 0){ 
            string = string + "= \
              IFERROR(QUERY('"+ itemm +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + beginningFirstLineCategory) +"&\"' \n" +
                "and "+ columnToLetter(subCategoryColumn) +" = '\"&C"+ (initialRowSpace + index) +"&\"' \n" +
                "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
              "label sum("+ columnToLetter(moneyColumn) +") ''\");0) \n" +
              "+ \n";
          }else if(indexx === arrayy.length - 1){
            string = string + "IFERROR(QUERY('"+ itemm +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + beginningFirstLineCategory) +"&\"' \n" +
                "and "+ columnToLetter(subCategoryColumn) +" = '\"&C"+ (initialRowSpace + index) +"&\"' \n" +
                "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
              "label sum("+ columnToLetter(moneyColumn) +") ''\");0)";
          }else{
            string = string + "IFERROR(QUERY('"+ itemm +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + beginningFirstLineCategory) +"&\"' \n" +
                "and "+ columnToLetter(subCategoryColumn) +" = '\"&C"+ (initialRowSpace + index) +"&\"' \n" +
                "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
              "label sum("+ columnToLetter(moneyColumn) +") ''\");0) \n" +
              "+ \n";
          }
        });
        statsSheet.getRange(initialRowSpace + index, initialColumnSpace + 1 + e).setValue(string);
      }
      statsSheet.getRange(initialRowSpace + index, initialColumnSpace + 15 /* 12 mesi + 2 colonna di spazio perché initialRow è 2 e non 3 come prima + la colonna in cui deve scrivere */).setValue("= \n" +
        "AVERAGE("+ columnToLetter(initialColumnSpace + 2) +""+ (initialRowSpace + index) +":"+ columnToLetter(initialColumnSpace + 1 + lastMonth)+""+ (initialRowSpace + index) +")");
      statsSheet.getRange(initialRowSpace + index, initialColumnSpace + 16 /* 12 mesi + 2 colonna di spazio + la colonna in cui deve scrivere */).setValue("= \n" +
        "IFERROR(("+ columnToLetter(initialColumnSpace + 1 + lastMonth) +""+ (initialRowSpace + index) +"/Q"+ (initialRowSpace + index) +")-1;\"-\")");
      // INCREMENTA DI 1 IL NUMERO DI RIGHE DA UNIRE QUANDO LE CATEGORIE SONO LA STESSA
      lastLineCategory++;
      previousCategoryName = item;
    }
  });
  statsSheet
    .getRange(initialRowSpace + cleanedSettingValue.length, initialColumnSpace)
    .setValue("=SUM("+ columnToLetter(initialColumnSpace + 2) +""+ (initialRowSpace + cleanedSettingValue.length) +":"+ columnToLetter(initialColumnSpace + 13) +""+ (initialRowSpace + cleanedSettingValue.length) +")");
  statsSheet
    .getRange(initialRowSpace + previousLineCategory + 1, initialColumnSpace, 1, 2).merge();
  for(e = 1; e <= 12; e++){
    statsSheet
      .getRange(initialRowSpace + cleanedSettingValue.length, initialColumnSpace + e + 1)
      .setValue("=SUM("+ columnToLetter(initialColumnSpace + e + 1) +""+(initialRowSpace)+":"+ columnToLetter(initialColumnSpace + e + 1) +""+(initialRowSpace + cleanedSettingValue.length-1)+")");
  }
}