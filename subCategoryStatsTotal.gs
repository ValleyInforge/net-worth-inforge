function subCategoryStatsTotal(type, settingsSheet, statsSheet, initialRowSpace, lastLineCategory, initialColumnSpace){

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
  initialColumnSpace = 2;
  initialRowSpace += 5;

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
