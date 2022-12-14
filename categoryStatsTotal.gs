// Funzione che genera le tabelle mensili suddivise per categorie con i valori totali.
function categoryStatsTotal(type, settingsSheet, statsSheet, categorySpace, initialRowSpace, initialColumnSpace, firstLineCategory){
  
  var settingsValue = settingsSheet.getDataRange().getValues();
  // Indica la colonna con i valori delle categorie da usare per generare le tabelle per le uscite o le entrate.
  var categoryColumn = type == "expenses" ? 4 : type == "income" ? 13 : false;
  // Indica la colonna con i valori da usare per generare le tabelle per le uscite o le entrate.
  var moneyColumn = type == "expenses" ? 3 : type == "income" ? 12 : false;
  // Indica la colonna con i valori dei mesi da usare per generare le tabelle per le uscite o le entrate.
  var monthColumn = type == "expenses" ? 2 : type == "income" ? 11 : false;
  // Indica la colonna con i valori dei mesi da usare per generare le tabelle per le uscite o le entrate.
  var cleanColumns = type == "expenses" ? 7 : type == "income" ? 9 : false;
  var newInitialRowSpace = categorySpace;
  // Indica lo spazio orizzontale (numero di colonne) da tenere prima della lista delle categorie.
  initialColumnSpace = 3;

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
      statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 2, initialColumnSpace + 5 + i).setValue("Totale");
      statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 2, initialColumnSpace + 6 + i).setValue("Grafici");
      statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 2, initialColumnSpace + 6 + i, 1, 2).merge();
    }else{
      statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 2, initialColumnSpace + 1 + i).setValue(months[i]);
      statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 1, initialColumnSpace + 1 + i).setValue(i + 1);
    }
  }
  statsSheet.hideRows((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 1);
  var cleanedSettingValue = [];
  // Esclude dai record le sottocategorie da escludere perch?? non sono ne uscite ne entrate 
  for(currentlyCategoryLine = 0; currentlyCategoryLine < settingsValue.length; currentlyCategoryLine++){
    if(settingsValue[currentlyCategoryLine][2] === false){
      cleanedSettingValue.push(currentlyCategoryLine);
    }
  }
  firstLineCategory = cleanedSettingValue[0];
  // Ricavo la lista dei mesi presenti nello sheet dei conti al fine di individuare l'ultimo mese inserito
  var monthList = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(banks[0]).getRange(columnToLetter(monthColumn) +"4:"+ columnToLetter(monthColumn)).getValues();
  monthList = monthList.reduce(function(prev, next) {
    return prev.concat(next);
  });
  monthList = monthList.filter(n => n);
  // Ricavo l'ultimo mese inserito
  var lastMonth = Math.max.apply(null, monthList);
  // Variabile utilizzata per sapere in anticipo il numero di categorie valide (e non sottocategorie). Viene usata per i grafici.
  var categoryLengthCheck = [];
  // Inserisco nella variabile tutte le categorie valide (e non sottocagetorie).
  cleanedSettingValue.forEach(function(item){
    if(categoryLengthCheck.indexOf(settingsValue[item][0]) === -1){
      categoryLengthCheck.push(settingsValue[item][0]);
    }
  });
  cleanedSettingValue.forEach(function (item, index){
    // Se le due categorie sono differenti o se si tratta dell'ultima sottocategoria.
    if(index === (cleanedSettingValue.length - 1) || settingsValue[item][0] !== settingsValue[firstLineCategory][0]){
      // Se si tratta dell'ultima sottocategoria.
      if(index === (cleanedSettingValue.length - 1)){
        // Scrive il nome della categoria attuale.
        statsSheet.getRange(initialRowSpace + categorySpace,initialColumnSpace - 1).setValue(settingsValue[item][0]);
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace - 1, 1, 2).merge();
        // Per ogni mese aggiunge la formula "query" per calcolare i valori per singola categoria.
        for(e = 1; e <= 12; e++){
          var string = "";
          // Per ogni banca aggiunge un pezzo della stringa cos?? da fare il totale.
          banks.forEach(function(item, index, array){
            // Se ?? la prima..
            if (index === 0){ 
              string = string + "= \
                IFERROR(QUERY('"+ item +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                  "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + categorySpace) +"&\"' \n" +
                  "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                  "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
                "label sum("+ columnToLetter(moneyColumn) +") ''\");0) \n" +
                "+ \n";
            // ..se ?? l'ultima.. 
            }else if(index === array.length - 1){
              string = string + "IFERROR(QUERY('"+ item +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                  "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + categorySpace) +"&\"' \n" +
                  "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                  "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
                "label sum("+ columnToLetter(moneyColumn) +") ''\");0)";
            // ..tutte le altre.
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
        // Inserisce la colonna con le medie.
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace + 14 /* 12 mesi + 1 colonna di spazio + la colonna in cui deve scrivere */).setValue("= \n" +
          "AVERAGE("+ columnToLetter(initialColumnSpace + 1) +""+ (initialRowSpace + categorySpace) +":"+ columnToLetter(initialColumnSpace + lastMonth)+""+ (initialRowSpace + categorySpace) +")");
        // Inserisce la colonna con il rapporto ultimo mese/media.
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace + 15 /* 12 mesi + 2 colonna di spazio + la colonna in cui deve scrivere */).setValue("= \n" +
          "IFERROR(("+ columnToLetter(initialColumnSpace + lastMonth) +""+ (initialRowSpace + categorySpace) +"/Q"+ (initialRowSpace + categorySpace) +")-1;\"-\")");
        // Inseriesce la colonna con il totale per categoria.
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace + 16 /* 12 mesi + 3 colonne di spazio + la colonna in cui deve scrivere */).setValue("= \n" +
          "SUM("+ columnToLetter(initialColumnSpace + 1) +""+ (initialRowSpace + categorySpace) +":"+ columnToLetter(initialColumnSpace + lastMonth)+""+ (initialRowSpace + categorySpace) +")");
        // Inserisce la colonna con il grafico a barre.
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace + 17 /* 12 mesi + 4 colonne di spazio + la colonna in cui deve scrivere */).setValue('= \n' +
          'SPARKLINE(S'+ (initialRowSpace + categorySpace) +';{"charttype"\\"bar";"color1"\\"teal"; "max"\\ ROUNDUP(max(S'+ (initialRowSpace) +':S'+ (initialRowSpace + categoryLengthCheck.length - 1) +'))})');
        // Inserisce la colonna con il grafico a linee.
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace + 18 /* 12 mesi + 5 colonne di spazio + la colonna in cui deve scrivere */).setValue('= \n' +
          'SPARKLINE('+ columnToLetter(initialColumnSpace + 1) +''+ (initialRowSpace + categorySpace) +':'+ columnToLetter(initialColumnSpace + lastMonth)+''+ (initialRowSpace + categorySpace) +';' +
          '{"color"\\IF('+ columnToLetter(initialColumnSpace + lastMonth) +''+ (initialRowSpace + categorySpace) +'>Q'+ (initialRowSpace + categorySpace) +';"red";"green");' +
          '"ymax"\\ROUNDUP(max('+ columnToLetter(initialColumnSpace + 1) +''+ (initialRowSpace + categorySpace) +':'+ columnToLetter(initialColumnSpace + lastMonth)+''+ (initialRowSpace + categorySpace) +')); "linewidth"\\2})');
        categorySpace++;
      // Se non si tratta dell'ultima sottocategoria.
      }else{
        // Scrive il nome della categoria precedente.
        statsSheet.getRange(initialRowSpace + categorySpace,initialColumnSpace).setValue(settingsValue[firstLineCategory][0]);
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace - 1, 1, 2).merge();
        // Per ogni mese aggiunge la formula "query" per calcolare i valori per singola categoria.
        for(e = 1; e <= 12; e++){
          var string = "";
          // Per ogni banca aggiunge un pezzo della stringa cos?? da fare il totale.
          banks.forEach(function(item, index, array){
            // Se ?? la prima..
            if (index === 0){ 
              string = string + "= \
                IFERROR(QUERY('"+ item +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                  "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + categorySpace) +"&\"' \n" +
                  "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                  "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
                "label sum("+ columnToLetter(moneyColumn) +") ''\");0) \n" +
                "+ \n";
            // ..se ?? l'ultima.. 
            }else if(index === array.length - 1){
              string = string + "IFERROR(QUERY('"+ item +"'!$A$4:$O$1000;\"select sum("+ columnToLetter(moneyColumn) +") \n" +
                  "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + categorySpace) +"&\"' \n" +
                  "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
                  "and "+ columnToLetter(cleanColumns) +" <> 1 \n" +
                "label sum("+ columnToLetter(moneyColumn) +") ''\");0)";
            // ..tutte le altre.
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
        // Inserisce la colonna con le medie.
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace + 14 /* 12 mesi + 1 colonna di spazio + la colonna in cui deve scrivere */).setValue("= \n" +
          "AVERAGE("+ columnToLetter(initialColumnSpace + 1) +""+ (initialRowSpace + categorySpace) +":"+ columnToLetter(initialColumnSpace + lastMonth)+""+ (initialRowSpace + categorySpace) +")");
        // Inserisce la colonna con il rapporto ultimo mese/media.
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace + 15 /* 12 mesi + 2 colonne di spazio + la colonna in cui deve scrivere */).setValue("= \n" +
          "IFERROR(("+ columnToLetter(initialColumnSpace + lastMonth) +""+ (initialRowSpace + categorySpace) +"/Q"+ (initialRowSpace + categorySpace) +")-1;\"-\")");
        // Inseriesce la colonna con il totale per categoria.
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace + 16 /* 12 mesi + 3 colonne di spazio + la colonna in cui deve scrivere */).setValue("= \n" +
          "SUM("+ columnToLetter(initialColumnSpace + 1) +""+ (initialRowSpace + categorySpace) +":"+ columnToLetter(initialColumnSpace + lastMonth)+""+ (initialRowSpace + categorySpace) +")");
        // Inserisce la colonna con il grafico a barre.
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace + 17 /* 12 mesi + 4 colonne di spazio + la colonna in cui deve scrivere */).setValue('= \n' +
          'SPARKLINE(S'+ (initialRowSpace + categorySpace) +';{"charttype"\\"bar";"color1"\\"teal"; "max"\\ ROUNDUP(max(S'+ (initialRowSpace) +':S'+ (initialRowSpace + categoryLengthCheck.length - 1) +'))})');
        // Inserisce la colonna con il grafico a linee.
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace + 18 /* 12 mesi + 5 colonne di spazio + la colonna in cui deve scrivere */).setValue('= \n' +
          'SPARKLINE('+ columnToLetter(initialColumnSpace + 1) +''+ (initialRowSpace + categorySpace) +':'+ columnToLetter(initialColumnSpace + lastMonth)+''+ (initialRowSpace + categorySpace) +';' +
          '{"color"\\IF('+ columnToLetter(initialColumnSpace + lastMonth) +''+ (initialRowSpace + categorySpace) +'>Q'+ (initialRowSpace + categorySpace) +';"red";"green");' +
          '"ymax"\\ROUNDUP(max('+ columnToLetter(initialColumnSpace + 1) +''+ (initialRowSpace + categorySpace) +':'+ columnToLetter(initialColumnSpace + lastMonth)+''+ (initialRowSpace + categorySpace) +')); "linewidth"\\2})');
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
