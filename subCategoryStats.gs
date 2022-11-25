// Valley - inforge.net //
// Funzione che genera le tabelle mensili per categorie e sottocategorie per ogni singolo conto
function subCategoryStats(type, settingsSheet, statsSheet, initialRowSpace, card, lastLineCategory, initialColumnSpace){
  // Indica tutti i valori presenti nello sheet "Impostazioni".
  var settingsValue = settingsSheet.getDataRange().getValues();
  // Indica la colonna con i valori delle categorie da usare per generare le tabelle per le uscite o le entrate.
  var subCategoryColumn = type == "expenses" ? 5 : type == "income" ? 14 : false;
  // Indica la colonna con i valori delle categorie da usare per generare le tabelle per le uscite o le entrate.
  var categoryColumn = type == "expenses" ? 4 : type == "income" ? 13 : false;
  // Indica la colonna con i valori economici da usare per generare le tabelle per le uscite o le entrate.
  var moneyColumn = type == "expenses" ? 3 : type == "income" ? 12: false;
  // Indica la colonna con i valori dei mesi da usare per generare le tabelle per le uscite o le entrate.
  var monthColumn = type == "expenses" ? 2 : type == "income" ? 11 : false;
  // Indica lo spazio orizzontale (numero di colonne) da tenere prima della lista delle categorie.
  initialColumnSpace = 2;
  // Indica quando sono finite le categorie e sotto-categorie.
  var lastSubCategory = 0;
  // Serve ad individuare la giusta posizione per la categoria successiva quando si genera la tabella. Impostato di partenza a 0.
  var firstLineCategory = 0;
  initialRowSpace += 5;

  // Crea l'intestazione per le categorie e le sotto-categorie.
  statsSheet.getRange(initialRowSpace - 2, initialColumnSpace).setValue("Categories");
  statsSheet.getRange(initialRowSpace - 2, initialColumnSpace + 1).setValue("Subcategories");
  //Crea l'intestazione dei mesi.
  for(i = 0; i < 12; i++){
    statsSheet.getRange(initialRowSpace - 2, initialColumnSpace + 2 + i).setValue(months[i]);
    statsSheet.getRange(initialRowSpace - 1, initialColumnSpace + 2 + i).setValue(i + 1);
  }
  statsSheet.hideRows(initialRowSpace - 1);

  for(currentlyCategoryLine = 0; currentlyCategoryLine <= (settingsValue.length); currentlyCategoryLine++){
    // Se le due categorie sono differenti o se si tratta dell'ultima categoria.
    if(currentlyCategoryLine === (settingsValue.length) || settingsValue[currentlyCategoryLine][0] !== settingsValue[firstLineCategory][0]){
      // Se si tratta dell'ultima categoria.
      if(currentlyCategoryLine === (settingsValue.length)){
        // Aumenta il numero di lastLineCategory di uno.
        lastLineCategory++;
        lastSubCategory = 1;
        // Scrive il nome della categoria.
        statsSheet.getRange(initialRowSpace + firstLineCategory, initialColumnSpace).setValue(settingsValue[firstLineCategory][0]);
        // Unisce le celle appartenenti alla stessa categoria.
        statsSheet.getRange(initialRowSpace + firstLineCategory, initialColumnSpace, lastLineCategory-1).mergeVertically();
      }else{
        // Scrive il nome della categoria.
        statsSheet.getRange(initialRowSpace + firstLineCategory, initialColumnSpace).setValue(settingsValue[firstLineCategory][0]);
        // Unisce le celle appartenenti alla stessa categoria.
        statsSheet.getRange(initialRowSpace + firstLineCategory, initialColumnSpace, lastLineCategory).mergeVertically();
        // Scrive il nome della sotto-categoria.
        statsSheet.getRange(initialRowSpace + currentlyCategoryLine, initialColumnSpace+1).setValue(settingsValue[currentlyCategoryLine][1]);
        // Aggiorna il valore di firstLineCategory al fine di individuare la giusta posizione per la categoria successiva.
        firstLineCategory = currentlyCategoryLine;
        for(e = 1; e <= 12; e++){
          statsSheet
            .getRange(initialRowSpace + currentlyCategoryLine, initialColumnSpace+1+e)
            .setValue("= \n" +
              "IFERROR(QUERY('"+ card +"'!$A$4:$O$1000;\" \n" +
              "select sum("+ columnToLetter(moneyColumn) +") \n" +
              "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + firstLineCategory) +"&\"' \n" +
              "and "+ columnToLetter(subCategoryColumn) +" = '\"&C"+ (initialRowSpace + currentlyCategoryLine) +"&\"' \n" +
              "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
              "label sum("+ columnToLetter(moneyColumn) +") ''\");0)"
            );
        }
      }
      // Resetta il valore di lastLineCategory a 1 (e non più 0).
      lastLineCategory = 1;
      if(lastSubCategory === 1){
        statsSheet
          .getRange(initialRowSpace + currentlyCategoryLine, initialColumnSpace)
          .setValue("=SUM("+ columnToLetter(initialColumnSpace + 2) +""+ (initialRowSpace + currentlyCategoryLine) +":"+ columnToLetter(initialColumnSpace + 13) +""+ (initialRowSpace + currentlyCategoryLine) +")");
        statsSheet
          .getRange(initialRowSpace + currentlyCategoryLine, initialColumnSpace, 1, 2).merge();
        for(e = 1; e <= 12; e++){
          statsSheet
            .getRange(initialRowSpace + currentlyCategoryLine, initialColumnSpace + e + 1)
            .setValue("=SUM("+ columnToLetter(initialColumnSpace + e + 1) +""+(initialRowSpace)+":"+ columnToLetter(initialColumnSpace + e + 1) +""+(initialRowSpace + currentlyCategoryLine-1)+")");
        }
      }
    
    }else{
      // Scrive il nome della sotto-categoria.
      statsSheet.getRange(initialRowSpace + currentlyCategoryLine, initialColumnSpace + 1).setValue(settingsValue[currentlyCategoryLine][1]);
      for(e = 1; e <= 12; e++){
        statsSheet
          .getRange(initialRowSpace + currentlyCategoryLine, initialColumnSpace + 1 + e)
          .setValue("= \n" +
            "IFERROR(QUERY('"+ card +"'!$A$4:$O$1000;\"\n" +
            "select sum("+ columnToLetter(moneyColumn) +") \n" +
            "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + firstLineCategory) +"&\"' \n" +
            "and "+ columnToLetter(subCategoryColumn) +" = '\"&C"+ (initialRowSpace + currentlyCategoryLine) +"&\"' \n" +
            "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
            "label sum("+ columnToLetter(moneyColumn) +") ''\");0)"
          );
      }
      // Incrementa di uno il numero di righe da unire in quando le due categorie sono la stessa (e quindi vanno unite).
      lastLineCategory++;

    }
  }
  return initialRowSpace + currentlyCategoryLine;
}
