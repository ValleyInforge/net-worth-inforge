// Valley - inforge.net //
// Funzione che genera le tabelle mensili per categorie per ogni singolo conto
function categoryStats(type, settingsSheet, statsSheet, categorySpace, card, initialRowSpace, initialColumnSpace, firstLineCategory){
  var settingsValue = settingsSheet.getDataRange().getValues();
  // Indica la colonna con i valori delle categorie da usare per generare le tabelle per le uscite o le entrate.
  var categoryColumn = type == "expenses" ? 4 : type == "income" ? 13 : false;
  // Indica la colonna con i valori da usare per generare le tabelle per le uscite o le entrate.
  var moneyColumn = type == "expenses" ? 3 : type == "income" ? 12 : false;
  // Indica la colonna con i valori dei mesi da usare per generare le tabelle per le uscite o le entrate.
  var monthColumn = type == "expenses" ? 2 : type == "income" ? 11 : false;
  var newInitialRowSpace = categorySpace;
  // Indica lo spazio orizzontale (numero di colonne) da tenere prima della lista delle categorie.
  initialColumnSpace = 3;
  

  // Crea l'intestazione per le categorie e le sotto-categorie.
  statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 2, initialColumnSpace - 1).setValue("Categories");
  statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 2, initialColumnSpace - 1, 1, 2).mergeAcross();;
  //Crea l'intestazione dei mesi.
  for(i = 0; i < 12; i++){
    statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 2, initialColumnSpace + 1 + i).setValue(months[i]);
    statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 1, initialColumnSpace + 1 + i).setValue(i + 1);
  }
  statsSheet.hideRows((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 1);
  for(currentlyCategoryLine = 0; currentlyCategoryLine <= (settingsValue.length); currentlyCategoryLine++){
    // Se le due categorie sono differenti o se si tratta dell'ultima categoria.
    if(currentlyCategoryLine === (settingsValue.length) || settingsValue[currentlyCategoryLine][0] !== settingsValue[firstLineCategory][0]){
      // Se si tratta dell'ultima categoria.
      if(currentlyCategoryLine === (settingsValue.length)){
        // Scrive il nome della categoria.
        statsSheet.getRange(initialRowSpace + categorySpace,initialColumnSpace - 1).setValue(settingsValue[firstLineCategory][0]);
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace - 1, 1, 2).merge();
        // Aggiorna il valore di firstLineCategory al fine di individuare la giusta posizione per la categoria successiva.
        firstLineCategory = currentlyCategoryLine;
        for(e = 1; e <= 12; e++){
          statsSheet
            .getRange(initialRowSpace + categorySpace, initialColumnSpace + e)
            .setValue("= \n" +
              "IFERROR(QUERY('"+ card +"'!$A$4:$O$1000;\" \n" +
              "select sum("+ columnToLetter(moneyColumn) +") \n" +
              "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + categorySpace) +"&\"' \n" +
              "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
              "label sum("+ columnToLetter(moneyColumn) +") ''\");0)"
            );
        }
        categorySpace++;
      }else{
        // Scrive il nome della categoria.
        statsSheet.getRange(initialRowSpace + categorySpace,initialColumnSpace).setValue(settingsValue[firstLineCategory][0]);
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace - 1, 1, 2).merge();
        // Aggiorna il valore di firstLineCategory al fine di individuare la giusta posizione per la categoria successiva.
        firstLineCategory = currentlyCategoryLine;
        for(e = 1; e <= 12; e++){
          statsSheet
            .getRange(initialRowSpace + categorySpace, initialColumnSpace + e)
            .setValue("= \n" +
              "IFERROR(QUERY('"+ card +"'!$A$4:$O$1000;\" \n" +
              "select sum("+ columnToLetter(moneyColumn) +") \n" +
              "where "+ columnToLetter(categoryColumn) +" = '\"&B"+ (initialRowSpace + categorySpace) +"&\"' \n" +
              "and "+ columnToLetter(monthColumn) +" = \"&"+ columnToLetter(e + 3) +"4&\" \n" +
              "label sum("+ columnToLetter(moneyColumn) +") ''\");0)"
            );
        }
        categorySpace++;
      }
    }
  }
  statsSheet
    .getRange(initialRowSpace + categorySpace, initialColumnSpace)
    .setValue("=SUM("+ columnToLetter(initialColumnSpace + 1) +""+ (initialRowSpace + categorySpace) +":"+ columnToLetter(initialColumnSpace + 13) +""+ (initialRowSpace + categorySpace) +")");
  for(e = 1; e <= 12; e++){
    statsSheet
      .getRange(initialRowSpace + categorySpace, initialColumnSpace + e)
      .setValue("=SUM("+ columnToLetter(initialColumnSpace + e) +""+((newInitialRowSpace === 0 ? initialRowSpace : newInitialRowSpace + initialRowSpace))+":"+ columnToLetter(initialColumnSpace + e) +""+(initialRowSpace + categorySpace-1)+")");
  }
  return categorySpace + initialRowSpace;
}
