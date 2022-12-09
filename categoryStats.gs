// Funzione che genera le tabelle mensili per categorie per ogni singolo conto
function categoryStats(type, settingsSheet, statsSheet, categorySpace, card, initialRowSpace, initialColumnSpace, firstLineCategory){
  // Variabile che conserva i valori delle categorie e sottocategorie.
  var settingsValue = settingsSheet.getDataRange().getValues();
  // Indica la colonna con i valori delle categorie da usare per generare le tabelle per le uscite o le entrate.
  var categoryColumn = type == "expenses" ? 4 : type == "income" ? 13 : false;
  // Indica la colonna con i valori da usare per generare le tabelle per le uscite o le entrate.
  var moneyColumn = type == "expenses" ? 3 : type == "income" ? 12 : false;
  // Indica la colonna con i valori dei mesi da usare per generare le tabelle per le uscite o le entrate.
  var monthColumn = type == "expenses" ? 2 : type == "income" ? 11 : false;
  // Variabile che identifica la nuova riga con la quale iniziare la tabella successiva.
  var newInitialRowSpace = categorySpace;
  // Indica lo spazio orizzontale (numero di colonne) da tenere prima della lista delle categorie.
  initialColumnSpace = 3;
  

  // Crea l'intestazione per le categorie e le sotto-categorie.
  statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 2, initialColumnSpace - 1).setValue("Categories");
  // Unisce le celle per mettere l'intestazione "Categories" su due colonne unite.
  statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 2, initialColumnSpace - 1, 1, 2).mergeAcross();;
  //Crea l'intestazione dei mesi.
  for(i = 0; i < 12; i++){
    // Aggiungo i mesi dell'anno nelle intestazioni.
    statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 2, initialColumnSpace + 1 + i).setValue(months[i]);
    // Oltre al mese testuale viene messo anche il mese in numero nella riga sotto (riga che viene nascosta). Questo dato serve per alcune funzioni.
    statsSheet.getRange((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 1, initialColumnSpace + 1 + i).setValue(i + 1);
  }
  // Nascondo la riga con i mesi in numero (fattore estetico).
  statsSheet.hideRows((categorySpace === 0 ? initialRowSpace : initialRowSpace + categorySpace) - 1);
  // Per ogni categoria presente..
  for(currentlyCategoryLine = 0; currentlyCategoryLine <= (settingsValue.length); currentlyCategoryLine++){
    // Se le due categorie sono differenti o se si tratta dell'ultima categoria.
    if(currentlyCategoryLine === (settingsValue.length) || settingsValue[currentlyCategoryLine][0] !== settingsValue[firstLineCategory][0]){
      // Se si tratta dell'ultima categoria.
      if(currentlyCategoryLine === (settingsValue.length)){
        // Scrive il nome della categoria.
        statsSheet.getRange(initialRowSpace + categorySpace,initialColumnSpace - 1).setValue(settingsValue[firstLineCategory][0]);
        // Unisce le celle per mettere la categoria su due colonne unite.
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace - 1, 1, 2).merge();
        // Aggiorna il valore di firstLineCategory al fine di individuare la giusta posizione per la categoria successiva.
        firstLineCategory = currentlyCategoryLine;
        // Per ogni mese..
        for(e = 1; e <= 12; e++){
          // Inserisce la formula dentro la cella.
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
        // Aumenta la variabile CategorySpace di uno.
        categorySpace++;
      }else{
        // Scrive il nome della categoria.
        statsSheet.getRange(initialRowSpace + categorySpace,initialColumnSpace).setValue(settingsValue[firstLineCategory][0]);
        // Unisce le celle per mettere la categoria su due colonne unite.
        statsSheet.getRange(initialRowSpace + categorySpace, initialColumnSpace - 1, 1, 2).merge();
        // Aggiorna il valore di firstLineCategory al fine di individuare la giusta posizione per la categoria successiva.
        firstLineCategory = currentlyCategoryLine;
        // Per ogni mese..
        for(e = 1; e <= 12; e++){
          // Inserisce la formula dentro la cella.
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
        // Aumenta la variabile CategorySpace di uno.
        categorySpace++;
      }
    }
  }
  // Inserisce la riga finale con la somma dei valori per ogni mese.
  statsSheet
    .getRange(initialRowSpace + categorySpace, initialColumnSpace)
    .setValue("=SUM("+ columnToLetter(initialColumnSpace + 1) +""+ (initialRowSpace + categorySpace) +":"+ columnToLetter(initialColumnSpace + 13) +""+ (initialRowSpace + categorySpace) +")");
  // Inserisce la riga finale con la somma dei valori per ogni mese.
  for(e = 1; e <= 12; e++){
    statsSheet
      .getRange(initialRowSpace + categorySpace, initialColumnSpace + e)
      .setValue("=SUM("+ columnToLetter(initialColumnSpace + e) +""+((newInitialRowSpace === 0 ? initialRowSpace : newInitialRowSpace + initialRowSpace))+":"+ columnToLetter(initialColumnSpace + e) +""+(initialRowSpace + categorySpace-1)+")");
  }
  return categorySpace + initialRowSpace;
}
