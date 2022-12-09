// Funzione eseguita ad ogni update dello spreadsheet. Serve a generare i menu a tendina delle sottocategorie negli sheets dei conti.
function onEdit(e) {
  // Variabili che selezionano prima la cella che l'utente ha modificato, poi seleziono la colonna.  
  var range = e.range;
  var column = range.getColumn();
  // Controlla se è stata modificata una cella delle colonne 4 o 13 (di qualsiasi foglio). Queste colonne corrispondono a quelle con i menu a tendina delle categorie.
  if(column === 4 || column === 13){
    // solo in tal caso va a definire le successive variabili sempre riferite alla cella modificata.
    var spreadSheet = e.source;
    var sheetName = spreadSheet.getActiveSheet().getName();
    var row = range.getRow();
    var value = e.value;
    // variabile che andrà a contenere i valori da mettere nel menu a tendina delle sottocategorie.
    var returnValues = [];
    
    
    // Se lo sheet è compreso tra gli sheets dei conti e la riga modificata è la 4 (uscite).
    if(banks.includes(sheetName) && column === 4){
      // Variabile che identifica lo sheet con l'elenco delle categorie e sottocategorie per le uscite.
      var dataSheet = ss.getSheetByName("» Expenses: Cat./Subcat.");
      var mainSheet = ss.getSheetByName(sheetName);
      // Ottengo il numero dell'ultima riga che presenta un valore nello sheet.
      var lastRowData = dataSheet.getLastRow(); 
      // Per ogni sottocategoria presa dallo sheet "» Expenses: Cat./Subcat."...
      for(var i = 1; i <= lastRowData; i++){
        // ...verifica se il valore appena messo come categoria nello sheet del conto è uguale alla categoria della sottocategoria "i"...
        if(value == dataSheet.getRange(i, 1).getValue()){
          // ... se vero inserisce la sottocategoria dentro la variabile returnValues.
          returnValues.push(dataSheet.getRange(i, 2).getValue());      
        }
      }
      // Pulisce la cella della sottocategoria nello sheet del conto da possibili vecchi valori già inseriti.
      mainSheet.getRange(row, column+1).clear();
      // Identifica la cella dove mettere il menu a tendina delle sottocategorie nello sheet del conto.
      var dropdown = mainSheet.getRange(row, column+1);
      // Genera il menu a tendina con le validation rule.
      var rule = SpreadsheetApp.newDataValidation().requireValueInList(returnValues).build();
      dropdown.setDataValidation(rule);
    
    
    // Altrimenti se lo sheet è compreso tra gli sheets dei conti e la riga modificata è la 13 (entrate).
    }else if(banks.includes(sheetName) && column === 13){
      // Variabile che identifica lo sheet con l'elenco delle categorie e sottocategorie per le entrate.
      var dataSheet = ss.getSheetByName("» Income: Cat./Subcat.");
      var mainSheet = ss.getSheetByName(sheetName);
      // Ottengo il numero dell'ultima riga che presenta un valore nello sheet.
      var lastRowData = dataSheet.getLastRow();   
      // Per ogni sottocategoria presa dallo sheet "» Income: Cat./Subcat."...
      for(var i = 1; i <= lastRowData; i++){
        // ...verifica se il valore appena messo come categoria nello sheet del conto è uguale alla categoria della sottocategoria "i"...
        if(value == dataSheet.getRange(i, 1).getValue()){
          // ... se vero inserisce la sottocategoria dentro la variabile returnValues.
          returnValues.push(dataSheet.getRange(i, 2).getValue());      
        }
      }
      // Pulisce la cella della sottocategoria nello sheet del conto da possibili vecchi valori già inseriti.
      mainSheet.getRange(row, column+1).clear();
      // Identifica la cella dove mettere il menu a tendina delle sottocategorie nello sheet del conto.
      var dropdown = mainSheet.getRange(row, column+1);
      // Genera il menu a tendina con le validation rule.
      var rule = SpreadsheetApp.newDataValidation().requireValueInList(returnValues).build();
      dropdown.setDataValidation(rule);
    }
  }
}

// Semplice funzione presa online che converte il numero di una colonna in lettera (es. la colonna "2" diventa "B").
function columnToLetter(column){
  var temp, letter = '';
  while (column > 0){
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// Semplice funzione presa online che converte la lettera di una colonna in numero (es. la colonna "B" diventa "2").
function letterToColumn(letter){
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++){
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}
