function netWorthStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var netWorthSheet = ss.getSheetByName("Annual NW");
  // Indica lo spazio verticale (numero di righe) da tenere prima della lista delle categorie.
  var initialRowSpace = 4;
  // Indica lo spazio orizzontale (numero di colonne) da tenere prima della lista delle categorie.
  var initialColumnSpace = 3;
  var months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

  // Crea l'intestazione per le categorie e le sotto-categorie.
  netWorthSheet.getRange(initialRowSpace - 2, initialColumnSpace - 1).setValue("Asset/Liability");
  //Crea l'intestazione dei mesi.
  for(i = 0; i < 12; i++){
    netWorthSheet.getRange(initialRowSpace - 2, initialColumnSpace + i).setValue(months[i]);
    netWorthSheet.getRange(initialRowSpace - 1, initialColumnSpace + i).setValue(i + 1);
  }
  netWorthSheet.hideRows(initialRowSpace - 1);
}
