// Valley - inforge.net //
var ss = SpreadsheetApp.getActiveSpreadsheet();
var months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
// Serve ad individuare la giusta posizione per la categoria successiva quando si genera la tabella. Impostato di partenza a 0.
var categorySpace = 0;
// Indica lo spazio verticale (numero di righe) da tenere prima della lista delle categorie.
var initialRowSpace = 5;
// Indica il numero di righe da unire perché appartenenti alla stessa categoria. Impostato di partenza a 0, al primo caso sale sempre a 1. Poi al reset tornerà sempre a 1.
var lastLineCategory = 0;
// Indica lo spazio orizzontale (numero di colonne) da tenere prima della lista delle categorie.
var initialColumnSpace;
var firstLineCategory = 0;
var banks = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu").getRange(3,5, 5).getValues();
// CONVERTO L'ARRAY 2D A 1D
banks = banks.reduce(function(prev, next) {
  return prev.concat(next);
});
banks = banks.filter(n => n);
