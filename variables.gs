// variabile che identifica lo spreadsheet.
var ss = SpreadsheetApp.getActiveSpreadsheet();
// I mesi dell'anno e.e
var months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
// Serve ad individuare la giusta posizione per la categoria successiva quando si genera la tabella. Impostato di partenza a 0.
var categorySpace = 0;
// Indica lo spazio verticale (numero di righe) da tenere prima della lista delle categorie.
var initialRowSpace = 5;
// Indica il numero di righe da unire perché appartenenti alla stessa categoria. Impostato di partenza a 0, al primo caso sale sempre a 1. Poi al reset tornerà sempre a 1.
var lastLineCategory = 0;
// Indica lo spazio orizzontale (numero di colonne) da tenere prima della lista delle categorie.
var initialColumnSpace;
// Indica la posizione da utilizzare per la prima linea delle tabelle per le categorie e sottocategorie.
var firstLineCategory = 0;
// Variabile che conserva la scelta dell'utente nella pagina del menu se vuole aggiornare le entrate o le uscite.
var type = ss.getSheetByName("Menu").getRange(3,2).getValue();
// Variabile che conserva i nomi dei fogli dei conti.
var banks = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu").getRange(3,5, 5).getValues();
// Converto l'array da 2D a 1D.
banks = banks.reduce(function(prev, next) {
  return prev.concat(next);
});
// Pulisco l'array da eventuali campi vuoti.
banks = banks.filter(n => n);
