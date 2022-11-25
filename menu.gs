// Valley - inforge.net //
function onOpen() {
  createMenuWithSubMenu();
}

function createMenuWithSubMenu() {
  /*var subMenu = SpreadsheetApp.getUi().createMenu("Advanced")
    .addItem("Setting D", "settingD")
    .addItem("Setting E", "settingE");
  */
  
  SpreadsheetApp.getUi().createMenu("⚙️ Menu")
    .addItem("Make with ✎ by Valley - Inforge.net", "inforge")
    .addItem("Settings", "settings")
    .addSeparator()
    .addItem("Monthly Income - Total", "monthlyIncomeTotal")
    .addItem("Monthly Expenses - Total", "monthlyExpensesTotal")
    .addSeparator()
    .addItem("Monthly Income - Bank account", "monthlyIncomeBankAccount")
    .addItem("Monthly Expenses - Bank account", "monthlyExpensesBankAccount")
    .addSeparator()
    .addItem("Hype - I/E", "hypeIE")
    .addItem("N26 - I/E", "n26IE")
    .addSeparator()
    .addItem("» Income: Cat./Subcat.", "incomeCatSubcat")
    .addItem("» Expenses: Cat./Subcat.", "expensesCatSubcat")
    .addSeparator()
    /*.addSubMenu(subMenu)*/
    .addToUi();
}

function inforge() {
  return true;
}

function settings() {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ss.getSheetByName('Menu'));
}

function monthlyIncomeTotal() {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ss.getSheetByName('Monthly Income - Total'));
}

function monthlyExpensesTotal() {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ss.getSheetByName('Monthly Expenses - Total'));
}

function monthlyIncomeBankAccount() {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ss.getSheetByName('Monthly Income / Bank account'));
}

function monthlyExpensesBankAccount() {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ss.getSheetByName('Monthly Expenses / Bank account'));
}

function hypeIE() {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ss.getSheetByName('Hype - I/E'));
}

function n26IE() {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ss.getSheetByName('N26 - I/E'));
}

function incomeCatSubcat() {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ss.getSheetByName('» Income: Cat./Subcat.'));
}

function expensesCatSubcat() {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ss.getSheetByName('» Expenses: Cat./Subcat.'));
}
