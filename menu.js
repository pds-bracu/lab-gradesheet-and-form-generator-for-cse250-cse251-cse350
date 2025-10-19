function onOpen(){
  var ui = getUI();
  ui.createMenu('LAB GRADESHEET MENU')
      .addItem('GENERATE LAB GRADESHEET', 'main')
      .addToUi();
  
  processStatusCell.setValue("");
  SpreadsheetApp.flush();
}
function onEdit() {
  processStatusCell.setValue("");
  SpreadsheetApp.flush();
}
