var ui = SpreadsheetApp.getUi();
function onOpen(){
  
  ui.createMenu('LAB GRADESHEET MENU')
      .addItem('GENERATE LAB GRADESHEET', 'generateGradesheet')
      .addToUi();
  
  processStatusCell.setValue("");
  SpreadsheetApp.flush();
}
function onEdit() {
  processStatusCell.setValue("");
  SpreadsheetApp.flush();
}
