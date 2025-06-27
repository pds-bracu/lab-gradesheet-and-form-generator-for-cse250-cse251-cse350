var semester = "Summer 2025";
var userSpreadsheetId = "1gZOXAFe_7TbV-WeuVg21c0F";
var urlSpreadsheetId = "1ib4Xw_0WhWD-dVkd6jOT1N";

const templatesAllCourse = {
  "CSE250": {
    gradesheetId: "1oEdfR3B8wmO-t2Cz8DwxFEqxmveMjj",
    reportFormId: "1gj599_RAAW8f3O8Z",
    destinationFolderId: "10du4hj4"
  },
  "CSE251": {
    gradesheetId: "1nJHPHn09mN1Nw7LO4",
    reportFormId: "1WANV7111zXNgBfXgcxThp",
    destinationFolderId: "1NwzGTATo"
  },
  "CSE350": {
    gradesheetId: "17zyhJKCD-vAar1e",
    reportFormId: "118A_McgLEciifblJ",
    destinationFolderId: "1FDF7Gp5SjV4Lo"
  },
};

const secConstraintsAllCourses = {
  "CSE250": {
    min: 1,
    max: 24,
    missing: [14, 16]
  },
  "CSE251": {
    min: 1,
    max: 24,
    missing: [13, 14]
  },
  "CSE350": {
    min: 1,
    max: 17,
    missing: []
  },
};

var userEmail = Session.getActiveUser();
var userSpreadsheet = SpreadsheetApp.openById(userSpreadsheetId);
var userSheet = userSpreadsheet.getSheetByName("User");
var processStatusCell = userSheet.getRange("D2");
var ui = SpreadsheetApp.getUi();
