var semester = "Summer 2025";
var userSpreadsheetId = "1gZOXAFe_7TbV-WeuVg21c0FB-Puy-7lFl97srzbZ-F4";
var urlSpreadsheetId = "1ib4Xw_0WhWD-dVkd6jOT1NSNMK_C1F05UW1JNYc98eg";

const templatesAllCourse = {
  "CSE250": {
    gradesheetId: "1oEdfR3B8wmO-t2Cz8DwxFEqxmveMjjgRy04E5udZvZc",
    reportFormId: "1gj599_RAAW8f3O8ZDV665bzGQx2PAS_eTlifHKpjoWE",
    destinationFolderId: "10du4hj4nMnrIiLr8lKoaGfT--MpZKtkO"
  },
  "CSE251": {
    gradesheetId: "1nJHPHn09mN1Nw7LO4wrtcW_hXwqjTfDhqnmP-GNj6H8",
    reportFormId: "1WANV7111zXNgBfXgcxThp7XMz6ajTS3I8YiqB0Itod4",
    destinationFolderId: "1NwzGTATo86FZe7QubL3JkVpk7WUTYTmP"
  },
  "CSE350": {
    gradesheetId: "17zyhJKCD-vAar1eQyMT7GzDD0CECSNFoWAd3Hvlp04U",
    reportFormId: "118A_McgLEciifblJH6Vg2Thk--E0zT1W_xSmk69cFxw",
    destinationFolderId: "1FDF7Gp5SjV4Lov34TS2FyPI_29j2BHVU"
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
