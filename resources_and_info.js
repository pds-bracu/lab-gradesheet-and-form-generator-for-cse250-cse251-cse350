var semester = "Fall 2025";
var userSpreadsheetId = "1GgvR9vRk68b0s5l2F-pdIhcCs...";
var urlSpreadsheetId = "1H4i1jf_CH2m_tixXQ7FURqPpD...";

const templatesAllCourse = {
  "CSE250": {
    gradesheetId: "1nvMUS_6sdnDRwMe9a0er1IE9W10q_...",
    reportFormId: "1QtQCRakEqgc8li6pdPmI...",
    destinationFolderId: "1pvE4TbgYB31nJl..."
  },
  "CSE251": {
    gradesheetId: "1PpTlVrv-N69iASREc5hwUM4mF7...",
    reportFormId: "1qDhm12ixoNMu2td969jzZFYPlju...",
    destinationFolderId: "1gLoGhqoO3HFABPPYdSFtGJhCGXpavkAC"
  },
  "CSE350": {
    gradesheetId: "1l0SuoAIH9F5zQTq_wluo7DKmXm1...",
    reportFormId: "1VOB3UnmiCFqZSsXeXv9xibofQ9...",
    destinationFolderId: "1tmYYL1dD_sSLh0VykF..."
  },
};

var userEmail = Session.getActiveUser();
var userSpreadsheet = SpreadsheetApp.openById(userSpreadsheetId);
var userSheet = userSpreadsheet.getSheetByName("User");
var processStatusCell = userSheet.getRange("D2");

function getUI() {
  return SpreadsheetApp.getUi();
}

// The following is optional
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
