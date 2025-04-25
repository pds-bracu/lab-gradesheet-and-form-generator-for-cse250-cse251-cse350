// // make everything "Anyone with the link can Edit" before running anything

var isAlreadyGenerated = 0;
var alreadyGeneratedBy;
var semester = "Spring 2024";
var userSpreadsheetId = "1PT1y8ZfyCIZbUIp0RYl8nOBs--jAmp5veYTSjh7_RVo";
var template_250_Gradesheet_Id = "1v4RuDRDXZ0Al0IEuwSQhXz1dD3QBTcOsqslOc1-YrbU";
var template_350_Gradesheet_Id = "1FbOGqRr7BILryUSlMRkB1D4NJT7VppR4r1p6nxQuC4E";
var template_250_Report_Form_Id = "1ucnvUb_xsYeIM9FXbfOTfMqa-rbiA_RfQO15hVrorHE";
var template_350_Report_Form_Id = "16C3ouRDK3AZH96IXhK7yJU6vnxi3MlO0mKFBp0Xb8qg";
var template_250_LtspiceClasswork_Form_Id = "1cXX5m8kHRR6RaIlmPYZAylZZ-SraLRFWezhqbMeahjI";
var urlSpreadsheetId = "";

var storeFolderId = "";  // where all the generated folders get stored
var storeFolder = DriveApp.getFolderById(storeFolderId);

var userSpreadsheet = SpreadsheetApp.openById(userSpreadsheetId);
var userSheet = userSpreadsheet.getSheetByName("User");
var processStatusCell = userSheet.getRange("D2");
var course = userSheet.getRange("B1").getValue();
var templateGradeSheetId;
var templateReportSubmissionFormId;

var userEmail = Session.getActiveUser();

function generateGradesheet() {
  if(course=="CSE250") {
    templateGradeSheetId = template_250_Gradesheet_Id;
    templateReportSubmissionFormId = template_250_Report_Form_Id;
  }
  else if(course=="CSE350") {
    templateGradeSheetId = template_350_Gradesheet_Id;
    templateReportSubmissionFormId = template_350_Report_Form_Id;
  }
  else {
   processStatusCell.setValue("Process terminated: select course (CSE250 or CSE350) and try again.");
   SpreadsheetApp.flush();
   return;
  }
  
  userSpreadsheet.toast("User Email:  " + userEmail);
  var section = Browser.inputBox("Enter section number", Browser.Buttons.OK_CANCEL);
  // // checking if the user section input is an integer and within range
  if(isNaN(section) | section.includes(".")) {
    processStatusCell.setValue("Process terminated: not a valid section");
    SpreadsheetApp.flush();
    Logger.log("Process terminated: not a valid section");
    return;
  }
  else {
    if((course=="CSE250" && (section<1 || section>34 || section==21 || section==31)) || (course=="CSE350" && (section < 1 || section > 16))) {
      processStatusCell.setValue("Process terminated: section out of range");
      SpreadsheetApp.flush();
      Logger.log("Process terminated: section out of range");
      return;
    }
  }


  // // Calling a function to check if already generated
  warnIfAlreadyGenerated(course, section);
  if(isAlreadyGenerated == 1) {
    var response = ui.alert(
      alreadyGeneratedBy + ', probably your co-faculty, has already generated for this section. Do you still want to continue?',
      ui.ButtonSet.YES_NO
    );

    if(response == ui.Button.NO) {
      processStatusCell.setValue("Process terminated at user's will");
      SpreadsheetApp.flush();
      return;
    }
  }
  processStatusCell.setValue("Processing please wait ... This will take just a few seconds");
  SpreadsheetApp.flush();

  userSpreadsheet.toast("Creating folder");
  newFolder = storeFolder.createFolder(course + "-" + section + "L - " + semester);
  newFolder1 = newFolder.createFolder("Forms");

  // // Generating Lab gradesheet
  userSpreadsheet.toast("Creating gradesheet");
  var fileNameGradesheet = course + "-" + section + "L Lab Gradesheet - " + semester;
  var newGradesheetFile = DriveApp.getFileById(templateGradeSheetId).makeCopy(fileNameGradesheet, newFolder);
  var newGradesheet = SpreadsheetApp.openById(newGradesheetFile.getId());
  var metaSheet = newGradesheet.getSheetByName("Meta");
  metaSheet.getRange("B2").setValue(parseInt(section));

  Drive.Permissions.insert(
    {
      'role': 'writer',
      'type': 'anyone',
    },
    newGradesheet.getId(),
  );

  
  saveGradesheetURLs(section, newGradesheet.getUrl(), userEmail);
  
  // // Generating Report submission form
  userSpreadsheet.toast("Creating necessary forms");
  var fileNameReportForm;
  if(course=="CSE250") {
    fileNameReportForm = course + "-" + section + "L Report Submission Form - " + semester;
  }
  else if(course=="CSE350") {
    fileNameReportForm = course + "-" + section + "L Report Submission Form - " + semester;
  }
  var reportSubmissionFormFile = DriveApp.getFileById(templateReportSubmissionFormId).makeCopy(fileNameReportForm, newFolder1);
  var reportSubmissionForm = FormApp.openById(reportSubmissionFormFile.getId());
  reportSubmissionForm.setAcceptingResponses(true);
  metaSheet.getRange("F3").setValue(reportSubmissionForm.getPublishedUrl());
  metaSheet.getRange("F4").setValue(reportSubmissionForm.getEditUrl());
  userSpreadsheet.toast("Almost there");
  setResponseSheet(newFolder1, reportSubmissionForm, newGradesheet, "F5", "Responses: " + fileNameReportForm);


  // Simulation classwork submission form for 250 Only
  // if(course=="CSE250") {
  //   var fileNameClasswork = course + "-" + section + "L Simulation Classwork Submissions " + semester;
  //   var classworkSubmissionFormFile = DriveApp.getFileById(template_250_LtspiceClasswork_Form_Id).makeCopy(fileNameClasswork, newFolder1);
  //   var classworkSubmissionForm = FormApp.openById(classworkSubmissionFormFile.getId());
  //   classworkSubmissionForm.setAcceptingResponses(true);
  //   metaSheet.getRange("F6").setValue(classworkSubmissionForm.getPublishedUrl());
  //   metaSheet.getRange("F7").setValue(classworkSubmissionForm.getEditUrl());
  //   setResponseSheet(newFolder1, classworkSubmissionForm, newGradesheet, "F8", "Responses: " + fileNameClasswork);
  // }
  
  // // making the user owner of the folder, that is, owner of everything generated
  Drive.Permissions.insert(
    {
      'role': 'owner',
      'type': 'user',
      'value': userEmail
    },
    newFolder.getId(),
    {
      'sendNotificationEmails': true
    }
  );

  processStatusCell.setValue("Done! Please check your email.");
  SpreadsheetApp.flush();
  Utilities.sleep(8000);   // 8 ms delay
  processStatusCell.setValue("");
}


// // Functions

// // function to create form response sheet and get url
function setResponseSheet(folder, form, gradesheet, destinationCell, fileName) {
  var responseSheet = SpreadsheetApp.create(fileName);
  responseFile = DriveApp.getFileById(responseSheet.getId());
  responseFile.moveTo(folder);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, responseSheet.getId());
  responseSheet.deleteSheet(responseSheet.getSheetByName("Sheet1"));
  
  // // Making the response sheet as 'anyone can view' due to hassle in allowing access manually
  Drive.Permissions.insert(
    {
      'role': 'reader',
      'type': 'anyone',
    },
    responseSheet.getId(),
  );
  responseSheetUrl = responseSheet.getUrl();
    gradesheet.getSheetByName("Meta").getRange(destinationCell).setValue(responseSheetUrl);
}

// // save the generated gradesheet urls
function saveGradesheetURLs (section, gradesheetUrl, userEmail) {
  var urlSpreadsheet = SpreadsheetApp.openById(urlSpreadsheetId);
  var urlSheet = urlSpreadsheet.getSheetByName(course);
  var data = urlSheet.getRange("A1:A").getValues();
  var lastRowWithData = data.filter(String).length;
  Logger.log(lastRowWithData);
  urlSheet.getRange("A" + (lastRowWithData+1)).setValue(section);
  urlSheet.getRange("B" + (lastRowWithData+1)).setValue(gradesheetUrl);
  urlSheet.getRange("C" + (lastRowWithData+1)).setValue(userEmail);
}

// // Check and warn if the cofaculty has already generated
function warnIfAlreadyGenerated(course, section) {
  var i = 0;
  var generatedSections = SpreadsheetApp.openById(urlSpreadsheetId).getSheetByName(course).getRange("A2:A").getValues().flat();
  var generatorEmails = SpreadsheetApp.openById(urlSpreadsheetId).getSheetByName(course).getRange("D2:D").getValues().flat();
  for(i=0; i<generatedSections.length; i++) {
    if(generatedSections[i] == section) {
      isAlreadyGenerated = 1;
      alreadyGeneratedBy = generatorEmails[i];
      break;
    }
  }
}
