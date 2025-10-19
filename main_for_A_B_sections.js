function main() {
  var ui = getUI();
  var course = userSheet.getRange("B1").getValue();
  const templates = templatesAllCourse[course];
  if(!templates) {
    processStatusCell.setValue("Select course (CSE250 or CSE251 or CSE350) and try again.");
    SpreadsheetApp.flush();
    return;
  }
  storeFolder = DriveApp.getFolderById(templates.destinationFolderId);
  userSpreadsheet.toast("User:  " + userEmail);
  
  
  var section = Browser.inputBox("Enter section", Browser.Buttons.OK_CANCEL);
  var pattern = /^\d{2}[AB]$/;
  if(!pattern.test(section)) {
    processStatusCell.setValue("Invalid section format. Please enter something like 01A or 10B");
    SpreadsheetApp.flush();
    return;
  };
  
  if(notifyIfAlreadyGenerated(course, section).isAlreadyGenerated) {
    var response = ui.alert(
      notifyIfAlreadyGenerated(course, section).generatedBy + ` has already generated for ${course}-${section}L. Do you still want to continue?`,
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

  createGradesheet(course, section, templates, storeFolder);
}

function createGradesheet(course, section, templates, storeFolder) {
  userSpreadsheet.toast("Creating folder");
  var secFolder = storeFolder.createFolder(course + "-" + section + "L - " + semester);
  var formFolder = secFolder.createFolder("Report Submission");
  userSpreadsheet.toast("Creating gradesheet");
  var fileNameGradesheet = course + "-" + section + "L Lab Gradesheet - " + semester;
  var newGradesheetFile = DriveApp.getFileById(templates.gradesheetId).makeCopy(fileNameGradesheet, secFolder);
  var newGradesheet = SpreadsheetApp.openById(newGradesheetFile.getId());
  
  var metaSheet = newGradesheet.getSheetByName("Meta");
  metaSheet.getRange("B2").setValue(section);

  saveGradesheetURLs(course, section, newGradesheet.getUrl());
  createForms(course, section, newGradesheet, templates, formFolder);

  permission(newGradesheet.getId(), 'reader', 'anyone', null, false);
  permission(secFolder.getId(), 'writer', 'user', getCofacultyEmail(newGradesheet), false);
  permission(secFolder.getId(), 'writer', 'user', 'purbayan.das@bracu.ac.bd', false);
  permission(secFolder.getId(), 'owner', 'user', userEmail, true);

  processStatusCell.setValue("Done! Please check your email or the course folders attached below.");
  SpreadsheetApp.flush();
  Utilities.sleep(5000);
  processStatusCell.setValue("");
  SpreadsheetApp.flush();
}

function getCofacultyEmail(gradesheet) {
  var facultyEmails = SpreadsheetApp.openById(urlSpreadsheetId).getSheetByName("Faculty").getRange("D2:D").getValues().flat();
  var facultyInitials = SpreadsheetApp.openById(urlSpreadsheetId).getSheetByName("Faculty").getRange("A2:A").getValues().flat();
  var userInitial = facultyInitials[facultyEmails.indexOf(userEmail.getEmail().toLowerCase())];
  var secFaculyIntials = gradesheet.getSheetByName("Meta").getRange("B3").getValue().split(",").map(s => s.trim());
  
  var coFacultyInitial = secFaculyIntials[0] === userInitial ? secFaculyIntials[1] : secFaculyIntials[0];
  return facultyEmails[facultyInitials.indexOf(coFacultyInitial)];
}

function permission(fileId, role, type, email, sentNotification) {
  
  const permissionDetails = {
    'role': role,
    'type': type,
  };
  if(type === 'anyone') {
    permissionDetails.withLink = true;
  }
  if(type === 'user' || type === 'group') {
    permissionDetails.value = email;
  }

  Drive.Permissions.insert(
    permissionDetails,  
    fileId,
    {
      'sendNotificationEmails': sentNotification
    }
  );
}

function createForms(course, section, gradesheet, templates, folder){
  userSpreadsheet.toast("Creating report submission form");
  var fileNameReportForm = course + "-" + section + "L Report Submission Form - " + semester;
  var reportSubmissionFormFile = DriveApp.getFileById(templates.reportFormId).makeCopy(fileNameReportForm, folder);
  var reportSubmissionForm = FormApp.openById(reportSubmissionFormFile.getId());
  reportSubmissionForm.setPublished(true);
  reportSubmissionForm.setAcceptingResponses(true);
  reportSubmissionForm.setProgressBar(true);
  gradesheet.getSheetByName("Meta").getRange("F3").setValue(reportSubmissionForm.shortenFormUrl(reportSubmissionForm.getPublishedUrl()));
  gradesheet.getSheetByName("Meta").getRange("F4").setValue(reportSubmissionForm.getEditUrl());
  userSpreadsheet.toast("Almost there");
  setResponseSheet(folder, reportSubmissionForm, gradesheet, "Responses: " + fileNameReportForm);
}

function setResponseSheet(folder, form, gradesheet, fileName) {
  var responseSheet = SpreadsheetApp.create(fileName);
  responseSheetFile = DriveApp.getFileById(responseSheet.getId());
  form.setDestination(FormApp.DestinationType.SPREADSHEET, responseSheet.getId());
  responseSheet.deleteSheet(responseSheet.getSheetByName("Sheet1"));
  responseSheetFile.moveTo(folder);
  gradesheet.getSheetByName("Meta").getRange("F5").setValue(responseSheet.getUrl());
  
  permission(responseSheet.getId(), 'reader', 'anyone', null, false);
}

function saveGradesheetURLs(course, section, gradesheetUrl) {
  var urlSpreadsheet = SpreadsheetApp.openById(urlSpreadsheetId);
  var urlSheet = urlSpreadsheet.getSheetByName(course);
  var data = urlSheet.getRange("A1:A").getValues();
  var lastRowWithData = data.filter(String).length;
  urlSheet.getRange("A" + (lastRowWithData+1)).setValue(section);
  urlSheet.getRange("B" + (lastRowWithData+1)).setValue(gradesheetUrl);
  urlSheet.getRange("C" + (lastRowWithData+1)).setValue(userEmail);
  urlSheet.getRange("E" + (lastRowWithData+1)).setValue(new Date());
}

function checkIfSectionIsValid(course, section) {
  var secRules = secConstraintsAllCourses[course];
  if(isNaN(section) || section < secRules.min || section > secRules.max || secRules.missing.includes(section)) {
    processStatusCell.setValue(`The section ${section} is closed or doesn't exist`);
    SpreadsheetApp.flush();
    return false;
  }
  return true;
}

function notifyIfAlreadyGenerated(course, section) {
  var generatedSections = SpreadsheetApp.openById(urlSpreadsheetId).getSheetByName(course).getRange("A2:A").getValues().flat();
  var generatorIntial = SpreadsheetApp.openById(urlSpreadsheetId).getSheetByName(course).getRange("D2:D").getValues().flat();

  for(let i=0; i<generatedSections.length; i++) {
    if(generatedSections[i] == section) {
      return {
        isAlreadyGenerated: true,
        generatedBy: generatorIntial[i]
      };
    }
  }
  return {
    isAlreadyGenerated: false,
    generatedBy: null
  };
}
