/**
 * Entry point and UI Triggers for the SQD Roster Application.
 */

/**
 * Builds the custom QC HAN UI menu upon opening the spreadsheet.
 */
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("QC HAN")
      .addItem("Publish as new revision", "publishAsNewRevision")
      .addSeparator()
      .addItem("Create new Roster", "createNewRoster")
      .addItem("Update A/C Schedules", "updateACSchedules")
      .addSeparator()
      .addItem("Update Public Roster", "updatePublicRoster")
      .addSeparator()
      .addItem("Xuất bảng chấm công", "ccExport")
      .addToUi();
  } catch (err) {
    console.error("Failed to build menu: " + err.message);
  }
}

/**
 * Displays a modal dialog with a link for the User to easily navigate to the Public roster.
 */
function showDialog() {
  try {
    const html = HtmlService.createHtmlOutput(
      `<a href="https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_IDS.PUBLIC_ROSTER}" target="_blank">Link</a>`
    )
      .setWidth(350)
      .setHeight(50);

    SpreadsheetApp.getUi().showModalDialog(html, 'Done');
  } catch (err) {
    console.error("Error in showDialog: " + err.message);
  }
}

/**
 * Increments the revision string in B2, changes active sheet name, and broadcasts an email to configured recipients.
 */
function publishAsNewRevision() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var ssURL = ss.getUrl();
    
    var currentRevision = sheet.getRange("B2").getValue();
    if (!currentRevision || currentRevision.toString().indexOf('R') === -1) {
      currentRevision = "R0";
    }

    var newRevision = `R${parseInt(currentRevision.toString().slice(1)) + 1}`;
    sheet.getRange("B2").setValue(newRevision);
    
    var currentName = sheet.getName();
    var newCleanName = `${currentName.slice(0, 5)} ${newRevision}`;
    sheet.setName(newCleanName);
    
    var recipientList = getEmailRecipients();
    if (recipientList) {
      MailApp.sendEmail(recipientList, `Schedule ${newCleanName} was updated!`, `Link to Schedule: ${ssURL}`);
    } else {
      SpreadsheetApp.getUi().alert("Warning: No email recipients found in " + CONFIG.SHEET_IDS.EMAILS_SHEET);
    }
    
    // Clear "UNPUBLISHED CHANGES DETECTED" notice if present
    sheet.getRange("C1").clearContent();
    SpreadsheetApp.getUi().alert("Successfully published version " + newCleanName);
  } catch (err) {
    console.error("Error in publishAsNewRevision: " + err.message);
    SpreadsheetApp.getUi().alert("Error during publish: " + err.message);
  }
}
