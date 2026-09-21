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
      .addSeparator()
      .addItem("Get Mobile Link", "showMobileLink")
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
 * Displays the deployed Web App URL so it can be copied/bookmarked for use on a phone.
 * Only resolves once a Web App deployment (test or versioned) exists; see Deploy menu in the
 * Apps Script editor.
 */
function showMobileLink() {
  try {
    var url = ScriptApp.getService().getUrl();
    if (!url) {
      SpreadsheetApp.getUi().alert("No Web App deployment found yet. In the Apps Script editor: Deploy > Test deployments (or New deployment > Web app).");
      return;
    }
    const html = HtmlService.createHtmlOutput(
      `<a href="${url}" target="_blank">${url}</a>`
    )
      .setWidth(350)
      .setHeight(60);
    SpreadsheetApp.getUi().showModalDialog(html, 'Mobile Panel Link');
  } catch (err) {
    console.error("Error in showMobileLink: " + err.message);
  }
}

/**
 * Menu entry point: runs the publish flow and surfaces errors/success via the Sheets UI.
 */
function publishAsNewRevision() {
  try {
    var result = doPublishAsNewRevision();
    if (!result.emailed) {
      SpreadsheetApp.getUi().alert("Warning: No email recipients found in " + CONFIG.SHEET_IDS.EMAILS_SHEET);
    }
    SpreadsheetApp.getUi().alert("Successfully published version " + result.newCleanName);
  } catch (err) {
    console.error("Error in publishAsNewRevision: " + err.message);
    SpreadsheetApp.getUi().alert("Error during publish: " + err.message);
  }
}

/**
 * Increments the revision string in B2, changes active sheet name, and broadcasts an email to configured recipients.
 * Contains no UI calls so it can also be driven from the Web App front end.
 * @returns {{newCleanName: string, emailed: boolean}}
 */
function doPublishAsNewRevision() {
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
    var emailed = false;
    if (recipientList) {
      MailApp.sendEmail(recipientList, `Schedule ${newCleanName} was updated!`, `Link to Schedule: ${ssURL}`);
      emailed = true;
    }

    // Clear "UNPUBLISHED CHANGES DETECTED" notice if present
    sheet.getRange("C1").clearContent();

    return { newCleanName: newCleanName, emailed: emailed };
}
