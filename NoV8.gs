/*
  * Copyright Laura Taylor
  * (https://github.com/techstreams/TSWorkflow)
  *
  * Permission is hereby granted, free of charge, to any person obtaining a copy
  * of this software and associated documentation files (the "Software"), to deal
  * in the Software without restriction, including without limitation the rights
  * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
  * copies of the Software, and to permit persons to whom the Software is
  * furnished to do so, subject to the following conditions:
  *
  * The above copyright notice and this permission notice shall be included in all
  * copies or substantial portions of the Software.
  *
  * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
  * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
  * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
  * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
  * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
  * SOFTWARE.
  */

/* 
 * This function adds a 'Purchase Request Workflow' menu to the workflow Sheet when opened
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();  // Sheet UI
  ui.createMenu('Purchase Request Workflow')
    .addSubMenu(ui.createMenu('⚙️ Configure')
      .addItem('⚙️ 1) Setup Workflow Config', 'configure')
      .addSeparator()
      .addItem('⚙️ 2) Setup Request Sheet', 'initialize'))
    .addSeparator()
    .addItem('✏️ Update Request', 'update')
    .addToUi();
}


/* 
 * This function populates the workflow Sheet 'Config' tab with workflow 
 * asset URLs and associates the workflow Form destination with the workflow Sheet
 */
function configure(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(), // active spreadsheet
      configSheet = ss.getSheetByName('Config'),  // config tab
      requestsFolder, requestForm, ssFolder, templateDoc;
  configSheet.activate();
  // Get spreadsheet parent folder - assumes all workflow asset documents in same folder
  ssFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  // Get workflow assets
  templateDoc = ssFolder.getFilesByType(MimeType.GOOGLE_DOCS).next();
  requestForm = ssFolder.getFilesByType(MimeType.GOOGLE_FORMS).next();
  requestsFolder = ssFolder.getFolders().next();
  // Add workflow asset URLs to ‘Config’ tab 
  configSheet.getRange(1, 2, 3).setValues([[requestForm.getUrl()], [templateDoc.getUrl()], [requestsFolder.getUrl()]]);
  // Set the workflow Form destination to the workflow Sheet
  FormApp.openById(requestForm.getId()).setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
}

/* 
 * This function adds additional fields and formatting to the form submission tab
 * and sets up the form submit trigger
 */
function initialize() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),  // active spreadsheet
        formSheet = ss.getSheets()[0],   // form submission tab - assumes first location
        triggerFunction = 'generate';    // name of function for submit trigger
  formSheet.activate();
  // Get form submission tab header row, update background color (yellow) and bold font
  formSheet.getRange(1, 1, 1, formSheet.getLastColumn())
           .setBackground('#fff2cc')
           .setFontWeight('bold');
  // Insert four workflow columns, set header values and update background color (green)
  formSheet.insertColumns(1, 4);
  formSheet.getRange(1, 1, 1, 4)
           .setValues([['Purchase Request Doc', 'Status', 'Status Comments', 'Last Update']])
           .setBackground('#A8D7BB');
  // Set data validation on status column to get dropdown on every form submit entry
  formSheet.getRange('B2:B')
           .setDataValidation(SpreadsheetApp.newDataValidation()
                                            .requireValueInList(['New', 'Pending', 'Approved', 'Declined'], true)
                                            .setHelpText('Please select a status')
                                            .build());
  // Set date format on 'Last Update' column
  formSheet.getRange('D2:D').setNumberFormat("M/d/yyyy hh:mm:ss");
  // Remove any existing form submit triggers and create new  
  ScriptApp.getProjectTriggers().filter(function(trigger) {
              return trigger.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT && trigger.getHandlerFunction() === triggerFunction;
           }).forEach(function(trigger){ ScriptApp.deleteTrigger(trigger) });
  ScriptApp.newTrigger(triggerFunction).forSpreadsheet(ss).onFormSubmit().create();
}

/* 
 * This function generates a new purchase request document from a form submission, 
 * replaces template markers, shares document with requester/supervisor and sends email notification
 * @param {Object} e - event object passed to form submit function
 */
function generate(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),  // active spreadsheet
      configSheet = ss.getSheetByName('Config'),   // config tab
      employeeSheet = ss.getSheetByName('Employees'), // employees stabheet
      formSheet = ss.getSheets()[0],   // form submission tab - assumes first tabsh
      date, doc, email, lastupdate, requestFile, submitDate, viewers;
  // Create and format submit date object from form submission timestamp
  date = new Date(e.namedValues['Timestamp'][0]);
  submitDate = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy hh:mm:ss a (z)");
  // Copy the purchase request template document and move copy to generated requests Drive folder
  requestFile = copyRequestTemplate_(configSheet, 'B2', e.namedValues['Requester Name'][0]);
  moveRequestFile_(configSheet, 'B3', requestFile);
  // Retrieve requester and requester supervisor information for request document sharing and email notifications
  viewers = getViewers_(employeeSheet, e.namedValues['Requester Name'][0]);
  // Open generated request document, replace template markers, update request status and save/close document
  doc = DocumentApp.openById(requestFile.getId());
  replaceTemplateMarkers_(doc, e.namedValues, viewers, submitDate);
  updateStatus_(doc, 'New', submitDate, '');
  // Add requester and supervisor (if exists) to generated request document and set 'VIEW' sharing
  if (viewers.emails.length > 0) {
    requestFile.addViewers(viewers.emails).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);
  }
  // Update workflow request range in form submission tab
  lastupdate = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "M/d/yyyy k:mm:ss");
  formSheet.getRange(e.range.getRow(), 1, 1, 4).setValues([[requestFile.getUrl(), 'New', '', lastupdate]]);
  // Generate notification email body and send to requester, supervisor and Sheet owner
  email = Utilities.formatString('New Purchase Request from: <strong>%s</strong><br><br>See request document <a href="%s">here<\/a>', viewers.requester.name, doc.getUrl());
  viewers.emails.push(Session.getEffectiveUser().getEmail());
  GmailApp.sendEmail(viewers.emails, Utilities.formatString('New %s', doc.getName()), '', { htmlBody: email });
}


/*
 * This function updates the purchase request document with status updates 
 * from form submission tab highlighted row and sends email notification
 */
function update() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),  // active spreadsheet
      configSheet = ss.getSheetByName('Config'),   // config tab
      employeeSheet = ss.getSheetByName('Employees'), // employees tab
      formSheet = ss.getSheets()[0],   // form submission tab - assumes first location
      activeRowRange, activeRowValues, email, date, doc, lastupdate, recipients;
  // Create and format date object for 'last update' timestamp
  date = new Date();
  lastupdate = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy hh:mm:ss a (z)");
  // Get updated workflow request range and process if valid
  activeRowRange = getWorkflowFields_();
  if (activeRowRange) {
    // Get valid workflow request range values
    activeRowValues = activeRowRange.getValues();
    // Get and open associated purchase request document
    doc = DocumentApp.openByUrl(activeRowValues[0][0]);
    // Get emails of document editors and viewers for email notification recipients
    recipients = doc.getEditors().map(function(editor) { return editor.getEmail() })
                                 .concat(doc.getViewers().map(function(viewer) { return viewer.getEmail() }));
    // Get request document status table (last table), populate and save/close       
    updateStatus_(doc, activeRowValues[0][1], lastupdate, activeRowValues[0][2]);
    // Update workflow request range 'Last Update' cell with formatted timestamp
    activeRowValues[0][3] = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "M/d/yyyy k:mm:ss");
    formSheet.getRange(activeRowRange.getRow(), 1, 1, 4).setValues(activeRowValues);
    // Generate notification email body and send to requester, supervisor and Sheet owner
    email = Utilities.formatString('Purchase Request Status Update: <strong>%s</strong><br><br>See request document <a href="%s">here<\/a>', activeRowValues[0][1], doc.getUrl());
    GmailApp.sendEmail(recipients.join(','), Utilities.formatString('Updated Status: %s', doc.getName()), '', { htmlBody: email });
    // Display request update message in Sheet
    ss.toast('Request has been updated.', 'Request Updated!');
  }
}


/* 
 * This function make a copy of the purchase request template and updates the file name
 * @param {Sheet} configSheet - config tab
 * @param {string} configRange - config range for purchase request URL in A1 notation
 * @param {string} requesterName - name of requester from form submission
 * @return {File} Google Drive file
 */
function copyRequestTemplate_(configSheet, configRange, requesterName) {
  var urlParts, templateFile, requestFile;
  // Retrieve purchase request template from Drive
  urlParts = configSheet.getRange(configRange).getValue().split('/');
  templateFile = DriveApp.getFileById(urlParts[urlParts.length - 2]);
  // Make a copy of the request template file and update new file name
  requestFile = templateFile.makeCopy();
  requestFile.setName(Utilities.formatString("Purchase Request - %s", requesterName));
  return requestFile;
}


/* 
 * This function moves the generated purchase request document to the generated requests folder in Google Drive
 * @param {Sheet} configSheet - config tab
 * @param {string} configRange - config range for generated requests folder URL in A1 notation
 * @param {File} requestFile - purchase request file
 */
function moveRequestFile_(configSheet, configRange, requestFile) {
  var urlParts, parentFolders, requestFolder;
  // Retrieve purchase requests folder from Drive
  urlParts = configSheet.getRange(configRange).getValue().split('/');
  requestFolder = DriveApp.getFolderById(urlParts[urlParts.length - 1]);
  // Add copied request file to generated requests folder
  requestFolder.addFile(requestFile);
  // Iterate through request file parent folders and remove file
  // from folders which don't match generated requests folder
  parentFolders = requestFile.getParents();
  while (parentFolders.hasNext()) {
    var f = parentFolders.next();
    if (f.getId() !== requestFolder.getId()) {
      f.removeFile(requestFile);
    }
  }
}


/* 
 * This function iterates over employee data to get requester and supervisor information
 * @param {Sheet} employeeSheet - employee tab
 * @param {string} requesterName - name of requester from form submission
 * @return {Object} requester and supervisor information for request sharing and notifications
 */
function getViewers_(employeeSheet, requesterName) {
  var employees = employeeSheet.getDataRange().getValues(),
      viewers = {},
      supervisor;
  // Shift off header row
  employees.shift();
  // Find form submit requester
  viewers.requester = employees.filter(function(row) { return row[0] === requesterName})
                               .map(function(row) { return { name:row[0], email:row[1], phone:row[2], supervisor:row[3] }})[0];
  viewers.emails = viewers.requester.email !== '' ? [viewers.requester.email] : [];
  // Find requester's supervisor
  supervisor = employees.filter(function(row) { return row[0] === viewers.requester.supervisor} )
                        .map(function(row) { return { name:row[0], email:row[1], phone:row[2] };});
  if (supervisor.length > 0) {
    viewers.supervisor = { name:supervisor[0].name, email:supervisor[0].email, phone:supervisor[0].phone };
    if (supervisor[0].email !== '') {
      viewers.emails.push(supervisor[0].email);
    }
  } else {
    viewers.supervisor = { name: 'N/a', email: 'N/a', phone: 'N/a' };
  }
  return viewers;
}


/* 
 * This function retrieves the workflow request range for selected row (if selection is valid)
 * If selection is invalid display a Sheet message
 * @return {Range} workflow fields range from active selection
*/
function getWorkflowFields_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),  // active spreadsheet
      activeSheet = ss.getActiveSheet(),  // active tab
      activeRowRange = null,  // active range
      activeRange, activeRowNum;
  // Ensure user is on form submission rab - if not show an error and exit
  if (activeSheet.getIndex() !== 1) {
    ss.toast('Select sheet containing purchase requests.', 'Operation Not Valid on Sheet!');
    return activeRowRange;
  }
  // Get the active range (selected row)
  activeRange = activeSheet.getActiveRange();
  // Ensure there is an active row selected - if not show an error and exit
  if (!activeRange) {
    ss.toast('Select a valid row to process.', 'No Row Selected!');
    return activeRowRange;
  }
  // Get the index of first row in the active range
  activeRowNum = activeRange.getRowIndex();
  // Ensure the active row is within the form submission range - if not show an error
  if (activeRowNum === 1 || activeRowNum > activeSheet.getLastRow()) {
    ss.toast('Select a valid row.', 'Selected Row Out Of Range!');
    return activeRowRange;
  }
  // Get the first 4 column range from active row
  activeRowRange = activeSheet.getRange(activeRowNum, 1, 1, 4);
  return activeRowRange;
}

/* 
 * This function replaces request document template markers with values passed from form submission and other data
 * @param {Document} doc - generated request document
 * @param {Object} requestVals - form submission fields
 * @param {Object} viewers - requester and supervisor information
 * @param {string} submitDate - formatted date string
 */
function replaceTemplateMarkers_(doc, requestVals, viewers, submitDate) {
  var docBody = doc.getBody();
  // Replace request document template markers with values passed from form submission
  Object.keys(requestVals).forEach(function(key) {
                             docBody.replaceText(Utilities.formatString("{{%s}}", key), requestVals[key][0]);
                          });
  // Replace submit date, requester and supervisor data
  // NOTE: Requester name replaced by requestVals
  docBody.replaceText("{{Submit Date}}", submitDate);
  docBody.replaceText("{{Requester Email}}", viewers.requester.email);
  docBody.replaceText("{{Requester Phone}}", viewers.requester.phone);
  docBody.replaceText("{{Supervisor Name}}", viewers.supervisor.name);
  docBody.replaceText("{{Supervisor Email}}", viewers.supervisor.email);
  docBody.replaceText("{{Supervisor Phone}}", viewers.supervisor.phone);
}


/* 
 * This function populates the request document status table and saves/closes document
 * @param {Document} doc - generated request document
 * @param {string} status - request status ('New','Pending','Approved','Declined')
 * @param {string} statusDate - formatted date string
 * @param {string} submitComments - request status comments
 */
function updateStatus_(doc, status, statusDate, statusComments) {
  var docBody = doc.getBody(),
      statusTable = docBody.getTables()[2]; 
  statusTable.getRow(0).getCell(1).editAsText().setText(status);
  statusTable.getRow(1).getCell(1).editAsText().setText(statusDate);
  statusTable.getRow(2).getCell(1).editAsText().setText(statusComments);
  doc.saveAndClose();
}

