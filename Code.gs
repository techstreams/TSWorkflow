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
 * This function adds a 'Purchase Request Workflow' menu to the Workflow Sheet when opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();  // Sheet UI
  ui.createMenu('Purchase Request Workflow')
    .addSubMenu(ui.createMenu('⚙️ Configure')
      .addItem('⚙️ 1) Setup Workflow Config', 'Workflow.configure')
      .addSeparator()
      .addItem('⚙️ 2) Setup Request Sheet', 'Workflow.initialize'))
    .addSeparator()
    .addItem('✏️ Update Request', 'Workflow.update')
    .addToUi();
}

/* 
 * Workflow Class - Purchase Requests
 */
class Workflow {

  /* 
   * Constructor function
   */
  constructor() {
    const self = this;
    self.ss = SpreadsheetApp.getActiveSpreadsheet();
    self.configSheet = self.ss.getSheetByName('Config');
    self.employeeSheet = self.ss.getSheetByName('Employees');
  }

  /* 
   * This static method populates the Workflow Sheet's 'Config' sheet with workflow 
   * asset URLs and associates the workflow Form destination with the workflow Sheet
   */
  static configure() {
    const workflow = new Workflow();
    workflow.setupConfig_();
  }

  /* 
   * This static method generates a new purchase request document from a form submission, 
   * replaces template markers, shares document with requester/supervisor and sends email notification
   * @param {Object} e - event object passed to onSubmit function
   */
  static generate(e) {
    const workflow = new Workflow();
    let date, doc, email, requestFile, submitDate, viewers;
    // Create and format submit date object from form submission timestamp
    date = new Date(e.namedValues['Timestamp'][0]);
    submitDate = workflow.getFormattedDate_(date, "MM/dd/yyyy hh:mm:ss a (z)");
    // Copy the purchase request template document and move copy to generated requests Drive folder
    requestFile = workflow.copyRequestTemplate_('B2', e.namedValues['Requester Name'][0]);
    workflow.moveRequestFile_('B3', requestFile);
    // Retrieve requester and requester supervisor information for request document sharing and email notifications
    viewers = workflow.getViewers_(e.namedValues['Requester Name'][0]);
    // Open generated request document, replace template markers, update request status and save/close document
    doc = DocumentApp.openById(requestFile.getId());
    workflow.replaceTemplateMarkers_(doc, e.namedValues, viewers, submitDate);
    workflow.updateStatus_(doc, 'New', submitDate);
    // Add requester and supervisor (if exists) to generated request document and set 'VIEW' sharing
    if (viewers.emails.length > 0) {
      requestFile.addViewers(viewers.emails).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);
    }
    // Update workflow request range in form submission sheet
    workflow.updateWorkflowFields_(e.range.getRow(), [[requestFile.getUrl(), 'New', '', workflow.getFormattedDate_(date, "M/d/yyyy k:mm:ss")]]);
    // Generate notification email body and send to requester/supervisor/business owner
    email = `New Purchase Request from: <strong>${viewers.requester.name}<\/strong><br><br>
    See request document <a href="${requestFile.getUrl()}">here<\/a>`;
    viewers.emails.push(Session.getEffectiveUser().getEmail());
    workflow.sendNotification_(viewers.emails, `New ${doc.getName()}`, email);
  }

  /* 
   * This static method adds additional fields and formatting to the form submission sheet
   * and setups the form submit trigger
   * @param {string} triggerFunction - name of trigger function to execute on form submission
   */
  static initialize(triggerFunction = 'Workflow.generate') {
    const workflow = new Workflow();
    workflow.initializeRequestSheet_(triggerFunction);
  }


  /*
   * This static method updates the purchase request document with status updates 
   * from form submission sheet highlighted row and sends email notification
   */
  static update() {
    const workflow = new Workflow();
    let activeRowRange, activeRowValues, email, date, doc, lastupdate, recipients;
    // Create and format date object for 'last update' timestamp
    date = new Date();
    lastupdate = workflow.getFormattedDate_(date, "MM/dd/yyyy hh:mm:ss a (z)");
    // Get updated workflow request range and process if valid
    activeRowRange = workflow.getWorkflowFields_();
    if (activeRowRange) {
      // Get valid workflow request range values
      activeRowValues = activeRowRange.getValues();
      // Get and open associated purchase request document
      doc = DocumentApp.openByUrl(activeRowValues[0][0]);
      // Get emails of document editors and viewers for email notification recipients
      recipients = doc.getEditors()
                      .map(editor => editor.getEmail())
                      .concat(doc.getViewers().map(viewer => viewer.getEmail()));
      // Get request document status table (last table), populate and save/close       
      workflow.updateStatus_(doc, activeRowValues[0][1], lastupdate, activeRowValues[0][2]);
      // Update workflow request range 'Last Update' cell with formatted timestamp
      activeRowValues[0][3] = workflow.getFormattedDate_(date, "M/d/yyyy k:mm:ss");
      workflow.updateWorkflowFields_(activeRowRange.getRow(), activeRowValues);
      // Generate notification email body and send to requester, supervisor and to Sheet owner
      email = `Purchase Request Status Update: <strong>${activeRowValues[0][1]}<\/strong><br><br>
      See request document <a href="${doc.getUrl()}">here<\/a>`;
      workflow.sendNotification_(recipients.join(','), `Updated Status: ${doc.getName()}`, email);
      // Display request update message in Sheet
      workflow.sendSSMsg_('Request has been updated.', 'Request Updated!');
    }
  }

  /* 
   * This method make a copy of the purchase request template and updates the file name
   * @param {string} configRange - config range for purchase request URL in A1 notation
   * @param {string} requesterName - name of requester from form submission
   * @return {File} Google Drive file
   */
  copyRequestTemplate_(configRange, requesterName) {
    const self = this;
    let urlParts, templateFile, requestFile;
    // Retrieve purchase request template from Drive
    urlParts = self.configSheet.getRange('B2').getValue().split('/');
    templateFile = DriveApp.getFileById(urlParts[urlParts.length - 2]);
    // Make a copy of the request template file and update new file name
    requestFile = templateFile.makeCopy();
    requestFile.setName(`Purchase Request - ${requesterName}`);
    return requestFile;
  }

  /* 
   * This method adds additional fields and formatting to the form submission sheet and setups the submit trigger
   * @param {string} triggerFunction - name of trigger function to execute on form submission
   * @return {Workflow} this object for chaining
   */
  initializeRequestSheet_(triggerFunction) {
    const self = this,  // active spreadsheet
          formSheet = self.ss.getSheets()[0];   // form submission sheet - assumes first sheet
    formSheet.activate();
    // Get form submission sheet header row, update background color (yellow) and fold font
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
    ScriptApp.getProjectTriggers()
      .filter(trigger => trigger.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT && trigger.getHandlerFunction() === triggerFunction)
      .forEach(trigger => ScriptApp.deleteTrigger(trigger));
    ScriptApp.newTrigger(triggerFunction)
      .forSpreadsheet(self.ss)
      .onFormSubmit()
      .create();
    return self;
  }

  /* 
   * This method formats a date using the Google Sheet timezone
   * @param {Date} date - Javascript date object
   * @param {string} format - string representing the desired date format
   * @return {string} formatted date string
   */
  getFormattedDate_(date, format) {
    const self = this;
    return Utilities.formatDate(date, self.ss.getSpreadsheetTimeZone(), format);
  }

  /* 
   * This method iterates over employee data to get requester and supervisor information
   * @param {string} requesterName - name of requester from form submission
   * @return {Object} requester and supervisor information for request sharing and notifications
   */
  getViewers_(requesterName) {
    const self = this,
          employees = self.employeeSheet.getDataRange().getValues(),
          viewers = {};
    let supervisor;
    // Shift off header row
    employees.shift();
    // Find requester who submitted the form request
    viewers.requester = employees.filter(row => row[0] === requesterName)
                       .map((row) => ({ name: row[0], email: row[1], phone: row[2], supervisor: row[3] }))[0];
    viewers.emails = viewers.requester.email !== '' ? [viewers.requester.email] : [];
    // Find requester's supervisor
    supervisor = employees.filter(row => row[0] === viewers.requester.supervisor)
                 .map((row) => ({ name: row[0], email: row[1], phone: row[2] }));
    if (supervisor.length > 0) {
      viewers.supervisor = { name: supervisor[0].name, email: supervisor[0].email, phone: supervisor[0].phone };
      if (supervisor[0].email !== '') {
        viewers.emails.push(supervisor[0].email);
      }
    } else {
      viewers.supervisor = { name: 'N/a', email: 'N/a', phone: 'N/a' };
    }
    return viewers;
  }

  /* 
   * This method retrieves the workflow request range of selected row (if selection is valid)
   * If selection is invalid display a Sheet message
   * @return {Range} workflow fields range from active selection
   */
  getWorkflowFields_() {
    const self = this,
          activeSheet = self.ss.getActiveSheet();
    let activeRowRange = null, activeRange, activeRowNum;
    // Ensure user is on form submission sheet - if not show an error and exit
    if (activeSheet.getIndex() !== 1) {
      self.sendSSMsg_('Select sheet containing purchase requests.', 'Operation Not Valid on Sheet!');
      return activeRowRange;
    }
    // Get the active range (selected row)
    activeRange = activeSheet.getActiveRange();
    // Ensure there is an active row selected - if not show an error and exit
    if (!activeRange) {
      self.sendSSMsg_('Select a valid row to process.', 'No Row Selected!');
      return activeRowRange;
    }
    // Get the index of first row in the active range
    activeRowNum = activeRange.getRowIndex();
    // Ensure the active row is within the form submission range - if not show an error
    if (activeRowNum === 1 || activeRowNum > activeSheet.getLastRow()) {
      self.sendSSMsg_('Select a valid row.', 'Selected Row Out Of Range!');
      return activeRowRange;
    }
    // Get the first 4 column range from active row
    activeRowRange = activeSheet.getRange(activeRowNum, 1, 1, 4);
    return activeRowRange;
  }

  /* 
   * This method moves the generated purchase request document to the generated requests folder in Google Drive
   * @param {string} configRange - config range for generated requests folder URL in A1 notation
   * @param {File} requestFile - purchase request file
   * @return {Workflow} this object for chaining
   */
  moveRequestFile_(configRange, requestFile) {
    const self = this;
    let urlParts, parentFolders, requestFolder;
    // Retrieve purchase requests folder from Drive
    urlParts = self.configSheet.getRange(configRange).getValue().split('/');
    requestFolder = DriveApp.getFolderById(urlParts[urlParts.length - 1]);
    // Add copied request file to generated requests folder
    requestFolder.addFile(requestFile);
    // Iterate through request file parent folders and remove file
    // from folders which don't match generated requests folder
    parentFolders = requestFile.getParents();
    while (parentFolders.hasNext()) {
      let f = parentFolders.next();
      if (f.getId() !== requestFolder.getId()) {
        f.removeFile(requestFile);
      }
    }
    return self;
  }

  /* 
   * This method replaces request document template markers with both values passed from form submission and other data
   * @param {Document} doc - generated request document
   * @param {Object} requestVals - form submission fields
   * @param {Object} viewers - requester and supervisor information
   * @param {string} submitDate - formatted date string
   * @return {Workflow} this object for chaining
   */
  replaceTemplateMarkers_(doc, requestVals, viewers, submitDate) {
    const self = this,
          docBody = doc.getBody();
    // Replace request document template markers with values passed from form submission
    Object.keys(requestVals).forEach(key => docBody.replaceText(Utilities.formatString("{{%s}}", key), requestVals[key][0]));
    // Replace submit date, requester and supervisor data
    // NOTE: Requester name replaced by requestVals
    docBody.replaceText("{{Submit Date}}", submitDate);
    docBody.replaceText("{{Requester Email}}", viewers.requester.email);
    docBody.replaceText("{{Requester Phone}}", viewers.requester.phone);
    docBody.replaceText("{{Supervisor Name}}", viewers.supervisor.name);
    docBody.replaceText("{{Supervisor Email}}", viewers.supervisor.email);
    docBody.replaceText("{{Supervisor Phone}}", viewers.supervisor.phone);
    return self;
  }

  /* 
   * This method sends email notifications
   * @param {string} emails - comma separated list of recipient emails
   * @param {string} subject - email subject 
   * @param {string} emailBody - email message body
   * @return {Workflow} this object for chaining
   */
  sendNotification_(emails, subject, emailBody) {
    const self = this;
    GmailApp.sendEmail(emails, subject, '', { htmlBody: emailBody });
    return self;
  }

  /* 
   * This method displays Sheet messages with toast()
   * @param {string} message - message content
   * @param {string} title - message title 
   * @return {Workflow} this object for chaining
   */
  sendSSMsg_(msg, title) {
    const self = this;
    self.ss.toast(msg, title);
    return self;
  }

  /* 
   * This method populates the 'Config' sheet with workflow asset URLs 
   * and associates the workflow Form destination with the workflow Sheet
   * @return {Workflow} this object for chaining
   */
  setupConfig_() {
    const self = this;
    let requestsFolder, requestForm, ssFolder, templateDoc;
    self.configSheet.activate();
    // Get spreadsheet parent folder - assumes all workflow documents in folder
    ssFolder = DriveApp.getFileById(self.ss.getId()).getParents().next();
    // Get workflow assets
    templateDoc = ssFolder.getFilesByType(MimeType.GOOGLE_DOCS).next();
    requestForm = ssFolder.getFilesByType(MimeType.GOOGLE_FORMS).next();
    requestsFolder = ssFolder.getFolders().next();
    // Add workflow asset URLs to ‘Config’ sheet 
    self.configSheet.getRange(1, 2, 3).setValues([[requestForm.getUrl()], [templateDoc.getUrl()], [requestsFolder.getUrl()]]);
    // Set the workflow Form destination to the workflow Sheet
    FormApp.openById(requestForm.getId()).setDestination(FormApp.DestinationType.SPREADSHEET, self.ss.getId());
    return self;
  }

  /* 
   * This method populates the request document status table and saves/closes document
   * @param {Document} doc - generated request document
   * @param {string} status - request status ('New','Pending','Approved','Declined')
   * @param {string} statusDate - formatted date string
   * @param {string} submitComments - request status comments
   * @return {Workflow} this object for chaining
   */
  updateStatus_(doc, status, statusDate, statusComments = '') {
    const self = this,
          docBody = doc.getBody(),
          statusTable = docBody.getTables()[2]; 
    statusTable.getRow(0).getCell(1).editAsText().setText(status);
    statusTable.getRow(1).getCell(1).editAsText().setText(statusDate);
    statusTable.getRow(2).getCell(1).editAsText().setText(statusComments);
    doc.saveAndClose();
    return self;
  }

  /* 
   * This method updates the selected request workflow range in the form submission sheet
   * @param {number} row - selected request row number
   * @param {string[][]} vals - two-dimensional array of workflow field values to be written to selected row
   * @return {Workflow} this object for chaining
   */
  updateWorkflowFields_(row, vals) {
    const self = this,
          formSheet = self.ss.getSheets()[0];
    formSheet.getRange(row, 1, 1, 4).setValues(vals);
    return self;
  }

}
