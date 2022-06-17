// Copyright 2018 Google Inc. All Rights Reserved.
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

/**
 * @fileoverview DCM Sheets addon syncs reports from DCM
 * at scheduled intervals to the spreadsheet.
 */

/**
 * Adds a custom menu to the spreadsheet.
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createMenu('CM_Reporting')
  menu.addItem('Select CM Report', 'showDialog')
      .addItem('Refresh current sheet', 'refreshCurrentSheet')
      .addItem('Refresh all sheets', 'refreshAllSheets')
      .addItem('Check Last sync time', 'showLastSyncDetails')
      .addItem('Schedule reports', 'schedulerDialog')
      .addSeparator()
      .addItem('Debug', 'debugInfo')
      .addSeparator()
      .addItem('Purge Properties', 'purgeProperties')
      .addToUi();
}

function debugInfo() {
  var documentProperties = PropertiesService.getDocumentProperties();
  var prop = documentProperties.getProperty(PROP_DCM_TRIGGER_ID)
  if (prop != null) {
    SpreadsheetApp.getUi().alert(prop);
  } else {
    SpreadsheetApp.getUi().alert('property is null');
  }
}

/**
 * Remove all properties and triggers in the current document.
 */
function purgeProperties() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
      'Are you sure?',
      'You will need to authenticate, link and schedule your reports' +
          ' again\nin this spreadsheet. Please confirm if you wish to continue.',
      ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    removeDocumentTriggers();
    unlinkAllReports();
    clearToken();
    PropertiesService.getDocumentProperties().deleteAllProperties();
    PropertiesService.getUserProperties().deleteAllProperties();
  }
}

/**
 * Remove any triggers that are currently set on the document.
 */
function removeDocumentTriggers() {
  var documentProperties = PropertiesService.getDocumentProperties();
  var triggers =
      ScriptApp.getUserTriggers(SpreadsheetApp.getActiveSpreadsheet());

  // Delete all triggers for this user for this add-on.
  for (i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  documentProperties.deleteProperty(PROP_DCM_SCHEDULE_FREQUENCY);
  documentProperties.deleteProperty(PROP_DCM_SCHEDULE_TIME);
  documentProperties.deleteProperty(PROP_DCM_SCHEDULE_TIME2);
  documentProperties.deleteProperty(PROP_DCM_TRIGGER_CREATED_BY);
  documentProperties.deleteProperty(PROP_DCM_TRIGGER_ID);

  SpreadsheetApp.getActiveSpreadsheet().toast(
      'All current document triggers deleted!');
}

/**
 * Runs when the add-on is installed. Calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  onOpen(e);
}

function checkAndGetOAuthService() {
  var oauthService = getOAuthService();
  if (!oauthService.hasAccess()) {
    var authorizationUrl = oauthService.getAuthorizationUrl();
    var template = HtmlService.createTemplate(
        '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
        'Please authenticate and then run the command again.');
    template.authorizationUrl = authorizationUrl;
    var page = template.evaluate();
    SpreadsheetApp.getUi().showModalDialog(page, 'Needs authentication');
    return null;
  } else {
    return oauthService;
  }
}

/**
 * Opens a dialog. The dialog structure is described in the Dialog.html or
 * SUDialog.html.
 */
function showDialog() {
  // Check if user needs authentication.
  var oauthService = checkAndGetOAuthService();
  
  if (!oauthService) {
    return;
  }
  
  var ui = HtmlService.createTemplateFromFile('Dialog')
      .evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(500)
      .setHeight(130);
  SpreadsheetApp.getUi().showModalDialog(ui, DIALOG_TITLE);
}

/**
 * Function to show ad dialog with last sync details.
 */
function showLastSyncDetails() {
  var ui = HtmlService.createTemplateFromFile('LastSyncDetails')
               .evaluate()
               .setSandboxMode(HtmlService.SandboxMode.IFRAME)
               .setWidth(400)
               .setHeight(80);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Last Sync Details');
}

/**
 * Function that fires when Refresh Current Sheet command is chosen.
 */
function refreshCurrentSheet() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Pulling new data', 'Status', -1);
  try {
    pullNewData();
  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Error: ' + err, 'Status', -1);
    return;
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Refresh Complete', 'Status', 3);
}

/**
 * Function that fires when Refresh All Sheet command is chosen.
 */
function refreshAllSheets() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Pulling new data', 'Status', -1);
  try {
    pullNewDataAll();
  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Error: ' + err, 'Status', -1);
    return;
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Refresh Complete', 'Status', 3);
}

/**
 * This is the main function which populates the data in the spreadsheet.
 * @param {Object} sheet The sheet object to populate the data in.
 */
function pullNewData(sheet) {  
  var documentProperties = PropertiesService.getDocumentProperties();
  var currentSheetId;
  var SpreadsheetName = SpreadsheetApp.getActiveSpreadsheet().getName();
  var SpreadsheetURL = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  if (sheet == null) {
    currentSheetId = SpreadsheetApp.getActiveSheet().getSheetId();
    sheet = SpreadsheetApp.getActiveSheet();
  } else
    currentSheetId = sheet.getSheetId();

  var networkId =
      documentProperties.getProperty(currentSheetId + SUFFIX_NETWORK_ID);
  var profileId =
      documentProperties.getProperty(currentSheetId + SUFFIX_PROFILE_ID);
  var reportId =
      documentProperties.getProperty(currentSheetId + SUFFIX_REPORT_ID);
  var reportName =
      documentProperties.getProperty(currentSheetId + SUFFIX_REPORT_NAME);
  
  if (reportId && profileId) {
    getReport(profileId,
              reportId,
              reportName,
              currentSheetId,
              networkId);
  }
}

/**
 * Refresh all linked sheets in the current spreadsheet.
 */
function pullNewDataAll() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    pullNewData(sheets[i]);
  }
}

/**
 * Function called from DialogJavaScript.html to show the list of reports for a
 * profile id.
 * @param {string} profileId The user profile to fetch existing reports for.
 * @return {Array<string>} The list of reports.
 */
function getReportList(profileId) {
    var values = {
    'sortField': 'ID',
    'sortOrder': "DESCENDING"
    }
  var reportList = [];
  var result = DoubleClickCampaigns.Reports.list(profileId,values);
  if (!result) {
    throw new Error('Unable to fetch reports from CM');
  }
  var reports = result.items;
  var reportList = [];
  var nextToken = result.nextPageToken;
  for (var i = 0; i < reports.length; i++) {
    if (reports[i].format == 'CSV') {
      reportList.push({'id' : reports[i].id, 'name' : reports[i].name});
    }
  }
  if(nextToken){
    getNextPage(profileId, nextToken, reportList)
  }
  console.log("Returning only reportList");
  return [reportList,nextToken];
  
}

/**
 * Helper function for getReportList function to fetch all report pages, now that the API returns only 10 results per page. 
 * Recursive function that will keep looking for nextPages until there's no nextPageToken left.
 * @param {string} profileId The user profile to fetch existing reports for.
 * @param {string} nextPageToken The next page Token returned when using the getReportList function.
 * @param {string} reportList The list with all the current reports already fetched from the requests.
 * @return {Array<string>} The list of reports.
 */
function getNextPage(profileId, nextPageToken, reportList){
  var values  = {
    'sortField': 'ID',
    'sortOrder': "DESCENDING",
    'pageToken': nextPageToken
    }
  var result = DoubleClickCampaigns.Reports.list(profileId,values);
    if (!result) {
    throw new Error('Unable to fetch reports from CM');
  }
  var reports = result.items;
  var nextToken = result.nextPageToken
  for (var i = 0; i < reports.length; i++) {
    if (reports[i].format == 'CSV') {
      reportList.push({'id' : reports[i].id, 'name' : reports[i].name});
    }
  }
  if(nextToken && reportList.length <= 40){
    return getNextPage(profileId,nextToken,reportList)
  } 
  return [reportList,nextToken]
}

/**
 * Check each worksheet and see if it is linked to a DCM report.
 * If yes, pull the latest data and sync.
 */
function DCM_offlineReportSync() {
  // Check if the script has all the authorization need to run offline sync
  var props = PropertiesService.getDocumentProperties();
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  var SpreadsheetName = SpreadsheetApp.getActiveSpreadsheet().getName();
  var SpreadsheetURL = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var documentProperties = PropertiesService.getDocumentProperties();

  // Check if the actions of the trigger requires authorization that has not
  // been granted yet; if so, warn the user via email. This check is required
  // when using triggers with add-ons to maintain functional triggers.
  if (authInfo.getAuthorizationStatus() ==
      ScriptApp.AuthorizationStatus.REQUIRED) {
    console.warn('CM Offline Report Sync: Auth Required for: ' +
                 SpreadsheetName + '. Url: ' + SpreadsheetURL);
    // Re-authorization is required. In this case, the user needs to be alerted
    // that they need to re-authorize; the normal trigger action is not
    // conducted, since it requires authorization first. Send at most one
    // "Authorization Required" email per day to avoid spamming users.
    var lastAuthEmailDate = props.getProperty('lastAuthEmailDate');
    var today = new Date().toDateString();
    if (lastAuthEmailDate != today) {
      if (MailApp.getRemainingDailyQuota() > 0) {
        var html = HtmlService.createTemplateFromFile('AuthorizationEmail');
        html.url = authInfo.getAuthorizationUrl();
        html.addonTitle = ADDON_TITLE;
        var message = html.evaluate();

        MailApp.sendEmail(
            SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail(),
            ADDON_TITLE + ' - Authorization Required',
            message.getContent(),
            {name : ADDON_TITLE, htmlBody : message.getContent()});
      }
      props.setProperty('lastAuthEmailDate', today);
    }
  } else {
    console.info('Started CM offline sync for: ' + SpreadsheetName +
                 '. Url: ' + SpreadsheetURL);
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var currentSheetId = sheets[i].getSheetId();
      var profileId =
          documentProperties.getProperty(currentSheetId + SUFFIX_PROFILE_ID);
      var reportId =
          documentProperties.getProperty(currentSheetId + SUFFIX_REPORT_ID);

      // Skip offline sync for this sheet, if it is not linked to DCM report
      if (!profileId || !reportId) {
        continue;
      }

      //try {
        console.info('Started CM offline sync for: ' + SpreadsheetName +
                     '. Sheet Name: ' + sheets[i].getName() +
                     '. Url: ' + SpreadsheetURL);
        pullNewData(sheets[i]);
        console.info('Finished CM offline sync for: ' + SpreadsheetName +
                     '. Sheet Name: ' + sheets[i].getName() +
                     '. Url: ' + SpreadsheetURL);
      /*} catch (e) {
        MailApp.sendEmail(
            SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail(),
            ADDON_TITLE + ' - Offline Sync Failed',
            'Sheet URL: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl() +
                '#gid=' + currentSheetId + '<br><br> Error: ' + e);
        console.error(
            ADDON_TITLE + ' - Offline Sync Failed',
            'Sheet URL: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl() +
                '#gid=' + currentSheetId + '<br><br> Error: ' + e);
      }*/
    }
  }
}

/**
 * Pull the report that has just been setup. This function is called from the
 * dialog box.
 * @param {string} profileId The current user's profile id.
 * @param {string} reportId The id of the report to fetch.
 * @param {string} reportName The name of the report to fetch.
 * @param {number} currentSheetId The id of the sheet to populate.
 * @param {string} networkId The cm network id.
 * @return {string} Human readable summary of the linked report.
 */
function pullReport(
    profileId, reportId, reportName, currentSheetId, networkId) {
  try {
    var strReadableLinkedReport = getReport(profileId,
                                            reportId,
                                            reportName,
                                            currentSheetId,
                                            networkId);
  } catch (e) {
    console.error(e);
    console.error(e.stack);
    throw (e);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(
      'CM report data added. You can now manually refresh or setup a scheduled sync',
      'Status',
      5);
  return strReadableLinkedReport;
}

/**
 * Download the latest report file corresponding to a report in DCM and put the
 * data in sheets or in drive.
 * @param {string} profileId The current user's profile id.
 * @param {string} reportId The id of the report to fetch.
 * @param {string} reportName The name of the report to fetch.
 * @param {number} currentSheetId The id of the sheet to populate.
 * @param {string} networkId The dcm network id.
 * @return {string} Human readable summary of the linked report.
 */
function getReport(
    profileId, reportId, reportName, currentSheetId, networkId) {
  console.log('started');
  var SpreadsheetName = SpreadsheetApp.getActiveSpreadsheet().getName();
  var SpreadsheetURL = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  SpreadsheetApp.getActiveSpreadsheet().toast(
      'Pulling CM report...', 'Status', -1);
  var result;
  var latestReportFile;
  var file;
  var OAuthToken;
  if (currentSheetId == null) {
    currentSheetId = SpreadsheetApp.getActiveSheet().getSheetId();
  }

  var sheet = getSheetById(currentSheetId);
  console.log('Started pulling report');
  
  // Get all reports under this user profile
  result = DoubleClickCampaigns.Reports.Files.list(
      profileId, reportId, {'sortField' : 'LAST_MODIFIED_TIME'});
  if (!result) {
    throw new Error(
        'Cannot fetch report files at this moment. Please check if there any CM report files');
  }
  if (result.items) {
    latestReportFile = result.items[0];
    if (!latestReportFile || latestReportFile.status == 'PROCESSING') {
      latestReportFile = result.items[1];
      if (!latestReportFile) {
        throw new Error('No files found corresponding to the report: ' +
                        reportName);
      }
    }
    console.info('Found CM Report for ' + SpreadsheetName +
                 '. Sheet Name: ' + sheet.getName() +
                 '. Url: ' + SpreadsheetURL);
  }

  file = DoubleClickCampaigns.Reports.Files.get(
      profileId, reportId, latestReportFile.id);  
  
  var oauthService = checkAndGetOAuthService();
  if (!oauthService) {
    return;
  }
  OAuthToken = oauthService.getAccessToken();
  
  // if the current file is still being process, pick the last file
  if (file.urls) {
    var httpOptions = {'headers' : {'Authorization' : 'Bearer ' + OAuthToken}};
    var contents = UrlFetchApp.fetch(file.urls.apiUrl, httpOptions);
    if (file.format == 'CSV') {
      console.info('Pulling CM Report File for ' + SpreadsheetName +
                   '. Sheet Name: ' + sheet.getName() +
                   '. Url: ' + SpreadsheetURL);
      populateSpreadsheet(
          contents.getContentText(), file.fileName, currentSheetId);
      console.info('Finished Pulling CM Report File for ' + SpreadsheetName +
                   '. Sheet Name: ' + sheet.getName() +
                   '. Url: ' + SpreadsheetURL);
    } else { // Store the Excel file directly in Drive
      DocsList.createFile(contents.getBlob()).rename(file.fileName);
    }
  }

  // store the profileId and reportId in Document Properties for later use in
  // offline sync
  var documentProperties = PropertiesService.getDocumentProperties();

  // Delete all the properties corresponding to other DDM add-ons
  documentProperties.deleteProperty(currentSheetId + SUFFIX_BUCKET_NAME);
  documentProperties.deleteProperty(currentSheetId +
                                    SUFFIX_DBM_REPORT_UPDATED_DATE);
  documentProperties.deleteProperty(currentSheetId + SUFFIX_WEBQUERY_URL);

  // store the Profile ID and Report ID in sheet specific document property
  // (with Sheet Id Prefix)
  // If not already stored
  // This allow each sheet to be independently synced to a DCM report
  documentProperties.setProperty(currentSheetId + SUFFIX_PROFILE_ID, profileId);
  documentProperties.setProperty(currentSheetId + SUFFIX_REPORT_ID, reportId);
  documentProperties.setProperty(currentSheetId + SUFFIX_REPORT_NAME,
                                 reportName);
  documentProperties.setProperty(currentSheetId + SUFFIX_DCM_REPORT_SETUP_USER,
                                 Session.getActiveUser().getEmail());
  if (networkId) {
    documentProperties.setProperty(currentSheetId + SUFFIX_NETWORK_ID,
                                   networkId);
  }

  var end_time = new Date();
  documentProperties.setProperty(currentSheetId + SUFFIX_LAST_SYNC, end_time);
  return retrieveLinkedReport();
}

/**
 * Fetch the human readable summary of the linked report.
 * @return {string} Human readable summary of the linked report.
 */
function retrieveLinkedReport() {
  var documentProperties = PropertiesService.getDocumentProperties();
  var currentSheetId = SpreadsheetApp.getActiveSheet().getSheetId();
  var networkId =
      documentProperties.getProperty(currentSheetId + SUFFIX_NETWORK_ID);
  var reportName =
      documentProperties.getProperty(currentSheetId + SUFFIX_REPORT_NAME);
  var reportSetupUser = documentProperties.getProperty(
      currentSheetId + SUFFIX_DCM_REPORT_SETUP_USER);

  var strReadableLinkedReport = '';
  if (networkId) {
    strReadableLinkedReport +=
        'Currently Linked Network ID: <b>' + networkId + '</b>';
  }
  if (reportName) {
    strReadableLinkedReport +=
        '.&nbsp;Currently Linked Report Name: <b>' + reportName + '</b>';
  }
  if (reportSetupUser) {
    strReadableLinkedReport +=
        '.&nbsp;Setup by: <b>' + reportSetupUser + '</b>';
  }

  return strReadableLinkedReport;
}

/**
 * Remove the linking of current report.
 * @return {string} Human readable form of the linked report.
 */
function unlinkReport() {
  var documentProperties = PropertiesService.getDocumentProperties();
  var currentSheetId = SpreadsheetApp.getActiveSheet().getSheetId();
  documentProperties.deleteProperty(currentSheetId + SUFFIX_NETWORK_ID);
  documentProperties.deleteProperty(currentSheetId + SUFFIX_REPORT_NAME);
  documentProperties.deleteProperty(currentSheetId +
                                    SUFFIX_DCM_REPORT_SETUP_USER);
  documentProperties.deleteProperty(currentSheetId + SUFFIX_REPORT_ID);
  documentProperties.deleteProperty(currentSheetId + SUFFIX_PROFILE_ID);
  return retrieveLinkedReport();
}

/**
 * Unlink all reports in the current spreadsheet.
 */
function unlinkAllReports() {
  var documentProperties = PropertiesService.getDocumentProperties();
  var allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < allSheets.length; i++) {
    var currentSheetId = allSheets[i].getSheetId();
    documentProperties.deleteProperty(currentSheetId + SUFFIX_NETWORK_ID);
    documentProperties.deleteProperty(currentSheetId + SUFFIX_REPORT_NAME);
    documentProperties.deleteProperty(currentSheetId +
                                      SUFFIX_DCM_REPORT_SETUP_USER);
    documentProperties.deleteProperty(currentSheetId + SUFFIX_REPORT_ID);
    documentProperties.deleteProperty(currentSheetId + SUFFIX_PROFILE_ID);
  }
}

/**
 * Helper Function to read content stream and put it in the current active
 * spreadsheet.
 * @param {Object} csvContent The CSV content of the fetched report.
 * @param {string} fileName The name of the file from DCM.
 * @param {number} currentSheetId The id of the current sheet.
 */
function populateSpreadsheet(csvContent, fileName, currentSheetId) {
  var rows = Utilities.parseCsv(csvContent);
  var documentProperties = PropertiesService.getDocumentProperties();

  if (rows && rows.length && rows[0] && rows[0].length) {
    // Get the existing spreadsheet using the specified filename
    var sheet = getSheetById(currentSheetId);

    // Clean up report metadata at the top of the report
    var numRowsToRemove = 0;
    for (var i = 0; i < rows.length; i++) {
      numRowsToRemove++;
      if (rows[i][0] == 'Report Fields') {
        break;
      }
    }

    if (numRowsToRemove > 0) {
      rows.splice(0, numRowsToRemove);
    }

    // Clean up the grand total at the bottom row
    rows.pop();

    // Clear the existing rows values. Only values are deleted, formatting is
    // retained.
    sheet.clearContents();

    // Output the result rows.
    // The values will use the formatting set for existing cells. Manually set
    // them if need be.
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

    // Truncate the sheet to number of rows with content.
    if (sheet.getMaxRows() > sheet.getLastRow()) {
      sheet.deleteRows(sheet.getLastRow() + 1,
                       sheet.getMaxRows() - sheet.getLastRow());
    }

    // Truncate the sheet to number of columns with content.
    if (sheet.getMaxColumns() > sheet.getLastColumn()) {
      sheet.deleteColumns(sheet.getLastColumn() + 1,
                          sheet.getMaxColumns() - sheet.getLastColumn());
    }

    var end_time = new Date();
    documentProperties.setProperty(currentSheetId + SUFFIX_LAST_SYNC, end_time);
  }
}

/**
 * Function called from DialogJavaScript.html to log certain activities.
 * @param {string} logText The text to log to browser console.
 */
function doLog(logText) {
  console.log(logText);
}

/**
 * Fetch the sheet object given the sheet id.
 * @param {number} sheetId The id of the sheet.
 * @return {Object} The sheet object.
 */
function getSheetById(sheetId) {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  // Iterate all sheets and compare ids.
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() == sheetId) {
      return sheets[i];
    }
  }
}

/**
 * Function that creates OAuth2 Object using OAuth2 Library for Appscript.
 * https://github.com/googlesamples/apps-script-oauth2.
 * @return {Object} The OAuth service object to use for the current user.
 */
function getOAuthService() {
  return OAuth2
      .createService(SERVICENAME)

      // Set the endpoint URLs, which are the same for all Google services.
      .setAuthorizationBaseUrl(AUTHORIZATION_URL)
      .setTokenUrl(TOKEN_ACCESS_URL)

      // Set the client ID and secret, from the Google Developers Console.
      .setClientId(CLIENT_ID)
      .setClientSecret(CLIENT_SECRET)

      // Set the project key of the script using this library.
      //.setProjectKey(ScriptApp.getProjectKey())

      // Set the name of the callback function in the script referenced
      // above that should be invoked to complete the OAuth flow.
      .setCallbackFunction('authCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())

      // Set the scopes to request (space-separated for Google services).
      .setScope(SCOPE)

      // Below are Google-specific OAuth2 parameters.

      // Sets the login hint, which will prevent the account chooser screen
      // from being shown to users logged in with multiple accounts.
      .setParam('login_hint', Session.getActiveUser().getEmail())

      // Requests offline access.
      .setParam('access_type', 'offline')

      // Forces the approval prompt every time. This is useful for testing,
      // but not desirable in a production application.
      .setParam('approval_prompt', 'force');
}

/**
 * Clear oAuth Token, if a user manually revokes access to this app under Google
 * Security, then we need to clear the OAuth Token running this function.
 */
function clearToken() {
  var OAuthService = getOAuthService();
  OAuthService.reset();
}

/**
 * Function is called after authentication request and displays relevant info
 * depending on whether the authentication is successful or not.
 * @param {Object} request The request object from OAuth service.
 * @return {Object} The html object to display.
 */
function authCallback(request) {
  var OAuthService = getOAuthService();
  var isAuthorized = OAuthService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput(
        '<p style="font-family:Arial">Success! You can close this tab and open the Dialog box again.</p>');
  } else {
    return HtmlService.createHtmlOutput(
        '<p style="font-family:Arial">Denied. You can close this tab. Please try opening the Dialog box again</p>');
  }
}

/**
 * Shows a dialog to schedule the report sync.
 */
function schedulerDialog() {
  var documentProperties = PropertiesService.getDocumentProperties();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var reportSetupUser = [];
  for (var i = 0; i < sheets.length; i++) {
    reportSetupUser.push(documentProperties.getProperty(
        sheets[i].getSheetId() + '_REPORT_SETUP_USER_EMAIL'));
  }

  
  var ui = HtmlService.createTemplateFromFile('Scheduler')
      .evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(700)
      .setHeight(100);

  SpreadsheetApp.getUi().showModalDialog(ui, 'Schedule Reports');
}

/**
 * Setup a Apps Script trigger to schedule the report sync.
 * @param {boolean} enableScheduler Whether scheduler is enabled or not.
 * @param {string} frequency The frequency to fetch the report.
 * @param {string} timer The scheduled time to fetch the report.
 * @param {string} timeOfDaySelected The scheduled time to fetch the report
 *     for weekly reports.
 * @return {string} Readable format of the schedule.
 */
function setScheduler(enableScheduler, frequency, timer, timeOfDaySelected) {
  var documentProperties = PropertiesService.getDocumentProperties();
  var triggerSettings = {};

  // Delete the trigger previously setup.
  var triggers = ScriptApp.getProjectTriggers();

  for (i = 0; i < triggers.length; i++) {
    // Delete the old trigger on the document.
    if (triggers[i].getUniqueId() ==
        documentProperties.getProperty(PROP_DCM_TRIGGER_ID)) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  if (!enableScheduler) {
    documentProperties.deleteProperty(PROP_DCM_SCHEDULE_FREQUENCY);
    documentProperties.deleteProperty(PROP_DCM_SCHEDULE_TIME);
    documentProperties.deleteProperty(PROP_DCM_SCHEDULE_TIME2);
    documentProperties.deleteProperty(PROP_DCM_TRIGGER_CREATED_BY);
    documentProperties.deleteProperty(PROP_DCM_TRIGGER_ID);
    return '';
  }

  switch (frequency) {
  case 'hourly':
    var triggerID = ScriptApp.newTrigger(TRIGGER_DCM_OFFLINE)
                        .timeBased()
                        .everyHours(timer)
                        .create()
                        .getUniqueId();

    triggerSettings[PROP_DCM_SCHEDULE_FREQUENCY] = frequency;
    triggerSettings[PROP_DCM_SCHEDULE_TIME] = timer;
    triggerSettings[PROP_DCM_SCHEDULE_TIME2] = null;
    triggerSettings[PROP_DCM_TRIGGER_CREATED_BY] = Session.getActiveUser().getEmail();
    triggerSettings[PROP_DCM_TRIGGER_ID] = triggerID;
    
    
    documentProperties.setProperties(triggerSettings);
    break;
  case 'daily':
    var triggerID = ScriptApp.newTrigger(TRIGGER_DCM_OFFLINE)
                        .timeBased()
                        .everyDays(1)
                        .atHour(timer)
                        .create()
                        .getUniqueId();

    triggerSettings[PROP_DCM_SCHEDULE_FREQUENCY] = frequency;
    triggerSettings[PROP_DCM_SCHEDULE_TIME] = timer;
    triggerSettings[PROP_DCM_SCHEDULE_TIME2] = null;
    triggerSettings[PROP_DCM_TRIGGER_CREATED_BY] = Session.getActiveUser().getEmail();
    triggerSettings[PROP_DCM_TRIGGER_ID] = triggerID;
    
    documentProperties.setProperties(triggerSettings);
    break;
  case 'weekly':
    var onWeekDay = ScriptApp.WeekDay.MONDAY;
    switch (timer) {
    case 'MONDAY':
      onWeekDay = ScriptApp.WeekDay.MONDAY;
      break;

    case 'TUESDAY':
      onWeekDay = ScriptApp.WeekDay.TUESDAY;
      break;

    case 'WEDNESDAY':
      onWeekDay = ScriptApp.WeekDay.WEDNESDAY;
      break;

    case 'THURSDAY':
      onWeekDay = ScriptApp.WeekDay.THURSDAY;
      break;

    case 'FRIDAY':
      onWeekDay = ScriptApp.WeekDay.FRIDAY;
      break;

    case 'SATURDAY':
      onWeekDay = ScriptApp.WeekDay.SATURDAY;
      break;

    case 'SUNDAY':
      onWeekDay = ScriptApp.WeekDay.SUNDAY;
      break;
    }
    var triggerID = ScriptApp.newTrigger(TRIGGER_DCM_OFFLINE)
                        .timeBased()
                        .everyWeeks(1)
                        .onWeekDay(onWeekDay)
                        .atHour(timeOfDaySelected)
                        .create()
                        .getUniqueId();

    triggerSettings[PROP_DCM_SCHEDULE_FREQUENCY] = frequency;
    triggerSettings[PROP_DCM_SCHEDULE_TIME] = timer;
    triggerSettings[PROP_DCM_SCHEDULE_TIME2] = timeOfDaySelected;
    triggerSettings[PROP_DCM_TRIGGER_CREATED_BY] = Session.getActiveUser().getEmail();
    triggerSettings[PROP_DCM_TRIGGER_ID] = triggerID;
    
    documentProperties.setProperties(triggerSettings);
    break;
  }

  return retrieveCurrentSchedule();
}

/**
 * Retrieve the current sync schedule in human readable format.
 * @return {string}
 */
function retrieveCurrentSchedule() {
  // Only returns the triggers for the current user for this document.
  var triggers =
      ScriptApp.getUserTriggers(SpreadsheetApp.getActiveSpreadsheet());
  var documentProperties = PropertiesService.getDocumentProperties();
  var doesTriggerExist = false;

  // If the user who setup the Scheduler is same as the current user, check if
  // the trigger stil exists. If it doesn't exist, delete the properties set.
  if (documentProperties.getProperty(PROP_DCM_TRIGGER_CREATED_BY) ==
      Session.getActiveUser().getEmail()) {
    for (i = 0; i < triggers.length; i++) {
      if (triggers[i].getUniqueId() ==
          documentProperties.getProperty(PROP_DCM_TRIGGER_ID)) {
        doesTriggerExist = true;
      }
    }

    if (!doesTriggerExist) {
      documentProperties.deleteProperty(PROP_DCM_SCHEDULE_FREQUENCY);
      documentProperties.deleteProperty(PROP_DCM_SCHEDULE_TIME);
      documentProperties.deleteProperty(PROP_DCM_SCHEDULE_TIME2);
      documentProperties.deleteProperty(PROP_DCM_TRIGGER_CREATED_BY);
      documentProperties.deleteProperty(PROP_DCM_TRIGGER_ID);
      return '';
    }
  }

  var documentProperties = PropertiesService.getDocumentProperties();
  var currentFrequency =
      documentProperties.getProperty(PROP_DCM_SCHEDULE_FREQUENCY);
  var currentTimer = documentProperties.getProperty(PROP_DCM_SCHEDULE_TIME);
  var currentTimer2 = documentProperties.getProperty(PROP_DCM_SCHEDULE_TIME2);
  var createdBy = documentProperties.getProperty(PROP_DCM_TRIGGER_CREATED_BY);
  var strReadableSchedule = '';

  switch (currentFrequency) {
  case 'hourly':
    strReadableSchedule +=
        '<b>Current Sync Schedule:</b> Every ' + currentTimer + ' Hours';
    break;

  case 'daily':
    strReadableSchedule += '<b>Current Sync Schedule:</b> Daily between ' +
                           DAILY_FREQUENCY[currentTimer].text;
    break;

  case 'weekly':
    strReadableSchedule += '<b>Current Sync Schedule:</b> Weekly on ' +
                           currentTimer + ' between ' +
                           DAILY_FREQUENCY[currentTimer2].text;
    break;
  }
  if (createdBy && strReadableSchedule) {
    strReadableSchedule += '. Created by: <b>' + createdBy + '</b>';
  }
  return strReadableSchedule;
}