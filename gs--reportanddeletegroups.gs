/**
 * Script to manage Google Workspace groups directly from Google Sheets, including listing, deleting empty groups, and sending reports.
 * Features:
 * - Lists empty groups (0 members and owners) in a Google Sheet.
 * - Provides an option to delete these groups after listing.
 * - Sends an email report with the count of empty groups and the primary domain name.
 * - Adds custom menu options for easy script execution.
 * Requirements:
 * - Admin SDK enabled in Google Cloud Project.
 * - An destination email address for reporting the amount of empty groups
 * - Proper scopes for Admin Directory API in appsscript.json.
 * Written by Jonas Lund, 2024.
 */

/**
 * Called automatically when the Google Sheet is opened. Adds a custom menu to the Google Sheets UI.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Group Management')
    .addItem('Start Listing Empty Groups', 'listGroupsWithZeroMembersAndOwners')
    .addItem('Reset and Start Over', 'resetAndStartOver')
    .addItem('Delete Found Empty Groups', 'promptToDeleteGroups')
    .addToUi();
}

/**
 * Prepares the Google Sheet by setting up headers. Clears existing content to start fresh.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The sheet ready for listing groups.
 */
function setupSheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName("Groups with 0 Members and Owners");
  if (!sheet) {
    sheet = spreadsheet.insertSheet("Groups with 0 Members and Owners");
  }
  
  sheet.clear(); // Clears the sheet to start fresh on each run
  var headers = [["Group Name", "Members", "Owners", "Creation Date", "Email Address"]];
  sheet.getRange('A1:E1').setValues(headers).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  return sheet;
}

/**
 * Lists Google Workspace groups with 0 members and owners. Sends an email report if empty groups are found.
 */
function listGroupsWithZeroMembersAndOwners() {
  var sheet = setupSheet();
  var emptyGroupsFound = 0;
  var options = { customer: 'my_customer', maxResults: 200 }; // Replace 'my_customer' with your actual customer ID

  do {
    var response = AdminDirectory.Groups.list(options);
    var groups = response.groups || [];
    options.pageToken = response.nextPageToken;

    groups.forEach(function(group) {
      var members = AdminDirectory.Members.list(group.id).members || [];
      var owners = AdminDirectory.Members.list(group.id, {'roles': 'OWNER'}).members || [];
      if (members.length === 0 && owners.length === 0) {
        emptyGroupsFound++;
        var createdDate = new Date(group.creationTime);
        var formattedDate = Utilities.formatDate(createdDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        sheet.appendRow([group.name, 0, 0, formattedDate, group.email]);
      }
    });
  } while (options.pageToken);

  if (emptyGroupsFound > 0) {
    var customer = AdminDirectory.Customers.get('my_customer'); // Fetch customer information to get primary domain
    sendEmailReport(emptyGroupsFound, customer.customerDomain);
  }
}

/**
 * Sends an email report with the count of found empty groups and the primary domain name.
 * @param {number} emptyGroupsFound - The number of empty groups found.
 * @param {string} primaryDomain - The primary domain name of the Google Workspace account.
 */
function sendEmailReport(emptyGroupsFound, primaryDomain) {
  MailApp.sendEmail({
    to: "groupreport@yourdomain.com", // Replace with your own desired email for the reports
    subject: "Empty Google Workspace Groups Report",
    body: `Found ${emptyGroupsFound} empty groups in the primary domain: ${primaryDomain}.`
  });
}

/**
 * Prompts the user with a dialog to confirm before proceeding with group deletion.
 */
function promptToDeleteGroups() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Delete Empty Groups', 
                           'Do you want to delete the found groups with 0 members?', 
                           ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    deleteGroups();
  } else {
    Logger.log('Group deletion cancelled.');
  }
}

/**
 * Deletes groups listed in the sheet and updates the sheet with deletion status and time.
 */
function deleteGroups() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Groups with 0 Members and Owners");
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var groupEmail = data[i][4];
    try {
      AdminDirectory.Groups.remove(groupEmail);
      var deletionTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      sheet.getRange(i + 1, 6).setValue(`Deleted at ${deletionTime}`);
    } catch (e) {
      Logger.log(`Failed to delete group: ${groupEmail}; Error: ${e.message}`);
      sheet.getRange(i + 1, 6).setValue('Failed to delete');
    }
  }
}

/**
 * Resets the script by clearing the sheet, any saved state, and restarts the group listing process.
 */
function resetAndStartOver() {
  clearTriggersAndPageToken();
  setupSheet(); // Reinitialize the sheet with headers
  listGroupsWithZeroMembersAndOwners(); // Restart the listing process
}

/**
 * Clears any saved state (like page tokens) and Google Apps Script triggers associated with this script.
 */
function clearTriggersAndPageToken() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('pageToken');
  
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
}
