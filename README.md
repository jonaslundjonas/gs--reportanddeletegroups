# reportanddeletegroups
Takes out a list of empty groups to a Google Sheet, report it to an email and offers the possebility to delete them.


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
