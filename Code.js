/**
 * ============================================================
 * QUEUE RESOLVE  (skills test — sample module, not production)
 * ------------------------------------------------------------
 * A small approval queue lives on a sheet called "Review_Queue".
 * A reviewer sets a Decision on each row (Approve / Reject).
 * resolveQueue() walks the queue and applies the approved rows:
 * it copies the approved value into the live "Accounts" sheet,
 * then stamps the queue row as Done so it is not processed again.
 *
 * Columns on Review_Queue:
 *   A Row_ID   B Account_ID   C Field   D New_Value   E Decision   F Status   G Resolved_At
 *
 * Columns on Accounts:
 *   A Account_ID   B Name   C Website   D Status
 *
 * A reviewer reports: "I approved a batch last week, the rows show
 * Done, but some of the accounts never actually got the new value."
 * ============================================================
 */

var QUEUE_SHEET = "Review_Queue";
var ACCOUNTS_SHEET = "Accounts";

function resolveQueue() {
  var ss = SpreadsheetApp.getActive();
  var q = ss.getSheetByName(QUEUE_SHEET);
  var accounts = ss.getSheetByName(ACCOUNTS_SHEET);

  var rows = q.getRange(2, 1, q.getLastRow() - 1, 7).getValues();
  var accData = accounts.getRange(2, 1, accounts.getLastRow() - 1, 4).getValues();

  var applied = 0;

  for (var i = 0; i < rows.length; i++) {
    var decision = rows[i][4];
    var status = rows[i][5];

    if (status === "Done") continue;
    if (decision !== "Approve") continue;

    var accountId = rows[i][1];
    var field = rows[i][2];
    var newValue = rows[i][3];

    // stamp the queue row as done
    q.getRange(i + 2, 6).setValue("Done");
    q.getRange(i + 2, 7).setValue(new Date());

    // find the account and write the new value
    for (var j = 0; j < accData.length; j++) {
      if (accData[j][0] === accountId) {
        if (field === "Name")    accData[j][1] = newValue;
        if (field === "Website") accData[j][2] = newValue;
        if (field === "Status")  accData[j][3] = newValue;
        accounts.getRange(j + 2, 1, 1, 4).setValues([accData[j]]);
        applied++;
      }
    }
  }

  Logger.log("Applied " + applied + " approved changes.");
}

/**
 * Helper: adds a new account to the Accounts sheet.
 * Used by other parts of the system. Leave the signature alone.
 */
function addAccount(id, name, website) {
  var accounts = SpreadsheetApp.getActive().getSheetByName(ACCOUNTS_SHEET);
  accounts.appendRow([id, name, website, "Active"]);
}

/**
 * Helper: adds a new account to the Accounts sheet (v2).
 * Same as above but also sets status. Some callers use this one.
 */
function addAccountWithStatus(id, name, website, status) {
  var accounts = SpreadsheetApp.getActive().getSheetByName(ACCOUNTS_SHEET);
  accounts.appendRow([id, name, website, status]);
}
