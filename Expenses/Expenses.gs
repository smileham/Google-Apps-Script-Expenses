/*
* Need to define the following Project Properties;
*
*  spreadsheet = 1P5FjdHBqDgGOqbylJWOK-iM7KDFrYMx8G4rK9gCXWmE (ID for the spreadsheet to store expenses)
*  referenceNumber = 0 (starting unique ID for expenses)
*  emailDomain = @gmail.com (ensure that the "Line Manager" email address in within your domain)
*  financeEmail = steven.mileham@gmail.com (User to share receipt images with)
*/
function doGet(e) {
  if (e.parameter.outcome!=null && e.parameter.outcome!="") {
    // Manage Approval
    return manageApproval_(e);
  }
  else {
    // Launch Expenses App
    return HtmlService.createTemplateFromFile("Home").evaluate()
        .addMetaTag("viewport", "width=device-width, initial-scale=1")
        .setTitle("Expenses")
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  }

}

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
}


function getData() {
    var theSpreadsheet = SpreadsheetApp.openById(ScriptProperties.getProperty("spreadsheet"));

    var staffDetails = { staffEmail: Session.getActiveUser().getEmail(),
                        staffNumber: "",
                        staffCostCentre: "",
                        staffCostLine: "",
                        staffLineManagerEmail: PropertiesService.getScriptProperties().getProperty("emailDomain") };

    staffDetails = lookupStaffDetails(staffDetails);
    staffDetails["costCentreListBox"] = buildSpreadsheetListBox_("CostCentre", staffDetails.staffCostCentre, theSpreadsheet);
    staffDetails["costLineListBox"] = buildSpreadsheetListBox_("CostLine", staffDetails.staffCostLine, theSpreadsheet);
    staffDetails["domain"] = PropertiesService.getScriptProperties().getProperty("emailDomain");
    staffDetails["codeString"] = buildCodeList_(theSpreadsheet);


    return staffDetails;
}

function buildSpreadsheetListBox_(theRangeName, selected, theSpreadsheet) {
    var theRange = theSpreadsheet.getRangeByName(theRangeName);

    var theValues = theRange.getValues();
    var newValues = new Array();

    for (var i = 0; i < theValues.length; i++) {
        if (theValues[i] == "") { break; }
        newValues.push(theValues[i]);
    }

    return newValues;
}

function submitExpenses(expenses) {
    /* Load Spreadsheet to track expenses */
    var theSpreadsheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty("spreadsheet"));
    var theSheet = theSpreadsheet.getSheetByName("Expenses");
    var lastRow = theSheet.getLastRow();
    var theEmail = Session.getActiveUser().getEmail();
    var theDate = new Date().toDateString();
    var today = new Date();

    Logger.log(expenses);

    storeStaffDetails_(theEmail, expenses.staffNo, expenses.CostCentre, expenses.CostLine, expenses.lineManagerEmail);

    //var referenceNumber = hash_(Session.getActiveUser()+new Date().getTime());
    var referenceNumber = generateRef_();
    var receiptURL = createReceiptFolder_(referenceNumber, theEmail + "," + PropertiesService.getScriptProperties().getProperty("financeEmail") + "," + expenses.lineManagerEmail);

    var numRows = Number(expenses.numRows);

    var userTable = '<table style="border-spacing: 5px; width:100%">' +
        '<tr><th style="background-color:#C0C0C0">User</th><td style="background-color:#f6f6f6;">' + theEmail + '</td></tr>' +
        '<tr><th style="background-color:#C0C0C0">Staff number</th><td>' + expenses.staffNo + '</td></tr>' +
        '<tr><th style="background-color:#C0C0C0">Cost Centre</th><td style="background-color:#f6f6f6;">' + expenses.CostCentre + '</td></tr>' +
        '<tr><th style="background-color:#C0C0C0">Cost Line</th><td>' + expenses.CostLine + '</td></tr>' +
        '<tr><th style="background-color:#C0C0C0">Line Manager</th><td style="background-color:#f6f6f6;">' + expenses.lineManagerEmail + '</td></tr>' +
        '</table><br>'

    var expenseTable = '<table style="border-spacing: 5px 0; width:100%"><tr><th style="background-color:#C0C0C0">#</th><th style="background-color:#C0C0C0">Description</th><th style="background-color:#C0C0C0">Date</th>' +
        '<th style="background-color:#C0C0C0">Code</th><th style="background-color:#C0C0C0">Amount</th></tr>';
    var total = 0;

    for (var i = 1; i < numRows + 1; i++) {
        var totalCode = "'00" + stripCode_(expenses.CostCentre) + stripCode_(expenses.CostLine) + stripCode_(expenses["type" + i]);
        theDate = new Date(expenses["date" + i]).toDateString();
        var theDescription = expenses["desc" + i] == "Description" ? "" : expenses["desc" + i];
        var zebra = (i % 2 != 0 ? 'style="background-color:#f6f6f6;"' : '');
        expenseTable += '<tr><td valign="top" ' + zebra + '">' + i + '</td>' +
            '<td valign="top" ' + zebra + '>' + theDescription + '</td>' +
            '<td valign="top" ' + zebra + '">' + theDate + '</td>' +
            '<td valign="top" ' + zebra + '>' + expenses["type" + i] + '</td>' +
            '<td valign="top" ' + zebra + '>£' + Number(expenses["amount" + i]).toFixed(2) + '</td></tr>';
        total += Number(expenses["amount" + i]);

        theSheet.appendRow([theEmail, today, "'" + expenses.staffNo, expenses.CostCentre, expenses.CostLine, theDate, theDescription, expenses["type" + i], totalCode, Number(expenses["amount" + i]).toFixed(2), "=HYPERLINK(\"" + receiptURL + "\",\"" + referenceNumber + "\")", expenses.lineManagerEmail]);

    }

    expenseTable += "<tr><th colspan='4'>Total</th><th>£" + Number(total).toFixed(2) + "</th></tr></table>";

    var receiptLink = '<div style="height:30px; text-align:center; width:100%; padding: 10px 0; border-radius: 5px; background-color:#f6f6f6; font-weight:bold; font-size:18px"><a style="font-face:sans-serif; color:black;" href="' + receiptURL + '">Receipts for this claim.</a></div>';

    /* Submitter Email */
    MailApp.sendEmail({
        to: theEmail,
        subject: "Expenses - ref: " + referenceNumber,
        htmlBody: userTable + expenseTable + receiptLink,
        noReply: true
    });


    var approvalTable = '<table width="100%"><tr><td width="50%">' +
        '<div style="height:30px; text-align:center; width:100%; padding: 10px 0; border-radius: 5px; background-color:#94ffa6; font-weight:bold; font-size:18px"><a style="font-face:sans-serif; color:black;" href="' + ScriptApp.getService().getUrl() + '?outcome=approve&startRow=' + lastRow + '&referenceNumber=' + referenceNumber + '">Approve</a></div></td>' +
        '<td><div style="height:30px; text-align:center; width:100%; padding: 10px 0; border-radius: 5px; background-color:#ff9494; font-weight:bold; font-size:18px"><a style="font-face:sans-serif; color:black;" href="' + ScriptApp.getService().getUrl() + '?outcome=reject&startRow=' + lastRow + '&referenceNumber=' + referenceNumber + '">Reject</a></div></td>' +
        '</tr></table>';

    /* Line Manager Email */
    MailApp.sendEmail({
        to: expenses.lineManagerEmail,
        subject: "[Approval] - Expenses - ref: " + referenceNumber,
        htmlBody: userTable + expenseTable + receiptLink + approvalTable,
        noReply: true
    });

    var payload = { "referenceNumber": referenceNumber, "receiptURL": receiptURL };
    return payload;

}
