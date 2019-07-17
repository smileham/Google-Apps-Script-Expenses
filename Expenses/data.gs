function createReceiptFolderLive_(refId, shareList) {

    var theExpenseFolder = null;

    var query = 'title="Expense Receipts" and ' +
        'mimeType = "application/vnd.google-apps.folder"';
    var folders, pageToken;
    do {
        folders = Drive.Files.list({
            q: query,
            maxResults: 10,
            pageToken: pageToken
        });
        if (folders.items && folders.items.length > 0) {
            for (var i = 0; i < folders.items.length; i++) {
                theExpenseFolder = folders.items[i];
            }
        } else {
            // Folder not found
            theExpenseFolder = Drive.Files.insert({
                "title": "Expense Receipts",
                "parents": [{ "id": "root" }],
                "mimeType": "application/vnd.google-apps.folder"
            });
        }
        pageToken = folders.nextPageToken;
    } while (pageToken);

    var theFolder = null;

    var query = 'title="Receipts: ' + refId + '" and ' +
        '"' + theExpenseFolder.id + '" in parents and ' +
        'mimeType = "application/vnd.google-apps.folder"';
    Logger.log(query);
    var folders, pageToken;
    do {
        folders = Drive.Files.list({
            q: query,
            maxResults: 10,
            pageToken: pageToken
        });
        if (folders.items && folders.items.length > 0) {
            for (var i = 0; i < folders.items.length; i++) {
                theFolder = folders.items[i];
            }
        } else {
            // Folder not found
            theFolder = Drive.Files.insert({
                "title": "Receipts: " + refId,
                "parents": [{ "id": theExpenseFolder.id }],
                "mimeType": "application/vnd.google-apps.folder"
            });
        }
        pageToken = folders.nextPageToken;
    } while (pageToken);

    var shareArray = shareList.split(",");
    for (var i = 0; i < shareArray.length; i++) {
        Drive.Permissions.insert(
            {
                'role': 'writer',
                'type': 'user',
                'value': shareArray[i]
            }, theFolder.id,
            {
                'sendNotificationEmails': 'false'
            })
    }

    return theFolder.alternateLink;
}

function createReceiptFolder_(refId, shareList) {

    var theExpenseFolder = null;

    var query = 'title="Expense Receipts (DEV)" and ' +
        'mimeType = "application/vnd.google-apps.folder"';
    var folders, pageToken;
    do {
        folders = Drive.Files.list({
            q: query,
            maxResults: 10,
            pageToken: pageToken
        });
        if (folders.items && folders.items.length > 0) {
            for (var i = 0; i < folders.items.length; i++) {
                theExpenseFolder = folders.items[i];
            }
        } else {
            // Folder not found
            theExpenseFolder = Drive.Files.insert({
                "title": "Expense Receipts (DEV)",
                "parents": [{ "id": "root" }],
                "mimeType": "application/vnd.google-apps.folder"
            });
        }
        pageToken = folders.nextPageToken;
    } while (pageToken);

    var theFolder = null;

    var query = 'title="Receipts (DEV): ' + refId + '" and ' +
        '"' + theExpenseFolder.id + '" in parents and ' +
        'mimeType = "application/vnd.google-apps.folder"';
    Logger.log(query);
    var folders, pageToken;
    do {
        folders = Drive.Files.list({
            q: query,
            maxResults: 10,
            pageToken: pageToken
        });
        if (folders.items && folders.items.length > 0) {
            for (var i = 0; i < folders.items.length; i++) {
                theFolder = folders.items[i];
            }
        } else {
            // Folder not found
            theFolder = Drive.Files.insert({
                "title": "Receipts (DEV): " + refId,
                "parents": [{ "id": theExpenseFolder.id }],
                "mimeType": "application/vnd.google-apps.folder"
            });
        }
        pageToken = folders.nextPageToken;
    } while (pageToken);


    var shareArray = shareList.split(",");
    for (var i = 0; i < shareArray.length; i++) {
        Drive.Permissions.insert(
            {
                'role': 'writer',
                'type': 'user',
                'value': shareArray[i]
            }, theFolder.id,
            {
                'sendNotificationEmails': 'false'
            })
    }

    return theFolder.alternateLink;
}

function lookupStaffDetails(staffDetails) {
    var theSpreadsheetId = PropertiesService.getScriptProperties().getProperty("inspectorateSpreadsheet");
    var theSpreadsheet = SpreadsheetApp.openById(theSpreadsheetId);
    var theSheet = theSpreadsheet.getSheetByName("Inspectors");
    var theEmails = theSheet.getRange("A:A").getValues();

    for (var i = 0; i < theEmails.length; i++) {
        if (theEmails[i][0] == staffDetails.staffEmail) {
            var staffDetails = JSON.parse(theSheet.getRange("B" + (i + 1) + ":B" + (i + 1)).getValue());
            staffDetails.rowNumber = i;
            return staffDetails;
            break;
        }
        else if (theEmails[i][0] == "") {
            break;
        }
    }
    return {
        staffEmail: "",
        staffNumber: "",
        staffCostCentre: "",
        staffCostLine: "",
        staffLineManagerEmail: ""
    };
}

function storeStaffDetails_(theEmail, staffNo, costCentre, costLine, lineManagerEmail) {
    var staffDetails = {
        staffEmail: theEmail,
        staffNumber: staffNo,
        staffCostCentre: costCentre,
        staffCostLine: costLine,
        staffLineManagerEmail: lineManagerEmail
    };
    var existingStaffDetail = lookupStaffDetails(staffDetails);

    if (existingStaffDetail != undefined) {
        staffDetails.rowNumber = existingStaffDetail.rowNumber;
    }
    updateStaffDetails_(staffDetails);
}

function updateStaffDetails_(staffDetails) {
    var theSpreadsheetId = PropertiesService.getScriptProperties().getProperty("inspectorateSpreadsheet");
    var theSpreadsheet = SpreadsheetApp.openById(theSpreadsheetId);
    var theSheet = theSpreadsheet.getSheetByName("Inspectors");

    if (staffDetails.rowNumber != undefined) {
        theSheet.getRange("A" + (staffDetails.rowNumber + 1) + ":B" + (staffDetails.rowNumber + 1))
            .setValues([[staffDetails.staffEmail, JSON.stringify(staffDetails)]]);
    }
    else {
        theSheet.appendRow([staffDetails.staffEmail, JSON.stringify(staffDetails)]);
    }

    return true;
}

function generateRef_() {
    var ref = Number(PropertiesService.getScriptProperties().getProperty("referenceNumber"));
    PropertiesService.getScriptProperties().setProperty("referenceNumber", ref + 1);
    return ref.toFixed(0);
}

function getFinanceEmail() {
    return PropertiesService.getScriptProperties().getProperty("financeEmail");
}

function buildCodeList_(theSpreadsheet) {
    var theCodeRange = theSpreadsheet.getRangeByName("ExpenseCodes");

    var theValues = theCodeRange.getValues();
    var codeString = "Select cost code";

    for (var i = 0; i < theValues.length; i++) {
        if (theValues[i][0] == "") { break; }
        codeString += "|" + theValues[i][0] + " - " + theValues[i][1];
    }

    return codeString;
}
