function lookupStaffDetails(staffDetails) {
  var db = ScriptDb.getMyDb();
  
  var result = db.query({staffEmail:staffDetails.staffEmail});
  if (result.hasNext()) {
    var staffDetails = result.next();
  }
  
  return staffDetails;
}

/* no approval */
function storeStaffDetails_(theEmail,staffNo, costCentre, costLine) {
  storeStaffDetails_(theEmail, staffNo, costCentre, costLine, "");
}

/* approval */
function storeStaffDetails_(theEmail,staffNo, costCentre, costLine, lineManagerEmail) {
  var db = ScriptDb.getMyDb();
  var result = db.query({staffEmail:theEmail});
  if (result.hasNext()) {
    var staffDetails = result.next();
    staffDetails.staffNumber = staffNo;
    staffDetails.staffCostCentre = costCentre;
    staffDetails.staffCostLine = costLine;
    staffDetails.staffLineManagerEmail = lineManagerEmail;
    
    db.save(staffDetails);
  }
  else {
    var staffDetails = {staffEmail:theEmail,
                        staffNumber: staffNo,
                        staffCostCentre: costCentre,
                        staffCostLine: costLine,
                        staffLineManagerEmail: lineManagerEmail};
    db.save(staffDetails);
  }
}

function storeReferenceNumber_(theEmail, theReferenceNumber, theTotal) {
  var db = ScriptDb.getMyDb();
  var referenceStore = {referenceEmail: theEmail, referenceNumber: theReferenceNumber, total: theTotal};
  db.save(referenceStore);
}

function retrieveReferenceNumbers(theEmail) {
  var db = ScriptDb.getMyDb();
  var results = db.query({referenceEmail:theEmail}).sortBy('referenceNumber', db.DESCENDING);;
  return results;
}


function generateRef_() {
  var ref = Number(ScriptProperties.getProperty("referenceNumber"));
  ScriptProperties.setProperty("referenceNumber", ref+1);
  return ref.toFixed(0);
}

function getFinanceEmail() {
  return ScriptProperties.getProperty("financeEmail");
}

function getSpreadsheetId() {
  return ScriptProperties.getProperty("spreadsheet");
}
