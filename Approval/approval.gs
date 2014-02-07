/*
* Uses the "Expenses" library, ensure it is included in the "Resources" > "Libraries" section;
* (Current Library code: MM2eWjjSwYU_HrcgVtbxfCdTczyYZLoOb)
*/

function doGet(e) {
  var theSpreadsheet = SpreadsheetApp.openById(Expenses.getSpreadsheetId());
  var theSheet       = theSpreadsheet.getSheetByName("Expenses");
  var theRange       = theSheet.getRange(2, 1, theSheet.getLastRow(), 12);
  var theEmail       = Session.getActiveUser().getEmail();
  var theValues      = theRange.getValues();
  var theSubmitterEmail = "";
  
  var outcome = "Error";
  // Need to test last row function
  for (var i=e.parameter.startRow-1; i<theValues.length; i++) {
    if (theValues[i][0]=="") {break;}
    if (theValues[i][11]=="" && theValues[i][9]==e.parameter.referenceNumber && theValues[i][10]==theEmail) {
      var theSubmitterEmail = theValues[i][0];
      var toUpdate = theSheet.getRange('L'+(i+2)+":M"+(i+2));
      if (e.parameter.outcome=="reject") {outcome = "Rejected"}
      else if (e.parameter.outcome=="approve") {outcome = "Approved"}
      else {outcome = "Error: "+e.parameter.outcome}
      toUpdate.setValues([[outcome, new Date().toUTCString()]]);
    }
  }
  
  if (theSubmitterEmail!="") {
    /* Submitter Email */  
    MailApp.sendEmail({
      to:       theSubmitterEmail,
      subject:  "Expenses - ref: "+e.parameter.referenceNumber,
      htmlBody: "<div style='width:100%; border-radius:5px; text-align:center; font-size:32px; font-weight:bold; background-color:#f6f6f6'>"+outcome+"</div>",
      noReply:  true
    });
  }
    
   /* Line Manager Email */
      MailApp.sendEmail({
        to:       theEmail,
        subject:  "[Approval] - Expenses - ref: "+e.parameter.referenceNumber,
        htmlBody: "<div style='width:100%; border-radius:5px; text-align:center; font-size:32px; font-weight:bold; background-color:#f6f6f6'>"+outcome+"</div>",
        noReply:  true
      });
  
  var app = UiApp.createApplication().setTitle("Expense Approval");
  app.add(app.createFlowPanel().add(app.createLabel(outcome)
                 .setStyleAttributes({width:"100%", backgroundColor:"#f6f6f6", borderRadius:"5px", textAlign:"center", fontSize:"32px", fontWeight:"bold"})));
  return app;      
}
