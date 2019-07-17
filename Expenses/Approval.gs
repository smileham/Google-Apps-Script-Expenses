function manageApproval_(e) {
    try {
    var theSpreadsheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty("spreadsheet"));
    var theSheet       = theSpreadsheet.getSheetByName("Expenses");
    var theRange       = theSheet.getRange(2, 1, theSheet.getLastRow(), 13);
    var theEmail       = Session.getActiveUser().getEmail();
    var theValues      = theRange.getValues();
    var theSubmitterEmail = "";
    
    var outcome = "Error";
    
    //Start less 150, just in case a purge has run and we begin half way through a set of approvals.
    var theStartRow = e.parameter.startRow-150;
    if (theStartRow<0) {theStartRow=0;}
    
    // While loop to ensure that we search from the top to cover any purged rows
    while (outcome=="Error") {
      Logger.log("While Loop");
      for (var i=theStartRow; i<theValues.length; i++) {
        Logger.log("Checking Row: "+i +" : "+outcome);
        if (theValues[i][0]=="") {break;}
        if (theValues[i][10]==e.parameter.referenceNumber && theValues[i][11].toString().toUpperCase()==theEmail.toUpperCase()) {
          var theSubmitterEmail = theValues[i][0];
          var toUpdate = theSheet.getRange('M'+(i+2)+":N"+(i+2));
          if (theValues[i][12]=="" && e.parameter.outcome=="reject") {outcome = "Rejected"}
          else if (theValues[i][12]=="" && e.parameter.outcome=="approve") {outcome = "Approved"}
          else {outcome = "Error: "+e.parameter.outcome; break;}
          toUpdate.setValues([[outcome, new Date()]]);
        }
      }
      if (outcome=="Error") {
        if (theStartRow>0) {theStartRow=0;}
        else {outcome="Error: Expense not found";}
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
      
      var app = HtmlService.createTemplateFromFile("Approve");
      
      app.outcome=outcome;
      
      return app.evaluate()
        .addMetaTag("viewport", "width=device-width, initial-scale=1")
        .setTitle("Expenses Approval")
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  }
  catch (e) {
    MailApp.sendEmail("steven.mileham@rspca.org.uk", "error in approval app", e.message);
  }
}
