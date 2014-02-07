/*
* Uses the "Expenses" library, ensure it is included in the "Resources" > "Libraries" section;
* (Current Library code: MM2eWjjSwYU_HrcgVtbxfCdTczyYZLoOb)
*/

function doGet() {
  try {
    var app = UiApp.createApplication().setTitle("Submit Receipts");
    var theForm = app.createFormPanel();
    var textStyle = {fontSize:"16px", width:"100%", textAlign:"center", paddingBottom:"5px"};
    var helpStyle = {fontSize:"10px", width:"100%", textAlign:"center", paddingBottom:"2px"};
  
    var staffDetails = {staffEmail:Session.getActiveUser().getEmail(), staffNumber:"",staffCostCentre:"",staffCostLine:"",staffLineManagerEmail:""};
    staffDetails = Expenses.lookupStaffDetails(staffDetails);
    
    if (staffDetails.staffNumber!="") {
      
      var fileNumber = app.createHidden("fileNumber","2").setId("fileNumber");
      var refId = app.createListBox().setId("refId").setName("refId").setStyleAttributes({fontSize:"16px", width:"100%", border:"1px solid black", borderRadius:"2px"});
      
      var referenceNumbers = Expenses.retrieveReferenceNumbers(Session.getActiveUser().getEmail());
      while (referenceNumbers.hasNext()) {
        var refNum = referenceNumbers.next();
        refId.addItem(refNum.referenceNumber + " : £" + refNum.total );
      }
      
      var shareList = app.createHidden("shareList", staffDetails.staffLineManagerEmail+","+Expenses.getFinanceEmail())
      var receiptHandler = app.createServerHandler("addFileUploadHandler").addCallbackElement(theForm);
      
      var pleaseWait = app.createLabel("Please Wait...").setVisible(false).setId("pleaseWait").setStyleAttributes(textStyle);
      var safeHandler = app.createClientHandler().forEventSource().setEnabled(false).forTargets([pleaseWait]).setVisible(true);
      var safeSubmitHandler = app.createClientHandler().forEventSource().setVisible(false).forTargets([pleaseWait]).setVisible(true);
      
      var addUploadButton = app.createButton("Add Receipt").setId("addButton").addClickHandler(safeHandler)
      .addClickHandler(receiptHandler)
      .setStyleAttributes({width:"100%", margin:"0", borderRadius:"5px", minHeight:"30px"});
      
      var receiptButton = app.createSubmitButton("Submit").addClickHandler(safeSubmitHandler).setStyleAttributes({width:"100%", margin:"0", borderRadius:"5px", minHeight:"30px"}).setId("submitButton");
      
      var filePanel = app.createFlowPanel().setWidth("100%").setId("filePanel");
      
      addFileRow_(1,app);
      
      app.add(theForm.add(app.createScrollPanel(app.createFlowPanel()
                          .add(app.createLabel("Select Expense Reference (reference id : total value)").setStyleAttributes(textStyle))
                          .add(refId).add(shareList).add(fileNumber)
                          .add(app.createLabel("Tap 'Choose File' and select the photograph of the receipt.").setStyleAttributes(helpStyle))
                          .add(app.createLabel("Tap 'Add Receipt' to add another 'Choose File' row and select another photograph.").setStyleAttributes(helpStyle))
                          .add(app.createLabel("Tap 'Submit' once all the receipts have been added and wait for photographs to upload.").setStyleAttributes(helpStyle))
                          .add(app.createLabel("The receipts will be shared with your Line Manager and Finance from the 'Expense Receipts' folder in Google Drive.").setStyleAttributes(helpStyle))
                          .add(filePanel)
                          .add(addUploadButton)
                          .add(app.createHTML("<br>"))
                          .add(receiptButton)
                          .add(pleaseWait)
                         )));
    }
    else {
      app.add(app.createLabel("No Expenses Found")
              .setStyleAttributes({fontSize:"16px", width:"100%", textAlign:"center"}));
    }
    
    
    return app;
  }
  catch (e) {
    MailApp.sendEmail("steven.mileham@rspca.org.uk", "error in submit app", e.message);
  }
}

function addFileRow_(fileNumber, app) {
  var panel = app.getElementById("filePanel");
  var fileUpload = app.createFileUpload().setStyleAttributes({width:"100%", margin:"0", borderRadius:"5px", minHeight:"30px", fontSize:"16px"})
    .setName("theFile"+fileNumber).setId("theFile"+fileNumber);
  
  panel.add(fileUpload);
  
}

function addFileUploadHandler(e) {
  var app = UiApp.getActiveApplication();
  var fileNumber = Number(e.parameter.fileNumber);
  app.getElementById("fileNumber").setValue(fileNumber+1);
  addFileRow_(fileNumber, app);
  
  app.getElementById("pleaseWait").setVisible(false);
  app.getElementById("addButton").setEnabled(true);
  
  return app;
}
  

function doPost(e) {
  
  var app = UiApp.getActiveApplication();
  var fileNumber = Number(e.parameter.fileNumber);
  var refId = e.parameter.refId.split(" : ")[0];
  
  var emailBody="<div>Following Receipts have been uploaded and shared;</div><ul>";
  
  var folderList = DriveApp.getFoldersByName("Expense Receipts");
  var theExpenseFolder = null;
  if (folderList.hasNext()) {
    theExpenseFolder = folderList.next();
  }
  else {
    theExpenseFolder = DriveApp.createFolder("Expense Receipts");
  }
  
  var folderList = theExpenseFolder.getFoldersByName("Receipts: "+refId);
  var theFolder = null;
  if (folderList.hasNext()) {
    theFolder = folderList.next();
  }
  else {
    theFolder = theExpenseFolder.createFolder("Receipts: "+refId);
  }
  //Fix fot email notification from DriveApp
  DocsList.getFolderById(theFolder.getId()).addEditors(e.parameter["shareList"].split(","));
  
  for (var i=1; i<fileNumber; i++) {
    var theBlob = e.parameter["theFile"+i];
    theBlob.setContentType("image/jpeg");
    theBlob.setName("Receipt "+i+" of "+(fileNumber-1)+" reference:"+refId);
    var theFile = theFolder.createFile(theBlob);
    
    emailBody += "<li>Receipt "+i+" of "+(fileNumber-1)+" reference:"+refId+"</li>";
  }
  
  emailBody+="</ul><div>They have been uploaded to: <a href='"+theFolder.getUrl()+"'>Receipts: "+refId+"</a>";
  
  MailApp.sendEmail({
    to:       e.parameter.shareList,
    subject:  "[Approval] - Expenses - ref: "+refId,
    htmlBody: emailBody,
    noReply:  true
  });
  
  MailApp.sendEmail({
    to:       Session.getActiveUser().getEmail(),
    subject:  "Expenses - ref: "+refId,
    htmlBody: emailBody,
    noReply:  true
  });
  
  app.remove(0);
    app.add(app.createFlowPanel()
            .add(app.createLabel("Receipts uploaded").setStyleAttributes({width:"100%",fontSize:"20px",textAlign:"center"}))
           );
  
  return app;
}
