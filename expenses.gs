function doGet() {
  var theSpreadsheet = SpreadsheetApp.openById(ScriptProperties.getProperty("spreadsheet"));
  var approval = UserProperties.getProperty("approvalProcess")=="true" || ScriptProperties.getProperty("approvalProcess")=="true";
  
  var app = UiApp.getActiveApplication();
  
  if (app.getId()==null) {
    var app = UiApp.createApplication().setTitle("Expenses");
  }
  else {
    try {app.remove(0);} catch (e) {}
  }
  
  var fontStyle = {fontSize:"16px"};
  
  var rowTable = app.createFlexTable().setId("rowTable").setWidth("100%");
  var mainPanel = app.createFlowPanel().setWidth("100%");
  var scrollPanel = app.createScrollPanel().setWidth("100%").setHeight("100%");
  
  if (approval) {
    var staffDetails = {staffEmail:Session.getActiveUser().getEmail(), staffNumber:"",staffCostCentre:"",staffCostLine:"",staffLineManagerEmail:""};
  }
  else {
    var staffDetails = {staffEmail:Session.getActiveUser().getEmail(), staffNumber:"",staffCostCentre:"",staffCostLine:""};
  }
  staffDetails = lookupStaffDetails(staffDetails);
  var costCentreListBox = buildSpreadsheetListBox_("CostCentre", staffDetails.staffCostCentre, app, theSpreadsheet).setWidth("100%");
  var costLineListBox = buildSpreadsheetListBox_("CostLine", staffDetails.staffCostLine,app, theSpreadsheet).setWidth("100%");
  
  if (approval) {
    var expenseGrid = app.createGrid(4,2)
    .setText(3,0,"Line Manager Email").setStyleAttributes(3, 0, fontStyle)
    .setWidget(3,1,app.createTextBox().setName('lineManagerEmail').setId("lineManagerEmail").setText(staffDetails.staffLineManagerEmail)
               .setStyleAttributes({fontSize:"16px", width:"100%", border:"1px solid black", borderRadius:"2px"}));
  }
  else {
    var expenseGrid = app.createGrid(3,2);
  }
  expenseGrid.setWidth("100%")
    .setText(0,0,"Staff No.").setStyleAttributes(0, 0, fontStyle)
    .setText(1,0,"Cost Centre").setStyleAttributes(1, 0, fontStyle)
    .setText(2,0,"Cost Line").setStyleAttributes(2, 0, fontStyle)
    .setWidget(0,1,app.createTextBox().setName('staffNo').setId("staffNo").setText(staffDetails.staffNumber)
               .setStyleAttributes({fontSize:"16px", width:"100%", border:"1px solid black", borderRadius:"2px"}))
    .setWidget(1,1,costCentreListBox.setStyleAttributes(fontStyle))
    .setWidget(2,1,costLineListBox.setStyleAttributes(fontStyle));
 
  var formPanel = app.createFormPanel().setWidth("100%");
  
  var numRows = app.createHidden().setValue("1").setName("numRows").setId("numRows");
  var codeStr = buildCodeList_(theSpreadsheet);
  var codeString = app.createHidden().setValue(codeStr).setName("codeString").setId("codeString");
  
  var safeHandler = app.createClientHandler().forEventSource().setEnabled(false);
  
  var buttonStyle = {width:"100%", margin:"0", borderRadius:"5px", minHeight:"30px"};
  var addRowClickHandler = app.createServerHandler("addRowHandler").addCallbackElement(formPanel);
  var deleteRowClickHandler = app.createServerHandler("deleteRowHandler").addCallbackElement(formPanel);
  var submitExpensesHandler = app.createServerHandler("submitExpenses").addCallbackElement(formPanel);
  var addRowButton = app.createButton("Add Row", addRowClickHandler).addClickHandler(safeHandler).setStyleAttributes(buttonStyle).setId("addButton");
  var deleteRowButton = app.createButton("Delete Row", deleteRowClickHandler).addClickHandler(safeHandler).setStyleAttributes(buttonStyle).setId("deleteButton");
  var submitButton = app.createButton("Done", submitExpensesHandler).addClickHandler(safeHandler).setStyleAttributes(buttonStyle).setId("submitButton");
  createTableRow_(1,codeStr,app);
  var buttonPanel = app.createGrid(1, 3).setWidth('100%').setWidget(0, 0, addRowButton).setWidget(0,1, deleteRowButton).setWidget(0,2,submitButton)
  .setStyleAttributes(0,0,{width:"33%"}).setStyleAttributes(0,1,{width:"33%"}).setStyleAttributes(0,2,{width:"33%"});
  mainPanel.add(expenseGrid).add(rowTable).add(buttonPanel).add(numRows).add(codeString);
  scrollPanel.add(mainPanel);
  formPanel.add(scrollPanel);
  app.add(formPanel);
  return app;
}

function buildSpreadsheetListBox_(theRangeName, selected, theApp, theSpreadsheet) {
  var theListBox = theApp.createListBox().addItem("Please select").setName(theRangeName).setId(theRangeName);
  var theRange = theSpreadsheet.getRangeByName(theRangeName);
  
  var theValues = theRange.getValues();
  var selectedIndex = 0;
  
  for (var i=0; i<theValues.length; i++) {
    if (theValues[i]=="") {break;}
    if (theValues[i]==selected) {selectedIndex=i}
    theListBox.addItem(theValues[i]);
  }
  
  return theListBox.setSelectedIndex(selected==""?0:selectedIndex+1);
}

function buildCodeList_(theSpreadsheet) {
  var theCodeRange = theSpreadsheet.getRangeByName("ExpenseCodes");
  
  var theValues = theCodeRange.getValues();
  var codeString= "Select cost code";
  
  for (var i=0; i<theValues.length; i++) {
    if (theValues[i][0]=="") {break;}
    codeString+="|"+theValues[i][0] + " - " + theValues[i][1];
  }
  
  return codeString;
}

function createTableRow_(rowNumber,codeString, app) {
  var theRow = app.getElementById("rowTable");
  var bottomRow = rowNumber*2-1;
  var topRow = bottomRow-1;
  
  var descriptionBox = app.createTextBox().setValue("Description").setName("desc"+rowNumber)
    .setStyleAttributes({width:"100%",fontSize:"16px", border:"1px solid black", borderRadius:"2px"});
  descriptionBox.addFocusHandler(app.createClientHandler().forEventSource().validateMatches(descriptionBox,"Description").setText(""));
  
  theRow.setText(topRow,0,rowNumber);
  theRow.setWidget(topRow,1,descriptionBox);
  
  var codeArray = codeString.split("|");
  var codeList = app.createListBox().setName("type"+rowNumber).setId("type"+rowNumber).setStyleAttributes({width:"100%",fontSize:"16px"});
  for (var i=0; i<codeArray.length; i++) {
    codeList.addItem(codeArray[i]);
  }
  
  var rowGrid = app.createGrid(1,3).setWidth("100%");
  
  var dateBox = app.createDateBox().setName("date"+rowNumber).setValue(new Date())
  .setId("date"+rowNumber).setFormat(UiApp.DateTimeFormat.DATE_SHORT)
  .setStyleAttributes({width:"100%",fontSize:"16px", border:"1px solid black", borderRadius:"2px"});
  rowGrid.setWidget(0,0, dateBox).setStyleAttributes(0,0,{width:"30%"});
  
  var amountBox = app.createTextBox()
    .setStyleAttributes({width:"100%",fontSize:"16px", border:"1px solid black", borderRadius:"2px"})
    .setName("amount"+rowNumber).setId("amount"+rowNumber).setValue("0.00");
  amountBox.addFocusHandler(app.createClientHandler().forEventSource().validateSum([amountBox],0).setText(""));
  rowGrid.setWidget(0,1,codeList).setStyleAttributes(0,1,{width:"50%"});
 
  rowGrid.setWidget(0,2,amountBox);
  theRow.setText(bottomRow,0,"");
  theRow.setWidget(bottomRow,1,rowGrid).setStyleAttributes(bottomRow,1,{padding:"0"});
}

function addRowHandler(e) {
  var app = UiApp.getActiveApplication();
  var numRowsVal = Number(e.parameter.numRows);
  if (validateForm_(app,e)) {
    app.getElementById("numRows").setValue(numRowsVal+1);
    createTableRow_(numRowsVal+1,e.parameter.codeString,app)
  }
  
  app.getElementById("addButton").setEnabled(true);
  
  return app;
}

function deleteRowHandler(e) {
  var app = UiApp.getActiveApplication();
  var numRowsVal = Number(e.parameter.numRows);
  var theRowNum = (numRowsVal)*2;
  if (numRowsVal>1) {
    app.getElementById("rowTable").removeRow(theRowNum-1).removeRow(theRowNum-2); 
    app.getElementById("numRows").setValue(numRowsVal-1);
  }
  
  app.getElementById("deleteButton").setEnabled(true);
  
  return app;
}

function validateForm_(app, e) {
 
  var numRows = Number(e.parameter.numRows);
  var isValid = true;
  var errorStyle = {backgroundColor:"#FF9494"};
  var normalStyle = {backgroundColor:"white"};
  
  var approval = UserProperties.getProperty("approvalProcess")=="true" || ScriptProperties.getProperty("approvalProcess")=="true";
  
  if (e.parameter["CostCentre"]=="Please select") {
    isValid = false;
    app.getElementById("CostCentre").setStyleAttributes(errorStyle);
  }
  else {
    app.getElementById("CostCentre").setStyleAttributes(normalStyle);
  }
 
  if (approval) {
    if (e.parameter["lineManagerEmail"]=="" || !endsWith_(e.parameter["lineManagerEmail"],ScriptProperties.getProperty("emailDomain"))) {
      isValid = false;
      app.getElementById("lineManagerEmail").setStyleAttributes(errorStyle);
    }
    else {
      app.getElementById("lineManagerEmail").setStyleAttributes(normalStyle);
    }
  }
  
  if (e.parameter["CostLine"]=="Please select") {
    isValid = false;
    app.getElementById("CostLine").setStyleAttributes(errorStyle);
  }
  else {
    app.getElementById("CostLine").setStyleAttributes(normalStyle);
  }
  
  if (e.parameter["staffNo"]=="") {
    isValid = false;
    app.getElementById("staffNo").setStyleAttributes(errorStyle);
  }
  else {
    app.getElementById("staffNo").setStyleAttributes(normalStyle);
  }
  
  for (var i=1; i<numRows+1; i++) {
    if (e.parameter["date"+i]=="" || e.parameter["date"+i] > new Date()) {
      isValid = false;
      app.getElementById("date"+i).setStyleAttributes(errorStyle);
    }
    else {
      app.getElementById("date"+i).setStyleAttributes(normalStyle);
    }
    
    if (e.parameter["type"+i]=="Select cost code") {
      isValid = false;
      app.getElementById("type"+i).setStyleAttributes(errorStyle);
    }
    else {
      app.getElementById("type"+i).setStyleAttributes(normalStyle);
    }
    
    if (e.parameter["amount"+i].trim()=="" || isNaN(Number(e.parameter["amount"+i])) || e.parameter["amount"+i]<=0 ) {
      isValid = false;
      app.getElementById("amount"+i).setStyleAttributes(errorStyle);
    }
    else {
      app.getElementById("amount"+i).setStyleAttributes(normalStyle).setValue(Number(e.parameter["amount"+i]).toFixed(2));
    }
  }
   return isValid;
}

function submitExpenses(e) {
  /* Load Spreadsheet to track expenses */
  var app = UiApp.getActiveApplication();
  if (validateForm_(app, e)) {
    var theSpreadsheet = SpreadsheetApp.openById(ScriptProperties.getProperty("spreadsheet"));
    var theSheet = theSpreadsheet.getSheetByName("Expenses");
    var lastRow = theSheet.getLastRow();
    var theEmail = Session.getActiveUser().getEmail();
    var theDate = new Date().toDateString();
    var today = new Date().toUTCString();
    
    var approval = UserProperties.getProperty("approvalProcess")=="true" || ScriptProperties.getProperty("approvalProcess")=="true";
    
    if (approval) {
      storeStaffDetails_(theEmail, e.parameter.staffNo, e.parameter.CostCentre, e.parameter.CostLine, e.parameter.lineManagerEmail);
    }
    else {
      storeStaffDetails_(theEmail, e.parameter.staffNo, e.parameter.CostCentre, e.parameter.CostLine);
    }
  
    //var referenceNumber = hash_(Session.getActiveUser()+new Date().getTime());
    var referenceNumber = generateRef_();
    var numRows = Number(e.parameter.numRows);
    
    var userTable = '<table style="border-spacing: 5px; width:100%">'+
      '<tr><th style="background-color:#C0C0C0">User</th><td style="background-color:#f6f6f6;">'+theEmail+'</td></tr>'+
      '<tr><th style="background-color:#C0C0C0">Staff number</th><td>'+e.parameter.staffNo+'</td></tr>'+
      '<tr><th style="background-color:#C0C0C0">Cost Centre</th><td style="background-color:#f6f6f6;">'+e.parameter.CostCentre+'</td></tr>'+
      '<tr><th style="background-color:#C0C0C0">Cost Line</th><td>'+e.parameter.CostLine+'</td></tr>';
    if (approval) {
      userTable+= '<tr><th style="background-color:#C0C0C0">Line Manager</th><td style="background-color:#f6f6f6;">'+e.parameter.lineManagerEmail+'</td></tr>';
    }
    userTable+='</table><br>'
    
    var expenseTable = '<table style="border-spacing: 5px 0; width:100%"><tr><th style="background-color:#C0C0C0">#</th><th style="background-color:#C0C0C0">Description</th><th style="background-color:#C0C0C0">Date</th>'+
      '<th style="background-color:#C0C0C0">Code</th><th style="background-color:#C0C0C0">Amount</th></tr>';
    var total = 0;
    
    for (var i=1; i<numRows+1; i++) {
      theDate = e.parameter["date"+i].toDateString();
      var theDescription = e.parameter["desc"+i]=="Description"?"":e.parameter["desc"+i];
      expenseTable += '<tr><td valign="top" style="'+(i%2!=0?'background-color:#f6f6f6;':'')+'">'+i+ '</td>'+
        '<td valign="top" '+(i%2!=0?'style="background-color:#f6f6f6;"':'')+'>'+ theDescription + '</td>'+
        '<td valign="top" style="'+(i%2!=0?'background-color:#f6f6f6;':'')+'">'+ theDate + '</td>'+
        '<td valign="top" '+(i%2!=0?'style="background-color:#f6f6f6;"':'')+'>'+ e.parameter["type"+i] + '</td>'+
        '<td valign="top" '+(i%2!=0?'style="background-color:#f6f6f6;"':'')+'>£' + Number(e.parameter["amount"+i]).toFixed(2) +'</td></tr>';
      total += Number(e.parameter["amount"+i]);
      if (approval) {
        theSheet.appendRow([theEmail, today, e.parameter.staffNo, e.parameter.CostCentre, e.parameter.CostLine, theDate, theDescription, e.parameter["type"+i], Number(e.parameter["amount"+i]).toFixed(2), referenceNumber, e.parameter.lineManagerEmail]);
      }
      else {
        theSheet.appendRow([theEmail, today, e.parameter.staffNo, e.parameter.CostCentre, e.parameter.CostLine, theDate, theDescription, e.parameter["type"+i], Number(e.parameter["amount"+i]).toFixed(2), referenceNumber]);
      }
    }
    
    expenseTable+="<tr><th colspan='4'>Total</th><th>£"+Number(total).toFixed(2)+"</th></tr></table>";
    storeReferenceNumber_(theEmail, referenceNumber, total);
    /* Submitter Email */
    MailApp.sendEmail({
      to:       theEmail,
      subject:  "Expenses - ref: "+referenceNumber,
      htmlBody: userTable + expenseTable,
      noReply:  true
    });
    
    if (approval) {
      var approvalTable='<table width="100%"><tr><td width="50%">'+
        '<div style="height:30px; text-align:center; width:100%; padding: 10px 0; border-radius: 5px; background-color:#94ffa6; font-weight:bold; font-size:18px"><a style="font-face:sans-serif; color:black;" href="'+ScriptProperties.getProperty("approvalURL")+'?outcome=approve&startRow='+lastRow+'&referenceNumber='+referenceNumber+'">Approve</a></div></td>'+
          '<td><div style="height:30px; text-align:center; width:100%; padding: 10px 0; border-radius: 5px; background-color:#ff9494; font-weight:bold; font-size:18px"><a style="font-face:sans-serif; color:black;" href="'+ScriptProperties.getProperty("approvalURL")+'?outcome=reject&startRow='+lastRow+'&referenceNumber='+referenceNumber+'">Reject</a></div></td>'+
            '</tr></table>';
      
      /* Line Manager Email */
      MailApp.sendEmail({
        to:       e.parameter.lineManagerEmail,
        subject:  "[Approval] - Expenses - ref: "+referenceNumber,
        htmlBody: userTable + expenseTable + approvalTable,
        noReply:  true
      });
    }
    
    app.remove(0);
    app.add(app.createFlowPanel()
            .add(app.createLabel("Expenses Submitted").setStyleAttributes({width:"100%",fontSize:"20px",textAlign:"center"}))
            .add(app.createLabel("Your reference number is").setStyleAttributes({width:"100%",fontSize:"16px",textAlign:"center"}))
            .add(app.createLabel(referenceNumber)
                 .setStyleAttributes({width:"100%", backgroundColor:"#f6f6f6", borderRadius:"5px", textAlign:"center", fontSize:"32px", fontWeight:"bold"}))
            .add(app.createLabel("Please submit receipts via the link below.").setStyleAttributes({width:"100%",fontSize:"16px",textAlign:"center"}))
            .add(app.createAnchor("Submit Receipts", ScriptProperties.getProperty("receiptURL"))
                 .setTarget("_self").setStyleAttributes({width:"100%",fontSize:"16px",textAlign:"center"}))
            .add(app.createHTML("<br>"))
            .add(app.createButton("Back to expenses",app.createServerHandler("doGet")).addClickHandler(app.createClientHandler().forEventSource().setEnabled(false))
                 .setStyleAttributes({width:"100%", margin:"0", borderRadius:"5px", minHeight:"30px"}))
           );
    
  }
  else {
    app.getElementById("submitButton").setEnabled(true);
  }
  return app;
}
