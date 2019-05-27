//---UI functions---

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Sidebar")
    .addItem("Sidebar", "showSidebar")
    .addToUi();
}

function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle("GASmerge");
  SpreadsheetApp.getUi().showSidebar(ui);
}

//---Sidebar Button functions---

function listFolderContents(folderID) {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    
    ss.setActiveSheet(sheet);
    sheet.clearContents();
    
    var columns = [['file link','name', 'email', 'email status', 'date sent']];
    sheet.getRange(1,1,1,5) //getRange(start row, start column, number of rows, number of columns) 
        .setValues(columns)
        .setFontWeight("bold")
        .setBorder(null, null, true, null, null, null, "gray", null); //setBorder(top, left, bottom, right, vertical, horizontal, color, style)

    var attachmentFolder = DriveApp.getFolderById(folderID).getFiles();
    
    while (attachmentFolder.hasNext()) {
        
        var file = attachmentFolder.next(); 
        var name = file.getName();
        var link = file.getUrl();
        var formula = "=hyperlink(\"" + link + "\";\"" + name + "\")";
            
        sheet.appendRow([formula]);
    }
};

function mailMerge() {

    var today = Utilities.formatDate(new Date(), "CST", "MM-dd-yyyy");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var rows = sheet.getDataRange();    
    var values = rows.getValues(); 
    var formulas = rows.getFormulas();
    
    //---Grab docIDs from links---
    var docIDs = [];
    
        for (var h = 1; h < formulas.length; h++) { 
            var link = formulas[h][0];
            var docID = link.match(/[-\w]{25,}/);
            docIDs.push(docID); 
        }
      
    //---Generate emails---
    for (var i = 1; i < values.length; i++) {
        
        if (values[i][3] == "") {  // Prevents sending duplicates
          var emailAddresses = values[i][2];
          var emailArray = emailAddresses.split(',');
          var message = "This is a successful test email.";
          var fileId = docIDs[i-1]; //offset by one to account for the values header row
  
          var attachment = DriveApp.getFileById(fileId).getBlob();
          
            for (var k  = 0; k < emailArray.length; k++) {
              GmailApp.sendEmail(emailArray[k], 'This was sent using GASmerge!', message,
                                {name:'The PDF is attached to this email',
                                 attachments: [attachment.getAs(MimeType.PDF)]})
            }
          
          sheet.getRange(i+1,4,1,2).setValues([['EMAIL SENT',today]]);
          SpreadsheetApp.flush();// Make sure the cell is updated right away in case the script is interrupted
        
        }
    }
};
