Script.gs

var ui = SpreadsheetApp.getUi();
function onOpen(e){
  
  ui.createMenu("Gmail Manager").addItem("Get Emails by Label", "getGmailEmails").addToUi();
  
}

function getGmailEmails(){
  var input = ui.prompt('Label Name', 'Enter the label name that is assigned to your emails:', Browser.Buttons.OK_CANCEL);
  
  if (input.getSelectedButton() == ui.Button.CANCEL){
    return;
  }
  
  var label = GmailApp.getUserLabelByName(input.getResponseText());
  var threads = label.getThreads();
  
  for(var i = threads.length - 1; i >=0; i--){
    var messages = threads[i].getMessages();
    
    for (var j = 0; j <messages.length; j++){
      var message = messages[j];
      extractDetails(message);
      }
    }
    threads[i].removeLabel(label);
  }
  
function extractDetails(message){
  var dateTime = message.getDate();
  var subjectText = message.getSubject();
  var senderDetails = message.getFrom();
  var bodyContents = message.getPlainBody();
  
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  activeSheet.appendRow([dateTime, senderDetails, subjectText, bodyContents]);
}