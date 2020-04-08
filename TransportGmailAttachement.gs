/**
 * The seperator for the end of the csv row.
 * @private {string}
 */
const endOfLineSep = '\n';
/**
 * The seperator for the end of the csv column.
 * @private {string}
 */
const columnSep = ';';
/**
 * The type name of allowed csv type.
 * @private {string}
 */
const allowedFileType = 'text/csv';
/**
 * The search query to look for in the gmail account.
 * @private {string}
 */
const searchQuery = 'label:GAS-New';
/**
 * The id of the spreadsheet where the data has to be imported to.
 * @private {string}
 */
const spreadSheetId = '1KBYEJ2SD1RT5VCxEkjbhu2M7HU0zk3wNVOj8BsscFvI';
/*
 * Contains the threads the code is currently working with.
 * @private {GmailThread?}
 */
var currentThread = null;
/**
 * The label name which the thread has at the beginning of that process.
 * @private {string}
 */
const beginningLabel = 'GAS-New';
/**
 * The label name which the thread has at the end of that process.
 * @private {string}
 */
const doneLabel = 'GAS-Old';
/**
 * The label name which the thread has get when an error occures.
 * @private {string}
 */
const errorLabel = 'GAS-Error';
/**
 * The email address of whom should receive the error.
 * @private {string}
 */
const emailToReport = 'any@email.com';


/*
 * Searchs threads by the const search query in the gmail account.
 */
function CheckGmailAccount() {
  var threads = GmailApp.search(searchQuery);
  
  if(threads !== undefined && threads.length > 0) {
    for(var i = 0; i < threads.length; i++) {
      currentThread = threads[i];
      GetAttachementFromThread(threads[i]);
    }
  }
}
/*
 * Itterates through the attachements and checks the file type.
 * If it is csv, it calls the corresponding method to proceed 
 * the content. 
 * @param {string} Gmail thread
 */
function GetAttachementFromThread(thread)
{
  if(thread.getMessages()[0].getAttachments()[0].getContentType() == allowedFileType) {
    AddContentToSpreadSheet(thread.getMessages()[0].getAttachments()[0].getDataAsString().split(endOfLineSep))
  } else {
    Logger.log("Wasn't able to read attachement. It doesn't seem to be a csv file.");
    SendEmail('Google App Script error!',"Wasn't able to read the email attachement. It doesn't seem to be a csv file.");
  }
}
/*
 * Appands the given content to the corresponding spreadsheet.
 * First row contains the titles.
 * @param {array} converted csv content
 */
function AddContentToSpreadSheet(contentRows)
{
  var ss = SpreadsheetApp.openById(spreadSheetId);
  var sheet = ss.getSheets()[0];
  
  for(var y = 1; y < contentRows.length; y++) {
    var columns = contentRows[y].split(columnSep);
    sheet.getRange(sheet.getLastRow()+1, 1, 1, columns.length).setValues([columns]);
  }
  SpreadsheetApp.flush();
  ChangeLabel(GmailApp.getUserLabelByName(doneLabel));
}
/*
 * Changes the label of the current thread to mark it as done.
 */
function ChangeLabel(newLabel)
{
  var currentLabel = GmailApp.getUserLabelByName(beginningLabel);
  currentThread.removeLabel(currentLabel);
  currentThread.addLabel(newLabel)
}
/*
 *
 * @param {string} subject of the email
 * @param {string} message of the email
 */
function SendEmail(subject, message)
{
  GmailApp.sendEmail(emailToReport, subject, message);
  ChangeLabel(GmailApp.getUserLabelByName(errorLabel)); 
}
