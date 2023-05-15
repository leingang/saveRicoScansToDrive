/**
 * MPL: Can you write a Google App Script to save all PDFs attached
 * to emails coming from a specific address to a Google Drive folder?
 *
 * ChatGPT: Yes, I can help you write a Google Apps Script to save all 
 * PDFs attached to emails coming from a specific address to a Google 
 * Drive folder. Here is an example script that you can modify to suit 
 * your specific needs:
 * 
 * 2023-05-12: Now a github repository
 * 
 */

/**
 * Get a folder from Google Drive
 * Creates the folder if it doesn't already exist.
 * 
 * @param {string} folderName 
 * @returns {Folder}
 */
function getDriveFolder(folderName) {
  // Get the root folder of the Google Drive
  var rootFolder = DriveApp.getRootFolder();
  // Check if the folder already exists in the root folder
  var folders = rootFolder.getFoldersByName(folderName);
  var folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    // Create the folder if it does not exist
    folder = rootFolder.createFolder(folderName);
  }
  return folder;
}

/**
 * Get all the threads that are sent from one of the printers
 * 
 * @param {RegExp} regex
 * @param {Function} callback 
 */
function findThreadsFromPrinter(regex, callback) {
  // Get all the threads in the Inbox
  var threads = GmailApp.search("in:inbox from:me has:attachment");
  for (var i = 0; i < threads.length; i++) {
    // Logger.log("Thread: " + i);
    var thread = threads[i];
    var subject = thread.getFirstMessageSubject();
    Logger.log("Subject: " + subject);  
    var messages = thread.getMessages();
    if (subject.match(regex)) {
      callback(thread);
    }
  }
}

/**
 * 
 * @param {GmailThread} thread 
 * @param {Folder} folder 
 */
function savePdfsFromThreadToFolder(thread,folder) {
  messages = thread.getMessages();
  for (var j = 0; j < messages.length; j++) {
    var message = messages[j];        
    savePdfsFromMessageToFolder(message,folder);
  }
}

/**
 * 
 * @param {Message} message 
 * @param {Folder} folder
 */
function savePdfsFromMessageToFolder(message,folder) {
  var attachments = message.getAttachments();
  for (var k = 0; k < attachments.length; k++) {
    var attachment = attachments[k];
    if (attachment.getContentType() === "application/pdf") {
      var file = folder.createFile(attachment);
      Logger.log("Saved PDF file: " + file.getName());
    }
  }
}


function savePDFsToDrive() {
  var folderName = "Downloads";
  var regex = /RNP[0-9A-F]{12}/;

  folder = getDriveFolder(folderName);  
  findThreadsFromPrinter(regex, function(t) {
    savePdfsFromThreadToFolder(t,folder);
    GmailApp.markThreadRead(t).moveThreadToArchive(t);
  });
}
