/**
 * Save PDFs sent from Ricoh Multifunction Scanners to Google Drive
 * 
 * Matthew Leingang and ChatGPT
 * May 2023
 */

/**
 * The main function
 */
function savePDFsToDrive() {
  var folderName = "Downloads";
  var regex = /RNP[0-9A-F]{12}/;

  folder = getDriveFolder(folderName);  
  getThreadsFromPrinter(regex).forEach(
    function(t) {
      savePdfsFromThreadToFolder(t,folder);
      GmailApp.markThreadRead(t).moveThreadToArchive(t);
    }
  ); 
}

/**
 * Get a folder nameed `folderName` from Google Drive
 * 
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
 * Get all the GMail threads that are sent from a printer matching `regex`
 *
 * @param {RegExp} regex - regular expression to match printer hosts
 * @returns {GmailThread[]} - array of threads
 */
function getThreadsFromPrinter(regex) {
  threads = GmailApp.search("in:inbox from:me has:attachment");
  return threads.filter(function (t) {
      return t.getFirstMessageSubject().match(regex);
    });
}

/**
 * Save all PDFs attached to messages in `thread` into `folder`
 *
 * @param {GmailThread} thread - the thread to search for attachments
 * @param {Folder} folder - the folder in which to save them
 */
function savePdfsFromThreadToFolder(thread,folder) {
  messages = thread.getMessages();
  messages.forEach(function (m) {
    savePdfsFromMessageToFolder(m,folder);
  });
}

/**
 * Save all PDFs attached to `message` into `folder`
 * 
 * @param {Message} message - the message to find attachments 
 * @param {Folder} folder - the folder in which to save them
 */
function savePdfsFromMessageToFolder(message,folder) {
  var attachments = message.getAttachments();
  attachments.forEach(function (a) {
    if (a.getContentType() === "application/pdf") {
      var file = folder.createFile(a);
      Logger.log("Saved PDF file: " + file.getName());
    }
  });
}


