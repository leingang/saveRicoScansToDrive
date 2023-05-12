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
function savePDFsToDrive() {
  // Set the email address to filter
  // TODO: maybe just use a query string?
  // "from:me has:attachment subject:RNP5838796D08BE"
  var emailAddress = "me";
  var printerName = "RNP5838796D08BE";

  // Set the name of the Google Drive folder to save the PDFs
  var folderName = "Downloads";

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

  // Get all the threads in the Inbox
  var threads = GmailApp.search("in:inbox from:" + emailAddress +" has:attachment subject:" + printerName);

  // Loop through all the threads in the Inbox
  for (var i = 0; i < threads.length; i++) {
    console.log("Thread: ", i);
    var thread = threads[i];

    // Get all the messages in the thread
    var messages = thread.getMessages();

    // Loop through all the messages in the thread
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      console.log("Subject: ", message.getSubject())

        // Get all the attachments in the message
        var attachments = message.getAttachments();

        // Loop through all the attachments in the message
        for (var k = 0; k < attachments.length; k++) {
          var attachment = attachments[k];

          // Check if the attachment is a PDF file
          if (attachment.getContentType() === "application/pdf") {

            // Save the PDF file to the Google Drive folder
            var file = folder.createFile(attachment);
            Logger.log("Saved PDF file: " + file.getName());
            GmailApp.markThreadRead(thread).moveThreadToArchive(thread);
          }
        }
      }
    }
  }

