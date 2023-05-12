/**
 * Save PDFs emailed from Multifunction Printer to Google Drive
 *
 * ChatGPT, Matthew Leingang
 * 2023-05-09
 *
 * To use this script, follow these steps:
 *
 * 1. Open your Google Drive account.
 * 2. Go to the Google Drive folder where you want to save the PDFs.
 * 3. Click on "New" and select "Google Script".
 * 4. Copy and paste this script into the script editor.
 * 5. Modify the emailAddress, folderName, and printerName variables to suit your specific needs.
 * 6. Click on "Run" and select "savePDFsToDrive".
 * 7. Grant the necessary permissions to the script when prompted.
 * 8. Wait for the script to finish running. The PDFs will be saved to the specified folder in your Google Drive.
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

  // Get all the threads from the inbox sent by the printer with attachments
  var threads = GmailApp.search("in:inbox from:" + emailAddress +" has:attachment subject:" + printerName);

  // Loop through all the threads
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
            // Mark thread as read and archive it
            GmailApp.markThreadRead(thread).moveThreadToArchive(thread);
          }
        }
      }
    }
  }
