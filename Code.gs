function extractCalendlyInvites() {
  // Open google sheet 
  var destinationSpreadsheet = SpreadsheetApp.openByUrl("google_sheet_url");
  var sheetName = "Desire Sheet Name or Already created Sheet Name";  //Sheet name
  var destinationSheet = destinationSpreadsheet.getSheetByName(sheetName); // Check for sheet name

  // If not sheet with sheet name, then insert sheet 
  if (!destinationSheet) {
    destinationSheet = destinationSpreadsheet.insertSheet(sheetName);
  }

  // If sheet is empty, create headers
  if (destinationSheet.getLastRow() === 0) {
    destinationSheet.appendRow(['Invitee Name', 'Invitee Email', 'Event Date/Time']);
  }
  // Search unread Email from:notifications@calendly.com 
  var threads = GmailApp.search('from:notifications@calendly.com is:unread');
  var labelName = "Calendy Scheduled";// Email label name
  var label = GmailApp.getUserLabelByName(labelName); // Check for label
  
  // If can't find label then create label
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }
  // Looping through the emails, to find "threads" = search unread eamil from:notifications@calendly.com
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j].getBody();
      
      // Log the raw email content for debugging
      Logger.log("Raw Email Body:\n" + message);
      
      // Clean HTML tags from the message
      var cleanMessage = message.replace(/<[^>]+>/g, ' ');

      // Extract Name, Email, and Date/Time using regex
      var nameMatch = cleanMessage.match(/Invitee:\s*([A-Za-z\s]+)(?=\s*Invitee Email:)/);
      var emailMatch = cleanMessage.match(/Invitee Email:\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
      var dateTimeMatch = cleanMessage.match(/Event Date\/Time:\s*([A-Za-z0-9\s:,-]+)/);
		
	  // Extract the name, email, date/time if found, otherwise default to "N/A"
      var name = nameMatch ? nameMatch[1].trim() : "N/A";
      var email = emailMatch ? emailMatch[1].trim() : "N/A";
      var dateTime = dateTimeMatch ? dateTimeMatch[1].trim() : "N/A";

      // Log extracted data for debugging
      Logger.log("Extracted Name: " + name);
      Logger.log("Extracted Email: " + email);
      Logger.log("Extracted Date/Time: " + dateTime);

      // If all fields are valid, then add to the extracted data to the sheet
      if (name !== "N/A" && email !== "N/A" && dateTime !== "N/A") {
        destinationSheet.appendRow([name, email, dateTime]);
      }
    }
	// Mark the email thread as read 
    threads[i].markRead();
	// Add the label 
    threads[i].addLabel(label);
  }
}
