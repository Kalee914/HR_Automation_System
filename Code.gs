function unlockColumns() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("sheet_name");
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  
  protections.forEach(function(protection) {
    protection.remove(); // Remove all existing protections
  });
}

function lockColumns() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("sheet_name");
  // Specify the ranges for columns to relock
  var columnsToRelock = ['F', 'G', 'I', 'J', 'L', 'M', 'O', 'Q', 'R', 'T'];

  // Loop through each specified column and relock it
  columnsToRelock.forEach(function(col) {
    var range = sheet.getRange(col + '1:' + col);  // Get the entire column from row 1 onward
    var protection = range.protect(); // Reapply protection
    protection.addEditor(Session.getEffectiveUser()); 
    protection.removeEditors(protection.getEditors()); 
  });
  var rowsToLock = sheet.getRange('1:2'); // Lock rows 1 and 2
  var rowProtection = rowsToLock.protect();
  rowProtection.addEditor(Session.getEffectiveUser()); 
  rowProtection.removeEditors(rowProtection.getEditors());
}

function moveData() {
  var sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = sourceSpreadsheet.getSheetByName("sheet_name");

  // Define destination spreadsheets and sheets
  var terminatedSheet = SpreadsheetApp.openById("sheet_id").getSheetByName("sheet_name");
  var internationInternsSheet = SpreadsheetApp.openById("sheet_id").getSheetByName("sheet_name");
  var onboardedSheet = SpreadsheetApp.openById("sheet_id").getSheetByName("sheet_name");

  var data = sourceSheet.getDataRange().getValues();
  var headers = data[0]; 
  var columnMap = {};
  headers.forEach((header, index) => {
    columnMap[header] = index;
  });



  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var preInterviewStatus = row[columnMap["Pre_Interview: Pass OR Fail"]];
    var interviewStatus = row[columnMap["Interview: Pass OR Fail"]];
    var selfTermination = row[columnMap["SELF Termination"]];
    var schoolRequirement = row[columnMap["School Requirement"]];
    var completeStatus = row[columnMap["Complete"]];
    
    if (String(preInterviewStatus).trim() === "Fail" || String(interviewStatus).trim() === "Fail" || selfTermination === true) {
      moveRow(row, terminatedSheet);
      sourceSheet.getRange(i + 1, 1, 1, sourceSheet.getLastColumn()).clearContent(); // Clear instead of delete

    } else if (schoolRequirement === "OPT" || schoolRequirement === "CPT") {
      moveRow(row, internationInternsSheet);
      sourceSheet.getRange(i + 1, 1, 1, sourceSheet.getLastColumn()).clearContent();

    } else if (completeStatus === true) {
      moveRow(row, onboardedSheet);
      sourceSheet.getRange(i + 1, 1, 1, sourceSheet.getLastColumn()).clearContent();

    }
  }

}

// Helper function to append a row to a target sheet
function moveRow(row, targetSheet) {
  targetSheet.appendRow(row);
}

function deleteEmptyRows() {
  var sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = sourceSpreadsheet.getSheetByName("sheet_name");

  var data = sourceSheet.getDataRange().getValues();
  var rowsToDelete = [];

  // Find empty rows (starting from the bottom to avoid index shifting)
  for (var i = data.length - 1; i >= 2; i--) {  
    if (data[i].every(cell => cell.toString().trim() === "")) { // Check if all cells in the row are empty
      //sourceSheet.deleteRow(i + 1); 
      rowsToDelete.push(i + 1); // Convert to 1-based row index
    }
  }

  // Delete all empty rows in bulk
  rowsToDelete.forEach(row => sourceSheet.deleteRow(row));
}

function sendEmailUsingGmailAPI(recipient, subject, htmlBody, aliasEmail) {
  var rawMessage = [
    "From: " + aliasEmail,
    "To: " + recipient,
    "Subject: " + subject,
    "MIME-Version: 1.0",
    "Content-Type: text/html; charset=UTF-8",
    "",
    htmlBody
  ].join("\r\n");

  var encodedMessage = Utilities.base64EncodeWebSafe(rawMessage); // Base64 encode the message

  var email = {
    raw: encodedMessage
  };

  try {
    // Send email using Gmail API
    Gmail.Users.Messages.send(email, "me");
  } catch (error) {
    Logger.log("Error sending email via Gmail API: " + error);
  }
}

function sendPreInterviewEmails() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName("sheet_name");
    var preInterviewTemplate = HtmlService.createTemplateFromFile('preInterviewTemplate');
    var preInterviewFollowUpTemplate = HtmlService.createTemplateFromFile('preInterviewFollowUpTemplate');
    var interviewSheet = ss.getSheetByName("Pre_Interview Testing");

    var data = mainSheet.getDataRange().getValues();
    var interviewData = interviewSheet.getDataRange().getValues();
    
    var roleToLink = {}; // Store Position-Question link mapping
    for (var i = 1; i < interviewData.length; i++) {
        roleToLink[interviewData[i][0]] = interviewData[i][1]; // Role -> Question Link
    }

    var emailCount = 0; //counter
    var followUpCount = 0; // Counter for follow-up emails
    var maxOnBoardEmails = 20; // Max pre-interview emails
    var maxFollowUpEmails = 10; // Max follow-up emails
    var aliasEmail = "alias_email"; // verified alias
    var today = new Date();

    for (var i = 1; i < data.length; i++) {
        if (emailCount >= maxOnBoardEmails && followUpCount >= maxFollowUpEmails) break; // Stop after max emails

        var email = data[i][2].trim(); // (Recipient Mail)
        var preInterviewSentStatus = data[i][5]; //(Pre-interview Sent)
        var preInterviewFollowUpStatus = data[i][6]; // (Follow-Up Sent)
        var preInterviewReceived = data[i][8]; //PreInterview Resonse Received
        var name = data[i][0]; //(Full Name)
        var position = data[i][4]; //(Position)
        
        var link = roleToLink[position] || "https://default-link.com"; // Get link from mapping

        if (/^[\w.-]+@[\w.-]+\.\w+$/.test(email)) {
            if (!preInterviewSentStatus && preInterviewReceived === "N/A" && emailCount < maxOnBoardEmails) {
                // Send Pre-Interview Email
                preInterviewTemplate.name = name;
                preInterviewTemplate.position = position;
                preInterviewTemplate.link = link;
          
                var htmlBody = preInterviewTemplate.evaluate().getContent();

                try {
                  sendEmailUsingGmailAPI(email, "Pre-Interview Steps for " + position + " Internship at __________", htmlBody, aliasEmail);

                    mainSheet.getRange(i + 1, 6).setValue(new Date().toLocaleDateString() + " - Pre-interview sent");
                    emailCount++;
                } catch (error) {
                    Logger.log("Failed to send Pre-Interview email to " + email + ": " + error.message);
                }
            } else if (!preInterviewFollowUpStatus && preInterviewReceived === "N/A" && preInterviewSentStatus) {
                // Check if exactly 7 days have passed since pre-interview email
                var sentDate = new Date(preInterviewSentStatus.split(" - ")[0]); // Extract date from status
                var diffDays = Math.floor((today - sentDate) / (1000 * 60 * 60 * 24));

                if (diffDays >= 7 && followUpCount < maxFollowUpEmails) {
                    // Send Follow-Up Email
                    preInterviewFollowUpTemplate.name = name;
                    preInterviewFollowUpTemplate.position = position;
                    preInterviewFollowUpTemplate.link = link;
                    
                    var htmlBody = preInterviewFollowUpTemplate.evaluate().getContent();

                    try {
                        sendEmailUsingGmailAPI(email, "Pre-Interview Steps for " + position + " Internship at ____________", htmlBody, aliasEmail);

                        mainSheet.getRange(i + 1, 7).setValue(new Date().toLocaleDateString() + " - Follow-up sent");
                        followUpCount++; // Increment follow-up count
                    } catch (error) {
                        Logger.log("Failed to send Follow-Up email to " + email + ": " + error.message);
                    }
                }
            }
        }
    }
}
function sendCalendlyEmails() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName("sheet_name");
    var calendlyTemplate = HtmlService.createTemplateFromFile('calendlyTemplate');
    var calendlyFollowUpTemplate = HtmlService.createTemplateFromFile('calendlyFollowUpTemplate');

    var data = mainSheet.getDataRange().getValues();
   
    var calendlyEmailCount = 0; // counter
    var calendlyFollowUpCount = 0; // Counter for follow-up emails
    var maxCalendlyEmails = 20; // Max calendly emails
    var maxCalendlyFollowUpEmails = 10; // Max follow-up emails
    var aliasEmail = "alias_email"; // verified alias
    var today = new Date();

    for (var i = 1; i < data.length; i++) {
        if (calendlyEmailCount >= maxCalendlyEmails && calendlyFollowUpCount >= maxCalendlyFollowUpEmails) break; // Stop after max emails
        
        var name = data[i][0]; // Column A (Full Name)
        var email = data[i][2].trim(); // Column C (Recipient email)
	var preInterviewStatus = data[i][10]; // Column K (Yu Yang: Pass OR Fail Sent)
	var calendlySentStatus = data[i][11]; // Column L (Calendly Email Sent)
	var	calendlyFollowUpStatus = data[i][12]; // Column M (Calendly Follow-Up Email)
        var calendlyReserved = data[i][14]; // Column O (Calendly Appointment Reserved)


        if (/^[\w.-]+@[\w.-]+\.\w+$/.test(email) && String(preInterviewStatus).trim() === "Pass") {
            if (!calendlySentStatus && calendlyReserved === "N/A" && calendlyEmailCount < maxCalendlyEmails) {
			
                // Send calendly Email
                calendlyTemplate.name = name;
                
                var htmlBody = calendlyTemplate.evaluate().getContent();

                try {
                    sendEmailUsingGmailAPI(email, "Invitation to Connect - ____________ Internship Program!", htmlBody, aliasEmail);
                    mainSheet.getRange(i + 1, 12).setValue(new Date().toLocaleDateString() + " - Calendly email sent");
                    calendlyEmailCount++;
                } catch (error) {
                    Logger.log("Failed to send calendly email to " + email + ": " + error.message);
                }
            } else if (!calendlyFollowUpStatus && calendlyReserved === "N/A" && calendlySentStatus) {
                // Check if exactly 7 days have passed since calendly email
                var sentDate = new Date(calendlySentStatus.split(" - ")[0]); // Extract date from status
                var diffDays = Math.floor((today - sentDate) / (1000 * 60 * 60 * 24));

                if (diffDays >= 7 && calendlyFollowUpCount < maxCalendlyFollowUpEmails) {
                    // Send Follow-Up Email
                    calendlyFollowUpTemplate.name = name;
                    
                    var htmlBody = calendlyFollowUpTemplate.evaluate().getContent();

                    try {
                        sendEmailUsingGmailAPI(email, "Reminder: Follow Up on Your ____________ Internship Interview", htmlBody, aliasEmail);
                        mainSheet.getRange(i + 1, 13).setValue(new Date().toLocaleDateString() + " - Follow-up sent");
                        calendlyFollowUpCount++; // Increment follow-up count
                    } catch (error) {
                        Logger.log("Failed to send Calendly Follow-Up email to " + email + ": " + error.message);
                    }
                }
            }
        }
    }
}
function sendContractEmails() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName("sheet_name");
    var contractTemplate = HtmlService.createTemplateFromFile('contractTemplate');
    var contractFollowUpTemplate= HtmlService.createTemplateFromFile('contractFollowUpTemplate');


    var data = mainSheet.getDataRange().getValues();
   
    var contractEmailCount = 0;
    var contractFollowUpCount = 0; // Counter for follow-up emails
    var maxContractEmails = 5; // Max Contract emails
    var maxContractFollowUpEmails = 5; // Max follow-up emails
    var aliasEmail = "alias_email"; // verified alias
    var today = new Date();

    for (var i = 1; i < data.length; i++) {
        if (contractEmailCount >= maxContractEmails && contractFollowUpCount >= maxContractFollowUpEmails) break; // Stop after max emails
        
        var name = data[i][0]; // (Full Name)
        var email = data[i][2].trim(); //  (Recipient email)
	var interviewStatus = data[i][15]; // (Jason: Pass OR Fail)
	var contractSentStatus = data[i][16]; //(contract Email Sent)
	var	contractFollowUPStatus = data[i][17]; //(contract Follow-Up Email)
        var contractReceived = data[i][19]; //(contract Appointment Reserved)


        if (/^[\w.-]+@[\w.-]+\.\w+$/.test(email) && String(interviewStatus).trim() === "Pass") {
            if (!contractSentStatus && contractReceived === "N/A" && contractEmailCount < maxContractEmails) {
			
                // Send contract Email
                contractTemplate.name = name;
                
                var htmlBody = contractTemplate.evaluate().getContent();

                try {
                    sendEmailUsingGmailAPI(email, "[Action Required] Internship Contract for ____________", htmlBody, aliasEmail);

                    mainSheet.getRange(i + 1, 17).setValue(new Date().toLocaleDateString() + " - Contract email sent");
                    contractEmailCount++;
                } catch (error) {
                    Logger.log("Failed to send contract email to " + email + ": " + error.message);
                }
            } else if (!contractFollowUPStatus && contractReceived === "N/A" && contractSentStatus) {
                // Check if exactly 7 days have passed since contract email
                var sentDate = new Date(contractSentStatus.split(" - ")[0]); // Extract date from status
                var diffDays = Math.floor((today - sentDate) / (1000 * 60 * 60 * 24));

                if (diffDays >= 7 && contractFollowUpCount < maxContractFollowUpEmails) {
                    // Send Follow-Up Email
                    contractFollowUpTemplate.name = name;
                    
                    var htmlBody = contractFollowUpTemplate.evaluate().getContent();

                    try {
                        sendEmailUsingGmailAPI(email, "Reminder: Follow Up on Your Internship Contract [____________]", htmlBody, aliasEmail);

                        mainSheet.getRange(i + 1, 18).setValue(new Date().toLocaleDateString() + " - Follow-up sent");
                        contractFollowUpCount++; // Increment follow-up count
                    } catch (error) {
                        Logger.log("Failed to send contract Follow-Up email to " + email + ": " + error.message);
                    }
                }
            }
        }
    }
}
