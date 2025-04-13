// Helper: Generate UUID
function generateUniqueId() {
  return Utilities.getUuid();
}

function onApprovalFormSubmit(e) {
  if (!e || !e.values) {
    Logger.log("ERROR: No event data detected. This script should be triggered by a form submission.");
    return;
  }

  var ss = SpreadsheetApp.openById("1048SyHN-e4XujUkJOXj1j5rMQdZ4R4ZwIEYcLTK0OJk");
  var sheet = e.range.getSheet();
  var sheetId = sheet.getSheetId();
  Logger.log("Sheet ID: " + sheetId);
  Logger.log("🚨 e.values received: " + JSON.stringify(e.values));


  var sourceSheet = ss.getSheetById(1447833263);
  var statusSheet = ss.getSheetById(684280098);
  var chairSheet = ss.getSheetById(362146665);
  var hrSheet = ss.getSheetById(1074532913);
  var oorSheet = ss.getSheetById(876106890);

// ----- 1. Chair Form -----
if (sheetId === 362146665) {
  Logger.log("Chair Form Submission (e.values): " + JSON.stringify(e.values));

  var chairApprovalStatus = e.values[5]; // Approval decision
  var rejectionReason = e.values[6] || "No comments provided."; // Comment
  var uniqueId = e.values[7]; // Unique ID

  Logger.log("Chair approval status: " + chairApprovalStatus);
  Logger.log("Chair form Unique ID: " + uniqueId);

  var rowIndex = e.range.getRow();
  var chairHeaders = chairSheet.getRange(1, 1, 1, chairSheet.getLastColumn()).getValues()[0];
  var uniqueIdCol = chairHeaders.indexOf("Unique ID") + 1;
  if (uniqueIdCol > 0) {
    chairSheet.getRange(rowIndex, uniqueIdCol).setValue(uniqueId);
  }

  var sourceRow = sourceSheet.getDataRange().getValues().find(row => row[23] === uniqueId);
  if (!sourceRow) {
    Logger.log("❌ No source row found for Unique ID: " + uniqueId);
    return;
  }

  var requesterEmail = sourceRow[1]; // Column B
  var requesterName = sourceRow[2];
  var volunteerName = sourceRow[6];

  var statusData = statusSheet.getRange(1, 1, statusSheet.getLastRow(), 1).getValues();
  for (var i = 0; i < statusData.length; i++) {
    if (statusData[i][0] === uniqueId) {
      if (chairApprovalStatus === "Approve") {
        statusSheet.getRange(i + 1, 2).setValue("HR");
        sendHRApprovalEmail(requesterName, volunteerName, sourceRow[16], sourceRow[17], uniqueId);
      } else {
        rejectionReason = rejectionReason.trim();
        statusSheet.getRange(i + 1, 2).setValue("Rejected");
        sendRejectionEmail(requesterEmail, requesterName, volunteerName, "Chair", rejectionReason);
      }
      break;
    }
  }
}


  // ----- 2. HR Form -----
  else if (sheetId === 1074532913) {
    var hrApprovalStatus = e.values[6]; // H column
    Logger.log("HR approval status: " + hrApprovalStatus);

    var uniqueId = e.values[e.values.length - 1];
    Logger.log("HR form Unique ID: " + uniqueId);

    var sourceRow = sourceSheet.getDataRange().getValues().find(row => row[23] === uniqueId);
    if (!sourceRow) {
      Logger.log("❌ No source row found for Unique ID: " + uniqueId);
      return;
    }

    var requesterEmail = sourceRow[1]; // Column B
    var requesterName = sourceRow[2];
    var volunteerName = sourceRow[6];

    var statusData = statusSheet.getRange(1, 1, statusSheet.getLastRow(), 1).getValues();
    for (var i = 0; i < statusData.length; i++) {
      if (statusData[i][0] === uniqueId) {
        if (hrApprovalStatus === "Approve") {
          statusSheet.getRange(i + 1, 2).setValue("OoR");
          sendOoRApprovalEmail(requesterName, volunteerName, "Office of Research", "vdradmin@ualberta.ca", uniqueId);
        } else {
          var rejectionReason = e.values[8] || "No comments provided."; // I column
          statusSheet.getRange(i + 1, 2).setValue("Rejected");
          sendRejectionEmail(requesterEmail, requesterName, volunteerName, "HR", rejectionReason);
        }
        break;
      }
    }
  }

  // ----- 3. OoR Form -----
  else if (sheetId === 876106890) {
    var oorApprovalStatus = e.values[9]; // J column
    Logger.log("OoR approval status: " + oorApprovalStatus);

    var uniqueId = e.values[e.values.length - 1].trim();
    Logger.log("OoR form Unique ID: " + uniqueId);

    var sourceRow = sourceSheet.getDataRange().getValues().find(row => row[23] === uniqueId);
    if (!sourceRow) {
      Logger.log("❌ Source row not found for Unique ID: " + uniqueId);
      return;
    }

    var requesterEmail = sourceRow[1]; // Column B
    var requesterName = sourceRow[2];
    var volunteerName = sourceRow[6];
    var chairEmail = sourceRow[5];
    var hrEmail = sourceRow[17];
    var oorEmail = "vdradmin@ualberta.ca";

    var statusData = statusSheet.getRange(2, 1, statusSheet.getLastRow() - 1, 2).getValues(); // Skip header
    for (var i = 0; i < statusData.length; i++) {
      var currentId = statusData[i][0] ? statusData[i][0].toString().trim() : "";
      if (currentId === uniqueId) {
        if (oorApprovalStatus === "Approved") {
          statusSheet.getRange(i + 2, 2).setValue("Approved"); // +2 to skip header
          Logger.log("✅ Status updated to Approved at row: " + (i + 2));
        } else {
          var rejectionReason = e.values[10] || "No comments provided."; // K column
          statusSheet.getRange(i + 2, 2).setValue("Rejected");
          Logger.log("❌ OoR rejected the request. Sending rejection email.");
          sendRejectionEmail(requesterEmail, requesterName, volunteerName, "Office of Research", rejectionReason);
          return;
        }
        break;
      }
    }

    var pdfFile = createFinalPDF(requesterName, volunteerName, uniqueId);
    if (!pdfFile) {
      Logger.log("❌ PDF creation failed.");
      return;
    }

    sendFinalApprovalEmail(requesterName, volunteerName, requesterEmail, chairEmail, hrEmail, oorEmail, pdfFile);
  }

  // ----- 4. Initial Form Submission -----
  else if (sheetId === 1447833263) {
    var uniqueId = generateUniqueId();
    Logger.log("Generated Unique ID: " + uniqueId);

    var sourceRowIndex = e.range.getRow();
    sourceSheet.getRange(sourceRowIndex, 24).setValue(uniqueId);

    statusSheet.appendRow([uniqueId, "Chair"]);
    SpreadsheetApp.flush();

    var sourceRow = sourceSheet.getRange(sourceRowIndex, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
    sendChairApprovalEmail(sourceRow[2], sourceRow[6], sourceRow[4], sourceRow[5], uniqueId);
  }
}



// Chair Approval Email
function sendChairApprovalEmail(requesterName, volunteerName, deptChairName, chairEmail, uniqueId) {
  if (!chairEmail || !chairEmail.includes("@") || !chairEmail.includes(".")) {
    Logger.log("ERROR: Invalid chair email detected: " + chairEmail + ". Skipping email.");
    return;
  }

  var baseApprovalFormUrl = "https://docs.google.com/forms/d/e/1FAIpQLScyVFTKA5CoUNiayt8ghkLUUU7uumrvhuXHReGbkPyPNcUL5Q/viewform";
  var preFilledApprovalLink = baseApprovalFormUrl +
    "?usp=pp_url" +
    "&entry.349323761=" + encodeURIComponent(requesterName) +
    "&entry.495112661=" + encodeURIComponent(volunteerName) +
    "&entry.809972930=" + encodeURIComponent(deptChairName) +
    "&entry.337630937=" + encodeURIComponent(uniqueId);

  var summaryHtml = buildEmailHtmlSummary(uniqueId);

  var subject = "Action Required: FoMD Volunteer Request for " + requesterName;
  var htmlBody = `
    <div style="font-family:sans-serif">
      Dear ${deptChairName},<br><br>
      A FoMD Volunteer Request has been submitted by <b>${requesterName}</b> that requires your approval.<br><br>
      <strong>Please review the submission summary below and click here to approve:</strong> <a href="${preFilledApprovalLink}">Chair Approval Form</a><br><br>
      ${summaryHtml}<br><br>
      This is an automated message. If you are unable to access the approval form or documentation, please contact the FoMD Office of Research (vdradmin@ualberta.ca).
    </div>
  `;

  MailApp.sendEmail({
    to: chairEmail,
    subject: subject,
    htmlBody: htmlBody
  });

  Logger.log("Chair Approval Email sent to: " + chairEmail);
}

function sendHRApprovalEmail(requesterName, volunteerName, hrName, hrEmail, uniqueId) {
    Logger.log("📩 Preparing HR approval email...");

    if (!hrEmail || !hrEmail.includes("@")) {
        Logger.log("ERROR: Invalid HR email: " + hrEmail);
        return;
    }

    var baseApprovalFormUrl = "https://docs.google.com/forms/d/e/1FAIpQLSdlnC4YaZrHqxulu6Qx3V3pZAytsP68OONmB1coUyALTcVsVw/viewform";
    var preFilledApprovalLink = baseApprovalFormUrl +
        "?usp=pp_url" +
        "&entry.1876530800=" + encodeURIComponent(requesterName) +
        "&entry.1887281903=" + encodeURIComponent(volunteerName) +
        "&entry.1774495000=" + encodeURIComponent(hrName) +
        "&entry.858445503=" + encodeURIComponent(uniqueId)

    var subject = "Action Required: HR Approval for Volunteer Request - " + volunteerName;

    var summaryHtml = buildEmailHtmlSummary(uniqueId);

    var htmlBody = `
        <div style="font-family:sans-serif">
            Dear ${hrName},<br><br>
            A FoMD Volunteer Request has been approved by a FoMD Department Chair and now requires your approval.<br><br>
            <strong>Please review the submission summary below and click here to approve:</strong> 
            <a href="${preFilledApprovalLink}">HR Approval Form</a><br><br>
            ${summaryHtml}<br>
            This is an automated message. If you are unable to access the approval form or documentation, please contact the FoMD Office of Research (vdradmin@ualberta.ca).
        </div>
    `;

    MailApp.sendEmail({
        to: hrEmail,
        subject: subject,
        htmlBody: htmlBody
    });

    Logger.log("✅ HR Approval Email sent to: " + hrEmail);
}

function sendOoRApprovalEmail(requesterName, volunteerName, oorName, oorEmail, uniqueId) {
  Logger.log("📩 Preparing Office of Research approval email...");

  if (!oorEmail || !oorEmail.includes("@")) {
    Logger.log("ERROR: Invalid OoR email: " + oorEmail);
    return;
  }

  var baseApprovalFormUrl = "https://docs.google.com/forms/d/e/1FAIpQLScX5PPBMboXUJZlOVERKUNYOWLWD089LEFDlufId91LrX7wKA/viewform";
  var preFilledApprovalLink = baseApprovalFormUrl +
    "?usp=pp_url" +
    "&entry.1928826297=" + encodeURIComponent(requesterName) +
    "&entry.198357457=" + encodeURIComponent(volunteerName) +
    "&entry.1874226895=" + encodeURIComponent(oorName) +
    "&entry.927231890=" + encodeURIComponent(uniqueId);

  var summaryHtml = buildEmailHtmlSummary(uniqueId);

  var subject = "Action Required: OoR Approval for Volunteer Request - " + volunteerName;
  var htmlBody = `
    <div style="font-family:sans-serif">
      Dear ${oorName},<br><br>
      A FoMD Volunteer Request for <strong>${volunteerName}</strong> has been approved by an FoMD Department Chair and HR Partner.<br><br>
      <strong>Please review the submission summary below and click here to provide final approval:</strong><br>
      <a href="${preFilledApprovalLink}">OoR Final Approval Form</a><br><br>
      ${summaryHtml}<br><br>
      This is an automated message. If you are unable to access the approval form or documentation, please contact the FoMD Office of Research (vdradmin@ualberta.ca).<br><br>

    </div>
  `;

  MailApp.sendEmail({
    to: oorEmail,
    subject: subject,
    htmlBody: htmlBody,
  });

  Logger.log("✅ OoR Approval Email sent to: " + oorEmail);
}



function sendFinalApprovalEmail(requesterName, volunteerName, requesterEmail, chairEmail, hrEmail, oorEmail, pdfFile) {
  if (!pdfFile) {
    Logger.log("❌ ERROR: No PDF file provided to sendFinalApprovalEmail.");
    return;
  }

  // Add the OoR Admin email address
  var oorAdminEmail = "vdradmin@ualberta.ca";

  // Collect all valid email recipients
  var allRecipients = [requesterEmail, chairEmail, hrEmail, oorEmail, oorAdminEmail]
    .filter(email => email && email.includes("@"))
    .join(",");

  var subject = "Completed: Volunteer Request for " + volunteerName;
  var body = "Dear all,\n\n" +
             "The volunteer request for " + volunteerName + " has received all required approvals.\n\n" +
             "Attached is the finalized PDF containing all submission details and approval history and will be archived by the Office of Research.\n\n" +
             "If you have any questions, please contact the Office of Research Administrator (vdradmin@ualberta.ca).\n\n";

  MailApp.sendEmail({
    to: allRecipients,
    subject: subject,
    body: body,
    attachments: [pdfFile.getAs(MimeType.PDF)]
  });

  Logger.log("✅ Final approval email sent to: " + allRecipients);
}


/**
 * Formats a timestamp to MM/DD/YYYY HH:MM:SS
 */
function formatTimestamp(timestamp) {
    if (!timestamp) return "";
    if (typeof timestamp === "string") return timestamp; // If already a string, return it as is.

    var date = new Date(timestamp);
    var month = date.getMonth() + 1; // Months are zero-based
    var day = date.getDate();
    var year = date.getFullYear();
    var hours = date.getHours();
    var minutes = date.getMinutes();
    var seconds = date.getSeconds();

    return month + "/" + day + "/" + year + " " + 
           (hours < 10 ? "0" + hours : hours) + ":" + 
           (minutes < 10 ? "0" + minutes : minutes) + ":" + 
           (seconds < 10 ? "0" + seconds : seconds);
}

function getApproverByColumn(sheet, uniqueId, nameColIndex) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol = headers.indexOf("Unique ID");
  for (var i = 1; i < data.length; i++) {
    if (data[i][idCol] === uniqueId) {
      return data[i][nameColIndex] || "N/A";
    }
  }
  return "N/A";
}

function getApprovalDateByColumn(sheet, uniqueId) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol = headers.indexOf("Unique ID");
  for (var i = 1; i < data.length; i++) {
    if (data[i][idCol] === uniqueId) {
      return data[i][0] ? Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") : "N/A";
    }
  }
  return "N/A";
}


function buildEmailHtmlSummary(form_id) {
  const ss = SpreadsheetApp.openById("1048SyHN-e4XujUkJOXj1j5rMQdZ4R4ZwIEYcLTK0OJk");
  const sheet = ss.getSheetById(1447833263); // Volunteer Form Responses
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  let matchedRow = null;
  for (let i = 1; i < data.length; i++) {
    const currentId = data[i][23]; // Column X (Unique ID)
    Logger.log(`Row ${i} - Comparing ${currentId} === ${form_id}`);
    if (currentId === form_id) {
      matchedRow = data[i];
      break;
    }
  }

  Logger.log("🔍 Searching for form_id: " + form_id);

  if (!matchedRow) {
    Logger.log("buildEmailHtmlSummary: No matching row found for form_id " + form_id);
    return "<p><em>No submission details found for this form.</em></p>";
  }

  let html = "<h3>Submission Summary</h3><table border='1' cellpadding='4' cellspacing='0'>";
  for (let i = 0; i < headers.length; i++) {
    let question = headers[i];
    let answer = matchedRow[i] ? matchedRow[i].toString() : "N/A";

    // Format DOB fields
    if (question.toLowerCase().includes("dob")) {
      const dateVal = new Date(answer);
      if (!isNaN(dateVal.getTime())) {
        answer = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "MM/dd/yyyy");
      }
    }

    // Make waiver/background links clickable
    if (question.toLowerCase().includes("waiver") || question.toLowerCase().includes("background")) {
      if (answer.startsWith("http")) {
        answer = `<a href="${answer}" target="_blank">View Document</a>`;
      }
    }

    html += `<tr><td><strong>${question}</strong></td><td>${answer}</td></tr>`;
  }

  html += "</table>";
  return html;
}

function sendRejectionEmail(toEmail, requesterName, volunteerName, rejectedBy, reason) {
  if (!reason || reason.trim().length < 3) {
    reason = "No comments provided.";
  }

  // Map short roles to full display names
  var displayNameMap = {
    "Chair": "the Department Chair",
    "HR": "the HR Partner",
    "Office of Research": "the Office of Research"
  };

  var displayRole = displayNameMap[rejectedBy] || rejectedBy;

  var subject = "Volunteer Request Rejected – " + volunteerName;

  var htmlBody = `
    <div style="font-family:Arial, sans-serif; font-size:14px; line-height:1.5;">
      <p>Dear ${requesterName},</p>

      <p>Your volunteer request for <strong>${volunteerName}</strong> has been reviewed and was not approved by <strong>${displayRole}</strong>.</p>

      <p><strong>Rejection Reason:</strong><br>
      ${reason.trim()}</p>

      <p>This is an automated message. If you have questions or wish to revise and resubmit, please contact the Office of Research at 
      <a href="mailto:vdradmin@ualberta.ca">vdradmin@ualberta.ca</a>.</p>

    </div>
  `;

  MailApp.sendEmail({
    to: toEmail,
    subject: subject,
    htmlBody: htmlBody
  });

  Logger.log("❌ Rejection email sent to: " + toEmail);
}

function createFinalPDF(requesterName, volunteerName, uniqueId) {
  Logger.log("📄 Starting PDF generation for: " + requesterName + " / " + volunteerName);

  var ss = SpreadsheetApp.openById("1048SyHN-e4XujUkJOXj1j5rMQdZ4R4ZwIEYcLTK0OJk");
  var sourceSheet = ss.getSheetById(1447833263);
  var chairSheet = ss.getSheetById(362146665);
  var hrSheet = ss.getSheetById(1074532913);
  var oorSheet = ss.getSheetById(876106890);

  var data = sourceSheet.getDataRange().getValues();
  var headers = data[0];
  var matchedRow = data.find(row => row[23] === uniqueId);
  if (!matchedRow) {
    Logger.log("❌ No row found for Unique ID: " + uniqueId);
    return null;
  }

  var doc = DocumentApp.create(`Volunteer Submission - ${volunteerName}`);
  var body = doc.getBody();

  try {
    var logoBlob = DriveApp.getFileById("1crAalSDlgI5zb2Ty_QrvbNWLOWdxsDok").getBlob();
    var logoParagraph = body.insertParagraph(0, "");
    var image = logoParagraph.appendInlineImage(logoBlob);
    image.setWidth(275);
    image.setHeight(30);
    logoParagraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    Logger.log("✅ Logo inserted");
  } catch (err) {
    Logger.log("⚠️ Failed to insert logo: " + err);
  }

  body.appendParagraph("Faculty of Medicine & Dentistry")
  .setFontSize(18)
  .setBold(true)
  .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph("Volunteer Form Submission")
  .setFontSize(14)
  .setBold(true)
  .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph(`Requester: ${requesterName}`)
    .setFontSize(10)
    .setBold(false);

  body.appendParagraph(`Volunteer: ${volunteerName}`)
    .setFontSize(10)
    .setBold(false);;

  body.appendParagraph(`Unique ID: ${uniqueId}`)
    .setFontSize(10)
    .setBold(false);;

  body.appendParagraph("\nSubmission Details:").setHeading(DocumentApp.ParagraphHeading.HEADING2);

  var table = body.appendTable();
  for (var i = 0; i < headers.length; i++) {
    var question = headers[i];
    var answer = matchedRow[i] ? matchedRow[i].toString() : "N/A";

    // ✅ Format DOB fields to MM/dd/yyyy
    if (question.toLowerCase().includes("dob")) {
      var dateVal = new Date(answer);
      if (!isNaN(dateVal.getTime())) {
        answer = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "MM/dd/yyyy");
      }
    }

    var row = table.appendTableRow();
    row.appendTableCell(question);

    if ((question.toLowerCase().includes("waiver") || question.toLowerCase().includes("background")) && answer.startsWith("http")) {
      var cell = row.appendTableCell();
      var linkParagraph = cell.appendParagraph("");
      var text = linkParagraph.appendText("View Document Here");
      text.setLinkUrl(answer);
      text.setForegroundColor("#0000FF");
      text.setUnderline(true);
    } else {
      row.appendTableCell(answer);
    }
  }

  body.appendParagraph("\nApprovals").setHeading(DocumentApp.ParagraphHeading.HEADING2);

  var approvalsTable = body.appendTable();
  var headerRow = approvalsTable.appendTableRow();
  headerRow.appendTableCell("Approver");
  headerRow.appendTableCell("Name");
  headerRow.appendTableCell("Timestamp");

  var roles = [
    ["Chair", chairSheet, 4],
    ["HR Partner", hrSheet, 3],
    ["Office of Research", oorSheet, 4]
  ];

  for (var i = 0; i < roles.length; i++) {
    var approverRow = approvalsTable.appendTableRow();
    approverRow.appendTableCell(roles[i][0]);
    approverRow.appendTableCell(getApproverByColumn(roles[i][1], uniqueId, roles[i][2]));
    approverRow.appendTableCell(getApprovalDateByColumn(roles[i][1], uniqueId));
  }

  var footer = doc.addFooter();
  var footerText = footer.appendParagraph("Confidential – For Internal Use Only | Faculty of Medicine & Dentistry - Office of Research");
  footerText.setFontSize(8).setForegroundColor("#666666").setAlignment(DocumentApp.HorizontalAlignment.CENTER);


  doc.saveAndClose();
  var pdfBlob = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
  Logger.log("✅ PDF Created Successfully");
  return pdfBlob;
}




