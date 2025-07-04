// === Constants ===
const SPREADSHEET_ID = "1nmEnqMrncho2ulbHqkXsYYt2rAz0Oro6LCbF3gklj9w";
const OOR_APPROVERS = ["gvilas@ualberta.ca", "mlemieux@ualberta.ca", "npannu@ualberta.ca", "underhil@ualberta.ca"];
const OOR_EMAIL = "vdradmin@ualberta.ca";

const SHEET_IDS = {
  requests: 1130224705,
  chair: 321002302,
  hr: 1075645015,
  oor: 108980097,
  status: 639691046, 
  log: 692656271
};

function doGet(e) {
  const tracking = e.parameter.tracking;
  const role = e.parameter.role;
  const action = e.parameter.action;

  if (e.parameter?.success === 'true') {
    // Serve different success pages depending on who submitted the form
    if (role === 'Chair' || role === 'HR' || role === 'OoR') {
      return HtmlService.createHtmlOutputFromFile("approver_success").setTitle("Approval Submitted");
    } else {
      return HtmlService.createHtmlOutputFromFile("success").setTitle("Submission Confirmed");
    }
  }

  if (action === 'getResubmissionLinkTool') {
    return HtmlService.createHtmlOutputFromFile('resubmissionSidebar')
      .setTitle('Generate Resubmission Link');
  }

  // === SANDBOX TEST BLOCK ===
  if (action === 'sandboxTest') {
    const sampleRequestData = [
      '', '', '', '',              // [0]-[3] Placeholder
      'John Requester',            // [4] Requester Name
      '', '', 
      'Jane Chair',                // [7] Chair Name
      'chair@example.com',         // [8] Chair Email
      'Alice Volunteer'            // [9] Volunteer Name
    ];
    
    const template = HtmlService.createTemplateFromFile("chairApprovalForm");
    template.requestData = sampleRequestData;
    template.uniqueId = "VR202405160001";

    return template.evaluate().setTitle("Chair Approval Form Sandbox Test");
  }

  // === Production Routes ===

  if (action === "getResubmissionForm" && tracking) {
    const template = HtmlService.createTemplateFromFile("resubmission_form");
    template.tracking = tracking;
    return template.evaluate().setTitle("Volunteer Request – Resubmission");
  }

  if (action === "getchairApprovalForm" && tracking) {
    const requestData = getRequestByTracking(tracking);
    if (requestData) {
      const template = HtmlService.createTemplateFromFile("chairApprovalForm");
      template.requestData = requestData.values;
      template.uniqueId = tracking;

      const submissionType = requestData.values[21]; // Column V
      const originalUid = requestData.values[20];    // Column U

      template.isResubmission = submissionType === "Resubmission";
      template.originalUid = originalUid;

      return template.evaluate().setTitle("Chair Approval Form");
    }
    return HtmlService.createHtmlOutput("No request found with the provided tracking number.");
  }

  if (action === "gethrApprovalForm" && tracking) {
    const requestData = getRequestByTracking(tracking);
    if (requestData) {
      const template = HtmlService.createTemplateFromFile("hrApprovalForm");
      template.requestData = requestData.values;
      template.uniqueId = tracking;

      const submissionType = requestData.values[21]; // Column V
      const originalUid = requestData.values[20];    // Column U

      template.isResubmission = submissionType === "Resubmission";
      template.originalUid = originalUid;

      return template.evaluate().setTitle("HR Approval Form");
    }
    return HtmlService.createHtmlOutput("No request found with the provided tracking number.");
  }


  if (action === "getoorApprovalForm" && tracking) {
    const requestData = getRequestByTracking(tracking);
    if (requestData) {
      const template = HtmlService.createTemplateFromFile("oorApprovalForm");
      template.requestData = requestData.values;
      template.headers = requestData.headers;
      template.uniqueId = tracking;

      const submissionType = requestData.values[21]; // Column V
      const originalUid = requestData.values[20];    // Column U

      template.isResubmission = submissionType === "Resubmission";
      template.originalUid = originalUid;

      return template.evaluate().setTitle("Office of Research Approval Form");
    }
    return HtmlService.createHtmlOutput("No request found with the provided tracking number.");
  }

  if (action === "getStatusChecker" && tracking) {
    const requestData = getRequestByTracking(tracking);
    if (requestData) {
      const template = HtmlService.createTemplateFromFile("status_checker");
      template.tracking = tracking;
      template.requestData = requestData.values;

      const submissionType = requestData.values[21]; // Column V
      const originalUid = requestData.values[20];    // Column U

      template.isResubmission = submissionType === "Resubmission";
      template.originalUid = originalUid;

      // ✅ New: Fetch OoR final comment
      const oorSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetById(SHEET_IDS.oor);
      const oorData = oorSheet.getDataRange().getValues();
      const oorHeaders = oorData[0];
      const oorRow = oorData.find(r => r[oorHeaders.indexOf("Unique ID")] === tracking);
      const finalComment = oorRow ? oorRow[10] : "No final comment provided.";

      template.oorFinalComment = finalComment;

      return template.evaluate().setTitle("Volunteer Request Status Tracker");
    }
  }


  if (action === "getSubmissionForm") {
    const template = HtmlService.createTemplateFromFile("submission_form");
    const tracking = e.parameter?.tracking || '';
    const isResubmission = !!(tracking && checkIfTrackingIsResubmission(tracking));
    template.tracking = tracking;
    template.isResubmission = isResubmission;
    return template.evaluate().setTitle("Volunteer Request Form");
  }

  // === Approver Generic View (without specific action) ===
  if (tracking && role) {
    const requestData = getRequestByTracking(tracking);
    if (requestData) {
      const template = HtmlService.createTemplateFromFile("approval");
      template.requestData = requestData;
      template.role = role;
      template.approver = getApproverDetails(role, requestData.values);
      return template.evaluate().setTitle("Volunteer Request Approval");
    }
    return HtmlService.createHtmlOutput("No request found with the provided tracking number.");
  }

  // === Fallback route ===
  const trackingParam = e.parameter?.tracking || '';
  const request = trackingParam ? getRequestByTracking(trackingParam) : null;

  const template = HtmlService.createTemplateFromFile("submission_form");
  template.tracking = request ? trackingParam : '';
  template.isResubmission = request ? checkIfTrackingIsResubmission(trackingParam) : false;
  return template.evaluate().setTitle("FoMD Volunteer Request");
}


function getApproversForForm() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetById(352748199); // Approvers Directory
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues(); // Role, Name, Email, Department

  const result = { Chair: [], HR: [], ADM: [] };

  data.forEach(([role, name, email, department]) => {
    if (role === "Chair") {
      result.Chair.push({ name, email, department });
    } else if (role === "HR") {
      result.HR.push({ name, email });
    } else if (role === "ADM") {
      result.ADM.push({ name, email, department });
    }
  });

  return result;
}

function submitForm(formData) {
  Logger.log("📦 Received tracking: " + formData.tracking);
  Logger.log("formData.isResubmission: %s", formData.isResubmission);
  Logger.log("Evaluated isResubmission: %s", String(formData.isResubmission).toLowerCase() === "true");
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetById(SHEET_IDS.requests);
  const statusSheet = ss.getSheetById(SHEET_IDS.status);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const isResubmission = String(formData.isResubmission).toLowerCase() === "true";
  const originalId = isResubmission ? String(formData.tracking).replace(/^"+|"+$/g, '') : "";
  const newUniqueId = generateUniqueId();
  const submissionType = isResubmission ? `Resubmission from ${originalId}` : "New Submission";

  const newRow = [
    new Date(),                              // A
    formData.consultedHR || '',              // B
    formData.hrName || '',                   // C
    formData.hrEmail || '',                  // D
    formData.requesterName || '',            // E
    formData.requesterEmail || '',           // F
    formData.unit || '',                     // G
    formData.departmentChairName || '',      // H
    formData.chairEmail || '',               // I
    formData.volunteerName || '',            // J
    formData.dob || '',                      // K
    formData.program || '',                  // L
    (formData.immigrationStatus || []).join(', '), // M
    formData.immigrationPermission || '',    // N
    formData.roleTitle || '',                // O
    formData.tasks || '',                    // P
    formData.justification || '',            // Q
    formData.startDate || '',                // R
    formData.endDate || '',                  // S
    formData.riskDetails || '',              // U
    newUniqueId,                             // Y (index 24)
    submissionType                           // Z (index 25) ← explicitly here
  ];

  sheet.appendRow(newRow);
  statusSheet.appendRow([newUniqueId, "Chair"]);

  const admInfo = getAdmByDepartment(formData.unit);

  sendChairApprovalEmail(
    formData.requesterName,
    formData.volunteerName,
    formData.departmentChairName,
    formData.chairEmail,
    newUniqueId,
    isResubmission
  );

  sendSubmissionConfirmationEmail(
    formData.requesterEmail,
    formData.requesterName,
    formData.volunteerName,
    newUniqueId,
    formData.departmentChairName,
    formData.chairEmail,
    formData.hrName,
    formData.hrEmail,
    admInfo.email,
    admInfo.name,
    isResubmission
  );

  return HtmlService.createHtmlOutputFromFile("success").getContent();
}

function checkDuplicateEarly(requesterName, volunteerName, dob) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetById(SHEET_IDS.requests);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const requesterIndex = headers.indexOf("Requester Name");
  const volunteerIndex = headers.indexOf("Volunteer Name");
  const dobIndex = headers.indexOf("DOB");

  Logger.log(`🧪 HEADERS: requesterIndex=${requesterIndex}, volunteerIndex=${volunteerIndex}, dobIndex=${dobIndex}`);
  Logger.log(`🟡 Submitted values: ${requesterName} / ${volunteerName} / ${dob}`);

  if (requesterIndex === -1 || volunteerIndex === -1 || dobIndex === -1) {
    Logger.log("❌ Column headers not found.");
    return false;
  }

  const submittedDob = Utilities.formatDate(new Date(dob), Session.getScriptTimeZone(), "yyyy-MM-dd");
  Logger.log(`📆 Submitted DOB (formatted): ${submittedDob}`);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowRequester = row[requesterIndex]?.toString().trim().toLowerCase();
    const rowVolunteer = row[volunteerIndex]?.toString().trim().toLowerCase();
    const rowDobRaw = row[dobIndex];

    let rowDobFormatted = "";
    if (rowDobRaw instanceof Date) {
      rowDobFormatted = Utilities.formatDate(rowDobRaw, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else {
      rowDobFormatted = Utilities.formatDate(new Date(rowDobRaw), Session.getScriptTimeZone(), "yyyy-MM-dd");
    }

    Logger.log(`🔍 Row ${i}: ${rowRequester} / ${rowVolunteer} / ${rowDobFormatted}`);

    if (
      rowRequester === requesterName.toLowerCase().trim() &&
      rowVolunteer === volunteerName.toLowerCase().trim() &&
      rowDobFormatted === submittedDob
    ) {
      Logger.log("✅ DUPLICATE DETECTED");
      return true;
    }
  }

  Logger.log("❌ No duplicate found.");
  return false;
}

// === Approval Submission ===
function submitChairApproval(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const chairSheet = ss.getSheetById(SHEET_IDS.chair);
    const requestSheet = ss.getSheetById(SHEET_IDS.requests);
    const statusSheet = ss.getSheetById(SHEET_IDS.status);

    // === Check for Duplicate Chair Approval ===
    const chairData = chairSheet.getDataRange().getValues();
    const chairHeaders = chairData[0];
    const chairUIDIndex = chairHeaders.indexOf("Unique ID");
    const alreadyApproved = chairData.some(row => row[chairUIDIndex] === data.uniqueId);
    if (alreadyApproved) {
      Logger.log(`Duplicate Chair Approval Attempt Detected for Unique ID: ${data.uniqueId}`);
      return 'duplicate';
    }

    // === Log Chair Approval ===
    chairSheet.appendRow([
      new Date().toISOString(),
      data.chairEmail,
      data.requesterName,
      data.volunteerName,
      data.chairName,
      data.approval,
      data.rejectionReason,
      data.uniqueId
    ]);

    // === Get matching request row ===
    const requestData = requestSheet.getDataRange().getValues();
    const requestHeaders = requestData[0];
    const idCol = requestHeaders.indexOf("Unique ID");
    const hrNameCol = requestHeaders.indexOf("Department HR Name");
    const hrEmailCol = requestHeaders.indexOf("Department HR Email");
    const departmentCol = requestHeaders.indexOf("Department");

    const matchingRow = requestData.find(row => row[idCol] === data.uniqueId);
    if (!matchingRow) {
      Logger.log(`submitChairApproval: No matching request found for Unique ID: ${data.uniqueId}`);
      return 'error: matching request not found';
    }

    const hrName = matchingRow[hrNameCol];
    const hrEmail = matchingRow[hrEmailCol];
    const department = matchingRow[departmentCol];
    const admInfo = getAdmByDepartment(department);

    // === Rejection Path ===
    if (data.approval === "Reject") {
      Logger.log(`Chair rejected the request. Notifying ADM: ${admInfo.email}`);

      sendChairRejectionEmail({
        requesterEmail: matchingRow[requestHeaders.indexOf("Requester Email")],
        requesterName: data.requesterName,
        volunteerName: data.volunteerName,
        rejectionReason: data.rejectionReason,
        uniqueId: data.uniqueId,
        admEmail: admInfo.email,
        admName: admInfo.name
      });

      // === Update Status Sheet to "Rejected" ===
      const statusData = statusSheet.getDataRange().getValues();
      const statusHeaders = statusData[0];
      const statusIdCol = statusHeaders.indexOf("Unique ID");
      const statusCol = statusHeaders.indexOf("Status");

      let statusUpdated = false;
      for (let i = 1; i < statusData.length; i++) {
        if (statusData[i][statusIdCol] === data.uniqueId) {
          statusSheet.getRange(i + 1, statusCol + 1).setValue("Rejected");
          statusUpdated = true;
          break;
        }
      }

      if (!statusUpdated) {
        Logger.log(`Chair rejection: No status row found for Unique ID: ${data.uniqueId}. Adding new "Rejected" row.`);
        statusSheet.appendRow([data.uniqueId, "Rejected"]);
      }

      return 'rejected';
    }

    // === Proceed to HR stage ===
    const statusData = statusSheet.getDataRange().getValues();
    const statusHeaders = statusData[0];
    const statusIdCol = statusHeaders.indexOf("Unique ID");
    const statusCol = statusHeaders.indexOf("Status");

    let found = false;
    for (let i = 1; i < statusData.length; i++) {
      if (statusData[i][statusIdCol] === data.uniqueId) {
        statusSheet.getRange(i + 1, statusCol + 1).setValue("HR");
        found = true;
        break;
      }
    }

    if (!found) {
      Logger.log(`submitChairApproval: No status row found for Unique ID: ${data.uniqueId}. Adding new row.`);
      statusSheet.appendRow([data.uniqueId, "HR"]);
    }

    if (!hrEmail || !hrEmail.includes('@')) {
      Logger.log(`submitChairApproval: Invalid HR email for Unique ID: ${data.uniqueId}`);
      return 'error: invalid HR email';
    }

    const submissionType = matchingRow[requestHeaders.indexOf("Submission Type")];
    const isResubmission = (submissionType || "").toLowerCase().startsWith("resubmission");
    sendHRApprovalEmail(data.requesterName, data.volunteerName, hrName, hrEmail, data.uniqueId, isResubmission);

      // Notify ADM that request has moved to HR
    if (admInfo.email) {
      MailApp.sendEmail({
        to: admInfo.email,
        subject: `Volunteer Request Moved to HR Stage – ${data.requesterName}`,
        htmlBody: `
          <p>Dear ${admInfo.name},</p>
          <p>The volunteer request for <strong>${data.volunteerName}</strong> submitted by <strong>${data.requesterName}</strong> has been approved by the Department Chair and is now pending HR approval.</p>
          <p><strong>Tracking ID:</strong> ${data.uniqueId}</p>
          <p>You can monitor its progress in the <a href="https://script.google.com/a/macros/ualberta.ca/s/AKfycbwXyUlQDD1Wts1BgQ5o2rWaAySpgn5Ps2Fzf7VjTmgFN1kMD5VQWmB_gFyfew6S2OrudA/exec?action=getStatusChecker" target="_blank">status tracker</a>.</p>
          <p>– FoMD Volunteer Management System</p>
        `
      });

      Logger.log(`📨 ADM notified: request moved to HR for UID ${data.uniqueId}`);
    }

    return 'success';

  } catch (err) {
    Logger.log(`submitChairApproval: Exception - ${err}`);
    return `error: ${err.message}`;
  }
}

function sendChairRejectionEmail({ requesterEmail, requesterName, volunteerName, rejectionReason, uniqueId, admEmail, admName }) {
  if (!requesterEmail || !requesterEmail.includes('@')) {
    Logger.log("❌ Invalid requester email, cannot send Chair rejection email.");
    return;
  }

  const subject = `Volunteer Request Rejected by Chair: ${volunteerName}`;
  const htmlBody = `
    <div style="font-family:sans-serif; font-size:14px;">
      <p>Dear ${requesterName},</p>

      <p>Your volunteer request for <strong>${volunteerName}</strong> has been <strong>rejected</strong> by your Department Chair.</p>

      <p><strong>Rejection Reason:</strong><br>
      ${rejectionReason || "No reason provided."}</p>

      <p><strong>Tracking ID:</strong> ${uniqueId}</p>

      <p>You can monitor its progress in the <a href="https://script.google.com/a/macros/ualberta.ca/s/AKfycbwXyUlQDD1Wts1BgQ5o2rWaAySpgn5Ps2Fzf7VjTmgFN1kMD5VQWmB_gFyfew6S2OrudA/exec?action=getStatusChecker" target="_blank">status tracker</a>.</p>

      <p>If you have questions or have been instructed to resubmit the request, please contact the FoMD Office of Research at 
      <a href="mailto:vdradmin@ualberta.ca">vdradmin@ualberta.ca</a>.</p>

      <p style="font-size:12px; color:#777;">This message was also copied to your department ADM for awareness.</p>
    </div>
  `;

  MailApp.sendEmail({
    to: requesterEmail,
    cc: admEmail || "",
    subject: subject,
    htmlBody: htmlBody
  });

  Logger.log(`📨 Chair rejection email sent to ${requesterEmail}, CC: ${admEmail}`);
}

function submitHRApproval(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hrSheet = ss.getSheetById(SHEET_IDS.hr);
  const requestSheet = ss.getSheetById(SHEET_IDS.requests);
  const statusSheet = ss.getSheetById(SHEET_IDS.status);
  const hrData = hrSheet.getDataRange().getValues();
  const hrHeaders = hrData[0];
  const hrUIDIndex = hrHeaders.indexOf("Unique ID");

  const alreadyApproved = hrData.some(row => row[hrUIDIndex] === data.uniqueId);
  if (alreadyApproved) {
    Logger.log(`Duplicate HR Approval Attempt Detected for Unique ID: ${data.uniqueId}`);
    return 'duplicate';
  }

  // === Log HR Approval ===
  hrSheet.appendRow([
    new Date().toISOString(),
    data.hrEmail,
    data.requesterName,
    data.volunteerName,
    data.hrName,
    data.compliance,
    data.approval,
    data.hrComment,
    data.uniqueId
  ]);

  // === Get matching request row ===
  const requestData = requestSheet.getDataRange().getValues();
  const requestHeaders = requestData[0];
  const idCol = requestHeaders.indexOf("Unique ID");
  const requesterEmailCol = requestHeaders.indexOf("Requester Email");
  const departmentCol = requestHeaders.indexOf("Department");

  const matchingRow = requestData.find(row => row[idCol] === data.uniqueId);
  if (!matchingRow) {
    Logger.log(`submitHRApproval: No matching request found for Unique ID: ${data.uniqueId}`);
    return 'error: matching request not found';
  }

  const requesterEmail = matchingRow[requesterEmailCol];
  const department = matchingRow[departmentCol];
  const admInfo = getAdmByDepartment(department);


  // === If rejected, update status and send rejection email ===
  if (data.approval === "Reject") {
    const statusData = statusSheet.getDataRange().getValues();
    const statusHeaders = statusData[0];
    const statusIdCol = statusHeaders.indexOf("Unique ID");
    const statusCol = statusHeaders.indexOf("Status");

    let found = false;
    for (let i = 1; i < statusData.length; i++) {
      if (statusData[i][statusIdCol] === data.uniqueId) {
        statusSheet.getRange(i + 1, statusCol + 1).setValue("Rejected");
        found = true;
        break;
      }
    }

    if (!found) {
      Logger.log(`submitHRApproval: No status row found for Unique ID: ${data.uniqueId}. Adding new "Rejected" row.`);
      statusSheet.appendRow([data.uniqueId, "Rejected"]);
    }

    sendHRRejectionEmail({
      requesterEmail: requesterEmail,
      requesterName: data.requesterName,
      volunteerName: data.volunteerName,
      rejectionReason: data.hrComment,
      uniqueId: data.uniqueId,
      admEmail: admInfo.email,
      admName: admInfo.name
    });

    return 'rejected';
  }

  // === If approved, move to OoR ===
  const statusData = statusSheet.getDataRange().getValues();
  const statusHeaders = statusData[0];
  const statusIdCol = statusHeaders.indexOf("Unique ID");
  const statusCol = statusHeaders.indexOf("Status");

  let found = false;
  for (let i = 1; i < statusData.length; i++) {
    if (statusData[i][statusIdCol] === data.uniqueId) {
      statusSheet.getRange(i + 1, statusCol + 1).setValue("OoR");
      found = true;
      break;
    }
  }

  if (!found) {
    Logger.log(`submitHRApproval: No status row found for Unique ID: ${data.uniqueId}. Adding new row.`);
    statusSheet.appendRow([data.uniqueId, "OoR"]);
  }

  const submissionType = matchingRow[requestHeaders.indexOf("Submission Type")];
  const isResubmission = (submissionType || "").toLowerCase().startsWith("resubmission");
  sendOoRApprovalEmail(data.requesterName, data.volunteerName, "Office of Research", OOR_EMAIL, data.uniqueId, isResubmission);


    // Notify ADM that request has moved to OoR
  if (admInfo.email) {
    MailApp.sendEmail({
      to: admInfo.email,
      subject: `Volunteer Request Moved to Office of Research – ${data.requesterName}`,
      htmlBody: `
        <p>Dear ${admInfo.name},</p>
        <p>The volunteer request for <strong>${data.volunteerName}</strong> submitted by <strong>${data.requesterName}</strong> has been approved by HR and is now pending approval from the Office of Research.</p>
        <p><strong>Tracking ID:</strong> ${data.uniqueId}</p>
        <p>You can monitor its progress in the <a href="https://script.google.com/a/macros/ualberta.ca/s/AKfycbwXyUlQDD1Wts1BgQ5o2rWaAySpgn5Ps2Fzf7VjTmgFN1kMD5VQWmB_gFyfew6S2OrudA/exec?action=getStatusChecker" target="_blank">status tracker</a>.</p>
        <p>– FoMD Volunteer Management System</p>
      `
    });

    Logger.log(`📨 ADM notified: request moved to OoR for UID ${data.uniqueId}`);
  }

  return 'success';
}

function sendHRRejectionEmail({ requesterEmail, requesterName, volunteerName, rejectionReason, uniqueId, admEmail, admName }) {
  if (!requesterEmail || !requesterEmail.includes('@')) {
    Logger.log("❌ Invalid requester email, cannot send HR rejection email.");
    return;
  }

  const subject = `Volunteer Request Rejected by HR: ${volunteerName}`;
  const htmlBody = `
    <div style="font-family:sans-serif; font-size:14px;">
      <p>Dear ${requesterName},</p>

      <p>Your volunteer request for <strong>${volunteerName}</strong> has been <strong>rejected</strong> by the Department HR partner.</p>

      <p><strong>Rejection Reason:</strong><br>
      ${rejectionReason || "No reason provided."}</p>

      <p><strong>Tracking ID:</strong> ${uniqueId}</p>

      <p>You can monitor its progress in the <a href="https://script.google.com/a/macros/ualberta.ca/s/AKfycbwXyUlQDD1Wts1BgQ5o2rWaAySpgn5Ps2Fzf7VjTmgFN1kMD5VQWmB_gFyfew6S2OrudA/exec?action=getStatusChecker" target="_blank">status tracker</a>.</p>

      <p>If you have questions or have been instructed to resubmit the request, please contact the FoMD Office of Research at 
      <a href="mailto:vdradmin@ualberta.ca">vdradmin@ualberta.ca</a>.</p>

      <p style="font-size:12px; color:#777;">This message was also copied to your department ADM for awareness.</p>
    </div>
  `;

  MailApp.sendEmail({
    to: requesterEmail,
    cc: admEmail || "",
    subject: subject,
    htmlBody: htmlBody
  });

  Logger.log(`📨 HR rejection email sent to ${requesterEmail}, CC: ${admEmail}`);
}

function submitOoRApproval(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const oorSheet = ss.getSheetById(SHEET_IDS.oor);
  const requestSheet = ss.getSheetById(SHEET_IDS.requests);
  const statusSheet = ss.getSheetById(SHEET_IDS.status);

  const oorData = oorSheet.getDataRange().getValues();
  const oorHeaders = oorData[0];
  const oorUIDIndex = oorHeaders.indexOf("Unique ID");

  const alreadyApproved = oorData.some(row => row[oorUIDIndex] === data.uniqueId);
  if (alreadyApproved) {
    Logger.log(`Duplicate OoR Approval Attempt Detected for Unique ID: ${data.uniqueId}`);
    return 'duplicate';
  }

  // 1. Log OoR Approval
  const oorRow = new Array(12).fill("");
  oorRow[0] = new Date().toISOString();          // A. Timestamp
  oorRow[1] = data.oorEmail;
  oorRow[2] = data.requesterName;
  oorRow[3] = data.volunteerName;
  oorRow[4] = data.oorName;
  oorRow[5] = data.alignment || "";
  oorRow[6] = data.alignmentComments || "";
  oorRow[7] = data.compliance || "";
  oorRow[8] = data.complianceComments || "";
  oorRow[9] = data.approval;
  oorRow[10] = data.rejectionReason || "";
  oorRow[11] = data.uniqueId;

  oorSheet.appendRow(oorRow);

  // 2. Handle rejection case
  if (data.approval === "Reject") {
    const statusData = statusSheet.getDataRange().getValues();
    const statusHeaders = statusData[0];
    const statusIdCol = statusHeaders.indexOf("Unique ID");
    const statusCol = statusHeaders.indexOf("Status");

    let found = false;
    for (let i = 1; i < statusData.length; i++) {
      if (statusData[i][statusIdCol] === data.uniqueId) {
        statusSheet.getRange(i + 1, statusCol + 1).setValue("Rejected");
        found = true;
        break;
      }
    }

    if (!found) {
      Logger.log(`submitOoRApproval: No status row found for Unique ID: ${data.uniqueId}. Adding new "Rejected" row.`);
      statusSheet.appendRow([data.uniqueId, "Rejected"]);
    }

    // Get requester + ADM info
    const requestData = requestSheet.getDataRange().getValues();
    const headers = requestData[0];
    const matchingRow = requestData.find(row => row[headers.indexOf("Unique ID")] === data.uniqueId);

    if (matchingRow) {
      const requesterName = matchingRow[4];
      const requesterEmail = matchingRow[5];
      const volunteerName = matchingRow[9];
      const department = matchingRow[headers.indexOf("Department")];
      const admInfo = getAdmByDepartment(department);

      sendOoRRejectionEmail({
        requesterEmail,
        requesterName,
        volunteerName,
        rejectionReason: data.rejectionReason,
        uniqueId: data.uniqueId,
        admEmail: admInfo.email,
        admName: admInfo.name
      });
    } else {
      Logger.log(`submitOoRApproval: No matching request row found for Unique ID: ${data.uniqueId}`);
    }

    return 'rejected';
  }

  // 3. If approved, update status to Complete
  const statusData = statusSheet.getDataRange().getValues();
  const statusHeaders = statusData[0];
  const statusIdCol = statusHeaders.indexOf("Unique ID");
  const statusCol = statusHeaders.indexOf("Status");

  let found = false;
  for (let i = 1; i < statusData.length; i++) {
    if (statusData[i][statusIdCol] === data.uniqueId) {
      statusSheet.getRange(i + 1, statusCol + 1).setValue("Complete");
      found = true;
      break;
    }
  }

  if (!found) {
    Logger.log(`submitOoRApproval: No status row found for Unique ID: ${data.uniqueId}. Adding new row.`);
    statusSheet.appendRow([data.uniqueId, "Complete"]);
  }

  // 4. Send Final Approval Email with PDF
  const matchingRow = requestSheet.getDataRange().getValues().find(row => row[requestSheet.getDataRange().getValues()[0].indexOf("Unique ID")] === data.uniqueId);
  if (matchingRow) {
    const requesterName = matchingRow[4];
    const volunteerName = matchingRow[9];
    const requesterEmail = matchingRow[5];
    const chairEmail = matchingRow[8];
    const hrEmail = matchingRow[3];

    sendFinalApprovalEmail(requesterName, volunteerName, requesterEmail, chairEmail, hrEmail, data.oorEmail, data.uniqueId, true);
  }

  return 'success';
}

function sendOoRRejectionEmail({ requesterEmail, requesterName, volunteerName, rejectionReason, uniqueId, admEmail, admName }) {
  if (!requesterEmail || !requesterEmail.includes('@')) {
    Logger.log("❌ Invalid requester email, cannot send OoR rejection email.");
    return;
  }

  const subject = `Volunteer Request Rejected by Office of Research: ${volunteerName}`;
  const htmlBody = `
    <div style="font-family:sans-serif; font-size:14px;">
      <p>Dear ${requesterName},</p>

      <p>Your volunteer request for <strong>${volunteerName}</strong> has been <strong>rejected</strong> by the Office of Research.</p>

      <p><strong>Rejection Reason:</strong><br>
      ${rejectionReason || "No reason provided."}</p>

      <p><strong>Tracking ID:</strong> ${uniqueId}</p>

      <p>You can monitor its progress in the <a href="https://script.google.com/a/macros/ualberta.ca/s/AKfycbwXyUlQDD1Wts1BgQ5o2rWaAySpgn5Ps2Fzf7VjTmgFN1kMD5VQWmB_gFyfew6S2OrudA/exec?action=getStatusChecker" target="_blank">status tracker</a>.</p>

      <p>If you have questions or have been instructed to resubmit the request, please contact the FoMD Office of Research at 
      <a href="mailto:vdradmin@ualberta.ca">vdradmin@ualberta.ca</a>.</p>

      <p style="font-size:12px; color:#777;">This message was also copied to your department ADM for awareness.</p>
    </div>
  `;

  MailApp.sendEmail({
    to: requesterEmail,
    cc: admEmail || "",
    subject: subject,
    htmlBody: htmlBody
  });

  Logger.log(`📨 OoR rejection email sent to ${requesterEmail}, CC: ${admEmail}`);
}

function getRequestByTracking(tracking) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets().find(s => s.getSheetId() === SHEET_IDS.requests);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const uidIndex = headers.indexOf("Unique ID");
  const typeIndex = headers.indexOf("Submission Type"); // Column V

  const cleanedTracking = String(tracking).trim().replace(/^"+|"+$/g, '');
  Logger.log("🔍 Looking for tracking: '" + cleanedTracking + "'");

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const sheetValue = String(row[uidIndex]).trim().replace(/^"+|"+$/g, '');
    if (sheetValue === cleanedTracking) {
      Logger.log("✅ Match found at row " + (i + 1));
      
      let isResubmission = false;
      let originalUid = '';
      const submissionType = (row[typeIndex] || '').toString().trim();
      
      if (submissionType.toLowerCase().startsWith('resubmission from')) {
        isResubmission = true;
        const match = submissionType.match(/VR\d+/);
        if (match) originalUid = match[0];
      }

      return {
        rowIndex: i + 1,
        headers: headers,
        values: row,
        isResubmission: isResubmission,
        originalUid: originalUid
      };
    }
  }

  return null;
}



// === Email Senders ===
function sendSubmissionConfirmationEmail(
  requesterEmail,
  requesterName,
  volunteerName,
  uniqueId,
  chairName,
  chairEmail,
  hrName,
  hrEmail,
  admEmail,
  admName,
  isResubmission // <-- new parameter
) {
  const subject = `${isResubmission ? '[RESUBMISSION] ' : ''}Confirmation: Volunteer Request Submitted for ${volunteerName}`;
  Logger.log("📨 [DEBUG] Final subject: " + subject);
  Logger.log("📨 [DEBUG] isResubmission received by email function: " + isResubmission);
  Logger.log("📨 [DEBUG] Typeof isResubmission: " + typeof isResubmission);

  const summaryTable = buildEmailHtmlSummary(uniqueId);

  const htmlBody = `
    <div style="font-family:sans-serif; font-size:14px;">
      <p>Dear ${requesterName},</p>

      <p>Thank you for submitting a volunteer request for <strong>${requesterName}</strong>.</p>

      <p>Your submission has been received and is pending approval from the following individuals:</p>

      <ul>
        <li><strong>Department Chair:</strong> ${chairName} (${chairEmail})</li>
        <li><strong>Department HR:</strong> ${stripParentheses(hrName)} (${hrEmail})</li>
        <li><strong>Office of Research Approvers:</strong> Gonzalo Vilas (gvilas@ualberta.ca), Joanne Lemieux (mlemieux@ualberta.ca), Neesh Pannu (npannu@ualberta.ca), Alan Underhill (underhil@ualberta.ca)</li>
      </ul>

      <p>You may check the current status of your request at any time using the 
      <a href="https://script.google.com/a/macros/ualberta.ca/s/AKfycbwXyUlQDD1Wts1BgQ5o2rWaAySpgn5Ps2Fzf7VjTmgFN1kMD5VQWmB_gFyfew6S2OrudA/exec" target="_blank">Volunteer Request Status Tracker</a>.</p>

      <p><strong>Unique ID:</strong> <code>${uniqueId}</code></p>

      ${admName ? `<p>This confirmation has also been sent to your department ADM, <strong>${admName}</strong>, and to the Office of Research Administrator at <strong>vdradmin@ualberta.ca</strong>.</p>` : `<p>This confirmation has also been sent to the Office of Research Administrator at <strong>vdradmin@ualberta.ca</strong>.</p>`}

      <p><strong>Request Summary:</strong></p>
      ${summaryTable}

      <p style="margin-top:20px; font-size:12px;">
        This is an automated message. If you have submitted this in error or wish to amend the information you have provided, please contact the FoMD Office of Research at 
        <a href="mailto:vdradmin@ualberta.ca">vdradmin@ualberta.ca</a>.
      </p>
    </div>
  `;

  MailApp.sendEmail({
    to: requesterEmail,
    cc: admEmail || "",
    subject: subject,
    htmlBody: htmlBody
  });

  Logger.log(`Confirmation email sent to: ${requesterEmail} (cc: ${admEmail})`);
}


function sendChairApprovalEmail(requesterName, volunteerName, deptChairName, chairEmail, uniqueId, isResubmission) {
  if (!chairEmail || !chairEmail.includes("@")) return;

  const subjectPrefix = isResubmission ? "[RESUBMISSION] " : "";
  const subject = `${subjectPrefix}Volunteer Request Approval Required for ${requesterName}`;

  const url = ScriptApp.getService().getUrl();
  const link = `${url}?action=getchairApprovalForm&tracking=${uniqueId}&role=Chair`;
  const summaryHtml = buildEmailHtmlSummary(uniqueId);

  const body = `
    <div style="font-family:sans-serif;">
      <p>Dear ${deptChairName},</p>

      <p>
        A volunteer request has been submitted by <strong>${requesterName}</strong> 
        for <strong>${volunteerName}</strong> and requires your approval.
      </p>

      <p>
        Please review the request summary and any attached documentation and click here to approve: 
        <a href="${link}">Chair Approval Form</a>
      </p>

      <p>
        Note that the chair approval form must be completed to move the request forward to the next approver. 
        Replying 'Approve' to this communication is insufficient.
      </p>

      <p><strong>Volunteer Request Summary:</strong><br>${summaryHtml}</p>

      <p style="font-size:12px; margin-top:20px;">
        This is an automated message. If you are unable to access the approval form or documentation, 
        please contact the FoMD Office of Research at 
        <a href="mailto:vdradmin@ualberta.ca">vdradmin@ualberta.ca</a>.
      </p>
    </div>
  `;

  MailApp.sendEmail({ to: chairEmail, subject: subject, htmlBody: body }); 
}

function sendHRApprovalEmail(requesterName, volunteerName, hrName, hrEmail, uniqueId, isResubmission) {
  Logger.log(`📨 Attempting to send HR Approval Email:
    Requester: ${requesterName},
    Volunteer: ${volunteerName},
    HR Name: ${hrName},
    HR Email: ${hrEmail},
    Unique ID: ${uniqueId},
    Resubmission: ${isResubmission}`);

  if (!hrEmail || !hrEmail.includes("@")) {
    Logger.log(`Invalid or missing HR Email for Unique ID: ${uniqueId}. Skipping email send.`);
    return;
  }

  const subjectPrefix = isResubmission ? "[RESUBMISSION] " : "";
  const subject = `${subjectPrefix}HR Approval Needed: ${requesterName}`;
  const url = ScriptApp.getService().getUrl();
  const link = `${url}?action=gethrApprovalForm&tracking=${uniqueId}&role=HR`;
  const summaryHtml = buildEmailHtmlSummary(uniqueId);

  const body = `
<div style="font-family:sans-serif;">
  <p>Dear ${stripParentheses(hrName)},</p>

  <p>
    A volunteer request has been submitted by <strong>${requesterName}</strong> 
    for <strong>${volunteerName}</strong>, has been approved by your department Chair, and now requires your approval.
  </p>

  <p>
    Please review the request summary and any attached documentation and click here to approve: 
    <a href="${link}">HR Approval Form</a>
  </p>

  <p>
    Note that the HR approval form must be completed to move the request forward to the next approver. 
    Replying 'Approve' to this communication is insufficient.
  </p>

  <p>This request has also been provided to your department ADM.</p>

  <p><strong>Volunteer Request Summary:</strong><br>${summaryHtml}</p>

  <div style="font-size:12px; margin-top:20px; color:#555;">
    <p>
      This is an automated message. If you are unable to access the approval form or documentation, 
      please contact the FoMD Office of Research at 
      <a href="mailto:vdradmin@ualberta.ca">vdradmin@ualberta.ca</a>.
    </p>
  </div>
</div>
`;

  Logger.log("HTML Body:\n" + body);

  MailApp.sendEmail({
    to: hrEmail,
    subject: subject,
    htmlBody: body
  });

  Logger.log(`HR Approval Email sent to: ${hrEmail} for Volunteer: ${volunteerName}`);
}

function sendOoRApprovalEmail(requesterName, volunteerName, oorName, oorEmail, uniqueId, isResubmission) {
  const subjectPrefix = isResubmission ? "[RESUBMISSION] " : "";
  const subject = `${subjectPrefix}OoR Approval Needed: ${requesterName}`;
  const url = ScriptApp.getService().getUrl();
  const link = `${url}?action=getoorApprovalForm&tracking=${uniqueId}&role=Office%20of%20Research`;
  const summaryHtml = buildEmailHtmlSummary(uniqueId);

  const body = `
<div style="font-family:sans-serif;">
  <p>Dear ${oorName},</p>

  <p>
    A volunteer request has been submitted by <strong>${requesterName}</strong> 
    for <strong>${volunteerName}</strong> and has been approved by their department Chair and HR service partner.
    This request will be included in the batch report for discussion and approval at the next VDR meeting.
  </p>

  <p>
    Please review the request summary and any attached documentation and click here to approve:  
    <a href="${link}">OoR Final Approval Form</a>
  </p>

  <p>
    Note that the final approval form must be completed to finalize this request.  
    Replying 'Approve' to this communication is insufficient.
  </p>

  <p><strong>Volunteer Request Summary:</strong></p>
  ${summaryHtml}

  <div style="font-size:12px; margin-top:20px; color:#555;">
    <p>
      This is an automated message. If you are unable to access the approval form or documentation,  
      please contact the FoMD Office of Research at  
      <a href="mailto:vdradmin@ualberta.ca">vdradmin@ualberta.ca</a>.
    </p>
  </div>
</div>
`;

  MailApp.sendEmail({
    to: oorEmail,
    cc: OOR_EMAIL,
    subject: subject,
    htmlBody: body
  });
}


function sendFinalApprovalEmail(requesterName, volunteerName, requesterEmail, chairEmail, hrEmail, oorEmail, uniqueId) {
  const pdfFile = createFinalPDF(requesterName, volunteerName, uniqueId);
  if (!pdfFile) {
    Logger.log("❌ Could not generate PDF for final approval.");
    return;
  }

  const toRecipients = [requesterEmail, chairEmail, hrEmail, oorEmail]
    .filter(email => email && email.includes("@"))
    .join(",");

  const department = getValueByUniqueId(uniqueId, "Department");
  const admInfo = getAdmByDepartment(department);
  const admEmail = admInfo.email;

  const ccRecipients = [
    "boisvenu@ualberta.ca",
    "vdradmin@ualberta.ca",
    "underhil@ualberta.ca",
    "nkosturi@ualberta.ca",
    "crilov@ualberta.ca",
    "shughes1@ualberta.ca",
    "mlemieux@ualberta.ca",
    "gvilas@ualberta.ca",
    "daerendi@ualberta.ca",
    "npannu@ualberta.ca"
  ]
  .concat(admEmail && admEmail.includes("@") ? [admEmail] : [])
  .join(",");

  const subject = `Final Approval Complete: Volunteer Request for ${requesterName}`;

  const htmlBody = `
    <div style="font-family:sans-serif; font-size:14px;">
      <p>Dear all,</p>

      <p>The volunteer request submitted by <strong>${requesterName}</strong> 
    for <strong>${volunteerName}</strong> has received all required approvals. A finalized copy of the approval with approver comments is attached for your reference and has also been provided to the corresponding department ADM.</p>

      <p><strong>Next Steps for the Requester:</strong></p>
      <ul>
        <li>Follow sections D + E in the <a href="https://www.ualberta.ca/en/medicine/media-library/research/research-resources/fomd-volunteer-procedure-april-9-2025.pdf.pdf" target="_blank"><strong>FoMD Volunteer Procedure</strong></a> to complete the remaining institutional level steps required after this approval.</li>
        <li>Retain the attached PDF for your records and onboarding.</li>
      </ul>

      <p>If you have any questions, please contact the Office of Research at 
        <a href="mailto:vdradmin@ualberta.ca">vdradmin@ualberta.ca</a>.</p>

      <p style="font-size:12px; color:#555; margin-top:20px;">
        This is an automated message from the FoMD Volunteer Management System.
      </p>
    </div>
  `;

  MailApp.sendEmail({
    to: toRecipients,
    cc: ccRecipients,
    subject: subject,
    htmlBody: htmlBody,
    attachments: [pdfFile]
  });

  Logger.log(`Final approval email sent to ${toRecipients} (cc: ${ccRecipients})`);
}

// === Helpers ===
function generateUniqueId() {
  const now = new Date();
  return 'VR' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHHmmssSSS');
}

function getValueByUniqueId(uniqueId, fieldLabel) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetById(SHEET_IDS.requests);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const index = headers.indexOf(fieldLabel);
  if (index === -1) return "";
  const row = data.find(r => r[headers.indexOf("Unique ID")] === uniqueId);
  return row ? row[index] : "";
}

function buildEmailHtmlSummary(trackingId) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets().find(s => s.getSheetId() === SHEET_IDS.requests);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const uidIndex = headers.indexOf("Unique ID");

  if (uidIndex === -1) return "<p>Unique ID column not found.</p>";

  const row = data.find(r => String(r[uidIndex]) === trackingId);
  if (!row) return "<p>No data found.</p>";

  let html = "<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse; font-family:sans-serif; font-size:13px;'>";

  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    let value = row[i];

    if (value !== "" && value !== undefined) {
      // Clean HR name only
      if (header === "Department HR Name") {
        value = stripParentheses(value);
      }
      html += `<tr><td><strong>${header}</strong></td><td>${value}</td></tr>`;
    }
  }

  html += "</table>";
  return html;
}




function getApproverDetails(role, row) {
  if (role === "Chair") return { name: row[4], email: row[5], role: "Chair" };
  if (role === "HR") return { name: row[16], email: row[17], role: "HR Partner" };
  if (role === "Office of Research") return { name: "Office of Research", email: OOR_EMAIL, role: "Office of Research" };
  return { name: "Unknown", email: "Unknown", role: "Unknown" };
}

function getApproverByColumn(sheet, uniqueId, nameColIndex) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf("Unique ID");
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === uniqueId) {
      return data[i][nameColIndex] || "N/A";
    }
  }
  return "N/A";
}

function getApprovalDateByColumn(sheet, uniqueId) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf("Unique ID");
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === uniqueId) {
      return data[i][0] ? Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") : "N/A";
    }
  }
  return "N/A";
}

function createFinalPDF(requesterName, volunteerName, uniqueId) {
  Logger.log("Starting PDF generation for: " + requesterName + " / " + volunteerName);

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sourceSheet = ss.getSheetById(SHEET_IDS.requests);
  var chairSheet = ss.getSheetById(SHEET_IDS.chair);
  var hrSheet = ss.getSheetById(SHEET_IDS.hr);
  var oorSheet = ss.getSheetById(SHEET_IDS.oor);


  var data = sourceSheet.getDataRange().getValues();
  var headers = data[0];
  const uidIndex = headers.indexOf("Unique ID");
  if (uidIndex === -1) {
    Logger.log("❌ 'Unique ID' column not found.");
    return null;
  }
  const matchedRow = data.find(row => row[uidIndex] === uniqueId);
  if (!matchedRow) {
    Logger.log("❌ No row found for Unique ID: " + uniqueId);
    return null;
  }

  let resubmissionBanner = null;
  const submissionTypeIndex = headers.indexOf("Submission Type");
  if (submissionTypeIndex !== -1) {
    const submissionType = (matchedRow[submissionTypeIndex] || '').toString().trim();
    if (submissionType.toLowerCase().startsWith('resubmission from')) {
      const match = submissionType.match(/VR\d+/);
      if (match) {
        const originalUid = match[0];
        resubmissionBanner = `This is a resubmission from UID ${originalUid}. Please review and re-complete the required fields.`;
      }
    }
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

  body.appendParagraph("").setSpacingAfter(12); // Adds vertical space

  // Core Facilities Warning Styled Box
  var warningTable = body.appendTable();
  var row = warningTable.appendTableRow();
  var cell = row.appendTableCell(getCoreFacilitiesTextBlurb());

  cell.setBackgroundColor("#E6F2EC");
  cell.setPaddingTop(6);
  cell.setPaddingBottom(6);
  cell.setPaddingLeft(8);
  cell.setPaddingRight(8);
  cell.setFontSize(10);
  cell.setFontFamily("Roboto");
  cell.setForegroundColor("#003C1E"); // Dark green text


  body.appendParagraph("\nRequest Details:").setHeading(DocumentApp.ParagraphHeading.HEADING2);

  var table = body.appendTable();
  for (var i = 0; i < headers.length; i++) {
    var question = headers[i];
    var answer = matchedRow[i] !== undefined && matchedRow[i] !== null ? matchedRow[i].toString() : "N/A";

    // Format DOB nicely
    if (question.toLowerCase().includes("dob")) {
      var dateVal = new Date(answer);
      if (!isNaN(dateVal.getTime())) {
        answer = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "MM/dd/yyyy");
      }
    }

    // Strip parentheses from HR name
    if (question === "Department HR Name") {
      answer = stripParentheses(answer);
    }

    var row = table.appendTableRow();
    row.appendTableCell(question);
    row.appendTableCell(answer);
  }


  body.appendParagraph("\nApprovals").setHeading(DocumentApp.ParagraphHeading.HEADING2);

  var approvalsTable = body.appendTable();
  var headerRow = approvalsTable.appendTableRow();
  headerRow.appendTableCell("Approver");
  headerRow.appendTableCell("Name");
  headerRow.appendTableCell("Timestamp");

  var roles = [
    ["Department Chair", chairSheet, 4],
    ["Department HR Service Partner", hrSheet, 4],
    ["Office of Research", oorSheet, 4]
  ];

  for (var i = 0; i < roles.length; i++) {
    var approverRow = approvalsTable.appendTableRow();
    approverRow.appendTableCell(roles[i][0]);
    const rawName = getApproverByColumn(roles[i][1], uniqueId, roles[i][2]);
    approverRow.appendTableCell(stripParentheses(rawName));
    approverRow.appendTableCell(getApprovalDateByColumn(roles[i][1], uniqueId));
  }

  var footer = doc.addFooter();
  var footerText = footer.appendParagraph("Confidential – For Internal Use Only | Faculty of Medicine & Dentistry - Office of Research");
  footerText.setFontSize(8).setForegroundColor("#666666").setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph("Approver Comments").setHeading(DocumentApp.ParagraphHeading.HEADING2);

  // Create table for grey box container
  var commentsTable = body.appendTable();
  var commentsRow = commentsTable.appendTableRow();
  var commentsCell = commentsRow.appendTableCell();

  // Style the cell like the cores snippet box
  commentsCell.setBackgroundColor("#E6E6E6"); // Light grey background
  commentsCell.setPaddingTop(0);
  commentsCell.setPaddingBottom(8);
  commentsCell.setPaddingLeft(10);
  commentsCell.setPaddingRight(10);
  commentsCell.setFontSize(10);
  commentsCell.setFontFamily("Roboto");
  commentsCell.setForegroundColor("#333333"); // Darker text

  // Gather comments as a single string with spacing
  var approverCommentsText = "";

  // Chair comment
  const chairComment = (() => {
    const chairData = chairSheet.getDataRange().getValues();
    const chairHeaders = chairData[0];
    const chairRow = chairData.find(r => r[chairHeaders.indexOf("Unique ID")] === uniqueId);
    return chairRow ? (chairRow[6] || "No comments.") : "Not available.";
  })();

  // HR comment
  const hrComment = (() => {
    const hrData = hrSheet.getDataRange().getValues();
    const hrHeaders = hrData[0];
    const hrRow = hrData.find(r => r[hrHeaders.indexOf("Unique ID")] === uniqueId);
    return hrRow ? (hrRow[7] || "No comments.") : "Not available.";
  })();

  // OoR comments
  const oorData = oorSheet.getDataRange().getValues();
  const oorHeaders = oorData[0];
  const oorRow = oorData.find(r => oorHeaders ? r[oorHeaders.indexOf("Unique ID")] === uniqueId : false);
  const oorAlignment = oorRow ? (oorRow[6] || "No alignment comments.") : "Not available.";
  const oorFinalNotes = oorRow ? (oorRow[10] || "No final notes.") : "Not available.";

  // Chair header (bold)
  commentsCell.appendParagraph("Chair:").setBold(true).setSpacingAfter(4);
  // Chair comment (normal)
  commentsCell.appendParagraph(chairComment)
    .setSpacingAfter(8)
    .setFontFamily("Roboto")
    .setFontSize(10)
    .setBold(false);  // <- here

  // HR header (bold)
  commentsCell.appendParagraph("HR:").setBold(true).setSpacingAfter(4);
  // HR comment (normal)
  commentsCell.appendParagraph(hrComment)
    .setSpacingAfter(8)
    .setFontFamily("Roboto")
    .setFontSize(10)
    .setBold(false);  // <- here

  // OoR header (bold, spelled out)
  commentsCell.appendParagraph("Office of Research:").setBold(true).setSpacingAfter(4);

  // OoR comments (indented, normal)
  commentsCell.appendParagraph("Alignment: " + oorAlignment)
    .setIndentStart(20)
    .setSpacingAfter(4)
    .setFontFamily("Roboto")
    .setFontSize(10)
    .setBold(false);  // <- here

  commentsCell.appendParagraph("Final Notes: " + oorFinalNotes)
    .setIndentStart(20)
    .setSpacingAfter(8)
    .setFontFamily("Roboto")
    .setFontSize(10)
    .setBold(false);  // <- here

  doc.saveAndClose();
  var pdfBlob = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
  Logger.log("PDF Created Successfully");
  return pdfBlob;
}

function redirectToSuccessPage(pageName) {
  return HtmlService.createHtmlOutputFromFile(pageName).getContent();
}

function getCoreFacilitiesTextBlurb() {
  return "Please note that volunteers are not permitted to access the FoMD Core Research Facilities instrumentation or services. Shadowing trained users or core staff may be permitted under certain circumstances and requires separate approval. Please contact Colleen Sunderland (ccreid1@ualberta.ca), Manager of Core Research Facilities, if applicable and to obtain approval. As per our user guidelines, research groups may have access privileges revoked if unauthorized individuals are found to be accessing core spaces or utilizing the instrumentation.";
}

function checkPendingApprovals() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 1. Chair Reminders
  sendReminders({
    sheetId: SHEET_IDS.requests,
    timestampCol: 0, // Column A
    statusSheetId: SHEET_IDS.status,
    targetStatus: "Chair",
    reminderRole: "Chair",
    emailIndex: 8, // Chair Email in requests sheet
    nameIndex: 7   // Chair Name in requests sheet
  });

  // 2. HR Reminders
  sendReminders({
    sheetId: SHEET_IDS.chair,
    timestampCol: 0, // Column A
    statusSheetId: SHEET_IDS.status,
    targetStatus: "HR",
    reminderRole: "HR",
    emailIndex: 1, // HR Email in HR sheet
    nameIndex: 4   // HR Name in HR sheet
  });

  // 3. OoR Reminders (based on HR approval timestamp)
  sendReminders({
    sheetId: SHEET_IDS.hr,
    timestampCol: 0, // Column A
    statusSheetId: SHEET_IDS.status,
    targetStatus: "OoR",
    reminderRole: "OoR",
    emailIndex: -1, // Not needed, single recipient
    nameIndex: -1
  });
}

function sendReminders(config) {
  const { sheetId, timestampCol, statusSheetId, targetStatus, reminderRole, emailIndex, nameIndex } = config;
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetById(sheetId);
  const statusSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetById(statusSheetId);

  const data = sheet.getDataRange().getValues();
  const statusData = statusSheet.getDataRange().getValues();
  const headers = statusData[0];
  const uidIndex = headers.indexOf("Unique ID");
  const statusIndex = headers.indexOf("Status");

  const now = new Date();
  const thresholds = [3, 5, 10];

  for (let i = 1; i < statusData.length; i++) {
    const uniqueId = statusData[i][uidIndex];
    const status = statusData[i][statusIndex];

    if (status !== targetStatus) continue;

    const record = data.find(row => row.includes(uniqueId));
    if (!record) continue;

    const timestamp = new Date(record[timestampCol]);
    const daysElapsed = Math.floor((now - timestamp) / (1000 * 60 * 60 * 24));

    if (thresholds.includes(daysElapsed)) {
      if (reminderRole === "Chair") {
        const email = record[emailIndex];
        const name = record[nameIndex];
        sendChairApprovalReminder(uniqueId, name, email);
      } else if (reminderRole === "HR") {
        const email = record[emailIndex];
        const name = record[nameIndex];
        sendHRApprovalReminder(uniqueId, name, email);
      } else if (reminderRole === "OoR") {
        sendOoRApprovalReminder(uniqueId);
      }
    }
  }
}

function sendChairApprovalReminder(uniqueId, chairName, chairEmail) {
  if (!chairEmail) return;
  const link = `${ScriptApp.getService().getUrl()}?action=getchairApprovalForm&tracking=${uniqueId}&role=Chair`;
  MailApp.sendEmail({
    to: chairEmail,
    subject: `Reminder: Chair Approval Needed – Volunteer Request`,
    htmlBody: `
      <p>Dear ${chairName},</p>
      <p>This is a reminder to review and approve the volunteer request:</p>
      <p><a href="${link}">Click here to open the Chair Approval Form</a></p>
      <p>Tracking ID: <strong>${uniqueId}</strong></p>
      <p>If you have questions, please contact <a href="mailto:vdradmin@ualberta.ca">vdradmin@ualberta.ca</a>.</p>`
  });
}

function sendHRApprovalReminder(uniqueId, hrName, hrEmail) {
  if (!hrEmail) return;
  const link = `${ScriptApp.getService().getUrl()}?action=gethrApprovalForm&tracking=${uniqueId}&role=HR`;
  MailApp.sendEmail({
    to: hrEmail,
    subject: `Reminder: HR Approval Needed – Volunteer Request`,
    htmlBody: `
      <p>Dear ${hrName},</p>
      <p>This is a reminder to review and approve the volunteer request:</p>
      <p><a href="${link}">Click here to open the HR Approval Form</a></p>
      <p>Tracking ID: <strong>${uniqueId}</strong></p>
      <p>If you have questions, please contact <a href="mailto:vdradmin@ualberta.ca">vdradmin@ualberta.ca</a>.</p>`
  });
}

function sendOoRApprovalReminder(uniqueId) {
  const link = `${ScriptApp.getService().getUrl()}?action=getoorApprovalForm&tracking=${uniqueId}&role=Office%20of%20Research`;
  MailApp.sendEmail({
    to: OOR_EMAIL,
    subject: `Reminder: Office of Research Approval Needed – Volunteer Request`,
    htmlBody: `
      <p>Reminder to review and approve the request for tracking ID: <strong>${uniqueId}</strong>.</p>
      <p><a href="${link}">Open OoR Approval Form</a></p>`
  });
}

function createReminderTrigger() {
  ScriptApp.newTrigger("checkPendingApprovals")
    .timeBased()
    .everyDays(1)
    .atHour(8) // adjust as needed
    .create();
}

function getAdmByDepartment(department) {
  const approvers = getApproversForForm();
  const match = approvers.ADM.find(a => a.department === department);
  return match || { name: "", email: "" };
}

function getSuccessPage(role) {
  if (role === 'Chair' || role === 'HR' || role === 'OoR') {
    return HtmlService.createTemplateFromFile("approver_success").evaluate().getContent();
  } else {
    return HtmlService.createTemplateFromFile("success").evaluate().getContent();
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Volunteer Tools")
    .addItem("Generate Resubmission Link", "showResubmissionSidebar")
    .addToUi();
}

function showResubmissionSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("resubmissionSidebar")
    .setTitle("Generate Resubmission Link");
  SpreadsheetApp.getUi().showSidebar(html);
}

function getResubmissionLinkForSelectedRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetById(SHEET_IDS.requests);
  const range = sheet.getActiveRange();
  const row = range.getRow();
  const uidCol = 21; // Column Y (index 24 in 1-based indexing)

  const uniqueId = sheet.getRange(row, uidCol).getValue();
  if (!uniqueId) return "❌ No Unique ID found in selected row.";

  const baseUrl = "https://script.google.com/a/macros/ualberta.ca/s/AKfycbyhDC6JQlmBUbjakE7xWVDhI_85S3z_lwwdlCKauanrrAtSYRzgV2vxqRa8oShyNEpy5Q/exec";
  const fullLink = `${baseUrl}?action=getResubmissionForm&tracking=${uniqueId}`;
  return fullLink;
}

function stripParentheses(name) {
  return name.replace(/\s*\(.*?\)\s*/g, "").trim();
}


