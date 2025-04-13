# Volunteer Request Approval System (FoMD)

This repository contains the source code for a Google Apps Script-based workflow used by the Faculty of Medicine & Dentistry (FoMD) to manage volunteer request approvals. The system handles form submissions, sequential multi-role approvals (Chair, HR, and Office of Research), PDF generation, and automated email notifications.

---

## 🔁 Workflow Overview

1. **Volunteer Request Submission (Initial Form)**
   - Populates the **Source Sheet** with request details
   - Assigns a **Unique ID**
   - Creates a record in the **Status Sheet**
   - Triggers an approval request email to the **Department Chair**

2. **Chair Approval Form**
   - Approves → HR is notified
   - Rejects → Rejection email sent to requester, status updated

3. **HR Approval Form**
   - Approves → Office of Research is notified
   - Rejects → Rejection email sent to requester, status updated

4. **OoR Approval Form**
   - Approves → Final PDF generated and emailed to all parties
   - Rejects → Rejection email sent to requester, status updated

---

## 📄 Script Files

### `onApprovalFormSubmit(e)`
Triggered on form submission. Handles logic based on which sheet the form data came from:

- **Initial Form (Source):** assigns UUID, sends Chair request
- **Chair Form:** logs decision, updates status, triggers HR email
- **HR Form:** logs decision, updates status, triggers OoR email
- **OoR Form:** logs decision, updates status, sends final approval PDF

### `send[Role]ApprovalEmail(...)`
Sends emails to each role with prefilled approval links and summary.

### `sendFinalApprovalEmail(...)`
Sends completed PDF with approvals to all involved emails (Chair, HR, OoR, Requester).

### `sendRejectionEmail(...)`
Handles formatting and dispatch of rejection notification email.

### `createFinalPDF(...)`
- Pulls form data and approval metadata
- Builds structured PDF using `DocumentApp`
- Includes clickable waiver/background links

### `getApproverByColumn(...)`
- Retrieves name of approver by Unique ID and column index.

### `getApprovalDateByColumn(...)`
- Formats timestamp of approval.

### `buildEmailHtmlSummary(...)`
- Builds HTML table preview of submission used in emails.

---

## 📋 Spreadsheet Setup

Google Sheet ID: `1048SyHN-e4XujUkJOXj1j5rMQdZ4R4ZwIEYcLTK0OJk`

### Tabs (Sheet IDs):
- **Volunteer Form Responses:** `1447833263`
- **Status:** `684280098`
- **Chair Approval:** `362146665`
- **HR Approval:** `1074532913`
- **OoR Approval:** `876106890`

### Column Reference Highlights:
- Column X (index 23): Unique ID
- Column B: Requester Email
- Column C: Requester Name
- Column G: Volunteer Name
- Column Q: HR Partner Name
- Column R: HR Partner Email

---

## 🖥️ Web Interface (HTML Page)

Provides a basic UI to list submissions and allow manual approval/denial:

```html
<table>
  <tr>
    <th>Name</th>
    <th>Department</th>
    <th>Waiver</th>
    <th>Background Check</th>
    <th>Approve/Deny</th>
  </tr>
</table>
```
Buttons call server-side `approveRequest(id)` or `denyRequest(id)` methods.

---

## 🕒 Timestamp Sync Helper

```js
function copyNewTimestampWithUniqueID() {
  // Copies the timestamp + default 'Chair' status to Status sheet
}
```
Used to sync data rows with status tracking on initial form submission.

---

## 📦 Notes & Caveats

- Files in "Shared with Me" (e.g., uploaded waiver/background docs) **do not count** against your Drive quota unless you copy or move them.
- Ensure column indices (e.g., approval decision in column G for HR) match the form structure.
- HR Name in the final PDF pulls from column index `3` in `hrSheet` to correct display issues.

---

## 🔐 Access and Permissions
- Make sure all forms and Drive folders used for uploading documents are shared appropriately.
- Approvers must be logged into their UAlberta accounts to access prefilled approval forms.

---

## 🛠️ Next Steps (Future Improvements)
- Add signature capture
- Add retry logic for email failures
- Add a submission dashboard with status filters
- Improve PDF formatting/styling

---

## 📬 Contact
For questions, contact:
**Office of Research** — vdradmin@ualberta.ca

