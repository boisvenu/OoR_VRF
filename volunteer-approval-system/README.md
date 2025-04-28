# Volunteer Approval System for FoMD

This Google Apps Script-based system automates a structured, multi-stage approval workflow for volunteer onboarding in the Faculty of Medicine & Dentistry (FoMD) at the University of Alberta. It integrates Google Forms, Google Sheets, and Drive to manage submissions, coordinate approvals, and generate final PDF summaries.

---

## Workflow Overview

1. **Initial Form Submission**  
   A requester submits a Google Form with details about the volunteer.

2. **Department Chair Approval**  
   The Department Chair receives a pre-filled approval form via email, reviews the request, and approves or rejects it.

3. **HR Partner Approval**  
   If Chair approval is granted, an HR Partner receives their own pre-filled approval form.

4. **Office of Research (OoR) Final Approval**  
   Upon HR approval, the Office of Research reviews the request and provides final approval.

5. **PDF Generation**  
   Once fully approved, the system compiles all form data and approvals into a PDF summary and emails it to all stakeholders.

6. **Status Tracking**  
   Each submission is assigned a unique ID and its progress is tracked across tabs in a central Google Sheet.

7. **Volunteer Request Status Tracker**  
   Submitters can enter their Unique ID into a web app to view real-time approval progress through each stage.

---

## Files and Structure

### `Code.gs`
Main script file containing:
- `generateUniqueId()` for UUID assignment
- `onApprovalFormSubmit(e)` to detect form submissions and route logic
- Email functions: `sendChairApprovalEmail()`, `sendHRApprovalEmail()`, `sendOoRApprovalEmail()`, `sendFinalApprovalEmail()`, `sendRejectionEmail()`, `sendSubmissionConfirmationEmail()`
- `createFinalPDF()` to compile the final approval package
- Helper functions: formatting, retrieval, and summary generation

### `index.html` (optional)
A web interface (if deployed as a Web App) for internal use to review submission data in real time.

### `utils/copyNewTimestampWithUniqueID.gs` (optional utility)
Used to initialize and time-stamp new submissions with a unique tracking ID.

### `VolunteerRequestTracker/`
Standalone web app for submitters to check the status of their volunteer request.
- `Code.gs` (Web app server-side logic for serving the tracker)
- `statusChecker.html` (Front-end HTML form and dynamic status display)

---

## Google Sheets Structure

| Sheet Name         | Sheet ID      | Description                                 |
|--------------------|---------------|---------------------------------------------|
| Source             | `1447833263`  | Stores raw form responses                   |
| Chair Approval     | `362146665`   | Stores Chair approval responses             |
| HR Approval        | `1074532913`  | Stores HR approval responses                |
| OoR Approval       | `876106890`   | Stores final Office of Research decisions   |
| Status             | `684280098`   | Tracks status progression for each request  |

---

## Technical Requirements

- A central Google Spreadsheet with the above tabs and column ordering aligned to the script
- Three separate Google Forms (Chair, HR, OoR), each configured to:
  - Accept pre-filled parameters (e.g., requester name, volunteer name, unique ID)
  - Store decisions and optional comments
- File access:
  - Submitters must share Waiver and Background Check files via Google Drive
- Script permissions:
  - Must be authorized to send email and access Google Drive
- Deployment:
  - Script must be deployed as a Web App only if using the optional interfaces

---

## Behavior Details

- **UUID Generation**: Each submission is assigned a unique identifier for tracking across forms and sheets.
- **Conditional Routing**: Form submission triggers determine the source sheet and route the request accordingly.
- **Automated Emails**:
  - Prefilled approval forms are emailed at each stage.
  - Rejection triggers notification with comments.
  - A final PDF is created and emailed only after all three approvals are received.
- **Volunteer Status Tracking**:
  - A web app allows submitters to track progress by entering their Unique ID.
  - Displays approval status, rejection reasons, and current approver in real-time.
- **PDF Composition**:
  - Includes a logo, form answers, clickable waiver/background links, and an approvals table.
  - Sent as an attachment to requester, approvers, and administrative contacts.

---

## Known Limitations

- If form structure changes (column order, field labels), script functions must be updated accordingly.
- Drive file permissions must be managed manually by submitters unless an automated sharing script is added.
- The system does not currently support re-submission or editing after rejection (can be implemented with version control logic).

---

## Support and Contact

For issues, improvements, or integration assistance, please contact:  
**FoMD Office of Research**  
**Email**: vdradmin@ualberta.ca

---

## License

This project is licensed under the [Creative Commons BY-NC 4.0 License](https://creativecommons.org/licenses/by-nc/4.0/).  
You may share and adapt the system for non-commercial use with proper attribution. Commercial use is prohibited without permission.

---
