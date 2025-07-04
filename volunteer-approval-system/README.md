# Volunteer Request Web App for FoMD

This Google Apps Script-based web application automates a structured, multi-stage volunteer request and approval process for the Faculty of Medicine & Dentistry (FoMD) at the University of Alberta. It replaces traditional Google Forms with fully custom HTML interfaces and integrates Google Sheets, Drive, and Gmail for end-to-end coordination, approval routing, and PDF generation.

---

## Workflow Overview

1. **Volunteer Submission**  
   A requester completes a styled, multi-page HTML form and uploads any required documents.

2. **Chair Approval**  
   The Department Chair receives a personalized approval link by email and reviews the request.

3. **HR Partner Approval**  
   If Chair approval is granted, the designated HR Partner receives their own approval link and fields.

4. **Office of Research (OoR) Final Approval**  
   Upon HR approval, the Office of Research provides final authorization.

5. **PDF Generation**  
   After full approval, the system compiles all data, signatures, and key metadata into a formatted PDF, which is emailed and stored in Drive.

6. **Status Tracking**  
   Each request is assigned a unique ID, and status is updated across tabs in a central Google Sheet.

7. **Resubmissions**  
   Rejected requests can be resubmitted with updated information, triggering a fresh approval sequence.

8. **Status Checker Web App**  
   Submitters can check the real-time status of their request using a UID-based search interface.

---

## Files and Structure

### Top-Level
- `appsscript.json`  
  Google Apps Script manifest file

### `code/` – Server-side logic
- `main.gs` – doGet() handler, form routing, and success pages
- `pdf_generation.gs` – PDF formatting and blob creation
- `approval_logic.gs` – Approval workflow functions for Chair, HR, and OoR
- `email_templates.gs` – Email formatting and sending logic
- `reminders.gs` – Reminder scheduling (3/5/10 days)
- `utils.gs` – UID generation, value lookups, and shared helpers

### `html/` – Web front-end templates
- `submission_form.html` – Volunteer submission (multi-page)
- `resubmission_form.html` – Resubmission form (pre-filled)
- `chairApprovalForm.html` – Chair approval interface
- `hrApprovalForm.html` – HR approval interface
- `oorApprovalForm.html` – Office of Research approval interface
- `status_checker.html` – Status lookup by UID
- `success.html` – Confirmation page after submission or final approval
- `resubmissionSidebar.html` – Internal sidebar tool to generate resubmission links

---

## Google Sheets Structure

| Sheet Name        | Sheet ID       | Description                                       |
|-------------------|----------------|---------------------------------------------------|
| Form Responses    | `1130224705`   | Raw submission entries                           |
| Chair Approval    | `362146665`    | Chair-level decisions                            |
| HR Approval       | `1074532913`   | HR approval responses                            |
| OoR Approval      | `876106890`    | Final decision by Office of Research             |
| Status            | `684280098`    | Tracks each request’s current status             |
| Approver Directory| `352748199`    | Lookup for Chair/HR/ADM names + emails           |

---

## Technical Requirements

- **Google Spreadsheet** with the above tabs
- **Deployed Web App** with:
  - `doGet(e)` serving all form and success routes
  - `doPost(e)` if needed for fetch-based submissions
- **Google Drive Folder** for storing final PDFs
- **Script Authorization** to access:
  - Google Sheets (read/write)
  - Gmail (sendEmail)
  - Drive (createFile)

---

## Behavior Details

- **UID Generation**  
  Each submission is assigned a persistent Unique ID for cross-sheet matching and status checks.

- **Conditional Email Routing**  
  Chair → HR → OoR approvals are sequentially triggered only after the prior approver signs.

- **Resubmission Support**  
  Rejected entries can be re-submitted via a special link, marked and tracked as "Resubmission."

- **Reminder Emails**  
  Automatically sent to approvers after 3, 5, and 10 days if no action is taken.

- **PDF Summary Generation**  
  Final approval triggers a custom PDF with all form data, links, comments, and signatures. This is attached to the final confirmation email and stored in Drive.

- **Status Checker**  
  Submitters can use the UID to check current approval status, rejection reasons, or resubmission needs.

---

## Known Limitations

- Sheet column order must remain consistent with the script logic.
- File uploads are limited to PDF format and must be stored in Drive.
- Drive sharing permissions must be manually reviewed if users lack access.
- Only one active submission per volunteer at a time is supported (based on duplicate logic).

---

## Support and Contact

For issues or integration support, please contact:

**FoMD Office of Research**  
**Email**: vdradmin@ualberta.ca

---

## License

This project is licensed under the [Creative Commons BY-NC 4.0 License](https://creativecommons.org/licenses/by-nc/4.0/).  
You may share and adapt the system for non-commercial purposes with attribution. Commercial use is prohibited without written consent.

---
