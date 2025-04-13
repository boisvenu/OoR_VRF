# Volunteer Approval System for FoMD

This Google Apps Script-based system automates multi-stage approvals for volunteer onboarding in the Faculty of Medicine & Dentistry (FoMD). It collects information through a Google Form, manages approvals through three sequential approvers, and generates a final PDF once all approvals are complete.

---

## 🔄 Workflow Overview

1. **Initial Form Submission**: The original requester submits a form.
2. **Chair Approval**: The department chair receives a pre-filled form and reviews the request.
3. **HR Approval**: If approved by the chair, the HR partner is emailed with a pre-filled form for approval.
4. **Office of Research Approval (OoR)**: Final review and approval by the Office of Research.
5. **PDF Generation**: A final summary PDF is created and distributed to all stakeholders.
6. **Status Tracking**: Each stage is tracked in a dedicated Google Sheet tab.

---

## 🗂️ Files

### `Code.gs`
Main script file. Includes:
- UUID generation
- `onApprovalFormSubmit()` routing logic
- Email functions for Chair, HR, and OoR
- PDF generation with form summary and approvals

### `index.html`
Web app interface for real-time review of submission data (optional utility panel for internal use).

### `utils/copyNewTimestampWithUniqueID.gs`
Utility function to append timestamp and initialize tracking.

---

## 🧾 Sheets Structure

| Sheet Name       | Purpose                                      |
|------------------|----------------------------------------------|
| Source (ID: `1447833263`) | Raw form responses                  |
| Chair Approval (ID: `362146665`) | Captures Chair decisions      |
| HR Approval (ID: `1074532913`) | Captures HR decisions          |
| OoR Approval (ID: `876106890`) | Captures Office of Research     |
| Status (ID: `684280098`) | Tracks workflow state per submission |

---

## ✅ Requirements

- Google Sheet with tab IDs matching those referenced above
- Three Google Forms (Chair, HR, OoR) with entry prefill mappings
- Apps Script deployed as a Web App (if using the web review panel)
- Permissions for the script to send email and access Drive

---

## ⚠️ Notes

- Waiver and background check links will appear in the Shared Drive section as viewed files but do not consume Drive space unless copied.
- All emails are sent automatically based on submission and approval logic.
- Ensure form column ordering matches expected indexes (e.g., approval decision must be in the correct column).

---

## 📬 Support

For help or updates, contact: **vdradmin@ualberta.ca**

