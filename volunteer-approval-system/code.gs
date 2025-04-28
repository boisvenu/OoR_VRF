// Serve the Volunteer Request Status form
function doGet() {
  return HtmlService.createHtmlOutputFromFile('statusChecker')
    .setTitle('Volunteer Request Status')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Handle the form submission from the user
function checkVolunteerStatus(trackingNumber) {
  return getVolunteerRequestStatus(trackingNumber);
}

// Pull approval information from all stages
function getVolunteerRequestStatus(trackingNumber) {
  const ss = SpreadsheetApp.openById('1048SyHN-e4XujUkJOXj1j5rMQdZ4R4ZwIEYcLTK0OJk');

  const sheets = [
    {
      id: 362146665,
      name: 'Chair Approval',
      uniqueIdCol: 8,   // Column H
      approverCol: 5,   // Column E
      statusCol: 6,     // Column F
      commentsCol: 7    // Column G
    },
    {
      id: 1074532913,
      name: 'HR Approval',
      uniqueIdCol: 9,   // Column I
      approverCol: 5,   // Column E
      statusCol: 7,     // Column G
      commentsCol: 8    // Column H
    },
    {
      id: 876106890,
      name: 'Office of Research Approval',
      uniqueIdCol: 12,  // Column L
      approverCol: 5,   // Column E
      statusCol: 10,    // Column J
      commentsCol: 11   // Column K
    }
  ];

  const result = [];

  sheets.forEach(sheetInfo => {
    const sheet = ss.getSheetById(sheetInfo.id);
    const data = sheet.getDataRange().getValues();

    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][sheetInfo.uniqueIdCol - 1] === trackingNumber) {
        let rawStatus = (data[i][sheetInfo.statusCol - 1] || '').toString().trim();
        let status = 'Pending';

        if (rawStatus.toLowerCase() === 'approve') {
          status = 'Approved';
        } else if (rawStatus.toLowerCase() === 'approved' || rawStatus.toLowerCase() === 'rejected') {
          status = rawStatus.charAt(0).toUpperCase() + rawStatus.slice(1).toLowerCase();
        }

        const approver = data[i][sheetInfo.approverCol - 1] || sheetInfo.name.replace(' Approval', '');
        const comments = data[i][sheetInfo.commentsCol - 1] || '';

        let stageStatus = {
          stage: sheetInfo.name,
          status: status,
          approver: approver,
          comments: ''
        };

        if (status.toLowerCase() === 'rejected' && comments) {
          stageStatus.comments = comments;
        }

        result.push(stageStatus);
        found = true;
        break;
      }
    }

    if (!found) {
      result.push({
        stage: sheetInfo.name,
        status: 'Pending',
        approver: sheetInfo.name.replace(' Approval', ''),
        comments: ''
      });
    }
  });

  return result;
}
