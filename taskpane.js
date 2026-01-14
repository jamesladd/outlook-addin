// Initialize Office.js
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Set up the Done button click handler
    document.getElementById('doneButton').addEventListener('click', closeTaskPane);

    // Load all data
    loadUserInformation();
    loadReportData();
  }
});

/**
 * Load user information from the current mailbox
 */
function loadUserInformation() {
  try {
    const mailbox = Office.context.mailbox;
    const userProfile = mailbox.userProfile;

    // Get user email address
    const userEmail = userProfile.emailAddress;
    document.getElementById('userEmail').textContent = userEmail;

    // Get user display name
    const userName = userProfile.displayName || 'Not available';
    document.getElementById('userName').textContent = userName;

    // Get account type (role) - this is a basic representation
    const accountType = userProfile.accountType || 'Not available';
    document.getElementById('userRole').textContent = accountType;

    // Permissions - Office.js doesn't directly expose all permissions
    // We show what we can access
    const permissions = [];
    if (mailbox.diagnostics) {
      permissions.push('Diagnostics: Enabled');
    }
    if (Office.context.mailbox.item) {
      permissions.push('Item Access: Enabled');
    }
    document.getElementById('userPermissions').textContent =
      permissions.length > 0 ? permissions.join(', ') : 'Basic permissions';

    // Get available mailboxes
    // Note: Office.js API has limited support for listing all mailboxes
    // We'll show the primary mailbox and any we can detect
    const mailboxList = document.getElementById('mailboxList');
    mailboxList.innerHTML = '';

    // Add primary mailbox
    const primaryItem = document.createElement('li');
    primaryItem.textContent = `${userEmail} (Primary)`;
    mailboxList.appendChild(primaryItem);

    // Try to get additional mailbox info if available
    if (mailbox.diagnostics && mailbox.diagnostics.hostName) {
      const hostItem = document.createElement('li');
      hostItem.textContent = `Host: ${mailbox.diagnostics.hostName}`;
      mailboxList.appendChild(hostItem);
    }

    // Show the user info section and hide loading
    document.getElementById('loading').style.display = 'none';
    document.getElementById('userInfo').style.display = 'block';

  } catch (error) {
    console.error('Error loading user information:', error);
    document.getElementById('loading').textContent =
      'Error loading user information: ' + error.message;
  }
}

/**
 * Load report data from the API
 */
async function loadReportData() {
  const reportSection = document.getElementById('reportSection');
  const reportError = document.getElementById('reportError');
  const reportTableBody = document.getElementById('reportTableBody');

  try {
    // API endpoint and credentials
    const apiUrl = 'https://inboxagent.dev.aportio.net/api/v2/reports/report_emails/';
    const authToken = 'UDhLVHQ5TWMxeDFUeEJXRmlaRGluMk9FcEZWaEVKaXM6YTZ4eVgyQ0t5MWtVQkl1djV3R1BnMHRqNDF3d3dpekNoRWpNQ2ozcUtWTWRCUEpwYTVycXBtWG02cHNUc1Jq';

    // Make the GET request with Basic authentication
    const response = await fetch(apiUrl, {
      method: 'GET',
      headers: {
        'Authorization': `Basic ${authToken}`,
        'Content-Type': 'application/json'
      }
    });

    // Check if request was successful
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    // Parse the JSON response
    const data = await response.json();

    // Clear any existing table rows
    reportTableBody.innerHTML = '';

    // Check if data exists and is an array
    if (data && Array.isArray(data)) {
      // If no results
      if (data.length === 0) {
        const row = reportTableBody.insertRow();
        const cell = row.insertCell(0);
        cell.colSpan = 4;
        cell.textContent = 'No email reports found';
        cell.style.textAlign = 'center';
        cell.style.fontStyle = 'italic';
      } else {
        // Populate table with data
        data.forEach(item => {
          const row = reportTableBody.insertRow();

          // ID column
          const idCell = row.insertCell(0);
          idCell.textContent = item.id || '-';

          // Subject column
          const subjectCell = row.insertCell(1);
          subjectCell.textContent = item.subject || '-';

          // Date column
          const dateCell = row.insertCell(2);
          dateCell.textContent = item.date ?
            new Date(item.date).toLocaleDateString() : '-';

          // Status column
          const statusCell = row.insertCell(3);
          statusCell.textContent = item.status || '-';
        });
      }
    } else if (data && typeof data === 'object') {
      // If the API returns a single object or different structure
      const row = reportTableBody.insertRow();
      const cell = row.insertCell(0);
      cell.colSpan = 4;
      cell.textContent = 'Data received in unexpected format';
      cell.style.textAlign = 'center';

      // Log for debugging
      console.log('API Response:', data);
    }

    // Show the report section
    reportSection.style.display = 'block';

  } catch (error) {
    console.error('Error loading report data:', error);

    // Show error message
    reportError.textContent = `Error loading report data: ${error.message}`;
    reportError.style.display = 'block';
    reportSection.style.display = 'block';

    // Show a message in the table
    reportTableBody.innerHTML = '';
    const row = reportTableBody.insertRow();
    const cell = row.insertCell(0);
    cell.colSpan = 4;
    cell.textContent = 'Failed to load report data';
    cell.style.textAlign = 'center';
    cell.style.color = '#d13438';
  }
}

/**
 * Close the task pane when Done button is clicked
 */
function closeTaskPane() {
  // In Outlook Web Add-ins, we can't programmatically close the task pane
  // But we can notify the user or clear the content
  if (Office.context.ui) {
    // Try to close if supported
    try {
      Office.context.ui.closeContainer();
    } catch (e) {
      // If closeContainer is not available, show a message
      alert('Please close the task pane using the X button in the top corner.');
    }
  } else {
    // Fallback: inform user to close manually
    alert('Please close the task pane using the X button in the top corner.');
  }
}
