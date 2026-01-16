/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("get-subject-btn").onclick = getSubject;
    document.getElementById("get-sender-btn").onclick = getSenderInfo;

    // Load initial email info
    loadEmailInfo();
  }
});

/**
 * Load basic email information when task pane opens
 */
function loadEmailInfo() {
  Office.context.mailbox.item.subject.getAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const emailInfoDiv = document.getElementById("email-info");
      emailInfoDiv.innerHTML = `
        <p><strong>Current Email:</strong></p>
        <p>${result.value}</p>
      `;
    }
  });
}

/**
 * Get and display the email subject
 */
function getSubject() {
  const resultsDiv = document.getElementById("results");

  Office.context.mailbox.item.subject.getAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      resultsDiv.className = "results-box success";
      resultsDiv.textContent = `Subject: ${result.value}`;
    } else {
      resultsDiv.className = "results-box error";
      resultsDiv.textContent = `Error: ${result.error.message}`;
    }
  });
}

/**
 * Get and display sender information
 */
function getSenderInfo() {
  const resultsDiv = document.getElementById("results");

  const item = Office.context.mailbox.item;

  // Get sender email address
  item.from.getAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const sender = result.value;
      resultsDiv.className = "results-box success";
      resultsDiv.textContent = `Sender Name: ${sender.displayName}\nEmail: ${sender.emailAddress}`;
    } else {
      resultsDiv.className = "results-box error";
      resultsDiv.textContent = `Error: ${result.error.message}`;
    }
  });
}