/* global Office */

// Global state
let eventCount = {
  total: 0,
  send: 0,
  change: 0,
  action: 0
};

let currentItem = null;
let eventLog = [];

// Initialize Office
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log("=== InboxAgent TaskPane Initialized ===");
    console.log("Host:", info.host);
    console.log("Platform:", info.platform);

    // Set up the task pane
    setupTaskPane();

    // Load current item
    loadCurrentItem();

    // Set up item changed handler
    setupItemChangedHandler();

    // Set up action tracking
    setupActionTracking();

    logEvent("System", "TaskPane Initialized", {
      host: info.host,
      platform: info.platform,
      timestamp: new Date().toISOString()
    });
  }
});

// Setup task pane UI and event handlers
function setupTaskPane() {
  console.log("Setting up TaskPane UI...");

  // Clear log button
  document.getElementById('clearLog').addEventListener('click', () => {
    clearEventLog();
  });

  // Track Reply button
  document.getElementById('trackReply').addEventListener('click', () => {
    trackAction('Reply');
  });

  // Track Forward button
  document.getElementById('trackForward').addEventListener('click', () => {
    trackAction('Forward');
  });

  // Track Reply All button
  document.getElementById('trackReplyAll').addEventListener('click', () => {
    trackAction('ReplyAll');
  });

  // Refresh Item button
  document.getElementById('refreshItem').addEventListener('click', () => {
    loadCurrentItem();
  });

  // Get Properties button
  document.getElementById('getProperties').addEventListener('click', () => {
    getAllItemProperties();
  });

  console.log("TaskPane UI setup complete");
}

// Load current item information
function loadCurrentItem() {
  console.log("Loading current item...");

  const item = Office.context.mailbox.item;

  if (!item) {
    console.log("No item currently selected");
    document.getElementById('currentItemInfo').innerHTML =
      '<p class="placeholder">No item selected</p>';
    return;
  }

  currentItem = item;

  const itemType = item.itemType;
  const subject = item.subject;
  const itemId = item.itemId || "N/A (Compose Mode)";
  const conversationId = item.conversationId || "N/A";

  let itemInfo = {
    type: itemType,
    subject: subject,
    itemId: itemId,
    conversationId: conversationId,
    mode: item.itemId ? "Read" : "Compose"
  };

  // Get additional properties based on item type
  if (itemType === Office.MailboxEnums.ItemType.Message) {
    itemInfo.from = getEmailAddress(item.from);
    itemInfo.to = getRecipients(item.to);
    itemInfo.cc = getRecipients(item.cc);
    itemInfo.internetMessageId = item.internetMessageId || "N/A";

    if (item.itemId) {
      // Read mode
      itemInfo.dateTimeCreated = item.dateTimeCreated ?
        new Date(item.dateTimeCreated).toLocaleString() : "N/A";
      itemInfo.dateTimeModified = item.dateTimeModified ?
        new Date(item.dateTimeModified).toLocaleString() : "N/A";
      itemInfo.sender = getEmailAddress(item.sender);
    }
  } else if (itemType === Office.MailboxEnums.ItemType.Appointment) {
    itemInfo.start = item.start ? new Date(item.start).toLocaleString() : "N/A";
    itemInfo.end = item.end ? new Date(item.end).toLocaleString() : "N/A";
    itemInfo.location = item.location || "N/A";
    itemInfo.organizer = getEmailAddress(item.organizer);
    itemInfo.requiredAttendees = getRecipients(item.requiredAttendees);
    itemInfo.optionalAttendees = getRecipients(item.optionalAttendees);
  }

  // Display item information
  displayItemInfo(itemInfo);

  // Log the item load
  logEvent("Item", "Item Loaded", itemInfo);

  console.log("Current item loaded:", itemInfo);
}

// Display item information in the UI
function displayItemInfo(itemInfo) {
  const container = document.getElementById('currentItemInfo');

  let html = '<div class="item-details">';

  for (const [key, value] of Object.entries(itemInfo)) {
    const displayKey = key.replace(/([A-Z])/g, ' $1').trim();
    const capitalizedKey = displayKey.charAt(0).toUpperCase() + displayKey.slice(1);

    html += `<p><strong>${capitalizedKey}:</strong> ${value || 'N/A'}</p>`;
  }

  html += '</div>';

  container.innerHTML = html;
}

// Helper function to get email address
function getEmailAddress(emailObj) {
  if (!emailObj) return "N/A";

  if (typeof emailObj === 'string') return emailObj;

  if (emailObj.emailAddress) {
    return `${emailObj.displayName || ''} <${emailObj.emailAddress}>`;
  }

  return "N/A";
}

// Helper function to get recipients
function getRecipients(recipientsObj) {
  if (!recipientsObj) return "N/A";

  // In compose mode, we need to use async methods
  if (typeof recipientsObj.getAsync === 'function') {
    return "Loading...";
  }

  // In read mode, it's an array
  if (Array.isArray(recipientsObj)) {
    return recipientsObj.map(r =>
      `${r.displayName || ''} <${r.emailAddress}>`
    ).join(', ') || "N/A";
  }

  return "N/A";
}

// Setup item changed handler
function setupItemChangedHandler() {
  console.log("Setting up item changed handler...");

  if (Office.context.mailbox && Office.context.mailbox.addHandlerAsync) {
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      onItemChanged,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Item changed handler registered successfully");
          logEvent("System", "Item Changed Handler Registered", {
            status: "Success"
          });
        } else {
          console.error("Failed to register item changed handler:", result.error);
          logEvent("Error", "Item Changed Handler Registration Failed", {
            error: result.error.message
          });
        }
      }
    );
  }
}

// Item changed event handler
function onItemChanged() {
  console.log("=== ITEM CHANGED EVENT ===");
  console.log("A new item has been selected");

  const item = Office.context.mailbox.item;

  const eventData = {
    timestamp: new Date().toISOString(),
    previousItem: currentItem ? {
      type: currentItem.itemType,
      subject: currentItem.subject,
      itemId: currentItem.itemId
    } : null,
    newItem: item ? {
      type: item.itemType,
      subject: item.subject,
      itemId: item.itemId
    } : null
  };

  console.log("Event Data:", eventData);

  logEvent("Change", "Item Changed", eventData, "event-change");

  // Reload the current item
  loadCurrentItem();
}

// Setup action tracking
function setupActionTracking() {
  console.log("Setting up action tracking...");

  const item = Office.context.mailbox.item;

  if (!item) {
    console.log("No item available for action tracking");
    return;
  }

  // Track body changes in compose mode
  if (item.body && typeof item.body.getAsync === 'function') {
    console.log("Setting up body change tracking...");
    // Note: There's no direct event for body changes,
    // but we can check periodically or on explicit actions
  }

  // Track attachment changes
  if (item.attachments) {
    console.log("Initial attachments:", item.attachments.length);
  }
}

// Track user actions (Reply, Forward, Reply All)
function trackAction(actionType) {
  console.log(`=== TRACKING ACTION: ${actionType} ===`);

  const item = Office.context.mailbox.item;

  if (!item) {
    console.log("No item available to track action");
    alert("No item selected");
    return;
  }

  const actionData = {
    action: actionType,
    timestamp: new Date().toISOString(),
    item: {
      type: item.itemType,
      subject: item.subject,
      itemId: item.itemId || "N/A",
      conversationId: item.conversationId
    }
  };

  // For read mode items, we can display the compose form
  if (item.itemId) {
    switch(actionType) {
      case 'Reply':
        console.log("Initiating Reply action");
        item.displayReplyForm({
          htmlBody: `<br><br><i>Tracked by InboxAgent at ${new Date().toLocaleString()}</i>`
        });
        actionData.details = "Reply form displayed with tracking note";
        break;

      case 'ReplyAll':
        console.log("Initiating Reply All action");
        item.displayReplyAllForm({
          htmlBody: `<br><br><i>Tracked by InboxAgent at ${new Date().toLocaleString()}</i>`
        });
        actionData.details = "Reply All form displayed with tracking note";
        break;

      case 'Forward':
        console.log("Initiating Forward action");
        item.displayReplyForm({
          htmlBody: `<br><br><i>Forwarded and tracked by InboxAgent at ${new Date().toLocaleString()}</i>`
        });
        actionData.details = "Forward form displayed with tracking note";
        break;
    }
  } else {
    actionData.details = "Item is in compose mode, action cannot be performed";
    console.log(actionData.details);
  }

  console.log("Action Data:", actionData);

  logEvent("Action", `User Action: ${actionType}`, actionData, "event-action");

  eventCount.action++;
  updateStatistics();
}

// Get all item properties
function getAllItemProperties() {
  console.log("=== GETTING ALL ITEM PROPERTIES ===");

  const item = Office.context.mailbox.item;

  if (!item) {
    console.log("No item available");
    alert("No item selected");
    return;
  }

  const properties = {
    // Basic properties
    itemType: item.itemType,
    itemId: item.itemId,
    conversationId: item.conversationId,
    subject: item.subject,

    // Dates
    dateTimeCreated: item.dateTimeCreated ?
      new Date(item.dateTimeCreated).toISOString() : null,
    dateTimeModified: item.dateTimeModified ?
      new Date(item.dateTimeModified).toISOString() : null,

    // Categories
    categories: item.categories || [],

    // Importance
    importance: item.importance,

    // Sensitivity
    sensitivity: item.sensitivity,

    // Internet headers (if available)
    internetMessageId: item.internetMessageId,

    // Normalized subject
    normalizedSubject: item.normalizedSubject,

    // Item class
    itemClass: item.itemClass
  };

  // Message-specific properties
  if (item.itemType === Office.MailboxEnums.ItemType.Message) {
    properties.from = item.from;
    properties.sender = item.sender;
    properties.to = item.to;
    properties.cc = item.cc;
    properties.bcc = item.bcc;
    properties.attachments = item.attachments ? item.attachments.length : 0;
  }

  // Appointment-specific properties
  if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
    properties.start = item.start ? new Date(item.start).toISOString() : null;
    properties.end = item.end ? new Date(item.end).toISOString() : null;
    properties.location = item.location;
    properties.organizer = item.organizer;
    properties.requiredAttendees = item.requiredAttendees;
    properties.optionalAttendees = item.optionalAttendees;
    properties.resources = item.resources;
    properties.recurrence = item.recurrence;
  }

  // User profile
  properties.userProfile = {
    emailAddress: Office.context.mailbox.userProfile.emailAddress,
    displayName: Office.context.mailbox.userProfile.displayName,
    timeZone: Office.context.mailbox.userProfile.timeZone
  };

  // Diagnostics
  properties.diagnostics = {
    hostName: Office.context.mailbox.diagnostics.hostName,
    hostVersion: Office.context.mailbox.diagnostics.hostVersion,
    OWAView: Office.context.mailbox.diagnostics.OWAView
  };

  console.log("All Item Properties:", properties);

  logEvent("Info", "All Properties Retrieved", properties);
}

// Log event to console and UI
function logEvent(type, message, data = null, cssClass = "") {
  const timestamp = new Date().toISOString();

  // Console logging with detailed information
  console.log(`[${timestamp}] [${type}] ${message}`);
  if (data) {
    console.log("Event Data:", data);
  }

  // Create log entry
  const logEntry = {
    timestamp,
    type,
    message,
    data: data ? JSON.stringify(data, null, 2) : null
  };

  eventLog.push(logEntry);

  // Update UI
  addLogEntryToUI(logEntry, cssClass);

  // Update statistics
  eventCount.total++;
  if (type === "Send") eventCount.send++;
  if (type === "Change") eventCount.change++;
  if (type === "Action") eventCount.action++;

  updateStatistics();
}

// Add log entry to UI
function addLogEntryToUI(logEntry, cssClass = "") {
  const container = document.getElementById('eventLogContainer');

  // Remove placeholder if it exists
  const placeholder = container.querySelector('.placeholder');
  if (placeholder) {
    placeholder.remove();
  }

  // Create log entry element
  const entryDiv = document.createElement('div');
  entryDiv.className = `log-entry ${cssClass}`;

  const timestampDiv = document.createElement('div');
  timestampDiv.className = 'log-timestamp';
  timestampDiv.textContent = new Date(logEntry.timestamp).toLocaleString();

  const typeDiv = document.createElement('div');
  typeDiv.className = 'log-type';
  typeDiv.textContent = `[${logEntry.type}] ${logEntry.message}`;

  entryDiv.appendChild(timestampDiv);
  entryDiv.appendChild(typeDiv);

  if (logEntry.data) {
    const detailsDiv = document.createElement('div');
    detailsDiv.className = 'log-details';
    detailsDiv.textContent = logEntry.data;
    entryDiv.appendChild(detailsDiv);
  }

  // Add to top of log
  container.insertBefore(entryDiv, container.firstChild);

  // Limit log entries to 50
  while (container.children.length > 50) {
    container.removeChild(container.lastChild);
  }
}

// Clear event log
function clearEventLog() {
  console.log("Clearing event log...");

  eventLog = [];

  const container = document.getElementById('eventLogContainer');
  container.innerHTML = '<p class="placeholder">Event log cleared. Waiting for events...</p>';

  logEvent("System", "Event Log Cleared");
}

// Update statistics display
function updateStatistics() {
  document.getElementById('totalEvents').textContent = eventCount.total;
  document.getElementById('sendEvents').textContent = eventCount.send;
  document.getElementById('changeEvents').textContent = eventCount.change;
  document.getElementById('userActions').textContent = eventCount.action;
}

// Listen for messages from event handlers
if (window.addEventListener) {
  window.addEventListener("message", (event) => {
    console.log("Received message from event handler:", event.data);

    if (event.data && event.data.type === "event-fired") {
      const eventInfo = event.data.eventInfo;

      logEvent(
        "LaunchEvent",
        `Event: ${eventInfo.eventType}`,
        eventInfo,
        "event-" + eventInfo.category
      );
    }
  }, false);
}

// Export functions for use in other scripts
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    logEvent,
    loadCurrentItem,
    trackAction
  };
}