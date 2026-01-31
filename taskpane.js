/* global Office */

// Global event counter and storage
let eventCounter = 0;
let activeListeners = 0;
let eventHistory = [];

// Initialize Office Add-in
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log('%c=== InboxAgent Taskpane Initialized ===', 'color: #0078d4; font-size: 16px; font-weight: bold;');
    console.log('Host:', info.host);
    console.log('Platform:', info.platform);

    // Setup event listeners for UI
    setupUIListeners();

    // Load current item information
    loadItemInfo();

    // Setup Office event listeners (taskpane-specific)
    setupOfficeEventListeners();

    // Log that taskpane is ready
    logEvent('TaskpaneReady', 'Taskpane has been initialized and is ready', {
      host: info.host,
      platform: info.platform,
      isPinned: isPinned()
    });
  }
});

// Check if taskpane is pinned
function isPinned() {
  // Note: There's no direct API to check if pinned, this is a placeholder
  return 'Pinning supported - check manually in Outlook';
}

// Setup UI Event Listeners
function setupUIListeners() {
  // Refresh button
  document.getElementById('refreshBtn').addEventListener('click', () => {
    logEvent('ButtonClick', 'Refresh button clicked', {});
    loadItemInfo();
  });

  // Test event button
  document.getElementById('testEventBtn').addEventListener('click', () => {
    logEvent('ButtonClick', 'Test Event button clicked', {});
    logEvent('TestEvent', 'This is a manually triggered test event', {
      timestamp: new Date().toISOString(),
      testData: 'Sample test data',
      random: Math.random()
    });
  });

  // Get properties button
  document.getElementById('getPropertiesBtn').addEventListener('click', () => {
    logEvent('ButtonClick', 'Get Properties button clicked', {});
    getItemProperties();
  });

  // Clear events button
  document.getElementById('clearEventsBtn').addEventListener('click', () => {
    clearEvents();
  });

  // Export events button
  document.getElementById('exportEventsBtn').addEventListener('click', () => {
    exportEvents();
  });

  console.log('✓ UI Event listeners configured');
}

// Setup Office Event Listeners (Taskpane-specific only)
function setupOfficeEventListeners() {
  try {
    if (Office.context.mailbox.item) {
      // ItemChanged event (when user switches items while taskpane is open)
      // This is the only event that can be registered from the taskpane
      if (Office.context.mailbox.addHandlerAsync) {
        Office.context.mailbox.addHandlerAsync(
          Office.EventType.ItemChanged,
          onItemChanged,
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log('✓ ItemChanged event listener registered');
              activeListeners++;
              updateActiveListeners();
            } else {
              console.error('✗ Failed to register ItemChanged listener:', asyncResult.error);
            }
          }
        );
      }

      // SelectedItemsChanged event (for read mode - mailbox level)
      if (Office.context.mailbox.addHandlerAsync && Office.EventType.SelectedItemsChanged) {
        Office.context.mailbox.addHandlerAsync(
          Office.EventType.SelectedItemsChanged,
          onSelectedItemsChanged,
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log('✓ SelectedItemsChanged event listener registered');
              activeListeners++;
              updateActiveListeners();
            } else {
              console.log('ℹ SelectedItemsChanged not available in this context');
            }
          }
        );
      }
    }

    // Note: Other events like RecipientsChanged, AttachmentsChanged, etc.
    // are handled via LaunchEvent in the manifest and launchevent.js
    // They cannot be registered from the taskpane directly

    console.log(`✓ Taskpane event listeners setup complete. Active listeners: ${activeListeners}`);
    console.log('ℹ Other events (Recipients, Attachments, Send, etc.) are handled via LaunchEvent in manifest');

  } catch (error) {
    console.error('Error setting up Office event listeners:', error);
    logEvent('Error', 'Failed to setup Office event listeners', { error: error.message });
  }
}

// Event Handlers
function onItemChanged(eventArgs) {
  console.log('%c[EVENT] ItemChanged', 'color: #10b981; font-weight: bold;', eventArgs);

  logEvent('ItemChanged', 'User switched to a different item', {
    eventType: eventArgs.type,
    eventArgs: JSON.stringify(eventArgs, null, 2)
  });

  // Reload item info
  loadItemInfo();
}

function onSelectedItemsChanged(eventArgs) {
  console.log('%c[EVENT] SelectedItemsChanged', 'color: #10b981; font-weight: bold;', eventArgs);

  logEvent('SelectedItemsChanged', 'Selected items in the mailbox changed', {
    eventType: eventArgs.type,
    eventArgs: JSON.stringify(eventArgs, null, 2)
  });

  // Reload item info
  loadItemInfo();
}

// Load current item information
function loadItemInfo() {
  const item = Office.context.mailbox.item;

  if (!item) {
    document.getElementById('itemType').textContent = 'No item selected';
    document.getElementById('itemSubject').textContent = 'N/A';
    document.getElementById('itemMode').textContent = 'N/A';
    document.getElementById('itemId').textContent = 'N/A';
    return;
  }

  // Item Type
  const itemType = item.itemType === Office.MailboxEnums.ItemType.Message ? 'Message' : 'Appointment';
  document.getElementById('itemType').textContent = itemType;

  // Subject
  if (item.subject) {
    if (typeof item.subject === 'string') {
      document.getElementById('itemSubject').textContent = item.subject || '(No subject)';
    } else {
      item.subject.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          document.getElementById('itemSubject').textContent = asyncResult.value || '(No subject)';
        }
      });
    }
  }

  // Mode (Read vs Compose)
  const isCompose = item.itemId === null || item.itemId === undefined;
  const mode = isCompose ? 'Compose' : 'Read';
  document.getElementById('itemMode').textContent = mode;

  // Item ID
  const itemId = item.itemId || 'New item (no ID yet)';
  document.getElementById('itemId').textContent = itemId.length > 30 ? itemId.substring(0, 30) + '...' : itemId;

  console.log('✓ Item information loaded');
}

// Get detailed item properties
function getItemProperties() {
  const item = Office.context.mailbox.item;

  if (!item) {
    logEvent('Error', 'No item available to get properties', {});
    return;
  }

  const properties = {
    itemType: item.itemType,
    itemClass: item.itemClass,
    itemId: item.itemId,
    conversationId: item.conversationId,
    dateTimeCreated: item.dateTimeCreated ? item.dateTimeCreated.toISOString() : 'N/A',
    dateTimeModified: item.dateTimeModified ? item.dateTimeModified.toISOString() : 'N/A',
    normalizedSubject: item.normalizedSubject
  };

  // Check if we're in compose or read mode
  const isCompose = item.itemId === null || item.itemId === undefined;
  properties.mode = isCompose ? 'Compose' : 'Read';

  // Add compose-specific properties
  if (isCompose) {
    // Get async properties for compose mode
    const asyncCallsNeeded = [];

    if (item.subject && item.subject.getAsync) {
      asyncCallsNeeded.push(
        new Promise((resolve) => {
          item.subject.getAsync((result) => {
            properties.subject = result.status === Office.AsyncResultStatus.Succeeded ? result.value : 'N/A';
            resolve();
          });
        })
      );
    }

    if (item.to && item.to.getAsync) {
      asyncCallsNeeded.push(
        new Promise((resolve) => {
          item.to.getAsync((result) => {
            properties.toRecipients = result.status === Office.AsyncResultStatus.Succeeded
              ? result.value.map(r => r.emailAddress)
              : [];
            resolve();
          });
        })
      );
    }

    if (item.cc && item.cc.getAsync) {
      asyncCallsNeeded.push(
        new Promise((resolve) => {
          item.cc.getAsync((result) => {
            properties.ccRecipients = result.status === Office.AsyncResultStatus.Succeeded
              ? result.value.map(r => r.emailAddress)
              : [];
            resolve();
          });
        })
      );
    }

    Promise.all(asyncCallsNeeded).then(() => {
      logPropertiesResult(properties);
    });
  } else {
    // Read mode - properties are directly accessible
    properties.subject = item.subject;
    properties.from = item.from ? item.from.emailAddress : 'N/A';
    properties.to = item.to ? item.to.map(r => r.emailAddress) : [];
    properties.cc = item.cc ? item.cc.map(r => r.emailAddress) : [];

    logPropertiesResult(properties);
  }
}

function logPropertiesResult(properties) {
  console.log('%cItem Properties:', 'color: #f59e0b; font-weight: bold;', properties);
  logEvent('PropertiesRetrieved', 'Item properties retrieved successfully', properties);
}

// Log event to UI and console
function logEvent(eventType, description, data) {
  eventCounter++;

  const timestamp = new Date().toISOString();
  const verboseLogging = document.getElementById('verboseLogging').checked;
  const timestampEvents = document.getElementById('timestampEvents').checked;
  const autoScroll = document.getElementById('autoScroll').checked;

  // Store in history
  const eventRecord = {
    id: eventCounter,
    type: eventType,
    description: description,
    data: data,
    timestamp: timestamp
  };
  eventHistory.push(eventRecord);

  // Console logging
  console.log(`[${timestamp}] ${eventType}: ${description}`, data);

  // Update UI
  const eventLog = document.getElementById('eventLog');
  const placeholder = eventLog.querySelector('.event-placeholder');
  if (placeholder) {
    placeholder.remove();
  }

  const eventItem = document.createElement('div');
  eventItem.className = 'event-item';

  const eventHeader = document.createElement('div');
  eventHeader.className = 'event-header';

  const eventTypeSpan = document.createElement('span');
  eventTypeSpan.className = 'event-type';
  eventTypeSpan.textContent = `#${eventCounter} - ${eventType}`;

  const eventTime = document.createElement('span');
  eventTime.className = 'event-time';
  if (timestampEvents) {
    eventTime.textContent = new Date(timestamp).toLocaleTimeString();
  }

  eventHeader.appendChild(eventTypeSpan);
  eventHeader.appendChild(eventTime);

  const eventDetails = document.createElement('div');
  eventDetails.className = 'event-details';
  eventDetails.textContent = description;

  eventItem.appendChild(eventHeader);
  eventItem.appendChild(eventDetails);

  if (verboseLogging && Object.keys(data).length > 0) {
    const eventData = document.createElement('div');
    eventData.className = 'event-data';
    eventData.textContent = JSON.stringify(data, null, 2);
    eventItem.appendChild(eventData);
  }

  eventLog.appendChild(eventItem);

  // Auto scroll
  if (autoScroll) {
    eventLog.scrollTop = eventLog.scrollHeight;
  }

  // Update counter
  document.getElementById('totalEvents').textContent = eventCounter;
}

// Update active listeners count
function updateActiveListeners() {
  document.getElementById('activeListeners').textContent = activeListeners;
}

// Clear events
function clearEvents() {
  eventCounter = 0;
  eventHistory = [];

  const eventLog = document.getElementById('eventLog');
  eventLog.innerHTML = `
        <div class="event-placeholder">
            <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                <path d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" stroke-width="2" stroke-linecap="round"/>
            </svg>
            <p>Waiting for events...</p>
            <small>Events will appear here as they occur</small>
        </div>
    `;

  document.getElementById('totalEvents').textContent = '0';

  console.log('%cEvents cleared', 'color: #ef4444; font-weight: bold;');
}

// Export events
function exportEvents() {
  if (eventHistory.length === 0) {
    alert('No events to export');
    return;
  }

  const exportData = {
    exportDate: new Date().toISOString(),
    totalEvents: eventHistory.length,
    events: eventHistory
  };

  const dataStr = JSON.stringify(exportData, null, 2);
  const dataBlob = new Blob([dataStr], { type: 'application/json' });
  const url = URL.createObjectURL(dataBlob);

  const link = document.createElement('a');
  link.href = url;
  link.download = `inboxagent-events-${Date.now()}.json`;
  link.click();

  URL.revokeObjectURL(url);

  console.log('%cEvents exported', 'color: #10b981; font-weight: bold;', exportData);
  logEvent('EventsExported', 'Event history exported to JSON file', {
    totalEvents: eventHistory.length
  });
}

// Listen for messages from event handlers (if using shared runtime)
if (window.addEventListener) {
  window.addEventListener('message', (event) => {
    if (event.data && event.data.type === 'LaunchEvent') {
      console.log('%c[MESSAGE] Received event from LaunchEvent handler', 'color: #8b5cf6; font-weight: bold;', event.data);

      logEvent(
        event.data.eventName || 'LaunchEvent',
        event.data.description || 'Event triggered from LaunchEvent handler',
        event.data.data || {}
      );
    }
  });
}