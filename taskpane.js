/* global Office */

// Global event counter and storage
let eventCounter = 0;
let activeListeners = 0;
let eventHistory = [];
const STORAGE_KEY = 'InboxAgent_Events';

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

    // Setup Office event listeners
    setupOfficeEventListeners();

    // Load stored events from previous sessions
    loadStoredEvents();

    // Log that taskpane is ready
    logEvent('TaskpaneReady', 'Taskpane has been initialized and is ready', {
      host: info.host,
      platform: info.platform
    });
  }
});

// Load stored events from roaming settings
function loadStoredEvents() {
  try {
    const storedEventsJson = Office.context.roamingSettings.get(STORAGE_KEY);

    if (storedEventsJson) {
      let storedEvents = [];
      try {
        storedEvents = JSON.parse(storedEventsJson);
      } catch (e) {
        console.error('Error parsing stored events:', e);
        storedEvents = [];
      }

      console.log(`📦 Loading ${storedEvents.length} stored events from previous sessions...`);

      // Display stored events in the taskpane
      storedEvents.forEach((event) => {
        // Add action class if present
        let actionClass = '';
        if (event.action === 'REPLY') actionClass = 'reply';
        else if (event.action === 'REPLY_ALL') actionClass = 'reply-all';
        else if (event.action === 'FORWARD') actionClass = 'forward';
        else if (event.action && event.action.includes('SEND')) actionClass = 'send';

        displayEventInUI(event.type, event.description, event.details || {}, actionClass, event.timestamp);
      });

      if (storedEvents.length > 0) {
        console.log('✓ Stored events loaded successfully');
        logEvent('StoredEventsLoaded', `Loaded ${storedEvents.length} events from storage`, {
          eventCount: storedEvents.length
        });
      }
    }
  } catch (error) {
    console.error('Error loading stored events:', error);
  }
}

// Setup UI Event Listeners
function setupUIListeners() {
  // Close reminder button
  document.getElementById('closeReminder').addEventListener('click', () => {
    document.getElementById('pinReminder').classList.add('hidden');
  });

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

// Setup Office Event Listeners
function setupOfficeEventListeners() {
  try {
    const item = Office.context.mailbox.item;

    if (!item) {
      console.warn('⚠️ No item available to setup event listeners');
      return;
    }

    // ItemChanged event (when user switches items while taskpane is open)
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

    // Check if we're in compose mode
    const isComposeMode = !item.itemId;

    if (isComposeMode) {
      console.log('📝 Compose mode detected - registering compose-specific listeners');

      // RecipientsChanged event
      if (item.to && typeof item.to.addHandlerAsync === 'function') {
        item.to.addHandlerAsync(
          Office.EventType.RecipientsChanged,
          onRecipientsChanged,
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log('✓ To RecipientsChanged event listener registered');
              activeListeners++;
              updateActiveListeners();
            }
          }
        );
      }

      // CC recipients
      if (item.cc && typeof item.cc.addHandlerAsync === 'function') {
        item.cc.addHandlerAsync(
          Office.EventType.RecipientsChanged,
          onCcChanged,
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log('✓ CC RecipientsChanged event listener registered');
              activeListeners++;
              updateActiveListeners();
            }
          }
        );
      }

      // BCC recipients
      if (item.bcc && typeof item.bcc.addHandlerAsync === 'function') {
        item.bcc.addHandlerAsync(
          Office.EventType.RecipientsChanged,
          onBccChanged,
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log('✓ BCC RecipientsChanged event listener registered');
              activeListeners++;
              updateActiveListeners();
            }
          }
        );
      }

      // AttachmentsChanged event
      if (item.addHandlerAsync) {
        item.addHandlerAsync(
          Office.EventType.AttachmentsChanged,
          onAttachmentsChanged,
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log('✓ AttachmentsChanged event listener registered');
              activeListeners++;
              updateActiveListeners();
            }
          }
        );
      }

      // EnhancedLocationsChanged event (for appointments)
      if (item.itemType === Office.MailboxEnums.ItemType.Appointment &&
        item.enhancedLocation &&
        typeof item.enhancedLocation.addHandlerAsync === 'function') {
        item.enhancedLocation.addHandlerAsync(
          Office.EventType.EnhancedLocationsChanged,
          onEnhancedLocationsChanged,
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log('✓ EnhancedLocationsChanged event listener registered');
              activeListeners++;
              updateActiveListeners();
            }
          }
        );
      }

      // RecurrenceChanged event (for appointments)
      if (item.itemType === Office.MailboxEnums.ItemType.Appointment &&
        item.recurrence &&
        typeof item.recurrence.addHandlerAsync === 'function') {
        item.recurrence.addHandlerAsync(
          Office.EventType.RecurrenceChanged,
          onRecurrenceChanged,
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log('✓ RecurrenceChanged event listener registered');
              activeListeners++;
              updateActiveListeners();
            }
          }
        );
      }
    } else {
      console.log('📖 Read mode detected - limited event listeners available');
    }

    console.log(`✓ Office event listeners setup complete. Total active: ${activeListeners}`);

  } catch (error) {
    console.error('Error setting up Office event listeners:', error);
    logEvent('Error', 'Failed to setup Office event listeners', {
      error: error.message,
      stack: error.stack
    });
  }
}

// Event Handlers
function onItemChanged(eventArgs) {
  console.log('%c[EVENT] ItemChanged', 'color: #10b981; font-weight: bold;', eventArgs);

  logEvent('ItemChanged', 'User switched to a different item', {
    eventType: eventArgs.type
  });

  // Reload item info and re-setup listeners for new item
  loadItemInfo();

  // Reset and re-setup event listeners for the new item
  activeListeners = 1; // Keep ItemChanged listener
  updateActiveListeners();
  setupOfficeEventListeners();
}

function onRecipientsChanged(eventArgs) {
  console.log('%c[EVENT] RecipientsChanged (To)', 'color: #10b981; font-weight: bold;', eventArgs);

  Office.context.mailbox.item.to.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      logEvent('RecipientsChanged', 'To recipients have been modified', {
        eventType: eventArgs.type,
        recipientCount: asyncResult.value.length,
        recipients: asyncResult.value.map(r => r.emailAddress)
      });
    }
  });
}

function onCcChanged(eventArgs) {
  console.log('%c[EVENT] RecipientsChanged (CC)', 'color: #10b981; font-weight: bold;', eventArgs);

  Office.context.mailbox.item.cc.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      logEvent('RecipientsChanged', 'CC recipients have been modified', {
        eventType: eventArgs.type,
        recipientType: 'CC',
        recipientCount: asyncResult.value.length,
        recipients: asyncResult.value.map(r => r.emailAddress)
      });
    }
  });
}

function onBccChanged(eventArgs) {
  console.log('%c[EVENT] RecipientsChanged (BCC)', 'color: #10b981; font-weight: bold;', eventArgs);

  Office.context.mailbox.item.bcc.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      logEvent('RecipientsChanged', 'BCC recipients have been modified', {
        eventType: eventArgs.type,
        recipientType: 'BCC',
        recipientCount: asyncResult.value.length
      });
    }
  });
}

function onAttachmentsChanged(eventArgs) {
  console.log('%c[EVENT] AttachmentsChanged', 'color: #10b981; font-weight: bold;', eventArgs);

  Office.context.mailbox.item.getAttachmentsAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      logEvent('AttachmentsChanged', 'Attachments have been added or removed', {
        eventType: eventArgs.type,
        attachmentCount: asyncResult.value.length,
        attachments: asyncResult.value.map(a => ({
          name: a.name,
          size: a.size,
          type: a.attachmentType
        }))
      });
    }
  });
}

function onEnhancedLocationsChanged(eventArgs) {
  console.log('%c[EVENT] EnhancedLocationsChanged', 'color: #10b981; font-weight: bold;', eventArgs);

  Office.context.mailbox.item.enhancedLocation.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      logEvent('EnhancedLocationsChanged', 'Appointment location has been modified', {
        eventType: eventArgs.type,
        locations: asyncResult.value
      });
    }
  });
}

function onRecurrenceChanged(eventArgs) {
  console.log('%c[EVENT] RecurrenceChanged', 'color: #10b981; font-weight: bold;', eventArgs);

  Office.context.mailbox.item.recurrence.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      logEvent('RecurrenceChanged', 'Appointment recurrence has been modified', {
        eventType: eventArgs.type,
        recurrence: asyncResult.value
      });
    }
  });
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

  // Mode
  const mode = item.itemId ? 'Read' : 'Compose';
  document.getElementById('itemMode').textContent = mode;

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

  // Item ID
  document.getElementById('itemId').textContent = item.itemId ? item.itemId.substring(0, 30) + '...' : 'New item';

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
    dateTimeCreated: item.dateTimeCreated,
    dateTimeModified: item.dateTimeModified,
    mode: item.itemId ? 'Read' : 'Compose'
  };

  // Get async properties
  if (item.subject) {
    if (typeof item.subject === 'string') {
      properties.subject = item.subject;
      logPropertiesResult(properties);
    } else {
      item.subject.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          properties.subject = asyncResult.value;
        }
        logPropertiesResult(properties);
      });
    }
  } else {
    logPropertiesResult(properties);
  }
}

function logPropertiesResult(properties) {
  console.log('%cItem Properties:', 'color: #f59e0b; font-weight: bold;', properties);
  logEvent('PropertiesRetrieved', 'Item properties retrieved successfully', properties);
}

// Display event in UI (helper function)
function displayEventInUI(eventType, description, data, actionClass = '', timestamp = null) {
  const eventLog = document.getElementById('eventLog');
  const placeholder = eventLog.querySelector('.event-placeholder');
  if (placeholder) {
    placeholder.remove();
  }

  const eventItem = document.createElement('div');
  eventItem.className = 'event-item';
  if (actionClass) {
    eventItem.classList.add(actionClass);
  }

  const eventHeader = document.createElement('div');
  eventHeader.className = 'event-header';

  const eventTypeSpan = document.createElement('span');
  eventTypeSpan.className = 'event-type';
  eventTypeSpan.textContent = `#${eventCounter} - ${eventType}`;

  const eventTime = document.createElement('span');
  eventTime.className = 'event-time';
  const timestampEvents = document.getElementById('timestampEvents').checked;
  if (timestampEvents) {
    const timeToDisplay = timestamp ? new Date(timestamp) : new Date();
    eventTime.textContent = timeToDisplay.toLocaleTimeString();
  }

  eventHeader.appendChild(eventTypeSpan);
  eventHeader.appendChild(eventTime);

  const eventDetails = document.createElement('div');
  eventDetails.className = 'event-details';
  eventDetails.textContent = description;

  eventItem.appendChild(eventHeader);
  eventItem.appendChild(eventDetails);

  const verboseLogging = document.getElementById('verboseLogging').checked;
  if (verboseLogging && data && Object.keys(data).length > 0) {
    const eventData = document.createElement('div');
    eventData.className = 'event-data';
    eventData.textContent = JSON.stringify(data, null, 2);
    eventItem.appendChild(eventData);
  }

  eventLog.appendChild(eventItem);

  // Auto scroll
  const autoScroll = document.getElementById('autoScroll').checked;
  if (autoScroll) {
    eventLog.scrollTop = eventLog.scrollHeight;
  }
}

// Log event to UI and console
function logEvent(eventType, description, data) {
  eventCounter++;

  const timestamp = new Date().toISOString();

  // Store in history
  const eventRecord = {
    id: eventCounter,
    type: eventType,
    description: description,
    data: data,
    timestamp: timestamp,
    action: data.action || data.detectedAction || null
  };
  eventHistory.push(eventRecord);

  // Console logging
  console.log(`[${timestamp}] ${eventType}: ${description}`, data);

  // Determine action class
  let actionClass = '';
  if (data.action === 'REPLY') actionClass = 'reply';
  else if (data.action === 'REPLY_ALL') actionClass = 'reply-all';
  else if (data.action === 'FORWARD') actionClass = 'forward';
  else if (data.action && data.action.includes('SEND')) actionClass = 'send';

  // Display in UI
  displayEventInUI(eventType, description, data, actionClass);

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

  // Clear roaming settings
  Office.context.roamingSettings.remove(STORAGE_KEY);
  Office.context.roamingSettings.saveAsync();

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