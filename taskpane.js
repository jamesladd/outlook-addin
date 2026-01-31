/* global Office, Queue */

// IIFE wrapper to execute immediately
(function () {
  'use strict';

  console.log('=== TASKPANE.JS LOADING (IIFE START) ===');
  console.log('Timestamp:', new Date().toISOString());

  let eventCounter = 0;
  let isMonitoring = true;
  let monitoringInterval = null;
  let previousItemState = null;
  let isInitialized = false;

  // Initialize Office
  Office.onReady((info) => {
    console.log('=== TASKPANE OFFICE.ONREADY FIRED ===');
    console.log('Timestamp:', new Date().toISOString());
    console.log('Host:', info.host);
    console.log('Platform:', info.platform);

    if (info.host === Office.HostType.Outlook) {
      // Use setTimeout to ensure DOM is ready
      setTimeout(() => {
        try {
          initializeTaskpane();
        } catch (error) {
          console.error('=== INITIALIZATION ERROR ===');
          console.error('Error:', error);
          console.error('Stack:', error.stack);
        }
      }, 100);
    }
  });

  function initializeTaskpane() {
    console.log('=== TASKPANE INITIALIZATION STARTED ===');
    console.log('Timestamp:', new Date().toISOString());

    try {
      // Verify DOM elements exist
      const clearLogBtn = document.getElementById('clear-log');
      const toggleMonitoringBtn = document.getElementById('toggle-monitoring');
      const triggerTestBtn = document.getElementById('trigger-test-event');

      console.log('DOM Elements Check:');
      console.log('  - clear-log:', clearLogBtn ? 'Found' : 'NOT FOUND');
      console.log('  - toggle-monitoring:', toggleMonitoringBtn ? 'Found' : 'NOT FOUND');
      console.log('  - trigger-test-event:', triggerTestBtn ? 'Found' : 'NOT FOUND');

      if (!clearLogBtn || !toggleMonitoringBtn || !triggerTestBtn) {
        throw new Error('Required DOM elements not found');
      }

      // Attach event handlers
      clearLogBtn.onclick = clearActivityLog;
      toggleMonitoringBtn.onclick = toggleMonitoring;
      triggerTestBtn.onclick = triggerTestEvent;

      console.log('Event handlers attached successfully');

      logActivity('info', 'InboxAgent taskpane initialized');

      // Check for event runtime
      checkEventRuntime();

      // Update current item
      updateCurrentItem();

      // Add Office event listeners
      addOfficeEventListeners();

      // Start deep monitoring immediately
      setTimeout(() => {
        startDeepMonitoring();
        logActivity('success', 'Deep monitoring started automatically');
      }, 500);

      isInitialized = true;

      console.log('=== INBOXAGENT TASKPANE INITIALIZED SUCCESSFULLY ===');
      console.log('Timestamp:', new Date().toISOString());
      console.log('Office Host:', Office.context.mailbox.diagnostics.hostName);
      console.log('Office Version:', Office.context.mailbox.diagnostics.hostVersion);
      console.log('Deep Monitoring: ACTIVE');

    } catch (error) {
      console.error('=== INITIALIZATION ERROR ===');
      console.error('Error:', error);
      console.error('Stack:', error.stack);
      logActivity('error', `Initialization failed: ${error.message}`);
    }
  }

  function checkEventRuntime() {
    console.log('=== CHECKING EVENT RUNTIME ===');
    console.log('Timestamp:', new Date().toISOString());

    const runtimeStatus = document.getElementById('runtime-status');

    if (!runtimeStatus) {
      console.error('runtime-status element not found');
      return;
    }

    if (Office.context.mailbox.item && Office.context.mailbox.addHandlerAsync) {
      runtimeStatus.textContent = 'Active';
      runtimeStatus.classList.add('active');
      logActivity('success', 'Event-based activation runtime is active');
      console.log('Event-based activation is supported');
    } else {
      runtimeStatus.textContent = 'Not Available';
      runtimeStatus.classList.add('inactive');
      logActivity('warning', 'Event-based activation not available');
      console.log('Event-based activation is NOT supported');
    }
  }

  // Helper function to get property value (handles both read and compose modes)
  function getPropertyValue(item, propertyName, callback) {
    if (!item) {
      console.log(`getPropertyValue: No item provided for ${propertyName}`);
      callback(null);
      return;
    }

    const property = item[propertyName];

    if (!property) {
      console.log(`getPropertyValue: Property ${propertyName} not found on item`);
      callback(null);
      return;
    }

    // Check if it's a compose mode property (has getAsync)
    if (typeof property.getAsync === 'function') {
      console.log(`getPropertyValue: Using getAsync for ${propertyName}`);
      try {
        property.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log(`getPropertyValue: Got value for ${propertyName}:`, result.value);
            callback(result.value);
          } else {
            console.error(`getPropertyValue: Failed to get ${propertyName}:`, result.error);
            callback(null);
          }
        });
      } catch (error) {
        console.error(`getPropertyValue: Exception getting ${propertyName}:`, error);
        callback(null);
      }
    } else {
      // Read mode - direct property access
      console.log(`getPropertyValue: Direct access for ${propertyName}:`, property);
      callback(property);
    }
  }

  function triggerTestEvent() {
    console.log('=== TRIGGERING TEST EVENT ===');
    console.log('Timestamp:', new Date().toISOString());

    logActivity('info', 'Test event triggered - check console for details');

    const item = Office.context.mailbox.item;
    if (item) {
      console.log('Current Item Details:');
      console.log('  - Item Type:', item.itemType);
      console.log('  - Item Class:', item.itemClass);
      console.log('  - Item ID:', item.itemId || 'No ID (new item)');
      console.log('  - Conversation ID:', item.conversationId);

      const testQueue = new Queue({ results: [], concurrency: 1 });

      testQueue.push(cb => {
        getPropertyValue(item, 'subject', (value) => {
          console.log('  - Subject:', value);
          logActivity('info', `Subject: ${value}`);
          cb();
        });
      });

      testQueue.push(cb => {
        getPropertyValue(item, 'from', (value) => {
          console.log('  - From:', JSON.stringify(value, null, 2));
          cb();
        });
      });

      testQueue.push(cb => {
        getPropertyValue(item, 'to', (value) => {
          console.log('  - To:', JSON.stringify(value, null, 2));
          cb();
        });
      });

      testQueue.push(cb => {
        getPropertyValue(item, 'categories', (value) => {
          console.log('  - Categories:', JSON.stringify(value, null, 2));
          logActivity('info', `Categories: ${JSON.stringify(value)}`);
          cb();
        });
      });

      if (item.attachments) {
        testQueue.push(cb => {
          console.log('  - Attachments:', item.attachments.length);
          item.attachments.forEach(att => {
            console.log(`    * ${att.name} (${att.size} bytes)`);
          });
          cb();
        });
      }

      testQueue.start((err) => {
        if (err) {
          console.error('Test queue error:', err);
        } else {
          console.log('Test queue completed successfully');
        }
      });
    } else {
      console.log('No item currently selected');
      logActivity('warning', 'No item currently selected');
    }
  }

  function addOfficeEventListeners() {
    console.log('=== ADDING OFFICE EVENT LISTENERS IN TASKPANE ===');
    console.log('Timestamp:', new Date().toISOString());

    try {
      // Item Changed Event
      if (Office.context.mailbox.addHandlerAsync) {
        Office.context.mailbox.addHandlerAsync(
          Office.EventType.ItemChanged,
          onItemChanged,
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              logActivity('success', 'ItemChanged listener attached');
              console.log('=== EVENT LISTENER ATTACHED ===');
              console.log('Event Type: ItemChanged');
              console.log('Timestamp:', new Date().toISOString());
            } else {
              logActivity('error', 'Failed to attach ItemChanged listener');
              console.error('=== EVENT LISTENER FAILED ===');
              console.error('Event Type: ItemChanged');
              console.error('Error:', result.error);
            }
          }
        );
      }

      // Recipients Changed Event (if in compose mode)
      const item = Office.context.mailbox.item;
      if (item && item.addHandlerAsync) {
        const eventTypes = [
          'RecipientsChanged',
          'RecurrenceChanged',
          'AppointmentTimeChanged'
        ];

        eventTypes.forEach(eventType => {
          if (Office.EventType[eventType]) {
            item.addHandlerAsync(
              Office.EventType[eventType],
              (args) => onItemPropertyChanged(eventType, args),
              (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                  logActivity('success', `${eventType} listener attached`);
                  console.log(`=== EVENT LISTENER ATTACHED: ${eventType} ===`);
                }
              }
            );
          }
        });
      }

      console.log('=== FINISHED ADDING OFFICE EVENT LISTENERS ===');
    } catch (error) {
      console.error('=== ERROR ADDING EVENT LISTENERS ===');
      console.error('Error:', error);
      console.error('Stack:', error.stack);
    }
  }

  function onItemChanged(args) {
    console.log('=== ITEM CHANGED EVENT FIRED (TASKPANE) ===');
    console.log('Timestamp:', new Date().toISOString());
    console.log('Event Args:', JSON.stringify(args, null, 2));

    logActivity('info', 'Item changed - Loading new item details');
    eventCounter++;
    updateEventCounter();
    updateCurrentItem();

    // Reset monitoring state for new item
    previousItemState = null;
    if (isMonitoring) {
      captureCurrentItemState();
    }
  }

  function onItemPropertyChanged(eventType, args) {
    console.log(`=== ${eventType.toUpperCase()} EVENT FIRED (TASKPANE) ===`);
    console.log('Timestamp:', new Date().toISOString());
    console.log('Event Args:', JSON.stringify(args, null, 2));

    logActivity('warning', `${eventType} detected`);
    eventCounter++;
    updateEventCounter();
  }

  function updateCurrentItem() {
    console.log('=== UPDATING CURRENT ITEM ===');

    const currentItemElement = document.getElementById('current-item');
    if (!currentItemElement) {
      console.error('current-item element not found');
      return;
    }

    const item = Office.context.mailbox.item;
    if (!item) {
      console.log('No item available');
      currentItemElement.textContent = 'No item selected';
      return;
    }

    console.log('Item available, getting subject...');

    getPropertyValue(item, 'subject', (subject) => {
      const displaySubject = subject || '(No Subject)';
      currentItemElement.textContent =
        displaySubject.substring(0, 30) + (displaySubject.length > 30 ? '...' : '');

      console.log('=== CURRENT ITEM UPDATED ===');
      console.log('Subject:', displaySubject);
      console.log('Item Type:', item.itemType);
      console.log('Item ID:', item.itemId || 'New item (no ID)');
    });
  }

  function toggleMonitoring() {
    console.log('=== TOGGLE MONITORING CLICKED ===');

    isMonitoring = !isMonitoring;
    const button = document.getElementById('toggle-monitoring');
    const statusElement = document.getElementById('monitoring-status');

    if (!button || !statusElement) {
      console.error('Button or status element not found');
      return;
    }

    if (isMonitoring) {
      button.textContent = 'Pause Monitoring';
      button.classList.remove('btn-success');
      button.classList.add('btn-warning');
      statusElement.textContent = 'Active';
      statusElement.classList.remove('paused');
      statusElement.classList.add('active');
      startDeepMonitoring();
      logActivity('success', 'Deep monitoring resumed');
    } else {
      button.textContent = 'Resume Monitoring';
      button.classList.remove('btn-warning');
      button.classList.add('btn-success');
      statusElement.textContent = 'Paused';
      statusElement.classList.remove('active');
      statusElement.classList.add('paused');
      stopDeepMonitoring();
      logActivity('warning', 'Deep monitoring paused');
    }

    console.log('=== MONITORING TOGGLED ===');
    console.log('Monitoring Active:', isMonitoring);
    console.log('Timestamp:', new Date().toISOString());
  }

  function startDeepMonitoring() {
    console.log('=== STARTING DEEP MONITORING ===');

    try {
      captureCurrentItemState();

      // Poll for changes every 2 seconds
      if (monitoringInterval) {
        clearInterval(monitoringInterval);
      }

      monitoringInterval = setInterval(() => {
        checkForItemChanges();
      }, 2000);

      console.log('=== DEEP MONITORING STARTED ===');
      console.log('Polling Interval: 2000ms');
      console.log('Timestamp:', new Date().toISOString());
    } catch (error) {
      console.error('=== ERROR STARTING MONITORING ===');
      console.error('Error:', error);
      console.error('Stack:', error.stack);
    }
  }

  function stopDeepMonitoring() {
    if (monitoringInterval) {
      clearInterval(monitoringInterval);
      monitoringInterval = null;
    }
    previousItemState = null;

    console.log('=== DEEP MONITORING STOPPED ===');
    console.log('Timestamp:', new Date().toISOString());
  }

  function captureCurrentItemState() {
    const item = Office.context.mailbox.item;
    if (!item) {
      console.log('captureCurrentItemState: No item available');
      return;
    }

    console.log('=== CAPTURING ITEM STATE ===');

    const captureQueue = new Queue({ results: [], concurrency: 1 });
    const state = {
      timestamp: new Date().toISOString(),
      itemType: item.itemType,
      itemId: item.itemId
    };

    // Capture subject
    captureQueue.push(cb => {
      getPropertyValue(item, 'subject', (value) => {
        state.subject = value;
        cb();
      });
    });

    // Capture categories
    captureQueue.push(cb => {
      getPropertyValue(item, 'categories', (value) => {
        state.categories = value;
        cb();
      });
    });

    // Capture internet message id
    if (item.internetMessageId) {
      captureQueue.push(cb => {
        state.internetMessageId = item.internetMessageId;
        cb();
      });
    }

    // Capture conversation id
    if (item.conversationId) {
      captureQueue.push(cb => {
        state.conversationId = item.conversationId;
        cb();
      });
    }

    // Capture from
    captureQueue.push(cb => {
      getPropertyValue(item, 'from', (value) => {
        state.from = value;
        cb();
      });
    });

    // Capture to recipients
    captureQueue.push(cb => {
      getPropertyValue(item, 'to', (value) => {
        state.to = value;
        cb();
      });
    });

    // Capture cc recipients
    captureQueue.push(cb => {
      getPropertyValue(item, 'cc', (value) => {
        state.cc = value;
        cb();
      });
    });

    // Capture attachments
    if (item.attachments) {
      captureQueue.push(cb => {
        state.attachments = item.attachments.map(att => ({
          id: att.id,
          name: att.name,
          size: att.size,
          attachmentType: att.attachmentType
        }));
        cb();
      });
    }

    captureQueue.push(cb => {
      previousItemState = state;
      console.log('=== ITEM STATE CAPTURED ===');
      console.log('State:', JSON.stringify(state, null, 2));
      cb();
    });

    captureQueue.start((err) => {
      if (err) {
        console.error('Capture queue error:', err);
      } else {
        console.log('Capture queue completed successfully');
      }
    });
  }

  function checkForItemChanges() {
    if (!previousItemState) {
      captureCurrentItemState();
      return;
    }

    const item = Office.context.mailbox.item;
    if (!item) return;

    const checkQueue = new Queue({ results: [], concurrency: 1 });
    const currentState = {
      timestamp: new Date().toISOString(),
      itemType: item.itemType,
      itemId: item.itemId
    };

    // Capture current state
    checkQueue.push(cb => {
      getPropertyValue(item, 'subject', (value) => {
        currentState.subject = value;
        cb();
      });
    });

    checkQueue.push(cb => {
      getPropertyValue(item, 'categories', (value) => {
        currentState.categories = value;
        cb();
      });
    });

    if (item.internetMessageId) {
      checkQueue.push(cb => {
        currentState.internetMessageId = item.internetMessageId;
        cb();
      });
    }

    if (item.conversationId) {
      checkQueue.push(cb => {
        currentState.conversationId = item.conversationId;
        cb();
      });
    }

    checkQueue.push(cb => {
      getPropertyValue(item, 'from', (value) => {
        currentState.from = value;
        cb();
      });
    });

    checkQueue.push(cb => {
      getPropertyValue(item, 'to', (value) => {
        currentState.to = value;
        cb();
      });
    });

    checkQueue.push(cb => {
      getPropertyValue(item, 'cc', (value) => {
        currentState.cc = value;
        cb();
      });
    });

    if (item.attachments) {
      checkQueue.push(cb => {
        currentState.attachments = item.attachments.map(att => ({
          id: att.id,
          name: att.name,
          size: att.size,
          attachmentType: att.attachmentType
        }));
        cb();
      });
    }

    // Compare states
    checkQueue.push(cb => {
      compareStates(previousItemState, currentState);
      previousItemState = currentState;
      cb();
    });

    checkQueue.start((err) => {
      if (err) {
        console.error('Check queue error:', err);
      }
    });
  }

  function compareStates(oldState, newState) {
    const oldJSON = JSON.stringify(oldState);
    const newJSON = JSON.stringify(newState);

    if (oldJSON !== newJSON) {
      console.log('=== ITEM STATE CHANGED ===');
      console.log('Timestamp:', new Date().toISOString());
      console.log('Previous State:', oldJSON);
      console.log('Current State:', newJSON);

      // Detailed change detection
      const changes = [];

      if (oldState.subject !== newState.subject) {
        changes.push(`Subject: "${oldState.subject}" → "${newState.subject}"`);
      }

      if (JSON.stringify(oldState.categories) !== JSON.stringify(newState.categories)) {
        changes.push(`Categories: ${JSON.stringify(oldState.categories)} → ${JSON.stringify(newState.categories)}`);
        logActivity('warning', `Categories changed`);
      }

      if (JSON.stringify(oldState.to) !== JSON.stringify(newState.to)) {
        changes.push(`To recipients changed`);
      }

      if (JSON.stringify(oldState.cc) !== JSON.stringify(newState.cc)) {
        changes.push(`CC recipients changed`);
      }

      if (JSON.stringify(oldState.attachments) !== JSON.stringify(newState.attachments)) {
        changes.push(`Attachments changed`);
      }

      if (oldState.conversationId !== newState.conversationId) {
        changes.push(`Conversation changed (possible reply/forward)`);
        detectEmailAction(oldState, newState);
      }

      changes.forEach(change => {
        console.log('Change detected:', change);
        logActivity('warning', change);
      });

      eventCounter++;
      updateEventCounter();
    }
  }

  function detectEmailAction(oldState, newState) {
    // Detect reply or forward actions
    if (oldState.conversationId && newState.conversationId) {
      if (oldState.conversationId === newState.conversationId &&
        oldState.itemId !== newState.itemId) {

        console.log('=== EMAIL ACTION DETECTED ===');
        console.log('Action Type: REPLY or FORWARD');
        console.log('Original Item ID:', oldState.itemId);
        console.log('New Item ID:', newState.itemId);
        console.log('Conversation ID:', newState.conversationId);
        console.log('Original Subject:', oldState.subject);
        console.log('New Subject:', newState.subject);

        let actionType = 'UNKNOWN';
        if (newState.subject && oldState.subject) {
          if (newState.subject.startsWith('RE:') || newState.subject.startsWith('Re:')) {
            actionType = 'REPLY';
          } else if (newState.subject.startsWith('FW:') || newState.subject.startsWith('Fw:')) {
            actionType = 'FORWARD';
          }
        }

        logActivity('success', `${actionType}: "${oldState.subject}"`);

        console.log('Detected Action:', actionType);
      }
    }
  }

  function logActivity(type, message) {
    try {
      const activityLog = document.getElementById('activity-log');
      if (!activityLog) {
        console.error('activity-log element not found');
        return;
      }

      const activityItem = document.createElement('div');
      activityItem.className = `activity-item ${type}`;

      const time = document.createElement('div');
      time.className = 'activity-time';
      time.textContent = new Date().toLocaleTimeString();

      const msg = document.createElement('div');
      msg.className = 'activity-message';
      msg.textContent = message;

      activityItem.appendChild(time);
      activityItem.appendChild(msg);

      // Insert at the top
      if (activityLog.firstChild) {
        activityLog.insertBefore(activityItem, activityLog.firstChild);
      } else {
        activityLog.appendChild(activityItem);
      }

      // Keep only last 50 items
      while (activityLog.children.length > 50) {
        activityLog.removeChild(activityLog.lastChild);
      }
    } catch (error) {
      console.error('Error logging activity:', error);
    }
  }

  function clearActivityLog() {
    console.log('=== CLEAR LOG CLICKED ===');

    try {
      const activityLog = document.getElementById('activity-log');
      if (!activityLog) {
        console.error('activity-log element not found');
        return;
      }

      activityLog.innerHTML = '';
      logActivity('info', 'Activity log cleared');
      console.log('=== ACTIVITY LOG CLEARED ===');
      console.log('Timestamp:', new Date().toISOString());
    } catch (error) {
      console.error('Error clearing log:', error);
    }
  }

  function updateEventCounter() {
    try {
      const counterElement = document.getElementById('events-tracked');
      if (counterElement) {
        counterElement.textContent = eventCounter;
      } else {
        console.error('events-tracked element not found');
      }
    } catch (error) {
      console.error('Error updating event counter:', error);
    }
  }

  console.log('=== TASKPANE.JS FULLY LOADED (IIFE END) ===');
  console.log('Timestamp:', new Date().toISOString());

})();