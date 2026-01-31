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
  const activityQueue = new Queue({ results: [], concurrency: 1 });

  // Initialize Office
  Office.onReady((info) => {
    console.log('=== TASKPANE OFFICE.ONREADY FIRED ===');
    console.log('Timestamp:', new Date().toISOString());
    console.log('Host:', info.host);
    console.log('Platform:', info.platform);

    if (info.host === Office.HostType.Outlook) {
      document.getElementById('clear-log').onclick = clearActivityLog;
      document.getElementById('toggle-monitoring').onclick = toggleMonitoring;
      document.getElementById('trigger-test-event').onclick = triggerTestEvent;

      initializeTaskpane();
    }
  });

  function initializeTaskpane() {
    console.log('=== TASKPANE INITIALIZATION STARTED ===');
    console.log('Timestamp:', new Date().toISOString());

    logActivity('info', 'InboxAgent taskpane initialized');

    // Check for event runtime
    checkEventRuntime();

    updateCurrentItem();

    // Add Office event listeners
    addOfficeEventListeners();

    // Start deep monitoring immediately
    startDeepMonitoring();
    logActivity('success', 'Deep monitoring started automatically');

    console.log('=== INBOXAGENT TASKPANE INITIALIZED ===');
    console.log('Timestamp:', new Date().toISOString());
    console.log('Office Host:', Office.context.mailbox.diagnostics.hostName);
    console.log('Office Version:', Office.context.mailbox.diagnostics.hostVersion);
    console.log('Deep Monitoring: ACTIVE');
  }

  function checkEventRuntime() {
    console.log('=== CHECKING EVENT RUNTIME ===');
    console.log('Timestamp:', new Date().toISOString());

    const runtimeStatus = document.getElementById('runtime-status');

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
      callback(null);
      return;
    }

    const property = item[propertyName];

    if (!property) {
      callback(null);
      return;
    }

    // Check if it's a compose mode property (has getAsync)
    if (typeof property.getAsync === 'function') {
      property.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          callback(result.value);
        } else {
          callback(null);
        }
      });
    } else {
      // Read mode - direct property access
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

      testQueue.start();
    } else {
      console.log('No item currently selected');
      logActivity('warning', 'No item currently selected');
    }
  }

  function addOfficeEventListeners() {
    console.log('=== ADDING OFFICE EVENT LISTENERS IN TASKPANE ===');
    console.log('Timestamp:', new Date().toISOString());

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
    const item = Office.context.mailbox.item;
    if (item) {
      activityQueue.push(cb => {
        getPropertyValue(item, 'subject', (subject) => {
          const displaySubject = subject || '(No Subject)';
          document.getElementById('current-item').textContent =
            displaySubject.substring(0, 30) + (displaySubject.length > 30 ? '...' : '');

          console.log('=== CURRENT ITEM UPDATED ===');
          console.log('Subject:', displaySubject);
          console.log('Item Type:', item.itemType);
          console.log('Item ID:', item.itemId || 'New item (no ID)');
          cb();
        });
      });

      activityQueue.start();
    }
  }

  function toggleMonitoring() {
    isMonitoring = !isMonitoring;
    const button = document.getElementById('toggle-monitoring');
    const statusElement = document.getElementById('monitoring-status');

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
    captureCurrentItemState();

    // Poll for changes every 2 seconds
    monitoringInterval = setInterval(() => {
      checkForItemChanges();
    }, 2000);

    console.log('=== DEEP MONITORING STARTED ===');
    console.log('Polling Interval: 2000ms');
    console.log('Timestamp:', new Date().toISOString());
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
    if (!item) return;

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

    captureQueue.start();
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

    checkQueue.start();
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
    const activityLog = document.getElementById('activity-log');
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
  }

  function clearActivityLog() {
    const activityLog = document.getElementById('activity-log');
    activityLog.innerHTML = '';
    logActivity('info', 'Activity log cleared');
    console.log('=== ACTIVITY LOG CLEARED ===');
    console.log('Timestamp:', new Date().toISOString());
  }

  function updateEventCounter() {
    document.getElementById('events-tracked').textContent = eventCounter;
  }

  console.log('=== TASKPANE.JS FULLY LOADED (IIFE END) ===');
  console.log('Timestamp:', new Date().toISOString());

})();