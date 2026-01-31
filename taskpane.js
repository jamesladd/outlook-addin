/* global Office, Queue */

(function () {
  'use strict';

  let eventCounter = 0;
  let isMonitoring = true; // Start as true - monitoring starts immediately
  let monitoringInterval = null;
  let previousItemState = null;
  const activityQueue = new Queue({ results: [], concurrency: 1 });

  // Initialize Office
  Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      document.getElementById('clear-log').onclick = clearActivityLog;
      document.getElementById('toggle-monitoring').onclick = toggleMonitoring;

      initializeTaskpane();
    }
  });

  function initializeTaskpane() {
    logActivity('info', 'InboxAgent taskpane initialized');

    // Check if event handlers are available (they should be loaded in commands.html runtime)
    console.log('=== CHECKING EVENT HANDLER AVAILABILITY ===');
    console.log('Window handlers check:');
    console.log('  - onNewMessageComposeHandler:', typeof window.onNewMessageComposeHandler);
    console.log('  - onMessageSendHandler:', typeof window.onMessageSendHandler);

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

  function addOfficeEventListeners() {
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
      // Try to add various handlers
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
  }

  function onItemChanged(args) {
    console.log('=== ITEM CHANGED EVENT FIRED ===');
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
    console.log(`=== ${eventType.toUpperCase()} EVENT FIRED ===`);
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
        item.subject.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const subject = result.value || '(No Subject)';
            document.getElementById('current-item').textContent =
              subject.substring(0, 30) + (subject.length > 30 ? '...' : '');

            console.log('=== CURRENT ITEM UPDATED ===');
            console.log('Subject:', subject);
            console.log('Item Type:', item.itemType);
            console.log('Item ID:', item.itemId || 'New item (no ID)');
          }
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
      item.subject.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          state.subject = result.value;
        }
        cb();
      });
    });

    // Capture categories
    if (item.categories) {
      captureQueue.push(cb => {
        item.categories.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            state.categories = result.value;
          }
          cb();
        });
      });
    }

    // Capture internet message id (for tracking replies/forwards)
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

    // Capture from (read mode)
    if (item.from) {
      captureQueue.push(cb => {
        item.from.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            state.from = result.value;
          }
          cb();
        });
      });
    }

    // Capture to recipients
    if (item.to) {
      captureQueue.push(cb => {
        item.to.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            state.to = result.value;
          }
          cb();
        });
      });
    }

    // Capture cc recipients
    if (item.cc) {
      captureQueue.push(cb => {
        item.cc.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            state.cc = result.value;
          }
          cb();
        });
      });
    }

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
      item.subject.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          currentState.subject = result.value;
        }
        cb();
      });
    });

    if (item.categories) {
      checkQueue.push(cb => {
        item.categories.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            currentState.categories = result.value;
          }
          cb();
        });
      });
    }

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

    if (item.from) {
      checkQueue.push(cb => {
        item.from.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            currentState.from = result.value;
          }
          cb();
        });
      });
    }

    if (item.to) {
      checkQueue.push(cb => {
        item.to.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            currentState.to = result.value;
          }
          cb();
        });
      });
    }

    if (item.cc) {
      checkQueue.push(cb => {
        item.cc.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            currentState.cc = result.value;
          }
          cb();
        });
      });
    }

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

})();