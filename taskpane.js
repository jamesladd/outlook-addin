/* global Office */

let eventLog = [];
const MAX_LOG_ENTRIES = 50;

Office.onReady((info) => {
  console.log('='.repeat(80));
  console.log('INBOXAGENT: Office.onReady fired');
  console.log('Host:', info.host);
  console.log('Platform:', info.platform);
  console.log('='.repeat(80));

  if (info.host === Office.HostType.Outlook) {
    initializeTaskpane();
    setupMailboxEventListeners();
  }
});

function initializeTaskpane() {
  logEvent('INFO', 'Taskpane Initialized', 'InboxAgent taskpane has been loaded successfully');

  if (Office.context.mailbox.item) {
    updateItemInfo();
    monitorItemChanges();
  } else {
    document.getElementById('mode').textContent = 'No item selected';
    logEvent('WARNING', 'No Active Item', 'No email or appointment is currently selected');
  }
}

function updateItemInfo() {
  const item = Office.context.mailbox.item;

  if (!item) return;

  const mode = item.itemType === Office.MailboxEnums.ItemType.Message ?
    (item.mode === Office.MailboxEnums.ItemMode.Read ? 'Read Mode' : 'Compose Mode') :
    'Appointment';

  document.getElementById('mode').textContent = mode;
  document.getElementById('itemType').textContent = item.itemType || '—';

  // Get subject
  if (item.subject) {
    if (typeof item.subject === 'string') {
      document.getElementById('subject').textContent = item.subject;
      logEvent('INFO', 'Item Loaded', `Subject: ${item.subject}`);
    } else if (item.subject.getAsync) {
      item.subject.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          document.getElementById('subject').textContent = result.value || '(No subject)';
          logEvent('INFO', 'Item Loaded', `Subject: ${result.value || '(No subject)'}`);
        }
      });
    }
  }

  // Get from address (read mode only)
  if (item.from && item.from.getAsync) {
    item.from.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const from = result.value;
        const fromText = from ? `${from.displayName} <${from.emailAddress}>` : '—';
        document.getElementById('from').textContent = fromText;

        console.log('='.repeat(80));
        console.log('ITEM DETAILS:');
        console.log('From:', fromText);
        console.log('Item Type:', item.itemType);
        console.log('Item Mode:', mode);
        console.log('Item ID:', item.itemId);
        console.log('Conversation ID:', item.conversationId);
        console.log('='.repeat(80));
      }
    });
  } else {
    document.getElementById('from').textContent = '—';
  }
}

function setupMailboxEventListeners() {
  console.log('='.repeat(80));
  console.log('SETTING UP MAILBOX EVENT LISTENERS');
  console.log('='.repeat(80));

  const item = Office.context.mailbox.item;

  if (!item) {
    console.warn('No item available to attach event listeners');
    return;
  }

  // Item Changed Event
  try {
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      onItemChanged,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log('✓ ItemChanged event listener added successfully');
          logEvent('SUCCESS', 'Event Listener Added', 'ItemChanged event listener is now active');
        } else {
          console.error('✗ Failed to add ItemChanged listener:', result.error);
        }
      }
    );
  } catch (error) {
    console.error('Error adding ItemChanged listener:', error);
  }

  // Recipients Changed Event (Compose mode)
  if (item.addHandlerAsync && item.to) {
    try {
      item.to.addHandlerAsync(
        Office.EventType.RecipientsChanged,
        onRecipientsChanged,
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log('✓ RecipientsChanged (TO) event listener added');
            logEvent('SUCCESS', 'Event Listener Added', 'RecipientsChanged (TO) listener active');
          }
        }
      );
    } catch (error) {
      console.error('Error adding TO RecipientsChanged listener:', error);
    }
  }

  if (item.addHandlerAsync && item.cc) {
    try {
      item.cc.addHandlerAsync(
        Office.EventType.RecipientsChanged,
        onRecipientsChanged,
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log('✓ RecipientsChanged (CC) event listener added');
            logEvent('SUCCESS', 'Event Listener Added', 'RecipientsChanged (CC) listener active');
          }
        }
      );
    } catch (error) {
      console.error('Error adding CC RecipientsChanged listener:', error);
    }
  }

  if (item.addHandlerAsync && item.bcc) {
    try {
      item.bcc.addHandlerAsync(
        Office.EventType.RecipientsChanged,
        onRecipientsChanged,
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log('✓ RecipientsChanged (BCC) event listener added');
            logEvent('SUCCESS', 'Event Listener Added', 'RecipientsChanged (BCC) listener active');
          }
        }
      );
    } catch (error) {
      console.error('Error adding BCC RecipientsChanged listener:', error);
    }
  }

  // Attachments Changed Event
  if (item.addHandlerAsync && item.attachments) {
    try {
      item.addHandlerAsync(
        Office.EventType.AttachmentsChanged,
        onAttachmentsChanged,
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log('✓ AttachmentsChanged event listener added');
            logEvent('SUCCESS', 'Event Listener Added', 'AttachmentsChanged listener active');
          }
        }
      );
    } catch (error) {
      console.error('Error adding AttachmentsChanged listener:', error);
    }
  }

  // Enhanced Context Changed Event
  if (item.addHandlerAsync) {
    try {
      item.addHandlerAsync(
        Office.EventType.EnhancedLocationsChanged,
        onEnhancedLocationsChanged,
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log('✓ EnhancedLocationsChanged event listener added');
            logEvent('SUCCESS', 'Event Listener Added', 'EnhancedLocationsChanged listener active');
          }
        }
      );
    } catch (error) {
      console.error('Error adding EnhancedLocationsChanged listener:', error);
    }
  }

  // Recurrence Changed Event (for appointments)
  if (item.addHandlerAsync && item.recurrence) {
    try {
      item.addHandlerAsync(
        Office.EventType.RecurrenceChanged,
        onRecurrenceChanged,
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log('✓ RecurrenceChanged event listener added');
            logEvent('SUCCESS', 'Event Listener Added', 'RecurrenceChanged listener active');
          }
        }
      );
    } catch (error) {
      console.error('Error adding RecurrenceChanged listener:', error);
    }
  }
}

// Event Handlers
function onItemChanged(eventArgs) {
  console.log('='.repeat(80));
  console.log('EVENT: ItemChanged');
  console.log('Type:', eventArgs.type);
  console.log('Timestamp:', new Date().toISOString());
  console.log('Event Args:', eventArgs);
  console.log('='.repeat(80));

  logEvent('EVENT', 'ItemChanged', 'User switched to a different email or item', eventArgs);
  updateItemInfo();
}

function onRecipientsChanged(eventArgs) {
  console.log('='.repeat(80));
  console.log('EVENT: RecipientsChanged');
  console.log('Type:', eventArgs.type);
  console.log('Recipient Type:', eventArgs.changedRecipientFields);
  console.log('Timestamp:', new Date().toISOString());
  console.log('Event Args:', eventArgs);
  console.log('='.repeat(80));

  const item = Office.context.mailbox.item;

  if (eventArgs.changedRecipientFields) {
    eventArgs.changedRecipientFields.forEach(field => {
      let recipientField;
      if (field === Office.MailboxEnums.RecipientField.To) {
        recipientField = item.to;
        logEvent('EVENT', 'RecipientsChanged', 'TO recipients have been modified', eventArgs);
      } else if (field === Office.MailboxEnums.RecipientField.Cc) {
        recipientField = item.cc;
        logEvent('EVENT', 'RecipientsChanged', 'CC recipients have been modified', eventArgs);
      } else if (field === Office.MailboxEnums.RecipientField.Bcc) {
        recipientField = item.bcc;
        logEvent('EVENT', 'RecipientsChanged', 'BCC recipients have been modified', eventArgs);
      }

      if (recipientField && recipientField.getAsync) {
        recipientField.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log('Updated Recipients:', result.value);
          }
        });
      }
    });
  }
}

function onAttachmentsChanged(eventArgs) {
  console.log('='.repeat(80));
  console.log('EVENT: AttachmentsChanged');
  console.log('Type:', eventArgs.type);
  console.log('Attachment Changes:', eventArgs.attachmentChanges);
  console.log('Attachment Status:', eventArgs.attachmentStatus);
  console.log('Timestamp:', new Date().toISOString());
  console.log('Event Args:', eventArgs);
  console.log('='.repeat(80));

  const item = Office.context.mailbox.item;
  const details = `Attachment modified. Status: ${eventArgs.attachmentStatus || 'Unknown'}`;

  logEvent('EVENT', 'AttachmentsChanged', details, eventArgs);

  if (item.attachments) {
    console.log('Current Attachments:', item.attachments);
  }
}

function onEnhancedLocationsChanged(eventArgs) {
  console.log('='.repeat(80));
  console.log('EVENT: EnhancedLocationsChanged');
  console.log('Type:', eventArgs.type);
  console.log('Timestamp:', new Date().toISOString());
  console.log('Event Args:', eventArgs);
  console.log('='.repeat(80));

  logEvent('EVENT', 'EnhancedLocationsChanged', 'Meeting location has been changed', eventArgs);
}

function onRecurrenceChanged(eventArgs) {
  console.log('='.repeat(80));
  console.log('EVENT: RecurrenceChanged');
  console.log('Type:', eventArgs.type);
  console.log('Timestamp:', new Date().toISOString());
  console.log('Event Args:', eventArgs);
  console.log('='.repeat(80));

  logEvent('EVENT', 'RecurrenceChanged', 'Appointment recurrence pattern has been modified', eventArgs);
}

// Monitor item changes for tracking actions
function monitorItemChanges() {
  const item = Office.context.mailbox.item;

  if (!item) return;

  // Monitor for body changes (could indicate reply/forward)
  if (item.body && item.body.getAsync) {
    let previousBodyLength = 0;

    setInterval(() => {
      item.body.getAsync('text', (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const currentLength = result.value.length;
          if (currentLength > previousBodyLength && previousBodyLength > 0) {
            console.log('='.repeat(80));
            console.log('POTENTIAL ACTION: Body content increased');
            console.log('Previous length:', previousBodyLength);
            console.log('Current length:', currentLength);
            console.log('Possible Reply or Forward action');
            console.log('='.repeat(80));
          }
          previousBodyLength = currentLength;
        }
      });
    }, 2000);
  }
}

// User Action Tracking Functions
function trackReplyAction() {
  console.log('='.repeat(80));
  console.log('USER ACTION: Reply Tracked');
  console.log('Timestamp:', new Date().toISOString());

  const item = Office.context.mailbox.item;

  if (item) {
    console.log('Email Details:');
    console.log('- Item ID:', item.itemId);
    console.log('- Conversation ID:', item.conversationId);
    console.log('- Item Type:', item.itemType);

    if (item.subject) {
      if (typeof item.subject === 'string') {
        console.log('- Subject:', item.subject);
        logEvent('ACTION', 'Reply Initiated', `User clicked Reply for: "${item.subject}"`, {
          itemId: item.itemId,
          conversationId: item.conversationId,
          action: 'reply'
        });
      } else if (item.subject.getAsync) {
        item.subject.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log('- Subject:', result.value);
            logEvent('ACTION', 'Reply Initiated', `User clicked Reply for: "${result.value}"`, {
              itemId: item.itemId,
              conversationId: item.conversationId,
              action: 'reply'
            });
          }
        });
      }
    }

    if (item.from && item.from.getAsync) {
      item.from.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log('- From:', result.value);
        }
      });
    }
  }

  console.log('='.repeat(80));
}

function trackForwardAction() {
  console.log('='.repeat(80));
  console.log('USER ACTION: Forward Tracked');
  console.log('Timestamp:', new Date().toISOString());

  const item = Office.context.mailbox.item;

  if (item) {
    console.log('Email Details:');
    console.log('- Item ID:', item.itemId);
    console.log('- Conversation ID:', item.conversationId);
    console.log('- Item Type:', item.itemType);

    if (item.subject) {
      if (typeof item.subject === 'string') {
        console.log('- Subject:', item.subject);
        logEvent('ACTION', 'Forward Initiated', `User clicked Forward for: "${item.subject}"`, {
          itemId: item.itemId,
          conversationId: item.conversationId,
          action: 'forward'
        });
      } else if (item.subject.getAsync) {
        item.subject.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log('- Subject:', result.value);
            logEvent('ACTION', 'Forward Initiated', `User clicked Forward for: "${result.value}"`, {
              itemId: item.itemId,
              conversationId: item.conversationId,
              action: 'forward'
            });
          }
        });
      }
    }
  }

  console.log('='.repeat(80));
}

// Event Logging System
function logEvent(type, eventName, description, eventArgs = null) {
  const timestamp = new Date().toISOString();
  const logEntry = {
    timestamp,
    type,
    eventName,
    description,
    eventArgs
  };

  eventLog.unshift(logEntry);

  if (eventLog.length > MAX_LOG_ENTRIES) {
    eventLog.pop();
  }

  console.log(`[${type}] ${eventName}: ${description}`);
  if (eventArgs) {
    console.log('Event Arguments:', eventArgs);
  }

  renderEventLog();
}

function renderEventLog() {
  const container = document.getElementById('eventLogContainer');

  if (eventLog.length === 0) {
    container.innerHTML = '<div class="loading">Waiting for events...</div>';
    return;
  }

  container.innerHTML = eventLog.map(entry => {
    const cssClass = entry.type === 'ERROR' ? 'error' :
      entry.type === 'WARNING' ? 'warning' : '';

    const time = new Date(entry.timestamp).toLocaleTimeString();

    return `
            <div class="event-item ${cssClass}">
                <div class="event-timestamp">${time}</div>
                <div class="event-type">${entry.eventName}</div>
                <div class="event-details">${entry.description}</div>
            </div>
        `;
  }).join('');
}

function clearEventLog() {
  eventLog = [];
  renderEventLog();
  console.log('='.repeat(80));
  console.log('Event log cleared');
  console.log('='.repeat(80));
  logEvent('INFO', 'Log Cleared', 'Event log has been cleared by user');
}