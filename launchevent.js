/* global Office */

// Global event tracking
const eventTracker = {
  events: [],
  statistics: {
    total: 0,
    byType: {}
  }
};

// Utility function to log events
function logEventToConsole(eventType, eventData) {
  const timestamp = new Date().toISOString();
  const logEntry = {
    timestamp,
    eventType,
    data: eventData
  };

  console.log("=".repeat(80));
  console.log(`EVENT FIRED: ${eventType}`);
  console.log(`Timestamp: ${timestamp}`);
  console.log("Event Details:");
  console.log(JSON.stringify(eventData, null, 2));
  console.log("=".repeat(80));

  // Track statistics
  eventTracker.events.push(logEntry);
  eventTracker.statistics.total++;
  eventTracker.statistics.byType[eventType] =
    (eventTracker.statistics.byType[eventType] || 0) + 1;

  console.log("Event Statistics:");
  console.log(`Total Events: ${eventTracker.statistics.total}`);
  console.log("Events by Type:", eventTracker.statistics.byType);

  return logEntry;
}

// Get detailed item information
function getItemDetails(mailboxItem) {
  const details = {
    itemType: mailboxItem.itemType,
    subject: mailboxItem.subject,
    itemId: mailboxItem.itemId,
    conversationId: mailboxItem.conversationId,
    internetMessageId: mailboxItem.internetMessageId,
    dateTimeCreated: mailboxItem.dateTimeCreated,
    dateTimeModified: mailboxItem.dateTimeModified
  };

  if (mailboxItem.itemType === Office.MailboxEnums.ItemType.Message) {
    // Message-specific details
    details.from = mailboxItem.from;
    details.sender = mailboxItem.sender;
  } else if (mailboxItem.itemType === Office.MailboxEnums.ItemType.Appointment) {
    // Appointment-specific details
    details.start = mailboxItem.start;
    details.end = mailboxItem.end;
    details.location = mailboxItem.location;
    details.organizer = mailboxItem.organizer;
  }

  return details;
}

// 1. OnNewMessageCompose Handler
function onNewMessageComposeHandler(event) {
  console.log("\n🆕 OnNewMessageCompose Event Triggered");

  const item = Office.context.mailbox.item;
  const eventData = {
    eventType: "OnNewMessageCompose",
    category: "compose",
    description: "User started composing a new message",
    itemDetails: getItemDetails(item),
    timestamp: new Date().toISOString()
  };

  logEventToConsole("OnNewMessageCompose", eventData);

  // Complete the event
  event.completed();
}

// 2. OnNewAppointmentOrganizer Handler
function onNewAppointmentOrganizerHandler(event) {
  console.log("\n📅 OnNewAppointmentOrganizer Event Triggered");

  const item = Office.context.mailbox.item;
  const eventData = {
    eventType: "OnNewAppointmentOrganizer",
    category: "compose",
    description: "User started organizing a new appointment",
    itemDetails: getItemDetails(item),
    timestamp: new Date().toISOString()
  };

  logEventToConsole("OnNewAppointmentOrganizer", eventData);

  event.completed();
}

// 3. OnMessageAttachmentsChanged Handler
function onMessageAttachmentsChangedHandler(event) {
  console.log("\n📎 OnMessageAttachmentsChanged Event Triggered");

  const item = Office.context.mailbox.item;

  item.attachments.getAttachmentsAsync((result) => {
    const attachments = result.value || [];

    const eventData = {
      eventType: "OnMessageAttachmentsChanged",
      category: "change",
      description: "Message attachments were added or removed",
      itemDetails: getItemDetails(item),
      attachmentCount: attachments.length,
      attachments: attachments.map(att => ({
        id: att.id,
        name: att.name,
        size: att.size,
        attachmentType: att.attachmentType
      })),
      timestamp: new Date().toISOString()
    };

    logEventToConsole("OnMessageAttachmentsChanged", eventData);

    event.completed();
  });
}

// 4. OnAppointmentAttachmentsChanged Handler
function onAppointmentAttachmentsChangedHandler(event) {
  console.log("\n📎 OnAppointmentAttachmentsChanged Event Triggered");

  const item = Office.context.mailbox.item;

  item.attachments.getAttachmentsAsync((result) => {
    const attachments = result.value || [];

    const eventData = {
      eventType: "OnAppointmentAttachmentsChanged",
      category: "change",
      description: "Appointment attachments were added or removed",
      itemDetails: getItemDetails(item),
      attachmentCount: attachments.length,
      attachments: attachments.map(att => ({
        id: att.id,
        name: att.name,
        size: att.size,
        attachmentType: att.attachmentType
      })),
      timestamp: new Date().toISOString()
    };

    logEventToConsole("OnAppointmentAttachmentsChanged", eventData);

    event.completed();
  });
}

// 5. OnMessageRecipientsChanged Handler
function onMessageRecipientsChangedHandler(event) {
  console.log("\n👥 OnMessageRecipientsChanged Event Triggered");

  const item = Office.context.mailbox.item;
  const eventData = {
    eventType: "OnMessageRecipientsChanged",
    category: "change",
    description: "Message recipients (To, Cc, or Bcc) were modified",
    itemDetails: getItemDetails(item),
    changedField: event.changedRecipientFields || "unknown",
    timestamp: new Date().toISOString()
  };

  // Get current recipients
  item.to.getAsync((toResult) => {
    item.cc.getAsync((ccResult) => {
      item.bcc.getAsync((bccResult) => {
        eventData.recipients = {
          to: toResult.value || [],
          cc: ccResult.value || [],
          bcc: bccResult.value || []
        };

        logEventToConsole("OnMessageRecipientsChanged", eventData);

        event.completed();
      });
    });
  });
}

// 6. OnAppointmentAttendeesChanged Handler
function onAppointmentAttendeesChangedHandler(event) {
  console.log("\n👥 OnAppointmentAttendeesChanged Event Triggered");

  const item = Office.context.mailbox.item;

  item.requiredAttendees.getAsync((requiredResult) => {
    item.optionalAttendees.getAsync((optionalResult) => {
      const eventData = {
        eventType: "OnAppointmentAttendeesChanged",
        category: "change",
        description: "Appointment attendees were added or removed",
        itemDetails: getItemDetails(item),
        attendees: {
          required: requiredResult.value || [],
          optional: optionalResult.value || []
        },
        timestamp: new Date().toISOString()
      };

      logEventToConsole("OnAppointmentAttendeesChanged", eventData);

      event.completed();
    });
  });
}

// 7. OnAppointmentTimeChanged Handler
function onAppointmentTimeChangedHandler(event) {
  console.log("\n🕐 OnAppointmentTimeChanged Event Triggered");

  const item = Office.context.mailbox.item;

  item.start.getAsync((startResult) => {
    item.end.getAsync((endResult) => {
      const eventData = {
        eventType: "OnAppointmentTimeChanged",
        category: "change",
        description: "Appointment start or end time was modified",
        itemDetails: getItemDetails(item),
        timeDetails: {
          start: startResult.value,
          end: endResult.value,
          duration: startResult.value && endResult.value ?
            (new Date(endResult.value) - new Date(startResult.value)) / 60000 + " minutes" :
            "unknown"
        },
        timestamp: new Date().toISOString()
      };

      logEventToConsole("OnAppointmentTimeChanged", eventData);

      event.completed();
    });
  });
}

// 8. OnAppointmentRecurrenceChanged Handler
function onAppointmentRecurrenceChangedHandler(event) {
  console.log("\n🔄 OnAppointmentRecurrenceChanged Event Triggered");

  const item = Office.context.mailbox.item;

  item.recurrence.getAsync((result) => {
    const eventData = {
      eventType: "OnAppointmentRecurrenceChanged",
      category: "change",
      description: "Appointment recurrence pattern was modified",
      itemDetails: getItemDetails(item),
      recurrence: result.value,
      timestamp: new Date().toISOString()
    };

    logEventToConsole("OnAppointmentRecurrenceChanged", eventData);

    event.completed();
  });
}

// 9. OnInfoBarDismissClicked Handler
function onInfoBarDismissClickedHandler(event) {
  console.log("\n❌ OnInfoBarDismissClicked Event Triggered");

  const eventData = {
    eventType: "OnInfoBarDismissClicked",
    category: "action",
    description: "User dismissed an information bar",
    infobarKey: event.infobarDetails ? event.infobarDetails.key : "unknown",
    timestamp: new Date().toISOString()
  };

  logEventToConsole("OnInfoBarDismissClicked", eventData);

  event.completed();
}

// 10. OnMessageSend Handler
function onMessageSendHandler(event) {
  console.log("\n📤 OnMessageSend Event Triggered");

  const item = Office.context.mailbox.item;

  // Get all message details before sending
  item.subject.getAsync((subjectResult) => {
    item.to.getAsync((toResult) => {
      item.cc.getAsync((ccResult) => {
        item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
          const eventData = {
            eventType: "OnMessageSend",
            category: "send",
            description: "User is attempting to send a message",
            itemDetails: getItemDetails(item),
            subject: subjectResult.value,
            recipients: {
              to: toResult.value || [],
              cc: ccResult.value || []
            },
            bodyPreview: bodyResult.value ?
              bodyResult.value.substring(0, 100) + "..." :
              "No body content",
            timestamp: new Date().toISOString()
          };

          logEventToConsole("OnMessageSend", eventData);

          // Allow the send operation to continue
          event.completed({ allowEvent: true });
        });
      });
    });
  });
}

// 11. OnAppointmentSend Handler
function onAppointmentSendHandler(event) {
  console.log("\n📤 OnAppointmentSend Event Triggered");

  const item = Office.context.mailbox.item;

  item.subject.getAsync((subjectResult) => {
    item.start.getAsync((startResult) => {
      item.end.getAsync((endResult) => {
        item.location.getAsync((locationResult) => {
          const eventData = {
            eventType: "OnAppointmentSend",
            category: "send",
            description: "User is attempting to send an appointment/meeting request",
            itemDetails: getItemDetails(item),
            subject: subjectResult.value,
            start: startResult.value,
            end: endResult.value,
            location: locationResult.value,
            timestamp: new Date().toISOString()
          };

          logEventToConsole("OnAppointmentSend", eventData);

          // Allow the send operation to continue
          event.completed({ allowEvent: true });
        });
      });
    });
  });
}

// 12. OnMessageFromChanged Handler
function onMessageFromChangedHandler(event) {
  console.log("\n📧 OnMessageFromChanged Event Triggered");

  const item = Office.context.mailbox.item;

  item.from.getAsync((result) => {
    const eventData = {
      eventType: "OnMessageFromChanged",
      category: "change",
      description: "The From field was changed (shared mailbox scenario)",
      itemDetails: getItemDetails(item),
      from: result.value,
      timestamp: new Date().toISOString()
    };

    logEventToConsole("OnMessageFromChanged", eventData);

    event.completed();
  });
}

// 13. OnSensitivityLabelChanged Handler
function onSensitivityLabelChangedHandler(event) {
  console.log("\n🔒 OnSensitivityLabelChanged Event Triggered");

  const item = Office.context.mailbox.item;

  const eventData = {
    eventType: "OnSensitivityLabelChanged",
    category: "change",
    description: "Sensitivity label was changed on the item",
    itemDetails: getItemDetails(item),
    sensitivity: item.sensitivity,
    timestamp: new Date().toISOString()
  };

  logEventToConsole("OnSensitivityLabelChanged", eventData);

  event.completed();
}

// Register all event handlers with Office
if (Office.actions) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onNewAppointmentOrganizerHandler", onNewAppointmentOrganizerHandler);
  Office.actions.associate("onMessageAttachmentsChangedHandler", onMessageAttachmentsChangedHandler);
  Office.actions.associate("onAppointmentAttachmentsChangedHandler", onAppointmentAttachmentsChangedHandler);
  Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
  Office.actions.associate("onAppointmentAttendeesChangedHandler", onAppointmentAttendeesChangedHandler);
  Office.actions.associate("onAppointmentTimeChangedHandler", onAppointmentTimeChangedHandler);
  Office.actions.associate("onAppointmentRecurrenceChangedHandler", onAppointmentRecurrenceChangedHandler);
  Office.actions.associate("onInfoBarDismissClickedHandler", onInfoBarDismissClickedHandler);
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
  Office.actions.associate("onMessageFromChangedHandler", onMessageFromChangedHandler);
  Office.actions.associate("onSensitivityLabelChangedHandler", onSensitivityLabelChangedHandler);

  console.log("✅ All event handlers registered successfully");
}

// Export for testing/debugging
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    eventTracker,
    logEventToConsole
  };
}