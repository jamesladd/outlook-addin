/* global Office */

// This file handles all event-based activation scenarios

Office.onReady(() => {
  console.log('%c=== InboxAgent Event Handler Initialized ===', 'color: #0078d4; font-size: 14px; font-weight: bold;');
});

// OnNewMessageCompose Event Handler
function onNewMessageComposeHandler(event) {
  console.log('%c[LAUNCH EVENT] OnNewMessageCompose', 'color: #10b981; font-weight: bold;', event);

  logDetailedEvent('OnNewMessageCompose', event, {
    description: 'User started composing a new message',
    itemType: 'Message',
    mode: 'Compose'
  });

  event.completed();
}

// OnNewAppointmentOrganizer Event Handler
function onNewAppointmentOrganizerHandler(event) {
  console.log('%c[LAUNCH EVENT] OnNewAppointmentOrganizer', 'color: #10b981; font-weight: bold;', event);

  logDetailedEvent('OnNewAppointmentOrganizer', event, {
    description: 'User started creating a new appointment',
    itemType: 'Appointment',
    mode: 'Organizer'
  });

  event.completed();
}

// OnMessageAttachmentsChanged Event Handler
function onMessageAttachmentsChangedHandler(event) {
  console.log('%c[LAUNCH EVENT] OnMessageAttachmentsChanged', 'color: #10b981; font-weight: bold;', event);

  Office.context.mailbox.item.getAttachmentsAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      logDetailedEvent('OnMessageAttachmentsChanged', event, {
        description: 'Message attachments have been modified',
        attachmentCount: asyncResult.value.length,
        attachments: asyncResult.value.map(att => ({
          id: att.id,
          name: att.name,
          size: att.size,
          type: att.attachmentType
        }))
      });
    }
  });

  event.completed();
}

// OnAppointmentAttachmentsChanged Event Handler
function onAppointmentAttachmentsChangedHandler(event) {
  console.log('%c[LAUNCH EVENT] OnAppointmentAttachmentsChanged', 'color: #10b981; font-weight: bold;', event);

  Office.context.mailbox.item.getAttachmentsAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      logDetailedEvent('OnAppointmentAttachmentsChanged', event, {
        description: 'Appointment attachments have been modified',
        attachmentCount: asyncResult.value.length,
        attachments: asyncResult.value.map(att => ({
          id: att.id,
          name: att.name,
          size: att.size
        }))
      });
    }
  });

  event.completed();
}

// OnMessageRecipientsChanged Event Handler
function onMessageRecipientsChangedHandler(event) {
  console.log('%c[LAUNCH EVENT] OnMessageRecipientsChanged', 'color: #10b981; font-weight: bold;', event);

  const item = Office.context.mailbox.item;
  const recipientData = {};

  item.to.getAsync((toResult) => {
    recipientData.to = toResult.value;

    item.cc.getAsync((ccResult) => {
      recipientData.cc = ccResult.value;

      item.bcc.getAsync((bccResult) => {
        recipientData.bcc = bccResult.value;

        logDetailedEvent('OnMessageRecipientsChanged', event, {
          description: 'Message recipients have been modified',
          changedRecipients: event.changedRecipientFields,
          toCount: recipientData.to.length,
          ccCount: recipientData.cc.length,
          bccCount: recipientData.bcc.length,
          recipients: {
            to: recipientData.to.map(r => r.emailAddress),
            cc: recipientData.cc.map(r => r.emailAddress),
            bcc: recipientData.bcc.map(r => r.emailAddress)
          }
        });
      });
    });
  });

  event.completed();
}

// OnAppointmentAttendeesChanged Event Handler
function onAppointmentAttendeesChangedHandler(event) {
  console.log('%c[LAUNCH EVENT] OnAppointmentAttendeesChanged', 'color: #10b981; font-weight: bold;', event);

  Office.context.mailbox.item.requiredAttendees.getAsync((reqResult) => {
    Office.context.mailbox.item.optionalAttendees.getAsync((optResult) => {
      logDetailedEvent('OnAppointmentAttendeesChanged', event, {
        description: 'Appointment attendees have been modified',
        requiredCount: reqResult.value.length,
        optionalCount: optResult.value.length,
        attendees: {
          required: reqResult.value.map(a => a.emailAddress),
          optional: optResult.value.map(a => a.emailAddress)
        }
      });
    });
  });

  event.completed();
}

// OnAppointmentTimeChanged Event Handler
function onAppointmentTimeChangedHandler(event) {
  console.log('%c[LAUNCH EVENT] OnAppointmentTimeChanged', 'color: #10b981; font-weight: bold;', event);

  Office.context.mailbox.item.start.getAsync((startResult) => {
    Office.context.mailbox.item.end.getAsync((endResult) => {
      logDetailedEvent('OnAppointmentTimeChanged', event, {
        description: 'Appointment time has been modified',
        startTime: startResult.value,
        endTime: endResult.value,
        duration: (new Date(endResult.value) - new Date(startResult.value)) / 60000 + ' minutes'
      });
    });
  });

  event.completed();
}

// OnAppointmentRecurrenceChanged Event Handler
function onAppointmentRecurrenceChangedHandler(event) {
  console.log('%c[LAUNCH EVENT] OnAppointmentRecurrenceChanged', 'color: #10b981; font-weight: bold;', event);

  Office.context.mailbox.item.recurrence.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      logDetailedEvent('OnAppointmentRecurrenceChanged', event, {
        description: 'Appointment recurrence pattern has been modified',
        recurrence: asyncResult.value,
        seriesTime: asyncResult.value ? asyncResult.value.seriesTime : null,
        recurrenceType: asyncResult.value ? asyncResult.value.recurrenceType : 'none'
      });
    }
  });

  event.completed();
}

// OnInfoBarDismissClicked Event Handler
function onInfoBarDismissClickedHandler(event) {
  console.log('%c[LAUNCH EVENT] OnInfoBarDismissClicked', 'color: #10b981; font-weight: bold;', event);

  logDetailedEvent('OnInfoBarDismissClicked', event, {
    description: 'User dismissed an information bar',
    infobarKey: event.infobarType
  });

  event.completed();
}

// OnMessageSend Event Handler
function onMessageSendHandler(event) {
  console.log('%c[LAUNCH EVENT] OnMessageSend', 'color: #10b981; font-weight: bold;', event);

  Office.context.mailbox.item.subject.getAsync((subjectResult) => {
    Office.context.mailbox.item.to.getAsync((toResult) => {
      logDetailedEvent('OnMessageSend', event, {
        description: 'User is attempting to send a message',
        subject: subjectResult.value,
        recipientCount: toResult.value.length,
        recipients: toResult.value.map(r => r.emailAddress)
      });

      // Always allow send for demo purposes
      event.completed({ allowEvent: true });
    });
  });
}

// OnAppointmentSend Event Handler
function onAppointmentSendHandler(event) {
  console.log('%c[LAUNCH EVENT] OnAppointmentSend', 'color: #10b981; font-weight: bold;', event);

  Office.context.mailbox.item.subject.getAsync((subjectResult) => {
    Office.context.mailbox.item.requiredAttendees.getAsync((attendeesResult) => {
      logDetailedEvent('OnAppointmentSend', event, {
        description: 'User is attempting to send an appointment',
        subject: subjectResult.value,
        attendeeCount: attendeesResult.value.length,
        attendees: attendeesResult.value.map(a => a.emailAddress)
      });

      // Always allow send for demo purposes
      event.completed({ allowEvent: true });
    });
  });
}

// OnMessageFromChanged Event Handler
function onMessageFromChangedHandler(event) {
  console.log('%c[LAUNCH EVENT] OnMessageFromChanged', 'color: #10b981; font-weight: bold;', event);

  Office.context.mailbox.item.from.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      logDetailedEvent('OnMessageFromChanged', event, {
        description: 'Message "From" field has been changed',
        from: asyncResult.value.emailAddress,
        displayName: asyncResult.value.displayName
      });
    }
  });

  event.completed();
}

// OnSensitivityLabelChanged Event Handler
function onSensitivityLabelChangedHandler(event) {
  console.log('%c[LAUNCH EVENT] OnSensitivityLabelChanged', 'color: #10b981; font-weight: bold;', event);

  logDetailedEvent('OnSensitivityLabelChanged', event, {
    description: 'Sensitivity label has been changed',
    sensitivityLabel: event.sensitivityLabel || 'Not available'
  });

  event.completed();
}

// Helper function to log detailed event information
function logDetailedEvent(eventName, event, additionalData) {
  const detailedLog = {
    eventName: eventName,
    timestamp: new Date().toISOString(),
    eventObject: {
      type: event.type,
      source: event.source,
      completed: typeof event.completed
    },
    mailboxInfo: {
      userProfile: Office.context.mailbox.userProfile.emailAddress,
      diagnostics: Office.context.mailbox.diagnostics
    },
    itemInfo: {
      itemId: Office.context.mailbox.item.itemId,
      itemType: Office.context.mailbox.item.itemType,
      itemClass: Office.context.mailbox.item.itemClass
    },
    additionalData: additionalData
  };

  console.log(`%c[DETAILED EVENT LOG] ${eventName}`, 'color: #f59e0b; font-weight: bold;');
  console.log('Event Details:', detailedLog);
  console.log('Raw Event Object:', event);
  console.log('─'.repeat(80));
}

// Register all event handlers with Office
if (typeof Office !== 'undefined' && Office.actions) {
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

  console.log('✓ All event handlers registered successfully');
}