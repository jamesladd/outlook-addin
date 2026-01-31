/* global Office */

Office.onReady(() => {
  console.log('='.repeat(80));
  console.log('INBOXAGENT: Commands runtime initialized');
  console.log('Event-based activation is ready');
  console.log('='.repeat(80));
});

// Event Handler: OnNewMessageCompose
function onNewMessageComposeHandler(event) {
  console.log('='.repeat(80));
  console.log('EVENT-BASED ACTIVATION: OnNewMessageCompose');
  console.log('Timestamp:', new Date().toISOString());
  console.log('Event Type:', event.type);
  console.log('Description: User started composing a new message');

  const item = Office.context.mailbox.item;
  if (item) {
    console.log('New Message Details:');
    console.log('- Item Type:', item.itemType);
    console.log('- Item Mode:', 'Compose');
  }

  console.log('='.repeat(80));

  event.completed();
}

// Event Handler: OnMessageAttachmentsChanged
function onMessageAttachmentsChangedHandler(event) {
  console.log('='.repeat(80));
  console.log('EVENT-BASED ACTIVATION: OnMessageAttachmentsChanged');
  console.log('Timestamp:', new Date().toISOString());
  console.log('Event Type:', event.type);
  console.log('Description: Attachments were added, removed, or modified');

  const item = Office.context.mailbox.item;
  if (item && item.attachments) {
    console.log('Current Attachments:');
    console.log('- Count:', item.attachments.length);
    item.attachments.forEach((attachment, index) => {
      console.log(`- Attachment ${index + 1}:`, attachment.name, '(', attachment.size, 'bytes)');
    });
  }

  console.log('='.repeat(80));

  event.completed();
}

// Event Handler: OnMessageRecipientsChanged
function onMessageRecipientsChangedHandler(event) {
  console.log('='.repeat(80));
  console.log('EVENT-BASED ACTIVATION: OnMessageRecipientsChanged');
  console.log('Timestamp:', new Date().toISOString());
  console.log('Event Type:', event.type);
  console.log('Description: Recipients (To, Cc, or Bcc) were changed');

  const item = Office.context.mailbox.item;
  if (item) {
    // Log TO recipients
    if (item.to && item.to.getAsync) {
      item.to.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log('TO Recipients:', result.value.map(r => r.emailAddress).join(', '));
        }
      });
    }

    // Log CC recipients
    if (item.cc && item.cc.getAsync) {
      item.cc.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log('CC Recipients:', result.value.map(r => r.emailAddress).join(', '));
        }
      });
    }
  }

  console.log('='.repeat(80));

  event.completed();
}

// Event Handler: OnMessageSend
function onMessageSendHandler(event) {
  console.log('='.repeat(80));
  console.log('EVENT-BASED ACTIVATION: OnMessageSend');
  console.log('Timestamp:', new Date().toISOString());
  console.log('Event Type:', event.type);
  console.log('Description: User is attempting to send a message');

  const item = Office.context.mailbox.item;
  if (item) {
    if (item.subject && item.subject.getAsync) {
      item.subject.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log('Message Subject:', result.value);
        }
      });
    }

    if (item.to && item.to.getAsync) {
      item.to.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log('Sending to:', result.value.map(r => r.emailAddress).join(', '));
        }
      });
    }
  }

  console.log('Action: Allowing send to proceed');
  console.log('='.repeat(80));

  // Allow the send to proceed
  event.completed({ allowEvent: true });
}

// Event Handler: OnAppointmentSend
function onAppointmentSendHandler(event) {
  console.log('='.repeat(80));
  console.log('EVENT-BASED ACTIVATION: OnAppointmentSend');
  console.log('Timestamp:', new Date().toISOString());
  console.log('Event Type:', event.type);
  console.log('Description: User is attempting to send an appointment/meeting invite');

  const item = Office.context.mailbox.item;
  if (item) {
    if (item.subject && item.subject.getAsync) {
      item.subject.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log('Appointment Subject:', result.value);
        }
      });
    }

    if (item.start && item.start.getAsync) {
      item.start.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log('Start Time:', result.value);
        }
      });
    }

    if (item.end && item.end.getAsync) {
      item.end.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log('End Time:', result.value);
        }
      });
    }
  }

  console.log('Action: Allowing send to proceed');
  console.log('='.repeat(80));

  event.completed({ allowEvent: true });
}

// Event Handler: OnNewAppointmentOrganizer
function onNewAppointmentOrganizerHandler(event) {
  console.log('='.repeat(80));
  console.log('EVENT-BASED ACTIVATION: OnNewAppointmentOrganizer');
  console.log('Timestamp:', new Date().toISOString());
  console.log('Event Type:', event.type);
  console.log('Description: User started creating a new appointment as organizer');

  const item = Office.context.mailbox.item;
  if (item) {
    console.log('New Appointment Details:');
    console.log('- Item Type:', item.itemType);
  }

  console.log('='.repeat(80));

  event.completed();
}

// Event Handler: OnMessageFromChanged
function onMessageFromChangedHandler(event) {
  console.log('='.repeat(80));
  console.log('EVENT-BASED ACTIVATION: OnMessageFromChanged');
  console.log('Timestamp:', new Date().toISOString());
  console.log('Event Type:', event.type);
  console.log('Description: The FROM address was changed (multi-account scenario)');

  const item = Office.context.mailbox.item;
  if (item && item.from && item.from.getAsync) {
    item.from.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log('New FROM address:', result.value.emailAddress);
        console.log('Display Name:', result.value.displayName);
      }
    });
  }

  console.log('='.repeat(80));

  event.completed();
}

// Event Handler: OnSensitivityLabelChanged
function onSensitivityLabelChangedHandler(event) {
  console.log('='.repeat(80));
  console.log('EVENT-BASED ACTIVATION: OnSensitivityLabelChanged');
  console.log('Timestamp:', new Date().toISOString());
  console.log('Event Type:', event.type);
  console.log('Description: The sensitivity label was changed');

  const item = Office.context.mailbox.item;
  if (item) {
    console.log('Sensitivity label changed for item:', item.itemType);
  }

  console.log('='.repeat(80));

  event.completed();
}

// Register all event handlers with Office
if (typeof Office !== 'undefined') {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onMessageAttachmentsChangedHandler", onMessageAttachmentsChangedHandler);
  Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
  Office.actions.associate("onNewAppointmentOrganizerHandler", onNewAppointmentOrganizerHandler);
  Office.actions.associate("onMessageFromChangedHandler", onMessageFromChangedHandler);
  Office.actions.associate("onSensitivityLabelChangedHandler", onSensitivityLabelChangedHandler);
}