/* global Office */

// Ensure Office is ready before doing anything
(function() {
  'use strict';

  // Track conversation IDs to detect replies/forwards
  const conversationTracker = new Map();
  let lastReadItemId = null;
  let lastReadConversationId = null;

  console.log('InboxAgent launchevent.js loading...');

  // Office.onReady with error handling
  Office.onReady(() => {
    console.log('%c=== InboxAgent Event Handler Initialized ===', 'color: #0078d4; font-size: 14px; font-weight: bold;');

    // Register all handlers immediately
    registerEventHandlers();
  }).catch((error) => {
    console.error('Office.onReady failed:', error);
  });

  // Register all event handlers
  function registerEventHandlers() {
    if (typeof Office === 'undefined' || !Office.actions) {
      console.error('Office.actions is not available!');
      return;
    }

    try {
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
    } catch (error) {
      console.error('Failed to register event handlers:', error);
    }
  }

  // Enhanced OnNewMessageCompose Event Handler with Reply/Forward Detection
  function onNewMessageComposeHandler(event) {
    console.log('%c[LAUNCH EVENT] OnNewMessageCompose', 'color: #10b981; font-weight: bold;', event);

    try {
      const item = Office.context.mailbox.item;

      if (!item) {
        console.error('No item available in OnNewMessageCompose');
        event.completed();
        return;
      }

      // Detect reply or forward by analyzing the compose item
      detectComposeAction(item, (actionType, details) => {
        logDetailedEvent('OnNewMessageCompose', event, {
          description: `User started composing a new message (${actionType})`,
          itemType: 'Message',
          mode: 'Compose',
          action: actionType, // 'NEW', 'REPLY', 'REPLY_ALL', 'FORWARD'
          ...details
        });
      });

      event.completed();
    } catch (error) {
      console.error('Error in onNewMessageComposeHandler:', error);
      event.completed();
    }
  }

  // Function to detect if this is a reply, reply all, forward, or new message
  function detectComposeAction(item, callback) {
    try {
      const details = {};

      // Get conversation ID
      const conversationId = item.conversationId;
      details.conversationId = conversationId;

      // Get subject to check for RE: or FW: prefixes
      item.subject.getAsync((subjectResult) => {
        if (subjectResult.status === Office.AsyncResultStatus.Succeeded) {
          const subject = subjectResult.value || '';
          details.subject = subject;

          // Check for reply/forward indicators in subject
          const isReply = subject.match(/^(RE:|Re:)/i);
          const isForward = subject.match(/^(FW:|Fw:|FWD:)/i);

          // Get body to check for quoted content
          item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
            if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
              const bodyText = bodyResult.value || '';
              details.bodyLength = bodyText.length;

              // Check for typical reply/forward markers in body
              const hasQuotedContent = bodyText.includes('From:') ||
                bodyText.includes('Sent:') ||
                bodyText.includes('-----Original Message-----') ||
                bodyText.match(/On .+ wrote:/);

              details.hasQuotedContent = hasQuotedContent;

              // Get recipients to help determine action type
              item.to.getAsync((toResult) => {
                if (toResult.status === Office.AsyncResultStatus.Succeeded) {
                  const recipients = toResult.value;
                  details.recipientCount = recipients.length;
                  details.recipients = recipients.map(r => r.emailAddress);

                  // Get CC recipients
                  item.cc.getAsync((ccResult) => {
                    if (ccResult.status === Office.AsyncResultStatus.Succeeded) {
                      const ccRecipients = ccResult.value;
                      details.ccCount = ccRecipients.length;
                      details.ccRecipients = ccRecipients.map(r => r.emailAddress);

                      // Determine action type based on collected data
                      let actionType = 'NEW';

                      if (isReply && hasQuotedContent) {
                        // Check if it's reply all (multiple recipients or CC present)
                        if (recipients.length > 1 || ccRecipients.length > 0) {
                          actionType = 'REPLY_ALL';
                          console.log('%c🔄 REPLY ALL DETECTED!', 'color: #f59e0b; font-size: 14px; font-weight: bold;');
                        } else {
                          actionType = 'REPLY';
                          console.log('%c↩️ REPLY DETECTED!', 'color: #10b981; font-size: 14px; font-weight: bold;');
                        }
                      } else if (isForward && hasQuotedContent) {
                        actionType = 'FORWARD';
                        console.log('%c➡️ FORWARD DETECTED!', 'color: #3b82f6; font-size: 14px; font-weight: bold;');
                      }

                      // Check against conversation tracker
                      if (conversationId && lastReadConversationId === conversationId) {
                        details.relatedToLastRead = true;
                        if (actionType === 'NEW' && hasQuotedContent) {
                          actionType = 'REPLY'; // Fallback detection
                        }
                      }

                      details.detectedAction = actionType;
                      callback(actionType, details);
                    }
                  });
                }
              });
            }
          });
        }
      });
    } catch (error) {
      console.error('Error in detectComposeAction:', error);
      callback('ERROR', { error: error.message });
    }
  }

  // OnNewAppointmentOrganizer Event Handler
  function onNewAppointmentOrganizerHandler(event) {
    console.log('%c[LAUNCH EVENT] OnNewAppointmentOrganizer', 'color: #10b981; font-weight: bold;', event);

    try {
      logDetailedEvent('OnNewAppointmentOrganizer', event, {
        description: 'User started creating a new appointment',
        itemType: 'Appointment',
        mode: 'Organizer'
      });
      event.completed();
    } catch (error) {
      console.error('Error in onNewAppointmentOrganizerHandler:', error);
      event.completed();
    }
  }

  // OnMessageAttachmentsChanged Event Handler
  function onMessageAttachmentsChangedHandler(event) {
    console.log('%c[LAUNCH EVENT] OnMessageAttachmentsChanged', 'color: #10b981; font-weight: bold;', event);

    try {
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
        event.completed();
      });
    } catch (error) {
      console.error('Error in onMessageAttachmentsChangedHandler:', error);
      event.completed();
    }
  }

  // OnAppointmentAttachmentsChanged Event Handler
  function onAppointmentAttachmentsChangedHandler(event) {
    console.log('%c[LAUNCH EVENT] OnAppointmentAttachmentsChanged', 'color: #10b981; font-weight: bold;', event);

    try {
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
        event.completed();
      });
    } catch (error) {
      console.error('Error in onAppointmentAttachmentsChangedHandler:', error);
      event.completed();
    }
  }

  // OnMessageRecipientsChanged Event Handler
  function onMessageRecipientsChangedHandler(event) {
    console.log('%c[LAUNCH EVENT] OnMessageRecipientsChanged', 'color: #10b981; font-weight: bold;', event);

    try {
      const item = Office.context.mailbox.item;
      const recipientData = {};

      item.to.getAsync((toResult) => {
        recipientData.to = toResult.value || [];

        item.cc.getAsync((ccResult) => {
          recipientData.cc = ccResult.value || [];

          item.bcc.getAsync((bccResult) => {
            recipientData.bcc = bccResult.value || [];

            logDetailedEvent('OnMessageRecipientsChanged', event, {
              description: 'Message recipients have been modified',
              changedRecipients: event.changedRecipientFields || [],
              toCount: recipientData.to.length,
              ccCount: recipientData.cc.length,
              bccCount: recipientData.bcc.length,
              recipients: {
                to: recipientData.to.map(r => r.emailAddress),
                cc: recipientData.cc.map(r => r.emailAddress),
                bcc: recipientData.bcc.map(r => r.emailAddress)
              }
            });

            event.completed();
          });
        });
      });
    } catch (error) {
      console.error('Error in onMessageRecipientsChangedHandler:', error);
      event.completed();
    }
  }

  // OnAppointmentAttendeesChanged Event Handler
  function onAppointmentAttendeesChangedHandler(event) {
    console.log('%c[LAUNCH EVENT] OnAppointmentAttendeesChanged', 'color: #10b981; font-weight: bold;', event);

    try {
      Office.context.mailbox.item.requiredAttendees.getAsync((reqResult) => {
        Office.context.mailbox.item.optionalAttendees.getAsync((optResult) => {
          logDetailedEvent('OnAppointmentAttendeesChanged', event, {
            description: 'Appointment attendees have been modified',
            requiredCount: reqResult.value ? reqResult.value.length : 0,
            optionalCount: optResult.value ? optResult.value.length : 0,
            attendees: {
              required: reqResult.value ? reqResult.value.map(a => a.emailAddress) : [],
              optional: optResult.value ? optResult.value.map(a => a.emailAddress) : []
            }
          });
          event.completed();
        });
      });
    } catch (error) {
      console.error('Error in onAppointmentAttendeesChangedHandler:', error);
      event.completed();
    }
  }

  // OnAppointmentTimeChanged Event Handler
  function onAppointmentTimeChangedHandler(event) {
    console.log('%c[LAUNCH EVENT] OnAppointmentTimeChanged', 'color: #10b981; font-weight: bold;', event);

    try {
      Office.context.mailbox.item.start.getAsync((startResult) => {
        Office.context.mailbox.item.end.getAsync((endResult) => {
          logDetailedEvent('OnAppointmentTimeChanged', event, {
            description: 'Appointment time has been modified',
            startTime: startResult.value,
            endTime: endResult.value,
            duration: (new Date(endResult.value) - new Date(startResult.value)) / 60000 + ' minutes'
          });
          event.completed();
        });
      });
    } catch (error) {
      console.error('Error in onAppointmentTimeChangedHandler:', error);
      event.completed();
    }
  }

  // OnAppointmentRecurrenceChanged Event Handler
  function onAppointmentRecurrenceChangedHandler(event) {
    console.log('%c[LAUNCH EVENT] OnAppointmentRecurrenceChanged', 'color: #10b981; font-weight: bold;', event);

    try {
      Office.context.mailbox.item.recurrence.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          logDetailedEvent('OnAppointmentRecurrenceChanged', event, {
            description: 'Appointment recurrence pattern has been modified',
            recurrence: asyncResult.value,
            seriesTime: asyncResult.value ? asyncResult.value.seriesTime : null,
            recurrenceType: asyncResult.value ? asyncResult.value.recurrenceType : 'none'
          });
        }
        event.completed();
      });
    } catch (error) {
      console.error('Error in onAppointmentRecurrenceChangedHandler:', error);
      event.completed();
    }
  }

  // OnInfoBarDismissClicked Event Handler
  function onInfoBarDismissClickedHandler(event) {
    console.log('%c[LAUNCH EVENT] OnInfoBarDismissClicked', 'color: #10b981; font-weight: bold;', event);

    try {
      logDetailedEvent('OnInfoBarDismissClicked', event, {
        description: 'User dismissed an information bar',
        infobarKey: event.infobarType || 'unknown'
      });
      event.completed();
    } catch (error) {
      console.error('Error in onInfoBarDismissClickedHandler:', error);
      event.completed();
    }
  }

  // OnMessageSend Event Handler
  function onMessageSendHandler(event) {
    console.log('%c[LAUNCH EVENT] OnMessageSend', 'color: #10b981; font-weight: bold;', event);

    try {
      Office.context.mailbox.item.subject.getAsync((subjectResult) => {
        Office.context.mailbox.item.to.getAsync((toResult) => {
          const subject = subjectResult.value || '';
          const isReply = subject.match(/^(RE:|Re:)/i);
          const isForward = subject.match(/^(FW:|Fw:|FWD:)/i);

          let sendAction = 'SEND_NEW';
          if (isReply) sendAction = 'SEND_REPLY';
          if (isForward) sendAction = 'SEND_FORWARD';

          console.log(`%c📤 ${sendAction} DETECTED!`, 'color: #ef4444; font-size: 14px; font-weight: bold;');

          logDetailedEvent('OnMessageSend', event, {
            description: `User is attempting to send a message (${sendAction})`,
            sendAction: sendAction,
            subject: subject,
            recipientCount: toResult.value ? toResult.value.length : 0,
            recipients: toResult.value ? toResult.value.map(r => r.emailAddress) : []
          });

          // Always allow send
          event.completed({ allowEvent: true });
        });
      });
    } catch (error) {
      console.error('Error in onMessageSendHandler:', error);
      // Always allow send even on error
      event.completed({ allowEvent: true });
    }
  }

  // OnAppointmentSend Event Handler
  function onAppointmentSendHandler(event) {
    console.log('%c[LAUNCH EVENT] OnAppointmentSend', 'color: #10b981; font-weight: bold;', event);

    try {
      Office.context.mailbox.item.subject.getAsync((subjectResult) => {
        Office.context.mailbox.item.requiredAttendees.getAsync((attendeesResult) => {
          logDetailedEvent('OnAppointmentSend', event, {
            description: 'User is attempting to send an appointment',
            subject: subjectResult.value || '',
            attendeeCount: attendeesResult.value ? attendeesResult.value.length : 0,
            attendees: attendeesResult.value ? attendeesResult.value.map(a => a.emailAddress) : []
          });

          // Always allow send
          event.completed({ allowEvent: true });
        });
      });
    } catch (error) {
      console.error('Error in onAppointmentSendHandler:', error);
      // Always allow send even on error
      event.completed({ allowEvent: true });
    }
  }

  // OnMessageFromChanged Event Handler
  function onMessageFromChangedHandler(event) {
    console.log('%c[LAUNCH EVENT] OnMessageFromChanged', 'color: #10b981; font-weight: bold;', event);

    try {
      Office.context.mailbox.item.from.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          logDetailedEvent('OnMessageFromChanged', event, {
            description: 'Message "From" field has been changed',
            from: asyncResult.value ? asyncResult.value.emailAddress : 'unknown',
            displayName: asyncResult.value ? asyncResult.value.displayName : 'unknown'
          });
        }
        event.completed();
      });
    } catch (error) {
      console.error('Error in onMessageFromChangedHandler:', error);
      event.completed();
    }
  }

  // OnSensitivityLabelChanged Event Handler
  function onSensitivityLabelChangedHandler(event) {
    console.log('%c[LAUNCH EVENT] OnSensitivityLabelChanged', 'color: #10b981; font-weight: bold;', event);

    try {
      logDetailedEvent('OnSensitivityLabelChanged', event, {
        description: 'Sensitivity label has been changed',
        sensitivityLabel: event.sensitivityLabel || 'Not available'
      });
      event.completed();
    } catch (error) {
      console.error('Error in onSensitivityLabelChangedHandler:', error);
      event.completed();
    }
  }

  // Helper function to log detailed event information
  function logDetailedEvent(eventName, event, additionalData) {
    try {
      const detailedLog = {
        eventName: eventName,
        timestamp: new Date().toISOString(),
        eventObject: {
          type: event.type,
          source: event.source,
          completed: typeof event.completed
        },
        mailboxInfo: {
          userProfile: Office.context.mailbox.userProfile ? Office.context.mailbox.userProfile.emailAddress : 'unknown',
          diagnostics: Office.context.mailbox.diagnostics
        },
        itemInfo: {
          itemId: Office.context.mailbox.item ? Office.context.mailbox.item.itemId : null,
          itemType: Office.context.mailbox.item ? Office.context.mailbox.item.itemType : null,
          itemClass: Office.context.mailbox.item ? Office.context.mailbox.item.itemClass : null
        },
        additionalData: additionalData
      };

      console.log(`%c[DETAILED EVENT LOG] ${eventName}`, 'color: #f59e0b; font-weight: bold;');
      console.log('Event Details:', detailedLog);
      console.log('Raw Event Object:', event);
      console.log('─'.repeat(80));
    } catch (error) {
      console.error('Error in logDetailedEvent:', error);
    }
  }

  console.log('InboxAgent launchevent.js loaded successfully');

})();