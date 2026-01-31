/* global Office */

(function() {
  'use strict';

  // Shared storage for events (persists across contexts)
  const STORAGE_KEY = 'InboxAgent_Events';

  // Track conversation IDs to detect replies/forwards
  const conversationTracker = new Map();
  let lastReadItemId = null;
  let lastReadConversationId = null;

  console.log('InboxAgent launchevent.js loading...');

  Office.onReady(() => {
    console.log('%c=== InboxAgent Event Handler Initialized ===', 'color: #0078d4; font-size: 14px; font-weight: bold;');
    registerEventHandlers();
  }).catch((error) => {
    console.error('Office.onReady failed:', error);
  });

  // Store event in roaming settings (persists across sessions)
  function storeEvent(eventData) {
    try {
      // Try to use roaming settings if available
      if (Office.context.roamingSettings) {
        const existingData = Office.context.roamingSettings.get(STORAGE_KEY);
        let events = [];

        try {
          events = existingData ? JSON.parse(existingData) : [];
        } catch (e) {
          console.warn('Failed to parse existing events:', e);
          events = [];
        }

        // Add new event
        events.push(eventData);

        // Keep only last 100 events
        if (events.length > 100) {
          events = events.slice(-100);
        }

        // Save back to roaming settings
        Office.context.roamingSettings.set(STORAGE_KEY, JSON.stringify(events));
        Office.context.roamingSettings.saveAsync((asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log('✓ Event stored successfully');
          } else {
            console.error('Failed to save event:', asyncResult.error);
          }
        });
      } else {
        console.warn('Roaming settings not available');
      }
    } catch (error) {
      console.error('Error storing event:', error);
    }
  }

  // Show notification to user
  function showNotification(title, message, eventType) {
    try {
      const item = Office.context.mailbox.item;
      if (item && item.notificationMessages) {
        const notificationKey = 'inboxagent_' + Date.now();

        item.notificationMessages.addAsync(
          notificationKey,
          {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: `${title}: ${message}`,
            icon: 'icon-16',
            persistent: false
          },
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log('✓ Notification shown:', title);

              // Auto-remove after 5 seconds
              setTimeout(() => {
                if (item.notificationMessages) {
                  item.notificationMessages.removeAsync(notificationKey);
                }
              }, 5000);
            } else {
              console.warn('Failed to show notification:', asyncResult.error);
            }
          }
        );
      }
    } catch (error) {
      console.error('Error showing notification:', error);
    }
  }

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

  // ============================================================================
  // EVENT HANDLERS
  // ============================================================================

  // OnNewMessageCompose Event Handler with Reply/Forward Detection
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
        const eventData = {
          id: Date.now(),
          type: 'OnNewMessageCompose',
          action: actionType,
          timestamp: new Date().toISOString(),
          description: `User started composing a new message (${actionType})`,
          details: details
        };

        logDetailedEvent('OnNewMessageCompose', event, eventData);
        storeEvent(eventData);

        // Show notification based on action type
        let notificationTitle = '✉️ New Email';
        let notificationMessage = 'InboxAgent is tracking this action';

        if (actionType === 'REPLY') {
          notificationTitle = '↩️ Reply';
          notificationMessage = 'Reply detected and logged';
        } else if (actionType === 'REPLY_ALL') {
          notificationTitle = '🔄 Reply All';
          notificationMessage = 'Reply All detected and logged';
        } else if (actionType === 'FORWARD') {
          notificationTitle = '➡️ Forward';
          notificationMessage = 'Forward detected and logged';
        }

        showNotification(notificationTitle, notificationMessage, actionType);
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
                  const recipients = toResult.value || [];
                  details.recipientCount = recipients.length;
                  details.recipients = recipients.map(r => r.emailAddress);

                  // Get CC recipients
                  item.cc.getAsync((ccResult) => {
                    if (ccResult.status === Office.AsyncResultStatus.Succeeded) {
                      const ccRecipients = ccResult.value || [];
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
      const eventData = {
        id: Date.now(),
        type: 'OnNewAppointmentOrganizer',
        action: 'NEW_APPOINTMENT',
        timestamp: new Date().toISOString(),
        description: 'User started creating a new appointment',
        details: {
          itemType: 'Appointment',
          mode: 'Organizer'
        }
      };

      logDetailedEvent('OnNewAppointmentOrganizer', event, eventData);
      storeEvent(eventData);
      showNotification('📅 New Appointment', 'Appointment creation detected', 'NEW_APPOINTMENT');

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
        const attachments = asyncResult.status === Office.AsyncResultStatus.Succeeded ?
          asyncResult.value : [];

        const eventData = {
          id: Date.now(),
          type: 'OnMessageAttachmentsChanged',
          action: 'ATTACHMENTS_MODIFIED',
          timestamp: new Date().toISOString(),
          description: 'Message attachments have been modified',
          details: {
            attachmentCount: attachments.length,
            attachments: attachments.map(att => ({
              id: att.id,
              name: att.name,
              size: att.size,
              type: att.attachmentType
            }))
          }
        };

        logDetailedEvent('OnMessageAttachmentsChanged', event, eventData);
        storeEvent(eventData);
        showNotification('📎 Attachments Changed', `${attachments.length} attachment(s)`, 'ATTACHMENTS');

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
        const attachments = asyncResult.status === Office.AsyncResultStatus.Succeeded ?
          asyncResult.value : [];

        const eventData = {
          id: Date.now(),
          type: 'OnAppointmentAttachmentsChanged',
          action: 'APPOINTMENT_ATTACHMENTS_MODIFIED',
          timestamp: new Date().toISOString(),
          description: 'Appointment attachments have been modified',
          details: {
            attachmentCount: attachments.length,
            attachments: attachments.map(att => ({
              id: att.id,
              name: att.name,
              size: att.size
            }))
          }
        };

        logDetailedEvent('OnAppointmentAttachmentsChanged', event, eventData);
        storeEvent(eventData);
        showNotification('📎 Appointment Attachments', `${attachments.length} attachment(s)`, 'ATTACHMENTS');

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

      item.to.getAsync((toResult) => {
        const toRecipients = toResult.status === Office.AsyncResultStatus.Succeeded ?
          toResult.value : [];

        item.cc.getAsync((ccResult) => {
          const ccRecipients = ccResult.status === Office.AsyncResultStatus.Succeeded ?
            ccResult.value : [];

          item.bcc.getAsync((bccResult) => {
            const bccRecipients = bccResult.status === Office.AsyncResultStatus.Succeeded ?
              bccResult.value : [];

            const eventData = {
              id: Date.now(),
              type: 'OnMessageRecipientsChanged',
              action: 'RECIPIENTS_MODIFIED',
              timestamp: new Date().toISOString(),
              description: 'Message recipients have been modified',
              details: {
                changedRecipientFields: event.changedRecipientFields || [],
                toCount: toRecipients.length,
                ccCount: ccRecipients.length,
                bccCount: bccRecipients.length,
                recipients: {
                  to: toRecipients.map(r => r.emailAddress),
                  cc: ccRecipients.map(r => r.emailAddress),
                  bcc: bccRecipients.map(r => r.emailAddress)
                }
              }
            };

            logDetailedEvent('OnMessageRecipientsChanged', event, eventData);
            storeEvent(eventData);

            const totalRecipients = toRecipients.length + ccRecipients.length + bccRecipients.length;
            showNotification('👥 Recipients Changed', `${totalRecipients} recipient(s)`, 'RECIPIENTS');

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
        const requiredAttendees = reqResult.status === Office.AsyncResultStatus.Succeeded ?
          reqResult.value : [];

        Office.context.mailbox.item.optionalAttendees.getAsync((optResult) => {
          const optionalAttendees = optResult.status === Office.AsyncResultStatus.Succeeded ?
            optResult.value : [];

          const eventData = {
            id: Date.now(),
            type: 'OnAppointmentAttendeesChanged',
            action: 'ATTENDEES_MODIFIED',
            timestamp: new Date().toISOString(),
            description: 'Appointment attendees have been modified',
            details: {
              requiredCount: requiredAttendees.length,
              optionalCount: optionalAttendees.length,
              attendees: {
                required: requiredAttendees.map(a => a.emailAddress),
                optional: optionalAttendees.map(a => a.emailAddress)
              }
            }
          };

          logDetailedEvent('OnAppointmentAttendeesChanged', event, eventData);
          storeEvent(eventData);

          const totalAttendees = requiredAttendees.length + optionalAttendees.length;
          showNotification('👥 Attendees Changed', `${totalAttendees} attendee(s)`, 'ATTENDEES');

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
        const startTime = startResult.status === Office.AsyncResultStatus.Succeeded ?
          startResult.value : null;

        Office.context.mailbox.item.end.getAsync((endResult) => {
          const endTime = endResult.status === Office.AsyncResultStatus.Succeeded ?
            endResult.value : null;

          const duration = (startTime && endTime) ?
            (new Date(endTime) - new Date(startTime)) / 60000 + ' minutes' : 'Unknown';

          const eventData = {
            id: Date.now(),
            type: 'OnAppointmentTimeChanged',
            action: 'TIME_MODIFIED',
            timestamp: new Date().toISOString(),
            description: 'Appointment time has been modified',
            details: {
              startTime: startTime,
              endTime: endTime,
              duration: duration
            }
          };

          logDetailedEvent('OnAppointmentTimeChanged', event, eventData);
          storeEvent(eventData);
          showNotification('🕒 Time Changed', 'Appointment time updated', 'TIME');

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
        const recurrence = asyncResult.status === Office.AsyncResultStatus.Succeeded ?
          asyncResult.value : null;

        const eventData = {
          id: Date.now(),
          type: 'OnAppointmentRecurrenceChanged',
          action: 'RECURRENCE_MODIFIED',
          timestamp: new Date().toISOString(),
          description: 'Appointment recurrence pattern has been modified',
          details: {
            recurrence: recurrence,
            seriesTime: recurrence ? recurrence.seriesTime : null,
            recurrenceType: recurrence ? recurrence.recurrenceType : 'none'
          }
        };

        logDetailedEvent('OnAppointmentRecurrenceChanged', event, eventData);
        storeEvent(eventData);
        showNotification('🔄 Recurrence Changed', 'Recurrence pattern updated', 'RECURRENCE');

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
      const eventData = {
        id: Date.now(),
        type: 'OnInfoBarDismissClicked',
        action: 'INFOBAR_DISMISSED',
        timestamp: new Date().toISOString(),
        description: 'User dismissed an information bar',
        details: {
          infobarKey: event.infobarType || 'unknown'
        }
      };

      logDetailedEvent('OnInfoBarDismissClicked', event, eventData);
      storeEvent(eventData);

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
          const recipients = toResult.value || [];

          const isReply = subject.match(/^(RE:|Re:)/i);
          const isForward = subject.match(/^(FW:|Fw:|FWD:)/i);

          let sendAction = 'SEND_NEW';
          if (isReply) sendAction = 'SEND_REPLY';
          if (isForward) sendAction = 'SEND_FORWARD';

          console.log(`%c📤 ${sendAction} DETECTED!`, 'color: #ef4444; font-size: 14px; font-weight: bold;');

          const eventData = {
            id: Date.now(),
            type: 'OnMessageSend',
            action: sendAction,
            timestamp: new Date().toISOString(),
            description: `User is sending a message (${sendAction})`,
            details: {
              sendAction: sendAction,
              subject: subject,
              recipientCount: recipients.length,
              recipients: recipients.map(r => r.emailAddress)
            }
          };

          logDetailedEvent('OnMessageSend', event, eventData);
          storeEvent(eventData);
          showNotification('📤 Sending Email', sendAction, sendAction);

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
          const subject = subjectResult.value || '';
          const attendees = attendeesResult.value || [];

          const eventData = {
            id: Date.now(),
            type: 'OnAppointmentSend',
            action: 'SEND_APPOINTMENT',
            timestamp: new Date().toISOString(),
            description: 'User is sending an appointment',
            details: {
              subject: subject,
              attendeeCount: attendees.length,
              attendees: attendees.map(a => a.emailAddress)
            }
          };

          logDetailedEvent('OnAppointmentSend', event, eventData);
          storeEvent(eventData);
          showNotification('📤 Sending Appointment', 'Appointment being sent', 'SEND_APPOINTMENT');

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
        const fromInfo = asyncResult.status === Office.AsyncResultStatus.Succeeded ?
          asyncResult.value : null;

        const eventData = {
          id: Date.now(),
          type: 'OnMessageFromChanged',
          action: 'FROM_CHANGED',
          timestamp: new Date().toISOString(),
          description: 'Message "From" field has been changed',
          details: {
            from: fromInfo ? fromInfo.emailAddress : 'unknown',
            displayName: fromInfo ? fromInfo.displayName : 'unknown'
          }
        };

        logDetailedEvent('OnMessageFromChanged', event, eventData);
        storeEvent(eventData);
        showNotification('📧 From Changed', 'Sender account changed', 'FROM');

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
      const eventData = {
        id: Date.now(),
        type: 'OnSensitivityLabelChanged',
        action: 'SENSITIVITY_CHANGED',
        timestamp: new Date().toISOString(),
        description: 'Sensitivity label has been changed',
        details: {
          sensitivityLabel: event.sensitivityLabel || 'Not available'
        }
      };

      logDetailedEvent('OnSensitivityLabelChanged', event, eventData);
      storeEvent(eventData);
      showNotification('🔒 Sensitivity Changed', 'Email sensitivity updated', 'SENSITIVITY');

      event.completed();
    } catch (error) {
      console.error('Error in onSensitivityLabelChangedHandler:', error);
      event.completed();
    }
  }

  // ============================================================================
  // HELPER FUNCTIONS
  // ============================================================================

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
          userProfile: Office.context.mailbox.userProfile ?
            Office.context.mailbox.userProfile.emailAddress : 'unknown',
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

  console.log('✓ InboxAgent launchevent.js loaded successfully');

})();