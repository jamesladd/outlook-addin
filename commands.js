/* global Office */

Office.onReady(() => {
  // Office is ready
});

/**
 * Shows the subject of the selected email(s) in a dialog
 * This function is called when the InboxAgent Action button is clicked
 */
function showSubject(event) {
  // Get the current item
  const item = Office.context.mailbox.item;

  // Get the subject
  item.subject.getAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const subject = result.value;

      // Display the subject in a dialog
      const message = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: `Subject: ${subject}`,
        icon: "Icon.80x80",
        persistent: false
      };

      // Show notification
      item.notificationMessages.addAsync("subjectNotification", message, (notificationResult) => {
        if (notificationResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to show notification:", notificationResult.error);
        }
      });

    } else {
      console.error("Failed to get subject:", result.error);
    }

    // Required: Signal that the function has completed
    event.completed();
  });
}

// Register the function
Office.actions = Office.actions || {};
Office.actions.showSubject = showSubject;