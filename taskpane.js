/* global Office */

// Global event counter and storage
let eventCounter = 0;
let activeListeners = 0;
let eventHistory = [];
let monitoringInterval = null;
let lastKnownState = {
  categories: null,
  importance: null,
  itemId: null,
  categoriesInitialized: false,
  importanceInitialized: false
};
let lastCategoryCheck = 0;
const CATEGORY_CHECK_THROTTLE = 3000; // 3 seconds between checks
let categoryCheckInProgress = false; // Prevent overlapping checks

const STORAGE_KEY = 'InboxAgent_Events';

// Initialize Office Add-in
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log('%c=== InboxAgent Taskpane Initialized ===', 'color: #0078d4; font-size: 16px; font-weight: bold;');
    console.log('Host:', info.host);
    console.log('Platform:', info.platform);

    // Setup event listeners for UI
    setupUIListeners();

    // Load current item information
    loadItemInfo();

    // Setup Office event listeners
    setupOfficeEventListeners();

    // Load stored events from previous sessions
    loadStoredEvents();

    // Start property monitoring if in read mode
    const item = Office.context.mailbox.item;
    if (item && item.itemId) {
      startPropertyMonitoring();
    }

    // Log that taskpane is ready
    logEvent('TaskpaneReady', 'Taskpane has been initialized and is ready', {
      host: info.host,
      platform: info.platform,
      isPinned: isPinned()
    });
  }
});

// Load stored events when taskpane opens
function loadStoredEvents() {
  try {
    const storedData = localStorage.getItem(STORAGE_KEY);
    if (storedData) {
      let storedEvents = [];
      try {
        storedEvents = JSON.parse(storedData);
      } catch (e) {
        storedEvents = [];
      }

      console.log(`📦 Loading ${storedEvents.length} stored events...`);

      // Set the counter to the last event ID to continue numbering
      if (storedEvents.length > 0) {
        const lastEvent = storedEvents[storedEvents.length - 1];
        if (lastEvent.id) {
          eventCounter = lastEvent.id;
        }
      }

      // Copy to history without triggering saves
      eventHistory = [...storedEvents];

      // Display stored events in the taskpane WITHOUT saving them again
      storedEvents.forEach((event) => {
        displayStoredEvent(event);
      });

      // Update the total counter display
      document.getElementById('totalEvents').textContent = eventCounter;

      if (storedEvents.length > 0) {
        console.log('✓ Stored events loaded successfully');
      }
    }
  } catch (error) {
    console.error('Error loading stored events:', error);
  }
}

// Display a stored event without saving it again or incrementing counter
function displayStoredEvent(event) {
  const eventLog = document.getElementById('eventLog');
  const placeholder = eventLog.querySelector('.event-placeholder');
  if (placeholder) {
    placeholder.remove();
  }

  const verboseLogging = document.getElementById('verboseLogging').checked;
  const timestampEvents = document.getElementById('timestampEvents').checked;

  const eventItem = document.createElement('div');
  eventItem.className = 'event-item stored-event';

  const eventHeader = document.createElement('div');
  eventHeader.className = 'event-header';

  const eventTypeSpan = document.createElement('span');
  eventTypeSpan.className = 'event-type';
  eventTypeSpan.textContent = `#${event.id} - ${event.type}`;

  const eventTime = document.createElement('span');
  eventTime.className = 'event-time';
  if (timestampEvents) {
    eventTime.textContent = new Date(event.timestamp).toLocaleTimeString();
  }

  eventHeader.appendChild(eventTypeSpan);
  eventHeader.appendChild(eventTime);

  const eventDetails = document.createElement('div');
  eventDetails.className = 'event-details';
  eventDetails.textContent = event.description;

  eventItem.appendChild(eventHeader);
  eventItem.appendChild(eventDetails);

  if (verboseLogging && event.details && Object.keys(event.details).length > 0) {
    const eventData = document.createElement('div');
    eventData.className = 'event-data';
    eventData.textContent = JSON.stringify(event.details, null, 2);
    eventItem.appendChild(eventData);
  }

  eventLog.appendChild(eventItem);
}

// Check if taskpane is pinned
function isPinned() {
  // Note: There's no direct API to check if pinned, this is a placeholder
  return 'Pinning supported - check manually in Outlook';
}

// Setup UI Event Listeners
function setupUIListeners() {
  // Refresh button
  document.getElementById('refreshBtn').addEventListener('click', () => {
    logEvent('ButtonClick', 'Refresh button clicked', {});
    loadItemInfo();
  });

  // Test event button
  document.getElementById('testEventBtn').addEventListener('click', () => {
    logEvent('ButtonClick', 'Test Event button clicked', {});
    logEvent('TestEvent', 'This is a manually triggered test event', {
      timestamp: new Date().toISOString(),
      testData: 'Sample test data',
      random: Math.random()
    });
  });

  // Get properties button
  document.getElementById('getPropertiesBtn').addEventListener('click', () => {
    logEvent('ButtonClick', 'Get Properties button clicked', {});
    getItemProperties();
  });

  // Clear events button
  document.getElementById('clearEventsBtn').addEventListener('click', () => {
    clearEvents();
  });

  // Export events button
  document.getElementById('exportEventsBtn').addEventListener('click', () => {
    exportEvents();
  });

  // Check properties button
  const checkPropertiesBtn = document.getElementById('checkPropertiesBtn');
  if (checkPropertiesBtn) {
    checkPropertiesBtn.addEventListener('click', () => {
      monitorItemProperties();
      logEvent('ManualCheck', 'Manual property check triggered', {});
    });
  }

  // Toggle monitoring button
  const toggleMonitoringBtn = document.getElementById('toggleMonitoringBtn');
  if (toggleMonitoringBtn) {
    toggleMonitoringBtn.addEventListener('click', () => {
      if (monitoringInterval) {
        stopPropertyMonitoring();
        toggleMonitoringBtn.textContent = '▶️ Start Monitoring';
        logEvent('MonitoringStopped', 'Property monitoring stopped', {});
      } else {
        startPropertyMonitoring();
        toggleMonitoringBtn.textContent = '⏸️ Stop Monitoring';
        logEvent('MonitoringStarted', 'Property monitoring started', {});
      }
    });
  }

  // Close reminder button
  const closeReminderBtn = document.getElementById('closeReminder');
  if (closeReminderBtn) {
    closeReminderBtn.addEventListener('click', () => {
      document.querySelector('.pin-reminder').style.display = 'none';
    });
  }

  console.log('✓ UI Event listeners configured');
}

// Helper function to compare arrays (more robust version)
function arraysEqual(a, b) {
  // Handle null/undefined cases
  if (!a && !b) return true;
  if (!a || !b) return false;

  // Check length
  if (a.length !== b.length) return false;

  // Sort and compare
  const sortedA = [...a].sort();
  const sortedB = [...b].sort();

  // Deep comparison
  for (let i = 0; i < sortedA.length; i++) {
    if (sortedA[i] !== sortedB[i]) {
      return false;
    }
  }

  return true;
}

// Create a deep copy of an array to prevent reference issues
function deepCopyArray(arr) {
  if (!arr) return [];
  return JSON.parse(JSON.stringify(arr));
}

// Helper function to compare arrays of category objects
function categoriesEqual(a, b) {
  // Handle null/undefined cases
  if (!a && !b) return true;
  if (!a || !b) return false;

  // Check length first
  if (a.length !== b.length) {
    console.log('Categories different length:', a.length, 'vs', b.length);
    return false;
  }

  // If both empty, they're equal
  if (a.length === 0) return true;

  // Compare by converting to sorted JSON strings
  // Sort by displayName to ensure consistent ordering
  const sortedA = [...a].sort((x, y) =>
    (x.displayName || '').localeCompare(y.displayName || '')
  );
  const sortedB = [...b].sort((x, y) =>
    (x.displayName || '').localeCompare(y.displayName || '')
  );

  const jsonA = JSON.stringify(sortedA);
  const jsonB = JSON.stringify(sortedB);

  const isEqual = jsonA === jsonB;

  console.log('Category comparison:', {
    aCount: a.length,
    bCount: b.length,
    jsonA: jsonA,
    jsonB: jsonB,
    isEqual: isEqual
  });

  return isEqual;
}

// Extract just the category names for comparison and display
function getCategoryNames(categories) {
  if (!categories) return [];
  return categories.map(cat => cat.displayName || cat);
}

// Property Monitoring Functions
function monitorItemProperties() {
  const item = Office.context.mailbox.item;

  if (!item || !item.itemId) {
    return;
  }

  if (lastKnownState.itemId !== item.itemId) {
    console.log('📌 New item detected, resetting monitoring state');
    resetMonitoringState(item);
    return;
  }

  if (!lastKnownState.categoriesInitialized) {
    console.log('⏳ Waiting for categories to initialize...');
    return;
  }

  // Throttle category checks
  const now = Date.now();
  if (now - lastCategoryCheck < CATEGORY_CHECK_THROTTLE) {
    return; // Skip this check
  }

  // Prevent overlapping async calls
  if (categoryCheckInProgress) {
    console.log('⏳ Category check already in progress, skipping...');
    return;
  }

  // Monitor Categories
  if (item.categories) {
    categoryCheckInProgress = true;
    lastCategoryCheck = now;

    item.categories.getAsync((result) => {
      // Always clear the lock when done
      categoryCheckInProgress = false;

      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error('Failed to get categories:', result.error);
        return;
      }

      const currentCategories = result.value || [];
      const previousCategories = lastKnownState.categories || [];

      console.log('═══════════════════════════════════════');
      console.log('CATEGORY CHECK at', new Date().toLocaleTimeString());
      console.log('Previous:', JSON.stringify(previousCategories));
      console.log('Current:', JSON.stringify(currentCategories));
      console.log('Previous length:', previousCategories.length);
      console.log('Current length:', currentCategories.length);

      // Check if actually different using category-aware comparison
      const areEqual = categoriesEqual(currentCategories, previousCategories);
      console.log('Categories equal?', areEqual);

      if (!areEqual) {
        console.log('%c🏷️ CATEGORIES CHANGED DETECTED!', 'color: #8b5cf6; font-size: 14px; font-weight: bold;');

        // Get category names for easier comparison
        const previousNames = getCategoryNames(previousCategories);
        const currentNames = getCategoryNames(currentCategories);

        console.log('Previous names:', previousNames);
        console.log('Current names:', currentNames);

        // Calculate added and removed by display name
        const addedNames = currentNames.filter(name => !previousNames.includes(name));
        const removedNames = previousNames.filter(name => !currentNames.includes(name));

        // Get full category objects for added/removed
        const addedCategories = currentCategories.filter(cat =>
          addedNames.includes(cat.displayName || cat)
        );
        const removedCategories = previousCategories.filter(cat =>
          removedNames.includes(cat.displayName || cat)
        );

        console.log('Added names:', addedNames);
        console.log('Removed names:', removedNames);
        console.log('Added categories:', JSON.stringify(addedCategories));
        console.log('Removed categories:', JSON.stringify(removedCategories));

        // CRITICAL: Only log if there's an actual change
        const hasAdditions = addedNames.length > 0;
        const hasRemovals = removedNames.length > 0;
        const hasChanges = hasAdditions || hasRemovals;

        // Check that added and removed names are completely different
        const overlap = addedNames.filter(name => removedNames.includes(name));
        const addedAndRemovedAreDifferent = overlap.length === 0;
        const isValidChange = hasChanges && addedAndRemovedAreDifferent;

        console.log('Change validation:', {
          hasAdditions: hasAdditions,
          hasRemovals: hasRemovals,
          hasChanges: hasChanges,
          overlap: overlap,
          addedAndRemovedAreDifferent: addedAndRemovedAreDifferent,
          isValidChange: isValidChange
        });

        if (isValidChange) {

          // Update state BEFORE logging the event to prevent loops
          lastKnownState.categories = deepCopyArray(currentCategories);
          console.log('✓ State updated to:', JSON.stringify(lastKnownState.categories));

          logEvent('CategoriesChanged', 'Email categories have been modified', {
            previousCategories: previousNames,
            currentCategories: currentNames,
            addedCategoryNames: addedNames,
            removedCategoryNames: removedNames,
            addedCategories: addedCategories,
            removedCategories: removedCategories,
            itemId: item.itemId,
            subject: item.subject,
            timestamp: new Date().toISOString()
          });

          // Show notifications
          if (addedNames.length > 0) {
            showInAppNotification('🏷️ Category Added', addedNames.join(', '));
          }
          if (removedNames.length > 0) {
            showInAppNotification('🏷️ Category Removed', removedNames.join(', '));
          }
        } else {
          if (!addedAndRemovedAreDifferent) {
            console.error('❌ CRITICAL ERROR: Added and Removed categories OVERLAP!');
            console.error('Overlap:', overlap);
            console.error('Added names:', addedNames);
            console.error('Removed names:', removedNames);
            console.error('This indicates a false positive - same category in both lists');
            console.error('Previous state:', JSON.stringify(previousCategories));
            console.error('Current state:', JSON.stringify(currentCategories));

            // Force a complete reset to break the loop
            console.log('🔧 Forcing state reset to break potential loop...');
            lastKnownState.categories = deepCopyArray(currentCategories);
            categoryCheckInProgress = false;
            lastCategoryCheck = Date.now() + 10000; // Block checks for 10 seconds

            // Show error notification
            showInAppNotification('⚠️ Monitoring Error', 'False positive detected - monitoring paused for 10s');
          } else {
            console.log('ℹ️ No actual changes detected (empty added and removed)');
            // Still update state to latest
            lastKnownState.categories = deepCopyArray(currentCategories);
          }
        }
      } else {
        console.log('✓ No category change detected');
        // Update state even when no change to ensure we have latest
        lastKnownState.categories = deepCopyArray(currentCategories);
      }
      console.log('═══════════════════════════════════════');
    });
  }

  // Monitor Importance
  const currentImportance = item.importance;

  if (lastKnownState.importanceInitialized &&
    currentImportance !== lastKnownState.importance) {
    console.log('%c⚠️ IMPORTANCE CHANGED!', 'color: #f59e0b; font-size: 14px; font-weight: bold;');
    console.log('Previous:', lastKnownState.importance);
    console.log('Current:', currentImportance);

    const importanceNames = {
      0: 'Low',
      1: 'Normal',
      2: 'High'
    };

    logEvent('ImportanceChanged', 'Email importance level has been modified', {
      previousImportance: importanceNames[lastKnownState.importance] || 'Unknown',
      currentImportance: importanceNames[currentImportance] || 'Unknown',
      itemId: item.itemId,
      subject: item.subject
    });

    lastKnownState.importance = currentImportance;
    showInAppNotification('⚠️ Importance Changed', importanceNames[currentImportance]);
  } else if (!lastKnownState.importanceInitialized) {
    lastKnownState.importance = currentImportance;
    lastKnownState.importanceInitialized = true;
    console.log('✓ Importance initialized:', currentImportance);
  }
}

// Reset monitoring state for new item
function resetMonitoringState(item) {
  console.log('🔄 Resetting monitoring state for new item');

  // Clear the check lock
  categoryCheckInProgress = false;
  lastCategoryCheck = 0;

  // Clear all state
  lastKnownState = {
    itemId: item.itemId,
    importance: item.importance,
    categories: [],
    categoriesInitialized: false,
    importanceInitialized: false
  };

  console.log('Fetching initial categories...');

  // Get initial categories (asynchronous)
  if (item.categories) {
    item.categories.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        // Deep copy to prevent reference issues
        lastKnownState.categories = deepCopyArray(result.value || []);

        console.log('✓ Initial categories captured:', JSON.stringify(lastKnownState.categories));
        console.log('✓ Category count:', lastKnownState.categories.length);

        // Mark as initialized immediately - no delay needed with proper locking
        lastKnownState.categoriesInitialized = true;
        console.log('✓ Categories monitoring now active');

      } else {
        console.error('Failed to get initial categories:', result.error);
        lastKnownState.categories = [];
        lastKnownState.categoriesInitialized = true;
      }
    });
  } else {
    lastKnownState.categories = [];
    lastKnownState.categoriesInitialized = true;
  }
}

// Start monitoring
function startPropertyMonitoring() {
  // Clear any existing interval
  if (monitoringInterval) {
    clearInterval(monitoringInterval);
  }

  const item = Office.context.mailbox.item;
  if (item && item.itemId) {
    console.log('🚀 Starting property monitoring...');

    // Reset state for this item
    resetMonitoringState(item);

    // Wait 2 seconds before starting polling to let initial state fully settle
    setTimeout(() => {
      // Poll every 3 seconds (increased from 2 to reduce race conditions)
      monitoringInterval = setInterval(monitorItemProperties, 3000);
      console.log('✓ Property monitoring active (polling every 3 seconds)');

      // Log monitoring started
      logEvent('MonitoringStarted', 'Started monitoring item properties', {
        itemId: item.itemId,
        itemType: item.itemType,
        subject: item.subject
      });

      // Update button state
      const toggleBtn = document.getElementById('toggleMonitoringBtn');
      if (toggleBtn) {
        toggleBtn.textContent = '⏸️ Stop Monitoring';
      }
    }, 2000);

  } else {
    console.log('⚠️ Cannot start monitoring in compose mode');
  }
}

// Stop monitoring
function stopPropertyMonitoring() {
  if (monitoringInterval) {
    clearInterval(monitoringInterval);
    monitoringInterval = null;
    console.log('✓ Property monitoring stopped');

    // Reset flags
    categoryCheckInProgress = false;
    lastCategoryCheck = 0;
    lastKnownState.categoriesInitialized = false;
    lastKnownState.importanceInitialized = false;

    // Update button state
    const toggleBtn = document.getElementById('toggleMonitoringBtn');
    if (toggleBtn) {
      toggleBtn.textContent = '▶️ Start Monitoring';
    }
  }
}

// Helper function to compare arrays
function arraysEqual(a, b) {
  if (a.length !== b.length) return false;
  const sortedA = [...a].sort();
  const sortedB = [...b].sort();
  return sortedA.every((val, index) => val === sortedB[index]);
}

// Show in-app notification in taskpane
function showInAppNotification(title, message) {
  let notificationArea = document.getElementById('notificationArea');

  if (!notificationArea) {
    // Create notification area if it doesn't exist
    notificationArea = document.createElement('div');
    notificationArea.id = 'notificationArea';
    notificationArea.className = 'notification-area';
    document.querySelector('.main-content').insertBefore(
      notificationArea,
      document.querySelector('.main-content').firstChild
    );
  }

  const notification = document.createElement('div');
  notification.className = 'in-app-notification';
  notification.innerHTML = `
    <div class="notification-content">
      <strong>${escapeHtml(title)}</strong>
      <p>${escapeHtml(message)}</p>
    </div>
    <button class="notification-close">×</button>
  `;

  notificationArea.appendChild(notification);

  // Add close handler
  notification.querySelector('.notification-close').addEventListener('click', () => {
    notification.remove();
  });

  // Auto-remove after 5 seconds
  setTimeout(() => {
    notification.classList.add('fade-out');
    setTimeout(() => notification.remove(), 300);
  }, 5000);
}

// Start monitoring
function startPropertyMonitoring() {
  // Clear any existing interval
  if (monitoringInterval) {
    clearInterval(monitoringInterval);
  }

  const item = Office.context.mailbox.item;
  if (item && item.itemId) {
    resetMonitoringState(item);

    // Wait 1 second before starting polling to let initial state settle
    setTimeout(() => {
      // Poll every 2 seconds
      monitoringInterval = setInterval(monitorItemProperties, 2000);
      console.log('✓ Property monitoring started (polling every 2 seconds)');

      // Update button state
      const toggleBtn = document.getElementById('toggleMonitoringBtn');
      if (toggleBtn) {
        toggleBtn.textContent = '⏸️ Stop Monitoring';
      }
    }, 1000);

  } else {
    console.log('⚠️ Cannot start monitoring in compose mode');
  }
}

// Stop monitoring
function stopPropertyMonitoring() {
  if (monitoringInterval) {
    clearInterval(monitoringInterval);
    monitoringInterval = null;
    console.log('✓ Property monitoring stopped');

    // Update button state
    const toggleBtn = document.getElementById('toggleMonitoringBtn');
    if (toggleBtn) {
      toggleBtn.textContent = '▶️ Start Monitoring';
    }
  }
}

// Setup Office Event Listeners
function setupOfficeEventListeners() {
  try {
    const item = Office.context.mailbox.item;

    if (!item) {
      console.warn('⚠️ No item available to setup event listeners');
      return;
    }

    // ItemChanged event (when user switches items while taskpane is open)
    if (Office.context.mailbox.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(
        Office.EventType.ItemChanged,
        onItemChanged,
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log('✓ ItemChanged event listener registered');
            activeListeners++;
            updateActiveListeners();
          } else {
            console.error('✗ Failed to register ItemChanged listener:', asyncResult.error);
          }
        }
      );
    }

    // Check if we're in compose mode (item doesn't have itemId yet)
    const isComposeMode = !item.itemId;

    if (isComposeMode) {
      console.log('📝 Compose mode detected - registering compose-specific listeners');

      // RecipientsChanged event - only available in compose mode
      if (item.to && typeof item.to.addHandlerAsync === 'function') {
        item.to.addHandlerAsync(
          Office.EventType.RecipientsChanged,
          onRecipientsChanged,
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log('✓ To RecipientsChanged event listener registered');
              activeListeners++;
              updateActiveListeners();
            } else {
              console.warn('⚠️ Failed to register To RecipientsChanged listener:', asyncResult.error);
            }
          }
        );
      }

      // CC recipients
      if (item.cc && typeof item.cc.addHandlerAsync === 'function') {
        item.cc.addHandlerAsync(
          Office.EventType.RecipientsChanged,
          onCcChanged,
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log('✓ CC RecipientsChanged event listener registered');
              activeListeners++;
              updateActiveListeners();
            } else {
              console.warn('⚠️ Failed to register CC RecipientsChanged listener:', asyncResult.error);
            }
          }
        );
      }

      // BCC recipients
      if (item.bcc && typeof item.bcc.addHandlerAsync === 'function') {
        item.bcc.addHandlerAsync(
          Office.EventType.RecipientsChanged,
          onBccChanged,
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log('✓ BCC RecipientsChanged event listener registered');
              activeListeners++;
              updateActiveListeners();
            } else {
              console.warn('⚠️ Failed to register BCC RecipientsChanged listener:', asyncResult.error);
            }
          }
        );
      }

      // AttachmentsChanged event - only in compose mode
      if (item.addHandlerAsync) {
        item.addHandlerAsync(
          Office.EventType.AttachmentsChanged,
          onAttachmentsChanged,
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log('✓ AttachmentsChanged event listener registered');
              activeListeners++;
              updateActiveListeners();
            } else {
              console.warn('⚠️ Failed to register AttachmentsChanged listener:', asyncResult.error);
            }
          }
        );
      }

      // EnhancedLocationsChanged event (for appointments in compose mode)
      if (item.itemType === Office.MailboxEnums.ItemType.Appointment &&
        item.enhancedLocation &&
        typeof item.enhancedLocation.addHandlerAsync === 'function') {
        item.enhancedLocation.addHandlerAsync(
          Office.EventType.EnhancedLocationsChanged,
          onEnhancedLocationsChanged,
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log('✓ EnhancedLocationsChanged event listener registered');
              activeListeners++;
              updateActiveListeners();
            } else {
              console.warn('⚠️ Failed to register EnhancedLocationsChanged listener:', asyncResult.error);
            }
          }
        );
      }

      // RecurrenceChanged event (for appointments in compose mode)
      if (item.itemType === Office.MailboxEnums.ItemType.Appointment &&
        item.recurrence &&
        typeof item.recurrence.addHandlerAsync === 'function') {
        item.recurrence.addHandlerAsync(
          Office.EventType.RecurrenceChanged,
          onRecurrenceChanged,
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log('✓ RecurrenceChanged event listener registered');
              activeListeners++;
              updateActiveListeners();
            } else {
              console.warn('⚠️ Failed to register RecurrenceChanged listener:', asyncResult.error);
            }
          }
        );
      }
    } else {
      console.log('📖 Read mode detected - limited event listeners available');
    }

    console.log(`✓ Office event listeners setup complete. Total active: ${activeListeners}`);

  } catch (error) {
    console.error('Error setting up Office event listeners:', error);
    logEvent('Error', 'Failed to setup Office event listeners', {
      error: error.message,
      stack: error.stack
    });
  }
}

// Event Handlers
function onItemChanged(eventArgs) {
  console.log('%c[EVENT] ItemChanged', 'color: #10b981; font-weight: bold;', eventArgs);

  logEvent('ItemChanged', 'User switched to a different item', {
    eventType: eventArgs.type,
    eventArgs: JSON.stringify(eventArgs, null, 2)
  });

  // Stop old monitoring
  stopPropertyMonitoring();

  // Reload item info
  loadItemInfo();

  // Reset and re-setup event listeners for new item
  activeListeners = 0;
  activeListeners = 1; // ItemChanged listener still active
  updateActiveListeners();
  setupOfficeEventListeners();

  // Start monitoring for new item
  setTimeout(() => {
    startPropertyMonitoring();
  }, 500);
}

function onRecipientsChanged(eventArgs) {
  console.log('%c[EVENT] RecipientsChanged (To)', 'color: #10b981; font-weight: bold;', eventArgs);

  Office.context.mailbox.item.to.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      logEvent('RecipientsChanged', 'To recipients have been modified', {
        eventType: eventArgs.type,
        recipientCount: asyncResult.value.length,
        recipients: asyncResult.value.map(r => r.emailAddress)
      });
    }
  });
}

function onCcChanged(eventArgs) {
  console.log('%c[EVENT] RecipientsChanged (CC)', 'color: #10b981; font-weight: bold;', eventArgs);

  Office.context.mailbox.item.cc.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      logEvent('RecipientsChanged', 'CC recipients have been modified', {
        eventType: eventArgs.type,
        recipientType: 'CC',
        recipientCount: asyncResult.value.length,
        recipients: asyncResult.value.map(r => r.emailAddress)
      });
    }
  });
}

function onBccChanged(eventArgs) {
  console.log('%c[EVENT] RecipientsChanged (BCC)', 'color: #10b981; font-weight: bold;', eventArgs);

  Office.context.mailbox.item.bcc.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      logEvent('RecipientsChanged', 'BCC recipients have been modified', {
        eventType: eventArgs.type,
        recipientType: 'BCC',
        recipientCount: asyncResult.value.length
      });
    }
  });
}

function onAttachmentsChanged(eventArgs) {
  console.log('%c[EVENT] AttachmentsChanged', 'color: #10b981; font-weight: bold;', eventArgs);

  Office.context.mailbox.item.getAttachmentsAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      logEvent('AttachmentsChanged', 'Attachments have been added or removed', {
        eventType: eventArgs.type,
        attachmentCount: asyncResult.value.length,
        attachments: asyncResult.value.map(a => ({
          name: a.name,
          size: a.size,
          type: a.attachmentType
        }))
      });
    }
  });
}

function onEnhancedLocationsChanged(eventArgs) {
  console.log('%c[EVENT] EnhancedLocationsChanged', 'color: #10b981; font-weight: bold;', eventArgs);

  Office.context.mailbox.item.enhancedLocation.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      logEvent('EnhancedLocationsChanged', 'Appointment location has been modified', {
        eventType: eventArgs.type,
        locations: asyncResult.value
      });
    }
  });
}

function onRecurrenceChanged(eventArgs) {
  console.log('%c[EVENT] RecurrenceChanged', 'color: #10b981; font-weight: bold;', eventArgs);

  Office.context.mailbox.item.recurrence.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      logEvent('RecurrenceChanged', 'Appointment recurrence has been modified', {
        eventType: eventArgs.type,
        recurrence: asyncResult.value
      });
    }
  });
}

// Load current item information
function loadItemInfo() {
  const item = Office.context.mailbox.item;

  if (!item) {
    document.getElementById('itemType').textContent = 'No item selected';
    document.getElementById('itemSubject').textContent = 'N/A';
    document.getElementById('itemMode').textContent = 'N/A';
    document.getElementById('itemId').textContent = 'N/A';
    return;
  }

  // Item Type
  const itemType = item.itemType === Office.MailboxEnums.ItemType.Message ? 'Message' : 'Appointment';
  document.getElementById('itemType').textContent = itemType;

  // Mode
  const mode = item.itemId ? 'Read' : 'Compose';
  document.getElementById('itemMode').textContent = mode;

  // Subject
  if (item.subject) {
    if (typeof item.subject === 'string') {
      document.getElementById('itemSubject').textContent = item.subject || '(No subject)';
    } else {
      item.subject.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          document.getElementById('itemSubject').textContent = asyncResult.value || '(No subject)';
        }
      });
    }
  }

  // Item ID
  document.getElementById('itemId').textContent = item.itemId ? item.itemId.substring(0, 30) + '...' : 'New item';

  console.log('✓ Item information loaded');
  logEvent('ItemInfoLoaded', 'Current item information loaded', {
    itemType: itemType,
    mode: mode,
    hasItemId: !!item.itemId
  });
}

// Get detailed item properties
function getItemProperties() {
  const item = Office.context.mailbox.item;

  if (!item) {
    logEvent('Error', 'No item available to get properties', {});
    return;
  }

  const properties = {
    itemType: item.itemType,
    itemClass: item.itemClass,
    itemId: item.itemId,
    conversationId: item.conversationId,
    dateTimeCreated: item.dateTimeCreated,
    dateTimeModified: item.dateTimeModified,
    mode: item.itemId ? 'Read' : 'Compose',
    importance: item.importance
  };

  // Get categories if available
  if (item.categories && item.itemId) {
    item.categories.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        properties.categories = result.value || [];
      }

      // Get async properties
      if (item.subject) {
        if (typeof item.subject === 'string') {
          properties.subject = item.subject;
          logPropertiesResult(properties);
        } else {
          item.subject.getAsync((asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              properties.subject = asyncResult.value;
            }
            logPropertiesResult(properties);
          });
        }
      } else {
        logPropertiesResult(properties);
      }
    });
  } else {
    // Get async properties
    if (item.subject) {
      if (typeof item.subject === 'string') {
        properties.subject = item.subject;
        logPropertiesResult(properties);
      } else {
        item.subject.getAsync((asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            properties.subject = asyncResult.value;
          }
          logPropertiesResult(properties);
        });
      }
    } else {
      logPropertiesResult(properties);
    }
  }
}

function logPropertiesResult(properties) {
  console.log('%cItem Properties:', 'color: #f59e0b; font-weight: bold;', properties);
  logEvent('PropertiesRetrieved', 'Item properties retrieved successfully', properties);
}

// Log event to UI and console
// Log event to UI and console
function logEvent(eventType, description, data) {
  eventCounter++;

  const timestamp = new Date().toISOString();
  const verboseLogging = document.getElementById('verboseLogging').checked;
  const timestampEvents = document.getElementById('timestampEvents').checked;
  const autoScroll = document.getElementById('autoScroll').checked;

  // Store in history
  const eventRecord = {
    id: eventCounter,
    type: eventType,
    description: description,
    details: data,
    timestamp: timestamp
  };
  eventHistory.push(eventRecord);

  // Keep only last 100 events in history
  if (eventHistory.length > 100) {
    eventHistory = eventHistory.slice(-100);
  }

  // Store in localStorage (debounced to avoid excessive writes)
  saveEventsToStorage();

  // Console logging
  console.log(`[${timestamp}] ${eventType}: ${description}`, data);

  // Update UI
  const eventLog = document.getElementById('eventLog');
  const placeholder = eventLog.querySelector('.event-placeholder');
  if (placeholder) {
    placeholder.remove();
  }

  const eventItem = document.createElement('div');
  eventItem.className = 'event-item';

  const eventHeader = document.createElement('div');
  eventHeader.className = 'event-header';

  const eventTypeSpan = document.createElement('span');
  eventTypeSpan.className = 'event-type';
  eventTypeSpan.textContent = `#${eventCounter} - ${eventType}`;

  const eventTime = document.createElement('span');
  eventTime.className = 'event-time';
  if (timestampEvents) {
    eventTime.textContent = new Date(timestamp).toLocaleTimeString();
  }

  eventHeader.appendChild(eventTypeSpan);
  eventHeader.appendChild(eventTime);

  const eventDetails = document.createElement('div');
  eventDetails.className = 'event-details';
  eventDetails.textContent = description;

  eventItem.appendChild(eventHeader);
  eventItem.appendChild(eventDetails);

  if (verboseLogging && Object.keys(data).length > 0) {
    const eventData = document.createElement('div');
    eventData.className = 'event-data';
    eventData.textContent = JSON.stringify(data, null, 2);
    eventItem.appendChild(eventData);
  }

  eventLog.appendChild(eventItem);

  // Auto scroll
  if (autoScroll) {
    eventLog.scrollTop = eventLog.scrollHeight;
  }

  // Update counter
  document.getElementById('totalEvents').textContent = eventCounter;
}

// Debounced save to localStorage to prevent excessive writes
let saveTimeout = null;
function saveEventsToStorage() {
  // Clear any pending save
  if (saveTimeout) {
    clearTimeout(saveTimeout);
  }

  // Schedule a save after 500ms of no new events
  saveTimeout = setTimeout(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(eventHistory));
      console.log('💾 Events saved to storage');
    } catch (e) {
      console.warn('Could not save to localStorage:', e);
    }
  }, 500);
}

// Update active listeners count
function updateActiveListeners() {
  document.getElementById('activeListeners').textContent = activeListeners;
}

// Clear events
function clearEvents() {
  eventCounter = 0;
  eventHistory = [];

  // Clear localStorage
  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch (e) {
    console.warn('Could not clear localStorage:', e);
  }

  const eventLog = document.getElementById('eventLog');
  eventLog.innerHTML = `
        <div class="event-placeholder">
            <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                <path d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" stroke-width="2" stroke-linecap="round"/>
            </svg>
            <p>Waiting for events...</p>
            <small>Events will appear here as they occur</small>
        </div>
    `;

  document.getElementById('totalEvents').textContent = '0';

  console.log('%cEvents cleared', 'color: #ef4444; font-weight: bold;');
}

// Export events
function exportEvents() {
  if (eventHistory.length === 0) {
    alert('No events to export');
    return;
  }

  const exportData = {
    exportDate: new Date().toISOString(),
    totalEvents: eventHistory.length,
    events: eventHistory
  };

  const dataStr = JSON.stringify(exportData, null, 2);
  const dataBlob = new Blob([dataStr], { type: 'application/json' });
  const url = URL.createObjectURL(dataBlob);

  const link = document.createElement('a');
  link.href = url;
  link.download = `inboxagent-events-${Date.now()}.json`;
  link.click();

  URL.revokeObjectURL(url);

  console.log('%cEvents exported', 'color: #10b981; font-weight: bold;', exportData);
  logEvent('EventsExported', 'Event history exported to JSON file', {
    totalEvents: eventHistory.length
  });
}

// Escape HTML helper
function escapeHtml(text) {
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  return String(text).replace(/[&<>"']/g, m => map[m]);
}