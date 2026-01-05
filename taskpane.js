// Supabase Configuration - HUB System
const CONFIG = {
  SUPABASE_URL: 'https://xaecuidoqzbrdpqqivpl.supabase.co',
  SUPABASE_ANON_KEY: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InhhZWN1aWRvcXpicmRwcXFpdnBsIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjMzOTExMjMsImV4cCI6MjA3ODk2NzEyM30.utebNO30MJIdvkHO4_-ja2Hw21tX8gkLGV0Rb58QscQ'
};

// Helper function to format date for PostgreSQL timestamp with time zone
// Preserves local (Danish) time and formats as ISO 8601
function formatDateForDatabase(date) {
  const pad = (n) => n.toString().padStart(2, '0');
  const year = date.getFullYear();
  const month = pad(date.getMonth() + 1);
  const day = pad(date.getDate());
  const hours = pad(date.getHours());
  const minutes = pad(date.getMinutes());
  const seconds = pad(date.getSeconds());
  
  // Get timezone offset in hours and minutes
  const tzOffset = -date.getTimezoneOffset();
  const tzHours = pad(Math.floor(Math.abs(tzOffset) / 60));
  const tzMinutes = pad(Math.abs(tzOffset) % 60);
  const tzSign = tzOffset >= 0 ? '+' : '-';
  
  return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}${tzSign}${tzHours}:${tzMinutes}`;
}

let currentEmail = null;
let selectedEntity = null;
let attachments = [];
let currentUser = null;
let viewingLoggedEmails = false;
let pendingMFA = null; // For storing MFA data during verification
let isComposeMode = false; // Track if we're in compose mode
let pendingLogOnSend = null; // Store entity to log when email is sent

// DOM Elements
const elements = {};

Office.onReady(function(info) {
  if (info.host === Office.HostType.Outlook) {
    initializeElements();
    checkStoredLogin();
    
    // Register ItemChanged event for pinned taskpane support
    // When pinned, the taskpane stays open when user switches emails
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      function(eventArgs) {
        // Refresh the current email when user selects a different email
        if (currentUser) {
          loadCurrentEmail();
          // Reset selection state when email changes
          selectedEntity = null;
          viewingLoggedEmails = false;
          if (elements.searchInput) elements.searchInput.value = '';
          if (elements.searchResults) elements.searchResults.innerHTML = '';
          if (elements.selectedEntity) elements.selectedEntity.classList.add('hidden');
          if (elements.actionSection) elements.actionSection.classList.add('hidden');
          if (elements.loggedEmailsSection) elements.loggedEmailsSection.classList.add('hidden');
        }
      }
    );
  }
});

function initializeElements() {
  elements.loginSection = document.getElementById('login-section');
  elements.mfaSection = document.getElementById('mfa-section');
  elements.appContent = document.getElementById('app-content');
  elements.loginEmail = document.getElementById('login-email');
  elements.loginPassword = document.getElementById('login-password');
  elements.btnLogin = document.getElementById('btn-login');
  elements.loginError = document.getElementById('login-error');
  elements.mfaCode = document.getElementById('mfa-code');
  elements.btnVerifyMfa = document.getElementById('btn-verify-mfa');
  elements.btnCancelMfa = document.getElementById('btn-cancel-mfa');
  elements.mfaError = document.getElementById('mfa-error');
  elements.userInfo = document.getElementById('user-info');
  elements.userEmail = document.getElementById('user-email');
  elements.btnLogout = document.getElementById('btn-logout');
  elements.loading = document.getElementById('loading');
  elements.emailPreview = document.getElementById('email-preview');
  elements.emailSubject = document.getElementById('email-subject');
  elements.emailFrom = document.getElementById('email-from');
  elements.emailTo = document.getElementById('email-to');
  elements.emailDate = document.getElementById('email-date');
  elements.attachmentsSection = document.getElementById('attachments-section');
  elements.attachmentsList = document.getElementById('attachments-list');
  elements.searchSection = document.getElementById('search-section');
  elements.searchInput = document.getElementById('search-input');
  elements.searchResults = document.getElementById('search-results');
  elements.notesSection = document.getElementById('notes-section');
  elements.notesInput = document.getElementById('notes-input');
  elements.actionSection = document.getElementById('action-section');
  elements.selectedEntity = document.getElementById('selected-entity');
  elements.btnLogEmail = document.getElementById('btn-log-email');
  elements.btnViewEmails = document.getElementById('btn-view-emails');
  elements.btnSearch = document.getElementById('btn-search');
  elements.messageContainer = document.getElementById('message-container');
  elements.loggedEmailsSection = document.getElementById('logged-emails-section');
  elements.loggedEmailsList = document.getElementById('logged-emails-list');
  elements.btnBackToSearch = document.getElementById('btn-back-to-search');
  
  // Task elements
  elements.btnAddTask = document.getElementById('btn-add-task');
  elements.taskModal = document.getElementById('task-modal');
  elements.btnCloseModal = document.getElementById('btn-close-modal');
  elements.btnCancelTask = document.getElementById('btn-cancel-task');
  elements.btnSaveTask = document.getElementById('btn-save-task');
  elements.taskType = document.getElementById('task-type');
  elements.taskDate = document.getElementById('task-date');
  elements.taskTime = document.getElementById('task-time');
  elements.taskNote = document.getElementById('task-note');

  // Login handlers
  elements.btnLogin.addEventListener('click', handleLogin);
  elements.loginPassword.addEventListener('keypress', function(e) {
    if (e.key === 'Enter') handleLogin();
  });
  elements.btnLogout.addEventListener('click', handleLogout);

  // MFA handlers
  if (elements.btnVerifyMfa) {
    elements.btnVerifyMfa.addEventListener('click', handleMFAVerification);
  }
  if (elements.btnCancelMfa) {
    elements.btnCancelMfa.addEventListener('click', cancelMFA);
  }
  if (elements.mfaCode) {
    elements.mfaCode.addEventListener('keypress', function(e) {
      if (e.key === 'Enter') handleMFAVerification();
    });
  }

  // Search handlers
  elements.btnSearch.addEventListener('click', performSearch);
  elements.searchInput.addEventListener('keypress', function(e) {
    if (e.key === 'Enter') performSearch();
  });
  elements.btnLogEmail.addEventListener('click', logEmail);
  
  // View emails button
  if (elements.btnViewEmails) {
    elements.btnViewEmails.addEventListener('click', loadLoggedEmails);
  }
  
  // Back button
  if (elements.btnBackToSearch) {
    elements.btnBackToSearch.addEventListener('click', backToSearch);
  }
  
  // Task handlers
  if (elements.btnAddTask) {
    elements.btnAddTask.addEventListener('click', openTaskModal);
  }
  if (elements.btnCloseModal) {
    elements.btnCloseModal.addEventListener('click', closeTaskModal);
  }
  if (elements.btnCancelTask) {
    elements.btnCancelTask.addEventListener('click', closeTaskModal);
  }
  if (elements.btnSaveTask) {
    elements.btnSaveTask.addEventListener('click', saveTask);
  }
  if (elements.taskModal) {
    elements.taskModal.addEventListener('click', function(e) {
      if (e.target === elements.taskModal) closeTaskModal();
    });
  }
  
  // Set default date to tomorrow
  if (elements.taskDate) {
    var tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    elements.taskDate.value = tomorrow.toISOString().split('T')[0];
  }
}

function checkStoredLogin() {
  try {
    var stored = localStorage.getItem('hub_outlook_user');
    if (stored) {
      currentUser = JSON.parse(stored);
      showMainApp();
    }
  } catch (e) {
    console.error('Error checking stored login:', e);
  }
}

function handleLogin() {
  var email = elements.loginEmail.value.trim();
  var password = elements.loginPassword.value;

  if (!email || !password) {
    showLoginError('Enter email and password');
    return;
  }

  elements.btnLogin.disabled = true;
  elements.btnLogin.textContent = 'Signing in...';
  elements.loginError.classList.add('hidden');

  // Verify credentials via edge function
  fetch(CONFIG.SUPABASE_URL + '/functions/v1/search-hub-entities', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CONFIG.SUPABASE_ANON_KEY
    },
    body: JSON.stringify({ 
      query: 'test', 
      type: 'all', 
      limit: 1,
      email: email,
      password: password
    })
  })
  .then(function(response) {
    if (!response.ok) {
      if (response.status === 401) {
        throw new Error('Incorrect email or password');
      }
      throw new Error('Login failed');
    }
    return response.json();
  })
  .then(function(data) {
    // Check if MFA is required
    if (data.mfaRequired) {
      console.log('MFA required, showing MFA screen');
      pendingMFA = { email: email, password: password, factorId: data.factorId };
      showMFASection();
      return;
    }
    
    // Store user info (without mfaVerified flag if no MFA)
    currentUser = { email: email, password: password, mfaVerified: true };
    localStorage.setItem('hub_outlook_user', JSON.stringify(currentUser));
    showMainApp();
  })
  .catch(function(error) {
    console.error('Login error:', error);
    showLoginError(error.message);
  })
  .finally(function() {
    elements.btnLogin.disabled = false;
    elements.btnLogin.textContent = 'Sign in';
  });
}

function showMFASection() {
  elements.loginSection.classList.add('hidden');
  elements.mfaSection.classList.remove('hidden');
  elements.mfaCode.value = '';
  elements.mfaError.classList.add('hidden');
  elements.mfaCode.focus();
}

function cancelMFA() {
  pendingMFA = null;
  elements.mfaSection.classList.add('hidden');
  elements.loginSection.classList.remove('hidden');
  elements.mfaCode.value = '';
}

function handleMFAVerification() {
  var code = elements.mfaCode.value.trim();
  
  if (!code || code.length !== 6) {
    showMFAError('Enter a 6-digit code');
    return;
  }
  
  if (!pendingMFA) {
    showMFAError('Session expired - please try again');
    cancelMFA();
    return;
  }
  
  elements.btnVerifyMfa.disabled = true;
  elements.btnVerifyMfa.textContent = 'Verifying...';
  elements.mfaError.classList.add('hidden');
  
  // Verify MFA code via edge function
  fetch(CONFIG.SUPABASE_URL + '/functions/v1/verify-mfa', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CONFIG.SUPABASE_ANON_KEY
    },
    body: JSON.stringify({
      email: pendingMFA.email,
      password: pendingMFA.password,
      factorId: pendingMFA.factorId,
      code: code
    })
  })
  .then(function(response) {
    if (!response.ok) {
      throw new Error('Invalid code - please try again');
    }
    return response.json();
  })
  .then(function(data) {
    if (data.error) {
      throw new Error(data.error);
    }
    
    // MFA verified successfully
    currentUser = { 
      email: pendingMFA.email, 
      password: pendingMFA.password, 
      mfaVerified: true 
    };
    localStorage.setItem('hub_outlook_user', JSON.stringify(currentUser));
    pendingMFA = null;
    elements.mfaSection.classList.add('hidden');
    showMainApp();
  })
  .catch(function(error) {
    console.error('MFA error:', error);
    showMFAError(error.message);
  })
  .finally(function() {
    elements.btnVerifyMfa.disabled = false;
    elements.btnVerifyMfa.textContent = 'Verify';
  });
}

function showMFAError(message) {
  elements.mfaError.textContent = message;
  elements.mfaError.classList.remove('hidden');
}

function showLoginError(message) {
  elements.loginError.textContent = message;
  elements.loginError.classList.remove('hidden');
}

function handleLogout() {
  currentUser = null;
  pendingMFA = null;
  localStorage.removeItem('hub_outlook_user');
  elements.loginSection.classList.remove('hidden');
  elements.mfaSection.classList.add('hidden');
  elements.appContent.classList.add('hidden');
  elements.loginEmail.value = '';
  elements.loginPassword.value = '';
  elements.mfaCode.value = '';
  elements.loginError.classList.add('hidden');
}

function showMainApp() {
  elements.loginSection.classList.add('hidden');
  elements.mfaSection.classList.add('hidden');
  elements.appContent.classList.remove('hidden');
  elements.userEmail.textContent = currentUser.email;
  loadCurrentEmail();
}

function checkForSuggestions() {
  if (!currentEmail || !currentEmail.from) {
    console.log('No email or sender for suggestions');
    return;
  }
  
  if (!currentUser || !currentUser.email || !currentUser.password) {
    console.log('No user credentials for suggestions');
    elements.searchResults.innerHTML = '<div class="no-results">Search for lead or booking</div>';
    return;
  }
  
  console.log('Checking for suggestions for sender:', currentEmail.from);
  elements.searchResults.innerHTML = '<div class="loading"><div class="spinner"></div></div>';
  
  fetch(CONFIG.SUPABASE_URL + '/functions/v1/search-hub-entities', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CONFIG.SUPABASE_ANON_KEY
    },
    body: JSON.stringify({ 
      query: '', 
      type: 'all',
      email: currentUser.email,
      password: currentUser.password,
      mfaVerified: currentUser.mfaVerified || false,
      senderEmail: currentEmail.from
    })
  })
  .then(function(response) {
    console.log('Suggestions response status:', response.status);
    if (!response.ok) return null;
    return response.json();
  })
  .then(function(data) {
    console.log('Suggestions data:', data);
    if (data && data.suggestion) {
      displaySuggestion(data.suggestion);
    } else {
      elements.searchResults.innerHTML = '<div class="no-results">Search for lead or booking</div>';
    }
  })
  .catch(function(error) {
    console.error('Suggestion error:', error);
    elements.searchResults.innerHTML = '<div class="no-results">Search for lead or booking</div>';
  });
}

function displaySuggestion(suggestion) {
  var entity = suggestion.entity;
  var isLead = suggestion.type === 'leads';
  var badgeClass = isLead ? 'lead' : 'booking';
  var badgeText = isLead ? 'Lead' : 'Booking';
  var displayName = entity.customer_name || entity.name || 'Unknown';
  
  var meta = [];
  if (entity.email || entity.customer_email) meta.push(entity.email || entity.customer_email);
  if (entity.destination) meta.push(entity.destination);
  if (entity.booking_number) meta.push('#' + entity.booking_number);
  
  var html = '<div class="suggestion-header">📌 Suggested based on previous emails:</div>' +
    '<div class="result-item suggested" data-id="' + entity.id + '" data-type="' + suggestion.type + '" data-name="' + displayName + '">' +
    '<div class="name">' + displayName + '<span class="badge ' + badgeClass + '">' + badgeText + '</span></div>' +
    (meta.length > 0 ? '<div class="meta">' + meta.join(' • ') + '</div>' : '') +
    '<div class="suggestion-reason">' + suggestion.reason + '</div>' +
    '</div>';
  
  elements.searchResults.innerHTML = html;
  
  var item = elements.searchResults.querySelector('.result-item');
  if (item) {
    item.addEventListener('click', function(e) { 
      selectEntity(e.currentTarget); 
    });
    item.addEventListener('dblclick', function(e) { 
      selectEntityAndLog(e.currentTarget); 
    });
    // Auto-select the suggestion
    selectEntity(item);
  }
}

function loadCurrentEmail() {
  try {
    var item = Office.context.mailbox.item;
    if (!item) {
      showMessage('No email selected', 'error');
      return;
    }
    
    // Detect if we're in compose mode (drafting/replying)
    // In compose mode, item.itemType is 'message' but item.from is undefined
    // and we use item.to.getAsync instead of item.to directly
    isComposeMode = typeof item.to === 'object' && typeof item.to.getAsync === 'function';
    
    console.log('Loading email, compose mode:', isComposeMode);
    
    if (isComposeMode) {
      loadComposeEmail(item);
    } else {
      loadReadEmail(item);
    }
  } catch (error) {
    console.error('Error loading email:', error);
    showMessage('Could not load email: ' + error.message, 'error');
  }
}

// Load email in READ mode (received emails)
function loadReadEmail(item) {
  currentEmail = {
    subject: item.subject || '(No subject)',
    from: '',
    to: [],
    cc: [],
    date: item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString('en-GB') : '-',
    dateISO: item.dateTimeCreated ? formatDateForDatabase(new Date(item.dateTimeCreated)) : formatDateForDatabase(new Date()),
    body: '',
    messageId: item.internetMessageId || item.itemId,
    conversationId: item.conversationId || null
  };
  
  if (item.from) {
    currentEmail.from = item.from.emailAddress || item.from.displayName || '-';
  }
  
  if (item.to && item.to.length > 0) {
    currentEmail.to = item.to.map(function(r) { return r.emailAddress || r.displayName; });
  }
  
  if (item.cc && item.cc.length > 0) {
    currentEmail.cc = item.cc.map(function(r) { return r.emailAddress || r.displayName; });
  }
  
  item.body.getAsync(Office.CoercionType.Html, function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      currentEmail.body = result.value;
    }
  });
  
  if (item.attachments && item.attachments.length > 0) {
    attachments = [];
    for (var i = 0; i < item.attachments.length; i++) {
      var att = item.attachments[i];
      if (!att.isInline) {
        attachments.push({
          name: att.name,
          contentType: att.contentType,
          size: att.size,
          id: att.id
        });
      }
    }
  }
  
  updateEmailPreview();
}

// Load email in COMPOSE mode (drafting/replying)
function loadComposeEmail(item) {
  currentEmail = {
    subject: '',
    from: currentUser ? currentUser.email : '',
    to: [],
    cc: [],
    date: new Date().toLocaleString('en-GB'),
    dateISO: formatDateForDatabase(new Date()),
    body: '',
    messageId: 'compose-' + Date.now(),
    conversationId: item.conversationId || null
  };
  
  // Get subject asynchronously in compose mode
  item.subject.getAsync(function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      currentEmail.subject = result.value || '(No subject)';
      if (elements.emailSubject) {
        elements.emailSubject.textContent = currentEmail.subject;
      }
    }
  });
  
  // Get recipients asynchronously
  item.to.getAsync(function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
      currentEmail.to = result.value.map(function(r) { return r.emailAddress || r.displayName; });
      if (elements.emailTo) {
        elements.emailTo.textContent = currentEmail.to.join(', ') || '-';
      }
      // In compose mode, search based on recipient email
      if (currentEmail.to.length > 0) {
        searchByRecipientEmail(currentEmail.to[0]);
      }
    }
  });
  
  // Get body
  item.body.getAsync(Office.CoercionType.Html, function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      currentEmail.body = result.value;
    }
  });
  
  // Reset attachments for compose mode
  attachments = [];
  
  updateEmailPreviewCompose();
}

// Search for lead/booking based on recipient email in compose mode
function searchByRecipientEmail(recipientEmail) {
  if (!recipientEmail || !currentUser) {
    return;
  }
  
  console.log('Searching for recipient email:', recipientEmail);
  elements.searchResults.innerHTML = '<div class="loading"><div class="spinner"></div></div>';
  
  fetch(CONFIG.SUPABASE_URL + '/functions/v1/search-hub-entities', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CONFIG.SUPABASE_ANON_KEY
    },
    body: JSON.stringify({ 
      query: '', 
      type: 'all',
      email: currentUser.email,
      password: currentUser.password,
      mfaVerified: currentUser.mfaVerified || false,
      senderEmail: recipientEmail // Use recipient as "sender" to find matching entity
    })
  })
  .then(function(response) {
    if (!response.ok) return null;
    return response.json();
  })
  .then(function(data) {
    if (data && data.suggestion) {
      displaySuggestion(data.suggestion);
    } else {
      elements.searchResults.innerHTML = '<div class="no-results">Search for lead or booking</div>';
    }
  })
  .catch(function(error) {
    console.error('Recipient search error:', error);
    elements.searchResults.innerHTML = '<div class="no-results">Search for lead or booking</div>';
  });
}

// Update preview for compose mode
function updateEmailPreviewCompose() {
  elements.loading.classList.add('hidden');
  elements.emailPreview.classList.remove('hidden');
  elements.searchSection.classList.remove('hidden');
  elements.notesSection.classList.remove('hidden');
  
  elements.emailSubject.textContent = currentEmail.subject || '(Loading...)';
  elements.emailFrom.textContent = 'You';
  elements.emailTo.textContent = currentEmail.to.join(', ') || '(Loading...)';
  elements.emailDate.textContent = 'Draft';
  
  // Hide attachments section in compose mode
  elements.attachmentsSection.classList.add('hidden');
  
  // Update the log button text for compose mode
  if (elements.btnLogEmail) {
    elements.btnLogEmail.textContent = 'Log when sent';
  }
}

function updateEmailPreview() {
  elements.loading.classList.add('hidden');
  elements.emailPreview.classList.remove('hidden');
  elements.searchSection.classList.remove('hidden');
  elements.notesSection.classList.remove('hidden');
  
  elements.emailSubject.textContent = currentEmail.subject;
  elements.emailFrom.textContent = currentEmail.from;
  elements.emailTo.textContent = currentEmail.to.join(', ') || '-';
  elements.emailDate.textContent = currentEmail.date;
  
  if (attachments.length > 0) {
    elements.attachmentsSection.classList.remove('hidden');
    elements.attachmentsList.innerHTML = attachments.map(function(att) {
      return '<li>' + att.name + ' (' + formatFileSize(att.size) + ')</li>';
    }).join('');
  }
  
  // Check for suggestions based on sender email
  checkForSuggestions();
}

function formatFileSize(bytes) {
  if (!bytes) return '-';
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function performSearch() {
  var query = elements.searchInput.value.trim();
  if (!query) {
    elements.searchResults.innerHTML = '<div class="no-results">Enter search term</div>';
    return;
  }
  
  elements.searchResults.innerHTML = '<div class="loading"><div class="spinner"></div></div>';
  
  fetch(CONFIG.SUPABASE_URL + '/functions/v1/search-hub-entities', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CONFIG.SUPABASE_ANON_KEY
    },
    body: JSON.stringify({ 
      query: query, 
      type: 'all',
      email: currentUser ? currentUser.email : null,
      password: currentUser ? currentUser.password : null,
      mfaVerified: currentUser ? currentUser.mfaVerified : false
    })
  })
  .then(function(response) {
    if (!response.ok) {
      if (response.status === 401) {
        handleLogout();
        throw new Error('Session expired - please sign in again');
      }
      throw new Error('Search failed');
    }
    return response.json();
  })
  .then(function(data) {
    // Combine leads and bookings with type info
    var results = [];
    if (data.leads) {
      data.leads.forEach(function(lead) {
        lead._type = 'leads';
        results.push(lead);
      });
    }
    if (data.bookings) {
      data.bookings.forEach(function(booking) {
        booking._type = 'bookings';
        results.push(booking);
      });
    }
    displaySearchResults(results);
  })
  .catch(function(error) {
    console.error('Search error:', error);
    elements.searchResults.innerHTML = '<div class="no-results">Error: ' + error.message + '</div>';
  });
}

function displaySearchResults(results) {
  if (results.length === 0) {
    elements.searchResults.innerHTML = '<div class="no-results">No results found</div>';
    return;
  }
  
  var html = results.map(function(result) {
    var isLead = result._type === 'leads';
    var badgeClass = isLead ? 'lead' : 'booking';
    var badgeText = isLead ? 'Lead' : 'Booking';
    var displayName = result.customer_name || result.name || 'Unknown';
    
    var meta = [];
    if (result.email || result.customer_email) meta.push(result.email || result.customer_email);
    if (result.destination) meta.push(result.destination);
    if (result.booking_number) meta.push('#' + result.booking_number);
    
    return '<div class="result-item" data-id="' + result.id + '" data-type="' + result._type + '" data-name="' + displayName + '">' +
      '<div class="name">' + displayName + '<span class="badge ' + badgeClass + '">' + badgeText + '</span></div>' +
      (meta.length > 0 ? '<div class="meta">' + meta.join(' • ') + '</div>' : '') +
      '</div>';
  }).join('');
  
  elements.searchResults.innerHTML = html;
  
  var items = elements.searchResults.querySelectorAll('.result-item');
  for (var i = 0; i < items.length; i++) {
    (function(item) {
      item.addEventListener('click', function(e) { 
        selectEntity(e.currentTarget); 
      });
      item.addEventListener('dblclick', function(e) { 
        selectEntityAndLog(e.currentTarget); 
      });
    })(items[i]);
  }
}

function selectEntity(element) {
  var items = elements.searchResults.querySelectorAll('.result-item');
  for (var i = 0; i < items.length; i++) {
    items[i].classList.remove('selected');
  }
  element.classList.add('selected');
  
  selectedEntity = {
    id: element.getAttribute('data-id'),
    type: element.getAttribute('data-type'),
    name: element.getAttribute('data-name')
  };
  
  updateActionSection();
}

// Double-click handler - select and immediately log
function selectEntityAndLog(element) {
  selectEntity(element);
  // Small delay to ensure selection is processed
  setTimeout(function() {
    logEmail();
  }, 100);
}

function updateActionSection() {
  if (selectedEntity) {
    elements.actionSection.classList.remove('hidden');
    var typeLabel = selectedEntity.type === 'leads' ? 'Lead' : 'Booking';
    elements.selectedEntity.textContent = 'Selected: ' + selectedEntity.name + ' (' + typeLabel + ')';
    elements.btnLogEmail.disabled = false;
    
    // Update button text based on mode
    if (isComposeMode) {
      elements.btnLogEmail.textContent = '📧 Log when sent';
    } else {
      elements.btnLogEmail.textContent = '📧 Log Email';
    }
    
    if (elements.btnAddTask) {
      elements.btnAddTask.disabled = false;
    }
    if (elements.btnViewEmails) {
      elements.btnViewEmails.classList.remove('hidden');
    }
  } else {
    elements.actionSection.classList.add('hidden');
    elements.btnLogEmail.disabled = true;
    if (elements.btnAddTask) {
      elements.btnAddTask.disabled = true;
    }
    if (elements.btnViewEmails) {
      elements.btnViewEmails.classList.add('hidden');
    }
  }
}

// ==================== TASK FUNCTIONS ====================

function openTaskModal() {
  if (!selectedEntity) {
    showMessage('Please select a lead or booking first', 'error');
    return;
  }
  
  // Reset form
  elements.taskType.value = 'task';
  elements.taskNote.value = '';
  
  // Set default date to tomorrow
  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  elements.taskDate.value = tomorrow.toISOString().split('T')[0];
  elements.taskTime.value = '09:00';
  
  elements.taskModal.classList.remove('hidden');
}

function closeTaskModal() {
  elements.taskModal.classList.add('hidden');
}

function saveTask() {
  if (!selectedEntity) {
    showMessage('Please select a lead or booking first', 'error');
    closeTaskModal();
    return;
  }
  
  if (!currentUser || !currentUser.email || !currentUser.password) {
    showMessage('Session expired - please sign in again', 'error');
    handleLogout();
    return;
  }
  
  var taskDate = elements.taskDate.value;
  var taskTime = elements.taskTime.value || '09:00';
  var taskType = elements.taskType.value;
  var taskNote = elements.taskNote.value.trim();
  
  if (!taskDate) {
    showMessage('Please select a date', 'error');
    return;
  }
  
  // Combine date and time into ISO datetime
  var reminderDatetime = taskDate + 'T' + taskTime + ':00';
  
  elements.btnSaveTask.disabled = true;
  elements.btnSaveTask.textContent = 'Saving...';
  
  var payload = {
    email: currentUser.email,
    password: currentUser.password,
    mfaVerified: currentUser.mfaVerified || false,
    reminderDatetime: reminderDatetime,
    actionType: taskType,
    note: taskNote || null,
    lead_id: selectedEntity.type === 'leads' ? selectedEntity.id : null,
    tour_booking_id: selectedEntity.type === 'bookings' ? selectedEntity.id : null
  };
  
  console.log('Creating task:', payload);
  
  fetch(CONFIG.SUPABASE_URL + '/functions/v1/create-outlook-reminder', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CONFIG.SUPABASE_ANON_KEY
    },
    body: JSON.stringify(payload)
  })
  .then(function(response) {
    if (!response.ok) {
      return response.json().then(function(err) {
        throw new Error(err.error || 'Could not create task');
      });
    }
    return response.json();
  })
  .then(function(data) {
    console.log('Task created:', data);
    showMessage('Task created!', 'success');
    closeTaskModal();
    setTimeout(hideMessage, 3000);
  })
  .catch(function(error) {
    console.error('Create task error:', error);
    showMessage('Error: ' + error.message, 'error');
  })
  .finally(function() {
    elements.btnSaveTask.disabled = false;
    elements.btnSaveTask.textContent = 'Save Task';
  });
}

// ==================== LOGGED EMAILS FUNCTIONS ====================

function loadLoggedEmails() {
  if (!selectedEntity) return;
  
  viewingLoggedEmails = true;
  
  // Hide search section and show logged emails section
  elements.searchSection.classList.add('hidden');
  elements.notesSection.classList.add('hidden');
  elements.actionSection.classList.add('hidden');
  elements.loggedEmailsSection.classList.remove('hidden');
  elements.loggedEmailsList.innerHTML = '<div class="loading"><div class="spinner"></div></div>';
  
  var entityType = selectedEntity.type === 'leads' ? 'lead' : 'booking';
  
  fetch(CONFIG.SUPABASE_URL + '/functions/v1/search-hub-entities', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CONFIG.SUPABASE_ANON_KEY
    },
    body: JSON.stringify({ 
      entityId: selectedEntity.id,
      entityType: entityType,
      email: currentUser ? currentUser.email : null,
      password: currentUser ? currentUser.password : null,
      mfaVerified: currentUser ? currentUser.mfaVerified : false
    })
  })
  .then(function(response) {
    if (!response.ok) throw new Error('Could not fetch emails');
    return response.json();
  })
  .then(function(data) {
    displayLoggedEmails(data.loggedEmails || []);
  })
  .catch(function(error) {
    console.error('Error loading logged emails:', error);
    elements.loggedEmailsList.innerHTML = '<div class="no-results">Error: ' + error.message + '</div>';
  });
}

function displayLoggedEmails(emails) {
  if (emails.length === 0) {
    elements.loggedEmailsList.innerHTML = '<div class="no-results">No logged emails on this case</div>';
    return;
  }
  
  var html = emails.map(function(email, index) {
    var date = email.received_at ? new Date(email.received_at).toLocaleDateString('en-GB', {
      day: '2-digit',
      month: 'short',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    }) : '-';
    
    var attachmentIcon = email.has_attachments ? ' 📎' : '';
    
    // Get email body - prefer text, fallback to html stripped
    var bodyContent = email.body_text || '';
    if (!bodyContent && email.body_html) {
      var tempDiv = document.createElement('div');
      tempDiv.innerHTML = email.body_html;
      bodyContent = tempDiv.textContent || tempDiv.innerText || '';
    }
    // Truncate body for preview
    var bodyPreview = bodyContent.substring(0, 500);
    if (bodyContent.length > 500) bodyPreview += '...';
    
    var toRecipients = Array.isArray(email.recipient_emails) ? email.recipient_emails.join(', ') : (email.to_emails || '');
    
    return '<div class="logged-email-item" data-email-index="' + index + '">' +
      '<div class="email-subject-line">' + escapeHtml(email.subject || '(No subject)') + attachmentIcon + '</div>' +
      '<div class="email-sender">From: ' + escapeHtml(email.sender_name || email.sender_email || email.from_email || 'Unknown') + '</div>' +
      '<div class="email-date">' + date + '</div>' +
      '<div class="expand-hint">Click to view content ▼</div>' +
      '<div class="email-body-preview">' +
        '<div class="email-meta">' +
          '<div><strong>From:</strong> ' + escapeHtml(email.sender_name || email.sender_email || email.from_email || '-') + '</div>' +
          '<div><strong>To:</strong> ' + escapeHtml(toRecipients || '-') + '</div>' +
          '<div><strong>Date:</strong> ' + date + '</div>' +
        '</div>' +
        '<div class="email-body-content">' + escapeHtml(bodyPreview || '(No content)') + '</div>' +
        '<div class="email-actions">' +
          '<button class="action-btn reply-btn" title="Reply">↩️ Reply</button>' +
          '<button class="action-btn forward-btn" title="Forward">➡️ Forward</button>' +
        '</div>' +
      '</div>' +
    '</div>';
  }).join('');
  
  // Store emails for reply/forward
  window.loggedEmailsData = emails;
  
  elements.loggedEmailsList.innerHTML = html;
  
  // Add click handlers for expand/collapse and actions
  var items = elements.loggedEmailsList.querySelectorAll('.logged-email-item');
  items.forEach(function(item) {
    var emailIndex = parseInt(item.getAttribute('data-email-index'), 10);
    var emailData = window.loggedEmailsData[emailIndex];
    
    // Click on item to expand/collapse
    item.addEventListener('click', function(e) {
      // Don't toggle if clicking on a button
      if (e.target.closest('.action-btn')) return;
      item.classList.toggle('expanded');
      var hint = item.querySelector('.expand-hint');
      if (hint) {
        hint.textContent = item.classList.contains('expanded') ? 'Click to hide ▲' : 'Click to view content ▼';
      }
    });
    
    var replyBtn = item.querySelector('.reply-btn');
    var forwardBtn = item.querySelector('.forward-btn');
    
    if (replyBtn) {
      replyBtn.addEventListener('click', function(e) {
        e.stopPropagation();
        composeReply(emailData);
      });
    }
    
    if (forwardBtn) {
      forwardBtn.addEventListener('click', function(e) {
        e.stopPropagation();
        composeForward(emailData);
      });
    }
  });
}

function escapeHtml(text) {
  if (!text) return '';
  var div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

function composeReply(emailData) {
  try {
    var subject = emailData.subject || '';
    var replySubject = subject.startsWith('Re:') ? subject : 'Re: ' + subject;
    var toEmail = emailData.from_email || emailData.sender_email || '';
    
    // Build reply body with original email quoted
    var originalDate = emailData.received_at ? new Date(emailData.received_at).toLocaleString('en-GB') : '';
    var originalFrom = emailData.sender_name || emailData.from_email || 'Unknown';
    var originalBody = emailData.body_text || emailData.body_html || '';
    
    // Strip HTML if we only have HTML body
    if (!emailData.body_text && emailData.body_html) {
      var tempDiv = document.createElement('div');
      tempDiv.innerHTML = emailData.body_html;
      originalBody = tempDiv.textContent || tempDiv.innerText || '';
    }
    
    var replyBody = '\n\n\n-------- Original message --------\n' +
      'From: ' + originalFrom + '\n' +
      'Date: ' + originalDate + '\n' +
      'Subject: ' + subject + '\n\n' +
      originalBody;
    
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: [toEmail],
      subject: replySubject,
      body: replyBody
    });
  } catch (error) {
    console.error('Error composing reply:', error);
    showMessage('Could not open reply window', 'error');
  }
}

function composeForward(emailData) {
  try {
    var subject = emailData.subject || '';
    var fwdSubject = subject.startsWith('Fwd:') || subject.startsWith('Fw:') ? subject : 'Fwd: ' + subject;
    
    // Build forward body with original email
    var originalDate = emailData.received_at ? new Date(emailData.received_at).toLocaleString('en-GB') : '';
    var originalFrom = emailData.sender_name || emailData.from_email || 'Unknown';
    var originalTo = emailData.to_emails || '';
    var originalBody = emailData.body_text || emailData.body_html || '';
    
    // Strip HTML if we only have HTML body
    if (!emailData.body_text && emailData.body_html) {
      var tempDiv = document.createElement('div');
      tempDiv.innerHTML = emailData.body_html;
      originalBody = tempDiv.textContent || tempDiv.innerText || '';
    }
    
    var forwardBody = '\n\n\n-------- Forwarded message --------\n' +
      'From: ' + originalFrom + '\n' +
      'Date: ' + originalDate + '\n' +
      'To: ' + originalTo + '\n' +
      'Subject: ' + subject + '\n\n' +
      originalBody;
    
    Office.context.mailbox.displayNewMessageForm({
      subject: fwdSubject,
      body: forwardBody
    });
  } catch (error) {
    console.error('Error composing forward:', error);
    showMessage('Could not open forward window', 'error');
  }
}

function backToSearch() {
  viewingLoggedEmails = false;
  elements.loggedEmailsSection.classList.add('hidden');
  elements.searchSection.classList.remove('hidden');
  elements.notesSection.classList.remove('hidden');
  updateActionSection();
}

// ==================== EMAIL LOGGING ====================

function logEmail() {
  if (!currentEmail || !selectedEntity) {
    showMessage('Please select a lead or booking first', 'error');
    return;
  }
  
  if (!currentUser || !currentUser.email || !currentUser.password) {
    showMessage('Session expired - please sign in again', 'error');
    handleLogout();
    return;
  }
  
  // Handle compose mode differently - mark for logging when sent
  if (isComposeMode) {
    handleComposeModelog();
    return;
  }
  
  console.log('Logging email with credentials:', { 
    hasEmail: !!currentUser.email, 
    hasPassword: !!currentUser.password,
    entity: selectedEntity
  });
  
  elements.btnLogEmail.disabled = true;
  elements.btnLogEmail.textContent = 'Logging...';
  elements.btnLogEmail.classList.add('loading');
  
  var attachmentPromises = [];
  if (attachments.length > 0) {
    for (var i = 0; i < attachments.length; i++) {
      attachmentPromises.push(getAttachmentContent(attachments[i]));
    }
  }
  
  Promise.all(attachmentPromises)
    .then(function(attachmentData) {
      var payload = {
        // Auth credentials - MUST be first and always present
        email: currentUser.email,
        password: currentUser.password,
        mfaVerified: currentUser.mfaVerified || false,
        // Email data
        subject: currentEmail.subject,
        from_email: currentEmail.from,
        to_emails: currentEmail.to,
        cc_emails: currentEmail.cc,
        html_body: currentEmail.body,
        sent_at: currentEmail.dateISO,
        message_id: currentEmail.messageId,
        conversation_id: currentEmail.conversationId,
        lead_id: selectedEntity.type === 'leads' ? selectedEntity.id : null,
        tour_booking_id: selectedEntity.type === 'bookings' ? selectedEntity.id : null,
        notes: elements.notesInput.value.trim() || null,
        // Attachments with correct field names matching edge function
        attachments: attachmentData.filter(function(a) { return a !== null; }).map(function(att) {
          return {
            fileName: att.name,
            contentType: att.content_type,
            content: att.content_base64,
            size: att.size
          };
        })
      };
      
      console.log('Sending payload with keys:', Object.keys(payload));
      
      return fetch(CONFIG.SUPABASE_URL + '/functions/v1/log-outlook-email', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + CONFIG.SUPABASE_ANON_KEY
        },
        body: JSON.stringify(payload)
      });
    })
    .then(function(response) {
      return response.json().then(function(data) {
        if (!response.ok) {
          // Check for duplicate error
          if (response.status === 409 && data.duplicate) {
            throw new Error('This email is already logged on this case');
          }
          throw new Error(data.error || 'Could not log email');
        }
        return data;
      });
    })
    .then(function(data) {
      showMessage('Email logged successfully!', 'success');
      setTimeout(function() {
        elements.notesInput.value = '';
        selectedEntity = null;
        elements.searchResults.innerHTML = '';
        elements.searchInput.value = '';
        updateActionSection();
        hideMessage();
      }, 2000);
    })
    .catch(function(error) {
      console.error('Log email error:', error);
      showMessage('Error: ' + error.message, 'error');
    })
    .finally(function() {
      elements.btnLogEmail.disabled = false;
      elements.btnLogEmail.textContent = isComposeMode ? 'Log when sent' : '📧 Log Email';
      elements.btnLogEmail.classList.remove('loading');
    });
}

// Handle logging in compose mode - captures email data to log when user sends
function handleComposeModelog() {
  var item = Office.context.mailbox.item;
  
  // First update currentEmail with latest compose content
  elements.btnLogEmail.disabled = true;
  elements.btnLogEmail.textContent = 'Preparing...';
  
  // Get subject
  item.subject.getAsync(function(subjectResult) {
    if (subjectResult.status === Office.AsyncResultStatus.Succeeded) {
      currentEmail.subject = subjectResult.value || '(No subject)';
    }
    
    // Get recipients
    item.to.getAsync(function(toResult) {
      if (toResult.status === Office.AsyncResultStatus.Succeeded && toResult.value) {
        currentEmail.to = toResult.value.map(function(r) { return r.emailAddress || r.displayName; });
      }
      
      // Get CC
      item.cc.getAsync(function(ccResult) {
        if (ccResult.status === Office.AsyncResultStatus.Succeeded && ccResult.value) {
          currentEmail.cc = ccResult.value.map(function(r) { return r.emailAddress || r.displayName; });
        }
        
        // Get body
        item.body.getAsync(Office.CoercionType.Html, function(bodyResult) {
          if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
            currentEmail.body = bodyResult.value;
          }
          
          // Now log the email immediately with current draft content
          // Update timestamp to now since this is when they're logging it
          currentEmail.dateISO = formatDateForDatabase(new Date());
          currentEmail.from = currentUser.email;
          
          // Generate a unique message ID for compose mode
          currentEmail.messageId = 'compose-' + Date.now() + '-' + Math.random().toString(36).substr(2, 9);
          
          // Now send to backend
          var payload = {
            email: currentUser.email,
            password: currentUser.password,
            mfaVerified: currentUser.mfaVerified || false,
            subject: currentEmail.subject,
            from_email: currentEmail.from,
            to_emails: currentEmail.to,
            cc_emails: currentEmail.cc,
            html_body: currentEmail.body,
            sent_at: currentEmail.dateISO,
            message_id: currentEmail.messageId,
            conversation_id: currentEmail.conversationId,
            lead_id: selectedEntity.type === 'leads' ? selectedEntity.id : null,
            tour_booking_id: selectedEntity.type === 'bookings' ? selectedEntity.id : null,
            notes: (elements.notesInput.value.trim() || '') + ' [Logged during composition]',
            attachments: []
          };
          
          console.log('Logging compose email:', payload);
          
          fetch(CONFIG.SUPABASE_URL + '/functions/v1/log-outlook-email', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'Authorization': 'Bearer ' + CONFIG.SUPABASE_ANON_KEY
            },
            body: JSON.stringify(payload)
          })
          .then(function(response) {
            return response.json().then(function(data) {
              if (!response.ok) {
                if (response.status === 409 && data.duplicate) {
                  throw new Error('This email is already logged on this case');
                }
                throw new Error(data.error || 'Could not log email');
              }
              return data;
            });
          })
          .then(function(data) {
            showMessage('Draft logged! Continue composing and send when ready.', 'success');
            setTimeout(function() {
              hideMessage();
            }, 4000);
          })
          .catch(function(error) {
            console.error('Log compose email error:', error);
            showMessage('Error: ' + error.message, 'error');
          })
          .finally(function() {
            elements.btnLogEmail.disabled = false;
            elements.btnLogEmail.textContent = 'Log when sent';
          });
        });
      });
    });
  });
}

function getAttachmentContent(att) {
  return new Promise(function(resolve) {
    Office.context.mailbox.item.getAttachmentContentAsync(att.id, function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve({
          name: att.name,
          content_type: att.contentType,
          size: att.size,
          content_base64: result.value.content
        });
      } else {
        resolve(null);
      }
    });
  });
}

function showMessage(text, type) {
  elements.messageContainer.textContent = text;
  elements.messageContainer.className = 'message-container ' + type;
  elements.messageContainer.classList.remove('hidden');
}

function hideMessage() {
  elements.messageContainer.classList.add('hidden');
}
