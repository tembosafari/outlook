// Supabase Configuration - HUB System
const CONFIG = {
  SUPABASE_URL: 'https://xaecuidoqzbrdpqqivpl.supabase.co',
  SUPABASE_ANON_KEY: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InhhZWN1aWRvcXpicmRwcXFpdnBsIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjMzOTExMjMsImV4cCI6MjA3ODk2NzEyM30.utebNO30MJIdvkHO4_-ja2Hw21tX8gkLGV0Rb58QscQ'
};

let currentEmail = null;
let selectedEntity = null;
let attachments = [];
let currentUser = null;
let viewingLoggedEmails = false;

// DOM Elements
const elements = {};

Office.onReady(function(info) {
  if (info.host === Office.HostType.Outlook) {
    initializeElements();
    checkStoredLogin();
  }
});

function initializeElements() {
  elements.loginSection = document.getElementById('login-section');
  elements.appContent = document.getElementById('app-content');
  elements.loginEmail = document.getElementById('login-email');
  elements.loginPassword = document.getElementById('login-password');
  elements.btnLogin = document.getElementById('btn-login');
  elements.loginError = document.getElementById('login-error');
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

  // Login handlers
  elements.btnLogin.addEventListener('click', handleLogin);
  elements.loginPassword.addEventListener('keypress', function(e) {
    if (e.key === 'Enter') handleLogin();
  });
  elements.btnLogout.addEventListener('click', handleLogout);

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
    showLoginError('Indtast email og adgangskode');
    return;
  }

  elements.btnLogin.disabled = true;
  elements.btnLogin.textContent = 'Logger ind...';
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
        throw new Error('Forkert email eller adgangskode');
      }
      throw new Error('Login fejlede');
    }
    return response.json();
  })
  .then(function() {
    // Store user info
    currentUser = { email: email, password: password };
    localStorage.setItem('hub_outlook_user', JSON.stringify(currentUser));
    showMainApp();
  })
  .catch(function(error) {
    console.error('Login error:', error);
    showLoginError(error.message);
  })
  .finally(function() {
    elements.btnLogin.disabled = false;
    elements.btnLogin.textContent = 'Log ind';
  });
}

function showLoginError(message) {
  elements.loginError.textContent = message;
  elements.loginError.classList.remove('hidden');
}

function handleLogout() {
  currentUser = null;
  localStorage.removeItem('hub_outlook_user');
  elements.loginSection.classList.remove('hidden');
  elements.appContent.classList.add('hidden');
  elements.loginEmail.value = '';
  elements.loginPassword.value = '';
  elements.loginError.classList.add('hidden');
}

function showMainApp() {
  elements.loginSection.classList.add('hidden');
  elements.appContent.classList.remove('hidden');
  elements.userEmail.textContent = currentUser.email;
  loadCurrentEmail();
}

function checkForSuggestions() {
  if (!currentEmail || !currentEmail.from) return;
  
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
      email: currentUser ? currentUser.email : null,
      password: currentUser ? currentUser.password : null,
      senderEmail: currentEmail.from
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
      elements.searchResults.innerHTML = '<div class="no-results">Søg efter lead eller booking</div>';
    }
  })
  .catch(function(error) {
    console.error('Suggestion error:', error);
    elements.searchResults.innerHTML = '<div class="no-results">Søg efter lead eller booking</div>';
  });
}

function displaySuggestion(suggestion) {
  var entity = suggestion.entity;
  var isLead = suggestion.type === 'leads';
  var badgeClass = isLead ? 'lead' : 'booking';
  var badgeText = isLead ? 'Lead' : 'Booking';
  var displayName = entity.customer_name || entity.name || 'Ukendt';
  
  var meta = [];
  if (entity.email || entity.customer_email) meta.push(entity.email || entity.customer_email);
  if (entity.destination) meta.push(entity.destination);
  if (entity.booking_number) meta.push('#' + entity.booking_number);
  
  var html = '<div class="suggestion-header">📌 Foreslået baseret på tidligere emails:</div>' +
    '<div class="result-item suggested" data-id="' + entity.id + '" data-type="' + suggestion.type + '" data-name="' + displayName + '">' +
    '<div class="name">' + displayName + '<span class="badge ' + badgeClass + '">' + badgeText + '</span></div>' +
    (meta.length > 0 ? '<div class="meta">' + meta.join(' • ') + '</div>' : '') +
    '<div class="suggestion-reason">' + suggestion.reason + '</div>' +
    '</div>';
  
  elements.searchResults.innerHTML = html;
  
  var item = elements.searchResults.querySelector('.result-item');
  if (item) {
    item.addEventListener('click', function() { selectEntity(this); });
    // Auto-select the suggestion
    selectEntity(item);
  }
}

function loadCurrentEmail() {
  try {
    var item = Office.context.mailbox.item;
    if (!item) {
      showMessage('Ingen email valgt', 'error');
      return;
    }
    
    currentEmail = {
      subject: item.subject || '(Intet emne)',
      from: '',
      to: [],
      cc: [],
      date: item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString('da-DK') : '-',
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
  } catch (error) {
    console.error('Error loading email:', error);
    showMessage('Kunne ikke indlæse email: ' + error.message, 'error');
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
    elements.searchResults.innerHTML = '<div class="no-results">Indtast søgeord</div>';
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
      password: currentUser ? currentUser.password : null
    })
  })
  .then(function(response) {
    if (!response.ok) {
      if (response.status === 401) {
        handleLogout();
        throw new Error('Session udløbet - log ind igen');
      }
      throw new Error('Søgning fejlede');
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
    elements.searchResults.innerHTML = '<div class="no-results">Fejl: ' + error.message + '</div>';
  });
}

function displaySearchResults(results) {
  if (results.length === 0) {
    elements.searchResults.innerHTML = '<div class="no-results">Ingen resultater fundet</div>';
    return;
  }
  
  var html = results.map(function(result) {
    var isLead = result._type === 'leads';
    var badgeClass = isLead ? 'lead' : 'booking';
    var badgeText = isLead ? 'Lead' : 'Booking';
    var displayName = result.customer_name || result.name || 'Ukendt';
    
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
    items[i].addEventListener('click', function() { selectEntity(this); });
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

function updateActionSection() {
  if (selectedEntity) {
    elements.actionSection.classList.remove('hidden');
    var typeLabel = selectedEntity.type === 'leads' ? 'Lead' : 'Booking';
    elements.selectedEntity.textContent = 'Valgt: ' + selectedEntity.name + ' (' + typeLabel + ')';
    elements.btnLogEmail.disabled = false;
    if (elements.btnViewEmails) {
      elements.btnViewEmails.classList.remove('hidden');
    }
  } else {
    elements.actionSection.classList.add('hidden');
    elements.btnLogEmail.disabled = true;
    if (elements.btnViewEmails) {
      elements.btnViewEmails.classList.add('hidden');
    }
  }
}

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
      password: currentUser ? currentUser.password : null
    })
  })
  .then(function(response) {
    if (!response.ok) throw new Error('Kunne ikke hente emails');
    return response.json();
  })
  .then(function(data) {
    displayLoggedEmails(data.loggedEmails || []);
  })
  .catch(function(error) {
    console.error('Error loading logged emails:', error);
    elements.loggedEmailsList.innerHTML = '<div class="no-results">Fejl: ' + error.message + '</div>';
  });
}

function displayLoggedEmails(emails) {
  if (emails.length === 0) {
    elements.loggedEmailsList.innerHTML = '<div class="no-results">Ingen loggede emails på denne sag</div>';
    return;
  }
  
  var html = emails.map(function(email) {
    var date = email.received_at ? new Date(email.received_at).toLocaleDateString('da-DK', {
      day: '2-digit',
      month: 'short',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    }) : '-';
    
    var attachmentIcon = email.has_attachments ? ' 📎' : '';
    
    return '<div class="logged-email-item" data-message-id="' + (email.outlook_message_id || '') + '" data-conversation-id="' + (email.conversation_id || '') + '">' +
      '<div class="email-subject-line">' + (email.subject || '(Intet emne)') + attachmentIcon + '</div>' +
      '<div class="email-sender">' + (email.sender_name || email.sender_email || 'Ukendt') + '</div>' +
      '<div class="email-date">' + date + '</div>' +
      '<div class="email-actions">' +
        '<button class="action-btn reply-btn" title="Besvar">↩️ Besvar</button>' +
        '<button class="action-btn forward-btn" title="Videresend">➡️ Videresend</button>' +
      '</div>' +
    '</div>';
  }).join('');
  
  elements.loggedEmailsList.innerHTML = html;
  
  // Add click handlers for actions
  var items = elements.loggedEmailsList.querySelectorAll('.logged-email-item');
  items.forEach(function(item) {
    var replyBtn = item.querySelector('.reply-btn');
    var forwardBtn = item.querySelector('.forward-btn');
    var conversationId = item.getAttribute('data-conversation-id');
    var messageId = item.getAttribute('data-message-id');
    var subject = item.querySelector('.email-subject-line').textContent;
    var sender = item.querySelector('.email-sender').textContent;
    
    if (replyBtn) {
      replyBtn.addEventListener('click', function(e) {
        e.stopPropagation();
        composeReply(sender, subject, messageId);
      });
    }
    
    if (forwardBtn) {
      forwardBtn.addEventListener('click', function(e) {
        e.stopPropagation();
        composeForward(subject, messageId);
      });
    }
  });
}

function composeReply(toEmail, subject, messageId) {
  try {
    // Create reply using Office.js
    var replySubject = subject.startsWith('Re:') ? subject : 'Re: ' + subject;
    
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: [toEmail],
      subject: replySubject,
      body: ''
    });
  } catch (error) {
    console.error('Error composing reply:', error);
    showMessage('Kunne ikke åbne svar-vindue', 'error');
  }
}

function composeForward(subject, messageId) {
  try {
    // Create forward using Office.js
    var fwdSubject = subject.startsWith('Fwd:') || subject.startsWith('Fw:') ? subject : 'Fwd: ' + subject;
    
    Office.context.mailbox.displayNewMessageForm({
      subject: fwdSubject,
      body: ''
    });
  } catch (error) {
    console.error('Error composing forward:', error);
    showMessage('Kunne ikke åbne videresend-vindue', 'error');
  }
}

function backToSearch() {
  viewingLoggedEmails = false;
  elements.loggedEmailsSection.classList.add('hidden');
  elements.searchSection.classList.remove('hidden');
  elements.notesSection.classList.remove('hidden');
  updateActionSection();
}

function logEmail() {
  if (!currentEmail || !selectedEntity) {
    showMessage('Vælg venligst et lead eller booking først', 'error');
    return;
  }
  
  elements.btnLogEmail.disabled = true;
  elements.btnLogEmail.textContent = 'Logger...';
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
        subject: currentEmail.subject,
        from_email: currentEmail.from,
        to_emails: currentEmail.to,
        cc_emails: currentEmail.cc,
        html_body: currentEmail.body,
        sent_at: currentEmail.date,
        message_id: currentEmail.messageId,
        conversation_id: currentEmail.conversationId,
        lead_id: selectedEntity.type === 'leads' ? selectedEntity.id : null,
        tour_booking_id: selectedEntity.type === 'bookings' ? selectedEntity.id : null,
        notes: elements.notesInput.value.trim() || null,
        attachments: attachmentData.filter(function(a) { return a !== null; })
      };
      
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
      if (!response.ok) {
        return response.json().then(function(err) {
          throw new Error(err.error || 'Kunne ikke logge email');
        });
      }
      showMessage('Email logget succesfuldt!', 'success');
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
      showMessage('Fejl: ' + error.message, 'error');
    })
    .finally(function() {
      elements.btnLogEmail.disabled = false;
      elements.btnLogEmail.textContent = 'Log Email';
      elements.btnLogEmail.classList.remove('loading');
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
