// Configuration - Update these with your actual values
const CONFIG = {
  // Replace with your actual Supabase URL and anon key
  SUPABASE_URL: 'https://xaecuidoqzbrdpqqivpl.supabase.co',
  SUPABASE_ANON_KEY: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InhhZWN1aWRvcXpicmRwcXFpdnBsIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjMzOTExMjMsImV4cCI6MjA3ODk2NzEyM30.utebNO30MJIdvkHO4_-ja2Hw21tX8gkLGV0Rb58QscQ'
};

// State
let currentEmail = null;
let selectedEntity = null;
let searchType = 'leads'; // 'leads' or 'bookings'
let attachments = [];

// DOM Elements
const elements = {
  loading: document.getElementById('loading'),
  emailPreview: document.getElementById('email-preview'),
  emailSubject: document.getElementById('email-subject'),
  emailFrom: document.getElementById('email-from'),
  emailTo: document.getElementById('email-to'),
  emailDate: document.getElementById('email-date'),
  attachmentsSection: document.getElementById('attachments-section'),
  attachmentsList: document.getElementById('attachments-list'),
  searchSection: document.getElementById('search-section'),
  searchInput: document.getElementById('search-input'),
  searchResults: document.getElementById('search-results'),
  notesSection: document.getElementById('notes-section'),
  notesInput: document.getElementById('notes-input'),
  actionSection: document.getElementById('action-section'),
  selectedEntity: document.getElementById('selected-entity'),
  btnLogEmail: document.getElementById('btn-log-email'),
  btnSearchLeads: document.getElementById('btn-search-leads'),
  btnSearchBookings: document.getElementById('btn-search-bookings'),
  btnSearch: document.getElementById('btn-search'),
  messageContainer: document.getElementById('message-container')
};

// Initialize Office
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initializeAddin();
  }
});

function initializeAddin() {
  // Set up event listeners
  elements.btnSearchLeads.addEventListener('click', () => setSearchType('leads'));
  elements.btnSearchBookings.addEventListener('click', () => setSearchType('bookings'));
  elements.btnSearch.addEventListener('click', performSearch);
  elements.searchInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') performSearch();
  });
  elements.btnLogEmail.addEventListener('click', logEmail);

  // Load current email
  loadCurrentEmail();
}

async function loadCurrentEmail() {
  try {
    const item = Office.context.mailbox.item;
    
    if (!item) {
      showMessage('Ingen email valgt', 'error');
      return;
    }

    // Get email details
    currentEmail = {
      subject: item.subject || '(Intet emne)',
      from: '',
      to: [],
      cc: [],
      date: item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString('da-DK') : '-',
      body: '',
      messageId: item.internetMessageId || item.itemId
    };

    // Get sender
    if (item.from) {
      currentEmail.from = item.from.emailAddress || item.from.displayName || '-';
    }

    // Get recipients
    if (item.to && item.to.length > 0) {
      currentEmail.to = item.to.map(r => r.emailAddress || r.displayName);
    }

    // Get CC
    if (item.cc && item.cc.length > 0) {
      currentEmail.cc = item.cc.map(r => r.emailAddress || r.displayName);
    }

    // Get body (async)
    item.body.getAsync(Office.CoercionType.Html, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        currentEmail.body = result.value;
      }
    });

    // Get attachments
    if (item.attachments && item.attachments.length > 0) {
      attachments = [];
      for (const att of item.attachments) {
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

    // Update UI
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

  // Show attachments if any
  if (attachments.length > 0) {
    elements.attachmentsSection.classList.remove('hidden');
    elements.attachmentsList.innerHTML = attachments
      .map(att => `<li>${att.name} (${formatFileSize(att.size)})</li>`)
      .join('');
  }
}

function formatFileSize(bytes) {
  if (!bytes) return '-';
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function setSearchType(type) {
  searchType = type;
  elements.btnSearchLeads.classList.toggle('active', type === 'leads');
  elements.btnSearchBookings.classList.toggle('active', type === 'bookings');
  elements.searchResults.innerHTML = '';
  selectedEntity = null;
  updateActionSection();
}

async function performSearch() {
  const query = elements.searchInput.value.trim();
  
  if (!query) {
    elements.searchResults.innerHTML = '<div class="no-results">Indtast søgeord</div>';
    return;
  }

  elements.searchResults.innerHTML = '<div class="loading"><div class="spinner"></div></div>';

  try {
    const response = await fetch(`${CONFIG.SUPABASE_URL}/functions/v1/search-hub-entities`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${CONFIG.SUPABASE_ANON_KEY}`
      },
      body: JSON.stringify({
        query,
        type: searchType
      })
    });

    if (!response.ok) {
      throw new Error('Søgning fejlede');
    }

    const data = await response.json();
    displaySearchResults(data.results || []);

  } catch (error) {
    console.error('Search error:', error);
    elements.searchResults.innerHTML = `<div class="no-results">Fejl: ${error.message}</div>`;
  }
}

function displaySearchResults(results) {
  if (results.length === 0) {
    elements.searchResults.innerHTML = '<div class="no-results">Ingen resultater fundet</div>';
    return;
  }

  elements.searchResults.innerHTML = results.map(result => {
    const isLead = searchType === 'leads';
    const badgeClass = isLead ? 'lead' : 'booking';
    const badgeText = isLead ? 'Lead' : 'Booking';
    
    let meta = [];
    if (result.email) meta.push(result.email);
    if (result.destination) meta.push(result.destination);
    if (result.booking_number) meta.push(`#${result.booking_number}`);
    
    return `
      <div class="result-item" data-id="${result.id}" data-type="${searchType}">
        <div class="name">
          ${result.customer_name || result.name}
          <span class="badge ${badgeClass}">${badgeText}</span>
        </div>
        ${meta.length > 0 ? `<div class="meta">${meta.join(' • ')}</div>` : ''}
      </div>
    `;
  }).join('');

  // Add click handlers
  elements.searchResults.querySelectorAll('.result-item').forEach(item => {
    item.addEventListener('click', () => selectEntity(item));
  });
}

function selectEntity(element) {
  // Remove previous selection
  elements.searchResults.querySelectorAll('.result-item').forEach(el => {
    el.classList.remove('selected');
  });

  // Select new
  element.classList.add('selected');
  selectedEntity = {
    id: element.dataset.id,
    type: element.dataset.type,
    name: element.querySelector('.name').textContent.trim().replace(/Lead|Booking/g, '').trim()
  };

  updateActionSection();
}

function updateActionSection() {
  if (selectedEntity) {
    elements.actionSection.classList.remove('hidden');
    const typeLabel = selectedEntity.type === 'leads' ? 'Lead' : 'Booking';
    elements.selectedEntity.textContent = `Logger til: ${selectedEntity.name} (${typeLabel})`;
    elements.btnLogEmail.disabled = false;
  } else {
    elements.actionSection.classList.add('hidden');
    elements.btnLogEmail.disabled = true;
  }
}

async function logEmail() {
  if (!currentEmail || !selectedEntity) {
    showMessage('Vælg venligst et lead eller booking først', 'error');
    return;
  }

  elements.btnLogEmail.disabled = true;
  elements.btnLogEmail.textContent = 'Logger...';
  elements.btnLogEmail.classList.add('loading');

  try {
    // Prepare attachment data
    const attachmentData = [];
    
    // Get attachment content if any
    if (attachments.length > 0) {
      for (const att of attachments) {
        try {
          const content = await getAttachmentContent(att.id);
          if (content) {
            attachmentData.push({
              name: att.name,
              content_type: att.contentType,
              size: att.size,
              content_base64: content
            });
          }
        } catch (err) {
          console.warn('Could not get attachment content:', err);
        }
      }
    }

    const payload = {
      subject: currentEmail.subject,
      from_email: currentEmail.from,
      to_emails: currentEmail.to,
      cc_emails: currentEmail.cc,
      html_body: currentEmail.body,
      sent_at: currentEmail.date,
      message_id: currentEmail.messageId,
      lead_id: selectedEntity.type === 'leads' ? selectedEntity.id : null,
      tour_booking_id: selectedEntity.type === 'bookings' ? selectedEntity.id : null,
      notes: elements.notesInput.value.trim() || null,
      attachments: attachmentData
    };

    const response = await fetch(`${CONFIG.SUPABASE_URL}/functions/v1/log-outlook-email`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${CONFIG.SUPABASE_ANON_KEY}`
      },
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      throw new Error(errorData.error || 'Kunne ikke logge email');
    }

    showMessage('✅ Email logget succesfuldt!', 'success');
    
    // Reset form after success
    setTimeout(() => {
      elements.notesInput.value = '';
      selectedEntity = null;
      elements.searchResults.innerHTML = '';
      elements.searchInput.value = '';
      updateActionSection();
      hideMessage();
    }, 2000);

  } catch (error) {
    console.error('Log email error:', error);
    showMessage('Fejl: ' + error.message, 'error');
  } finally {
    elements.btnLogEmail.disabled = false;
    elements.btnLogEmail.textContent = 'Log Email';
    elements.btnLogEmail.classList.remove('loading');
  }
}

function getAttachmentContent(attachmentId) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value.content);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

function showMessage(text, type) {
  elements.messageContainer.textContent = text;
  elements.messageContainer.className = `message-container ${type}`;
  elements.messageContainer.classList.remove('hidden');
}

function hideMessage() {
  elements.messageContainer.classList.add('hidden');
}
