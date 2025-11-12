// Configuration - Replace with your Google Apps Script Web App URLs
const GET_URL = 'https://script.google.com/macros/s/AKfycbyjazSKGZNiQJ3s0KWapU7wO5cPg8gafsjE0BQZI73i1E3OK_LikkmQ14JZcc2AUOF4/exec';
const POST_URL = 'https://script.google.com/macros/s/AKfycbyjazSKGZNiQJ3s0KWapU7wO5cPg8gafsjE0BQZI73i1E3OK_LikkmQ14JZcc2AUOF4/exec';
const AUTH_STORAGE_KEY = 'authorized';

let currentRoom = '';
let allBookings = [];
let pendingRoom = '';
let displayIntervalId = null;
const renderCache = new Map();
const ADMIN_USERS = ['admin', 'Admin', 'Administrator', 'Manager'];
let editingBookingId = null;
let editingBookingOriginal = null;
let pendingDiscardAction = null;

// Load bookings when page loads
document.addEventListener('DOMContentLoaded', () => {
    initializeTabs();
    initializeBookingCardActions();
    initializeDiscardModal();
    startRealTimeUpdates();
    loadBookings();
    // Refresh bookings every 30 seconds
    setInterval(loadBookings, 30000);
});

function startRealTimeUpdates() {
    if (displayIntervalId) {
        return;
    }

    updateTimeAndDisplays();
    displayIntervalId = setInterval(updateTimeAndDisplays, 1000);
}

function updateTimeAndDisplays() {
    const now = new Date();
    updateClockDisplays(now);
    displayBookings(now);
}

function updateClockDisplays(now) {
    const clockElements = document.querySelectorAll('[data-clock]');

    clockElements.forEach(element => {
        const target = element.dataset.clock || 'global';
        element.textContent = formatClockForTarget(now, target);
    });
}

function formatClockForTarget(now, target) {
    const timeOptions = { hour: '2-digit', minute: '2-digit', second: '2-digit' };
    const dateOptions = { month: 'short', day: 'numeric', year: 'numeric' };

    const time = now.toLocaleTimeString('en-US', timeOptions);
    const date = now.toLocaleDateString('en-US', dateOptions);

    if (target === 'global') {
        const weekday = now.toLocaleDateString('en-US', { weekday: 'long' });
        return `${weekday}, ${date} • ${time}`;
    }

    return `${time} • ${date}`;
}

// Load all bookings from Google Sheets
async function loadBookings() {
    try {
        const response = await fetch(GET_URL);
        if (!response.ok) {
            throw new Error('Failed to fetch bookings');
        }
        
        const data = await response.json();
        allBookings = Array.isArray(data) ? data.map(normalizeBooking) : [];
        displayBookings(new Date());
    } catch (error) {
        console.error('Error loading bookings:', error);
        showError('Failed to load bookings. Please check your configuration.');
        allBookings = [];
        updateNowRunningBox([], new Date());
    }
}

function normalizeBooking(rawBooking = {}) {
    const safeBooking = { ...rawBooking };

    const title = typeof safeBooking.title === 'string' ? safeBooking.title.trim() : '';
    const bookedBy = typeof safeBooking.bookedBy === 'string' ? safeBooking.bookedBy.trim() : '';
    const room = typeof safeBooking.room === 'string' ? safeBooking.room.trim() : '';
    const roomKey = extractRoomKey(room || safeBooking.roomKey || '');
    const roomLabel = formatRoomLabel(roomKey);
    const noteValue = pickFirstString([
        safeBooking.note,
        safeBooking.notes,
        safeBooking.Note,
        safeBooking.Notes,
        safeBooking.noteText,
        safeBooking.description
    ]);

    safeBooking.title = title || 'Untitled Meeting';
    safeBooking.bookedBy = bookedBy || 'Unknown';
    safeBooking.room = roomLabel;
    safeBooking.roomKey = roomKey;
    safeBooking.start = safeBooking.start;
    safeBooking.end = safeBooking.end;
    safeBooking.note = noteValue;
    safeBooking.id = deriveBookingId(safeBooking);

    return safeBooking;
}

function pickFirstString(candidates = []) {
    for (const value of candidates) {
        if (typeof value === 'string') {
            const trimmed = value.trim();
            if (trimmed) {
                return trimmed;
            }
        }
    }
    return '';
}

// Display bookings in respective panels
function displayBookings(now = new Date()) {
    if (!Array.isArray(allBookings) || allBookings.length === 0) {
        updateNowRunningBox([], now);
        return;
    }

    const sortedBookings = [...allBookings].sort((a, b) => new Date(a.start) - new Date(b.start));
    
    // Filter upcoming bookings for room panels
    const upcomingBookings = sortedBookings.filter(booking => {
        const endTime = new Date(booking.end);
        return endTime > now;
    });

    // Sort by start time
    upcomingBookings.sort((a, b) => new Date(a.start) - new Date(b.start));

    // Display for each room (upcoming and running only)
    displayRoomBookings('A', upcomingBookings.filter(b => b.room === 'Room A'), now);
    displayRoomBookings('B', upcomingBookings.filter(b => b.room === 'Room B'), now);
    displayRoomBookings('C', upcomingBookings.filter(b => b.room === 'Room C'), now);
    
    // Display all bookings
    displayAllBookings(sortedBookings, now);
}

// Display bookings for a specific room
function displayRoomBookings(room, bookings, now) {
    const container = document.getElementById(`room${room}-bookings`);
    if (!container) {
        return;
    }
    
    if (bookings.length === 0) {
        renderIfChanged(container, '<p class="no-bookings">No upcoming bookings</p>');
        return;
    }

    const markup = bookings.map(booking => createBookingCard(booking, { showRoom: false, now })).join('');
    renderIfChanged(container, markup);
}

// Display all bookings
function displayAllBookings(bookings, now) {
    const container = document.getElementById('all-bookings');
    if (!container) {
        return;
    }
    
    if (bookings.length === 0) {
        renderIfChanged(container, '<p class="no-bookings">No upcoming bookings</p>');
        updateNowRunningBox([], now);
        return;
    }

    const markup = bookings.map(booking => createBookingCard(booking, { showRoom: true, now })).join('');
    renderIfChanged(container, markup);

    const runningBookings = bookings.filter(booking => getBookingStatus(booking, now) === 'running');
    updateNowRunningBox(runningBookings, now);
}

// Tabs
function initializeTabs() {
    const buttons = document.querySelectorAll('.tab-button');
    const contents = document.querySelectorAll('.tab-content');

    if (!buttons.length || !contents.length) {
        return;
    }

    buttons.forEach(button => {
        button.addEventListener('click', () => {
            setActiveTab(button.dataset.tab);
        });
    });

    const activeButton = Array.from(buttons).find(btn => btn.classList.contains('active'));
    if (activeButton) {
        setActiveTab(activeButton.dataset.tab);
    } else {
        setActiveTab(buttons[0].dataset.tab);
    }
}

function setActiveTab(tabKey) {
    const buttons = document.querySelectorAll('.tab-button');
    const contents = document.querySelectorAll('.tab-content');

    buttons.forEach(button => {
        button.classList.toggle('active', button.dataset.tab === tabKey);
    });

    contents.forEach(content => {
        content.classList.toggle('active', content.id === `tab-${tabKey}`);
    });
}

function initializeBookingCardActions() {
    document.addEventListener('click', handleBookingCardAction);
}

function initializeDiscardModal() {
    const discardModal = document.getElementById('discardModal');
    if (!discardModal) {
        return;
    }

    const cancelBtn = discardModal.querySelector('[data-discard-cancel]');
    const confirmBtn = discardModal.querySelector('[data-discard-confirm]');

    if (cancelBtn) {
        cancelBtn.addEventListener('click', () => {
            closeDiscardModal(false);
        });
    }

    if (confirmBtn) {
        confirmBtn.addEventListener('click', () => {
            closeDiscardModal(true);
        });
    }
}

function handleBookingCardAction(event) {
    const editButton = event.target.closest('[data-action="edit-booking"]');
    if (!editButton) {
        return;
    }

    event.preventDefault();

    const bookingId = editButton.dataset.bookingId;
    const booking = findBookingById(bookingId);

    if (!booking) {
        console.warn('Unable to locate booking for edit', bookingId);
        return;
    }

    if (!isAuthorized()) {
        const roomKey = extractRoomKey(booking.room);
        pendingRoom = roomKey;
        openAuthModal(roomKey);
        return;
    }

    if (!canEditBooking(booking)) {
        alert('You are not permitted to edit this booking.');
        return;
    }

    openEditModal(booking);
}

function openEditModal(booking) {
    if (!booking) {
        return;
    }

    editingBookingId = booking.id;
    editingBookingOriginal = { ...booking };

    const roomKey = extractRoomKey(booking.room);
    showBookingModal(roomKey, { booking, mode: 'edit' });
}

// Create HTML for a booking card
function createBookingCard(booking, options = {}) {
    const { showRoom = false, now = new Date() } = options;
    const status = getBookingStatus(booking, now);
    const statusLabel = getBookingStatusLabel(status);
    const startTime = formatDateTime(booking.start);
    const endTime = formatDateTime(booking.end);
    const noteText = typeof booking.note === 'string' ? booking.note.trim() : '';
    const safeTitle = escapeHtml(booking.title);
    const safeBookedBy = escapeHtml(booking.bookedBy);
    const safeRoom = escapeHtml(booking.room);
    const safeNote = noteText ? escapeHtml(noteText) : '';
    const noteMarkup = `<div class="booking-note"><strong>Note:</strong> ${safeNote || '<span class="muted">No note provided.</span>'}</div>`;
    const canEdit = canEditBooking(booking);
    const editButtonMarkup = canEdit
        ? `<button type="button" class="edit-booking-btn" data-action="edit-booking" data-booking-id="${escapeHtml(String(booking.id))}" title="Edit meeting">Edit</button>`
        : '';

    return `
        <div class="booking-card booking-status-${status}" data-status="${status}" data-booking-id="${escapeHtml(String(booking.id))}">
            <div class="booking-card-header">
                <h3>${safeTitle}</h3>
                <div class="booking-card-header-actions">
                    <span class="status-badge status-badge-${status}">${statusLabel}</span>
                    ${editButtonMarkup}
                </div>
            </div>
            <div class="booking-info">
                ${showRoom ? `<span><strong>Room:</strong> ${safeRoom}</span>` : ''}
                <span><strong>Booked by:</strong> ${safeBookedBy}</span>
                <span><strong>Start:</strong> ${startTime}</span>
                <span><strong>End:</strong> ${endTime}</span>
            </div>
            ${noteMarkup}
        </div>
    `;
}

function getBookingStatus(booking, now) {
    const start = new Date(booking.start);
    const end = new Date(booking.end);

    if (now >= start && now < end) {
        return 'running';
    }

    if (now < start) {
        return 'upcoming';
    }

    return 'past';
}

function getBookingStatusLabel(status) {
    switch (status) {
        case 'running':
            return 'Running';
        case 'upcoming':
            return 'Upcoming';
        case 'past':
        default:
            return 'Finished';
    }
}

function updateNowRunningBox(runningBookings, now) {
    const box = document.getElementById('nowRunningBox');
    const content = document.getElementById('nowRunningContent');

    if (!box || !content) {
        return;
    }

    if (!runningBookings.length) {
        content.innerHTML = '';
        box.classList.add('hidden');
        return;
    }

    box.classList.remove('hidden');
    const sortedRunning = [...runningBookings].sort((a, b) => new Date(a.start) - new Date(b.start));

    const markup = sortedRunning
        .map(booking => createRunningNoteMarkup(booking))
        .join('');

    renderIfChanged(content, markup);
}

function createRunningNoteMarkup(booking) {
    const noteText = typeof booking.note === 'string' ? booking.note.trim() : '';
    const safeTitle = escapeHtml(booking.title);
    const safeRoom = escapeHtml(booking.room);
    const safeBookedBy = escapeHtml(booking.bookedBy);
    const safeNote = noteText ? escapeHtml(noteText) : '';
    const timeRange = formatTimeRange(booking);
    const noteMarkup = safeNote
        ? `<p class="now-running-note"><strong>Note:</strong> ${safeNote}</p>`
        : '<p class="now-running-note muted">No note provided.</p>';

    return `
        <div class="now-running-item">
            <div class="now-running-meta">
                <span class="now-running-room">${safeRoom}</span>
                <span class="now-running-time">${timeRange}</span>
            </div>
            <h4>${safeTitle}</h4>
            ${noteMarkup}
            <span class="now-running-host">Hosted by ${safeBookedBy}</span>
        </div>
    `;
}

function formatTimeRange(booking) {
    const options = { hour: '2-digit', minute: '2-digit' };
    const start = new Date(booking.start).toLocaleTimeString('en-US', options);
    const end = new Date(booking.end).toLocaleTimeString('en-US', options);
    return `${start} - ${end}`;
}

// Format date and time
function formatDateTime(dateString) {
    const date = new Date(dateString);
    const options = {
        year: 'numeric',
        month: 'short',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
    };
    return date.toLocaleString('en-US', options);
}

function escapeHtml(value) {
    if (value === null || value === undefined) {
        return '';
    }

    return String(value).replace(/[&<>"']/g, (match) => {
        switch (match) {
            case '&':
                return '&amp;';
            case '<':
                return '&lt;';
            case '>':
                return '&gt;';
            case '"':
                return '&quot;';
            case '\'':
                return '&#39;';
            default:
                return match;
        }
    });
}

function renderIfChanged(element, markup) {
    if (!element) {
        return;
    }

    const previousMarkup = renderCache.get(element);
    if (previousMarkup === markup) {
        return;
    }

    element.innerHTML = markup;
    renderCache.set(element, markup);
}

function findBookingById(id) {
    if (!id) {
        return null;
    }

    return allBookings.find(booking => String(booking.id) === String(id)) || null;
}

function getAuthorizedUser() {
    return sessionStorage.getItem('authorizedUser') || '';
}

function normalizeName(value) {
    return (value || '').toString().trim().toLowerCase();
}

function isAdminUser(name = getAuthorizedUser()) {
    const normalized = normalizeName(name);
    return ADMIN_USERS.some(adminName => normalizeName(adminName) === normalized);
}

function canEditBooking(booking) {
    if (!booking || !isAuthorized()) {
        return false;
    }

    const user = getAuthorizedUser();
    if (!user) {
        return false;
    }

    if (isAdminUser(user)) {
        return true;
    }

    return normalizeName(booking.bookedBy) === normalizeName(user);
}

function formatRoomLabel(roomKey = '') {
    const trimmed = (roomKey || '').toString().trim();
    if (!trimmed) {
        return '';
    }

    if (trimmed.toLowerCase().startsWith('room ')) {
        return `Room ${trimmed.slice(5).trim()}`;
    }

    return trimmed.startsWith('Room ') ? trimmed : `Room ${trimmed}`;
}

function extractRoomKey(roomLabel = '') {
    const value = (roomLabel || '').toString().trim();
    if (!value) {
        return '';
    }

    return value.toLowerCase().startsWith('room ')
        ? value.slice(5).trim()
        : value;
}

function deriveBookingId(booking) {
    if (!booking || typeof booking !== 'object') {
        return '';
    }

    const candidateKeys = [
        'id',
        'ID',
        'bookingId',
        'bookingID',
        'recordId',
        'recordID',
        'rowId',
        'rowID',
        'entryId',
        'entryID',
        'uuid',
        'UID',
        'guid',
        'timestamp',
        'createdAt'
    ];

    for (const key of candidateKeys) {
        if (key in booking) {
            const candidate = booking[key];
            if (candidate !== undefined && candidate !== null && String(candidate).trim() !== '') {
                return String(candidate);
            }
        }
    }

    const fallback = [
        booking.room || booking.roomKey || '',
        booking.start || '',
        booking.end || '',
        booking.title || '',
        booking.bookedBy || ''
    ]
        .map(part => String(part || '').trim().toLowerCase())
        .join('|');

    return fallback || `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

function toDatetimeLocalValue(value) {
    if (!value) {
        return '';
    }

    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
        return '';
    }

    const year = date.getFullYear();
    const month = padNumber(date.getMonth() + 1);
    const day = padNumber(date.getDate());
    const hours = padNumber(date.getHours());
    const minutes = padNumber(date.getMinutes());

    return `${year}-${month}-${day}T${hours}:${minutes}`;
}

function padNumber(value) {
    return String(value).padStart(2, '0');
}

function getBookingFormSnapshot() {
    const form = document.getElementById('bookingForm');
    if (!form) {
        return {};
    }

    const getValue = (id) => {
        const element = document.getElementById(id);
        return element ? element.value : '';
    };

    return {
        mode: form.dataset.mode || 'create',
        room: getValue('room'),
        title: getValue('title'),
        start: getValue('startTime'),
        end: getValue('endTime'),
        bookedBy: getValue('bookedBy'),
        note: getValue('note'),
        accessCode: getValue('accessCode')
    };
}

function recordBookingFormInitialState() {
    const form = document.getElementById('bookingForm');
    if (!form) {
        return;
    }

    const snapshot = JSON.stringify(getBookingFormSnapshot());
    form.dataset.initialState = snapshot;
}

function hasBookingFormChanges() {
    const form = document.getElementById('bookingForm');
    if (!form) {
        return false;
    }

    const initialState = form.dataset.initialState;
    if (!initialState) {
        return true;
    }

    try {
        const currentState = JSON.stringify(getBookingFormSnapshot());
        return currentState !== initialState;
    } catch (error) {
        console.warn('Unable to compare booking form state:', error);
        return true;
    }
}

function isBookingModalOpen() {
    const bookingModal = document.getElementById('bookingModal');
    return Boolean(bookingModal && bookingModal.style.display === 'block');
}

function attemptCloseBookingModal() {
    if (!isBookingModalOpen()) {
        return;
    }

    const hasChanges = hasBookingFormChanges();
    const message = hasChanges
        ? 'Discard your changes to this booking?'
        : 'Close the booking form without saving?';

    openDiscardModal({
        message,
        cancelLabel: 'Keep Editing',
        confirmLabel: hasChanges ? 'Discard Changes' : 'Close Form',
        confirmTone: hasChanges ? 'destructive' : 'accent',
        onConfirm: () => {
            closeModal();
        }
    });
}

function openDiscardModal(options = {}) {
    const discardModal = document.getElementById('discardModal');
    if (!discardModal) {
        if (typeof options.onConfirm === 'function') {
            options.onConfirm();
        }
        return;
    }

    const messageEl = discardModal.querySelector('[data-discard-message]');
    if (messageEl) {
        messageEl.textContent = options.message || 'Discard changes?';
    }

    const cancelBtn = discardModal.querySelector('[data-discard-cancel]');
    const confirmBtn = discardModal.querySelector('[data-discard-confirm]');

    if (cancelBtn) {
        if (!cancelBtn.dataset.defaultLabel) {
            cancelBtn.dataset.defaultLabel = cancelBtn.textContent.trim();
        }
        if (!cancelBtn.dataset.defaultClass) {
            cancelBtn.dataset.defaultClass = cancelBtn.className;
        }
        cancelBtn.textContent = options.cancelLabel || cancelBtn.dataset.defaultLabel || 'Keep Editing';
        cancelBtn.className = cancelBtn.dataset.defaultClass || cancelBtn.className;
    }

    if (confirmBtn) {
        if (!confirmBtn.dataset.defaultLabel) {
            confirmBtn.dataset.defaultLabel = confirmBtn.textContent.trim();
        }
        if (!confirmBtn.dataset.defaultClass) {
            confirmBtn.dataset.defaultClass = confirmBtn.className;
        }
        confirmBtn.textContent = options.confirmLabel || confirmBtn.dataset.defaultLabel || 'Discard';

        const tone = options.confirmTone || 'destructive';
        if (tone === 'accent') {
            confirmBtn.className = 'submit-btn accent';
        } else {
            confirmBtn.className = 'submit-btn destructive';
        }
    }

    discardModal.classList.remove('hidden');
    discardModal.style.display = 'block';
    pendingDiscardAction = typeof options.onConfirm === 'function' ? options.onConfirm : null;

    const focusTarget = discardModal.querySelector('[data-discard-confirm]');
    if (focusTarget) {
        requestAnimationFrame(() => {
            focusTarget.focus();
        });
    }
}

function closeDiscardModal(shouldConfirm = false) {
    const discardModal = document.getElementById('discardModal');
    if (!discardModal) {
        return;
    }

    discardModal.style.display = 'none';
    discardModal.classList.add('hidden');

    const cancelBtn = discardModal.querySelector('[data-discard-cancel]');
    const confirmBtn = discardModal.querySelector('[data-discard-confirm]');

    if (cancelBtn) {
        if (cancelBtn.dataset.defaultLabel) {
            cancelBtn.textContent = cancelBtn.dataset.defaultLabel;
        }
        if (cancelBtn.dataset.defaultClass) {
            cancelBtn.className = cancelBtn.dataset.defaultClass;
        }
    }

    if (confirmBtn) {
        if (confirmBtn.dataset.defaultLabel) {
            confirmBtn.textContent = confirmBtn.dataset.defaultLabel;
        }
        if (confirmBtn.dataset.defaultClass) {
            confirmBtn.className = confirmBtn.dataset.defaultClass;
        }
    }

    const action = pendingDiscardAction;
    pendingDiscardAction = null;

    if (shouldConfirm && typeof action === 'function') {
        action();
    }
}

function attemptDeleteBooking() {
    const bookingForm = document.getElementById('bookingForm');
    if (!bookingForm || bookingForm.dataset.mode !== 'edit') {
        return;
    }

    const bookingId = bookingForm.dataset.bookingId;
    const booking = editingBookingOriginal || findBookingById(bookingId);

    if (!booking) {
        alert('Unable to locate this booking. Please refresh and try again.');
        return;
    }

    if (!canEditBooking(booking)) {
        alert('You are not permitted to delete this booking.');
        return;
    }

    openDiscardModal({
        message: 'Delete this meeting? This action cannot be undone.',
        cancelLabel: 'Keep Meeting',
        confirmLabel: 'Delete Meeting',
        confirmTone: 'destructive',
        onConfirm: () => performDeleteBooking(booking)
    });
}

async function performDeleteBooking(booking) {
    const sessionUser = getAuthorizedUser();
    const payload = {
        action: 'delete',
        id: booking.id,
        room: booking.room,
        roomKey: booking.roomKey || extractRoomKey(booking.room),
        title: booking.title,
        start: booking.start,
        end: booking.end,
        bookedBy: booking.bookedBy,
        note: booking.note || '',
        deletedBy: sessionUser || booking.bookedBy || ''
    };

    try {
        await fetch(POST_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(payload)
        });

        alert('Meeting deleted successfully.');
        closeModal();
        setTimeout(() => {
            loadBookings();
        }, 1200);
    } catch (error) {
        console.error('Error deleting booking:', error);
        alert('Failed to delete meeting. Please try again.');
    }
}

// Authorization helpers
function isAuthorized() {
    return sessionStorage.getItem(AUTH_STORAGE_KEY) === 'true';
}

function openAuthModal(room) {
    pendingRoom = room;
    const authModal = document.getElementById('authModal');
    const authForm = document.getElementById('authForm');
    const authError = document.getElementById('authError');

    authForm.reset();
    authError.textContent = '';
    authModal.style.display = 'block';

    requestAnimationFrame(() => {
        const nameField = document.getElementById('authName');
        if (nameField) {
            nameField.focus();
        }
    });
}

function closeAuthModal() {
    const authModal = document.getElementById('authModal');
    document.getElementById('authForm').reset();
    document.getElementById('authError').textContent = '';
    authModal.style.display = 'none';
    pendingRoom = '';
}

// Open modal for adding booking
function openModal(room) {
    if (!isAuthorized()) {
        openAuthModal(room);
        return;
    }

    showBookingModal(room);
}

function showBookingModal(room, options = {}) {
    const bookingModal = document.getElementById('bookingModal');
    const bookingForm = document.getElementById('bookingForm');
    if (!bookingModal || !bookingForm) {
        return;
    }

    const booking = options.booking || null;
    const mode = options.mode || (booking ? 'edit' : 'create');
    const roomKey = extractRoomKey(booking ? booking.room : room);
    const roomLabel = formatRoomLabel(roomKey);

    currentRoom = roomKey;
    pendingRoom = '';

    bookingForm.reset();
    bookingForm.dataset.mode = mode;
    bookingForm.dataset.bookingId = booking && booking.id ? String(booking.id) : '';

    const modalTitle = bookingModal.querySelector('.modal-header h2');
    if (modalTitle) {
        modalTitle.textContent = mode === 'edit' ? 'Edit Booking' : 'Add New Booking';
    }

    const roomField = document.getElementById('room');
    const titleField = document.getElementById('title');
    const startField = document.getElementById('startTime');
    const endField = document.getElementById('endTime');
    const bookedByField = document.getElementById('bookedBy');
    const accessCodeField = document.getElementById('accessCode');
    const noteField = document.getElementById('note');
    const submitButton = document.getElementById('submitBookingBtn');
    const deleteButton = document.getElementById('deleteBookingBtn');

    roomField.value = roomLabel;

    if (mode === 'edit' && booking) {
        titleField.value = booking.title || '';
        startField.value = toDatetimeLocalValue(booking.start);
        endField.value = toDatetimeLocalValue(booking.end);
        startField.min = '';
        endField.min = '';
        bookedByField.value = booking.bookedBy || '';
        noteField.value = booking.note || '';
        accessCodeField.required = false;
        accessCodeField.placeholder = 'Access code (required only if changing host)';
    } else {
        const now = new Date();
        const dateTimeString = now.toISOString().slice(0, 16);
        startField.min = dateTimeString;
        endField.min = dateTimeString;
        bookedByField.value = getAuthorizedUser() || '';
        noteField.value = '';
        accessCodeField.required = true;
        accessCodeField.placeholder = 'Enter access code';
    }

    accessCodeField.value = '';
    if (submitButton) {
        submitButton.textContent = mode === 'edit' ? 'Save Changes' : 'Submit Booking';
    }
    if (deleteButton) {
        deleteButton.classList.toggle('hidden', mode !== 'edit');
    }
    bookingModal.style.display = 'block';
    pendingDiscardAction = null;
    recordBookingFormInitialState();

    requestAnimationFrame(() => {
        if (titleField) {
            titleField.focus();
        }
    });
}

// Close modal
function closeModal() {
    const bookingModal = document.getElementById('bookingModal');
    const bookingForm = document.getElementById('bookingForm');
    const accessCodeField = document.getElementById('accessCode');

    closeDiscardModal(false);

    if (bookingModal) {
        bookingModal.style.display = 'none';
    }

    if (bookingForm) {
        bookingForm.reset();
        bookingForm.dataset.mode = 'create';
        bookingForm.dataset.bookingId = '';
        delete bookingForm.dataset.initialState;
    }

    if (accessCodeField) {
        accessCodeField.required = true;
        accessCodeField.placeholder = 'Enter access code';
    }

    editingBookingId = null;
    editingBookingOriginal = null;
    currentRoom = '';
    pendingRoom = '';
    pendingDiscardAction = null;
}

// Close modals when clicking outside
window.onclick = function(event) {
    const modal = document.getElementById('bookingModal');
    const authModal = document.getElementById('authModal');
    const discardModal = document.getElementById('discardModal');

    if (event.target === modal) {
        attemptCloseBookingModal();
    }

    if (event.target === authModal) {
        closeAuthModal();
    }

    if (event.target === discardModal) {
        closeDiscardModal(false);
    }
}

// Handle authorization form
document.getElementById('authForm').addEventListener('submit', (e) => {
    e.preventDefault();

    const name = document.getElementById('authName').value.trim();
    const accessCode = document.getElementById('authCode').value;
    const authError = document.getElementById('authError');

    if (validUsers[name] && validUsers[name] === accessCode) {
        sessionStorage.setItem(AUTH_STORAGE_KEY, 'true');
        sessionStorage.setItem('authorizedUser', name);
        const roomToOpen = pendingRoom || currentRoom || 'A';

        closeAuthModal();
        displayBookings(new Date());
        showBookingModal(roomToOpen);
        return;
    }

    authError.textContent = 'Access denied.';
    document.getElementById('authCode').value = '';
    requestAnimationFrame(() => {
        document.getElementById('authCode').focus();
    });
});

// Handle form submission
document.getElementById('bookingForm').addEventListener('submit', async (e) => {
    e.preventDefault();

    const bookingForm = e.target;
    const mode = bookingForm.dataset.mode || 'create';
    const editingId = bookingForm.dataset.bookingId || '';
    const formData = new FormData(bookingForm);

    const roomInput = (formData.get('room') || '').toString().trim();
    const roomKey = extractRoomKey(roomInput);
    const roomLabel = formatRoomLabel(roomKey);
    const title = (formData.get('title') || '').trim();
    const name = (formData.get('bookedBy') || '').trim();
    const accessCode = (formData.get('accessCode') || '').trim();
    const note = (formData.get('note') || '').trim();
    const startValue = formData.get('startTime');
    const endValue = formData.get('endTime');

    if (!title) {
        alert('Meeting title is required.');
        return;
    }

    if (!roomLabel) {
        alert('Room selection is required.');
        return;
    }

    const startTime = new Date(startValue);
    const endTime = new Date(endValue);

    if (Number.isNaN(startTime.getTime()) || Number.isNaN(endTime.getTime())) {
        alert('Please provide valid start and end times.');
        return;
    }

    if (endTime <= startTime) {
        alert('End time must be after start time');
        return;
    }

    const sessionUser = getAuthorizedUser();
    const isAdmin = isAdminUser(sessionUser);
    const targetBooking = mode === 'edit' ? findBookingById(editingId) : null;
    const originalHost = targetBooking ? targetBooking.bookedBy : '';
    const isOwner = targetBooking ? normalizeName(targetBooking.bookedBy) === normalizeName(sessionUser) : false;

    if (mode === 'create') {
        if (!validUsers[name] || validUsers[name] !== accessCode) {
            alert('Access denied: Invalid name or access code');
            return;
        }
    } else {
        if (!isAuthorized() || (!isOwner && !isAdmin)) {
            alert('You are not authorized to edit this booking.');
            return;
        }

        const hostChanged = targetBooking && normalizeName(name) !== normalizeName(originalHost);
        if (hostChanged) {
            if (!validUsers[name] || validUsers[name] !== accessCode) {
                alert('Please provide the access code for the updated host.');
                return;
            }
        } else if (accessCode && (!validUsers[originalHost] || validUsers[originalHost] !== accessCode)) {
            alert('Access code does not match the current host.');
            return;
        }
    }

    const hasConflict = allBookings.some(booking => {
        if (booking.room !== roomLabel) return false;
        if (mode === 'edit' && booking.id === editingId) return false;

        const bookingStart = new Date(booking.start);
        const bookingEnd = new Date(booking.end);

        return (
            (startTime >= bookingStart && startTime < bookingEnd) ||
            (endTime > bookingStart && endTime <= bookingEnd) ||
            (startTime <= bookingStart && endTime >= bookingEnd)
        );
    });

    if (hasConflict) {
        alert('This time slot conflicts with an existing booking. Please choose a different time.');
        return;
    }

    const bookingPayload = {
        action: mode === 'edit' ? 'update' : 'create',
        room: roomLabel,
        roomKey,
        title,
        start: startTime.toISOString(),
        end: endTime.toISOString(),
        bookedBy: name,
        note,
        submittedBy: sessionUser || name
    };

    if (mode === 'edit') {
        bookingPayload.id = editingId;
        bookingPayload.originalId = editingBookingOriginal?.id || editingId;
        bookingPayload.originalRoom = editingBookingOriginal?.room || roomLabel;
        bookingPayload.originalStart = editingBookingOriginal?.start || '';
        bookingPayload.originalEnd = editingBookingOriginal?.end || '';
        bookingPayload.originalBookedBy = editingBookingOriginal?.bookedBy || '';
        bookingPayload.originalTitle = editingBookingOriginal?.title || '';
        bookingPayload.updatedBy = sessionUser || name;
    } else {
        bookingPayload.createdBy = name;
    }

    try {
        await fetch(POST_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(bookingPayload)
        });

        alert(mode === 'edit' ? 'Booking updated successfully!' : 'Booking submitted successfully!');
        closeModal();

        setTimeout(() => {
            loadBookings();
        }, 1500);
    } catch (error) {
        console.error('Error submitting booking:', error);
        alert(mode === 'edit' ? 'Failed to update booking. Please try again.' : 'Failed to submit booking. Please try again.');
    }
});

// Show error message
function showError(message) {
    const containers = [
        'roomA-bookings',
        'roomB-bookings',
        'roomC-bookings',
        'all-bookings'
    ];
    
    containers.forEach(id => {
        const container = document.getElementById(id);
        renderIfChanged(container, `<p class="no-bookings" style="color: #ff6b6b;">${message}</p>`);
    });

    updateNowRunningBox([], new Date());
}
