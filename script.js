// Configuration
// TIMEZONE: Asia/Kolkata (IST, GMT+5:30) - All timestamps and display use this timezone
const GET_URL = 'https://script.google.com/macros/s/AKfycbyCeT7yF3ZWv9LRCNywiwDpm_KqRPdQQs07WhBR8zON7BFsk-x4-RPFcCa28M3mJ_Ci/exec';
const POST_URL = 'https://script.google.com/macros/s/AKfycbyCeT7yF3ZWv9LRCNywiwDpm_KqRPdQQs07WhBR8zON7BFsk-x4-RPFcCa28M3mJ_Ci/exec';
const IST_TIMEZONE = 'Asia/Kolkata'; // GMT+5:30

let allBookings = [];
let currentWeekStart = {};
let currentUser = null;
let currentEventId = null;

// Helper function to get IST time
function getISTTime(date) {
    const istFormatter = new Intl.DateTimeFormat('en-US', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false,
        timeZone: IST_TIMEZONE
    });
    
    const parts = istFormatter.formatToParts(date);
    const istDate = new Date();
    
    const yearPart = parts.find(p => p.type === 'year')?.value;
    const monthPart = parts.find(p => p.type === 'month')?.value;
    const dayPart = parts.find(p => p.type === 'day')?.value;
    const hourPart = parts.find(p => p.type === 'hour')?.value;
    const minutePart = parts.find(p => p.type === 'minute')?.value;
    const secondPart = parts.find(p => p.type === 'second')?.value;
    
    istDate.setFullYear(yearPart);
    istDate.setMonth(monthPart - 1);
    istDate.setDate(dayPart);
    istDate.setHours(hourPart);
    istDate.setMinutes(minutePart);
    istDate.setSeconds(secondPart);
    
    return istDate;
}

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    initializeTabs();
    initializeViewToggles();
    initializeWeekNavigation();
    initializeCalendarClicks();
    initializeLogin();
    startClock();
    checkLoginStatus();
    initializeNoteAutosize();
    initializeParticipantsAutosize();
    initializeDateTimeValidation();
    loadBookings();
    setInterval(loadBookings, 30000);
    initializeTabsAutoScroll();
});

function openDialog(opts) {
    const existing = document.getElementById('appNotifyOverlay');
    const overlay = existing || document.createElement('div');
    overlay.id = 'appNotifyOverlay';
    overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.35);display:flex;align-items:center;justify-content:center;z-index:1000;';
    const box = document.createElement('div');
    box.style.cssText = 'max-width:440px;width:90%;border-radius:10px;box-shadow:0 10px 30px rgba(0,0,0,0.2);background:#fff;overflow:hidden;';
    const header = document.createElement('div');
    header.style.cssText = 'padding:14px 18px;font-weight:600;color:#fff;';
    const type = opts.type || 'info';
    header.style.background = type === 'success' ? '#00bcd4' : type === 'error' ? '#c62828' : type === 'warning' ? '#ed6c02' : '#1976d2';
    header.textContent = type === 'success' ? 'Success' : type === 'error' ? 'Error' : type === 'warning' ? 'Warning' : 'Notice';
    const body = document.createElement('div');
    body.style.cssText = 'padding:18px;color:#333;font-size:15px;';
    body.textContent = opts.message || '';
    const actions = document.createElement('div');
    actions.style.cssText = 'display:flex;gap:10px;justify-content:flex-end;padding:12px 18px 18px;';
    const buttons = opts.buttons && opts.buttons.length ? opts.buttons : [{ label: 'OK', value: true, variant: 'primary' }];
    return new Promise(resolve => {
        buttons.forEach(b => {
            const btn = document.createElement('button');
            btn.textContent = b.label;
            btn.style.cssText = 'padding:8px 14px;border-radius:6px;border:none;cursor:pointer;font-weight:600;';
            btn.style.background = b.variant === 'danger' ? '#c62828' : b.variant === 'secondary' ? '#e0e0e0' : '#1976d2';
            btn.style.color = b.variant === 'secondary' ? '#333' : '#fff';
            btn.addEventListener('click', () => {
                document.body.removeChild(overlay);
                resolve(b.value);
            });
            actions.appendChild(btn);
        });
        box.appendChild(header);
        box.appendChild(body);
        box.appendChild(actions);
        overlay.innerHTML = '';
        overlay.appendChild(box);
        if (!existing) document.body.appendChild(overlay);
    });
}

function showPopup(message, type) {
    return openDialog({ message, type });
}

function showConfirm(message, okLabel) {
    return openDialog({ message, type: 'warning', buttons: [
        { label: 'Cancel', value: false, variant: 'secondary' },
        { label: okLabel || 'OK', value: true, variant: 'danger' }
    ]});
}

// Authentication Functions
function initializeLogin() {
    document.getElementById('loginBtn').addEventListener('click', openLoginModal);
    document.getElementById('loginForm').addEventListener('submit', handleLogin);
}

function openLoginModal() {
    document.getElementById('loginModal').style.display = 'block';
}

function closeLoginModal() {
    document.getElementById('loginModal').style.display = 'none';
    document.getElementById('loginForm').reset();
}

async function handleLogin(e) {
    e.preventDefault();
    const form = document.getElementById('loginForm');
    const submitBtn = form ? form.querySelector('button[type="submit"]') : null;
    if (submitBtn) {
        submitBtn.disabled = true;
        submitBtn.dataset.prevText = submitBtn.textContent;
        submitBtn.textContent = 'Signing in...';
    }
    const name = document.getElementById('loginName').value;
    const accessCode = document.getElementById('loginPassword').value;
    try {
        const url = `${GET_URL}?action=auth&name=${encodeURIComponent(name)}&accessCode=${encodeURIComponent(accessCode)}`;
        const response = await fetch(url);
        if (!response.ok) {
            await showPopup('Login failed', 'error');
            return;
        }
        const data = await response.json();
        if (data && data.status === 'success' && data.user) {
            currentUser = { name: data.user.name, role: data.user.role, email: data.user.email };
            sessionStorage.setItem('currentUser', JSON.stringify(currentUser));
            updateUIForLogin();
            closeLoginModal();
            await showPopup('Login successful', 'success');
        } else {
            await showPopup('Invalid name or access code', 'error');
        }
    } catch (err) {
        await showPopup('Login error', 'error');
    } finally {
        if (submitBtn) {
            submitBtn.disabled = false;
            submitBtn.textContent = submitBtn.dataset.prevText || submitBtn.textContent;
            delete submitBtn.dataset.prevText;
        }
    }
}

function logout() {
    currentUser = null;
    sessionStorage.removeItem('currentUser');
    updateUIForLogin();
}

function checkLoginStatus() {
    const savedUser = sessionStorage.getItem('currentUser');
    if (savedUser) {
        currentUser = JSON.parse(savedUser);
        updateUIForLogin();
    }
}

function updateUIForLogin() {
    const loginBtn = document.getElementById('loginBtn');
    const userInfo = document.getElementById('userInfo');
    const userName = document.getElementById('userName');
    
    if (currentUser) {
        loginBtn.classList.add('hidden');
        userInfo.classList.remove('hidden');
        userName.textContent = `${currentUser.name} (${currentUser.role})`;
    } else {
        loginBtn.classList.remove('hidden');
        userInfo.classList.add('hidden');
    }
}

function canEditBooking(booking) {
    if (!currentUser) return false;
    if (currentUser.role === 'admin') return true;
    return booking.bookedBy === currentUser.name;
}

// Clock Functions
function startClock() {
    updateClock();
    setInterval(updateClock, 1000);
}

function updateClock() {
    const now = new Date();
    const clocks = document.querySelectorAll('[data-clock]');
    
    clocks.forEach(clock => {
        const type = clock.dataset.clock;
        const weekdayFormatter = new Intl.DateTimeFormat('en-US', { weekday: 'long', timeZone: IST_TIMEZONE });
        const dateFormatter = new Intl.DateTimeFormat('en-US', { month: 'short', day: 'numeric', year: 'numeric', timeZone: IST_TIMEZONE });
        const timeFormatter = new Intl.DateTimeFormat('en-US', { hour: '2-digit', minute: '2-digit', second: '2-digit', timeZone: IST_TIMEZONE });
        
        const weekday = weekdayFormatter.format(now);
        const date = dateFormatter.format(now);
        const time = timeFormatter.format(now);
        
        if (type === 'global') {
            clock.textContent = `${weekday}, ${date} • ${time}`;
        } else {
            clock.textContent = `${time} • ${date}`;
        }
    });
    
    updateCurrentTimeLine();
}

// Booking Functions
async function loadBookings() {
    try {
        const response = await fetch(GET_URL);
        if (!response.ok) throw new Error('Failed to fetch');
        
        const data = await response.json();
        allBookings = Array.isArray(data) ? data.map(normalizeBooking) : [];
        displaySummaryViews();
        refreshAllCalendars();
    } catch (error) {
        console.error('Error loading bookings:', error);
        allBookings = [];
    }
}

function normalizeBooking(booking) {
    return {
        id: booking.id || Date.now(),
        room: booking.room || 'Room A',
        roomKey: booking.roomKey || (booking.room ? String(booking.room).replace('Room ', '') : ''),
        title: booking.title || 'Untitled Meeting',
        start: booking.start,
        end: booking.end,
        bookedBy: booking.bookedBy || 'Unknown',
        note: booking.note || '',
        participants: booking.participants || '',
        emailSent: booking.emailSent || false,
        createdBy: booking.createdBy || '',
        updatedBy: booking.updatedBy || '',
        createdAt: booking.createdAt || '',
        updatedAt: booking.updatedAt || ''
    };
}

// Tab Functions
function initializeTabs() {
    const buttons = document.querySelectorAll('.tab-button');
    buttons.forEach(button => {
        button.addEventListener('click', () => {
            const tab = button.dataset.tab;
            document.querySelectorAll('.tab-button').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            button.classList.add('active');
            document.getElementById(`tab-${tab}`).classList.add('active');
        });
    });
}

// View Toggle Functions
function initializeViewToggles() {
    const toggleButtons = document.querySelectorAll('.view-toggle-btn');
    toggleButtons.forEach(button => {
        button.addEventListener('click', () => {
            const view = button.dataset.view;
            const room = button.dataset.room;
            
            const container = button.closest('.room-panel');
            container.querySelectorAll('.view-toggle-btn').forEach(b => b.classList.remove('active'));
            button.classList.add('active');
            
            const summaryView = container.querySelector('.summary-view');
            const calendarView = container.querySelector('.calendar-view');
            
            if (view === 'calendar') {
                summaryView.classList.remove('active');
                calendarView.classList.add('active');
                if (!currentWeekStart[room]) {
                    currentWeekStart[room] = getWeekStart(new Date());
                }
                renderCalendar(room);
            } else {
                calendarView.classList.remove('active');
                summaryView.classList.add('active');
            }
        });
    });
}

// Week Navigation Functions
function initializeWeekNavigation() {
    document.querySelectorAll('.week-nav-btn').forEach(button => {
        button.addEventListener('click', () => {
            const action = button.dataset.action;
            const room = button.dataset.room;
            
            if (!currentWeekStart[room]) {
                currentWeekStart[room] = getWeekStart(new Date());
            }
            
            if (action === 'prev-week') {
                currentWeekStart[room] = addDays(currentWeekStart[room], -1);
            } else if (action === 'next-week') {
                currentWeekStart[room] = addDays(currentWeekStart[room], 1);
            } else if (action === 'today') {
                currentWeekStart[room] = getWeekStart(new Date());
            }
            
            renderCalendar(room);
        });
    });
}

function initializeCalendarClicks() {
    document.addEventListener('click', (e) => {
        const hourCell = e.target.closest('.hour-cell');
        if (hourCell && !e.target.closest('.calendar-booking')) {
            const room = hourCell.dataset.room;
            const date = hourCell.dataset.date;
            const hour = hourCell.dataset.hour;
            if (room && date && hour) {
                // Get current IST time
                const now = new Date();
                const istFormatter = new Intl.DateTimeFormat('en-US', {
                    year: 'numeric',
                    month: '2-digit',
                    day: '2-digit',
                    hour: '2-digit',
                    minute: '2-digit',
                    hour12: false,
                    timeZone: IST_TIMEZONE
                });
                
                const timeParts = istFormatter.formatToParts(now);
                const currentISTYear = parseInt(timeParts.find(p => p.type === 'year')?.value);
                const currentISTMonth = parseInt(timeParts.find(p => p.type === 'month')?.value);
                const currentISTDay = parseInt(timeParts.find(p => p.type === 'day')?.value);
                const currentISTHour = parseInt(timeParts.find(p => p.type === 'hour')?.value);
                const currentISTMinute = parseInt(timeParts.find(p => p.type === 'minute')?.value);
                
                // Parse selected slot date
                const parts = String(date).split('-').map(Number);
                const slotYear = parts[0];
                const slotMonth = parts[1] || 1;
                const slotDay = parts[2] || 1;
                const slotHour = Number(hour);
                
                // Check if selected slot is in the past (with 5-minute margin)
                const isSlotToday = (slotYear === currentISTYear && slotMonth === currentISTMonth && slotDay === currentISTDay);
                const isSlotHourInPast = isSlotToday && slotHour < currentISTHour;
                const isPastDate = (slotYear < currentISTYear) || 
                                  (slotYear === currentISTYear && slotMonth < currentISTMonth) ||
                                  (slotYear === currentISTYear && slotMonth === currentISTMonth && slotDay < currentISTDay);
                
                if (isPastDate || isSlotHourInPast) {
                    showPopup('Selected time is in the past', 'warning');
                    return;
                }
                
                // Allow booking with current IST time if slot is in current hour
                openModalWithTime(room, date, hour, isSlotToday && slotHour === currentISTHour ? currentISTMinute : null);
            }
        }
    });
}

function getWeekStart(date) {
    const d = new Date(date);
    d.setHours(0, 0, 0, 0);
    return d;
}

function addDays(date, days) {
    const result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
}

function formatWeekRange(weekStart) {
    const dayBefore = addDays(weekStart, -1);
    const dayAfter5 = addDays(weekStart, 5);
    const startStr = dayBefore.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
    const endStr = dayAfter5.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
    return `${startStr} - ${endStr}`;
}

// Calendar Rendering
function renderCalendar(room) {
    if (!currentWeekStart[room]) {
        currentWeekStart[room] = getWeekStart(new Date());
    }
    
    const weekStart = currentWeekStart[room];
    const container = document.querySelector(`.calendar-grid[data-room="${room}"]`);
    const weekRangeEl = document.querySelector(`.week-range[data-room="${room}"]`);
    
    if (weekRangeEl) {
        weekRangeEl.textContent = formatWeekRange(weekStart);
    }
    
    if (!container) return;
    
    let html = '<div class="calendar-header">';
    html += '<div class="calendar-header-cell">Time</div>';
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const formatLocalDate = (d) => {
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const da = String(d.getDate()).padStart(2, '0');
        return `${y}-${m}-${da}`;
    };
    for (let i = -1; i < 6; i++) {
        const date = addDays(weekStart, i);
        const dateStr = formatLocalDate(date);
        const dayName = date.toLocaleDateString('en-US', { weekday: 'short' });
        const dayNum = date.getDate();
        const isToday = date.getTime() === today.getTime();
        html += `<div class="calendar-header-cell${isToday ? ' today' : ''}" data-date="${dateStr}">${dayName}<br>${dayNum}</div>`;
    }
    html += '</div>';
    
    html += '<div class="calendar-body">';
    html += '<div class="calendar-days">';
    
    let timeLabelsColumn = '';
    for (let hour = 8; hour <= 19; hour++) {
        timeLabelsColumn += `<div class="time-label" data-hour="${hour}">${formatHour(hour)}</div>`;
    }
    html += `<div class="time-labels-column">${timeLabelsColumn}</div>`;
    
    for (let day = -1; day < 6; day++) {
        const date = addDays(weekStart, day);
        const dateStr = formatLocalDate(date);
        const isToday = date.getTime() === today.getTime();
        
        let dayColumnHtml = `<div class="day-column${isToday ? ' today' : ''}" data-room="${room}" data-date="${dateStr}">`;
        for (let hour = 8; hour <= 19; hour++) {
            dayColumnHtml += `<div class="hour-cell" data-room="${room}" data-date="${dateStr}" data-hour="${hour}"></div>`;
        }
        
        const roomLabel = `Room ${room}`;
        const visibleStart = new Date(date);
        visibleStart.setHours(8, 0, 0, 0);
        const visibleEnd = new Date(date);
        visibleEnd.setHours(20, 0, 0, 0);
        
        const dayBookings = allBookings.filter(b => {
            if (b.room !== roomLabel) return false;
            const start = new Date(b.start);
            const end = new Date(b.end);
            const dayStart = new Date(date);
            dayStart.setHours(0, 0, 0, 0);
            const dayEnd = new Date(date);
            dayEnd.setHours(23, 59, 59, 999);
            return start < dayEnd && end > dayStart;
        });
        
        dayBookings.forEach(booking => {
            const bookingStart = new Date(booking.start);
            const bookingEnd = new Date(booking.end);
            const clampedStart = bookingStart > visibleStart ? bookingStart : visibleStart;
            const clampedEnd = bookingEnd < visibleEnd ? bookingEnd : visibleEnd;
            if (clampedEnd <= clampedStart) return;
            const minutesFromVisibleStart = (clampedStart - visibleStart) / (1000 * 60);
            const durationMinutes = (clampedEnd - clampedStart) / (1000 * 60);
            const topPx = minutesFromVisibleStart;
            const heightPx = durationMinutes;
            const now = new Date();
            const isRunning = bookingStart <= now && bookingEnd > now;
            const startTime = bookingStart.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' });
            const endTime = bookingEnd.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' });
            dayColumnHtml += `
                <div class="calendar-booking${isRunning ? ' running' : ''}" style="top: ${topPx}px; height: ${heightPx}px;" onclick="viewBookingDetails('${booking.id}')">
                    <div class="calendar-booking-title">${escapeHtml(booking.title)}</div>
                    <div class="calendar-booking-time">${startTime} - ${endTime}</div>
                </div>
            `;
        });
        
        dayColumnHtml += '</div>';
        html += dayColumnHtml;
    }
    
    html += '</div>';
    html += '</div>';
    container.innerHTML = html;

    initializeCalendarHover();
}

let lastHoverDate = null;
let lastHoverHour = null;

function initializeCalendarHover() {
    document.addEventListener('mouseover', (e) => {
        const cell = e.target.closest('.hour-cell');
        if (!cell) return;
        const date = cell.dataset.date;
        const hour = cell.dataset.hour;
        if (!date || !hour) return;

        if (lastHoverDate === date && lastHoverHour === hour) return;
        clearHoverHighlights();
        applyHoverHighlights(date, hour);
        lastHoverDate = date;
        lastHoverHour = hour;
    });

    document.addEventListener('mouseout', (e) => {
        const cell = e.target.closest('.hour-cell');
        if (!cell) return;
        clearHoverHighlights();
        lastHoverDate = null;
        lastHoverHour = null;
    });
}

function applyHoverHighlights(date, hour) {
    const headerCell = document.querySelector(`.calendar-header-cell[data-date="${date}"]`);
    if (headerCell) headerCell.classList.add('hover');
    const timeLabel = document.querySelector(`.time-labels-column .time-label[data-hour="${hour}"]`);
    if (timeLabel) timeLabel.classList.add('hover');
    document.querySelectorAll(`.hour-cell[data-hour="${hour}"]`).forEach(el => el.classList.add('hover-row'));
}

function clearHoverHighlights() {
    document.querySelectorAll('.calendar-header-cell.hover').forEach(el => el.classList.remove('hover'));
    document.querySelectorAll('.time-label.hover').forEach(el => el.classList.remove('hover'));
    document.querySelectorAll('.hour-cell.hover-row').forEach(el => el.classList.remove('hover-row'));
}

function formatHour(hour) {
    if (hour === 0) return '12 AM';
    if (hour < 12) return `${hour} AM`;
    if (hour === 12) return '12 PM';
    return `${hour - 12} PM`;
}

function refreshAllCalendars() {
    ['A', 'B', 'C'].forEach(room => {
        const calendarView = document.querySelector(`.calendar-view[data-room="${room}"]`);
        if (calendarView && calendarView.classList.contains('active')) {
            renderCalendar(room);
        }
    });
    updateCurrentTimeLine();
}

function updateCurrentTimeLine() {
    const now = new Date();
    
    // Get current IST time
    const istFormatter = new Intl.DateTimeFormat('en-US', {
        hour: '2-digit',
        minute: '2-digit',
        hour12: false,
        timeZone: IST_TIMEZONE
    });
    
    const timeParts = istFormatter.formatToParts(now);
    const currentHour = parseInt(timeParts.find(p => p.type === 'hour').value);
    const currentMinutes = parseInt(timeParts.find(p => p.type === 'minute').value);
    
    // Only show line if within calendar hours (8-19)
    if (currentHour < 8 || currentHour >= 20) {
        document.querySelectorAll('.current-time-line').forEach(line => line.remove());
        return;
    }
    
    // Calculate position: each hour is 60px, each minute is 1px
    const offsetFromTop = (currentHour - 8) * 60 + currentMinutes;
    
    document.querySelectorAll('.day-column').forEach(column => {
        // Remove existing line
        const existingLine = column.querySelector('.current-time-line');
        if (existingLine) existingLine.remove();
        
        // Add new line
        const line = document.createElement('div');
        line.className = 'current-time-line';
        line.style.top = offsetFromTop + 'px';
        column.appendChild(line);
    });
}

// Summary View
function displaySummaryViews() {
    const now = new Date();
    const upcomingBookings = allBookings.filter(b => new Date(b.end) > now);
    
    displayRoomSummary('A', upcomingBookings.filter(b => b.room === 'Room A'));
    displayRoomSummary('B', upcomingBookings.filter(b => b.room === 'Room B'));
    displayRoomSummary('C', upcomingBookings.filter(b => b.room === 'Room C'));
    displayAllBookings(upcomingBookings);
}

function displayRoomSummary(room, bookings) {
    const container = document.getElementById(`room${room}-bookings`);
    if (!container) return;
    
    if (bookings.length === 0) {
        container.innerHTML = '<p class="no-bookings">No upcoming bookings</p>';
        return;
    }
    
    bookings.sort((a, b) => new Date(a.start) - new Date(b.start));
    container.innerHTML = bookings.map(b => createBookingCard(b, false)).join('');
}

function displayAllBookings(bookings) {
    const container = document.getElementById('all-bookings');
    if (!container) return;
    
    if (bookings.length === 0) {
        container.innerHTML = '<p class="no-bookings">No upcoming bookings</p>';
        return;
    }
    
    bookings.sort((a, b) => new Date(a.start) - new Date(b.start));
    container.innerHTML = bookings.map(b => createBookingCard(b, true)).join('');
}

function createBookingCard(booking, showRoom) {
    const now = new Date();
    const start = new Date(booking.start);
    const end = new Date(booking.end);
    
    let status = 'upcoming';
    if (now >= start && now < end) status = 'running';
    else if (now >= end) status = 'past';
    
    const statusLabel = status === 'running' ? 'Running' : status === 'upcoming' ? 'Upcoming' : 'Finished';
    
    const roomLabel = getRoomDisplayName(booking);
    return `
        <div class="booking-card booking-status-${status}" onclick="viewBookingDetails('${booking.id}')" style="cursor: pointer;">
            <div class="booking-card-header">
                <h3>${escapeHtml(booking.title)}</h3>
                <span class="status-badge status-badge-${status}">${statusLabel}</span>
            </div>
            <div class="booking-info">
                ${showRoom ? `<span><strong>Room:</strong> ${escapeHtml(roomLabel)}</span>` : ''}
                <span><strong>Booked by:</strong> ${escapeHtml(booking.bookedBy)}</span>
                <span><strong>Start:</strong> ${formatDateTime(booking.start)}</span>
                <span><strong>End:</strong> ${formatDateTime(booking.end)}</span>
                ${booking.participants ? `<span><strong>Participants:</strong> ${escapeHtml(booking.participants)}</span>` : ''}
            </div>
            ${booking.note ? `<div class="booking-note"><strong>Note:</strong> ${escapeHtml(booking.note)}</div>` : ''}
        </div>
    `;
}

function formatDateTime(dateString) {
    const date = new Date(dateString);
    // Format in IST (Asia/Kolkata, GMT+5:30)
    const formatter = new Intl.DateTimeFormat('en-US', {
        year: 'numeric',
        month: 'short',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        timeZone: 'Asia/Kolkata'
    });
    return formatter.format(date);
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// Event Details Modal
function viewBookingDetails(bookingId) {
    const booking = allBookings.find(b => String(b.id) === String(bookingId));
    if (!booking) return;
    
    currentEventId = bookingId;
    const canEdit = canEditBooking(booking);
    
    const now = new Date();
    const start = new Date(booking.start);
    const end = new Date(booking.end);
    
    let status = 'Upcoming';
    if (now >= start && now < end) status = 'Running';
    else if (now >= end) status = 'Finished';
    
    let detailsHtml = `
        <div class="event-details-content">
            <div class="event-detail-row">
                <strong>Meeting Title</strong>
                <span>${escapeHtml(booking.title)}</span>
            </div>
            <div class="event-detail-row">
                <strong>Room</strong>
                <span>${escapeHtml(getRoomDisplayName(booking))}</span>
            </div>
            <div class="event-detail-row">
                <strong>Booked By</strong>
                <span>${escapeHtml(booking.bookedBy)}</span>
            </div>
            <div class="event-detail-row">
                <strong>Status</strong>
                <span>${status}</span>
            </div>
            <div class="event-detail-row">
                <strong>Start Time</strong>
                <span>${formatDateTime(booking.start)}</span>
            </div>
            <div class="event-detail-row">
                <strong>End Time</strong>
                <span>${formatDateTime(booking.end)}</span>
            </div>
            ${booking.participants ? `
            <div class="event-detail-row">
                <strong>Participants</strong>
                <span>${escapeHtml(booking.participants)}</span>
            </div>
            ` : ''}
            ${booking.note ? `
            <div class="event-detail-row">
                <strong>Note</strong>
                <span>${escapeHtml(booking.note)}</span>
            </div>
            ` : ''}
            ${booking.emailSent ? `
            <div class="event-detail-row">
                <strong>Email Status</strong>
                <span style="color: #00bcd4;">✓ Invitations sent</span>
            </div>
            ` : ''}
        </div>
    `;
    
    if (canEdit) {
        detailsHtml += `
            <div class="event-actions">
                <button class="edit-btn" onclick="editBooking('${bookingId}')">Edit Meeting</button>
                <button class="delete-btn" onclick="deleteBooking('${bookingId}')">Delete Meeting</button>
            </div>
        `;
    }
    
    document.getElementById('eventDetails').innerHTML = detailsHtml;
    document.getElementById('eventModal').style.display = 'block';
}

function closeEventModal() {
    document.getElementById('eventModal').style.display = 'none';
    currentEventId = null;
}

// Edit and Delete Functions
function editBooking(bookingId) {
    const booking = allBookings.find(b => String(b.id) === String(bookingId));
    if (!booking || !canEditBooking(booking)) {
        showPopup('You do not have permission to edit this booking', 'error');
        return;
    }
    
    closeEventModal();
    
    const roomLetter = booking.room.replace('Room ', '');
    const modal = document.getElementById('bookingModal');
    const form = document.getElementById('bookingForm');
    
    form.dataset.editMode = 'true';
    form.dataset.bookingId = bookingId;
    
    document.getElementById('room').value = booking.room;
    document.getElementById('title').value = booking.title;
    document.getElementById('startTime').value = toDateTimeLocal(new Date(booking.start));
    document.getElementById('endTime').value = toDateTimeLocal(new Date(booking.end));
    document.getElementById('note').value = booking.note || '';
    document.getElementById('participants').value = booking.participants || '';
    
    modal.querySelector('h2').textContent = 'Edit Booking';
    modal.style.display = 'block';
    autosizeNote();
    autosizeParticipants();
}

async function deleteBooking(bookingId) {
    const booking = allBookings.find(b => String(b.id) === String(bookingId));
    if (!booking || !canEditBooking(booking)) {
        showPopup('You do not have permission to delete this booking', 'error');
        return;
    }
    
    const confirmed = await showConfirm('Are you sure you want to delete this booking? Cancellation emails will be sent to all participants.', 'Delete');
    if (!confirmed) {
        return;
    }
    
    try {
        await fetch(POST_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                action: 'delete',
                id: bookingId,
                sendCancellation: true,
                deletedBy: currentUser ? currentUser.name : 'Unknown'
            })
        });
        
        await showPopup('Booking deleted successfully. Cancellation emails are being sent.', 'success');
        closeEventModal();
        setTimeout(loadBookings, 1500);
    } catch (error) {
        console.error('Error:', error);
        await showPopup('Failed to delete booking. Please try again.', 'error');
    }
}

// Modal Functions
function openModal(room) {
    if (!currentUser) {
        showPopup('Please login to create a booking', 'info');
        openLoginModal();
        return;
    }
    
    const modal = document.getElementById('bookingModal');
    const form = document.getElementById('bookingForm');
    form.reset();
    delete form.dataset.editMode;
    delete form.dataset.bookingId;
    
    document.getElementById('room').value = `Room ${room}`;
    document.getElementById('sendEmailNotification').checked = true;
    modal.querySelector('h2').textContent = 'Add New Booking';
    modal.style.display = 'block';
    autosizeNote();
    autosizeParticipants();
    
    // Get current IST time with 5-minute margin
    const now = new Date();
    const istFormatter = new Intl.DateTimeFormat('en-US', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false,
        timeZone: IST_TIMEZONE
    });
    
    const timeParts = istFormatter.formatToParts(now);
    const currentISTYear = parseInt(timeParts.find(p => p.type === 'year')?.value);
    const currentISTMonth = parseInt(timeParts.find(p => p.type === 'month')?.value);
    const currentISTDay = parseInt(timeParts.find(p => p.type === 'day')?.value);
    const currentISTHour = parseInt(timeParts.find(p => p.type === 'hour')?.value);
    const currentISTMinute = parseInt(timeParts.find(p => p.type === 'minute')?.value);
    
    let currentIST = new Date(currentISTYear, currentISTMonth - 1, currentISTDay, currentISTHour, currentISTMinute, 0, 0);
    currentIST.setMinutes(currentIST.getMinutes() - 5);
    
    const nowLocal = toDateTimeLocalIST(currentIST);
    const startInput = document.getElementById('startTime');
    const endInput = document.getElementById('endTime');
    if (startInput) startInput.min = nowLocal;
    if (endInput) endInput.min = nowLocal;
    validateDateTimeInputs();
}

function openModalWithTime(room, dateStr, hour, startMinute = null) {
    if (!currentUser) {
        showPopup('Please login to create a booking', 'info');
        openLoginModal();
        return;
    }
    
    const modal = document.getElementById('bookingModal');
    const form = document.getElementById('bookingForm');
    form.reset();
    delete form.dataset.editMode;
    delete form.dataset.bookingId;
    
    document.getElementById('room').value = `Room ${room}`;
    
    const parts = String(dateStr).split('-').map(Number);
    // If startMinute is provided, use current IST minutes for start time
    const startMinutes = startMinute !== null ? startMinute : 0;
    const startDate = new Date(parts[0], (parts[1] || 1) - 1, parts[2] || 1, Number(hour), startMinutes, 0, 0);
    
    // End time is always 1 hour after start time
    const endDate = new Date(startDate);
    endDate.setHours(endDate.getHours() + 1);
    
    document.getElementById('startTime').value = toDateTimeLocal(startDate);
    document.getElementById('endTime').value = toDateTimeLocal(endDate);
    document.getElementById('sendEmailNotification').checked = true;
    
    modal.querySelector('h2').textContent = 'Add New Booking';
    modal.style.display = 'block';
    autosizeNote();
    autosizeParticipants();
    
    // Get current IST time with 5-minute margin
    const now = new Date();
    const istFormatter = new Intl.DateTimeFormat('en-US', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false,
        timeZone: IST_TIMEZONE
    });
    
    const timeParts = istFormatter.formatToParts(now);
    const currentISTYear = parseInt(timeParts.find(p => p.type === 'year')?.value);
    const currentISTMonth = parseInt(timeParts.find(p => p.type === 'month')?.value);
    const currentISTDay = parseInt(timeParts.find(p => p.type === 'day')?.value);
    const currentISTHour = parseInt(timeParts.find(p => p.type === 'hour')?.value);
    const currentISTMinute = parseInt(timeParts.find(p => p.type === 'minute')?.value);
    
    let currentIST = new Date(currentISTYear, currentISTMonth - 1, currentISTDay, currentISTHour, currentISTMinute, 0, 0);
    currentIST.setMinutes(currentIST.getMinutes() - 5);
    
    const nowLocal = toDateTimeLocalIST(currentIST);
    const startInput = document.getElementById('startTime');
    const endInput = document.getElementById('endTime');
    if (startInput) startInput.min = nowLocal;
    if (endInput) endInput.min = startInput.value || nowLocal;
    validateDateTimeInputs();
}

function toDateTimeLocal(date) {
    // Convert to IST format for datetime-local input
    const istFormatter = new Intl.DateTimeFormat('en-US', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false,
        timeZone: IST_TIMEZONE
    });
    
    const parts = istFormatter.formatToParts(date);
    const year = parts.find(p => p.type === 'year')?.value;
    const month = parts.find(p => p.type === 'month')?.value;
    const day = parts.find(p => p.type === 'day')?.value;
    const hours = parts.find(p => p.type === 'hour')?.value;
    const minutes = parts.find(p => p.type === 'minute')?.value;
    
    return `${year}-${month}-${day}T${hours}:${minutes}`;
}

function toDateTimeLocalIST(date) {
    // Convert to IST format for datetime-local input (uses IST as reference timezone)
    const istFormatter = new Intl.DateTimeFormat('en-US', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false,
        timeZone: IST_TIMEZONE
    });
    
    const parts = istFormatter.formatToParts(date);
    const year = parts.find(p => p.type === 'year')?.value;
    const month = parts.find(p => p.type === 'month')?.value;
    const day = parts.find(p => p.type === 'day')?.value;
    const hours = parts.find(p => p.type === 'hour')?.value;
    const minutes = parts.find(p => p.type === 'minute')?.value;
    
    return `${year}-${month}-${day}T${hours}:${minutes}`;
}

function closeModal() {
    const form = document.getElementById('bookingForm');
    document.getElementById('bookingModal').style.display = 'none';
    form.reset();
    delete form.dataset.editMode;
    delete form.dataset.bookingId;
}

window.onclick = async function(event) {
    const bookingModal = document.getElementById('bookingModal');
    const eventModal = document.getElementById('eventModal');
    const loginModal = document.getElementById('loginModal');
    
    if (event.target === bookingModal) {
        const confirmed = await showConfirm('Do you want to discard settings?', 'Discard');
        if (confirmed) {
            closeModal();
        }
    } else if (event.target === eventModal) {
        closeEventModal();
    } else if (event.target === loginModal) {
        closeLoginModal();
    }
}

// Form Submission
document.getElementById('bookingForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    
    if (!currentUser) {
        showPopup('Please login to create or edit a booking', 'info');
        return;
    }
    
    const form = e.target;
    const isEditMode = form.dataset.editMode === 'true';
    const bookingId = form.dataset.bookingId;
    
    const formData = new FormData(form);
    const room = formData.get('room');
    const title = formData.get('title');
    const startTime = formData.get('startTime');
    const endTime = formData.get('endTime');
    const note = formData.get('note');
    const participants = formData.get('participants');
    const sendEmail = document.getElementById('sendEmailNotification').checked;
    
    const start = new Date(startTime);
    const end = new Date(endTime);
    const now = new Date();
    if (start < now) {
        showPopup('Start time cannot be in the past', 'warning');
        return;
    }
    
    if (end <= start) {
        showPopup('End time must be after start time', 'warning');
        return;
    }
    
    // Check for conflicts (exclude current booking if editing)
    const hasConflict = allBookings.some(booking => {
        if (isEditMode && String(booking.id) === String(bookingId)) return false;
        if (booking.room !== room) return false;
        const bookingStart = new Date(booking.start);
        const bookingEnd = new Date(booking.end);
        return (start < bookingEnd && end > bookingStart);
    });
    
    if (hasConflict) {
        showPopup('This time slot conflicts with an existing booking', 'warning');
        return;
    }
    
    // Extract room key from room name (e.g., "Room A" -> "A")
    const roomKey = room.replace('Room ', '').trim();
    
    const booking = {
        action: isEditMode ? 'update' : 'create',
        room,
        roomKey,
        title,
        start: start.toISOString(),
        end: end.toISOString(),
        bookedBy: currentUser.name,
        createdBy: currentUser.name,
        note: note || '',
        participants: participants || '',
        sendEmail: sendEmail
    };
    
    if (isEditMode) {
        booking.id = bookingId;
        booking.updatedBy = currentUser.name;
    }
    
    const submitBtn = form.querySelector('button[type="submit"]');
    if (submitBtn) {
        submitBtn.disabled = true;
        submitBtn.dataset.prevText = submitBtn.textContent;
        submitBtn.textContent = isEditMode ? 'Updating...' : 'Submitting...';
    }
    
    try {
        await fetch(POST_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(booking)
        });
        
        const successMsg = isEditMode 
            ? 'Booking updated successfully' + (sendEmail && participants ? '. Email notifications are being sent.' : '')
            : 'Booking submitted successfully' + (sendEmail && participants ? '. Email invitations are being sent.' : '');
        
        await showPopup(successMsg, 'success');
        closeModal();
        setTimeout(loadBookings, 1500);
    } catch (error) {
        console.error('Error:', error);
        await showPopup(`Failed to ${isEditMode ? 'update' : 'submit'} booking. Please try again.`, 'error');
    } finally {
        if (submitBtn) {
            submitBtn.disabled = false;
            submitBtn.textContent = submitBtn.dataset.prevText || submitBtn.textContent;
            delete submitBtn.dataset.prevText;
        }
    }
});

function validateDateTimeInputs() {
    const form = document.getElementById('bookingForm');
    const startInput = document.getElementById('startTime');
    const endInput = document.getElementById('endTime');
    const errorEl = document.getElementById('dateTimeError');
    const submitBtn = form ? form.querySelector('button[type="submit"]') : null;
    if (!startInput || !endInput) return;
    
    // Get current IST time with 5-minute margin
    const now = new Date();
    const istFormatter = new Intl.DateTimeFormat('en-US', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false,
        timeZone: IST_TIMEZONE
    });
    
    const timeParts = istFormatter.formatToParts(now);
    const currentISTYear = parseInt(timeParts.find(p => p.type === 'year')?.value);
    const currentISTMonth = parseInt(timeParts.find(p => p.type === 'month')?.value);
    const currentISTDay = parseInt(timeParts.find(p => p.type === 'day')?.value);
    const currentISTHour = parseInt(timeParts.find(p => p.type === 'hour')?.value);
    const currentISTMinute = parseInt(timeParts.find(p => p.type === 'minute')?.value);
    
    // Create current IST time and subtract 5-minute margin (allowing bookings made up to 5 minutes before current time)
    let currentIST = new Date(currentISTYear, currentISTMonth - 1, currentISTDay, currentISTHour, currentISTMinute, 0, 0);
    currentIST.setMinutes(currentIST.getMinutes() - 5);
    
    const startVal = startInput.value;
    const endVal = endInput.value;
    let valid = true;
    let errorMsg = 'Selected date/time is in the past';
    
    if (startVal) {
        const s = new Date(startVal);
        if (s < currentIST) valid = false;
    }
    
    if (endVal && startVal) {
        const s = new Date(startVal);
        const e = new Date(endVal);
        if (e <= s) {
            valid = false;
            errorMsg = 'End time must be after start time';
        }
        if (e < currentIST) {
            valid = false;
            errorMsg = 'Selected date/time is in the past';
        }
    }
    
    if (endInput) endInput.min = startVal || toDateTimeLocalIST(currentIST);
    if (errorEl) {
        errorEl.style.display = valid ? 'none' : 'block';
        errorEl.textContent = errorMsg;
    }
    if (submitBtn) submitBtn.disabled = !valid;
}

function initializeDateTimeValidation() {
    const startInput = document.getElementById('startTime');
    const endInput = document.getElementById('endTime');
    if (!startInput || !endInput) return;
    
    // Get current IST time with 5-minute margin
    const now = new Date();
    const istFormatter = new Intl.DateTimeFormat('en-US', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false,
        timeZone: IST_TIMEZONE
    });
    
    const timeParts = istFormatter.formatToParts(now);
    const currentISTYear = parseInt(timeParts.find(p => p.type === 'year')?.value);
    const currentISTMonth = parseInt(timeParts.find(p => p.type === 'month')?.value);
    const currentISTDay = parseInt(timeParts.find(p => p.type === 'day')?.value);
    const currentISTHour = parseInt(timeParts.find(p => p.type === 'hour')?.value);
    const currentISTMinute = parseInt(timeParts.find(p => p.type === 'minute')?.value);
    
    let currentIST = new Date(currentISTYear, currentISTMonth - 1, currentISTDay, currentISTHour, currentISTMinute, 0, 0);
    currentIST.setMinutes(currentIST.getMinutes() - 5);
    
    const nowLocal = toDateTimeLocalIST(currentIST);
    startInput.min = nowLocal;
    endInput.min = nowLocal;
    const handler = () => validateDateTimeInputs();
    startInput.addEventListener('input', handler);
    endInput.addEventListener('input', handler);
    validateDateTimeInputs();
}

function initializeNoteAutosize() {
    const note = document.getElementById('note');
    if (!note) return;
    note.style.overflowY = 'hidden';
    note.style.resize = 'none';
    const handler = () => {
        note.style.height = 'auto';
        note.style.height = note.scrollHeight + 'px';
    };
    note.addEventListener('input', handler);
    note.addEventListener('change', handler);
    handler();
}

function autosizeNote() {
    const note = document.getElementById('note');
    if (!note) return;
    note.style.height = 'auto';
    note.style.height = note.scrollHeight + 'px';
}

function initializeParticipantsAutosize() {
    const participants = document.getElementById('participants');
    if (!participants) return;
    participants.style.overflowY = 'hidden';
    participants.style.resize = 'none';
    const handler = () => {
        participants.style.height = 'auto';
        participants.style.height = participants.scrollHeight + 'px';
    };
    participants.addEventListener('input', handler);
    participants.addEventListener('change', handler);
    handler();
}

function autosizeParticipants() {
    const participants = document.getElementById('participants');
    if (!participants) return;
    participants.style.height = 'auto';
    participants.style.height = participants.scrollHeight + 'px';
}

function getRoomDisplayName(booking) {
    const key = booking.roomKey || (booking.room && booking.room.replace('Room ', '')) || '';
    if (key === 'A') return 'A- BLOCK A BOARDROOM';
    if (key === 'B') return 'B- BLOCK C BOARDROOM';
    if (key === 'C') return 'C-BLOCK D AUDITORIUM';
    if (booking.room) return booking.room;
    return 'Unknown Room';

}



