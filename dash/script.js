// Configuration
const API_URL = 'https://script.google.com/macros/s/AKfycbyhDgoCXuDTgVtbYtI6DvrSCiFyYkY-6QN4bZElSefnJf2ZfGo60HQsxnzhcFDhvmE4/exec';
const REFRESH_INTERVAL = 60000; // 1 minute
const CAROUSEL_AUTO_ADVANCE = 10000; // 10 seconds

let currentRoomKey = '';
let cameraStream = null;
let refreshTimer = null;
let carouselTimer = null;
let currentCarouselIndex = 0;
let totalMeetings = 0;

// Timezone Configuration
const IST_TIMEZONE = 'Asia/Kolkata'; // +5:30 GMT (Used in Sri Lanka)

// Room name mapping
const ROOM_NAMES = {
    'A': 'BLOCK A BOARDROOM',
    'B': 'BLOCK C BOARDROOM',
    'C': 'BLOCK D AUDITORIUM'
};

// ===========================
// Initialization
// ===========================
document.addEventListener('DOMContentLoaded', () => {
    initializeApp();
});

function initializeApp() {
    // Get room key from URL parameter
    const urlParams = new URLSearchParams(window.location.search);
    currentRoomKey = urlParams.get('room') || 'A';
    
    // Set room name
    document.getElementById('roomName').textContent = ROOM_NAMES[currentRoomKey] || `Room ${currentRoomKey}`;
    
    // Start clock
    updateClock();
    setInterval(updateClock, 1000);

    // Set initial loading state
    setLoadingState();
    
    // Initial load
    loadRoomSchedule();
    
    // Set up auto-refresh
    refreshTimer = setInterval(loadRoomSchedule, REFRESH_INTERVAL);
    
    // Attach swipe gestures
    attachSwipeGestures();
}

// ===========================
// Clock Functions
// ===========================
function updateClock() {
    const now = new Date();
    const timeString = formatDateIST(now, 'time');
    const dateString = formatDateIST(now, 'date');
    
    document.getElementById('currentTime').textContent = `${dateString} • ${timeString}`;
}

// ===========================
// Data Loading
// ===========================
async function loadRoomSchedule() {
    try {
        const url = `${API_URL}?action=getRoomSchedule&room=${currentRoomKey}`;
        const response = await fetch(url);
        
        if (!response.ok) {
            throw new Error('Failed to fetch schedule');
        }
        
        const data = await response.json();
        
        // Update displays
        updateCurrentMeeting(data.currentMeeting);
        updateUpcomingMeetings(data.upcomingMeetings);
        
        // Update last update time
        const now = new Date();
        document.getElementById('lastUpdate').textContent = formatDateIST(now, 'time-only');
        
    } catch (error) {
        console.error('Error loading schedule:', error);
        displayError();
    }
}

// ===========================
// Current Meeting Display
// ===========================
function updateCurrentMeeting(meeting) {
    const displayContainer = document.getElementById('currentMeetingDisplay');
    const statusIndicator = document.getElementById('statusIndicator');

    // Clear loading state
    statusIndicator.classList.remove('loading');

    if (meeting) {
        // Normalize meeting data
        const m = normalizeMeeting(meeting);

        // Update status
        statusIndicator.classList.add('occupied');
        statusIndicator.querySelector('.status-text').textContent = 'In Use';

        // Enable camera
        enableCamera();

        // Calculate time remaining
        const now = new Date();
        const end = m.endDate;
        const remainingMs = end - now;
        const remainingMins = Math.max(0, Math.floor(remainingMs / 60000));
        const remainingHours = Math.floor(remainingMins / 60);
        const remainingMinsDisplay = remainingMins % 60;

        let timeRemainingText = '';
        if (remainingHours > 0) {
            timeRemainingText = `${remainingHours}h ${remainingMinsDisplay}m`;
        } else {
            timeRemainingText = `${remainingMinsDisplay} minutes`;
        }

        // Calculate progress
        const start = m.startDate;
        const totalDuration = end - start;
        const elapsed = now - start;
        const progressPercent = totalDuration > 0 ? Math.min(100, (elapsed / totalDuration) * 100) : 0;

        // Render meeting display
        displayContainer.innerHTML = `
            <div class="meeting-active">
                <div class="meeting-title">${escapeHtml(m.title)}</div>
                <div class="meeting-details">
                    <div class="meeting-detail-row">
                        <span class="detail-icon">👤</span>
                        <span class="detail-label">Organizer:</span>
                        <span class="detail-value">${escapeHtml(m.bookedBy)}</span>
                    </div>
                    <div class="meeting-detail-row">
                        <span class="detail-icon">🕐</span>
                        <span class="detail-label">Started:</span>
                        <span class="detail-value">${formatTime(m.startDate)}</span>
                    </div>
                    <div class="meeting-detail-row">
                        <span class="detail-icon">🕐</span>
                        <span class="detail-label">Ends:</span>
                        <span class="detail-value">${formatTime(m.endDate)}</span>
                    </div>
                    ${m.participants ? `
                    <div class="meeting-detail-row">
                        <span class="detail-icon">👥</span>
                        <span class="detail-label">Participants:</span>
                        <span class="detail-value">${escapeHtml(m.participants)}</span>
                    </div>
                    ` : ''}
                </div>
                <div class="meeting-time-progress">
                    <div class="time-remaining">
                        <span class="time-remaining-label">Time Remaining</span>
                        <span class="time-remaining-value">${timeRemainingText}</span>
                    </div>
                    <div class="progress-bar">
                        <div class="progress-fill" style="width: ${progressPercent}%"></div>
                    </div>
                </div>
            </div>
        `;
    } else {
        // No meeting - room available
        statusIndicator.classList.remove('occupied');
        statusIndicator.querySelector('.status-text').textContent = 'Available';

        // Disable camera
        disableCamera();

        displayContainer.innerHTML = `
            <div class="no-meeting">
                <div class="no-meeting-icon">✓</div>
                <div class="no-meeting-text">Room Available</div>
                <div class="no-meeting-subtext">No ongoing meeting</div>
            </div>
        `;
    }
}

// ===========================
// Upcoming Meetings Display
// ===========================
function updateUpcomingMeetings(meetings) {
    const upcomingList = document.getElementById('upcomingList');
    const counterDiv = document.getElementById('carouselCounter');

    if (!meetings || meetings.length === 0) {
        upcomingList.innerHTML = '<div class="no-upcoming">No more meetings scheduled for today</div>';
        counterDiv.style.display = 'none';
        stopCarousel();
        return;
    }

    totalMeetings = meetings.length;
    currentCarouselIndex = 0;

    let html = '';
    meetings.forEach((meeting, index) => {
        const m = normalizeMeeting(meeting);
        html += `
            <div class="upcoming-item ${index === 0 ? 'active' : ''}">
                <div class="upcoming-header">
                    <div class="upcoming-title">${escapeHtml(m.title)}</div>
                    <div class="upcoming-time">${formatTimeRange(m.startDate, m.endDate)}</div>
                </div>
                <div class="upcoming-details">
                    <div class="upcoming-detail">
                        <strong>Organizer:</strong>
                        <span>${escapeHtml(m.bookedBy)}</span>
                    </div>
                    ${m.participants ? `
                    <div class="upcoming-detail">
                        <strong>Participants:</strong>
                        <span>${escapeHtml(m.participants)}</span>
                    </div>
                    ` : ''}
                    ${m.note ? `
                    <div class="upcoming-detail">
                        <strong>Note:</strong>
                        <span>${escapeHtml(m.note)}</span>
                    </div>
                    ` : ''}
                </div>
            </div>
        `;
    });

    upcomingList.innerHTML = html;
    
    // Show counter
    counterDiv.style.display = 'block';
    updateCarouselCounter();
    
    // Start carousel
    if (totalMeetings > 1) {
        startCarousel();
    }
}

// ===========================
// Carousel Functions
// ===========================
function showCarouselSlide(index) {
    if (index < 0 || index >= totalMeetings) return;
    
    const items = document.querySelectorAll('.upcoming-item');
    items.forEach((item, i) => {
        item.classList.remove('active', 'prev');
        if (i === index) {
            item.classList.add('active');
        } else if (i < index) {
            item.classList.add('prev');
        }
    });
    
    currentCarouselIndex = index;
    updateCarouselCounter();
    
    // Reset timer
    stopCarousel();
    if (totalMeetings > 1) {
        startCarousel();
    }
}

function updateCarouselCounter() {
    document.getElementById('currentCount').textContent = currentCarouselIndex + 1;
    document.getElementById('totalCount').textContent = totalMeetings;
}

function startCarousel() {
    stopCarousel();
    carouselTimer = setTimeout(() => {
        const nextIndex = (currentCarouselIndex + 1) % totalMeetings;
        showCarouselSlide(nextIndex);
    }, CAROUSEL_AUTO_ADVANCE);
}

function stopCarousel() {
    if (carouselTimer) {
        clearTimeout(carouselTimer);
        carouselTimer = null;
    }
}

// ===========================
// Touch Gestures
// ===========================
function attachSwipeGestures() {
    const upcomingList = document.getElementById('upcomingList');
    let touchStartX = 0;
    let touchEndX = 0;

    upcomingList.addEventListener('touchstart', (e) => {
        touchStartX = e.changedTouches[0].screenX;
    }, { passive: true });

    upcomingList.addEventListener('touchend', (e) => {
        touchEndX = e.changedTouches[0].screenX;
        handleSwipe();
    }, { passive: true });

    function handleSwipe() {
        const swipeThreshold = 50;
        const diff = touchStartX - touchEndX;

        if (Math.abs(diff) > swipeThreshold) {
            if (diff > 0) {
                // Swipe left - next
                const nextIndex = (currentCarouselIndex + 1) % totalMeetings;
                showCarouselSlide(nextIndex);
            } else {
                // Swipe right - previous
                const prevIndex = (currentCarouselIndex - 1 + totalMeetings) % totalMeetings;
                showCarouselSlide(prevIndex);
            }
        }
    }
}

// ===========================
// Camera Functions
// ===========================
async function enableCamera() {
    const cameraContainer = document.getElementById('cameraContainer');
    const cameraFeed = document.getElementById('cameraFeed');

    // If camera already active, just show it
    if (cameraStream && cameraFeed.srcObject) {
        cameraContainer.style.display = 'block';
        return;
    }

    try {
        cameraStream = await navigator.mediaDevices.getUserMedia({
            video: {
                width: { ideal: 1280 },
                height: { ideal: 720 }
            },
            audio: false
        });

        cameraFeed.srcObject = cameraStream;
        cameraContainer.style.display = 'block';

        console.log('Camera enabled');
    } catch (error) {
        console.error('Camera access error:', error);
        cameraContainer.style.display = 'none';
    }
}

function disableCamera() {
    const cameraContainer = document.getElementById('cameraContainer');
    const cameraFeed = document.getElementById('cameraFeed');
    
    if (cameraStream) {
        cameraStream.getTracks().forEach(track => track.stop());
        cameraStream = null;
        cameraFeed.srcObject = null;
    }
    
    cameraContainer.style.display = 'none';
    console.log('Camera disabled');
}

// ===========================
// Utility Functions
// ===========================
function setLoadingState() {
    const statusIndicator = document.getElementById('statusIndicator');
    statusIndicator.classList.add('loading');
    statusIndicator.querySelector('.status-text').textContent = 'Loading...';
    
    const currentDisplay = document.getElementById('currentMeetingDisplay');
    currentDisplay.innerHTML = '<div class="loading-message">Loading current meeting...</div>';
}

function formatDateIST(date, format) {
    const formatter = new Intl.DateTimeFormat('en-US', {
        timeZone: IST_TIMEZONE,
        year: 'numeric',
        month: 'long',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: true,
        weekday: 'long'
    });
    
    const parts = formatter.formatToParts(date);
    const result = {};
    
    parts.forEach(part => {
        result[part.type] = part.value;
    });
    
    if (format === 'date') {
        return `${result.weekday}, ${result.month} ${result.day}, ${result.year}`;
    } else if (format === 'time') {
        return `${result.hour}:${result.minute}:${result.second} ${result.dayPeriod}`;
    } else if (format === 'time-only') {
        return `${result.hour}:${result.minute}:${result.second} ${result.dayPeriod}`;
    }
    
    return formatter.format(date);
}

function parseDate(value) {
    if (!value && value !== 0) return new Date(NaN);
    if (value instanceof Date) return value;
    if (typeof value === 'number') {
        // Convert Excel/Sheets serial to JS timestamp
        return new Date(Math.round((value - 25569) * 86400 * 1000));
    }
    return new Date(value);
}

function normalizeMeeting(row) {
    if (!row) return null;

    if (Array.isArray(row)) {
        const [id, room, roomKey, title, start, end, bookedBy, note, participants] = row;
        return {
            id,
            room,
            roomKey,
            title: title || '',
            startDate: parseDate(start),
            endDate: parseDate(end),
            bookedBy: bookedBy || '',
            note: note || '',
            participants: participants || ''
        };
    }

    return {
        ...row,
        title: row.title || row.Subject || '',
        startDate: parseDate(row.start || row.startDate || row.start_time),
        endDate: parseDate(row.end || row.endDate || row.end_time),
        bookedBy: row.bookedBy || row.booker || row.organizer || '',
        participants: row.participants || '',
        note: row.note || row.description || ''
    };
}

function formatTime(dateInput) {
    const date = dateInput instanceof Date ? dateInput : new Date(dateInput);
    if (isNaN(date)) return '';
    return formatDateIST(date, 'time-only');
}

function formatTimeRange(startInput, endInput) {
    return `${formatTime(startInput)} - ${formatTime(endInput)}`;
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function displayError() {
    const displayContainer = document.getElementById('currentMeetingDisplay');
    displayContainer.innerHTML = `
        <div class="no-meeting">
            <div class="no-meeting-icon" style="color: #ff4c4c;">⚠</div>
            <div class="no-meeting-text" style="color: #ff8a8a;">Connection Error</div>
            <div class="no-meeting-subtext">Unable to load schedule</div>
        </div>
    `;
    
    const upcomingList = document.getElementById('upcomingList');
    upcomingList.innerHTML = '<div class="no-upcoming">Unable to load schedule</div>';
}

// ===========================
// Cleanup
// ===========================
window.addEventListener('beforeunload', () => {
    disableCamera();
    if (refreshTimer) clearInterval(refreshTimer);
    stopCarousel();
});
