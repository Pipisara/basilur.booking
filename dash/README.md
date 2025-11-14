# Meeting Room Booking System - Display Dashboard

A modern, real-time meeting room display system built with Google Apps Script, HTML5, CSS3, and JavaScript. Perfect for office environments to show room availability, current meetings, and upcoming bookings.

## 📋 Table of Contents

- [Features](#features)
- [System Architecture](#system-architecture)
- [Setup Instructions](#setup-instructions)
- [Opening Rooms Using Links](#opening-rooms-using-links)
- [Deployment](#deployment)
- [Configuration](#configuration)
- [File Structure](#file-structure)
- [Technical Details](#technical-details)
- [Troubleshooting](#troubleshooting)

## ✨ Features

### Room Display Features
- **Real-time Room Status**: Live availability indicators (Available/In Use)
- **Current Meeting Display**: Shows active meeting details with countdown timer
- **Upcoming Meetings**: Today's schedule with organizer and participant information
- **Live Clock**: Current time in IST (Asia/Kolkata) timezone
- **Progress Bar**: Visual representation of meeting progress
- **Auto-refresh**: Updates every 60 seconds automatically
- **Responsive Design**: Works on any display size (TVs, tablets, monitors)
- **Live Camera Feed**: Optional webcam display during active meetings

### Backend Features
- **Google Apps Script Integration**: Serverless backend using Google Sheets
- **Data Persistence**: Stores bookings in Google Sheets
- **Email Notifications**: Automatic booking confirmation emails
- **Multi-room Support**: Manage multiple meeting rooms
- **User Authentication**: Role-based access control
- **Timezone Support**: IST (GMT+5:30) timezone for all timestamps

## 🏗️ System Architecture

```
Meeting.Window
├── code.gs          (Google Apps Script backend)
├── index.html       (Main UI)
├── script.js        (Frontend logic)
├── style.css        (Styling)
└── logo.png         (Company logo)
```

### Component Breakdown

| Component | Purpose |
|-----------|---------|
| `code.gs` | Google Apps Script handling API requests, data storage, and email notifications |
| `index.html` | HTML structure for the display dashboard |
| `script.js` | Frontend JavaScript for real-time updates and room logic |
| `style.css` | Modern CSS styling with gradient backgrounds and animations |

## 🚀 Setup Instructions

### Step 1: Create a Google Apps Script Project

1. Go to [Google Apps Script](https://script.google.com)
2. Click **New Project**
3. Copy and paste the contents of `code.gs` into the editor
4. Save the project with a name (e.g., "Meeting Room Dashboard")

### Step 2: Set Up Google Sheets

1. Create a new Google Sheet linked to your Apps Script project
2. The script will automatically create three sheets:
   - **Bookings**: Stores all room bookings
   - **Users**: Stores user credentials and roles
   - **Email Log**: Tracks email notifications

### Step 3: Enable Gmail API

1. In Apps Script editor, click **Services** (+ button)
2. Search for and enable **Gmail API**
3. Ensure **Advanced Google Services** is enabled for Gmail

### Step 4: Deploy as Web App

1. In Apps Script editor, click **Deploy** → **New Deployment**
2. Select type: **Web app**
3. Set as follows:
   - Execute as: Your Google account
   - Who has access: **Anyone**
4. Click **Deploy**
5. Copy the generated deployment URL

### Step 5: Update API URL in Frontend

1. Open `script.js`
2. Find line 1: `const API_URL = '...'`
3. Replace with your deployment URL

## 🔗 Opening Rooms Using Links

The system supports multiple meeting rooms via URL parameters. Use the room query parameter to display different rooms.

### Available Rooms

| Room Key | Room Name | Code |
|----------|-----------|------|
| A | BLOCK A BOARDROOM | `?room=A` |
| B | BLOCK C BOARDROOM | `?room=B` |
| C | BLOCK D AUDITORIUM | `?room=C` |

### Room Link Examples

Use these URLs to open specific rooms:

**Room A (Block A Boardroom)**
```
https://[your-deployment-url]?room=A
```

**Room B (Block C Boardroom)**
```
https://[your-deployment-url]?room=B
```

**Room C (Block D Auditorium)**
```
https://[your-deployment-url]?room=C
```

**Default (Room A if no parameter provided)**
```
https://[your-deployment-url]
```

### How It Works

1. The frontend reads the URL parameter `?room=X` using `URLSearchParams`
2. If no parameter provided, defaults to Room A
3. The room key is sent to the backend API
4. Backend filters bookings by room and returns current/upcoming meetings
5. Display updates with room-specific information

### Example: Setting Up Wall Displays

For office wall displays, you can:

1. **Create bookmarks** in a browser with the direct links:
   - Conference Room A: `https://[url]?room=A`
   - Conference Room B: `https://[url]?room=B`
   - Auditorium: `https://[url]?room=C`

2. **Use QR codes** to link to each room:
   - Generate QR codes for each room link
   - Display QR codes outside each meeting room
   - Scan to view live room status

3. **Display on tablets/TVs** using digital signage software:
   - Load the room-specific URL in fullscreen mode
   - Auto-refresh ensures real-time updates

## 🔧 Deployment

### Deploy as Google Apps Script Web App

1. **Create Deployment**
   - Go to Deploy → New Deployment
   - Select "Web app" type
   - Set "Execute as" to your Google account
   - Set "Who has access" to "Anyone"

2. **Get Your URL**
   - After deployment, you'll receive a URL like:
   ```
   https://script.google.com/macros/s/[DEPLOYMENT-ID]/userweb
   ```

3. **Update code.gs with API URL**
   - If needed, update any API endpoints in the code

4. **Access the Dashboard**
   - Room A: `[YOUR-URL]?room=A`
   - Room B: `[YOUR-URL]?room=B`
   - Room C: `[YOUR-URL]?room=C`

## ⚙️ Configuration

### Timezone Settings

All timestamps use **Asia/Kolkata (IST, GMT+5:30)** timezone.

To change timezone:

1. In `code.gs`, find `formatDateAsIST()` function
2. Modify the timezone string (e.g., `'America/New_York'`, `'Europe/London'`)
3. In `script.js`, find `IST_TIMEZONE` constant and update similarly

### Refresh Interval

Change how often the display updates:

In `script.js`, line 3:
```javascript
const REFRESH_INTERVAL = 60000; // 60 seconds
// Change to: 30000 for 30 seconds, 120000 for 2 minutes, etc.
```

### Room Names

Customize room names in `script.js`:

```javascript
const ROOM_NAMES = {
    'A': 'BLOCK A BOARDROOM',
    'B': 'BLOCK C BOARDROOM',
    'C': 'BLOCK D AUDITORIUM'
};
```

### Email Configuration

Add custom email settings in `code.gs`:

```javascript
// Look for the sendEmailNotification() function
// Customize email subjects, templates, and recipients
```

### Company Branding

1. Replace `logo.png` with your company logo
2. Update footer text in `index.html`:
   ```html
   <span>Company Name © 2025</span>
   ```
3. Modify colors in `style.css`:
   - Background gradient: Lines 14-18
   - Primary color: Search for `#00bcd4`
   - Secondary colors: Search for `#b0f0ff`

## 📁 File Structure

### code.gs (Google Apps Script Backend)
- **doGet()**: Handles GET requests for bookings
- **doPost()**: Handles POST requests for creating/updating/deleting bookings
- **handleGetRoomSchedule()**: Returns current and upcoming meetings for a specific room
- **handleAuthentication()**: User login verification
- **getOrCreateBookingsSheet()**: Manages bookings data
- **getOrCreateUsersSheet()**: Manages user credentials
- **getOrCreateEmailLogSheet()**: Tracks email notifications
- **sendEmailNotification()**: Sends booking confirmation emails

### index.html (Frontend Structure)
- Header with room name and time display
- Status indicator (Available/In Use)
- Current meeting display section
- Upcoming meetings list
- Camera feed container (optional)
- Footer with timestamp

### script.js (Frontend Logic)
- **loadRoomSchedule()**: Fetches data from backend API
- **updateCurrentMeeting()**: Updates current meeting display
- **updateUpcomingMeetings()**: Updates upcoming meetings list
- **enableCamera()**: Activates webcam for live feed
- **formatDateIST()**: Converts dates to IST timezone
- **normalizeMeeting()**: Standardizes meeting data format
- **Utility functions**: HTML escaping, date parsing, error handling

### style.css (Styling)
- Modern gradient background
- Responsive flexbox layout
- Animated status indicators
- Meeting cards with detailed information
- Progress bars for time tracking
- Mobile-friendly design

## 🔍 Technical Details

### API Endpoints

**GET Bookings**
```
?action=bookings
```
Returns all bookings in the system.

**GET Room Schedule**
```
?action=getRoomSchedule&room=A
```
Returns current and upcoming meetings for a specific room (A, B, or C).

**POST Create/Update/Delete**
```
POST with JSON payload:
{
  "action": "create|update|delete",
  "room": "Room Name",
  "roomKey": "A|B|C",
  "title": "Meeting Title",
  "start": "2025-11-14 14:00:00",
  "end": "2025-11-14 15:00:00",
  "bookedBy": "John Doe",
  "participants": "john@company.com, jane@company.com",
  "note": "Optional notes"
}
```

### Data Structure

**Bookings Sheet Columns**
| Column | Type | Description |
|--------|------|-------------|
| ID | String | Unique booking identifier |
| Room | String | Display name (e.g., "Block A Boardroom") |
| RoomKey | String | Single letter code (A, B, C) |
| Title | String | Meeting title |
| Start | DateTime | Meeting start time (IST) |
| End | DateTime | Meeting end time (IST) |
| Booked By | String | Organizer name/email |
| Note | String | Meeting notes |
| Participants | String | Comma-separated participant list |
| Email Sent | Boolean | Whether confirmation email was sent |
| Created By | String | User who created booking |
| Updated By | String | User who last updated booking |
| Created At | DateTime | Booking creation timestamp |
| Updated At | DateTime | Last update timestamp |

### Browser Compatibility

- ✅ Chrome 90+
- ✅ Firefox 88+
- ✅ Safari 14+
- ✅ Edge 90+
- ✅ Mobile browsers (iOS Safari, Chrome Mobile)

### Performance

- **Auto-refresh**: 60 seconds (configurable)
- **API calls**: Minimal bandwidth usage
- **Real-time updates**: No polling required for time display
- **Camera stream**: Optional, disabled by default

## 🐛 Troubleshooting

### Issue: Display shows "Connection Error"

**Solution:**
1. Verify the deployment URL is correct in `script.js`
2. Check that the Google Apps Script is deployed and accessible
3. Ensure **Anyone** access is enabled in deployment settings
4. Check browser console (F12) for specific error messages

### Issue: Times showing incorrectly

**Solution:**
1. Verify timezone is set to `Asia/Kolkata` in both `code.gs` and `script.js`
2. Check system timezone on the display device
3. Ensure Google Sheet is set to IST timezone in sheet settings

### Issue: Camera feed not showing

**Solution:**
1. Camera is only shown during active meetings
2. Check browser permissions (Allow Camera access)
3. Ensure device has webcam connected
4. Try in different browser (some have stricter security)
5. HTTPS is required for camera access in most browsers

### Issue: Bookings not appearing

**Solution:**
1. Verify bookings exist in Google Sheet
2. Check that room key (A, B, or C) matches exactly
3. Verify booking is within the correct date range
4. Check booking start/end times are in IST

### Issue: Emails not sending

**Solution:**
1. Verify Gmail API is enabled in Apps Script services
2. Check email addresses in booking records
3. Review Email Log sheet for error messages
4. Ensure sending account has Gmail quota available

### Issue: Multiple rooms showing same content

**Solution:**
1. Check that different room parameters are used in URLs
2. Verify API is correctly filtering by room key
3. Clear browser cache and refresh
4. Try different browser or incognito window

## 📝 Booking Data Entry

Add bookings to the system by:

1. **Direct Sheet Entry**: Add rows to the Bookings sheet manually
2. **API POST**: Send booking data via HTTP POST request
3. **Booking Form**: Create a separate form that submits to the API

### Manual Entry Format

When adding bookings directly to the sheet, ensure:
- Date/Time format: `YYYY-MM-DD HH:MM:SS`
- Room Key must be: `A`, `B`, or `C`
- All required fields are filled

## 🔐 Security Notes

- The deployed web app has "Anyone" access for public viewing
- Authentication is available but optional for viewing
- User credentials are stored in Google Sheet (encrypted by Google)
- All data is encrypted in transit (HTTPS)
- Consider restricting access via domain if sensitive

## 📞 Support & Customization

### Common Customizations

1. **Add more rooms**: Add entries to `ROOM_NAMES` and modify backend
2. **Change colors**: Edit CSS variables in `style.css`
3. **Adjust fonts**: Modify `font-family` and `font-size` in CSS
4. **Add logo**: Replace `logo.png` file
5. **Customize messages**: Edit text in `index.html` and `script.js`

### Integration Points

- **Google Calendar**: Sync bookings from Google Calendar via API
- **Slack**: Send booking notifications to Slack
- **Teams**: Integrate with Microsoft Teams calendar
- **Custom CRM**: Export bookings to external systems

## 📄 License

This project is provided as-is. Modify as needed for your organization.

## 🎯 Version History

- **v3.2** (2024-11-13): Current version without ICS files
- Removed ICS file generation
- Improved timezone handling
- Enhanced email notifications
- Better error handling

---

**Last Updated**: November 14, 2025  
**Timezone**: Asia/Kolkata (IST, GMT+5:30)  
**Company**: Basilur Tea Export © 2025
