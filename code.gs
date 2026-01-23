// ============================================================================
// Meeting Room Booking System with Gmail Integration - Google Apps Script
// ============================================================================
// Deploy as Web App with "Anyone" access
// Enable Gmail API in Advanced Google Services
// Version: 3.2 without ICS files
// Last Updated: 2024-11-13
// TIMEZONE: Asia/Kolkata (IST, GMT+5:30) - Used for all timestamp operations
// ============================================================================

/**
 * Handle GET requests - Retrieve all bookings or authenticate
 */
function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action ? e.parameter.action : 'bookings';
    
    if (action === 'auth') {
      return handleAuthentication(e);
    }
    
    if (action === 'getRoomSchedule') {
      return handleGetRoomSchedule(e);
    }
    
    // Return all bookings
    const sheet = getOrCreateBookingsSheet();
    const data = sheet.getDataRange().getValues();
    const bookings = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      
      bookings.push({
        id: row[0] || '',
        room: row[1] || '',
        roomKey: row[2] || '',
        title: row[3] || '',
        start: row[4] || '',
        end: row[5] || '',
        bookedBy: row[6] || '',
        note: row[7] || '',
        participants: row[8] || '',
        emailSent: row[9] || false,
        createdBy: row[10] || '',
        updatedBy: row[11] || '',
        createdAt: row[12] || '',
        updatedAt: row[13] || ''
      });
    }
    
    return ContentService
      .createTextOutput(JSON.stringify(bookings))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log('doGet Error: ' + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({
        error: error.toString(),
        message: 'Failed to retrieve data'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handle POST requests - Create, update, or delete bookings
 */
function doPost(e) {
  try {
    const booking = JSON.parse(e.postData.contents);
    const sheet = getOrCreateBookingsSheet();
    const action = booking.action || 'create';
    
    Logger.log('doPost: Action = ' + action + ', ID = ' + (booking.id || 'new'));
    
    if (action === 'update') {
      return handleUpdate(sheet, booking);
    } else if (action === 'delete') {
      return handleDelete(sheet, booking);
    } else {
      return handleCreate(sheet, booking);
    }
    
  } catch (error) {
    Logger.log('doPost Error: ' + error.toString());
    
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handle user authentication
 */
function handleAuthentication(e) {
  const name = e.parameter && e.parameter.name ? e.parameter.name : '';
  const accessCode = e.parameter && e.parameter.accessCode ? e.parameter.accessCode : '';
  
  const usersSheet = getOrCreateUsersSheet();
  const usersData = usersSheet.getDataRange().getValues();
  
  let found = null;
  for (let i = 1; i < usersData.length; i++) {
    const row = usersData[i];
    if (row[0] && String(row[0]).trim() === String(name).trim() && 
        String(row[1]).trim() === String(accessCode).trim()) {
      found = { 
        name: row[0], 
        email: row[2] || '', 
        role: row[3] || 'user' 
      };
      break;
    }
  }
  
  if (found) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', user: found }))
      .setMimeType(ContentService.MimeType.JSON);
  } else {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Invalid credentials' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handle getting room schedule (for display boards)
 */
function handleGetRoomSchedule(e) {
  const roomKey = e.parameter && e.parameter.room ? e.parameter.room : '';
  
  if (!roomKey) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: 'Room key is required'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const sheet = getOrCreateBookingsSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  
  let currentMeeting = null;
  const upcomingMeetings = [];
  
  // Filter bookings by room and time
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Column indices based on: id, room, roomKey, title, start, end, bookedBy, note, participants, emailSent, createdBy, updatedBy, createdAt, updatedAt
    const bookingRoomKey = row[2] || '';
    
    if (bookingRoomKey.toString() !== roomKey.toString()) {
      continue;
    }
    
    const start = new Date(row[4]);
    const end = new Date(row[5]);
    
    // Check if meeting is currently active
    if (now >= start && now < end) {
      currentMeeting = {
        id: row[0] || '',
        room: row[1] || '',
        roomKey: row[2] || '',
        title: row[3] || '',
        start: row[4] || '',
        end: row[5] || '',
        bookedBy: row[6] || '',
        note: row[7] || '',
        participants: row[8] || ''
      };
    }
    // Check if meeting is in the future (today)
    else if (now < start && start.toDateString() === now.toDateString()) {
      upcomingMeetings.push({
        id: row[0] || '',
        room: row[1] || '',
        roomKey: row[2] || '',
        title: row[3] || '',
        start: row[4] || '',
        end: row[5] || '',
        bookedBy: row[6] || '',
        note: row[7] || '',
        participants: row[8] || ''
      });
    }
  }
  
  // Sort upcoming meetings by start time
  upcomingMeetings.sort((a, b) => new Date(a.start) - new Date(b.start));
  
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'success',
      currentMeeting: currentMeeting,
      upcomingMeetings: upcomingMeetings,
      timestamp: formatDateAsIST(now, 'yyyy-MM-dd HH:mm:ss')
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Get or create Users sheet
 */
function getOrCreateUsersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Users');
  
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    const headers = ['Name', 'Access Code', 'Email', 'Role', 'Created At', 'Updated At'];
    sheet.appendRow(headers);
    
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#34a853');
    headerRange.setFontColor('#ffffff');
    
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 250);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 160);
    sheet.setColumnWidth(6, 160);
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * Get or create Bookings sheet
 */
function getOrCreateBookingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Bookings');
  
  if (!sheet) {
    Logger.log('Creating new Bookings sheet with headers');
    
    sheet = ss.insertSheet('Bookings');
    
    const headers = [
      'ID', 'Room', 'RoomKey', 'Title', 'Start', 'End', 
      'Booked By', 'Note', 'Participants', 'Email Sent',
      'Created By', 'Updated By', 'Created At', 'Updated At', 'Calendar Event ID'
    ];
    
    sheet.appendRow(headers);
    
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    
    sheet.setColumnWidth(1, 120);  // ID
    sheet.setColumnWidth(2, 80);   // Room
    sheet.setColumnWidth(3, 60);   // RoomKey
    sheet.setColumnWidth(4, 200);  // Title
    sheet.setColumnWidth(5, 150);  // Start
    sheet.setColumnWidth(6, 150);  // End
    sheet.setColumnWidth(7, 150);  // Booked By
    sheet.setColumnWidth(8, 300);  // Note
    sheet.setColumnWidth(9, 300);  // Participants
    sheet.setColumnWidth(10, 100); // Email Sent
    sheet.setColumnWidth(11, 120); // Created By
    sheet.setColumnWidth(12, 120); // Updated By
    sheet.setColumnWidth(13, 150); // Created At
    sheet.setColumnWidth(14, 150); // Updated At
    sheet.setColumnWidth(15, 150); // Calendar Event ID
    
    sheet.setFrozenRows(1);
  } else {
    // Check if new column needs to be added to existing sheet
    if (sheet.getLastColumn() < 15) {
      Logger.log('Adding missing "Calendar Event ID" column');
      sheet.insertColumnAfter(14);
      const headerCell = sheet.getRange(1, 15);
      headerCell.setValue('Calendar Event ID');
      headerCell.setFontWeight('bold');
      headerCell.setBackground('#4285f4');
      headerCell.setFontColor('#ffffff');
      sheet.setColumnWidth(15, 150);
    }
  }
  
  return sheet;
}

/**
 * Get or create Email Log sheet
 */
function getOrCreateEmailLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Email Log');
  
  if (!sheet) {
    sheet = ss.insertSheet('Email Log');
    
    const headers = [
      'Timestamp', 'Booking ID', 'Meeting Title', 'Recipient Email', 
      'Subject', 'Status', 'Error Message'
    ];
    
    sheet.appendRow(headers);
    
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#ea4335');
    headerRange.setFontColor('#ffffff');
    
    sheet.setColumnWidth(1, 150);  // Timestamp
    sheet.setColumnWidth(2, 120);  // Booking ID
    sheet.setColumnWidth(3, 200);  // Meeting Title
    sheet.setColumnWidth(4, 250);  // Recipient Email
    sheet.setColumnWidth(5, 300);  // Subject
    sheet.setColumnWidth(6, 100);  // Status
    sheet.setColumnWidth(7, 300);  // Error Message
    
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

function handleCreate(sheet, booking) {
  const now = new Date();
  const bookingId = now.getTime().toString();
  
  Logger.log('Creating booking: ' + bookingId + ' - ' + booking.title);
  
  const emailSent = false;
  let calendarEventId = '';
  let googleCalendarUrl = '';
  
  // 1. Google Calendar Integration
  if (booking.manualAddToCalendar === true) {
    Logger.log('Manual add requested. Generating ID only.');
    googleCalendarUrl = generateGoogleCalendarLink(booking);
  } else {
    // Auto-create on Google Calendar
    calendarEventId = createGoogleCalendarEvent(booking) || '';
  }
  
  // 2. Add to Sheet (including new Calendar Event ID column)
  sheet.appendRow([
    bookingId,
    booking.room || '',
    booking.roomKey || '',
    booking.title || 'Untitled Meeting',
    booking.start || '',
    booking.end || '',
    booking.bookedBy || '',
    booking.note || '',
    booking.participants || '',
    emailSent,
    booking.createdBy || booking.bookedBy || '',
    '',
    formatDateAsIST(now, 'yyyy-MM-dd HH:mm:ss'),
    '',
    calendarEventId // Column 15
  ]);
  
  // 3. Send HTML Notification Email (Branding/Notification only)
  if (booking.sendEmail && booking.participants) {
    try {
      sendMeetingInvitation(booking, bookingId, false);
      
      // Update email sent status
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString() === bookingId.toString()) {
          sheet.getRange(i + 1, 10).setValue(true);
          break;
        }
      }
    } catch (emailError) {
      Logger.log('Email Error: ' + emailError.toString());
      logEmailError(bookingId, booking.title, '', 'Failed to send invitation', emailError.toString());
    }
  }
  
  Logger.log('Booking created successfully: ' + bookingId);
  
  const response = {
    status: 'success',
    action: 'create',
    id: bookingId,
    message: 'Booking created successfully',
    calendarEventId: calendarEventId
  };
  
  if (googleCalendarUrl) {
    response.googleCalendarUrl = googleCalendarUrl;
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Update an existing booking
 */
function handleUpdate(sheet, booking) {
  const bookingId = booking.id || booking.originalId;
  
  if (!bookingId) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Booking ID required'})).setMimeType(ContentService.MimeType.JSON);
  }
  
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  let calendarEventId = '';
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === bookingId.toString()) {
      rowIndex = i + 1;
      calendarEventId = data[i][14] || ''; // Get existing Calendar ID
      break;
    }
  }
  
  if (rowIndex === -1) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Booking not found'})).setMimeType(ContentService.MimeType.JSON);
  }
  
  // Update Google Calendar Event
  if (calendarEventId) {
    updateGoogleCalendarEvent(calendarEventId, booking);
  } else if (!booking.manualAddToCalendar) {
    // If no ID exists (old booking), try to create one now?
    // User didn't strictly specify, but good practice.
    // For now, let's just create if missing and not manual
    calendarEventId = createGoogleCalendarEvent(booking) || '';
  }
  
  const now = new Date();
  const existingCreatedAt = data[rowIndex - 1][12];
  const existingCreatedBy = data[rowIndex - 1][10];
  const existingEmailSent = data[rowIndex - 1][9];
  
  // Note: We retrieve 15 columns now (indexes 0-14)
  // But we need to be careful if the sheet doesn't have 15 cols yet (creating it handles headers, but existing data rows might be short)
  // Arrays in Apps Script extend automatically.
  
  sheet.getRange(rowIndex, 1, 1, 15).setValues([[
    bookingId,
    booking.room || '',
    booking.roomKey || '',
    booking.title || 'Untitled Meeting',
    booking.start || '',
    booking.end || '',
    booking.bookedBy || '',
    booking.note || '',
    booking.participants || '',
    existingEmailSent,
    existingCreatedBy || '',
    booking.updatedBy || booking.submittedBy || '',
    existingCreatedAt || '',
    formatDateAsIST(now, 'yyyy-MM-dd HH:mm:ss'),
    calendarEventId
  ]]);
  
  // Send updated HTML invitation if requested
  if (booking.sendEmail && booking.participants) {
    try {
      sendMeetingInvitation(booking, bookingId, true);
      sheet.getRange(rowIndex, 10).setValue(true);
    } catch (emailError) {
      Logger.log('Email Error: ' + emailError.toString());
      logEmailError(bookingId, booking.title, '', 'Failed to send update', emailError.toString());
    }
  }
  
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'success',
      action: 'update',
      id: bookingId,
      message: 'Booking updated successfully'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Delete a booking
 */
function handleDelete(sheet, booking) {
  const bookingId = booking.id;
  
  if (!bookingId) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Booking ID required'})).setMimeType(ContentService.MimeType.JSON);
  }
  
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  let bookingData = null;
  let calendarEventId = '';
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === bookingId.toString()) {
      rowIndex = i + 1;
      calendarEventId = data[i][14] || '';
      bookingData = {
        title: data[i][3],
        start: data[i][4],
        end: data[i][5],
        room: data[i][1],
        bookedBy: data[i][6],
        participants: data[i][8],
        note: data[i][7]
      };
      break;
    }
  }
  
  if (rowIndex === -1) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Booking not found'})).setMimeType(ContentService.MimeType.JSON);
  }
  
  // Delete from Google Calendar
  if (calendarEventId) {
    deleteGoogleCalendarEvent(calendarEventId);
  }
  
  // Send cancellation email
  if (booking.sendCancellation && bookingData && bookingData.participants) {
    try {
      // Logic: Google Calendar sends cancellation if we deleted the event and sendUpdates='all'.
      // But we also want to send our HTML branded email.
      sendCancellationEmail(bookingData, bookingId, booking.deletedBy || 'Unknown');
    } catch (emailError) {
      Logger.log('Cancellation Email Error: ' + emailError.toString());
      logEmailError(bookingId, bookingData.title, '', 'Failed to send cancellation', emailError.toString());
    }
  }
  
  sheet.deleteRow(rowIndex);
  
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'success',
      action: 'delete',
      id: bookingId,
      message: 'Booking deleted successfully'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Send meeting invitation email with MIME calendar format
 */
function sendMeetingInvitation(booking, bookingId, isUpdate) {
  if (!booking.participants) return;
  
  const emails = booking.participants.split(',').map(function(e) { return e.trim(); }).filter(function(e) { return e; });
  if (emails.length === 0) return;
  
  const startDate = new Date(booking.start);
  const endDate = new Date(booking.end);
  const IST_TIMEZONE = 'Asia/Kolkata';
  
  const formattedDate = Utilities.formatDate(startDate, IST_TIMEZONE, 'EEEE, MMMM dd, yyyy');
  const formattedStartTime = Utilities.formatDate(startDate, IST_TIMEZONE, 'hh:mm a');
  const formattedEndTime = Utilities.formatDate(endDate, IST_TIMEZONE, 'hh:mm a');
  
  const subject = (isUpdate ? 'Updated: ' : '') + booking.title;
  
  // Build HTML email
  const emailHtml = getMeetingHtmlTemplate(isUpdate)
    .replace(/{{MEETING_TITLE}}/g, booking.title)
    .replace(/{{MEETING_ROOM}}/g, booking.room)
    .replace(/{{MEETING_DATE}}/g, formattedDate)
    .replace(/{{START_TIME}}/g, formattedStartTime)
    .replace(/{{END_TIME}}/g, formattedEndTime)
    .replace(/{{ORGANIZER}}/g, booking.bookedBy)
    .replace(/{{NOTES}}/g, booking.note || 'No additional notes')
    .replace(/{{BOOKING_ID}}/g, bookingId);
  
  const organizerEmail = Session.getActiveUser().getEmail() || 'noreply@basilurtea.com';
  const organizerDisplayName = booking && booking.bookedBy ? booking.bookedBy : 'Meeting Organizer';
  
  emails.forEach(function(email) {
    try {
      // Send simple HTML email
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: emailHtml,
        name: organizerDisplayName
      });
      
      logEmailSuccess(bookingId, booking.title, email, subject, isUpdate ? 'Updated' : 'Scheduled');
      Logger.log('HTML invitation sent to: ' + email);
    } catch (error) {
      Logger.log('Failed to send email to ' + email + ': ' + error.toString());
      logEmailError(bookingId, booking.title, email, subject, error.toString());
    }
  });
}

/**
 * Send cancellation email with calendar cancellation
 */
function sendCancellationEmail(bookingData, bookingId, deletedBy) {
  if (!bookingData.participants) return;
  
  const emails = bookingData.participants.split(',').map(function(e) { return e.trim(); }).filter(function(e) { return e; });
  if (emails.length === 0) return;
  
  const startDate = new Date(bookingData.start);
  const IST_TIMEZONE = 'Asia/Kolkata';
  const formattedDate = Utilities.formatDate(startDate, IST_TIMEZONE, 'EEEE, MMMM dd, yyyy');
  const formattedStartTime = Utilities.formatDate(startDate, IST_TIMEZONE, 'hh:mm a');
  
  const subject = 'Canceled: ' + bookingData.title;
  
  const emailHtml = getCancellationHtmlTemplate()
    .replace(/{{MEETING_TITLE}}/g, bookingData.title)
    .replace(/{{MEETING_ROOM}}/g, bookingData.room)
    .replace(/{{MEETING_DATE}}/g, formattedDate)
    .replace(/{{START_TIME}}/g, formattedStartTime)
    .replace(/{{CANCELED_BY}}/g, deletedBy)
    .replace(/{{BOOKING_ID}}/g, bookingId);

  const organizerDisplayName = bookingData && bookingData.bookedBy ? bookingData.bookedBy : 'Organizer';

  emails.forEach(function(email) {
    try {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: emailHtml,
        name: organizerDisplayName
      });

      logEmailSuccess(bookingId, bookingData.title, email, subject, 'Canceled');
      Logger.log('Cancellation sent to: ' + email);
    } catch (error) {
      Logger.log('Failed to send cancellation email to ' + email + ': ' + error.toString());
      logEmailError(bookingId, bookingData.title, email, subject, error.toString());
    }
  });
}

/**
 * HTML Template for Meeting Invitation
 */
function getMeetingHtmlTemplate(isUpdate) {
  const headerText = isUpdate ? 'Meeting Updated' : 'Meeting Invitation';
  const introText = isUpdate ? 'This meeting has been <strong style="color:#ff9800;">UPDATED</strong>. Please note the new details:' : 'You have been invited to a meeting:';
  
  return `
<!DOCTYPE html> 
<html lang="en">
<head>
<meta charset="UTF-8">
<link rel="icon" href="//www.basilurtea.com/cdn/shop/files/Basilur.png" type="image/png">
<link rel="apple-touch-icon" href="//www.basilurtea.com/cdn/shop/files/Basilur.png">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${headerText}</title>
</head>
<body style="margin:0;padding:0;background:linear-gradient(135deg,#0f2027,#203a43,#2c5364);font-family:'Poppins',sans-serif;color:#f5f5f5;">

  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="padding:40px 0;">
    <tr>
      <td align="center">
        <table width="600" cellpadding="0" cellspacing="0" border="0" style="background:rgba(15,32,39,0.95);border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.4);">

          <!-- Header -->
          <tr>
            <td align="center" style="padding:20px;background:#2c5364;border-bottom:2px solid #00bcd4;">
            <img src="https://www.basilurtea.com/cdn/shop/files/Basilur.png" alt="Company Logo" width="60" style="display:block;margin-bottom:10px;">              <h1 style="color:#00bcd4;font-size:20px;margin:0;">${headerText}</h1>
            </td>
          </tr>

          <!-- Body -->
          <tr>
            <td style="padding:30px 40px;">
              <p style="font-size:16px;margin-bottom:10px;">${introText}</p>

              <div style="background:rgba(32,58,67,0.9);padding:20px;border-radius:10px;border:1px solid rgba(0,188,212,0.35);margin-bottom:25px;line-height:1.6;">
                <p style="margin:0 0 8px;"><strong style="color:#00bcd4;">&#128197; Meeting:</strong> {{MEETING_TITLE}}</p>
                <p style="margin:0 0 8px;"><strong style="color:#00bcd4;">&#127970; Room:</strong> {{MEETING_ROOM}}</p>
                <p style="margin:0 0 8px;"><strong style="color:#00bcd4;">&#128198; Date:</strong> {{MEETING_DATE}}</p>
                <p style="margin:0 0 8px;"><strong style="color:#00bcd4;">&#128336; Start Time:</strong> {{START_TIME}}</p>
                <p style="margin:0 0 8px;"><strong style="color:#00bcd4;">&#128337; End Time:</strong> {{END_TIME}}</p>
                <p style="margin:0 0 8px;"><strong style="color:#00bcd4;">&#128100; Organized by:</strong> {{ORGANIZER}}</p>
                <p style="margin:15px 0 8px;"><strong style="color:#00bcd4;">&#128221; Notes:</strong> {{NOTES}}</p>
                <p style="margin:10px 0 0;font-size:14px;color:#b0b0b0;">Booking ID: {{BOOKING_ID}}</p>
              </div>

              <p style="font-size:15px;color:#00bcd4;font-weight:600;margin-top:20px;text-align:center;">
                📅 This invitation has been added to your calendar
              </p>
              
              <p style="font-size:13px;color:#888;margin-top:15px;text-align:center;">
                Check your calendar app (Gmail, Outlook, etc.) to accept or decline.
              </p>
            </td>
          </tr>

          <!-- Footer -->
          <tr>
            <td align="center" style="background:#203a43;padding:15px;border-top:1px solid rgba(245,245,245,0.08);font-size:12px;color:#b0b0b0;">
              © 2025 Basilur Tea Export. All rights reserved.
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>

</body>
</html>
`;
}

/**
 * HTML Template for Meeting Cancellation
 */
function getCancellationHtmlTemplate() {
  return `
<!DOCTYPE html> 
<html lang="en">
<head>
<meta charset="UTF-8">
<link rel="icon" href="//www.basilurtea.com/cdn/shop/files/Basilur.png" type="image/png">
<link rel="apple-touch-icon" href="//www.basilurtea.com/cdn/shop/files/Basilur.png">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Meeting Canceled</title>
</head>
<body style="margin:0;padding:0;background:linear-gradient(135deg,#0f2027,#203a43,#2c5364);font-family:'Poppins',sans-serif;color:#f5f5f5;">

  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="padding:40px 0;">
    <tr>
      <td align="center">
        <table width="600" cellpadding="0" cellspacing="0" border="0" style="background:rgba(15,32,39,0.95);border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.4);">

          <!-- Header -->
          <tr>
            <td align="center" style="padding:20px;background:#c62828;border-bottom:2px solid #ff5252;">
              <img src="https://www.basilurtea.com/cdn/shop/files/Basilur.png" alt="Company Logo" width="60" style="display:block;margin-bottom:10px;">
              <h1 style="color:#ff5252;font-size:20px;margin:0;">&#10060; Meeting Canceled</h1>
            </td>
          </tr>

          <!-- Body -->
          <tr>
            <td style="padding:30px 40px;">
              <p style="font-size:16px;margin-bottom:10px;color:#ff5252;font-weight:600;">The following meeting has been CANCELED:</p>

              <div style="background:rgba(32,58,67,0.9);padding:20px;border-radius:10px;border:1px solid rgba(255,82,82,0.35);margin-bottom:25px;line-height:1.6;">
                <p style="margin:0 0 8px;"><strong style="color:#ff5252;">&#128197; Meeting:</strong> {{MEETING_TITLE}}</p>
                <p style="margin:0 0 8px;"><strong style="color:#ff5252;">&#127970; Room:</strong> {{MEETING_ROOM}}</p>
                <p style="margin:0 0 8px;"><strong style="color:#ff5252;">&#128198; Date:</strong> {{MEETING_DATE}}</p>
                <p style="margin:0 0 8px;"><strong style="color:#ff5252;">&#128336; Time:</strong> {{START_TIME}}</p>
                <p style="margin:15px 0 8px;"><strong style="color:#ff5252;">&#10060; Canceled by:</strong> {{CANCELED_BY}}</p>
                <p style="margin:10px 0 0;font-size:14px;color:#b0b0b0;">Booking ID: {{BOOKING_ID}}</p>
              </div>

              <p style="font-size:13px;color:#888;margin-top:25px;text-align:center;">
                Please update your calendar accordingly.
              </p>
            </td>
          </tr>

          <!-- Footer -->
          <tr>
            <td align="center" style="background:#203a43;padding:15px;border-top:1px solid rgba(245,245,245,0.08);font-size:12px;color:#b0b0b0;">
              © 2025 Basilur Tea Export. All rights reserved.
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>

</body>
</html>
`;
}

/**
 * Log successful email
 */
function logEmailSuccess(bookingId, meetingTitle, recipient, subject, status) {
  const emailLog = getOrCreateEmailLogSheet();
  emailLog.appendRow([
    formatDateAsIST(new Date(), 'yyyy-MM-dd HH:mm:ss'),
    bookingId,
    meetingTitle,
    recipient,
    subject,
    status,
    ''
  ]);
}

/**
 * Log email error
 */
function logEmailError(bookingId, meetingTitle, recipient, subject, errorMessage) {
  const emailLog = getOrCreateEmailLogSheet();
  emailLog.appendRow([
    formatDateAsIST(new Date(), 'yyyy-MM-dd HH:mm:ss'),
    bookingId,
    meetingTitle,
    recipient,
    subject,
    'Failed',
    errorMessage
  ]);
}

// ============================================================================
// UTILITY FUNCTIONS (Run Manually)
// ============================================================================

/**
 * Clean up old bookings (run manually or set up time-based trigger)
 */
function cleanupOldBookings() {
  const sheet = getOrCreateBookingsSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const rowsToDelete = [];
  const daysToKeep = 30;
  
  for (let i = 1; i < data.length; i++) {
    const endTime = new Date(data[i][5]);
    const timeDiff = now - endTime;
    const daysDiff = timeDiff / (1000 * 60 * 60 * 24);
    
    if (daysDiff > daysToKeep) {
      rowsToDelete.push(i + 1);
    }
  }
  
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }
  
  Logger.log('Cleaned up ' + rowsToDelete.length + ' bookings older than ' + daysToKeep + ' days');
}

/**
 * Get booking statistics
 */
function getBookingStats() {
  const sheet = getOrCreateBookingsSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  
  let totalBookings = data.length - 1;
  let upcomingBookings = 0;
  let runningBookings = 0;
  let pastBookings = 0;
  let emailsSent = 0;
  let roomStats = {};
  
  for (let i = 1; i < data.length; i++) {
    const start = new Date(data[i][4]);
    const end = new Date(data[i][5]);
    const room = data[i][1];
    const emailSent = data[i][9];
    
    if (now >= start && now < end) {
      runningBookings++;
    } else if (now < start) {
      upcomingBookings++;
    } else {
      pastBookings++;
    }
    
    if (emailSent) emailsSent++;
    
    roomStats[room] = (roomStats[room] || 0) + 1;
  }
  
  const stats = {
    total: totalBookings,
    upcoming: upcomingBookings,
    running: runningBookings,
    past: pastBookings,
    emailsSent: emailsSent,
    byRoom: roomStats,
    generatedAt: formatDateAsIST(now, 'yyyy-MM-dd HH:mm:ss')
  };
  
  Logger.log(JSON.stringify(stats, null, 2));
  return stats;
}

/**
 * Format date as IST (GMT+5:30) - Utility function
 */
function formatDateAsIST(dateInput, format) {
  format = format || 'MM/dd/yyyy HH:mm:ss';
  const IST_TIMEZONE = 'Asia/Kolkata';
  const date = dateInput instanceof Date ? dateInput : new Date(dateInput);
  return Utilities.formatDate(date, IST_TIMEZONE, format);
}

/**
 * Get current time in IST
 */
function getCurrentTimeIST() {
  return formatDateAsIST(new Date(), 'yyyy-MM-dd\'T\'HH:mm:ss');
}

/**
 * Test email functionality
 */
function testEmail() {
  const testBooking = {
    title: 'Test Meeting',
    room: 'Room A',
    start: new Date(Date.now() + 86400000).toISOString(),
    end: new Date(Date.now() + 90000000).toISOString(),
    bookedBy: 'Test User',
    note: 'This is a test email',
    participants: 'pipisara.design@gmail.com'  // Replace with your email
  };
  
  try {
    sendMeetingInvitation(testBooking, 'TEST123', false);
    Logger.log('Test email sent successfully!');
  } catch (error) {
    Logger.log('Test email failed: ' + error.toString());
  }
}

/**
 * Test cancellation email
 */
function testCancellationEmail() {
  const testBooking = {
    title: 'Test Meeting Cancellation',
    room: 'Room A',
    start: new Date(Date.now() + 86400000).toISOString(),
    end: new Date(Date.now() + 90000000).toISOString(),
    bookedBy: 'Test User',
    note: 'This is a test cancellation',
    participants: 'pipisara.design@gmail.com'  // Replace with your email
  };
  
  try {
    sendCancellationEmail(testBooking, 'TEST456', 'Admin User');
    Logger.log('Test cancellation email sent successfully!');
  } catch (error) {
    Logger.log('Test cancellation email failed: ' + error.toString());
  }
}

// ============================================================================
// GOOGLE CALENDAR INTEGRATION
// ============================================================================

/**
 * Create a Google Calendar Event
 * Requires "Google Calendar API" Advanced Service
 */
function createGoogleCalendarEvent(booking) {
  try {
    const calendarId = 'primary';
    const attendees = [];
    
    if (booking.participants) {
      const parts = booking.participants.split(',');
      parts.forEach(function(email) {
        const trimmed = email.trim();
        if (trimmed) {
          attendees.push({ email: trimmed });
        }
      });
    }

    const event = {
      summary: booking.title,
      location: booking.room,
      description: booking.note,
      start: {
        dateTime: booking.start,
        timeZone: 'Asia/Kolkata'
      },
      end: {
        dateTime: booking.end,
        timeZone: 'Asia/Kolkata'
      },
      attendees: attendees,
      reminders: {
        useDefault: true
      }
    };

    const newEvent = Calendar.Events.insert(event, calendarId, {
      sendUpdates: 'all'
    });
    
    Logger.log('Created Google Calendar Event: ' + newEvent.id);
    return newEvent.id;

  } catch (e) {
    Logger.log('Error creating calendar event: ' + e.toString());
    // We don't want to break the booking flow if calendar fails, so return null
    return null;
  }
}

/**
 * Update a Google Calendar Event
 */
function updateGoogleCalendarEvent(eventId, booking) {
  if (!eventId) return;

  try {
    const calendarId = 'primary';
    const attendees = [];
    
    if (booking.participants) {
      const parts = booking.participants.split(',');
      parts.forEach(function(email) {
        const trimmed = email.trim();
        if (trimmed) {
          attendees.push({ email: trimmed });
        }
      });
    }

    const event = {
      summary: booking.title,
      location: booking.room,
      description: booking.note,
      start: {
        dateTime: booking.start,
        timeZone: 'Asia/Kolkata'
      },
      end: {
        dateTime: booking.end,
        timeZone: 'Asia/Kolkata'
      },
      attendees: attendees
    };

    Calendar.Events.patch(event, calendarId, eventId, {
      sendUpdates: 'all'
    });
    
    Logger.log('Updated Google Calendar Event: ' + eventId);

  } catch (e) {
    Logger.log('Error updating calendar event: ' + e.toString());
  }
}

/**
 * Delete a Google Calendar Event
 */
function deleteGoogleCalendarEvent(eventId) {
  if (!eventId) return;

  try {
    const calendarId = 'primary';
    Calendar.Events.remove(calendarId, eventId, {
      sendUpdates: 'all'
    });
    Logger.log('Deleted Google Calendar Event: ' + eventId);
  } catch (e) {
    Logger.log('Error deleting calendar event: ' + e.toString());
  }
}

/**
 * Generate a Manual "Add to Calendar" Link
 */
function generateGoogleCalendarLink(booking) {
  // Format: https://calendar.google.com/calendar/render?action=TEMPLATE&text=...
  
  const formatDateForLink = function(isoStr) {
    // Converts ISO string to YYYYMMDDTHHMMSSZ
    var d = new Date(isoStr);
    return d.toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z';
  };

  const baseUrl = 'https://calendar.google.com/calendar/render?action=TEMPLATE';
  const text = '&text=' + encodeURIComponent(booking.title || 'Meeting');
  const details = '&details=' + encodeURIComponent(booking.note || '');
  const location = '&location=' + encodeURIComponent(booking.room || '');
  const dates = '&dates=' + formatDateForLink(booking.start) + '/' + formatDateForLink(booking.end);
  const add = '&add=' + encodeURIComponent(booking.participants || ''); // This pre-fills guests

  return baseUrl + text + dates + details + location + add;
}


// ============================================================================
// END OF SCRIPT
// ============================================================================
