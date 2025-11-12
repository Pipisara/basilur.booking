// ============================================================================
// Meeting Room Booking System - Google Apps Script Backend
// ============================================================================
// Deploy this as a Web App with "Anyone" access
// Version: 2.0
// Last Updated: 2024-11-12
// ============================================================================

/**
 * Handle GET requests - Retrieve all bookings
 * @param {Object} e - Event object from GET request
 * @returns {TextOutput} JSON array of all bookings
 */
function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action ? e.parameter.action : 'bookings';
    if (action === 'auth') {
      const name = e.parameter && e.parameter.name ? e.parameter.name : '';
      const accessCode = e.parameter && e.parameter.accessCode ? e.parameter.accessCode : '';
      const usersSheet = getOrCreateUsersSheet();
      const usersData = usersSheet.getDataRange().getValues();
      let found = null;
      for (let i = 1; i < usersData.length; i++) {
        const row = usersData[i];
        if (row[0] && String(row[0]).trim() === String(name).trim() && String(row[1]).trim() === String(accessCode).trim()) {
          found = { name: row[0], email: row[2] || '', role: row[3] || 'user' };
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
        createdBy: row[8] || '',
        updatedBy: row[9] || '',
        createdAt: row[10] || '',
        updatedAt: row[11] || ''
      });
    }
    return ContentService
      .createTextOutput(JSON.stringify(bookings))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        error: error.toString(),
        message: 'Failed to retrieve data'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handle POST requests - Create or update bookings
 * @param {Object} e - Event object from POST request
 * @returns {TextOutput} JSON response with status
 */
function doPost(e) {
  try {
    // Parse incoming booking data
    const booking = JSON.parse(e.postData.contents);
    const sheet = getOrCreateBookingsSheet();
    const action = booking.action || 'create';
    
    Logger.log(`doPost: Action = ${action}, ID = ${booking.id || 'new'}`);
    
    // Route to appropriate handler
    if (action === 'update') {
      return handleUpdate(sheet, booking);
    } else if (action === 'delete') {
      return handleDelete(sheet, booking);
    } else {
      return handleCreate(sheet, booking);
    }
    
  } catch (error) {
    Logger.log(`doPost Error: ${error.toString()}`);
    
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

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
 * Get existing Bookings sheet or create new one with headers
 * @returns {Sheet} The Bookings sheet
 */
function getOrCreateBookingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Bookings');
  
  // Create sheet if it doesn't exist
  if (!sheet) {
    Logger.log('Creating new Bookings sheet with headers');
    
    sheet = ss.insertSheet('Bookings');
    
    // Add header row
    const headers = [
      'ID',
      'Room',
      'RoomKey',
      'Title',
      'Start',
      'End',
      'Booked By',
      'Note',
      'Created By',
      'Updated By',
      'Created At',
      'Updated At'
    ];
    
    sheet.appendRow(headers);
    
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    
    // Set column widths
    sheet.setColumnWidth(1, 120);  // ID
    sheet.setColumnWidth(2, 80);   // Room
    sheet.setColumnWidth(3, 60);   // RoomKey
    sheet.setColumnWidth(4, 200);  // Title
    sheet.setColumnWidth(5, 150);  // Start
    sheet.setColumnWidth(6, 150);  // End
    sheet.setColumnWidth(7, 150);  // Booked By
    sheet.setColumnWidth(8, 300);  // Note
    sheet.setColumnWidth(9, 120);  // Created By
    sheet.setColumnWidth(10, 120); // Updated By
    sheet.setColumnWidth(11, 150); // Created At
    sheet.setColumnWidth(12, 150); // Updated At
    
    // Freeze header row
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * Create a new booking
 * @param {Sheet} sheet - The Bookings sheet
 * @param {Object} booking - Booking data
 * @returns {TextOutput} JSON response
 */
function handleCreate(sheet, booking) {
  const now = new Date();
  const bookingId = now.getTime().toString();
  
  Logger.log(`Creating booking: ${bookingId} - ${booking.title}`);
  
  // Append new row with booking data
  sheet.appendRow([
    bookingId,                                    // ID
    booking.room || '',                           // Room
    booking.roomKey || '',                        // RoomKey
    booking.title || 'Untitled Meeting',          // Title
    booking.start || '',                          // Start
    booking.end || '',                            // End
    booking.bookedBy || '',                       // Booked By
    booking.note || '',                           // Note
    booking.createdBy || booking.bookedBy || '',  // Created By
    '',                                           // Updated By (empty for new)
    now.toISOString(),                            // Created At
    ''                                            // Updated At (empty for new)
  ]);
  
  Logger.log(`Booking created successfully: ${bookingId}`);
  
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'success',
      action: 'create',
      id: bookingId,
      message: 'Booking created successfully'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Update an existing booking
 * @param {Sheet} sheet - The Bookings sheet
 * @param {Object} booking - Updated booking data
 * @returns {TextOutput} JSON response
 */
function handleUpdate(sheet, booking) {
  const bookingId = booking.id || booking.originalId;
  
  if (!bookingId) {
    Logger.log('Update failed: No booking ID provided');
    
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: 'Booking ID is required for updates'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  
  // Find the row with matching ID
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === bookingId.toString()) {
      rowIndex = i + 1; // +1 because sheet rows are 1-indexed
      break;
    }
  }
  
  if (rowIndex === -1) {
    Logger.log(`Update failed: Booking ${bookingId} not found`);
    
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: 'Booking not found'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  Logger.log(`Updating booking: ${bookingId} at row ${rowIndex}`);
  
  const now = new Date();
  const existingCreatedAt = data[rowIndex - 1][10];
  const existingCreatedBy = data[rowIndex - 1][8];
  
  // Update the entire row
  sheet.getRange(rowIndex, 1, 1, 12).setValues([[
    bookingId,                                    // ID (keep same)
    booking.room || '',                           // Room
    booking.roomKey || '',                        // RoomKey
    booking.title || 'Untitled Meeting',          // Title
    booking.start || '',                          // Start
    booking.end || '',                            // End
    booking.bookedBy || '',                       // Booked By
    booking.note || '',                           // Note
    existingCreatedBy || '',                      // Created By (preserve)
    booking.updatedBy || booking.submittedBy || '', // Updated By
    existingCreatedAt || '',                      // Created At (preserve)
    now.toISOString()                             // Updated At
  ]]);
  
  Logger.log(`Booking updated successfully: ${bookingId}`);
  
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
 * @param {Sheet} sheet - The Bookings sheet
 * @param {Object} booking - Booking data with ID
 * @returns {TextOutput} JSON response
 */
function handleDelete(sheet, booking) {
  const bookingId = booking.id;
  
  if (!bookingId) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: 'Booking ID is required for deletion'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  
  // Find the row with matching ID
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === bookingId.toString()) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: 'Booking not found'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  Logger.log(`Deleting booking: ${bookingId} at row ${rowIndex}`);
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

// ============================================================================
// UTILITY FUNCTIONS (Optional - Run Manually)
// ============================================================================

/**
 * Clean up old bookings (run manually or set up time-based trigger)
 * Deletes bookings that ended more than 30 days ago
 */
function cleanupOldBookings() {
  const sheet = getOrCreateBookingsSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const rowsToDelete = [];
  const daysToKeep = 30;
  
  // Find rows with end times older than specified days
  for (let i = 1; i < data.length; i++) {
    const endTime = new Date(data[i][5]); // End time column
    const timeDiff = now - endTime;
    const daysDiff = timeDiff / (1000 * 60 * 60 * 24);
    
    if (daysDiff > daysToKeep) {
      rowsToDelete.push(i + 1);
    }
  }
  
  // Delete rows in reverse order to maintain correct indices
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }
  
  Logger.log(`Cleaned up ${rowsToDelete.length} bookings older than ${daysToKeep} days`);
}

/**
 * Get booking statistics (run manually for insights)
 */
function getBookingStats() {
  const sheet = getOrCreateBookingsSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  
  let totalBookings = data.length - 1; // Exclude header
  let upcomingBookings = 0;
  let runningBookings = 0;
  let pastBookings = 0;
  let roomStats = {};
  
  for (let i = 1; i < data.length; i++) {
    const start = new Date(data[i][4]);
    const end = new Date(data[i][5]);
    const room = data[i][1];
    
    // Count by status
    if (now >= start && now < end) {
      runningBookings++;
    } else if (now < start) {
      upcomingBookings++;
    } else {
      pastBookings++;
    }
    
    // Count by room
    roomStats[room] = (roomStats[room] || 0) + 1;
  }
  
  const stats = {
    total: totalBookings,
    upcoming: upcomingBookings,
    running: runningBookings,
    past: pastBookings,
    byRoom: roomStats,
    generatedAt: now.toISOString()
  };
  
  Logger.log(JSON.stringify(stats, null, 2));
  return stats;
}

/**
 * Test function - Creates sample bookings for testing
 */
function createSampleBookings() {
  const sheet = getOrCreateBookingsSheet();
  const now = new Date();
  
  const samples = [
    {
      room: 'Room A',
      roomKey: 'A',
      title: 'Daily Standup',
      start: new Date(now.getTime() + 1 * 60 * 60 * 1000).toISOString(),
      end: new Date(now.getTime() + 2 * 60 * 60 * 1000).toISOString(),
      bookedBy: 'Admin',
      note: 'Quick team sync',
      createdBy: 'Admin'
    },
    {
      room: 'Room B',
      roomKey: 'B',
      title: 'Client Presentation',
      start: new Date(now.getTime() + 3 * 60 * 60 * 1000).toISOString(),
      end: new Date(now.getTime() + 4 * 60 * 60 * 1000).toISOString(),
      bookedBy: 'Manager',
      note: 'Q4 review with stakeholders',
      createdBy: 'Manager'
    },
    {
      room: 'Room C',
      roomKey: 'C',
      title: 'Training Session',
      start: new Date(now.getTime() + 24 * 60 * 60 * 1000).toISOString(),
      end: new Date(now.getTime() + 26 * 60 * 60 * 1000).toISOString(),
      bookedBy: 'Pipisara Chandrabhanu',
      note: 'New employee onboarding',
      createdBy: 'Pipisara Chandrabhanu'
    }
  ];
  
  samples.forEach(sample => {
    handleCreate(sheet, sample);
  });
  
  Logger.log(`Created ${samples.length} sample bookings`);
}

// ============================================================================
// END OF SCRIPT
// ============================================================================
