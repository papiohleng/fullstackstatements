/**
 * ============================================
 * Code.gs - MABON SUITES HOTEL PLATFORM
 * ============================================
 * 
 * Full Stack Google App Script Application
 * for Mabon Suites Hotel - Luxury Accommodations
 * 
 * Features:
 * - Guest booking form with real-time pricing
 * - Admin dashboard for managing reservations
 * - PDF booking confirmation generation
 * - Client billing & statement system
 * - Visa document management
 * - Multi-room type support
 * ============================================
 */

const CONFIG = {
  SPREADSHEET_ID: '1SXUFidabeo_9GE3Nh-FwNKmmaMrlcyoziI0wtDL9yK4',
  DRIVE_FOLDER_ID: '1EO8gDYZhZEMctWTN8pQV1g9nnjopqj0E',
  BOOKINGS_SHEET: 'Bookings',
  DOCUMENTS_SHEET: 'Documents',
  COMMENTS_SHEET: 'Comments',
  TRANSACTIONS_SHEET: 'Transactions',
  DELETED_SHEET: 'Deleted',
  ROOMS_SHEET: 'Rooms',
  APP_NAME: 'Mabon Suites',
  COMPANY_EMAIL: 'reservations@mabonsuites.com',
  COMPANY_PHONE: '+27 11 XXX XXXX',
  COMPANY_ADDRESS: 'Johannesburg, South Africa',
  TOURIST_LEVY: 100,
  VISA_FEE_PER_FILE: 500,
  MAX_VISA_FEE: 1500
};

/* =========================
   ENTRY POINTS
========================= */

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'form';

  if (action === 'form' || action === 'booking')
    return HtmlService.createHtmlOutputFromFile('Form')
      .setTitle('Book Your Stay | ' + CONFIG.APP_NAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  if (action === 'admin' || action === 'dashboard')
    return HtmlService.createHtmlOutputFromFile('Submissions')
      .setTitle('Admin Dashboard | ' + CONFIG.APP_NAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  if (action === 'deleted')
    return HtmlService.createHtmlOutputFromFile('Deleted')
      .setTitle('Deleted Bookings | ' + CONFIG.APP_NAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  if (action === 'attachments')
    return HtmlService.createHtmlOutputFromFile('Attachments')
      .setTitle('Documents Manager | ' + CONFIG.APP_NAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  if (action === 'get-bookings')
    return jsonResponse(getAllBookings());

  if (action === 'get-booking')
    return jsonResponse(getBookingByReference(e.parameter.reference));

  if (action === 'get-comments')
    return jsonResponse({ comments: getCommentsForBooking(e.parameter.reference) });

  if (action === 'get-deleted')
    return jsonResponse(getDeletedBookings());

  if (action === 'get-stats')
    return jsonResponse(getStats());

  if (action === 'get-deleted-count')
    return jsonResponse(getDeletedBookingsCount());

  if (action === 'get-bookings-with-comments')
    return jsonResponse(getAllBookingsWithComments());

  if (action === 'get-rooms')
    return jsonResponse(getRoomsFromSheet());

  if (action === 'get-transactions')
    return jsonResponse(getTransactionsByReference(e.parameter.reference));

  if (action === 'get-statement-pdf')
    return generateStatementPDFResponse(e.parameter.reference);

  return jsonResponse({ error: 'Unknown action' });
}

function doPost(e) {
  const data = e.postData?.contents ? JSON.parse(e.postData.contents) : {};
  const action = data.action;

  if (action === 'submit') return jsonResponse(submitBooking(data));
  if (action === 'update-field') return jsonResponse(updateField(data));
  if (action === 'add-comment') return jsonResponse(addCommentInternal(data));
  if (action === 'send-email') return jsonResponse(sendEmailInternal(data));
  if (action === 'delete') return jsonResponse(moveToDeleted(data));
  if (action === 'restore') return jsonResponse(restoreFromDeleted(data));
  if (action === 'permanent-delete') return jsonResponse(permanentDelete(data));
  if (action === 'send-confirmation') return jsonResponse(sendConfirmationEmail(data));
  if (action === 'add-transaction') return jsonResponse(addTransaction(data));
  if (action === 'delete-transaction') return jsonResponse(deleteTransaction(data));
  if (action === 'send-statement') return jsonResponse(sendStatementEmail(data));

  return jsonResponse({ error: 'Unknown action' });
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify({
    success: !data?.error,
    ...data
  })).setMimeType(ContentService.MimeType.JSON);
}

/* =========================
   BOOKING SUBMISSION
========================= */

function submitBooking(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    if (!data.email)
      return { error: 'Email address is required' };

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const bookSheet = getOrCreateBookingsSheet(ss);
    const docSheet = getOrCreateDocumentsSheet(ss);

    const reference = data.reference || generateReference(data.guestName);
    const timestamp = new Date().toISOString();

    // Calculate fees
    const nights = parseInt(data.nights) || 1;
    const roomRate = parseFloat(data.roomRate) || 0;
    const roomTotal = roomRate * nights;
    const touristLevy = CONFIG.TOURIST_LEVY;
    const visaFileCount = data.visaFileCount || 0;
    const visaFee = Math.min(visaFileCount * CONFIG.VISA_FEE_PER_FILE, CONFIG.MAX_VISA_FEE);
    const totalAmount = roomTotal + touristLevy + visaFee;

    // Append booking row (24 columns)
    bookSheet.appendRow([
      reference,                              // 0  Reference
      timestamp,                              // 1  Timestamp
      'new',                                  // 2  Status
      sanitize(data.guestName || ''),         // 3  Guest_Name
      sanitize(data.email),                   // 4  Email
      sanitize(data.phone || ''),             // 5  Phone
      sanitize(data.country || 'South Africa'), // 6  Country
      sanitize(data.checkInDate || ''),       // 7  Check_In_Date
      sanitize(data.checkOutDate || ''),      // 8  Check_Out_Date
      nights,                                 // 9  Nights
      sanitize(data.roomType || 'deluxe'),    // 10 Room_Type
      sanitize(data.roomName || 'Deluxe Suite'), // 11 Room_Name
      roomRate,                               // 12 Room_Rate
      roomTotal,                              // 13 Room_Total
      touristLevy,                            // 14 Tourist_Levy
      visaFee,                                // 15 Visa_Fee
      totalAmount,                            // 16 Total_Amount
      sanitize(data.specialRequests || ''),   // 17 Special_Requests
      data.requestOptions || '[]',            // 18 Request_Options (JSON)
      'pending',                              // 19 Payment_Status
      data.visaRequired || false,             // 20 Visa_Required
      false,                                  // 21 Confirmation_Sent
      false,                                  // 22 Guest_Arrived
      timestamp                               // 23 Last_Updated
    ]);

    // Upload visa documents
    uploadDocuments(reference, data, docSheet);

    // Add system comment
    addCommentInternal({ reference, author: 'SYSTEM', text: 'Booking submitted' });

    // Send confirmation email with PDF
    sendBookingConfirmation(data, reference, {
      nights, roomRate, roomTotal, touristLevy, visaFee, totalAmount
    });

    return { success: true, reference };
  } catch(e) {
    Logger.log('submitBooking error: ' + e.message);
    return { error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// Called from Form.html via google.script.run
function processFormSubmission(formData) {
  const result = submitBooking(formData);
  if (result.error) throw new Error(result.error);
  return { success: true, reference: result.reference };
}

/* =========================
   DOCUMENTS
========================= */

function uploadDocuments(reference, data, docSheet) {
  try {
    const parent = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    let folder;
    
    // Try to find existing folder or create new one
    const folders = parent.getFoldersByName(reference);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = parent.createFolder(reference);
    }

    // Upload visa/passport files
    const fileKeys = ['visaFile1', 'visaFile2', 'visaFile3', 'passportFile'];
    const fileTypes = ['Visa Document 1', 'Visa Document 2', 'Visa Document 3', 'Passport'];
    
    fileKeys.forEach((key, idx) => {
      const f = data[key];
      if (!f?.data) return;

      const file = uploadFile(folder, f, fileTypes[idx]);
      docSheet.appendRow([
        generateUUID(),
        reference,
        fileTypes[idx],
        f.name,
        file.getUrl(),
        data[key + 'Comment'] || '',
        new Date().toISOString()
      ]);
    });
  } catch (e) {
    Logger.log('Upload error: ' + e.message);
  }
}

function uploadFile(folder, fileData, prefix) {
  const blob = Utilities.newBlob(
    Utilities.base64Decode(fileData.data),
    fileData.type || 'application/octet-stream',
    prefix + '_' + fileData.name
  );
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file;
}

/* =========================
   READ OPERATIONS
========================= */

function getBookingByReference(reference) {
  if (!reference) return null;

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const bookings = ss.getSheetByName(CONFIG.BOOKINGS_SHEET).getDataRange().getValues();
  const docs = getDocumentsMap()[reference] || [];

  for (let i = 1; i < bookings.length; i++) {
    if (bookings[i][0] === reference) {
      return {
        reference,
        timestamp: bookings[i][1],
        status: bookings[i][2],
        guestName: bookings[i][3],
        email: bookings[i][4],
        phone: bookings[i][5],
        country: bookings[i][6],
        checkInDate: bookings[i][7],
        checkOutDate: bookings[i][8],
        nights: bookings[i][9],
        roomType: bookings[i][10],
        roomName: bookings[i][11],
        roomRate: bookings[i][12],
        roomTotal: bookings[i][13],
        touristLevy: bookings[i][14],
        visaFee: bookings[i][15],
        totalAmount: bookings[i][16],
        specialRequests: bookings[i][17],
        requestOptions: bookings[i][18],
        paymentStatus: bookings[i][19],
        visaRequired: bookings[i][20],
        confirmationSent: bookings[i][21],
        guestArrived: bookings[i][22],
        lastUpdated: bookings[i][23],
        documents: docs,
        comments: getCommentsForBooking(reference)
      };
    }
  }
  return null;
}

function getAllBookings() {
  try {
    Logger.log('getAllBookings: Starting...');
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.BOOKINGS_SHEET);
    
    if (!sheet) {
      Logger.log('getAllBookings: Bookings sheet not found - creating it');
      // Try to create the sheet if it doesn't exist
      const newSheet = getOrCreateBookingsSheet(ss);
      return { bookings: [] };
    }
    
    const lastRow = sheet.getLastRow();
    Logger.log('getAllBookings: Sheet has ' + lastRow + ' rows total');
    
    if (lastRow < 2) {
      Logger.log('getAllBookings: No data rows (only header or empty)');
      return { bookings: [] };
    }
    
    const rows = sheet.getDataRange().getValues();
    Logger.log('getAllBookings: Read ' + rows.length + ' rows');
    
    // Log first data row for debugging
    if (rows.length > 1) {
      Logger.log('getAllBookings: First data row reference: ' + rows[1][0]);
    }
    
    var docMap = {};
    try {
      docMap = getDocumentsMap();
    } catch(e) {
      Logger.log('getAllBookings: Error getting documents map: ' + e.message);
    }
    
    const bookings = [];
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[0]) {
        Logger.log('getAllBookings: Skipping empty row at index ' + i);
        continue;
      }
      
      var docs = docMap[r[0]] || [];
      
      // Ensure status has a default value
      var status = r[2] ? String(r[2]).toLowerCase().trim() : 'new';
      
      // Handle date conversion safely
      var timestamp = '';
      if (r[1]) {
        if (r[1] instanceof Date) {
          timestamp = r[1].toISOString();
        } else {
          timestamp = String(r[1]);
        }
      } else {
        timestamp = new Date().toISOString();
      }
      
      var checkIn = r[7];
      var checkOut = r[8];
      if (checkIn instanceof Date) checkIn = checkIn.toISOString().split('T')[0];
      if (checkOut instanceof Date) checkOut = checkOut.toISOString().split('T')[0];
      
      bookings.push({
        reference: String(r[0] || ''),
        timestamp: timestamp,
        status: status,
        guestName: String(r[3] || ''),
        email: String(r[4] || ''),
        phone: String(r[5] || ''),
        country: String(r[6] || 'South Africa'),
        checkInDate: String(checkIn || ''),
        checkOutDate: String(checkOut || ''),
        nights: parseInt(r[9]) || 0,
        roomType: String(r[10] || 'deluxe'),
        roomName: String(r[11] || ''),
        roomRate: parseFloat(r[12]) || 0,
        roomTotal: parseFloat(r[13]) || 0,
        touristLevy: parseFloat(r[14]) || 0,
        visaFee: parseFloat(r[15]) || 0,
        totalAmount: parseFloat(r[16]) || 0,
        paymentStatus: r[19] ? String(r[19]).toLowerCase() : 'pending',
        confirmationSent: r[21] === true || r[21] === 'TRUE' || r[21] === 'true',
        guestArrived: r[22] === true || r[22] === 'TRUE' || r[22] === 'true',
        lastUpdated: String(r[23] || ''),
        documents: docs
      });
    }

    Logger.log('getAllBookings: Returning ' + bookings.length + ' bookings');
    return { bookings: bookings };
    
  } catch(e) {
    Logger.log('getAllBookings ERROR: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    return { bookings: [], error: e.message };
  }
}

function getAllBookingsWithComments() {
  try {
    Logger.log('getAllBookingsWithComments: Starting...');
    
    var result = getAllBookings();
    var bookings = result.bookings || [];
    
    Logger.log('getAllBookingsWithComments: Got ' + bookings.length + ' bookings from getAllBookings');
    
    if (result.error) {
      Logger.log('getAllBookingsWithComments: getAllBookings returned error: ' + result.error);
      return { bookings: [], error: result.error };
    }
    
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var commentsSheet = ss.getSheetByName(CONFIG.COMMENTS_SHEET);
    
    // Build comment count map
    var commentCounts = {};
    if (commentsSheet && commentsSheet.getLastRow() > 1) {
      var rows = commentsSheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        var ref = rows[i][1];
        if (ref) {
          commentCounts[ref] = (commentCounts[ref] || 0) + 1;
        }
      }
    }
    
    // Add comment count to each booking
    for (var j = 0; j < bookings.length; j++) {
      bookings[j].commentCount = commentCounts[bookings[j].reference] || 0;
    }
    
    Logger.log('getAllBookingsWithComments: Returning ' + bookings.length + ' bookings with comments');
    return { bookings: bookings };
    
  } catch(e) {
    Logger.log('getAllBookingsWithComments ERROR: ' + e.message);
    return { bookings: [], error: e.message };
  }
}

function getRoomsFromSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.ROOMS_SHEET);
  
  if (!sheet) return { rooms: [] };
  
  const rows = sheet.getDataRange().getValues();
  const rooms = [];
  
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] || r[8] === false) continue; // Skip inactive rooms
    
    let features = [];
    try {
      features = JSON.parse(r[6] || '[]');
    } catch(e) {
      features = [];
    }
    
    rooms.push({
      id: r[0],
      type: r[1],
      name: r[2],
      price: r[3],
      size: r[4],
      description: r[5],
      features: features,
      maxGuests: r[7]
    });
  }
  
  return { rooms };
}

function getStats() {
  const { bookings } = getAllBookings();
  
  const stats = {
    total: bookings.length,
    new: 0,
    pending: 0,
    confirmed: 0,
    checked_in: 0,
    checked_out: 0,
    cancelled: 0,
    totalRevenue: 0,
    paidRevenue: 0
  };
  
  bookings.forEach(b => {
    const status = (b.status || 'new').toLowerCase();
    if (stats.hasOwnProperty(status)) {
      stats[status]++;
    }
    stats.totalRevenue += b.totalAmount || 0;
    if (b.paymentStatus === 'paid') {
      stats.paidRevenue += b.totalAmount || 0;
    }
  });
  
  return stats;
}

/* =========================
   COMMENTS
========================= */

function addCommentInternal({ reference, author, text }) {
  if (!reference || !text) return { error: 'Invalid comment' };

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  getOrCreateCommentsSheet(ss).appendRow([
    generateUUID(), reference, author || 'Concierge',
    sanitize(text), new Date().toISOString()
  ]);

  return { success: true };
}

function addComment(reference, text, author) {
  const result = addCommentInternal({ reference, author: author || 'Concierge', text });
  if (result.error) throw new Error(result.error);
  return result;
}

function getCommentsForBooking(reference) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.COMMENTS_SHEET);
  if (!sheet) return [];

  return sheet.getDataRange().getValues()
    .slice(1)
    .filter(r => r[1] === reference)
    .map(r => ({ id: r[0], author: r[2], text: r[3], timestamp: r[4] }))
    .reverse();
}

/* =========================
   UPDATE FIELD
========================= */

function updateField(data) {
  const { reference, field, value } = data;
  if (!reference || !field) return { error: 'Missing reference or field' };

  // Column indices are 1-based for getRange (Column A = 1)
  const fieldMap = {
    status: 3,
    guestName: 4,
    email: 5,
    phone: 6,
    paymentStatus: 20,
    confirmationSent: 22,
    guestArrived: 23
  };

  const col = fieldMap[field];
  if (!col) return { error: 'Invalid field' };

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.BOOKINGS_SHEET);
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === reference) {
      sheet.getRange(i + 1, col).setValue(value);
      sheet.getRange(i + 1, 24).setValue(new Date().toISOString()); // Update Last_Updated
      
      if (field === 'status') {
        addCommentInternal({ reference, author: 'Concierge', text: 'Status changed to: ' + value });
      }
      if (field === 'paymentStatus') {
        addCommentInternal({ reference, author: 'Concierge', text: 'Payment status changed to: ' + value });
      }
      if (field === 'guestArrived' && value === true) {
        addCommentInternal({ reference, author: 'SYSTEM', text: 'Guest checked in' });
      }
      return { success: true };
    }
  }
  return { error: 'Not found' };
}

function updateBookingField(reference, field, value) {
  const result = updateField({ reference, field, value });
  if (result.error) throw new Error(result.error);
  return result;
}

/* =========================
   DELETE / RESTORE
========================= */

function moveToDeleted({ reference, deleteReason, deletedBy }) {
  if (!reference) return { error: 'Reference required' };
  if (!deleteReason) return { error: 'Delete reason required' };

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const bookSheet = ss.getSheetByName(CONFIG.BOOKINGS_SHEET);
  const delSheet = getOrCreateDeletedSheet(ss);

  const data = bookSheet.getDataRange().getValues();
  const idx = data.findIndex(r => r[0] === reference);
  if (idx < 1) return { error: 'Not found' };

  delSheet.appendRow([
    reference,
    new Date().toISOString(),
    deleteReason || '',
    deletedBy || 'Admin',
    JSON.stringify(data[idx])
  ]);

  bookSheet.deleteRow(idx + 1);
  addCommentInternal({ reference, author: 'SYSTEM', text: 'Booking cancelled: ' + deleteReason });

  return { success: true };
}

function deleteBooking(reference, reason, deletedBy) {
  const result = moveToDeleted({ reference, deleteReason: reason, deletedBy });
  if (result.error) throw new Error(result.error);
  return result;
}

function getDeletedBookings() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.DELETED_SHEET);
  if (!sheet) return { deleted: [] };

  const rows = sheet.getDataRange().getValues();
  const deleted = [];
  
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    
    let snapshot = {};
    try { snapshot = JSON.parse(r[4] || '[]'); } catch(e) {}
    
    deleted.push({
      reference: r[0],
      deletedAt: r[1],
      deleteReason: r[2],
      deletedBy: r[3],
      guestName: snapshot[3] || '',
      email: snapshot[4] || '',
      roomName: snapshot[11] || '',
      checkInDate: snapshot[7] || '',
      checkOutDate: snapshot[8] || '',
      totalAmount: snapshot[16] || 0,
      snapshot: r[4]
    });
  }

  return { deleted };
}

function restoreFromDeleted({ reference }) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const delSheet = ss.getSheetByName(CONFIG.DELETED_SHEET);
  const bookSheet = getOrCreateBookingsSheet(ss);

  const rows = delSheet.getDataRange().getValues();
  const idx = rows.findIndex(r => r[0] === reference);
  if (idx < 1) return { error: 'Not found' };

  const snapshot = JSON.parse(rows[idx][4]);
  snapshot[2] = 'pending'; // Reset status
  snapshot[23] = new Date().toISOString(); // Update Last_Updated

  bookSheet.appendRow(snapshot);
  delSheet.deleteRow(idx + 1);
  addCommentInternal({ reference, author: 'SYSTEM', text: 'Booking restored' });

  return { success: true };
}

function restoreBooking(reference) {
  const result = restoreFromDeleted({ reference });
  if (result.error) throw new Error(result.error);
  return result;
}

function permanentDelete({ reference }) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const delSheet = ss.getSheetByName(CONFIG.DELETED_SHEET);
  const rows = delSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === reference) {
      delSheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'Not found' };
}

function permanentDeleteBooking(reference) {
  const result = permanentDelete({ reference });
  if (result.error) throw new Error(result.error);
  return result;
}

function getDeletedBookingsCount() {
  try {
    var result = getDeletedBookings();
    var deleted = result.deleted || [];
    return { deletedCount: deleted.length };
  } catch(e) {
    Logger.log('getDeletedBookingsCount ERROR: ' + e.message);
    return { deletedCount: 0, error: e.message };
  }
}

/* =========================
   EMAIL - BOOKING CONFIRMATION WITH PDF
========================= */

function sendBookingConfirmation(data, reference, pricing) {
  Logger.log('[sendBookingConfirmation] Starting for reference: ' + reference);
  Logger.log('[sendBookingConfirmation] Email: ' + data.email);
  
  try {
    const { guestName, email, phone, roomName, checkInDate, checkOutDate } = data;
    const { nights, roomRate, roomTotal, touristLevy, visaFee, totalAmount } = pricing;
    
    if (!email) {
      Logger.log('[sendBookingConfirmation] No email provided, skipping');
      return false;
    }
    
    const year = new Date().getFullYear();
    
    // Parse request options
    let requestOptions = [];
    try {
      requestOptions = JSON.parse(data.requestOptions || '[]');
    } catch(e) {
      requestOptions = [];
    }

    const subject = `Booking Confirmation - ${reference} | ${CONFIG.APP_NAME}`;

    // Plain text version
    const plainText = `
Dear ${guestName || 'Valued Guest'},

Thank you for choosing ${CONFIG.APP_NAME}!

Your Booking Reference: ${reference}

BOOKING DETAILS:
Room: ${roomName || 'Deluxe Suite'}
Check-in: ${checkInDate || 'TBD'}
Check-out: ${checkOutDate || 'TBD'}
Duration: ${nights} night(s)

PRICING:
Room Rate: R ${roomRate.toLocaleString()} x ${nights} nights = R ${roomTotal.toLocaleString()}
Tourist Levy: R ${touristLevy.toLocaleString()}
${visaFee > 0 ? `Visa Processing: R ${visaFee.toLocaleString()}` : ''}
Total: R ${totalAmount.toLocaleString()}

${requestOptions.length > 0 ? `Special Requests: ${requestOptions.join(', ')}` : ''}

A PDF confirmation with full details has been attached to this email.

We look forward to welcoming you!

Best regards,
${CONFIG.APP_NAME} Team
${CONFIG.COMPANY_ADDRESS}
${CONFIG.COMPANY_PHONE}
    `.trim();

    // HTML version
    const html = `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Booking Confirmation</title>
</head>
<body style="margin:0; padding:0; background-color:#0a0a0a; font-family:'Segoe UI',system-ui,-apple-system,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="padding:24px 16px;">
    <tr>
      <td align="center">
        <table width="100%" cellpadding="0" cellspacing="0" 
               style="max-width:600px; background:#141414; border-radius:16px; overflow:hidden; 
                      border:1px solid #2a2a2a;">
          
          <!-- Header -->
          <tr>
            <td style="background:linear-gradient(135deg, #c9a962 0%, #a88a4a 100%); padding:32px 24px; text-align:center;">
              <h1 style="margin:0; color:#0a0a0a; font-size:28px; font-weight:700; font-family:'Playfair Display',Georgia,serif;">
                ${CONFIG.APP_NAME}
              </h1>
              <p style="margin:8px 0 0; color:rgba(0,0,0,0.7); font-size:14px; letter-spacing:2px; text-transform:uppercase;">
                Booking Confirmation
              </p>
            </td>
          </tr>

          <!-- Content -->
          <tr>
            <td style="padding:36px 32px; color:#faf9f6;">
              <p style="font-size:16px; margin:0 0 20px;">
                Dear <strong>${escapeHtml(guestName || 'Valued Guest')}</strong>,
              </p>

              <p style="font-size:15px; line-height:1.6; margin:0 0 24px; color:#a3a3a3;">
                Thank you for choosing ${CONFIG.APP_NAME}. We are delighted to confirm your reservation.
              </p>

              <!-- Reference box -->
              <table width="100%" cellpadding="0" cellspacing="0" 
                     style="background:rgba(201,169,98,0.1); border:1px solid rgba(201,169,98,0.3); border-radius:12px; margin:28px 0; padding:24px; text-align:center;">
                <tr>
                  <td>
                    <div style="font-size:14px; color:#a3a3a3; margin-bottom:8px;">
                      Your Booking Reference
                    </div>
                    <div style="font-size:24px; font-weight:700; color:#c9a962; letter-spacing:2px;">
                      ${reference}
                    </div>
                  </td>
                </tr>
              </table>

              <!-- Booking Details -->
              <table width="100%" cellpadding="0" cellspacing="0" style="margin:24px 0;">
                <tr>
                  <td style="font-size:14px; color:#a3a3a3; padding-bottom:12px; border-bottom:1px solid #2a2a2a;">
                    <strong style="color:#c9a962;">Booking Details</strong>
                  </td>
                </tr>
                <tr>
                  <td style="padding:16px 0;">
                    <table width="100%" cellpadding="0" cellspacing="0">
                      <tr>
                        <td style="padding:8px 0; color:#a3a3a3; font-size:14px;">Room:</td>
                        <td style="padding:8px 0; color:#faf9f6; font-size:14px; text-align:right;">${roomName || 'Deluxe Suite'}</td>
                      </tr>
                      <tr>
                        <td style="padding:8px 0; color:#a3a3a3; font-size:14px;">Check-in:</td>
                        <td style="padding:8px 0; color:#faf9f6; font-size:14px; text-align:right;">${checkInDate || 'TBD'}</td>
                      </tr>
                      <tr>
                        <td style="padding:8px 0; color:#a3a3a3; font-size:14px;">Check-out:</td>
                        <td style="padding:8px 0; color:#faf9f6; font-size:14px; text-align:right;">${checkOutDate || 'TBD'}</td>
                      </tr>
                      <tr>
                        <td style="padding:8px 0; color:#a3a3a3; font-size:14px;">Duration:</td>
                        <td style="padding:8px 0; color:#faf9f6; font-size:14px; text-align:right;">${nights} night(s)</td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>

              <!-- Total -->
              <table width="100%" cellpadding="0" cellspacing="0" 
                     style="background:rgba(201,169,98,0.1); border-radius:12px; padding:20px; margin:24px 0;">
                <tr>
                  <td>
                    <table width="100%" cellpadding="0" cellspacing="0">
                      <tr>
                        <td style="color:#a3a3a3; font-size:14px; padding:6px 0;">Room (${nights} nights):</td>
                        <td style="text-align:right; color:#faf9f6; font-size:14px; padding:6px 0;">R ${roomTotal.toLocaleString()}</td>
                      </tr>
                      <tr>
                        <td style="color:#a3a3a3; font-size:14px; padding:6px 0;">Tourist Levy:</td>
                        <td style="text-align:right; color:#faf9f6; font-size:14px; padding:6px 0;">R ${touristLevy.toLocaleString()}</td>
                      </tr>
                      ${visaFee > 0 ? `
                      <tr>
                        <td style="color:#a3a3a3; font-size:14px; padding:6px 0;">Visa Processing:</td>
                        <td style="text-align:right; color:#faf9f6; font-size:14px; padding:6px 0;">R ${visaFee.toLocaleString()}</td>
                      </tr>
                      ` : ''}
                      <tr>
                        <td style="color:#faf9f6; font-size:16px; font-weight:700; padding:12px 0 0; border-top:1px solid #2a2a2a;">Total Amount:</td>
                        <td style="text-align:right; color:#c9a962; font-size:22px; font-weight:700; padding:12px 0 0; border-top:1px solid #2a2a2a;">R ${totalAmount.toLocaleString()}</td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>

              <p style="font-size:14px; line-height:1.6; margin:24px 0; color:#a3a3a3;">
                A detailed PDF confirmation is attached to this email for your records.
              </p>

              <p style="margin-top:32px; font-size:15px; color:#faf9f6;">
                We look forward to welcoming you,<br>
                <strong style="color:#c9a962;">${CONFIG.APP_NAME} Team</strong>
              </p>
            </td>
          </tr>

          <!-- Footer -->
          <tr>
            <td style="background:#1f1f1f; padding:20px; text-align:center; 
                       font-size:12px; color:#a3a3a3; border-top:1px solid #2a2a2a;">
              <p style="margin:0;">${CONFIG.APP_NAME} | ${CONFIG.COMPANY_ADDRESS}</p>
              <p style="margin:8px 0 0;">${CONFIG.COMPANY_PHONE} | ${CONFIG.COMPANY_EMAIL}</p>
              <p style="margin:12px 0 0; color:#6b6b6b;">&copy; ${year} ${CONFIG.APP_NAME} - All rights reserved</p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>
    `.trim();

    // Generate PDF confirmation
    Logger.log('[sendBookingConfirmation] Generating PDF...');
    let pdfBlob;
    try {
      pdfBlob = generateBookingPDF(data, reference, pricing);
      Logger.log('[sendBookingConfirmation] PDF generated: ' + (pdfBlob ? pdfBlob.getName() : 'null'));
    } catch (pdfError) {
      Logger.log('[sendBookingConfirmation] PDF generation failed: ' + pdfError.message);
      pdfBlob = null;
    }

    // Send the email with PDF attachment
    Logger.log('[sendBookingConfirmation] Sending email to: ' + email);
    
    const emailOptions = {
      to: email,
      subject: subject,
      body: plainText,
      htmlBody: html
    };
    
    // Only attach PDF if it was generated successfully
    if (pdfBlob) {
      emailOptions.attachments = [pdfBlob];
      Logger.log('[sendBookingConfirmation] PDF attached to email');
    }
    
    MailApp.sendEmail(emailOptions);
    Logger.log('[sendBookingConfirmation] Email sent successfully');

    return true;

  } catch (error) {
    Logger.log('[sendBookingConfirmation] Failed to send email: ' + error.message);
    return false;
  }
}

/* =========================
   PDF BOOKING CONFIRMATION GENERATION
========================= */

function generateBookingPDF(data, reference, pricing) {
  const { guestName, email, phone, roomName, roomType, checkInDate, checkOutDate, specialRequests, country } = data;
  const { nights, roomRate, roomTotal, touristLevy, visaFee, totalAmount } = pricing;
  
  const today = new Date();
  const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMMM dd, yyyy');
  
  try {
    // Create a temporary Google Doc
    const tempDoc = DocumentApp.create('Booking_' + reference + '_' + Date.now());
    const docBody = tempDoc.getBody();
    
    // Header with hotel name
    const headerPara = docBody.appendParagraph(CONFIG.APP_NAME);
    headerPara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    headerPara.setForegroundColor('#c9a962');
    headerPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    const subtitle = docBody.appendParagraph('Luxury Hotel Johannesburg');
    subtitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    subtitle.setForegroundColor('#6b7280');
    
    docBody.appendParagraph('');
    
    const titlePara = docBody.appendParagraph('BOOKING CONFIRMATION');
    titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    titlePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    docBody.appendParagraph('Reference: ' + reference).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    docBody.appendParagraph('Date: ' + formattedDate).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    docBody.appendHorizontalRule();
    
    // Guest Information Section
    const guestHeader = docBody.appendParagraph('GUEST INFORMATION');
    guestHeader.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    
    docBody.appendParagraph('Name: ' + (guestName || 'N/A'));
    docBody.appendParagraph('Email: ' + (email || 'N/A'));
    docBody.appendParagraph('Phone: ' + (phone || 'N/A'));
    docBody.appendParagraph('Country: ' + (country || 'South Africa'));
    docBody.appendHorizontalRule();
    
    // Reservation Details Section
    const resHeader = docBody.appendParagraph('RESERVATION DETAILS');
    resHeader.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    
    const resTableData = [
      ['Room Type', roomName || 'Deluxe Suite'],
      ['Check-in Date', checkInDate || 'TBD'],
      ['Check-out Date', checkOutDate || 'TBD'],
      ['Duration', nights + ' night(s)']
    ];
    
    const resTable = docBody.appendTable(resTableData);
    for (let i = 0; i < resTable.getNumRows(); i++) {
      resTable.getRow(i).getCell(0).setBackgroundColor('#f3f4f6');
    }
    
    docBody.appendHorizontalRule();
    
    // Pricing Section
    const priceHeader = docBody.appendParagraph('PRICING SUMMARY');
    priceHeader.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    
    const priceTableData = [
      ['Description', 'Amount (ZAR)'],
      ['Room Rate (' + nights + ' nights @ R' + roomRate.toLocaleString() + ')', 'R ' + roomTotal.toLocaleString()],
      ['Tourist Levy', 'R ' + touristLevy.toLocaleString()]
    ];
    
    if (visaFee > 0) {
      priceTableData.push(['Visa Processing Fee', 'R ' + visaFee.toLocaleString()]);
    }
    
    priceTableData.push(['TOTAL', 'R ' + totalAmount.toLocaleString()]);
    
    const priceTable = docBody.appendTable(priceTableData);
    priceTable.getRow(0).setBackgroundColor('#c9a962');
    for (let i = 0; i < priceTable.getRow(0).getNumCells(); i++) {
      priceTable.getRow(0).getCell(i).setForegroundColor('#ffffff');
    }
    // Style total row
    const lastRow = priceTable.getRow(priceTable.getNumRows() - 1);
    lastRow.setBackgroundColor('#f0fdf4');
    
    docBody.appendParagraph('');
    
    // Special Requests (if any)
    if (specialRequests && specialRequests.trim()) {
      const reqHeader = docBody.appendParagraph('SPECIAL REQUESTS');
      reqHeader.setHeading(DocumentApp.ParagraphHeading.HEADING3);
      docBody.appendParagraph(specialRequests);
      docBody.appendHorizontalRule();
    }
    
    // Hotel Policies
    const polHeader = docBody.appendParagraph('HOTEL POLICIES');
    polHeader.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    
    docBody.appendListItem('Check-in time: 2:00 PM');
    docBody.appendListItem('Check-out time: 10:00 AM');
    docBody.appendListItem('Early check-in and late check-out available upon request (subject to availability)');
    docBody.appendListItem('Cancellation policy: Free cancellation up to 24 hours before check-in');
    docBody.appendListItem('Valid ID or passport required at check-in');
    
    docBody.appendParagraph('');
    
    // Footer
    const footer = docBody.appendParagraph(CONFIG.APP_NAME + ' | ' + CONFIG.COMPANY_ADDRESS);
    footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    footer.setForegroundColor('#9ca3af');
    
    const contact = docBody.appendParagraph(CONFIG.COMPANY_PHONE + ' | ' + CONFIG.COMPANY_EMAIL);
    contact.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    contact.setForegroundColor('#9ca3af');
    
    const thanks = docBody.appendParagraph('We look forward to welcoming you!');
    thanks.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    thanks.setForegroundColor('#c9a962');
    thanks.setBold(true);
    
    // Save and close the document
    tempDoc.saveAndClose();
    
    // Get the document as PDF
    const docFile = DriveApp.getFileById(tempDoc.getId());
    const pdfBlob = docFile.getAs('application/pdf').setName(CONFIG.APP_NAME + '_Booking_' + reference + '.pdf');
    
    // Clean up - delete the temp doc
    docFile.setTrashed(true);
    
    Logger.log('PDF generated successfully for reference: ' + reference);
    return pdfBlob;
    
  } catch (pdfError) {
    Logger.log('PDF generation error: ' + pdfError.message);
    return null;
  }
}

/* =========================
   SEND CUSTOM EMAIL
========================= */

function sendEmailInternal(data) {
  const { reference, to, subject, body } = data;
  
  if (!to || !subject || !body) {
    return { error: 'Missing email parameters' };
  }
  
  try {
    MailApp.sendEmail({
      to: to,
      subject: subject + ' | ' + CONFIG.APP_NAME,
      body: body + '\n\n--\n' + CONFIG.APP_NAME + '\n' + CONFIG.COMPANY_EMAIL
    });
    
    if (reference) {
      addCommentInternal({ reference, author: 'Concierge', text: 'Email sent: ' + subject });
    }
    
    return { success: true };
  } catch (e) {
    return { error: 'Failed to send email: ' + e.message };
  }
}

function sendConfirmationEmail(data) {
  const booking = getBookingByReference(data.reference);
  if (!booking) return { error: 'Booking not found' };
  
  const pricing = {
    nights: booking.nights,
    roomRate: booking.roomRate,
    roomTotal: booking.roomTotal,
    touristLevy: booking.touristLevy,
    visaFee: booking.visaFee,
    totalAmount: booking.totalAmount
  };
  
  const result = sendBookingConfirmation(booking, booking.reference, pricing);
  
  if (result) {
    updateField({ reference: data.reference, field: 'confirmationSent', value: true });
    return { success: true };
  }
  
  return { error: 'Failed to send confirmation' };
}

/* =========================
   HELPER FUNCTIONS
========================= */

function generateReference(name) {
  const prefix = 'MAB';
  const timestamp = Date.now().toString().slice(-8);
  const random = Math.random().toString(36).substr(2, 4).toUpperCase();
  return prefix + '-' + timestamp + random;
}

function generateUUID() {
  return Utilities.getUuid();
}

function sanitize(str) {
  if (typeof str !== 'string') return str;
  return str.replace(/[<>]/g, '').trim();
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getDocumentsMap() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.DOCUMENTS_SHEET);
  if (!sheet) return {};

  const map = {};
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const ref = rows[i][1];
    if (!map[ref]) map[ref] = [];
    map[ref].push({
      id: rows[i][0],
      type: rows[i][2],
      fileName: rows[i][3],
      url: rows[i][4],
      comment: rows[i][5],
      uploadedAt: rows[i][6]
    });
  }
  return map;
}

/* =========================
   ATTACHMENTS / DOCUMENTS
========================= */

function getAllAttachments() {
  Logger.log('[getAllAttachments] Starting...');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.DOCUMENTS_SHEET);
    
    if (!sheet) {
      Logger.log('[getAllAttachments] Documents sheet not found');
      return { success: true, attachments: [], stats: { images: 0, documents: 0, total: 0, storage: '0 MB' } };
    }
    
    const rows = sheet.getDataRange().getValues();
    Logger.log('[getAllAttachments] Found ' + rows.length + ' rows (including header)');
    
    const attachments = [];
    let imageCount = 0;
    let documentCount = 0;
    
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row[0]) continue; // Skip empty rows
      
      const fileName = row[3] || '';
      const fileType = getFileType(fileName);
      
      if (fileType === 'image') imageCount++;
      else documentCount++;
      
      attachments.push({
        id: row[0],
        reference: row[1] || '',
        type: row[2] || 'Document',
        fileName: fileName,
        url: row[4] || '',
        comment: row[5] || '',
        uploadedAt: row[6] ? (row[6] instanceof Date ? row[6].toISOString() : String(row[6])) : new Date().toISOString(),
        fileType: fileType
      });
    }
    
    // Sort by upload date (newest first)
    attachments.sort((a, b) => new Date(b.uploadedAt) - new Date(a.uploadedAt));
    
    Logger.log('[getAllAttachments] Returning ' + attachments.length + ' attachments');
    
    return {
      success: true,
      attachments: attachments,
      stats: {
        images: imageCount,
        documents: documentCount,
        total: attachments.length,
        storage: estimateStorageUsed(attachments.length) + ' MB'
      }
    };
    
  } catch (error) {
    Logger.log('[getAllAttachments] Error: ' + error.message);
    return { success: false, error: error.message, attachments: [] };
  }
}

function getFileType(fileName) {
  if (!fileName) return 'document';
  const ext = fileName.toLowerCase().split('.').pop();
  const imageExts = ['jpg', 'jpeg', 'png', 'gif', 'webp', 'bmp', 'svg'];
  const pdfExts = ['pdf'];
  const docExts = ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'txt', 'csv'];
  
  if (imageExts.includes(ext)) return 'image';
  if (pdfExts.includes(ext)) return 'pdf';
  if (docExts.includes(ext)) return 'document';
  return 'document';
}

function estimateStorageUsed(fileCount) {
  // Rough estimate: average 500KB per file
  return (fileCount * 0.5).toFixed(1);
}

function deleteAttachment(documentId) {
  if (!documentId) return { error: 'Document ID required' };
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.DOCUMENTS_SHEET);
    if (!sheet) return { error: 'Documents sheet not found' };
    
    const rows = sheet.getDataRange().getValues();
    
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === documentId) {
        sheet.deleteRow(i + 1);
        Logger.log('[deleteAttachment] Deleted document: ' + documentId);
        return { success: true };
      }
    }
    
    return { error: 'Document not found' };
  } catch (error) {
    Logger.log('[deleteAttachment] Error: ' + error.message);
    return { error: error.message };
  }
}

/* =========================
   SHEET INITIALIZATION
========================= */

function getOrCreateBookingsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.BOOKINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.BOOKINGS_SHEET);
    sheet.appendRow([
      'Reference', 'Timestamp', 'Status', 'Guest_Name', 'Email', 'Phone', 'Country',
      'Check_In_Date', 'Check_Out_Date', 'Nights', 'Room_Type', 'Room_Name',
      'Room_Rate', 'Room_Total', 'Tourist_Levy', 'Visa_Fee', 'Total_Amount',
      'Special_Requests', 'Request_Options', 'Payment_Status', 'Visa_Required',
      'Confirmation_Sent', 'Guest_Arrived', 'Last_Updated'
    ]);
  }
  return sheet;
}

function getOrCreateDocumentsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.DOCUMENTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.DOCUMENTS_SHEET);
    sheet.appendRow(['Document_ID', 'Reference', 'Type', 'File_Name', 'File_URL', 'Comment', 'Uploaded_At']);
  }
  return sheet;
}

function getOrCreateCommentsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.COMMENTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.COMMENTS_SHEET);
    sheet.appendRow(['Comment_ID', 'Reference', 'Author', 'Text', 'Timestamp']);
  }
  return sheet;
}

function getOrCreateDeletedSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.DELETED_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.DELETED_SHEET);
    sheet.appendRow(['Reference', 'Deleted_At', 'Reason', 'Deleted_By', 'Snapshot']);
  }
  return sheet;
}

function getOrCreateTransactionsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.TRANSACTIONS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.TRANSACTIONS_SHEET);
    sheet.appendRow(['Transaction_ID', 'Reference', 'Date', 'Description', 'Amount', 'Type', 'Created_By', 'Created_At']);
  }
  return sheet;
}

/* =========================
   TRANSACTIONS / BILLING
   Bank-style statement system
========================= */

/**
 * Add a new transaction for a client
 * @param {Object} data - { reference, date, description, amount, type, createdBy }
 */
function addTransaction(data) {
  const { reference, date, description, amount, type, createdBy } = data;
  
  if (!reference) return { error: 'Reference (Client ID) is required' };
  if (!description) return { error: 'Description is required' };
  if (amount === undefined || amount === null) return { error: 'Amount is required' };
  
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = getOrCreateTransactionsSheet(ss);
    
    const transactionId = generateUUID();
    const transactionDate = date || new Date().toISOString().split('T')[0];
    const transactionType = type || (parseFloat(amount) >= 0 ? 'CHARGE' : 'PAYMENT');
    const timestamp = new Date().toISOString();
    
    sheet.appendRow([
      transactionId,                           // 0 - Transaction_ID
      reference,                               // 1 - Reference (Client ID)
      transactionDate,                         // 2 - Date
      sanitize(description),                   // 3 - Description
      parseFloat(amount),                      // 4 - Amount
      transactionType.toUpperCase(),           // 5 - Type
      createdBy || 'Admin',                    // 6 - Created_By
      timestamp                                // 7 - Created_At
    ]);
    
    // Add comment to booking
    addCommentInternal({ 
      reference, 
      author: 'BILLING', 
      text: `Transaction added: ${description} - R ${Math.abs(parseFloat(amount)).toLocaleString()} (${transactionType})`
    });
    
    Logger.log('[addTransaction] Added transaction ' + transactionId + ' for ' + reference);
    
    return { success: true, transactionId };
  } catch (error) {
    Logger.log('[addTransaction] Error: ' + error.message);
    return { error: error.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Delete a transaction by ID
 * @param {Object} data - { transactionId, deletedBy }
 */
function deleteTransaction(data) {
  const { transactionId, deletedBy } = data;
  
  if (!transactionId) return { error: 'Transaction ID is required' };
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.TRANSACTIONS_SHEET);
    
    if (!sheet) return { error: 'Transactions sheet not found' };
    
    const rows = sheet.getDataRange().getValues();
    
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === transactionId) {
        const reference = rows[i][1];
        const description = rows[i][3];
        const amount = rows[i][4];
        
        sheet.deleteRow(i + 1);
        
        // Add comment to booking
        addCommentInternal({ 
          reference, 
          author: 'BILLING', 
          text: `Transaction deleted by ${deletedBy || 'Admin'}: ${description} - R ${Math.abs(amount).toLocaleString()}`
        });
        
        Logger.log('[deleteTransaction] Deleted transaction: ' + transactionId);
        return { success: true };
      }
    }
    
    return { error: 'Transaction not found' };
  } catch (error) {
    Logger.log('[deleteTransaction] Error: ' + error.message);
    return { error: error.message };
  }
}

/**
 * Get all transactions for a client with running balance
 * @param {string} reference - The booking/client reference
 */
function getTransactionsByReference(reference) {
  if (!reference) return { error: 'Reference is required', transactions: [], summary: {} };
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.TRANSACTIONS_SHEET);
    
    if (!sheet) {
      return { 
        success: true, 
        transactions: [], 
        summary: { totalCharges: 0, totalPayments: 0, balance: 0 } 
      };
    }
    
    const rows = sheet.getDataRange().getValues();
    const transactions = [];
    let totalCharges = 0;
    let totalPayments = 0;
    
    // Filter and collect transactions for this reference
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (row[1] !== reference) continue;
      
      const amount = parseFloat(row[4]) || 0;
      
      if (amount >= 0) {
        totalCharges += amount;
      } else {
        totalPayments += Math.abs(amount);
      }
      
      transactions.push({
        id: row[0],
        reference: row[1],
        date: row[2] instanceof Date ? row[2].toISOString().split('T')[0] : String(row[2]),
        description: row[3] || '',
        amount: amount,
        type: row[5] || 'CHARGE',
        createdBy: row[6] || 'Admin',
        createdAt: row[7] ? (row[7] instanceof Date ? row[7].toISOString() : String(row[7])) : ''
      });
    }
    
    // Sort by date (oldest first for proper balance calculation)
    transactions.sort((a, b) => new Date(a.date) - new Date(b.date));
    
    // Calculate running balance
    let runningBalance = 0;
    transactions.forEach(t => {
      runningBalance += t.amount;
      t.balance = runningBalance;
    });
    
    // Return newest first for display
    transactions.reverse();
    
    const balance = totalCharges - totalPayments;
    
    Logger.log('[getTransactionsByReference] Found ' + transactions.length + ' transactions for ' + reference);
    
    return {
      success: true,
      transactions,
      summary: {
        totalCharges,
        totalPayments,
        balance
      }
    };
  } catch (error) {
    Logger.log('[getTransactionsByReference] Error: ' + error.message);
    return { error: error.message, transactions: [], summary: {} };
  }
}

/**
 * Get transactions for a client (callable from HTML)
 */
function getClientTransactions(reference) {
  return getTransactionsByReference(reference);
}

/**
 * Add transaction (callable from HTML)
 */
function addClientTransaction(reference, date, description, amount, type, createdBy) {
  return addTransaction({ reference, date, description, amount, type, createdBy });
}

/**
 * Delete transaction (callable from HTML)
 */
function deleteClientTransaction(transactionId, deletedBy) {
  return deleteTransaction({ transactionId, deletedBy });
}

/* =========================
   STATEMENT PDF GENERATION
   Professional bank-style statement
========================= */

function generateStatementPDF(reference) {
  Logger.log('[generateStatementPDF] Starting for reference: ' + reference);
  
  if (!reference) return null;
  
  try {
    // Get booking details
    const booking = getBookingByReference(reference);
    if (!booking) {
      Logger.log('[generateStatementPDF] Booking not found');
      return null;
    }
    
    // Get transactions
    const { transactions, summary } = getTransactionsByReference(reference);
    
    const today = new Date();
    const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMMM dd, yyyy');
    const statementPeriod = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMMM yyyy');
    
    // Create a temporary Google Doc
    const tempDoc = DocumentApp.create('Statement_' + reference + '_' + Date.now());
    const docBody = tempDoc.getBody();
    
    // ===== HEADER =====
    const headerPara = docBody.appendParagraph(CONFIG.APP_NAME);
    headerPara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    headerPara.setForegroundColor('#c9a962');
    headerPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    const subtitlePara = docBody.appendParagraph('ACCOUNT STATEMENT');
    subtitlePara.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    subtitlePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    subtitlePara.setForegroundColor('#374151');
    
    docBody.appendParagraph('Statement Date: ' + formattedDate).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    docBody.appendHorizontalRule();
    
    // ===== CLIENT INFORMATION =====
    const clientHeader = docBody.appendParagraph('ACCOUNT HOLDER');
    clientHeader.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    clientHeader.setForegroundColor('#1f2937');
    
    docBody.appendParagraph('Name: ' + (booking.guestName || 'N/A'));
    docBody.appendParagraph('Account Reference: ' + reference);
    docBody.appendParagraph('Email: ' + (booking.email || 'N/A'));
    docBody.appendParagraph('Phone: ' + (booking.phone || 'N/A'));
    
    docBody.appendHorizontalRule();
    
    // ===== ACCOUNT SUMMARY =====
    const summaryHeader = docBody.appendParagraph('ACCOUNT SUMMARY');
    summaryHeader.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    summaryHeader.setForegroundColor('#1f2937');
    
    const summaryData = [
      ['Total Charges', 'R ' + (summary.totalCharges || 0).toLocaleString()],
      ['Total Payments', 'R ' + (summary.totalPayments || 0).toLocaleString()],
      ['Current Balance', 'R ' + (summary.balance || 0).toLocaleString()]
    ];
    
    const summaryTable = docBody.appendTable(summaryData);
    summaryTable.getRow(0).setBackgroundColor('#f3f4f6');
    summaryTable.getRow(1).setBackgroundColor('#f3f4f6');
    
    // Style the balance row based on amount
    const balanceRow = summaryTable.getRow(2);
    if (summary.balance > 0) {
      balanceRow.setBackgroundColor('#fef2f2'); // Red tint for amount due
      balanceRow.getCell(1).setForegroundColor('#dc2626');
    } else if (summary.balance < 0) {
      balanceRow.setBackgroundColor('#f0fdf4'); // Green tint for credit
      balanceRow.getCell(1).setForegroundColor('#16a34a');
    } else {
      balanceRow.setBackgroundColor('#f0fdf4');
      balanceRow.getCell(1).setForegroundColor('#16a34a');
    }
    balanceRow.getCell(0).setBold(true);
    balanceRow.getCell(1).setBold(true);
    
    docBody.appendParagraph('');
    
    // ===== TRANSACTION HISTORY =====
    const transHeader = docBody.appendParagraph('TRANSACTION HISTORY');
    transHeader.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    transHeader.setForegroundColor('#1f2937');
    
    if (transactions && transactions.length > 0) {
      // Reverse to show oldest first (chronological order for statement)
      const sortedTrans = [...transactions].reverse();
      
      // Create transaction table
      const transTableData = [['Date', 'Description', 'Debit', 'Credit', 'Balance']];
      
      sortedTrans.forEach(t => {
        const debit = t.amount > 0 ? 'R ' + t.amount.toLocaleString() : '';
        const credit = t.amount < 0 ? 'R ' + Math.abs(t.amount).toLocaleString() : '';
        const balance = 'R ' + (t.balance || 0).toLocaleString();
        
        transTableData.push([
          t.date || '',
          t.description || '',
          debit,
          credit,
          balance
        ]);
      });
      
      const transTable = docBody.appendTable(transTableData);
      
      // Style header row
      const headerRow = transTable.getRow(0);
      headerRow.setBackgroundColor('#c9a962');
      for (let c = 0; c < 5; c++) {
        headerRow.getCell(c).setForegroundColor('#ffffff');
        headerRow.getCell(c).setBold(true);
      }
      
      // Alternate row colors
      for (let r = 1; r < transTable.getNumRows(); r++) {
        if (r % 2 === 0) {
          transTable.getRow(r).setBackgroundColor('#f9fafb');
        }
      }
    } else {
      docBody.appendParagraph('No transactions recorded for this account.');
    }
    
    docBody.appendParagraph('');
    docBody.appendHorizontalRule();
    
    // ===== FOOTER =====
    const footerNote = docBody.appendParagraph('This is an official statement from ' + CONFIG.APP_NAME + '. Please retain for your records.');
    footerNote.setForegroundColor('#6b7280');
    footerNote.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    const footer = docBody.appendParagraph(CONFIG.APP_NAME + ' | ' + CONFIG.COMPANY_ADDRESS);
    footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    footer.setForegroundColor('#9ca3af');
    
    const contact = docBody.appendParagraph(CONFIG.COMPANY_PHONE + ' | ' + CONFIG.COMPANY_EMAIL);
    contact.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    contact.setForegroundColor('#9ca3af');
    
    // Save and close the document
    tempDoc.saveAndClose();
    
    // Get the document as PDF
    const docFile = DriveApp.getFileById(tempDoc.getId());
    const pdfBlob = docFile.getAs('application/pdf').setName(CONFIG.APP_NAME + '_Statement_' + reference + '.pdf');
    
    // Clean up - delete the temp doc
    docFile.setTrashed(true);
    
    Logger.log('[generateStatementPDF] PDF generated successfully for: ' + reference);
    return pdfBlob;
    
  } catch (error) {
    Logger.log('[generateStatementPDF] Error: ' + error.message);
    return null;
  }
}

/**
 * Generate and return PDF as base64 for download
 */
function generateStatementPDFResponse(reference) {
  const pdfBlob = generateStatementPDF(reference);
  if (!pdfBlob) {
    return ContentService.createTextOutput('Error generating PDF').setMimeType(ContentService.MimeType.TEXT);
  }
  return ContentService
    .createTextOutput(Utilities.base64Encode(pdfBlob.getBytes()))
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Generate statement PDF and return base64 (callable from HTML)
 */
function generateStatementPDFBase64(reference) {
  const pdfBlob = generateStatementPDF(reference);
  if (!pdfBlob) {
    return 'Error generating PDF';
  }
  return Utilities.base64Encode(pdfBlob.getBytes());
}

/**
 * Send statement via email with PDF attachment
 */
function sendStatementEmail(data) {
  const { reference, recipientEmail, customMessage } = data;
  
  if (!reference) return { error: 'Reference is required' };
  
  try {
    // Get booking details
    const booking = getBookingByReference(reference);
    if (!booking) return { error: 'Booking not found' };
    
    const email = recipientEmail || booking.email;
    if (!email) return { error: 'No email address available' };
    
    // Get transaction summary
    const { transactions, summary } = getTransactionsByReference(reference);
    
    // Generate PDF
    Logger.log('[sendStatementEmail] Generating PDF for: ' + reference);
    const pdfBlob = generateStatementPDF(reference);
    
    if (!pdfBlob) {
      return { error: 'Failed to generate statement PDF' };
    }
    
    const today = new Date();
    const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMMM dd, yyyy');
    
    const subject = `Account Statement - ${reference} | ${CONFIG.APP_NAME}`;
    
    // Determine balance status
    let balanceStatus = '';
    let balanceColor = '#16a34a';
    if (summary.balance > 0) {
      balanceStatus = 'Amount Due';
      balanceColor = '#dc2626';
    } else if (summary.balance < 0) {
      balanceStatus = 'Credit Balance';
      balanceColor = '#16a34a';
    } else {
      balanceStatus = 'Paid in Full';
      balanceColor = '#16a34a';
    }
    
    // Plain text version
    const plainText = `
Dear ${booking.guestName || 'Valued Guest'},

Please find attached your account statement from ${CONFIG.APP_NAME}.

Account Reference: ${reference}
Statement Date: ${formattedDate}

ACCOUNT SUMMARY:
Total Charges: R ${(summary.totalCharges || 0).toLocaleString()}
Total Payments: R ${(summary.totalPayments || 0).toLocaleString()}
Current Balance: R ${(summary.balance || 0).toLocaleString()} (${balanceStatus})

${customMessage ? 'Message from ' + CONFIG.APP_NAME + ':\n' + customMessage + '\n\n' : ''}
A detailed PDF statement is attached for your records.

If you have any questions about your account, please contact us.

Best regards,
${CONFIG.APP_NAME} Team
${CONFIG.COMPANY_ADDRESS}
${CONFIG.COMPANY_PHONE}
    `.trim();
    
    // HTML version
    const html = `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Account Statement</title>
</head>
<body style="margin:0; padding:0; background-color:#f3f4f6; font-family:'Segoe UI',system-ui,-apple-system,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="padding:24px 16px;">
    <tr>
      <td align="center">
        <table width="100%" cellpadding="0" cellspacing="0" 
               style="max-width:600px; background:#ffffff; border-radius:16px; overflow:hidden; 
                      box-shadow:0 4px 6px rgba(0,0,0,0.1);">
          
          <!-- Header -->
          <tr>
            <td style="background:linear-gradient(135deg, #c9a962 0%, #a88a4a 100%); padding:32px 24px; text-align:center;">
              <h1 style="margin:0; color:#0a0a0a; font-size:28px; font-weight:700;">
                ${CONFIG.APP_NAME}
              </h1>
              <p style="margin:8px 0 0; color:rgba(0,0,0,0.7); font-size:14px; letter-spacing:2px; text-transform:uppercase;">
                Account Statement
              </p>
            </td>
          </tr>

          <!-- Content -->
          <tr>
            <td style="padding:36px 32px; color:#1f2937;">
              <p style="font-size:16px; margin:0 0 20px;">
                Dear <strong>${escapeHtml(booking.guestName || 'Valued Guest')}</strong>,
              </p>

              <p style="font-size:15px; line-height:1.6; margin:0 0 24px; color:#4b5563;">
                Please find attached your account statement. Here is a summary of your account:
              </p>

              <!-- Reference box -->
              <table width="100%" cellpadding="0" cellspacing="0" 
                     style="background:#f9fafb; border:1px solid #e5e7eb; border-radius:12px; margin:20px 0; padding:20px;">
                <tr>
                  <td>
                    <div style="font-size:13px; color:#6b7280; margin-bottom:4px;">Account Reference</div>
                    <div style="font-size:18px; font-weight:700; color:#c9a962; letter-spacing:1px;">${reference}</div>
                  </td>
                  <td style="text-align:right;">
                    <div style="font-size:13px; color:#6b7280; margin-bottom:4px;">Statement Date</div>
                    <div style="font-size:16px; color:#1f2937;">${formattedDate}</div>
                  </td>
                </tr>
              </table>

              <!-- Summary Table -->
              <table width="100%" cellpadding="0" cellspacing="0" style="margin:24px 0; border-collapse:collapse;">
                <tr>
                  <td style="padding:12px 16px; background:#f9fafb; border:1px solid #e5e7eb; color:#4b5563;">Total Charges</td>
                  <td style="padding:12px 16px; background:#f9fafb; border:1px solid #e5e7eb; text-align:right; font-weight:600;">R ${(summary.totalCharges || 0).toLocaleString()}</td>
                </tr>
                <tr>
                  <td style="padding:12px 16px; border:1px solid #e5e7eb; color:#4b5563;">Total Payments</td>
                  <td style="padding:12px 16px; border:1px solid #e5e7eb; text-align:right; font-weight:600; color:#16a34a;">R ${(summary.totalPayments || 0).toLocaleString()}</td>
                </tr>
                <tr>
                  <td style="padding:16px; background:${summary.balance > 0 ? '#fef2f2' : '#f0fdf4'}; border:1px solid #e5e7eb; font-weight:700; font-size:16px;">
                    Current Balance
                  </td>
                  <td style="padding:16px; background:${summary.balance > 0 ? '#fef2f2' : '#f0fdf4'}; border:1px solid #e5e7eb; text-align:right; font-weight:700; font-size:18px; color:${balanceColor};">
                    R ${Math.abs(summary.balance || 0).toLocaleString()}
                    <div style="font-size:12px; font-weight:normal; color:#6b7280;">${balanceStatus}</div>
                  </td>
                </tr>
              </table>

              ${customMessage ? `
              <div style="background:#fffbeb; border-left:4px solid #c9a962; padding:16px; margin:24px 0; border-radius:0 8px 8px 0;">
                <div style="font-size:12px; color:#92400e; margin-bottom:8px; text-transform:uppercase; letter-spacing:1px;">Message from ${CONFIG.APP_NAME}</div>
                <div style="font-size:14px; color:#78350f;">${escapeHtml(customMessage)}</div>
              </div>
              ` : ''}

              <p style="font-size:14px; line-height:1.6; margin:24px 0; color:#4b5563;">
                A detailed PDF statement is attached to this email for your records. If you have any questions about your account, please don't hesitate to contact us.
              </p>

              <p style="margin-top:32px; font-size:15px; color:#1f2937;">
                Best regards,<br>
                <strong style="color:#c9a962;">${CONFIG.APP_NAME} Team</strong>
              </p>
            </td>
          </tr>

          <!-- Footer -->
          <tr>
            <td style="background:#f9fafb; padding:20px; text-align:center; 
                       font-size:12px; color:#6b7280; border-top:1px solid #e5e7eb;">
              <p style="margin:0;">${CONFIG.APP_NAME} | ${CONFIG.COMPANY_ADDRESS}</p>
              <p style="margin:8px 0 0;">${CONFIG.COMPANY_PHONE} | ${CONFIG.COMPANY_EMAIL}</p>
              <p style="margin:12px 0 0; color:#9ca3af;">&copy; ${today.getFullYear()} ${CONFIG.APP_NAME} - All rights reserved</p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>
    `.trim();
    
    // Send the email with PDF attachment
    Logger.log('[sendStatementEmail] Sending to: ' + email);
    
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: plainText,
      htmlBody: html,
      attachments: [pdfBlob]
    });
    
    // Add comment to booking
    addCommentInternal({ 
      reference, 
      author: 'BILLING', 
      text: `Statement emailed to ${email}. Balance: R ${(summary.balance || 0).toLocaleString()}`
    });
    
    Logger.log('[sendStatementEmail] Statement sent successfully');
    
    return { success: true, sentTo: email };
    
  } catch (error) {
    Logger.log('[sendStatementEmail] Error: ' + error.message);
    return { error: 'Failed to send statement: ' + error.message };
  }
}

/**
 * Send statement (callable from HTML)
 */
function sendClientStatement(reference, recipientEmail, customMessage) {
  return sendStatementEmail({ reference, recipientEmail, customMessage });
}

/* =========================
   TEST & DEBUG FUNCTIONS
========================= */

/**
 * Test function to verify bookings retrieval
 * Run this from the GAS editor to check if data is loading
 */
function testGetAllBookings() {
  Logger.log('=== Testing getAllBookings ===');
  
  var result = getAllBookings();
  
  Logger.log('Result: ' + JSON.stringify(result, null, 2));
  Logger.log('Number of bookings: ' + (result.bookings ? result.bookings.length : 0));
  
  if (result.error) {
    Logger.log('ERROR: ' + result.error);
  }
  
  if (result.bookings && result.bookings.length > 0) {
    Logger.log('First booking: ' + JSON.stringify(result.bookings[0], null, 2));
  }
  
  return result;
}

/**
 * Test function to verify comments retrieval
 */
function testGetAllBookingsWithComments() {
  Logger.log('=== Testing getAllBookingsWithComments ===');
  
  var result = getAllBookingsWithComments();
  
  Logger.log('Result bookings count: ' + (result.bookings ? result.bookings.length : 0));
  
  if (result.error) {
    Logger.log('ERROR: ' + result.error);
  }
  
  return result;
}

/**
 * Test sheet access
 */
function testSheetAccess() {
  Logger.log('=== Testing Sheet Access ===');
  
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    Logger.log('Spreadsheet name: ' + ss.getName());
    
    var sheets = ['Bookings', 'Documents', 'Comments', 'Deleted', 'Rooms', 'Transactions'];
    
    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i];
      var sheet = ss.getSheetByName(name);
      if (sheet) {
        Logger.log(name + ': OK (' + sheet.getLastRow() + ' rows, ' + sheet.getLastColumn() + ' cols)');
      } else {
        Logger.log(name + ': MISSING');
      }
    }
    
    return { success: true };
  } catch(e) {
    Logger.log('ERROR: ' + e.message);
    return { error: e.message };
  }
}
