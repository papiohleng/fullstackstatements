/**
 * ============================================
 * Setup.gs - MABON SUITES HOTEL PLATFORM SETUP
 * ============================================
 *
 * RUN: runSetup()
 *
 * Creates / fixes these sheets:
 * 1. Bookings (24 columns) - Guest hotel bookings
 * 2. Documents   (1-to-many, FK = Reference)
 * 3. Comments    (event log)
 * 4. Deleted     (tombstone with JSON snapshot)
 * 5. Rooms       (catalog of available rooms)
 * 6. Transactions (billing statements)
 *
 * Synced with Code.gs for Mabon Suites Hotel
 * ============================================
 */

function runSetup() {
  Logger.log('========================================');
  Logger.log('MABON SUITES HOTEL PLATFORM SETUP');
  Logger.log('========================================');

  var SPREADSHEET_ID = '1SXUFidabeo_9GE3Nh-FwNKmmaMrlcyoziI0wtDL9yK4';
  var DRIVE_FOLDER_ID = '1EO8gDYZhZEMctWTN8pQV1g9nnjopqj0E';

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log('Spreadsheet: ' + ss.getName());

  /* =====================================================
     BOOKINGS (SOURCE OF TRUTH)
     24 columns – MUST MATCH Code.gs APPEND ORDER
  ===================================================== */

  var bookingsHeaders = [
    'Reference',           // 0  - Unique booking reference (MAB-XXXXXXXX)
    'Timestamp',           // 1  - Submission timestamp
    'Status',              // 2  - new|pending|confirmed|checked_in|checked_out|cancelled
    'Guest_Name',          // 3  - Full name
    'Email',               // 4  - Email address
    'Phone',               // 5  - Phone number
    'Country',             // 6  - Country
    'Check_In_Date',       // 7  - Check-in date
    'Check_Out_Date',      // 8  - Check-out date
    'Nights',              // 9  - Number of nights
    'Room_Type',           // 10 - deluxe|executive|presidential
    'Room_Name',           // 11 - Display name of room
    'Room_Rate',           // 12 - Rate per night in ZAR
    'Room_Total',          // 13 - Room cost (rate * nights)
    'Tourist_Levy',        // 14 - Tourist levy (R100)
    'Visa_Fee',            // 15 - Visa processing fee
    'Total_Amount',        // 16 - Total amount in ZAR
    'Special_Requests',    // 17 - Special requests text
    'Request_Options',     // 18 - JSON array of selected options
    'Payment_Status',      // 19 - pending|partial|paid|refunded
    'Visa_Required',       // 20 - true|false
    'Confirmation_Sent',   // 21 - true|false
    'Guest_Arrived',       // 22 - true|false
    'Last_Updated'         // 23 - Last update timestamp
  ];

  setupSheet(
    ss,
    'Bookings',
    bookingsHeaders,
    '#c9a962'  // Mabon Suites gold accent color
  );

  /* =====================================================
     DOCUMENTS (1-to-many) - Visa docs, passports, etc.
  ===================================================== */

  var documentsHeaders = [
    'Document_ID',   // UUID
    'Reference',     // FK → Bookings.Reference
    'Type',          // Passport|Visa|ID Document|Other
    'File_Name',
    'File_URL',
    'Comment',       // File-specific comment
    'Uploaded_At'
  ];

  setupSheet(
    ss,
    'Documents',
    documentsHeaders,
    '#a88a4a'
  );

  /* =====================================================
     COMMENTS (EVENT LOG)
  ===================================================== */

  var commentsHeaders = [
    'Comment_ID',    // UUID
    'Reference',     // FK
    'Author',        // Admin | SYSTEM | Concierge | Guest Name
    'Text',
    'Timestamp'
  ];

  setupSheet(
    ss,
    'Comments',
    commentsHeaders,
    '#8b7340'
  );

  /* =====================================================
     TRANSACTIONS (BILLING / STATEMENTS)
     For client billing - like bank statements
     Running balance computed on retrieval (not stored)
  ===================================================== */

  var transactionsHeaders = [
    'Transaction_ID',  // 0 - UUID
    'Reference',       // 1 - FK → Bookings.Reference (Client ID)
    'Date',            // 2 - Transaction date
    'Description',     // 3 - What the transaction is for
    'Amount',          // 4 - Positive=charge, Negative=payment/credit
    'Type',            // 5 - CHARGE|PAYMENT|REFUND|ADJUSTMENT
    'Created_By',      // 6 - Admin who created
    'Created_At'       // 7 - Timestamp
  ];

  setupSheet(
    ss,
    'Transactions',
    transactionsHeaders,
    '#059669'  // Green for financial
  );

  /* =====================================================
     DELETED (TOMBSTONE PATTERN)
     MUST MATCH Code.gs:
     appendRow([reference, deletedAt, reason, deletedBy, snapshot])
  ===================================================== */

  var deletedHeaders = [
    'Reference',     // 0
    'Deleted_At',    // 1
    'Reason',        // 2
    'Deleted_By',    // 3
    'Snapshot'       // 4 (JSON string)
  ];

  setupSheet(
    ss,
    'Deleted',
    deletedHeaders,
    '#dc3545'
  );

  /* =====================================================
     ROOMS (CATALOG)
     Pre-populate with Mabon Suites room types
  ===================================================== */

  var roomsHeaders = [
    'Room_ID',       // Unique room identifier
    'Room_Type',     // deluxe|executive|presidential
    'Name',          // Display name
    'Price_ZAR',     // Price per night in ZAR
    'Size_SQM',      // Room size in square meters
    'Description',   // Short description
    'Features',      // JSON array of features
    'Max_Guests',    // Maximum occupancy
    'Active'         // true|false
  ];

  var roomsSheet = setupSheet(
    ss,
    'Rooms',
    roomsHeaders,
    '#c9a962'
  );

  // Populate rooms if sheet is empty (only headers)
  if (roomsSheet.getLastRow() === 1) {
    populateRooms(roomsSheet);
  }

  /* =====================================================
     VERIFY DRIVE FOLDER
  ===================================================== */

  try {
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    Logger.log('Drive folder OK: ' + folder.getName());
  } catch (e) {
    Logger.log('ERROR: Cannot access Drive folder');
    throw e;
  }

  Logger.log('');
  Logger.log('========================================');
  Logger.log('SETUP COMPLETE');
  Logger.log('========================================');
  Logger.log('');
  Logger.log('ARCHITECTURE:');
  Logger.log('  Bookings (1) ----< Documents (N)');
  Logger.log('        |');
  Logger.log('        +----< Comments (N)');
  Logger.log('        |');
  Logger.log('        +----< Transactions (N) [BILLING]');
  Logger.log('        |');
  Logger.log('        +----< Deleted (1, tombstone)');
  Logger.log('');
  Logger.log('  Rooms (catalog) - Reference for room pricing');
  Logger.log('  Transactions - Client billing statements (bank-style)');
  Logger.log('');

  return 'Setup complete. Sheets are architecturally synced.';
}

/* =====================================================
   HELPERS
===================================================== */

function setupSheet(ss, name, headers, color) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    Logger.log('Created sheet: ' + name);
  }

  // Clear existing header row only
  sheet.getRange(1, 1, 1, sheet.getLastColumn() || headers.length).clear();

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground(color)
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  Logger.log(name + ': ' + headers.length + ' columns verified');
  return sheet;
}

function populateRooms(sheet) {
  var rooms = [
    // Deluxe Suite
    [
      'deluxe',
      'deluxe',
      'Deluxe Suite',
      2500,
      45,
      'Elegant suite with modern amenities and city views',
      '["King Bed", "45 m²", "City View", "Free WiFi", "Air Conditioning", "Mini Bar", "Smart TV"]',
      2,
      true
    ],
    // Executive Suite
    [
      'executive',
      'executive',
      'Executive Suite',
      3800,
      60,
      'Spacious suite with dedicated work area and premium amenities',
      '["King Bed", "60 m²", "Work Desk", "Lounge Area", "Free WiFi", "Mini Bar", "Smart TV", "Espresso Machine"]',
      2,
      true
    ],
    // Presidential Suite
    [
      'presidential',
      'presidential',
      'Presidential Suite',
      8500,
      120,
      'Ultimate luxury with private jacuzzi and butler service',
      '["King Bed", "120 m²", "Jacuzzi", "Butler Service", "Private Terrace", "Dining Area", "Premium Bar", "Smart Home"]',
      4,
      true
    ]
  ];

  sheet.getRange(2, 1, rooms.length, 9).setValues(rooms);
  Logger.log('Rooms: ' + rooms.length + ' room types populated');
}

/* =====================================================
   TEST FUNCTION
===================================================== */

function testSetup() {
  var ss = SpreadsheetApp.openById('1SXUFidabeo_9GE3Nh-FwNKmmaMrlcyoziI0wtDL9yK4');
  var sheets = ['Bookings', 'Documents', 'Comments', 'Deleted', 'Rooms', 'Transactions'];
  
  for (var i = 0; i < sheets.length; i++) {
    var s = sheets[i];
    var sh = ss.getSheetByName(s);
    Logger.log(s + ': ' + (sh ? 'OK (' + sh.getLastColumn() + ' cols)' : 'MISSING'));
  }
}
