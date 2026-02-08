# Mabon Suites Hotel - Google Apps Script Full-Stack Application

## Overview

A complete hotel booking management system built with Google Apps Script, featuring:
- **Public Booking Form** - Guest-facing reservation form with room selection and visa document upload
- **Admin Dashboard** - Manage bookings, update statuses, send confirmations
- **Billing/Statements** - Bank-style transaction management with PDF generation and email delivery
- **Document Management** - View and manage uploaded guest documents
- **Deleted Records** - Soft-delete with restore capability

## File Structure

```
gas-project/
├── Code.gs          # Main backend logic (all server functions)
├── Setup.gs         # One-time setup script to create sheets
├── Form.html        # Guest booking form (public-facing)
├── Submissions.html # Admin dashboard for managing bookings
├── Deleted.html     # Deleted records management
├── Attachments.html # Document viewer/manager
└── README.md        # This file
```

## Setup Instructions

### Step 1: Create a new Google Apps Script Project

1. Go to [script.google.com](https://script.google.com)
2. Click **"New Project"**
3. Name it "Mabon Suites Hotel Booking System"

### Step 2: Add the Code Files

1. In the Apps Script editor, replace the default `Code.gs` content with the content from `Code.gs`
2. Click **File > New > Script** and create `Setup.gs`, paste its content
3. Click **File > New > HTML file** for each HTML file:
   - `Form.html`
   - `Submissions.html`
   - `Deleted.html`
   - `Attachments.html`

### Step 3: Configure Resource IDs

In `Code.gs`, update these constants with your own IDs:

```javascript
const CONFIG = {
  SPREADSHEET_ID: 'YOUR_GOOGLE_SHEET_ID',
  DRIVE_FOLDER_ID: 'YOUR_GOOGLE_DRIVE_FOLDER_ID',
  // ... rest of config
};
```

Also update in `Setup.gs`:
```javascript
var SPREADSHEET_ID = 'YOUR_GOOGLE_SHEET_ID';
var DRIVE_FOLDER_ID = 'YOUR_GOOGLE_DRIVE_FOLDER_ID';
```

### Step 4: Run Initial Setup

1. In Apps Script editor, open `Setup.gs`
2. Select `runSetup` from the function dropdown
3. Click **Run**
4. Authorize when prompted (this creates all required sheets)

### Step 5: Deploy as Web App

1. Click **Deploy > New deployment**
2. Click the gear icon next to "Select type" and choose **Web app**
3. Configure:
   - **Description**: "Mabon Suites Hotel Booking System v1.0"
   - **Execute as**: "Me"
   - **Who has access**: "Anyone" (for public booking form)
4. Click **Deploy**
5. Copy the Web App URL

## Accessing the Application

Use URL parameters to access different pages:

| Page | URL |
|------|-----|
| Booking Form (Default) | `YOUR_WEB_APP_URL` |
| Booking Form | `YOUR_WEB_APP_URL?action=form` |
| Admin Dashboard | `YOUR_WEB_APP_URL?action=admin` |
| Deleted Records | `YOUR_WEB_APP_URL?action=deleted` |
| Attachments | `YOUR_WEB_APP_URL?action=attachments` |

## Key Features

### Room Selection Fix

The **main issue** you reported was that room selection wasn't working in the deployed version. This has been fixed by:

1. **Changed from clickable divs to a native `<select>` element** - Native form elements work reliably in Google Apps Script's sandboxed environment
2. **Added hidden input fields** to store room data for form submission
3. **Used proper event handlers** that work in deployed context

### Booking Features

- Real-time price calculation
- Date validation (check-out after check-in)
- Visa document upload (max 3 files, 10MB each)
- Special requests checkboxes
- PDF booking confirmation sent via email

### Admin Dashboard Features

- Filter bookings by status, payment status, room type
- Search by guest name, email, reference
- Update booking status with one click
- Send confirmation emails
- View booking details and activity log
- Add internal notes/comments

### Billing/Statement System

- Add charges, payments, refunds, adjustments
- Running balance calculation
- Generate professional PDF statements
- Email statements directly to guests
- Download statements as PDF

### Document Management

- View all uploaded documents
- Filter by type (images, documents)
- Preview images directly
- Download documents

## Google Sheets Structure

The system creates these sheets:

| Sheet | Purpose |
|-------|---------|
| Bookings | Main booking data (24 columns) |
| Documents | Uploaded files (visa docs, passports) |
| Comments | Activity log / notes |
| Transactions | Billing records |
| Deleted | Soft-deleted bookings (with JSON snapshot) |
| Rooms | Room catalog (pre-populated) |

## Troubleshooting

### Room Selection Not Working

The original issue was that clickable div elements don't reliably receive click events in Google Apps Script's iframe sandbox. The fix uses:

```html
<select class="room-select" id="roomSelect" name="roomSelect" required>
  <option value="" disabled selected>-- Choose a room --</option>
  <option value="deluxe" data-price="2500">Deluxe Suite - R 2,500 / night</option>
  ...
</select>
```

Instead of:
```html
<div class="room-option" onclick="selectRoom(this)">...</div>
```

### Form Not Submitting

Make sure all required fields are filled:
- Full Name
- Email
- Phone
- Check-in/Check-out dates
- Room selection

### PDF Not Generating

The system creates temporary Google Docs, converts to PDF, and deletes the temp doc. Ensure your account has permission to:
- Create documents in Drive
- Access the configured Drive folder

### Email Not Sending

Check:
- Valid email address entered
- MailApp daily quota not exceeded (100 emails/day for free accounts)

## Security Notes

- Files uploaded to Drive are set to "Anyone with link can view"
- No authentication is implemented (add your own if needed)
- Spreadsheet and Drive folder should have appropriate sharing settings

## Customization

### Room Types

Edit the `ROOMS` object in `Form.html` and the room data in `Setup.gs` to add/modify room types.

### Pricing

- `TOURIST_LEVY`: R100 (fixed)
- `VISA_FEE_PER_FILE`: R500 per document
- `MAX_VISA_FEE`: R1,500 cap

### Branding

Update `CONFIG` in `Code.gs`:
```javascript
APP_NAME: 'Your Hotel Name',
COMPANY_EMAIL: 'your@email.com',
COMPANY_PHONE: '+XX XX XXX XXXX',
COMPANY_ADDRESS: 'Your Address',
```

## Support

This is a self-contained Google Apps Script application. For issues:
1. Check the Execution Logs in Apps Script
2. Verify spreadsheet and folder permissions
3. Test with the `testSetup()` function in Setup.gs
