/**
 * ================================================================
 * LIVELIGOOD — Google Apps Script
 * ================================================================
 * Deploy this as a Web App to:
 *   1. Handle bulk inquiry form submissions (POST)
 *   2. Serve enterprise/product data as JSON (GET)
 *
 * DEPLOYMENT STEPS:
 *   1. Open your Google Sheet → Extensions → Apps Script
 *   2. Paste this entire file into Code.gs
 *   3. Click Deploy → New Deployment → Web App
 *   4. Set "Execute as" = Me, "Who has access" = Anyone
 *   5. Copy the Web App URL → paste into CONFIG.appsScriptUrl in index.html
 * ================================================================
 */

// ----------------------------------------------------------------
// GET — Return data from any sheet as JSON
// ----------------------------------------------------------------
function doGet(e) {
  const sheet = e.parameter.sheet || 'enterprises';
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ws = ss.getSheetByName(sheet);
    if (!ws) {
      return jsonResponse({ error: `Sheet "${sheet}" not found` }, 404);
    }
    const data = sheetToJSON(ws);
    return jsonResponse({ success: true, sheet, rows: data.length, data });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message }, 500);
  }
}

// ----------------------------------------------------------------
// POST — Handle bulk inquiry form submission
// ----------------------------------------------------------------
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Write to bulk_inquiries sheet
    let ws = ss.getSheetByName('bulk_inquiries');
    if (!ws) {
      ws = ss.insertSheet('bulk_inquiries');
      // Add headers if sheet is new
      ws.appendRow([
        'timestamp', 'name', 'organisation', 'email', 'phone',
        'product_interest', 'quantity', 'budget_range', 'message', 'status'
      ]);
    }

    ws.appendRow([
      payload.timestamp || new Date().toISOString(),
      payload.name        || '',
      payload.org         || '',
      payload.email       || '',
      payload.phone       || '',
      payload.product     || '',
      payload.quantity    || '',
      payload.budget      || '',
      payload.message     || '',
      'New',              // status column for team tracking
    ]);

    // Send confirmation email to the inquirer
    if (payload.email) {
      sendConfirmationEmail(payload);
    }

    // Notify internal team
    notifyTeam(payload);

    return jsonResponse({ success: true, message: 'Inquiry received successfully.' });
  } catch (err) {
    console.error('doPost error:', err);
    return jsonResponse({ success: false, error: err.message }, 500);
  }
}

// ----------------------------------------------------------------
// Send confirmation email to the person who submitted
// ----------------------------------------------------------------
function sendConfirmationEmail(data) {
  const subject = 'Liveligood — Your Bulk Inquiry Has Been Received! 🌱';
  const body = `
Hi ${data.name || 'there'},

Thank you for reaching out about bulk / social procurement through Liveligood!

Here's a summary of your inquiry:
━━━━━━━━━━━━━━━━━━━━━━━━
Organisation   : ${data.org || '—'}
Product Interest: ${data.product || 'All Products'}
Estimated Qty  : ${data.quantity || 'Not specified'}
Budget Range   : ${data.budget || 'Not specified'}
━━━━━━━━━━━━━━━━━━━━━━━━

Our social procurement team will get back to you within 24 hours
with a customised proposal and a free SROI (Social Return on
Investment) estimate for your order.

While you wait, feel free to explore our marketplace:
👉 https://liveligood.in

With impact,
The Liveligood Team
📧 bulk@liveligood.in | 📞 +91 98765 43210
🌱 India's Social Impact Marketplace

---
You're receiving this because you submitted a bulk inquiry on Liveligood.
`.trim();

  GmailApp.sendEmail(data.email, subject, body);
}

// ----------------------------------------------------------------
// Notify internal procurement team
// ----------------------------------------------------------------
function notifyTeam(data) {
  const TEAM_EMAIL = 'bulk@liveligood.in'; // Change to your team email
  const subject = `[Liveligood] New Bulk Inquiry from ${data.org || 'Unknown Org'}`;
  const body = `
New bulk inquiry received on Liveligood!

Name         : ${data.name}
Organisation : ${data.org}
Email        : ${data.email}
Phone        : ${data.phone || '—'}
Product      : ${data.product}
Quantity     : ${data.quantity || '—'}
Budget       : ${data.budget || '—'}
Message      : ${data.message || '—'}
Timestamp    : ${data.timestamp}

View all inquiries in the bulk_inquiries sheet.
`.trim();

  GmailApp.sendEmail(TEAM_EMAIL, subject, body);
}

// ----------------------------------------------------------------
// Convert a sheet to array of JSON objects
// ----------------------------------------------------------------
function sheetToJSON(ws) {
  const rows = ws.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i].toString() : ''; });
    return obj;
  });
}

// ----------------------------------------------------------------
// Helper: return JSON ContentService response
// ----------------------------------------------------------------
function jsonResponse(obj, statusCode) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------
// Trigger: Send a weekly impact digest email to subscribers
// (Set up in Apps Script → Triggers → Weekly)
// ----------------------------------------------------------------
function weeklyImpactDigest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ents = sheetToJSON(ss.getSheetByName('enterprises'));
  const prods = sheetToJSON(ss.getSheetByName('products'));
  
  const totalLivelihoods = ents.reduce((s, e) => s + (+e.total_livelihoods || 0), 0);
  const totalWater = ents.reduce((s, e) => s + (+e.total_water_saved_litres || 0), 0);
  
  console.log(`Weekly Digest: ${ents.length} enterprises, ${prods.length} products, ${totalLivelihoods} livelihoods`);
  // Add email logic here if needed
}
