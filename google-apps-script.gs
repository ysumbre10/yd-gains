/**
 * YD GAINS — Google Sheets backend
 * ---------------------------------
 * This receives every set you log on the YD Gains site and stores it
 * in a Google Sheet as a permanent, lifelong training log.
 *
 * It UPSERTS: re-logging the same set (same Date + Exercise + Set) updates
 * that row instead of adding a duplicate — so the sheet stays clean.
 *
 * ============================================================
 * ONE-TIME SETUP (about 2 minutes)
 * ============================================================
 * 1. Go to https://sheets.google.com and create a new blank spreadsheet.
 *    Name it anything, e.g. "YD Gains Log".
 * 2. In that sheet: Extensions ▸ Apps Script.
 * 3. Delete whatever code is there, paste THIS ENTIRE FILE, and click Save (💾).
 * 4. Click Deploy ▸ New deployment.
 *      - Click the gear ⚙ next to "Select type" ▸ choose "Web app".
 *      - Description: anything.
 *      - Execute as:        Me
 *      - Who has access:    Anyone
 *    Click Deploy. Approve/authorize when Google asks (it's your own script).
 * 5. Copy the "Web app" URL — it ends in /exec.
 * 6. This URL is hard-coded into the site (the SHEET_URL constant near the top
 *    of the <script> in index.html). It syncs silently in the background — there
 *    is no sync UI. If you ever redeploy to a NEW url, update SHEET_URL to match.
 *
 * That's it. From now on every set you log is appended/updated in the sheet.
 * TIP: if you change this code, use Deploy ▸ Manage deployments ▸ Edit ▸
 * Version: New version, so the SAME /exec URL keeps working (no code change needed).
 */

var HEADER = ['Timestamp', 'Date', 'Day', 'Workout', 'Exercise', 'Set', 'Weight', 'Unit', 'Reps', 'Done'];

// Normalise a Date column value to a 'YYYY-MM-DD' string. Sheets sometimes
// coerces the date string we send into a real Date cell; this brings it back.
function ymd_(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  return String(v);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = ss.getSpreadsheetTimeZone();
    var sh = ss.getSheetByName('Log') || ss.insertSheet('Log');
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADER);
      sh.setFrozenRows(1);
      sh.getRange('B:B').setNumberFormat('@'); // keep the Date column as plain text
    }

    var d = JSON.parse(e.postData.contents);
    var row = [new Date(), d.date, d.day, d.title, d.exercise, d.set, d.weight, (d.unit || 'kg'), d.reps, d.done];

    // Upsert: find an existing row with the same Date + Exercise + Set.
    var values = sh.getDataRange().getValues();
    var target = -1;
    for (var i = 1; i < values.length; i++) {
      if (ymd_(values[i][1], tz) === String(d.date) &&
          values[i][4] === d.exercise &&
          String(values[i][5]) === String(d.set)) {
        target = i + 1; // 1-based row
        break;
      }
    }
    if (target > 0) sh.getRange(target, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Read-back: returns the whole log as JSON. Supports JSONP via ?callback=fn
// so the site can read it cross-origin without CORS headers.
function doGet(e) {
  var out = { ok: true, service: 'YD Gains Log', rows: [] };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = ss.getSpreadsheetTimeZone();
    var sh = ss.getSheetByName('Log');
    if (sh && sh.getLastRow() > 1) {
      var values = sh.getDataRange().getValues();
      for (var i = 1; i < values.length; i++) {
        var r = values[i];
        if (!r[1] && !r[4]) continue; // skip blank rows
        var done = r[9];
        out.rows.push({
          date: ymd_(r[1], tz),
          day: r[2],
          title: r[3],
          exercise: r[4],
          set: String(r[5]),
          weight: r[6],
          unit: r[7] || 'kg',
          reps: r[8],
          done: (done === true || done === 'TRUE' || done === 'true')
        });
      }
    }
  } catch (err) {
    out.ok = false;
    out.error = String(err);
  }
  var json = JSON.stringify(out);
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
