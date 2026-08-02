/**
 * Products backend — Apps Script Web App
 *
 * Deploy: Deploy > New deployment > Web app
 *   Execute as:      Me
 *   Who has access:  Anyone
 * Copy the /exec URL into WEB_APP_URL in products.html.
 *
 * Because the script executes as the deploying account, it can only open
 * spreadsheets that account can reach. The client makes their sheet reachable
 * once, either by sharing it with the deploying account or by setting
 * link-sharing to "Anyone with the link -> Editor".
 *
 * Every response is a JSON envelope: { ok: true, ... } or { ok: false, code: "..." }.
 * The client branches on `ok`, never on shape, so a failure can never be
 * mistaken for a successful payload.
 */

var SHEET_NAME = 'Products';
var HEADERS = ['Name', 'SKU', 'Barcode', 'Category', 'Price', 'Quantity', 'Description', 'Image URL', 'Added At'];

function doGet() {
  return json({ ok: true, service: 'products', version: 1 });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json({ ok: false, code: 'BAD_REQUEST' });
    }

    var req = JSON.parse(e.postData.contents);

    if (req.action === 'validate') return json(validateSheet(req.sheetUrl));
    if (req.action === 'addProduct') return json(addProduct(req.sheetUrl, req.product));

    return json({ ok: false, code: 'BAD_REQUEST' });
  } catch (err) {
    // Log the real error for the developer; the client only ever sees a code.
    Logger.log('doPost failed: ' + err);
    return json({ ok: false, code: 'SERVER_ERROR' });
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Opens the spreadsheet, distinguishing "bad URL" from "no permission" so the
 * client can tell the user which one to fix.
 */
function openSheet(sheetUrl) {
  if (!sheetUrl || String(sheetUrl).indexOf('/spreadsheets/d/') === -1) {
    throw { code: 'BAD_URL' };
  }

  var file;
  try {
    file = SpreadsheetApp.openByUrl(String(sheetUrl).trim());
  } catch (err) {
    // openByUrl throws the same way for "not found" and "not shared with me",
    // and from the user's side the fix is identical: share the sheet.
    throw { code: 'NO_ACCESS' };
  }

  // Creating the tab and writing headers are the first writes we attempt.
  // If the sheet opened but these fail, the account has view-only access —
  // report that specifically, since the user can fix it by changing sharing.
  var tab;
  try {
    tab = file.getSheetByName(SHEET_NAME);
    if (!tab) {
      tab = file.insertSheet(SHEET_NAME);
    }
    ensureHeaders(tab);
  } catch (err) {
    throw { code: 'READ_ONLY' };
  }

  return { file: file, tab: tab };
}

function ensureHeaders(tab) {
  if (tab.getLastRow() > 0) return;

  tab.getRange(1, 1, 1, HEADERS.length)
     .setValues([HEADERS])
     .setFontWeight('bold')
     .setBackground('#f1f3f4');
  tab.setFrozenRows(1);
}

function validateSheet(sheetUrl) {
  try {
    var target = openSheet(sheetUrl);

    // A read-only open still succeeds, so probe for write access explicitly —
    // otherwise the first Save is where the user would discover the problem.
    try {
      target.tab.getRange(1, 1).setValue(HEADERS[0]);
    } catch (err) {
      return { ok: false, code: 'READ_ONLY' };
    }

    return { ok: true, name: target.file.getName() };
  } catch (err) {
    return { ok: false, code: err.code || 'SERVER_ERROR' };
  }
}

function addProduct(sheetUrl, product) {
  try {
    if (!product) return { ok: false, code: 'BAD_REQUEST' };

    var name = trim(product.name);
    var sku = trim(product.sku);
    var price = product.price;
    var quantity = product.quantity;

    if (!name || !sku) return { ok: false, code: 'MISSING_FIELDS' };
    if (price === '' || price === null || isNaN(Number(price))) return { ok: false, code: 'MISSING_FIELDS' };
    if (quantity === '' || quantity === null || isNaN(Number(quantity))) return { ok: false, code: 'MISSING_FIELDS' };
    // Reject negative price/quantity server-side (the client blocks them, a direct POST does not).
    if (Number(price) < 0 || Number(quantity) < 0) return { ok: false, code: 'INVALID_VALUE' };

    // Serialize the dedup-check + append so two concurrent saves of the same SKU
    // can't both pass the check and append duplicate rows.
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var target = openSheet(sheetUrl);
      var tab = target.tab;

      // Guard against duplicates. This also makes a double-tapped Save idempotent:
      // the retry lands on the existing SKU instead of appending a second row.
      var skuColumn = HEADERS.indexOf('SKU') + 1;
      var lastRow = tab.getLastRow();
      if (lastRow > 1) {
        var existing = tab.getRange(2, skuColumn, lastRow - 1, 1).getValues();
        for (var i = 0; i < existing.length; i++) {
          if (String(existing[i][0]).trim().toLowerCase() === sku.toLowerCase()) {
            return { ok: false, code: 'DUPLICATE_SKU' };
          }
        }
      }

      // A write failure here means access was downgraded to view-only after the
      // sheet was connected. Tell the user that rather than a generic error.
      try {
        tab.appendRow([
          sanitizeCell(name),
          sanitizeCell(sku),
          sanitizeCell(trim(product.barcode)),
          sanitizeCell(trim(product.category)),
          Number(price),
          Number(quantity),
          sanitizeCell(trim(product.description)),
          sanitizeCell(trim(product.imageUrl)),
          Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
        ]);
      } catch (err) {
        return { ok: false, code: 'READ_ONLY' };
      }

      return { ok: true };
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    Logger.log('addProduct failed: ' + err);
    return { ok: false, code: err.code || 'SERVER_ERROR' };
  }
}

function trim(value) {
  return String(value == null ? '' : value).trim();
}

// Neutralize spreadsheet formula/CSV injection: any value starting with = + - @
// is prefixed with a quote so the sheet stores it as text, not a live formula.
function sanitizeCell(value) {
  var s = String(value == null ? '' : value);
  if (/^[=+\-@\t\r]/.test(s)) return "'" + s;
  return s;
}
