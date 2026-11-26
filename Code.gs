/**
 * Warehouse Inventory Scanner — Apps Script backend
 * ---------------------------------------------------
 * Bind this script to the "Warehouse Inventory Scanner" Google Sheet
 * (Extensions > Apps Script from inside the sheet). It:
 *   1. Creates three tabs: "Products", "Current Locations", and "History"
 *   2. Accepts Add/Remove scan events posted from the scanner web app (doPost)
 *   3. Tracks a quantity per product PER LOCATION (the same product can sit
 *      in several locations at once with different counts), and logs every
 *      add/remove to History
 *   4. Generates a new unique product code on request (doPost, op:
 *      'addProduct') so Steve never has to invent or track numbers by hand
 *      — see "AUTOMATIC PRODUCT CODES" below
 *   5. Flags low stock right on the "Products" tab — a live "Total On Hand"
 *      column plus a red highlight once it drops to or below a product's
 *      Restock Level — see "LOW-STOCK FLAGGING" below
 *   6. Moves stock between two locations as one atomic step (doPost, op:
 *      'moveStock') instead of a manual Remove-then-Add, so a move can't
 *      end up only half-done — see "MOVING STOCK" below
 *
 * MOVING STOCK (2026-10-12)
 * A move is still, underneath, a Remove at the old location and an Add at
 * the new one — there's no other way to represent "this left here and
 * arrived there." What changed is that the app now does both steps as a
 * single request (handleMoveStock_, under the same script lock as
 * everything else) instead of Steve manually doing a Remove scan and then
 * a separate Add scan — the old two-scan way was easy to get only half
 * done (he hit exactly this: an item ended up listed at both the old and
 * new location because only the Add half ever happened). Per Steve
 * (2026-10-12): if there's less at the origin than asked to move, move
 * whatever's actually there rather than rejecting the whole thing — same
 * spirit as a plain Remove's clamp-at-0 behavior below. Logged in History
 * as a "Move Out" row and a "Move In" row rather than a new column, so
 * this needed no schema change at all — see handleMoveStock_.
 *
 * LOW-STOCK FLAGGING (2026-10-01)
 * Most of what Steve sells is unique, one-off inventory he can't restock —
 * once it's sold, it's gone, and there's no "reorder point" for it. Some
 * items ARE restockable multi-unit products, though, and for those he wants
 * a heads-up when supply is getting low. So:
 *   - "Restock Level" (Products, column C) stays exactly what it's always
 *     been: a plain number Steve types in BY HAND, only for the items he
 *     actually restocks. Left blank (the default for everything else), it
 *     means "not a restockable item" and nothing is ever flagged for it —
 *     see the AND($C2<>"", ...) guard in applyRestockFormatting_ below.
 *   - "Total On Hand" (Products, column E, new) is a live formula, not
 *     something the script writes — SUMIF("'Current Locations'!A:A", <this
 *     product's code>, "'Current Locations'!D:D"), i.e. the sum of that
 *     product's Quantity across every location it currently sits in. Since
 *     a 0-quantity row gets deleted outright from Current Locations (see
 *     the data model note below), SUMIF naturally returns 0 once none are
 *     left anywhere — no special case needed. It updates itself the instant
 *     any phone scans a change; the backend never has to touch it after the
 *     row is first created.
 *   - A conditional-formatting rule on "Products" (applied by setupSheet,
 *     see applyRestockFormatting_) highlights the whole row in red the
 *     moment Total On Hand drops to or below Restock Level. That's the only
 *     notification Steve asked for — no in-app warning banner, no email —
 *     so this is deliberately just a Sheets-native visual flag, nothing
 *     more.
 *   - Run setupSheet again after deploying this (same as any time the
 *     sheet's shape changes) — it backfills the Total On Hand formula onto
 *     any product rows added before this feature existed (e.g. ones typed
 *     in by hand) and (re)installs the highlight rule.
 *
 * AUTOMATIC PRODUCT CODES (2026-10-01)
 * New codes are generated, not typed — tap "Add New Product" in the app,
 * type the name, and the server hands back the next code in the sequence:
 * PRODUCT_CODE_PREFIX ("P") + a zero-padded number (e.g. "P0001", "P0042").
 * See getNextProductCode_ / handleAddProduct_ below. This is purely about
 * how NEW codes are chosen going forward — the scanner's validation (does
 * this code exist in "Products"?) never cares about a code's shape, so any
 * product registered by hand before this feature existed stays valid.
 *
 * DATA MODEL (2026-09-30 redesign — read this before changing anything below)
 * A product code represents a PRODUCT TYPE, not one physical unit. Scanning
 * a product's QR code doesn't mean "this one thing is here" — it means "this
 * many units of this product are being added to / removed from this
 * location right now." So:
 *   - "Products" is the registry of product codes you've created. A code
 *     must be listed here, with a name, before the scanner will accept it —
 *     this is what stops a random QR code (a link, someone else's product
 *     barcode, anything) from silently getting recorded as if it were real
 *     inventory. Add a row here once per new product (Product Code + Name;
 *     leave Restock Level blank until that feature is built), print its QR
 *     label, and the scanner will recognize it from then on. Restock Level
 *     is optional and manual — see "LOW-STOCK FLAGGING" above — and Total
 *     On Hand is a live formula, not something you (or the scanner) set.
 *   - "Current Locations" is one row per (Product Code, Location) PAIR, with
 *     a running Quantity — but ONLY while that quantity is above 0. At the
 *     ledger level, moving stock is still just a Remove at the old location
 *     plus an Add at the new one (see "MOVING STOCK" above for how that's
 *     now done as one atomic app action instead of two manual scans). A
 *     quantity dropping to 0 deletes that row outright (per Steve,
 *     2026-10-01 — this reverses an earlier 2026-09-30 decision to keep
 *     0-quantity rows visible). This tab is meant to answer "where is this
 *     right now," so once there are none of a product at a location, that
 *     row has no reason to still be here.
 *   - "History" logs every Add/Remove/Move event with its quantity and the
 *     resulting new total at that location, for a full audit trail. A move
 *     logs as a paired "Move Out" + "Move In" row rather than its own
 *     column shape — see "MOVING STOCK" above.
 *
 * SETUP (one time):
 *   1. Extensions > Apps Script, delete any placeholder code, paste this whole file.
 *   2. In the toolbar function dropdown pick "setupSheet", click Run, approve
 *      the permission prompts (this is your own script touching your own sheet).
 *      It finishes in a couple of seconds — watch the "Execution log" panel at
 *      the bottom for "Setup complete", then go check the three tabs on the sheet.
 *      (There's deliberately no popup dialog here: Apps Script can only show
 *      those when a function is run from inside the Sheets UI itself, e.g. a
 *      custom menu — run from the editor's Run button, a popup call just hangs
 *      until Google kills the execution at the 6-minute limit.)
 *   3. Deploy > New deployment > select type "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone
 *      - Click Deploy, approve access again if asked.
 *   4. Copy the "Web app URL" it gives you (ends in /exec).
 *   5. Paste that URL into the scanner app's Settings screen (gear icon), once —
 *      every phone that opens the app link will pick it up automatically.
 *
 * If you ever change the code, you must Deploy > Manage deployments > edit
 * (pencil) > New version, or phones will keep hitting the old code.
 *
 * PRODUCT CODE COLUMNS ARE FORCED TO PLAIN TEXT ON PURPOSE.
 * Google Sheets auto-converts a purely-numeric string typed or written into a
 * normal cell into a real Number — which silently drops any leading zero
 * (e.g. "091926401" becomes 91926401). If that happens, a code that's
 * scanned again later (still with its leading zero, exactly as the QR
 * encodes it) will no longer text-match the stored value. Formatting the
 * Product Code columns as Plain Text ("@") before writing into them stops
 * Sheets from doing that conversion, so the full code (leading zeros and
 * all) is kept exactly as entered/scanned.
 *
 * "CURRENT LOCATIONS" QUANTITY COLUMN IS FORCED TO A PLAIN NUMBER, FOR A
 * SIMILAR REASON (2026-10-01). Column D used to hold a date ("Last Scanned
 * At") back when this tab had a different layout (see the 2026-09-30 data
 * model note above) — a cell keeps its number FORMAT even after its old
 * value is cleared out. So when this redesign started writing plain
 * quantities (0, 1, 2…) into that same column position, Sheets kept
 * displaying them using the leftover date format — "0" rendered as
 * "12/30/1899", the date equivalent of day zero — even though the
 * underlying number was always correct. `setNumberFormat('0')` is applied
 * to that cell on every write (alongside the Plain-Text formatting on
 * Product Code) so this can't happen again, the same way the leading-zero
 * fix above works. If you ever see a date instead of a number anywhere a
 * quantity should be, this is almost certainly why — reformat that column
 * as Number in the Sheet (Format > Number > Number); the data underneath is
 * fine, only the display was ever wrong.
 *
 * PRODUCT-LOCATIONS LOOKUP (2026-11-26)
 * New doGet action (?action=productLocations&code=P0005) that returns every
 * Current Locations row for one product code — just location + quantity,
 * nothing else. Added so the app's Move Stock flow can ask "where is this
 * currently recorded?" purely from scanning the product, instead of
 * requiring Steve to already know/scan the correct origin location first.
 * See getProductLocations_ below, and index.html's Move Stock flow.
 */

const PRODUCTS_SHEET_NAME = 'Products';
const CURRENT_SHEET_NAME = 'Current Locations';
const HISTORY_SHEET_NAME = 'History';

const PRODUCTS_HEADERS = ['Product Code', 'Name', 'Restock Level', 'Date Added', 'Total On Hand'];
const CURRENT_HEADERS = ['Product Code', 'Product Name', 'Location', 'Quantity', 'Last Updated By', 'Last Updated At'];
const HISTORY_HEADERS = ['Timestamp', 'Product Code', 'Product Name', 'Action', 'Quantity', 'Location', 'New Total At Location', 'Scanned By'];

// New product codes are auto-generated as this letter followed by a
// zero-padded number — e.g. "P0001", "P0042" — short and simple to retype
// by hand into a label maker, and visually distinct from location codes
// (which start with a single "L" — see LOCATION_PREFIX_RE in index.html).
// 4 digits (0001–9999) comfortably covers the "under 10,000 items" scale
// Steve described; if that's ever not enough, raising PRODUCT_CODE_DIGITS
// to 5 is a one-line change and doesn't invalidate any existing code
// (validation only checks the Products tab, never a fixed code length).
const PRODUCT_CODE_PREFIX = 'P';
const PRODUCT_CODE_DIGITS = 4;

// How many rows' worth of each Product Code column to pre-format as Plain
// Text during setup, so ordinary typing/pasting into the sheet also behaves
// (this matters most for the Products tab, where you type codes by hand).
// New rows written by doPost are always force-formatted individually too,
// so this is just a head start — it's fine if your product list grows past it.
const TEXT_FORMAT_ROW_BUFFER = 5000;

/**
 * Run this once by hand from the Apps Script editor to build the sheet.
 * Safe to re-run — it won't wipe existing data, but it DOES overwrite the
 * header row text/formatting on each tab to match the current schema above,
 * so if you're upgrading from an older version of this script, clear out
 * any old data rows that used a different column layout first.
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const products = getOrCreateSheet_(ss, PRODUCTS_SHEET_NAME, PRODUCTS_HEADERS);
  const current = getOrCreateSheet_(ss, CURRENT_SHEET_NAME, CURRENT_HEADERS);
  const history = getOrCreateSheet_(ss, HISTORY_SHEET_NAME, HISTORY_HEADERS);

  formatHeader_(products, PRODUCTS_HEADERS.length);
  formatHeader_(current, CURRENT_HEADERS.length);
  formatHeader_(history, HISTORY_HEADERS.length);

  // Force the Product Code columns to Plain Text so leading zeros never get
  // silently stripped by Sheets' automatic number detection. See the note
  // at the top of this file.
  products.getRange(2, 1, TEXT_FORMAT_ROW_BUFFER, 1).setNumberFormat('@');
  current.getRange(2, 1, TEXT_FORMAT_ROW_BUFFER, 1).setNumberFormat('@');
  history.getRange(2, 2, TEXT_FORMAT_ROW_BUFFER, 1).setNumberFormat('@');

  // Force the Quantity column to a plain Number so it never inherits a
  // leftover date/time format from this tab's old layout (or anything else
  // that might apply a format here later) and silently display 0 as
  // "12/30/1899". See the note at the top of this file. doPost also
  // force-formats this column on every individual write, so this is just a
  // head start covering every row up front, including ones no scan has
  // touched yet.
  current.getRange(2, 4, TEXT_FORMAT_ROW_BUFFER, 1).setNumberFormat('0');

  // Same defensive formatting for the new "Total On Hand" formula column —
  // it should always read as a plain number, never inherit something odd.
  products.getRange(2, 5, TEXT_FORMAT_ROW_BUFFER, 1).setNumberFormat('0');

  // Backfill the Total On Hand formula onto every existing product row —
  // this covers rows added before this feature existed (typed in by hand)
  // as well as ones added through the app, and is safe to re-run (it just
  // overwrites each cell with the same formula). New rows going forward get
  // this written once, directly, in handleAddProduct_.
  const productsLastRow = products.getLastRow();
  if (productsLastRow >= 2) {
    const formulas = [];
    for (let r = 2; r <= productsLastRow; r++) {
      formulas.push([totalOnHandFormula_(r)]);
    }
    products.getRange(2, 5, formulas.length, 1).setFormulas(formulas);
  }

  // Red-highlight any product row where Total On Hand has dropped to or
  // below its (manually-set) Restock Level. setConditionalFormatRules
  // REPLACES whatever rules were on the sheet before, so re-running this is
  // always safe — it never piles up duplicate rules. See "LOW-STOCK
  // FLAGGING" at the top of this file.
  applyRestockFormatting_(products);

  products.setColumnWidths(1, 1, 140);
  products.setColumnWidths(2, 1, 220);
  products.setColumnWidths(3, 1, 120);
  products.setColumnWidths(4, 1, 140);
  products.setColumnWidths(5, 1, 130);

  current.setColumnWidths(1, 1, 140);
  current.setColumnWidths(2, 1, 200);
  current.setColumnWidths(3, 1, 180);
  current.setColumnWidths(4, 1, 100);
  current.setColumnWidths(5, 1, 140);
  current.setColumnWidths(6, 1, 160);

  history.setColumnWidths(1, 1, 160);
  history.setColumnWidths(2, 1, 140);
  history.setColumnWidths(3, 1, 200);
  history.setColumnWidths(4, 1, 90);
  history.setColumnWidths(5, 1, 90);
  history.setColumnWidths(6, 1, 180);
  history.setColumnWidths(7, 1, 160);
  history.setColumnWidths(8, 1, 140);

  // Remove the default empty "Sheet1" if it's still there and untouched.
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && defaultSheet.getLastRow() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  // No popup here on purpose — SpreadsheetApp.getUi() only works when a
  // function is run from inside the Sheets UI (e.g. a custom menu). Run from
  // the Apps Script editor's own Run button, a ui.alert() call just hangs
  // until the execution is killed at the 6-minute limit. Logger.log shows up
  // in the editor's "Execution log" panel instead, safely, every time.
  Logger.log('Setup complete. "Products", "Current Locations" and "History" tabs are ready.');
}

/**
 * Gets or creates a sheet, and always (re)writes its header row to match
 * `headers` — even on a sheet that already has data below it. This is safe
 * (row 1 is just column titles) and ensures the header reflects the current
 * schema if this script's column layout ever changes again later.
 */
function getOrCreateSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  return sheet;
}

function formatHeader_(sheet, numCols) {
  const range = sheet.getRange(1, 1, 1, numCols);
  range.setFontWeight('bold').setBackground('#1f2933').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
}

/**
 * The "Total On Hand" formula for a given Products row: the live sum of
 * that row's Product Code (column A) across every row in "Current
 * Locations", matched against that tab's Quantity column. A plain SUMIF —
 * no IFERROR needed, since SUMIF already returns 0 (not an error) when
 * nothing matches, which is exactly "none of this product anywhere."
 */
function totalOnHandFormula_(row) {
  return "=SUMIF('" + CURRENT_SHEET_NAME + "'!$A:$A,A" + row + ",'" + CURRENT_SHEET_NAME + "'!$D:$D)";
}

/**
 * Installs the single conditional-format rule that is the whole of the
 * "low stock" notification Steve asked for: highlight a Products row red
 * when Restock Level is set (not blank — blank means "not a restockable
 * item, never flag it") AND Total On Hand has fallen to or below it.
 * setConditionalFormatRules replaces this sheet's whole rule list, so
 * calling this again (e.g. every time setupSheet runs) is always safe.
 */
function applyRestockFormatting_(products) {
  const range = products.getRange(2, 1, TEXT_FORMAT_ROW_BUFFER, PRODUCTS_HEADERS.length);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($C2<>"", $E2<=$C2)')
    .setBackground('#f8d7da')
    .setRanges([range])
    .build();
  products.setConditionalFormatRules([rule]);
}

/**
 * Visit the /exec URL in a browser with no parameters to confirm the
 * deployment is live — shows plain text.
 * Add ?action=recent (optionally &limit=25) to get the most recent History
 * rows as JSON instead — this is what the scanner app polls to show a
 * shared "live activity" feed across every phone, since GitHub Pages has
 * no direct line to Claude's own real-time store.
 * Add ?action=recentProducts (optionally &limit=10) to get the most
 * recently registered Products rows — used by the app's "Add New Product"
 * screen to show confirmation of what was just created without switching
 * over to the Sheet.
 * Add ?action=productLocations&code=P0005 to get every Current Locations
 * row for that one product code (just location + quantity) — used by the
 * app's Move Stock flow to find where a product is currently recorded
 * without requiring an origin location to already be scanned/active.
 */
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'recent') {
    return jsonOutput_({ ok: true, recent: getRecentHistory_(e) });
  }
  if (action === 'recentProducts') {
    return jsonOutput_({ ok: true, products: getRecentProducts_(e) });
  }
  if (action === 'productLocations') {
    return jsonOutput_({ ok: true, locations: getProductLocations_(e) });
  }
  return ContentService
    .createTextOutput('Warehouse Scanner backend is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function getRecentHistory_(e) {
  const limit = Math.min(50, Math.max(1, parseInt((e && e.parameter && e.parameter.limit) || '25', 10) || 25));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const history = getOrCreateSheet_(ss, HISTORY_SHEET_NAME, HISTORY_HEADERS);
  const lastRow = history.getLastRow();
  if (lastRow < 2) return [];
  const startRow = Math.max(2, lastRow - limit + 1);
  const numRows = lastRow - startRow + 1;
  const values = history.getRange(startRow, 1, numRows, HISTORY_HEADERS.length).getValues();
  const rows = values.map(function(r) {
    return {
      timestamp: r[0] instanceof Date ? r[0].getTime() : r[0],
      productCode: String(r[1] || ''),
      productName: String(r[2] || ''),
      action: String(r[3] || ''),
      quantity: r[4],
      location: String(r[5] || ''),
      newTotal: r[6],
      scannedBy: String(r[7] || '')
    };
  });
  rows.reverse(); // newest first
  return rows;
}

function getRecentProducts_(e) {
  const limit = Math.min(50, Math.max(1, parseInt((e && e.parameter && e.parameter.limit) || '10', 10) || 10));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const products = getOrCreateSheet_(ss, PRODUCTS_SHEET_NAME, PRODUCTS_HEADERS);
  const lastRow = products.getLastRow();
  if (lastRow < 2) return [];
  const startRow = Math.max(2, lastRow - limit + 1);
  const numRows = lastRow - startRow + 1;
  const values = products.getRange(startRow, 1, numRows, PRODUCTS_HEADERS.length).getValues();
  const rows = values.map(function(r) {
    return {
      productCode: String(r[0] || ''),
      name: String(r[1] || ''),
      restockLevel: r[2],
      dateAdded: r[3] instanceof Date ? r[3].getTime() : r[3]
    };
  });
  rows.reverse(); // newest first
  return rows;
}

/**
 * Returns every "Current Locations" row for one product code, as a plain
 * array of { location, quantity } — nothing else. Used by the app's Move
 * Stock flow (2026-11-26) to figure out, purely from scanning the product,
 * whether it's currently recorded nowhere, in exactly one spot (the normal
 * case — auto-pick it as the move's origin, no prompt needed), or split
 * across more than one location at once (rare — let the person on the
 * phone pick which one, right there, instead of guessing). Order isn't
 * guaranteed to mean anything in particular; the app sorts/displays as
 * needed.
 */
function getProductLocations_(e) {
  const productCode = String((e && e.parameter && e.parameter.code) || '').trim();
  if (!productCode) return [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const current = getOrCreateSheet_(ss, CURRENT_SHEET_NAME, CURRENT_HEADERS);
  const currentData = current.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < currentData.length; i++) {
    const cellCode = String(currentData[i][0]).trim();
    if (codesMatch_(cellCode, productCode)) {
      const qty = Number(currentData[i][3]) || 0;
      if (qty > 0) {
        rows.push({ location: String(currentData[i][2] || ''), quantity: qty });
      }
    }
  }
  return rows;
}

/**
 * True only for a string made up entirely of digits (no letters, spaces, or
 * punctuation) — used to decide when it's safe to compare two codes with
 * their leading zeros ignored (see stripLeadingZeros_ below).
 */
function isNumericId_(s) {
  return /^\d+$/.test(s);
}

/** "091926401" -> "91926401". Only meaningful for isNumericId_ strings. */
function stripLeadingZeros_(s) {
  const stripped = s.replace(/^0+/, '');
  return stripped === '' ? '0' : stripped;
}

/**
 * Compares two codes for equality the forgiving way: exact text match first,
 * and — only when both are purely numeric — also equal once leading zeros
 * are ignored. This lets the sheet "self-heal" a code that lost a leading
 * zero to Sheets' old auto-number behavior, by still recognizing it as the
 * same code on the next scan.
 */
function codesMatch_(a, b) {
  if (a === b) return true;
  return isNumericId_(a) && isNumericId_(b) && stripLeadingZeros_(a) === stripLeadingZeros_(b);
}

/**
 * Looks up a product code in the "Products" tab. Returns { rowIndex, name,
 * restockLevel } if found (rowIndex is 1-indexed, for use with getRange),
 * or null if the code isn't registered.
 */
function findProduct_(productsData, productCode) {
  for (let i = 1; i < productsData.length; i++) {
    const cellCode = String(productsData[i][0]).trim();
    if (codesMatch_(cellCode, productCode)) {
      return {
        rowIndex: i + 1,
        name: String(productsData[i][1] || ''),
        restockLevel: productsData[i][2]
      };
    }
  }
  return null;
}

/**
 * Looks up the existing "Current Locations" row for a (productCode,
 * location) PAIR — not just the product code alone, since the same product
 * can have separate rows (and separate quantities) at different locations.
 * Returns { rowIndex, quantity } if found, or null if this product has no
 * row yet at this specific location.
 */
function findCurrentRow_(currentData, productCode, location) {
  for (let i = 1; i < currentData.length; i++) {
    const cellCode = String(currentData[i][0]).trim();
    const cellLocation = String(currentData[i][2] || '').trim();
    if (codesMatch_(cellCode, productCode) && cellLocation === location) {
      return { rowIndex: i + 1, quantity: Number(currentData[i][3]) || 0 };
    }
  }
  return null;
}

/**
 * Every POST is JSON with an "op" field saying what kind of request it is:
 *   - op: 'addProduct'  → handleAddProduct_  (generate + register a new code)
 *   - op: 'moveStock'   → handleMoveStock_   (Remove at one location, Add at
 *     another, as a single atomic step — see "MOVING STOCK" at the top)
 *   - op: 'scanEvent', or no op at all (older clients never sent one) →
 *     handleScanEvent_ (the plain Add/Remove-stock flow)
 * All three run under the same script lock, so two phones can never be
 * handed the same generated product code, corrupt each other's quantity
 * update, or race a move against a plain scan of the same row.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const body = JSON.parse(e.postData.contents);
    const op = String(body.op || 'scanEvent').trim();
    if (op === 'addProduct') {
      return handleAddProduct_(body);
    }
    if (op === 'moveStock') {
      return handleMoveStock_(body);
    }
    return handleScanEvent_(body);
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Receives { op: 'addProduct', name: string, scannedBy, timestamp }.
 * Generates the next unique product code (see getNextProductCode_) and
 * registers it in "Products" with the given name. Runs inside doPost's
 * script lock, so the same code is never handed out twice even if two
 * phones tap "Add New Product" at nearly the same moment.
 */
function handleAddProduct_(body) {
  const name = String(body.name || '').trim();
  if (!name) {
    return jsonOutput_({ ok: false, error: 'name_required', message: 'Enter a product name first.' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const products = getOrCreateSheet_(ss, PRODUCTS_SHEET_NAME, PRODUCTS_HEADERS);
  const productsData = products.getDataRange().getValues();
  const productCode = getNextProductCode_(productsData);
  const timestamp = body.timestamp ? new Date(body.timestamp) : new Date();

  const newRow = products.getLastRow() + 1;
  products.getRange(newRow, 1).setNumberFormat('@'); // plain text — keep leading zeros
  products.getRange(newRow, 5).setNumberFormat('0');
  products.getRange(newRow, 1, 1, 4).setValues([[productCode, name, '', timestamp]]);
  products.getRange(newRow, 5).setFormula(totalOnHandFormula_(newRow));

  return jsonOutput_({ ok: true, productCode: productCode, name: name });
}

/**
 * Finds the next unused product code: PRODUCT_CODE_PREFIX followed by a
 * zero-padded number one higher than the highest existing code using that
 * same prefix. Codes that don't match the pattern (e.g. ones registered by
 * hand before this feature existed, with no "P" prefix) are simply ignored
 * when finding the highest number — never cause an error, and stay valid
 * forever since acceptance only ever checks Products-tab membership, not
 * code shape. Starts at P0001 if nothing matching the pattern exists yet.
 */
function getNextProductCode_(productsData) {
  const pattern = new RegExp('^' + PRODUCT_CODE_PREFIX + '(\\d+)$', 'i');
  let maxNum = 0;
  for (let i = 1; i < productsData.length; i++) {
    const code = String(productsData[i][0] || '').trim();
    const m = pattern.exec(code);
    if (m) {
      const n = parseInt(m[1], 10);
      if (n > maxNum) maxNum = n;
    }
  }
  const next = maxNum + 1;
  const padded = String(next).padStart(PRODUCT_CODE_DIGITS, '0');
  return PRODUCT_CODE_PREFIX + padded;
}

/**
 * Receives one Add/Remove scan event as JSON:
 *   { productCode: string, location: string, action: 'add'|'remove',
 *     quantity: number, scannedBy: string, timestamp: number }
 * Validates the product code is registered in "Products", then updates (or
 * creates) that product's quantity row at that location in "Current
 * Locations", and appends an entry to "History".
 */
function handleScanEvent_(body) {
  const productCode = String(body.productCode || '').trim();
  const location = String(body.location || '').trim();
  const action = String(body.action || '').trim().toLowerCase();
  const quantity = Math.floor(Number(body.quantity));
  const scannedBy = String(body.scannedBy || '').trim();
  const timestamp = body.timestamp ? new Date(body.timestamp) : new Date();

  if (!productCode || !location) {
    return jsonOutput_({ ok: false, error: 'productCode and location are required' });
  }
  if (action !== 'add' && action !== 'remove') {
    return jsonOutput_({ ok: false, error: 'action must be "add" or "remove"' });
  }
  if (!Number.isFinite(quantity) || quantity <= 0) {
    return jsonOutput_({ ok: false, error: 'quantity must be a positive number' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const products = getOrCreateSheet_(ss, PRODUCTS_SHEET_NAME, PRODUCTS_HEADERS);
  const current = getOrCreateSheet_(ss, CURRENT_SHEET_NAME, CURRENT_HEADERS);
  const history = getOrCreateSheet_(ss, HISTORY_SHEET_NAME, HISTORY_HEADERS);

  const productsData = products.getDataRange().getValues();
  const product = findProduct_(productsData, productCode);
  if (!product) {
    return jsonOutput_({
      ok: false,
      error: 'unknown_product',
      message: 'That code isn\'t a registered product yet — add it to the "Products" tab first.'
    });
  }

  const currentData = current.getDataRange().getValues();
  const existingRow = findCurrentRow_(currentData, productCode, location);
  const previousQuantity = existingRow ? existingRow.quantity : 0;
  const delta = action === 'add' ? quantity : -quantity;
  // Removing more than is on hand clamps at 0 rather than going negative.
  const newQuantity = Math.max(0, previousQuantity + delta);

  if (newQuantity > 0) {
    writeCurrentRow_(current, existingRow, productCode, product.name, location, newQuantity, scannedBy, timestamp);
  } else if (existingRow) {
    // Hit 0 — this (Product Code, Location) pair is no longer "currently"
    // anywhere, so the row is deleted outright rather than left showing a
    // 0 (per Steve, 2026-10-01 — reverses the earlier 2026-09-30 decision
    // to keep 0-quantity rows visible). History below still gets its own
    // row for this event regardless, so the full record of where this
    // product was and when it left is never lost — only the "current
    // snapshot" view loses the row.
    current.deleteRow(existingRow.rowIndex);
  }
  // (If there's no existingRow AND newQuantity is 0 — e.g. a Remove was sent
  // for a product/location that was never added — there's nothing to delete
  // and nothing to write; History still logs the attempt.)

  appendHistoryRow_(history, timestamp, productCode, product.name, action === 'add' ? 'Add' : 'Remove', quantity, location, newQuantity, scannedBy);

  return jsonOutput_({ ok: true, productName: product.name, newQuantity: newQuantity });
}

/**
 * Writes (or creates) a Current Locations row for (productCode, location)
 * with the given quantity, applying the same format guards every write in
 * this file uses — Plain Text on Product Code (leading zeros), plain
 * Number on Quantity (no leftover date formatting). See the top-of-file
 * notes for why both exist. `existingRow` is whatever findCurrentRow_
 * returned — pass null/falsy to create a brand-new row instead.
 */
function writeCurrentRow_(current, existingRow, productCode, productName, location, quantity, scannedBy, timestamp) {
  const rowIndex = existingRow ? existingRow.rowIndex : current.getLastRow() + 1;
  current.getRange(rowIndex, 1).setNumberFormat('@');
  current.getRange(rowIndex, 4).setNumberFormat('0');
  current.getRange(rowIndex, 1, 1, 6)
    .setValues([[productCode, productName, location, quantity, scannedBy, timestamp]]);
}

/**
 * Appends one row to History (Add, Remove, Move Out, or Move In), applying
 * the same Plain-Text guard on Product Code every write in this file uses.
 */
function appendHistoryRow_(history, timestamp, productCode, productName, action, quantity, location, newQuantity, scannedBy) {
  const newHistRow = history.getLastRow() + 1;
  history.getRange(newHistRow, 2).setNumberFormat('@');
  history.getRange(newHistRow, 1, 1, 8).setValues([[
    timestamp, productCode, productName, action, quantity, location, newQuantity, scannedBy
  ]]);
}

/**
 * Receives { op: 'moveStock', productCode, fromLocation, toLocation,
 * quantity, scannedBy, timestamp }. See "MOVING STOCK" at the top of this
 * file for the full story — in short, this does a Remove at fromLocation
 * and an Add at toLocation as one atomic step (same script lock as every
 * other op), so a move can't end up only half-done the way two separate
 * manual scans could.
 *
 * If fewer units are actually at fromLocation than requested, the move is
 * clamped to however many ARE there (same clamp-at-0 spirit as a plain
 * Remove above) rather than rejected outright — the response reports
 * movedQuantity so the app can tell Steve exactly how many moved.
 */
function handleMoveStock_(body) {
  const productCode = String(body.productCode || '').trim();
  const fromLocation = String(body.fromLocation || '').trim();
  const toLocation = String(body.toLocation || '').trim();
  const requestedQuantity = Math.floor(Number(body.quantity));
  const scannedBy = String(body.scannedBy || '').trim();
  const timestamp = body.timestamp ? new Date(body.timestamp) : new Date();

  if (!productCode || !fromLocation || !toLocation) {
    return jsonOutput_({ ok: false, error: 'productCode, fromLocation and toLocation are required' });
  }
  if (fromLocation === toLocation) {
    return jsonOutput_({
      ok: false,
      error: 'same_location',
      message: 'That\'s the same location — scan a different one to move stock to.'
    });
  }
  if (!Number.isFinite(requestedQuantity) || requestedQuantity <= 0) {
    return jsonOutput_({ ok: false, error: 'quantity must be a positive number' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const products = getOrCreateSheet_(ss, PRODUCTS_SHEET_NAME, PRODUCTS_HEADERS);
  const current = getOrCreateSheet_(ss, CURRENT_SHEET_NAME, CURRENT_HEADERS);
  const history = getOrCreateSheet_(ss, HISTORY_SHEET_NAME, HISTORY_HEADERS);

  const productsData = products.getDataRange().getValues();
  const product = findProduct_(productsData, productCode);
  if (!product) {
    return jsonOutput_({
      ok: false,
      error: 'unknown_product',
      message: 'That code isn\'t a registered product yet — add it to the "Products" tab first.'
    });
  }

  const currentData = current.getDataRange().getValues();
  const originRow = findCurrentRow_(currentData, productCode, fromLocation);
  const availableAtOrigin = originRow ? originRow.quantity : 0;
  if (availableAtOrigin <= 0) {
    return jsonOutput_({
      ok: false,
      error: 'nothing_to_move',
      message: 'That product isn\'t currently at ' + fromLocation + ' — nothing to move.'
    });
  }

  // Clamp to what's actually there rather than rejecting the whole move —
  // see the doc comment above.
  const movedQuantity = Math.min(requestedQuantity, availableAtOrigin);
  const newOriginQuantity = availableAtOrigin - movedQuantity;

  if (newOriginQuantity > 0) {
    writeCurrentRow_(current, originRow, productCode, product.name, fromLocation, newOriginQuantity, scannedBy, timestamp);
  } else {
    current.deleteRow(originRow.rowIndex);
  }
  appendHistoryRow_(history, timestamp, productCode, product.name, 'Move Out', movedQuantity, fromLocation, newOriginQuantity, scannedBy);

  // Re-read Current Locations fresh before touching the destination row —
  // deleteRow() above shifts every row below it up by one, so any rowIndex
  // found before that delete could now point at the wrong row.
  const currentDataAfterOrigin = current.getDataRange().getValues();
  const destRow = findCurrentRow_(currentDataAfterOrigin, productCode, toLocation);
  const previousDestQuantity = destRow ? destRow.quantity : 0;
  const newDestQuantity = previousDestQuantity + movedQuantity;
  writeCurrentRow_(current, destRow, productCode, product.name, toLocation, newDestQuantity, scannedBy, timestamp);
  appendHistoryRow_(history, timestamp, productCode, product.name, 'Move In', movedQuantity, toLocation, newDestQuantity, scannedBy);

  return jsonOutput_({
    ok: true,
    productName: product.name,
    movedQuantity: movedQuantity,
    requestedQuantity: requestedQuantity,
    fromLocation: fromLocation,
    toLocation: toLocation,
    newOriginQuantity: newOriginQuantity,
    newDestQuantity: newDestQuantity
  });
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
