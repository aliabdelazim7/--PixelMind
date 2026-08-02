/**
 * Executes the REAL products.gs source against a simulated Google Sheets API.
 * Verifies the logic we control. Does NOT verify Google's actual behaviour.
 */
const fs = require('fs');
const vm = require('vm');
const path = require('path');

const SRC = path.resolve(__dirname, '..', 'products.gs');

/* ---- Fake Google Sheets ------------------------------------------------ */

class FakeSheet {
  constructor(readOnly = false) { this.grid = []; this.readOnly = readOnly; this.frozen = 0; }

  _cell(r, c) {
    while (this.grid.length < r) this.grid.push([]);
    const row = this.grid[r - 1];
    while (row.length < c) row.push('');
    return row;
  }

  getLastRow() {
    let last = 0;
    this.grid.forEach((row, i) => {
      if (row.some(v => v !== '' && v != null)) last = i + 1;
    });
    return last;
  }

  getRange(row, col, numRows = 1, numCols = 1) {
    const sheet = this;
    return {
      setValues(values) {
        if (sheet.readOnly) throw new Error('read-only');
        values.forEach((rowVals, dr) => {
          const target = sheet._cell(row + dr, col + rowVals.length - 1);
          rowVals.forEach((v, dc) => { target[col + dc - 1] = v; });
        });
        return this;
      },
      getValues() {
        const out = [];
        for (let r = 0; r < numRows; r++) {
          const line = [];
          for (let c = 0; c < numCols; c++) {
            const gridRow = sheet.grid[row + r - 1] || [];
            line.push(gridRow[col + c - 1] ?? '');
          }
          out.push(line);
        }
        return out;
      },
      setValue(v) {
        if (sheet.readOnly) throw new Error('read-only');
        sheet._cell(row, col)[col - 1] = v;
        return this;
      },
      setFontWeight() { return this; },
      setBackground() { return this; }
    };
  }

  appendRow(values) {
    if (this.readOnly) throw new Error('read-only');
    this.grid[this.getLastRow()] = values.slice();
  }

  setFrozenRows(n) { this.frozen = n; }
}

class FakeSpreadsheet {
  constructor(name, readOnly = false) { this.name = name; this.tabs = {}; this.readOnly = readOnly; }
  getName() { return this.name; }
  getSheetByName(n) { return this.tabs[n] || null; }
  insertSheet(n) { this.tabs[n] = new FakeSheet(this.readOnly); return this.tabs[n]; }
}

const REGISTRY = {};

const sandbox = {
  SpreadsheetApp: {
    openByUrl(url) {
      const doc = REGISTRY[url];
      if (!doc) throw new Error('No item with the given ID could be found');
      return doc;
    }
  },
  ContentService: {
    MimeType: { JSON: 'json' },
    createTextOutput(t) { return { _text: t, setMimeType() { return this; } }; }
  },
  Utilities: { formatDate: () => '2026-07-20 10:00' },
  Session: { getScriptTimeZone: () => 'UTC' },
  Logger: { log: () => {} },
  console
};

vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(SRC, 'utf8'), sandbox);

/* ---- Harness ----------------------------------------------------------- */

let pass = 0, fail = 0;
const failures = [];

function check(label, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (ok) { pass++; console.log(`  PASS  ${label}`); }
  else {
    fail++;
    failures.push(label);
    console.log(`  FAIL  ${label}\n          expected ${JSON.stringify(expected)}\n          actual   ${JSON.stringify(actual)}`);
  }
}

function post(payload) {
  const res = sandbox.doPost({ postData: { contents: JSON.stringify(payload) } });
  return JSON.parse(res._text);
}

const GOOD = 'https://docs.google.com/spreadsheets/d/abc123/edit';
const RO   = 'https://docs.google.com/spreadsheets/d/readonly1/edit';

function reset() {
  Object.keys(REGISTRY).forEach(k => delete REGISTRY[k]);
  REGISTRY[GOOD] = new FakeSpreadsheet('My Inventory');
  REGISTRY[RO]   = new FakeSpreadsheet('Locked Sheet', true);
}

const product = (over = {}) => Object.assign({
  name: 'Blue T-Shirt', sku: 'TS-BLU-M', price: '19.99', quantity: '40',
  barcode: '', category: 'Apparel', description: '', imageUrl: ''
}, over);

/* ---- Step 5: validation succeeds --------------------------------------- */
console.log('\n[5] Connect / validate');
reset();
check('valid sheet connects, returns name', post({ action: 'validate', sheetUrl: GOOD }), { ok: true, name: 'My Inventory' });
check('Products tab auto-created', !!REGISTRY[GOOD].getSheetByName('Products'), true);
check('headers written', REGISTRY[GOOD].getSheetByName('Products').getRange(1, 1, 1, 9).getValues()[0],
  ['Name', 'SKU', 'Barcode', 'Category', 'Price', 'Quantity', 'Description', 'Image URL', 'Added At']);

/* ---- Step 11: invalid sheet URL ---------------------------------------- */
console.log('\n[11] Invalid sheet URL');
reset();
check('garbage string', post({ action: 'validate', sheetUrl: 'hello world' }), { ok: false, code: 'BAD_URL' });
check('non-sheets google url', post({ action: 'validate', sheetUrl: 'https://google.com/docs' }), { ok: false, code: 'BAD_URL' });
check('empty url', post({ action: 'validate', sheetUrl: '' }), { ok: false, code: 'BAD_URL' });
check('unshared sheet -> NO_ACCESS', post({ action: 'validate', sheetUrl: 'https://docs.google.com/spreadsheets/d/nope/edit' }), { ok: false, code: 'NO_ACCESS' });

/* ---- Step 12: read-only sheet ------------------------------------------ */
console.log('\n[12] Read-only sheet');
reset();
check('view-only, no tab yet -> READ_ONLY', post({ action: 'validate', sheetUrl: RO }), { ok: false, code: 'READ_ONLY' });

// Harder case: the Products tab already exists WITH headers, so ensureHeaders
// is a no-op and nothing writes until the explicit probe. Regression guard for
// the bug where this surfaced as SERVER_ERROR.
reset();
const roDoc = REGISTRY[RO];
roDoc.readOnly = false;                       // seed it as if previously writable
roDoc.insertSheet('Products');
roDoc.getSheetByName('Products').appendRow(
  ['Name', 'SKU', 'Barcode', 'Category', 'Price', 'Quantity', 'Description', 'Image URL', 'Added At']);
roDoc.readOnly = true;
roDoc.getSheetByName('Products').readOnly = true;
check('view-only, tab already exists -> READ_ONLY', post({ action: 'validate', sheetUrl: RO }), { ok: false, code: 'READ_ONLY' });
check('saving to view-only sheet -> READ_ONLY', post({ action: 'addProduct', sheetUrl: RO, product: product() }), { ok: false, code: 'READ_ONLY' });

/* ---- Steps 6/7: add a product ------------------------------------------ */
console.log('\n[6,7] Add product / row appended');
reset();
post({ action: 'validate', sheetUrl: GOOD });
check('save returns ok', post({ action: 'addProduct', sheetUrl: GOOD, product: product() }), { ok: true });
const tab = REGISTRY[GOOD].getSheetByName('Products');
check('sheet now has 2 rows (header + product)', tab.getLastRow(), 2);
check('row contents correct', tab.getRange(2, 1, 1, 9).getValues()[0],
  ['Blue T-Shirt', 'TS-BLU-M', '', 'Apparel', 19.99, 40, '', '', '2026-07-20 10:00']);
check('price stored as number', typeof tab.getRange(2, 5).getValues()[0][0], 'number');
check('quantity stored as number', typeof tab.getRange(2, 6).getValues()[0][0], 'number');

/* ---- Step 9: second product -------------------------------------------- */
console.log('\n[9] Second product after refresh');
check('different SKU accepted', post({ action: 'addProduct', sheetUrl: GOOD, product: product({ name: 'Red Hat', sku: 'HAT-RED' }) }), { ok: true });
check('sheet now has 3 rows', tab.getLastRow(), 3);
check('first product still intact', tab.getRange(2, 1, 1, 2).getValues()[0], ['Blue T-Shirt', 'TS-BLU-M']);

/* ---- Step 10: duplicate SKU -------------------------------------------- */
console.log('\n[10] Duplicate SKU protection');
check('exact duplicate rejected', post({ action: 'addProduct', sheetUrl: GOOD, product: product() }), { ok: false, code: 'DUPLICATE_SKU' });
check('case-insensitive duplicate rejected', post({ action: 'addProduct', sheetUrl: GOOD, product: product({ sku: 'ts-blu-m' }) }), { ok: false, code: 'DUPLICATE_SKU' });
check('whitespace-padded duplicate rejected', post({ action: 'addProduct', sheetUrl: GOOD, product: product({ sku: '  TS-BLU-M  ' }) }), { ok: false, code: 'DUPLICATE_SKU' });
check('no extra row appended by rejects', tab.getLastRow(), 3);

/* ---- Required-field validation ----------------------------------------- */
console.log('\n[extra] Server-side required fields');
reset(); post({ action: 'validate', sheetUrl: GOOD });
check('missing name', post({ action: 'addProduct', sheetUrl: GOOD, product: product({ name: '' }) }), { ok: false, code: 'MISSING_FIELDS' });
check('missing sku', post({ action: 'addProduct', sheetUrl: GOOD, product: product({ sku: '  ' }) }), { ok: false, code: 'MISSING_FIELDS' });
check('missing price', post({ action: 'addProduct', sheetUrl: GOOD, product: product({ price: '' }) }), { ok: false, code: 'MISSING_FIELDS' });
check('non-numeric price', post({ action: 'addProduct', sheetUrl: GOOD, product: product({ price: 'abc' }) }), { ok: false, code: 'MISSING_FIELDS' });
check('missing quantity', post({ action: 'addProduct', sheetUrl: GOOD, product: product({ quantity: '' }) }), { ok: false, code: 'MISSING_FIELDS' });
check('zero price allowed', post({ action: 'addProduct', sheetUrl: GOOD, product: product({ sku: 'FREE-1', price: '0', quantity: '0' }) }), { ok: true });

/* ---- Malformed requests ------------------------------------------------ */
console.log('\n[extra] Malformed requests never 500 to the user');
reset();
check('unknown action', post({ action: 'explode', sheetUrl: GOOD }), { ok: false, code: 'BAD_REQUEST' });
check('null product', post({ action: 'addProduct', sheetUrl: GOOD, product: null }), { ok: false, code: 'BAD_REQUEST' });
check('no postData', JSON.parse(sandbox.doPost({})._text), { ok: false, code: 'BAD_REQUEST' });
check('invalid JSON body', JSON.parse(sandbox.doPost({ postData: { contents: '{oops' } })._text), { ok: false, code: 'SERVER_ERROR' });
check('addProduct to unshared sheet', post({ action: 'addProduct', sheetUrl: 'https://docs.google.com/spreadsheets/d/gone/edit', product: product() }), { ok: false, code: 'NO_ACCESS' });

/* ---- Envelope invariant ------------------------------------------------ */
console.log('\n[invariant] No response can be mistaken for success');
const probes = [
  { action: 'explode' },
  { action: 'validate', sheetUrl: 'junk' },
  { action: 'addProduct', sheetUrl: GOOD, product: null }
];
let clean = true;
probes.forEach(p => { const r = post(p); if (r.ok !== false || !r.code) clean = false; });
check('every failure has ok:false + a code', clean, true);

console.log(`\n${'='.repeat(52)}\n  ${pass} passed, ${fail} failed`);
if (fail) { console.log('  Failures: ' + failures.join(', ')); process.exit(1); }
