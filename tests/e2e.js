/**
 * End-to-end: real Chromium -> real products.html -> real HTTP -> real
 * products.gs logic running on a simulated Sheets API.
 *
 * The only thing NOT real is Google's own infrastructure.
 */
const http = require('http');
const fs = require('fs');
const vm = require('vm');
const path = require('path');
const { chromium, devices } = require('playwright');

const HTML = fs.readFileSync(path.resolve(__dirname, '..', 'products.html'), 'utf8');
const GS = fs.readFileSync(path.resolve(__dirname, '..', 'products.gs'), 'utf8');

/* ---- Simulated Sheets API (same fakes as the backend suite) ------------- */

class FakeSheet {
  constructor(ro = false) { this.grid = []; this.readOnly = ro; }
  _row(r, c) {
    while (this.grid.length < r) this.grid.push([]);
    const row = this.grid[r - 1];
    while (row.length < c) row.push('');
    return row;
  }
  getLastRow() {
    let last = 0;
    this.grid.forEach((row, i) => { if (row.some(v => v !== '' && v != null)) last = i + 1; });
    return last;
  }
  getRange(row, col, nR = 1, nC = 1) {
    const s = this;
    return {
      setValues(vals) {
        if (s.readOnly) throw new Error('ro');
        vals.forEach((rv, dr) => { const t = s._row(row + dr, col + rv.length - 1); rv.forEach((v, dc) => { t[col + dc - 1] = v; }); });
        return this;
      },
      getValues() {
        const out = [];
        for (let r = 0; r < nR; r++) { const line = []; for (let c = 0; c < nC; c++) { line.push((s.grid[row + r - 1] || [])[col + c - 1] ?? ''); } out.push(line); }
        return out;
      },
      setValue(v) { if (s.readOnly) throw new Error('ro'); s._row(row, col)[col - 1] = v; return this; },
      setFontWeight() { return this; },
      setBackground() { return this; }
    };
  }
  appendRow(v) { if (this.readOnly) throw new Error('ro'); this.grid[this.getLastRow()] = v.slice(); }
  setFrozenRows() {}
}
class FakeDoc {
  constructor(n, ro = false) { this.name = n; this.tabs = {}; this.readOnly = ro; }
  getName() { return this.name; }
  getSheetByName(n) { return this.tabs[n] || null; }
  insertSheet(n) { this.tabs[n] = new FakeSheet(this.readOnly); return this.tabs[n]; }
}

const GOOD = 'https://docs.google.com/spreadsheets/d/abc123/edit';
const RO = 'https://docs.google.com/spreadsheets/d/readonly1/edit';
let REGISTRY = {};

function resetSheets() {
  REGISTRY = { [GOOD]: new FakeDoc('My Inventory'), [RO]: new FakeDoc('Locked Sheet', true) };
}
resetSheets();

const sandbox = {
  SpreadsheetApp: { openByUrl(u) { const d = REGISTRY[u]; if (!d) throw new Error('not found'); return d; } },
  ContentService: { MimeType: { JSON: 'json' }, createTextOutput(t) { return { _text: t, setMimeType() { return this; } }; } },
  Utilities: { formatDate: () => '2026-07-20 10:00' },
  Session: { getScriptTimeZone: () => 'UTC' },
  Logger: { log: () => {} }, console
};
vm.createContext(sandbox);
vm.runInContext(GS, sandbox);

/* ---- Server ------------------------------------------------------------ */

let mode = 'normal';   // normal | netfail | hang
let timeoutMs = null;  // override client timeout for the abort test

const server = http.createServer((req, res) => {
  if (req.url.startsWith('/exec')) {
    if (mode === 'netfail') { req.socket.destroy(); return; }         // real connection failure
    if (mode === 'hang') { return; }                                   // never responds
    let body = '';
    req.on('data', c => body += c);
    req.on('end', () => {
      const send = () => {
        const out = sandbox.doPost({ postData: { contents: body } });
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(out._text);
      };
      // 'slow' widens the in-flight window so a double-tap can be tested.
      if (mode === 'slow') setTimeout(send, 700); else send();
    });
    return;
  }
  let page = HTML.replace('const WEB_APP_URL = "";', `const WEB_APP_URL = "http://localhost:${PORT}/exec";`);
  if (timeoutMs) page = page.replace('const TIMEOUT_MS = 20000;', `const TIMEOUT_MS = ${timeoutMs};`);
  res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
  res.end(page);
});

const PORT = 8787;
const BASE = `http://localhost:${PORT}/`;

/* ---- Harness ----------------------------------------------------------- */

let pass = 0, fail = 0; const failures = [];
function check(label, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (ok) { pass++; console.log(`  PASS  ${label}`); }
  else { fail++; failures.push(label); console.log(`  FAIL  ${label}\n          expected ${JSON.stringify(expected)}\n          actual   ${JSON.stringify(actual)}`); }
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

async function fillProduct(page, p) {
  await page.fill('#p-name', p.name ?? '');
  await page.fill('#p-sku', p.sku ?? '');
  await page.fill('#p-price', p.price ?? '');
  await page.fill('#p-qty', p.qty ?? '');
  if (p.category) await page.fill('#p-category', p.category);
}

const toastText = page => page.locator('#toast').textContent();
const visible = (page, sel) => page.locator(sel).isVisible();

async function connectTo(page, url) {
  await page.fill('#sheet-url', url);
  await page.click('#connect-btn');
  await sleep(400);
}

(async () => {
  await new Promise(r => server.listen(PORT, r));
  const browser = await chromium.launch();

  /* ===== Steps 3-7: connect, add product, verify row ==================== */
  console.log('\n[3,4,5] Connect flow');
  let ctx = await browser.newContext({ ...devices['iPhone 12'] });
  let page = await ctx.newPage();
  const pageErrors = [];
  page.on('pageerror', e => pageErrors.push(e.message));
  await page.goto(BASE);

  check('starts on Connect screen', await visible(page, '#screen-connect'), true);
  check('entry screen hidden initially', await visible(page, '#screen-entry'), false);

  await connectTo(page, GOOD);
  check('connect succeeds -> entry screen', await visible(page, '#screen-entry'), true);
  check('sheet name shown to user', (await page.locator('#sheet-label').textContent()).trim(), 'Saving to My Inventory');

  console.log('\n[6,7] Add product -> row appended');
  await page.click('#add-btn');
  await sleep(350);
  check('modal opened', await visible(page, '#backdrop'), true);
  check('name field autofocused', await page.evaluate(() => document.activeElement.id), 'p-name');

  await fillProduct(page, { name: 'Blue T-Shirt', sku: 'TS-BLU-M', price: '19.99', qty: '40', category: 'Apparel' });
  await page.click('#save-btn');
  await sleep(500);

  check('modal closed after save', await visible(page, '#backdrop'), false);
  check('success toast shown', (await toastText(page)).trim(), '✅ Product added successfully.');
  const tab = () => REGISTRY[GOOD].getSheetByName('Products');
  check('ROW APPENDED to sheet', tab().getRange(2, 1, 1, 6).getValues()[0], ['Blue T-Shirt', 'TS-BLU-M', '', 'Apparel', 19.99, 40]);
  check('session counter updated', (await page.locator('#session-count').textContent()).trim(), '1 added this session');

  console.log('\n[form reset] Fields cleared after success');
  await page.click('#add-btn');
  await sleep(350);
  const vals = await page.evaluate(() => ['p-name','p-sku','p-price','p-qty','p-category','p-desc','p-image','p-barcode'].map(i => document.getElementById(i).value));
  check('all fields empty on reopen', vals, ['','','','','','','','']);

  /* ===== Step 10: duplicate SKU ======================================== */
  console.log('\n[10] Duplicate SKU protection');
  await fillProduct(page, { name: 'Blue T-Shirt Copy', sku: 'TS-BLU-M', price: '19.99', qty: '5' });
  await page.click('#save-btn');
  await sleep(500);
  check('modal STAYS open on duplicate', await visible(page, '#backdrop'), true);
  check('inline SKU error shown', (await page.locator('#p-sku-err').textContent()).trim(), 'A product with that SKU already exists.');
  check('typed values preserved', await page.inputValue('#p-name'), 'Blue T-Shirt Copy');
  check('no duplicate row written', tab().getLastRow(), 2);

  console.log('\n[10b] Correcting the SKU succeeds');
  await page.fill('#p-sku', 'TS-BLU-L');
  await page.click('#save-btn');
  await sleep(500);
  check('save succeeds after fix', await visible(page, '#backdrop'), false);
  check('row now appended', tab().getLastRow(), 3);

  /* ===== Step 8: refresh persistence =================================== */
  console.log('\n[8] Refresh page');
  await page.reload();
  await sleep(400);
  check('still connected after refresh', await visible(page, '#screen-entry'), true);
  check('connect screen NOT shown again', await visible(page, '#screen-connect'), false);
  check('counter resets to 0 (session-scoped)', (await page.locator('#session-count').textContent()).trim(), '');

  console.log('\n[9] Add another product after refresh');
  await page.click('#add-btn');
  await sleep(350);
  await fillProduct(page, { name: 'Red Hat', sku: 'HAT-RED', price: '9.50', qty: '12' });
  await page.click('#save-btn');
  await sleep(500);
  check('third row appended', tab().getLastRow(), 4);
  check('earlier rows intact', tab().getRange(2, 1, 1, 2).getValues()[0], ['Blue T-Shirt', 'TS-BLU-M']);

  /* ===== Client-side validation ======================================== */
  console.log('\n[validation] Required fields');
  await page.click('#add-btn');
  await sleep(350);
  await page.click('#save-btn');
  await sleep(250);
  check('empty form blocked', await visible(page, '#backdrop'), true);
  check('name error', (await page.locator('#p-name-err').textContent()).trim(), 'Please enter a product name.');
  check('sku error', (await page.locator('#p-sku-err').textContent()).trim(), 'Please enter a SKU.');
  check('price error', (await page.locator('#p-price-err').textContent()).trim(), 'Please enter a price.');
  check('qty error', (await page.locator('#p-qty-err').textContent()).trim(), 'Please enter a quantity.');
  check('focus moved to first bad field', await page.evaluate(() => document.activeElement.id), 'p-name');
  check('nothing written to sheet', tab().getLastRow(), 4);

  await fillProduct(page, { name: 'X', sku: 'X-1', price: '-5', qty: '3' });
  await page.click('#save-btn');
  await sleep(250);
  check('negative price rejected', (await page.locator('#p-price-err').textContent()).trim(), 'Please enter a valid price.');

  /* ===== Step 13: network failure ====================================== */
  console.log('\n[13] Network failure');
  mode = 'netfail';
  await page.fill('#p-price', '5');
  await page.click('#save-btn');
  await sleep(700);
  check('friendly toast, no technical detail', (await toastText(page)).trim(), "Couldn't save the product. Please try again.");
  check('modal stays open so data is not lost', await visible(page, '#backdrop'), true);
  check('values still in form', await page.inputValue('#p-name'), 'X');
  check('save button re-enabled', await page.locator('#save-btn').isDisabled(), false);
  check('save button label restored', (await page.locator('#save-btn').textContent()).trim(), 'Save');
  mode = 'normal';

  console.log('\n[13b] Retry after network recovers');
  await page.click('#save-btn');
  await sleep(500);
  check('retry succeeds', await visible(page, '#backdrop'), false);
  check('row appended on retry', tab().getLastRow(), 5);

  await ctx.close();

  /* ===== Step 13c: request timeout ===================================== */
  console.log('\n[13c] Server hang -> client timeout');
  timeoutMs = 1500;
  mode = 'hang';
  ctx = await browser.newContext({ ...devices['iPhone 12'] });
  page = await ctx.newPage();
  await page.goto(BASE);
  await page.evaluate(([u, n]) => localStorage.setItem('product_entry_sheet', JSON.stringify({ url: u, name: n })), [GOOD, 'My Inventory']);
  await page.reload(); await sleep(300);
  await page.click('#add-btn'); await sleep(350);
  await fillProduct(page, { name: 'Timeout Test', sku: 'TO-1', price: '1', qty: '1' });
  await page.click('#save-btn');
  await sleep(2500);
  check('aborts and shows friendly message', (await toastText(page)).trim(), "Couldn't save the product. Please try again.");
  check('button not stuck in saving state', await page.locator('#save-btn').isDisabled(), false);
  mode = 'normal'; timeoutMs = null;
  await ctx.close();

  /* ===== Steps 11,12: bad URL and read-only ============================ */
  console.log('\n[11] Invalid sheet URL (client + server)');
  ctx = await browser.newContext({ ...devices['iPhone 12'] });
  page = await ctx.newPage();
  await page.goto(BASE);

  await page.click('#connect-btn'); await sleep(200);
  check('empty URL blocked', (await page.locator('#sheet-url-err').textContent()).trim(), 'Please paste your Google Sheet link.');

  await connectTo(page, 'hello world');
  check('garbage URL message', (await page.locator('#sheet-url-err').textContent()).trim(), "That doesn't look like a Google Sheet link. Please check and try again.");
  check('stayed on connect screen', await visible(page, '#screen-connect'), true);

  await connectTo(page, 'https://docs.google.com/spreadsheets/d/doesnotexist/edit');
  check('unshared sheet message', (await page.locator('#sheet-url-err').textContent()).trim(), 'We can\'t open that sheet. Make sure it\'s shared as "Anyone with the link → Editor".');

  console.log('\n[12] Read-only sheet');
  await connectTo(page, RO);
  check('read-only message', (await page.locator('#sheet-url-err').textContent()).trim(), 'That sheet is view-only. Change sharing to "Editor" and try again.');
  check('did not enter app', await visible(page, '#screen-entry'), false);

  console.log('\n[recovery] Change sheet escape hatch');
  await connectTo(page, GOOD);
  check('recovers to entry screen', await visible(page, '#screen-entry'), true);
  await page.click('text=Use a different sheet'); await sleep(250);
  check('back to connect screen', await visible(page, '#screen-connect'), true);
  check('stored sheet cleared', await page.evaluate(() => localStorage.getItem('product_entry_sheet')), null);
  await ctx.close();

  /* ===== Double-tap Save =============================================== */
  console.log('\n[stress] Impatient double-tap on Save');
  resetSheets();
  ctx = await browser.newContext({ ...devices['iPhone 12'] });
  page = await ctx.newPage();
  await page.goto(BASE);
  await connectTo(page, GOOD);

  mode = 'slow';
  await page.click('#add-btn'); await sleep(350);
  await fillProduct(page, { name: 'Double Tap', sku: 'DT-1', price: '3', qty: '1' });
  await page.click('#save-btn');
  await sleep(80);
  check('Save disabled while in flight', await page.locator('#save-btn').isDisabled(), true);
  check('Cancel disabled while in flight', await page.locator('#cancel-btn').isDisabled(), true);
  await page.locator('#save-btn').dispatchEvent('click');   // bypasses pointer guard, worst case
  await page.locator('#save-btn').dispatchEvent('click');
  await sleep(1600);
  mode = 'normal';
  const dtTab = () => REGISTRY[GOOD].getSheetByName('Products');
  check('exactly ONE row despite 3 clicks', dtTab().getLastRow(), 2);
  check('modal closed once', await visible(page, '#backdrop'), false);

  console.log('\n[stress] Escape key cannot close mid-save');
  mode = 'slow';
  await page.click('#add-btn'); await sleep(350);
  await fillProduct(page, { name: 'Esc Test', sku: 'ESC-1', price: '2', qty: '1' });
  await page.click('#save-btn');
  await sleep(100);
  await page.keyboard.press('Escape');
  await sleep(100);
  check('modal still open during save', await visible(page, '#backdrop'), true);
  await sleep(1200);
  mode = 'normal';
  check('row saved after escape attempt', dtTab().getLastRow(), 3);

  /* ===== Bulk entry ==================================================== */
  console.log('\n[stress] 25 products back to back');
  resetSheets();
  await page.reload(); await sleep(300);
  const t0 = Date.now();
  for (let i = 1; i <= 25; i++) {
    await page.click('#add-btn');
    await page.waitForFunction(() => document.activeElement.id === 'p-name');
    await fillProduct(page, { name: `Product ${i}`, sku: `SKU-${i}`, price: String(i), qty: String(i * 2) });
    await page.click('#save-btn');
    await page.waitForSelector('#backdrop:not(.open)', { state: 'attached' });
    await page.waitForFunction(() => !document.getElementById('backdrop').classList.contains('open'));
  }
  const bulkTab = REGISTRY[GOOD].getSheetByName('Products');
  check('all 25 rows present (+header)', bulkTab.getLastRow(), 26);
  check('first row correct', bulkTab.getRange(2, 1, 1, 2).getValues()[0], ['Product 1', 'SKU-1']);
  check('last row correct', bulkTab.getRange(26, 1, 1, 2).getValues()[0], ['Product 25', 'SKU-25']);
  check('no rows lost or duplicated', new Set(bulkTab.grid.slice(1).map(r => r[1])).size, 25);
  check('counter shows 25', (await page.locator('#session-count').textContent()).trim(), '25 added this session');
  console.log(`         (25 products in ${((Date.now() - t0) / 1000).toFixed(1)}s)`);
  await ctx.close();
  resetSheets();

  /* ===== Step 14: mobile layout ======================================== */
  console.log('\n[14] Mobile layout at each width');
  const WIDTHS = [320, 360, 375, 390, 414, 768];
  for (const w of WIDTHS) {
    const c = await browser.newContext({ viewport: { width: w, height: 720 }, deviceScaleFactor: 2, isMobile: w < 768, hasTouch: true });
    const p = await c.newPage();
    await p.goto(BASE);
    await p.evaluate(([u, n]) => localStorage.setItem('product_entry_sheet', JSON.stringify({ url: u, name: n })), [GOOD, 'My Inventory']);
    await p.reload(); await sleep(250);
    await p.click('#add-btn'); await sleep(400);

    const m = await p.evaluate(() => {
      const inputs = [...document.querySelectorAll('input, textarea, select')];
      const smallFont = inputs.filter(el => parseFloat(getComputedStyle(el).fontSize) < 16).length;
      const btns = [...document.querySelectorAll('button')].filter(b => b.offsetParent !== null && !b.hidden);
      const smallTap = btns.filter(b => b.getBoundingClientRect().height < 44).length;
      const save = document.getElementById('save-btn').getBoundingClientRect();
      return {
        overflow: document.documentElement.scrollWidth > window.innerWidth,
        smallFont, smallTap,
        saveVisible: save.top >= 0 && save.bottom <= window.innerHeight && save.width > 0
      };
    });

    // Regression: an empty toast used to peek above the bottom edge because
    // its hidden transform was a percentage of its own (tiny) height.
    const toastHidden = await p.evaluate(() => {
      const t = document.getElementById('toast');
      const s = getComputedStyle(t);
      return s.visibility === 'hidden' || parseFloat(s.opacity) === 0;
    });
    check(`${w}px: idle toast fully hidden`, toastHidden, true);

    check(`${w}px: no horizontal scroll`, m.overflow, false);
    check(`${w}px: no input under 16px (iOS zoom)`, m.smallFont, 0);
    check(`${w}px: no tap target under 44px`, m.smallTap, 0);
    check(`${w}px: Save button on screen`, m.saveVisible, true);

    await p.screenshot({ path: `shot-${w}.png`, fullPage: false });
    await c.close();
  }

  check('no uncaught JS errors during run', pageErrors, []);

  await browser.close();
  server.close();

  console.log(`\n${'='.repeat(52)}\n  ${pass} passed, ${fail} failed`);
  if (fail) { console.log('  Failures:\n   - ' + failures.join('\n   - ')); process.exit(1); }
})();
