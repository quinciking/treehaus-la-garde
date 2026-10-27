/**
 * Treehaus · La Garde Hub — Google Sheets backend
 *
 * One-time setup (about 3 minutes):
 *  1. Go to https://script.google.com, click New project, and paste this whole file in.
 *  2. Fill in GUEST_CODE_TO_SET and HOST_CODE_TO_SET below, then choose `setup` in the
 *     toolbar and press Run. Approve the permissions. This creates the tabs in the
 *     "Treehaus Hub · Live data" sheet and loads the starting data.
 *  3. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
 *     Copy the web app URL into config.js in the GitHub repo.
 * After setup you can clear the two codes below; they're stored in Script properties.
 */
const SHEET_ID = '1ZpkvFmMKMdgqh60cI1MfPldcUCFyRRTvDLQ1gX0a9fc';
const GUEST_CODE_TO_SET = '';  // passcode guests type to open the hub
const HOST_CODE_TO_SET = '';   // separate code for Quinci and Tristan (unlocks trip settings)

const COLLECTIONS = ['guests', 'meals', 'workshops', 'expenses', 'settings'];
const HOST_ONLY = ['settings'];

function setup() {
  const props = PropertiesService.getScriptProperties();
  if (GUEST_CODE_TO_SET) props.setProperty('GUEST_CODE', GUEST_CODE_TO_SET);
  if (HOST_CODE_TO_SET) props.setProperty('HOST_CODE', HOST_CODE_TO_SET);
  if (!props.getProperty('GUEST_CODE') || !props.getProperty('HOST_CODE')) {
    throw new Error('Set GUEST_CODE_TO_SET and HOST_CODE_TO_SET, then run setup again.');
  }
  const ss = SpreadsheetApp.openById(SHEET_ID);
  COLLECTIONS.forEach(function (c) {
    let sh = ss.getSheetByName(c);
    if (!sh) {
      sh = ss.insertSheet(c);
      sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setNumberFormat('@');
      sh.getRange(1, 1).setValue('id');
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, sh.getMaxColumns()).setFontWeight('bold');
    }
  });
  const blank = ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1) ss.deleteSheet(blank);
  Object.keys(SEED).forEach(function (c) {
    const sh = ss.getSheetByName(c);
    if (sh.getLastRow() <= 1) SEED[c].forEach(function (d) { upsert_(sh, d.id, d); });
  });
}

function doGet(e) {
  const p = e.parameter || {};
  const role = role_(p.key, p.hostKey);
  if (!role) return json_({ ok: false, error: 'unauthorized' });
  if (p.action === 'all') return json_({ ok: true, host: role === 'host', data: readAll_() });
  return json_({ ok: true, host: role === 'host' });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  const role = role_(b.key, b.hostKey);
  if (!role) return json_({ ok: false, error: 'unauthorized' });
  if (COLLECTIONS.indexOf(b.collection) < 0 || !b.id || !/^[\w.~:@+-]{1,200}$/.test(b.id)) return json_({ ok: false, error: 'bad_request' });
  if (HOST_ONLY.indexOf(b.collection) >= 0 && role !== 'host') return json_({ ok: false, error: 'host_only' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(b.collection);
    if (b.op === 'set') upsert_(sh, b.id, b.data || {});
    else if (b.op === 'delete') remove_(sh, b.id);
    else return json_({ ok: false, error: 'bad_request' });
  } finally { lock.releaseLock(); }
  return json_({ ok: true });
}

function role_(key, hostKey) {
  const props = PropertiesService.getScriptProperties();
  const g = props.getProperty('GUEST_CODE'), h = props.getProperty('HOST_CODE');
  if (h && (hostKey === h || key === h)) return 'host';
  if (g && key === g) return 'guest';
  return null;
}

function readAll_() {
  const ss = SpreadsheetApp.openById(SHEET_ID), out = {};
  COLLECTIONS.forEach(function (c) {
    const sh = ss.getSheetByName(c);
    if (!sh) { out[c] = []; return; }
    const v = sh.getDataRange().getDisplayValues(), head = v[0];
    out[c] = v.slice(1).filter(function (r) { return r[0] !== ''; }).map(function (r) {
      const o = {};
      head.forEach(function (k, i) { if (k && r[i] !== '') o[k] = decode_(r[i]); });
      o.id = String(r[0]);
      return o;
    });
  });
  return out;
}

function upsert_(sh, id, data) {
  data = Object.assign({}, data); delete data.id;
  let head = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
  Object.keys(data).forEach(function (k) {
    if (head.indexOf(k) < 0) {
      const col = head.length + 1;
      if (col > sh.getMaxColumns()) sh.insertColumnAfter(sh.getMaxColumns());
      sh.getRange(1, col, sh.getMaxRows(), 1).setNumberFormat('@');
      sh.getRange(1, col).setValue(k).setFontWeight('bold');
      head.push(k);
    }
  });
  const row = head.map(function (k) { return k === 'id' ? id : (k in data ? encode_(data[k]) : ''); });
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().map(function (r) { return r[0]; }) : [];
  const at = ids.indexOf(id);
  const r = at >= 0 ? at + 2 : sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, head.length).setNumberFormat('@').setValues([row]);
}

function remove_(sh, id) {
  if (sh.getLastRow() <= 1) return;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().map(function (r) { return r[0]; });
  const at = ids.indexOf(id);
  if (at >= 0) sh.deleteRow(at + 2);
}

function encode_(v) { return typeof v === 'string' ? v : JSON.stringify(v); }
function decode_(s) {
  if (/^(true|false|null|-?\d+(\.\d+)?|\[[\s\S]*\]|\{[\s\S]*\})$/.test(s)) { try { return JSON.parse(s); } catch (e) {} }
  return s;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// Starting data, copied from the Claude version of the hub on 26 Oct 2026. Only loaded into empty tabs.
const SEED = {"guests":[{"id":"r1-kayra","host":false,"name":"Kayra Theodore","retreat":"r1","room":"Room 1","status":"confirmed"},{"id":"r1-naomi","host":false,"name":"Naomi Lilly","retreat":"r1","room":"Dorm room 3","status":"confirmed"},{"id":"r1-nina","host":false,"name":"Christina Alcinay","retreat":"r1","room":"Room 4","status":"confirmed"},{"id":"r1-quinci","host":true,"name":"Quinci King","retreat":"r1","room":"Room 5","status":"confirmed"},{"id":"r1-raisa","host":false,"name":"Raisa Gates","retreat":"r1","room":"Dorm room 3","status":"confirmed"},{"id":"r1-sharjeel","host":false,"name":"Sharjeel Khan","retreat":"r1","room":"Room 2","status":"confirmed"},{"id":"r1-tessa","host":false,"name":"Tessa Kaye","retreat":"r1","room":"Dorm room 3","status":"confirmed"},{"id":"r1-tristan","host":true,"name":"Tristan Bouan","retreat":"r1","room":"Room 2","status":"confirmed"},{"id":"r1-zai","host":false,"name":"Zai","retreat":"r1","room":"Room 1","status":"confirmed"},{"id":"r2-aida","host":false,"name":"Aida Delpuech","retreat":"r2","status":"confirmed"},{"id":"r2-emmanuel","host":false,"name":"Emmanuel Muhammad","retreat":"r2","status":"confirmed"},{"id":"r2-hannah","host":false,"name":"Hannah Godefa","retreat":"r2","status":"confirmed"},{"id":"r2-jack","host":false,"name":"Jack Wingate","retreat":"r2","status":"confirmed"},{"id":"r2-nicolas","host":false,"name":"Nicolas Block","retreat":"r2","status":"confirmed"},{"id":"r2-quinci","host":true,"name":"Quinci King","retreat":"r2","status":"confirmed"},{"id":"r2-tristan","host":true,"name":"Tristan Bouan","retreat":"r2","status":"confirmed"}],"meals":[{"id":"r1-2026-11-03-dinner","date":"2026-11-03","menu":"Cider-steamed mussels with leeks, shallots, fresh herbs and roasted potatoes","resident":true,"retreat":"r1","shopping":"","slot":"dinner","team":["Christina Alcinay"]},{"id":"r1-2026-11-04-breakfast","date":"2026-11-04","menu":"Haitian-Breton fusion day\nPotato and celeriac hash with caramelized onions, chard and fried eggs","resident":true,"retreat":"r1","shopping":"","slot":"breakfast","team":["Christina Alcinay"]},{"id":"r1-2026-11-04-dinner","date":"2026-11-04","menu":"Haitian epis-roasted chicken with pumpkin, carrots, parsnips, crispy Savoy cabbage and cider pan gravy","resident":true,"retreat":"r1","shopping":"","slot":"dinner","team":["Christina Alcinay"]},{"id":"r1-2026-11-05-breakfast","date":"2026-11-05","menu":"Buttermilk pancakes with warm apple-pear compote and toasted walnuts","resident":true,"retreat":"r1","shopping":"","slot":"breakfast","team":["Christina Alcinay"]},{"id":"r1-2026-11-05-dinner","date":"2026-11-05","menu":"Pan-roasted mackerel with potato-cauliflower purée, pickled red cabbage and apple-cider sauce","resident":true,"retreat":"r1","shopping":"","slot":"dinner","team":["Christina Alcinay"]},{"id":"r1-2026-11-06-breakfast","date":"2026-11-06","menu":"Savory crêpes with roasted pumpkin, wilted spinach, caramelized onions and cheese","resident":true,"retreat":"r1","shopping":"","slot":"breakfast","team":["Christina Alcinay"]},{"id":"r1-2026-11-06-dinner","date":"2026-11-06","menu":"Cider-braised beef with carrots, turnips, onions, mushrooms and creamy potato-celeriac mash","resident":true,"retreat":"r1","shopping":"","slot":"dinner","team":["Christina Alcinay"]},{"id":"r1-2026-11-07-breakfast","date":"2026-11-07","menu":"Haitian-Creole seafood fusion day\nApple and chestnut French toast with caramelized pears and walnut crumble","resident":true,"retreat":"r1","shopping":"","slot":"breakfast","team":["Christina Alcinay"]},{"id":"r1-2026-11-07-dinner","date":"2026-11-07","menu":"Epis-seared scallops with creamy Creole risotto (grits-like), sautéed leeks, crispy shallots, apple pikliz and cider brown-butter sauce","resident":true,"retreat":"r1","shopping":"","slot":"dinner","team":["Christina Alcinay"]},{"id":"r1-2026-11-08-dinner","date":"2026-11-08","menu":"Feast","retreat":"r1","shopping":"","slot":"dinner","team":[]},{"id":"r1-2026-11-09-dinner","date":"2026-11-09","menu":"Feast","retreat":"r1","shopping":"","slot":"dinner","team":[]},{"id":"r1-2026-11-10-breakfast","date":"2026-11-10","menu":"Farewell brunch","retreat":"r1","shopping":"","slot":"breakfast","team":[]}],"workshops":[{"id":"r1-ws-farmtable","date":"","lead":"Christina Alcinay","length":"","notes":"Zai offered to document.","retreat":"r1","status":"confirmed","time":"","title":"Experimental farm-to-table culinary experience"},{"id":"r1-ws-mayor","date":"2026-11-04","lead":"Patrice Gaultier, Mayor of Évran","length":"Right after lunch","notes":"The Mayor comes to the house to welcome residents to the area and co-design programming for the week, such as local producer tours and volunteering at the town's elderly care home (EHPAD).","retreat":"r1","status":"confirmed","time":"","title":"Welcome visit from the Mayor of Évran"},{"id":"r1-ws-producers","date":"","lead":"Mayor of Évran","length":"","notes":"Offered by Patrice Gaultier. Details to be co-designed with the Mayor at his welcome visit on Wed 4 Nov.","retreat":"r1","status":"tbc","time":"","title":"Tour of local producers"},{"id":"r1-ws-senior","date":"","lead":"Tristan Bouan","length":"","notes":"Volunteering at Évran's government-run elderly care home, co-designed with the Mayor, Patrice Gaultier, at his visit on Wed 4 Nov. Likely on a feast and partner day (Sun 8 or Mon 9 Nov). Christina offered to cook; Kayra offered to style guests in traditional Breton looks. Tristan to confirm the date.","retreat":"r1","status":"tbc","time":"","title":"Visit to the local senior home (EHPAD)"},{"id":"r1-ws-stretch","date":"","lead":"Zai","length":"30 min","notes":"","retreat":"r1","status":"confirmed","time":"","title":"Morning stretch"},{"id":"r1-ws-supplies-1","date":"2026-11-03","lead":"Treehaus","length":"","notes":"Optional group trip to the local grocery and art supply stores. Pick your own materials with your $250 supplies budget.","retreat":"r1","status":"confirmed","time":"","title":"Grocery and art supply run"},{"id":"r1-ws-supplies-2","date":"2026-11-06","lead":"Treehaus","length":"","notes":"A second drive to the art supply store for anything you've run out of or newly need.","retreat":"r1","status":"confirmed","time":"","title":"Second art supply run"},{"id":"r1-ws-writing","date":"","lead":"Naomi Lilly","length":"","notes":"Naomi's writing process: pieces that combine personal reflection with cultural analysis.","retreat":"r1","status":"confirmed","time":"","title":"Creative writing for impact"}],"settings":[{"id":"r1","drinks":50,"food":150,"fuelPlan":[{"amount":300,"what":"Diesel for heating, 200 L (deposit)"},{"amount":150,"what":"Gas for the cars and firewood"}],"gas":50,"note":"","payHow":""},{"id":"r2","drinks":50,"food":150,"fuelPlan":[{"amount":225,"what":"Diesel for heating, 150 L (deposit)"},{"amount":112.5,"what":"Gas for the cars and firewood"}],"gas":50,"note":"","payHow":""}],"expenses":[]};
