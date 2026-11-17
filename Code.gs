/**
 * SMART MAINTAIN – Gonda LPG Plant
 * Google Sheets + Apps Script backend (V15)
 *
 * Setup:
 * 1. Create a dedicated Google Sheet owned by an approved BPCL account.
 * 2. Open Extensions > Apps Script and paste this complete file into Code.gs.
 * 3. Run setupSmartMaintain() once and approve permissions.
 * 4. Deploy > New deployment > Web app.
 *    Execute as: Me. Access: Anyone (only if approved by BPCL IT).
 * 5. Copy the /exec URL into SMART_MAINTAIN_SCRIPT_URL in index.html.
 *
 * SECURITY:
 * This public-web-app configuration permits writes by anyone who can reach its URL.
 * Do NOT use for confidential or sensitive BPCL information without BPCL IT approval.
 * For production, restrict access to authorized Google Workspace identities and add
 * server-side authorization. Never put mailbox passwords or other secrets in this file.
 */
const DB_SHEET_NAME = 'SMART_MAINTAIN_DB';
const DATA_KEYS = ['assets','contracts','logs','reminderPrefs','plantData'];

function setupSmartMaintain() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this script from the target Google Sheet (Extensions > Apps Script).');
  let sh = ss.getSheetByName(DB_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(DB_SHEET_NAME);
  // Non-destructive initialization: preserve all existing database rows.
  // Create the header only when the tab is completely empty.
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,3).setValues([['KEY','JSON_DATA','UPDATED_AT']]);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1,180);
    sh.setColumnWidth(2,600);
    sh.setColumnWidth(3,220);
  }
  const existing = readRows_(sh);
  const now = new Date();
  const defaults = {
    assets: '[]',
    contracts: '[]',
    logs: '[]',
    reminderPrefs: '{}',
    plantData: '{}'
  };
  const rowsToAdd = [];
  DATA_KEYS.forEach(key => {
    if (!Object.prototype.hasOwnProperty.call(existing, key)) {
      rowsToAdd.push([key, defaults[key], now]);
    }
  });
  if (rowsToAdd.length) {
    sh.getRange(sh.getLastRow()+1,1,rowsToAdd.length,3).setValues(rowsToAdd);
  }
  return 'SMART MAINTAIN database checked. Existing rows were preserved.';
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action !== 'get') return output_({ok:true,service:'SMART MAINTAIN backend'});
  const data = readAll_();
  const result = {ok:true, ...data, updatedAt:new Date().toISOString()};
  const json = JSON.stringify(result);
  if (p.callback && /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(p.callback)) {
    return ContentService.createTextOutput(p.callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return output_(result);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const raw = e && e.postData && e.postData.contents;
    if (!raw) return output_({ok:false,error:'Empty request body'});
    const payload = JSON.parse(raw);
    if (payload.action !== 'sync') return output_({ok:false,error:'Unsupported action'});
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('No active spreadsheet. Run setupSmartMaintain first.');
    let sh = ss.getSheetByName(DB_SHEET_NAME);
    if (!sh) { setupSmartMaintain(); sh = ss.getSheetByName(DB_SHEET_NAME); }
    const current = readRows_(sh);
    const now = new Date();
    DATA_KEYS.forEach(key => {
      if (Object.prototype.hasOwnProperty.call(payload,key)) {
        current[key] = JSON.stringify(payload[key]);
      }
    });
    const rows = [['KEY','JSON_DATA','UPDATED_AT']];
    DATA_KEYS.forEach(key => rows.push([key, current[key] || (key==='assets'||key==='contracts'||key==='logs'?'[]':'{}'), now]));
    sh.clearContents();
    sh.getRange(1,1,rows.length,3).setValues(rows);
    sh.setFrozenRows(1);
    return output_({ok:true,message:'Saved',updatedAt:now.toISOString()});
  } catch (err) {
    return output_({ok:false,error:String(err && err.message || err)});
  } finally {
    lock.releaseLock();
  }
}

function readRows_(sh) {
  const values = sh.getDataRange().getValues();
  const out = {};
  for (let i=1;i<values.length;i++) {
    const key = String(values[i][0] || '');
    if (DATA_KEYS.indexOf(key) >= 0) out[key] = String(values[i][1] || '');
  }
  return out;
}
function readAll_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss && ss.getSheetByName(DB_SHEET_NAME);
  if (!sh) return {assets:[],contracts:[],logs:[],reminderPrefs:{},plantData:{}};
  const rows = readRows_(sh), out = {};
  DATA_KEYS.forEach(key => {
    try { out[key] = rows[key] ? JSON.parse(rows[key]) : (key==='assets'||key==='contracts'||key==='logs'?[]:{}); }
    catch (_) { out[key] = key==='assets'||key==='contracts'||key==='logs'?[]:{}; }
  });
  return out;
}
function output_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
