/***********************
 * CONFIG
 ***********************/
const SHEET_ID = '1ouJkXXl6n-dALneQVdJjwdai1QspP_rE1_OOjJKkKq4';
const SHEET_NAME = 'RileyHoods';
const ARCHIVE_SHEET_NAME = 'RileyHoodsArchive';
const RANGE_CACHE_SECONDS = 45;

const CAL_MAP = {
  'TC1-1': '93b3bb9a2622fc1691f59eb488c81344b7d068252b80cf8dc3713037e52c70df@group.calendar.google.com',
  'TC1-2': '93b3bb9a2622fc1691f59eb488c81344b7d068252b80cf8dc3713037e52c70df@group.calendar.google.com',
  'TC1-3': '93b3bb9a2622fc1691f59eb488c81344b7d068252b80cf8dc3713037e52c70df@group.calendar.google.com',
  'TC2-1': '98e45c8ddf4afbfde1cb203c40e3a2ea27e7bc909fc412d7a4a4542ad33004f2@group.calendar.google.com',
  'TC2-2': '98e45c8ddf4afbfde1cb203c40e3a2ea27e7bc909fc412d7a4a4542ad33004f2@group.calendar.google.com',
  'TC2-3': '98e45c8ddf4afbfde1cb203c40e3a2ea27e7bc909fc412d7a4a4542ad33004f2@group.calendar.google.com',
  'TC3-1': '88a5ed82a5f53fd29a13aba89e1b5ad6f44198a32ec99becf88f5ddf8198eefe@group.calendar.google.com',
  'TC3-2': '88a5ed82a5f53fd29a13aba89e1b5ad6f44198a32ec99becf88f5ddf8198eefe@group.calendar.google.com',
  'TC3-3': '88a5ed82a5f53fd29a13aba89e1b5ad6f44198a32ec99becf88f5ddf8198eefe@group.calendar.google.com',
};

const HEADERS = ['id','room','hood','calendarKey','title','start_iso','end_iso','gcal_id','gcal_link','created_by','created_at'];

/***********************
 * HELPERS
 ***********************/
let __spreadsheet = null;
let __sheet = null;

function _ss() {
  if (!__spreadsheet) __spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  if (!__sheet) {
    __sheet = __spreadsheet.getSheetByName(SHEET_NAME);
    if (!__sheet) throw new Error('Tab "' + SHEET_NAME + '" not found.');
  }
  return __sheet;
}

function _nowIso() {
  return new Date().toISOString();
}

function _corsJson(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function _idxFromHeaders(headers) {
  return Object.fromEntries(headers.map((h, i) => [h, i]));
}

function initSheet() {
  const sh = _ss();
  const firstRow = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (firstRow[0] !== 'id') {
    sh.clear();
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
}

function _readTable_() {
  initSheet();
  const sh = _ss();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { sh, headers: HEADERS.slice(), idx: _idxFromHeaders(HEADERS), rows: [] };
  const headers = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const rows = sh.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  return { sh, headers, idx: _idxFromHeaders(headers), rows };
}

function _rangeCacheKey_(start, end) {
  return Utilities.base64EncodeWebSafe(`${start.toISOString()}|${end.toISOString()}`).slice(0, 190);
}

function _clearRangeCache_() {
  // CacheService has no wildcard delete. Versioning invalidates all prior range keys.
  PropertiesService.getScriptProperties().setProperty('rangeCacheVersion', String(Date.now()));
}

function _currentCacheVersion_() {
  const props = PropertiesService.getScriptProperties();
  let version = props.getProperty('rangeCacheVersion');
  if (!version) {
    version = String(Date.now());
    props.setProperty('rangeCacheVersion', version);
  }
  return version;
}

function _eventFromRow_(r, idx) {
  return {
    id: r[idx.id],
    calendarId: r[idx.calendarKey],
    title: r[idx.title],
    start: r[idx.start_iso],
    end: r[idx.end_iso],
  };
}

function _findRowIndexById_(rows, idx, id) {
  return rows.findIndex(r => r[idx.id] === id);
}

function _hasConflictInRows_(rows, idx, calendarKey, start, end, ignoreId) {
  for (const r of rows) {
    if (!r[idx.id]) continue;
    if (r[idx.calendarKey] !== calendarKey) continue;
    if (ignoreId && r[idx.id] === ignoreId) continue;
    const s = new Date(r[idx.start_iso]);
    const t = new Date(r[idx.end_iso]);
    if (s < end && start < t) return true;
  }
  return false;
}

function _withWriteLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok:false, error:'busy' };
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/***********************
 * HTTP HANDLERS
 ***********************/
function doGet(e) {
  try {
    const startStr = e.parameter.start || e.parameter.timeMin;
    const endStr = e.parameter.end || e.parameter.timeMax;
    const start = new Date(startStr);
    const end = new Date(endStr);
    if (isNaN(start) || isNaN(end)) return _corsJson([]);

    const cache = CacheService.getScriptCache();
    const cacheKey = `range:${_currentCacheVersion_()}:${_rangeCacheKey_(start, end)}`;
    const cached = cache.get(cacheKey);
    if (cached) return _corsJson(JSON.parse(cached));

    const { idx, rows } = _readTable_();
    const out = [];
    rows.forEach(r => {
      if (!r[idx.id]) return;
      const s = new Date(r[idx.start_iso]);
      const t = new Date(r[idx.end_iso]);
      if (t <= start || s >= end) return;
      out.push(_eventFromRow_(r, idx));
    });

    try {
      cache.put(cacheKey, JSON.stringify(out), RANGE_CACHE_SECONDS);
    } catch (_) {}
    return _corsJson(out);
  } catch (err) {
    return _corsJson({ ok:false, error:String(err) });
  }
}

function doPost(e) {
  try {
    let body = {};
    if (e.postData && e.postData.contents && e.postData.type && e.postData.type.indexOf('application/json') !== -1) {
      body = JSON.parse(e.postData.contents || '{}');
    } else if (e.parameter && e.parameter.payload) {
      body = JSON.parse(e.parameter.payload);
    } else if (e.postData && e.postData.contents) {
      try { body = JSON.parse(e.postData.contents); } catch (_) { body = {}; }
    }

    const action = body.action;
    const ev = body.event || {};

    if (action === 'create') return _corsJson(_withWriteLock_(() => createBooking_(ev)));
    if (action === 'update') return _corsJson(_withWriteLock_(() => updateBooking_(ev)));
    if (action === 'delete') return _corsJson(_withWriteLock_(() => deleteBooking_(ev.id)));
    if (action === 'createMany') return _corsJson(_withWriteLock_(() => createManyBookings_(body.events || [])));

    return _corsJson({ ok:false, error:'unknown action' });
  } catch (err) {
    return _corsJson({ ok:false, error:String(err) });
  }
}

/***********************
 * CORE CRUD
 ***********************/
function createBooking_(ev) {
  const table = _readTable_();
  const { sh, idx, rows } = table;
  const id = ev.id || Utilities.getUuid();
  const start = new Date(ev.start);
  const end = new Date(ev.end);
  const [room, hood] = (ev.calendarId || '').split('-');

  const rowIndex = _findRowIndexById_(rows, idx, id);
  if (rowIndex >= 0) return updateBooking_({ id, calendarId: ev.calendarId, title: ev.title, start: ev.start, end: ev.end });

  if (_hasConflictInRows_(rows, idx, ev.calendarId, start, end)) return { ok:false, error:'conflict' };

  const gcalId = CAL_MAP[ev.calendarId];
  if (!gcalId) return { ok:false, error:`No calendar for ${ev.calendarId}` };
  const cal = CalendarApp.getCalendarById(gcalId);
  if (!cal) return { ok:false, error:`Calendar not found ${gcalId}` };
  const ge = cal.createEvent(ev.title || `${room}/${hood} booking`, start, end);

  const now = _nowIso();
  sh.appendRow([
    id, room, hood, ev.calendarId, ev.title || '',
    start.toISOString(), end.toISOString(),
    ge.getId(), ge.getHtmlLink ? ge.getHtmlLink() : '',
    Session.getActiveUser().getEmail(), now
  ]);
  _clearRangeCache_();

  return { ok:true, id, gcal_id: ge.getId() };
}

function createManyBookings_(events) {
  if (!Array.isArray(events) || !events.length) return { ok:false, error:'no events' };
  const saved = [];
  for (const ev of events) {
    const result = createBooking_(ev);
    if (!result || result.ok !== true) return { ok:false, error:result && result.error || 'failed', saved };
    saved.push({ ...ev, id: result.id });
  }
  return { ok:true, saved };
}

function updateBooking_(ev) {
  const table = _readTable_();
  const { sh, idx, rows } = table;
  const rowIndex = _findRowIndexById_(rows, idx, ev.id);
  if (rowIndex < 0) return { ok:false, error:'not found' };

  const r = rows[rowIndex];
  const oldCalKey = r[idx.calendarKey];
  const newCalKey = ev.calendarId || oldCalKey;
  const start = new Date(ev.start || r[idx.start_iso]);
  const end = new Date(ev.end || r[idx.end_iso]);
  const title = ev.title ?? r[idx.title];

  if (_hasConflictInRows_(rows, idx, newCalKey, start, end, ev.id)) return { ok:false, error:'conflict' };

  const oldGId = r[idx.gcal_id];
  const sameCal = newCalKey === oldCalKey;
  let ge = null;
  const oldCalId = CAL_MAP[oldCalKey];
  if (oldCalId) {
    const oldCal = CalendarApp.getCalendarById(oldCalId);
    if (oldCal) {
      try { ge = oldCal.getEventById(oldGId); } catch (_) {}
    }
  }
  if (!ge) {
    try { ge = CalendarApp.getEventById(oldGId); } catch (_) {}
  }

  if (ge && sameCal) {
    ge.setTime(start, end);
    ge.setTitle(title);
  } else {
    if (ge) ge.deleteEvent();
    const cal = CalendarApp.getCalendarById(CAL_MAP[newCalKey]);
    if (!cal) return { ok:false, error:`Calendar not found ${newCalKey}` };
    ge = cal.createEvent(title, start, end);
  }

  const [room, hood] = String(newCalKey).split('-');
  const row = rowIndex + 2;
  sh.getRange(row, 1, 1, HEADERS.length).setValues([[
    ev.id,
    room,
    hood,
    newCalKey,
    title,
    start.toISOString(),
    end.toISOString(),
    ge.getId(),
    ge.getHtmlLink ? ge.getHtmlLink() : '',
    r[idx.created_by],
    r[idx.created_at],
  ]]);
  _clearRangeCache_();

  return { ok:true };
}

function deleteBooking_(id) {
  const table = _readTable_();
  const { sh, idx, rows } = table;
  const rowIndex = _findRowIndexById_(rows, idx, id);
  if (rowIndex < 0) return { ok:false, error:'not found' };

  const r = rows[rowIndex];
  const calKey = r[idx.calendarKey];
  const calId = CAL_MAP[calKey];

  try {
    let ge = null;
    if (calId) {
      const cal = CalendarApp.getCalendarById(calId);
      if (cal) {
        try { ge = cal.getEventById(r[idx.gcal_id]); } catch (_) {}
      }
    }
    if (!ge) {
      try { ge = CalendarApp.getEventById(r[idx.gcal_id]); } catch (_) {}
    }
    if (!ge && calId) {
      const cal = CalendarApp.getCalendarById(calId);
      if (cal) {
        const s = new Date(r[idx.start_iso]);
        const e = new Date(r[idx.end_iso]);
        const candidates = cal.getEvents(s, e, { search:r[idx.title] || '' });
        ge = candidates.find(event => event.getId() === r[idx.gcal_id]) || (candidates.length === 1 ? candidates[0] : null);
      }
    }
    if (ge) ge.deleteEvent();
  } catch (_) {}

  sh.deleteRow(rowIndex + 2);
  _clearRangeCache_();
  return { ok:true };
}

/***********************
 * MAINTENANCE
 ***********************/
function archiveBookingsBefore_(cutoffIso) {
  const cutoff = new Date(cutoffIso);
  if (isNaN(cutoff)) throw new Error('Invalid cutoff date');

  const table = _readTable_();
  const { sh, idx, rows } = table;
  const archive = __spreadsheet.getSheetByName(ARCHIVE_SHEET_NAME) || __spreadsheet.insertSheet(ARCHIVE_SHEET_NAME);
  if (archive.getLastRow() === 0) archive.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

  const toArchive = [];
  const rowsToDelete = [];
  rows.forEach((r, index) => {
    const end = new Date(r[idx.end_iso]);
    if (r[idx.id] && !isNaN(end) && end < cutoff) {
      toArchive.push(r);
      rowsToDelete.push(index + 2);
    }
  });

  if (toArchive.length) {
    archive.getRange(archive.getLastRow() + 1, 1, toArchive.length, HEADERS.length).setValues(toArchive);
    rowsToDelete.reverse().forEach(row => sh.deleteRow(row));
    _clearRangeCache_();
  }
  return { archived:toArchive.length };
}

function archiveBookingsOlderThan90Days() {
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 90);
  return archiveBookingsBefore_(cutoff.toISOString());
}

function _firstAuth() {
  _ss().getLastRow();
  CalendarApp.getAllCalendars();
  LockService.getScriptLock().tryLock(1);
  CacheService.getScriptCache().put('auth', 'ok', 1);
  PropertiesService.getScriptProperties().setProperty('auth', 'ok');
}
