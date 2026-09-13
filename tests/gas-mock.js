/**
 * ============================================================
 *  شبیه‌ساز کامل محیط Google Apps Script برای تست سامانه پرمیت
 *  (بدون هیچ وابستگی خارجی — فقط Node.js خالص)
 * ============================================================
 *  این فایل تمام سرویس‌های Apps Script که اسکریپت بک‌اند استفاده
 *  می‌کند را به‌صورت درون‌حافظه‌ای شبیه‌سازی می‌کند:
 *  SpreadsheetApp, DriveApp, PropertiesService, ScriptApp,
 *  LockService, ContentService, HtmlService, Session, Utilities,
 *  Logger, UrlFetchApp, DocumentApp, MimeType
 *
 *  ویژگی کلیدی: «ساعت قابل کنترل» برای شبیه‌سازی سناریوهای
 *  بستن خودکار پرمیت در مرز شیفت‌های روز (۱۹:۰۰) و شب (۰۷:۰۰).
 */
'use strict';
const crypto = require('crypto');

/* ---------------- ساعت شبیه‌سازی‌شده ---------------- */
// همه توابع تاریخ بک‌اند در نهایت به این ساعت وابسته هستند.
const RealDate = Date;
let FAKE_NOW = null; // Date یا null (= زمان واقعی)
function setFakeNow(d) { FAKE_NOW = d ? new RealDate(d.getTime()) : null; }
function nowDate() { return FAKE_NOW ? new RealDate(FAKE_NOW.getTime()) : new RealDate(); }

/**
 * کلاس Date جعلی: هر جا بک‌اند «new Date()» (بدون آرگومان) یا «Date.now()» را
 * صدا بزند، زمان شبیه‌سازی‌شده را می‌گیرد؛ ساخت تاریخ با آرگومان واقعی می‌ماند.
 */
function FakeDate(...args) {
  if (args.length === 0) {
    return FAKE_NOW ? new RealDate(FAKE_NOW.getTime()) : new RealDate();
  }
  switch (args.length) {
    case 1: return new RealDate(args[0]);
    case 2: return new RealDate(args[0], args[1]);
    case 3: return new RealDate(args[0], args[1], args[2]);
    case 4: return new RealDate(args[0], args[1], args[2], args[3]);
    case 5: return new RealDate(args[0], args[1], args[2], args[3], args[4]);
    case 6: return new RealDate(args[0], args[1], args[2], args[3], args[4], args[5]);
    default: return new RealDate(args[0], args[1], args[2], args[3], args[4], args[5], args[6]);
  }
}
FakeDate.now = function () { return FAKE_NOW ? FAKE_NOW.getTime() : RealDate.now(); };
FakeDate.parse = RealDate.parse;
FakeDate.UTC = RealDate.UTC;
FakeDate.prototype = RealDate.prototype;

// تفاوت زمانی هر منطقه (ساعت) — برای سادگی و سازگاری با منطق بک‌اند،
// تست‌ها با TZ=Asia/Tehran اجرا می‌شوند تا مؤلفه‌های محلی Date هم‌راستا باشند.
const TZ_OFFSET_HOURS = {
  'Asia/Tehran': 3.5, 'Iran': 3.5, 'GMT': 0, 'UTC': 0, '': 3.5
};

function _two(n) { n = String(n); return n.length === 1 ? '0' + n : n; }

/* ---------------- Utilities ---------------- */
const Utilities = {
  Charset: { UTF_8: 'UTF-8', US_ASCII: 'US_ASCII' },
  DigestAlgorithm: { SHA_256: 'SHA_256', MD5: 'MD5' },
  base64Decode: function (s) {
    const clean = String(s).replace(/-/g, '+').replace(/_/g, '/');
    return Array.from(Buffer.from(clean, 'base64'));
  },
  base64Encode: function (bytesOrStr) {
    const buf = typeof bytesOrStr === 'string' ? Buffer.from(bytesOrStr, 'utf8') : Buffer.from(bytesOrStr);
    return buf.toString('base64');
  },
  newBlob: function (bytes, mime, name) { return makeBlob(bytes, mime, name); },
  formatDate: function (date, tz, fmt) {
    // date را به منطقه زمانی هدف می‌بریم
    const offsetMin = (TZ_OFFSET_HOURS[tz] != null ? TZ_OFFSET_HOURS[tz] : 3.5) * 60;
    const localOffMin = -date.getTimezoneOffset();
    const shifted = new Date(date.getTime() + (offsetMin - localOffMin) * 60000);
    return String(fmt)
      .replace(/yyyy/g, shifted.getFullYear())
      .replace(/MM/g, _two(shifted.getMonth() + 1))
      .replace(/dd/g, _two(shifted.getDate()))
      .replace(/HH/g, _two(shifted.getHours()))
      .replace(/mm/g, _two(shifted.getMinutes()))
      .replace(/ss/g, _two(shifted.getSeconds()));
  },
  computeDigest: function (algo, str, charset) {
    const a = algo === Utilities.DigestAlgorithm.SHA_256 ? 'sha256' : 'md5';
    const buf = crypto.createHash(a).update(String(str), 'utf8').digest();
    return Array.from(new Int8Array(buf.buffer, buf.byteOffset, buf.length));
  },
  sleep: function () {}
};

function makeBlob(bytes, mime, name) {
  const data = Buffer.from(bytes || []);
  return {
    _data: data, _mime: mime || 'application/octet-stream', _name: name || 'blob',
    getName: function () { return this._name; },
    setName: function (n) { this._name = n; return this; },
    getContentType: function () { return this._mime; },
    getBytes: function () { return Array.from(this._data); },
    getDataAsString: function () { return this._data.toString('utf8'); },
    copyBlob: function () { return makeBlob(Array.from(this._data), this._mime, this._name); },
    getAs: function (mime) {
      // شبیه‌سازی تبدیل (مثلاً HTML به PDF): فقط mime/پسوند عوض می‌شود
      const nm = this._name.replace(/\.[^.]+$/, '') + '.' + String(mime).split('/')[1];
      return makeBlob(Array.from(this._data), mime, nm);
    }
  };
}

/* ---------------- MimeType ---------------- */
const MimeType = {
  JSON: 'application/json', JAVASCRIPT: 'text/javascript', PLAIN_TEXT: 'text/plain',
  HTML: 'text/html', PDF: 'application/pdf', CSV: 'text/csv',
  GOOGLE_SHEETS: 'application/vnd.google-apps.spreadsheet',
  GOOGLE_DOCS: 'application/vnd.google-apps.document'
};

/* ---------------- Logger ---------------- */
const _logs = [];
const Logger = { log: function (s) { _logs.push(String(s)); }, _all: _logs };

/* ---------------- PropertiesService ---------------- */
function makeProps() {
  const store = {};
  return {
    getProperty: (k) => (Object.prototype.hasOwnProperty.call(store, k) ? store[k] : null),
    setProperty: (k, v) => { store[k] = String(v); },
    deleteProperty: (k) => { delete store[k]; },
    getProperties: () => Object.assign({}, store),
    _store: store
  };
}
const _scriptProps = makeProps();
const _userProps = makeProps();
const PropertiesService = {
  getScriptProperties: () => _scriptProps,
  getUserProperties: () => _userProps,
  getDocumentProperties: () => makeProps()
};

/* ---------------- LockService ---------------- */
const LockService = {
  getScriptLock: () => ({ waitLock: () => true, releaseLock: () => {}, tryLock: () => true, hasLock: () => true }),
  getUserLock: () => ({ waitLock: () => true, releaseLock: () => {}, tryLock: () => true, hasLock: () => true })
};

/* ---------------- ContentService / HtmlService ---------------- */
const ContentService = {
  MimeType: { JSON: MimeType.JSON, JAVASCRIPT: MimeType.JAVASCRIPT, TEXT: MimeType.PLAIN_TEXT },
  createTextOutput: function (s) {
    let content = String(s == null ? '' : s);
    let mime = MimeType.PLAIN_TEXT;
    return {
      setContent: function (c) { content = String(c); return this; },
      setMimeType: function (m) { mime = m; return this; },
      getContent: function () { return content; },
      getMimeType: function () { return mime; }
    };
  }
};
const HtmlService = {
  createHtmlOutput: function (html) {
    return {
      _html: String(html),
      setTitle: function () { return this; },
      setWidth: function () { return this; },
      getContent: function () { return this._html; }
    };
  }
};

/* ---------------- Session ---------------- */
let _scriptTimeZone = 'Asia/Tehran';
const Session = {
  getScriptTimeZone: () => _scriptTimeZone,
  getTemporaryActiveUserKey: () => 'TEMP-KEY-1',
  getActiveUser: () => ({ getEmail: () => '' }),
  getEffectiveUser: () => ({ getEmail: () => '' })
};

/* ---------------- ScriptApp (تریگرها) ---------------- */
const _projectTriggers = [];
const ScriptApp = {
  getProjectTriggers: () => _projectTriggers.slice(),
  deleteTrigger: function (t) {
    const i = _projectTriggers.indexOf(t);
    if (i >= 0) _projectTriggers.splice(i, 1);
  },
  newTrigger: function (handler) {
    const t = {
      _handler: handler, _type: null, _minutes: 0, _days: 0, _hour: 0, _eventType: null, _forSS: null,
      getHandlerFunction: function () { return this._handler; },
      getEventType: function () { return this._eventType; },
      timeBased: function () { this._type = 'time'; return this; },
      everyMinutes: function (m) { this._minutes = m; return this; },
      everyHours: function (h) { this._minutes = h * 60; return this; },
      everyDays: function (d) { this._days = d; return this; },
      atHour: function (h) { this._hour = h; return this; },
      forSpreadsheet: function (ss) { this._forSS = ss; return this; },
      onEdit: function () { this._eventType = 'ON_EDIT'; return this; },
      onChange: function () { this._eventType = 'ON_CHANGE'; return this; },
      create: function () { _projectTriggers.push(this); return this; }
    };
    return t;
  },
  AuthMode: { FULL: 'FULL' }
};

/* ---------------- DriveApp ---------------- */
let _driveIdSeq = 1;
function makeDriveFile(name, mime, contentBytesOrString, folder) {
  const isBytes = Array.isArray(contentBytesOrString);
  let data = isBytes ? Buffer.from(contentBytesOrString) : Buffer.from(String(contentBytesOrString == null ? '' : contentBytesOrString), 'utf8');
  const f = {
    _id: 'DRV' + (_driveIdSeq++), _name: name, _mime: mime || MimeType.PLAIN_TEXT,
    _data: data, _trashed: false, _folder: folder || null, _shared: null,
    getId: function () { return this._id; },
    getName: function () { return this._name; },
    setName: function (n) { this._name = n; return this; },
    getMimeType: function () { return this._mime; },
    getUrl: function () { return 'https://drive.google.com/file/d/' + this._id + '/view'; },
    getSize: function () { return this._data.length; },
    getBlob: function () { return makeBlob(Array.from(this._data), this._mime, this._name); },
    setContent: function (s) { this._data = Buffer.from(String(s), 'utf8'); return this; },
    setTrashed: function (t) { this._trashed = !!t; return this; },
    isTrashed: function () { return this._trashed; },
    setSharing: function (access, perm) { this._shared = { access: access, perm: perm }; return this; }
  };
  return f;
}
function makeDriveFolder(name, parent) {
  const folder = {
    _id: 'FLD' + (_driveIdSeq++), _name: name, _parent: parent || null,
    _subfolders: [], _files: [], _trashed: false,
    getId: function () { return this._id; },
    getName: function () { return this._name; },
    getUrl: function () { return 'https://drive.google.com/drive/folders/' + this._id; },
    getParent: function () { return this._parent; },
    setTrashed: function (t) { this._trashed = !!t; },
    createFolder: function (nm) { const f = makeDriveFolder(nm, this); this._subfolders.push(f); return f; },
    getFolders: function () { return iter(this._subfolders.filter(x => !x._trashed)); },
    getFoldersByName: function (nm) { return iter(this._subfolders.filter(x => !x._trashed && x._name === nm)); },
    getFiles: function () { return iter(this._files.filter(x => !x._trashed)); },
    getFilesByName: function (nm) { return iter(this._files.filter(x => !x._trashed && x._name === nm)); },
    createFile: function (a, b, c) {
      let file;
      if (arguments.length === 1 && a && typeof a.getBytes === 'function') {
        file = makeDriveFile(a.getName(), a.getContentType(), a.getBytes(), this);
      } else {
        file = makeDriveFile(a, c || MimeType.PLAIN_TEXT, b, this);
      }
      this._files.push(file);
      return file;
    },
    addFile: function (file) { this._files.push(file); return this; },
    removeFile: function (file) {
      const i = this._files.indexOf(file);
      if (i >= 0) this._files.splice(i, 1);
      return this;
    }
  };
  return folder;
}
function iter(arr) {
  let i = 0;
  return {
    hasNext: function () { return i < arr.length; },
    next: function () { return arr[i++]; }
  };
}
const _rootFolder = makeDriveFolder('My Drive', null);
const _driveRegistry = {}; // id → ss (برای فایل‌های شیت)
const DriveApp = {
  Access: { ANYONE: 'ANYONE', ANYONE_WITH_LINK: 'ANYONE_WITH_LINK' },
  Permission: { VIEW: 'VIEW', EDIT: 'EDIT', COMMENT: 'COMMENT' },
  getRootFolder: function () { return _rootFolder; },
  getFoldersByName: function (nm) {
    // جستجوی بازگشتی در کل درایو (مثل رفتار واقعی برای پوشه‌های ریشه)
    const out = [];
    (function walk(f) {
      f._subfolders.forEach(function (s) {
        if (!s._trashed) { if (s._name === nm) out.push(s); walk(s); }
      });
    })(_rootFolder);
    return iter(out);
  },
  createFolder: function (nm) { return _rootFolder.createFolder(nm); },
  createFile: function (blob) { return _rootFolder.createFile(blob); },
  getFileById: function (id) {
    if (_driveRegistry[id]) return _driveRegistry[id];
    let found = null;
    (function walk(f) {
      f._files.forEach(function (x) { if (!x._trashed && x._id === id) found = x; });
      f._subfolders.forEach(walk);
    })(_rootFolder);
    if (!found) throw new Error('Drive file not found: ' + id);
    return found;
  },
  getFilesByName: function (nm) {
    const out = [];
    (function walk(f) {
      f._files.forEach(function (x) { if (!x._trashed && x._name === nm) out.push(x); });
      f._subfolders.forEach(walk);
    })(_rootFolder);
    return iter(out);
  },
  _root: _rootFolder
};

/* ---------------- SpreadsheetApp ---------------- */
let _ssIdSeq = 1;

function makeRange(sheet, row, col, numRows, numCols) {
  return {
    _sheet: sheet, _row: row, _col: col, _nr: numRows, _nc: numCols,
    getValues: function () {
      const out = [];
      for (let r = 0; r < this._nr; r++) {
        const rowArr = [];
        for (let c = 0; c < this._nc; c++) rowArr.push(this._sheet._cell(this._row + r, this._col + c));
        out.push(rowArr);
      }
      return out;
    },
    getDisplayValues: function () {
      return this.getValues().map(function (r) {
        return r.map(function (v) { return v == null ? '' : String(v); });
      });
    },
    getValue: function () { return this._sheet._cell(this._row, this._col); },
    setValue: function (v) { this._sheet._setCell(this._row, this._col, v); return this; },
    setValues: function (arr) {
      for (let r = 0; r < arr.length; r++)
        for (let c = 0; c < arr[r].length; c++)
          this._sheet._setCell(this._row + r, this._col + c, arr[r][c]);
      return this;
    },
    setBackground: function () { return this; },
    setFontColor: function () { return this; },
    setFontWeight: function () { return this; },
    setFontSize: function () { return this; },
    setHorizontalAlignment: function () { return this; },
    setNumberFormat: function () { return this; },
    setBorder: function () { return this; },
    clear: function () {
      for (let r = 0; r < this._nr; r++)
        for (let c = 0; c < this._nc; c++) this._sheet._setCell(this._row + r, this._col + c, '');
      return this;
    },
    clearContent: function () { return this.clear(); },
    getCell: function (r, c) { return makeRange(this._sheet, this._row + r - 1, this._col + c - 1, 1, 1); },
    getRow: function () { return this._row; },
    getColumn: function () { return this._col; },
    getSheet: function () { return this._sheet; }
  };
}

function makeSheet(ss, name) {
  const sh = {
    _ss: ss, _name: name, _cells: {}, _lastRow: 0, _lastCol: 0, _hidden: false, _frozen: 0,
    getName: function () { return this._name; },
    setName: function (n) { this._name = n; return this; },
    hideSheet: function () { this._hidden = true; return this; },
    isHidden: function () { return this._hidden; },
    setFrozenRows: function (n) { this._frozen = n; },
    setColumnWidth: function () {},
    _key: function (r, c) { return r + ':' + c; },
    _cell: function (r, c) {
      const v = this._cells[this._key(r, c)];
      return v == null ? '' : v;
    },
    _setCell: function (r, c, v) {
      if (v === '' || v == null) { delete this._cells[this._key(r, c)]; }
      else { this._cells[this._key(r, c)] = v; }
      if (r > this._lastRow) {
        // فقط اگر واقعاً مقداری در این ردیف هست، lastRow را بالا ببر
        let hasAny = false;
        for (let cc = 1; cc <= Math.max(this._lastCol, c, 30); cc++) {
          if (this._cells[this._key(r, cc)] != null && this._cells[this._key(r, cc)] !== '') { hasAny = true; break; }
        }
        if (hasAny) this._lastRow = r;
      }
      if (c > this._lastCol) {
        if (v != null && v !== '') this._lastCol = c;
      }
      // اگر سلول خالی شد، ممکن است lastRow کم شود
      this._recalcLasts();
    },
    _recalcLasts: function () {
      let lr = 0, lc = 0;
      for (const k in this._cells) {
        const parts = k.split(':');
        const r = parseInt(parts[0], 10), c = parseInt(parts[1], 10);
        if (r > lr) lr = r;
        if (c > lc) lc = c;
      }
      this._lastRow = lr; this._lastCol = lc;
    },
    getLastRow: function () { return this._lastRow; },
    getLastColumn: function () { return Math.max(this._lastCol, 1); },
    getMaxRows: function () { return Math.max(this._lastRow + 50, 100); },
    getMaxColumns: function () { return Math.max(this._lastCol + 10, 26); },
    appendRow: function (arr) {
      const r = this._lastRow + 1;
      for (let c = 0; c < arr.length; c++) this._setCell(r, c + 1, arr[c]);
      return this;
    },
    deleteRow: function (r) { return this.deleteRows(r, 1); },
    deleteRows: function (r, howMany) {
      howMany = howMany || 1;
      // بازچینی سلول‌ها از ردیف r به بالا
      const newCells = {};
      for (const k in this._cells) {
        const parts = k.split(':');
        const rr = parseInt(parts[0], 10), cc = parseInt(parts[1], 10);
        if (rr < r || rr >= r + howMany) {
          const nr = rr >= r + howMany ? rr - howMany : rr;
          newCells[nr + ':' + cc] = this._cells[k];
        }
      }
      this._cells = newCells;
      this._recalcLasts();
    },
    insertRowAfter: function (r) {
      const newCells = {};
      for (const k in this._cells) {
        const parts = k.split(':');
        const rr = parseInt(parts[0], 10), cc = parseInt(parts[1], 10);
        const nr = rr > r ? rr + 1 : rr;
        newCells[nr + ':' + cc] = this._cells[k];
      }
      this._cells = newCells;
      this._recalcLasts();
    },
    getRange: function (a, b, nr, nc) {
      if (typeof a === 'string') {
        // پشتیبانی از 'A1:D1'
        const m = a.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
        if (!m) throw new Error('range A1notation unsupported: ' + a);
        const c1 = colFromLetters(m[1]), r1 = parseInt(m[2], 10);
        const c2 = colFromLetters(m[3]), r2 = parseInt(m[4], 10);
        return makeRange(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
      }
      return makeRange(this, a, b, nr || 1, nc || 1);
    },
    getDataRange: function () {
      const rows = Math.max(this._lastRow, 1);
      const cols = Math.max(this._lastCol, 1);
      return makeRange(this, 1, 1, rows, cols);
    },
    clear: function () { this._cells = {}; this._lastRow = 0; this._lastCol = 0; return this; },
    clearContents: function () { return this.clear(); },
    getActiveCell: function () { return makeRange(this, 1, 1, 1, 1); },
    getTabColor: function () { return ''; }, setTabColor: function () { return this; }
  };
  return sh;
}
function colFromLetters(s) {
  let n = 0;
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n;
}

function makeSpreadsheet(name) {
  const ss = {
    _id: 'SS' + (_ssIdSeq++), _name: name || 'Untitled spreadsheet', _sheets: [],
    getId: function () { return this._id; },
    getName: function () { return this._name; },
    getUrl: function () { return 'https://docs.google.com/spreadsheets/d/' + this._id + '/edit'; },
    getSheets: function () { return this._sheets.slice(); },
    getSheetByName: function (n) {
      for (let i = 0; i < this._sheets.length; i++) if (this._sheets[i].getName() === n) return this._sheets[i];
      return null;
    },
    insertSheet: function (n) {
      const sh = makeSheet(this, n || ('Sheet' + (this._sheets.length + 1)));
      this._sheets.push(sh);
      return sh;
    },
    deleteSheet: function (sh) {
      const i = this._sheets.indexOf(sh);
      if (i >= 0 && this._sheets.length > 1) this._sheets.splice(i, 1);
    },
    getActiveSheet: function () { return this._sheets[0]; },
    addEditor: function () { return this; },
    rename: function (n) { this._name = n; }
  };
  return ss;
}

let _activeSS = null;
let _uiEnabled = false; // false = زمینه وب‌اپ (بدون UI شیت)
const SpreadsheetApp = {
  getActiveSpreadsheet: function () {
    if (!_activeSS) throw new Error('no active spreadsheet (standalone script)');
    return _activeSS;
  },
  openById: function (id) {
    if (_activeSS && _activeSS.getId() === id) return _activeSS;
    if (_driveRegistry[id]) return _driveRegistry[id];
    throw new Error('Spreadsheet not found: ' + id);
  },
  open: function (file) {
    // file یک شیء فایل درایو است؛ شیت متناظر را از رجیستری می‌گیریم
    if (file && file.getId && _driveRegistry[file.getId()]) return _driveRegistry[file.getId()];
    throw new Error('Cannot open spreadsheet');
  },
  create: function (name) {
    const ss = makeSpreadsheet(name);
    ss.insertSheet('Sheet1');
    // مثل واقعیت: فایل درایو هم ساخته می‌شود
    const df = makeDriveFile(name, MimeType.GOOGLE_SHEETS, '', null);
    df.getId = function () { return ss._id; };
    _driveRegistry[ss._id] = ss;
    _rootFolder._files.push(df);
    return ss;
  },
  getUi: function () {
    if (!_uiEnabled) throw new Error('UI only available in spreadsheet-bound context');
    return makeUi();
  },
  _setActive: function (ss) { _activeSS = ss; },
  _setUi: function (b) { _uiEnabled = b; },
  _registry: _driveRegistry
};

function makeUi() {
  return {
    ButtonSet: { OK: 'OK', YES: 'YES', NO: 'NO', OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO' },
    Button: { OK: 'OK', YES: 'YES', NO: 'NO', CANCEL: 'CANCEL' },
    createMenu: function () {
      return { addItem: function () { return this; }, addSeparator: function () { return this; }, addToUi: function () { return this; } };
    },
    alert: function () { return 'OK'; },
    prompt: function () { return { getSelectedButton: function () { return 'CANCEL'; }, getResponseText: function () { return ''; } }; },
    showSidebar: function () {}
  };
}

/* ---------------- UrlFetchApp ---------------- */
let _urlFetchImpl = null;
const UrlFetchApp = {
  fetch: function (url, options) {
    if (_urlFetchImpl) return _urlFetchImpl(url, options);
    throw new Error('network disabled in tests');
  },
  _setImpl: function (fn) { _urlFetchImpl = fn; }
};

/* ---------------- DocumentApp (برای مسیر پشتیبان ساخت PDF چک‌لیست) ---------------- */
function _docParagraph(text) {
  return {
    _text: text || '',
    setHeading: function () { return this; },
    setAlignment: function () { return this; },
    editAsText: function () { return { setFontSize: function () { return this; }, setBold: function () { return this; } }; },
    appendImage: function () { return { setWidth: function () { return this; }, setHeight: function () { return this; }, getWidth: function () { return 200; }, getHeight: function () { return 80; } }; },
    appendParagraph: function (t) { return _docParagraph(t); }
  };
}
function _docTable(rows) {
  const data = (rows || []).slice();
  return {
    _rows: data,
    setBorderColor: function () { return this; },
    getNumRows: function () { return data.length; },
    getRow: function (i) {
      const cells = data[i] || [];
      return {
        getNumCells: function () { return cells.length; },
        getCell: function (c) { return _docCell(cells[c]); },
        appendTableCell: function () { const cell = _docCell(''); data[i].push(''); return cell; }
      };
    },
    appendTableRow: function () {
      data.push([]);
      return this.getRow(data.length - 1);
    }
  };
}
function _docCell(v) {
  return {
    clear: function () { return this; },
    appendImage: function () { return { setWidth: function () { return this; }, setHeight: function () { return this; }, getWidth: function () { return 200; }, getHeight: function () { return 80; } }; },
    appendParagraph: function (t) { return _docParagraph(t); },
    editAsText: function () { return { setFontSize: function () { return this; }, setBold: function () { return this; } }; },
    setBackgroundColor: function () { return this; },
    _v: v
  };
}
const DocumentApp = {
  ParagraphHeading: { HEADING1: 'H1' },
  HorizontalAlignment: { CENTER: 'CENTER' },
  create: function (name) {
    const body = {
      appendParagraph: function (t) { return _docParagraph(t); },
      appendHorizontalRule: function () { return this; },
      appendTable: function (rows) { return _docTable(rows); }
    };
    return {
      _name: name, _id: 'DOC' + (_driveIdSeq++),
      getBody: function () { return body; },
      getId: function () { return this._id; },
      saveAndClose: function () {
        // تبدیل به فایل درایو که خروجی PDF بدهد
        const df = makeDriveFile(name, MimeType.GOOGLE_DOCS, '', null);
        const realId = this._id;
        df.getId = function () { return realId; };
        df.getAs = function (mime) { return makeBlob([1, 2, 3], mime, name + '.pdf'); };
        _driveRegistry[realId] = df;
        _rootFolder._files.push(df);
      }
    };
  }
};

/* ---------------- CacheService (استفاده احتمالی آینده) ---------------- */
const CacheService = {
  getScriptCache: function () {
    const m = {};
    return {
      get: function (k) { return Object.prototype.hasOwnProperty.call(m, k) ? m[k] : null; },
      put: function (k, v) { m[k] = v; },
      remove: function (k) { delete m[k]; }
    };
  }
};

/* ---------------- خروجی‌گیری و ریست ---------------- */
function resetWorld(opts) {
  opts = opts || {};
  FAKE_NOW = null;
  _logs.length = 0;
  _projectTriggers.length = 0;
  _driveIdSeq = 1;
  _ssIdSeq = 1;
  // پاک‌سازی درایو
  _rootFolder._subfolders.length = 0;
  _rootFolder._files.length = 0;
  for (const k in _driveRegistry) delete _driveRegistry[k];
  for (const k in _scriptProps._store) delete _scriptProps._store[k];
  for (const k in _userProps._store) delete _userProps._store[k];
  // ساخت شیت اصلی متصل (مثل حالت Container-bound)
  if (opts.standalone) {
    _activeSS = null;
  } else {
    _activeSS = makeSpreadsheet('PTW HSE - شبیه‌سازی');
    _activeSS.insertSheet('Sheet1');
  }
  _uiEnabled = !!opts.ui;
  _urlFetchImpl = null;
}

module.exports = {
  SpreadsheetApp, DriveApp, PropertiesService, ScriptApp, LockService,
  ContentService, HtmlService, Session, Utilities, Logger, UrlFetchApp,
  DocumentApp, CacheService, MimeType,
  FakeDate, RealDate,
  setFakeNow, nowDate, resetWorld,
  _internal: { scriptProps: _scriptProps, logs: _logs, triggers: _projectTriggers, rootFolder: _rootFolder, registry: _driveRegistry }
};
