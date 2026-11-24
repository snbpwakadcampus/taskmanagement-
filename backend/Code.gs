/**
 * SNBP University – Wakad Campus | Daily Activity Command Centre
 * Google Apps Script backend (Google Sheet = database, Gmail = OTP + reminders)
 *
 * SETUP (one time):
 *  1. Create a new Google Sheet → Extensions → Apps Script → paste this whole file.
 *  2. Edit CONFIG below (ADMIN_EMAIL, DASHBOARD_URL).
 *  3. Select the function `setup` → Run → allow permissions.
 *     (Creates the sheets, the admin account, and the daily 9 AM reminder.)
 *     Your temporary admin password is emailed to ADMIN_EMAIL.
 *  4. Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone → Deploy.
 *  5. Copy the Web app URL into API_URL at the top of index.html.
 */

/* Google Sheet used as the database (ID from its URL). Leave '' if this script is bound to the sheet. */
const SHEET_ID = '1CrXEBN2uW_keKzYbXvBxHMLR8fAjsqc3HX4rms-p8tA';
function book() { return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }

const CONFIG = {
  ADMIN_EMAIL: 'snbpwakadcampus@gmail.com',
  ADMIN_NAME: 'Dr. Atul Kathole',
  ADMIN_DESIGNATION: 'Admin – Strategy & Branding',
  ORG: 'SNBP University · School of Business Management · Wakad Campus',
  DASHBOARD_URL: 'https://snbpwakadcampus.github.io/taskmanagement-/',
  TZ: 'Asia/Kolkata',
  REMINDER_HOUR: 9,          // daily email reminder time (IST)
  WHATSAPP_HOUR: 8,          // daily WhatsApp group update time (IST)
  SESSION_HOURS: 6,          // login session length
  OTP_MINUTES: 10
};

const SHEETS = {
  Users: ['id', 'name', 'email', 'designation', 'department', 'access', 'phone', 'color', 'active', 'salt', 'hash', 'mustChange', 'createdAt', 'lastLogin'],
  Tasks: ['id', 'code', 'activity', 'title', 'assignees', 'priority', 'startDate', 'targetDate', 'closedDate', 'status', 'progress', 'budget', 'remarks', 'createdBy', 'createdAt', 'updatedAt', 'updates'],
  Budget: ['id', 'head', 'description', 'taskId', 'responsible', 'vendor', 'planned', 'approved', 'paid', 'paymentStatus', 'mode', 'dueDate', 'paidDate', 'reference', 'remarks', 'updatedAt', 'updatedBy'],
  Settings: ['key', 'value'],
  Log: ['time', 'user', 'action', 'detail']
};

const DEFAULT_SETTINGS = {
  title: 'Daily Activity Command Centre',
  subtitle: 'Admission & Branding Action Plan',
  publicBudget: 'yes',
  reminderOn: 'yes',
  dueSoonDays: '1',
  adminDigest: 'yes',
  whatsappOn: 'yes',
  whatsappEmail: 'yes'
};

const COLORS = ['#7C3AED', '#E8590C', '#0CA678', '#1C7ED6', '#D6336C', '#F59F00', '#5C7CFA', '#12B886', '#AE3EC9', '#FA5252'];
const STATUSES = ['Not Started', 'In Progress', 'On Hold', 'Completed'];

/* ------------------------------------------------------------------ setup */
function setup() {
  const ss = book();
  Object.keys(SHEETS).forEach(function (name) {
    let s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    s.getRange('A:Z').setNumberFormat('@');
    s.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold').setBackground('#2E1A47').setFontColor('#ffffff');
    s.setFrozenRows(1);
  });
  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  const st = getSettings();
  Object.keys(DEFAULT_SETTINGS).forEach(function (k) { if (st[k] === undefined) setSetting(k, DEFAULT_SETTINGS[k]); });

  const email = CONFIG.ADMIN_EMAIL.toLowerCase();
  if (!rows('Users').some(function (u) { return u.email === email; })) {
    const pw = tempPassword();
    const salt = Utilities.getUuid();
    put('Users', {
      id: newId('U'), name: CONFIG.ADMIN_NAME, email: email, designation: CONFIG.ADMIN_DESIGNATION,
      department: 'School of Business Management', access: 'admin', phone: '', color: COLORS[0], active: 'yes',
      salt: salt, hash: hashPw(pw, salt), mustChange: 'yes', createdAt: nowStr(), lastLogin: ''
    });
    sendMail(email, 'Your admin account is ready', mailWrap('Welcome, ' + esc(CONFIG.ADMIN_NAME),
      '<p>Your <b>Admin</b> account for the Daily Activity Command Centre is ready.</p>' +
      credBox(email, pw) + '<p>Admin login is protected with a one-time code (OTP) sent to this email each time.</p>'));
    Logger.log('Admin created. Temporary password: ' + pw);
  }
  installTrigger();
  Logger.log('Setup complete.');
}

function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyReminder') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'whatsappDaily') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyReminder').timeBased().everyDays(1).atHour(CONFIG.REMINDER_HOUR).inTimezone(CONFIG.TZ).create();
  ScriptApp.newTrigger('whatsappDaily').timeBased().everyDays(1).atHour(CONFIG.WHATSAPP_HOUR).inTimezone(CONFIG.TZ).create();
}

/* ------------------------------------------------------------------ http */
function doGet(e) { return out(handle((e && e.parameter) || {})); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (x) { p = {}; }
  return out(handle(p));
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function handle(p) {
  try {
    const pub = { public: apiPublic, login: apiLogin, verifyOtp: apiVerifyOtp, forgot: apiForgot, resetWithCode: apiResetWithCode };
    if (pub[p.action]) return pub[p.action](p);
    const s = getSession(p.token);
    if (!s) return { ok: false, auth: false, error: 'Session expired. Please log in again.' };
    const me = rows('Users').filter(function (u) { return u.id === s.id; })[0];
    if (!me || me.active !== 'yes') return { ok: false, auth: false, error: 'Account is inactive.' };
    const priv = {
      data: apiData, saveTask: apiSaveTask, updateTask: apiUpdateTask, deleteTask: apiDeleteTask, bulkTasks: apiBulkTasks,
      saveUser: apiSaveUser, deleteUser: apiDeleteUser, resetPassword: apiResetPassword, bulkUsers: apiBulkUsers,
      changePassword: apiChangePassword, saveBudget: apiSaveBudget, deleteBudget: apiDeleteBudget, bulkBudget: apiBulkBudget,
      saveSettings: apiSaveSettings, sendReminders: apiSendReminders, waPreview: apiWaPreview, waSend: apiWaSend, waGroups: apiWaGroups, logout: apiLogout
    };
    if (!priv[p.action]) return { ok: false, error: 'Unknown action' };
    return priv[p.action](p, me);
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}

/* ------------------------------------------------------------------ table helpers */
function sheet(n) { return book().getSheetByName(n); }
function cell(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TZ, 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}
function rows(n) {
  const v = sheet(n).getDataRange().getValues();
  const h = v.shift() || [];
  return v.map(function (r, i) {
    const o = { _r: i + 2 };
    h.forEach(function (k, j) { o[k] = cell(r[j]); });
    return o;
  }).filter(function (o) { return o[h[0]] !== ''; });
}
function toRow(n, o) { return SHEETS[n].map(function (k) { return o[k] === undefined || o[k] === null ? '' : String(o[k]); }); }
function put(n, o, existing) {
  const s = sheet(n);
  const ex = existing || rows(n).filter(function (r) { return r.id === o.id; })[0];
  if (ex) s.getRange(ex._r, 1, 1, SHEETS[n].length).setValues([toRow(n, o)]);
  else s.appendRow(toRow(n, o));
}
function appendMany(n, list) {
  if (!list.length) return;
  const s = sheet(n);
  s.getRange(s.getLastRow() + 1, 1, list.length, SHEETS[n].length).setValues(list.map(function (o) { return toRow(n, o); }));
}
function removeRow(n, id) {
  const ex = rows(n).filter(function (r) { return r.id === id; })[0];
  if (ex) sheet(n).deleteRow(ex._r);
  return !!ex;
}
function withLock(fn) {
  const l = LockService.getScriptLock();
  l.waitLock(25000);
  try { return fn(); } finally { l.releaseLock(); }
}
function newId(p) { return p + Date.now().toString(36).toUpperCase() + Math.floor(Math.random() * 1296).toString(36).toUpperCase(); }
function nowStr() { return Utilities.formatDate(new Date(), CONFIG.TZ, "yyyy-MM-dd'T'HH:mm:ss"); }
function todayStr() { return Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd'); }
function addDays(s, n) { const d = new Date(s + 'T12:00:00'); d.setDate(d.getDate() + n); return Utilities.formatDate(d, CONFIG.TZ, 'yyyy-MM-dd'); }
function clip(v, n) { return String(v === undefined || v === null ? '' : v).slice(0, n || 500); }
function num(v) { const x = Number(String(v || '').replace(/[^0-9.\-]/g, '')); return isNaN(x) ? 0 : x; }
function dateOk(v) { v = clip(v, 10); return /^\d{4}-\d{2}-\d{2}$/.test(v) ? v : ''; }
function esc(s) { return String(s || '').replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
function addLog(user, action, detail) {
  try { sheet('Log').appendRow([nowStr(), user, action, clip(detail, 300)]); } catch (e) { }
}

function getSettings() {
  const o = {};
  rows('Settings').forEach(function (r) { o[r.key] = r.value; });
  return o;
}
function setSetting(k, v) {
  const ex = rows('Settings').filter(function (r) { return r.key === k; })[0];
  if (ex) sheet('Settings').getRange(ex._r, 2).setValue(String(v));
  else sheet('Settings').appendRow([k, String(v)]);
}

/* ------------------------------------------------------------------ security */
function hashPw(pw, salt) {
  let h = salt + '|' + pw;
  for (let i = 0; i < 150; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt));
  }
  return h;
}
function tempPassword() {
  const c = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
  let s = '';
  for (let i = 0; i < 9; i++) s += c.charAt(Math.floor(Math.random() * c.length));
  return s + '@' + Math.floor(10 + Math.random() * 89);
}
function otpCode() { return String(Math.floor(100000 + Math.random() * 900000)); }
function cache() { return CacheService.getScriptCache(); }
function getSession(t) {
  if (!t) return null;
  const v = cache().get('ses_' + t);
  return v ? JSON.parse(v) : null;
}
function newSession(u) {
  const t = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
  cache().put('ses_' + t, JSON.stringify({ id: u.id }), CONFIG.SESSION_HOURS * 3600);
  u.lastLogin = nowStr();
  put('Users', u, u);
  addLog(u.name, 'login', u.access);
  return { ok: true, token: t, user: safeUser(u, true), mustChange: u.mustChange === 'yes' };
}
function failCount(email, add) {
  const k = 'fail_' + email;
  let n = Number(cache().get(k) || 0);
  if (add) { n++; cache().put(k, String(n), 900); }
  return n;
}
function safeUser(u, withEmail) {
  const o = { id: u.id, name: u.name, designation: u.designation, department: u.department, access: u.access, color: u.color, active: u.active };
  if (withEmail) { o.email = u.email; o.phone = u.phone; o.lastLogin = u.lastLogin; o.createdAt = u.createdAt; }
  return o;
}
function isMgr(me) { return me.access === 'admin' || me.access === 'manager'; }
function mask(e) { const p = e.split('@'); return p[0].slice(0, 2) + '***@' + p[1]; }

/* ------------------------------------------------------------------ public */
function apiPublic() {
  const st = getSettings();
  const users = rows('Users').filter(function (u) { return u.active === 'yes'; }).map(function (u) { return safeUser(u, false); });
  const res = { ok: true, settings: publicSettings(st), users: users, tasks: rows('Tasks').map(cleanTask), today: todayStr() };
  res.budget = st.publicBudget === 'yes' ? rows('Budget').map(cleanBudget) : [];
  return res;
}
function publicSettings(st) { return { title: st.title, subtitle: st.subtitle, publicBudget: st.publicBudget, reminderOn: st.reminderOn, dueSoonDays: st.dueSoonDays, adminDigest: st.adminDigest, whatsappOn: st.whatsappOn, whatsappEmail: st.whatsappEmail }; }
function cleanTask(t) { delete t._r; return t; }
function cleanBudget(b) { delete b._r; return b; }

function apiLogin(p) {
  const email = clip(p.email, 120).trim().toLowerCase();
  if (failCount(email) >= 6) return { ok: false, error: 'Too many attempts. Try again after 15 minutes.' };
  const u = rows('Users').filter(function (x) { return x.email === email && x.active === 'yes'; })[0];
  if (!u || hashPw(String(p.password || ''), u.salt) !== u.hash) {
    failCount(email, true);
    addLog(email, 'login-failed', '');
    return { ok: false, error: 'Incorrect email or password.' };
  }
  if (u.access === 'admin') {
    const code = otpCode();
    cache().put('otp_' + email, JSON.stringify({ code: code, tries: 0 }), CONFIG.OTP_MINUTES * 60);
    sendMail(email, 'Admin login code: ' + code, mailWrap('Your admin login code',
      '<p>Use this one-time code to finish signing in:</p><p style="font-size:34px;letter-spacing:8px;font-weight:800;color:#2E1A47;margin:16px 0">' + code + '</p>' +
      '<p>It expires in ' + CONFIG.OTP_MINUTES + ' minutes. If you did not try to log in, change your password now.</p>'));
    return { ok: true, otp: true, sentTo: mask(email) };
  }
  return newSession(u);
}
function apiVerifyOtp(p) {
  const email = clip(p.email, 120).trim().toLowerCase();
  const raw = cache().get('otp_' + email);
  if (!raw) return { ok: false, error: 'Code expired. Please log in again.' };
  const o = JSON.parse(raw);
  if (o.tries >= 5) { cache().remove('otp_' + email); return { ok: false, error: 'Too many wrong codes. Please log in again.' }; }
  if (String(p.code || '').trim() !== o.code) {
    o.tries++;
    cache().put('otp_' + email, JSON.stringify(o), CONFIG.OTP_MINUTES * 60);
    return { ok: false, error: 'Wrong code. ' + (5 - o.tries) + ' tries left.' };
  }
  cache().remove('otp_' + email);
  const u = rows('Users').filter(function (x) { return x.email === email && x.active === 'yes'; })[0];
  if (!u) return { ok: false, error: 'Account not found.' };
  return newSession(u);
}
function apiForgot(p) {
  const email = clip(p.email, 120).trim().toLowerCase();
  const u = rows('Users').filter(function (x) { return x.email === email && x.active === 'yes'; })[0];
  if (u && !cache().get('rst_lock_' + email)) {
    const code = otpCode();
    cache().put('rst_' + email, JSON.stringify({ code: code, tries: 0 }), 900);
    cache().put('rst_lock_' + email, '1', 60);
    sendMail(email, 'Password reset code: ' + code, mailWrap('Reset your password',
      '<p>Your password reset code is</p><p style="font-size:34px;letter-spacing:8px;font-weight:800;color:#2E1A47">' + code + '</p><p>Valid for 15 minutes.</p>'));
  }
  return { ok: true, message: 'If this email is registered, a reset code has been sent.' };
}
function apiResetWithCode(p) {
  const email = clip(p.email, 120).trim().toLowerCase();
  const raw = cache().get('rst_' + email);
  if (!raw) return { ok: false, error: 'Code expired. Request a new one.' };
  const o = JSON.parse(raw);
  if (o.tries >= 5 || String(p.code || '').trim() !== o.code) {
    o.tries++; cache().put('rst_' + email, JSON.stringify(o), 900);
    return { ok: false, error: 'Wrong code.' };
  }
  if (String(p.newPassword || '').length < 6) return { ok: false, error: 'Password must be at least 6 characters.' };
  return withLock(function () {
    const u = rows('Users').filter(function (x) { return x.email === email; })[0];
    u.salt = Utilities.getUuid(); u.hash = hashPw(String(p.newPassword), u.salt); u.mustChange = 'no';
    put('Users', u, u);
    cache().remove('rst_' + email);
    addLog(u.name, 'password-reset', '');
    return { ok: true, message: 'Password changed. Please log in.' };
  });
}

/* ------------------------------------------------------------------ private */
function apiLogout(p) { cache().remove('ses_' + p.token); return { ok: true }; }

function apiData(p, me) {
  const st = getSettings();
  const all = rows('Users');
  const res = {
    ok: true, today: todayStr(), me: safeUser(me, true), settings: publicSettings(st),
    users: all.filter(function (u) { return u.active === 'yes'; }).map(function (u) { return safeUser(u, false); }),
    tasks: rows('Tasks').map(cleanTask), budget: rows('Budget').map(cleanBudget)
  };
  if (me.access === 'admin') {
    res.allUsers = all.map(function (u) { return safeUser(u, true); });
    res.log = rows('Log').slice(-300).reverse().map(function (l) { delete l._r; return l; });
    res.mailQuota = MailApp.getRemainingDailyQuota();
    const pr = PropertiesService.getScriptProperties();
    res.wa = { instance: pr.getProperty('waInstance') || '', token: pr.getProperty('waToken') ? '••••••••' : '', groupId: pr.getProperty('waGroupId') || '', apiUrl: pr.getProperty('waApiUrl') || 'https://api.green-api.com', last: pr.getProperty('waLast') || '' };
  }
  return res;
}

function normTask(t, base, me) {
  const o = base ? JSON.parse(JSON.stringify(base)) : { id: newId('T'), createdBy: me.id, createdAt: nowStr(), updates: '[]' };
  ['code', 'activity', 'title', 'priority', 'remarks'].forEach(function (k) { if (t[k] !== undefined) o[k] = clip(t[k], k === 'remarks' || k === 'title' ? 1000 : 200); });
  ['startDate', 'targetDate', 'closedDate'].forEach(function (k) { if (t[k] !== undefined) o[k] = dateOk(t[k]); });
  if (t.assignees !== undefined) o.assignees = (Array.isArray(t.assignees) ? t.assignees : String(t.assignees).split(',')).map(function (x) { return String(x).trim(); }).filter(Boolean).join(',');
  if (t.status !== undefined) o.status = STATUSES.indexOf(t.status) >= 0 ? t.status : 'Not Started';
  if (t.progress !== undefined) o.progress = String(Math.max(0, Math.min(100, num(t.progress))));
  if (t.budget !== undefined) o.budget = String(num(t.budget));
  if (!o.status) o.status = 'Not Started';
  if (!o.priority) o.priority = 'Medium';
  if (o.status === 'Completed') { if (!o.closedDate) o.closedDate = todayStr(); o.progress = '100'; }
  else if (t.status !== undefined) o.closedDate = '';
  o.updatedAt = nowStr();
  return o;
}
function pushUpdate(o, me, msg) {
  let u = [];
  try { u = JSON.parse(o.updates || '[]'); } catch (e) { u = []; }
  u.push({ t: nowStr(), by: me.name, msg: clip(msg, 400) });
  o.updates = JSON.stringify(u.slice(-40));
}
function canEditTask(t, me) {
  return isMgr(me) || String(t.assignees).split(',').indexOf(me.id) >= 0 || t.createdBy === me.id;
}

function apiSaveTask(p, me) {
  return withLock(function () {
    const t = p.task || {};
    const ex = t.id ? rows('Tasks').filter(function (r) { return r.id === t.id; })[0] : null;
    if (ex && !canEditTask(ex, me)) return { ok: false, error: 'You can edit only your own tasks.' };
    if (!clip(t.title).trim() && !(ex && ex.title)) return { ok: false, error: 'Task title is required.' };
    if (!isMgr(me)) {
      if (!ex) t.assignees = me.id;
      else delete t.assignees;
    }
    const before = ex ? String(ex.assignees).split(',') : [];
    const o = normTask(t, ex, me);
    if (!o.assignees) return { ok: false, error: 'Assign the task to at least one person.' };
    pushUpdate(o, me, ex ? (t.note ? t.note : 'Edited · status ' + o.status) : 'Created');
    put('Tasks', o, ex);
    addLog(me.name, ex ? 'task-edit' : 'task-create', (o.code ? o.code + ' ' : '') + o.title);
    const added = o.assignees.split(',').filter(function (id) { return before.indexOf(id) < 0 && id !== me.id; });
    if (added.length && p.notify !== false) notifyAssigned(added, [o], me);
    delete o._r;
    return { ok: true, task: o };
  });
}
function apiUpdateTask(p, me) {
  return withLock(function () {
    const ex = rows('Tasks').filter(function (r) { return r.id === p.id; })[0];
    if (!ex) return { ok: false, error: 'Task not found.' };
    if (!canEditTask(ex, me)) return { ok: false, error: 'You can update only your own tasks.' };
    const ch = {};
    if (p.status) ch.status = p.status;
    if (p.progress !== undefined && p.progress !== '') ch.progress = p.progress;
    if (p.closedDate) ch.closedDate = p.closedDate;
    const o = normTask(ch, ex, me);
    if (ch.status === 'Completed' && p.closedDate) o.closedDate = dateOk(p.closedDate);
    pushUpdate(o, me, (p.status ? p.status : 'Progress ' + o.progress + '%') + (p.note ? ' — ' + p.note : ''));
    put('Tasks', o, ex);
    addLog(me.name, 'task-update', (o.code || '') + ' ' + o.title + ' → ' + o.status);
    delete o._r;
    return { ok: true, task: o };
  });
}
function apiDeleteTask(p, me) {
  return withLock(function () {
    const ex = rows('Tasks').filter(function (r) { return r.id === p.id; })[0];
    if (!ex) return { ok: false, error: 'Task not found.' };
    if (!(isMgr(me) || ex.createdBy === me.id)) return { ok: false, error: 'Not allowed.' };
    sheet('Tasks').deleteRow(ex._r);
    addLog(me.name, 'task-delete', (ex.code || '') + ' ' + ex.title);
    return { ok: true };
  });
}
function apiBulkTasks(p, me) {
  if (!isMgr(me)) return { ok: false, error: 'Only Admin / Manager can bulk upload.' };
  return withLock(function () {
    const all = (p.tasks || []).slice(0, 1000).map(function (t, i) {
      const o = normTask(t, null, me);
      o.id = o.id + i;
      pushUpdate(o, me, 'Imported from Excel');
      return o;
    });
    const list = all.filter(function (o) { return o.title; });
    appendMany('Tasks', list);
    let budgets = [];
    if (p.budget && p.budget.length) {
      budgets = p.budget.map(function (b) {
        const t = all[Number(b.taskIndex)];
        if (t) b.taskId = t.id;
        return normBudget(b, null, me);
      });
      appendMany('Budget', budgets);
    }
    addLog(me.name, 'bulk-tasks', list.length + ' tasks, ' + budgets.length + ' budget lines');
    if (p.notify !== false) {
      const by = {};
      list.forEach(function (t) { t.assignees.split(',').filter(Boolean).forEach(function (id) { (by[id] = by[id] || []).push(t); }); });
      Object.keys(by).forEach(function (id) { if (id !== me.id) notifyAssigned([id], by[id], me); });
    }
    return { ok: true, count: list.length, budgetCount: budgets.length };
  });
}

/* users */
function apiSaveUser(p, me) {
  if (me.access !== 'admin') return { ok: false, error: 'Only Admin can manage users.' };
  return withLock(function () {
    const u = p.user || {};
    const all = rows('Users');
    const email = clip(u.email, 120).trim().toLowerCase();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return { ok: false, error: 'Valid email required.' };
    const ex = u.id ? all.filter(function (x) { return x.id === u.id; })[0] : null;
    if (all.some(function (x) { return x.email === email && (!ex || x.id !== ex.id); })) return { ok: false, error: 'Email already exists.' };
    const access = ['admin', 'manager', 'member'].indexOf(u.access) >= 0 ? u.access : 'member';
    if (ex && ex.access === 'admin' && (access !== 'admin' || u.active === 'no') &&
      all.filter(function (x) { return x.access === 'admin' && x.active === 'yes'; }).length < 2) return { ok: false, error: 'Keep at least one active admin.' };
    const o = ex ? JSON.parse(JSON.stringify(ex)) : { id: newId('U'), createdAt: nowStr(), color: COLORS[all.length % COLORS.length], salt: Utilities.getUuid(), mustChange: 'yes', lastLogin: '' };
    o.name = clip(u.name, 100).trim(); o.email = email; o.designation = clip(u.designation, 100); o.department = clip(u.department, 100);
    o.access = access; o.phone = clip(u.phone, 20); o.active = u.active === 'no' ? 'no' : 'yes';
    if (u.color) o.color = clip(u.color, 9);
    if (!o.name) return { ok: false, error: 'Name required.' };
    let pw = null;
    if (!ex) { pw = tempPassword(); o.hash = hashPw(pw, o.salt); }
    put('Users', o, ex);
    addLog(me.name, ex ? 'user-edit' : 'user-create', o.name + ' (' + o.access + ')');
    if (pw) welcomeMail(o, pw, me);
    return { ok: true, user: safeUser(o, true) };
  });
}
function apiBulkUsers(p, me) {
  if (me.access !== 'admin') return { ok: false, error: 'Only Admin can manage users.' };
  return withLock(function () {
    const all = rows('Users');
    const seen = {};
    all.forEach(function (x) { seen[x.email] = 1; });
    const list = [];
    (p.users || []).slice(0, 300).forEach(function (u, i) {
      const email = clip(u.email, 120).trim().toLowerCase();
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || seen[email] || !clip(u.name).trim()) return;
      seen[email] = 1;
      const pw = tempPassword();
      const o = {
        id: newId('U') + i, name: clip(u.name, 100).trim(), email: email, designation: clip(u.designation, 100), department: clip(u.department, 100),
        access: ['manager', 'member'].indexOf(u.access) >= 0 ? u.access : 'member', phone: clip(u.phone, 20),
        color: COLORS[(all.length + list.length) % COLORS.length], active: 'yes', salt: Utilities.getUuid(), mustChange: 'yes', createdAt: nowStr(), lastLogin: ''
      };
      o.hash = hashPw(pw, o.salt);
      list.push(o);
      welcomeMail(o, pw, me);
    });
    appendMany('Users', list);
    addLog(me.name, 'bulk-users', list.length + ' users');
    return { ok: true, count: list.length, users: list.map(function (u) { return safeUser(u, true); }) };
  });
}
function apiDeleteUser(p, me) {
  if (me.access !== 'admin') return { ok: false, error: 'Only Admin can manage users.' };
  if (p.id === me.id) return { ok: false, error: 'You cannot delete yourself.' };
  return withLock(function () {
    const ex = rows('Users').filter(function (x) { return x.id === p.id; })[0];
    if (!ex) return { ok: false, error: 'User not found.' };
    sheet('Users').deleteRow(ex._r);
    addLog(me.name, 'user-delete', ex.name);
    return { ok: true };
  });
}
function apiResetPassword(p, me) {
  if (me.access !== 'admin') return { ok: false, error: 'Only Admin can reset passwords.' };
  return withLock(function () {
    const u = rows('Users').filter(function (x) { return x.id === p.id; })[0];
    if (!u) return { ok: false, error: 'User not found.' };
    const pw = tempPassword();
    u.salt = Utilities.getUuid(); u.hash = hashPw(pw, u.salt); u.mustChange = 'yes';
    put('Users', u, u);
    sendMail(u.email, 'Your password was reset', mailWrap('Hello ' + esc(u.name), '<p>' + esc(me.name) + ' reset your password.</p>' + credBox(u.email, pw)));
    addLog(me.name, 'password-reset', u.name);
    return { ok: true, message: 'New password emailed to ' + u.email };
  });
}
function apiChangePassword(p, me) {
  if (hashPw(String(p.oldPassword || ''), me.salt) !== me.hash) return { ok: false, error: 'Current password is wrong.' };
  if (String(p.newPassword || '').length < 6) return { ok: false, error: 'New password must be at least 6 characters.' };
  return withLock(function () {
    me.salt = Utilities.getUuid(); me.hash = hashPw(String(p.newPassword), me.salt); me.mustChange = 'no';
    put('Users', me, me);
    addLog(me.name, 'password-change', '');
    return { ok: true };
  });
}

/* budget */
function normBudget(b, base, me) {
  const o = base ? JSON.parse(JSON.stringify(base)) : { id: newId('B') + Math.floor(Math.random() * 99) };
  ['head', 'description', 'taskId', 'responsible', 'vendor', 'mode', 'reference', 'remarks'].forEach(function (k) { if (b[k] !== undefined) o[k] = clip(b[k], 500); });
  ['planned', 'approved', 'paid'].forEach(function (k) { if (b[k] !== undefined) o[k] = String(num(b[k])); });
  ['dueDate', 'paidDate'].forEach(function (k) { if (b[k] !== undefined) o[k] = dateOk(b[k]); });
  const ps = ['Pending', 'Partially Paid', 'Paid', 'On Hold'];
  if (b.paymentStatus !== undefined) o.paymentStatus = ps.indexOf(b.paymentStatus) >= 0 ? b.paymentStatus : 'Pending';
  if (!o.paymentStatus) o.paymentStatus = 'Pending';
  if (o.paymentStatus === 'Paid' && !o.paidDate) o.paidDate = todayStr();
  o.updatedAt = nowStr(); o.updatedBy = me.name;
  return o;
}
function apiSaveBudget(p, me) {
  return withLock(function () {
    const b = p.item || {};
    const ex = b.id ? rows('Budget').filter(function (r) { return r.id === b.id; })[0] : null;
    if (!isMgr(me) && !(ex && ex.responsible === me.id)) return { ok: false, error: 'Only Admin / Manager or the responsible person can edit this.' };
    if (!isMgr(me)) { delete b.planned; delete b.approved; delete b.responsible; }
    const o = normBudget(b, ex, me);
    if (!o.head && !o.description) return { ok: false, error: 'Budget head or description required.' };
    put('Budget', o, ex);
    addLog(me.name, ex ? 'budget-edit' : 'budget-create', (o.head || '') + ' ' + o.description + ' paid ' + o.paid);
    delete o._r;
    return { ok: true, item: o };
  });
}
function apiDeleteBudget(p, me) {
  if (!isMgr(me)) return { ok: false, error: 'Not allowed.' };
  return withLock(function () { const ok = removeRow('Budget', p.id); addLog(me.name, 'budget-delete', p.id); return { ok: ok }; });
}
function apiBulkBudget(p, me) {
  if (!isMgr(me)) return { ok: false, error: 'Not allowed.' };
  return withLock(function () {
    const list = (p.items || []).slice(0, 1000).map(function (b) { return normBudget(b, null, me); });
    appendMany('Budget', list);
    addLog(me.name, 'bulk-budget', list.length + ' lines');
    return { ok: true, count: list.length };
  });
}

/* settings & reminders */
function apiSaveSettings(p, me) {
  if (me.access !== 'admin') return { ok: false, error: 'Only Admin.' };
  const s = p.settings || {};
  Object.keys(DEFAULT_SETTINGS).forEach(function (k) { if (s[k] !== undefined) setSetting(k, clip(s[k], 200)); });
  const pr = PropertiesService.getScriptProperties();
  ['waInstance', 'waToken', 'waGroupId', 'waApiUrl'].forEach(function (k) { if (s[k] !== undefined && String(s[k]).indexOf('••') < 0) pr.setProperty(k, clip(s[k], 200).trim()); });
  addLog(me.name, 'settings', Object.keys(s).filter(function (k) { return k !== 'waToken'; }).join(', ') + (s.waToken ? ' (+ API token updated)' : ''));
  return { ok: true, settings: publicSettings(getSettings()) };
}
function apiSendReminders(p, me) {
  if (me.access !== 'admin') return { ok: false, error: 'Only Admin.' };
  const n = dailyReminder(true);
  addLog(me.name, 'reminders-manual', n + ' emails');
  return { ok: true, message: n + ' reminder email(s) sent.' };
}

/** Runs every morning (trigger). Emails each person their overdue + due-soon tasks and payments. */
function dailyReminder(force) {
  const st = getSettings();
  if (st.reminderOn !== 'yes' && force !== true) return 0;
  const td = todayStr();
  const soon = addDays(td, Number(st.dueSoonDays || 1));
  const users = rows('Users').filter(function (u) { return u.active === 'yes'; });
  const tasks = rows('Tasks').filter(function (t) { return t.status !== 'Completed' && t.targetDate; });
  const budget = rows('Budget').filter(function (b) { return b.paymentStatus !== 'Paid' && b.dueDate && b.dueDate <= soon; });
  let sent = 0;
  users.forEach(function (u) {
    const mine = tasks.filter(function (t) { return t.assignees.split(',').indexOf(u.id) >= 0; });
    const over = mine.filter(function (t) { return t.targetDate < td; });
    const due = mine.filter(function (t) { return t.targetDate >= td && t.targetDate <= soon; });
    const pay = budget.filter(function (b) { return b.responsible === u.id; });
    if (!u.email || (!over.length && !due.length && !pay.length)) return;
    let body = '<p>Good morning ' + esc(u.name) + ', here is your task status for <b>' + td + '</b>.</p>';
    if (over.length) body += '<h3 style="color:#C92A2A;margin:18px 0 6px">⚠ Overdue (' + over.length + ')</h3>' + taskTable(over, td);
    if (due.length) body += '<h3 style="color:#E8590C;margin:18px 0 6px">Due soon (' + due.length + ')</h3>' + taskTable(due, td);
    if (pay.length) body += '<h3 style="color:#2E1A47;margin:18px 0 6px">Payments to close (' + pay.length + ')</h3>' + payTable(pay);
    body += '<p>Please update the status on the dashboard once done — reminders stop automatically when a task is marked Completed.</p>';
    if (MailApp.getRemainingDailyQuota() > 1) {
      sendMail(u.email, (over.length ? over.length + ' overdue · ' : '') + (due.length ? due.length + ' due soon · ' : '') + 'Daily task reminder', mailWrap('Your pending tasks', body));
      sent++;
    }
  });
  if (st.adminDigest === 'yes') {
    const over = tasks.filter(function (t) { return t.targetDate < td; });
    const admins = users.filter(function (u) { return u.access === 'admin'; });
    if (over.length || budget.length) {
      const names = {};
      users.forEach(function (u) { names[u.id] = u.name; });
      let body = '<p>Team summary for <b>' + td + '</b>: <b>' + over.length + '</b> overdue task(s), <b>' + budget.length + '</b> payment(s) due.</p>';
      if (over.length) body += '<h3 style="color:#C92A2A">Overdue across the team</h3>' + taskTable(over, td, names);
      if (budget.length) body += '<h3 style="color:#2E1A47">Payments due</h3>' + payTable(budget, names);
      admins.forEach(function (a) { if (MailApp.getRemainingDailyQuota() > 1) { sendMail(a.email, 'Admin digest · ' + over.length + ' overdue', mailWrap('Team digest', body)); sent++; } });
    }
  }
  return sent;
}

/* ------------------------------------------------------------------ email */
function sendMail(to, subject, html) {
  if (!to) return;
  try {
    MailApp.sendEmail({ to: to, subject: '[SNBP Wakad] ' + subject, htmlBody: html, name: 'SNBP Wakad Command Centre', replyTo: CONFIG.ADMIN_EMAIL });
  } catch (e) { addLog('system', 'mail-error', to + ' ' + e); }
}
function mailWrap(title, body) {
  return '<div style="font-family:Segoe UI,Arial,sans-serif;background:#F4F1F8;padding:24px">' +
    '<div style="max-width:640px;margin:auto;background:#fff;border-radius:14px;overflow:hidden">' +
    '<div style="background:#2E1A47;color:#fff;padding:20px 24px"><div style="font-size:12px;opacity:.8">' + esc(CONFIG.ORG) + '</div>' +
    '<div style="font-size:20px;font-weight:700;margin-top:4px">' + title + '</div></div>' +
    '<div style="padding:20px 24px;color:#1f1830;font-size:14px;line-height:1.55">' + body +
    '<p style="margin-top:22px"><a href="' + CONFIG.DASHBOARD_URL + '" style="background:#FF6B4A;color:#fff;text-decoration:none;padding:11px 18px;border-radius:8px;font-weight:700">Open dashboard</a></p>' +
    '</div><div style="padding:12px 24px;font-size:11px;color:#7a7090;background:#FAF8FC">Automated message from the Daily Activity Command Centre.</div></div></div>';
}
function credBox(email, pw) {
  return '<table style="background:#F4F1F8;border-radius:10px;padding:12px 16px;margin:12px 0"><tr><td style="padding:4px 12px 4px 0">Login</td><td><b>' + esc(email) + '</b></td></tr>' +
    '<tr><td style="padding:4px 12px 4px 0">Temporary password</td><td><b style="font-family:monospace;font-size:16px">' + esc(pw) + '</b></td></tr></table>' +
    '<p>You will be asked to set your own password after the first login.</p>';
}
function welcomeMail(u, pw, me) {
  sendMail(u.email, 'Your login for the Daily Activity Command Centre', mailWrap('Welcome, ' + esc(u.name),
    '<p>' + esc(me.name) + ' has created your account (' + esc(u.access) + ') so you can see your assigned activities and update them.</p>' + credBox(u.email, pw)));
}
function taskTable(list, td, names) {
  let h = '<table style="width:100%;border-collapse:collapse;font-size:13px"><tr style="background:#F4F1F8;text-align:left"><th style="padding:6px">Task</th>' + (names ? '<th style="padding:6px">Who</th>' : '') + '<th style="padding:6px">Target</th><th style="padding:6px">Late</th></tr>';
  list.sort(function (a, b) { return a.targetDate < b.targetDate ? -1 : 1; }).forEach(function (t) {
    const late = Math.round((new Date(td) - new Date(t.targetDate)) / 86400000);
    h += '<tr style="border-bottom:1px solid #eee"><td style="padding:6px"><b>' + esc(t.code) + '</b> ' + esc(t.title) + '<div style="color:#7a7090;font-size:11px">' + esc(t.activity) + ' · ' + esc(t.status) + '</div></td>' +
      (names ? '<td style="padding:6px">' + esc(t.assignees.split(',').map(function (i) { return names[i] || '?'; }).join(', ')) + '</td>' : '') +
      '<td style="padding:6px;white-space:nowrap">' + t.targetDate + '</td><td style="padding:6px;color:' + (late > 0 ? '#C92A2A' : '#2B8A3E') + ';white-space:nowrap">' + (late > 0 ? late + ' day(s)' : 'on time') + '</td></tr>';
  });
  return h + '</table>';
}
function payTable(list, names) {
  let h = '<table style="width:100%;border-collapse:collapse;font-size:13px"><tr style="background:#F4F1F8;text-align:left"><th style="padding:6px">Item</th>' + (names ? '<th style="padding:6px">Who</th>' : '') + '<th style="padding:6px">Due</th><th style="padding:6px;text-align:right">Balance (₹)</th></tr>';
  list.forEach(function (b) {
    const bal = (num(b.approved) || num(b.planned)) - num(b.paid);
    h += '<tr style="border-bottom:1px solid #eee"><td style="padding:6px"><b>' + esc(b.head) + '</b> ' + esc(b.description) + '<div style="color:#7a7090;font-size:11px">' + esc(b.vendor) + ' · ' + esc(b.paymentStatus) + '</div></td>' +
      (names ? '<td style="padding:6px">' + esc(names[b.responsible] || '—') + '</td>' : '') +
      '<td style="padding:6px">' + b.dueDate + '</td><td style="padding:6px;text-align:right">' + bal.toLocaleString('en-IN') + '</td></tr>';
  });
  return h + '</table>';
}
function notifyAssigned(ids, tasks, me) {
  const users = rows('Users');
  ids.forEach(function (id) {
    const u = users.filter(function (x) { return x.id === id && x.active === 'yes'; })[0];
    if (!u || !u.email) return;
    const td = todayStr();
    sendMail(u.email, tasks.length === 1 ? 'New task: ' + tasks[0].title : tasks.length + ' new tasks assigned to you',
      mailWrap('New task' + (tasks.length > 1 ? 's' : '') + ' for you', '<p>' + esc(me.name) + ' assigned you the following:</p>' + taskTable(tasks, td)));
  });
}


/* ------------------------------------------------------------------ WhatsApp daily group update */
function shortName(n) { return String(n || '').replace(/^(Dr|Mr|Mrs|Ms|Prof)\.?\s+/i, '').split(' ')[0]; }
function buildDailyUpdate() {
  const td = todayStr(), yd = addDays(td, -1);
  const names = {};
  rows('Users').forEach(function (u) { names[u.id] = shortName(u.name); });
  const tasks = rows('Tasks');
  const who = function (t) { return t.assignees.split(',').filter(Boolean).map(function (i) { return names[i] || '?'; }).join('/'); };
  const title = function (t) { return (t.code ? t.code + ' ' : '') + clip(t.title, 50) + (t.title.length > 50 ? '…' : ''); };
  const done = tasks.filter(function (t) { return t.status === 'Completed'; });
  const prog = tasks.filter(function (t) { return t.status === 'In Progress'; });
  const doneY = done.filter(function (t) { return t.closedDate === yd; });
  const today = tasks.filter(function (t) { return t.status !== 'Completed' && t.targetDate === td; });
  const over = tasks.filter(function (t) { return t.status !== 'Completed' && t.targetDate && t.targetDate < td; })
    .sort(function (a, b) { return a.targetDate < b.targetDate ? -1 : 1; });
  const pct = tasks.length ? Math.round(done.length / tasks.length * 100) : 0;
  const bar = '▓'.repeat(Math.round(pct / 10)) + '░'.repeat(10 - Math.round(pct / 10));
  const starts = tasks.map(function (t) { return t.startDate || t.targetDate; }).filter(String).sort();
  const ends = tasks.map(function (t) { return t.targetDate; }).filter(String).sort();
  const d = new Date(td + 'T12:00:00');
  let m = '🌅 *Good morning, Team SNBP Wakad!*\n📅 ' + Utilities.formatDate(d, CONFIG.TZ, 'EEE, d MMM yyyy');
  if (starts.length) {
    const span = Math.round((new Date(ends[ends.length - 1]) - new Date(starts[0])) / 864e5) + 1;
    const day = Math.round((new Date(td) - new Date(starts[0])) / 864e5) + 1;
    if (day >= 1 && day <= span) m += ' · Day ' + day + ' of ' + span;
  }
  m += '\n\n📊 *Progress*  ' + bar + '  *' + pct + '%*\n✅ ' + done.length + ' done · 🔄 ' + prog.length + ' in progress · ⚠️ ' + over.length + ' overdue\n';
  m += '\n🏆 *Done yesterday (' + doneY.length + ')*\n' + (doneY.length ? doneY.slice(0, 5).map(function (t) { return '• ' + title(t) + ' — ' + who(t); }).join('\n') + (doneY.length > 5 ? '\n  +' + (doneY.length - 5) + ' more' : '') : '• No closures logged yesterday') + '\n';
  m += '\n📌 *Today’s focus (' + today.length + ')*\n';
  if (today.length) {
    const by = {};
    today.forEach(function (t) { t.assignees.split(',').filter(Boolean).forEach(function (i) { (by[i] = by[i] || []).push(t); }); });
    Object.keys(by).forEach(function (i) { m += '👤 *' + (names[i] || '?') + '*\n' + by[i].slice(0, 4).map(function (t) { return '   • ' + title(t); }).join('\n') + (by[i].length > 4 ? '\n   +' + (by[i].length - 4) + ' more' : '') + '\n'; });
  } else m += '• Nothing due today — push ahead on upcoming tasks\n';
  if (over.length) {
    m += '\n⚠️ *Overdue (' + over.length + ')*\n' + over.slice(0, 5).map(function (t) {
      const late = Math.round((new Date(td) - new Date(t.targetDate)) / 864e5);
      return '• ' + title(t) + ' — ' + who(t) + ' _(' + late + 'd)_';
    }).join('\n') + (over.length > 5 ? '\n  +' + (over.length - 5) + ' more' : '') + '\n';
  }
  const b = rows('Budget');
  if (b.length) {
    let base = 0, paid = 0;
    b.forEach(function (x) { base += num(x.approved) || num(x.planned); paid += num(x.paid); });
    m += '\n💰 Budget: ₹' + paid.toLocaleString('en-IN') + ' paid of ₹' + base.toLocaleString('en-IN') + '\n';
  }
  m += '\n🔗 Update your tasks: ' + CONFIG.DASHBOARD_URL + '\n_Let’s make today count!_ 💪';
  return m;
}
/** One-time clean-up: removes API tokens that older versions wrote into the Log sheet. */
function redactLogTokens() {
  const sh = sheet('Log');
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][3]).indexOf('waToken') >= 0) sh.getRange(i + 1, 4).setValue('WhatsApp settings saved (token hidden)');
  }
}
function waConfig() {
  const pr = PropertiesService.getScriptProperties();
  return { url: (pr.getProperty('waApiUrl') || 'https://api.green-api.com').replace(/\/$/, ''), id: pr.getProperty('waInstance') || '', token: pr.getProperty('waToken') || '', group: pr.getProperty('waGroupId') || '' };
}
function waPost(method, payload) {
  const c = waConfig();
  if (!c.id || !c.token) throw new Error('WhatsApp gateway not connected yet (add Instance ID and API token in Admin → Settings).');
  const r = UrlFetchApp.fetch(c.url + '/waInstance' + c.id + '/' + method + '/' + c.token, {
    method: payload ? 'post' : 'get', contentType: 'application/json', payload: payload ? JSON.stringify(payload) : undefined, muteHttpExceptions: true
  });
  if (r.getResponseCode() >= 300) throw new Error('WhatsApp gateway error ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 200));
  return JSON.parse(r.getContentText() || '{}');
}
/** Runs every morning at 8 AM (trigger): posts the update to the WhatsApp group, and/or emails it to admins with a one-tap share link. */
function whatsappDaily(force) {
  const st = getSettings();
  if (st.whatsappOn !== 'yes' && force !== true) return 'off';
  const text = buildDailyUpdate();
  const c = waConfig();
  let status = '';
  if (c.id && c.token && c.group) {
    try { waPost('sendMessage', { chatId: c.group, message: text }); status = 'posted to group'; }
    catch (e) { status = 'gateway failed: ' + e.message; }
  } else status = 'gateway not connected';
  if (st.whatsappEmail === 'yes' || status.indexOf('posted') < 0) {
    const share = 'https://wa.me/?text=' + encodeURIComponent(text);
    const body = '<p>Today\'s WhatsApp update' + (status.indexOf('posted') === 0 ? ' was posted to the group automatically.' : ' is ready — tap the button on your phone and choose the team group.') + '</p>' +
      '<pre style="white-space:pre-wrap;background:#F4F1F8;padding:14px;border-radius:10px;font-family:inherit">' + esc(text) + '</pre>' +
      '<p><a href="' + share + '" style="background:#25D366;color:#fff;text-decoration:none;padding:11px 18px;border-radius:8px;font-weight:700">Send on WhatsApp</a></p>';
    rows('Users').filter(function (u) { return u.access === 'admin' && u.active === 'yes'; }).forEach(function (a) { sendMail(a.email, 'WhatsApp daily update · ' + todayStr(), mailWrap('Daily WhatsApp update', body)); });
  }
  PropertiesService.getScriptProperties().setProperty('waLast', nowStr() + ' · ' + status);
  addLog('system', 'whatsapp', status);
  return status;
}
function apiWaPreview(p, me) {
  if (me.access !== 'admin') return { ok: false, error: 'Only Admin.' };
  return { ok: true, text: buildDailyUpdate() };
}
function apiWaSend(p, me) {
  if (me.access !== 'admin') return { ok: false, error: 'Only Admin.' };
  const st = whatsappDaily(true);
  return { ok: st.indexOf('posted') === 0, message: 'WhatsApp: ' + st + '. Copy also emailed to admins.', error: 'WhatsApp: ' + st };
}
/** Lists the WhatsApp groups the linked number belongs to, so the admin can pick the team group. */
function apiWaGroups(p, me) {
  if (me.access !== 'admin') return { ok: false, error: 'Only Admin.' };
  try {
    const list = waPost('getContacts') || [];
    return { ok: true, groups: list.filter(function (c) { return String(c.id || '').indexOf('@g.us') > 0; }).map(function (c) { return { id: c.id, name: c.name || c.contactName || c.id }; }) };
  } catch (e) { return { ok: false, error: String(e.message) }; }
}

/* One-time import of the 22–30 Nov 2026 Action Plan Tracker (safe to re-run: skips if tasks exist). */
function importActionPlan() {
  const PLAN = [["1.1", "Programme Flyer – School of Management", "Design professional flyer covering all School of Management programmes (BBA, BCA, MBA, MCA)", ["Mahesh", "Atul"], "2026-11-24", 0, "Soft copy for digital use and print"], ["1.2", "Programme Flyer – School of Management", "Verify website address, official contact number and email ID; generate QR code", ["Mahesh", "Atul"], "2026-11-24", 0, "Cross-check every detail before release"], ["1.3", "Programme Flyer – School of Management", "Final approval and release of print-ready and social-media versions", ["Vishal"], "2026-11-25", 0, ""], ["2.1", "Digital Marketing & Promotion", "Organic circulation on Facebook, Instagram, LinkedIn and other relevant platforms", ["Atul"], "2026-11-25", 0, "Continue daily till 30 Nov"], ["2.2", "Digital Marketing & Promotion", "Paid promotion – Class XII students and 17–25 age group, 25 km radius around Wakad Campus", ["Atul"], "2026-11-26", 5000, "Total digital budget ₹5,000 up to 30 Nov (incl. Google Ads)"], ["2.3", "Digital Marketing & Promotion", "Explore and run Google Ads within the available budget", ["Atul"], "2026-11-27", 0, "Within ₹5,000 above"], ["2.4", "Digital Marketing & Promotion", "Track reach, clicks and enquiries from each platform", ["Atul"], "2026-11-30", 0, "Share summary report"], ["3.1", "School / Junior College / Coaching Centre Outreach", "Prepare ~15-slide presentation on career options after Class XII + 3 slides on SNBP University, Wakad Campus", ["Mahesh"], "2026-11-25", 0, ""], ["3.2", "School / Junior College / Coaching Centre Outreach", "Prepare 4-minute promotional video on Wakad Campus and its programmes", ["Mahesh"], "2026-11-27", 0, ""], ["3.3", "School / Junior College / Coaching Centre Outreach", "Identify 3–4 schools / junior colleges / coaching centres and obtain permission for a 30-minute session", ["Mahesh"], "2026-11-26", 0, "Record institution, contact person, visit date"], ["3.4", "School / Junior College / Coaching Centre Outreach", "Purchase 100 promotional articles with SBM – SNBP University, Wakad Campus branding (₹150–₹200 each)", ["Mahesh"], "2026-11-27", 20000, "Maximum budget ₹20,000"], ["3.5", "School / Junior College / Coaching Centre Outreach", "Prepare attractive A3 flyer/poster for institutions' notice boards", ["Mahesh"], "2026-11-25", 0, ""], ["3.6", "School / Junior College / Coaching Centre Outreach", "Conduct visits and presentations; distribute articles; request poster display", ["Mahesh"], "2026-11-30", 0, "Minimum 3–4 institutions"], ["4.1", "Corporate & HR Outreach", "Contact at least 3 companies / HR representatives; introduce SNBP brand and MBA programme", ["Mahesh"], "2026-11-27", 0, ""], ["4.2", "Corporate & HR Outreach", "Arrange a visit to at least one organisation to promote the MBA programme", ["Mahesh"], "2026-11-30", 0, ""], ["5.1", "Newspaper Pamphlet Distribution", "Prepare one-page promotional flyer for newspaper insertion", ["Mahesh"], "2026-11-25", 0, ""], ["5.2", "Newspaper Pamphlet Distribution", "Print at least 5,000 pamphlets", ["Mahesh"], "2026-11-27", 0, "Budget as per actual expenses"], ["5.3", "Newspaper Pamphlet Distribution", "Arrange distribution through newspapers", ["Mahesh"], "2026-11-29", 0, "Record areas covered"], ["6.1", "Strategic Planning & Budget Finalisation", "Meeting of Dr. Atul Katole and Mr. Mahesh with Dr. Ketan Desale", ["Atul"], "2026-11-26", 0, ""], ["6.2", "Strategic Planning & Budget Finalisation", "Comprehensive strategic plan with time-wise execution calendar and detailed budget", ["Atul"], "2026-12-01", 0, "Deadline 1 Dec 2026"], ["7.1", "Hoarding & Outdoor Branding", "Identify prominent hoarding locations in Wakad – highways and major roads", ["Avinash"], "2026-11-27", 0, "Primary approval by Dr. Avinash Thakur"], ["7.2", "Hoarding & Outdoor Branding", "Discuss options, obtain quotations and evaluate monthly expenditure", ["Avinash"], "2026-11-30", 0, "Primary approval by Dr. Avinash Thakur; budget after quotations"], ["8.1", "Finalisation of Promotional Agencies", "Finalise engagement with College Wollage – Mr. Nadeem Patel", ["Wadhwa", "Avinash"], "2026-11-30", 60000, "Proposed ₹50,000–₹60,000 for both agencies"], ["8.2", "Finalisation of Promotional Agencies", "Finalise engagement with Digital News – Mr. Vikas Kadam", ["Wadhwa", "Avinash"], "2026-11-30", 0, "Included in 8.1 budget"], ["9.1", "Student Lead Generation", "Identify and collect at least 300 valid prospective student leads", ["Mahesh"], "2026-11-30", 0, "Target: 300 leads"], ["9.2", "Student Lead Generation", "Start calling and follow-up of collected leads", ["Mahesh"], "2026-12-07", 0, "Begins after first week of December"], ["10.1", "CRM Implementation", "Ensure proper implementation of the University's CRM system for Wakad Campus", ["Mahesh"], "2026-11-30", 0, ""], ["11.1", "Office Readiness – Wakad Campus", "Make office area fully ready in all respects; complete essential arrangements", ["Avinash"], "2026-11-30", 0, "Follow-up by Dr. Avinash Thakur; operational before 1 Dec 2026"], ["12.1", "Mobile Purchase", "Mobile Purchase", ["Atul"], "2026-11-30", 0, ""]];
  if (rows('Tasks').length) { Logger.log('Tasks already exist – import skipped.'); return; }
  const people = {
    Atul: { name: CONFIG.ADMIN_NAME },
    Mahesh: { name: 'Mr. Mahesh Karad', designation: 'Admissions Manager', access: 'member' },
    Avinash: { name: 'Dr. Avinash Thakur', designation: 'Director', access: 'manager' },
    Wadhwa: { name: 'Dr. L. K. Wadhwa', designation: 'Advisor', access: 'manager' },
    Vishal: { name: 'Dr. Vishal', designation: 'Flyer approval', access: 'member' }
  };
  const users = rows('Users');
  const ids = {};
  Object.keys(people).forEach(function (k, i) {
    const p = people[k];
    let u = users.filter(function (x) { return x.name === p.name; })[0];
    if (!u) {
      u = { id: newId('U') + i, name: p.name, email: '', designation: p.designation, department: 'School of Business Management', access: p.access, phone: '', color: COLORS[(users.length + i) % COLORS.length], active: 'yes', salt: Utilities.getUuid(), hash: '', mustChange: 'yes', createdAt: nowStr(), lastLogin: '' };
      put('Users', u);
    }
    ids[k] = u.id;
  });
  const me = rows('Users').filter(function (x) { return x.access === 'admin'; })[0];
  const tasks = [], budget = [];
  PLAN.forEach(function (r, i) {
    const o = normTask({ code: r[0], activity: r[1], title: r[2], assignees: r[3].map(function (k) { return ids[k]; }).join(','), startDate: '2026-11-22', targetDate: r[4], status: 'Not Started', priority: 'Medium', budget: r[5], remarks: r[6] }, null, me);
    o.id = o.id + i;
    pushUpdate(o, me, 'Imported from Action Plan Tracker');
    tasks.push(o);
    if (Number(r[5]) > 0) budget.push(normBudget({ head: r[1], description: r[2], taskId: o.id, responsible: ids[r[3][0]], planned: r[5], dueDate: r[4], remarks: r[6] }, null, me));
  });
  appendMany('Tasks', tasks);
  appendMany('Budget', budget);
  setSetting('subtitle', 'Admission & Branding Action Plan · 22–30 Nov 2026');
  addLog(me.name, 'bulk-tasks', tasks.length + ' tasks, ' + budget.length + ' budget lines (Action Plan Tracker)');
  Logger.log('Imported ' + tasks.length + ' tasks and ' + budget.length + ' budget lines.');
}
