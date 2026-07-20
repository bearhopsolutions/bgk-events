// BGK Events — server-side (Google Apps Script, bound to the "BGK Events DB" Google Sheet)

var SHEET_USERS = 'Users';
var SHEET_EVENTS = 'Events';
var SHEET_ENROLLMENTS = 'Enrollments';
var SESSION_TTL_MS = 30 * 24 * 60 * 60 * 1000; // 30 days

// This project is now a JSON API only — the frontend lives on GitHub Pages
// and calls this web app's URL with fetch(). See index.html in the repo.

var ROUTES = {
  login: function (p) { return login(p.email, p.pin); },
  validateSession: function (p) { return validateSession(p.token); },
  logout: function (p) { return logout(p.token); },
  forgotPin: function (p) { return forgotPin(p.email); },
  getEventsForUser: function (p) { return getEventsForUser(p.token); },
  enrollInEvent: function (p) { return enrollInEvent(p.token, p.eventId); },
  adminGetEvents: function (p) { return adminGetEvents(p.token); },
  adminSaveEvent: function (p) { return adminSaveEvent(p.token, p.event); },
  adminGetUsers: function (p) { return adminGetUsers(p.token); }
};

function doGet(e) {
  return jsonOutput_({ ok: true, message: 'BGK Events API is running.' });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var handler = ROUTES[body.action];
    if (!handler) throw new Error('Unknown action: ' + body.action);
    var result = handler(body.payload || {});
    return jsonOutput_(result);
  } catch (err) {
    return jsonOutput_({ error: err.message });
  }
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Run this once from the Apps Script editor (select "setup" in the function dropdown, click Run)
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var users = ss.getSheetByName(SHEET_USERS) || ss.insertSheet(SHEET_USERS);
  users.clear();
  users.appendRow(['Email', 'Name', 'PIN', 'IsAdmin']);
  users.appendRow(['kevin412l@hotmail.com', 'Kevin', '041295', true]);

  var events = ss.getSheetByName(SHEET_EVENTS) || ss.insertSheet(SHEET_EVENTS);
  events.clear();
  events.appendRow(['EventID', 'EventName', 'Status', 'Visibility', 'Description']);

  var enrollments = ss.getSheetByName(SHEET_ENROLLMENTS) || ss.insertSheet(SHEET_ENROLLMENTS);
  enrollments.clear();
  enrollments.appendRow(['Email', 'EventID', 'DateEnrolled']);

  SpreadsheetApp.flush();
  return 'Setup complete';
}

// ---------- helpers ----------

function sheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function rowsAsObjects_(sheet) {
  var values = sheet.getDataRange().getValues();
  var headers = values.shift();
  return values
    .filter(function (row) { return row.join('') !== ''; })
    .map(function (row) {
      var obj = {};
      headers.forEach(function (h, i) { obj[h] = row[i]; });
      return obj;
    });
}

function findUserRow_(sheet, email) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).toLowerCase() === String(email).toLowerCase()) {
      return i + 1; // 1-indexed sheet row
    }
  }
  return -1;
}

function makeToken_() {
  return Utilities.getUuid();
}

function saveSession_(token, user) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('session_' + token, JSON.stringify({
    email: user.Email,
    name: user.Name,
    isAdmin: !!user.IsAdmin,
    expires: Date.now() + SESSION_TTL_MS
  }));
}

function getSession_(token) {
  if (!token) return null;
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('session_' + token);
  if (!raw) return null;
  var session = JSON.parse(raw);
  if (Date.now() > session.expires) {
    props.deleteProperty('session_' + token);
    return null;
  }
  return session;
}

function requireSession_(token) {
  var session = getSession_(token);
  if (!session) throw new Error('Not logged in. Please log in again.');
  return session;
}

function requireAdmin_(token) {
  var session = requireSession_(token);
  if (!session.isAdmin) throw new Error('Admins only.');
  return session;
}

function randomPin_() {
  return String(Math.floor(100000 + Math.random() * 900000));
}

// ---------- auth ----------

function login(email, pin) {
  var sheet = sheet_(SHEET_USERS);
  var users = rowsAsObjects_(sheet);
  var user = users.filter(function (u) {
    return String(u.Email).toLowerCase() === String(email).toLowerCase() && String(u.PIN) === String(pin);
  })[0];

  if (!user) {
    return { success: false, message: 'Email or PIN is incorrect.' };
  }

  var token = makeToken_();
  saveSession_(token, user);
  return {
    success: true,
    token: token,
    user: { email: user.Email, name: user.Name, isAdmin: !!user.IsAdmin }
  };
}

function validateSession(token) {
  var session = getSession_(token);
  if (!session) return { valid: false };
  return { valid: true, user: { email: session.email, name: session.name, isAdmin: session.isAdmin } };
}

function logout(token) {
  if (token) PropertiesService.getScriptProperties().deleteProperty('session_' + token);
  return true;
}

function forgotPin(email) {
  var sheet = sheet_(SHEET_USERS);
  var rowIndex = findUserRow_(sheet, email);
  if (rowIndex === -1) {
    // Don't reveal whether the email exists
    return { success: true, message: 'If that email is registered, a new PIN has been sent.' };
  }

  var isAdmin = sheet.getRange(rowIndex, 4).getValue();
  if (isAdmin) {
    return { success: false, message: 'Admin PINs can only be reset directly in the database. Contact the site owner.' };
  }

  var newPin = randomPin_();
  sheet.getRange(rowIndex, 3).setValue(newPin);
  var name = sheet.getRange(rowIndex, 2).getValue();

  MailApp.sendEmail({
    to: email,
    subject: 'Your BGK Events PIN was reset',
    body: 'Hi ' + name + ',\n\nYour new PIN is: ' + newPin + '\n\nUse it to log in at BGK Events.'
  });

  return { success: true, message: 'If that email is registered, a new PIN has been sent.' };
}

// ---------- events / home / nav ----------

function visibleToUser_(event, session, enrollmentsByEvent) {
  if (session.isAdmin) return true;
  if (event.Visibility === 'All') return true;
  if (event.Visibility === 'Enrolled') {
    var enrolled = enrollmentsByEvent[event.EventID] || [];
    return enrolled.indexOf(session.email) !== -1;
  }
  return false; // Hidden
}

function getEventsForUser(token) {
  var session = requireSession_(token);
  var events = rowsAsObjects_(sheet_(SHEET_EVENTS));
  var enrollments = rowsAsObjects_(sheet_(SHEET_ENROLLMENTS));

  var enrollmentsByEvent = {};
  enrollments.forEach(function (e) {
    if (!enrollmentsByEvent[e.EventID]) enrollmentsByEvent[e.EventID] = [];
    enrollmentsByEvent[e.EventID].push(String(e.Email).toLowerCase());
  });

  var mySet = {};
  enrollments.forEach(function (e) {
    if (String(e.Email).toLowerCase() === session.email.toLowerCase()) mySet[e.EventID] = true;
  });

  var visible = events.filter(function (ev) { return visibleToUser_(ev, session, enrollmentsByEvent); });

  var enrolling = visible
    .filter(function (ev) { return ev.Status === 'Enrolling'; })
    .map(function (ev) { return decorate_(ev, mySet); });

  var active = visible
    .filter(function (ev) { return ev.Status === 'Active'; })
    .map(function (ev) { return decorate_(ev, mySet); });

  return { enrolling: enrolling, active: active };
}

function decorate_(ev, mySet) {
  return {
    eventId: ev.EventID,
    name: ev.EventName,
    status: ev.Status,
    isEnrolled: !!mySet[ev.EventID]
  };
}

function enrollInEvent(token, eventId) {
  var session = requireSession_(token);
  var events = rowsAsObjects_(sheet_(SHEET_EVENTS));
  var event = events.filter(function (e) { return String(e.EventID) === String(eventId); })[0];
  if (!event) throw new Error('Event not found.');
  if (event.Status !== 'Enrolling') throw new Error('This event is not open for enrollment.');

  var enrollSheet = sheet_(SHEET_ENROLLMENTS);
  var existing = rowsAsObjects_(enrollSheet);
  var already = existing.some(function (e) {
    return String(e.EventID) === String(eventId) && String(e.Email).toLowerCase() === session.email.toLowerCase();
  });
  if (already) return { success: true };

  enrollSheet.appendRow([session.email, eventId, new Date()]);
  return { success: true };
}

// ---------- admin ----------

function adminGetEvents(token) {
  requireAdmin_(token);
  return rowsAsObjects_(sheet_(SHEET_EVENTS));
}

function adminSaveEvent(token, eventObj) {
  requireAdmin_(token);
  var sheet = sheet_(SHEET_EVENTS);
  var values = sheet.getDataRange().getValues();

  if (!eventObj.EventID) {
    eventObj.EventID = Utilities.getUuid();
    sheet.appendRow([eventObj.EventID, eventObj.EventName, eventObj.Status, eventObj.Visibility, eventObj.Description || '']);
    return eventObj;
  }

  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(eventObj.EventID)) {
      sheet.getRange(i + 1, 1, 1, 5).setValues([[
        eventObj.EventID, eventObj.EventName, eventObj.Status, eventObj.Visibility, eventObj.Description || ''
      ]]);
      return eventObj;
    }
  }
  throw new Error('Event not found.');
}

function adminGetUsers(token) {
  requireAdmin_(token);
  return rowsAsObjects_(sheet_(SHEET_USERS)).map(function (u) {
    return { email: u.Email, name: u.Name, isAdmin: !!u.IsAdmin };
  });
}
