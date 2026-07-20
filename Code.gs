// BGK Events — server-side (Google Apps Script, bound to the "BGK Events DB" Google Sheet)
// JSON API only — the frontend lives on GitHub Pages and calls this web app's URL with fetch().

var SHEET_USERS = 'Users';
var SHEET_EVENTS = 'Events';
var SHEET_ENROLLMENTS = 'Enrollments';
var SHEET_WL_PARTICIPANTS = 'WeightLossParticipants';
var SHEET_WEIGHINS = 'WeighIns';
var SESSION_TTL_MS = 30 * 24 * 60 * 60 * 1000; // 30 days

var ROUTES = {
  login: function (p) { return login(p.email, p.pin); },
  validateSession: function (p) { return validateSession(p.token); },
  logout: function (p) { return logout(p.token); },
  forgotPin: function (p) { return forgotPin(p.email); },
  getEventsForUser: function (p) { return getEventsForUser(p.token); },
  enrollInEvent: function (p) { return enrollInEvent(p.token, p.eventId); },
  adminGetEvents: function (p) { return adminGetEvents(p.token); },
  adminSaveEvent: function (p) { return adminSaveEvent(p.token, p.event); },
  adminGetUsers: function (p) { return adminGetUsers(p.token); },
  getModuleData: function (p) { return getModuleData(p.token, p.eventId); },
  joinWeightLossChallenge: function (p) { return joinWeightLossChallenge(p.token, p.eventId, p.initialWeight); },
  submitWeighIn: function (p) { return submitWeighIn(p.token, p.eventId, p.week, p.weight, p.photoBase64, p.photoMimeType); },
  adminGetWeighIns: function (p) { return adminGetWeighIns(p.token, p.eventId); },
  adminVerifyWeighIn: function (p) { return adminVerifyWeighIn(p.token, p.eventId, p.email, p.week, p.verified); },
  adminSetPaid: function (p) { return adminSetPaid(p.token, p.eventId, p.email, p.paid); }
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

// Run this once from the Apps Script editor (select "setup" in the function dropdown, click Run).
// NOTE: this clears and re-seeds Users/Events/Enrollments — if you've already added other
// users by hand, note them down before re-running this.
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var users = ss.getSheetByName(SHEET_USERS) || ss.insertSheet(SHEET_USERS);
  users.clear();
  users.appendRow(['Email', 'Name', 'PIN', 'IsAdmin']);
  users.appendRow(['kevin412l@hotmail.com', 'Kevin', '041295', true]);

  var events = ss.getSheetByName(SHEET_EVENTS) || ss.insertSheet(SHEET_EVENTS);
  events.clear();
  events.appendRow(['EventID', 'EventName', 'ModuleType', 'Visibility', 'RegistrationStart', 'RegistrationEnd', 'EventEndDate', 'Description']);

  var enrollments = ss.getSheetByName(SHEET_ENROLLMENTS) || ss.insertSheet(SHEET_ENROLLMENTS);
  enrollments.clear();
  enrollments.appendRow(['Email', 'EventID', 'DateEnrolled']);

  var wlParticipants = ss.getSheetByName(SHEET_WL_PARTICIPANTS) || ss.insertSheet(SHEET_WL_PARTICIPANTS);
  wlParticipants.clear();
  wlParticipants.appendRow(['EventID', 'Email', 'InitialWeight', 'Paid']);

  var weighIns = ss.getSheetByName(SHEET_WEIGHINS) || ss.insertSheet(SHEET_WEIGHINS);
  weighIns.clear();
  weighIns.appendRow(['EventID', 'Email', 'Week', 'Date', 'Weight', 'PhotoFileId', 'PhotoUrl', 'Verified']);

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

function parseDate_(s) {
  return s ? new Date(s + 'T00:00:00') : '';
}

function userNameMap_() {
  var map = {};
  rowsAsObjects_(sheet_(SHEET_USERS)).forEach(function (u) { map[String(u.Email).toLowerCase()] = u.Name; });
  return map;
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

// Status is computed from dates, not stored:
// - past EventEndDate                      -> Archived
// - before RegistrationStart                -> Upcoming (not shown to regular users)
// - between RegistrationStart/End           -> Enrolling
// - otherwise                                -> Active
function computeStatus_(event) {
  var today = new Date();
  var end = event.EventEndDate ? new Date(event.EventEndDate) : null;
  if (end && today > end) return 'Archived';

  var regStart = event.RegistrationStart ? new Date(event.RegistrationStart) : null;
  var regEnd = event.RegistrationEnd ? new Date(event.RegistrationEnd) : null;
  if (regStart && today < regStart) return 'Upcoming';
  if (regStart && regEnd && today >= regStart && today <= regEnd) return 'Enrolling';
  return 'Active';
}

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

  var visible = events
    .map(function (ev) { ev._status = computeStatus_(ev); return ev; })
    .filter(function (ev) { return visibleToUser_(ev, session, enrollmentsByEvent); });

  var enrolling = visible
    .filter(function (ev) { return ev._status === 'Enrolling'; })
    .map(function (ev) { return decorate_(ev, mySet); });

  var active = visible
    .filter(function (ev) { return ev._status === 'Active'; })
    .map(function (ev) { return decorate_(ev, mySet); });

  return { enrolling: enrolling, active: active };
}

function decorate_(ev, mySet) {
  return {
    eventId: ev.EventID,
    name: ev.EventName,
    moduleType: ev.ModuleType,
    status: ev._status,
    isEnrolled: !!mySet[ev.EventID]
  };
}

function enrollUser_(email, eventId) {
  var enrollSheet = sheet_(SHEET_ENROLLMENTS);
  var existing = rowsAsObjects_(enrollSheet);
  var already = existing.some(function (e) {
    return String(e.EventID) === String(eventId) && String(e.Email).toLowerCase() === email.toLowerCase();
  });
  if (already) return;
  enrollSheet.appendRow([email, eventId, new Date()]);
}

function enrollInEvent(token, eventId) {
  var session = requireSession_(token);
  var events = rowsAsObjects_(sheet_(SHEET_EVENTS));
  var event = events.filter(function (e) { return String(e.EventID) === String(eventId); })[0];
  if (!event) throw new Error('Event not found.');
  if (computeStatus_(event) !== 'Enrolling') throw new Error('This event is not open for enrollment.');
  enrollUser_(session.email, eventId);
  return { success: true };
}

// ---------- weight loss challenge module ----------

function getModuleData(token, eventId) {
  var session = requireSession_(token);
  var events = rowsAsObjects_(sheet_(SHEET_EVENTS));
  var event = events.filter(function (e) { return String(e.EventID) === String(eventId); })[0];
  if (!event) throw new Error('Event not found.');
  var status = computeStatus_(event);

  if (event.ModuleType !== 'WeightLoss') {
    return { moduleType: event.ModuleType, name: event.EventName, status: status };
  }

  var participants = rowsAsObjects_(sheet_(SHEET_WL_PARTICIPANTS)).filter(function (p) { return String(p.EventID) === String(eventId); });
  var mine = participants.filter(function (p) { return String(p.Email).toLowerCase() === session.email.toLowerCase(); })[0];
  var weighIns = rowsAsObjects_(sheet_(SHEET_WEIGHINS)).filter(function (w) { return String(w.EventID) === String(eventId); });
  var names = userNameMap_();

  var regEnd = event.RegistrationEnd ? new Date(event.RegistrationEnd) : null;
  var currentWeek = 1;
  if (regEnd) {
    var days = Math.floor((new Date() - regEnd) / (24 * 60 * 60 * 1000));
    currentWeek = Math.max(1, Math.floor(days / 7) + 1);
  }

  var leaderboard = participants.map(function (p) {
    var verified = weighIns
      .filter(function (w) { return String(w.Email).toLowerCase() === String(p.Email).toLowerCase() && w.Verified; })
      .sort(function (a, b) { return Number(b.Week) - Number(a.Week); });
    var latestWeight = verified[0] ? Number(verified[0].Weight) : Number(p.InitialWeight);
    var lbsLost = Number(p.InitialWeight) - latestWeight;
    var pctLost = Number(p.InitialWeight) ? (lbsLost / Number(p.InitialWeight)) * 100 : 0;
    return {
      name: names[String(p.Email).toLowerCase()] || p.Email,
      lbsLost: Math.round(lbsLost * 10) / 10,
      pctLost: Math.round(pctLost * 100) / 100
    };
  }).sort(function (a, b) { return b.pctLost - a.pctLost; });

  var myWeighIns = weighIns
    .filter(function (w) { return String(w.Email).toLowerCase() === session.email.toLowerCase(); })
    .map(function (w) { return { week: w.Week, weight: w.Weight, verified: !!w.Verified }; });

  return {
    moduleType: 'WeightLoss',
    name: event.EventName,
    status: status,
    isParticipant: !!mine,
    myInitialWeight: mine ? mine.InitialWeight : null,
    currentWeek: currentWeek,
    myWeighIns: myWeighIns,
    leaderboard: leaderboard
  };
}

function joinWeightLossChallenge(token, eventId, initialWeight) {
  var session = requireSession_(token);
  var events = rowsAsObjects_(sheet_(SHEET_EVENTS));
  var event = events.filter(function (e) { return String(e.EventID) === String(eventId); })[0];
  if (!event) throw new Error('Event not found.');
  if (event.ModuleType !== 'WeightLoss') throw new Error('Not a weight loss event.');
  if (computeStatus_(event) !== 'Enrolling') throw new Error('This challenge is not open for enrollment.');
  if (!initialWeight || Number(initialWeight) <= 0) throw new Error('Enter a starting weight.');

  enrollUser_(session.email, eventId);

  var sheet = sheet_(SHEET_WL_PARTICIPANTS);
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(eventId) && String(values[i][1]).toLowerCase() === session.email.toLowerCase()) {
      return { success: true };
    }
  }
  sheet.appendRow([eventId, session.email, Number(initialWeight), false]);
  return { success: true };
}

function weighInFolder_(eventId) {
  var props = PropertiesService.getScriptProperties();
  var key = 'wlFolder_' + eventId;
  var folderId = props.getProperty(key);
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (err) { /* fall through and recreate */ }
  }
  var folder = DriveApp.createFolder('BGK Events - WeighIn Photos - ' + eventId);
  props.setProperty(key, folder.getId());
  return folder;
}

function submitWeighIn(token, eventId, week, weight, photoBase64, photoMimeType) {
  var session = requireSession_(token);
  var participants = rowsAsObjects_(sheet_(SHEET_WL_PARTICIPANTS));
  var mine = participants.filter(function (p) {
    return String(p.EventID) === String(eventId) && String(p.Email).toLowerCase() === session.email.toLowerCase();
  })[0];
  if (!mine) throw new Error('You are not enrolled in this challenge.');
  if (!weight || Number(weight) <= 0) throw new Error('Enter a valid weight.');

  var photoUrl = '';
  var photoFileId = '';
  if (photoBase64) {
    var blob = Utilities.newBlob(Utilities.base64Decode(photoBase64), photoMimeType || 'image/jpeg', session.email + '_week' + week + '.jpg');
    var file = weighInFolder_(eventId).createFile(blob);
    photoFileId = file.getId();
    photoUrl = file.getUrl();
  }

  var sheet = sheet_(SHEET_WEIGHINS);
  var values = sheet.getDataRange().getValues();
  var row = [eventId, session.email, Number(week), new Date(), Number(weight), photoFileId, photoUrl, false];
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(eventId) && String(values[i][1]).toLowerCase() === session.email.toLowerCase() && Number(values[i][2]) === Number(week)) {
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return { success: true };
    }
  }
  sheet.appendRow(row);
  return { success: true };
}

function adminGetWeighIns(token, eventId) {
  requireAdmin_(token);
  var names = userNameMap_();
  return rowsAsObjects_(sheet_(SHEET_WEIGHINS))
    .filter(function (r) { return String(r.EventID) === String(eventId); })
    .map(function (r) {
      return {
        email: r.Email,
        name: names[String(r.Email).toLowerCase()] || r.Email,
        week: r.Week,
        weight: r.Weight,
        photoUrl: r.PhotoUrl,
        verified: !!r.Verified
      };
    });
}

function adminVerifyWeighIn(token, eventId, email, week, verified) {
  requireAdmin_(token);
  var sheet = sheet_(SHEET_WEIGHINS);
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(eventId) && String(values[i][1]).toLowerCase() === String(email).toLowerCase() && Number(values[i][2]) === Number(week)) {
      sheet.getRange(i + 1, 8).setValue(!!verified);
      return { success: true };
    }
  }
  throw new Error('Weigh-in not found.');
}

function adminSetPaid(token, eventId, email, paid) {
  requireAdmin_(token);
  var sheet = sheet_(SHEET_WL_PARTICIPANTS);
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(eventId) && String(values[i][1]).toLowerCase() === String(email).toLowerCase()) {
      sheet.getRange(i + 1, 4).setValue(!!paid);
      return { success: true };
    }
  }
  throw new Error('Participant not found.');
}

// ---------- admin ----------

function adminGetEvents(token) {
  requireAdmin_(token);
  return rowsAsObjects_(sheet_(SHEET_EVENTS)).map(function (ev) {
    return {
      EventID: ev.EventID,
      EventName: ev.EventName,
      ModuleType: ev.ModuleType,
      Visibility: ev.Visibility,
      RegistrationStart: ev.RegistrationStart,
      RegistrationEnd: ev.RegistrationEnd,
      EventEndDate: ev.EventEndDate,
      Description: ev.Description,
      computedStatus: computeStatus_(ev)
    };
  });
}

function adminSaveEvent(token, eventObj) {
  requireAdmin_(token);
  var sheet = sheet_(SHEET_EVENTS);
  var values = sheet.getDataRange().getValues();

  var row = [
    eventObj.EventID || Utilities.getUuid(),
    eventObj.EventName || '',
    eventObj.ModuleType || '',
    eventObj.Visibility || 'All',
    parseDate_(eventObj.RegistrationStart),
    parseDate_(eventObj.RegistrationEnd),
    parseDate_(eventObj.EventEndDate),
    eventObj.Description || ''
  ];

  if (!eventObj.EventID) {
    sheet.appendRow(row);
    return { EventID: row[0] };
  }

  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(eventObj.EventID)) {
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return { EventID: row[0] };
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
