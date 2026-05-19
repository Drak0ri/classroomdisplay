// ============================================================
// IESV Kiosk Relay — Apps Script
// Deploy as: Web App, Execute as: Me, Access: Anyone
// Paste into barryjshaw Google account Apps Script
// ============================================================

// ---- CONFIGURATION -----------------------------------------
const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID_HERE'; // Replace after creating sheet
const ADMIN_PIN = '123456'; // Change immediately after setup — 6 digits
const SHARED_SECRET = 'iesv-kiosk-2025'; // Must match in dashboard + wall HTML

// ---- ROOM DEFAULTS -----------------------------------------
const DEFAULT_ROOMS = ['D40', 'D41', 'D42', 'D43', 'D44', 'D45'];

// ============================================================
// MAIN ENTRY POINTS
// ============================================================

function doGet(e) {
  const params = e.parameter;
  const secret = params.secret;
  const action = params.action;

  if (secret !== SHARED_SECRET) {
    return jsonResponse({ error: 'Unauthorised' });
  }

  try {
    switch (action) {
      case 'getState':
        return jsonResponse(getRoomState(params.room));
      case 'getAllStates':
        return jsonResponse(getAllStates());
      case 'getConfig':
        return jsonResponse(getConfig());
      case 'getSchedule':
        return jsonResponse(getSchedule(params.room));
      default:
        return jsonResponse({ error: 'Unknown action' });
    }
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ error: 'Invalid JSON' });
  }

  if (body.secret !== SHARED_SECRET) {
    return jsonResponse({ error: 'Unauthorised' });
  }

  try {
    switch (body.action) {
      case 'verifyAdminPin':
        return jsonResponse({ ok: body.pin === ADMIN_PIN });

      case 'verifyRoomPin':
        return jsonResponse({ ok: verifyRoomPin(body.room, body.pin) });

      case 'setWidget':
        return jsonResponse(setWidget(body.room, body.widget, body.options || {}));

      case 'setCountdown':
        return jsonResponse(setCountdown(body.room, body.data));

      case 'setSoundOptions':
        return jsonResponse(setSoundOptions(body.room, body.data));

      case 'updateConfig':
        if (body.pin !== ADMIN_PIN) return jsonResponse({ error: 'Bad admin PIN' });
        return jsonResponse(updateConfig(body.config));

      case 'uploadSchedule':
        if (body.pin !== ADMIN_PIN && !verifyRoomPin(body.room, body.pin)) return jsonResponse({ error: 'Bad PIN' });
        return jsonResponse(uploadSchedule(body.room, body.schedule));

      case 'setRoomPin':
        if (body.pin !== ADMIN_PIN) return jsonResponse({ error: 'Bad admin PIN' });
        return jsonResponse(setRoomPin(body.room, body.newPin));

      default:
        return jsonResponse({ error: 'Unknown action' });
    }
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ============================================================
// STATE SHEET — "RoomState" tab
// Columns: Room | Widget | CountdownEnd | CountdownPaused | CountdownRemaining
//          | CountdownEndAction | SoundStyle | SoundThreshold | SoundAlerts
//          | SoundAlertMsg | Options (JSON) | Updated
// ============================================================

function getSheet(tabName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    if (tabName === 'RoomState') {
      sheet.appendRow([
        'Room','Widget','CountdownEnd','CountdownPaused','CountdownRemaining',
        'CountdownEndAction','SoundStyle','SoundThreshold','SoundAlerts',
        'SoundAlertMsg','Options','Updated'
      ]);
      DEFAULT_ROOMS.forEach(room => {
        sheet.appendRow([room, 'clock', '', 'false', '', 'clock', 'bar', '65', 'true', 'Too loud!', '{}', new Date().toISOString()]);
      });
    }
    if (tabName === 'Config') {
      sheet.appendRow(['Key', 'Value']);
      DEFAULT_ROOMS.forEach(room => {
        sheet.appendRow([`room_class_${room}`, '']);
        sheet.appendRow([`room_pin_${room}`, '0000']);
        sheet.appendRow([`room_enabled_${room}`, 'true']);
      });
    }
    if (tabName === 'Schedules') {
      sheet.appendRow(['Room', 'ScheduleJSON', 'Updated']);
      DEFAULT_ROOMS.forEach(room => {
        sheet.appendRow([room, '[]', new Date().toISOString()]);
      });
    }
  }
  return sheet;
}

function getRoomRow(sheet, room) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === room) return { row: i + 1, data: data[i] };
  }
  return null;
}

function getRoomState(room) {
  const sheet = getSheet('RoomState');
  const result = getRoomRow(sheet, room);
  if (!result) return { error: 'Room not found' };
  const d = result.data;
  return {
    room: d[0], widget: d[1],
    countdown: {
      end: d[2], paused: d[3] === 'true', remaining: d[4], endAction: d[5]
    },
    sound: {
      style: d[6], threshold: parseInt(d[7]) || 65,
      alerts: d[8] === 'true', alertMsg: d[9]
    },
    options: safeParseJSON(d[10]),
    updated: d[11]
  };
}

function getAllStates() {
  const sheet = getSheet('RoomState');
  const data = sheet.getDataRange().getValues();
  const states = {};
  for (let i = 1; i < data.length; i++) {
    const d = data[i];
    if (!d[0]) continue;
    states[d[0]] = {
      room: d[0], widget: d[1],
      countdown: { end: d[2], paused: d[3] === 'true', remaining: d[4], endAction: d[5] },
      sound: { style: d[6], threshold: parseInt(d[7]) || 65, alerts: d[8] === 'true', alertMsg: d[9] },
      options: safeParseJSON(d[10]),
      updated: d[11]
    };
  }
  return states;
}

function setWidget(room, widget, options) {
  const sheet = getSheet('RoomState');
  const result = getRoomRow(sheet, room);
  if (!result) return { error: 'Room not found' };
  const r = result.row;
  sheet.getRange(r, 2).setValue(widget);
  sheet.getRange(r, 11).setValue(JSON.stringify(options));
  sheet.getRange(r, 12).setValue(new Date().toISOString());
  return { ok: true, room, widget };
}

function setCountdown(room, data) {
  const sheet = getSheet('RoomState');
  const result = getRoomRow(sheet, room);
  if (!result) return { error: 'Room not found' };
  const r = result.row;
  sheet.getRange(r, 2).setValue('countdown');
  sheet.getRange(r, 3).setValue(data.end || '');
  sheet.getRange(r, 4).setValue(data.paused ? 'true' : 'false');
  sheet.getRange(r, 5).setValue(data.remaining || '');
  sheet.getRange(r, 6).setValue(data.endAction || 'clock');
  sheet.getRange(r, 12).setValue(new Date().toISOString());
  return { ok: true };
}

function setSoundOptions(room, data) {
  const sheet = getSheet('RoomState');
  const result = getRoomRow(sheet, room);
  if (!result) return { error: 'Room not found' };
  const r = result.row;
  sheet.getRange(r, 2).setValue('sound');
  sheet.getRange(r, 7).setValue(data.style || 'bar');
  sheet.getRange(r, 8).setValue(data.threshold || 65);
  sheet.getRange(r, 9).setValue(data.alerts ? 'true' : 'false');
  sheet.getRange(r, 10).setValue(data.alertMsg || 'Too loud!');
  sheet.getRange(r, 12).setValue(new Date().toISOString());
  return { ok: true };
}

// ============================================================
// CONFIG SHEET — "Config" tab (key/value)
// ============================================================

function getConfig() {
  const sheet = getSheet('Config');
  const data = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) config[data[i][0]] = data[i][1];
  }
  return config;
}

function updateConfig(newConfig) {
  const sheet = getSheet('Config');
  const data = sheet.getDataRange().getValues();
  const keyMap = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) keyMap[data[i][0]] = i + 1;
  }
  Object.entries(newConfig).forEach(([key, value]) => {
    if (keyMap[key]) {
      sheet.getRange(keyMap[key], 2).setValue(value);
    } else {
      sheet.appendRow([key, value]);
    }
  });
  return { ok: true };
}

function verifyRoomPin(room, pin) {
  const config = getConfig();
  return config[`room_pin_${room}`] === pin;
}

function setRoomPin(room, newPin) {
  return updateConfig({ [`room_pin_${room}`]: newPin });
}

// ============================================================
// SCHEDULES SHEET — "Schedules" tab
// ============================================================

function getSchedule(room) {
  const sheet = getSheet('Schedules');
  const result = getRoomRow(sheet, room);
  if (!result) return { schedule: [] };
  return { schedule: safeParseJSON(result.data[1]) };
}

function uploadSchedule(room, schedule) {
  const sheet = getSheet('Schedules');
  const result = getRoomRow(sheet, room);
  const json = JSON.stringify(schedule);
  if (result) {
    sheet.getRange(result.row, 2).setValue(json);
    sheet.getRange(result.row, 3).setValue(new Date().toISOString());
  } else {
    sheet.appendRow([room, json, new Date().toISOString()]);
  }
  return { ok: true, count: schedule.length };
}

// ============================================================
// HELPERS
// ============================================================

function safeParseJSON(str) {
  try { return JSON.parse(str); } catch { return {}; }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
