/**
 * HPP Seva Connect — Sabha attendance (marking, live counts, reports, analytics).
 * Merged from HP-Adajan-Attendance: behaviour aligned with Sabhas + Attendance sheets.
 */

/**
 * @param {Object} user
 * @return {boolean}
 */
function _attendanceRoleCanMark_(user) {
  if (!user) return false;
  var r = String(user.role || '').toLowerCase();
  return (
    r === ROLES.ADMIN ||
    r === ROLES.LEADER ||
    r === ROLES.SUB_LEADER ||
    r === ROLES.KARYAKARTA
  );
}

/**
 * @param {Object} user
 * @return {string}
 */
function _defaultMandalForAttendance_(user) {
  var mid = String(user.mandalID || '').trim();
  if (mid) return mid;
  if (String(user.role || '').toLowerCase() === ROLES.ADMIN) {
    try {
      var sh = getSheet(SHEET_MANDALS);
      var lr = sh.getLastRow();
      if (lr >= 2) {
        return String(sh.getRange(2, MANDALS_COL.MandalID + 1).getValue() || '').trim();
      }
    } catch (e) {
      logError('_defaultMandalForAttendance_', {}, e);
    }
  }
  return '';
}

/**
 * @param {string} dateStr yyyy-MM-dd
 * @param {Object} user
 * @return {Object|null} { rowIndex, id, date, title, notes, mandalID }
 */
function _findVisibleSabhaForDate_(dateStr, user) {
  var sheet = getSheet(SHEET_SABHAS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var numCols = SABHAS_COL.CreatedOn + 1;
  var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var i;
  for (i = 0; i < data.length; i++) {
    var r = data[i];
    if (!_dashboardSabhaVisible(r, user)) continue;
    var cell = r[SABHAS_COL.Date];
    var ds = cell instanceof Date && !isNaN(cell.getTime()) ? formatDate(cell) : '';
    if (ds === dateStr) {
      return {
        rowIndex: i + 2,
        id: String(r[SABHAS_COL.SabhaID] || '').trim(),
        date: ds,
        title: String(r[SABHAS_COL.Title] || '').trim(),
        notes: String(r[SABHAS_COL.Notes] || '').trim(),
        mandalID: String(r[SABHAS_COL.MandalID] || '').trim(),
      };
    }
  }
  return null;
}

/**
 * @param {Array} attRow
 * @return {boolean}
 */
function _attendanceRowIsPresentOrLate_(attRow) {
  var st = String(attRow[ATTENDANCE_COL.Status] || '').toLowerCase();
  return st === ATTENDANCE_STATUS.PRESENT || st === ATTENDANCE_STATUS.LATE;
}

/**
 * @param {Object} user
 * @return {Array<Object>}
 */
function _getAttendanceDevoteeSummaries_(user) {
  var values = _getDevoteesRowsCached_();
  var out = [];
  var i;
  for (i = 0; i < values.length; i++) {
    var row = values[i];
    if (!_devoteeRowVisibleForRole(row, user)) continue;
    var st = String(row[DEVOTEES_COL.Status] || '').toLowerCase();
    if (st && st !== DEVOTEE_STATUS.ACTIVE) continue;
    out.push({
      id: String(row[DEVOTEES_COL.DevoteeID] || '').trim(),
      firstName: String(row[DEVOTEES_COL.FirstName] || '').trim(),
      middleName: String(row[DEVOTEES_COL.MiddleName] || '').trim(),
      lastName: String(row[DEVOTEES_COL.LastName] || '').trim(),
      gender: String(row[DEVOTEES_COL.Gender] || '').trim() || 'Male',
      mobile: String(row[DEVOTEES_COL.Mobile] || '').trim(),
      birthday: row[DEVOTEES_COL.DOB] instanceof Date && !isNaN(row[DEVOTEES_COL.DOB].getTime())
        ? formatDate(row[DEVOTEES_COL.DOB])
        : '',
      karyakarta: String(row[DEVOTEES_COL.KaryakartaID] || '').trim(),
    });
  }
  return out;
}

/**
 * Initial payload for the attendance page (sabha for date + list + present ids).
 * @param {string} dateStr yyyy-MM-dd
 * @return {{ success: boolean, data?: Object, error?: string }}
 */
function getAttendanceBootstrap(dateStr) {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');
    if (!_attendanceRoleCanMark_(user)) {
      return errorResponse('Only admins, leaders, sub-leaders, and karyakartas can mark attendance.');
    }
    var ds = String(dateStr || '').trim();
    if (!ds) return errorResponse('Date is required');

    var sabhaRes = getOrCreateSabhaForDateAttendance(ds);
    if (!sabhaRes || !sabhaRes.success || !sabhaRes.data) return sabhaRes;

    var sabha = sabhaRes.data;
    var presentRes = getSabhaPresentDevoteeIds(sabha.id);
    if (!presentRes || !presentRes.success) return presentRes;

    var devotees = _getAttendanceDevoteeSummaries_(user);
    return successResponse({
      sabha: sabha,
      presentIds: presentRes.data || [],
      devotees: devotees,
    });
  } catch (e) {
    logError('getAttendanceBootstrap', { dateStr: dateStr }, e);
    return errorResponse(e.message || String(e));
  }
}

/**
 * Find or create a Sabha row for the given calendar date (visible mandal).
 * @param {string} dateStr yyyy-MM-dd
 * @return {{ success: boolean, data?: Object, error?: string }}
 */
function getOrCreateSabhaForDateAttendance(dateStr) {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');
    if (!_attendanceRoleCanMark_(user)) {
      return errorResponse('Only admins, leaders, sub-leaders, and karyakartas can mark attendance.');
    }
    var ds = String(dateStr || '').trim();
    if (!ds) return errorResponse('Date is required');

    var found = _findVisibleSabhaForDate_(ds, user);
    if (found && found.id) return successResponse(found);

    var mandalId = _defaultMandalForAttendance_(user);
    if (!mandalId) {
      return errorResponse('No Mandal is configured. Add Mandals in the sheet or set Mandal on your user profile.');
    }

    var sheet = getSheet(SHEET_SABHAS);
    var newId = generateId(sheet, 'SAB');
    var dateObj = new Date(ds + 'T12:00:00');
    var displayDate = Utilities.formatDate(dateObj, _appTimeZone(), 'dd MMM yyyy');
    var title = 'Weekly Sabha – ' + displayDate;
    var now = new Date();
    var row = [];
    var c;
    for (c = 0; c <= SABHAS_COL.CreatedOn; c++) row[c] = '';
    row[SABHAS_COL.SabhaID] = newId;
    row[SABHAS_COL.Title] = title;
    row[SABHAS_COL.Type] = 'weekly';
    row[SABHAS_COL.Date] = dateObj;
    row[SABHAS_COL.Time] = '';
    row[SABHAS_COL.Venue] = '';
    row[SABHAS_COL.MandalID] = mandalId;
    row[SABHAS_COL.RecurringType] = 'weekly';
    row[SABHAS_COL.Mode] = 'in-person';
    row[SABHAS_COL.MeetingLink] = '';
    row[SABHAS_COL.Speaker] = '';
    row[SABHAS_COL.KathaaTopic] = '';
    row[SABHAS_COL.Notes] = '';
    row[SABHAS_COL.MaxCapacity] = '';
    row[SABHAS_COL.GoogleCalendarEventID] = '';
    row[SABHAS_COL.CreatedBy] = user.userID || user.email || 'system';
    row[SABHAS_COL.CreatedOn] = now;

    sheet.appendRow(row);
    SpreadsheetApp.flush();

    return successResponse({
      rowIndex: sheet.getLastRow(),
      id: newId,
      date: ds,
      title: title,
      notes: '',
      mandalID: mandalId,
    });
  } catch (e) {
    logError('getOrCreateSabhaForDateAttendance', { dateStr: dateStr }, e);
    return errorResponse(e.message || String(e));
  }
}

/**
 * @param {string} sabhaId
 * @return {{ success: boolean, data?: string[], error?: string }}
 */
function getSabhaPresentDevoteeIds(sabhaId) {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');
    if (!_attendanceRoleCanMark_(user)) {
      return errorResponse('Only admins, leaders, sub-leaders, and karyakartas can mark attendance.');
    }
    var sheet = getSheet(SHEET_SABHAS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return successResponse([]);
    var sid = String(sabhaId || '').trim();
    var data = sheet.getRange(2, 1, lastRow - 1, SABHAS_COL.MandalID + 1).getValues();
    var i;
    var ok = false;
    for (i = 0; i < data.length; i++) {
      if (String(data[i][SABHAS_COL.SabhaID] || '').trim() === sid && _dashboardSabhaVisible(data[i], user)) {
        ok = true;
        break;
      }
    }
    if (!ok) return errorResponse('Sabha not found or access denied');

    var att = getSheet(SHEET_ATTENDANCE);
    var lr = att.getLastRow();
    if (lr < 2) return successResponse([]);
    var attData = att.getRange(2, 1, lr - 1, ATTENDANCE_COL.IsStreak + 1).getValues();
    var ids = [];
    for (i = 0; i < attData.length; i++) {
      var r = attData[i];
      if (String(r[ATTENDANCE_COL.SabhaID] || '').trim() !== sid) continue;
      if (_attendanceRowIsPresentOrLate_(r)) ids.push(String(r[ATTENDANCE_COL.DevoteeID] || '').trim());
    }
    return successResponse(ids);
  } catch (e) {
    logError('getSabhaPresentDevoteeIds', { sabhaId: sabhaId }, e);
    return errorResponse(e.message || String(e));
  }
}

/**
 * @param {string} sabhaId
 * @return {{ success: boolean, data?: Object, error?: string }}
 */
function getLiveSabhaAttendance(sabhaId) {
  try {
    var res = getSabhaPresentDevoteeIds(sabhaId);
    if (!res || !res.success) return res;
    var ids = res.data || [];
    return successResponse({
      presentIds: ids,
      count: ids.length,
      ts: new Date().getTime(),
    });
  } catch (e) {
    logError('getLiveSabhaAttendance', { sabhaId: sabhaId }, e);
    return errorResponse(e.message || String(e));
  }
}

/**
 * @param {string} sabhaId
 * @param {string} devoteeId
 * @return {{ success: boolean, data?: Object, error?: string }}
 */
function markSabhaPresent(sabhaId, devoteeId) {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');
    if (!_attendanceRoleCanMark_(user)) {
      return errorResponse('Only admins, leaders, sub-leaders, and karyakartas can mark attendance.');
    }
    if (!sabhaId || !devoteeId) return errorResponse('Sabha and devotee are required');

    var verify = getSabhaPresentDevoteeIds(sabhaId);
    if (!verify || !verify.success) return verify;

    var values = _getDevoteesRowsCached_();
    var idx = _findDevoteeIndexInRows_(values, devoteeId);
    if (idx < 0) return errorResponse('Devotee not found');
    if (!_devoteeRowVisibleForRole(values[idx], user)) return errorResponse('Access denied for this devotee');

    var lock = LockService.getScriptLock();
    var lockAcquired = false;
    try {
      lock.waitLock(15000);
      lockAcquired = true;

      var sheet = getSheet(SHEET_ATTENDANCE);
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        var data = sheet.getRange(2, 1, lastRow - 1, ATTENDANCE_COL.IsStreak + 1).getValues();
        var i;
        for (i = 0; i < data.length; i++) {
          var r = data[i];
          if (
            String(r[ATTENDANCE_COL.SabhaID] || '').trim() === String(sabhaId).trim() &&
            String(r[ATTENDANCE_COL.DevoteeID] || '').trim() === String(devoteeId).trim() &&
            _attendanceRowIsPresentOrLate_(r)
          ) {
            return successResponse({ status: 'already_marked' });
          }
        }
      }

      var newId = generateId(sheet, 'ATT');
      var row = [];
      for (var c = 0; c <= ATTENDANCE_COL.IsStreak; c++) row[c] = '';
      row[ATTENDANCE_COL.AttendanceID] = newId;
      row[ATTENDANCE_COL.SabhaID] = String(sabhaId).trim();
      row[ATTENDANCE_COL.DevoteeID] = String(devoteeId).trim();
      row[ATTENDANCE_COL.Status] = ATTENDANCE_STATUS.PRESENT;
      row[ATTENDANCE_COL.MarkedAt] = new Date();
      row[ATTENDANCE_COL.MarkedBy] = user.userID || user.email || '';
      row[ATTENDANCE_COL.CheckInMethod] = 'app';
      row[ATTENDANCE_COL.IsStreak] = false;

      sheet.appendRow(row);
      SpreadsheetApp.flush();
      return successResponse({ status: 'success', attendanceId: newId });
    } catch (e) {
      logError('markSabhaPresent', { sabhaId: sabhaId, devoteeId: devoteeId }, e);
      return errorResponse(e.message || String(e));
    } finally {
      if (lockAcquired) {
        try {
          lock.releaseLock();
        } catch (le) {
          Logger.log('releaseLock: ' + le.message);
        }
      }
    }
  } catch (e2) {
    logError('markSabhaPresent.outer', { sabhaId: sabhaId, devoteeId: devoteeId }, e2);
    return errorResponse(e2.message || String(e2));
  }
}

/**
 * @param {string} sabhaId
 * @param {string} devoteeId
 * @return {{ success: boolean, data?: Object, error?: string }}
 */
function unmarkSabhaPresent(sabhaId, devoteeId) {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');
    if (!_attendanceRoleCanMark_(user)) {
      return errorResponse('Only admins, leaders, sub-leaders, and karyakartas can mark attendance.');
    }
    if (!sabhaId || !devoteeId) return errorResponse('Sabha and devotee are required');

    var verify = getSabhaPresentDevoteeIds(sabhaId);
    if (!verify || !verify.success) return verify;

    var values = _getDevoteesRowsCached_();
    var idx = _findDevoteeIndexInRows_(values, devoteeId);
    if (idx < 0) return errorResponse('Devotee not found');
    if (!_devoteeRowVisibleForRole(values[idx], user)) return errorResponse('Access denied for this devotee');

    var lock = LockService.getScriptLock();
    var lockAcquired = false;
    try {
      lock.waitLock(15000);
      lockAcquired = true;

      var sheet = getSheet(SHEET_ATTENDANCE);
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return successResponse({ status: 'not_found' });

      var data = sheet.getRange(2, 1, lastRow - 1, ATTENDANCE_COL.IsStreak + 1).getValues();
      var i;
      for (i = data.length - 1; i >= 0; i--) {
        var r = data[i];
        if (
          String(r[ATTENDANCE_COL.SabhaID] || '').trim() === String(sabhaId).trim() &&
          String(r[ATTENDANCE_COL.DevoteeID] || '').trim() === String(devoteeId).trim() &&
          _attendanceRowIsPresentOrLate_(r)
        ) {
          sheet.deleteRow(i + 2);
          SpreadsheetApp.flush();
          return successResponse({ status: 'success' });
        }
      }
      return successResponse({ status: 'not_found' });
    } catch (e) {
      logError('unmarkSabhaPresent', { sabhaId: sabhaId, devoteeId: devoteeId }, e);
      return errorResponse(e.message || String(e));
    } finally {
      if (lockAcquired) {
        try {
          lock.releaseLock();
        } catch (le2) {}
      }
    }
  } catch (e2) {
    logError('unmarkSabhaPresent.outer', { sabhaId: sabhaId, devoteeId: devoteeId }, e2);
    return errorResponse(e2.message || String(e2));
  }
}

/**
 * @return {{ success: boolean, data?: Array<Object>, error?: string }}
 */
function getAllSabhasForAttendance() {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');
    if (!_attendanceRoleCanMark_(user)) {
      return errorResponse('Only admins, leaders, sub-leaders, and karyakartas can view attendance reports.');
    }
    var sheet = getSheet(SHEET_SABHAS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return successResponse([]);
    var data = sheet.getRange(2, 1, lastRow - 1, SABHAS_COL.CreatedOn + 1).getValues();
    var out = [];
    var i;
    for (i = 0; i < data.length; i++) {
      var r = data[i];
      if (!_dashboardSabhaVisible(r, user)) continue;
      var cell = r[SABHAS_COL.Date];
      var ds = cell instanceof Date && !isNaN(cell.getTime()) ? formatDate(cell) : '';
      out.push({
        id: String(r[SABHAS_COL.SabhaID] || '').trim(),
        date: ds,
        title: String(r[SABHAS_COL.Title] || '').trim(),
        notes: String(r[SABHAS_COL.Notes] || '').trim(),
      });
    }
    out.sort(function (a, b) {
      if (a.date === b.date) return 0;
      return a.date < b.date ? 1 : -1;
    });
    return successResponse(out);
  } catch (e) {
    logError('getAllSabhasForAttendance', {}, e);
    return errorResponse(e.message || String(e));
  }
}

/**
 * @param {string} sabhaId
 * @param {string} notes
 * @return {{ success: boolean, data?: Object, error?: string }}
 */
function updateSabhaNotesAttendance(sabhaId, notes) {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');
    if (!_attendanceRoleCanMark_(user)) return errorResponse('Access denied');

    var sheet = getSheet(SHEET_SABHAS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return successResponse({ status: 'not_found' });
    var data = sheet.getRange(2, 1, lastRow - 1, SABHAS_COL.Notes + 1).getValues();
    var i;
    for (i = 0; i < data.length; i++) {
      if (!_dashboardSabhaVisible(data[i], user)) continue;
      if (String(data[i][SABHAS_COL.SabhaID] || '').trim() === String(sabhaId || '').trim()) {
        sheet.getRange(i + 2, SABHAS_COL.Notes + 1).setValue(notes == null ? '' : String(notes));
        SpreadsheetApp.flush();
        return successResponse({ status: 'success' });
      }
    }
    return successResponse({ status: 'not_found' });
  } catch (e) {
    logError('updateSabhaNotesAttendance', { sabhaId: sabhaId }, e);
    return errorResponse(e.message || String(e));
  }
}

/**
 * @param {string} sabhaId
 * @return {{ success: boolean, data?: Object, error?: string }}
 */
function getSabhaAttendanceReport(sabhaId) {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');
    if (!_attendanceRoleCanMark_(user)) return errorResponse('Access denied');

    var devotees = _getAttendanceDevoteeSummaries_(user);
    var sheet = getSheet(SHEET_SABHAS);
    var lastRow = sheet.getLastRow();
    var session = null;
    if (lastRow >= 2) {
      var sesData = sheet.getRange(2, 1, lastRow - 1, SABHAS_COL.Notes + 1).getValues();
      var j;
      for (j = 0; j < sesData.length; j++) {
        var sr = sesData[j];
        if (!_dashboardSabhaVisible(sr, user)) continue;
        if (String(sr[SABHAS_COL.SabhaID] || '').trim() === String(sabhaId || '').trim()) {
          var dc = sr[SABHAS_COL.Date];
          session = {
            id: String(sr[SABHAS_COL.SabhaID] || '').trim(),
            date: dc instanceof Date && !isNaN(dc.getTime()) ? formatDate(dc) : '',
            title: String(sr[SABHAS_COL.Title] || '').trim(),
            notes: String(sr[SABHAS_COL.Notes] || '').trim(),
          };
          break;
        }
      }
    }

    var attSheet = getSheet(SHEET_ATTENDANCE);
    var lr = attSheet.getLastRow();
    var presentIds = {};
    if (lr >= 2) {
      var attData = attSheet.getRange(2, 1, lr - 1, ATTENDANCE_COL.IsStreak + 1).getValues();
      var k;
      for (k = 0; k < attData.length; k++) {
        var ar = attData[k];
        if (String(ar[ATTENDANCE_COL.SabhaID] || '').trim() !== String(sabhaId || '').trim()) continue;
        if (_attendanceRowIsPresentOrLate_(ar)) {
          presentIds[String(ar[ATTENDANCE_COL.DevoteeID] || '').trim()] = true;
        }
      }
    }

    var present = [];
    var absent = [];
    var d;
    for (d = 0; d < devotees.length; d++) {
      var dev = devotees[d];
      if (presentIds[dev.id]) present.push(dev);
      else absent.push(dev);
    }
    var total = devotees.length;
    var pct = total > 0 ? Math.round((present.length / total) * 100) : 0;

    return successResponse({
      session: session,
      total: total,
      presentCount: present.length,
      absentCount: absent.length,
      percentage: pct,
      present: present,
      absent: absent,
    });
  } catch (e) {
    logError('getSabhaAttendanceReport', { sabhaId: sabhaId }, e);
    return errorResponse(e.message || String(e));
  }
}

/**
 * @param {string} devoteeId
 * @return {{ success: boolean, data?: Object, error?: string }}
 */
function getDevoteeAttendanceHistory(devoteeId) {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');
    if (!_attendanceRoleCanMark_(user)) return errorResponse('Access denied');

    var values = _getDevoteesRowsCached_();
    var idx = _findDevoteeIndexInRows_(values, devoteeId);
    if (idx < 0) return errorResponse('Devotee not found');
    if (!_devoteeRowVisibleForRole(values[idx], user)) return errorResponse('Access denied');

    var sessionsRes = getAllSabhasForAttendance();
    if (!sessionsRes || !sessionsRes.success) return sessionsRes;
    var sessions = sessionsRes.data || [];

    var attSheet = getSheet(SHEET_ATTENDANCE);
    var lastRow = attSheet.getLastRow();
    var attendedSet = {};
    if (lastRow >= 2) {
      var data = attSheet.getRange(2, 1, lastRow - 1, ATTENDANCE_COL.IsStreak + 1).getValues();
      var i;
      for (i = 0; i < data.length; i++) {
        var r = data[i];
        if (String(r[ATTENDANCE_COL.DevoteeID] || '').trim() !== String(devoteeId || '').trim()) continue;
        if (_attendanceRowIsPresentOrLate_(r)) {
          attendedSet[String(r[ATTENDANCE_COL.SabhaID] || '').trim()] = true;
        }
      }
    }

    var history = [];
    for (var s = 0; s < sessions.length; s++) {
      var sess = sessions[s];
      history.push({
        sessionId: sess.id,
        date: sess.date,
        title: sess.title,
        attended: !!attendedSet[sess.id],
      });
    }
    var attCount = Object.keys(attendedSet).length;
    var pct = sessions.length > 0 ? Math.round((attCount / sessions.length) * 100) : 0;
    return successResponse({
      history: history,
      totalSessions: sessions.length,
      attended: attCount,
      percentage: pct,
    });
  } catch (e) {
    logError('getDevoteeAttendanceHistory', { devoteeId: devoteeId }, e);
    return errorResponse(e.message || String(e));
  }
}

/**
 * @return {{ success: boolean, data?: Object, error?: string }}
 */
function getAttendanceAnalytics() {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');
    if (!_attendanceRoleCanMark_(user)) return errorResponse('Access denied');

    var devotees = _getAttendanceDevoteeSummaries_(user);
    var sessionsRes = getAllSabhasForAttendance();
    if (!sessionsRes || !sessionsRes.success) return sessionsRes;
    var sessions = sessionsRes.data || [];

    var attSheet = getSheet(SHEET_ATTENDANCE);
    var lastRow = attSheet.getLastRow();
    var countMap = {};
    var di;
    for (di = 0; di < devotees.length; di++) countMap[devotees[di].id] = 0;

    var totalMarks = 0;
    var attMatrix = {};
    if (lastRow >= 2) {
      var data = attSheet.getRange(2, 1, lastRow - 1, ATTENDANCE_COL.IsStreak + 1).getValues();
      var i;
      for (i = 0; i < data.length; i++) {
        var r = data[i];
        if (!_attendanceRowIsPresentOrLate_(r)) continue;
        var did = String(r[ATTENDANCE_COL.DevoteeID] || '').trim();
        var sid = String(r[ATTENDANCE_COL.SabhaID] || '').trim();
        if (countMap[did] !== undefined) countMap[did]++;
        totalMarks++;
        if (!attMatrix[sid]) attMatrix[sid] = 0;
        attMatrix[sid]++;
      }
    }

    var avgPct =
      sessions.length > 0 && devotees.length > 0
        ? Math.round((totalMarks / (sessions.length * devotees.length)) * 100)
        : 0;

    var top5 = devotees
      .map(function (d) {
        var c = countMap[d.id] || 0;
        return {
          id: d.id,
          firstName: d.firstName,
          middleName: d.middleName,
          lastName: d.lastName,
          gender: d.gender,
          mobile: d.mobile,
          birthday: d.birthday,
          karyakarta: d.karyakarta,
          count: c,
          pct: sessions.length > 0 ? Math.round((c / sessions.length) * 100) : 0,
        };
      })
      .sort(function (a, b) {
        return b.count - a.count;
      })
      .slice(0, 5);

    var todayMD = Utilities.formatDate(new Date(), _appTimeZone(), 'MM-dd');
    var birthdays = devotees.filter(function (d) {
      if (!d.birthday || d.birthday.length < 5) return false;
      return d.birthday.substring(5) === todayMD;
    });

    var trend = [];
    var si;
    for (si = 0; si < Math.min(8, sessions.length); si++) {
      var s = sessions[si];
      var cnt = attMatrix[s.id] || 0;
      trend.push({
        sessionId: s.id,
        date: s.date,
        title: s.title,
        count: cnt,
        pct: devotees.length > 0 ? Math.round((cnt / devotees.length) * 100) : 0,
      });
    }

    var recentSessionIds = [];
    var ri;
    for (ri = 0; ri < Math.min(10, sessions.length); ri++) recentSessionIds.push(sessions[ri].id);

    var streakDevoteeIds = [];
    if (recentSessionIds.length >= 3 && lastRow >= 2) {
      var slice3 = recentSessionIds.slice(0, 3);
      var dvAll = attSheet.getRange(2, 1, lastRow - 1, ATTENDANCE_COL.IsStreak + 1).getValues();
      for (var dj = 0; dj < devotees.length; dj++) {
        var dev = devotees[dj];
        var allThree = true;
        var t;
        for (t = 0; t < slice3.length; t++) {
          var found = false;
          var u;
          for (u = 0; u < dvAll.length; u++) {
            var rr = dvAll[u];
            if (
              String(rr[ATTENDANCE_COL.SabhaID] || '').trim() === slice3[t] &&
              String(rr[ATTENDANCE_COL.DevoteeID] || '').trim() === dev.id &&
              _attendanceRowIsPresentOrLate_(rr)
            ) {
              found = true;
              break;
            }
          }
          if (!found) {
            allThree = false;
            break;
          }
        }
        if (allThree) streakDevoteeIds.push(dev.id);
      }
    }

    return successResponse({
      totalDevotees: devotees.length,
      totalSessions: sessions.length,
      averagePct: avgPct,
      top5: top5,
      birthdays: birthdays,
      trend: trend,
      streakDevoteeIds: streakDevoteeIds,
    });
  } catch (e) {
    logError('getAttendanceAnalytics', {}, e);
    return errorResponse(e.message || String(e));
  }
}

/**
 * @param {string} sabhaId
 * @return {{ success: boolean, data?: Object, error?: string }}
 */
function exportSabhaAttendanceReport(sabhaId) {
  try {
    var rep = getSabhaAttendanceReport(sabhaId);
    if (!rep || !rep.success || !rep.data) return rep;
    var report = rep.data;
    var s = report.session;
    var lines = [
      'HPP Akshar Setu — Attendance Report',
      (s ? s.title : sabhaId) + ' | Date: ' + (s ? s.date : ''),
      'Present: ' + report.presentCount + ' | Absent: ' + report.absentCount + ' | ' + report.percentage + '%',
      '',
      '--- PRESENT ---',
    ];
    var i;
    for (i = 0; i < report.present.length; i++) {
      var p = report.present[i];
      lines.push(String(i + 1) + '. ' + [p.firstName, p.middleName, p.lastName].filter(Boolean).join(' '));
    }
    lines.push('', '--- ABSENT ---');
    for (i = 0; i < report.absent.length; i++) {
      var a = report.absent[i];
      lines.push(String(i + 1) + '. ' + [a.firstName, a.middleName, a.lastName].filter(Boolean).join(' '));
    }
    lines.push('', 'Jai Swaminarayan');

    var csvRows = [['Name', 'Gender', 'Mobile', 'Karyakarta', 'Status']];
    for (i = 0; i < report.present.length; i++) {
      var pr = report.present[i];
      csvRows.push([
        [pr.firstName, pr.middleName, pr.lastName].filter(Boolean).join(' '),
        pr.gender,
        pr.mobile,
        pr.karyakarta,
        'Present',
      ]);
    }
    for (i = 0; i < report.absent.length; i++) {
      var ab = report.absent[i];
      csvRows.push([
        [ab.firstName, ab.middleName, ab.lastName].filter(Boolean).join(' '),
        ab.gender,
        ab.mobile,
        ab.karyakarta,
        'Absent',
      ]);
    }
    var csv = csvRows
      .map(function (r) {
        return r
          .map(function (c) {
            return '"' + String(c).replace(/"/g, '""') + '"';
          })
          .join(',');
      })
      .join('\n');

    return successResponse({ status: 'success', text: lines.join('\n'), csv: csv });
  } catch (e) {
    logError('exportSabhaAttendanceReport', { sabhaId: sabhaId }, e);
    return errorResponse(e.message || String(e));
  }
}

/**
 * @param {string} sabhaId
 * @return {{ success: boolean, data?: Array<Object>, error?: string }}
 */
function getKaryakartaAttendanceReport(sabhaId) {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');
    if (!_attendanceRoleCanMark_(user)) return errorResponse('Access denied');

    var rep = getSabhaAttendanceReport(sabhaId);
    if (!rep || !rep.success || !rep.data) return rep;
    var report = rep.data;

    var presentSet = {};
    for (var i = 0; i < report.present.length; i++) presentSet[report.present[i].id] = true;

    var devotees = _getAttendanceDevoteeSummaries_(user);
    function mobileForKaryakartaLabel(label) {
      var L = String(label || '').trim();
      if (!L || L === '(Unassigned)') return '';
      var idx;
      for (idx = 0; idx < devotees.length; idx++) {
        var cand = devotees[idx];
        var fn = [cand.firstName, cand.middleName, cand.lastName].filter(Boolean).join(' ');
        if (fn === L || String(cand.id) === L) return String(cand.mobile || '').trim();
      }
      return '';
    }

    var karyaMap = {};
    for (var d = 0; d < devotees.length; d++) {
      var dev = devotees[d];
      var k = dev.karyakarta || '(Unassigned)';
      if (!karyaMap[k]) {
        karyaMap[k] = { name: k, mobile: mobileForKaryakartaLabel(k), total: 0, present: [], absent: [] };
      }
      karyaMap[k].total++;
      if (presentSet[dev.id]) karyaMap[k].present.push(dev);
      else karyaMap[k].absent.push(dev);
    }

    var list = [];
    for (var key in karyaMap) {
      if (Object.prototype.hasOwnProperty.call(karyaMap, key)) list.push(karyaMap[key]);
    }
    list.sort(function (a, b) {
      var ap = a.present.length / (a.total || 1);
      var bp = b.present.length / (b.total || 1);
      return bp - ap;
    });
    return successResponse(list);
  } catch (e) {
    logError('getKaryakartaAttendanceReport', { sabhaId: sabhaId }, e);
    return errorResponse(e.message || String(e));
  }
}
