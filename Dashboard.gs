/**
 * HPP Seva Connect — dashboard aggregates (KPIs, charts, lists).
 */

/**
 * @return {{ success: boolean, data?: Object, error?: string }}
 */
 
function getBootstrapDashboardData() {
  try {
    var startedAt = Date.now();
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');
    var snap = _getDashboardSnapshotFastForUser_(user);
    if (!snap || !snap.success || !snap.data) {
      return errorResponse((snap && snap.error) || 'Could not load dashboard snapshot');
    }
    return successResponse({
      user: user,
      dashboard: snap.data,
      perf: {
        totalMs: Date.now() - startedAt,
      },
    });
  } catch (e) {
    var errorId = logError('getBootstrapDashboardData', {}, e);
    return errorResponse((e.message || String(e)) + ' (Ref: ' + errorId + ')', errorId);
  }
}

/**
 * @return {{ success: boolean, data?: Object, error?: string }}
 */
function getDashboardSnapshotFast() {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');
    return _getDashboardSnapshotFastForUser_(user);
  } catch (e) {
    var errorId = logError('getDashboardSnapshotFast', {}, e);
    return errorResponse((e.message || String(e)) + ' (Ref: ' + errorId + ')', errorId);
  }
}

function _getDashboardSnapshotFastForUser_(user) {
  var role = String((user && user.role) || '').toLowerCase();
  var uid = String((user && (user.userID || user.email)) || '').trim();
  var mandal = String((user && user.mandalID) || '').trim();
  var ver = _devoteesCacheVersion_();
  var snapshotKey = 'dash_fast|' + ver + '|' + role + '|' + uid + '|' + mandal;
  try {
    var c = CacheService.getScriptCache();
    var rawCached = c.get(snapshotKey);
    if (rawCached) {
      var parsedCached = JSON.parse(rawCached);
      if (parsedCached) return successResponse(parsedCached);
    }
  } catch (ignoreFastCacheRead) {}

  if (role === ROLES.ADMIN) {
    var adminSnap = getDashboardSnapshotCached();
    if (adminSnap) {
      adminSnap.cacheStamp = _devoteesCacheVersion_();
      try {
        CacheService.getScriptCache().put(snapshotKey, JSON.stringify(adminSnap), 1800);
      } catch (ignoreAdminFastCacheWrite) {}
      return successResponse(adminSnap);
    }
  }
  var perfStart = Date.now();
  var perf = {};
  var t0;

  t0 = Date.now();
  var devValues = _getDevoteesRowsCached_();
  perf.readDevoteesMs = Date.now() - t0;

  var totalVisible = 0;
  var activeDevotees = 0;
  var familiesSet = {};
  var todayBirthdays = [];
  var upcomingBirthdays = [];
  var karyakartaYuvakMap = {};
  var now = new Date();
  var isAdmin = String((user && user.role) || '').toLowerCase() === ROLES.ADMIN;

  t0 = Date.now();
  for (var i = 0; i < devValues.length; i++) {
    var row = devValues[i];
    if (!isAdmin && !_devoteeRowVisibleForRole(row, user)) continue;
    totalVisible++;
    var st = String(row[DEVOTEES_COL.Status] || '').toLowerCase();
    if (st === DEVOTEE_STATUS.ACTIVE) activeDevotees++;
    var fid = String(row[DEVOTEES_COL.FamilyID] || '').trim();
    if (fid) familiesSet[fid] = true;
    var kName = String(row[DEVOTEES_COL.KaryakartaName] || '').trim() || '—';
    if (!karyakartaYuvakMap[kName]) karyakartaYuvakMap[kName] = { name: kName, yuvakCount: 0, totalCount: 0 };
    karyakartaYuvakMap[kName].totalCount++;
    var dtype = String(row[DEVOTEES_COL.DevoteeType] || '').toLowerCase();
    if (dtype.indexOf('yuvak') !== -1 || dtype.indexOf('youth') !== -1) {
      karyakartaYuvakMap[kName].yuvakCount++;
    }
    if (st === DEVOTEE_STATUS.ACTIVE) {
      var b = _dashboardBirthdayInfo(row, now, _appTimeZone(), {}, {});
      if (b) {
        if (b.daysAway === 0) todayBirthdays.push(b);
        else if (b.daysAway <= 30) upcomingBirthdays.push(b);
      }
    }
  }
  upcomingBirthdays.sort(function (a, b) { return a.daysAway - b.daysAway; });
  var kRows = Object.keys(karyakartaYuvakMap).map(function (k) { return karyakartaYuvakMap[k]; })
    .sort(function (a, b) {
      if (b.yuvakCount !== a.yuvakCount) return b.yuvakCount - a.yuvakCount;
      if (b.totalCount !== a.totalCount) return b.totalCount - a.totalCount;
      return a.name.localeCompare(b.name);
    });
  perf.processDevoteesMs = Date.now() - t0;
  perf.totalMs = Date.now() - perfStart;

  var computed = {
    totalVisible: totalVisible,
    activeDevotees: activeDevotees,
    inactiveDevotees: 0,
    relocatedDevotees: 0,
    deceasedDevotees: 0,
    familiesCount: Object.keys(familiesSet).length,
    maleCount: 0,
    femaleCount: 0,
    otherCount: 0,
    donationEnabledCount: 0,
    whatsAppOptedInCount: 0,
    karyakartaCount: kRows.length,
    pendingFollowups: 0,
    weekSabhaPercent: null,
    todayBirthdays: todayBirthdays.slice(0, 60),
    upcomingBirthdays30: upcomingBirthdays.slice(0, 120),
    karyakartaYuvakCounts: kRows.slice(0, 200),
    attendanceLast4Weeks: [0, 0, 0, 0],
    devoteeGrowthMonths: ['M1', 'M2', 'M3', 'M4', 'M5', 'M6'],
    devoteeGrowthCounts: [0, 0, 0, 0, 0, 0],
    statusLabels: ['Active', 'Inactive', 'Relocated', 'Deceased'],
    statusCounts: [activeDevotees, 0, 0, 0],
    genderLabels: ['Male', 'Female', 'Other'],
    genderCounts: [0, 0, 0],
    ageBandLabels: ['Kids', 'Youth', 'Adult', 'Senior', 'Unknown'],
    ageBandCounts: [0, 0, 0, 0, 0],
    birthdayTrendLabels: ['W1', 'W2', 'W3', 'W4', 'W5'],
    birthdayTrendCounts: [0, 0, 0, 0, 0],
    motivationQuote: { text: 'હે હરિ ! બસ એક, તું રાજી થા.', author: 'Satsang' },
    cacheStamp: _devoteesCacheVersion_(),
    perf: perf,
  };
  try {
    CacheService.getScriptCache().put(snapshotKey, JSON.stringify(computed), 1800);
  } catch (ignoreFastCacheWrite) {}
  return successResponse(computed);
}

/**
 * @return {{ success: boolean, data?: Object, error?: string }}
 */
function getDashboardSummary() {
  try {
    var perf = {};
    var perfStart = Date.now();
    var t0 = perfStart;
    var user = getCurrentUser();
    perf.getCurrentUserMs = Date.now() - t0;
    if (!user) return errorResponse('Not signed in');

    var tz = _appTimeZone();
    var now = new Date();

    t0 = Date.now();
    var devValues = _getDevoteesRowsCached_();
    perf.readDevoteesMs = Date.now() - t0;
    var userNameById = {};
    try {
      var userSheet = getSheet(SHEET_USERS);
      var uLast = userSheet.getLastRow();
      if (uLast >= 2) {
        t0 = Date.now();
        var uRows = userSheet.getRange(2, 1, uLast - 1, 11).getValues();
        perf.readUsersMs = Date.now() - t0;
        t0 = Date.now();
        for (var ui = 0; ui < uRows.length; ui++) {
          var uid = String(uRows[ui][USERS_COL.UserID] || '').trim();
          if (!uid) continue;
          userNameById[uid] = String(uRows[ui][USERS_COL.Name] || uid).trim();
        }
        perf.buildUserMapMs = Date.now() - t0;
      }
    } catch (ignoreUsers) {}
    var totalVisible = 0;
    var activeDevotees = 0;
    var inactiveDevotees = 0;
    var relocatedDevotees = 0;
    var deceasedDevotees = 0;
    var maleCount = 0;
    var femaleCount = 0;
    var otherCount = 0;
    var elderlyCount = 0;
    var specialNeedsCount = 0;
    var donationEnabledCount = 0;
    var whatsAppOptedInCount = 0;
    var karyakartaCount = 0;
    var karyakartaYuvakMap = {};
    var familiesSet = {};
    var todayBirthdays = [];
    var upcomingBirthdays = [];
    var devById = {};
    var i, row;

    t0 = Date.now();
    for (i = 0; i < devValues.length; i++) {
      row = devValues[i];
      if (!_devoteeRowVisibleForRole(row, user)) continue;
      totalVisible++;
      var did = String(row[DEVOTEES_COL.DevoteeID] || '').trim();
      if (did) devById[did] = row;
      var st = String(row[DEVOTEES_COL.Status] || '').toLowerCase();
      if (st === DEVOTEE_STATUS.ACTIVE) activeDevotees++;
      else if (st === DEVOTEE_STATUS.INACTIVE) inactiveDevotees++;
      else if (st === DEVOTEE_STATUS.RELOCATED) relocatedDevotees++;
      else if (st === DEVOTEE_STATUS.DECEASED) deceasedDevotees++;
      var fid = String(row[DEVOTEES_COL.FamilyID] || '').trim();
      if (fid) familiesSet[fid] = true;
      var g = String(row[DEVOTEES_COL.Gender] || '').toLowerCase();
      if (g === 'male' || g === 'm') maleCount++;
      else if (g === 'female' || g === 'f') femaleCount++;
      else if (g) otherCount++;
      if (String(row[DEVOTEES_COL.IsElderly] || '').toLowerCase() === 'true') elderlyCount++;
      if (String(row[DEVOTEES_COL.SpecialNeeds] || '').trim()) specialNeedsCount++;
      if (String(row[DEVOTEES_COL.DonationEnabled] || '').toLowerCase() === 'true') donationEnabledCount++;
      if (String(row[DEVOTEES_COL.WhatsAppOptedIn] || '').toLowerCase() === 'true') whatsAppOptedInCount++;
      var dType = String(row[DEVOTEES_COL.DevoteeType] || '').toLowerCase();
      var kNameRaw = String(row[DEVOTEES_COL.KaryakartaName] || row[DEVOTEES_COL.KaryakartaID] || '').trim();
      var kName = _dashboardResolveKaryakartaName_(kNameRaw, userNameById, devById);
      if (kName && kName !== '—') {
        if (!karyakartaYuvakMap[kName]) {
          karyakartaYuvakMap[kName] = { name: kName, yuvakCount: 0, totalCount: 0 };
        }
        karyakartaYuvakMap[kName].totalCount++;
        if (_dashboardIsYuvakType_(dType)) karyakartaYuvakMap[kName].yuvakCount++;
      }

      var b = _dashboardBirthdayInfo(row, now, tz, userNameById, devById);
      if (b && st === DEVOTEE_STATUS.ACTIVE) {
        if (b.daysAway === 0) todayBirthdays.push(b);
        else if (b.daysAway <= 30) upcomingBirthdays.push(b);
      }
    }
    perf.processDevoteesMs = Date.now() - t0;
    upcomingBirthdays.sort(function (a, b) { return a.daysAway - b.daysAway; });

    var fuSheet = getSheet(SHEET_FOLLOWUPS);
    var fuLast = fuSheet.getLastRow();
    t0 = Date.now();
    var fuValues = fuLast >= 2 ? fuSheet.getRange(2, 1, fuLast - 1, 13).getValues() : [];
    perf.readFollowupsMs = Date.now() - t0;
    var pendingFollowups = 0;
    t0 = Date.now();
    for (i = 0; i < fuValues.length; i++) {
      row = fuValues[i];
      if (String(row[FOLLOWUPS_COL.Status] || '').toLowerCase() !== FOLLOWUP_STATUS.PENDING) continue;
      var devId = String(row[FOLLOWUPS_COL.DevoteeID] || '').trim();
      var drow = devById[devId];
      if (!drow || !_devoteeRowVisibleForRole(drow, user)) continue;
      pendingFollowups++;
    }
    perf.processFollowupsMs = Date.now() - t0;

    var sabSheet = getSheet(SHEET_SABHAS);
    var sabLast = sabSheet.getLastRow();
    t0 = Date.now();
    var sabValues = sabLast >= 2 ? sabSheet.getRange(2, 1, sabLast - 1, 17).getValues() : [];
    perf.readSabhasMs = Date.now() - t0;

    var attSheet = getSheet(SHEET_ATTENDANCE);
    var attLast = attSheet.getLastRow();
    t0 = Date.now();
    var attValues = attLast >= 2 ? attSheet.getRange(2, 1, attLast - 1, 8).getValues() : [];
    perf.readAttendanceMs = Date.now() - t0;

    t0 = Date.now();
    var weekSabhaPercent = _dashboardWeekSabhaPercent(sabValues, attValues, user, now, tz);
    var attendanceLast4Weeks = _dashboardAttendanceByWeek(sabValues, attValues, user, now, tz);
    var growth = _dashboardDevoteeGrowth(devValues, user, now, tz);
    var ageBand = _dashboardAgeBandCounts(devValues, user);
    var birthdayTrend = _dashboardUpcomingBirthdayWeeklyCounts(upcomingBirthdays);
    perf.computeChartsMs = Date.now() - t0;
    var karyakartaYuvakCounts = Object.keys(karyakartaYuvakMap)
      .map(function (k) { return karyakartaYuvakMap[k]; })
      .sort(function (a, b) {
        if (b.yuvakCount !== a.yuvakCount) return b.yuvakCount - a.yuvakCount;
        if (b.totalCount !== a.totalCount) return b.totalCount - a.totalCount;
        return a.name.localeCompare(b.name);
      });
    karyakartaCount = karyakartaYuvakCounts.length;
    var familiesCount = Object.keys(familiesSet).length;
    t0 = Date.now();
    var quote = _dashboardMotivationQuote(now, tz);
    perf.quoteMs = Date.now() - t0;
    perf.totalMs = Date.now() - perfStart;

    return successResponse({
      totalVisible: totalVisible,
      activeDevotees: activeDevotees,
      inactiveDevotees: inactiveDevotees,
      relocatedDevotees: relocatedDevotees,
      deceasedDevotees: deceasedDevotees,
      familiesCount: familiesCount,
      maleCount: maleCount,
      femaleCount: femaleCount,
      otherCount: otherCount,
      elderlyCount: elderlyCount,
      specialNeedsCount: specialNeedsCount,
      donationEnabledCount: donationEnabledCount,
      whatsAppOptedInCount: whatsAppOptedInCount,
      karyakartaCount: karyakartaCount,
      pendingFollowups: pendingFollowups,
      weekSabhaPercent: weekSabhaPercent,
      todayBirthdays: todayBirthdays,
      upcomingBirthdays30: upcomingBirthdays.slice(0, 80),
      karyakartaYuvakCounts: karyakartaYuvakCounts.slice(0, 120),
      attendanceLast4Weeks: attendanceLast4Weeks,
      devoteeGrowthMonths: growth.labels,
      devoteeGrowthCounts: growth.counts,
      statusLabels: ['Active', 'Inactive', 'Relocated', 'Deceased'],
      statusCounts: [activeDevotees, inactiveDevotees, relocatedDevotees, deceasedDevotees],
      genderLabels: ['Male', 'Female', 'Other'],
      genderCounts: [maleCount, femaleCount, otherCount],
      ageBandLabels: ['Kids', 'Youth', 'Adult', 'Senior', 'Unknown'],
      ageBandCounts: [ageBand.kids, ageBand.youth, ageBand.adult, ageBand.senior, ageBand.unknown],
      birthdayTrendLabels: birthdayTrend.labels,
      birthdayTrendCounts: birthdayTrend.counts,
      motivationQuote: quote,
      perf: perf,
    });
  } catch (e) {
    var errorId = logError('getDashboardSummary', {}, e);
    return errorResponse((e.message || String(e)) + ' (Ref: ' + errorId + ')', errorId);
  }
}

/**
 * @param {Array} sabValues
 * @param {Array} attValues
 * @param {Object} user
 * @param {Date} now
 * @param {string} tz
 * @return {number|null}
 */
function _dashboardWeekSabhaPercent(sabValues, attValues, user, now, tz) {
  var start = _dashboardStartOfWeek(now, tz);
  var end = new Date(start.getTime());
  end.setDate(end.getDate() + 7);

  var pcts = [];
  var i;
  for (i = 0; i < sabValues.length; i++) {
    var srow = sabValues[i];
    if (!_dashboardSabhaVisible(srow, user)) continue;
    var d = srow[SABHAS_COL.Date];
    if (!(d instanceof Date) || isNaN(d.getTime())) continue;
    if (d < start || d >= end) continue;
    var sid = String(srow[SABHAS_COL.SabhaID] || '').trim();
    if (!sid) continue;
    var pct = _dashboardAttendancePctForSabha(sid, attValues);
    if (pct !== null) pcts.push(pct);
  }
  if (!pcts.length) return null;
  var sum = 0;
  for (i = 0; i < pcts.length; i++) sum += pcts[i];
  return Math.round(sum / pcts.length);
}

/**
 * @param {Array} sabValues
 * @param {Array} attValues
 * @param {Object} user
 * @param {Date} now
 * @param {string} tz
 * @return {number[]}
 */
function _dashboardAttendanceByWeek(sabValues, attValues, user, now, tz) {
  var out = [0, 0, 0, 0];
  var weekStarts = [];
  var w;
  var anchor = _dashboardStartOfWeek(now, tz);
  for (w = 3; w >= 0; w--) {
    var s = new Date(anchor.getTime());
    s.setDate(s.getDate() - w * 7);
    weekStarts.push(s);
  }

  for (w = 0; w < 4; w++) {
    var ws = weekStarts[w];
    var we = new Date(ws.getTime());
    we.setDate(we.getDate() + 7);
    var bucket = [];
    var i;
    for (i = 0; i < sabValues.length; i++) {
      var srow = sabValues[i];
      if (!_dashboardSabhaVisible(srow, user)) continue;
      var d = srow[SABHAS_COL.Date];
      if (!(d instanceof Date) || isNaN(d.getTime())) continue;
      if (d < ws || d >= we) continue;
      var sid = String(srow[SABHAS_COL.SabhaID] || '').trim();
      if (!sid) continue;
      var pct = _dashboardAttendancePctForSabha(sid, attValues);
      if (pct !== null) bucket.push(pct);
    }
    if (bucket.length) {
      var sum = 0;
      for (i = 0; i < bucket.length; i++) sum += bucket[i];
      out[w] = Math.round(sum / bucket.length);
    }
  }
  return out;
}

/**
 * Cumulative active devotees (visible) created on or before each month end.
 * @param {Array} devValues
 * @param {Object} user
 * @param {Date} now
 * @param {string} tz
 * @return {{ labels: string[], counts: number[] }}
 */
function _dashboardDevoteeGrowth(devValues, user, now, tz) {
  var labels = [];
  var counts = [];
  var m;
  for (m = 5; m >= 0; m--) {
    var ref = new Date(now.getFullYear(), now.getMonth() - m, 1);
    var lastDay = new Date(ref.getFullYear(), ref.getMonth() + 1, 0, 23, 59, 59);
    labels.push(Utilities.formatDate(ref, tz, 'MMM'));
    var c = 0;
    var i;
    for (i = 0; i < devValues.length; i++) {
      var row = devValues[i];
      if (!_devoteeRowVisibleForRole(row, user)) continue;
      if (String(row[DEVOTEES_COL.Status] || '').toLowerCase() !== DEVOTEE_STATUS.ACTIVE) continue;
      var created = row[DEVOTEES_COL.CreatedOn];
      if (!(created instanceof Date) || isNaN(created.getTime())) {
        c++;
        continue;
      }
      if (created <= lastDay) c++;
    }
    counts.push(c);
  }
  return { labels: labels, counts: counts };
}

function _dashboardBirthdayInfo(row, now, tz, userNameById, devById) {
  var dobDate = _dashboardParseDob_(row[DEVOTEES_COL.DOB]);
  if (!dobDate || isNaN(dobDate.getTime())) return null;
  var y = parseInt(Utilities.formatDate(now, tz, 'yyyy'), 10);
  var next = new Date(y, dobDate.getMonth(), dobDate.getDate());
  next.setHours(0, 0, 0, 0);
  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  if (next < today) next = new Date(y + 1, dobDate.getMonth(), dobDate.getDate());
  var daysAway = Math.round((next.getTime() - today.getTime()) / 86400000);
  var turns = next.getFullYear() - dobDate.getFullYear();
  var mobile = String(row[DEVOTEES_COL.Mobile] || '').trim();
  var digits = _dashboardDigitsOnly_(mobile);
  var waLink = digits
    ? 'https://wa.me/' +
      digits +
      '?text=' +
      encodeURIComponent('Jai Swaminarayan. Happy birthday! Stay blessed always.')
    : '';
  var karyakartaName = _dashboardResolveKaryakartaName_(
    String(row[DEVOTEES_COL.KaryakartaName] || row[DEVOTEES_COL.KaryakartaID] || ''),
    userNameById,
    devById
  );
  return {
    name: _devoteeFullName(row),
    dateLabel: Utilities.formatDate(next, tz, 'dd MMM'),
    daysAway: daysAway,
    turns: turns,
    mobile: mobile,
    waLink: waLink,
    karyakartaName: karyakartaName,
  };
}

function _dashboardParseDob_(dob) {
  if (!dob && dob !== 0) return null;
  if (dob instanceof Date && !isNaN(dob.getTime())) return new Date(dob.getFullYear(), dob.getMonth(), dob.getDate());
  if (typeof dob === 'number' && isFinite(dob)) {
    // Google Sheets serial date: days since 1899-12-30
    var ms = Math.round((dob - 25569) * 86400000);
    var dt = new Date(ms);
    if (!isNaN(dt.getTime())) return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
  }
  var s = String(dob).trim();
  if (!s) return null;
  var m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) {
    var d = parseInt(m[1], 10);
    var mo = parseInt(m[2], 10);
    // Prefer dd/mm/yyyy; if impossible, fallback mm/dd/yyyy.
    if (mo > 12 && d <= 12) {
      var tmp = d;
      d = mo;
      mo = tmp;
    }
    return new Date(parseInt(m[3], 10), mo - 1, d);
  }
  var parsed = new Date(s);
  if (!isNaN(parsed.getTime())) return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
  return null;
}

function _dashboardIsYuvakType_(typeText) {
  var t = String(typeText || '').toLowerCase();
  return t.indexOf('yuvak') !== -1 || t.indexOf('youth') !== -1;
}

function _dashboardResolveKaryakartaName_(kId, userNameById, devById) {
  var id = String(kId || '').trim();
  if (!id) return '—';
  if (userNameById && userNameById[id]) return userNameById[id];
  if (devById && devById[id]) return _devoteeFullName(devById[id]);
  return id;
}

function _dashboardDigitsOnly_(s) {
  return String(s || '').replace(/\D/g, '');
}

function _dashboardAgeBandCounts(devValues, user) {
  var out = { kids: 0, youth: 0, adult: 0, senior: 0, unknown: 0 };
  for (var i = 0; i < devValues.length; i++) {
    var row = devValues[i];
    if (!_devoteeRowVisibleForRole(row, user)) continue;
    var age = _devoteeAgeYears(row[DEVOTEES_COL.DOB]);
    if (age === null) out.unknown++;
    else if (age <= 12) out.kids++;
    else if (age <= 35) out.youth++;
    else if (age <= 59) out.adult++;
    else out.senior++;
  }
  return out;
}

function _dashboardUpcomingBirthdayWeeklyCounts(upcomingBirthdays) {
  var labels = ['W1', 'W2', 'W3', 'W4', 'W5'];
  var counts = [0, 0, 0, 0, 0];
  for (var i = 0; i < upcomingBirthdays.length; i++) {
    var d = parseInt(upcomingBirthdays[i].daysAway, 10);
    if (isNaN(d) || d < 1 || d > 35) continue;
    var idx = Math.min(4, Math.floor((d - 1) / 7));
    counts[idx]++;
  }
  return { labels: labels, counts: counts };
}

function _dashboardMotivationQuote(now, tz) {
  var pool = _dashboardSlogansFromSheet_();
  if (!pool.length) {
    pool = [
      { text: 'દિવ્યભાવ અને નમ્રતા એ જ સચ્ચી ભક્તિનો માર્ગ છે.', author: 'પ્રમુખસ્વામી મહારાજ' },
      { text: 'સ્વામિનારાયણ ભગવાન પર અખંડ શ્રદ્ધા રાખો, શાંતિ મળેશે.', author: 'ભગવાન સ્વામિનારાયણ' },
      { text: 'સેવા, સમર્પણ અને એકતા થી satsang ફુલેફાલે છે.', author: 'પ્રબોધ સ્વામી મહારાજ' },
    ];
  }
  var dateKey = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  var seed = 0;
  for (var i = 0; i < dateKey.length; i++) seed = (seed * 31 + dateKey.charCodeAt(i)) % 2147483647;
  return pool[seed % pool.length];
}

function _dashboardSlogansFromSheet_() {
  try {
    var ss = getSpreadsheet();
    var sh = ss.getSheetByName(SHEET_SLOGANS);
    if (!sh) return [];
    var last = sh.getLastRow();
    if (last < 2) return [];
    var rows = sh.getRange(2, 1, last - 1, 3).getValues();
    var out = [];
    for (var i = 0; i < rows.length; i++) {
      var txt = String(rows[i][0] || '').trim();
      if (!txt) continue;
      var author = String(rows[i][1] || '').trim() || 'Satsang';
      var active = String(rows[i][2] == null ? 'true' : rows[i][2]).toLowerCase();
      if (active === 'false' || active === '0' || active === 'no') continue;
      out.push({ text: txt, author: author });
    }
    return out;
  } catch (e) {
    logError('_dashboardSlogansFromSheet_', {}, e);
    return [];
  }
}

function hppEnsureSlogansSheet() {
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName(SHEET_SLOGANS);
  if (!sh) sh = ss.insertSheet(SHEET_SLOGANS);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, 3).setValues([['Slogan', 'Author', 'IsActive']]);
    sh.setFrozenRows(1);
  }
  return successResponse({ sheet: SHEET_SLOGANS });
}

/**
 * @param {Date} d
 * @param {string} tz
 * @return {Date} Start of week (Monday 00:00 in local calendar parts).
 */
function _dashboardStartOfWeek(d, tz) {
  var y = parseInt(Utilities.formatDate(d, tz, 'yyyy'), 10);
  var m = parseInt(Utilities.formatDate(d, tz, 'MM'), 10) - 1;
  var day = parseInt(Utilities.formatDate(d, tz, 'dd'), 10);
  var local = new Date(y, m, day);
  var dow = local.getDay();
  var diff = dow === 0 ? -6 : 1 - dow;
  local.setDate(local.getDate() + diff);
  local.setHours(0, 0, 0, 0);
  return local;
}

/**
 * @param {Array} row
 * @param {Object} user
 * @return {boolean}
 */
function _dashboardSabhaVisible(row, user) {
  var role = String(user.role || '').toLowerCase();
  if (role === ROLES.ADMIN) return true;
  var mid = String(row[SABHAS_COL.MandalID] || '').trim();
  if (
    role === ROLES.LEADER ||
    role === ROLES.SUB_LEADER ||
    role === ROLES.KARYAKARTA
  ) {
    return mid === String(user.mandalID || '').trim();
  }
  return false;
}

/**
 * @param {string} sabhaId
 * @param {Array} attValues
 * @return {number|null}
 */
function _dashboardAttendancePctForSabha(sabhaId, attValues) {
  var p = 0;
  var l = 0;
  var ab = 0;
  var i;
  for (i = 0; i < attValues.length; i++) {
    var ar = attValues[i];
    if (String(ar[ATTENDANCE_COL.SabhaID] || '').trim() !== sabhaId) continue;
    var st = String(ar[ATTENDANCE_COL.Status] || '').toLowerCase();
    if (st === ATTENDANCE_STATUS.PRESENT) p++;
    else if (st === ATTENDANCE_STATUS.LATE) l++;
    else if (st === ATTENDANCE_STATUS.ABSENT) ab++;
  }
  var t = p + l + ab;
  if (!t) return null;
  return Math.round(((p + l) / t) * 100);
}
