/**
 * HPP Seva Connect — Devotees CRUD, search, stats, audit.
 */

/**
 * @param {Object} filters mandalID, subLeaderID, karyakartaID, status, devoteeType, searchQuery
 */
function getDevotees(filters) {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');
    var bundle = _computeDevoteesBundle_(user, filters || {});
    return successResponse(bundle.devotees);
  } catch (e) {
    logError('getDevotees', filters, e);
    return errorResponse(e.message || String(e));
  }
}

function getDevoteesBundle(filters) {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');
    var f = filters || {};
    var cached = _getDevoteesBundleCache_(user, f);
    if (cached) {
      return successResponse(cached);
    }
    var bundle = _computeDevoteesBundle_(user, f);
    _setDevoteesBundleCache_(user, f, {
      devotees: bundle.devotees,
      stats: bundle.stats,
      paging: bundle.paging,
    });
    return successResponse({
      devotees: bundle.devotees,
      stats: bundle.stats,
      paging: bundle.paging,
    });
  } catch (e) {
    var errorId = logError('getDevoteesBundle', filters, e);
    return errorResponse((e.message || String(e)) + ' (Ref: ' + errorId + ')', errorId);
  }
}

function getDevoteeById(devoteeID) {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');
    var values = _getDevoteesRowsCached_();
    var found = _findDevoteeIndexInRows_(values, devoteeID);
    if (found < 0) return errorResponse('Devotee not found');
    var row = values[found];
    if (!_devoteeRowVisibleForRole(row, user)) return errorResponse('Access denied');
    return successResponse(_devoteeRowToObject(row, found + 2));
  } catch (e) {
    logError('getDevoteeById', { devoteeID: devoteeID }, e);
    return errorResponse(e.message || String(e));
  }
}

function addDevotee(data) {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');

    var first = _dStr(data, ['FirstName', 'firstName']);
    var last = _dStr(data, ['LastName', 'lastName']);
    if (!first || !last) return errorResponse('First name and last name are required');

    var devSheet = getSheet(SHEET_DEVOTEES);
    var famSheet = getSheet(SHEET_FAMILIES);

    var relation = _dStr(data, ['RelationWithHead', 'relationWithHead']) || 'Self';
    var isSelf = String(relation).toLowerCase() === 'self';

    var mandaId =
      _dStr(data, ['Mandal', 'mandal', 'MandaID', 'mandalID', 'MandalID']) || String(user.mandalID || '');
    var familyId = '';
    var familyName = _dStr(data, ['FamilyName', 'familyName']) || (first + ' ' + last).trim();

    if (isSelf) {
      familyId = _nextFamilyId_(devSheet);
    } else {
      var refId = _dStr(data, [
        'ReferenceLookupID',
        'referenceLookupID',
        'ReferenceDevoteeID',
        'referenceDevoteeID',
        'RelatedDevoteeID',
        'relatedDevoteeID',
      ]);
      if (!refId) return errorResponse('Reference devotee is required when relation is not Self');
      var famFromRef = _findFamilyIdByDevoteeId(devSheet, refId);
      if (!famFromRef) return errorResponse('Reference devotee not found');
      familyId = famFromRef;
    }

    var devoteeId = _nextDevoteeId_(devSheet);
    var now = new Date();
    var row = _emptyDevoteeRow();

    row[DEVOTEES_COL.DevoteeID] = devoteeId;
    row[DEVOTEES_COL.FamilyID] = familyId;
    row[DEVOTEES_COL.MandaID] = mandaId;
    row[DEVOTEES_COL.SubLeaderID] = _dStr(data, [
      'KaryakartaLeaderName',
      'karyakartaLeaderName',
      'SubLeaderID',
      'subLeaderID',
    ]);
    row[DEVOTEES_COL.KaryakartaID] = _dStr(data, ['KaryakartaName', 'karyakartaName', 'KaryakartaID', 'karyakartaID']);
    row[DEVOTEES_COL.FirstName] = first;
    row[DEVOTEES_COL.MiddleName] = _dStr(data, ['MiddleName', 'middleName']);
    row[DEVOTEES_COL.LastName] = last;
    row[DEVOTEES_COL.Gender] = _dStr(data, ['Gender', 'gender']);
    row[DEVOTEES_COL.DOB] = _parseDobCell(data.DOB || data.dob);
    var inPhoto = _dStr(data, ['Photo', 'photo', 'photoUrl']);
    var savedPhoto = _saveDevoteePhotoToDrive_(inPhoto, devoteeId);
    if (savedPhoto === '__PHOTO_UPLOAD_FAILED__') {
      return errorResponse('Photo upload failed. Please authorize Drive access and try again.');
    }
    row[DEVOTEES_COL.Photo] = _truncateForSheet_(savedPhoto, 2048);
    row[DEVOTEES_COL.Mobile] = _dStr(data, ['Mobile', 'mobile']);
    row[DEVOTEES_COL.WhatsApp] = _dStr(data, ['WhatsApp', 'whatsApp', 'whatsapp']);
    row[DEVOTEES_COL.Email] = _dStr(data, ['Email', 'email']);
    row[DEVOTEES_COL.Address] = _dStr(data, ['Address', 'address']);
    row[DEVOTEES_COL.NativePlace] = _dStr(data, ['NativePlace', 'nativePlace']);
    row[DEVOTEES_COL.DevoteeType] = _dStr(data, ['DevoteeType', 'devoteeType']);
    row[DEVOTEES_COL.Status] =
      _dStr(data, ['Status', 'status']) || DEVOTEE_STATUS.ACTIVE;
    row[DEVOTEES_COL.BloodGroup] = _dStr(data, ['BloodGroup', 'bloodGroup']);
    row[DEVOTEES_COL.IsElderly] = _dBool(data, ['IsElderly', 'isElderly']);
    row[DEVOTEES_COL.SpecialNeeds] = _dStr(data, ['SpecialNeeds', 'specialNeeds']);
    row[DEVOTEES_COL.EmergencyContact] = _dStr(data, ['EmergencyContact', 'emergencyContact']);
    row[DEVOTEES_COL.DikashaLevel] = _dStr(data, ['DikashaLevel', 'dikashaLevel', 'DikshaLevel']);
    row[DEVOTEES_COL.DikashaDate] = _parseDobCell(data.DikashaDate || data.dikashaDate);
    row[DEVOTEES_COL.Panchamrut] = _dStr(data, ['Panchamrut', 'panchamrut']);
    row[DEVOTEES_COL.Skills] = _truncateForSheet_(_mergeSkillsEducation(data));
    row[DEVOTEES_COL.WhatsAppOptedIn] = _dBool(data, ['WhatsAppOptedIn', 'whatsAppOptedIn']);
    row[DEVOTEES_COL.ReferenceDevoteeID] = isSelf
      ? ''
      : _dStr(data, ['Reference', 'reference', 'ReferenceName', 'referenceName', 'ReferenceDevoteeID', 'referenceDevoteeID']);
    row[DEVOTEES_COL.RelationWithHead] = relation;
    row[DEVOTEES_COL.LanguagePref] = _dStr(data, ['LanguagePref', 'languagePref']);
    row[DEVOTEES_COL.City] = _dStr(data, ['City', 'city']);
    row[DEVOTEES_COL.State] = _dStr(data, ['State', 'state']);
    row[DEVOTEES_COL.Country] = _dStr(data, ['Country', 'country']) || 'India';
    row[DEVOTEES_COL.DonationEnabled] = _dBool(data, ['DonationEnabled', 'donationEnabled']);
    row[DEVOTEES_COL.PANNumber] = _dStr(data, ['PANNumber', 'panNumber']);
    row[DEVOTEES_COL.CreatedBy] = user.userID || user.email || '';
    row[DEVOTEES_COL.CreatedOn] = now;
    row[DEVOTEES_COL.UpdatedOn] = now;

    if (isSelf) {
      var famRow = ['', '', '', '', '', ''];
      famRow[FAMILIES_COL.FamilyID] = familyId;
      famRow[FAMILIES_COL.HeadDevoteeID] = devoteeId;
      famRow[FAMILIES_COL.MandaID] = mandaId;
      famRow[FAMILIES_COL.FamilyName] = familyName;
      famRow[FAMILIES_COL.TotalMembers] = 1;
      famRow[FAMILIES_COL.CreatedOn] = now;
      famSheet.appendRow(famRow);
    }

    devSheet.appendRow(row);
    _appendAuditLog('INSERT', SHEET_DEVOTEES, devoteeId, null, row);
    _bustDevoteesCache_();

    return successResponse({ devoteeID: devoteeId, familyID: familyId });
  } catch (e) {
    logError('addDevotee', data, e);
    return errorResponse(e.message || String(e));
  }
}

function updateDevotee(data) {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');

    var rowIndex = parseInt(data && data._rowIndex, 10);
    if (!rowIndex || rowIndex < 2) return errorResponse('Invalid row index');

    var sheet = getSheet(SHEET_DEVOTEES);
    var oldRow = sheet.getRange(rowIndex, 1, 1, 38).getValues()[0];
    if (!_devoteeRowVisibleForRole(oldRow, user)) return errorResponse('Access denied');

    var newRow = oldRow.slice();
    _applyDevoteeDataToRow(newRow, data);
    if (newRow[DEVOTEES_COL.Photo] === '__PHOTO_UPLOAD_FAILED__') {
      return errorResponse('Photo upload failed. Please authorize Drive access and try again.');
    }
    newRow[DEVOTEES_COL.UpdatedOn] = new Date();

    sheet.getRange(rowIndex, 1, 1, 38).setValues([newRow]);
    _appendAuditLog('UPDATE', SHEET_DEVOTEES, String(oldRow[DEVOTEES_COL.DevoteeID]), oldRow, newRow);
    _bustDevoteesCache_();

    return successResponse({ message: 'Devotee updated' });
  } catch (e) {
    logError('updateDevotee', data, e);
    return errorResponse(e.message || String(e));
  }
}

function deleteDevotee(devoteeID, password) {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');

    var expected = String(PropertiesService.getScriptProperties().getProperty('DELETE_PASSWORD') || '').trim();
    var provided = String(password || '').trim();
    var role = String((user && user.role) || '').toLowerCase();
    if (expected) {
      if (provided !== expected) return errorResponse('Invalid delete password');
    } else if (role !== ROLES.ADMIN) {
      return errorResponse('Delete password is not configured. Contact admin.');
    }

    var sheet = getSheet(SHEET_DEVOTEES);
    var rowIndex = _findDevoteeRowIndex(sheet, devoteeID);
    if (!rowIndex) return errorResponse('Devotee not found');

    var oldRow = sheet.getRange(rowIndex, 1, 1, 38).getValues()[0];
    if (!_devoteeRowVisibleForRole(oldRow, user)) return errorResponse('Access denied');

    var newRow = oldRow.slice();
    newRow[DEVOTEES_COL.Status] = DEVOTEE_STATUS.INACTIVE;
    newRow[DEVOTEES_COL.UpdatedOn] = new Date();
    sheet.getRange(rowIndex, 1, 1, 38).setValues([newRow]);

    _appendAuditLog('SOFT_DELETE', SHEET_DEVOTEES, devoteeID, oldRow, newRow);
    _bustDevoteesCache_();

    return successResponse({ message: 'Devotee marked inactive' });
  } catch (e) {
    logError('deleteDevotee', { devoteeID: devoteeID }, e);
    return errorResponse(e.message || String(e));
  }
}

function _nextDevoteeId_(sheet) {
  return _nextIdWithPrefixAndWidth_(sheet, DEVOTEES_COL.DevoteeID, 'SEVAK', 5);
}

function _nextFamilyId_(sheet) {
  return _nextIdWithPrefixAndWidth_(sheet, DEVOTEES_COL.FamilyID, 'FAM', 4);
}

function _nextIdWithPrefixAndWidth_(sheet, columnIndex, prefix, width) {
  var p = String(prefix || '').trim();
  var w = Math.max(1, parseInt(width, 10) || 1);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return p + _padIdNumber(1, w);
  var col = (parseInt(columnIndex, 10) || 0) + 1;
  var vals = sheet.getRange(2, col, lastRow - 1, 1).getValues();
  var re = new RegExp('^' + _escapeRegExp(p) + '(\\d+)$', 'i');
  var maxNum = 0;
  for (var i = 0; i < vals.length; i++) {
    var s = String(vals[i][0] || '').trim();
    var m = s.match(re);
    if (!m) continue;
    var n = parseInt(m[1], 10);
    if (!isNaN(n) && n > maxNum) maxNum = n;
  }
  return p + _padIdNumber(maxNum + 1, w);
}

function searchDevotees(query) {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');

    var q = String(query || '').trim().toLowerCase();
    if (q.length < 2) return errorResponse('Enter at least 2 characters');

    var values = _getDevoteesRowsCached_();
    if (!values.length) return successResponse([]);
    var out = [];

    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      if (!_devoteeRowVisibleForRole(row, user)) continue;

      var id = String(row[DEVOTEES_COL.DevoteeID] || '').toLowerCase();
      var fam = String(row[DEVOTEES_COL.FamilyID] || '').toLowerCase();
      var mob = String(row[DEVOTEES_COL.Mobile] || '').toLowerCase();
      var name = _devoteeFullName(row).toLowerCase();

      if (id.indexOf(q) === -1 && fam.indexOf(q) === -1 && mob.indexOf(q) === -1 && name.indexOf(q) === -1) {
        continue;
      }

      out.push({
        DevoteeID: String(row[DEVOTEES_COL.DevoteeID] || '').trim(),
        Name: _devoteeFullName(row),
        Mobile: String(row[DEVOTEES_COL.Mobile] || '').trim(),
        FamilyID: String(row[DEVOTEES_COL.FamilyID] || '').trim(),
        _rowIndex: i + 2,
      });
      if (out.length >= 8) break;
    }

    return successResponse(out);
  } catch (e) {
    logError('searchDevotees', { query: query }, e);
    return errorResponse(e.message || String(e));
  }
}

function checkMobileExists(mobile, excludeDevoteeID) {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');

    var m = _normalizeMobile(mobile);
    if (!m) return successResponse({ exists: false, name: '' });

    var values = _getDevoteesRowsCached_();
    if (!values.length) return successResponse({ exists: false, name: '' });
    var ex = String(excludeDevoteeID || '').trim();

    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      if (!_devoteeRowVisibleForRole(row, user)) continue;
      var did = String(row[DEVOTEES_COL.DevoteeID] || '').trim();
      if (ex && did === ex) continue;
      if (_normalizeMobile(row[DEVOTEES_COL.Mobile]) === m) {
        return successResponse({ exists: true, name: _devoteeFullName(row) });
      }
    }
    return successResponse({ exists: false, name: '' });
  } catch (e) {
    logError('checkMobileExists', { mobile: mobile }, e);
    return errorResponse(e.message || String(e));
  }
}

function getFamilyMembers(familyID) {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');

    var fid = String(familyID || '').trim();
    if (!fid) return errorResponse('Family ID required');

    var values = _getDevoteesRowsCached_();
    if (!values.length) return successResponse([]);
    var out = [];

    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      if (String(row[DEVOTEES_COL.FamilyID] || '').trim() !== fid) continue;
      if (!_devoteeRowVisibleForRole(row, user)) continue;
      out.push(_devoteeRowToObject(row, i + 2));
    }
    return successResponse(out);
  } catch (e) {
    logError('getFamilyMembers', { familyID: familyID }, e);
    return errorResponse(e.message || String(e));
  }
}

function getDevoteeStats() {
  try {
    var user = getCurrentUser();
    if (!user) return errorResponse('Not signed in');
    var bundle = _computeDevoteesBundle_(user, {});
    return successResponse(bundle.stats);
  } catch (e) {
    logError('getDevoteeStats', {}, e);
    return errorResponse(e.message || String(e));
  }
}

function _computeDevoteesBundle_(user, f) {
  var values = _getDevoteesRowsCached_();
  if (!values.length) {
    return {
      devotees: [],
      stats: {
        total: 0, active: 0, families: 0, ambrish: 0, male: 0, female: 0, kids: 0, youth: 0, adult: 0, senior: 0
      },
      paging: { page: 0, pageSize: 50, total: 0, totalPages: 1, hasPrev: false, hasNext: false },
    };
  }
  var out = [];
  var page = parseInt(f && f.page, 10);
  if (isNaN(page) || page < 0) page = 0;
  var pageSize = parseInt(f && f.pageSize, 10);
  if (isNaN(pageSize) || pageSize <= 0) pageSize = 50;
  if (pageSize > 500) pageSize = 500;
  var start = page * pageSize;
  var end = start + pageSize;
  var filteredCount = 0;
  var stats = {
    total: 0,
    active: 0,
    families: 0,
    ambrish: 0,
    male: 0,
    female: 0,
    kids: 0,
    youth: 0,
    adult: 0,
    senior: 0,
  };
  var famSet = {};
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (!_devoteeRowVisibleForRole(row, user)) continue;
    stats.total++;
    var st = String(row[DEVOTEES_COL.Status] || '').toLowerCase();
    if (st === DEVOTEE_STATUS.ACTIVE) stats.active++;
    var fid = String(row[DEVOTEES_COL.FamilyID] || '').trim();
    if (fid) famSet[fid] = true;
    var dtype = String(row[DEVOTEES_COL.DevoteeType] || '').toLowerCase();
    if (dtype.indexOf('ambrish') !== -1) stats.ambrish++;
    var g = String(row[DEVOTEES_COL.Gender] || '').toLowerCase();
    if (g === 'male' || g === 'm') stats.male++;
    else if (g === 'female' || g === 'f') stats.female++;
    var age = _devoteeAgeYears(row[DEVOTEES_COL.DOB]);
    if (age !== null) {
      if (age <= 12) stats.kids++;
      else if (age <= 35) stats.youth++;
      else if (age <= 59) stats.adult++;
      else stats.senior++;
    }
    if (!_devoteeRowMatchesFilters(row, f || {})) continue;
    if (filteredCount >= start && filteredCount < end) out.push(_devoteeRowToObject(row, i + 2));
    filteredCount++;
  }
  for (var fk in famSet) {
    if (Object.prototype.hasOwnProperty.call(famSet, fk)) stats.families++;
  }
  var totalPages = Math.max(1, Math.ceil(filteredCount / pageSize));
  var safePage = page;
  if (safePage > totalPages - 1) safePage = totalPages - 1;
  return {
    devotees: out,
    stats: stats,
    paging: {
      page: safePage,
      pageSize: pageSize,
      total: filteredCount,
      totalPages: totalPages,
      hasPrev: safePage > 0,
      hasNext: safePage < totalPages - 1,
    },
  };
}

var DEVOTEES_CACHE_TTL_SEC = 25;
var DEVOTEES_CACHE_VER_KEY = 'DEVOTEES_CACHE_VER';
var DEVOTEES_DATA_CACHE_KEY = 'aksharSetuDevotees_v4';
var DEVOTEES_DATA_PROP_KEY = 'devotees_data';
var DEVOTEES_DATA_CACHE_TTL_SEC = 21600;
var DEVOTEES_CACHE_CHUNK_SIZE = 90000;
var DEVOTEES_PROP_CHUNK_SIZE = 8000;
var DASHBOARD_SNAPSHOT_CACHE_KEY = 'aksharSetuDashboardSnapshot_v1';
var DASHBOARD_SNAPSHOT_PROP_KEY = 'dashboard_snapshot_v1';
var DEVOTEES_META_PROP_KEY = 'devotees_cache_meta_v1';

function _devoteesCacheVersion_() {
  var props = PropertiesService.getScriptProperties();
  return String(props.getProperty(DEVOTEES_CACHE_VER_KEY) || '1');
}

function _bustDevoteesCache_() {
  var props = PropertiesService.getScriptProperties();
  var cur = parseInt(props.getProperty(DEVOTEES_CACHE_VER_KEY) || '1', 10);
  if (!cur || isNaN(cur)) cur = 1;
  props.setProperty(DEVOTEES_CACHE_VER_KEY, String(cur + 1));
  try {
    refreshCache();
  } catch (e) {
    logError('_bustDevoteesCache_.refreshCache', {}, e);
  }
}

/**
 * Read once and persist devotees payload in cache + script properties.
 * @return {{ success: boolean, data?: Object, error?: string }}
 */
function refreshCache() {
  try {
    var sheet = getSheet(SHEET_DEVOTEES);
    var all = sheet.getDataRange().getValues();
    var headers = all.length ? all[0] : [];
    var rows = all.length > 1 ? all.slice(1) : [];
    var payload = {
      headers: headers,
      rows: rows,
      refreshedAt: new Date().toISOString(),
    };
    var raw = JSON.stringify(payload);
    _storeChunkedInCache_(DEVOTEES_DATA_CACHE_KEY, raw, DEVOTEES_DATA_CACHE_TTL_SEC, DEVOTEES_CACHE_CHUNK_SIZE);
    // Do NOT store full devotees payload in ScriptProperties; it exceeds quota on larger datasets.
    PropertiesService.getScriptProperties().setProperty(
      DEVOTEES_META_PROP_KEY,
      JSON.stringify({ refreshedAt: payload.refreshedAt, rows: rows.length })
    );
    var snapshot = _buildDashboardSnapshotFromRows_(rows);
    var snapshotRaw = JSON.stringify(snapshot);
    _storeChunkedInCache_(DASHBOARD_SNAPSHOT_CACHE_KEY, snapshotRaw, DEVOTEES_DATA_CACHE_TTL_SEC, DEVOTEES_CACHE_CHUNK_SIZE);
    try {
      _storeChunkedInProps_(DASHBOARD_SNAPSHOT_PROP_KEY, snapshotRaw, DEVOTEES_PROP_CHUNK_SIZE);
    } catch (ignoreSnapshotPropsQuota) {
      // Snapshot in cache is sufficient for speed-first mode.
    }
    console.log('Devotees cache refreshed at ' + payload.refreshedAt + ' rows=' + rows.length);
    return successResponse({ refreshedAt: payload.refreshedAt, rows: rows.length });
  } catch (e) {
    var errorId = logError('refreshCache', {}, e);
    return errorResponse((e.message || String(e)) + ' (Ref: ' + errorId + ')', errorId);
  }
}

/**
 * Installable trigger handler: refresh devotees cache on sheet edit.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEdit(e) {
  try {
    var range = e && e.range;
    var sheet = range && range.getSheet ? range.getSheet() : null;
    if (!sheet) return;
    if (String(sheet.getName() || '').trim() !== SHEET_DEVOTEES) return;
    _bustDevoteesCache_();
  } catch (err) {
    logError('onEdit', {}, err);
  }
}

/**
 * One-time setup: populate cache and install onEdit trigger.
 * @return {{ success: boolean, data?: Object, error?: string }}
 */
function initializeCache() {
  try {
    _cleanupLegacyDevoteesProps_();
    var init = refreshCache();
    if (!init || !init.success) return init;
    var ss = getSpreadsheet();
    var triggers = ScriptApp.getProjectTriggers();
    var hasOnEdit = false;
    for (var i = 0; i < triggers.length; i++) {
      var t = triggers[i];
      if (
        t.getHandlerFunction &&
        t.getHandlerFunction() === 'onEdit' &&
        t.getEventType &&
        t.getEventType() === ScriptApp.EventType.ON_EDIT
      ) {
        hasOnEdit = true;
        break;
      }
    }
    if (!hasOnEdit) {
      ScriptApp.newTrigger('onEdit').forSpreadsheet(ss).onEdit().create();
    }
    return successResponse({ cacheInitialized: true, triggerExists: true });
  } catch (e) {
    var errorId = logError('initializeCache', {}, e);
    return errorResponse((e.message || String(e)) + ' (Ref: ' + errorId + ')', errorId);
  }
}

/**
 * Returns devotees cache payload. Reads sheet only when cache/property missing.
 * @return {{ headers: Array, rows: Array, refreshedAt: string }}
 */
function getDevoteesCacheData() {
  var raw = '';
  raw = _readChunkedFromCache_(DEVOTEES_DATA_CACHE_KEY);
  if (!raw) {
    var init = refreshCache();
    if (!init || !init.success) return { headers: [], rows: [], refreshedAt: '' };
    raw = _readChunkedFromCache_(DEVOTEES_DATA_CACHE_KEY);
  }
  try {
    var parsed = JSON.parse(raw);
    return {
      headers: parsed && parsed.headers ? parsed.headers : [],
      rows: parsed && parsed.rows ? parsed.rows : [],
      refreshedAt: parsed && parsed.refreshedAt ? parsed.refreshedAt : '',
    };
  } catch (e) {
    var refill = refreshCache();
    if (!refill || !refill.success) return { headers: [], rows: [], refreshedAt: '' };
    try {
      var reRaw = _readChunkedFromCache_(DEVOTEES_DATA_CACHE_KEY) || '';
      var reparsed = JSON.parse(reRaw || '{}');
      return {
        headers: reparsed && reparsed.headers ? reparsed.headers : [],
        rows: reparsed && reparsed.rows ? reparsed.rows : [],
        refreshedAt: reparsed && reparsed.refreshedAt ? reparsed.refreshedAt : '',
      };
    } catch (ignore) {
      return { headers: [], rows: [], refreshedAt: '' };
    }
  }
}

function _cleanupLegacyDevoteesProps_() {
  try {
    var props = PropertiesService.getScriptProperties();
    var metaRaw = props.getProperty(_chunkMetaKey_(DEVOTEES_DATA_PROP_KEY));
    if (!metaRaw) return;
    var count = 0;
    try { count = Number(JSON.parse(metaRaw).count || 0); } catch (ignoreCount) {}
    props.deleteProperty(_chunkMetaKey_(DEVOTEES_DATA_PROP_KEY));
    for (var i = 0; i < count; i++) {
      props.deleteProperty(_chunkPartKey_(DEVOTEES_DATA_PROP_KEY, i));
    }
  } catch (e) {
    logError('_cleanupLegacyDevoteesProps_', {}, e);
  }
}

function _chunkMetaKey_(baseKey) {
  return baseKey + '__meta';
}

function _chunkPartKey_(baseKey, idx) {
  return baseKey + '__part_' + idx;
}

function _splitChunks_(raw, size) {
  var out = [];
  var s = String(raw || '');
  if (!s) return [''];
  for (var i = 0; i < s.length; i += size) out.push(s.slice(i, i + size));
  return out;
}

function _storeChunkedInCache_(baseKey, raw, ttlSec, chunkSize) {
  var cache = CacheService.getScriptCache();
  var chunks = _splitChunks_(raw, chunkSize);
  var priorMetaRaw = cache.get(_chunkMetaKey_(baseKey));
  var priorCount = 0;
  if (priorMetaRaw) {
    try { priorCount = Number(JSON.parse(priorMetaRaw).count || 0); } catch (ignore) {}
  }
  for (var i = 0; i < chunks.length; i++) {
    cache.put(_chunkPartKey_(baseKey, i), chunks[i], ttlSec);
  }
  cache.put(_chunkMetaKey_(baseKey), JSON.stringify({ count: chunks.length }), ttlSec);
  for (var j = chunks.length; j < priorCount; j++) {
    cache.remove(_chunkPartKey_(baseKey, j));
  }
}

function _readChunkedFromCache_(baseKey) {
  var cache = CacheService.getScriptCache();
  var metaRaw = cache.get(_chunkMetaKey_(baseKey));
  if (!metaRaw) return '';
  var count = 0;
  try { count = Number(JSON.parse(metaRaw).count || 0); } catch (ignore) { return ''; }
  if (!count || count < 1) return '';
  var out = [];
  for (var i = 0; i < count; i++) {
    var p = cache.get(_chunkPartKey_(baseKey, i));
    if (p === null || p === undefined) return '';
    out.push(p);
  }
  return out.join('');
}

function _storeChunkedInProps_(baseKey, raw, chunkSize) {
  var props = PropertiesService.getScriptProperties();
  var chunks = _splitChunks_(raw, chunkSize);
  var priorMetaRaw = props.getProperty(_chunkMetaKey_(baseKey));
  var priorCount = 0;
  if (priorMetaRaw) {
    try { priorCount = Number(JSON.parse(priorMetaRaw).count || 0); } catch (ignore) {}
  }
  var bulk = {};
  for (var i = 0; i < chunks.length; i++) {
    bulk[_chunkPartKey_(baseKey, i)] = chunks[i];
  }
  bulk[_chunkMetaKey_(baseKey)] = JSON.stringify({ count: chunks.length });
  props.setProperties(bulk, false);
  if (priorCount > chunks.length) {
    for (var j = chunks.length; j < priorCount; j++) {
      props.deleteProperty(_chunkPartKey_(baseKey, j));
    }
  }
}

function _readChunkedFromProps_(baseKey) {
  var props = PropertiesService.getScriptProperties();
  var metaRaw = props.getProperty(_chunkMetaKey_(baseKey));
  if (!metaRaw) return '';
  var count = 0;
  try { count = Number(JSON.parse(metaRaw).count || 0); } catch (ignore) { return ''; }
  if (!count || count < 1) return '';
  var out = [];
  for (var i = 0; i < count; i++) {
    var p = props.getProperty(_chunkPartKey_(baseKey, i));
    if (p === null || p === undefined) return '';
    out.push(p);
  }
  return out.join('');
}

function _getDevoteesRowsCached_() {
  var data = getDevoteesCacheData();
  var rows = (data && data.rows) ? data.rows : [];
  if (rows && rows.length) return rows;
  // Safety fallback: if chunked cache/properties are unavailable, hydrate from sheet once.
  try {
    var sheet = getSheet(SHEET_DEVOTEES);
    var all = sheet.getDataRange().getValues();
    var directRows = all.length > 1 ? all.slice(1) : [];
    if (directRows.length) {
      try { refreshCache(); } catch (ignoreRefresh) {}
      return directRows;
    }
  } catch (e) {
    logError('_getDevoteesRowsCached_.fallback', {}, e);
  }
  return [];
}

function getDashboardSnapshotCached() {
  var raw = _readChunkedFromCache_(DASHBOARD_SNAPSHOT_CACHE_KEY);
  if (!raw) {
    raw = _readChunkedFromProps_(DASHBOARD_SNAPSHOT_PROP_KEY);
    if (raw) {
      _storeChunkedInCache_(DASHBOARD_SNAPSHOT_CACHE_KEY, raw, DEVOTEES_DATA_CACHE_TTL_SEC, DEVOTEES_CACHE_CHUNK_SIZE);
    }
  }
  if (!raw) {
    var init = refreshCache();
    if (!init || !init.success) return null;
    raw = _readChunkedFromProps_(DASHBOARD_SNAPSHOT_PROP_KEY);
  }
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function _buildDashboardSnapshotFromRows_(rows) {
  var totalVisible = 0;
  var activeDevotees = 0;
  var familiesSet = {};
  var todayBirthdays = [];
  var upcomingBirthdays = [];
  var karyakartaYuvakMap = {};
  var now = new Date();
  var tz = _appTimeZone();
  var list = rows || [];
  for (var i = 0; i < list.length; i++) {
    var row = list[i];
    totalVisible++;
    var st = String(row[DEVOTEES_COL.Status] || '').trim().toLowerCase();
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
      var b = _dashboardBirthdayInfo(row, now, tz, {}, {});
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
  return {
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
    motivationQuote: _devoteesMotivationQuote_(),
    snapshotAt: new Date().toISOString(),
    cacheStamp: _devoteesCacheVersion_(),
  };
}

function _devoteesMotivationQuote_() {
  try {
    var cacheKey = 'dashboard:slogan:daily:v1';
    var cache = CacheService.getScriptCache();
    var cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (ignoreCached) {}
    }
    var quote = { text: 'હે હરિ ! બસ એક, તું રાજી થા.', author: 'Satsang' };
    try {
      if (typeof _dashboardMotivationQuote === 'function') {
        var q = _dashboardMotivationQuote(new Date(), _appTimeZone());
        if (q && q.text) quote = q;
      }
    } catch (ignoreQuote) {}
    try {
      cache.put(cacheKey, JSON.stringify(quote), 3600);
    } catch (ignorePut) {}
    return quote;
  } catch (e) {
    return { text: 'હે હરિ ! બસ એક, તું રાજી થા.', author: 'Satsang' };
  }
}

function getDevoteesCacheStamp() {
  try {
    return successResponse({
      cacheStamp: _devoteesCacheVersion_(),
    });
  } catch (e) {
    var errorId = logError('getDevoteesCacheStamp', {}, e);
    return errorResponse((e.message || String(e)) + ' (Ref: ' + errorId + ')', errorId);
  }
}

function _devoteesCacheKey_(user, filters) {
  var role = String((user && user.role) || '');
  var uid = String((user && (user.userID || user.email)) || '');
  var mandal = String((user && user.mandalID) || '');
  var f = filters || {};
  var fKey = [
    String(f.mandalID || ''),
    String(f.subLeaderID || ''),
    String(f.karyakartaID || ''),
    String(f.status || ''),
    String(f.devoteeType || ''),
    String(f.searchQuery || ''),
    String(f.gender || ''),
    String(f.bloodGroup || ''),
    String(f.ageBand || ''),
    String(f.page || ''),
    String(f.pageSize || ''),
  ].join('|');
  return 'dev_bundle|' + _devoteesCacheVersion_() + '|' + role + '|' + uid + '|' + mandal + '|' + fKey;
}

function _getDevoteesBundleCache_(user, filters) {
  try {
    var cache = CacheService.getScriptCache();
    var key = _devoteesCacheKey_(user, filters);
    var raw = cache.get(key);
    if (!raw) return null;
    var parsed = JSON.parse(raw);
    if (!parsed || !parsed.devotees || !parsed.stats) return null;
    return parsed;
  } catch (e) {
    logError('_getDevoteesBundleCache_', {}, e);
    return null;
  }
}

function _setDevoteesBundleCache_(user, filters, payload) {
  try {
    var cache = CacheService.getScriptCache();
    var key = _devoteesCacheKey_(user, filters);
    cache.put(key, JSON.stringify(payload), DEVOTEES_CACHE_TTL_SEC);
  } catch (e) {
    logError('_setDevoteesBundleCache_', {}, e);
  }
}

/* ---------- helpers ---------- */

function _devoteeRowVisibleForRole(row, user) {
  var role = String(user.role || '').toLowerCase();
  if (role === ROLES.ADMIN) return true;
  if (role === ROLES.LEADER) {
    return (
      String(row[DEVOTEES_COL.MandaID] || '').trim().toLowerCase() ===
      String(user.mandalID || '').trim().toLowerCase()
    );
  }
  if (role === ROLES.SUB_LEADER) {
    var sl = String(row[DEVOTEES_COL.SubLeaderID] || '').trim().toLowerCase();
    var uid = String(user.userID || '').trim().toLowerCase();
    var uname = String(user.name || '').trim().toLowerCase();
    return sl === uid || (!!uname && sl === uname);
  }
  if (role === ROLES.KARYAKARTA) {
    var k = String(row[DEVOTEES_COL.KaryakartaID] || '').trim().toLowerCase();
    var uid = String(user.userID || '').trim().toLowerCase();
    var uname = String(user.name || '').trim().toLowerCase();
    return k === uid || (!!uname && k === uname);
  }
  return false;
}

function _devoteeRowMatchesFilters(row, f) {
  if (
    f.mandalID &&
    String(row[DEVOTEES_COL.MandaID] || '').trim().toLowerCase() !== String(f.mandalID).trim().toLowerCase()
  ) {
    return false;
  }
  if (
    f.subLeaderID &&
    String(row[DEVOTEES_COL.SubLeaderID] || '').trim().toLowerCase() !== String(f.subLeaderID).trim().toLowerCase()
  ) {
    return false;
  }
  if (f.karyakartaID && String(row[DEVOTEES_COL.KaryakartaID] || '').trim() !== String(f.karyakartaID).trim()) {
    return false;
  }
  if (f.status && String(f.status).toLowerCase() !== 'all') {
    if (String(row[DEVOTEES_COL.Status] || '').toLowerCase() !== String(f.status).toLowerCase()) return false;
  }
  if (f.devoteeType && String(f.devoteeType).toLowerCase() !== 'all') {
    var dt = String(row[DEVOTEES_COL.DevoteeType] || '').toLowerCase();
    if (dt.indexOf(String(f.devoteeType).toLowerCase()) === -1) return false;
  }
  if (f.searchQuery) {
    var q = String(f.searchQuery).trim().toLowerCase();
    if (q) {
      var hay =
        _devoteeFullName(row).toLowerCase() +
        ' ' +
        String(row[DEVOTEES_COL.Mobile] || '').toLowerCase() +
        ' ' +
        String(row[DEVOTEES_COL.DevoteeID] || '').toLowerCase() +
        ' ' +
        String(row[DEVOTEES_COL.FamilyID] || '').toLowerCase() +
        ' ' +
        String(row[DEVOTEES_COL.Address] || '').toLowerCase() +
        ' ' +
        String(row[DEVOTEES_COL.NativePlace] || '').toLowerCase() +
        ' ' +
        String(row[DEVOTEES_COL.MandaID] || '').toLowerCase() +
        ' ' +
        String(row[DEVOTEES_COL.DevoteeType] || '').toLowerCase();
      if (hay.indexOf(q) === -1) return false;
    }
  }
  if (f.gender && String(f.gender).toLowerCase() !== 'all') {
    var g = String(row[DEVOTEES_COL.Gender] || '').toLowerCase();
    if (g !== String(f.gender).toLowerCase()) return false;
  }
  if (f.bloodGroup && String(f.bloodGroup).toLowerCase() !== 'all') {
    var bg = String(row[DEVOTEES_COL.BloodGroup] || '').toUpperCase().trim();
    if (bg !== String(f.bloodGroup).toUpperCase().trim()) return false;
  }
  if (f.ageBand && String(f.ageBand).toLowerCase() !== 'all') {
    var age = _devoteeAgeYears(row[DEVOTEES_COL.DOB]);
    var band = 'unknown';
    if (age !== null) {
      if (age <= 12) band = 'kids';
      else if (age <= 35) band = 'youth';
      else if (age <= 59) band = 'adult';
      else band = 'senior';
    }
    if (band !== String(f.ageBand).toLowerCase()) return false;
  }
  return true;
}

function _devoteeFullName(row) {
  return [row[DEVOTEES_COL.FirstName], row[DEVOTEES_COL.MiddleName], row[DEVOTEES_COL.LastName]]
    .map(function (p) {
      return String(p || '').trim();
    })
    .filter(function (p) {
      return p;
    })
    .join(' ');
}

function _devoteeRowToObject(row, rowIndex1) {
  var keys = Object.keys(DEVOTEES_COL);
  var o = { _rowIndex: rowIndex1 };
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    o[k] = _sanitizeDevoteeCell(row[DEVOTEES_COL[k]], k);
  }
  o.DisplayName = [o.FirstName, o.MiddleName, o.LastName]
    .map(function (p) {
      return String(p || '').trim();
    })
    .filter(function (p) {
      return p;
    })
    .join(' ');
  return o;
}

/**
 * Ensures values are JSON-serializable for google.script.run (no RichText objects, etc.).
 * @param {*} v
 * @param {string} colKey
 * @return {string|number|boolean}
 */
function _sanitizeDevoteeCell(v, colKey) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return '';
    if (colKey === 'DOB' || colKey === 'DikashaDate') return formatDate(v);
    return formatDateTime(v);
  }
  var t = typeof v;
  if (t === 'string' || t === 'number') return v;
  if (t === 'boolean') return v;
  if (v && typeof v.getText === 'function') {
    try {
      return v.getText();
    } catch (ignore) {
      return '';
    }
  }
  try {
    return String(v);
  } catch (e2) {
    return '';
  }
}

function _findDevoteeRowIndex(sheet, devoteeID) {
  var id = String(devoteeID || '').trim();
  if (!id) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var col = DEVOTEES_COL.DevoteeID + 1;
  var vals = sheet.getRange(2, col, lastRow - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim() === id) return i + 2;
  }
  return 0;
}

function _findDevoteeIndexInRows_(rows, devoteeID) {
  var target = String(devoteeID || '').trim();
  if (!target) return -1;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][DEVOTEES_COL.DevoteeID] || '').trim() === target) return i;
  }
  return -1;
}

function _findFamilyIdByDevoteeId(sheet, devoteeId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return '';
  var ids = sheet.getRange(2, DEVOTEES_COL.DevoteeID + 1, lastRow - 1, 1).getValues();
  var fams = sheet.getRange(2, DEVOTEES_COL.FamilyID + 1, lastRow - 1, 1).getValues();
  var target = String(devoteeId || '').trim();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === target) return String(fams[i][0] || '').trim();
  }
  return '';
}

function _emptyDevoteeRow() {
  var a = [];
  for (var i = 0; i < 38; i++) a[i] = '';
  return a;
}

function _dStr(data, keys) {
  if (!data) return '';
  for (var i = 0; i < keys.length; i++) {
    var v = data[keys[i]];
    if (v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim();
  }
  return '';
}

function _dBool(data, keys) {
  for (var i = 0; i < keys.length; i++) {
    var v = data[keys[i]];
    if (v === true || v === false) return v;
    if (v === 'true' || v === 'TRUE' || v === 1 || v === '1') return true;
    if (v === 'false' || v === 'FALSE' || v === 0 || v === '0') return false;
  }
  return false;
}

function _parseDobCell(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  var s = String(v).trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    var d = new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
    if (!isNaN(d.getTime())) return d;
  }
  var iso = new Date(s);
  if (!isNaN(iso.getTime())) return iso;
  return s;
}

function _mergeSkillsEducation(data) {
  var base = _dStr(data, ['Skills', 'skills']);
  var edu = _dStr(data, ['Qualification', 'qualification', 'Education', 'education']);
  var occ = _dStr(data, ['Occupation', 'occupation']);
  var parts = [];
  if (base) parts.push(base);
  if (edu) parts.push('Qualification: ' + edu);
  if (occ) parts.push('Occupation: ' + occ);
  return parts.join(' | ');
}

function _applyDevoteeDataToRow(row, data) {
  if (!data) return;
  var map = [
    ['Mandal', DEVOTEES_COL.MandaID, ['Mandal', 'mandal', 'MandaID', 'mandalID']],
    [
      'KaryakartaLeaderName',
      DEVOTEES_COL.SubLeaderID,
      ['KaryakartaLeaderName', 'karyakartaLeaderName', 'SubLeaderID', 'subLeaderID'],
    ],
    ['KaryakartaName', DEVOTEES_COL.KaryakartaID, ['KaryakartaName', 'karyakartaName', 'KaryakartaID', 'karyakartaID']],
    ['FirstName', DEVOTEES_COL.FirstName, ['FirstName', 'firstName']],
    ['MiddleName', DEVOTEES_COL.MiddleName, ['MiddleName', 'middleName']],
    ['LastName', DEVOTEES_COL.LastName, ['LastName', 'lastName']],
    ['Gender', DEVOTEES_COL.Gender, ['Gender', 'gender']],
    ['Mobile', DEVOTEES_COL.Mobile, ['Mobile', 'mobile']],
    ['WhatsApp', DEVOTEES_COL.WhatsApp, ['WhatsApp', 'whatsApp']],
    ['Email', DEVOTEES_COL.Email, ['Email', 'email']],
    ['Address', DEVOTEES_COL.Address, ['Address', 'address']],
    ['NativePlace', DEVOTEES_COL.NativePlace, ['NativePlace', 'nativePlace']],
    ['DevoteeType', DEVOTEES_COL.DevoteeType, ['DevoteeType', 'devoteeType']],
    ['Status', DEVOTEES_COL.Status, ['Status', 'status']],
    ['BloodGroup', DEVOTEES_COL.BloodGroup, ['BloodGroup', 'bloodGroup']],
    ['SpecialNeeds', DEVOTEES_COL.SpecialNeeds, ['SpecialNeeds', 'specialNeeds']],
    ['EmergencyContact', DEVOTEES_COL.EmergencyContact, ['EmergencyContact', 'emergencyContact']],
    ['DikashaLevel', DEVOTEES_COL.DikashaLevel, ['DikashaLevel', 'dikashaLevel']],
    ['Panchamrut', DEVOTEES_COL.Panchamrut, ['Panchamrut', 'panchamrut']],
    ['LanguagePref', DEVOTEES_COL.LanguagePref, ['LanguagePref', 'languagePref']],
    ['City', DEVOTEES_COL.City, ['City', 'city']],
    ['State', DEVOTEES_COL.State, ['State', 'state']],
    ['Country', DEVOTEES_COL.Country, ['Country', 'country']],
    ['PANNumber', DEVOTEES_COL.PANNumber, ['PANNumber', 'panNumber']],
    ['Reference', DEVOTEES_COL.ReferenceDevoteeID, [
      'Reference',
      'reference',
      'ReferenceName',
      'referenceName',
      'ReferenceDevoteeID',
      'referenceDevoteeID',
      'RelatedDevoteeID',
    ]],
    ['RelationWithHead', DEVOTEES_COL.RelationWithHead, ['RelationWithHead', 'relationWithHead']],
    ['FamilyID', DEVOTEES_COL.FamilyID, ['FamilyID', 'familyID']],
  ];

  for (var i = 0; i < map.length; i++) {
    var keys = map[i][2];
    var has = false;
    var val = '';
    for (var j = 0; j < keys.length; j++) {
      if (Object.prototype.hasOwnProperty.call(data, keys[j])) {
        has = true;
        val = data[keys[j]];
        break;
      }
    }
    if (has) row[map[i][1]] = _truncateForSheet_(val === null || val === undefined ? '' : val);
  }
  if (data.Photo !== undefined || data.photo !== undefined || data.photoUrl !== undefined) {
    var inPhoto =
      data.Photo !== undefined
        ? data.Photo
        : data.photo !== undefined
          ? data.photo
          : data.photoUrl;
    var devId = row[DEVOTEES_COL.DevoteeID] || '';
    row[DEVOTEES_COL.Photo] = _truncateForSheet_(_saveDevoteePhotoToDrive_(inPhoto, devId), 2048);
  }

  if (data.DOB !== undefined || data.dob !== undefined) {
    row[DEVOTEES_COL.DOB] = _parseDobCell(data.DOB !== undefined ? data.DOB : data.dob);
  }
  if (data.DikashaDate !== undefined || data.dikashaDate !== undefined) {
    row[DEVOTEES_COL.DikashaDate] = _parseDobCell(
      data.DikashaDate !== undefined ? data.DikashaDate : data.dikashaDate
    );
  }
  if (data.IsElderly !== undefined || data.isElderly !== undefined) {
    row[DEVOTEES_COL.IsElderly] = !!(data.IsElderly !== undefined ? data.IsElderly : data.isElderly);
  }
  if (data.DonationEnabled !== undefined || data.donationEnabled !== undefined) {
    row[DEVOTEES_COL.DonationEnabled] = !!(data.DonationEnabled !== undefined
      ? data.DonationEnabled
      : data.donationEnabled);
  }
  if (data.WhatsAppOptedIn !== undefined || data.whatsAppOptedIn !== undefined) {
    row[DEVOTEES_COL.WhatsAppOptedIn] = !!(data.WhatsAppOptedIn !== undefined
      ? data.WhatsAppOptedIn
      : data.whatsAppOptedIn);
  }

  var mergedSkills = _mergeSkillsEducation(data);
  if (mergedSkills || data.Qualification || data.qualification || data.Education || data.education || data.Occupation || data.occupation || data.Skills || data.skills) {
    row[DEVOTEES_COL.Skills] = _truncateForSheet_(mergedSkills);
  }
}

/** Google Sheets max cell length is 50,000 characters — stay safely under. */
var SHEET_CELL_MAX = 49900;

function _truncateForSheet_(val, maxLen) {
  var max = maxLen == null ? SHEET_CELL_MAX : maxLen;
  var s = val === null || val === undefined ? '' : String(val);
  if (s.length <= max) return s;
  return s.substring(0, Math.max(0, max - 3)) + '...';
}

function _appPhotosFolder_() {
  var appName =
    PropertiesService.getScriptProperties().getProperty('APP_NAME') ||
    'Hari Prabodham Parivar HPP - Akshar Setu';
  var rootName = String(appName).trim() || 'Akshar Setu';
  var folders = DriveApp.getFoldersByName(rootName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(rootName);
}

function _saveDevoteePhotoToDrive_(photoVal, devoteeId) {
  var raw = String(photoVal || '').trim();
  if (!raw) return '';
  /** Never persist base64 / huge blobs to a sheet cell — only short URLs or empty. */
  if (!/^data:image\//i.test(raw)) {
    if (/^https?:\/\//i.test(raw)) return raw;
    return '';
  }
  try {
    var m = raw.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,([\s\S]+)$/);
    if (!m) return '__PHOTO_UPLOAD_FAILED__';
    var mime = m[1];
    var b64 = m[2].replace(/\s/g, '');
    var ext = 'jpg';
    if (mime.indexOf('png') !== -1) ext = 'png';
    else if (mime.indexOf('webp') !== -1) ext = 'webp';
    var bytes = Utilities.base64Decode(b64);
    var idPart = String(devoteeId || 'DEV').replace(/[^A-Za-z0-9_-]/g, '');
    var fileName = 'devotee_' + idPart + '_' + new Date().getTime() + '.' + ext;
    var blob = Utilities.newBlob(bytes, mime, fileName);
    var folder = _appPhotosFolder_();
    var file = folder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      /** Some Google Workspace domains block public link sharing; keep private URL fallback. */
      logError('_saveDevoteePhotoToDrive_setSharing', { devoteeId: devoteeId }, shareErr);
    }
    var fileId = file.getId();
    /** Use thumbnail endpoint; works better than direct uc in restricted domains. */
    return 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1200';
  } catch (e) {
    logError('_saveDevoteePhotoToDrive_', { devoteeId: devoteeId }, e);
    return '__PHOTO_UPLOAD_FAILED__';
  }
}

function _normalizeMobile(v) {
  return String(v || '')
    .replace(/\D/g, '')
    .replace(/^91/, '');
}

function _devoteeAgeYears(dobCell) {
  var d = null;
  if (dobCell instanceof Date && !isNaN(dobCell.getTime())) d = dobCell;
  else {
    var p = _parseDobCell(dobCell);
    if (p instanceof Date && !isNaN(p.getTime())) d = p;
  }
  if (!d) return null;
  var today = new Date();
  var age = today.getFullYear() - d.getFullYear();
  var m = today.getMonth() - d.getMonth();
  if (m < 0 || (m === 0 && today.getDate() < d.getDate())) age--;
  return age;
}

function _appendAuditLog(action, tableAffected, recordId, oldValue, newValue) {
  var sh = getSheet(SHEET_AUDIT);
  var uid = (getCurrentUser() && getCurrentUser().userID) || '';
  var logId = generateId(sh, 'LOG');
  var row = ['', '', '', '', '', '', '', ''];
  row[AUDIT_COL.LogID] = logId;
  row[AUDIT_COL.UserID] = uid;
  row[AUDIT_COL.Action] = action;
  row[AUDIT_COL.TableAffected] = tableAffected;
  row[AUDIT_COL.RecordID] = String(recordId || '');
  row[AUDIT_COL.OldValue] = oldValue === null || oldValue === undefined ? '' : JSON.stringify(oldValue);
  row[AUDIT_COL.NewValue] = newValue === null || newValue === undefined ? '' : JSON.stringify(newValue);
  row[AUDIT_COL.Timestamp] = new Date();
  sh.appendRow(row);
}

/**
 * Run once from Apps Script editor to grant Drive scope.
 * Helps when web-app photo uploads fail due missing authorization.
 */
function authorizeDriveAccess() {
  var appName =
    PropertiesService.getScriptProperties().getProperty('APP_NAME') ||
    'Hari Prabodham Parivar HPP - Akshar Setu';
  var rootName = String(appName).trim() || 'Akshar Setu';
  var folders = DriveApp.getFoldersByName(rootName);
  if (folders.hasNext()) return 'Drive access already granted: ' + folders.next().getName();
  var folder = DriveApp.createFolder(rootName);
  return 'Drive access granted. Folder ready: ' + folder.getName();
}
