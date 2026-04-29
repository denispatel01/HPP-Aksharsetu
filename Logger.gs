/**
 * Centralized error logger (server + client) with persistent sheet records.
 */

/**
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateErrorLogSheet_() {
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName(SHEET_ERROR_LOGS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_ERROR_LOGS);
  }
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, 9).setValues([
      ['LogID', 'Timestamp', 'Source', 'FunctionName', 'UserEmail', 'Message', 'Stack', 'Context', 'Page'],
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/**
 * Run once manually (or can be called safely multiple times).
 * Ensures ErrorLogs tab exists with headers.
 * @return {{success:boolean,data:{sheet:string}}}
 */
function hppEnsureErrorLogSheet() {
  var sh = getOrCreateErrorLogSheet_();
  return successResponse({ sheet: sh.getName() });
}

/**
 * @param {string} source "server" | "client"
 * @param {string} fn
 * @param {*} context
 * @param {*} err
 * @param {string=} page
 * @return {string} log id
 */
function writeErrorLog_(source, fn, context, err, page) {
  try {
    var sh = getOrCreateErrorLogSheet_();
    var ts = new Date();
    var logId = 'ERR-' + Utilities.formatDate(ts, _appTimeZone(), 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 1000);
    var msg = err && err.message ? err.message : String(err || 'Unknown error');
    var stack = err && err.stack ? String(err.stack) : '';
    var ctx =
      context === null || context === undefined
        ? ''
        : typeof context === 'object'
          ? JSON.stringify(context)
          : String(context);
    var user = getCurrentUser();
    var email = (user && user.email) || '';
    sh.appendRow([logId, ts, String(source || 'server'), String(fn || ''), email, msg, stack, ctx, String(page || '')]);
    return logId;
  } catch (e) {
    Logger.log('[HPP ERROR LOGGER FAILED] %s', e && e.message ? e.message : String(e));
    return 'ERR-LOGGER-FAILED';
  }
}

/**
 * Called from frontend to log browser/runtime errors.
 * @param {{message?:string,stack?:string,functionName?:string,context?:*,page?:string}} payload
 * @return {{success:boolean,data:{logId:string}}}
 */
function logClientError(payload) {
  var p = payload || {};
  var errObj = { message: String(p.message || 'Client error'), stack: String(p.stack || '') };
  var id = writeErrorLog_('client', p.functionName || 'client', p.context || {}, errObj, p.page || '');
  return successResponse({ logId: id });
}

/**
 * @param {number=} limit
 * @return {{success:boolean,data:Array}}
 */
function getRecentErrorLogs(limit) {
  try {
    var sh = getOrCreateErrorLogSheet_();
    var last = sh.getLastRow();
    if (last < 2) return successResponse([]);
    var lim = parseInt(limit, 10);
    if (!lim || lim < 1) lim = 100;
    if (lim > 500) lim = 500;
    var start = Math.max(2, last - lim + 1);
    var rows = sh.getRange(start, 1, last - start + 1, 9).getValues();
    return successResponse(rows);
  } catch (e) {
    var id = writeErrorLog_('server', 'getRecentErrorLogs', { limit: limit }, e, '');
    return errorResponse('Could not read error logs. Ref: ' + id, id);
  }
}
