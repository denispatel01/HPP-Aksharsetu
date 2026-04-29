/**
 * HPP Seva Connect — session and role checks (UserProperties).
 * Session shape: { userID, email, name, role, mandalID }
 */

var AUTH_SESSION_KEY = 'SESSION_USER';

/**
 * Sign in: find Users row by email or append a new user, then persist session.
 * @param {string} googleEmail Optional; falls back to Session.getActiveUser() when deployed as user.
 * @param {string} googleName Optional display name.
 * @return {{ success: boolean, data?: Object, error?: string }}
 */
function loginUser(googleEmail, googleName) {
  try {
    var email = _authNormalizeEmail(googleEmail);
    var nameHint = googleName ? String(googleName).trim() : '';

    if (!email) {
      try {
        email = _authNormalizeEmail(Session.getActiveUser().getEmail());
      } catch (activeErr) {
        logError('loginUser', { step: 'resolveEmail' }, activeErr);
      }
    }
    if (!email) {
      return { success: false, error: 'Could not determine Google email. Allow access and try again.' };
    }

    try {
      var activeEmail = _authNormalizeEmail(Session.getActiveUser().getEmail());
      if (activeEmail && activeEmail !== email) {
        return { success: false, error: 'Email does not match the signed-in Google account.' };
      }
    } catch (verifyErr) {
      // Execute-as-owner / no active user: skip strict check
    }

    var sheet = getSheet(SHEET_USERS);
    var lastRow = sheet.getLastRow();
    var emailCol = USERS_COL.Email + 1;
    var foundRow = 0;

    if (lastRow >= 2) {
      var emails = sheet.getRange(2, emailCol, lastRow, emailCol).getValues();
      for (var i = 0; i < emails.length; i++) {
        if (_authNormalizeEmail(emails[i][0]) === email) {
          foundRow = i + 2;
          break;
        }
      }
    }

    var now = new Date();
    var numCols = 11;
    var userId;
    var displayName;
    var role;
    var mandalID;

    if (foundRow) {
      var range = sheet.getRange(foundRow, 1, 1, numCols);
      var row = range.getValues()[0];
      userId = String(row[USERS_COL.UserID] || '').trim();
      displayName = String(row[USERS_COL.Name] || '').trim() || nameHint || email;
      role = String(row[USERS_COL.Role] || ROLES.DEVOTEE).trim() || ROLES.DEVOTEE;
      mandalID = row[USERS_COL.MandalID] != null && row[USERS_COL.MandalID] !== ''
        ? String(row[USERS_COL.MandalID])
        : '';

      if (nameHint) {
        range.getCell(1, USERS_COL.Name + 1).setValue(nameHint);
        displayName = nameHint;
      }
      range.getCell(1, USERS_COL.LastLogin + 1).setValue(now);
    } else {
      userId = generateId(sheet, 'USR');
      displayName = nameHint || email.split('@')[0];
      role = ROLES.DEVOTEE;
      mandalID = '';

      var newRow = [];
      for (var c = 0; c < numCols; c++) newRow[c] = '';
      newRow[USERS_COL.UserID] = userId;
      newRow[USERS_COL.Email] = email;
      newRow[USERS_COL.Name] = displayName;
      newRow[USERS_COL.Role] = role;
      newRow[USERS_COL.MandalID] = mandalID;
      newRow[USERS_COL.SubGroupID] = '';
      newRow[USERS_COL.GoogleUID] = '';
      newRow[USERS_COL.IsActive] = true;
      newRow[USERS_COL.LastLogin] = now;
      newRow[USERS_COL.LanguagePref] = '';
      newRow[USERS_COL.CreatedOn] = now;

      sheet.appendRow(newRow);
    }

    var userObject = {
      userID: userId,
      email: email,
      name: displayName,
      role: role,
      mandalID: mandalID,
    };

    PropertiesService.getUserProperties().setProperty(
      AUTH_SESSION_KEY,
      JSON.stringify(userObject)
    );

    return { success: true, data: userObject };
  } catch (err) {
    logError('loginUser', { email: googleEmail }, err);
    return { success: false, error: err.message || String(err) };
  }
}

/**
 * @return {Object|null} Parsed session user or null.
 */
function getCurrentUser() {
  var raw = PropertiesService.getUserProperties().getProperty(AUTH_SESSION_KEY);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    logError('getCurrentUser', { rawLength: raw ? raw.length : 0 }, e);
    return null;
  }
}

/**
 * Clears all user-scoped script properties (session).
 * @return {{ success: boolean }}
 */
function logoutUser() {
  PropertiesService.getUserProperties().deleteAllProperties();
  return { success: true };
}

/**
 * @param {string[]} allowedRoles
 */
function checkRole(allowedRoles) {
  var user = getCurrentUser();
  var msg = 'Access denied. Required roles: ' + allowedRoles.join(', ');
  if (!user || !user.role) {
    throw new Error(msg);
  }
  var role = String(user.role);
  for (var i = 0; i < allowedRoles.length; i++) {
    if (allowedRoles[i] === role) return;
  }
  throw new Error(msg);
}

/** @param {*} v @return {string} */
function _authNormalizeEmail(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim().toLowerCase();
}
