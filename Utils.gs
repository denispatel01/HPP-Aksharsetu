/**
 * HPP Seva Connect — spreadsheet helpers, IDs, responses, session checks, sheet bootstrap.
 */

/** @return {GoogleAppsScript.Spreadsheet.Spreadsheet} */
function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error('No spreadsheet: set Script property SPREADSHEET_ID or open the bound spreadsheet.');
  }
  return active;
}

/**
 * @param {string} name Sheet tab name (use SHEET_* from Config).
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getSheet(name) {
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    throw new Error('Sheet not found: ' + name);
  }
  return sh;
}

/**
 * Next ID in column A: prefix + zero-padded number (e.g. DEV001). Scans existing IDs.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} prefix e.g. 'DEV'
 * @return {string}
 */
function generateId(sheet, prefix) {
  var p = String(prefix || '');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return p + _padIdNumber(1, 3);
  }
  var values = sheet.getRange(2, 1, lastRow, 1).getValues();
  var maxNum = 0;
  var re = new RegExp('^' + _escapeRegExp(p) + '(\\d+)$', 'i');
  for (var i = 0; i < values.length; i++) {
    var raw = values[i][0];
    if (raw === null || raw === '') continue;
    var str = String(raw).trim();
    var m = str.match(re);
    if (m) {
      var n = parseInt(m[1], 10);
      if (!isNaN(n) && n > maxNum) maxNum = n;
    }
  }
  var next = maxNum + 1;
  var width = Math.max(3, String(next).length);
  return p + _padIdNumber(next, width);
}

/**
 * @param {Date} date
 * @return {string} yyyy-MM-dd
 */
function formatDate(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return '';
  }
  return Utilities.formatDate(date, _appTimeZone(), 'yyyy-MM-dd');
}

/**
 * @param {Date} date
 * @return {string} yyyy-MM-dd HH:mm
 */
function formatDateTime(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return '';
  }
  return Utilities.formatDate(date, _appTimeZone(), 'yyyy-MM-dd HH:mm');
}

/**
 * @return {string} Today's date yyyy-MM-dd in app timezone.
 */
function todayStr() {
  return formatDate(new Date());
}

/**
 * Structured error logging for debugging and support.
 * @param {string} fn Function name
 * @param {*} context Serializable context (object or string)
 * @param {*} err Error or message
 */
function logError(fn, context, err) {
  var msg = err && err.message ? err.message : String(err);
  var stack = err && err.stack ? err.stack : '';
  var ctx =
    context === null || context === undefined
      ? ''
      : typeof context === 'object'
        ? JSON.stringify(context)
        : String(context);
  Logger.log('[HPP ERROR] fn=%s | context=%s | message=%s | stack=%s', fn, ctx, msg, stack);
  return writeErrorLog_('server', fn, context, err, '');
}

/**
 * @param {*} data
 * @return {{ success: boolean, data: * }}
 */
function successResponse(data) {
  return { success: true, data: data };
}

/**
 * @param {string} message
 * @return {{ success: boolean, error: string }}
 */
function errorResponse(message, errorId) {
  var out = { success: false, error: String(message || 'Unknown error') };
  if (errorId) out.errorId = String(errorId);
  return out;
}

/**
 * One-time (or repeat) setup: create all 13 tabs, header row, header style, freeze row 1.
 * Run manually from the Apps Script editor after opening/binding the spreadsheet.
 */
function setupSheets() {
  var ss = getSpreadsheet();
  var defs = _sheetSetupDefinitions();
  var navy = '#162240';
  var white = '#FFFFFF';

  for (var i = 0; i < defs.length; i++) {
    var def = defs[i];
    var sh = ss.getSheetByName(def.name);
    if (!sh) {
      sh = ss.insertSheet(def.name);
    }
    var headers = def.headers;
    var numCols = headers.length;
    sh.getRange(1, 1, 1, numCols).setValues([headers]);
    var headerRange = sh.getRange(1, 1, 1, numCols);
    headerRange.setBackground(navy);
    headerRange.setFontColor(white);
    headerRange.setFontWeight('bold');
    sh.setFrozenRows(1);
  }
}

/**
 * HPP Seva Connect — one-click database bootstrap (run from Apps Script editor ▶).
 * Creates all tabs missing from the spreadsheet (Users, Devotees, …) and writes
 * row 1 headers + navy header styling + freeze row 1. Safe to run again (refreshes headers).
 *
 * Prerequisites:
 * • Web app / standalone: File → Project settings → Script properties → add
 *   SPREADSHEET_ID = the Google Sheet ID from the sheet URL, OR
 * • Open the target spreadsheet → Extensions → Apps Script (container-bound project).
 */
function hppSetupAllSheets() {
  try {
    setupSheets();
    var ss = getSpreadsheet();
    Logger.log(
      'hppSetupAllSheets: OK — spreadsheet "' +
        ss.getName() +
        '" has ' +
        ss.getSheets().length +
        ' tab(s).'
    );
  } catch (e) {
    var msg =
      'hppSetupAllSheets failed: ' +
      (e && e.message ? e.message : String(e)) +
      ' — Set Script property SPREADSHEET_ID to your Sheet ID, or bind this script to the spreadsheet.';
    Logger.log(msg);
    throw new Error(msg);
  }
}

/** @return {string} */
function _appTimeZone() {
  try {
    return Session.getScriptTimeZone() || 'Asia/Kolkata';
  } catch (e) {
    return 'Asia/Kolkata';
  }
}

/** @param {number} n @param {number} width @return {string} */
function _padIdNumber(n, width) {
  var s = String(n);
  while (s.length < width) s = '0' + s;
  return s;
}

/** @param {string} s @return {string} */
function _escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Tab names and header rows (must match Config column maps order).
 * @return {{ name: string, headers: string[] }[]}
 */
function _sheetSetupDefinitions() {
  return [
    {
      name: SHEET_DEVOTEES,
      headers: [
        'DevoteeID',
        'FamilyID',
        'Mandal',
        'KaryakartaLeaderName',
        'KaryakartaName',
        'FirstName',
        'MiddleName',
        'LastName',
        'Gender',
        'DOB',
        'Photo',
        'Mobile',
        'WhatsApp',
        'Email',
        'Address',
        'NativePlace',
        'DevoteeType',
        'Status',
        'BloodGroup',
        'IsElderly',
        'SpecialNeeds',
        'EmergencyContact',
        'DikashaLevel',
        'DikashaDate',
        'Panchamrut',
        'Skills',
        'WhatsAppOptedIn',
        'Reference',
        'RelationWithHead',
        'LanguagePref',
        'City',
        'State',
        'Country',
        'DonationEnabled',
        'PANNumber',
        'CreatedBy',
        'CreatedOn',
        'UpdatedOn',
      ],
    },
    {
      name: SHEET_FAMILIES,
      headers: [
        'FamilyID',
        'HeadDevoteeID',
        'MandaID',
        'FamilyName',
        'TotalMembers',
        'CreatedOn',
      ],
    },
    {
      name: SHEET_MANDALS,
      headers: ['MandalID', 'MandalName', 'City', 'LeaderUserID', 'CreatedOn'],
    },
    {
      name: SHEET_USERS,
      headers: [
        'UserID',
        'Email',
        'Name',
        'Role',
        'MandalID',
        'SubGroupID',
        'GoogleUID',
        'IsActive',
        'LastLogin',
        'LanguagePref',
        'CreatedOn',
      ],
    },
    {
      name: SHEET_SABHAS,
      headers: [
        'SabhaID',
        'Title',
        'Type',
        'Date',
        'Time',
        'Venue',
        'MandalID',
        'RecurringType',
        'Mode',
        'MeetingLink',
        'Speaker',
        'KathaaTopic',
        'Notes',
        'MaxCapacity',
        'GoogleCalendarEventID',
        'CreatedBy',
        'CreatedOn',
      ],
    },
    {
      name: SHEET_ATTENDANCE,
      headers: [
        'AttendanceID',
        'SabhaID',
        'DevoteeID',
        'Status',
        'MarkedAt',
        'MarkedBy',
        'CheckInMethod',
        'IsStreak',
      ],
    },
    {
      name: SHEET_FOLLOWUPS,
      headers: [
        'FollowUpID',
        'DevoteeID',
        'AssignedTo',
        'Type',
        'Status',
        'DueDate',
        'ContactMode',
        'Priority',
        'Notes',
        'EscalatedTo',
        'CreatedBy',
        'CreatedOn',
        'CompletedOn',
      ],
    },
    {
      name: SHEET_EVENTS,
      headers: [
        'EventID',
        'Name',
        'Type',
        'StartDate',
        'EndDate',
        'Venue',
        'MandalID',
        'Capacity',
        'WaitlistCount',
        'RegistrationDeadline',
        'PaymentRequired',
        'Amount',
        'FeedbackFormURL',
        'TargetAudience',
        'CreatedBy',
        'CreatedOn',
      ],
    },
    {
      name: SHEET_REGISTRATIONS,
      headers: [
        'RegID',
        'EventID',
        'DevoteeID',
        'FamilyMembers',
        'PaymentStatus',
        'AmountPaid',
        'TransportNeeded',
        'FoodPreference',
        'SpecialNeeds',
        'RegisteredBy',
        'RegisteredOn',
      ],
    },
    {
      name: SHEET_SEVA_TYPES,
      headers: ['SevaTypeID', 'Name', 'Description', 'SkillsRequired', 'CreatedOn'],
    },
    {
      name: SHEET_SEVA,
      headers: [
        'SevaID',
        'SevaTypeID',
        'EventID',
        'SabhaID',
        'Date',
        'TimeSlot',
        'DevoteeID',
        'Status',
        'AssignedBy',
        'ConfirmedAt',
        'Notes',
      ],
    },
    {
      name: SHEET_NOTIFICATIONS,
      headers: ['NotifID', 'UserID', 'Type', 'Message', 'IsRead', 'CreatedOn'],
    },
    {
      name: SHEET_AUDIT,
      headers: [
        'LogID',
        'UserID',
        'Action',
        'TableAffected',
        'RecordID',
        'OldValue',
        'NewValue',
        'Timestamp',
      ],
    },
    {
      name: SHEET_ERROR_LOGS,
      headers: ['LogID', 'Timestamp', 'Source', 'FunctionName', 'UserEmail', 'Message', 'Stack', 'Context', 'Page'],
    },
    {
      name: SHEET_SLOGANS,
      headers: ['Slogan', 'Author', 'IsActive'],
    },
  ];
}

/**
 * HPP Seva Connect — insert DEMO / sample rows into every tab (run from editor ▶).
 *
 * WARNING: Deletes ALL data rows (row 2 onward) on every HPP sheet, then inserts
 * fresh sample data. Do NOT run on a production sheet with real data.
 *
 * Run AFTER hppSetupAllSheets() at least once (this function calls setupSheets()
 * first so headers exist).
 */
function hppInsertSampleData() {
  setupSheets();
  var now = new Date();
  var dSabha = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 3);
  var dEventStart = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 14);
  var dEventEnd = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 15);
  var dDeadline = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 10);
  var dFollowDue = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 7);

  var sheetOrder = [
    SHEET_DEVOTEES,
    SHEET_FAMILIES,
    SHEET_MANDALS,
    SHEET_USERS,
    SHEET_SABHAS,
    SHEET_ATTENDANCE,
    SHEET_FOLLOWUPS,
    SHEET_EVENTS,
    SHEET_REGISTRATIONS,
    SHEET_SEVA_TYPES,
    SHEET_SEVA,
    SHEET_NOTIFICATIONS,
    SHEET_AUDIT,
    SHEET_ERROR_LOGS,
    SHEET_SLOGANS,
  ];

  for (var s = 0; s < sheetOrder.length; s++) {
    _hppClearDataBelowHeader(getSheet(sheetOrder[s]));
  }

  /** @param {GoogleAppsScript.Spreadsheet.Sheet} sh */
  function appendRows(sh, rows) {
    if (!rows || !rows.length) return;
    var start = sh.getLastRow() + 1;
    // Sheet.getRange(row, column, numRows, numColumns) — 3rd arg is COUNT of rows, not end row.
    sh.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
  }

  appendRows(getSheet(SHEET_MANDALS), [
    ['M001', 'Adhyatmik Mandal North', 'Surat', 'USR002', now],
    ['M002', 'Adhyatmik Mandal South', 'Surat', 'USR002', now],
  ]);

  appendRows(getSheet(SHEET_USERS), [
    [
      'USR001',
      'admin.sample@hpp.local',
      'Demo Admin',
      ROLES.ADMIN,
      'M001',
      '',
      '',
      true,
      now,
      'gu',
      now,
    ],
    [
      'USR002',
      'leader.sample@hpp.local',
      'Demo Leader',
      ROLES.LEADER,
      'M001',
      '',
      '',
      true,
      now,
      'gu',
      now,
    ],
    [
      'USR003',
      'subleader.sample@hpp.local',
      'Demo Sub-Leader',
      ROLES.SUB_LEADER,
      'M001',
      'SG1',
      '',
      true,
      now,
      'gu',
      now,
    ],
    [
      'USR004',
      'karyakarta.sample@hpp.local',
      'Demo Karyakarta',
      ROLES.KARYAKARTA,
      'M001',
      '',
      '',
      true,
      now,
      'gu',
      now,
    ],
    [
      'USR005',
      'devotee.sample@hpp.local',
      'Demo Bhakt',
      ROLES.DEVOTEE,
      'M001',
      '',
      '',
      true,
      now,
      'gu',
      now,
    ],
  ]);

  function devRow(o) {
    var r = [];
    for (var i = 0; i < 38; i++) r[i] = '';
    r[DEVOTEES_COL.DevoteeID] = o.id;
    r[DEVOTEES_COL.FamilyID] = o.fam;
    r[DEVOTEES_COL.MandaID] = o.mandal || 'M001';
    r[DEVOTEES_COL.SubLeaderID] = o.sub || 'USR003';
    r[DEVOTEES_COL.KaryakartaID] = o.kar || 'Demo Karyakarta';
    r[DEVOTEES_COL.FirstName] = o.fn;
    r[DEVOTEES_COL.MiddleName] = o.mn || '';
    r[DEVOTEES_COL.LastName] = o.ln;
    r[DEVOTEES_COL.Gender] = o.gender;
    r[DEVOTEES_COL.DOB] = o.dob;
    r[DEVOTEES_COL.Photo] = '';
    r[DEVOTEES_COL.Mobile] = o.mob;
    r[DEVOTEES_COL.WhatsApp] = o.mob;
    r[DEVOTEES_COL.Email] = o.email || '';
    r[DEVOTEES_COL.Address] = o.addr || 'Surat, Gujarat';
    r[DEVOTEES_COL.NativePlace] = 'Surat';
    r[DEVOTEES_COL.DevoteeType] = o.dtype || 'yuvak';
    r[DEVOTEES_COL.Status] = DEVOTEE_STATUS.ACTIVE;
    r[DEVOTEES_COL.BloodGroup] = o.blood || 'B+';
    r[DEVOTEES_COL.IsElderly] = !!o.elderly;
    r[DEVOTEES_COL.SpecialNeeds] = '';
    r[DEVOTEES_COL.EmergencyContact] = o.em || '9876500099';
    r[DEVOTEES_COL.DikashaLevel] = o.diksha || 'naam';
    r[DEVOTEES_COL.DikashaDate] = '';
    r[DEVOTEES_COL.Panchamrut] = o.panch || 'yes';
    r[DEVOTEES_COL.Skills] = 'Education: B.Com | Occupation: Service';
    r[DEVOTEES_COL.WhatsAppOptedIn] = true;
    r[DEVOTEES_COL.ReferenceDevoteeID] = o.ref || '';
    r[DEVOTEES_COL.RelationWithHead] = o.rel || 'Self';
    r[DEVOTEES_COL.LanguagePref] = 'gu';
    r[DEVOTEES_COL.City] = 'Surat';
    r[DEVOTEES_COL.State] = 'Gujarat';
    r[DEVOTEES_COL.Country] = 'India';
    r[DEVOTEES_COL.DonationEnabled] = false;
    r[DEVOTEES_COL.PANNumber] = '';
    r[DEVOTEES_COL.CreatedBy] = 'USR001';
    r[DEVOTEES_COL.CreatedOn] = now;
    r[DEVOTEES_COL.UpdatedOn] = now;
    return r;
  }

  var dob1 = new Date(1992, 2, 15);
  var dob2 = new Date(1995, 7, 8);
  var dob3 = new Date(2012, 10, 1);
  var dob4 = new Date(1988, 0, 22);
  var dob5 = new Date(1990, 5, 30);

  appendRows(
    getSheet(SHEET_DEVOTEES),
    [
      devRow({
        id: 'DEV001',
        fam: 'FAM001',
        fn: 'Hitesh',
        ln: 'Shah',
        gender: 'male',
        dob: dob1,
        mob: '9876500001',
        email: 'hitesh.sample@hpp.local',
        dtype: 'yuvak,ambrish',
        rel: 'Self',
      }),
      devRow({
        id: 'DEV002',
        fam: 'FAM001',
        fn: 'Kiran',
        ln: 'Shah',
        gender: 'female',
        dob: dob2,
        mob: '9876500002',
        email: 'kiran.sample@hpp.local',
        dtype: 'yuvti',
        rel: 'Spouse',
        ref: 'DEV001',
      }),
      devRow({
        id: 'DEV003',
        fam: 'FAM001',
        fn: 'Arjun',
        ln: 'Shah',
        gender: 'male',
        dob: dob3,
        mob: '9876500003',
        dtype: 'kishore',
        rel: 'Child',
        ref: 'DEV001',
      }),
      devRow({
        id: 'DEV004',
        fam: 'FAM002',
        fn: 'Pooja',
        ln: 'Mehta',
        gender: 'female',
        dob: dob4,
        mob: '9876500004',
        email: 'pooja.sample@hpp.local',
        dtype: 'haribhakt',
        rel: 'Self',
      }),
      devRow({
        id: 'DEV005',
        fam: 'FAM002',
        fn: 'Rohan',
        ln: 'Mehta',
        gender: 'male',
        dob: dob5,
        mob: '9876500005',
        dtype: 'yuvak',
        rel: 'Spouse',
        ref: 'DEV004',
      }),
    ]
  );

  appendRows(getSheet(SHEET_FAMILIES), [
    ['FAM001', 'DEV001', 'M001', 'Shah Family', 3, now],
    ['FAM002', 'DEV004', 'M001', 'Mehta Family', 2, now],
  ]);

  appendRows(getSheet(SHEET_SABHAS), [
    [
      'SAB001',
      'Friday Satsang Sabha',
      'weekly',
      dSabha,
      '19:00',
      'Mandir Hall, Varachha',
      'M001',
      'weekly',
      'offline',
      '',
      'P.Pu. Swami',
      'Satsang Katha',
      'Sample sabha row',
      200,
      '',
      'USR002',
      now,
    ],
    [
      'SAB002',
      'Yuvak Sabha',
      'special',
      dSabha,
      '17:30',
      'Community Hall',
      'M001',
      'none',
      'hybrid',
      'https://meet.sample.invalid/hpp',
      'Karyakarta',
      'Niyam & Maryada',
      '',
      80,
      '',
      'USR004',
      now,
    ],
  ]);

  appendRows(getSheet(SHEET_ATTENDANCE), [
    ['ATT001', 'SAB001', 'DEV001', ATTENDANCE_STATUS.PRESENT, now, 'USR004', 'app', true],
    ['ATT002', 'SAB001', 'DEV002', ATTENDANCE_STATUS.PRESENT, now, 'USR004', 'app', true],
    ['ATT003', 'SAB001', 'DEV003', ATTENDANCE_STATUS.LATE, now, 'USR004', 'manual', false],
  ]);

  appendRows(getSheet(SHEET_FOLLOWUPS), [
    [
      'FU001',
      'DEV005',
      'USR004',
      'visit',
      FOLLOWUP_STATUS.PENDING,
      dFollowDue,
      'call',
      'high',
      'Welcome visit after relocation',
      '',
      'USR002',
      now,
      '',
    ],
    [
      'FU002',
      'DEV003',
      'USR004',
      'class',
      FOLLOWUP_STATUS.DONE,
      dFollowDue,
      'whatsapp',
      'normal',
      'Bal Sabha follow-up',
      '',
      'USR004',
      now,
      now,
    ],
  ]);

  appendRows(getSheet(SHEET_EVENTS), [
    [
      'EVT001',
      'Yuva Shibir 2026',
      'shibir',
      dEventStart,
      dEventEnd,
      'Ashram Campus',
      'M001',
      120,
      5,
      dDeadline,
      false,
      0,
      '',
      'yuvak,yuvti',
      'USR002',
      now,
    ],
  ]);

  appendRows(getSheet(SHEET_REGISTRATIONS), [
    [
      'REG001',
      'EVT001',
      'DEV001',
      '2',
      PAYMENT_STATUS.WAIVER,
      0,
      false,
      'regular',
      '',
      'USR004',
      now,
    ],
  ]);

  appendRows(getSheet(SHEET_SEVA_TYPES), [
    ['ST001', 'Parking Seva', 'Vehicle guidance', 'communication', now],
    ['ST002', 'Prasad Seva', 'Kitchen help', 'hygiene', now],
  ]);

  appendRows(getSheet(SHEET_SEVA), [
    [
      'SEV001',
      'ST001',
      '',
      'SAB001',
      dSabha,
      '18:00-19:00',
      'DEV002',
      SEVA_STATUS.CONFIRMED,
      'USR004',
      now,
      'Gate A',
    ],
  ]);

  appendRows(getSheet(SHEET_NOTIFICATIONS), [
    ['NTF001', 'USR005', 'info', 'Welcome to Hari Prabodham Parivar HPP - Akshar Setu (sample).', false, now],
  ]);

  appendRows(getSheet(SHEET_AUDIT), [
    [
      'LOG001',
      'USR001',
      'SAMPLE_SEED',
      'ALL',
      'hppInsertSampleData',
      '',
      '{"note":"demo data"}',
      now,
    ],
  ]);

  Logger.log('hppInsertSampleData: sample rows inserted on all sheets.');
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function _hppClearDataBelowHeader(sheet) {
  var lr = sheet.getLastRow();
  if (lr > 1) {
    sheet.deleteRows(2, lr - 1);
  }
}
