/**
 * HPP Seva Connect — configuration and column maps (0-based).
 * All .gs files share global scope in Apps Script.
 */

/** Google Sheet backing the app (set in Script properties: SPREADSHEET_ID). */
const SPREADSHEET_ID =
  PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '';

/** Tab names */
const SHEET_DEVOTEES = 'Devotees';
const SHEET_FAMILIES = 'Families';
const SHEET_MANDALS = 'Mandals';
const SHEET_USERS = 'Users';
const SHEET_SABHAS = 'Sabhas';
const SHEET_ATTENDANCE = 'Attendance';
const SHEET_FOLLOWUPS = 'FollowUps';
const SHEET_EVENTS = 'Events';
const SHEET_REGISTRATIONS = 'EventRegistrations';
const SHEET_SEVA_TYPES = 'SevaTypes';
const SHEET_SEVA = 'SevaAssignments';
const SHEET_NOTIFICATIONS = 'Notifications';
const SHEET_AUDIT = 'AuditLog';
const SHEET_ERROR_LOGS = 'ErrorLogs';
const SHEET_SLOGANS = 'Slogans';

/** Role values (lowercase strings). */
const ROLES = {
  ADMIN: 'admin',
  LEADER: 'leader',
  SUB_LEADER: 'sub-leader',
  KARYAKARTA: 'karyakarta',
  DEVOTEE: 'devotee',
};

const DEVOTEE_STATUS = {
  ACTIVE: 'active',
  INACTIVE: 'inactive',
  RELOCATED: 'relocated',
  DECEASED: 'deceased',
};

const FOLLOWUP_STATUS = {
  PENDING: 'pending',
  DONE: 'done',
  NOT_REACHABLE: 'not_reachable',
  DECLINED: 'declined',
};

const ATTENDANCE_STATUS = {
  PRESENT: 'present',
  ABSENT: 'absent',
  LATE: 'late',
};

const PAYMENT_STATUS = {
  PAID: 'paid',
  PENDING: 'pending',
  WAIVER: 'waiver',
};

const SEVA_STATUS = {
  ASSIGNED: 'assigned',
  CONFIRMED: 'confirmed',
  DONE: 'done',
  ABSENT: 'absent',
};

/**
 * Devotees — column indices (0-based), order matches sheet headers.
 */
const DEVOTEES_COL = {
  DevoteeID: 0,
  FamilyID: 1,
  Mandal: 2,
  MandaID: 2,
  KaryakartaLeaderName: 3,
  SubLeaderID: 3,
  KaryakartaName: 4,
  KaryakartaID: 4,
  FirstName: 5,
  MiddleName: 6,
  LastName: 7,
  Gender: 8,
  DOB: 9,
  Photo: 10,
  Mobile: 11,
  WhatsApp: 12,
  Email: 13,
  Address: 14,
  NativePlace: 15,
  DevoteeType: 16,
  Status: 17,
  BloodGroup: 18,
  IsElderly: 19,
  SpecialNeeds: 20,
  EmergencyContact: 21,
  DikashaLevel: 22,
  DikashaDate: 23,
  Panchamrut: 24,
  Skills: 25,
  WhatsAppOptedIn: 26,
  Reference: 27,
  ReferenceDevoteeID: 27,
  RelationWithHead: 28,
  LanguagePref: 29,
  City: 30,
  State: 31,
  Country: 32,
  DonationEnabled: 33,
  PANNumber: 34,
  CreatedBy: 35,
  CreatedOn: 36,
  UpdatedOn: 37,
};

const FAMILIES_COL = {
  FamilyID: 0,
  HeadDevoteeID: 1,
  MandaID: 2,
  FamilyName: 3,
  TotalMembers: 4,
  CreatedOn: 5,
};

const MANDALS_COL = {
  MandalID: 0,
  MandalName: 1,
  City: 2,
  LeaderUserID: 3,
  CreatedOn: 4,
};

const USERS_COL = {
  UserID: 0,
  Email: 1,
  Name: 2,
  Role: 3,
  MandalID: 4,
  SubGroupID: 5,
  GoogleUID: 6,
  IsActive: 7,
  LastLogin: 8,
  LanguagePref: 9,
  CreatedOn: 10,
};

const SABHAS_COL = {
  SabhaID: 0,
  Title: 1,
  Type: 2,
  Date: 3,
  Time: 4,
  Venue: 5,
  MandalID: 6,
  RecurringType: 7,
  Mode: 8,
  MeetingLink: 9,
  Speaker: 10,
  KathaaTopic: 11,
  Notes: 12,
  MaxCapacity: 13,
  GoogleCalendarEventID: 14,
  CreatedBy: 15,
  CreatedOn: 16,
};

const ATTENDANCE_COL = {
  AttendanceID: 0,
  SabhaID: 1,
  DevoteeID: 2,
  Status: 3,
  MarkedAt: 4,
  MarkedBy: 5,
  CheckInMethod: 6,
  IsStreak: 7,
};

const FOLLOWUPS_COL = {
  FollowUpID: 0,
  DevoteeID: 1,
  AssignedTo: 2,
  Type: 3,
  Status: 4,
  DueDate: 5,
  ContactMode: 6,
  Priority: 7,
  Notes: 8,
  EscalatedTo: 9,
  CreatedBy: 10,
  CreatedOn: 11,
  CompletedOn: 12,
};

const EVENTS_COL = {
  EventID: 0,
  Name: 1,
  Type: 2,
  StartDate: 3,
  EndDate: 4,
  Venue: 5,
  MandalID: 6,
  Capacity: 7,
  WaitlistCount: 8,
  RegistrationDeadline: 9,
  PaymentRequired: 10,
  Amount: 11,
  FeedbackFormURL: 12,
  TargetAudience: 13,
  CreatedBy: 14,
  CreatedOn: 15,
};

const EVENT_REGISTRATIONS_COL = {
  RegID: 0,
  EventID: 1,
  DevoteeID: 2,
  FamilyMembers: 3,
  PaymentStatus: 4,
  AmountPaid: 5,
  TransportNeeded: 6,
  FoodPreference: 7,
  SpecialNeeds: 8,
  RegisteredBy: 9,
  RegisteredOn: 10,
};

const SEVA_TYPES_COL = {
  SevaTypeID: 0,
  Name: 1,
  Description: 2,
  SkillsRequired: 3,
  CreatedOn: 4,
};

const SEVA_ASSIGNMENTS_COL = {
  SevaID: 0,
  SevaTypeID: 1,
  EventID: 2,
  SabhaID: 3,
  Date: 4,
  TimeSlot: 5,
  DevoteeID: 6,
  Status: 7,
  AssignedBy: 8,
  ConfirmedAt: 9,
  Notes: 10,
};

const NOTIFICATIONS_COL = {
  NotifID: 0,
  UserID: 1,
  Type: 2,
  Message: 3,
  IsRead: 4,
  CreatedOn: 5,
};

const AUDIT_COL = {
  LogID: 0,
  UserID: 1,
  Action: 2,
  TableAffected: 3,
  RecordID: 4,
  OldValue: 5,
  NewValue: 6,
  Timestamp: 7,
};
