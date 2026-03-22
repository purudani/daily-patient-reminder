"""
Configuration for the daily patient reminder automation.
Edit paths and options to match your setup.
"""
import os

# --- Paths (local folder where files are placed) ---
BASE_FOLDER = os.path.expanduser(
    os.environ.get("REMINDER_BASE_FOLDER", "~/Desktop/Git_projects/daily-patient-reminder")
)

# Action report: appointments with actions (Create, Reschedule, Cancel, Delete)
ACTION_REPORT_PATH = os.environ.get(
    "ACTION_REPORT_PATH",
    os.path.join(BASE_FOLDER, "action report.xlsx")
)
ACTION_SHEET_NAME = os.environ.get("ACTION_SHEET_NAME", "Action")  # or index 0

# Actual report: current appointment state used when "has newer action"
ACTUAL_REPORT_PATH = os.environ.get(
    "ACTUAL_REPORT_PATH",
    os.path.join(BASE_FOLDER, "actual report.xlsx")
)
ACTUAL_SHEET_NAME = os.environ.get("ACTUAL_SHEET_NAME", "Actual")  # or index 0

# Note: Action and Actual are treated as separate files by default.
# If your office exports both tabs into one workbook, set both *_REPORT_PATH to that file.

# Mailchimp export: PN -> email lookup
MAILCHIMP_EXPORT_PATH = os.environ.get(
    "MAILCHIMP_EXPORT_PATH",
    os.path.join(BASE_FOLDER, "processed_mailchimp_export.xlsx")
)
MAILCHIMP_SHEET_NAME = os.environ.get("MAILCHIMP_SHEET_NAME", 0)  # first sheet

# Persistent store for ICS UID + SEQUENCE (so reschedules/cancels match the same calendar entry)
EVENT_ID_STORE_PATH = os.environ.get(
    "EVENT_ID_STORE_PATH",
    os.path.join(BASE_FOLDER, "event_id_store.json")
)

# One row per outbound patient email (audit trail). Set INVITE_LOG_PATH= to disable.
INVITE_LOG_PATH = os.environ.get(
    "INVITE_LOG_PATH",
    os.path.join(BASE_FOLDER, "invite_sent_log.xlsx"),
)

# --- Excel options ---
SKIP_FIRST_N_ROWS = 2
SKIP_SAME_DAY = True
SKIP_NEXT_DAY = False  # set True to also skip tomorrow's appts
SKIP_BLANK_PN = True
# If multiple actions for same appointment, use latest
MULTIPLE_ACTIONS_USE_LAST = True

# --- Action / Actual report column names (Scheduler Activity Report style) ---
# Typical layout: rows 0-1 title/subtitle, row 2 = header (see SKIP_FIRST_N_ROWS).
COL_ACTION = "Action"  # CREATE, RESCHEDULE, DELETE, CANCEL w. remove, etc.
COL_PN = "PN"
COL_APPT_DATE = "Date"  # appointment date (M/D/YYYY in export)
COL_APPT_TIME = "Time"  # e.g. 04:30p, 11:30a
COL_LOCATION = "Location"  # LIB | LIBN | LIBJ
COL_APPT_TYPE = "Type"  # 30DN, MT50, PTDN, etc. (used for duration)
COL_PATIENT_NAME = "Patient Name"  # "Last, First" — optional display only
COL_APPT_ID = "Appointment ID"  # optional; if missing we key by PN + date + time
# Production uses one text column for reschedule changes:
COL_RESCHEDULE_INTO = "Reschedule Into"
# Legacy / optional separate columns (used if Reschedule Into is empty):
COL_RESCHEDULE_DATE = "Reschedule Date"
COL_RESCHEDULE_TIME = "Reschedule Time"
COL_HAS_NEWER_ACTION = "Has newer"  # Yes or blank (also matches "Has Newer Action" via case-insensitive lookup)

# If the Actual sheet uses different headers for date/time, set these (empty = same as action columns):
ACTUAL_COL_DATE = os.environ.get("ACTUAL_COL_DATE", "").strip() or None  # e.g. "Last Schedul"
ACTUAL_COL_TIME = os.environ.get("ACTUAL_COL_TIME", "").strip() or None
ACTUAL_COL_LOCATION = os.environ.get("ACTUAL_COL_LOCATION", "").strip() or None
ACTUAL_COL_TYPE = os.environ.get("ACTUAL_COL_TYPE", "").strip() or None

# --- Mailchimp / audience export (real export uses separate First/Last) ---
COL_MAILCHIMP_PN = "PN"
COL_MAILCHIMP_EMAIL = "Email"
COL_MAILCHIMP_FIRST = "First"
COL_MAILCHIMP_LAST = "Last"
COL_MAILCHIMP_NAME = "Name"  # if present, used instead of First+Last

# Duration policy: 60 minutes for all appointments except MT30 = 30 minutes.
APPOINTMENT_TYPE_DURATION_MINUTES = {
    "MT30": 30,
}

# --- Locations (code -> address) ---
LOCATION_MAP = {
    "LIB": "115 Columbus Dr, Ste 300, Jersey City, NJ 07302",
    "LIBN": "132 Newark Ave, Jersey City, NJ 07302",
    "LIBJ": "2 Journal Sq Plaza, Jersey City, NJ 07306",
}

# --- Calendar / ICS defaults (patient email + invite.ics) ---
DEFAULT_DURATION_MINUTES = 60
MT30_DURATION_MINUTES = 30  # legacy; prefer APPOINTMENT_TYPE_DURATION_MINUTES
ORGANIZER_NAME = "Liberty PT & Wellness"
CALENDAR_TITLE = "Liberty PT & Wellness"
WEBSITE_URL = "https://libertyptnj.com/"
REMINDER_MINUTES_BEFORE = [48 * 60, 2 * 60]  # first value used for ICS VALARM
# Graph accepts Windows names (e.g. "Eastern Standard Time") or IANA IDs.
# IANA is often clearer for Google Calendar and other non-Outlook clients parsing the invite.
TIMEZONE = os.environ.get("TIMEZONE", "America/New_York").strip() or "America/New_York"

# --- Email (confirmation body + meeting invite HTML) ---
CONFIRMATION_PHONE = "201-366-1115"
# Optional: full URL to logo image for HTML body (if empty, a styled text header is used)
EMAIL_LOGO_URL = os.environ.get("EMAIL_LOGO_URL", "").strip()
# Inbox preview line (some clients show first line / hidden preheader)
EMAIL_PREVIEW_TEXT = os.environ.get(
    "EMAIL_PREVIEW_TEXT",
    "Your appointment at Liberty PT & Wellness is confirmed.",
)

# Outlook on the web “Add to calendar” link base (M365 default). Use https://outlook.live.com for Outlook.com.
OUTLOOK_CALENDAR_WEB_BASE = os.environ.get(
    "OUTLOOK_CALENDAR_WEB_BASE", "https://outlook.office.com"
).strip() or "https://outlook.office.com"

# --- Microsoft Graph ---
# Put real values in .env only — never commit secrets to the repo.
#
# App-only (no sign-in): set GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, GRAPH_TENANT_ID,
# and GRAPH_MAILBOX_USER (organizer mailbox UPN, e.g. deepak@libertyptnj.com).
# Azure app needs Application permission: Mail.Send (+ admin consent). Patient invites are email + .ics.
#
# Delegated (device code / interactive): leave GRAPH_CLIENT_SECRET empty; set GRAPH_CLIENT_ID
# and optionally GRAPH_TENANT_ID. Sign in when prompted. Optional GRAPH_MAILBOX_USER: if set,
# API uses /users/{email}/... (must match the signed-in account).
GRAPH_TENANT_ID = os.environ.get("GRAPH_TENANT_ID", "")
GRAPH_CLIENT_ID = os.environ.get("GRAPH_CLIENT_ID", "")
GRAPH_CLIENT_SECRET = os.environ.get("GRAPH_CLIENT_SECRET", "")
# Organizer / shared mailbox that owns the calendar and sends invites (required for app-only)
GRAPH_MAILBOX_USER = os.environ.get("GRAPH_MAILBOX_USER", "")

# User.Read: only needed for delegated mode when GRAPH_MAILBOX_USER is unset (GET /me for organizer email).
GRAPH_SCOPES = [
    "https://graph.microsoft.com/Mail.Send",
    "https://graph.microsoft.com/User.Read",
]
GRAPH_APP_ONLY_SCOPE = ["https://graph.microsoft.com/.default"]

# --- Logging ---
LOG_FOLDER = os.environ.get("REMINDER_LOG_FOLDER", BASE_FOLDER)
LOG_INVITES_AND_CHANGES = True  # log every invite/change sent
