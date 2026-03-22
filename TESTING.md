# Testing the daily patient reminder

Use dummy data so all invites go to one inbox and you can sign in as the clinic account.

## 1. Generate dummy data

From the project root:

```bash
pip install -r requirements.txt
python scripts/create_dummy_data.py
```

Use **`python scripts/create_dummy_data.py --full`** only if you want the older **large** matrix (many rows and edge cases).

If the script errors, ensure `pandas` and `openpyxl` are installed (`pip install pandas openpyxl`).

### Default (minimal) outputs

In the project folder:

| File | Purpose |
|------|---------|
| **dummy_action_report.xlsx** | **3 rows**, one scenario each: CREATE, RESCHEDULE (Reschedule Into), RESCHEDULE + Has newer → **Actual** |
| **dummy_actual_report.xlsx** | One **Actual** row (for **APT-M-ACTUAL** only) |
| **dummy_mailchimp.xlsx** | **3 PNs** → test email |
| **dummy_cancel_followup.xlsx** | **Run 2 only**: one **DELETE** for **APT-M-CREATE** (after run 1 stored the invite UID) |
| **dummy_actual_report_cancel_only.xlsx** | Optional **Actual** for run 2 (or reuse **dummy_actual_report.xlsx**) |

All confirmation emails go **to** `purudani.2015@gmail.com`.  
They are sent **from** the mailbox in **`GRAPH_MAILBOX_USER`** (app-only) or the account you sign in with (delegated).

### Invite audit log (Excel)

Every successful send appends **one row** to **`invite_sent_log.xlsx`** (or **`INVITE_LOG_PATH`**) with: time sent (**Eastern**, from `TIMEZONE` in config, default `America/New_York`), from, to, action, subject, PN, appointment id/key, date/time, location, ICS UID/sequence/method, attachment name.  
Set **`INVITE_LOG_PATH=`** empty in `.env` to disable.

### Full matrix (`--full`)

Produces the legacy multi-row workbooks (many PNs). See the script output and use **second-run cancel** notes below; dedupe (“keep last”) can reorder what you expect from a quick skim.

## 2. Point the app at dummy files

Either use environment variables or a `.env` file in the project folder:

```bash
# .env (create from .env.example)
GRAPH_CLIENT_ID=your-azure-app-client-id
ACTION_REPORT_PATH=/full/path/to/daily-patient-reminder/dummy_action_report.xlsx
ACTUAL_REPORT_PATH=/full/path/to/daily-patient-reminder/dummy_actual_report.xlsx
MAILCHIMP_EXPORT_PATH=/full/path/to/daily-patient-reminder/dummy_mailchimp.xlsx
```

If your **production** workbook uses tab names like `3-16 action appt`, set:

```bash
ACTION_SHEET_NAME=3-16 action appt
ACTUAL_SHEET_NAME=3-16 actual appt
```

Use absolute paths so the 8 PM job finds the files.

Optional: use a separate event-id store for testing:

```bash
EVENT_ID_STORE_PATH=/full/path/to/daily-patient-reminder/dummy_event_id_store.json
```

## 3. Run 1 — minimal smoke (3 emails)

1. Sign in as **deepak@libertyptnj.com** when the script prompts (device code or browser), or use app-only auth.
2. Run:

   ```bash
   python run_daily.py
   ```

3. Check **`invite_sent_log.xlsx`**: **3 rows** — `create`, `reschedule`, `reschedule` (the third line is reschedule fed from **Actual**).
4. Check **`reminder_YYYYMMDD.log`** and the inbox for **invite.ics** on each message.

## 4. Run 2 — cancel/delete (1 email)

**Cancel/Delete** needs a **stored ICS UID** for that appointment. After run 1, **`APT-M-CREATE`** is in **`event_id_store.json`**.

Point `.env` at the follow-up Action file (and any Actual file you prefer):

```bash
ACTION_REPORT_PATH=/full/path/to/dummy_cancel_followup.xlsx
ACTUAL_REPORT_PATH=/full/path/to/dummy_actual_report_cancel_only.xlsx
```

Run **`python run_daily.py`** again. Expect **one** email with **cancel.ics** and a new **`invite_sent_log.xlsx`** row with **`action=delete`**.

## 5. Minimal scenario reference

| Appointment ID | PN | What to verify |
|----------------|-----|----------------|
| **APT-M-CREATE** | 100000201 | CREATE → **invite.ics**, log `action=create` |
| **APT-M-RESCHED** | 100000202 | RESCHEDULE from **Reschedule Into** → updated **invite.ics** |
| **APT-M-ACTUAL** | 100000203 | RESCHEDULE + **Has newer** → times from **Actual** sheet |
| **APT-M-CREATE** (run 2) | 100000201 | DELETE → **cancel.ics**, log `action=delete` |

## 6. Quick checklist

- [ ] `python scripts/create_dummy_data.py` — dummy files in project folder.
- [ ] `.env` has `GRAPH_CLIENT_ID` and paths to dummy workbooks.
- [ ] `ACTION_SHEET_NAME` / `ACTUAL_SHEET_NAME` match your tab names (`Action`/`Actual` for dummies).
- [ ] Run 1: **3** **invite.ics** emails; **invite_sent_log.xlsx** has **3** rows.
- [ ] Run 2: **1** **cancel.ics**; log row with **`delete`** for **APT-M-CREATE**.
