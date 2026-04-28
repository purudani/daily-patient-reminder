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

The script creates **only the same three workbook types** you use in production:

| File | Purpose |
|------|---------|
| **dummy_action_report.xlsx** | **3 rows**, one scenario each: CREATE, RESCHEDULE (Reschedule Into), RESCHEDULE + Has newer → **Actual** |
| **dummy_actual_report.xlsx** | **Actual** rows needed for the “has newer” case |
| **dummy_mailchimp.xlsx** | **3 PNs** → test email |

We **do not** generate extra “cancel follow-up” files. Those existed only as a shortcut for a second test run; real data is always **Action + Actual + Mailchimp** only.

By default, all confirmation emails go **to** `ddmittalp@gmail.com` because `DEFAULT_RECIPIENT_EMAIL` is set in `config.py`. Set `DEFAULT_RECIPIENT_EMAIL=` in `.env` to test real Mailchimp recipients.
They are sent **from** the mailbox in **`GRAPH_MAILBOX_USER`** (app-only) or the account you sign in with (delegated).

### Invite audit log (Excel)

Every successful send appends **one row** to **`invite_sent_log.xlsx`** (or **`INVITE_LOG_PATH`**) with: time sent (**Eastern**, from `TIMEZONE` in config, default `America/New_York`), from, to, action, subject, PN, appointment key, date/time, location, ICS UID/sequence/method, attachment name.  
Set **`INVITE_LOG_PATH=`** empty in `.env` to disable.

### Full matrix (`--full`)

Produces the legacy multi-row workbooks (many PNs). Dedupe (“keep last”) can reorder what you expect from a quick skim.

## 2. Point the app at dummy files

Either use environment variables or a `.env` file in the project folder:

```bash
# .env (create from .env.example)
GRAPH_CLIENT_ID=your-azure-app-client-id
```

Place the test files in:

- `Excel/action.xlsx`
- `Excel/actual.xlsx`
- `Excel/processed_mailchimp_export.xlsx`

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
5. After a successful `run_daily.py` run, `Excel/action.xlsx` and `Excel/actual.xlsx` are renamed with the run date, so place fresh copies back under those names before another run.

## 4. Testing cancel/delete (optional second run)

**Why not in the same dummy Action file?**  
If you put **CREATE** and **DELETE** for the **same** visit (same PN + Date + Time) in one Action export, **last row wins** — you’d only process one of them. So cancel is tested with a **separate** Action file (or a new export from the scheduler), same as production.

**After run 1** (so `event_id_store.json` has the UID for the create):

1. Make a **copy** of your Action workbook (or export a new Action sheet) that contains **only** a **DELETE** (or **CANCEL w. remove**) row for the **same** PN, **Date**, and **Time** as the minimal **create** row (PN **100000201**, same date/time as the create line in `dummy_action_report.xlsx`).
2. Replace **`Excel/action.xlsx`** with that file. **`Excel/actual.xlsx`** can stay the same (or any valid Actual file with compatible columns).
3. Run **`python run_daily.py`** again. Expect **one** email with **cancel.ics** and a new **`invite_sent_log.xlsx`** row with **`action=delete`**.

**Cancel row in the Action report** — use **`CANCEL w. remove`** or **`DELETE`** in the **Action** column (same style as your scheduler export). The job maps both to a cancel email + **`cancel.ics`**.

**Why cancel sometimes didn’t send (key mismatch)**  
The app looks up the prior invite in **`event_id_store.json`** using **`PN + Date + Time`**. After a **reschedule**, the store stays keyed by the **original** appointment slot (e.g. **12:00**), not the new time (**10:00**). If your cancel row shows the **new** time only, the keys won’t match. **Fix:** either put the **original** Date/Time on the cancel row, or rely on the **single-visit fallback** (if there is only **one** stored visit for that PN, the app now resolves it automatically — see `resolve_store_key_for_cancel` in `event_id_store.py`). If a patient has **two** upcoming visits, you must match the correct **Date/Time** to the line in **`event_id_store.json`**.

**Where to see “cancel” in data**  
- **Action** sheet: the row whose **Action** column is cancel/delete.  
- **`event_id_store.json`**: that visit’s key disappears after a successful cancel send.  
- **`invite_sent_log.xlsx`**: a row with **`action=cancel`** or **`action=delete`**.

## 5. Minimal scenario reference

| Scenario | PN | What to verify |
|----------|-----|----------------|
| Create row | 100000201 | CREATE → **invite.ics**, log `action=create` |
| Reschedule Into row | 100000202 | RESCHEDULE from **Reschedule Into** → updated **invite.ics** |
| Has newer + Actual | 100000203 | RESCHEDULE + **Has newer** → times from **Actual** sheet |
| Cancel (separate Action file, run 2) | 100000201 | DELETE → **cancel.ics**, log `action=delete` |

## 6. Quick checklist

- [ ] `python scripts/create_dummy_data.py` — **three** dummy files in project folder.
- [ ] `.env` has `GRAPH_CLIENT_ID` and paths to dummy workbooks.
- [ ] `ACTION_SHEET_NAME` / `ACTUAL_SHEET_NAME` match your tab names (`Action`/`Actual` for dummies).
- [ ] Run 1: **3** **invite.ics** emails; **invite_sent_log.xlsx** has **3** rows.
- [ ] Run 2 (optional): **1** **cancel.ics** using a **separate** Action-only file for the delete row.
