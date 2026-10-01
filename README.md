# Plan IO 2026Z → Google Calendar (Semestr 1, 2 or 3)

## Files
- `main.py` – CLI (dry run by default)
- `parse_plan.py` – reads the Excel file, maps codes to full names, validates dates, assigns colours
- `gcal.py` – Google Calendar auth + create/update/delete
- `config.toml` – all settings (calendar name, excluded electives, name/date/colour overrides)

Needs Python 3.11+.

## Setup
1. `pip install -r requirements.txt`
2. Put `plan_zajęć_IO_2026Z.xlsx` next to `main.py` (or pass `--excel path`).
3. Google credentials (one-time):
   - https://console.cloud.google.com → create a project → *APIs & Services* → enable **Google Calendar API**.
   - *OAuth consent screen*: type External, add yourself as a **test user**.
   - *Credentials* → *Create credentials* → *OAuth client ID* → **Desktop app** → download JSON,
     save it as `credentials.json` next to `main.py`.
   - First `--push` opens a browser to log in; the token is cached in `token.json`.

## Usage
The semester number (1, 2 or 3) is a required first argument:

    python main.py 1                  # dry run: prints summary + warnings, writes preview_sem1.csv
    python main.py 1 --push           # create/update semester 1 events
    python main.py 1 --push --prune   # also remove events that disappeared from the plan
    python main.py 1 --delete-all     # remove everything this script created for semester 1

Re-running is safe: every event has a deterministic ID, so it is updated rather than duplicated.
By default each semester goes into its own calendar, e.g. "IO 2026Z – Semestr 2" (created if missing),
and events are tagged per semester, so `--prune`/`--delete-all` never touch another semester.

## How the sheet is read
- Only the column block under the `SEMESTR <n>` header of each week is used.
- Merged date/time cells are resolved, so parallel classes in one slot (e.g. PROJ-OG + NLPa-30) both become events.
- `PROa-do 17:45` → PROa ending at 17:45.
- Blocks merged across all semesters (Dzień rektorski, inauguracja, przerwa świąteczna) are added as
  notices in the calendar's default colour (`include_notices = false` to drop them).
- Each slot is a separate event (breaks between slots are kept).
- Colours: nearest Google colour to the cell colour in the sheet, one per subject. Google only has
  11 event colours; if a semester has more than 11 subjects, the ones with the fewest events share.
