# Attendance Detention Tracker

Counts absence codes from a pasted attendance report and tells you who has
earned a detention. Replaces doing it by hand with a tally sheet.

**Everything runs on your own computer.** There are no network calls anywhere in
this program. Student names never leave the machine you run it on.

## What it does

You paste in the daily attendance report. The program picks out the codes you
care about, adds them to a running total per student, and reports the students
who have just hit the threshold. The Excel file you point it at is updated in
place. Once a week it can also tell you everyone who's earned a detention
*since the last time you asked* -- see [Weekly Detention Report](#weekly-detention-report)
below.

## Two ways to use it

| | |
|---|---|
| **Desktop app (recommended)** | Double-click `dist\AttendanceTracker.exe`. A window opens with fields to fill in and buttons to click -- no terminal, no typing commands, no Python required on the machine that runs it. |
| **Command line** | Run `attendance.py` directly with Python. Useful for scripting, or if you're the one maintaining the code. |

Both use the exact same logic (`attendance.py`) and read/write the exact same
kind of workbook, so you can mix and match freely.

### Desktop app

Nothing to install. Just run `AttendanceTracker.exe`. If you don't have that
file yet, or you've changed the code and need a new one, see
[Rebuilding the desktop app](#rebuilding-the-desktop-app-exe) below.

### Command line

You need Python 3.10 or newer.

```
pip install openpyxl
python make_template.py attendance.xlsx
```

That creates a blank workbook. Keep it somewhere student records normally live.

```
python attendance.py --workbook attendance.xlsx --input todays_report.txt
```

Or run it without `--workbook` / `--input` and it will ask for them:

```
python attendance.py
Path to the Excel workbook: attendance.xlsx
Paste the attendance report below.
When you are finished, press Ctrl-D (Mac/Linux) or Ctrl-Z then Enter (Windows).
```

Useful flags:

| Flag | What it does |
|---|---|
| `--workbook attendance.xlsx` | The Excel file to update. Asked for interactively if left out. |
| `--input todays_report.txt` | Read the pasted report from a file instead of typing/pasting it in |
| `--date 2026-09-14` | Record under a date other than today |
| `--dry-run` | Show what would change without saving |
| `--force` | Record anyway when that date is already logged |

## Creating a new workbook

- **Desktop app:** click **New Template...**, choose where to save it, and it
  fills in the Workbook field for you automatically.
- **Command line:** `python make_template.py attendance.xlsx`

Either one produces an empty workbook with the sheets described below, ready
to start pasting reports into.

## Daily use

**Desktop app:** paste the report into the big text box, check the workbook
path and date at the top, and click **Run**. The output panel shows what was
read, any warnings, the report date, any detentions earned that day, and any
detentions earned since the last weekly report was marked as sent (see below).
If it's been more than 7 days since a weekly report was last sent (or since
attendance started being logged, if none has ever been sent), a reminder is
printed too. Tick **Dry run** to preview without saving, or **Force** to
re-record a date that's already logged.

**Command line:** see the flags table above.

## Weekly Detention Report

Some schools only need the detention list on a set schedule (for example: check
every Tuesday, send the notice Thursday) even though attendance gets pasted in
every day. The desktop app has a separate section for this so the daily
workflow doesn't have to change at all:

- **Generate Weekly Report** lists every detention earned *since the last time
  you clicked Mark as Sent* -- not just today's, not a fixed 7-day window, but
  everything genuinely new. The very first time you use it (nothing marked
  sent yet), it shows the full history so far.
- **Mark as Sent** tells the app "I've got this list, start the next one from
  here." Click it only after you've actually copied the list into the email --
  running Generate again before that just shows the same list, nothing is lost.
  It also permanently archives the report to the **Weekly Reports** sheet: who
  was on it, the date range it covered, and the date it was sent -- a
  standing record you can always look back on, separate from and unaffected
  by the running totals.

**Where the "last sent" marker is stored:** the date you last marked as sent is written into the
workbook file itself, in a hidden sheet called `_Internal` -- not on the
computer, not in the app, not in any settings file. That means:

- Resetting, reinstalling, or replacing the computer running the app has no
  effect on it at all, because the app keeps no state of its own.
- As long as you keep using that same workbook file, it will always remember
  where it left off, on any machine.
- The only way to lose that memory is to lose or replace *that specific file*
  (e.g. restoring an old backup, or accidentally generating a new template
  instead of opening the existing one).

Back up the workbook the same way you'd back up any important file (File
Explorer copy, OneDrive, a network drive, etc.) and always point the app at
that same file.

## The workbook

| Sheet | What it holds |
|---|---|
| **Log** | Every code ever recorded. One row = one student, one code, one period, one date. **The only sheet holding real data.** |
| **Students** | Running totals. Every cell is a formula reading the Log, so you can click any number and see where it came from. Rebuilt each run. |
| **Detentions** | Every detention earned, dated. Columns F and G are yours to fill in. |
| **Weekly Reports** | A permanent record of every weekly report ever sent: who was on it, the date range it covered, and when it was sent. Only added to when you click **Mark as Sent** -- never rebuilt, so past reports never change. |
| **Settings** | Which codes count and how many it takes. Edit the shaded cells. |
| **_Internal** | Hidden. Just one thing: the date the weekly report was last marked as sent. Not meant to be edited by hand. |

**To fix a mistake:** delete the row from the Log sheet, save, run the program
again. Everything recalculates. Nothing is stored anywhere else (except the
weekly-report marker above, which is unaffected by Log edits).

## Changing which codes count

Open the Settings sheet and change `NO` to `YES`. Out of the box only `L`
counts, at three per detention. No code changes needed.

## How the counting works

Three occurrences earns a detention and the count restarts. Because the count
only ever resets by exactly three, the arithmetic is plain division: seven
occurrences is two detentions with one left on the clock. That is why deleting
a Log row always produces the right answer -- there is no hidden state.

By default, all periods with the same code on the same day count as **one**
occurrence -- a student marked `L` in three periods on one day still only
gains one `L` toward the threshold, not three. This is controlled by
`COUNT_MODE` near the top of `attendance.py`:

- `"per_day"` (the default) -- one per student, per code, per day, no matter
  how many periods it appears in.
- `"per_period"` -- each period column counts separately, so four periods of
  `L` in one day would add four.

## Safeguards

- **Duplicate pastes are refused.** If that date and code are already in the
  Log, the program stops rather than issue detentions nobody earned.
- **Near-miss names are flagged.** `Halvorson` arriving when `Halvorsen` already
  exists prints a warning instead of quietly starting a second tally.
- **Unreadable lines are reported**, never silently dropped.
- **`--dry-run`** (or the Dry Run checkbox) shows the outcome before anything
  is written.
- **A locked file gives a plain-English message.** If the workbook is open in
  Excel when you try to save, you're told to close it and try again --
  instead of a Python error, and nothing you entered is lost.

## Reading the report format

A line looks like:

```
Ellison, Margot<TAB>8<TAB>L-
Okafor, Chidi<TAB>7<TAB>T-PA T-PA T-PA
```

Everything after the second tab is period columns, five characters each, so the
position of a code is what tells the program which period it belongs to. If your
report uses a different column width, change `PERIOD_FIELD_WIDTH` at the top of
`attendance.py`.

## Rebuilding the desktop app (.exe)

`AttendanceTracker.exe` is a frozen snapshot -- it does not update itself.
You only need to rebuild it when the code actually changes: a new feature, a
bug fix, a tweak to the window or its labels, and so on. If nothing in
`attendance.py` or `attendance_gui.py` changes, the exe you already have is
good forever and none of this is necessary.

When you do need to rebuild:

```
pip install pyinstaller
python -m PyInstaller --onefile --windowed --name "AttendanceTracker" attendance_gui.py
```

This always creates a fresh `dist\` folder and puts the new exe at
`dist\AttendanceTracker.exe`, regardless of where you moved any earlier copy.
It's self-contained -- move that one file out of `dist\` to wherever you keep
it (Desktop, a shared drive, etc.) and it runs without Python installed. The
`dist\` and `build\` folders left behind are just build output; safe to
delete once you've moved the exe out.

## Privacy

The workbook contains student records. Treat it the way you treat any other
student record: keep it on protected storage, do not email it, do not upload it,
and do not commit a real one to this repository. The template in the repo is
empty by design.

## License

MIT, or whichever license you prefer -- add a LICENSE file before publishing.
