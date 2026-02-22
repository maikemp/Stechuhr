# Stechuhr

A privacy-friendly CLI time tracker that stores everything in plain Excel files.

- **Fast**: Clock in/out with a single terminal command
- **Private**: All data stays local in Excel files - no cloud, no account
- **Transparent**: Open and edit your timesheet in any spreadsheet app
- **Simple**: One Python package, no database, no config servers
- **Yours**: Plain .xlsx files you can backup, share, or archive however you want

## Features

- **Clock in/out** from the terminal with `stempel ein` / `stempel aus`
- **Home office vs. Office**: On first clock-in, use `--home` for home office or omit it for office work. This is set once per day.
- **Travel time offset**: For office days, automatically adds configurable minutes (default: 2) to the first arrival and last departure. Home office days get no offset.
- **Automatic break deduction**: If you work more than 6 hours without a gap, 30 minutes are automatically deducted (proportionally for 6-6.5h).
- **Sick leave**: Use `--krank` on clock-out to count the day as fully worked when you have to leave early due to illness.
- **Manual entries**: Forgot to stamp? Use `stempel nachtrag` to add entries retroactively.
- **Overtime balance**: `stempel saldo` shows your cumulative overtime across all tracked days.
- **Week overview**: `stempel woche` shows the current week at a glance.
- **Live tracking**: `stempel status` and `stempel woche` include time worked so far when you're currently clocked in.
- **Open Excel**: `stempel excel` opens the spreadsheet directly.
- **Forgotten stamp warning**: When clocking in, you get a warning if the previous workday has an unclosed entry.
- **Color-coded output**: Positive saldo in green, negative in red.
- **Missing days**: Days without any stamps are assumed to have been worked as expected (no overtime, no undertime). This also covers **vacation and sick days** -- if you don't stamp on a day, it counts as if you worked your expected hours (saldo = 0).
- **Flexible schedule**: Only days with expected hours > 0 appear in the Excel. You can clock in on any day (including weekends) -- those days are added automatically and count as overtime.
- **Excel output**: One `.xlsx` file per year, 12 sheets (one per month), with formatted rows for each day with expected hours.
- **Manual Excel editing**: Edit times directly in the spreadsheet, then run `stempel update` to recalculate hours and balances.
- **Shell completion**: Tab completion for bash, zsh, and fish.

## Installation

Requires Python 3.9+.

```bash
pip install .
```

This installs the `stempel` command globally.

### Shell completion (optional)

For **zsh** (add to `~/.zshrc`):

```bash
eval "$(_STEMPEL_COMPLETE=zsh_source stempel)"
```

For **bash** 4.4+ (add to `~/.bashrc`):

```bash
eval "$(_STEMPEL_COMPLETE=bash_source stempel)"
```

For **fish** (run once):

```bash
_STEMPEL_COMPLETE=fish_source stempel > ~/.config/fish/completions/stempel.fish
```

## Usage

### Clock in and out

```bash
stempel ein                     # Clock in (office)
stempel ein --home              # Clock in (home office)
stempel aus                     # Clock out
```

The `--home` flag is only relevant on the **first clock-in of the day**. It determines whether the day is treated as a home office day (no travel time offset) or an office day (+2 minutes on first arrival and last departure).

### Multiple clock-ins per day

You can clock in and out multiple times per day (e.g. for a lunch break):

```bash
stempel ein                     # 09:00 - arrive at office
stempel aus                     # 12:00 - lunch break
stempel ein                     # 12:45 - back from lunch
stempel aus                     # 17:00 - leave for the day
```

The travel time offset is always applied to the **first arrival** and **last departure** only. When you clock in again after a break, the offset on the previous clock-out is automatically removed (recalculated), and will be applied to the new last clock-out instead.

### Sick leave

If you get sick during the day and need to leave early:

```bash
stempel aus --krank
```

This records your actual clock-out time but sets the day's total to the expected hours, so you don't lose any hours. The Status column is set to "Krank" and your actual stamps are preserved in the Excel for reference.

For full sick days where you don't come in at all, simply don't stamp -- the day will automatically count as if you worked the expected hours (saldo = 0).

### Override time or date

```bash
stempel ein --time 08:30
stempel aus --time 17:00 --date 2026-02-18
```

### Manual entry (Nachtrag)

For days you forgot to stamp:

```bash
stempel nachtrag --date 2026-02-17 --ein 09:00 --aus 15:30
stempel nachtrag --date 2026-02-17 --ein 09:00 --aus 15:30 --home
```

### Recalculate after manual Excel edits

If you edit times directly in the Excel file, run `update` to recalculate hours and balances:

```bash
stempel update                      # Recalculate entire current year
stempel update --month 2026-02      # Recalculate one month
stempel update --date 2026-02-18    # Recalculate one day
```

### Check overtime

```bash
stempel saldo                   # Cumulative balance up to yesterday
```

### Today's status

```bash
stempel status
```

When you're currently clocked in, this shows a live estimate (marked with `~`) of hours worked so far.

### Week overview

```bash
stempel woche
```

Example output (while clocked in on Wednesday):

```
Woche 16.02. - 20.02.2026:

  Mo 16.02.  Gesamt:  8:00  Soll:  8:00  Saldo: +0:00
  Di 17.02.  Gesamt:  8:00  Soll:  8:00  Saldo: +0:00
  Mi 18.02.  Gesamt: ~5:30  Soll:  8:00  Saldo: ~-2:30 <--
  Do 19.02.  (noch offen)
  Fr 20.02.  (noch offen)

  Woche:     Gesamt: 21:30  Soll: 40:00  Saldo: -18:30
```

The `~` prefix indicates a live estimate that includes time since the last clock-in.

### Open Excel file

```bash
stempel excel                   # Opens the current year's .xlsx
```

### View configuration

```bash
stempel config
```

## Configuration

On first run, a config file is created at `~/.config/stechuhr/config.json`:

```json
{
  "data_dir": "~/.config/stechuhr/data",
  "travel_offset_minutes": 2,
  "expected_hours": {
    "monday": 8.0,
    "tuesday": 8.0,
    "wednesday": 8.0,
    "thursday": 8.0,
    "friday": 8.0
  },
  "auto_break_threshold_hours": 6.0,
  "auto_break_deduction_minutes": 30,
  "carry_over_balance": {}
}
```

| Field | Description |
|-------|-------------|
| `data_dir` | Where `.xlsx` files are stored |
| `travel_offset_minutes` | Minutes added/subtracted for commute time (office days only) |
| `expected_hours` | Target hours per weekday (days with 0 hours won't appear in Excel unless you clock in) |
| `auto_break_threshold_hours` | Threshold above which break is auto-deducted |
| `auto_break_deduction_minutes` | Maximum auto-deduction (30 min) |
| `carry_over_balance` | Year-to-year overtime carry-over (e.g. `{"2025": 5.0}`) |

Edit this file directly to change your settings.

## Excel file structure

Files are stored as `{YYYY}.xlsx` (e.g. `2026.xlsx`) in the data directory, with one sheet per month.

Each sheet contains:

| Column | Content |
|--------|---------|
| Datum | Date (DD.MM.YYYY) |
| Tag | Weekday (Mo, Di, Mi, Do, Fr) |
| Ein 1, Aus 1, Stunden 1 | First stamp block |
| Ein 2, Aus 2, Stunden 2 | Second stamp block |
| Ein 3, Aus 3, Stunden 3 | Third stamp block (expands if needed) |
| Status | "Home" (home office), "Office", or "Krank" (sick leave) |
| Gesamt | Total hours worked (after break deduction) |
| Soll | Expected hours |
| Saldo | Balance (Gesamt - Soll) |

At the bottom of each sheet: **Summe** (monthly totals), **Uebertrag** (carry-over from previous month), and **Kumuliert** (cumulative balance).

You can edit the Excel manually. After editing, run `stempel update` to recalculate all derived values.

## How the travel time offset works

The travel time offset (default: 2 minutes) accounts for the time walking from the office door to your desk.

| Situation | What happens |
|-----------|--------------|
| Office day (no `--home`) | +2 min on first clock-in, +2 min on last clock-out |
| Home office (`--home`) | No offset applied |
| Multiple clock-ins/outs | Offset only on first arrival and final departure |

The offset is recalculated from scratch every time you clock in or out, so it always applies correctly to the current first/last stamps.

## How the break deduction works

| Situation | What happens |
|-----------|--------------|
| Worked <= 6h | No deduction |
| Worked 6h-6.5h | Deduct the amount exceeding 6h (e.g. 6h15m -> deduct 15min -> 6h) |
| Worked > 6.5h | Deduct 30 minutes |

## Vacation and sick days

- **Full day off** (vacation, sick, holiday): Simply don't stamp. The day automatically counts as if you worked the expected hours (saldo = 0).
- **Leaving early due to illness**: Use `stempel aus --krank`. This records your actual times but credits you with the full expected hours.

## License

MIT
