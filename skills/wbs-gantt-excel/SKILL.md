---
name: wbs-gantt-excel
description: Generate a Japanese-style WBS / Gantt-chart Excel workbook (項番/タスクのタイトル/担当/SP/開始/終了/期間/完了率 columns, plus a day-by-day calendar grid with month/SP/week bands and colored Gantt bars) from a list of features and task assignees/durations. Use this whenever the user asks to create, build, or update a WBS, スケジュール表, 工程表, or ガントチャート in Excel/xlsx form, especially when sprints ("SP") and Japanese public holidays need to be accounted for in the scheduling. Supports both uniform-length SPs and SPs with individually-set, variable-length start/end dates per sprint. Also use it if the user references a WBS file shaped like "dPM WBS" (機能名 groups, each broken into a fixed sequence of コーディング/実装/試験/リリース-style steps with per-step assignees and business-day durations) and wants a new or refreshed version of it. Don't use this for generic/ad-hoc spreadsheets that aren't schedule/timeline-shaped — plain data tables should just be written directly with openpyxl or pandas.
---

# WBS / Gantt Chart Excel Generator

Builds a two-sheet Excel workbook: a "ガントチャート" sheet (task table + a business-day calendar
with colored Gantt bars) and a "設定" sheet (the holiday list that drives all the date math). The
whole thing is produced by `scripts/generate_wbs.py`, which only needs `openpyxl` — no other
Python packages.

## Why this shape

Every feature ("機能") goes through the *same* fixed 6-step process, so the task-name column never
needs to be typed out per feature — it's a template:

```
コーディング → 実装 → 実装後コーダー確認 → 試験 → 受け入れ試験 → リリース
```

A feature can skip any of these steps (e.g. no 受け入れ試験), but the order and names are fixed.
What varies per feature is: who's assigned to each step and how many business days it takes. Dates
are never typed by hand — they're computed by chaining business days (skipping weekends and the
holidays in the 設定 sheet) from each feature's start date through its steps in order.

## Step-by-step workflow

1. **Get the tasks CSV.** Ask the user for a CSV describing their features, or help them create one.
   - The required shape is `assets/tasks_template.csv` — copy it to a working location and either
     have the user fill it in, or fill it in yourself from whatever the user describes in chat
     (feature names, assignees, durations in business days). Columns:
     ```
     機能名,開始日,
     コーディング_担当,コーディング_日数,
     実装_担当,実装_日数,
     実装後コーダー確認_担当,実装後コーダー確認_日数,
     試験_担当,試験_日数,
     受け入れ試験_担当,受け入れ試験_日数,
     リリース_担当,リリース_日数
     ```
   - One row per feature. 開始日 is the calendar date (YYYY-MM-DD or YYYY/MM/DD) the *first*
     included step should start on — the script rolls it forward to the next business day if it
     lands on a weekend/holiday and reports the adjustment.
   - Leave a step's `_担当` cell blank to skip that step for that feature (then its `_日数` must
     also be blank). If `_担当` is filled, `_日数` is required and must be a positive integer
     (business days for that step).
   - Every other step's start is simply "the next business day after the previous step's end" —
     don't ask the user to compute this, the script does it.

2. **Get the holiday list.** Default to `assets/holidays_2026.csv` (real 2026 Japanese national
   holidays, same `date,name` shape as the reference file's 設定 sheet). If the schedule spans a
   different year or the user has their own company holiday list, ask for/build a CSV in the same
   `date,name` format (dates as `YYYY-MM-DD`) instead. This file becomes the workbook's 設定 sheet
   verbatim, so the user can hand-edit holidays later directly in Excel too.

3. **Confirm the SP (sprint) settings.** There are two mutually exclusive ways to define SPs — ask
   the user (or infer from context) which fits: do all sprints run the same length, or does the
   length vary sprint to sprint (e.g. a longer QA sprint, a short one around a holiday week)?

   **Option A — uniform SP length.** Settings live in a small `key,value` CSV —
   `assets/sp_config_template.csv` — rather than being typed on the command line each time, so the
   user can hand-edit and reuse it. Copy the template and fill in whichever of these the user cares
   about (blank/omitted rows fall back to sensible defaults, so only ask about settings they want
   to change):
   ```
   key,value
   chart_start_date,2026-01-19
   chart_end_date,2026-03-31
   sp_start_number,26
   sp_length_workdays,10
   num_sps,
   ```
   - `chart_start_date`: first date shown on the calendar grid. If left blank, the script falls
     back to the earliest 開始日 across the tasks CSV.
   - `chart_end_date`: last date the calendar must reach. Leave blank if you'd rather specify
     `num_sps` directly, or leave both blank to auto-size the calendar to cover every task's end
     date plus one buffer sprint. **`chart_end_date` and `num_sps` are mutually exclusive** — fill
     in at most one of them, otherwise the script errors out asking you to pick one.
   - `sp_start_number` (default 26): the SP number of the first sprint band.
   - `sp_length_workdays` (default 10): business days per SP band (10 = two 5-day weeks), applied
     uniformly to every SP.
   - `num_sps`: an explicit fixed number of SPs to render, as an alternative to `chart_end_date`.

   **Option B — variable SP length, one row per sprint.** Use this whenever SPs don't all share the
   same duration. Instead of `sp_config`, fill in `assets/sp_list_template.csv`: one row per SP,
   each with its own explicit start and end date, so lengths can differ freely between sprints:
   ```
   SP,開始日,終了日
   26,2026-01-19,2026-01-30
   27,2026-02-02,2026-02-13
   28,2026-02-16,2026-03-06
   29,2026-03-09,2026-03-20
   ```
   - Rows don't need to be pre-sorted — the script sorts by 開始日. Overlapping ranges are a hard
     error (fix the dates and rerun); a business-day gap between one SP's end and the next SP's
     start is only a warning (that stretch of days renders with a blank SP band — often fine if
     it's intentional buffer time).
   - This mode replaces `sp_start_number` / `sp_length_workdays` / `num_sps` entirely — don't set
     those (via `--sp-config` or the matching CLI flags) alongside `--sp-list`, the script errors
     out asking you to pick one mode. `chart_start_date` / `chart_end_date` still work on top of it
     (e.g. to extend the calendar earlier/later than the SP list itself covers) and default to the
     SP list's own earliest start / latest end when omitted.

   Either way: if a task's computed end date falls beyond whatever the calendar ends up covering
   (e.g. a `chart_end_date` set too early, or an SP list that doesn't reach far enough), the script
   still generates the file but prints a warning per overflowing task — that bar just won't be
   visible until the range is widened and regenerated.

4. **Run the generator:**
   ```bash
   # uniform SP length
   python3 scripts/generate_wbs.py \
     --tasks <path-to-tasks.csv> \
     --holidays <path-to-holidays.csv> \
     --output <output.xlsx> \
     --sp-config <path-to-sp_config.csv>

   # variable SP length
   python3 scripts/generate_wbs.py \
     --tasks <path-to-tasks.csv> \
     --holidays <path-to-holidays.csv> \
     --output <output.xlsx> \
     --sp-list <path-to-sp_list.csv>
   ```
   Individual `--chart-start-date` / `--chart-end-date` / `--sp-start-number` / `--sp-length-workdays`
   / `--num-sps` CLI flags are also accepted for one-off tweaks in uniform mode — any flag passed on
   the command line overrides the matching value from `--sp-config`, so you don't need to edit the
   file for a quick experiment. In variable mode only `--chart-start-date` / `--chart-end-date` are
   meaningful alongside `--sp-list`.

   The script prints a summary (feature/task counts, date range, SP range, any start-date
   rollovers, and any overflow/SP-gap warnings) and exits with a non-zero code plus a Japanese error
   message if the input has a problem (missing column, bad date, missing/non-positive duration,
   conflicting `chart_end_date`/`num_sps`, overlapping SPs, mixing both SP modes, etc.) — read that
   message back to the user and fix the input rather than guessing.

5. **Report the output path to the user** and mention the summary line (date range, SP range,
   number of features/tasks, any adjustments). If they want to tweak durations or assignees, edit
   the tasks CSV and rerun; if they want to tweak sprint boundaries, edit `sp_config`/`sp_list` and
   rerun — don't hand-edit the generated xlsx's date logic directly, since the whole point is that
   it's regenerable.

## Notes on what the workbook contains

- Column D (SP) and the day-by-day grid's month/SP/week bands are derived purely from the
  computed dates — you never need to fill these in by hand.
- The Gantt bars are live Excel conditional formatting (`=AND(I$4>=$E<row>,I$4<=$F<row>)`), so if
  someone opens the file and drags a task's 開始/終了 dates around by hand, the bar moves with it.
  The fill used by that rule sets both `fgColor` and `bgColor` to the same color — a dxf (the style
  block a conditional-formatting rule points to) is rendered differently from a normal cell style,
  and some Excel builds (notably Excel for Mac) paint a dxf's "solid" pattern from `bgColor` rather
  than `fgColor`; setting only `fgColor` renders fine in openpyxl/most Windows Excel but shows up
  blank/white on those builds, so both are always set together for this fill.
  The day-header row (row 4) is also a live `WORKDAY()` formula chain against the 設定 sheet, so
  editing/adding holidays there reflows the visible calendar dates — but the merged month/SP/week
  *band boundaries* are fixed at generation time (same limitation as hand-built Excel Gantt charts:
  regenerate via the script if holidays change enough to shift band widths).
- 期間 (column G) is a live `NETWORKDAYS()` formula; タスク完了率 (column H) is left at 0% for the
  user to fill in manually — it's not derived from the input CSV.
