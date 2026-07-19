---
name: wbs-gantt-excel
description: Generate a Japanese-style WBS / Gantt-chart Excel workbook (項番/タスクのタイトル/担当/SP/開始/終了/期間/完了率 columns, plus a day-by-day calendar grid with month/SP/week bands and colored Gantt bars) from a list of features, each broken into a free-form, ordered list of steps with their own assignee/start date/duration. Use this whenever the user asks to create, build, or update a WBS, スケジュール表, 工程表, or ガントチャート in Excel/xlsx form, especially when sprints ("SP") and Japanese public holidays need to be accounted for in the scheduling. SPs are defined by an explicit per-sprint start/end date list, so sprint lengths can vary freely. Also use it if the user references a WBS file shaped like "dPM WBS" (機能名 groups, each broken into コーディング/実装/試験/リリース-style steps with per-step assignees and business-day durations) and wants a new or refreshed version of it — step names, order, and count are all free-form per feature, and steps within a feature may have gaps between them or overlap. Don't use this for generic/ad-hoc spreadsheets that aren't schedule/timeline-shaped — plain data tables should just be written directly with openpyxl or pandas.
---

# WBS / Gantt Chart Excel Generator

Builds a two-sheet Excel workbook: a "ガントチャート" sheet (task table + a business-day calendar
with colored Gantt bars) and a "設定" sheet (the holiday list that drives all the date math). The
whole thing is produced by `scripts/generate_wbs.py`, which needs `openpyxl` and `PyYAML` (`pip
install pyyaml` if the environment doesn't already have it).

## Why this shape

Each feature ("機能") is broken into an ordered list of steps — typically something like
コーディング → 実装 → 試験 → リリース, but step names, order, and count are entirely free-form per
feature (add a custom step like テストデータ作成, skip steps a feature doesn't need, reorder them,
whatever fits). What's fixed per step is its shape: a name, an assignee, a start date, and a
duration in business days.

Every step carries its **own explicit start date** — there is no chaining that forces one step to
begin the business day after the previous one ends. That's deliberate: it lets steps within a
feature have a gap between them (e.g. a buffer week before 受け入れ試験) or overlap freely (e.g. a
テストデータ作成 step that runs in parallel with 実装), simply by choosing that step's start date.
End dates are still never typed by hand — each step's end is computed from its own start + duration
by walking business days (skipping weekends and the holidays in the 設定 sheet).

## Step-by-step workflow

1. **Get the tasks YAML.** Ask the user for a YAML file describing their features, or help them
   create one.
   - The required shape is `assets/tasks_template.yaml` — copy it to a working location and either
     have the user fill it in, or fill it in yourself from whatever the user describes in chat
     (feature names, step names, assignees, start dates, durations in business days). Shape:
     ```yaml
     features:
       - name: <機能名>
         steps:
           - name: <工程名>       # free-form, e.g. コーディング / 実装 / テストデータ作成 / ...
             assignee: <担当>     # may be "" if not yet assigned
             start: YYYY-MM-DD    # this step's own start date
             duration: <int>      # business days, must be a positive integer
           - name: <次の工程>
             ...
       - name: <次の機能>
         steps: [...]
     ```
   - `features` is a list of one or more features; each feature's `steps` is a list of one or more
     steps, in whatever order/names/count fits that feature — nothing is templated or fixed across
     features, so one feature can have a custom step another doesn't.
   - Every step's `start` is a real calendar date (YAML parses unquoted `YYYY-MM-DD` as a date
     automatically) — the script rolls it forward to the next business day if it lands on a
     weekend/holiday and reports the adjustment, then computes that step's end from `start` +
     `duration` by walking business days.
   - **Steps are scheduled independently — there is no auto-chaining.** If the user wants a
     traditional back-to-back schedule (each step starting the business day after the previous
     one ends), compute those start dates yourself when filling in the YAML — don't leave gaps
     unless the user actually wants a buffer, and don't assume the script will close them for you.
     Conversely, if the user wants a gap (buffer time) or an overlap (parallel work) between two
     steps in the same feature, just set that step's `start` accordingly — both are fully
     supported, no special flag needed.

2. **Get the holiday list.** Default to `assets/holidays_2026.csv` (real 2026 Japanese national
   holidays, same `date,name` shape as the reference file's 設定 sheet). If the schedule spans a
   different year or the user has their own company holiday list, ask for/build a CSV in the same
   `date,name` format (dates as `YYYY-MM-DD`) instead. This file becomes the workbook's 設定 sheet
   verbatim, so the user can hand-edit holidays later directly in Excel too.

3. **Confirm the SP (sprint) settings.** SPs are defined by one row per sprint in
   `assets/sp_list_template.csv`, each with its own explicit start and end date, so lengths can
   differ freely between sprints (e.g. a longer QA sprint, a short one around a holiday week):
   ```
   SP,開始日,終了日
   26,2026-01-19,2026-01-30
   27,2026-02-02,2026-02-13
   28,2026-02-16,2026-03-06
   29,2026-03-09,2026-03-20
   ```
   - This file is the default — if the user doesn't hand you a different SP list, the script reads
     `assets/sp_list_template.csv` as-is. Have the user fill it in (or fill it in yourself from
     what they describe), then either edit that template in place or pass a different file via
     `--sp-list` if they want to keep multiple schedules around.
   - Rows don't need to be pre-sorted — the script sorts by 開始日. Overlapping ranges are a hard
     error (fix the dates and rerun); a business-day gap between one SP's end and the next SP's
     start is only a warning (that stretch of days renders with a blank SP band — often fine if
     it's intentional buffer time).
   - `chart_start_date` / `chart_end_date` (see step 4) work on top of the SP list, e.g. to extend
     the calendar earlier/later than the SP list itself covers, and default to the SP list's own
     earliest start / latest end when omitted.

   If a task's computed end date falls beyond whatever the calendar ends up covering (e.g. a
   `chart_end_date` set too early, or an SP list that doesn't reach far enough), the script still
   generates the file but prints a warning per overflowing task — that bar just won't be visible
   until the range is widened and regenerated.

4. **Run the generator:**
   ```bash
   python3 scripts/generate_wbs.py \
     --tasks <path-to-tasks.yaml> \
     --holidays <path-to-holidays.csv> \
     --output <output.xlsx>
   ```
   `--sp-list <path-to-sp_list.csv>` is optional — omit it to use `assets/sp_list_template.csv` by
   default. `--chart-start-date` / `--chart-end-date` are also accepted for one-off tweaks to trim
   or extend the rendered calendar beyond what the SP list covers.

   The script prints a summary (feature/task counts, date range, SP range, any start-date
   rollovers, and any overflow/SP-gap warnings) and exits with a non-zero code plus a Japanese error
   message if the input has a problem (missing/invalid YAML key, bad date, missing/non-positive
   duration, overlapping SPs, PyYAML not installed, etc.) — read that message back to the user and
   fix the input rather than guessing.

5. **Report the output path to the user** and mention the summary line (date range, SP range,
   number of features/tasks, any adjustments). If they want to tweak durations, assignees, or
   step start dates (including adding gaps/overlaps), edit the tasks YAML and rerun; if they want
   to tweak sprint boundaries, edit the SP list and rerun — don't hand-edit the generated xlsx's
   date logic directly, since the whole point is that it's regenerable.

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
  user to fill in manually — it's not derived from the input YAML.
