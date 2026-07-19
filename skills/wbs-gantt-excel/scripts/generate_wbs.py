#!/usr/bin/env python3
"""Generate a Japanese WBS / Gantt-chart Excel workbook from a tasks YAML file and a holiday CSV.

Usage:
    python generate_wbs.py --tasks tasks.yaml --holidays holidays.csv --output out.xlsx \
        [--sp-list sp_list.csv] [--chart-start-date 2026-01-19] [--chart-end-date 2026-03-31]

--sp-list points to a CSV of SP,開始日,終了日 rows (see assets/sp_list_template.csv) that gives
each SP its own explicit, independently-sized date range, so SP lengths can vary freely from
sprint to sprint. If omitted, it defaults to assets/sp_list_template.csv next to this script.
--chart-start-date/--chart-end-date still work on top of it to trim/extend the rendered calendar
beyond what the SP list itself covers; both default to the SP list's own earliest start / latest
end when omitted.

--tasks points to a YAML file (see assets/tasks_template.yaml): a `features:` list, each with a
free-form, ordered `steps:` list (name/assignee/start/duration). Every step carries its own
explicit start date, so steps within a feature may have gaps between them or overlap freely --
there is no forced chaining from one step's end to the next step's start.

Depends on openpyxl and PyYAML (`pip install pyyaml` if not already available). See the skill's
SKILL.md for the file formats.
"""
import argparse
import csv
import datetime as dt
import os
import sys

from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FIRST_DAY_COL = 9  # column I
HOLIDAY_SHEET_LAST_ROW = 500  # generous fixed range for WORKDAY/NETWORKDAYS formulas
DEFAULT_SP_LIST_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "assets", "sp_list_template.csv"
)

HEADER_FILL = PatternFill("solid", fgColor="FF0B5394")
SP_FILL = PatternFill("solid", fgColor="FF990000")
WEEK_FILL = PatternFill("solid", fgColor="FF2F75B5")
DATE_FILL = PatternFill("solid", fgColor="FFBDD7EE")
GROUP_FILL = PatternFill("solid", fgColor="FFCCCCCC")
# Both fgColor and bgColor are set (not just fgColor) because this fill is used inside a
# conditional-formatting dxf: unlike a normal cell style, some Excel builds (notably Excel for
# Mac) render a dxf's "solid" pattern from bgColor rather than fgColor, so a fgColor-only fill
# shows as blank/white there even though it looks correct in openpyxl/Windows Excel.
GANTT_FILL = PatternFill(fill_type="solid", start_color="FF5B9BD5", end_color="FF5B9BD5")

HEADER_FONT = Font(bold=True, size=9, color="FFFFFFFF")
DATE_FONT = Font(bold=True, size=9, color="FF000000")
GROUP_NUM_FONT = Font(bold=True, size=11, color="FF000000")
GROUP_TITLE_FONT = Font(bold=True, size=11, color="FF0000FF")
TASK_FONT = Font(size=10, color="FF434343")
COL_HEADER_FONT = Font(bold=True, size=8, color="FF000000")

CENTER = Alignment(horizontal="center", vertical="center")
CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")
THIN_LEFT = Border(left=Side(style="thin"))


class InputError(ValueError):
    pass


# ---------------------------------------------------------------------------
# CSV / date helpers
# ---------------------------------------------------------------------------

def read_csv_rows(path):
    last_err = None
    for enc in ("utf-8-sig", "cp932"):
        try:
            with open(path, newline="", encoding=enc) as f:
                reader = csv.DictReader(f)
                rows = list(reader)
                fieldnames = reader.fieldnames
            return fieldnames, rows
        except UnicodeDecodeError as e:
            last_err = e
            continue
    raise InputError(f"CSVの文字コードを判定できませんでした(utf-8/cp932で読めません): {path} ({last_err})")


def parse_date(s, context=""):
    s = (s or "").strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise InputError(f"日付を解釈できません{(' (' + context + ')') if context else ''}: {s!r}")


def is_workday(d, holidays):
    return d.weekday() < 5 and d not in holidays


def next_workday(d, holidays):
    d = d + dt.timedelta(days=1)
    while not is_workday(d, holidays):
        d += dt.timedelta(days=1)
    return d


def roll_to_workday(d, holidays):
    while not is_workday(d, holidays):
        d += dt.timedelta(days=1)
    return d


def add_workdays(start, n, holidays):
    """start must already be a workday. Returns the date n workdays after start (n=0 -> start)."""
    d = start
    for _ in range(n):
        d = next_workday(d, holidays)
    return d


def workday_index(chart_start, date, holidays):
    """0-based position of `date` in the business-day sequence starting at chart_start."""
    if date < chart_start:
        return -1
    idx = 0
    d = chart_start
    while d < date:
        d = next_workday(d, holidays)
        idx += 1
    return idx


# ---------------------------------------------------------------------------
# Loading holidays / tasks
# ---------------------------------------------------------------------------

def load_holidays(path):
    fieldnames, rows = read_csv_rows(path)
    if not fieldnames or "date" not in fieldnames:
        raise InputError(f"祝日CSVに 'date' 列が見つかりません: {path}")
    holidays = {}
    for row in rows:
        raw = (row.get("date") or "").strip()
        if not raw:
            continue
        d = parse_date(raw, context=f"祝日ファイル {path}")
        holidays[d] = (row.get("name") or "").strip()
    return holidays


def load_yaml_module():
    try:
        import yaml
    except ImportError:
        raise InputError(
            "タスクYAMLの読み込みには PyYAML が必要です。`pip install pyyaml` を実行してから再実行してください。"
        )
    return yaml


def _require_mapping(obj, context):
    if not isinstance(obj, dict):
        raise InputError(f"{context} はマッピング(key: value)にしてください")
    return obj


def _coerce_date(value, context):
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    return parse_date(str(value), context=context)


def load_tasks(path):
    yaml = load_yaml_module()
    try:
        with open(path, encoding="utf-8") as f:
            data = yaml.safe_load(f)
    except yaml.YAMLError as e:
        raise InputError(f"タスクYAMLの構文エラーです: {path}\n{e}")

    if not isinstance(data, dict) or "features" not in data:
        raise InputError(f"タスクYAMLのトップレベルは 'features:' キーを持つマッピングにしてください: {path}")
    raw_features = data["features"]
    if not isinstance(raw_features, list) or not raw_features:
        raise InputError(f"タスクYAMLの 'features' は1件以上のリストにしてください: {path}")

    features = []
    for fi, raw_feat in enumerate(raw_features, start=1):
        _require_mapping(raw_feat, f"features[{fi}]")
        name = str(raw_feat.get("name") or "").strip()
        if not name:
            raise InputError(f"features[{fi}]: name が未指定です")

        raw_steps = raw_feat.get("steps")
        if not isinstance(raw_steps, list) or not raw_steps:
            raise InputError(f"「{name}」: steps は1件以上のリストにしてください")

        steps = []
        for si, raw_step in enumerate(raw_steps, start=1):
            _require_mapping(raw_step, f"「{name}」の steps[{si}]")
            step_name = str(raw_step.get("name") or "").strip()
            if not step_name:
                raise InputError(f"「{name}」の steps[{si}]: name が未指定です")

            assignee = str(raw_step.get("assignee") or "").strip()

            start_raw = raw_step.get("start")
            if start_raw is None or str(start_raw).strip() == "":
                raise InputError(f"「{name}」の「{step_name}」: start が未指定です")
            raw_start = _coerce_date(start_raw, context=f"「{name}」の「{step_name}」の start")

            dur_raw = raw_step.get("duration")
            if dur_raw is None or str(dur_raw).strip() == "":
                raise InputError(f"「{name}」の「{step_name}」: duration が未指定です")
            try:
                dur = int(dur_raw)
            except (TypeError, ValueError):
                raise InputError(f"「{name}」の「{step_name}」: duration は整数にしてください: {dur_raw!r}")
            if dur <= 0:
                raise InputError(f"「{name}」の「{step_name}」: duration は正の整数にしてください: {dur}")

            steps.append({"name": step_name, "assignee": assignee, "raw_start": raw_start, "duration": dur})
        features.append({"name": name, "steps": steps})

    if not features:
        raise InputError(f"タスクYAMLに有効なfeatureがありません: {path}")
    return features


def parse_int_setting(value, field_name):
    try:
        return int(value)
    except (TypeError, ValueError):
        raise InputError(f"{field_name} は整数で指定してください: {value!r}")


def load_sp_list(path):
    """Load an explicit per-SP date-range list from a SP,開始日,終了日 CSV.

    Each row gives one SP's own start/end date, so SPs can have different lengths. Returns a list
    of {"sp": int, "start": date, "end": date} sorted by start date.
    """
    fieldnames, rows = read_csv_rows(path)
    required = ["SP", "開始日", "終了日"]
    missing = [c for c in required if c not in (fieldnames or [])]
    if missing:
        raise InputError(f"SP一覧CSVに必要な列がありません: {missing}\n必要な列: {required}")

    entries = []
    for i, row in enumerate(rows, start=2):
        sp_raw = (row.get("SP") or "").strip()
        if not sp_raw:
            continue
        sp_number = parse_int_setting(sp_raw, f"SP一覧 行{i} の SP番号")
        start = parse_date(row.get("開始日"), context=f"SP一覧 行{i} (SP{sp_number}) の開始日")
        end = parse_date(row.get("終了日"), context=f"SP一覧 行{i} (SP{sp_number}) の終了日")
        if end < start:
            raise InputError(f"SP一覧 行{i} (SP{sp_number}): 終了日({end})が開始日({start})より前です")
        entries.append({"sp": sp_number, "start": start, "end": end})

    if not entries:
        raise InputError(f"SP一覧CSVに有効な行がありません: {path}")

    entries.sort(key=lambda e: e["start"])
    for a, b in zip(entries, entries[1:]):
        if a["end"] >= b["start"]:
            raise InputError(
                f"SP一覧: SP{a['sp']}({a['start']}〜{a['end']}) と "
                f"SP{b['sp']}({b['start']}〜{b['end']}) の期間が重複しています"
            )
    return entries


def sp_list_gap_warnings(entries, holidays):
    """Business days between consecutive SPs that no SP row covers (informational only)."""
    warnings = []
    for a, b in zip(entries, entries[1:]):
        d = next_workday(a["end"], holidays)
        if d < b["start"]:
            warnings.append(
                f"SP{a['sp']}の終了({a['end']})とSP{b['sp']}の開始({b['start']})の間に、"
                f"どのSPにも属さない営業日があります"
            )
    return warnings


def sp_label_for_date(d, entries):
    for e in entries:
        if e["start"] <= d <= e["end"]:
            return e["sp"]
    return None


def schedule_features(features, holidays, adjustments):
    """Each step carries its own explicit start date, rolled forward to the next business day if
    needed. Steps are scheduled independently of one another -- unlike the old fixed-step chain,
    nothing forces a step to start right after the previous one ends, so gaps and overlaps between
    steps within a feature are both allowed."""
    for feat in features:
        for step in feat["steps"]:
            start = roll_to_workday(step["raw_start"], holidays)
            if start != step["raw_start"]:
                adjustments.append(
                    f"「{feat['name']}」の「{step['name']}」: 開始日 {step['raw_start']} は非稼働日のため {start} に調整しました"
                )
            step["start"] = start
            step["end"] = add_workdays(start, step["duration"] - 1, holidays)
    return features


# ---------------------------------------------------------------------------
# Calendar / SP grid
# ---------------------------------------------------------------------------

def build_calendar_explicit(chart_start, chart_end, holidays):
    """Business-day calendar spanning chart_start..chart_end inclusive (for --sp-list mode).

    Chains next_workday() and stops once it would pass chart_end, rather than pre-computing a
    day count via workday_index() -- that count-based approach overshoots by one whenever
    chart_end itself falls on a weekend/holiday (it rounds up to the next workday on or after
    chart_end instead of stopping at the last workday on or before it).
    """
    if chart_end < chart_start:
        raise InputError(f"終了日(chart_end_date={chart_end})が開始日(chart_start_date={chart_start})より前です。")
    calendar = [chart_start]
    d = chart_start
    while True:
        nxt = next_workday(d, holidays)
        if nxt > chart_end:
            break
        calendar.append(nxt)
        d = nxt
    return calendar


def week_of_month(d):
    return (d.day - 1) // 7 + 1


def merge_band(ws, row, calendar, key_func, label_func, fill, font, align=CENTER):
    col = FIRST_DAY_COL
    i = 0
    n = len(calendar)
    while i < n:
        key = key_func(calendar[i])
        j = i
        while j + 1 < n and key_func(calendar[j + 1]) == key:
            j += 1
        start_col = FIRST_DAY_COL + i
        end_col = FIRST_DAY_COL + j
        for c in range(start_col, end_col + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = fill
            cell.font = font
            cell.alignment = align
            cell.border = THIN_LEFT
        if end_col > start_col:
            ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
        ws.cell(row=row, column=start_col, value=label_func(calendar[i]))
        i = j + 1


def merge_sp_band(ws, row, sp_labels, fill, font):
    """Render the SP band from a precomputed per-day label list (works for both uniform SP
    length and explicit --sp-list mode -- consecutive equal labels get merged into one cell)."""
    n = len(sp_labels)
    i = 0
    while i < n:
        label = sp_labels[i]
        j = i
        while j + 1 < n and sp_labels[j + 1] == label:
            j += 1
        start_col = FIRST_DAY_COL + i
        end_col = FIRST_DAY_COL + j
        for c in range(start_col, end_col + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = fill
            cell.font = font
            cell.alignment = CENTER
            cell.border = THIN_LEFT
        if end_col > start_col:
            ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
        ws.cell(row=row, column=start_col, value=(f"SP{label}" if label is not None else ""))
        i = j + 1


# ---------------------------------------------------------------------------
# Workbook assembly
# ---------------------------------------------------------------------------

def build_workbook(features, holidays, chart_start, sp_entries, chart_end=None):
    adjustments = []
    schedule_features(features, holidays, adjustments)

    for feat in features:
        for step in feat["steps"]:
            if workday_index(chart_start, step["start"], holidays) < 0:
                raise InputError(
                    f"「{feat['name']}」の「{step['name']}」の開始日({step['start']})が"
                    f"チャート開始日({chart_start})より前です。chart_start_date を早めてください。"
                )

    sp_warnings = []
    if chart_end is None:
        chart_end = max(e["end"] for e in sp_entries)
    calendar = build_calendar_explicit(chart_start, chart_end, holidays)
    num_sps = len(sp_entries)
    sp_labels = [sp_label_for_date(d, sp_entries) for d in calendar]
    sp_warnings.extend(sp_list_gap_warnings(sp_entries, holidays))
    if any(lbl is None for lbl in sp_labels):
        sp_warnings.append("表示範囲内にどのSPにも属さない営業日があります(SP列は空欄になります)")
    sp_numbers = sorted(e["sp"] for e in sp_entries)
    sp_range_label = f"SP{sp_numbers[0]}-SP{sp_numbers[-1]}"

    def sp_number_for(d):
        return sp_label_for_date(d, sp_entries)

    overflow = []
    for feat in features:
        for step in feat["steps"]:
            if step["end"] > calendar[-1]:
                overflow.append(
                    f"「{feat['name']}」の「{step['name']}」の終了日({step['end']})が"
                    f"表示範囲の終了({calendar[-1]})を超えています"
                )
    total_workdays = len(calendar)
    last_col = FIRST_DAY_COL + total_workdays - 1
    last_col_letter = get_column_letter(last_col)
    first_col_letter = get_column_letter(FIRST_DAY_COL)

    wb = Workbook()
    ws = wb.active
    ws.title = "ガントチャート"

    # --- fixed left-hand headers (A1:H5) ---
    headers = ["項番", "タスクのタイトル", "担当", "SP", "開始", "終了", "期間", "タスク\n完了率"]
    for idx, label in enumerate(headers, start=1):
        for r in range(1, 6):
            cell = ws.cell(row=r, column=idx)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
        ws.merge_cells(start_row=1, start_column=idx, end_row=5, end_column=idx)
        top = ws.cell(row=1, column=idx, value=label)
        top.alignment = CENTER_WRAP

    # --- month / SP / week / date / weekday bands ---
    merge_band(ws, 1, calendar, key_func=lambda d: (d.year, d.month),
               label_func=lambda d: f"{d.month}月", fill=HEADER_FILL, font=HEADER_FONT)
    merge_sp_band(ws, 2, sp_labels, fill=SP_FILL, font=HEADER_FONT)
    merge_band(ws, 3, calendar, key_func=lambda d: (d.year, d.month, week_of_month(d)),
               label_func=lambda d: f"第 {week_of_month(d)} 週", fill=WEEK_FILL, font=HEADER_FONT)

    for i, d in enumerate(calendar):
        col = FIRST_DAY_COL + i
        letter = get_column_letter(col)
        cell = ws.cell(row=4, column=col)
        if i == 0:
            cell.value = d
        else:
            prev_letter = get_column_letter(col - 1)
            cell.value = f"=WORKDAY({prev_letter}4,1,'設定'!$A$2:$A${HOLIDAY_SHEET_LAST_ROW})"
        cell.number_format = "d"
        cell.fill = DATE_FILL
        cell.font = DATE_FONT
        cell.alignment = CENTER
        cell.border = THIN_LEFT

        wcell = ws.cell(row=5, column=col, value=f'=TEXT({letter}4,"ddd")')
        wcell.fill = DATE_FILL
        wcell.font = DATE_FONT
        wcell.alignment = CENTER
        wcell.border = THIN_LEFT

    # --- group / task rows ---
    row = 6
    for gi, feat in enumerate(features, start=1):
        ws.cell(row=row, column=1, value=gi).font = GROUP_NUM_FONT
        ws.cell(row=row, column=2, value=feat["name"]).font = GROUP_TITLE_FONT
        for c in range(1, 9):
            ws.cell(row=row, column=c).fill = GROUP_FILL
        ws.cell(row=row, column=1).alignment = LEFT
        ws.row_dimensions[row].height = 21
        row += 1

        group_start_row = row
        for si, step in enumerate(feat["steps"], start=1):
            sp_number = sp_number_for(step["start"])
            values = {
                1: si,
                2: step["name"],
                3: step["assignee"],
                4: sp_number if sp_number is not None else "",
            }
            for col, val in values.items():
                c = ws.cell(row=row, column=col, value=val)
                c.font = TASK_FONT
                c.alignment = CENTER if col in (1, 4) else LEFT

            ecell = ws.cell(row=row, column=5, value=step["start"])
            ecell.number_format = 'm"/"d'
            ecell.font = TASK_FONT
            ecell.alignment = LEFT

            fcell = ws.cell(row=row, column=6, value=step["end"])
            fcell.number_format = 'm"/"d'
            fcell.font = TASK_FONT
            fcell.alignment = LEFT

            gcell = ws.cell(row=row, column=7,
                             value=f"=NETWORKDAYS(E{row},F{row},'設定'!$A$2:$A${HOLIDAY_SHEET_LAST_ROW})")
            gcell.font = TASK_FONT
            gcell.alignment = CENTER

            hcell = ws.cell(row=row, column=8, value=0)
            hcell.number_format = "0%"
            hcell.font = TASK_FONT
            hcell.alignment = CENTER

            row += 1
        group_end_row = row - 1

        formula = f"AND({first_col_letter}$4>=$E{group_start_row},{first_col_letter}$4<=$F{group_start_row})"
        rng = f"{first_col_letter}{group_start_row}:{last_col_letter}{group_end_row}"
        ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=GANTT_FILL))

    last_row = row - 1
    if last_row >= 6:
        ws.conditional_formatting.add(
            f"H6:H{last_row}",
            ColorScaleRule(start_type="num", start_value=0, start_color="FFF8696B",
                            mid_type="num", mid_value=0.5, mid_color="FFFFEB84",
                            end_type="num", end_value=1, end_color="FF63BE7B"),
        )

    # --- column widths / row heights / freeze panes ---
    widths = {"A": 6.13, "B": 27.75, "C": 10.0, "D": 4.38, "E": 5.0, "F": 5.0, "G": 5.13, "H": 6.38}
    for letter, w in widths.items():
        ws.column_dimensions[letter].width = w
    for col in range(FIRST_DAY_COL, last_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 3.0
    for r in range(1, 6):
        ws.row_dimensions[r].height = 17.25
    ws.freeze_panes = "I6"

    # --- 設定 sheet (holidays) ---
    ws2 = wb.create_sheet("設定")
    for i, (d, name) in enumerate(sorted(holidays.items()), start=2):
        ws2.cell(row=i, column=1, value=dt.datetime(d.year, d.month, d.day)).number_format = "yyyy/m/d"
        ws2.cell(row=i, column=2, value=name)
    ws2.column_dimensions["A"].width = 12
    ws2.column_dimensions["B"].width = 16

    summary = {
        "features": len(features),
        "tasks": sum(len(f["steps"]) for f in features),
        "chart_start": chart_start,
        "chart_end": calendar[-1],
        "num_sps": num_sps,
        "sp_range": sp_range_label,
        "adjustments": adjustments,
        "overflow": overflow,
        "sp_warnings": sp_warnings,
    }
    return wb, summary


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Generate a Japanese WBS/Gantt xlsx from a tasks CSV and holiday CSV.")
    p.add_argument("--tasks", required=True,
                    help="Tasks YAML path (features/steps, see assets/tasks_template.yaml)")
    p.add_argument("--holidays", required=True, help="Holiday CSV path (columns: date,name)")
    p.add_argument("--output", required=True, help="Output .xlsx path")
    p.add_argument("--sp-list", default=None,
                    help="CSV of SP,開始日,終了日 rows giving each SP its own explicit date range "
                         "(see assets/sp_list_template.csv). Defaults to that template file next "
                         "to this script if omitted.")
    p.add_argument("--chart-start-date", default=None,
                    help="First day shown on the chart, YYYY-MM-DD (default: the SP list's own "
                         "earliest 開始日)")
    p.add_argument("--chart-end-date", default=None,
                    help="Last day the calendar must cover, YYYY-MM-DD (default: the SP list's own "
                         "latest 終了日)")
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    try:
        sp_list_path = args.sp_list or DEFAULT_SP_LIST_PATH
        sp_entries = load_sp_list(sp_list_path)

        chart_end = parse_date(args.chart_end_date, context="終了日(chart_end_date)") if args.chart_end_date else None

        holidays = load_holidays(args.holidays)
        features = load_tasks(args.tasks)

        chart_start_auto = False
        if args.chart_start_date:
            chart_start = parse_date(args.chart_start_date, context="開始日(chart_start_date)")
        else:
            chart_start = min(e["start"] for e in sp_entries)
            chart_start_auto = True

        wb, summary = build_workbook(features, holidays, chart_start, sp_entries, chart_end)

        out_dir = os.path.dirname(os.path.abspath(args.output))
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        wb.save(args.output)
    except InputError as e:
        print(f"エラー: {e}", file=sys.stderr)
        return 1

    print(f"作成しました: {args.output}")
    print(f"機能数: {summary['features']} / タスク数: {summary['tasks']}")
    if chart_start_auto:
        print(f"(開始日が未指定のため、SP一覧の最早開始日 {summary['chart_start']} を使用しました)")
    print(f"期間: {summary['chart_start']} 〜 {summary['chart_end']} ({summary['sp_range']}, 全{summary['num_sps']}SP)")
    if summary["adjustments"]:
        print("開始日の調整:")
        for a in summary["adjustments"]:
            print(f"  - {a}")
    if summary["overflow"]:
        print("表示範囲を超えるタスク(SP一覧やchart_end_dateを広げてください):")
        for o in summary["overflow"]:
            print(f"  - {o}")
    if summary["sp_warnings"]:
        print("SPに関する注意:")
        for w in summary["sp_warnings"]:
            print(f"  - {w}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
