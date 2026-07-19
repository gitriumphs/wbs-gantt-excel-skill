---
name: wbs-gantt-excel
description: Generate a Japanese-style WBS / Gantt-chart Excel workbook (項番/タスクのタイトル/担当/SP/開始/終了/期間/完了率 columns, plus a day-by-day calendar grid with month/SP/week bands and colored Gantt bars) from a list of features, each broken into a free-form, ordered list of steps with their own assignee/start date/duration. Use this whenever the user asks to create, build, or update a WBS, スケジュール表, 工程表, or ガントチャート in Excel/xlsx form, especially when sprints ("SP") and Japanese public holidays need to be accounted for in the scheduling. SPs are defined by an explicit per-sprint start/end date list, so sprint lengths can vary freely. Also use it if the user references a WBS file shaped like "WBS" (機能名 groups, each broken into コーディング/実装/試験/リリース-style steps with per-step assignees and business-day durations) and wants a new or refreshed version of it — step names, order, and count are all free-form per feature, and steps within a feature may have gaps between them or overlap. Don't use this for generic/ad-hoc spreadsheets that aren't schedule/timeline-shaped — plain data tables should just be written directly with openpyxl or pandas.
---

# WBS / ガントチャート Excel ジェネレーター

2枚のシートからなる Excel ワークブックを作成します。「ガントチャート」シート（タスク表 + 色付きガントバー付きの営業日カレンダー）と「設定」シート（日付計算の基になる祝日リスト）です。すべて `scripts/generate_wbs.py` によって生成され、`openpyxl` と `PyYAML` が必要です（環境に入っていなければ `pip install pyyaml`）。

## この構成にしている理由

各機能（「機能」）は、順序付きの工程リストに分解されます。典型的には コーディング → 実装 → 試験 → リリース のような流れですが、工程名・順序・数は機能ごとに完全に自由です（テストデータ作成のようなカスタム工程を追加したり、その機能に不要な工程を省いたり、順序を入れ替えたり、何でも構いません）。工程ごとに固定されているのは、その形だけです。すなわち、名前・担当・開始日・所要日数（営業日）です。

各工程は**それぞれ独自の明示的な開始日**を持ちます。前の工程が終わった翌営業日に自動的に始まる、というような連鎖はありません。これは意図的な仕様です。これにより、機能内の工程間に間隔を空けたり（例：受け入れ試験の前にバッファ週を挟む）、自由に重複させたり（例：実装と並行して走るテストデータ作成工程）が、その工程の開始日を選ぶだけで実現できます。終了日は手入力されることはなく、常に各工程の「開始日 + 所要日数」から、営業日を数える（土日と設定シートの祝日を除く）ことで計算されます。

## 作業手順

1. **タスク YAML を用意する。** ユーザーに、機能を記述した YAML ファイルを求めるか、一緒に作成します。
   - 必要な形式は `assets/tasks_template.yaml` です。これを作業用の場所にコピーし、ユーザーに記入してもらうか、チャットで説明された内容（機能名、工程名、担当、開始日、所要日数（営業日））をもとに自分で埋めます。形式：
     ```yaml
     features:
       - name: <機能名>
         steps:
           - name: <工程名>       # 自由形式。例: コーディング / 実装 / テストデータ作成 / ...
             assignee: <担当>     # 未アサインの場合は "" でも可
             start: YYYY-MM-DD    # この工程自体の開始日
             duration: <int>      # 営業日数。正の整数であること
           - name: <次の工程>
             ...
       - name: <次の機能>
         steps: [...]
     ```
   - `features` は1つ以上の機能のリストで、各機能の `steps` は1つ以上の工程のリストです。順序・名前・数はその機能に合わせて自由に決められ、機能間でテンプレート化・固定はされません。つまり、ある機能だけにしかないカスタム工程があっても構いません。
   - すべての工程の `start` は実際のカレンダー日付です（YAML はクォートなしの `YYYY-MM-DD` を日付として自動的にパースします）。土日・祝日に当たる場合、スクリプトが次の営業日へ繰り上げ、その調整内容を報告した上で、`start` + `duration` からその工程の終了日を営業日ベースで計算します。
   - **各工程は独立してスケジュールされ、自動連鎖はありません。** ユーザーが従来型の連続スケジュール（各工程が前の工程の終了翌営業日に開始する形）を望む場合は、YAML を記入する際に開始日を自分で計算してください。ユーザーが実際にバッファを望んでいない限り間隔を空けたままにせず、スクリプトが自動で詰めてくれると仮定しないでください。逆に、同じ機能内の2工程間に間隔（バッファ時間）や重複（並行作業）を持たせたい場合は、その工程の `start` を適切に設定するだけで構いません。どちらも特別なフラグなしで完全にサポートされています。

2. **祝日リストを用意する。** デフォルトは `assets/holidays_2026.csv`（2026年の実際の日本の祝日一覧。参照ファイルの設定シートと同じ `date,name` 形式）です。スケジュールが別の年にまたがる場合や、ユーザー独自の会社の休日リストがある場合は、同じ `date,name` 形式（日付は `YYYY-MM-DD`）の CSV を求めるか作成してください。このファイルはそのままワークブックの設定シートになるため、ユーザーは後で Excel 上で直接祝日を手編集することもできます。

3. **SP（スプリント）設定を確認する。** SP は `assets/sp_list_template.csv` に1スプリント1行の形で定義され、それぞれ独自の明示的な開始日・終了日を持つため、スプリントごとに長さを自由に変えられます（例：長めの QA スプリント、祝日週前後の短いスプリントなど）：
   ```
   SP,開始日,終了日
   26,2026-01-19,2026-01-30
   27,2026-02-02,2026-02-13
   28,2026-02-16,2026-03-06
   29,2026-03-09,2026-03-20
   ```
   - このファイルがデフォルトです。ユーザーが別の SP リストを渡さない限り、スクリプトは `assets/sp_list_template.csv` をそのまま読み込みます。ユーザーに記入してもらうか、説明された内容から自分で記入し、そのままテンプレートを編集するか、複数のスケジュールを保持したい場合は `--sp-list` で別ファイルを指定してください。
   - 行は事前にソートされている必要はありません。スクリプトが開始日でソートします。範囲が重複している場合はエラーになります（日付を修正して再実行してください）。あるSPの終了日と次のSPの開始日の間の営業日ギャップは警告のみです（その期間は SP 帯が空欄で描画されます。意図的なバッファ期間であれば問題ないことが多いです）。
   - `chart_start_date` / `chart_end_date`（手順4参照）は SP リストの上に重ねて機能し、例えば SP リストがカバーする範囲よりもカレンダーを前後に拡張する際に使います。省略した場合は SP リスト自体の最も早い開始日／最も遅い終了日がデフォルトになります。

   タスクの計算上の終了日が、カレンダーがカバーする範囲を超える場合（例：`chart_end_date` が早すぎる、または SP リストの範囲が足りない場合）、スクリプトはファイルは生成しますが、はみ出したタスクごとに警告を出します。その場合、範囲を広げて再生成するまで、そのバーは表示されません。

4. **ジェネレーターを実行する：**
   ```bash
   python3 scripts/generate_wbs.py \
     --tasks <path-to-tasks.yaml> \
     --holidays <path-to-holidays.csv> \
     --output <output.xlsx>
   ```
   `--sp-list <path-to-sp_list.csv>` は省略可能です。省略するとデフォルトで `assets/sp_list_template.csv` が使われます。`--chart-start-date` / `--chart-end-date` も、SP リストがカバーする範囲を超えて描画カレンダーを一時的に調整（トリムまたは拡張）したい場合に指定できます。

   スクリプトはサマリー（機能数／タスク数、日付範囲、SP範囲、開始日の繰り上げ調整、はみ出し／SPギャップの警告など）を出力し、入力に問題がある場合（YAML キーの欠落・不正、不正な日付、所要日数の欠落・非正数、SP の重複、PyYAML 未インストールなど）は日本語のエラーメッセージとともに0以外の終了コードで終了します。そのメッセージをユーザーに伝え、推測で直さずに入力を修正してください。

5. **出力パスをユーザーに報告し**、サマリー行（日付範囲、SP範囲、機能数／タスク数、調整内容）を伝えます。所要日数・担当・工程の開始日（間隔や重複の追加を含む）を調整したい場合はタスク YAML を編集して再実行し、スプリントの境界を調整したい場合は SP リストを編集して再実行してください。生成された xlsx の日付ロジックを直接手編集しないでください。これは再生成可能であることこそがこの仕組みの目的だからです。

## ワークブックの内容について

- D列（SP）と日別グリッドの月／SP／週の帯は、すべて計算済みの日付から導出されます。手動で埋める必要はありません。
- ガントバーはライブの Excel 条件付き書式（`=AND(I$4>=$E<row>,I$4<=$F<row>)`）です。そのため、誰かがファイルを開いてタスクの開始／終了日をドラッグして手動で変更すると、バーも連動して動きます。この規則が使う塗りつぶしは `fgColor` と `bgColor` の両方に同じ色を設定しています。dxf（条件付き書式のルールが参照するスタイルブロック）は通常のセルスタイルとは異なる描画がされ、一部の Excel ビルド（特に Mac 版 Excel）では dxf の「ソリッド」パターンを `fgColor` ではなく `bgColor` から描画します。`fgColor` だけを設定すると openpyxl やほとんどの Windows 版 Excel では正しく表示されますが、これらのビルドでは空白／白のまま表示されてしまうため、この塗りつぶしについては常に両方を設定しています。
  日付ヘッダー行（4行目）もライブの `WORKDAY()` 数式チェーンとして設定シートを参照しているため、そこで祝日を編集・追加すると表示上のカレンダー日付が再計算されます。ただし、結合された月／SP／週の**帯の境界**は生成時点で固定されます（手作業で作った Excel ガントチャートと同じ制約です。帯の幅が変わるほど祝日が変化した場合は、スクリプトで再生成してください）。
- 期間（G列）はライブの `NETWORKDAYS()` 数式です。タスク完了率（H列）は 0% のままユーザーが手動で入力する想定であり、入力 YAML から自動導出されるものではありません。
