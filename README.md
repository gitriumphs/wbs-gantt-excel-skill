# wbs-gantt-excel

日本語形式のWBS/ガントチャートExcelワークブック（項番/タスクのタイトル/担当/SP/開始/終了/期間/完了率の列と、
月・SP・週バンド＋色付きガントバー付きの日次カレンダーグリッド）を生成する [Agent Skill](https://code.claude.com/docs/en/skills) です。

祝日（`assets/holidays_2026.csv`）と週末を考慮した営業日ベースの日付計算に対応し、SPごとに固定長の場合と可変長（開始/終了日を個別設定）の場合の両方をサポートします。

## インストール

[skills CLI](https://github.com/vercel-labs/skills) を使ってインストールできます。

```
npx skills add https://github.com/gitriumphs/wbs-gantt-excel-skill --skill wbs-gantt-excel
```

## 使い方

Claude（または対応するコーディングエージェント）に「WBSを作って」「ガントチャートを更新して」のように依頼すると、
このスキルが `scripts/generate_wbs.py`（依存パッケージは `openpyxl` のみ）を使ってExcelワークブックを生成します。

詳細な手順・CSVフォーマットは [`skills/wbs-gantt-excel/SKILL.md`](skills/wbs-gantt-excel/SKILL.md) を参照してください。

## ライセンス

[MIT](LICENSE)
