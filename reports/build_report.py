"""マーケティング本部 月次報告書（Excel）を作る。

坂田さんの「202609_マーケティング本部月次報告書.xlsx」と同じ2シート構成:
  - WEB・SNS数値: data/metrics.csv の数値を年度（4月始まり）ごとに並べる
  - トピック: topics/YYYY-MM.tsv（カテゴリー1, カテゴリー2, 詳細 のタブ区切り）

使い方:
  python3 build_report.py 2026-10
  → out/202610_マーケティング本部月次報告書.xlsx
"""

import csv
import sys
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

BASE = Path(__file__).resolve().parent
MONTHS = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]
SNS = ["X", "Instagram", "YouTube", "LINE"]
WEB = ["アクセス数", "WEBサイト（UU数）"]

HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOLD = Font(bold=True)


def fiscal_year(year, month):
    return year if month >= 4 else year - 1


def load_metrics():
    """{metric: {fiscal_year: {month: value}}}"""
    data = defaultdict(lambda: defaultdict(dict))
    with open(BASE / "data" / "metrics.csv", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            y, m = map(int, row["year_month"].split("-"))
            data[row["metric"]][fiscal_year(y, m)][m] = int(row["value"])
    return data


def load_topics(ym):
    path = BASE / "topics" / f"{ym}.tsv"
    if not path.exists():
        return []
    with open(path, encoding="utf-8") as f:
        return [
            (line.split("\t") + ["", "", ""])[:3]
            for line in f.read().splitlines()
            if line.strip() and not line.startswith("#")
        ]


def write_row(ws, r, label, values, extra=()):
    ws.cell(r, 1, label)
    for i, v in enumerate(values, start=2):
        ws.cell(r, i, v)
    for i, v in enumerate(extra, start=2 + len(values)):
        ws.cell(r, i, v)
    for c in range(1, 2 + len(values) + len(extra)):
        ws.cell(r, c).border = BORDER


def header(ws, r, title, extra):
    labels = [title] + [f"{m}月" for m in MONTHS] + list(extra)
    for c, text in enumerate(labels, start=1):
        cell = ws.cell(r, c, text)
        cell.font = BOLD
        cell.fill = HEADER_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center")


def stats(values):
    present = [v for v in values if v is not None]
    if not present:
        return None, None
    return sum(present), round(sum(present) / len(present))


def sns_sheet(ws, data, fy, month):
    r = 1
    header(ws, r, "フォロワー数", ["合計", "平均", "前月増減", "前年同月比"])
    for name in SNS:
        for label, year in (("前年度", fy - 1), ("今年度", fy)):
            vals = [data[name][year].get(m) for m in MONTHS]
            total, avg = stats(vals)
            extra = [total, avg]
            if year == fy:
                cur = data[name][fy].get(month)
                prev_m = 12 if month == 1 else month - 1
                prev = data[name][fy - 1 if month == 4 else fy].get(prev_m)
                last_year = data[name][fy - 1].get(month)
                extra += [
                    cur - prev if cur is not None and prev is not None else None,
                    cur - last_year if cur is not None and last_year is not None else None,
                ]
            r += 1
            write_row(ws, r, f"{name}（{label}）", vals, extra)
    r += 2
    for name in WEB:
        header(ws, r, name, ["合計", "平均"])
        years = sorted(data[name])
        for year in years:
            vals = [data[name][year].get(m) for m in MONTHS]
            r += 1
            write_row(ws, r, str(year), vals, stats(vals))
        r += 2
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0"
    ws.column_dimensions["A"].width = 22
    for col in "BCDEFGHIJKLM":
        ws.column_dimensions[col].width = 9
    for col in "NOPQ":
        ws.column_dimensions[col].width = 11


def topic_sheet(ws, year, month, topics):
    ws["A1"] = f"{year}年{month}月度 マーケティング本部月次報告"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "概要"
    ws["A3"], ws["B3"] = "キーメトリクス", "別シート参照"
    ws["A4"] = "実施内容"
    for c, text in enumerate(["カテゴリー1", "カテゴリー2", "詳細"], start=1):
        cell = ws.cell(5, c, text)
        cell.font = BOLD
        cell.fill = HEADER_FILL
        cell.border = BORDER
    last1 = None
    for r, (c1, c2, detail) in enumerate(topics, start=6):
        ws.cell(r, 1, c1 if c1 != last1 else "")
        ws.cell(r, 2, c2)
        ws.cell(r, 3, detail)
        for c in range(1, 4):
            ws.cell(r, c).border = BORDER
        last1 = c1 or last1
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 60


def main():
    ym = sys.argv[1]
    year, month = map(int, ym.split("-"))
    data = load_metrics()
    wb = Workbook()
    sns_sheet(wb.active, data, fiscal_year(year, month), month)
    wb.active.title = "WEB・SNS数値"
    topic_sheet(wb.create_sheet("トピック"), year, month, load_topics(ym))
    out = BASE / "out" / f"{year}{month:02d}_マーケティング本部月次報告書.xlsx"
    out.parent.mkdir(exist_ok=True)
    wb.save(out)
    print(out)


if __name__ == "__main__":
    main()
