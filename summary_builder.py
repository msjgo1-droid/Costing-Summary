"""
SUMMARY 차트(엑셀)를 새로 만들거나 기존 SUMMARY 파일에 새 제출 데이터를 병합해서
다시 만들어주는 모듈.

기존 SUMMARY 파일(COSTING BULK ... SUMMARY.xlsx)의 레이아웃을 최대한 그대로
재현한다:
  SEASON | NO | STYLE | DESCRIPTION | [FOB, CM, MARGIN%] x N (날짜별) | TOTAL FOB SAVING | REMARK
"""
import io
import re
import datetime as dt
from dataclasses import dataclass, field
from typing import Optional

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference

THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BLUE = "FF0070C0"
HEADER_FILL = PatternFill("solid", fgColor="FFD9E1F2")
FONT_NAME = "Calibri"


@dataclass
class SummaryRow:
    season: str
    style_no: str
    colorway: str  # "SOLID" / "HEATHER" / "" (컬러웨이 구분 없는 단일 항목)
    description: str
    no: str = ""
    dates: dict = field(default_factory=dict)  # {date: {"fob":float,"cm":float,"margin":float|None}}
    remark: str = ""

    @property
    def key(self):
        return (self.style_no, self.colorway)

    @property
    def style_cell_text(self):
        return f"{self.style_no} / {self.colorway}" if self.colorway else self.style_no


DATE_IN_TEXT_RE = re.compile(r"(\d{1,2})/(\d{1,2})")


def _parse_style_cell(text: str):
    if text is None:
        return "", ""
    text = str(text).strip()
    if "/" in text:
        style_no, colorway = text.split("/", 1)
        return style_no.strip(), colorway.strip()
    return text, ""


def parse_existing_summary(file_like) -> list[SummaryRow]:
    """기존 SUMMARY 엑셀에서 행 데이터를 읽어 SummaryRow 리스트로 반환."""
    wb = openpyxl.load_workbook(file_like, data_only=True)
    ws = wb[wb.sheetnames[0]]

    # 헤더 행 찾기 (SEASON/NO/STYLE/DESCRIPTION 이 있는 행)
    header_row = None
    for r in range(1, min(ws.max_row, 10) + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, 5)]
        if vals and isinstance(vals[0], str) and vals[0].strip().upper() == "SEASON":
            header_row = r
            break
    if header_row is None:
        return []

    # 날짜 그룹 컬럼 찾기: FOB.. 로 시작하는 헤더들을 순서대로 그룹핑
    col_headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if isinstance(v, str):
            col_headers[c] = v.strip().upper()

    date_groups = []  # [{"fob_col":c, "cm_col":c, "margin_col":c|None, "date":date}]
    cols_sorted = sorted(col_headers)
    i = 0
    while i < len(cols_sorted):
        c = cols_sorted[i]
        label = col_headers[c]
        if label.startswith("FOB"):
            fob_col = c
            cm_col = cols_sorted[i + 1] if i + 1 < len(cols_sorted) and col_headers[cols_sorted[i + 1]] == "CM" else None
            margin_col = None
            j = i + 2
            if cm_col and j < len(cols_sorted) and "MARGIN" in col_headers[cols_sorted[j]]:
                margin_col = cols_sorted[j]
                j += 1
            # 날짜 그룹 라벨은 보통 한 행 위(header_row-1) 병합 셀에 있음
            group_label = ws.cell(row=header_row - 1, column=fob_col).value or label
            m = DATE_IN_TEXT_RE.search(str(group_label))
            gdate = None
            if m:
                month, day = int(m.group(1)), int(m.group(2))
                try:
                    gdate = dt.date(dt.date.today().year, month, day)
                except ValueError:
                    gdate = None
            date_groups.append({"fob_col": fob_col, "cm_col": cm_col, "margin_col": margin_col, "date": gdate})
            i = j if j > i + 1 else i + 1
        else:
            i += 1

    # REMARK / TOTAL FOB SAVING 컬럼
    remark_col = None
    for c, label in col_headers.items():
        if label == "REMARK":
            remark_col = c

    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        season = ws.cell(row=r, column=1).value
        style_cell = ws.cell(row=r, column=3).value
        description = ws.cell(row=r, column=4).value
        if not style_cell and not description:
            continue
        style_no, colorway = _parse_style_cell(style_cell)
        no_val = ws.cell(row=r, column=2).value
        remark_val = ws.cell(row=r, column=remark_col).value if remark_col else ""

        dates = {}
        for g in date_groups:
            fob = ws.cell(row=r, column=g["fob_col"]).value if g["fob_col"] else None
            cm = ws.cell(row=r, column=g["cm_col"]).value if g["cm_col"] else None
            margin = ws.cell(row=r, column=g["margin_col"]).value if g["margin_col"] else None
            if fob is None and cm is None:
                continue
            key_date = g["date"] or dt.date(1900, 1, 1) + dt.timedelta(days=len(dates))
            dates[key_date] = {
                "fob": fob if isinstance(fob, (int, float)) else None,
                "cm": cm if isinstance(cm, (int, float)) else None,
                "margin": margin if isinstance(margin, (int, float)) else None,
            }

        rows.append(
            SummaryRow(
                season=season or "",
                style_no=style_no,
                colorway=colorway,
                description=description or "",
                no=str(no_val) if no_val is not None else "",
                dates=dates,
                remark=remark_val or "",
            )
        )
    return rows


def merge_rows(existing: list[SummaryRow], new: list[SummaryRow]) -> list[SummaryRow]:
    """기존 행과 새로 파싱한 행을 style_no+colorway 기준으로 병합. 새 날짜는 추가,
    같은 날짜가 있으면 새 값으로 덮어쓴다."""
    by_key = {r.key: r for r in existing}
    ordered_keys = [r.key for r in existing]
    for nr in new:
        if nr.key in by_key:
            er = by_key[nr.key]
            er.dates.update(nr.dates)
            if nr.description:
                er.description = nr.description
            if nr.season:
                er.season = nr.season
            if nr.remark:
                er.remark = nr.remark
        else:
            by_key[nr.key] = nr
            ordered_keys.append(nr.key)
    return [by_key[k] for k in ordered_keys]


def _date_label(idx: int, total: int, d: dt.date) -> str:
    txt = f"{d.month}/{d.day}"
    if idx == 0:
        return f"SUBMITTED {txt}"
    if idx == total - 1:
        return f"NEGOTIATED {txt}"
    return f"FOB {txt}"


def build_summary_workbook(rows: list[SummaryRow], season_label: str, brand: str) -> openpyxl.Workbook:
    all_dates = sorted({d for r in rows for d in r.dates.keys()})
    n_dates = len(all_dates)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"

    # ---- 레이아웃 계산 ----
    # A SEASON, B NO, C STYLE, D DESCRIPTION, 이후 날짜그룹(FOB,CM,MARGIN%) x n, TOTAL FOB SAVING, REMARK
    fixed_cols = ["SEASON", "NO", "STYLE", "DESCRIPTION"]
    first_date_col = len(fixed_cols) + 1
    group_width = 3  # FOB, CM, MARGIN%
    total_saving_col = first_date_col + n_dates * group_width
    remark_col = total_saving_col + 1

    header_row = 2
    data_start_row = 3

    def style_header(cell, bold=True, color=None, fill=None, size=9, font_name=FONT_NAME):
        cell.font = Font(name=font_name, bold=bold, size=size, color=color)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
        if fill:
            cell.fill = fill

    # 고정 컬럼 헤더
    for i, name in enumerate(fixed_cols):
        c = ws.cell(row=header_row, column=i + 1, value=name)
        style_header(c, fill=HEADER_FILL)
        ws.merge_cells(start_row=header_row - 1, start_column=i + 1, end_row=header_row, end_column=i + 1)

    # 날짜 그룹 헤더
    for gi, d in enumerate(all_dates):
        base_col = first_date_col + gi * group_width
        label = _date_label(gi, n_dates, d)
        top = ws.cell(row=header_row - 1, column=base_col, value=label)
        top.font = Font(name="맑은 고딕", bold=True, size=11, color=BLUE)
        top.alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(start_row=header_row - 1, start_column=base_col, end_row=header_row - 1, end_column=base_col + 2)
        for k, sub in enumerate(["FOB", "CM", "MARGIN %"]):
            c = ws.cell(row=header_row, column=base_col + k, value=sub if sub != "FOB" else f"FOB {d.month}/{d.day}")
            style_header(c, color=BLUE if sub != "MARGIN %" else None, fill=HEADER_FILL)

    tsc = ws.cell(row=header_row, column=total_saving_col, value="TOTAL FOB SAVING")
    style_header(tsc, fill=HEADER_FILL)
    ws.merge_cells(start_row=header_row - 1, start_column=total_saving_col, end_row=header_row, end_column=total_saving_col)
    rc = ws.cell(row=header_row, column=remark_col, value="REMARK")
    style_header(rc, fill=HEADER_FILL)
    ws.merge_cells(start_row=header_row - 1, start_column=remark_col, end_row=header_row, end_column=remark_col)

    currency_fmt = '_(\\$* #,##0.00_);_(\\$* \\(#,##0.00\\);_(\\$* "-"??_);_(@_)'

    # ---- 데이터 행 ----
    r = data_start_row
    review_flags = []  # (row, col, message) - UI에 검토 필요 항목 표시용
    for row_data in rows:
        c_season = ws.cell(row=r, column=1, value=row_data.season)
        c_no = ws.cell(row=r, column=2, value=row_data.no)
        c_style = ws.cell(row=r, column=3, value=row_data.style_cell_text)
        c_desc = ws.cell(row=r, column=4, value=row_data.description)
        for cell in (c_season, c_no, c_style, c_desc):
            cell.font = Font(name=FONT_NAME, size=9)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BORDER

        first_fob_col_letter = None
        last_fob_col_letter = None
        for gi, d in enumerate(all_dates):
            base_col = first_date_col + gi * group_width
            info = row_data.dates.get(d, {})
            fob, cm, margin = info.get("fob"), info.get("cm"), info.get("margin")

            fob_cell = ws.cell(row=r, column=base_col, value=fob)
            fob_cell.number_format = currency_fmt
            fob_cell.font = Font(name=FONT_NAME, size=9, bold=True, color=BLUE)
            fob_cell.alignment = Alignment(horizontal="center", vertical="center")
            fob_cell.border = BORDER
            if fob is not None:
                col_letter = get_column_letter(base_col)
                last_fob_col_letter = f"{col_letter}{r}"
                if first_fob_col_letter is None:
                    first_fob_col_letter = f"{col_letter}{r}"

            cm_cell = ws.cell(row=r, column=base_col + 1, value=cm)
            cm_cell.font = Font(name=FONT_NAME, size=9)
            cm_cell.alignment = Alignment(horizontal="center", vertical="center")
            cm_cell.border = BORDER

            margin_cell = ws.cell(row=r, column=base_col + 2, value=margin)
            margin_cell.number_format = "0.00%"
            margin_cell.font = Font(name="맑은 고딕", size=9)
            margin_cell.alignment = Alignment(horizontal="center", vertical="center")
            margin_cell.border = BORDER
            if margin is None and fob is not None:
                review_flags.append((r, base_col + 2, f"{row_data.style_cell_text}: MARGIN% 수동 확인 필요"))

        saving_cell = ws.cell(row=r, column=total_saving_col)
        if first_fob_col_letter and last_fob_col_letter and first_fob_col_letter != last_fob_col_letter:
            saving_cell.value = f"={last_fob_col_letter}-{first_fob_col_letter}"
        saving_cell.number_format = currency_fmt
        saving_cell.font = Font(name=FONT_NAME, size=9, bold=True, color=BLUE)
        saving_cell.alignment = Alignment(horizontal="center", vertical="center")
        saving_cell.border = BORDER

        remark_cell = ws.cell(row=r, column=remark_col, value=row_data.remark)
        remark_cell.font = Font(name=FONT_NAME, size=9)
        remark_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        remark_cell.border = BORDER

        r += 1

    # ---- 열 너비 ----
    widths = {1: 9, 2: 7, 3: 16, 4: 24}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w
    for gi in range(n_dates):
        base_col = first_date_col + gi * group_width
        ws.column_dimensions[get_column_letter(base_col)].width = 13
        ws.column_dimensions[get_column_letter(base_col + 1)].width = 8
        ws.column_dimensions[get_column_letter(base_col + 2)].width = 11
    ws.column_dimensions[get_column_letter(total_saving_col)].width = 13
    ws.column_dimensions[get_column_letter(remark_col)].width = 38

    ws.freeze_panes = ws.cell(row=data_start_row, column=first_date_col)

    # ---- 비교 차트 (제출일이 3개 이상인 스타일만) ----
    chart_ws = None
    for row_idx, row_data in enumerate(rows):
        dated = sorted(row_data.dates.items())
        if len(dated) <= 2:
            continue
        if chart_ws is None:
            chart_ws = wb.create_sheet("Comparison Charts")
            chart_ws.sheet_view.showGridLines = False
            chart_ws["A1"] = "날짜별 FOB/CM 비교 차트 (제출 3회 이상 스타일)"
            chart_ws["A1"].font = Font(bold=True, size=12)

        data_row = 4 + row_idx * 12
        title = f"{row_data.style_no} / {row_data.colorway or '-'}  {row_data.description}"
        chart_ws.cell(row=data_row, column=1, value=title).font = Font(bold=True, size=10)
        chart_ws.cell(row=data_row + 1, column=1, value="DATE")
        chart_ws.cell(row=data_row + 1, column=2, value="FOB")
        chart_ws.cell(row=data_row + 1, column=3, value="CM")
        for k, (d, info) in enumerate(dated):
            chart_ws.cell(row=data_row + 2 + k, column=1, value=f"{d.month}/{d.day}")
            chart_ws.cell(row=data_row + 2 + k, column=2, value=info.get("fob"))
            chart_ws.cell(row=data_row + 2 + k, column=3, value=info.get("cm"))

        n = len(dated)
        chart = LineChart()
        chart.title = title
        chart.height, chart.width = 7, 14
        cats = Reference(chart_ws, min_col=1, min_row=data_row + 2, max_row=data_row + 1 + n)
        fob_ref = Reference(chart_ws, min_col=2, min_row=data_row + 1, max_row=data_row + 1 + n)
        cm_ref = Reference(chart_ws, min_col=3, min_row=data_row + 1, max_row=data_row + 1 + n)
        chart.add_data(fob_ref, titles_from_data=True)
        chart.add_data(cm_ref, titles_from_data=True)
        chart.set_categories(cats)
        chart.y_axis.title = "USD"
        chart.x_axis.title = "Submitted Date"
        anchor = f"E{data_row}"
        chart_ws.add_chart(chart, anchor)

    return wb, review_flags


def workbook_to_bytes(wb: openpyxl.Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
