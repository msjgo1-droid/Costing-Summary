"""
SUMMARY 차트(엑셀)를 새로 만들거나 기존 SUMMARY 파일에 새 제출 데이터를 병합해서
다시 만들어주는 모듈.

컬럼 구성 (스타일/컬러웨이별로 가장 최근 제출분 한 줄만 표시):
  Season | Style/Colour | Style Description | Stage | FOB | CM | MARGIN % | REMARK
"""
import io
from dataclasses import dataclass
from typing import Optional

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

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
    stage: str = ""  # 참고용: 이 FOB/CM/MARGIN%가 어느 단계(P1/PP/SMS 등)에서 나온 값인지
    fob: Optional[float] = None
    cm: Optional[float] = None  # 우측("open"/개정) CM 값
    margin: Optional[float] = None
    remark: str = ""

    @property
    def key(self):
        return (self.style_no, self.colorway)

    @property
    def style_cell_text(self):
        return f"{self.style_no} / {self.colorway}" if self.colorway else self.style_no


def _parse_style_cell(text: str):
    if text is None:
        return "", ""
    text = str(text).strip()
    if "/" in text:
        style_no, colorway = text.split("/", 1)
        return style_no.strip(), colorway.strip()
    return text, ""


def parse_existing_summary(file_like) -> list[SummaryRow]:
    """기존 SUMMARY 엑셀에서 행 데이터를 읽어 SummaryRow 리스트로 반환.

    예전 방식(날짜/단계별로 FOB·CM·MARGIN% 그룹이 여러 개 있던 파일)과 새 방식
    (스타일당 한 줄, FOB·CM·MARGIN% 한 세트)을 모두 인식한다. 예전 방식 파일이면
    값이 채워진 여러 그룹 중 가장 마지막(=가장 최근) 그룹의 값을 대표값으로 사용한다."""
    wb = openpyxl.load_workbook(file_like, data_only=True)
    ws = wb[wb.sheetnames[0]]

    # 헤더 행 찾기 (SEASON 이 적힌 행)
    header_row = None
    for r in range(1, min(ws.max_row, 10) + 1):
        for c in range(1, 6):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip().upper() == "SEASON":
                header_row = r
                break
        if header_row:
            break
    if header_row is None:
        return []

    col_headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if isinstance(v, str):
            col_headers[c] = v.strip().upper()

    season_col = next((c for c, v in col_headers.items() if v == "SEASON"), 1)
    style_col = next((c for c, v in col_headers.items() if v in ("STYLE", "STYLE/COLOUR", "STYLE/COLOR")), None)
    desc_col = next((c for c, v in col_headers.items() if v in ("DESCRIPTION", "STYLE DESCRIPTION")), None)
    stage_col = next((c for c, v in col_headers.items() if v == "STAGE"), None)
    remark_col = next((c for c, v in col_headers.items() if v == "REMARK"), None)

    if style_col is None or desc_col is None:
        return []

    # FOB 컬럼(들) 찾기: 새 방식이면 1개, 예전 방식이면 여러 개(날짜/단계 그룹별로 하나씩)
    fob_cols = sorted(c for c, v in col_headers.items() if v.startswith("FOB"))

    def group_for(fob_col):
        cm_col = fob_col + 1 if col_headers.get(fob_col + 1) == "CM" else None
        margin_col = fob_col + 2 if cm_col and "MARGIN" in col_headers.get(fob_col + 2, "") else None
        # 예전 방식은 그룹 라벨(예: "PP Jul", "SUBMITTED 4/29")이 헤더 한 행 위 병합 셀에 있었다.
        label = str(ws.cell(row=header_row - 1, column=fob_col).value or "").strip()
        return {"fob_col": fob_col, "cm_col": cm_col, "margin_col": margin_col, "label": label}

    groups = [group_for(c) for c in fob_cols]

    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        style_cell = ws.cell(row=r, column=style_col).value
        description = ws.cell(row=r, column=desc_col).value
        if not style_cell and not description:
            continue
        style_no, colorway = _parse_style_cell(style_cell)
        season = ws.cell(row=r, column=season_col).value
        stage_val = ws.cell(row=r, column=stage_col).value if stage_col else ""
        remark_val = ws.cell(row=r, column=remark_col).value if remark_col else ""

        stage_label = str(stage_val).strip() if stage_val else ""
        fob = cm = margin = None
        # 값이 채워진 그룹을 순서대로 덮어써서, 가장 마지막(=가장 최근)에 값이 있는
        # 그룹의 데이터가 최종 대표값으로 남도록 한다.
        for g in groups:
            gf = ws.cell(row=r, column=g["fob_col"]).value if g["fob_col"] else None
            gc = ws.cell(row=r, column=g["cm_col"]).value if g["cm_col"] else None
            gm = ws.cell(row=r, column=g["margin_col"]).value if g["margin_col"] else None
            if gf is None and gc is None:
                continue
            if isinstance(gf, (int, float)):
                fob = gf
            if isinstance(gc, (int, float)):
                cm = gc
            if isinstance(gm, (int, float)):
                margin = gm
            if not stage_label and g["label"]:
                stage_label = g["label"]

        rows.append(
            SummaryRow(
                season=season or "",
                style_no=style_no,
                colorway=colorway,
                description=description or "",
                stage=stage_label,
                fob=fob,
                cm=cm,
                margin=margin,
                remark=remark_val or "",
            )
        )
    return rows


def merge_rows(existing: list[SummaryRow], new: list[SummaryRow]) -> list[SummaryRow]:
    """기존 행과 새로 파싱한 행을 style_no+colorway 기준으로 병합.
    같은 스타일/컬러웨이가 이미 있으면 새로 추출한(더 최근 제출분) 값으로 덮어쓴다."""
    by_key = {r.key: r for r in existing}
    ordered_keys = [r.key for r in existing]
    for nr in new:
        if nr.key in by_key:
            er = by_key[nr.key]
            if nr.fob is not None:
                er.fob = nr.fob
            if nr.cm is not None:
                er.cm = nr.cm
            if nr.margin is not None:
                er.margin = nr.margin
            if nr.stage:
                er.stage = nr.stage
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


def build_summary_workbook(rows: list[SummaryRow], season_label: str, brand: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"

    headers = ["Season", "Style/Colour", "Style Description", "Stage", "FOB", "CM", "MARGIN %", "REMARK"]
    header_row = 1
    data_start_row = 2

    def style_header(cell):
        cell.font = Font(name=FONT_NAME, bold=True, size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
        cell.fill = HEADER_FILL

    for i, name in enumerate(headers):
        c = ws.cell(row=header_row, column=i + 1, value=name)
        style_header(c)

    currency_fmt = '_(\\$* #,##0.00_);_(\\$* \\(#,##0.00\\);_(\\$* "-"??_);_(@_)'

    r = data_start_row
    review_flags = []  # (row, col, message) - UI에 검토 필요 항목 표시용
    for row_data in rows:
        c_season = ws.cell(row=r, column=1, value=row_data.season)
        c_style = ws.cell(row=r, column=2, value=row_data.style_cell_text)
        c_desc = ws.cell(row=r, column=3, value=row_data.description)
        c_stage = ws.cell(row=r, column=4, value=row_data.stage)
        for cell in (c_season, c_style, c_desc, c_stage):
            cell.font = Font(name=FONT_NAME, size=9)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BORDER

        fob_cell = ws.cell(row=r, column=5, value=row_data.fob)
        fob_cell.number_format = currency_fmt
        fob_cell.font = Font(name=FONT_NAME, size=9, bold=True, color=BLUE)
        fob_cell.alignment = Alignment(horizontal="center", vertical="center")
        fob_cell.border = BORDER

        cm_cell = ws.cell(row=r, column=6, value=row_data.cm)
        cm_cell.font = Font(name=FONT_NAME, size=9)
        cm_cell.alignment = Alignment(horizontal="center", vertical="center")
        cm_cell.border = BORDER

        margin_cell = ws.cell(row=r, column=7, value=row_data.margin)
        margin_cell.number_format = "0.00%"
        margin_cell.font = Font(name="맑은 고딕", size=9)
        margin_cell.alignment = Alignment(horizontal="center", vertical="center")
        margin_cell.border = BORDER
        if row_data.margin is None and row_data.fob is not None:
            review_flags.append((r, 7, f"{row_data.style_cell_text}: MARGIN% 수동 확인 필요"))

        remark_cell = ws.cell(row=r, column=8, value=row_data.remark)
        remark_cell.font = Font(name=FONT_NAME, size=9)
        remark_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        remark_cell.border = BORDER

        r += 1

    widths = {1: 9, 2: 18, 3: 26, 4: 10, 5: 13, 6: 9, 7: 11, 8: 42}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.freeze_panes = ws.cell(row=data_start_row, column=5)

    return wb, review_flags


def workbook_to_bytes(wb: openpyxl.Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
