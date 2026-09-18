"""parser.py 의 결과(스타일 정보 + 컬러웨이별 DateBlock 목록)를
summary_builder.SummaryRow 로 변환하는 연결 코드.

날짜 그대로 컬럼을 늘리면 컬럼 수가 너무 많아지고 빈 칸이 많아지므로, 시트 탭 이름에
적힌 샘플 단계(P1/P2/P3, GTM1/GTM2, SMS, PP)와 CBS 날짜의 '월'만 묶어서 하나의
컬럼(예: "PP Jul", "GTM2 Jun")으로 합친다. 같은 단계+월에 여러 번 제출된 경우
가장 최근(날짜가 늦은) 값으로 덮어써서 대표값 하나만 남긴다.
"""
import calendar

from summary_builder import SummaryRow
from parser import extract_stage


def stage_month_label(sheet_name: str, date) -> str:
    stage = extract_stage(sheet_name)
    mon = calendar.month_abbr[date.month]  # Jan, Feb, ... Jul, ...
    return f"{stage} {mon}"


def _compose_remark(blocks: list) -> tuple[str, bool]:
    """CM 옆 노란색 REMARK + 단계별 변경점을 하나의 REMARK 텍스트로 합친다.
    반환: (remark_text, needs_review)
    """
    dated = [b for b in blocks if b.date is not None and b.fob is not None]
    if not dated:
        dated = blocks

    needs_review = False
    lines = []

    latest = dated[-1]
    base_note = latest.remark.strip() if latest.remark else ""
    if base_note:
        lines.append(base_note)
        if not latest.remark_is_yellow:
            needs_review = True
            lines[-1] += "  [검토 필요: 노란색 강조 아님]"
    else:
        needs_review = True
        lines.append("[검토 필요: REMARK 미검출]")

    for b in dated:
        if not b.diffs:
            continue
        date_txt = f"{b.date.month}/{b.date.day}" if b.date else b.sheet_name
        # CM 변경을 우선 노출, 그 다음 항목 최대 2개까지만
        cm_diffs = [d for d in b.diffs if d.upper().startswith("CM")]
        other_diffs = [d for d in b.diffs if not d.upper().startswith("CM")]
        picked = (cm_diffs + other_diffs)[:3]
        for d in picked:
            lines.append(f"[{date_txt}] {d}")

    return "\n".join(lines), needs_review


def _bucket_by_stage_month(blocks: list) -> dict:
    """(단계+월) 라벨별로 가장 최근 제출 값 하나만 남긴다."""
    dated = [b for b in blocks if b.date is not None and b.fob is not None]
    buckets: dict[str, object] = {}
    for b in dated:
        label = stage_month_label(b.sheet_name, b.date)
        if label not in buckets or b.date > buckets[label].date:
            buckets[label] = b
    return {
        label: {"fob": b.fob, "cm": b.cm, "margin": b.margin, "date": b.date}
        for label, b in buckets.items()
    }


def submitted_to_summary_rows(style_info, blocks_by_colorway: dict, season_override: str = "") -> list:
    rows = []
    for colorway, blocks in blocks_by_colorway.items():
        dated_blocks = _bucket_by_stage_month(blocks)
        remark_text, needs_review = _compose_remark(blocks)
        row = SummaryRow(
            season=season_override or style_info.season,
            style_no=style_info.style_no,
            colorway=colorway,
            description=style_info.description,
            no="",
            dates=dated_blocks,
            remark=remark_text,
        )
        row.needs_review = needs_review  # type: ignore[attr-defined]
        rows.append(row)
    return rows
