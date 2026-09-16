"""parser.py 의 결과(스타일 정보 + 컬러웨이별 DateBlock 목록)를
summary_builder.SummaryRow 로 변환하는 연결 코드.
"""
from summary_builder import SummaryRow


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


def submitted_to_summary_rows(style_info, blocks_by_colorway: dict, season_override: str = "") -> list:
    rows = []
    for colorway, blocks in blocks_by_colorway.items():
        dated_blocks = {b.date: {"fob": b.fob, "cm": b.cm, "margin": b.margin} for b in blocks if b.date is not None and b.fob is not None}
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
