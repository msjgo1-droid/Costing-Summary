"""parser.py 의 결과(스타일 정보 + 컬러웨이별 DateBlock 목록)를
summary_builder.SummaryRow 로 변환하는 연결 코드.

날짜/단계별로 컬럼을 늘리면 컬럼 수가 너무 많아지고 빈 칸이 많아지므로, 컬러웨이별로
가장 최근(날짜가 가장 늦은, 동점이면 시트 탭이 더 왼쪽/최신인 쪽) 제출분 딱 하나만
SUMMARY의 한 줄에 반영한다. CM은 우측(open/개정) 값을 사용한다.
"""
import datetime as dt

from summary_builder import SummaryRow
from parser import extract_stage


def _pick_latest(blocks: list):
    """컬러웨이의 여러 블록 중 SUMMARY에 반영할 가장 최근 스냅샷 하나를 고른다.
    날짜가 있는 블록을 우선하고, 없으면 전체 블록 중에서 고른다. 동점이면
    시트 탭 순서(sheet_order, 값이 클수록 더 왼쪽/최신)가 더 큰 쪽을 선택한다."""
    dated = [b for b in blocks if b.date is not None and b.fob is not None]
    candidates = dated if dated else blocks
    if not candidates:
        return None
    return max(candidates, key=lambda b: (b.date or dt.date.min, b.sheet_order))


def _compose_remark(blocks: list) -> tuple[str, bool]:
    """CM 옆 노란색 REMARK + 단계별 변경점을 하나의 REMARK 텍스트로 합친다.
    반환: (remark_text, needs_review)
    """
    dated = [b for b in blocks if b.date is not None and b.fob is not None]
    if not dated:
        dated = blocks
    if not dated:
        return "", False

    needs_review = False
    lines = []

    latest = dated[-1]
    base_note = latest.remark.strip() if latest.remark else ""
    if base_note:
        lines.append(base_note)
        # 노란색 강조가 아니어도 REMARK 텍스트 자체는 그대로 사용하고,
        # 화면에 노출되는 문구는 남기지 않는다 (review_flags로만 내부 추적).
        if not latest.remark_is_yellow:
            needs_review = True
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
        latest = _pick_latest(blocks)
        remark_text, needs_review = _compose_remark(blocks)

        stage_label = extract_stage(latest.sheet_name) if latest else ""
        fob = latest.fob if latest else None
        # 우측("open"/개정) CM을 우선 사용하고, 없으면 좌측 값으로 보완
        cm = None
        margin = None
        if latest is not None:
            cm = latest.cm_right if latest.cm_right is not None else latest.cm
            margin = latest.margin

        row = SummaryRow(
            season=season_override or style_info.season,
            style_no=style_info.style_no,
            colorway=colorway,
            description=style_info.description,
            stage=stage_label,
            fob=fob,
            cm=cm,
            margin=margin,
            remark=remark_text,
        )
        row.needs_review = needs_review  # type: ignore[attr-defined]
        rows.append(row)
    return rows
