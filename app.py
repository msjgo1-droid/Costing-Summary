"""
Costing Bulk SUMMARY 자동 생성 Streamlit 앱

- SUBMITTED로 끝나는 원가(costing) 엑셀 파일을 1개 이상 업로드하면
  스타일/시즌/DESCRIPTION과, 컬러웨이별 가장 최근 제출분의 FOB·CM·MARGIN%·REMARK를
  자동으로 추출해 SUMMARY 엑셀을 만들어 준다.
- 기존 SUMMARY 파일을 함께 업로드하면 새 제출 데이터로 값을 덮어써서 업데이트한다.
"""
import io

import pandas as pd
import streamlit as st

from parser import parse_submitted_workbook
from pipeline import submitted_to_summary_rows
from summary_builder import (
    SummaryRow,
    build_summary_workbook,
    workbook_to_bytes,
    merge_rows,
    parse_existing_summary,
)

st.set_page_config(page_title="Costing Bulk SUMMARY 생성기", layout="wide")

st.title("📦 Costing Bulk SUMMARY 자동 생성기")
st.caption(
    "SUBMITTED 원가 엑셀에서 시즌·스타일·DESCRIPTION과 가장 최근 제출분의 FOB/CM/MARGIN%/REMARK를 "
    "추출해 SUMMARY 파일을 만들거나 업데이트합니다. 자동 추출 결과는 생성 전에 반드시 검토·수정할 수 있습니다."
)

if "parsed_rows" not in st.session_state:
    st.session_state.parsed_rows = []  # list[SummaryRow]
if "existing_rows" not in st.session_state:
    st.session_state.existing_rows = []

# ---------------- 사이드바: 파일명 옵션 ----------------
with st.sidebar:
    st.header("① 파일명 설정")
    season_label = st.text_input("시즌 (예: S27, F27, S28)", value="F27")
    brand = st.text_input("브랜드/바이어명", value="HOKA")
    suffix = st.text_input("파일명 뒤에 붙일 설명 (선택)", value="", placeholder="예: Negotiated 06.01")
    filename_preview = f"Costing Bulk {season_label} {brand} - SUMMARY"
    if suffix:
        filename_preview += f" {suffix}"
    st.text_input("생성될 파일명 미리보기", value=filename_preview + ".xlsx", disabled=True)

st.header("② 파일 업로드")
col1, col2 = st.columns(2)
with col1:
    submitted_files = st.file_uploader(
        "SUBMITTED 원가 엑셀 (여러 개 선택 가능)",
        type=["xlsx"],
        accept_multiple_files=True,
        help="파일명이 ...-SUBMITTED.xlsx 로 끝나는 스타일별 원가 시트",
    )
with col2:
    existing_summary_file = st.file_uploader(
        "업데이트할 기존 SUMMARY 파일 (선택 사항)",
        type=["xlsx"],
        help="비워두면 새 SUMMARY를 처음부터 생성합니다.",
    )

st.header("③ 데이터 추출")
if st.button("📥 업로드한 파일에서 데이터 추출", type="primary", disabled=not submitted_files):
    all_rows = []
    warnings = []
    for f in submitted_files:
        try:
            info, blocks, skip_notes = parse_submitted_workbook(io.BytesIO(f.getvalue()), filename_for_parsing=f.name)
            rows = submitted_to_summary_rows(info, blocks, season_override="")
            if not rows:
                warnings.append(f"⚠ {f.name}: 컬러웨이/CM 블록을 찾지 못했습니다. 파일 구조를 확인해주세요.")
            all_rows.extend(rows)
            for note in skip_notes:
                warnings.append(f"⚠ {f.name}: {note}")
        except Exception as e:
            warnings.append(f"⚠ {f.name}: 처리 중 오류 발생 ({e})")

    existing_rows = []
    if existing_summary_file is not None:
        try:
            existing_rows = parse_existing_summary(io.BytesIO(existing_summary_file.getvalue()))
        except Exception as e:
            warnings.append(f"⚠ 기존 SUMMARY 파일을 읽는 중 오류 발생 ({e})")

    st.session_state.existing_rows = existing_rows
    st.session_state.parsed_rows = merge_rows(existing_rows, all_rows)
    for w in warnings:
        st.warning(w)
    st.success(f"{len(all_rows)}개 컬러웨이 행을 추출했습니다. (기존 SUMMARY 행 {len(existing_rows)}개와 병합됨)")

rows: list[SummaryRow] = st.session_state.parsed_rows

if rows:
    st.header("④ 추출 결과 확인 및 수정")
    st.caption(
        "스타일/컬러웨이별로 가장 최근 제출분의 FOB·CM·MARGIN%만 한 줄로 표시합니다. "
        "STAGE 칸은 그 값이 어느 단계(P1/PP/SMS 등)에서 나온 것인지 참고용입니다. "
        "자동으로 채워지지 않았거나 틀린 값은 표에서 직접 수정한 뒤 아래에서 SUMMARY를 생성하세요."
    )

    master_df = pd.DataFrame(
        [
            {
                "_idx": i,
                "SEASON": r.season,
                "STYLE NO": r.style_no,
                "COLORWAY": r.colorway,
                "DESCRIPTION": r.description,
                "STAGE": r.stage,
                "FOB": r.fob,
                "CM": r.cm,
                # 표에서는 사람이 보기 편하게 %값(예: 16.12)으로 표시/입력하고,
                # 최종 저장 시 소수(0.1612)로 다시 변환한다.
                "MARGIN %": (r.margin * 100) if r.margin is not None else None,
                "REMARK": r.remark,
            }
            for i, r in enumerate(rows)
        ]
    )
    edited_master = st.data_editor(
        master_df,
        column_config={
            "_idx": None,
            "STAGE": st.column_config.TextColumn(
                help="이 FOB/CM/MARGIN% 값이 어느 단계(P1/P2/P3, GTM1/GTM2, SMS, PP)에서 나온 것인지 참고용 표시입니다. 잘못 인식됐으면 직접 수정하세요."
            ),
            "MARGIN %": st.column_config.NumberColumn(
                help="자동 추출된 값입니다. 비어 있거나 틀렸으면 직접 입력/수정하세요. (예: 16.12 = 16.12%)",
                format="%.2f",
            ),
            "REMARK": st.column_config.TextColumn(width="large"),
        },
        num_rows="fixed",
        use_container_width=True,
        key="master_editor",
    )

    st.header("⑤ SUMMARY 생성")
    if st.button("📊 SUMMARY 엑셀 생성", type="primary"):
        final_rows = []
        for _, mrow in edited_master.iterrows():
            final_rows.append(
                SummaryRow(
                    season=mrow["SEASON"],
                    style_no=str(mrow["STYLE NO"]),
                    colorway=mrow["COLORWAY"],
                    description=mrow["DESCRIPTION"],
                    stage=mrow["STAGE"] or "",
                    fob=mrow["FOB"] if pd.notna(mrow["FOB"]) else None,
                    cm=mrow["CM"] if pd.notna(mrow["CM"]) else None,
                    margin=(mrow["MARGIN %"] / 100.0) if pd.notna(mrow["MARGIN %"]) else None,
                    remark=mrow["REMARK"] or "",
                )
            )

        wb, review_flags = build_summary_workbook(final_rows, season_label=season_label, brand=brand)
        data = workbook_to_bytes(wb)

        final_filename = f"Costing Bulk {season_label} {brand} - SUMMARY"
        if suffix:
            final_filename += f" {suffix}"
        final_filename += ".xlsx"

        st.success("SUMMARY 파일이 생성되었습니다.")
        if review_flags:
            with st.expander(f"⚠ 검토가 필요한 항목 {len(review_flags)}건 (주로 MARGIN %)", expanded=False):
                for _, _, msg in review_flags:
                    st.write("- " + msg)

        st.download_button(
            "⬇️ SUMMARY 엑셀 다운로드",
            data=data,
            file_name=final_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("먼저 SUBMITTED 파일을 업로드하고 '데이터 추출' 버튼을 눌러주세요.")

with st.expander("ℹ️ 추출 규칙 안내"):
    st.markdown(
        """
- **시즌 / 스타일 / DESCRIPTION**: 파일명 시작 부분의 `F27-1234567_M FlyLux ...` 형식에서 자동 추출합니다.
- **가장 최근 값만 표시**: 여러 번 제출된 이력이 있어도 컬럼이 늘어나지 않도록, 스타일/컬러웨이별로
  가장 최근(제출일이 가장 늦은, 동점이면 시트 탭이 더 왼쪽/최신인 쪽) 제출분의 FOB/CM/MARGIN%만
  한 줄로 보여줍니다. STAGE 칸에는 그 값이 어느 단계(P1/P2/P3, GTM1/GTM2, SMS, PP)에서 나온
  것인지 참고용으로 표시됩니다.
- **CM**: CM 행의 우측("open"/개정 CM) TTL 값을 사용합니다. (우측 값이 없으면 좌측 값으로 대체)
- **REMARK**: CM 옆 우측 REMARK 칸 중 노란색으로 강조된 내용을 가져오고, 노란색이 아니면 `[검토 필요]`로 표시됩니다.
  이전 단계 대비 변경점(노란색 강조 항목 기준)도 참고용으로 함께 기록됩니다.
- **MARGIN %**: CBS 문구와 같은 행(그 블록의 최종 합계 행), LOSS 열에 적힌 값을 그대로 가져옵니다.
  SOLID/HEATHER가 한 시트에 같이 있는 경우(주로 파일명에 FlyLux가 들어간 스타일) 컬러웨이별로
  서로 다른 값을 각각 읽어옵니다. 찾지 못하면 비워두고 검토 표시를 남깁니다.
- **시즌 표기 제외**: 시트 탭 이름에서 단계(P1~PP 등)를 못 찾을 경우, S27/F27 같은 시즌 표기는
  단계 라벨로 쓰지 않습니다.
- **스타일번호 불일치 차단**: 시트 내용에 적힌 스타일번호가 파일명의 스타일번호와 다르면 데이터
  품질 문제로 보고 해당 블록은 무시합니다 (경고 메시지로 안내됩니다).
        """
    )
