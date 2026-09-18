"""
Costing Bulk SUMMARY 자동 생성 Streamlit 앱

- SUBMITTED로 끝나는 원가(costing) 엑셀 파일을 1개 이상 업로드하면
  스타일/시즌/DESCRIPTION, 날짜별 FOB·CM, REMARK를 자동으로 추출하고
  SUMMARY 차트(엑셀)를 만들어 준다.
- 기존 SUMMARY 파일을 함께 업로드하면 새 제출 데이터를 병합해서 업데이트한다.
- 시트가 3개 이상(제출 3회 이상)인 스타일은 날짜별 비교 표 + 꺾은선 차트를 자동 생성.
"""
import io
import datetime as dt

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
    "SUBMITTED 원가 엑셀에서 시즌·스타일·DESCRIPTION·날짜별 FOB/CM/REMARK를 추출해 "
    "SUMMARY 차트를 만들거나 업데이트합니다. 자동 추출 결과는 생성 전에 반드시 검토·수정할 수 있습니다."
)

if "parsed_rows" not in st.session_state:
    st.session_state.parsed_rows = []  # list[SummaryRow]
if "existing_rows" not in st.session_state:
    st.session_state.existing_rows = []

# ---------------- 사이드바: 파일명 옵션 ----------------
with st.sidebar:
    st.header("① 파일명 설정")
    season_label = st.text_input("시즌 (예: S27, F27, S27/F27)", value="S27")
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
            info, blocks = parse_submitted_workbook(io.BytesIO(f.getvalue()), filename_for_parsing=f.name)
            rows = submitted_to_summary_rows(info, blocks, season_override="")
            if not rows:
                warnings.append(f"⚠ {f.name}: 컬러웨이/CM 블록을 찾지 못했습니다. 파일 구조를 확인해주세요.")
            all_rows.extend(rows)
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
    st.caption("자동으로 채워지지 않은 항목(특히 MARGIN %, 스타일명/DESCRIPTION 오탐)은 표에서 직접 수정한 뒤 아래에서 SUMMARY를 생성하세요.")

    # ---- 스타일 마스터 테이블 (수정 가능) ----
    master_df = pd.DataFrame(
        [
            {
                "_idx": i,
                "SEASON": r.season,
                "NO": r.no,
                "STYLE NO": r.style_no,
                "COLORWAY": r.colorway,
                "DESCRIPTION": r.description,
                "REMARK": r.remark,
            }
            for i, r in enumerate(rows)
        ]
    )
    edited_master = st.data_editor(
        master_df,
        column_config={
            "_idx": None,
            "REMARK": st.column_config.TextColumn(width="large"),
        },
        num_rows="fixed",
        use_container_width=True,
        key="master_editor",
    )

    st.subheader("단계(STAGE) / 월별 FOB · CM · MARGIN %")
    st.caption(
        "정확한 날짜 대신 시트 탭 이름의 샘플 단계(P1/P2/P3, GTM1/GTM2, SMS, PP)와 "
        "제출월만 묶어서 하나의 컬럼으로 표시합니다 (예: 'PP Jul'). 같은 단계+월에 "
        "제출이 여러 번 있으면 가장 최근 값만 남깁니다. STAGE 라벨은 표에서 직접 수정할 수 있습니다. "
        "MARGIN %는 각 시트 하단(CBS와 같은 행, LOSS 열)의 값을 그대로 읽어온 것입니다."
    )
    date_rows = []
    for i, r in enumerate(rows):
        for label, info in sorted(r.dates.items(), key=lambda kv: kv[1].get("date") or dt.date.min):
            margin = info.get("margin")
            date_rows.append(
                {
                    "_idx": i,
                    "STYLE NO": r.style_no,
                    "COLORWAY": r.colorway,
                    "STAGE": label,
                    "FOB": info.get("fob"),
                    "CM": info.get("cm"),
                    # 표에서는 사람이 보기 편하게 %값(예: 16.12)으로 표시/입력하고,
                    # 최종 저장 시 소수(0.1612)로 다시 변환한다.
                    "MARGIN %": (margin * 100) if margin is not None else None,
                    # 정렬용 원본 날짜 (화면에는 안 보이지만 SUMMARY 생성 시 컬럼 순서를 정하는 데 씀)
                    "_date": info.get("date"),
                }
            )
    date_df = pd.DataFrame(date_rows)
    if not date_df.empty:
        edited_dates = st.data_editor(
            date_df,
            column_config={
                "_idx": None,
                "_date": None,
                "STAGE": st.column_config.TextColumn(help="예: 'PP Jul', 'GTM2 Jun', 'SMS Sep'. 잘못 인식됐으면 직접 수정하세요."),
                "MARGIN %": st.column_config.NumberColumn(
                    help="자동 추출된 값입니다. 비어 있거나 틀렸으면 직접 입력/수정하세요. (예: 16.12 = 16.12%)",
                    format="%.2f",
                ),
            },
            num_rows="dynamic",
            use_container_width=True,
            key="date_editor",
        )
    else:
        edited_dates = date_df

    st.header("⑤ SUMMARY 생성")
    if st.button("📊 SUMMARY 엑셀 생성", type="primary"):
        # 편집된 내용을 다시 SummaryRow 리스트로 조립
        final_rows = []
        for _, mrow in edited_master.iterrows():
            idx = mrow["_idx"]
            dsub = edited_dates[edited_dates["_idx"] == idx] if not edited_dates.empty else edited_dates
            dates = {}
            for _, drow in dsub.iterrows():
                label = str(drow["STAGE"]).strip()
                if not label or label.lower() == "nan":
                    continue
                raw_date = drow.get("_date")
                if raw_date is None or (not isinstance(raw_date, dt.date) and pd.isna(raw_date)):
                    raw_date = None
                elif isinstance(raw_date, pd.Timestamp):
                    raw_date = raw_date.date()
                elif not isinstance(raw_date, dt.date):
                    raw_date = None
                dates[label] = {
                    "fob": drow["FOB"] if pd.notna(drow["FOB"]) else None,
                    "cm": drow["CM"] if pd.notna(drow["CM"]) else None,
                    "margin": (drow["MARGIN %"] / 100.0) if pd.notna(drow["MARGIN %"]) else None,
                    "date": raw_date,
                }
            final_rows.append(
                SummaryRow(
                    season=mrow["SEASON"],
                    style_no=str(mrow["STYLE NO"]),
                    colorway=mrow["COLORWAY"],
                    description=mrow["DESCRIPTION"],
                    no=str(mrow["NO"]) if mrow["NO"] else "",
                    dates=dates,
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
- **시즌 / 스타일 / DESCRIPTION**: 파일명 시작 부분의 `S27-1234567_M FlyLux ...` 형식에서 자동 추출합니다.
- **제출일자 / FOB**: 각 시트의 `CBS <날짜> $<가격>` 문구를 찾아 사용합니다. (시트 탭 이름이 아닌 CBS 문구 기준)
- **컬럼 구성 (단계+월)**: 정확한 날짜별로 컬럼을 나누면 빈 칸/컬럼이 너무 많아지므로, 시트 탭 이름의
  샘플 단계(P1/P2/P3, GTM1/GTM2, SMS, PP)와 CBS 날짜의 월만 묶어 하나의 컬럼(예: `PP Jul`)으로
  만듭니다. 같은 단계+월에 여러 번 제출됐으면 가장 최근 값만 남습니다. 기존 SUMMARY 파일의
  `SUBMITTED 4/29`처럼 예전 방식으로 저장된 컬럼은 그대로 유지됩니다.
- **CM**: CM 행의 좌측(제안가 블록) TTL 값을 사용합니다.
- **REMARK**: CM 옆 우측 REMARK 칸 중 노란색으로 강조된 내용을 가져오고, 노란색이 아니면 `[검토 필요]`로 표시됩니다.
- **단계별 변경점**: 시트가 여러 개인 경우, 노란색으로 강조된 값이 이전 단계 대비 달라진 항목을 자동으로 REMARK에 함께 기록합니다. (참고용이며 확인이 필요합니다)
- **MARGIN %**: CBS 문구와 같은 행(그 블록의 최종 합계 행), LOSS 열에 적힌 값을 그대로 가져옵니다. SOLID/HEATHER가 한 시트에 같이 있는 경우(주로 파일명에 FlyLux가 들어간 스타일) 컬러웨이별로 서로 다른 값을 각각 읽어옵니다. 찾지 못하면 비워두고 검토 표시를 남깁니다.
- **비교 차트**: 제출 이력이 3회 이상인 스타일은 'Comparison Charts' 시트에 날짜별 FOB/CM 꺾은선 차트가 자동 추가됩니다.
        """
    )
