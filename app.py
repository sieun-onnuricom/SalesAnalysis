# -*- coding: utf-8 -*-
"""
app.py
======
데이터 입력 화면(Streamlit). 한 화면에서 입력 -> 보고서 2종 생성/다운로드.

입력
----
- xlsx 업로드(매출_raw / 일간데이터)
- 관리한계 배수 K 슬라이더(낮을수록 이상치를 더 많이 잡음)

출력
----
- 결과물1 보고용  : report_business.build_business_report
- 결과물2 검토용  : report_method.build_method_report

실행:  streamlit run app.py
"""
from __future__ import annotations

import os
import tempfile

import pandas as pd
import streamlit as st

import sales_anomaly_engine as E
import report_business as RB
import report_method as RM


st.set_page_config(page_title="매출 이상치 탐지", layout="wide")
st.title("매출 예측 · 이상치 탐지 엔진")
st.caption(f"대상: {E.TARGET_BRAND} {E.TARGET_PRODUCT_DAILY}  ·  "
           "베이스라인 OLS(추세+요일+프로모션+광고비) · 예측 ±Kσ 관리한계")


# ─────────────────────────────────────────────────────────────────────────
# 보조
# ─────────────────────────────────────────────────────────────────────────
def _save_upload(uploaded) -> str:
    """업로드(BytesIO)를 임시파일로 저장하고 경로 반환."""
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    with os.fdopen(fd, "wb") as f:
        f.write(uploaded.getbuffer())
    return path


def _sheet_arg(s: str):
    """시트 인자: 숫자면 인덱스(int), 아니면 시트명(str)."""
    s = (s or "").strip()
    return int(s) if s.isdigit() else (s or 0)


# ─────────────────────────────────────────────────────────────────────────
# 사이드바: 입력
# ─────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("1) 데이터 입력")
    same_file = st.checkbox("매출_raw와 일간데이터가 같은 파일", value=False,
                            help="한 워크북에 두 시트가 모두 있으면 체크")

    sales_path = daily_path = None
    up_sales = st.file_uploader("매출_raw 파일(.xlsx)", type=["xlsx"])
    up_daily = (up_sales if same_file
                else st.file_uploader("일간데이터 파일(.xlsx)", type=["xlsx"]))
    if up_sales:
        sales_path = _save_upload(up_sales)
    if up_daily:
        daily_path = sales_path if (same_file and sales_path) else _save_upload(up_daily)

    st.header("2) 시트 이름")
    sales_sheet = st.text_input("매출_raw 시트", value="매출_raw")
    daily_sheet = st.text_input("일간데이터 시트(이름 또는 0=첫시트)", value="0")

    st.header("3) 관리한계 K (±Kσ)")
    sigma_k = st.slider("K (낮을수록 이상치를 더 많이 잡음)",
                        min_value=1.0, max_value=3.0, value=1.5, step=0.1)
    st.caption("정규가정 기준 이탈률 ≈ ±1.0σ:32% · ±1.5σ:13% · ±2.0σ:5% · ±2.5σ:1%. "
               "실제 검출 수는 데이터에 따라 다름 → 결과의 '시그마민감도' 시트 확인.")

    run = st.button("분석 실행", type="primary", use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────
# 실행
# ─────────────────────────────────────────────────────────────────────────
if run:
    try:
        if not (sales_path and daily_path):
            st.error("매출_raw와 일간데이터를 모두 입력하세요.")
            st.stop()

        with st.spinner("분석 중… (적합 → 이상치 탐지 → 진단)"):
            result = E.run_analysis(
                sales_path, daily_path,
                sales_sheet=_sheet_arg(sales_sheet),
                daily_sheet=_sheet_arg(daily_sheet),
                sigma_k=sigma_k,
            )
            biz_path = os.path.join(tempfile.gettempdir(), "보고용_리포트.xlsx")
            mth_path = os.path.join(tempfile.gettempdir(), "검토용_리포트.xlsx")
            RB.build_business_report(result, biz_path)
            RM.build_method_report(result, mth_path)

        an = result["anomalies"]
        n_anom = int(an["이상치"].sum())
        st.success(f"완료 · 분석 {an['날짜'].nunique()}일 · "
                   f"R²={result['ols'].rsquared:.3f} · σ={result['sigma']:.4f} · "
                   f"이상일 {n_anom}일 (±{sigma_k:.1f}σ)")

        # ── 다운로드 ──
        c1, c2 = st.columns(2)
        with c1:
            with open(biz_path, "rb") as f:
                st.download_button("⬇ 결과물1 · 보고용 리포트(xlsx)", f.read(),
                                   file_name="보고용_리포트.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
        with c2:
            with open(mth_path, "rb") as f:
                st.download_button("⬇ 결과물2 · 검토용 리포트(xlsx)", f.read(),
                                   file_name="검토용_리포트.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)

        # ── 화면 미리보기 ──
        st.subheader("요약")
        st.dataframe(
            pd.read_excel(biz_path, sheet_name="00_요약"),
            use_container_width=True, hide_index=True)

        col_a, col_b = st.columns(2)
        with col_a:
            st.subheader("K별 이상일 수 (K 선택 근거)")
            ss = pd.read_excel(mth_path, sheet_name="03_시그마민감도")
            st.dataframe(ss.dropna(subset=["SIGMA_K"]) if "SIGMA_K" in ss.columns else ss,
                         use_container_width=True, hide_index=True)
        with col_b:
            st.subheader("결론 및 유의점(자동)")
            cc = pd.read_excel(mth_path, sheet_name="07_결론_및_유의점")
            st.dataframe(cc.dropna(subset=["점검항목"]) if "점검항목" in cc.columns else cc,
                         use_container_width=True, hide_index=True)

        st.subheader("이상일 요약")
        st.dataframe(pd.read_excel(biz_path, sheet_name="04_이상일_요약"),
                     use_container_width=True, hide_index=True)
        st.info("시각화(그라데이션 밴드 차트)는 보고용 리포트의 '03_매출_밴드차트' 시트에서 확인하세요.")

    except Exception as e:
        st.error(f"오류: {e}")
        st.exception(e)
else:
    st.info("좌측에서 매출_raw와 일간데이터 xlsx를 올리고 [분석 실행]을 누르세요.")


# ─────────────────────────────────────────────────────────────────────────
# 푸터(모든 화면 하단에 표시)
# ─────────────────────────────────────────────────────────────────────────
st.divider()
st.markdown(
    "<div style='text-align:center; color:#888; font-size:0.85rem; padding:6px 0;'>"
    "온누리커뮤니티 · 박시은 · 2026-06-05"
    "</div>",
    unsafe_allow_html=True,
)
