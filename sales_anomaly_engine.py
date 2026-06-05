# -*- coding: utf-8 -*-
"""
sales_anomaly_engine.py
=======================
매출 예측 -> 예측구간(±2σ) 이탈일(이상치) 탐지 -> 이상일 마케팅 활동 매핑 엔진.

설계 원칙
---------
- 백엔드 통계 엔진 전용. Streamlit 등 UI 의존성 없음(main_ui()는 별도 파일).
- 헤드리스 단독 실행 가능. 다중시트 엑셀 리포트 산출.
- 베이스라인: 단일 OLS (추세 + 요일 더미 + 프로모션 더미 + 광고비).
  * 평균 재합산 안 함. 추세선 절편에 레벨 포함(이중계산 방지).
- 이상치 기준: 예측 ±SIGMA_K·σ 관리한계 (SIGMA_K=2.0). σ는 OLS 잔차 표준편차.
- 프로모션: 이상치 처리 X -> 회귀 더미로 투입(계획된 이벤트). 광고비(연속)와 별개 축.
- LightGBM+SHAP: 매체 기여도 레이어 '선택 모듈'. 미설치 시 자동 비활성.
- 분석 대상: 란시노 유리젖병. SKU 단위까지 분해.

데이터 소스 (확정 사양)
----------------------
[매출_raw] 헤더 1행, 거래 단위.
  A=결제일자 B=브랜드 C=제품 D=SKU E=SKU판매가 F=원가 G=실결제금액 H=수량 I=계정별칭
  -> 매출(Y)=G합계, 수량=H합계.
[일간데이터] 헤더 2행(1행 병합 그룹 / 2행 세부). 조인키 A·B·C.
  J=광고비(paid). ※ 사용자 확정: 메타 외 유료채널 존재 -> J를 '총광고비'로 사용.
  마케팅 활동 매핑용 활동열: Q,R(마케팅) / S,T,V(바이럴) / W,X,Z(숏폼) / AA,AB,AD(사이다)
  결과/누수열 제외: D,E(전체) H,I(paid 전환) K,L,M,N(paid제외) O,P(마케팅매출)
                    U,Y,AC(채널 판매수량) G(ROAS) AE(오가닉판매수량, 데이터스튜디오용)

조인키 = (날짜, 브랜드, 제품). 대상 분석은 브랜드=란시노, 제품=유리젖병 고정.
유리젖병 매칭(매출_raw 제품열): '젖병' 포함 AND 'PP' 미포함.
  -> 의심 매칭건(부속/세정류 추정)은 별도 리스팅.
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass, field
from datetime import date, datetime

import re

import numpy as np
import pandas as pd
import statsmodels.api as sm


# ============================================================================
# 설정
# ============================================================================
SIGMA_K = 2.0          # 관리한계 배수 (±2σ)
MIN_DAYS_FIT = 30      # OLS 적합 최소 일수
MIN_DAYS_CV = 60       # 교차검증 최소 일수

TARGET_BRAND = "란시노"
TARGET_PRODUCT_DAILY = "유리젖병"   # 일간데이터 제품열의 카테고리 값

# 베이스라인 단순 유지: 교차항/SHAP은 기본 OFF
USE_CROSS_PROMO_ADSPEND = False     # 광고비×프로모션 교차항
ENABLE_SHAP = False                 # LightGBM+SHAP 매체 기여도 레이어

# 자기상관/이분산 보정 (잔차진단 결과 반영)
# ※ lag1 제거 요청 반영: USE_LAG1=False. 단점은 자기상관(DW) 재발 가능 -> 11_잔차진단 재확인 필요.
#   대신 광고비-전일매출 다중공선성이 사라져 광고비 계수 해석이 깨끗해짐(21_VIF 참고).
USE_LAG1 = False           # 전일 매출(lag-1) 항. False=제거
USE_LOG_TARGET = True      # 로그매출로 적합 -> 이분산 완화(관리폭이 매출수준에 비례)
HAC_MAXLAGS = 7            # Newey-West(HAC) 강건 표준오차 lag(주간). 계수 p값 보정용
SPIKE_PCT = 30.0           # 예측이 전일 대비 ±이 %p 이상 튄 날을 '예측 급변일'로 진단
USE_WEEKDAY_SIGMA = True   # ±2σ 관리한계를 요일별 σ로 (요일별 변동성 차이 반영 -> 과검출 교정)
WEEKDAY_SIGMA_MIN_N = 10   # 요일별 σ 산출 최소 표본(미만이면 전체 σ 대체)

# 매출_raw 유리젖병 매칭 시 '의심'으로 분류할 부속/비매출 키워드(검토 표시용, 제외 아님)
ACCESSORY_KEYWORDS = ["젖꼭지", "꼭지", "세정", "세척", "소독", "솔", "브러시",
                      "케이스", "커버", "거치", "건조", "보관", "리필", "패드"]

# 수기 검토 후 유리젖병에서 '제외'할 제품명(완전일치). 규칙으로 못 거르는 혼합박스 등.
# (내추럴웨이브 젖꼭지/뚜껑류는 _is_glass_bottle 규칙이 자동 제외하므로 여기 둘 필요 없음)
EXCLUDE_PRODUCTS: list[str] = [
    "[임신출산선물] 란시노 베스트셀러 박스 A(젖병, 젖꼭지, 수유패드 포함 + 파우치 증정)",
    "[임신출산선물] 란시노 베스트셀러 박스 B(젖병, 젖꼭지, 수유패드 포함 + 파우치 증정)",
]

# 프로모션 일정 (행사명, 시작, 종료) — 종료일 포함
PROMO_SCHEDULE = [
    ("설 선물 기획전",        "2025-01-20", "2025-01-22"),
    ("새봄맞이 프로모션",     "2025-03-17", "2025-03-19"),
    ("가정의 달 빅프로모션",  "2025-05-19", "2025-05-22"),
    ("추석맞이 선물대전",     "2025-09-22", "2025-09-25"),
    ("란시노 블랙 프라이데이","2025-11-24", "2025-11-27"),
    ("봄맞이 프로모션",       "2026-03-03", "2026-03-06"),
]

WEEKDAY_KR = ["월", "화", "수", "목", "금", "토", "일"]

# 일간데이터 열 위치(0-based). 헤더가 병합/중복이라 위치 기반 추출.
DAILY_COL_POS = {
    "날짜": 0, "브랜드": 1, "제품": 2,
    "광고비": 9,                       # J (paid 총광고비)
    "마케팅_조회수": 16, "마케팅_비용": 17,          # Q, R
    "바이럴_1~3위건수": 18, "바이럴_1~3위조회수": 19, "바이럴_비용": 21,  # S,T,V
    "숏폼_발행수": 22, "숏폼_조회수": 23, "숏폼_비용": 25,             # W,X,Z
    "사이다_발행수": 26, "사이다_조회수": 27, "사이다_비용": 29,        # AA,AB,AD
}
DAILY_ACTIVITY_COLS = [k for k in DAILY_COL_POS if k not in ("날짜", "브랜드", "제품")]

# 매출_raw 열 위치(0-based, 헤더 1행)
RAW_COLS = ["결제일자", "브랜드", "제품", "SKU", "SKU판매가", "원가",
            "실결제금액", "수량", "계정별칭"]


# ============================================================================
# 진단 컨테이너
# ============================================================================
@dataclass
class JoinDiagnostics:
    raw_matched_products: list = field(default_factory=list)      # 매칭된 distinct 제품명
    raw_suspect_products: list = field(default_factory=list)      # 의심 매칭 제품명
    n_raw_rows_total: int = 0
    n_raw_rows_matched: int = 0
    n_dates_sales: int = 0
    n_dates_daily: int = 0
    n_dates_joined: int = 0
    n_dates_adspend_filled0: int = 0   # 광고비 결측 -> 0 대체 일수
    daily_header_ok: bool = True
    messages: list = field(default_factory=list)

    def log(self, msg: str):
        self.messages.append(msg)
        print(f"[진단] {msg}")


# ============================================================================
# 1. 로더
# ============================================================================
def load_sales_raw(path: str, sheet_name="매출_raw") -> pd.DataFrame:
    """매출_raw 로드(헤더 1행). 위치 기반으로 9개 열 매핑."""
    df = pd.read_excel(path, sheet_name=sheet_name, header=0)
    if df.shape[1] < len(RAW_COLS):
        raise ValueError(f"매출_raw 열 수 부족: {df.shape[1]} < {len(RAW_COLS)}")
    df = df.iloc[:, :len(RAW_COLS)].copy()
    df.columns = RAW_COLS
    df["결제일자"] = pd.to_datetime(df["결제일자"], errors="coerce").dt.normalize()
    for c in ["SKU판매가", "원가", "실결제금액", "수량"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df.dropna(subset=["결제일자"])
    df["브랜드"] = df["브랜드"].astype(str).str.strip()
    df["제품"] = df["제품"].astype(str).str.strip()
    df["SKU"] = df["SKU"].astype(str).str.strip()
    return df


def load_daily_data(path: str, sheet_name=0, diag: JoinDiagnostics | None = None
                    ) -> pd.DataFrame:
    """일간데이터 로드(헤더 2행). 위치 기반 추출 + 헤더 정합성 검증."""
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
    # 헤더 정합성: 2행(인덱스1) 세부 라벨이 기대값과 일치하는지 확인
    def _h(idx):
        try:
            return str(raw.iloc[1, idx]).strip()
        except Exception:
            return ""
    checks = {"날짜": (0, "날짜"), "제품": (2, "제품"), "광고비": (9, "광고비")}
    ok = all(_h(pos) == label for label, (pos, _) in
             [(k, (v[0], v[1])) for k, v in checks.items()])
    if diag is not None and not ok:
        diag.daily_header_ok = False
        found = {lbl: _h(pos) for lbl, (pos, _) in checks.items()}
        diag.log(f"[경고] 일간데이터 헤더 위치 불일치 추정. 기대 vs 실제={found}. "
                 f"열 위치(DAILY_COL_POS) 재확인 필요.")

    data = raw.iloc[2:].copy()   # 3행부터 데이터
    out = pd.DataFrame()
    for name, pos in DAILY_COL_POS.items():
        out[name] = data.iloc[:, pos] if pos < data.shape[1] else np.nan
    out["날짜"] = pd.to_datetime(out["날짜"], errors="coerce").dt.normalize()
    out["브랜드"] = out["브랜드"].astype(str).str.strip()
    out["제품"] = out["제품"].astype(str).str.strip()
    for c in DAILY_ACTIVITY_COLS:
        out[c] = pd.to_numeric(out[c], errors="coerce")
    out = out.dropna(subset=["날짜"])
    # #REF! 등 깨진 키 행 제거
    out = out[~out["브랜드"].isin(["#REF!", "nan", "None", ""])]
    return out


# ============================================================================
# 2. 대상 필터 + 유리젖병 매칭 진단
# ============================================================================
def _is_glass_bottle(name: str) -> bool:
    """유리젖병 본품 판정.
    - 확정: '유리젖병'/'유리 젖병' 과 용량(ml)이 함께 있으면 본품(사용자 규칙).
    - 제외: 위 확정 신호가 없으면서 젖꼭지/뚜껑류(내추럴웨이브 등)면 부속.
    - 그 외: 기존 규칙('젖병' 포함 & 'PP' 미포함).
    """
    s = str(name)
    if "PP" in s.upper():
        return False
    is_glass = ("유리젖병" in s) or ("유리 젖병" in s)
    has_ml = bool(re.search(r"\d+\s*ml", s, re.IGNORECASE))
    if is_glass and has_ml:                 # 유리젖병 + ml -> 본품 확정
        return True
    non_body = ["내추럴웨이브", "내추럴 웨이브", "젖병꼭지", "뚜껑", "캡", "보관"]
    if any(k in s for k in non_body) and not is_glass:
        return False
    return "젖병" in s


def _is_suspect(name: str) -> bool:
    """검토용 '의심' 라벨. 단, 유리젖병+ml로 본품이 확정된 것은 의심에서 제외."""
    s = str(name)
    is_glass = ("유리젖병" in s) or ("유리 젖병" in s)
    has_ml = bool(re.search(r"\d+\s*ml", s, re.IGNORECASE))
    if is_glass and has_ml:                 # 본품 확정 -> 의심 아님
        return False
    return any(kw in s for kw in ACCESSORY_KEYWORDS)


def filter_target_raw(df_raw: pd.DataFrame, diag: JoinDiagnostics) -> pd.DataFrame:
    """매출_raw에서 란시노 유리젖병 거래 추출 + 매칭 진단."""
    diag.n_raw_rows_total = len(df_raw)
    brand_mask = df_raw["브랜드"] == TARGET_BRAND
    sub = df_raw[brand_mask].copy()
    sub["_is_glass"] = sub["제품"].map(_is_glass_bottle)
    matched = sub[sub["_is_glass"]].copy()

    # 수기 확정 제외 적용
    if EXCLUDE_PRODUCTS:
        before = len(matched)
        matched = matched[~matched["제품"].isin(EXCLUDE_PRODUCTS)]
        diag.log(f"EXCLUDE_PRODUCTS 적용: {before - len(matched)}행 제외 "
                 f"({len(EXCLUDE_PRODUCTS)}종 지정)")
    diag.n_raw_rows_matched = len(matched)

    distinct = sorted(matched["제품"].unique().tolist())
    diag.raw_matched_products = distinct
    diag.raw_suspect_products = [p for p in distinct if _is_suspect(p)]

    diag.log(f"매출_raw 총 {diag.n_raw_rows_total}행 중 란시노 {int(brand_mask.sum())}행, "
             f"유리젖병 매칭 {diag.n_raw_rows_matched}행 / distinct 제품 {len(distinct)}종")
    diag.log(f"매칭 제품명: {distinct}")
    if diag.raw_suspect_products:
        diag.log(f"[의심 매칭] 부속/비매출 추정 {len(diag.raw_suspect_products)}종 "
                 f"-> 수기 확인 권장: {diag.raw_suspect_products}")
    else:
        diag.log("의심 매칭 없음.")
    return matched.drop(columns=["_is_glass"])


def filter_target_daily(df_daily: pd.DataFrame, diag: JoinDiagnostics) -> pd.DataFrame:
    """일간데이터에서 란시노·유리젖병 행 추출(일자 단위)."""
    m = (df_daily["브랜드"] == TARGET_BRAND) & (df_daily["제품"] == TARGET_PRODUCT_DAILY)
    sub = df_daily[m].copy()
    # 동일 (날짜) 중복 시 합산
    agg = {c: "sum" for c in DAILY_ACTIVITY_COLS}
    out = sub.groupby("날짜", as_index=False).agg(agg)
    diag.n_dates_daily = out["날짜"].nunique()
    diag.log(f"일간데이터 란시노·{TARGET_PRODUCT_DAILY} 일자 {diag.n_dates_daily}일")
    return out


# ============================================================================
# 3. 일별 집계 + 조인 + 모델링 프레임
# ============================================================================
def aggregate_daily_sales(df_matched: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """일별 매출/수량 집계(Y) + SKU 단위 일별 집계(분해용)."""
    daily = (df_matched.groupby("결제일자", as_index=False)
             .agg(매출=("실결제금액", "sum"), 수량=("수량", "sum"))
             .rename(columns={"결제일자": "날짜"}))
    sku = (df_matched.groupby(["결제일자", "SKU"], as_index=False)
           .agg(매출=("실결제금액", "sum"), 수량=("수량", "sum"))
           .rename(columns={"결제일자": "날짜"}))
    return daily, sku


def build_model_frame(daily_sales: pd.DataFrame, daily_act: pd.DataFrame,
                      diag: JoinDiagnostics) -> pd.DataFrame:
    """매출(Y) + 일간 활동/광고비 조인 + 추세·요일·프로모션 특성 생성."""
    diag.n_dates_sales = daily_sales["날짜"].nunique()
    df = daily_sales.merge(daily_act, on="날짜", how="left")
    diag.n_dates_joined = df["날짜"].nunique()

    # 광고비 결측 -> 0 (집행 미기록으로 해석). 결측 여부 플래그 보존(가짜0 진단용)
    df["광고비결측"] = df["광고비"].isna()
    fill0 = int(df["광고비결측"].sum())
    diag.n_dates_adspend_filled0 = fill0
    if fill0:
        diag.log(f"광고비 결측 {fill0}일 -> 0 대체")
    for c in DAILY_ACTIVITY_COLS:
        df[c] = df[c].fillna(0.0)

    df = df.sort_values("날짜").reset_index(drop=True)
    # 추세 t = 기준시작일로부터의 '달력 일수'. (구글시트 수식이 날짜만으로 t를 재현하려면
    #  행 인덱스가 아니라 달력일 기준이어야 함. 결측일이 있어도 수식이 정확해짐.)
    base_date = df["날짜"].min()
    df["t"] = (df["날짜"] - base_date).dt.days.astype(float)
    df.attrs["base_date"] = base_date
    df["요일"] = df["날짜"].dt.weekday                  # 0=월
    df["요일명"] = df["요일"].map(lambda i: WEEKDAY_KR[i])

    # 프로모션 더미(union) + 행사명
    df["프로모션"] = 0
    df["행사명"] = ""
    for name, s, e in PROMO_SCHEDULE:
        m = (df["날짜"] >= pd.Timestamp(s)) & (df["날짜"] <= pd.Timestamp(e))
        df.loc[m, "프로모션"] = 1
        df.loc[m, "행사명"] = name

    # 로그매출(이분산 완화) + 전일 로그매출(lag-1, 자기상관 흡수)
    floor = 1.0  # 0/음수 방지 바닥
    df["매출_clip"] = df["매출"].clip(lower=floor)
    df["logY"] = np.log(df["매출_clip"])
    df["lag_logY"] = df["logY"].shift(1)   # 직전 행(달력 결측 시 직전 존재일)
    return df


# ============================================================================
# 4. 베이스라인 OLS
# ============================================================================
def _design_matrix(df: pd.DataFrame) -> tuple[pd.DataFrame, list]:
    """설계행렬: const + [lag_logY] + t + 요일더미(월 기준 drop) + 프로모션 + 광고비 [+교차항]."""
    X = pd.DataFrame(index=df.index)
    if USE_LAG1:
        X["lag_logY"] = df["lag_logY"].astype(float)   # 전일 로그매출
    X["t"] = df["t"]
    wd = pd.get_dummies(df["요일명"], prefix="요일")
    drop_col = "요일_월"
    if drop_col in wd.columns:
        wd = wd.drop(columns=[drop_col])           # 월요일 기준
    X = pd.concat([X, wd.astype(float)], axis=1)
    X["프로모션"] = df["프로모션"].astype(float)
    X["광고비"] = df["광고비"].astype(float)
    if USE_CROSS_PROMO_ADSPEND:
        X["프로모션x광고비"] = X["프로모션"] * X["광고비"]
    X = sm.add_constant(X, has_constant="add")
    return X, list(X.columns)


def _target(df: pd.DataFrame) -> np.ndarray:
    """적합 대상 y. 로그 옵션 시 logY, 아니면 매출(원)."""
    return df["logY"].values if USE_LOG_TARGET else df["매출"].astype(float).values


def fit_baseline_ols(df: pd.DataFrame):
    """OLS 적합(로그매출 + HAC 강건 표준오차). 반환: results, sigma(모델스케일 잔차σ), 열명."""
    if len(df) < MIN_DAYS_FIT:
        raise ValueError(f"적합 표본 부족: {len(df)}일 < MIN_DAYS_FIT={MIN_DAYS_FIT}")
    X, cols = _design_matrix(df)
    y = _target(df)
    XY = X.astype(float).copy()
    XY["_y"] = y
    XY = XY.dropna()                       # lag-1로 인한 첫 행 결측 제거
    yv = XY.pop("_y").values
    # HAC(Newey-West): 자기상관·이분산에 강건한 표준오차(계수 p값용)
    res = sm.OLS(yv, XY).fit(cov_type="HAC", cov_kwds={"maxlags": HAC_MAXLAGS})
    sigma = float(np.sqrt(res.scale))      # 잔차 표준편차(모델 스케일)
    return res, sigma, cols


def compute_sigmas(res, sigma_global: float, df: pd.DataFrame) -> dict:
    """요일별 잔차 σ 산출(모델 스케일). 표본 부족 요일은 전체 σ로 대체.
    반환: {요일명: σ}. USE_WEEKDAY_SIGMA=False면 전부 전체 σ."""
    if not USE_WEEKDAY_SIGMA:
        return {wd: sigma_global for wd in WEEKDAY_KR}
    X, _ = _design_matrix(df)
    pred = np.asarray(res.predict(X.astype(float)), dtype=float)
    y = _target(df)
    resid = y - pred
    tmp = pd.DataFrame({"요일명": df["요일명"].values, "resid": resid}).dropna()
    out = {}
    for wd in WEEKDAY_KR:
        r = tmp.loc[tmp["요일명"] == wd, "resid"]
        out[wd] = float(r.std(ddof=1)) if len(r) >= WEEKDAY_SIGMA_MIN_N else sigma_global
    return out


def reference_means(res, df: pd.DataFrame) -> pd.DataFrame:
    """계수 보고표(요일 효과 별도 보고 포함). 요일계수=월요일 대비 증분."""
    params = res.params
    pvals = res.pvalues
    rows = []
    for name in params.index:
        rows.append({
            "변수": name,
            "계수": params[name],
            "표준오차": res.bse[name],
            "t값": res.tvalues[name],
            "p값": pvals[name],
            "유의(p<0.05)": "O" if pvals[name] < 0.05 else "",
        })
    out = pd.DataFrame(rows)
    out.attrs["note"] = ("요일계수=월요일 대비 증분(원). 추세 t=일당 변화량. "
                         "절편에 레벨 포함(평균 재합산 안 함).")
    return out


def build_sheet_formula(res, sigma: float, df: pd.DataFrame, sigma_map: dict = None
                        ) -> pd.DataFrame:
    """구글시트(B방안)용: 베이스라인 계수+σ를 고정하고, 신규일을 '날짜·광고비·프로모션'만으로
    예측/상하한/이상여부를 시트가 자동 계산하도록 하는 수식과 값을 산출.

    시트에 [날짜], [광고비], [프로모션(0/1)], [실매출] 열이 있다고 가정.
    아래 수식의 대괄호를 실제 셀 주소로 바꿔 붙여넣으면 됨.
    """
    p = res.params
    def g(name):  # 계수 안전 추출(없으면 0)
        return float(p[name]) if name in p.index else 0.0

    base = df.attrs.get("base_date", df["날짜"].min())
    bd = pd.Timestamp(base)
    const, b_t, b_promo, b_ad = g("const"), g("t"), g("프로모션"), g("광고비")
    b_lag = g("lag_logY")
    wd = {"월": 0.0, "화": g("요일_화"), "수": g("요일_수"), "목": g("요일_목"),
          "금": g("요일_금"), "토": g("요일_토"), "일": g("요일_일")}
    k2 = SIGMA_K
    sm_ = sigma_map or {w: sigma for w in WEEKDAY_KR}   # 요일별 σ

    base_date_str = f"DATE({bd.year},{bd.month},{bd.day})"
    f_dow = (f"CHOOSE(WEEKDAY([날짜],2),{wd['월']:.4f},{wd['화']:.4f},{wd['수']:.4f},"
             f"{wd['목']:.4f},{wd['금']:.4f},{wd['토']:.4f},{wd['일']:.4f})")
    # 요일별 σ 선택식(WEEKDAY 1=월..7=일)
    f_sig = (f"CHOOSE(WEEKDAY([날짜],2),{sm_['월']:.5f},{sm_['화']:.5f},{sm_['수']:.5f},"
             f"{sm_['목']:.5f},{sm_['금']:.5f},{sm_['토']:.5f},{sm_['일']:.5f})")
    lag_term = f"{b_lag:.6f}*LN([전일매출])+" if USE_LAG1 else ""
    lp = (f"{const:.4f}+{lag_term}{b_t:.6f}*([날짜]-{base_date_str})+{f_dow}"
          f"+{b_promo:.4f}*[프로모션]+{b_ad:.8f}*[광고비]")

    if USE_LOG_TARGET:
        # 원단위 예측 = EXP(LP). 관리폭은 요일별 σ로 곱셈.
        f_pred = f"EXP({lp})"
        f_upper = f"[예측셀]*EXP({k2}*{f_sig})"
        f_lower = f"[예측셀]*EXP(-{k2}*{f_sig})"
    else:
        f_pred = lp
        f_upper = f"[예측셀]+{k2}*{f_sig}"
        f_lower = f"[예측셀]-{k2}*{f_sig}"
    f_flag = 'IF([실매출]>[상한셀],"이상(상회)",IF([실매출]<[하한셀],"이상(하회)","정상"))'
    f_dev = "([실매출]-[예측셀])/[예측셀]*100"

    rows = [
        ("■ 사용법", "각 '수식'을 시트에 붙여넣을 때 맨 앞에 '='를 추가. "
                     "[날짜]/[광고비]/[프로모션]/[실매출]/[전일매출]은 해당 셀 주소로, "
                     "[예측셀]/[상한셀]/[하한셀]도 계산된 셀 주소로 치환."),
        ("모델", f"로그매출 + {'lag1(전일매출) + ' if USE_LAG1 else ''}추세 + 요일 + 프로모션 + 광고비"
                 f"{' / 요일별σ' if USE_WEEKDAY_SIGMA else ''}"),
        ("기준시작일(t=0)", bd.strftime("%Y-%m-%d")),
        ("절편(const, 로그)", round(const, 4)),
        ("전일매출 계수 b_lag", round(b_lag, 6)),
        ("추세 b_t (로그/일)", round(b_t, 6)),
        ("프로모션 계수(로그)", round(b_promo, 4)),
        ("광고비 계수(로그/원)", round(b_ad, 8)),
        ("요일_화 증분", round(wd["화"], 4)),
        ("요일_수 증분", round(wd["수"], 4)),
        ("요일_목 증분", round(wd["목"], 4)),
        ("요일_금 증분", round(wd["금"], 4)),
        ("요일_토 증분", round(wd["토"], 4)),
        ("요일_일 증분", round(wd["일"], 4)),
        ("σ_월", round(sm_["월"], 5)), ("σ_화", round(sm_["화"], 5)),
        ("σ_수", round(sm_["수"], 5)), ("σ_목", round(sm_["목"], 5)),
        ("σ_금", round(sm_["금"], 5)), ("σ_토", round(sm_["토"], 5)),
        ("σ_일", round(sm_["일"], 5)),
        ("── 수식 ──", ""),
        ("예측 수식", f_pred),
        ("상한 수식(요일별σ)", f_upper),
        ("하한 수식(요일별σ)", f_lower),
        ("이상여부 수식", f_flag),
        ("이탈% 수식", f_dev),
    ]
    out = pd.DataFrame(rows, columns=["항목", "수식/값"])
    out.attrs["note"] = "베이스라인 계수 고정 -> 시트가 신규일을 자동 판정(predict-only)."
    return out


# ============================================================================
# 5. 이상치 탐지
# ============================================================================
def detect_anomalies(res, sigma: float, df: pd.DataFrame, sigma_map: dict = None
                     ) -> pd.DataFrame:
    """예측 ±SIGMA_K·σ 이탈일 탐지. 로그모델이면 원단위로 역변환.
    sigma_map(요일별 σ) 주어지면 요일마다 다른 관리폭 적용."""
    X, _ = _design_matrix(df)
    pred_m = np.asarray(res.predict(X.astype(float)), dtype=float)  # 모델스케일 예측

    # 요일별 σ 벡터(없으면 전체 σ)
    if sigma_map:
        sig = df["요일명"].map(lambda w: sigma_map.get(w, sigma)).astype(float).values
    else:
        sig = np.full(len(df), sigma, dtype=float)

    out = df[["날짜", "요일명", "프로모션", "행사명", "매출", "수량", "광고비"]].copy()
    out["σ적용"] = sig
    if USE_LOG_TARGET:
        out["예측"] = np.exp(pred_m)
        out["상한"] = np.exp(pred_m + SIGMA_K * sig)
        out["하한"] = np.exp(pred_m - SIGMA_K * sig)
        out["z"] = np.where(sig > 0, (df["logY"].values - pred_m) / sig, 0.0)
    else:
        out["예측"] = pred_m
        out["상한"] = pred_m + SIGMA_K * sig
        out["하한"] = pred_m - SIGMA_K * sig
        out["z"] = np.where(sig > 0, (out["매출"].values - pred_m) / sig, 0.0)

    out["잔차"] = out["매출"] - out["예측"]
    out["이탈%"] = np.where(out["예측"] > 0,
                           (out["매출"] - out["예측"]) / out["예측"] * 100, np.nan).round(1)
    out["이상치"] = (out["매출"] > out["상한"]) | (out["매출"] < out["하한"])
    out.loc[~np.isfinite(pred_m), "이상치"] = False
    out["방향"] = np.where(out["이상치"] & (out["매출"] > out["상한"]), "상회",
                          np.where(out["이상치"] & (out["매출"] < out["하한"]), "하회", ""))
    return out


def residual_diagnostics(res, sigma: float) -> pd.DataFrame:
    """±2σ 가정의 통계적 정당성 검증: 잔차 정규성·자기상관·등분산.
    R²가 낮아도 이 검정이 양호하면 ±2σ 판정은 정당함.
    """
    resid = np.asarray(res.resid, dtype=float)
    rows = []

    # 정규성 (Jarque-Bera: statsmodels 내장)
    try:
        from statsmodels.stats.stattools import jarque_bera
        jb, jb_p, skew, kurt = jarque_bera(resid)
        rows.append(("잔차 정규성 (Jarque-Bera p)", round(float(jb_p), 4),
                     "p>0.05면 정규에 가까움(±2σ 의미 성립)"))
        rows.append(("  - 왜도(skew)", round(float(skew), 3), "0에 가까울수록 대칭"))
        rows.append(("  - 첨도(kurtosis)", round(float(kurt), 3), "3에 가까울수록 정규"))
    except Exception as e:
        rows.append(("잔차 정규성", "계산불가", str(e)[:40]))

    # 자기상관 (Durbin-Watson)
    try:
        from statsmodels.stats.stattools import durbin_watson
        dw = float(durbin_watson(resid))
        rows.append(("잔차 자기상관 (Durbin-Watson)", round(dw, 3),
                     "2 근처=무상관 / 2미만=양의상관(시계열 주의)"))
    except Exception as e:
        rows.append(("Durbin-Watson", "계산불가", str(e)[:40]))

    # 자기상관 (Ljung-Box, lag10)
    try:
        from statsmodels.stats.diagnostic import acorr_ljungbox
        lb = acorr_ljungbox(resid, lags=[10], return_df=True)
        rows.append(("잔차 자기상관 (Ljung-Box lag10 p)",
                     round(float(lb["lb_pvalue"].iloc[0]), 4),
                     "p>0.05면 자기상관 없음(양호)"))
    except Exception as e:
        rows.append(("Ljung-Box", "계산불가", str(e)[:40]))

    # 등분산 (Breusch-Pagan)
    try:
        from statsmodels.stats.diagnostic import het_breuschpagan
        bp = het_breuschpagan(resid, res.model.exog)
        rows.append(("등분산성 (Breusch-Pagan p)", round(float(bp[1]), 4),
                     "p>0.05면 등분산(양호) / 낮으면 고매출일 과검출 위험"))
    except Exception as e:
        rows.append(("Breusch-Pagan", "계산불가", str(e)[:40]))

    rows.append(("σ (잔차표준편차)", round(sigma, 1), "±2σ 관리한계의 폭 근거"))
    return pd.DataFrame(rows, columns=["검정", "값", "해석"])


def sigma_sensitivity(res, sigma: float, df: pd.DataFrame, sigma_map: dict = None
                      ) -> pd.DataFrame:
    """SIGMA_K 민감도: 관리한계 배수별 이상일 수. 2.0 근처에서 급변하면 기준 불안정."""
    X, _ = _design_matrix(df)
    pred_m = np.asarray(res.predict(X.astype(float)), dtype=float)
    ym = df["logY"].values if USE_LOG_TARGET else df["매출"].astype(float).values
    if sigma_map:
        sig = df["요일명"].map(lambda w: sigma_map.get(w, sigma)).astype(float).values
    else:
        sig = np.full(len(df), sigma, dtype=float)
    valid = np.isfinite(pred_m) & np.isfinite(ym)
    pred_m, ym, sig = pred_m[valid], ym[valid], sig[valid]
    n = len(ym)
    rows = []
    from scipy.stats import norm
    k_list = sorted({1.0, 1.25, 1.5, 2.0, 2.5, 3.0, round(float(SIGMA_K), 2)})
    for k in k_list:
        hi, lo = pred_m + k * sig, pred_m - k * sig
        n_up = int((ym > hi).sum())
        n_dn = int((ym < lo).sum())
        exp_pct = float(2 * (1 - norm.cdf(k)) * 100)   # 정규가정 기대 이탈률
        rows.append({
            "SIGMA_K": k, "이상일수": n_up + n_dn, "상회": n_up, "하회": n_dn,
            "비율(%)": round(100 * (n_up + n_dn) / n, 1),
            "정규기대(%)": round(exp_pct, 2),
            "현재선택": "◀ 현재" if abs(k - float(SIGMA_K)) < 1e-9 else "",
        })
    out = pd.DataFrame(rows)
    out.attrs["note"] = ("관리한계 배수 K가 낮을수록 이상일을 더 많이 잡음(밴드 좁힘). "
                         "실제비율이 정규기대와 가까우면 ±Kσ 기준이 데이터에 부합. "
                         "정규기대=2·(1−Φ(K))·100, 잔차 정규가정 기준이라 실제와 다를 수 있음(11_잔차진단 참고).")
    return out


def diagnose_pred_spikes(anomaly_df: pd.DataFrame) -> pd.DataFrame:
    """예측이 전일 대비 크게 튄 날을 찾아 원인(광고비 급변/프로모션) 추정.
    '예측이 왜 불쑥 튀나'에 대한 근거표. 전일매출 lag는 약하게만(약 6%) 작용하므로
    광고비·프로모션이 더 크게 예측을 움직인다.
    """
    d = anomaly_df[["날짜", "예측", "광고비", "프로모션", "행사명", "매출"]].copy()
    d = d.sort_values("날짜").reset_index(drop=True)
    d["전일예측"] = d["예측"].shift(1)
    d["전일광고비"] = d["광고비"].shift(1)
    d["예측변동%"] = np.where(d["전일예측"] > 0,
                            (d["예측"] - d["전일예측"]) / d["전일예측"] * 100, np.nan)
    d["광고비변동%"] = np.where(d["전일광고비"] > 0,
                             (d["광고비"] - d["전일광고비"]) / d["전일광고비"] * 100,
                             np.where(d["광고비"] > 0, 100.0, 0.0))
    spikes = d[d["예측변동%"].abs() >= SPIKE_PCT].copy()

    def cause(r):
        if r["프로모션"] == 1:
            return f"프로모션({r['행사명']})" if r["행사명"] else "프로모션"
        if abs(r["광고비변동%"]) >= SPIKE_PCT:
            return "광고비 급변"
        return "기타(요일/추세/전일매출)"
    if not spikes.empty:
        spikes["추정원인"] = spikes.apply(cause, axis=1)
    else:
        spikes["추정원인"] = pd.Series(dtype=str)

    cols = ["날짜", "예측", "전일예측", "예측변동%", "광고비", "전일광고비",
            "광고비변동%", "프로모션", "행사명", "추정원인"]
    spikes = spikes[cols].round({"예측": 0, "전일예측": 0, "예측변동%": 1,
                                 "광고비변동%": 1})
    spikes.attrs["note"] = (f"예측이 전일 대비 ±{SPIKE_PCT:.0f}%p 이상 변한 날. "
                            "전일매출 계수가 작아(≈0.06) 광고비·프로모션이 예측을 더 크게 움직임.")
    return spikes.reset_index(drop=True)


def decompose_prediction(res, sigma: float, df: pd.DataFrame
                         ) -> tuple[pd.DataFrame, pd.DataFrame]:
    """예측을 변수별 기여로 분해(로그스케일 가법 = 원단위 곱셈).
    반환: (전체 분해표, 봉우리 주원인표). '예측이 왜 솟나'를 변수별 숫자로 규명."""
    X, cols = _design_matrix(df)
    p = res.params
    def g(n): return float(p[n]) if n in p.index else 0.0

    c = pd.DataFrame({"날짜": df["날짜"].values})
    c["상수"] = g("const")
    c["추세"] = g("t") * df["t"].values
    dow_cols = [col for col in cols if col.startswith("요일_")]
    dow = np.zeros(len(df))
    for col in dow_cols:
        dow += g(col) * X[col].astype(float).values
    c["요일"] = dow
    c["전일매출"] = g("lag_logY") * df["lag_logY"].values if USE_LAG1 else 0.0
    c["프로모션"] = g("프로모션") * df["프로모션"].values
    c["광고비"] = g("광고비") * df["광고비"].values
    terms = ["상수", "추세", "요일", "전일매출", "프로모션", "광고비"]
    c["log예측"] = c[terms].sum(axis=1)
    c["예측(원)"] = np.exp(c["log예측"])

    # 봉우리 비율 = 예측 / 14일 이동중앙값 (국소적으로 얼마나 솟았나)
    med = pd.Series(c["예측(원)"].values).rolling(14, min_periods=3, center=True).median()
    c["봉우리비율"] = c["예측(원)"].values / med.values

    full = c.copy()
    full[terms] = full[terms].round(4)
    full["예측(원)"] = full["예측(원)"].round(0)
    full["봉우리비율"] = full["봉우리비율"].round(2)

    # 봉우리 주원인: 변동 가능한 항(요일/전일매출/프로모션/광고비) 중
    # 자기 평균 대비 가장 크게 끌어올린 항을 원인으로
    var_terms = ["요일", "전일매출", "프로모션", "광고비"]
    means = {t: float(c[t].mean()) for t in var_terms}
    sp = c.dropna(subset=["봉우리비율"]).nlargest(10, "봉우리비율")
    rows = []
    for _, r in sp.iterrows():
        devs = {t: r[t] - means[t] for t in var_terms}
        cause = max(devs, key=devs.get)            # 평균 대비 최대 상승 항
        rows.append({
            "날짜": r["날짜"], "예측(원)": round(r["예측(원)"], 0),
            "봉우리비율": round(r["봉우리비율"], 2),
            "주원인": cause,
            "주원인_평균대비(로그)": round(devs[cause], 4),
            "주원인_배수": round(float(np.exp(devs[cause])), 3),
            "요일기여": round(r["요일"], 3), "전일매출기여": round(r["전일매출"], 3),
            "프로모션기여": round(r["프로모션"], 3), "광고비기여": round(r["광고비"], 3),
        })
    spike_cause = pd.DataFrame(rows)
    spike_cause.attrs["note"] = ("봉우리비율=예측/14일중앙값. 주원인_배수=그 항이 평균 대비 "
                                 "예측을 몇 배로 올렸나(EXP). 로그모델이라 고매출 구간서 더 크게 보임.")
    return full, spike_cause


def weekday_diagnostics(res, anomaly_df: pd.DataFrame, sigma_map: dict = None
                        ) -> pd.DataFrame:
    """요일별 반영 점검: 절편(월=기준)·요일계수·평균이탈·잔차분산·이상치율·요일σ."""
    p = res.params
    const = float(p.get("const", 0.0))
    rows = []
    for wd in WEEKDAY_KR:                       # 월~일 순
        sub = anomaly_df[anomaly_df["요일명"] == wd]
        n = len(sub)
        coef = 0.0 if wd == "월" else float(p.get(f"요일_{wd}", 0.0))
        rows.append({
            "요일": wd,
            "일수": n,
            "요일계수(로그)": round(coef, 4),
            "요일배수(EXP)": round(float(np.exp(coef)), 3),   # 월 대비 몇 배
            "요일σ(로그)": round(sigma_map.get(wd), 4) if sigma_map else np.nan,
            "평균매출": round(sub["매출"].mean(), 0) if n else np.nan,
            "평균이탈%": round(sub["이탈%"].mean(), 1) if n else np.nan,
            "이탈%_표준편차": round(sub["이탈%"].std(), 1) if n else np.nan,
            "이상치수": int(sub["이상치"].sum()) if n else 0,
            "이상치율%": round(100 * sub["이상치"].mean(), 1) if n else np.nan,
        })
    out = pd.DataFrame(rows)
    out.attrs["note"] = (f"절편(const, 로그)={const:.4f} -> 월요일·t=0 기준 예측=EXP(const). "
                         "요일σ=그 요일 전용 관리폭(요일별σ 적용 시). 요일σ 큰 요일은 폭이 넓어 "
                         "과검출 교정됨. 평균이탈%가 0 근처면 평균은 잘 맞은 것.")
    return out


def vif_diagnostics(res) -> pd.DataFrame:
    """다중공선성(VIF). 광고비 계수 음수/불안정의 원인 점검. VIF>5 주의, >10 심각."""
    from statsmodels.stats.outliers_influence import variance_inflation_factor
    exog = np.asarray(res.model.exog, dtype=float)
    names = list(res.model.exog_names)
    rows = []
    for i, nm in enumerate(names):
        if nm == "const":
            continue
        try:
            v = float(variance_inflation_factor(exog, i))
        except Exception:
            v = np.nan
        flag = "심각(>10)" if v > 10 else ("주의(>5)" if v > 5 else "양호")
        rows.append({"변수": nm, "VIF": round(v, 2), "판정": flag,
                     "계수": round(float(res.params.get(nm, np.nan)), 8),
                     "p값": round(float(res.pvalues.get(nm, np.nan)), 4)})
    out = pd.DataFrame(rows).sort_values("VIF", ascending=False).reset_index(drop=True)
    out.attrs["note"] = ("VIF=다른 변수로 그 변수를 얼마나 설명되나(공선성). 높을수록 계수 불안정. "
                         "광고비 VIF가 높으면 음수 계수는 공선성 산물로 해석.")
    return out


def relationship_correlations(df: pd.DataFrame) -> pd.DataFrame:
    """매출/수량과 각 변수의 상관(원시 + 추세제거). 대표님 '매출↑때 광고비↑' 검증 +
    회귀계수(음수)와 원시상관(양수?)의 차이를 추세교란으로 설명."""
    vars_ = ["광고비"] + [c for c in DAILY_ACTIVITY_COLS if c != "광고비"]
    t = df["t"].values.astype(float)

    def detrend(col):
        x = df[col].values.astype(float)
        b = np.polyfit(t, x, 1)
        return pd.Series(x - (b[0] * t + b[1]))

    dm_sales, dm_qty = detrend("매출"), detrend("수량")
    rows = []
    for v in vars_:
        dv = detrend(v)
        rows.append({
            "변수": v,
            "매출_상관(원시)": round(df["매출"].corr(df[v]), 3),
            "수량_상관(원시)": round(df["수량"].corr(df[v]), 3),
            "매출_상관(추세제거)": round(dm_sales.corr(dv), 3),
            "수량_상관(추세제거)": round(dm_qty.corr(dv), 3),
        })
    out = pd.DataFrame(rows)
    rs = round(df["매출"].corr(df["수량"]), 3)
    out.attrs["note"] = (
        f"매출↔수량 상관={rs}. 원시상관이 +인데 추세제거 후 0/−로 바뀌면 '추세 교란'"
        "(둘 다 시간에 따라 함께 성장)이지 직접 효과가 아님. 회귀계수(음수)는 추세·요일 "
        "통제 후 값이라 추세제거 상관과 방향이 비슷함. 대표님 '매출↑때 광고비↑'는 "
        "원시 광고비-매출 상관 부호로 확인.")
    return out


def _to_calendar(df: pd.DataFrame) -> pd.DataFrame:
    """매출 있는 날만 있는 프레임을 '달력일 전체'로 재색인(빠진 날=NaN).
    -> .diff()가 결측일 경계에서 NaN이 되어, '진짜 하루 차이'만 남음."""
    d = df.sort_values("날짜").set_index("날짜")
    full = pd.date_range(d.index.min(), d.index.max(), freq="D")
    return d.reindex(full).rename_axis("날짜").reset_index()


def adspend_increment_effect(df: pd.DataFrame) -> pd.DataFrame:
    """광고비 '증액'이 매출 '증가'와 연결되나(차분 분석). 추세는 차분으로 자동 제거.
    - 일별/주별 Δ광고비 ↔ Δ매출 상관(당기·시차)
    - 당기만 +이고 시차가 0이면 '같이 움직일 뿐'(공통요인), 시차도 +면 광고가 미는 근거.
    """
    d = df.sort_values("날짜").reset_index(drop=True).copy()
    rows = []

    # ── 일별 차분 (달력일 기준: 결측일 경계 차분은 NaN으로 제외) ──
    cal = _to_calendar(d)
    d_sales = cal["매출"].diff()
    d_ad = cal["광고비"].diff()
    rows.append({"구간": "일별", "관계": "Δ광고비 ↔ Δ매출(당일)",
                 "상관": round(d_ad.corr(d_sales), 3), "표본": int((d_ad.notna() & d_sales.notna()).sum())})
    rows.append({"구간": "일별", "관계": "Δ광고비(전일) ↔ Δ매출(당일)",
                 "상관": round(d_ad.shift(1).corr(d_sales), 3),
                 "표본": int((d_ad.shift(1).notna() & d_sales.notna()).sum())})

    # ── 주별 집계 후 차분(노이즈 완화, 운영 단위) ──
    w = d.set_index("날짜").resample("W").agg(매출=("매출", "sum"),
                                              광고비=("광고비", "sum")).reset_index()
    dw_sales = w["매출"].diff()
    dw_ad = w["광고비"].diff()
    rows.append({"구간": "주별", "관계": "Δ광고비 ↔ Δ매출(당주)",
                 "상관": round(dw_ad.corr(dw_sales), 3), "표본": int(dw_ad.notna().sum())})
    rows.append({"구간": "주별", "관계": "Δ광고비(전주) ↔ Δ매출(당주)",
                 "상관": round(dw_ad.shift(1).corr(dw_sales), 3),
                 "표본": int((dw_ad.shift(1).notna() & dw_sales.notna()).sum())})

    # ── 광고비 크게 늘린 주 vs 줄인 주의 매출 변화 ──
    up = dw_sales[dw_ad > 0]
    down = dw_sales[dw_ad < 0]
    rows.append({"구간": "주별", "관계": "광고비 늘린 주의 평균 매출증감",
                 "상관": round(float(up.mean()), 0) if len(up) else np.nan,
                 "표본": int(len(up))})
    rows.append({"구간": "주별", "관계": "광고비 줄인 주의 평균 매출증감",
                 "상관": round(float(down.mean()), 0) if len(down) else np.nan,
                 "표본": int(len(down))})

    out = pd.DataFrame(rows)
    out.attrs["note"] = (
        "차분=전기 대비 변화. 추세 자동 제거됨. 당기 상관만 +이고 시차(전일/전주)가 0이면 "
        "'같은 시기에 같이 움직일 뿐'(요일·프로모 등 공통요인) -> 광고 고유효과 약함. "
        "시차에서도 +면 '광고 증액이 이후 매출을 밀어올림' 근거. "
        "단, 어느 결과든 '광고 중단 가능'은 소규모 축소 실험으로만 확정(관측데이터 한계).")
    return out


def adspend_diff_regression(df: pd.DataFrame) -> pd.DataFrame:
    """광고비 '차분 회귀': Δ로그매출 ~ Δ광고비(당일)+Δ광고비(전일)+요일+프로모션.
    차분으로 추세·레벨 제거, 회귀로 기울기·유의성·시차 동시 추정(상관보다 의사결정에 직접)."""
    d = df.sort_values("날짜").reset_index(drop=True).copy()
    cal = _to_calendar(d)              # 달력일 재색인 -> 결측일 경계 차분은 NaN
    ylevel = cal["logY"].values if USE_LOG_TARGET else cal["매출"].astype(float).values
    dy = pd.Series(ylevel).diff()                       # Δ(로그)매출 (하루차이만)
    d_ad = cal["광고비"].astype(float).diff()           # Δ광고비(당일)
    d_ad_lag = d_ad.shift(1)                            # Δ광고비(전일)

    X = pd.DataFrame({"Δ광고비_당일": d_ad.values, "Δ광고비_전일": d_ad_lag.values})
    wd = pd.get_dummies(cal["요일명"], prefix="요일").drop(columns=["요일_월"], errors="ignore")
    X = pd.concat([X, wd.astype(float).reset_index(drop=True)], axis=1)
    X["프로모션"] = cal["프로모션"].astype(float).values
    XY = sm.add_constant(X, has_constant="add").astype(float)
    XY["_y"] = dy.values
    XY = XY.dropna()                   # 결측일 경계(NaN) 자동 제외
    yy = XY.pop("_y").values
    m = sm.OLS(yy, XY).fit(cov_type="HAC", cov_kwds={"maxlags": HAC_MAXLAGS})

    rows = []
    for name in ["Δ광고비_당일", "Δ광고비_전일"]:
        coef = float(m.params.get(name, np.nan))
        pv = float(m.pvalues.get(name, np.nan))
        # 로그종속이므로 계수×100 ≈ 광고비1원당 매출 %변화
        rows.append({
            "변수": name,
            "계수": round(coef, 9),
            "광고비1원당_매출%변화": round(coef * 100, 7) if USE_LOG_TARGET else np.nan,
            "표준오차": round(float(m.bse.get(name, np.nan)), 9),
            "t값": round(float(m.tvalues.get(name, np.nan)), 2),
            "p값": round(pv, 4),
            "부호": "양수(+)" if coef > 0 else "음수(−)",
            "유의(p<0.05)": "O" if pv < 0.05 else "",
        })
    out = pd.DataFrame(rows)
    out.attrs["note"] = (
        f"모델: Δln(매출) ~ Δ광고비(당일)+Δ광고비(전일)+요일+프로모션, HAC표준오차. "
        f"표본 {int(m.nobs)}, R²={m.rsquared:.3f}. "
        "차분으로 추세·레벨 제거됨. 당일/전일 계수가 +·유의면 '광고비 증액이 (즉시/하루뒤) "
        "매출을 올린다'는 근거. 로그종속이라 계수×100 ≈ 광고비 1원 증액당 매출 %변화.")
    return out


def weekday_reflection_test(df: pd.DataFrame) -> pd.DataFrame:
    """'위클리(요일) 반영이 미미하다'는 피드백 검증.
    요일 더미 블록의 (1)결합 유의성 F검정, (2)요일 추가로 인한 ΔR², (3)유의한 요일 수,
    (4)요일 최대-최소 매출배수 차이로 '요일이 충분히 반영됐는지' 판정."""
    y = df["logY"].values if USE_LOG_TARGET else df["매출"].astype(float).values

    # 전체 모델(요일 포함) vs 축소 모델(요일 제외). 한글 컬럼명 파싱 회피 위해 numpy로 적합
    Xfull_df, cols = _design_matrix(df)
    Xfull_df = sm.add_constant(Xfull_df, has_constant="add").astype(float)
    names = list(Xfull_df.columns)
    dow_cols = [c for c in names if c.startswith("요일_")]
    dow_idx = [names.index(c) for c in dow_cols]
    Xfull = Xfull_df.values
    res_full = sm.OLS(y, Xfull).fit()
    keep = [i for i in range(len(names)) if i not in dow_idx]
    res_red = sm.OLS(y, Xfull[:, keep]).fit()

    # 요일 블록 결합 F검정 (요일 계수 전부 0인가?) — 제약행렬
    R = np.zeros((len(dow_idx), len(names)))
    for r, j in enumerate(dow_idx):
        R[r, j] = 1.0
    ftest = res_full.f_test(R)
    f_val = float(np.ravel(ftest.fvalue)[0])
    f_p = float(np.ravel(ftest.pvalue)[0])

    dR2 = res_full.rsquared - res_red.rsquared
    pvals_full = res_full.pvalues            # numpy fit -> positional
    n_sig = int(sum(pvals_full[j] < 0.05 for j in dow_idx))

    # 요일 효과 크기: 최대-최소 요일계수 -> 매출 배수
    dow_coef = {c: float(res_full.params[j]) for c, j in zip(dow_cols, dow_idx)}
    dow_coef["요일_월(기준)"] = 0.0
    hi = max(dow_coef.values()); lo = min(dow_coef.values())
    span_mult = float(np.exp(hi - lo)) if USE_LOG_TARGET else np.nan

    verdict = ("충분히 반영됨" if (f_p < 0.05 and dR2 >= 0.03 and n_sig >= 3)
               else ("약하게 반영됨(보강 고려)" if f_p < 0.05 else "거의 반영 안 됨"))
    rows = [
        ("요일 블록 결합 F", round(f_val, 2), "요일 더미 전체가 동시에 0인지 검정"),
        ("결합 F p값", round(f_p, 6), "<0.05면 요일 효과 통계적으로 분명"),
        ("ΔR² (요일 추가 기여)", round(dR2, 4),
         "요일 넣어 설명력이 이만큼 증가. 0.03↑면 의미 있는 기여"),
        ("R² (요일 포함)", round(res_full.rsquared, 4), ""),
        ("R² (요일 제외)", round(res_red.rsquared, 4), ""),
        ("유의한 요일 수", f"{n_sig}/{len(dow_cols)}", "월 기준 대비 유의(p<0.05)한 요일 개수"),
        ("최고-최저 요일 매출배수", round(span_mult, 3) if span_mult == span_mult else "—",
         "가장 높은 요일이 가장 낮은 요일의 몇 배(EXP)"),
        ("판정", verdict, "F유의+ΔR²≥0.03+유의요일≥3 => 충분히 반영됨"),
    ]
    out = pd.DataFrame(rows, columns=["항목", "값", "해석"])
    out.attrs["note"] = (
        "'위클리 반영 미미' 피드백 검증용. 결합 F가 유의하고 ΔR²·매출배수가 크면 요일은 "
        "이미 충분히 반영된 것(보정 불필요). 반대면 요일×월 상호작용 등 보강 검토.")
    return out


def multi_variable_effect(df: pd.DataFrame, vars_list: list) -> pd.DataFrame:
    """여러 변수의 매출 영향을 수준/차분 양쪽으로 한 표에 비교(광고비 vs 숏폼 등).
    수준 상관(원시·추세제거) + 차분 동행(일/주, 당기·시차) + 늘린주/줄인주 매출증감."""
    t = df["t"].values.astype(float)
    def detrend(col):
        x = df[col].values.astype(float)
        b = np.polyfit(t, x, 1)
        return pd.Series(x - (b[0] * t + b[1]))
    dm_sales = detrend("매출")
    d_sales = df["매출"].diff()
    wk = df.set_index("날짜").resample("W").sum(numeric_only=True)
    dws = wk["매출"].diff()

    rows = []
    for v in vars_list:
        if v not in df.columns:
            continue
        dv = df[v].diff()
        dwv = wk[v].diff() if v in wk.columns else pd.Series(dtype=float)
        up = dws[dwv > 0]
        down = dws[dwv < 0]
        rows.append({
            "변수": v,
            "수준_원시상관": round(df["매출"].corr(df[v]), 3),
            "수준_추세제거상관": round(dm_sales.corr(detrend(v)), 3),
            "차분_당일상관": round(dv.corr(d_sales), 3),
            "차분_전일상관": round(dv.shift(1).corr(d_sales), 3),
            "주차분_당주상관": round(dwv.corr(dws), 3) if len(dwv) else np.nan,
            "주차분_전주상관": round(dwv.shift(1).corr(dws), 3) if len(dwv) else np.nan,
            "늘린주_평균매출증감": round(float(up.mean()), 0) if len(up) else np.nan,
            "줄인주_평균매출증감": round(float(down.mean()), 0) if len(down) else np.nan,
        })
    out = pd.DataFrame(rows)
    out.attrs["note"] = (
        "수준_추세제거상관: 장기동력 여부(높으면 성장에 기여). 차분상관: 단기동행(늘리면 같이 늚). "
        "추세제거상관은 0인데 차분상관만 +면 '단기 동행은 하나 장기 성장동력은 아님'(광고비형). "
        "추세제거상관도 +면 '장기 성장과도 동행'(성장동력 후보).")
    return out


def analysis_method_definitions() -> pd.DataFrame:
    """각 분석의 기법·수식·입력·해석을 고정 문서화(부호가 갈리는 이유 추적용)."""
    rows = [
        ("수준 원시상관", "피어슨 상관(레벨)",
         "r = cov(광고비, 매출) / (sd(광고비)·sd(매출))",
         "광고비·매출의 그날 값 그대로",
         "추세 포함. '높은 수준 시점끼리' 관계. 대표님 '매출↑때 광고비↑' 확인용."),
        ("수준 추세제거상관", "추세제거 후 피어슨",
         "각 변수 x를 x=b0+b1·t로 회귀 → 잔차 e=x−x̂ 를 상관",
         "직선추세 뺀 잔차(레벨)",
         "공통 성장추세 제거. +면 장기 성장과 동행(성장동력), 0/−면 추세교란."),
        ("차분 당일상관", "1차 차분 후 피어슨",
         "Δx=x_t−x_{t-1} 변환 후 r(Δ광고비, Δ매출)",
         "전일 대비 변화량",
         "'늘린 날 매출도 늚?' 단기 동행. 추세·레벨 자동 제거."),
        ("차분 시차상관", "차분 + 시차 피어슨",
         "r(Δx_{t-1}, Δy_t)  (전일 증감 → 당일 매출증감)",
         "전기 변화량 vs 당기 변화량",
         "+면 증액이 이후 매출로 이어짐(누적효과). 0이면 즉시 동행만."),
        ("이상탐지 회귀계수", "다중 OLS + HAC 표준오차",
         "ln(매출_t)=b0+b_t·t+Σγ_d·요일_d+b_p·프로모+b_a·광고비_t+ε",
         "수준값, 추세·요일·프로모션 동시통제",
         "b_a가 광고비 계수. 통제 후 부분효과. 역인과·추세교란 섞여 음수 가능. 효과측정 아님."),
        ("VIF", "분산팽창계수",
         "VIF_i = 1/(1−R_i²), R_i²=그 변수를 나머지로 회귀한 설명력",
         "설계행렬 각 변수",
         ">5 주의, >10 심각. 계수 불안정(공선성) 여부."),
        ("이상치 판정", "OLS 예측 ±K·σ 관리한계",
         "이상 = |실측−예측| > K·σ (K=2). 로그모델이라 원단위는 예측×exp(±Kσ)",
         "요일별 σ 적용",
         "예측구간 이탈일. σ는 요일별 잔차 표준편차."),
    ]
    return pd.DataFrame(rows, columns=["분석", "기법", "수식", "입력(재료)", "해석"])
def decompose_sku_on_anomaly(anomaly_df: pd.DataFrame, sku_daily: pd.DataFrame,
                             model_frame: pd.DataFrame) -> pd.DataFrame:
    """이상일별 SKU 기여 분해. SKU 베이스라인=적합구간 해당 SKU 일평균 매출."""
    anom_dates = anomaly_df.loc[anomaly_df["이상치"], "날짜"]
    if anom_dates.empty:
        return pd.DataFrame(columns=["날짜", "방향", "SKU", "매출", "수량",
                                     "SKU기준선", "편차", "당일내비중"])
    base = (sku_daily.groupby("SKU")["매출"].mean().rename("SKU기준선"))
    dir_map = dict(zip(anomaly_df["날짜"], anomaly_df["방향"]))
    rows = []
    for d in anom_dates:
        day = sku_daily[sku_daily["날짜"] == d]
        tot = day["매출"].sum()
        for _, r in day.iterrows():
            sku_base = float(base.get(r["SKU"], np.nan))
            rows.append({
                "날짜": d, "방향": dir_map.get(d, ""), "SKU": r["SKU"],
                "매출": r["매출"], "수량": r["수량"],
                "SKU기준선": sku_base,
                "편차": r["매출"] - sku_base,
                "당일내비중": (r["매출"] / tot) if tot else np.nan,
            })
    out = pd.DataFrame(rows)
    out = out.sort_values(["날짜", "편차"], ascending=[True, False]).reset_index(drop=True)
    return out


def map_activity_on_anomaly(anomaly_df: pd.DataFrame, model_frame: pd.DataFrame
                            ) -> pd.DataFrame:
    """이상일에 수행된 마케팅 활동 매핑(광고비/조회수/발행수 등)."""
    anom = anomaly_df[anomaly_df["이상치"]][["날짜", "방향", "매출", "예측", "z"]]
    act = model_frame[["날짜"] + DAILY_ACTIVITY_COLS + ["프로모션", "행사명"]]
    out = anom.merge(act, on="날짜", how="left").sort_values("날짜")
    return out.reset_index(drop=True)


# ============================================================================
# 7. 교차검증 비교 (OLS vs LightGBM) — 선택
# ============================================================================
def crossval_compare(df: pd.DataFrame, n_splits: int = 5) -> pd.DataFrame:
    """TimeSeriesSplit R2 비교. LightGBM 미설치 시 OLS만 보고."""
    from sklearn.model_selection import TimeSeriesSplit
    from sklearn.metrics import r2_score

    rows = []
    if len(df) < MIN_DAYS_CV:
        return pd.DataFrame([{
            "모델": "(교차검증 생략)", "평균R2": np.nan, "폴드수": 0,
            "비고": f"표본 {len(df)}일 < MIN_DAYS_CV={MIN_DAYS_CV}"}])

    X, _ = _design_matrix(df)
    y = _target(df)                                  # 모델 스케일(로그) 타깃
    mask = np.isfinite(X.astype(float).values).all(axis=1) & np.isfinite(y)
    Xv = X.astype(float).values[mask]
    y = y[mask]
    df_v = df.loc[mask].reset_index(drop=True)
    n_splits = max(2, min(n_splits, len(Xv) // 20))
    tss = TimeSeriesSplit(n_splits=n_splits)
    scale = "로그매출" if USE_LOG_TARGET else "매출(원)"

    # OLS
    ols_r2 = []
    for tr, te in tss.split(Xv):
        m = sm.OLS(y[tr], Xv[tr]).fit()
        ols_r2.append(r2_score(y[te], m.predict(Xv[te])))
    rows.append({"모델": "OLS(베이스라인)", "평균R2": float(np.mean(ols_r2)),
                 "폴드수": n_splits, "폴드별R2": str([round(v, 3) for v in ols_r2]),
                 "비고": f"{scale}/{'lag1+' if USE_LAG1 else ''}추세+요일+프로모+광고비"})

    # LightGBM (선택)
    try:
        import lightgbm as lgb
        feat = df_v[DAILY_ACTIVITY_COLS + ["t", "요일", "프로모션"]].astype(float).values
        lgb_r2 = []
        for tr, te in tss.split(feat):
            model = lgb.LGBMRegressor(n_estimators=300, learning_rate=0.05,
                                      num_leaves=15, min_child_samples=10,
                                      verbose=-1)
            model.fit(feat[tr], y[tr])
            lgb_r2.append(r2_score(y[te], model.predict(feat[te])))
        rows.append({"모델": "LightGBM", "평균R2": float(np.mean(lgb_r2)),
                     "폴드수": n_splits, "비고": "매체 기여도 레이어(선택)"})
    except ImportError:
        rows.append({"모델": "LightGBM", "평균R2": np.nan, "폴드수": 0,
                     "비고": "미설치 -> 비활성"})
    return pd.DataFrame(rows)


# ============================================================================
# 8. SHAP 매체 기여도 — 선택
# ============================================================================
def shap_report(df: pd.DataFrame) -> pd.DataFrame:
    """LightGBM+SHAP 평균 |기여| 보고. 미설치/비활성 시 빈 표."""
    if not ENABLE_SHAP:
        return pd.DataFrame([{"비고": "ENABLE_SHAP=False -> 비활성"}])
    try:
        import lightgbm as lgb
        import shap
    except ImportError:
        return pd.DataFrame([{"비고": "lightgbm/shap 미설치 -> 비활성"}])
    feats = DAILY_ACTIVITY_COLS + ["t", "요일", "프로모션"]
    X = df[feats].astype(float)
    y = df["매출"].astype(float)
    model = lgb.LGBMRegressor(n_estimators=400, learning_rate=0.05,
                              num_leaves=15, min_child_samples=10, verbose=-1)
    model.fit(X, y)
    sv = shap.TreeExplainer(model).shap_values(X)
    imp = np.abs(sv).mean(axis=0)
    return (pd.DataFrame({"특성": feats, "평균|SHAP|": imp})
            .sort_values("평균|SHAP|", ascending=False).reset_index(drop=True))


# ============================================================================
# 9. 오케스트레이션
# ============================================================================
def _analyze_one(daily_sales: pd.DataFrame, sku_daily: pd.DataFrame,
                 daily_act: pd.DataFrame, diag: JoinDiagnostics) -> dict:
    """공용 분석 코어: model_frame -> 적합 -> 탐지 -> SKU분해 -> 진단.
    daily_sales는 '매출' 슬롯에 분석 대상값(매출 또는 수량)이 들어 있어야 함."""
    mf = build_model_frame(daily_sales, daily_act, diag)
    res, sigma, _ = fit_baseline_ols(mf)
    sigma_map = compute_sigmas(res, sigma, mf)
    anomalies = detect_anomalies(res, sigma, mf, sigma_map)
    pred_full, spike_cause = decompose_prediction(res, sigma, mf)
    return {
        "model_frame": mf, "ols": res, "sigma": sigma, "sigma_map": sigma_map,
        "coef": reference_means(res, mf),
        "anomalies": anomalies,
        "sku_decomp": decompose_sku_on_anomaly(anomalies, sku_daily, mf),
        "cv": crossval_compare(mf),
        "sheet_formula": build_sheet_formula(res, sigma, mf, sigma_map),
        "resid_diag": residual_diagnostics(res, sigma),
        "sigma_sens": sigma_sensitivity(res, sigma, mf, sigma_map),
        "pred_spikes": diagnose_pred_spikes(anomalies),
        "pred_decomp": pred_full, "spike_cause": spike_cause,
        "weekday_diag": weekday_diagnostics(res, anomalies, sigma_map),
        "weekday_reflect": weekday_reflection_test(mf),
        "vif": vif_diagnostics(res),
        "correlations": relationship_correlations(mf),
        "ad_increment": adspend_increment_effect(mf),
        "ad_diff_reg": adspend_diff_regression(mf),
        "var_effect": multi_variable_effect(
            mf, ["광고비", "숏폼_발행수", "숏폼_조회수", "숏폼_비용",
                 "마케팅_조회수", "바이럴_1~3위건수"]),
        "method_def": analysis_method_definitions(),
    }


def run_analysis(sales_path: str, daily_path: str,
                 sales_sheet="매출_raw", daily_sheet=0,
                 sigma_k: float | None = None) -> dict:
    # UI/CLI에서 관리한계 배수를 동적으로 지정(미지정 시 모듈 기본값 SIGMA_K 사용).
    # 여러 함수가 모듈 전역 SIGMA_K를 직접 참조하므로 실행 전에 전역을 갱신한다.
    global SIGMA_K
    if sigma_k is not None:
        SIGMA_K = float(sigma_k)
    diag = JoinDiagnostics()
    raw = load_sales_raw(sales_path, sales_sheet)
    daily = load_daily_data(daily_path, daily_sheet, diag)

    matched = filter_target_raw(raw, diag)
    if matched.empty:
        raise ValueError("유리젖병 매칭 0건. 매출_raw 제품명/브랜드 확인 필요.")
    daily_act = filter_target_daily(daily, diag)
    daily_sales, sku_daily = aggregate_daily_sales(matched)

    # ── 매출 레이어 ──
    rev = _analyze_one(daily_sales, sku_daily, daily_act, diag)
    activity = map_activity_on_anomaly(rev["anomalies"], rev["model_frame"])
    shap_tbl = shap_report(rev["model_frame"])
    n_anom = int(rev["anomalies"]["이상치"].sum())
    diag.log(f"[매출] 적합 {len(rev['model_frame'])}일(로그), σ={rev['sigma']:.4f}, "
             f"R2={rev['ols'].rsquared:.3f}, 이상치 {n_anom}일")

    # ── 수량 레이어 (매출_raw 수량 합을 종속변수로. 일간데이터 E열은 미사용=역인과 방지) ──
    ds_q = daily_sales[["날짜"]].copy()
    ds_q["매출"] = daily_sales["수량"].values    # '매출' 슬롯에 수량 투입
    ds_q["수량"] = daily_sales["수량"].values
    sku_q = sku_daily[["날짜", "SKU"]].copy()
    sku_q["매출"] = sku_daily["수량"].values
    sku_q["수량"] = sku_daily["수량"].values
    diag_q = JoinDiagnostics()
    qty = _analyze_one(ds_q, sku_q, daily_act, diag_q)
    n_anom_q = int(qty["anomalies"]["이상치"].sum())
    diag.log(f"[수량] 적합 {len(qty['model_frame'])}일(로그), σ={qty['sigma']:.4f}, "
             f"R2={qty['ols'].rsquared:.3f}, 이상치 {n_anom_q}일")

    return {
        "diag": diag, "activity": activity, "shap": shap_tbl,
        # 매출
        "model_frame": rev["model_frame"], "ols": rev["ols"], "sigma": rev["sigma"],
        "coef": rev["coef"], "anomalies": rev["anomalies"],
        "sku_decomp": rev["sku_decomp"], "cv": rev["cv"],
        "sheet_formula": rev["sheet_formula"], "resid_diag": rev["resid_diag"],
        "sigma_sens": rev["sigma_sens"], "pred_spikes": rev["pred_spikes"],
        "pred_decomp": rev["pred_decomp"], "spike_cause": rev["spike_cause"],
        "weekday_diag": rev["weekday_diag"], "weekday_reflect": rev["weekday_reflect"],
        "vif": rev["vif"],
        "correlations": rev["correlations"], "ad_increment": rev["ad_increment"],
        "ad_diff_reg": rev["ad_diff_reg"],
        "var_effect": rev["var_effect"], "method_def": rev["method_def"],
        # 수량
        "qty": qty,
    }


# ============================================================================
# 10. 엑셀 리포트
# ============================================================================
def build_excel(result: dict, out_path: str) -> str:
    diag: JoinDiagnostics = result["diag"]
    res = result["ols"]

    diag_rows = [
        ("매출_raw 총행수", diag.n_raw_rows_total),
        ("유리젖병 매칭 행수", diag.n_raw_rows_matched),
        ("매칭 distinct 제품수", len(diag.raw_matched_products)),
        ("의심 매칭 제품수", len(diag.raw_suspect_products)),
        ("매출 일수", diag.n_dates_sales),
        ("일간데이터 대상 일수", diag.n_dates_daily),
        ("조인 일수", diag.n_dates_joined),
        ("광고비 결측->0 일수", diag.n_dates_adspend_filled0),
        ("일간 헤더 정합성", "정상" if diag.daily_header_ok else "불일치(확인필요)"),
        ("σ(잔차표준편차)", round(result["sigma"], 1)),
        ("R2", round(float(res.rsquared), 4)),
        ("이상치 일수", int(result["anomalies"]["이상치"].sum())),
    ]
    df_diag = pd.DataFrame(diag_rows, columns=["항목", "값"])
    df_matched = pd.DataFrame({"매칭_제품명": diag.raw_matched_products})
    df_matched["의심여부"] = df_matched["매칭_제품명"].isin(diag.raw_suspect_products).map(
        {True: "의심", False: ""})

    # 차트용 데이터: 실측/예측/상한/하한 + 이상일 마커 + 월라벨(매월 1일만 표기)
    an = result["anomalies"]
    chart_df = an[["날짜", "매출", "예측", "상한", "하한"]].copy()
    chart_df["이상일매출"] = np.where(an["이상치"].values, an["매출"].values, np.nan)

    qa = result["qty"]["anomalies"]
    qty_chart_df = qa[["날짜", "매출", "예측", "상한", "하한"]].copy()
    qty_chart_df.columns = ["날짜", "수량", "예측", "상한", "하한"]
    qty_chart_df["이상일수량"] = np.where(qa["이상치"].values, qa["매출"].values, np.nan)

    def _month_labels(dates: pd.Series) -> list:
        """매월 첫 데이터일에만 'YYYY-MM', 나머지는 빈 문자열."""
        d = pd.to_datetime(dates).reset_index(drop=True)
        ym = d.dt.strftime("%Y-%m")
        labels, prev = [], None
        for m in ym:
            labels.append(m if m != prev else "")
            prev = m
        return labels

    def _add_trend_chart(data_ws, target_ws, n, title, ylabel, label_col, anchor):
        """카테고리축 라인차트 + 이상일 마커. 월라벨 가로표시. data_ws 데이터를 target_ws에 앵커."""
        from openpyxl.chart import LineChart, Reference
        from openpyxl.chart.marker import Marker
        from openpyxl.chart.text import RichText
        from openpyxl.drawing.text import (RichTextProperties, Paragraph,
                                            ParagraphProperties, CharacterProperties)
        chart = LineChart()
        chart.title = title
        chart.height, chart.width = 9, 30
        chart.y_axis.title = ylabel
        chart.y_axis.delete = False
        chart.x_axis.title = "날짜(월)"
        chart.x_axis.delete = False
        chart.x_axis.tickLblPos = "low"
        # x축 라벨 가로 고정(세로로 쪼개지는 현상 방지)
        chart.x_axis.txPr = RichText(
            bodyPr=RichTextProperties(rot=0, vert="horz"),
            p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties()),
                         endParaRPr=CharacterProperties())])
        data = Reference(data_ws, min_col=2, max_col=6, min_row=1, max_row=n + 1)
        cats = Reference(data_ws, min_col=label_col, min_row=2, max_row=n + 1)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        if len(chart.series) >= 5:
            s = chart.series[4]
            s.graphicalProperties.line.noFill = True
            s.marker = Marker(symbol="circle", size=7)
        chart.series[0].graphicalProperties.line.width = 8000
        target_ws.add_chart(chart, anchor)

    # ── 통합 이상일 비교표 (날짜 기준: 매출/수량 중 하나라도 이상인 날) ──
    m = an[["날짜", "매출", "예측", "이탈%", "이상치"]].rename(
        columns={"예측": "매출예측", "이탈%": "매출이탈%", "이상치": "매출이상"})
    qn = qa[["날짜", "매출", "예측", "이탈%", "이상치"]].rename(
        columns={"매출": "수량", "예측": "수량예측", "이탈%": "수량이탈%", "이상치": "수량이상"})
    comp = m.merge(qn, on="날짜", how="outer").sort_values("날짜")
    comp = comp[comp["매출이상"].fillna(False) | comp["수량이상"].fillna(False)].copy()
    def _gubun(r):
        if r["매출이상"] and r["수량이상"]:
            return "매출+수량"
        return "매출만" if r["매출이상"] else "수량만"
    comp["구분"] = comp.apply(_gubun, axis=1)
    comp["매출이상"] = comp["매출이상"].map({True: "O", False: ""})
    comp["수량이상"] = comp["수량이상"].map({True: "O", False: ""})
    comp = comp[["날짜", "구분", "매출", "매출예측", "매출이탈%", "매출이상",
                 "수량", "수량예측", "수량이탈%", "수량이상"]].round(
        {"매출": 0, "매출예측": 0, "수량": 0, "수량예측": 0})

    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        # 00 대시보드: 통합표 + 매출/수량 차트(위아래)
        comp.to_excel(xw, sheet_name="00_대시보드", index=False, startrow=1)
        dash = xw.sheets["00_대시보드"]
        dash["A1"] = "■ 통합 이상일 (매출/수량 중 하나라도 ±2σ 이탈). 차트는 우측 위=매출, 아래=수량."

        # 차트 원본 데이터(별도 헬퍼 시트)
        chart_df["월라벨"] = _month_labels(chart_df["날짜"])
        chart_df.to_excel(xw, sheet_name="_매출차트데이터", index=False)
        qty_chart_df["월라벨"] = _month_labels(qty_chart_df["날짜"])
        qty_chart_df.to_excel(xw, sheet_name="_수량차트데이터", index=False)

        _add_trend_chart(xw.sheets["_매출차트데이터"], dash, len(chart_df),
                         "일별 매출: 실측 vs 예측 (±2σ) · 이상일 강조", "매출(원)", 7, "L2")
        _add_trend_chart(xw.sheets["_수량차트데이터"], dash, len(qty_chart_df),
                         "일별 수량: 실측 vs 예측 (±2σ) · 이상일 강조", "수량(개)", 7, "L21")
        xw.sheets["_매출차트데이터"].sheet_state = "hidden"
        xw.sheets["_수량차트데이터"].sheet_state = "hidden"

        df_diag.to_excel(xw, sheet_name="01_조인진단", index=False)
        df_matched.to_excel(xw, sheet_name="02_매칭제품명", index=False)
        result["coef"].to_excel(xw, sheet_name="03_OLS계수", index=False)
        result["anomalies"].to_excel(xw, sheet_name="04_일별_실측vs예측", index=False)
        (result["anomalies"][result["anomalies"]["이상치"]]
         .to_excel(xw, sheet_name="05_이상일요약", index=False))
        result["sku_decomp"].to_excel(xw, sheet_name="06_이상일_SKU분해", index=False)
        result["activity"].to_excel(xw, sheet_name="07_이상일_활동매핑", index=False)
        result["cv"].to_excel(xw, sheet_name="08_교차검증비교", index=False)
        result["shap"].to_excel(xw, sheet_name="09_SHAP기여도", index=False)
        result["sheet_formula"].to_excel(xw, sheet_name="10_구글시트_수식", index=False)
        result["resid_diag"].to_excel(xw, sheet_name="11_잔차진단", index=False)
        result["sigma_sens"].to_excel(xw, sheet_name="12_시그마민감도", index=False)
        result["pred_spikes"].to_excel(xw, sheet_name="13_예측급변일_원인", index=False)
        result["spike_cause"].to_excel(xw, sheet_name="18_봉우리_주원인", index=False)
        result["pred_decomp"].to_excel(xw, sheet_name="19_예측분해_전체", index=False)
        result["weekday_diag"].to_excel(xw, sheet_name="20_요일진단", index=False)
        result["weekday_reflect"].to_excel(xw, sheet_name="20b_요일반영검증", index=False)
        result["vif"].to_excel(xw, sheet_name="21_다중공선성VIF", index=False)
        result["correlations"].to_excel(xw, sheet_name="22_상관관계", index=False)
        result["ad_increment"].to_excel(xw, sheet_name="23_광고증액효과", index=False)
        result["ad_diff_reg"].to_excel(xw, sheet_name="23b_광고비_차분회귀", index=False)
        result["var_effect"].to_excel(xw, sheet_name="24_변수효과비교", index=False)
        result["method_def"].to_excel(xw, sheet_name="25_분석방법_정의", index=False)

        # ── 수량 레이어 시트 ──
        q = result["qty"]
        q_an = q["anomalies"].drop(columns=["수량"]).rename(columns={"매출": "수량"})
        q_an.to_excel(xw, sheet_name="14_수량_일별", index=False)
        q_an[q_an["이상치"]].to_excel(xw, sheet_name="15_수량_이상일", index=False)
        (q["sku_decomp"].drop(columns=["수량"], errors="ignore")
         .rename(columns={"매출": "수량", "SKU기준선": "수량기준선"})
         .to_excel(xw, sheet_name="16_수량_SKU분해", index=False))
        q["resid_diag"].to_excel(xw, sheet_name="17_수량_잔차진단", index=False)
    print(f"[리포트] 저장 완료: {out_path}")
    return out_path


# ============================================================================
# 합성 데이터 (헤드리스 검증용)
# ============================================================================
def _make_synthetic(sales_path: str, daily_path: str, seed: int = 7):
    """실제 시트 레이아웃(매출_raw 1행 / 일간데이터 2행 헤더)대로 합성 파일 생성."""
    rng = np.random.default_rng(seed)
    dates = pd.date_range("2025-01-01", "2026-03-31", freq="D")
    skus = ["유리젖병160_2P", "유리젖병240_2P", "유리젖병160_1P"]
    promo = set()
    for _, s, e in PROMO_SCHEDULE:
        for d in pd.date_range(s, e):
            promo.add(d.normalize())

    rows = []
    for i, d in enumerate(dates):
        wd = d.weekday()
        base = 800_000 + 600 * i + (250_000 if wd >= 5 else 0)   # 추세+주말
        if d.normalize() in promo:
            base *= 2.4
        if rng.random() < 0.03:                                  # 이상치 주입
            base *= rng.choice([0.25, 3.0])
        n_tx = max(1, int(rng.poisson(6) * (2 if d.normalize() in promo else 1)))
        for _ in range(n_tx):
            sku = rng.choice(skus)
            qty = int(rng.integers(1, 4))
            amt = max(0, (base / n_tx) * rng.uniform(0.7, 1.3))
            rows.append(["란시노", "란시노 " + sku.replace("_", " ") + " 역류방지 젖병",
                         sku, d, 19900, 9000, round(amt), qty, "acct_a"])
        # 노이즈: 다른 브랜드/제품
        rows.append(["기타브랜드", "기타 PP젖병 세트", "PP_X", d, 9900, 4000,
                     round(rng.uniform(10000, 50000)), 1, "acct_b"])
    raw = pd.DataFrame(rows, columns=["브랜드", "제품", "SKU", "결제일자",
                                      "SKU판매가", "원가", "실결제금액", "수량", "계정별칭"])
    raw = raw[["결제일자", "브랜드", "제품", "SKU", "SKU판매가", "원가",
               "실결제금액", "수량", "계정별칭"]]
    raw.to_excel(sales_path, sheet_name="매출_raw", index=False)

    # 일간데이터: 2행 헤더 구조로 직접 기록
    ncol = 31
    arr = [[None] * ncol for _ in range(2 + len(dates))]
    # 1행 그룹
    arr[0][0] = "구분"; arr[0][4] = "전체"; arr[0][7] = "paid"
    arr[0][10] = "paid 제외"; arr[0][14] = "마케팅"; arr[0][18] = "바이럴 상세"
    arr[0][22] = "숏폼 상세"; arr[0][26] = "기획사이다 상세"
    # 2행 세부
    hdr2 = ["날짜", "브랜드", "제품", "매출", "판매수량", "비용", "ROAS",
            "전환값", "판매수량", "광고비", "매출", "판매수량", "전주비교", "전월비교",
            "매출", "판매수량", "조회수", "비용", "1~3위 건수", "1~3위 조회수",
            "판매수량", "비용", "발행수", "조회수", "판매수량", "비용",
            "발행수", "조회수", "판매수량", "비용", "오가닉판매수량"]
    arr[1][:len(hdr2)] = hdr2
    for k, d in enumerate(dates):
        r = arr[2 + k]
        r[0] = d; r[1] = "란시노"; r[2] = "유리젖병"
        adspend = max(0, rng.normal(150_000, 40_000))
        if d.normalize() in promo:
            adspend *= 2.0
        r[9] = round(adspend)                  # J 광고비
        r[16] = round(rng.uniform(5000, 20000))   # Q 마케팅 조회수
        r[17] = round(rng.uniform(0, 30000))      # R 마케팅 비용
        r[18] = int(rng.integers(0, 3))           # S 바이럴 1~3위건수
        r[19] = round(rng.uniform(0, 50000))      # T 바이럴 조회수
        r[21] = round(rng.uniform(0, 20000))      # V 바이럴 비용
        r[22] = int(rng.integers(0, 4))           # W 숏폼 발행수
        r[23] = round(rng.uniform(0, 80000))      # X 숏폼 조회수
        r[25] = round(rng.uniform(0, 15000))      # Z 숏폼 비용
        r[26] = int(rng.integers(0, 2))           # AA 사이다 발행수
        r[27] = round(rng.uniform(0, 30000))      # AB 사이다 조회수
        r[29] = round(rng.uniform(0, 10000))      # AD 사이다 비용
    pd.DataFrame(arr).to_excel(daily_path, sheet_name="시트1",
                               index=False, header=False)
    print(f"[합성] 생성: {sales_path}, {daily_path}")


# ============================================================================
# CLI
# ============================================================================
def main():
    ap = argparse.ArgumentParser(description="매출 이상치 탐지 엔진(헤드리스)")
    ap.add_argument("--sales", help="매출_raw 엑셀 경로")
    ap.add_argument("--daily", help="일간데이터 엑셀 경로")
    ap.add_argument("--sales-sheet", default="매출_raw",
                    help="매출_raw 탭 이름(기본: 매출_raw)")
    ap.add_argument("--daily-sheet", default="0",
                    help="일간데이터 탭 이름 또는 인덱스(기본: 0=첫 번째 탭)")
    ap.add_argument("--out", default="anomaly_report.xlsx", help="리포트 출력 경로")
    ap.add_argument("--synthetic", action="store_true",
                    help="합성 데이터로 엔드투엔드 검증")
    args = ap.parse_args()

    if args.synthetic:
        sales_p, daily_p = "_syn_sales.xlsx", "_syn_daily.xlsx"
        _make_synthetic(sales_p, daily_p)
        args.sales, args.daily = sales_p, daily_p
        args.sales_sheet, args.daily_sheet = "매출_raw", "시트1"
    if not args.sales or not args.daily:
        ap.error("--sales 와 --daily 가 필요합니다(또는 --synthetic).")

    # 탭 지정: 숫자면 인덱스, 아니면 시트명
    daily_sheet = int(args.daily_sheet) if args.daily_sheet.isdigit() else args.daily_sheet
    result = run_analysis(args.sales, args.daily,
                          sales_sheet=args.sales_sheet, daily_sheet=daily_sheet)
    build_excel(result, args.out)
    print("[완료]")


if __name__ == "__main__":
    main()
