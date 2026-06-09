# -*- coding: utf-8 -*-
"""
report_business.py
==================
결과물1: 보고용 리포트(xlsx). **판매수량 우선 / 매출 보조** 버전.

목적
----
- 판매수량 예측 및 이상치 판정 결과를 '보고'에 바로 쓸 수 있게 압축(수량=주지표).
- 이상일에 어떤 변수(광고비/프로모션/요일/추세)가 그날 예측을 끌어올렸는지(변수 역할),
  그리고 어떤 활동/이슈가 있었는지 한 시트에서 확인.
- 시각화: 상·하한을 '그라데이션 밴드'로 표현(범위를 시각적으로 전달).
- 매출은 '보조요약' 시트로 수량과 비교(수량 정상인데 매출만 이탈 = 단가/할인 이슈).

수량을 우선으로 두는 이유
------------------------
매출은 가격·할인·환율에 흔들리지만 수량은 순수 수요 신호라 이상일·성장동력 판단에 더 깨끗함.
(매출↔수량 상관이 매우 높아 결론은 거의 동일하나, 수량 기준이 해석상 더 정확)

시트 구성(primary='수량')
------------------------
00_요약            : 대상·기간·R²·σ·관리한계 K·이상치 건수·핵심 메시지(+매출 보조)
01_일별_수량현황    : 날짜/요일/수량/예측/상한/하한/이탈%/이상여부/방향/프로모션
02_베이스라인_계수  : 변수·계수·p값·유의·해석(광고비/프로모션/요일/추세)
03_수량_밴드차트    : 실측 vs 예측 ±Kσ(그라데이션 밴드) + 이상일 마커
04_이상일_요약      : 이상일별 방향·이탈%·주원인·변수기여·활동/이슈
05_이상일_SKU분해   : 이상일별 SKU 수량·기준선·편차·당일내비중
06_일별_예측분해    : 전체 일자의 변수별 예측 기여(상수/추세/요일/프로모션/광고비)
07_시그마민감도     : K별 이상일수·실제비율 vs 정규기대
08_다중공선성_VIF   : 광고비 계수 안정성 점검
09_상관_차분_비교   : 수량 우선 상관/추세제거/차분/주차분/차분회귀(매출 상관 보조열)
10_잔차진단         : 정규성·자기상관·등분산·σ
11_신뢰도_결론      : 규칙기반 자동 진단(경고/주의/양호)
12_매출_보조요약    : 매출 이상일 + 수량↔매출 비교(단가/할인 이슈 식별)
"""
from __future__ import annotations

import numpy as np
import pandas as pd

from openpyxl.chart import AreaChart, LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.fill import (GradientFillProperties, GradientStop,
                                   LinearShadeProperties)
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (RichTextProperties, Paragraph,
                                   ParagraphProperties, CharacterProperties)
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

import sales_anomaly_engine as E


# ─────────────────────────────────────────────────────────────────────────
# 보조
# ─────────────────────────────────────────────────────────────────────────
def _month_labels(dates: pd.Series) -> list:
    """매월 첫 데이터일에만 'YYYY-MM', 나머지는 빈 문자열(차트 가로축 정리용)."""
    d = pd.to_datetime(dates).reset_index(drop=True)
    ym = d.dt.strftime("%Y-%m")
    labels, prev = [], None
    for m in ym:
        labels.append(m if m != prev else "")
        prev = m
    return labels


def _relabel_metric(df: pd.DataFrame, metric: str) -> pd.DataFrame:
    """수량 레이어 테이블의 '매출' 슬롯(실제 값=수량)을 metric 이름으로 정리.
    수량 레이어는 '매출'과 '수량' 컬럼이 모두 동일한 수량값이라 중복 '수량'은 제거 후 rename."""
    if metric == "매출":
        return df.copy()
    df = df.copy()
    if "매출" in df.columns and "수량" in df.columns:
        df = df.drop(columns=["수량"])
    df = df.rename(columns={"매출": metric})
    return df


def _coef_interpret(name: str, metric: str = "수량") -> str:
    if name == "const":
        return f"절편(기준레벨): 월요일·추세0·프로모션0·광고비0일 때의 로그{metric}"
    if name == "t":
        return f"추세: 하루 경과당 로그{metric} 변화(>0이면 우상향 성장)"
    if name == "프로모션":
        return f"프로모션 진행일의 로그{metric} 증분(월요일 동일조건 대비)"
    if name == "광고비":
        return f"광고비 1원당 로그{metric} 변화(추세·요일·프로모션 통제 후 부분효과)"
    if name == "lag_logY":
        return f"전일 로그{metric}의 영향(자기상관 흡수항)"
    if name.startswith("요일_"):
        return f"{name[3:]}요일의 월요일 대비 로그{metric} 증분"
    return ""


def _style_header(ws, n_cols: int, row: int = 1):
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True, size=10)
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font, cell.alignment = fill, font, align
    ws.freeze_panes = ws.cell(row=row + 1, column=1)


# ─────────────────────────────────────────────────────────────────────────
# 그라데이션 밴드 차트
# ─────────────────────────────────────────────────────────────────────────
def _add_band_chart(data_ws, target_ws, n, title, ylabel, anchor):
    """상·하한 그라데이션 밴드 + 예측/실측 라인 + 이상일 마커.
    data_ws 컬럼 순서(필수): A=월라벨 B=하한 C=밴드(상한-하한) D=예측 E=실측 F=이상일실측
    """
    area = AreaChart()
    area.grouping = "stacked"
    area.overlap = 100
    area.add_data(Reference(data_ws, min_col=2, min_row=1, max_row=n + 1),
                  titles_from_data=True)   # 하한
    area.add_data(Reference(data_ws, min_col=3, min_row=1, max_row=n + 1),
                  titles_from_data=True)   # 밴드

    s_low = area.series[0]
    s_low.graphicalProperties = GraphicalProperties()
    s_low.graphicalProperties.noFill = True
    s_low.graphicalProperties.line = LineProperties(noFill=True)

    gf = GradientFillProperties()
    gf.gsLst = [GradientStop(pos=0, srgbClr="9DC3E6"),
                GradientStop(pos=50000, srgbClr="EAF2FB"),
                GradientStop(pos=100000, srgbClr="9DC3E6")]
    gf.lin = LinearShadeProperties(ang=5400000, scaled=True)
    s_band = area.series[1]
    s_band.graphicalProperties = GraphicalProperties()
    s_band.graphicalProperties.gradFill = gf
    s_band.graphicalProperties.line = LineProperties(noFill=True)

    line = LineChart()
    line.add_data(Reference(data_ws, min_col=4, min_row=1, max_row=n + 1),
                  titles_from_data=True)   # 예측
    line.add_data(Reference(data_ws, min_col=5, min_row=1, max_row=n + 1),
                  titles_from_data=True)   # 실측
    line.add_data(Reference(data_ws, min_col=6, min_row=1, max_row=n + 1),
                  titles_from_data=True)   # 이상일
    line.series[0].graphicalProperties = GraphicalProperties()
    line.series[0].graphicalProperties.line = LineProperties(
        solidFill="2F5597", prstDash="dash", w=14000)
    line.series[1].graphicalProperties = GraphicalProperties()
    line.series[1].graphicalProperties.line = LineProperties(
        solidFill="404040", w=20000)
    s_an = line.series[2]
    s_an.graphicalProperties = GraphicalProperties()
    s_an.graphicalProperties.line = LineProperties(noFill=True)
    s_an.marker = Marker(symbol="circle", size=8)
    s_an.marker.graphicalProperties = GraphicalProperties()
    s_an.marker.graphicalProperties.solidFill = "C00000"

    cats = Reference(data_ws, min_col=1, min_row=2, max_row=n + 1)
    area.set_categories(cats)
    line.set_categories(cats)
    line.y_axis.axId = area.y_axis.axId

    area += line
    area.title = title
    area.height, area.width = 9.5, 30
    area.y_axis.title = ylabel
    area.y_axis.delete = False
    area.x_axis.title = "날짜(월)"
    area.x_axis.delete = False
    area.x_axis.tickLblPos = "low"
    area.x_axis.txPr = RichText(
        bodyPr=RichTextProperties(rot=0, vert="horz"),
        p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties()),
                     endParaRPr=CharacterProperties())])
    target_ws.add_chart(area, anchor)


def _band_chart_df(an: pd.DataFrame, metric: str) -> pd.DataFrame:
    """밴드차트용 데이터프레임(월라벨/하한/밴드/예측/실측/이상일)."""
    cd = an[["날짜", metric, "예측", "상한", "하한", "이상치"]].copy().sort_values("날짜")
    return pd.DataFrame({
        "월라벨": _month_labels(cd["날짜"]),
        "하한": cd["하한"].values,
        "밴드(상한-하한)": (cd["상한"] - cd["하한"]).values,
        "예측": cd["예측"].values,
        metric: cd[metric].values,
        f"이상일{metric}": np.where(cd["이상치"].values, cd[metric].values, np.nan),
    })


def _add_weekday_band_charts(xw, target_ws, an: pd.DataFrame, metric: str, unit: str):
    """요일(월~일)별 개별 밴드차트를 target_ws에 세로로 쌓아 배치.
    각 요일의 시계열(실측 vs 예측 ±Kσ 밴드 + 이상일 마커)을 따로 보여 요일별 추세 비교."""
    anchor_row = 3
    for wd in E.WEEKDAY_KR:                       # 월~일 순
        sub = an[an["요일명"] == wd]
        if len(sub) < 2:
            continue
        cdf = _band_chart_df(sub, metric)
        dsheet = f"_wd_{wd}"
        cdf.to_excel(xw, sheet_name=dsheet, index=False)
        xw.sheets[dsheet].sheet_state = "hidden"
        _add_band_chart(
            xw.sheets[dsheet], target_ws, len(cdf),
            f"{wd}요일 {metric}: 실측 vs 예측 (±{E.SIGMA_K:.2f}σ 밴드) · 이상일 강조",
            f"{metric}({unit})", f"A{anchor_row}")
        anchor_row += 20                          # 차트 높이(약 9.5행) 간격


# ─────────────────────────────────────────────────────────────────────────
# 이상일 요약(변수 역할 + 활동/이슈)
# ─────────────────────────────────────────────────────────────────────────
def _anomaly_summary(P: dict, metric: str) -> pd.DataFrame:
    an = _relabel_metric(P["anomalies"], metric)
    anom = an[an["이상치"]].copy()
    if anom.empty:
        return pd.DataFrame(columns=["날짜", "방향", metric, "예측", "이탈%", "주원인"])

    dec = P["pred_decomp"].copy()           # 일자별 변수 기여(로그)
    var_terms = ["요일", "프로모션", "광고비"]
    if "전일매출" in dec.columns and dec["전일매출"].abs().sum() > 0:
        var_terms = ["요일", "전일매출", "프로모션", "광고비"]
    var_terms = [t for t in var_terms if t in dec.columns]
    means = {t: float(dec[t].mean()) for t in var_terms}

    contrib = dec[["날짜"] + var_terms].rename(
        columns={t: f"기여_{t}" for t in var_terms})
    m = anom.merge(contrib, on="날짜", how="left")

    mf = P["model_frame"]
    base_issue = [c for c in ["행사명", "광고비"] if c in m.columns]
    extra_src = [c for c in ["숏폼_조회수", "사이다_조회수"]
                 if c in mf.columns and c not in m.columns]
    if extra_src:
        m = m.merge(mf[["날짜"] + extra_src], on="날짜", how="left")
    rename_issue = {"사이다_조회수": "기획사이다_조회수"}
    m = m.rename(columns=rename_issue)
    issue_cols = base_issue + [rename_issue.get(c, c) for c in extra_src]

    def _cause(r):
        devs = {t: float(r.get(f"기여_{t}", 0.0)) - means.get(t, 0.0) for t in means}
        if not devs:
            return ""
        top = max(devs, key=devs.get)
        return f"{top}(평소대비 {np.exp(devs[top]):.2f}배)"
    m["주원인"] = m.apply(_cause, axis=1)

    for t in means:
        m[f"{t}_기여배수"] = np.exp(m[f"기여_{t}"].astype(float) - means[t]).round(3)

    keep = ["날짜", "방향", metric, "예측", "이탈%", "주원인"]
    keep += [f"{t}_기여배수" for t in means]
    keep += issue_cols
    keep = [c for c in keep if c in m.columns]
    out = m[keep].sort_values("날짜").reset_index(drop=True)
    out = out.round({metric: 0, "예측": 0, "이탈%": 1})
    out.attrs["note"] = ("주원인=그날 예측을 평소 대비 가장 크게 끌어올린 변수. "
                         "기여배수=그 변수가 예측을 평소 대비 몇 배로 만들었나(EXP). "
                         "행사명/광고비/숏폼조회수/기획사이다조회수는 그날 실제 활동(이슈 확인용). "
                         "※ 조회수 활동열은 베이스라인 예측에 사용되지 않음(참고용). "
                         "베이스라인 = 추세 + 요일 + 프로모션 + 광고비.")
    return out


# ─────────────────────────────────────────────────────────────────────────
# 신뢰도 자동 결론
# ─────────────────────────────────────────────────────────────────────────
def _get_val(df: pd.DataFrame, label_col: str, label: str, val_col: str):
    hit = df[df[label_col].astype(str).str.contains(label, regex=False, na=False)]
    if hit.empty:
        return None
    try:
        return float(hit.iloc[0][val_col])
    except (TypeError, ValueError):
        return None


def _reliability_conclusions(P: dict, correlations: pd.DataFrame, metric: str) -> pd.DataFrame:
    """신뢰도 낮음/모순(부호 충돌)/가정 위반을 규칙으로 자동 플래그(주지표=metric)."""
    res = P["ols"]
    rd = P["resid_diag"]
    flags = []

    def add(item, value, verdict, meaning, advice):
        flags.append({"점검항목": item, "값": value, "판정": verdict,
                      "의미": meaning, "권고": advice})

    r2 = float(res.rsquared)
    add("설명력 R²", round(r2, 3),
        "참고" if r2 >= 0.5 else "주의(설명력 낮음)",
        f"베이스라인이 {metric} 변동의 일부만 설명. 단 이상치 판정은 '예측 대비 잔차'라 "
        "R²가 낮아도 잔차진단이 양호하면 ±Kσ 판정 자체는 유효.",
        "예측값 신뢰가 중요하면 변수 보강. 이상탐지 목적이면 잔차진단·시그마민감도로 정당성 확인.")

    jb = _get_val(rd, "검정", "Jarque-Bera", "값")
    if jb is not None:
        add("잔차 정규성(JB p)", round(jb, 4),
            "양호" if jb >= 0.05 else "주의",
            "정규에 가까우면 ±Kσ의 '몇 %' 해석 성립. 벗어나면 정규기대 %는 부정확.",
            "K는 시그마민감도의 '실제비율'을 보고 결정(정규기대 맹신 금지).")

    dw = _get_val(rd, "검정", "Durbin-Watson", "값")
    if dw is not None:
        add("잔차 자기상관(DW)", round(dw, 3),
            "양호" if 1.5 <= dw <= 2.5 else "주의",
            "2 근처면 자기상관 없음. 벗어나면 시계열 의존 → σ·표준오차 과소평가 가능.",
            "양호면 조치 불필요. 아니면 HAC 적용 확인 / 전일값 항 재도입 검토.")

    bp = _get_val(rd, "검정", "Breusch-Pagan", "값")
    if bp is not None:
        add("등분산성(BP p)", round(bp, 4),
            "양호" if bp >= 0.05 else "주의",
            f"등분산이면 양호. 이분산이면 고{metric}일 과검출 위험(로그변환으로 완화 중).",
            "양호면 조치 불필요. 아니면 잔차 vs 적합값 산점 확인.")

    vif = P.get("vif")
    if vif is not None and "VIF" in vif.columns and len(vif):
        vser = pd.to_numeric(vif["VIF"], errors="coerce")
        maxv = float(vser.max()); who = vif.iloc[vser.idxmax()]["변수"]
        verdict = "경고(심각)" if maxv > 10 else ("주의" if maxv > 5 else "양호")
        add(f"다중공선성 VIF(최대: {who})", round(maxv, 2), verdict,
            "VIF 높은 변수는 계수 불안정(부호·크기 신뢰 저하). 기준: >5 주의 / >10 심각.",
            "VIF>10이면 해당 변수 계수 해석 보류. 변수 통합/제거 검토.")

    coef = P["coef"]
    raw_col = f"{metric}_상관(원시)"
    ad_coef = _get_val(coef, "변수", "광고비", "계수")
    raw_corr = None
    if (correlations is not None and "변수" in correlations.columns
            and raw_col in correlations.columns):
        hit = correlations[correlations["변수"] == "광고비"]
        if not hit.empty:
            try:
                raw_corr = float(hit.iloc[0][raw_col])
            except (TypeError, ValueError):
                raw_corr = None
    if ad_coef is not None and raw_corr is not None:
        contradiction = (ad_coef < 0 and raw_corr > 0)
        add("광고비: 회귀계수 vs 원시상관",
            f"계수 {ad_coef:.2e} / 원시상관 {raw_corr:+.2f}",
            "경고(모순)" if contradiction else "참고",
            ("통제후 회귀계수는 음수인데 원시상관은 양수 → 추세교란·역인과 가능. "
             f"'광고비가 {metric}을(를) 올린다/내린다'로 단정 불가."
             if contradiction else "회귀계수와 원시상관 부호가 모순되지 않음."),
            "광고비 차분회귀(09 시트) 부호·유의로 재확인. 인과 확정은 소규모 축소실험 필요.")

    ss = P.get("sigma_sens")
    if ss is not None and "현재선택" in ss.columns:
        cur = ss[ss["현재선택"].astype(str).str.contains("현재", na=False)]
        if not cur.empty:
            act = float(cur.iloc[0]["비율(%)"]); exp = float(cur.iloc[0]["정규기대(%)"])
            verdict = "주의" if abs(act - exp) >= max(3.0, 0.5 * exp) else "양호"
            add(f"이상치 실제비율 vs 정규기대(±{E.SIGMA_K:.2f}σ)",
                f"실제 {act:.1f}% / 기대 {exp:.2f}%", verdict,
                "비슷하면 ±Kσ 기준이 데이터에 부합. 크게 다르면 정규가정 부적합.",
                "차이가 크면 정규기대% 대신 '실제비율'로 K를 정할 것.")

    out = pd.DataFrame(flags)
    out.attrs["note"] = ("규칙 기반 자동 진단. '경고'는 우선 확인, '주의'는 해석 시 유의, "
                         "'참고/양호'는 통과. 인과(특히 광고비)는 실험 없이 단정하지 않음. "
                         "코드·수식·기준 상세는 '분석방법_정리.html' 참고.")
    return out


def _corr_diff_block(xw, sheet: str, P: dict, correlations: pd.DataFrame, metric: str):
    """상관·차분 관련 3표를 한 시트에 위아래로 기록(수량 우선)."""
    ws = xw.book.create_sheet(sheet)

    # 상관표: 수량 컬럼을 앞으로, 매출 컬럼은 보조로 뒤에 배치
    corr = correlations.copy() if correlations is not None else None
    if corr is not None and metric != "매출":
        front = ["변수", f"{metric}_상관(원시)", f"{metric}_상관(추세제거)"]
        rest = [c for c in corr.columns if c not in front]
        corr = corr[[c for c in front if c in corr.columns] + rest]

    # var_effect / ad_diff_reg는 수량 레이어 기준(매출→수량 라벨 정리)
    def _relabel_cols(df_):
        if df_ is None or metric == "매출":
            return df_
        return df_.rename(columns={c: c.replace("매출", metric) for c in df_.columns})

    blocks = [("[원시/추세제거 상관]", corr, correlations.attrs.get("note", "") if correlations is not None else ""),
              ("[변수효과 비교(수준/차분)]", _relabel_cols(P.get("var_effect")),
               (P.get("var_effect").attrs.get("note", "") if P.get("var_effect") is not None else "")),
              ("[광고비 차분회귀]", _relabel_cols(P.get("ad_diff_reg")),
               (P.get("ad_diff_reg").attrs.get("note", "").replace("매출", metric)
                if P.get("ad_diff_reg") is not None else ""))]

    r = 1
    for title, df_, note in blocks:
        if df_ is None or len(df_) == 0:
            continue
        ws.cell(row=r, column=1, value=title).font = Font(bold=True, size=11)
        r += 1
        df_.to_excel(xw, sheet_name=sheet, index=False, startrow=r - 1)
        r += len(df_) + 1
        if note:
            ws.cell(row=r, column=1, value="해설:")
            ws.cell(row=r, column=2, value=note)
            r += 2


def _revenue_appendix(result: dict, metric: str) -> pd.DataFrame:
    """매출 보조요약: 수량 vs 매출 이상일 비교(둘 중 하나라도 이탈한 날).
    수량 정상인데 매출만 이탈 = 단가/할인 이슈로 식별."""
    rev = result["anomalies"]                          # 실제 매출
    q = _relabel_metric(result["qty"]["anomalies"], "수량")   # 실제 수량
    rcol = ["날짜", "매출", "예측", "이탈%", "이상치"]
    m_rev = rev[rcol].rename(columns={"예측": "매출예측", "이탈%": "매출이탈%",
                                      "이상치": "매출이상"})
    m_q = q[["날짜", "수량", "예측", "이탈%", "이상치"]].rename(
        columns={"예측": "수량예측", "이탈%": "수량이탈%", "이상치": "수량이상"})
    comp = m_q.merge(m_rev, on="날짜", how="outer").sort_values("날짜")
    comp = comp[comp["수량이상"].fillna(False) | comp["매출이상"].fillna(False)].copy()

    def _gubun(r):
        if r["수량이상"] and r["매출이상"]:
            return "수량+매출"
        return "수량만" if r["수량이상"] else "매출만"
    comp["구분"] = comp.apply(_gubun, axis=1)
    comp["수량이상"] = comp["수량이상"].map({True: "O", False: ""})
    comp["매출이상"] = comp["매출이상"].map({True: "O", False: ""})
    comp = comp[["날짜", "구분", "수량", "수량예측", "수량이탈%", "수량이상",
                 "매출", "매출예측", "매출이탈%", "매출이상"]].round(
        {"수량": 0, "수량예측": 0, "매출": 0, "매출예측": 0})
    comp = comp.reset_index(drop=True)
    comp.attrs["note"] = ("주지표=수량. 이 시트는 매출 관점 보조요약. "
                          "'매출만' 이탈(수량 정상)=단가/할인/믹스 이슈로 매출만 변동. "
                          "'수량만' 이탈(매출 정상)=저가 SKU 위주 수요 변동. "
                          "'수량+매출' 동시 이탈=실제 수요 급변. "
                          "수량·매출 이탈% 부호가 같고 크기 비슷하면 단가 영향 작음.")
    return comp


# ─────────────────────────────────────────────────────────────────────────
# 메인
# ─────────────────────────────────────────────────────────────────────────
def build_business_report(result: dict, out_path: str, primary: str = "수량") -> str:
    """보고용 리포트 생성. primary='수량'(기본)이면 판매수량 우선 + 매출 보조요약,
    primary='매출'이면 매출 우선(레거시)."""
    metric = primary
    unit = "개" if metric == "수량" else "원"

    # 주지표 레이어 선택
    if primary == "수량":
        P = dict(result["qty"])
    else:
        P = {k: result[k] for k in
             ("model_frame", "ols", "sigma", "sigma_map", "coef", "anomalies",
              "sku_decomp", "pred_decomp", "resid_diag", "sigma_sens", "vif",
              "var_effect", "ad_diff_reg") if k in result}
    # 상관표는 항상 top-level(매출·수량 둘 다 정확히 계산된 표) 사용
    correlations = result.get("correlations")

    diag = result["diag"]
    res = P["ols"]
    an = _relabel_metric(P["anomalies"], metric)
    n_up = int(((an["이상치"]) & (an["방향"] == "상회")).sum())
    n_dn = int(((an["이상치"]) & (an["방향"] == "하회")).sum())
    period = f'{an["날짜"].min():%Y-%m-%d} ~ {an["날짜"].max():%Y-%m-%d}'

    # 매출 보조 카운트(요약용)
    rev_an = result["anomalies"]
    n_rev = int(rev_an["이상치"].sum())

    # 00 요약
    summary = pd.DataFrame([
        ("분석 대상", f"{E.TARGET_BRAND} {E.TARGET_PRODUCT_DAILY}"),
        ("주지표", f"판매수량 (매출은 12_매출_보조요약 참조)"),
        ("분석 기간", period),
        ("분석 일수", int(an["날짜"].nunique())),
        ("모델", f"로그{metric} OLS · 추세 + 요일 + 프로모션 + 광고비"
                 + (" + 전일값" if E.USE_LAG1 else "")),
        ("설명력 R²", round(float(res.rsquared), 4)),
        (f"잔차 표준편차 σ(로그)", round(float(P["sigma"]), 4)),
        ("관리한계 배수 K", f"±{E.SIGMA_K:.2f}σ (낮을수록 이상치를 더 많이 잡음)"),
        ("수량 이상일 총건수", int(an["이상치"].sum())),
        ("  ├ 상회(예측보다 높음)", n_up),
        ("  └ 하회(예측보다 낮음)", n_dn),
        ("(보조) 매출 이상일 건수", n_rev),
        ("핵심 메시지",
         f"주지표=판매수량. 전체 {int(an['날짜'].nunique())}일 중 {int(an['이상치'].sum())}일이 "
         f"예측 ±{E.SIGMA_K:.2f}σ 밴드를 벗어남. 이상일별 원인·활동은 04, 시각화는 03, "
         f"변수 영향은 02·06. 신뢰도·모순 점검은 07~11(특히 11_신뢰도_결론). "
         f"매출 관점 비교는 12_매출_보조요약. 코드·수식·기준 상세는 동봉 '분석방법_정리.html'."),
    ], columns=["항목", "값"])

    # 01 일별 수량현황
    daily = an[["날짜", "요일명", metric, "예측", "상한", "하한",
                "이탈%", "이상치", "방향", "프로모션", "행사명"]].copy()
    daily["이상여부"] = daily["이상치"].map({True: "이상", False: ""})
    daily = daily.drop(columns=["이상치"])
    daily = daily.round({metric: 0, "예측": 0, "상한": 0, "하한": 0, "이탈%": 1})

    # 02 계수
    coef = P["coef"].copy()
    coef["배수(EXP)"] = np.exp(coef["계수"].astype(float)).round(4)
    coef["해석"] = coef["변수"].map(lambda n: _coef_interpret(n, metric))
    coef = coef.round({"계수": 5, "표준오차": 5, "t값": 2, "p값": 4})

    # 04 / 05 / 06
    anom_summary = _anomaly_summary(P, metric)
    sku = _relabel_metric(P["sku_decomp"], metric).copy()
    sku_round = {metric: 0, "SKU기준선": 0, "편차": 0, "당일내비중": 3}
    sku = sku.round({k: v for k, v in sku_round.items() if k in sku.columns})
    pred_decomp = P["pred_decomp"].copy()
    if "예측(원)" in pred_decomp.columns and metric != "매출":
        pred_decomp = pred_decomp.rename(columns={"예측(원)": f"예측({unit})"})

    # 차트 데이터(밴드 = 상한 - 하한)
    chart_df = _band_chart_df(an, metric)

    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="00_요약", index=False)
        daily.to_excel(xw, sheet_name=f"01_일별_{metric}현황", index=False)
        coef.to_excel(xw, sheet_name="02_베이스라인_계수", index=False)

        # 03 차트 시트 + 숨김 데이터
        wb = xw.book
        ws_chart = wb.create_sheet(f"03_{metric}_밴드차트")
        ws_chart["A1"] = ("■ 실측 vs 예측 ±Kσ 밴드(그라데이션=예측구간). "
                          "빨강 점=이상일. 점선=예측, 굵은 회색=실측.")
        ws_chart["A1"].font = Font(bold=True, size=11)
        chart_df.to_excel(xw, sheet_name="_차트데이터", index=False)
        ws_data = xw.sheets["_차트데이터"]
        _add_band_chart(ws_data, ws_chart, len(chart_df),
                        f"일별 {metric}: 실측 vs 예측 (±{E.SIGMA_K:.2f}σ 밴드) · 이상일 강조",
                        f"{metric}({unit})", "A3")
        ws_data.sheet_state = "hidden"

        # 03b 요일별 개별 밴드차트(월~일 세로 배치)
        ws_wd = wb.create_sheet("03b_요일별_밴드차트")
        ws_wd["A1"] = ("■ 요일별 밴드차트(월~일). 각 차트는 그 요일만의 시계열 "
                       "실측 vs 예측 ±Kσ 밴드 + 이상일(빨강 점). 요일별 추세·변동성 비교용.")
        ws_wd["A1"].font = Font(bold=True, size=11)
        _add_weekday_band_charts(xw, ws_wd, an, metric, unit)

        anom_summary.to_excel(xw, sheet_name="04_이상일_요약", index=False)
        sku.to_excel(xw, sheet_name="05_이상일_SKU분해", index=False)
        pred_decomp.to_excel(xw, sheet_name="06_일별_예측분해", index=False)

        # ── 신뢰도 레이어 ──
        sigma_sens = P["sigma_sens"]
        vif = P["vif"]
        resid_diag = P["resid_diag"]
        conclusions = _reliability_conclusions(P, correlations, metric)

        sigma_sens.to_excel(xw, sheet_name="07_시그마민감도", index=False)
        vif.to_excel(xw, sheet_name="08_다중공선성_VIF", index=False)
        _corr_diff_block(xw, "09_상관_차분_비교", P, correlations, metric)
        resid_diag.to_excel(xw, sheet_name="10_잔차진단", index=False)
        conclusions.to_excel(xw, sheet_name="11_신뢰도_결론", index=False)

        # 12 매출 보조요약
        rev_appendix = _revenue_appendix(result, metric)
        rev_appendix.to_excel(xw, sheet_name="12_매출_보조요약", index=False)

        # 시트별 노트
        for sht, df_ in [("04_이상일_요약", anom_summary),
                         ("06_일별_예측분해", pred_decomp),
                         ("07_시그마민감도", sigma_sens),
                         ("08_다중공선성_VIF", vif),
                         ("10_잔차진단", resid_diag),
                         ("11_신뢰도_결론", conclusions),
                         ("12_매출_보조요약", rev_appendix)]:
            note = df_.attrs.get("note", "")
            if note:
                ws = xw.sheets[sht]
                ws.cell(row=ws.max_row + 2, column=1, value="해설:")
                ws.cell(row=ws.max_row, column=2, value=note)

        # 헤더 스타일
        for sht, ncol in [(f"01_일별_{metric}현황", daily.shape[1]),
                          ("02_베이스라인_계수", coef.shape[1]),
                          ("04_이상일_요약", anom_summary.shape[1]),
                          ("05_이상일_SKU분해", sku.shape[1]),
                          ("06_일별_예측분해", pred_decomp.shape[1]),
                          ("07_시그마민감도", sigma_sens.shape[1]),
                          ("08_다중공선성_VIF", vif.shape[1]),
                          ("11_신뢰도_결론", conclusions.shape[1]),
                          ("12_매출_보조요약", rev_appendix.shape[1])]:
            if sht in xw.sheets and ncol > 0:
                _style_header(xw.sheets[sht], ncol)

    print(f"[보고용/{metric}우선] 저장 완료: {out_path}")
    return out_path
