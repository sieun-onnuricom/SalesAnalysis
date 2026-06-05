# 매출 이상치 탐지 엔진

란시노 유리젖병 매출을 베이스라인(추세+요일+프로모션+광고비)으로 예측하고,
예측 ±Kσ 밴드를 벗어난 날(이상일)을 탐지·해석하는 도구입니다.

## 구성

| 파일 | 역할 |
|------|------|
| `sales_anomaly_engine.py` | 통계 엔진(로더·OLS·이상치 탐지·진단). UI 의존성 없음 |
| `report_business.py` | **보고용** xlsx(요약·일별현황·계수·그라데이션 밴드차트·이상일·SKU + 시그마민감도·VIF·상관차분·잔차진단·신뢰도결론) |
| `app.py` | **데이터 입력 화면**(Streamlit). 보고용 리포트 1종 다운로드 |
| `분석방법_정리.html` | 코드·기법·수식·판정기준 정리(정적 참고문서, 데이터 무관) |

## 로컬 실행

```bash
pip install -r requirements.txt
streamlit run app.py
```

입력 방식: ① xlsx 업로드  또는  ② GitHub raw URL(공개/무료).
관리한계 K는 슬라이더로 조절(낮을수록 이상치를 더 많이 잡음).

## CLI(엔진 단독)

```bash
# 합성 데이터로 동작 검증
python sales_anomaly_engine.py --synthetic --out report.xlsx

# 실제 파일
python sales_anomaly_engine.py --sales 매출.xlsx --daily 일간.xlsx \
       --daily-sheet 0 --out report.xlsx
```

## 관리한계 K 가이드

정규가정 기준 이탈률: ±1.0σ≈32% · ±1.5σ≈13% · ±2.0σ≈5% · ±2.5σ≈1%.
**실제 검출 수는 데이터에 따라 다르므로**, 검토용 리포트의 `03_시그마민감도`에서
K별 실제 이상일 수를 보고 결정하는 것을 권장합니다(정규기대% 맹신 금지).
