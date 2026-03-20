-- 목적       : 알바몬 2025년도 월별 발생매출 집계
-- 작성자     : 신건희(사업성장팀) Claude
-- 모수 정의  : 알바몬(code_oem.ws_code = 'AM') 기준 SAP 발생 매출
--              정상 건(_row_status = 1), 취소·환불 제외
-- 기간 정의  : 2025년 전체 (2025년 1월 ~ 12월)
--              SQL 반열린구간: valid_crr_date >= '2025-01-01' AND < '2026-01-01'
-- 지표·차원  : 월별 발생매출 합계(sppl_value), 매출 건수
-- 사용 테이블: box_magma.stt_stats_fnnc_daily_hist_sap (발생매출 기준)
--              box_magma.code_oem                      (알바몬 OEM 식별: ws_code = 'AM')
-- 가정·TODO  : ⚠️ valid_crr_date = SAP 발생(인식)일 컬럼 여부 GQ 정의 우선 확인
--              ⚠️ sppl_value = 공급가액(발생매출) 기준 확인 (vs sales_price 판매가)
--              ⚠️ sap_apply_stat 정상 처리 코드값 확인 후 주석 해제

WITH params AS (
    SELECT
        DATE '2025-01-01' AS start_dt,
        DATE '2026-01-01' AS end_dt
),

-- 알바몬 OEM 번호 목록 (ws_code = 'AM')
albamon_oem AS (
    SELECT oem_no
    FROM box_magma.code_oem
    WHERE ws_code  = 'AM'
      AND del_stat = 0
),

-- 발생매출 기준 원본 (stt_stats_fnnc_daily_hist_sap)
base AS (
    SELECT
        f.valid_crr_date,
        f.sppl_value,      -- 공급가액 (발생매출 기준)
        f.sales_price      -- 판매가 (참고용)
    FROM box_magma.stt_stats_fnnc_daily_hist_sap f
    JOIN albamon_oem a
      ON f.sales_oem_no = a.oem_no
    CROSS JOIN params
    WHERE CAST(f.valid_crr_date AS DATE) >= params.start_dt
      AND CAST(f.valid_crr_date AS DATE) <  params.end_dt
      AND f._row_status = 1
      -- AND f.sap_apply_stat = 1  -- ⚠️ 정상 코드값 확인 후 해제
),

-- 월별 집계
monthly_agg AS (
    SELECT
        DATE_TRUNC('month', CAST(valid_crr_date AS DATE)) AS revenue_month,
        COUNT(*)                                           AS sales_cnt,
        SUM(sppl_value)                                    AS monthly_revenue,
        SUM(sales_price)                                   AS monthly_sales_price
    FROM base
    GROUP BY 1
),

-- 모수 검증
validation AS (
    SELECT
        COUNT(*)        AS total_row_cnt,
        SUM(sppl_value) AS total_revenue_check
    FROM base
)

SELECT
    m.revenue_month,
    m.sales_cnt,
    m.monthly_revenue,
    m.monthly_sales_price,
    SUM(m.monthly_revenue) OVER (
        ORDER BY m.revenue_month
        ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW
    ) AS cumulative_revenue,
    v.total_row_cnt,
    v.total_revenue_check
FROM monthly_agg m
CROSS JOIN validation v
ORDER BY m.revenue_month
