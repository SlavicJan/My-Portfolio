--NEW FIXED ABC + XYZ ANALYSIS 
WITH 
-- 1. Агрегаты для ABC
mt AS (
    SELECT 
        dr_ndrugs                AS product,  
        SUM(dr_kol)              AS sales,  
        SUM(dr_kol * dr_croz - dr_sdisc)  
                                AS revenue_abc,  
        SUM(dr_kol * (dr_croz - dr_czak) - dr_sdisc)  
                                AS profit_abc  
    FROM sales
    GROUP BY product
),

-- 2. Недельные продажи для XYZ
weeks AS (
    SELECT 
        dr_ndrugs                AS product,
        SUM(dr_kol)              AS sales,
        TO_CHAR(dr_dat, 'YYYY-WW') AS yw,
        TO_CHAR(dr_dat, 'YYYY-MM') AS ym
    FROM sales
    GROUP BY product, yw
),

-- 3. Классификация по стабильности (XYZ)
xyz_t AS (
    SELECT 
        product,
        CASE
            WHEN STDDEV_SAMP(sales)/AVG(sales) > 0.25 THEN 'Z'
            WHEN STDDEV_SAMP(sales)/AVG(sales) <= 0.10 THEN 'X'
            WHEN STDDEV_SAMP(sales)/AVG(sales) BETWEEN 0.10 AND 0.25 THEN 'Y'
            ELSE NULL
        END AS xyz
    FROM weeks
    GROUP BY product
    HAVING 
        COUNT(DISTINCT yw) >= 4       -- минимум 4 разных недели
        AND COUNT(DISTINCT yw) < 7    -- исключаем товары, продаваемые все недели подряд
)

-- 4. ABC + XYZ итог
SELECT 
    mt.product,
    CASE
        WHEN SUM(mt.sales)  OVER (ORDER BY mt.sales  DESC)
           / SUM(mt.sales)  OVER () <= 0.80 THEN 'A'
        WHEN SUM(mt.sales)  OVER (ORDER BY mt.sales  DESC)
           / SUM(mt.sales)  OVER () <= 0.95 THEN 'B'
        ELSE 'C'
    END AS abc_amount,
    CASE
        WHEN SUM(mt.profit_abc) OVER (ORDER BY mt.profit_abc DESC)
           / SUM(mt.profit_abc) OVER () <= 0.80 THEN 'A'
        WHEN SUM(mt.profit_abc) OVER (ORDER BY mt.profit_abc DESC)
           / SUM(mt.profit_abc) OVER () <= 0.95 THEN 'B'
        ELSE 'C'
    END AS abc_profit,
    CASE
        WHEN SUM(mt.revenue_abc) OVER (ORDER BY mt.revenue_abc DESC)
           / SUM(mt.revenue_abc) OVER () <= 0.80 THEN 'A'
        WHEN SUM(mt.revenue_abc) OVER (ORDER BY mt.revenue_abc DESC)
           / SUM(mt.revenue_abc) OVER () <= 0.95 THEN 'B'
        ELSE 'C'
    END AS abc_revenue,
    xyz_t.xyz
FROM mt
LEFT JOIN xyz_t ON xyz_t.product = mt.product
ORDER BY mt.product
