/*Общий обзор запроса
Этот SQL-запрос реализует стандартный анализ ABC (по объёмам продаж, выручке и прибыли) и XYZ (по стабильности продаж) на уровне товаров. Структура построена через три CTE:

mt – агрегирует суммарные продажи, выручку и прибыль.

weeks – рассчитывает недельные продажи.

xyz_t – на основе недельных продаж классифицирует товары по степени вариативности (X, Y, Z).

В финальном SELECT происходит объединение ABC-классификаций (по объёму, прибыли, выручке) и присоединение XYZ-тегов
*/

WITH mt AS (
    SELECT 
        dr_ndrugs AS product, 
        SUM(dr_kol) AS sales, 
        SUM(dr_kol * dr_croz - dr_sdisc) AS revenue_abc, 
        SUM(dr_kol * (dr_croz - dr_czak) - dr_sdisc) AS profit_abc,
	avg(dr_kol)
    FROM sales s
    GROUP BY product
),
weeks as (
select dr_ndrugs as product, sum(dr_kol) as sales,  to_char(dr_dat, 'YYYY-WW') as yw
from sales s
group by product, yw),
xyz_t as(
select product, case
	when stddev_samp(sales)/avg(sales)>0.25 then 'Z'
	when stddev_samp(sales)/avg(sales)<=0.1 then 'X'
	when stddev_samp(sales)/avg(sales) between 0.1 and 0.25 then 'Y'
	else Null
end as xyz
from weeks
group by product
having count(yw)>=4--distinct and =????
)
SELECT 
    xyz.product, 
    CASE
        WHEN SUM(sales) OVER (ORDER BY sales DESC) / SUM(sales) OVER () <= 0.8 THEN 'A'
        WHEN SUM(sales) OVER (ORDER BY sales DESC) / SUM(sales) OVER () <= 0.95 THEN 'B'
        ELSE 'C'
    END AS abc_amount,
    CASE
        WHEN SUM(profit_abc) OVER (ORDER BY profit_abc DESC) / SUM(profit_abc) OVER () <= 0.8 THEN 'A'
        WHEN SUM(profit_abc) OVER (ORDER BY profit_abc DESC) / SUM(profit_abc) OVER () <= 0.95 THEN 'B'
        ELSE 'C'
    END AS profit_abc,
    CASE
        WHEN SUM(revenue_abc) OVER (ORDER BY revenue_abc DESC) / SUM(revenue_abc) OVER () <= 0.8 THEN 'A'
        WHEN SUM(revenue_abc) OVER (ORDER BY revenue_abc DESC) / SUM(revenue_abc) OVER () <= 0.95 THEN 'B'
        ELSE 'C'
    END AS revenue_abc/*,
    CASE
        WHEN STDDEV(sales / AVG(sales) ) <= 0.1 THEN 'X'
        WHEN STDDEV(sales / AVG(sales) ) <= 0.25 THEN 'Y'
        ELSE 'Z'
    END AS xyz_sales*/, xyz
FROM mt
left join xyz_t xyz on xyz.product=mt.product
ORDER BY product;