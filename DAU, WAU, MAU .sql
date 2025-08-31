    SELECT date_trunc('day', entry_at) as day, COUNT(DISTINCT user_id) as dau_count
FROM userentry
GROUP BY day
ORDER BY day;

WITH dau AS (
    SELECT date_trunc('day', entry_at) as day, COUNT(DISTINCT user_id) as dau_count
    FROM userentry
    GROUP BY day
), wau AS (
    SELECT date_trunc('week', entry_at) as week_start, COUNT(DISTINCT user_id) as wau_count
    FROM userentry
    GROUP BY week_start
), mau AS (
    SELECT date_trunc('month', entry_at) as month, COUNT(DISTINCT user_id) as MAU
    FROM userentry
    GROUP BY month
)
SELECT 
      round( AVG(dau_count)) as DAU, round( AVG(wau_count)) as WAU,    round( AVG(MAU)) as MAU
FROM mau, wau, dau