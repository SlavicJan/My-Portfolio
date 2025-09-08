--N-day retention 

with user_entry_and_register_info as (
  -- Формируем сводную таблицу с информацией о посещениях,
  -- датой регистрации и разницей в днях
  select
    u.user_id,
    date(u.entry_at) as entry_at,
    date(u2.date_joined) as date_joined,
    extract(days from u.entry_at - u2.date_joined) as diff,
    to_char(u2.date_joined, 'MM') as cohort
  from userentry u
  join users u2
  on u.user_id = u2.id
  where to_char(u2.date_joined, 'YYYY') = '2022'
),
count_users_by_cohort_and_activity as (
  -- Считаем количество пользователей с активностью
  -- в N-дни с разбивкой по когортам
  select
      cohort,
       diff,
       count(distinct user_id) as cnt
  from user_entry_and_register_info
  where diff in (0, 1, 3, 7, 14, 30)
  group by cohort, diff
)
-- Собираем результат в сводную wide table
select
  to_char(to_date(cohort, 'MM'), 'Month') as cohort,
  max(case when diff = 0 then cnt end) as "Day 0",
  max(case when diff = 1 then cnt end) as "Day 1",
  max(case when diff = 3 then cnt end) as "Day 3",
  max(case when diff = 7 then cnt end) as "Day 7",
  max(case when diff = 14 then cnt end) as "Day 14",
  max(case when diff = 30 then cnt end) as "Day 30"
from count_users_by_cohort_and_activity
group by cohort;

–Более правильных подход или расчет (снижает размер таблицы сразу)

with a as (
  select
    to_char(u2.date_joined, 'MM') as cohort,
    extract(days from u.entry_at - u2.date_joined) as diff,
    count(distinct u.user_id) as cnt
  from userentry u
  join users u2
  on u.user_id = u2.id
  where to_char(u2.date_joined, 'YYYY') = '2022'
  and extract(days from u.entry_at - u2.date_joined) in (0, 1, 3, 7, 14, 30)
  group by cohort, diff
)
select
  to_char(to_date(cohort, 'MM'), 'Month') as cohort,
  max(case when diff = 0 then cnt end) as "Day 0",
  max(case when diff = 1 then cnt end) as "Day 1",
  max(case when diff = 3 then cnt end) as "Day 3",
  max(case when diff = 7 then cnt end) as "Day 7",
  max(case when diff = 14 then cnt end) as "Day 14",
  max(case when diff = 30 then cnt end) as "Day 30"
from a
group by cohort