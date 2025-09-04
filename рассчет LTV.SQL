with cohorts as (
   select
       u.id as user_id,
       date_trunc('month', u.date_joined) as cohort_month
   from users u
),
revenue as (
   select
       t.user_id,
       sum(t.value) as total_revenue
   from transaction t
   where type_id in (1,23,24,25,26,27,28)
   group by t.user_id
)
select
   to_char(c.cohort_month, 'mm-yyyy') cohort_month,
   round(avg(coalesce(r.total_revenue, 0)), 2) as avg_ltv
from cohorts c
left join revenue r on c.user_id = r.user_id
group by cohort_month
order by cohort_month


