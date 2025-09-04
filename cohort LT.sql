with cohorts as (
select
	id as user_id,
	DATE_TRUNC('month',
	date_joined) as cohort_month
from
	users
),
last_activity as (
select
	user_id,
	MAX(entry_at) as last_entry_at
from
	userentry
group by
	user_id
)
select
	to_char(c.cohort_month,
	'mm-yyyy') as my,
	ROUND(AVG(extract(EPOCH from (la.last_entry_at - u.date_joined)) / 86400)::numeric,
	2) as avg_lt_days
from
	cohorts c
join users u on
	u.id = c.user_id
join last_activity la on
	la.user_id = c.user_id
where
	la.last_entry_at >= u.date_joined
group by
	my
order by
	my