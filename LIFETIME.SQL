with last_activity as (
select
	user_id,
	MAX(entry_at) as last_entry_at
from
	userentry
group by
	user_id
)
select
	ROUND(AVG(extract(EPOCH from (la.last_entry_at - u.date_joined)) / 86400)::numeric,
	2) as avg_lt_days
from
	users u
join last_activity la on
	la.user_id = u.id
where
	la.last_entry_at >= u.date_joined