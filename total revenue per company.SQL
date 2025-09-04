select
	c.name as company_name,
	SUM(t.value) as total_revenue
from
	transaction t
join users u on
	t.user_id = u.id
join company c on
	u.company_id = c.id
where
	type_id in (2, 23, 24, 25, 26, 27, 28)
group by
	c.name
order by
	total_revenue desc