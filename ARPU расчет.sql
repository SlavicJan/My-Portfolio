select round(avg(sum),3) from (select user_id, sum(value)
from "transaction" t 
where type_id in (1,23,24,25,26,27,28)
group by user_id) arpu