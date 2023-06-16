select
    w.offer_id,
    w.full_path,
    w.unique_urs,
    w.commercialdev,
    w.jointdev,
    count(w.unique_urs) as urs_per_file,
    count(o.ur_current) as urs_per_offers,
    sum(m.lsev_dec19) as lsevdec19,
    sum(m.ppa) as ppa
from ws_hist_offers as w
cross join unnest(w.unique_urs) as f (urs)
left join master_tape as m
    on f.urs = m.ur_current
left join offers_data as o
    on w.offer_id = o.offerid
group by all;
