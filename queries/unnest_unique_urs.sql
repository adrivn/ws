with all_offers as (
  select * from ws_hist_offers w1
  union all
  select * from ws_current_offers w2
),
-- que queremos traer del tape de ofertas semanal??
-- tenemos que traerlo por offer_id?
-- NOTE:
-- no, hay que traer los agregados por offer_id pero no lincar via offer_id, me explico?
-- los conteos de urs_per_offers_table y urs_per_file son IGUALES de modo que todo está ok
-- WARN: Error en ws_hist_offers porque no existe, hay que meter un algo para que solo pille las disponibles,
-- tal vez un regex que pille las ws_
enriched_offers as (
  select  a.*,
          array_agg(DISTINCT o.ur_current) AS all_urs_pot,
          array_agg(DISTINCT o.commercialdev) AS all_commercialdevs_pot,
          array_agg(DISTINCT o.jointdev) AS all_jointdevs_pot,
          array_agg(DISTINCT o.offerstatus) AS offer_status
  from all_offers a
  left join offers o
  on a.offer_id = o.offerid
  group by all
),
final_piece as(select
    w.*,
    l1.unique_urs,
    l1.commercialdevs,
    l1.jointdevs,
    len(l1.unique_urs) as total_urs,
from enriched_offers as w,
    lateral (
        select list_distinct(list_concat(w.unique_urs,w.all_urs_pot)),
               cast(list_distinct(list_concat(w.commercialdev,w.all_commercialdevs_pot)) as int[]),
               cast(list_distinct(list_concat(w.jointdev,w.all_jointdevs_pot)) as int[])
    ) l1(unique_urs, commercialdevs, jointdevs)
group by all)
select  f.unique_id,
        cast(f.offer_id as int) as offer_id,
        f.offer_status[1] as offer_status,
        list_aggregate(f.unique_urs, 'string_agg', ' | ') as unique_urs,
        list_aggregate(f.commercialdevs, 'string_agg', ' | ') as commercialdevs,
        list_aggregate(f.jointdevs, 'string_agg', ' | ') as jointdevs,
        f.total_urs,
        sum(coalesce(m.lsev_dec19,a.lsev_dec19)) as lsevdec19,
        sum(m.ppa) as ppa
from final_piece f
cross join unnest(f.unique_urs) as r (urs)
left join master_tape as m
    on r.urs = m.ur_current
left join allocation_new as a
    on r.urs = a.ur
group by all;
