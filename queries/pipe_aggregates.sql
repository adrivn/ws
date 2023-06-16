-- TODO: El fichero debe traer información de:
-- 1. Ventas actuales
-- 2. Compromisos
-- 3. Ofertas antiguas sobre las mismas promos/urs

select
    o.offerid,
    string_agg(distinct o.ur_current::int, ', ') as unique_urs,
    string_agg(distinct o.commercialdev::int, ', ') as commercialdev,
    string_agg(distinct o.jointdev::int, ', ') as jointdev,
    string_agg(distinct o.offerstatus, ', ') as offer_status,
    max(s.saledate) as actual_sale_date,
    max(s.commitmentdate) as commitment_date,
    sum(m.lsev_dec19) as lsev_offer,
    sum(m.ppa) as ppa_offer,
    string_agg(distinct m.direccion_territorial, '|') as dts
from offers_data as o
left join master_tape as m
    on o.ur_current = m.ur_current
left join sales2023 as s
    on o.ur_current = s.ur_5000
group by all;
