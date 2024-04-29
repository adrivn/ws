-- TODO: El fichero debe traer información de:
-- 1. Ventas actuales
-- 2. Compromisos
-- 3. Ofertas antiguas sobre las mismas promos/urs

with offers_expanded as (
    select
        o.*,
        m.direccion_territorial,
        m.ppa,
        m.lsev_dec19
    from offers as o
    left join master_tape as m
        on o.ur_current = m.ur_current
)

select
    o.offerid,
    string_agg(distinct o.ur_current::int, ', ') as unique_urs,
    string_agg(distinct o.commercialdev::int, ', ') as commercialdev,
    string_agg(distinct o.jointdev::int, ', ') as jointdev,
    string_agg(distinct o.offerstatus, ', ') as offer_status,
    sum(s.saleprice) filter (where o.offerid = s.offerid) as sale_price,
    max(s.saledate) filter (where o.offerid = s.offerid) as actual_sale_date,
    max(s.commitmentdate) filter (where o.offerid = s.offerid) as commitment_date,
    sum(o.lsev_dec19) as lsev_offer,
    sum(o.ppa) as ppa_offer,
    string_agg(distinct o.direccion_territorial, '|') as dts
from offers_expanded as o
left join sales2024 as s
    on o.ur_current = s.ur_current
group by all;
