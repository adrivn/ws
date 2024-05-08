with all_offers as (
  select * from ws.ws_offers_2024
  union all 
  select * from ws.ws_offers_2023
  union all 
  select * from ws.ws_offers_2022
  union all 
  select * from ws.ws_offers_2021
  union all 
  select * from ws.ws_offers_2020
  union all 
  select * from ws.ws_offers_2019 
)
select coalesce(f.offer_id, nullif(regexp_extract(f.full_path, '5\d{6}'), '')::int, m.offerid) as offer_id, 
       f.unique_id,
       len(f.unique_urs) as total_urs,
       nullif(count(m.ur_current) filter (where m.updatedcategory = 'Sale Agreed' and m.offerid = f.offer_id),0) as urs_committed,
       nullif(sum(m.commitmentprice) filter (where m.updatedcategory = 'Sale Agreed' and m.offerid = f.offer_id),0) as commitment_amount,
       max(m.commitmentdate) filter (where m.updatedcategory = 'Sale Agreed' and m.offerid = f.offer_id) as commitment_date,
       nullif(count(m.ur_current) filter (where m.updatedcategory = 'Sold Assets' and m.offerid = f.offer_id),0) as urs_sold,
       nullif(sum(m.saleprice) filter (where m.updatedcategory = 'Sold Assets' and m.offerid = f.offer_id),0) as sold_amount,
       max(m.saledate) filter (where m.updatedcategory = 'Sold Assets' and m.offerid = f.offer_id) as sale_date,
       nullif(count(m.ur_current) filter (where m.updatedcategory in ('Sold Assets', 'Sale Agreed') and m.offerid = f.offer_id),0) as total_urs_sold,
       nullif(sum(coalesce(m.saleprice,0) + coalesce(m.commitmentprice,0)) filter (where m.updatedcategory in ('Sold Assets', 'Sale Agreed') and m.offerid = f.offer_id),0) as total_proceeds,
       nullif(sum(m.lsev_dec19), 0) as lsevdec19,
       nullif(sum(m.ppa),0) as ppa,
       mode(m.bucketi_ha) as asset_bucket,
       mode(m.city) as province,
       mode(m.address) as address,
       mode(m.direccion_territorial) as dt,
       f.* exclude (offer_id, unique_id, address, servihabitat_opinion, updated_at)
       replace( array_to_string(f.unique_urs, '-') as unique_urs, 
                array_to_string(f.commercialdev, '-') as commercialdev,
                array_to_string(f.jointdev, '-') as jointdev,
                case when probable_date is not null then probable_date else offer_date end as offer_date,
        )
from all_offers f
cross join unnest(f.unique_urs) as r (urs)
left join master_tape as m
    on r.urs = m.ur_current
left join allocation as a
    on r.urs = a.ur,
lateral (
  select strptime(regexp_extract(full_path, '(2\d{7})', 1), '%Y%m%d') as probable_date
)
group by all
order by offer_date desc;
