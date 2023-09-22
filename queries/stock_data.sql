-- TODO: Discriminar por tipo de estrategias
select
    lat.asset_id,
    case
        when lat.asset_id = m.commercialdev
            then 'Promo Comercial'
        when lat.asset_id = m.jointdev
            then 'Promo Conjunta'
        else 'Unidad Registral'
    end as tipo_agrupacion,
    string_agg(distinct lat.category, ' | ') as categories,
    (count(*) filter (where m.updatedcategory = 'Sold Assets') * 1.0 / count(*)) as percent_sold,
    (count(*) filter (where m.updatedcategory = 'Sale Agreed') * 1.0 / count(*)) as percent_committed,
    (count(*) filter (where m.updatedcategory = 'Remaining Stock') * 1.0 / count(*)) as percent_stock,
     case when count(distinct m.ur_current) > 2 then
         array_to_string((array_agg(distinct m.ur_current::int))[1:2], ', ') || ' + other ' || count(distinct m.ur_current::int) - 2
      else
         string_agg(distinct m.ur_current::int, ', ')
     end as ur_current,
     case when count(distinct m.commercialdev::int) > 2 then
         array_to_string((array_agg(distinct m.commercialdev::int))[1:2], ', ') || ' + other ' || count(distinct m.commercialdev::int) - 2
      else
         string_agg(distinct m.commercialdev::int, ', ')
     end as commercialdev,
     case when count(distinct m.jointdev::int) > 2 then
         array_to_string((array_agg(distinct m.jointdev::int))[1:2], ', ') || ' + other ' || count(distinct m.jointdev::int) - 2
      else
         string_agg(distinct m.jointdev::int, ', ')
     end as jointdev,
    string_agg(distinct w.estrategia, ' | ') as estrategias,
     case when count(distinct m.address) > 2 then
         array_to_string((array_agg(distinct m.address))[1:2], ', ') || ' + other ' || count(distinct m.address) - 2
      else
         string_agg(distinct m.address, ', ')
     end as address,
     case when count(distinct m.town) > 2 then
         array_to_string((array_agg(distinct m.town))[1:2], ', ') || ' + other ' || count(distinct m.town) - 2
      else
         string_agg(distinct m.town, ', ')
     end as town,
     case when count(distinct m.city) > 2 then
         array_to_string((array_agg(distinct m.city))[1:2], ', ') || ' + other ' || count(distinct m.city) - 2
      else
         string_agg(distinct m.city, ', ')
     end as city,
     case when count(distinct m.cadastralreference) > 2 then
         array_to_string((array_agg(distinct m.cadastralreference))[1:2], ', ') || ' + other ' || count(distinct m.cadastralreference) - 2
      else
         string_agg(distinct m.cadastralreference, ', ')
     end as cadastralreference,
    string_agg(distinct m.direccion_territorial, ', ') as dts,
    sum(coalesce(m.currentsaleprice, m.last_listing_price)) as listed_price,
    max(m.last_listing_date) as latest_listing_date,
    max(m.saledate) as latest_sale_date,
    max(m.commitmentdate) as latest_commitment_date,
    sum(m.val2020) as valuation2020,
    sum(m.val2021) as valuation2021,
    sum(m.val2022) as valuation2022,
    count(*) as total_urs,
    count(*) filter (where m.updatedcategory = 'Sold Assets') as sold_urs,
    sum(m.saleprice) filter (where m.updatedcategory = 'Sold Assets') as sold_amount,
    sum(m.lsev_dec19) filter (where m.updatedcategory = 'Sold Assets') as sold_lsev,
    sum(m.ppa) filter (where m.updatedcategory = 'Sold Assets') as sold_ppa,
    sum(m.nbv) filter (where m.updatedcategory = 'Sold Assets') as sold_nbv,
    count(*) filter (where m.updatedcategory = 'Sale Agreed') as committed_urs,
    sum(m.commitmentprice) filter (where m.updatedcategory = 'Sale Agreed') as committed_amount,
    sum(m.lsev_dec19) filter (where m.updatedcategory = 'Sale Agreed') as committed_lsev,
    sum(m.ppa) filter (where m.updatedcategory = 'Sale Agreed') as committed_ppa,
    sum(m.nbv) filter (where m.updatedcategory = 'Sale Agreed') as committed_nbv,
    count(*) filter (where m.updatedcategory = 'Remaining Stock') as remaining_urs,
    sum(m.lsev_dec19) filter (where m.updatedcategory = 'Remaining Stock') as remaining_lsev,
    sum(m.ppa) filter (where m.updatedcategory = 'Remaining Stock') as remaining_ppa,
    sum(m.nbv) filter (where m.updatedcategory = 'Remaining Stock') as remaining_nbv

from master_tape as m
left join disaggregated_assets as w
    on m.ur_current = w.unidad_registral,
    lateral(
        select
            coalesce(w.unidad_registral, nullif(commercialdev, 0), nullif(jointdev, 0), ur_current) as asset_id,
            case
                when label is not null
                    then
                        case
                            when publishingchannel_description is null
                                then 'Wholesale'
                            when publishingchannel_description = 'Retail'
                                then 'No longer in Wholesale'
                            else publishingchannel_description
                        end
                else 
                    case when publishingchannel_description = 'Wholesale'
                      then 'Wholesale - from Retail'
                    end
            end as category
    ) as lat
where
     m.updatedcategory != 'Exclusions'
group by all
order by 1, 2
