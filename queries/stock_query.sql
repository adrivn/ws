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
    array_agg(distinct lat.category) as categories,
    (count(*) filter (where m.updatedcategory = 'Sold Assets') * 1.0 / count(*)) as percent_sold,
    (count(*) filter (where m.updatedcategory = 'Sale Agreed') * 1.0 / count(*)) as percent_committed,
    (count(*) filter (where m.updatedcategory = 'Remaining Stock') * 1.0 / count(*)) as percent_stock,
    array_agg(distinct m.ur_current) as unique_urs,
    array_agg(distinct m.commercialdev) as commercialdevs,
    array_agg(distinct m.jointdev) as jointdevs,
    array_agg(distinct w.estrategia) as estrategias,
    array_agg(distinct m.address) as addresses,
    array_agg(distinct m.town) as municipality,
    array_agg(distinct m.city) as city_province,
    array_agg(distinct m.direccion_territorial) as dts,
    array_agg(distinct m.cadastralreference) as cadastral_references,
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
left join ws_segregated as w
    on m.ur_current = w.unidad_registral,
    lateral(
        select
            coalesce(w.unidad_registral, nullif(commercialdev, 0), nullif(jointdev, 0), ur_current) as asset_id,
            case
                when label is not null
                    then
                        case
                            when publishingchannel_description is null
                                then 'Wholesale (sold)'
                            when publishingchannel_description = 'Retail'
                                then 'Formerly Wholesale'
                            else publishingchannel_description
                        end
                when publishingchannel_description = 'Wholesale'
                    then 'Wholesale (formerly Retail)'
            end as category
    ) as lat
where lat.category is not null
group by all
order by 1, 2
