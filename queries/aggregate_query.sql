select  case when asset_id = commercialdev then 'Promo Comercial'
            when asset_id = jointdev then 'Promo Conjunta'
            else 'Unidad Registral'
        end as tipo_agrupacion,
        string_agg(distinct category, '|') as categories,
        (count(*) filter (where updatedcategory = 'Sold Assets') * 1.0 / count(*)) as 'percent_sold',
        (count(*) filter (where updatedcategory = 'Sale Agreed') * 1.0 / count(*)) as 'percent_committed',
        (count(*) filter (where updatedcategory = 'Remaining Stock') * 1.0 / count(*)) as 'percent_stock',
        asset_id,
        string_agg(distinct ur_current, '|') as unique_urs,
        string_agg(distinct commercialdev, '|') as commercialdevs,
        string_agg(distinct jointdev, '|') as jointdevs,
        string_agg(distinct w.ESTRATEGIA, '|') as estrategias,
        string_agg(distinct town,'|') as municipality,
        string_agg(distinct city,'|') as city_province,
        string_agg(distinct direccion_territorial,'|') as dts,
        string_agg(distinct cadastralreference,'|') as cadastral_references,
        sum(currentsaleprice) as listed_price,
        max(last_listing_date) as latest_listing_date,
        max(saledate) as latest_sale_date,
        max(commitmentdate) as latest_commitment_date,
        sum(val2020) as valuation2020,
        sum(val2021) as valuation2021,
        sum(val2022) as valuation2022,
        count(*) as total_urs,
        count(*) filter (where updatedcategory = 'Sold Assets') as sold_urs,
        sum(saleprice) filter (where updatedcategory = 'Sold Assets') as sold_amount,
        sum(lsev_dec19) filter (where updatedcategory = 'Sold Assets') as sold_lsev,
        sum(ppa) filter (where updatedcategory = 'Sold Assets') as sold_ppa,
        sum(nbv) filter (where updatedcategory = 'Sold Assets') as sold_nbv,
        count(*) filter (where updatedcategory = 'Sale Agreed') as committed_urs,
        sum(commitmentprice) filter (where updatedcategory = 'Sale Agreed') as committed_amount,
        sum(lsev_dec19) filter (where updatedcategory = 'Sale Agreed') as committed_lsev,
        sum(ppa) filter (where updatedcategory = 'Sale Agreed') as committed_ppa,
        sum(nbv) filter (where updatedcategory = 'Sale Agreed') as committed_nbv,
        count(*) filter (where updatedcategory = 'Remaining Stock') as remaining_urs,
        sum(lsev_dec19) filter (where updatedcategory = 'Remaining Stock') as remaining_lsev,
        sum(ppa) filter (where updatedcategory = 'Remaining Stock') as remaining_ppa,
        sum(nbv) filter (where updatedcategory = 'Remaining Stock') as remaining_nbv,

from master_tape m left join ws_segregated w on m.ur_current = w.UNIDAD_REGISTRAL,
lateral (
    select coalesce(w.UNIDAD_REGISTRAL, nullif(commercialdev, 0), nullif(jointdev,0), ur_current) AS asset_id,
            case when label is not null then
                case when publishingchannel_description is null then 'Wholesale (sold)'
                    when publishingchannel_description = 'Retail' then 'Formerly Wholesale'
                    else publishingchannel_description
                end
            else
                case when publishingchannel_description = 'Wholesale' then 'Wholesale (formerly Retail)' end
            end as category
) lat
where category is not null
group by all
order by 1, 2
