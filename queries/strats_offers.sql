-- of_sold_by_week_l3m
WITH pivot_data AS (
   PIVOT (select  *, 
                  week(offer_date) as week, 
                  year(offer_date) as year
         from offers_enriched_table where offer_date > now() - interval 3 months 
        and (total_proceeds is not null)
  )
      ON (year)
      USING   
            COUNT(unique_id) as assets, 
            SUM(total_urs) as total_urs, 
            SUM(offer_price) as offer_amount, 
            SUM(commitment_amount) as commitments, 
            SUM(sold_amount) as sales,
            SUM(total_proceeds) as total_proceeds,
            SUM(lsevdec19) as lsev_dec19,
            SUM(ppa) as ppa
    GROUP BY asset_bucket
)
select *,
        round(("2023_total_proceeds" / "2023_ppa" - 1) * 100.0, 2)  || '%' as perf_vs_ppa,
        round(("2023_total_proceeds" / "2023_lsev_dec19" - 1) * 100.0, 2) || '%' as perf_vs_lsev19,
from pivot_data;
-- of_pipeline_by_week_l3m
WITH pivot_data AS (
   PIVOT (select  *, 
                  week(offer_date) as week, 
                  year(offer_date) as year
         from offers_enriched_table where offer_date > now() - interval 3 months 
        and (total_proceeds is null)
  )
      ON (year)
      USING   
            COUNT(unique_id) as assets, 
            SUM(total_urs) as total_urs, 
            SUM(offer_price) as offer_amount, 
            SUM(lsevdec19) as lsev_dec19,
            SUM(ppa) as ppa
    GROUP BY asset_bucket
)
select *,
        round(("2023_offer_amount" / "2023_ppa" - 1) * 100.0, 2)  || '%' as perf_vs_ppa,
        round(("2023_offer_amount" / "2023_lsev_dec19" - 1) * 100.0, 2) || '%' as perf_vs_lsev19,
from pivot_data;
