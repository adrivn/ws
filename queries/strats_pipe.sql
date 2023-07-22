-- aggregates_by_year_and_offer_status
pivot 
pipeline 
on year_planned_ep 
using 
  count(id_offer) as offers,
  sum(amount_offer) as initial_offer,
  sum(accepted_offer) as accepted_offer
group by status_offer 
order by status_offer;
-- 2023_pipeline_by_expected_closing_month
with pivot_data as (
  pivot 
  (select * from pipeline where year_planned_ep = 2023 and fecha_rechazadacada is null)
  using 
    count(id_offer) as offers,
    sum(amount_offer) as initial_offer,
    sum(importe_oferta_aprobada__pdte_aprobar) as approved_offer
  group by tipo_inmueble_agrupado_coral, month_planned_ep
)
select * from pivot_data
order by tipo_inmueble_agrupado_coral, month_planned_ep;
-- rejected_offers_by_month_and_reason
with pivot_data as (pivot 
(select *
        replace (case when fecha_rechazadacada is not null and motivo is null then 
                  'No especificado' 
                else motivo 
                end as motivo),
        month(fecha_rechazadacada) as rejection_month,
        year(fecha_rechazadacada) as rejection_year,
from pipeline
where motivo is not null)
using 
  count(id_offer) as offers,
  sum(amount_offer) as initial_offer,
  sum(importe_oferta_aprobada__pdte_aprobar) as approved_offer
group by rejection_year, rejection_month, motivo)
select * from pivot_data
order by rejection_year, rejection_month, motivo;
