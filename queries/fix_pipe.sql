-- Renombrado de columnas
alter table {table_name} rename promo__ur to asset_id;
-- Cambio de datatypes
alter table {table_name} alter id_offer set data type int;
alter table {table_name} alter asset_id set data type int[] using string_split(asset_id, '/');
alter table {table_name} alter month_planned_ep set data type tinyint using month(planned_signing_date);
alter table {table_name} alter q_planned_ep set data type tinyint using quarter(planned_signing_date);
alter table {table_name} alter year_planned_ep set data type smallint using year(planned_signing_date);
alter table {table_name} alter month_signed_deposit_agreement set data type tinyint using month(signing_date_deposit_agreement);
alter table {table_name} alter year_signed_deposit_agreement set data type smallint using year(signing_date_deposit_agreement);
alter table {table_name} alter month_sale_ep set data type tinyint using month(date_of_sale);
-- Arreglo de valores
