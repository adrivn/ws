update {table_name} set offer_date = regexp_replace(trim(offer_date), '[\.\/]', '-', 'g') where regexp_matches(offer_date, '[\.\/]') = true;
update {table_name} set offer_date = strptime(offer_date, '%d-%m-%Y') where regexp_matches(offer_date, '^\d{2}-.');
update {table_name} set offer_date = strptime(regexp_extract(full_path, '(2\d{7})', 1), '%Y%m%d') where regexp_matches(offer_date, '^[345789]') or offer_date is null;
update {table_name} set offer_date = make_date(2023,month(offer_date),day(offer_date)) where year(offer_date) > 2023;
update {table_name} set offer_price = nullif(regexp_extract(offer_price, '\d+\.?\d+'), '') where regexp_matches(offer_price, '[a-zA-Z\s]');
update {table_name} set appraisal_price = nullif(regexp_extract(appraisal_price, '\d+\.?\d+'), '') where regexp_matches(appraisal_price, '[a-zA-Z\s]');
update {table_name} set web_price = NULL where regexp_matches(web_price, '\D');
update {table_name} set sap_price = NULL where regexp_matches(sap_price, '\D');
update {table_name} set unique_urs = string_to_array(regexp_replace(unique_urs,'[\[\]]', '', 'g'), ',');
update {table_name} set commercialdev = string_to_array(regexp_replace(commercialdev,'[\[\]]', '', 'g'), ',');
update {table_name} set jointdev = string_to_array(regexp_replace(jointdev,'[\[\]]', '', 'g'), ',');
update {table_name} set offer_id = regexp_replace(trim(offer_id), '[\n\t\W]', '', 'g')[:7] where regexp_matches(offer_id, '[\n\t\W]');
update {table_name} set offer_id = NULL where len(offer_id) < 6 or regexp_matches(offer_id, '[a-zA-Z]');
update {table_name} set contract_deposit = 0 where contract_deposit = '-';
update {table_name} set client_description = NULL where client_description = 'NOMBRE';
alter table {table_name} alter column offer_date set data type date;
alter table {table_name} alter column offer_price set data type double;
alter table {table_name} alter column appraisal_price set data type double;
alter table {table_name} alter column web_price set data type double;
alter table {table_name} alter column sap_price set data type double;
alter table {table_name} alter column unique_urs set data type int[];
alter table {table_name} alter column commercialdev set data type int[];
alter table {table_name} alter column jointdev set data type int[];
alter table {table_name} alter column offer_id set data type int;
alter table {table_name} add column if not exists unique_id varchar; 
update {table_name} set unique_id = md5(full_path); 
