import duckdb
from pathlib import Path
from datetime import datetime

directorio = Path("//EURFL01/advisors.hal/non-hudson/Coral Homes/CoralHudson/6. Stock/Parquet").absolute()

db = duckdb.connect()
db.execute(f"CREATE VIEW master_tape AS SELECT * FROM parquet_scan('{directorio}/master_tape.parquet')")
db.execute(f"CREATE VIEW offers_data AS SELECT * FROM parquet_scan('{directorio}/offers.parquet')")
print("Accessing data...")
agg_data = db.execute("""select coalesce(nullif(commercialdev, 0), nullif(jointdev,0), ur_current) as asset_id,
            string_agg(distinct city,'|') as city_province,
            string_agg(distinct direccion_territorial,'|') as dts,
            count(*) as total_urs,
            count(*) filter (where updatedcategory = 'Sold Assets') as sold_urs,
            sum(lsev_dec19) filter (where updatedcategory = 'Sold Assets') as sold_lsev,
            sum(ppa) filter (where updatedcategory = 'Sold Assets') as sold_ppa,
            sum(nbv) filter (where updatedcategory = 'Sold Assets') as sold_nbv,
            count(*) filter (where updatedcategory = 'Sale Agreed') as committed_urs,
            sum(lsev_dec19) filter (where updatedcategory = 'Sale Agreed') as committed_lsev,
            sum(ppa) filter (where updatedcategory = 'Sale Agreed') as committed_ppa,
            sum(nbv) filter (where updatedcategory = 'Sale Agreed') as committed_nbv,
            count(*) filter (where updatedcategory = 'Remaining Stock') as remaining_urs,
            sum(lsev_dec19) filter (where updatedcategory = 'Remaining Stock') as remaining_lsev,
            sum(ppa) filter (where updatedcategory = 'Remaining Stock') as remaining_ppa,
            sum(nbv) filter (where updatedcategory = 'Remaining Stock') as remaining_nbv,
                from master_tape
            where label is not null
            group by all""").df()
print("Creating Excel output file...")
now_rendered = datetime.strftime(datetime.now(), "%d-%m-%Y")
agg_data.to_excel("stock_file.xlsx", index=False, sheet_name=f"Wholesale Stock {now_rendered}")
