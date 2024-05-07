import polars as pl

tipos_datos = {
    "offer_id": pl.Float32,
    # "offer_date": pl.Date,
    "offer_price": pl.Float32,
    "appraisal_price": pl.Float32,
    "sap_price": pl.Float32,
    "web_price": pl.Float32,
    "due_diligence_percent": pl.Float32,
    "contract_deposit_percent": pl.Float32,
    "public_deed_percent": pl.Float32,
    "land_area": pl.Float32,
    "buildable_area": pl.Float32,
    "client_phone": pl.Float32,
}

second_pass_cast = {
    "offer_id": pl.UInt32,
    "client_phone": pl.UInt64,
}
