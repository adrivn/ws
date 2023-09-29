WITH rankings AS (
    SELECT
        m.ur_current,
        m.city AS province_svh,
        p.*,
        jaro_similarity(p.province, m.city) score,
        rank() over (puntuacion) AS rank
    FROM
        iso3166_provincias p,
        master_tape m WINDOW puntuacion AS (
            PARTITION by m.ur_current
            ORDER BY
                score DESC
        )
    ORDER BY
        1,
        4 DESC
)
SELECT
    lat.asset_id,
    CASE
        WHEN lat.asset_id = m.commercialdev THEN 'Promo Comercial'
        WHEN lat.asset_id = m.jointdev THEN 'Promo Conjunta'
        ELSE 'Unidad Registral'
    END AS tipo_agrupacion,
    lat.category AS categories,
    lat.source_label,
    trim(MODE(w.estrategia)) AS strategy,
    (
        count(*) filter (
            WHERE
                m.updatedcategory = 'Sold Assets'
        ) * 1.0 / count(*)
    ) AS percent_sold,
    (
        count(*) filter (
            WHERE
                m.updatedcategory = 'Sale Agreed'
        ) * 1.0 / count(*)
    ) AS percent_committed,
    (
        count(*) filter (
            WHERE
                m.updatedcategory = 'Remaining Stock'
        ) * 1.0 / count(*)
    ) AS percent_stock,
    array_to_string(
        list_distinct(array_agg(trim(m.ur_current :: int))),
        ' | '
    ) AS ur_current,
    array_to_string(
        list_distinct(array_agg(trim(m.commercialdev :: int))),
        ' | '
    ) AS commercialdev,
    array_to_string(
        list_distinct(array_agg(trim(m.jointdev :: int))),
        ' | '
    ) AS jointdev,
    trim(MODE(m.address)) AS address,
    trim(MODE(m.town)) AS municipality,
    trim(MODE(m.city)) AS province,
    array_to_string(
        list_distinct(array_agg(trim(m.cadastralreference))),
        ' | '
    ) AS cadastralreference,
    trim(MODE(m.direccion_territorial)) AS dts,
    sum(
        coalesce(m.currentsaleprice, m.last_listing_price)
    ) AS listed_price,
    max(m.last_listing_date) AS latest_listing_date,
    max(m.saledate) AS latest_sale_date,
    max(m.commitmentdate) AS latest_commitment_date,
    sum(m.val2020) AS valuation2020,
    sum(m.val2021) AS valuation2021,
    sum(m.val2022) AS valuation2022,
    count(*) AS total_urs,
    count(*) filter (
        WHERE
            m.updatedcategory = 'Sold Assets'
    ) AS sold_urs,
    sum(m.saleprice) filter (
        WHERE
            m.updatedcategory = 'Sold Assets'
    ) AS sold_amount,
    sum(m.lsev_dec19) filter (
        WHERE
            m.updatedcategory = 'Sold Assets'
    ) AS sold_lsev,
    sum(m.ppa) filter (
        WHERE
            m.updatedcategory = 'Sold Assets'
    ) AS sold_ppa,
    sum(m.nbv) filter (
        WHERE
            m.updatedcategory = 'Sold Assets'
    ) AS sold_nbv,
    count(*) filter (
        WHERE
            m.updatedcategory = 'Sale Agreed'
    ) AS committed_urs,
    sum(m.commitmentprice) filter (
        WHERE
            m.updatedcategory = 'Sale Agreed'
    ) AS committed_amount,
    sum(m.lsev_dec19) filter (
        WHERE
            m.updatedcategory = 'Sale Agreed'
    ) AS committed_lsev,
    sum(m.ppa) filter (
        WHERE
            m.updatedcategory = 'Sale Agreed'
    ) AS committed_ppa,
    sum(m.nbv) filter (
        WHERE
            m.updatedcategory = 'Sale Agreed'
    ) AS committed_nbv,
    count(*) filter (
        WHERE
            m.updatedcategory = 'Remaining Stock'
    ) AS remaining_urs,
    sum(m.lsev_dec19) filter (
        WHERE
            m.updatedcategory = 'Remaining Stock'
    ) AS remaining_lsev,
    sum(m.ppa) filter (
        WHERE
            m.updatedcategory = 'Remaining Stock'
    ) AS remaining_ppa,
    sum(m.nbv) filter (
        WHERE
            m.updatedcategory = 'Remaining Stock'
    ) AS remaining_nbv,
    MODE(r.code) AS province_code,
    MODE(r.region_code) AS region_code,
    AVG(CAST(ll.Y_GOOGLE AS DOUBLE)) AS latitude,
    AVG(CAST(ll.X_GOOGLE AS DOUBLE)) AS longitude,
    m.updated_at
FROM
    master_tape AS m
    LEFT JOIN rankings r ON r.ur_current = m.ur_current
    LEFT JOIN (
        SELECT
            UNIDAD_REGISTRAL,
            MODE(X_GOOGLE) AS X_GOOGLE,
            MODE(Y_GOOGLE) AS Y_GOOGLE
        FROM
            (
                SELECT
                    *
                FROM
                    channels_historic
                UNION
                ALL
                SELECT
                    load_date,
                    UNIDAD_REGISTRAL,
                    X_GOOGLE,
                    Y_GOOGLE,
                    CHANNEL
                FROM
                    latest_operations
            ) AS coordinates
        WHERE
            X_GOOGLE IS NOT NULL
        GROUP BY
            1
    ) ll ON m.ur_current = ll.UNIDAD_REGISTRAL
    LEFT JOIN (
        SELECT
            UNIDAD_REGISTRAL,
            CHANNEL,
            SNAPSHOT_DATE
        FROM
            (
                SELECT
                    *
                FROM
                    channels_historic
                UNION
                ALL
                SELECT
                    load_date,
                    UNIDAD_REGISTRAL,
                    X_GOOGLE,
                    Y_GOOGLE,
                    CHANNEL
                FROM
                    latest_operations
            ) AS outie
        WHERE
            SNAPSHOT_DATE = (
                SELECT
                    MAX(SNAPSHOT_DATE)
                FROM
                    channels_historic AS innie
                WHERE
                    innie.UNIDAD_REGISTRAL = outie.UNIDAD_REGISTRAL
            )
    ) ch ON m.ur_current = ch.UNIDAD_REGISTRAL
    LEFT JOIN disaggregated_assets AS w ON m.ur_current = w.unidad_registral,
    lateral(
        SELECT
            coalesce(
                w.unidad_registral,
                nullif(m.commercialdev, 0),
                nullif(m.jointdev, 0),
                m.ur_current
            ) :: int AS asset_id,
            CASE
                WHEN m.salechannel IS NOT NULL THEN CASE
                    WHEN m.salechannel = 'Mayorista' THEN 'Wholesale'
                    ELSE 'Retail'
                END
                ELSE coalesce(ch.CHANNEL, 'Retail')
            END AS category,
            CASE
                WHEN m.salechannel IS NOT NULL THEN 'Sale Channel (Reporting)'
                ELSE 'Operations Tape'
            END AS source_label,
            -- case
            -- when label is not null
            --     then
            --         case
            --             when publishingchannel_description is null
            --                 then 'Wholesale'
            --             when publishingchannel_description = 'Retail'
            --                 then 'No longer in Wholesale'
            --             else publishingchannel_description
            --         end
            -- else
            --     case when publishingchannel_description = 'Wholesale'
            --       then 'Wholesale - from Retail'
            --     end
            -- end as category
    ) AS lat
WHERE
    m.updatedcategory != 'Exclusions'
    AND lat.category = 'Wholesale'
    AND r.rank = 1
GROUP BY
    ALL
ORDER BY
    1,
    2
