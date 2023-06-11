SELECT
    o.offerid,
    string_agg(DISTINCT o.ur_current, '|') AS unique_urs,
    string_agg(DISTINCT o.commercialdev, '|') AS commercialdev,
    string_agg(DISTINCT o.jointdev, '|') AS jointdev,
    string_agg(DISTINCT o.offerstatus, '|') AS offer_status,
    count(o.ur_current) AS total_urs,
    sum(m.ppa) AS ppa,
    sum(m.lsev_dec19) AS lsev_dec19
FROM offers_data AS o
INNER JOIN master_tape AS m
    ON o.ur_current = m.ur_current
GROUP BY 1;
