SELECT DISTINCT a.PlanVersion,
                a.SKU + ' - ' + a.Period AS MissingObject,
                                'No FAP for SKU in period' AS ErrorType
FROM tblFacts AS a
LEFT JOIN tblFAP AS b ON a.planversion = b.planversion
AND a.SKU = b.SKU
AND a.Period = b.Period
WHERE b.FAPBox IS NULL
  AND a.Forecast = 'yes'
UNION ALL
SELECT DISTINCT a.PlanVersion,
                a.SKU,
                'SKU not in SKU table'
FROM tblFacts AS a
LEFT JOIN tblSKU AS b ON a.SKU = b.SKU
WHERE b.SKU IS NULL
UNION ALL
SELECT DISTINCT a.PlanVersion,
                a.SKU,
                'Missing Weight or Packs per Box'
FROM tblFacts AS a
LEFT JOIN tblSKU AS b ON a.SKU = b.SKU
WHERE b.SKU IS NOT NULL
  AND (b.WeightInKg = 0
       OR b.PackPerBox = 0
       OR ISNULL(b.WeightInKg)
       OR ISNULL(b.PackPerBox))