SELECT DISTINCT a.planversion, 
                a.sku + ' - ' + a.period   AS MissingObject, 
                'No FAP for SKU in period' AS ErrorType 
FROM   tblfacts AS a 
       LEFT JOIN tblfap AS b 
              ON a.planversion = b.planversion 
                 AND a.sku = b.sku 
                 AND a.period = b.period 
WHERE  b.fapbox IS NULL 
       AND a.forecast = 'yes' 
UNION ALL 
SELECT DISTINCT a.planversion, 
                a.sku, 
                'SKU not in SKU table' 
FROM   tblfacts AS a 
       LEFT JOIN tblsku AS b 
              ON a.sku = b.sku 
WHERE  b.sku IS NULL 
UNION ALL 
SELECT DISTINCT a.planversion, 
                a.sku, 
                'Missing Weight or Packs per Box' 
FROM   tblfacts AS a 
       LEFT JOIN tblsku AS b 
              ON a.sku = b.sku 
WHERE  b.sku IS NOT NULL 
       AND ( b.weightinkg = 0 
              OR b.packperbox = 0 
              OR Isnull(b.weightinkg) 
              OR Isnull(b.packperbox) ) 
UNION ALL 
SELECT DISTINCT a.planversion, 
                a.sku + ' - ' + a.customer + '-' + a.period, 
                'No NIS for SKU and Customer in period' 
FROM   tblfacts AS a 
       LEFT JOIN tblnis AS b 
              ON a.planversion = b.planversion 
                 AND a.sku = b.sku 
                 AND a.customer = b.customer 
                 AND a.period = b.period 
WHERE  a.forecast = 'yes' 
       AND b.nisbox IS NULL 
GROUP  BY a.planversion, 
          a.sku, 
          a.customer, 
          a.period, 
          b.sku, 
          b.nisbox 
HAVING Sum(a.volume) <> 0 
UNION ALL 
SELECT DISTINCT a.planversion, 
                a.sku + ' - ' + a.period, 
                'No COGS for SKU in period' 
FROM   tblfacts AS a 
       LEFT JOIN tblcostinput AS b 
              ON a.planversion = b.planversion 
                 AND a.sku = b.sku 
                 AND a.period = b.period 
WHERE  a.forecast = 'yes' 
       AND b.sku IS NULL 
GROUP  BY a.planversion, 
          a.sku, 
          a.period, 
          b.sku 
HAVING Sum(a.volume) <> 0;   