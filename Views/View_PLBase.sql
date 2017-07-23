SELECT a.planversion, 
       a.period, 
       Val(RIGHT(a.period, 2))      AS Mnth, 
       a.sourcetype, 
       a.forecast, 
       a.sku, 
       b.description                AS SKUDescription, 
       b.prdha4, 
       b.prdha3, 
       b.prdha2, 
       b.prdha1, 
       b.promotionalsku, 
       b.profitcenter, 
       b.brand, 
       b.bridgehierarchy, 
       b.reportingcategory, 
       a.customer, 
       c.customername, 
       c.planningcustomer, 
       a.promononpromo, 
       a.onoffinvoice, 
       Sum(a.volume)                AS Volume, 
       Sum(a.volpromo)              AS VolPromo, 
       Sum(a.volnonpromo)           AS VolNonPromo, 
       Sum(a.ebit)                  AS Ebit, 
       Sum(a.osa)                   AS OSA, 
       Sum(a.cm)                    AS CM, 
       Sum(a.wd)                    AS WD, 
       Sum(a.marketingcm)           AS MarketingCM, 
       Sum(a.gp)                    AS GP, 
       Sum(a.totalap)               AS TotalAP, 
       Sum(a.advertising)           AS Advertising, 
       Sum(a.promotion)             AS Promotion, 
       Sum(a.nosinclct)             AS NOS, 
       Sum(a.grosssalesvalueinclct) AS GOS, 
       Sum(a.tradespend)            AS TradeSpend, 
       Sum(a.ppr_lpa)               AS LPA, 
       Sum(a.ppr)                   AS PPR, 
       Sum(a.tpr)                   AS TPR, 
       Sum(a.oninvoiceconditions)   AS OnInvoiceConditions, 
       Sum(a.grosssalesvalueinclct + a.ppr_lpa 
           + a.oninvoiceconditions) AS NIS, 
       Sum(a.bdf)                   AS BDF, 
       Sum(a.bmc)                   AS BMC, 
       Sum(a.costofgoodsexclct)     AS COGS, 
       Sum(a.totdisplaycosts)       AS DisplayCosts, 
       Sum(a.totmb)                 AS MB, 
       Sum(a.totgreendot)           AS Greendot,
       Sum(a.ListingFees)           AS ListingFees,
       Sum(a.BDFexLF)               AS BDFexLF
FROM   (view_facts AS a 
        LEFT JOIN tblsku AS b 
               ON a.sku = b.sku) 
       LEFT JOIN tblcustomer AS c 
              ON a.customer = c.customer 
GROUP  BY a.planversion, 
          a.period, 
          a.sourcetype, 
          a.forecast, 
          a.sku, 
          b.description, 
          b.prdha4, 
          b.prdha3, 
          b.prdha2, 
          b.prdha1, 
          b.promotionalsku, 
          b.profitcenter, 
          b.brand, 
          b.bridgehierarchy, 
          b.reportingcategory, 
          a.customer, 
          c.customername, 
          c.planningcustomer, 
          a.promononpromo, 
          a.onoffinvoice;   