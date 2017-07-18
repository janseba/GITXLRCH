SELECT [a.101_salesreturnsoffinvtas] 
       + [a.101_salesreturnsoffinvfin] + a.fap1 
       + [a.101_salesreturnsoninv] + [a.105_gosfap1]           AS 
       GrossSalesValueInclCT, 
       a.lpa + [a.102_pproninv]                                AS PPR_LPA, 
       [a.106_tabmconinv] + [a.107_tabdfoninv] 
       + a.discount1perc + [a.103_tproninv] 
       + a.cpincentivesoninv + a.discount1eur 
       + a.discount1fix + [a.104_taefficiencyoninv]            AS 
       OnInvoiceConditions, 
       [a.102_pproffinvfin] + [a.102_pproffinvtas]             AS PPR, 
       [a.discount2eur] + [a.103_tproffinvfin] 
       + [a.103_tproffinvtas] 
       + [a.discount4percfap1] 
       + [a.discount4percnis] + [a.discount4eur] 
       + [a.discount4fix]                                      AS TPR, 
       [a.discount2percnis] 
       + [a.cpincentivesoffinvfin] 
       + [a.107_tabdfoffinvtas] 
       + [a.14_3termofpayment] 
       + [a.104_taefficiencyoffinvfin] 
       + [a.104_taefficiencyoffinvtas] 
       + [a.107_tabdfoffinvfin] 
       + [a.discount5percnis] + [a.discount5fix] 
       + [a.discount5percfap1] + [a.discount5eur] 
       + [a.discount2fix] 
       + [a.cpincentivesoffinvtas] 
       + [a.discount2percfap1] + [a.hq] + [a.growth] 
       + [a.yer] + [a.termsofpayment] + [a.listingfees]        AS BDF, 
       [a.106_tabmcoffinvtas] 
       + [a.18royaltieslncome3rdparty] 
       + [a.106_tabmcoffinvfin] 
       + [a.discount3percfap1] 
       + [a.discount3percnis] + [a.discount3eur] 
       + [a.discount3fix] + [a.placement] + [a.folders]        AS BMC, 
       [a.greendot] + [a.22_5ecotax] + [a.ecotax]              AS TotGreendot, 
       [a.22_3_5extracosts] + [a.displaycosts]                 AS 
       TotDisplayCosts, 
       [a.22_6excisetax] + [a.coffeetax]                       AS TotCoffeeTax, 
       [a.brewersupport] + [a.promotionother]                  AS Promotion, 
       [a.advworkingmedia] 
       + [a.advnonworkingmedia]                                AS Advertising, 
       promotion + advertising                                 AS TotalAP, 
       a.mb + a.costofsales                                    AS TotMB, 
       totgreendot + totmb + a.mbadjustments 
       + totdisplaycosts                                       AS 
       CostOfGoodsExclCT, 
       costofgoodsexclct + totcoffeetax                        AS 
       CostOfGoodsInclCT, 
       [ppr_lpa] + bmc + bdf + tpr + ppr + oninvoiceconditions AS TradeSpend, 
       grosssalesvalueinclct + tradespend                      AS NOSInclCT, 
       totcoffeetax + nosinclct                                AS NOSExclCT, 
       totcoffeetax + grosssalesvalueinclct                    AS 
       GrossSalesValueExclCT, 
       nosexclct + costofgoodsexclct                           AS GP, 
       gp + totalap                                            AS MarketingCM, 
       a.shipping + a.warehouse                                AS WD, 
       wd + marketingcm                                        AS CM, 
       a.otherga + a.selling + a.marketing 
       + a.sellingoffice + a.sellingfield 
       + a.controlling + a.finance + a.tax + a.it 
       + a.humanresources + a.facility + a.legal 
       + a.supplychain + a.generalmanagement                   AS OSA, 
       cm + osa                                                AS Ebit, 
       grosssalesvalueinclct + [ppr_lpa]                       AS FAP2, 
       oninvoiceconditions + fap2                              AS NISInclCT, 
       a.discount5percfap1 + a.discount5percnis 
       + a.discount5eur + a.discount5fix                       AS Discount5, 
       [a.17_1onelistfee] + [a.listingfees]                    AS ListingFees, 
       [a.15_2specialoccasion]                                 AS FixedAmounts, 
       a.*, 
       b.description                                           AS SKUDescription 
       , 
       b.prdha4, 
       b.prdha3, 
       b.prdha2, 
       b.prdha1, 
       b.salesplanning, 
       b.profitcenter, 
       b.alternativehierarchy, 
       b.brand, 
       b.otherdetails, 
       b.weightinkg, 
       b.tdiscperkg, 
       b.cupsperkg, 
       b.bridgehierarchy, 
       b.reportingcategory, 
       b.promotionalsku, 
       c.customername, 
       c.planningcustomer, 
       c.conditioncustomer, 
       c.distributionchannel 
FROM   (tblfacts AS a 
        LEFT JOIN tblsku AS b 
               ON a.sku = b.sku) 
       LEFT JOIN tblcustomer AS c 
              ON a.customer = c.customer;   