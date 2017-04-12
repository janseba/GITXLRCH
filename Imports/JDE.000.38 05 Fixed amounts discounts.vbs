Sub XLCode()
    Dim vWorksheets As Variant, w As Variant, wks As Worksheet, iCol As Integer, periodFrom As String, planVersion As String
    Dim iRow As Integer, sPrdha3 As String, sql As String, dealFrom As String, dealTo As String, noPeriods As Integer, period As Integer
    Dim NIS As Double, rsNIS As Object, rsFacts As Object, connection As Object, AllocatedDiscount As Double, rsSKU As Object
    Dim sSKU As String
    
    vWorksheets = Array("TPR", "Growth Bonus", "YER", "ECR", "Special Occasion", "Placement", "FolderAds")
    planVersion = GetPar([A1], "Plan Version=")
    If planVersion = "" Then XLImp "Error", "There was no planversion selected": Exit Sub
    periodFrom = GetSQL("SELECT FromPeriod FROM Sources WHERE Source = " & Quot(planVersion))
    
    RunSQL "DELETE FROM tblFacts WHERE SourceType = 'fixdisc' AND PlanVersion = " & Quot(planVersion)
    
    For Each w In vWorksheets
        Set wks = ActiveWorkbook.Worksheets(w)
        With wks
            For iCol = 2 To wks.UsedRange.Columns.Count
                dealTo = .Cells(8, iCol)
                'Check if Customer Nr and Amount are filled and if forecast periodFrom <= end period of deal
                If Not IsEmpty(.Cells(6, iCol)) And .Cells(10, iCol) <> 0 And periodFrom <= dealTo Then
                    sPrdha3 = ""
                    'make a list of prdha3
                    For iRow = 14 To .UsedRange.Rows.Count
                        If .Cells(iRow, iCol) = "X" Then sPrdha3 = sPrdha3 & "," & Quot(.Cells(iRow, 1))
                    Next iRow
                    sPrdha3 = Mid(sPrdha3, 2)
                    
                    'get number of periods
                    If periodFrom < .Cells(7, iCol) Then dealFrom = .Cells(7, iCol) Else dealFrom = periodFrom
                    noPeriods = Right(dealTo, 2) - Right(dealFrom, 2) + 1
                    
                    'get recordset with NIS details and totals
                    sql = "SELECT 'Detail' AS Type, b.Period, b.Customer, b.SKU, b.PromoNonPromo, SUM(b.NIS) AS NIS FROM View_PLBase AS b WHERE b.PlanVersion = " & Quot(planVersion) & " AND b.Period BETWEEN " & Quot(dealFrom) & " AND " & Quot(dealTo) & " AND Customer = " & Quot(.Cells(6, iCol)) & " AND Prdha3 IN (" & sPrdha3 & ") GROUP BY b.Period, b.PromoNonPromo, b.Customer, b.SKU UNION ALL SELECT 'Total' AS Type, a.Period, " & Quot(.Cells(6, iCol)) & " AS Customer,'TOTAL' AS SKU, 'TOTAL' AS PromoNonPromo, SUM(a.NIS) AS NIS FROM View_PLBase AS a WHERE a.PlanVersion = " & Quot(planVersion) & " AND a.Period BETWEEN " & _
                        Quot(dealFrom) & " AND " & Quot(dealTo) & " AND a.Customer = " & Quot(.Cells(6, iCol)) & " AND a.Prdha3 IN (" & sPrdha3 & ")" & _
                        " GROUP BY Period"
                    Set rsNIS = GetRecordSet(sql)
                    Set rsFacts = GetRecordSet("SELECT * FROM tblFacts WHERE PlanVersion IS NULL")
                    
                    
                    'add records to tblFacts for each Period
                    For period = 1 To noPeriods
                        AllocatedDiscount = 0
                        rsNIS.Filter = "Type = 'Total' AND Period = " & Quot(dealFrom + period - 1)
                        NIS = rsNIS.Fields("NIS").Value
                        rsNIS.Filter = 0
                        rsNIS.Filter = "Type = 'Detail' AND Period = " & Quot(dealFrom + period - 1)
                        rsNIS.MoveFirst
                        Do Until rsNIS.EOF
                            rsFacts.AddNew
                            rsFacts.Fields("Country") = "CH"
                            rsFacts.Fields("PlanVersion") = planVersion
                            rsFacts.Fields("Period") = rsNIS.Fields("Period").Value
                            rsFacts.Fields("SourceType") = "fixdisc"
                            rsFacts.Fields("Forecast") = "yes"
                            rsFacts.Fields("SKU") = rsNIS.Fields("SKU").Value
                            rsFacts.Fields("Customer") = rsNIS.Fields("Customer").Value
                            rsFacts.Fields("PromoNonPromo") = rsNIS.Fields("PromoNonPromo").Value
                            rsFacts.Fields("Discount4Fix") = -(rsNIS.Fields("NIS") / NIS) * (.Cells(10, iCol) / .Cells(9, iCol))
                            AllocatedDiscount = AllocatedDiscount + rsFacts.Fields("Discount4Fix")
                            rsNIS.MoveNext
                        Loop
                        If -(.Cells(10, iCol) / .Cells(9, iCol)) <> AllocatedDiscount Then
                            'Determine number of profit centers
                            Set rsSKU = GetRecordSet("SELECT DISTINCT ProfitCenter FROM tblSKU WHERE Prdha3 IN (" & sPrdha3 & ")")
                            'More than one profit center than put unallocated fixed discount on Beans
                            If rsSKU.RecordCount > 1 Then
                                rsFacts.AddNew
                                rsFacts.Fields("Country") = "CH"
                                rsFacts.Fields("PlanVersion") = planVersion
                                rsFacts.Fields("Period") = dealFrom + period - 1
                                rsFacts.Fields("SourceType") = "fixdisc"
                                rsFacts.Fields("Forecast") = "yes"
                                rsFacts.Fields("SKU") = "D-BEANS"
                                rsFacts.Fields("Customer") = .Cells(6, iCol)
                                rsFacts.Fields("PromoNonPromo") = "NonPromo"
                                rsFacts.Fields("Discount4Fix") = -(.Cells(10, iCol) / .Cells(9, iCol)) - AllocatedDiscount
                            Else
                                'Get dummy sku for profit center
                                Set rsSKU = GetRecordSet("SELECT  TOP 1 SKU FROM tblSKU WHERE ProfitCenter = " & Quot(rsSKU.Fields("ProfitCenter").Value) & " AND Description LIKE '%Dummy%' AND SKU <> 'D-FINANCE'")
                                If rsSKU.RecordCount <> 1 Then sSKU = "D-BEANS" Else sSKU = rsSKU.Fields("SKU").Value
                                rsFacts.AddNew
                                rsFacts.Fields("Country") = "CH"
                                rsFacts.Fields("PlanVersion") = planVersion
                                rsFacts.Fields("Period") = dealFrom + period - 1
                                rsFacts.Fields("SourceType") = "fixdisc"
                                rsFacts.Fields("Forecast") = "yes"
                                rsFacts.Fields("SKU") = sSKU
                                rsFacts.Fields("Customer") = .Cells(6, iCol)
                                rsFacts.Fields("PromoNonPromo") = "NonPromo"
                                rsFacts.Fields("Discount4Fix") = -(.Cells(10, iCol) / .Cells(9, iCol)) - AllocatedDiscount
                            End If
                        End If
                    Next period
                Set connection = GetDBConnection: connection.Open
                rsFacts.ActiveConnection = connection
                rsFacts.UpdateBatch
                connection.Close
                End If
            Next iCol
        End With
    Next w
End Sub

Function GetRecordSet(ByVal sTable As String) As Object
    Dim rsData As Object, connection As Object
    
    Set connection = GetDBConnection()
    connection.Open
    Set rsData = CreateObject("ADODB.Recordset")
    With rsData
        .CursorLocation = 3 'adUseClient
        .CursorType = 1 'adOpenKeyset
        .LockType = 4 'adLockBatchOptimistic
        .Open sTable, connection
        .ActiveConnection = Nothing
    End With
    
    connection.Close
    Set GetRecordSet = rsData
End Function
Function GetDBConnection() As Object
    Dim pw As String, connectionString As String, dbConnection As Object, sDbName As String
    
    pw = "xlsysjs14"
    sDbName = GetSQL("SELECT ParValue FROM XLControl WHERE Code = 'Database'")
    connectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & GetPref(9) & sDbName & "; Jet OLEDB:Database password=" & pw
    Set dbConnection = CreateObject("ADODB.Connection")
    dbConnection.Open connectionString: dbConnection.Close
    Set GetDBConnection = dbConnection
End Function