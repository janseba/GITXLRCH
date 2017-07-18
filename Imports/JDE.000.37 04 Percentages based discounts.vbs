Sub XLCode()
    Dim wks As Worksheet, row As Long, rs As Object, period As Integer, Planversion As String, PeriodFrom As String
    Dim connection As Object, startPeriod As Integer, PeriodTo As String, noPeriods As Integer, vSheetNames As Variant
    Dim vWksName As Variant, n As Name, rngCustomer As Range, prdha3 As String, colCustomers As Collection, customerNumbers As Range
    
    vSheetNames = Array("Headquarters", "Growth Bonus", "YER", "Payment Terms", "Placement", "Folders")
    Set customerNumbers = Range("rngCustomers")
    Set colCustomers = New Collection
    With customerNumbers
        For row = 1 To .Rows.Count
            colCustomers.Add .Cells(row, 2).Value, .Cells(row, 1).Value
        Next row
    End With
    
    Planversion = GetPar([A1], "Plan Version=")
    PeriodFrom = GetPar([A1], "Period From=")
    PeriodTo = GetPar([A1], "Period To=")
    noPeriods = Right(PeriodTo, 2) - Right(PeriodFrom, 2) + 1
    startPeriod = Right(PeriodFrom, 2)
    If Left(PeriodFrom, 4) <> Left(PeriodTo, 4) Then
        XLImp "ERROR", "Period from and period to are not within the same year"
    End If
    If GetSQL("SELECT Locked FROM sources WHERE Source = " & Quot(Planversion)) = "y" Then
        XLImp "ERROR", "The plan version has been locked for input": Exit Sub
    End If
    Set rs = GetEmptyRecordSet("SELECT * FROM tblPercentageDiscounts WHERE PlanVersion IS NULL")
    
    For Each vWksName In vSheetNames
        Set wks = ActiveWorkbook.Worksheets(vWksName)
        For Each n In wks.Names
            If Mid(n.Name, InStr(1, n.Name, "!") + 1, 1) = "C" Then 'we are only intersted in customer ranges, they all start with C
                Set rngCustomer = n.RefersToRange
                For period = startPeriod To 12
                    With rngCustomer
                        For row = 5 To .Rows.Count
                            If UCase(.Cells(row, period)) = "X" Then
                                prdha3 = wks.Cells(.Cells(row, period).row, 1)
                                rs.AddNew
                                rs.Fields("DiscountCategory") = vWksName
                                rs.Fields("PlanVersion") = Planversion
                                rs.Fields("Period") = Left(PeriodFrom, 4) * 100 + period
                                rs.Fields("Prdha3") = prdha3
                                rs.Fields("Customer") = colCustomers.Item(Mid(n.Name, InStr(1, n.Name, "!") + 1))
                                rs.Fields("Discount") = .Cells(2, period)
                            End If
                        Next row
                    End With
                Next period
            End If
        Next n
    Next vWksName
       
    Set connection = GetDBConnection: connection.Open
    connection.Execute "DELETE FROM tblPercentageDiscounts WHERE PlanVersion = " & Quot(Planversion) & _
        " AND Period BETWEEN " & Quot(PeriodFrom) & " AND " & Quot(PeriodTo)
    rs.ActiveConnection = connection
    rs.UpdateBatch
    Application.Wait (Now + TimeValue("0:00:01"))
    XLImp "SELECT COUNT(code) FROM Companies", rs.RecordCount & " lines were added to database in 1 batch update"
    connection.Close

    XLImp "UPDATE tblFacts " & _
        "SET tblFacts.HQ = 0, tblFacts.Growth = 0, tblFacts.YER = 0, tblFacts.TermsOfPayment = 0, tblFacts.Placement = 0, tblFacts.Folders = 0 " & _
        "WHERE tblFacts.PlanVersion = " & Quot(Planversion) & " AND tblFacts.Forecast = 'yes' AND tblFacts.Period BETWEEN " & Quot(PeriodFrom) & " AND " & Quot(PeriodTo), "Reset values to 0..."
    
    UpdateField "HQ", "Headquarters", Planversion, PeriodFrom, PeriodTo
    UpdateField "Growth", "Growth Bonus", Planversion, PeriodFrom, PeriodTo
    UpdateField "YER", "YER", Planversion, PeriodFrom, PeriodTo
    UpdateField "TermsOfPayment", "Payment Terms", Planversion, PeriodFrom, PeriodTo
    UpdateField "Placement", "Placement", Planversion, PeriodFrom, PeriodTo
    UpdateField "Folders", "Folders", Planversion, PeriodFrom, PeriodTo
    UpdateField "ListingFees", "Listing Fees", Planversion, PeriodFrom, PeriodTo

End Sub
Sub UpdateField(ByVal fieldName As String, ByVal category As String, ByVal Planversion As String, ByVal PeriodFrom As String, ByVal PeriodTo As String)
    XLImp "UPDATE (tblFACTS INNER JOIN tblSKU ON tblFacts.SKU = tblSKU.SKU) INNER JOIN tblPercentageDiscounts ON " & _
        "tblPercentageDiscounts.PlanVersion = tblFacts.PlanVersion AND tblPercentageDiscounts.Period = tblFacts.Period AND " & _
        "tblPercentageDiscounts.Prdha3 = tblSKU.Prdha3 AND tblPercentageDiscounts.Customer = tblFacts.Customer " & _
        "Set tblFacts." & fieldName & " = (tblFacts.FAP1 + tblFacts.LPA + tblFacts.discount1eur) * -tblPercentageDiscounts.Discount " & _
        "WHERE tblPercentageDiscounts.DiscountCategory = " & Quot(category) & " AND tblFacts.PlanVersion = " & Quot(Planversion) & _
        " AND tblFacts.Forecast ='yes' AND tblFacts.Period BETWEEN " & Quot(PeriodFrom) & " AND " & Quot(PeriodTo)
End Sub
Function GetEmptyRecordSet(ByVal sTable As String) As Object
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
    Set GetEmptyRecordSet = rsData
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