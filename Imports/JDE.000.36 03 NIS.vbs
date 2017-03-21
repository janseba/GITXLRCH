Sub XLCode()
    Dim wks As Worksheet, row As Long, rs As Object, period As Integer, planVersion As String, periodFrom As String
    Dim connection As Object, startPeriod As Integer, periodTo As String, noPeriods As Integer, tbl As ListObject
    Dim vData As Variant
    Set wks = ActiveSheet
    Set tbl = wks.ListObjects("tblNIS")
    vData = tbl.DataBodyRange
    planVersion = GetPar([A1], "Plan Version=")
    periodFrom = GetPar([A1], "Period From=")
    periodTo = GetPar([A1], "Period To=")
    noPeriods = Right(periodTo, 2) - Right(periodFrom, 2) + 1
    If Left(periodFrom, 4) <> Left(periodTo, 4) Then
        XLImp "ERROR", "Period from and period to are not within the same year"
    End If
    If GetSQL("SELECT Locked FROM sources WHERE Source = " & Quot(planVersion)) = "y" Then
        XLImp "ERROR", "The plan version has been locked for input": Exit Sub
    End If
    Set rs = GetEmptyRecordSet("SELECT * FROM tblNIS WHERE PlanVersion IS NULL")
    
        For row = LBound(vData) To UBound(vData)
            For period = Right(periodFrom, 2) To Right(periodTo, 2)
                rs.AddNew
                rs.Fields("PlanVersion") = planVersion
                rs.Fields("Customer") = vData(row, 2)
                rs.Fields("SKU") = vData(row, 4)
                rs.Fields("Period") = CLng(Left(periodFrom, 4)) * 100 + period
                rs.Fields("NISBox") = vData(row, 7)
                rs.Fields("PaymentDiscount") = vData(row, 8)
            Next period
        Next row
    
    Set connection = GetDBConnection: connection.Open
    connection.Execute "DELETE FROM tblNIS WHERE PlanVersion = " & Quot(planVersion) & _
        " AND Period BETWEEN " & Quot(periodFrom) & " AND " & Quot(periodTo)
    rs.ActiveConnection = connection
    rs.UpdateBatch
    Application.Wait (Now + TimeValue("0:00:01"))
    XLImp "SELECT COUNT(code) FROM Companies", rs.RecordCount & " lines were added to database in 1 batch update"
    connection.Close
    XLImp "UPDATE tblNIS LEFT JOIN tblSKU ON tblNIS.SKU = tblSKU.SKU SET tblNIS.NISKg = IIf(IIf(IsNull(tblSKU.PackPerBox), 0, tblSKU.PackPerBox) = 0 " & _
        "OR IIf(IsNull(tblSKU.WeightInKg), 0, tblSKU.WeightInKg) =0, 0, tblNIS.NISBox / tblSKU.PackPerBox / tblSKU.WeightInKg) " & _
        "WHERE PlanVersion = " & Quot(planVersion) & " AND Period BETWEEN " & Quot(periodFrom) & " AND " & _
        Quot(periodTo), "Calculate NIS per Kg"

    XLImp "UPDATE tblFacts " & _
        "SET tblFacts.LPA = 0, tblFacts.discount1eur = 0 " & _
        "WHERE tblFacts.PlanVersion = " & Quot(planVersion) & " AND tblFacts.Forecast = 'yes' AND tblFacts.Period BETWEEN " & Quot(periodFrom) & " AND " & Quot(periodTo), "Reset values to 0..."

    XLImp "UPDATE tblFacts INNER JOIN tblNIS ON tblFacts.PlanVersion = tblNIS.PlanVersion AND tblFacts.SKU = tblNIS.SKU AND " & _
        "tblFacts.Customer = tblNIS.Customer AND tblFacts.Period = tblNIS.Period " & _
        "SET tblFacts.LPA = (tblFacts.Volume * tblNIS.NISKg) - tblFacts.FAP1 " & _
        "WHERE tblFacts.PlanVersion = " & Quot(planVersion) & " AND tblFacts.Forecast = 'yes' AND tblFacts.Period BETWEEN " & Quot(periodFrom) & " AND " & Quot(periodTo), "Calculate LPA..."

    XLImp "UPDATE tblFacts INNER JOIN tblNIS ON tblFacts.PlanVersion = tblNIS.PlanVersion AND tblFacts.SKU = tblNIS.SKU AND " & _
        "tblFacts.Customer = tblNIS.Customer AND tblFacts.Period = tblNIS.Period " & _
        "SET tblFacts.discount1eur = - tblNIS.PaymentDiscount * (tblFacts.FAP1 + tblFacts.LPA) " & _
        "WHERE tblFacts.PlanVersion = " & Quot(planVersion) & " AND tblFacts.Forecast = 'yes' AND tblFacts.Period BETWEEN " & Quot(periodFrom) & " AND " & Quot(periodTo), "Calculate Payment Discount..."

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