Sub XLCode()
    Dim wks As Worksheet, row As Long, rs As Object, period As Integer, planVersion As String, periodFrom As String
    Dim connection As Object, country As String, startPeriod As Integer, periodTo As String, noPeriods As Integer
    Set wks = ActiveSheet
    planVersion = GetPar([A1], "Plan Version=")
    country = GetPar([A1], "Country=")
    periodFrom = GetPar([A1], "Period From=")
    periodTo = GetPar([A1], "Period To=")
    noPeriods = Right(periodTo, 2) - Right(periodFrom, 2) + 1
    If Left(periodFrom, 4) <> Left(periodTo, 4) Then
        XLImp "ERROR", "Period from and period to are not within the same year"
    End If
    If GetSQL("SELECT Locked FROM sources WHERE Source = " & Quot(planVersion)) = "y" Then
        XLImp "ERROR", "The plan version has been locked for input": Exit Sub
    End If
    Set rs = GetEmptyRecordSet("SELECT * FROM tblCostInput WHERE PlanVersion IS NULL")
    
    With wks
        For row = 5 To wks.UsedRange.Rows.Count
            If Not IsEmpty(.Cells(row, 4)) And .Cells(row, 3) = "L133" Then
                For period = Right(periodFrom, 2) To Right(periodTo, 2)
                    rs.AddNew
                    rs.Fields("Country") = country
                    rs.Fields("PlanVersion") = planVersion
                    rs.Fields("LineType") = "COGS"
                    rs.Fields("SKU") = .Cells(row, 4)
                    rs.Fields("Period") = CLng(Left(periodFrom, 4)) * 100 + period
                    rs.Fields("Amount") = .Cells(row, 39)
                Next period
            End If
        Next row
    End With
    Set connection = GetDBConnection: connection.Open
    connection.Execute "DELETE FROM tblCostInput WHERE PlanVersion = " & Quot(planVersion) & " AND Country = " & Quot(country) & _
        " AND Period BETWEEN " & Quot(periodFrom) & " AND " & Quot(periodTo)
    rs.ActiveConnection = connection
    rs.UpdateBatch
    XLImp "SELECT COUNT(code) FROM Companies", rs.RecordCount & " lines were added to database in 1 batch update"
    connection.Close

    XLImp "UPDATE tblFacts LEFT JOIN tblCostInput ON tblFacts.PlanVersion = tblCostInput.PlanVersion AND tblFacts.Period = tblCostInput.Period AND tblFacts.SKU = tblCostInput.SKU " & _
        "SET tblFacts.mb = 0 WHERE tblFacts.PlanVersion = " & Quot(planVersion) & " AND tblFacts.Forecast = 'yes' " & _
        "AND tblFacts.Period BETWEEN " & Quot(periodFrom) & " AND " & Quot(periodTo)

    XLImp "UPDATE tblFacts LEFT JOIN tblCostInput ON tblFacts.PlanVersion = tblCostInput.PlanVersion AND tblFacts.Period = tblCostInput.Period AND tblFacts.SKU = tblCostInput.SKU " & _
        "SET tblFacts.mb = -1 * tblFacts.Volume * IIf(ISNULL(tblCostInput.Amount), 0, tblCostInput.Amount) WHERE tblFacts.PlanVersion = " & Quot(planVersion) & " AND tblFacts.Forecast = 'yes' " & _
        "AND tblFacts.Period BETWEEN " & Quot(periodFrom) & " AND " & Quot(periodTo)
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