Sub XLCode()
    Dim wks As Worksheet, row As Long, rs As Object, period As Integer, planVersion As String, periodFrom As String, periodTo As String
    Dim connection As Object, country As String, startPeriod As Integer, endPeriod As Integer, rngInput As Range

    planVersion = GetPar([A1], "Plan Version=")
    country = GetPar([A1], "Country=")
    periodFrom = GetPar([A1], "Period From=")
    periodTo = GetPar([A1], "Period To=")
    startPeriod = CInt(Right(periodFrom, 2))
    endPeriod = CInt(Right(periodTo, 2))
    Set rs = GetEmptyRecordSet("SELECT * FROM tblFacts WHERE PlanVersion IS NULL")

    Set wks = ActiveSheet
    Set rngInput = wks.Names(wks.Name & "!XLRInput").RefersToRange

    With rngInput
        For row = 3 To rngInput.Rows.Count
            If Not IsEmpty(.Cells(row, 1)) Then
                For period = startPeriod To endPeriod
                    If .Cells(row, period + 1) <> 0 Then
                        rs.AddNew
                        rs.Fields("Country") = country
                        rs.Fields("PlanVersion") = planVersion
                        rs.Fields("Period") = CLng(Left(periodFrom, 4)) * 100 + period
                        rs.Fields("SourceType") = "SGAActuals"
                        rs.Fields("Forecast") = "no"
                        rs.Fields("sku") = wks.Names(wks.Name & "!XLRsku").RefersToRange.Value
                        rs.Fields("Customer") = wks.Names(wks.Name & "!XLRcustomer").RefersToRange.Value
                        rs.Fields("PromoNonPromo") = "NonPromo"
                        rs.Fields(CStr(.Cells(row, 1))) = -1 * IIf(IsError(.Cells(row, period + 1)), 0, .Cells(row, period + 1))
                    End If
                Next period
            End If
        Next row
    End With

    Set connection = GetDBConnection: connection.Open
    connection.Execute "DELETE FROM tblFacts WHERE SourceType = 'SGAActuals' AND PlanVersion = " & Quot(planVersion) & " AND Country = " & _
        Quot(country) & " AND Period BETWEEN " & Quot(periodFrom) & " AND " & Quot(periodTo)
    rs.ActiveConnection = connection
    rs.UpdateBatch
    XLImp "SELECT COUNT(code) FROM Companies", rs.RecordCount & " lines were added to database in 1 batch update"
    connection.Close
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