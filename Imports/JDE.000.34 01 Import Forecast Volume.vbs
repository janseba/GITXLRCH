Sub XLCode()
    Dim wks As Worksheet, row As Long, rs As Object, planVersion As String, iStartMonth As Integer
    Dim connection As Object, country As String, bladen As Variant, sht As Variant, col As Long, periodFrom As String

    planVersion = GetPar([A1], "Plan Version=")
    country = GetPar([A1], "Country=")
    If GetSQL("SELECT Locked FROM sources WHERE Source = " & Quot(planVersion)) = "y" Then
        XLImp "ERROR", "The plan version has been locked for input": Exit Sub
    End If

    periodFrom = GetSQL("SELECT FromPeriod FROM Sources WHERE Source = " & Quot(planVersion))
    iStartMonth = CInt(Right(periodFrom, 2))

    Set rs = GetEmptyRecordSet("SELECT * FROM tblFacts WHERE PlanVersion IS NULL")

    Set wks = ActiveSheet

    With wks
        For row = 2 To wks.UsedRange.Rows.Count
            For col = 9 + iStartMonth To 21
                If Not IsEmpty(.Cells(row, 5)) And Not IsEmpty(.Cells(row, 7)) And .Cells(row, col) <> 0 And Not IsEmpty(.Cells(row, col)) Then
                    rs.AddNew
                    rs.Fields("Country") = country
                    rs.Fields("PlanVersion") = planVersion
                    rs.Fields("Period") = CStr(CLng(Left(periodFrom, 4)) * 100 + col - 9)
                    rs.Fields("SourceType") = "Volume"
                    rs.Fields("Forecast") = "yes"
                    rs.Fields("SKU") = .Cells(row, 7)
                    rs.Fields("Customer") = .Cells(row, 5)
                    If .Cells(row, 9) = "Base" Then
                        rs.Fields("PromoNonPromo") = "NonPromo"
                        rs.Fields("VolNonPromo") = .Cells(row, col) * 1000
                    Else
                        rs.Fields("PromoNonPromo") = "Promo"
                        rs.Fields("VolPromo") = .Cells(row, col) * 1000
                    End If
                    rs.Fields("OnOffInvoice") = ""
                    rs.Fields("Volume") = .Cells(row, col) * 1000
                End If
            Next col
        Next row
    End With
    Set connection = GetDBConnection: connection.Open
    connection.Execute "DELETE FROM tblFacts WHERE SourceType = 'Volume' AND PlanVersion = " & Quot(planVersion) & " AND Country = " & Quot(country)
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