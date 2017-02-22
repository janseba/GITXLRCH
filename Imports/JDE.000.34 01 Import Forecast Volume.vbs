Sub XLCode()
    Dim wks As Worksheet, row As Long, rs As Object, planVersion As String, period As String
    Dim connection As Object, country As String, bladen As Variant, sht As Variant, col as Long

    planVersion = GetPar([A1], "Plan Version=")
    country = GetPar([A1], "Country=")
If GetSQL("SELECT Locked FROM sources WHERE Source = " & Quot(planVersion)) = "y" Then
    XLImp "ERROR", "The plan version has been locked for input": Exit Sub
End If

    Set rs = GetEmptyRecordSet("SELECT * FROM tblFacts WHERE PlanVersion IS NULL")

    Set wks = ActiveSheet

    With wks
        For row = 14 To wks.UsedRange.Rows.Count
            For col = 6 to wks.UsedRange.Columns.Count
                If Not IsEmpty(.Cells(row, 3)) And Not IsEmpty(.Cells(5, col)) And .Cells(row,col) <> 0 Then
                    rs.AddNew
                    rs.Fields("Country") = country
                    rs.Fields("PlanVersion") = planVersion
                    rs.Fields("Period") = .Cells(5, col)
                    rs.Fields("SourceType") = "AOP16"
                    rs.Fields("Forecast") = "yes"
                    rs.Fields("SKU") = .Cells(row, 3)
                    rs.Fields("Customer") =  .Cells(1, col)
                    rs.Fields("PromoNonPromo") = "NonPromo"
                    rs.Fields("OnOffInvoice") = ""
                    if .Cells(2,col) = "Volume" Then
                        rs.Fields("Volume") = .Cells(row, col)
                    ElseIf .Cells(2, col) = "FAP1" Then
                        rs.Fields("FAP1") = .Cells(row, col)
                    ElseIf .Cells(2,col) = "Discount2Fix" Then
                        rs.Fields("Discount2Fix") = -1 * .Cells(row, col)
                    ElseIf .Cells(2, col) = "COGS" Then
                        rs.Fields("MB") = -1 * .Cells(row, col)
                    End If
                End If
            Next col
        Next row
    End With
    Set connection = GetDBConnection: connection.Open
    'connection.Execute "DELETE FROM tblFactsAOP WHERE SourceType = 'AOP16' AND PlanVersion = " & Quot(planVersion) & " AND Country = " & Quot(country)
    'connection.Execute "DELETE FROM tblFacts WHERE SourceType = 'AOP16' AND PlanVersion = " & Quot(planVersion) & " AND Country = " & Quot(country)
    'rs.ActiveConnection = connection
    'rs.UpdateBatch
    XLImp "SELECT COUNT(code) FROM Companies", rs.RecordCount & " lines were added to database in 1 batch update"
    XLIMP "INSERT INTO tblFacts(Country, PlanVersion, Period, SourceType, Forecast, SKU, Customer, PromoNonPromo, OnOffInvoice, Volume, FAP1, Discount2Fix, MB) " & _
		"SELECT Country, PlanVersion, Period, SourceType, Forecast, SKU, Customer, PromoNonPromo, OnOffInvoice, SUM(Volume), SUM(FAP1), SUM(Discount2Fix), SUM(MB) " & _
		"FROM tblFactsAOP " & _
		"WHERE PlanVersion = " & Quot(planVersion) & " AND Country = " & Quot(country) & _
		" GROUP BY Country, PlanVersion, Period, SourceType, Forecast, SKU, Customer, PromoNonPromo, OnOffInvoice", "Insert AOP in database"
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
    Dim pw As String, connectionString As String, dbConnection As Object
    
    pw = "xlsysjs14"
    connectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & GetPref(9) & "XLReporting_JDE_Retail_CH.dat; Jet OLEDB:Database password=" & pw
    Set dbConnection = CreateObject("ADODB.Connection")
    dbConnection.Open connectionString: dbConnection.Close
    Set GetDBConnection = dbConnection
End Function

