Sub XLCode()
    Dim wks As Worksheet, row As Long, rs As Object, planVersion As String, period As String, dblValue As Double
    Dim connection As Object, country As String, bladen As Variant, sht As Variant, col As Long, sSQL As String

    planVersion = GetPar([A1], "Plan Version=")
    country = GetPar([A1], "Country=")
    If GetSQL("SELECT Locked FROM sources WHERE Source = " & Quot(planVersion)) = "y" Then
        XLImp "ERROR", "The plan version has been locked for input": Exit Sub
    End If

    Set rs = GetEmptyRecordSet("SELECT * FROM tblFacts WHERE PlanVersion IS NULL")

    Set wks = ActiveSheet

    With wks
        For row = 2 To wks.UsedRange.Rows.Count
            For col = 8 To wks.UsedRange.Columns.Count
                If .Cells(row, col) <> 0 And Not IsEmpty(.Cells(row, col)) Then
                    dblValue = .Cells(row, col) * 1000
                    rs.AddNew
                    rs.Fields("Country") = country
                    rs.Fields("PlanVersion") = planVersion
                    rs.Fields("Period") = .Cells(1, col)
                    rs.Fields("SourceType") = "AOP17"
                    rs.Fields("Forecast") = "yes"
                    rs.Fields("SKU") = .Cells(row, 5)
                    rs.Fields("Customer") = .Cells(1, 4)
                    rs.Fields("PromoNonPromo") = "NonPromo"
                    rs.Fields("OnOffInvoice") = ""
                    If .Cells(row, 7) = "Volume" Then
                        rs.Fields("Volume") = dblValue
                    ElseIf .Cells(row, 7) = "FAP" Then
                        rs.Fields("FAP1") = dblValue
                    ElseIf .Cells(row, 7) = "PPR" Then
                        rs.Fields("LPA") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "Del" Then
                        rs.Fields("discount1eur") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "BDF1" Then
                        rs.Fields("discount2fix") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "TPR" Then
                        rs.Fields("discount2eur") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "WB" Then
                        rs.Fields("discount3eur") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "Folder" Then
                        rs.Fields("discount3percnis") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "LF" Then
                        rs.Fields("107_TABDFOffInvTAS") = -1 * dblValue
                        rs.Fields("17_1OneListFee") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "COGS" Then
                        rs.Fields("MB") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "Working Media" Then
                        rs.Fields("AdvWorkingMedia") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "Non Working Media" Then
                        rs.Fields("AdvNonWorkingMedia") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "Promotion - Maschine" Then
                        rs.Fields("BrewerSupport") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "Promotion - Other" Then
                        rs.Fields("PromotionOther") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "Distribution - Warehouse Total" Then
                        rs.Fields("Warehouse") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "Distribution - Shipping" Then
                        rs.Fields("Shipping") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "Marketing" Then
                        rs.Fields("Marketing") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "Selling Office" Then
                        rs.Fields("SellingOffice") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "Selling Field" Then
                        rs.Fields("SellingField") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "Controlling" Then
                        rs.Fields("Controlling") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "Finance" Then
                        rs.Fields("Finance") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "Tax" Then
                        rs.Fields("Tax") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "IT" Then
                        rs.Fields("IT") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "Human Resources" Then
                        rs.Fields("HumanResources") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "Facility" Then
                        rs.Fields("Facility") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "Legal" Then
                        rs.Fields("Legal") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "Supply Chain" Then
                        rs.Fields("SupplyChain") = -1 * dblValue
                    ElseIf .Cells(row, 7) = "General Management" Then
                        rs.Fields("GeneralManagement") = -1 * dblValue
                    End If
                End If
            Next col
        Next row
    End With
    Set connection = GetDBConnection: connection.Open
    connection.Execute "DELETE FROM tblFacts WHERE SourceType = 'AOP17' AND PlanVersion = " & Quot(planVersion) & " AND Country = " & Quot(country)
    rs.ActiveConnection = connection
    rs.UpdateBatch
    XLImp "SELECT COUNT(code) FROM Companies", rs.RecordCount & " lines were added to database in 1 batch update"
    connection.Close

    sSQL = "UPDATE tblFacts AS a LEFT JOIN tblSKU AS b ON a.SKU = b.SKU " & _
      "SET a.PromoNonPromo = IIf(a.PromoNonPromo = 'NonPromo' AND b.PromotionalSKU = 'yes', 'Promo', a.PromoNonPromo)" & _
      " WHERE a.PlanVersion = " & Quot(planVersion)

    XLImp sSQL, "Check promotional SKUs"

    sSQL = "UPDATE tblFacts SET VolPromo = IIf(PromoNonPromo = 'Promo', Volume, 0), " & _
      "VolNonPromo = IIf(PromoNonPromo = 'NonPromo', Volume, 0) WHERE PlanVersion = " & Quot(planVersion)

    XLImp sSQL, "Split volumes into promo and non promo."
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