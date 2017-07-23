Sub XLCode()
    Dim wkbTemplate As Workbook, wksData As Worksheet, template As String, baseData As Variant, baseTable As ListObject
    Dim tmpRange As Range, planNew As String, planOld As String, mnth As Variant, monthFrom As Integer, monthTo As Integer
    Dim detailsData As Variant, detailsTable As ListObject
    
    'Get variables
    Set wksData = ActiveSheet
    template = GetPref(9) & "Templates\Template_PriceVariance.xlsx"
    planNew = GetPar([A2], "Plan Version=")
    planOld = GetPar([A2], "Reference 1=")
    mnth = GetPar(wksData.[A2], "Period=")
    If mnth <> "" Then
        mnth = Split(mnth, "-")
        monthFrom = CInt(mnth(0))
        monthTo = CInt(mnth(1))
    Else
        monthFrom = 1
        monthTo = 12
    End If

    'Prepare empty template
    Set wkbTemplate = Application.Workbooks.Open(Filename:=template, ReadOnly:=True)
    wkbTemplate.Sheets("By Customer").Move Before:=wksData
    wkbTemplate.Sheets("By Category").Move Before:=wksData
    wkbTemplate.Sheets("BaseData").Move Before:=wksData
    wkbTemplate.Sheets("Details").Move Before:=wksData
    ActiveWorkbook.Names("ptrPlanVersionOld").RefersToRange.Value = planOld
    ActiveWorkbook.Names("ptrPlanVersionNew").RefersToRange.Value = planNew
    
    'Extract data
    baseData = Application.Transpose(GetDBData("SELECT DISTINCT PlanningCustomer, IIf(ISNULL(Prdha4),'#NA',Prdha4), IIf(ISNULL(Prdha4),'#NA',Prdha3), IIf(ISNULL(Prdha4),'#NA',Prdha2) FROM View_PLBase WHERE (PlanVersion = " & Quot(planNew) & _
        " OR PlanVersion = " & Quot(planOld) & ") AND Mnth BETWEEN " & monthFrom & " AND " & monthTo))
    detailsData = Application.Transpose(GetDBData("SELECT PlanVersion, PlanningCustomer, IIf(ISNULL(Prdha4), '#NA', Prdha4), SUM(Volume), SUM(NIS) " & _
        ", SUM(TPR), SUM(BDF), SUM(BDFexLF), SUM(ListingFees), SUM(BMC), SUM(NOS), SUM(COGS), SUM(GP) " & _
        "FROM View_PLBase WHERE (PlanVersion = " & Quot(planNew) & " OR PlanVersion = " & Quot(planOld) & ") AND Mnth BETWEEN " & monthFrom & _
        " AND " & monthTo & " GROUP BY Prdha4, PlanningCustomer, PlanVersion"))

    'Fill Template
    Set baseTable = ActiveWorkbook.Sheets("BaseData").ListObjects("tblBaseData")
    FillTable baseData, baseTable
    Set detailsTable = ActiveWorkbook.Sheets("Details").ListObjects("tblDetails")
    FillTable detailsData, detailsTable
    
    'Add Pivots
    CreatePivotByCustomer
    CreatePivotByCategory
    
    ActiveWorkbook.Sheets("By Customer").Activate

End Sub

Sub FillTable(ByRef aData As Variant, ByRef tbl As ListObject)
    Dim rng As Range
    Set rng = tbl.ListRows.Add.Range
    rng.Cells(1, 1).Resize(UBound(aData, 1), UBound(aData, 2)).Offset(-1) = aData
End Sub

Function GetDBData(ByVal sql As String) As Variant
    Dim pw As String, connectionString As String, dbConnection As Object, rst As Object, vResult As Variant, sDbName As String
    
    pw = "xlsysjs14"
    sDbName = GetSQL("SELECT ParValue FROM XLControl WHERE Code = 'Database'")
    connectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & GetPref(9) & sDbName & "; Jet OLEDB:Database password=" & pw
    Set dbConnection = CreateObject("ADODB.Connection")
    dbConnection.Open connectionString
    Set rst = CreateObject("ADODB.Recordset")
    rst.Open sql, dbConnection, 3, 1
    If Not rst.EOF Then
        vResult = rst.GetRows
    Else
        vResult = ""
    End If
    dbConnection.Close
    Set dbConnection = Nothing
    GetDBData = vResult
End Function

Sub ResetReport()
    Dim w As Worksheet, n As Name
    For Each w In ActiveWorkbook.Sheets
        If Left(w.Name, 5) <> "XLRep" Then
            Application.DisplayAlerts = False
            w.Delete
            Application.DisplayAlerts = True
        End If
    Next w
    For Each n In ActiveWorkbook.Names
        If Left(n.Name, 5) <> "_xlfn" Then
            n.Delete
        End If
    Next n
End Sub
Sub CreatePivotByCustomer()
    
    Dim wksPivot As Worksheet, pt As PivotTable
    
    Set wksPivot = ActiveWorkbook.Sheets("By Customer")
    

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "tblBaseData", Version:=6).CreatePivotTable TableDestination:="'By Customer'!R1C1" _
        , TableName:="pivotPriceDifference", DefaultVersion:=6
    
    Set pt = wksPivot.PivotTables("pivotPriceDifference")
    
    CreatePivotSum pt, "Delta Volume", ChrW(916) & " Volume", "#,##0"
    CreatePivotSum pt, "Delta NIS", ChrW(916) & " NIS", "#,##0"
    CreatePivotSum pt, "Delta TPR", ChrW(916) & " TPR", "#,##0"
    CreatePivotSum pt, "Delta BDFexLF", ChrW(916) & " BDFexLF", "#,##0"
    CreatePivotSum pt, "Delta LF", ChrW(916) & " LF", "#,##0"
    CreatePivotSum pt, "Delta BMC", ChrW(916) & " BMC", "#,##0"
    CreatePivotSum pt, "Delta NOS", ChrW(916) & " NOS", "#,##0"
    CreatePivotSum pt, "Delta COGS", ChrW(916) & " COGS", "#,##0"
    CreatePivotSum pt, "Delta GP", ChrW(916) & " GP", "#,##0"
    
    With pt
        With .PivotFields("Prdha2")
            .Orientation = xlPageField
            .Position = 1
        End With
        With .PivotFields("Prdha3")
            .Orientation = xlPageField
            .Position = 1
        End With
        With .PivotFields("Prdha4")
            .Orientation = xlPageField
            .Position = 1
        End With
        With .PivotFields("Customer")
            .Orientation = xlRowField
            .Position = 1
        End With
    End With
        
End Sub
Sub CreatePivotByCategory()
    
    Dim wksPivot As Worksheet, pt As PivotTable
    
    Set wksPivot = ActiveWorkbook.Sheets("By Category")
    

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "tblBaseData", Version:=6).CreatePivotTable TableDestination:="'By Category'!R1C1" _
        , TableName:="pivotPriceDifference", DefaultVersion:=6
    
    Set pt = wksPivot.PivotTables("pivotPriceDifference")
    
    CreatePivotSum pt, "Delta Volume", ChrW(916) & " Volume", "#,##0"
    CreatePivotSum pt, "Delta NIS", ChrW(916) & " NIS", "#,##0"
    CreatePivotSum pt, "Delta TPR", ChrW(916) & " TPR", "#,##0"
    CreatePivotSum pt, "Delta BDFexLF", ChrW(916) & " BDFexLF", "#,##0"
    CreatePivotSum pt, "Delta LF", ChrW(916) & " LF", "#,##0"
    CreatePivotSum pt, "Delta BMC", ChrW(916) & " BMC", "#,##0"
    CreatePivotSum pt, "Delta NOS", ChrW(916) & " NOS", "#,##0"
    CreatePivotSum pt, "Delta COGS", ChrW(916) & " COGS", "#,##0"
    CreatePivotSum pt, "Delta GP", ChrW(916) & " GP", "#,##0"
    
    With pt
        With .PivotFields("Prdha3")
            .Orientation = xlPageField
            .Position = 1
        End With
        With .PivotFields("Prdha4")
            .Orientation = xlPageField
            .Position = 1
        End With
        With .PivotFields("Customer")
            .Orientation = xlPageField
            .Position = 1
        End With
        With .PivotFields("Prdha2")
            .Orientation = xlRowField
            .Position = 1
        End With
    End With
        
End Sub

Sub CreatePivotSum(ByRef pt As PivotTable, ByVal sourceName As String, ByVal targetName As String, ByVal numberFormat As String)
    With pt
        .AddDataField .PivotFields(sourceName), targetName, xlSum
        .PivotFields(targetName).numberFormat = numberFormat
    End With
End Sub