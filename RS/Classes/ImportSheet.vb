Imports System.Globalization
Imports SpreadsheetGear
Imports System.Text.RegularExpressions
Imports System.IO


Public MustInherit Class RetailSalesImporter

    Property ImportedRetailData As DataTable
    Property retailImportType As RetailImportType

    Property SalesWeek As String
    Property SalesPeriod As String
    Property DepartmentCode As String
    Property CustomerCode As String

    Property IsValid As Boolean
    Property ValidationErrors As String
    Property FileName As String
    Property FileSize As Int32

    ''' <summary>
    ''' Fill ImportedRetailData datatable with retail sales data
    ''' </summary>
    ''' <remarks></remarks>
    Public MustOverride Sub Import()

    Protected Overridable Function InitializeForImport() As Boolean
        Return True
    End Function

    Public Shared Function CreateImporter(ByVal importFileName As String) As RetailSalesImporter
        Dim fi = New FileInfo(importFileName)

        If ASCDATA1.GetDataValue("SELECT COUNT(*) FROM RSTFILE1 WHERE IMPORT_FILENAME=:PARM1 AND IMPORT_FILESIZE=:PARM2", "VN", New Object() {fi.Name, fi.Length}) > 0 Then
            Throw New Exception(String.Format("File {0} has already been imported", fi.Name))
        End If

        Select Case fi.Extension.ToUpper
            Case ".XLS", ".XLSX", ".XLSM"
                Return RetailExcelImport.CreateExcelImporter(importFileName)
        End Select

        Return Nothing
    End Function

    Sub New(ByVal importFileName As String)
        Dim fi = New FileInfo(importFileName)
        FileSize = fi.Length
        FileName = fi.Name

        ImportedRetailData = New DataTable()
        ImportedRetailData.Columns.Add("CUST_STORE_NO", GetType(String))
        ImportedRetailData.Columns.Add("CUST_DEPT_CODE", GetType(String))
        ImportedRetailData.Columns.Add("AMT_SOLD", GetType(Decimal))
        ImportedRetailData.Columns.Add("AMT_SOLD_MTD", GetType(Decimal))
        ImportedRetailData.Columns.Add("AMT_SOLD_STD", GetType(Decimal))
        ImportedRetailData.Columns.Add("AMT_PLAN", GetType(Decimal))
        ImportedRetailData.Columns.Add("OPS_YYYYPP", GetType(String))
        ImportedRetailData.Columns.Add("OPS_YYYYWW", GetType(String))
        ImportedRetailData.Columns.Add("MEMO", GetType(String))
        ImportedRetailData.Columns.Add("MATL_CODE", GetType(String))
        ImportedRetailData.PrimaryKey = New DataColumn() {ImportedRetailData.Columns("CUST_STORE_NO"), ImportedRetailData.Columns("CUST_DEPT_CODE"), ImportedRetailData.Columns("MATL_CODE")}
    End Sub

End Class

Public MustInherit Class RetailExcelImport
    Inherits RetailSalesImporter

    Private Property WorkbookSet As IWorkbookSet
    Protected Property ImportFile As IWorkbook


    Public Shared Function CreateExcelImporter(ByVal importFileName As String) As RetailExcelImport
        SetGemboxLicense()

        Dim excelDocument As IWorkbook = LoadExcelFile(importFileName)

        Dim importType As RetailImportType = DetermineCustomerFromExcel(excelDocument)

        Select Case importType
            Case retailImportType.SaksExcel
                Return New SaksRetailImportExcel(importFileName, excelDocument)
            Case retailImportType.SaksConsignmentExcel
                Return New SaksConsignmentImportExcel(importFileName, excelDocument)
            Case retailImportType.NordstromExcel
                Return New NordstromRetailImportExcel(importFileName, excelDocument)
            Case retailImportType.BloomingdalesExcel
                Return New BloomingdalesRetailImport(importFileName, excelDocument)
            Case retailImportType.BloomingdalesExcel2
                Return New BloomingdalesRetailImport2(importFileName, excelDocument)
            Case retailImportType.BloomingdalesExcel3
                Return New BloomingdalesRetailImport3(importFileName, excelDocument)
            Case retailImportType.NeimanMarcusFlashExcel
                Return New NeimanMarcusFlashImport(importFileName, excelDocument)
            Case retailImportType.NeimanMarcusRetailExcel
                Return New NeimanMarcusRetailImport(importFileName, excelDocument)
            Case retailImportType.NeimanMarcusConsignmentExcel
                Return New NeimanMarcusConsignmentImport(importFileName, excelDocument)
        End Select

        Return Nothing
    End Function

    Protected Shared Sub SetGemboxLicense()
        'SpreadsheetInfo.SetLicense("EMPX-L9BW-EL8E-4GKJ")

    End Sub

    Private Shared Function LoadExcelFile(ByVal importFileName As String) As IWorkbook
        Dim workbookSet As SpreadsheetGear.IWorkbookSet = SpreadsheetGear.Factory.GetWorkbookSet()

        Dim importFile As IWorkbook = workbookSet.Workbooks.Open(importFileName)

        Return importFile
    End Function

    Private Shared Function DetermineCustomerFromExcel(ByVal importFile As IWorkbook) As RetailImportType
        Try
            For Each sht As IWorksheet In importFile.Worksheets
                If (sht.Cells("W1").Value IsNot Nothing AndAlso sht.Cells("W1").Value = "NORDSTROM") Or (sht.Cells("V1").Value IsNot Nothing AndAlso sht.Cells("V1").Value = "NORDSTROM") Then
                    Return retailImportType.NordstromExcel
                End If

                If sht.Cells("A2").Value IsNot Nothing AndAlso sht.Cells("A2").Value.ToString.StartsWith("DEPTCL") Then
                    Return retailImportType.BloomingdalesExcel
                End If


                Try
                    If sht.Cells("C9").Value & "" = "BLM" OrElse (sht.PageSetup.RightFooter IsNot Nothing AndAlso sht.PageSetup.RightFooter.ToUpper.Contains("MACY'S")) Then
                        Return retailImportType.BloomingdalesExcel2
                    End If
                Catch ex As Exception
                End Try

                Try
                    If (sht.Cells("A6").Value & "").ToString().ToUpper().StartsWith("SALES AND STOCK") Then
                        Return retailImportType.BloomingdalesExcel3
                    End If
                Catch ex As Exception
                End Try

                If sht.Cells("A2").Value IsNot Nothing AndAlso sht.Cells("A2").Value.ToString.StartsWith("Consigned Sales") Then
                    Return retailImportType.NeimanMarcusConsignmentExcel
                End If

                If sht.Cells("H1").Value IsNot Nothing AndAlso sht.Cells("H1").Value.ToString.StartsWith("Neiman Marcus") Then
                    Return retailImportType.NeimanMarcusFlashExcel
                End If

                If (sht.Cells("A1").Value IsNot Nothing AndAlso sht.Cells("A1").Value.ToString.ToUpper.Contains("TRUNK")) _
                    Or (sht.Cells("A1").Value IsNot Nothing AndAlso sht.Cells("A1").Value.ToString.ToUpper.StartsWith("JH D76 WEEKLY")) _
                    Or (sht.Cells("A5").Value IsNot Nothing AndAlso sht.Cells("A5").Value.ToString.ToUpper.StartsWith("DEPT") AndAlso sht.Cells("C5").Value IsNot Nothing AndAlso sht.Cells("C5").Value.ToString.ToUpper.StartsWith("76")) _
                     Or (sht.Cells("A2").Value IsNot Nothing AndAlso sht.Cells("A2").Value.ToString.ToUpper.Contains("TRUNK")) _
                     Or (sht.Cells("B1").Value IsNot Nothing AndAlso sht.Cells("B1").Value.ToString.ToUpper.Contains("TOTAL_SALES_R")) _
                     Or (sht.Cells("A3").Value IsNot Nothing AndAlso sht.Cells("A3").Value.ToString.ToUpper.Contains("ITEM #")) _
                     Or (sht.Cells("D4").Value IsNot Nothing AndAlso sht.Cells("D4").Value.ToString.ToUpper.Contains("SALES $ WTD")) _
                     Or (sht.Cells("D5").Value IsNot Nothing AndAlso sht.Cells("D5").Value.ToString.ToUpper.Contains("SALES $ WTD")) _
                     Or (sht.Cells("D6").Value IsNot Nothing AndAlso sht.Cells("D6").Value.ToString.ToUpper.Contains("SALES $ WTD")) _
                     Or (sht.Cells("D7").Value IsNot Nothing AndAlso sht.Cells("D7").Value.ToString.ToUpper.Contains("SALES $ WTD")) Then
                    Return retailImportType.SaksConsignmentExcel

                End If

                If sht.Name = "TOTAL DOLLARS" Then
                    If sht.Cells("R2").Value IsNot Nothing AndAlso sht.Cells("R2").Value.ToString.Trim = "Display:" Then
                        Return retailImportType.NeimanMarcusRetailExcel
                    End If
                    If sht.Cells("R3").Value IsNot Nothing AndAlso sht.Cells("R3").Value.ToString.Trim = "Display:" Then
                        Return retailImportType.NeimanMarcusRetailExcel
                    End If
                End If

                If (sht.Cells("O4").Value IsNot Nothing AndAlso sht.Cells("O4").Value.ToString.ToUpper.Contains("FLASH")) OrElse
                    (sht.Cells("N4").Value IsNot Nothing AndAlso sht.Cells("N4").Value.ToString.ToUpper.Contains("FLASH")) Then
                    Return retailImportType.SaksExcel
                End If
            Next
        Catch ex As Exception
            '"Unable to verify customer"
        End Try

        Return Nothing
    End Function

    Public Sub New(ByVal importFileName As String, ByVal excelDocument As IWorkbook)
        MyBase.New(importFileName)

        Me.ImportFile = excelDocument
        IsValid = False
        If InitializeForImport() Then
            IsValid = True
        End If
    End Sub
End Class

Public Class NordstromRetailImportExcel
    Inherits RetailExcelImport

    Private Property ImportWorksheet As IWorksheet

    Private Const StartDateCell As String = "E1"
    Private Const EndDateCell As String = "G1"
    Private Const DeptCell As String = "E3"

    Private Const StoreNoColumnIndex As Integer = 2
    Private Const SalesAmtColumnIndex As Integer = 5

    Public Sub New(ByVal importFileName As String, ByVal excelDocument As IWorkbook)
        MyBase.New(importFileName, excelDocument)
        CustomerCode = "NORDSTR10"
        retailImportType = retailImportType.NordstromExcel
    End Sub

    Protected Overrides Function InitializeForImport() As Boolean
        ImportWorksheet = GetSheetForImport(ImportFile)
        If ImportWorksheet IsNot Nothing Then
            Return True
        End If
        Return False
    End Function

    Private Function GetSheetForImport(ByVal excelFile As IWorkbook) As IWorksheet
        Dim sheetToImport As IWorksheet = Nothing
        For Each sht As IWorksheet In excelFile.Worksheets
            If IsSheetValid(sht) Then
                Dim shtSalesWeek As String = GetTimeFrame(sht)
                If Convert.ToInt32(shtSalesWeek) > Convert.ToInt32(SalesWeek) Then
                    sheetToImport = sht
                    SalesWeek = shtSalesWeek
                End If
            End If
        Next
        Return sheetToImport
    End Function

    Protected Function GetTimeFrame(ByVal worksheet As IWorksheet) As String
        Dim startDate, endDate As Date
        Dim areDatesValid As Boolean

        areDatesValid = Date.TryParse(workSheet.Cells(StartDateCell).Value, startDate) And
                        Date.TryParse(workSheet.Cells(EndDateCell).Value, endDate)

        If areDatesValid Then
            If endDate <> startDate.AddDays(6) Then
                Return Nothing
            End If

            Dim YYYYWW As String = ASCDATA1.GetDataValue("SELECT YYYYWW FROM GLTPARM3 WHERE WEEK_END_DATE=:PARM1", "D", New Object() {Convert.ToDateTime(endDate)})

            Return YYYYWW
        End If

        Return Nothing
    End Function

    Private Function IsSheetValid(ByRef worksheet As IWorksheet) As Boolean
        If worksheet.Cells("A1").Value <> "" Then 'IF THE SHEET DOES NOT CONTAIN A BLANK COLUMN A
            worksheet.Cells("A").Insert(InsertShiftDirection.Right)
            'worksheet.Cells.InsertColumn(0) 'ADD BLANK COLUMN TO SHIFT COLUMNS INTO PROPER POSITION FOR VALIDATION
        End If

        If worksheet.Cells("W1").Value = "NORDSTROM" And worksheet.Cells("A1").Value = "" Then 'this appears to be a Nordstrom sheet
            Return True
        End If

        Return False
    End Function

    Public Overrides Sub Import()
        If ImportWorksheet.Cells("A1").Value <> "" Then
            ImportWorksheet.Cells("A").Insert() 'some sheets do not have the blank column A
        End If

        LoadDateData(ImportWorksheet)
        LoadDeptData(ImportWorksheet)
        LoadSalesData(ImportWorksheet)
    End Sub

#Region "Header Data"

    Private Sub LoadDateData(ByVal salesDataSheet As IWorksheet)
        Dim weekEndDate As String = salesDataSheet.Cells(EndDateCell).Value

        Dim drDateInfo As DataRow = ASCDATA1.GetDataRow("SELECT * FROM GLTPARM3 WHERE WEEK_END_DATE=:PARM1", "D", New Object() {Convert.ToDateTime(weekEndDate)})
        SalesWeek = drDateInfo.Item("YYYYWW")
        SalesPeriod = drDateInfo.Item("YYYYPP")
    End Sub

    Private Sub LoadDeptData(ByVal salesDataSheet As IWorksheet)
        DepartmentCode = salesDataSheet.Cells(DeptCell).Value
    End Sub

#End Region

#Region "Sales Data"
    Private Sub LoadSalesData(ByVal salesDataSheet As IWorksheet)
        'Loop through the sheet and create a new datarow for each row of sales data info
        For Each salesDataRow In salesDataSheet.Cells.Rows
            If isSalesDataRow(salesDataRow) Then
                '(For Nordstroms) the relevant data is the store# (col C) and sales amt (col F)
                Dim storeNo = Convert.ToString(salesDataRow(StoreNoColumnIndex).Value)
                Dim salesAmt = Convert.ToDecimal(salesDataRow(SalesAmtColumnIndex).Value)

                AddSalesDataRow(storeNo, salesAmt)
            End If
        Next
    End Sub

    Private Function isSalesDataRow(ByVal salesDataRow As IRange) As Boolean
        If Not salesDataRow(StoreNoColumnIndex).Value & "" = "" Then
            If Integer.TryParse(salesDataRow(StoreNoColumnIndex).Value, Nothing) Then
                Return True
            End If
        End If
        Return False
    End Function

    Private Sub AddSalesDataRow(ByVal storeNo As String, ByVal salesAmt As Decimal)
        Dim sdr As DataRow = ImportedRetailData.NewRow()
        sdr.Item("CUST_STORE_NO") = storeNo.PadLeft(6, "0")
        sdr.Item("CUST_DEPT_CODE") = DepartmentCode
        sdr.Item("AMT_SOLD") = salesAmt
        sdr.Item("MEMO") = "0"
        sdr.Item("MATL_CODE") = "S"
        sdr.Item("OPS_YYYYPP") = SalesPeriod
        sdr.Item("OPS_YYYYWW") = SalesWeek
        ImportedRetailData.Rows.Add(sdr)
    End Sub
#End Region

End Class

Public Class BloomingdalesRetailImport
    Inherits RetailExcelImport

    Private Const DateCell As String = "A3"
    Private Const DeptCell As String = "A2"

    Private Const StoreNoColumnIndex As Integer = 0
    Private SalesAmtColumnIndex As Integer = 2 'might be 2 if hidden columns are not counted
    Private SalesAmtMTDColumnIndex As Integer = 5
    Private SalesAmtSTDColumnIndex As Integer = 8


    Public Sub New(ByVal fileName As String, ByVal excelDocument As IWorkbook)
        MyBase.New(fileName, excelDocument)
        retailImportType = retailImportType.BloomingdalesExcel
        CustomerCode = "BLOOMIES10"
    End Sub

    Protected Overrides Function InitializeForImport() As Boolean
        'Bloomingdales IWorkbooks contain multiple sheets
        'there are sheets for asset vs memo -- these are different departments
        'there may also be separate sheets for gold -- gold sheets are separate
        Return True
    End Function

    Public Overrides Sub Import()
        For Each sht As IWorksheet In ImportFile.Worksheets
            If sht.Visible = SheetVisibility.Visible Then
                LoadDateData(sht)
                LoadDeptData(sht)
                LoadSalesData(sht)
            End If
        Next
    End Sub

    Private Sub LoadDateData(sht As IWorksheet)
        Dim dateString As String = sht.Cells(DateCell).Value
        'dateString is in format: Per x Wk y z
        Dim splitDateString As String() = dateString.Split(" ")
        Dim periodNo As Int32 = Integer.Parse(splitDateString(1))
        Dim weekNo As Int32 = Integer.Parse(splitDateString(3))
        Dim yearNo As Int32 = Integer.Parse(splitDateString(4))

        If periodNo = 12 Then 'Bloomingdales period is offset -1 from current month (i.e. December = Per 11)
            periodNo = 1
            yearNo += 1
        Else
            periodNo += 1
        End If
        Dim YYYYMM As String = String.Format("20" & (yearNo).ToString.PadLeft(2, "0") & periodNo.ToString.PadLeft(2, "0"))
        Dim drDateInfo As DataRow = ASCDATA1.GetDataRow("SELECT * FROM GLTPARM3 WHERE YYYYMM=:PARM1 AND REL_WEEK=:PARM2", "VN", New Object() {YYYYMM, weekNo})
        'Dim drDateInfo As DataRow = ASCDATA1.GetDataRow("SELECT * FROM GLTPARM3 WHERE YYYYWW=(SELECT TO_NUMBER(YYYYWW) FROM GLTPARM3 WHERE YYYYMM=:PARM1 AND REL_WEEK=:PARM2)-100", "VN", New Object() {YYYYMM, weekNo})

        SalesWeek = drDateInfo.Item("YYYYWW")
        SalesPeriod = drDateInfo.Item("YYYYPP")
    End Sub

    Private Sub LoadDeptData(sht As IWorksheet)
        Dim deptString As String = sht.Cells(DeptCell).Value
        'deptString is in format: DEPTCL xxx yy - need clarification
        DepartmentCode = deptString.Substring(4)
    End Sub

    Private Sub LoadSalesData(sht As IWorksheet)
        Dim sheetName As String = sht.Name
        Dim memo As Boolean = sheetName.ToUpper.Contains("MEMO")
        Dim material As String = DepartmentCode 'material can be derived from dept code

        SetColumnIndices(sht)

        For Each salesDataRow As IRange In sht.Cells.Rows
            If isSalesDataRow(salesDataRow) Then
                '(For Bloomingdales) the relevant data is the store# (col A) and WTD sales amt (col C)
                Dim storeNo = Convert.ToString(salesDataRow(StoreNoColumnIndex).Value).PadLeft(6, "0")
                Dim salesAmt = Convert.ToDecimal(salesDataRow(SalesAmtColumnIndex).Value) * 1000
                Dim salesAmtMTD As Decimal
                Dim salesAmtSTD As Decimal

                Try
                    salesAmtMTD = Convert.ToDecimal(salesDataRow(SalesAmtMTDColumnIndex).Value) * 1000
                    salesAmtSTD = Convert.ToDecimal(salesDataRow(SalesAmtSTDColumnIndex).Value) * 1000
                Catch ex As Exception

                End Try
                AddSalesDataRow(storeNo, memo, material, salesAmt, salesAmtMTD, salesAmtSTD)
            End If
        Next
    End Sub

    Private Sub SetColumnIndices(ByVal excelWorksheet As IWorksheet)
        'WTD, PTD, STD columns are not always the same
        Dim lastColumnHeader As String = ""
        For i As Integer = 1 To excelWorksheet.Cells.Columns.Count - 1
            Dim columnHeader As String = excelWorksheet.Cells(4, i).Value & ""
            If columnHeader <> lastColumnHeader Then
                Select Case columnHeader
                    Case "WTD Sales"
                        If excelWorksheet.Cells(5, i).Value & "" = "TY" Then
                            SalesAmtColumnIndex = i '+ 1
                        End If
                    Case "PTD Sales"
                        If excelWorksheet.Cells(5, i).Value & "" = "TY" Then
                            SalesAmtMTDColumnIndex = i '+ 1
                        End If
                    Case "STD Sales"
                        If excelWorksheet.Cells(5, i).Value & "" = "TY" Then
                            SalesAmtSTDColumnIndex = i '+ 1
                        End If
                End Select
            End If
            lastColumnHeader = columnHeader
        Next
    End Sub

    Private Function isSalesDataRow(ByVal salesDataRow As IRange) As Boolean
        If Not salesDataRow(StoreNoColumnIndex).Value & "" = "" Then
            If Integer.TryParse(salesDataRow(StoreNoColumnIndex).Value, Nothing) Then
                Return True
            End If
        End If
        Return False
    End Function

    Private Sub AddSalesDataRow(ByVal storeNo As String, ByVal memo As Boolean, ByVal materialCode As String, ByVal salesAmt As Decimal, ByVal salesAmtMTD As Decimal, ByVal salesAmtSTD As Decimal)
        Dim sdr As DataRow

        sdr = ImportedRetailData.NewRow()
        sdr.Item("CUST_STORE_NO") = storeNo.PadLeft(6, "0")
        sdr.Item("CUST_DEPT_CODE") = DepartmentCode
        sdr.Item("MEMO") = If(memo, "1", "0")
        sdr.Item("MATL_CODE") = materialCode
        sdr.Item("AMT_SOLD") = salesAmt
        sdr.Item("AMT_SOLD_MTD") = salesAmtMTD
        sdr.Item("AMT_SOLD_STD") = salesAmtSTD
        sdr.Item("OPS_YYYYPP") = SalesPeriod
        sdr.Item("OPS_YYYYWW") = SalesWeek
        ImportedRetailData.Rows.Add(sdr)

    End Sub

End Class


Public Class BloomingdalesRetailImport2
    Inherits RetailExcelImport

    Private Const DateCell As String = "A3"
    Private Const DeptCell As String = "A2"

    Private Const StoreNoColumnIndex As Integer = 3
    Private SalesAmtColumnIndex As Integer = 6
    Private SalesAmtMTDColumnIndex As Integer = 10
    Private SalesAmtSTDColumnIndex As Integer = 14


    Public Sub New(ByVal fileName As String, ByVal excelDocument As IWorkbook)
        MyBase.New(fileName, excelDocument)
        retailImportType = retailImportType.BloomingdalesExcel2
        CustomerCode = "BLOOMIES10"
    End Sub

    Protected Overrides Function InitializeForImport() As Boolean
        'Bloomingdales IWorkbooks contain multiple sheets
        'there are sheets for asset vs memo -- these are different departments
        'there may also be separate sheets for gold -- gold sheets are separate
        Return True
    End Function

    Public Overrides Sub Import()
        For Each sht As IWorksheet In ImportFile.Worksheets
            If sht.Visible = SheetVisibility.Visible And Not sht.Name.StartsWith("#7") And Not sht.Name.Contains("TTL") Then 'only process visible sheets, and not the totals sheet
                LoadDateData(sht)
                LoadDeptData(sht)
                LoadSalesData(sht)
            End If
        Next
    End Sub

    Private Sub LoadDateData(sht As IWorksheet)
        'get date info from header
        'dateString is in format: AS OF PERIOD x, WEEK x, YYYY
        Dim dateString As String = sht.PageSetup.CenterHeader
        Dim rgx As Match = Regex.Match(dateString, "AS OF PERIOD (\d+), WEEK (\d+), (\d{4})")
        Dim periodNo As Int32 = Integer.Parse(rgx.Groups(1).Value)
        Dim weekNo As Int32 = Integer.Parse(rgx.Groups(2).Value)
        Dim yearNo As Int32 = Integer.Parse(rgx.Groups(3).Value)

        If periodNo = 12 Then 'Bloomingdales period is offset -1 from current month (i.e. December = Per 11)
            periodNo = 1
            yearNo += 1
        Else
            periodNo += 1
        End If
        Dim YYYYMM As String = String.Format(yearNo & periodNo.ToString.PadLeft(2, "0"))
        Dim drDateInfo As DataRow = ASCDATA1.GetDataRow("SELECT * FROM GLTPARM3 WHERE YYYYMM=:PARM1 AND REL_WEEK=:PARM2", "VN", New Object() {YYYYMM, weekNo})

        SalesWeek = drDateInfo.Item("YYYYWW")
        SalesPeriod = drDateInfo.Item("YYYYPP")
    End Sub

    Private Sub LoadDeptData(sht As IWorksheet)
        Dim deptString As String = sht.Name
        DepartmentCode = deptString
    End Sub

    Private Sub LoadSalesData(sht As IWorksheet)
        Dim sheetName As String = sht.Name
        Dim memo As Boolean = sheetName.ToUpper.Contains("MEMO")
        Dim material As String = DepartmentCode 'material can be derived from dept code

        For Each salesDataRow As IRange In sht.Cells.Rows
            If isSalesDataRow(salesDataRow) Then
                Dim storeNo = Convert.ToString(salesDataRow(StoreNoColumnIndex).Value).PadLeft(6, "0")
                Dim salesAmt = Convert.ToDecimal(salesDataRow(SalesAmtColumnIndex).Value) * 1000
                Dim salesAmtMTD As Decimal
                Dim salesAmtSTD As Decimal

                Try
                    salesAmtMTD = Convert.ToDecimal(salesDataRow(SalesAmtMTDColumnIndex).Value) * 1000
                    salesAmtSTD = Convert.ToDecimal(salesDataRow(SalesAmtSTDColumnIndex).Value) * 1000
                Catch ex As Exception

                End Try
                AddSalesDataRow(storeNo, memo, material, salesAmt, salesAmtMTD, salesAmtSTD)
            End If
        Next
    End Sub

    Private Function isSalesDataRow(ByVal salesDataRow As IRange) As Boolean
        If salesDataRow(StoreNoColumnIndex).Value & "" <> "" Then
            If Integer.TryParse(salesDataRow(StoreNoColumnIndex).Value, Nothing) Then
                Return True
            End If
        End If
        Return False
    End Function

    Private Sub AddSalesDataRow(ByVal storeNo As String, ByVal memo As Boolean, ByVal materialCode As String, ByVal salesAmt As Decimal, ByVal salesAmtMTD As Decimal, ByVal salesAmtSTD As Decimal)
        Dim sdr As DataRow

        sdr = ImportedRetailData.NewRow()
        sdr.Item("CUST_STORE_NO") = storeNo.PadLeft(6, "0")
        sdr.Item("CUST_DEPT_CODE") = DepartmentCode
        sdr.Item("MEMO") = If(memo, "1", "0")
        sdr.Item("MATL_CODE") = materialCode
        sdr.Item("AMT_SOLD") = salesAmt
        sdr.Item("AMT_SOLD_MTD") = salesAmtMTD
        sdr.Item("AMT_SOLD_STD") = salesAmtSTD
        sdr.Item("OPS_YYYYPP") = SalesPeriod
        sdr.Item("OPS_YYYYWW") = SalesWeek
        ImportedRetailData.Rows.Add(sdr)

    End Sub

End Class



Public Class BloomingdalesRetailImport3
    Inherits RetailExcelImport

    Private Const DateCell As String = "A7"

    Private Const StoreNameColumnIndex As Integer = 0
    Private SalesAmtColumnIndex As Integer = 2
    Private SalesAmtMTDColumnIndex As Integer = 12
    Private SalesAmtSTDColumnIndex As Integer = 22


    Public Sub New(ByVal fileName As String, ByVal excelDocument As IWorkbook)
        MyBase.New(fileName, excelDocument)
        retailImportType = retailImportType.BloomingdalesExcel3
        CustomerCode = "BLOOMIES10"
    End Sub

    Public Function GetDateTable() As DataTable

        Dim dateString As String = ImportFile.Worksheets(0).Cells(DateCell).Value

        dateString = dateString.Substring(dateString.IndexOf(":") + 2)
        dateString = dateString.Substring(0, dateString.IndexOf(" "))

        Dim sheetDate As Date = Date.Parse(dateString)

        Dim dateDT As DataTable = ASCDATA1.GetDataTable("SELECT * FROM (SELECT * FROM GLTPARM3 WHERE WEEK_END_DATE < :PARM1 ORDER BY YYYYWW DESC) WHERE ROWNUM < 20", "", "D", New Object() {sheetDate.Date})
        Return dateDT
    End Function

    Protected Overrides Function InitializeForImport() As Boolean
        'Bloomingdales IWorkbooks contain multiple sheets
        'there are sheets for asset vs memo -- these are different departments
        'there may also be separate sheets for gold -- gold sheets are separate
        Return True
    End Function

    Public Overrides Sub Import()
        Throw New Exception("Error importing -- date not specified for Bloomingdales import #3...")
    End Sub

    Public Overloads Sub Import(ByVal YYYYWW As String)
        SalesWeek = YYYYWW
        SalesPeriod = ASCDATA1.GetDataValue("SELECT YYYYPP FROM GLTPARM3 WHERE YYYYWW=:PARM1", "V", New Object() {YYYYWW})

        For Each sht As IWorksheet In ImportFile.Worksheets
            If sht.Visible = SheetVisibility.Visible And Not sht.Name.StartsWith("#7") And Not sht.Name.ToUpper.Contains("TOTAL") And Not sht.Name.ToUpper.Contains("ASSET") Then 'only process visible sheets, and not the totals sheet
                LoadDeptData(sht)
                LoadSalesData(sht)
            End If
        Next
    End Sub

    Private Sub LoadDeptData(sht As IWorksheet)
        Dim deptString As String = sht.Name
        DepartmentCode = deptString
    End Sub

    Private Sub LoadSalesData(sht As IWorksheet)
        Dim sheetName As String = sht.Name
        Dim memo As Boolean = sheetName.ToUpper.Contains("MEMO")
        Dim material As String = DepartmentCode 'material can be derived from dept code

        For Each salesDataRow As IRange In sht.Cells.Rows
            If isSalesDataRow(salesDataRow) Then
                Dim storeNo = If(salesDataRow(StoreNameColumnIndex).Value.ToString() = "TTL COM", "TTL COM", salesDataRow(StoreNameColumnIndex + 1).Value.ToString.PadLeft(6, "0"))
                Dim salesAmt = Convert.ToDecimal(salesDataRow(SalesAmtColumnIndex).Value) * 1000
                Dim salesAmtMTD As Decimal
                Dim salesAmtSTD As Decimal

                Try
                    salesAmtMTD = Convert.ToDecimal(salesDataRow(SalesAmtMTDColumnIndex).Value) * 1000
                    salesAmtSTD = Convert.ToDecimal(salesDataRow(SalesAmtSTDColumnIndex).Value) * 1000
                Catch ex As Exception

                End Try
                AddSalesDataRow(storeNo, memo, material, salesAmt, salesAmtMTD, salesAmtSTD)
            End If
        Next
    End Sub

    Private Function isSalesDataRow(ByVal salesDataRow As IRange) As Boolean
        If salesDataRow(StoreNameColumnIndex).Style.Font.Bold = False Or (salesDataRow(StoreNameColumnIndex).Value IsNot Nothing AndAlso salesDataRow(StoreNameColumnIndex).Value.ToString() = "TTL COM") Then
            If Decimal.TryParse(salesDataRow(SalesAmtMTDColumnIndex).Value, Nothing) Then
                Return True
            End If
        End If
        Return False
    End Function

    Private Sub AddSalesDataRow(ByVal storeNo As String, ByVal memo As Boolean, ByVal materialCode As String, ByVal salesAmt As Decimal, ByVal salesAmtMTD As Decimal, ByVal salesAmtSTD As Decimal)
        Dim sdr As DataRow

        sdr = ImportedRetailData.NewRow()
        sdr.Item("CUST_STORE_NO") = storeNo
        sdr.Item("CUST_DEPT_CODE") = DepartmentCode
        sdr.Item("MEMO") = If(memo, "1", "0")
        sdr.Item("MATL_CODE") = materialCode
        sdr.Item("AMT_SOLD") = salesAmt
        sdr.Item("AMT_SOLD_MTD") = salesAmtMTD
        sdr.Item("AMT_SOLD_STD") = salesAmtSTD
        sdr.Item("OPS_YYYYPP") = SalesPeriod
        sdr.Item("OPS_YYYYWW") = SalesWeek
        ImportedRetailData.Rows.Add(sdr)
    End Sub

End Class

Public Class NeimanMarcusRetailImport
    Inherits RetailExcelImport

    Private Const dateColumnIndex As Integer = 18
    Private Const storeNoColumnIndex As Integer = 3

    Private Const deptCell As String = "F3"

    Public Sub New(ByVal importFileName As String, ByVal excelDocument As IWorkbook)
        MyBase.New(importFileName, excelDocument)
        CustomerCode = "NEIMANM10"
        retailImportType = retailImportType.NeimanMarcusRetailExcel
    End Sub

    Public Overrides Sub Import()
        'Date data is in S column -- maybe check R column for "Wk End"
        For Each sht As IWorksheet In ImportFile.Worksheets
            If sht.Name = "TOTAL DOLLARS" Then
                LoadDateData(sht)
                LoadDeptData(sht)
                LoadSalesData(sht)
            End If
        Next
    End Sub

    Private Sub LoadDateData(ByVal excelWorksheet As IWorksheet)
        Dim weekEndDate As Date

        For i As Integer = 0 To excelWorksheet.Cells.Rows.Count - 1
            If excelWorksheet.Cells(i, dateColumnIndex - 1).Value IsNot Nothing AndAlso excelWorksheet.Cells(i, dateColumnIndex - 1).Value.ToString.Trim() = "WK End:" Then
                Dim dateString As String = excelWorksheet.Cells(i, dateColumnIndex).Value.ToString()
                If dateString.Contains("(") Then dateString = dateString.Substring(0, dateString.IndexOf("(") - 1)
                If Date.TryParse(dateString, weekEndDate) Then
                    Exit For
                End If
            End If
        Next

        Dim dateDataRow As DataRow = ASCDATA1.GetDataRow("SELECT YYYYWW,YYYYPP FROM GLTPARM3 WHERE WEEK_END_DATE=:PARM1", False, "D", New Object() {weekEndDate})
        Me.SalesPeriod = dateDataRow.Item("YYYYPP")
        Me.SalesWeek = dateDataRow.Item("YYYYWW")
    End Sub

    Private Sub LoadDeptData(ByVal excelWorksheet As IWorksheet)
        DepartmentCode = excelWorksheet.Cells(deptCell).Value
    End Sub

    Private Sub LoadSalesData(ByVal excelWorksheet As IWorksheet)
        Dim collectData As Boolean = (DepartmentCode = "622")

        For i As Integer = 6 To excelWorksheet.Cells.Rows.Count
            If collectData Then
                If isSalesDataRow(excelWorksheet.Cells.Rows(i)) Then
                    Dim storeNo As String = excelWorksheet.Cells.Rows(i)(3).Value.ToString.Substring(0, 3)
                    Dim amtSold As Decimal = Decimal.Parse(excelWorksheet.Cells.Rows(i)(9).Value) * 1000

                    AddSalesDataRow(storeNo, amtSold)
                End If

                If excelWorksheet.Cells.Rows(i)(3).Value IsNot Nothing AndAlso (excelWorksheet.Cells.Rows(i)(3).Value.ToString.ToUpper() = "VENDOR TOTALS FOR" Or excelWorksheet.Cells.Rows(i)(3).Value.ToString.ToUpper() = "TOTAL COMPANY SUMMARY" Or excelWorksheet.Cells.Rows(i)(3).Value.ToString.ToUpper() = "TOTAL VENDORS BY STORE") Then
                    Exit For
                End If
            Else
                If excelWorksheet.Cells.Rows(i)(3).Value IsNot Nothing AndAlso excelWorksheet.Cells.Rows(i)(3).Value.ToString.ToUpper() = "TOTAL DEPARTMENTS  BY STORE" Then
                    collectData = True
                End If
            End If
        Next

    End Sub

    Private Sub AddSalesDataRow(ByVal storeNo As String, ByVal amtSold As Integer)
        Dim sdr As DataRow

        sdr = ImportedRetailData.NewRow()
        sdr.Item("CUST_STORE_NO") = storeNo
        sdr.Item("CUST_DEPT_CODE") = DepartmentCode
        sdr.Item("MEMO") = "0"
        sdr.Item("MATL_CODE") = "S"
        sdr.Item("AMT_SOLD") = amtSold
        sdr.Item("OPS_YYYYPP") = SalesPeriod
        sdr.Item("OPS_YYYYWW") = SalesWeek
        ImportedRetailData.Rows.Add(sdr)
    End Sub

    Private Function isSalesDataRow(ByVal excelRow As IRange) As Boolean
        Dim storeInfoString As String = excelRow(3).Value
        Dim x As Decimal
        If Not String.IsNullOrEmpty(storeInfoString) AndAlso storeInfoString.IndexOf("-") = 3 Then
            If Decimal.TryParse(excelRow(9).Value, x) Then
                Return True
            End If
        End If
        Return False
    End Function
End Class

Public Class NeimanMarcusConsignmentImport
    Inherits RetailExcelImport

    Public Sub New(ByVal importFileName As String, ByVal excelDocument As IWorkbook)
        MyBase.New(importFileName, excelDocument)
        CustomerCode = "NEIMANM10"
        retailImportType = retailImportType.NeimanMarcusConsignmentExcel
    End Sub

    Public Overrides Sub Import()
        For Each sht As IWorksheet In ImportFile.Worksheets
            LoadDateData(sht)
            LoadDeptData(sht)
            LoadSalesData(sht)
        Next
    End Sub

    Private Sub LoadDateData(ByVal excelWorksheet As IWorksheet)
        Dim weekEndDate As Date
        Dim dateString As String = excelWorksheet.Cells("A1").Value.ToString() 'Date Range: mm/dd/yyyy - mm/dd/yyyy
        Dim dateSplit = dateString.Split(" ")
        If Not Date.TryParse(dateSplit(5), weekEndDate) Then
            'loop through the rest of date column looking for the sheet date
        End If
        Dim dateDataRow As DataRow = ASCDATA1.GetDataRow("SELECT YYYYWW,YYYYPP FROM GLTPARM3 WHERE WEEK_END_DATE=:PARM1", False, "D", New Object() {weekEndDate})
        Me.SalesPeriod = dateDataRow.Item("YYYYPP")
        Me.SalesWeek = dateDataRow.Item("YYYYWW")
    End Sub

    Private Sub LoadDeptData(ByVal excelWorksheet As IWorksheet)
        DepartmentCode = "WM"
    End Sub

    Private Sub LoadSalesData(ByVal excelWorksheet As IWorksheet)
        For i As Integer = 3 To excelWorksheet.Cells.Rows.Count
            If isSalesDataRow(excelWorksheet.Cells.Rows(i)) Then
                Dim storeNo As String = excelWorksheet.Cells.Rows(i)(1).Value.ToString()
                Dim amtSold As Decimal = Decimal.Parse(excelWorksheet.Cells.Rows(i)(3).Value)
                AddSalesDataRow(storeNo, amtSold)
            End If
        Next
    End Sub

    Private Sub AddSalesDataRow(ByVal storeNo As String, ByVal amtSold As Integer)
        Dim sdr As DataRow

        If ImportedRetailData.Select(String.Format("CUST_STORE_NO='{0}'", storeNo.Replace("'", "''"))).Count > 0 Then
            sdr = ImportedRetailData.Select(String.Format("CUST_STORE_NO='{0}'", storeNo.Replace("'", "''")))(0)
            sdr.Item("AMT_SOLD") += amtSold
        Else
            sdr = ImportedRetailData.NewRow()
            sdr.Item("CUST_STORE_NO") = storeNo
            sdr.Item("CUST_DEPT_CODE") = DepartmentCode
            sdr.Item("MEMO") = "0"
            sdr.Item("MATL_CODE") = "G"
            sdr.Item("AMT_SOLD") = amtSold
            sdr.Item("OPS_YYYYPP") = SalesPeriod
            sdr.Item("OPS_YYYYWW") = SalesWeek
            ImportedRetailData.Rows.Add(sdr)
        End If
    End Sub

    Private Function isSalesDataRow(ByVal excelRow As IRange) As Boolean
        If excelRow(0).Value IsNot Nothing AndAlso excelRow(0).Value.ToString.Trim() = "Total" Then
            Return True
        End If
        Return False
    End Function
End Class

Public Class NeimanMarcusFlashImport
    Inherits RetailExcelImport

    Private Const DateCell As String = "A1" '"as of <Month> <dayNum>, <yearNum>"
    Private Const DateCell2 As String = "A2" '"<Month> Week <wkNum>"

    Private Const StoreNoColumnIndex As Integer = 0
    Private Const SalesAmtMTDColumnIndex As Integer = 2

    Private Const PlanColumnIndex As Integer = 5

    Private Const SalesAmtSTDColumnIndex As Integer = 8

    Private planData As Dictionary(Of String, Decimal)

    Public Sub New(ByVal fileName As String, ByVal excelDocument As IWorkbook)
        MyBase.New(fileName, excelDocument)
        CustomerCode = "NEIMANM10"
        retailImportType = retailImportType.NeimanMarcusFlashExcel
        planData = New Dictionary(Of String, Decimal)
    End Sub

    Protected Overrides Function InitializeForImport() As Boolean
        Return True
    End Function

    Public Overrides Sub Import()
        'Neiman sheets do not contain WTD info, only PTD
        'To extract currect weekly info,subtract the previous "period to date" info (unless we are on week 1)
        'This means we can only use a sheet if we have imported all the prior sheet from the period
        For Each sht As IWorksheet In ImportFile.Worksheets
            If sht.Visible = SheetVisibility.Visible Then
                If sht.Name = "Women" Or sht.Name = "Men" Then
                    LoadDateData(sht)
                    LoadDeptData(sht)
                    LoadPTDSalesData(sht)
                    'SubtractPriorSales()
                End If
            End If
        Next
    End Sub

    Private Sub LoadPTDSalesData(ByVal excelWorksheet As IWorksheet)
        Dim memo As Boolean = False

        Dim material As String = "S" 'Neiman sheets include gold/cinta data as well as total sales data

        For Each salesDataRow As IRange In excelWorksheet.Cells.Rows
            If (salesDataRow(StoreNoColumnIndex).Value & "").ToString.ToUpper.Contains("RECAP") And salesDataRow(SalesAmtMTDColumnIndex).Value & "" = "" Then
                material = ""
            End If
            If material = "" AndAlso ((salesDataRow(StoreNoColumnIndex).Value & "").ToString.ToUpper.Contains("TOTAL GOLD") Or (salesDataRow(StoreNoColumnIndex).Value & "").ToString.ToUpper.Contains("GOLD RECAP")) Then
                material = "G"
            End If
            If material = "G" AndAlso (salesDataRow(StoreNoColumnIndex).Value & "").ToString.ToUpper.Contains("TOTAL CINTA") Then
                material = "C"
            End If
            If material = "" AndAlso isSalesDataRow(salesDataRow) Then
                Dim storeNo = Convert.ToString(salesDataRow(StoreNoColumnIndex).Value)
                If storeNo.Contains("- CLOSED") Then
                    Dim removeIndex = storeNo.IndexOf("- CLOSED")
                    storeNo = storeNo.Substring(0, removeIndex)
                End If
                storeNo = storeNo.Trim()
                Dim planAmt = Convert.ToDecimal(salesDataRow(PlanColumnIndex).Value) * 1000
                planData.Add(storeNo, planAmt)
            End If
            If material <> "" AndAlso isSalesDataRow(salesDataRow) Then
                Dim storeNo = Convert.ToString(salesDataRow(StoreNoColumnIndex).Value)
                If storeNo.Contains("- CLOSED") Then
                    Dim removeIndex = storeNo.IndexOf("- CLOSED")
                    storeNo = storeNo.Substring(0, removeIndex)
                End If
                storeNo = storeNo.Trim()
                Dim salesAmtMtd = Convert.ToDecimal(salesDataRow(SalesAmtMTDColumnIndex).Value) * 1000
                Dim salesAmtStd = Convert.ToDecimal(salesDataRow(SalesAmtSTDColumnIndex).Value) * 1000
                Dim planAmt = Convert.ToDecimal(salesDataRow(PlanColumnIndex).Value) * 1000

                If material = "G" Or material = "C" Then
                    If planData.ContainsKey(storeNo) Then
                        planAmt = planData(storeNo)
                        planData(storeNo) = 0
                    End If
                    ImportedRetailData.Rows.Find(New Object() {storeNo, DepartmentCode, "S"}).Item("AMT_SOLD_MTD") -= salesAmtMtd
                    ImportedRetailData.Rows.Find(New Object() {storeNo, DepartmentCode, "S"}).Item("AMT_SOLD_STD") -= salesAmtStd
                    ImportedRetailData.Rows.Find(New Object() {storeNo, DepartmentCode, "S"}).Item("AMT_PLAN") -= planAmt
                End If

                AddSalesDataRow(storeNo, memo, material, salesAmtMtd, salesAmtStd, planAmt)

                If material = "G" Or material = "C" And planData.ContainsKey(storeNo) Then
                    planData(storeNo) = 0
                End If
            End If
        Next
    End Sub

    Private Sub AddSalesDataRow(ByVal storeNo As String, ByVal memo As Boolean, ByVal material As String, ByVal salesAmtMtd As Decimal, ByVal salesAmtStd As Decimal, ByVal planAmt As Decimal)
        Dim sdr As DataRow

        sdr = ImportedRetailData.NewRow()
        sdr.Item("CUST_STORE_NO") = storeNo
        sdr.Item("CUST_DEPT_CODE") = DepartmentCode
        sdr.Item("MEMO") = If(memo, "1", "0")
        sdr.Item("MATL_CODE") = material
        sdr.Item("AMT_SOLD_MTD") = salesAmtMtd
        sdr.Item("AMT_SOLD_STD") = salesAmtStd
        sdr.Item("AMT_PLAN") = planAmt
        sdr.Item("OPS_YYYYPP") = SalesPeriod
        sdr.Item("OPS_YYYYWW") = SalesWeek
        ImportedRetailData.Rows.Add(sdr)
    End Sub

    Private Function isSalesDataRow(ByVal salesDataRow As IRange) As Boolean
        If salesDataRow(StoreNoColumnIndex).Value & "" <> "" Then
            If Not salesDataRow(StoreNoColumnIndex).Value.ToString.ToUpper.Contains("TOTAL") Then
                If Not salesDataRow(StoreNoColumnIndex).Value.ToString.ToUpper.Trim = "OTHER" Then
                    Dim x As Decimal
                    If Decimal.TryParse(salesDataRow(SalesAmtMTDColumnIndex).Value, x) Then
                        Return True
                    End If
                End If
            End If
        End If
        Return False
    End Function

    Private Sub LoadDateData(ByVal excelWorksheet As IWorksheet)
        Dim dateString As String = excelWorksheet.Cells(DateCell).Value
        Dim splitDate As String() = dateString.Split(" ")

        Dim dateString2 As String = excelWorksheet.Cells(DateCell2).Value
        Dim splitDate2 As String() = dateString2.Split(" ")

        Dim yyyy As Integer = Convert.ToInt32(splitDate(splitDate.Length - 1))
        Dim month As String = splitDate2(0).Substring(0, 3)
        Dim weekNo As Integer = Convert.ToInt32(splitDate2(2))
        Dim mm As Integer
        Try
            mm = DateTime.ParseExact(month, "MMM", CultureInfo.InvariantCulture).Month
        Catch ex As Exception
            mm = DateTime.ParseExact(month, "MMMM", CultureInfo.InvariantCulture).Month
        End Try

        Dim YYYYMM As String = String.Format(yyyy.ToString() & mm.ToString.PadLeft(2, "0"))
        Dim drDateInfo As DataRow = ASCDATA1.GetDataRow("SELECT * FROM GLTPARM3 WHERE YYYYMM=:PARM1 AND REL_WEEK=:PARM2", "VN", New Object() {YYYYMM, weekNo})

        If SalesWeek <> "" And SalesWeek <> drDateInfo.Item("YYYYWW") Then
            Throw New Exception("File contains data for multiple weeks")
        End If
        SalesWeek = drDateInfo.Item("YYYYWW")
        SalesPeriod = drDateInfo.Item("YYYYPP")
    End Sub

    Private Sub LoadDeptData(ByVal excelWorksheet As IWorksheet)
        DepartmentCode = excelWorksheet.Name
    End Sub

End Class

Public Class SaksRetailImportExcel
    Inherits RetailExcelImport

    Private Const DateCell As String = "B9"
    Private Const DeptCell As String = "A7"

    Private Const StoreNoColumnIndex As Integer = 0
    Private Const SalesAmtMTDColumnIndex As Integer = 4
    Private Const SalesAmtSTDColumnIndex As Integer = 11
    Private Const PlanAmtColumnIndex As Integer = 8

    Public Sub New(ByVal fileName As String, ByVal excelDocument As IWorkbook)
        MyBase.New(fileName, excelDocument)
        CustomerCode = "SAKSFIF10"
        retailImportType = retailImportType.SaksExcel
    End Sub

    Public Overrides Sub Import()
        For Each sht As IWorksheet In ImportFile.Worksheets
            If sht.Visible = SheetVisibility.Visible Then
                LoadDateData(sht)
                LoadDeptData(sht)
                LoadSalesData(sht)
            End If
        Next
    End Sub

    Private Sub LoadSalesData(ByVal excelWorksheet As IWorksheet)
        For Each excelRow As IRange In excelWorksheet.Cells.Rows
            If isSalesDataRow(excelRow) Then
                Dim storeNo As String = excelRow(StoreNoColumnIndex).Value.ToString().Split("-")(0).Trim()

                Dim salesAmtMTD As Decimal = Decimal.Parse(excelRow(SalesAmtMTDColumnIndex).Value) * 1000
                Dim salesAmtSTD As Decimal = Decimal.Parse(If(excelRow(SalesAmtSTDColumnIndex).Value, 0)) * 1000
                Dim planAmt As Decimal = Decimal.Parse(excelRow(PlanAmtColumnIndex).Value) * 1000

                AddSalesDataRow(storeNo, salesAmtMTD, salesAmtSTD, planAmt)
            End If
        Next
    End Sub

    Private Sub AddSalesDataRow(ByVal storeNo As String, ByVal salesAmtMtd As Decimal, ByVal salesAmtStd As Decimal, ByVal planAmt As Decimal)
        Dim sdr As DataRow = ImportedRetailData.NewRow()
        sdr.Item("CUST_STORE_NO") = storeNo.PadLeft(6, "0")
        sdr.Item("CUST_DEPT_CODE") = DepartmentCode
        sdr.Item("AMT_SOLD_MTD") = salesAmtMtd
        sdr.Item("AMT_SOLD_STD") = salesAmtStd
        sdr.Item("AMT_PLAN") = planAmt
        sdr.Item("MEMO") = "0"
        sdr.Item("MATL_CODE") = DepartmentCode
        sdr.Item("OPS_YYYYPP") = SalesPeriod
        sdr.Item("OPS_YYYYWW") = SalesWeek
        ImportedRetailData.Rows.Add(sdr)
    End Sub

    Private Function isSalesDataRow(ByVal excelRow As IRange) As Boolean
        If excelRow(StoreNoColumnIndex).Style.Font.Bold = False Then
            Return Decimal.TryParse(excelRow(SalesAmtMTDColumnIndex).Value, Nothing) 'AndAlso Decimal.TryParse(excelRow(SalesAmtSTDColumnIndex).Value, Nothing)
        End If
        Return False
    End Function

    Private Sub LoadDeptData(ByVal excelWorksheet As IWorksheet)
        If excelWorksheet.Cells(DeptCell).Value IsNot Nothing Then
            DepartmentCode = excelWorksheet.Cells(DeptCell).Value.ToString.Split("-")(0).Trim()
        Else
            DepartmentCode = "001"
        End If
    End Sub

    Private Sub LoadDateData(ByVal excelWorksheet As IWorksheet)
        Dim shtDate As Date = excelWorksheet.Cells(DateCell).Value

        Dim drDateInfo As DataRow = ASCDATA1.GetDataRow("SELECT * FROM GLTPARM3 WHERE WEEK_END_DATE=:PARM1", "D", New Object() {shtDate})

        If drDateInfo Is Nothing Then
            drDateInfo = ASCDATA1.GetDataRow("SELECT * FROM GLTPARM3 WHERE WEEK_END_DATE+1=:PARM1", "D", New Object() {shtDate})
        End If

        If drDateInfo Is Nothing Then
            drDateInfo = ASCDATA1.GetDataRow("SELECT * FROM GLTPARM3 WHERE WEEK_END_DATE-1=:PARM1", "D", New Object() {shtDate})
        End If

        SalesWeek = drDateInfo.Item("YYYYWW")
        SalesPeriod = drDateInfo.Item("YYYYPP")
    End Sub
End Class

Public Class SaksConsignmentImportExcel
    Inherits RetailExcelImport

    Public Sub New(ByVal fileName As String, ByVal excelDocument As IWorkbook)
        MyBase.New(fileName, excelDocument)
        CustomerCode = "SAKSFIF10"
        DepartmentCode = "CON"
        retailImportType = retailImportType.SaksConsignmentExcel
        ImportedRetailData.PrimaryKey = New DataColumn() {ImportedRetailData.Columns("YYYYWW"), ImportedRetailData.Columns("CUST_STORE_NO"), ImportedRetailData.Columns("CUST_DEPT_CODE")}
    End Sub

    Public Overrides Sub Import()
        Dim sht As IWorksheet = ImportFile.Worksheets(ImportFile.Worksheets.Count - 1)
        LoadDateData(sht)
        LoadSalesData(sht)
    End Sub

    Public Overloads Sub Import(ByVal worksheetName As String)
        Dim sht As IWorksheet = ImportFile.Worksheets(worksheetName)
        LoadDateData(sht)
        LoadSalesData(sht)
    End Sub

    Public Function GetWorksheets() As DataTable
        Dim wsNames As New DataTable()
        wsNames.Columns.Add("Worksheet")
        For Each ws As IWorksheet In ImportFile.Worksheets
            wsNames.Rows.Add(New Object() {ws.Name})
        Next
        Return wsNames
    End Function

    Private Sub LoadSalesData(ByVal excelWorksheet As IWorksheet)
        Dim salesAmtCol As Integer = -1
        For Each excelRow As IRange In excelWorksheet.Cells.Rows
            If salesAmtCol = -1 Then
                For i As Integer = 0 To 7
                    If excelRow(i) IsNot Nothing AndAlso excelRow(i).Value IsNot Nothing AndAlso (excelRow(i).Value.ToString() = "Sales $" Or excelRow(i).Value.ToString().ToUpper().Contains("SALES $ WTD")) Then
                        salesAmtCol = i
                        Exit For
                    End If
                Next
            ElseIf salesAmtCol > 0 AndAlso isSalesDataRow(excelRow, salesAmtCol) Then
                Dim storeNo As String = excelRow(0).Value.ToString().Split("_")(0).PadLeft(6, "0")
                Dim salesAmt As Decimal = excelRow(salesAmtCol).Value

                AddSalesDataRow(storeNo, salesAmt)
            End If
        Next
    End Sub

    Private Sub AddSalesDataRow(ByVal storeNo As String, ByVal salesAmt As Decimal)
        Dim sdr As DataRow = ImportedRetailData.NewRow()
        sdr.Item("CUST_STORE_NO") = storeNo.PadLeft(6, "0")
        sdr.Item("CUST_DEPT_CODE") = DepartmentCode
        sdr.Item("AMT_SOLD") = salesAmt
        sdr.Item("MEMO") = "0"
        sdr.Item("MATL_CODE") = "G"
        sdr.Item("OPS_YYYYPP") = SalesPeriod
        sdr.Item("OPS_YYYYWW") = SalesWeek
        ImportedRetailData.Rows.Add(sdr)
    End Sub

    Private Function isSalesDataRow(ByVal excelRow As IRange, salesAmtCol As Integer) As Boolean
        If excelRow(0).Value IsNot Nothing AndAlso excelRow(salesAmtCol).Value IsNot Nothing Then
            Return True
        End If
        Return False
    End Function


    Private Sub LoadDateData(ByVal excelWorksheet As IWorksheet)
        Dim monthNo As Integer

        Dim mtch = Regex.Match(excelWorksheet.Name, "(\w+) (Wk|Week|week) ?(\d+)", RegexOptions.IgnoreCase)

        Dim culture = New CultureInfo("en-US")
        culture.DateTimeFormat.AbbreviatedMonthNames = New String() {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sept", "Oct", "Nov", "Dec", ""}

        Try
            monthNo = DateTime.ParseExact(mtch.Groups(1).Value, "MMM", culture).Month
        Catch
            monthNo = DateTime.ParseExact(mtch.Groups(1).Value, "MMMM", culture).Month

        End Try

        Dim importYear As Integer = If(Today.Month = 1 And monthNo = 12, Today.Year - 1, Today.Year)


        Dim drDateInfo As DataRow = ASCDATA1.GetDataRow("SELECT * FROM RSTCLND1 WHERE CUST_CODE='SAKSFIF10' AND YM=:PARM1 AND MM=:PARM2 AND REL_WEEK=:PARM3", "NNN", New Object() {importYear, monthNo, mtch.Groups(3).Value})

        SalesWeek = drDateInfo.Item("YYYYWW")
        SalesPeriod = drDateInfo.Item("YYYYPP")
    End Sub
End Class

Public Enum RetailImportType
    SaksPDF
    SaksExcel
    SaksConsignmentExcel
    NeimanMarcusFlashExcel
    NeimanMarcusRetailExcel
    NeimanMarcusConsignmentExcel
    NordstromExcel
    NordstromPDF
    BloomingdalesExcel
    BloomingdalesExcel2
    BloomingdalesExcel3
    HoltRenfrewPDF
End Enum