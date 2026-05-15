Imports Aspose.Cells
Imports System.Globalization

Public Class RetailExport
#Region "Styling"
    Protected Sub SetStyleAlignCenter(ByRef cell As Cell)
        Dim centerAlignStyle = cell.GetStyle()
        centerAlignStyle.HorizontalAlignment = TextAlignmentType.Center
        Dim centerAlignFlags As New StyleFlag()
        centerAlignFlags.HorizontalAlignment = True

        cell.SetStyle(centerAlignStyle, centerAlignFlags)
    End Sub

    Protected Sub SetStyleAlignCenter(ByRef rng As Range)
        Dim centerAlignStyle = rng.Worksheet.Workbook.Styles(rng.Worksheet.Workbook.Styles.Add())
        centerAlignStyle.HorizontalAlignment = TextAlignmentType.Center
        Dim centerAlignFlags As New StyleFlag()
        centerAlignFlags.HorizontalAlignment = True

        rng.ApplyStyle(centerAlignStyle, centerAlignFlags)
    End Sub

    Protected Sub SetStyleAlignCenterVertical(ByRef cell As Cell)
        Dim centerAlignStyle = cell.GetStyle()
        centerAlignStyle.VerticalAlignment = TextAlignmentType.Center
        Dim centerAlignFlags As New StyleFlag()
        centerAlignFlags.VerticalAlignment = True

        cell.SetStyle(centerAlignStyle, centerAlignFlags)
    End Sub

    Protected Sub SetStyleAlignCenterVertical(ByRef rng As Range)
        Dim centerAlignStyle = rng.Worksheet.Workbook.Styles(rng.Worksheet.Workbook.Styles.Add())
        centerAlignStyle.VerticalAlignment = TextAlignmentType.CenterAcross
        Dim centerAlignFlags As New StyleFlag()
        centerAlignFlags.VerticalAlignment = True

        rng.ApplyStyle(centerAlignStyle, centerAlignFlags)
    End Sub

    Protected Sub SetStyleForegroundColor(ByRef cell As Cell, ByVal color As Color)
        Dim foregroundColorStyle = cell.GetStyle()
        foregroundColorStyle.Pattern = BackgroundType.Solid
        foregroundColorStyle.ForegroundColor = color

        Dim foregroundFlag As New StyleFlag()
        foregroundFlag.CellShading = True

        cell.SetStyle(foregroundColorStyle, foregroundFlag)
    End Sub

    Protected Sub SetStyleForegroundColor(ByRef row As Row, ByVal color As Color)
        Dim foregroundColorStyle = row.Style
        foregroundColorStyle.Pattern = BackgroundType.Solid
        foregroundColorStyle.ForegroundColor = color

        Dim foregroundFlag As New StyleFlag()
        foregroundFlag.CellShading = True

        row.ApplyStyle(foregroundColorStyle, foregroundFlag)
    End Sub

    Protected Sub SetStyleBackgroundColor(ByRef cell As Cell, ByVal color As Color)
        Dim backgroundColorStyle = cell.GetStyle()
        backgroundColorStyle.Pattern = BackgroundType.Solid
        backgroundColorStyle.BackgroundColor = color

        Dim backgroundFlag As New StyleFlag()
        backgroundFlag.CellShading = True

        cell.SetStyle(backgroundColorStyle, backgroundFlag)
    End Sub

    Protected Sub SetStyleBackgroundColor(ByRef row As Row, ByVal color As Color)
        Dim backgroundColorStyle = row.Style
        backgroundColorStyle.Pattern = BackgroundType.Solid
        backgroundColorStyle.BackgroundColor = color

        Dim backgroundFlag As New StyleFlag()
        backgroundFlag.CellShading = True

        row.ApplyStyle(backgroundColorStyle, backgroundFlag)
    End Sub

    Protected Sub SetStyleBackgroundColor(ByRef rng As Range, ByVal color As Color)
        Dim backgroundColorStyle = rng.Worksheet.Workbook.Styles(rng.Worksheet.Workbook.Styles.Add())
        backgroundColorStyle.Pattern = BackgroundType.Solid
        backgroundColorStyle.BackgroundColor = color

        Dim backgroundFlag As New StyleFlag()
        backgroundFlag.CellShading = True

        rng.ApplyStyle(backgroundColorStyle, backgroundFlag)
    End Sub

    Protected Sub SetStyleFontBold(ByRef cell As Cell)
        Dim fontBoldStyle = cell.GetStyle()
        fontBoldStyle.Font.IsBold = True

        Dim fontBoldFlag As New StyleFlag()
        fontBoldFlag.FontBold = True

        cell.SetStyle(fontBoldStyle, fontBoldFlag)
    End Sub

    Protected Sub SetStyleFontBold(ByRef rng As Range)
        Dim fontBoldStyle = rng.Worksheet.Workbook.Styles(rng.Worksheet.Workbook.Styles.Add())
        fontBoldStyle.Font.IsBold = True

        Dim fontBoldFlag As New StyleFlag()
        fontBoldFlag.FontBold = True

        rng.ApplyStyle(fontBoldStyle, fontBoldFlag)
    End Sub

    Protected Sub SetStyleFontBold(ByRef row As Row)
        Dim fontBoldStyle = row.Style
        fontBoldStyle.Font.IsBold = True

        Dim fontBoldFlag As New StyleFlag()
        fontBoldFlag.FontBold = True

        row.ApplyStyle(fontBoldStyle, fontBoldFlag)
    End Sub

    Protected Sub SetStyleFontColor(ByRef cell As Cell, ByVal color As Color)
        Dim fontColorStyle = cell.GetStyle()
        fontColorStyle.Font.Color = color

        Dim fontColorFlag As New StyleFlag()
        fontColorFlag.FontColor = True

        cell.SetStyle(fontColorStyle, fontColorFlag)
    End Sub

    Protected Sub SetStyleFontColor(ByRef rng As Range, ByVal color As Color)
        Dim fontColorStyle = rng.Worksheet.Workbook.Styles(rng.Worksheet.Workbook.Styles.Add())
        fontColorStyle.Font.Color = color

        Dim fontColorFlag As New StyleFlag()
        fontColorFlag.FontColor = True

        rng.ApplyStyle(fontColorStyle, fontColorFlag)
    End Sub

    Protected Sub SetStyleForegroundColor(ByRef rng As Range, ByVal color As Color)
        Dim foregroundColorStyle As Style = rng.Worksheet.Workbook.Styles(rng.Worksheet.Workbook.Styles.Add())
        foregroundColorStyle.Pattern = BackgroundType.Solid
        foregroundColorStyle.ForegroundColor = color

        Dim foregroundFlag As New StyleFlag()
        foregroundFlag.CellShading = True

        rng.ApplyStyle(foregroundColorStyle, foregroundFlag)
    End Sub

    Protected Sub SetStyleWrapText(ByRef rng As Range)
        Dim wrapTextStyle As Style = rng.Worksheet.Workbook.Styles(rng.Worksheet.Workbook.Styles.Add())
        wrapTextStyle.IsTextWrapped = True

        Dim wrapTextFlag As New StyleFlag()
        wrapTextFlag.WrapText = True

        rng.ApplyStyle(wrapTextStyle, wrapTextFlag)
    End Sub

    Protected Sub SetStyleWrapText(ByRef cell As Cell)
        Dim wrapTextStyle As Style = cell.GetStyle()
        wrapTextStyle.IsTextWrapped = True

        Dim wrapTextFlag As New StyleFlag()
        wrapTextFlag.WrapText = True

        cell.SetStyle(wrapTextStyle, wrapTextFlag)
    End Sub

    Protected Sub SetNumberFormat(ByRef cell As Cell, ByVal formatString As String)
        Dim numberFormatFlag As New StyleFlag()
        numberFormatFlag.NumberFormat = True

        Dim customNumberFormat As Style = cell.GetStyle()
        customNumberFormat.Custom = formatString

        cell.SetStyle(customNumberFormat, numberFormatFlag)
    End Sub

    Protected Sub SetNumberFormat(ByRef rng As Range, ByVal formatString As String)
        Dim numberFormatFlag As New StyleFlag()
        numberFormatFlag.NumberFormat = True

        Dim customNumberFormat As Style = rng.Worksheet.Workbook.Styles(rng.Worksheet.Workbook.Styles.Add())
        customNumberFormat.Custom = formatString

        rng.ApplyStyle(customNumberFormat, numberFormatFlag)
    End Sub

    Protected Sub SetThousandsFormat(ByVal cell As Cell)
        SetNumberFormat(cell, "$#,##0.0,")
    End Sub

    Protected Sub SetThousandsFormat(ByVal rng As Range)
        SetNumberFormat(rng, "$#,##0,0,")
    End Sub

    Protected Sub SetDollarFormat(ByVal cell As Cell)
        SetNumberFormat(cell, "$#,##0")
    End Sub

    Protected Sub SetDollarFormat(ByVal rng As Range)
        SetNumberFormat(rng, "$#,##0")
    End Sub

    Protected Sub SetPercentageFormat(ByVal cell As Cell)
        SetNumberFormat(cell, "#0.#%")
    End Sub

    Protected Sub SetPercentageFormat(ByVal rng As Range)
        SetNumberFormat(rng, "#0.#%")
    End Sub

    Protected Function GetDistinct(ByVal columnName As String, ByVal sourceTable As DataTable) As List(Of String)
        Dim view As New DataView(sourceTable)
        Dim departments As DataTable = view.ToTable(True, columnName)
        Return departments.AsEnumerable().Select(Function(row) row(columnName).ToString()).ToList()
    End Function

    Protected Function GetExcelColumnName(ByVal columnNumber As Integer) As String
        Dim dividend As Integer = columnNumber
        Dim columnName As String = String.Empty
        Dim modulo As Integer

        While (dividend > 0)
            modulo = (dividend - 1) Mod 26
            columnName = Convert.ToChar(65 + modulo).ToString() + columnName
            dividend = CType((dividend - modulo) / 26, Integer)
        End While

        Return columnName
    End Function

    Protected Function ConvertFromR1C1(ByVal rowNumber As Integer, ByVal columnNumber As Integer) As String
        Return (GetExcelColumnName(columnNumber) & rowNumber)
    End Function

#End Region
End Class

Public Class WeeklyFlashByAccount
    Inherits RetailExport

    Dim _customerCode As String
    Dim _reportYYYYWW As String
    Dim _reportData As DataTable
    Dim _showWTD As Boolean

    Dim nameCol, dsaCol, wtdTYcol, wtdLYcol, mtdTYcol, mtdLYcol, pctMTDcol, mtdPlanCol, pctMTDplanCol, stdTYcol, stdLYcol, pctSTDcol, stdPlanCol, pctSTDplanCol As String

    Public Sub New(ByVal customerCode As String, ByVal reportYYYYWW As String)
        _customerCode = customerCode
        _reportYYYYWW = reportYYYYWW
    End Sub

    Protected Shared Sub SetAsposeLicense()
        Dim license As Aspose.Cells.License = New Aspose.Cells.License()
        license.SetLicense("Aspose.Total.lic")
    End Sub

    Private Sub SetColumns()
        nameCol = "A"
        dsaCol = "B"
        wtdTYcol = "C"
        wtdLYcol = "D"
        Dim wtdAdjust As Integer = 0
        If _showWTD = True Then
            wtdAdjust = 2
        End If
        mtdTYcol = Chr(Asc("C") + wtdAdjust)
        mtdLYcol = Chr(Asc("D") + wtdAdjust)
        pctMTDcol = Chr(Asc("E") + wtdAdjust)
        mtdPlanCol = Chr(Asc("F") + wtdAdjust)
        pctMTDplanCol = Chr(Asc("G") + wtdAdjust)

        stdTYcol = Chr(Asc("I") + wtdAdjust)
        stdLYcol = Chr(Asc("J") + wtdAdjust)
        pctSTDcol = Chr(Asc("K") + wtdAdjust)
        stdPlanCol = Chr(Asc("L") + wtdAdjust)
        pctSTDplanCol = Chr(Asc("M") + wtdAdjust)
    End Sub

    Public Sub CreateReport(ByVal fileName As String, Optional ByVal showWTDbyStore As Boolean = False)
        GetReportData()
        SetAsposeLicense()
        Dim reportWorkbook As Workbook = New Workbook()
        reportWorkbook.Worksheets.RemoveAt(0)

        _showWTD = showWTDbyStore
        SetColumns()

        reportWorkbook.DefaultStyle.Font.Name = "Times New Roman"

        Dim departments = GetDistinct("DEPT_CODE", _reportData)
        If departments.Count = 2 Then
            departments.Add("B")
        End If

        For Each department As String In departments
            Dim reportSheet As Worksheet = reportWorkbook.Worksheets.Add(department)
            reportSheet.PageSetup.FitToPagesWide = 1
            reportSheet.PageSetup.FitToPagesTall = Nothing
            WriteDepartmentSheet(department, reportSheet)
            reportSheet.AutoFitColumn(0)
        Next

        Try
            reportWorkbook.Save(fileName, SaveFormat.Xlsx)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub GetReportData()
        _reportData = ASCDATA1.GetDataTable("SELECT * FROM RSTFLASH WHERE CUST_CODE=:PARM1 AND YYYYWW=:PARM2", "", "VV", New Object() {_customerCode, _reportYYYYWW})
    End Sub

    Private Sub WriteDepartmentSheet(ByVal departmentCode As String, ByVal reportSheet As Worksheet)
        Dim departmentName As String

        If departmentCode <> "B" Then
            departmentName = ASCDATA1.GetDataValue("SELECT DEPT_DESC FROM ICTDEPT1 WHERE DEPT_CODE=:PARM1", "V", New Object() {departmentCode})
        Else
            departmentName = "All Departments"
        End If

        reportSheet.Name = departmentName

        WriteReportHeader(departmentName, reportSheet)
        WriteColumnHeadings(True, 3, reportSheet)

        Dim materialsToReport = New String() {"S", "G", "C", "T"}

        Dim summaryRows As New List(Of Integer)
        Dim currentRow As Integer = 5
        currentRow += 1
        For Each rgn As String In GetDistinct("REGION_DESC", _reportData.AsEnumerable().Where(Function(row) (row.Item("DEPT_CODE") = departmentCode Or departmentCode = "B") And row.Item("CUST_STORE_NAME") <> "DIRECT").CopyToDataTable()).OrderBy(Function(x) If(x = "OTHER", "ZZZ", x))
            Dim rgnStartRow As Integer = currentRow
            currentRow = WriteDataRows(departmentCode, rgn, materialsToReport, -1, True, currentRow, reportSheet)
            currentRow = WriteSubTotalRow(rgn, rgnStartRow, True, currentRow, reportSheet)
            summaryRows.Add(currentRow - 1)
            currentRow += 1
        Next

        currentRow = WriteTotals("All Stores", currentRow, summaryRows, reportSheet)

        If _reportData.AsEnumerable().Any(Function(row) (row.Item("DEPT_CODE") = departmentCode Or departmentCode = "B") And row.Item("CUST_STORE_NAME") = "DIRECT") Then
            If (_customerCode = "SAKSFIF10" Or _customerCode = "NEIMANM10") And departmentCode <> "B" Then
                'Break out of this section if all DIRECT data not filled in
                Dim sql As String = "SELECT " & _
                                "  COUNT(*) CNT, NVL(MAX(C1.REL_WEEK),1) REL_WEEK " & _
                                "FROM " & _
                                "  RSTRETL5 R5 " & _
                                "JOIN RSTCLND1 C1 " & _
                                "ON (R5.CUST_CODE=C1.CUST_CODE AND C1.YYYYWW =:PARM1) " & _
                                "WHERE  R5.CUST_CODE=:PARM2 AND R5.CUST_STORE_NO='DIRECT' " & _
                                "AND R5.DEPT_CODE=:PARM3 AND R5.OPS_YYYYWW BETWEEN " & _
                                " (SELECT MIN(YYYYWW) FROM RSTCLND1 WHERE CUST_CODE=:PARM2 " & _
                                "    AND CUST_YYYYPP=(SELECT CUST_YYYYPP FROM RSTCLND1 WHERE CUST_CODE=:PARM2 AND YYYYWW =:PARM1)) " & _
                                "AND :PARM1"
                Dim dr As DataRow = ASCDATA1.GetDataRow(sql, "VVV", New String() {_reportYYYYWW, _customerCode, departmentCode})
                If dr.Item("CNT") < dr.Item("REL_WEEK") Then
                    GoTo BypassDirect
                End If
            End If

            Try
                summaryRows.Clear()
                summaryRows.Add(currentRow - 1)
                currentRow += 1
                summaryRows.Add(currentRow)
                currentRow = WriteDirect(departmentCode, materialsToReport, currentRow, summaryRows, reportSheet)
                currentRow += 1

                currentRow = WriteTotals("Stores/Direct", currentRow, summaryRows, reportSheet)
            Catch ex As Exception
                'There is no sales info for DIRECT
            End Try
        End If
BypassDirect:

        currentRow += 1
        currentRow = WriteTotalWTD(departmentCode, currentRow, reportSheet)

        reportSheet.HorizontalPageBreaks.Add(currentRow + 1)


        If _reportData.AsEnumerable().Any(Function(row) (row.Item("DEPT_CODE") = departmentCode Or departmentCode = "B") And New String() {"G", "C", "T"}.Contains(row.Item("MATL_CODE"))) Then 'CONSIGNMENT DATA EXISTS
            currentRow += 2

            reportSheet.Cells.CreateRange(currentRow, 0, 1, 13).SetOutlineBorder(BorderType.TopBorder, CellBorderType.Medium, Color.Black)
            reportSheet.Cells(currentRow, 0).Value = "CONSIGNMENT RECAP - THESE NUMBERS ARE INCLUDED IN TOTAL SALES RECAP ABOVE"
            SetStyleFontColor(reportSheet.Cells(currentRow, 0), Color.Red)
            SetStyleFontBold(reportSheet.Cells(currentRow, 0))
            currentRow += 1
            currentRow = WriteColumnHeadings(True, currentRow, reportSheet)

            Dim startRow = currentRow
            currentRow = WriteDataRows(departmentCode, "", New String() {"G", "C", "T"}, 0, True, currentRow, reportSheet)

            WriteSubTotalRow("Consignment", startRow, True, currentRow, reportSheet)

            currentRow += 2

            'Do Gold/Cinta breakdown if Cinta data exists
            If currentRow - startRow >= 20 Then
                reportSheet.HorizontalPageBreaks.Add(currentRow - 1)
            End If

            If _reportData.AsEnumerable().Any(Function(row) (row.Item("DEPT_CODE") = departmentCode Or departmentCode = "B") And row.Item("MATL_CODE") = "C") _
                Or _reportData.AsEnumerable().Any(Function(row) (row.Item("DEPT_CODE") = departmentCode Or departmentCode = "B") And row.Item("MATL_CODE") = "T") Then
                reportSheet.Cells(currentRow, 0).Value = "TOTAL GOLD"
                SetStyleFontColor(reportSheet.Cells(currentRow, 0), Color.Red)
                SetStyleFontBold(reportSheet.Cells(currentRow, 0))
                currentRow = WriteColumnHeadings(False, currentRow, reportSheet)
                startRow = currentRow
                currentRow = WriteDataRows(departmentCode, "", New String() {"G"}, 0, False, currentRow, reportSheet)
                WriteSubTotalRow("Gold", startRow, False, currentRow, reportSheet)

                If _reportData.AsEnumerable().Any(Function(row) (row.Item("DEPT_CODE") = departmentCode Or departmentCode = "B") And row.Item("MATL_CODE") = "C") Then
                    currentRow += 2

                    reportSheet.Cells(currentRow, 0).Value = "TOTAL CINTA"
                    SetStyleFontColor(reportSheet.Cells(currentRow, 0), Color.Red)
                    SetStyleFontBold(reportSheet.Cells(currentRow, 0))
                    currentRow = WriteColumnHeadings(False, currentRow, reportSheet)
                    startRow = currentRow
                    currentRow = WriteDataRows(departmentCode, "", New String() {"C"}, 0, False, currentRow, reportSheet)
                    WriteSubTotalRow("Cinta", startRow, False, currentRow, reportSheet)
                End If

                If _reportData.AsEnumerable().Any(Function(row) (row.Item("DEPT_CODE") = departmentCode Or departmentCode = "B") And row.Item("MATL_CODE") = "T") Then
                    currentRow += 2

                    reportSheet.Cells(currentRow, 0).Value = "TOTAL TRUNK"
                    SetStyleFontColor(reportSheet.Cells(currentRow, 0), Color.Red)
                    SetStyleFontBold(reportSheet.Cells(currentRow, 0))
                    currentRow = WriteColumnHeadings(False, currentRow, reportSheet)
                    startRow = currentRow
                    currentRow = WriteDataRows(departmentCode, "", New String() {"T"}, 0, False, currentRow, reportSheet)
                    WriteSubTotalRow("Trunk", startRow, False, currentRow, reportSheet)
                End If

            End If
        Else
            'No Consignment Data
        End If
    End Sub


    Private Function WriteTotalWTD(ByVal departmentCode As String, ByVal currentRow As Integer, ByVal reportSheet As Worksheet) As Integer
        Dim storesWTD As Decimal
        Dim directWTD As Decimal
        Dim storesWTD_LY As Decimal
        Dim directWTD_LY As Decimal


        Try
            storesWTD = _reportData.AsEnumerable.Where(Function(row) row.Item("CUST_STORE_NAME") <> "DIRECT" And (row.Item("DEPT_CODE") = departmentCode Or departmentCode = "B")).Sum(Function(row) row.Field(Of Decimal?)("AMT_SOLD_WTD_TY"))
            directWTD = _reportData.AsEnumerable.Where(Function(row) row.Item("CUST_STORE_NAME") = "DIRECT" And (row.Item("DEPT_CODE") = departmentCode Or departmentCode = "B")).Sum(Function(row) row.Field(Of Decimal?)("AMT_SOLD_WTD_TY"))
        Catch ex As Exception

        End Try

        Try
            storesWTD_LY = _reportData.AsEnumerable.Where(Function(row) row.Item("CUST_STORE_NAME") <> "DIRECT" And (row.Item("DEPT_CODE") = departmentCode Or departmentCode = "B")).Sum(Function(row) row.Field(Of Decimal?)("AMT_SOLD_WTD_LY"))
            directWTD_LY = _reportData.AsEnumerable.Where(Function(row) row.Item("CUST_STORE_NAME") = "DIRECT" And (row.Item("DEPT_CODE") = departmentCode Or departmentCode = "B")).Sum(Function(row) row.Field(Of Decimal?)("AMT_SOLD_WTD_LY"))
        Catch ex As Exception

        End Try

        reportSheet.Cells(currentRow - 1, 1).Value = "TY"
        reportSheet.Cells(currentRow - 1, 2).Value = "LY"
        reportSheet.Cells(currentRow - 1, 3).Value = "% Chg"
        reportSheet.Cells(currentRow, 0).Value = "WTD Stores: "
        reportSheet.Cells(currentRow, 1).Value = storesWTD
        SetThousandsFormat(reportSheet.Cells(currentRow, 1))

        reportSheet.Cells(currentRow, 2).Value = storesWTD_LY
        SetThousandsFormat(reportSheet.Cells(currentRow, 2))
        reportSheet.Cells(currentRow, 3).Formula = String.Format("=(B{0}-C{0})/ABS(C{0})", currentRow + 1)
        SetPercentageFormat(reportSheet.Cells(currentRow, 3))

        If directWTD <> 0 Or directWTD_LY <> 0 Then

            reportSheet.Cells(currentRow + 1, 0).Value = "WTD Direct: "
            reportSheet.Cells(currentRow + 1, 1).Value = directWTD
            SetThousandsFormat(reportSheet.Cells(currentRow + 1, 1))
            reportSheet.Cells(currentRow + 1, 2).Value = directWTD_LY
            SetThousandsFormat(reportSheet.Cells(currentRow + 1, 2))
            reportSheet.Cells(currentRow + 1, 3).Formula = String.Format("=(B{0}-C{0})/ABS(C{0})", currentRow + 2)
            SetPercentageFormat(reportSheet.Cells(currentRow + 1, 3))
        Else
            Return currentRow + 1
        End If

        Return currentRow + 2
    End Function

    Private Function WriteDirect(ByVal departmentCode As String, ByVal materialsToReport As String(), ByVal currentRow As Object, ByVal summaryRows As List(Of Integer), ByVal reportSheet As Worksheet) As Integer
        currentRow = WriteDataRows(departmentCode, "DIRECT", materialsToReport, 1, True, currentRow, reportSheet)
        Dim cellsToStyle As Range = reportSheet.Cells.CreateRange("A" & currentRow - 1, "M" & currentRow - 1)
        SetTotalsStyle(cellsToStyle)
        Return currentRow
    End Function

    Private Sub WriteReportHeader(ByVal departmentName As String, ByVal reportSheet As Worksheet)

        Dim dateRow As DataRow = ASCDATA1.GetDataRow("SELECT TO_CHAR(WEEK_END_DATE,'FMMonth DD, YYYY') AS_OF_DATE, " & _
                                                     " TO_CHAR(TO_DATE(YYYYMM,'YYYYMM'),'FMMonth') || ' Week ' || REL_WEEK WK_DESC FROM GLTPARM3 WHERE YYYYWW=:PARM1", False, "V", New Object() {_reportYYYYWW})
        reportSheet.Cells("M1").Value = " " 'tell Gembox that we have at least M columns
        reportSheet.Cells("M2").Value = " " 'tell Gembox that we have at least M columns
        reportSheet.Cells("M3").Value = " " 'tell Gembox that we have at least M columns

        reportSheet.Cells("A1").Value = "as of " & dateRow("AS_OF_DATE")
        SetHeaderFont("A1", reportSheet)

        dateRow.Table.Columns("WK_DESC").ReadOnly = False

        If _customerCode = "NEIMANM10" Then 'might need to adjust week
            If _reportYYYYWW.EndsWith("53") Or _
                (Convert.ToInt32(_reportYYYYWW.Substring(4, 2)) <= 26 AndAlso New String() {"2001", "2007", "2013", "2018", "2023", "2028"}.Contains(_reportYYYYWW.Substring(0, 4))) Then 'year after a 53 week year or 53rd week
                If _reportYYYYWW.EndsWith("26") Then
                    dateRow("WK_DESC") = "July Week 5"
                Else
                    dateRow = ASCDATA1.GetDataRow("SELECT TO_CHAR(WEEK_END_DATE,'FMMonth DD, YYYY') AS_OF_DATE, " & _
                                                     " TO_CHAR(TO_DATE(YYYYMM,'YYYYMM'),'FMMonth') || ' Week ' || REL_WEEK WK_DESC FROM GLTPARM3 WHERE YYYYWW=:PARM1", False, "V", New Object() {ASCMAIN1.Week_Calc(_reportYYYYWW, 1)})
                End If
            End If
        End If

        Dim cellsToStyle As Range = reportSheet.Cells.CreateRange("C1", "M3")
        Dim rangeStyle As Style = reportSheet.Workbook.Styles(reportSheet.Workbook.Styles.Add())
        rangeStyle.Pattern = BackgroundType.Solid
        rangeStyle.ForegroundColor = Color.FromArgb(255, 255, 255, 153)
        rangeStyle.SetBorder(BorderType.Horizontal, CellBorderType.None, Nothing)
        Dim rangeFlags As New StyleFlag()
        rangeFlags.CellShading = True
        rangeFlags.Borders = True

        cellsToStyle.ApplyStyle(rangeStyle, rangeFlags)


        reportSheet.Cells.SetColumnWidth(0, 20.29)


        Dim bottomBorderStyle As New Style()
        bottomBorderStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black)
        Dim borderFlag As New StyleFlag()
        borderFlag.BottomBorder = True

        Dim centerAlignStyle As New Style()
        centerAlignStyle.HorizontalAlignment = TextAlignmentType.Center
        Dim alignFlag As New StyleFlag()
        alignFlag.HorizontalAlignment = True

        For i As Integer = 0 To 12
            If i <> 0 And i <> 1 And i <> 7 Then reportSheet.Cells.SetColumnWidth(i, 11)
            If i <> 0 Then reportSheet.Cells.Columns(i).ApplyStyle(centerAlignStyle, alignFlag)
            reportSheet.Cells(2, i).SetStyle(bottomBorderStyle, borderFlag)
        Next

        Dim rightBorderStyle As New Style()
        rightBorderStyle.SetBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black)

        Dim rightBorderFlag As New StyleFlag()
        rightBorderFlag.RightBorder = True

        reportSheet.Cells("B1").SetStyle(rightBorderStyle, rightBorderFlag)
        reportSheet.Cells("B2").SetStyle(rightBorderStyle, rightBorderFlag)
        reportSheet.Cells("B3").SetStyle(rightBorderStyle, rightBorderFlag)


        reportSheet.Cells("A2").Value = dateRow("WK_DESC")
        SetHeaderFont("A2", reportSheet)

        Dim custRow As DataRow = ASCDATA1.GetDataRow("SELECT CUST_NAME FROM ARTCUST1 WHERE CUST_CODE=:PARM1", False, "V", New Object() {_customerCode})
        reportSheet.Cells("H1").Value = String.Format("{0} {1} Flash Reporting", custRow("CUST_NAME"), departmentName)
        SetHeaderFont("H1", reportSheet)

        reportSheet.Cells("H2").Value = "By Store"
        SetHeaderFont("H2", reportSheet)

        reportSheet.Cells("H3").Value = "Store Summary"
        SetHeaderFont("H3", reportSheet)

        reportSheet.FreezePanes(5, 1, 5, 1)
    End Sub

    Private Sub SetHeaderFont(ByVal cellString As String, ByVal reportSheet As Worksheet)
        Dim headerStyle As New Style()
        headerStyle.Font.IsBold = True
        headerStyle.Font.Size = 9

        Dim fontFlag As New StyleFlag()
        fontFlag.FontBold = True
        fontFlag.FontSize = True

        reportSheet.Cells(cellString).SetStyle(headerStyle, fontFlag)
    End Sub

    Private Function WriteColumnHeadings(ByVal includePlan As Boolean, ByVal rowIndex As Integer, ByVal reportsheet As Worksheet) As Integer

        Dim wtdAdjust As Integer = If(_showWTD, 2, 0)

        If includePlan Then
            reportsheet.Cells.CreateRange(rowIndex, 2, 1, 4 + wtdAdjust).Merge()
            reportsheet.Cells.CreateRange(rowIndex, 8 + wtdAdjust, 1, 4).Merge()
        Else
            reportsheet.Cells.CreateRange(rowIndex, 2, 1, 3 + wtdAdjust).Merge()
            reportsheet.Cells.CreateRange(rowIndex, 8 + wtdAdjust, 1, 3).Merge()
        End If

        Dim boldStyle As New Style()
        boldStyle.Font.IsBold = True
        Dim fontFlag As New StyleFlag()
        fontFlag.FontBold = True

        reportsheet.Cells.Rows(rowIndex + 1).ApplyStyle(boldStyle, fontFlag)

        Dim boldAndCenterStyle As New Style()
        boldAndCenterStyle.Font.IsBold = True
        boldAndCenterStyle.HorizontalAlignment = TextAlignmentType.Center
        Dim boldAndCenterFlag As New StyleFlag()
        boldAndCenterFlag.FontBold = True
        boldAndCenterFlag.HorizontalAlignment = True


        reportsheet.Cells(nameCol & (rowIndex + 2)).Value = "Store Name"

        If _showWTD Then
            reportsheet.Cells(rowIndex, 2).Value = "Week/Month to Date Analysis"
            reportsheet.Cells(rowIndex, 2).SetStyle(boldAndCenterStyle, boldAndCenterFlag)

            reportsheet.Cells(wtdTYcol & (rowIndex + 2)).Value = "WTD TY Act"
            reportsheet.Cells(wtdLYcol & (rowIndex + 2)).Value = "WTD LY Act"
        Else
            reportsheet.Cells(rowIndex, 2).Value = "Month to Date Analysis"
            reportsheet.Cells(rowIndex, 2).SetStyle(boldAndCenterStyle, boldAndCenterFlag)
        End If

        reportsheet.Cells(dsaCol & (rowIndex + 2)).Value = "DSA"
        reportsheet.Cells(mtdTYcol & (rowIndex + 2)).Value = If(_showWTD, "MTD ", "") & "TY Act"
        reportsheet.Cells(mtdLYcol & (rowIndex + 2)).Value = If(_showWTD, "MTD ", "") & "LY Act"

        reportsheet.Cells(pctMTDcol & (rowIndex + 2)).Value = "% Chg"

        If includePlan Then
            reportsheet.Cells(mtdPlanCol & (rowIndex + 2)).Value = "Plan"
            reportsheet.Cells(pctMTDplanCol & (rowIndex + 2)).Value = "% To Plan"
            Dim headerRng As Range = reportsheet.Cells.CreateRange(rowIndex + 1, 0, 1, 7 + wtdAdjust)
            headerRng.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Double, Color.Black)
        Else
            Dim headerRng As Range = reportsheet.Cells.CreateRange(rowIndex + 1, 0, 1, 5 + wtdAdjust)
            headerRng.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Double, Color.Black)
        End If

        reportsheet.Cells(rowIndex, 8 + wtdAdjust).Value = "Season to Date Analysis"
        reportsheet.Cells(rowIndex, 8 + wtdAdjust).SetStyle(boldAndCenterStyle, boldAndCenterFlag)

        reportsheet.Cells(stdTYcol & (rowIndex + 2)).Value = "TY Act"
        reportsheet.Cells(stdLYcol & (rowIndex + 2)).Value = "LY Act"

        reportsheet.Cells(pctSTDcol & (rowIndex + 2)).Value = "% Chg"

        If includePlan Then
            reportsheet.Cells(stdPlanCol & (rowIndex + 2)).Value = "Plan"
            reportsheet.Cells(pctSTDplanCol & (rowIndex + 2)).Value = "% To Plan"
            Dim headerRng As Range = reportsheet.Cells.CreateRange(rowIndex + 1, 8 + wtdAdjust, 1, 5)
            headerRng.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Double, Color.Black)
        Else
            Dim headerRng As Range = reportsheet.Cells.CreateRange(rowIndex + 1, 8 + wtdAdjust, 1, 3)
            headerRng.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Double, Color.Black)
        End If

        Return rowIndex + 3
    End Function

    Private Function WriteDataRows(ByVal departmentCode As String, ByVal rgn As String, ByVal materials As String(), ByVal direct As Integer, ByVal includePlan As Boolean, ByVal currentRow As Integer, ByVal reportSheet As Worksheet) As Integer
        Dim rowsToConsider As DataTable

        Try
            If direct = 1 Then
                rowsToConsider = _reportData.AsEnumerable().Where(Function(row) (row.Item("DEPT_CODE") = departmentCode Or departmentCode = "B") And row.Item("CUST_STORE_NAME") = "DIRECT" And materials.Contains(row.Item("MATL_CODE"))).CopyToDataTable()
            ElseIf rgn = "" And direct = -1 Then
                rowsToConsider = _reportData.AsEnumerable().Where(Function(row) (row.Item("DEPT_CODE") = departmentCode Or departmentCode = "B") And row.Item("CUST_STORE_NAME") <> "DIRECT" And materials.Contains(row.Item("MATL_CODE"))).CopyToDataTable()
            ElseIf rgn = "" And direct = 0 Then
                rowsToConsider = _reportData.AsEnumerable().Where(Function(row) (row.Item("DEPT_CODE") = departmentCode Or departmentCode = "B") And materials.Contains(row.Item("MATL_CODE"))).CopyToDataTable()
            Else
                rowsToConsider = _reportData.AsEnumerable().Where(Function(row) (row.Item("DEPT_CODE") = departmentCode Or departmentCode = "B") And row.Item("REGION_DESC") = rgn And row.Item("CUST_STORE_NAME") <> "DIRECT" And materials.Contains(row.Item("MATL_CODE"))).CopyToDataTable()
            End If
        Catch ex As Exception
            Return currentRow
        End Try

        Dim groupedByMaterial = rowsToConsider. _
                                AsEnumerable().GroupBy(Function(row) New With {Key .StoreName = row.Field(Of String)("CUST_STORE_NAME"), Key .DSA = row.Field(Of String)("DSA")}). _
                                [Select](Function(g) New With {.StoreName = g.Key.StoreName, _
                                                               .DSA = g.Key.DSA, _
                                                               .SoldThisYearWTD = g.Sum(Function(x) x.Field(Of Decimal?)("AMT_SOLD_WTD_TY")), _
                                                               .SoldLastYearWTD = g.Sum(Function(x) x.Field(Of Decimal?)("AMT_SOLD_WTD_LY")), _
                                                               .SoldThisYearMTD = g.Sum(Function(x) x.Field(Of Decimal?)("AMT_SOLD_MTD_TY")), _
                                                               .SoldLastYearMTD = g.Sum(Function(x) x.Field(Of Decimal?)("AMT_SOLD_MTD_LY")), _
                                                               .SoldThisYearSTD = g.Sum(Function(x) x.Field(Of Decimal?)("AMT_SOLD_STD_TY")), _
                                                               .SoldLastYearSTD = g.Sum(Function(x) x.Field(Of Decimal?)("AMT_SOLD_STD_LY")), _
                                                               .PlanMTD = g.Sum(Function(x) x.Field(Of Decimal?)("AMT_PLAN")), _
                                                               .PlanSTD = g.Sum(Function(x) x.Field(Of Decimal?)("AMT_PLAN_STD"))}).OrderBy(Function(x) If(x.StoreName = "OTHER" Or x.StoreName = "DIRECT", If(x.StoreName = "DIRECT", "ZZ", "ZZZ"), x.StoreName))

        For Each salesRow In groupedByMaterial
            reportSheet.Cells(nameCol & currentRow).Value = salesRow.StoreName
            reportSheet.Cells(dsaCol & currentRow).Value = salesRow.DSA
            If _showWTD Then
                reportSheet.Cells(wtdTYcol & currentRow).Value = salesRow.SoldThisYearWTD
                SetThousandsFormat(reportSheet.Cells(wtdTYcol & currentRow))
                reportSheet.Cells(wtdLYcol & currentRow).Value = salesRow.SoldLastYearWTD
                SetThousandsFormat(reportSheet.Cells(wtdLYcol & currentRow))
            End If
            reportSheet.Cells(mtdTYcol & currentRow).Value = salesRow.SoldThisYearMTD
            SetThousandsFormat(reportSheet.Cells(mtdTYcol & currentRow))
            reportSheet.Cells(mtdLYcol & currentRow).Value = salesRow.SoldLastYearMTD
            SetThousandsFormat(reportSheet.Cells(mtdLYcol & currentRow))
            reportSheet.Cells(pctMTDcol & currentRow).Formula = String.Format("=({1}{0}-{2}{0})/ABS({2}{0})", currentRow, mtdTYcol, mtdLYcol)
            SetPercentageFormat(reportSheet.Cells(pctMTDcol & currentRow))

            If includePlan Then
                reportSheet.Cells(mtdPlanCol & currentRow).Value = salesRow.PlanMTD
                SetThousandsFormat(reportSheet.Cells(mtdPlanCol & currentRow))
                reportSheet.Cells(pctMTDplanCol & currentRow).Formula = String.Format("={1}{0}/{2}{0}", currentRow, mtdTYcol, mtdPlanCol)
                SetPercentageFormat(reportSheet.Cells(pctMTDplanCol & currentRow))
            End If

            reportSheet.Cells(stdTYcol & currentRow).Value = salesRow.SoldThisYearSTD
            SetThousandsFormat(reportSheet.Cells(stdTYcol & currentRow))
            reportSheet.Cells(stdLYcol & currentRow).Value = salesRow.SoldLastYearSTD
            SetThousandsFormat(reportSheet.Cells(stdLYcol & currentRow))
            reportSheet.Cells(pctSTDcol & currentRow).Formula = String.Format("=({1}{0}-{2}{0})/ABS({2}{0})", currentRow, stdTYcol, stdLYcol)
            SetPercentageFormat(reportSheet.Cells(pctSTDcol & currentRow))

            If includePlan Then
                reportSheet.Cells(stdPlanCol & currentRow).Value = salesRow.PlanSTD
                SetThousandsFormat(reportSheet.Cells(stdPlanCol & currentRow))
                reportSheet.Cells(pctSTDplanCol & currentRow).Formula = String.Format("={1}{0}/{2}{0}", currentRow, stdTYcol, stdPlanCol)
                SetPercentageFormat(reportSheet.Cells(pctSTDplanCol & currentRow))
            End If

            currentRow += 1
        Next
        Return currentRow
    End Function

    Private Sub SetTotalsStyle(ByVal totalsRange As Range)
        Dim totalsStyle As Style = totalsRange.Worksheet.Workbook.Styles(totalsRange.Worksheet.Workbook.Styles.Add())
        totalsStyle.Font.IsBold = True
        totalsStyle.Pattern = BackgroundType.Solid
        totalsStyle.ForegroundColor = Color.Yellow
        totalsStyle.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black)
        totalsStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black)

        Dim totalsFlag As New StyleFlag()
        totalsFlag.FontBold = True
        totalsFlag.CellShading = True
        totalsFlag.Borders = True
        totalsFlag.TopBorder = True
        totalsFlag.BottomBorder = True
        totalsRange.ApplyStyle(totalsStyle, totalsFlag)
    End Sub

    Private Function WriteSubTotalRow(ByVal summaryText As String, ByVal rgnStartRow As Integer, ByVal includePlan As Boolean, ByVal currentRow As Integer, ByVal reportSheet As Worksheet) As Integer
        Dim rgnEndRow As Integer = currentRow - 1

        Dim cellsToStyle As Range = reportSheet.Cells.CreateRange(nameCol & currentRow, pctSTDplanCol & currentRow)
        SetTotalsStyle(cellsToStyle)

        reportSheet.Cells(nameCol & currentRow).Value = String.Format("Total {0}", summaryText)

        If (_showWTD) Then
            reportSheet.Cells(wtdTYcol & currentRow).Formula = String.Format("=SUM({2}{0}:{2}{1})", rgnStartRow, rgnEndRow, wtdTYcol)
            SetThousandsFormat(reportSheet.Cells(wtdTYcol & currentRow))
            reportSheet.Cells(wtdLYcol & currentRow).Formula = String.Format("=SUM({2}{0}:{2}{1})", rgnStartRow, rgnEndRow, wtdLYcol)
            SetThousandsFormat(reportSheet.Cells(wtdLYcol & currentRow))
        End If

        reportSheet.Cells(mtdTYcol & currentRow).Formula = String.Format("=SUM({2}{0}:{2}{1})", rgnStartRow, rgnEndRow, mtdTYcol)
        SetThousandsFormat(reportSheet.Cells(mtdTYcol & currentRow))
        reportSheet.Cells(mtdLYcol & currentRow).Formula = String.Format("=SUM({2}{0}:{2}{1})", rgnStartRow, rgnEndRow, mtdLYcol)
        SetThousandsFormat(reportSheet.Cells(mtdLYcol & currentRow))
        reportSheet.Cells(pctMTDcol & currentRow).Formula = String.Format("=({1}{0}-{2}{0})/ABS({2}{0})", currentRow, mtdTYcol, mtdLYcol)
        SetPercentageFormat(reportSheet.Cells(pctMTDcol & currentRow))

        If includePlan Then
            reportSheet.Cells(mtdPlanCol & currentRow).Formula = String.Format("=SUM({2}{0}:{2}{1})", rgnStartRow, rgnEndRow, mtdPlanCol)
            SetThousandsFormat(reportSheet.Cells(mtdPlanCol & currentRow))
            reportSheet.Cells(pctMTDplanCol & currentRow).Formula = String.Format("={1}{0}/{2}{0}", currentRow, mtdTYcol, mtdPlanCol)
            SetPercentageFormat(reportSheet.Cells(pctMTDplanCol & currentRow))
        End If

        reportSheet.Cells(stdTYcol & currentRow).Formula = String.Format("=SUM({2}{0}:{2}{1})", rgnStartRow, rgnEndRow, stdTYcol)
        SetThousandsFormat(reportSheet.Cells(stdTYcol & currentRow))
        reportSheet.Cells(stdLYcol & currentRow).Formula = String.Format("=SUM({2}{0}:{2}{1})", rgnStartRow, rgnEndRow, stdLYcol)
        SetThousandsFormat(reportSheet.Cells(stdLYcol & currentRow))
        reportSheet.Cells(pctSTDcol & currentRow).Formula = String.Format("=({1}{0}-{2}{0})/ABS({2}{0})", currentRow, stdTYcol, stdLYcol)
        SetPercentageFormat(reportSheet.Cells(pctSTDcol & currentRow))

        If includePlan Then
            reportSheet.Cells(stdPlanCol & currentRow).Formula = String.Format("=SUM({2}{0}:{2}{1})", rgnStartRow, rgnEndRow, stdPlanCol)
            SetThousandsFormat(reportSheet.Cells(stdPlanCol & currentRow))
            reportSheet.Cells(pctSTDplanCol & currentRow).Formula = String.Format("={1}{0}/{2}{0}", currentRow, stdTYcol, stdPlanCol)
            SetPercentageFormat(reportSheet.Cells(pctSTDplanCol & currentRow))
        End If

        Return currentRow + 1
    End Function

    Private Function WriteTotals(ByVal totalText As String, ByVal currentRow As Integer, ByVal summaryRows As List(Of Integer), ByVal reportSheet As Worksheet) As Integer

        Dim cellsToStyle As Range = reportSheet.Cells.CreateRange(nameCol & currentRow, pctSTDplanCol & currentRow)
        SetTotalsStyle(cellsToStyle)

        Dim thousandsStyle As Style = New Style()
        thousandsStyle.Custom = "$#0.0,"

        Dim percentageStyle As Style = New Style()
        percentageStyle.Custom = "#0.#%"

        Dim numberFormatFlag As New StyleFlag()
        numberFormatFlag.NumberFormat = True

        reportSheet.Cells("A" & currentRow).Value = "Total " & totalText

        For Each col In New String() {wtdTYcol, wtdLYcol, mtdTYcol, mtdLYcol, mtdPlanCol, stdTYcol, stdLYcol, stdPlanCol}
            Dim sumFormatString As String = ""
            Dim plus As String = "="
            For Each summaryRow In summaryRows
                sumFormatString &= plus & col & summaryRow
                plus = "+"
            Next
            sumFormatString &= ""
            reportSheet.Cells(col & currentRow).Formula = sumFormatString
            reportSheet.Cells(col & currentRow).SetStyle(thousandsStyle, numberFormatFlag)
        Next

        reportSheet.Cells(pctMTDcol & currentRow).Formula = String.Format("=({1}{0}-{2}{0})/ABS({2}{0})", currentRow, mtdTYcol, mtdLYcol)
        reportSheet.Cells(pctMTDcol & currentRow).SetStyle(percentageStyle, numberFormatFlag)
        reportSheet.Cells(pctMTDplanCol & currentRow).Formula = String.Format("={1}{0}/{2}{0}", currentRow, mtdTYcol, mtdPlanCol)
        reportSheet.Cells(pctMTDplanCol & currentRow).SetStyle(percentageStyle, numberFormatFlag)
        reportSheet.Cells(pctSTDcol & currentRow).Formula = String.Format("=({1}{0}-{2}{0})/ABS({2}{0})", currentRow, stdTYcol, stdLYcol)
        reportSheet.Cells(pctSTDcol & currentRow).SetStyle(percentageStyle, numberFormatFlag)
        reportSheet.Cells(pctSTDplanCol & currentRow).Formula = String.Format("={1}{0}/{2}{0}", currentRow, stdTYcol, stdPlanCol)
        reportSheet.Cells(pctSTDplanCol & currentRow).SetStyle(percentageStyle, numberFormatFlag)

        Return currentRow + 1
    End Function
End Class

Public Class NationalWeeklySalesFlash
    Inherits RetailExport

    Dim _reportYYYYPP As String
    Dim _reportData As DataTable
    Dim _eComSales As DataRow

    Public Sub New(ByVal reportYYYYPP As String)
        _reportYYYYPP = reportYYYYPP
    End Sub

    Protected Shared Sub SetAsposeLicense()
        Dim license As Aspose.Cells.License = New Aspose.Cells.License()
        license.SetLicense("Aspose.Total.lic")
    End Sub

    Public Sub CreateReport(ByVal fileName As String)
        GetReportData()
        SetAsposeLicense()
        Dim reportWorkbook As Workbook = New Workbook()
        reportWorkbook.Worksheets.RemoveAt(0)

        reportWorkbook.DefaultStyle.Font.Name = "Times New Roman"

        Dim reportSheet As Worksheet = reportWorkbook.Worksheets.Add("SheetName")

        reportSheet.PageSetup.FitToPagesWide = 1
        reportSheet.PageSetup.Orientation = PageOrientationType.Landscape
        reportSheet.PageSetup.FitToPagesTall = Nothing

        PrepareSheet(reportSheet)
        FillSheet(reportSheet)

        reportSheet.AutoFitColumn(0)
        Try
            reportWorkbook.Save(fileName, SaveFormat.Xlsx)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub FillSheet(ByVal reportSheet As Worksheet)
        Dim totalRows As New Dictionary(Of String, List(Of Integer))

        HideWeek5IfNecessary(reportSheet)
        Dim rowIndex As Integer = WriteSummarySection("", 1, "CUST_CODE", "", reportSheet, totalRows)
        reportSheet.Cells(rowIndex + 1, 0).Value = "*Please note that Holt Renfrew sales are in Canadian dollars"
        rowIndex = WriteSummarySection("Stores", rowIndex + 3, "DEPT_NAME", "AND DIRECT=0", reportSheet, totalRows)
        rowIndex = WriteSummarySection("Direct", rowIndex + 3, "DEPT_NAME", "AND DIRECT=1", reportSheet, totalRows)
        rowIndex += 2
        _eComSales.Item("CUST_NAME") = "JH E-Commerce"
        _eComSales.Item("DEPT_NAME") = ""
        WriteDataRow(_eComSales, rowIndex, reportSheet)
        SetEcommerceStyle(rowIndex, reportSheet)
        WriteTotals(rowIndex + 2, totalRows, reportSheet)
    End Sub

    Private Sub HideWeek5IfNecessary(reportSheet As Worksheet)
        Dim maxWeek As Integer = ASCDATA1.GetDataValue("SELECT MAX(REL_WEEK) FROM RSTCLND1 WHERE YYYYPP=:PARM1", "V", New Object() {_reportYYYYPP})
        If maxWeek < 5 Then
            reportSheet.Cells.HideColumns(13, 3)
        End If
    End Sub

    Private Function WriteSummarySection(ByVal summaryType As String, ByVal rowIndex As Integer, ByVal groupByColumn As String, ByVal filter As String, reportSheet As Worksheet, ByRef totalRows As Dictionary(Of String, List(Of Integer))) As Integer
        WriteWeekHeaders(rowIndex, reportSheet)
        WriteColumnHeaders(rowIndex + 1, reportSheet)

        rowIndex += 2
        Dim groupByItems = GetDistinct(groupByColumn, _reportData)
        If groupByColumn = "DEPT_NAME" Then
            groupByItems = groupByItems.OrderByDescending(Function(x) x).ToList()
        End If
        Dim subTotalsRows As New List(Of Integer)
        For Each groupByItem As String In groupByItems
            Dim startIndex As Integer = rowIndex
            Dim customerName As String = ""
            For Each row As DataRow In _reportData.Select(String.Format("{0}='{1}' " & filter, groupByColumn, groupByItem), "DIRECT, DEPT_NAME DESC")
                WriteDataRow(row, rowIndex, reportSheet)
                customerName = row.Item("CUST_NAME")
                rowIndex += 1
            Next
            WriteSubTotalRow(summaryType, If(groupByColumn = "CUST_CODE", customerName, groupByItem), startIndex, rowIndex, reportSheet)
            subTotalsRows.Add(rowIndex)
            If summaryType <> "" Then
                If Not totalRows.ContainsKey(groupByItem) Then
                    totalRows.Add(groupByItem, New List(Of Integer))
                End If
                totalRows(groupByItem).Add(rowIndex)
            End If
            rowIndex += 1
        Next

        rowIndex += 1

        WriteSummaryTotals(summaryType, rowIndex, subTotalsRows, reportSheet)

        Return rowIndex
    End Function

    Private Sub WriteDataRow(ByVal dataRow As DataRow, ByVal rowIndex As Integer, ByVal reportSheet As Worksheet)
        reportSheet.Cells(rowIndex, 0).Value = dataRow.Item("CUST_NAME") & " " & dataRow.Item("DEPT_NAME") & If(Val(dataRow.Item("DIRECT") & "") = 1, " Direct", "")

        For i As Integer = 0 To 4
            reportSheet.Cells(rowIndex, 1 + (i * 3)).Value = dataRow.Item(String.Format("WEEK_{0}_TY", i + 1))
            reportSheet.Cells(rowIndex, 2 + (i * 3)).Value = dataRow.Item(String.Format("WEEK_{0}_LY", i + 1))
            reportSheet.Cells(rowIndex, 3 + (i * 3)).Formula = String.Format("({0}{1}-{2}{1})/{2}{1}", Chr(Asc("A") + (1 + i * 3)), rowIndex + 1, Chr(Asc("A") + (2 + i * 3)))
        Next

        reportSheet.Cells(String.Format("Q{0}", rowIndex + 1)).Formula = String.Format("B{0}+E{0}+H{0}+K{0}+N{0}", rowIndex + 1)
        reportSheet.Cells(String.Format("R{0}", rowIndex + 1)).Formula = String.Format("C{0}+F{0}+I{0}+L{0}+O{0}", rowIndex + 1)
        reportSheet.Cells(String.Format("S{0}", rowIndex + 1)).Formula = String.Format("(Q{0}-R{0})/R{0}", rowIndex + 1)

        SetRowBorders(False, rowIndex, reportSheet)
    End Sub

    Private Sub WriteSubTotalRow(ByVal summaryType As String, ByVal description As String, ByVal startIndex As Integer, ByVal rowIndex As Integer, ByVal reportSheet As Worksheet)
        Dim subTotalStyle As Style = reportSheet.Workbook.Styles(reportSheet.Workbook.Styles.Add())
        subTotalStyle.Font.IsBold = True
        subTotalStyle.Pattern = BackgroundType.Solid
        subTotalStyle.ForegroundColor = Color.FromArgb(255, 255, 255, 153)
        subTotalStyle.SetBorder(BorderType.TopBorder, CellBorderType.Medium, Color.Black)
        subTotalStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black)
        Dim subTotalFlags As New StyleFlag()
        subTotalFlags.FontBold = True
        subTotalFlags.CellShading = True
        subTotalFlags.Borders = True

        reportSheet.Cells.CreateRange(rowIndex, 0, 1, 19).ApplyStyle(subTotalStyle, subTotalFlags)

        reportSheet.Cells(rowIndex, 0).Value = "Total " & description & " " & summaryType

        For i As Integer = 0 To 4
            Dim colTY As String = Chr(Asc("A") + (1 + i * 3))
            Dim colLY As String = Chr(Asc("A") + (2 + i * 3))
            reportSheet.Cells(rowIndex, 1 + (i * 3)).Formula = String.Format("SUM({0}{1}:{0}{2})", colTY, startIndex + 1, rowIndex)
            reportSheet.Cells(rowIndex, 2 + (i * 3)).Formula = String.Format("SUM({0}{1}:{0}{2})", colLY, startIndex + 1, rowIndex)
            reportSheet.Cells(rowIndex, 3 + (i * 3)).Formula = String.Format("({0}{1}-{2}{1})/{2}{1}", colTY, rowIndex + 1, colLY)
        Next

        reportSheet.Cells(String.Format("Q{0}", rowIndex + 1)).Formula = String.Format("B{0}+E{0}+H{0}+K{0}+N{0}", rowIndex + 1)
        reportSheet.Cells(String.Format("R{0}", rowIndex + 1)).Formula = String.Format("C{0}+F{0}+I{0}+L{0}+O{0}", rowIndex + 1)
        reportSheet.Cells(String.Format("S{0}", rowIndex + 1)).Formula = String.Format("(Q{0}-R{0})/R{0}", rowIndex + 1)

        SetRowBorders(True, rowIndex, reportSheet)
    End Sub

    Private Sub WriteSummaryTotals(ByVal summaryType As String, ByVal rowIndex As Integer, ByVal subTotalsRows As List(Of Integer), ByVal reportSheet As Worksheet)
        Dim subTotalStyle As Style = reportSheet.Workbook.Styles(reportSheet.Workbook.Styles.Add())
        subTotalStyle.Font.IsBold = True
        subTotalStyle.Pattern = BackgroundType.Solid
        subTotalStyle.ForegroundColor = Color.FromArgb(255, 204, 255, 204)
        subTotalStyle.SetBorder(BorderType.TopBorder, CellBorderType.Medium, Color.Black)
        subTotalStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black)
        Dim subTotalFlags As New StyleFlag()
        subTotalFlags.FontBold = True
        subTotalFlags.CellShading = True
        subTotalFlags.Borders = True

        reportSheet.Cells.CreateRange(rowIndex, 0, 1, 19).ApplyStyle(subTotalStyle, subTotalFlags)

        reportSheet.Cells(rowIndex, 0).Value = "Total " & summaryType

        For i As Integer = 0 To 5
            Dim colTY As String = Chr(Asc("A") + (1 + i * 3))
            Dim colLY As String = Chr(Asc("A") + (2 + i * 3))
            Dim formula1 As String = ""
            Dim formula2 As String = ""
            Dim delimiter As String = ""
            For Each subRowIndex As Integer In subTotalsRows
                formula1 &= delimiter & colTY & (subRowIndex + 1)
                formula2 &= delimiter & colLY & (subRowIndex + 1)
                delimiter = "+"
            Next

            reportSheet.Cells(rowIndex, 1 + (i * 3)).Formula = formula1
            reportSheet.Cells(rowIndex, 2 + (i * 3)).Formula = formula2
            reportSheet.Cells(rowIndex, 3 + (i * 3)).Formula = String.Format("({0}{1}-{2}{1})/{2}{1}", colTY, rowIndex + 1, colLY)
        Next

        SetRowBorders(True, rowIndex, reportSheet)
    End Sub

    Private Sub WriteTotals(ByVal rowIndex As Integer, ByVal totalRows As Dictionary(Of String, List(Of Integer)), ByVal reportSheet As Worksheet)
        For Each totalType As String In totalRows.Keys
            Dim totalsStyle As Style = reportSheet.Workbook.Styles(reportSheet.Workbook.Styles.Add())
            totalsStyle.Font.IsBold = True
            totalsStyle.Pattern = BackgroundType.Solid
            totalsStyle.ForegroundColor = Color.FromArgb(255, 255, 153, 204)
            totalsStyle.SetBorder(BorderType.TopBorder, CellBorderType.Medium, Color.Black)
            totalsStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black)
            Dim totalsFlags As New StyleFlag()
            totalsFlags.FontBold = True
            totalsFlags.CellShading = True
            totalsFlags.Borders = True

            reportSheet.Cells.CreateRange(rowIndex, 0, 1, 19).ApplyStyle(totalsStyle, totalsFlags)

            reportSheet.Cells(rowIndex, 0).Value = "Total " & totalType

            For i As Integer = 0 To 5
                Dim colTY As String = Chr(Asc("A") + (1 + i * 3))
                Dim colLY As String = Chr(Asc("A") + (2 + i * 3))
                Dim formula1 As String = ""
                Dim formula2 As String = ""
                Dim delimiter As String = ""
                For Each totalRowIndex As Integer In totalRows(totalType)
                    formula1 &= delimiter & colTY & (totalRowIndex + 1)
                    formula2 &= delimiter & colLY & (totalRowIndex + 1)
                    delimiter = "+"
                Next

                reportSheet.Cells(rowIndex, 1 + (i * 3)).Formula = formula1
                reportSheet.Cells(rowIndex, 2 + (i * 3)).Formula = formula2
                reportSheet.Cells(rowIndex, 3 + (i * 3)).Formula = String.Format("({0}{1}-{2}{1})/{2}{1}", colTY, rowIndex + 1, colLY)
            Next

            SetRowBorders(True, rowIndex, reportSheet)

            rowIndex += 2
        Next
    End Sub

    Private Sub SetEcommerceStyle(ByVal rowIndex As Integer, ByVal reportSheet As Worksheet)
        Dim totalsStyle As Style = reportSheet.Workbook.Styles(reportSheet.Workbook.Styles.Add())
        totalsStyle.Font.IsBold = True
        totalsStyle.Pattern = BackgroundType.Solid
        totalsStyle.ForegroundColor = Color.FromArgb(255, 255, 204, 0)
        totalsStyle.SetBorder(BorderType.TopBorder, CellBorderType.Medium, Color.Black)
        totalsStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black)
        Dim totalsFlags As New StyleFlag()
        totalsFlags.FontBold = True
        totalsFlags.CellShading = True
        totalsFlags.Borders = True

        reportSheet.Cells.CreateRange(rowIndex, 0, 1, 19).ApplyStyle(totalsStyle, totalsFlags)

        'reportSheet.Cells(rowIndex, 0).Value = "Total "

        SetRowBorders(True, rowIndex, reportSheet)
    End Sub

    Private Sub WriteWeekHeaders(ByVal rowIndex As Integer, ByVal reportSheet As Worksheet)
        Dim wkRangeStyle As Style = reportSheet.Workbook.Styles(reportSheet.Workbook.Styles.Add())
        wkRangeStyle.Font.IsBold = True
        Dim wkRangeFlags As New StyleFlag()
        wkRangeFlags.FontBold = True
        For i As Integer = 0 To 4
            Dim curIndex = i
            Dim wkRange As Range = reportSheet.Cells.CreateRange(rowIndex, 1 + (3 * i), 1, 2)
            wkRange.Merge()
            wkRange.Value = _reportData.AsEnumerable.Where(Function(row) row.Item(String.Format("WEEK_{0}_RNG", curIndex + 1)) & "" <> "").Select(Function(row) row.Item(String.Format("WEEK_{0}_RNG", curIndex + 1)).ToString()).FirstOrDefault()
            If wkRange.Value(0, 0) Is Nothing Then
                Dim firstWeekRange As String = _reportData.AsEnumerable.Where(Function(row) row.Item("WEEK_1_RNG") & "" <> "").Select(Function(row) row.Item("WEEK_1_RNG").ToString()).FirstOrDefault()
                Dim firstDay As String = firstWeekRange.Split("-")(0)
                Dim firstDate = Convert.ToDateTime(firstDay, New System.Globalization.CultureInfo("en-US"))
                wkRange.Value = firstDate.AddDays(7 * i).ToString("M/d") & "-" & firstDate.AddDays(7 * i + 6).ToString("M/d")
            End If
            wkRange.SetOutlineBorders(CellBorderType.Medium, Color.Black)
            wkRange.ApplyStyle(wkRangeStyle, wkRangeFlags)
        Next
    End Sub

    Private Sub WriteColumnHeaders(ByVal rowIndex As Integer, ByVal reportSheet As Worksheet)
        Dim headerCellStyle As New Style()
        headerCellStyle.Font.IsBold = True
        headerCellStyle.SetBorder(BorderType.TopBorder, CellBorderType.Medium, Color.Black)
        headerCellStyle.SetBorder(BorderType.LeftBorder, CellBorderType.Medium, Color.Black)
        headerCellStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black)
        headerCellStyle.SetBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black)
        Dim headerStyleFlags As New StyleFlag()
        headerStyleFlags.FontBold = True
        headerStyleFlags.Borders = True

        Dim columnHeaders = New String() {"Store Name", "TY WK 1", "LY WK 1", "% Chg", "TY WK 2", "LY WK 2", "% Chg", "TY WK 3", "LY WK 3", "% Chg", "TY WK 4", "LY WK 4", "% Chg", "TY WK 5", "LY WK 5", "% Chg", "MTD TY", "MTD LY", "% Chg"}


        For i As Integer = 0 To 18
            reportSheet.Cells(rowIndex, i).SetStyle(headerCellStyle, headerStyleFlags)
            reportSheet.Cells(rowIndex, i).Value = columnHeaders(i)
        Next

    End Sub


    Private Sub SetRowBorders(ByVal boldBottom As Boolean, ByVal rowIndex As Integer, reportSheet As Worksheet)
        Dim borderFlag As New StyleFlag()
        borderFlag.BottomBorder = True
        borderFlag.RightBorder = True
        borderFlag.LeftBorder = True

        For i As Integer = 0 To 18
            Dim cellStyle = reportSheet.Cells(rowIndex, i).GetStyle()
            If boldBottom Then
                cellStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black)
            Else
                cellStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black)
            End If

            If i = 0 Then
                cellStyle.SetBorder(BorderType.LeftBorder, CellBorderType.Medium, Color.Black) 'Set left border on first column
            End If

            cellStyle.SetBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black)

            reportSheet.Cells(rowIndex, i).SetStyle(cellStyle)
        Next
    End Sub


    Private Sub PrepareSheet(ByVal reportSheet As Worksheet)

        Dim numberFormatFlag As New StyleFlag()
        numberFormatFlag.NumberFormat = True

        Dim thousandsNumberFormat As New Style()
        thousandsNumberFormat.Custom = "$#0.0,"

        Dim percentageFormat As New Style()
        percentageFormat.Custom = "#0.0%"

        Dim centerAlignmentStyle As New Style()
        centerAlignmentStyle.HorizontalAlignment = TextAlignmentType.Center
        Dim alignmentFlags As New StyleFlag()
        alignmentFlags.HorizontalAlignment = True

        Dim fontStyle As New Style()
        fontStyle.Font.IsBold = True
        fontStyle.Font.Size = 20
        Dim fontFlag As New StyleFlag()
        fontFlag.FontSize = True
        fontFlag.FontBold = True

        reportSheet.FreezePanes(2, 1, 2, 1)

        Dim topRowRange = reportSheet.Cells.CreateRange("A1", "S1")
        topRowRange.Merge()
        topRowRange.ApplyStyle(centerAlignmentStyle, alignmentFlags)
        topRowRange.ApplyStyle(fontStyle, fontFlag)

        Dim titleString As String = ASCDATA1.GetDataValue("SELECT 'MTD ' || TO_CHAR(TO_DATE(OPS_YYYYMM,'YYYYMM'),'Month YYYY') FROM GLTPARM2 WHERE OPS_YYYYPP=:PARM1", "V", New Object() {_reportYYYYPP})
        topRowRange.Value = titleString

        For Each i As Integer In New Integer() {1, 2, 4, 5, 7, 8, 10, 11, 13, 14, 16, 17}
            reportSheet.Cells.Columns(i).ApplyStyle(thousandsNumberFormat, numberFormatFlag)
        Next

        For Each i As Integer In New Integer() {3, 6, 9, 12, 15, 18}
            reportSheet.Cells.Columns(i).ApplyStyle(percentageFormat, numberFormatFlag)
        Next

        For i As Integer = 1 To 18
            reportSheet.Cells.Columns(i).ApplyStyle(centerAlignmentStyle, alignmentFlags)
        Next

        reportSheet.Cells.Columns(0).Width = 22
        reportSheet.Cells.Rows(0).Height *= 2
        reportSheet.Cells.Rows(1).Height *= 2
        reportSheet.Cells.Rows(2).Height *= 2

        reportSheet.Cells("A2").Value = String.Format("As of {0}", Date.Today.ToString("MM/dd/yyyy"))
        SetStyleFontBold(reportSheet.Cells("A2"))
        SetStyleFontColor(reportSheet.Cells("A2"), Color.Red)

    End Sub


    Private Sub GetReportData()
        _reportData = ASCDATA1.GetDataTable("SELECT * FROM RSTNATF1 WHERE OPS_YYYYPP=:PARM1 AND CUST_CODE <> 'ECOMSALE10'", "", "V", New Object() {_reportYYYYPP})
        _eComSales = ASCDATA1.GetDataRow("SELECT * FROM RSTNATF1 WHERE OPS_YYYYPP=:PARM1 AND CUST_CODE = 'ECOMSALE10'", True, "V", New Object() {_reportYYYYPP})
    End Sub

End Class

Public Class MonthlyNationalAndIndependentRecap
    Inherits RetailExport

    Dim _reportYYYYPP As String
    Dim _reportData As DataTable
    Dim _dateData As DataRow

    Public Sub New(ByVal reportYYYYPP As String)
        _reportYYYYPP = reportYYYYPP
    End Sub

    Protected Shared Sub SetAsposeLicense()
        Dim license As License = New License()
        license.SetLicense("Aspose.Total.lic")
    End Sub

    Public Sub CreateReport(ByVal fileName As String)
        GetReportData()
        SetAsposeLicense()
        Dim reportWorkbook As Workbook = New Workbook()
        reportWorkbook.Worksheets.RemoveAt(0)

        reportWorkbook.DefaultStyle.Font.Name = "Times New Roman"

        Dim reportSheet As Worksheet = reportWorkbook.Worksheets.Add("SheetName")

        reportSheet.PageSetup.FitToPagesWide = 1
        reportSheet.PageSetup.FitToPagesTall = Nothing

        reportSheet.PageSetup.Orientation = PageOrientationType.Landscape
        reportSheet.PageSetup.PaperSize = PaperSizeType.PaperLetter

        PrepareSheet(reportSheet)
        FillSheet(reportSheet)

        reportSheet.AutoFitColumn(0)
        Try
            reportWorkbook.Save(fileName, SaveFormat.Xlsx)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub PrepareSheet(ByVal reportSheet As Worksheet)
        Dim numberFormatFlag As New StyleFlag()
        numberFormatFlag.NumberFormat = True

        Dim thousandsNumberFormat As New Style()
        thousandsNumberFormat.Custom = "$#,##0,"

        Dim percentageFormat As New Style()
        percentageFormat.Custom = "#0.0%"

        Dim centerAlignmentStyle As New Style()
        centerAlignmentStyle.HorizontalAlignment = TextAlignmentType.Center
        Dim alignmentFlags As New StyleFlag()
        alignmentFlags.HorizontalAlignment = True

        Dim fontStyle As New Style()
        fontStyle.Font.IsBold = True
        Dim fontFlag As New StyleFlag()
        fontFlag.FontBold = True

        reportSheet.Cells.Columns(0).ApplyStyle(fontStyle, fontFlag)

        For Each i As Integer In New Integer() {2, 3, 5, 6, 8, 9, 11, 12, 13, 15, 16, 18}
            reportSheet.Cells.Columns(i).ApplyStyle(thousandsNumberFormat, numberFormatFlag)
        Next

        For Each i As Integer In New Integer() {4, 7, 10, 14, 17, 19}
            reportSheet.Cells.Columns(i).ApplyStyle(percentageFormat, numberFormatFlag)
        Next

        For i As Integer = 0 To 19
            reportSheet.Cells.Columns(i).ApplyStyle(centerAlignmentStyle, alignmentFlags)
        Next

        reportSheet.Cells.Columns(0).Width = 25

        reportSheet.Zoom = 80

        reportSheet.PageSetup.PrintTitleRows = "$1:$2"
        Dim dateRange As String = _dateData.Item("DATE_RANGE_FY").ToString()
        dateRange = dateRange.Substring(dateRange.Length - 2)
        reportSheet.PageSetup.SetHeader(1, String.Format("John Hardy" & vbCrLf &
                                                        "Retail Recap FY20{0}" & vbCrLf &
                                                        "{1} 20{2}", _dateData.Item("FYY"), _dateData.Item("MM"), dateRange))
        reportSheet.PageSetup.SetFooter(0, Today.ToString("MM/dd/yyyy"))
        reportSheet.PageSetup.SetFooter(1, String.Format("John Hardy {1} 20{0}", _dateData.Item("FYY"), _dateData.Item("MM")))
    End Sub

    Private Sub FillSheet(ByVal reportSheet As Worksheet)
        Dim rowIndex As Integer = 0
        rowIndex = WriteHeaderRow(rowIndex, reportSheet)

        Dim totalRows As New List(Of Integer)
        Dim totalWomenRows As New List(Of Integer)
        Dim totalMenRows As New List(Of Integer)
        Dim totalDirectRows As New List(Of Integer)
        SetRowBorders(False, rowIndex + 1, reportSheet)
        rowIndex = WriteNationalGrouping(rowIndex + 2, "Women Stores", "DEPT_NAME='Women' AND INDEPENDENT='0' AND DIRECT='0'", reportSheet, groupDepartment:=True)
        totalRows.Add(rowIndex)
        totalWomenRows.Add(rowIndex)
        SetRowBorders(False, rowIndex + 1, reportSheet)

        rowIndex = WriteNationalGrouping(rowIndex + 2, "Men Stores", "DEPT_NAME='Men' AND INDEPENDENT='0' AND DIRECT='0'", reportSheet, groupDepartment:=True)
        totalRows.Add(rowIndex)
        totalMenRows.Add(rowIndex)
        SetRowBorders(False, rowIndex + 1, reportSheet)

        rowIndex = WriteNationalGrouping(rowIndex + 2, "Stores (Women + Men)", "INDEPENDENT='0' AND DIRECT='0'", reportSheet, groupDepartment:=False)
        SetRowBorders(False, rowIndex + 1, reportSheet)

        rowIndex = WriteNationalGrouping(rowIndex + 2, "Women/Men Direct", "INDEPENDENT='0' AND DIRECT='1'", reportSheet, groupDepartment:=True)
        totalRows.Add(rowIndex)
        totalDirectRows.Add(rowIndex)
        SetRowBorders(False, rowIndex + 1, reportSheet)

        rowIndex = WriteTotals(rowIndex + 2, "Total Nationals", totalRows, reportSheet, Color.FromArgb(255, 204, 255, 204))
        totalRows.Clear()
        totalRows.Add(rowIndex)
        SetRowBorders(False, rowIndex + 1, reportSheet)

        rowIndex = WriteIndependents(rowIndex + 2, reportSheet)
        totalRows.Add(rowIndex)

        rowIndex = WriteTotals(rowIndex + 2, "Grand Total", totalRows, reportSheet, Color.FromArgb(255, 204, 255, 204)) 'national total + independent total

        rowIndex = WriteEcommerce(rowIndex + 2, reportSheet)
        totalDirectRows.Add(rowIndex)

        rowIndex = WriteLastCallorSaksOff5th("NM Last Call", rowIndex + 2, reportSheet)
        rowIndex = WriteLastCallorSaksOff5th("Saks Off 5th", rowIndex + 1, reportSheet)

        Dim nationalsTable As DataTable = _reportData.AsEnumerable.Where(Function(row) row.Item("INDEPENDENT") = 0).CopyToDataTable()
        Dim customerView As New DataView(nationalsTable)
        Dim custsTable As DataTable = customerView.ToTable(True, New String() {"CUST_CODE", "CUST_NAME"})

        For Each custRow As DataRow In custsTable.Rows
            rowIndex = WriteHeaderRow(rowIndex + 2, reportSheet)
            rowIndex = WriteNationalGrouping(rowIndex + 1, custRow.Item("CUST_NAME"), String.Format("CUST_CODE='{0}' AND INDEPENDENT='0'", custRow.Item("CUST_CODE")), reportSheet, groupDepartment:=True)
            totalRows.Add(rowIndex)
        Next

        rowIndex = WriteTotals(rowIndex + 2, "Total Women Stores", totalWomenRows, reportSheet, Color.FromArgb(255, 255, 153, 204))
        rowIndex = WriteTotals(rowIndex + 2, "Total Men's Stores", totalMenRows, reportSheet, Color.FromArgb(255, 255, 153, 204))
        rowIndex = WriteEcommerce(rowIndex + 2, reportSheet)
        rowIndex = WriteTotals(rowIndex + 2, "Total Direct", totalDirectRows, reportSheet, Color.FromArgb(255, 255, 255, 153))

        reportSheet.Cells.Columns(17).IsHidden = True
        reportSheet.Cells.Columns(19).IsHidden = True
        reportSheet.Cells.DeleteColumn(11)
    End Sub

    Private Function WriteLastCallorSaksOff5th(ByVal custName As String, ByVal rowIndex As Integer, ByVal reportSheet As Worksheet)
        Dim totalRng As Range = reportSheet.Cells.CreateRange(rowIndex, 0, 1, 20)
        SetStyleForegroundColor(totalRng, Color.FromArgb(255, 218, 238, 243))
        SetStyleFontBold(totalRng)
        totalRng.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Medium, Color.Black)
        totalRng.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black)

        reportSheet.Cells(rowIndex, 0).Value = custName
        SetRowBorders(True, rowIndex, reportSheet)


        Dim colHeaders = New String() {"FYTD", "RYTD", "RMTD", "F3M", "FYB"}
        custName = If(custName.ToUpper().Contains("LAST"), "NMLASTCALL", "SAKSOFF5TH")

        Dim dr As DataRow = _reportData.AsEnumerable.Where(Function(row) row.Item("CUST_CODE") = custName).FirstOrDefault()

        If dr IsNot Nothing Then
            Dim offset As Integer = 0
            Dim offset2 As Integer = 0
            For i As Integer = 0 To 4
                reportSheet.Cells(rowIndex, 2 + (i * 3) + offset).Value = dr.Item(colHeaders(i) & "_TY")
                If i = 4 Then
                    offset2 += 1
                    reportSheet.Cells(rowIndex, 3 + (i * 3) + offset).Formula = String.Format("({0}{1}-{2}{1})/{2}{1}", Chr(Asc("A") + (1 + i * 3) + offset), rowIndex + 1, Chr(Asc("A") + (2 + i * 3) + offset))
                End If
                reportSheet.Cells(rowIndex, 3 + (i * 3) + offset + offset2).Value = dr.Item(colHeaders(i) & "_LY")
                reportSheet.Cells(rowIndex, 4 + (i * 3) + offset + offset2).Formula = String.Format("({0}{1}-{2}{1})/{2}{1}", Chr(Asc("A") + (2 + i * 3) + offset), rowIndex + 1, Chr(Asc("A") + (3 + i * 3) + offset + offset2))
                If i = 2 Then
                    offset += 1
                    reportSheet.Cells(rowIndex, 4 + (i * 3) + offset).Value = dr.Item("TM2YB")
                End If

                If i = 3 Then
                    offset += 1
                    reportSheet.Cells(rowIndex, 4 + (i * 3) + offset).Value = dr.Item("AMT_PLAN")
                End If
            Next
        End If

        Return rowIndex
    End Function

    Private Function WriteEcommerce(ByVal rowIndex As Integer, ByVal reportSheet As Worksheet) As Integer
        Dim totalRng As Range = reportSheet.Cells.CreateRange(rowIndex, 0, 1, 20)
        SetStyleForegroundColor(totalRng, Color.FromArgb(255, 252, 213, 180))
        SetStyleFontBold(totalRng)
        totalRng.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Medium, Color.Black)
        totalRng.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black)

        reportSheet.Cells(rowIndex, 0).Value = "JH E Commerce"
        SetRowBorders(True, rowIndex, reportSheet)

        Dim colHeaders = New String() {"FYTD", "RYTD", "RMTD", "F3M", "FYB"}


        Dim dr As DataRow = _reportData.AsEnumerable.Where(Function(row) row.Item("CUST_CODE") = "ECOMSALE10").FirstOrDefault()


        If dr IsNot Nothing Then
            Dim offset As Integer = 0
            Dim offset2 As Integer = 0
            For i As Integer = 0 To 4
                reportSheet.Cells(rowIndex, 2 + (i * 3) + offset).Value = dr.Item(colHeaders(i) & "_TY")
                If i = 4 Then
                    offset2 += 1
                    reportSheet.Cells(rowIndex, 3 + (i * 3) + offset).Formula = String.Format("({0}{1}-{2}{1})/{2}{1}", Chr(Asc("A") + (1 + i * 3) + offset), rowIndex + 1, Chr(Asc("A") + (2 + i * 3) + offset))
                End If
                reportSheet.Cells(rowIndex, 3 + (i * 3) + offset + offset2).Value = dr.Item(colHeaders(i) & "_LY")
                reportSheet.Cells(rowIndex, 4 + (i * 3) + offset + offset2).Formula = String.Format("({0}{1}-{2}{1})/{2}{1}", Chr(Asc("A") + (2 + i * 3) + offset), rowIndex + 1, Chr(Asc("A") + (3 + i * 3) + offset + offset2))
                If i = 2 Then
                    offset += 1
                    reportSheet.Cells(rowIndex, 4 + (i * 3) + offset).Value = dr.Item("TM2YB")
                End If
                If i = 3 Then
                    offset += 1
                    reportSheet.Cells(rowIndex, 4 + (i * 3) + offset).Value = dr.Item("AMT_PLAN")
                End If
            Next
        End If

        Return rowIndex
    End Function


    Private Function WriteNationalGrouping(ByVal rowIndex As Integer, ByVal totalText As String, ByVal filter As String, ByVal reportSheet As Worksheet, ByVal groupDepartment As Boolean) As Integer

        Dim startIndex As Integer = rowIndex

        If Not groupDepartment Then

            Dim dtr As DataTable = _reportData.Clone() 'create a clone of the report data structure, which will be filled by the below LINQ query (by way of the LoadDataRow command in the Select clause)

            Dim dt = From row In _reportData.Select(filter, "SEQ, DIRECT, DEPT_NAME DESC").AsEnumerable()
                     Group row By YYYYPP = row.Item("YYYYPP"), CUST_CODE = row.Item("CUST_CODE"), CUST_NAME = row.Item("CUST_NAME"),
                                  INDEPENDENT = row.Item("INDEPENDENT"), DIRECT = row.Item("DIRECT"), DOOR_W = row.Item("DOOR_W"), DOOR_M = row.Item("DOOR_M") Into FYTD_TY = Sum(Convert.ToDecimal(row.Item("FYTD_TY"))),
                                                                         FYTD_LY = Sum(Convert.ToDecimal(row.Item("FYTD_LY"))),
                                                                         RYTD_TY = Sum(Convert.ToDecimal(row.Item("RYTD_TY"))),
                                                                         RYTD_LY = Sum(Convert.ToDecimal(row.Item("RYTD_LY"))),
                                                                         RMTD_TY = Sum(Convert.ToDecimal(row.Item("RMTD_TY"))),
                                                                         RMTD_LY = Sum(Convert.ToDecimal(row.Item("RMTD_LY"))),
                                                                         F3M_TY = Sum(Convert.ToDecimal(row.Item("F3M_TY"))),
                                                                         F3M_LY = Sum(Convert.ToDecimal(row.Item("F3M_LY"))),
                                                                         TM2YB = Sum(Convert.ToDecimal(row.Item("TM2YB"))),
                                                                         FYB_TY = Sum(Convert.ToDecimal(row.Item("FYB_TY"))),
                                                                         FYB_LY = Sum(Convert.ToDecimal(row.Item("FYB_LY"))),
                                                                         AMT_PLAN = Sum(Convert.ToDecimal(Val(row.Item("AMT_PLAN") & ""))),
                                                                         SEQ = Min(Convert.ToInt32(row.Item("SEQ")))
                     Select dtr.LoadDataRow(New Object() {YYYYPP, CUST_CODE, CUST_NAME, "WM", "Women", SEQ, INDEPENDENT, DIRECT, FYTD_TY, FYTD_LY, RYTD_TY, RYTD_LY, RMTD_TY, RMTD_LY, F3M_TY, F3M_LY, TM2YB, FYB_TY, FYB_LY, DOOR_W + DOOR_M, DOOR_M, AMT_PLAN}, True)

            dt.ToArray() 'this line forces the LINQ query to evaluate, populating the DataTable dtr

            For Each row In dtr.Rows
                WriteDataRow(row, groupDepartment, rowIndex, reportSheet)
                rowIndex += 1
            Next
        Else
            For Each row As DataRow In _reportData.Select(filter, "SEQ, DIRECT, DEPT_NAME DESC")
                WriteDataRow(row, groupDepartment, rowIndex, reportSheet)
                rowIndex += 1
            Next
        End If

        WriteSubTotalRow(startIndex, rowIndex, totalText, reportSheet)

        Return rowIndex
    End Function

    Private Function WriteIndependents(ByVal rowIndex As Integer, ByVal reportSheet As Worksheet) As Integer
        Dim startIndex As Integer = rowIndex
        For Each row As DataRow In _reportData.Select("INDEPENDENT='1' AND CUST_CODE NOT IN ('NMLASTCALL','SAKSOFF5TH','ECOMSALE10')", "FYB_TY DESC")
            WriteDataRow(row, False, rowIndex, reportSheet)
            SetRowBorders(False, rowIndex + 1, reportSheet)
            rowIndex += 2
        Next

        WriteSubTotalRow(startIndex, rowIndex, "Independent", reportSheet)

        Return rowIndex
    End Function

    Private Sub WriteDataRow(ByVal dataRow As DataRow, ByVal writeDepartment As Boolean, ByVal rowIndex As Integer, ByVal reportSheet As Worksheet)

        If dataRow.Item("CUST_CODE") = "BCCLARK10" Then dataRow.Item("CUST_NAME") = "BC Clark"

        reportSheet.Cells(rowIndex, 0).Value = dataRow.Item("CUST_NAME") &
                    If(writeDepartment, If(dataRow.Item("DIRECT") = 1, " Direct", "") & " " & dataRow.Item("DEPT_NAME"), "")

        reportSheet.Cells(rowIndex, 1).Value = If(dataRow.Item("DEPT_NAME") = "Women", dataRow.Item("DOOR_W"), dataRow.Item("DOOR_M"))

        Dim colHeaders = New String() {"FYTD", "RYTD", "RMTD", "F3M", "FYB"}

        Dim offset As Integer = 0
        Dim offset2 As Integer = 0
        For i As Integer = 0 To 4
            reportSheet.Cells(rowIndex, 2 + (i * 3) + offset).Value = dataRow.Item(colHeaders(i) & "_TY")
            If i = 4 Then
                reportSheet.Cells(rowIndex, 3 + (i * 3) + offset).Formula = String.Format("({0}{1}-{2}{1})/{2}{1}", Chr(Asc("A") + (1 + i * 3) + offset), rowIndex + 1, Chr(Asc("A") + (2 + i * 3) + offset))
                offset2 = 1
            End If
            reportSheet.Cells(rowIndex, 3 + (i * 3) + offset + offset2).Value = dataRow.Item(colHeaders(i) & "_LY")
            reportSheet.Cells(rowIndex, 4 + (i * 3) + offset + offset2).Formula = String.Format("({0}{1}-{2}{1})/{2}{1}", Chr(Asc("A") + (2 + i * 3) + offset), rowIndex + 1, Chr(Asc("A") + (3 + i * 3) + offset + offset2))
            If i = 2 Then
                offset = 1
                reportSheet.Cells(rowIndex, 4 + (i * 3) + offset).Value = dataRow.Item("TM2YB")
            End If
            If i = 3 Then
                offset += 1
                reportSheet.Cells(rowIndex, 4 + (i * 3) + offset).Value = dataRow.Item("AMT_PLAN")
            End If
        Next

        SetRowBorders(False, rowIndex, reportSheet)
    End Sub

    Private Sub SetRowBorders(ByVal totalRow As Boolean, ByVal rowIndex As Integer, reportSheet As Worksheet)
        Dim mediumBorderCols = New Integer() {0, 1, 4, 7, 10, 11, 14, 18}
        For i As Integer = 0 To 19
            Dim cellStyle = reportSheet.Cells(rowIndex, i).GetStyle()
            If i = 0 Then
                cellStyle.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black)
            End If

            If Not totalRow Then
                cellStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black)
            End If
            If mediumBorderCols.Contains(i) Or totalRow Then
                cellStyle.SetBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black)
            Else
                cellStyle.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black)
            End If

            reportSheet.Cells(rowIndex, i).SetStyle(cellStyle)
        Next
    End Sub

    Private Function WriteHeaderRow(ByVal rowIndex As Integer, ByVal reportSheet As Worksheet) As Integer
        Dim accountRng As Range = reportSheet.Cells.CreateRange(rowIndex, 0, 2, 1)
        accountRng.Merge()
        SetStyleForegroundColor(accountRng, Color.FromArgb(255, 153, 204, 255))
        SetStyleFontBold(accountRng)
        SetStyleAlignCenter(accountRng)
        SetStyleAlignCenterVertical(accountRng)
        accountRng.SetOutlineBorders(CellBorderType.Medium, Color.Black)
        accountRng.Value = "Account"

        Dim doorRng As Range = reportSheet.Cells.CreateRange(rowIndex, 1, 2, 1)
        doorRng.Merge()
        SetStyleForegroundColor(doorRng, Color.FromArgb(255, 153, 204, 255))
        SetStyleFontBold(doorRng)
        SetStyleAlignCenter(doorRng)
        SetStyleAlignCenterVertical(doorRng)
        SetStyleWrapText(doorRng)
        doorRng.SetOutlineBorders(CellBorderType.Medium, Color.Black)
        doorRng.Value = "Door Count"

        Dim colHeaders = New String() {String.Format("JH FY{0} YTD ({1})", _dateData.Item("FYY"), _dateData.Item("DATE_RANGE_FY")),
                                       String.Format("RY{0} ({1})", _dateData.Item("RYY"), _dateData.Item("DATE_RANGE_RY")),
                                       String.Format("{0}", _dateData.Item("MM")),
                                       String.Format("3 Month Trend ({0})", _dateData.Item("DATE_RANGE_3MB")),
                                       ""}

        Dim offset As Integer = 0
        Dim offset2 As Integer = 0
        For i As Integer = 0 To 4
            Dim TyLyRng As Range = reportSheet.Cells.CreateRange(rowIndex, 2 + (i * 3) + offset, 2, 3)
            If i = 4 Then
                TyLyRng = reportSheet.Cells.CreateRange(rowIndex, 1 + (i * 3) + offset, 2, 5)
            End If
            SetStyleForegroundColor(TyLyRng, Color.FromArgb(255, 204, 255, 204))
            SetStyleFontBold(TyLyRng)
            SetStyleAlignCenter(TyLyRng)
            TyLyRng.SetOutlineBorders(CellBorderType.Medium, Color.Black)

            Dim headerTextRng As Range = reportSheet.Cells.CreateRange(rowIndex, 2 + (i * 3) + offset, 1, 3)
            headerTextRng.Merge()
            headerTextRng.Value = colHeaders(i)

            reportSheet.Cells(rowIndex + 1, 2 + (i * 3) + offset).Value = If(i = 4, String.Format("FY{0}" & vbCrLf & "Actual", _dateData.Item("FYY") - 1), "TY")
            SetStyleWrapText(reportSheet.Cells(rowIndex + 1, 2 + (i * 3) + offset))
            If i = 4 Then
                reportSheet.Cells(rowIndex + 1, 3 + (i * 3) + offset).Value = "% CHG" & If(i = 4, String.Format(" FY{0} Plan vs FY{1} Act", _dateData.Item("FYY"), _dateData.Item("FYY") - 1), "")
                SetStyleWrapText(reportSheet.Cells(rowIndex + 1, 3 + (i * 3) + offset))
                offset2 = 1
            End If
            reportSheet.Cells(rowIndex + 1, 3 + (i * 3) + offset + offset2).Value = If(i = 4, String.Format("FY{0}", _dateData.Item("FYY") - 2), "LY")
            SetStyleWrapText(reportSheet.Cells(rowIndex + 1, 3 + (i * 3) + offset + offset2))
            reportSheet.Cells(rowIndex + 1, 4 + (i * 3) + offset + offset2).Value = "% CHG" & If(i = 4, String.Format(" FY{0} Act vs FY{1} Act", _dateData.Item("FYY") - 1, _dateData.Item("FYY") - 2), "")
            SetStyleWrapText(reportSheet.Cells(rowIndex + 1, 4 + (i * 3) + offset + offset2))

            If i = 2 Then
                offset += 1
                Dim hdrRange As Range = reportSheet.Cells.CreateRange(rowIndex, 4 + (i * 3) + offset, 2, 1)
                SetStyleForegroundColor(hdrRange, Color.FromArgb(255, 204, 255, 204))
                SetStyleFontBold(hdrRange)
                SetStyleAlignCenter(hdrRange)
                hdrRange.SetOutlineBorders(CellBorderType.Medium, Color.Black)
                reportSheet.Cells(rowIndex, 4 + (i * 3) + offset).Value = _dateData.Item("MM")
                reportSheet.Cells(rowIndex + 1, 4 + (i * 3) + offset).Value = String.Format("FY 20{0}", _dateData.Item("FYY") - 2)
            End If

            If i = 3 Then
                offset += 1
                Dim hdrRange As Range = reportSheet.Cells.CreateRange(rowIndex, 4 + (i * 3) + offset, 2, 1)
                SetStyleForegroundColor(hdrRange, Color.FromArgb(255, 204, 255, 204))
                SetStyleFontBold(hdrRange)
                SetStyleAlignCenter(hdrRange)
                reportSheet.Cells(rowIndex + 1, 4 + (i * 3) + offset).Value = String.Format("FY {0} Plan", _dateData.Item("FYY"))
                SetStyleWrapText(reportSheet.Cells(rowIndex + 1, 4 + (i * 3) + offset))
            End If

            If i = 4 Then
                Dim hdrRange As Range = reportSheet.Cells.CreateRange(rowIndex, 1 + (i * 3) + offset, 2, 4)
                hdrRange.SetOutlineBorders(CellBorderType.Medium, Color.Black)
            End If
        Next

        reportSheet.AutoFitRow(rowIndex + 1)

        Return rowIndex + 1
    End Function

    Private Sub WriteSubTotalRow(ByVal startIndex As Integer, ByVal rowIndex As Integer, ByVal totalText As String, ByVal reportSheet As Worksheet)

        Dim subTotalRange As Range = reportSheet.Cells.CreateRange(rowIndex, 0, 1, 20)
        SetStyleFontBold(subTotalRange)
        SetStyleForegroundColor(subTotalRange, Color.FromArgb(255, 255, 255, 204))

        reportSheet.Cells(rowIndex, 0).Value = "Total " & totalText
        reportSheet.Cells(rowIndex, 1).Formula = String.Format("SUM({0}{1}:{0}{2})", "B", startIndex + 1, rowIndex)
        Dim offset = 0
        Dim offset2 = 0
        For i As Integer = 0 To 4
            'If i = 3 Then offset = 2
            Dim colTY As String = Chr(Asc("A") + (2 + i * 3) + offset)
            Dim colLY As String = Chr(Asc("A") + (3 + i * 3) + offset)

            reportSheet.Cells(rowIndex, 2 + (i * 3) + offset).Formula = String.Format("SUM({0}{1}:{0}{2})", colTY, startIndex + 1, rowIndex)
            If i = 4 Then
                offset2 += 1
                Dim colX As String = Chr(Asc("A") + (1 + i * 3) + offset)
                reportSheet.Cells(rowIndex, 2 + (i * 3) + offset + offset2).Formula = String.Format("({0}{1}-{2}{1})/{2}{1}", colX, rowIndex + 1, colTY)
                colLY = Chr(Asc("A") + (3 + i * 3) + offset + offset2)
            End If
            reportSheet.Cells(rowIndex, 3 + (i * 3) + offset + offset2).Formula = String.Format("SUM({0}{1}:{0}{2})", colLY, startIndex + 1, rowIndex)
            reportSheet.Cells(rowIndex, 4 + (i * 3) + offset + offset2).Formula = String.Format("({0}{1}-{2}{1})/{2}{1}", colTY, rowIndex + 1, colLY)
            If i = 2 Then
                offset += 1
                Dim colTM2YB As String = Chr(Asc("A") + (4 + i * 3) + offset)
                reportSheet.Cells(rowIndex, 4 + (i * 3) + offset).Formula = String.Format("SUM({0}{1}:{0}{2})", colTM2YB, startIndex + 1, rowIndex)
            End If

            If i = 3 Then
                offset += 1
                Dim colX As String = Chr(Asc("A") + (4 + i * 3) + offset)
                reportSheet.Cells(rowIndex, 4 + (i * 3) + offset).Formula = String.Format("SUM({0}{1}:{0}{2})", colX, startIndex + 1, rowIndex)
            End If
        Next

        SetRowBorders(False, rowIndex, reportSheet)
    End Sub

    Private Function WriteTotals(ByVal rowIndex As Integer, ByVal totalText As String, ByVal totalRows As List(Of Integer), ByVal reportSheet As Worksheet, ByVal fgColor As Color) As Integer
        Dim totalRng As Range = reportSheet.Cells.CreateRange(rowIndex, 0, 1, 20)
        SetStyleForegroundColor(totalRng, fgColor)
        SetStyleFontBold(totalRng)
        totalRng.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Medium, Color.Black)
        totalRng.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black)

        reportSheet.Cells(rowIndex, 0).Value = totalText

        Dim doorFormula As String = ""
        Dim doorDelim As String = ""
        For Each subRowIndex As Integer In totalRows
            doorFormula &= doorDelim & "B" & (subRowIndex + 1)
            doorDelim = "+"
        Next
        reportSheet.Cells(rowIndex, 1).Formula = doorFormula

        Dim offset As Integer = 0
        Dim offset2 As Integer = 0
        For i As Integer = 0 To 4
            If i = 4 Then offset2 += 1
            Dim colX As String = Chr(Asc("A") + (1 + i * 3) + offset)
            Dim colTY As String = Chr(Asc("A") + (2 + i * 3) + offset)
            Dim colLY As String = Chr(Asc("A") + (3 + i * 3) + offset + offset2)
            Dim colTM2YB As String = Chr(Asc("A") + (5 + i * 3) + offset)
            Dim formula1 As String = ""
            Dim formula2 As String = ""
            Dim formula3 As String = ""
            Dim formula4 As String = ""
            Dim delimiter As String = ""
            For Each subRowIndex As Integer In totalRows
                formula1 &= delimiter & colTY & (subRowIndex + 1)
                formula2 &= delimiter & colLY & (subRowIndex + 1)
                formula3 &= delimiter & colTM2YB & (subRowIndex + 1)
                formula4 &= delimiter & Chr(Asc(colLY) + 2) & (subRowIndex + 1)
                delimiter = "+"
            Next
            reportSheet.Cells(rowIndex, 2 + (i * 3) + offset).Formula = formula1
            If i = 4 Then
                reportSheet.Cells(rowIndex, 2 + (i * 3) + offset + offset2).Formula = String.Format("({0}{1}-{2}{1})/{2}{1}", colX, rowIndex + 1, colTY)
                colLY = Chr(Asc("A") + (3 + i * 3) + offset + offset2)
            End If
            reportSheet.Cells(rowIndex, 3 + (i * 3) + offset + offset2).Formula = formula2
            reportSheet.Cells(rowIndex, 4 + (i * 3) + offset + offset2).Formula = String.Format("({0}{1}-{2}{1})/{2}{1}", colTY, rowIndex + 1, colLY)
            If i = 2 Then
                offset += 1
                reportSheet.Cells(rowIndex, 4 + (i * 3) + offset).Formula = formula3
            End If

            If i = 3 Then
                offset += 1
                Dim colY As String = Chr(Asc("A") + (4 + i * 3) + offset)
                reportSheet.Cells(rowIndex, 4 + (i * 3) + offset).Formula = formula4
            End If
        Next

        SetRowBorders(True, rowIndex, reportSheet)
        Return rowIndex
    End Function

    Private Sub GetReportData()
        _reportData = ASCDATA1.GetDataTable("SELECT * FROM RSTNATS1 WHERE YYYYPP=:PARM1", "", "V", New Object() {_reportYYYYPP})
        _dateData = ASCDATA1.GetDataRow("SELECT * FROM RSTFTPDP WHERE YYYYPP=:PARM1", "V", New Object() {_reportYYYYPP})
    End Sub

End Class

Public Class MonthlyByDoorRecap
    Inherits RetailExport

    Dim _reportYYYYPP As String
    Dim _reportData As DataTable
    Dim _custCode As String
    Dim _dateData As DataRow

    Public Sub New(ByVal reportYYYYPP As String, ByVal custCode As String)
        _reportYYYYPP = reportYYYYPP
        _custCode = custCode
    End Sub

    Protected Shared Sub SetAsposeLicense()
        Dim license As License = New License()
        license.SetLicense("Aspose.Total.lic")
    End Sub

    Public Sub CreateReport(ByVal fileName As String)
        GetReportData()
        SetAsposeLicense()
        Dim reportWorkbook As Workbook = New Workbook()
        reportWorkbook.Worksheets.RemoveAt(0)

        reportWorkbook.DefaultStyle.Font.Name = "Times New Roman"

        For Each deptCode In New String() {"WM", "MN"}
            Dim reportSheet As Worksheet = reportWorkbook.Worksheets.Add(If(deptCode = "WM", "Women", "Men"))


            reportSheet.PageSetup.Orientation = PageOrientationType.Landscape
            reportSheet.PageSetup.FitToPagesWide = 1
            reportSheet.PageSetup.FitToPagesTall = Nothing

            PrepareSheet(reportSheet, deptCode)
            FillSheet(deptCode, reportSheet)
            reportSheet.AutoFitColumn(0)
        Next
        
        Try
            reportWorkbook.Save(fileName, SaveFormat.Xlsx)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub PrepareSheet(ByVal reportSheet As Worksheet, ByVal deptCode As String)
        Dim numberFormatFlag As New StyleFlag()
        numberFormatFlag.NumberFormat = True

        Dim thousandsNumberFormat As New Style()
        thousandsNumberFormat.Custom = "$#,##0,"

        Dim percentageFormat As New Style()
        percentageFormat.Custom = "#0.0%"

        Dim centerAlignmentStyle As New Style()
        centerAlignmentStyle.HorizontalAlignment = TextAlignmentType.Center
        Dim alignmentFlags As New StyleFlag()
        alignmentFlags.HorizontalAlignment = True

        Dim fontStyle As New Style()
        fontStyle.Font.IsBold = True
        Dim fontFlag As New StyleFlag()
        fontFlag.FontBold = True

        reportSheet.Cells.Columns(0).ApplyStyle(fontStyle, fontFlag)

        For Each i As Integer In New Integer() {1, 2, 4, 5, 7, 8, 10, 11, 13, 14}
            reportSheet.Cells.Columns(i).ApplyStyle(thousandsNumberFormat, numberFormatFlag)
        Next

        For Each i As Integer In New Integer() {3, 6, 9, 12, 15}
            reportSheet.Cells.Columns(i).ApplyStyle(percentageFormat, numberFormatFlag)
        Next

        For i As Integer = 0 To 17
            reportSheet.Cells.Columns(i).ApplyStyle(centerAlignmentStyle, alignmentFlags)
        Next

        reportSheet.Cells.Columns(0).Width = 25

        reportSheet.Zoom = 80

        Dim custName As String = Nothing

        Select Case _custCode
            Case "SAKSFIF10"
                custName = "Saks"
            Case "NORDSTR10"
                custName = "Nordstrom"
            Case "BLOOMIES10"
                custName = "Bloomingdales"
            Case "NEIMANM10"
                custName = "Neiman Marcus"
            Case "HOLTREN10"
                custName = "Holt Renfrew"
        End Select

        reportSheet.PageSetup.PrintTitleRows = "$1:$2"
        reportSheet.PageSetup.SetHeader(1, String.Format("{0} {1}" & vbCrLf &
                                                        "Retail Recap" & vbCrLf &
                                                        "{2} EOM",
                                                        custName,
                                                        If(deptCode = "WM", "Women's", "Men's"),
                                                        _dateData.Item("MM")))
        reportSheet.PageSetup.SetFooter(0, Today.ToString("MM/dd/yyyy"))
        reportSheet.PageSetup.SetFooter(1, String.Format("{0} {1} {2} EOM", custName,
                                                        If(deptCode = "WM", "Women's", "Men's"),
                                                        _dateData.Item("MM")))
    End Sub

    Private Sub FillSheet(ByVal deptCode As String, ByVal reportSheet As Worksheet)
        Dim rowIndex As Integer = 0
        rowIndex = WriteHeaderRow(rowIndex, reportSheet)

        Dim totalRows As New List(Of Integer)
        SetRowBorders(False, rowIndex + 1, reportSheet)
        rowIndex += 1
        For Each rgn In _reportData.AsEnumerable().Where(Function(row) row.Item("DEPT_CODE") = deptCode).OrderBy(Function(row) row.Item("CUST_REGION_SEQ") & If(row.Item("CUST_REGION_CODE") = "OTHER", "ZZZZ", row.Item("CUST_REGION_DESC"))).Select(Function(row) row.Item("CUST_REGION_CODE")).Distinct()
            rowIndex = WriteRegionGrouping(rowIndex, deptCode, rgn, reportSheet)
            totalRows.Add(rowIndex)
            rowIndex += 1
        Next

        rowIndex = WriteTotals(rowIndex + 1, "Grand Total", totalRows, reportSheet)
    End Sub

    Private Function WriteRegionGrouping(ByVal rowIndex As Integer, ByVal deptCode As String, ByVal regionCode As String, ByVal reportSheet As Worksheet) As Integer

        Dim startIndex As Integer = rowIndex
        For Each row As DataRow In _reportData.Select(String.Format("CUST_REGION_CODE='{0}' AND DEPT_CODE='{1}'", regionCode, deptCode), "CUST_STORE_NAME DESC")
            WriteDataRow(row, rowIndex, reportSheet)
            rowIndex += 1
        Next
        Dim regionDesc As String = _reportData.AsEnumerable().Where(Function(row) row.Item("CUST_REGION_CODE") = regionCode).Select(Function(row) row.Item("CUST_REGION_DESC")).Distinct().First()
        WriteRegionTotalRow(startIndex, rowIndex, regionDesc, reportSheet)

        Return rowIndex
    End Function

    Private Sub WriteDataRow(ByVal dataRow As DataRow, ByVal rowIndex As Integer, ByVal reportSheet As Worksheet)
        reportSheet.Cells(rowIndex, 0).Value = dataRow.Item("CUST_STORE_NAME")

        Dim colHeaders = New String() {"FYTD", "RYTD", "MTD", "F3M", "FYB"}

        Dim offset As Integer = 0
        For i As Integer = 0 To 4
            'If i = 3 Then offset = 2
            reportSheet.Cells(rowIndex, 1 + (i * 3) + offset).Value = dataRow.Item(colHeaders(i))
            reportSheet.Cells(rowIndex, 2 + (i * 3) + offset).Value = dataRow.Item(colHeaders(i) & "_LY")
            reportSheet.Cells(rowIndex, 3 + (i * 3) + offset).Formula = String.Format("({0}{1}-{2}{1})/{2}{1}", Chr(Asc("A") + (1 + i * 3) + offset), rowIndex + 1, Chr(Asc("A") + (2 + i * 3) + offset))
        Next

        SetRowBorders(False, rowIndex, reportSheet)
    End Sub


    Private Sub WriteRegionTotalRow(ByVal startIndex As Integer, ByVal rowIndex As Integer, ByVal regionDesc As String, ByVal reportSheet As Worksheet)

        Dim subTotalRange As Range = reportSheet.Cells.CreateRange(rowIndex, 0, 1, 16)
        SetStyleFontBold(subTotalRange)
        SetStyleForegroundColor(subTotalRange, Color.FromArgb(255, 255, 255, 204))

        reportSheet.Cells(rowIndex, 0).Value = "Total " & regionDesc
        Dim offset = 0
        For i As Integer = 0 To 4
            Dim colTY As String = Chr(Asc("A") + (1 + i * 3) + offset)
            Dim colLY As String = Chr(Asc("A") + (2 + i * 3) + offset)
            reportSheet.Cells(rowIndex, 1 + (i * 3) + offset).Formula = String.Format("SUM({0}{1}:{0}{2})", colTY, startIndex + 1, rowIndex)
            reportSheet.Cells(rowIndex, 2 + (i * 3) + offset).Formula = String.Format("SUM({0}{1}:{0}{2})", colLY, startIndex + 1, rowIndex)
            reportSheet.Cells(rowIndex, 3 + (i * 3) + offset).Formula = String.Format("({0}{1}-{2}{1})/{2}{1}", colTY, rowIndex + 1, colLY)
        Next

        SetRowBorders(False, rowIndex, reportSheet)
    End Sub


    Private Sub SetRowBorders(ByVal totalRow As Boolean, ByVal rowIndex As Integer, reportSheet As Worksheet)
        Dim mediumBorderCols = New Integer() {0, 3, 6, 9, 12, 15}
        For i As Integer = 0 To 15
            Dim cellStyle = reportSheet.Cells(rowIndex, i).GetStyle()
            If i = 0 Then
                cellStyle.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black)
            End If

            If Not totalRow Then
                cellStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black)
            End If
            If mediumBorderCols.Contains(i) Or totalRow Then
                cellStyle.SetBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black)
            Else
                cellStyle.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black)
            End If

            reportSheet.Cells(rowIndex, i).SetStyle(cellStyle)
        Next
    End Sub

    Private Function WriteHeaderRow(ByVal rowIndex As Integer, ByVal reportSheet As Worksheet) As Integer
        Dim accountRng As Range = reportSheet.Cells.CreateRange(rowIndex, 0, 2, 1)
        accountRng.Merge()
        SetStyleForegroundColor(accountRng, Color.FromArgb(255, 153, 204, 255))
        SetStyleFontBold(accountRng)
        SetStyleAlignCenter(accountRng)
        SetStyleAlignCenterVertical(accountRng)
        accountRng.SetOutlineBorders(CellBorderType.Medium, Color.Black)
        accountRng.Value = "Account"

        Dim colHeaders = New String() {String.Format("JH FY{0} ({1})", _dateData.Item("FYY"), _dateData.Item("DATE_RANGE_FY")),
                                       String.Format("RY{0} ({1})", If(_custCode = "NEIMANM10", _dateData.Item("FYY"), _dateData.Item("RYY")), If(_custCode = "NEIMANM10", _dateData.Item("DATE_RANGE_FY"), _dateData.Item("DATE_RANGE_RY"))),
                                       String.Format("{0}", _dateData.Item("MM")),
                                       String.Format("3 Month Trend ({0})", _dateData.Item("DATE_RANGE_3MB")),
                                       ""}

        Dim offset As Integer = 0
        For i As Integer = 0 To 4
            'If i = 3 Then offset = 2
            Dim TyLyRng As Range = reportSheet.Cells.CreateRange(rowIndex, 1 + (i * 3) + offset, 2, 3)
            SetStyleForegroundColor(TyLyRng, Color.FromArgb(255, 204, 255, 204))
            SetStyleFontBold(TyLyRng)
            SetStyleAlignCenter(TyLyRng)
            TyLyRng.SetOutlineBorders(CellBorderType.Medium, Color.Black)

            Dim headerTextRng As Range = reportSheet.Cells.CreateRange(rowIndex, 1 + (i * 3) + offset, 1, 3)
            headerTextRng.Merge()
            headerTextRng.Value = colHeaders(i)

            reportSheet.Cells(rowIndex + 1, 1 + (i * 3) + offset).Value = If(i = 4, String.Format("FY{0}", _dateData.Item("FYY") - 1), "TY") '"TY"
            reportSheet.Cells(rowIndex + 1, 2 + (i * 3) + offset).Value = If(i = 4, String.Format("FY{0}", _dateData.Item("FYY") - 2), "LY") ' "LY"
            reportSheet.Cells(rowIndex + 1, 3 + (i * 3) + offset).Value = "% CHG"

        Next

        Return rowIndex + 1
    End Function

    Private Function WriteTotals(ByVal rowIndex As Integer, ByVal totalText As String, ByVal totalRows As List(Of Integer), ByVal reportSheet As Worksheet) As Integer
        Dim totalRng As Range = reportSheet.Cells.CreateRange(rowIndex, 0, 1, 16)
        SetStyleForegroundColor(totalRng, Color.FromArgb(255, 204, 255, 204))
        SetStyleFontBold(totalRng)
        totalRng.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Medium, Color.Black)
        totalRng.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black)

        reportSheet.Cells(rowIndex, 0).Value = totalText

        Dim offset As Integer = 0
        For i As Integer = 0 To 4
            'If i = 3 Then offset = 2
            Dim colTY As String = Chr(Asc("A") + (1 + i * 3) + offset)
            Dim colLY As String = Chr(Asc("A") + (2 + i * 3) + offset)
            Dim formula1 As String = ""
            Dim formula2 As String = ""
            Dim delimiter As String = ""
            For Each subRowIndex As Integer In totalRows
                formula1 &= delimiter & colTY & (subRowIndex + 1)
                formula2 &= delimiter & colLY & (subRowIndex + 1)
                delimiter = "+"
            Next
            reportSheet.Cells(rowIndex, 1 + (i * 3) + offset).Formula = formula1
            reportSheet.Cells(rowIndex, 2 + (i * 3) + offset).Formula = formula2
            reportSheet.Cells(rowIndex, 3 + (i * 3) + offset).Formula = String.Format("({0}{1}-{2}{1})/{2}{1}", colTY, rowIndex + 1, colLY)
        Next

        SetRowBorders(True, rowIndex, reportSheet)
        Return rowIndex
    End Function

    Private Sub GetReportData()
        _reportData = ASCDATA1.GetDataTable("SELECT * FROM RSTNATM1 WHERE YYYYPP=:PARM1 AND CUST_CODE=:PARM2", "", "VV", New Object() {_reportYYYYPP, _custCode})
        _dateData = ASCDATA1.GetDataRow("SELECT * FROM RSTFTPDP WHERE YYYYPP=:PARM1", "V", New Object() {_reportYYYYPP})
    End Sub

End Class

Public Class IndependentsByRegionReport
    Inherits RetailExport

    Dim _reportData As DataTable
    Dim _fiscalYear As String
    Dim _regionCode As String
    Dim maxMonth As Integer

    Public Sub New(ByVal fiscalYear As String)
        _fiscalYear = fiscalYear
    End Sub

    Public Sub New(ByVal fiscalYear As String, ByVal salesRep As String)
        _fiscalYear = fiscalYear
        _regionCode = salesRep
    End Sub

    Protected Shared Sub SetAsposeLicense()
        Dim license As License = New License()
        license.SetLicense("Aspose.Total.lic")
    End Sub

    Private Sub GetReportData()
        If _regionCode Is Nothing Then
            _reportData = ASCDATA1.GetDataTable("SELECT * FROM RSTINDY1 WHERE YYYY=:PARM1", "", "V", New Object() {_fiscalYear})
        Else
            _reportData = ASCDATA1.GetDataTable("SELECT * FROM RSTINDY1 WHERE YYYY=:PARM1 AND REGION_CODE=:PARM2", "", "VV", New Object() {_fiscalYear, _regionCode})
        End If

        Dim mm As Integer = 1
        For i As Integer = 1 To 12
            Dim curIndex = i
            If _reportData.AsEnumerable.Any(Function(row) row.Item(curIndex.ToString.PadLeft(2, "0") & "_TY") IsNot DBNull.Value) Then
                mm = i
            End If
        Next

        maxMonth = mm
    End Sub

    Public Sub CreateReport(ByVal filename As String)
        SetAsposeLicense()
        GetReportData()

        Dim reportWorkbook As Workbook = New Workbook()
        reportWorkbook.Worksheets.RemoveAt(0)

        reportWorkbook.DefaultStyle.Font.Name = "Arial"

        Dim regions = GetDistinct("REGION_CODE", _reportData)

        If _regionCode Is Nothing Then 'Create summary sheet
            Dim reportSheet As Worksheet = reportWorkbook.Worksheets.Add("FISCAL " & _fiscalYear)

            reportSheet.PageSetup.FitToPagesWide = 1
            reportSheet.PageSetup.FitToPagesTall = Nothing

            PrepareSheet(reportSheet)
            FillSheet("", reportSheet)
            reportSheet.AutoFitColumn(0)
        End If

        For Each regionCode In regions
            Dim reportSheet As Worksheet = reportWorkbook.Worksheets.Add(regionCode)

            reportSheet.PageSetup.FitToPagesWide = 1
            reportSheet.PageSetup.FitToPagesTall = Nothing

            PrepareSheet(reportSheet)
            FillSheet(regionCode, reportSheet)
            reportSheet.AutoFitColumn(0)
        Next
        Try
            filename = filename.Replace("/", "")
            reportWorkbook.Save(filename, SaveFormat.Xlsx)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub PrepareSheet(ByVal reportSheet As Worksheet)
        reportSheet.Cells.Columns(2).Width = 0.1
        Dim index = reportSheet.ConditionalFormattings.Add()

        Dim fcs As FormatConditionCollection = reportSheet.ConditionalFormattings(index)

        Dim conditionIndex As Integer = fcs.AddCondition(FormatConditionType.ContainsBlanks, Nothing, Nothing, Nothing)

        Dim fc As FormatCondition = fcs(conditionIndex)
        fc.Style.BackgroundColor = Color.Yellow

    End Sub

    Private Sub FillSheet(ByVal regionCode As String, ByVal reportSheet As Worksheet)
        reportSheet.Name = String.Format("Fiscal {0}" & If(regionCode <> "", "-" & regionCode, ""), _fiscalYear)

        Dim rowIndex As Integer = 2

        reportSheet.Cells(rowIndex, 0).Value = "RETAIL SALES REPORT"
        SetStyleFontBold(reportSheet.Cells(rowIndex, 0))
        reportSheet.Cells(rowIndex + 1, 0).Value = "FISCAL " & _fiscalYear
        SetStyleFontBold(reportSheet.Cells(rowIndex, 1))

        rowIndex += 4
        rowIndex = WriteMonthHeaders(rowIndex, reportSheet)

        Dim regionCodes = (From row In _reportData.AsEnumerable()
                Where row.Item("REGION_CODE") = If(regionCode = "", row.Item("REGION_CODE"), regionCode)
                Select row.Field(Of String)("REGION_CODE") Distinct).ToList()

        Dim regionSumList As String = ""
        Dim regionDelimiter = "="

        For Each regionCode In regionCodes
            Dim currentRegionCode = regionCode
            Dim regionDesc As String = _reportData.AsEnumerable().Where(Function(row) row.Item("REGION_CODE") = currentRegionCode).Select(Function(row) row.Item("REGION_DESC")).FirstOrDefault()

            rowIndex = WriteYearHeaders(regionDesc, rowIndex + 1, reportSheet)


            Dim srepCodes = (From row In _reportData.AsEnumerable()
                    Where row.Item("REGION_CODE") = currentRegionCode
                    Select row.Field(Of String)("SREP_CODE") Distinct).ToList()

            Dim srepSumList As String = ""
            Dim delimiter As String = "="
            For Each srepCode In srepCodes
                rowIndex = WriteSalesRepSection(srepCode, rowIndex + 1, reportSheet)
                rowIndex += 1

                srepSumList &= delimiter & "{0}" & rowIndex
                delimiter = "+"
            Next

            rowIndex = WriteTotals(regionDesc, srepSumList, rowIndex + 1, reportSheet)
            rowIndex += 1

            regionSumList &= regionDelimiter & "{0}" & rowIndex
            regionDelimiter = "+"
        Next

        If regionCode = "" Then
            WriteTotals("GRAND", regionSumList, rowIndex + 1, reportSheet)
        End If
    End Sub

    Private Function WriteMonthHeaders(ByVal rowIndex As Integer, ByVal reportSheet As Worksheet) As Integer
        Dim offset As Integer = 0
        Dim yyyy As Integer = Convert.ToInt32(_fiscalYear) - 1
        For i As Integer = 1 To 12 'Each month
            If i = 6 Then
                yyyy += 1
            End If
            Dim rng = reportSheet.Cells.CreateRange(rowIndex, 3 * i + offset, 1, 3)
            rng.Merge()
            rng.SetOutlineBorders(CellBorderType.Medium, Color.Black)
            SetStyleFontBold(reportSheet.Cells.Rows(rowIndex))
            SetStyleForegroundColor(rng, Color.FromArgb(255, 204, 255, 255))
            rng.Value = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName((i + 6) Mod 12 + 1)

            If i Mod 6 = 0 Then
                offset = 3
                If i = 12 Then offset = 6
                Dim stdRng = reportSheet.Cells.CreateRange(rowIndex, 3 * i + offset, 1, 3)
                stdRng.Merge()
                stdRng.SetOutlineBorders(CellBorderType.Medium, Color.Black)
                SetStyleForegroundColor(stdRng, Color.FromArgb(255, 204, 255, 255))

                Dim seasonName As String = If(i > 6, "SPRING ", "FALL ") & If(i > 6, yyyy, yyyy - 1)
                stdRng.Value = seasonName

                If i = 12 Then
                    offset = 9
                    Dim fytdRng = reportSheet.Cells.CreateRange(rowIndex, 3 * i + offset, 1, 3)
                    fytdRng.Merge()
                    SetStyleForegroundColor(fytdRng, Color.FromArgb(255, 204, 255, 255))
                    fytdRng.SetOutlineBorders(CellBorderType.Medium, Color.Black)
                    fytdRng.Value = "TOTAL FISCAL " & _fiscalYear
                End If
            End If
        Next

        Return rowIndex
    End Function

    Private Function WriteYearHeaders(ByVal regionDesc As String, ByVal rowIndex As Integer, ByVal reportSheet As Worksheet) As Integer
        Dim offset As Integer = 0
        Dim yyyy As Integer = Convert.ToInt32(_fiscalYear) - 1

        reportSheet.Cells(rowIndex, 0).Value = String.Format("Account / {0}", regionDesc)
        reportSheet.Cells(rowIndex, 1).Value = "# of doors"
        SetStyleFontBold(reportSheet.Cells.Rows(rowIndex))
        SetStyleForegroundColor(reportSheet.Cells.CreateRange(rowIndex, 0, 1, 2), Color.FromArgb(255, 204, 255, 255))
        SetRowBorders(True, rowIndex, reportSheet)

        For i As Integer = 1 To 12 'Each month
            If i = 6 Then
                yyyy += 1
            End If

            SetStyleForegroundColor(reportSheet.Cells.CreateRange(rowIndex, 3 * i + offset, 1, 3), Color.FromArgb(255, 204, 255, 255))
            reportSheet.Cells(rowIndex, 3 * i + offset).Value = yyyy
            reportSheet.Cells(rowIndex, 3 * i + offset + 1).Value = yyyy - 1
            reportSheet.Cells(rowIndex, 3 * i + offset + 2).Value = "%VAR"

            If i Mod 6 = 0 Then
                offset = 3
                If i = 12 Then offset = 6

                SetStyleForegroundColor(reportSheet.Cells.CreateRange(rowIndex, 3 * i + offset, 1, 3), Color.FromArgb(255, 204, 255, 255))
                reportSheet.Cells(rowIndex, 3 * i + offset).Value = If(i = 6, yyyy - 1, yyyy)
                reportSheet.Cells(rowIndex, 3 * i + offset + 1).Value = If(i = 6, yyyy - 1, yyyy) - 1
                reportSheet.Cells(rowIndex, 3 * i + offset + 2).Value = "%VAR"

                If i = 12 Then
                    offset = 9
                    SetStyleForegroundColor(reportSheet.Cells.CreateRange(rowIndex, 3 * i + offset, 1, 3), Color.FromArgb(255, 204, 255, 255))
                    reportSheet.Cells(rowIndex, 3 * i + offset).Value = yyyy
                    reportSheet.Cells(rowIndex, 3 * i + offset + 1).Value = yyyy - 1
                    reportSheet.Cells(rowIndex, 3 * i + offset + 2).Value = "%VAR"
                End If
            End If
        Next

        Return rowIndex
    End Function

    Private Sub SetRowBorders(ByVal boldHorizontals As Boolean, ByVal rowIndex As Integer, ByVal reportsheet As Worksheet)
        If boldHorizontals Then
            Dim cellStyle1 = reportsheet.Cells(rowIndex, 0).GetStyle()
            cellStyle1.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black)
            cellStyle1.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black)
            cellStyle1.SetBorder(BorderType.TopBorder, CellBorderType.Medium, Color.Black)
            cellStyle1.SetBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black)
            reportsheet.Cells(rowIndex, 0).SetStyle(cellStyle1)
            Dim cellStyle2 = reportsheet.Cells(rowIndex, 1).GetStyle()
            cellStyle2.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black)
            cellStyle2.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black)
            cellStyle2.SetBorder(BorderType.TopBorder, CellBorderType.Medium, Color.Black)
            cellStyle2.SetBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black)
            reportsheet.Cells(rowIndex, 1).SetStyle(cellStyle2)
        End If

        For i As Integer = 3 To 45 Step 3
            Dim cellStyle = reportsheet.Cells(rowIndex, i).GetStyle()
            cellStyle.SetBorder(BorderType.LeftBorder, CellBorderType.Medium, Color.Black)
            cellStyle.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black)
            If boldHorizontals Then
                cellStyle.SetBorder(BorderType.TopBorder, CellBorderType.Medium, Color.Black)
                cellStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black)
            Else
                cellStyle.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black)
                cellStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black)
            End If
            reportsheet.Cells(rowIndex, i).SetStyle(cellStyle)
        Next
        For i As Integer = 4 To 46 Step 3
            Dim cellStyle = reportsheet.Cells(rowIndex, i).GetStyle()
            cellStyle.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black)
            cellStyle.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black)
            If boldHorizontals Then
                cellStyle.SetBorder(BorderType.TopBorder, CellBorderType.Medium, Color.Black)
                cellStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black)
            Else
                cellStyle.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black)
                cellStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black)
            End If
            reportsheet.Cells(rowIndex, i).SetStyle(cellStyle)
        Next
        For i As Integer = 5 To 47 Step 3
            Dim cellStyle = reportsheet.Cells(rowIndex, i).GetStyle()
            cellStyle.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black)
            cellStyle.SetBorder(BorderType.RightBorder, CellBorderType.Medium, Color.Black)
            If boldHorizontals Then
                cellStyle.SetBorder(BorderType.TopBorder, CellBorderType.Medium, Color.Black)
                cellStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black)
            Else
                cellStyle.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black)
                cellStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black)
            End If
            reportsheet.Cells(rowIndex, i).SetStyle(cellStyle)
        Next
    End Sub

    Private Function WriteSalesRepSection(ByVal srepCode As String, ByVal rowIndex As Integer, ByVal reportSheet As Worksheet) As Integer
        Dim custCode As String = ""
        Dim startIndex As Integer = rowIndex
        Dim sumList As String = ""
        Dim delimiter As String = "="
        For Each salesRow As DataRow In _reportData.Select(String.Format("SREP_CODE='{0}'", srepCode))
            If custCode = "" Then
                custCode = salesRow.Item("CUST_CODE")
            End If

            If custCode <> salesRow.Item("CUST_CODE") Then
                custCode = salesRow.Item("CUST_CODE")
                If _reportData.AsEnumerable.Where(Function(row) row.Item("CUST_CODE") = custCode).Select(Function(row) Convert.ToInt32(row.Item("GRP"))).Max() > 0 Then
                    SetRowBorders(False, rowIndex, reportSheet)
                    rowIndex += 1
                End If
            End If

            WriteSalesRow(salesRow, rowIndex, reportSheet)


            rowIndex += 1

            If salesRow.Item("GRP") = "1" Then
                SetRowBorders(False, rowIndex, reportSheet)
                rowIndex += 1
            Else
                sumList &= delimiter & "{0}" & rowIndex.ToString
                delimiter = "+"
            End If
        Next

        Dim srepName As String = _reportData.AsEnumerable().Where(Function(row) row.Item("SREP_CODE") = srepCode).Select(Function(row) row.Item("SREP_NAME")).FirstOrDefault()

        rowIndex = WriteSalesRepTotalRow(srepName, sumList, rowIndex, reportSheet)

        Return rowIndex
    End Function

    Private Function WriteSalesRepTotalRow(ByVal srepName As String, ByVal sumList As String, ByVal rowIndex As Integer, ByVal reportSheet As Worksheet) As Integer
        SetStyleForegroundColor(reportSheet.Cells.Rows(rowIndex), Color.FromArgb(255, 255, 255, 153))
        SetStyleFontBold(reportSheet.Cells.Rows(rowIndex))
        SetRowBorders(True, rowIndex, reportSheet)
        'FormatRow(rowIndex, reportSheet)

        Dim offset As Integer = 0
        Dim yyyy As Integer = Convert.ToInt32(_fiscalYear) - 1


        reportSheet.Cells(rowIndex, 0).Value = srepName & " TOTAL"
        reportSheet.Cells(rowIndex, 1).Formula = String.Format(sumList, "B") 'sum startIndex to rowIndex

        Dim colListTY As String = Nothing
        Dim colListLY As String = Nothing
        Dim delimiter As String = "="

        For i As Integer = 1 To 12 'Each month
            If i = 6 Then
                yyyy += 1
            End If

            colListTY &= delimiter & GetExcelColumnName(3 * i + offset + 1) & "{0}"
            colListLY &= delimiter & GetExcelColumnName(3 * i + offset + 2) & "{0}"
            delimiter = "+"

            reportSheet.Cells(rowIndex, 3 * i + offset).Formula = String.Format(sumList, GetExcelColumnName(3 * i + offset + 1))
            reportSheet.Cells(rowIndex, 3 * i + offset + 1).Formula = String.Format(sumList, GetExcelColumnName(3 * i + offset + 2))
            reportSheet.Cells(rowIndex, 3 * i + offset + 2).Formula = String.Format("=({0}-{1})/ABS({1})", ConvertFromR1C1(rowIndex + 1, 3 * i + offset + 1), ConvertFromR1C1(rowIndex + 1, 3 * i + offset + 2))

            If i Mod 6 = 0 Then
                offset = 3
                If i = 12 Then offset = 6

                reportSheet.Cells(rowIndex, 3 * i + offset).Formula = String.Format(colListTY, rowIndex + 1)
                reportSheet.Cells(rowIndex, 3 * i + offset + 1).Formula = String.Format(colListLY, rowIndex + 1)
                reportSheet.Cells(rowIndex, 3 * i + offset + 2).Formula = String.Format("=({0}-{1})/ABS({1})", ConvertFromR1C1(rowIndex + 1, 3 * i + offset + 1), ConvertFromR1C1(rowIndex + 1, 3 * i + offset + 2))

                colListTY = ""
                colListLY = ""
                delimiter = "="

                If i = 12 Then
                    offset = 9
                    reportSheet.Cells(rowIndex, 3 * i + offset).Formula = String.Format("=V{0}+AQ{0}", rowIndex + 1)
                    reportSheet.Cells(rowIndex, 3 * i + offset + 1).Formula = String.Format("=W{0}+AR{0}", rowIndex + 1)
                    reportSheet.Cells(rowIndex, 3 * i + offset + 2).Formula = String.Format("=(AT{0}-AU{0})/AU{0}", rowIndex + 1)
                End If
            End If
        Next

        FormatRow(rowIndex, reportSheet)

        Return rowIndex
    End Function

    Private Sub WriteSalesRow(ByVal salesRow As DataRow, ByVal rowIndex As Integer, ByVal reportSheet As Worksheet)
        Dim offset As Integer = 0
        Dim yyyy As Integer = Convert.ToInt32(_fiscalYear) - 1

        SetRowBorders(False, rowIndex, reportSheet)
        SetStyleForegroundColor(reportSheet.Cells(rowIndex, 0), Color.FromArgb(255, 255, 255, 153))

        FormatRow(rowIndex, reportSheet)

        If salesRow.Item("GRP") = "1" Then
            SetStyleFontBold(reportSheet.Cells.CreateRange(rowIndex, 0, 1, 47))
            salesRow.Item("CUST_STORE_NAME") &= " TOTAL"
        End If

        Dim startBlanksColumn As Integer = 3
        Dim doorOpenPeriod As String = salesRow.Item("YYYYPP_DOOR_OPENED") & ""
        If Not String.IsNullOrEmpty(doorOpenPeriod) Then
            If doorOpenPeriod.Substring(0, 4) = _fiscalYear Then
                startBlanksColumn = 2 + CType(doorOpenPeriod.Substring(4, 2), Integer) + If(CType(doorOpenPeriod.Substring(4, 2), Integer) > 6, 3, 0)
            End If
        End If

        Dim ca As CellArea = CellArea.CreateCellArea(rowIndex, startBlanksColumn, rowIndex, 2 + 3 * maxMonth + If(maxMonth > 6, 3, 0))
        Dim fcs = reportSheet.ConditionalFormattings(0)
        If (salesRow.Item("OPEN_DOOR") = "1") Then
            fcs.AddArea(ca)
        End If

        reportSheet.Cells(rowIndex, 0).Value = If(salesRow.Item("GRP") = "0", salesRow.Item("CUST_STORE_NAME"), salesRow.Item("CUST_NAME"))
        reportSheet.Cells(rowIndex, 1).Value = salesRow.Item("STR_CNT")

        Dim colListTY As String = Nothing
        Dim colListLY As String = Nothing
        Dim delimiter As String = "="

        For i As Integer = 1 To 12 'Each month
            If i = 6 Then
                yyyy += 1
            End If

            colListTY &= delimiter & GetExcelColumnName(3 * i + offset + 1) & "{0}"
            colListLY &= delimiter & GetExcelColumnName(3 * i + offset + 2) & "{0}"
            delimiter = "+"

            If salesRow.Item("GRP") > 0 Then
                reportSheet.Cells(rowIndex, 3 * i + offset).Formula = String.Format("=SUM({0}:{1})", ConvertFromR1C1(rowIndex - salesRow.Item("ACTUAL_CNT") + 1, 3 * i + offset + 1), ConvertFromR1C1(rowIndex, 3 * i + offset + 1)) 'salesRow.Item(i.ToString("00") & "_TY") & ""
                reportSheet.Cells(rowIndex, 3 * i + offset + 1).Formula = String.Format("=SUM({0}:{1})", ConvertFromR1C1(rowIndex - salesRow.Item("ACTUAL_CNT") + 1, 3 * i + offset + 2), ConvertFromR1C1(rowIndex, 3 * i + offset + 2)) 'salesRow.Item(i.ToString("00") & "_LY") & ""
            Else
                reportSheet.Cells(rowIndex, 3 * i + offset).Formula = salesRow.Item(i.ToString("00") & "_TY") & ""
                reportSheet.Cells(rowIndex, 3 * i + offset + 1).Formula = salesRow.Item(i.ToString("00") & "_LY") & ""
            End If
            reportSheet.Cells(rowIndex, 3 * i + offset + 2).Formula = String.Format("=({0}-{1})/ABS({1})", ConvertFromR1C1(rowIndex + 1, 3 * i + offset + 1), ConvertFromR1C1(rowIndex + 1, 3 * i + offset + 2))

            If i Mod 6 = 0 Then
                offset = 3
                If i = 12 Then offset = 6

                SetStyleForegroundColor(reportSheet.Cells.CreateRange(rowIndex, 3 * i + offset, 1, 3), Color.FromArgb(255, 255, 204, 153))

                reportSheet.Cells(rowIndex, 3 * i + offset).Formula = String.Format(colListTY, rowIndex + 1) & ""
                reportSheet.Cells(rowIndex, 3 * i + offset + 1).Formula = String.Format(colListLY, rowIndex + 1) & ""
                reportSheet.Cells(rowIndex, 3 * i + offset + 2).Formula = String.Format("=({0}-{1})/ABS({1})", ConvertFromR1C1(rowIndex + 1, 3 * i + offset + 1), ConvertFromR1C1(rowIndex + 1, 3 * i + offset + 2))

                colListTY = ""
                colListLY = ""
                delimiter = "="

                If i = 12 Then
                    offset = 9
                    reportSheet.Cells(rowIndex, 3 * i + offset).Formula = String.Format("=V{0}+AQ{0}", rowIndex + 1)
                    reportSheet.Cells(rowIndex, 3 * i + offset + 1).Formula = String.Format("=W{0}+AR{0}", rowIndex + 1)
                    reportSheet.Cells(rowIndex, 3 * i + offset + 2).Formula = String.Format("=(AT{0}-AU{0})/AU{0}", rowIndex + 1)
                End If
            End If
        Next
    End Sub

    Private Sub FormatRow(ByVal rowIndex As Integer, ByVal reportSheet As Worksheet)
        For i As Integer = 1 To 15
            SetDollarFormat(reportSheet.Cells(rowIndex, 3 * i))
            SetDollarFormat(reportSheet.Cells(rowIndex, 3 * i + 1))
            SetPercentageFormat(reportSheet.Cells(rowIndex, 3 * i + 2))
        Next
    End Sub

    Private Function WriteTotals(ByVal regionDesc As String, ByVal sumList As String, ByVal rowIndex As Integer, ByVal reportSheet As Worksheet) As Integer
        reportSheet.Cells(rowIndex, 0).Value = If(regionDesc = "GRAND", "GRAND TOTAL", regionDesc & " Region TOTAL")
        reportSheet.Cells(rowIndex, 1).Formula = String.Format(sumList, "B")
        FormatRow(rowIndex, reportSheet)
        SetStyleForegroundColor(reportSheet.Cells.Rows(rowIndex), If(regionDesc = "GRAND", Color.FromArgb(255, 255, 204, 153), Color.FromArgb(255, 255, 255, 0)))
        SetStyleFontBold(reportSheet.Cells.Rows(rowIndex))
        SetRowBorders(True, rowIndex, reportSheet)

        For i As Integer = 1 To 15
            reportSheet.Cells(rowIndex, i * 3).Formula = String.Format(sumList, GetExcelColumnName(i * 3 + 1))
            reportSheet.Cells(rowIndex, i * 3 + 1).Formula = String.Format(sumList, GetExcelColumnName(i * 3 + 2))
            reportSheet.Cells(rowIndex, i * 3 + 2).Formula = String.Format("=({0}-{1})/ABS({1})", ConvertFromR1C1(rowIndex + 1, 3 * i + 1), ConvertFromR1C1(rowIndex + 1, 3 * i + 2))
        Next

        Return rowIndex
    End Function

End Class