Public Class DPRFCPR1

#Region "Declarations"
    Dim DPTFCPR1 As String
    Dim SOTPRICX As String
    Dim YYYY As String

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim YEARs As New List(Of String)
        For Y As Integer = Val(Now.Year) - 2 To Val(Mid(ASCMAIN1.CYP, 1, 4)) + 1
            YEARs.Add(Format(Y, "0000"))
        Next
        cmbOPS_YYYY.DataSource = YEARs
        cmbOPS_YYYY.Value = Now.Year

    End Sub

    Protected Overrides Sub Build_Workfile()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Building Work File")

        Create_Pivot()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Overrides Sub Build_Report_File_Post_Process()

    End Sub

    Public Overrides Sub Print_Report()

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            ASCMAIN1.sql = "SELECT CUST_CODE, COUNT (*) CUSTS, MIN (MARKET_CODE) MARKET_CODE1, MAX(MARKET_CODE) MARKET_CODE2 FROM SOTMKTC1 WHERE CUST_CODE IS NOT NULL GROUP BY CUST_CODE HAVING COUNT (*) > 1"
            Dim row As DataRow = ASCDATA1.GetDataRow
            If row IsNot Nothing Then
                EMsg &= vbCr & "Problem with Customer " & row.Item("CUST_CODE") & " - in Markets: " & row.Item("MARKET_CODE1") & " and " & row.Item("MARKET_CODE2")
            End If

            Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("MARKET_CODE")
            If rowASTDSQLA.Item("EXCLUDE") & "" = "1" Then
                EMsg &= vbCr & "You may not use Exclude in a Filter"
            Else
                Dim MARKET_CODEs As String = SQLA("MARKET_CODE")
                If MARKET_CODEs <> "" Then
                    For Each MARKET_CODE As String In Split(MARKET_CODEs, ",")
                        Dim rowSOTMKTC1 As DataRow = LookUp("SOTMKTC1", MARKET_CODE)
                        If rowSOTMKTC1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Market Code (" & MARKET_CODE & ")"
                        ElseIf rowSOTMKTC1.Item("CUST_CODE") & "" = "" Then
                            EMsg &= vbCr & "Invalid Market Code (" & MARKET_CODE & ") - no Customer Defined"
                        End If
                    Next
                End If
            End If
        End If
    End Sub

    Sub Create_Pivot()

        YYYY = cmbOPS_YYYY.Value
        Dim YYYY_ly As String = Format(Val(YYYY) - 1, "0000")

        Dim CM As Integer = 0 ' CURRENT MONTH INDEX, 0 = BEFORE 1, 13 = AFTER 12
        If YYYY < Mid(ASCMAIN1.CYP, 1, 4) Then
            CM = 13
        ElseIf YYYY > Mid(ASCMAIN1.CYP, 1, 4) Then
            CM = 0
        Else
            CM = Val(Mid(ASCMAIN1.CYP, 5, 2))
        End If

        Dim DATES(12) As Date
        ASCMAIN1.sql = "Select * from GLTPARM2 where OPS_YYYYPP between '" & YYYY & "01' and '" & YYYY & "12'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "OPS_YYYYPP")
            Dim PP As Integer = Val(Mid(row.Item("OPS_YYYYPP"), 5, 2))
            DATES(PP) = row.Item("PRD_END_DATE")
        Next

        Dim YPC As String = ASCMAIN1.CYP

        ' YPA1 AND YPA2 ARE THE STARTING AND ENDING PERIODS USED FOR ACTUALS
        Dim YPA1 As String = YYYY & "01"
        Dim YPA2 As String = YYYY & "12"
        ' YPF1 AND YPF2 ARE THE STARTING AND ENDING PERIODS USED FOR FORECASTS
        Dim YPF1 As String = YYYY & "01"
        Dim YPF2 As String = YYYY & "12"
        If YPA1 >= YPC Then ' THIS ONE IS A HARD ONE TO REASON ABOUT
            ' NO HISTORY
            YPA1 = ""
            YPA2 = ""
        ElseIf YYYY = Mid(YPC, 1, 4) Then
            ' HISTORY UP UNTIL LAST MONTH
            YPA2 = ASCMAIN1.Period_Calc(YPC, -1)
            YPF1 = YPC
        Else
            ' NO FORECAST
            YPF1 = ""
            YPF2 = ""
        End If



        Dim SQLW As String = ""
        SQLW &= SQLA_filter("MARKET_CODE", "DPTITMF1")

        ASCMAIN1.Progress("Now Creating Workbook")

        Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & ".xlsx"
        Dim DataTable As DataTable
        Dim r As Integer = 0

        Dim excel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlSourceRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim xlDestRange As Microsoft.Office.Interop.Excel.Range = Nothing

        excel = New Microsoft.Office.Interop.Excel.Application
        wb = excel.Workbooks.Open(FILENAME)

        Dim sqlSumS As String = ""
        Dim sqlSumF As String = ""

        Dim sqlSum As String = ""
        Dim sqlAll As String = ""
        For M As Integer = 1 To 12

            Dim MM As String = Format(M, "00") '
            sqlAll &= ", SOTPRICX.P" & MM & " P" & MM
            For Each T As String In New String() {"U", "S", "C", "M"}
                sqlSum &= ", Sum(" & T & MM & ") " & T & MM
                sqlAll &= ", " & T & MM
            Next

            Dim sqlSumSpfx As String = ", Sum (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & YYYY & MM & "',NVL(SOTINVH2.ORDR_QTY_SHIP,0)"
            sqlSumS &= sqlSumSpfx & ",0)) U" & MM
            sqlSumS &= sqlSumSpfx & " * NVL(SOTINVH2.ORDR_UNIT_PRICE,0)/1000,0)) S" & MM
            sqlSumS &= sqlSumSpfx & " * NVL(SOTINVH2.ITEM_UNIT_COST,0)/1000,0)) C" & MM
            sqlSumS &= ", 0 M" & MM

            Dim sqlSumFpfx As String = ", Sum (CASE WHEN DPTITMF1.OPS_YYYYPP_FC = '" & YYYY & MM & "' THEN NVL(DPTITMF1.FORECAST,0)"
            sqlSumF &= sqlSumFpfx & " ELSE 0 END) U" & MM
            sqlSumF &= ", 0 S" & MM ' sqlSumFpfx & " * NVL(SOTPRIC2.ITEM_PRICE,0) ELSE 0 END) S" & MM
            sqlSumF &= ", 0 C" & MM '  sqlSumFpfx & " * NVL(ICTITEM1.ITEM_COST_STD,0) ELSE 0 END) C" & MM
            sqlSumF &= ", 0 M" & MM
        Next

        ASCMAIN1.Progress("Work Table")
        ASCMAIN1.sql = "Select X.MARKET_CODE, X.ITEM_CODE, ARTCUST1.PRICE_LIST_CODE, SOTTCLS1.CHANNEL_CODE, SOTMKTC1.CUST_CODE" & vbCrLf _
            & sqlSum & vbCrLf _
            & " from ((" & vbCrLf _
            & "Select MARKET_CODE, ITEM_CODE" & sqlSumF & vbCrLf _
            & " from DPTITMF1" & vbCrLf _
            & " where OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "   and OPS_YYYYPP_FC between '" & YPF1 & "' and '" & YPF2 & "'" & vbCrLf _
            & Replace(SQLW, "DPTITMF1", "DPTITMF1") & vbCrLf _
            & " group by MARKET_CODE, ITEM_CODE" & vbCrLf _
            & ") union (" & vbCrLf _
            & "Select SOTMKTC1.MARKET_CODE, SOTINVH2.ITEM_CODE" & sqlSumS & vbCrLf _
            & " from SOTINVH2,ARTCUST1,SOTMKTC1" & vbCrLf _
            & " where SOTINVH2.ORDR_YYYYPP_UPDATED between '" & YPA1 & "' and '" & YPA2 & "'" & vbCrLf _
            & "   and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
            & "   and SOTMKTC1.CUST_CODE = ARTCUST1.CUST_CODE" & vbCrLf _
            & Replace(SQLW, "DPTITMF1", "SOTMKTC1") & vbCrLf _
            & " group by SOTMKTC1.MARKET_CODE, SOTINVH2.ITEM_CODE" & vbCrLf _
            & ")) X, SOTMKTC1, ARTCUST1, SOTTCLS1" & vbCrLf _
            & " where SOTMKTC1.MARKET_CODE = X.MARKET_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTMKTC1.CUST_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & " group by X.MARKET_CODE, X.ITEM_CODE, ARTCUST1.PRICE_LIST_CODE, SOTTCLS1.CHANNEL_CODE, SOTMKTC1.CUST_CODE"

        DPTFCPR1 = ASCMAIN1.Temp_Table

        ASCMAIN1.sql = "Insert into " & DPTFCPR1 & " DPTFCPR1 (MARKET_CODE, ITEM_CODE, PRICE_LIST_CODE,CHANNEL_CODE,CUST_CODE)" & vbCrLf _
            & "Select SOTMKTC1.MARKET_CODE, SOTPRIC2.ITEM_CODE, SOTPRIC2.PRICE_LIST_CODE, SOTTCLS1.CHANNEL_CODE, SOTMKTC1.CUST_CODE" & vbCrLf _
            & " from SOTPRIC2,ARTCUST1,SOTMKTC1,ICTITEM1,SOTTCLS1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = SOTMKTC1.CUST_CODE" & vbCrLf _
            & "   and SOTPRIC2.PRICE_LIST_CODE = ARTCUST1.PRICE_LIST_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = SOTPRIC2.ITEM_CODE" & vbCrLf _
            & "   and ICTITEM1.ITEM_STATUS = 'A'" & vbCrLf _
            & Replace(SQLW, "DPTITMF1", "SOTMKTC1") & vbCrLf _
            & " minus" & vbCrLf _
            & "Select MARKET_CODE, ITEM_CODE, PRICE_LIST_CODE, CHANNEL_CODE, CUST_CODE from " & DPTFCPR1 & " DPTFCPR1"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into " & DPTFCPR1 & " DPTFCPR1 (MARKET_CODE, ITEM_CODE, PRICE_LIST_CODE,CHANNEL_CODE,CUST_CODE)" & vbCrLf _
            & "Select Distinct SOTMKTC1.MARKET_CODE, SOTINVH2.ITEM_CODE, ARTCUST1.PRICE_LIST_CODE, SOTTCLS1.CHANNEL_CODE, SOTMKTC1.CUST_CODE" & vbCrLf _
            & " from SOTINVH2,ARTCUST1,SOTMKTC1,ICTITEM1,SOTTCLS1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = SOTMKTC1.CUST_CODE" & vbCrLf _
            & "   and SOTINVH2.CUST_CODE = ARTCUST1.CUST_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED between '" & YYYY_ly & "01' and '" & YYYY_ly & "12'" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and ICTITEM1.ITEM_STATUS = 'A'" & vbCrLf _
            & Replace(SQLW, "DPTITMF1", "SOTMKTC1") & vbCrLf _
            & " minus" & vbCrLf _
            & "Select MARKET_CODE, ITEM_CODE, PRICE_LIST_CODE, CHANNEL_CODE, CUST_CODE from " & DPTFCPR1 & " DPTFCPR1"
        ASCDATA1.ExecuteSQL()


        Dim PRICES As String = ""
        Dim RETAILS As String = ""
        For M As Integer = 1 To 12
            PRICES &= ", CASE WHEN SOTPRIC2.ITEM_NEW_PRICE_DATE <= '" & Format(DATES(M).AddDays(1).AddMonths(-1), "dd-MMM-yyyy") & "' THEN SOTPRIC2.ITEM_NEW_PRICE ELSE SOTPRIC2.ITEM_PRICE END P" & Format(M, "00") & vbCrLf
            RETAILS &= ", CASE WHEN ICTITEM1.ITEM_NEW_RETAIL_PRICE_DATE <= '" & Format(DATES(M).AddDays(1).AddMonths(-1), "dd-MMM-yyyy") & "' THEN ICTITEM1.ITEM_NEW_RETAIL_PRICE ELSE ICTITEM1.ITEM_RETAIL_PRICE END R" & Format(M, "00") & vbCrLf
        Next

        ASCMAIN1.sql = "Select X.PRICE_LIST_CODE, X.MARKET_CODE, X.ITEM_CODE" & vbCrLf _
            & ", SOTMKTC1.CUST_CODE, ARTCUST1.PRICE_CLASS_CODE, SOTPRIC2.ITEM_PRICE" & vbCrLf _
            & ", SOTPCLS1.PRICE_BASIS, SOTPCLS1.PRICE_BASE_DPCT" & vbCrLf _
            & PRICES _
            & RETAILS _
            & " from " & DPTFCPR1 & " X, ARTCUST1, SOTPRIC2, SOTMKTC1, ICTITEM1, SOTPCLS1" & vbCrLf _
            & " where SOTMKTC1.MARKET_CODE = X.MARKET_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTMKTC1.CUST_CODE" & vbCrLf _
            & "   and SOTPRIC2.PRICE_LIST_CODE (+) = X.PRICE_LIST_CODE" & vbCrLf _
            & "   and SOTPRIC2.ITEM_CODE (+) = X.ITEM_CODE" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
            & "   and SOTPCLS1.PRICE_CLASS_CODE (+) = ARTCUST1.PRICE_CLASS_CODE"
        SOTPRICX = ASCMAIN1.Temp_Table

        ' 09/05/18 - disabling this constraint because HDQ/SLSRESERVE does not have a Price List code 
        '  and Yvonne/Lauren want to be able to run this report even if customer does not have a price list
        ' the report was designed for sales forecasts - and now we are using it for forecasted no charge shipments
        ' ASCDATA1.ExecuteSQL("Alter Table " & SOTPRICX & " Add Primary Key (PRICE_LIST_CODE,MARKET_CODE,ITEM_CODE)")

        For M As Integer = 1 To 12
            ASCMAIN1.sql = Replace("Update " & SOTPRICX & " Set PXX = RXX * (100 - PRICE_BASE_DPCT) / 100 where PRICE_BASIS = 'R' and PXX is Null", "XX", Format(M, "00"))
            ASCDATA1.ExecuteSQL()
        Next




        ' WHAT TO DO ABOUT OVERSHIPPED MONTHS

        Dim rTotal As Int64
        Dim r0 As Integer

        ASCMAIN1.Progress("-", "Data")

        ws = wb.Worksheets("Data")

        ASCMAIN1.sql = "Select " & vbCrLf _
            & "  X.MARKET_CODE" & vbCrLf _
            & ", '" & YYYY & "' YEAR" & vbCrLf _
            & ", X.CUST_CODE" & vbCrLf _
            & ", SOTCHAN1.CHANNEL_DESC" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", ICTCOLL0.HC_NAME" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_NAME" & vbCrLf _
            & ", ICTITEM1.PROD_CODE" & vbCrLf _
            & ", X.ITEM_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.ITEM_STATUS" & vbCrLf _
            & ", DECODE(SOTPRIC2.ITEM_CODE,NULL,'N','Y') ON_PRICE_LIST" & vbCrLf _
            & ", SOTPRIC2.ITEM_PRICE" & vbCrLf _
            & ", ICTITEM1.ITEM_COST_STD" & vbCrLf _
            & sqlAll & vbCrLf _
            & " from " & DPTFCPR1 & " X" & vbCrLf _
            & ",SOTMKTC1,ICTCOLL1,ICTITEM1," & SOTPRICX & " SOTPRICX,SOTCHAN1,ICTCOLL0,SOTPRIC2" & vbCrLf _
            & " where SOTMKTC1.MARKET_CODE = X.MARKET_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
            & "   and SOTCHAN1.CHANNEL_CODE (+) = X.CHANNEL_CODE" & vbCrLf _
            & "   and ICTCOLL0.HC_CODE (+) = ICTCOLL1.HC_CODE" & vbCrLf _
            & "   and SOTPRIC2.PRICE_LIST_CODE (+) = X.PRICE_LIST_CODE" & vbCrLf _
            & "   and SOTPRIC2.ITEM_CODE (+) = X.ITEM_CODE" & vbCrLf _
            & "   and SOTPRICX.PRICE_LIST_CODE (+) = X.PRICE_LIST_CODE" & vbCrLf _
            & "   and SOTPRICX.MARKET_CODE (+) = X.MARKET_CODE" & vbCrLf _
            & "   and SOTPRICX.ITEM_CODE (+) = X.ITEM_CODE"

        DataTable = ASCDATA1.GetDataTable

        r0 = 2 '4
        rTotal = DataTable.Select("").Length
        r = 0
        For Each row As DataRow In DataTable.Select("", "MARKET_CODE,BRAND_CODE,COLLECTION_NAME,ITEM_CODE")
            r += 1
            ws.Range("A" & CStr(r0 + r) & ":BU" & CStr(r0 + r)).Value2 = row.ItemArray
            If r Mod 100 = 0 Then
                ASCMAIN1.Progress("-", CStr(r) & "/" & CStr(rTotal))
            End If
        Next
        wb.Names.Add("PivotBase", "=Data!$A$" & CStr(r0) & ":$CD$" & CStr(r0 + DataTable.Rows.Count))

        ASCMAIN1.Progress("-", "Formulas")

        Dim CPM As Integer = 5 ' Cells per Month: Price, Units, Sales, CGS, Margin
        Dim HdgCols As Integer = 14 ' Number of columns of Item Data, before Jan

        Dim units_ytd As String = "" ' Formula for Units YTD

        For M As Integer = 1 To 12
            If M < CM Then
                units_ytd &= "+" & Excel_Cell(1, HdgCols + (M - 1) * CPM + 2)
            End If
            For T As Integer = 3 To CPM
                If T = CPM Or M >= CM Then

                    Dim J As Integer = HdgCols + (M - 1) * CPM + T

                    ' Copy the formulae on row 1 down to all of the rows
                    ' Only copy the GM% formula (ie, T=CPM) for actuals (ie, M < CM)
                    ' Copy all 3 formulae for Plans (ie, M >= CM)
                    Dim XC As String = Excel_Cell(0, J)
                    xlSourceRange = ws.Range(XC & CStr(1) & ":" & XC & CStr(1))
                    xlDestRange = ws.Range(XC & CStr(r0 + 1) & ":" & XC & CStr(r0 + DataTable.Rows.Count))
                    xlSourceRange.Copy(xlDestRange)
                End If
            Next
        Next

        Dim XCT As Integer = HdgCols + 12 * CPM ' Last col for Dec

        ' Set Units YTD, and then copy to Sales & CGS
        ws.Cells(1, XCT + 1).Formula = "=" & Mid(units_ytd, 2)
        xlSourceRange = ws.Range(Excel_Cell(1, XCT + 1) & ":" & Excel_Cell(1, XCT + 1))
        xlDestRange = ws.Range(Excel_Cell(1, XCT + 2) & ":" & Excel_Cell(1, XCT + 3))
        xlSourceRange.Copy(xlDestRange)

        ' Copy down 4 columns formulae for YTD and 4 columns for Total Year
        xlSourceRange = ws.Range(Excel_Cell(1, XCT + 1) & ":" & Excel_Cell(1, XCT + 4 + 4))
        xlDestRange = ws.Range(Excel_Cell(r0 + 1, XCT + 1) & ":" & Excel_Cell(r0 + DataTable.Rows.Count, XCT + 4 + 4))
        xlSourceRange.Copy(xlDestRange)

        ws.Cells(1, 1).Value = Now
        ws.Cells(1, 2).Value = "'" & YYYY

        ASCMAIN1.Progress("-", "Pivots")
        '   excel.Run("ResetData")

        Try
            wb.Worksheets("PIVOT").Activate()
            wb.ActiveSheet.PivotTables("PivotBase").PivotCache.Refresh()

        Catch ex As Exception

        End Try




        ASCMAIN1.Progress("Now Saving Workbook")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = "Monthly_Forecast_and_Actuals"
                XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsx"

                Dim objOpt As Object = Nothing ' Missing.Value
                wb.SaveAs(ASCMAIN1.Folders("Work") & XLS_FILENAME _
                          , Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook)
                wb.Close(False, objOpt, objOpt)

                success = True

            Catch ex As Exception
                If ASCMAIN1.Running_in_VS Then Stop
            End Try
        Loop

        excel.Quit()
        ws = Nothing
        wb = Nothing
        excel = Nothing
        xlSourceRange = Nothing
        xlDestRange = Nothing

        ReleaseCOMObject(xlDestRange)
        ReleaseCOMObject(xlSourceRange)
        ReleaseCOMObject(ws)
        ReleaseCOMObject(wb)
        ReleaseCOMObject(excel)

        Add_Document_to_ASTSPRF1(ASCMAIN1.Folders("Work") & XLS_FILENAME)

        ASCMAIN1.Progress("")

    End Sub

    Public Overrides Function Get_Custom_Filter_for_Codes_Selection(COLUMN_NAME As String) As String

        Select Case COLUMN_NAME
            Case "MARKET_CODE"
                Return "CUST_CODE is Not Null"

            Case Else
                Return MyBase.Get_Custom_Filter_for_Codes_Selection(COLUMN_NAME)
        End Select
    End Function

     
End Class