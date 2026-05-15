Imports System.IO
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports SpreadsheetGear

Public Class RSCXLSC1

    Public Shared Sub Update_RSTRETLx(
    ByVal EDI_DOC_SEQ_NO As String,
    Optional ByVal plus_or_minus As String = "")

        Dim S As Int32 = 1
        If plus_or_minus = "-" Then
            S = -1
        End If

        ASCMAIN1.Record_Event("EDT852T1", EDI_DOC_SEQ_NO, "", Now, ASCMAIN1.USER_ID, "852" & plus_or_minus, "852 Update", "")

        ASCDATA1.ExecuteSP("RSPRETLX", "VN", New Object() {EDI_DOC_SEQ_NO, S}, New String() {"EDI_DOC_SEQ_NO_IN", "S"})
    End Sub

    Public Shared Sub Import_Retail_Sales_For(frm As ASFBASE0, ByVal XLS_DOC_SEQ_NO As String)

        Dim rowRSTXLSQ1 As DataRow = frm.LookUp("RSTXLSQ1", XLS_DOC_SEQ_NO)
        Dim EDI_DOC_SEQ_NO As String = ""
        frm.dst.Tables("RSTXLSQE").Rows.Clear()
        frm.dst.Tables("RSTRETL1").Rows.Clear()
        frm.Fill_Records("ICTITEMU")

        Try

            If rowRSTXLSQ1 IsNot Nothing Then
                '\\ABSNASQ\Public\SLP\Share\SLP\XLS\InBound\0000000001-49041526390970070544200402524.csv
                Dim RS_PARM_XLS_FOLDER As String = frm.ROWs("RSTPARM1").Item("RS_PARM_XLS_FOLDER") & ""
                Dim CUST_CODE As String = rowRSTXLSQ1("CUST_CODE") & ""
                Dim XLS_DOC_FILENAME As String = $"{CUST_CODE}-{XLS_DOC_SEQ_NO}-{rowRSTXLSQ1("XLS_DOC_FILENAME")}"

                frm.Fill_Records("EDTUPCX1", CUST_CODE)
                Dim XLS_DATE_FROM As DateTime = rowRSTXLSQ1("XLS_DATE_FROM")
                Dim XLS_DATE_TO As DateTime = rowRSTXLSQ1("XLS_DATE_TO")
                Dim EDI_SOURCE As String = "XLS"
                Dim importFile As String = System.IO.Path.Combine(RS_PARM_XLS_FOLDER, XLS_DOC_FILENAME)
                If My.Computer.FileSystem.FileExists(importFile) Then
                    Dim rowEDT852T1 As DataRow = Create_EDT852T1(frm, CUST_CODE, XLS_DATE_FROM, XLS_DATE_TO, EDI_SOURCE)
                    EDI_DOC_SEQ_NO = rowEDT852T1("EDI_DOC_SEQ_NO")
                    Dim wb As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(importFile)
                    Dim ws As IWorksheet = wb.Worksheets(0)
                    Dim cells As IRange = ws.UsedRange

                    Select Case CUST_CODE
                        Case "REVOLVE"

                            Dim ITEM_CODE_col As Integer = ColumnLetterToIndex("C")
                            Dim QTY_SOLD_col As Integer = ColumnLetterToIndex("L")
                            Dim QTY_EOW_col As Integer = ColumnLetterToIndex("Q")
                            Dim AMT_SOLD_col As Integer = ColumnLetterToIndex("AC")
                            Dim STARTING_ROW As Integer = 1

                            For row As Integer = STARTING_ROW To cells.RowCount - 1
                                Dim rowRSTRETL1 As DataRow = frm.dst.Tables("RSTRETL1").NewRow
                                rowRSTRETL1("EDI_DOC_SEQ_NO") = rowEDT852T1("EDI_DOC_SEQ_NO")
                                rowRSTRETL1("CUST_CODE") = CUST_CODE
                                rowRSTRETL1("CUST_STORE_NO") = "000000"
                                rowRSTRETL1("ITEM_CODE") = cells(row, ITEM_CODE_col).Value
                                rowRSTRETL1("QTY_SOLD") = cells(row, QTY_SOLD_col).Value
                                rowRSTRETL1("AMT_SOLD") = cells(row, AMT_SOLD_col).Value
                                rowRSTRETL1("OPS_YYYYPP") = rowEDT852T1("OPS_YYYYPP")
                                rowRSTRETL1("OPS_YYYYWW") = rowEDT852T1("OPS_YYYYWW")
                                rowRSTRETL1("QTY_EOW") = cells(row, QTY_EOW_col).Value
                                frm.dst.Tables("RSTRETL1").Rows.Add(rowRSTRETL1)
                            Next

                        Case "ANTHRO"
                            If wb.Worksheets.Count > 1 Then
                                For Each sheet As IWorksheet In wb.Worksheets
                                    If sheet.Name.Equals("Total", StringComparison.OrdinalIgnoreCase) Then
                                        ws = sheet
                                        Exit For
                                    End If
                                Next
                            End If

                            Dim ITEM_CODE_col As Integer = ColumnLetterToIndex("D")

                            Dim QTY_SOLD_RTL_col As Integer = ColumnLetterToIndex("Q")
                            Dim AMT_SOLD_RTL_col As Integer = ColumnLetterToIndex("P")
                            Dim QTY_EOW_RTL_col As Integer = ColumnLetterToIndex("R")

                            Dim QTY_SOLD_ECOM_col As Integer = ColumnLetterToIndex("X")
                            Dim AMT_SOLD_ECOM_col As Integer = ColumnLetterToIndex("W")
                            Dim QTY_EOW_ECOM_col As Integer = ColumnLetterToIndex("Y")

                            Dim STARTING_ROW As Integer = 7

                            For row As Integer = STARTING_ROW To cells.RowCount - 1

                                Dim ITEM_CODE As String = Validate_Item_Code(frm, cells(row, ITEM_CODE_col).Value, CUST_CODE)

                                If ITEM_CODE <> "" AndAlso ITEM_CODE <> "IGNORE" Then
                                    Dim QTY_SOLD As Int32 = Val(cells(row, QTY_SOLD_RTL_col).Value & "")
                                    Dim AMT_SOLD As Int32 = Val(cells(row, AMT_SOLD_RTL_col).Value & "")
                                    Dim QTY_EOW As Int32 = Val(cells(row, QTY_EOW_RTL_col).Value & "")
                                    If QTY_SOLD <> 0 Or AMT_SOLD <> 0 Or QTY_EOW <> 0 Then
                                        Dim rowRSTRETL1 As DataRow = frm.dst.Tables("RSTRETL1").Rows.Find(New String() {
                                            CStr(rowEDT852T1("EDI_DOC_SEQ_NO")), CUST_CODE, "000002", ITEM_CODE,
                                            CStr(rowEDT852T1("OPS_YYYYPP")), CStr(rowEDT852T1("OPS_YYYYWW"))
                                        })
                                        If rowRSTRETL1 Is Nothing Then
                                            rowRSTRETL1 = frm.dst.Tables("RSTRETL1").NewRow()
                                            rowRSTRETL1("EDI_DOC_SEQ_NO") = rowEDT852T1("EDI_DOC_SEQ_NO")
                                            rowRSTRETL1("CUST_CODE") = CUST_CODE
                                            rowRSTRETL1("CUST_STORE_NO") = "000002"
                                            rowRSTRETL1("ITEM_CODE") = ITEM_CODE
                                            rowRSTRETL1("OPS_YYYYPP") = rowEDT852T1("OPS_YYYYPP")
                                            rowRSTRETL1("OPS_YYYYWW") = rowEDT852T1("OPS_YYYYWW")
                                            frm.dst.Tables("RSTRETL1").Rows.Add(rowRSTRETL1)
                                        End If
                                        rowRSTRETL1("QTY_SOLD") = Val(rowRSTRETL1("QTY_SOLD") & "") + QTY_SOLD
                                        rowRSTRETL1("AMT_SOLD") = Val(rowRSTRETL1("AMT_SOLD") & "") + AMT_SOLD
                                        rowRSTRETL1("QTY_EOW") = Val(rowRSTRETL1("QTY_EOW") & "") + QTY_EOW
                                    End If

                                    Dim QTY_SOLD_ECOM As Int32 = Val(cells(row, QTY_SOLD_ECOM_col).Value & "")
                                    Dim AMT_SOLD_ECOM As Int32 = Val(cells(row, AMT_SOLD_ECOM_col).Value & "")
                                    Dim QTY_EOW_ECOM As Int32 = Val(cells(row, QTY_EOW_ECOM_col).Value & "")
                                    If QTY_SOLD_ECOM <> 0 Or AMT_SOLD_ECOM <> 0 Or QTY_EOW_ECOM <> 0 Then

                                        Dim rowRSTRETL1 As DataRow = frm.dst.Tables("RSTRETL1").Rows.Find(New String() {
                                        CStr(rowEDT852T1("EDI_DOC_SEQ_NO")), CUST_CODE, "000000", ITEM_CODE,
                                        CStr(rowEDT852T1("OPS_YYYYPP")), CStr(rowEDT852T1("OPS_YYYYWW"))
                                    })
                                        If rowRSTRETL1 Is Nothing Then
                                            rowRSTRETL1 = frm.dst.Tables("RSTRETL1").NewRow()
                                            rowRSTRETL1("EDI_DOC_SEQ_NO") = rowEDT852T1("EDI_DOC_SEQ_NO")
                                            rowRSTRETL1("CUST_CODE") = CUST_CODE
                                            rowRSTRETL1("CUST_STORE_NO") = "000000"
                                            rowRSTRETL1("ITEM_CODE") = ITEM_CODE
                                            rowRSTRETL1("OPS_YYYYPP") = rowEDT852T1("OPS_YYYYPP")
                                            rowRSTRETL1("OPS_YYYYWW") = rowEDT852T1("OPS_YYYYWW")
                                            frm.dst.Tables("RSTRETL1").Rows.Add(rowRSTRETL1)
                                        End If
                                        rowRSTRETL1("QTY_SOLD") = Val(rowRSTRETL1("QTY_SOLD") & "") + QTY_SOLD_ECOM
                                        rowRSTRETL1("AMT_SOLD") = Val(rowRSTRETL1("AMT_SOLD") & "") + AMT_SOLD_ECOM
                                        rowRSTRETL1("QTY_EOW") = Val(rowRSTRETL1("QTY_EOW") & "") + QTY_EOW_ECOM

                                    End If
                                End If

                            Next

                        Case "CONTAINER"
                            Dim STARTING_ROW As Integer = 4
                            Dim HEADER_ROW As Integer = STARTING_ROW - 2
                            Dim COL_ITEM As Integer = ColumnLetterToIndex("C")
                            Dim COL_QTY As Integer = Find_Header_Column(cells, HEADER_ROW, {"Sales Qty WTD"})
                            Dim COL_AMT As Integer = Find_Header_Column(cells, HEADER_ROW, {"Sales $ WTD"})
                            If COL_QTY < 0 OrElse COL_AMT < 0 Then Throw New ApplicationException("Could not locate 'Sales Qty WTD' or 'Sales $ WTD'.")

                            For r As Integer = STARTING_ROW To cells.RowCount - 1
                                'Dim vpn As String = Trim(CStr(cells(r, COL_ITEM).Text))
                                Dim vpn As String = Try_Read_UPC(cells(r, COL_ITEM))
                                If vpn = "" Then Continue For
                                Dim item As String = Validate_Item_Code(frm, vpn, CUST_CODE)
                                If item = "" OrElse item = "IGNORE" Then Continue For

                                Dim qty As Integer = ToInt(CStr(cells(r, COL_QTY).Text))
                                Dim amt As Decimal = ToDec(CStr(cells(r, COL_AMT).Text))
                                If qty = 0 AndAlso amt = 0D Then Continue For
                                Add_Qtys_And_Amts(frm, rowEDT852T1, CUST_CODE, "000000", item, qty, amt, DBNull.Value)
                            Next

                        Case "HUTGROUP"

                            Dim HR As Integer = 0

                            Dim COL_WEND = Find_Header_Column(cells, HR, {"Week_End_Date", "Week End Date", "end_date", "Week End"})
                            Dim COL_BAR = Find_Header_Column(cells, HR, {"Barcode (UPC)", "Barcode", "UPC"})
                            Dim COL_AMT = Find_Header_Column(cells, HR, {"Revenue_USD", "Revenue USD", "Revenue", "Sales $", "Sales_USD"})
                            Dim COL_QTY = Find_Header_Column(cells, HR, {"Sale_Volume_Units", "Sale Volume Units", "Units", "Qty Sold", "Qty" & ChrW(160) & "Sold", "Quantity"})

                            If {COL_WEND, COL_BAR, COL_AMT, COL_QTY}.Any(Function(i) i < 0) Then
                                Throw New ApplicationException("HUTGROUP: missing one or more headers (week end, barcode, revenue, units).")
                            End If

                            Dim yMap As New Dictionary(Of Date, (YPP As String, YWW As String))()
                            Dim existingOrAdded As New HashSet(Of String)(StringComparer.Ordinal)

                            For r As Integer = HR + 1 To cells.RowCount - 1
                                Dim upc As String = Try_Read_UPC(cells(r, COL_BAR)) : If upc = "" Then Continue For
                                Dim item As String = Validate_Item_Code(frm, upc, CUST_CODE)
                                If item = "" OrElse item = "IGNORE" Then Continue For

                                Dim rawWendStr As String = Trim(CStr(cells(r, COL_WEND).Text)) : If rawWendStr = "" Then Continue For

                                Dim rawWend As Date
                                If Not Date.TryParse(rawWendStr, rawWend) Then
                                    Dim v = cells(r, COL_WEND).Value
                                    If v Is Nothing OrElse Not Date.TryParse(CStr(v), rawWend) Then Continue For
                                End If

                                Dim satWend As Date = Prev_Saturday(rawWend)
                                If Not yMap.ContainsKey(satWend) Then
                                    Dim yinfo = Get_Week_And_Year(satWend)
                                    If yinfo Is Nothing Then Continue For
                                    yMap(satWend) = yinfo.Value
                                End If
                                Dim ypp = yMap(satWend).YPP
                                Dim yww = yMap(satWend).YWW

                                If Exists_In_DB_Or_Added(frm, CUST_CODE, "000000", item, ypp, yww, existingOrAdded) Then Continue For

                                Dim qty As Integer = CellInt(cells(r, COL_QTY))
                                Dim amt As Decimal = CellDec(cells(r, COL_AMT))
                                If qty = 0 AndAlso amt = 0D Then Continue For

                                Add_Qtys_And_Amts(frm, rowEDT852T1, CUST_CODE, "000000", item, qty, amt, DBNull.Value, ypp, yww)

                                existingOrAdded.Add(item & "|" & ypp & "|" & yww)
                            Next


                        Case "SEPHORACA"
                            Dim ws852 = Get_Sheet(wb, "852 Data")
                            If ws852 Is Nothing Then Throw New ApplicationException("Worksheet '852 Data' not found.")
                            Dim c852 As IRange = ws852.UsedRange

                            Dim HR As Integer = 0
                            Dim COL_STORE As Integer = ColumnLetterToIndex("D")
                            Dim COL_ITEM As Integer = ColumnLetterToIndex("N")
                            Dim COL_EOW As Integer = ColumnLetterToIndex("Q")

                            Dim COL_Q As Integer = Find_Header_Column(c852, HR, {"Qty Sold", "Qty Sold (Net)", "Qty Sold (N)", "Qty" & ChrW(160) & "Sold"})
                            Dim COL_A As Integer = Find_Header_Column(c852, HR, {"$ Sold", "$" & ChrW(160) & "Sold", "Sales $", "Sales $" & ChrW(160) & "(Net)", "$ Sold (Net)"})
                            If COL_Q < 0 Then COL_Q = ColumnLetterToIndex("AB")
                            If COL_A < 0 Then COL_A = ColumnLetterToIndex("BC")

                            For r As Integer = 1 To c852.RowCount - 1
                                Dim storeRaw As String = Trim(CStr(c852(r, COL_STORE).Text))
                                If storeRaw = "" Then Continue For
                                Dim store As String = storeRaw.PadLeft(6, "0"c)

                                Dim upc As String = Try_Read_UPC(c852(r, COL_ITEM))
                                If upc = "" Then Continue For
                                Dim item As String = Validate_Item_Code(frm, upc, CUST_CODE)
                                If item = "" OrElse item = "IGNORE" Then Continue For

                                Dim qty = ToInt(CStr(c852(r, COL_Q).Text))
                                Dim amt = ToDec(CStr(c852(r, COL_A).Text))
                                Dim eow = ToInt(CStr(c852(r, COL_EOW).Text))
                                Add_Qtys_And_Amts(frm, rowEDT852T1, CUST_CODE, store, item, qty, amt, eow)
                            Next

                        Case "VONMAUR"
                            Dim COL_STYLE = ColumnLetterToIndex("C")
                            Dim COL_COLOR = ColumnLetterToIndex("D")
                            Dim COL_UNITS_REG = ColumnLetterToIndex("F")
                            Dim COL_UNITS_MD = ColumnLetterToIndex("G")
                            Dim COL_DOLLARS_REG = ColumnLetterToIndex("I")
                            Dim COL_DOLLARS_MD = ColumnLetterToIndex("J")
                            Dim COL_QTY_EOW = ColumnLetterToIndex("L")

                            Dim hasMarkdownHeader As Boolean = False
                            For rr = 0 To cells.RowCount - 1
                                For cc = 0 To cells.ColumnCount - 1
                                    If Trim(CStr(cells(rr, cc).Text)).Equals("ITEMS AT MARKDOWN", StringComparison.OrdinalIgnoreCase) Then
                                        hasMarkdownHeader = True : Exit For
                                    End If
                                Next
                                If hasMarkdownHeader Then Exit For
                            Next

                            For r As Integer = 10 To cells.RowCount - 1
                                Dim style As String = Trim(CStr(cells(r, COL_STYLE).Text))
                                If style = "" OrElse
                           style.Equals("STYLE", StringComparison.OrdinalIgnoreCase) OrElse
                           style.Equals("TOTALS", StringComparison.OrdinalIgnoreCase) OrElse
                           style.StartsWith("DEPT. SELL", StringComparison.OrdinalIgnoreCase) Then Continue For

                                If style.StartsWith("VENDOR TOTALS", StringComparison.OrdinalIgnoreCase) Then
                                    If Not hasMarkdownHeader Then Exit For Else Continue For
                                End If

                                Dim color As String = Trim(CStr(cells(r, COL_COLOR).Text))
                                Dim rawKey As String = If(color = "", style, $"{style}-{color}")

                                Dim item As String = Validate_Item_Code(frm, rawKey, CUST_CODE)
                                If item = "" Then item = rawKey
                                If item = "" OrElse item = "IGNORE" Then Continue For

                                Dim uReg = ToInt(CStr(cells(r, COL_UNITS_REG).Text))
                                Dim uMd = ToInt(CStr(cells(r, COL_UNITS_MD).Text))
                                Dim dReg = ToDec(CStr(cells(r, COL_DOLLARS_REG).Text))
                                Dim dMd = ToDec(CStr(cells(r, COL_DOLLARS_MD).Text))
                                Dim eow = ToInt(CStr(cells(r, COL_QTY_EOW).Text))

                                Dim qty As Integer, amt As Decimal
                                If uReg <> 0 OrElse dReg <> 0D Then
                                    qty = uReg : amt = dReg
                                ElseIf uMd <> 0 OrElse dMd <> 0D Then
                                    qty = uMd : amt = dMd
                                Else
                                    Continue For
                                End If

                                Add_Qtys_And_Amts(frm, rowEDT852T1, CUST_CODE, "000000", item, qty, amt, eow)
                            Next

                        Case "HOLT"
                            Import_HOLT(frm, wb, rowEDT852T1, CUST_CODE)

                        Case "SHOPBOP"
                            Dim COL_FLAG = ColumnLetterToIndex("C")
                            Dim COL_ITEM = ColumnLetterToIndex("F")
                            Dim COL_UNITS = ColumnLetterToIndex("N")
                            Dim COL_DOLLARS = ColumnLetterToIndex("O")
                            Dim COL_EOW = ColumnLetterToIndex("P")

                            Dim COL_COLOR = ColumnLetterToIndex("E")

                            For r As Integer = 0 To cells.RowCount - 1
                                Dim marker As String = Trim(CStr(cells(r, COL_FLAG).Text))
                                If marker = "" OrElse Not marker.StartsWith("SLIPP", StringComparison.OrdinalIgnoreCase) Then Continue For

                                Dim rawItem As String = Trim(CStr(cells(r, COL_ITEM).Text))
                                If rawItem = "" Then Continue For

                                Dim fixedItem As String = rawItem
                                Dim noHyphen As String = rawItem.Replace("-", "").Trim()
                                Dim isMaskPlaceholder As Boolean =
                                rawItem.Equals("SILK-MASK", StringComparison.OrdinalIgnoreCase) OrElse
                                rawItem.Equals("SLK-MASK", StringComparison.OrdinalIgnoreCase) OrElse
                                noHyphen.Equals("SILKMASK", StringComparison.OrdinalIgnoreCase)

                                If isMaskPlaceholder AndAlso COL_COLOR >= 0 Then
                                    Dim color As String = (cells(r, COL_COLOR).Text & "").Trim()
                                    If color.StartsWith("caramel", StringComparison.OrdinalIgnoreCase) Then
                                        fixedItem = "853218006117"
                                    ElseIf color.StartsWith("charcoal", StringComparison.OrdinalIgnoreCase) Then
                                        fixedItem = "853218006124"
                                    End If
                                End If

                                Dim item As String = Validate_Item_Code(frm, fixedItem, CUST_CODE)
                                If item = "" OrElse item = "IGNORE" Then Continue For

                                Dim qty = ToInt(CStr(cells(r, COL_UNITS).Text))
                                Dim amt = ToDec(CStr(cells(r, COL_DOLLARS).Text))
                                Dim eow = ToInt(CStr(cells(r, COL_EOW).Text))
                                Add_Qtys_And_Amts(frm, rowEDT852T1, CUST_CODE, "000000", item, qty, amt, eow)
                            Next


                        Case "SDM"
                            Dim wsSDM = Get_Sheet_Like(wb, "SKU")
                            If wsSDM Is Nothing Then Throw New ApplicationException("Worksheet with 'SKU' in the name was not found.")
                            Dim cSDM As IRange = wsSDM.UsedRange

                            ' UPC column is fixed
                            Dim COL_ITEM As Integer = ColumnLetterToIndex("D")

                            Dim salesPos = Find_Header_ColAndRow_Anywhere(cSDM, {"Sales", "Sales $", "Sales CAD", "Sales $ CAD", "Sales USD"})
                            Dim unitsPos = Find_Header_ColAndRow_Anywhere(cSDM, {"Units", "Units Wk", "Units Qty"})

                            If salesPos.Col < 0 Then Throw New ApplicationException("SDM: Could not locate the Sales column header.")
                            If unitsPos.Col < 0 Then Throw New ApplicationException("SDM: Could not locate the Units column header.")

                            Dim HEADER_ROW As Integer = Math.Max(salesPos.Row, unitsPos.Row)

                            Dim COL_AMT As Integer = salesPos.Col
                            Dim COL_QTY As Integer = unitsPos.Col

                            Dim COL_LABEL As Integer = ColumnLetterToIndex("B")
                            If COL_LABEL >= cSDM.ColumnCount Then COL_LABEL = ColumnLetterToIndex("A")

                            For r As Integer = HEADER_ROW + 1 To cSDM.RowCount - 1
                                ' stop on grand total row
                                Dim label As String = Trim(CStr(cSDM(r, COL_LABEL).Text))
                                If label.Equals("Total", StringComparison.OrdinalIgnoreCase) Then Exit For

                                Dim upc As String = Try_Read_UPC(cSDM(r, COL_ITEM))
                                If upc = "" Then Continue For

                                Dim item As String = Validate_Item_Code(frm, upc, CUST_CODE)
                                If item = "" OrElse item = "IGNORE" Then Continue For

                                Dim amt As Decimal = CellDec(cSDM(r, COL_AMT))
                                Dim qty As Integer = CellInt(cSDM(r, COL_QTY))

                                ' July Pt.2, merge into existing period row instead of adding a new one
                                Add_Qtys_And_Amts_SDM_Merge(frm, rowEDT852T1, CUST_CODE, item, qty, amt)
                            Next


                        Case "GOOP"

                            For Each wsYr As IWorksheet In wb.Worksheets
                                Dim yrName As String = wsYr.Name.Trim()
                                Dim sheetYear As Integer
                                If Not Integer.TryParse(yrName, sheetYear) Then Continue For

                                Dim c As IRange = wsYr.UsedRange

                                Dim posUPC = Find_Header_ColAndRow_Anywhere(c, {"UPC", "UPCS"})
                                If posUPC.Col < 0 Then Continue For

                                Dim posUnitsGT = Find_Header_ColAndRow_Anywhere(c, {"Grand Total"})
                                Dim posUnitsEC = Find_Header_ColAndRow_Anywhere(c, {"E-Commerce", "E-Commerce"})
                                Dim posUnitsRT = Find_Header_ColAndRow_Anywhere(c, {"Retail"})

                                Dim posAmt = Find_Header_ColAndRow_Anywhere(c, {
                                    "Total Sales", "GRAND TTL SLS", "GRAND TOTAL SALES", "Sales $", "Sales USD", "Total $"
                                })
                                If posAmt.Col < 0 Then
                                    Continue For
                                End If

                                Dim COL_MONTHLBL As Integer = ColumnLetterToIndex("A")
                                Dim COL_UPC As Integer = posUPC.Col
                                Dim COL_UNITS_GT As Integer = posUnitsGT.Col
                                Dim COL_UNITS_EC As Integer = posUnitsEC.Col
                                Dim COL_UNITS_RT As Integer = posUnitsRT.Col
                                Dim COL_AMT As Integer = posAmt.Col

                                Dim headerRow As Integer = Math.Max(Math.Max(posUPC.Row, posAmt.Row),
                                                            Math.Max(Math.Max(posUnitsGT.Row, posUnitsEC.Row), posUnitsRT.Row))

                                Dim curYPP As String = Nothing
                                Dim curYWW As String = Nothing
                                Dim curMM As Integer = 0
                                Dim curYY As Integer = sheetYear

                                For r As Integer = 0 To c.RowCount - 1

                                    Dim aTxt As String = Trim(CStr(c(r, COL_MONTHLBL).Text))
                                    If aTxt <> "" Then
                                        Dim parts = aTxt.Replace("/", "-").Replace(".", "-").Trim().
                                                        Split(New Char() {"-"c, " "c}, StringSplitOptions.RemoveEmptyEntries)

                                        If parts.Length >= 1 Then
                                            Dim probe As DateTime
                                            If DateTime.TryParseExact(parts(0),
                                                New String() {"MMM", "MMMM", "MMM.", "MMMM."},
                                                CultureInfo.InvariantCulture,
                                                DateTimeStyles.None, probe) Then

                                                Dim mmTmp As Integer = probe.Month
                                                Dim yyTmp As Integer = curYY
                                                If parts.Length >= 2 Then
                                                    Dim yyRaw As Integer
                                                    If Integer.TryParse(parts(1), yyRaw) Then
                                                        yyTmp = If(yyRaw < 100, 2000 + yyRaw, yyRaw)
                                                    End If
                                                End If

                                                Dim YYYYPP As String = $"{yyTmp:0000}{mmTmp:00}"
                                                Dim rowGL As DataRow = ASCDATA1.GetDataRow(
                                                    "Select MAX(WEEK_END_DATE) DATE_TO, MIN(YYYYPP) YYYYPP, MIN(YYYYWW) YYYYWW " &
                                                    "from GLTPARM3 where YYYYPP = :PARM1",
                                                    "V", New Object() {YYYYPP})

                                                If rowGL IsNot Nothing AndAlso Not rowGL.IsNull("DATE_TO") Then
                                                    curYPP = CStr(rowGL("YYYYPP"))
                                                    curYWW = CStr(rowGL("YYYYWW"))
                                                    curMM = mmTmp : curYY = yyTmp
                                                Else
                                                    curYPP = Nothing : curYWW = Nothing : curMM = 0
                                                End If

                                                Continue For
                                            End If
                                        End If

                                        If aTxt.StartsWith("Grand Total", StringComparison.OrdinalIgnoreCase) OrElse
                                           aTxt.StartsWith("Total", StringComparison.OrdinalIgnoreCase) Then
                                            curMM = 0 : curYPP = Nothing : curYWW = Nothing
                                            Continue For
                                        End If
                                    End If

                                    If r <= headerRow Then Continue For
                                    If curMM = 0 OrElse String.IsNullOrEmpty(curYPP) Then Continue For

                                    Dim upc As String = Try_Read_UPC(c(r, COL_UPC))
                                    If upc = "" Then Continue For

                                    Dim item As String = Validate_Item_Code(frm, upc, CUST_CODE)
                                    If item = "" OrElse item = "IGNORE" Then Continue For

                                    Dim qty As Integer = 0
                                    If COL_UNITS_GT >= 0 Then
                                        qty = CellInt(c(r, COL_UNITS_GT))
                                    Else
                                        Dim ec As Integer = If(COL_UNITS_EC >= 0, CellInt(c(r, COL_UNITS_EC)), 0)
                                        Dim rt As Integer = If(COL_UNITS_RT >= 0, CellInt(c(r, COL_UNITS_RT)), 0)
                                        qty = ec + rt
                                    End If

                                    Dim amt As Decimal = CellDec(c(r, COL_AMT))

                                    Add_Qtys_And_Amts(frm, rowEDT852T1, CUST_CODE, "000000", item, qty, amt, DBNull.Value, curYPP, curYWW)

                                Next
                            Next

                        Case "EQUINOX"
                            ' --- does this workbook have the April dual-tab layout?
                            Dim hasStores As Boolean = False, hasEcomm As Boolean = False
                            For Each s As IWorksheet In wb.Worksheets
                                If s.Name.Equals("Stores", StringComparison.OrdinalIgnoreCase) Then hasStores = True
                                If s.Name.Equals("ECOMM", StringComparison.OrdinalIgnoreCase) Then hasEcomm = True
                            Next

                            ' =============== STORES (tab) ===============
                            If hasStores Then
                                Dim wsStores As IWorksheet = Get_Sheet(wb, "Stores")
                                Dim cs As IRange = wsStores.UsedRange

                                ' UPC column is now provided as "UPCS" (sometimes "UPC")
                                Dim posUPC = Find_Header_ColAndRow_Anywhere(cs, {"UPCS", "UPC", "UPC(s)", "UPC #"})
                                ' Legacy description/color (fallback)
                                Dim posDesc = Find_Header_ColAndRow_Anywhere(cs, {"Description", "Descriptor", "Descriptor/Color"})
                                Dim posColor = Find_Header_ColAndRow_Anywhere_Exact(cs, {"Color", "Colour"})

                                ' Units / Dollars columns (keep legacy sniffing too)
                                Dim posU = Find_Header_ColAndRow_Anywhere(cs, {"Sum of MTD U", "Sum of MTD Units", "MTD Units", "MTD U", "Units"})
                                Dim posS = Find_Header_ColAndRow_Anywhere(cs, {"Sum of MTD $", "Sum of MTD Dollars", "MTD $", "MTD Dollars", "Sum of MTD Sales", "MTD Sales", "Sales $"})

                                Dim COL_UNITS As Integer = If(posU.Col >= 0, posU.Col, ColumnLetterToIndex("T"))
                                Dim COL_SALES As Integer = If(posS.Col >= 0, posS.Col, ColumnLetterToIndex("U"))

                                If (posUPC.Col < 0) AndAlso (posDesc.Col < 0 OrElse posColor.Col < 0) Then
                                    Throw New ApplicationException("EQUINOX (Stores): could not locate UPC/Description/Color headers.")
                                End If

                                Dim hdrRow As Integer = Math.Max(Math.Max(posUPC.Row, posDesc.Row), posColor.Row)

                                For r As Integer = hdrRow + 1 To cs.RowCount - 1
                                    Dim itemKey As String = ""
                                    ' 1) Prefer true UPC
                                    If posUPC.Col >= 0 Then
                                        itemKey = Try_Read_UPC(cs(r, posUPC.Col))
                                    End If
                                    ' 2) Fallback to legacy Description|Color
                                    If String.IsNullOrWhiteSpace(itemKey) AndAlso posDesc.Col >= 0 AndAlso posColor.Col >= 0 Then
                                        Dim descRaw As String = Normalize_Spaces(CStr(cs(r, posDesc.Col).Text))
                                        Dim colorRaw As String = Normalize_Spaces(CStr(cs(r, posColor.Col).Text))
                                        itemKey = Build_Equinox_Store_Key(descRaw, colorRaw)
                                    End If
                                    If String.IsNullOrWhiteSpace(itemKey) Then Continue For

                                    Dim item As String = Validate_Item_Code(frm, itemKey, CUST_CODE)
                                    If item = "" OrElse item = "IGNORE" Then Continue For

                                    Dim qty As Integer = CellInt(cs(r, COL_UNITS))
                                    Dim amt As Decimal = CellDec(cs(r, COL_SALES))
                                    If qty = 0 AndAlso amt = 0D Then Continue For

                                    Add_Qtys_And_Amts(frm, rowEDT852T1, CUST_CODE, "000000", item, qty, amt, DBNull.Value)
                                Next
                            End If

                            ' =============== ECOMM (tab) ===============
                            If hasEcomm Then
                                Dim wsEC As IWorksheet = Get_Sheet(wb, "ECOMM")
                                Dim ce As IRange = wsEC.UsedRange

                                ' New format has "UPC" at far right and "Net Items"/"Total sales"
                                Dim posUPC = Find_Header_ColAndRow_Anywhere(ce, {"UPC", "UPCS", "UPC(s)", "UPC #"})
                                Dim posUnits = Find_Header_ColAndRow_Anywhere(ce, {"Net Items", "net_quantity", "Net items", "Units", "Qty"})
                                Dim posSales = Find_Header_ColAndRow_Anywhere(ce, {"Total sales", "total_sales", "Sales $", "Net Sales", "Gross Sales"})
                                ' Fallback keys if UPC missing (older Shopify export)
                                Dim posProd = Find_Header_ColAndRow_Anywhere(ce, {"product_title", "Product title at time of sale"})
                                Dim posVar = Find_Header_ColAndRow_Anywhere(ce, {"Product variant SKU", "Variant", "Variant Title", "Variant_Title", "Option", "Colour/Size", "Color/Size", "Style/Size"})

                                If posUnits.Col < 0 OrElse posSales.Col < 0 Then
                                    Throw New ApplicationException("EQUINOX (ECOMM): missing Units/Sales headers.")
                                End If

                                Dim hdrRow As Integer = Math.Max(Math.Max(posUPC.Row, posUnits.Row), posSales.Row)

                                For r As Integer = hdrRow + 1 To ce.RowCount - 1
                                    Dim itemKey As String = ""
                                    ' 1) Prefer UPC
                                    If posUPC.Col >= 0 Then
                                        itemKey = Try_Read_UPC(ce(r, posUPC.Col))
                                    End If
                                    ' 2) Fallback to Product|Variant
                                    If String.IsNullOrWhiteSpace(itemKey) AndAlso (posProd.Col >= 0 OrElse posVar.Col >= 0) Then
                                        Dim prod As String = If(posProd.Col >= 0, Trim(CStr(ce(r, posProd.Col).Text)), "")
                                        Dim varT As String = If(posVar.Col >= 0, Trim(CStr(ce(r, posVar.Col).Text)), "")
                                        itemKey = Build_Equinox_Ecom_Key(prod, varT)
                                    End If
                                    If String.IsNullOrWhiteSpace(itemKey) Then Continue For

                                    Dim item As String = Validate_Item_Code(frm, itemKey, CUST_CODE)
                                    If item = "" OrElse item = "IGNORE" Then Continue For

                                    Dim qty As Integer = CellInt(ce(r, posUnits.Col))
                                    Dim amt As Decimal = CellDec(ce(r, posSales.Col))
                                    Dim eow As Integer = 0 ' no EOW on this sheet (keep as 0)
                                    If qty = 0 AndAlso amt = 0D AndAlso eow = 0 Then Continue For

                                    Add_Qtys_And_Amts(frm, rowEDT852T1, CUST_CODE, "ECOM", item, qty, amt, eow)
                                Next

                                Exit Select ' handled both tabs
                            End If

                            ' ================= SINGLE-SHEET FALLBACKS =================
                            Dim origName As String = rowRSTXLSQ1("XLS_DOC_FILENAME") & ""
                            Dim wsEq As IWorksheet = wb.Worksheets(0)
                            Dim c As IRange = wsEq.UsedRange

                            If Not IsEquinoxEcommFile(origName) Then
                                ' ---- Single-sheet STORES ----
                                Dim posUPC = Find_Header_ColAndRow_Anywhere(c, {"UPCS", "UPC", "UPC(s)", "UPC #"})
                                Dim posDesc = Find_Header_ColAndRow_Anywhere(c, {"Description", "Descriptor", "Descriptor/Color"})
                                Dim posColor = Find_Header_ColAndRow_Anywhere_Exact(c, {"Color", "Colour"})

                                Dim posU = Find_Header_ColAndRow_Anywhere(c, {"Sum of MTD U", "Sum of MTD Units", "MTD Units", "MTD U", "Units"})
                                Dim posS = Find_Header_ColAndRow_Anywhere(c, {"Sum of MTD $", "Sum of MTD Dollars", "MTD $", "MTD Dollars", "Sum of MTD Sales", "MTD Sales", "Sales $"})

                                Dim COL_UNITS2 As Integer = If(posU.Col >= 0, posU.Col, ColumnLetterToIndex("T"))
                                Dim COL_SALES2 As Integer = If(posS.Col >= 0, posS.Col, ColumnLetterToIndex("U"))

                                If (posUPC.Col < 0) AndAlso (posDesc.Col < 0 OrElse posColor.Col < 0) Then
                                    Throw New ApplicationException("EQUINOX (Stores): could not locate UPC/Description/Color headers.")
                                End If

                                Dim headerRow As Integer = Math.Max(Math.Max(posUPC.Row, posDesc.Row), posColor.Row)

                                For r As Integer = headerRow + 1 To c.RowCount - 1
                                    Dim key As String = ""
                                    If posUPC.Col >= 0 Then key = Try_Read_UPC(c(r, posUPC.Col))
                                    If String.IsNullOrWhiteSpace(key) AndAlso posDesc.Col >= 0 AndAlso posColor.Col >= 0 Then
                                        key = Build_Equinox_Store_Key(Normalize_Spaces(CStr(c(r, posDesc.Col).Text)),
                                              Normalize_Spaces(CStr(c(r, posColor.Col).Text)))
                                    End If
                                    If key = "" Then Continue For

                                    Dim item As String = Validate_Item_Code(frm, key, CUST_CODE)
                                    If item = "" OrElse item = "IGNORE" Then Continue For

                                    Dim qty As Integer = CellInt(c(r, COL_UNITS2))
                                    Dim amt As Decimal = CellDec(c(r, COL_SALES2))
                                    If qty = 0 AndAlso amt = 0D Then Continue For

                                    Add_Qtys_And_Amts(frm, rowEDT852T1, CUST_CODE, "000000", item, qty, amt, DBNull.Value)
                                Next

                                Exit Select
                            End If

                            ' ---- Single-sheet ECOMM ----
                            Dim posUPC1 = Find_Header_ColAndRow_Anywhere(c, {"UPC", "UPCS", "UPC(s)", "UPC #"})
                            Dim posUnits1 = Find_Header_ColAndRow_Anywhere(c, {"Net Items", "net_quantity", "Net items", "Units", "Qty Sold", "Quantity"})
                            Dim posSales1 = Find_Header_ColAndRow_Anywhere(c, {"Total sales", "total_sales", "Sales $", "Net Sales", "Gross Sales"})
                            Dim posProd1 = Find_Header_ColAndRow_Anywhere(c, {"product_title", "Product title at time of sale"})
                            Dim posVar1 = Find_Header_ColAndRow_Anywhere(c, {"Product variant SKU", "Variant", "Variant Title", "Variant_Title", "Option", "Colour/Size", "Color/Size", "Style/Size"})
                            If posUnits1.Col < 0 OrElse posSales1.Col < 0 Then
                                Throw New ApplicationException("EQUINOX (ECOM): could not locate Units/Sales headers.")
                            End If

                            Dim headerRowN As Integer = Math.Max(Math.Max(posUPC1.Row, posUnits1.Row), posSales1.Row)

                            For r As Integer = headerRowN + 1 To c.RowCount - 1
                                Dim key As String = ""
                                If posUPC1.Col >= 0 Then key = Try_Read_UPC(c(r, posUPC1.Col))
                                If String.IsNullOrWhiteSpace(key) AndAlso (posProd1.Col >= 0 OrElse posVar1.Col >= 0) Then
                                    Dim prod As String = If(posProd1.Col >= 0, Trim(CStr(c(r, posProd1.Col).Text)), "")
                                    Dim varT As String = If(posVar1.Col >= 0, Trim(CStr(c(r, posVar1.Col).Text)), "")
                                    key = Build_Equinox_Ecom_Key(prod, varT)
                                End If
                                If key = "" Then Continue For

                                Dim item As String = Validate_Item_Code(frm, key, CUST_CODE)
                                If item = "" OrElse item = "IGNORE" Then Continue For

                                Dim qty As Integer = CellInt(c(r, posUnits1.Col))
                                Dim amt As Decimal = CellDec(c(r, posSales1.Col))
                                Dim eow As Integer = 0
                                If qty = 0 AndAlso amt = 0D AndAlso eow = 0 Then Continue For

                                Add_Qtys_And_Amts(frm, rowEDT852T1, CUST_CODE, "ECOM", item, qty, amt, eow)
                            Next


                        Case "BEAUTYSPAC"

                            ' --- Format A: “Vendor Report …” → read tab "SKU Report"
                            If XLS_DOC_FILENAME.IndexOf("Vendor Report", StringComparison.OrdinalIgnoreCase) >= 0 Then
                                Dim wsSKU As IWorksheet = Nothing
                                For Each s As IWorksheet In wb.Worksheets
                                    If s.Name.Equals("SKU Report", StringComparison.OrdinalIgnoreCase) Then
                                        wsSKU = s : Exit For
                                    End If
                                Next
                                If wsSKU Is Nothing Then wsSKU = wb.Worksheets(0)

                                Dim c As IRange = wsSKU.UsedRange

                                Dim COL_UPC As Integer = ColumnLetterToIndex("B")   ' Item Barcode
                                Dim COL_AMT As Integer = ColumnLetterToIndex("I")   ' $ Sold 
                                Dim COL_QTY As Integer = ColumnLetterToIndex("L")   ' Units Sold 

                                Dim firstDataRow As Integer = 0
                                For r As Integer = 0 To Math.Min(15, c.RowCount - 1)
                                    Dim upcProbe As String = Try_Read_UPC(c(r, COL_UPC))
                                    Dim hdrA As String = Trim(CStr(c(r, 0).Text))
                                    If Not String.IsNullOrWhiteSpace(upcProbe) AndAlso Not upcProbe.Equals("Item Barcode", StringComparison.OrdinalIgnoreCase) AndAlso Not hdrA.Equals("TOTAL", StringComparison.OrdinalIgnoreCase) Then
                                        firstDataRow = r : Exit For
                                    End If
                                Next

                                For r As Integer = firstDataRow To c.RowCount - 1
                                    Dim aTxt As String = Trim(CStr(c(r, 0).Text))
                                    If aTxt.Equals("TOTAL", StringComparison.OrdinalIgnoreCase) Then Continue For

                                    Dim upc As String = Try_Read_UPC(c(r, COL_UPC))
                                    If String.IsNullOrWhiteSpace(upc) OrElse upc.Equals("Item Barcode", StringComparison.OrdinalIgnoreCase) Then Continue For

                                    Dim item As String = Validate_Item_Code(frm, upc, CUST_CODE)
                                    If item = "" OrElse item = "IGNORE" Then Continue For

                                    Dim amt As Decimal = CellDec(c(r, COL_AMT))
                                    Dim qty As Integer = CellInt(c(r, COL_QTY))

                                    If qty = 0 AndAlso amt = 0D Then Continue For

                                    Add_Qtys_And_Amts(frm, rowEDT852T1, CUST_CODE, "000000", item, qty, amt, DBNull.Value)
                                Next
                                Exit Select
                            End If

                            ' --- Format B: filenames containing "SKUReport" → single sheet
                            If XLS_DOC_FILENAME.IndexOf("SKUREPORT", StringComparison.OrdinalIgnoreCase) >= 0 Then
                                Dim wsSKU As IWorksheet = Get_Sheet_Like(wb, "SKU")
                                If wsSKU Is Nothing Then wsSKU = wb.Worksheets(0)
                                Dim c As IRange = wsSKU.UsedRange

                                Dim startCol As Integer = c.Column
                                Dim COL_UPC As Integer = ColumnLetterToIndex("D") - startCol   ' UPC
                                Dim COL_AMT As Integer = ColumnLetterToIndex("M") - startCol   ' TW $ Sales
                                Dim COL_QTY As Integer = ColumnLetterToIndex("Q") - startCol   ' TW Units

                                Dim firstDataRow As Integer = 0
                                For r As Integer = 0 To Math.Min(15, c.RowCount - 1)
                                    Dim upcProbe As String = Try_Read_UPC(c(r, COL_UPC))
                                    If Not String.IsNullOrWhiteSpace(upcProbe) AndAlso
                                       Not upcProbe.Equals("UPC", StringComparison.OrdinalIgnoreCase) AndAlso
                                       Not upcProbe.Equals("Item Barcode", StringComparison.OrdinalIgnoreCase) Then
                                        firstDataRow = r : Exit For
                                    End If
                                Next

                                For r As Integer = firstDataRow To c.RowCount - 1
                                    Dim aTxt As String = Trim(CStr(c(r, 0).Text))
                                    If aTxt.Equals("TOTAL", StringComparison.OrdinalIgnoreCase) Then Continue For

                                    Dim upc As String = Try_Read_UPC(c(r, COL_UPC))
                                    If String.IsNullOrWhiteSpace(upc) OrElse upc.Equals("UPC", StringComparison.OrdinalIgnoreCase) _
                                       OrElse upc.Equals("Item Barcode", StringComparison.OrdinalIgnoreCase) Then Continue For

                                    Dim item As String = Validate_Item_Code(frm, upc, CUST_CODE)
                                    If item = "" OrElse item = "IGNORE" Then Continue For

                                    Dim amt As Decimal = CellDec(c(r, COL_AMT))
                                    Dim qty As Integer = CellInt(c(r, COL_QTY))
                                    If qty = 0 AndAlso amt = 0D Then Continue For

                                    Add_Qtys_And_Amts(frm, rowEDT852T1, CUST_CODE, "000000", item, qty, amt, DBNull.Value)
                                Next

                                Exit Select
                            End If

                        Case "MODA"
                            Dim HDR As Integer = 0
                            Dim COL_UPC As Integer = Find_Header_Column(cells, HDR, {"VID", "UPC", "Barcode", "SKU"})
                            Dim COL_QTY As Integer = Find_Header_Column(cells, HDR, {"Net Unit Sales", "Net Units", "Units Sold", "Units"})
                            Dim COL_AMT As Integer = Find_Header_Column(cells, HDR, {"Net $ Sales", "Net Sales $", "Sales $", "$ Sales"})
                            Dim COL_EOW As Integer = Find_Header_Column(cells, HDR, {"Units On Hand", "On Hand", "EOW", "OH Units"})

                            If COL_UPC < 0 Then COL_UPC = ColumnLetterToIndex("B")
                            If COL_QTY < 0 Then COL_QTY = ColumnLetterToIndex("U")
                            If COL_AMT < 0 Then COL_AMT = ColumnLetterToIndex("V")
                            If COL_EOW < 0 Then COL_EOW = ColumnLetterToIndex("N")

                            For r As Integer = HDR + 1 To cells.RowCount - 1
                                Dim upc As String = Try_Read_UPC(cells(r, COL_UPC))
                                If String.IsNullOrWhiteSpace(upc) Then Continue For

                                Dim item As String = Validate_Item_Code(frm, upc, CUST_CODE)
                                If item = "" OrElse item = "IGNORE" Then Continue For

                                Dim qty As Integer = CellInt(cells(r, COL_QTY))
                                Dim amt As Decimal = CellDec(cells(r, COL_AMT))
                                Dim eow As Integer = CellInt(cells(r, COL_EOW))

                                If qty = 0 AndAlso amt = 0D AndAlso eow = 0 Then Continue For

                                Add_Qtys_And_Amts(frm, rowEDT852T1, CUST_CODE, "000000", item, qty, amt, eow)
                            Next

                    End Select

                End If

                If frm.dst.Tables("RSTXLSQE").Rows.Count > 0 Then
                    Dim sqlEDTTRPM1_check As String = "Select * FROM EDTTRPM1 WHERE EDI_TP_QUAL = :PARM1 AND EDI_TP_ID = :PARM2 AND EDI_DOC_NO = :PARM3"
                    Dim rowEDTTRPM1 As DataRow = ASCDATA1.GetDataRow(sqlEDTTRPM1_check, "VVV", New Object() {"ZZ", CUST_CODE, "852"})

                    If IsNothing(rowEDTTRPM1) Then
                        Dim sqlEDTTRPM1_ins As String = $"INSERT INTO EDTTRPM1 (EDI_TP_QUAL, EDI_TP_ID, EDI_DOC_NO, CUST_CODE, EDI_STATUS) VALUES ('ZZ', '{CUST_CODE}', '852', '{CUST_CODE}', 'P')"
                        ASCDATA1.ExecuteSQL(sqlEDTTRPM1_ins)
                    End If

                    frm.Update_Record_TDA("EDTUPCX1")
                    Using frmmsg As New ASFMSGBF
                        frmmsg.Show_grd(frm.dst.Tables("RSTXLSQE"), frm, "Invalid UPC Codes")
                    End Using
                Else

                    frm.BeginTrans()

                    ASCMAIN1.sql = $"Update RSTXLSQ1 SET XLS_DOC_STATUS = '1', LAST_DATE = SYSDATE, LAST_OPER = '{ASCMAIN1.USER_ID}', EDI_DOC_SEQ_NO = '{EDI_DOC_SEQ_NO}' WHERE XLS_DOC_SEQ_NO = :PARM1"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {XLS_DOC_SEQ_NO})

                    frm.Update_Record_TDA("RSTRETL1")
                    frm.Update_Record_TDA("EDT852T1")
                    Update_RSTRETLx(EDI_DOC_SEQ_NO)

                    frm.CommitTrans($"XLS Doc: {XLS_DOC_SEQ_NO} imported.")

                    System.IO.File.Move(importFile, importFile.Replace("InBound", "Inbound_Archive"))
                End If


            End If

        Catch ex As Exception
            ASCMAIN1.Record_Event("RSTXLSQ1", XLS_DOC_SEQ_NO, "", Now, ASCMAIN1.USER_ID, "XLSE", "XLS Import Error", "")
            ASCMAIN1.sql = "Update RSTXLSQ1 SET XLS_DOC_STATUS = 'E' WHERE XLS_DOC_SEQ_NO = :PARM1"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {XLS_DOC_SEQ_NO})
            Stop
        End Try
    End Sub

    Public Shared Function Create_EDT852T1(frm As ASFBASE0, ByVal CUST_CODE As String, EDI_FROM_DATE As Date, EDI_TO_DATE As Date, EDI_SOURCE As String) As DataRow

        Dim EDI_DOC_SEQ_NO As String = ASCMAIN1.Next_Control_No("EDTJRNL3.EDI_DOC_SEQ_NO")

        Dim GEN_DOC_NO As String = ""
        Dim EDI_ISA_NO As String = ""
        Dim EDI_TP_QUAL As String = "" 'rowEDTTRPM1.Item("EDI_TP_QUAL") & ""
        Dim EDI_TP_ID As String = "" 'rowEDTTRPM1.Item("EDI_TP_ID") & ""
        Dim EDI_OUR_QUAL As String = "" 'rowEDTTRPM1.Item("EDI_OUR_QUAL") & ""
        Dim EDI_OUR_ID As String = "" 'rowEDTTRPM1.Item("EDI_OUR_ID") & ""

        Dim rowEDT852T1 As DataRow = frm.dst.Tables("EDT852T1").NewRow
        rowEDT852T1.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
        rowEDT852T1.Item("GEN_DOC_NO") = GEN_DOC_NO
        rowEDT852T1.Item("EDI_ISA_NO") = EDI_ISA_NO
        rowEDT852T1.Item("EDI_TP_QUAL") = EDI_TP_QUAL
        rowEDT852T1.Item("EDI_TP_ID") = EDI_TP_ID
        rowEDT852T1.Item("EDI_OUR_QUAL") = EDI_OUR_QUAL
        rowEDT852T1.Item("EDI_OUR_ID") = EDI_OUR_ID
        rowEDT852T1.Item("EDI_FROM_DATE") = EDI_FROM_DATE
        rowEDT852T1.Item("EDI_TO_DATE") = EDI_TO_DATE

        rowEDT852T1.Item("EDI_STATUS") = "1"

        Dim sqlGLTPARM3 As String = "Select MIN(YYYYPP) YYYYPP, MIN(YYYYWW) YYYYWW FROM GLTPARM3 WHERE WEEK_END_DATE >= :PARM1"
        Dim rowGLTPARM3 As DataRow = ASCDATA1.GetDataRow(sqlGLTPARM3, "D", New Object() {EDI_FROM_DATE})

        rowEDT852T1.Item("OPS_YYYYPP") = rowGLTPARM3.Item("YYYYPP")
        rowEDT852T1.Item("OPS_YYYYWW") = rowGLTPARM3.Item("YYYYWW")
        rowEDT852T1.Item("CUST_CODE") = CUST_CODE
        rowEDT852T1.Item("INIT_DATE") = Now
        rowEDT852T1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowEDT852T1.Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
        rowEDT852T1.Item("EDI_SOURCE") = EDI_SOURCE
        frm.dst.Tables("EDT852T1").Rows.Add(rowEDT852T1)
        Return rowEDT852T1
    End Function

    Public Shared Function ColumnLetterToIndex(columnLetter As String, Optional zeroBased As Boolean = True) As Integer
        Dim col As Integer = 0
        columnLetter = columnLetter.ToUpperInvariant()

        For i As Integer = 0 To columnLetter.Length - 1
            col = col * 26 + (Asc(columnLetter(i)) - Asc("A"c) + 1)
        Next

        If zeroBased Then
            Return col - 1 ' SpreadsheetGear / DataTable style
        Else
            Return col     ' Excel 1-based style
        End If
    End Function

    Public Shared Function BuildFileImportList(CUST_CODE As String, directory As String) As List(Of FileDateInfo)
        Dim results As New List(Of FileDateInfo)()
        If Not System.IO.Directory.Exists(directory) Then Return results

        Select Case CUST_CODE.ToUpperInvariant()
            Case "MODA"

                Dim rx As New Regex("^\s*(?<m>\d{1,2})[.\-](?<d>\d{1,2})(?:[.\-](?<y>\d{2,4}))?\s*(?:by[\s\-_]*item)?",
                                    RegexOptions.IgnoreCase)

                Dim seenBase As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                For Each pattern In New String() {"*.csv", "*.xlsx", "*.xls"}
                    For Each filePath In System.IO.Directory.EnumerateFiles(directory, pattern)
                        Dim fileName = Path.GetFileName(filePath)
                        If fileName.StartsWith("~$") Then Continue For

                        Dim baseName = Path.GetFileNameWithoutExtension(fileName)
                        If seenBase.Contains(baseName) Then Continue For

                        Dim m = rx.Match(baseName)
                        If Not m.Success Then Continue For

                        Dim mm As Integer, dd As Integer, yy As Integer

                        If Not Integer.TryParse(m.Groups("m").Value, mm) Then Continue For
                        If Not Integer.TryParse(m.Groups("d").Value, dd) Then Continue For
                        If mm < 1 OrElse mm > 12 OrElse dd < 1 OrElse dd > 31 Then Continue For

                        If m.Groups("y").Success AndAlso m.Groups("y").Value <> "" Then
                            If Not Integer.TryParse(m.Groups("y").Value, yy) Then Continue For
                            If yy < 100 Then yy += 2000
                        Else
                            yy = 2024
                        End If

                        results.Add(New FileDateInfo With {
                            .FileName = fileName,
                            .Month = mm,
                            .Day = dd,
                            .Year = yy
                        })
                        seenBase.Add(baseName)
                    Next
                Next

            Case "BEAUTYSPAC"
                Dim seenBase As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                Dim rxLeading As New Regex("^\s*(?<m>\d{1,2})\.(?<d>\d{1,2})(?:\.(?<y>\d{2,4}))?\s*", RegexOptions.IgnoreCase)

                Dim rxSlipTail As New Regex("SLIP[-_\s]*(?<m>\d{1,2})\.(?<d>\d{1,2})(?:\.(?<y>\d{2,4}))?\b",
                                RegexOptions.IgnoreCase)

                For Each pattern In New String() {"*.xlsx", "*.xls"}
                    For Each filePath In System.IO.Directory.EnumerateFiles(directory, pattern)
                        Dim fileName = Path.GetFileName(filePath)
                        If fileName.StartsWith("~$") Then Continue For

                        Dim baseName = Path.GetFileNameWithoutExtension(fileName)
                        If seenBase.Contains(baseName) Then Continue For

                        Dim mm As Integer = 0, dd As Integer = 0, yy As Integer = 0
                        Dim got As Boolean = False

                        Dim mLead = rxLeading.Match(fileName)
                        If mLead.Success Then
                            If Integer.TryParse(mLead.Groups("m").Value, mm) AndAlso
                   Integer.TryParse(mLead.Groups("d").Value, dd) Then

                                If mLead.Groups("y").Success AndAlso mLead.Groups("y").Value <> "" Then
                                    If Integer.TryParse(mLead.Groups("y").Value, yy) Then
                                        If yy < 100 Then yy += 2000
                                    End If
                                Else
                                    yy = 2023
                                End If
                                got = (mm >= 1 AndAlso mm <= 12 AndAlso dd >= 1 AndAlso dd <= 31 AndAlso yy > 0)
                            End If
                        End If

                        If Not got Then
                            Dim mSlip = rxSlipTail.Match(fileName)
                            If mSlip.Success Then
                                If Integer.TryParse(mSlip.Groups("m").Value, mm) AndAlso
                       Integer.TryParse(mSlip.Groups("d").Value, dd) Then

                                    If mSlip.Groups("y").Success AndAlso mSlip.Groups("y").Value <> "" Then
                                        If Integer.TryParse(mSlip.Groups("y").Value, yy) Then
                                            If yy < 100 Then yy += 2000
                                        End If
                                    Else
                                        yy = 2023
                                    End If
                                    got = (mm >= 1 AndAlso mm <= 12 AndAlso dd >= 1 AndAlso dd <= 31 AndAlso yy > 0)
                                End If
                            End If
                        End If

                        If got Then
                            results.Add(New FileDateInfo With {.FileName = fileName, .Month = mm, .Day = dd, .Year = yy})
                            seenBase.Add(baseName)
                        End If
                    Next
                Next

            Case "GOOP"
                For Each filePath In System.IO.Directory.EnumerateFiles(directory, "*.xlsx")
                    Dim fileName = Path.GetFileName(filePath)
                    If fileName.StartsWith("~$") Then Continue For

                    Dim mYR = Regex.Matches(fileName, "(?<y>\d{4})")
                    Dim yy As Integer = If(mYR.Count > 0, Integer.Parse(mYR(mYR.Count - 1).Groups("y").Value), Date.Now.Year)

                    results.Add(New FileDateInfo With {.FileName = fileName, .Month = 12, .Day = 0, .Year = yy})
                Next

            Case "SDM"

                Dim rxMonthYearLoose As New Regex(
                    "^\s*(?<mon>[A-Za-z]{3,9})\s+(?<yr>\d{2,4})\b",
                    RegexOptions.IgnoreCase)

                Dim rxPnum As New Regex("(?<!\d)P\s*(?<m>\d{1,2})\D+(?<y>\d{2,4})(?!\d)", RegexOptions.IgnoreCase)
                Dim rxNumYear As New Regex("(?<!\d)(?<m>\d{1,2})\s+(?<y>\d{2,4})(?!\d)", RegexOptions.IgnoreCase)

                Dim seenBase As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                For Each filePath In System.IO.Directory.EnumerateFiles(directory, "*.xlsx")
                    Dim fileName = Path.GetFileName(filePath)
                    If fileName.StartsWith("~$") Then Continue For

                    Dim baseName = Path.GetFileNameWithoutExtension(fileName)
                    If seenBase.Contains(baseName) Then Continue For

                    Dim mm As Integer = 0, yy As Integer = 0
                    Dim ok As Boolean = False

                    Dim m0 = rxMonthYearLoose.Match(baseName)
                    If m0.Success Then
                        Dim monStr = m0.Groups("mon").Value
                        If monStr.Equals("SEPT", StringComparison.OrdinalIgnoreCase) Then monStr = "Sep"
                        Dim dt As DateTime
                        If DateTime.TryParseExact(monStr,
                                                  New String() {"MMM", "MMMM", "MMM.", "MMMM.", "Sep", "Sept", "Sept."},
                                                  CultureInfo.InvariantCulture, DateTimeStyles.None, dt) Then
                            mm = dt.Month
                            yy = CInt(m0.Groups("yr").Value)
                            If yy < 100 Then yy += 2000
                            ok = (mm >= 1 AndAlso mm <= 12)
                        End If
                    End If

                    If ok Then
                        results.Add(New FileDateInfo With {.FileName = fileName, .Month = mm, .Year = yy, .Day = 0})
                        seenBase.Add(baseName)
                    End If
                Next

            Case "SHOPBOP"
                Dim rx As New Regex("^\s*(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?\b", RegexOptions.IgnoreCase)

                Dim seenBase As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                For Each pattern In New String() {"*.xlsx", "*.xls"}
                    For Each filePath In System.IO.Directory.EnumerateFiles(directory, pattern)
                        Dim fileName = Path.GetFileName(filePath)
                        If fileName.StartsWith("~$") Then Continue For

                        Dim baseName = Path.GetFileNameWithoutExtension(fileName)
                        If seenBase.Contains(baseName) Then Continue For

                        Dim m = rx.Match(fileName)
                        If Not m.Success Then Continue For

                        Dim mm As Integer, dd As Integer, yy As Integer

                        If Not Integer.TryParse(m.Groups(1).Value, mm) Then Continue For
                        If Not Integer.TryParse(m.Groups(2).Value, dd) Then Continue For
                        If mm < 1 OrElse mm > 12 OrElse dd < 1 OrElse dd > 31 Then Continue For

                        If m.Groups(3).Success AndAlso m.Groups(3).Value <> "" Then
                            If Not Integer.TryParse(m.Groups(3).Value, yy) Then Continue For
                            If yy < 100 Then yy += 2000
                        Else
                            yy = 2024
                        End If

                        results.Add(New FileDateInfo With {.FileName = fileName, .Month = mm, .Day = dd, .Year = yy})
                        seenBase.Add(baseName)
                    Next
                Next


            Case "HOLT"
                Dim rxRange As New Regex("(\d{1,2})[.\-/](\d{1,2})(?:[.\-/](\d{2,4}))?\s*[-–]\s*(\d{1,2})[.\-/](\d{1,2})(?:[.\-/](\d{2,4}))?",
                             RegexOptions.IgnoreCase)

                Dim FixYear As Func(Of Integer?, Integer, Integer) =
                Function(optYY As Integer?, guessYr As Integer) As Integer
                    If optYY.HasValue Then
                        Dim y = optYY.Value
                        If y < 100 Then y += 2000
                        Return y
                    End If
                    Return guessYr
                End Function

                Dim seenBase As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each pattern In New String() {"*.xlsx", "*.csv"}
                    For Each filePath In System.IO.Directory.EnumerateFiles(directory, pattern)
                        Dim fileName = Path.GetFileName(filePath)
                        If fileName.StartsWith("~$") Then Continue For

                        Dim baseName = Path.GetFileNameWithoutExtension(fileName)
                        If seenBase.Contains(baseName) Then Continue For

                        Dim m = rxRange.Match(fileName)
                        If Not m.Success Then Continue For

                        ' start date parts
                        Dim sm As Integer = Integer.Parse(m.Groups(1).Value)
                        Dim sd As Integer = Integer.Parse(m.Groups(2).Value)
                        Dim sYOpt As Integer? = If(m.Groups(3).Success AndAlso m.Groups(3).Value <> "", Integer.Parse(m.Groups(3).Value), CType(Nothing, Integer?))

                        ' end date parts (WEEK END)
                        Dim em As Integer = Integer.Parse(m.Groups(4).Value)
                        Dim ed As Integer = Integer.Parse(m.Groups(5).Value)
                        Dim eYOpt As Integer? = If(m.Groups(6).Success AndAlso m.Groups(6).Value <> "", Integer.Parse(m.Groups(6).Value), CType(Nothing, Integer?))

                        Dim ts As Date = System.IO.File.GetLastWriteTime(filePath)
                        Dim refYear As Integer = ts.Year

                        Dim endYear As Integer
                        Dim startYear As Integer

                        If eYOpt.HasValue Then
                            endYear = FixYear(eYOpt, refYear)
                            startYear = If(sYOpt.HasValue, FixYear(sYOpt, refYear), endYear)
                            If Not sYOpt.HasValue AndAlso em < sm Then startYear -= 1
                        ElseIf sYOpt.HasValue Then
                            startYear = FixYear(sYOpt, refYear)
                            endYear = startYear
                            If em < sm Then endYear += 1
                        Else
                            Dim candidates = New Integer() {refYear - 1, refYear, refYear + 1}
                            Dim best As Integer = refYear : Dim bestDiff As Double = Double.MaxValue
                            For Each y In candidates
                                Dim tryDate As Date
                                If Date.TryParse($"{y}-{em:00}-{ed:00}", tryDate) Then
                                    Dim d = Math.Abs((tryDate - ts).TotalDays)
                                    If d < bestDiff Then bestDiff = d : best = y
                                End If
                            Next
                            endYear = best
                            startYear = endYear
                            If em < sm Then startYear -= 1
                        End If

                        Dim endDt As Date
                        If Not Date.TryParse($"{endYear}-{em:00}-{ed:00}", endDt) Then Continue For

                        results.Add(New FileDateInfo With {.FileName = fileName, .Month = em, .Day = ed, .Year = endYear})
                        seenBase.Add(baseName)
                    Next
                Next


            Case "HUTGROUP"

                Dim seenBase As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim rxRange As New Regex("Week\s*Ending\s*(\d{1,2})\.(\d{1,2})\.(\d{2,4})\s*-\s*(\d{1,2})\.(\d{1,2})\.(\d{2,4})",
                                         RegexOptions.IgnoreCase)
                Dim rxSimple As New Regex("^\s*(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?\b",
                                          RegexOptions.IgnoreCase)

                For Each pattern In New String() {"*.xlsx", "*.csv"}
                    For Each filePath In System.IO.Directory.EnumerateFiles(directory, pattern)
                        Dim fileName = Path.GetFileName(filePath)
                        If fileName.StartsWith("~$") Then Continue For

                        Dim baseName = Path.GetFileNameWithoutExtension(fileName)
                        If seenBase.Contains(baseName) Then Continue For

                        Dim mR = rxRange.Match(fileName)
                        If mR.Success Then
                            Dim mm As Integer = Integer.Parse(mR.Groups(4).Value)
                            Dim dd As Integer = Integer.Parse(mR.Groups(5).Value)
                            Dim yy As Integer = Integer.Parse(mR.Groups(6).Value)
                            If yy < 100 Then yy += 2000
                            results.Add(New FileDateInfo With {.FileName = fileName, .Month = mm, .Day = dd, .Year = yy})
                            seenBase.Add(baseName)
                            Continue For
                        End If

                        Dim mS = rxSimple.Match(fileName)
                        If mS.Success Then
                            Dim mm As Integer = Integer.Parse(mS.Groups(1).Value)
                            Dim dd As Integer = Integer.Parse(mS.Groups(2).Value)
                            Dim yy As Integer = If(mS.Groups(3).Success AndAlso mS.Groups(3).Value <> "",
                                                   Integer.Parse(mS.Groups(3).Value), 2024)
                            If yy < 100 Then yy += 2000
                            results.Add(New FileDateInfo With {.FileName = fileName, .Month = mm, .Day = dd, .Year = yy})
                            seenBase.Add(baseName)
                        End If
                    Next
                Next

            Case "EQUINOX"
                Dim rxEq As New Regex("^\s*(?:Equinox_)?SLIP\s+([A-Za-z]{3,9})\s+(\d{2,4})\b",
                          RegexOptions.IgnoreCase)

                For Each filePath In System.IO.Directory.EnumerateFiles(directory, "*.xlsx")
                    Dim fileName = Path.GetFileName(filePath)
                    If fileName.StartsWith("~$") Then Continue For

                    Dim m = rxEq.Match(fileName)
                    If Not m.Success Then Continue For

                    Dim monthStr As String = m.Groups(1).Value
                    If monthStr.Equals("SEPT", StringComparison.OrdinalIgnoreCase) Then
                        monthStr = "Sep"
                    End If

                    Dim dt As DateTime
                    If Not DateTime.TryParseExact(monthStr,
            New String() {"MMM", "MMMM", "MMM.", "MMMM.", "Sep", "Sept", "Sept.", "SEPT", "SEPTEMBER"},
            CultureInfo.InvariantCulture, DateTimeStyles.None, dt) Then
                        Continue For
                    End If

                    Dim yy As Integer
                    If Not Integer.TryParse(m.Groups(2).Value, yy) Then Continue For
                    If yy < 100 Then yy += 2000

                    results.Add(New FileDateInfo With {.FileName = fileName, .Month = dt.Month, .Year = yy})
                Next

            Case "SEPHORACA"
                Dim rx As New Regex("Week\s+of\s+(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{2,4})",
                                    RegexOptions.IgnoreCase)

                For Each filePath In System.IO.Directory.EnumerateFiles(directory, "*.xlsx")
                    Dim fileName = Path.GetFileName(filePath)
                    If fileName.StartsWith("~$") Then Continue For

                    Dim m = rx.Match(fileName)
                    If Not m.Success Then Continue For

                    Dim mm As Integer, dd As Integer, yy As Integer
                    If Not Integer.TryParse(m.Groups(1).Value, mm) Then Continue For
                    If Not Integer.TryParse(m.Groups(2).Value, dd) Then Continue For
                    If Not Integer.TryParse(m.Groups(3).Value, yy) Then Continue For
                    If yy < 100 Then yy += 2000

                    If mm < 1 OrElse mm > 12 OrElse dd < 1 OrElse dd > 31 Then Continue For

                    results.Add(New FileDateInfo With {
                        .FileName = fileName,
                        .Month = mm,
                        .Day = dd,
                        .Year = yy
                    })
                Next
            Case "CONTAINER"
                Dim rxDot As New Regex("^\s*(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?\b", RegexOptions.IgnoreCase)

                For Each filePath In System.IO.Directory.EnumerateFiles(directory, "*.xlsx")
                    Dim fileName = Path.GetFileName(filePath)
                    If fileName.StartsWith("~$") Then Continue For   ' skip temp files

                    Dim m = rxDot.Match(fileName)
                    If Not m.Success Then Continue For

                    Dim month As Integer, day As Integer, year As Integer

                    If Not Integer.TryParse(m.Groups(1).Value, month) Then Continue For
                    If Not Integer.TryParse(m.Groups(2).Value, day) Then Continue For

                    If m.Groups(3).Success AndAlso m.Groups(3).Value <> "" Then
                        If Not Integer.TryParse(m.Groups(3).Value, year) Then Continue For
                        If year < 100 Then year += 2000
                    Else
                        year = 2024
                    End If

                    If month < 1 OrElse month > 12 OrElse day < 1 OrElse day > 31 Then Continue For

                    results.Add(New FileDateInfo With {.FileName = fileName, .Month = month, .Day = day, .Year = year})
                Next


            Case "VONMAUR"
                Dim rxWeekOf As New Regex("Week\s+of\s+(\d{1,2})[.\-\/](\d{1,2})(?:[.\-\/](\d{2,4}))?",
                              RegexOptions.IgnoreCase)

                Dim rxWeekEnding As New Regex("week\s+ending\s+(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{2,4})",
                                  RegexOptions.IgnoreCase)
                Dim rxMonthWk As New Regex("^\s*([A-Za-z]{3,9})\s+Wk\.?\s*(\d+)\b",
                           RegexOptions.IgnoreCase)

                For Each filePath In System.IO.Directory.EnumerateFiles(directory, "*.xlsx")
                    Dim fileName = Path.GetFileName(filePath)
                    If fileName.StartsWith("~$") Then Continue For

                    Dim mEnd = rxWeekEnding.Match(fileName)
                    If mEnd.Success Then
                        Dim mm As Integer, dd As Integer, yy As Integer
                        If Integer.TryParse(mEnd.Groups(1).Value, mm) AndAlso
                           Integer.TryParse(mEnd.Groups(2).Value, dd) AndAlso
                           Integer.TryParse(mEnd.Groups(3).Value, yy) Then

                            If yy < 100 Then yy += 2000
                            If mm >= 1 AndAlso mm <= 12 AndAlso dd >= 1 AndAlso dd <= 31 Then
                                results.Add(New FileDateInfo With {.FileName = fileName, .Month = mm, .Day = dd, .Year = yy})
                            End If
                        End If
                        Continue For
                    End If

                    Dim mOf = rxWeekOf.Match(fileName)
                    If mOf.Success Then
                        Dim mm As Integer, dd As Integer
                        If Not Integer.TryParse(mOf.Groups(1).Value, mm) Then Continue For
                        If Not Integer.TryParse(mOf.Groups(2).Value, dd) Then Continue For
                        If mm < 1 OrElse mm > 12 OrElse dd < 1 OrElse dd > 31 Then Continue For

                        Dim yy As Integer
                        If mOf.Groups(3).Success AndAlso mOf.Groups(3).Value <> "" Then
                            If Not Integer.TryParse(mOf.Groups(3).Value, yy) Then Continue For
                            If yy < 100 Then yy += 2000
                        Else
                            Dim ts = System.IO.File.GetLastWriteTime(filePath)
                            Dim candidates = New Integer() {ts.Year - 1, ts.Year, ts.Year + 1}
                            yy = candidates _
                    .OrderBy(Function(y) Math.Abs((New Date(y, mm, dd) - ts).TotalDays)) _
                    .First()
                        End If

                        results.Add(New FileDateInfo With {.FileName = fileName, .Month = mm, .Day = dd, .Year = yy})
                        Continue For
                    End If
                    Dim mMon = rxMonthWk.Match(fileName)
                    If mMon.Success Then
                        Dim dt As DateTime
                        If Not DateTime.TryParseExact(mMon.Groups(1).Value,
                                New String() {"MMM", "MMMM", "MMM.", "MMMM.", "Sept", "Sept."},
                                CultureInfo.InvariantCulture, DateTimeStyles.None, dt) Then
                            Continue For
                        End If

                        Dim wk As Integer
                        If Not Integer.TryParse(mMon.Groups(2).Value, wk) Then Continue For
                        If wk < 1 OrElse wk > 5 Then Continue For  ' sanity

                        Dim inferredYear As Integer = System.IO.File.GetLastWriteTime(filePath).Year

                        results.Add(New FileDateInfo With {
                            .FileName = fileName, .Month = dt.Month, .Day = 0, .Year = inferredYear, .WeekOfMonth = wk})
                    End If

                Next

            Case Else
                Dim rx As New Regex("([A-Za-z]+)\s+(\d{2,4})", RegexOptions.IgnoreCase)

                For Each filePath In System.IO.Directory.EnumerateFiles(directory, "*.xlsx")
                    Dim fileName = Path.GetFileName(filePath)
                    If fileName.StartsWith("~$") Then Continue For

                    Dim m = rx.Match(fileName)
                    If Not m.Success Then Continue For

                    Dim monthStr As String = m.Groups(1).Value
                    Dim yearStr As String = m.Groups(2).Value

                    Dim year As Integer
                    If Not Integer.TryParse(yearStr, year) Then Continue For
                    If year < 100 Then year += 2000

                    Dim dt As DateTime
                    If Not DateTime.TryParseExact(monthStr,
                                              New String() {"MMM", "MMMM"},
                                              CultureInfo.InvariantCulture,
                                              DateTimeStyles.None,
                                              dt) Then Continue For

                    results.Add(New FileDateInfo With {.FileName = fileName, .Month = dt.Month, .Year = year})
                Next
        End Select

        Return results.OrderBy(Function(r) r.Year).ThenBy(Function(r) r.Month).ThenBy(Function(r) r.Day).ToList()

    End Function


    Public Class FileDateInfo
        Public Property FileName As String
        Public Property Month As Integer
        Public Property Year As Integer
        Public Property Day As Integer
        Public Property WeekOfMonth As Integer
    End Class

    Public Shared Sub Create_RSTXLSQ1_Records(frm As ASFBASE0, ByVal CUST_CODE As String, filesToImport As List(Of FileDateInfo))

        frm.dst.Tables("RSTXLSQ1").Rows.Clear()

        Dim PrevSat As Func(Of Date, Date) =
        Function(d As Date) d.AddDays(-(((CInt(d.DayOfWeek) + 1) Mod 7)))

        For Each f In filesToImport

            Dim XLS_DOC_SEQ_NO As String = ASCMAIN1.Next_Control_No("RSTXLSQ1.XLS_DOC_SEQ_NO")
            Dim XLS_DOC_FILENAME As String = f.FileName

            Dim importIsWeekly As Boolean = True
            Select Case CUST_CODE.ToUpperInvariant()
                Case "ANTHRO", "EQUINOX", "SDM", "GOOP"
                    importIsWeekly = False
            End Select

            Dim rowRSTXLSQ1 As DataRow = frm.dst.Tables("RSTXLSQ1").NewRow
            rowRSTXLSQ1.Item("XLS_DOC_SEQ_NO") = XLS_DOC_SEQ_NO
            rowRSTXLSQ1.Item("XLS_DOC_FILENAME") = XLS_DOC_FILENAME
            rowRSTXLSQ1.Item("CUST_CODE") = CUST_CODE
            rowRSTXLSQ1.Item("XLS_DOC_STATUS") = "0"
            rowRSTXLSQ1.Item("XLS_DOC_SOURCE") = "XLS"
            rowRSTXLSQ1.Item("INIT_DATE") = Now
            rowRSTXLSQ1.Item("INIT_OPER") = ASCMAIN1.USER_ID

            Dim XLS_DATE_TO As Date
            Dim XLS_DATE_FROM As Date

            If importIsWeekly Then
                If f.Day > 0 Then
                    Dim dt As New Date(f.Year, f.Month, f.Day)
                    Dim sat As Date = PrevSat(dt)
                    Dim rowGLTPARM3 As DataRow =
                    ASCDATA1.GetDataRow("Select YYYYPP, YYYYWW, WEEK_END_DATE DATE_TO from GLTPARM3 where WEEK_END_DATE = :PARM1",
                                        "D", New Object() {sat})
                    If rowGLTPARM3 Is Nothing Then
                        Throw New ApplicationException($"GLTPARM3 not found for week ending {sat:d}.")
                    End If
                    XLS_DATE_TO = CDate(rowGLTPARM3("DATE_TO"))
                ElseIf CUST_CODE.Equals("VONMAUR", StringComparison.OrdinalIgnoreCase) AndAlso f.WeekOfMonth > 0 Then
                    Dim firstOfMonth As New Date(f.Year, f.Month, 1)
                    Dim offsetToSat As Integer = (6 - CInt(firstOfMonth.DayOfWeek) + 7) Mod 7
                    Dim firstSat As Date = firstOfMonth.AddDays(offsetToSat)

                    Dim wk1 As Date = If(firstSat.Day <= 6, firstSat.AddDays(7), firstSat)

                    Dim targetSat As Date = wk1.AddDays(7 * (f.WeekOfMonth - 1))

                    Dim rowGLTPARM3 As DataRow =
                        ASCDATA1.GetDataRow("Select YYYYPP, YYYYWW, WEEK_END_DATE DATE_TO " &
                                            "from GLTPARM3 where WEEK_END_DATE = :PARM1",
                                            "D", New Object() {targetSat})
                    If rowGLTPARM3 Is Nothing Then
                        Throw New ApplicationException($"GLTPARM3 not found for week ending {targetSat:d}.")
                    End If
                    XLS_DATE_TO = CDate(rowGLTPARM3("DATE_TO"))

                Else
                    ' no day parsed -> use the latest week end date for that accounting period
                    Dim YYYYPP As String = $"{f.Year:0000}{f.Month:00}"
                    Dim rowGLTPARM3 As DataRow =
                    ASCDATA1.GetDataRow("Select MAX(WEEK_END_DATE) DATE_TO, MIN(YYYYPP) YYYYPP, MIN(YYYYWW) YYYYWW " &
                                        "from GLTPARM3 where YYYYPP = :PARM1",
                                        "V", New Object() {YYYYPP})
                    If rowGLTPARM3 Is Nothing OrElse rowGLTPARM3.IsNull("DATE_TO") Then
                        Throw New ApplicationException($"GLTPARM3 not found for YYYYPP={YYYYPP}.")
                    End If
                    XLS_DATE_TO = CDate(rowGLTPARM3("DATE_TO"))
                End If

                XLS_DATE_FROM = XLS_DATE_TO.AddDays(-6)

            Else
                ' monthly import: last week ending of the period
                Dim YYYYPP As String = $"{f.Year:0000}{f.Month:00}"
                Dim rowGLTPARM3 As DataRow =
            ASCDATA1.GetDataRow("Select MIN(YYYYPP) YYYYPP, MIN(YYYYWW) YYYYWW, MIN(WEEK_END_DATE) DATE_TO " &
                                "from GLTPARM3 where WEEK_END_DATE = " &
                                "(select MAX(WEEK_END_DATE) from GLTPARM3 where YYYYPP = :PARM1)",
                                "V", New Object() {YYYYPP})
                If rowGLTPARM3 Is Nothing OrElse rowGLTPARM3.IsNull("DATE_TO") Then
                    Throw New ApplicationException($"GLTPARM3 not found for YYYYPP={YYYYPP}.")
                End If
                XLS_DATE_TO = CDate(rowGLTPARM3("DATE_TO"))
                XLS_DATE_FROM = XLS_DATE_TO.AddDays(-6)
            End If

            rowRSTXLSQ1.Item("XLS_DATE_TO") = XLS_DATE_TO
            rowRSTXLSQ1.Item("XLS_DATE_FROM") = XLS_DATE_FROM

            frm.dst.Tables("RSTXLSQ1").Rows.Add(rowRSTXLSQ1)


            '\\ABSNASQ\Public\SLP\Share\SLP\XLS\InBound\0000000001_49041526390970070544200402524.csv
            Dim RS_PARM_XLS_FOLDER As String = frm.ROWs("RSTPARM1").Item("RS_PARM_XLS_FOLDER") & ""
            Dim fileSource As String = System.IO.Path.Combine(RS_PARM_XLS_FOLDER.Replace("InBound", "History"), CUST_CODE, XLS_DOC_FILENAME)
            Dim fileDest As String = System.IO.Path.Combine(RS_PARM_XLS_FOLDER, $"{CUST_CODE}-{XLS_DOC_SEQ_NO}-{XLS_DOC_FILENAME}")
            System.IO.File.Copy(fileSource, fileDest)

        Next

        If frm.dst.Tables("RSTXLSQ1").Rows.Count > 0 Then
            Dim msg As String = $"{frm.dst.Tables("RSTXLSQ1").Rows.Count} workbooks queued for import"
            frm.BeginTrans()
            frm.Update_Record_TDA("RSTXLSQ1")
            frm.CommitTrans()
        End If


    End Sub
    Public Shared Function Validate_Item_Code(frm As ASFBASE0, ITEM_UPC_CODE_XLS As String, CUST_CODE As String) As String
        Dim ITEM_CODE_VALID As String = ""
        ASCMAIN1.sql = "Select * from ICTITEM1 where ITEM_CODE_UPC = :PARM1"

        Dim rowICTITEMU As DataRow = frm.dst.Tables("ICTITEMU").Rows.Find(New String() {ITEM_UPC_CODE_XLS})
        If rowICTITEMU IsNot Nothing Then
            ITEM_CODE_VALID = rowICTITEMU.Item("ITEM_CODE")
        Else
            Dim rowEDTUPCX1 As DataRow = frm.dst.Tables("EDTUPCX1").Rows.Find(New String() {CUST_CODE, ITEM_UPC_CODE_XLS})
            If rowEDTUPCX1 IsNot Nothing Then
                If rowEDTUPCX1.Item("ITEM_CODE") & "" <> "" Then
                    ITEM_CODE_VALID = rowEDTUPCX1.Item("ITEM_CODE")
                ElseIf rowEDTUPCX1.Item("IGNORE") & "" = "1" Then
                    ITEM_CODE_VALID = "IGNORE"
                Else
                    If frm.dst.Tables("RSTXLSQE").Rows.Find(New String() {CUST_CODE, ITEM_UPC_CODE_XLS}) Is Nothing Then
                        frm.dst.Tables("RSTXLSQE").Rows.Add(New String() {CUST_CODE, ITEM_UPC_CODE_XLS})
                    End If
                End If
            Else
                ITEM_UPC_CODE_XLS = ITEM_UPC_CODE_XLS & ""
                If frm.dst.Tables("RSTXLSQE").Rows.Find(New String() {CUST_CODE, ITEM_UPC_CODE_XLS}) Is Nothing Then
                    frm.dst.Tables("RSTXLSQE").Rows.Add(New String() {CUST_CODE, ITEM_UPC_CODE_XLS})
                    frm.dst.Tables("EDTUPCX1").Rows.Add(New String() {CUST_CODE, ITEM_UPC_CODE_XLS})
                End If
            End If

        End If
        Return ITEM_CODE_VALID
    End Function

    Public Shared Sub Rollback_Import_For_Customer(frm As ASFBASE0, CUST_CODE As String)

        Dim RS_PARM_XLS_FOLDER As String = frm.ROWs("RSTPARM1").Item("RS_PARM_XLS_FOLDER") & ""

        Try
            For Each filePath As String In Directory.GetFiles(RS_PARM_XLS_FOLDER)
                Dim fileName As String = Path.GetFileName(filePath)
                If fileName.IndexOf(CUST_CODE, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    File.Delete(filePath)
                End If
            Next
            For Each filePath As String In Directory.GetFiles(RS_PARM_XLS_FOLDER.Replace("InBound", "InBound_Archive"))
                Dim fileName As String = Path.GetFileName(filePath)
                If fileName.IndexOf(CUST_CODE, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    File.Delete(filePath)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "Error Deleting Workbook")
        End Try

        frm.BeginTrans()
        For Each TABLE_NAME As String In New String() {"RSTXLSQ1", "RSTRETL1", "EDT852T1", "RSTRETL2", "RSTRETL4"}
            ASCMAIN1.sql = $"DELETE FROM {TABLE_NAME} WHERE CUST_CODE = :PARM1"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {CUST_CODE})
        Next
        frm.CommitTrans("Rollback Complete!")

    End Sub
    'scans a row to find the first column whose header matches one of ours
    Private Shared Function Find_Header_Column(cells As IRange, headerRow As Integer, headerNames As IEnumerable(Of String)) As Integer
        For c As Integer = 0 To cells.ColumnCount - 1
            Dim hdr As String = Trim(CStr(cells(headerRow, c).Text))
            If hdr.Length = 0 Then Continue For
            For Each name In headerNames
                If String.Equals(hdr, name, StringComparison.OrdinalIgnoreCase) Then Return c
            Next
        Next
        Return -1
    End Function
    'tries to return the requested worksheet
    Private Shared Function Get_Sheet(wb As IWorkbook, ParamArray names() As String) As IWorksheet
        If names IsNot Nothing Then
            For Each n In names
                For Each s As IWorksheet In wb.Worksheets
                    If s.Name.Equals(n, StringComparison.OrdinalIgnoreCase) Then Return s
                Next
            Next
        End If
        Return wb.Worksheets(0)
    End Function
    'tries to return a worksheet who contains "SKU"
    Private Shared Function Get_Sheet_Like(wb As IWorkbook, containsText As String) As IWorksheet
        For Each s As IWorksheet In wb.Worksheets
            If s.Name.IndexOf(containsText, StringComparison.OrdinalIgnoreCase) >= 0 Then Return s
        Next
        Return Nothing
    End Function
    'strips number of $, spaces, commas, parentheses and converts to int/decimal
    Private Shared Function Clean_Numeric(t As String) As String
        Dim s As String = If(t, "")
        s = s.Replace(ChrW(160), " ").Replace("$", "").Replace(",", "").Trim()
        If s.StartsWith("("c) AndAlso s.EndsWith(")"c) Then s = "-" & s.Substring(1, s.Length - 2)
        Return s
    End Function
    Private Shared Function ToInt(t As String) As Integer
        Dim n As Integer
        Integer.TryParse(Clean_Numeric(t), NumberStyles.Integer, CultureInfo.InvariantCulture, n)
        Return n
    End Function
    Private Shared Function ToDec(t As String) As Decimal
        Dim n As Decimal
        Decimal.TryParse(Clean_Numeric(t), NumberStyles.Number Or NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, n)
        Return n
    End Function
    'reads UPCs whether stored as text, scientific notation, or numbers
    Private Shared Function Try_Read_UPC(c As IRange) As String
        Dim v As Object = c.Value
        If v Is Nothing Then Return ""
        If TypeOf v Is Double OrElse TypeOf v Is Decimal Then
            Return Decimal.Truncate(Convert.ToDecimal(v)).ToString("0", CultureInfo.InvariantCulture)
        End If
        Dim t As String = Trim(CStr(c.Text))
        Dim d As Double
        If t.IndexOf("E"c) >= 0 AndAlso Double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, d) Then
            Return Decimal.Truncate(Convert.ToDecimal(d)).ToString("0", CultureInfo.InvariantCulture)
        End If
        Return t
    End Function
    'Finds or creates RSTRETL1 row for Doc/Cust/Store/Item combo
    Private Shared Function Get_Or_Add_rowRSTRETL1(
    frm As ASFBASE0, rowEDT852T1 As DataRow,
    custCode As String, store As String, item As String,
    Optional ypp As String = Nothing, Optional yww As String = Nothing
) As DataRow

        Dim useYPP As String = If(String.IsNullOrEmpty(ypp), CStr(rowEDT852T1("OPS_YYYYPP")), ypp)
        Dim useYWW As String = If(String.IsNullOrEmpty(yww), CStr(rowEDT852T1("OPS_YYYYWW")), yww)

        Dim key = New String() {
        CStr(rowEDT852T1("EDI_DOC_SEQ_NO")), custCode, store, item, useYPP, useYWW
    }

        Dim r As DataRow = frm.dst.Tables("RSTRETL1").Rows.Find(key)
        If r Is Nothing Then
            r = frm.dst.Tables("RSTRETL1").NewRow()
            r("EDI_DOC_SEQ_NO") = rowEDT852T1("EDI_DOC_SEQ_NO")
            r("CUST_CODE") = custCode
            r("CUST_STORE_NO") = store
            r("ITEM_CODE") = item
            r("OPS_YYYYPP") = useYPP
            r("OPS_YYYYWW") = useYWW
            r("QTY_SOLD") = 0
            r("AMT_SOLD") = CDec(0)
            r("QTY_EOW") = 0
            frm.dst.Tables("RSTRETL1").Rows.Add(r)
        End If
        Return r
    End Function
    Private Shared Sub Add_Qtys_And_Amts(
    frm As ASFBASE0, rowEDT852T1 As DataRow,
    custCode As String, store As String, item As String,
    qty As Object, amt As Object, eow As Object,
    Optional ypp As String = Nothing, Optional yww As String = Nothing
)
        Dim r = Get_Or_Add_rowRSTRETL1(frm, rowEDT852T1, custCode, store, item, ypp, yww)
        If Not IsDBNull(qty) Then r("QTY_SOLD") = CInt(If(IsDBNull(r("QTY_SOLD")), 0, r("QTY_SOLD"))) + Convert.ToInt32(qty)
        If Not IsDBNull(amt) Then r("AMT_SOLD") = CDec(If(IsDBNull(r("AMT_SOLD")), 0D, r("AMT_SOLD"))) + Convert.ToDecimal(amt)
        If Not IsDBNull(eow) Then r("QTY_EOW") = CInt(If(IsDBNull(r("QTY_EOW")), 0, r("QTY_EOW"))) + Convert.ToInt32(eow)
    End Sub

    ' Returns 0 for blank/placeholder text and never throws.
    Private Shared Function CellDec(c As SpreadsheetGear.IRange) As Decimal
        Dim v As Object = c.Value
        If v Is Nothing OrElse v Is DBNull.Value Then Return 0D

        ' Fast path for numeric cells
        If TypeOf v Is Double OrElse TypeOf v Is Decimal OrElse TypeOf v Is Integer OrElse TypeOf v Is Long Then
            Return Convert.ToDecimal(v, CultureInfo.InvariantCulture)
        End If

        ' Fall back to text parsing (handle $, commas, NBSP, parentheses, accounting dash, etc.)
        Dim t As String = (c.Text & "").Trim()
        If t.Length = 0 Then Return 0D
        t = t.Replace(ChrW(160), " "c).Trim()          ' NBSP -> space
        If t = "-" OrElse t = "—" OrElse t = "–" Then Return 0D    ' accounting dash variants
        If t.Equals("N/A", StringComparison.OrdinalIgnoreCase) Then Return 0D

        Dim isNeg As Boolean = t.StartsWith("(") AndAlso t.EndsWith(")")
        If isNeg Then t = t.Substring(1, t.Length - 2)

        ' strip currency and thousand separators/spaces
        t = t.Replace("$", "").Replace(",", "").Replace(" ", "")

        Dim d As Decimal
        If Decimal.TryParse(t, NumberStyles.Number Or NumberStyles.AllowLeadingSign Or NumberStyles.AllowDecimalPoint,
                        CultureInfo.InvariantCulture, d) Then
            Return If(isNeg, -d, d)
        End If
        Return 0D
    End Function

    Private Shared Function CellInt(c As SpreadsheetGear.IRange) As Integer
        Dim v As Object = c.Value
        If v Is Nothing OrElse v Is DBNull.Value Then Return 0

        ' Fast path for numeric cells
        If TypeOf v Is Double OrElse TypeOf v Is Decimal OrElse TypeOf v Is Integer OrElse TypeOf v Is Long Then
            ' Truncate like Excel displays integers from numeric cells
            Return CInt(Math.Truncate(Convert.ToDecimal(v, CultureInfo.InvariantCulture)))
        End If

        ' Fall back to text parsing
        Dim t As String = (c.Text & "").Trim()
        If t.Length = 0 Then Return 0
        t = t.Replace(ChrW(160), " "c).Trim()
        If t = "-" OrElse t = "—" OrElse t = "–" Then Return 0
        If t.Equals("N/A", StringComparison.OrdinalIgnoreCase) Then Return 0

        Dim isNeg As Boolean = t.StartsWith("(") AndAlso t.EndsWith(")")
        If isNeg Then t = t.Substring(1, t.Length - 2)

        t = t.Replace("$", "").Replace(",", "").Replace(" ", "")

        ' Sometimes units sneak in with a decimal (e.g., "12.0") — accept them.
        Dim d As Decimal
        If Decimal.TryParse(t, NumberStyles.Number Or NumberStyles.AllowLeadingSign Or NumberStyles.AllowDecimalPoint,
                        CultureInfo.InvariantCulture, d) Then
            Dim i As Integer = CInt(Math.Truncate(d))
            Return If(isNeg, -i, i)
        End If
        Return 0
    End Function


    'normalizes date to the Saturday of the week prior
    Private Shared Function Prev_Saturday(d As Date) As Date
        ' Saturday=6
        Return d.AddDays(-(((CInt(d.DayOfWeek) + 1) Mod 7)))
    End Function
    'looks up GLTPARM3 to get period/week for given Saturday
    Private Shared Function Get_Week_And_Year(weekEndSaturday As Date) As (String, String)?
        ASCMAIN1.sql = "Select * from GLTPARM3 where WEEK_END_DATE = :PARM1"
        Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, False, "D", New Object() {weekEndSaturday})
        If row Is Nothing Then Return Nothing
        Return (CStr(row("YYYYPP")), CStr(row("YYYYWW")))
    End Function
    'Skips duplicate rows for HUTGROUP
    Private Shared Function AlreadyImported(frm As ASFBASE0, cust As String, item As String, yww As String, ByRef seen As HashSet(Of String)) As Boolean
        Dim key = item & "|" & yww
        If seen.Contains(key) Then Return True
        ASCMAIN1.sql = "select 1 from RSTRETL1 where CUST_CODE = :PARM1 and CUST_STORE_NO = '000000' and ITEM_CODE = :PARM2 and OPS_YYYYWW = :PARM3"
        Dim hit As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVV", New Object() {cust, item, yww})
        If hit IsNot Nothing Then
            seen.Add(key)
            Return True
        End If
        Return False
    End Function
    'Locates header row for SDM
    Private Shared Function DetectHeaderRow(cells As IRange, colAmt As Integer, colQty As Integer, colLabel As Integer, Optional fallback As Integer = 2) As Integer
        For r As Integer = 0 To Math.Min(15, cells.RowCount - 1)
            Dim eHdr As String = Trim(CStr(cells(r, colAmt).Text))
            Dim iHdr As String = Trim(CStr(cells(r, colQty).Text))
            Dim bHdr As String = Trim(CStr(cells(r, colLabel).Text))
            If (eHdr.StartsWith("Sales", StringComparison.OrdinalIgnoreCase)) AndAlso
               (iHdr.StartsWith("Units", StringComparison.OrdinalIgnoreCase) OrElse bHdr.Equals("SKU", StringComparison.OrdinalIgnoreCase)) Then
                Return r
            End If
        Next
        Return fallback
    End Function
    'loads main sheet/data sheet, reads fields, adds to RSTRETL1 
    Private Shared Sub Import_HOLT(frm As ASFBASE0, wb As IWorkbook, rowEDT852T1 As DataRow, CUST_CODE As String)
        Dim ws As IWorksheet = Get_Sheet(wb, "852 DATA", "852 Data", "Data", "Sheet1")
        Dim cells As IRange = ws.UsedRange

        Dim priceMap As New Dictionary(Of String, Decimal)(StringComparer.Ordinal)
        Dim wsData As IWorksheet = Nothing
        For Each s As IWorksheet In wb.Worksheets
            If s.Name.Equals("DATA", StringComparison.OrdinalIgnoreCase) Then wsData = s : Exit For
        Next
        If wsData IsNot Nothing Then
            Dim d As IRange = wsData.UsedRange
            Dim COL_UPC_D As Integer = ColumnLetterToIndex("A")
            Dim COL_PRICE_D As Integer = ColumnLetterToIndex("C")
            For r As Integer = 1 To d.RowCount - 1
                Dim upc As String = Try_Read_UPC(d(r, COL_UPC_D))
                If upc <> "" AndAlso Not priceMap.ContainsKey(upc) Then
                    Dim pr As Decimal = CellDec(d(r, COL_PRICE_D))
                    If pr > 0D Then priceMap(upc) = pr
                End If
            Next
        End If

        Dim COL_STORE As Integer = ColumnLetterToIndex("D")
        Dim COL_ITEM As Integer = ColumnLetterToIndex("N")
        Dim COL_EOW As Integer = ColumnLetterToIndex("Q")
        Dim COL_QS As Integer = ColumnLetterToIndex("AB")
        Dim COL_PRICE As Integer = ColumnLetterToIndex("BC")
        Dim COL_AMT As Integer = ColumnLetterToIndex("BD")

        For r As Integer = 1 To cells.RowCount - 1
            Dim storeRaw As String = Trim(CStr(cells(r, COL_STORE).Text))
            If storeRaw = "" Then Continue For
            Dim store As String = storeRaw.PadLeft(6, "0"c)

            Dim upc As String = Try_Read_UPC(cells(r, COL_ITEM))
            If upc = "" Then Continue For

            Dim item As String = Validate_Item_Code(frm, upc, CUST_CODE)
            If item = "" OrElse item = "IGNORE" Then Continue For

            Dim qtySold As Integer = ToInt(CStr(cells(r, COL_QS).Text))
            Dim qtyEow As Integer = ToInt(CStr(cells(r, COL_EOW).Text))

            Dim amtSold As Decimal
            Dim vAmt As Object = cells(r, COL_AMT).Value
            If TypeOf vAmt Is Double OrElse TypeOf vAmt Is Decimal Then
                amtSold = Convert.ToDecimal(vAmt, CultureInfo.InvariantCulture)
            Else
                amtSold = ToDec(CStr(cells(r, COL_AMT).Text))
            End If

            If (amtSold = 0D OrElse Not amtSold.Equals(amtSold)) AndAlso qtySold > 0 Then
                Dim unitPrice As Decimal = 0D
                If priceMap.TryGetValue(upc, unitPrice) Then
                    amtSold = unitPrice * qtySold
                Else
                    unitPrice = ToDec(CStr(cells(r, COL_PRICE).Text))
                    If unitPrice <= 0D Then
                        Dim vP As Object = cells(r, COL_PRICE).Value
                        If TypeOf vP Is Double OrElse TypeOf vP Is Decimal Then
                            unitPrice = Convert.ToDecimal(vP, CultureInfo.InvariantCulture)
                        End If
                    End If
                    If unitPrice > 0D Then amtSold = unitPrice * qtySold
                End If
                If amtSold <> 0D Then amtSold = Math.Round(amtSold, 2, MidpointRounding.AwayFromZero)
            End If

            Add_Qtys_And_Amts(frm, rowEDT852T1, CUST_CODE, store, item, qtySold, amtSold, qtyEow)
        Next
    End Sub
    ' SDM: merge this file’s qty/amt into an existing period row if one exists.
    Private Shared Function SqlQ(s As String) As String
        If s Is Nothing Then s = ""
        Return "'" & s.Replace("'", "''") & "'"
    End Function

    Private Shared Sub Add_Qtys_And_Amts_SDM_Merge(frm As ASFBASE0,
                                                   rowEDT852T1 As DataRow,
                                                   custCode As String,
                                                   item As String,
                                                   qty As Integer,
                                                   amt As Decimal)
        If qty = 0 AndAlso amt = 0D Then Exit Sub

        Const store As String = "000000"
        Dim ypp As String = CStr(rowEDT852T1("OPS_YYYYPP"))

        Dim litCust = SqlQ(custCode)
        Dim litItem = SqlQ(item)
        Dim litYpp = SqlQ(ypp)
        Dim litQty = qty.ToString(CultureInfo.InvariantCulture)
        Dim litAmt = amt.ToString(CultureInfo.InvariantCulture)

        ASCMAIN1.sql =
            "select EDI_DOC_SEQ_NO from RSTRETL1 " &
            "where CUST_CODE = " & litCust & " and CUST_STORE_NO = '000000' " &
            "and ITEM_CODE = " & litItem & " and OPS_YYYYPP = " & litYpp

        Dim existing As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)

        If existing IsNot Nothing Then
            Dim docNo As String = CStr(existing("EDI_DOC_SEQ_NO"))
            Dim litDoc = SqlQ(docNo)

            ASCMAIN1.sql =
                "update RSTRETL1 " &
                "set QTY_SOLD = NVL(QTY_SOLD,0) + " & litQty & ", " &
                "    AMT_SOLD = NVL(AMT_SOLD,0) + " & litAmt & " " &
                "where EDI_DOC_SEQ_NO = " & litDoc & " " &
                "  and CUST_CODE = " & litCust & " and CUST_STORE_NO = '000000' " &
                "  and ITEM_CODE = " & litItem & " and OPS_YYYYPP = " & litYpp

            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            Exit Sub
        End If
        Add_Qtys_And_Amts(frm, rowEDT852T1, custCode, store, item, qty, amt, DBNull.Value)
    End Sub
    ' Scan the first few rows for a header that equals/starts-with any of the names.
    Private Shared Function Find_Header_ColAndRow_Anywhere(cells As IRange,
                                                       headerNames As IEnumerable(Of String),
                                                       Optional maxScanRows As Integer = 15) As (Col As Integer, Row As Integer)
        Dim limit As Integer = Math.Min(maxScanRows, cells.RowCount - 1)
        For r As Integer = 0 To limit
            For c As Integer = 0 To cells.ColumnCount - 1
                Dim hdr As String = Trim(CStr(cells(r, c).Text)).Replace(ChrW(160), " ")
                If hdr = "" Then Continue For
                For Each name In headerNames
                    If hdr.Equals(name, StringComparison.OrdinalIgnoreCase) _
                   OrElse hdr.StartsWith(name, StringComparison.OrdinalIgnoreCase) Then
                        Return (c, r)
                    End If
                Next
            Next
        Next
        Return (-1, -1)
    End Function
    ' True if the row already exists in DB or we’ve already added it in this run
    Private Shared Function Exists_In_DB_Or_Added(frm As ASFBASE0,
                                              cust As String,
                                              store As String,
                                              item As String,
                                              ypp As String,
                                              yww As String,
                                              cache As HashSet(Of String)) As Boolean
        Dim key As String = item & "|" & ypp & "|" & yww
        If cache.Contains(key) Then Return True

        ASCMAIN1.sql = "select 1 from RSTRETL1 " &
                   "where CUST_CODE = :P1 and CUST_STORE_NO = :P2 " &
                   "  and ITEM_CODE = :P3 and OPS_YYYYPP = :P4 and OPS_YYYYWW = :P5"
        Dim hit As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVVVV",
                                             New Object() {cust, store, item, ypp, yww})
        If hit IsNot Nothing Then
            cache.Add(key)
            Return True
        End If
        Return False
    End Function
    Private Shared Function IsEquinoxEcommFile(fileName As String) As Boolean
        Dim n As String = (fileName & "")
        Return Regex.IsMatch(n, "\b(ECOMM?|ONLINE)\b", RegexOptions.IgnoreCase)
    End Function

    Private Shared Function Normalize_Spaces(s As String) As String
        Return Regex.Replace((s & "").Trim(), "\s+", " ")
    End Function
    Private Shared Function Build_Equinox_Ecom_Key(productTitle As String, variantTitle As String) As String
        Dim p As String = Normalize_Item_Code(productTitle, True)
        Dim v As String = Normalize_Item_Code(variantTitle, True)
        If v = "" Then Return p
        Return p & "|" & v
    End Function
    Private Shared Function Build_Equinox_Store_Key(descText As String, colorText As String) As String
        Dim d As String = Normalize_Spaces(descText)
        Dim k As String = Normalize_Spaces(colorText)
        If d = "" AndAlso k = "" Then Return ""
        If k = "" Then Return d
        If d = "" Then Return k
        Return d & "|" & k
    End Function


    Private Shared Function Normalize_Item_Code(raw As String,
                                             Optional replaceSpaces As Boolean = True) As String
        Dim s As String = Regex.Replace((raw & "").Trim(), "\s+", " ")
        If replaceSpaces Then s = s.Replace(" "c, "_"c)
        s = Regex.Replace(s, "[\u0000-\u001F\u007F]", "")
        Return s
    End Function
    ' Exact-match version: only returns a column if the header equals one of the names.
    Private Shared Function Find_Header_ColAndRow_Anywhere_Exact(
    cells As IRange, headerNames As IEnumerable(Of String),
    Optional maxScanRows As Integer = 15) As (Col As Integer, Row As Integer)

        Dim limit As Integer = Math.Min(maxScanRows, cells.RowCount - 1)
        For r As Integer = 0 To limit
            For c As Integer = 0 To cells.ColumnCount - 1
                Dim hdr As String = Trim(CStr(cells(r, c).Text)).Replace(ChrW(160), " ")
                If hdr = "" Then Continue For
                For Each name In headerNames
                    If hdr.Equals(name, StringComparison.OrdinalIgnoreCase) Then
                        Return (c, r)
                    End If
                Next
            Next
        Next
        Return (-1, -1)
    End Function
    Public Shared Sub Import_ADS_For(frm As ASFBASE0, ByVal XLS_DOC_SEQ_NO As String)
        Dim rowRSTXLSQ1 As DataRow = frm.LookUp("RSTXLSQ1", XLS_DOC_SEQ_NO)
        If rowRSTXLSQ1 Is Nothing Then Exit Sub

        Dim CUST_CODE As String = CStr(rowRSTXLSQ1("CUST_CODE") & "")
        If Not String.Equals(CUST_CODE, "ADS", StringComparison.OrdinalIgnoreCase) Then Exit Sub

        Dim RS_PARM_XLS_FOLDER As String = frm.ROWs("RSTPARM1").Item("RS_PARM_XLS_FOLDER") & ""
        Dim baseName As String = CStr(rowRSTXLSQ1("XLS_DOC_FILENAME"))

        Dim candidates As String() = {
            $"{CUST_CODE}-{XLS_DOC_SEQ_NO}-{baseName}",
            $"{XLS_DOC_SEQ_NO}-{baseName}",
            baseName
        }

        Dim importFile As String = Nothing
        For Each fn In candidates
            Dim path = System.IO.Path.Combine(RS_PARM_XLS_FOLDER, fn)
            If My.Computer.FileSystem.FileExists(path) Then
                importFile = path
                Exit For
            End If
        Next

        If String.IsNullOrEmpty(importFile) Then
            Exit Sub
        End If

        Dim wb As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(importFile)
        Dim ws As SpreadsheetGear.IWorksheet = wb.Worksheets(0)
        Dim cells As SpreadsheetGear.IRange = ws.UsedRange

        Dim HDR As Integer = 10

        Dim COL_ORDER_DATE = Find_Header_Column(cells, HDR, {"Order Date", "ORDER_DATE"})
        Dim COL_START_SHIP = Find_Header_Column(cells, HDR, {"Start Ship", "START_SHIP"})
        Dim COL_REQ_DATE = Find_Header_Column(cells, HDR, {"Req. Date", "REQ_DATE", "Rep Date", "REP_DATE"})
        Dim COL_ORDER_NO = Find_Header_Column(cells, HDR, {"Order #", "Order No", "ORDER_NO"})
        Dim COL_SOLD_TO = Find_Header_Column(cells, HDR, {"Sold-To", "Sold To", "SOLD_TO"})
        Dim COL_CUSTOMER_NAME = Find_Header_Column(cells, HDR, {"Customer Name", "CUSTOMER_NAME"})
        Dim COL_CUST_ORDER_REF = Find_Header_Column(cells, HDR, {"Cust. Order Ref", "Cust Order Ref", "PO #", "CUST_ORDER_REF"})
        Dim COL_CARRIER = Find_Header_Column(cells, HDR, {"Carrier", "CARRIER"})
        Dim COL_IMPORT_WINDOW = Find_Header_Column(cells, HDR, {"Import", "Import Window", "IMPORT_WINDOW"})
        Dim COL_PICKING_WINDOW = Find_Header_Column(cells, HDR, {"Picking", "Picking Window", "PICKING_WINDOW"})
        Dim COL_DELIVERY_WINDOW = Find_Header_Column(cells, HDR, {"Delivery", "Delivery Window", "DELIVERY_WINDOW"})
        Dim COL_BOL_NO = Find_Header_Column(cells, HDR, {"BOL #", "BOL", "BOL_NO"})
        Dim COL_PRO_NO = Find_Header_Column(cells, HDR, {"PRO #", "PRO_NO"})
        Dim COL_EDI_INV = Find_Header_Column(cells, HDR, {"EDI Invoice", "EDI_INVOICE"})
        Dim COL_EDI_ASN = Find_Header_Column(cells, HDR, {"EDI ASN", "EDI_ASN"})
        Dim COL_ORDER_VALUE = Find_Header_Column(cells, HDR, {"Order Value", "ORDER_VALUE", "Value"})

        If COL_ORDER_NO < 0 Then Throw New ApplicationException("ADS import: could not locate 'Order No' column.")

        frm.BeginTrans()
        Try
            Dim ADS_DOC_SEQ_NO As String = ASCMAIN1.Next_Control_No("ZORDSTA2.ADS_DOC_SEQ_NO")
            Dim hdrDateTxt As String = Read_Header_Text(ws, "P3", "Q3")
            Dim hdrTimeTxt As String = Read_Header_Text(ws, "P6", "Q6")

            Dim reportDate As Object = Parse_Header_Date(hdrDateTxt)
            Dim reportTimeTxt As String = hdrTimeTxt

            ASCMAIN1.sql =
            "INSERT INTO ZORDSTA1 (ADS_DOC_SEQ_NO, REPORT_DATE, REPORT_TIME_TXT, INIT_DATE, INIT_OPER) VALUES (" &
            SQL_Txt(ADS_DOC_SEQ_NO) & "," &
            SQL_Date(reportDate) & "," &
            SQL_Txt(reportTimeTxt) & "," &
            "SYSDATE," & SQL_Txt(ASCMAIN1.USER_ID) & ")"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            For r As Integer = HDR + 1 To cells.RowCount - 1
                Dim orderNo As String = Trim(CStr(cells(r, COL_ORDER_NO).Text))
                If orderNo = "" Then Continue For

                Dim dtOrder = Read_Date(cells, r, COL_ORDER_DATE)
                Dim dtStartShip = Read_Date(cells, r, COL_START_SHIP)
                Dim dtReq = Read_Date(cells, r, COL_REQ_DATE)

                Dim soldTo = Read_Txt(cells, r, COL_SOLD_TO, 30)
                Dim custNameRaw = Read_Txt(cells, r, COL_CUSTOMER_NAME, 200)
                Dim custName = CleanCustomerName(soldTo, custNameRaw)

                Dim custPo = Read_Txt(cells, r, COL_CUST_ORDER_REF, 60)
                Dim carrier = Read_Txt(cells, r, COL_CARRIER, 60)
                Dim importWin = Read_Txt(cells, r, COL_IMPORT_WINDOW, 60)
                Dim pickingWin = Read_Txt(cells, r, COL_PICKING_WINDOW, 60)
                Dim deliveryWin = Read_Txt(cells, r, COL_DELIVERY_WINDOW, 60)
                Dim bolNo = Read_Txt(cells, r, COL_BOL_NO, 60)
                Dim proNo = Read_Txt(cells, r, COL_PRO_NO, 60)
                Dim ediInv = Read_Txt(cells, r, COL_EDI_INV, 40)
                Dim ediAsn = Read_Txt(cells, r, COL_EDI_ASN, 40)
                Dim orderVal = If(COL_ORDER_VALUE >= 0, CellDec(cells(r, COL_ORDER_VALUE)), 0D)

                ASCMAIN1.sql =
                    "INSERT INTO ZORDSTA2 (" &
                    "ADS_DOC_SEQ_NO, ORDER_DATE, START_SHIP, REQ_DATE, ORDER_NO, SOLD_TO, CUSTOMER_NAME, " &
                    "CUST_ORDER_REF, CARRIER, IMPORT_WINDOW, PICKING_WINDOW, DELIVERY_WINDOW, BOL_NO, PRO_NO, " &
                    "EDI_INVOICE, EDI_ASN, ORDER_VALUE" &
                    ") VALUES (" &
                    SQL_Txt(ADS_DOC_SEQ_NO) & "," &
                    SQL_Date(dtOrder) & "," &
                    SQL_Date(dtStartShip) & "," &
                    SQL_Date(dtReq) & "," &
                    SQL_Txt(orderNo) & "," &
                    SQL_Txt(soldTo) & "," &
                    SQL_Txt(custName) & "," &
                    SQL_Txt(custPo) & "," &
                    SQL_Txt(carrier) & "," &
                    SQL_Txt(importWin) & "," &
                    SQL_Txt(pickingWin) & "," &
                    SQL_Txt(deliveryWin) & "," &
                    SQL_Txt(bolNo) & "," &
                    SQL_Txt(proNo) & "," &
                    SQL_Txt(ediInv) & "," &
                    SQL_Txt(ediAsn) & "," &
                    SQL_Num(orderVal) &
                    ")"

                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            Next

            ASCMAIN1.sql = "Update RSTXLSQ1 SET XLS_DOC_STATUS = '1', LAST_DATE = SYSDATE, LAST_OPER = " & SQL_Txt(ASCMAIN1.USER_ID) &
                           " WHERE XLS_DOC_SEQ_NO = " & SQL_Txt(XLS_DOC_SEQ_NO)
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            frm.CommitTrans($"ADS workbook imported to ZORDSTA2 (Doc {XLS_DOC_SEQ_NO}).")

            Try
                Dim adsArchiveFolder As String = RS_PARM_XLS_FOLDER.Replace("InBound", "ADS_ARCHIVE")
                If Not System.IO.Directory.Exists(adsArchiveFolder) Then
                    System.IO.Directory.CreateDirectory(adsArchiveFolder)
                End If

                Dim newFileName As String = $"ADS-{ADS_DOC_SEQ_NO}-{baseName}"
                Dim destPath As String = Unique_Dest_Path(adsArchiveFolder, newFileName)

                System.IO.File.Move(importFile, destPath)
            Catch moveEx As Exception
                ASCMAIN1.Record_Event("ZORDSTA2", XLS_DOC_SEQ_NO, "", Now, ASCMAIN1.USER_ID, "ADSM", "ADS file move warning: " & moveEx.Message, "")
            End Try

        Catch ex As Exception
            ASCMAIN1.Record_Event("ZORDSTA2", XLS_DOC_SEQ_NO, "", Now, ASCMAIN1.USER_ID, "ADSE", "ADS Import Error", "")
            ASCMAIN1.sql = "Update RSTXLSQ1 SET XLS_DOC_STATUS = 'E' WHERE XLS_DOC_SEQ_NO = " & SQL_Txt(XLS_DOC_SEQ_NO)
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            Throw
        End Try
    End Sub
    Private Shared Function Unique_Dest_Path(baseFolder As String, fileName As String) As String
        Dim dest As String = System.IO.Path.Combine(baseFolder, fileName)
        If Not System.IO.File.Exists(dest) Then Return dest
        Dim name As String = System.IO.Path.GetFileNameWithoutExtension(fileName)
        Dim ext As String = System.IO.Path.GetExtension(fileName)
        Dim i As Integer = 1
        Do
            Dim tryPath = System.IO.Path.Combine(baseFolder, $"{name} ({i}){ext}")
            If Not System.IO.File.Exists(tryPath) Then Return tryPath
            i += 1
        Loop
    End Function

    Private Shared Function CleanCustomerName(soldTo As String, nameIn As String) As String
        Dim code As String = (soldTo & "").Trim()

        Select Case code.ToUpperInvariant()
            Case "ABS CODE", "AAFESCOM", "JOY", "NORDDROP", "QVC",
             "SLIPCOMC", "SLIPCOMU", "WAYFAIR", "WESTELM",
             "POOSH", "POTTERBAR"
                Return Nothing
        End Select

        If String.IsNullOrWhiteSpace(nameIn) Then Return nameIn
        Dim s As String = nameIn.Trim()
        If s.Length > 200 Then s = s.Substring(0, 200)
        Return s
    End Function

    Private Shared Function Read_Header_Text(ws As SpreadsheetGear.IWorksheet, a1 As String, Optional a1Alt As String = Nothing) As String
        Dim t As String = Trim(CStr(ws.Cells(a1).Text))
        If t = "" AndAlso Not String.IsNullOrEmpty(a1Alt) Then
            t = Trim(CStr(ws.Cells(a1Alt).Text))
        End If
        Return If(t = "", Nothing, t)
    End Function

    Private Shared Function Parse_Header_Date(dateTxt As String) As Object
        If String.IsNullOrEmpty(dateTxt) Then Return DBNull.Value
        Dim d As DateTime
        If DateTime.TryParse(dateTxt, d) Then
            Return New DateTime(d.Year, d.Month, d.Day)
        End If
        Return DBNull.Value
    End Function


    Private Shared Function Read_Txt(c As SpreadsheetGear.IRange, r As Integer, col As Integer, maxLen As Integer) As String
        If col < 0 Then Return Nothing
        Dim t As String = Trim(CStr(c(r, col).Text))
        If t = "" Then Return Nothing
        If t.Length > maxLen Then t = t.Substring(0, maxLen)
        Return t
    End Function

    Private Shared Function Read_Date(c As SpreadsheetGear.IRange, r As Integer, col As Integer) As Object
        If col < 0 Then Return DBNull.Value
        Dim v As Object = c(r, col).Value
        Dim d As DateTime
        If TypeOf v Is Double OrElse TypeOf v Is Decimal Then
            Dim base As Date = #1/1/1900#
            Dim dd As Integer = CInt(Math.Truncate(Convert.ToDecimal(v)))
            If dd >= 1 Then Return base.AddDays(dd - 2)
        End If
        Dim t As String = Trim(CStr(c(r, col).Text))
        If t <> "" AndAlso Date.TryParse(t, d) Then Return d
        Return DBNull.Value
    End Function

    Private Shared Function SQL_Date(v As Object) As String
        If v Is Nothing OrElse v Is DBNull.Value Then Return "NULL"
        Dim d As DateTime = CDate(v)
        Return "DATE '" & d.ToString("yyyy-MM-dd") & "'"
    End Function

    Private Shared Function SQL_Num(n As Object) As String
        If n Is Nothing OrElse n Is DBNull.Value Then Return "NULL"
        Return Convert.ToDecimal(n, Globalization.CultureInfo.InvariantCulture).ToString(Globalization.CultureInfo.InvariantCulture)
    End Function

    Private Shared Function SQL_Txt(s As String) As String
        If String.IsNullOrEmpty(s) Then Return "NULL"
        Return SqlQ(s)
    End Function

End Class
