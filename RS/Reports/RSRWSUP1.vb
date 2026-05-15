Public Class RSRWSUP1
    Dim RSTWSUP1 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYW("RYW", ASCMAIN1.CYW, -300, 0, 0)
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        Create_Work_File()

        With dst
            ASCMAIN1.sql = "Select * from " & RSTWSUP1
            Create_TDA(.Tables.Add, "RSTWSUP1", "**", 0, False, "", 3)
            'With .Tables("RSTBOWS1").Columns
            '    .Add("INXIT_XWKS", GetType(System.Decimal), "ISNULL(TYQTYL1,0)+ISNULL(TYQTYL2,0)+ISNULL(TYQTYL3,0)")
            '    .Add("LYN10", GetType(System.Decimal), "ISNULL(RSLYN01,0)+ISNULL(RSLYN02,0)+ISNULL(RSLYN03,0)+ISNULL(RSLYN04,0)+ISNULL(RSLYN05,0)+ISNULL(RSLYN06,0)+ISNULL(RSLYN07,0)+ISNULL(RSLYN08,0)+ISNULL(RSLYN09,0)+ISNULL(RSLYN10,0)")
            '    'LY Next 10 wks SLS: LY-N01+LY-N02+LY-N03+LY-N04+LY-N05+LY-N06+LY-N07+LY-N08+LY-N09+LY-N10

            '    .Add("QTY_NEEDED_CALC", GetType(System.Decimal), "ISNULL(NXX_QTY,0)-ISNULL(QTY_EOW,0)-ISNULL(ORDR_QTY_OPEN,0)-ISNULL(SHIPW0,0)-ISNULL(SHIPW1,0)-ISNULL(SHIPW2,0)")
            '    .Add("QTY_NEEDED", GetType(System.Decimal), "IIF(QTY_NEEDED_CALC<0,0,QTY_NEEDED_CALC)")
            'End With

            ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, SREP_CODE, TRADE_CLASS_CODE" & vbCrLf _
                & " from ARTCUST1 where CUST_CODE in (Select Distinct CUST_CODE from " & RSTWSUP1 & ")"
            Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_NAME, ARTCUST2.CUST_STORE_STATE, ARTCUST2.SELL_CODE" & vbCrLf _
                & ", SOTSELL1.SELL_NAME, SOTSELL1.REGION_CODE, SOTSREG1.REGION_DESC" & vbCrLf _
                & " from ARTCUST2,SOTSELL1,SOTSREG1" & vbCrLf _
                & " where (ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO)" & vbCrLf _
                & " in (Select Distinct CUST_CODE, CUST_STORE_NO from " & RSTWSUP1 & ")" & vbCrLf _
                & "   and SOTSELL1.SELL_CODE (+) = ARTCUST2.SELL_CODE" & vbCrLf _
                & "   and SOTSREG1.REGION_CODE (+) = SOTSELL1.REGION_CODE"
            Create_TDA(.Tables.Add, "ARTCUST2", "**", 0, False, "", 2)

            'ASCMAIN1.sql = "Select ITEM_CODE, ITEM_DESC, ITEM_RETAIL_PRICE, COLLECTION_CODE, ITEM_EAN_CODE" & vbCrLf _
            '    & " from ICTITEM1 where ITEM_CODE in (Select Distinct ITEM_CODE from " & RSTWSUP1 & ")"
            'Create_TDA(.Tables.Add, "ICTITEM1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select COLLECTION_CODE, COLLECTION_NAME, BRAND_CODE" & vbCrLf _
                & " from ICTCOLL1 where COLLECTION_CODE in (Select Distinct COLLECTION_CODE from " & RSTWSUP1 & ")"
            Create_TDA(.Tables.Add, "ICTCOLL1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select * from SOTSELL1"
            Create_TDA(.Tables.Add, "SOTSELL1", "**", 0, False, "", 1)

        End With

        EnforceConstraints(False)

        Fill_Records("RSTWSUP1")


        Dim WKS As Integer = Val(numWksAvgSls.Value & "")
        Dim P As Integer = 0

        Fill_Records("ARTCUST1")
        Fill_Records("ARTCUST2")
        Fill_Records("ICTCOLL1")
        Fill_Records("SOTSELL1")

        EnforceConstraints(True)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Creating Workbook")


        Dim FACTOR As Decimal = Val(numFACTOR.Value & "")
        Dim WKSAVGSLS As Integer = Val(numWksAvgSls.Value & "")
        Dim WKSPOXIT As Integer = Val(numWksPOXit.Value & "")

        Dim D As String = optUR.Value 'Data Column Name Prefix QTY/RTL

        Dim xls_path As String = ASCMAIN1.Folders("Work")
        Dim xls_name As String = ""

        Dim FILENAME As String = ""

        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0

        Do Until success
            Try
                XLS_NO += 1
                xls_name = ASCMAIN1.DBS_SESSION_ID
                xls_name &= "-" & Format(XLS_NO, "000") & ".xlsx"
                FILENAME = xls_path & "\" & xls_name & ".XLSx"

                If Not My.Computer.FileSystem.FileExists(FILENAME) Then
                    success = True
                End If
            Catch ex As Exception
                Stop
            End Try
        Loop

        Dim oWB As SpreadsheetGear.IWorkbook
        oWB = SpreadsheetGear.Factory.GetWorkbook()

        Dim SI As Integer = 0
        Dim oSheet As SpreadsheetGear.IWorksheet

        oSheet = oWB.Sheets(SI)
        oSheet.Name = "Parameters"
        Dim tbl As New DataTable
        With tbl.Columns
            .Add("RTL_WEEK", GetType(System.String))
            .Add("WEEK_END_DATE", GetType(System.DateTime))
            .Add("WKSAVGSLS", GetType(System.Int32))
            .Add("WKSPOXIT", GetType(System.Int32))
            .Add("FACTOR", GetType(System.Int32))
            .Add("DATA")
        End With


        Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", RYW)
        tbl.Rows.Add(New Object() {RYW, rowGLTPARM3.Item("WEEK_END_DATE"), WKSAVGSLS, WKSPOXIT, FACTOR, D})
        Load_DataTable_into_SGXLS(1, 1, tbl, oSheet)

        ' BUILD XLS HERE

        'oSheet = oWB.Worksheets.Add()
        'Load_DataTable_into_SGXLS(1, 1, dst.Tables("RSTWSUP1"), oSheet)
        'oSheet.Name = "KitchenSink"


        Dim Rx As Integer = 0

        For Each rowARTCUST1 As DataRow In dst.Tables("ARTCUST1").Select("", "CUST_CODE")
            oSheet = oWB.Worksheets.Add()
            Dim CUST_CODE = rowARTCUST1.Item("CUST_CODE")
            Dim sqlCUST_CODE As String = "CUST_CODE = '" & CUST_CODE & "'"

            oSheet.Name = CUST_CODE
            Rx = 5

            Dim Cx As Integer = 0
            ' sheet headings
            Cx += 1 : oSheet.Cells(Rx, Cx).Value = "Customer"
            With oSheet.Cells(Rx, Cx, Rx, Cx + 5)
                .Interior.Color = SpreadsheetGear.Colors.LightGray
                .EntireColumn.NumberFormat = "@"
            End With

            Cx += 1 : oSheet.Cells(Rx, Cx).Value = "Store"
            Cx += 1 : oSheet.Cells(Rx, Cx).Value = "Store Name"
            oSheet.Cells(Rx, Cx).EntireColumn.ColumnWidth = 30
            Cx += 1 : oSheet.Cells(Rx, Cx).Value = "State"
            Cx += 1 : oSheet.Cells(Rx, Cx).Value = "AE"
            oSheet.Cells(Rx, Cx, Rx, Cx + 1).Interior.Color = SpreadsheetGear.Colors.Gold
            Cx += 1 : oSheet.Cells(Rx, Cx).Value = "AE Name"
            oSheet.Cells(Rx, Cx).EntireColumn.ColumnWidth = 30

            Cx += 1 : oSheet.Cells(Rx, Cx).Value = "Collection"
            With oSheet.Cells(Rx, Cx, Rx, Cx + 5)
                .Interior.Color = SpreadsheetGear.Colors.LightBlue
                .EntireColumn.NumberFormat = "@"
                .EntireColumn.ColumnWidth = 10
            End With

            For w As Integer = 13 To 0 Step -1
                Dim C As String = CStr(w) & "LW"
                If w = 1 Then C = "LW"
                If w = 0 Then C = "TW"
                Cx += 1 : oSheet.Cells(Rx, Cx).Value = C
                If w = 13 Then
                    With oSheet.Cells(Rx, Cx, Rx, Cx + 13)
                        .Interior.Color = SpreadsheetGear.Colors.LightGreen
                        .EntireColumn.NumberFormat = "#,##0"
                        .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    End With
                    With oSheet.Cells(Rx - 1, Cx, Rx - 1, Cx + 13)
                        .Merge()
                        .Interior.Color = SpreadsheetGear.Colors.LawnGreen

                        If Absx1.optFor("OPTIT").Value = "I" Then
                            .Value = "Shipping History, in " & optUR.Text
                        Else
                            .Value = "Sell-Thru (Retail) History, in " & optUR.Text
                        End If

                        .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    End With
                End If
                If w > Val(numWksAvgSls.Value) Then
                    oSheet.Cells(Rx, Cx).EntireColumn.Hidden = True
                End If
            Next

            Cx += 1 : oSheet.Cells(Rx, Cx).Value = "Avg Wk"
            With oSheet.Cells(Rx, Cx, Rx, Cx + 1)
                .Interior.Color = SpreadsheetGear.Colors.LightSeaGreen
                .EntireColumn.NumberFormat = "#,##0.00"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .EntireColumn.ColumnWidth = 10
            End With
            Cx += 1 : oSheet.Cells(Rx, Cx).Value = "Proj/Wk"

            Cx += 1 : oSheet.Cells(Rx, Cx).Value = "Open+Pick"
            With oSheet.Cells(Rx, Cx, Rx, Cx + 2)
                .Interior.Color = SpreadsheetGear.Colors.PeachPuff
                .EntireColumn.NumberFormat = "#,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .EntireColumn.ColumnWidth = 10
            End With
            Cx += 1 : oSheet.Cells(Rx, Cx).Value = "On Hand"
            Cx += 1 : oSheet.Cells(Rx, Cx).Value = "In Transit"

            Cx += 1 : oSheet.Cells(Rx, Cx).Value = "Total Supply"
            With oSheet.Cells(Rx, Cx, Rx, Cx)
                .Interior.Color = SpreadsheetGear.Colors.Violet
                .EntireColumn.NumberFormat = "#,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .EntireColumn.ColumnWidth = 10
            End With

            Cx += 1 : oSheet.Cells(Rx, Cx).Value = "WOS"
            With oSheet.Cells(Rx, Cx, Rx, Cx)
                .Interior.Color = SpreadsheetGear.Colors.Violet
                .EntireColumn.NumberFormat = "#,##0.0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .EntireColumn.ColumnWidth = 10
            End With


            For Each rowARTCUST2 As DataRow In dst.Tables("ARTCUST2").Select(sqlCUST_CODE, "CUST_STORE_NO")
                Dim CUST_STORE_NO As String = rowARTCUST2.Item("CUST_STORE_NO")
                Dim sqlCUST_STORE_NO As String = " and CUST_STORE_NO = '" & CUST_STORE_NO & "'"

                Dim CUST_STORE_NAME As String = rowARTCUST2.Item("CUST_STORE_NAME") & ""
                Dim CUST_STORE_STATE As String = rowARTCUST2.Item("CUST_STORE_STATE") & ""
                Dim SELL_CODE As String = rowARTCUST2.Item("SELL_CODE") & ""

                Dim rowSOTSELL1 As DataRow = dst.Tables("SOTSELL1").Rows.Find(SELL_CODE)
                Dim SELL_NAME As String = ""
                If rowSOTSELL1 IsNot Nothing Then
                    SELL_NAME = rowSOTSELL1.Item("SELL_NAME") & ""
                End If

                ' Rx += 1 ' this line will put a space before the start of each store

                For Each rowRSTWSUP1 As DataRow In dst.Tables("RSTWSUP1").Select(sqlCUST_CODE & sqlCUST_STORE_NO, "CUST_CODE, CUST_STORE_NO, COLLECTION_CODE")
                    Dim COLLECTION_CODE = rowRSTWSUP1.Item("COLLECTION_CODE")
                    Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)

                    Rx += 1

                    Cx = 0
                    Cx += 1 : oSheet.Cells(Rx, Cx).Value = CUST_CODE

                    Cx += 1 : oSheet.Cells(Rx, Cx).Value = CUST_STORE_NO
                    Cx += 1 : oSheet.Cells(Rx, Cx).Value = CUST_STORE_NAME
                    Cx += 1 : oSheet.Cells(Rx, Cx).Value = CUST_STORE_STATE
                    Cx += 1 : oSheet.Cells(Rx, Cx).Value = SELL_CODE
                    Cx += 1 : oSheet.Cells(Rx, Cx).Value = SELL_NAME

                    Cx += 1 : oSheet.Cells(Rx, Cx).Value = COLLECTION_CODE

                    For W As Integer = 13 To 0 Step -1
                        Dim V As Int32 = rowRSTWSUP1.Item(D & "_L" & Format(W, "00"))
                        Cx += 1 : oSheet.Cells(Rx, Cx).Value = V
                    Next

                    Dim CxTW As Integer = Cx ' TW column

                    Cx += 1 : oSheet.Cells(Rx, Cx).Formula = "=Sum(" & Excel_Cell0(Rx, CxTW - 1 - WKSAVGSLS + 1) & ":" & Excel_Cell0(Rx, CxTW - 1) & ")/" & CStr(WKSAVGSLS)
                    Cx += 1 : oSheet.Cells(Rx, Cx).Formula = "=" & Excel_Cell0(Rx, Cx - 1) & "*" & CStr(FACTOR)
                    Dim CXProj As Integer = Cx

                    Cx += 1 : oSheet.Cells(Rx, Cx).Value = Val(rowRSTWSUP1.Item(D & "_O") & "") + Val(rowRSTWSUP1.Item(D & "_P") & "")
                    Cx += 1 : oSheet.Cells(Rx, Cx).Value = Val(rowRSTWSUP1.Item(D & "_EOW") & "")

                    Cx += 1 : oSheet.Cells(Rx, Cx).Formula = "=Sum(" & Excel_Cell0(Rx, CxTW - WKSPOXIT) & ":" & Excel_Cell0(Rx, CxTW) & ")"

                    Cx += 1 : oSheet.Cells(Rx, Cx).Formula = "=Sum(" & Excel_Cell0(Rx, Cx - 3) & ":" & Excel_Cell0(Rx, Cx - 1) & ")"
                    Cx += 1 : oSheet.Cells(Rx, Cx).Formula = "=IFERROR(" & Excel_Cell0(Rx, Cx - 1) & " / " & Excel_Cell0(Rx, CXProj) & ",0)"

                Next

            Next
        Next


        oWB.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(FILENAME)

        oWB = Nothing

        Add_Document_to_ASTSPRF1(FILENAME)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")

    End Sub

    Overrides Sub Build_Report_File_Post_Process()

    End Sub

    Public Overrides Sub Print_Report()

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.cmbFor("RYW").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify a Reporting Week"
            End If
        End If
    End Sub

    Sub Create_Work_File()

        Get_SQL("*")

        Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", RYW)
        Dim RYP As String = rowGLTPARM3("YYYYPP")

        Dim LYW As String = ASCMAIN1.WEEK_CALC(RYW, -1)

        ' SOTINVH2 - Gross Shipments History by Week

        Dim sqlTYL13 As String = ""
        Dim sqlTYL13_0 As String = ""
        Dim sqlTYL13_Sum As String = ""
        For w As Integer = 0 To 13
            Dim YW As String = ASCMAIN1.Week_Calc(RYW, -1 * w)

            If Absx1.optFor("OPTIT").Value = "I" Then
                sqlTYL13 &= ", Sum (Decode(SOTINVH2.OPS_YYYYWW,'" & YW & "',ORDR_QTY_SHIP,0)) QTY_L" & Format(w, "00")
                sqlTYL13 &= ", Sum (Decode(SOTINVH2.OPS_YYYYWW,'" & YW & "',ORDR_QTY_SHIP,0) * SOTINVH2.ORDR_UNIT_PRICE) WSL_L" & Format(w, "00")
                sqlTYL13 &= ", Sum (Decode(SOTINVH2.OPS_YYYYWW,'" & YW & "',ORDR_QTY_SHIP,0) * SOTINVH2.ITEM_RETAIL_PRICE) RTL_L" & Format(w, "00") & vbCrLf
            Else
                sqlTYL13 &= ", Sum (Decode(RSTRETL1.OPS_YYYYWW,'" & YW & "',QTY_SOLD,0)) QTY_L" & Format(w, "00")
                sqlTYL13 &= ", Sum (Decode(RSTRETL1.OPS_YYYYWW,'" & YW & "',QTY_SOLD,0) * ICTITEM1.ITEM_RETAIL_PRICE * .6) WSL_L" & Format(w, "00")
                sqlTYL13 &= ", Sum (Decode(RSTRETL1.OPS_YYYYWW,'" & YW & "',AMT_SOLD,0)) RTL_L" & Format(w, "00") & vbCrLf
            End If

            sqlTYL13_0 &= ", 0 QTY_L" & Format(w, "00")
            sqlTYL13_0 &= ", 0 WSL_L" & Format(w, "00")
            sqlTYL13_0 &= ", 0 RTL_L" & Format(w, "00") & vbCrLf

            sqlTYL13_Sum &= ", Sum (QTY_L" & Format(w, "00") & ") QTY_L" & Format(w, "00")
            sqlTYL13_Sum &= ", Sum (WSL_L" & Format(w, "00") & ") WSL_L" & Format(w, "00")
            sqlTYL13_Sum &= ", Sum (RTL_L" & Format(w, "00") & ") RTL_L" & Format(w, "00") & vbCrLf
        Next

        ' cust_code
        ' collection_code
        ' ae sell_code
        ' basic_promo

        Dim ITEM_BASIC_PROMO As String = optBP.Value ' Dim ITEM_BASIC_PROMO = "B"

        ASCMAIN1.sql = "Select X.CUST_CODE, X.CUST_STORE_NO, X.COLLECTION_CODE" & vbCrLf _
            & sqlTYL13_Sum _
            & ", Sum (X.QTY_O) QTY_O, Sum (X.WSL_O) WSL_O, Sum (X.RTL_O) RTL_O, Sum (X.QTY_P) QTY_P, Sum (X.WSL_P) WSL_P, Sum (X.RTL_P) RTL_P" & vbCrLf _
            & ", Sum (X.QTY_EOW) QTY_EOW, Sum (X.WSL_EOW) WSL_EOW, Sum (X.RTL_EOW) RTL_EOW" & vbCrLf _
            & " from ARTCUST2, ICTCOLL1, (" & vbCrLf _
            & IIf(Absx1.optFor("OPTIT").Value = "I",
                "Select ICTITEM1.COLLECTION_CODE, SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO" & vbCrLf _
                & sqlTYL13 _
                & ", 0 QTY_O, 0 WSL_O, 0 RTL_O,  0 QTY_P, 0 WSL_P, 0 RTL_P" & vbCrLf _
                & ", 0 QTY_EOW, 0 WSL_EOW, 0 RTL_EOW" & vbCrLf _
                & " from SOTINVH2,ICTITEM1" & vbCrLf _
                & " where SOTINVH2.OPS_YYYYWW between '" & ASCMAIN1.Week_Calc(RYW, -13) & "' and '" & RYW & "'" & vbCrLf _
                & "   and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
                & "   and ICTITEM1.ITEM_BASIC_PROMO = '" & ITEM_BASIC_PROMO & "'" & vbCrLf _
                & "   and ICTITEM1.ITEM_SNU_CODE = 'S'" & vbCrLf _
                & SQLA_filter("CUST_CODE", "SOTINVH2") _
                & SQLA_filter("COLLECTION_CODE", "ICTITEM1") & vbCrLf _
                & "group by ICTITEM1.COLLECTION_CODE, SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO",
                "Select ICTITEM1.COLLECTION_CODE, RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO" & vbCrLf _
                & sqlTYL13 _
                & ", 0 QTY_O, 0 WSL_O, 0 RTL_O,  0 QTY_P, 0 WSL_P, 0 RTL_P" & vbCrLf _
                & ", 0 QTY_EOW, 0 WSL_EOW, 0 RTL_EOW" & vbCrLf _
                & " from RSTRETL1,ICTITEM1" & vbCrLf _
                & " where RSTRETL1.OPS_YYYYWW between '" & ASCMAIN1.Week_Calc(RYW, -13) & "' and '" & RYW & "'" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE" & vbCrLf _
                & "   and ICTITEM1.ITEM_BASIC_PROMO = '" & ITEM_BASIC_PROMO & "'" & vbCrLf _
                & "   and ICTITEM1.ITEM_SNU_CODE = 'S'" & vbCrLf _
                & SQLA_filter("CUST_CODE", "RSTRETL1") _
                & SQLA_filter("COLLECTION_CODE", "ICTITEM1") & vbCrLf _
                & "group by ICTITEM1.COLLECTION_CODE, RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO") & vbCrLf _
            & " union " & vbCrLf _
            & "Select ICTITEM1.COLLECTION_CODE, SOTORDR2.CUST_CODE, SOTORDR2.CUST_STORE_NO" & vbCrLf _
            & sqlTYL13_0 _
            & ", SUM (DECODE(SOTORDR2.ORDR_STATUS,'O',SOTORDR2.ORDR_QTY_OPEN,0)) QTY_O" & vbCrLf _
            & ", SUM (DECODE(SOTORDR2.ORDR_STATUS,'O',SOTORDR2.ORDR_QTY_OPEN,0) * SOTORDR2.ORDR_UNIT_PRICE) WSL_O" & vbCrLf _
            & ", SUM (DECODE(SOTORDR2.ORDR_STATUS,'O',SOTORDR2.ORDR_QTY_OPEN,0) * SOTORDR2.ITEM_RETAIL_PRICE) RTL_O" & vbCrLf _
            & ", SUM (DECODE(SOTORDR2.ORDR_STATUS,'P',SOTORDR2.ORDR_QTY_PICK,0)) QTY_P" & vbCrLf _
            & ", SUM (DECODE(SOTORDR2.ORDR_STATUS,'P',SOTORDR2.ORDR_QTY_PICK,0) * SOTORDR2.ORDR_UNIT_PRICE) WSL_P" & vbCrLf _
            & ", SUM (DECODE(SOTORDR2.ORDR_STATUS,'P',SOTORDR2.ORDR_QTY_PICK,0) * SOTORDR2.ITEM_RETAIL_PRICE) RTL_P" & vbCrLf _
            & ", 0 QTY_EOW, 0 WSL_EOW, 0 RTL_EOW" & vbCrLf _
            & " from SOTORDR2,ICTITEM1" & vbCrLf _
            & " where SOTORDR2.ORDR_STATUS BETWEEN 'O' AND 'P'" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE" & vbCrLf _
            & "   and ICTITEM1.ITEM_BASIC_PROMO = '" & ITEM_BASIC_PROMO & "'" & vbCrLf _
            & "   and ICTITEM1.ITEM_SNU_CODE = 'S'" & vbCrLf _
            & SQLA_filter("CUST_CODE", "SOTORDR2") _
            & SQLA_filter("COLLECTION_CODE", "ICTITEM1") & vbCrLf _
            & "group by ICTITEM1.COLLECTION_CODE, SOTORDR2.CUST_CODE, SOTORDR2.CUST_STORE_NO" & vbCrLf & vbCrLf _
            & " union " & vbCrLf _
            & "Select ICTITEM1.COLLECTION_CODE, RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO" & vbCrLf _
            & sqlTYL13_0 _
            & ", 0 QTY_O, 0 WSL_O, 0 RTL_O,  0 QTY_P, 0 WSL_P, 0 RTL_P" & vbCrLf _
            & ", Sum (RSTRETL1.QTY_EOW) QTY_EOW" & vbCrLf _
            & ", Sum (RSTRETL1.QTY_EOW * ICTITEM1.ITEM_RETAIL_PRICE * 0.6) WSL_EOW" & vbCrLf _
            & ", Sum (RSTRETL1.QTY_EOW * ICTITEM1.ITEM_RETAIL_PRICE) RTL_EOW" & vbCrLf _
            & " from RSTRETL1,ICTITEM1" & vbCrLf _
            & " where RSTRETL1.OPS_YYYYWW = '" & LYW & "'" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE" & vbCrLf _
            & "   and ICTITEM1.ITEM_BASIC_PROMO = '" & ITEM_BASIC_PROMO & "'" & vbCrLf _
            & "   and ICTITEM1.ITEM_SNU_CODE = 'S'" & vbCrLf _
            & SQLA_filter("CUST_CODE", "RSTRETL1") _
            & SQLA_filter("COLLECTION_CODE", "ICTITEM1") & vbCrLf _
            & " group by ICTITEM1.COLLECTION_CODE, RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO" & vbCrLf _
            & ") X" & vbCrLf _
            & " where ICTCOLL1.COLLECTION_CODE = X.COLLECTION_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_CODE = X.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_STORE_NO = X.CUST_STORE_NO" & vbCrLf _
            & "   and NVL(ARTCUST2.CUST_STORE_STATUS,'?') = 'A'" & vbCrLf _
            & "   group by X.CUST_CODE, X.CUST_STORE_NO, X.COLLECTION_CODE"

        ASCMAIN1.sql = "Select RSTWSUP1.* from (" & ASCMAIN1.sql & ") RSTWSUP1" & sql_TABLE_NAMEs & vbCrLf _
             & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN)


        RSTWSUP1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & RSTWSUP1 & " Add Primary Key (CUST_CODE,CUST_STORE_NO,COLLECTION_CODE)")

        'ASCMAIN1.sql = "" _
        '        & "DELETE FROM " & RSTWSUP1 & vbCrLf _
        '        & "WHERE (CUST_CODE, CUST_STORE_NO) IN (" & vbCrLf _
        '        & "SELECT DISTINCT CUST_CODE, CUST_STORE_NO FROM " & RSTWSUP1 & vbCrLf _
        '        & "MINUS" & vbCrLf _
        '        & "SELECT CUST_CODE, CUST_STORE_NO FROM ARTCUST2)"
        'ASCDATA1.ExecuteSQL()

    End Sub
End Class