Public Class SORALLO1

#Region "General Declarations"

    Dim ALLO_DATE1 As Date = Now.Date
    Dim ALLO_DATE2 As Date = Now.Date

    Dim SOTALLO1 As String
    Dim SOTALLO2 As String
    Dim SOTALLOZ As String
    Dim SOTALLO3 As String
    Dim SOTALLOY As String

    Dim sqlSOTALLOC As String
    Dim SOTALLOC As String
    'Dim SOTALLOC_I As String
    Public RECAP_FILE As String
    Dim ALLO_CTL_NOs As New List(Of String)


#End Region
    Dim CUST_CODE_RECAP As New Dictionary(Of String, Integer)
    Dim START_END_DATEs As New Dictionary(Of String, String)

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("ARTPARM1")
        Get_PARM("ICTPARM1")

        Absx1.dteFor("ALLO_DATE1").Value = ALLO_DATE1
        Absx1.dteFor("ALLO_DATE2").Value = ALLO_DATE2
    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim sqlw As String = ""
        Dim sqlCUST_CODE As String = ""

        sqlw &= SQL_in("ITEM_BASIC_PROMO", "ICTITEM1.ITEM_BASIC_PROMO") & vbCrLf
        sqlw &= SQL_in("ITEM_SNU_CODE", "ICTITEM1.ITEM_SNU_CODE") & vbCrLf
        sqlw &= SQL_in("ITEM_CODE", "SOTALLO1.ITEM_CODE") & vbCrLf
        sqlCUST_CODE &= SQL_in("CUST_CODE")  ' SQL_in("CUST_CODE", "SOTALLO2.CUST_CODE") & vbCrLf
        sqlw &= SQL_in("COLLECTION_CODE", "ICTITEM1.COLLECTION_CODE") & vbCrLf
        sqlw &= SQL_in("BRAND_CODE", "ICTCOLL1.BRAND_CODE") & vbCrLf
        sqlw &= SQL_in("PROD_CODE", "ICTITEM1.PROD_CODE") & vbCrLf
        sqlw &= SQL_in("COLLECTION_GENDER", "ICTCOLL1.COLLECTION_GENDER") & vbCrLf
        sqlw &= SQL_in("ALLO_GROUP_CODE", "SOTALLO1.ALLO_GROUP_CODE") & vbCrLf
        sqlw &= SQL_in("CHANNEL_CODE", "SOTTCLS1.CHANNEL_CODE") & vbCrLf

        Prepare_dst(True, New String() {sqlw, sqlCUST_CODE})

        Dim channelCodeFilters As New List(Of String)()
        Dim regex As New System.Text.RegularExpressions.Regex("(SOTTCLS1\.CHANNEL_CODE\s*(?:[\<\>\=]+\s*'[^']+'|IN\s*\([^\)]+\)|NOT IN\s*\([^\)]+\)))", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        Dim matches As System.Text.RegularExpressions.MatchCollection = regex.Matches(sqlw)
        For Each match As System.Text.RegularExpressions.Match In matches
            channelCodeFilters.Add(match.Value)
        Next

        If channelCodeFilters.Count > 0 Then
            Dim combinedFilter As String = String.Join(" AND ", channelCodeFilters)
            ASCMAIN1.sql = "SELECT DISTINCT CUST_CODE FROM ARTCUST1, SOTTCLS1 WHERE ARTCUST1.TRADE_CLASS_CODE = SOTTCLS1.TRADE_CLASS_CODE AND " & combinedFilter
            Dim customerTable As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

            If customerTable.Rows.Count > 0 Then
                Dim customerCodes As String = String.Join(",", customerTable.AsEnumerable().Select(Function(row) "'" & row.Field(Of String)("CUST_CODE") & "'").ToArray())
                sqlCUST_CODE &= " AND CUST_CODE IN (" & customerCodes & ")"
            End If
        End If

        If sqlCUST_CODE <> "" Then
            ASCDATA1.DeleteRows("SOTALLOZ", "NOT (" & Mid(sqlCUST_CODE, 6) & ")")
            'For Each rowSOTALLOZ As DataRow In dst.Tables("SOTALLOZ").Select("")
            '    rowSOTALLOZ.Item("HIDE") = "1"
            'Next
            'For Each rowSOTALLOZ As DataRow In dst.Tables("SOTALLOZ").Select(Mid(sqlCUST_CODE, 6))
            '    rowSOTALLOZ.Item("HIDE") = "0"
            'Next
        End If


        If Not remotely_controlled Then
            grdSOTALLOC.DataSource = dst.Tables("SOTALLOC")
            Dim tblARTCUST1 As DataTable = dst.Tables("ARTCUST1").Copy
            If sqlCUST_CODE <> "" Then
                tblARTCUST1.DefaultView.RowFilter = Mid(sqlCUST_CODE, 5)
            End If
            grdARTCUST1.DataSource = tblARTCUST1 ' dst.Tables("ARTCUST1")
            grdSOTALLOX.DataSource = dst.Tables("SOTALLOX")
            grdSOTALLOX.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand
            grdSOTALLOX.DisplayLayout.Bands(0).Columns("QTY_ALLO_RPT").Header.Caption = "Qty Allo Rpt"
            grdSOTALLOX.DisplayLayout.Bands(0).Columns("QTY_ALLO_RPT").Header.VisiblePosition = grdSOTALLOX.DisplayLayout.Bands(0).Columns("BALANCE").Header.VisiblePosition + 1

            grdSOTALLOS.DataSource = dst.Tables("SOTALLOS")

            Dim tblARTCUST2 As DataTable = dst.Tables("ARTCUST2").Copy
            If sqlCUST_CODE <> "" Then
                tblARTCUST2.DefaultView.RowFilter = Mid(sqlCUST_CODE, 5)
            End If
            grdARTCUST2.DataSource = tblARTCUST2 ' dst.Tables("ARTCUST2")

            If grdSOTALLOC.DisplayLayout.ValueLists.Count = 0 Then
                ASCMAIN1.Add_Value_List(grdSOTALLOC, "ITEM_BASIC_PROMO")
                ASCMAIN1.Add_Value_List(grdSOTALLOC, "ITEM_SNU_CODE")
                '  ASCMAIN1.Add_Value_List(grdSOTALLOC, "PROD_CODE")
            End If

            If grdSOTALLOS.DisplayLayout.ValueLists.Count = 0 Then
                ASCMAIN1.Add_Value_List(grdSOTALLOS, "ITEM_BASIC_PROMO")
                ASCMAIN1.Add_Value_List(grdSOTALLOS, "ITEM_SNU_CODE")
                '  ASCMAIN1.Add_Value_List(grdSOTALLOS, "PROD_CODE")
            End If

            Dim COLUMN_NAMEs_to_show() As String = {"ITEM_CODE", "QTY_ALLO", "DATE_START", "DATE_END", "ALLO_GROUP_CODE",
                                            "ITEM_DESC", "ITEM_EAN_CODE", "ITEM_RETAIL_PRICE", "ITEM_SO_QTY_MULT",
                                            "ITEM_SNU_CODE", "ITEM_BASIC_PROMO", "PROD_CODE", "COLLECTION_CODE", "BRAND_CODE",
                                             "AMT_ALLO",
                                             "TY_SI_QTY", "TY_ST_QTY", "TY_ST_AMT", "TY_SIST",
                                             "LY_SI_QTY", "LY_ST_QTY", "LY_ST_AMT", "LY_SIST", "TY_LY_CHG",
                                             "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "BALANCE_OPEN", "ALLO_NOTES", "CUST_01", "CUST_02",
                                             "CUST_03", "CUST_04", "CUST_05", "CUST_06", "CUST_07", "CUST_08", "CUST_09", "CUST_10"}

            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTALLOC, grdSOTALLOS}
                With grd.DisplayLayout.Bands(0)
                    Dim gcol As UltraWinGrid.UltraGridColumn
                    For Each gcol In .Columns
                        gcol.Hidden = Not COLUMN_NAMEs_to_show.Contains(gcol.Key)
                    Next

                    Dim VP As Integer = -1

                    With .Columns("BRAND_CODE")
                        .Header.Caption = "Brand"
                        .Width = 60
                        VP += 1 : .Header.VisiblePosition = VP
                    End With
                    With .Columns("COLLECTION_CODE")
                        .Header.Caption = "Cllctn"
                        .Width = 80
                        VP += 1 : .Header.VisiblePosition = VP
                    End With

                    With .Columns("ITEM_CODE")
                        .Header.Caption = "Item Code"
                        .Width = 120
                        VP += 1 : .Header.VisiblePosition = VP
                    End With

                    With .Columns("ITEM_EAN_CODE")
                        .Header.Caption = "EAN/UPC"
                        .Width = 150
                        VP += 1 : .Header.VisiblePosition = VP
                    End With

                    With .Columns("DATE_START")
                        .Header.Caption = "Start"
                        .Width = 70
                        .Format = "MM/dd"
                        VP += 1 : .Header.VisiblePosition = VP
                    End With
                    With .Columns("DATE_END")
                        .Header.Caption = "End"
                        .Width = 70
                        .Format = "MM/dd"
                        VP += 1 : .Header.VisiblePosition = VP
                    End With
                    With .Columns("ALLO_GROUP_CODE")
                        .Header.Caption = "Group"
                        .Width = 130
                        VP += 1 : .Header.VisiblePosition = VP
                    End With
                    With .Columns("ITEM_DESC")
                        .Header.Caption = "Description"
                        .Width = 200
                        VP += 1 : .Header.VisiblePosition = VP
                    End With

                    With .Columns("ITEM_RETAIL_PRICE")
                        .Header.Caption = "Retail"
                        .Width = 80
                        .Format = "$#,##0.00"
                        VP += 1 : .Header.VisiblePosition = VP
                    End With
                    With .Columns("ITEM_SO_QTY_MULT")
                        .Header.Caption = "Mult"
                        .Width = 50
                        VP += 1 : .Header.VisiblePosition = VP
                    End With
                    With .Columns("ITEM_SNU_CODE")
                        .Header.Caption = "SN"
                        .Width = 60
                        VP += 1 : .Header.VisiblePosition = VP
                    End With
                    With .Columns("ITEM_BASIC_PROMO")
                        .Header.Caption = "BP"
                        .Width = 60
                        VP += 1 : .Header.VisiblePosition = VP
                    End With
                    With .Columns("PROD_CODE")
                        .Header.Caption = "Prod"
                        .Width = 90
                        VP += 1 : .Header.VisiblePosition = VP
                    End With

                    If grd.Name = "grdSOTALLOS" Then
                        With .Columns("CUST_STORE_NO")
                            .Header.Caption = "Store"
                            .Width = 70
                            .Hidden = False
                            VP += 1 : .Header.VisiblePosition = VP
                        End With
                    End If

                    With .Columns("QTY_ALLO")
                        .Header.Caption = "Qty Allo"
                        .Width = 70
                        .Format = "#,##0"
                        VP += 1 : .Header.VisiblePosition = VP
                    End With

                    With .Columns("AMT_ALLO")
                        .Header.Caption = "Amt Allo"
                        .Width = 70
                        .Format = "#,##0"
                        VP += 1 : .Header.VisiblePosition = VP
                    End With

                    With .Columns("TY_SI_QTY")
                        .Header.Caption = "TY SI Units"
                        .Width = 90
                        .Format = "#,##0"
                        VP += 1 : .Header.VisiblePosition = VP
                    End With
                    With .Columns("TY_ST_QTY")
                        .Header.Caption = "TY ST Units"
                        .Width = 90
                        .Format = "#,##0"
                        VP += 1 : .Header.VisiblePosition = VP
                    End With
                    With .Columns("TY_ST_AMT")
                        .Header.Caption = "TY ST $Amt"
                        .Width = 90
                        .Format = "#,##0"
                        VP += 1 : .Header.VisiblePosition = VP
                    End With
                    With .Columns("TY_SIST")
                        .Header.Caption = "TY%SThru"
                        .Width = 60
                        .Format = "#,##0.0"
                        VP += 1 : .Header.VisiblePosition = VP
                    End With

                    With .Columns("LY_SI_QTY")
                        .Header.Caption = "LY SI Units"
                        .Width = 90
                        .Format = "#,##0"
                        VP += 1 : .Header.VisiblePosition = VP
                    End With
                    With .Columns("LY_ST_QTY")
                        .Header.Caption = "LY ST Units"
                        .Width = 90
                        .Format = "#,##0"
                        VP += 1 : .Header.VisiblePosition = VP
                    End With
                    With .Columns("LY_ST_AMT")
                        .Header.Caption = "LY ST $Amt"
                        .Width = 90
                        .Format = "#,##0"
                        VP += 1 : .Header.VisiblePosition = VP
                    End With
                    With .Columns("LY_SIST")
                        .Header.Caption = "LY%SThru"
                        .Width = 60
                        .Format = "#,##0.0"
                        VP += 1 : .Header.VisiblePosition = VP
                    End With

                    With .Columns("TY_LY_CHG")
                        .Header.Caption = "AlloVsLY %Chg"
                        .Width = 60
                        .Hidden = True
                        .Format = "#,##0.0"
                        VP += 1 : .Header.VisiblePosition = VP
                    End With

                    With .Columns("ORDR_QTY_OPEN")
                        .Header.Caption = "Open Units"
                        .Width = 90
                        .Format = "#,##0"
                        VP += 1 : .Header.VisiblePosition = VP
                    End With
                    With .Columns("ORDR_QTY_PICK")
                        .Header.Caption = "Pick Units"
                        .Width = 90
                        .Format = "#,##0"
                        VP += 1 : .Header.VisiblePosition = VP
                    End With
                    With .Columns("BALANCE_OPEN")
                        .Header.Caption = "Bal Open"
                        .Width = 90
                        .Format = "#,##0"
                        VP += 1 : .Header.VisiblePosition = VP
                    End With
                    With .Columns("ALLO_NOTES")
                        .Header.Caption = "Notes"
                        .Width = 90
                        VP += 1 : .Header.VisiblePosition = VP
                    End With
                End With
            Next

            Sort_grdColumns(grdARTCUST1, "CUST_CODE")
            Setup_grdSOTALLOC()

            Sort_grdColumns(grdARTCUST2, "CUST_CODE,CUST_STORE_NO")
            Setup_grdSOTALLOS()

            If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
                With grdSOTALLOX.DisplayLayout.Bands(0)
                    .Columns("BRAND_CODE").Hidden = True
                    .Columns("COLLECTION_CODE").Hidden = True
                    .Columns("ITEM_EAN_CODE").Hidden = True
                End With
                With grdSOTALLOC.DisplayLayout.Bands(0)
                    .Columns("BRAND_CODE").Hidden = True
                    .Columns("COLLECTION_CODE").Hidden = True
                    .Columns("ITEM_EAN_CODE").Hidden = True
                End With
                With grdSOTALLOS.DisplayLayout.Bands(0)
                    .Columns("BRAND_CODE").Hidden = True
                    .Columns("COLLECTION_CODE").Hidden = True
                    .Columns("ITEM_EAN_CODE").Hidden = True
                End With
            Else
                With grdSOTALLOX.DisplayLayout.Bands(0)
                    .Columns("ITEM_UPC_CODE").Hidden = True
                End With
                With grdSOTALLOC.DisplayLayout.Bands(0)
                    .Columns("ITEM_UPC_CODE").Hidden = True
                End With
                With grdSOTALLOS.DisplayLayout.Bands(0)
                    .Columns("ITEM_UPC_CODE").Hidden = True
                End With
            End If


            If chkCHKEXCEL.Checked Then

                Dim myWorkbook As GemBox.Spreadsheet.ExcelFile = Nothing
                CUST_CODE_RECAP.Clear()
                Dim LNO_RECAP As Integer

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Creating Workbook")
                For Each grow As UltraWinGrid.UltraGridRow In grdARTCUST1.Rows
                    grow.Activate()
                    Dim CUST_CODE As String = grow.Cells("CUST_CODE").Value
                    grdSOTALLOC.Text = CUST_CODE
                    Application.DoEvents()
                    If myWorkbook Is Nothing Then
                        myWorkbook = Gembox_Export_to_Excel(grdSOTALLOC, False, False, "", "X")
                        'myWorksheet = myWorkbook.Worksheets(0)
                    Else
                        Gembox_Export_to_Excel_Add_grd(myWorkbook, grdSOTALLOC, False, "", "X")
                    End If
                    If CUST_CODE_RECAP.ContainsKey(CUST_CODE) Then
                    Else
                        CUST_CODE_RECAP.Add(CUST_CODE, LNO_RECAP + 1)
                    End If
                    LNO_RECAP += 1

                Next

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading Workbook")
                Dim xlsFileName As String = GetFileName(myWorkbook, "SORALLO1", ".xlsx")

                ' Gembox_Export_to_Excel_Show(myWorkbook, , ".xlsx")

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Creating Workbook")
                For Each grow As UltraWinGrid.UltraGridRow In grdARTCUST2.Rows
                    grow.Activate()
                    Dim CUST_CODE As String = grow.Cells("CUST_CODE").Value
                    Dim CUST_STORE_NO As String = grow.Cells("CUST_STORE_NO").Value
                    Dim CUST_STORE_NAME As String = grow.Cells("CUST_STORE_NAME").Value & ""
                    grdSOTALLOS.Text = CUST_CODE & "-" & CUST_STORE_NO & " " & CUST_STORE_NAME
                    If grdSOTALLOS.Rows.Count <> 0 Then
                        Application.DoEvents()
                        If myWorkbook Is Nothing Then
                            myWorkbook = Gembox_Export_to_Excel(grdSOTALLOS, False, False, "", "X")
                            'myWorksheet = myWorkbook.Worksheets(0)
                        Else
                            Gembox_Export_to_Excel_Add_grd(myWorkbook, grdSOTALLOS, False, "", "X")
                        End If
                    End If
                Next
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading Workbook")
                If grdARTCUST2.Rows.Count <> 0 Then
                    ' Gembox_Export_to_Excel_Show(myWorkbook, , ".xlsx")
                    Gembox_Export_to_Excel_Show(myWorkbook, "SORALLO1_CUST_STORE", ".xlsx")
                End If


                '  Sort_grdColumns(grdSOTALLOX, "ITEM_CODE")
                Sort_grdColumns(grdSOTALLOX, "BRAND_CODE,COLLECTION_CODE,ITEM_SNU_CODE,PROD_CODE,ITEM_CODE")
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Creating Workbook")
                myWorkbook = Gembox_Export_to_Excel(grdSOTALLOX, False, False, "", "X")
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading Workbook")
                'Gembox_Export_to_Excel_Show(myWorkbook, , ".xlsx")
                Gembox_Export_to_Excel_Show(myWorkbook, "SORALLO1_SUMMARY", ".xlsx")

                GENERATE_RECAP()

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
            End If

            Sort_grdColumns(grdARTCUST1, "CUST_CODE")
            Setup_grdSOTALLOC()

            Sort_grdColumns(grdARTCUST2, "CUST_CODE,CUST_STORE_NO")
            Setup_grdSOTALLOS()

            Sort_grdColumns(grdSOTALLOX, "ITEM_CODE")


            If ASCMAIN1.CLIENT = "INT" Then
                grdSOTALLOX.DisplayLayout.Bands(0).Columns("ITEM_EAN_CODE").Header.Caption = "EAN/UPC"
                grdSOTALLOS.DisplayLayout.Bands(0).Columns("ITEM_EAN_CODE").Header.Caption = "EAN/UPC"
                grdSOTALLOC.DisplayLayout.Bands(0).Columns("ITEM_EAN_CODE").Header.Caption = "EAN/UPC"
            End If


            tabAllocations.Visible = True
        End If



        For Each row As DataRow In dst.Tables("SOTALLOX").Select("")
            Dim ALLO_CTL_NO As String = row.Item("ALLO_CTL_NO")
            Dim QTY_ALLO_RPT As Int64 = 0
            For Each rowZ As DataRow In dst.Tables("SOTALLOZ").Select("ALLO_CTL_NO = '" & ALLO_CTL_NO & "'")
                Dim QTY_ALLO_RPT_Z As Int64 = Val(rowZ.Item("QTY_ALLO_RPT") & "")
                QTY_ALLO_RPT += QTY_ALLO_RPT_Z
            Next
            row.Item("QTY_ALLO_RPT") = QTY_ALLO_RPT
        Next

        'Create_Relation("SOTALLOX", "SOTALLOZ", "ALLO_CTL_NO")
        '.Tables("SOTALLOX").Columns("QTY_ALLO_RPT").Expression = "SUM(CHILD.QTY_ALLO_RPT)"

        Check_if_Empty("SOTALLO1")
    End Sub

    Public Overrides Sub Print_Report()

        If Format(ALLO_DATE1, "MM/dd/yyyy") = Format(ALLO_DATE2, "MM/dd/yyyy") Then
            SUBT = "Allocations Active on " & Format(ALLO_DATE1, "MM/dd/yy")
        Else
            SUBT = "Allocations Active from " & Format(ALLO_DATE1, "MM/dd/yy") & " to " & Format(ALLO_DATE2, "MM/dd/yy")
        End If
        If Absx1.chkFor("CHKEXC_ONLY").Checked Then
            SUBT &= ", Exceeded Only"
        End If

        CR_params.Add("SUBT", "Item/Customer")
        CR_params.Add("PAGE_EJECT", IIf(Absx1.chkFor("CHKPAGE_EJECT").Checked, "1", "0"))
        CR_params.Add("EXC_ONLY", IIf(Absx1.chkFor("CHKEXC_ONLY").Checked, "1", "0"))
        CR_params.Add("SUMMARY", "0")
        Generate_Report(RPT, , SUBT)

        CR_params.Add("SUBT", "Item Summary")
        CR_params.Add("PAGE_EJECT", IIf(Absx1.chkFor("CHKPAGE_EJECT").Checked, "1", "0"))
        CR_params.Add("EXC_ONLY", IIf(Absx1.chkFor("CHKEXC_ONLY").Checked, "1", "0"))
        CR_params.Add("SUMMARY", "1")
        Generate_Report(RPT, , SUBT)

        CR_params.Add("SUBT", "Customer/Item")
        CR_params.Add("PAGE_EJECT", IIf(Absx1.chkFor("CHKPAGE_EJECT").Checked, "1", "0"))
        CR_params.Add("EXC_ONLY", IIf(Absx1.chkFor("CHKEXC_ONLY").Checked, "1", "0"))
        RPT = "SORALLO2"
        Generate_Report(RPT, , SUBT)

        If chkCHKEXCEL_STORE.Checked Then
            Create_Pivot_by_Store()
        End If
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            If Format(ALLO_DATE1, "yyyyMMdd") > Format(ALLO_DATE2, "yyyyMMdd") Then
                EMsg &= vbCr & "From Date must not be later than To Date"
            End If

            If EMsg = "" Then
                ALLO_DATE1 = Absx1.dteFor("ALLO_DATE1").Value
                ALLO_DATE2 = Absx1.dteFor("ALLO_DATE2").Value
            End If

        End If
    End Sub

    Overrides Function Prepare_dst(
    ByVal perform_fill As Boolean,
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Get_PARM("SOTPARM1")
        Get_PARM("ARTPARM1")

        Dim sqlw As String = CStr(parms(0))
        Dim sqlwCUST_CODE As String = CStr(parms(1))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        Create_Temp_Tables(sqlw, sqlwCUST_CODE)

        With dst

            ASCMAIN1.sql = "Select SOTALLO1.*, X.QTY_ALLO_TOTAL from " & SOTALLO1 & " SOTALLO1" & vbCrLf _
                & ", (Select ALLO_CTL_NO, SUM (QTY_ALLO) QTY_ALLO_TOTAL" & vbCrLf _
                & " from SOTALLO2 where ALLO_CTL_NO in (Select ALLO_CTL_NO from " & SOTALLO1 & ")" & vbCrLf _
                & " group by ALLO_CTL_NO) X where X.ALLO_CTL_NO (+) = SOTALLO1.ALLO_CTL_NO"
            Create_TDA(.Tables.Add, "SOTALLO1", "**", 0, False, "", 1)
            With dst.Tables("SOTALLO1").Columns
                .Add("BALANCE", GetType(System.Int64))
                '.Add("AVA2SHIP", GetType(System.Int64))
                '.Add("AVA2ALLO", GetType(System.Int64))
            End With

            ASCMAIN1.sql = "Select ICTITEM1.* from ICTITEM1" & vbCrLf _
                & " where ITEM_CODE in (Select Distinct ITEM_CODE from " & SOTALLO1 & " SOTALLO1)"
            Create_TDA(.Tables.Add, "ICTITEM1", "**", 0, False, "", 1)
            .Tables("ICTITEM1").Columns.Add("ITEM_IMAGE", GetType(System.Byte()))

            ASCMAIN1.sql = "Select SOTALLO2.* from " & SOTALLO2 & " SOTALLO2 where SOTALLO2.ALLO_CTL_NO in (Select Distinct ALLO_CTL_NO from " & SOTALLO1 & ")"
            Create_TDA(.Tables.Add, "SOTALLO2", "**", 0, False, "", 2)

            Create_Relation("SOTALLO1", "SOTALLO2", "ALLO_CTL_NO")

            ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME from ARTCUST1" _
                & " where CUST_CODE in " & vbCrLf _
                & " (Select Distinct CUST_CODE from " & SOTALLO2 & " SOTALLO2" & vbCrLf _
                & " union" & vbCrLf _
                & "  Select Distinct CUST_CODE from " & SOTALLOZ & ")"
            Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO, ARTCUST2.CUST_STORE_NAME" & vbCrLf _
                & " from ARTCUST1,ARTCUST2" _
                & " where ARTCUST2.CUST_CODE = ARTCUST1.CUST_CODE" & vbCrLf _
                & "   and ARTCUST1.CUST_ALLOCATE_BY_STORE = '1'" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE in " & vbCrLf _
                & " (Select Distinct CUST_CODE from " & SOTALLO2 & " SOTALLO2" & vbCrLf _
                & " union" & vbCrLf _
                & "  Select Distinct CUST_CODE from " & SOTALLOZ & ")"
            Create_TDA(.Tables.Add, "ARTCUST2", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select ICTSTAT2.ITEM_CODE" & vbCrLf _
                 & ", SUM (ICTSTAT2.WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND" & vbCrLf _
                 & ", SUM (ICTSTAT2.WHSE_QTY_ONPO) WHSE_QTY_ONPO" & vbCrLf _
                 & ", SUM (ICTSTAT2.WHSE_QTY_PLAN) WHSE_QTY_PLAN" & vbCrLf _
                 & ", SUM (ICTSTAT2.WHSE_QTY_OPEN) WHSE_QTY_OPEN" & vbCrLf _
                 & ", SUM (ICTSTAT2.WHSE_QTY_PICK) WHSE_QTY_PICK" & vbCrLf _
                 & ", SUM (ICTSTAT2.WHSE_QTY_COMM) WHSE_QTY_COMM" & vbCrLf _
                 & "  from ICTSTAT2,ICTWHSE1" & vbCrLf _
                 & " where ICTSTAT2.ITEM_CODE in (Select Distinct ITEM_CODE from " & SOTALLO1 & " SOTALLO1)" & vbCrLf _
                 & "   and ICTWHSE1.WHSE_CODE = ICTSTAT2.WHSE_CODE" & vbCrLf _
                 & "   and NVL(ICTWHSE1.WHSE_MRP_EXC_IND,'0') <> '1'" & vbCrLf _
                 & " group by ICTSTAT2.ITEM_CODE"
            Create_TDA(.Tables.Add, "ICTSTAT2", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select SOTALLOZ.* from " & SOTALLOZ & " SOTALLOZ"
            Create_TDA(.Tables.Add, "SOTALLOZ", "**", 0, False, "", 2)
            dst.Tables("SOTALLOZ").Columns.Add("HIDE")
            dst.Tables("SOTALLOZ").Columns.Add("QTY_ALLO_RPT", GetType(System.Int64))

            sqlSOTALLOC = "Select SOTALLOZ.*, SOTALLO2.QTY_ALLO, SOTALLO2.ALLO_NOTES" & vbCrLf _
                & ", SOTALLO1.ITEM_CODE, SOTALLO1.DATE_START, SOTALLO1.DATE_END, SOTALLO1.ALLO_GROUP_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_UPC_CODE, NVL(ICTITEM1.ITEM_EAN_CODE,ICTITEM1.ITEM_UPC_CODE) ITEM_EAN_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_SO_QTY_MULT, ICTITEM1.ITEM_DATE_TO_SHIP" & vbCrLf _
                & ", ICTITEM1.COST_CATGY_CODE, ICTITEM1.ITEM_SNU_CODE, ICTITEM1.ITEM_BASIC_PROMO" & vbCrLf _
                & ", ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.ITEM_STATUS" & vbCrLf _
                & ", ICTITEM1.PROD_CODE, ICTITEM1.COLLECTION_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
                & " from " & SOTALLOZ & " SOTALLOZ, SOTALLO1, SOTALLO2, ICTITEM1, ICTCOLL1" & vbCrLf _
                & " where SOTALLO1.ALLO_CTL_NO = SOTALLOZ.ALLO_CTL_NO" & vbCrLf _
                & "   and SOTALLO2.ALLO_CTL_NO (+) = SOTALLOZ.ALLO_CTL_NO" & vbCrLf _
                & "   and SOTALLO2.CUST_CODE (+) = SOTALLOZ.CUST_CODE" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = SOTALLO1.ITEM_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE"
            Create_SOTALLOC("")

            ASCMAIN1.sql = "Select * from " & SOTALLOC
            Create_TDA(.Tables.Add, "SOTALLOC", "**", 0, False, "", 2)
            With dst.Tables("SOTALLOC").Columns
                .Add("BALANCE_CALC", GetType(System.Int64), "ISNULL(QTY_ALLO,0)-ISNULL(ORDR_QTY_SHIP,0)-ISNULL(ORDR_QTY_PICK,0)")
                .Add("BALANCE", GetType(System.Int64), "IIF(BALANCE_CALC<0,0,BALANCE_CALC)")
                .Add("AMT_ALLO", GetType(System.Int64), "QTY_ALLO * ITEM_RETAIL_PRICE")
                .Add("TY_SI_QTY", GetType(System.Int64))
                .Add("TY_SI_AMT", GetType(System.Decimal))
                .Add("LY_SI_QTY", GetType(System.Int64))
                .Add("LY_SI_AMT", GetType(System.Decimal))
                .Add("TY_ST_QTY", GetType(System.Int64))
                .Add("TY_ST_AMT", GetType(System.Decimal))
                .Add("LY_ST_QTY", GetType(System.Int64))
                .Add("LY_ST_AMT", GetType(System.Decimal))
                .Add("TY_SIST", GetType(System.Decimal), "IIF(ISNULL(TY_SI_QTY,0)=0,0,100*ISNULL(TY_ST_QTY,0)/ISNULL(TY_SI_QTY,0))")
                .Add("LY_SIST", GetType(System.Decimal), "IIF(ISNULL(LY_SI_QTY,0)=0,0,100*ISNULL(LY_ST_QTY,0)/ISNULL(LY_SI_QTY,0))")
                .Add("TY_LY_CHG", GetType(System.Decimal), "IIF(ISNULL(LY_ST_QTY,0)=0,0,100*(ISNULL(QTY_ALLO,0)/ISNULL(LY_ST_QTY,0)-1))")
                .Add("BALANCE_OPEN", GetType(System.Int64), "ISNULL(QTY_ALLO,0)-ISNULL(TY_SI_QTY,0)-ISNULL(ORDR_QTY_OPEN,0)-ISNULL(ORDR_QTY_PICK,0)")
                For i As Integer = 1 To 10
                    Dim COL_NAME As String = $"CUST_{i:00}"
                    If Not .Contains(COL_NAME) Then
                        .Add(COL_NAME, GetType(Int64))
                    End If
                Next
            End With


            ASCMAIN1.sql = "Select SOTALLOY.* from " & SOTALLOY & " SOTALLOY"
            Create_TDA(.Tables.Add, "SOTALLOY", "**", 0, False, "", 3)

            ASCMAIN1.sql = "Select SOTALLOY.*, SOTTCLS1.CHANNEL_CODE, SOTALLO3.QTY_ALLO, SOTALLO3.ALLO_NOTES" & vbCrLf _
                & ", SOTALLO1.ITEM_CODE, SOTALLO1.DATE_START, SOTALLO1.DATE_END, SOTALLO1.ALLO_GROUP_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_UPC_CODE, NVL(ICTITEM1.ITEM_EAN_CODE,ICTITEM1.ITEM_UPC_CODE) ITEM_EAN_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_SO_QTY_MULT, ICTITEM1.ITEM_DATE_TO_SHIP" & vbCrLf _
                & ", ICTITEM1.COST_CATGY_CODE, ICTITEM1.ITEM_SNU_CODE, ICTITEM1.ITEM_BASIC_PROMO" & vbCrLf _
                & ", ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.ITEM_STATUS" & vbCrLf _
                & ", ICTITEM1.PROD_CODE, ICTITEM1.COLLECTION_CODE, ICTCOLL1.BRAND_CODE, ARTCUST2.CUST_STORE_NAME" & vbCrLf _
                & " from " & SOTALLOY & " SOTALLOY, SOTALLO1, SOTALLO3, ICTITEM1, ICTCOLL1, ARTCUST2, ARTCUST1, SOTTCLS1" & vbCrLf _
                & " where SOTALLO1.ALLO_CTL_NO = SOTALLOY.ALLO_CTL_NO" & vbCrLf _
                & "   and SOTALLO3.ALLO_CTL_NO (+) = SOTALLOY.ALLO_CTL_NO" & vbCrLf _
                & "   and SOTALLO3.CUST_CODE (+) = SOTALLOY.CUST_CODE" & vbCrLf _
                & "   and SOTALLO3.CUST_STORE_NO (+) = SOTALLOY.CUST_STORE_NO" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = SOTALLO1.ITEM_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTALLOY.CUST_CODE" & vbCrLf _
                & "   and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_CODE = SOTALLOY.CUST_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_STORE_NO = SOTALLOY.CUST_STORE_NO"
            Create_TDA(.Tables.Add, "SOTALLOS", "**", 0, False, "", 3)
            With dst.Tables("SOTALLOS")
                .Columns("ORDR_QTY").DataType = GetType(System.Int64)
                .Columns("ORDR_QTY_OPEN").DataType = GetType(System.Int64)
                .Columns("ORDR_QTY_PICK").DataType = GetType(System.Int64)
                .Columns("ORDR_QTY_SHIP").DataType = GetType(System.Int64)
                .Columns("ORDR_QTY_CANC").DataType = GetType(System.Int64)
            End With


            With dst.Tables("SOTALLOS").Columns
                .Add("BALANCE_CALC", GetType(System.Int64), "ISNULL(QTY_ALLO,0)-ISNULL(ORDR_QTY_SHIP,0)-ISNULL(ORDR_QTY_PICK,0)")
                .Add("BALANCE", GetType(System.Int64), "IIF(BALANCE_CALC<0,0,BALANCE_CALC)")
                .Add("AMT_ALLO", GetType(System.Int64), "QTY_ALLO * ITEM_RETAIL_PRICE")
                .Add("TY_SI_QTY", GetType(System.Int64))
                .Add("TY_SI_AMT", GetType(System.Decimal))
                .Add("LY_SI_QTY", GetType(System.Int64))
                .Add("LY_SI_AMT", GetType(System.Decimal))
                .Add("TY_ST_QTY", GetType(System.Int64))
                .Add("TY_ST_AMT", GetType(System.Decimal))
                .Add("LY_ST_QTY", GetType(System.Int64))
                .Add("LY_ST_AMT", GetType(System.Decimal))
                .Add("TY_SIST", GetType(System.Decimal), "IIF(ISNULL(TY_SI_QTY,0)=0,0,100*ISNULL(TY_ST_QTY,0)/ISNULL(TY_SI_QTY,0))")
                .Add("LY_SIST", GetType(System.Decimal), "IIF(ISNULL(LY_SI_QTY,0)=0,0,100*ISNULL(LY_ST_QTY,0)/ISNULL(LY_SI_QTY,0))")
                .Add("TY_LY_CHG", GetType(System.Decimal), "IIF(ISNULL(LY_ST_QTY,0)=0,0,100*(ISNULL(QTY_ALLO,0)/ISNULL(LY_ST_QTY,0)-1))")
                .Add("BALANCE_OPEN", GetType(System.Int64), "ISNULL(QTY_ALLO,0)-ISNULL(TY_SI_QTY,0)-ISNULL(ORDR_QTY_OPEN,0)-ISNULL(ORDR_QTY_PICK,0)")
            End With

            ASCMAIN1.sql = "Select SOTALLO1.* from " & SOTALLO1 & " SOTALLO1"
            Create_TDA(.Tables.Add, "SOTALLOX", "**", 0, False, "", 1)
            With .Tables("SOTALLOX").Columns
                .Add("QTY_ALLO", GetType(System.Int64))
                .Add("ORDR_QTY_SHIP", GetType(System.Int64))
                .Add("ORDR_QTY_PICK", GetType(System.Int64))
                .Add("BALANCE_CALC", GetType(System.Int64), "ISNULL(QTY_ALLO,0)-ISNULL(ORDR_QTY_SHIP,0)-ISNULL(ORDR_QTY_PICK,0)")
                .Add("BALANCE", GetType(System.Int64), "IIF(BALANCE_CALC<0,0,BALANCE_CALC)")
                .Add("ORDR_QTY_OPEN", GetType(System.Int64))
                .Add("QTY_ALLO_RSRV", GetType(System.Int64), "IIF(ISNULL(QTY_ALLO_PLAN,0)=0,0,QTY_ALLO_PLAN-ISNULL(QTY_ALLO,0))")
                .Add("QTY_ALLO_RPT", GetType(System.Int64))
            End With

            ASCMAIN1.sql = "SELECT ARTCUST1.CUST_CODE_ALLO, SOTORDR2.CUST_CODE, " & vbCrLf _
             & "SUM(SOTORDR2.ORDR_QTY_SHIP) AS QTY, SOTORDR2.ITEM_CODE" & vbCrLf _
             & "FROM SOTORDR2, ARTCUST1 WHERE ROWNUM < 1" & vbCrLf _
             & "GROUP BY ARTCUST1.CUST_CODE_ALLO, SOTORDR2.CUST_CODE, SOTORDR2.ITEM_CODE"
            Create_TDA(.Tables.Add, "SOTALLON", "**", 0, False, "")

            'Create_Relation("SOTALLOX", "SOTALLOZ", "ALLO_CTL_NO")
            '.Tables("SOTALLOX").Columns("QTY_ALLO_RPT").Expression = "SUM(CHILD.QTY_ALLO_RPT)"

        End With

        If perform_fill Then
            Fill_Records_RPT(New String() {sqlw, sqlwCUST_CODE})
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = parms(0) ' MUST HAVE PARMS - THERE IS A DATE
        Dim sqlwCUST_CODE As String = parms(1)

        Create_Temp_Tables(sqlw, sqlwCUST_CODE)

        EnforceConstraints(False)

        'If Not remotely_controlled Then
        '    If SQLA("CUST_CODE") <> "" Then
        '        ASCDATA1.ExecuteSQL("Delete from " & SOTALLOZ & " where CUST_CODE in (Select Distinct CUST_CODE from " & SOTALLOZ & " minus Select Distinct CUST_CODE from " & SOTALLOZ & " " & ASCMAIN1.SQL_Add_WHERE(SQL_in("CUST_CODE")) & ")")
        '    End If
        'End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Filling Dataset")

        Fill_Records("SOTALLO1")
        Fill_Records("SOTALLO2")
        Fill_Records("SOTALLOZ")
        Fill_Records("ICTSTAT2")
        Fill_Records("ARTCUST1")

        Create_SOTALLOC(sqlwCUST_CODE)
        Fill_Records("SOTALLOC")

        Fill_Records("SOTALLOX")
        Fill_Records("ARTCUST2")
        Fill_Records("SOTALLOS")
        ALLO_CTL_NOs.Clear()

        For Each rowSOTALLOX As DataRow In dst.Tables("SOTALLOX").Select("")
            Dim ALLO_CTL_NO As String = rowSOTALLOX.Item("ALLO_CTL_NO")
            Dim QTY_ALLO As Int64 = 0
            Dim ORDR_QTY_SHIP As Int64 = 0
            Dim ORDR_QTY_PICK As Int64 = 0
            Dim ORDR_QTY_OPEN As Int64 = 0
            'Dim QTY_ALLO_RPT As Int64 = 0
            For Each rowSOTALLOC As DataRow In dst.Tables("SOTALLOC").Select("ALLO_CTL_NO = '" & ALLO_CTL_NO & "'", "")
                QTY_ALLO += Val(rowSOTALLOC.Item("QTY_ALLO") & "")
                ORDR_QTY_SHIP += Val(rowSOTALLOC.Item("ORDR_QTY_SHIP") & "")
                ORDR_QTY_PICK += Val(rowSOTALLOC.Item("ORDR_QTY_PICK") & "")
                ORDR_QTY_OPEN += Val(rowSOTALLOC.Item("ORDR_QTY_OPEN") & "")

                Dim CUST_CODE As String = rowSOTALLOC.Item("CUST_CODE")
                Dim rowSOTALLOZ As DataRow = dst.Tables("SOTALLOZ").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE})
                If rowSOTALLOZ IsNot Nothing Then
                    rowSOTALLOZ.Item("QTY_ALLO_RPT") = Val(rowSOTALLOZ.Item("QTY_ALLO_RPT") & "") + Val(rowSOTALLOC.Item("QTY_ALLO") & "")
                End If
            Next
            rowSOTALLOX.Item("QTY_ALLO") = QTY_ALLO
            rowSOTALLOX.Item("ORDR_QTY_SHIP") = ORDR_QTY_SHIP
            rowSOTALLOX.Item("ORDR_QTY_PICK") = ORDR_QTY_PICK
            rowSOTALLOX.Item("ORDR_QTY_OPEN") = ORDR_QTY_OPEN
            'rowSOTALLOX.Item("QTY_ALLO_RPT") = QTY_ALLO_RPT

            Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO)
            Dim BALANCE As Int64 = Val(dst.Tables("SOTALLOC").Compute("SUM(BALANCE)", "ALLO_CTL_NO = '" & ALLO_CTL_NO & "'"))
            rowSOTALLO1.Item("BALANCE") = BALANCE
            'If Val(rowSOTALLOX.Item("BALANCE") & "") > 0 Then
            '    rowSOTALLO1.Item("BALANCE") = rowSOTALLOX.Item("BALANCE")
            'End If
        Next




        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Gathering Sell-Thru Stats")


        ASCMAIN1.sql = "Select DISTINCT X.ALLO_CTL_NO, X.ITEM_CODE, X.DATE_START, X.DATE_END" & vbCrLf _
            & ", NVL(SOTALLO1.ITEM_CODE_COMPARE_TO,ICTITEM1.ITEM_CODE_COMPARE_TO) ITEM_CODE_COMPARE_TO" & vbCrLf _
            & " from " & SOTALLOC & " X,SOTALLO1,ICTITEM1" & vbCrLf _
            & " where SOTALLO1.ALLO_CTL_NO = X.ALLO_CTL_NO" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = X.ITEM_CODE"

        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim ITEM_CODE_COMPARE_TO As String = row.Item("ITEM_CODE_COMPARE_TO") & ""
            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            If ITEM_CODE_COMPARE_TO = "" Then
                If rowICTITEM1.Item("ITEM_CODE_COMPARE_TO") & "" <> "" Then
                    ITEM_CODE_COMPARE_TO = rowICTITEM1.Item("ITEM_CODE_COMPARE_TO")
                Else
                    ITEM_CODE_COMPARE_TO = ITEM_CODE
                End If
            End If
            Dim ALLO_CTL_NO As String = row.Item("ALLO_CTL_NO")
            ALLO_CTL_NOs.Add(ALLO_CTL_NO)
            Dim DATE_START As Date = row.Item("DATE_START")
            Dim DATE_END As Date = row.Item("DATE_END")

            Dim YP1 As String = Format(DATE_START, "yyyyMM")
            Dim YP2 As String = Format(DATE_END, "yyyyMM")

            ITEM_CODE = row.Item("ITEM_CODE")
            Dim ST_QTY As String = "TY_ST_QTY"
            Dim ST_AMT As String = "TY_ST_AMT"
            For y As Integer = 0 To 1
                If y = 1 Then
                    YP1 = Format(Val(Mid(YP1, 1, 4)) - 1, "0000") & Mid(YP1, 5, 2)
                    YP2 = Format(Val(Mid(YP2, 1, 4)) - 1, "0000") & Mid(YP2, 5, 2)
                    ST_QTY = "LY_ST_QTY"
                    ST_AMT = "LY_ST_AMT"
                    ITEM_CODE = ITEM_CODE_COMPARE_TO
                End If

                ASCMAIN1.sql = "Select CUST_CODE, SUM (QTY_SOLD) QTY, SUM (AMT_SOLD) AMT" & vbCrLf _
                    & " from RSTRETL1" & vbCrLf _
                    & " where ITEM_CODE = '" & ITEM_CODE & "'" & vbCrLf _
                    & "   and OPS_YYYYPP between '" & YP1 & "' and '" & YP2 & "'" & vbCrLf _
                    & sqlwCUST_CODE & vbCrLf _
                    & " group by CUST_CODE"

                For Each rowC_I As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim CUST_CODE As String = rowC_I.Item("CUST_CODE")
                    Dim rowSOTALLOC As DataRow = dst.Tables("SOTALLOC").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE})
                    If rowSOTALLOC IsNot Nothing Then
                        rowSOTALLOC.Item(ST_QTY) = rowC_I.Item("QTY")
                        rowSOTALLOC.Item(ST_AMT) = rowC_I.Item("AMT")
                    End If
                Next
            Next

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Gathering Sell-In Stats")

            YP1 = Format(DATE_START, "yyyyMM")
            YP2 = Format(DATE_END, "yyyyMM")

            ITEM_CODE = row.Item("ITEM_CODE")
            Dim SI_QTY As String = "TY_SI_QTY"
            Dim SI_AMT As String = "TY_SI_AMT"
            For y As Integer = 0 To 1
                If y = 1 Then
                    YP1 = Format(Val(Mid(YP1, 1, 4)) - 1, "0000") & Mid(YP1, 5, 2)
                    YP2 = Format(Val(Mid(YP2, 1, 4)) - 1, "0000") & Mid(YP2, 5, 2)
                    SI_QTY = "LY_SI_QTY"
                    SI_AMT = "LY_SI_AMT"
                    ITEM_CODE = ITEM_CODE_COMPARE_TO
                End If

                'the replace was only looking for cust code= cust code before, not cust code allo = cust code
                Dim sqlCUST_CODE As String = ""
                If Not String.IsNullOrEmpty(sqlwCUST_CODE) Then
                    sqlCUST_CODE = " AND (" & Replace(sqlwCUST_CODE, " and CUST_CODE", "SOTINVH2.CUST_CODE") _
                        & " OR " & Replace(sqlwCUST_CODE, " and CUST_CODE", "ARTCUST1.CUST_CODE_ALLO") & ")"
                Else
                    sqlCUST_CODE = ""
                End If

                ASCMAIN1.sql = "Select NVL(ARTCUST1.CUST_CODE_ALLO,SOTINVH2.CUST_CODE) CUST_CODE, SUM (SOTINVH2.ORDR_QTY_SHIP) QTY, SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0)) AMT" & vbCrLf _
                    & " from SOTINVH2,ARTCUST1" & vbCrLf _
                    & " where SOTINVH2.ITEM_CODE = '" & ITEM_CODE & "'" & vbCrLf _
                    & "   and ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
                    & "   and SOTINVH2.ORDR_YYYYPP_UPDATED between '" & YP1 & "' and '" & YP2 & "'" & vbCrLf _
                    & "   and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
                    & sqlCUST_CODE & vbCrLf _
                    & " group by NVL(ARTCUST1.CUST_CODE_ALLO,SOTINVH2.CUST_CODE)"

                If y = 0 Then
                    If Not String.IsNullOrEmpty(sqlwCUST_CODE) Then
                        sqlCUST_CODE = " AND (" & Replace(sqlwCUST_CODE, " and CUST_CODE", "SOTORDR2.CUST_CODE") _
                            & " OR " & Replace(sqlwCUST_CODE, " and CUST_CODE", "ARTCUST1.CUST_CODE_ALLO") & ")"
                    Else
                        sqlCUST_CODE = ""
                    End If

                    ASCMAIN1.sql = "Select NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE) CUST_CODE, SUM (SOTORDR2.ORDR_QTY_SHIP) QTY, SUM (NVL(SOTORDR2.ORDR_QTY_SHIP,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) AMT" & vbCrLf _
                        & " from SOTORDR2,ARTCUST1" & vbCrLf _
                        & " where SOTORDR2.ITEM_CODE = '" & ITEM_CODE & "'" & vbCrLf _
                        & "   and ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE" & vbCrLf _
                        & "   and SOTORDR2. ALLO_CTL_NO = '" & ALLO_CTL_NO & "'" & vbCrLf _
                        & sqlCUST_CODE & vbCrLf _
                        & " group by NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE)"

                End If

                For Each rowC_I As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim CUST_CODE As String = rowC_I.Item("CUST_CODE")
                    Dim rowSOTALLOC As DataRow = dst.Tables("SOTALLOC").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE})
                    If rowSOTALLOC IsNot Nothing Then
                        rowSOTALLOC.Item(SI_QTY) = rowC_I.Item("QTY")
                        rowSOTALLOC.Item(SI_AMT) = rowC_I.Item("AMT")
                    End If
                Next

            Next
        Next
        If ALLO_CTL_NOs.Count > 0 Then
            Dim ALLO_CTL_NOs_Formatted As New List(Of String)

            For i As Integer = 0 To ALLO_CTL_NOs.Count - 1 Step 1000
                Dim endIndex As Integer = Math.Min(i + 999, ALLO_CTL_NOs.Count - 1)
                Dim builder As New Text.StringBuilder()
                builder.Append("SOTORDR2.ALLO_CTL_NO IN (")

                For j As Integer = i To endIndex
                    builder.Append("'").Append(ALLO_CTL_NOs(j)).Append("'")
                    If j < endIndex Then builder.Append(",")
                Next

                builder.Append(")")
                ALLO_CTL_NOs_Formatted.Add(builder.ToString())
            Next

            Dim sqlWhere As String = String.Join(" OR ", ALLO_CTL_NOs_Formatted)

            ASCMAIN1.sql = "SELECT ARTCUST1.CUST_CODE_ALLO, SOTORDR2.CUST_CODE, " &
                   "SUM(SOTORDR2.ORDR_QTY_SHIP) AS QTY, SOTORDR2.ITEM_CODE " & vbCrLf &
                   "FROM SOTORDR2, ARTCUST1 " & vbCrLf &
                   "WHERE ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE " & vbCrLf &
                   "AND (" & sqlWhere & ") " & vbCrLf &
                   "GROUP BY ARTCUST1.CUST_CODE_ALLO, SOTORDR2.CUST_CODE, SOTORDR2.ITEM_CODE"

            Fill_Records("SOTALLON", , , ASCMAIN1.sql)
        End If



        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Item Master Data & Images")

        Fill_Records("ICTITEM1")
        If ASCMAIN1.CLIENT = "INT" Then
            For Each rowICTITEM1 As DataRow In dst.Tables("ICTITEM1").Select("ISNULL(ITEM_EAN_CODE,'?') = '?'")
                rowICTITEM1.Item("ITEM_EAN_CODE") = rowICTITEM1.Item("ITEM_UPC_CODE")
            Next
        End If

        Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
        Dim imgba() As Byte = Nothing
        If ASCMAIN1.CLIENT = "INT" Then
            ' NO IMAGES - MEMORY ERROR
        Else
            For Each row As DataRow In dst.Tables("ICTITEM1").Select("")
                Dim ITEM_CODE As String = row.Item("ITEM_CODE")
                Dim IMAGE_FILENAME As String = FOLDER_NAME & "\" & ITEM_CODE & ".JPG"
                If My.Computer.FileSystem.FileExists(IMAGE_FILENAME) Then
                    row.Item("ITEM_IMAGE") = ASCMAIN1.GetImageData(IMAGE_FILENAME)
                    'ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, False, , , imgba)
                    'row.Item("ITEM_IMAGE") = imgba
                Else
                    IMAGE_FILENAME = FOLDER_NAME & "\" & ITEM_CODE & ".PNG"
                    If My.Computer.FileSystem.FileExists(IMAGE_FILENAME) Then
                        row.Item("ITEM_IMAGE") = ASCMAIN1.GetImageData(IMAGE_FILENAME)
                    End If
                End If
            Next
        End If

        EnforceConstraints(True)

    End Sub

    Sub Create_Temp_Tables(SQLW As String, sqlwCUST_CODE As String)

        If remotely_controlled Then
            ALLO_DATE1 = Now.Date
            ALLO_DATE2 = Now.Date
        Else
            ALLO_DATE1 = Absx1.dteFor("ALLO_DATE1").Value
            ALLO_DATE2 = Absx1.dteFor("ALLO_DATE2").Value
        End If

        ASCMAIN1.sql = "Select SOTALLO2.*, SOTTCLS1.CHANNEL_CODE" & vbCrLf _
            & " from SOTALLO2,SOTALLO1,ICTITEM1,ICTCOLL1, SOTTCLS1, ARTCUST1 " & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = SOTALLO1.ITEM_CODE" & vbCrLf _
            & " AND SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & " AND ARTCUST1.CUST_CODE = SOTALLO2.CUST_CODE" & vbCrLf _
            & "   and SOTALLO2.ALLO_CTL_NO = SOTALLO1.ALLO_CTL_NO" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and SOTALLO1.DATE_START <= '" & Format(ALLO_DATE2, "dd-MMM-yyyy") & "'" & vbCrLf _
            & "   and SOTALLO1.DATE_END >= '" & Format(ALLO_DATE1, "dd-MMM-yyyy") & "'" & vbCrLf _
            & SQLW
        If SOTALLO2 = "" Then
            ASCMAIN1.sql &= " and ROWNUM < 1"
            SOTALLO2 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTALLO2 & " Add Primary Key (ALLO_CTL_NO,CUST_CODE)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTALLO2)
            ASCDATA1.ExecuteSQL("Insert into " & SOTALLO2 & " " & ASCMAIN1.sql)
        End If

        ASCMAIN1.sql = "SELECT SOTALLO3.*, SOTTCLS1.CHANNEL_CODE" & vbCrLf _
            & " FROM SOTALLO3, SOTALLO2, SOTALLO1, ICTITEM1, ICTCOLL1, SOTTCLS1, ARTCUST1" & vbCrLf _
            & " WHERE ICTITEM1.ITEM_CODE = SOTALLO1.ITEM_CODE" & vbCrLf _
            & " AND SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & " AND ARTCUST1.CUST_CODE = SOTALLO3.CUST_CODE" & vbCrLf _
            & " AND SOTALLO3.ALLO_CTL_NO = SOTALLO1.ALLO_CTL_NO" & vbCrLf _
            & " AND SOTALLO2.ALLO_CTL_NO = SOTALLO3.ALLO_CTL_NO" & vbCrLf _
            & " AND SOTALLO2.CUST_CODE = SOTALLO3.CUST_CODE" & vbCrLf _
            & " AND ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " AND SOTALLO1.DATE_START <= '" & Format(ALLO_DATE2, "dd-MMM-yyyy") & "'" & vbCrLf _
            & " AND SOTALLO1.DATE_END >= '" & Format(ALLO_DATE1, "dd-MMM-yyyy") & "'" & vbCrLf _
            & SQLW
        If SOTALLO3 = "" Then
            ASCMAIN1.sql &= " and ROWNUM < 1"
            SOTALLO3 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTALLO3 & " Add Primary Key (ALLO_CTL_NO,CUST_CODE,CUST_STORE_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTALLO3)
            ASCDATA1.ExecuteSQL("Insert into " & SOTALLO3 & " " & ASCMAIN1.sql)
        End If

        ASCMAIN1.sql = "Select SOTALLO1.*" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_SNU_CODE, ICTITEM1.ITEM_BASIC_PROMO" & vbCrLf _
            & ", ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.ITEM_STATUS" & vbCrLf _
            & ", ICTITEM1.PROD_CODE, ICTITEM1.ITEM_NOT_ALLOCATED, ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_SO_QTY_MULT" & vbCrLf _
            & ", ICTITEM1.ITEM_UPC_CODE, NVL(ICTITEM1.ITEM_EAN_CODE,ICTITEM1.ITEM_UPC_CODE) ITEM_EAN_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_DATE_TO_SHIP" & vbCrLf _
            & " from SOTALLO1,ICTITEM1,ICTCOLL1 " & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = SOTALLO1.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and SOTALLO1.ALLO_CTL_NO in (Select Distinct ALLO_CTL_NO from " & SOTALLO2 & ASCMAIN1.SQL_Add_WHERE(sqlwCUST_CODE) & ")"


        If SOTALLO1 = "" Then
            SOTALLO1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTALLO1 & " Add Primary Key (ALLO_CTL_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTALLO1)
            ASCDATA1.ExecuteSQL("Insert into " & SOTALLO1 & " " & ASCMAIN1.sql)
        End If

        ASCMAIN1.sql = "Select SOTORDR2.ALLO_CTL_NO" & vbCrLf _
            & ", NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE) CUST_CODE" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'O',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_OPEN" & vbCrLf _
            & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'P',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_PICK" & vbCrLf _
            & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'F',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_SHIP" & vbCrLf _
            & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'F',SOTORDR1.ORDR_DATE_SHIPPED,NULL)) ORDR_DATE_SHIPPED" & vbCrLf _
            & " from SOTORDR2,SOTORDR1,ARTCUST1" & vbCrLf _
            & "   where (SOTORDR2.ALLO_CTL_NO) in (Select Distinct ALLO_CTL_NO from " & SOTALLO2 & ")" & vbCrLf _
            & "     and ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE" & vbCrLf _
            & "     and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & " group by SOTORDR2.ALLO_CTL_NO, NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE)"

        If SOTALLOZ = "" Then
            SOTALLOZ = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTALLOZ & " Add Primary Key (ALLO_CTL_NO,CUST_CODE)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTALLOZ)
            ASCDATA1.ExecuteSQL("Insert into " & SOTALLOZ & " " & ASCMAIN1.sql)

            ASCDATA1.ExecuteSQL("Delete from " & SOTALLOZ & " where ORDR_QTY_OPEN = 0 and ORDR_QTY_PICK = 0 and ORDR_QTY_SHIP = 0")
            ASCDATA1.ExecuteSQL("Insert into " & SOTALLOZ & " (ALLO_CTL_NO,CUST_CODE) Select ALLO_CTL_NO, CUST_CODE from " & SOTALLO2 & " where (ALLO_CTL_NO,CUST_CODE) in (Select ALLO_CTL_NO,CUST_CODE from " & SOTALLO2 & " minus Select ALLO_CTL_NO,CUST_CODE from " & SOTALLOZ & ")")
        End If



        ASCMAIN1.sql = "Select SOTORDR2.ALLO_CTL_NO" & vbCrLf _
            & ", SOTORDR2.CUST_CODE" & vbCrLf _
            & ", SOTORDR2.CUST_STORE_NO" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'O',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_OPEN" & vbCrLf _
            & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'P',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_PICK" & vbCrLf _
            & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'F',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_SHIP" & vbCrLf _
            & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'F',SOTORDR1.ORDR_DATE_SHIPPED,NULL)) ORDR_DATE_SHIPPED" & vbCrLf _
            & " from SOTORDR2,SOTORDR1,ARTCUST1" & vbCrLf _
            & "   where (SOTORDR2.ALLO_CTL_NO) in (Select Distinct ALLO_CTL_NO from " & SOTALLO2 & ")" & vbCrLf _
            & "     and ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE" & vbCrLf _
            & "     and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "     and ARTCUST1.CUST_ALLOCATE_BY_STORE = '1'" & vbCrLf _
            & " group by SOTORDR2.ALLO_CTL_NO, SOTORDR2.CUST_CODE, SOTORDR2.CUST_STORE_NO"

        If SOTALLOY = "" Then
            SOTALLOY = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTALLOY & " Add Primary Key (ALLO_CTL_NO,CUST_CODE,CUST_STORE_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTALLOY)
            ASCDATA1.ExecuteSQL("Insert into " & SOTALLOY & " " & ASCMAIN1.sql)

            ASCDATA1.ExecuteSQL("Delete from " & SOTALLOY & " where ORDR_QTY_OPEN = 0 and ORDR_QTY_PICK = 0 and ORDR_QTY_SHIP = 0")
            ASCDATA1.ExecuteSQL("Insert into " & SOTALLOY & " (ALLO_CTL_NO,CUST_CODE,CUST_STORE_NO) Select ALLO_CTL_NO, CUST_CODE, CUST_STORE_NO from " & SOTALLO3 & " where (ALLO_CTL_NO,CUST_CODE,CUST_STORE_NO) in (Select ALLO_CTL_NO,CUST_CODE,CUST_STORE_NO from " & SOTALLO3 & " minus Select ALLO_CTL_NO,CUST_CODE,CUST_STORE_NO from " & SOTALLOY & ")")
        End If

    End Sub

    Private Sub grdARTCUST1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdARTCUST1.AfterRowActivate
        Setup_grdSOTALLOC()
    End Sub

    Sub Setup_grdSOTALLOC()
        If grdARTCUST1.ActiveRow Is Nothing Then
        Else
            Dim CUST_CODE As String = grdARTCUST1.ActiveRow.Cells("CUST_CODE").Value
            Dim dvw As DataView = DirectCast(grdSOTALLOC.DataSource, DataTable).DefaultView
            dvw.RowFilter = "CUST_CODE = '" & CUST_CODE & "'"
            grdSOTALLOC.Text = "Allocations for " & CUST_CODE
            Sort_grdColumns(grdSOTALLOC, "BRAND_CODE,COLLECTION_CODE,ITEM_SNU_CODE,PROD_CODE,ITEM_CODE")
        End If
    End Sub

    Sub Setup_grdSOTALLOS()
        If grdARTCUST2.ActiveRow Is Nothing Then
            grdSOTALLOS.Visible = False
        Else
            Dim CUST_CODE As String = grdARTCUST2.ActiveRow.Cells("CUST_CODE").Value
            Dim CUST_STORE_NO As String = grdARTCUST2.ActiveRow.Cells("CUST_STORE_NO").Value
            Dim CUST_STORE_NAME As String = grdARTCUST2.ActiveRow.Cells("CUST_STORE_NAME").Value & ""
            Dim dvw As DataView = DirectCast(grdSOTALLOS.DataSource, DataTable).DefaultView
            dvw.RowFilter = "CUST_CODE = '" & CUST_CODE & "' and CUST_STORE_NO = '" & CUST_STORE_NO & "'"
            grdSOTALLOS.Text = "Allocations for " & CUST_CODE & ", Store " & CUST_STORE_NO & " " & CUST_STORE_NAME
            Sort_grdColumns(grdSOTALLOS, "BRAND_CODE,COLLECTION_CODE,ITEM_SNU_CODE,PROD_CODE,ITEM_CODE")
            grdSOTALLOS.Visible = True
        End If
    End Sub

    Private Sub grdARTCUST2_AfterRowActivate(sender As Object, e As EventArgs) Handles grdARTCUST2.AfterRowActivate
        Setup_grdSOTALLOS()
    End Sub

    Sub Create_SOTALLOC(sqlwCUST_CODE As String)

        ' Dim sqlSOTALLOC_I As String = "SELECT DISTINCT ALLO_CTL_NO, ITEM_CODE, DATE_START, DATE_END FROM " & SOTALLOC

        If SOTALLOC = "" Then
            SOTALLOC = ASCMAIN1.Temp_Table(sqlSOTALLOC & sqlwCUST_CODE)

            'sqlSOTALLOC_I = "SELECT DISTINCT ALLO_CTL_NO, ITEM_CODE, DATE_START, DATE_END FROM " & SOTALLOC
            'SOTALLOC_I = ASCMAIN1.Temp_Table(sqlSOTALLOC_I)

            'ASCDATA1.ExecuteSQL("Alter Table " & SOTALLOC_i & " Add Primary Key (ALLO_CTL_NO)")
            'ASCDATA1.ExecuteSQL("Alter Table " & SOTALLOC_i & " Add TY_ST_QTY NUMBER (10,0)")
            'ASCDATA1.ExecuteSQL("Alter Table " & SOTALLOC_i & " Add TY_ST_AMT NUMBER (13,2)")
            'ASCDATA1.ExecuteSQL("Alter Table " & SOTALLOC_i & " Add LY_ST_AMT NUMBER (10,0)")
            'ASCDATA1.ExecuteSQL("Alter Table " & SOTALLOC_i & " Add LY_ST_AMT NUMBER (13,2)")

        Else
            ASCMAIN1.sql = "Truncate Table " & SOTALLOC
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Insert into " & SOTALLOC & " " & sqlSOTALLOC ' & Replace(sqlwCUST_CODE, "CUST_CODE", "SOTALLO2.CUST_CODE")
            ASCDATA1.ExecuteSQL()

            'sqlSOTALLOC_I = "SELECT DISTINCT ALLO_CTL_NO, ITEM_CODE, DATE_START, DATE_END FROM " & SOTALLOC

            'ASCMAIN1.sql = "Truncate Table " & SOTALLOC_I
            'ASCDATA1.ExecuteSQL()
            'ASCMAIN1.sql = "Insert into " & SOTALLOC_I & " " & sqlSOTALLOC_I
            'ASCDATA1.ExecuteSQL()

        End If
    End Sub

    Sub Create_Pivot_by_Store()

        Dim useSSG As Boolean = False
        ASCMAIN1.Progress("Now Creating Allocations Status Workbook by Store")

        Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & ".xlsx"
        Dim DataTable As DataTable
        Dim r As Integer = 0

        Dim excel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlSourceRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim xlDestRange As Microsoft.Office.Interop.Excel.Range = Nothing

        Dim workbook As SpreadsheetGear.IWorkbook = Nothing
        Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing

        If useSSG Then
            workbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
            worksheet = workbook.Sheets("Data")
        Else
            excel = New Microsoft.Office.Interop.Excel.Application
            ' wb = excel.Workbooks.Open(ASCMAIN1.Folders("Work") & "SORALLO1.xlsx")
            wb = excel.Workbooks.Open(FILENAME)
            ws = wb.Worksheets("Data")
        End If

        DataTable = dst.Tables("SOTALLOS")

        If useSSG Then
            ' NOT USED - BECAUSE THERE IS A PIVOT TABLE IN THIS XLS - AND CODE IS OLD AND NOT REFLECTIVE OF LASTEST CHANGES
            range = worksheet.Cells("A1")
            range.CopyFromDataTable(DataTable, SpreadsheetGear.Data.SetDataFlags.NoColumnHeaders)
            workbook.Names.Add("PivotBase", "=DATA!$R$3:$AE$" & CStr(3 + DataTable.Rows.Count))
        Else
            r = 0
            For Each row As DataRow In DataTable.Select("")
                r += 1
                ws.Range("A" & CStr(1 + r) & ":AK" & CStr(1 + r)).Value2 = row.ItemArray
            Next

            xlSourceRange = ws.Range("AL2:AL2")
            xlDestRange = ws.Range("AL2:AL" & CStr(1 + DataTable.Rows.Count))
            xlSourceRange.Copy(xlDestRange)

            wb.Names.Add("PivotBase", "=Data!$A$1:$AL$" & CStr(1 + DataTable.Rows.Count))
        End If

        If useSSG Then
            ' worksheet.Cells("A4:Q4").Copy(worksheet.Cells("A4:Q" & CStr(3 + DataTable.Rows.Count)))
            '  worksheet.Cells("C1").Value = Now
        Else
            'xlSourceRange = ws.Range("A1:AG1")
            'xlDestRange = ws.Range("A4:Q" & CStr(3 + DataTable.Rows.Count))
            'xlSourceRange.Copy(xlDestRange)
            ' ws.Cells(1, 3).Value = Now
        End If

        'If useSSG Then
        '    worksheet.Visible = False
        'Else
        '    ws.Visible = False
        'End If

        If useSSG Then

        Else
            'excel.Run("ResetData")
        End If


        ASCMAIN1.Progress("Now Saving Workbook")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = "Allocations_Status_by_Store"
                XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsx"

                'If My.Computer.FileSystem.FileExists(ASCMAIN1.Folders("Work") & XLS_FILENAME) Then
                '    Try
                '        My.Computer.FileSystem.DeleteFile(ASCMAIN1.Folders("Work") & XLS_FILENAME)
                '    Catch ex As Exception

                '    End Try
                'End If


                If useSSG Then
                    workbook.SaveAs(ASCMAIN1.Folders("Work") & XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook) ' SpreadsheetGear.FileFormat.OpenXMLWorkbook) ' SpreadsheetGear.FileFormat.OpenXMLWorkbookMacroEnabled)
                Else
                    Dim objOpt As Object = Nothing ' Missing.Value
                    wb.SaveAs(ASCMAIN1.Folders("Work") & XLS_FILENAME _
                              , Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook)
                    wb.Close(False, objOpt, objOpt)
                End If

                success = True

            Catch ex As Exception
                'Stop
            End Try
        Loop

        If useSSG Then
        Else
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
        End If

        Show_Document(ASCMAIN1.Folders("Work") & XLS_FILENAME)

        ASCMAIN1.Progress("")
    End Sub

    Private Sub chkCHKEXCEL_CheckedChanged(sender As Object, e As EventArgs) Handles chkCHKEXCEL.CheckedChanged

    End Sub
    Sub GENERATE_RECAP()
        Dim xls_filename As String = ASCMAIN1.Folders("Work") & "SORALLO1.xlsx"
        Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(xls_filename)
        Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets.Add()
        oSheet.Name = "RECAP"
        Dim range As SpreadsheetGear.IRange = oSheet.Cells("A1:Z999")

        Dim RX As Int32 = 1
        Dim CX As Int32 = 0
        Dim TOTAL_COL As Int32 = 13 '12


        oSheet.Cells.Font.Name = "Verdana"
        oSheet.Cells.Font.Size = 10
        ''    oSheet.Cells.Columns.AutoFit()
        CX = TOTAL_COL
        oSheet.Cells(RX, CX).Value = "           Totals"
        oSheet.Cells(RX, CX).Interior.Color = SpreadsheetGear.Colors.LightSkyBlue
        oSheet.Cells(RX, CX + 1).Interior.Color = SpreadsheetGear.Colors.LightSkyBlue

        Dim RECAP_CUST_CODE As String = ""
        For Each RECAP_CUST_CODE In CUST_CODE_RECAP.Keys
            CX = TOTAL_COL + (CUST_CODE_RECAP(RECAP_CUST_CODE) * 2)
            oSheet.Cells(RX, CX).Value = "           " & RECAP_CUST_CODE & "           "
            oSheet.Cells(RX, CX).Interior.Color = SpreadsheetGear.Colors.Yellow
            oSheet.Cells(RX, CX).NumberFormat = "@"
            oSheet.Hyperlinks.Add(oSheet.Cells(RX, CX),
                                                  "",
                                                  "'" & RECAP_CUST_CODE & "'!A4",
                                                 "Click Here to Navigate to " & RECAP_CUST_CODE,
                                                  "")
            oSheet.Cells(RX, CX + 1).Interior.Color = SpreadsheetGear.Colors.Yellow
            oSheet.Cells(RX, CX + 1).NumberFormat = "@"

            oSheet.Cells(RX + 1, CX).Value = "Qty Allo"
            oSheet.Cells(RX + 1, CX).ColumnWidth = 11
            oSheet.Cells(RX + 1, CX).NumberFormat = "@"
            oSheet.Cells(RX + 1, CX).Interior.Color = SpreadsheetGear.Colors.LightGray


            oSheet.Cells(RX + 1, CX + 1).Value = "Bal Open "
            oSheet.Cells(RX + 1, CX + 1).ColumnWidth = 11
            oSheet.Cells(RX + 1, CX + 1).NumberFormat = "@"
            oSheet.Cells(RX + 1, CX + 1).Interior.Color = SpreadsheetGear.Colors.LightGray


        Next

        oSheet.Cells(1, 0).RowHeight = 19.1
        oSheet.Cells(2, 0).RowHeight = 19.1

        RX = 2
        CX = -1
        Dim NUMBER_COLUMNS As Integer = 14

        CX += 1 : oSheet.Cells(RX, CX).Value = "Brand"
        oSheet.Cells(RX, CX).ColumnWidth = 7.67
        CX += 1 : oSheet.Cells(RX, CX).Value = "Cllctn"
        oSheet.Cells(RX, CX).ColumnWidth = 10.44
        CX += 1 : oSheet.Cells(RX, CX).Value = "Item Code"
        oSheet.Cells(RX, CX).ColumnWidth = 16.11
        CX += 1 : oSheet.Cells(RX, CX).Value = "EAN/UPC"
        oSheet.Cells(RX, CX).ColumnWidth = 20.33
        CX += 1 : oSheet.Cells(RX, CX).Value = "Description"
        oSheet.Cells(RX, CX).ColumnWidth = 37
        CX += 1 : oSheet.Cells(RX, CX).Value = "Retail"
        oSheet.Cells(RX, CX).ColumnWidth = 10.44
        CX += 1 : oSheet.Cells(RX, CX).Value = "Mult"
        oSheet.Cells(RX, CX).ColumnWidth = 6.22
        CX += 1 : oSheet.Cells(RX, CX).Value = "SN"
        oSheet.Cells(RX, CX).ColumnWidth = 7.67
        CX += 1 : oSheet.Cells(RX, CX).Value = "BP"
        oSheet.Cells(RX, CX).ColumnWidth = 7.67
        CX += 1 : oSheet.Cells(RX, CX).Value = "Prod"
        oSheet.Cells(RX, CX).ColumnWidth = 11.89
        CX += 1 : oSheet.Cells(RX, CX).Value = "Start"
        oSheet.Cells(RX, CX).ColumnWidth = 13
        CX += 1 : oSheet.Cells(RX, CX).Value = "End"
        oSheet.Cells(RX, CX).ColumnWidth = 13
        CX += 1 : oSheet.Cells(RX, CX).Value = "Group"
        oSheet.Cells(RX, CX).ColumnWidth = 13
        CX += 1 : oSheet.Cells(RX, CX).Value = "Qty Allo"
        oSheet.Cells(RX, CX).ColumnWidth = 11
        CX += 1 : oSheet.Cells(RX, CX).Value = "Bal Open "
        oSheet.Cells(RX, CX).ColumnWidth = 11

        For CX = 0 To NUMBER_COLUMNS
            If CX = 10 Or CX = 11 Then
                oSheet.Cells(RX, CX).Interior.Color = SpreadsheetGear.Colors.LightSalmon
            Else
                oSheet.Cells(RX, CX).Interior.Color = SpreadsheetGear.Colors.LightGray
            End If

            oSheet.Cells(RX, CX).Font.Color = SpreadsheetGear.Colors.Black
            '   oSheet.Cells(RX, CX).Font.Bold = True
        Next

        oSheet.Cells(RX + 1, 0).Activate()
        oSheet.WindowInfo.FreezePanes = True

        RX = 2
        CX = 0

        Dim TOT_CT As Integer = 0
        Dim TOT_ALLO As Double = 0
        Dim TOT_BAL As Double = 0
        Dim CURR_ITEM_CODE As String = ""
        Dim CURR_CUST_CODE As String = ""
        Dim START_DATE As Date = Now.Date
        Dim END_DATE As Date = Now.Date
        Dim START_END_DATE As String = ""
        Dim ALLO_GROUP_CODE As String = ""

        For Each rowSOTALLOC As DataRow In dst.Tables("SOTALLOC").Select("", "BRAND_CODE, COLLECTION_CODE, ITEM_SNU_CODE, PROD_CODE, ITEM_CODE, DATE_START DESC, DATE_END")
            CURR_CUST_CODE = rowSOTALLOC.Item("CUST_CODE") & ""
            START_END_DATE = Format(rowSOTALLOC.Item("DATE_START"), "yyyyMMdd") & Format(rowSOTALLOC.Item("DATE_END"), "yyyyMMdd")
            ALLO_GROUP_CODE = rowSOTALLOC.Item("ALLO_GROUP_CODE") & ""

            If CUST_CODE_RECAP.ContainsKey(CURR_CUST_CODE) Then
                If CURR_ITEM_CODE = "" Or CURR_ITEM_CODE <> rowSOTALLOC.Item("ITEM_CODE") & "" Or Not START_END_DATEs.ContainsKey(START_END_DATE) Then
                    If CURR_ITEM_CODE <> "" Then
                        ' PRINT TOTALS 
                        oSheet.Cells(RX, TOTAL_COL).Value = TOT_ALLO
                        oSheet.Cells(RX, TOTAL_COL + 1).Value = TOT_BAL
                        CX += 1 : oSheet.Cells(RX, TOTAL_COL - 3).Value = START_DATE
                        oSheet.Cells(RX, TOTAL_COL - 2).NumberFormat = "MM/DD"
                        CX += 1 : oSheet.Cells(RX, TOTAL_COL - 2).Value = END_DATE
                        oSheet.Cells(RX, TOTAL_COL - 1).NumberFormat = "MM/DD"
                        CX += 1 : oSheet.Cells(RX, TOTAL_COL - 1).Value = ALLO_GROUP_CODE
                        oSheet.Cells(RX, TOTAL_COL - 1).NumberFormat = "@"
                    End If
                    CX = -1
                    RX += 1
                    TOT_ALLO = 0
                    TOT_BAL = 0


                    If CURR_ITEM_CODE = rowSOTALLOC.Item("ITEM_CODE") & "" Then ' DO NOT REPRINT CERTAIN FIELDS WHEN ONLY START AND END DATE DIFF
                        CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTALLOC.Item("BRAND_CODE")
                        CX += 1 : oSheet.Cells(RX, CX).Value = "'" & rowSOTALLOC.Item("COLLECTION_CODE")
                        CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTALLOC.Item("ITEM_CODE")
                        oSheet.Cells(RX, CX).Interior.Color = SpreadsheetGear.Colors.Beige
                        CX += 1 : oSheet.Cells(RX, CX).NumberFormat = "@" : oSheet.Cells(RX, CX).Value = rowSOTALLOC.Item("ITEM_EAN_CODE")
                        CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTALLOC.Item("ITEM_DESC")
                        CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTALLOC.Item("ITEM_RETAIL_PRICE")
                        oSheet.Cells(RX, CX).NumberFormat = "$#,##0.00"
                        CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTALLOC.Item("ITEM_SO_QTY_MULT")
                        oSheet.Cells(RX, CX).NumberFormat = "@"
                        CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTALLOC.Item("ITEM_SNU_CODE")
                        CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTALLOC.Item("ITEM_BASIC_PROMO")
                        CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTALLOC.Item("PROD_CODE")
                    Else
                        START_END_DATEs.Clear()
                        START_DATE = rowSOTALLOC.Item("DATE_START")
                        END_DATE = rowSOTALLOC.Item("DATE_END")
                        CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTALLOC.Item("BRAND_CODE")
                        CX += 1 : oSheet.Cells(RX, CX).Value = "'" & rowSOTALLOC.Item("COLLECTION_CODE")
                        CX += 1 : oSheet.Cells(RX, CX).NumberFormat = "@" : oSheet.Cells(RX, CX).Value = rowSOTALLOC.Item("ITEM_CODE")
                        oSheet.Cells(RX, CX).Interior.Color = SpreadsheetGear.Colors.Beige
                        CX += 1 : oSheet.Cells(RX, CX).NumberFormat = "@" : oSheet.Cells(RX, CX).Value = rowSOTALLOC.Item("ITEM_EAN_CODE")
                        CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTALLOC.Item("ITEM_DESC")
                        CX += 1 : oSheet.Cells(RX, CX).NumberFormat = "$#,##0.00" : oSheet.Cells(RX, CX).Value = rowSOTALLOC.Item("ITEM_RETAIL_PRICE")
                        CX += 1 : oSheet.Cells(RX, CX).NumberFormat = "@" : oSheet.Cells(RX, CX).Value = rowSOTALLOC.Item("ITEM_SO_QTY_MULT")
                        CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTALLOC.Item("ITEM_SNU_CODE")
                        CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTALLOC.Item("ITEM_BASIC_PROMO")
                        CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTALLOC.Item("PROD_CODE")
                    End If

                    CURR_ITEM_CODE = rowSOTALLOC.Item("ITEM_CODE") & ""
                    START_END_DATEs.Add(START_END_DATE, START_END_DATE)

                Else
                    ''START_END_DATE = Format(rowSOTALLOC.Item("DATE_START"), "yyyyMMdd") & Format(rowSOTALLOC.Item("DATE_END"), "yyyyMMdd")
                    ''If START_END_DATEs.ContainsKey(START_END_DATE) Then
                    ''    '  Stop
                    ''Else
                    ''    START_END_DATEs.Add(START_END_DATE, START_END_DATE)
                    ''End If
                End If
                Dim CUST_CX As Int16 = 0
                CURR_CUST_CODE = rowSOTALLOC.Item("CUST_CODE") & ""


                START_DATE = rowSOTALLOC.Item("DATE_START")
                END_DATE = rowSOTALLOC.Item("DATE_END")

                If CUST_CODE_RECAP.ContainsKey(CURR_CUST_CODE) Then
                    CX = TOTAL_COL + (CUST_CODE_RECAP(CURR_CUST_CODE) * 2)
                    oSheet.Cells(RX, CX).Value = Val(oSheet.Cells(RX, CX).Value) + Val(rowSOTALLOC.Item("QTY_ALLO") & "")
                    oSheet.Cells(RX, CX + 1).Value = Val(oSheet.Cells(RX, CX + 1).Value) + Val(rowSOTALLOC.Item("BALANCE_OPEN") & "")
                    TOT_ALLO = TOT_ALLO + Val(rowSOTALLOC.Item("QTY_ALLO") & "")
                    TOT_BAL = TOT_BAL + Val(rowSOTALLOC.Item("BALANCE_OPEN") & "")
                Else
                End If

            End If
        Next
        '     If TOT_ALLO <> 0 Then
        oSheet.Cells(RX, TOTAL_COL).Value = TOT_ALLO
        oSheet.Cells(RX, TOTAL_COL + 1).Value = TOT_BAL
        CX += 1 : oSheet.Cells(RX, TOTAL_COL - 2).Value = START_DATE
        oSheet.Cells(RX, TOTAL_COL - 3).NumberFormat = "MM/DD"
        CX += 1 : oSheet.Cells(RX, TOTAL_COL - 1).Value = END_DATE
        oSheet.Cells(RX, TOTAL_COL - 2).NumberFormat = "MM/DD"
        CX += 1 : oSheet.Cells(RX, TOTAL_COL - 1).Value = ALLO_GROUP_CODE
        oSheet.Cells(RX, TOTAL_COL - 1).NumberFormat = "@"
        '      End If
        xls_filename = ASCMAIN1.Folders("Work") & "SORALLO1_Recap.xlsx"

        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        If My.Computer.FileSystem.FileExists(xls_filename) Then
            xls_filename = ASCMAIN1.Folders("Work") & "SORALLO1_Recap"
            Do Until success
                Try
                    XLS_NO += 1
                    xls_filename = ASCMAIN1.Folders("Work") & "SORALLO1_Recap" & "-" & Format(XLS_NO, "00000") & ".xlsx"

                    If Not My.Computer.FileSystem.FileExists(xls_filename) Then
                        success = True
                    End If
                Catch ex As Exception
                    Stop
                End Try
            Loop

        End If

        oWB.SaveAs(xls_filename, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        oWB.Close()
        Show_Document(xls_filename)

    End Sub
    Private Sub grdSOTALLOC_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTALLOC.InitializeRow
        Dim rowSOTALLOC As DataRowView = CType(e.Row.ListObject, DataRowView)
        Dim CUST_CODE_ALLO As String = rowSOTALLOC("CUST_CODE").ToString()
        Dim ITEM_CODE As String = rowSOTALLOC("ITEM_CODE").ToString()
        Dim subordinateRows = dst.Tables("SOTALLON").Select($"CUST_CODE_ALLO = '{CUST_CODE_ALLO}' AND ITEM_CODE = '{ITEM_CODE}'")

        For i As Integer = 0 To 9
            Dim colName As String = $"CUST_{i + 1:00}"
            Dim custCol = grdSOTALLOC.DisplayLayout.Bands(0).Columns(colName)

            e.Row.Cells(colName).Value = 0
            custCol.Hidden = custCol.Header.Caption.StartsWith("CUST_")
        Next

        For Each subordinateRow As DataRow In subordinateRows
            Dim CUST_CODE As String = subordinateRow("CUST_CODE").ToString()
            Dim QTY As Integer = If(IsDBNull(subordinateRow("QTY")), 0, Convert.ToInt32(subordinateRow("QTY")))

            For i As Integer = 0 To 9
                Dim colName As String = $"CUST_{i + 1:00}"
                Dim custCol = grdSOTALLOC.DisplayLayout.Bands(0).Columns(colName)

                If custCol.Header.Caption = CUST_CODE Then
                    e.Row.Cells(colName).Value = QTY
                    custCol.Hidden = (QTY = 0)
                    Exit For
                ElseIf custCol.Header.Caption.StartsWith("CUST_") Then
                    custCol.Header.Caption = CUST_CODE
                    e.Row.Cells(colName).Value = QTY
                    custCol.Hidden = (QTY = 0)
                    Exit For
                End If
            Next
        Next
        rowSOTALLOC.EndEdit()
        grdSOTALLOC.UpdateData()
    End Sub

End Class