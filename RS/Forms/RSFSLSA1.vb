Imports Infragistics.Win
Imports Infragistics.UltraChart.Resources

Public Class RSFSLSA1
    Dim RYP As String
    Dim generated As Boolean = False
    Dim RSTSLSAX As String
    Dim RSTSLSA0 As String
    Dim YPP(,) As String
    Dim YPPD() As Date
    Dim sqlRSTSLSAP As String
    Dim RSTSLSAQ As String
    Dim CN As New Dictionary(Of String, String)
    Dim RSTSLSB0 As String
    Dim sqlRSTSLSB0 As String
    'Dim RSTSLSB1 As String
    Dim sqlRSTSLSB1 As String
    Dim sqlRSTSLSB2 As String
    Dim sqlRSTSLSA3 As String
    Dim sqlRSTSLSBR As String
    Dim excel_preparation As Boolean = False
    Dim grdRSTSLSB1_HCs As String()
    Dim grdRSTSLSBR_HCs As String()

    Dim SUMCOLs As New Dictionary(Of String, String)
    Dim filter_count As Int32 = 0
    Dim filter_desc As String = ""
    Dim sql_filter As String = ""
    Dim sql_filters As New Dictionary(Of String, String)
    Dim sql_filter_no_CUST_CODE As String = ""
    Dim filter_CUST_CODE As String = ""
    Dim ICTITEM1 As String = "ICTITEM1"
    Dim WKS60() As String
    Dim OPS_YYYYWW_BASE As String = ASCMAIN1.CYW
    Dim RYW_LGI As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP", ASCMAIN1.CYP, -36, 0, -1)

        grdRSTSLSA1.Tag = "COLLECTION_CODE"
        grdRSTSLSA2.Tag = "ITEM_CLASS_CODE"
        'Dim tlb_sbt As UltraWinToolbars.StateButtonTool
        'tlb_sbt = DirectCast(tlb.Tools("A1COLLECTION_CODE"), UltraWinToolbars.StateButtonTool)
        'tlb_sbt.Checked = True
        'tlb_sbt = DirectCast(tlb.Tools("A2ITEM_CLASS_CODE"), UltraWinToolbars.StateButtonTool)
        'tlb_sbt.Checked = True

        With dst

            Create_TDA(.Tables.Add, "ASTFLTR1", "*", 0, False)
            .Tables("ASTFLTR1").Columns("CODE_VALUES").MaxLength = -1

            Setup_RSTSLSA0(ASCMAIN1.CYP, True)
            Setup_RSTSLSAX(False)

            ASCMAIN1.sql = Get_sqlRSTSLSAX("1", False)
            For Each TABLE_NAME As String In New String() {"RSTSLSA1", "RSTSLSA2", "RSTSLSA3"}
                'ASCMAIN1.sql = Get_sqlRSTSLSAX(Mid(TABLE_NAME, 8, 1), False)

                If TABLE_NAME = "RSTSLSA3" Then
                    ASCMAIN1.sql = Replace(Replace(ASCMAIN1.sql, "NVL(RSTSLSAX.COLLECTION_CODE,'?') CODE_VALUE,", "NVL(RSTSLSAX.COLLECTION_CODE,'?') LINK_VALUE, NVL(RSTSLSAX.COLLECTION_CODE,'?') CODE_VALUE,"), " group by NVL(RSTSLSAX.COLLECTION_CODE,'?')", " group by NVL(RSTSLSAX.COLLECTION_CODE,'?'),NVL(RSTSLSAX.COLLECTION_CODE,'?')")
                End If

                Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, False, "", IIf(TABLE_NAME = "RSTSLSA3", 2, 1))

                With .Tables(TABLE_NAME).Columns
                    .Add("TY_SELL_THRU", GetType(System.Decimal), "IIF((ISNULL(TY_QTY,0)+ISNULL(TY_QOH,0))=0,0,100 * ISNULL(TY_QTY,0)/(ISNULL(TY_QTY,0)+ISNULL(TY_QOH,0)))")
                    .Add("TY_PCT_TOTAL", GetType(System.Decimal), "IIF(0=0,0,TY_QTY/0)")
                    .Add("LY_SELL_THRU", GetType(System.Decimal), "IIF((ISNULL(LY_QTY,0)+ISNULL(LY_QOH,0))=0,0,100 * ISNULL(LY_QTY,0)/(ISNULL(LY_QTY,0)+ISNULL(LY_QOH,0)))")
                    .Add("LY_PCT_TOTAL", GetType(System.Decimal), "IIF(0=0,0,LY_QTY/0)")
                    .Add("TY_QTY_RANK", GetType(System.Int32))
                    .Add("TY_AMT_RANK", GetType(System.Int32))
                    .Add("LY_QTY_RANK", GetType(System.Int32))
                    .Add("LY_AMT_RANK", GetType(System.Int32))
                    .Add("VAR", GetType(System.Decimal), "TY_AMT - LY_AMT")
                    .Add("VAR_PCT", GetType(System.Decimal), "IIF(LY_AMT=0,0,100*VAR/LY_AMT)")
                End With
                With .Tables(TABLE_NAME)
                    .Columns("CODE_VALUE").MaxLength = -1
                    .Columns("DESC_VALUE").MaxLength = -1
                    .Columns("CODE2_VALUE").MaxLength = -1
                End With
            Next

            .Relations.Add("RSTSLSA2_RSTSLSA3", _
                    .Tables("RSTSLSA2").Columns("CODE_VALUE"), _
                    .Tables("RSTSLSA3").Columns("LINK_VALUE"))

            ASCMAIN1.sql = sqlRSTSLSAP
            Create_TDA(.Tables.Add, "RSTSLSAP", "**", 0, False, "", 1)
            'With these limits set to 20/60/20 and 40/60/40 kept on getting the freddie error
            With .Tables("RSTSLSAP")
                .Columns("CODE_VALUE").MaxLength = -1
                .Columns("DESC_VALUE").MaxLength = -1
                .Columns("CODE2_VALUE").MaxLength = -1
            End With

            ASCMAIN1.sql = Replace(sqlRSTSLSB1, ":PARM1", "''")
            Create_TDA(.Tables.Add, "RSTSLSB1", "**", 0, False, "", 1)

            ASCMAIN1.sql = sqlRSTSLSB2 ' Replace(sqlRSTSLSB2, ":PARM1", "''")
            Create_TDA(.Tables.Add, "RSTSLSB2", "**", 0, False, "V", 2)

            For Each TABLE_NAME As String In New String() {"RSTSLSB1", "RSTSLSB2"}
                With .Tables(TABLE_NAME).Columns
                    .Add("TY_SELL_THRU", GetType(System.Decimal), "IIF((ISNULL(TY_QTY,0)+ISNULL(TY_QOH,0))=0,0,100 * ISNULL(TY_QTY,0)/(ISNULL(TY_QTY,0)+ISNULL(TY_QOH,0)))")
                    .Add("LY_SELL_THRU", GetType(System.Decimal), "IIF((ISNULL(LY_QTY,0)+ISNULL(LY_QOH,0))=0,0,100 * ISNULL(LY_QTY,0)/(ISNULL(LY_QTY,0)+ISNULL(LY_QOH,0)))")
                    .Add("TY_QTY_RANK", GetType(System.Int32))
                    .Add("TY_AMT_RANK", GetType(System.Int32))
                    .Add("LY_QTY_RANK", GetType(System.Int32))
                    .Add("LY_AMT_RANK", GetType(System.Int32))


                    If TABLE_NAME = "RSTSLSB1" Then
                        ' THESE COLUMNS SHOULD COME IN WITH VALUES, AND NOT NEED TO BE CALCULATED
                        ' AND THEN THEY WILL WORK FOR RSTSLSB2 ALSO
                        '.Add("TY_AOH", GetType(System.Decimal), "ISNULL(TY_QOH,0) * ISNULL(ITEM_RETAIL_PRICE,0)")
                        '.Add("LY_AOH", GetType(System.Decimal), "ISNULL(LY_QOH,0) * ISNULL(ITEM_RETAIL_PRICE,0)")

                        .Add("N", GetType(System.Int32))
                        .Add("PICTURE", GetType(System.Byte()))
                    Else
                        '.Add("TY_AOH", GetType(System.Decimal))
                        '.Add("LY_AOH", GetType(System.Decimal))
                    End If
                End With
            Next

            ASCMAIN1.sql = sqlRSTSLSBR
            Create_TDA(.Tables.Add, "RSTSLSBR", "**", 0, False, "", 1)
        End With

        grdRSTSLSA1.DataSource = dst.Tables("RSTSLSA1")
        grdRSTSLSA2.DataSource = dst.Tables("RSTSLSA2")
        grdRSTSLSAP.DataSource = dst.Tables("RSTSLSAP")
        grdRSTSLSB1.DataSource = dst.Tables("RSTSLSB1")
        grdRSTSLSB2.DataSource = dst.Tables("RSTSLSB2")
        grdRSTSLSBR.DataSource = dst.Tables("RSTSLSBR")

        Create_Summary(grdRSTSLSB1, "ITEM_CODE", "Count")
        Create_Summary(grdRSTSLSB2, "CUST_CODE", "Count")

        ASCMAIN1.Add_Value_List(grdRSTSLSB1, "ITEM_CATGY_CODE")

        Format_grids()

        TAC.RSCMAIN1.Load_Filter(New String() _
        {"COLLECTION_CODE", "PROD_CODE", "ITEM_CLASS_CODE", "CUST_CODE",
         "MATL_CODE", "ITEM_CATGY_CODE", "DEPT_CODE", "PRICE_POINT_CODE", "ITEM_ABC_CODE", "ITEM_CODE"}, grdASTFLTR1, Me)

        tabBottom.Tabs("Pivot").Visible = False

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                'Validate_Code("OPS_YYYYPP")
            Case "Excel"

                'Dim DVW As DataView = DirectCast(grdRSTSLSB1.DataSource, DataTable).DefaultView
                If tabMain.SelectedTab.Key = "Summary" Then
                    EMsg &= vbCr & "Use Individual Grid Export-to-Excel icons"
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "L"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Excel"
                excel_preparation = True
                Me.Cursor = Cursors.WaitCursor

                Dim myWorkbook As New Infragistics.Documents.Excel.Workbook

                If tabMain.SelectedTab.Key = "Summary" Then
                    myWorkSheet = myWorkbook.Worksheets.Add("Sheet1")
                    Excel_Export_Summary(myWorkbook)
                    ' gets an error
                ElseIf tabMain.SelectedTab.Key = "Top N" Then

                    Excel_Export_TopN(myWorkbook)
                End If

                Export_to_Excel_Show(myWorkbook)

                Call ASCMAIN1.Progress("")
                Me.Cursor = Cursors.Default

                excel_preparation = False

            Case "Done"
                Call Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Excel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Options").Visible = ScreenMode
                .Groups("Filter").Visible = ScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)
        tabMain.Visible = ScreenMode
        Setup_tabMain()


        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("RSTSLSA1").Rows.Clear()
        dst.Tables("RSTSLSA2").Rows.Clear()
        dst.Tables("RSTSLSA3").Rows.Clear()
        dst.Tables("RSTSLSB1").Rows.Clear()
        dst.Tables("RSTSLSB2").Rows.Clear()
        dst.Tables("RSTSLSAP").Rows.Clear()
        dst.EnforceConstraints = True

        filter_CUST_CODE = ""

        lblLGIData.Visible = False

        For Each ROW In dst.Tables("ASTFLTR1").Rows
            ROW.ITEM("CODE_VALUES") = ""
            ROW.ITEM("EXCLUDE") = ""
        Next
        'Absx1.CMBFor("RYP").Value = ""
    End Sub

    Sub Load_Record()

        Call ASCMAIN1.Progress("Now Loading Data ...")

        Call Save_Header_Fields(UltraGroupBox1)

        RYP = Absx1.cmbFor("RYP").Value
        'ICTITEM1 = "ICTITEM1"
        'If chkHISTCAT.Checked Then
        '    ICTITEM1 = RSCMAIN1.Get_ICTITEM1_Hist_CATGY(RYP)
        'End If
        ICTITEM1 = RSCMAIN1.Get_ICTITEM1_Hist_CATGY(RYP, chkHISTCAT.Checked)

        ASCMAIN1.sql = "Select Max (OPS_YYYYWW) from EDT852T1 where EDI_CUST_BATCH_NO = 'LGI'"
        RYW_LGI = ASCDATA1.GetDataValue
        ASCMAIN1.sql = "Select Max (YYYYWW) from GLTPARM3 where YYYYPP = '" & RYP & "'"
        Dim RYP_last_week As String = ASCDATA1.GetDataValue
        Dim N As Integer = ASCMAIN1.Week_Diff(RYP_last_week, RYW_LGI)
        If N < 0 And N > -5 Then
            ' OK TO USE LGI WEEK
            Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", RYW_LGI)
            lblLGIData.Text = "LGI Data as of " & rowGLTPARM3.Item("LEGEND")
            lblLGIData.Visible = False '  True
        Else
            RYW_LGI = ""
        End If

        EnforceConstraints(False)
        Setup_RSTSLSA0(RYP, False)
        Setup_RSTSLSAX(True)
        EnforceConstraints(True)

        Calculate_N()

        ASCMAIN1.Progress("Now Formatting Grids")

        With grdRSTSLSAP.DisplayLayout.Bands(0)
            For P As Integer = 0 To 24
                With .Columns("P" & Format(P, "00"))
                    .Header.Caption = YPP(P, 1)
                End With
            Next
        End With

        Dim VLTrend As ValueListItem = optTrend.ValueList.ValueListItems(1)
        Dim VLTopN As ValueListItem = optTopN.ValueList.ValueListItems(1)
        If optSales.Value = "W" Then
            VLTrend.DisplayText = "$ Sales"
            VLTopN.DisplayText = "$ Sales"
        Else
            VLTrend.DisplayText = "$ Retail"
            VLTopN.DisplayText = "$ Retail"
        End If

        Dim b As Int32 = 0
        For Each grd As UltraWinGrid.UltraGrid In _
        New UltraWinGrid.UltraGrid() {grdRSTSLSA1, grdRSTSLSA2, grdRSTSLSA2}
            With grd.DisplayLayout.Bands(b)
                For Each Y As String In New String() {"TY", "LY"}
                    .Columns(Y & "_SELL_THRU").Hidden = (optSales.Value = "W")
                    .Columns(Y & "_QOH").Hidden = (optSales.Value = "W")
                    .Columns(Y & "_AOH").Hidden = (optSales.Value = "W")

                    .Columns(Y & "_AMT").Header.Caption = IIf(optSales.Value = "W", "Sales", "Retail")
                    .Columns(Y & "_AMT").Header.ToolTipText = IIf(optSales.Value = "W", "W/S", "Retail") & " $ Sold " & Y
                    .Columns(Y & "_AMT_RANK").Header.ToolTipText = "Rank based on $" & IIf(optSales.Value = "W", "Sales", "Retail") & ", " & Y
                Next

            End With

            If grd.Name = "grdRSTSLSA2" Then
                b += 1
            End If
        Next
        For Each grd As UltraWinGrid.UltraGrid In _
        New UltraWinGrid.UltraGrid() {grdRSTSLSB1, grdRSTSLSB2}
            With grd.DisplayLayout.Bands(0)
                For Each Y As String In New String() {"TY", "LY"}
                    .Columns(Y & "_SELL_THRU").Hidden = (optSales.Value = "W")
                    .Columns(Y & "_QOH").Hidden = (optSales.Value = "W")
                    .Columns(Y & "_AOH").Hidden = (optSales.Value = "W")
                Next
            End With
        Next

        grdRSTSLSB1.Text = "All Records"

        ASCMAIN1.Progress("Now Setting up Charts")
        Show_Selected()
        Set_Charts()
        Setup_grdRSTSLSA2()

        tabMain.Tabs("Sales && On Hand by Week").Visible = chkShow60.Checked
        If chkShow60.Checked Then
            ASCMAIN1.Progress("Now Loading 60 Weeks of Sales & On Hand")
            Fill_Records("RSTSLSBR")
            grdRSTSLSBR.Text = "Sales & On Hand by Item, by Week"
        End If


        Call ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Call BeginTrans()
        Call CommitTrans("Update Complete")

    End Sub

    Sub Delete_Record()
        'Call BeginTrans()
        'Call Delete_Records("DPTITMF1")
        'Call CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where MARKET_CODE = '" & HFs("MARKET_CODE") & "'" _
        '    & "   and OPS_YYYYPP = '" & ASCMAIN1.CYP & "'")
    End Sub
#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()

        SUMCOLs.Add("COLLECTION_CODE", "Collection")
        SUMCOLs.Add("HC_CODE", "High Collection")
        SUMCOLs.Add("PROD_CODE", "Product")
        SUMCOLs.Add("ITEM_CLASS_CODE", "Item Class")
        SUMCOLs.Add("CUST_CODE", "Customer")
        SUMCOLs.Add("PRICE_POINT_CODE", "Price Point")
        SUMCOLs.Add("MATL_CODE", "Material")
        SUMCOLs.Add("ITEM_CATGY_CODE", "Category")
        SUMCOLs.Add("DEPT_CODE", "Department")
        SUMCOLs.Add("ITEM_ABC_CODE", "ABC")
        SUMCOLs.Add("MARKET_CODE", "Market")
        SUMCOLs.Add("ITEM_CODE", "Item Code")

        ' it would be great of the list supplied to load_popup_menu 
        ' could be derived from SUMCOLs

        Load_Popup_Menu(grdRSTSLSA1, "SSOOOOOOOOOOOO" _
                    , "Show Filter", "Show GroupBox" _
                    , "A1:A1ITEM_CATGY_CODE|By Category" _
                    , "A1:A1COLLECTION_CODE|By Collection" _
                    , "A1:A1HC_CODE|By High Collection" _
                    , "A1:A1PROD_CODE|By Product" _
                    , "A1:A1CUST_CODE|By Customer" _
                    , "A1:A1DEPT_CODE|By Department" _
                    , "A1:A1ITEM_CLASS_CODE|By Item Class" _
                    , "A1:A1MATL_CODE|By Material" _
                    , "A1:A1PRICE_POINT_CODE|By Price Poin  t" _
                    , "A1:A1ITEM_ABC_CODE|By ABC" _
                    , "A1:A1MARKET_CODE|By Market" _
                    , "A1:A1ITEM_CODE|By Item Code")

        'Call Load_Popup_Menu(grdRSTSLSA1, "SS", "Show Filter", "Show GroupBox")

        Load_Popup_Menu(grdRSTSLSA2, "OOOOOOOOOOOO" _
                     , "A2:A2ITEM_CATGY_CODE|By Category" _
                     , "A2:A2COLLECTION_CODE|By Collection" _
                     , "A2:A2HC_CODE|By High Collection" _
                     , "A2:A2PROD_CODE|By Product" _
                     , "A2:A2CUST_CODE|By Customer" _
                     , "A2:A2DEPT_CODE|By Department" _
                     , "A2:A2ITEM_CLASS_CODE|By Item Class" _
                     , "A2:A2MATL_CODE|By Material" _
                     , "A2:A2PRICE_POINT_CODE|By Price Point" _
                     , "A2:A2ITEM_ABC_CODE|By ABC" _
                     , "A2:A2MARKET_CODE|By Market" _
                     , "A2:A2ITEM_CODE|By Item Code")


        grdRSTSLSB1_HCs = New String() {"ITEM_CODE", "ITEM_DESC" _
                            , "N" _
                             , "COLLECTION_CODE" _
                             , "HC_CODE" _
                             , "PROD_CODE" _
                             , "ITEM_CATGY_CODE", "DEPT_CODE" _
                             , "STYLE_CODE", "STYLE_CODE_NC" _
                             , "COLOR_CODE", "SIZE_CODE" _
                             , "ITEM_CLASS_CODE", "MATL_CODE" _
                             , "PRICE_POINT_CODE" _
                             , "ITEM_RETAIL_PRICE", "ITEM_PRICE", "ITEM_ABC_CODE"}

        grdRSTSLSBR_HCs = New String() {"ITEM_CODE", "ITEM_DESC" _
                     , "COLLECTION_CODE" _
                     , "HC_CODE" _
                     , "PROD_CODE" _
                     , "ITEM_CATGY_CODE", "DEPT_CODE" _
                     , "STYLE_CODE", "STYLE_CODE_NC" _
                     , "COLOR_CODE", "SIZE_CODE" _
                     , "ITEM_CLASS_CODE", "MATL_CODE" _
                     , "PRICE_POINT_CODE" _
                     , "ITEM_RETAIL_PRICE", "ITEM_PRICE", "ITEM_ABC_CODE"}


        Load_Popup_Menu(grdRSTSLSB1, "BSS" & "SSSSSSSSSSSSSSSSSSSSS" _
                             , "Show Image", "Show Filter", "Show GroupBox" _
                             , "ITEM_CODE|Show Item Code" _
                             , "N|Show N" _
                             , "ITEM_DESC|Show Description" _
                             , "COLLECTION_CODE|Show Collection" _
                             , "HC_CODE|Show High Collection" _
                             , "PROD_CODE|Show Product" _
                             , "ITEM_CATGY_CODE|Show Catgy", "DEPT_CODE|Show Dept" _
                             , "STYLE_CODE|Show Style", "STYLE_CODE_NC|Show Style NC" _
                             , "COLOR_CODE|Show Color", "SIZE_CODE|Show Size" _
                             , "ITEM_CLASS_CODE|Show Class", "MATL_CODE|Show Matl" _
                             , "PRICE_POINT_CODE|Price Point" _
                             , "ITEM_RETAIL_PRICE|Show Retail" _
                             , "ITEM_PRICE|Show W/S Price" _
                             , "ITEM_ABC_CODE|Show ABC")
        '                        , "MARKET_CODE|Show Market")


        Load_Popup_Menu(grdRSTSLSBR, "BSS" & "SSSSSSSSSSSSSSSSSSSSS" _
                     , "Show Image", "Show Filter", "Show GroupBox" _
                     , "ITEM_CODE|Show Item Code" _
                     , "N|Show N" _
                     , "ITEM_DESC|Show Description" _
                     , "COLLECTION_CODE|Show Collection" _
                     , "HC_CODE|Show High Collection" _
                     , "PROD_CODE|Show Product" _
                     , "ITEM_CATGY_CODE|Show Catgy", "DEPT_CODE|Show Dept" _
                     , "STYLE_CODE|Show Style", "STYLE_CODE_NC|Show Style NC" _
                     , "COLOR_CODE|Show Color", "SIZE_CODE|Show Size" _
                     , "ITEM_CLASS_CODE|Show Class", "MATL_CODE|Show Matl" _
                     , "PRICE_POINT_CODE|Price Point" _
                     , "ITEM_RETAIL_PRICE|Show Retail" _
                     , "ITEM_PRICE|Show W/S Price" _
                     , "ITEM_ABC_CODE|Show ABC")

        '             , "MARKET_CODE|Show Market")

        Load_Popup_Menu(grdASTFLTR1, "S", "Single Value") ', "Load List", "Maintain Lists")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        'If tlb_pop.Tools.Exists("Load List") Then
        '    tlb_btn = DirectCast(tlb_pop.Tools("Load List"), UltraWinToolbars.ButtonTool)
        '    tlb_btn.SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow)
        'End If
        'If tlb_pop.Tools.Exists("Maintain Lists") Then
        '    tlb_btn = DirectCast(tlb_pop.Tools("Maintain Lists"), UltraWinToolbars.ButtonTool)
        '    tlb_btn.SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow)
        'End If


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else

            Select Case e.SourceControl.Name

                Case "grdASTFLTR1"
                    'If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                    '    e.Cancel = True
                    'End If
                    tlb_sbt = DirectCast(tlb_pop.Tools("Single Value"), UltraWinToolbars.StateButtonTool)
                    tlb_sbt.Tag = "X"
                    If grd.ActiveRow.Cells("CODE_VALUES").Style = UltraWinGrid.ColumnStyle.DropDownList Then
                        tlb_sbt.Checked = True
                    Else
                        tlb_sbt.Checked = False
                    End If
                    tlb_sbt.Tag = ""

                Case "grdRSTSLSA1"
                    Dim COL As String = grdRSTSLSA1.Tag
                    tlb_sbt = DirectCast(tlb_pop.Tools("A1" & COL), UltraWinToolbars.StateButtonTool)
                    If Not tlb_sbt.Checked Then
                        tlb_sbt.Tag = "X"
                        tlb_sbt.Checked = True
                        tlb_sbt.Tag = ""
                    End If

                Case "grdRSTSLSA2"
                    Dim COL As String = grdRSTSLSA2.Tag
                    tlb_sbt = DirectCast(tlb_pop.Tools("A2" & COL), UltraWinToolbars.StateButtonTool)
                    If Not tlb_sbt.Checked Then
                        tlb_sbt.Tag = "X"
                        tlb_sbt.Checked = True
                        tlb_sbt.Tag = ""
                    End If

                Case "grdRSTSLSB1"
                    For Each c As String In grdRSTSLSB1_HCs
                        Dim COL As String = Split(c, "|")(0)
                        tlb_sbt = DirectCast(tlb_pop.Tools(COL), UltraWinToolbars.StateButtonTool)
                        tlb_sbt.Tag = "X"
                        If grdRSTSLSB1.DisplayLayout.Bands(0).Columns(COL).Hidden Then
                            tlb_sbt.Checked = False
                        Else
                            tlb_sbt.Checked = True
                        End If
                        tlb_sbt.SharedProps.Caption = "Show " & grdRSTSLSB1.DisplayLayout.Bands(0).Columns(COL).Header.Caption
                        tlb_sbt.Tag = ""
                    Next

                Case "grdRSTSLSBR"
                    For Each c As String In grdRSTSLSBR_HCs
                        Dim COL As String = Split(c, "|")(0)
                        tlb_sbt = DirectCast(tlb_pop.Tools(COL), UltraWinToolbars.StateButtonTool)
                        tlb_sbt.Tag = "X"
                        If grdRSTSLSBR.DisplayLayout.Bands(0).Columns(COL).Hidden Then
                            tlb_sbt.Checked = False
                        Else
                            tlb_sbt.Checked = True
                        End If
                        tlb_sbt.SharedProps.Caption = "Show " & grdRSTSLSBR.DisplayLayout.Bands(0).Columns(COL).Header.Caption
                        tlb_sbt.Tag = ""
                    Next

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        Select Case grd.Name
            Case "grdRSTSLSA1"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool
                tlb_sbt = DirectCast(tlb.Tools(e.Tool.Key), UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag <> "X" Then
                    If grdRSTSLSA1.ActiveRow IsNot Nothing AndAlso grdRSTSLSA1.ActiveRow.IsDataRow Then
                        grdRSTSLSA1.Tag = Mid(e.Tool.Key, 3)
                        Get_sqlRSTSLSAX("1")
                        CreateGraph_Trend()
                        CreateGraph_Totals()
                    End If
                End If

            Case "grdRSTSLSA2"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool
                tlb_sbt = DirectCast(tlb.Tools(e.Tool.Key), UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag <> "X" Then
                    grdRSTSLSA2.Tag = Mid(e.Tool.Key, 3)
                    If grdRSTSLSA1.ActiveRow Is Nothing Then
                        Exit Sub
                    End If
                    Get_sqlRSTSLSAX("2")
                End If

            Case "grdRSTSLSB1"
                If Join(grdRSTSLSB1_HCs, ",").Contains(e.Tool.Key) Then
                    Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                    If tlb_sbt.Tag <> "" Then Exit Sub
                    Dim COL As String = tlb_sbt.Key
                    grdRSTSLSB1.DisplayLayout.Bands(0).Columns(COL).Hidden = Not tlb_sbt.Checked
                    Exit Sub
                End If

            Case "grdRSTSLSBR"
                If Join(grdRSTSLSBR_HCs, ",").Contains(e.Tool.Key) Then
                    Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                    If tlb_sbt.Tag <> "" Then Exit Sub
                    Dim COL As String = tlb_sbt.Key
                    grdRSTSLSBR.DisplayLayout.Bands(0).Columns(COL).Hidden = Not tlb_sbt.Checked
                    Exit Sub
                End If

        End Select



        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If


        Select Case e.Tool.Key

            'Case "Job Order Inquiry"
            '    Dim JOB_NO As String = grd.ActiveRow.Cells("JOB_NO").Text
            '    Context_Launch("Load", JOB_NO, e.Tool.Key, "DEFJOBMI")

            Case "Show Image"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Dim FILENAME As String = "C:\Documents and Settings\wjz\Desktop\JHI\Images\" & ITEM_CODE & ".JPG"
                If My.Computer.FileSystem.FileExists(FILENAME) Then
                    Me.Cursor = Cursors.WaitCursor
                    Call ASCMAIN1.Progress("Now Loading Image Viewer")
                    System.Diagnostics.Process.Start(FILENAME)
                    Me.Cursor = Cursors.Default
                    Call ASCMAIN1.Progress("")
                End If

            Case "Single Value"

                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag <> "" Then Exit Sub
                Dim COLUMN_NAME As String = grd.ActiveRow.Cells("COLUMN_NAME").Text
                If tlb_sbt.Checked Then
                    'grd.DisplayLayout.Bands(0).Columns("COLUMN_CAPTION").CellAppearance.BackColor = Color.LightGreen
                    ASCMAIN1.sql = ASCMAIN1.CodeSelector.Get_SQL(grd.ActiveRow.Cells("COLUMN_NAME").Text) & " ORDER BY 1"
                    Dim tbl As DataTable = ASCDATA1.GetDataTable

                    Dim udd As New UltraWinGrid.UltraDropDown
                    udd.DataSource = tbl
                    udd.ValueMember = tbl.Columns(0).ColumnName
                    udd.DisplayMember = tbl.Columns(1).ColumnName
                    udd.DisplayLayout.Bands(0).Columns(0).Header.Caption = ASCMAIN1.CodeSelector.grdColumns(0).Item(5)
                    udd.DisplayLayout.Bands(0).Columns(1).Header.Caption = ASCMAIN1.CodeSelector.grdColumns(1).Item(5)
                    udd.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.SortSingle
                    grd.ActiveRow.Cells("CODE_VALUES").ValueList = udd
                    grd.ActiveRow.Cells("CODE_VALUES").Style = UltraWinGrid.ColumnStyle.DropDownList
                Else
                    grd.ActiveRow.Cells("CODE_VALUES").ValueList = Nothing
                    grd.ActiveRow.Cells("CODE_VALUES").Style = UltraWinGrid.ColumnStyle.EditButton
                End If

                'Case "Load List"
                '    Dim LIST_CODE As String = View_Lookup(Nothing, "LIST_CODE", "", "", "COLUMN_NAME = '" & grdASTFLTR1.ActiveRow.Cells("COLUMN_NAME").Text & "'")
                '    If LIST_CODE <> "" Then
                '        ASCMAIN1.sql = "Select CODE_VALUE from ASTLIST2 where LIST_CODE = '" & LIST_CODE & "'"
                '        Dim CODE_VALUEs As String = ""
                '        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                '            CODE_VALUEs &= "," & row.Item("CODE_VALUE")
                '        Next
                '        grdASTFLTR1.ActiveRow.Cells("CODE_VALUES").Value = Mid(CODE_VALUEs, 2)
                '        grdASTFLTR1.UpdateData()
                '    End If

                'Case "Maintain Lists"
                '    ASCMAIN1.CodeSelector.Get_SQL(grdASTFLTR1.ActiveRow.Cells("COLUMN_NAME").Text)
                '    ASCMAIN1.CodeSelector.MultipleSelections = True
                '    Using frmASFCODE1 As New ASFCODE1
                '        frmASFCODE1.ShowDialog()
                '    End Using


            Case Else
                'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        'Select Case Absx1.GetABSColumnName(sender)
        '    Case "OPS_YYYYPP"
        '        If e.KeyCode = Windows.Forms.Keys.Enter Then
        '            Call Click_Command("Load", e)
        '        End If
        'End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "OPS_YYYYPP"
                Call Click_Command("Load")
        End Select
    End Sub
#End Region

    Private Sub chkFullScreen_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFullScreen.CheckedChanged
        splSummary.Panel2Collapsed = chkFullScreen.Checked
    End Sub

    Private Sub optTD_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTD.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        ReSummarize()
    End Sub

    Private Sub tabMain_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()
        If SELECTION_NO = 0 Then Exit Sub
        UltraExplorerBar1.Groups("Summary").Visible = (tabMain.SelectedTab.Key = "Summary") And ScreenMode
        UltraExplorerBar1.Groups("Top N").Visible = (tabMain.SelectedTab.Key = "Top N") And ScreenMode
        'UltraExplorerBar1.Groups("Top/Bottom").Visible = (tabMain.SelectedTab.Key = "Top/Bottom") And ScreenMode

        'Select Case tabMain.SelectedTab.Key
        '    Case "Summary"
        '        tabBottom.Parent = splSummary.Panel2
        '    Case "Top N"
        '        tabBottom.Parent = splTopN.Panel2
        '    Case "Top/Bottom"
        '        tabBottom.Parent = spltopbottom.Panel2
        'End Select
    End Sub

    Function Setup_RSTSLSA0(ByVal RYP As String, Optional ByVal initialize As Boolean = False) As String

        ASCMAIN1.Get_Period_Range(-24, YPPD, YPP, RYP)

        ASCMAIN1.sql = "Select MAX(YYYYWW) YYYYWW from GLTPARM3 where YYYYPP = :PARM1"
        Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", RYP)
        OPS_YYYYWW_BASE = row.Item("YYYYWW")

        Dim Period_Number As Int32 = 0

        If optCalendar.Value = "R" Then
            Dim RYM As String = ASCMAIN1.Get_YYYYMM(RYP, 0)
            Period_Number = (Mid(RYM, 5, 2)) - 1
            If Period_Number = 0 Then Period_Number = 12
        Else
            Period_Number = (Mid(RYP, 5, 2))
        End If

        Dim sqlRSTSLSA0 As String = "Select" _
        & " RSTRETL1.ITEM_CODE, RSTRETL1.CUST_CODE" & vbCrLf
        For P As Int32 = 0 To 24
            For Each C As String In New String() {"QTY_SOLD", "AMT_SOLD", "QTY_EOW", "AMT_EOW"}
                Dim C2 As String = Mid(C, 1, 3)
                If C = "QTY_EOW" Then
                    If optSales.Value = "W" Then
                        sqlRSTSLSA0 &= ", 0 " & C2 & Format(P, "00") & vbCrLf
                    Else
                        C2 = "QOH"
                        If P = 0 Then
                            ASCMAIN1.sql = "Select Max (OPS_YYYYWW) from RSTRETL4 where OPS_YYYYPP = '" & YPP(P, 0) & "'"
                            Dim YW As String = ASCDATA1.GetDataValue
                            Dim YWZ As String = "'" & YW & "'"
                            If RYW_LGI <> "" Then YWZ = "DECODE(EDT852T1.EDI_CUST_BATCH_NO,'LGI','" & RYW_LGI & "','" & YW & "')"
                            sqlRSTSLSA0 &= ", Sum (DECODE(RSTRETL1.OPS_YYYYWW," & YWZ & ",RSTRETL1." & C & ",0)) " & C2 & Format(P, "00") & vbCrLf
                            'sqlRSTSLSA0 &= ", Sum (DECODE(RSTRETL1.OPS_YYYYWW,X.YW,RSTRETL1." & C & ",0)) " & C2 & Format(P, "00") & vbCrLf
                        Else
                            ASCMAIN1.sql = "Select Max (YYYYWW) from GLTPARM3 where YYYYPP = '" & YPP(P, 0) & "'"
                            Dim YW As String = ASCDATA1.GetDataValue
                            sqlRSTSLSA0 &= ", Sum (DECODE(RSTRETL1.OPS_YYYYWW,'" & YW & "',RSTRETL1." & C & ",0)) " & C2 & Format(P, "00") & vbCrLf
                        End If
                    End If
                ElseIf C = "AMT_EOW" Then
                    If optSales.Value = "W" Then
                        sqlRSTSLSA0 &= ", 0 " & C2 & Format(P, "00") & vbCrLf
                    Else
                        C2 = "AOH"
                        Dim Camt As String = "QTY_EOW * ICTITEM1.ITEM_RETAIL_PRICE"
                        If P = 0 Then
                            ASCMAIN1.sql = "Select Max (OPS_YYYYWW) from RSTRETL4 where OPS_YYYYPP = '" & YPP(P, 0) & "'"
                            Dim YW As String = ASCDATA1.GetDataValue
                            Dim YWZ As String = "'" & YW & "'"
                            If RYW_LGI <> "" Then YWZ = "DECODE(EDT852T1.EDI_CUST_BATCH_NO,'LGI','" & RYW_LGI & "','" & YW & "')"
                            sqlRSTSLSA0 &= ", Sum (DECODE(RSTRETL1.OPS_YYYYWW," & YWZ & ",RSTRETL1." & Camt & ",0)) " & C2 & Format(P, "00") & vbCrLf
                            'sqlRSTSLSA0 &= ", Sum (DECODE(RSTRETL1.OPS_YYYYWW,X.YW,RSTRETL1." & C & ",0)) " & C2 & Format(P, "00") & vbCrLf
                        Else
                            ASCMAIN1.sql = "Select Max (YYYYWW) from GLTPARM3 where YYYYPP = '" & YPP(P, 0) & "'"
                            Dim YW As String = ASCDATA1.GetDataValue
                            sqlRSTSLSA0 &= ", Sum (DECODE(RSTRETL1.OPS_YYYYWW,'" & YW & "',RSTRETL1." & Camt & ",0)) " & C2 & Format(P, "00") & vbCrLf
                        End If
                    End If
                Else
                    If optSales.Value = "W" Then
                        If C = "QTY_SOLD" Then
                            C = "ORDR_QTY_SHIP"
                        Else
                            C = "ORDR_QTY_SHIP * ORDR_UNIT_PRICE"
                        End If
                        sqlRSTSLSA0 &= ", Sum (DECODE(RSTRETL1.OPS_YYYYPP,'" & YPP(P, 0) & "',RSTRETL1." & C & ",0)) " & C2 & Format(P, "00") & vbCrLf
                    Else
                        sqlRSTSLSA0 &= ", Sum (DECODE(RSTRETL1.OPS_YYYYPP,'" & YPP(P, 0) & "',RSTRETL1." & C & ",0)) " & C2 & Format(P, "00") & vbCrLf
                    End If
                End If
            Next
        Next
        For Each TD As String In New String() {"M", "Q", "S", "Y"}
            Dim Period_Number_in_TD As Int32 = 0
            If TD = "M" Then
                Period_Number_in_TD = ((Period_Number - 1) Mod 1) + 1 ' ALWAYS 1
            ElseIf TD = "Q" Then
                Period_Number_in_TD = ((Period_Number - 1) Mod 3) + 1
            ElseIf TD = "S" Then
                Period_Number_in_TD = ((Period_Number - 1) Mod 6) + 1
            ElseIf TD = "Y" Then
                Period_Number_in_TD = ((Period_Number - 1) Mod 12) + 1 ' ALWAYS PN
            End If

            Dim YP1 As String = YPP(0, 0)
            Dim YP0 As String = YPP(Period_Number_in_TD - 1, 0)

            For Each Y As String In New String() {"TY", "LY"}
                If Y = "LY" Then
                    YP0 = ASCMAIN1.Period_Calc(YP0, -12)
                    YP1 = ASCMAIN1.Period_Calc(YP1, -12)
                End If


                ASCMAIN1.sql = "Select Max (YYYYWW) from GLTPARM3 where YYYYPP = '" & YP1 & "'"
                If Y = "TY" Then
                    ASCMAIN1.sql = "Select Max (OPS_YYYYWW) from RSTRETL4 where OPS_YYYYPP = '" & YP1 & "'"
                End If
                Dim YW As String = ASCDATA1.GetDataValue
                Dim YWZ As String = "'" & YW & "'"
                If RYW_LGI <> "" And Y = "TY" Then YWZ = "DECODE(EDT852T1.EDI_CUST_BATCH_NO,'LGI','" & RYW_LGI & "','" & YW & "')"

                If optSales.Value = "W" Then
                    Dim sqlc As String = "CASE WHEN RSTRETL1.OPS_YYYYPP BETWEEN '" & YP0 & "' AND '" & YP1 & "' THEN "
                    sqlRSTSLSA0 &= "" _
                    & ", Sum (" & sqlc & "RSTRETL1.ORDR_QTY_SHIP ELSE 0 END) " & Y & "_QTY_" & TD & "TD" & vbCrLf _
                    & ", Sum (" & sqlc & "RSTRETL1.ORDR_QTY_SHIP * RSTRETL1.ORDR_UNIT_PRICE ELSE 0 END) " & Y & "_AMT_" & TD & "TD" & vbCrLf _
                    & ", 0 " & Y & "_QOH_" & TD & "TD" & vbCrLf _
                    & ", 0 " & Y & "_AOH_" & TD & "TD" & vbCrLf
                Else
                    Dim sqlc As String = "CASE WHEN RSTRETL1.OPS_YYYYPP BETWEEN '" & YP0 & "' AND '" & YP1 & "' THEN "
                    sqlRSTSLSA0 &= "" _
                    & ", Sum (" & sqlc & "RSTRETL1.QTY_SOLD ELSE 0 END) " & Y & "_QTY_" & TD & "TD" & vbCrLf _
                    & ", Sum (" & sqlc & "RSTRETL1.AMT_SOLD ELSE 0 END) " & Y & "_AMT_" & TD & "TD" & vbCrLf _
                    & ", Sum (CASE WHEN RSTRETL1.OPS_YYYYWW = " & YWZ & " Then RSTRETL1.QTY_EOW ELSE 0 END) " & Y & "_QOH_" & TD & "TD" & vbCrLf _
                    & ", Sum (CASE WHEN RSTRETL1.OPS_YYYYWW = " & YWZ & " Then RSTRETL1.QTY_EOW * ICTITEM1.ITEM_RETAIL_PRICE ELSE 0 END) " & Y & "_AOH_" & TD & "TD" & vbCrLf
                End If
            Next
        Next

        'Dim sqlRSTSLSA0_filter As String = ""
        'For Each rowASTFLTR1 As DataRow In dst.Tables("ASTFLTR1").Rows
        '    If rowASTFLTR1.Item("CODE_VALUES") & "" <> "" Then
        '        Dim COLUMN_NAME As String = rowASTFLTR1.Item("COLUMN_NAME")
        '        Dim TABLE_NAME As String = "ICTITEM1"
        '        If COLUMN_NAME = "CUST_CODE" Then
        '            TABLE_NAME = "RSTRETL1"
        '        End If
        '        sqlRSTSLSA0_filter &= " and " & rowASTFLTR1.Item("COLUMN_NAME") & " in ('" & Replace(rowASTFLTR1.Item("CODE_VALUES"), ",", "','") & "')"
        '    End If
        'Next
        'Dim sqlRSTSLSB0_filter As String = ""
        'For Each rowASTFLTR1 As DataRow In dst.Tables("ASTFLTR1").Rows
        '    If rowASTFLTR1.Item("CODE_VALUES") & "" <> "" Then
        '        Dim COLUMN_NAME As String = rowASTFLTR1.Item("COLUMN_NAME")
        '        Dim TABLE_NAME As String = "ICTITEM1"
        '        If COLUMN_NAME = "CUST_CODE" Then
        '            TABLE_NAME = "RSTSLSA0"
        '        End If
        '        sqlRSTSLSB0_filter &= " and " & rowASTFLTR1.Item("COLUMN_NAME") & " in ('" & Replace(rowASTFLTR1.Item("CODE_VALUES"), ",", "','") & "')"
        '    End If
        'Next


        'SELECT CUST_CODE, MAX(OPS_YYYYWW) YW FROM RSTRETL1
        'WHERE OPS_YYYYPP = '200903'
        'GROUP BY CUST_CODE

        '        & ", (Select CUST_CODE, Max(OPS_YYYYWW) YW from RSTRETL1 " _
        '        & " where OPS_YYYYPP = '" & YPP(0, 0) & "' group by CUST_CODE) X" & vbCrLf _
        '        & "   and X.CUST_CODE (+) = RSTRETL1.CUST_CODE" & vbCrLf _

        sqlRSTSLSA0 &= "" _
        & " from RSTRETL1," & ICTITEM1 & " ICTITEM1" & IIf(optSales.Value = "W", "", ",EDT852T1") & vbCrLf _
        & " where RSTRETL1.OPS_YYYYPP BETWEEN '" & YPP(24, 0) & "' AND '" & YPP(0, 0) & "'" & vbCrLf _
        & " and ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE" & vbCrLf _
        & IIf(optSales.Value = "W", "", " and EDT852T1.EDI_DOC_SEQ_NO (+) = RSTRETL1.EDI_DOC_SEQ_NO" & vbCrLf) _
        & IIf(initialize, " and ROWNUM <1", "") & vbCrLf _
        & Replace(filter_CUST_CODE, "X.", "RSTRETL1.") _
        & " group by" & vbCrLf _
        & " RSTRETL1.ITEM_CODE, RSTRETL1.CUST_CODE"


        If optSales.Value = "W" Then
            sqlRSTSLSA0 = Replace(sqlRSTSLSA0, "OPS_YYYYPP", "ORDR_YYYYPP_UPDATED")
            sqlRSTSLSA0 = Replace(sqlRSTSLSA0, "RSTRETL1", "SOTINVH2")
        End If

        If RSTSLSA0 = "" Then
            RSTSLSA0 = ASCMAIN1.Temp_Table(sqlRSTSLSA0)

            ASCMAIN1.sql = "Select * from " & RSTSLSA0 & " where ROWNUM < 1"
            Dim DT As DataTable = ASCDATA1.GetDataTable
            sqlRSTSLSB0 = "Select RSTSLSA0.ITEM_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.DEPT_CODE" & vbCrLf
            For I As Int32 = 2 To DT.Columns.Count - 1
                Dim COLUMN_NAME As String = DT.Columns(I).ColumnName
                sqlRSTSLSB0 &= ", Sum (RSTSLSA0." & COLUMN_NAME & ") " & COLUMN_NAME & vbCrLf
            Next
            sqlRSTSLSB0 &= " from " & RSTSLSA0 & " RSTSLSA0, " & ICTITEM1 & " ICTITEM1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = RSTSLSA0.ITEM_CODE" & vbCrLf _
            & " group by RSTSLSA0.ITEM_CODE" _
            & ", ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.DEPT_CODE" & vbCrLf

            RSTSLSB0 = ASCMAIN1.Temp_Table(sqlRSTSLSB0)
            ASCDATA1.ExecuteSQL("Alter Table " & RSTSLSB0 & " Add Primary Key (ITEM_CODE)")
        Else
            sqlRSTSLSA0 = Replace(sqlRSTSLSA0, "ICTITEM1 ICTITEM1", ICTITEM1 & " ICTITEM1")
            sqlRSTSLSB0 = Replace(sqlRSTSLSB0, "ICTITEM1 ICTITEM1", ICTITEM1 & " ICTITEM1")

            ASCDATA1.ExecuteSQL("Truncate Table " & RSTSLSA0)
            ASCDATA1.ExecuteSQL("Insert into " & RSTSLSA0 & " " & sqlRSTSLSA0)

            ASCDATA1.ExecuteSQL("Truncate Table " & RSTSLSB0)
            ASCDATA1.ExecuteSQL("Insert into " & RSTSLSB0 & " " & sqlRSTSLSB0)
        End If

        Return sqlRSTSLSA0
    End Function

    Function Setup_RSTSLSAX(Optional ByVal fill_RSTSLSA1 As Boolean = True) As String

        Dim TD As String = optTD.Value

        Dim sqlRSTSLSB1x As String = ""

        If optItem.Value = "ITEM_CODE" Then
            sqlRSTSLSB1x = "Select" _
            & "  RSTSLSB0.ITEM_CODE, ICTITEM1.ITEM_DESC, RSTSLSB0.ITEM_CATGY_CODE, RSTSLSB0.DEPT_CODE" _
            & ", ICTITEM1.STYLE_CODE, ICTITEM1.STYLE_CODE_NC, ICTITEM1.COLOR_CODE, ICTITEM1.SIZE_CODE" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_CLASS_CODE, ICTITEM1.MATL_CODE" & vbCrLf _
            & ", PRICE_POINT(ICTITEM1.ITEM_RETAIL_PRICE) PRICE_POINT_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_PRICE, ICTITEM1.ITEM_ABC_CODE" & vbCrLf
        Else
            sqlRSTSLSB1x = "Select" _
            & "  ICTITEM1." & optItem.Value & " ITEM_CODE, MIN (ICTITEM1.ITEM_DESC) ITEM_DESC, MIN (ICTITEM1.ITEM_CATGY_CODE) ITEM_CATGY_CODE, MIN (ICTITEM1.DEPT_CODE) DEPT_CODE" _
            & ", MIN (" & IIf(optItem.Value = "STYLE_CODE", "ICTITEM1.ITEM_CODE", "ICTITEM1.STYLE_CODE") & ") STYLE_CODE" _
            & ", MIN (" & IIf(optItem.Value = "STYLE_CODE_NC", "ICTITEM1.ITEM_CODE", "ICTITEM1.STYLE_CODE_NC") & ") STYLE_CODE_NC" _
            & ", MIN (ICTITEM1.COLOR_CODE) COLOR_CODE, MIN (ICTITEM1.SIZE_CODE) SIZE_CODE" & vbCrLf _
            & ", MIN (ICTITEM1.COLLECTION_CODE) COLLECTION_CODE, MIN (ICTCOLL1.HC_CODE) HC_CODE, MIN (ICTITEM1.PROD_CODE) PROD_CODE, MIN (ICTITEM1.ITEM_CLASS_CODE) ITEM_CLASS_CODE, MIN (ICTITEM1.MATL_CODE) MATL_CODE" & vbCrLf _
            & ", MIN (PRICE_POINT(ICTITEM1.ITEM_RETAIL_PRICE)) PRICE_POINT_CODE" & vbCrLf _
            & ", MIN (ICTITEM1.ITEM_RETAIL_PRICE) ITEM_RETAIL_PRICE, MIN (ICTITEM1.ITEM_PRICE) ITEM_PRICE, MIN (ICTITEM1.ITEM_ABC_CODE) ITEM_ABC_CODE" & vbCrLf
        End If

        ''DGJ 
        ''If optItem.Value = "ITEM_CODE" Then
        ''    sqlRSTSLSB1x = "Select" _
        ''    & "  RSTSLSB0.ITEM_CODE, ICTITEM1.ITEM_DESC, RSTSLSB0.ITEM_CATGY_CODE, RSTSLSB0.DEPT_CODE, RSTSLSB0.CUST_CODE" _
        ''    & ", ICTITEM1.STYLE_CODE, ICTITEM1.STYLE_CODE_NC, ICTITEM1.COLOR_CODE, ICTITEM1.SIZE_CODE" & vbCrLf _
        ''    & ", ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_CLASS_CODE, ICTITEM1.MATL_CODE" & vbCrLf _
        ''    & ", PRICE_POINT(ICTITEM1.ITEM_RETAIL_PRICE) PRICE_POINT_CODE" & vbCrLf _
        ''    & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_PRICE, ICTITEM1.ITEM_ABC_CODE, SOTTCLS1.MARKET_CODE" & vbCrLf
        ''Else
        ''    sqlRSTSLSB1x = "Select" _
        ''    & "  ICTITEM1." & optItem.Value & " ITEM_CODE, MIN (ICTITEM1.ITEM_DESC) ITEM_DESC, MIN (ICTITEM1.ITEM_CATGY_CODE) ITEM_CATGY_CODE, MIN (ICTITEM1.DEPT_CODE) DEPT_CODE, MIN (ICTITEM1.CUST_CODE) CUST_CODE" _
        ''    & ", MIN (" & IIf(optItem.Value = "STYLE_CODE", "ICTITEM1.ITEM_CODE", "ICTITEM1.STYLE_CODE") & ") STYLE_CODE" _
        ''    & ", MIN (" & IIf(optItem.Value = "STYLE_CODE_NC", "ICTITEM1.ITEM_CODE", "ICTITEM1.STYLE_CODE_NC") & ") STYLE_CODE_NC" _
        ''    & ", MIN (ICTITEM1.COLOR_CODE) COLOR_CODE, MIN (ICTITEM1.SIZE_CODE) SIZE_CODE" & vbCrLf _
        ''    & ", MIN (ICTITEM1.COLLECTION_CODE) COLLECTION_CODE, MIN (ICTCOLL1.HC_CODE) HC_CODE, MIN (ICTITEM1.PROD_CODE) PROD_CODE, MIN (ICTITEM1.ITEM_CLASS_CODE) ITEM_CLASS_CODE, MIN (ICTITEM1.MATL_CODE) MATL_CODE" & vbCrLf _
        ''    & ", MIN (PRICE_POINT(ICTITEM1.ITEM_RETAIL_PRICE)) PRICE_POINT_CODE" & vbCrLf _
        ''    & ", MIN (ICTITEM1.ITEM_RETAIL_PRICE) ITEM_RETAIL_PRICE, MIN (ICTITEM1.ITEM_PRICE) ITEM_PRICE, MIN (ICTITEM1.ITEM_ABC_CODE) ITEM_ABC_CODE, MIN (SOTTCLS1.MARKET_CODE) MARKET_CODE" & vbCrLf
        ''End If


        sqlRSTSLSBR = sqlRSTSLSB1x

        sqlRSTSLSB2 = "Select" _
        & "  RSTSLSA0.ITEM_CODE, RSTSLSA0.CUST_CODE, ARTCUST1.CUST_NAME"


        'Dim sqlRSTSLSAX_filter As String = ""
        'For Each rowASTFLTR1 As DataRow In dst.Tables("ASTFLTR1").Rows
        '    If rowASTFLTR1.Item("CODE_VALUES") & "" <> "" Then
        '        sqlRSTSLSAX_filter &= " and " & rowASTFLTR1.Item("COLUMN_NAME") & " in ('" & Replace(rowASTFLTR1.Item("CODE_VALUES"), ",", "','") & "')"
        '    End If
        'Next
        '        & sqlRSTSLSAX_filter _

        sqlRSTSLSA3 = "Select X.X LINK_VALUE, RSTSLSA0.ITEM_CODE CODE_VALUE, ICTITEM1.ITEM_DESC DESC_VALUE"

        Dim sqlw As String = ""
        Dim sqlRSTSLSAX As String = "Select" _
        & " ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_CLASS_CODE, RSTSLSA0.CUST_CODE" _
        & ", PRICE_POINT(ICTITEM1.ITEM_RETAIL_PRICE) PRICE_POINT_CODE" _
        & ", ICTITEM1.MATL_CODE, ICTITEM1.DEPT_CODE, ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.ITEM_ABC_CODE, SOTTCLS1.MARKET_CODE, ICTITEM1.ITEM_CODE"
        For Each Y As String In New String() {"TY", "LY"}
            For Each C As String In New String() {"QTY", "AMT", "QOH", "AOH"}
                sqlRSTSLSAX &= ", Sum (RSTSLSA0." & Y & "_" & C & "_" & TD & "TD" & ") " & Y & "_" & C
                sqlRSTSLSA3 &= ", Sum (RSTSLSA0." & Y & "_" & C & "_" & TD & "TD" & ") " & Y & "_" & C
                sqlRSTSLSB1x &= ", Sum (RSTSLSB0." & Y & "_" & C & "_" & TD & "TD" & ") " & Y & "_" & C
                sqlRSTSLSB2 &= ", Sum (RSTSLSA0." & Y & "_" & C & "_" & TD & "TD" & ") " & Y & "_" & C
                sqlw &= " or ." & Y & "_" & C & "_" & TD & "TD" & " <> 0"
            Next
        Next

        'sqlRSTSLSBR &= "{}"

        If chkShow60.Checked Or Not ScreenMode Then
            ReDim WKS60(60)
            Dim sql_sums As String = ""
            Dim OPS_YYYYWW As String = ""
            For W As Integer = 1 To 60
                OPS_YYYYWW = ASCMAIN1.Week_Calc(OPS_YYYYWW_BASE, -60 + W)
                WKS60(W) = OPS_YYYYWW
                sql_sums &= ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,'" & OPS_YYYYWW & "',NVL(RSTRETL1.QTY_SOLD,0),0)) SLS" & Format(W, "00")
                sql_sums &= ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,'" & OPS_YYYYWW & "',NVL(RSTRETL1.QTY_EOW,0),0)) EOW" & Format(W, "00")
                If grdRSTSLSBR.DisplayLayout.Bands(0).Groups.Count <> 0 Then
                    Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", WKS60(W))
                    grdRSTSLSBR.DisplayLayout.Bands(0).Groups("W" & Format(W, "00")).Header.Caption = rowGLTPARM3.Item("LEGEND")
                End If
            Next
            sqlRSTSLSBR &= sql_sums
        End If

        ' sqlw &= " OR 1=1" ' THIS WAS PUT INTO PLACE BECAUSE THE EXPORT TO EXCEL BY CUSTOMER WAS SHOWING SKUS BY CUSTOMER THAT DID NOT APPEAR IN THE TOTAL BECAUSE THE TOTAL NETTED TO 0, SO WE NEED ALL OF THE ROWS AND WE NEED TO FILTER THEM ONCE THEY ARE IN THE DATATABLE

        sqlRSTSLSAX &= "" _
        & " from " & RSTSLSA0 & " RSTSLSA0," & ICTITEM1 & " ICTITEM1, ICTCOLL1, ARTCUST1, SOTTCLS1" _
        & " where ICTITEM1.ITEM_CODE = RSTSLSA0.ITEM_CODE" _
        & "   and (" & Mid(Replace(sqlw, ".", "RSTSLSA0."), 4) & ")" _
        & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" _
        & "   and ARTCUST1.CUST_CODE = RSTSLSA0.CUST_CODE" _
        & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE " _
        & " group by" _
        & " ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_CLASS_CODE, RSTSLSA0.CUST_CODE" _
        & ", PRICE_POINT(ICTITEM1.ITEM_RETAIL_PRICE)" _
        & ", ICTITEM1.MATL_CODE, ICTITEM1.DEPT_CODE, ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.ITEM_ABC_CODE, SOTTCLS1.MARKET_CODE, ICTITEM1.ITEM_CODE"

        sqlRSTSLSA3 &= "" _
        & " from " & RSTSLSA0 & " RSTSLSA0," & ICTITEM1 & " ICTITEM1, ICTCOLL1, ARTCUST1, SOTTCLS1" _
        & " where ICTITEM1.ITEM_CODE = RSTSLSA0.ITEM_CODE" _
        & "   and (" & Mid(Replace(sqlw, ".", "RSTSLSA0."), 4) & ")" _
        & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" _
        & "   and ARTCUST1.CUST_CODE = RSTSLSA0.CUST_CODE" _
        & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" _
        & " group by" _
        & " X.X, RSTSLSA0.ITEM_CODE, ICTITEM1.ITEM_DESC"



        If RSTSLSAX = "" Then
            RSTSLSAX = ASCMAIN1.Temp_Table(sqlRSTSLSAX)
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & RSTSLSAX)
            ASCDATA1.ExecuteSQL("Insert into " & RSTSLSAX & " " & sqlRSTSLSAX)
        End If

        If fill_RSTSLSA1 Then
            Get_sqlRSTSLSAX("1")
        End If

        sqlRSTSLSB1x &= "" _
        & " from " & RSTSLSB0 & " RSTSLSB0," & ICTITEM1 & " ICTITEM1, ICTCOLL1" _
        & " where ICTITEM1.ITEM_CODE = RSTSLSB0.ITEM_CODE" _
        & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE"
        '  & "   and ARTCUST1.CUST_CODE = RSTSLSB0.CUST_CODE" _
        '  & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE"
        If sql_filter_no_CUST_CODE <> "" Then
            sqlRSTSLSB1x &= "   and (" & Mid(Replace(sqlw, ".", "RSTSLSB0."), 4) & ")"
            'sqlRSTSLSB2
        End If

        sqlRSTSLSBR &= "" _
        & " from RSTRETL1," & RSTSLSB0 & " RSTSLSB0," & ICTITEM1 & " ICTITEM1, ICTCOLL1" _
        & " where ICTITEM1.ITEM_CODE = RSTSLSB0.ITEM_CODE" _
        & "   and RSTRETL1.ITEM_CODE = ICTITEM1.ITEM_CODE" _
        & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE"
        ' & "   and ARTCUST1.CUST_CODE = ICTITEM1.CUST_CODE" _
        ' & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE"

        If sql_filter_no_CUST_CODE <> "" Then
        'sqlRSTSLSBR &= " and RSTRETL1.CUST_CODE in ()"
        End If
        sqlRSTSLSBR &= " and RSTRETL1.OPS_YYYYWW between '" & ASCMAIN1.Week_Calc(OPS_YYYYWW_BASE, -60 + 1) & "' and '" & OPS_YYYYWW_BASE & "'"


        sqlRSTSLSB2 &= "" _
        & " from " & RSTSLSA0 & " RSTSLSA0,ARTCUST1" _
        & " where ARTCUST1.CUST_CODE = RSTSLSA0.CUST_CODE" _
        & "   and RSTSLSA0.ITEM_CODE = :PARM1"
        If sql_filter_no_CUST_CODE <> "" Then
            sqlRSTSLSB2 &= "   and (" & Mid(Replace(sqlw, ".", "RSTSLSA0."), 4) & ")"
        End If

        If optItem.Value = "ITEM_CODE" Then
            sqlRSTSLSB1x &= "" _
            & " group by" _
            & "  RSTSLSB0.ITEM_CODE, ICTITEM1.ITEM_DESC, RSTSLSB0.ITEM_CATGY_CODE, RSTSLSB0.DEPT_CODE" & vbCrLf _
            & ", ICTITEM1.STYLE_CODE, ICTITEM1.STYLE_CODE_NC, ICTITEM1.COLOR_CODE, ICTITEM1.SIZE_CODE" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_CLASS_CODE, ICTITEM1.MATL_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_PRICE, ICTITEM1.ITEM_ABC_CODE" & vbCrLf

            sqlRSTSLSBR &= "" _
             & " group by" _
             & "  RSTSLSB0.ITEM_CODE, ICTITEM1.ITEM_DESC, RSTSLSB0.ITEM_CATGY_CODE, RSTSLSB0.DEPT_CODE" & vbCrLf _
             & ", ICTITEM1.STYLE_CODE, ICTITEM1.STYLE_CODE_NC, ICTITEM1.COLOR_CODE, ICTITEM1.SIZE_CODE" & vbCrLf _
             & ", ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_CLASS_CODE, ICTITEM1.MATL_CODE" & vbCrLf _
             & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_PRICE, ICTITEM1.ITEM_ABC_CODE" & vbCrLf
            '' DGJ
            ''sqlRSTSLSB1x &= "" _
            ''& " group by" _
            ''& "  RSTSLSB0.ITEM_CODE, ICTITEM1.ITEM_DESC, RSTSLSB0.ITEM_CATGY_CODE, RSTSLSB0.DEPT_CODE, RSTSLSB0.CUST_CODE" & vbCrLf _
            ''& ", ICTITEM1.STYLE_CODE, ICTITEM1.STYLE_CODE_NC, ICTITEM1.COLOR_CODE, ICTITEM1.SIZE_CODE" & vbCrLf _
            ''& ", ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_CLASS_CODE, ICTITEM1.MATL_CODE" & vbCrLf _
            ''& ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_PRICE, ICTITEM1.ITEM_ABC_CODE, SOTTCLS1.MARKET_CODE" & vbCrLf

            ''sqlRSTSLSBR &= "" _
            '' & " group by" _
            '' & "  RSTSLSB0.ITEM_CODE, ICTITEM1.ITEM_DESC, RSTSLSB0.ITEM_CATGY_CODE, RSTSLSB0.DEPT_CODE, RSTSLSB0.CUST_CODE" & vbCrLf _
            '' & ", ICTITEM1.STYLE_CODE, ICTITEM1.STYLE_CODE_NC, ICTITEM1.COLOR_CODE, ICTITEM1.SIZE_CODE" & vbCrLf _
            '' & ", ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_CLASS_CODE, ICTITEM1.MATL_CODE" & vbCrLf _
            '' & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_PRICE, ICTITEM1.ITEM_ABC_CODE, SOTTCLS1.MARKET_CODE" & vbCrLf

        Else
            sqlRSTSLSB1x &= "" _
            & " group by" _
            & "  ICTITEM1." & optItem.Value & vbCrLf

            sqlRSTSLSBR &= "" _
           & " group by" _
           & "  ICTITEM1." & optItem.Value & vbCrLf
        End If

        sqlRSTSLSB2 &= "" _
        & " group by" _
        & "  RSTSLSA0.ITEM_CODE, RSTSLSA0.CUST_CODE, ARTCUST1.CUST_NAME"

        sqlRSTSLSB1 = sqlRSTSLSB1x

        If fill_RSTSLSA1 Then
            splRSTSLSB1.Panel2Collapsed = True
            dst.Tables("RSTSLSB2").Rows.Clear()
            Fill_Records("RSTSLSB1", , , Replace(sqlRSTSLSB1, " group by ", sql_filter_no_CUST_CODE & " group by "))
            Calculate_Rank(dst.Tables("RSTSLSB1"))
        End If

        '-------------------
        ''DGJ
        ''      & " ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_CLASS_CODE, RSTSLSA0.CUST_CODE" _

        Dim SQL As String = "Select" _
        & " ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_CLASS_CODE, RSTSLSA0.CUST_CODE" _
        & ", PRICE_POINT(ICTITEM1.ITEM_RETAIL_PRICE) PRICE_POINT_CODE" _
        & ", ICTITEM1.MATL_CODE, ICTITEM1.DEPT_CODE, ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.ITEM_ABC_CODE, SOTTCLS1.MARKET_CODE, ICTITEM1.ITEM_CODE"
        For P As Integer = 0 To 24
            For Each C As String In New String() {"QTY", "AMT", "QOH"}
                SQL &= ", Sum (RSTSLSA0." & C & Format(P, "00") & ") " & C & Format(P, "00")
            Next
        Next
        SQL &= "" _
        & " from " & RSTSLSA0 & " RSTSLSA0," & ICTITEM1 & " ICTITEM1, ICTCOLL1, ARTCUST1, SOTTCLS1" _
        & " where ICTITEM1.ITEM_CODE = RSTSLSA0.ITEM_CODE" _
        & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" _
        & "   and ARTCUST1.CUST_CODE = RSTSLSA0.CUST_CODE" _
        & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" _
        & " group by" _
        & " ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_CLASS_CODE, RSTSLSA0.CUST_CODE" _
        & ", PRICE_POINT(ICTITEM1.ITEM_RETAIL_PRICE)" _
        & ", ICTITEM1.MATL_CODE, ICTITEM1.DEPT_CODE, ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.ITEM_ABC_CODE, SOTTCLS1.MARKET_CODE, ICTITEM1.ITEM_CODE"
        ''DGJ
        ''       & " ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_CLASS_CODE, RSTSLSA0.CUST_CODE" _

        If RSTSLSAQ = "" Then
            RSTSLSAQ = ASCMAIN1.Temp_Table(SQL)
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & RSTSLSAQ)
            ASCDATA1.ExecuteSQL("Insert into " & RSTSLSAQ & " " & SQL)
        End If

        '-------------------

        Return sqlRSTSLSAX
    End Function

    Function Get_sqlRSTSLSAX(ByVal A12 As String, Optional ByVal fill As Boolean = True)

        Dim grd As UltraWinGrid.UltraGrid

        Dim COLUMN_NAME As String = ""

        Select Case A12
            Case "1"
                grd = grdRSTSLSA1
                COLUMN_NAME = grdRSTSLSA1.Tag
            Case Else
                grd = grdRSTSLSA2
                COLUMN_NAME = grdRSTSLSA2.Tag
        End Select

        Dim TABLE_NAME As String = ""
        Dim COLUMN_NAME_DESC As String = ""
        Dim COLUMN_NAME_CODE2 As String = ""
        Dim COLUMN_NAME_KEY As String = ""

        Select Case COLUMN_NAME
            Case "COLLECTION_CODE"
                TABLE_NAME = "ICTCOLL1"
                COLUMN_NAME_DESC = "COLLECTION_NAME"
                COLUMN_NAME_CODE2 = "BRAND_CODE"
                Set_Caption(grd, "Collection", "Name", "Brand")
            Case "HC_CODE"
                TABLE_NAME = "ICTCOLL0"
                COLUMN_NAME_DESC = "HC_NAME"
                COLUMN_NAME_CODE2 = "NULL"
                Set_Caption(grd, "High Collection", "Name", "")
            Case "PROD_CODE"
                TABLE_NAME = "ICTPROD1"
                COLUMN_NAME_DESC = "PROD_DESC"
                COLUMN_NAME_CODE2 = "NULL"
                Set_Caption(grd, "Product", "Name", "")
            Case "ITEM_CLASS_CODE"
                TABLE_NAME = "ICTCLAS1"
                COLUMN_NAME_DESC = "ITEM_CLASS_DESC"
                COLUMN_NAME_CODE2 = "NULL"
                Set_Caption(grd, "Class", "Description", "")
            Case "CUST_CODE"
                TABLE_NAME = "ARTCUST1"
                COLUMN_NAME_DESC = "CUST_NAME"
                COLUMN_NAME_CODE2 = "SREP_CODE"
                Set_Caption(grd, "Customer", "Name", "SRep")
            Case "PRICE_POINT_CODE"
                TABLE_NAME = "ICTPRPT1"
                COLUMN_NAME_DESC = "PRICE_POINT_DESC"
                COLUMN_NAME_CODE2 = "NULL"
                Set_Caption(grd, "Code", "Price Range", "")
            Case "MATL_CODE"
                TABLE_NAME = "ICTMATL1"
                COLUMN_NAME_DESC = "MATL_DESC"
                COLUMN_NAME_CODE2 = "NULL"
                Set_Caption(grd, "Code", "Material", "")
            Case "ITEM_CATGY_CODE"
                TABLE_NAME = "ICTCATG1"
                COLUMN_NAME_DESC = "ITEM_CATGY_DESC"
                COLUMN_NAME_CODE2 = "NULL"
                Set_Caption(grd, "Code", "Category", "")
            Case "DEPT_CODE"
                TABLE_NAME = "ICTDEPT1"
                COLUMN_NAME_DESC = "DEPT_DESC"
                COLUMN_NAME_CODE2 = "NULL"
                Set_Caption(grd, "Code", "Department", "")
            Case "ITEM_ABC_CODE"
                TABLE_NAME = "DPTABCP1"
                COLUMN_NAME_DESC = "ABC_DESC"
                COLUMN_NAME_CODE2 = "NULL"
                Set_Caption(grd, "Code", "ABC", "")
                COLUMN_NAME_KEY = "ABC_CODE"
            Case "MARKET_CODE"
                TABLE_NAME = "SOTMKTC1"
                COLUMN_NAME_DESC = "MARKET_DESC"
                COLUMN_NAME_CODE2 = "NULL"
                Set_Caption(grd, "Code", "Market", "")
                COLUMN_NAME_KEY = "MARKET_CODE"
            Case "ITEM_CODE"
                TABLE_NAME = "ICTITEM1"
                COLUMN_NAME_DESC = "ITEM_DESC"
                COLUMN_NAME_CODE2 = "NULL"
                Set_Caption(grd, "Code", "Description", "")
                COLUMN_NAME_KEY = "ITEM_CODE"

        End Select

        If COLUMN_NAME_KEY = "" Then COLUMN_NAME_KEY = COLUMN_NAME

        Dim SQLW As String = ""
        If A12 = "2" Then
            SQLW = " And RSTSLSAX." & grdRSTSLSA1.Tag & " = '" & grdRSTSLSA1.ActiveRow.Cells("CODE_VALUE").Text & "'"
        End If

        Set_Filter()

        Dim sqlRSTSLSAX As String = "Select NVL(RSTSLSAX." & COLUMN_NAME & ",'?') CODE_VALUE" _
        & ", " & TABLE_NAME & "." & COLUMN_NAME_DESC & " DESC_VALUE" _
        & ", " & IIf(COLUMN_NAME_CODE2 = "NULL", "NULL", TABLE_NAME & "." & COLUMN_NAME_CODE2) & " CODE2_VALUE" _
        & ", SUM (TY_QTY) TY_QTY, SUM (TY_AMT) TY_AMT, SUM (TY_QOH) TY_QOH, SUM (TY_AOH) TY_AOH" _
        & ", SUM (LY_QTY) LY_QTY, SUM (LY_AMT) LY_AMT, SUM (LY_QOH) LY_QOH, SUM (LY_AOH) LY_AOH" _
        & " from " & RSTSLSAX & " RSTSLSAX," & TABLE_NAME _
        & " where " & TABLE_NAME & "." & COLUMN_NAME_KEY & " (+) = RSTSLSAX." & COLUMN_NAME _
        & Replace(sql_filter, "X.", "RSTSLSAX.") _
        & SQLW _
        & " group by NVL(RSTSLSAX." & COLUMN_NAME & ",'?')" _
        & ", " & TABLE_NAME & "." & COLUMN_NAME_DESC _
        & IIf(COLUMN_NAME_CODE2 = "NULL", "", ", " & TABLE_NAME & "." & COLUMN_NAME_CODE2)

        If A12 = "1" Then
            CN.Clear()
            CN.Add("COLUMN_NAME", COLUMN_NAME)
            CN.Add("TABLE_NAME", TABLE_NAME)
            CN.Add("COLUMN_NAME_DESC", COLUMN_NAME_DESC)
            CN.Add("COLUMN_NAME_CODE2", COLUMN_NAME_CODE2)
            CN.Add("COLUMN_NAME_KEY", COLUMN_NAME_KEY)
            Get_sqlRSTSLSAP(fill)
        End If


        If fill Then
            EnforceConstraints(False)
            If A12 = "2" Then
                dst.Tables("RSTSLSA3").Rows.Clear()
            End If
            Fill_Records("RSTSLSA" & A12, "", True, sqlRSTSLSAX)
            Calculate_Rank(dst.Tables("RSTSLSA" & A12))
            If Not excel_preparation Then
                Sort_grdColumns(grd, "CODE_VALUE")
            End If

            Dim CAPTION As String
            If A12 = "1" Then
                CAPTION = filter_desc & " by " & SUMCOLs(grdRSTSLSA1.Tag)
            Else
                If grdRSTSLSA1.ActiveRow Is Nothing Then
                    CAPTION = grdRSTSLSA1.DisplayLayout.Bands(0).Columns("CODE_VALUE").Header.Caption & " " & "?" & " by " & SUMCOLs(grdRSTSLSA2.Tag)
                Else
                    CAPTION = grdRSTSLSA1.DisplayLayout.Bands(0).Columns("CODE_VALUE").Header.Caption & " " & grdRSTSLSA1.ActiveRow.Cells("CODE_VALUE").Text & " by " & SUMCOLs(grdRSTSLSA2.Tag)
                End If
            End If
            grd.Text = CAPTION
            grd.DisplayLayout.CaptionVisible = DefaultableBoolean.False
            grd.DisplayLayout.Bands(0).Groups("CODES").Header.Caption = CAPTION
            grd.DisplayLayout.Bands(0).Groups("TY").Header.Caption = "This Year, " & optTD.Text
            grd.DisplayLayout.Bands(0).Groups("LY").Header.Caption = "Last Year, " & optTD.Text

            EnforceConstraints(True)
        End If

        Return sqlRSTSLSAX
    End Function

    Sub Get_sqlRSTSLSAP(ByVal fill As Boolean)

        Dim COLUMN_NAME As String = CN("COLUMN_NAME")
        Dim TABLE_NAME As String = CN("TABLE_NAME")
        Dim COLUMN_NAME_DESC As String = CN("COLUMN_NAME_DESC")
        Dim COLUMN_NAME_CODE2 As String = CN("COLUMN_NAME_CODE2")

        Dim COLUMN_NAME_KEY As String = COLUMN_NAME
        If COLUMN_NAME_KEY = "ITEM_ABC_CODE" Then COLUMN_NAME_KEY = "ABC_CODE"

        Dim sql As String = "Select NVL(RSTSLSAQ." & COLUMN_NAME & ",'?') CODE_VALUE" _
        & ", " & TABLE_NAME & "." & COLUMN_NAME_DESC & " DESC_VALUE" _
        & ", " & IIf(COLUMN_NAME_CODE2 = "NULL", "NULL", TABLE_NAME & "." & COLUMN_NAME_CODE2) & " CODE2_VALUE"
        For P As Integer = 0 To 24
            Dim C2 As String = optTrend.Value

            sql &= ", Sum (" & C2 & Format(P, "00") & ") P" & Format(P, "00") & vbCrLf
        Next
        sql &= "" _
        & " from " & RSTSLSAQ & " RSTSLSAQ," & TABLE_NAME _
        & " where " & TABLE_NAME & "." & COLUMN_NAME_KEY & " (+) = RSTSLSAQ." & COLUMN_NAME _
        & Replace(sql_filter, "X.", "RSTSLSAQ.") _
        & " group by NVL(RSTSLSAQ." & COLUMN_NAME & ",'?')" _
        & ", " & TABLE_NAME & "." & COLUMN_NAME_DESC _
        & IIf(COLUMN_NAME_CODE2 = "NULL", "", ", " & TABLE_NAME & "." & COLUMN_NAME_CODE2)
        sqlRSTSLSAP = sql

        If fill Then
            EnforceConstraints(False)
            Fill_Records("RSTSLSAP", "", True, sqlRSTSLSAP)
            Sort_grdColumns(grdRSTSLSAP, "CODE_VALUE")
            grdRSTSLSAP.Text = "25 Month Trend, Showing " & optTrend.Text & " By " & SUMCOLs(grdRSTSLSA1.Tag)

            EnforceConstraints(True)
        End If
        'Dim X As Integer = dst.Tables("RSTSLSA3").Rows.Count

    End Sub

    Sub Set_Caption(ByVal grd As UltraWinGrid.UltraGrid, _
                    ByVal CODE_VALUE_caption As String, _
                    ByVal DESC_VALUE_caption As String, _
                    ByVal C0DE2_VALUE_caption As String)

        If SELECTION_NO = 0 Then Exit Sub

        With grd.DisplayLayout.Bands(0)
            .Columns("CODE_VALUE").Header.Caption = CODE_VALUE_caption
            .Columns("DESC_VALUE").Header.Caption = DESC_VALUE_caption
            .Columns("CODE2_VALUE").Header.Caption = C0DE2_VALUE_caption
        End With

        If grd.Name = "grdRSTSLSA1" Then
            Set_Caption(grdRSTSLSAP, CODE_VALUE_caption, DESC_VALUE_caption, C0DE2_VALUE_caption)
        End If
    End Sub

    Private Sub chkShowLY_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowLY.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_Selected()
    End Sub

    Private Sub grdRSTSLSA1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdRSTSLSA1.AfterRowActivate
        Setup_grdRSTSLSA2()
    End Sub

    Sub Setup_grdRSTSLSA2()
        If grdRSTSLSA1.ActiveRow Is Nothing OrElse Not grdRSTSLSA1.ActiveRow.IsDataRow Then
            grdRSTSLSA2.Visible = False
            splChart.Visible = False
            grdRSTSLSAP.Visible = False
        Else
            grdRSTSLSA2.Visible = True
            splChart.Visible = True
            grdRSTSLSAP.Visible = True

            Get_sqlRSTSLSAX("2")

            If optChart.Value = "2" Then
                CreateGraph_Trend()
            End If
        End If
        grdRSTSLSA2.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand
    End Sub

    Sub Format_grids()

        Dim b As Int32 = 0
        For Each grd As UltraWinGrid.UltraGrid In _
        New UltraWinGrid.UltraGrid() {grdRSTSLSA1, grdRSTSLSA2, grdRSTSLSA2}

            Call Create_Summary(grd, "CODE_VALUE", "Count", grd.DisplayLayout.Bands(b).Key)

            Dim G As UltraWinGrid.UltraGridGroup

            With grd.DisplayLayout.Bands(b)
                G = .Groups.Add("CODES", "Codes")
                G.Header.Appearance.BackColor = Color.CornflowerBlue
                G.Header.Appearance.BackGradientStyle = GradientStyle.None
                G.Header.Appearance.ForeColor = Color.White
                G.Header.Fixed = True
                For Each COLUMN_NAME As String In New String() _
                    {"CODE_VALUE", "DESC_VALUE", "CODE2_VALUE"}
                    With .Columns(COLUMN_NAME)
                        .Group = G
                        .CellAppearance.BackColor = Drawing.Color.Beige
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                    End With
                Next
                .Columns("CODE_VALUE").Width = 90
                .Columns("DESC_VALUE").Width = 110
                .Columns("CODE2_VALUE").Width = 70

                For Each Y As String In New String() {"TY", "LY"}
                    G = .Groups.Add(Y, IIf(Y = "TY", "This Year", "Last Year"))
                    G.Header.Appearance.BackColor = IIf(Y = "TY", Color.LightGreen, Color.Orange)
                    G.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                    For Each COLUMN_NAME As String In New String() _
                         {"_QTY", "_AMT", "_QOH", "_AOH"}
                        With .Columns(Y & COLUMN_NAME)

                            Call Create_Summary(grd, Y & COLUMN_NAME, , grd.DisplayLayout.Bands(b).Key, "#,##0")

                            .Group = G
                            .CellActivation = UltraWinGrid.Activation.NoEdit
                            .Format = "#,##0"
                            If COLUMN_NAME Like "*AMT" Then
                                .Width = 90
                            Else
                                .Width = 70
                            End If
                        End With
                    Next

                    For Each COLUMN_NAME As String In New String() _
                        {"_SELL_THRU", "_PCT_TOTAL"}
                        With .Columns(Y & COLUMN_NAME)
                            .Group = G
                            .CellAppearance.BackColor = Drawing.Color.LightGray
                            .CellActivation = UltraWinGrid.Activation.NoEdit
                            .Format = "#,##0.0"
                            .Width = 60
                        End With
                    Next
                    Call Create_Summary(grd, Y & "_PCT_TOTAL", , grd.DisplayLayout.Bands(b).Key, "#,##0.0")

                    For Each COLUMN_NAME As String In New String() _
                        {"_QTY_RANK", "_AMT_RANK"}
                        With .Columns(Y & COLUMN_NAME)
                            .Group = G
                            .CellAppearance.BackColor = Drawing.Color.LightPink
                            .CellActivation = UltraWinGrid.Activation.NoEdit
                            .Format = "#,##0"
                            .Width = 60
                        End With
                    Next

                    .Columns(Y & "_QTY").Header.Caption = "Units"
                    .Columns(Y & "_QTY").Header.ToolTipText = "Units Sold " & Y
                    .Columns(Y & "_AMT").Header.Caption = "Retail"
                    .Columns(Y & "_AMT").Header.ToolTipText = "Retail $ Sold " & Y
                    .Columns(Y & "_QOH").Header.Caption = "OH Units"
                    .Columns(Y & "_QOH").Header.ToolTipText = "On Hand Units " & Y
                    .Columns(Y & "_AOH").Header.Caption = "OH Retail"
                    .Columns(Y & "_AOH").Header.ToolTipText = "On Hand Retail " & Y
                    .Columns(Y & "_QTY_RANK").Header.Caption = "#Rank"
                    .Columns(Y & "_QTY_RANK").Header.ToolTipText = "Rank based on Units, " & Y
                    .Columns(Y & "_AMT_RANK").Header.Caption = "$Rank"
                    .Columns(Y & "_AMT_RANK").Header.ToolTipText = "Rank based on $Retail"
                    .Columns(Y & "_SELL_THRU").Header.Caption = "Sell%"
                    .Columns(Y & "_SELL_THRU").Header.ToolTipText = "Sell Thru % " & Y
                    .Columns(Y & "_PCT_TOTAL").Header.Caption = "%Total"
                    .Columns(Y & "_PCT_TOTAL").Header.ToolTipText = "% of Total " & Y
                Next

                If b = 0 Then
                    COLUMN_NAME = "VAR"
                    With .Columns(COLUMN_NAME)
                        .Group = G
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                        .Format = "#,##0"
                        .Header.Caption = "Variance"
                        .Width = 70
                        .CellAppearance.BackColor = Color.LightCyan
                        Create_Summary(grd, COLUMN_NAME, , grd.DisplayLayout.Bands(b).Key)
                    End With

                    COLUMN_NAME = "VAR_PCT"
                    With .Columns(COLUMN_NAME)
                        .Group = G
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                        .Format = "#,##0.0"
                        .Header.Caption = "Var%"
                        .Width = 50
                        .CellAppearance.BackColor = Color.LightGoldenrodYellow
                        Create_Summary(grd, COLUMN_NAME, , grd.DisplayLayout.Bands(b).Key)
                    End With
                End If


            End With

            If grd.Name = "grdRSTSLSA2" Then
                grd.DisplayLayout.Override.AllowColSizing = UltraWinGrid.AllowColSizing.Synchronized
                If b = 1 Then
                    grd.DisplayLayout.Bands(b).SummaryFooterCaption = "Totals for [SCROLLTIPFIELD]"
                    With grdRSTSLSA2.DisplayLayout.Bands(1)
                        .Columns("CODE_VALUE").Header.Caption = "Item Code"
                        .Columns("DESC_VALUE").Header.Caption = "Description"
                        .Columns("CODE2_VALUE").Header.Caption = ""
                        .Columns("CODE2_VALUE").Hidden = True
                        .GroupHeadersVisible = False
                    End With
                End If

                b += 1
            End If
        Next


        With grdRSTSLSAP.DisplayLayout.Bands(0)
            For P As Integer = 0 To 24
                With .Columns("P" & Format(P, "00"))
                    .Width = 60
                    .Format = "#,##0"
                    Call Create_Summary(grdRSTSLSAP, "P" & Format(P, "00"), , , "#,##0")
                End With
            Next
            For Each COLUMN_NAME As String In New String() _
                {"CODE_VALUE", "DESC_VALUE", "CODE2_VALUE"}
                With grdRSTSLSAP.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                    .CellAppearance.BackColor = Drawing.Color.Beige
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                End With
            Next
            With grdRSTSLSAP.DisplayLayout.Bands(0)
                .Columns("CODE_VALUE").Width = 90
                .Columns("DESC_VALUE").Width = 110
                .Columns("CODE2_VALUE").Width = 70
                .Columns("CODE_VALUE").Header.Fixed = True
                .Columns("DESC_VALUE").Header.Fixed = True
                .Columns("CODE2_VALUE").Header.Fixed = True
            End With
        End With


        For Each grd As UltraWinGrid.UltraGrid In _
        New UltraWinGrid.UltraGrid() {grdRSTSLSB1, grdRSTSLSB2}
            With grd.DisplayLayout.Bands(0)
                For Each Y As String In New String() {"TY", "LY"}
                    For Each COLUMN_NAME As String In New String() _
                         {"_QTY", "_AMT", "_QOH", "_AOH"}
                        With .Columns(Y & COLUMN_NAME)
                            Call Create_Summary(grd, Y & COLUMN_NAME, , , "#,##0")
                            .CellActivation = UltraWinGrid.Activation.NoEdit
                            .Format = "#,##0"
                            If COLUMN_NAME Like "*AMT" Then
                                .Width = 90
                            Else
                                .Width = 70
                            End If
                            .Header.Appearance.BackColor = IIf(Y = "TY", Color.LightGreen, Color.Orange)
                            .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                        End With
                    Next

                    For Each COLUMN_NAME As String In New String() _
                        {"_SELL_THRU"}
                        With .Columns(Y & COLUMN_NAME)
                            .CellAppearance.BackColor = Drawing.Color.LightGray
                            .CellActivation = UltraWinGrid.Activation.NoEdit
                            .Format = "#,##0.0"
                            .Width = 60
                        End With
                    Next

                    For Each COLUMN_NAME As String In New String() _
                        {"_QTY_RANK", "_AMT_RANK"}
                        With .Columns(Y & COLUMN_NAME)
                            .CellAppearance.BackColor = Drawing.Color.LightPink
                            .CellActivation = UltraWinGrid.Activation.NoEdit
                            .Format = "#,##0"
                            .Width = 60
                        End With
                    Next
                Next

                If grd.Name = "grdRSTSLSB1" Then
                    .Columns("ITEM_CODE").Header.Fixed = True
                Else
                    .Columns("CUST_CODE").Header.Fixed = True
                End If
            End With

        Next

        With grdRSTSLSBR.DisplayLayout.Bands(0)
            .LevelCount = 2

            Dim G As UltraWinGrid.UltraGridGroup

            G = .Groups.Add("CODES", "Codes")
            G.Header.Appearance.BackColor = Color.CornflowerBlue
            G.Header.Appearance.BackGradientStyle = GradientStyle.None
            G.Header.Appearance.ForeColor = Color.White
            G.Header.Fixed = True
            For Each COLUMN_NAME As String In New String() _
                {"ITEM_CODE", "ITEM_DESC", "ITEM_CATGY_CODE", "DEPT_CODE", "STYLE_CODE",
                 "STYLE_CODE_NC", "COLOR_CODE", "SIZE_CODE", "COLLECTION_CODE", "HC_CODE", "PROD_CODE", "ITEM_CLASS_CODE",
                 "MATL_CODE", "PRICE_POINT_CODE", "ITEM_RETAIL_PRICE", "ITEM_PRICE", "ITEM_ABC_CODE"}
                With .Columns(COLUMN_NAME)
                    .Group = G
                    .CellAppearance.BackColor = Drawing.Color.Beige
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                End With
            Next

            G = .Groups.Add("DATA", "")
            G.Header.Appearance.BackColor = Color.Orange
            G.Header.Appearance.BackColor2 = Color.White
            G.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            With .Columns.Add("SLS")
                .Group = G
                .Level = 0
                .Header.Caption = ""
            End With
            With .Columns.Add("EOW")
                .Group = G
                .Level = 1
                .Header.Caption = ""
            End With
            G.Width = 50

            For W As Integer = 1 To 60
                'Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", WKS60(W))
                G = .Groups.Add("W" & Format(W, "00"), "") ' rowGLTPARM3.Item("LEGEND"))
                G.Header.Appearance.BackColor = Color.Orange
                G.Header.Appearance.BackColor2 = Color.White
                G.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                With .Columns("SLS" & Format(W, "00"))
                    .Group = G
                    .Level = 0
                    .Header.Caption = "Qty"
                    .Format = "###,##0"
                End With
                With .Columns("EOW" & Format(W, "00"))
                    .Group = G
                    .Level = 1
                    .Header.Caption = "O/H"
                    .Format = "###,##0"
                End With
                G.Width = 70
            Next

            .Override.FilterOperatorLocation = UltraWinGrid.FilterOperatorLocation.WithOperand
        End With

    End Sub

    Private Sub chkShowRank_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowRank.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_Selected()
    End Sub

    Private Sub chkShowPcts_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowPcts.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_Selected()
    End Sub

    Sub Show_Selected()
        Dim B As Int32 = 0
        For Each grd As UltraWinGrid.UltraGrid In _
        New UltraWinGrid.UltraGrid() {grdRSTSLSA1, grdRSTSLSA2, grdRSTSLSA2}
            With grd.DisplayLayout.Bands(B)
                For Each Y As String In New String() {"TY", "LY"}
                    .Columns(Y & "_SELL_THRU").Hidden = Not chkShowPcts.Checked Or (optSales.Value = "W")
                    .Columns(Y & "_PCT_TOTAL").Hidden = Not chkShowPcts.Checked
                    .Columns(Y & "_QTY_RANK").Hidden = Not chkShowRank.Checked
                    .Columns(Y & "_AMT_RANK").Hidden = Not chkShowRank.Checked
                    .Columns(Y & "_QOH").Hidden = Not chkShowOH.Checked
                    .Columns(Y & "_AOH").Hidden = Not chkShowOH.Checked
                Next

                .Groups("LY").Hidden = Not chkShowLY.Checked
            End With

            If grd.Name = "grdRSTSLSA2" Then
                B += 1
            End If
        Next

        For Each grd As UltraWinGrid.UltraGrid In _
        New UltraWinGrid.UltraGrid() {grdRSTSLSB1, grdRSTSLSB2}
            With grd.DisplayLayout.Bands(0)
                .Columns("LY_QTY").Hidden = Not chkShowLY.Checked
                .Columns("LY_AMT").Hidden = Not chkShowLY.Checked
                .Columns("LY_QOH").Hidden = Not chkShowLY.Checked Or (optSales.Value = "W")
                .Columns("LY_AOH").Hidden = Not chkShowLY.Checked Or (optSales.Value = "W")
                .Columns("LY_SELL_THRU").Hidden = Not chkShowLY.Checked Or (optSales.Value = "W")
                .Columns("LY_QTY_RANK").Hidden = Not chkShowLY.Checked
                .Columns("LY_AMT_RANK").Hidden = Not chkShowLY.Checked
            End With
        Next
    End Sub

    Sub Calculate_Rank(ByVal tbl As DataTable)
        Dim RANK As Integer

        For Each Y As String In New String() {"TY", "LY"}
            Dim QTY_TOTAL As Int32 = 0
            For Each DT As String In New String() {"QTY", "AMT"}
                RANK = 0
                For Each row As DataRow In tbl.Select("", Y & "_" & DT & " DESC")
                    RANK += 1
                    row.Item(Y & "_" & DT & "_RANK") = RANK
                    If DT = "QTY" Then
                        QTY_TOTAL += Val(row.Item(Y & "_" & DT) & "")
                    End If
                Next
            Next
            If tbl.Columns.Contains(Y & "_PCT_TOTAL") Then
                tbl.Columns(Y & "_PCT_TOTAL").Expression = "IIF(" & CStr(QTY_TOTAL) & "=0,0,100 * " & Y & "_QTY/" & CStr(QTY_TOTAL) & ")"
            End If
        Next


    End Sub

    Sub CreateGraph_Totals()

        Dim chtIsVisible As Boolean = chtTotals.Visible
        chtTotals.Visible = False

        chtTotals.DataSource = Nothing

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Charting Data")

        Me.SuspendLayout()

        Dim RL() As String

        chtTotals.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtTotals.LabelHash = labelHash

        chtTotals.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtTotals.Tooltips.FormatString = "<HIGHLOW>"

        chtTotals.TitleTop.Text = "Totals " & optTD.Text & " " & optTrend.Text & ", by " & SUMCOLs(grdRSTSLSA1.Tag)

        Dim RLi As Integer = 0

        Dim DTY As New DataTable
        With DTY
            .Columns.Add("CODE")
            .Columns.Add("VALUE", GetType(System.Decimal))
        End With

        Dim DTX As DataTable = dst.Tables("RSTSLSA1")
        ReDim RL(DTX.Rows.Count - 1)
        For Each row As DataRow In DTX.Select("", "CODE_VALUE")
            RL(RLi) = row("CODE_VALUE") & "" ' & ":" & row("DESC_VALUE")
            RLi += 1
            DTY.Rows.Add(New Object() {row.Item("CODE_VALUE"), row.Item("TY_QTY")})
        Next

        chtTotals.DataSource = DTY
        chtTotals.PieChart.ColumnIndex = -1
        chtTotals.PieChart.OthersCategoryPercent = 2
        chtTotals.DataBind()

        chtTotals.Visible = True

        Me.ResumeLayout()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()
    End Sub

    Sub CreateGraph_Trend()

        Dim chtIsVisible As Boolean = chtTrend.Visible
        chtTrend.Visible = False

        Dim periods As Integer
        If optChart.Value = "1" Then
            periods = 24
        Else
            periods = 12
        End If

        chtTrend.DataSource = Nothing

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Charting Data")

        Me.SuspendLayout()

        Dim RL() As String
        Dim CL() As String
        ReDim CL(Periods)

        For i As Integer = 0 To periods
            CL(i) = YPP(periods - i, 1)
        Next

        chtTrend.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtTrend.LabelHash = labelHash

        chtTrend.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtTrend.Tooltips.FormatString = "<HIGHLOW>"

        Dim DT As New DataTable
        DT.Columns.Add("CODE_VALUE")
        DT.Columns.Add("DESC_VALUE")
        For P As Integer = periods To 0 Step -1
            DT.Columns.Add("P" & Format(P, "00"), GetType(System.Decimal))
        Next

        Dim RLi As Integer = 0
        Dim DTX As DataTable = dst.Tables("RSTSLSAP")

        If optChart.Value = "1" Then
            ReDim RL(DTX.Rows.Count - 1)
            chtTrend.TitleTop.Text = "Trend " & optTD.Text & " " & optTrend.Text & ", by " & SUMCOLs(grdRSTSLSA1.Tag)

            For Each row As DataRow In DTX.Select("", "CODE_VALUE")
                RL(RLi) = row("CODE_VALUE") & "" ' & ":" & row("DESC_VALUE")
                RLi += 1

                Dim rowDT As DataRow = DT.NewRow
                rowDT.Item("CODE_VALUE") = row("CODE_VALUE")
                rowDT.Item("DESC_VALUE") = row("DESC_VALUE")
                For P As Integer = 1 To periods
                    rowDT.Item("P" & Format(P, "00")) = row("P" & Format(P, "00"))
                Next
                DT.Rows.Add(rowDT)
            Next
        Else

            Dim grd As UltraWinGrid.UltraGrid = grdRSTSLSA1
            'If "" = "" Then
            '    grd = grdRSTSLSA1
            'Else
            '    grd = grdRSTSLSA2 - need to prepare RSTSLSAP before trying to plot RSTSLSA2
            'End If

            If grd.ActiveRow Is Nothing Then
                Exit Sub
            Else
                ReDim RL(1)
                Dim CODE_VALUE As String = grd.ActiveRow.Cells("CODE_VALUE").Text
                Dim DESC_VALUE As String = grd.ActiveRow.Cells("DESC_VALUE").Text

                chtTrend.TitleTop.Text = SUMCOLs(grd.Tag) & " " & CODE_VALUE & ", " & optTD.Text & " " & optTrend.Text & ", TY vs LY"

                Dim row As DataRow = DTX.Rows.Find(CODE_VALUE)
                For Each Y As String In New String() {"TY", "LY"}
                    RL(RLi) = Y
                    RLi += 1

                    Dim rowDT As DataRow = DT.NewRow
                    rowDT.Item("CODE_VALUE") = Y
                    rowDT.Item("DESC_VALUE") = DESC_VALUE

                    Dim PX As Int32 = 0
                    If Y = "LY" Then
                        PX = 12
                    End If

                    If row IsNot Nothing Then
                        For P As Integer = 0 To periods
                            rowDT.Item("P" & Format(P, "00")) = row("P" & Format(P + PX, "00"))
                        Next
                    End If
                    DT.Rows.Add(rowDT)
                Next
            End If
        End If
        chtTrend.Data.SetRowLabels(RL)
        chtTrend.Data.SetColumnLabels(CL)

        chtTrend.DataSource = DT
        chtTrend.DataBind()
        chtTrend.Visible = True

        Me.ResumeLayout()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()
    End Sub


    Public Class MyCustomTooltip
        Implements IRenderLabel

        Public Sub New()

        End Sub 'New

        Public Overloads Function ToString(ByVal Context As System.Collections.Hashtable) As String Implements IRenderLabel.ToString
            'Return Context("ITEM_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))
            'Return Context("SERIES_LABEL") & vbCrLf & Context("ITEM_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))
            Return Context("SERIES_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))

        End Function 'ToString 
    End Class 'MyCustomTooltip

    Private Sub optTrend_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTrend.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_Charts()
    End Sub

    Sub Set_Charts()
        Get_sqlRSTSLSAP(True)

        CreateGraph_Trend()
        CreateGraph_Totals()
    End Sub

    Private Sub optChart_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optChart.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_Charts()
    End Sub

    Private Sub grdRSTSLSB1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdRSTSLSB1.AfterRowActivate
        If Not splRSTSLSB1.Panel2Collapsed Then
            Load_RSTSLSB2()
        End If
    End Sub

    Private Sub grdRSTSLSB1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdRSTSLSB1.DoubleClickRow
        splRSTSLSB1.Panel2Collapsed = False
        Load_RSTSLSB2()
    End Sub

    Sub Load_RSTSLSB2()
        If grdRSTSLSB1.ActiveRow.IsDataRow Then
            Dim ITEM_CODE As String = grdRSTSLSB1.ActiveRow.Cells("ITEM_CODE").Text
            Fill_Records("RSTSLSB2", , , Replace(sqlRSTSLSB2, ":PARM1", "'" & ITEM_CODE & "'"))
            Dim ITEM_RETAIL_PRICE As Decimal = Val(grdRSTSLSB1.ActiveRow.Cells("ITEM_RETAIL_PRICE").Value)
            For Each rowRSTSLSB2 As DataRow In dst.Tables("RSTSLSB2").Rows
                rowRSTSLSB2.Item("TY_AOH") = Val(rowRSTSLSB2.Item("TY_QOH") & "") * ITEM_RETAIL_PRICE
                rowRSTSLSB2.Item("LY_AOH") = Val(rowRSTSLSB2.Item("LY_QOH") & "") * ITEM_RETAIL_PRICE
            Next

            Calculate_Rank(dst.Tables("RSTSLSB2"))

            grdRSTSLSB2.Text = "Item " & ITEM_CODE & " Breakout by Customer"
        Else
            splRSTSLSB1.Panel2Collapsed = True
            dst.Tables("RSTSLSB2").Rows.Clear()
        End If

    End Sub

    Private Sub grdRSTSLSB1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdRSTSLSB1.InitializeRow
        Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Value
        'If My.Computer.FileSystem.FileExists("C:\Documents and Settings\wjz\Desktop\JHI\Images\" & ITEM_CODE & ".JPG") Then
        '    e.Row.Cells("IMAGE").Appearance.ImageHAlign = HAlign.Center
        '    e.Row.Cells("IMAGE").Appearance.Image = ASCMAIN1.Get_Image("C:\Documents and Settings\wjz\Desktop\JHI\Images\", ITEM_CODE & ".JPG")
        'End If
    End Sub


    Sub Excel_Export_Summary(ByVal myWorkbook As Infragistics.Documents.Excel.Workbook)

        Call Export_to_Excel_Add_grd(myWorkbook, grdRSTSLSA1, False, grdRSTSLSA1.Text)

        Dim cx As Int32 = 0
        For Each gcol As UltraWinGrid.UltraGridColumn In grdRSTSLSA1.DisplayLayout.Bands(0).Columns
            If Not gcol.Hidden And Not gcol.Group.Hidden Then
                cx += 1
            End If
        Next

        Dim c As Int32 = 0
        For Each grow As UltraWinGrid.UltraGridRow In grdRSTSLSA1.Rows
            grdRSTSLSA1.ActiveRow = grow
            c += cx + 1
            Export_to_Excel_Add_grd_to_Sheet(grdRSTSLSA2, 4, c)
        Next
    End Sub

    Sub Excel_Export_TopN(ByVal myWorkbook As Infragistics.Documents.Excel.Workbook)

        'Sort_grdColumns(grdRSTSLSB1, "ITEM_CODE")

        Calculate_N()
        Call Export_to_Excel_Add_grd(myWorkbook, grdRSTSLSB1, False, grdRSTSLSB1.Text)

        Dim cx As Int32 = 0
        For Each gcol As UltraWinGrid.UltraGridColumn In grdRSTSLSB1.DisplayLayout.Bands(0).Columns
            If Not gcol.Hidden Then
                cx += 1
            End If
        Next

        Dim tblN As DataTable = dst.Tables("RSTSLSB1").Copy

        Dim HCs As New List(Of String)

        Dim gh As Int32 = 0
        Dim ch As Int32 = 0
        With grdRSTSLSB1.DisplayLayout.Bands(0)

            For Each COLUMN_NAME As String In grdRSTSLSB1_HCs
                If Not .Columns(COLUMN_NAME).Hidden And COLUMN_NAME <> "ITEM_CODE" And COLUMN_NAME <> "N" Then
                    .Columns(COLUMN_NAME).Hidden = True
                    HCs.Add(COLUMN_NAME)
                    ch += 1
                End If
            Next

            If .SortedColumns.Count <> 0 Then
                For Each GC As UltraWinGrid.UltraGridColumn In .SortedColumns
                    If GC.IsGroupByColumn Then
                        gh += 1
                    End If
                Next
            End If
        End With

        Dim SQL As String = ""

        myWorkSheet = myWorkbook.Worksheets(0)

        Dim Original_Text As String = grdRSTSLSB1.Text
        Dim c As Int32 = cx + 1 + gh

        Dim tblRSTSLSB1 As DataTable = dst.Tables("RSTSLSB1").Clone
        tblRSTSLSB1.Merge(dst.Tables("RSTSLSB1"))
        For Each rowRSTSLSB1 As DataRow In tblRSTSLSB1.Rows
            rowRSTSLSB1.Item("TY_QTY") = DBNull.Value
            rowRSTSLSB1.Item("TY_AMT") = DBNull.Value
            rowRSTSLSB1.Item("TY_QOH") = DBNull.Value
            rowRSTSLSB1.Item("LY_QTY") = DBNull.Value
            rowRSTSLSB1.Item("LY_AMT") = DBNull.Value
            rowRSTSLSB1.Item("LY_QOH") = DBNull.Value
            'rowRSTSLSB1.Item("TY_SELL_THRU") = DBNull.Value
            'rowRSTSLSB1.Item("LY_SELL_THRU") = DBNull.Value
            rowRSTSLSB1.Item("TY_QTY_RANK") = DBNull.Value
            rowRSTSLSB1.Item("TY_AMT_RANK") = DBNull.Value
            rowRSTSLSB1.Item("LY_QTY_RANK") = DBNull.Value
            rowRSTSLSB1.Item("LY_AMT_RANK") = DBNull.Value
            rowRSTSLSB1.Item("TY_AOH") = DBNull.Value
            rowRSTSLSB1.Item("LY_AOH") = DBNull.Value
            rowRSTSLSB1.Item("N") = DBNull.Value
        Next


        ASCMAIN1.sql = "Select Distinct CUST_CODE from " & RSTSLSAX & " RSTSLSAX " & Replace(ASCMAIN1.SQL_Add_WHERE(sql_filters("CUST_CODE")), "X.", "RSTSLSAX.")
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows

            Dim CUST_CODE As String = row.Item(0)
            ASCMAIN1.Progress("-", CUST_CODE)

            ASCDATA1.ExecuteSQL("Truncate Table " & RSTSLSB0)
            Dim sql2 As String = Replace(Replace(sqlRSTSLSB0, "Sum (", "Sum (CASE WHEN RSTSLSA0.CUST_CODE = '" & CUST_CODE & "' THEN "), ") ", " ELSE 0 END)")

            ASCDATA1.ExecuteSQL("Insert into " & RSTSLSB0 & " " & sql2)

            SQL = Replace(sqlRSTSLSB1, " group by ", sql_filter_no_CUST_CODE & " group by ")
            splRSTSLSB1.Panel2Collapsed = True
            dst.Tables("RSTSLSB2").Rows.Clear()
            Fill_Records("RSTSLSB1", "", True, SQL)

            dst.Tables("RSTSLSB1").Merge(tblRSTSLSB1, True)

            'Sort_grdColumns(grdRSTSLSB1, "ITEM_CODE")

            If optTopN_ATB.Value <> "A" Then
                For Each rowRSTSLSB1 As DataRow In dst.Tables("RSTSLSB1").Rows
                    Dim rowN As DataRow = tblN.Rows.Find(rowRSTSLSB1.Item("ITEM_CODE"))
                    rowRSTSLSB1.Item("N") = rowN.Item("N")
                Next
            Else
                Calculate_N()
            End If

            grdRSTSLSB1.Text = "Customer " & CUST_CODE

            Try
                Export_to_Excel_Add_grd_to_Sheet(grdRSTSLSB1, 4, c)
            Catch ex As Exception
                MsgBox(ex.Message)
                Exit Sub
            End Try


            c += cx + 1 - ch + gh

        Next

        With grdRSTSLSB1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In HCs
                .Columns(COLUMN_NAME).Hidden = False
            Next
        End With


        grdRSTSLSB1.Text = Original_Text

        ASCDATA1.ExecuteSQL("Truncate Table " & RSTSLSB0)
        ASCDATA1.ExecuteSQL("Insert into " & RSTSLSB0 & " " & sqlRSTSLSB0)

        SQL = Replace(sqlRSTSLSB1, " group by ", sql_filter_no_CUST_CODE & " group by ")
        splRSTSLSB1.Panel2Collapsed = True
        dst.Tables("RSTSLSB2").Rows.Clear()
        Fill_Records("RSTSLSB1", "", True, SQL)

        Calculate_N()

    End Sub

    'Sub Load_Filter(ByVal COLUMN_NAMEs() As String)

    '    grdASTFLTR1.DataSource = dst.Tables("ASTFLTR1")

    '    For Each COLUMN_NAME As String In COLUMN_NAMEs
    '        Dim rowASTDSQLK As DataRow = LookUp("ASTDSQLK", COLUMN_NAME)
    '        Dim COLUMN_CAPTION As String = ""

    '        If rowASTDSQLK Is Nothing Then
    '            COLUMN_CAPTION = ASCMAIN1.Make_Caption(COLUMN_NAME)
    '        Else
    '            COLUMN_CAPTION = rowASTDSQLK.Item("COLUMN_CAPTION") & ""
    '        End If
    '        dst.Tables("ASTFLTR1").Rows.Add(New String() {COLUMN_NAME, COLUMN_CAPTION})
    '    Next
    'End Sub

    Private Sub grdASTFLTR1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdASTFLTR1.BeforeRowUpdate

        Dim COLUMN_NAME As String = e.Row.Cells("COLUMN_NAME").Text
        Dim sql As String = ASCMAIN1.CodeSelector.Get_SQL(COLUMN_NAME)
        If sql <> "" Then
            Dim CODE_VALUES_new As String = ""
            Dim CODE_VALUES As String = e.Row.Cells("CODE_VALUES").Text
            Dim KEY_EXPRESSION As String = ASCMAIN1.CodeSelector.grdColumns(0).Item("COLUMN_NAME")

            If CODE_VALUES <> "" Then
                Dim CODE_VALUES_old As String = ""
                For Each txt As String In Split(Replace(CODE_VALUES, "'", ""), ",")
                    CODE_VALUES_old = CODE_VALUES_old & ",'" & ASCMAIN1.Format_Field(txt, COLUMN_NAME, , True) & "'"
                Next
                CODE_VALUES_old = Mid$(CODE_VALUES_old, 2)
                Dim where_or_and As String = " where "
                If ASCMAIN1.CodeSelector.tblASTVIEW1.Rows(0).Item("WHERE_CLAUSE") & "" <> "" Then
                    where_or_and = " and "
                End If

                For Each dr As DataRow In ASCDATA1.GetDataTable(sql & where_or_and & KEY_EXPRESSION & " IN (" & CODE_VALUES_old & ")").Rows
                    CODE_VALUES_new = CODE_VALUES_new & "," & dr.Item(0)
                Next
            End If

            CODE_VALUES_new = Mid(CODE_VALUES_new, 2)
            If CODE_VALUES_new <> CODE_VALUES Then
                grdASTFLTR1.DisplayLayout.Rows(e.Row.Index).Cells("CODE_VALUES").Value = CODE_VALUES_new
            End If
        End If
    End Sub

    Private Sub grdASTFLTR1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTFLTR1.ClickCellButton
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL(grdASTFLTR1.ActiveRow.Cells("COLUMN_NAME").Text)
        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = Replace(grdASTFLTR1.ActiveRow.Cells("CODE_VALUES").Text & "", ",", Chr(0))
            Dim F As New ASFCODE1
            F.ShowDialog()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                grdASTFLTR1.ActiveRow.Cells("CODE_VALUES").Value = Mid$(Replace(ASCMAIN1.CodeSelector.SelectedCodes0, Chr(0), ","), 2)
                grdASTFLTR1.UpdateData()
            End If
        End If
    End Sub

    Private Sub cmdFilterApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFilterApply.Click
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("Now Applying Filters")

        Get_sqlRSTSLSAX("1")
        grdRSTSLSB1.Text = filter_desc
        splRSTSLSB1.Panel2Collapsed = True
        dst.Tables("RSTSLSB2").Rows.Clear()
        Fill_Records("RSTSLSB1", , , Replace(sqlRSTSLSB1, " group by ", sql_filter_no_CUST_CODE & " group by "))
        Calculate_Rank(dst.Tables("RSTSLSB1"))

        ReSummarize()

        'EnforceConstraints(False)
        'Setup_RSTSLSA0(RYP, False)
        'Setup_RSTSLSAX(True)
        'EnforceConstraints(True)

        If chkShow60.Checked Then
            ASCMAIN1.Progress("Now Loading 60 Weeks of Sales & On Hand with Filters")
            Fill_Records("RSTSLSBR", , , Replace(sqlRSTSLSBR, " group by ", sql_filter_no_CUST_CODE & Replace(filter_CUST_CODE, "X.", "RSTRETL1.") & " group by "))
            grdRSTSLSBR.Text = "Sales & On Hand by Item, by Week for " & filter_desc
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub optTopNTB_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTopN_ATB.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        'Calculate_N()
        ReSummarize()
    End Sub

    Private Sub optTopN_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTopN.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        'Calculate_N()
        ReSummarize()
    End Sub

    Sub Calculate_N()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Calculating N")

        For Each rowRSTSLSB1 As DataRow In dst.Tables("RSTSLSB1").Rows
            rowRSTSLSB1.Item("N") = 0
        Next

        Dim COLs As New List(Of String)
        With grdRSTSLSB1.DisplayLayout.Bands(0)
            If .SortedColumns.Count <> 0 Then
                Dim GC As UltraWinGrid.UltraGridColumn
                For i As Int32 = .SortedColumns.Count - 1 To 0 Step -1
                    GC = .SortedColumns(i)
                    If GC.IsGroupByColumn Then
                        COLs.Add(GC.Key)
                    Else
                        .SortedColumns.Remove(GC)
                    End If
                Next
            End If
        End With

        If COLs.Count = 0 Then
            Calculate_N_for_a_Group("")
        Else
            For Each row As DataRow In ASCDATA1.SelectDistinct("RSTSLSB1", COLs.ToArray).Rows
                Dim SQL As String = ""
                For Each COL As String In COLs
                    SQL &= " and " & COL & " = '" & row.Item(COL) & "'"
                Next
                Calculate_N_for_a_Group(Mid(SQL, 6))
            Next
        End If

        Dim dvw As DataView = DirectCast(grdRSTSLSB1.DataSource, DataTable).DefaultView
        If optTopN_ATB.Value = "A" Then
            dvw.RowFilter = ""
            If COLs.Count = 0 Then
                Sort_grdColumns(grdRSTSLSB1, "ITEM_CODE")
            End If
        Else
            dvw.RowFilter = "N <= " & CStr(numN.Value)

            grdRSTSLSB1.DisplayLayout.Bands(0).SortedColumns.Add("N", False, False)
            If grdRSTSLSB1.Rows.Count <> 0 Then grdRSTSLSB1.ActiveRow = grdRSTSLSB1.Rows(0)

            'Sort_grdColumns(grdRSTSLSB1, "N")
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("")

    End Sub

    Sub Calculate_N_for_a_Group(ByVal sql As String)
        Dim N As Int32 = 0
        Dim SORT As String = "TY_" & optTopN.Value
        If optTopN_ATB.Value <> "B" Then
            SORT &= " DESC"
        End If
        For Each rowRSTSLSB1 As DataRow In dst.Tables("RSTSLSB1") _
        .Select(sql, SORT)
            N += 1
            rowRSTSLSB1.Item("N") = N
        Next
    End Sub
    Private Sub numN_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles numN.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            'Calculate_N()
            ReSummarize()
        End If
    End Sub

    Private Sub numN_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numN.ValueChanged
    End Sub

    Private Sub optItem_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optItem.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        ReSummarize()
    End Sub

    Sub ReSummarize()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Summarizing", "")

        Setup_RSTSLSAX()
        CreateGraph_Totals()

        Calculate_N()
        grdRSTSLSB1.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")

        With grdRSTSLSB1.DisplayLayout.Bands(0).Columns("ITEM_CODE")
            Select Case optItem.Value
                Case "ITEM_CODE"
                    .Header.Caption = "Item"
                Case "STYLE_CODE"
                    .Header.Caption = "Style"
                Case "STYLE_CODE_NC"
                    .Header.Caption = "Style NC"
            End Select
        End With
        With grdRSTSLSB1.DisplayLayout.Bands(0).Columns("STYLE_CODE")
            Select Case optItem.Value
                Case "STYLE_CODE"
                    .Header.Caption = "Item"
                Case Else
                    .Header.Caption = "Style"
            End Select
        End With
        With grdRSTSLSB1.DisplayLayout.Bands(0).Columns("STYLE_CODE_NC")
            Select Case optItem.Value
                Case "STYLE_CODE_NC"
                    .Header.Caption = "Item"
                Case Else
                    .Header.Caption = "Style NC"
            End Select
        End With
    End Sub

    Private Sub cmdTopN_Apply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTopN_Apply.Click
        ReSummarize()
    End Sub

    Private Sub cmdFilterClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFilterClear.Click
        For Each rowASTFLTR1 As DataRow In dst.Tables("ASTFLTR1").Rows
            rowASTFLTR1.Item("CODE_VALUES") = ""
        Next
    End Sub

    Private Sub grdRSTSLSA1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdRSTSLSA1.DoubleClickRow
        If grdRSTSLSA1.Tag <> "CUST_CODE" Then

            Dim CODE_VALUE As String = e.Row.Cells("CODE_VALUE").Value & ""
            If grdRSTSLSA1.Tag = "ITEM_CATGY_CODE" Then
                CODE_VALUE = e.Row.Cells("DESC_VALUE").Value & ""
            End If

            optTopN_ATB.Value = "A"
            With grdRSTSLSB1.DisplayLayout
                .ClearGroupByColumns()
                .Bands(0).ColumnFilters.ClearAllFilters()
                .Rows.ColumnFilters(grdRSTSLSA1.Tag).FilterConditions.Add _
                (Infragistics.Win.UltraWinGrid.FilterComparisionOperator.Equals, CODE_VALUE)
                Show_Filter(grdRSTSLSB1, True)
                .Bands(0).Columns(grdRSTSLSA1.Tag).Hidden = False

                If grdRSTSLSB1.Rows.FilteredInRowCount <> 0 Then
                    grdRSTSLSB1.ActiveRow = grdRSTSLSB1.Rows.GetFilteredInNonGroupByRows()(0)
                End If
            End With
            tabMain.Tabs("Top N").Selected = True
        End If
    End Sub

    Sub Set_Filter()

        filter_count = 0
        filter_desc = ""
        sql_filter = ""
        sql_filter_no_CUST_CODE = ""
        sql_filters.Clear()
        Dim sql As String = ""
        For Each rowASTFLTR1 As DataRow In dst.Tables("ASTFLTR1").Rows
            Dim CODE_VALUES As String = rowASTFLTR1.Item("CODE_VALUES") & ""
            Dim filter_COL As String = rowASTFLTR1.Item("COLUMN_NAME") & ""
            Dim filter_CAP As String = rowASTFLTR1.Item("COLUMN_CAPTION") & ""
            sql = ""
            If CODE_VALUES <> "" Then
                filter_count += 1
                If InStr(CODE_VALUES, ",") = 0 Then
                    sql = " and X." & filter_COL & " = '" & CODE_VALUES & "'"
                    filter_desc &= "," & filter_CAP & " " & CODE_VALUES
                    'filter_desc = filter_CAP & " " & CODE_VALUES
                Else
                    sql = " and X." & filter_COL & " in ('" & Replace(CODE_VALUES, ",", "','") & "')"
                    filter_desc &= "," & filter_CAP & " " & CODE_VALUES
                    'filter_desc = "Selected " & filter_CAP & "s"
                End If
            End If
            If filter_COL <> "CUST_CODE" Then
                sql_filter_no_CUST_CODE &= sql
            End If
            sql_filter &= sql
            sql_filters.Add(filter_COL, sql)
        Next
        sql_filter_no_CUST_CODE = Replace(sql_filter_no_CUST_CODE, "X.", "ICTITEM1.")
        If filter_count = 0 Then
            filter_desc = "All Records"
        Else
            filter_desc = Mid(filter_desc, 2)
            If filter_count > 1 Then
                'filter_desc = "Filtered Records"
            End If
        End If



        If sql_filters.ContainsKey("CUST_CODE") Then
            If filter_CUST_CODE <> sql_filters("CUST_CODE") Then
                filter_CUST_CODE = sql_filters("CUST_CODE")
                EnforceConstraints(False)
                Setup_RSTSLSA0(RYP, False)
                Setup_RSTSLSAX(True)
                EnforceConstraints(True)
            End If
        End If
    End Sub

    Private Sub grdRSTSLSA2_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdRSTSLSA2.DoubleClickRow
        grdRSTSLSA2.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.MultiBand
        Dim sql As String = Replace(sqlRSTSLSA3, "X.X", grdRSTSLSA2.Tag)
        sql = Replace(sql, " group by ", " and " & grdRSTSLSA1.Tag & " = '" & grdRSTSLSA1.ActiveRow.Cells("CODE_VALUE").Text & "' group by ")
        sql = Replace(sql, " group by ", " and " & grdRSTSLSA2.Tag & " = '" & grdRSTSLSA2.ActiveRow.Cells("CODE_VALUE").Text & "' group by ")
        'Stop

        'grdRSTSLSA2.DisplayLayout.Bands(1).Groups(0).Header.Caption = grdRSTSLSA2.ActiveRow.Cells("CODE_VALUE").Text & ", By Item"
        grdRSTSLSA2.DisplayLayout.Bands(1).Groups(0).Header.Caption = "By Item"
        'grdRSTSLSA2.DisplayLayout .Bands(1).
        Fill_Records("RSTSLSA3", , False, sql)
        grdRSTSLSA2.ActiveRow.Expanded = True
        With grdRSTSLSA2.DisplayLayout.Bands(1)
            .Columns("DESC_VALUE").Width = grdRSTSLSA2.DisplayLayout.Bands(0).Groups(0).Width - (.Columns("CODE_VALUE").Width + .RowSelectorWidthResolved)
        End With
    End Sub

    Private Sub grdRSTSLSA2_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdRSTSLSA2.InitializeLayout

    End Sub

    Private Sub grdRSTSLSA1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdRSTSLSA1.InitializeLayout

    End Sub

    Private Sub grdRSTSLSB1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdRSTSLSB1.InitializeLayout

    End Sub

    Private Sub grdASTFLTR1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdASTFLTR1.InitializeLayout

    End Sub

    Private Sub grdRSTSLSBR_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdRSTSLSBR.InitializeLayout

    End Sub

    Private Sub grdRSTSLSBR_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdRSTSLSBR.InitializeRow
        e.Row.Cells("SLS").Value = "Sold"
        e.Row.Cells("EOW").Value = "EOW"
    End Sub

    Private Sub chkShowOH_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowOH.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_Selected()
    End Sub
End Class