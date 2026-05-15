Imports Infragistics.UltraChart.Resources
Imports Infragistics.UltraChart.Shared.Styles

Public Class SAFSLSI1

    Dim RYP As String
    Dim SATSLSI1 As String
    Dim CUST_CODEs() As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst

            Create_SATSLSI1("")
            ASCMAIN1.sql = "Select * from " & SATSLSI1
            Create_TDA(.Tables.Add, "SATSLSI1", "**", 0, False, "", 0)

            .Tables("SATSLSI1").Columns("QTY_RTL_LY_MTD").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns("QTY_RTL_TY_MTD").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns("QTY_RTL_LS").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns("QTY_RTL_TS").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns("AMT_RTL_LS").DataType = GetType(System.Decimal)
            .Tables("SATSLSI1").Columns("AMT_RTL_TS").DataType = GetType(System.Decimal)
            .Tables("SATSLSI1").Columns("QTY_EOW").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns.Add("QTY_RTL", GetType(System.Int64), "QTY_RTL_LS + QTY_RTL_TS")
            .Tables("SATSLSI1").Columns.Add("AMT_RTL", GetType(System.Decimal), "AMT_RTL_LS + AMT_RTL_TS")
            .Tables("SATSLSI1").Columns.Add("AMT_RTL_WS", GetType(System.Decimal), "QTY_RTL * ITEM_PRICE")
            .Tables("SATSLSI1").Columns.Add("AMT_EOW", GetType(System.Decimal), "QTY_EOW * ITEM_RETAIL_PRICE")
            .Tables("SATSLSI1").Columns.Add("AMT_EOW_WS", GetType(System.Decimal), "QTY_EOW * ITEM_PRICE")

            .Tables("SATSLSI1").Columns("QTY_EOW_C0").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns("QTY_EOW_C1").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns("QTY_EOW_C2").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns("QTY_EOW_C3").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns("QTY_EOW_C4").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns("QTY_EOW_C5").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns("QTY_EOW_C6").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns("QTY_EOW_C7").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns("QTY_EOW_C8").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns("QTY_EOW_C9").DataType = GetType(System.Int64)

            .Tables("SATSLSI1").Columns("QTY_SHP_LY_MTD").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns("QTY_SHP_TY_MTD").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns("QTY_SHP_LS").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns("QTY_SHP_TS").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns("AMT_SHP_LS").DataType = GetType(System.Decimal)
            .Tables("SATSLSI1").Columns("AMT_SHP_TS").DataType = GetType(System.Decimal)
            .Tables("SATSLSI1").Columns.Add("QTY_SHP", GetType(System.Int64), "QTY_SHP_LS + QTY_SHP_TS")
            .Tables("SATSLSI1").Columns.Add("AMT_SHP", GetType(System.Decimal), "AMT_SHP_LS + AMT_SHP_TS")
            .Tables("SATSLSI1").Columns.Add("AMT_SHP_WS", GetType(System.Int64), "QTY_SHP * ITEM_PRICE")

            .Tables("SATSLSI1").Columns("QOH_SHP").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns("QOH_RTN").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns("QOH_OTH").DataType = GetType(System.Int64)
            .Tables("SATSLSI1").Columns.Add("QOH_ALL", GetType(System.Int64), "QOH_SHP + QOH_RTN + QOH_OTH")
            .Tables("SATSLSI1").Columns.Add("WOH_ALL", GetType(System.Decimal), "QOH_ALL * ITEM_PRICE")

            .Tables("SATSLSI1").Columns.Add("SELL_THRU_PCT", GetType(System.Decimal), "IIF(QTY_RTL + QTY_EOW = 0, 0 , 100 * QTY_RTL / (QTY_RTL + QTY_EOW))")

        End With

        grdSATSLSI1.DataSource = dst.Tables("SATSLSI1")

        Create_Summary(grdSATSLSI1, "ITEM_CODE", "Count")

        Dim G As UltraWinGrid.UltraGridGroup
        With grdSATSLSI1.DisplayLayout.Bands(0)
            G = .Groups.Add("ITEM")
            G.Header.Appearance.TextHAlign = HAlign.Center
            G.Header.Caption = "Items with Activity in Current and/or Previous Season"
            G.Header.Appearance.BackColor = Drawing.Color.White
            G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

            For Each COLUMN_NAME In New String() {"ITEM_CODE", "ITEM_SPEC_NO", "ITEM_DESC", "DEPT_CODE", "PROD_CODE", "COLLECTION_CODE", "MATL_CODE", _
                                                            "ITEM_CATGY_CODE_PREV", "ITEM_CATGY_CODE", "ITEM_RETAIL_PRICE", "ITEM_PRICE"}
                With .Columns(COLUMN_NAME)
                    .Group = G
                    .Header.Appearance.BackColor = Drawing.Color.White
                    .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    .CellAppearance.BackColor = Drawing.Color.Beige
                End With
                'Create_Summary(grdSATSLSI1, COLUMN_NAME)
            Next


            G = .Groups.Add("RTL")
            G.Header.Appearance.TextHAlign = HAlign.Center
            G.Header.Caption = "Sell Thru - Current and Prior Seasons"
            G.Header.Appearance.BackColor2 = Drawing.Color.LightPink
            G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

            For Each COLUMN_NAME In New String() _
            {"QTY_RTL_LY_MTD", "QTY_RTL_TY_MTD", "QTY_RTL_LS", "QTY_RTL_TS", "QTY_RTL", "AMT_RTL", "AMT_RTL_WS", "SELL_THRU_PCT", "QTY_EOW", "AMT_EOW", "AMT_EOW_WS"}
                With .Columns(COLUMN_NAME)
                    .Group = G
                    .Header.Appearance.BackColor2 = Drawing.Color.LightPink
                    .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    .CellAppearance.BackColor = Drawing.Color.Linen
                    .Width = 70
                    .Format = "###,##0"
                    If COLUMN_NAME = "SELL_THRU_PCT" Then
                        .Format = "###,##0.0"
                    End If

                    'If COLUMN_NAME = "QTY_RTL_LY_MTD" Then
                    '    .Hidden = True
                    'End If

                    If COLUMN_NAME <> "SELL_THRU_PCT" Then
                        Create_Summary(grdSATSLSI1, COLUMN_NAME)
                    End If
                End With
            Next



            G = .Groups.Add("RTL_CUST")
            G.Header.Appearance.TextHAlign = HAlign.Center
            G.Header.Caption = "Retail OH by Customer"
            G.Header.Appearance.BackColor2 = Drawing.Color.Orange
            G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

            For Each COLUMN_NAME In New String() {"QTY_EOW_C0", "QTY_EOW_C1", "QTY_EOW_C2", "QTY_EOW_C3", "QTY_EOW_C4", "QTY_EOW_C5", "QTY_EOW_C6", "QTY_EOW_C7", "QTY_EOW_C8", "QTY_EOW_C9"}
                With .Columns(COLUMN_NAME)
                    Dim C As Int16 = Val(COLUMN_NAME.Substring(COLUMN_NAME.Length - 1, 1))
                    .Header.Caption = CUST_CODEs(C)
                    .Group = G
                    .Header.Appearance.BackColor2 = Drawing.Color.Orange
                    .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    .CellAppearance.BackColor = Drawing.Color.LemonChiffon
                    .Width = 70
                    .Format = "###,##0"
                End With
                Create_Summary(grdSATSLSI1, COLUMN_NAME)
            Next


            G = .Groups.Add("SHP")
            G.Header.Appearance.TextHAlign = HAlign.Center
            G.Header.Caption = "Gross Shipments - Current and Prior Seasons"
            G.Header.Appearance.BackColor2 = Drawing.Color.Violet
            G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

            For Each COLUMN_NAME In New String() {"QTY_SHP_LY_MTD", "QTY_SHP_TY_MTD", "QTY_SHP_LS", "QTY_SHP_TS", "QTY_SHP", "AMT_SHP_WS"}
                With .Columns(COLUMN_NAME)
                    .Group = G
                    .Header.Appearance.BackColor2 = Drawing.Color.Violet
                    .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    '.CellAppearance.BackColor = Drawing.Color.LightGreen
                    .Width = 70
                    .Format = "###,##0"
                End With
                Create_Summary(grdSATSLSI1, COLUMN_NAME)
            Next


            G = .Groups.Add("OH")
            G.Header.Appearance.TextHAlign = HAlign.Center
            G.Header.Caption = "OH Inventory " & Format(Now, "MM/dd/yyyy")
            G.Header.Appearance.BackColor2 = Drawing.Color.Lime
            G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

            For Each COLUMN_NAME In New String() {"QOH_SHP", "QOH_RTN", "QOH_OTH", "QOH_ALL", "WOH_ALL"}
                With .Columns(COLUMN_NAME)
                    .Group = G
                    .Header.Appearance.BackColor2 = Drawing.Color.Lime
                    .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    .CellAppearance.BackColor = Drawing.Color.LightGreen
                    .Width = 70
                    .Format = "###,##0"
                End With
                Create_Summary(grdSATSLSI1, COLUMN_NAME)
            Next
        End With


        grdSATSLSI1.DisplayLayout.UseFixedHeaders = True
        With grdSATSLSI1.DisplayLayout.Bands("SATSLSI1")
            .Columns("ITEM_CODE").Header.Fixed = True
            '.Columns("ITEM_DESC").Header.Fixed = True
            '.Columns("PROD_CODE").Header.Fixed = True
            '.Columns("COLLECTION_CODE").Header.Fixed = True
            '.Columns("MATL_CODE").Header.Fixed = True
        End With

        Absx1.txtFor("OPS_YYYYPP").Text = ASCMAIN1.CYP ' .ToString.Substring(0, 4) & "12"

        ASCMAIN1.Add_Value_List(grdSATSLSI1, "ITEM_CATGY_CODE")
        ASCMAIN1.Add_Value_List(grdSATSLSI1, "ITEM_CATGY_CODE_PREV", "Select ITEM_CATGY_CODE, ITEM_CATGY_DESC from ICTCATG1")
        ASCMAIN1.Add_Value_List(grdSATSLSI1, "MATL_CODE", , New String() {":", ":Unknown"})

        'With grdSATSLSI1.DisplayLayout.Bands(0).SortedColumns
        '    .Clear()
        '    .Add("MATL_CODE", False, True)
        '    .Add("PROD_CODE", False, True)
        '    .Add("COLLECTION_CODE", False, True)
        'End With
        Show_Filter(grdSATSLSI1, True)
        grdSATSLSI1.DisplayLayout.GroupByBox.Hidden = False

        Show_Inactive(False)

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Call Validate_Code("OPS_YYYYPP")
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
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Done"
                Call Mode_Settings(False)

            Case "Print"
                Call Print_Report()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                '.Groups("Screen Control").Items("Print").Settings.Enabled = iScreenMode

            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdSATSLSI1.Visible = tf
        'UltraExplorerBar1.Groups("Summaries").Visible = False
        lblTS.Visible = tf
        lblLS.Visible = tf

        If ScreenMode Then

        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        dst.EnforceConstraints = False
        dst.Tables("SATSLSI1").Rows.Clear()
        dst.EnforceConstraints = True
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Compiling Historical Data")
        Application.DoEvents()

        Call Save_Header_Fields(UltraGroupBox1)

        Dim z As String = Absx1.txtFor("OPS_YYYYPP").Text
        RYP = z
        Create_SATSLSI1(RYP)

        Call ASCMAIN1.Progress("Now Loading Data from Database")

        dst.EnforceConstraints = False

        Fill_Records("SATSLSI1")

        grdSATSLSI1.Rows.ExpandAll(True)

        EnforceConstraints(True)

        ASCMAIN1.Progress("Now Setting Up Screen")

        Sort_grdColumns(grdSATSLSI1, "DEPT_CODE,PROD_CODE,MATL_CODE,ITEM_CATGY_CODE")
        Setup_Grid_Caption()

        Dim RYP_LEGEND = ASCMAIN1.Get_Legend(RYP, False, True)
        Dim LYP_LEGEND = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(RYP, -12), False, True)
        With grdSATSLSI1.DisplayLayout.Bands(0)
            .Columns("QTY_RTL_LY_MTD").Header.Caption = LYP_LEGEND
            .Columns("QTY_RTL_TY_MTD").Header.Caption = RYP_LEGEND
            .Columns("QTY_SHP_LY_MTD").Header.Caption = LYP_LEGEND
            .Columns("QTY_SHP_TY_MTD").Header.Caption = RYP_LEGEND
        End With

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")

    End Sub

    Sub Update_Record()
        Call BeginTrans()
        Call CommitTrans("Update Complete")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdSATSLSI1, "SSBSB", "Show Filter", "Show GroupBox", "Item Status Inquiry", "Show Inactive", "Restore Original Sort")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
        End If

        'If tlb_pop.Tools.Exists("Include Inactive") Then
        'End If


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If e.Tool.Key <> "grdSATCSLSS" Then
            '    Dim tlb_sbt As UltraWinToolbars.StateButtonTool
            '    tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            '    tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
            '    tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            '    tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
            'End If

            Select Case e.SourceControl.Name
                Case "grdSATCSLS1"



            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked

            Case "Show Inactive"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Inactive(tlb_sbt.Checked)

            Case "Restore Original Sort"
                Sort_grdColumns(grdSATSLSI1, "DEPT_CODE,PROD_CODE,MATL_CODE,ITEM_CATGY_CODE")
        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
        End Select


    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "OPS_YYYYPP"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    Call Click_Command("Load", e)
                End If
        End Select

    End Sub

#End Region

    Sub Create_SATSLSI1(ByVal RYP As String)

        Dim LS_P1 As String = ""
        Dim LS_P2 As String = ""
        Dim TS_P1 As String = ""
        Dim TS_P2 As String = ""

        Dim OPS_YYYY_LS As String = ""
        Dim SEASON_LS As String = ""

        If RYP = "" Then
            ReDim CUST_CODEs(9)
            CUST_CODEs(0) = "LGI"
            CUST_CODEs(1) = "NEIMANM10"
            CUST_CODEs(2) = "NORDSTR10"
            CUST_CODEs(3) = "VONMAUR10"
            CUST_CODEs(4) = "FINKS10"
            CUST_CODEs(5) = "SAKSFIF10"
            CUST_CODEs(6) = "BLOOMIES10"
            CUST_CODEs(7) = "HOLTREN10"
            CUST_CODEs(8) = "CARLYLE10"
            CUST_CODEs(9) = "BAILEYBA10"
        End If

        If RYP <> "" Then
            TS_P2 = Mid(RYP, 1, 4) & IIf(Val(Mid(RYP, 5, 2)) < 7, "06", "12")
            TS_P1 = ASCMAIN1.Period_Calc(TS_P2, -5)
            LS_P2 = ASCMAIN1.Period_Calc(TS_P1, -1)
            LS_P1 = ASCMAIN1.Period_Calc(LS_P2, -5)

            OPS_YYYY_LS = Mid(LS_P1, 1, 4)
            If Mid(LS_P1, 5, 2) = "01" Then
                SEASON_LS = "F"
            Else
                SEASON_LS = "S"
            End If

            lblTS.Text = "TS (This Season): " & ASCMAIN1.Get_Legend(TS_P1) & " thru " & ASCMAIN1.Get_Legend(TS_P2)
            lblLS.Text = "LS (Last Season): " & ASCMAIN1.Get_Legend(LS_P1) & " thru " & ASCMAIN1.Get_Legend(LS_P2)
        End If

        Dim SQL_RS As String = "Select RSTRETL1.ITEM_CODE" & vbCrLf _
        & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYPP BETWEEN '" & ASCMAIN1.Period_Calc(RYP, -12) & "' AND '" & ASCMAIN1.Period_Calc(RYP, -12) & "' THEN RSTRETL1.QTY_SOLD ELSE 0 END) QTY_RTL_LY_MTD" & vbCrLf _
        & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYPP BETWEEN '" & RYP & "' AND '" & RYP & "' THEN RSTRETL1.QTY_SOLD ELSE 0 END) QTY_RTL_TY_MTD" & vbCrLf _
        & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYPP BETWEEN '" & LS_P1 & "' AND '" & LS_P2 & "' THEN RSTRETL1.QTY_SOLD ELSE 0 END) QTY_RTL_LS" & vbCrLf _
        & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYPP BETWEEN '" & TS_P1 & "' AND '" & TS_P2 & "' THEN RSTRETL1.QTY_SOLD ELSE 0 END) QTY_RTL_TS" & vbCrLf _
        & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYPP BETWEEN '" & LS_P1 & "' AND '" & LS_P2 & "' THEN RSTRETL1.AMT_SOLD ELSE 0 END) AMT_RTL_LS" & vbCrLf _
        & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYPP BETWEEN '" & TS_P1 & "' AND '" & TS_P2 & "' THEN RSTRETL1.AMT_SOLD ELSE 0 END) AMT_RTL_TS" & vbCrLf _
        & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYWW = X.OPS_YYYYWW_max THEN RSTRETL1.QTY_EOW ELSE 0 END) QTY_EOW" & vbCrLf _
        & ", SUM (CASE WHEN EDT852T1.EDI_DEPT_NO = 'LGI' and RSTRETL1.OPS_YYYYWW = X.OPS_YYYYWW_max THEN RSTRETL1.QTY_EOW ELSE 0 END) QTY_EOW_C0" & vbCrLf _
        & ", SUM (CASE WHEN RSTRETL1.CUST_CODE = '" & CUST_CODEs(1) & "' and RSTRETL1.OPS_YYYYWW = X.OPS_YYYYWW_max THEN RSTRETL1.QTY_EOW ELSE 0 END) QTY_EOW_C1" & vbCrLf _
        & ", SUM (CASE WHEN RSTRETL1.CUST_CODE = '" & CUST_CODEs(2) & "' and RSTRETL1.OPS_YYYYWW = X.OPS_YYYYWW_max THEN RSTRETL1.QTY_EOW ELSE 0 END) QTY_EOW_C2" & vbCrLf _
        & ", SUM (CASE WHEN RSTRETL1.CUST_CODE = '" & CUST_CODEs(3) & "' and RSTRETL1.OPS_YYYYWW = X.OPS_YYYYWW_max THEN RSTRETL1.QTY_EOW ELSE 0 END) QTY_EOW_C3" & vbCrLf _
        & ", SUM (CASE WHEN RSTRETL1.CUST_CODE = '" & CUST_CODEs(4) & "' and RSTRETL1.OPS_YYYYWW = X.OPS_YYYYWW_max THEN RSTRETL1.QTY_EOW ELSE 0 END) QTY_EOW_C4" & vbCrLf _
        & ", SUM (CASE WHEN RSTRETL1.CUST_CODE = '" & CUST_CODEs(5) & "' and RSTRETL1.OPS_YYYYWW = X.OPS_YYYYWW_max THEN RSTRETL1.QTY_EOW ELSE 0 END) QTY_EOW_C5" & vbCrLf _
        & ", SUM (CASE WHEN RSTRETL1.CUST_CODE = '" & CUST_CODEs(6) & "' and RSTRETL1.OPS_YYYYWW = X.OPS_YYYYWW_max THEN RSTRETL1.QTY_EOW ELSE 0 END) QTY_EOW_C6" & vbCrLf _
        & ", SUM (CASE WHEN RSTRETL1.CUST_CODE = '" & CUST_CODEs(7) & "' and RSTRETL1.OPS_YYYYWW = X.OPS_YYYYWW_max THEN RSTRETL1.QTY_EOW ELSE 0 END) QTY_EOW_C7" & vbCrLf _
        & ", SUM (CASE WHEN RSTRETL1.CUST_CODE = '" & CUST_CODEs(8) & "' and RSTRETL1.OPS_YYYYWW = X.OPS_YYYYWW_max THEN RSTRETL1.QTY_EOW ELSE 0 END) QTY_EOW_C8" & vbCrLf _
        & ", SUM (CASE WHEN RSTRETL1.CUST_CODE = '" & CUST_CODEs(9) & "' and RSTRETL1.OPS_YYYYWW = X.OPS_YYYYWW_max THEN RSTRETL1.QTY_EOW ELSE 0 END) QTY_EOW_C9" & vbCrLf _
        & ", 0 QTY_SHP_LY_MTD, 0 QTY_SHP_TY_MTD, 0 QTY_SHP_LS, 0 QTY_SHP_TS, 0 AMT_SHP_LS, 0 AMT_SHP_TS" & vbCrLf _
        & ", 0 QOH_SHP, 0 QOH_RTN, 0 QOH_OTH" & vbCrLf _
        & " from RSTRETL1,EDT852T1,(" & "Select CUST_CODE, MAX(OPS_YYYYWW) OPS_YYYYWW_max from EDT852T1 " & vbCrLf _
        & "   where EDI_STATUS in ('1','M') " & vbCrLf _
        & " and OPS_YYYYPP between '" & ASCMAIN1.Period_Calc(RYP, -2) & "' and '" & RYP & "'" & vbCrLf _
        & " group by CUST_CODE) X" & vbCrLf _
        & " where RSTRETL1.OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(RYP, -12) & "' and RSTRETL1.OPS_YYYYPP <= '" & RYP & "'" & vbCrLf _
        & " and EDT852T1.EDI_DOC_SEQ_NO = RSTRETL1.EDI_DOC_SEQ_NO" & vbCrLf _
        & " and X.CUST_CODE (+) = RSTRETL1.CUST_CODE" & vbCrLf _
        & " and (NVL(RSTRETL1.QTY_SOLD,0) <> 0 OR NVL(RSTRETL1.QTY_EOW,0) <> 0)" & vbCrLf _
        & " group by RSTRETL1.ITEM_CODE"


        Dim SQL_WS As String = "Select SATSSUMI.ITEM_CODE" & vbCrLf _
        & ", 0 QTY_RTL_LY_MTD, 0 QTY_RTL_TY_MTD, 0 QTY_RTL_LS, 0 QTY_RTL_TS, 0 AMT_RTL_LS, 0 AMT_RTL_TS, 0 QTY_EOW" & vbCrLf _
        & ", 0 QTY_EOW_C0, 0 QTY_EOW_C1, 0 QTY_EOW_C2, 0 QTY_EOW_C3, 0 QTY_EOW_C4, 0 QTY_EOW_C5, 0 QTY_EOW_C6, 0 QTY_EOW_C7, 0 QTY_EOW_C8, 0 QTY_EOW_C9" & vbCrLf _
        & ", SUM (CASE WHEN SATSSUMI.OPS_YYYYPP BETWEEN '" & ASCMAIN1.Period_Calc(RYP, -12) & "' AND '" & ASCMAIN1.Period_Calc(RYP, -12) & "' THEN SATSSUMI.ORDR_QTY_SHIP ELSE 0 END) QTY_SHP_LY_MTD" & vbCrLf _
        & ", SUM (CASE WHEN SATSSUMI.OPS_YYYYPP BETWEEN '" & RYP & "' AND '" & RYP & "' THEN SATSSUMI.ORDR_QTY_SHIP ELSE 0 END) QTY_SHP_TY_MTD" & vbCrLf _
        & ", SUM (CASE WHEN SATSSUMI.OPS_YYYYPP BETWEEN '" & LS_P1 & "' AND '" & LS_P2 & "' THEN SATSSUMI.ORDR_QTY_SHIP ELSE 0 END) QTY_SHP_LS" & vbCrLf _
        & ", SUM (CASE WHEN SATSSUMI.OPS_YYYYPP BETWEEN '" & TS_P1 & "' AND '" & TS_P2 & "' THEN SATSSUMI.ORDR_QTY_SHIP ELSE 0 END) QTY_SHP_TS" & vbCrLf _
        & ", SUM (CASE WHEN SATSSUMI.OPS_YYYYPP BETWEEN '" & LS_P1 & "' AND '" & LS_P2 & "' THEN SATSSUMI.ORDR_AMT_SHIP ELSE 0 END) AMT_SHP_LS" & vbCrLf _
        & ", SUM (CASE WHEN SATSSUMI.OPS_YYYYPP BETWEEN '" & TS_P1 & "' AND '" & TS_P2 & "' THEN SATSSUMI.ORDR_AMT_SHIP ELSE 0 END) AMT_SHP_TS" & vbCrLf _
        & ", 0 QOH_SHP, 0 QOH_RTN, 0 QOH_OTH" & vbCrLf _
        & " from SATSSUMI,ARTCUST1" & vbCrLf _
        & " where SATSSUMI.OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(RYP, -12) & "' and SATSSUMI.OPS_YYYYPP <= '" & RYP & "'" & vbCrLf _
        & "  and SATSSUMI.INV_TYPE = 'I' and SATSSUMI.ORDR_QTY_SHIP <> 0" & vbCrLf _
        & "  and ARTCUST1.CUST_CODE = SATSSUMI.CUST_CODE" & vbCrLf _
        & "  and NVL(ARTCUST1.TRADE_CLASS_CODE,'XXX') <> 'CON'" & vbCrLf _
        & " group by SATSSUMI.ITEM_CODE"


        Dim SQL_OH As String = "Select ICTSTAT2.ITEM_CODE" & vbCrLf _
        & ", 0 QTY_RTL_LY_MTD, 0 QTY_RTL_TY_MTD, 0 QTY_RTL_LS, 0 QTY_RTL_TS, 0 AMT_RTL_LS, 0 AMT_RTL_TS, 0 QTY_EOW" & vbCrLf _
        & ", 0 QTY_EOW_C0, 0 QTY_EOW_C1, 0 QTY_EOW_C2, 0 QTY_EOW_C3, 0 QTY_EOW_C4, 0 QTY_EOW_C5, 0 QTY_EOW_C6, 0 QTY_EOW_C7, 0 QTY_EOW_C8, 0 QTY_EOW_C9" & vbCrLf _
        & ", 0 QTY_SHP_LY_MTD, 0 QTY_SHP_TY_MTD, 0 QTY_SHP_LS, 0 QTY_SHP_TS, 0 AMT_SHP_LS, 0 AMT_SHP_TS" & vbCrLf _
        & ", SUM (CASE WHEN ICTWHSE1.WHSE_TYPE = 'S' THEN ICTSTAT2.WHSE_QTY_ON_HAND ELSE 0 END) QOH_SHP" & vbCrLf _
        & ", SUM (CASE WHEN ICTWHSE1.WHSE_TYPE = 'R' THEN ICTSTAT2.WHSE_QTY_ON_HAND ELSE 0 END) QOH_RTN" & vbCrLf _
        & ", SUM (CASE WHEN NVL(ICTWHSE1.WHSE_TYPE,'?') NOT IN ('S','R','N') THEN ICTSTAT2.WHSE_QTY_ON_HAND ELSE 0 END) QOH_OTH" & vbCrLf _
        & " from ICTSTAT2,ICTWHSE1 where ICTSTAT2.WHSE_QTY_ON_HAND <> 0" & vbCrLf _
        & " and ICTWHSE1.WHSE_CODE = ICTSTAT2.WHSE_CODE" & vbCrLf _
        & " group by ICTSTAT2.ITEM_CODE"

        Dim SQL_I As String = "Select ITEM_CODE" & vbCrLf _
        & ", SUM (QTY_RTL_LY_MTD) QTY_RTL_LY_MTD, SUM (QTY_RTL_TY_MTD) QTY_RTL_TY_MTD, SUM (QTY_RTL_LS) QTY_RTL_LS, SUM (QTY_RTL_TS) QTY_RTL_TS, SUM (AMT_RTL_LS) AMT_RTL_LS, SUM (AMT_RTL_TS) AMT_RTL_TS, SUM (QTY_EOW) QTY_EOW" & vbCrLf _
        & ", SUM (QTY_EOW_C0) QTY_EOW_C0, SUM (QTY_EOW_C1) QTY_EOW_C1, SUM (QTY_EOW_C2) QTY_EOW_C2, SUM (QTY_EOW_C3) QTY_EOW_C3, SUM (QTY_EOW_C4) QTY_EOW_C4" & vbCrLf _
        & ", SUM (QTY_EOW_C5) QTY_EOW_C5, SUM (QTY_EOW_C6) QTY_EOW_C6, SUM (QTY_EOW_C7) QTY_EOW_C7, SUM (QTY_EOW_C8) QTY_EOW_C8, SUM (QTY_EOW_C9) QTY_EOW_C9" & vbCrLf _
        & ", SUM (QTY_SHP_LY_MTD) QTY_SHP_LY_MTD, SUM (QTY_SHP_TY_MTD) QTY_SHP_TY_MTD, SUM (QTY_SHP_LS) QTY_SHP_LS, SUM (QTY_SHP_TS) QTY_SHP_TS, SUM (AMT_SHP_LS) AMT_SHP_LS, SUM (AMT_SHP_TS) AMT_SHP_TS" & vbCrLf _
        & ", SUM (QOH_SHP) QOH_SHP, SUM (QOH_RTN) QOH_RTN, SUM (QOH_OTH) QOH_OTH" & vbCrLf _
        & " from (" & SQL_RS & " UNION " & SQL_WS & " UNION " & SQL_OH & ")" & vbCrLf _
        & " group by ITEM_CODE"

        ASCMAIN1.sql = "Select I.*, ICTITEM1.ITEM_SPEC_NO, ICTITEM1.ITEM_DESC" & vbCrLf _
        & ", ICTITEM1.DEPT_CODE, ICTITEM1.PROD_CODE, ICTITEM1.COLLECTION_CODE, ICTITEM1.MATL_CODE" & vbCrLf _
        & ", DPTABCP0.ITEM_CATGY_CODE ITEM_CATGY_CODE_PREV, ICTITEM1.ITEM_CATGY_CODE" & vbCrLf _
        & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_PRICE" & vbCrLf _
        & " from ICTITEM1, DPTABCP0, (" & SQL_I & ") I" & vbCrLf _
        & " where I.ITEM_CODE = ICTITEM1.ITEM_CODE" & vbCrLf _
        & " and DPTABCP0.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
        & " and DPTABCP0.OPS_YYYY (+) = '" & OPS_YYYY_LS & "'" & vbCrLf _
        & " and DPTABCP0.SEASON (+) = '" & SEASON_LS & "'"

        If SATSLSI1 = "" Then
            SATSLSI1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SATSLSI1 & " Add Primary Key (ITEM_CODE)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SATSLSI1)
            ASCDATA1.ExecuteSQL("Insert into " & SATSLSI1 & " " & ASCMAIN1.sql)
        End If

    End Sub

    Sub Print_Report()
        Dim SUBT As String = ""

        Call Print_Report_Begin()

        'CR_params.Add("DETAIL", "Y")
        'Generate_Report("SARSLSJ5", "Sales Analysis by Rep/Customer-Rank", SUBT)

        Call Print_Report_End()
    End Sub

    Private Sub optCode_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Setup_Summary()
    End Sub

    Sub Setup_Summary()
        If SELECTION_NO = 0 Then Exit Sub
    End Sub

    Private Sub grdSATSLSI1_AfterGroupPosChanged(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterGroupPosChangedEventArgs) Handles grdSATSLSI1.AfterGroupPosChanged

    End Sub

    Private Sub grdSATSLSI1_AfterSortChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BandEventArgs) Handles grdSATSLSI1.AfterSortChange
        Setup_Grid_Caption()

    End Sub

    Private Sub grdSATSLSI1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSATSLSI1.InitializeRow
        If e.Row.Cells("ITEM_CATGY_CODE_PREV").Value & "" <> "" And _
           e.Row.Cells("ITEM_CATGY_CODE").Value & "" <> "" And _
           e.Row.Cells("ITEM_CATGY_CODE_PREV").Value & "" <> e.Row.Cells("ITEM_CATGY_CODE").Value & "" Then
            e.Row.Cells("ITEM_CATGY_CODE").Appearance.BackColor = Drawing.Color.Yellow
        End If
    End Sub

    Sub Show_Inactive(ByVal tf As Boolean)
        Dim dvw As DataView = DirectCast(grdSATSLSI1.DataSource, DataTable).DefaultView
        If tf Then
            dvw.RowFilter = ""
        Else
            dvw.RowFilter = "ITEM_CATGY_CODE <> 'I'"
        End If

    End Sub

    Sub Setup_Grid_Caption()
        Dim CAPTION As String = Me.Text
        Dim GROUPS As String = ""
        Dim SORTS As String = ""
        For Each GCOL As UltraWinGrid.UltraGridColumn In grdSATSLSI1.DisplayLayout.Bands(0).SortedColumns
            If GCOL.IsGroupByColumn Then
                GROUPS &= "," & GCOL.Header.Caption
            Else
                SORTS &= "," & GCOL.Header.Caption
            End If
        Next
        If GROUPS <> "" Then
            CAPTION &= ", Grouped by " & Mid(GROUPS, 2)
        End If
        If SORTS <> "" Then
            CAPTION &= ", Sorted by " & Mid(SORTS, 2)
        End If

        grdSATSLSI1.Text = CAPTION
    End Sub
End Class