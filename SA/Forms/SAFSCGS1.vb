Public Class SAFSCGS1

    Dim RYP0 As String
    Dim RYP1 As String
    Dim SATSCGS1 As String
    Dim sqlSATSCGS1 As String
    Dim sqlBRAND_CODEs As String
    Dim BRAND_CODEs As New List(Of String)
    Dim Periods As Integer

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Create_SATSCGS1(True)
            ASCMAIN1.sql = "Select * from " & SATSCGS1
            Create_TDA(.Tables.Add, "SATSCGS1", "**", 0, False, "", 0)

            ' WGP -> WGP

            Dim T(5) As String
            With .Tables("SATSCGS1")
                For Each BRAND_CODE As String In BRAND_CODEs
                    .Columns.Add("'" & BRAND_CODE & "'_WGP", GetType(System.Decimal), _
                                 "['" & BRAND_CODE & "'_SLS_WSL" & "]-[" & "'" & BRAND_CODE & "'_CGS]")
                    .Columns.Add("'" & BRAND_CODE & "'_WGPPCT", GetType(System.Decimal), _
                                "IIF(['" & BRAND_CODE & "'_SLS_WSL]=0,0," & "100*['" & BRAND_CODE & "'_WGP]" & "/" & "['" & BRAND_CODE & "'_SLS_WSL]" & ")")
                    .Columns.Add("'" & BRAND_CODE & "'_RCGSPCT", GetType(System.Decimal), _
                               "IIF(['" & BRAND_CODE & "'_SLS_RTL]=0,0," & "100*['" & BRAND_CODE & "'_CGS]" & "/" & "['" & BRAND_CODE & "'_SLS_RTL]" & ")")
                    .Columns.Add("'" & BRAND_CODE & "'_WCGSPCT", GetType(System.Decimal), _
                               "IIF(['" & BRAND_CODE & "'_SLS_WSL]=0,0," & "100*['" & BRAND_CODE & "'_CGS]" & "/" & "['" & BRAND_CODE & "'_SLS_WSL]" & ")")

                    T(1) &= "+ISNULL(['" & BRAND_CODE & "'_SLS_WSL],0)"
                    T(2) &= "+ISNULL(['" & BRAND_CODE & "'_SLS_RTL],0)"
                    T(3) &= "+ISNULL(['" & BRAND_CODE & "'_CGS],0)"
                    T(4) &= "+ISNULL(['" & BRAND_CODE & "'_SLSNC_WSL],0)"
                    T(5) &= "+ISNULL(['" & BRAND_CODE & "'_CGSNC],0)"

                Next

                Dim BRAND_CODE_total As String = "TOTAL"

                With dst.Tables("SATSCGS1").Columns
                    .Add("'" & BRAND_CODE_total & "'_SLS_WSL", GetType(System.Decimal), Mid(T(1), 2))
                    .Add("'" & BRAND_CODE_total & "'_SLS_RTL", GetType(System.Decimal), Mid(T(2), 2))
                    .Add("'" & BRAND_CODE_total & "'_CGS", GetType(System.Decimal), Mid(T(3), 2))
                    .Add("'" & BRAND_CODE_total & "'_SLSNC_WSL", GetType(System.Decimal), Mid(T(4), 2))
                    .Add("'" & BRAND_CODE_total & "'_CGSNC", GetType(System.Decimal), Mid(T(5), 2))
                End With

                .Columns.Add("'" & BRAND_CODE_total & "'_WGP", GetType(System.Decimal), _
             "['" & BRAND_CODE_total & "'_SLS_WSL" & "]-[" & "'" & BRAND_CODE_total & "'_CGS]")
                .Columns.Add("'" & BRAND_CODE_total & "'_WGPPCT", GetType(System.Decimal), _
                            "IIF(['" & BRAND_CODE_total & "'_SLS_WSL]=0,0," & "100*['" & BRAND_CODE_total & "'_WGP]" & "/" & "['" & BRAND_CODE_total & "'_SLS_WSL]" & ")")
                .Columns.Add("'" & BRAND_CODE_total & "'_RCGSPCT", GetType(System.Decimal), _
                           "IIF(['" & BRAND_CODE_total & "'_SLS_RTL]=0,0," & "100*['" & BRAND_CODE_total & "'_CGS]" & "/" & "['" & BRAND_CODE_total & "'_SLS_RTL]" & ")")
                .Columns.Add("'" & BRAND_CODE_total & "'_WCGSPCT", GetType(System.Decimal), _
                           "IIF(['" & BRAND_CODE_total & "'_SLS_WSL]=0,0," & "100*['" & BRAND_CODE_total & "'_CGS]" & "/" & "['" & BRAND_CODE_total & "'_SLS_WSL]" & ")")

            End With
        End With

        grdSATSCGS1.DataSource = dst.Tables("SATSCGS1")
        With grdSATSCGS1.DisplayLayout.Bands(0)

            '.SummaryFooterCaption = "X"

            Dim G As UltraWinGrid.UltraGridGroup = .Groups.Add
            G.Key = "CUST_CODE"
            G.Header.Appearance.BackColor = Drawing.Color.White
            G.Header.Appearance.BackColor2 = Drawing.Color.Gold
            G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            G.Header.Caption = "Customer"
            G.Header.Fixed = True
            With .Columns("CUST_CODE")
                .Group = G
                .Header.Caption = "Code"
            End With
            With .Columns("TRADE_CLASS_CODE")
                .Group = G
                .Header.Caption = "Trade Class"
                .HiddenWhenGroupBy = DefaultableBoolean.True

            End With

            For i As Integer = 0 To BRAND_CODEs.Count ' For Each BRAND_CODE As String In BRAND_CODEs
                Dim BRAND_CODE As String = ""
                If i = 0 Then
                    BRAND_CODE = "TOTAL"
                Else
                    BRAND_CODE = BRAND_CODEs(i - 1)
                End If
                G = .Groups.Add
                G.Key = BRAND_CODE
                G.Header.Caption = BRAND_CODE
                G.Header.Appearance.BackColor = Drawing.Color.White
                G.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

                For Each SFX As String In New String() {"SLS_WSL", "SLS_RTL", "CGS", "RCGSPCT", "WGP", "WGPPCT", "WCGSPCT", "SLSNC_WSL", "CGSNC"}
                    Dim COLUMN_NAME As String = "'" & BRAND_CODE & "'_" & SFX
                    With .Columns(COLUMN_NAME)
                        .Group = G
                        Select Case SFX
                            Case "SLS_WSL"
                                .Header.Caption = "RS Sls@W/S"
                                .Header.ToolTipText = "Re-Saleable Items Shipped @ Wholesale Price"
                            Case "SLS_RTL"
                                .Header.Caption = "RS Sls@Rtl"
                                .Header.ToolTipText = "Re-Saleable Items Shipped @ Retail Price"
                            Case "CGS"
                                .Header.Caption = "RS CGS"
                                .Header.ToolTipText = "Re-Saleable Items Shipped @ Standard Cost"
                            Case "RCGSPCT"
                                .Header.Caption = "RS C%R"
                                .Header.ToolTipText = "Re-Saleable Items Shipped, Std Cost as a % of Retail"
                                .CellAppearance.BackColor = Drawing.Color.LightBlue
                                .Format = "##0.0"
                                .Width = 90
                            Case "WGP"
                                .Header.Caption = "RS $GP"
                                .Header.ToolTipText = "Re-Saleable Items Shipped, Wholesale Gross Profit"
                            Case "WGPPCT"
                                .Header.Caption = "RS GP%"
                                .Header.ToolTipText = "Re-Saleable Items Shipped, Wholesale Gross Profit%"
                                .Format = "##0.0"
                                .Width = 90
                            Case "WCGSPCT"
                                .Header.Caption = "RS C%W"
                                .Header.ToolTipText = "Re-Saleable Items Shipped, Std Cost as a % of Wholesale"
                                .CellAppearance.BackColor = Drawing.Color.LightBlue
                                .Format = "##0.0"
                                .Width = 90
                            Case "SLSNC_WSL"
                                .Header.Caption = "NC Sls"
                                .Header.ToolTipText = "No-Charge Items Shipped, Wholesale Sales"
                                .CellAppearance.BackColor = Drawing.Color.LightGray
                            Case "CGSNC"
                                .Header.Caption = "NC CGS"
                                .Header.ToolTipText = "No-Charge Items Shipped @ Standard Cost"
                                .CellAppearance.BackColor = Drawing.Color.LightGray
                        End Select

                        .Header.Appearance.BackColor = Drawing.Color.White
                        .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                        .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    End With
                    If SFX = "WGPPCT" Or SFX = "RCGSPCT" Or SFX = "WCGSPCT" Then
                        Create_Summary(grdSATSCGS1, COLUMN_NAME, "Custom")
                    Else
                        Create_Summary(grdSATSCGS1, COLUMN_NAME)
                    End If

                Next
            Next
        End With

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 12, -11)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -60, 12, 0)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                RYP0 = Absx1.cmbFor("RYP0").Value
                RYP1 = Absx1.cmbFor("RYP1").Value
                Periods = ASCMAIN1.Period_Diff(RYP0, RYP1) + 1
                If Periods < 1 Or Periods > 12 Then
                    EMsg &= "Period Range must be between 1 and 12"
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Print"
                Print_Report()

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                '.Groups("Screen Control").Items("Print").Settings.Enabled = iScreenMode
                .Groups("View Options").Visible = False ' ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdSATSCGS1.Visible = ScreenMode

        lblMonths.Visible = ScreenMode
        lblMonths.Text = CStr(Periods) & " Mos"

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        dst.Tables("SATSCGS1").Rows.Clear()
        EnforceConstraints(True)
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Compiling Historical Data")

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("Now Loading Data from Database")

        Create_SATSCGS1()
        EnforceConstraints(False)
        Fill_Records("SATSCGS1")
        EnforceConstraints(True)

        Sort_grdColumns(grdSATSCGS1, "CUST_CODE")
        grdSATSCGS1.DisplayLayout.Bands(0).SortedColumns.Add("TRADE_CLASS_CODE", False, True)
        grdSATSCGS1.Text = "Sales & CGS Analysis for " & IIf(RYP0 = RYP1, Absx1.cmbFor("RYP0").Text, Absx1.cmbFor("RYP0").Text & " thru " & Absx1.cmbFor("RYP1").Text)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()
        CommitTrans("Update Complete")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSATSCGS1, "SSS", "Show Filter", "Show GroupBox", "Show Pins", "Sales Summary by Customer")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdSATCSLS1"

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Sales Summary by Customer"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Context_Launch("View", CUST_CODE, e.Tool.Key, "SAFCSLS1")
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
                    Click_Command("Load", e)
                End If
        End Select
    End Sub

#End Region

    Sub Create_SATSCGS1(Optional initialize As Boolean = False)

        If initialize Then
            'ASCMAIN1.sql = "Select Distinct BRAND_CODE from ICTCOLL1" & vbCrLf _
            '    & " where COLLECTION_CODE in " & vbCrLf _
            '    & "(Select Distinct COLLECTION_CODE from SOTINVH2,ICTITEM1" & vbCrLf _
            '    & " where ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
            '    & "   and ORDR_YYYYPP_UPDATED between :PARM1 and :PARM2"

            BRAND_CODEs.Clear()
            sqlBRAND_CODEs = ""
            ASCMAIN1.sql = "Select BRAND_CODE from ICTBRAN1"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "BRAND_CODE")
                Dim BRAND_CODE As String = row.Item("BRAND_CODE")
                sqlBRAND_CODEs &= ",'" & BRAND_CODE & "'"
                BRAND_CODEs.Add(BRAND_CODE)
            Next

            sqlSATSCGS1 = "Select * from ( " & vbCrLf _
                & "Select SOTINVH2.CUST_CODE, ARTCUST1.TRADE_CLASS_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
                & ", SUM (DECODE(ICTITEM1.ITEM_SNU_CODE,'S',SOTINVH2.ORDR_QTY_SHIP,0) * SOTINVH2.ORDR_UNIT_PRICE) SLS_WSL" & vbCrLf _
                & ", SUM (DECODE(ICTITEM1.ITEM_SNU_CODE,'S',SOTINVH2.ORDR_QTY_SHIP,0) * SOTINVH2.ITEM_RETAIL_PRICE) SLS_RTL" & vbCrLf _
                & ", SUM (DECODE(ICTITEM1.ITEM_SNU_CODE,'S',SOTINVH2.ORDR_QTY_SHIP,0) * DECODE(SOTINVH2.WHSE_CODE,NULL,0,SOTINVH2.ITEM_UNIT_COST)) CGS" & vbCrLf _
                & ", SUM (DECODE(ICTITEM1.ITEM_SNU_CODE,'S',0,SOTINVH2.ORDR_QTY_SHIP) * SOTINVH2.ORDR_UNIT_PRICE) SLSNC_WSL" & vbCrLf _
                & ", SUM (DECODE(ICTITEM1.ITEM_SNU_CODE,'S',0,SOTINVH2.ORDR_QTY_SHIP) * DECODE(SOTINVH2.WHSE_CODE,NULL,0,SOTINVH2.ITEM_UNIT_COST)) CGSNC" & vbCrLf _
                & " from SOTINVH2,ICTITEM1,ICTCOLL1,ARTCUST1" & vbCrLf _
                & " where SOTINVH2.ORDR_YYYYPP_UPDATED between :PARM1 and :PARM2" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & " group by ARTCUST1.TRADE_CLASS_CODE, SOTINVH2.CUST_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
                & ")" & vbCrLf _
                & " Pivot " & vbCrLf _
                & "(" & vbCrLf _
                & "  Sum(SLS_WSL) AS SLS_WSL, Sum(SLS_RTL) AS SLS_RTL, Sum(CGS) AS CGS, SUM (SLSNC_WSL) AS SLSNC_WSL, SUM (CGSNC) AS CGSNC" & vbCrLf _
                & "  for BRAND_CODE" & vbCrLf _
                & "  in (" & Mid(sqlBRAND_CODEs, 2) & ")" & vbCrLf _
                & ")" & vbCrLf _
                & "order by CUST_CODE"
            SATSCGS1 = ASCMAIN1.Temp_Table(Replace(Replace(sqlSATSCGS1, ":PARM1", "''"), ":PARM2", "''"))
            ASCDATA1.ExecuteSQL("Alter Table " & SATSCGS1 & " Add Primary Key (CUST_CODE)")
        Else

            ASCDATA1.ExecuteSQL("Truncate Table " & SATSCGS1)
            ASCDATA1.ExecuteSQL("Insert into " & SATSCGS1 & " " & sqlSATSCGS1, "VV", New String() {RYP0, RYP1})

        End If
    End Sub

    Sub Print_Report()
        Dim SUBT As String = ""

        Print_Report_Begin()

        'CR_params.Add("DETAIL", "Y")
        'Generate_Report("SARSLSJ5", "Sales Analysis by Rep/Customer-Rank", SUBT)

        Print_Report_End()
    End Sub

    Private Sub grdSATSCGS1_InitializeGroupByRow(sender As Object, e As UltraWinGrid.InitializeGroupByRowEventArgs) Handles grdSATSCGS1.InitializeGroupByRow
        e.Row.Description = e.Row.Value
    End Sub

    Public Overrides Function CustomSummary_End( _
        ByVal summarySettings As UltraWinGrid.SummarySettings, _
        ByVal rows As UltraWinGrid.RowsCollection, _
        ByVal CustomValue As Double, _
        ByVal grd As UltraWinGrid.UltraGrid) As Double

        CustomValue = 0
        Dim TOTALS As New Dictionary(Of String, Decimal)

        Select Case grd.Name
            Case "grdSATSCGS1"


                'For Each BRAND_CODE As String In BRAND_CODEs
                '    .Columns.Add("'" & BRAND_CODE & "'_WGP", GetType(System.Decimal), _
                '                 "['" & BRAND_CODE & "'_SLS_WSL" & "]-[" & "'" & BRAND_CODE & "'_CGS]")
                '    .Columns.Add("'" & BRAND_CODE & "'_WGPPCT", GetType(System.Decimal), _
                '                "IIF(['" & BRAND_CODE & "'_SLS_WSL]=0,0," & "100*['" & BRAND_CODE & "'_WGP]" & "/" & "['" & BRAND_CODE & "'_SLS_WSL]" & ")")
                '    .Columns.Add("'" & BRAND_CODE & "'_RCGSPCT", GetType(System.Decimal), _
                '               "IIF(['" & BRAND_CODE & "'_SLS_RTL]=0,0," & "100*['" & BRAND_CODE & "'_CGS]" & "/" & "['" & BRAND_CODE & "'_SLS_RTL]" & ")")
                '    .Columns.Add("'" & BRAND_CODE & "'_WCGSPCT", GetType(System.Decimal), _
                '               "IIF(['" & BRAND_CODE & "'_SLS_WSL]=0,0," & "100*['" & BRAND_CODE & "'_CGS]" & "/" & "['" & BRAND_CODE & "'_SLS_WSL]" & ")")
                'Next


                Dim KEY As String = summarySettings.Key
                If KEY.EndsWith("_RCGSPCT") Then
                    Dim CGS As String = Replace(KEY, "_RCGSPCT", "_CGS")
                    Dim SLS_RTL As String = Replace(KEY, "_RCGSPCT", "_SLS_RTL")
                    TOTALS.Add(CGS, 0)
                    TOTALS.Add(SLS_RTL, 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS(SLS_RTL) <> 0 Then CustomValue = 100 * TOTALS(CGS) / TOTALS(SLS_RTL)
                ElseIf KEY.EndsWith("_WGPPCT") Then
                    Dim WGP As String = Replace(KEY, "_WGPPCT", "_WGP")
                    Dim SLS_WSL As String = Replace(KEY, "_WGPPCT", "_SLS_WSL")
                    TOTALS.Add(WGP, 0)
                    TOTALS.Add(SLS_WSL, 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS(SLS_WSL) <> 0 Then CustomValue = 100 * TOTALS(WGP) / TOTALS(SLS_WSL)
                ElseIf KEY.EndsWith("_WCGSPCT") Then
                    Dim CGS As String = Replace(KEY, "_WCGSPCT", "_CGS")
                    Dim SLS_WSL As String = Replace(KEY, "_WCGSPCT", "_SLS_WSL")
                    TOTALS.Add(CGS, 0)
                    TOTALS.Add(SLS_WSL, 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS(SLS_WSL) <> 0 Then CustomValue = 100 * TOTALS(CGS) / TOTALS(SLS_WSL)
                ElseIf KEY.EndsWith("_WCGSPCT") Then
                    Dim CGS As String = Replace(KEY, "_WCGSPCT", "_CGS")
                    Dim SLS_WSL As String = Replace(KEY, "_WCGSPCT", "_SLS_WSL")
                    TOTALS.Add(CGS, 0)
                    TOTALS.Add(SLS_WSL, 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS(SLS_WSL) <> 0 Then CustomValue = 100 * TOTALS(CGS) / TOTALS(SLS_WSL)
                Else
                    Stop
                End If

            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Public Overrides Function CustomStringSummary_End( _
        ByVal summarySettings As UltraWinGrid.SummarySettings, _
        ByVal rows As UltraWinGrid.RowsCollection, _
        ByVal CustomValue As String, _
        ByVal grd As UltraWinGrid.UltraGrid) As String

        Select Case grd.Name
            Case "grdSATSCGS1"
                Dim KEY As String = summarySettings.Key
                CustomValue = "Totals"
            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Sub CustomSummary_Calculate_Totals( _
       ByVal rows As UltraWinGrid.RowsCollection, _
       ByRef TOTALS As Dictionary(Of String, Decimal), _
       ByVal KEY As String)

        For Each grow2 As UltraWinGrid.UltraGridRow In rows
            If grow2.IsGroupByRow Then
                Dim gbrow As UltraWinGrid.UltraGridGroupByRow = DirectCast(grow2, UltraWinGrid.UltraGridGroupByRow)
                CustomSummary_Calculate_Totals(gbrow.Rows, TOTALS, KEY)
            Else
                If KEY.EndsWith("_RCGSPCT") Then
                    Dim CGS As String = Replace(KEY, "_RCGSPCT", "_CGS")
                    Dim SLS_RTL As String = Replace(KEY, "_RCGSPCT", "_SLS_RTL")
                    TOTALS(CGS) += Val(grow2.Cells(CGS).Value & "")
                    TOTALS(SLS_RTL) += Val(grow2.Cells(SLS_RTL).Value & "")
                ElseIf KEY.EndsWith("_WGPPCT") Then
                    Dim WGP As String = Replace(KEY, "_WGPPCT", "_WGP")
                    Dim SLS_WSL As String = Replace(KEY, "_WGPPCT", "_SLS_WSL")
                    TOTALS(WGP) += Val(grow2.Cells(WGP).Value & "")
                    TOTALS(SLS_WSL) += Val(grow2.Cells(SLS_WSL).Value & "")
                ElseIf KEY.EndsWith("_WCGSPCT") Then
                    Dim CGS As String = Replace(KEY, "_WCGSPCT", "_CGS")
                    Dim SLS_WSL As String = Replace(KEY, "_WCGSPCT", "_SLS_WSL")
                    TOTALS(CGS) += Val(grow2.Cells(CGS).Value & "")
                    TOTALS(SLS_WSL) += Val(grow2.Cells(SLS_WSL).Value & "")
                ElseIf KEY.EndsWith("_WCGSPCT") Then
                    Dim CGS As String = Replace(KEY, "_WCGSPCT", "_CGS")
                    Dim SLS_WSL As String = Replace(KEY, "_WCGSPCT", "_SLS_WSL")
                    TOTALS(CGS) += Val(grow2.Cells(CGS).Value & "")
                    TOTALS(SLS_WSL) += Val(grow2.Cells(SLS_WSL).Value & "")
                ElseIf KEY = "TRADE_CLASS_CODE" Then
                    '  TOTALS(KEY) = "Totals"
                End If
            End If
        Next
    End Sub



End Class