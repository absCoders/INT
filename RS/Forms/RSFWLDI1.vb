Public Class RSFWLDI1

    Dim RYP0 As String
    Dim RYP1 As String
    Dim RYW As String
    Dim RSTWLDI1 As String
    Dim ICTITEM1 As String
    Dim EDT852TC As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")

        With dst
            Create_RSTWLDI1()

            Dim sqlRQ As String = "" _
            & ", SUM (DECODE(ICTITEM1.ITEM_CATGY_CODE,'?',RSTWLDI1.RQ,0)) RQ?" _
            & ", SUM (DECODE(ICTITEM1.ITEM_CATGY_CODE,'?',RSTWLDI1.RQ,0) * ICTITEM1.ITEM_PRICE) RQWSL?"

            Dim sqlWQ As String = "" _
            & ", SUM (DECODE(ICTITEM1.ITEM_CATGY_CODE,'?',RSTWLDI1.WQ,0)) WQ?" _
            & ", SUM (DECODE(ICTITEM1.ITEM_CATGY_CODE,'?',RSTWLDI1.WQ,0) * ICTITEM1.ITEM_PRICE) WQWSL?"

            ASCMAIN1.sql = "Select ICTITEM1.COLLECTION_CODE " _
            & ", DECODE(ICTITEM1.ITEM_CATGY_CODE,'C','A','E','A','N','A','P','A','I') STATUS" _
            & ", ICTCOLL1.HC_CODE" _
            & ", SUM(RSTWLDI1.RS) RS, SUM (RSTWLDI1.RU * ICTITEM1.ITEM_PRICE) RUWSL" _
            & ", SUM (RSTWLDI1.RU) RU" _
            & ", SUM (RSTWLDI1.RQ) RQ" _
            & ", SUM (RSTWLDI1.RQ * ICTITEM1.ITEM_RETAIL_PRICE) RQRTL" _
            & ", SUM (RSTWLDI1.RQ * ICTITEM1.ITEM_PRICE) RQWSL" _
            & Replace(sqlRQ, "?", "C") _
            & Replace(sqlRQ, "?", "E") _
            & Replace(sqlRQ, "?", "P") _
            & Replace(sqlRQ, "?", "N") _
            & Replace(sqlRQ, "?", "I") _
            & ", SUM (RSTWLDI1.WQ) WQ" _
            & ", SUM (RSTWLDI1.WQ * ICTITEM1.ITEM_RETAIL_PRICE) WQRTL" _
            & ", SUM (RSTWLDI1.WQ * ICTITEM1.ITEM_PRICE) WQWSL" _
            & Replace(sqlWQ, "?", "C") _
            & Replace(sqlWQ, "?", "E") _
            & Replace(sqlWQ, "?", "P") _
            & Replace(sqlWQ, "?", "N") _
            & Replace(sqlWQ, "?", "I") _
            & " from " _
            & RSTWLDI1 & " RSTWLDI1, " & ICTITEM1 & " ICTITEM1, ICTCOLL1" _
            & " where ICTITEM1.ITEM_CODE = RSTWLDI1.ITEM_CODE" _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" _
            & " group by " _
            & "  ICTITEM1.COLLECTION_CODE" _
            & ", DECODE(ICTITEM1.ITEM_CATGY_CODE,'C','A','E','A','N','A','P','A','I')" _
            & ", ICTCOLL1.HC_CODE"
            Create_TDA(.Tables.Add, "RSTWLDI0", "**", 0, False, , 2)
            .Tables("RSTWLDI0").Columns.Add("R_SELLTHRU", GetType(System.Decimal), "IIF(RU+RQ=0,0,100*RU/(RU+RQ))")
            .Tables("RSTWLDI0").Columns.Add("RQWSL_PTL", GetType(System.Decimal), "IIF(RU+RQ=0,0,100*RU/(RU+RQ))")

            For Each GT As String In New String() {"R", "W", "T"}
                For Each ITEM_CATGY_CODE As String In New String() {"", "C", "E", "P", "N", "I"}
                    If GT = "T" Then
                        .Tables("RSTWLDI0").Columns.Add("TQ" & ITEM_CATGY_CODE, GetType(System.Decimal), "RQ" & ITEM_CATGY_CODE & "+WQ" & ITEM_CATGY_CODE)
                        If ITEM_CATGY_CODE = "" Then
                            .Tables("RSTWLDI0").Columns.Add("TQRTL" & ITEM_CATGY_CODE, GetType(System.Decimal), "RQRTL" & ITEM_CATGY_CODE & "+WQRTL" & ITEM_CATGY_CODE)
                        End If
                        .Tables("RSTWLDI0").Columns.Add("TQWSL" & ITEM_CATGY_CODE, GetType(System.Decimal), "RQWSL" & ITEM_CATGY_CODE & "+WQWSL" & ITEM_CATGY_CODE)
                    End If

                    If ITEM_CATGY_CODE = "" Then
                        COLUMN_NAME = GT & "QWSL" & ITEM_CATGY_CODE & "_PCT_TTL"
                        .Tables("RSTWLDI0").Columns.Add(COLUMN_NAME, GetType(System.Decimal))

                    Else
                        COLUMN_NAME = GT & "Q" & ITEM_CATGY_CODE & "_PCT"
                        Dim COLUMN_NAME1 As String = GT & "Q" & ITEM_CATGY_CODE
                        Dim COLUMN_NAME2 As String = GT & "Q"
                        .Tables("RSTWLDI0").Columns.Add(COLUMN_NAME, GetType(System.Decimal), "IIF(" & COLUMN_NAME2 & "=0,0,100*" & COLUMN_NAME1 & " / " & COLUMN_NAME2 & ")")

                    End If
                Next
            Next

            ASCMAIN1.sql = "Select * from ICTCATG1"
            Create_TDA(.Tables.Add, "ICTCATG1", "**", 0, True, , 1)

            ASCMAIN1.sql = "Select CUST_CODE" _
            & ", MAX (OPS_YYYYWW) OPS_YYYYWW, '0' SEL from EDT852T1 group by CUST_CODE"
            EDT852TC = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & EDT852TC & " Add Primary Key (CUST_CODE)")
            Create_TDA(.Tables.Add("EDT852TC"), EDT852TC, "**", 0, True, , 2)

            With .Tables.Add("RSTWLDIQ")
                .Columns.Add("RECAP_DESC")
                .Columns.Add("RECAP_VALUE", GetType(System.Decimal))
                .Columns.Add("RECAP_PCT", GetType(System.Decimal))
            End With
        End With

        grdRSTWLDI0.DataSource = dst.Tables("RSTWLDI0")
        grdRSTWLDIQ.DataSource = dst.Tables("RSTWLDIQ")

        Fill_Records("EDT852TC")
        Fill_Records("ICTCATG1")

        With grdRSTWLDI0.DisplayLayout.Bands(0)
            .ColHeaderLines = 2

            For Each COLUMN_NAME As String In New String() _
            {"STATUS", "HC_CODE", "COLLECTION_CODE"}
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.Beige
            Next
            .Columns("COLLECTION_CODE").Header.Fixed = True
        End With

        Dim G As UltraWinGrid.UltraGridGroup

        With grdRSTWLDI0.DisplayLayout.Bands(0)
            G = .Groups.Add("CODES")
            G.Header.Caption = "Collection Activity"
            G.Header.Appearance.BackColor = System.Drawing.Color.LightGreen
            .Columns("COLLECTION_CODE").Group = G

            Dim CIx As Integer = 0
            For Each COLUMN_NAME As String In New String() {"RS", "RUWSL", "RU", "R_SELLTHRU"}
                With .Columns(COLUMN_NAME)
                    .Group = G
                    If COLUMN_NAME = "R_SELLTHRU" Then
                        .Format = "##0.0"
                        Create_Summary(grdRSTWLDI0, COLUMN_NAME, "Custom")
                    Else
                        .Format = "###,##0"
                        Create_Summary(grdRSTWLDI0, COLUMN_NAME)
                    End If
                    .Header.Caption = vbCrLf & New String() {"$Sell Thru", "#ST@Wsl", "#Sell Thru", "Sell Thru%"}(CIx)
                    .Header.Appearance.BackColor = System.Drawing.Color.Lime

                End With
                CIx += 1
            Next

            For Each GT As String In New String() {"R", "W", "T"}
                Dim gi As Integer = InStr("RWT", GT)
                G = .Groups.Add(GT)
                G.Header.Appearance.TextHAlign = HAlign.Center
                G.Header.Caption = New String() {"Retailer", "In Stock", "Total"}(gi - 1)
                G.Header.Appearance.BackColor = New System.Drawing.Color() _
                    {System.Drawing.Color.Violet, System.Drawing.Color.Orange, System.Drawing.Color.DodgerBlue}(gi - 1)


                For Each ITEM_CATGY_CODE As String In New String() {"", "C", "E", "P", "N", "I"}
                    Dim CI As Integer = InStr("CEPNI", ITEM_CATGY_CODE)
                    If ITEM_CATGY_CODE = "" Then CI = 0
                    Dim clr As System.Drawing.Color = New System.Drawing.Color() _
                        {System.Drawing.Color.LightBlue, _
                         System.Drawing.Color.Yellow, _
                         System.Drawing.Color.Orange, _
                         System.Drawing.Color.LimeGreen, _
                         System.Drawing.Color.CornflowerBlue, _
                         System.Drawing.Color.Pink}(CI)

                    COLUMN_NAME = GT & "Q" & ITEM_CATGY_CODE
                    .Columns(COLUMN_NAME).Group = G
                    .Columns(COLUMN_NAME).Format = "###,##0"
                    Dim ITEM_CATGY_DESC As String = ""
                    If ITEM_CATGY_CODE <> "" Then
                        Dim rowICTCATG1 As DataRow = dst.Tables("ICTCATG1").Rows.Find(ITEM_CATGY_CODE)
                        ITEM_CATGY_DESC = rowICTCATG1.Item("ITEM_CATGY_DESC")
                    End If

                    .Columns(COLUMN_NAME).Header.Caption = ITEM_CATGY_DESC & vbCrLf & "#OH"
                    Create_Summary(grdRSTWLDI0, COLUMN_NAME)

                    .Columns(COLUMN_NAME).Header.Appearance.BackColor = clr

                    If ITEM_CATGY_CODE = "" Then
                        COLUMN_NAME = GT & "QRTL" & ITEM_CATGY_CODE
                        .Columns(COLUMN_NAME).Group = G
                        .Columns(COLUMN_NAME).Format = "###,##0"
                        .Columns(COLUMN_NAME).Header.Caption = ITEM_CATGY_DESC & vbCrLf & "$OH"
                        Create_Summary(grdRSTWLDI0, COLUMN_NAME)
                        .Columns(COLUMN_NAME).Header.Appearance.BackColor = clr
                    End If

                    COLUMN_NAME = GT & "QWSL" & ITEM_CATGY_CODE
                    .Columns(COLUMN_NAME).Group = G
                    .Columns(COLUMN_NAME).Format = "###,##0"
                    .Columns(COLUMN_NAME).Header.Caption = ITEM_CATGY_DESC & vbCrLf & "#OH@Wsl"
                    Create_Summary(grdRSTWLDI0, COLUMN_NAME)
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor = clr

                    If ITEM_CATGY_CODE = "" Then
                        COLUMN_NAME = GT & "QWSL" & ITEM_CATGY_CODE & "_PCT_TTL"
                        .Columns(COLUMN_NAME).Group = G
                        .Columns(COLUMN_NAME).Format = "##0.0"
                        .Columns(COLUMN_NAME).Header.Caption = ITEM_CATGY_DESC & vbCrLf & "@Wsl %Ttl"
                        Create_Summary(grdRSTWLDI0, COLUMN_NAME)
                        .Columns(COLUMN_NAME).Header.Appearance.BackColor = clr

                    Else
                        COLUMN_NAME = GT & "Q" & ITEM_CATGY_CODE & "_PCT"
                        .Columns(COLUMN_NAME).Group = G
                        .Columns(COLUMN_NAME).Format = "##0.0"
                        .Columns(COLUMN_NAME).Header.Caption = ITEM_CATGY_DESC & vbCrLf & "#OH%"
                        Create_Summary(grdRSTWLDI0, COLUMN_NAME, "Custom")
                        .Columns(COLUMN_NAME).Header.Appearance.BackColor = clr

                    End If
                Next
            Next

            .SummaryFooterCaption = "Total [GROUPBYROWVALUE]"

        End With

        Absx1.cmbFor("RYP0").Value = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -11)
        Absx1.cmbFor("RYP1").Value = ASCMAIN1.CYP

        'ASCMAIN1.Add_Value_List(grdRSTWLDI0, "ITEM_CATGY_CODE")
        'ASCMAIN1.Add_Value_List(grdRSTWLDI0, "MATL_CODE")
        grdRSTWLDI0.DisplayLayout.GroupByBox.Hidden = True
        'Show_Filter(grdRSTWLDI0, True)

        grdRSTWLDI0.DisplayLayout.Bands(0).Columns("HC_CODE").SortComparer = New srtComparer

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                RYP0 = Absx1.cmbFor("RYP0").Value
                RYP1 = Absx1.cmbFor("RYP1").Value

                If RYP0 > RYP1 Then
                    EMsg &= vbCr & "Starting Period may not be later than Ending Period"
                End If

                If EMsg = "" Then
                    Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", RYP1)
                    ASCMAIN1.sql = "Select Max (YYYYWW) from GLTPARM3 where YYYYPP = '" & RYP1 & "'"
                    RYW = ASCDATA1.GetDataValue
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

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Scope").Visible = Not ScreenMode
                .Groups("Options").Visible = ScreenMode
                .Groups("OH @Retailers").Visible = ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdRSTWLDI0.Visible = tf
        Set_Read_Only_for_ctl(chkHISTCAT, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        dst.Tables("RSTWLDI0").Rows.Clear()
        dst.Tables("RSTWLDIQ").Rows.Clear()
        EnforceConstraints(True)
        grdRSTWLDI0.DisplayLayout.Rows.ColumnFilters.ClearAllFilters()
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Save_Header_Fields(UltraGroupBox1)

        grdRSTWLDI0.DisplayLayout.Bands(0).SortedColumns.Clear()

        EnforceConstraints(False)
        Create_RSTWLDI1()
        Fill_Records("RSTWLDI0")
        EnforceConstraints(True)

        For Each GT As String In New String() {"R", "W", "T"}
            COLUMN_NAME = GT & "QWSL" & "" & "_PCT_TTL"
            Dim PCT_TTL As String = "0"
            Dim TTL As Decimal = Val(dst.Tables("RSTWLDI0").Compute("SUM(" & GT & "QWSL)", "") & "")
            If TTL <> 0 Then
                PCT_TTL = "100 * " & GT & "QWSL" & " / " & CStr(TTL)
            End If
            dst.Tables("RSTWLDI0").Columns(COLUMN_NAME).Expression = PCT_TTL
        Next

        ASCMAIN1.Progress("Now Setting Up Screen")

        With grdRSTWLDI0.DisplayLayout
            .Rows.ColumnFilters.ClearAllFilters()
            With .Bands(0)
                .SortedColumns.Clear()
                .SortedColumns.Add("STATUS", False, True)
                .SortedColumns.Add("HC_CODE", True, True)
                .SortedColumns.Add("COLLECTION_CODE", False, False)
            End With
        End With

        Show_Columns()
        Dim MOs As String = Format(ASCMAIN1.Period_Diff(RYP0, RYP1) + 1, "00")
        grdRSTWLDI0.Text = "World Inventory Sell-In / Sell-Thru for the " & MOs & " Months from " & Absx1.cmbFor("RYP0").Text & " to " & Absx1.cmbFor("RYP1").Text

        'grdRSTWLDI1.Rows.Refresh(UltraWinGrid.RefreshRow.RefreshDisplay)

        grdRSTWLDI0.Rows.CollapseAll(True)

        For Each grow As UltraWinGrid.UltraGridRow In grdRSTWLDI0.Rows
            grow.Expanded = True
        Next

        dst.Tables("RSTWLDIQ").Rows.Clear()

        Dim INACTIVE As Decimal = Val(dst.Tables("RSTWLDI0").Compute("SUM(RQWSLI)", "") & "")
        Dim PHASE_OUT As Decimal = Val(dst.Tables("RSTWLDI0").Compute("SUM(RQWSLP)", "") & "")
        Dim TOTAL As Decimal = Val(dst.Tables("RSTWLDI0").Compute("SUM(RQWSL)", "") & "")

        Dim ACTIVE As Decimal = Val(dst.Tables("RSTWLDI0").Compute("SUM(RQWSL)", "STATUS = 'A'") & "")

        Dim UPCT As Decimal = 0 : If TOTAL <> 0 Then UPCT = 100 * (INACTIVE + PHASE_OUT) / TOTAL ' IIf(TOTAL = 0, 0, 100 * (INACTIVE + PHASE_OUT) / TOTAL)
        Dim PPCT As Decimal = 0 : If TOTAL <> 0 Then PPCT = 100 * (ACTIVE - PHASE_OUT) / TOTAL ' (TOTAL = 0, 0, 100 * (ACTIVE - PHASE_OUT) / TOTAL)

        dst.Tables("RSTWLDIQ").Rows.Add(New Object() {"Unproductive", (INACTIVE + PHASE_OUT) / 1000, UPCT})
        dst.Tables("RSTWLDIQ").Rows.Add(New Object() {"Productive", (ACTIVE - PHASE_OUT) / 1000, PPCT})

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Update_Record()
        Call BeginTrans()
        Call CommitTrans("Update Complete")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdRSTWLDI0, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                'Case "grdSATCSLS1"

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Item Status Inquiry"
            '    Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
            '    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            'Case "OPS_YYYYPP"
            '    If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
            '        Call Click_Command("Load", e)
            '    End If
        End Select

    End Sub

    Public Overrides Sub chk_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.chk_CheckedChanged(sender, e)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Setting Column Visibility")

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(DirectCast(sender, Control))
        If COLUMN_NAME.StartsWith("G") Then
            Show_Groups()
        End If
        If COLUMN_NAME.StartsWith("C") Then
            Show_Columns()
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub
#End Region

    Sub Create_RSTWLDI1()

        Dim sql_filter As String = ""
        Dim sql_joins As String = ""
        Dim sql_tables As String = ""

        Dim sqlRSTWLDI1 As String = "Select ITEM_CODE" _
        & ", SUM (RU) RU, SUM (RS) RS, SUM (RQ) RQ, SUM (WU) WU, SUM (WS) WS, SUM (WQ) WQ" _
        & " from (" _
        & " Select X.ITEM_CODE" _
        & ", 0 RU, 0 RS, 0 RQ, SUM (X.ORDR_QTY_SHIP) WU, SUM (X.ORDR_QTY_SHIP * X.ORDR_UNIT_PRICE) WS, 0 WQ" _
        & " from SOTINVH2 X" & sql_tables _
        & " where X.ORDR_YYYYPP_UPDATED >= 'TYP000' and X.ORDR_YYYYPP_UPDATED <= 'TYP001'" _
        & IIf(chkGross.Checked, " and X.INV_TYPE = 'I'", "") _
        & sql_joins & sql_filter _
        & " group by X.ITEM_CODE" _
        & " union " _
        & " Select X.ITEM_CODE" _
        & ", 0 RU, 0 RS, 0 RQ, 0 WU, 0 WS, SUM (X.WHSE_QTY_ON_HAND) WQ" _
        & " from " & IIf(RYP1 = ASCMAIN1.CYP, "ICTSTAT2", "ICTSTAT5") & " X" & sql_tables _
        & IIf(RYP1 = ASCMAIN1.CYP, "", " where X.OPS_YYYYPP = 'TYP001'") _
        & sql_joins & sql_filter _
        & " group by X.ITEM_CODE" _
        & " union " _
        & " Select X.ITEM_CODE" _
        & ", SUM (X.QTY_SOLD) RU, SUM (X.AMT_SOLD) RS, SUM (DECODE(X.OPS_YYYYWW,'TYW000',X.QTY_EOW,0)) RQ, 0 WU, 0 WS, 0 WQ" _
        & " from RSTRETL1 X" & sql_tables _
        & " where X.OPS_YYYYPP >= 'TYP000' and X.OPS_YYYYPP <= 'TYP001'" _
        & sql_joins & sql_filter _
        & " group by X.ITEM_CODE" _
        & ") group by ITEM_CODE"

        If RYP1 = "" Then
            ASCMAIN1.sql = sqlRSTWLDI1
            RSTWLDI1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & RSTWLDI1 & " Add Primary Key (ITEM_CODE)")
            ICTITEM1 = RSCMAIN1.Get_ICTITEM1_Hist_CATGY(RYP1, chkHISTCAT.Checked)
        Else
            RSCMAIN1.Get_ICTITEM1_Hist_CATGY(RYP1, chkHISTCAT.Checked, ICTITEM1)
            ASCDATA1.ExecuteSQL("Truncate Table " & RSTWLDI1)
            ASCMAIN1.sql = sqlRSTWLDI1
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "TYP000", RYP0)
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "TYP001", RYP1)
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "TYW000", RYW)
            ASCDATA1.ExecuteSQL("Insert into " & RSTWLDI1 & " " & ASCMAIN1.sql)
        End If
    End Sub

    Private Sub grdRSTWLDI0_InitializeGroupByRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeGroupByRowEventArgs) Handles grdRSTWLDI0.InitializeGroupByRow
        e.Row.Description = e.Row.ValueAsDisplayText & ""
        If e.Row.Column.Key = "STATUS" Then
            If e.Row.Value = "A" Then e.Row.Description = "Active"
            If e.Row.Value = "I" Then e.Row.Description = "Inactive"
            e.Row.Appearance.BackColor = Drawing.Color.LightBlue
        ElseIf e.Row.Column.Key = "HC_CODE" Then
            e.Row.Appearance.BackColor = Drawing.Color.Linen
        End If
    End Sub

    Private Sub grdRSTWLDI1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdRSTWLDI0.InitializeRow

    End Sub

    Sub Show_Groups()
        For Each G As String In New String() {"R", "W", "T"}
            grdRSTWLDI0.DisplayLayout.Bands(0).Groups(G).Hidden = Not Absx1.chkFor("G" & G).Checked
        Next
    End Sub

    Sub Show_Columns()
        For Each C As String In New String() {"*", "C", "E", "P", "N", "I"}
            For Each G As String In New String() {"R", "W", "T"}
                Dim ITEM_CATGY_CODE As String = C
                If C = "*" Then ITEM_CATGY_CODE = ""
                grdRSTWLDI0.DisplayLayout.Bands(0).Columns(G & "Q" & ITEM_CATGY_CODE).Hidden = Not Absx1.chkFor("C" & C).Checked
                If ITEM_CATGY_CODE = "" Then
                    grdRSTWLDI0.DisplayLayout.Bands(0).Columns(G & "QRTL" & ITEM_CATGY_CODE).Hidden = Not Absx1.chkFor("C" & C).Checked
                End If
                grdRSTWLDI0.DisplayLayout.Bands(0).Columns(G & "QWSL" & ITEM_CATGY_CODE).Hidden = Not Absx1.chkFor("C" & C).Checked
                If ITEM_CATGY_CODE = "" Then
                    grdRSTWLDI0.DisplayLayout.Bands(0).Columns(G & "QWSL" & ITEM_CATGY_CODE & "_PCT_TTL").Hidden = Not Absx1.chkFor("C" & C).Checked
                Else
                    grdRSTWLDI0.DisplayLayout.Bands(0).Columns(G & "Q" & ITEM_CATGY_CODE & "_PCT").Hidden = Not Absx1.chkFor("C" & C).Checked
                End If
            Next
        Next
    End Sub

    Public Overrides Function CustomSummary_End( _
    ByVal summarySettings As UltraWinGrid.SummarySettings, _
    ByVal rows As UltraWinGrid.RowsCollection, _
    ByVal CustomValue As Double, _
    ByVal grd As UltraWinGrid.UltraGrid) As Double

        CustomValue = 0
        Dim TOTALS As New Dictionary(Of String, Decimal)

        Select Case grd.Name
            Case "grdRSTWLDI0"
                Dim KEY As String = summarySettings.Key
                If KEY.EndsWith("_SELLTHRU") Then
                    TOTALS.Add("U", 0)
                    TOTALS.Add("Q", 0)
                    grdRSTWLDI0_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS("U") + TOTALS("Q") <> 0 Then CustomValue = 100 * TOTALS("U") / (TOTALS("U") + TOTALS("Q"))

                ElseIf KEY.EndsWith("_PCT") Then
                    TOTALS.Add("Q", 0)
                    TOTALS.Add("QX", 0)
                    grdRSTWLDI0_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS("Q") <> 0 Then CustomValue = 100 * TOTALS("QX") / TOTALS("Q")
                End If

            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Sub grdRSTWLDI0_Calculate_Totals( _
        ByVal rows As UltraWinGrid.RowsCollection, _
        ByRef TOTALS As Dictionary(Of String, Decimal), _
        ByVal KEY As String)

        For Each grow2 As UltraWinGrid.UltraGridRow In rows
            If grow2.IsGroupByRow Then
                Dim gbrow As UltraWinGrid.UltraGridGroupByRow = DirectCast(grow2, UltraWinGrid.UltraGridGroupByRow)
                grdRSTWLDI0_Calculate_Totals(gbrow.Rows, TOTALS, KEY)
            Else
                If KEY.EndsWith("_SELLTHRU") Then
                    TOTALS("U") += Val(grow2.Cells(Mid(KEY, 1, 1) & "U").Value & "")
                    TOTALS("Q") += Val(grow2.Cells(Mid(KEY, 1, 1) & "Q").Value & "")
                ElseIf KEY.EndsWith("_PCT") Then
                    TOTALS("Q") += Val(grow2.Cells(Mid(KEY, 1, 2)).Value & "")
                    TOTALS("QX") += Val(grow2.Cells(Mid(KEY, 1, 3)).Value & "")
                End If
            End If
        Next
    End Sub
End Class

Public Class srtComparer
    Implements IComparer

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare

        Dim xCell As UltraWinGrid.UltraGridCell = DirectCast(x, UltraWinGrid.UltraGridCell)
        Dim yCell As UltraWinGrid.UltraGridCell = DirectCast(y, UltraWinGrid.UltraGridCell)

        ' you can assume that xCell and yCell are 2 cells in the column using this sort comparer
        ' return 0 if the values in xCell and yCell are equal
        ' return -1 if x < y
        ' retyrb +1 if x > y

        If xCell.Value & "" = yCell.Value & "" Then ' Or Not xCell.Row.IsGroupByRow
            Return 0
        Else
            Dim tbl As DataTable = DirectCast(xCell.Column.Band.Layout.Grid.DataSource, DataTable)

            Dim STATUS As String = xCell.Row.Cells("STATUS").Value & ""
            Dim HC_CODE As String = xCell.Row.Cells("HC_CODE").Value & ""
            Dim xv As Decimal = Val(tbl.Compute("SUM(RS)", "STATUS = '" & STATUS & "' and HC_CODE = '" & HC_CODE & "'") & "")

            STATUS = yCell.Row.Cells("STATUS").Value & ""
            HC_CODE = yCell.Row.Cells("HC_CODE").Value & ""
            Dim yv As Decimal = Val(tbl.Compute("SUM(RS)", "STATUS = '" & STATUS & "' and HC_CODE = '" & HC_CODE & "'") & "")


            Dim COLUMN_NAME As String = xCell.Column.Key

            If xv = yv Or xCell.Value & "" = yCell.Value & "" Then
                Return 0
            Else
                Return IIf(xv < yv, -1, 1)
            End If
        End If

    End Function
End Class