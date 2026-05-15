Imports Infragistics.Win.UltraWinGrid

Public Class RSFRETLJ

    Dim IMPORT_NO As String

    Dim CUST_CODE As String
    Dim OPS_YYYYPP As String
    Dim OPS_YYYYWW As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GLTPARM1")
        With dst
            ASCMAIN1.sql = "SELECT RSTCUSTS.*,ARTCUST2.CUST_STORE_NAME FROM RSTCUSTS " &
                            "LEFT JOIN ARTCUST2 ON (RSTCUSTS.CUST_CODE=ARTCUST2.CUST_CODE AND RSTCUSTS.CUST_STORE_NO=ARTCUST2.CUST_STORE_NO)" &
                            " WHERE RSTCUSTS.CUST_CODE=:PARM1"
            Create_TDA(.Tables.Add, "RSTCUSTS", "**", 0, True, "V")

            ASCMAIN1.sql = "WITH CUR_YP(OPS_YYYYPP) AS " &
                            "  (SELECT CURR_YEAR || LPAD(CURR_PERIOD, 2, '0') OPS_YYYYPP FROM ASTPCTL1 " &
                            "  ), PRDS(OPS_YYYYPP) AS " &
                            "  (SELECT OPS_YYYYPP " &
                            "  FROM GLTPARM2 " &
                            "  WHERE OPS_YYYYPP BETWEEN TAPPRDA1( " &
                            "    (SELECT OPS_YYYYPP FROM CUR_YP " &
                            "    ), -35) " &
                            "  AND (SELECT OPS_YYYYPP FROM CUR_YP) " &
                            "  ) " &
                            "SELECT YP.OPS_YYYYPP, NVL(R5.CNT, 0) - NVL(RF1.NUM_CUSTS, 0) AUD_CUSTS, NVL(R5.AMT_SOLD, 0) - NVL(RF1.AMT_SOLD, 0) AUD_SLS, NVL(R1.CNT, 0) EDI_CUSTS, NVL(R1.AMT_SOLD, 0) EDI_SLS, NVL(RF1.NUM_CUSTS, 0) FF_CUSTS, NVL(RF1.AMT_SOLD, 0) FF_SLS " &
                            "FROM PRDS YP " &
                            "LEFT JOIN " &
                            "  (SELECT OPS_YYYYPP, COUNT( * ) CNT, SUM(AMT_SOLD_SUM) AMT_SOLD FROM RSSRETL5 GROUP BY OPS_YYYYPP " &
                            "  ) R5 ON (YP.OPS_YYYYPP = R5.OPS_YYYYPP) " &
                            "LEFT JOIN " &
                            "  (SELECT OPS_YYYYPP, COUNT( * ) CNT, SUM(AMT_SOLD_SUM) AMT_SOLD FROM RSSRETL1 GROUP BY OPS_YYYYPP " &
                            "  ) R1                 ON (YP.OPS_YYYYPP = R1.OPS_YYYYPP) " &
                            "LEFT JOIN RSTR5FF1 RF1 ON (YP.OPS_YYYYPP = RF1.OPS_YYYYPP) " &
                            "ORDER BY YP.OPS_YYYYPP desc"
            Create_TDA(.Tables.Add, "RSTSSUM1", "**", 0, False, "")

            ASCMAIN1.sql = "SELECT NVL(R5.CUST_CODE, R1.CUST_CODE) CUST_CODE, NVL(R5.AMT_SOLD_SUM, 0) - NVL(RF2.AMT_SOLD, 0) AUD_SLS, NVL(R1.AMT_SOLD_SUM, 0) EDI_SLS, NVL(RF2.AMT_SOLD, 0) FF_SLS " &
                            "FROM RSSRETL5 R5 " &
                            "FULL OUTER JOIN RSSRETL1 R1 ON (R5.CUST_CODE = R1.CUST_CODE AND R5.OPS_YYYYPP = R1.OPS_YYYYPP) " &
                            "LEFT JOIN RSTR5FF2 RF2      ON (RF2.CUST_CODE = R5.CUST_CODE AND RF2.OPS_YYYYPP = R5.OPS_YYYYPP) " &
                            "WHERE NVL(R5.OPS_YYYYPP, R1.OPS_YYYYPP) = :PARM1 " &
                            "ORDER BY CUST_CODE"
            Create_TDA(.Tables.Add, "RSTSSUM2", "**", 0, False, "V")

            ASCMAIN1.sql = "SELECT RSTRETL5.* FROM RSTRETL5 " &
                            " WHERE RSTRETL5.CUST_CODE=:PARM1 AND RSTRETL5.OPS_YYYYPP=:PARM2"
            Create_TDA(.Tables.Add, "RSTRETL5", "**", 0, True, "VV")

            ASCMAIN1.sql = "SELECT * from TATEVNT1 where TABLE_NAME = 'RSTSSUM1'"
            Create_TDA(.Tables.Add, "TATEVNT1", "**", 0, False)

            Create_TDA(.Tables.Add, "ARTCUST2", "*", 1, False)
            Create_TDA(.Tables.Add, "ICTCOLL1", "*", 0, False)
            Fill_Records("ICTCOLL1")

        End With

        grdRSTRETL5.DataSource = dst.Tables("RSTRETL5")
        grdRSTSSUM1.DataSource = dst.Tables("RSTSSUM1")
        grdRSTSSUM2.DataSource = dst.Tables("RSTSSUM2")

        grdTATEVNT1.DataSource = dst.Tables("TATEVNT1")
        'Fill_Records("TATEVNT1")
        'Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)



        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdRSTRETL5, grdRSTSSUM1, grdRSTSSUM2, grdTATEVNT1}
            grd.DisplayLayout.Override.ActiveRowAppearance.BackColor = System.Drawing.Color.LightGreen
            grd.DisplayLayout.Override.ActiveRowAppearance.ForeColor = System.Drawing.Color.Black

            grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

            If grd.Name = "grdRSTRETL5" Then
                grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            End If

            With grd.DisplayLayout.Bands(0)
                For Each c As UltraWinGrid.UltraGridColumn In .Columns
                    c.CellActivation = Activation.NoEdit
                    c.Header.Appearance.BackColor = System.Drawing.Color.White
                    c.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    c.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

                    If grd.Name = "grdRSTRETL5" Then
                        c.CellActivation = Activation.AllowEdit
                    End If

                    'If grd.Name = "grdSOTRMAF2" Then
                    '    If c.Key = "RA_QTY" Or c.Key = "RA_PRICE" Then
                    '        c.CellActivation = Activation.AllowEdit
                    '        c.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                    '    End If
                    '    If c.Key = "LAST_RECD_DATE" Or c.Key = "LAST_RECD_COST" Or c.Key = "VAR_LAST_COST" Or c.Key = "VAR_LAST_COST_PCT" Then
                    '        c.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                    '    End If
                    '    If c.Key = "FRAME_COST" Or c.Key = "VAR_FRAME_COST" Or c.Key = "VAR_FRAME_COST_PCT" Or c.Key = "EXT_FRAME_COST" Then
                    '        c.Header.Appearance.BackColor2 = System.Drawing.Color.LightPink
                    '    End If
                    '    If c.Key = "AMT_RTV" Or c.Key = "AMT_REC" Or c.Key = "AMT_CRD" Or c.Key = "QTY_REC" Or c.Key = "QTY_DMG" Or c.Key = "QTY_CRD" Or c.Key = "QTY_RTV" Or c.Key = "VAR_REC" Or c.Key = "VAR_RTV" Then
                    '        c.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                    '    End If
                    'End If
                Next
            End With
        Next



        Create_Summary(grdRSTRETL5, "AMT_SOLD", "Sum")

        Create_Summary(grdRSTSSUM2, "CUST_CODE", "Count")
        Create_Summary(grdRSTSSUM2, New String() {"AUD_SLS", "EDI_SLS", "FF_SLS"})

        'Dim sql_where As String = ""
        'sql_where = "OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -36) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "'"
        'Load_Drop_Down("OPS_YYYYPP", sql_where)


        'For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In cmbOPS_YYYYPP.Rows
        '    If row.Cells(0).Value.ToString = ASCMAIN1.CYP Then
        '        cmbOPS_YYYYPP.SelectedRow = row
        '        Exit For
        '    End If
        'Next
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("OPS_YYYYPP")

                If Absx1.txtFor("OPS_YYYYPP").Text < ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") Then
                    MsgBox("GL Period " & Absx1.txtFor("OPS_YYYYPP").Text & " has been closed, You may Load but Updating is Prohibited", MsgBoxStyle.OkOnly, "Audit Retail Entry")
                End If

                Validate_Code("CUST_CODE")

                If EMsg = "" Then
                    CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                    OPS_YYYYPP = Absx1.txtFor("OPS_YYYYPP").Text
                    OPS_YYYYWW = ASCDATA1.GetDataValue("SELECT YYYYWW FROM GLTPARM3 WHERE REL_WEEK=1 AND YYYYPP=:PARM1", "V", New Object() {OPS_YYYYPP})
                End If

            Case "Update"

                If Absx1.txtFor("OPS_YYYYPP").Text & "" < ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") Then
                    EMsg &= vbCr & "You cannot Update a Closed GL Period"
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
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode

                End With
            End With
        End If

        grdRSTRETL5.Visible = ScreenMode = True

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"RSTRETL5"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        LoadSummaryData()

        Fill_Records("TATEVNT1")
        Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        IMPORT_NO = ASCMAIN1.Next_Control_No("RSTIMPR1.IMPORT_NO")

        EnforceConstraints(False)
        Fill_Records("ARTCUST2", New Object() {CUST_CODE})
        Fill_Records("RSTRETL5", New Object() {CUST_CODE, OPS_YYYYPP})
        EnforceConstraints(True)

        grdRSTRETL5.Text = $"Retail Sales by Store for {CUST_CODE} for {OPS_YYYYPP}"

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()
        Update_Record_TDA("RSTRETL5")
        CommitTrans("Update Complete")
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Call Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "ITEM_CODE"
            '    If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
            '        Click_Command("Edit", e)
            '    End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "ITEM_CODE"
            '    Click_Command("Edit")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "ITEM_CODE"

        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)

        End Select
    End Sub

#End Region

#Region "Popup_Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdRSTRETL5, "SS", "Show Filter", "Show GroupBox")
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

        Select Case grd.Name
            'Case "grdSOTRMAF2"
            '    tlb_btn = DirectCast(tlb_pop.Tools("Style Multi-Color"), UltraWinToolbars.ButtonTool)
            '    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Item Status Inquiry"
            '    Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
            '    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            '    If rowICTITEM1 IsNot Nothing Then
            '        Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
            '    End If

        End Select
    End Sub

#End Region


#Region "RSTRETL5"

    Private Sub grdRSTRETL5_BeforeCellListDropDown(sender As System.Object, e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles grdRSTRETL5.BeforeCellListDropDown
        If Not (e.Cell.Row.IsAddRow Or e.Cell.Row.IsFilterRow) Then
            e.Cancel = True
        End If
    End Sub

    Private Sub grdRSTRETL5_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdRSTRETL5.AfterExitEditMode
        With grdRSTRETL5
            Select Case .ActiveCell.Column.Key
                Case "CUST_STORE_NO"
                    Dim CUST_STORE_NO As String = .ActiveCell.Text
                    If CUST_STORE_NO <> "" Then
                        .ActiveCell.Value = CUST_STORE_NO.PadLeft(6, "0").ToUpper()
                    End If
                Case "COLLECTION_CODE"
                    Dim COLLECTION_CODE As String = .ActiveCell.Text
                    If COLLECTION_CODE <> UCase(COLLECTION_CODE) Then
                        .ActiveCell.Value = COLLECTION_CODE.ToUpper()
                    End If

            End Select
        End With
    End Sub

    Private Sub grdRSTRETL5_BeforeRowUpdate(sender As System.Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdRSTRETL5.BeforeRowUpdate
        Dim validMsg = ""
        Select Case isRowValid(e.Row)
            Case RowValidity.Valid
                grdRSTRETL5.DisplayLayout.RowScrollRegions(0).ScrollRowIntoView(e.Row)
            Case RowValidity.InvalidStore
                validMsg = "Invalid Store"
            Case RowValidity.InvalidCollection
                validMsg = "Invalid Collection"
            Case RowValidity.InvalidStoreAndCollection
                validMsg = "Invalid Store and Collection"
        End Select

        If validMsg <> "" Then
            e.Cancel = True
            MsgBox(validMsg)
        End If

        If Not e.Cancel Then
            With grdRSTRETL5
                If .ActiveRow.Cells("IMPORT_NO").Text = "" Then
                    .ActiveRow.Cells("IMPORT_NO").Value = IMPORT_NO
                    .ActiveRow.Cells("CUST_CODE").Value = CUST_CODE

                    e.Row.Cells("OPS_YYYYWW").Value = OPS_YYYYWW
                    e.Row.Cells("OPS_YYYYPP").Value = OPS_YYYYPP
                End If
            End With

        End If
    End Sub

    Private Sub grdRSTRETL5_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdRSTRETL5.ClickCellButton
        Dim COLUMN_NAME As String = e.Cell.Column.Key

        Dim sql_where As String = ""

        Select Case COLUMN_NAME
            Case "CUST_STORE_NO"
                sql_where = $"CUST_CODE='{CUST_CODE}'"

            Case "COLLECTION_CODE"
                sql_where = ""
        End Select

        grdClickCellButton(grdRSTRETL5, sql_where)


        'Dim SQL As String = ""

        'If e.Cell.Column.CellActivation = UltraWinGrid.Activation.NoEdit And Not e.Cell.Row.IsAddRow Then
        '    Exit Sub
        'End If

        'Dim TABLE_NAME As String = ""
        'Select Case COLUMN_NAME
        '    Case "CUST_STORE_NO"
        '        TABLE_NAME = "ARTCUST2"
        '        ' If you need to limit the select then add a where clause
        '   '     SQL = String.Format("SELECT CUST_STORE_NO, CUST_STORE_NAME FROM ARTCUST2 WHERE CUST_CODE='{0}'", CUST_CODE)

        '    Case "COLLECTION_CODE"
        '        TABLE_NAME = "ICTCOLL1"
        '        '   SQL = String.Format("SELECT COLLECTION_CODE, COLLECTION_NAME FROM ICTCOLL1")
        'End Select

        ''ASCMAIN1.CodeSelector.SQL = SQL
        'ASCMAIN1.CodeSelector.UseDataFromTable = dst.Tables(TABLE_NAME)
        'ASCMAIN1.CodeSelector.VIEW_NAME = COLUMN_NAME
        'ASCMAIN1.CodeSelector.MultipleSelections = False
        'ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = e.Cell.Text

        'Dim F As New ASFCODE1
        'F.ShowDialog()
        'F.Dispose()

        'If ASCMAIN1.CodeSelector.Selections <> 0 Then

        '    Select Case COLUMN_NAME
        '        Case "CUST_STORE_NO", "COLLECTION_CODE"
        '            grdRSTRETL5.ActiveRow.Cells(COLUMN_NAME).Value = ASCMAIN1.CodeSelector.SelectedRows(0).Item(0)
        '    End Select
        'End If
    End Sub

#End Region

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME

            Case "CUST_CODE"
                'sql_where = "CUST_STATUS = 'A' AND TRADE_CLASS_CODE IN ('IND','NAT')"

            Case "OPS_YYYYPP"
                sql_where = "OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -36) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "'"
        End Select
    End Sub

    Private Sub grdRSTSSUM1_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdRSTSSUM1.DoubleClickRow

        Dim YP As String = grdRSTSSUM1.ActiveRow.Cells("OPS_YYYYPP").Value & ""
        If YP < ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") Then
            MsgBox("You cannot Process Flood Fill for a Closed GL Period " & YP, MsgBoxStyle.OkOnly, "Audit Retail Entry")
        Else
            Dim floodFillExists As Boolean = False
            If e.Row.Cells("FF_SLS").Value > 0 Then
                floodFillExists = True
            End If

            If MsgBox(IIf(floodFillExists,
                          $"Flood Fill Exists for {YP} - continuing will zero out Flood Fill.",
                          $"Flood Fill Does NOT Exist for {YP} - continuing will establish Flood Fill.") _
                          & vbCrLf & vbCrLf & "Continue?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

            ASCMAIN1.Progress("Processing Flood Fill...")

            DATETIME_STAMP = Now + ASCMAIN1.NowTSD

            If floodFillExists Then
                RemoveFloodFill(e.Row.Cells("OPS_YYYYPP").Value)
                TAC.TACMAIN1.Record_Event("RSTSSUM1", YP, DATETIME_STAMP, ASCMAIN1.USER_ID, "ZEROFF", "Zero Flood Fill")
            Else
                PerformFloodFill(e.Row.Cells("OPS_YYYYPP").Value)
                TAC.TACMAIN1.Record_Event("RSTSSUM1", YP, DATETIME_STAMP, ASCMAIN1.USER_ID, "UPDTFF", "Update Flood Fill")
            End If

            Fill_Records("TATEVNT1")
            Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)

            ASCMAIN1.Progress("")

            LoadSummaryData(e.Row.Cells("OPS_YYYYPP").Value)
            grdRSTSSUM2.Visible = True
            grdRSTSSUM2.Text = $"Retail Sales by Customer for Period {YP}"
        End If
    End Sub

    Private Sub RemoveFloodFill(ByVal period As String)
        ASCDATA1.ExecuteSP("RSPFFDEL", "VV", New Object() {period, ASCMAIN1.USER_ID}, New String() {"P_OPS_YYYYPP", "P_INIT_OPER"})
    End Sub

    Private Sub PerformFloodFill(ByVal period As String)
        ASCDATA1.ExecuteSP("RSPRTLFF", "VV", New Object() {period, ASCMAIN1.USER_ID}, New String() {"P_OPS_YYYYPP", "P_INIT_OPER"})
    End Sub

    Private Sub LoadSummaryData(Optional ByVal period As String = "")
        Fill_Records("RSTSSUM1", , True)

        Dim selectedPeriod As String = ""
        If grdRSTSSUM1.ActiveRow IsNot Nothing Then
            selectedPeriod = grdRSTSSUM1.ActiveRow.Cells("OPS_YYYYPP").Value
        End If

        If selectedPeriod = "" Then selectedPeriod = period

        Fill_Records("RSTSSUM2", New Object() {selectedPeriod}, True)
        grdRSTSSUM2.Text = "Retail Sales by Customer"
    End Sub


    Private Sub grdRSTSSUM2_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdRSTSSUM2.InitializeRow
        If e.Row.Cells("FF_SLS").Value > 0 Then
            e.Row.Appearance.BackColor = Color.PaleGoldenrod
        End If
    End Sub

    Private Function isRowValid(ByVal row As UltraGridRow) As RowValidity

        Dim CUST_STORE_NO As String = row.Cells("CUST_STORE_NO").Value & ""
        Dim validStore As Boolean = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO}) IsNot Nothing

        Dim COLLECTION_CODE As String = row.Cells("COLLECTION_CODE").Value & ""
        Dim validCollection As Boolean = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE) IsNot Nothing
        If validCollection Then
            If COLLECTION_CODE <> COLLECTION_CODE.ToUpper Then
                validCollection = False
            End If
        End If

        If validStore And validCollection Then Return RowValidity.Valid
        If validStore And Not validCollection Then Return RowValidity.InvalidCollection
        If Not validStore And validCollection Then Return RowValidity.InvalidStore

        Return RowValidity.InvalidStoreAndCollection
    End Function

    Enum RowValidity
        Valid
        InvalidStore
        InvalidCollection
        InvalidStoreAndCollection
    End Enum

End Class