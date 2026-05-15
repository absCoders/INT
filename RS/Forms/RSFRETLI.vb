Imports Infragistics.Win.UltraWinGrid

Public Class RSFRETLI

    'Dim filestoImport As New List(Of RetailSalesImporter)
    Dim editImportNumber As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

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

        End With

        grdRSTRETL5.DataSource = dst.Tables("RSTRETL5")
        grdRSTSSUM1.DataSource = dst.Tables("RSTSSUM1")
        grdRSTSSUM2.DataSource = dst.Tables("RSTSSUM2")

        grdTATEVNT1.DataSource = dst.Tables("TATEVNT1")
        'Fill_Records("TATEVNT1")
        'Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)

        Create_Summary(grdRSTRETL5, "AMT_SOLD", "Sum")

        Dim sql_where As String = ""
        sql_where = "OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -36) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "'"
        Load_Drop_Down("OPS_YYYYPP", sql_where)


        For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In cmbOPS_YYYYPP.Rows
            If row.Cells(0).Value.ToString = ASCMAIN1.CYP Then
                cmbOPS_YYYYPP.SelectedRow = row
                Exit For
            End If
        Next
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("OPS_YYYYPP")
                If cmbOPS_YYYYPP.Text < ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") Then
                    MsgBox("GL Period " & cmbOPS_YYYYPP.Text & " has been closed, You may Load but Updating is Prohibited", MsgBoxStyle.OkOnly, "Audit Retail Entry")
                End If

            Case "Update"

                If cmbOPS_YYYYPP.Text & "" < ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") Then
                    EMsg &= vbCr & "You cannot Update a Closed GL Period"
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

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
            End With

            grdRSTRETL5.Visible = ScreenMode = True

            Set_Read_Only(UltraGroupBox6, ScreenMode = True)
        End If

        'If Not tf Then
        '    LoadSummaryData()
        'End If

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        dst.EnforceConstraints = False
        If dst.Tables.Count > 0 Then
            dst.Tables("RSTRETL5").Clear()
        End If
        dst.EnforceConstraints = True

        LoadSummaryData()

        Fill_Records("TATEVNT1")
        Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)
    End Sub


    Sub Load_Record()
        ASCMAIN1.Progress("Now Loading Data ...")
        editImportNumber = ASCMAIN1.Next_Control_No("RSTIMPR1.IMPORT_NO")

        Fill_Records("RSTRETL5", New Object() {txtCustomerCode.Value, cmbOPS_YYYYPP.Value})
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()
        Update_Record_TDA("RSTRETL5")
        CommitTrans()
    End Sub


#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdRSTRETL5, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)
        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                'Case "grdRSTBUDR1"
                '    If grdRSTRETLA.Tag = "" Then
                '        e.Cancel = True
                '    End If
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"

            Case "Show GroupBox"
                'If grd IsNot Nothing Then
                '    Dim tlb_sbt As StateButtonTool = DirectCast(e.Tool, StateButtonTool)
                '    grd.DisplayLayout.Bands(0).ColHeadersVisible = tlb_sbt.Checked
                'End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            
        End Select
    End Sub
#End Region

    Public Sub cmbOPS_YYYYPP_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmbOPS_YYYYPP.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case Absx1.GetABSColumnName(sender)
                Case "OPS_YYYYPP"
                    If Not ScreenMode Then
                        Proceed_PreReq("Load")
                    End If
            End Select
        End If
    End Sub

#Region "RSTRETL5"

    Private Sub grdRSTRETL5_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdRSTRETL5.InitializeLayout
        'Dim storeNamesColumn As UltraGridColumn = Me.grdRSTRETL5.DisplayLayout.Bands(0).Columns("CUST_STORE_NAME")
        'storeNamesColumn.ValueList = Me.cmbNationalStores
        'storeNamesColumn.Style = ColumnStyle.DropDownValidate
        'storeNamesColumn.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.SuggestAppend
    End Sub

    Private Sub grdRSTRETL5_BeforeCellListDropDown(sender As System.Object, e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles grdRSTRETL5.BeforeCellListDropDown
        If Not (e.Cell.Row.IsAddRow Or e.Cell.Row.IsFilterRow) Then
            e.Cancel = True
        End If
    End Sub

    Private Sub grdRSTRETL5_BeforeRowActivate(sender As System.Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdRSTRETL5.BeforeRowActivate
        If e.Row.IsAddRow Or e.Row.IsFilterRow Then
            For Each cell In e.Row.Cells
                cell.IgnoreRowColActivation = True
                cell.Activation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
            Next
        End If
    End Sub

    Private Sub grdRSTRETL5_AfterRowInsert(sender As System.Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdRSTRETL5.AfterRowInsert
        e.Row.Cells("IMPORT_NO").Value = editImportNumber
        e.Row.Cells("CUST_CODE").Value = txtCustomerCode.Text

        e.Row.Cells("OPS_YYYYWW").Value = ASCDATA1.GetDataValue("SELECT YYYYWW FROM GLTPARM3 WHERE REL_WEEK=1 AND YYYYPP=:PARM1", "V", New Object() {cmbOPS_YYYYPP.Value})
        e.Row.Cells("OPS_YYYYPP").Value = cmbOPS_YYYYPP.Value



        With grdRSTRETL5.DisplayLayout.Bands(0)
            '.ColumnFilters.ClearAllFilters()
            'Show_Filter(grdRSTIMPR1, False)
        End With
    End Sub

    Private Sub grdRSTRETL5_BeforeCellUpdate(sender As System.Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdRSTRETL5.BeforeCellUpdate
        If e.Cell.Column.Key = "CUST_STORE_NO" Then
            'If Not Regex.IsMatch(e.Cell.Text & "", "^\d+$") And e.Cell.Text.ToUpper() <> "DIRECT" And txtCustomerCode.Text <> "ECOMSALE10" Then
            '    e.Cancel = True
            '    MsgBox("Invalid store #")
            'End If
        End If
    End Sub

    Private Sub grdRSTRETL5_AfterCellUpdate(sender As System.Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTRETL5.AfterCellUpdate
        Dim grid As UltraGrid = CType(sender, UltraGrid)
        grid.EventManager.SetEnabled(GridEventIds.AfterCellUpdate, False)

        Try
            If e.Cell.Column.Key = "CUST_STORE_NO" Then
                Dim newVal As String = e.Cell.Value.ToString().PadLeft(6, "0").ToUpper()
                e.Cell.SetValue(newVal, False)

                'Dim storeName As String = ASCDATA1.GetDataValue("SELECT CUST_STORE_NAME FROM ARTCUST2 WHERE CUST_CODE=:PARM1 AND CUST_STORE_NO=:PARM2", "VV", New Object() {txtCustomerCode.Text, newVal}) & ""
                'e.Cell.Row.Cells("CUST_STORE_NAME").SetValue(storeName, False)

                With grdRSTRETL5.DisplayLayout.Bands(0)
                    '.ColumnFilters.ClearAllFilters()
                    '.ColumnFilters("CUST_STORE_NO").FilterConditions.Add(Infragistics.Win.UltraWinGrid.FilterComparisionOperator.Match, newVal)

                    'Show_Filter(grdRSTIMPR1, True)
                End With
                'ElseIf e.Cell.Column.Key = "CUST_STORE_NAME" Then
                '    Dim parenIndex As Integer = e.Cell.Text.LastIndexOf("(")
                '    Dim storeNo As String = e.Cell.Text.Substring(parenIndex + 1, 6)
                '    grid.EventManager.SetEnabled(GridEventIds.BeforeCellUpdate, False)
                '    e.Cell.Row.Cells("CUST_STORE_NO").SetValue(storeNo, False)
                '    grid.EventManager.SetEnabled(GridEventIds.BeforeCellUpdate, True)
            End If
        Finally
            grid.EventManager.SetEnabled(GridEventIds.AfterCellUpdate, True)
        End Try
    End Sub


    Private Sub grdRSTRETL5_BeforeRowUpdate(sender As System.Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdRSTRETL5.BeforeRowUpdate


        Dim validMsg = ""
        Select Case isRowValid(e.Row)
            Case RowValidity.Valid
                GoTo Valid
            Case RowValidity.InvalidStore
                validMsg = "Invalid Store"
            Case RowValidity.InvalidCollection
                validMsg = "Invalid Collection"
            Case RowValidity.InvalidStoreAndCollection
                validMsg = "Invalid Store and Collection"
        End Select
        e.Cancel = True
        MsgBox(validMsg)

Valid:  grdRSTRETL5.DisplayLayout.RowScrollRegions(0).ScrollRowIntoView(e.Row)
    End Sub

#End Region


    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                'sql_where = "CUST_STATUS = 'A' AND TRADE_CLASS_CODE IN ('IND','NAT')"
        End Select
    End Sub

    Private Sub grdRSTRETL5_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdRSTRETL5.ClickCellButton
        Dim COLUMN_NAME As String = e.Cell.Column.Key
        Dim SQL As String = ""

        If e.Cell.Column.CellActivation = UltraWinGrid.Activation.NoEdit And Not e.Cell.Row.IsAddRow Then
            Exit Sub
        End If

        Select Case COLUMN_NAME
            Case "CUST_STORE_NO"
                ' If you need to limit the select then add a where clause
                SQL = String.Format("SELECT CUST_STORE_NO, CUST_STORE_NAME FROM ARTCUST2 WHERE CUST_CODE='{0}'", txtCustomerCode.Text)

            Case "COLLECTION_CODE"
                SQL = String.Format("SELECT COLLECTION_CODE, COLLECTION_NAME FROM ICTCOLL1")
        End Select

        ASCMAIN1.CodeSelector.SQL = SQL

        ASCMAIN1.CodeSelector.VIEW_NAME = ""
        ASCMAIN1.CodeSelector.MultipleSelections = False
        ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = e.Cell.Text
        Dim F As New ASFCODE1
        F.ShowDialog()
        F.Dispose()

        If ASCMAIN1.CodeSelector.Selections <> 0 Then

            Select Case COLUMN_NAME
                Case "CUST_STORE_NO", "COLLECTION_CODE"
                    grdRSTRETL5.ActiveRow.Cells(COLUMN_NAME).Value = ASCMAIN1.CodeSelector.SelectedRows(0).Item(0)
            End Select
        End If
    End Sub

    Private Sub grdRSTSSUM1_AfterSelectChange(sender As Object, e As AfterSelectChangeEventArgs) Handles grdRSTSSUM1.AfterSelectChange

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
    End Sub

    Private Sub grdRSTSSUM2_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdRSTSSUM2.InitializeRow
        If e.Row.Cells("FF_SLS").Value > 0 Then
            e.Row.Appearance.BackColor = Color.PaleGoldenrod
        End If
    End Sub

    Private Function isRowValid(ByVal row As UltraGridRow) As RowValidity

        Dim validStore As Boolean = ASCDATA1.GetDataValue("SELECT COUNT(*) FROM ARTCUST2 WHERE CUST_CODE=:PARM1 And CUST_STORE_NO=:PARM2", "VV", New Object() {Absx1.txtFor("CUST_CODE").Text, row.Cells("CUST_STORE_NO").Value}) > 0
        Dim validCollection As Boolean = ASCDATA1.GetDataValue("SELECT COUNT(*) FROM ICTCOLL1 WHERE COLLECTION_CODE=:PARM1", "V", New Object() {row.Cells("COLLECTION_CODE").Value}) > 0

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
