Public Class RSFBUDF1
    Dim RSTBUDFX As String
    Dim MOS As Integer
    Dim YPs() As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            ASCMAIN1.sql = "Select * from ICTCOLL1"
            Create_TDA(.Tables.Add, "ICTCOLL1", "**", 0, False)

            Create_TDA(.Tables.Add, "RSTBUDF1", "*")
            Create_BAs("RSTBUDF1")

            Create_Work_Tables()

            ASCMAIN1.sql = "Select * from " & RSTBUDFX
            Create_TDA(.Tables.Add, "RSTBUDFX", "**", 0, False, "", 3)

            With .Tables("RSTBUDFX").Columns
                .Add("TOTAL", GetType(System.Decimal), _
                      "ISNULL(BUDGET_P01,0)+ISNULL(BUDGET_P02,0)+ISNULL(BUDGET_P03,0)+" _
                    & "ISNULL(BUDGET_P04,0)+ISNULL(BUDGET_P05,0)+ISNULL(BUDGET_P06,0)+" _
                    & "ISNULL(BUDGET_P07,0)+ISNULL(BUDGET_P08,0)+ISNULL(BUDGET_P09,0)+" _
                    & "ISNULL(BUDGET_P10,0)+ISNULL(BUDGET_P11,0)+ISNULL(BUDGET_P12,0)")
            End With

        End With

        Fill_Records("ICTCOLL1")

        grdRSFBUDFX.DataSource = dst.Tables("RSTBUDFX")

        Create_Summary(grdRSFBUDFX, "COLLECTION_CODE", "Count")
        For M As Integer = 1 To 12
            Create_Summary(grdRSFBUDFX, "BUDGET_P" & Format(M, "00"), , , "###,##0")
        Next
        Create_Summary(grdRSFBUDFX, "TOTAL", , , "###,##0")

        With grdRSFBUDFX.DisplayLayout.Bands("RSTBUDFX")
            For Each COLUMN_NAME As String In New String() {"COLLECTION_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next

            For i As Integer = 1 To 12
                Dim COLUMN_NAME As String = "BUDGET_P" & Format(i, "00")
                .Columns(COLUMN_NAME).Format = "#,##0"
                .Columns(COLUMN_NAME).Width = 80
            Next
            .Columns("TOTAL").Format = "#,##0"
            .Columns("TOTAL").Width = 100

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = System.Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.Gold
                If New String() {"COLLECTION_CODE", "TOTAL"}.Contains(gcol.Key) Then
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    If gcol.Key <> "TOTAL" Then gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
        End With

        Absx1.cmbFor("RYP0").Value = Mid(ASCMAIN1.CYP, 1, 4) & "01"
        Absx1.cmbFor("RYP1").Value = Mid(ASCMAIN1.CYP, 1, 4) & "12"

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Calc_MOs()
                If MOS < 1 Or MOS > 12 Then
                    EMsg &= vbCr & "Period Range must span between 1 and 12 months"
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("RSTBUDF1", "*") Then
                        Exit Sub
                    End If
                End If

            Case "Update"

                For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("RSTBUDFX"), New String() {"COLLECTION_CODE"}).Select("")
                    Dim COLLECTION_CODE As String = row.Item(0)
                    Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(New String() {COLLECTION_CODE})
                    If rowICTCOLL1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Collection (" & COLLECTION_CODE & ")"
                    End If
                Next

            Case "Import from XLS"

                If MsgBox("This function will Import Financial Retail Sales Budget data" _
                & vbCrLf & " from a specifically formatted spreadsheet" _
                & vbCrLf & " and use that data to replace the data currently on file" _
                & vbCrLf _
                & vbCrLf & " for the Period Range from " & Absx1.cmbFor("RYP0").Text & " thru " & Absx1.cmbFor("RYP1").Text _
                & vbCrLf _
                & vbCrLf & "Once you click 'Yes' to proceed," _
                & vbCrLf & " you will be asked for the location of the spreadsheet, " _
                & vbCrLf & " and the data will be imported and displayed in the grid below." _
                & vbCrLf _
                & vbCrLf & "You will have an opportunity to review it before clicking 'Update'." _
                & vbCrLf _
                & vbCrLf & "Proceed with the Import?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
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
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Import from XLS"

                With grdRSFBUDFX.DisplayLayout.Bands(0)
                    For Each COLUMN_NAME As String In New String() {"COLLECTION_CODE"}
                        .Columns(COLUMN_NAME).Hidden = False
                        .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                    Next
                End With

                Excel_Import_SG(grdRSFBUDFX)
                Sort_grdColumns(grdRSFBUDFX, "COLLECTION_CODE")
                Setup_grd()
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
                    .Items("Import from XLS").Visible = ScreenMode '  Not tf
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        grdRSFBUDFX.Visible = ScreenMode


        If ScreenMode Then
            Setup_grd()
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"RSTBUDFX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)
        Create_Work_Tables()
        Fill_Records("RSTBUDFX")
        EnforceConstraints(True)

        Sort_grdColumns(grdRSFBUDFX, "COLLECTION_CODE")
        Set_Month_Headings()
        Setup_grd()

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        Dim RYP0 As String = Absx1.cmbFor("RYP0").Value
        Dim RYP1 As String = Absx1.cmbFor("RYP1").Value

        Dim sql_Delete As String = "Delete from RSTBUDF1" _
            & " where OPS_YYYYPP between '" & RYP0 & "' and '" & RYP1 & "'"

        dst.Tables("RSTBUDF1").Rows.Clear()

        For Each rowRSTBUDFX As DataRow In dst.Tables("RSTBUDFX").Select("")
            For I As Integer = 1 To MOS
                Dim BUDGET As Decimal = Val(rowRSTBUDFX.Item("BUDGET_P" & Format(I, "00")) & "")
                If BUDGET <> 0 Then
                    Dim rowRSTBUDF1 As DataRow = dst.Tables("RSTBUDF1").NewRow
                    rowRSTBUDF1.Item("COLLECTION_CODE") = rowRSTBUDFX.Item("COLLECTION_CODE")
                    rowRSTBUDF1.Item("OPS_YYYYPP") = YPs(I)
                    rowRSTBUDF1.Item("BUDGET") = BUDGET
                    dst.Tables("RSTBUDF1").Rows.Add(rowRSTBUDF1)
                End If
            Next
        Next

        ASCDATA1.ExecuteSQL(sql_Delete)
        Update_BAs("RSTBUDF1")

        Dim rowSOTCHAN1 As DataRow = LookUp("SOTCHAN1", "1")
        Dim CUST_CODE As String = rowSOTCHAN1.Item("CUST_CODE")

        ASCMAIN1.sql = "Delete from RSTBUDR1" & vbCrLf _
            & " where OPS_YYYYPP >= '" & YPs(1) & "' and OPS_YYYYPP <= '" & YPs(MOS) & "'" & vbCrLf _
            & "   and CUST_CODE = '" & CUST_CODE & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into RSTBUDR1" & vbCrLf _
            & "Select OPS_YYYYPP, '" & CUST_CODE & "' CUST_CODE, '000001' CUST_STORE_NO" & vbCrLf _
            & ", COLLECTION_CODE, 'E' ITEM_CATGY_CODE, FIN - SLS BUDGET from (" & vbCrLf _
            & "Select COLLECTION_CODE, OPS_YYYYPP, SUM (SLS) SLS, SUM (FIN) FIN from (" & vbCrLf _
            & "Select COLLECTION_CODE, OPS_YYYYPP, SUM (BUDGET) SLS, 0 FIN from RSTBUDR1 " & vbCrLf _
            & " where OPS_YYYYPP >= '" & YPs(1) & "' AND OPS_YYYYPP <= '" & YPs(MOS) & "'" & vbCrLf _
            & " group by COLLECTION_CODE, OPS_YYYYPP" & vbCrLf _
            & " union " & vbCrLf _
            & "Select COLLECTION_CODE, OPS_YYYYPP, 0 SLS, SUM (BUDGET) FIN FROM RSTBUDF1 " & vbCrLf _
            & " where OPS_YYYYPP >= '" & YPs(1) & "' AND OPS_YYYYPP <= '" & YPs(MOS) & "'" & vbCrLf _
            & " group by COLLECTION_CODE, OPS_YYYYPP" & vbCrLf _
            & ") group by COLLECTION_CODE, OPS_YYYYPP" & vbCrLf _
            & ")"
        ASCDATA1.ExecuteSQL()

        CommitTrans("Update Complete")
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup_Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdRSFBUDFX, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdRSTBUDFX"
                    'tlb_btn = DirectCast(tlb_pop.Tools("Clear Column"), UltraWinToolbars.ButtonTool)
                    'tlb_btn.SharedProps.Visible = ScreenMode And (EntryMode = "L")
                    'tlb_btn = DirectCast(tlb_pop.Tools("Copy Value"), UltraWinToolbars.ButtonTool)
                    'tlb_btn.SharedProps.Visible = ScreenMode And (EntryMode = "L")

                    If grdRSFBUDFX.Tag = "" Then
                        'e.Cancel = True
                    End If
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
            'Case "Clear Column"
            '    Dim COLUMN_NAME As String = grdRSTBUDFX.Tag
            '    If COLUMN_NAME = "" Then Exit Sub
            '    If COLUMN_NAME = "CUST_STORE_NO" Then Exit Sub
            '    For Each row As DataRow In dst.Tables("ARTCUST2").Select("", "", DataViewRowState.CurrentRows)
            '        row.Item(COLUMN_NAME) = DBNull.Value
            '    Next
            'Case "Copy Value"
            '    Dim COLUMN_NAME As String = grdRSTBUDFX.Tag
            '    If COLUMN_NAME = "" Then Exit Sub
            '    If grdRSTBUDFX.ActiveRow Is Nothing OrElse grdRSTBUDFX.ActiveRow.IsAddRow OrElse Not grdRSTBUDFX.ActiveRow.IsDataRow Then Exit Sub
            '    Dim COPY_VALUE As String = grdRSTBUDFX.ActiveRow.Cells(COLUMN_NAME).Value
            '    For Each row As DataRow In dst.Tables("ARTCUST2").Select("", "", DataViewRowState.CurrentRows)
            '        row.Item(COLUMN_NAME) = COPY_VALUE
            '    Next

            Case Else
                'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"
    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "CUST_CODE"
            '    If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Click_Command("Load", e)
            '    End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "CUST_CODE"
            '    Click_Command("Load")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "CUST_CODE"
            '    If EntryMode = "" Then
            '        If Absx1.txtFor("CUST_CODE").Text <> "" Then
            '            LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
            '            If cdr IsNot Nothing Then

            '            End If
            '        End If
            '    End If
        End Select
    End Sub

    Public Overrides Sub cmb_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.cmb_ValueChanged(sender, e)

        Calc_MOS()

    End Sub
 #End Region

    Sub Calc_MOs()

        If Absx1.cmbFor("RYP0").Value & "" <> "" And Absx1.cmbFor("RYP1").Value & "" <> "" Then
            MOS = 1 + ASCMAIN1.Period_Diff(Absx1.cmbFor("RYP0").Value, Absx1.cmbFor("RYP1").Value)
            lblMonths.Text = CStr(MOS) & " Mos"
        Else
            lblMonths.Text = ""
        End If

        Dim RYP0 As String = Absx1.cmbFor("RYP0").Value
        If MOS >= 1 And MOS <= 12 Then
            ReDim YPs(MOS)
            For i As Integer = 1 To MOS
                YPs(i) = ASCMAIN1.Period_Calc(RYP0, (i - 1))
            Next
        End If

    End Sub
#Region "grdRSTBUDFX"

    Private Sub grdRSTBUDFX_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdRSFBUDFX.AfterExitEditMode
        With grdRSFBUDFX
            Select Case .ActiveCell.Column.Key
                Case "COLLECTION_CODE"
                    Dim COLLECTION_CODE As String = .ActiveCell.Text
                    If COLLECTION_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(COLLECTION_CODE, .ActiveCell.Column.Key)
                    End If
            End Select
        End With
    End Sub

    Private Sub grdRSTBUDFX_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdRSFBUDFX.AfterRowActivate
        With grdRSFBUDFX.DisplayLayout.Bands(0)
            If grdRSFBUDFX.ActiveRow.IsAddRow Then
                .Columns("COLLECTION_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .Columns("COLLECTION_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdRSTBUDFX_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdRSFBUDFX.BeforeRowUpdate
        With grdRSFBUDFX

        End With
    End Sub

    Private Sub grdRSTBUDFX_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSFBUDFX.ClickCellButton
        Select Case grdRSFBUDFX.ActiveCell.Column.Key
            Case "COLLECTION_CODE"
                grdClickCellButton(grdRSFBUDFX)
        End Select
    End Sub
#End Region

    Sub Setup_grd()


        With grdRSFBUDFX.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            .SortedColumns.Add("COLLECTION_CODE", False)
        End With

        Dim COLLS As String = ""
        Dim allow_modifications As Boolean = True

        Dim RYP_LEGENDS As String = Absx1.cmbFor("RYP0").Text & " thru " & Absx1.cmbFor("RYP1").Text

        grdRSFBUDFX.Text = "Retail Sales Budgets, by Store/Month, for " & RYP_LEGENDS & " - " & COLLS
        If allow_modifications Then
            grdRSFBUDFX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdRSFBUDFX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            grdRSFBUDFX.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
        Else
            grdRSFBUDFX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdRSFBUDFX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            grdRSFBUDFX.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        End If
    End Sub

    Sub Set_Month_Headings()

        Dim RYP0 As String = Absx1.cmbFor("RYP0").Value
        Dim RYP1 As String = Absx1.cmbFor("RYP1").Value

        For M As Integer = 1 To 12
            With grdRSFBUDFX.DisplayLayout.Bands(0).Columns("BUDGET_P" & Format(M, "00"))
                Dim YP As String = ASCMAIN1.Period_Calc(RYP0, (M - 1))
                If YP > RYP1 Then
                    .Hidden = True
                Else
                    Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", YP)
                    Dim LEGEND As String = rowGLTPARM2.Item("LEGEND")
                    .Header.Caption = Mid(LEGEND, 10, 6)
                    .Width = 100
                    .Hidden = False
                End If
            End With
        Next
    End Sub

    Overrides Function Excel_Import_Pre_Process_SG _
    (ByVal grd As UltraWinGrid.UltraGrid, dt As DataTable,
     Optional ByRef load_by_table As Boolean = False, _
     Optional ByRef load_handled As Boolean = False, _
     Optional ByRef F As ASFEXCL1 = Nothing) As Int64

        Dim dtbad As DataTable = dt.Clone
        dtbad.Columns.Add("ERROR")

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Budgets from XLS")

        Dim RowsMax As Int64 = dt.Rows.Count
        Dim r As Int64 = 0

        load_handled = True
        If dt.Rows.Count = 0 Then
            MsgBox("No Rows Loaded", MsgBoxStyle.OkOnly, "Import Failed")
        Else
            dst.Tables("RSTBUDFX").Rows.Clear()
        End If

        For Each row As DataRow In dt.Select("")
            r += 1
            If r Mod 100 = 0 Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading Budgets from XLS")
                RowsMax = dt.Rows.Count
                ASCMAIN1.Progress("-", CStr(r) & "/" & CStr(RowsMax))
            End If

            Try
                Dim rowRSTBUDFX As DataRow = dst.Tables("RSTBUDFX").NewRow
                With rowRSTBUDFX
                    For Each C As String In New String() {"COLLECTION_CODE"}
                        .Item(C) = row.Item(C)
                    Next
                    For I As Integer = 1 To MOS
                        Dim C As String = "BUDGET_P" & Format(I, "00")
                        .Item(C) = row.Item(C)
                    Next
                End With

                dst.Tables("RSTBUDFX").Rows.Add(rowRSTBUDFX)

            Catch ex As Exception
                Dim rowbad As DataRow = dtbad.NewRow
                rowbad.ItemArray = row.ItemArray
                rowbad.Item("ERROR") = ex.Message
                dtbad.Rows.Add(rowbad)
            End Try
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        If dtbad.Rows.Count > 0 Then
            Using fr As New ASFMSGBF
                fr.Show_grd(dtbad, Me, "Some Rows Failed to Update - Please Check Last Column for Messages")
            End Using
        End If

    End Function

    Sub Create_Work_Tables()

        Dim RYP0 As String = Absx1.cmbFor("RYP0").Value
        Dim RYP1 As String = Absx1.cmbFor("RYP1").Value

        Dim sqlM As String = ""
        For I As Integer = 1 To 12
            Dim YP As String = ASCMAIN1.Period_Calc(RYP0, I - 1)
            If YP > RYP1 Then YP = ""
            sqlM &= ", Sum (Decode(RSTBUDF1.OPS_YYYYPP,'" & YP & "',BUDGET,0)) BUDGET_P" & Format(I, "00")
        Next

        ASCMAIN1.sql = "Select RSTBUDF1.COLLECTION_CODE" & vbCrLf _
            & sqlM _
            & " from RSTBUDF1" _
            & " where RSTBUDF1.OPS_YYYYPP between '" & RYP0 & "' and '" & RYP1 & "'" & vbCrLf _
            & " group by RSTBUDF1.COLLECTION_CODE"

        If RSTBUDFX = "" Then
            RSTBUDFX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & RSTBUDFX & " Add Primary Key (COLLECTION_CODE)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & RSTBUDFX)
            ASCDATA1.ExecuteSQL("Insert into " & RSTBUDFX & " " & ASCMAIN1.sql)
        End If
    End Sub
End Class