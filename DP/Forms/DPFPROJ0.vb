Public Class DPFPROJ0

    Dim YP(,) As String
    Dim YP_LY(,) As String
    Dim OPS_YYYY As String
    Dim SEASON As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")

        With dst
            Create_TDA(.Tables.Add, "ICTBRAN1", "*")

            Create_TDA(.Tables.Add, "DPTPROJ0", "*")

            ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
            & ", NVL(DPTPROJ0.ITEM_CATGY_CODE,'I') ITEM_CATGY_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_CATGY_CODE ITEM_CATGY_CODE_CURR " & vbCrLf _
            & ", ICTITEM1.HIDE_FROM_3PL ITEM_CATGY_CODE_3PL" & vbCrLf _
            & " from ICTITEM1,ICTCOLL1,DPTPROJ0 " & vbCrLf _
            & " where ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " and ICTCOLL1.BRAND_CODE = :PARM1 " & vbCrLf _
            & " and DPTPROJ0.OPS_YYYY (+) = :PARM2" & vbCrLf _
            & " and DPTPROJ0.SEASON (+) = :PARM3" & vbCrLf _
            & " and DPTPROJ0.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
            & " and (DPTPROJ0.ITEM_CODE IS NOT NULL OR NVL(ICTITEM1.HIDE_FROM_3PL,'I') <> 'I')"
            Create_TDA(.Tables.Add, "ICTITEMI", "**", 0, False, "VVV", 1)
            .Tables("ICTITEMI").Columns.Add("BAD_ITEM")
        End With

        grdICTITEMI.DataSource = dst.Tables("ICTITEMI")

        With grdICTITEMI.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key <> "ITEM_CATGY_CODE" Then
                    gcol.CellAppearance.BackColor = Color.Beige
                End If
            Next
            '.Columns("ITEM_CATGY_CODE").CellAppearance.BackColor = Color.Beige
        End With

        Create_Summary(grdICTITEMI, "ITEM_CODE", "Count")

        Dim YEARs As New List(Of String)
        For Y As Integer = Val(Now.Year) - 1 To Val(Now.Year + 1)
            YEARs.Add(Format(Y, "0000"))
        Next
        Absx1.cbeFor("OPS_YYYY").DataSource = YEARs

        ASCMAIN1.Add_Value_List(grdICTITEMI, "ITEM_CATGY_CODE")
        ASCMAIN1.Add_Value_List(grdICTITEMI, "ITEM_CATGY_CODE_CURR", "Select ITEM_CATGY_CODE, ITEM_CATGY_DESC from ICTCATG1")
        ASCMAIN1.Add_Value_List(grdICTITEMI, "ITEM_CATGY_CODE_3PL", "Select ITEM_CATGY_CODE, ITEM_CATGY_DESC from ICTCATG1")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                Validate_Code("BRAND_CODE")

                If cmbOPS_YYYY.Text = "" Then
                    EMsg &= vbCr & "You Must Specify a Year"
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("DPTFCSTD", Absx1.txtFor("BRAND_CODE").Text) Then
                        Exit Sub
                    End If
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

            Case "Update"
                Call Update_Record()
                Call Mode_Settings(False)

            Case "Cancel", "Done"
                Call Mode_Settings(False)

            Case "Load from Spreadsheet"

                Load_from_Spreadsheet()

            Case "Load from NAV"
                For Each row As DataRow In dst.Tables("ICTITEMI").Rows
                    row.Item("ITEM_CATGY_CODE") = row.Item("ITEM_CATGY_CODE_3PL")
                Next
                MsgBox("NAV values for Item Category have been copied", MsgBoxStyle.OkOnly, "Success")
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Load from Spreadsheet").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Load from NAV").Settings.Enabled = iScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdICTITEMI.Visible = ScreenMode
        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        dst.Tables("ICTITEMI").Rows.Clear()
        EnforceConstraints(True)

        Absx1.txtFor("BRAND_CODE").Text = ""
        If Absx1.cbeFor("OPS_YYYY").Value & "" = "" Then
            Absx1.cbeFor("OPS_YYYY").Value = Now.Year
        End If

        'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Differences Only"), UltraWinToolbars.StateButtonTool)
        'tlb_sbt.Checked = False - causes error in toolclick event

    End Sub

    Sub Load_Record()

        Call ASCMAIN1.Progress("Now Loading Data ...")
        Call Save_Header_Fields(UltraGroupBox1)
        OPS_YYYY = Absx1.cbeFor("OPS_YYYY").Text
        EnforceConstraints(False)
        Fill_Records("ICTITEMI", New String() {HFs("BRAND_CODE"), HFs("OPS_YYYY"), HFs("SEASON")})
        EnforceConstraints(True)
        Call ASCMAIN1.Progress("")

        grdICTITEMI.DisplayLayout.Bands(0).Columns("ITEM_CATGY_CODE").Header.Caption = "Catgy (" & HFs("SEASON") & Mid(HFs("OPS_YYYY"), 3, 2) & ")"
    End Sub

    Sub Update_Record()

        BeginTrans()

        dst.Tables("DPTPROJ0").Rows.Clear()
        For Each rowICTITEMI As DataRow In dst.Tables("ICTITEMI") _
        .Select("(ITEM_CATGY_CODE = 'C' OR ITEM_CATGY_CODE = 'N' OR ITEM_CATGY_CODE = 'E' OR ITEM_CATGY_CODE = 'F' OR ITEM_CATGY_CODE = 'P') AND ISNULL(BAD_ITEM,'0') <> '1'")
            Dim rowDPTPROJ0 As DataRow = dst.Tables("DPTPROJ0").NewRow
            rowDPTPROJ0.Item("OPS_YYYY") = HFs("OPS_YYYY")
            rowDPTPROJ0.Item("SEASON") = HFs("SEASON")
            rowDPTPROJ0.Item("ITEM_CODE") = rowICTITEMI.Item("ITEM_CODE")
            rowDPTPROJ0.Item("ITEM_CATGY_CODE") = rowICTITEMI.Item("ITEM_CATGY_CODE")
            dst.Tables("DPTPROJ0").Rows.Add(rowDPTPROJ0)
        Next

        Dim sql_Delete As String = "Delete from DPTPROJ0" _
            & " where OPS_YYYY = '" & HFs("OPS_YYYY") & "'" _
            & " and SEASON = '" & HFs("SEASON") & "'"
        Update_Record_TDA("DPTPROJ0", sql_Delete)

        ' IF CURRENT SEASON, THEN FIX ITEM MASTER
        ' 11/18/09 ANNA confirms THAT FISCAL 201001, WHICH IS Aug'2009, is SEASON F OPS_YYYY 2009 - so Projections use Calendar Year

        Dim CURR_SEASON As String = "S"
        Dim CURR_YEAR As String = Mid(ASCMAIN1.CYP, 1, 4)
        If Val(Mid(ASCMAIN1.CYP, 5, 2)) >= 1 And Val(Mid(ASCMAIN1.CYP, 5, 2)) <= 6 Then
            CURR_SEASON = "F"
            CURR_YEAR = Format(Val(CURR_YEAR) - 1, "0000")
        End If

        If HFs("OPS_YYYY") = CURR_YEAR And HFs("SEASON") = CURR_SEASON Then
            ASCMAIN1.sql = "UPDATE ICTITEM1 SET ITEM_CATGY_CODE = 'I'"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "UPDATE ICTITEM1 SET ITEM_CATGY_CODE = " _
            & " (SELECT ITEM_CATGY_CODE from DPTPROJ0" _
            & " where OPS_YYYY = '" & CURR_YEAR & "'" _
            & "   and SEASON = '" & CURR_SEASON & "'" _
            & "   and ITEM_CODE = ICTITEM1.ITEM_CODE)"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "UPDATE ICTITEM1 SET ITEM_STATUS = DECODE(ITEM_CATGY_CODE,'I','I','A')"
            ASCDATA1.ExecuteSQL()
        End If

        CommitTrans("Update Complete")

    End Sub
#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdICTITEMI, "SSSB", "Show Filter", "Show GroupBox", "Show Differences Only", "Load from Spreadsheet")
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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdDPTFCSTD"
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

            Case "Load from Spreadsheet"
                Load_from_Spreadsheet()

            Case "Show Differences Only"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Dim dvw As DataView = DirectCast(grdICTITEMI.DataSource, DataTable).DefaultView
                If tlb_sbt.Checked Then
                    dvw.RowFilter = "ISNULL(ITEM_CATGY_CODE,'?') <> ISNULL(ITEM_CATGY_CODE_CURR,'?')"
                    ASCMAIN1.Notify("Now Showing Only those Items with Differences")
                    grdICTITEMI.Text = "Item Category Definitions - Differences Only"
                Else
                    dvw.RowFilter = ""
                    ASCMAIN1.Notify("Now Showing All Items")
                    grdICTITEMI.Text = "Item Category Definitions"
                End If

        End Select

        Select Case grd.Name
            'Case "grdDPTFCSTD"
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
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
            Case "BRAND_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    'Call Click_Command("Load", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "BRAND_CODE"
                'Call Click_Command("Load")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

            Case "BRAND_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("BRAND_CODE").Text <> "" Then
                        Call LookUp("ICTBRAN1", Absx1.txtFor("BRAND_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

        End Select
    End Sub

#End Region

#Region "Excel Upload"

    Sub Load_from_Spreadsheet()
        'dst.Tables("ICTITEMI").Rows.Clear()

        Dim tbl As DataTable = dst.Tables("ICTITEMI").Clone
        tbl.Merge(dst.Tables("ICTITEMI"))
        dst.Tables("ICTITEMI").Rows.Clear()

        Excel_Import(grdICTITEMI)

        For Each row As DataRow In tbl.Rows
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim rowICTITEMI As DataRow = dst.Tables("ICTITEMI").Rows.Find(ITEM_CODE)
            If rowICTITEMI Is Nothing Then
                dst.Tables("ICTITEMI").Rows.Add(row.ItemArray)
            Else
                rowICTITEMI.Item("ITEM_CATGY_CODE_CURR") = row.Item("ITEM_CATGY_CODE_CURR")
                rowICTITEMI.Item("ITEM_CATGY_CODE_3PL") = row.Item("ITEM_CATGY_CODE_3PL")
                rowICTITEMI.Item("BAD_ITEM") = row.Item("BAD_ITEM")
            End If
        Next
        Sort_grdColumns(grdICTITEMI, "ITEM_CODE")
        grdICTITEMI.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

    End Sub

    Overrides Sub Excel_Import_Pre_Process _
    (ByVal grd As UltraWinGrid.UltraGrid, _
     Optional ByRef load_by_table As Boolean = False, _
     Optional ByRef load_handled As Boolean = False, _
     Optional ByRef F As ASFEXCL1 = Nothing)

        If grd.Name = "grdICTITEMI" Then
            load_by_table = True
        End If

    End Sub

    Overrides Sub Excel_Import_Post_Process(ByVal grd As UltraWinGrid.UltraGrid, F As ASFEXCL1)

        Dim BAD_ITEM_count As Int64 = dst.Tables("ICTITEMI").Select("BAD_ITEM = '1'").Length
        If BAD_ITEM_count <> 0 Then
            MsgBox("There were " & CStr(BAD_ITEM_count) & " Bad Items Loaded from Spreadsheet", MsgBoxStyle.OkOnly, CStr(dst.Tables("ICTITEMI").Rows.Count) & " Records Loaded")
        Else
            MsgBox("All Items Loaded Successfully", MsgBoxStyle.OkOnly, CStr(dst.Tables("ICTITEMI").Rows.Count) & " Records Loaded")
        End If

    End Sub

    Overrides Sub Excel_Import_Custom_Processing_row _
    (ByVal row As DataRow, ByVal grow As UltraWinGrid.UltraGridRow, _
     Optional ByVal TBL As DataTable = Nothing)
        'If optCI.Value = "I" Then
        'Else
        '    row.Item("CUST_CODE") = Absx1.txtFor("CUST_CODE").Text
        'End If
        If Len(row.Item("ITEM_CATGY_CODE") & "") > 1 Then
            row.Item("ITEM_CATGY_CODE") = Mid(row.Item("ITEM_CATGY_CODE"), 1, 1)
        End If

        Dim ITEM_CODE As String = row.Item("ITEM_CODE") & ""
        row.Item("ITEM_DESC") = ""
        row.Item("BAD_ITEM") = ""
        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
        If rowICTITEM1 IsNot Nothing Then
            row.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
            row.Item("ITEM_CATGY_CODE_CURR") = rowICTITEM1.Item("ITEM_CATGY_CODE")
            row.Item("ITEM_CATGY_CODE_3PL") = rowICTITEM1.Item("HIDE_FROM_3PL")
        Else
            row.Item("BAD_ITEM") = "1"
        End If
    End Sub
#End Region

#Region "grdICTITEMI"

    Private Sub grdICTITEMI_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTITEMI.AfterCellUpdate

        If e.Cell.Column.Key = "ITEM_CODE" Then
            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", e.Cell.Text)
            If rowICTITEM1 IsNot Nothing Then
                e.Cell.Row.Cells("ITEM_DESC").Value = rowICTITEM1.Item("ITEM_DESC") & ""
                e.Cell.Row.Cells("BAD_ITEM").Value = ""
            Else
                e.Cell.Row.Cells("ITEM_DESC").Value = ""
                e.Cell.Row.Cells("BAD_ITEM").Value = "1"
            End If
        End If

    End Sub

    Private Sub grdICTITEMI_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTITEMI.AfterRowActivate

        If grdICTITEMI.ActiveRow Is Nothing Then Exit Sub

        With grdICTITEMI.DisplayLayout.Bands(0)
        End With
    End Sub

    Private Sub grdICTITEMI_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTITEMI.AfterRowUpdate
    End Sub

    Private Sub grdICTITEMI_BeforeCellActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles grdICTITEMI.BeforeCellActivate
    End Sub

    Private Sub grdICTITEMI_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdICTITEMI.BeforeCellUpdate
    End Sub

    Private Sub grdICTITEMI_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdICTITEMI.BeforeExitEditMode
        Call grdFieldFormat(grdICTITEMI)
    End Sub

    Private Sub grdICTITEMI_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdICTITEMI.BeforeRowsDeleted

    End Sub

    Private Sub grdICTITEMI_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTITEMI.BeforeRowUpdate
        With grdICTITEMI
            If LookUp("ICTITEM1", e.Row.Cells("ITEM_CODE").Text) Is Nothing Then
                MsgBox("Invalid Value entered for Item Code (" & e.Row.Cells("ITEM_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
                Exit Sub
            End If
        End With
    End Sub

    Private Sub grdICTITEMI_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTITEMI.ClickCellButton
        Dim sql_where As String = ""
        Call grdClickCellButton(grdICTITEMI, sql_where, False)
    End Sub

    Private Sub grdICTITEMI_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdICTITEMI.Error
        grdICTITEMI.ActiveRow.CancelUpdate()
    End Sub

    Private Sub grdICTITEMI_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTITEMI.InitializeRow
        If e.Row.Cells("ITEM_CATGY_CODE").Value & "" _
        <> e.Row.Cells("ITEM_CATGY_CODE_CURR").Value & "" And e.Row.Cells("ITEM_CATGY_CODE_CURR").Value & "" <> "" Then
            e.Row.Cells("ITEM_CATGY_CODE").Appearance.BackColor = Color.Yellow
        Else
            e.Row.Cells("ITEM_CATGY_CODE").Appearance.BackColor = Color.Empty
        End If

        If e.Row.Cells("BAD_ITEM").Value & "" = "1" Then
            e.Row.Cells("BAD_ITEM").Appearance.BackColor = Color.Red
        Else
            e.Row.Cells("BAD_ITEM").Appearance.BackColor = Color.Empty
        End If
    End Sub
#End Region

End Class