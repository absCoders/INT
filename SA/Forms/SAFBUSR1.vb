Public Class SAFBUSR1

    Dim rowARTCUST1 As DataRow
    Dim ADDRESS() As String = {"CUST_STORE_NAME", "CUST_STORE_ADDR1", "CUST_STORE_ADDR2", "CUST_STORE_ADDR3", _
             "CUST_STORE_CITY", "CUST_STORE_STATE", "CUST_STORE_ZIP_CODE", "CUST_STORE_COUNTRY", _
             "CUST_STORE_CONTACT", "CUST_STORE_PHONE", "CUST_STORE_EXT", "CUST_STORE_FAX", "CUST_STORE_EMAIL", "GLOBAL_LOCATION_NUMBER"}

    Dim LAST_CHANGE_COLUMN_NAME As String
    Dim LAST_CHANGE_CELL_VALUE As String
    Dim COPY_VALUE_clipboard As String
    Dim COLUMN_NAME_clipboard As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Create_TDA(.Tables.Add, "ARTCUST1", "*")
            Create_TDA(.Tables.Add, "ARTCUST2", "*", 1)
        End With

        grdARTCUST2.DataSource = dst.Tables("ARTCUST2")

        Create_Summary(grdARTCUST2, "CUST_STORE_NO", "Count")

        With grdARTCUST2.DisplayLayout.Bands("ARTCUST2")
            .Columns("CUST_STORE_NO").Header.Fixed = True
        End With

        For Each COLUMN_NAME As String In ADDRESS
            With grdARTCUST2.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Appearance
                .BackColor = Drawing.Color.Yellow
                .BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            End With
            grdARTCUST2.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = True
            If COLUMN_NAME = "CUST_STORE_NAME" Then grdARTCUST2.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = False
        Next
        grdARTCUST2.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").CellAppearance.BackColor = Drawing.Color.Beige

        ASCMAIN1.Add_Value_List(grdARTCUST2, "CUST_STORE_STATUS")

        If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
            grdARTCUST2.DisplayLayout.Bands(0).Columns("SREP_CODE").Hidden = False
        Else
            grdARTCUST2.DisplayLayout.Bands(0).Columns("SREP_CODE").Hidden = True
        End If

        Dim udd As New UltraWinGrid.UltraDropDown
        Dim DVW As DataView = New DataView(dst.Tables("ARTCUST2"), "CUST_DC_IND = '1'", "CUST_STORE_NO", DataViewRowState.CurrentRows)
        udd.DataSource = DVW
        For Each GC As UltraWinGrid.UltraGridColumn In udd.DisplayLayout.Bands(0).Columns
            If GC.Key <> "CUST_STORE_NO" Then
                GC.Hidden = True
            Else
                GC.Header.Caption = "DC No"
            End If
        Next
        udd.ValueMember = "CUST_STORE_NO"
        With grdARTCUST2.DisplayLayout.Bands(0).Columns("CUST_DC_NO")
            .ValueList = udd
            .ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
        End With

        grpStores.Visible = False
        grpARTCUST1.Visible = False
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("CUST_CODE")

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("ARTCUST1", Absx1.txtFor("CUST_CODE").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Update"
                'If Absx1.txtFor("STAX_CODE").Text = "" And Val(Absx1.numFor("INV_STAX").Value & "") <> 0 Then
                '    EMsg &= vbCr & "You Must Specify a Tax Code"
                'End If

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

                .Groups("Display Options").Visible = ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grpStores.Visible = tf
        grpARTCUST1.Visible = tf

        Set_Read_Only(grpARTCUST1, False)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ARTCUST1", "ARTCUST2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Absx1.txtFor("CUST_CODE").Text = ""

        LAST_CHANGE_CELL_VALUE = ""
        LAST_CHANGE_COLUMN_NAME = ""
        COPY_VALUE_clipboard = ""
        COLUMN_NAME_clipboard = ""
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        'If EntryMode = "N" Then
        '    Absx1.txtFor("INV_NO").Text = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
        'End If

        EnforceConstraints(False)
        rowARTCUST1 = Fill_Record("ARTCUST1", HFs("CUST_CODE"))
        Fill_Records("ARTCUST2", HFs("CUST_CODE"))
        EnforceConstraints(True)

        Sort_grdColumns(grdARTCUST2, "CUST_STORE_NO")
        optShow.Value = "A"

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        rowARTCUST1.Item("LAST_DATE") = DATETIME_STAMP
        rowARTCUST1.Item("LAST_OPER") = ASCMAIN1.USER_ID

        For Each rowARTCUST2 As DataRow In dst.Tables("ARTCUST2").Select("", "", DataViewRowState.CurrentRows)
            If rowARTCUST2.RowState = DataRowState.Added Then
                rowARTCUST2.Item("LAST_DATE") = DATETIME_STAMP
                rowARTCUST2.Item("LAST_OPER") = ASCMAIN1.USER_ID
            End If

            rowARTCUST2.Item("LAST_DATE") = DATETIME_STAMP
            rowARTCUST2.Item("LAST_OPER") = ASCMAIN1.USER_ID
        Next

        Update_Record_TDA("ARTCUST1")
        Update_Record_TDA("ARTCUST2")

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
        Load_Popup_Menu(grdARTCUST2, "SSSSSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Show Vertically", "Add New Stores", "Show Full Address", "Clear Column", "Copy Value and Paste to All Stores", "Copy Value to Clipboard", "Paste Value to Selected Stores")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
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
            Case "grdARTCUST2"

                tlb_btn = DirectCast(tlb_pop.Tools("Copy Value and Paste to All Stores"), UltraWinToolbars.ButtonTool)
                If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
                    tlb_btn.SharedProps.Visible = False
                Else
                    tlb_btn.SharedProps.Caption = "Copy Value and Paste to All Stores"
                    tlb_btn.SharedProps.Visible = True
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Copy Value to Clipboard"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (grd.Selected.Rows.Count = 0)
                If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Or grd.ActiveCell Is Nothing Then
                    tlb_btn.SharedProps.Visible = False
                Else
                    tlb_btn.SharedProps.Caption = "Copy '" & grd.ActiveCell.Value & "' to Clipboard"
                    tlb_btn.SharedProps.Visible = True
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Paste Value to Selected Stores"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (COLUMN_NAME_clipboard <> "") And grd.Selected.Rows.Count > 0
                tlb_btn.SharedProps.Caption = "Paste '" & COPY_VALUE_clipboard & "' to Selected Stores"
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdARTCUST2"
                    'If grdARTCUST2.Tag = "" Then
                    '    e.Cancel = True
                    'End If
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Show Vertically"
                tlb_sbt = DirectCast(tlb.Tools("Show Vertically"), UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    Dim tlb_sbt2 As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Add New Stores"), UltraWinToolbars.StateButtonTool)
                    tlb_sbt2.Checked = False
                End If
                grdARTCUST2.DisplayLayout.Bands(0).CardView = tlb_sbt.Checked

            Case "Add New Stores"
                tlb_sbt = DirectCast(tlb.Tools("Add New Stores"), UltraWinToolbars.StateButtonTool)

                If tlb_sbt.Checked Then
                    Dim tlb_sbt2 As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Vertically"), UltraWinToolbars.StateButtonTool)
                    tlb_sbt2.Checked = False
                End If

                If tlb_sbt.Checked Then
                    grdARTCUST2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                Else
                    grdARTCUST2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                End If

            Case "Show Full Address"
                tlb_sbt = DirectCast(tlb.Tools("Show Full Address"), UltraWinToolbars.StateButtonTool)
                With grdARTCUST2.DisplayLayout.Bands(0)
                    For Each COLUMN_NAME As String In ADDRESS
                        If COLUMN_NAME = "CUST_STORE_NAME" Then
                        Else
                            .Columns(COLUMN_NAME).Hidden = Not tlb_sbt.Checked
                        End If
                    Next
                End With

            Case "Paste Value to Selected Stores"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    grow.Cells(COLUMN_NAME_clipboard).Value = COPY_VALUE_clipboard
                    grow.Update()
                Next
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Clear Column"
                Dim COLUMN_NAME As String = grd.ActiveCell.Column.Key
                If COLUMN_NAME = "" Then Exit Sub
                If COLUMN_NAME = "CUST_STORE_NO" Then Exit Sub
                For Each row As DataRow In dst.Tables("ARTCUST2").Select("", "", DataViewRowState.CurrentRows)
                    row.Item(COLUMN_NAME) = DBNull.Value
                Next

            Case "Copy Value and Paste to All Stores"
                Dim COLUMN_NAME As String = grd.ActiveCell.Column.Key
                If COLUMN_NAME = "" Then Exit Sub
                If grdARTCUST2.ActiveRow Is Nothing OrElse grdARTCUST2.ActiveRow.IsAddRow OrElse Not grdARTCUST2.ActiveRow.IsDataRow Then Exit Sub
                Dim COPY_VALUE As String = grdARTCUST2.ActiveRow.Cells(COLUMN_NAME).Value & ""
                For Each row As DataRow In dst.Tables("ARTCUST2").Select("", "", DataViewRowState.CurrentRows)
                    row.Item(COLUMN_NAME) = COPY_VALUE
                Next

            Case "Copy Value to Clipboard"
                Dim COLUMN_NAME As String = grd.ActiveCell.Column.Key
                If COLUMN_NAME = "" Then Exit Sub
                If COLUMN_NAME = "CUST_STORE_NO" Then
                    MsgBox("Cannot Copy and Paste Store Numbers")
                    Exit Sub
                End If
                If grdARTCUST2.ActiveRow Is Nothing OrElse grdARTCUST2.ActiveRow.IsAddRow OrElse Not grdARTCUST2.ActiveRow.IsDataRow Then Exit Sub
                COPY_VALUE_clipboard = grdARTCUST2.ActiveRow.Cells(COLUMN_NAME).Value
                COLUMN_NAME_clipboard = COLUMN_NAME

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Load", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Click_Command("Load")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                        LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

        End Select
    End Sub

#End Region

#Region "grdARTCUST2"

    Private Sub grdARTCUST2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUST2.AfterCellUpdate
        If e.Cell.Column.Key = "CUST_DC_IND" Then
            'grdARTCUST2.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
            grdARTCUST2.UpdateData()
        End If

        If e.Cell.Value & "" <> "" Then
            LAST_CHANGE_CELL_VALUE = e.Cell.Value
            LAST_CHANGE_COLUMN_NAME = e.Cell.Column.Key
        End If

    End Sub

    Private Sub grdARTCUST2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTCUST2.BeforeRowUpdate
        If e.Row.IsAddRow Then
            e.Row.Cells("CUST_CODE").Value = HFs("CUST_CODE")
        End If
    End Sub

    Private Sub grdARTCUST2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUST2.ClickCellButton
        Dim sql_where As String = ""
        Select Case grdARTCUST2.ActiveCell.Column.Key
            Case "SELL_CODE"
        End Select

        grdClickCellButton(grdARTCUST2, sql_where, False)
    End Sub

    Private Sub grdARTCUST2_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdARTCUST2.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In grdARTCUST2.Selected.Rows
            If dst.Tables("ARTCUST2").Rows(grow.ListIndex).RowState = DataRowState.Added Then
            Else
                MsgBox("Cannot Delete Existing Store Records", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                e.Cancel = True
                Exit For
            End If
        Next
    End Sub

    Private Sub grdARTCUST2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTCUST2.AfterRowActivate
        If grdARTCUST2.ActiveRow.IsAddRow Then
            grdARTCUST2.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            grdARTCUST2.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").CellActivation = UltraWinGrid.Activation.NoEdit
        End If
    End Sub

    Private Sub grdARTCUST2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTCUST2.AfterExitEditMode
        With grdARTCUST2
            Select Case .ActiveCell.Column.Key
                Case "CUST_STORE_NO"
                    If .ActiveCell.Text <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(.ActiveCell.Text, .ActiveCell.Column.Key)
                    End If
            End Select
        End With
    End Sub
#End Region

    Private Sub optShow_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optShow.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Dim DVW As DataView = DirectCast(grdARTCUST2.DataSource, DataTable).DefaultView
        Select Case optShow.Value
            Case "A"
                DVW.RowFilter = ""
            Case "S"
                DVW.RowFilter = "CUST_DC_IND IS NULL OR CUST_DC_IND <> '1'"
            Case "D"
                DVW.RowFilter = "CUST_DC_IND = '1'"
        End Select

    End Sub
End Class