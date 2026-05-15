Public Class SOTSREP1

    Dim sqlSOTSREP2 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select SOTSREP2.*, ARTCUST1.CUST_NAME" _
                & " from SOTSREP2,ARTCUST1" _
                & " where SOTSREP2.SREP_CODE = :PARM1" _
                & "   and ARTCUST1.CUST_CODE = SOTSREP2.CUST_CODE"
            Create_TDA(.Tables.Add, "SOTSREP2", "**", 0, True, "V", 2)

            ASCMAIN1.sql = "Select ARTCUST1.*" _
                & " from ARTCUST1" _
                & " where ARTCUST1.SREP_CODE_OVER = :PARM1"
            Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, True, "V", 1, "SREP_CODE_OVER,SREP_COMM_PCT_OVER")
            .Tables("ARTCUST1").Columns("CURR_CODE").AllowDBNull = True
        End With

        grdSOTSREP2.DataSource = dst.Tables("SOTSREP2")
        grdARTCUST1.DataSource = dst.Tables("ARTCUST1")
        'Create_Summary(grdSOTSREP2, "CUST_CODE", "Count")
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTSREP2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
        Load_Popup_Menu(grdARTCUST1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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
 
            Case "grdSOTSREP2"
                tlb_btn = DirectCast(tlb_pop.Tools("Add Codes"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")

            Case "grdARTCUST1"
                tlb_btn = DirectCast(tlb_pop.Tools("Add Codes"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Select Case e.Tool.Key
            Case "Add Codes"
                If grd.Name = "grdSOTSREP2" Then
                    Add_Codes(grdSOTSREP2, "ARTCUST1", "CUST_CODE", "Customers")
                ElseIf grd.Name = "grdARTCUST1" Then
                    Add_Codes(grdARTCUST1, "ARTCUST1", "CUST_CODE", "Customers")
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub
#End Region
#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
            Case "Edit"

            Case "Update"
                For Each row As DataRow In dst.Tables("ARTCUST1").Rows
                    Dim CUST_CODE As String = ""
                    If row.RowState = DataRowState.Deleted Then
                        CUST_CODE = row.Item("CUST_CODE", DataRowVersion.Original)
                    Else
                        CUST_CODE = row.Item("CUST_CODE")
                    End If
                    If Not ASCMAIN1.Logical_Lock("ARTCUST1", CUST_CODE, False, , , 1) Then
                        EMsg &= "Unable to Update until exclusive access to all Override Customers is obtained"
                        Exit For
                    End If
                Next
        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim SREP_CODE As String = Absx1.txtFor("SREP_CODE").Text
        Dim sqlDelete = "SREP_CODE = '" & SREP_CODE & "'"
        Update_Record_TDA("SOTSREP2", sqlDelete)

        ASCDATA1.ExecuteSQL("Update ARTCUST1 Set SREP_CODE_OVER = Null, SREP_COMM_PCT_OVER = NULL" _
                            & " where SREP_CODE_OVER = '" & SREP_CODE & "'")
        dst.Tables("ARTCUST1").AcceptChanges()
        For Each rowARTCUST1 As DataRow In dst.Tables("ARTCUST1").Select("")
            rowARTCUST1.SetModified()
        Next
        'For Each rowARTCUST1 As DataRow In dst.Tables("ARTCUST1").Select("", "", DataViewRowState.Deleted)
        '    rowARTCUST1.AcceptChanges()
        'Next
        'For Each rowARTCUST1 As DataRow In dst.Tables("ARTCUST1").Select("", "", DataViewRowState.Added)
        '    rowARTCUST1.AcceptChanges()
        '    rowARTCUST1.SetModified()
        'Next
        Update_Record_TDA("ARTCUST1")
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Fill_Records("SOTSREP2", New String() {Absx1.txtFor("SREP_CODE").Text})
        Sort_grdColumns(grdSOTSREP2, "CUST_CODE")
        Fill_Records("ARTCUST1", New String() {Absx1.txtFor("SREP_CODE").Text})
        Sort_grdColumns(grdARTCUST1, "CUST_CODE")
        EnforceConstraints(True)
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"SOTSREP2", "ARTCUST1"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdSOTSREP2.Enabled = tf
        grdARTCUST1.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTSREP2, grdARTCUST1}
            With grd.DisplayLayout.Override
                If EntryMode = "New" Or EntryMode = "Edit" Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.True
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                End If
            End With
        Next
    End Sub

#End Region

#Region "grdSOTSREP2"

    Private Sub grdSOTSREP2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTSREP2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "CUST_CODE"
                grdCodeDesc(grdSOTSREP2, "ARTCUST1", "CUST_CODE", "CUST_NAME")
        End Select
    End Sub

    Private Sub grdSOTSREP2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTSREP2.AfterRowActivate

        With grdSOTSREP2.DisplayLayout.Bands(0).Columns("CUST_CODE")
            If grdSOTSREP2.ActiveRow.IsAddRow Then
                .CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdSOTSREP2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTSREP2.AfterRowsDeleted

    End Sub

    Private Sub grdSOTSREP2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTSREP2.AfterRowUpdate

    End Sub

    Private Sub grdSOTSREP2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTSREP2.BeforeRowsDeleted
        'For Each grow As UltraWinGrid.UltraGridRow In e.Rows
        '    Dim SREP_CODE As String = grow.Cells("SREP_CODE").Value
        '    Dim CUST_CODE As String = grow.Cells("CUST_CODE").Value
        '    Dim rowSOTSREP2 As DataRow = dst.Tables("SOTSREP2").Rows.Find(New String() {SREP_CODE, CUST_CODE})
        '    If Not rowSOTSREP2.RowState = DataRowState.Added Then
        '        MsgBox("Cannot Delete Previously Updated Colors", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
        '        e.Cancel = True
        '    End If
        'Next
    End Sub

    Private Sub grdSOTSREP2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTSREP2.BeforeRowUpdate

        Dim row As DataRow = LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Text)
        If row Is Nothing Then
            e.Cancel = True
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("SREP_CODE").Value = Absx1.txtFor("SREP_CODE").Text
        End If

    End Sub

    Private Sub grdSOTSREP2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTSREP2.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "CUST_CODE"
                Dim sql_where As String = Get_List_of_Codes("ARTCUST1.CUST_CODE not in", "SOTSREP2", "CUST_CODE")
                grdClickCellButton(grdSOTSREP2, sql_where, True)
        End Select
    End Sub

    Private Sub grdSOTSREP2_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTSREP2.InitializeRow

    End Sub

#End Region

#Region "grdARTCUST1"

    Private Sub grdARTCUST1_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUST1.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "CUST_CODE"
                grdCodeDesc(grdARTCUST1, "ARTCUST1", "CUST_CODE", "CUST_NAME")
        End Select
    End Sub

    Private Sub grdARTCUST1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdARTCUST1.AfterRowActivate

        With grdARTCUST1.DisplayLayout.Bands(0).Columns("CUST_CODE")
            If grdARTCUST1.ActiveRow.IsAddRow Then
                .CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdARTCUST1_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdARTCUST1.AfterRowsDeleted

    End Sub

    Private Sub grdARTCUST1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdARTCUST1.AfterRowUpdate

    End Sub

    Private Sub grdARTCUST1_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdARTCUST1.BeforeRowsDeleted
        'For Each grow As UltraWinGrid.UltraGridRow In e.Rows
        '    Dim SREP_CODE As String = grow.Cells("SREP_CODE").Value
        '    Dim CUST_CODE As String = grow.Cells("CUST_CODE").Value
        '    Dim rowARTCUST1 As DataRow = dst.Tables("ARTCUST1").Rows.Find(New String() {SREP_CODE, CUST_CODE})
        '    If Not rowARTCUST1.RowState = DataRowState.Added Then
        '        MsgBox("Cannot Delete Previously Updated Colors", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
        '        e.Cancel = True
        '    End If
        'Next
    End Sub

    Private Sub grdARTCUST1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTCUST1.BeforeRowUpdate

        Dim row As DataRow = LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Text)
        If row Is Nothing Then
            e.Cancel = True
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("SREP_CODE_OVER").Value = Absx1.txtFor("SREP_CODE").Text
        End If

    End Sub

    Private Sub grdARTCUST1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUST1.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "CUST_CODE"
                Dim sql_where As String = Get_List_of_Codes("ARTCUST1.CUST_CODE not in", "ARTCUST1", "CUST_CODE")
                grdClickCellButton(grdARTCUST1, sql_where, True)
        End Select
    End Sub

#End Region
End Class