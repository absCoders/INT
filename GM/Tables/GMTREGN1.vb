Public Class GMTREGN1

    Dim sqlGMTSTOR1 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select GMTSTOR1.*, GMTODIV1.OP_GRP_CODE" _
                & " from GMTSTOR1,GMTODIV1" _
                & " where GMTSTOR1.REGION_CODE = :PARM1" _
                & "   and GMTODIV1.OP_DIV_CODE (+) = GMTSTOR1.OP_DIV_CODE"
            Create_TDA(.Tables.Add, "GMTSTOR1", "**", 0, True, "V", 1)
        End With

        grdGMTSTOR1.DataSource = dst.Tables("GMTSTOR1")

        'With grdGMTSTOR1.DisplayLayout.Bands(0)
        '    .Columns("ITEM_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
        'End With
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdGMTSTOR1, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
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

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()

    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Fill_Records("GMTSTOR1", New String() {Absx1.txtFor("REGION_CODE").Text})
        Sort_grdColumns(grdGMTSTOR1, "STORE_NO")
        EnforceConstraints(True)
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"GMTSTOR1"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdGMTSTOR1.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdGMTSTOR1}
            With grd.DisplayLayout.Override
                'If EntryMode = "New" Or EntryMode = "Edit" Then
                '    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                '    .AllowUpdate = DefaultableBoolean.True
                '    .AllowDelete = DefaultableBoolean.True
                'Else
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
                'End If
            End With
        Next
    End Sub

#End Region

#Region "grdGMTSTOR1"

    Private Sub grdGMTSTOR1_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGMTSTOR1.AfterCellUpdate
        'Select Case e.Cell.Column.Key
        '    Case "ITEM_CODE"
        '        grdCodeDesc(grdGMTSTOR1, "ICTITEM1", "ITEM_CODE", "ITEM_DESC")
        'End Select
    End Sub

    Private Sub grdGMTSTOR1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdGMTSTOR1.AfterRowActivate

        'With grdGMTSTOR1.DisplayLayout.Bands(0).Columns("ITEM_CODE")
        '    If grdGMTSTOR1.ActiveRow.IsAddRow Then
        '        .CellActivation = UltraWinGrid.Activation.AllowEdit
        '    Else
        '        .CellActivation = UltraWinGrid.Activation.NoEdit
        '    End If
        'End With
    End Sub

    Private Sub grdGMTSTOR1_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdGMTSTOR1.AfterRowsDeleted

    End Sub

    Private Sub grdGMTSTOR1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdGMTSTOR1.AfterRowUpdate

    End Sub

    Private Sub grdGMTSTOR1_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdGMTSTOR1.BeforeRowsDeleted
        'For Each grow As UltraWinGrid.UltraGridRow In e.Rows
        '    Dim SREP_CODE As String = grow.Cells("SREP_CODE").Value
        '    Dim CUST_CODE As String = grow.Cells("CUST_CODE").Value
        '    Dim rowGMTSTOR1 As DataRow = dst.Tables("GMTSTOR1").Rows.Find(New String() {SREP_CODE, CUST_CODE})
        '    If Not rowGMTSTOR1.RowState = DataRowState.Added Then
        '        MsgBox("Cannot Delete Previously Updated Colors", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
        '        e.Cancel = True
        '    End If
        'Next
    End Sub

    Private Sub grdGMTSTOR1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdGMTSTOR1.BeforeRowUpdate

        'Dim row As DataRow = LookUp("ICTITEM1", e.Row.Cells("ITEM_CODE").Text)
        'If row Is Nothing Then
        '    e.Cancel = True
        'End If

        'If e.Row.IsAddRow Then
        '    e.Row.Cells("ORDR_FORM_CODE").Value = Absx1.txtFor("ORDR_FORM_CODE").Text
        '    e.Row.Cells("ORDR_FORM_LNO").Value = Val(dst.Tables("GMTSTOR1").Compute("MAX(ORDR_FORM_LNO)", "") & "") + 10
        'End If

    End Sub

    Private Sub grdGMTSTOR1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGMTSTOR1.ClickCellButton
        'Select Case e.Cell.Column.Key
        '    Case "ITEM_CODE"
        '        Dim sql_where As String = Get_List_of_Codes("ICTITEM1.ITEM_CODE not in", "GMTSTOR1", "ITEM_CODE")
        '        grdClickCellButton(grdGMTSTOR1, sql_where, True)
        'End Select
    End Sub

#End Region
End Class