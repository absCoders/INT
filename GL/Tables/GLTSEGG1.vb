Public Class GLTSEGG1

    Dim sqlGLTSEGG2 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GLTPARM1")

        With dst
            ASCMAIN1.sql = "Select GLTSEGG2.*, GLTSEGM1.ACCT_SEG_DESC" _
                & " from GLTSEGG2,GLTSEGM1" _
                & " where GLTSEGG2.ACCT_SEG_ID = :PARM1" _
                & "   and GLTSEGG2.ACCT_SEG_GROUP_CODE = :PARM2" _
                & "   and GLTSEGM1.ACCT_SEG_ID = GLTSEGG2.ACCT_SEG_ID" _
                & "   and GLTSEGM1.ACCT_SEG_CODE = GLTSEGG2.ACCT_SEG_CODE"
            Create_TDA(.Tables.Add, "GLTSEGG2", "**", 0, True, "VV", 3)
        End With

        grdGLTSEGG2.DataSource = dst.Tables("GLTSEGG2")

        Create_Summary(grdGLTSEGG2, "ACCT_SEG_CODE", "Count")

        With grdGLTSEGG2.DisplayLayout.Bands(0)
            .Columns("ACCT_SEG_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
        End With
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdGLTSEGG2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
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

            Case "grdGLTSEGG2"
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
                If grd.Name = "grdGLTSEGG2" Then
                    Add_Codes(grdGLTSEGG2, "GLTSEGM1", "ACCT_SEG_CODE", "Segments")
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

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim ACCT_SEG_ID As String = Absx1.txtFor("ACCT_SEG_ID").Text
        Dim ACCT_SEG_GROUP_CODE As String = Absx1.txtFor("ACCT_SEG_GROUP_CODE").Text
        Dim sqlDelete = "ACCT_SEG_ID = '" & ACCT_SEG_ID & "' and ACCT_SEG_GROUP_CODE = '" & ACCT_SEG_GROUP_CODE & "'"
        Update_Record_TDA("GLTSEGG2", sqlDelete)
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Fill_Records("GLTSEGG2", New String() {Absx1.txtFor("ACCT_SEG_ID").Text, Absx1.txtFor("ACCT_SEG_GROUP_CODE").Text})
        Sort_grdColumns(grdGLTSEGG2, "ACCT_SEG_CODE")
        EnforceConstraints(True)

        Dim ACCT_SEG_ID As String = Absx1.txtFor("ACCT_SEG_ID").Text
        Dim GL_PARM_SEG_DESC As String = ROWs("GLTPARM1").Item("GL_PARM_SEG" & ACCT_SEG_ID & "_DESC")
        grdGLTSEGG2.DisplayLayout.Bands(0).Columns("ACCT_SEG_CODE").Header.Caption = GL_PARM_SEG_DESC
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"GLTSEGG2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdGLTSEGG2.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdGLTSEGG2}
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

#Region "grdGLTSEGG2"

    Private Sub grdGLTSEGG2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTSEGG2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "ACCT_SEG_CODE"
                grdCodeDesc(grdGLTSEGG2, "GLTSEGM1", "ACCT_SEG_CODE", "ACCT_SEG_DESC")
        End Select
    End Sub

    Private Sub grdGLTSEGG2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdGLTSEGG2.AfterRowActivate

        With grdGLTSEGG2.DisplayLayout.Bands(0).Columns("ACCT_SEG_CODE")
            If grdGLTSEGG2.ActiveRow.IsAddRow Then
                .CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdGLTSEGG2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdGLTSEGG2.AfterRowsDeleted

    End Sub

    Private Sub grdGLTSEGG2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdGLTSEGG2.AfterRowUpdate

    End Sub

    Private Sub grdGLTSEGG2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdGLTSEGG2.BeforeRowsDeleted
        'For Each grow As UltraWinGrid.UltraGridRow In e.Rows
        '    Dim SREP_CODE As String = grow.Cells("SREP_CODE").Value
        '    Dim CUST_CODE As String = grow.Cells("CUST_CODE").Value
        '    Dim rowGLTSEGG2 As DataRow = dst.Tables("GLTSEGG2").Rows.Find(New String() {SREP_CODE, CUST_CODE})
        '    If Not rowGLTSEGG2.RowState = DataRowState.Added Then
        '        MsgBox("Cannot Delete Previously Updated Colors", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
        '        e.Cancel = True
        '    End If
        'Next
    End Sub

    Private Sub grdGLTSEGG2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdGLTSEGG2.BeforeRowUpdate

        Dim row As DataRow = LookUp("GLTSEGM1", New String() {e.Row.Cells("ACCT_SEG_ID").Text, e.Row.Cells("ACCT_SEG_CODE").Text})
        If row Is Nothing Then
            e.Cancel = True
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("ACCT_SEG_ID").Value = Absx1.txtFor("ACCT_SEG_ID").Text
            e.Row.Cells("ACCT_SEG_GROUP_CODE").Value = Absx1.txtFor("ACCT_SEG_GROUP_CODE").Text
        End If

    End Sub

    Private Sub grdGLTSEGG2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTSEGG2.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "ACCT_SEG_CODE"
                Dim sql_where As String = Get_List_of_Codes("GLTSEGM1.ACCT_SEG_CODE not in", "GLTSEGG2", "ACCT_SEG_CODE")
                grdClickCellButton(grdGLTSEGG2, sql_where, True)
        End Select
    End Sub

#End Region

End Class