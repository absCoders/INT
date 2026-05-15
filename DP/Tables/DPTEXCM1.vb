Public Class DPTEXCM1

    Dim tblASTVIEWC As New DataTable

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select DPTEXCS1.*, ICTITEM1.ITEM_DESC" _
            & " from DPTEXCS1,ICTITEM1 " _
            & " where DPTEXCS1.EXC_MSG_CODE = :PARM1" _
            & " and ICTITEM1.ITEM_CODE = DPTEXCS1.ITEM_CODE"
            Call Create_TDA(.Tables.Add, "DPTEXCS1", "**", 0, True, "V", 2)
        End With
        grdDPTEXCS1.DataSource = dst.Tables("DPTEXCS1")

    End Sub

#Region "Overrides"

    Overrides Sub Proceed_Update_Special_Pre()

        'ASCDATA1.ExecuteSQL("Delete from ASTVIEW2 where VIEW_NAME = '" & Absx1.txtFor("VIEW_NAME").Text & "' and TABLE_NAME = '" & Absx1.txtFor("TABLE_NAME").Text & "'")

        Call Update_Record_TDA("DPTEXCS1")
    End Sub

    Overrides Sub Show_Record_Special()
        Call Fill_Records("DPTEXCS1", Absx1.txtFor("EXC_MSG_CODE").Text)
    End Sub

    Overrides Sub Clear_Record_Special()

        If SELECTION_NO = 0 Then Exit Sub
        If ScreenMode Then
            dst.Tables("DPTEXCS1").Rows.Clear()
        End If

    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdDPTEXCS1.Enabled = tf
    End Sub

    Overrides Sub Proceed_Update_Special_Post()
        'Call ASCMAIN1.Load_Views()
    End Sub
#End Region

#Region "grdDPTEXCS1"

    Private Sub grdDPTEXCS1_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdDPTEXCS1.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"

                grdCodeDesc(grdDPTEXCS1, "ICTITEM1", "ITEM_CODE", "ITEM_DESC")
                If cdr IsNot Nothing Then
                    'e.Cell.Row.Cells("ORDR_QTY").Value = Val(ROWs("SOTPARM1").Item("SO_PARM_DEFAULT_QTY") & "")
                    'e.Cell.Row.Cells("PRICE_CATGY_CODE").Value = cdr.Item("PRICE_CATGY_CODE")
                    'e.Cell.Row.Cells("ITEM_ORDER_CODE").Value = cdr.Item("ITEM_ORDER_CODE")
                Else
                    grdDPTEXCS1.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If

            Case "ORDR_QTY"
                'With grdSOTORDR2.ActiveRow
                '    .Cells("LINE_AMOUNT").Value = Val(.Cells("ORDR_QTY").Value & "") * Val(.Cells("ORDR_UNIT_PRICE").Value & "")
                'End With
        End Select
    End Sub

    Private Sub grdDPTEXCS1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdDPTEXCS1.BeforeRowUpdate
        If e.Row.IsAddRow Then
            e.Row.Cells("EXC_MSG_CODE").Value = Absx1.txtFor("EXC_MSG_CODE").Text
        End If
    End Sub

    Private Sub grdDPTEXCS1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdDPTEXCS1.ClickCellButton
        Dim sql_where As String = ""
        Select Case grdDPTEXCS1.ActiveCell.Column.Key
            Case "ITEM_CODE"
        End Select

        Call grdClickCellButton(grdDPTEXCS1, sql_where, False)
    End Sub

    Private Sub grdDPTEXCS1_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdDPTEXCS1.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In grdDPTEXCS1.Selected.Rows
            If dst.Tables("DPTEXCS1").Rows(grow.ListIndex).RowState = DataRowState.Added Then
            Else
                '    MsgBox("Cannot Delete Existing Records", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                '    e.Cancel = True
                '    Exit For
            End If
        Next
    End Sub

    Private Sub grdDPTEXCS1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdDPTEXCS1.AfterRowActivate
        If grdDPTEXCS1.ActiveRow.IsAddRow Then
            grdDPTEXCS1.DisplayLayout.Bands(0).Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            grdDPTEXCS1.DisplayLayout.Bands(0).Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        End If
    End Sub

    Private Sub grdDPTEXCS1_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdDPTEXCS1.AfterExitEditMode
        With grdDPTEXCS1
            Select Case .ActiveCell.Column.Key
                Case "ITEM_CODE"
                    If .ActiveCell.Text <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(.ActiveCell.Text, .ActiveCell.Column.Key)
                    End If
            End Select
        End With
    End Sub
#End Region
End Class