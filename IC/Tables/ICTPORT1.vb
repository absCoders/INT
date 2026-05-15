Public Class ICTPORT1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select ICTPORT2.*, ICTWHSE1.WHSE_DESC from ICTPORT2,ICTWHSE1 where ICTWHSE1.WHSE_CODE = ICTPORT2.WHSE_CODE and ICTPORT2.PORT_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTPORT2", "**", 0, True, "V", 2)
        End With

        grdICTPORT2.DataSource = dst.Tables("ICTPORT2")

        Create_Summary(grdICTPORT2, "WHSE_CODE", "Count")

    End Sub

    Private Sub grdICTPORT2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTPORT2.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "WHSE_CODE"
                Dim WHSE_CODE As String = e.Cell.Value & ""
                grdCodeDesc(grdICTPORT2, "ICTWHSE1", "WHSE_CODE", "WHSE_DESC")
        End Select
    End Sub

    Private Sub grdICTPORT2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTPORT2.AfterExitEditMode
        With grdICTPORT2
            Select Case .ActiveCell.Column.Key
                Case "WHSE_CODE"
                    Dim WHSE_CODE As String = .ActiveCell.Text
                    If WHSE_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(WHSE_CODE, .ActiveCell.Column.Key)
                    End If
            End Select
        End With

    End Sub

    Private Sub grdICTPORT2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTPORT2.AfterRowActivate
        With grdICTPORT2
            If .ActiveRow.IsAddRow Then
                .DisplayLayout.Bands(0).Columns("WHSE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .ActiveCell = grdICTPORT2.ActiveRow.Cells("WHSE_CODE")
                .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                '.DisplayLayout.Bands(0).Columns("WHSE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                ' why cant we edit the acct code?
            End If
        End With
    End Sub

#Region "Overrides"

    Overrides Sub Proceed_Update_Special_Pre()

        Dim sql As String = "Delete from ICTPORT2 where PORT_CODE = '" & Absx1.txtFor("PORT_CODE").Text & "'"
        Update_Record_TDA("ICTPORT2", sql)
    End Sub

    Overrides Sub Show_Record_Special()
        Dim txtctl As UltraWinEditors.UltraTextEditor
        txtctl = Absx1.txtFor("PORT_CODE")
        Call Clear_Record_Special()
        Call Load_Report_Form(txtctl.Text)
    End Sub

    Sub Load_Report_Form(ByVal PORT_CODE As String)

        Fill_Records("ICTPORT2", PORT_CODE)
        Sort_grdColumns(grdICTPORT2, "WHSE_CODE")

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("ICTPORT2").Rows.Clear()
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdICTPORT2.Enabled = tf

        '  MyBase.Mode_Settings(tf, MODE_description)
        With grdICTPORT2.DisplayLayout.Override
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

    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Update"
                

        End Select

    End Sub
#End Region

    Private Sub grdICTPORT2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTPORT2.BeforeRowUpdate
        With grdICTPORT2
            If e.Row.Cells("WHSE_CODE").Text = "" Then
                e.Cancel = True
            Else
                LookUp("ICTWHSE1", e.Row.Cells("WHSE_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Whse Code (" & e.Row.Cells("WHSE_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If Not e.Cancel Then
                If e.Row.Cells("PORT_CODE").Text = "" Then
                    .ActiveRow.Cells("PORT_CODE").Value = Absx1.CtlFor("PORT_CODE").Text
                End If
            End If
        End With

    End Sub

    Private Sub grdICTPORT2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTPORT2.ClickCellButton
        Dim sql_where As String = ""
        Call grdClickCellButton(grdICTPORT2, sql_where, sql_where <> "")
    End Sub

End Class