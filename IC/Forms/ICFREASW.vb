Imports ABSolution
Imports Infragistics.Win.UltraWinGrid

Public Class ICFREASW

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "ICTREASW", "*")
            .Tables("ICTREASW").Columns.Add("REASON_DESC", GetType(String), "REASON_CODE")

            Create_TDA(.Tables.Add, "ICTREAS1", "*")
            Fill_Records("ICTREAS1", String.Empty, True, "SELECT * FROM ICTREAS1")
        End With

        grdICTREASW.DataSource = dst.Tables("ICTREASW")
        ASCMAIN1.Add_Value_List(grdICTREASW, "REASON_DESC", "SELECT REASON_CODE, REASON_DESC FROM ICTREAS1")

        Mode_Settings(False)

        Create_Summary(grdICTREASW, "WHSE_REASON_CODE", "Count")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = ""

        Select Case eItemKey

            Case "Load"

            Case "Cancel"

            Case "Update"

        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Cancel"
                Mode_Settings(False)

            Case "Update"
                Update_Record()
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
            End With
        End If

        grdICTREASW.Visible = ScreenMode

    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)
        dst.Tables("ICTREASW").Rows.Clear()
        EnforceConstraints(True)
    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("Now Loading information")

        EnforceConstraints(False)
        Fill_Records("ICTREASW", String.Empty, True, "SELECT * FROM ICTREASW")
        Sort_grdColumns(grdICTREASW, "WHSE_REASON_CODE")

        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()

        Try
            BeginTrans()

            ASCDATA1.ExecuteSQL("DELETE FROM ICTREASW")
            dst.Tables("ICTREASW").AcceptChanges()
            For Each drICTREASW As DataRow In dst.Tables("ICTREASW").Select("")
                drICTREASW.SetAdded()
            Next

            Update_Record_TDA("ICTREASW")

            CommitTrans("Update Complete")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try
    End Sub

    Private Sub grdICTREASW_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdICTREASW.BeforeRowUpdate

        e.Row.Cells("WHSE_REASON_CODE").Value = (e.Row.Cells("WHSE_REASON_CODE").Value & String.Empty).ToString.ToUpper.Trim

        If e.Row.Cells("WHSE_REASON_CODE").Value & String.Empty = String.Empty Then
            MessageBox.Show("Whse Reason Code is required.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        Dim REASON_CODE As String = (e.Row.Cells("REASON_CODE").Value & String.Empty).ToString.ToUpper.Trim
        If REASON_CODE.Length = 0 Then
            MessageBox.Show("Reason Code is required.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        Dim drICTREAS1 As DataRow = dst.Tables("ICTREAS1").Rows.Find(REASON_CODE)
        If drICTREAS1 Is Nothing Then
            MessageBox.Show("Invalid Reason Code.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        e.Row.Cells("REASON_CODE").Value = REASON_CODE

    End Sub

    Private Sub grdICTREASW_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdICTREASW.ClickCellButton
        If grdICTREASW.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "REASON_CODE"
                sql_where = ""
            Case Else
                Exit Sub
        End Select
        grdClickCellButton(grdICTREASW, sql_where, False)
    End Sub

#End Region

End Class