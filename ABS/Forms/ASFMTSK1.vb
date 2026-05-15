Public Class ASFMTSK1

    Private Sub ASFMTSK1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ASCMAIN1.Center(Me)
    End Sub

    Private Sub ASFMTSK1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        With dst
            Create_TDA(.Tables.Add, "ASTMTSK1", "*", 0)
            Create_TDA(.Tables.Add, "ASTMTSK2", "*", 0)

            Create_Relation("ASTMTSK1", "ASTMTSK2", "ENTITY_TYPE,ENTITY")

            .Tables.Add(ASCDATA1.GetDataTable("*", "ASTUSER1"))
        End With

        EnforceConstraints(False)

        Fill_Records("ASTMTSK1")
        Fill_Records("ASTMTSK2")

        For Each rowASTMTSK2 As DataRow In dst.Tables("ASTMTSK2").Select
            Dim ENTITY_TYPE As String = rowASTMTSK2.Item("ENTITY_TYPE") & String.Empty
            Dim ENTITY As String = rowASTMTSK2.Item("ENTITY") & String.Empty

            If dst.Tables("ASTMTSK1").Rows.Find(New String() {ENTITY_TYPE, ENTITY}) Is Nothing Then
                Dim rowASTMTSK1 As DataRow = dst.Tables("ASTMTSK1").NewRow
                rowASTMTSK1.Item("ENTITY_TYPE") = ENTITY_TYPE
                rowASTMTSK1.Item("ENTITY") = ENTITY
                dst.Tables("ASTMTSK1").Rows.Add(rowASTMTSK1)
            End If
        Next
        EnforceConstraints(True)

        grdASTMTSK1.DataSource = dst.Tables("ASTMTSK1")
        grdASTUSER1.DataSource = dst.Tables("ASTUSER1")

    End Sub

    Private Sub grdASTMTSK1_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTMTSK1.AfterRowsDeleted
        For Each row As DataRow In dst.Tables("ASTMTSK1").Select("MT_ACTION = 'O'")
            row.Item("OPEN_COUNT") = row.GetChildRows("ASTMTSK1_ASTMTSK2").Length
        Next
        ASCDATA1.DeleteRows(dst.Tables("ASTMTSK1"), "MT_ACTION = 'O' and OPEN_COUNT = 0")
        For Each row As DataRow In dst.Tables("ASTMTSK1").Select("MT_ACTION = 'L'")
            If row.GetChildRows("ASTMTSK1_ASTMTSK2").Length = 0 Then
                row.Item("OPEN_COUNT") = -1
            End If
        Next
        ASCDATA1.DeleteRows(dst.Tables("ASTMTSK1"), "MT_ACTION = 'L' and OPEN_COUNT = -1")
        Update_Record_TDA("ASTMTSK2")
        Update_Record_TDA("ASTMTSK1")
    End Sub

    Private Sub cmdClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearAll.Click
        BeginTrans()
        For Each row As DataRow In dst.Tables("ASTMTSK1").Select("")
            Dim ENTITY_TYPE As String = row.Item("ENTITY_TYPE")
            Dim ENTITY As String = row.Item("ENTITY")
            ASCMAIN1.TACMAIN1.Record_Event("ASTMTSK1", ENTITY_TYPE & ":" & ENTITY, Now, ASCMAIN1.USER_ID, "CMT", "Clear All Multi-Tasking Locks", "", "ASFMTSK1")
        Next
        Delete_Rows("ASTMTSK1", "1=1")
        Update_Record_TDA("ASTMTSK2")
        Update_Record_TDA("ASTMTSK1")
        CommitTrans()
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub grdASTMTSK1_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdASTMTSK1.BeforeRowsDeleted
        e.DisplayPromptMsg = False
        Dim ENTITY_TYPE As String = e.Rows(0).Cells("ENTITY_TYPE").Value
        Dim ENTITY As String = e.Rows(0).Cells("ENTITY").Value
        ASCMAIN1.TACMAIN1.Record_Event("ASTMTSK1", ENTITY_TYPE & ":" & ENTITY, Now, ASCMAIN1.USER_ID, "CMT", "Clear Lock", "", "ASFMTSK1")
    End Sub
End Class