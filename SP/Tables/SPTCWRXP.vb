Public Class SPTCWRXP
    Private Sub UltraLabel1_Click(sender As Object, e As EventArgs) Handles UltraLabel1.Click

    End Sub

    Private Sub UltraLabel4_Click(sender As Object, e As EventArgs) Handles UltraLabel4.Click

    End Sub

    Private Sub btnEXPINTEGRITY_Click(sender As Object, e As EventArgs) Handles btnEXPINTEGRITY.Click
        ASCMAIN1.sql = "select CHECKBOOK,SUM(CWRX_EXP_PCT) DIST from  SPTCWRXP WHERE CHECKBOOK IN ('MENS','WMEN') GROUP BY CHECKBOOK" _
         & " HAVING SUM(CWRX_EXP_PCT) <> 100"
        'Dim tblIntegrity As DataTable = ASCDATA1.GetDataTable()
        'If tblIntegrity.Rows.Count > 0 Then
        '    EMsg &= vbCr & "Order Contains Lots which are Pre-Sold"
        'End If
        Dim MESSAGE_INT As String = ""
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            MESSAGE_INT = MESSAGE_INT & row.Item("CHECKBOOK") & " Exp Distribution " & row.Item("DIST") & "% must equal 100%" & vbCr
        Next
        If MESSAGE_INT = "" Then
            MESSAGE_INT = "Check Books Exp are Distributed 100%"
        End If
        MsgBox(MESSAGE_INT, MsgBoxStyle.OkOnly, "Check Book Integrity Check")






    End Sub
End Class