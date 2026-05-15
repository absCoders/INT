Public Class SOTPARM1

    Private Sub SOTPARM1_Load(sender As Object, e As EventArgs) Handles Me.Load
        'If ASCMAIN1.CLIENT = "INT" Then
        'grpPriceSecurity.Visible = False
        'End If

        ' 2/15/2017 - I set the Hasbutton on the two controls on grpPriceSecurity to false. Was naot able to Update without getting error messages that the fields had wrong values. EWZ
    End Sub

    Sub One_Time_Update()

        Dim CYPdt As Date = CDate("10/31/2017")

        ASCMAIN1.sql = "Select * from SOTPRIC2" & vbCrLf _
            & " where ITEM_NEW_PRICE_DATE <= '" & Format(CYPdt.AddDays(1), "dd-MMM-yyyy") & "'"
        Fill_Records("SOTPRIC2", "", True, ASCMAIN1.sql)
        If dst.Tables("SOTPRIC2").Rows.Count <> 0 Then
            For Each rowSOTPRIC2 As DataRow In dst.Tables("SOTPRIC2").Select("")
                DATETIME_STAMP = CDate(rowSOTPRIC2.Item("ITEM_NEW_PRICE_DATE"))
                rowSOTPRIC2.Item("ITEM_PRICE") = rowSOTPRIC2.Item("ITEM_NEW_PRICE")
                rowSOTPRIC2.Item("ITEM_NEW_PRICE") = DBNull.Value
                rowSOTPRIC2.Item("ITEM_NEW_PRICE_DATE") = DBNull.Value
                Write_Audit_Trail(rowSOTPRIC2, "E")
            Next
            Update_Record_TDA("SOTPRIC2")
        End If

        ASCMAIN1.sql = "Select * from SOTPRIC2" & vbCrLf _
            & "  where ITEM_NEW_SRP_DATE <= '" & Format(CYPdt.AddDays(1), "dd-MMM-yyyy") & "'"
        Fill_Records("SOTPRIC2", "", True, ASCMAIN1.sql)
        If dst.Tables("SOTPRIC2").Rows.Count <> 0 Then
            For Each rowSOTPRIC2 As DataRow In dst.Tables("SOTPRIC2").Select("")
                DATETIME_STAMP = CDate(rowSOTPRIC2.Item("ITEM_NEW_SRP_DATE"))
                rowSOTPRIC2.Item("ITEM_SRP") = rowSOTPRIC2.Item("ITEM_NEW_SRP")
                rowSOTPRIC2.Item("ITEM_NEW_SRP") = DBNull.Value
                rowSOTPRIC2.Item("ITEM_NEW_SRP_DATE") = DBNull.Value
                Write_Audit_Trail(rowSOTPRIC2, "E")
            Next
            Update_Record_TDA("SOTPRIC2")
        End If

        MsgBox("Update Complete", MsgBoxStyle.OkOnly)
    End Sub

    Private Sub UltraButton1_Click(sender As Object, e As EventArgs) Handles UltraButton1.Click
        If MsgBox("OK to update all Prices prior to 11/01/2017?", MsgBoxStyle.OkCancel, "Verification") = MsgBoxResult.Cancel Then
            Exit Sub
        End If

        With dst
            ASCMAIN1.sql = "Select SOTPRIC2.* from SOTPRIC2"
            Create_TDA(.Tables.Add, "SOTPRIC2", "**", 0, True, "", 2)
        End With

        One_Time_Update()
    End Sub
End Class