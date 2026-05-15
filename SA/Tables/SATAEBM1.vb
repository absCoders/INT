Public Class SATAEBM1


    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Dim BUS_MGR_CODE As String = ASCMAIN1.Next_Control_No("SATAEBM1.BUS_MGR_CODE")
        Absx1.txtFor("BUS_MGR_CODE").Text = BUS_MGR_CODE
        Click_Command("New")
    End Sub

    Overrides Sub Show_Record_Special()
        btnAdd.Visible = False
    End Sub
    Overrides Sub Set_ScreenMode_Special(tf As Boolean)
        btnAdd.Visible = Not tf
    End Sub


End Class