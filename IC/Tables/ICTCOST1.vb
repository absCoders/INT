Public Class ICTCOST1

    Private Sub ICTCOST1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Absx1.chkFor("EXP_AT_PURCHASE").Visible = Not (ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT")
    End Sub
End Class