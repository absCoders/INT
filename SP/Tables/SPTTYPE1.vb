Public Class SPTTYPE1

    Private Sub SPTTYPE1_Load(sender As Object, e As EventArgs) Handles Me.Load

        If ASCMAIN1.CLIENT = "AHA" Then chkEXPENSE_TYPE_INCL_SIST.Visible = True
    End Sub
End Class