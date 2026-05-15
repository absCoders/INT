Public Class ASFTEXT1
    Public t As String

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub VCA_TEXT_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        t = Replace(t, vbCrLf, vbCr)
        t = Replace(t, vbLf, vbCr)
        t = Replace(t, vbCr, vbCrLf)
        txt.Text = t
    End Sub

End Class