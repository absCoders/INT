Public Class ICFIRTVF

    Public RANo As String = ""
    Public TrackingNo As String = ""

    Public Sub New()
        Me.InitializeComponent()
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

        RANo = txtRANo.Text
        TrackingNo = txtTrackingNo.Text

        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class