Public Class ARTREAS1

    Private Sub ARTREAS1_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Not ASCMAIN1.CLIENT = "AHA" Then
        Else
            chkShippingViolation.Visible = False
        End If

        grpSegments.Visible = False
    End Sub
End Class