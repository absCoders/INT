Public Class SORSORDX

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "U"

    End Sub

    Public Overrides Sub Print_Report()
        'Generate_Report(RPT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            If ASCDATA1.GetDataValue("Select PRD_CLOSE_IND from ASTPCTL1") & "" = "1" Then
                EMsg = EMsg & vbCr & "Period-End Already Initialized"

            End If

        End If

    End Sub

    Overrides Sub Update_Record()

    End Sub
End Class