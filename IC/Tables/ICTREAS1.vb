Imports Infragistics.Win.UltraWinGrid

Public Class ICTREAS1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
        End With

    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"

            Case "Update"

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sqlDelete = ""
    End Sub

    Overrides Sub Show_Record_Special()
        If EntryMode = "New" Then
        End If
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            EnforceConstraints(True)
        End If
    End Sub

End Class