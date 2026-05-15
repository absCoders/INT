Imports Infragistics.Win.UltraWinGrid

Public Class ICTREASW

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
        End With

    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"

            Case "Update"
                Dim REASON_CODE As String = Absx1.txtFor("REASON_CODE").Text.Trim

                If REASON_CODE.Length = 0 Then
                    EMsg &= vbCr & "Reason Code is required."
                Else
                    Dim drICTREAS1 As DataRow = LookUp("ICTREAS1", REASON_CODE)
                    If drICTREAS1 Is Nothing Then
                        EMsg &= vbCr & "Reason Code is invalid."
                    End If
                End If

                Absx1.txtFor("WHSE_REASON_CODE").Text = Absx1.txtFor("WHSE_REASON_CODE").Text.Trim.ToUpper

                ' Need to have a key that does not have spaces. This value is hidden on the screen
                If Absx1.txtFor("WHSE_REASON_CTL").TextLength = 0 Then
                    Absx1.txtFor("WHSE_REASON_CTL").Text = ASCMAIN1.Next_Control_No("ICTREASW.WHSE_REASON_CTL")
                End If

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