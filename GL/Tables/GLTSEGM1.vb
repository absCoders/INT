Public Class GLTSEGM1
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("GLTPARM1")
    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)
        If SELECTION_NO = 0 Then Exit Sub
        Select Case eItemKey
            Case "New"
            Case "Edit"
            Case "Update"

                If Absx1.txtFor("ACCT_SEG_DESC").Text & "" = "" Then
                    EMsg &= vbCr & "Segment Description is Mandatory"
                End If

                Dim ACCT_SEG_ID As String = Absx1.txtFor("ACCT_SEG_ID").Text
                Dim SC As String = ROWs("GLTPARM1").Item("GL_PARM_SEG" & ACCT_SEG_ID & "_CLASS_DESC") & ""
                If SC <> "" Then
                    Dim ACCT_SEG_CLASS As String = Absx1.txtFor("ACCT_SEG_CLASS").Text
                    If ACCT_SEG_CLASS = "" Then
                        EMsg &= vbCr & "Segment Class is Mandatory (" & SC & ")"
                    End If
                End If

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()

    End Sub

    Overrides Sub Proceed_Update_Special_Post()

    End Sub
#End Region
End Class