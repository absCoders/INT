Public Class GLTACCT1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GLTPARM1")


        If ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & "" = "" Then
            lblGL_PARM_SEG2_DESC.Visible = False
            optACCT_SEG2_MAND.Visible = False
        Else
            lblGL_PARM_SEG2_DESC.Text = ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC")
        End If

        If ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & "" = "" Then
            lblGL_PARM_SEG3_DESC.Visible = False
            optACCT_SEG3_MAND.Visible = False
        Else
            lblGL_PARM_SEG3_DESC.Text = ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC")
        End If

        If ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & "" = "" Then
            lblGL_PARM_SEG4_DESC.Visible = False
            optACCT_SEG4_MAND.Visible = False
        Else
            lblGL_PARM_SEG4_DESC.Text = ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC")
        End If

    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
            Case "Edit"
            Case "Update"
                If Absx1.txtFor("ACCT_DESC").Text = "" Then
                    EMsg &= vbCr & "You must enter a value for Account Description"
                End If
                If Absx1.optFor("ACCT_STATUS").Value & "" = "" Then
                    EMsg &= vbCr & "You must select a value for Account Status"
                End If
                If Absx1.optFor("ACCT_DR_CR_IND").Value & "" = "" Then
                    EMsg &= vbCr & "You must select a value for Normal Posting (DR/CR)"
                End If
                If Absx1.optFor("ACCT_TYPE").Value & "" = "" Then
                    EMsg &= vbCr & "You must select a value for Account Type"
                End If

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()

    End Sub

    Overrides Sub Proceed_Update_Special_Post()

    End Sub

    Overrides Sub Show_Record_Special()

    End Sub

    Overrides Sub Clear_Record_Special()

    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

    End Sub

#End Region

End Class