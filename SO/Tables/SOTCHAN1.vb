Public Class SOTCHAN1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Get_PARM("GLTPARM1")

        With dst
            Create_TDA(.Tables.Add, "GLTSEGM1", "*")
        End With

        lblCUST_CODE.Visible = (ASCMAIN1.CLIENT = "INT")
        txtCUST_CODE.Visible = (ASCMAIN1.CLIENT = "INT")
        txtCUST_NAME.Visible = (ASCMAIN1.CLIENT = "INT")

        lblSEG2_CODE.Visible = Not (ASCMAIN1.CLIENT = "INT")
        txtSEG2_CODE.Visible = Not (ASCMAIN1.CLIENT = "INT")
        txtSEG2_DESC.Visible = Not (ASCMAIN1.CLIENT = "INT")
    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
                If LookUp("SOTTCLS1", Absx1.txtFor("CHANNEL_CODE").Text) IsNot Nothing Then
                    EMsg &= vbCr & "Cannot have a Trade Class Code that is the same as a Channel Code"
                End If

            Case "Edit"
            Case "Update"

                'If Absx1.optFor("CHANNEL_STATUS").Value & "" = "" Then
                '    EMsg &= vbCr & "Status is Mandatory"
                'End If
                If Absx1.txtFor("CHANNEL_DESC").Text & "" = "" Then
                    EMsg &= vbCr & "Channel Name is Mandatory"
                End If
                If LookUp("SOTTCLS1", Absx1.txtFor("CHANNEL_CODE").Text) IsNot Nothing Then
                    EMsg &= vbCr & "Cannot have a Trade Class Code that is the same as a Channel Code"
                End If

                If ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & "" = "" _
                    OrElse ROWs("SOTPARM1").Item("SO_PARM_DTL_SEG3") & "" <> "1" Then
                    ' NO NEED TO CHECK SEG4
                Else
                    'Dim SEG3_CODE As String = Absx1.txtFor("SEG3_CODE").Text
                    'If SEG3_CODE <> "" And SEG3_CODE <> Absx1.txtFor("CHANNEL_CODE").Text Then
                    '    If LookUp("SOTTCLS1", SEG3_CODE) Is Nothing _
                    '        AndAlso LookUp("SOTCHAN1", SEG3_CODE) Is Nothing _
                    '        AndAlso LookUp("GLTSEGM1", New String() {"3", SEG3_CODE}) Is Nothing Then
                    '        EMsg &= vbCr & "Invalid Value Specified for Segment 3"
                    '    End If
                    'End If
                End If
        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        'Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        'Dim sqlDelete = "CUST_CODE = '" & CUST_CODE & "'"
        'Update_Record_TDA("SPTDCOM2", sqlDelete)
    End Sub

    Overrides Sub Proceed_Update_Special_Post()
        ' NOTE SOTCHAN1 DOES NOT (YET) HAVE OVERRIDING SEG3_CODE
        'Dim SEG3_CODE As String = Absx1.txtFor("SEG3_CODE").Text
        'If SEG3_CODE = "" Then
        '    SEG3_CODE = Absx1.txtFor("CHANNEL_CODE").Text
        'End If
        Dim SEG3_CODE As String = Absx1.txtFor("CHANNEL_CODE").Text
        Dim row As DataRow = LookUp("GLTSEGM1", New String() {"3", SEG3_CODE})
        If row Is Nothing Then
            Dim rowGLTSEGM1 As DataRow = dst.Tables("GLTSEGM1").NewRow
            rowGLTSEGM1.Item("ACCT_SEG_ID") = "3"
            rowGLTSEGM1.Item("ACCT_SEG_CODE") = SEG3_CODE
            rowGLTSEGM1.Item("ACCT_SEG_STATUS") = "A"
            rowGLTSEGM1.Item("ACCT_SEG_CLASS") = Absx1.txtFor("CHANNEL_CODE").Text
            rowGLTSEGM1.Item("ACCT_SEG_DESC") = Absx1.txtFor("CHANNEL_DESC").Text
            dst.Tables("GLTSEGM1").Rows.Add(rowGLTSEGM1)
            Update_Record_TDA("GLTSEGM1")
        End If
    End Sub
#End Region
     
End Class