Public Class SOTTCLS1
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "GLTSEGM1", "*")
        End With

    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)
        If SELECTION_NO = 0 Then Exit Sub
        Select Case eItemKey
            Case "New"
                If LookUp("SOTCHAN1", Absx1.txtFor("TRADE_CLASS_CODE").Text) IsNot Nothing Then
                    EMsg &= vbCr & "Cannot have a Trade Class Code that is the same as a Channel Code"
                End If

            Case "Edit"
            Case "Update"

                If LookUp("SOTCHAN1", Absx1.txtFor("TRADE_CLASS_CODE").Text) IsNot Nothing Then
                    EMsg &= vbCr & "Cannot have a Trade Class Code that is the same as a Channel Code"
                End If
                If Absx1.txtFor("TRADE_CLASS_DESC").Text & "" = "" Then
                    EMsg &= vbCr & "Trade Class Description is Mandatory"
                End If

                Dim CHANNEL_CODE As String = Absx1.txtFor("CHANNEL_CODE").Text
                If CHANNEL_CODE = "" Then
                    EMsg &= vbCr & "Channel Code is Mandatory"
                Else
                    If LookUp("SOTCHAN1", CHANNEL_CODE) Is Nothing Then
                        EMsg &= vbCr & "Invalid Value Specified for Channel"
                    End If
                End If

                Dim SEG3_CODE As String = Absx1.txtFor("SEG3_CODE").Text
                If SEG3_CODE <> "" And SEG3_CODE <> Absx1.txtFor("TRADE_CLASS_CODE").Text Then
                    If LookUp("SOTTCLS1", SEG3_CODE) Is Nothing Then
                        EMsg &= vbCr & "Invalid Value Specified for Segment 3"
                    End If
                End If

                'If SEG3_CODE = "" Then
                '    Absx1.txtFor("SEG3_CODE").Text = Absx1.txtFor("TRADE_CLASS_CODE").Text
                '    Synch_TABLE_NAME("SOTTCLS1")
                'End If
        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()

    End Sub

    Overrides Sub Proceed_Update_Special_Post()
        Dim SEG3_CODE As String = Absx1.txtFor("SEG3_CODE").Text
        If SEG3_CODE = "" Then
            SEG3_CODE = Absx1.txtFor("TRADE_CLASS_CODE").Text
        End If
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