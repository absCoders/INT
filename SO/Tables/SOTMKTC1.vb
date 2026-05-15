Public Class SOTMKTC1
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ' Create_TDA(.Tables.Add, "GLTSEGM1", "*")
        End With

    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)
        If SELECTION_NO = 0 Then Exit Sub
        Select Case eItemKey
            Case "New"
 

            Case "Edit"
            Case "Update"
                Dim MARKET_CODE As String = Absx1.txtFor("MARKET_CODE").Text
                Dim MARKET_CODE_FC = Absx1.txtFor("MARKET_CODE_FC").Text
                Dim CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                If MARKET_CODE_FC <> "" Then
                    If MARKET_CODE = MARKET_CODE_FC Then
                        EMsg &= vbCr & "Cannot use Same Market Code as FC Market Code - leave blank for same"
                    End If
                    If CUST_CODE <> "" Then
                        EMsg &= vbCr & "Cannot set up a Customer-Specific Market Code when using another Market Code for FC"
                    End If
                    Dim row As DataRow = LookUp("SOTMKTC1", MARKET_CODE_FC)
                    If row Is Nothing Then
                        EMsg &= vbCr & "Invalid Value specified for FC Market Code - " & MARKET_CODE_FC
                    Else
                        If row.Item("MARKET_CODE_FC") & "" <> "" Then
                            EMsg &= vbCr & "Cannot use a Market which itself uses another Market - " & MARKET_CODE_FC
                        End If
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