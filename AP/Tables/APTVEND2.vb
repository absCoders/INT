Public Class APTVEND2


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        grpBankingInfo.Visible = (ASCMAIN1.USER_SECURITY_CODEs.Split(",").Contains("P3"))

    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Update"

                Dim VEND_ALT_BANK_ROUTING_NO As String = Absx1.txtFor("VEND_ALT_BANK_ROUTING_NO").Text
                Dim VEND_ALT_BANK_SWIFT_NO As String = Absx1.txtFor("VEND_ALT_BANK_SWIFT_NO").Text
                Dim VEND_ALT_BANK_ACCT_ID As String = Absx1.txtFor("VEND_ALT_BANK_ACCT_ID").Text
                Dim VEND_ALT_BANK_COUNTRY As String = Absx1.txtFor("VEND_ALT_BANK_COUNTRY").Text
                Dim VEND_ALT_BANK_ACCT_CLASS As String = Absx1.optFor("VEND_ALT_BANK_ACCT_CLASS").Value & ""
                Dim VEND_ALT_BANK_ACCT_TYPE As String = Absx1.optFor("VEND_ALT_BANK_ACCT_TYPE").Value & ""

                If VEND_ALT_BANK_ROUTING_NO <> "" Or VEND_ALT_BANK_ACCT_ID <> "" Or VEND_ALT_BANK_SWIFT_NO <> "" Then
                    If VEND_ALT_BANK_ACCT_ID = "" Then EMsg &= vbCr & "Bank Account No is Mandatory"
                    If VEND_ALT_BANK_COUNTRY = "" Then
                        EMsg &= vbCr & "Bank Country is Mandatory"
                    Else
                        Dim rowTATCNTRY As DataRow = LookUp("TATCNTRY", VEND_ALT_BANK_COUNTRY)
                        If rowTATCNTRY Is Nothing Then
                            EMsg &= vbCr & "Invalid Value specified for Bank Country"
                        Else
                            If VEND_ALT_BANK_COUNTRY = "USA" Then
                                If VEND_ALT_BANK_ROUTING_NO = "" Then
                                    EMsg &= vbCr & "You must have a Routing No for a US Bank"
                                Else
                                    If Format(Val(VEND_ALT_BANK_ROUTING_NO), "000000000") <> VEND_ALT_BANK_ROUTING_NO Then
                                        EMsg &= vbCr & "Bank Routing No should be 9 digits all numeric"
                                    End If
                                End If
                                If VEND_ALT_BANK_SWIFT_NO <> "" Then EMsg &= vbCr & "You cannot have a Swift No for a US Bank"
                            Else
                                If VEND_ALT_BANK_ROUTING_NO <> "" Then EMsg &= vbCr & "You cannot have a Routing No for a non-US Bank"
                                If VEND_ALT_BANK_SWIFT_NO = "" Then
                                    EMsg &= vbCr & "You must have a Swift No for a non-US Bank"
                                Else
                                    'https://stackoverflow.com/questions/3028150/what-is-proper-regex-expression-for-swift-codes
                                    ' Dim rx As String = "[A-Z]{6,6}[A-Z2-9][A-NP-Z0-9]([A-Z0-9]{3,3}){0,1}" from JPMC spec
                                    Dim rx As String = "^[A-Z]{6}[A-Z0-9]{2}([A-Z0-9]{3})?$"
                                    Dim r As New System.Text.RegularExpressions.Regex(rx)
                                    ' RBOSGB2L
                                    If r.IsMatch(VEND_ALT_BANK_SWIFT_NO) Then
                                    Else
                                        EMsg &= vbCr & $"{VEND_ALT_BANK_SWIFT_NO} has Special Characters which are not allowed"
                                        EMsg &= vbCr & "A swift code should be 8 or 11 letters or digits where the first six must be letters."
                                    End If
                                End If
                            End If
                        End If
                    End If

                    If VEND_ALT_BANK_ACCT_CLASS = "" Then EMsg &= vbCr & "Bank Account Class is Mandatory"
                    If VEND_ALT_BANK_ACCT_TYPE = "" Then EMsg &= vbCr & "Bank Account Type is Mandatory"
                End If
        End Select

    End Sub
End Class