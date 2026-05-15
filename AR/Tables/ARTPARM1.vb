Public Class ARTPARM1
#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
            Case "Edit"

            Case "Update"

                If Absx1.chkFor("AR_PARM_USE_DISC").Checked Then
                    If Absx1.txtFor("AR_PARM_HDG_DISC").Text = "" Then
                        EMsg &= vbCr & "Please provide a Heading for Discounts"
                    End If
                    If LookUp("ARTREAS1", Absx1.txtFor("AR_PARM_REASON_CODE_DISC").Text) Is Nothing Then
                        EMsg &= vbCr & "Please provide a Valid Reason Code for Discounts"
                    End If
                End If
                If Absx1.chkFor("AR_PARM_USE_WOFF").Checked Then
                    If Absx1.txtFor("AR_PARM_HDG_WOFF").Text = "" Then
                        EMsg &= vbCr & "Please provide a Heading for Write-Offs"
                    End If
                    If LookUp("ARTREAS1", Absx1.txtFor("AR_PARM_REASON_CODE_WOFF").Text) Is Nothing Then
                        EMsg &= vbCr & "Please provide a Valid Reason Code for Write-Offs"
                    End If
                End If

        End Select
    End Sub

    Private Sub btnWHTSCL_Click(sender As Object, e As EventArgs) Handles btnWHTSCL.Click
        ShowUsersWithThisSecurity(Absx1.txtFor("AR_PARM_SEC_ISSUE_CRD").Text)
    End Sub

    Private Sub btnWHTSCC_Click(sender As Object, e As EventArgs) Handles btnWHTSCC.Click
        ShowUsersWithThisSecurity(Absx1.txtFor("AR_PARM_SEC_CREDIT_CRD").Text)
    End Sub


    Private Sub ShowUsersWithThisSecurity(ByVal SECURITY_CODE As String)
        Try
            If Not ScreenMode Then
                Exit Sub
            End If

            SECURITY_CODE = SECURITY_CODE.Trim
            SECURITY_CODE = SECURITY_CODE.Replace("'", "")
            If SECURITY_CODE.Length = 0 Then
                Exit Sub
            End If

            ASCMAIN1.sql = $"SELECT ASTUSER1.USER_ID, ASTUSER1.USER_NAME
                                FROM ASTUSER1, ASTUSER2
                                WHERE ASTUSER1.USER_ID = ASTUSER2.USER_ID
                                AND ASTUSER2.SECURITY_CODE = :PARM1
                                ORDER BY USER_NAME"

            Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {SECURITY_CODE})
            If tbl.Rows.Count = 0 Then
                MessageBox.Show($"No users have this security code", "Who has this security", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Using frmASFMSGBF As New ASFMSGBF
                frmASFMSGBF.gridColumnCaptions = "User ID, Username"
                frmASFMSGBF.Show_grd(tbl, Me, $"Users with Security Code {SECURITY_CODE}")
                frmASFMSGBF.Dispose()
            End Using

        Catch ex As Exception
            MessageBox.Show($"Error showing users with this security {ex.Message}", "Who has this security", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

#End Region

End Class