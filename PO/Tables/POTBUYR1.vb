Public Class POTBUYR1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            'ASCMAIN1.sql = "Select APTVEND9.*, GLTACCT1.ACCT_DESC from APTVEND9,GLTACCT1 where GLTACCT1.ACCT_CODE = APTVEND9.ACCT_CODE and APTVEND9.VEND_CODE = :PARM1"
            'Create_TDA(.Tables.Add, "APTVEND9", "**", 0, True, "V", 5)
        End With
      
    End Sub

#Region "Overrides"

    Overrides Sub Proceed_Update_Special_Pre()

    End Sub

    Overrides Sub Show_Record_Special()
    End Sub

    Sub Load_Report_Form(ByVal VEND_CODE As String)

    End Sub

    Overrides Sub Clear_Record_Special()

    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        cmdAddUser.Visible = Not tf

    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Edit"
                If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                    If Absx1.txtFor("VEND_BUYER_CODE").Text = ASCMAIN1.USER_ID Then
                        EMsg &= EMsg & "You may NOT maintain your own record"
                    End If
                End If

            Case "Update"
                    Dim row As DataRow = LookUp("ASTUSER1", Absx1.txtFor("VEND_BUYER_CODE").Text)
                    If row Is Nothing Then
                        EMsg &= EMsg & "Buyer Code must be a valid User ID"
                    End If
        End Select

    End Sub
#End Region
     Private Sub cmdAddUser_Click(sender As Object, e As EventArgs) Handles cmdAddUser.Click
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("USER_ID")

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = False
            ASCMAIN1.CodeSelector.Custom_sql_where = " and USER_ID Not in (Select VEND_BUYER_CODE from POTBUYR1)"
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Absx1.txtFor("VEND_BUYER_CODE").Text = ASCMAIN1.CodeSelector.SelectedCode
                Click_Command("New")
                Absx1.txtFor("VEND_BUYER_NAME").Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("USER_NAME")
            End If
        End If
    End Sub
End Class