Public Class ASTPARMP

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        cmdInitializeEncryption.Visible = False
        If ASCMAIN1.Running_in_VS Then
            cmdInitializeEncryption.Visible = True
        End If

        Dim rowASTPARMP As DataRow = LookUp("ASTPARMP", "Z")
        If rowASTPARMP.Item("AS_PARM_PWD_ENCRYPTED") & "" = "1" Then
            cmdInitializeEncryption.Visible = False
        End If

    End Sub
    Public Overrides Sub Proceed_PreReq_Special(eItemKey As String)
        MyBase.Proceed_PreReq_Special(eItemKey)

        Absx1.chkFor("AS_PARM_USE_ENCRYPTION").Enabled = False
        Absx1.chkFor("AS_PARM_PWD_ENCRYPTED").Enabled = False

    End Sub

    Public Overrides Sub Mode_Settings(tf As Boolean, Optional MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        Absx1.chkFor("AS_PARM_USE_ENCRYPTION").Enabled = False
        Absx1.chkFor("AS_PARM_PWD_ENCRYPTED").Enabled = False
        Set_Read_Only_for_ctl(Absx1.txtFor("AS_PARM_SEC_ALERT_EMAIL"), True)
        Set_Read_Only_for_ctl(Absx1.txtFor("AS_PARM_USERDOMAIN"), True)

    End Sub

    Sub IntializeEncryption()

        Dim rowASTPARMP As DataRow = LookUp("ASTPARMP", "Z")
        If rowASTPARMP.Item("AS_PARM_PWD_ENCRYPTED") & "" = "1" Then
            MsgBox("Password Encryption is already enabled")
            Exit Sub
        End If

        BeginTrans()

        Try

            Create_TDA(dst.Tables.Add, "ASTAESC1", "*", 0)
            Fill_Records("ASTAESC1")
            Dim rowAETAESC1 As DataRow = Nothing
            Dim ASTAESC1s As New List(Of String)
            ASTAESC1s.Add("ASTUSER1.USER_PASSWORD")
            ASTAESC1s.Add("ASTPARM1.AS_PARM_EMAIL_PASSWORD")
            ASTAESC1s.Add("ASTPARM1.AS_PARM_EFAX_PASSWORD")

            For Each TC As String In ASTAESC1s
                Dim TABLE_NAME As String = Split(TC, ".")(0)
                Dim COLUMN_NAME As String = Split(TC, ".")(1)
                rowAETAESC1 = dst.Tables("ASTAESC1").Rows.Find(New String() {TABLE_NAME, COLUMN_NAME})
                If rowAETAESC1 Is Nothing Then dst.Tables("ASTAESC1").Rows.Add(New String() {TABLE_NAME, COLUMN_NAME})

            Next
            Update_Record_TDA("ASTAESC1")

            Create_TDA(dst.Tables.Add, "ASTPARM1", "*", 0)
            Fill_Records("ASTPARM1")
            For Each row As DataRow In dst.Tables("ASTPARM1").Select("")
                Dim AS_PARM_EMAIL_PASSWORD As String = row.Item("AS_PARM_EMAIL_PASSWORD") & ""
                AS_PARM_EMAIL_PASSWORD = ASCMAIN1.EncryptAES(AS_PARM_EMAIL_PASSWORD)
                row.Item("AS_PARM_EMAIL_PASSWORD") = AS_PARM_EMAIL_PASSWORD

                Dim AS_PARM_EFAX_PASSWORD As String = row.Item("AS_PARM_EFAX_PASSWORD") & ""
                AS_PARM_EFAX_PASSWORD = ASCMAIN1.EncryptAES(AS_PARM_EFAX_PASSWORD)
                row.Item("AS_PARM_EFAX_PASSWORD") = AS_PARM_EFAX_PASSWORD
            Next
            Update_Record_TDA("ASTPARM1")

            Create_TDA(dst.Tables.Add, "ASTUSER1", "*", 0)
            Fill_Records("ASTUSER1")
            For Each row As DataRow In dst.Tables("ASTUSER1").Select("")
                Dim USER_PASSWORD As String = row.Item("USER_PASSWORD") & ""
                USER_PASSWORD = ASCMAIN1.EncryptAES(USER_PASSWORD)
                row.Item("USER_PASSWORD") = USER_PASSWORD
            Next
            Update_Record_TDA("ASTUSER1")

            Create_TDA(dst.Tables.Add, "ASTUSERP", "*", 0)
            Fill_Records("ASTUSERP")
            For Each row As DataRow In dst.Tables("ASTUSERP").Select("")
                Dim USER_PASSWORD As String = row.Item("USER_PASSWORD") & ""
                USER_PASSWORD = ASCMAIN1.EncryptAES(USER_PASSWORD)
                row.Item("USER_PASSWORD") = USER_PASSWORD
                row.AcceptChanges()
                row.SetAdded()
            Next
            ASCMAIN1.sql = "Delete from ASTUSERP"
            ASCDATA1.ExecuteSQL()
            Update_Record_TDA("ASTUSERP")

            ' Create_TDA(dst.Tables.Add, "ASTAUDT1", "*", 0)
            ASCMAIN1.sql = "Select * from ASTAUDT1 where COLUMN_NAME Like '%PASSWORD%' or COLUMN_NAME LIKE '%_PWD'"
            Fill_Records("ASTAUDT1",,, ASCMAIN1.sql)
            For Each row As DataRow In dst.Tables("ASTAUDT1").Select("")
                Dim OLD_VALUE As String = row.Item("OLD_VALUE") & ""
                If OLD_VALUE <> "" Then
                    OLD_VALUE = ASCMAIN1.EncryptAES(OLD_VALUE)
                    row.Item("OLD_VALUE") = OLD_VALUE
                End If
                Dim NEW_VALUE As String = row.Item("NEW_VALUE") & ""
                If NEW_VALUE <> "" Then
                    NEW_VALUE = ASCMAIN1.EncryptAES(NEW_VALUE)
                    row.Item("NEW_VALUE") = NEW_VALUE
                    row.AcceptChanges()
                    row.SetAdded()
                End If
            Next
            ASCMAIN1.sql = "Delete from ASTAUDT1 where COLUMN_NAME like '%PASSWORD%' or COLUMN_NAME LIKE '%_PWD'"
            ASCDATA1.ExecuteSQL()
            Update_Record_TDA("ASTAUDT1")


            ASCMAIN1.sql = "Update ASTPARMP Set AS_PARM_PWD_ENCRYPTED = '1'"
            ASCDATA1.ExecuteSQL()

            CommitTrans("Password Encryption has been Initialized")

        Catch ex As Exception

        End Try

    End Sub

    Private Sub cmdInitializeEncryption_Click(sender As Object, e As EventArgs) Handles cmdInitializeEncryption.Click
        If Not ASCMAIN1.Running_in_VS Then
            Exit Sub
        End If
        IntializeEncryption()
    End Sub

End Class