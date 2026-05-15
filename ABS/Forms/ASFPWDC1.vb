Imports System.Windows.Forms
Imports System.Math

Public Class ASFPWDC1

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        ASCMAIN1.Message = ""
        Me.Close()
    End Sub

    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click

        Dim Emsg As String = String.Empty
        Dim newPassword As String = String.Empty

        txtUSER_ID.Text = txtUSER_ID.Text.Trim
        If txtUSER_ID.Text <> ASCMAIN1.USER_ID Then
            Emsg = vbCr & "Invalid User ID" & Emsg
        End If

        txtUSER_PASSWORD.Text = txtUSER_PASSWORD.Text.Trim
        If txtUSER_PASSWORD.Text.Trim <> ASCMAIN1.USER_PASSWORD Then
            Emsg = vbCr & "Invalid User Password" & Emsg
        End If

        txtNewPass.Text = txtNewPass.Text.Trim
        txtNewPassVer.Text = txtNewPass.Text.Trim

        If txtNewPass.Text <> txtNewPassVer.Text Then
            Emsg = vbCr & "New Password entries are not equal." & Emsg
        ElseIf txtNewPass.Text.Length = 0 Then
            Emsg = vbCr & "New Password entries are missing." & Emsg
        ElseIf txtNewPass.Text.Length > 25 Then
            Emsg = vbCr & "Max length of Password is 25 characters" & Emsg
        End If

        If txtNewPass.Text.ToUpper = txtUSER_PASSWORD.Text.ToUpper Then
            Emsg = vbCr & "New Password may not be the same as your current password." & Emsg
        End If

        If Emsg.Length > 0 Then
            Emsg = Emsg.Substring(1)
            MsgBox(Emsg, MsgBoxStyle.OkOnly, "Change Password Error")
            Exit Sub
        End If

        Dim tblASTPARMP As DataTable = ASCDATA1.GetDataTable("Select * from ASTPARMP where AS_PARM_KEY = 'Z'", "ASTPARMP")
        If tblASTPARMP.Rows.Count = 0 Then
            MsgBox("Table ASTPARMP is missing parameters.", MsgBoxStyle.OkOnly, "Change Password Error")
            Exit Sub
        End If

        newPassword = txtNewPass.Text

        Emsg = ASCMAIN1.Validate_User_Password(False, txtUSER_ID.Text, newPassword, tblASTPARMP.Rows(0))

        If Emsg.Length > 0 Then
            Emsg = Emsg.Substring(1)
            MsgBox(Emsg, MsgBoxStyle.OkOnly, "Change Password Error")
            Exit Sub
        End If

        If tblASTPARMP.Rows(0).Item("AS_PARM_PWD_ENCRYPTED").ToString = "1" Then
            newPassword = ASCMAIN1.EncryptAES(newPassword)
        End If

        ASCMAIN1.sql = "Update ASTUSER1 Set USER_PASSWORD = :PARM1" &
            ", USER_PASSWORD_LAST_DATE = TRUNC(SYSDATE)" &
            ", USER_PASSWORD_TEMP = NULL, USER_MUST_CHG_PWD = NULL" &
            " where USER_ID = :PARM2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {newPassword, txtUSER_ID.Text})

        ASCMAIN1.sql = "Insert Into ASTUSERP (USER_ID, INIT_DATE, USER_PASSWORD)" &
            " VALUES (:PARM1, TRUNC(SYSDATE), :PARM2)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {ASCMAIN1.USER_ID, newPassword})

        MsgBox("Password successfully changed.", MsgBoxStyle.OkOnly, "Change Password")
        ASCMAIN1.Message = "X"

        Me.Close()

    End Sub

    ' Move to a centralized location so user maintenance cmay use it
#Region "Validate User Passwords"

    Public Const ABSPWKEY = "Pm6#9&LG%<?"

    Public Structure UserPasswordError
        Dim UserID As String
        Dim UserName As String
        Dim EMsg As String
    End Structure


#End Region

    Private Sub ASFPWDC1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'UltraPictureBox1.Image = System.Drawing.Image.FromFile(ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.CLIENT_CODE & ".bmp")
        UltraPictureBox1.Image = System.Drawing.Image.FromFile(ASCMAIN1.Folders("Images") & "ABS\" & "ABS" & ".bmp")
        txtUSER_ID.Text = ASCMAIN1.USER_ID
    End Sub
End Class
