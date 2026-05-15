Public Class APRCHKR1

    Dim APTCHKR1 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Get_PARM("GLTPARM1")
        Call Get_PARM("APTPARM1")
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        APTCHKR1 = TAC.APCMAIN1.Prepare_Check_Register(Me, dst, True)

        'If dst Is Nothing Then
        '    Stop
        'End If
        Check_if_Empty("APTCHKR1")
    End Sub

    Public Overrides Sub Print_Report()
        Generate_Report(RPT)
        Call Print_GL()
    End Sub

    Overrides Sub Update_Record()

        Dim sql As String

        sql = "Update APTCHCK1 " _
        & " Set REGISTER_IND = '1'" _
        & ", REGISTER_XNO = '" & XNO & "'" _
        & " where (BANK_CODE, CHECK_NUM) in " _
        & "(Select BANK_CODE, CHECK_NUM from " & APTCHKR1 _
        & " where RECORD_TYPE = 'I')"
        ASCDATA1.ExecuteSQL(sql)

        sql = "Update APTCHCK1 " _
        & " Set REGISTER_IND_F = '1'" _
        & ", REGISTER_XNO_F = '" & XNO & "'" _
        & " where (BANK_CODE, CHECK_NUM) in " _
        & "(Select BANK_CODE, CHECK_NUM from " & APTCHKR1 _
        & " where RECORD_TYPE = 'V')"
        ASCDATA1.ExecuteSQL(sql)

        GL_Update()

        If ASCMAIN1.DBS_SERVER = "EXP" Or ASCMAIN1.DBS_COMPANY = "EXP" Then
            For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("GLTINTF1"), New String() {"JOURNAL_NO"}).Select("")
                Dim JOURNAL_NO As String = row.Item(0)
                ASCMAIN1.sql = "Insert into GLTDETL1_OBX Select GLTDETL1.*, NULL DATETIME_STAMP, 'APCD' JOURNAL_TYPE from GLTDETL1 where JOURNAL_NO = '" & JOURNAL_NO & "'"
                ASCDATA1.ExecuteSQL()
            Next
        End If

    End Sub
End Class