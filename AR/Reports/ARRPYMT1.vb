Public Class ARRPYMT1

    Dim ARTPYMT0 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Get_PARM("GLTPARM1")
        Call Get_PARM("ARTPARM1")
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        Dim sqlw As String = ""

        sqlw &= Get_Filter("PYMT_BATCH_NO", "ARTPYMT1.PYMT_BATCH_NO")

        If chkMYBATCHESONLY.Checked Then
            sqlw &= "   and ARTPYMT1.LAST_OPER = '" & ASCMAIN1.USER_ID & "'"
        End If

        ASCMAIN1.sql = "Select PYMT_BATCH_NO from ARTPYMT1 where STATUS = '0'" & sqlw

        ARTPYMT0 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ARTPYMT0 & " Add Primary Key (PYMT_BATCH_NO)")
        ASCMAIN1.AnalyzeTable(ARTPYMT0)

        With dst
            sql = "Select ARTPYMT1.* from ARTPYMT1," & ARTPYMT0 & " ARTPYMT0 where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT0.PYMT_BATCH_NO"
            .Tables.Add(ASCDATA1.GetDataTable(sql, "ARTPYMT1", 1))

            sql = "Select ARTPYMT2.* from ARTPYMT2," & ARTPYMT0 & " ARTPYMT0 where ARTPYMT2.PYMT_BATCH_NO = ARTPYMT0.PYMT_BATCH_NO"
            .Tables.Add(ASCDATA1.GetDataTable(sql, "ARTPYMT2", 2))

            .Tables.Add(ASCDATA1.GetDataTable("*", "GLTBANK1"))
            .Tables.Add(ASCDATA1.GetDataTable("*", "GLTACCT1"))
            .Tables.Add(ASCDATA1.GetDataTable("*", "ARTPOST1"))
        End With

        Check_if_Empty("ARTPYMT1")
    End Sub

    Public Overrides Sub Print_Report()
        CR_params.Add("GL_PARM_CURR_CODE", ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE"))

        Generate_Report(RPT)
    End Sub

    Overrides Sub Update_Record()

        Dim sql As String

        sql = "Update ARTPYMT1 Set STATUS = '1'" _
        & " where PYMT_BATCH_NO in " _
        & "(Select PYMT_BATCH_NO from " & ARTPYMT0 & ")"
        ASCDATA1.ExecuteSQL(sql)

        sql = "Update ARTPYMT2 Set PYMT_STATUS = '1'" _
        & " where PYMT_BATCH_NO in " _
        & "(Select PYMT_BATCH_NO from " & ARTPYMT0 & ")"
        ASCDATA1.ExecuteSQL(sql)

    End Sub
End Class