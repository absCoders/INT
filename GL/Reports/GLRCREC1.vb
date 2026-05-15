Public Class GLRCREC1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GLTPARM1")
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, -1)

    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim TT As String = GL_Prep(Mid$(RYP, 1, 4), Mid$(RYP, 1, 4))
        Dim pp As Integer = Val(Mid$(RYP, 5, 2))

        sql = "Select ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE, ACCT_BEG_BAL"
        For i As Integer = 1 To pp
            sql = sql & " + ACCT_ACT_P" & Format(i, "00")
        Next i
        sql = sql & " AMT_GL from " & TT
        sql = sql & " where ACCT_YEAR = '" & Mid$(RYP, 1, 4) & "'"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTCRECG", 4))

        Get_WKCodes("GLTCRECG", "ACCT_CODE", "GLTACCT1", "*")

        Dim sqlw As String = ""
        sqlw &= SQLA_filter("CREC_TYPE_CODE")

        sql = "Select GLTCREC1.* from GLTCREC1" & ASCMAIN1.SQL_Add_WHERE(sqlw)
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTCREC1", 1))

        sql = "Select GLTCREC2.* from GLTCREC2" & ASCMAIN1.SQL_Add_WHERE(sqlw)
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTCREC2", 2))

        sql = "Select * from GLTCREC4"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTCREC4", 2))

        If Absx1.chkFor("SHOW_CVX").Checked Then
            sql = "Select * from GLTCREC3 where OPS_YYYYPP = '" & RYP & "'" & sqlw
        Else
            sql = "SELECT OPS_YYYYPP, CREC_TYPE_CODE, CREC_CLASS_CODE, ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE" _
            & ", 'X' DETL_CTL_TYPE, 'X' DETL_CTL_NO, 'X' DETL_CVX_TYPE, 'X' DETL_CVX_NO, SUM (CREC_AMT) CREC_AMT" _
            & " from GLTCREC3 where OPS_YYYYPP = '" & RYP & "'" & sqlw _
            & " group by OPS_YYYYPP, CREC_TYPE_CODE, CREC_CLASS_CODE, ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE"
        End If
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTCREC3", 0))
    End Sub

    Public Overrides Sub Print_Report()
        CR_params.Add("SHOW_CVX", IIf(Absx1.chkFor("SHOW_CVX").Checked, "1", "0"))
        SUBT = Absx1.cmbFor("RYP").Text
        Generate_Report(RPT, , SUBT)
    End Sub
End Class