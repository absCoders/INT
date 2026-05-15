Public Class APRAEXP2

    Dim APTAEXP2 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Get_PARM("GLTPARM1")
        Call Get_PARM("APTPARM1")

        Set_cmbYP("RYP", ASCMAIN1.CYP, -36, -1, -1)
    End Sub

    Protected Overrides Sub Build_Workfile()
        MyBase.Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        Dim sql As String

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        With dst
            sql = "SELECT APTINVH2.*" _
            & ", APTINVH1.VEND_CODE, APTVEND1.VEND_NAME" _
            & ", APTINVH1.INV_NUM, APTINVH1.INV_DATE" _
            & ", APTVEND1.PROCESSOR_CODE, APTVEND1.VEND_TYPE" _
            & ", DECODE(APTINVH2.INV_COMMENT_DTL, NULL, APTINVH1.INV_REF, APTINVH2.INV_COMMENT_DTL) DTL_DESC" _
            & " from APTINVH1, APTINVH2, APTVEND1 " _
            & " where APTINVH1.VOUCHER_NO = APTINVH2.VOUCHER_NO" _
            & "   and APTINVH1.OPS_YYYYPP_ACCRUE " & IIf(Absx1.optFor("OPTXO").Value = "X", "=", "<=") & " '" & RYP & "'" _
            & "   and APTVEND1.VEND_CODE = APTINVH1.VEND_CODE" _
            & "   and APTINVH2.INV_LTYP IS NULL" _
            & "   and NVL(APTINVH1.REGISTER_IND,'0') <> 'D'" _
            & "   and APTINVH1.OPS_YYYYPP > '" & RYP & "'"

            sql &= SQL_in("VEND_CODE", "APTINVH1.VEND_CODE")
            sql &= SQL_in("ACCT_CODE", "APTINVH2.ACCT_CODE")
            sql &= SQL_in("PROCESSOR_CODE", "APTVEND1.PROCESSOR_CODE")
            sql &= SQL_in("VEND_TYPE", "APTVEND1.VEND_TYPE")

            APTAEXP2 = ASCMAIN1.Temp_Table(sql)

            ASCMAIN1.sql = "Select * from " & APTAEXP2
            .Tables.Add(ASCDATA1.GetDataTable("**", "APTAEXP2", 2))
        End With


        Call MyBase.Get_SQL("*", APTAEXP2)
        Call ASCMAIN1.Progress("Building Tiers")

        sql = "Select " & sql_SELECT_cols & vbCr
        sql &= ", APTAEXP2.VOUCHER_NO" & vbCr
        sql &= ", APTAEXP2.VOUCHER_LNO" & vbCr
        sql &= ", APTAEXP2.INV_LINE_AMT" & vbCr
        sql &= " from " & APTAEXP2 & " APTAEXP2 " & sql_TABLE_NAMEs & vbCr
        sql &= ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCr
        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")

        Check_if_Empty("APTAEXP2")

    End Sub

    Public Overrides Sub Print_Report()
        Dim SUBT As String = "For " & RYPLEGEND
        CR_params.Add("DTL", Absx1.optFor("OPTDS").Value)

        Generate_Report(RPT, , SUBT)
    End Sub

    Public Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        If eItemKey = "Proceed" Then

        End If
    End Sub

End Class