Imports System.Math

Public Class APRAEXP1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Get_PARM("GLTPARM1")
        Call Get_PARM("APTPARM1")

        'Set_cmbYP("RYP", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1), -35, -35, 0)
        Set_cmbYP("RYP", ASCMAIN1.CYP, -36, -1, -1)
    End Sub

    Protected Overrides Sub Build_Workfile()

        With dst
            sql = "SELECT APTINVH2.*" _
            & ", APTINVH1.VEND_CODE, APTVEND1.VEND_NAME" _
            & ", APTINVH1.INV_NUM, APTINVH1.INV_DATE" _
            & " from APTINVH1, APTINVH2, APTVEND1 " _
            & " where APTINVH1.VOUCHER_NO = APTINVH2.VOUCHER_NO" _
            & "   and APTINVH1.OPS_YYYYPP_ACCRUE = '" & RYP & "' " _
            & "   and APTVEND1.VEND_CODE = APTINVH1.VEND_CODE" _
            & "   and APTINVH2.INV_LTYP IS NULL"

            sql &= SQL_in("VEND_CODE", "APTINVH1.VEND_CODE")
            sql &= SQL_in("ACCT_CODE", "APTINVH2.ACCT_CODE")

            .Tables.Add(ASCDATA1.GetDataTable(sql, "APTAEXP1", 2))
        End With

        Call Get_WKCodes("APTAEXP1", "ACCT_CODE", "GLTACCT1")

        Check_if_Empty("APTAEXP1")

    End Sub

    Public Overrides Sub Print_Report()
        Dim SUBT As String = "For " & RYPLEGEND
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            'If Absx1.cmbFor("RYP").Text = "" Then
            '    EMsg &= "You Must Select a Period"
            'End If
        End If

    End Sub
End Class