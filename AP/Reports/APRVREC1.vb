Public Class APRVREC1

    Dim APTINVH1 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Call Range_Events(grpINV_DATE_RANGE)

        Call Get_PARM("GLTPARM1")
        Call Get_PARM("APTPARM1")

    End Sub

    Protected Overrides Sub Build_Workfile()

        ' Prepare Working Variables

        Call ASCMAIN1.Progress("Building Open AP Work File")

        Dim sql As String = ""

        If Not Absx1.chkFor("CHKINV_DATE_F").Checked Then
            Dim z As String = Format(Absx1.dteFor("INV_DATE_F").Value, "dd-MMM-yyyy")
            sql = sql & " and APTINVH1.INV_DATE >= '" & z & "'"
            Page0.Add("Invoices dated >= " & z)
        End If
        If Not Absx1.chkFor("CHKINV_DATE_L").Checked Then
            Dim z As String = Format(Absx1.dteFor("INV_DATE_L").Value, "dd-MMM-yyyy")
            sql = sql & " and APTINVH1.INV_DATE <= '" & z & "'"
            Page0.Add("Invoices dated <= " & z)
        End If

        If Absx1.optFor("OPTAOH").Value = "O" Or Absx1.optFor("OPTAOH").Value = "H" Then
            sql = sql & " and APTINVH1.INV_STATUS = '" & Absx1.optFor("OPTAOH").Value & "'"
            If Absx1.optFor("OPTAOH").Value = "H" Then
                Page0.Add("Invoices On Hold for Payment")
            Else
                Page0.Add("Invoices OK to Pay")
            End If
        Else
            sql = sql & " and APTINVH1.INV_STATUS in ('O','H')"
        End If

        Dim not_all_types_selected As Boolean = False
        Dim INV_TYPEs As String = ""
        For Each INV_TYPE As String In New String() {"I", "B", "D", "R", "C", "A"}
            If Absx1.chkFor("INV_TYPE_" & INV_TYPE).Checked Then
                INV_TYPEs &= ",'" & INV_TYPE & "'"
            Else
                not_all_types_selected = True
            End If
        Next
        If not_all_types_selected Then
            sql = sql & " and APTINVH1.INV_TYPE in (" & Mid(INV_TYPEs, 2) & ")"
        End If

        sql = sql & SQL_in("VEND_CODE", "APTINVH1.VEND_CODE")

        sql = "Select APTINVH1.* " _
        & " from APTINVH1, APTVEND1" & vbCr _
        & " where APTVEND1.VEND_CODE = APTINVH1.VEND_CODE" & vbCr _
        & sql

        APTINVH1 = ASCMAIN1.Temp_Table(sql)
        ASCMAIN1.sql = "Select * from " & APTINVH1
        dst.Tables.Add(ASCDATA1.GetDataTable("", "APTATBR1", 1))


        ASCMAIN1.sql = "Select APTINVH8.*" _
        & " from APTINVH8," & APTINVH1 & " APTINVH1" _
        & " where APTINVH1.VOUCHER_NO = APTINVH8.VOUCHER_NO"
        Dim APTINVH8 As String = ASCMAIN1.Temp_Table
        ASCMAIN1.sql = "Alter Table " & APTINVH8 & " Add VOUCHER_ADJ_AMT_ALLOWED NUMBER (13,2)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into " & APTINVH8 _
        & " Select APTINVH5.VOUCHER_NO, 0 VOUCHER_ANO" _
        & ", 'Qty or Price Variance' VOUCHER_ADJ_DESC" _
        & ", SUM (DECODE(NVL(APTINVH5.CB,'0'),'1',-1 * NVL(APTINVH5.VAR_AMT,0),0)) VOUCHER_ADJ_AMT" _
        & ", SUM (DECODE(NVL(APTINVH5.CB,'0'),'0',-1 * NVL(APTINVH5.VAR_AMT,0),0)) VOUCHER_ADJ_AMT_ALLOWED" _
        & " from APTINVH5, " & APTINVH1 & " APTINVH1" _
        & " where APTINVH1.VOUCHER_NO = APTINVH5.VOUCHER_NO" _
        & " group by APTINVH5.VOUCHER_NO"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select * from " & APTINVH8
        dst.Tables.Add(ASCDATA1.GetDataTable("**", "APTINVH8", 2))

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Call MyBase.Get_SQL("*")

        sql = "Select " & sql_SELECT_cols & vbCr _
        & ", APTINVH1.VOUCHER_NO" & vbCr _
        & ", APTINVH1.INV_BALANCE" & vbCr

        sql = sql & " from " & APTINVH1 & " APTINVH1 " & sql_TABLE_NAMEs & vbCr
        sql = sql & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCr
        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")


    End Sub

    Public Overrides Sub Print_Report()

        Dim SUBT As String = ""

        If Not Absx1.chkFor("CHKINV_DATE_F").Checked _
        Or Not Absx1.chkFor("CHKINV_DATE_L").Checked Then
            SUBT = SUBT & ", "
            SUBT = SUBT & "Showing A/P Items Dated"
            If Not Absx1.chkFor("CHKINV_DATE_F").Checked Then
                SUBT = SUBT & " from " & Format(Absx1.dteFor("INV_DATE_F").Value, "MM/dd/yyyy")
            End If
            If Not Absx1.chkFor("CHKINV_DATE_L").Checked Then
                SUBT = SUBT & " thru " & Format(Absx1.dteFor("INV_DATE_L").Value, "MM/dd/yyyy")
            End If
        End If

        Generate_Report(RPT, , SUBT)
    End Sub

End Class