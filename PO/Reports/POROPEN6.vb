Public Class POROPEN6

    Dim POTOPEN6 As String = "" ' Logical Name of Work Table, which will be used as the name of the DataTable
    Dim H(12) As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("POTPARM1")

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)
    End Sub

    Protected Overrides Sub Build_Workfile()
        MyBase.Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        Dim sql As String = ""

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        Dim DTX As String = "ICTIREC1.OPS_YYYYPP"

        ' Extracts from Data Sources

        Dim P(12) As String
        Dim YP As String = RYP0
        For i As Int32 = 0 To 12
            P(i) = ASCMAIN1.Period_Calc(YP, i - 1)
            If P(i) <= RYP1 Then
                H(i) = ASCMAIN1.Get_Legend(P(i), False, True)
            Else
                H(i) = ""
            End If
        Next

        Dim VX As String = ""

        ' Create the Work File - as a flattened out result set containing all PB columns and data fields that you would like to put in the detail section of the report

        sql = "Select POTORDR1.VEND_CODE, ICTIREC2.ITEM_CODE, POTORDR1.MARKET_CODE" & vbCrLf

        VX = "NVL(ICTIREC2.QTY_REC,0) * NVL(ICTIREC2.PO_COST,0)"
        Dim sqlsumPO As String = ""
        For i As Int32 = 1 To 12
            sql &= ", SUM (CASE WHEN " & DTX & " > '" & P(i - 1) & "' and " & DTX & " <= '" & P(i) & "' THEN " & VX & " ELSE 0 END) PO_" & Format(i, "00") & vbCrLf
            sqlsumPO &= "+NVL(PO_" & Format(i, "00") & ",0)"
        Next
        sqlsumPO = "PO_TOT = " & Mid(sqlsumPO, 2)
        sql &= ", 0 PO_TOT" & vbCrLf

        VX = "NVL(ICTIREC2.QTY_REC,0)"
        Dim sqlsumPP As String = ""
        For i As Int32 = 1 To 12
            sql &= ", SUM (CASE WHEN " & DTX & " > '" & P(i - 1) & "' and " & DTX & " <= '" & P(i) & "' THEN " & VX & " ELSE 0 END) PP_" & Format(i, "00") & vbCrLf
            sqlsumPP &= "+NVL(PP_" & Format(i, "00") & ",0)"
        Next
        sqlsumPP = "PP_TOT = " & Mid(sqlsumPP, 2)
        sql &= ", 0 PP_TOT" & vbCrLf

        sql &= " from ICTIREC1,ICTIREC2,POTORDR1,ICTITEM1" & vbCrLf _
        & " where POTORDR1.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO" & vbCrLf _
        & "   and ICTIREC2.RECEIPT_NO = ICTIREC1.RECEIPT_NO" & vbCrLf _
        & "   and ICTITEM1.ITEM_CODE = ICTIREC2.ITEM_CODE"


        ' one of these for each filter in Report Maintenance
        sql &= SQL_in("VEND_CODE", "POTORDR1.VEND_CODE")
        'sql &= SQL_in("BUYER_CODE", "BUYER_CODE")
        'sql &= SQL_in("PO_ORDER_TYPE", "PO_ORDER_TYPE")
        'sql &= SQL_in("PO_ORDER_NO", "PO_ORDER_NO")
        sql &= SQL_in("ITEM_CODE", "ICTIREC2.ITEM_CODE")
        sql &= SQL_in("MARKET_CODE", "POTORDR1.MARKET_CODE")

        sql = sql & " and ICTIREC1.OPS_YYYYPP >= '" & RYP0 & "'" & vbCrLf
        sql = sql & " and ICTIREC1.OPS_YYYYPP <= '" & RYP1 & "'" & vbCrLf

        sql &= " group by POTORDR1.VEND_CODE, ICTIREC2.ITEM_CODE, POTORDR1.MARKET_CODE" & vbCrLf

        POTOPEN6 = ASCMAIN1.Temp_Table(sql)
        ASCDATA1.ExecuteSQL("Update " & POTOPEN6 & " Set " & sqlsumPO)
        ASCDATA1.ExecuteSQL("Update " & POTOPEN6 & " Set " & sqlsumPP)

        MyBase.Get_SQL("*", POTOPEN6)
        ASCMAIN1.Progress("Building Tiers")

        sql = "Select " & sql_SELECT_cols & ASTSRPT1_sum_columns & vbCr
        sql &= " from " & POTOPEN6 & " POTOPEN6 " & sql_TABLE_NAMEs & vbCr
        sql &= ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCr
        sql &= " group by " & sql_GROUP_BY_cols
        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")

        Dim rowASTDSQLS As DataRow
        For i As Int32 = 1 To 12
            rowASTDSQLS = tblASTDSQLS.Select("COLUMN_NAME = 'PO_" & Format(i, "00") & "'")(0)
            rowASTDSQLS.Item("COLUMN_CAPTION") = "$" & H(i)
            rowASTDSQLS = tblASTDSQLS.Select("COLUMN_NAME = 'PP_" & Format(i, "00") & "'")(0)
            rowASTDSQLS.Item("COLUMN_CAPTION") = "#" & H(i)
        Next
        rowASTDSQLS = tblASTDSQLS.Select("COLUMN_NAME = 'PO_TOT'")(0)
        rowASTDSQLS.Item("COLUMN_CAPTION") = "$Total"
        rowASTDSQLS = tblASTDSQLS.Select("COLUMN_NAME = 'PP_TOT'")(0)
        rowASTDSQLS.Item("COLUMN_CAPTION") = "#Total"

    End Sub

    Public Overrides Sub Print_Report()
        Dim SUBT As String = "PO Receipts"
        SUBT = SUBT & " from " & RYPLEGEND0 & " thru " & RYPLEGEND1

        CR_params.Add("PO", IIf(Absx1.chkFor("CHKPO").Checked, "1", "0"))
        CR_params.Add("PP", IIf(Absx1.chkFor("CHKPP").Checked, "1", "0"))
        CR_params.Add("PO_TEXT", "$Value")
        CR_params.Add("PP_TEXT", "#Units")

        For i As Int32 = 1 To 12
            CR_params.Add("H" & Format(i, "00"), H(i))
        Next
        Generate_Report(RPT, , SUBT)
    End Sub

    Public Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        If eItemKey = "Proceed" Then
            If Not Absx1.chkFor("CHKPO").Checked And Not Absx1.chkFor("CHKPP").Checked Then
                EMsg &= vbCr & "You must select at least 1: $Value Received and/or Units Received"
            End If
        End If
    End Sub

End Class
