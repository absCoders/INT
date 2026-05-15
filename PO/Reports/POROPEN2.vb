Imports ABSolution
Imports Infragistics.Win

Public Class POROPEN2
    Dim POTOPEN2 As String = "" ' Logical Name of Work Table, which will be used as the name of the DataTable

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Range_Events(grpPO_DATE_RANGE)
        Range_Events(grpDateRequired)
        Get_PARM("POTPARM1")

        'Set_cmbYP("RYP", ASCMAIN1.CYP, -36, -1, -1)
    End Sub

    Protected Overrides Sub Build_Workfile()
        MyBase.Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        Dim sql As String

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Dim SORTBY_CODE As String = ""
        If Absx1.optFor("SORTBY").Value = "P" Then
            SORTBY_CODE = "POTORDR2.PO_ORDER_NO"
        Else
            SORTBY_CODE = "POTORDR2.ITEM_CODE"
        End If
        sql = "Select " & SORTBY_CODE & " SORTBY, " & vbCrLf _
        & "POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO, 'X' PO_LINE_TYPE, " & vbCrLf _
        & "POTORDR1.VEND_CODE, POTORDR1.PO_REFERENCE, " & vbCrLf _
        & "POTORDR2.ITEM_CODE, POTORDR2.ITEM_DESC, " & vbCrLf _
        & "POTORDR2.BM_ISSUE_NO, POTORDR2.PO_COST, " & vbCrLf _
        & "POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_REC, " & vbCrLf _
        & "POTORDR2.PO_QTY_INV, POTORDR2.PO_QTY_OPN, " & vbCrLf _
        & "POTORDR2.PO_DATE_REQUIRED , POTORDR2.WHSE_CODE, " & vbCrLf _
        & "POTORDR2.PO_QTY_ORD * POTORDR2.PO_COST PO_AMT_ORD, " & vbCrLf _
        & "POTORDR2.PO_QTY_OPN * POTORDR2.PO_COST PO_AMT_OPN, " & vbCrLf _
        & "POTORDR1.VEND_NAME, POTORDR1.PO_DATE_ORDERED " & vbCrLf _
        & " from POTORDR2,POTORDR1,ICTITEM1,ICTCOLL1" & vbCrLf _
        & " where POTORDR1.PO_ORDER_NO  = POTORDR2.PO_ORDER_NO" & vbCrLf _
        & "   and POTORDR1.PO_STATUS = 'O'" & vbCrLf _
        & "   and ICTITEM1.ITEM_CODE (+) = POTORDR2.ITEM_CODE" & vbCrLf _
        & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf

        If Not Absx1.chkFor("CHKPO_DATE_F").Checked Then
            Dim z As String = Format(Absx1.dteFor("PO_DATE_F").Value, "dd-MMM-yyyy")
            sql = sql & " and POTORDR1.PO_DATE_ORDERED >= '" & z & "'"
            Page0.Add("Purchase Orders dated >= " & z)
        End If
        If Not Absx1.chkFor("CHKPO_DATE_L").Checked Then
            Dim z As String = Format(Absx1.dteFor("PO_DATE_L").Value, "dd-MMM-yyyy")
            sql = sql & " and POTORDR1.PO_DATE_ORDERED <= '" & z & "'"
            Page0.Add("Purchase Orders dated <= " & z)
        End If

        If Not Absx1.chkFor("CHKDATEREQ_F").Checked Then
            Dim z As String = Format(Absx1.dteFor("DATEREQ_F").Value, "dd-MMM-yyyy")
            sql = sql & " and POTORDR2.PO_DATE_REQUIRED >= '" & z & "'"
            Page0.Add("Date Required >= " & z)
        End If
        If Not Absx1.chkFor("CHKDATEREQ_L").Checked Then
            Dim z As String = Format(Absx1.dteFor("DATEREQ_L").Value, "dd-MMM-yyyy")
            sql = sql & " and POTORDR2.PO_DATE_REQUIRED <= '" & z & "'"
            Page0.Add("Date Required <= " & z)
        End If

        ' one of these for each filter in Report Maintenance
        sql &= SQL_in("VEND_CODE", "POTORDR1.VEND_CODE")
        sql &= SQL_in("ITEM_CODE", "POTORDR2.ITEM_CODE")
        sql &= SQL_in("PO_ORDER_NO", "POTORDR1.PO_ORDER_NO")
        sql &= SQL_in("COLLECTION_CODE", "ICTITEM1.COLLECTION_CODE")
        sql &= SQL_in("BRAND_CODE", "ICTCOLL1.BRAND_CODE")

        POTOPEN2 = ASCMAIN1.Temp_Table(sql)

        ASCMAIN1.sql = "Select * from " & POTOPEN2
        dst.Tables.Add(ASCDATA1.GetDataTable("**", "POTOPEN2", 0))

        Check_if_Empty("POTOPEN2")

    End Sub

    Public Overrides Sub Print_Report()
        Dim SUBT As String = ""
        If Not Absx1.chkFor("CHKPO_DATE_F").Checked _
        Or Not Absx1.chkFor("CHKPO_DATE_L").Checked Then
            SUBT = SUBT & "Showing Purchase Orders Dated"
            If Not Absx1.chkFor("CHKPO_DATE_F").Checked Then
                SUBT = SUBT & " from " & Format(Absx1.dteFor("PO_DATE_F").Value, "MM/dd/yyyy")
            End If
            If Not Absx1.chkFor("CHKPO_DATE_L").Checked Then
                SUBT = SUBT & " thru " & Format(Absx1.dteFor("PO_DATE_L").Value, "MM/dd/yyyy")
            End If
        End If

        CR_params.Add("SORTBY", Absx1.optFor("SORTBY").Value)

        Generate_Report(RPT, , SUBT)
    End Sub

    Public Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        If eItemKey = "Proceed" Then

        End If
    End Sub
End Class