Public Class APRPODL1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -24, 0, 0)
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Work Tables

        ASCMAIN1.sql = "Select ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO" & vbCrLf _
            & ", ICTIREC2.ITEM_CODE, ICTIREC2.QTY_REC, ICTIREC2.QTY_INV, ICTIREC2.PO_COST" & vbCrLf _
            & ", ICTIREC2.PO_ORDER_NO, ICTIREC2.PO_ORDER_LNO, ICTIREC2.OPS_YYYYPP" & vbCrLf _
            & ", ICTIREC1.VEND_CODE, ICTIREC1.RECEIPT_DATE, ICTIREC1.WHSE_CODE, ICTIREC1.SOURCE_DOC_NO" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC" & vbCrLf _
            & " from ICTIREC2,ICTIREC1,ICTITEM1" & vbCrLf _
            & " where ICTIREC2.RECEIPT_NO = ICTIREC1.RECEIPT_NO" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf

        If RYP = ASCMAIN1.CYP Then
            ASCMAIN1.sql &= "" _
                & "   and ICTIREC1.ACCRUAL_STATUS = '0'" & vbCrLf _
                & "   and ICTIREC2.ACCRUAL_STATUS = '0'"
        Else
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "ICTIREC2.QTY_INV", "NVL(ICTIREC2.QTY_REC,0) - NVL(GLTCREC5.CREC_QTY,0) QTY_INV")
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "from ICTIREC2,ICTIREC1,ICTITEM1", "from ICTIREC2,ICTIREC1,ICTITEM1,GLTCREC5")
            ASCMAIN1.sql &= vbCrLf _
                & "   and GLTCREC5.OPS_YYYYPP = '" & RYP & "'" & vbCrLf _
                & "   and GLTCREC5.CREC_TYPE_CODE = 'ICP'" & vbCrLf _
                & "   and GLTCREC5.DETL_CTL_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
                & "   and GLTCREC5.DETL_CTL_LNO = ICTIREC2.RECEIPT_LNO" & vbCrLf
        End If

        Dim APTPODL1 As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & APTPODL1 & " Add Primary Key (RECEIPT_NO,RECEIPT_LNO)")

        If optANP.Value = "N" Then
            ASCMAIN1.sql = "Delete from " & APTPODL1 & " where NVL(QTY_INV,0) <> 0"
            ASCDATA1.ExecuteSQL()
        ElseIf optANP.Value = "P" Then
            ASCMAIN1.sql = "Delete from " & APTPODL1 & " where NVL(QTY_INV,0) = 0"
            ASCDATA1.ExecuteSQL()
        End If
        '  ASCDATA1.ExecuteSQL("Update " & APTPODL1 & " Set QTY_INV = 0")

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Dim sql_Data As String = ""
        Dim sql_Cols As String = ""

        ASCMAIN1.Progress("Accrued Purchases")
        MyBase.Get_SQL("*", APTPODL1)

        sql_Data = "" _
            & ", SUM (APTPODL1.QTY_REC) QTY_REC" & vbCrLf _
            & ", SUM (APTPODL1.QTY_REC * APTPODL1.PO_COST) AMT_REC" & vbCrLf _
            & ", SUM (APTPODL1.QTY_INV) QTY_INV" & vbCrLf _
            & ", SUM ((NVL(APTPODL1.QTY_REC,0) - NVL(APTPODL1.QTY_INV,0)) * APTPODL1.PO_COST) AMT_REC_NOT_INV"

        sql_Cols = "" _
            & ",QTY_REC,AMT_REC,QTY_INV,AMT_REC_NOT_INV"

        If ASCMAIN1.CLIENT = "AHA" Then
            sql_Data &= "" _
                & ", SUM ((NVL(APTPODL1.QTY_REC,0) - NVL(APTPODL1.QTY_INV,0)) * NVL(ICTCOSTA.ITEM_COST_VCOST,0)) VCOST" & vbCrLf _
                & ", SUM ((NVL(APTPODL1.QTY_REC,0) - NVL(APTPODL1.QTY_INV,0)) * NVL(ICTCOSTA.ITEM_COST_LANDG,0)) LANDG" & vbCrLf _
                & ", SUM ((NVL(APTPODL1.QTY_REC,0) - NVL(APTPODL1.QTY_INV,0)) * NVL(ICTCOSTA.ITEM_COST_TOOLG,0)) TOOLG" & vbCrLf _
                & ", SUM ((NVL(APTPODL1.QTY_REC,0) - NVL(APTPODL1.QTY_INV,0)) * NVL(ICTCOSTA.ITEM_COST_OVRHD,0)) OVRHD"

            sql_Cols &= "" _
                & ",VCOST,LANDG,TOOLG,OVRHD"

        End If

        sql_filter = "" _
            & ""

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & ", APTPODL1.RECEIPT_NO, APTPODL1.RECEIPT_LNO" & vbCrLf _
            & sql_Data _
            & " from " & APTPODL1 & " APTPODL1" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & IIf(sql_GROUP_BY_cols = "", "''", sql_GROUP_BY_cols) & vbCrLf _
            & ", APTPODL1.RECEIPT_NO, APTPODL1.RECEIPT_LNO"

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "Select * from " & APTPODL1 & " APTPODL1"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTPODL1", 2))

    End Sub

    Public Overrides Sub Print_Report()
        SUBT = ""  

        If optANP.Value = "N" Then
            SUBT = "Showing Open Accruals which have Never been Invoiced Only, "
        ElseIf optANP.Value = "P" Then
            SUBT = "Showing Open Accruals which beeen Partially Invoiced Only. "
        End If

        If RYP = ASCMAIN1.CYP Then
            SUBT &= "Current Period " & RYPLEGEND
        Else
            SUBT &= "End of Month " & RYPLEGEND
        End If
        Generate_Report(RPT, , SUBT)
        
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        'If eItemKey = "Proceed" Then
        '    If Absx1.cmbFor("RYP").Value & "" = "" Then
        '        EMsg &= vbCr & "You must Specify a Reporting Period"
        '    End If
        'End If
    End Sub

    Overrides Sub Build_Report_File_Post_Process()
        'dst.Tables("ASTSRPT1").Columns("BOOKED_TYTM").Expression = "ISNULL(CARRIED_FWD,0)+ISNULL(BOOKED_PRV,0)+ISNULL(BOOKED_CUR,0)"
        'dst.Tables("ASTSRPT1").Columns("PROJECTED_TYTM").Expression = "ISNULL(SHIPPED_TYTM,0)+ISNULL(OTS_M01,0)"
    End Sub
      
End Class