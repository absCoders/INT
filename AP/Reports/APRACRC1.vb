Imports System.Drawing

Public Class APRACRC1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -24, 0, 0)
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        ' Prepare Work Tables

        ASCMAIN1.sql = "Select APTACRC1.CTL_NO,
                APTACRC1.VEND_CODE_ACC,
                APTACRC1.ACCRUAL_CODE,
                APTACRC1.COST_ACT,
                APTACRC1.COST_ACC,
                APTACRC1.CHARGEBACK_IND,
                APTACRC1.VOUCHER_NO,
                APTACRC1.CTL_STATUS,
                APTACRC1.OPS_YYYYPP,
                APTACRC1.VOUCHER_NO_ORIG,
                APTACRC1.CTL_DATE,
                APTACRC1.CTL_NOTE,
                APTACRC1.INV_PRINT_IND,
                APTACRC1.INV_PRINT_DATE,
                APTACRC1.INV_PRINT_USER,
                APTACRC1.COST_ORIG,
                APTACRC1.RECEIPT_NO,
                APTACRC1.RECEIPT_LNO,
                APTACRC1.PO_ORDER_NO,
                APTACRC1.PO_ORDER_LNO,
                APTACRC1.ITEM_CODE,
                APTACRC1.CTL_TYPE,
                APTACRC1.COST_CATGY_CODE,
                APTACRC1.SOURCE_DOC_NO
                from APTACRC1"

        If RYP = ASCMAIN1.CYP Then
            ASCMAIN1.sql &= " WHERE (" & vbCrLf
            ASCMAIN1.sql &= " (CTL_STATUS = '0' or (CTL_STATUS = '1' AND NVL(PPD_IND,'0') = '1' AND NVL(PPD_MATCHED,'0') = '0'))"
        ASCMAIN1.sql &= ")"
        Else
            'ASCMAIN1.sql = Replace(ASCMAIN1.sql, "APTACRC1.COST_ACT", "NVL(APTACRC1.COST_ACT,0) COST_ACT")
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "from APTACRC1", "from APTACRC1, GLTCREC5")

            ASCMAIN1.sql &= vbCrLf _
            & " where GLTCREC5.OPS_YYYYPP = '" & RYP & "'" & vbCrLf _
            & "   and APTACRC1.CTL_NO = GLTCREC5.DETL_CVX_NO"

            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "APTACRC1.COST_ACT", "CASE WHEN NVL(APTACRC1.PPD_IND,'0') = '1' THEN -1 * GLTCREC5.CREC_AMT ELSE 0 END COST_ACT")
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "APTACRC1.COST_ACC", "CASE WHEN NVL(APTACRC1.PPD_IND,'0') = '1' THEN 0 ELSE GLTCREC5.CREC_AMT END COST_ACC")

            'ASCMAIN1.sql &= vbCrLf _
            '& " WHERE GLTCREC5.OPS_YYYYPP = '" & RYP & "'" & vbCrLf _
            '& " AND APTACRC1.RECEIPT_NO = GLTCREC5.DETL_CTL_NO AND APTACRC1.RECEIPT_LNO = GLTCREC5.DETL_CTL_LNO" & vbCrLf _
            '& " AND GLTCREC5.CREC_TYPE_CODE = DECODE(APTACRC1.ACCRUAL_CODE,'TRF','ICT','FRT','ICF','')"

            ' need to support both joins until we repair the FRT records in GLTCREC5, replacing ITEM_CODE with CTL_NO in DETL_CVX_NO
            ' the reason CTL_NO is important for TRF is that we don't have a RECEIPT_NO for PPD
            'ASCMAIN1.sql &= vbCrLf _
            '& " where GLTCREC5.OPS_YYYYPP = '" & RYP & "'" & vbCrLf _
            '& "   and ((APTACRC1.ACCRUAL_CODE = 'FRT' and APTACRC1.RECEIPT_NO = GLTCREC5.DETL_CTL_NO AND APTACRC1.RECEIPT_LNO = GLTCREC5.DETL_CTL_LNO)" & vbCrLf _
            '& "    or  (APTACRC1.ACCRUAL_CODE = 'TRF' and APTACRC1.CTL_NO = GLTCREC5.DETL_CVX_NO))" & vbCrLf _
            '& "   and GLTCREC5.CREC_TYPE_CODE = DECODE(APTACRC1.ACCRUAL_CODE,'TRF','ICT','FRT','ICF','')"

        End If

        Dim APTACRC1 As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & APTACRC1 & " Add Primary Key (CTL_NO)")
        ASCDATA1.ExecuteSQL("Alter Table " & APTACRC1 & " Add QTY_REC Number (6,0)")

        ASCDATA1.ExecuteSQL("Update " & APTACRC1 & " X Set QTY_REC = (Select QTY_REC from ICTIREC2 where RECEIPT_NO = X.RECEIPT_NO and RECEIPT_LNO = X.RECEIPT_LNO)")

        If RYP = ASCMAIN1.CYP Then
        Else
            ' ASCDATA1.ExecuteSQL("Update " & APTACRC1 & " X Set COST_ACT = 0 where ACCRUAL_CODE = 'FRT'")

        End If

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Dim sql_Data As String = ""
        Dim sql_Cols As String = ""

        ASCMAIN1.Progress("Accrued Other")
        MyBase.Get_SQL("*", APTACRC1)

        sql_Data = "" _
            & ", SUM (APTACRC1.COST_ACC) COST_ACC" & vbCrLf _
            & ", SUM (APTACRC1.COST_ACT) COST_ACT"

        sql_Cols = "" _
            & ",COST_ACC,COST_ACT"

        sql_filter = "" _
            & ""

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & ", APTACRC1.RECEIPT_NO, APTACRC1.RECEIPT_LNO, APTACRC1.CTL_NO" & vbCrLf _
            & sql_Data _
            & " from " & APTACRC1 & " APTACRC1" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & IIf(sql_GROUP_BY_cols = "", "''", sql_GROUP_BY_cols) & vbCrLf _
            & ", APTACRC1.RECEIPT_NO, APTACRC1.RECEIPT_LNO, APTACRC1.CTL_NO"

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select * from " & APTACRC1 & " APTACRC1"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTACRC1", 1))
        dst.Tables("APTACRC1").Columns.Add("NET_ACCRUAL", GetType(System.Decimal), "ISNULL(COST_ACC,0) - ISNULL(COST_ACT,0)")

        ASCMAIN1.sql = "Select * from APTVEND1 where VEND_CODE in (Select Distinct VEND_CODE_ACC from " & APTACRC1 & ")"
        Create_TDA(dst.Tables.Add, "APTVEND1", "**", 0, False, "", 1)
        Fill_Records("APTVEND1")

        ASCMAIN1.sql = "Select * from APTACRM1 where ACCRUAL_CODE in (Select Distinct ACCRUAL_CODE from " & APTACRC1 & ")"
        Create_TDA(dst.Tables.Add, "APTACRM1", "**", 0, False, "", 1)
        Fill_Records("APTACRM1")
    End Sub

    Public Overrides Sub Print_Report()

        If RYP = ASCMAIN1.CYP Then
            SUBT = "Current Period - " & RYPLEGEND
        Else
            SUBT = "Previous Period - " & RYPLEGEND
        End If

        Generate_Report(RPT, , SUBT)
        Prepare_Data_Extracts()
    End Sub
    Sub Prepare_Data_Extracts()

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        grdASTEXPT1.DataSource = dst.Tables("APTACRC1")
        grdASTEXPT1.Text = "Open Accrued Other"
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")
        Set_DX_Column(grdASTEXPT1, "VEND_CODE_ACC", "Vendor Code", 100,, "Count", Color.SteelBlue)
        Set_DX_Column(grdASTEXPT1, "ACCRUAL_CODE", "Accrual Code", 90,,, Color.DarkTurquoise)
        Set_DX_Column(grdASTEXPT1, "CTL_TYPE", "Type", 40,,, Color.DarkTurquoise)
        Set_DX_Column(grdASTEXPT1, "PO_ORDER_NO", "PO No", 80,,, Color.DarkKhaki)
        Set_DX_Column(grdASTEXPT1, "ITEM_CODE", "Item Code", 110,,, Color.SteelBlue)
        Set_DX_Column(grdASTEXPT1, "COST_ACC", "Cost Accrued", 100, "#,##0.00", "Sum", Color.MediumPurple)
        Set_DX_Column(grdASTEXPT1, "COST_ACT", "Cost Actual", 100, "#,##0.00", "Sum", Color.MediumPurple)
        Set_DX_Column(grdASTEXPT1, "NET_ACCRUAL", "Net Accrual", 100, "#,##0.00", "Sum", Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "VOUCHER_NO", "PPD Voucher", 90,,, Color.DarkKhaki)
        Set_DX_Column(grdASTEXPT1, "OPS_YYYYPP", "Period", 80,,, Color.DarkSeaGreen)
        Set_DX_Column(grdASTEXPT1, "RECEIPT_NO", "Rec No", 80,,, Color.DarkKhaki)
        Set_DX_Column(grdASTEXPT1, "CTL_NO", "Ctl No", 100,,, Color.DarkKhaki)
        Set_DX_Column(grdASTEXPT1, "SOURCE_DOC_NO", "Bill of Lading", 130,,, Color.DarkKhaki)
        Set_DX_Column(grdASTEXPT1, "CTL_DATE", "Control Date", 100, "MM/dd/yyyy", , Color.LightSlateGray)

    End Sub


    Public Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        Select Case eItemKey
            Case "Proceed"

                If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length < 1 Then
                    EMsg &= vbCr & "You Must Select a Sort when running Other AP Accruals Report"
                End If

        End Select
    End Sub

    Overrides Sub Build_Report_File_Post_Process()
        'dst.Tables("ASTSRPT1").Columns("BOOKED_TYTM").Expression = "ISNULL(CARRIED_FWD,0)+ISNULL(BOOKED_PRV,0)+ISNULL(BOOKED_CUR,0)"
        'dst.Tables("ASTSRPT1").Columns("PROJECTED_TYTM").Expression = "ISNULL(SHIPPED_TYTM,0)+ISNULL(OTS_M01,0)"
    End Sub

End Class