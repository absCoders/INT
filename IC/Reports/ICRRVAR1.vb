Public Class ICRRVAR1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP", ASCMAIN1.CYP, -36, 0, 0)
        optCOST.Value = "F" '"S"
    End Sub

    Protected Overrides Sub Build_Workfile()

        ' Prepare Working Variables
        MyBase.RWU = "N"

        ASCMAIN1.Progress("Run-Time Options", "")

        sql = "Select * from ICTCOLL1"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "ICTCOLL1"))

        sql = "Select * from ICTBRAN1"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "ICTBRAN1"))

        Dim costType As String = optCOST.Value

        Dim col_sfx As String = ""
        Dim cTableName As String = "ICTCOSTA"
        If RYP = ASCMAIN1.CYP Then
            Select Case costType
                Case "S"
                    cTableName = "ICTCOSTC"
                    col_sfx = "CUR"
                Case "F"
                    cTableName = "ICTCOSTF"
                    col_sfx = "FUT"
            End Select
        End If

        ' **************************************************************
        ' ******************* Not sure this is correct *****************
        ' **************************************************************
        If costType <> "S" Then
            Dim TABLES As New Dictionary(Of String, String)

            ASCMAIN1.Progress("ReCalculate Costs", "")
            TABLES = ICCMAIN1.ReCalculate_Costs(Me, costType, "R")

            If TABLES.ContainsKey("ICTCOSTX") Then
                cTableName = TABLES("ICTCOSTX")
            End If

        End If

        sql = "Select ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO" & vbCrLf
        sql &= ", ICTIREC1.PO_ORDER_NO, ICTIREC2.PO_ORDER_LNO, ICTIREC2.BM_ISSUE_NO PO_DTL_BM_ISSUE_NO" & vbCrLf
        sql &= ", ICTIREC1.VEND_CODE, ICTIREC1.RECEIPT_DATE, ICTIREC2.QTY_REC" & vbCrLf
        sql &= ", APTACRC1_F.XXX_CLASS_CODE FRT_CLASS_CODE, APTACRC1_F.XXX_CLASS_PCT FRT_CLASS_PCT" & vbCrLf
        sql &= ", APTACRC1_T.XXX_CLASS_CODE TRF_CLASS_CODE, APTACRC1_T.XXX_CLASS_PCT TRF_CLASS_PCT" & vbCrLf
        sql &= ", ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTCOLL1.COLLECTION_CODE, ICTBRAN1.BRAND_CODE" & vbCrLf
        sql &= ", ICTIREC2.PO_COST, ROUND(DECODE(NVL(ICTIREC2.QTY_REC, 0),0,0,NVL(ICTIREC2.EXT_COST_MATLS, 0) / ICTIREC2.QTY_REC), 4) PO_COST_MATLS, APTVEND1.CURR_CODE" & vbCrLf
        sql &= ", ICTIREC2.TRAN_PV, ICTIREC2.TRAN_MV, 0.01 TRAN_FV, 0.01 TRAN_TV, 0.01 TOTAL_VAR, 0.01 TOTAL_VAR_PCT" & vbCrLf
        sql &= ", NVL(APTACRC1_F.COST_ACC,0) FRT_COST, NVL(APTACRC1_T.COST_ACC,0) TRF_COST" & vbCrLf
        sql &= ", ICTCOSTX.ITEM_COST_VCOST, ICTCOSTX.ITEM_COST_MATLS, ICTCOSTX.ITEM_COST_CURR_CODE, ICTCOSTX.BM_ISSUE_NO ITEM_ISSUE_NO" & vbCrLf
        sql &= $", ICTCOSTX.ITEM_COST_LANDG, ICTCOSTX.ITEM_COST_FRT_CLASS, ICTFRTC1.FRT_CLASS_PCT_{col_sfx} FRT_CLASS_PCT_STD" & vbCrLf
        sql &= $", ICTCOSTX.ITEM_COST_TOOLG, ICTCOSTX.ITEM_COST_TRF_CLASS, ICTTRFC1.TRF_CLASS_PCT_{col_sfx} TRF_CLASS_PCT_STD" & vbCrLf
        sql &= " from ICTIREC1, ICTIREC2, ICTITEM1, ICTCOLL1, ICTBRAN1, APTVEND1, " & cTableName & " ICTCOSTX" & vbCrLf
        sql &= ", APTACRC1 APTACRC1_F, APTACRC1 APTACRC1_T, ICTFRTC1, ICTTRFC1" & vbCrLf
        sql &= "  where ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf
        sql &= "  and ICTIREC2.ITEM_CODE = ICTCOSTX.ITEM_CODE" & vbCrLf
        sql &= "  and ICTIREC2.ITEM_CODE = ICTITEM1.ITEM_CODE" & vbCrLf
        sql &= "  and ICTITEM1.COLLECTION_CODE = ICTCOLL1.COLLECTION_CODE" & vbCrLf
        sql &= "  and ICTCOLL1.BRAND_CODE = ICTBRAN1.BRAND_CODE " & vbCrLf
        sql &= "  and ICTIREC2.OPS_YYYYPP = '" & RYP & "'" & vbCrLf
        sql &= "  and ICTIREC1.REVERSED_BY_RECEIPT_NO IS NULL and ICTIREC1.REVERSES_RECEIPT_NO IS NULL" & vbCrLf
        sql &= "  and ICTIREC1.VEND_CODE = APTVEND1.VEND_CODE" & vbCrLf
        sql &= "  and ICTCOSTX.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf
        sql &= "  and APTACRC1_F.RECEIPT_NO (+) = ICTIREC2.RECEIPT_NO" & vbCrLf
        sql &= "  and APTACRC1_F.RECEIPT_LNO (+) = ICTIREC2.RECEIPT_LNO" & vbCrLf
        sql &= "  and APTACRC1_F.ACCRUAL_CODE (+) = 'FRT'" & vbCrLf
        sql &= "  and APTACRC1_T.RECEIPT_NO (+) = ICTIREC2.RECEIPT_NO" & vbCrLf
        sql &= "  and APTACRC1_T.RECEIPT_LNO (+) = ICTIREC2.RECEIPT_LNO" & vbCrLf
        sql &= "  and APTACRC1_T.ACCRUAL_CODE (+) = 'TRF'" & vbCrLf
        sql &= "  and ICTFRTC1.FRT_CLASS_CODE (+) = ICTCOSTX.ITEM_COST_FRT_CLASS" & vbCrLf
        sql &= "  and ICTTRFC1.TRF_CLASS_CODE (+) = ICTCOSTX.ITEM_COST_TRF_CLASS" & vbCrLf
        If RYP <> ASCMAIN1.CYP Then
            sql &= "  and ICTCOSTX.OPS_YYYYPP = '" & RYP & "'" & vbCrLf
        End If

        sql &= SQL_in("COLLECTION_CODE", "ICTITEM1.COLLECTION_CODE")
        sql &= SQL_in("BRAND_CODE", "ICTCOLL1.BRAND_CODE")
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "ICTRVAR1", 0))

        dst.Tables("ICTRVAR1").Columns("TOTAL_VAR").ReadOnly = False
        dst.Tables("ICTRVAR1").Columns("TOTAL_VAR_PCT").ReadOnly = False
        dst.Tables("ICTRVAR1").Columns("PO_COST_MATLS").ReadOnly = False
        dst.Tables("ICTRVAR1").Columns("TRAN_PV").ReadOnly = False
        dst.Tables("ICTRVAR1").Columns("TRAN_MV").ReadOnly = False
        dst.Tables("ICTRVAR1").Columns("TRAN_FV").ReadOnly = False
        dst.Tables("ICTRVAR1").Columns("TRAN_TV").ReadOnly = False

        ' ICTIREC4
        Dim sqlICTIREC4 = "Select sum (QTY_CON * (Decode(ICTCOSTX.ITEM_COST_VCOST ,Null, 0, ICTCOSTX.ITEM_COST_VCOST) + Decode(ICTCOSTX.ITEM_COST_MATLS, Null ,0, ICTCOSTX.ITEM_COST_MATLS)))" & vbCrLf
        sqlICTIREC4 &= " from ICTIREC4, " & cTableName & " ICTCOSTX" & vbCrLf
        sqlICTIREC4 &= " where ICTIREC4.RECEIPT_NO = :PARM1" & vbCrLf
        sqlICTIREC4 &= " and ICTIREC4.RECEIPT_LNO = :PARM2" & vbCrLf
        sqlICTIREC4 &= " and ICTCOSTX.ITEM_CODE = ICTIREC4.ITEM_CODE" & vbCrLf
        If RYP <> ASCMAIN1.CYP Then
            sqlICTIREC4 &= "  and ICTCOSTX.OPS_YYYYPP = '" & RYP & "'" & vbCrLf
        End If

        ASCMAIN1.Progress("Process Items", "")
        Dim pcost As Decimal
        Dim mcost As Decimal
        Dim fcost As Decimal
        Dim tcost As Decimal
        Dim pcost_std As Decimal
        Dim mcost_std As Decimal
        Dim fcost_std As Decimal
        Dim tcost_std As Decimal
        Dim mcost_tot As Decimal
        Dim fcost_tot As Decimal
        Dim tcost_tot As Decimal
        Dim AMT As Double
        Dim QTY_REC As Int32
        For Each rowICTRVAR1 As DataRow In dst.Tables("ICTRVAR1").Select("", "RECEIPT_NO, RECEIPT_LNO")

            ASCMAIN1.Progress("-", rowICTRVAR1.Item("RECEIPT_NO") & ", " & rowICTRVAR1.Item("RECEIPT_LNO"))

            rowICTRVAR1.Item("TOTAL_VAR") = Val(rowICTRVAR1.Item("TRAN_PV") & String.Empty) + Val(rowICTRVAR1.Item("TRAN_MV") & String.Empty) + Val(rowICTRVAR1.Item("TRAN_FV") & String.Empty) + Val(rowICTRVAR1.Item("TRAN_TV") & String.Empty)
            QTY_REC = Val(rowICTRVAR1.Item("QTY_REC") & String.Empty)

            fcost_tot = Val(rowICTRVAR1.Item("FRT_COST") & String.Empty)
            tcost_tot = Val(rowICTRVAR1.Item("TRF_COST") & String.Empty)

            pcost = Val(rowICTRVAR1.Item("PO_COST") & String.Empty)
            mcost = Val(rowICTRVAR1.Item("PO_COST_MATLS") & String.Empty)
            fcost = fcost_tot / QTY_REC
            tcost = tcost_tot / QTY_REC

            pcost_std = Val(rowICTRVAR1.Item("ITEM_COST_VCOST") & String.Empty)
            mcost_std = Val(rowICTRVAR1.Item("ITEM_COST_MATLS") & String.Empty)
            fcost_std = Val(rowICTRVAR1.Item("ITEM_COST_LANDG") & String.Empty)
            tcost_std = Val(rowICTRVAR1.Item("ITEM_COST_TOOLG") & String.Empty)

            If RYP = ASCMAIN1.CYP Then
                mcost_tot = 0
                If rowICTRVAR1.Item("PO_DTL_BM_ISSUE_NO") & String.Empty <> String.Empty Then
                    mcost_tot = Val(ASCDATA1.GetDataValue(sqlICTIREC4, "VN", New Object() {rowICTRVAR1.Item("RECEIPT_NO") & String.Empty, _
                                                                                           rowICTRVAR1.Item("RECEIPT_LNO") & String.Empty}) & String.Empty)
                End If
                If QTY_REC <> 0 Then
                    mcost = mcost_tot / QTY_REC
                Else
                    mcost = 0
                End If
                rowICTRVAR1.Item("PO_COST_MATLS") = mcost
                rowICTRVAR1.Item("TRAN_PV") = QTY_REC * (pcost - pcost_std)
                rowICTRVAR1.Item("TRAN_MV") = QTY_REC * (mcost - mcost_std)
                rowICTRVAR1.Item("TRAN_FV") = fcost_tot - QTY_REC * (fcost_std)
                rowICTRVAR1.Item("TRAN_TV") = tcost_tot - QTY_REC * (tcost_std)
            End If
            AMT = QTY_REC * (pcost_std + mcost_std)
            rowICTRVAR1.Item("TOTAL_VAR") = Val(rowICTRVAR1.Item("TRAN_PV") & String.Empty) + Val(rowICTRVAR1.Item("TRAN_MV") & String.Empty) + Val(rowICTRVAR1.Item("TRAN_FV") & String.Empty) + Val(rowICTRVAR1.Item("TRAN_TV") & String.Empty)
            rowICTRVAR1.Item("TOTAL_VAR") = System.Math.Round(Val(rowICTRVAR1.Item("TOTAL_VAR") & ""), 2)
            If AMT = 0 Then
                rowICTRVAR1.Item("TOTAL_VAR_PCT") = 0
            Else
                rowICTRVAR1.Item("TOTAL_VAR_PCT") = 100 * rowICTRVAR1.Item("TOTAL_VAR") / AMT
            End If
        Next

        Dim rowICTIVAR1s() As DataRow = dst.Tables("ICTRVAR1").Select("TOTAL_VAR <= " & numVariance.Value)
        If rowICTIVAR1s.Length > 0 Then
            For i As Int32 = rowICTIVAR1s.Length - 1 To 0 Step -1
                Dim rowICTIVAR1 As DataRow = rowICTIVAR1s(i)
                rowICTIVAR1.Delete()
            Next
        End If

        'For Each rowICTRVAR1 As DataRow In dst.Tables("ICTRVAR1").Select("TOTAL_VAR <= " & numVariance.Value)
        '    rowICTRVAR1.Delete()
        'Next

        dst.Tables("ICTRVAR1").AcceptChanges()

    End Sub

    Public Overrides Sub Print_Report()
        SUBT = "Using " & optCOST.Text
        Generate_Report(RPT, , SUBT)
        Prepare_Data_Extracts()
    End Sub

    Sub Prepare_Data_Extracts()

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand
        grdASTEXPT1.DataSource = dst.Tables("ICTRVAR1")

        grdASTEXPT1.Text = "PO Receipts Variance Details"
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")


        Set_DX_Column(grdASTEXPT1, "ITEM_CODE", "Item Code", 120, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "ITEM_DESC", "Item Description", 200, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "COLLECTION_CODE", "Collection", 80, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "BRAND_CODE", "Brand", 60, , , Color.LightBlue)

        Set_DX_Column(grdASTEXPT1, "RECEIPT_NO", "Rec No", 70, , , Color.CornflowerBlue)
        Set_DX_Column(grdASTEXPT1, "RECEIPT_LNO", "Ln", 40, , , Color.CornflowerBlue)
        Set_DX_Column(grdASTEXPT1, "PO_ORDER_NO", "PO No", 70, , , Color.CornflowerBlue)
        Set_DX_Column(grdASTEXPT1, "PO_ORDER_LNO", "Ln", 40, , , Color.CornflowerBlue)

        Set_DX_Column(grdASTEXPT1, "PO_DTL_BM_ISSUE_NO", "BM#", 50, , , Color.CornflowerBlue)

        Set_DX_Column(grdASTEXPT1, "VEND_CODE", "Vendor", 100, , , Color.CornflowerBlue)
        Set_DX_Column(grdASTEXPT1, "RECEIPT_DATE", "Rec Date", 110, "MM/dd/yyyy", , Color.CornflowerBlue)
        Set_DX_Column(grdASTEXPT1, "QTY_REC", "Qty Rec", 80, "#,##0", , Color.CornflowerBlue)
        Set_DX_Column(grdASTEXPT1, "FRT_CLASS_CODE", "Frt", 50, , , Color.LightGray)
        Set_DX_Column(grdASTEXPT1, "FRT_CLASS_PCT", "Frt%", 50, "##.00", , Color.LightGray)
        Set_DX_Column(grdASTEXPT1, "FRT_COST", "$Frt", 80, "#,##0.00", "Sum", Color.LightGray)
        Set_DX_Column(grdASTEXPT1, "TRF_CLASS_CODE", "Trf", 50, , , Color.PaleVioletRed)
        Set_DX_Column(grdASTEXPT1, "TRF_CLASS_PCT", "Trf%", 50, "##.00", , Color.PaleVioletRed)
        Set_DX_Column(grdASTEXPT1, "TRF_COST", "$Trf", 80, "#,##0.00", "Sum", Color.PaleVioletRed)



        Set_DX_Column(grdASTEXPT1, "PO_COST", "PO Cost", 80, "#,##0.00", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "CURR_CODE", "Curr", 40, , , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "PO_COST_MATLS", "Matls", 80, "#,##0.00", , Color.LightGreen)

        Set_DX_Column(grdASTEXPT1, "TRAN_PV", "PPV", 80, "#,##0.00", "Sum", Color.PaleTurquoise)
        Set_DX_Column(grdASTEXPT1, "TRAN_MV", "MPV", 80, "#,##0.00", "Sum", Color.PaleTurquoise)
        Set_DX_Column(grdASTEXPT1, "TRAN_FV", "FPV", 80, "#,##0.00", "Sum", Color.PaleTurquoise)
        Set_DX_Column(grdASTEXPT1, "TRAN_TV", "TPV", 80, "#,##0.00", "Sum", Color.PaleTurquoise)
        Set_DX_Column(grdASTEXPT1, "TOTAL_VAR", "Total Var", 100, "#,##0.00", "Sum", Color.PaleTurquoise)
        Set_DX_Column(grdASTEXPT1, "TOTAL_VAR_PCT", "Var%", 50, "##.00", , Color.PaleTurquoise)

        Set_DX_Column(grdASTEXPT1, "ITEM_COST_VCOST", "Std VCost", 80, "#,##0.00", , Color.LightPink)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_CURR_CODE", "Curr", 40, , , Color.LightPink)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_MATLS", "Std Matls", 80, "#,##0.00", , Color.LightPink)
        Set_DX_Column(grdASTEXPT1, "ITEM_ISSUE_NO", "BM#", 50, , , Color.LightPink)

        Set_DX_Column(grdASTEXPT1, "ITEM_COST_LANDG", "Std Frt", 80, "#,##0.00", , Color.LightGray)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_FRT_CLASS", "FrtCls", 50, , , Color.LightGray)
        Set_DX_Column(grdASTEXPT1, "FRT_CLASS_PCT_STD", "Frt%", 50, "##.00", , Color.LightGray)

        Set_DX_Column(grdASTEXPT1, "ITEM_COST_TOOLG", "Std Trf", 80, "#,##0.00", , Color.PaleVioletRed)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_TRF_CLASS", "Std Trf", 50, , , Color.PaleVioletRed)
        Set_DX_Column(grdASTEXPT1, "TRF_CLASS_PCT_STD", "Std Frt%", 50, "##.00", , Color.PaleVioletRed)


        Sort_grdColumns(grdASTEXPT1, "BRAND_CODE, COLLECTION_CODE, ITEM_CODE")

        'grdASTEXPT1.DisplayLayout.Bands(0).Columns("ITEM_CODE").Header.Fixed = True
        'Sort_grdColumns(grdASTEXPT1, "ITEM_CODE")

    End Sub
    Overrides Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Public Overrides Sub Verify_Special(eItemKey As String)
        MyBase.Verify_Special(eItemKey)

        Select Case eItemKey
            Case "Proceed"
                RYPLEGEND = cmbOPS.Text
                RYP = Mid$(RYPLEGEND, 1, 4) & Mid$(RYPLEGEND, 6, 2)

                If RYP <> ASCMAIN1.CYP AndAlso optCOST.Value <> "S" Then
                    EMsg &= vbCr & "Future costs are available only for the current period."
                End If
        End Select
    End Sub

End Class