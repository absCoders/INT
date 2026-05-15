Public Class ICFCOST1
    Dim ICTSTATX As String
    Dim sqlICTSTATX As String
    Dim sqlICTIVAL1 As String

    Dim rowICTITEM1 As DataRow
    Dim rowICTCOSTM As DataRow

    Dim RYP As String = ASCMAIN1.CYP
    Dim ITEM_CODE As String

    Dim ICTCOSTP As String
    Dim sqlICTCOSTP As String

    ' AP RECORDS - RELATED TO RECEIPTS
    ' STAT1/STAT2 - NEC?

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")
        Get_PARM("GLTPARM1")

        If MENU_ITEM_OBJECT = "ICFCOSTI" Then InquiryMode = True

        With dst

            ASCMAIN1.sql = "Select ICTCOSTA.* from ICTCOSTA"
            Create_TDA(.Tables.Add, "ICTCOSTM", "**", 0, False, "", 0)
            .Tables("ICTCOSTM").Columns.Add("SEQ")
            .Tables("ICTCOSTM").Columns.Add("SEQ_DESC")
            .Tables("ICTCOSTM").Columns.Add("ITEM_COST_TOTAL_D", GetType(System.Decimal), "ISNULL(ITEM_COST_VCOST,0)+ISNULL(ITEM_COST_LANDG,0)+ISNULL(ITEM_COST_TOOLG,0)+ISNULL(ITEM_COST_OVRHD,0)")
            .Tables("ICTCOSTM").Columns.Add("ITEM_COST_TOTAL_M", GetType(System.Decimal), "ISNULL(ITEM_COST_MATLS,0)+ISNULL(ITEM_COST_LANDGI,0)+ISNULL(ITEM_COST_TOOLGI,0)+ISNULL(ITEM_COST_OVRHDI,0)")

            ASCMAIN1.sql = "Select ICTCOSTH.* from ICTCOSTH where ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTCOSTH", "**", 0, True, "V")

            ASCMAIN1.sql = "Select ICTCOSTC.* from ICTCOSTC where ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTCOSTC", "**", 0, True, "V")

            ASCMAIN1.sql = "Select ICTCOSTF.* from ICTCOSTF where ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTCOSTF", "**", 0, True, "V")

            ' NOTE THERE IS ALREADY AN EPHEMERAL ICTCOSTP IN THIS SCREEN. CALLING THIS ONE ICTCOSTP1
            ASCMAIN1.sql = "Select ICTCOSTP.* from ICTCOSTP where ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTCOSTP1", "**", 0, False, "V")

            ASCMAIN1.sql = "Select ICTIVAR1.* from ICTIVAR1 where ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTIVAR1", "**", 0, True, "V")

            ASCMAIN1.sql = "Select ICTIREC2.*" & vbCrLf _
                & ", DECODE(ICTIREC2.ACCRUAL_STATUS,'1',0,NVL(ICTIREC2.QTY_REC,0) - NVL(ICTIREC2.QTY_INV,0)) QTY_REC_NOT_INV" & vbCrLf _
                & ", DECODE(ICTIREC2.ACCRUAL_STATUS,'1',0,(NVL(ICTIREC2.QTY_REC,0) - NVL(ICTIREC2.QTY_INV,0)) * NVL(ICTIREC2.PO_COST,0)) AMT_REC_NOT_INV" & vbCrLf _
                & ", ICTIREC1.RECEIPT_DATE, ICTIREC1.VEND_CODE" & vbCrLf _
                & ", ICTIREC1.WHSE_CODE" & vbCrLf _
                & ", ICTIREC1.SOURCE_DOC_NO" & vbCrLf _
                & " from ICTIREC2,ICTIREC1" & vbCrLf _
                & " where ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
                & " and ICTIREC2.ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTIRECX", "**", 0, False, "V")

            ASCMAIN1.sql = "Select APTINVH5.*, APTINVH1.INV_NUM, APTINVH1.INV_DATE" & vbCrLf _
                & " from ICTIREC2,APTINVH5,APTINVH1" & vbCrLf _
                & " where ICTIREC2.ITEM_CODE = :PARM1" & vbCrLf _
                & " and (APTINVH5.INV_QTY <> 0 or APTINVH5.CLOSE_LINE = '1')" & vbCrLf _
                & " and APTINVH1.VOUCHER_NO = APTINVH5.VOUCHER_NO" & vbCrLf _
                & " and APTINVH1.INV_STATUS <> 'D'" & vbCrLf _
                & " and APTINVH5.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
                & " and APTINVH5.RECEIPT_LNO = ICTIREC2.RECEIPT_LNO"
            Create_TDA(.Tables.Add, "APTINVH5", "**", 0, False, "V")

            Create_Relation("ICTIRECX", "APTINVH5", "RECEIPT_NO,RECEIPT_LNO")

            ASCMAIN1.sql = "Select APTACRC1.*, ICTIREC2.QTY_REC" & vbCrLf _
                & " from APTACRC1,ICTIREC2 " & vbCrLf _
                & " where APTACRC1.ITEM_CODE = :PARM1" & vbCrLf _
                & " and ICTIREC2.RECEIPT_NO = APTACRC1.RECEIPT_NO" & vbCrLf _
                & " and ICTIREC2.RECEIPT_LNO = APTACRC1.RECEIPT_LNO"
            Create_TDA(.Tables.Add, "APTACRC1", "**", 0, False, "V")
            .Tables("APTACRC1").Columns.Add("VARIANCE", GetType(System.Decimal), "IIF(ISNULL(VOUCHER_NO,'?')<>'?',ISNULL(COST_ACT,0)-ISNULL(COST_ACC,0),0)")
            .Tables("APTACRC1").Columns.Add("COST_VAR_ITEM_PER_UNIT", GetType(System.Decimal), "IIF(ISNULL(QTY_REC,0)=0,0,(ISNULL(COST_VAR_ITEM,0) + ISNULL(TPV_ADJ,0))/ISNULL(QTY_REC,0))")
            .Tables("APTACRC1").Columns.Add("COST_ACT_CALC", GetType(System.Decimal), "IIF(ISNULL(VOUCHER_NO,'?')<>'?',ISNULL(COST_ACT,0), IIF(ISNULL(CTL_STATUS,'0')='0', NULL, ISNULL(COST_VAR_ITEM,0) + ISNULL(COST_ACC,0) + ISNULL(TPV_ADJ,0)))")

            Create_Relation("ICTIRECX", "APTACRC1", "RECEIPT_NO,RECEIPT_LNO")

            '           & " and (APTINVH7.INV_QTY <> 0 or APTINVH7.CLOSE_LINE = '1')" & vbCrLf _
            ASCMAIN1.sql = "Select APTINVH7.*, APTINVH1.INV_NUM, APTINVH1.INV_DATE" & vbCrLf _
                & " from APTACRC1,APTINVH7,APTINVH1,ICTIREC2" & vbCrLf _
                & " where APTACRC1.ITEM_CODE = :PARM1" & vbCrLf _
                & " and APTINVH1.VOUCHER_NO = APTINVH7.VOUCHER_NO" & vbCrLf _
                & " and APTINVH1.INV_STATUS <> 'D'" & vbCrLf _
                & " and APTINVH7.CTL_NO = APTACRC1.CTL_NO" & vbCrLf _
                & " and ICTIREC2.RECEIPT_NO = APTACRC1.RECEIPT_NO" & vbCrLf _
                & " and ICTIREC2.RECEIPT_LNO = APTACRC1.RECEIPT_LNO"
            Create_TDA(.Tables.Add, "APTINVH7", "**", 0, False, "V")
            .Tables("APTINVH7").Columns.Add("VARIANCE", GetType(System.Decimal), "ISNULL(TOTAL_INV,0)-ISNULL(TOTAL_ACC,0)")
            '.Tables("APTINVH7").Columns.Add("COST_VAR_ITEM_PER_UNIT", GetType(System.Decimal), "ISNULL(COST_VAR_ITEM_PER_UNIT,0)/ISNULL(TOTAL_ACC,0)")

            Create_Relation("APTACRC1", "APTINVH7", "CTL_NO")

            ASCMAIN1.sql = "Select POTORDR2.*" & vbCrLf _
                & ", POTORDR1.VEND_CODE, POTORDR1.PO_DATE_ORDERED, POTORDR1.PO_DATE_CANCEL" & vbCrLf _
                & ", POTORDR1.MARKET_CODE" & vbCrLf _
                & ", POTORDR1.PO_ORDR_NOTES_INTERNAL, POTORDR1.PO_ORDR_NOTES_EXTERNAL" & vbCrLf _
                & ", POTORDR1.PO_ORDER_TYPE, POTORDR1.PO_TYPE, POTORDR1.PO_DATE_SHIPPED" & vbCrLf _
                & ", POTORDR1.PO_DATE_ETA, POTORDR1.PO_SHIP_VESSEL, POTORDR1.CONTAINER_NO" & vbCrLf _
                & " from POTORDR2,POTORDR1" & vbCrLf _
                & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR2.ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "POTORDRX", "**", 0, False, "V", 2)

            Dim sqlICTCOSTP_Cols As String = ""
            Dim sqlICTCOSTP_non_zero As String = ""
            ASCMAIN1.sql = "Select ITEM_CODE" & vbCrLf
            For i As Int16 = 1 To 13
                Dim C As String = "ITEM_COST_P" & Format(i, "00")
                sqlICTCOSTP_Cols &= ", " & C
                sqlICTCOSTP_non_zero &= " or NVL(" & C & ",0) <> 0"
                ASCMAIN1.sql &= ", SUM (DECODE(OPS_YYYYPP,'XXXX" & Format(i, "00") & "',ITEM_COST_TOTAL,0)) " & C & vbCrLf
            Next
            ASCMAIN1.sql &= " from ICTCOSTA" & vbCrLf _
                & " where OPS_YYYYPP >= 'XXXX01' and OPS_YYYYPP <= 'XXXX13'" & vbCrLf _
                & "   and NVL(ITEM_COST_TOTAL,0) <> 0" & vbCrLf _
                & " group by ITEM_CODE" & vbCrLf

            Dim sqlICTCOSTP_OH As String = ASCMAIN1.sql

            ASCMAIN1.sql = "Select ITEM_CODE" & vbCrLf
            For i As Int16 = 1 To 13
                Dim C As String = "BOM_P" & Format(i, "00")
                sqlICTCOSTP_Cols &= ", " & C
                sqlICTCOSTP_non_zero &= " or NVL(" & C & ",0) <> 0"
                ASCMAIN1.sql &= ", SUM (DECODE(OPS_YYYYPP,'XXXX" & Format(i, "00") & "',WHSE_QTY_BEG,0)) " & C & vbCrLf
            Next
            ASCMAIN1.sql &= " from ICTSTAT1" & vbCrLf _
                & " where OPS_YYYYPP >= 'XXXX01' and OPS_YYYYPP <= 'XXXX13'" & vbCrLf _
                & "   and NVL(WHSE_QTY_BEG,0) <> 0" & vbCrLf _
                & " group by ITEM_CODE" & vbCrLf

            Dim sqlICTCOSTP_COST As String = ASCMAIN1.sql

            '   sqlICTCOSTP &= vbCrLf & " union " & vbCrLf & ASCMAIN1.sql

            ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_CATGY_CODE" & vbCrLf _
                & ", ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
                & ", ICTITEM1.PROD_CODE, ICTITEM1.COST_CATGY_CODE" & vbCrLf _
                & sqlICTCOSTP_Cols _
                & " from ICTITEM1, ICTCOLL1" & vbCrLf _
                & ", (" & sqlICTCOSTP_COST & ") X_COST" & vbCrLf _
                & ", (" & sqlICTCOSTP_OH & ") X_OH" & vbCrLf _
                & " where X_COST.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
                & "   and X_OH.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & "   and (" & Mid(sqlICTCOSTP_non_zero, 4) & ")"
            sqlICTCOSTP = ASCMAIN1.sql
            ICTCOSTP = ASCMAIN1.Temp_Table(sqlICTCOSTP)

            ASCMAIN1.sql = "Select * from " & ICTCOSTP
            Create_TDA(.Tables.Add, "ICTCOSTP", "**", 0, False, "", 1)
            .Tables("ICTCOSTP").Columns.Add("COST_VAR", GetType(System.Decimal))
            .Tables("ICTCOSTP").Columns.Add("COST_VAR_PCT", GetType(System.Decimal))
            For I As Integer = 2 To 13
                .Tables("ICTCOSTP").Columns.Add("VAR_P" & Format(I, "00"), GetType(System.Decimal))
                .Tables("ICTCOSTP").Columns("VAR_P" & Format(I, "00")).Expression = "ISNULL(BOM_P" & Format(I, "00") & ",0)*(ISNULL(ITEM_COST_P" & Format(I, "00") & ",0)- ISNULL(ITEM_COST_P" & Format(I - 1, "00") & ",0))"
            Next

            sqlICTSTATX = "Select ICTSTAT1.ITEM_CODE, ICTSTAT1.OPS_YYYYPP" & vbCrLf _
                & ", ICTSTAT1.WHSE_QTY_BEG" & vbCrLf _
                & ", ICTSTAT1.WHSE_QTY_SHP" & vbCrLf _
                & ", ICTSTAT1.WHSE_QTY_RTN" & vbCrLf _
                & ", ICTSTAT1.WHSE_QTY_REC" & vbCrLf _
                & ", ICTSTAT1.WHSE_QTY_ADJ" & vbCrLf _
                & ", ICTSTAT1.WHSE_QTY_XFR" & vbCrLf _
                & ", ICTSTAT1.WHSE_QTY_CON" & vbCrLf _
                & ", ICTSTAT1.WHSE_QTY_RTV" & vbCrLf _
                & ", ICTSTAT1.WHSE_QTY_PHY" & vbCrLf _
                & ", ICTSTAT2.WHSE_QTY_ON_HAND" & vbCrLf _
                & ", ICTSTAT2.WHSE_QTY_ONPO" & vbCrLf _
                & ", ICTSTAT2.WHSE_QTY_PLAN" & vbCrLf _
                & ", ICTSTAT2.WHSE_QTY_OPEN" & vbCrLf _
                & ", ICTSTAT2.WHSE_QTY_PICK" & vbCrLf _
                & ", ICTSTAT2.WHSE_QTY_COMM" & vbCrLf _
                & ", ICTSTAT2.WHSE_QTY_HOLD" & vbCrLf _
                & " from ICTSTAT1,ICTSTAT2"
            ASCMAIN1.sql = sqlICTSTATX & " where ROWNUM < 1"
            ICTSTATX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ICTSTATX & " Add Primary Key (ITEM_CODE,OPS_YYYYPP)")

            sqlICTIVAL1 = "Select ICTSTATX.*, P.ITEM_COST_TOTAL ITEM_COST_TOTAL_P, C.ITEM_COST_TOTAL" & vbCrLf _
                & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & ", ICTITEM1.COST_CATGY_CODE, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_SNU_CODE, ICTITEM1.ITEM_BASIC_PROMO, ICTITEM1.ITEM_COST_STATUS" & vbCrLf _
                & " from ICTITEM1," & ICTSTATX & " ICTSTATX,ICTCOSTA P,ICTCOSTA C" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = ICTSTATX.ITEM_CODE" & vbCrLf _
                & " and C.ITEM_CODE (+) = ICTSTATX.ITEM_CODE" & vbCrLf _
                & " and C.OPS_YYYYPP (+) = :PARM1" & vbCrLf _
                & " and P.ITEM_CODE (+) = ICTSTATX.ITEM_CODE" & vbCrLf _
                & " and P.OPS_YYYYPP (+) = :PARM2" & vbCrLf
            ASCMAIN1.sql = sqlICTIVAL1
            Create_TDA(.Tables.Add, "ICTIVAL1", "**", 0, False, "VV", 2)
            .Tables("ICTIVAL1").Columns.Add("COST_REVAL", GetType(System.Decimal), "ISNULL(ITEM_COST_TOTAL,0) - ISNULL(ITEM_COST_TOTAL_P,0)")
            .Tables("ICTIVAL1").Columns.Add("COST_REVAL_PCT", GetType(System.Decimal), "IIF(ISNULL(ITEM_COST_TOTAL_P,0)=0,0,COST_REVAL / ISNULL(ITEM_COST_TOTAL_P,0))")
            .Tables("ICTIVAL1").Columns.Add("COST_REVAL_AMT", GetType(System.Decimal), "COST_REVAL * ISNULL(WHSE_QTY_BEG,0)")
            .Tables("ICTIVAL1").Columns.Add("WHSE_AMT_BEG", GetType(System.Decimal), "ISNULL(ITEM_COST_TOTAL_P,0) * ISNULL(WHSE_QTY_BEG,0)")
            .Tables("ICTIVAL1").Columns.Add("WHSE_AMT_ON_HAND", GetType(System.Decimal), "ISNULL(ITEM_COST_TOTAL,0) * ISNULL(WHSE_QTY_ON_HAND,0)")

            ASCMAIN1.sql = "Select ICTITEM1.*" & vbCrLf _
                & " from ICTITEM1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = :PARM1"
            For Each TABLE_NAME As String In New String() {"ICTITEM1", "ICTITEM1_RECENT"}
                Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, False, "V", 1)
                .Tables(TABLE_NAME).Columns.Add("IMAGE", GetType(System.Byte()))
            Next

            For Each TABLE_NAME As String In New String() _
                {"ICTCOLL1", "ICTCATG1", "ICTFRTC1", "ICTTRFC1", "ICTWHSE1", "ICTCOST1", "ICTPROD1"}
                Create_TDA(.Tables.Add, TABLE_NAME, "*", 0, False)
                Fill_Records(TABLE_NAME)
            Next

            With .Tables.Add("ICTIVARX")
                .Columns.Add("XTD")
                .Columns.Add("PPV", GetType(System.Decimal))
                .Columns.Add("MUV", GetType(System.Decimal))
                .Columns.Add("CPV", GetType(System.Decimal))
                .Columns.Add("REV", GetType(System.Decimal))
                .Columns.Add("FPV", GetType(System.Decimal))
                .Columns.Add("TPV", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() {.Columns("XTD")}
            End With

            With .Tables.Add("ICTITEMX")
                .Columns.Add("COLUMN_NAME")
                .Columns.Add("COLUMN_CAPTION")
                .Columns.Add("COLUMN_CODE")
                .Columns.Add("COLUMN_DESC")
                .PrimaryKey = New DataColumn() {.Columns("COLUMN_NAME")}
            End With
        End With

        ASCMAIN1.sql = "Select OPS_YYYYPP, LEGEND from GLTPARM2" _
         & " where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -36) & "'" _
         & "   and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "'"

        Dim dvw As DataView = ASCDATA1.GetDataTable.DefaultView
        dvw.Sort = "OPS_YYYYPP DESC"
        cbeYP.DataSource = dvw
        cbeYP.DataSource = ASCDATA1.GetDataTable.DefaultView
        cbeYP.ValueMember = "OPS_YYYYPP"
        cbeYP.DisplayMember = "LEGEND"
        cbeYP.Value = ASCMAIN1.CYP

        Dim dvw1 As DataView = ASCDATA1.GetDataTable.DefaultView
        dvw1.Sort = "OPS_YYYYPP DESC"
        cbeYP1.DataSource = dvw1
        cbeYP1.ValueMember = "OPS_YYYYPP"
        cbeYP1.DisplayMember = "LEGEND"
        cbeYP1.Value = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -3)

        Dim dvw2 As DataView = ASCDATA1.GetDataTable.DefaultView
        dvw2.Sort = "OPS_YYYYPP DESC"
        cbeYP2.DataSource = dvw2
        cbeYP2.ValueMember = "OPS_YYYYPP"
        cbeYP2.DisplayMember = "LEGEND"
        cbeYP2.Value = ASCMAIN1.CYP

        grdICTCOSTM.DataSource = dst.Tables("ICTCOSTM")
        grdICTCOSTP.DataSource = dst.Tables("ICTCOSTP")
        grdICTCOSTH.DataSource = dst.Tables("ICTCOSTH")
        grdICTIVAL1.DataSource = dst.Tables("ICTIVAL1")
        grdICTIVAR1.DataSource = dst.Tables("ICTIVAR1")
        grdICTIVARX.DataSource = dst.Tables("ICTIVARX")
        grdICTIRECX.DataSource = dst.Tables("ICTIRECX")
        grdPOTORDRX.DataSource = dst.Tables("POTORDRX")
        grdICTITEM1_Recent.DataSource = dst.Tables("ICTITEM1_RECENT")
        grdICTITEMX.DataSource = dst.Tables("ICTITEMX")
        grdICTCOSTP1.DataSource = dst.Tables("ICTCOSTP1")

        Show_Filter(grdICTITEM1_Recent, True)
        grdICTITEM1_Recent.DisplayLayout.GroupByBox.Hidden = False
        Create_Summary(grdICTITEM1_Recent, "ITEM_CODE", "Count")

        Create_Summary(grdICTIRECX, "RECEIPT_NO", "Count")
        Create_Summary(grdICTIRECX, New String() {"TRAN_PV", "TRAN_MV", "TRAN_FV", "TRAN_TV", "EXT_COST_MATLS", "QTY_REC", "QTY_INV", "QTY_REC_NOT_INV", "AMT_REC_NOT_INV"})

        With grdICTCOSTM.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ITEM_CODE", "ITEM_COST_TOTAL", "SEQ_DESC", "OPS_YYYYPP"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            .Columns("ITEM_COST_TOTAL").CellAppearance.BackColor = Color.LightBlue
            '.Columns("ITEM_COST_VCURR").CellAppearance.BackColor = Color.Yellow
            '.Columns("ITEM_COST_FRT_CLASS").CellAppearance.BackColor = Color.Yellow
            '.Columns("ITEM_COST_MAKE_BUY").CellAppearance.BackColor = Color.Yellow
        End With

        With grdICTCOSTH.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ITEM_COST_TOTAL", "OPS_YYYYPP"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdICTIVAL1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"OPS_YYYYPP", "ITEM_CODE", "ITEM_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() _
                     {"WHSE_QTY_BEG", "WHSE_QTY_SHP", "WHSE_QTY_RTN", "WHSE_QTY_REC", "WHSE_QTY_ADJ", _
                      "WHSE_QTY_XFR", "WHSE_QTY_CON", "WHSE_QTY_RTV", "WHSE_QTY_PHY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                    gcol.Width = 80
                    gcol.Format = "#,##0"
                    'Create_Summary(grdICTIVAL1, gcol.Key)
                ElseIf New String() _
                    {"WHSE_QTY_ON_HAND", "WHSE_QTY_ONPO", "WHSE_QTY_PLAN", "WHSE_QTY_OPEN", _
                    "WHSE_QTY_PICK", "WHSE_QTY_COMM", "WHSE_QTY_HOLD"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.Goldenrod
                    gcol.Width = 80
                    gcol.Format = "#,##0"
                    'Create_Summary(grdICTIVAL1, gcol.Key)
                ElseIf New String() _
                   {"WHSE_AMT_ON_HAND", "WHSE_AMT_BEG"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                    gcol.Width = 90
                    gcol.Format = "#,##0"
                    'Create_Summary(grdICTIVAL1, gcol.Key)
                ElseIf New String() _
                    {"ITEM_COST_TOTAL", "ITEM_COST_TOTAL_P"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.Pink
                    gcol.Width = 80
                    gcol.Format = "#,##0.0000"
                ElseIf New String() _
                    {"COST_REVAL", "COST_REVAL_AMT", "COST_REVAL_PCT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.Orange
                    gcol.Width = 80
                Else
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                End If

                For Each COLUMN_NAME As String In New String() {"WHSE_QTY_BEG", "ITEM_COST_TOTAL_P", "WHSE_AMT_BEG"}
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Color.Turquoise
                Next
                For Each COLUMN_NAME As String In New String() {"WHSE_QTY_ON_HAND", "ITEM_COST_TOTAL", "WHSE_AMT_ON_HAND"}
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Color.Gold
                Next
            Next
            .Columns("COST_REVAL").Format = "#,##0.0000"
            .Columns("COST_REVAL_AMT").Format = "#,##0"
            .Columns("COST_REVAL_PCT").Format = "#,##0.0"

        End With

        'Create_Summary(grdICTIVAL1, "ITEM_CODE", "Count")
        'Create_Summary(grdICTIVAL1, "COST_REVAL_AMT")

        Create_Summary(grdICTCOSTP, "ITEM_CODE", "Count")

        grdICTCOSTP.DisplayLayout.Bands(0).ColHeaderLines = 2

        With grdICTCOSTP.DisplayLayout.Bands(0)
            For I As Int16 = 1 To 13
                Dim COLUMN_NAME = "ITEM_COST_P" & Format(I, "00")
                .Columns(COLUMN_NAME).Format = "###,##0.0000"
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Color.LightBlue
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                COLUMN_NAME = "BOM_P" & Format(I, "00")
                .Columns(COLUMN_NAME).Format = "#,##0"
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Color.LightGreen
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                Create_Summary(grdICTCOSTP, COLUMN_NAME)
                If I > 1 Then

                    COLUMN_NAME = "VAR_P" & Format(I, "00")
                    .Columns(COLUMN_NAME).Format = "#,##0.00"
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor = Color.White
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Color.Orange
                    .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    Create_Summary(grdICTCOSTP, COLUMN_NAME)
                End If
            Next
        End With


        For Each COLUMN_NAME As String In New String() {"ITEM_CODE", "COLLECTION_CODE", "HC_CODE", "BRAND_CODE", "ITEM_CATGY_CODE", "ITEM_DESC", "PROD_CODE", "COST_CATGY_CODE"}
            With grdICTCOSTP.DisplayLayout.Bands(0)
                If COLUMN_NAME <> "ITEM_CODE" Then .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
                .Columns(COLUMN_NAME).Header.Fixed = True


            End With
        Next
         
        ASCMAIN1.Add_Value_List(grdICTCOSTP, "ITEM_CATGY_CODE")
        ASCMAIN1.Add_Value_List(grdICTCOSTH, "ITEM_CATGY_CODE")

        ASCMAIN1.Add_Value_List(grdICTCOSTM, "ITEM_COST_MAKE_BUY", Nothing, New String() {":", "M:Make", "B:Buy"})

        Set_Read_Only(grpItemMasterData, True)
        Bind_Controls(grpICTCOSTM, "ICTCOSTM")

        Dim blnTOOLG As Boolean = False
        If ASCMAIN1.CLIENT = "AHA" Then blnTOOLG = True
        If ASCMAIN1.CLIENT = "INT" Then blnTOOLG = True

        lblTOOLG.Visible = blnTOOLG
        lblOVRHD.Visible = False
        Absx1.numFor("ITEM_COST_TOOLG").Visible = blnTOOLG
        Absx1.numFor("ITEM_COST_TOOLGI").Visible = False ' blnTOOLG
        Absx1.numFor("ITEM_COST_OVRHD").Visible = False
        Absx1.numFor("ITEM_COST_OVRHDI").Visible = False

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTCOSTM, grdICTCOSTH}
            grd.DisplayLayout.Bands(0).Columns("ITEM_COST_TOOLG").Hidden = Not blnTOOLG
            grd.DisplayLayout.Bands(0).Columns("ITEM_COST_OVRHD").Hidden = True
            grd.DisplayLayout.Bands(0).Columns("ITEM_COST_TOOLGI").Hidden = True ' Not blnTOOLG
            grd.DisplayLayout.Bands(0).Columns("ITEM_COST_OVRHDI").Hidden = True
        Next

        '   Absx1.numFor("ITEM_COST_VCURR").MaskInput = "nnn.nnnnnn"

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View", "Edit"

                If Absx1.txtFor("ITEM_CODE").Text = "" Then
                    EMsg &= vbCr & "You must specify an Item Code to View"
                Else
                    rowICTITEM1 = LookUp("ICTITEM1", Absx1.txtFor("ITEM_CODE").Text)
                    If rowICTITEM1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Item " & Absx1.txtFor("ITEM_CODE").Text & " on File"
                    End If
                End If

                If eItemKey = "Edit" Then
                    rowICTITEM1 = LookUp("ICTITEM1", Absx1.txtFor("ITEM_CODE").Text)
                    If rowICTITEM1 IsNot Nothing AndAlso rowICTITEM1.Item("ITEM_COST_STATUS") & "" <> "P" Then
                        EMsg &= vbCr & "May only edit items pending cost initialization. Item " & Absx1.txtFor("ITEM_CODE").Text & " is not pending."
                    End If

                    ASCMAIN1.Logical_Lock("ICTITEM1", Absx1.txtFor("ITEM_CODE").Text)
                End If

            Case "Update"

                If rowICTITEM1.Item("ITEM_COST_STATUS") & "" <> "P" Then
                    EMsg &= vbCr & "May only edit items pending cost initialization. Item " & Absx1.txtFor("ITEM_CODE").Text & " is not pending."
                End If
                Calculate_Cost()

                EMsg &= Validate_Field("ITEM_COST_FRT_CLASS", Nothing, "Freight Class Code", "ICTFRTC1", Nothing)
                EMsg &= Validate_Field("ITEM_CATGY_CODE", rowICTITEM1, "Item Category Code", "ICTCATG1", Nothing)
                EMsg &= Validate_Field("COST_CATGY_CODE", rowICTITEM1, "Cost Category Code", "ICTCOST1", Nothing)
                EMsg &= Validate_Field("PROD_CODE", rowICTITEM1, "Product Code", "ICTPROD1", Nothing)
                EMsg &= Validate_Field("ITEM_SNU_CODE", rowICTITEM1, "Item SNU Code", "", New String() {"S", "N", "U"})
                EMsg &= Validate_Field("ITEM_BASIC_PROMO", rowICTITEM1, "Item Basic/Promo Code", "", New String() {"B", "P"})
                EMsg &= Validate_Field("ITEM_COST_MAKE_BUY", rowICTCOSTM, "Make/Buy", "", New String() {"M", "B"})

                'If Absx1.optFor("ITEM_COST_MAKE_BUY").Value = "M" Then
                '    EMsg &= vbCrLf & "Feature not supported (yet) for Make Items"
                'End If
                 
                If Absx1.optFor("ITEM_COST_MAKE_BUY").Value = "B" Then
                    rowICTCOSTM.Item("BM_ISSUE_NO") = ""
                Else
                    Dim BM_ISSUE_NO As String = rowICTCOSTM.Item("BM_ISSUE_NO") & ""
                    If BM_ISSUE_NO = "" Then
                        EMsg &= vbCrLf & "Item Must have a BM Issue to be a Make Item"
                    Else
                        Dim rowBMTMAIN2 As DataRow = LookUp("BMTMAIN2", New String() {ITEM_CODE, BM_ISSUE_NO})
                        If rowBMTMAIN2 Is Nothing Then
                            EMsg &= vbCrLf & "BM Issue Selected Not on File"
                        Else
                            If BM_ISSUE_NO = "00" Or rowBMTMAIN2.Item("BM_ISSUE_USE_FOR_STD") & "" <> "1" Then
                                EMsg &= vbCrLf & "Invalid BM Issue to use for Std Cost"
                            End If
                        End If
                    End If
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Dim ITEM_CODE_save As String = ITEM_CODE
                Update_Record()
                Mode_Settings(False)
                Absx1.txtFor("ITEM_CODE").Text = ITEM_CODE_save
                Click_Command("View")

            Case "Done"

                'If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" And Format(Now, "yyyyMMdd") = "20251016" Then
                '    Stop
                '    Fix_Variances("202508")
                '    Fix_Variances("202509")
                '    Stop
                'End If

                Mode_Settings(False)

            Case "Cancel"
                Dim ITEM_CODE_save As String = ITEM_CODE
                Mode_Settings(False)
                Absx1.txtFor("ITEM_CODE").Text = ITEM_CODE_save
                Click_Command("View")
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    If EntryMode = "V" And Not InquiryMode Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode

                    .Items("Edit").Visible = (Not InquiryMode)
                    .Items("Done").Visible = Not (EntryMode = "N" Or EntryMode = "E")
                    .Items("Update").Visible = (Not InquiryMode And EntryMode <> "V")
                    .Items("Cancel").Visible = (Not InquiryMode And EntryMode <> "V")

                    If InquiryMode Then
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                        .Items("Update").Settings.Enabled = not_iScreenMode
                        .Items("Cancel").Settings.Enabled = not_iScreenMode
                    End If
                End With

                .Groups("Valuation").Visible = Not ScreenMode
                .Groups("Cost History").Visible = Not ScreenMode
                .Groups("Item Master").Visible = ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ASCMAIN1.CLIENT = "AHA" Then
            lblTOOLG.Text = "Tooling"
        End If

        splMain.Visible = ScreenMode
        grdICTCOSTP.Visible = False
        chkNonZeroVarianceOnly.Visible = False
        grdICTIVAL1.Visible = False
        tabSummary.Visible = Not ScreenMode

        If Not ScreenMode Then
            grdICTCOSTM.Parent = splItems.Panel2
            grdICTCOSTM.Text = "Items with Future Costs Pending"
            grdICTCOSTM.DisplayLayout.Bands(0).Columns("ITEM_CODE").Hidden = False
            Show_Filter(grdICTCOSTM, True)

            With grdICTIVAL1.DisplayLayout.Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    If New String() _
                         {"WHSE_QTY_BEG", "WHSE_QTY_SHP", "WHSE_QTY_RTN", "WHSE_QTY_REC", "WHSE_QTY_ADJ", _
                          "WHSE_QTY_XFR", "WHSE_QTY_CON", "WHSE_QTY_RTV", "WHSE_QTY_PHY"}.Contains(gcol.Key) Then
                        Create_Summary(grdICTIVAL1, gcol.Key)
                    ElseIf New String() _
                        {"WHSE_QTY_ON_HAND", "WHSE_QTY_ONPO", "WHSE_QTY_PLAN", "WHSE_QTY_OPEN", _
                        "WHSE_QTY_PICK", "WHSE_QTY_COMM", "WHSE_QTY_HOLD"}.Contains(gcol.Key) Then
                        Create_Summary(grdICTIVAL1, gcol.Key)
                    ElseIf New String() _
                       {"WHSE_AMT_ON_HAND", "WHSE_AMT_BEG"}.Contains(gcol.Key) Then
                        Create_Summary(grdICTIVAL1, gcol.Key)
                    End If

                Next
            End With

            Create_Summary(grdICTIVAL1, "ITEM_CODE", "Count")
            Create_Summary(grdICTIVAL1, "COST_REVAL_AMT")

        Else
            grdICTCOSTM.Parent = splMain.Panel1
            grdICTCOSTM.Text = "Cost Details"
            grdICTCOSTM.DisplayLayout.Bands(0).Columns("ITEM_CODE").Hidden = True
            Show_Filter(grdICTCOSTM, False)

            grdICTIVAL1.DisplayLayout.Bands(0).Summaries.Clear()
        End If

        cmdCalculateCost.Visible = (EntryMode = "E")

        lblInitialCost.Visible = (EntryMode = "E" AndAlso rowICTITEM1.Item("ITEM_COST_STATUS") & "" = "P")

        grpItemMasterData.Visible = ScreenMode
        Setup_tabDetails()

        With grdICTIVAL1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ITEM_CODE", "ITEM_DESC", "COLLECTION_CODE", "ITEM_CATGY_CODE", "COST_CATGY_CODE", "PROD_CODE", "ITEM_SNU_CODE", "ITEM_BASIC_PROMO"}
                .Columns(COLUMN_NAME).Hidden = ScreenMode
            Next
            .Columns("OPS_YYYYPP").Hidden = Not ScreenMode
        End With



        If ScreenMode Then
            grdICTIVAL1.Parent = tabDetails.Tabs("Valuation History").TabPage
            grdICTIVAL1.Visible = True ''

            Set_Read_Only(grpICTCOSTM, (EntryMode <> "E"))
            ' grdICTCOSTM.Enabled = Not (EntryMode = "E")

            grpRetail.Visible = True
            Set_Read_Only(grpRetail, True)
        Else
            Clear_Record()
            grdICTIVAL1.Parent = tabSummary.Tabs("Valuation").TabPage
            grpRetail.Visible = False
        End If

    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"APTINVH7", "APTACRC1", "ICTCOSTM", "ICTCOSTH", "ICTCOSTF", "ICTITEM1", "APTINVH5", "ICTIRECX", "POTORDRX", "ICTIVAR1", "ICTIVARX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
        Setup_tabSummary()
        Absx1.txtFor("ITEM_CODE").Text = ""

        ASCMAIN1.sql = "Select * from ICTCOSTF where ITEM_EXP_IMP_IND = 'E'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            rowICTCOSTM = Add_ICTCOSTM(2, row)
        Next
        Sort_grdColumns(grdICTCOSTM, "INIT_DATE".ToLower)

        tabSummary.SelectedTab = tabSummary.Tabs("Recently Viewed")
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)
        ITEM_CODE = Absx1.txtFor("ITEM_CODE").Text

        '' Dim rowICTCOSTP1 As DataRow = dst.Tables("ICTCOSTP1").Rows.Find(New String() {"202604", ITEM_CODE})
        '' Modify the below to use parms
        'ASCMAIN1.sql = $"select * from ICTCOSTP where item_code=:PARM1 and OPS_YYYYPP=:PARM2"
        'Dim rowICTCOSTP1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, , "VV", New String() {ITEM_CODE, ASCMAIN1.CYP})
        'If rowICTCOSTP1 IsNot Nothing Then
        '    numVCost.Value = Val(rowICTCOSTP1.Item("ITEM_COST_VCOST"))
        'Else
        '    numVCost.Value = DBNull.Value
        'End If

        EnforceConstraints(False)

        rowICTITEM1 = Fill_Record("ICTITEM1", ITEM_CODE)
        Dim ITEM_YYYYPP_PRV_COST As String = rowICTITEM1.Item("ITEM_YYYYPP_PRV_COST") & ""
        Dim ITEM_YYYYPP_CUR_COST As String = rowICTITEM1.Item("ITEM_YYYYPP_CUR_COST") & ""

        dst.Tables("ICTITEMX").Rows.Clear()
        Add_ICTITEMX("ITEM_CATGY_CODE", "Catgy", "ITEM_CATGY_DESC", "ICTCATG1")
        Add_ICTITEMX("COST_CATGY_CODE", "Cost", "COST_CATGY_DESC", "ICTCOST1")
        Add_ICTITEMX("ITEM_BASIC_PROMO", "BP", "", "")
        Add_ICTITEMX("ITEM_SNU_CODE", "SNU", "", "")

        Fill_Records("ICTIVAR1", ITEM_CODE)
        Sort_grdColumns(grdICTIVAR1, "OPS_YYYYPP".ToLower)

        Fill_Records("POTORDRX", ITEM_CODE)
        Sort_grdColumns(grdPOTORDRX, "PO_ORDER_NO".ToLower)

        Fill_Records("ICTIRECX", ITEM_CODE)
        Sort_grdColumns(grdICTIRECX, "RECEIPT_DATE".ToLower)
        Fill_Records("APTINVH5", ITEM_CODE)

        Fill_Records("APTACRC1", ITEM_CODE)
        Fill_Records("APTINVH7", ITEM_CODE)

        Load_ICTSTATX("", ITEM_CODE)
        ASCMAIN1.sql = Replace(Replace(sqlICTIVAL1, ":PARM1", "ICTSTATX.OPS_YYYYPP"), ":PARM2", "PERIOD_CALC(ICTSTATX.OPS_YYYYPP,-1)")
        Fill_Records("ICTIVAL1", "", True, ASCMAIN1.sql)
        Sort_grdColumns(grdICTIVAL1, "OPS_YYYYPP".ToLower)

        'Fill_Records("ICTCOSTA", ITEM_CODE)
        'Sort_grdColumns(grdICTCOSTH, "OPS_YYYYPP".ToLower)

        Fill_Records("ICTCOSTC", ITEM_CODE)
        Fill_Records("ICTCOSTH", ITEM_CODE)
        Sort_grdColumns(grdICTCOSTH, "OPS_YYYYPP".ToLower)
        Fill_Records("ICTCOSTF", ITEM_CODE)
        Fill_Records("ICTCOSTP1", ITEM_CODE)

        Dim PV_EXP_MTD As Decimal = 0
        Dim PV_EXP_YTD As Decimal = 0
        Dim PV_DEF As Decimal = 0
        Dim MV_EXP_MTD As Decimal = 0
        Dim MV_EXP_YTD As Decimal = 0
        Dim MV_DEF As Decimal = 0
        Dim CV_EXP_MTD As Decimal = 0
        Dim CV_EXP_YTD As Decimal = 0
        Dim RV_EXP_MTD As Decimal = 0
        Dim RV_EXP_YTD As Decimal = 0
        Dim FV_EXP_MTD As Decimal = 0
        Dim FV_EXP_YTD As Decimal = 0
        Dim FV_DEF As Decimal = 0
        Dim TV_EXP_MTD As Decimal = 0
        Dim TV_EXP_YTD As Decimal = 0
        Dim TV_DEF As Decimal = 0

        Dim PV_DEF_UM As Decimal = 0
        Dim MV_DEF_UM As Decimal = 0
        Dim FV_DEF_UM As Decimal = 0
        Dim TV_DEF_UM As Decimal = 0

        Dim LYP As String = ASCMAIN1.Period_Calc(Mid(ASCMAIN1.CYP, 1, 4) & "01", -1)
        Dim YLP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)
        For Each rowICTIVAR1 As DataRow In dst.Tables("ICTIVAR1").Select($"OPS_YYYYPP >= '{LYP}'", "OPS_YYYYPP")
            With rowICTIVAR1
                Dim YP As String = .Item("OPS_YYYYPP")
                If YP = YLP Then
                    PV_DEF = Val(.Item("PV_DEF") & "")
                    MV_DEF = Val(.Item("MV_DEF") & "")
                    FV_DEF = Val(.Item("FV_DEF") & "")
                    TV_DEF = Val(.Item("TV_DEF") & "")
                    ASCMAIN1.sql = $"Select Sum (WHSE_QTY_ON_HAND) from ICTSTAT5 where ITEM_CODE = '{ITEM_CODE}' and OPS_YYYYPP = '{YLP}'"
                    Dim OH As Int32 = Val(ASCDATA1.GetDataValue)
                    If OH <> 0 Then
                        PV_DEF_UM = PV_DEF / OH
                        MV_DEF_UM = MV_DEF / OH
                        FV_DEF_UM = FV_DEF / OH
                        TV_DEF_UM = TV_DEF / OH
                    End If
                End If
                If Mid(YP, 1, 4) = Mid(ASCMAIN1.CYP, 1, 4) Then
                    PV_EXP_YTD += Val(.Item("PV_EXP") & "")
                    MV_EXP_YTD += Val(.Item("MV_EXP") & "")
                    FV_EXP_YTD += Val(.Item("FV_EXP") & "")
                    TV_EXP_YTD += Val(.Item("TV_EXP") & "")
                    If Mid(YP, 5, 2) = Mid(ASCMAIN1.CYP, 5, 2) Then
                        PV_EXP_MTD += Val(.Item("PV_EXP") & "")
                        MV_EXP_MTD += Val(.Item("MV_EXP") & "")
                        FV_EXP_MTD += Val(.Item("FV_EXP") & "")
                        TV_EXP_MTD += Val(.Item("TV_EXP") & "")
                    End If
                End If
            End With
        Next

        dst.Tables("ICTIVARX").Rows.Clear()
        dst.Tables("ICTIVARX").Rows.Add(New Object() {"Mtd", PV_EXP_MTD, MV_EXP_MTD, CV_EXP_MTD, RV_EXP_MTD, FV_EXP_MTD, TV_EXP_MTD})
        dst.Tables("ICTIVARX").Rows.Add(New Object() {"Ytd", PV_EXP_YTD, MV_EXP_YTD, CV_EXP_YTD, RV_EXP_YTD, FV_EXP_YTD, TV_EXP_YTD})
        dst.Tables("ICTIVARX").Rows.Add(New Object() {"Def", PV_DEF, MV_DEF, DBNull.Value, DBNull.Value, FV_DEF, TV_DEF})
        dst.Tables("ICTIVARX").Rows.Add(New Object() {"/UM", PV_DEF_UM, MV_DEF_UM, DBNull.Value, DBNull.Value, FV_DEF_UM, TV_DEF_UM})
        grdICTIVARX.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.Select

        dst.Tables("ICTCOSTM").Rows.Clear()
        rowICTCOSTM = Add_ICTCOSTM(0, dst.Tables("ICTCOSTH").Rows.Find(New String() {ITEM_YYYYPP_PRV_COST, ITEM_CODE}))
        rowICTCOSTM = Add_ICTCOSTM(1, dst.Tables("ICTCOSTC").Rows.Find(New String() {ITEM_CODE}))
        rowICTCOSTM.Item("OPS_YYYYPP") = ITEM_YYYYPP_CUR_COST
        rowICTCOSTM = Add_ICTCOSTM(2, dst.Tables("ICTCOSTF").Rows.Find(New String() {ITEM_CODE}))
        Sort_grdColumns(grdICTCOSTM, "SEQ")
        grdICTCOSTM.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.Select

        For Each grow As UltraWinGrid.UltraGridRow In grdICTCOSTM.Rows
            If (grow.Cells("SEQ").Value = 2 And (EntryMode = "E")) _
            Or (grow.Cells("SEQ").Value = 1 And (EntryMode = "V")) Then
                grdICTCOSTM.ActiveRow = grow
                Exit For
            End If
        Next

        Dim blnCalculate As Boolean = False

        If EntryMode = "E" Then
            rowICTCOSTM = dst.Tables("ICTCOSTM").Select("SEQ = 2")(0)
            If rowICTCOSTM.Item("ITEM_COST_MAKE_BUY") & "" = "" Then
                Dim rowBMTMAIN1 As DataRow = LookUp("BMTMAIN1", ITEM_CODE)
                If rowBMTMAIN1 IsNot Nothing Then
                    rowICTCOSTM.Item("ITEM_COST_MAKE_BUY") = "M"
                Else
                    rowICTCOSTM.Item("ITEM_COST_MAKE_BUY") = "B"
                End If
            End If

            If rowICTCOSTM.Item("ITEM_COST_MAKE_BUY") = "M" Then

                ASCMAIN1.sql = "Select Max (BM_ISSUE_NO) from BMTMAIN2 " & vbCrLf _
                    & " where BM_PROD_ITEM = '" & ITEM_CODE & "' and BM_ISSUE_USE_FOR_STD = '1' and BM_ISSUE_NO <> '00'"
                Dim BMI As String = ASCDATA1.GetDataValue

                If rowICTCOSTM.Item("BM_ISSUE_NO") & "" <> BMI Then
                    blnCalculate = True
                    rowICTCOSTM.Item("BM_ISSUE_NO") = BMI
                End If

            End If

            If rowICTCOSTM.Item("ITEM_COST_FRT_CLASS") & "" = "" Then
                rowICTCOSTM.Item("ITEM_COST_FRT_CLASS") = ROWs("ICTPARM1").Item("IC_PARM_FRT_CLASS")
            End If
        End If

        Setup_MB()


        Dim rowICTITEM1_RECENT As DataRow = dst.Tables("ICTITEM1_RECENT").Rows.Find(ITEM_CODE)
        If rowICTITEM1_RECENT Is Nothing Then
            rowICTITEM1_RECENT = dst.Tables("ICTITEM1_RECENT").NewRow
            rowICTITEM1_RECENT.ItemArray = rowICTITEM1.ItemArray
            dst.Tables("ICTITEM1_RECENT").Rows.Add(rowICTITEM1_RECENT)
        End If

        lblCostPending.Visible = (rowICTITEM1.Item("ITEM_COST_STATUS") & "" = "P")

        EnforceConstraints(True)

        If blnCalculate Then Calculate_Cost()

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        Dim ITEM_COST_STATUS As String = rowICTITEM1.Item("ITEM_COST_STATUS") & ""
        Dim TABLE_NAME_ICTCOSTX As String = IIf(ITEM_COST_STATUS = "P", "ICTCOSTC", "ICTCOSTF")

        dst.Tables(TABLE_NAME_ICTCOSTX).Rows.Clear()
        Dim rowICTCOSTX As DataRow = dst.Tables(TABLE_NAME_ICTCOSTX).NewRow
        For Each dcol As DataColumn In dst.Tables(TABLE_NAME_ICTCOSTX).Columns
            rowICTCOSTX.Item(dcol.ColumnName) = rowICTCOSTM.Item(dcol.ColumnName)
        Next
        dst.Tables(TABLE_NAME_ICTCOSTX).Rows.Add(rowICTCOSTX)
        Update_Record_TDA(TABLE_NAME_ICTCOSTX, IIf(TABLE_NAME_ICTCOSTX = "ICTCOSTC", "", "ITEM_CODE = '" & ITEM_CODE & "'"))

        If ITEM_COST_STATUS = "P" Then
            ASCMAIN1.sql = "" _
                 & "Begin " & vbCrLf _
                 & " Declare Cursor C1 is Select * from ICTCOSTC where ITEM_CODE = '" & ITEM_CODE & "';" & vbCrLf _
                 & " Begin " & vbCrLf _
                 & "  For R1 in C1 Loop" & vbCrLf _
                 & "   Update ICTITEM1 " & vbCrLf _
                 & "   Set ITEM_COST_STATUS = NULL" & vbCrLf _
                 & ", ITEM_COST_STD = R1.ITEM_COST_TOTAL" & vbCrLf _
                 & ", ITEM_COST_MAKE_BUY = R1.ITEM_COST_MAKE_BUY" & vbCrLf _
                 & ", ITEM_COST_CURR_CODE = R1.ITEM_COST_CURR_CODE" & vbCrLf _
                 & ", ITEM_COST_FRT_CLASS = R1.ITEM_COST_FRT_CLASS" & vbCrLf _
                 & ", ITEM_COST_WASTE_PCT = R1.ITEM_COST_WASTE_PCT" & vbCrLf _
                 & ", ITEM_YYYYPP_CUR_COST = '" & ASCMAIN1.CYP & "'" _
                 & "    where ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
                 & "  End Loop;" & vbCrLf _
                 & " End;" & vbCrLf _
                 & "End;"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, ICTCOSTC.* from ICTCOSTC where ITEM_CODE = '" & ITEM_CODE & "'"
            ASCDATA1.ExecuteSQL("Insert into ICTCOSTA " & ASCMAIN1.sql)
            ASCDATA1.ExecuteSQL("Insert into ICTCOSTH " & ASCMAIN1.sql)
            lblCostPending.Visible = False
        End If

        CommitTrans("Update Complete")

    End Sub

    Sub Delete_Record()
        Call BeginTrans()
        Stop
        'Call Delete_Records("table")
        Call CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub


    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            Case "BM_ISSUE_NO"
                sql_where = "BM_PROD_ITEM = '" & Absx1.txtFor("ITEM_CODE").Text & "'"

        End Select
    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

            Case "View"
                Absx1.txtFor("ITEM_CODE").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()

        Load_Popup_Menu(grdICTCOSTP, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdICTIVAL1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Demand Planning")
        Load_Popup_Menu(grdPOTORDRX, "B", "PO Inquiry")
        Load_Popup_Menu(grdICTCOSTH, "B", "BM Inquiry")
        Load_Popup_Menu(grdICTCOSTM, "B", "BM Inquiry")
        Load_Popup_Menu(grdICTIRECX, "BBBB", "PO Inquiry", "Voucher Inquiry", "BM Inquiry", "PO Receipts Inquiry")
        Load_Popup_Menu(grdICTITEM1_Recent, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Item Status Inquiry")

    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdICTIRECX"

                    tlb_btn = DirectCast(tlb_pop.Tools("Voucher Inquiry"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = grd.ActiveRow IsNot Nothing AndAlso (grd.ActiveRow.Band.Key = "ICTIRECX_APTINVH5" Or grd.ActiveRow.Band.Key = "ICTIRECX_APTACRC1")
                    tlb_btn = DirectCast(tlb_pop.Tools("PO Inquiry"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.Band.Key = "ICTIRECX"
                    tlb_btn = DirectCast(tlb_pop.Tools("PO Receipts Inquiry"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.Band.Key = "ICTIRECX"
                    tlb_btn = DirectCast(tlb_pop.Tools("BM Inquiry"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.Band.Key = "ICTIRECX"

                    'Case "grdARTSTMT1"
                    '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                    '        e.Cancel = True
                    '    End If

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Demand Planning"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                Dim STYLE_CODE As String = rowICTITEM1.Item("STYLE_CODE") & ""
                If STYLE_CODE <> "" Then
                    Context_Launch("Load", STYLE_CODE, e.Tool.Key, "DPFPLAN1")
                End If

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("Select", ITEM_CODE, e.Tool.Key, "ICTSTAT1")
                End If

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI")

            Case "PO Receipts Inquiry"
                Dim RECEIPT_NO As String = grd.ActiveRow.Cells("RECEIPT_NO").Text
                Context_Launch("View", RECEIPT_NO, e.Tool.Key, "ICFIRECI")

            Case "BM Inquiry"
                Dim BM_ISSUE_NO As String = grd.ActiveRow.Cells("BM_ISSUE_NO").Text
                If BM_ISSUE_NO <> "" Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "BMFMAINI")
                End If

            Case "Voucher Inquiry"
                Dim VOUCHER_NO As String = grd.ActiveRow.Cells("VOUCHER_NO").Text
                If VOUCHER_NO = "" And grd.ActiveRow.Band.Key = "ICTIRECX_APTACRC1" Then
                    VOUCHER_NO = grd.ActiveRow.Cells("VOUCHER_NO_MATCHED").Text
                End If
                If VOUCHER_NO <> "" Then
                    Context_Launch("Load", VOUCHER_NO, e.Tool.Key, "APFINVHI")
                End If
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "ITEM_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "ITEM_CODE"
                Call Click_Command("View")
        End Select
    End Sub
#End Region

    Sub Load_ICTSTATX(OPS_YYYYPP As String, ITEM_CODE As String)

        ASCDATA1.ExecuteSQL("Truncate Table " & ICTSTATX)

        ASCMAIN1.sql = "" _
        & "Insert into " & ICTSTATX _
        & " Select ITEM_CODE, OPS_YYYYPP" _
        & ", SUM (WHSE_QTY_BEG) WHSE_QTY_BEG" _
        & ", SUM (WHSE_QTY_SHP) WHSE_QTY_SHP" _
        & ", SUM (WHSE_QTY_RTN) WHSE_QTY_RTN" _
        & ", SUM (WHSE_QTY_REC) WHSE_QTY_REC" _
        & ", SUM (WHSE_QTY_ADJ) WHSE_QTY_ADJ" _
        & ", SUM (WHSE_QTY_XFR) WHSE_QTY_XFR" _
        & ", SUM (WHSE_QTY_CON) WHSE_QTY_CON" _
        & ", SUM (WHSE_QTY_RTV) WHSE_QTY_RTV" _
        & ", SUM (WHSE_QTY_PHY) WHSE_QTY_PHY" _
        & ", SUM (WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND" _
        & ", SUM (WHSE_QTY_ONPO) WHSE_QTY_ONPO" _
        & ", SUM (WHSE_QTY_PLAN) WHSE_QTY_PLAN" _
        & ", SUM (WHSE_QTY_OPEN) WHSE_QTY_OPEN" _
        & ", SUM (WHSE_QTY_PICK) WHSE_QTY_PICK" _
        & ", SUM (WHSE_QTY_COMM) WHSE_QTY_COMM" _
        & ", SUM (WHSE_QTY_HOLD) WHSE_QTY_HOLD" _
        & " from (" _
        & "Select ITEM_CODE, OPS_YYYYPP" _
        & ", SUM (WHSE_QTY_BEG) WHSE_QTY_BEG" _
        & ", SUM (WHSE_QTY_SHP) WHSE_QTY_SHP" _
        & ", SUM (WHSE_QTY_RTN) WHSE_QTY_RTN" _
        & ", SUM (WHSE_QTY_REC) WHSE_QTY_REC" _
        & ", SUM (WHSE_QTY_ADJ) WHSE_QTY_ADJ" _
        & ", SUM (WHSE_QTY_XFR) WHSE_QTY_XFR" _
        & ", SUM (WHSE_QTY_CON) WHSE_QTY_CON" _
        & ", SUM (WHSE_QTY_RTV) WHSE_QTY_RTV" _
        & ", SUM (WHSE_QTY_PHY) WHSE_QTY_PHY" _
        & ", 0 WHSE_QTY_ON_HAND" _
        & ", 0 WHSE_QTY_ONPO" _
        & ", 0 WHSE_QTY_PLAN" _
        & ", 0 WHSE_QTY_OPEN" _
        & ", 0 WHSE_QTY_PICK" _
        & ", 0 WHSE_QTY_COMM" _
        & ", 0 WHSE_QTY_HOLD" _
        & " from ICTSTAT1 " & IIf(OPS_YYYYPP <> "", _
                                        "where OPS_YYYYPP = '" & OPS_YYYYPP & "'", _
                                        "where ITEM_CODE = '" & ITEM_CODE & "'") _
        & " group by ITEM_CODE, OPS_YYYYPP" _
        & " union " _
        & "Select ITEM_CODE, OPS_YYYYPP" _
        & ", 0 WHSE_QTY_BEG" _
        & ", 0 WHSE_QTY_SHP" _
        & ", 0 WHSE_QTY_RTN" _
        & ", 0 WHSE_QTY_REC" _
        & ", 0 WHSE_QTY_ADJ" _
        & ", 0 WHSE_QTY_XFR" _
        & ", 0 WHSE_QTY_CON" _
        & ", 0 WHSE_QTY_RTV" _
        & ", 0 WHSE_QTY_PHY" _
        & ", SUM (WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND" _
        & ", SUM (WHSE_QTY_ONPO) WHSE_QTY_ONPO" _
        & ", SUM (WHSE_QTY_PLAN) WHSE_QTY_PLAN" _
        & ", SUM (WHSE_QTY_OPEN) WHSE_QTY_OPEN" _
        & ", SUM (WHSE_QTY_PICK) WHSE_QTY_PICK" _
        & ", SUM (WHSE_QTY_COMM) WHSE_QTY_COMM" _
        & ", SUM (WHSE_QTY_HOLD) WHSE_QTY_HOLD" _
        & " from ICTSTAT5 " & IIf(OPS_YYYYPP <> "", _
                                        "where OPS_YYYYPP = '" & OPS_YYYYPP & "'", _
                                        "where ITEM_CODE = '" & ITEM_CODE & "'") _
        & " group by ITEM_CODE, OPS_YYYYPP" _
        & " union " _
        & "Select ITEM_CODE, '" & ASCMAIN1.CYP & "' OPS_YYYYPP" _
        & ", 0 WHSE_QTY_BEG" _
        & ", 0 WHSE_QTY_SHP" _
        & ", 0 WHSE_QTY_RTN" _
        & ", 0 WHSE_QTY_REC" _
        & ", 0 WHSE_QTY_ADJ" _
        & ", 0 WHSE_QTY_XFR" _
        & ", 0 WHSE_QTY_CON" _
        & ", 0 WHSE_QTY_RTV" _
        & ", 0 WHSE_QTY_PHY" _
        & ", SUM (WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND" _
        & ", SUM (WHSE_QTY_ONPO) WHSE_QTY_ONPO" _
        & ", SUM (WHSE_QTY_PLAN) WHSE_QTY_PLAN" _
        & ", SUM (WHSE_QTY_OPEN) WHSE_QTY_OPEN" _
        & ", SUM (WHSE_QTY_PICK) WHSE_QTY_PICK" _
        & ", SUM (WHSE_QTY_COMM) WHSE_QTY_COMM" _
        & ", SUM (WHSE_QTY_HOLD) WHSE_QTY_HOLD" _
        & " from ICTSTAT2 " & IIf(OPS_YYYYPP <> "", _
                                        "where '" & ASCMAIN1.CYP & "' = '" & OPS_YYYYPP & "'", _
                                        "where ITEM_CODE = '" & ITEM_CODE & "'") _
        & " group by ITEM_CODE" _
        & ") group by ITEM_CODE, OPS_YYYYPP"
        ASCDATA1.ExecuteSQL()

    End Sub

    Private Sub tabDetails_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabDetails.SelectedTabChanged
        Setup_tabDetails()
    End Sub

    Sub Setup_tabDetails()
        If SELECTION_NO = 0 Then Exit Sub
        'UltraExplorerBar1.Groups("Invoices").Visible = (tabDetails.SelectedTab.Key = "Invoices") And ScreenMode
    End Sub

    Private Sub cmdFetchHistory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFetchHistory.Click
        Dim YP1 As String = cbeYP1.Value
        Dim YP2 As String = cbeYP2.Value

        Dim NP As Int16 = ASCMAIN1.Period_Diff(YP1, YP2) + 1
        If NP > 13 Or NP < 2 Then
            MsgBox("Range of Periods must be from 2 to 13 periods", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Fetching Cost History")

        Dim sql As String = sqlICTCOSTP
        Dim YP As String = ASCMAIN1.Period_Calc(YP1, -1)
        For i As Int16 = 1 To NP
            YP = ASCMAIN1.Period_Calc(YP, 1)
            sql = Replace(sql, "XXXX" & Format(i, "00"), YP)
            Dim LEGEND As String = ASCMAIN1.Get_Legend(YP, False, True)
            With grdICTCOSTP.DisplayLayout.Bands(0).Columns("ITEM_COST_P" & Format(i, "00"))
                .Header.Appearance.BackColor = Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Hidden = False
                .Header.Caption = LEGEND
                .Width = 80
                .Format = "###.0000"
            End With
            With grdICTCOSTP.DisplayLayout.Bands(0).Columns("BOM_P" & Format(i, "00"))
                .Header.Appearance.BackColor = Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Hidden = False
                .Header.Caption = LEGEND & vbCrLf & "Beg OH"
                .Width = 80
                .Format = "#,##0"
            End With
            If i > 1 Then
                With grdICTCOSTP.DisplayLayout.Bands(0).Columns("VAR_P" & Format(i, "00"))
                    .Header.Appearance.BackColor = Color.White
                    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    .Hidden = False
                    .Header.Caption = LEGEND & vbCrLf & "Reval"
                    .Width = 80
                    .Format = "#,##0"
                End With
            End If

        Next
        sql = Replace(sql, "XXXX13", YP2)

        With grdICTCOSTP.DisplayLayout.Bands(0).Columns("COST_VAR")
            .Header.Appearance.BackColor = Color.White
            .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            .Header.Appearance.BackColor2 = Color.Pink
        End With
        With grdICTCOSTP.DisplayLayout.Bands(0).Columns("COST_VAR_PCT")
            .Header.Appearance.BackColor = Color.White
            .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            .Header.Appearance.BackColor2 = Color.Pink
        End With


        If NP < 13 Then
            sql = Replace(sql, "XXXX13", YP2)
            For I As Int16 = NP + 1 To 13
                With grdICTCOSTP.DisplayLayout.Bands(0)
                    .Columns("ITEM_COST_P" & Format(I, "00")).Hidden = True
                    .Columns("BOM_P" & Format(I, "00")).Hidden = True
                    .Columns("VAR_P" & Format(I, "00")).Hidden = True
                End With
            Next
        End If

        With dst.Tables("ICTCOSTP")
            .Columns("COST_VAR").Expression = "ITEM_COST_P" & Format(NP, "00") & " - ITEM_COST_P" & Format(NP - 1, "00")
            .Columns("COST_VAR_PCT").Expression = "IIF(ISNULL(ITEM_COST_P" & Format(NP - 1, "00") & ",0) = 0, 0, 100 * COST_VAR / ITEM_COST_P" & Format(NP - 1, "00" & ")")
        End With

        ASCDATA1.ExecuteSQL("Truncate Table " & ICTCOSTP)
        ASCDATA1.ExecuteSQL("Insert into " & ICTCOSTP & " " & sql)

        Fill_Records("ICTCOSTP")
        Setup_ICTCOSTP()
        grdICTCOSTP.Visible = True
        chkNonZeroVarianceOnly.Visible = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub cmdValuation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdValuation.Click
        Dim YP As String = cbeYP.Value
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Fetching Inventory Valuation")

        Dim RYP As String = cbeYP.Value
        Dim PYP As String = ASCMAIN1.Period_Calc(RYP, -1)
        Load_ICTSTATX(RYP, "")

        Fill_Records("ICTIVAL1", New String() {RYP, PYP})
        grdICTIVAL1.Visible = True
        ' splICTCOSTP.Panel2Collapsed = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Print_Record()

        Print_Report_Begin()
        CR_params.Add("SUBT", "")
        Dim RPT As String = "ICRCOST1"
        Generate_Report(RPT, "Cost Maintenance")
    End Sub

    Sub Setup_tabSummary()
        With UltraExplorerBar1
            .Groups("Cost History").Visible = (tabSummary.SelectedTab.Key = "Cost History")
            .Groups("Valuation").Visible = (tabSummary.SelectedTab.Key = "Valuation")
        End With
    End Sub
    Private Sub tabSummary_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabSummary.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabSummary()
    End Sub

    Private Sub grdICTCOSTP_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTCOSTP.DoubleClickRow
        If Not ScreenMode Then
            Absx1.txtFor("ITEM_CODE").Text = e.Row.Cells("ITEM_CODE").Value
            Click_Command("View")
        End If
    End Sub
     
    Private Sub grdICTIVAL1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTIVAL1.DoubleClickRow
        If Not ScreenMode Then
            Absx1.txtFor("ITEM_CODE").Text = e.Row.Cells("ITEM_CODE").Value
            Click_Command("View")
        End If
    End Sub

    Function Add_ICTCOSTM(SEQ As Integer, row As DataRow) As DataRow

        Dim rowICTCOSTM As DataRow = dst.Tables("ICTCOSTM").NewRow
        If row IsNot Nothing Then
            For Each dcol As DataColumn In row.Table.Columns
                rowICTCOSTM.Item(dcol.ColumnName) = row.Item(dcol.ColumnName)
            Next
        Else
            rowICTCOSTM.Item("ITEM_CODE") = ITEM_CODE
        End If

        Dim SEQ_DESC As String = ""
        If SEQ = 0 Then SEQ_DESC = "Prv"
        If SEQ = 1 Then SEQ_DESC = "Cur"
        If SEQ = 2 Then SEQ_DESC = "Fut"

        rowICTCOSTM.Item("SEQ") = SEQ
        rowICTCOSTM.Item("SEQ_DESC") = SEQ_DESC

        If rowICTCOSTM.Item("OPS_YYYYPP") & String.Empty = "" Then
            rowICTCOSTM.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        End If


        dst.Tables("ICTCOSTM").Rows.Add(rowICTCOSTM)
        Return rowICTCOSTM
    End Function

    Private Sub optMB_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optMB.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_MB()
    End Sub

    Sub Setup_MB()
        lblMaterials.Visible = (optMB.Value & "" = "M")
        Absx1.numFor("ITEM_COST_MATLS").Visible = (optMB.Value & "" = "M")
        Absx1.numFor("ITEM_COST_LANDGI").Visible = (optMB.Value & "" = "M")
        Absx1.numFor("ITEM_COST_TOOLGI").Visible = (optMB.Value & "" = "M") And lblTOOLG.Visible
        Absx1.numFor("ITEM_COST_OVRHDI").Visible = (optMB.Value & "" = "M") And lblOVRHD.Visible

        lblBM.Visible = (optMB.Value & "" = "M")
        Absx1.txtFor("BM_ISSUE_NO").Visible = (optMB.Value & "" = "M")
        Absx1.txtFor("BM_ISSUE_COMMENT").Visible = (optMB.Value & "" = "M")

        Absx1.numFor("ITEM_COST_TOTAL_D").Visible = (optMB.Value & "" = "M")
        Absx1.numFor("ITEM_COST_TOTAL_M").Visible = (optMB.Value & "" = "M")
        lblTotal.Visible = (optMB.Value & "" = "M")

        lblMBWarning.Visible = (optMB.Value & "" = "M")

        If (optMB.Value & "" = "M") Then
            ASCMAIN1.sql = "Select Max (BM_ISSUE_NO) from BMTMAIN2 " & vbCrLf _
                & " where BM_PROD_ITEM = '" & ITEM_CODE & "' and BM_ISSUE_USE_FOR_STD = '1' and BM_ISSUE_NO <> '00'"
            Dim BMI As String = ASCDATA1.GetDataValue

            If Absx1.txtFor("BM_ISSUE_NO").Text <> BMI Then

                Absx1.txtFor("BM_ISSUE_NO").Text = BMI
            End If
        End If

    End Sub

    Private Sub cmdCalculateCost_Click(sender As System.Object, e As System.EventArgs) Handles cmdCalculateCost.Click

        Calculate_Cost()

    End Sub

    Sub Calculate_Cost()
        Synch_TABLE_NAME("ICTCOSTM")

        Dim FRT_CLASS_CODE As String = rowICTCOSTM.Item("ITEM_COST_FRT_CLASS") & ""
        Dim rowICTFRTC1 As DataRow = dst.Tables("ICTFRTC1").Rows.Find(FRT_CLASS_CODE)
        Dim FRT_CLASS_PCT_CUR As Decimal = 0
        If rowICTFRTC1 IsNot Nothing Then
            FRT_CLASS_PCT_CUR = Val(rowICTFRTC1.Item("FRT_CLASS_PCT_CUR") & "")
        End If

        Dim TRF_CLASS_CODE As String = rowICTCOSTM.Item("ITEM_COST_TRF_CLASS") & ""
        Dim rowICTTRFC1 As DataRow = dst.Tables("ICTTRFC1").Rows.Find(TRF_CLASS_CODE)
        Dim TRF_CLASS_PCT_CUR As Decimal = 0
        If rowICTTRFC1 IsNot Nothing Then
            TRF_CLASS_PCT_CUR = Val(rowICTTRFC1.Item("TRF_CLASS_PCT_FUT") & "")
        End If

        rowICTCOSTM.Item("ITEM_COST_MATLS") = DBNull.Value
        rowICTCOSTM.Item("ITEM_COST_LANDGI") = DBNull.Value
        rowICTCOSTM.Item("ITEM_COST_TOOLGI") = DBNull.Value
        rowICTCOSTM.Item("ITEM_COST_OVRHDI") = DBNull.Value


        If ASCMAIN1.CLIENT = "AHA" Then
            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            If rowICTITEM1.Item("VEND_CODE") = "DEADSEALAB" Then
                Dim ITEM_COST_VCOST As Decimal = Val(rowICTCOSTM.Item("ITEM_COST_VCURR") & "")
                If ITEM_COST_VCOST <> 0 Then
                    Dim DISC_PCT As Decimal = -5 ' HARD CODED -5%
                    Dim ITEM_COST_TOOLG As Decimal = ITEM_COST_VCOST * DISC_PCT / 100
                    rowICTCOSTM.Item("ITEM_COST_TOOLG") = ITEM_COST_TOOLG
                    lblTOOLG.Text = "Disc " & Format(DISC_PCT, "##0") & "%"
                End If
            End If
        End If

        If optMB.Value = "B" Then
        Else
            Dim BM_ISSUE_NO As String = Absx1.txtFor("BM_ISSUE_NO").Text
            Dim M As String = " * NVL(BMTMAIN3.BM_QTY_PER_ASSY,0) * (100 + NVL(ICTCOSTC.ITEM_COST_WASTE_PCT,0)) / 100"
            ASCMAIN1.sql = "Select " _
                & "  SUM ((NVL(ICTCOSTC.ITEM_COST_VCOST,0)+NVL(ICTCOSTC.ITEM_COST_MATLS,0))" & M & ") ITEM_COST_MATLS" & vbCrLf _
                & ", SUM ((NVL(ICTCOSTC.ITEM_COST_LANDG,0)+NVL(ICTCOSTC.ITEM_COST_LANDGI,0))" & M & ") ITEM_COST_LANDGI" & vbCrLf _
                & ", SUM ((NVL(ICTCOSTC.ITEM_COST_TOOLG,0)+NVL(ICTCOSTC.ITEM_COST_TOOLGI,0))" & M & ") ITEM_COST_TOOLGI" & vbCrLf _
                & ", SUM ((NVL(ICTCOSTC.ITEM_COST_OVRHD,0)+NVL(ICTCOSTC.ITEM_COST_OVRHDI,0))" & M & ") ITEM_COST_OVRHDI" & vbCrLf _
                & " from ICTCOSTC,BMTMAIN3" & vbCrLf _
                & " where BMTMAIN3.BM_PROD_ITEM = '" & ITEM_CODE & "'" & vbCrLf _
                & "   and BMTMAIN3.BM_ISSUE_NO = '" & BM_ISSUE_NO & "'" & vbCrLf _
                & "   and NVL(BMTMAIN3.BM_VEND_SUPP_MATL,'0') <> '1'" & vbCrLf _
                & "   and NVL(BMTMAIN3.BM_WHEN_EXHAUSTED,'?') = '?'" & vbCrLf _
                & "   and ICTCOSTC.ITEM_CODE = BMTMAIN3.BM_COMP_ITEM"
            Dim row As DataRow = ASCDATA1.GetDataRow
            If row IsNot Nothing Then
                rowICTCOSTM.Item("ITEM_COST_MATLS") = row.Item("ITEM_COST_MATLS")
                rowICTCOSTM.Item("ITEM_COST_LANDGI") = row.Item("ITEM_COST_LANDGI")
                rowICTCOSTM.Item("ITEM_COST_TOOLGI") = row.Item("ITEM_COST_TOOLGI")
                rowICTCOSTM.Item("ITEM_COST_OVRHDI") = row.Item("ITEM_COST_OVRHDI")
            End If
        End If

        Dim ITEM_COST_VCURR As Decimal = Val(rowICTCOSTM.Item("ITEM_COST_VCURR") & "")
        rowICTCOSTM.Item("ITEM_COST_VCOST") = ITEM_COST_VCURR
        rowICTCOSTM.Item("ITEM_COST_LANDG") = ITEM_COST_VCURR * FRT_CLASS_PCT_CUR / 100
        rowICTCOSTM.Item("ITEM_COST_TOOLG") = ITEM_COST_VCURR * TRF_CLASS_PCT_CUR / 100

        Dim ITEM_COST_TOTAL As Decimal = Val(rowICTCOSTM.Item("ITEM_COST_VCOST") & "") _
                                       + Val(rowICTCOSTM.Item("ITEM_COST_LANDG") & "") _
                                       + Val(rowICTCOSTM.Item("ITEM_COST_TOOLG") & "") _
                                       + Val(rowICTCOSTM.Item("ITEM_COST_OVRHD") & "") _
                                       + Val(rowICTCOSTM.Item("ITEM_COST_MATLS") & "") _
                                       + Val(rowICTCOSTM.Item("ITEM_COST_LANDGI") & "") _
                                       + Val(rowICTCOSTM.Item("ITEM_COST_TOOLGI") & "") _
                                       + Val(rowICTCOSTM.Item("ITEM_COST_OVRHDI") & "")

        rowICTCOSTM.Item("ITEM_COST_TOTAL") = ITEM_COST_TOTAL
        rowICTCOSTM.Item("ITEM_EXP_IMP_IND") = "E"

        rowICTCOSTM.Item("COLLECTION_CODE") = rowICTITEM1.Item("COLLECTION_CODE")
        rowICTCOSTM.Item("ITEM_CLASS_CODE") = rowICTITEM1.Item("ITEM_CLASS_CODE")
        rowICTCOSTM.Item("COST_CATGY_CODE") = rowICTITEM1.Item("COST_CATGY_CODE")
        rowICTCOSTM.Item("ITEM_CATGY_CODE") = rowICTITEM1.Item("ITEM_CATGY_CODE")
        rowICTCOSTM.Item("ITEM_COST_CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
        rowICTCOSTM.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowICTCOSTM.Item("INIT_DATE") = DATETIME_STAMP

        'BM_ISSUE_NO(VARCHAR2(2))
        'ITEM_COST_WASTE_TYPE(VARCHAR2(1))
        'ITEM_BM_ISSUE_SEL(VARCHAR2(1))

        Absx1.numFor("ITEM_COST_LANDG").Focus()
        Absx1.numFor("ITEM_COST_TOOLG").Focus()
        Absx1.numFor("ITEM_COST_TOTAL").Focus()
        Absx1.numFor("FRT_CLASS_PCT_CUR").Focus()
        Absx1.numFor("TRF_CLASS_PCT_CUR").Focus()

        cmdCalculateCost.Focus()
    End Sub

    Private Sub grdICTITEM1_Recent_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTITEM1_Recent.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("ITEM_CODE").Text = e.Row.Cells("ITEM_CODE").Value
            Click_Command("View")
        End If
    End Sub

    Sub Add_ICTITEMX(COLUMN_NAME As String, COLUMN_CAPTION As String, COLUMN_NAME_DESC As String, TABLE_NAME As String)
        Dim COLUMN_CODE As String = rowICTITEM1.Item(COLUMN_NAME) & ""
        Dim COLUMN_DESC As String = "?"
        If COLUMN_NAME_DESC <> "" Then
            Dim row As DataRow = dst.Tables(TABLE_NAME).Rows.Find(COLUMN_CODE)
            If row IsNot Nothing Then
                COLUMN_DESC = row.Item(COLUMN_NAME_DESC) & ""
            End If
        Else
            Dim row As DataRow = LookUp("ASTCODE1", New String() {"ICTITEM1", COLUMN_NAME, COLUMN_CODE})
            If row IsNot Nothing Then
                COLUMN_DESC = row.Item("T_DESC") & ""
            End If
        End If
        dst.Tables("ICTITEMX").Rows.Add(New String() {COLUMN_NAME, COLUMN_CAPTION, COLUMN_CODE, COLUMN_DESC})
    End Sub

    Private Sub grdICTCOSTM_BeforeRowDeactivate(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles grdICTCOSTM.BeforeRowDeactivate
        If Not IsLoading Then
            If EntryMode = "E" Then
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub grdICTCOSTM_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTCOSTM.InitializeRow
        If EntryMode = "E" And e.Row.Cells("SEQ").Value = 2 Then
            e.Row.Appearance.BackColor = Color.LightGreen
        ElseIf EntryMode = "V" And e.Row.Cells("SEQ").Value = 1 Then
            e.Row.Appearance.BackColor = Color.LightGreen
        End If
    End Sub

    Function Validate_Field(COLUMN_NAME As String, rowICTITEM1 As DataRow, COLUMN_CAPTION As String, TABLE_NAME As String, VALUE_LIST() As String) As String
        Dim EMsg As String = ""

        Dim CODE_VALUE As String = ""
        If rowICTITEM1 Is Nothing Then
            CODE_VALUE = Absx1.txtFor(COLUMN_NAME).Text
        Else
            CODE_VALUE = rowICTITEM1.Item(COLUMN_NAME) & ""
        End If
        If CODE_VALUE = "" Then
            EMsg &= vbCr & COLUMN_CAPTION & " is Mandatory"
        Else
            If VALUE_LIST IsNot Nothing Then
                If Not VALUE_LIST.Contains(CODE_VALUE) Then
                    EMsg &= vbCr & "Invalid Value Specified for " & COLUMN_CAPTION
                End If
            Else
                Dim row As DataRow
                If dst.Tables.Contains(TABLE_NAME) Then
                    row = dst.Tables(TABLE_NAME).Rows.Find(CODE_VALUE)
                Else
                    row = LookUp(TABLE_NAME, CODE_VALUE)
                End If
                If row Is Nothing Then
                    EMsg &= vbCr & "Invalid Value Specified for " & COLUMN_CAPTION
                End If
            End If
        End If

        Return EMsg
    End Function

    Private Sub grdICTIVAL1_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdICTIVAL1.InitializeLayout

    End Sub

    Private Sub grdICTIVAL1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTIVAL1.InitializeRow
        If e.Row.Cells("ITEM_COST_STATUS").Value & "" = "P" Then
            e.Row.Appearance.BackColor = Color.Yellow
            e.Row.ToolTipText = "Item does NOT have a Standard Cost"
        End If
    End Sub

    Private Sub chkNonZeroVarianceOnly_CheckedChanged(sender As Object, e As EventArgs) Handles chkNonZeroVarianceOnly.CheckedChanged
        Setup_ICTCOSTP()
    End Sub

    Sub Setup_ICTCOSTP()
        Dim DVW As DataView = DirectCast(grdICTCOSTP.DataSource, DataTable).DefaultView
        Dim sql As String = ""
        If chkNonZeroVarianceOnly.Checked Then
            sql = "ISNULL(COST_VAR,0) <> 0"
        End If
        DVW.RowFilter = Sql
    End Sub

    Private Sub grdICTCOSTM_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdICTCOSTM.InitializeLayout

    End Sub

    Private Sub grdICTCOSTM_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTCOSTM.DoubleClickRow
        If Not ScreenMode AndAlso e.Row.IsDataRow Then
            Absx1.txtFor("ITEM_CODE").Text = e.Row.Cells("ITEM_CODE").Value
            Click_Command("View")
        End If
    End Sub

    Sub Fix_Variances(YP As String)

        ' THIS CODE (FROM TARPEND1) WAS WIRED INTO THE DONE BUTTON TO FIX ICTIREC2.TRAN_xV AND ICTIVAR1.xV_EXP

        Dim YP_SAVE As String = ASCMAIN1.CYP
        ASCMAIN1.CYP = YP

        ' Set Inventory Cost Variances

        ' Reset TRAN_xV fields (because of Re-Valuation) - ICTIREC2 contributions only

        ASCMAIN1.sql = $"Begin Declare Cursor C1 is
            Select ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO
            , Sum (ICTIREC2.QTY_REC * (ICTIREC2.PO_COST - ICTCOSTA.ITEM_COST_VCOST)) TRAN_PV
            , Sum (NVL(ICTIREC2.EXT_COST_MATLS,0) - ICTIREC2.QTY_REC * (
                    NVL(ICTCOSTA.ITEM_COST_MATLS,0) +
                    NVL(ICTCOSTA.ITEM_COST_LANDGI,0) +
                    NVL(ICTCOSTA.ITEM_COST_TOOLGI,0) +
                    NVL(ICTCOSTA.ITEM_COST_OVRHDI,0)
                    )) TRAN_MV
            , Sum (ICTIREC2.QTY_REC * (ICTIREC2.PO_COST_FRT - ICTCOSTA.ITEM_COST_LANDG)) TRAN_FV
            , Sum (ICTIREC2.QTY_REC * (ICTIREC2.PO_COST_TRF - ICTCOSTA.ITEM_COST_TOOLG)) TRAN_TV
                from ICTIREC2,ICTCOSTA
                where ICTCOSTA.OPS_YYYYPP = ICTIREC2.OPS_YYYYPP
                and ICTCOSTA.ITEM_CODE = ICTIREC2.ITEM_CODE
                and ICTIREC2.OPS_YYYYPP = '{ASCMAIN1.CYP}'
            group by ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO;
            Begin For R1 in C1 Loop
            Update ICTIREC2 Set TRAN_PV = R1.TRAN_PV, TRAN_MV = R1.TRAN_MV, TRAN_FV = R1.TRAN_FV, TRAN_TV = R1.TRAN_TV
            where RECEIPT_NO = R1.RECEIPT_NO and RECEIPT_LNO = R1.RECEIPT_LNO;
            End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

        ' PO Receipts

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Update ICTIVAR1 Set PV_EXP = 0, MV_EXP = 0, CV_EXP = 0, FV_EXP = 0, TV_EXP = 0" & vbCrLf _
            & "  where OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
            & "Begin Declare Cursor C1 is" & vbCrLf _
            & "Select ICTIREC2.ITEM_CODE, ICTIREC2.OPS_YYYYPP" & vbCrLf _
            & ", Sum (ICTIREC2.TRAN_PV) TRAN_PV" & vbCrLf _
            & ", Sum (ICTIREC2.TRAN_MV) TRAN_MV" & vbCrLf _
            & ", Sum (ICTIREC2.TRAN_CV) TRAN_CV" & vbCrLf _
            & ", Sum (ICTIREC2.TRAN_FV) TRAN_FV" & vbCrLf _
            & ", Sum (ICTIREC2.TRAN_TV) TRAN_TV" & vbCrLf _
            & " from ICTIREC2,ICTCOSTA" & vbCrLf _
            & " where (NVL(TRAN_PV,0) <> 0 " & vbCrLf _
            & "     or NVL(TRAN_MV,0) <> 0 or NVL(TRAN_CV,0) <> 0 or NVL(TRAN_FV,0) <> 0 or NVL(TRAN_TV,0) <> 0)" & vbCrLf _
            & "   and ICTCOSTA.OPS_YYYYPP = ICTIREC2.OPS_YYYYPP" & vbCrLf _
            & "   and ICTCOSTA.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf _
            & "   and ICTIREC2.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "group by ICTIREC2.ITEM_CODE, ICTIREC2.OPS_YYYYPP;" & vbCrLf _
            & "Begin For R1 in C1 Loop" & vbCrLf _
            & " Update ICTIVAR1 " & vbCrLf _
            & "  Set PV_EXP = R1.TRAN_PV, MV_EXP = R1.TRAN_MV, CV_EXP = R1.TRAN_CV" & vbCrLf _
            & "    , FV_EXP = R1.TRAN_FV, TV_EXP = R1.TRAN_TV" & vbCrLf _
            & "  where ITEM_CODE = R1.ITEM_CODE AND OPS_YYYYPP = R1.OPS_YYYYPP;" & vbCrLf _
            & " If SQL%NOTFOUND Then" & vbCrLf _
            & "  Insert into ICTIVAR1 (ITEM_CODE, OPS_YYYYPP, PV_EXP, MV_EXP, CV_EXP, FV_EXP)" & vbCrLf _
            & "   values (R1.ITEM_CODE, R1.OPS_YYYYPP, R1.TRAN_PV, R1.TRAN_MV, R1.TRAN_CV, R1.TRAN_FV);" & vbCrLf _
            & " End If;" & vbCrLf _
            & "End Loop; End; End; End;" & vbCrLf
        ASCDATA1.ExecuteSQL()

        ' PV from AP Invoicing

        Dim sqlAPTINVH5 As String = "" _
            & "Begin Declare Cursor C1 is" & vbCrLf _
            & "Select ICTIREC2.ITEM_CODE, APTINVH1.OPS_YYYYPP" & vbCrLf _
            & ", Sum (APTINVH5.VAR_AMT) PV" & vbCrLf _
            & " from APTINVH5,ICTIREC2,APTINVH1" & vbCrLf _
            & " where ICTIREC2.RECEIPT_NO = APTINVH5.RECEIPT_NO" & vbCrLf _
            & "   and ICTIREC2.RECEIPT_LNO = APTINVH5.RECEIPT_LNO" & vbCrLf _
            & "   and APTINVH1.REGISTER_IND = '1'" & vbCrLf _
            & "   and APTINVH5.VAR_AMT <> 0" & vbCrLf _
            & "   and APTINVH1.VOUCHER_NO = APTINVH5.VOUCHER_NO" & vbCrLf _
            & "   and APTINVH1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & " group by ICTIREC2.ITEM_CODE, APTINVH1.OPS_YYYYPP;" & vbCrLf

        ASCMAIN1.sql = sqlAPTINVH5 _
            & "Begin For R1 in C1 Loop" & vbCrLf _
            & "Update ICTIVAR1 " & vbCrLf _
            & "Set PV_EXP = NVL(PV_EXP,0) + R1.PV" & vbCrLf _
            & " where ITEM_CODE = R1.ITEM_CODE and OPS_YYYYPP = R1.OPS_YYYYPP;" & vbCrLf _
            & "If SQL%NOTFOUND THEN" & vbCrLf _
            & "Insert into ICTIVAR1 (ITEM_CODE, OPS_YYYYPP, PV_EXP)" & vbCrLf _
            & " Values (R1.ITEM_CODE, R1.OPS_YYYYPP, R1.PV);" & vbCrLf _
            & "End If;" & vbCrLf _
            & "End Loop; End; End;" & vbCrLf
        ASCDATA1.ExecuteSQL()

        sqlAPTINVH5 = Replace(sqlAPTINVH5, "ICTIREC2.ITEM_CODE, APTINVH1.OPS_YYYYPP", "ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO")
        ASCMAIN1.sql = sqlAPTINVH5 _
            & "Begin For R1 in C1 Loop" & vbCrLf _
            & "Update ICTIREC2 " & vbCrLf _
            & "Set TRAN_PV = NVL(TRAN_PV,0) + R1.PV" & vbCrLf _
            & " where RECEIPT_NO = R1.RECEIPT_NO and RECEIPT_LNO = R1.RECEIPT_LNO;" & vbCrLf _
            & "End Loop; End; End;" & vbCrLf
        ASCDATA1.ExecuteSQL()


        ' FV & TV from Invoicing

        Dim sqlAPTACRC1 As String = "" _
            & "Begin Declare Cursor C1 is" & vbCrLf _
            & "Select ITEM_CODE, OPS_YYYYPP" & vbCrLf _
            & ", Sum (CASE WHEN ACCRUAL_CODE = 'FRT' THEN VAR ELSE 0 END) FV" & vbCrLf _
            & ", Sum (CASE WHEN ACCRUAL_CODE = 'TRF' THEN VAR ELSE 0 END) TV" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select APTACRC1.ITEM_CODE, APTINVH1.OPS_YYYYPP, APTACRC1.ACCRUAL_CODE" & vbCrLf _
            & " , APTACRC1.RECEIPT_NO, APTACRC1.RECEIPT_LNO" & vbCrLf _
            & ", Sum (NVL(APTINVH7.TOTAL_INV,0) - NVL(APTINVH7.TOTAL_ACC,0)) VAR" & vbCrLf _
            & " from APTINVH7,APTACRC1,APTINVH1" & vbCrLf _
            & " where APTACRC1.CTL_NO = APTINVH7.CTL_NO" & vbCrLf _
            & "   and APTINVH1.REGISTER_IND = '1'" & vbCrLf _
            & "   and (APTACRC1.ACCRUAL_CODE = 'FRT' or APTACRC1.ACCRUAL_CODE = 'TRF')" & vbCrLf _
            & "   and NVL(APTACRC1.PPD_IND,'0') = '0' and PPD_MATCHED_XNO is Null" & vbCrLf _
            & "   and APTACRC1.ITEM_CODE is Not Null" & vbCrLf _
            & "   and NVL(APTINVH7.TOTAL_INV,0) - NVL(APTINVH7.TOTAL_ACC,0) <> 0" & vbCrLf _
            & "   and APTINVH1.VOUCHER_NO = APTINVH7.VOUCHER_NO" & vbCrLf _
            & "   and APTINVH1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & " group by APTACRC1.ITEM_CODE, APTINVH1.OPS_YYYYPP, APTACRC1.ACCRUAL_CODE" & vbCrLf _
            & ", APTACRC1.RECEIPT_NO, APTACRC1.RECEIPT_LNO)" & vbCrLf _
            & " group by ITEM_CODE, OPS_YYYYPP;" & vbCrLf

        ASCMAIN1.sql = sqlAPTACRC1 _
            & "Begin For R1 in C1 Loop" & vbCrLf _
            & " Update ICTIVAR1 " & vbCrLf _
            & "  Set FV_EXP = NVL(FV_EXP,0) + R1.FV, TV_EXP = NVL(TV_EXP,0) + R1.TV" & vbCrLf _
            & "   where ITEM_CODE = R1.ITEM_CODE and OPS_YYYYPP = R1.OPS_YYYYPP;" & vbCrLf _
            & " If SQL%NOTFOUND Then" & vbCrLf _
            & "  Insert into ICTIVAR1 (ITEM_CODE, OPS_YYYYPP, FV_EXP, TV_EXP)" & vbCrLf _
            & "   Values (R1.ITEM_CODE, R1.OPS_YYYYPP, R1.FV, R1.TV);" & vbCrLf _
            & " End If;" & vbCrLf _
            & "End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

        sqlAPTACRC1 = Replace(sqlAPTACRC1, "ITEM_CODE, OPS_YYYYPP", "RECEIPT_NO, RECEIPT_LNO")
        ASCMAIN1.sql = sqlAPTACRC1 _
            & "Begin For R1 in C1 Loop" & vbCrLf _
            & " Update ICTIREC2 " & vbCrLf _
            & "  Set TRAN_FV = NVL(TRAN_FV,0) + R1.FV, TRAN_TV = NVL(TRAN_TV,0) + R1.TV" & vbCrLf _
            & "   where RECEIPT_NO = R1.RECEIPT_NO and RECEIPT_LNO = R1.RECEIPT_LNO;" & vbCrLf _
            & "End Loop; End; End;"
        ASCDATA1.ExecuteSQL()


        ' TV from PPD Match

        Dim sqlAPTACRCM As String = "" _
            & "Begin Declare Cursor C1 is" & vbCrLf _
            & "Select ITEM_CODE, OPS_YYYYPP_MATCHED, Sum (COST_VAR_ITEM) TV" & vbCrLf _
            & " from APTACRC1" & vbCrLf _
            & " where OPS_YYYYPP_MATCHED = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "   and NVL(PPD_IND,'0') = '0'" & vbCrLf _
            & " group by ITEM_CODE, OPS_YYYYPP_MATCHED;" & vbCrLf

        ASCMAIN1.sql = sqlAPTACRCM _
            & "Begin For R1 in C1 Loop" & vbCrLf _
            & " Update ICTIVAR1 " & vbCrLf _
            & "  Set TV_EXP = NVL(TV_EXP,0) + R1.TV" & vbCrLf _
            & "   where ITEM_CODE = R1.ITEM_CODE and OPS_YYYYPP = R1.OPS_YYYYPP_MATCHED;" & vbCrLf _
            & " If SQL%NOTFOUND Then" & vbCrLf _
            & "  Insert into ICTIVAR1 (ITEM_CODE, OPS_YYYYPP, TV_EXP)" & vbCrLf _
            & "   Values (R1.ITEM_CODE, R1.OPS_YYYYPP_MATCHED, R1.TV);" & vbCrLf _
            & " End If;" & vbCrLf _
            & "End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

        sqlAPTACRCM = Replace(sqlAPTACRCM, "ITEM_CODE, OPS_YYYYPP_MATCHED", "RECEIPT_NO, RECEIPT_LNO")
        ASCMAIN1.sql = sqlAPTACRCM _
            & "Begin For R1 in C1 Loop" & vbCrLf _
            & " Update ICTIREC2 " & vbCrLf _
            & "  Set TRAN_TV = NVL(TRAN_TV,0) + R1.TV" & vbCrLf _
            & "   where RECEIPT_NO = R1.RECEIPT_NO and RECEIPT_LNO = R1.RECEIPT_LNO;" & vbCrLf _
            & "End Loop; End; End;"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.CYP = YP_SAVE
    End Sub



    Private Sub UltraNumericEditor17_ValueChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub UltraLabel10_Click(sender As Object, e As EventArgs)

    End Sub
End Class