Public Class SOFSCOM1
    Dim CUST_CODE As String
    Dim rowARTCUST1 As DataRow
    Dim SREP_CODE As String
    Dim rowSOTSREP1 As DataRow
    Dim OPS_YYYYPP As String
    Dim LYP As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If MENU_ITEM_OBJECT = "SOFSCOMI" Then
            InquiryMode = True
        End If

        With dst

            ASCMAIN1.sql = "Select SOTINVH1.*, ARTCUST2.CUST_STORE_STATE" _
                & " from SOTINVH1, ARTCUST2" _
                & " where SOTINVH1.CUST_CODE = :PARM1 and SOTINVH1.ORDR_YYYYPP_UPDATED = :PARM2" _
                & " and SOTINVH1.CUST_CODE = ARTCUST2.CUST_CODE (+)  AND SOTINVH1.CUST_STORE_NO = ARTCUST2.CUST_STORE_NO (+)"
            Create_TDA(.Tables.Add, "SOTINVH1", "**", 0, True, "VV", 2, "SREP_COMM_PCT,SREP_CODE,SREP_COMM_ADJ,INV_COMMENT,SREP_COMM_XNO,SREP_COMM_IND")
            .Tables("SOTINVH1").Columns.Add("SREP_COMM_CALC", GetType(System.Decimal), "ISNULL(INV_SALES,0) * ISNULL(SREP_COMM_PCT,0) / 100")
            .Tables("SOTINVH1").Columns.Add("SREP_COMM", GetType(System.Decimal), "ISNULL(SREP_COMM_CALC,0) + ISNULL(SREP_COMM_ADJ,0)")

            ASCMAIN1.sql = "Select SOTSCOM2.*" _
                & " from SOTSCOM2 where NVL(SOTSCOM2.OPS_YYYYPP,'" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1) & "') = :PARM1"
            Create_TDA(.Tables.Add, "SOTSCOM2", "**", 0, True, "V", 1)

            ASCMAIN1.sql = "Select SOTSCOMO.*, ARTCUST2.CUST_STORE_NAME, ARTCUST2.CUST_STORE_STATE" _
                & " from SOTSCOMO, ARTCUST2" _
                & " where SOTSCOMO.CUST_CODE = :PARM1 AND SOTSCOMO.OPS_YYYYPP = :PARM2" _
                & " and SOTSCOMO.CUST_CODE = ARTCUST2.CUST_CODE (+) AND SOTSCOMO.CUST_STORE_NO = ARTCUST2.CUST_STORE_NO (+)"
            Create_TDA(.Tables.Add, "SOTSCOMO", "**", 0, True, "VV", 3)
            .Tables("SOTSCOMO").Columns.Add("SREP_COMM_CALC", GetType(System.Decimal), "ISNULL(INV_SALES,0) * ISNULL(SREP_COMM_PCT,0) / 100")
            .Tables("SOTSCOMO").Columns.Add("SREP_COMM", GetType(System.Decimal), "ISNULL(SREP_COMM_CALC,0) + ISNULL(SREP_COMM_ADJ,0)")

            ASCMAIN1.sql = "Select RSTRETLC.*, ARTCUST2.CUST_STORE_NAME, ARTCUST2.CUST_STORE_STATE" _
                & " from RSTRETLC, ARTCUST2" _
                & " where RSTRETLC.CUST_CODE = :PARM1 AND RSTRETLC.OPS_YYYYPP = :PARM2" _
                & " and RSTRETLC.CUST_CODE = ARTCUST2.CUST_CODE (+) AND RSTRETLC.CUST_STORE_NO = ARTCUST2.CUST_STORE_NO (+)"
            Create_TDA(.Tables.Add, "RSTRETLC", "**", 0, True, "VV", 4)
            .Tables("RSTRETLC").Columns.Add("SELL_COMM_CALC", GetType(System.Decimal), "ISNULL(AMT_SOLD,0) * ISNULL(SELL_COMM_PCT,0) / 100")
            .Tables("RSTRETLC").Columns.Add("SELL_COMM", GetType(System.Decimal), "ISNULL(SELL_COMM_CALC,0) + ISNULL(SELL_COMM_ADJ,0)")

            ASCMAIN1.sql = "Select SOTINVH2.*,ICTITEM1.ITEM_DESC" _
                & " from SOTINVH2,ICTITEM1 where ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" _
                & " and SOTINVH2.INV_TYPE = :PARM1 and SOTINVH2.INV_NO = :PARM2"
            Create_TDA(.Tables.Add, "SOTINVH2", "**", 0, True, "VV", 3)
            .Tables("SOTINVH2").Columns.Add("EXT_NET", GetType(System.Decimal), "ISNULL(ORDR_QTY_SHIP,0) * ISNULL(ORDR_UNIT_PRICE,0)")

            ASCMAIN1.sql = "Select X.*" & vbCrLf _
                & ",ARTCUST1.CUST_NAME,ARTCUST4.SREP_CODE,ARTCUST4.SREP_CODE_OVER,ARTCUST4.SREP_COMM_PCT_OVER" & vbCrLf _
                & " from ARTCUST1,ARTCUST4,(" & vbCrLf _
                & " SELECT CUST_CODE, SUM(RECS) RECS, SUM(SALES) SALES, SUM(COMMS) COMMS, SUM(COMMS_ADJ) COMMS_ADJ " & vbCrLf _
                & ", SUM(INVS) INVS, SUM(RTNS) RTNS, SUM(RTLS) RTLS" & vbCrLf _
                & " FROM (" & vbCrLf _
                & " Select CUST_CODE, COUNT (*) RECS" & vbCrLf _
                & ", SUM (INV_SALES) SALES" & vbCrLf _
                & ", SUM (INV_SALES * SREP_COMM_PCT / 100) COMMS" & vbCrLf _
                & ", SUM (SREP_COMM_ADJ) COMMS_ADJ" & vbCrLf _
                & ", SUM (DECODE(INV_TYPE,'I',1,0)) INVS" & vbCrLf _
                & ", SUM (DECODE(INV_TYPE,'C',1,0)) RTNS" & vbCrLf _
                & ", 0 RTLS" & vbCrLf _
                & " from SOTINVH1 where ORDR_YYYYPP_UPDATED = :PARM1" & vbCrLf _
                & " group by CUST_CODE" & vbCrLf _
                & " Union " & vbCrLf _
                & " Select CUST_CODE, COUNT (*) RECS" & vbCrLf _
                & ", SUM (INV_SALES) SALES" & vbCrLf _
                & ", SUM (INV_SALES * SREP_COMM_PCT / 100) COMMS" & vbCrLf _
                & ", SUM (SREP_COMM_ADJ) COMMS_ADJ" & vbCrLf _
                & ", SUM (DECODE(INV_TYPE,'I',1,0)) INVS" & vbCrLf _
                & ", SUM (DECODE(INV_TYPE,'C',1,0)) RTNS" & vbCrLf _
                & ", 0 RTLS" & vbCrLf _
                & " from SOTSCOMO where OPS_YYYYPP = :PARM1" & vbCrLf _
                & " group by CUST_CODE" & vbCrLf _
                & " Union " & vbCrLf _
                & " Select CUST_CODE, COUNT(*) RECS" & vbCrLf _
                & ", 0 SALES" & vbCrLf _
                & ", SUM (AMT_SOLD * SELL_COMM_PCT / 100) COMMS" & vbCrLf _
                & ", SUM (SELL_COMM_ADJ) COMMS_ADJ" & vbCrLf _
                & ", 0 INVS" & vbCrLf _
                & ", 0 RTNS" & vbCrLf _
                & ", COUNT(*) RTLS" & vbCrLf _
                & " from RSTRETLC where OPS_YYYYPP = :PARM1" & vbCrLf _
                & " group by CUST_CODE" & vbCrLf _
                & " ) " & vbCrLf _
                & " group by CUST_CODE ) X " & vbCrLf _
                & " where ARTCUST1.CUST_CODE (+) = X.CUST_CODE" & vbCrLf _
                & "   and ARTCUST4.OPS_YYYYPP (+) = :PARM1" & vbCrLf _
                & "   and ARTCUST4.CUST_CODE (+) = X.CUST_CODE"
            Create_TDA(.Tables.Add, "SOTSCOMC", "**", 0, False, "V", 1)
            .Tables("SOTSCOMC").Columns("INVS").DataType = GetType(System.Int64)
            .Tables("SOTSCOMC").Columns("RTNS").DataType = GetType(System.Int64)
            .Tables("SOTSCOMC").Columns("RTLS").DataType = GetType(System.Int64)
            .Tables("SOTSCOMC").Columns.Add("COMMS_OVER", GetType(System.Decimal), "SALES * SREP_COMM_PCT_OVER / 100")
            .Tables("SOTSCOMC").Columns.Add("COMMS_TOTAL", GetType(System.Decimal), "ISNULL(COMMS,0) + ISNULL(COMMS_ADJ,0) + ISNULL(COMMS_OVER,0)")

            ASCMAIN1.sql = "Select X.*,SOTSREP1.SREP_NAME,SOTSREP1.VEND_CODE FROM SOTSREP1," & vbCrLf _
                & "(Select SREP_CODE, SUM (RECS) RECS, SUM (SALES) SALES, SUM (COMMS) COMMS, SUM (COMMS_ADJ) COMMS_ADJ" & vbCrLf _
                & ", SUM (INVS) INVS, SUM (RTNS) RTNS, SUM(RTLS) RTLS" & vbCrLf _
                & ", SUM (COMMS_OVER) COMMS_OVER" & vbCrLf _
                & ", SUM (COMMS_MISC) COMMS_MISC" & vbCrLf _
                & " from (" & vbCrLf _
                & "Select SREP_CODE, COUNT (*) RECS" & vbCrLf _
                & ", SUM (INV_SALES) SALES" & vbCrLf _
                & ", SUM (INV_SALES * SREP_COMM_PCT / 100) COMMS" & vbCrLf _
                & ", SUM (SREP_COMM_ADJ) COMMS_ADJ" & vbCrLf _
                & ", SUM (DECODE(INV_TYPE,'I',1,0)) INVS" & vbCrLf _
                & ", SUM (DECODE(INV_TYPE,'C',1,0)) RTNS" & vbCrLf _
                & ", 0 RTLS, 0 COMMS_OVER" & vbCrLf _
                & ", 0 COMMS_MISC" & vbCrLf _
                & " from SOTINVH1 where ORDR_YYYYPP_UPDATED = :PARM1" & vbCrLf _
                & " group by SREP_CODE" & vbCrLf _
                & " union " & vbCrLf _
                & "Select SREP_CODE, COUNT (*) RECS" & vbCrLf _
                & ", SUM (INV_SALES) SALES" & vbCrLf _
                & ", SUM (INV_SALES * SREP_COMM_PCT / 100) COMMS" & vbCrLf _
                & ", SUM (SREP_COMM_ADJ) COMMS_ADJ" & vbCrLf _
                & ", SUM (DECODE(INV_TYPE,'I',1,0)) INVS" & vbCrLf _
                & ", SUM (DECODE(INV_TYPE,'C',1,0)) RTNS" & vbCrLf _
                & ", 0 RTLS, 0 COMMS_OVER" & vbCrLf _
                & ", 0 COMMS_MISC" & vbCrLf _
                & " from SOTSCOMO where OPS_YYYYPP = :PARM1" & vbCrLf _
                & " group by SREP_CODE" & vbCrLf _
                & " Union " & vbCrLf _
                & "Select ARTCUST4.SREP_CODE_OVER SREP_CODE, 0 RECS" & vbCrLf _
                & ", 0 SALES" & vbCrLf _
                & ", 0 COMMS" & vbCrLf _
                & ", 0 COMMS_ADJ" & vbCrLf _
                & ", 0 INVS" & vbCrLf _
                & ", 0 RTNS" & vbCrLf _
                & ", 0 RTLS" & vbCrLf _
                & ", SUM (SOTINVH1.INV_SALES * ARTCUST4.SREP_COMM_PCT_OVER / 100) COMMS_OVER" & vbCrLf _
                & ", 0 COMMS_MISC" & vbCrLf _
                & " from SOTINVH1,ARTCUST4 where SOTINVH1.ORDR_YYYYPP_UPDATED = :PARM1" & vbCrLf _
                & "   and ARTCUST4.OPS_YYYYPP = :PARM1" & vbCrLf _
                & "   and ARTCUST4.CUST_CODE = SOTINVH1.CUST_CODE" & vbCrLf _
                & "   and ARTCUST4.SREP_CODE_OVER is not null" & vbCrLf _
                & " group by ARTCUST4.SREP_CODE_OVER" _
                & " union " & vbCrLf _
                & "Select RSTRETLC.SELL_CODE, 0 RECS" & vbCrLf _
                & ", SUM (AMT_SOLD) SALES" & vbCrLf _
                & ", SUM (AMT_SOLD * SELL_COMM_PCT / 100) COMMS" & vbCrLf _
                & ", SUM (SELL_COMM_ADJ) COMMS_ADJ" & vbCrLf _
                & ", 0 INVS" & vbCrLf _
                & ", 0 RTNS" & vbCrLf _
                & ", COUNT(*) RTLS" & vbCrLf _
                & ", 0 COMMS_OVER" & vbCrLf _
                & ", 0 COMMS_MISC" & vbCrLf _
                & " from RSTRETLC where RSTRETLC.OPS_YYYYPP = :PARM1" & vbCrLf _
                & " group by RSTRETLC.SELL_CODE" _
                & " union " & vbCrLf _
                & "Select SOTSCOM2.SREP_CODE, 0 RECS" & vbCrLf _
                & ", 0 SALES" & vbCrLf _
                & ", 0 COMMS" & vbCrLf _
                & ", 0 COMMS_ADJ" & vbCrLf _
                & ", 0 INVS" & vbCrLf _
                & ", 0 RTNS" & vbCrLf _
                & ", 0 RTLS, 0 COMMS_OVER" & vbCrLf _
                & ", SUM (SREP_COMM_ADJ_MISC) COMMS_MISC" & vbCrLf _
                & " from SOTSCOM2 where NVL(SOTSCOM2.OPS_YYYYPP,'" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1) & "') = :PARM1" & vbCrLf _
                & " group by SOTSCOM2.SREP_CODE" _
                & ") group by SREP_CODE) X" & vbCrLf _
                & " where SOTSREP1.SREP_CODE (+) = X.SREP_CODE"
            Create_TDA(.Tables.Add, "SOTSCOMS", "**", 0, False, "V", 1)
            .Tables("SOTSCOMS").Columns("INVS").DataType = GetType(System.Int64)
            .Tables("SOTSCOMS").Columns("RTNS").DataType = GetType(System.Int64)
            .Tables("SOTSCOMS").Columns.Add("COMMS_TOTAL", GetType(System.Decimal), "ISNULL(COMMS,0) + ISNULL(COMMS_ADJ,0) + ISNULL(COMMS_OVER,0) + ISNULL(COMMS_MISC,0)")

            ASCMAIN1.sql = "Select * from SOTSREP1"
            Create_TDA(.Tables.Add, "SOTSREP1", "**", 0, False)

            .Tables.Add("TOTALS")
            .Tables("TOTALS").Columns.Add("CODE", GetType(System.String))
            .Tables("TOTALS").Columns.Add("DESC", GetType(System.String))
            .Tables("TOTALS").Columns.Add("AMOUNT", GetType(System.Decimal))

            ASCMAIN1.sql = "SELECT '0' SEL, SOTSCOM1.* FROM SOTSCOM1 WHERE OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "SOTSCOM1", "**", 0, False, "V", 0)

        End With

        Fill_Records("SOTSREP1")

        'grdSOTSCOMS.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
        'grdSOTSCOMC.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
        grdSOTSCOM2.DataSource = dst.Tables("SOTSCOM2")
        grdSOTSCOMS.DataSource = dst.Tables("SOTSCOMS")
        grdSOTSCOMC.DataSource = dst.Tables("SOTSCOMC")
        grdSOTINVH1.DataSource = dst.Tables("SOTINVH1")
        grdSOTINVH2.DataSource = dst.Tables("SOTINVH2")
        grdRSTRETLC.DataSource = dst.Tables("RSTRETLC")
        grdSOTSCOMO.DataSource = dst.Tables("SOTSCOMO")
        grdSOTSCOM1.DataSource = dst.Tables("SOTSCOM1")
        grdTotals.DataSource = dst.Tables("TOTALS")

        Create_Summary(grdTotals, "AMOUNT", "Sum")

        Create_Summary(grdSOTSCOMC, "CUST_CODE", "Count")
        Create_Summary(grdSOTSCOMC, New String() {"INVS", "RTNS", "SALES", "COMMS", "RTLS", "COMMS_ADJ", "COMMS_OVER", "COMMS_TOTAL"})

        Create_Summary(grdSOTSCOMS, "SREP_CODE", "Count")
        Create_Summary(grdSOTSCOMS, New String() {"INVS", "RTNS", "SALES", "COMMS", "RTLS", "COMMS_ADJ", "COMMS_OVER", "COMMS_MISC", "COMMS_TOTAL"})

        Create_Summary(grdSOTINVH1, "INV_NO", "Count")
        Create_Summary(grdSOTINVH1, New String() {"INV_SALES", "SREP_COMM_CALC", "SREP_COMM_ADJ", "SREP_COMM"})
        With grdSOTINVH1.DisplayLayout.Bands(0)
            .Columns("INV_TYPE").Header.Fixed = True
            .Columns("INV_NO").Header.Fixed = True

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                Dim COLUMN_NAME As String = gcol.Key
                If New String() {"SREP_CODE", "SREP2_CODE", "SREP_COMM_ADJ", "SREP_COMM_PCT", "SREP2_COMM_PCT"}.Contains(gcol.Key) Then
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                ElseIf COLUMN_NAME = "SREP_COMM_XNO" Then
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.Disabled
                Else
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        Create_Summary(grdSOTINVH2, "INV_LNO", "Count")
        Create_Summary(grdSOTINVH2, New String() {"ORDR_QTY_SHIP", "EXT_NET"})

        Create_Summary(grdRSTRETLC, "CUST_CODE", "Count")
        Create_Summary(grdRSTRETLC, New String() {"AMT_SOLD", "SELL_COMM_CALC", "SELL_COMM_ADJ", "SELL_COMM"})
        grdRSTRETLC.DisplayLayout.UseFixedHeaders = True
        With grdRSTRETLC.DisplayLayout.Bands(0)
            .Columns("CUST_CODE").Header.Fixed = True
            .Columns("CUST_STORE_NO").Header.Fixed = True

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                Dim COLUMN_NAME As String = gcol.Key
                If New String() {"SELL_COMM_PCT", "SELL_CODE", "SELL_COMM_ADJ"}.Contains(gcol.Key) Then
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            Next
        End With

        Create_Summary(grdSOTSCOMO, "CUST_CODE", "Count")
        Create_Summary(grdSOTSCOMO, New String() {"INV_SALES", "SREP_COMM_CALC", "SREP_COMM_ADJ", "SREP_COMM"})
        grdSOTSCOMO.DisplayLayout.UseFixedHeaders = True
        With grdSOTSCOMO.DisplayLayout.Bands(0)
            .Columns("CUST_CODE").Header.Fixed = True
            .Columns("CUST_STORE_NO").Header.Fixed = True

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                Dim COLUMN_NAME As String = gcol.Key
                If New String() {"SREP_COMM_PCT", "SREP_CODE", "SREP_COMM_ADJ"}.Contains(gcol.Key) Then
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            Next
        End With

        grdSOTSCOM2.DisplayLayout.UseFixedHeaders = True
        With grdSOTSCOM2.DisplayLayout.Bands(0)
            .Columns("SREP_COMM_ADJ_NO").Header.Fixed = True
            .Columns("SREP_COMM_ADJ_MISC").Header.Fixed = True
            .Columns("SREP_COMM_ADJ_NOTE").Header.Fixed = True
        End With

        grdSOTINVH1.DisplayLayout.UseFixedHeaders = True
        With grdSOTINVH1.DisplayLayout.Bands(0)
            .Columns("INV_TYPE").Header.Fixed = True
            .Columns("INV_NO").Header.Fixed = True

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                Dim COLUMN_NAME As String = gcol.Key
                '              .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                '              .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                If New String() {"SREP_COMM_PCT", "SREP_CODE", "INV_COMMENT", "SREP_COMM_ADJ"}.Contains(gcol.Key) Then
                    '                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    '                    .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightYellow
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                    '               Else
                    '                   .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        With grdSOTINVH2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                Dim COLUMN_NAME As String = gcol.Key
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"EXT_NET"}.Contains(gcol.Key) Then
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        ASCMAIN1.Add_Value_List(grdSOTSCOM1, "SREP_CODE", "SELECT SREP_CODE, '(' || SREP_CODE || ') ' || INITCAP(SREP_NAME) FROM SOTSREP1")
        ASCMAIN1.Add_Value_List(grdSOTSCOM1, "SREP_CODE_MGR", "SELECT SREP_CODE, '(' || SREP_CODE || ') ' || INITCAP(SREP_NAME) FROM SOTSREP1")
        grdSOTSCOM1.DisplayLayout.UseFixedHeaders = True
        With grdSOTSCOM1.DisplayLayout.Bands(0)
            .Columns("SEL").Header.Fixed = True
            .Columns("SREP_CODE").Header.Fixed = True
            .Columns("SREP_CODE_MGR").Header.Fixed = True
        End With
        Create_Summary(grdSOTSCOM1, "SEL", "Sum")
        Create_Summary(grdSOTSCOM1, "SREP_CODE", "Count")
        Create_Summary(grdSOTSCOM1, "SREP_COMM_CALC", "Sum")
        Create_Summary(grdSOTSCOM1, "SREP_COMM_ADJ", "Sum")
        Create_Summary(grdSOTSCOM1, "SREP_COMM_ADJ_MISC", "Sum")
        Create_Summary(grdSOTSCOM1, "SREP_COMM_OVER", "Sum")
        Create_Summary(grdSOTSCOM1, "SREP_COMM_TOTAL", "Sum")

        grdSOTSCOMC.DisplayLayout.UseFixedHeaders = True
        With grdSOTSCOMC.DisplayLayout.Bands("SOTSCOMC")
            .Columns("CUST_CODE").Header.Fixed = True
            .Columns("CUST_NAME").Header.Fixed = True
        End With

        grdSOTSCOMS.DisplayLayout.UseFixedHeaders = True
        With grdSOTSCOMS.DisplayLayout.Bands("SOTSCOMS")
            .Columns("SREP_CODE").Header.Fixed = True
            .Columns("SREP_NAME").Header.Fixed = True
        End With

        ' Oldest period should be the previous period. Commissions are done after month end
        ASCMAIN1.sql = "Select OPS_YYYYPP, LEGEND from GLTPARM2" _
            & " where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -13) & "' and OPS_YYYYPP <= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1) & "'"
        cbeOPS_YYYYPP.DataSource = ASCDATA1.GetDataTable

        ASCMAIN1.sql = "Select SREP_CODE, SREP_NAME from SOTSREP1"
        Dim DT As DataTable = ASCDATA1.GetDataTable
        Dim DVW As DataView = DT.DefaultView
        DVW.Sort = "SREP_CODE"
        cmbSREP_CODE.DataSource = DVW.ToTable
        ' ASCMAIN1.Add_Value_List(grdICTCOSTS, "COST_TYPE", , New String() {":", "1:Direct", "2:Materials"})

        LYP = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)
        Position_Controls(grdSOTINVH1)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Edit", "View"
                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    Validate_Code("CUST_CODE")
                    If EMsg = "" Then
                        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                        rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                        If rowARTCUST1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Customer Code " & CUST_CODE
                        Else
                            SREP_CODE = rowARTCUST1.Item("SREP_CODE")
                            Absx1.txtFor("SREP_CODE").Text = SREP_CODE
                        End If
                    End If
                ElseIf Absx1.txtFor("SREP_CODE").Text <> "" Then
                    Validate_Code("SREP_CODE")
                    If EMsg = "" Then
                        CUST_CODE = ""
                        SREP_CODE = Absx1.txtFor("SREP_CODE").Text
                        rowSOTSREP1 = LookUp("SOTSREP1", SREP_CODE)
                        If rowSOTSREP1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Sales Rep Code " & SREP_CODE
                        End If
                    End If
                Else
                    EMsg &= vbCr & "You Must Enter either a Customer Code or a Sales Rep Code"
                End If

                If EMsg = "" And eItemKey = "Edit" Then
                    If Not ASCMAIN1.Logical_Lock("SOTSCOM1", OPS_YYYYPP) Then
                        Exit Sub
                    End If
                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo, _
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Update", "Save"

                If dst.Tables("SOTINVH1").Select("SREP_COMM_IND = '1'", "", DataViewRowState.ModifiedCurrent).Length <> 0 Then
                    If eItemKey = "Save" Then
                        EMsg = "You are not permitted to Save when there are Commission Adjustments"
                        Exit Select
                    End If

                    If MsgBox("OK to Continue with Update?", MsgBoxStyle.YesNo, _
                              "This Update will create Commission Adjustment Records") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Import Retail Sales"
                Dim sql As String = String.Empty
                sql = "select distinct SELL_COMM_XNO from RSTRETLC Where SELL_COMM_XNO is not null and OPS_YYYYPP = '" & OPS_YYYYPP & "'"
                sql &= " Union"
                sql &= " select distinct SREP_COMM_XNO from SOTSCOMO Where SREP_COMM_XNO is not null and OPS_YYYYPP = '" & OPS_YYYYPP & "'"
                sql &= " Union"
                sql &= " select distinct SREP_COMM_XNO from SOTSCOM1 Where VOUCHER_NO is not null and OPS_YYYYPP = '" & OPS_YYYYPP & "'"
                If ASCDATA1.GetDataTable(sql).Rows.Count > 0 Then
                    EMsg &= vbCr & "The selected period has been finalized you cannot Import / Recalculate Retails Sales Commissions."
                Else
                    If MessageBox.Show("Do you want to Import / Recalculate Retails Sales Commissions? All existing Retails Sales Commissions will be deleted." _
                                        & Environment.NewLine & Environment.NewLine _
                                        & "Continue?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
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

            Case "Edit", "View"
                If eItemKey = "View" Then
                    EntryMode = "V"
                Else
                    EntryMode = "E"
                End If
                Load_Record()
                Mode_Settings(True)

            Case "Update", "Save"
                Update_Record()
                If eItemKey = "Update" Then
                    Mode_Settings(False)
                End If

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "Import Retail Sales"
                SetMonthlyCommissions(OPS_YYYYPP)
                Mode_Settings(False)
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

                    .Items("Import Retail Sales").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Save").Settings.Enabled = iScreenMode

                    .Items("Import Retail Sales").Visible = (EntryMode = "N")
                    .Items("Edit").Visible = (Not InquiryMode)
                    .Items("Done").Visible = (EntryMode = "N" Or EntryMode <> "E")
                    .Items("Print").Visible = False ' (InquiryMode Or EntryMode = "V")
                    .Items("Update").Visible = (Not InquiryMode And EntryMode <> "V")
                    .Items("Cancel").Visible = (Not InquiryMode And EntryMode <> "V")
                    .Items("Save").Visible = (Not InquiryMode And EntryMode <> "V") AndAlso _
                        dst.Tables("SOTINVH1").Select("SREP_COMM_IND = '1'", "", DataViewRowState.ModifiedCurrent).Length = 0
                End With
                .Groups("Totals").Visible = tf
            End With
        End If

        lblCUST_CODE.Visible = Not ScreenMode Or (CUST_CODE <> "")
        txtCUST_CODE.Visible = Not ScreenMode Or (CUST_CODE <> "")
        txtCUST_NAME.Visible = Not ScreenMode Or (CUST_CODE <> "")

        If ScreenMode Then
            If InquiryMode Or (EntryMode = "V") Then
                grdSOTINVH1.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdSOTINVH1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdSOTINVH1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

                grdRSTRETLC.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdRSTRETLC.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdRSTRETLC.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

                grdSOTSCOMO.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdSOTSCOMO.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdSOTSCOMO.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            Else
                grdSOTINVH1.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdSOTINVH1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdSOTINVH1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

                grdRSTRETLC.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdRSTRETLC.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdRSTRETLC.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

                grdSOTSCOMO.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                grdSOTSCOMO.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdSOTSCOMO.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            End If

            With grdSOTINVH1.DisplayLayout.Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    Dim COLUMN_NAME As String = gcol.Key
                    .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                    If EntryMode = "E" And New String() {"SREP_COMM_PCT", "SREP_CODE", "INV_COMMENT", "SREP_COMM_ADJ"}.Contains(gcol.Key) Then
                        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                        .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightYellow
                    Else
                        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                        .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Empty
                    End If
                Next
            End With

            With grdRSTRETLC.DisplayLayout.Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    Dim COLUMN_NAME As String = gcol.Key
                    .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                    If EntryMode = "E" And New String() {"SELL_COMM_PCT", "SELL_CODE", "SELL_COMM_ADJ"}.Contains(gcol.Key) Then
                        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                        .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightYellow
                    Else
                        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                        .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Empty
                    End If
                Next
            End With

            With grdSOTSCOMO.DisplayLayout.Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    Dim COLUMN_NAME As String = gcol.Key
                    .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                    If EntryMode = "E" And New String() {"SREP_COMM_PCT", "SREP_CODE", "SREP_COMM_ADJ"}.Contains(gcol.Key) Then
                        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                        .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightYellow
                    Else
                        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                        .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Empty
                    End If
                Next
            End With
        End If

        Setup_tabSummary()
        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(cbeOPS_YYYYPP, ScreenMode)

        cmbSREP_CODE.Visible = (EntryMode = "E")
        numSREP_COMM_PCT.Visible = (EntryMode = "E")

        tabSummary.Visible = Not ScreenMode
        tabInvoices.Visible = ScreenMode

        If tabInvoices.SelectedTab.Key = tabInvoices.Tabs("Invoices").Key Then
            Position_Controls(grdSOTINVH1)
        ElseIf tabInvoices.SelectedTab.Key = tabInvoices.Tabs("Retail Sales").Key Then
            Position_Controls(grdRSTRETLC)
        Else
            Position_Controls(grdSOTSCOMO)
        End If

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTINVH1", "SOTINVH2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        If cbeOPS_YYYYPP.Value & "" = "" Then
            cbeOPS_YYYYPP.Value = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)
        Else
            Load_SOTSCOMC()
        End If


        EnforceConstraints(True)

        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("SREP_CODE").Text = ""
        EntryMode = ""
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        If CUST_CODE = "" Then
            'ASCMAIN1.sql = "Select SOTINVH1.*" _
            '  & " from SOTINVH1 where SREP_CODE = '" & SREP_CODE & "' and ORDR_YYYYPP_UPDATED = '" & OPS_YYYYPP & "'"

            ASCMAIN1.sql = "Select SOTINVH1.*, ARTCUST2.CUST_STORE_STATE" _
                 & " from SOTINVH1, ARTCUST2" _
                 & " where SOTINVH1.SREP_CODE = '" & SREP_CODE & "' and SOTINVH1.ORDR_YYYYPP_UPDATED = '" & OPS_YYYYPP & "'" _
                 & " and SOTINVH1.CUST_CODE = ARTCUST2.CUST_CODE (+)  AND SOTINVH1.CUST_STORE_NO = ARTCUST2.CUST_STORE_NO (+)"

            Fill_Records("SOTINVH1", "", , ASCMAIN1.sql)
            grdSOTINVH1.Text = "Sales Invoices for Sales Rep " & SREP_CODE & " posted in " & cbeOPS_YYYYPP.Text

            ASCMAIN1.sql = "Select RSTRETLC.*, ARTCUST2.CUST_STORE_NAME, ARTCUST2.CUST_STORE_STATE" _
                & " from RSTRETLC, ARTCUST2" _
                & " where RSTRETLC.SELL_CODE = '" & SREP_CODE & "' and RSTRETLC.OPS_YYYYPP = '" & OPS_YYYYPP & "'" _
                & " and RSTRETLC.CUST_CODE = ARTCUST2.CUST_CODE (+) AND RSTRETLC.CUST_STORE_NO = ARTCUST2.CUST_STORE_NO (+)"
            Fill_Records("RSTRETLC", "", , ASCMAIN1.sql)
            grdRSTRETLC.Text = "Retail Sales for Sales Rep " & SREP_CODE & " posted in " & cbeOPS_YYYYPP.Text

            ASCMAIN1.sql = "Select SOTSCOMO.*, ARTCUST2.CUST_STORE_NAME, ARTCUST2.CUST_STORE_STATE" _
                & " from SOTSCOMO, ARTCUST2" _
                & " where SOTSCOMO.SREP_CODE = '" & SREP_CODE & "' and SOTSCOMO.OPS_YYYYPP = '" & OPS_YYYYPP & "'" _
                & " and SOTSCOMO.CUST_CODE = ARTCUST2.CUST_CODE (+) AND SOTSCOMO.CUST_STORE_NO = ARTCUST2.CUST_STORE_NO (+)"
            Fill_Records("SOTSCOMO", "", , ASCMAIN1.sql)
            grdSOTSCOMO.Text = "Override Sales for Sales Rep " & SREP_CODE & " posted in " & cbeOPS_YYYYPP.Text

        Else
            Fill_Records("SOTINVH1", New String() {CUST_CODE, OPS_YYYYPP})
            grdSOTINVH1.Text = "Sales Invoices for Customer " & CUST_CODE & " posted in " & cbeOPS_YYYYPP.Text

            Fill_Records("RSTRETLC", New String() {CUST_CODE, OPS_YYYYPP})
            grdRSTRETLC.Text = "Retail Sales for Customer " & CUST_CODE & " posted in " & cbeOPS_YYYYPP.Text

            Fill_Records("SOTSCOMO", New String() {CUST_CODE, OPS_YYYYPP})
            grdSOTSCOMO.Text = "Override Sales for Customer " & CUST_CODE & " posted in " & cbeOPS_YYYYPP.Text
        End If

        Sort_grdColumns(grdSOTINVH1, "INV_NO")
        Sort_grdColumns(grdRSTRETLC, "CUST_CODE")
        Sort_grdColumns(grdSOTSCOMO, "CUST_CODE")

        EnforceConstraints(True)

        cmbSREP_CODE.Value = SREP_CODE
        Dim SREP_COMM_PCT As Decimal = 0
        Dim rowARTCUST4 As DataRow
        If CUST_CODE <> "" Then
            rowARTCUST4 = LookUp("ARTCUST4", New String() {OPS_YYYYPP, CUST_CODE})
        Else
            ASCMAIN1.sql = "Select * from ARTCUST4 where OPS_YYYYPP = '" & OPS_YYYYPP & "' and SREP_CODE = '" & SREP_CODE & "'"
            rowARTCUST4 = ASCDATA1.GetDataRow
        End If

        If rowARTCUST4 IsNot Nothing Then
            SREP_COMM_PCT = Val(rowARTCUST4.Item("SREP_COMM_PCT") & "")
        End If

        numSREP_COMM_PCT.Value = SREP_COMM_PCT
        numSREP_COMM_PCT.Focus()

        SETUP_grdSOTINVH2()
        GetTotals()

        'Display_Totals()
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        ' do we write this row if the old comm amt was 0
        Try
            For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("SREP_COMM_IND = '1'", "", DataViewRowState.ModifiedCurrent)
                Dim rowSOTSCOM2 As DataRow = dst.Tables("SOTSCOM2").NewRow
                With rowSOTSCOM2
                    .Item("SREP_COMM_ADJ_NO") = ASCMAIN1.Next_Control_No("SOTSCOM2.SREP_COMM_ADJ_NO")
                    .Item("SREP_CODE") = rowSOTINVH1.Item("SREP_CODE", DataRowVersion.Original)
                    .Item("SREP_COMM_ADJ_MISC") = -1 * (Val(rowSOTINVH1.Item("SREP_COMM_CALC", DataRowVersion.Original) & "") + Val(rowSOTINVH1.Item("SREP_COMM_ADJ", DataRowVersion.Original) & ""))
                    .Item("SREP_COMM_ADJ_NOTE") = "Prior Period Adjustment"
                    .Item("INV_TYPE") = rowSOTINVH1.Item("INV_TYPE", DataRowVersion.Original)
                    .Item("INV_NO") = rowSOTINVH1.Item("INV_NO", DataRowVersion.Original)
                    .Item("INV_SALES") = rowSOTINVH1.Item("INV_SALES", DataRowVersion.Original)
                    .Item("INV_COMMENT") = rowSOTINVH1.Item("INV_COMMENT", DataRowVersion.Original)
                    .Item("SREP_COMM_PCT") = rowSOTINVH1.Item("SREP_COMM_PCT", DataRowVersion.Original)
                    .Item("SREP_COMM_CALC") = rowSOTINVH1.Item("SREP_COMM_CALC", DataRowVersion.Original)
                    .Item("SREP_COMM_ADJ") = rowSOTINVH1.Item("SREP_COMM_ADJ", DataRowVersion.Original)
                    .Item("SREP_COMM_XNO_ADJ") = rowSOTINVH1.Item("SREP_COMM_XNO")
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("OPS_YYYYPP") = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)
                End With
                dst.Tables("SOTSCOM2").Rows.Add(rowSOTSCOM2)

                rowSOTINVH1.Item("SREP_COMM_XNO") = DBNull.Value
                rowSOTINVH1.Item("SREP_COMM_IND") = DBNull.Value
            Next

            Try
                BeginTrans()
                Update_Record_TDA("SOTINVH1")
                Update_Record_TDA("SOTSCOM2")
                Update_Record_TDA("RSTRETLC")
                Update_Record_TDA("SOTSCOMO")
                CommitTrans("Update Complete")
            Catch ex As Exception
                Rollback(ex.Message)
            End Try

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTSCOMC, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTINVH1, "SSSBBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "Change to SRep", "Change to Comm%")
        Load_Popup_Menu(grdSOTINVH2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Inventory Status")
        Load_Popup_Menu(grdRSTRETLC, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Change to SRep", "Change to Comm%")
        Load_Popup_Menu(grdSOTSCOMO, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Change to SRep", "Change to Comm%")
        Load_Popup_Menu(grdSOTSCOM1, "SSSBBB", "Show Filter", "Show GroupBox", "Show Pins", "Select All", "De-Select All", "Email Reports")
    End Sub

    Public Overrides Sub tlb_beforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdSOTINVH1", "grdRSTRETLC", "grdSOTSCOMO"
                    tlb_btn = DirectCast(tlb.Tools("Change to SRep"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Caption = "Change to SRep " & cmbSREP_CODE.Value
                    tlb_btn.SharedProps.Visible = (EntryMode = "E")
                    tlb_btn = DirectCast(tlb.Tools("Change to Comm%"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Caption = "Change to " & Format(numSREP_COMM_PCT.Value, "#.00") & "% Comm"
                    tlb_btn.SharedProps.Visible = (EntryMode = "E")
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim rowsSkipped As Int16 = 0

        Select Case e.Tool.Key

            Case "Change to SRep"
                Dim SREP_CODE As String = "SREP_CODE"
                Dim SREP_COMM_XNO As String = "SREP_COMM_XNO"

                If grd.Name = "grdRSTRETLC" Then
                    SREP_CODE = "SELL_CODE"
                    SREP_COMM_XNO = "SELL_COMM_XNO"
                End If

                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    If grow.Cells(SREP_COMM_XNO).Value & String.Empty <> String.Empty Then
                        rowsSkipped += 1
                        Continue For
                    End If
                    grow.Cells(SREP_CODE).Value = cmbSREP_CODE.Value
                    grow.Update()
                Next
                grd.Selected.Rows.Clear()

            Case "Change to Comm%"
                Dim SREP_COMM_PCT As String = "SREP_COMM_PCT"
                Dim SREP_COMM_XNO As String = "SREP_COMM_XNO"

                If grd.Name = "grdRSTRETLC" Then
                    SREP_COMM_PCT = "SELL_COMM_PCT"
                    SREP_COMM_XNO = "SELL_COMM_XNO"
                End If

                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    If grow.Cells(SREP_COMM_XNO).Value & String.Empty <> String.Empty Then
                        rowsSkipped += 1
                        Continue For
                    End If

                    grow.Cells(SREP_COMM_PCT).Value = numSREP_COMM_PCT.Value
                    grow.Update()
                Next
                grd.Selected.Rows.Clear()

            Case Else

        End Select

        If rowsSkipped > 0 Then
            MessageBox.Show(rowsSkipped & " of the selected row(s) are already paid; therefore, those rows were not updated.", "Change", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("Select", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Inventory Status"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("BM_COMP_ITEM").Value
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Select All"
                For Each row As DataRow In dst.Tables("SOTSCOM1").Select()
                    row.Item("SEL") = "1"
                Next
                dst.Tables("SOTSCOM1").AcceptChanges()

            Case "De-Select All"
                For Each row As DataRow In dst.Tables("SOTSCOM1").Select()
                    row.Item("SEL") = "0"
                Next
                dst.Tables("SOTSCOM1").AcceptChanges()

            Case "Email Reports"
                EmailCommissionReports()
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    LOAD_SOTSCOMC()
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                LOAD_SOTSCOMC()
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

            Case "CUST_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                        LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

        End Select
    End Sub

#End Region

#Region "grdRSTRETLC"

    Private Sub grdRSTRETLC_AfterColPosChanged(sender As Object, e As Infragistics.Win.UltraWinGrid.AfterColPosChangedEventArgs) Handles grdRSTRETLC.AfterColPosChanged
        Position_Controls(grdRSTRETLC)
    End Sub

    Private Sub grdRSTRETLC_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdRSTRETLC.AfterRowUpdate
        GetTotals()
    End Sub

    Private Sub grdRSTRETLC_SizeChanged(sender As Object, e As System.EventArgs) Handles grdRSTRETLC.SizeChanged
        Position_Controls(grdRSTRETLC)
    End Sub

    Private Sub grdRSTRETLC_AfterColRegionScroll(sender As Object, e As Infragistics.Win.UltraWinGrid.ColScrollRegionEventArgs) Handles grdRSTRETLC.AfterColRegionScroll
        Position_Controls(grdRSTRETLC)
    End Sub

    Private Sub grdRSTRETLC_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdRSTRETLC.BeforeRowUpdate
        Dim SREP_CODE As String = e.Row.Cells("SELL_CODE").Value & ""
        If LookUp("SOTSREP1", SREP_CODE) Is Nothing Then
            MessageBox.Show("Invalid/Missing Sales Rep", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If
    End Sub

    Private Sub grdRSTRETLC_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdRSTRETLC.InitializeRow
        If e.Row.IsDataRow Then
            e.Row.Cells("CUST_CODE").Appearance.BackColor = Drawing.Color.Empty
            If e.Row.Cells("SELL_COMM_IND").Value & "" = "1" Then
                e.Row.Cells("CUST_CODE").Appearance.BackColor = Drawing.Color.Yellow
                e.Row.Cells("CUST_CODE").ToolTipText = "Commission on this Invoice has already been Posted"
            End If
        End If
    End Sub

    Private Sub grdRSTRETLC_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdRSTRETLC.AfterRowActivate
        '   No edits if already painf the commission for this line. User must make manual adjustment
        If EntryMode = "E" Then
            If grdRSTRETLC.ActiveRow.Cells("SELL_COMM_XNO").Value & String.Empty <> String.Empty Then
                grdRSTRETLC.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdRSTRETLC.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            Else
                grdRSTRETLC.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdRSTRETLC.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            End If
        End If
    End Sub

#End Region

#Region "grdSOTSCOMO"

    Private Sub grdSOTSCOMO_AfterColPosChanged(sender As Object, e As Infragistics.Win.UltraWinGrid.AfterColPosChangedEventArgs) Handles grdSOTSCOMO.AfterColPosChanged
        Position_Controls(grdSOTSCOMO)
    End Sub

    Private Sub grdSOTSCOMO_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTSCOMO.AfterRowUpdate
        GetTotals()
    End Sub

    Private Sub grdSOTSCOMO_SizeChanged(sender As Object, e As System.EventArgs) Handles grdSOTSCOMO.SizeChanged
        Position_Controls(grdSOTSCOMO)
    End Sub

    Private Sub grdSOTSCOMO_AfterColRegionScroll(sender As Object, e As Infragistics.Win.UltraWinGrid.ColScrollRegionEventArgs) Handles grdSOTSCOMO.AfterColRegionScroll
        Position_Controls(grdSOTSCOMO)
    End Sub

    Private Sub grdSOTSCOMO_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTSCOMO.BeforeRowUpdate

        ' Validate Customer Code
        e.Row.Cells("CUST_CODE").Value = (e.Row.Cells("CUST_CODE").Value & String.Empty).ToString.ToUpper
        Dim CUST_CODE As String = e.Row.Cells("CUST_CODE").Value

        e.Row.Cells("CUST_STORE_NO").Value = (e.Row.Cells("CUST_STORE_NO").Value & String.Empty).ToString.ToUpper
        Dim CUST_STORE_NO As String = e.Row.Cells("CUST_STORE_NO").Value

        If LookUp("ARTCUST1", CUST_CODE) Is Nothing Then
            MessageBox.Show("Invalid Customer")
            e.Cancel = True
            Exit Sub
        End If

        Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
        If rowARTCUST2 Is Nothing Then
            MessageBox.Show("Invalid Customer / Store combination")
            e.Cancel = True
            Exit Sub
        End If

        'ARTCUST2.CUST_STORE_NAME, ARTCUST2.CUST_STORE_STATE
        If e.Row.IsAddRow Then
            e.Row.Cells("CUST_STORE_NAME").Value = rowARTCUST2.Item("CUST_STORE_NAME") & ""
            e.Row.Cells("CUST_STORE_STATE").Value = rowARTCUST2.Item("CUST_STORE_STATE") & ""
        End If

        Dim SREP_COMM_PCT As Double = Val(e.Row.Cells("SREP_COMM_PCT").Value & String.Empty)
        e.Row.Cells("SREP_COMM_PCT").Value = SREP_COMM_PCT
        If SREP_COMM_PCT < 0 OrElse SREP_COMM_PCT > 100 Then
            MessageBox.Show("Commission Percentage must be between 0 and 100")
            e.Cancel = True
            Exit Sub
        End If

        Dim INV_SALES As Double = Val(e.Row.Cells("INV_SALES").Value & String.Empty)
        e.Row.Cells("INV_SALES").Value = INV_SALES

        Dim SREP_CODE As String = e.Row.Cells("SREP_CODE").Value & ""
        If LookUp("SOTSREP1", SREP_CODE) Is Nothing Then
            MessageBox.Show("Invalid/Missing Sales Rep", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        If e.Row.IsAddRow AndAlso e.Row.Cells("INV_NO").Value & String.Empty = String.Empty Then
            e.Row.Cells("INV_TYPE").Value = "O"
            e.Row.Cells("INV_NO").Value = ASCMAIN1.Next_Control_No("SOTSCOMO.INV_NO")
        End If

        If e.Row.Cells("OPS_YYYYPP").Value & String.Empty = String.Empty Then
            e.Row.Cells("OPS_YYYYPP").Value = OPS_YYYYPP
        End If

    End Sub

    Private Sub grdSOTSCOMO_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTSCOMO.InitializeRow
        If e.Row.IsDataRow Then
            e.Row.Cells("CUST_CODE").Appearance.BackColor = Drawing.Color.Empty
            If e.Row.Cells("SREP_COMM_IND").Value & "" = "1" Then
                e.Row.Cells("CUST_CODE").Appearance.BackColor = Drawing.Color.Yellow
                e.Row.Cells("CUST_CODE").ToolTipText = "Commission on this Invoice has already been Posted"
            End If
        End If
    End Sub

    Private Sub grdSOTSCOMO_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTSCOMO.AfterRowActivate

        '   No edits if already painf the commission for this line. User must make manual adjustment
        If EntryMode = "E" Then
            If grdSOTSCOMO.ActiveRow.Cells("SREP_COMM_XNO").Value & String.Empty <> String.Empty Then
                grdSOTSCOMO.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdSOTSCOMO.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            Else
                grdSOTSCOMO.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdSOTSCOMO.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            End If
        End If

        If grdSOTSCOMO.ActiveRow.IsAddRow Then
            If CUST_CODE.Length > 0 AndAlso grdSOTSCOMO.ActiveRow.Cells("CUST_CODE").Value & String.Empty = String.Empty Then
                grdSOTSCOMO.ActiveRow.Cells("CUST_CODE").Value = CUST_CODE
            ElseIf grdSOTSCOMO.ActiveRow.Cells("SREP_CODE").Value & String.Empty = String.Empty Then
                grdSOTSCOMO.ActiveRow.Cells("SREP_CODE").Value = SREP_CODE
            End If

            If grdSOTSCOMO.ActiveRow.Cells("INV_COMMENT").Value & String.Empty = String.Empty Then
                grdSOTSCOMO.ActiveRow.Cells("INV_COMMENT").Value = "Manual Entry"
            End If
        End If

        If grdSOTSCOMO.DisplayLayout.Override.AllowAddNew <> UltraWinGrid.AllowAddNew.No Then
            If grdSOTSCOMO.ActiveRow.IsAddRow Then
                If CUST_CODE.Length > 0 Then
                    grdSOTSCOMO.DisplayLayout.Bands(0).Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                    grdSOTSCOMO.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    grdSOTSCOMO.DisplayLayout.Bands(0).Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                    grdSOTSCOMO.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            Else
                grdSOTSCOMO.DisplayLayout.Bands(0).Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                grdSOTSCOMO.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End If

    End Sub

    Private Sub grdSOTSCOMO_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTSCOMO.ClickCellButton

        Select Case e.Cell.Column.Key
            Case "CUST_CODE"
                grdClickCellButton(grdSOTSCOMO, "")
            Case "CUST_STORE_NO"
                grdClickCellButton(grdSOTSCOMO, "CUST_CODE = '" & e.Cell.Row.Cells("CUST_CODE").Value & "'")
            Case "SREP_CODE"
                grdClickCellButton(grdSOTSCOMO, "")
        End Select
    End Sub

    Private Sub grdSOTSCOMC_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTSCOMC.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("CUST_CODE").Text = e.Row.Cells("CUST_CODE").Value
            Click_Command("View")
        End If
    End Sub

#End Region

#Region "grdSOTINVH1"

    Private Sub grdSOTINVH1_AfterColRegionScroll(sender As Object, e As Infragistics.Win.UltraWinGrid.ColScrollRegionEventArgs) Handles grdSOTINVH1.AfterColRegionScroll
        Position_Controls(grdSOTINVH1)
    End Sub

    Private Sub grdSOTINVH1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTINVH1.AfterRowUpdate
        GetTotals()
    End Sub

    Private Sub grdSOTINVH1_AfterColPosChanged(sender As Object, e As Infragistics.Win.UltraWinGrid.AfterColPosChangedEventArgs) Handles grdSOTINVH1.AfterColPosChanged
        Position_Controls(grdSOTINVH1)
    End Sub

    Private Sub grdSOTINVH1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTINVH1.AfterRowActivate
        SETUP_grdSOTINVH2()

        '   No edits if already painf the commission for this line. User must make manual adjustment
        If EntryMode = "E" Then
            If grdSOTINVH1.ActiveRow.Cells("SREP_COMM_XNO").Value & String.Empty <> String.Empty Then
                grdSOTINVH1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdSOTINVH1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            Else
                grdSOTINVH1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdSOTINVH1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            End If
        End If
    End Sub

    Private Sub grdSOTINVH1_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTINVH1.BeforeRowUpdate
        Dim SREP_CODE As String = e.Row.Cells("SREP_CODE").Value & ""
        If LookUp("SOTSREP1", SREP_CODE) Is Nothing Then
            MessageBox.Show("Invalid/Missing Sales Rep", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If
    End Sub

    Private Sub grdSOTINVH1_SizeChanged(sender As Object, e As System.EventArgs) Handles grdSOTINVH1.SizeChanged
        Position_Controls(grdSOTINVH1)
    End Sub

    Private Sub grdSOTINVH1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTINVH1.InitializeRow
        If e.Row.IsDataRow Then
            e.Row.Cells("INV_NO").Appearance.BackColor = Drawing.Color.Empty
            If e.Row.Cells("SREP_COMM_IND").Value & "" = "1" Then
                e.Row.Cells("INV_NO").Appearance.BackColor = Drawing.Color.Yellow
                e.Row.Cells("INV_NO").ToolTipText = "Commission on this Invoice has already been Posted"
            End If

            ' not easy - because we need to pack SREP_CODE_CUST and SREP_CODE_CUST_STORE into SOTINVH1 from Oracle, and this comes from ARTCUST4,ARTCUST7
            'e.Row.Cells("SREP_CODE").Appearance.ForeColor = Drawing.Color.Empty
            'Dim SREP_CODE_this As String = e.Row.Cells("SREP_CODE_CUST").Value & ""
            'If e.Row.Cells("SREP_CODE_CUST_STORE").Value & "" <> "" Then
            '    SREP_CODE_this = e.Row.Cells("SREP_CODE_CUST_STORE").Value
            'End If
            'If e.Row.Cells("SREP_CODE").Value & "" <> SREP_CODE_THIS Then
            '    e.Row.Cells("SREP_CODE").Appearance.BackColor = Drawing.Color.Red
            '    e.Row.Cells("SREP_CODE").ToolTipText = "Sales Rep for this Account/Store is Normally " & SREP_CODE_this
            'End If

            ' not easy - because we do not have sales rep commission % history, especially with account exceptions and overrides
            'e.Row.Cells("SREP_COMM_PCT").Appearance.ForeColor = Drawing.Color.Empty
            'Dim SREP_COMM_PCT_this As Decimal = 0
            'If e.Row.Cells("SREP_CODE").Value & "" <> SREP_COMM_PCT_this Then
            '    e.Row.Cells("SREP_COMM_PCT").Appearance.BackColor = Drawing.Color.Red
            '    e.Row.Cells("SREP_COMM_PCT").ToolTipText = "Sales Rep Commission for this Account is Normally " & Format(SREP_COMM_PCT_this, "##.0") & "%"
            'End If
        End If
    End Sub

#End Region

#Region "grdSOTSCOMS"

    Private Sub grdSOTSCOMS_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTSCOMS.AfterRowActivate
        Setup_grdSOTSCOM2()
    End Sub

    Private Sub grdSOTSCOMS_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTSCOMS.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("CUST_CODE").Text = ""
            Absx1.txtFor("SREP_CODE").Text = e.Row.Cells("SREP_CODE").Value & String.Empty
            Click_Command("View")
        End If
    End Sub

    Private Sub grdSOTSCOM2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTSCOM2.InitializeRow
        If e.Row.Cells("OPS_YYYYPP").Value & "" = "" Then
            e.Row.Cells("SREP_COMM_ADJ_NO").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("SREP_COMM_ADJ_NO").ToolTipText = "Not Posted Yet"
        End If
    End Sub

    Private Sub grdSOTSCOM2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTSCOM2.AfterRowsDeleted
        'Dim SREP_COMM_ADJ_NOs As List(Of String) = DirectCast(grdSOTSCOM2.Tag, List(Of String))
        'ASCMAIN1.sql = "Delete from SOTSCOM2 where SREP_COMM_ADJ_NO in ('" & Join(SREP_COMM_ADJ_NOs.ToArray, "','") & "')"
        ASCMAIN1.sql = CStr(grdSOTSCOM2.Tag)
        ASCDATA1.ExecuteSQL()
    End Sub

    Private Sub grdSOTSCOM2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTSCOM2.BeforeRowsDeleted
        'Dim SREP_COMM_ADJ_NOs As New List(Of String)
        Dim SREP_COMM_ADJ_NOs As String = ""
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If grow.Cells("OPS_YYYYPP").Value & "" <> "" Then
                e.Cancel = True
                MsgBox("Cannot Delete Rows which have been Already Posted", MsgBoxStyle.OkOnly, "Cannot Delete")
                Exit For
            Else
                'SREP_COMM_ADJ_NOs.Add(grow.Cells("SREP_COMM_ADJ_NO").Value)
                SREP_COMM_ADJ_NOs &= ",'" & grow.Cells("SREP_COMM_ADJ_NO").Value & "'"
            End If
        Next
        'grdSOTSCOM2.Tag = SREP_COMM_ADJ_NOs
        grdSOTSCOM2.Tag = "Delete from SOTSCOM2 where SREP_COMM_ADJ_NO in (" & Mid(SREP_COMM_ADJ_NOs, 2) & ")"
    End Sub

#End Region

#Region "From Controls"

    Private Sub cbeOPS_YYYYPP_ValueChanged(sender As System.Object, e As System.EventArgs) Handles cbeOPS_YYYYPP.ValueChanged
        Load_SOTSCOMC()
    End Sub

    Private Sub cmdAddMiscAdj_Click(sender As System.Object, e As System.EventArgs) Handles cmdAddMiscAdj.Click
        Dim SREP_COMM_ADJ_MISC As Decimal = Val(numSREP_COMM_ADJ_MISC.Value)
        Dim SREP_COMM_ADJ_NOTE As String = txtSREP_COMM_ADJ_NOTE.Text
        If SREP_COMM_ADJ_MISC = 0 Or SREP_COMM_ADJ_NOTE = "" Then
            MsgBox("Amount and Note are Mandatory", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        If MsgBox("OK to add a Commission Adjustment for " & Format(SREP_COMM_ADJ_MISC, "$#,##0.00"), MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
            Exit Sub
        End If

        Dim SREP_CODE As String = grdSOTSCOMS.ActiveRow.Cells("SREP_CODE").Value

        Dim rowSOTSCOM2 As DataRow = dst.Tables("SOTSCOM2").NewRow
        With rowSOTSCOM2
            .Item("SREP_COMM_ADJ_NO") = ASCMAIN1.Next_Control_No("SOTSCOM2.SREP_COMM_ADJ_NO")
            .Item("SREP_CODE") = SREP_CODE
            .Item("SREP_COMM_ADJ_MISC") = SREP_COMM_ADJ_MISC
            .Item("SREP_COMM_ADJ_NOTE") = SREP_COMM_ADJ_NOTE
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
        End With
        dst.Tables("SOTSCOM2").Rows.Add(rowSOTSCOM2)
        Update_Record_TDA("SOTSCOM2")

        Dim COMMS_MISC As Decimal = Val(grdSOTSCOMS.ActiveRow.Cells("COMMS_MISC").Value & "") + SREP_COMM_ADJ_MISC
        grdSOTSCOMS.ActiveRow.Cells("COMMS_MISC").Value = COMMS_MISC
        grdSOTSCOMS.ActiveRow.Update()

        MsgBox("Misc Commission Adjustment has been Updated", MsgBoxStyle.OkOnly, "Verification")

        txtSREP_COMM_ADJ_NOTE.Clear()
        numSREP_COMM_ADJ_MISC.Value = 0
    End Sub

    Private Sub tabSummary_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabSummary.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabSummary()
    End Sub

    Private Sub tabInvoices_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabInvoices.SelectedTabChanged
        Select Case e.Tab.Key
            Case "Invoices"
                Position_Controls(grdSOTINVH1)
            Case "Retail Sales"
                Position_Controls(grdRSTRETLC)
            Case "Overrides"
                Position_Controls(grdSOTSCOMO)
        End Select
    End Sub

#End Region

#Region "Form Procedures"

    Sub SETUP_grdSOTINVH2()
        If grdSOTINVH1.ActiveRow Is Nothing OrElse Not grdSOTINVH1.ActiveRow.IsDataRow Then
            splSOTINVH1.Panel2Collapsed = True
        Else
            Dim INV_TYPE As String = grdSOTINVH1.ActiveRow.Cells("INV_TYPE").Value
            Dim INV_NO As String = grdSOTINVH1.ActiveRow.Cells("INV_NO").Value
            Fill_Records("SOTINVH2", New String() {INV_TYPE, INV_NO})
            Sort_grdColumns(grdSOTINVH2, "INV_LNO")
            grdSOTINVH2.Text = "Sales Invoice Details for Invoice " & INV_TYPE & ":" & INV_NO
            splSOTINVH1.Panel2Collapsed = False
        End If
    End Sub

    Sub Setup_grdSOTSCOM2()

        txtSREP_COMM_ADJ_NOTE.Clear()
        numSREP_COMM_ADJ_MISC.Value = 0

        If grdSOTSCOMS.ActiveRow Is Nothing OrElse Not grdSOTSCOMS.ActiveRow.IsDataRow Then
            splSREP.Panel2Collapsed = True
        Else
            Dim SREP_CODE As String = grdSOTSCOMS.ActiveRow.Cells("SREP_CODE").Value & String.Empty
            Dim dvw As DataView = DirectCast(grdSOTSCOM2.DataSource, DataTable).DefaultView
            dvw.RowFilter = "SREP_CODE = '" & SREP_CODE & "'"
            Sort_grdColumns(grdSOTSCOM2, "SREP_CODE")
            grdSOTSCOM2.Text = "Misc Commission Adjustments for Sales Rep " & SREP_CODE
            splSREP.Panel2Collapsed = False
        End If
    End Sub

    Sub Print_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Printing Salesman Commission Report")

        Print_Report_Begin()
        CR_params.Add("NOTES", "1")
        Generate_Report("BMRLIST1", "Bill of Materials", "")
        Print_Report_End()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Load_SOTSCOMC()
        If SELECTION_NO = 0 Then Exit Sub
        OPS_YYYYPP = cbeOPS_YYYYPP.Value & ""
        If OPS_YYYYPP = "" Then Exit Sub

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Creating Summary Sheets")

        Fill_Records("SOTSCOMC", OPS_YYYYPP)
        Sort_grdColumns(grdSOTSCOMC, "CUST_CODE")
        grdSOTSCOMC.Text = "Sales Summary by Customer for " & cbeOPS_YYYYPP.Text

        Fill_Records("SOTSCOM2", OPS_YYYYPP)

        Fill_Records("SOTSCOMS", OPS_YYYYPP)
        Sort_grdColumns(grdSOTSCOMS, "SREP_CODE")
        grdSOTSCOMS.Text = "Sales Summary by Sales Rep for " & cbeOPS_YYYYPP.Text

        'Fill_Records("SOTSCOM1", String.Empty, True, "SELECT '0', SOTSCOM1.* FROM SOTSCOM1 WHERE OPS_YYYYPP = '" & OPS_YYYYPP & "'")
        Fill_Records("SOTSCOM1", OPS_YYYYPP)
        Sort_grdColumns(grdSOTSCOM1, "SREP_CODE")
        grdSOTSCOM1.Text = "Commission Reports by Sales Rep for " & cbeOPS_YYYYPP.Text

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Position_Controls(ByRef grd As Infragistics.Win.UltraWinGrid.UltraGrid)
        If SELECTION_NO = 0 Then Exit Sub

        If grd.Rows.Count = 0 Then Exit Sub
        If grd.Rows.Count <> 0 AndAlso Not grd.Rows(0).IsDataRow Then Exit Sub

        If Not ScreenMode Then Exit Sub
        Try
            Dim r As System.Drawing.Rectangle
            Dim SREP_CODE As String = "SREP_CODE"
            Dim SREP_COMM_PCT As String = "SREP_COMM_PCT"

            If grd.Name = "grdRSTRETLC" Then
                SREP_CODE = "SELL_CODE"
                SREP_COMM_PCT = "SELL_COMM_PCT"
            End If

            cmbSREP_CODE.Parent = grd
            r = grd.ActiveRowScrollRegion.FirstRow.Cells(SREP_CODE).GetUIElement().ClipRect
            cmbSREP_CODE.Width = grd.DisplayLayout.Bands(0).Columns(SREP_CODE).Header.SizeResolved.Width
            cmbSREP_CODE.Left = r.Left
            'txtStore.Top = grdSOTPICK1.Top

            numSREP_COMM_PCT.Parent = grd
            r = grd.ActiveRowScrollRegion.FirstRow.Cells(SREP_COMM_PCT).GetUIElement().ClipRect
            numSREP_COMM_PCT.Width = grd.DisplayLayout.Bands(0).Columns(SREP_COMM_PCT).Header.SizeResolved.Width
            numSREP_COMM_PCT.Left = r.Left
            'txtStore.Top = grdSOTPICK1.Top

        Catch ex As Exception

        End Try

    End Sub

    Sub Setup_tabSummary()
        UltraExplorerBar1.Groups("Commission Adjustments").Visible = ScreenMode And tabSummary.SelectedTab.Key = "by Sales Rep" And OPS_YYYYPP < LYP
    End Sub

    Private Sub GetTotals()

        dst.Tables("TOTALS").Clear()

        Dim sales As Decimal = Val(dst.Tables("SOTINVH1").Compute("SUM(SREP_COMM)", "") & String.Empty)
        Dim retails As Decimal = Val(dst.Tables("RSTRETLC").Compute("SUM(SELL_COMM)", "") & String.Empty)
        Dim override As Decimal = Val(dst.Tables("SOTSCOMO").Compute("SUM(SREP_COMM)", "") & String.Empty)

        dst.Tables("TOTALS").Rows.Add(New Object() {"A", "Sales", sales})
        dst.Tables("TOTALS").Rows.Add(New Object() {"B", "Retails", retails})
        dst.Tables("TOTALS").Rows.Add(New Object() {"C", "Overrides", override})

        Sort_grdColumns(grdTotals, "CODE")

    End Sub

    Private Sub EmailCommissionReports()

        Try
            'If ASCMAIN1.DBS_SERVER <> "AHA" OrElse ASCMAIN1.DBS_COMPANY <> "AHA" Then
            If ASCMAIN1.CLIENT <> "AHA" Then
                Exit Sub
            End If

            EMsg = String.Empty
            Dim reportFolder As String = ASCMAIN1.Folders("Archive") & "Reports\" & ASCMAIN1.SOLUTION & "\"
            If Not reportFolder.EndsWith("\") Then
                reportFolder &= "\"
            End If

            If Not My.Computer.FileSystem.DirectoryExists(reportFolder) Then
                MessageBox.Show("The reports Archive Directory cannot be located: " & reportFolder, "Email Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            If dst.Tables("SOTSCOM1").Select("SEL = '1'").Length = 0 Then
                MessageBox.Show("You must select atleast one Sales Rep", "Email Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim tblSOTSREP1 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM SOTSREP1", "SOTSREP1")
            For Each row As DataRow In dst.Tables("SOTSCOM1").Select("SEL = '1'", "SREP_CODE_MGR")

                row.Item("SEL") = "0"

                Dim srep_code As String = row.Item("SREP_CODE_MGR") & String.Empty
                If srep_code.Length = 0 Then
                    srep_code = row.Item("SREP_CODE") & String.Empty
                End If
                Dim rowSOTSREP1 As DataRow = tblSOTSREP1.Rows.Find(srep_code)
                If rowSOTSREP1 Is Nothing Then
                    EMsg &= vbCrLf & "Cannot locate master file for sales rep: " & srep_code
                    Continue For
                End If

                If rowSOTSREP1.Item("SREP_EMAIL") & String.Empty = String.Empty Then
                    EMsg &= vbCrLf & "Cannot locate email address for sales rep: " & srep_code
                    Continue For
                End If

                Dim reportNo As String = row.Item("REPORT_NO")

                If reportNo.Length = 0 Then
                    EMsg &= vbCrLf & "Cannot locate Report (" & reportNo & ") file for sales rep: " & srep_code
                    Continue For
                End If

                Dim reportPath As String = reportFolder & reportNo.Substring(0, 5) & "\" & ASCMAIN1.SOLUTION & "_" & reportNo & ".PDF"

                If Not My.Computer.FileSystem.FileExists(reportPath) Then
                    EMsg &= vbCrLf & "Cannot locate Report (" & reportNo & ") file for sales rep: " & srep_code
                    Continue For
                End If

                Dim ATTACHMENTs As New Dictionary(Of String, String)
                ATTACHMENTs.Add(reportPath, reportPath)

                Dim SUBJECT As String = ""
                SUBJECT = "Commission Report for Sales Rep (" & rowSOTSREP1.Item("SREP_NAME") & ") for Period " & OPS_YYYYPP

                ' Concatentate and process all email addresses
                Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
                EMAIL_ADDRESSs.Add(rowSOTSREP1.Item("SREP_EMAIL") & String.Empty, rowSOTSREP1.Item("SREP_EMAIL") & String.Empty)

                Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                       (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs, _
                         SUBJECT, "INV", True, False, CUST_CODE, row.Item("SREP_CODE"), "Sales Rep")


            Next

        Catch ex As Exception

        End Try

        If EMsg.Length > 0 Then
            MessageBox.Show(EMsg, "Email Errors", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

    End Sub

    Private Function SetMonthlyCommissions(ByVal Period As String) As Boolean

        Try
            BeginTrans()

            SetMonthlyCommissions = True

            Dim sql As String = String.Empty

            ' Clear any Retail Sales already imported - Do not clear customers who retail sales commissions come from the Import Screen - RSFCUST1
            sql = "Delete from RSTRETLC where OPS_YYYYPP = '" & Period & "' and CUST_CODE NOT IN (SELECT CUST_CODE FROM ARTCUST1 WHERE CUST_EDI_COMM_SEP = '1')"
            ASCDATA1.ExecuteSQL(sql)

            ' Retail Sales - Do not insert customers who retail sales commissions come from the Import Screen - RSFCUST1
            sql = " INSERT INTO RSTRETLC  SELECT CUST_CODE, CUST_STORE_NO, OPS_YYYYPP, SUM(QTY_SOLD) QTY_SOLD, SUM(AMT_SOLD) AMT_SOLD"
            sql &= " , NULL, NULL, NULL, NULL, NULL"
            sql &= "  FROM RSTRETL4"
            sql &= " where OPS_YYYYPP = '" & Period & "'"
            sql &= " and CUST_CODE NOT IN (SELECT CUST_CODE FROM ARTCUST1 WHERE CUST_EDI_COMM_SEP = '1')"
            sql &= " GROUP BY CUST_CODE, CUST_STORE_NO, OPS_YYYYPP"
            ASCDATA1.ExecuteSQL(sql)

            ' Update Retail Sales with Sell-Thru Rep
            sql = " BEGIN DECLARE CURSOR C1 IS"
            sql &= " SELECT CUST_CODE, CUST_STORE_NO, SELL_CODE "
            sql &= " 	   FROM ARTCUST2 "
            sql &= " 	   WHERE (CUST_CODE, CUST_STORE_NO) IN"
            sql &= " 	   (SELECT DISTINCT CUST_CODE, CUST_STORE_NO FROM RSTRETLC WHERE OPS_YYYYPP = '" & Period & "');"
            sql &= " BEGIN FOR R1 IN C1 LOOP"
            sql &= "                   UPDATE RSTRETLC "
            sql &= " 				  	SET SELL_CODE = R1.SELL_CODE"
            sql &= " 						 WHERE  OPS_YYYYPP = '" & Period & "' "
            sql &= " 						 AND CUST_CODE = R1.CUST_CODE "
            sql &= " 						 AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= " END LOOP;"
            sql &= " END;"
            sql &= " END;"
            ASCDATA1.ExecuteSQL(sql)

            ' Update Sell Thru Comm %
            sql = " UPDATE RSTRETLC SET SELL_COMM_PCT = "
            sql &= " (SELECT NVL(SELL_COMM_PCT, 0) FROM SOTSREP1 WHERE SREP_CODE = RSTRETLC.SELL_CODE)"
            sql &= " WHERE OPS_YYYYPP = '" & Period & "'"
            ASCDATA1.ExecuteSQL(sql)

            ' Set all non assigned Retail Sales to 98 show they show up on the commission report
            sql = " UPDATE RSTRETLC SET SELL_CODE = '98', SELL_COMM_PCT = 0 WHERE SELL_CODE IS NULL AND OPS_YYYYPP = '" & Period & "'"
            ASCDATA1.ExecuteSQL(sql)

            CommitTrans()

            Load_SOTSCOMC()

        Catch ex As Exception
            Rollback(ex.Message)
            SetMonthlyCommissions = False
        End Try

    End Function

#End Region

End Class