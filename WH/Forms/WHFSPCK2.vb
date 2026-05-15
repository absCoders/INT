Imports System.Drawing
Imports Infragistics.Win.UltraWinGrid

Public Class WHFSPCK2
    Dim WHSE_CODE As String = String.Empty
    Dim LP_CODE As String = String.Empty
    Dim rowICTWHSE1 As DataRow = Nothing
    Dim rowWHTTPLP1 As DataRow = Nothing

    Dim WHTSTYLX As String = String.Empty

    ' NEED TO MT AGAINST REL & DE-REL FOR THE WHSE SELECTED
    ' NEED TO MT FOR SENDITEMS
    ' DO NOT ALLOW PT PRINT FOR A LP WHSE
    ' CHECK MT ON DESIGN RECALL OF SHIPMENTS/PICK TICKETS
    ' CHECK EVENT PROCEDURE FIRING WHEN CLICKING CANCEL - WHEN CYCLING MODES TO FALSE, DONT WANT TO LOAD_SOTSHIPX

    Dim SOTSHIPX As String = String.Empty
    Dim Shipments As Integer = 0
    Dim LP_XNO As String = String.Empty
    Dim ASW As New Dictionary(Of String, String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        btnCARS.Visible = ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz"
        chkInitialize.Visible = ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz"

        With dst

            ASCMAIN1.sql = "Select SHIP_BOL_NO, '1' SEL, '1' EDI856, '1' SHIP_CART_REQD from SOTSHIP1 where ROWNUM < 1"
            SOTSHIPX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add Primary Key (SHIP_BOL_NO)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add ORDR_NO VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add ORDR_ADDR_TYPE_ST VARCHAR2(2)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add CUST_CODE VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("CREATE INDEX I_" & SOTSHIPX & "_1 ON " & SOTSHIPX & " (ORDR_NO)")

            ASCMAIN1.sql = "Select SOTSHIP1.*, SOTSHIPX.SEL, SOTSHIPX.EDI856" & vbCrLf _
                & ", SOTSHIPX.ORDR_NO, SOTSHIP1.SHIP_ADDR_TYPE ORDR_ADDR_TYPE_ST" & vbCrLf _
                & ", WHTSHIPX.LP_XNO LP_XNO_XMIT" & vbCrLf _
                & ", SOTORDR0.CUST_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
                & ", SOTORDR0.ORDR_CUST_PO, SOTORDR0.ORDR_DATE, SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE" & vbCrLf _
                & " from " & SOTSHIPX & " SOTSHIPX, SOTSHIP1, SOTORDR0, ARTCUST1, WHTSHIPX" & vbCrLf _
                & " where SOTSHIP1.SHIP_BOL_NO = SOTSHIPX.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE" & vbCrLf _
                & "   and WHTSHIPX.LP_XNO (+) = SOTSHIP1.LP_XNO" & vbCrLf _
                & "   and WHTSHIPX.SHIP_BOL_NO (+) = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_STATUS = 'P'"
            Create_TDA(.Tables.Add("SOTSHIPX"), SOTSHIPX, "**", 0, True, "", 1, "SEL")
            '.Tables("SOTSHIPX").Columns.Add("SEL")

            Create_TDA(.Tables.Add, "WHTSHIPX", "*")

            ASCMAIN1.sql = "Select ICTWHSE1.WHSE_CODE, ICTWHSE1.WHSE_DESC, ICTWHSE1.LP_CODE" & vbCrLf _
                & ", X.SHIPS" & vbCrLf _
                & "  from ICTWHSE1" & vbCrLf _
                & ", (Select SOTSHIP1.WHSE_CODE, Count (*) SHIPS" & vbCrLf _
                & "  from SOTSHIP1" & vbCrLf _
                & "where SOTSHIP1.SHIP_STATUS = 'P' and NVL(SOTSHIP1.LP_STATUS, '0') = '0'" & vbCrLf _
                & " group by SOTSHIP1.WHSE_CODE) X" & vbCrLf _
                & " where ICTWHSE1.LP_CODE is Not Null" & vbCrLf _
                & "   and X.WHSE_CODE (+) = ICTWHSE1.WHSE_CODE"
            Create_TDA(.Tables.Add, "ICTWHSEX", "**", 0, False)
            .Tables("ICTWHSEX").Columns("SHIPS").DataType = GetType(System.Int32)


            ASCMAIN1.sql = "Select SOTPICK1.*, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE FROM SOTPICK1, SOTORDR1 WHERE SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO AND SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTORDR2.ITEM_CODE, SOTORDR2.ITEM_DESC" _
                & ", SUM (SOTPICK2.PICK_QTY) PICK_QTY" _
                & ", SUM (SOTPICK2.PICK_QTY_CONF) PICK_QTY_CONF" _
                & ", SUM (SOTPICK2.PICK_QTY_CANC) PICK_QTY_CANC" _
                & ", SUM (SOTPICK2.PICK_QTY_BACK) PICK_QTY_BACK" _
                & ", SUM (SOTPICK2.PICK_QTY_CANC_REL) PICK_QTY_CANC_REL" _
                & ", SUM (SOTPICK2.PICK_QTY_BACK_REL) PICK_QTY_BACK_REL" _
                & " from SOTPICK2, SOTORDR2, SOTPICK1 " _
                & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" _
                & " and SOTPICK2.PICK_NO = SOTPICK1.PICK_NO and SOTPICK1.SHIP_BOL_NO = :PARM1" _
                & " group by SOTORDR2.ITEM_CODE, SOTORDR2.ITEM_DESC"
            Create_TDA(.Tables.Add, "SOTPICKX", "**", 0, False, "V", 0)

            If ASCMAIN1.CLIENT = "AHA" Then
                ASCMAIN1.sql = "Select * from WHTLPXN1 where LP_XNO_SOURCE IN ('" & MENU_ITEM_OBJECT & "', 'SERVICE')" _
                    & " and INIT_DATE >= :PARM1 and INIT_DATE -1  < :PARM2"
            Else
                ASCMAIN1.sql = "Select * from WHTLPXN1 where LP_XNO_SOURCE = '" & MENU_ITEM_OBJECT & "'" _
                    & " and INIT_DATE >= :PARM1 and INIT_DATE -1  < :PARM2"
            End If
            Create_TDA(.Tables.Add, "WHTLPXN1", "**", 0, True, "DD", 1)

            Create_TDA(dst.Tables.Add, "TATCNTRY", "*", 0, False)

            If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" Then

                ' NOTE: This may no longer be accurate after default warehouse change (hard-coded CLA) Dec 2025
                ASCMAIN1.sql = "Select SOTSHIP1.*, SOTORDR0.CUST_CODE" & vbCrLf _
                    & " from SOTSHIP1,SOTORDR0" & vbCrLf _
                    & " where SOTSHIP1.SHIP_STATUS = 'P'" & vbCrLf _
                    & "   and SOTSHIP1.SHIP_ROUTING_REQD = '1' and NVL(SOTSHIP1.SHIP_ROUTING_IND,'0') = '0'" & vbCrLf _
                    & "   and SOTSHIP1.WHSE_CODE in ('CLA','CLARTN')" & vbCrLf _
                    & "   and SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO"
                Create_TDA(.Tables.Add, "SOTSHIP1_CARS", "**", 0, False, "", 1)

                ASCMAIN1.sql = "Select SOTPICK1.* from SOTPICK1" & vbCrLf _
                    & " where SHIP_BOL_NO = :PARM1" & vbCrLf _
                    & "   and PICK_STATUS = 'P'"
                Create_TDA(.Tables.Add, "SOTPICK1_CARS", "**", 0, False, "V", 1)

                ASCMAIN1.sql = "Select SOTSHIPD.* from SOTSHIPD" & vbCrLf _
                    & " where CUST_CODE = :PARM1" & vbCrLf _
                    & "   and CUST_STORE_NO = :PARM2"
                Create_TDA(.Tables.Add, "SOTSHIPD", "**", 0, True, "VV", 1)

                ASCMAIN1.sql = "Select SOTSHIPG.*" & vbCrLf _
                    & " from SOTSHIPG" & vbCrLf _
                    & " where SOTSHIPG.SHIP_TO_CODE = :PARM1" & vbCrLf _
                    & "   and SOTSHIPG.ROUTING_RULE_STATUS = 'A'"
                Create_TDA(.Tables.Add, "SOTSHIPG", "**", 0, False, "V", 1)

                ASCMAIN1.sql = "Select SOTSHIPY.*" & vbCrLf _
                    & ", SOTSHIPS.DEL_METHOD, SOTSHIPS.SHIP_VIA_DESC, SOTSHIPS.SHIP_VIA_SCAC" & vbCrLf _
                    & ", SOTSHIPS.CARRIER, SOTSHIPS.SHIP_METHOD" & vbCrLf _
                    & " from SOTSHIPY, SOTSHIPG, SOTSHIPS" & vbCrLf _
                    & " where SOTSHIPG.SHIP_TO_CODE = :PARM1" & vbCrLf _
                    & "   and SOTSHIPG.ROUTING_RULE_STATUS = 'A'" & vbCrLf _
                    & "   and SOTSHIPY.ROUTING_RULE_NO = SOTSHIPG.ROUTING_RULE_NO" & vbCrLf _
                    & "   and SOTSHIPS.SHIP_VIA (+) = SOTSHIPY.SHIP_VIA"
                Create_TDA(.Tables.Add, "SOTSHIPY", "**", 0, False, "V", 1)


                ASCMAIN1.sql = "Select SOTSHIPI.* from SOTSHIPI" & vbCrLf _
                    & " where SHIP_TO_CODE = :PARM1" & vbCrLf _
                    & "   and CALL_IN_RULE_STATUS = 'A'" ' & vbCrLf _
                Create_TDA(.Tables.Add, "SOTSHIPI", "**", 0, False, "V", 1)

                Create_TDA(.Tables.Add, "SOTSHIPR", "*")

                ASCMAIN1.sql = "Select Sum (SOTPICK2.PICK_QTY) PICK_QTY, Sum (SOTPICK2.PICK_QTY * ICTITEM1.ITEM_WEIGHT) PICK_WGT" & vbCrLf _
                    & ", MIN (ICTITEM1.ITEM_WEIGHT) ITEM_WGT_MIN" & vbCrLf _
                    & ", MIN (CASE WHEN NVL(ICTITEM1.ITEM_WEIGHT,0) = 0 THEN ICTITEM1.ITEM_CODE ELSE NULL END) ITEM_CODE" & vbCrLf _
                    & "  from SOTPICK1,SOTPICK2,SOTORDR2,ICTITEM1" & vbCrLf _
                    & "  where SOTPICK1.PICK_NO = :PARM1 and SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                    & "    and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                    & "    and ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE"
                Create_TDA(.Tables.Add, "SOTPICK1_CARS_WGT", "**", 0, False, "V", 1)

                ASCMAIN1.sql = "Select
                          SOTSHIP1.SHIP_BOL_NO, SOTSHIP1.SHIP_ADDR_TYPE, SOTSHIP1.SHIP_ADDR_CODE
                        , SOTSHIP1.ORDR_GROUP_NO, SOTSHIP1.FRT_TERMS, SOTSHIP1.WHSE_CODE
                        , SOTSHIP1.LP_STATUS, SOTSHIP1.LP_XNO, SOTSHIP1.LP_XMIT_DATE
                        , SOTSHIP1.SHIP_ROUTING_REQD, SOTSHIP1.SHIP_ROUTING_IND
                        , SOTSHIP1.SHIP_PRE_AUTH_REQD, SOTSHIP1.SHIP_PRE_AUTH_NO
                        , SOTSHIP1.SHIP_SHIP_DATE, SOTSHIP1.SHIP_CANCEL_DATE, SOTSHIP1.INIT_DATE SHIP_DATE_REL
                        , SOTSHIP1.SHIP_PICK_TICKET_COUNT, SOTSHIP1.SHIP_PICK_QTY, SOTSHIP1.SHIP_ITEMS_COUNT, SOTSHIP1.SHIP_LINE_ITEMS_COUNT
                        , SOTSHIPR.ORDR_SHIP_DATE,SOTSHIPR.ORDR_CANCEL_DATE,SOTSHIPR.SHIP_TO_CODE,SOTSHIPR.DEL_METHOD
                        , SOTSHIPR.OPS_LEAD_TIME_DAYS,SOTSHIPR.CALL_IN_NOTICE_DAYS
                        , SOTSHIPR.SHIP_CALL_IN_EARLIEST,SOTSHIPR.SHIP_CALL_IN_LATEST
                        , SOTSHIPR.SHIP_XMIT_EARLIEST,SOTSHIPR.SHIP_XMIT_LATEST
                        , SOTSHIPR.SHIP_METHOD,SOTSHIPR.SHIP_CARRIER,SOTSHIPR.SHIP_WEIGHT
                        , SOTSHIPR.SHIP_SUGG_CALL_IN_DATE,SOTSHIPR.SHIP_SUGG_PICK_UP_DATE
                        , SOTSHIPR.SHIP_ROUTING_RULE_USED,SOTSHIPR.SHIP_CALL_IN_RULE_USED
                        , SOTSHIPR.SHIP_ROUTING_ATTEMPTS, SOTSHIPR.SHIP_BILL_FRT_TO
                        , SOTSHIPR.INIT_DATE SHIP_DATE_ROUTED_INIT, SOTSHIPR.LAST_DATE SHIP_DATE_ROUTED_LAST
                        , SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO
                        , SOTSHIPR.CARRIER_ACCT_TYPE, SOTSHIPR.CARRIER_ACCT_NO" & vbCrLf _
                    & " from SOTSHIPR,SOTSHIP1,SOTORDR0" & vbCrLf _
                    & " where SOTSHIP1.SHIP_BOL_NO = SOTSHIPR.SHIP_BOL_NO" & vbCrLf _
                    & "   and SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
                    & "   and SOTSHIP1.SHIP_STATUS = 'P' AND SOTSHIP1.SHIP_ROUTING_REQD = '1'"
                Create_TDA(.Tables.Add, "SOTSHIPR_VIEW", "**", 0, False, "", 1)


                ASCMAIN1.sql = "Select SOTSHIPV.*" & vbCrLf _
                    & " from SOTSHIPV" & vbCrLf _
                    & " where EVENT_DATE > TRUNC(SYSDATE)"
                Create_TDA(.Tables.Add, "SOTSHIPV_ALL", "**", 0, False)

                ASCMAIN1.sql = "Select SOTSHIPG.*" & vbCrLf _
                    & " from SOTSHIPG"
                Create_TDA(.Tables.Add, "SOTSHIPG_ALL", "**", 0, False)

                ASCMAIN1.sql = "Select SOTSHIPI.*" & vbCrLf _
                    & " from SOTSHIPI"
                Create_TDA(.Tables.Add, "SOTSHIPI_ALL", "**", 0, False)

                Create_TDA(.Tables.Add, "SOTPARMH", "*", 0, False)

                ASCMAIN1.sql = "Select SOTSHIPV.*" & vbCrLf _
                    & " from SOTSHIPV" & vbCrLf _
                    & " where SOTSHIPV.SHIP_BOL_NO = :PARM1"
                Create_TDA(.Tables.Add, "SOTSHIPV", "**", 0, False, "V", 0)

                ASCMAIN1.sql = "Select SOTORDR2.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_WEIGHT" & vbCrLf _
                    & ", SUM(SOTPICK2.PICK_QTY) PICK_QTY, SUM(SOTPICK2.PICK_QTY * ICTITEM1.ITEM_WEIGHT) PICK_WGT" & vbCrLf _
                    & " from SOTORDR2, SOTPICK2, SOTPICK1, ICTITEM1" & vbCrLf _
                    & " where SOTPICK1.SHIP_BOL_NO = :PARM1" & vbCrLf _
                    & "   and SOTPICK2.PICK_NO = SOTPICK1.PICK_NO and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
                    & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                    & "   and ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE" & vbCrLf _
                    & " group by SOTORDR2.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_WEIGHT"
                Create_TDA(.Tables.Add, "SOTSHIP1_WGT", "**", 0, False, "V", 1)

                ASCMAIN1.sql = "Select SOTPICK1.*" & vbCrLf _
                    & " from SOTPICK1, SOTORDR1" & vbCrLf _
                    & " where SOTPICK1.SHIP_BOL_NO = :PARM1" & vbCrLf _
                    & "   and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO"
                Create_TDA(.Tables.Add, "SOTSHIPP", "**", 0, False, "V", 1)


                ASCMAIN1.sql = "Select SOTSHIPE.*" & vbCrLf _
                    & " from SOTSHIPE" & vbCrLf _
                    & " where SOTSHIPE.SHIP_TO_TYPE = :PARM1 and CUST_CODE = :PARM2"
                Create_TDA(.Tables.Add, "SOTSHIPE", "**", 0, True, "VV", 2)

                ASCMAIN1.sql = "Select SOTSHIPQ.* from SOTSHIPQ"
                Create_TDA(.Tables.Add, "SOTSHIPQ", "**", 0, False, "", 2)
                Fill_Records("SOTSHIPQ")

                ASCMAIN1.sql = "Select SOTSHIPF.* from SOTSHIPF"
                Create_TDA(.Tables.Add, "SOTSHIPF", "**", 0, False, "", 2)
                Fill_Records("SOTSHIPF")

                ASCMAIN1.sql = "Select SOTSHIPZ.* from SOTSHIPZ"
                Create_TDA(.Tables.Add, "SOTSHIPZ", "**", 0, False, "", 2)
                Fill_Records("SOTSHIPZ")

                ASCMAIN1.sql = "Select SOTSHIPW.*" & vbCrLf _
                    & " from SOTSHIPW"
                Create_TDA(.Tables.Add, "SOTSHIPW", "**", 0, False)
                Fill_Records("SOTSHIPW")

                ASCMAIN1.sql = "Select SOTSHIPS.*" & vbCrLf _
                    & " from SOTSHIPS"
                Create_TDA(.Tables.Add, "SOTSHIPS", "**", 0, False)
                Fill_Records("SOTSHIPS")

                With .Tables.Add("SOTSHIP1_MATRIX")
                    .Columns.Add("CALL_IN_DATE", GetType(System.DateTime))
                    .Columns.Add("CALL_IN_DAY")
                    For I As Integer = 1 To 20
                        .Columns.Add($"PICK_UP_{Format(I, "00")}", GetType(System.Int32))
                    Next
                    .PrimaryKey = New DataColumn() { .Columns("CALL_IN_DATE")}
                End With

            End If

            .Tables.Add("TASKS")
            With .Tables("TASKS")
                .Columns.Add("SEQ_NO", GetType(Int32))
                .Columns.Add("TASK_TIME", GetType(DateTime))
                .Columns.Add("TASK_DESC", GetType(String))
            End With
        End With

        grdICTWHSEX.DataSource = dst.Tables("ICTWHSEX")
        grdSOTSHIPX.DataSource = dst.Tables("SOTSHIPX")

        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdSOTPICKX.DataSource = dst.Tables("SOTPICKX")
        grdWHTLPXN1.DataSource = dst.Tables("WHTLPXN1")

        grdTasks.DataSource = dst.Tables("TASKS")


        If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" Then
            grdSOTSHIPR_VIEW.DataSource = dst.Tables("SOTSHIPR_VIEW")
            grdSOTSHIP1_MATRIX.DataSource = dst.Tables("SOTSHIP1_MATRIX")
            Create_Summary(grdSOTSHIP1_MATRIX, "CALL_IN_DATE", "Count")

            grdSOTSHIPV.DataSource = dst.Tables("SOTSHIPV")
            grdSOTSHIPP.DataSource = dst.Tables("SOTSHIPP")
            grdSOTSHIP1_WGT.DataSource = dst.Tables("SOTSHIP1_WGT")

            grdSOTSHIPV_ALL.DataSource = dst.Tables("SOTSHIPV_ALL")

            With grdSOTSHIPR_VIEW.DisplayLayout.Bands(0)
                .Columns("SHIP_BOL_NO").Header.Fixed = True
                .Columns("SHIP_SHIP_DATE").Header.Fixed = True
                .Columns("SHIP_CANCEL_DATE").Header.Fixed = True
                .Columns("CUST_CODE").Header.Fixed = True
                .Columns("SHIP_ADDR_TYPE").Header.Fixed = True
                .Columns("SHIP_ADDR_CODE").Header.Fixed = True

                For Each gcol As UltraWinGrid.UltraGridColumn In grdSOTSHIPR_VIEW.DisplayLayout.Bands(0).Columns
                    gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                    gcol.Header.Appearance.BackColor = System.Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                    If New String() {"SHIP_CALL_IN_EARLIEST", "SHIP_CALL_IN_LATEST"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    End If
                    If New String() {"SHIP_SHIP_DATE", "SHIP_CANCEL_DATE"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                    End If
                    If New String() {"SHIP_XMIT_EARLIEST", "SHIP_XMIT_LATEST"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                    End If
                    If New String() {"CALL_IN_NOTICE_DAYS"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = System.Drawing.Color.Yellow
                    End If
                    If New String() {"OPS_LEAD_TIME_DAYS"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = System.Drawing.Color.PaleVioletRed
                    End If
                    If New String() {"SHIP_SUGG_CALL_IN_DATE"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = System.Drawing.Color.Gold
                    End If
                    If New String() {"SHIP_PICK_TICKET_COUNT", "SHIP_PICK_QTY", "SHIP_ITEMS_COUNT", "SHIP_LINE_ITEMS_COUNT"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = System.Drawing.Color.PaleTurquoise
                    End If
                Next
            End With

            Set_Read_Only(grpCallInRules, True)

            ASCMAIN1.Add_Value_List(grdSOTSHIPR_VIEW, "SHIP_ROUTING_IND", "Select EVENT_ERROR_TYPE, EVENT_ERROR_DESC From SOTSHIPM")
            Create_Summary(grdSOTSHIPR_VIEW, "SHIP_BOL_NO", "Count")
        Else
            tab0.Tabs("CARS").Visible = False
        End If




        Fill_Records("ICTWHSEX")
        Fill_Records("TATCNTRY")

        Create_Summary(grdICTWHSEX, "WHSE_CODE", "Count")
        Create_Summary(grdICTWHSEX, "SHIPS")



        Create_Summary(grdSOTSHIP1_WGT, "ITEM_CODE", "Count")
        Create_Summary(grdSOTSHIP1_WGT, "PICK_WGT")


        Create_Summary(grdSOTSHIPX, "SHIP_BOL_NO", "Count")
        Create_Summary(grdSOTSHIPX, "SEL")
        Create_Summary(grdSOTPICK1, "PICK_NO", "Count")

        Create_Summary(grdSOTPICKX, "ITEM_CODE", "Count")
        Create_Summary(grdSOTPICKX, New String() _
                       {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK" _
                       , "PICK_QTY_CANC_REL", "PICK_QTY_BACK_REL"})

        Create_Summary(grdWHTLPXN1, "LP_XNO", "Count")


        grdSOTSHIPX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        For Each C As UltraWinGrid.UltraGridColumn In grdSOTSHIPX.DisplayLayout.Bands(0).Columns
            If C.Key = "SEL" Then
                C.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                C.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next

        With grdSOTSHIPX.DisplayLayout.Bands("SOTSHIPX")
            .Columns("SHIP_BOL_NO").Header.Fixed = True
            .Columns("SEL").Header.Fixed = True
            .Columns("CUST_CODE").Header.Fixed = True
        End With
        With grdSOTPICKX.DisplayLayout.Bands("SOTPICKX")
            For Each COLUMN_NAME As String In New String() {"ITEM_CODE", "ITEM_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        calFrom.Value = Now.Date.AddDays(-10)
        calTo.Value = Now.Date

        optPending.ValueList.ValueListItems(1).DisplayText = "Transmitted"

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCrLf & "You must specify a Warehouse"
                Else
                    rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If rowICTWHSE1 Is Nothing Then
                        EMsg &= vbCrLf & "Invalid Value specified for Warehouse"
                    Else
                        If rowICTWHSE1.Item("LP_CODE") & "" = "" Then
                            EMsg &= vbCr & "Warehouse " & Absx1.txtFor("WHSE_CODE").Text & " is not set up as a 3PL"
                        Else
                            rowWHTTPLP1 = LookUp("WHTTPLP1", rowICTWHSE1.Item("LP_CODE"))
                            If rowWHTTPLP1 Is Nothing Then
                                EMsg &= vbCrLf & "Warehouse " & Absx1.txtFor("WHSE_CODE").Text & " Does NOT have a valid value specified for its 3PL"
                            End If
                        End If

                    End If
                End If

                If EMsg = "" Then
                    WHSE_CODE = rowICTWHSE1.Item("WHSE_CODE")
                    LP_CODE = rowICTWHSE1.Item("LP_CODE")

                    If Not ASCMAIN1.Logical_Lock("WHTSPCK1", WHSE_CODE) Then Exit Sub

                End If

            Case "Send"
                Shipments = dst.Tables("SOTSHIPX").Select("SEL = '1' And LP_XMIT_DATE is not Null").Length
                If Shipments > 0 Then
                    EMsg &= vbCr & "There are " & Shipments & " De-Released Shipments Selected. You are not permitted to select De-Released Shipments"
                    Exit Select
                End If

                Shipments = dst.Tables("SOTSHIPX").Select("SEL = '1'").Length
                If Shipments = 0 Then
                    EMsg &= vbCr & "No Shipments Selected"
                    Exit Select
                End If

                ' All shipments for an Order group must be selected
                If Not chkOrderGroup.Checked Then
                    Dim ORDR_GROUP_NO As String = String.Empty
                    For Each row As DataRow In dst.Tables("SOTSHIPX").Select("SEL = '1'", "ORDR_GROUP_NO")
                        If ORDR_GROUP_NO = row.Item("ORDR_GROUP_NO") & String.Empty Then
                            Continue For
                        End If

                        ORDR_GROUP_NO = row.Item("ORDR_GROUP_NO") & String.Empty
                        If dst.Tables("SOTSHIPX").Select("SEL <> '1' AND ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'").Length > 0 Then
                            EMsg &= vbCr & "Order Group (" & ORDR_GROUP_NO & ") does not have all its sales orders selected."
                        End If
                    Next
                End If

                If EMsg = "" Then
                    If optPending.Value = "0" Then
                        If MessageBox.Show("You are about to send " & Shipments & " Shipments Electronically over to the 3PL." _
                                    & vbCrLf _
                                    & "No Changes or De-Releases are Permitted" _
                                    & " to these Orders once they are sent to the 3PL" _
                                    & " without getting the 3PL to Void the corresponding Record in their System." _
                                    & vbCrLf _
                                    & vbCrLf & "OK To Proceed?", "Transmit", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            Exit Sub
                        End If
                    Else
                        EMsg &= vbCr & "Transmit option not available for Transmitted Sales Orders"
                    End If
                End If

            Case "Request 940 Cancel"
                Shipments = dst.Tables("SOTSHIPX").Select("SEL='1'").Length
                If Shipments = 0 Then
                    EMsg &= vbCr & "No Shipments Selected"
                End If

                If EMsg = "" Then
                    If MsgBox("You are about to send a Request to Cancel " & Shipments & " Shipments to the 3PL" _
                             & vbCrLf _
                             & vbCrLf & "You should have communicated with your CSR before doing this" _
                             & vbCrLf & " to make sure that these Pick Tickets are able to be Cancelled." _
                             & vbCrLf _
                             & vbCrLf & "Once you get a confirmation email, you should then De-Transmit these Shipments." _
                             & vbCrLf _
                             & vbCrLf & "OK To Proceed?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Send"
                If SendShipmentsTo3PL(Absx1.txtFor("WHSE_CODE").Text) Then
                    Mode_Settings(False)
                End If

            Case "Request 940 Cancel"
                Request_940_Cancel()
                Load_SOTSHIPX()

            Case "Cancel"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Send").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Transmission Controls").Visible = ScreenMode
                .Groups("Screen Control").Items("Request 940 Cancel").Settings.Enabled = DefaultableBoolean.False
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tab0.Visible = Not ScreenMode
        splShipments.Visible = ScreenMode

        If ScreenMode Then
            grdSOTSHIPX.Dock = DockStyle.None
            grdSOTSHIPX.Parent = splShipments.Panel1
            grdSOTSHIPX.Dock = DockStyle.Fill
            grdSOTSHIPX.Text = "Shipments Pending Transmission to 3PL"
            grdSOTSHIPX.DisplayLayout.Bands(0).Columns("SEL").Hidden = False
            grdSOTSHIPX.Visible = True
        Else
            Clear_Record()
            grdSOTSHIPX.Dock = DockStyle.None
            grdSOTSHIPX.Parent = splTransmissions.Panel2
            grdSOTSHIPX.Dock = DockStyle.Fill
            grdSOTSHIPX.Text = "Shipments Transmitted"
            grdSOTSHIPX.DisplayLayout.Bands(0).Columns("SEL").Hidden = True
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTSHIPX", "SOTPICK1", "SOTPICKX", "WHTLPXN1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        optPending.Value = "0"
        Load_WHTLPXN1()
        Fill_Records("ICTWHSEX")
        Setup_tab0()
        UltraExplorerBar1.Groups("Screen Control").Items("Request 940 Cancel").Settings.Enabled = DefaultableBoolean.False
        Clear_All_Filters(grdSOTSHIPX)
        Clear_All_Filters(grdSOTPICK1)
        Clear_All_Filters(grdSOTPICKX)

        Sort_grdColumns(grdTasks, "SEQ_NO")

    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        Load_SOTSHIPX()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

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

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTWHSEX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTSHIPV_ALL, "SS", "Show Filter", "Show GroupBox")

        Load_Popup_Menu(grdSOTSHIPX, "SSSPBBPBBB", "Show Filter", "Show GroupBox", "Show Pins" _
                        , "Select All", "De-Select All", "Select All in Group", "Select All for Customer", "Recall Shipment")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdSOTSHIPX"
                    tlb_pop.Tools("Select All").SharedProps.Visible = True
                    tlb_pop.Tools("De-Select All").SharedProps.Visible = True
                    tlb_pop.Tools("Recall Shipment").SharedProps.Visible = Not ScreenMode

                    If optPending.Value = "D" Then
                        'tlb_pop.Tools("Select All").SharedProps.Visible = False
                    End If
                    tlb_pop.Tools("Recall Shipment").SharedProps.Visible = False

                    tlb_pop.Tools("Select All in Group").SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow)
                    tlb_pop.Tools("Select All for Customer").SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow)

                    If Not ScreenMode Then
                        If grd.ActiveRow.Cells("LP_XNO").Value & "" <> grd.ActiveRow.Cells("LP_XNO_XMIT").Value & "" _
                        Or grd.ActiveRow.Cells("LP_STATUS").Value & "" <> "1" Then
                            tlb_pop.Tools("Recall Shipment").SharedProps.Visible = False
                        End If
                    End If
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Select All"
                Dim selectAll As Int16 = 0

                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows.GetFilteredInNonGroupByRows
                    If EntryMode = "L" AndAlso optPending.Value = "0" Then
                        If grow.Cells("LP_XMIT_DATE").Value & String.Empty = String.Empty Then
                            grow.Cells("SEL").Value = "1"
                            selectAll += 1
                        End If
                    Else
                        grow.Cells("SEL").Value = "1"
                        selectAll += 1
                    End If
                    grow.Update()
                Next

                MsgBox("You have selected " & selectAll & " Records by Selecting All", MsgBoxStyle.OkOnly, "Verification")

            Case "De-Select All"

                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows.GetFilteredInNonGroupByRows
                    grow.Cells("SEL").Value = "0"
                    grow.Update()
                Next

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Recall Shipment"
                Dim SHIP_BOL_NO As String = grdSOTSHIPX.ActiveRow.Cells("SHIP_BOL_NO").Value
                Dim LP_XNO As String = grdSOTSHIPX.ActiveRow.Cells("LP_XNO").Value
                If MsgBox("Are you sure you want to Recall all Pick Tickets and Shipments for Shipment " & SHIP_BOL_NO,
                          MsgBoxStyle.YesNo,
                          "Verification to Recall Pick Tickets and Shipments from a 3PL") <> MsgBoxResult.Yes Then
                    Exit Sub
                End If

                Recall_Shipment(SHIP_BOL_NO, LP_XNO)

            Case "Select All in Group"
                For Each rowSOTWSHIPX As DataRow In dst.Tables("SOTSHIPX").Select("ORDR_GROUP_NO = '" & grd.ActiveRow.Cells("ORDR_GROUP_NO").Value & "'")
                    'rowSOTWSHIPX.Item("SEL") = "1"
                    If EntryMode = "L" AndAlso optPending.Value = "0" Then
                        If rowSOTWSHIPX.Item("LP_XMIT_DATE") & String.Empty = String.Empty Then rowSOTWSHIPX.Item("SEL") = "1"
                    Else
                        rowSOTWSHIPX.Item("SEL") = "1"
                    End If
                Next

            Case "Select All for Customer"
                For Each rowSOTWSHIPX As DataRow In dst.Tables("SOTSHIPX").Select("CUST_CODE = '" & grd.ActiveRow.Cells("CUST_CODE").Value & "'")
                    'rowSOTWSHIPX.Item("SEL") = "1"
                    If EntryMode = "L" AndAlso optPending.Value = "0" Then
                        If rowSOTWSHIPX.Item("LP_XMIT_DATE") & String.Empty = String.Empty Then rowSOTWSHIPX.Item("SEL") = "1"
                    Else
                        rowSOTWSHIPX.Item("SEL") = "1"
                    End If
                Next

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Load", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                Click_Command("Load")
        End Select
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdICTWHSEX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTWHSEX.DoubleClickRow
        Absx1.txtFor("WHSE_CODE").Text = e.Row.Cells("WHSE_CODE").Text
        Click_Command("Load")
    End Sub

    Private Sub grdWHTLPXN1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWHTLPXN1.AfterRowActivate
        Setup_WHTLPXN1()
    End Sub

    Private Sub grdSOTSHIPX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTSHIPX.AfterRowActivate
        Setup_SOTSHIPX()
    End Sub

    Private Sub grdSOTSHIPX_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTSHIPX.BeforeCellUpdate
        If EntryMode = "L" AndAlso optPending.Value = "0" Then
            If e.Cell.Column.Key = "SEL" Then
                If e.Cell.Row.Cells("LP_XMIT_DATE").Value & String.Empty <> String.Empty Then
                    e.Cancel = True
                End If
            End If
        End If
    End Sub

    Private Sub grdSOTSHIPX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTSHIPX.InitializeRow
        If Not ScreenMode And EntryMode = "" Then
            If grdWHTLPXN1.ActiveRow IsNot Nothing Then
                If e.Row.Cells("LP_XNO_XMIT").Value & "" <> grdWHTLPXN1.ActiveRow.Cells("LP_XNO").Value & "" Then
                    e.Row.CellAppearance.BackColor = Drawing.Color.Tomato
                Else
                    e.Row.CellAppearance.BackColor = Drawing.Color.Empty
                End If
            End If

        ElseIf EntryMode = "L" AndAlso optPending.Value = "0" Then
            If e.Row.Cells("LP_XMIT_DATE").Value & String.Empty <> String.Empty Then
                e.Row.CellAppearance.BackColor = Drawing.Color.Tomato
            Else
                e.Row.CellAppearance.BackColor = Drawing.Color.Empty
            End If
        Else
            e.Row.CellAppearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

    Private Sub optPending_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optPending.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        ' dont want to be here if closing down
        Load_SOTSHIPX()
    End Sub

    Private Sub tab0_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tab0()
    End Sub

    Private Sub btnLoadHistory_Click(sender As System.Object, e As System.EventArgs) Handles btnLoadHistory.Click
        Load_WHTLPXN1()
    End Sub

#End Region

#Region "Form Procedures"

    Private Sub AddTask(ByVal TaskDescription As String)
        dst.Tables("TASKS").Rows.Add({dst.Tables("TASKS").Rows.Count + 1, DateTime.Now, TaskDescription})
    End Sub


    Public Function Build_List_of_Objects(Of C As {New})(sql As String) As List(Of C)

        Dim objList As New List(Of C)
        Dim ALL_COLUMNS As Dictionary(Of String, System.Reflection.FieldInfo) _
            = Get_Columns_from_Class(GetType(C))

        Dim tbl As DataTable = ASCDATA1.GetDataTable(sql)
        Dim row_count_total As Int32 = tbl.Rows.Count
        Dim row_counter As Int32 = 0

        For Each row As DataRow In tbl.Rows
            row_counter += 1

            Dim objItem As New C

            If 1 <> 1 Then
                ALL_COLUMNS = Get_Columns_from_Class(GetType(C))
            End If

            For Each COLUMN_NAME In ALL_COLUMNS.Keys
                If row.Item(COLUMN_NAME) & "" = "" Then
                Else
                    'Dim P As System.Reflection.MemberInfo = ALL_COLUMNS(COLUMN_NAME)

                    'If row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.DateTime" Then
                    '    P.GetType().GetProperty(COLUMN_NAME).SetValue(P, row.Item(COLUMN_NAME), Nothing)
                    'ElseIf row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.String" Then
                    '    P.GetType().GetProperty(COLUMN_NAME).SetValue(P, row.Item(COLUMN_NAME), Nothing)
                    'ElseIf row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.Double" Then
                    '    Dim V As Decimal = Val(row.Item(COLUMN_NAME))
                    '    P.GetType().GetProperty(COLUMN_NAME).SetValue(P, V, Nothing)
                    'ElseIf row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.Int32" Then
                    '    Dim V As Int32 = Val(row.Item(COLUMN_NAME))
                    '    P.GetType().GetProperty(COLUMN_NAME).SetValue(P, V, Nothing)
                    'ElseIf row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.Int16" Then
                    '    Dim V As Int16 = Val(row.Item(COLUMN_NAME))
                    '    P.GetType().GetProperty(COLUMN_NAME).SetValue(P, V, Nothing)
                    'Else
                    '    P.GetType().GetProperty(COLUMN_NAME).SetValue(P, Val(row.Item(COLUMN_NAME)), Nothing)
                    'End If

                    If row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.DateTime" Then
                        ALL_COLUMNS(COLUMN_NAME).SetValue(objItem, row.Item(COLUMN_NAME))
                    ElseIf row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.String" Then
                        ALL_COLUMNS(COLUMN_NAME).SetValue(objItem, row.Item(COLUMN_NAME))
                    ElseIf row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.Double" Then
                        Dim V As Decimal = Val(row.Item(COLUMN_NAME))
                        ALL_COLUMNS(COLUMN_NAME).SetValue(objItem, V)
                    ElseIf row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.Int32" Then
                        Dim V As Int32 = Val(row.Item(COLUMN_NAME))
                        ALL_COLUMNS(COLUMN_NAME).SetValue(objItem, V)
                    ElseIf row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.Int16" Then
                        Dim V As Int16 = Val(row.Item(COLUMN_NAME))
                        ALL_COLUMNS(COLUMN_NAME).SetValue(objItem, V)
                    Else
                        ALL_COLUMNS(COLUMN_NAME).SetValue(objItem, Val(row.Item(COLUMN_NAME)))
                    End If

                End If
            Next
            objList.Add(objItem)

            'If row_counter > 100 Then
            '    Exit For
            'End If
        Next

        Return objList
    End Function

    Public Shared Function Get_Columns_from_Class(T As Type) _
        As Dictionary(Of String, System.Reflection.FieldInfo)

        Dim COLUMN_NAMEs As New Dictionary(Of String, System.Reflection.FieldInfo)
        ' Dim COLUMN_NAMEs As New Dictionary(Of String, System.Reflection.PropertyInfo)

        'Dim t As Type = XX.GetType
        Dim fieldName As String
        ' Dim propertyValue As Object

        ' Use each property of the business object passed in 
        'For Each pi As System.Reflection.PropertyInfo In _
        '        T.GetProperties(System.Reflection.BindingFlags.Instance Or _
        '                        System.Reflection.BindingFlags.Public Or _
        '                        System.Reflection.BindingFlags.NonPublic)
        '    ' Get the name and value of the property 
        '    If pi.Name <> "ExtensionData" Then
        '        fieldName = pi.Name
        '        COLUMN_NAMEs.Add(fieldName, pi)
        '    End If

        '    ' Get the value of the property 
        '    ' propertyValue = pi.GetValue(XX, Nothing)
        '    'Console.WriteLine(fieldName & ": " &
        '    'If(propertyValue Is Nothing, "Nothing", propertyValue.ToString))
        'Next

        For Each pi As System.Reflection.FieldInfo In
               T.GetFields(System.Reflection.BindingFlags.Instance Or
                               System.Reflection.BindingFlags.Public Or
                               System.Reflection.BindingFlags.NonPublic)
            If pi.MemberType = Reflection.MemberTypes.Field Then
                fieldName = pi.Name
                If fieldName <> "SQL" Then
                    ' Debug.Write(pi.Name & ":" & pi.MemberType.ToString)
                    COLUMN_NAMEs.Add(fieldName, pi)
                End If
            End If
        Next
        Return COLUMN_NAMEs
    End Function

    Sub Setup_WHTLPXN1()
        If grdWHTLPXN1.ActiveRow Is Nothing Then
            grdSOTSHIPX.Visible = False
        Else
            grdSOTSHIPX.Visible = True
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTSHIPX)
            Dim LP_XNO As String = grdWHTLPXN1.ActiveRow.Cells("LP_XNO").Value
            ASCDATA1.ExecuteSQL("Insert into " & SOTSHIPX & " (SHIP_BOL_NO) Select SHIP_BOL_NO from WHTSHIPX where LP_XNO = '" & LP_XNO & "'")
            Fill_Records("SOTSHIPX")
            Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO")
        End If
    End Sub

    Sub Setup_tab0()
        UltraExplorerBar1.Groups("Transmission History").Visible = Not ScreenMode And tab0.SelectedTab.Key = "Transmissions"
    End Sub

    Sub Load_WHTLPXN1()
        Fill_Records("WHTLPXN1", New Object() {calFrom.Value, calTo.Value})
        Sort_grdColumns(grdWHTLPXN1, "LP_XNO".ToLower)
        Setup_WHTLPXN1()
    End Sub

    Sub Setup_SOTSHIPX()
        If grdSOTSHIPX.ActiveRow Is Nothing OrElse Not grdSOTSHIPX.ActiveRow.IsDataRow Then
            tabShipment.Visible = False
        Else
            tabShipment.Visible = True
            Dim SHIP_BOL_NO As String = grdSOTSHIPX.ActiveRow.Cells("SHIP_BOL_NO").Value & ""
            grdSOTPICK1.Text = "Pick Tickets for Shipment No " & SHIP_BOL_NO
            Fill_Records("SOTPICK1", SHIP_BOL_NO)
            Sort_grdColumns(grdSOTPICK1, "PICK_NO")
            grdSOTPICKX.Text = "Item Summary for Shipment No " & SHIP_BOL_NO
            Fill_Records("SOTPICKX", SHIP_BOL_NO)
            Sort_grdColumns(grdSOTPICKX, "ITEM_CODE")
        End If
    End Sub

    Sub Recall_Shipment(SHIP_BOL_NO As String, LP_XNO As String)

        If Not ASCMAIN1.Logical_Lock("WHTSPCK1", WHSE_CODE) Then Exit Sub
        If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then Exit Sub

        Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NO)
        If rowSOTSHIP1.Item("LP_XNO") & "" <> LP_XNO _
        Or rowSOTSHIP1.Item("LP_STATUS") <> "1" Then
            ' DO NOTHING, SOMETHING HAS CHANGED
        Else
            ASCMAIN1.sql = "Update SOTSHIP1 Set LP_STATUS = '0', LP_XNO = NULL where SHIP_BOL_NO = :PARM1"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New String() {SHIP_BOL_NO})
        End If

        Mode_Settings(False)

    End Sub

    Sub De_Transmit_3PL()

        Dim sqlSHIP_BOL_NO As String = ""
        For Each row As DataRow In dst.Tables("SOTSHIPX").Select("")
            Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO")
            sqlSHIP_BOL_NO &= ",'" & SHIP_BOL_NO & "'"
        Next
        sqlSHIP_BOL_NO = " where SHIP_BOL_NO in (" & Mid(sqlSHIP_BOL_NO, 2) & ")"
        Dim sqlPICK_NO As String = " where PICK_NO in (Select PICK_NO from ADS.SOTPICK1_3PL@ADSIIS" & sqlSHIP_BOL_NO & ")"


        For Each TABLE_NAME As String In New String() _
            {"SOTCART2", "SOTCART1", "SOTPICK2", "SOTPICK1", "SOTSHIP1"}

            ASCMAIN1.Progress("-", TABLE_NAME)

        Next
    End Sub

    Sub Write_Notes(EDI_OUTBOUND_DOC_NO As String, EDI_NTE_TYPE As String, NOTES As String)

        Dim EDI_NTE_SEQ_NO As Int32 = 0
        For Each NOTE As String In Split(NOTES, vbCrLf)
            NOTE = Trim(NOTE)
            Do While NOTE <> ""
                EDI_NTE_SEQ_NO += 1

                Dim EDI_NTE As String
                If NOTE.Length > 40 Then
                    EDI_NTE = Mid(NOTE, 1, 40)
                    NOTE = Mid(NOTE, 41)
                Else
                    EDI_NTE = NOTE
                    NOTE = ""
                End If

                EDI_NTE = Replace(EDI_NTE, "*", "@")

                Dim rowEDT940O4 As DataRow = dst.Tables("EDT940O4").NewRow
                With rowEDT940O4
                    .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                    .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    .Item("EDI_NTE_TYPE") = EDI_NTE_TYPE
                    .Item("EDI_NTE_SEQ_NO") = EDI_NTE_SEQ_NO
                    .Item("EDI_NTE") = EDI_NTE
                End With
                dst.Tables("EDT940O4").Rows.Add(rowEDT940O4)
            Loop
        Next

    End Sub

    Sub Load_SOTSHIPX()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Shipments Queue")

        If optPending.Value = "0" Then
            ASCMAIN1.sql = "Select SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                & ", '0' SEL, SHIP_856_IND EDI856, SOTSHIP1.SHIP_CART_REQD, NULL ORDR_NO" & vbCrLf _
                & ", SHIP_ADDR_TYPE ORDR_ADDR_TYPE_ST, NULL CUST_CODE from SOTSHIP1" _
                & " where SOTSHIP1.WHSE_CODE = '" & WHSE_CODE & "'"

            UltraExplorerBar1.Groups("Screen Control").Items("Send").Text = "Transmit"
            ASCMAIN1.sql &= " and nvl(SOTSHIP1.LP_STATUS,'0') = '0'"
            grdSOTSHIPX.Text = "Shipments Pending Transmission to 3PL (" & WHSE_CODE & ")"

            UltraExplorerBar1.Groups("Screen Control").Items("Request 940 Cancel").Settings.Enabled = DefaultableBoolean.False
            UltraExplorerBar1.Groups("Screen Control").Items("Send").Settings.Enabled = DefaultableBoolean.True
        Else
            ASCMAIN1.sql = "Select SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                 & ", '0' SEL, SHIP_856_IND EDI856, SOTSHIP1.SHIP_CART_REQD, NULL ORDR_NO" & vbCrLf _
                 & ", SHIP_ADDR_TYPE ORDR_ADDR_TYPE_ST, NULL CUST_CODE from SOTSHIP1" _
                 & " where SOTSHIP1.WHSE_CODE = '" & WHSE_CODE & "'"

            ASCMAIN1.sql &= " and nvl(SOTSHIP1.LP_STATUS, '0') = '1'" & vbCrLf
            grdSOTSHIPX.Text = "Shipments Sent to 3PL (" & WHSE_CODE & ")"

            UltraExplorerBar1.Groups("Screen Control").Items("Request 940 Cancel").Settings.Enabled = DefaultableBoolean.True
            UltraExplorerBar1.Groups("Screen Control").Items("Send").Settings.Enabled = DefaultableBoolean.False
        End If

        For Each colName As String In New String() {"LP_XNO"} ', "LP_XMIT_DATE"}
            grdSOTSHIPX.DisplayLayout.Bands(0).Columns(colName).Hidden = optPending.Value = "0"
        Next

        ASCDATA1.ExecuteSQL("Truncate Table " & SOTSHIPX)
        ASCDATA1.ExecuteSQL("Insert into " & SOTSHIPX & " " & ASCMAIN1.sql)

        'ASCDATA1.ExecuteSQL("Update " & SOTSHIPX & " SOTSHIPX " _
        '                    & "Set ORDR_NO = (Select Min (ORDR_NO) ORDR_NO from SOTPICK1 " _
        '                    & " where SHIP_BOL_NO = SOTSHIPX.SHIP_BOL_NO)")
        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS " _
            & " SELECT SOTPICK1.SHIP_BOL_NO, MIN(SOTPICK1.ORDR_NO) ORDR_NO" _
            & " FROM SOTPICK1, " & SOTSHIPX & " SOTSHIPX " _
            & " WHERE SOTPICK1.SHIP_BOL_NO = SOTSHIPX.SHIP_BOL_NO GROUP BY SOTPICK1.SHIP_BOL_NO;" _
            & " BEGIN FOR R1 IN C1 LOOP" _
            & " UPDATE " & SOTSHIPX & " SET ORDR_NO = R1.ORDR_NO WHERE SHIP_BOL_NO = R1.SHIP_BOL_NO;" _
            & " END LOOP; END; END;"
        'ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        'ASCDATA1.ExecuteSQL("Update " & SOTSHIPX & " SOTSHIPX " _
        '            & "Set CUST_CODE = (Select CUST_CODE from SOTORDR1 where ORDR_NO = SOTSHIPX.ORDR_NO)")

        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS " _
            & " SELECT SOTORDR1.ORDR_NO, SOTORDR1.CUST_CODE" _
            & " FROM SOTORDR1," & SOTSHIPX & " SOTSHIPX " _
            & " WHERE SOTORDR1.ORDR_NO = SOTSHIPX.ORDR_NO;" _
            & " BEGIN FOR R1 IN C1 LOOP" _
            & " UPDATE " & SOTSHIPX & " SET CUST_CODE = R1.CUST_CODE WHERE ORDR_NO = R1.ORDR_NO;" _
            & " END LOOP; END; END;"
        'ASCDATA1.ExecuteSQL(ASCMAIN1.sql)


        Fill_Records("SOTSHIPX")
        Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO".ToLower)

        Setup_SOTSHIPX()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Request_940_Cancel()

        ' ADS does not use 940's
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing to Cancel Shipments")

        Try
            Dim shipmentList As New List(Of String)

            For Each rowSOTSHIPX As DataRow In dst.Tables("SOTSHIPX").Select("SEL='1'")
                Dim SHIP_BOL_NO As String = rowSOTSHIPX.Item("SHIP_BOL_NO")
                Dim ORDR_GROUP_NO As String = rowSOTSHIPX.Item("ORDR_GROUP_NO")

                If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO) Then
                    Continue For
                End If

                Dim rowSOTSHIP1 As DataRow = ASCDATA1.GetDataRow("Select * from SOTSHIP1 where SHIP_BOL_NO = :PARM1", "V", SHIP_BOL_NO)
                If rowSOTSHIP1 Is Nothing Then
                    MessageBox.Show("Shipment (" & SHIP_BOL_NO & ") cannot be found.", "940 Cancel", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Continue For
                ElseIf rowSOTSHIP1.Item("SHIP_STATUS") & String.Empty <> "P" Then
                    MessageBox.Show("Shipment (" & SHIP_BOL_NO & ") is not 'In Pick'.", "940 Cancel", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Continue For
                End If

                If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then
                    Continue For
                End If

                shipmentList.Add(SHIP_BOL_NO)
            Next

            Try
                BeginTrans()

                Dim listOfShipments As String = String.Join("', '", shipmentList.ToArray)
                listOfShipments = "'" & listOfShipments & "'"

                ASCMAIN1.sql = "Update SOTSHIP1 Set LP_STATUS = '0', LP_XNO = NULL, SHIP_PICK_PRINTED = NULL where SHIP_BOL_NO in (" & listOfShipments & ")"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                CommitTrans("Request Successful")

            Catch ex As Exception
                Rollback(ex.Message)
            End Try

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Cancel 940", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
            ASCMAIN1.MultiTask_Release()
        End Try

        '*******************************************************************************************************************
        '*******************************************************************************************************************
        '*******************************************************************************************************************
        '*******************************************************************************************************************
        '*******************************************************************************************************************


        If 1 = 1 Then
            Exit Sub
        End If

        BeginTrans()

        If Not dst.Tables.Contains("EDT940O1") Then
            Create_TDA(dst.Tables.Add, "EDT940O1", "*")
            Create_TDA(dst.Tables.Add, "EDT940O2", "*")
            Create_TDA(dst.Tables.Add, "EDT940O4", "*")
            Create_TDA(dst.Tables.Add, "EDT940O5", "*")
        Else
            For Each tableName As String In New String() {"EDT940O1", "EDT940O2", "EDT940O4", "EDT940O"}
                dst.Tables(tableName).Clear()
            Next
        End If

        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
        Dim rowEDTTRPM1 As DataRow = LookUp("EDTTRPM1",
                                            New String() {rowICTWHSE1.Item("WHSE_EDI_QUAL"), rowICTWHSE1.Item("WHSE_EDI_ID"), "943"})


        For Each rowSOTSHIPX As DataRow In dst.Tables("SOTSHIPX").Select("SEL='1'")
            Dim CUST_CODE As String = rowSOTSHIPX.Item("CUST_CODE")
            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
            Dim rowEDTSLSP1 As DataRow = LookUp("EDTSLSP1", CUST_CODE)
            Dim SHIP_BOL_NO As String = rowSOTSHIPX.Item("SHIP_BOL_NO")
            Dim SHIP_VIA_CODE As String = rowSOTSHIPX.Item("SHIP_VIA_CODE") & ""
            Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", SHIP_VIA_CODE, True)

            ASCMAIN1.sql = "Update SOTSHIP1 Set LP_STATUS = '0', LP_XNO = NULL, SHIP_PICK_PRINTED = NULL" _
               & " where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCMAIN1.sql = "Select SOTPICK1.* from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and PICK_STATUS = 'P'"
            For Each rowSOTPICK1 As DataRow In ASCDATA1.GetDataTable.Select("")

                Dim EDI_OUTBOUND_DOC_NO As String = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")
                Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                'Dim EDI_DOC_SEQ_NO As String = rowSOTORDR1.Item("EDI_DOC_SEQ_NO") & ""
                'Dim rowEDT850T1 As DataRow = LookUp("EDT850T1", EDI_DOC_SEQ_NO)

                Dim CUST_STORE_NO As String = rowSOTORDR1.Item("CUST_STORE_NO") & ""
                Dim CUST_DC_NO As String = rowSOTORDR1.Item("CUST_DC_NO") & ""

                Dim rowEDT940O1 As DataRow = dst.Tables("EDT940O1").NewRow
                With rowEDT940O1
                    .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                    .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    .Item("CUST_CODE") = CUST_CODE
                    .Item("CUST_STORE_NO") = CUST_STORE_NO
                    .Item("PICK_NO") = PICK_NO
                    .Item("SHIP_BOL_NO") = SHIP_BOL_NO
                    .Item("ORDR_CUST_PO") = rowSOTORDR1.Item("ORDR_CUST_PO")
                    .Item("EDI_SUPPLIER_NO") = rowARTCUST1.Item("CUST_VEND_REF")
                    .Item("ORDR_DEPT") = rowSOTORDR1.Item("ORDR_DEPT")
                    .Item("ORDR_SHIP_DATE") = rowSOTORDR1.Item("ORDR_SHIP_DATE")
                    .Item("ORDR_CANCEL_DATE") = rowSOTORDR1.Item("ORDR_CANCEL_DATE")
                    .Item("ORDR_PO_DATE") = rowSOTORDR1.Item("ORDR_DATE")

                    Dim FRT_TERMS As String = rowSOTSHIPX.Item("FRT_TERMS") & ""
                    Dim FRT_TERMS_EDI As String = ""
                    Select Case FRT_TERMS
                        Case "PPD", "PPA"
                            FRT_TERMS_EDI = "PP"
                        Case "COL"
                            FRT_TERMS_EDI = "CC"
                    End Select
                    .Item("FRT_TERMS") = FRT_TERMS_EDI

                    '.Item("EDI_TRANS_METH_CODE") = "?"
                    '.Item("EDI_SERVICE_LEVEL") = "?"
                    '.Item("EDI_TP_BILLING_ACCT") = "? ' IF FRT_TERMS WAS 3RD PARTY WE WOULD SEND THE 3RD PARTY ACCT NUMBER
                    .Item("EDI_SCAC_CODE") = "ROUT" ' rowSOTSVIA1.Item("SHIP_VIA_SCAC")

                    Dim EDI_DIVISION_CODE As String = rowICTWHSE1.Item("LP_WHSE_ID") & ""

                    ASCMAIN1.sql = "Select * from EDT940O1 where PICK_NO = '" & PICK_NO & "'"
                    Dim rowEDT940O1_prior As DataRow = ASCDATA1.GetDataRow
                    If rowEDT940O1_prior IsNot Nothing Then
                        EDI_DIVISION_CODE = rowEDT940O1_prior.Item("EDI_DIVISION_CODE") & ""
                    End If
                    .Item("EDI_DIVISION_CODE") = EDI_DIVISION_CODE

                    ' .Item("EDI_LABEL_FORMAT") = rowARTCUST1.Item("CUST_VEND_REF")
                    .Item("EDI_LABEL_FORMAT") = rowARTCUST1.Item("LABEL_TEMPLATE_CODE")
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("ORDR_TYPE_CODE") = rowSOTORDR1.Item("ORDR_TYPE_CODE")
                    'If rowEDT850T1 IsNot Nothing Then
                    '    .Item("EDI_MERCH_TYPE") = rowEDT850T1.Item("EDI_MERCH_TYPE")
                    'End If
                    .Item("EDI_MERCH_TYPE") = rowSOTORDR1.Item("EDI_MERCH_TYPE")
                    .Item("ORDR_STATUS_CODE") = "V"
                End With
                dst.Tables("EDT940O1").Rows.Add(rowEDT940O1)

                ASCMAIN1.sql = "Insert into EDTSYSIH (COMPANY_CODE,EDI_OUTBOUND_DOC_NO,EDI_APPLICATION_ID,EDI_PROCESS_IND," _
                    & "EDI_OUR_ID,EDI_TP_ID,INIT_DATE,INIT_OPER)" _
                    & " Values (:PARM1,:PARM2,:PARM3,:PARM4,:PARM5,:PARM6,SYSDATE,'" & ASCMAIN1.USER_ID & "')"
                Dim EDI_APPLICATION_ID As String = "OW"
                Dim EDI_PROCESS_IND As String = "1"
                ' EDI_PROCESS_IND = "T"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVVVV",
                        New Object() {ASCMAIN1.DBS_COMPANY, EDI_OUTBOUND_DOC_NO, EDI_APPLICATION_ID, EDI_PROCESS_IND,
                                      rowEDTTRPM1.Item("EDI_OUR_ID"), rowICTWHSE1.Item("WHSE_EDI_ID")})
            Next
        Next

        Update_Record_TDA("EDT940O1")
        Update_Record_TDA("EDT940O2")
        Update_Record_TDA("EDT940O4")
        Update_Record_TDA("EDT940O5")

        CommitTrans("Update Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Function SendShipmentsTo3PL(ByVal WHSE_CODE As String) As Boolean

        Try
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Preparing to Transfer Shipments")

            Dim clsSOCADSO1 As New TAC.SOCADSO1
            Dim listOfShipments As New List(Of String)
            Dim shipmentsSentTo3pl As String = String.Empty

            For Each rowSOTSHIPX As DataRow In dst.Tables("SOTSHIPX").Select("SEL='1'")
                listOfShipments.Add(rowSOTSHIPX.Item("SHIP_BOL_NO"))
            Next

            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
            If rowICTWHSE1 Is Nothing Then
                MessageBox.Show($"Unknown Warehouse {WHSE_CODE}", "Send Shipments To 3PL", MessageBoxButtons.OK, MessageBoxIcon.Error)
                SendShipmentsTo3PL = False
                Exit Function
            End If

            Dim LP_CODE As String = rowICTWHSE1.Item("LP_CODE") & String.Empty

            dst.Tables("TASKS").Rows.Clear()
            clsSOCADSO1.tblTasks = Nothing
            clsSOCADSO1.tblTasks = dst.Tables("TASKS")
            Sort_grdColumns(grdTasks, "SEQ_NO")

            Select Case LP_CODE

                Case "ADS"
                    AddTask("WHFSPCK1 Call Prepare ADS Sales Orders File")

                    If Not clsSOCADSO1.PrepareADSSalesOrdersFile(listOfShipments, WHSE_CODE) Then
                        MessageBox.Show("The following occured when releasing to Warehouse (" & WHSE_CODE & "): " & vbCr & clsSOCADSO1.LastError, "Process Shipments", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        SendShipmentsTo3PL = False
                        Exit Function
                    End If

                    shipmentsSentTo3pl = "'" & String.Join("', '", listOfShipments.ToArray) & "'"
                    SendShipmentsTo3PL = True

                Case "CLA"
                    AddTask("WHFSPCK1 Call Prepare Clarins File")
                    SendShipmentsTo3PL = clsSOCADSO1.Prepare_Clarins_File(listOfShipments)

                    If SendShipmentsTo3PL = False Then
                        MessageBox.Show("The following occured when releasing to Warehouse (" & WHSE_CODE & "): " & vbCr & clsSOCADSO1.LastError, "Process Shipments", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Function
                    End If

                    shipmentsSentTo3pl = "'" & String.Join("', '", listOfShipments.ToArray) & "'"
                    If clsSOCADSO1.LastError.Length > 0 Then
                        Dim zmsg As String = "Some Shipments were successfuly sent to the 3PL." & Environment.NewLine & Environment.NewLine
                        zmsg &= "However, the following occured when releasing to Warehouse (" & WHSE_CODE & "): " & vbCr & clsSOCADSO1.LastError
                        MessageBox.Show(zmsg, "Process Shipments", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If

                Case Else
                    MessageBox.Show($"Invalid Warehouse {WHSE_CODE}, LP Code {LP_CODE} - No Shipments processed", "Process Shipments", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    SendShipmentsTo3PL = False
                    Exit Function
            End Select

            Try
                Select Case LP_CODE

                    Case "ADS"
                        LP_XNO = clsSOCADSO1.ReleaseShipmentNo

                        BeginTrans()

                        AddTask("WHFSPCK1 Update SOTSHIP1")

                        ASCMAIN1.sql = "Update SOTSHIP1 Set LP_STATUS = '1', LP_XNO = '" & LP_XNO & "', LP_XMIT_DATE = SYSDATE"
                        ASCMAIN1.sql &= ", SHIP_PICK_PRINTED = SYSDATE"
                        ASCMAIN1.sql &= " where SHIP_BOL_NO in (" & shipmentsSentTo3pl & ")"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                        AddTask("WHFSPCK1 Update SOTPICK1")

                        ASCMAIN1.sql = "Update SOTPICK1 Set PICK_PRINTED = SYSDATE, ORDR_NO_3PL = PICK_NO"
                        ASCMAIN1.sql &= ", PICK_PRINTED_OPER = '" & ASCMAIN1.USER_ID & "'"
                        ASCMAIN1.sql &= " where SHIP_BOL_NO in (" & shipmentsSentTo3pl & ")"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                        dst.Tables("WHTSHIPX").Rows.Clear()
                        dst.Tables("WHTLPXN1").Rows.Clear()

                        For Each shipBol As String In shipmentsSentTo3pl.Split(",")
                            dst.Tables("WHTSHIPX").Rows.Add(New Object() {LP_XNO, shipBol.Replace("'", "").Trim})
                        Next

                        AddTask("WHFSPCK1 Update WHTSHIPX")
                        Update_Record_TDA("WHTSHIPX")

                        Dim rowWHTLPXN1 As DataRow = dst.Tables("WHTLPXN1").NewRow
                        rowWHTLPXN1.Item("LP_XNO") = LP_XNO
                        rowWHTLPXN1.Item("LP_XNO_SOURCE") = MENU_ITEM_OBJECT
                        rowWHTLPXN1.Item("LP_XNO_RECORDS") = dst.Tables("WHTSHIPX").Rows.Count
                        rowWHTLPXN1.Item("LP_XNO_NOTES") = ""
                        rowWHTLPXN1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                        rowWHTLPXN1.Item("INIT_DATE") = DateTime.Now
                        dst.Tables("WHTLPXN1").Rows.Add(rowWHTLPXN1)

                        AddTask("WHFSPCK1 Update WHTLPXN1")
                        Update_Record_TDA("WHTLPXN1")

                        AddTask("WHFSPCK1 CommitTrans")
                        CommitTrans()
                        SendShipmentsTo3PL = True
                End Select

                Try
                    ' Show Address where the goods are shipping
                    Dim specialDeliveries As String = String.Empty
                    Dim rowSOTORDR5 As DataRow = Nothing
                    Dim addressLine As String = String.Empty
                    ',CUST_ADDR1,CUST_ADDR2,CUST_ADDR3,CUST_CITY,CUST_STATE,CUST_ZIP_CODE,CUST_COUNTRY
                    Dim sql As String = "Select * from SOTORDR5 where CUST_ADDR_TYPE = 'ST' and ORDR_NO = (SELECT ORDR_NO from SOTPICK1 where PICK_NO = :PARM1)"

                    AddTask("WHFSPCK1 GetPriorityDeliveries")
                    If clsSOCADSO1.GetPriorityDeliveries.Count > 0 Then
                        specialDeliveries &= Environment.NewLine & "The following shipments are PRIORITY Deliveries"
                        For Each priorityDelivery As String In clsSOCADSO1.GetPriorityDeliveries
                            priorityDelivery = priorityDelivery.Trim
                            If priorityDelivery.Length > 0 Then
                                rowSOTORDR5 = ASCDATA1.GetDataRow(sql, "V", priorityDelivery)
                                addressLine = String.Empty
                                If rowSOTORDR5 IsNot Nothing Then
                                    addressLine = Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item("CUST_NAME") & String.Empty
                                    For Each field As String In New String() {"CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3"}
                                        If rowSOTORDR5.Item(field) & String.Empty <> String.Empty Then
                                            addressLine &= Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item(field) & String.Empty
                                        End If
                                    Next
                                    addressLine &= Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item("CUST_CITY") & ", " & rowSOTORDR5.Item("CUST_STATE") & "  " & rowSOTORDR5.Item("CUST_ZIP_CODE")
                                End If
                                specialDeliveries &= Environment.NewLine & vbTab & vbTab & priorityDelivery & addressLine & Environment.NewLine
                            End If
                        Next

                    End If

                    AddTask("WHFSPCK1 GetNextDayDeliveries")
                    If clsSOCADSO1.GetNextDayDeliveries.Count > 0 Then
                        specialDeliveries &= Environment.NewLine & "The following shipments are Next Day Deliveries"
                        For Each nextDayDelivery As String In clsSOCADSO1.GetNextDayDeliveries
                            nextDayDelivery = nextDayDelivery.Trim
                            If nextDayDelivery.Length > 0 Then
                                rowSOTORDR5 = ASCDATA1.GetDataRow(sql, "V", nextDayDelivery)
                                addressLine = String.Empty
                                If rowSOTORDR5 IsNot Nothing Then
                                    addressLine = Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item("CUST_NAME") & String.Empty
                                    For Each field As String In New String() {"CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3"}
                                        If rowSOTORDR5.Item(field) & String.Empty <> String.Empty Then
                                            addressLine &= Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item(field) & String.Empty
                                        End If
                                    Next
                                    addressLine &= Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item("CUST_CITY") & ", " & rowSOTORDR5.Item("CUST_STATE") & "  " & rowSOTORDR5.Item("CUST_ZIP_CODE")
                                End If
                                specialDeliveries &= Environment.NewLine & vbTab & vbTab & nextDayDelivery & addressLine & Environment.NewLine
                            End If
                        Next
                    End If

                    AddTask("WHFSPCK1 GetSecondDayDeliveries")
                    If clsSOCADSO1.GetSecondDayDeliveries.Count > 0 Then
                        specialDeliveries &= Environment.NewLine & "The following shipments are 2nd Day Deliveries"
                        For Each secondDayDelivery As String In clsSOCADSO1.GetSecondDayDeliveries
                            secondDayDelivery = secondDayDelivery.Trim
                            If secondDayDelivery.Length > 0 Then
                                rowSOTORDR5 = ASCDATA1.GetDataRow(sql, "V", secondDayDelivery)
                                addressLine = String.Empty
                                If rowSOTORDR5 IsNot Nothing Then
                                    addressLine = Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item("CUST_NAME") & String.Empty
                                    For Each field As String In New String() {"CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3"}
                                        If rowSOTORDR5.Item(field) & String.Empty <> String.Empty Then
                                            addressLine &= Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item(field) & String.Empty
                                        End If
                                    Next
                                    addressLine &= Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item("CUST_CITY") & ", " & rowSOTORDR5.Item("CUST_STATE") & "  " & rowSOTORDR5.Item("CUST_ZIP_CODE")
                                End If
                                specialDeliveries &= Environment.NewLine & vbTab & vbTab & secondDayDelivery & addressLine & Environment.NewLine
                            End If
                        Next
                    End If

                    AddTask("WHFSPCK1 GetPosDeliveries")
                    If clsSOCADSO1.GetPosDeliveries.Count > 0 Then
                        specialDeliveries &= Environment.NewLine & "The following shipments are POS Deliveries"
                        For Each PosDelivery As String In clsSOCADSO1.GetPosDeliveries
                            PosDelivery = PosDelivery.Trim
                            If PosDelivery.Length > 0 Then
                                rowSOTORDR5 = ASCDATA1.GetDataRow(sql, "V", PosDelivery)
                                addressLine = String.Empty
                                If rowSOTORDR5 IsNot Nothing Then
                                    addressLine = Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item("CUST_NAME") & String.Empty
                                    For Each field As String In New String() {"CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3"}
                                        If rowSOTORDR5.Item(field) & String.Empty <> String.Empty Then
                                            addressLine &= Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item(field) & String.Empty
                                        End If
                                    Next
                                    addressLine &= Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item("CUST_CITY") & ", " & rowSOTORDR5.Item("CUST_STATE") & "  " & rowSOTORDR5.Item("CUST_ZIP_CODE")
                                End If
                                specialDeliveries &= Environment.NewLine & vbTab & vbTab & PosDelivery & addressLine & Environment.NewLine
                            End If
                        Next
                    End If

                    Dim pickData As String = String.Empty
                    sql = " SELECT SOTPICK1.PICK_NO, SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.ORDR_CUST_PO, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_ADDR_TYPE_ST"
                    sql &= " FROM SOTORDR1, SOTPICK1"
                    sql &= " WHERE SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO "
                    sql &= " AND SOTPICK1.SHIP_BOL_NO IN (" & shipmentsSentTo3pl & ")"

                    pickData = "<TABLE BORDER=""0"" CELLPADDING =""5"" CELLSPACING=""5"" WIDTH=""710"">" & Environment.NewLine

                    pickData &= " <TR>"
                    pickData &= "   <TH style=""text-align:left"" ""font-weight:bold"" width=""100"">Pick Ticket Detail:</TH>"
                    pickData &= " </TR>" & Environment.NewLine

                    pickData &= " <TR>"
                    pickData &= "   <TH style=""text-align:left"" width=""100"">Pick No</TH>"
                    pickData &= "   <TH style=""text-align:left"" width=""150"">Customer</TH>"
                    pickData &= "   <TH style=""text-align:left"" width=""300"">Customer Name</TH>"
                    pickData &= "   <TH style=""text-align:left"" width=""200"">PO</TH>"
                    pickData &= "   <TH style=""text-align:left"" width=""10""></TH>"
                    pickData &= " </TR>" & Environment.NewLine

                    AddTask("WHFSPCK1 Create Email Pick Information")
                    For Each row As DataRow In ASCDATA1.GetDataTable(sql).Select("", "PICK_NO")
                        Dim CUST_CODE As String = row.Item("CUST_CODE")
                        Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")

                        Dim XREF_CUST_CODE_SHIP_TO As String = row.Item("CUST_CODE")
                        Dim XREF_CUST_STORE_NO_SHIP_TO As String = String.Empty

                        If WHSE_CODE = "CLA" Then
                            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                            Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})

                            If rowARTCUST1.Item("CUST_SHIP_TO_MANUAL") & String.Empty = "1" Then
                                XREF_CUST_CODE_SHIP_TO = rowARTCUST1.Item("CUST_NO_3PL") & "/" & rowARTCUST1.Item("CUST_STORE_NO_3PL") & String.Empty
                            Else
                                XREF_CUST_CODE_SHIP_TO = rowARTCUST2.Item("CUST_NO_3PL") & String.Empty
                                XREF_CUST_STORE_NO_SHIP_TO = rowARTCUST2.Item("CUST_STORE_NO_3PL") & String.Empty

                                Select Case row.Item("ORDR_ADDR_TYPE_ST") & String.Empty
                                    Case "DC"
                                        Dim CUST_DC_NO As String = rowARTCUST2.Item("CUST_DC_NO") & String.Empty
                                        rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_DC_NO})

                                        ' Set to Blank so if the DC record is missing the Clarins Settings an error is generated
                                        XREF_CUST_CODE_SHIP_TO = String.Empty
                                        XREF_CUST_STORE_NO_SHIP_TO = String.Empty

                                        If rowARTCUST2 IsNot Nothing Then
                                            XREF_CUST_CODE_SHIP_TO = rowARTCUST2.Item("CUST_NO_3PL") & String.Empty
                                            XREF_CUST_STORE_NO_SHIP_TO = rowARTCUST2.Item("CUST_STORE_NO_3PL") & String.Empty
                                        End If
                                End Select
                            End If
                            XREF_CUST_CODE_SHIP_TO &= "/" & XREF_CUST_STORE_NO_SHIP_TO
                        Else
                            XREF_CUST_CODE_SHIP_TO = CUST_CODE & "/" & CUST_STORE_NO
                        End If

                        pickData &= " <TR>"
                        pickData &= "   <TD>" & row.Item("PICK_NO") & "</TD>"
                        pickData &= "   <TD>" & XREF_CUST_CODE_SHIP_TO & "</TD>"
                        pickData &= "   <TD>" & row.Item("CUST_NAME") & "</TD>"
                        pickData &= "   <TD>" & row.Item("ORDR_CUST_PO") & "</TD>"
                        pickData &= "   <TD>" & "" & "</TD>"
                        pickData &= " </TR>" & Environment.NewLine
                    Next
                    pickData &= "</TABLE>"

                    AddTask("WHFSPCK1 Select MAX(LP_XNO) from WHTSHIPX")
                    ASCMAIN1.sql = "Select MAX(LP_XNO) from WHTSHIPX WHERE SHIP_BOL_NO IN (" & shipmentsSentTo3pl & ")"
                    LP_XNO = ASCDATA1.GetDataValue(ASCMAIN1.sql) & String.Empty

                    If LP_XNO.Length > 0 Then
                        AddTask("WHFSPCK1 Start EmailIPLBShipments")
                        EmailIPLBShipments(LP_XNO, WHSE_CODE, listOfShipments.Count, specialDeliveries, pickData)
                        AddTask("WHFSPCK1 End EmailIPLBShipments")
                    End If

                Catch ex As Exception
                    MessageBox.Show("The following error occurred: " & ex.Message, "Process Shipments", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Catch ex As Exception
                Rollback(ex.Message)
                SendShipmentsTo3PL = False
            End Try

            MessageBox.Show("Shipments transferred successfully.", "Process Shipments", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("The following error occurred: " & ex.Message, "Process Shipments", MessageBoxButtons.OK, MessageBoxIcon.Error)
            SendShipmentsTo3PL = False

        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("", "")
            AddTask("WHFSPCK1 Process Complete")
            Sort_grdColumns(grdTasks, "SEQ_NO")
        End Try

    End Function

    Private Sub EmailIPLBShipments(ByVal LP_XNO As String,
                                  ByVal WHSE_CODE As String,
                                  ByVal NumOfShipments As Int32,
                                  ByVal specialDeliveries As String,
                                  ByVal pickData As String)

        If LP_XNO.Length > 0 Then
            Dim emailTransferred As Boolean = False

            While Not emailTransferred
                Dim objASCNOTEE As New TAC.ASCNOTEE(ASCMAIN1.Folders, "SHIP_" & WHSE_CODE, dst)
                objASCNOTEE.Note = $"Batch Number {LP_XNO} uploaded from IPLB contains {NumOfShipments} shipments."
                objASCNOTEE.Note &= Environment.NewLine & Environment.NewLine & specialDeliveries & Environment.NewLine & Environment.NewLine & pickData
                objASCNOTEE.CreateComponents()
                objASCNOTEE.EmailDocument()

                If objASCNOTEE.lastError.Length = 0 Then
                    Exit While
                End If

                MessageBox.Show($"The following error occurred when generating the Shipments email: {objASCNOTEE.lastError}", "Process Shipments", MessageBoxButtons.OK, MessageBoxIcon.Error)

                If MessageBox.Show("Do you want to try and resend the email? If you choose 'Yes' the system will wait 3 seconds then retry.", "Process Shipments", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    emailTransferred = True
                Else
                    System.Threading.Thread.Sleep(3000)
                End If
            End While
        End If

    End Sub


#End Region

#Region "CARS"

    Private Sub btnCARS_Click(sender As Object, e As EventArgs) Handles btnCARS.Click

        Dim DATE_LAST_EMAIL As Date = Nothing

        If chkInitialize.Checked Then
            ASCMAIN1.sql = "UPDATE SOTSHIP1 SET SHIP_ROUTING_IND = '0' WHERE SHIP_STATUS = 'P' AND SHIP_ROUTING_REQD = '1'"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "TRUNCATE TABLE SOTSHIPR"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "TRUNCATE TABLE SOTSHIPV"
            ASCDATA1.ExecuteSQL()

            chkInitialize.Checked = False
        End If


        Dim CHECK_EMAIL_ONCE_AFTER_END As Boolean = False
        Dim SHIP_ROUTING_XNO As String = ""

        Dim DATE_LAST_NWD_LOAD As Date = Nothing
        Dim DATE_LAST_PARM_LOAD As Date = Nothing
        Dim rowSOTPARMH As DataRow = Nothing

        Do
            ' Parameters

            Dim rowSOTPARMR As DataRow
            Dim SO_PARM_CALC_PAUSE_SECS As Int32
            Dim SO_PARM_EMAIL_FREQ As Int32
            Dim SO_PARM_CARS_START As Int32
            Dim SO_PARM_CARS_END As Int32
            Dim SO_PARM_RUN_CARS_TODAY As Date
            Dim SO_PARM_DEF_LTL_CARRIER As String = ""
            Dim SO_PARM_REFRESH_MINS As Int32 = 2
            Dim SO_PARM_NOM_NET_WGT_PER_CTN As Decimal = 9.25
            Dim SO_PARM_TARE_WGT_PER_CTN As Decimal = 0.75

            If DATE_LAST_PARM_LOAD.Year < 2000 Then
                DATE_LAST_PARM_LOAD = Now.AddDays(-1)
            End If
            If Now.Subtract(DATE_LAST_PARM_LOAD).TotalMinutes > SO_PARM_REFRESH_MINS Or Format(Now, "yyyyMMdd") <> Format(DATE_LAST_PARM_LOAD, "yyyyMMdd") Then
                rowSOTPARMR = LookUp("SOTPARMR", "Z")
                DATE_LAST_PARM_LOAD = Now

                SO_PARM_CALC_PAUSE_SECS = Val(rowSOTPARMR.Item("SO_PARM_CALC_PAUSE_SECS") & "")
                SO_PARM_EMAIL_FREQ = Val(rowSOTPARMR.Item("SO_PARM_EMAIL_FREQ") & "")
                SO_PARM_CARS_START = Val(rowSOTPARMR.Item("SO_PARM_CARS_START") & "")
                SO_PARM_CARS_END = Val(rowSOTPARMR.Item("SO_PARM_CARS_END") & "")
                SO_PARM_RUN_CARS_TODAY = IIf(rowSOTPARMR.Item("SO_PARM_RUN_CARS_TODAY") & "" = "", Nothing, rowSOTPARMR.Item("SO_PARM_RUN_CARS_TODAY") & "")
                SO_PARM_DEF_LTL_CARRIER = rowSOTPARMR.Item("SO_PARM_DEF_LTL_CARRIER") & ""
                SO_PARM_REFRESH_MINS = Val(rowSOTPARMR.Item("SO_PARM_REFRESH_MINS") & "")
                SO_PARM_NOM_NET_WGT_PER_CTN = Val(rowSOTPARMR.Item("SO_PARM_NOM_NET_WGT_PER_CTN") & "")
                SO_PARM_TARE_WGT_PER_CTN = Val(rowSOTPARMR.Item("SO_PARM_TARE_WGT_PER_CTN") & "")

                If SO_PARM_DEF_LTL_CARRIER = "" Then SO_PARM_DEF_LTL_CARRIER = "DEFAULT"
                If SO_PARM_CARS_START <= 0 Or SO_PARM_CARS_START > 23 Then SO_PARM_CARS_START = 0
                If SO_PARM_CARS_END <= 0 Or SO_PARM_CARS_END > 23 Then SO_PARM_CARS_START = 15
                If SO_PARM_CARS_START > SO_PARM_CARS_END Then SO_PARM_CARS_START = SO_PARM_CARS_END

                If SO_PARM_NOM_NET_WGT_PER_CTN <= 0 Then SO_PARM_NOM_NET_WGT_PER_CTN = 9.25
                If SO_PARM_TARE_WGT_PER_CTN <= 0 Then SO_PARM_TARE_WGT_PER_CTN = 0.75

                ' *******************************************************
                If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" Then SO_PARM_CARS_END = 23

                DATE_LAST_NWD_LOAD = Nothing
            End If


            If DATE_LAST_NWD_LOAD.Year < 2000 Then
                DATE_LAST_NWD_LOAD = Now.AddDays(-1)
            End If

            If Now.Subtract(DATE_LAST_NWD_LOAD).TotalMinutes > SO_PARM_REFRESH_MINS Or Format(Now, "yyyyMMdd") <> Format(DATE_LAST_NWD_LOAD, "yyyyMMdd") Then
                Fill_Records("SOTPARMH")
                DATE_LAST_NWD_LOAD = Now

                If Format(SO_PARM_RUN_CARS_TODAY, "yyyyMMdd") = Format(Now, "yyyyMMdd") Then
                    rowSOTPARMH = Nothing
                Else
                    rowSOTPARMH = dst.Tables("SOTPARMH").Rows.Find(Now.Date)
                End If

                Fill_Records("SOTSHIPW")
                Fill_Records("SOTSHIPS")
                Fill_Records("SOTSHIPU")
                Fill_Records("SOTSHIPF")
                Fill_Records("SOTSHIPQ")
                Fill_Records("SOTSHIPZ")

            End If

            Dim WORKING_DAY As Boolean = True
            If (Now.DayOfWeek = Global.System.DayOfWeek.Saturday Or Now.DayOfWeek = Global.System.DayOfWeek.Sunday) And Format(SO_PARM_RUN_CARS_TODAY, "yyyyMMdd") <> Format(Now, "yyyyMMdd") Then WORKING_DAY = False
            If rowSOTPARMH IsNot Nothing Then WORKING_DAY = False

            ' *******************************************************
            If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" Then WORKING_DAY = True

            Do While (WORKING_DAY _
                    Or (SO_PARM_RUN_CARS_TODAY.Year > 2000 And Format(SO_PARM_RUN_CARS_TODAY, "yyyyMMdd") = Format(Now, "yyyyMMdd"))) _
                AndAlso (Now.Hour >= SO_PARM_CARS_START And Now.Hour <= SO_PARM_CARS_END)
                CHECK_EMAIL_ONCE_AFTER_END = True

                Dim OPS_LEAD_TIME_DAYS As Int32 = 0
                SHIP_ROUTING_XNO = ASCMAIN1.Next_Control_No("SOTSHIPR.SHIP_ROUTING_XNO")

                Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Now starting To process shipments {SHIP_ROUTING_XNO}")

                Fill_Records("SOTSHIP1_CARS")
                Dim RECORDS As Int32 = dst.Tables("SOTSHIP1_CARS").Rows.Count
                If RECORDS = 0 Then
                    Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"No Shipments To Process - sleeping For {CStr(SO_PARM_CALC_PAUSE_SECS)} seconds")
                    Exit Do
                Else
                    Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Now Processing {RECORDS} Shipments", EVENT_KEY_COUNT:=RECORDS)
                End If

                For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1_CARS").Select("", "SHIP_BOL_NO")

                    Dim SHIP_BOL_NO As String = ""
                    Dim CUST_CODE As String = ""
                    Dim SHIP_ADDR_CODE As String = ""
                    Dim ORDR_SHIP_DATE As Date
                    Dim ORDR_CANCEL_DATE As Date
                    Dim SHIP_PRE_AUTH_REQD As String = ""
                    Dim SHIP_TO_CODE As String = ""

                    Try
                        DATETIME_STAMP = Now + ASCMAIN1.NowTSD

                        SHIP_BOL_NO = rowSOTSHIP1.Item("SHIP_BOL_NO")
                        CUST_CODE = rowSOTSHIP1.Item("CUST_CODE") & ""
                        SHIP_ADDR_CODE = rowSOTSHIP1.Item("SHIP_ADDR_CODE") & ""

                        ORDR_SHIP_DATE = rowSOTSHIP1.Item("SHIP_SHIP_DATE")
                        ORDR_CANCEL_DATE = rowSOTSHIP1.Item("SHIP_CANCEL_DATE")

                        ' If SHIP or CANCEL falls on a NWD, do we need to adjust them?
                        ' From BH
                        '	if ship = sat Or sun move to next mon (for calc purposes)
                        '	if canc = sat Or sun move to prev fri (for calc purposes)

                        Dim dSHIP As Date = ORDR_SHIP_DATE
                        If dSHIP.DayOfWeek = Global.System.DayOfWeek.Saturday Or dSHIP.DayOfWeek = Global.System.DayOfWeek.Sunday Then
                            dSHIP = Calc_WD("*", 0, dSHIP, 0, 1)
                        End If
                        'If dSHIP.DayOfWeek = Global.System.DayOfWeek.Saturday Then
                        '    dSHIP = dSHIP.AddDays(2) ' prob need to check for holiday
                        'ElseIf dSHIP.DayOfWeek = Global.System.DayOfWeek.Sunday Then
                        '    dSHIP = dSHIP.AddDays(1) ' prob need to check for holiday
                        'End If
                        Dim dCANC As Date = ORDR_CANCEL_DATE
                        If dCANC.DayOfWeek = Global.System.DayOfWeek.Saturday Or dCANC.DayOfWeek = Global.System.DayOfWeek.Sunday Then
                            dCANC = Calc_WD("*", 0, dCANC, 0, -1)
                        End If
                        'If dCANC.DayOfWeek = Global.System.DayOfWeek.Saturday Then
                        '    dCANC = dCANC.AddDays(-1) ' prob need to check for holiday
                        'ElseIf dCANC.DayOfWeek = Global.System.DayOfWeek.Sunday Then
                        '    dCANC = dCANC.AddDays(-2) ' prob need to check for holiday
                        'End If
                        Dim daysShipWindow As Int32 = dCANC.Subtract(dSHIP).TotalDays + 1 ' days in Ship Window

                        SHIP_PRE_AUTH_REQD = rowSOTSHIP1.Item("SHIP_PRE_AUTH_REQD") & ""

                        If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" Then
                            'If SHIP_BOL_NO = "0000438023" Then Stop
                        End If

                        Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Now Processing Shipment", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE)

                        Dim rowSOTSHIPR As DataRow = Fill_Record("SOTSHIPR", SHIP_BOL_NO)
                        If rowSOTSHIPR Is Nothing Then
                            rowSOTSHIPR = dst.Tables("SOTSHIPR").NewRow
                            With rowSOTSHIPR
                                .Item("SHIP_BOL_NO") = SHIP_BOL_NO
                                .Item("ORDR_SHIP_DATE") = ORDR_SHIP_DATE
                                .Item("ORDR_CANCEL_DATE") = ORDR_CANCEL_DATE
                                .Item("SHIP_TO_CODE") = SHIP_TO_CODE
                                .Item("INIT_DATE") = DATETIME_STAMP
                                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                                .Item("LAST_DATE") = DATETIME_STAMP
                                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                            End With

                            dst.Tables("SOTSHIPR").Rows.Add(rowSOTSHIPR)
                        End If

                        With rowSOTSHIPR
                            .Item("SHIP_ROUTING_ATTEMPTS") = Val(.Item("SHIP_ROUTING_ATTEMPTS") & "") + 1
                        End With

                        Update_Record_TDA("SOTSHIPR")

                        If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then
                            Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Unable To Lock Shipment", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="L")
                            Exit Try
                        End If

                        Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Lock Shipment Successful", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE)



                        ' Weight

                        Dim SHIP_WEIGHT As Decimal = 0
                        Dim SHIP_CARTONS As Int32 = 0

                        Fill_Records("SOTPICK1_CARS", SHIP_BOL_NO)
                        Dim PT_Count As Int32 = dst.Tables("SOTPICK1_CARS").Rows.Count
                        Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Now Processing {PT_Count} Pick Tickets For Shipment {SHIP_BOL_NO}", EVENT_KEY_COUNT:=PT_Count, SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE)

                        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1_CARS").Select("", "PICK_NO")
                            Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                            Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
                            If Not ASCMAIN1.Logical_Lock("SOTPICK1", PICK_NO) Or Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO) Then
                                Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Unable To Lock Order Or Pick Ticket", EVENT_KEY_TYPE:="P", EVENT_KEY:=PICK_NO, SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, EVENT_ERROR_TYPE:="L")
                                Exit Try
                            End If

                            Dim rowSOTPICK1_CARS_WGT As DataRow = Fill_Record("SOTPICK1_CARS_WGT", PICK_NO)
                            ' GET TOTAL CARTONS, WEIGHT, VOLUME *********************
                            Dim PICK_WGT As Decimal = Val(rowSOTPICK1_CARS_WGT.Item("PICK_WGT") & "")
                            If PICK_WGT = 0 Then
                                Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"No Weight For Pick Ticket {PICK_NO}", EVENT_KEY_TYPE:="P", EVENT_KEY:=PICK_NO, SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="W")
                                Exit Try
                            Else
                                Dim ITEM_WGT_MIN As Decimal = Val(rowSOTPICK1_CARS_WGT.Item("ITEM_WGT_MIN") & "")
                                If ITEM_WGT_MIN = 0 Then
                                    For Each rowSOTPICK1_CARS_WGT_ITEM As DataRow In dst.Tables("SOTPICK1_CARS_WGT").Select("ISNULL(ITEM_WGT,0) = 0")
                                        Dim ITEM_CODE As String = rowSOTPICK1_CARS_WGT_ITEM.Item("ITEM_CODE") & ""
                                        Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Item ({ITEM_CODE}) has No Weight on Pick Ticket {PICK_NO}", EVENT_KEY_TYPE:="I", EVENT_KEY:=ITEM_CODE, SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="W")
                                    Next
                                    Exit Try
                                End If
                            End If

                            SHIP_WEIGHT += PICK_WGT
                        Next

                        Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Calculated Weight for Shipment: {Format(SHIP_WEIGHT, "#,##0.00")} LBS", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE)

                        Dim CARTONS As Decimal = SHIP_WEIGHT / SO_PARM_NOM_NET_WGT_PER_CTN
                        Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Carton count estimated using {Format(SO_PARM_NOM_NET_WGT_PER_CTN, "#.00")} LBs per Carton: {Format(CARTONS, "#,##0.00")}", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE)
                        If CARTONS > Math.Floor(CARTONS) Then
                            CARTONS = Math.Floor(CARTONS) + 1
                            Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Carton count rounded up to next whole number: {Format(CARTONS, "#,##0.00")}", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE)
                        End If

                        SHIP_CARTONS = CARTONS

                        Dim TARE As Decimal = SHIP_CARTONS * SO_PARM_TARE_WGT_PER_CTN
                        SHIP_WEIGHT += TARE
                        Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Tare @{Format(SO_PARM_TARE_WGT_PER_CTN, "#.00")} LB per Carton: {Format(TARE, "#,##0.00")}, Adjusted Total Weight: {Format(SHIP_WEIGHT, "#,##0.00")}", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE)

                        If SHIP_WEIGHT > Math.Floor(SHIP_WEIGHT) Then
                            SHIP_WEIGHT = Math.Floor(SHIP_WEIGHT) + 1
                            Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Weight rounded up to next whole LB: {Format(SHIP_WEIGHT, "#,##0.00")} LBS", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE)
                        End If


                        ' Ship-to & Customer Master

                        Dim rowSOTSHIPD As DataRow = Get_SHIP_TO_Record(SHIP_ROUTING_XNO, SHIP_BOL_NO, CUST_CODE, SHIP_ADDR_CODE)
                        Dim rowSOTSHIPE As DataRow = dst.Tables("SOTSHIPE").Rows.Find(New String() {"C", CUST_CODE})
                        If rowSOTSHIPE Is Nothing Then
                            rowSOTSHIPE = Fill_Record("SOTSHIPE", New String() {"C", CUST_CODE})
                        End If

                        If rowSOTSHIPD Is Nothing Then
                            Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Unable To Find Ship-To Record", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="S")
                            Exit Try
                        ElseIf rowSOTSHIPE Is Nothing Then
                            Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Unable To Find Ship-To Prefix (Customer) Record", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="S")
                            Exit Try
                        End If

                        SHIP_TO_CODE = rowSOTSHIPD.Item("SHIP_TO_CODE")
                        Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Found Ship-To Code", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE)

                        ' Ship Via

                        Dim SHIP_VIA_CODE As String = rowSOTSHIP1.Item("SHIP_VIA_CODE") & ""
                        Dim CARRIER_ACCT_TYPE As String = ""

                        Dim FRT_TERMS As String = rowSOTSHIP1.Item("FRT_TERMS") & "" ' 3PY & COL mean "Use the Customer's SPS account"
                        If FRT_TERMS = "" Then
                            Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Shipment has no Freight Terms Specified", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="D")
                            Exit Try
                        ElseIf FRT_TERMS <> "PPD" And FRT_TERMS <> "PPA" And FRT_TERMS <> "COL" And FRT_TERMS <> "3PY" Then
                            Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Shipment has invalid Freight Terms ({FRT_TERMS})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="D")
                            Exit Try
                        End If

                        Dim DEL_METHOD As String = ""

                        Dim SHIP_CARRIER As String = ""
                        Dim SHIP_METHOD As String = ""

                        Dim SHIP_VIA As String = ""
                        Dim CARRIER As String = ""

                        Dim rowSOTSHIPS As DataRow = Nothing
                        If SHIP_VIA_CODE = "" Then
                            Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Shipment has no Ship Via Specified", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="D")
                            Exit Try

                        ElseIf SHIP_VIA_CODE <> "RGOF" Then
                            rowSOTSHIPS = dst.Tables("SOTSHIPS").Rows.Find(SHIP_VIA_CODE)
                            If rowSOTSHIPS Is Nothing Then
                                Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Shipment has invalid Ship Via ({SHIP_VIA_CODE})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="D")
                                Exit Try
                            Else
                                DEL_METHOD = rowSOTSHIPS.Item("DEL_METHOD") & ""
                                CARRIER_ACCT_TYPE = rowSOTSHIPS.Item("CARRIER_ACCT_TYPE") & ""
                                If DEL_METHOD <> "SPS" And DEL_METHOD <> "LTL" Then
                                    Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Invalid Delivery Method ({DEL_METHOD}) specified in provided Ship Via ({SHIP_VIA_CODE})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="V")
                                    Exit Try
                                End If

                                Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Ship Via resolved to value provided in Shipment-Order: ({SHIP_VIA_CODE})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE)
                                If CARRIER_ACCT_TYPE <> "" Then
                                    Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Carrier Type resolved to value provided in Ship Via: ({CARRIER_ACCT_TYPE})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE)
                                End If
                            End If

                        ElseIf SHIP_VIA_CODE = "RGOF" Then

                            CARRIER_ACCT_TYPE = rowSOTSHIPD.Item("CARRIER_ACCT_TYPE") & ""
                            SHIP_VIA = rowSOTSHIPD.Item("SHIP_VIA") & ""

                            If SHIP_VIA = "" Then
                                If rowSOTSHIPE.Item("SHIP_VIA") & "" <> "" Then
                                    SHIP_VIA = rowSOTSHIPE.Item("SHIP_VIA") & ""
                                    Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Ship Via resolved to value provided in Customer Master: ({SHIP_VIA_CODE})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE)
                                End If
                            Else
                                Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Ship Via resolved to value provided in Ship-To Master: ({SHIP_VIA_CODE})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE)
                            End If

                            If CARRIER_ACCT_TYPE = "" Then
                                If rowSOTSHIPE.Item("CARRIER_ACCT_TYPE") & "" <> "" Then
                                    CARRIER_ACCT_TYPE = rowSOTSHIPE.Item("CARRIER_ACCT_TYPE") & ""
                                    Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Carrier Type resolved to value provided in Customer Master: ({CARRIER_ACCT_TYPE})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE)
                                End If
                            Else
                                Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Carrier Type resolved to value provided in Ship-To Master: ({CARRIER_ACCT_TYPE})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE)
                            End If

                            If SHIP_VIA <> "" Then
                                rowSOTSHIPS = dst.Tables("SOTSHIPS").Rows.Find(SHIP_VIA)
                                If rowSOTSHIPS Is Nothing Then
                                    Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Invalid Ship Via ({SHIP_VIA}) coming from Customer or Ship-To", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="V")
                                    Exit Try
                                Else
                                    DEL_METHOD = rowSOTSHIPS.Item("DEL_METHOD") & ""
                                    If DEL_METHOD <> "SPS" And DEL_METHOD <> "LTL" Then
                                        Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Invalid Delivery Method ({DEL_METHOD}) specified in resolved Ship Via ({SHIP_VIA})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="V")
                                        Exit Try
                                    End If
                                End If
                            End If
                        End If

                        ' Routing Rules

                        Dim ROUTING_RULE_NO As String = ""

                        If DEL_METHOD = "" Then

                            Dim rowSOTSHIPG As DataRow = Fill_Record("SOTSHIPG", New String() {SHIP_TO_CODE})
                            If rowSOTSHIPG Is Nothing Then
                                rowSOTSHIPG = Fill_Record("SOTSHIPG", New String() {Mid(SHIP_TO_CODE, 1, 8)})
                            End If
                            If dst.Tables("SOTSHIPG").Rows.Count = 0 Then
                                Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Could Not Find any Routing Records For Ship-To {SHIP_TO_CODE}", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="R")
                                Exit Try
                            End If

                            ROUTING_RULE_NO = rowSOTSHIPG.Item("ROUTING_RULE_NO")
                            Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Found Routing Record ({ROUTING_RULE_NO}) For Shipment {SHIP_BOL_NO}", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE)

                            Fill_Records("SOTSHIPY", ROUTING_RULE_NO)
                            For Each rowSOTSHIPY As DataRow In dst.Tables("SOTSHIPY").Select("", "FR_WGT")
                                Dim FR_WGT As Decimal = Val(rowSOTSHIPY.Item("FR_WGT") & "")
                                'Dim TO_WGT As Decimal = Val(rowSOTSHIPG.Item("TO_WGT") & "")
                                If SHIP_WEIGHT >= FR_WGT Then


                                    SHIP_VIA = rowSOTSHIPY.Item("SHIP_VIA") & ""
                                    rowSOTSHIPS = dst.Tables("SOTSHIPS").Rows.Find(SHIP_VIA)
                                    If rowSOTSHIPS Is Nothing Then
                                        Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Invalid Ship Via ({SHIP_VIA}) coming from Routing Rule ({ROUTING_RULE_NO})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="V")
                                        Exit Try
                                    Else
                                        DEL_METHOD = rowSOTSHIPS.Item("DEL_METHOD") & ""
                                        If DEL_METHOD <> "SPS" And DEL_METHOD <> "LTL" Then
                                            Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Invalid Delivery Method ({DEL_METHOD}) specified in provided Ship Via ({SHIP_VIA})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="V")
                                            Exit Try
                                        End If
                                    End If

                                    If rowSOTSHIPS.Item("CARRIER_ACCT_TYPE") & "" <> "" Then
                                        ' Ship Via has a specific Carrier Account Type to use
                                        CARRIER_ACCT_TYPE = rowSOTSHIPS.Item("CARRIER_ACCT_TYPE")
                                    End If
                                Else

                                    Exit For

                                End If
                            Next

                            If ROUTING_RULE_NO = "" Then
                                Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Could Not Find the correct Routing Rules For Shipment {SHIP_BOL_NO} With {SHIP_WEIGHT} lbs", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, EVENT_ERROR_TYPE:="G")
                                Exit Try
                            End If
                            Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Found Routing Rule For Shipment {SHIP_BOL_NO} With {SHIP_WEIGHT} lbs", EVENT_KEY:=ROUTING_RULE_NO, EVENT_KEY_TYPE:="R", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE)

                            If SHIP_VIA = "" Then
                                Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Could Not Determine Ship Via from Routing Rule For Shipment {SHIP_BOL_NO} With {SHIP_WEIGHT} lbs", EVENT_KEY:=ROUTING_RULE_NO, EVENT_KEY_TYPE:="R", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, EVENT_ERROR_TYPE:="G")
                                Exit Try
                            Else
                                If DEL_METHOD <> "SPS" And DEL_METHOD <> "LTL" Then
                                    Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Could Not Determine SPS Or LTL from Routing Rule For Shipment {SHIP_BOL_NO} With {SHIP_WEIGHT} lbs", EVENT_KEY:=ROUTING_RULE_NO, EVENT_KEY_TYPE:="R", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, EVENT_ERROR_TYPE:="G")
                                    Exit Try
                                End If
                            End If

                        End If

                        ' Get OPS Lead Time based on Total Weight

                        Dim rowSOTSHIPW As DataRow = dst.Tables("SOTSHIPW").Rows.Find(DEL_METHOD)
                        For I As Integer = 1 To 3
                            Dim MIN_SHIP_WGT_X As Int32 = Val(rowSOTSHIPW.Item($"MIN_SHIP_WGT_{Format(I, "0")}") & "")
                            Dim OPS_LEAD_TIME_DAYS_X As Int32 = Val(rowSOTSHIPW.Item($"OPS_LEAD_TIME_DAYS_{Format(I, "0")}") & "")
                            If SHIP_WEIGHT >= MIN_SHIP_WGT_X Then
                                OPS_LEAD_TIME_DAYS = OPS_LEAD_TIME_DAYS_X
                            Else
                                Exit For
                            End If
                        Next

                        Dim SHIP_CALL_IN_EARLIEST As Date
                        Dim SHIP_CALL_IN_LATEST As Date
                        Dim SHIP_XMIT_EARLIEST As Date
                        Dim SHIP_XMIT_LATEST As Date

                        Dim CALL_IN_RULE_NO As String = ""
                        Dim CALL_IN_NOTICE_DAYS As Int32 = 0

                        Dim SHIP_SUGG_CALL_IN_DATE As Date
                        Dim SHIP_SUGG_PICK_UP_DATE As Date

                        Dim CARRIER_ACCT_NO As String = ""
                        Dim SHIP_BILL_FRT_TO As String = ""

                        If DEL_METHOD = "SPS" Or SHIP_PRE_AUTH_REQD = "1" Then
                            ' No Routing Required, No Call-In Required

                            If DEL_METHOD = "SPS" Then
                                SHIP_METHOD = rowSOTSHIPS.Item("SHIP_METHOD") & ""

                                CARRIER = rowSOTSHIPS.Item("CARRIER") & ""
                                If CARRIER = "" Then
                                    Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"No Carrier specified in provided Ship Via ({SHIP_VIA_CODE})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="V")
                                    Exit Try
                                Else
                                    Dim rowSOTSHIPZ As DataRow = dst.Tables("SOTSHIPZ").Rows.Find(CARRIER)
                                    If rowSOTSHIPZ Is Nothing Then
                                        Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Invalid Carrier specified ({CARRIER}) in provided Ship Via ({SHIP_VIA_CODE})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="V")
                                        Exit Try
                                    End If
                                End If

                                Dim rowSOTSHIPQ As DataRow = dst.Tables("SOTSHIPQ").Rows.Find(New String() {CARRIER, SHIP_METHOD})
                                If rowSOTSHIPQ Is Nothing Then
                                    Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Invalid Shipping Method ({SHIP_METHOD}) specified in Ship Via ({SHIP_VIA_CODE})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="V")
                                    Exit Try
                                End If

                                If CARRIER_ACCT_TYPE = "" Then CARRIER_ACCT_TYPE = "DEFAULT"
                                Dim rowSOTSHIPU As DataRow = dst.Tables("SOTSHIPU").Rows.Find(New String() {CARRIER, CARRIER_ACCT_TYPE})
                                If rowSOTSHIPU Is Nothing Then
                                    Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Cannot Find IPLBs SPS Account No to use for Carrier ({CARRIER}) and Account Type ({CARRIER_ACCT_TYPE})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="V")
                                    Exit Try
                                Else
                                    CARRIER_ACCT_NO = rowSOTSHIPU.Item("CARRIER_ACCT_NO") & ""
                                    If CARRIER_ACCT_NO = "" Then
                                        Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"SPS Carrier Account No missing for Carrier ({CARRIER}) and Account Type ({CARRIER_ACCT_TYPE})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="V")
                                        Exit Try
                                    End If
                                End If

                                ' this definition is according to Malek
                                SHIP_CARRIER = CARRIER_ACCT_NO & SHIP_METHOD

                                If FRT_TERMS = "3PY" Or FRT_TERMS = "COL" Then
                                    SHIP_BILL_FRT_TO = ""
                                    If CARRIER = "" Then
                                        Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Cannot Find SPS Account No to use for Ship-To ({SHIP_TO_CODE}) - Unsupported Carrier ({CARRIER})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="V")
                                        Exit Try
                                    End If

                                    Dim rowSOTSHIPF As DataRow = dst.Tables("SOTSHIPF").Rows.Find(New String() {SHIP_TO_CODE, CARRIER})
                                    If rowSOTSHIPF Is Nothing OrElse rowSOTSHIPF.Item("CARRIER_ACCT_NO") & "" = "" Then
                                        Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Cannot Find SPS Account No to use for Ship-To ({SHIP_TO_CODE}) and Carrier ({CARRIER})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="V")
                                        Exit Try
                                    Else
                                        SHIP_BILL_FRT_TO = rowSOTSHIPD.Item("CARRIER_ACCT_NO") & ""
                                    End If
                                End If
                            End If

                            If SHIP_PRE_AUTH_REQD = "1" Then
                                ' not sure what to do
                            End If

                            SHIP_XMIT_EARLIEST = dSHIP.AddDays(-1 * OPS_LEAD_TIME_DAYS)
                            SHIP_XMIT_LATEST = dCANC.AddDays(-1 * OPS_LEAD_TIME_DAYS)

                        Else

                            SHIP_CARRIER = SO_PARM_DEF_LTL_CARRIER
                            SHIP_METHOD = "" ' ***************************  HB: Use GND, ask SB

                            Dim j As Integer = 0

                            ' Call-In Rules

                            'If SHIP_BOL_NO = "0000436730" Then Stop

                            Fill_Record("SOTSHIPI", New String() {SHIP_TO_CODE})
                            If dst.Tables("SOTSHIPI").Rows.Count = 0 Then
                                ' ************************** if no rule for ship to, see if we have a rule for the customer
                                Fill_Record("SOTSHIPI", New String() {Mid(SHIP_TO_CODE, 1, 8)})
                            End If

                            If dst.Tables("SOTSHIPI").Rows.Count = 0 Then
                                Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Could Not Find any Call-In Records For Ship-To {SHIP_TO_CODE}", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="I")
                                Exit Try
                            ElseIf dst.Tables("SOTSHIPI").Rows.Count > 1 Then
                                Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Found Multiple Call-In Records For Ship-To {SHIP_TO_CODE}", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="I")
                                Exit Try
                            End If

                            Dim RegEx01 As New System.Text.RegularExpressions.Regex("^[01]*$")

                            ' Call-In Calculation

                            Dim rowSOTSHIPI As DataRow = dst.Tables("SOTSHIPI").Rows(0)
                            CALL_IN_NOTICE_DAYS = Val(rowSOTSHIPI.Item("CALL_IN_NOTICE_DAYS") & "")
                            Dim CALL_IN_HOL_EXC As String = rowSOTSHIPI.Item("CALL_IN_HOL_EXC") & ""

                            CALL_IN_RULE_NO = rowSOTSHIPI.Item("CALL_IN_RULE_NO")
                            Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Found Call-In Record For Shipment {SHIP_BOL_NO}", EVENT_KEY:=CALL_IN_RULE_NO, EVENT_KEY_TYPE:="C", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE)

                            Dim CALL_IN_DAYS As String = rowSOTSHIPI.Item("CALL_IN_DAYS") & ""
                            If CALL_IN_DAYS.Length <> 7 Or Not RegEx01.IsMatch(CALL_IN_DAYS) Then
                                Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Issue with Call-In Days ({CALL_IN_DAYS}) For Ship-To {SHIP_TO_CODE}", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="I")
                                Exit Try
                            End If

                            ' Stick to Mon-Fri for Call-In Days
                            If CALL_IN_DAYS = "0000000" Then CALL_IN_DAYS = "0111110"
                            Mid(CALL_IN_DAYS, 1, 1) = "0"
                            Mid(CALL_IN_DAYS, 7, 1) = "0"

                            ' **********************************
                            If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" Then
                                'If SHIP_BOL_NO = "0000438023" Then Stop
                            End If

                            SHIP_CALL_IN_EARLIEST = Calc_WD("*", CALL_IN_DAYS, dSHIP, -1 * CALL_IN_NOTICE_DAYS)
                            SHIP_CALL_IN_LATEST = Calc_WD("*", CALL_IN_DAYS, dCANC, -1 * CALL_IN_NOTICE_DAYS)
                            SHIP_XMIT_EARLIEST = Calc_WD("*", CALL_IN_DAYS, SHIP_CALL_IN_EARLIEST, -1 * OPS_LEAD_TIME_DAYS)
                            SHIP_XMIT_LATEST = Calc_WD("*", CALL_IN_DAYS, SHIP_CALL_IN_LATEST, -1 * OPS_LEAD_TIME_DAYS)

                            SHIP_SUGG_CALL_IN_DATE = SHIP_CALL_IN_EARLIEST

                            ' the earliest sugg call-in date s/b at least OPS LT days after XMIT date
                            ' the XMIT date (which hasn't happened yet) is today for calc purposes
                            Dim EARLIEST_CALL_IN_DATE As Date = Calc_WD("*", CALL_IN_DAYS, Now.Date, OPS_LEAD_TIME_DAYS)
                            If Format(EARLIEST_CALL_IN_DATE, "yyyyMMdd") > Format(SHIP_SUGG_CALL_IN_DATE, "yyyyMMdd") Then '
                                ' the suggested call in date is earlier than the earliest permissible, so advance it to the earliest permissible
                                SHIP_SUGG_CALL_IN_DATE = EARLIEST_CALL_IN_DATE
                            End If

                            Dim SUGG_CALL_IN_DAY_NO As Int32 = 1 ' 1 = Sun, ..., 7 = Sat
                            Do
                                If SHIP_SUGG_CALL_IN_DATE.AddDays(1 - SUGG_CALL_IN_DAY_NO).DayOfWeek = Global.System.DayOfWeek.Sunday Then
                                    Exit Do
                                End If
                                SUGG_CALL_IN_DAY_NO += 1
                            Loop

                            Dim CALL_IN_DAY_RANGE_DAYS As Int32 = SHIP_CALL_IN_LATEST.Subtract(SHIP_CALL_IN_EARLIEST).TotalDays + 1
                            Dim CALL_IN_DAY_RANGE As String = CALL_IN_DAYS
                            Do Until CALL_IN_DAY_RANGE.Length > CALL_IN_DAY_RANGE_DAYS
                                CALL_IN_DAY_RANGE &= CALL_IN_DAYS
                            Loop
                            CALL_IN_DAY_RANGE &= CALL_IN_DAYS
                            CALL_IN_DAY_RANGE &= CALL_IN_DAYS

                            Dim SHIP_SUGG_CALL_IN_DATE_new As Date
                            Dim found_a_CALL_IN_date As Boolean = False
                            j = 0

                            Do Until found_a_CALL_IN_date
                                ' Find Next Call-In Date
                                Dim i As Integer = 0
                                If CALL_IN_DAY_RANGE.Length > SUGG_CALL_IN_DAY_NO + j Then
                                    i = InStr(Mid(CALL_IN_DAY_RANGE, SUGG_CALL_IN_DAY_NO + j), "1")
                                End If

                                If i = 0 Or (j + i - 1) > CALL_IN_DAY_RANGE_DAYS Then
                                    Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Could Not Find an available Call-In For Shipment {SHIP_BOL_NO}", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="J")
                                    Exit Try
                                End If

                                SHIP_SUGG_CALL_IN_DATE_new = Calc_WD(CALL_IN_HOL_EXC, CALL_IN_DAYS, SHIP_SUGG_CALL_IN_DATE, j + i - 1, , , True)

                                If SHIP_SUGG_CALL_IN_DATE_new.Subtract(Now.Date).TotalDays >= 0 Then
                                    found_a_CALL_IN_date = True
                                End If

                                If Not found_a_CALL_IN_date Then
                                    j += 1
                                End If
                            Loop

                            SHIP_SUGG_CALL_IN_DATE = SHIP_SUGG_CALL_IN_DATE_new


                            ' Pick-Up Calculation

                            ' **********************************
                            If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" Then
                                'If SHIP_BOL_NO = "0000438023" Then Stop
                            End If

                            Dim PICK_UP_HOL_EXC As String = rowSOTSHIPI.Item("PICK_UP_HOL_EXC") & ""

                            Dim PICK_UP_DAYS As String = rowSOTSHIPI.Item("PICK_UP_DAYS") & ""
                            If PICK_UP_DAYS.Length <> 7 Or Not RegEx01.IsMatch(PICK_UP_DAYS) Then
                                Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Issue with Pick-Up Days ({PICK_UP_DAYS}) For Ship-To {SHIP_TO_CODE}", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="I")
                                Exit Try
                            End If

                            ' Stick to Mon-Fri for Pick-Up-In Days
                            If PICK_UP_DAYS = "0000000" Then PICK_UP_DAYS = "0111110"
                            Mid(PICK_UP_DAYS, 1, 1) = "0"
                            Mid(PICK_UP_DAYS, 7, 1) = "0"


                            SHIP_SUGG_PICK_UP_DATE = Calc_WD("*", PICK_UP_DAYS, SHIP_SUGG_CALL_IN_DATE, 1) ' you must give retailer at least 24 hours

                            Dim SUGG_PICK_UP_DAY_NO As Int32 = 1 ' 1 = Sun, ..., 7 = Sat
                            Do
                                If SHIP_SUGG_PICK_UP_DATE.AddDays(1 - SUGG_PICK_UP_DAY_NO).DayOfWeek = Global.System.DayOfWeek.Sunday Then
                                    Exit Do
                                End If
                                SUGG_PICK_UP_DAY_NO += 1
                            Loop

                            Dim PICK_UP_DAY_RANGE_DAYS As Int32 = SHIP_CALL_IN_LATEST.Subtract(SHIP_CALL_IN_EARLIEST).TotalDays

                            Dim PICK_UP_DAY_RANGE As String = PICK_UP_DAYS
                            Do Until PICK_UP_DAY_RANGE.Length > PICK_UP_DAY_RANGE_DAYS
                                PICK_UP_DAY_RANGE &= PICK_UP_DAYS
                            Loop
                            PICK_UP_DAY_RANGE &= PICK_UP_DAYS

                            Dim SHIP_SUGG_PICK_UP_DATE_new As Date
                            Dim found_a_PICK_UP_date As Boolean = False
                            j = 0

                            Do Until found_a_PICK_UP_date
                                ' Find Next Pick-Up-In Date
                                Dim i As Integer = 0
                                If PICK_UP_DAY_RANGE.Length > SUGG_PICK_UP_DAY_NO + j Then
                                    i = InStr(Mid(PICK_UP_DAY_RANGE, SUGG_PICK_UP_DAY_NO + j), "1")
                                End If

                                If i = 0 Or (j + i - 1) > PICK_UP_DAY_RANGE_DAYS Then
                                    'If i <> 0 Then Stop
                                    Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Could Not Find an available Pick-Up Date for Shipment {SHIP_BOL_NO}", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="P")
                                    Exit Try
                                End If

                                SHIP_SUGG_PICK_UP_DATE_new = Calc_WD(PICK_UP_HOL_EXC, PICK_UP_DAYS, SHIP_SUGG_PICK_UP_DATE, j + i - 1, ,, True)

                                If SHIP_SUGG_PICK_UP_DATE_new.Subtract(Now.Date).TotalDays >= 0 Then
                                    found_a_PICK_UP_date = True
                                End If

                                If Not found_a_PICK_UP_date Then
                                    j += 1
                                End If
                            Loop

                            SHIP_SUGG_PICK_UP_DATE = SHIP_SUGG_PICK_UP_DATE_new

                        End If

                        DATETIME_STAMP = Now + ASCMAIN1.NowTSD
                        rowSOTSHIPR = Fill_Record("SOTSHIPR", SHIP_BOL_NO)
                        With rowSOTSHIPR

                            .Item("DEL_METHOD") = DEL_METHOD
                            .Item("OPS_LEAD_TIME_DAYS") = OPS_LEAD_TIME_DAYS

                            If DEL_METHOD = "LTL" Then

                                .Item("CALL_IN_NOTICE_DAYS") = CALL_IN_NOTICE_DAYS

                                .Item("SHIP_CALL_IN_EARLIEST") = SHIP_CALL_IN_EARLIEST
                                .Item("SHIP_CALL_IN_LATEST") = SHIP_CALL_IN_LATEST
                                .Item("SHIP_XMIT_EARLIEST") = SHIP_XMIT_EARLIEST
                                .Item("SHIP_XMIT_LATEST") = SHIP_XMIT_LATEST
                                .Item("SHIP_ROUTING_RULE_USED") = ROUTING_RULE_NO
                                .Item("SHIP_CALL_IN_RULE_USED") = CALL_IN_RULE_NO

                                .Item("SHIP_SUGG_CALL_IN_DATE") = SHIP_SUGG_CALL_IN_DATE
                                .Item("SHIP_SUGG_PICK_UP_DATE") = SHIP_SUGG_PICK_UP_DATE

                            Else

                                .Item("CALL_IN_NOTICE_DAYS") = DBNull.Value

                                .Item("SHIP_CALL_IN_EARLIEST") = DBNull.Value
                                .Item("SHIP_CALL_IN_LATEST") = DBNull.Value
                                .Item("SHIP_XMIT_EARLIEST") = SHIP_XMIT_EARLIEST
                                .Item("SHIP_XMIT_LATEST") = SHIP_XMIT_LATEST
                                .Item("SHIP_ROUTING_RULE_USED") = ""
                                .Item("SHIP_CALL_IN_RULE_USED") = ""

                                .Item("SHIP_SUGG_CALL_IN_DATE") = DBNull.Value
                                .Item("SHIP_SUGG_PICK_UP_DATE") = DBNull.Value
                            End If

                            .Item("SHIP_BILL_FRT_TO") = SHIP_BILL_FRT_TO

                            .Item("SHIP_CARRIER") = SHIP_CARRIER
                            .Item("SHIP_METHOD") = SHIP_METHOD

                            .Item("CARRIER_ACCT_TYPE") = CARRIER_ACCT_TYPE
                            .Item("CARRIER_ACCT_NO") = CARRIER_ACCT_NO

                            .Item("SHIP_WEIGHT") = SHIP_WEIGHT
                            .Item("SHIP_CARTONS") = SHIP_CARTONS

                            .Item("LAST_DATE") = DATETIME_STAMP
                            .Item("LAST_OPER") = ASCMAIN1.USER_ID
                        End With

                        Update_Record_TDA("SOTSHIPR")
                        Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Recording Routing & Call-In Date Results for Shipment {SHIP_BOL_NO}", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE)

                        DATETIME_STAMP = Now + ASCMAIN1.NowTSD

                        ' If SHIP_BOL_NO = "0000433652" Then Stop

                        If Format(SHIP_XMIT_LATEST, "yyyyMMdd") < Format(DATETIME_STAMP, "yyyyMMdd") Then
                            Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Latest Transmit Date ({Format(SHIP_XMIT_LATEST, "MM/dd/yyyy")}) has passed", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="K")
                            Exit Try
                        End If

                        ASCMAIN1.sql = "Update SOTSHIP1 Set SHIP_ROUTING_IND = '1' where SHIP_BOL_NO = :PARM1"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {SHIP_BOL_NO})

                    Catch ex As Exception
                        Dim MSG As String = ex.Message
                        If MSG.Length > 100 Then
                            MSG = Mid(MSG, 1, 100)
                        End If
                        Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"{MSG}", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=SHIP_ADDR_CODE, SHIP_TO_CODE:=SHIP_TO_CODE, EVENT_ERROR_TYPE:="X")
                    End Try
                Next

                ASCMAIN1.Progress("")
                MsgBox("Complete")
                ASCMAIN1.MultiTask_Release()

                Fill_Records("SOTSHIPR_VIEW")

                Fill_Records("SOTSHIPV_ALL")
                Sort_grdColumns(grdSOTSHIPV_ALL, "EVENT_DATE")

                dst.Tables("SOTSHIP1_MATRIX").Rows.Clear()
                ' *************************** THIS TABLE NEEDS TO BE LOADED WITH ALL SHIPMENTS THAT WERE ROUTED AND NOT XMITTED

                Dim filter As String = "DEL_METHOD = 'LTL' and SHIP_SUGG_CALL_IN_DATE is not null and SHIP_SUGG_PICK_UP_DATE is not null"

                If dst.Tables("SOTSHIPR_VIEW").Select(filter).Length > 0 Then

                    Dim CALL_IN_MIN As Date = dst.Tables("SOTSHIPR_VIEW").Compute("MIN(SHIP_SUGG_CALL_IN_DATE)", filter)
                    Dim CALL_IN_MAX As Date = dst.Tables("SOTSHIPR_VIEW").Compute("MAX(SHIP_SUGG_CALL_IN_DATE)", filter)
                    Dim PICK_UP_MIN As Date = dst.Tables("SOTSHIPR_VIEW").Compute("MIN(SHIP_SUGG_PICK_UP_DATE)", filter)
                    Dim PICK_UP_MAX As Date = dst.Tables("SOTSHIPR_VIEW").Compute("MAX(SHIP_SUGG_PICK_UP_DATE)", filter)

                    Dim DTE As Date = CALL_IN_MIN
                    Do
                        dst.Tables("SOTSHIP1_MATRIX").Rows.Add(New Object() {DTE, Mid(DTE.DayOfWeek.ToString, 1, 3)})
                        DTE = DTE.AddDays(1)
                    Loop Until Format(DTE, "yyyyMMdd") > Format(CALL_IN_MAX, "yyyyMMdd")

                    Dim dx As Integer = 0
                    DTE = PICK_UP_MIN
                    grdSOTSHIP1_MATRIX.DisplayLayout.Bands(0).ColHeaderLines = 2
                    Do
                        dx += 1
                        With grdSOTSHIP1_MATRIX.DisplayLayout.Bands(0).Columns($"PICK_UP_{Format(dx, "00")}")
                            .Header.Caption = Format(DTE, "MM/dd") & vbCrLf & Mid(DTE.DayOfWeek.ToString, 1, 3)
                            .Hidden = False
                            .Header.Appearance.TextHAlign = HAlign.Center
                            .CellAppearance.TextHAlign = HAlign.Center
                            If DTE.DayOfWeek = Global.System.DayOfWeek.Sunday Then
                                .Header.Appearance.ForeColor = System.Drawing.Color.Red
                            Else
                                .Header.Appearance.ForeColor = System.Drawing.Color.Empty
                            End If
                        End With
                        DTE = DTE.AddDays(1)
                    Loop Until Format(DTE, "yyyyMMdd") > Format(PICK_UP_MAX, "yyyyMMdd")

                    If dx < 20 Then
                        For i As Integer = dx + 1 To 20
                            grdSOTSHIP1_MATRIX.DisplayLayout.Bands(0).Columns($"PICK_UP_{Format(i, "00")}").Hidden = True
                        Next
                    End If

                    For Each rowSOTSHIPR_VIEW As DataRow In dst.Tables("SOTSHIPR_VIEW").Select(filter)
                        Dim CALL_IN_DATE As Date = rowSOTSHIPR_VIEW.Item("SHIP_SUGG_CALL_IN_DATE")
                        Dim PICK_UP_DATE As Date = rowSOTSHIPR_VIEW.Item("SHIP_SUGG_PICK_UP_DATE")
                        Dim PICK_UP_I As Integer = PICK_UP_DATE.Subtract(PICK_UP_MIN).TotalDays + 1
                        Dim rowSOTSHIP1_MATRIX As DataRow = dst.Tables("SOTSHIP1_MATRIX").Rows.Find(CALL_IN_DATE)
                        Dim PICK_UP_X As Integer = Val(rowSOTSHIP1_MATRIX.Item($"PICK_UP_{Format(PICK_UP_I, "00")}") & "")
                        rowSOTSHIP1_MATRIX.Item($"PICK_UP_{Format(PICK_UP_I, "00")}") = PICK_UP_X + 1
                    Next
                    Sort_grdColumns(grdSOTSHIP1_MATRIX, "CALL_IN_DATE")
                End If

                tab0.SelectedTab = tab0.Tabs("CARS")

                Check_email(SO_PARM_EMAIL_FREQ, SO_PARM_CARS_END, DATE_LAST_EMAIL)

                '*************************************
                If ASCMAIN1.Running_in_VS Then Exit Do
                '*************************************
            Loop

            '*************************************
            If ASCMAIN1.Running_in_VS Then
                dst.Tables("ASTSQLX1").Rows.Clear()
            End If
            '*************************************

            If CHECK_EMAIL_ONCE_AFTER_END Then
                Check_email(SO_PARM_EMAIL_FREQ, SO_PARM_CARS_END, DATE_LAST_EMAIL)
                CHECK_EMAIL_ONCE_AFTER_END = False
            End If


            '*************************************
            If ASCMAIN1.Running_in_VS Then Exit Do
            '*************************************

            If SO_PARM_CALC_PAUSE_SECS <> 0 Then
                System.Threading.Thread.Sleep(SO_PARM_CALC_PAUSE_SECS * 1000)
            End If
        Loop

    End Sub

    Function Calc_WD(CALL_IN_HOL_EXC As String, X_DAYS As String, dt As Date, days As Int32,
      Optional s As Integer = 0, Optional last_day As Boolean = False, Optional check_NWD_only As Boolean = False) As Date

        'SHIP_CALL_IN_EARLIEST = Calc_WD("*", CALL_IN_DAYS = "0101010", ORDR_SHIP_DATE = "05/16/2025", -1 * CALL_IN_NOTICE_DAYS = 4)

        If s = 0 Then s = Math.Sign(days)
        If s = 0 Then s = 1

        Do While dt.DayOfWeek = Global.System.DayOfWeek.Saturday Or dt.DayOfWeek = Global.System.DayOfWeek.Sunday
            dt = dt.AddDays(s)
        Loop

        If CALL_IN_HOL_EXC = "*" And days = 0 Then
            ' what about if we are now on a holiday?
            Return dt
        End If

        Dim rowSOTPARMH As DataRow = Nothing
        Do
            rowSOTPARMH = dst.Tables("SOTPARMH").Rows.Find(dt)
            If rowSOTPARMH IsNot Nothing Then
                Dim NATIONAL_HOLIDAY As String = rowSOTPARMH.Item("NATIONAL_HOLIDAY") & ""
                If last_day And NATIONAL_HOLIDAY = "1" And "ELS".Contains(CALL_IN_HOL_EXC) Then

                    If CALL_IN_HOL_EXC = "E" Then
                        dt = dt.AddDays(-1)
                        dt = Calc_WD("*", X_DAYS, dt, 0, -1)
                    End If
                    If CALL_IN_HOL_EXC = "L" Then
                        dt = dt.AddDays(1)
                        dt = Calc_WD("*", X_DAYS, dt, 0, 1)
                    End If
                    If CALL_IN_HOL_EXC = "S" Then
                        Dim Z As String = X_DAYS & X_DAYS
                        Dim DNO As Integer = 0
                        If dt.DayOfWeek = Global.System.DayOfWeek.Sunday Then DNO = 1
                        If dt.DayOfWeek = Global.System.DayOfWeek.Monday Then DNO = 2
                        If dt.DayOfWeek = Global.System.DayOfWeek.Tuesday Then DNO = 3
                        If dt.DayOfWeek = Global.System.DayOfWeek.Wednesday Then DNO = 4
                        If dt.DayOfWeek = Global.System.DayOfWeek.Thursday Then DNO = 5
                        If dt.DayOfWeek = Global.System.DayOfWeek.Friday Then DNO = 6
                        If dt.DayOfWeek = Global.System.DayOfWeek.Saturday Then DNO = 7
                        '0101010'
                        Dim DNO_ADJ As Integer = InStr(Mid(X_DAYS, DNO + 1), "1")
                        'dt = dt.AddDays(DNO_ADJ)
                        dt = Calc_WD("*", X_DAYS, dt, 0, DNO_ADJ)
                    End If
                Else
                    dt = dt.AddDays(s) ' Move to the next day, either forward or back
                End If
            Else
                If Not check_NWD_only Then ' not sure why
                    dt = dt.AddDays(s)
                    dt = Calc_WD("*", X_DAYS, dt, 0, Math.Sign(s))
                End If
            End If
        Loop Until rowSOTPARMH Is Nothing

        If days = 0 Then
            Return dt
        End If

        Dim zz As String = CALL_IN_HOL_EXC
        If zz = "*" Then zz = "**"

        For i As Integer = 1 To Math.Abs(days) - 1 ' 0 To Math.Abs(days) - 1
            'If i = 0 Then
            '    dt = Calc_WD(CALL_IN_HOL_EXC, X_DAYS, dt, 0, s)
            'Else
            'dt = Calc_WD(CALL_IN_HOL_EXC, X_DAYS, dt, s, 0, i = Math.Abs(days))
            dt = Calc_WD(zz, X_DAYS, dt, 0, s, (i = Math.Abs(days) - 1))
            'End If
        Next

        Return dt

    End Function

    Sub Check_email(SO_PARM_EMAIL_FREQ As Int32, SO_PARM_CARS_END As Int32, DATE_LAST_EMAIL As Date)

        ' NEED TO CHECK FOR EMAIL SENDING ONCE AFTER WE EXCEED SO_PARM_CARS_END 

        ASCMAIN1.sql = $"Select 
            SOTSHIPV.EVENT_DATE,
            SOTSHIPV.EVENT_MSG,
            SOTSHIPV.SHIP_BOL_NO,
            SOTSHIPV.CUST_CODE,
            SOTSHIPV.SHIP_ADDR_CODE,
            SOTSHIPV.SHIP_TO_CODE,
            SOTSHIPV.EVENT_ERROR_TYPE
             from SOTSHIPV
            where SOTSHIPV.EVENT_ERROR_TYPE between 'A' and 'Z'
              and SOTSHIPV.EVENT_ERROR_EMAIL_IND = '0'"

        'and SOTSHIPV.SHIP_ROUTING_XNO = '{SHIP_ROUTING_XNO}'"

        Dim dt As DataTable = ASCDATA1.GetDataTable

        If dt.Rows.Count > 0 Then
            If DATE_LAST_EMAIL.Year < 2000 Then
                DATE_LAST_EMAIL = Now
            End If
            If SO_PARM_EMAIL_FREQ = 0 Or ((Now.Subtract(DATE_LAST_EMAIL).TotalSeconds > SO_PARM_EMAIL_FREQ Or Now.Hour > SO_PARM_CARS_END)) Then
                Send_emails_CARS_Errors(dt)
                DATE_LAST_EMAIL = Now
            End If
        End If
    End Sub

    Sub Record_CARS_Event(SHIP_ROUTING_XNO As String, Optional EVENT_KEY_TYPE As String = "", Optional EVENT_KEY As String = "", Optional EVENT_MSG As String = "", Optional EVENT_KEY_COUNT As Int32 = 0,
                          Optional SHIP_BOL_NO As String = "", Optional CUST_CODE As String = "", Optional SHIP_ADDR_CODE As String = "", Optional SHIP_TO_CODE As String = "", Optional EVENT_ERROR_TYPE As String = "")

        ASCMAIN1.Progress(EVENT_MSG)

        ASCMAIN1.sql = "Insert into SOTSHIPV (EVENT_DATE, EVENT_KEY_TYPE, EVENT_KEY, EVENT_MSG, EVENT_KEY_COUNT," & vbCrLf _
            & "SHIP_BOL_NO, CUST_CODE, SHIP_ADDR_CODE, SHIP_TO_CODE, EVENT_ERROR_TYPE, SHIP_ROUTING_XNO, EVENT_ERROR_EMAIL_IND) " & vbCrLf _
            & " Values (SYSDATE, :PARM1, :PARM2, :PARM3, :PARM4, :PARM5, :PARM6, :PARM7, :PARM8, :PARM9, :PARM10, '0')"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVNVVVVVV", New Object() {EVENT_KEY_TYPE, EVENT_KEY, EVENT_MSG, EVENT_KEY_COUNT, SHIP_BOL_NO, CUST_CODE, SHIP_ADDR_CODE, SHIP_TO_CODE, EVENT_ERROR_TYPE, SHIP_ROUTING_XNO})

        If EVENT_ERROR_TYPE <> "" Then
            If EVENT_ERROR_TYPE <> "L" Then
                ASCMAIN1.sql = "Update SOTSHIP1 Set SHIP_ROUTING_IND = :PARM1 where SHIP_BOL_NO = :PARM2"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {EVENT_ERROR_TYPE, SHIP_BOL_NO})
            End If
            ASCMAIN1.MultiTask_Release()
        End If

    End Sub

    Function Get_SHIP_TO_Record(SHIP_ROUTING_XNO As String, SHIP_BOL_NO As String, CUST_CODE As String, CUST_STORE_NO As String) As DataRow
        Dim SHIP_TO_TYPE As String = "S" ' This method works for SHIP_TO_TYPE S (Store)
        Dim SHIP_TO_CODE As String = ""

        Dim rowSOTSHIPD As DataRow = Fill_Record("SOTSHIPD", New String() {CUST_CODE, CUST_STORE_NO})
        If rowSOTSHIPD Is Nothing Then
            Dim SHIP_TO_CODE_PFX As String = Get_SHIP_TO_CODE_PFX(SHIP_ROUTING_XNO, SHIP_BOL_NO, CUST_CODE)
            rowSOTSHIPD = dst.Tables("SOTSHIPD").NewRow
            DATETIME_STAMP = Now + ASCMAIN1.NowTSD
            SHIP_TO_CODE = SHIP_TO_CODE_PFX & "P" & CUST_STORE_NO
            With rowSOTSHIPD
                .Item("SHIP_TO_CODE") = SHIP_TO_CODE
                .Item("SHIP_TO_TYPE") = SHIP_TO_TYPE
                .Item("SHIP_TO_CODE_PFX") = SHIP_TO_CODE_PFX
                .Item("CUST_CODE") = CUST_CODE
                .Item("CUST_STORE_NO") = CUST_STORE_NO
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_DATE") = DATETIME_STAMP
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
            End With
            dst.Tables("SOTSHIPD").Rows.Add(rowSOTSHIPD)

            Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Established New Ship-To-Code ({SHIP_TO_CODE})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=CUST_STORE_NO, SHIP_TO_CODE:=SHIP_TO_CODE)

            Update_Record_TDA("SOTSHIPD")
        Else
            SHIP_TO_CODE = rowSOTSHIPD.Item("SHIP_TO_CODE")

            Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Found Ship-To-Code ({SHIP_TO_CODE})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE, SHIP_ADDR_CODE:=CUST_STORE_NO, SHIP_TO_CODE:=SHIP_TO_CODE)

        End If

        Return rowSOTSHIPD
    End Function

    Function Get_SHIP_TO_CODE_PFX(SHIP_ROUTING_XNO As String, SHIP_BOL_NO As String, CUST_CODE As String) As String
        Dim SHIP_TO_TYPE As String = "C" ' This method works for SHIP_TO_TYPE C (Customer)
        Dim SHIP_TO_CODE_PFX As String = ""
        ' write event determining Ship-To PFX
        Dim rowSOTSHIPE As DataRow = Fill_Record("SOTSHIPE", New String() {SHIP_TO_TYPE, CUST_CODE})
        If rowSOTSHIPE Is Nothing Then
            SHIP_TO_CODE_PFX = ASCMAIN1.Next_Control_No("SOTSHIPE.SHIP_TO_CODE_PFX")
            rowSOTSHIPE = dst.Tables("SOTSHIPE").NewRow
            DATETIME_STAMP = Now + ASCMAIN1.NowTSD
            With rowSOTSHIPE
                .Item("SHIP_TO_TYPE") = SHIP_TO_TYPE
                .Item("SHIP_TO_KEY") = CUST_CODE
                .Item("SHIP_TO_CODE_PFX") = SHIP_TO_CODE_PFX
                .Item("CUST_CODE") = CUST_CODE
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_DATE") = DATETIME_STAMP
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
            End With
            dst.Tables("SOTSHIPE").Rows.Add(rowSOTSHIPE)

            Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Established New Ship-To-Code-Prefix ({SHIP_TO_CODE_PFX})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE)

            Update_Record_TDA("SOTSHIPE")
        Else
            SHIP_TO_CODE_PFX = rowSOTSHIPE.Item("SHIP_TO_CODE_PFX")

            Record_CARS_Event(SHIP_ROUTING_XNO, EVENT_MSG:=$"Found Ship-To-Code-Prefix ({SHIP_TO_CODE_PFX})", SHIP_BOL_NO:=SHIP_BOL_NO, CUST_CODE:=CUST_CODE)

        End If

        Return SHIP_TO_CODE_PFX
    End Function

    Private Sub grdSOTSHIPR_VIEW_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTSHIPR_VIEW.AfterRowActivate
        If Not grdSOTSHIPR_VIEW.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Dim SHIP_BOL_NO As String = grdSOTSHIPR_VIEW.ActiveRow.Cells("SHIP_BOL_NO").Value

        Fill_Records("SOTSHIPV", SHIP_BOL_NO)
        Sort_grdColumns(grdSOTSHIPV, "EVENT_DATE")
        grdSOTSHIPV.Text = $"Event Log for Shipment {SHIP_BOL_NO}"

        Fill_Records("SOTSHIPP", SHIP_BOL_NO)
        Sort_grdColumns(grdSOTSHIPP, "PICK_NO")
        grdSOTSHIPP.Text = $"Pick Tickets for Shipment {SHIP_BOL_NO}"

        Fill_Records("SOTSHIP1_WGT", SHIP_BOL_NO)
        Sort_grdColumns(grdSOTSHIP1_WGT, "ITEM_CODE")
        grdSOTSHIP1_WGT.Text = $"Items and Weights for Shipment {SHIP_BOL_NO}"

        Create_Timeline(SHIP_BOL_NO)

    End Sub

    Function GetDay(text As String, DT As Date) As String
        Return text & vbCrLf & Format(DT, "MM/dd") & " " & Mid(DT.DayOfWeek.ToString(), 1, 3)
    End Function
    Sub Create_Timeline(SHIP_BOL_NO As String)

        Dim rowTL As DataRow = dst.Tables("SOTSHIPR_VIEW").Rows.Find(SHIP_BOL_NO)

        Dim ORDR_SHIP_DATE As Date = rowTL.Item("SHIP_SHIP_DATE")
        Dim ORDR_CANCEL_DATE As Date = rowTL.Item("SHIP_CANCEL_DATE")
        Dim OPS_LT1 As Int32 = Val(rowTL.Item("OPS_LEAD_TIME_DAYS") & "")
        Dim OPS_LT2 As Int32 = Val(rowTL.Item("OPS_LEAD_TIME_DAYS") & "")
        Dim SHIP_CALL_IN_RULE_USED As String = rowTL.Item("SHIP_CALL_IN_RULE_USED") & ""
        Dim CALL_IN_DAYS As String = "0000000" ' "0100010"
        Dim PICK_UP_DAYS As String = "0000000" ' "0100010"
        Dim CALL_IN_NOTICE_DAYS As Int32 = Val(rowTL.Item("CALL_IN_NOTICE_DAYS") & "")

        Dim SHIP_SUGG_CALL_IN_DATE As Date = Nothing
        Dim SHIP_SUGG_PICK_UP_DATE As Date = Nothing
        If (rowTL.Item("SHIP_SUGG_CALL_IN_DATE") & "") <> "" Then SHIP_SUGG_CALL_IN_DATE = rowTL.Item("SHIP_SUGG_CALL_IN_DATE")
        If (rowTL.Item("SHIP_SUGG_CALL_IN_DATE") & "") <> "" Then SHIP_SUGG_PICK_UP_DATE = rowTL.Item("SHIP_SUGG_PICK_UP_DATE")

        Dim rowSOTSHIPI As DataRow = Nothing
        If SHIP_CALL_IN_RULE_USED <> "" Then
            rowSOTSHIPI = LookUp("SOTSHIPI", SHIP_CALL_IN_RULE_USED)
            If rowSOTSHIPI IsNot Nothing Then
                CALL_IN_DAYS = rowSOTSHIPI.Item("CALL_IN_DAYS") & ""
                If CALL_IN_DAYS = "" Then CALL_IN_DAYS = "0000000"
                PICK_UP_DAYS = rowSOTSHIPI.Item("PICK_UP_DAYS") & ""
                If PICK_UP_DAYS = "" Then PICK_UP_DAYS = "0000000"
                txtHolidayCallIn.Text = rowSOTSHIPI.Item("CALL_IN_HOL_EXC") & ""
                txtHolidayPickUp.Text = rowSOTSHIPI.Item("PICK_UP_HOL_EXC") & ""
            Else
                txtHolidayCallIn.Text = ""
                txtHolidayPickUp.Text = ""
            End If
        End If

        chkC_SUN.Checked = (Mid(CALL_IN_DAYS, 1, 1) = "1")
        chkC_MON.Checked = (Mid(CALL_IN_DAYS, 2, 1) = "1")
        chkC_TUE.Checked = (Mid(CALL_IN_DAYS, 3, 1) = "1")
        chkC_WED.Checked = (Mid(CALL_IN_DAYS, 4, 1) = "1")
        chkC_THU.Checked = (Mid(CALL_IN_DAYS, 5, 1) = "1")
        chkC_FRI.Checked = (Mid(CALL_IN_DAYS, 6, 1) = "1")
        chkC_SAT.Checked = (Mid(CALL_IN_DAYS, 7, 1) = "1")

        chkP_SUN.Checked = (Mid(PICK_UP_DAYS, 1, 1) = "1")
        chkP_MON.Checked = (Mid(PICK_UP_DAYS, 2, 1) = "1")
        chkP_TUE.Checked = (Mid(PICK_UP_DAYS, 3, 1) = "1")
        chkP_WED.Checked = (Mid(PICK_UP_DAYS, 4, 1) = "1")
        chkP_THU.Checked = (Mid(PICK_UP_DAYS, 5, 1) = "1")
        chkP_FRI.Checked = (Mid(PICK_UP_DAYS, 6, 1) = "1")
        chkP_SAT.Checked = (Mid(PICK_UP_DAYS, 7, 1) = "1")

        Dim lineWidth As Int32 = 2

        Dim marginTop As Int32 = 20 ' Top margin
        Dim marginBot As Int32 = 20 ' Bottom margin

        Dim surf As Graphics = Panel1.CreateGraphics
        surf.Clear(Color.White)
        Dim pBlack As New Pen(System.Drawing.Color.Black, 1)
        Dim pGreen As New Pen(System.Drawing.Color.Green, 1)
        Dim pBlue As New Pen(System.Drawing.Color.Blue, 1)
        Dim pRed As New Pen(System.Drawing.Color.Red, 1)
        Dim pRed2 As New Pen(System.Drawing.Color.Red, 4)
        Dim pYellow As New Pen(System.Drawing.Color.Yellow, 1)
        Dim pGray As New Pen(System.Drawing.Color.Gray, 1)
        Dim pDarkOrange As New Pen(System.Drawing.Color.DarkOrange, lineWidth)
        Dim pPurple As New Pen(System.Drawing.Color.Purple, lineWidth)

        Dim marginLeft As Int32 = 50 ' 50 pixel margin
        Dim timelineWidth As Int32 = Panel1.Width - marginLeft * 2 ' width of timeline = total width - (left+right margin)
        Dim timelineY As Int32 = CInt(Panel1.Height / 2) + marginTop ' vertical position of timeline

        surf.DrawLine(pBlack, marginLeft, timelineY, timelineWidth + marginLeft, timelineY) ' This is the main Timeline

        Dim dSHIP As Date = ORDR_SHIP_DATE
        If dSHIP.DayOfWeek = Global.System.DayOfWeek.Saturday Then
            dSHIP = dSHIP.AddDays(2) ' prob need to check for holiday
        ElseIf dSHIP.DayOfWeek = Global.System.DayOfWeek.Sunday Then
            dSHIP = dSHIP.AddDays(1) ' prob need to check for holiday
        End If
        Dim dCANC As Date = ORDR_CANCEL_DATE
        If dCANC.DayOfWeek = Global.System.DayOfWeek.Saturday Then
            dCANC = dCANC.AddDays(-1) ' prob need to check for holiday
        ElseIf dCANC.DayOfWeek = Global.System.DayOfWeek.Sunday Then
            dCANC = dCANC.AddDays(-2) ' prob need to check for holiday
        End If
        Dim daysShipWindow As Int32 = dCANC.Subtract(dSHIP).TotalDays + 1 ' days in Ship Window

        Dim bufferLeft As Int32 = 2 ' number of days on timeline for margin left before plotting
        Dim bufferRight As Int32 = 2 ' number of days on timeline for margin right before plotting
        If Format(Now, "yyyyMMdd") > Format(dSHIP, "yyyyMMdd") Then
            bufferRight += 5
        End If
        If Format(Now, "yyyyMMdd") < Format(dSHIP, "yyyyMMdd") Then
            Dim DADJ As Integer = dSHIP.Subtract(Now.Date).TotalDays
            If DADJ > 10 Then
                'bufferLeft += DADJ
                bufferLeft = CInt(DADJ * 0.75)
            End If
        End If

        If daysShipWindow < 10 Then
            bufferLeft += 3
            bufferRight += 3
        End If
        Dim days As Int32 = daysShipWindow * 2 + bufferLeft + bufferRight ' total days in timeline, 2 Ship Windows + 2 buffers

        Dim d1 As Date = dCANC.AddDays(-1 * days + bufferRight) ' starting date of timeline for plotting

        Dim rectHeight As Int32 = 40
        Dim xx As Int32 = 0 ' working variable
        Dim yy As Int32 = 0 ' working variable
        Dim legendWidth As Int32 = 100 ' width of a Flag Rectangle
        Dim gap As Int32 = 10

        ' Draw flag for Today Purple

        Dim yNow As Int32 = timelineY - marginTop - rectHeight * 2 - (rectHeight + gap) * 2
        Dim xNow As Int32 = Now.Date.Subtract(d1).TotalDays * timelineWidth / days
        surf.DrawLine(pPurple, marginLeft + xNow, yNow, marginLeft + xNow, timelineY)

        Dim xAdjLR_of_Now As Int32 = legendWidth * 2
        'If Now.Date.Subtract(d1).TotalDays > days / 2 Then
        '    xAdjLR_of_Now = -1 * xAdjLR_of_Now ' if Now is > half of Timline, show Suggs on the left else right
        '    ' if on left, maybe need an extra legendWidth to the left (ie, when we -1 *)
        'End If

        Dim rectTodayLegend As New Rectangle(marginLeft + xNow, yNow, legendWidth, rectHeight)
        surf.DrawRectangle(pPurple, rectTodayLegend)
        AddText(surf, rectTodayLegend, GetDay("Today", Now), Brushes.Purple, 9)

        ' Draw flag for Suggested Call-In Date Red

        If SHIP_SUGG_CALL_IN_DATE <> "01/01/0001" Then
            yy = timelineY - marginTop - rectHeight * 2 - (rectHeight + gap) * 2
            Dim xxSugg As Int32 = SHIP_SUGG_CALL_IN_DATE.Subtract(d1).TotalDays * timelineWidth / days
            If Format(SHIP_SUGG_CALL_IN_DATE, "yyyyMMdd") > Format(Now, "yyyyMMdd") Then
                xx = xNow + xAdjLR_of_Now
            Else
                xx = xNow - xAdjLR_of_Now
            End If
            surf.DrawLine(pRed, marginLeft + xx + CInt(100 / 2), yy + rectHeight, marginLeft + xxSugg, timelineY)

            Dim rectSuggLegend As New Rectangle(marginLeft + xx, yy, legendWidth, rectHeight)
            surf.DrawRectangle(pRed, rectSuggLegend)
            AddText(surf, rectSuggLegend, GetDay("Sugg Call-In", SHIP_SUGG_CALL_IN_DATE), Brushes.Red, 9)
        End If

        ' Draw flag for Suggested Pick-Up Date Red

        If SHIP_SUGG_PICK_UP_DATE <> "01/01/0001" Then
            Dim ySugg As Int32 = timelineY - marginTop - rectHeight * 2 - (rectHeight + gap) * 2
            Dim xxSugg As Int32 = SHIP_SUGG_PICK_UP_DATE.Subtract(d1).TotalDays * timelineWidth / days
            xx = xNow + xAdjLR_of_Now + legendWidth
            If Format(SHIP_SUGG_PICK_UP_DATE, "yyyyMMdd") > Format(Now, "yyyyMMdd") Then
                xx = xNow + xAdjLR_of_Now + legendWidth
            Else
                xx = xNow - xAdjLR_of_Now + legendWidth
            End If

            surf.DrawLine(pRed, marginLeft + xx + CInt(100 / 2), ySugg + rectHeight, marginLeft + xxSugg, timelineY)

            Dim rectSuggLegend As New Rectangle(marginLeft + xx, ySugg, 100, rectHeight)
            surf.DrawRectangle(pRed, rectSuggLegend)
            AddText(surf, rectSuggLegend, GetDay("Sugg Pick-Up", SHIP_SUGG_PICK_UP_DATE), Brushes.Red, 9)
        End If

        ' Turn SHIP & CANC into #days from d1, and then scale #days to an x coord on the timeline
        ' Create Ship Window & draw Green Rectangle

        Dim xxSHIP As Int32 = ORDR_SHIP_DATE.Subtract(d1).TotalDays * timelineWidth / days
        Dim xxSHIPadj As Int32 = dSHIP.Subtract(d1).TotalDays * timelineWidth / days

        Dim extraDay As Int32 = 0
        If ORDR_CANCEL_DATE.DayOfWeek = Global.System.DayOfWeek.Saturday Then
        ElseIf ORDR_CANCEL_DATE.DayOfWeek = Global.System.DayOfWeek.Sunday Then
        Else
            extraDay += 1
        End If

        Dim yyCANC As Int32 = ORDR_CANCEL_DATE.Subtract(d1).TotalDays * timelineWidth / days
        Dim yyCANCx As Int32 = (ORDR_CANCEL_DATE.Subtract(d1).TotalDays + extraDay) * timelineWidth / days
        Dim yyCANCadj As Int32 = dCANC.Subtract(d1).TotalDays * timelineWidth / days

        Dim rectShipWindow As New Rectangle(marginLeft + xxSHIP, timelineY + marginBot, (yyCANCx - xxSHIP), rectHeight)
        surf.FillRectangle(Brushes.Green, rectShipWindow)
        Dim daysShipWindowReally As Int32 = dCANC.Subtract(dSHIP).TotalDays + 1 ' days really in Ship Window

        AddText(surf, rectShipWindow, "Ship Window" & $" {daysShipWindowReally} days", Brushes.White)

        ' Draw flags for Ship Date and Cancel Date

        surf.DrawLine(pGreen, marginLeft + xxSHIP, timelineY, marginLeft + xxSHIP, timelineY + 80)
        Dim rectSHIPWindowLegend As New Rectangle(marginLeft + xxSHIP, timelineY + 50 + marginBot, 100, rectHeight)
        surf.DrawRectangle(pGreen, rectSHIPWindowLegend)
        AddText(surf, rectSHIPWindowLegend, GetDay("Ship Date", ORDR_SHIP_DATE), Brushes.Green, 9)

        surf.DrawLine(pGreen, marginLeft + yyCANC, timelineY, marginLeft + yyCANC, timelineY + 80)
        Dim rectCANCWindowLegend As New Rectangle(marginLeft + yyCANC, timelineY + 50 + marginBot, 100, rectHeight)
        surf.DrawRectangle(pGreen, rectCANCWindowLegend)
        AddText(surf, rectCANCWindowLegend, GetDay("Cancel Date", ORDR_CANCEL_DATE), Brushes.Green, 9)

        ' Create Call-In Lead Time blocks Yellow

        Dim xxCALL1 As Int32 = 0
        Dim xxCALL2 As Int32 = 0

        Dim SHIP_CALL_IN_EARLIEST As Date = Nothing
        Dim SHIP_CALL_IN_LATEST As Date = Nothing

        If rowTL.Item("SHIP_CALL_IN_EARLIEST") & "" <> "" Then
            SHIP_CALL_IN_EARLIEST = rowTL.Item("SHIP_CALL_IN_EARLIEST")
            SHIP_CALL_IN_LATEST = rowTL.Item("SHIP_CALL_IN_LATEST")

            xxCALL1 = SHIP_CALL_IN_EARLIEST.Subtract(d1).TotalDays * timelineWidth / days
            If CALL_IN_NOTICE_DAYS <> 0 Then
                Dim rectLT1Window As New Rectangle(marginLeft + xxCALL1, timelineY - 0 + marginBot, (xxSHIPadj - xxCALL1), rectHeight)
                surf.FillRectangle(Brushes.Yellow, rectLT1Window)
                AddText(surf, rectLT1Window, "Notice" & vbCrLf & $"{CALL_IN_NOTICE_DAYS} wd", Brushes.Black, 9)
            End If

            xxCALL2 = SHIP_CALL_IN_LATEST.Subtract(d1).TotalDays * timelineWidth / days
            If CALL_IN_NOTICE_DAYS <> 0 Then
                Dim rectLT2Window As New Rectangle(marginLeft + xxCALL2, timelineY - rectHeight * 2 - marginTop, (yyCANCadj - xxCALL2), rectHeight * 2)
                surf.FillRectangle(Brushes.Yellow, rectLT2Window)
                AddText(surf, rectLT2Window, "Notice" & vbCrLf & $"{CALL_IN_NOTICE_DAYS} wd", Brushes.Black, 9)
            End If
        End If

        ' Create OPs Lead Time blocks Gray

        Dim dOPs_LT1 As Date = Nothing
        Dim dOPs_LT2 As Date = Nothing
        Dim yOps As Int32 = 0
        If rowTL.Item("SHIP_XMIT_EARLIEST") & "" <> "" Then
            dOPs_LT1 = rowTL.Item("SHIP_XMIT_EARLIEST")
            dOPs_LT2 = rowTL.Item("SHIP_XMIT_LATEST")

            yOps = timelineY - marginTop - rectHeight * 2 - (rectHeight + 10)
            Dim xWidth As Decimal = 2 * timelineWidth / days
            If CALL_IN_NOTICE_DAYS + OPS_LT1 <> 0 Then
                xx = dOPs_LT1.Subtract(d1).TotalDays * timelineWidth / days
                Dim rectOPsLT1Window As New Rectangle(marginLeft + xx, timelineY + 0 + marginBot, (xxCALL1 - xx), rectHeight)
                'Dim rectOPsLT1Window As New Rectangle(marginLeft + xx, timelineY + 0 + marginBot, xWidth, rectHeight)
                surf.FillRectangle(Brushes.Gray, rectOPsLT1Window)
                AddText(surf, rectOPsLT1Window, "OPs" & vbCrLf & $"{OPS_LT1} wd", Brushes.White, 9)
                surf.DrawLine(pDarkOrange, marginLeft + xx, yOps, marginLeft + xx, timelineY)

                Dim rectOpsLT1WindowLegend As New Rectangle(marginLeft + xx, yOps, legendWidth, rectHeight)
                surf.DrawRectangle(pDarkOrange, rectOpsLT1WindowLegend)
                AddText(surf, rectOpsLT1WindowLegend, GetDay("Earliest Xmit", dOPs_LT1), Brushes.DarkOrange, 9)
            End If

            If CALL_IN_NOTICE_DAYS + OPS_LT2 <> 0 Then
                xx = dOPs_LT2.Subtract(d1).TotalDays * timelineWidth / days
                Dim rectOpsLT2Window As New Rectangle(marginLeft + xx, timelineY - rectHeight * 2 - marginTop, (xxCALL2 - xx), rectHeight)
                'Dim rectOpsLT2Window As New Rectangle(marginLeft + xx, timelineY - rectHeight * 2 - marginTop, xWidth, rectHeight)
                surf.FillRectangle(Brushes.Gray, rectOpsLT2Window)
                AddText(surf, rectOpsLT2Window, "OPs" & vbCrLf & $"{OPS_LT2} wd", Brushes.White, 9)
                surf.DrawLine(pDarkOrange, marginLeft + xx, yOps, marginLeft + xx, timelineY)

                Dim rectOpsLT2WindowLegend As New Rectangle(marginLeft + xx, yOps, legendWidth, rectHeight)
                surf.DrawRectangle(pDarkOrange, rectOpsLT2WindowLegend)
                AddText(surf, rectOpsLT2WindowLegend, GetDay("Latest Xmit", dOPs_LT2), Brushes.DarkOrange, 9)
            End If
        End If

        ' Create Call-In Window & draw Blue Rectangle

        Dim daysCallInWindow As Int32 = SHIP_CALL_IN_LATEST.Subtract(SHIP_CALL_IN_EARLIEST).TotalDays
        xx = SHIP_CALL_IN_EARLIEST.Subtract(d1).TotalDays * timelineWidth / days
        yy = SHIP_CALL_IN_LATEST.Subtract(d1).TotalDays * timelineWidth / days

        Dim rectCallInWindow As New Rectangle(marginLeft + xx, timelineY - rectHeight - marginTop, (yy - xx), rectHeight)
        surf.FillRectangle(Brushes.Blue, rectCallInWindow)
        AddText(surf, rectCallInWindow, "Call-In Window" & $" {daysCallInWindow} cd", Brushes.White)

        ' Earliest / Latest Call-In Date flags

        If CALL_IN_NOTICE_DAYS <> 0 Then
            surf.DrawLine(pBlue, marginLeft + xx, timelineY, marginLeft + xx, timelineY + legendWidth + rectHeight)
            Dim rectCallinLT1WindowLegend As New Rectangle(marginLeft + xx, timelineY + legendWidth + marginBot, legendWidth, rectHeight)
            surf.DrawRectangle(pBlue, rectCallinLT1WindowLegend)
            AddText(surf, rectCallinLT1WindowLegend, GetDay("Earliest Call-In", SHIP_CALL_IN_EARLIEST), Brushes.Blue, 9)
        End If

        If CALL_IN_NOTICE_DAYS <> 0 Then
            surf.DrawLine(pBlue, marginLeft + yy, timelineY, marginLeft + yy, timelineY + legendWidth + rectHeight)
            Dim rectCallinLT2WindowLegend As New Rectangle(marginLeft + yy, timelineY + legendWidth + marginBot, legendWidth, rectHeight)
            surf.DrawRectangle(pBlue, rectCallinLT2WindowLegend)
            AddText(surf, rectCallinLT2WindowLegend, GetDay("Latest Call-In", SHIP_CALL_IN_LATEST), Brushes.Blue, 9)
        End If

        ' Draw blue daily ticks, wide red on Sunday

        Dim p As Pen
        Dim dx As Int32 = 0
        Dim tWidth As Int32 = 0


        Do
            dx += 1
            xx = dx * timelineWidth / days
            Dim d1x As Date = d1.AddDays(dx)
            If d1x.DayOfWeek = Global.System.DayOfWeek.Sunday Then
                p = pRed2
                tWidth = 2
            Else
                p = pBlue
                tWidth = 0
            End If

            xx = marginLeft + xx - CInt(tWidth / 2)
            surf.DrawLine(p, xx, timelineY + 1, xx, timelineY + 7)

            Dim dLegend As String = ""
            Dim rectDay As New Rectangle(xx - 2, timelineY + 3, 20, 20)
            If d1x.DayOfWeek = Global.System.DayOfWeek.Saturday Then
                dLegend &= "Sa"
            ElseIf d1x.DayOfWeek = Global.System.DayOfWeek.Sunday Then
                dLegend &= "Su"
            End If
            Dim rowSOTPARMH As DataRow = dst.Tables("SOTPARMH").Rows.Find(d1x)
            If rowSOTPARMH IsNot Nothing AndAlso rowSOTPARMH.Item("NATIONAL_HOLIDAY") & "" = "1" Then
                dLegend &= "H"
            End If
            If dLegend <> "" Then
                AddText(surf, rectDay, dLegend, Brushes.DarkOrange, 9)
            End If
        Loop Until dx + 1 >= days


        ' Relevant Call-In Date blue circles on Timeline
        ' Relevant Pick-Up Date red circles on Timeline

        Dim DayOfWeek As Int32 = 0
        dx = 0

        Do
            dx += 1
            xx = dx * timelineWidth / days

            Dim d1x As Date = d1.AddDays(dx)
            If d1x.DayOfWeek = Global.System.DayOfWeek.Sunday Then DayOfWeek = 1
            If d1x.DayOfWeek = Global.System.DayOfWeek.Monday Then DayOfWeek = 2
            If d1x.DayOfWeek = Global.System.DayOfWeek.Tuesday Then DayOfWeek = 3
            If d1x.DayOfWeek = Global.System.DayOfWeek.Wednesday Then DayOfWeek = 4
            If d1x.DayOfWeek = Global.System.DayOfWeek.Thursday Then DayOfWeek = 5
            If d1x.DayOfWeek = Global.System.DayOfWeek.Friday Then DayOfWeek = 6
            If d1x.DayOfWeek = Global.System.DayOfWeek.Saturday Then DayOfWeek = 7

            'If DayOfWeek <> 0 And d1x.Subtract(Now.Date).TotalDays >= 0 And SHIP_CALL_IN_LATEST.Subtract(d1x).TotalDays >= 0 Then
            If DayOfWeek <> 0 Then
                If Mid(CALL_IN_DAYS, DayOfWeek, 1) = "1" Then
                    surf.FillEllipse(Brushes.Blue, marginLeft + xx - 5, timelineY - 8, 10, 10)
                End If
            End If

            'If DayOfWeek <> 0 And d1x.Subtract(Now.Date).TotalDays >= 0 And SHIP_CALL_IN_LATEST.Subtract(d1x).TotalDays >= 0 Then
            If DayOfWeek <> 0 Then
                If Mid(PICK_UP_DAYS, DayOfWeek, 1) = "1" Then
                    surf.FillEllipse(Brushes.Red, marginLeft + xx - 5, timelineY + 0, 10, 10)
                End If
            End If
        Loop Until dx + 1 >= days

    End Sub
    Sub AddText(g As Graphics, rect As Rectangle, sText As String, brush As Brush, Optional fontSize As Int32 = 12)
        Using font As Font = New Font("Arial", fontSize, Drawing.FontStyle.Regular, GraphicsUnit.Point)

            Dim stringFormat As StringFormat = New StringFormat()
            stringFormat.Alignment = StringAlignment.Center
            stringFormat.LineAlignment = StringAlignment.Center
            g.DrawString(sText, font, brush, rect, stringFormat)
        End Using
    End Sub

    Private Sub grdSOTSHIPR_VIEW_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTSHIPR_VIEW.InitializeRow
        If Not e.Row.IsDataRow Then
            Exit Sub
        End If

        Dim SHIP_ROUTING_IND As String = e.Row.Cells("SHIP_ROUTING_IND").Value & ""
        If SHIP_ROUTING_IND = "1" Then
            e.Row.Cells("SHIP_ROUTING_IND").Appearance.ForeColor = System.Drawing.Color.Green
        Else
            e.Row.Cells("SHIP_ROUTING_IND").Appearance.ForeColor = System.Drawing.Color.Red
        End If

        Dim DEL_METHOD As String = e.Row.Cells("DEL_METHOD").Value & ""
        If DEL_METHOD = "SPS" Then
            e.Row.Cells("DEL_METHOD").Appearance.ForeColor = System.Drawing.Color.Red
        Else
            e.Row.Cells("DEL_METHOD").Appearance.ForeColor = System.Drawing.Color.Blue
        End If
    End Sub

    Private Sub btnDraw_Click(sender As Object, e As EventArgs) Handles btnDraw.Click
        Dim SHIP_BOL_NO As String = grdSOTSHIPR_VIEW.ActiveRow.Cells("SHIP_BOL_NO").Value
        Create_Timeline(SHIP_BOL_NO)
    End Sub

    Sub Send_emails_CARS_Errors(dt As DataTable)

        For Each rowType As DataRow In ASCDATA1.SelectDistinct(dt, "EVENT_ERROR_TYPE").Select("")
            Dim EVENT_ERROR_TYPE As String = rowType.Item("EVENT_ERROR_TYPE")
            Dim rowSOTSHIPM As DataRow = LookUp("SOTSHIPM", EVENT_ERROR_TYPE)

            Dim EMAIL_SUBJECT As String = "CARS Error - " & rowSOTSHIPM.Item("EVENT_ERROR_DESC") & ""
            Dim EMAIL_USER As String = rowSOTSHIPM.Item("EVENT_ERROR_EMAIL_TO") & ""

            If EMAIL_USER <> "" Then
                Dim EMAIL_BODY As String = ""
                For Each row As DataRow In dt.Select($"EVENT_ERROR_TYPE = '{EVENT_ERROR_TYPE}'", "EVENT_MSG")
                    Dim EVENT_DATE As String = Format(row.Item("EVENT_DATE"), "MM/dd/yy HH:mm:ss")
                    Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO") & ""
                    Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1_CARS").Rows.Find(SHIP_BOL_NO)
                    Dim SHIP_SHIP_DATE As Date = rowSOTSHIP1.Item("SHIP_SHIP_DATE")
                    Dim SHIP_CANCEL_DATE As Date = rowSOTSHIP1.Item("SHIP_CANCEL_DATE")
                    Dim SHIP_WINDOW As String = Format(SHIP_SHIP_DATE, "MM/dd") & " " & Format(SHIP_CANCEL_DATE, "MM/dd")
                    Dim EVENT_MSG As String = row.Item("EVENT_MSG") & ""
                    Dim SHIP_TO_CODE As String = row.Item("SHIP_TO_CODE") & ""
                    Dim CUST_CODE As String = row.Item("CUST_CODE") & ""
                    Dim SHIP_ADDR_CODE As String = row.Item("SHIP_ADDR_CODE") & ""
                    Dim CUST_STORE As String = CUST_CODE & " " & SHIP_ADDR_CODE

                    'EMAIL_BODY &= $"<br/><tr><td>{EVENT_DATE}</td><td>{SHIP_BOL_NO}</td><td>{EVENT_MSG}</td><td>{CUST_CODE}</td><td>{SHIP_ADDR_CODE}</td><td>{SHIP_TO_CODE}</td></tr>"
                    EMAIL_BODY &= $"<br/><tr><td>{EVENT_DATE}</td><td>{SHIP_WINDOW}</td><td>{EVENT_MSG}</td><td>{CUST_STORE}</td><td>{SHIP_TO_CODE}</td></tr>"
                Next

                Dim EVENT_ERROR_EMAIL_XNO As String = ASCMAIN1.Next_Control_No("SOTSHIPV.EVENT_ERROR_EMAIL_XNO")

                EMAIL_BODY = "<html><style>th, td {border:1px solid black;text-align: center;}</style><body>" & $"<h2>CARS email Control No {EVENT_ERROR_EMAIL_XNO}</h2>" &
                        "</br></br><table style='width:100%'><tr><th>Date Time</th><th>Ship Window</th><th>Message</th><th>Customer DC-Store</th><th>Ship-To Code</th></tr>" _
                        & EMAIL_BODY & "</br></table></body></html>"

                email_CARS_Errors(EMAIL_SUBJECT, EMAIL_USER, EMAIL_BODY)

                ASCMAIN1.sql = $"Update SOTSHIPV Set EVENT_ERROR_EMAIL_XNO = '{EVENT_ERROR_EMAIL_XNO}', EVENT_ERROR_EMAIL_DATE = SYSDATE, EVENT_ERROR_EMAIL_IND = '1'"
                ASCDATA1.ExecuteSQL()

            End If
        Next

    End Sub

    Sub email_CARS_Errors(EMAIL_SUBJECT As String, EMAIL_USER As String, EMAIL_BODY As String)

        ' PROBABLY NEED A TRY/CATCH IN HERE

        Dim em As New TAC.ASCNOTEE(ASCMAIN1.Folders, "CARS_ERROR", Nothing)
        em.CreateComponents()

        em.SetEmailSubject(EMAIL_SUBJECT)

        'If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then
        '    USER_EMAIL = "wjz@absolution.com"
        'End If

        em.SetEmailTo(EMAIL_USER)
        em.SetDocumentBody(EMAIL_BODY)

        ' em.SaveEmail = True
        em.EmailDocument()

    End Sub

    Private Sub grdSOTSHIP1_MATRIX_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTSHIP1_MATRIX.InitializeRow
        Dim CALL_IN_DAY As String = e.Row.Cells("CALL_IN_DAY").Value
        If CALL_IN_DAY = "Sun" Then
            e.Row.Cells("CALL_IN_DAY").Appearance.ForeColor = System.Drawing.Color.Red
        Else
            e.Row.Cells("CALL_IN_DAY").Appearance.ForeColor = System.Drawing.Color.Empty
        End If

    End Sub

    Private Sub btnDraw_DragEnter(sender As Object, e As DragEventArgs) Handles btnDraw.DragEnter

    End Sub

#End Region

End Class