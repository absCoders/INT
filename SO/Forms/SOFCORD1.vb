Imports Infragistics.UltraChart.Resources
Imports Infragistics.UltraChart.Resources.Appearance
Imports Infragistics.UltraChart.Core
Imports Infragistics.UltraChart.Core.ColorModel
Imports Infragistics.UltraChart.Data
Imports Infragistics.UltraChart.Core.Layers
Imports Infragistics.UltraChart.Core.Primitives
Imports Infragistics.UltraChart.Shared.Styles

Imports Infragistics.Win

Public Class SOFCORD1

#Region "Declarations"
    Dim CUST_CODE As String
    Dim rowARTCUST1 As DataRow
    Dim sqlSOTORDR0 As String
    Dim sqlSOTPICK2 As String
    Dim sqlSOTORDRS As String
    Dim sqlSOTRSRVS As String
    Dim ORDR_GROUP_NO As String
    Private sqlSOTSHIPX As String = String.Empty
    Dim SOTORDR0 As String
    Dim SOTORDR0_ERRORS As String
    Dim ORDR_REL_HOLD_CODES_list As New Dictionary(Of String, String)
    Dim SOTORDRT As String

    Private lstOrderGroupNos As New List(Of String)
    Private lstAlreadySent As New List(Of String)
    Private num855Selected As Int16 = 0

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Get_PARM("ARTPARM1")

        AUDIT.Add("SOTORDR0", "*")

        With dst

            sqlSOTORDR0 = "Select SOTORDR0.ORDR_GROUP_NO, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO, TRUNC(EDT850T1.EDI_RECEIVED_DATE) EDI_RECEIVED_DATE
                                , SOTORDR0.CUST_DC_NO, SOTORDR0.ORDR_DEPT, SOTORDR0.SALES_DIVISION_CODE, SOTORDR0.ORDR_DATE
                                , SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE, SOTORDR0.ORDR_ORIG_SHIP_DATE, SOTORDR0.ORDR_ORIG_CANCEL_DATE
                                , SOTORDR0.WHSE_CODE, SOTORDR0.ORDR_SOURCE
                                , SOTORDR0.ORDR_AMT, SOTORDR0.ORDR_AMT_OPEN, SOTORDR0.ORDR_AMT_PICK, SOTORDR0.ORDR_AMT_SHIP, SOTORDR0.ORDR_AMT_CANC
                                , SOTORDR0.ORDR_QTY, SOTORDR0.ORDR_QTY_OPEN, SOTORDR0.ORDR_QTY_PICK, SOTORDR0.ORDR_QTY_SHIP, SOTORDR0.ORDR_QTY_CANC
                                , SOTORDR0.ORDR_CNT, SOTORDR0.ORDR_CNT_OPEN, SOTORDR0.ORDR_CNT_PICK, SOTORDR0.ORDR_NO_MIN, SOTORDR0.ORDR_ARRIVAL_DATE, SOTORDR0.ORDR_ALLO_DATE
                                , DECODE(NVL(E1.EDI_STATUS, NULL), NULL, '0', '1') CUST_855, E8.SENT_855, SOTORDR0.EDI_DOC_SEQ_NO
                                , SOTORDR0.ORDR_INTERNAL_NOTES
                                , SOTORDR1.ORDR_HOLD, SOTORDR1.REVERSE_PO
                                 from SOTORDR0, SOTORDR1, EDT850T1,
                                 (SELECT * FROM EDTTRPM1 WHERE EDI_DOC_NO = '855') E1,
                                 (SELECT ORDR_GROUP_NO, MAX(INIT_DATE) SENT_855 FROM EDT855O1 GROUP BY ORDR_GROUP_NO) E8
                                 where SOTORDR1.ORDR_NO (+) = SOTORDR0.ORDR_NO_MIN 
                                    and SOTORDR0.CUST_CODE = E1.CUST_CODE(+)
                                    and SOTORDR0.ORDR_GROUP_NO = E8.ORDR_GROUP_NO (+)
                                    and SOTORDR1.EDI_DOC_SEQ_NO = EDT850T1.EDI_DOC_SEQ_NO (+)"

            SOTORDR0 = ASCMAIN1.Temp_Table(sqlSOTORDR0 & " and ROWNUM < 1")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add Primary Key (ORDR_GROUP_NO)")

            ASCMAIN1.sql = "Select SOTORDR1.ORDR_GROUP_NO, SOTPICK1.ERROR_REASON ERRORS from SOTORDR1,SOTPICK1 where ROWNUM < 1"
            SOTORDR0_ERRORS = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0_ERRORS & " Add Primary Key (ORDR_GROUP_NO)")

            SOTORDR0_ERRORS = ASCMAIN1.Temp_Table

            ASCMAIN1.sql = "Select X.*, ARTCUST1.CUST_NAME, SOTORDR1.CUST_STORE_NO, NVL(SOTORDR1.CUST_STORE_LOCATION,SOTORDR5.CUST_NAME) CUST_STORE_LOCATION, SOTORDR0_ERRORS.ERRORS, SOTORDR1.ORDR_HIGH_PRIORITY, SOTORDR1.ORDR_HIGH_PRIORITY_NOTE" & vbCrLf _
                & " from " & SOTORDR0 & " X, SOTORDR1, ARTCUST1, SOTORDR5, " & SOTORDR0_ERRORS & " SOTORDR0_ERRORS" & vbCrLf _
                & " where SOTORDR1.ORDR_NO = X.ORDR_NO_MIN and ARTCUST1.CUST_CODE = X.CUST_CODE" & vbCrLf _
                & "   and SOTORDR5.ORDR_NO (+) = X.ORDR_NO_MIN" & vbCrLf _
                & "   and SOTORDR5.CUST_ADDR_TYPE (+) = 'ST'" & vbCrLf _
                & "   and SOTORDR0_ERRORS.ORDR_GROUP_NO (+) = X.ORDR_GROUP_NO"
            Create_TDA(.Tables.Add, "SOTORDR0", "**", 0, True, "V", 1, "ORDR_INTERNAL_NOTES")

            sqlSOTSHIPX = "SELECT DISTINCT ORDR_GROUP_NO FROM SOTSHIP1 WHERE SHIP_STATUS = 'P' AND LP_STATUS = '1'"
            Create_TDA(.Tables.Add, "SOTSHIPX", sqlSOTSHIPX, 0, False, String.Empty, 1)

            Dim TBL As DataTable = .Tables("SOTORDR0").Clone
            TBL.TableName = "SOTCORDG"
            .Tables.Add(TBL)

            ASCMAIN1.sql = "Select SOTORDR1.ORDR_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_DATE" & vbCrLf _
                & ", SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE, SOTORDR1.ORDR_ORIG_SHIP_DATE, SOTORDR1.ORDR_ORIG_CANCEL_DATE" & vbCrLf _
                & ", SOTORDR1.CUST_STORE_NO, SOTORDR1.SALES_DIVISION_CODE, SOTORDR1.ORDR_SOURCE, SOTORDR1.ORDR_DEPT, SOTORDR1.ORDR_ADDR_TYPE_ST" & vbCrLf _
                & ", SOTORDR1.ORDR_STATUS, NVL(SOTORDR1.CUST_STORE_LOCATION, SOTORDR5.CUST_NAME) CUST_STORE_LOCATION, SOTORDR1.SREP_CODE, SOTORDR1.ORDR_PRIORITY, SOTORDR1.ORDR_HOLD" & vbCrLf _
                & ", SOTORDR1.ORDR_REL_HOLD_CODES, SOTORDR1.CUST_DC_NO, SOTORDR1.WHSE_CODE, SOTORDR1.EDI_DOC_SEQ_NO, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.SHIP_VIA_CODE, SOTORDR1.ORDR_ARRIVAL_DATE, SOTORDR1.ORDR_ALLO_DATE" & vbCrLf _
                & " from SOTORDR1, SOTORDR5 where SOTORDR1.ORDR_GROUP_NO = :PARM1 and SOTORDR1.ORDR_STATUS <> 'D'" _
                & " and SOTORDR5.ORDR_NO (+) = SOTORDR1.ORDR_NO and SOTORDR5.CUST_ADDR_TYPE (+) = 'ST'"

            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False, "V", 1)

            sqlSOTORDRS = "Select SOTORDR2.ITEM_CODE, SOTORDR2.ITEM_DESC, ICTITEM1.ITEM_EAN_CODE" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY * ICTITEM1.ITEM_COST_STD) ORDR_CGS" & vbCrLf _
                & " from SOTORDR2,SOTORDR1,ICTITEM1" & vbCrLf _
                & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE" & vbCrLf _
                & "   and SOTORDR1.ORDR_STATUS <> 'D'" & vbCrLf _
                & " group by SOTORDR2.ITEM_CODE, SOTORDR2.ITEM_DESC, ICTITEM1.ITEM_EAN_CODE"
            ASCMAIN1.sql = Replace(sqlSOTORDRS, " group by ", " and ROWNUM < 1 group by ")
            Create_TDA(.Tables.Add, "SOTORDRS", "**", 0, False, "", 0)
            .Tables("SOTORDRS").Columns.Add("ORDR_UNIT_PRICE", GetType(System.Decimal), "IIF(ISNULL(ORDR_QTY,0)=0,0,ISNULL(ORDR_AMT,0) / ISNULL(ORDR_QTY,0))")
            .Tables("SOTORDRS").Columns.Add("ORDR_GP", GetType(System.Decimal), "ISNULL(ORDR_AMT,0)-ISNULL(ORDR_CGS,0)")
            .Tables("SOTORDRS").Columns.Add("ORDR_GP_PCT", GetType(System.Decimal), "IIF(ISNULL(ORDR_AMT,0)=0,0,100*ORDR_GP/ISNULL(ORDR_AMT,0))")
            .Tables("SOTORDRS").Columns.Add("ORDR_AMT_CANC", GetType(System.Decimal), "ISNULL(ORDR_QTY_CANC,0)*ISNULL(ORDR_UNIT_PRICE,0)")

            ASCMAIN1.sql = "Select SOTPICK1.*, SOTORDR1.CUST_STORE_NO, SOTORDR1.CUST_STORE_LOCATION CUST_STORE_NAME" & vbCrLf _
                & ", SOTPICK0.PICK_FORCED" & vbCrLf _
                & " from SOTPICK1,SOTORDR1,SOTPICK0 " & vbCrLf _
                & " where SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTPICK0.PICK_BATCH_NO = SOTPICK1.PICK_BATCH_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_GROUP_NO = :PARM1" & vbCrLf _
                & "   and SOTORDR1.ORDR_STATUS <> 'D'"

            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTPICK2.*, SOTORDR2.ITEM_CODE" & vbCrLf _
                & ", SOTORDR2.ITEM_DESC, SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & " from SOTPICK1,SOTPICK2,SOTORDR2" & vbCrLf _
                & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
                & "   and SOTPICK2.PICK_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICK2", "**", 0, False, "V", 2)

            Create_Relation("SOTPICK1", "SOTPICK2", "PICK_NO")

            ASCMAIN1.sql = "Select SOTSHIP1.SHIP_BOL_NO,SOTSHIP1.SHIP_DATE_SHIPPED,SOTSHIP1.SHIP_VIA_CODE,SOTSHIP1.SHIP_REF" & vbCrLf _
                & ",SOTSHIP1.SHIP_TOTAL_WGT,SOTSHIP1.SHIP_CNT_CARTONS,SOTSHIP1.SHIP_ADDR_TYPE,SOTSHIP1.SHIP_ADDR_CODE" & vbCrLf _
                & ",SHIP_PICK_PRINTED,PICK_BATCH_NO,SHIP_STATUS,LP_STATUS" & vbCrLf _
                & ",SOTSHIP1.BILL_OF_LADING_NO,SOTSHIP1.FRT_TERMS,SOTSHIP1.SHIP_PULL_BY_STYLE" & vbCrLf _
                & ",SOTSHIP1.SHIP_856_BATCH_NO,SOTSHIP1.SHIP_810_BATCH_NO,SOTSHIP1.WHSE_CODE,SOTSHIP1.INV_DATE,SOTSHIP1.SHIP_MANIFEST_NO" & vbCrLf _
                & ",SOTSHIP1.SHIP_BOL_NO_REV,SOTSHIP1.SHIP_NOTES,SOTSHIP1.SHIPPED_ACTUAL,SOTSHIP1.SHIP_SEAL_NO" & vbCrLf _
                & ",SOTSHIP1.SHIP_BOL_NO_ORIG,SOTSHIP1.SHIP_BOL_NO_SPLIT,SOTSHIP1.BOL_PRINTED,SOTSHIP1.SHIP_SPEC_INST" & vbCrLf _
                & ",SOTSHIP1.MASTER_SHIP_BOL_NO,SOTSHIP1.SHIP_940_BATCH_NO,SOTSHIP1.SHIP_753_IND,SOTSHIP1.SHIP_DATE_PACKED" & vbCrLf _
                & ",SOTSHIP1.INIT_DATE,SOTSHIP1.INIT_OPER, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
                & " from SOTSHIP1, SOTORDR0" & vbCrLf _
                & " where SOTSHIP1.ORDR_GROUP_NO = :PARM1 AND SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO"
            ASCMAIN1.sql = $"Select X.*, P.ORDR_NO_MAX, P.PICKS, SOTORDR5.CUST_NAME from ({ASCMAIN1.sql}) X
, (SELECT SOTPICK1.SHIP_BOL_NO, MAX (SOTPICK1.ORDR_NO) ORDR_NO_MAX, COUNT (*) PICKS FROM SOTPICK1 GROUP BY SOTPICK1.SHIP_BOL_NO) P
, SOTORDR5
WHERE P.SHIP_BOL_NO (+) = X.SHIP_BOL_NO
AND SOTORDR5.ORDR_NO (+) = P.ORDR_NO_MAX
AND SOTORDR5.CUST_ADDR_TYPE (+) = 'ST'"
            Create_TDA(.Tables.Add, "SOTSHIP1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTPICK1.SHIP_BOL_NO" _
                & ", SOTORDR2.ITEM_CODE, SOTORDR2.ITEM_DESC" _
                & ", SUM (SOTPICK2.PICK_QTY) PICK_QTY" _
                & ", SUM (SOTPICK2.PICK_QTY * SOTPICK2.PICK_UNIT_PRICE) PICK_AMT" _
                & ", SUM (SOTPICK2.PICK_QTY_CONF) PICK_QTY_CONF" _
                & ", SUM (SOTPICK2.PICK_QTY_CANC) PICK_QTY_CANC" _
                & ", SUM (SOTPICK2.PICK_QTY_BACK) PICK_QTY_BACK" _
                & ", SUM (SOTPICK2.PICK_QTY_CANC_REL) PICK_QTY_CANC_REL" _
                & ", SUM (SOTPICK2.PICK_QTY_BACK_REL) PICK_QTY_BACK_REL" _
                & " from SOTPICK1,SOTPICK2,SOTORDR2" _
                & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" _
                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" _
                & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" _
                & "   and SOTPICK1.SHIP_BOL_NO = :PARM1" _
                & " group by SOTPICK1.SHIP_BOL_NO" _
                & ", SOTORDR2.ITEM_CODE, SOTORDR2.ITEM_DESC"
            Create_TDA(.Tables.Add, "SOTSHIP2", "**", 0, False, "V", 0)
            .Tables("SOTSHIP2").Columns.Add("PICK_UNIT_PRICE", GetType(System.Decimal), "IIF(PICK_QTY=0,0,PICK_AMT/PICK_QTY)")
            With .Tables("SOTSHIP2")
                .Columns("PICK_QTY").DataType = GetType(System.Int64)
                .Columns("PICK_QTY_CONF").DataType = GetType(System.Int64)
                .Columns("PICK_QTY_CANC").DataType = GetType(System.Int64)
                .Columns("PICK_QTY_BACK").DataType = GetType(System.Int64)
                .Columns("PICK_QTY_CANC_REL").DataType = GetType(System.Int64)
                .Columns("PICK_QTY_BACK_REL").DataType = GetType(System.Int64)
            End With

            Create_Relation("SOTSHIP1", "SOTSHIP2", "SHIP_BOL_NO")

            With .Tables.Add("SOTCORDR")
                .Columns.Add("ITEM_CODE")
                .Columns.Add("ITEM_DESC")
                .Columns.Add("ORDR", GetType(System.Int64))
                .Columns.Add("OPEN", GetType(System.Int64))
                .Columns.Add("PICK", GetType(System.Int64))
                .Columns.Add("SHIP", GetType(System.Int64))
                .Columns.Add("CANC", GetType(System.Int64))
                .Columns.Add("ORDR_AMT", GetType(System.Decimal))

                .Columns("ORDR").DefaultValue = 0
                .Columns("OPEN").DefaultValue = 0
                .Columns("PICK").DefaultValue = 0
                .Columns("SHIP").DefaultValue = 0
                .Columns("CANC").DefaultValue = 0
                .Columns("ORDR_AMT").DefaultValue = 0

                .PrimaryKey = New DataColumn() { .Columns("ITEM_CODE")}
            End With

            With .Tables.Add("SOTORDRX")
                .Columns.Add("ORDR_NO")
                .Columns.Add("CUST_STORE_NO")
            End With

            With .Tables.Add("SOTORDRM")
                .Columns.Add("ITEM_CODE")
                .Columns.Add("QTY")
            End With

            With .Tables.Add("SOTCORDX")
                .Columns.Add("SORT_SEQ")
                .Columns.Add("CODE_VALUE")
                Dim TOT As String = ""
                Dim YTD As String = ""
                For I As Integer = 1 To 12
                    Dim TC As String = "V" & Format(I, "00")
                    .Columns.Add(TC, GetType(System.Decimal))
                    TOT &= "+ISNULL(" & TC & ",0)"
                    If I > 12 - Val(Mid(ASCMAIN1.CYP, 5, 2)) Then
                        YTD &= "+ISNULL(" & TC & ",0)"
                    End If
                Next
                .Columns.Add("TOT", GetType(System.Decimal), TOT)
                .Columns.Add("YTD", GetType(System.Decimal), YTD)
                .Columns.Add("TOTPCT", GetType(System.Decimal), "TOT / 1")
                .Columns.Add("YTDPCT", GetType(System.Decimal), "YTD / 1")
                .PrimaryKey = New DataColumn() { .Columns("SORT_SEQ")}
            End With

            Dim T As DataTable = .Tables("SOTCORDX").Clone
            T.TableName = "SOTCORDD"
            T.PrimaryKey = New DataColumn() {T.Columns("SORT_SEQ"), T.Columns("CODE_VALUE")}
            .Tables.Add(T)

            ASCMAIN1.sql = "Select SOTINVH1.*" & vbCrLf _
                & " from SOTINVH1 " & vbCrLf _
                & " where SOTINVH1.CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False, "V", 2)

            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1)

            ASCMAIN1.sql = "SELECT SOTINVH2.INV_NO, SOTINVH2.ITEM_CODE" _
                & ", SOTINVH2.ORDR_UNIT_PRICE, SOTINVH2.ORDR_QTY_SHIP" _
                & ", SOTINVH1.ORDR_CUST_PO, SOTINVH1.ORDR_NO, SOTINVH1.CUST_STORE_NO, SOTINVH1.WHSE_CODE" _
                & " from SOTINVH2,SOTINVH1,SOTORDR1" _
                & " where SOTINVH2.CUST_CODE = :PARM1" _
                & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" _
                & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" _
                & "   and SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO" _
                & "   and SOTINVH2.ITEM_CODE = :PARM2 AND SOTINVH2.ORDR_YYYYPP_UPDATED = :PARM3"
            Create_TDA(.Tables.Add, "SOTCORDY", "**", 0, False, "VVV", 0)
            .Tables("SOTCORDY").Columns.Add("AMT", GetType(System.Decimal), "ORDR_QTY_SHIP * ORDR_UNIT_PRICE")

            ASCMAIN1.sql = "Select SOTCART1.*,SOTPICK1.SHIP_BOL_NO,SOTSHIP1.SHIP_ADDR_TYPE,SOTSHIP1.SHIP_ADDR_CODE, SOTORDR1.ORDR_CUST_PO" _
                & " from SOTCART1,SOTPICK1,SOTSHIP1, SOTORDR1" _
                & " where SOTPICK1.PICK_NO = SOTCART1.PICK_NO" _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" _
                & "   and SOTSHIP1.ORDR_GROUP_NO = :PARM1" _
                & "   and SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO"
            Create_TDA(.Tables.Add, "SOTCART1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTCART2.*" _
                & " from SOTCART2,SOTCART1,SOTPICK1,SOTSHIP1" _
                & " where SOTCART1.CART_NO = SOTCART2.CART_NO" _
                & "   and SOTPICK1.PICK_NO = SOTCART1.PICK_NO" _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" _
                & "   and SOTSHIP1.ORDR_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTCART2", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select SOTCART3.*, LOT_EXPIRATION_DATE LOT_EXP, LOT_FIFO_DATE LOT_FIFO
                                from SOTCART3, SOTCART2, SOTCART1, SOTPICK1, SOTSHIP1
                                where SOTCART3.CART_NO = SOTCART2.CART_NO
                                AND SOTCART2.CART_LNO=SOTCART3.CART_LNO
                                AND SOTCART1.CART_NO = SOTCART2.CART_NO
                                and SOTPICK1.PICK_NO = SOTCART1.PICK_NO
                                and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO
                                and SOTSHIP1.ORDR_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTCART3", "**", 0, False, "V", 3)

            Create_Relation("SOTCART1", "SOTCART2", "CART_NO")
            Create_Relation("SOTCART2", "SOTCART3", "CART_NO,CART_LNO")

            ASCMAIN1.sql = "Select SOTCART2.*" & vbCrLf _
                & ",ICTITEM1.ITEM_DESC" & vbCrLf _
                & ", SOTCART2.QTY_PACKED ORDR_QTY, SOTCART2.QTY_PACKED ORDR_QTY_CANC" & vbCrLf _
                & " from SOTCART2,ICTITEM1 where ROWNUM < 1"
            SOTORDRT = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDRT & " Add Primary Key (CART_NO, CART_LNO)")
            ASCMAIN1.sql = "Select SOTCART2.*" & vbCrLf _
                & ",SOTCART1.CART_TRACKING_NO, SOTCART1.PICK_NO, SOTCART1.CART_TOTAL_UNITS,SOTCART1.CART_TOTAL_WGT_ACTUAL" & vbCrLf _
                & ",SOTORDR1.ORDR_SHIP_DATE,SOTORDR1.ORDR_STATUS,SOTORDR1.CUST_STORE_NO,SOTORDR1.ORDR_CUST_PO,NVL(SOTORDR1.CUST_STORE_LOCATION,SOTSELL1.SELL_NAME) CUST_STORE_LOCATION,SOTINVH1.INV_DATE, SOTSELL1.SELL_NAME, SOTSHIP1.SHIP_REF" & vbCrLf _
                & " from SOTSELL1," & SOTORDRT & " SOTCART2,SOTCART1,SOTORDR1,SOTINVH1,SOTSHIP1" & vbCrLf _
                & "where SOTCART1.CART_NO (+) = SOTCART2.CART_NO" & vbCrLf _
                & "  and SOTORDR1.ORDR_NO = SOTCART2.ORDR_NO" & vbCrLf _
                & "  and SOTINVH1.ORDR_NO = SOTCART2.ORDR_NO" & vbCrLf _
                & "  and SOTSELL1.SELL_CODE (+) = SOTORDR1.SELL_CODE" & vbCrLf _
                & "  and SOTSHIP1.SHIP_BOL_NO (+) = SOTINVH1.SHIP_BOL_NO"
            Dim sqlSOTORDRT As String = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SOTORDRT", "**", 0, False, "", 0)

            ASCMAIN1.sql = "SELECT CART_NO, ORDR_NO," & vbCrLf _
                & "CASE WHEN CART_TRACKING_NO IS NULL AND SHIP_REF IS NOT NULL AND QTY_PACKED <> 0 THEN SHIP_REF ELSE CART_TRACKING_NO END CART_TRACKING_NO," _
                & "PICK_NO," & vbCrLf _
                & "SUM (QTY_PACKED) QTY_PACKED, COUNT (*) ITEMS," & vbCrLf _
                & "MIN (ITEM_CODE) ITEM1, MAX (ITEM_CODE) ITEM2," & vbCrLf _
                & "CART_TOTAL_UNITS, CART_TOTAL_WGT_ACTUAL," & vbCrLf _
                & "SUM (ORDR_QTY) ORDR_QTY, SUM (ORDR_QTY_CANC) ORDR_QTY_CANC," & vbCrLf _
                & "ORDR_SHIP_DATE, ORDR_STATUS, CUST_STORE_NO, ORDR_CUST_PO, CUST_STORE_LOCATION, INV_DATE, SELL_NAME, SHIP_REF" & vbCrLf _
                & " from ( " & sqlSOTORDRT & " ) " & vbCrLf _
                & " group by CART_NO, ORDR_NO, " & vbCrLf _
                & "CASE WHEN CART_TRACKING_NO IS NULL AND SHIP_REF IS NOT NULL AND QTY_PACKED <> 0 THEN SHIP_REF ELSE CART_TRACKING_NO END," & vbCrLf _
                & "PICK_NO, ORDR_SHIP_DATE, CART_TOTAL_UNITS, CART_TOTAL_WGT_ACTUAL," _
                & "ORDR_STATUS, CUST_STORE_NO, ORDR_CUST_PO, CUST_STORE_LOCATION, INV_DATE, SELL_NAME, SHIP_REF"
            Create_TDA(.Tables.Add, "SOTORDRU", "**", 0, False, "", 0)
            With dst.Tables("SOTORDRU")
                .Columns("QTY_PACKED").DataType = GetType(System.Int64)
                .Columns("ITEMS").DataType = GetType(System.Int64)
                .Columns("CART_TOTAL_UNITS").DataType = GetType(System.Int64)
                .Columns("ORDR_QTY").DataType = GetType(System.Int64)
                .Columns("ORDR_QTY_CANC").DataType = GetType(System.Int64)
            End With

            With .Tables.Add("SOTORDCC")
                .Columns.Add("ORDR_NO")
                .Columns.Add("TOT_CTNS", GetType(System.Int64))
                .Columns.Add("TOT_WGT", GetType(System.Decimal))
            End With

            Create_TDA(.Tables.Add, "EDTTRPM1", "*")
            Fill_Records("EDTTRPM1", String.Empty, True, "SELECT * FROM EDTTRPM1 WHERE EDI_DOC_NO = '855'")

        End With

        grdSOTORDRT.DataSource = dst.Tables("SOTORDRT")
        grdSOTORDRU.DataSource = dst.Tables("SOTORDRU")
        grdSOTORDR0.DataSource = dst.Tables("SOTORDR0")
        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdSOTSHIP1.DataSource = dst.Tables("SOTSHIP1")
        grdSOTORDRS.DataSource = dst.Tables("SOTORDRS")
        grdSOTCORDX.DataSource = dst.Tables("SOTCORDX")
        grdSOTCORDD.DataSource = dst.Tables("SOTCORDD")
        grdSOTCORDY.DataSource = dst.Tables("SOTCORDY")
        grdSOTORDR1.DataSource = dst.Tables("SOTORDR1")
        grdSOTORDRX.DataSource = grdSOTORDR1.DataSource
        grdSOTORDRM.DataSource = dst.Tables("SOTORDRM")
        grdSOTCART1.DataSource = dst.Tables("SOTCART1")
        grdSOTORDCC.DataSource = dst.Tables("SOTORDCC")

        grdSOTORDR1.DisplayLayout.UseFixedHeaders = True
        With grdSOTORDR1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ORDR_NO", "CUST_STORE_NO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        grdSOTORDR0.DisplayLayout.UseFixedHeaders = True
        With grdSOTORDR0.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_CODE", "CUST_NAME", "ORDR_GROUP_NO", "ORDR_CUST_PO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdSOTORDR0.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "ORDR_INTERNAL_NOTES" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If

                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

                If New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"ORDR_CNT", "ORDR_CNT_OPEN", "ORDR_CNT_PICK"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"ORDR_DATE", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "ORDR_ORIG_SHIP_DATE", "ORDR_ORIG_CANCEL_DATE", "ORDR_ARRIVAL_DATE", "ORDR_ALLO_DATE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                ElseIf New String() {"ORDR_CUST_PO", "CUST_DC_NO", "ORDR_DEPT", "WHSE_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Yellow
                ElseIf New String() {"ORDR_SOURCE", "CUST_STORE_NO", "CUST_STORE_LOCATION", "ORDR_NO_MIN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                ElseIf New String() {"CUST_CODE", "CUST_NAME", "ORDR_GROUP_NO", "SALES_DIVISION_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf New String() {"CUST_855", "SENT_855"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightSeaGreen
                ElseIf New String() {"ORDR_INTERNAL_NOTES"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.PaleVioletRed
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        Create_Summary(grdSOTORDR0, "ORDR_GROUP_NO", "Count")
        Create_Summary(grdSOTORDR0, New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC", "ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "ORDR_CNT", "ORDR_CNT_OPEN", "ORDR_CNT_PICK"}, , , "#,##0")

        Create_Summary(grdSOTORDRS, "ITEM_CODE", "Count")
        Create_Summary(grdSOTORDRS, New String() {"ORDR_QTY", "ORDR_AMT", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "ORDR_AMT_CANC"})

        ' Create_Summary(grdSOTCART1, "CART_NO", "Count")

        Create_Summary(grdSOTORDRT, "ORDR_NO", "Count")
        Create_Summary(grdSOTORDRT, New String() {"ORDR_QTY", "ORDR_QTY_CANC", "QTY_PACKED"})
        Show_Filter(grdSOTORDRT, True)
        grdSOTORDRT.DisplayLayout.GroupByBox.Hidden = False

        Create_Summary(grdSOTORDCC, "ORDR_NO", "Count")
        Create_Summary(grdSOTORDCC, New String() {"TOT_CTNS", "TOT_WGT"})
        Show_Filter(grdSOTORDCC, True)
        grdSOTORDCC.DisplayLayout.GroupByBox.Hidden = False
        grdSOTORDCC.DisplayLayout.Bands(0).Columns("TOT_CTNS").Format = "#,##0"
        grdSOTORDCC.DisplayLayout.Bands(0).Columns("TOT_WGT").Format = "#,##0.00"
        grdSOTORDCC.DisplayLayout.Override.RowSelectorWidth = 20
        grdSOTORDCC.DisplayLayout.ScrollBarLook.VerticalScrollBarWidth = 20

        Create_Summary(grdSOTORDRU, "ORDR_NO", "Count")
        Create_Summary(grdSOTORDRU, New String() {"ORDR_QTY", "ORDR_QTY_CANC", "QTY_PACKED", "CART_TOTAL_UNITS", "CART_TOTAL_WGT_ACTUAL"})
        Show_Filter(grdSOTORDRU, True)
        grdSOTORDRU.DisplayLayout.GroupByBox.Hidden = False

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTCORDD, grdSOTCORDX}
            grd.DisplayLayout.UseFixedHeaders = True
            With grd.DisplayLayout.Bands(0)
                For I As Integer = 1 To 12
                    .Columns("V" & Format(I, "00")).Format = "#,##0"
                Next
                .Columns("TOT").Format = "#,##0"
                .Columns("YTD").Format = "#,##0"
                .Columns("TOTPCT").Format = "#,##0.0"
                .Columns("YTDPCT").Format = "#,##0.0"
                For Each COLUMN_NAME As String In New String() {"CODE_VALUE", "TOT", "YTD", "TOTPCT", "YTDPCT"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                Next
            End With
        Next
        Create_Summary(grdSOTCORDD, "CODE_VALUE", "Count")
        Create_Summary(grdSOTCORDD, New String() {"TOT", "TOTPCT", "YTD", "YTDPCT", "V01", "V02", "V03", "V04", "V05", "V06", "V07", "V08", "V09", "V10", "V11", "V12"})

        Create_Summary(grdSOTCORDY, "INV_NO", "Count")
        Create_Summary(grdSOTCORDY, New String() {"ORDR_QTY_SHIP", "AMT"})


        Show_Filter(grdSOTORDR0, True)
        grdSOTORDR0.DisplayLayout.GroupByBox.Hidden = False

        ASCMAIN1.Add_Value_List(grdSOTORDRT, "ORDR_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTORDRU, "ORDR_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTPICK1, "PICK_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTSHIP1, "SHIP_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTSHIP1, "LP_STATUS", Nothing, New String() {":", "0:Pending Xmit", "1:Transmitted", "2:Shipped", "3:Confirmed", "V:Xmit/De-Rel", "D:De-Released"})
        ASCMAIN1.Add_Value_List(grdSOTORDR0, "ORDR_SOURCE", Nothing, New String() {":", "K:Keyed", "E:EDI", "W:Web", "S:SRep", "G:Gratis"})

        '  Set_cmbYP("RYP_TO", ASCMAIN1.CYP, -36, 12, 0)
        cmb12Months.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cmb12Months.SelectedRow = cmb12Months.Rows(0)

        ASCMAIN1.sql = "Select * from ASTCODE1 where TABLE_NAME = 'SOTORDR1' and COLUMN_NAME = 'ORDR_REL_HOLD_CODES'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "T_CODE")
            ORDR_REL_HOLD_CODES_list.Add(row.Item("T_CODE"), row.Item("T_DESC"))
        Next

        dteOSFrom.Value = Now.Date.AddMonths(-1)
        dteOSTo.Value = Now.Date

        If ASCMAIN1.CLIENT = "INT" Then
            tabMain.Tabs("Status && Tracking").Visible = True
            With grdSOTORDRT.DisplayLayout.Bands(0)
                .Columns("CUST_STORE_LOCATION").Header.Caption = "Name"
                .Columns("CUST_STORE_NO").Hidden = True
            End With
            With grdSOTORDRU.DisplayLayout.Bands(0)
                .Columns("CUST_STORE_LOCATION").Header.Caption = "Name"
                .Columns("CUST_STORE_NO").Hidden = True
            End With
        Else
            tabMain.Tabs("Status && Tracking").Visible = False
            With grdSOTORDRT.DisplayLayout.Bands(0)
                .Columns("SELL_NAME").Hidden = True
            End With
            With grdSOTORDRU.DisplayLayout.Bands(0)
                .Columns("SELL_NAME").Hidden = True
            End With
        End If

        grdSOTORDRS.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.CellSelect
        grdSOTORDR1.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.CellSelect

        MakeTransparent(chkEditInternalNotes)

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Select"

                'If ASCMAIN1.Running_in_VS Then
                '    Get850Data()
                '    Exit Sub
                'End If

                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Customer"
                Else
                    'Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                    'rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE) ' Absx1.txtFor("CUST_CODE").Text)
                    rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 IsNot Nothing Then
                        If Trim(ASCMAIN1.USER_CODES) = "FS" Then
                            If rowARTCUST1.Item("SREP_CODE") & "" <> TAC.TACMAIN1.SREP_CODE _
                                And Not TAC.TACMAIN1.SREP_CODEs.Contains(rowARTCUST1.Item("SREP_CODE") & "") Then

                                Dim found_store As Boolean = False
                                ASCMAIN1.sql = "Select Distinct SREP_CODE from ARTCUST2 where CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                                For Each rowARTCUST2_SREP As DataRow In ASCDATA1.GetDataTable.Select("")
                                    If rowARTCUST2_SREP.Item("SREP_CODE") & "" <> TAC.TACMAIN1.SREP_CODE _
                                        And Not TAC.TACMAIN1.SREP_CODEs.Contains(rowARTCUST2_SREP.Item("SREP_CODE") & "") Then
                                    Else
                                        found_store = True
                                    End If
                                Next

                                If Not found_store Then
                                    EMsg &= vbCr & "Customer " & Absx1.txtFor("CUST_CODE").Text & " is not connected to Sales Rep code " & TAC.TACMAIN1.SREP_CODE
                                End If
                            End If
                        End If

                        If EMsg = "" Then
                            CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                        End If
                    Else
                        EMsg &= vbCr & "No Record of Customer " & Absx1.txtFor("CUST_CODE").Text
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

            Case "Select"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Refresh"
                Dim savedFilters As New Dictionary(Of String, Infragistics.Win.UltraWinGrid.FilterCondition) 'save filters for each column in dictionary, key column name and value is filter cond

                For Each column As Infragistics.Win.UltraWinGrid.UltraGridColumn In grdSOTORDR0.DisplayLayout.Bands(0).Columns 'loop through each col in grid and see if it has a filter
                    If column.Band.ColumnFilters(column.Key).FilterConditions.Count > 0 Then
                        Dim filter As Infragistics.Win.UltraWinGrid.FilterCondition = column.Band.ColumnFilters(column.Key).FilterConditions(0)
                        savedFilters.Add(column.Key, filter)
                    End If
                Next

                Load_SOTORDR0()
                optGROUP_STATUS_ValueChanged(Nothing, Nothing)

                For Each savedFilter As KeyValuePair(Of String, Infragistics.Win.UltraWinGrid.FilterCondition) In savedFilters 'reapply any filters from before the refresh
                    grdSOTORDR0.DisplayLayout.Bands(0).ColumnFilters(savedFilter.Key).FilterConditions.Add(savedFilter.Value)
                Next



        End Select

    End Sub

    Private Sub Get850Data()

        Dim FOLDER As String = "C:\Users\wjz\Desktop\SLP\850 examples from SPS\"
        Dim FILES As New Dictionary(Of String, String)
        FILES.Add("documentDownloads20231117_1122_728253 (BLM Holiday)", "BLOOMIES")
        FILES.Add("documentDownloads20231117_1124_998486 (Nordstrom Holiday)", "NORDSTROM")
        FILES.Add("documentDownloads20231117_1126_35521 (Ulta Holiday)", "ULTA")
        FILES.Add("documentDownloads20231117_1142_254490 (Nordstrom Core Replen)", "NORDSTROM")
        FILES.Add("po5702844_318360 (BLM Core Replen)", "BLOOMIES")
        FILES.Add("po0100909593_271089 (Ulta Core Replen)", "ULTA")

        Create_TDA(dst.Tables.Add, "SLP_850", "*")

        Dim COLS As New List(Of String)
        For Each DCOL As DataColumn In dst.Tables("SLP_850").Columns
            COLS.Add(DCOL.ColumnName)
        Next

        Dim tblColsMissing As New DataTable
        With tblColsMissing
            .Columns.Add("CUST_CODE")
            .Columns.Add("FILENAME")
            .Columns.Add("COLUMN_NAME")
            .PrimaryKey = New DataColumn() { .Columns("CUST_CODE"), .Columns("FILENAME"), .Columns("COLUMN_NAME")}
        End With

        Try
            Dim IFILE As Integer = 0
            Dim dt As DataTable = Nothing

            For Each FILENAME_root As String In FILES.Keys

                Dim FILENAME As String = FOLDER & FILENAME_root & ".CSV"
                Dim COLSX As New List(Of String)

                IFILE += 1
                Dim CUST_CODE As String = FILES(FILENAME_root)

                Dim iline As Int16 = 0
                Dim inputLine As String = String.Empty



                Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(FILENAME)

                    MyReader.TextFieldType = FileIO.FieldType.Delimited
                    MyReader.SetDelimiters(",")

                    Dim currentRow As String()

                    While Not MyReader.EndOfData
                        currentRow = MyReader.ReadFields()

                        iline += 1

                        If iline = 1 Then
                            If IFILE = 1 Then
                                dt = New DataTable("SLP_850X")
                                For Each fieldName As String In currentRow
                                    fieldName = fieldName.Trim.Replace(" ", "_").Replace("/", "_")
                                    dt.Columns.Add(fieldName)
                                Next
                            Else
                                dt.Rows.Clear()
                            End If
                            Continue While
                        End If

                        dt.Rows.Add(currentRow)

                    End While
                    MyReader.Close()
                    MyReader.Dispose()

                    For Each row As DataRow In dt.Select("")
                        Dim row850 As DataRow = dst.Tables("SLP_850").NewRow
                        row850.Item("CUST_CODE") = CUST_CODE
                        For Each dcol As DataColumn In dt.Columns
                            Dim CN As String = dcol.ColumnName
                            Dim V As String = row.Item(CN) & ""
                            If V <> "" Then
                                Dim CNO As String = Replace(Replace(CN.ToUpper, "#", "_"), "%", "PCT")
                                If CNO = "SIZE" Then CNO = "SIZE_"
                                If Not COLS.Contains(CNO) Then
                                    If Not COLSX.Contains(CN) Then
                                        COLSX.Add(CN)
                                        tblColsMissing.Rows.Add(New String() {CUST_CODE, FILENAME_root, CN})
                                    End If
                                Else
                                    row850.Item(CNO) = V
                                End If

                            End If
                        Next
                        dst.Tables("SLP_850").Rows.Add(row850)
                    Next


                End Using
            Next

        Catch ex As Exception
            EMsg &= vbCr & ex.Message
        End Try


        If tblColsMissing.Rows.Count > 1 Then
            Using F As New ASFMSGBF
                F.Show_grd(tblColsMissing, Me, "Columns Missing")

            End Using
        End If

        Stop

        ASCMAIN1.sql = "Truncate Table SLP_850"
        ASCDATA1.ExecuteSQL()

        Update_Record_TDA("SLP_850")

        Stop

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Select").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Refresh").Settings.Enabled = not_iScreenMode

            .Groups("Find Customer By").Visible = Not ScreenMode
            .Groups("Show Orders").Visible = ScreenMode
            .Groups("Items").Visible = False
            .Groups("12 Month History").Visible = False
            .Groups("Status && Tracking").Visible = ScreenMode

        End With

        grpSOTORDR0.Visible = Not ScreenMode
        tabMain.Visible = ScreenMode

        grdSOTORDR0.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = ScreenMode

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        spl12Months.Visible = False
        chtSATCSLS1_X.Visible = False
        'chkEditInternalNotes.Visible = ScreenMode

        grdSOTORDR0.DisplayLayout.Bands(0).Columns("CUST_NAME").Hidden = ScreenMode

        If ScreenMode Then
            grdSOTORDR0.Parent = splSOTORDR0.Panel1
            chkEditInternalNotes.Parent = grdSOTORDR0

            Setup_tabMain()
            Setup_Summary()
        Else
            Clear_Record()
            grdSOTORDR0.Parent = grpSOTORDR0
            optGROUP_STATUS_ValueChanged(Nothing, Nothing)
        End If
    End Sub

    Sub Clear_Record()

        Absx1.txtFor("CUST_CODE").Text = ""
        CUST_CODE = ""
        chkEditInternalNotes.Visible = False

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTORDR1", "SOTORDRS", "SOTPICK1", "SOTPICK2", "SOTSHIP1", "SOTSHIP2",
             "SOTCORDX", "SOTCORDY", "SOTCORDD", "SOTORDRM", "SOTCART1", "SOTCART2", "SOTCART3", "SOTORDRT", "SOTORDRU"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        optOrders.Value = "OP"

        tabMain.SelectedTab = tabMain.Tabs("Orders")
        tabMonth.Tabs("Details").Visible = False

        grdSOTORDR0.Tag = ""
        Load_SOTORDR0("")
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")
        grdSOTORDR0.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        Load_SOTORDR0("", CUST_CODE)
        Setup_SOTORDR0()

        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control,
         ByVal COLUMN_NAME As String,
         Optional ByRef sql_where As String = "",
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            Case "CUST_CODE"
                'If Trim(ASCMAIN1.USER_CODES) = "FS" Then
                '    If TAC.TACMAIN1.SREP_CODEs.Count > 1 Then
                '        sql_where = " and ARTCUST1.SREP_CODE in ('" & Join(TAC.TACMAIN1.SREP_CODEs.ToArray, "','") & "')"
                '    Else
                '        sql_where = " and ARTCUST1.SREP_CODE = '" & TAC.TACMAIN1.SREP_CODE & "'"
                '    End If
                'End If

                If Trim(ASCMAIN1.USER_CODES) = "FS" Then
                    If TAC.TACMAIN1.SREP_CODEs.Count > 1 Then
                        sql_where = " and (ARTCUST1.SREP_CODE in ('" & Join(TAC.TACMAIN1.SREP_CODEs.ToArray, "','") & "')" _
                            & " or ARTCUST1.CUST_CODE in (Select Distinct CUST_CODE from ARTCUST2 where ARTCUST2.SREP_CODE in ('" & Join(TAC.TACMAIN1.SREP_CODEs.ToArray, "','") & "')))"
                    Else
                        sql_where = " and (ARTCUST1.SREP_CODE = '" & TAC.TACMAIN1.SREP_CODE & "'" _
                            & " or ARTCUST1.CUST_CODE in (Select Distinct CUST_CODE from ARTCUST2 where ARTCUST2.SREP_CODE = '" & TAC.TACMAIN1.SREP_CODE & "'))"
                    End If
                End If
        End Select
    End Sub

    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "Select"

                Dim CUST_CODE As String = Split(key, ":")(0)
                Dim ORDR_GROUP_NO As String = Split(key & ":", ":")(1)
                Absx1.txtFor("CUST_CODE").Text = CUST_CODE
                Click_Command("Select")
                If ORDR_GROUP_NO <> "" Then
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Rows
                        If grow.Cells("ORDR_GROUP_NO").Value = ORDR_GROUP_NO Then
                            grdSOTORDR0.ActiveRow = grow
                            grdSOTORDR0.ActiveRowScrollRegion.FirstRow = grow
                            grow.Selected = True
                        End If

                    Next
                End If
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "ARTCUST1"
            E.COLUMN_NAME = "CUST_CODE"
            E.CODE_VALUE = Absx1.txtFor("CUST_CODE").Text
            E.DESC_VALUE = "Customer"
            E.ATTACHMENT_NOTES = ""
            'If rowSOTORDR1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTORDR0, "SSSBBSSBBBBB", "Show Filter", "Show GroupBox", "Show Pins",
                        "Store Configuration Report", "Customer Order Summary", "Show Original Ship/Cancel", "Show Orders with Changed Ship/Cancel",
                        "Sales Order Entry", "Show Raw EDI", "Create EDI 855", "Edit Internal Notes", "Show Internal Notes Audit Trail")
        Load_Popup_Menu(grdSOTORDR1, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "Sales Order Entry", "Show Raw EDI")
        Load_Popup_Menu(grdSOTORDRS, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Item Status Inquiry")
        Load_Popup_Menu(grdSOTPICK1, "SSSBBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Show Invoice", "Show Pro-Forma Invoice", "email Invoice", "email Pro-Forma Invoice", "EDI Data")
        Load_Popup_Menu(grdSOTSHIP1, "SSSBBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Show Invoice", "Show Pro-Forma Invoice", "email Invoice", "email Pro-Forma Invoice")
        Load_Popup_Menu(grdSOTCORDY, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry")
        Load_Popup_Menu(grdSOTORDRX, "BB", "Sales Order Inquiry", "Show Raw EDI")
        Load_Popup_Menu(grdSOTCART1, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Track Shipment", "View Package POD")
        Load_Popup_Menu(grdSOTORDRT, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry")
        Load_Popup_Menu(grdSOTORDRU, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry")
        Load_Popup_Menu(grdSOTORDCC, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "Export Grids")

    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        If Not GRDs.ContainsKey(Mid(e.SourceControl.Name, 4)) Then Exit Sub

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            '   e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
                Case "grdSOTORDR0"
                    num855Selected = 0

                    tlb_pop = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
                    tlb_btn = DirectCast(tlb_pop.Tools("Store Configuration Report"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = True
                    tlb_btn = DirectCast(tlb_pop.Tools("Customer Order Summary"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ScreenMode

                    tlb_btn = DirectCast(tlb_pop.Tools("Edit Internal Notes"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = True

                    tlb_btn = DirectCast(tlb_pop.Tools("Create EDI 855"), UltraWinToolbars.ButtonTool)
                    If EntryMode <> "V" Then
                        tlb_btn.SharedProps.Visible = False
                    ElseIf optOrders.Value <> "O" Then
                        tlb_btn.SharedProps.Visible = False
                    ElseIf dst.Tables("EDTTRPM1").Select($"CUST_CODE = '{HFs("CUST_CODE")}'").Length = 0 Then
                        tlb_btn.SharedProps.Visible = False
                    ElseIf grdSOTORDR0.Selected.Rows.Count = 0 Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        lstOrderGroupNos.Clear()
                        lstAlreadySent.Clear()
                        num855Selected = grdSOTORDR0.Selected.Rows.Count

                        For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSOTORDR0.Selected.Rows
                            Dim EDI_DOC_SEQ_NO As String = grdRow.Cells("EDI_DOC_SEQ_NO").Value & String.Empty
                            Dim ORDR_SOURCE As String = grdRow.Cells("ORDR_SOURCE").Value & String.Empty
                            Dim ORDR_GROUP_NO As String = grdRow.Cells("ORDR_GROUP_NO").Value & String.Empty
                            Dim CUST_855 As String = grdRow.Cells("CUST_855").Value & String.Empty
                            Dim SENT_855 As String = grdRow.Cells("SENT_855").Value & String.Empty

                            If EDI_DOC_SEQ_NO.Length > 0 AndAlso ORDR_SOURCE = "E" AndAlso CUST_855 = "1" Then
                                If SENT_855.Length = 0 AndAlso Not lstOrderGroupNos.Contains(ORDR_GROUP_NO) Then
                                    lstOrderGroupNos.Add(ORDR_GROUP_NO)
                                Else
                                    lstAlreadySent.Add(ORDR_GROUP_NO)
                                End If
                            End If
                        Next

                        If lstOrderGroupNos.Count > 0 OrElse lstAlreadySent.Count > 0 Then
                            tlb_btn.SharedProps.Visible = True
                        Else
                            tlb_btn.SharedProps.Visible = False
                        End If
                    End If

                Case "grdSOTPICK1"
                    Dim PICK_STATUS As String = ""
                    If grd.ActiveRow.Band.Key = "SOTPICK1" Then
                        PICK_STATUS = grd.ActiveRow.Cells("PICK_STATUS").Value & ""
                    Else
                        PICK_STATUS = grd.ActiveRow.ParentRow.Cells("PICK_STATUS").Value & ""
                    End If

                    tlb_btn = DirectCast(tlb_pop.Tools("Show Invoice"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (PICK_STATUS = "F")
                    tlb_btn = DirectCast(tlb_pop.Tools("Show Pro-Forma Invoice"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (PICK_STATUS = "P")
                    tlb_btn = DirectCast(tlb_pop.Tools("email Invoice"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (PICK_STATUS = "F")
                    tlb_btn = DirectCast(tlb_pop.Tools("email Pro-Forma Invoice"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (PICK_STATUS = "P")

                Case "grdSOTSHIP1"

                    Dim SHIP_STATUS As String = ""
                    If grd.ActiveRow.Band.Index = 0 Then
                        SHIP_STATUS = grd.ActiveRow.Cells("SHIP_STATUS").Value & ""
                    Else
                        SHIP_STATUS = grd.ActiveRow.ParentRow.Cells("SHIP_STATUS").Value & ""
                    End If
                    tlb_btn = DirectCast(tlb_pop.Tools("Show Invoice"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (SHIP_STATUS = "F")
                    tlb_btn = DirectCast(tlb_pop.Tools("Show Pro-Forma Invoice"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (SHIP_STATUS = "P")
                    tlb_btn = DirectCast(tlb_pop.Tools("email Invoice"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (SHIP_STATUS = "F")
                    tlb_btn = DirectCast(tlb_pop.Tools("email Pro-Forma Invoice"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (SHIP_STATUS = "P")

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Show Filter", "Show GroupBox"
                Dim tlb_sbt_Filter As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
                Dim tlb_sbt_GroupBox As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
                If grd.Name = "grdSOTCART1" Then
                    If tlb_sbt_Filter.Checked Or tlb_sbt_GroupBox.Checked Then
                        grdSOTCART1.DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.OutlookGroupBy
                    Else
                        grdSOTCART1.DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.Horizontal
                    End If
                End If

            Case "Show Orders with Changed Ship/Cancel"
                Toggle_ChgShipCancel()
            Case "Show Original Ship/Cancel"
                Toggle_ShowShipCancel()
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Show Internal Notes Audit Trail"
                If grdSOTORDR0.Selected.Rows.Count = 0 Then
                    If grdSOTORDR0.ActiveRow IsNot Nothing Then
                        grdSOTORDR0.ActiveRow.Selected = True
                    End If
                End If
                If grdSOTORDR0.Selected.Rows.Count = 1 Then
                    Dim ORDR_GROUP_NO As String = grdSOTORDR0.Selected.Rows(0).Cells("ORDR_GROUP_NO").Value
                    Using F As New ASFMSGBF
                        F.grdGroupBy = True
                        F.grdFilter = True
                        ASCMAIN1.sql = $"Select COLUMN_NAME, INIT_DATE, USER_ID, OLD_VALUE, NEW_VALUE, XNO, FM_MODE from ASTAUDT1 where TABLE_NAME = 'SOTORDR0' and (KEY_VALUE = 'O:{ORDR_GROUP_NO}' or KEY_VALUE like '{ORDR_GROUP_NO}%') and COLUMN_NAME = 'ORDR_INTERNAL_NOTES' ORDER BY INIT_DATE DESC"
                        F.Show_grd(ASCDATA1.GetDataTable, Me, $"Audit Trail for Table: SOTORDR0; Record: {ORDR_GROUP_NO}")
                    End Using
                End If

            Case "Edit Internal Notes"

                If grdSOTORDR0.Selected.Rows.Count = 0 Then
                    If grdSOTORDR0.ActiveRow IsNot Nothing Then
                        grdSOTORDR0.ActiveRow.Selected = True
                    End If
                End If


                If grdSOTORDR0.Selected.Rows.Count = 0 Then
                    MsgBox("You must select at least 1 order group to Edit the Internal Notes", MsgBoxStyle.OkOnly, "Cannot Edit Internal Notes")
                Else
                    Dim CUST_CODE_selected As String = ""
                    Dim ORDR_INTERNAL_NOTES_selected As String = ""
                    Dim ORDR_GROUP_NOs As New List(Of String)
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Selected.Rows
                        Dim ORDR_GROUP_NO As String = grow.Cells("ORDR_GROUP_NO").Value
                        Dim CUST_CODE As String = grow.Cells("CUST_CODE").Value
                        If CUST_CODE_selected = "" Then
                            CUST_CODE_selected = CUST_CODE
                            ORDR_INTERNAL_NOTES_selected = grow.Cells("ORDR_INTERNAL_NOTES").Value & ""
                        Else
                            If CUST_CODE <> CUST_CODE_selected Then
                                MsgBox("You cannot edit multiple Order Groups across more than 1 customer to Edit the Internal Notes", MsgBoxStyle.OkOnly, "Cannot Edit Internal Notes")
                                ASCMAIN1.MultiTask_Release()
                                ORDR_GROUP_NOs.Clear()
                                Exit For
                            Else
                                If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO) Then
                                    ORDR_GROUP_NOs.Clear()
                                    Exit For
                                End If
                            End If
                        End If
                        ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)
                    Next

                    If ORDR_GROUP_NOs.Count > 0 Then
                        Using frm As New ASFMSGBF

                            Dim ORDR_INTERNAL_NOTES As String = frm.Get_txtblock_from_User("Internal Notes", $"Enter Internal Notes for {ORDR_GROUP_NOs.Count} order(s) for {CUST_CODE_selected}", ORDR_INTERNAL_NOTES_selected,, 1000)

                            If frm.user_option = -1 Then
                                ' user clicked cancel
                            Else

                                DATETIME_STAMP = Now + ASCMAIN1.NowTSD

                                For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs
                                    Dim rowSOTORDR0 As DataRow = LookUp("SOTORDR0", ORDR_GROUP_NO)
                                    Dim ORDR_INTERNAL_NOTES_orig As String = rowSOTORDR0.Item("ORDR_INTERNAL_NOTES") & ""
                                    rowSOTORDR0 = dst.Tables("SOTORDR0").Rows.Find(New String() {ORDR_GROUP_NO})
                                    rowSOTORDR0.Item("ORDR_INTERNAL_NOTES") = ORDR_INTERNAL_NOTES_orig
                                    rowSOTORDR0.AcceptChanges()
                                    rowSOTORDR0.Item("ORDR_INTERNAL_NOTES") = ORDR_INTERNAL_NOTES
                                    Write_Audit_Trail(rowSOTORDR0, "E")

                                    ASCMAIN1.sql = "Update SOTORDR0 Set ORDR_INTERNAL_NOTES = :PARM1 where ORDR_GROUP_NO = :PARM2"
                                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {ORDR_INTERNAL_NOTES, ORDR_GROUP_NO})
                                    ASCMAIN1.sql = "Update SOTORDR1 Set ORDR_INTERNAL_NOTES = :PARM1 where ORDR_GROUP_NO = :PARM2"
                                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {ORDR_INTERNAL_NOTES, ORDR_GROUP_NO})
                                Next

                                'For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Selected.Rows
                                '    Dim ORDR_GROUP_NO As String = grow.Cells("ORDR_GROUP_NO").Value
                                '    Dim ORDR_TYPE As String = grow.Cells("ORDR_TYPE").Value
                                '    Dim rowSOTORDR0 As DataRow = LookUp("SOTORDR0", ORDR_GROUP_NO)
                                '    Dim ORDR_INTERNAL_NOTES_orig As String = rowSOTORDR0.Item("ORDR_INTERNAL_NOTES") & ""
                                '    rowSOTORDR0 = dst.Tables("SOTORDR0").Rows.Find(New String() {ORDR_TYPE, ORDR_GROUP_NO})
                                '    rowSOTORDR0.Item("ORDR_INTERNAL_NOTES") = ORDR_INTERNAL_NOTES_orig
                                '    rowSOTORDR0.AcceptChanges()
                                '    rowSOTORDR0.Item("ORDR_INTERNAL_NOTES") = ORDR_INTERNAL_NOTES
                                '    Write_Audit_Trail(rowSOTORDR0, "E")
                                '    'grow.Cells("ORDR_INTERNAL_NOTES").Value = ORDR_INTERNAL_NOTES
                                '    'grow.Update()
                                'Next
                            End If


                            grdSOTORDR0.Selected.Rows.Clear()
                        End Using


                        ASCMAIN1.MultiTask_Release()

                    End If
                End If


            Case "Create EDI 855"
                Try
                    If num855Selected = 0 Then
                        MessageBox.Show("There are no 855s to create.", "Create EDI 855", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Exit Sub
                    End If

                    Dim zmsg As String = String.Empty
                    zmsg = $"There are {num855Selected} Order Groups selected."

                    If lstAlreadySent.Count > 0 Then
                        zmsg &= Environment.NewLine & $"WARNING: {lstAlreadySent.Count} of the selected Order Groups were previoulsy sent and new 855s will be created."
                    End If

                    zmsg &= Environment.NewLine & Environment.NewLine & $"Do you want to generate 855s for the {num855Selected} selected Order Groups."

                    num855Selected = 0

                    If MessageBox.Show(zmsg, "Create EDI 855", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        Exit Sub
                    End If

                    For Each ORDR_GROUP_NO As String In lstAlreadySent
                        If Not lstOrderGroupNos.Contains(ORDR_GROUP_NO) Then
                            lstOrderGroupNos.Add(ORDR_GROUP_NO)
                        End If
                    Next

                    Generate855s(lstOrderGroupNos)

                Catch ex As Exception
                    MessageBox.Show($"Error: {ex.Message}", "Create EDI 855", MessageBoxButtons.OK, MessageBoxIcon.Error)

                Finally
                    lstAlreadySent.Clear()
                    lstOrderGroupNos.Clear()
                    num855Selected = 0
                End Try

            Case "Export Grids"
                Export_Status_And_Tracking()

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Sales Order Entry"
                Dim ORDR_NO As String = ""
                If grd.Name = "grdSOTORDR0" Then
                    Dim ORDR_GROUP_NO = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value
                    ORDR_NO = ASCDATA1.GetDataValue("Select Min (ORDR_NO) from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'")
                Else
                    ORDR_NO = grd.ActiveRow.Cells("ORDR_NO").Value
                End If

                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDR1")
                End If

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Value
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Show Raw EDI"

                If grdSOTORDR1.ActiveRow IsNot Nothing Then
                    Dim EDI_DOC_SEQ_NO As String = grdSOTORDR1.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & ""
                    Dim RAW_EDI As String = TAC.SOCMAIN1.Get_Raw_EDI(EDI_DOC_SEQ_NO)
                    Using frm As New ASFTEXT1
                        frm.t = RAW_EDI
                        frm.Text = "Raw EDI for " & CUST_CODE & " PO No " & grdSOTORDR1.ActiveRow.Cells("ORDR_CUST_PO").Value
                        frm.ShowDialog()
                    End Using
                End If

            Case "Store Configuration Report"
                Dim ORDR_GROUP_NO As String = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value
                Store_Configuration_Report(ORDR_GROUP_NO)

            Case "Customer Order Summary"

                dst.Tables("SOTCORDR").Rows.Clear()
                dst.Tables("SOTCORDG").Rows.Clear()

                Dim TBL As DataTable = dst.Tables("SOTORDRS").Copy

                'Dim TBL As DataTable = Nothing
                'If grd.Selected.Rows.Count = 1 Or chkShowSelectedOrder.Checked Then
                '    TBL = dst.Tables("SOTORDRS").Copy
                'End If

                If grd.Selected.Rows.Count = 0 Then grd.ActiveRow.Selected = True
                If chkShowSelectedOrder.Checked Then
                    grd.Selected.Rows.Clear()
                    grd.ActiveRow.Selected = True
                End If

                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    Dim ORDR_GROUP_NO As String = grow.Cells("ORDR_GROUP_NO").Value
                    Dim ORDR_CUST_PO As String = grow.Cells("ORDR_CUST_PO").Value & ""

                    If grd.Selected.Rows.Count = 1 Or chkShowSelectedOrder.Checked Then
                    Else
                        Dim sql As String = Replace(sqlSOTORDRS, " group by ", " and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' group by ")
                        Fill_Records("SOTORDRS", "", True, sql)
                    End If
                    Summarize_Group(ORDR_GROUP_NO)
                Next

                If grd.Selected.Rows.Count = 1 Or chkShowSelectedOrder.Checked Then
                    dst.Tables("SOTORDRS").Rows.Clear()
                    dst.Tables("SOTORDRS").Merge(TBL)
                End If

                Print_Report_Begin()
                Dim SUBT As String
                If grd.Selected.Rows.Count = 1 Or chkShowSelectedOrder.Checked Then
                    With grd.Selected.Rows(0)
                        Dim ORDR_GROUP_NO As String = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value
                        Dim ORDR_CUST_PO As String = grd.ActiveRow.Cells("ORDR_CUST_PO").Value
                        Dim CUST_DC_NO As String = grd.ActiveRow.Cells("CUST_DC_NO").Value & ""

                        Dim ORDR_DATE As Date = grd.ActiveRow.Cells("ORDR_DATE").Value
                        Dim ORDR_SHIP_DATE As Date = grd.ActiveRow.Cells("ORDR_SHIP_DATE").Value
                        Dim ORDR_CANCEL_DATE As Date = grd.ActiveRow.Cells("ORDR_CANCEL_DATE").Value

                        Dim CUST_STORE_NO As String = ""
                        If chkShowSelectedOrder.Checked Then
                            If grdSOTORDRX.ActiveRow IsNot Nothing Then
                                CUST_STORE_NO = grdSOTORDRX.ActiveRow.Cells("CUST_STORE_NO").Value & ""
                            End If
                        End If

                        SUBT = CUST_CODE _
                            & ", PO " & ORDR_CUST_PO & " " & Format(ORDR_DATE, "MM/dd/yy") _
                            & ", Ship " & Format(ORDR_SHIP_DATE, "MM/dd/yy") _
                            & ", Cancel " & Format(ORDR_CANCEL_DATE, "MM/dd/yy") _
                            & IIf(chkShowSelectedOrder.Checked, ", Store No " & CUST_STORE_NO, "") _
                            & ", Order Group No " & ORDR_GROUP_NO
                    End With
                Else
                    SUBT = CUST_CODE & ", " & grd.Selected.Rows.Count & " Selected POs"
                End If

                Generate_Report("SORCORDR", "Customer Order Summary", SUBT, , , , False)
                Print_Report_End()

            Case "Show Invoice", "Show Pro-Forma Invoice", "email Invoice", "email Pro-Forma Invoice"
                Dim SHIP_BOL_NO As String = ""
                Dim PICK_NO As String = ""
                Dim FILENAME As String = ""
                If grd.Name = "grdSOTSHIP1" Then
                    SHIP_BOL_NO = grd.ActiveRow.Cells("SHIP_BOL_NO").Value
                Else
                    PICK_NO = grd.ActiveRow.Cells("PICK_NO").Value
                End If

                Dim make_pdf As Boolean = (e.Tool.Key Like "*email*")
                If e.Tool.Key Like "*Pro-Forma*" Then
                    FILENAME = Create_Invoice(SHIP_BOL_NO, PICK_NO, make_pdf, True)
                Else
                    FILENAME = Create_Invoice(SHIP_BOL_NO, PICK_NO, make_pdf)
                End If

                If e.Tool.Key Like "*email*" Then
                    Dim ATTACHMENT As String = FILENAME & ".PDF"
                    FILENAME = ASCMAIN1.Folders("Temp") & ASCMAIN1.DBS_COMPANY & "_" & FILENAME & ".PDF"
                    Dim SUBJECT As String = ""
                    Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    Dim emailAddress As String = rowARTCUST1.Item("CUST_EMAIL") & String.Empty
                    emailAddress = emailAddress.Trim
                    If rowARTCUST1.Item("CUST_INV_EMAIL") & "" <> "" Then
                        If emailAddress.Length > 0 Then
                            emailAddress &= ";"
                        End If
                        emailAddress &= rowARTCUST1.Item("CUST_INV_EMAIL") & ""
                    End If

                    TAC.TACMAIN1.Send_email_with_Attachment(Me,
                         FILENAME, IIf(ATTACHMENT = "", FILENAME, ATTACHMENT), SUBJECT,
                         emailAddress, rowARTCUST1.Item("CUST_CONTACT") & "",
                         "INV", Absx1.txtFor("CUST_CODE").Text, Absx1.txtFor("CUST_NAME").Text, "Customer")
                End If

            Case "Track Shipment"
                If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.Band.Key = "SOTCART1" Then
                    Dim CART_TRACKING_NO As String = grd.ActiveRow.Cells("CART_TRACKING_NO").Value & ""
                    Dim PICK_NO As String = grd.ActiveRow.Cells("PICK_NO").Value & ""
                    If CART_TRACKING_NO <> "" Then
                        ASCMAIN1.sql = "Select SOTSHIP1.SHIP_VIA_CODE from SOTSHIP1,SOTPICK1" _
                            & " where SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO and SOTPICK1.PICK_NO = '" & PICK_NO & "'"
                        Dim SHIP_VIA_CODE As String = ASCDATA1.GetDataValue
                        TAC.SOCMAIN1.Track_Shipment(SHIP_VIA_CODE, CART_TRACKING_NO)
                    End If
                End If

            Case "View Package POD"
                Try
                    If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.Band.Key = "SOTCART1" Then
                        Dim CART_TRACKING_NO As String = grd.ActiveRow.Cells("CART_TRACKING_NO").Value & ""
                        Dim PICK_NO As String = grd.ActiveRow.Cells("PICK_NO").Value & ""
                        If CART_TRACKING_NO <> "" Then
                            ASCMAIN1.sql = "Select * from SOTSHIPT where PICK_NO = :PARM1 AND TRACKING_NO = :PARM2"
                            Dim rowSOTSHIPT As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", {PICK_NO, CART_TRACKING_NO})
                            If rowSOTSHIPT Is Nothing Then
                                MessageBox.Show("No POD record for this tracking no.", "View Package POD", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ElseIf rowSOTSHIPT.Item("POD") & String.Empty = String.Empty Then
                                MessageBox.Show("POD has net been received from shipper.", "View Package POD", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Else
                                Dim frm As New ASFMSGBF
                                frm.Show_Xml_File("View Package POD", rowSOTSHIPT.Item("POD") & String.Empty, Me)
                            End If
                        End If
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "View Package POD", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
        End Select
    End Sub

#End Region

    Sub Summarize_Group(ORDR_GROUP_NO As String)

        Dim rowSOTORDR0 As DataRow = dst.Tables("SOTORDR0").Rows.Find(New Object() {ORDR_GROUP_NO})
        dst.Tables("SOTCORDG").Rows.Add(rowSOTORDR0.ItemArray)

        For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select("", "ITEM_CODE")
            Dim ITEM_CODE As String = rowSOTORDRS.Item("ITEM_CODE")
            Dim rowSOTCORDR As DataRow
            rowSOTCORDR = dst.Tables("SOTCORDR").Rows.Find(ITEM_CODE)
            If rowSOTCORDR Is Nothing Then
                rowSOTCORDR = dst.Tables("SOTCORDR").NewRow
                rowSOTCORDR.Item("ITEM_CODE") = rowSOTORDRS.Item("ITEM_CODE")
                rowSOTCORDR.Item("ITEM_DESC") = rowSOTORDRS.Item("ITEM_DESC")
                dst.Tables("SOTCORDR").Rows.Add(rowSOTCORDR)
            End If

            rowSOTCORDR.Item("ORDR") += Val(rowSOTORDRS.Item("ORDR_QTY") & "")
            rowSOTCORDR.Item("OPEN") += Val(rowSOTORDRS.Item("ORDR_QTY_OPEN") & "")
            rowSOTCORDR.Item("PICK") += Val(rowSOTORDRS.Item("ORDR_QTY_PICK") & "")
            rowSOTCORDR.Item("SHIP") += Val(rowSOTORDRS.Item("ORDR_QTY_SHIP") & "")
            rowSOTCORDR.Item("CANC") += Val(rowSOTORDRS.Item("ORDR_QTY_CANC") & "")
            rowSOTCORDR.Item("ORDR_AMT") += Val(rowSOTORDRS.Item("ORDR_AMT") & "")
            'dst.Tables("SOTCORDR").Rows.Add(rowSOTCORDR)
        Next

    End Sub
#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"

                If ASCMAIN1.USER_ID = "wjz" And Absx1.txtFor("CUST_CODE").Text = "X" Then
                    Dim row As DataRow = Nothing
                    If row.Item("Y") = "Z" Then
                        MsgBox("MAGIC")
                    End If
                End If

                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Select")
                End If

            Case "SREP_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    grdSOTORDR0.Tag = "SREP_CODE"
                    Load_SOTORDR0(Absx1.txtFor("SREP_CODE").Text)
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "CUST_CODE"
            '    If Not ScreenMode Then
            '        Load_SOTORDRX()
            '    End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Click_Command("Select")
            Case "SREP_CODE"
                grdSOTORDR0.Tag = "SREP_CODE"
                Load_SOTORDR0(Absx1.txtFor("SREP_CODE").Text)
        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select
    End Sub
#End Region

    Sub Load_SOTORDR0(Optional PARM1 As String = "", Optional CUST_CODE As String = "")

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Order Summary", "")

        Fill_Records("SOTSHIPX", String.Empty, True, sqlSOTSHIPX)

        If CUST_CODE <> "" Then ' ScreenMode Then
            ASCMAIN1.sql = sqlSOTORDR0
            Dim sqlw As String = " and SOTORDR0.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf
            If cmbSALES_DIVISION_CODE.Value & "" <> "" AndAlso cmbSALES_DIVISION_CODE.SelectedRow IsNot Nothing Then
                sqlw &= " and SOTORDR0.SALES_DIVISION_CODE = '" & cmbSALES_DIVISION_CODE.SelectedRow.Cells("SALES_DIVISION_CODE").Value & "'" & vbCrLf
            End If

            chkEditInternalNotes.Visible = False
            If optOrders.Value = "O" Then
                sqlw &= " AND SOTORDR0.ORDR_QTY_OPEN <> 0" & vbCrLf
                chkEditInternalNotes.Visible = True
            ElseIf optOrders.Value = "P" Then
                sqlw &= " AND SOTORDR0.ORDR_QTY_PICK <> 0" & vbCrLf
            ElseIf optOrders.Value = "S" Then
                sqlw &= " AND SOTORDR0.ORDR_QTY_SHIP <> 0" & vbCrLf
            ElseIf optOrders.Value = "C" Then
                sqlw &= " AND SOTORDR0.ORDR_QTY_CANC <> 0" & vbCrLf
            ElseIf optOrders.Value = "OP" Then
                sqlw &= " AND (NVL(SOTORDR0.ORDR_QTY_OPEN,0) <> 0 OR NVL(SOTORDR0.ORDR_QTY_PICK,0) <> 0)" & vbCrLf
            End If

            grdSOTORDR0.Text = "Orders for " & CUST_CODE & "; Status: " & optOrders.Text

            ASCMAIN1.sql &= sqlw
        Else
            ASCMAIN1.sql = sqlSOTORDR0
            PARM1 = Replace(Replace(PARM1, ";", ""), "'", "")

            Dim sqlORDR_STATUS As String = ""

            Select Case grdSOTORDR0.Tag & ""
                Case ""
                    grdSOTORDR0.Text = "Orders which are either Open or In Pick"
                    ASCMAIN1.sql &= " and (ORDR_CNT_OPEN <> 0 or ORDR_CNT_PICK <> 0)"

                Case "SREP_CODE"
                    grdSOTORDR0.Text = "Open Orders for Sales Rep " & PARM1
                    ASCMAIN1.sql &= " and SOTORDR0.ORDR_GROUP_NO in " _
                        & " (Select Distinct ORDR_GROUP_NO from SOTORDR1 " _
                        & " where ORDR_STATUS >= 'O' and ORDR_STATUS <= 'P'" _
                        & "   and SREP_CODE = '" & PARM1 & "')"

                Case "ORDR_CUST_PO"
                    grdSOTORDR0.Text = "All Customer Orders using Customer PO " & PARM1
                    ASCMAIN1.sql &= " and SOTORDR1.ORDR_GROUP_NO in " _
                        & " (Select Distinct ORDR_GROUP_NO from SOTORDR1 where ORDR_CUST_PO = '" & PARM1 & "' and ORDR_STATUS <> 'D')"
            End Select
        End If

        ASCDATA1.ExecuteSQL("Delete from " & SOTORDR0)
        ASCDATA1.ExecuteSQL("Insert into " & SOTORDR0 & " " & ASCMAIN1.sql)

        ASCMAIN1.sql = "Select SOTORDR1.ORDR_GROUP_NO, Min(SOTPICK1.ERROR_REASON) ERROR" & vbCrLf _
            & " from " & SOTORDR0 & " SOTORDR0, SOTORDR1, SOTPICK1" & vbCrLf _
            & " where SOTORDR1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO" & vbCrLf _
            & "   and SOTORDR1.ORDR_STATUS = 'P'" & vbCrLf _
            & "   and SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            & "   and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
            & "   and SOTPICK1.ERROR_REASON is Not Null" & vbCrLf _
            & " group by SOTORDR1.ORDR_GROUP_NO"

        ASCDATA1.ExecuteSQL("Delete from " & SOTORDR0_ERRORS)
        ASCDATA1.ExecuteSQL("Insert into " & SOTORDR0_ERRORS & " " & ASCMAIN1.sql)

        ASCMAIN1.sql = "Select X.*, ARTCUST1.CUST_NAME, SOTORDR1.CUST_STORE_NO, NVL(SOTORDR1.CUST_STORE_LOCATION,SOTORDR5.CUST_NAME) CUST_STORE_LOCATION, SOTORDR0_ERRORS.ERRORS, SOTORDR1.ORDR_HIGH_PRIORITY, SOTORDR1.ORDR_HIGH_PRIORITY_NOTE" & vbCrLf _
            & " from " & SOTORDR0 & " X, SOTORDR1, ARTCUST1, SOTORDR5, " & SOTORDR0_ERRORS & " SOTORDR0_ERRORS" & vbCrLf _
            & " where SOTORDR1.ORDR_NO = X.ORDR_NO_MIN and ARTCUST1.CUST_CODE = X.CUST_CODE" & vbCrLf _
            & "   and SOTORDR5.ORDR_NO (+) = X.ORDR_NO_MIN" & vbCrLf _
            & "   and SOTORDR5.CUST_ADDR_TYPE (+) = 'ST'" & vbCrLf _
            & "   and SOTORDR0_ERRORS.ORDR_GROUP_NO (+) = X.ORDR_GROUP_NO"

        If Trim(ASCMAIN1.USER_CODES) = "FS" Then
            If TAC.TACMAIN1.SREP_CODEs.Count > 1 Then
                ASCMAIN1.sql &= " and (ARTCUST1.SREP_CODE in ('" & Join(TAC.TACMAIN1.SREP_CODEs.ToArray, "','") & "')" _
                    & " or ARTCUST1.CUST_CODE in (Select Distinct CUST_CODE from ARTCUST2 where ARTCUST2.SREP_CODE in ('" & Join(TAC.TACMAIN1.SREP_CODEs.ToArray, "','") & "')))"
            Else
                ASCMAIN1.sql &= " and (ARTCUST1.SREP_CODE = '" & TAC.TACMAIN1.SREP_CODE & "'" _
                    & " or ARTCUST1.CUST_CODE in (Select Distinct CUST_CODE from ARTCUST2 where ARTCUST2.SREP_CODE = '" & TAC.TACMAIN1.SREP_CODE & "'))"
            End If
        End If

        If ASCMAIN1.CLIENT = "AHA" And CUST_CODE = "CONSUMER" And (optOrders.Value = "A" Or optOrders.Value = "S" Or optOrders.Value = "C") Then
            ASCMAIN1.sql &= vbCrLf & " and X.ORDR_DATE > SYSDATE - 365"
            grdSOTORDR0.Text &= " - Ordered in last 365 days"
        End If

        Fill_Records("SOTORDR0", "", , ASCMAIN1.sql)
        Setup_SOTORDR0()

        Dim sqlFilter As String = String.Empty
        For iIndex As Int16 = 0 To optGROUP_STATUS.Items.Count - 1
            Select Case optGROUP_STATUS.Items(iIndex).DataValue
                Case "A"
                    optGROUP_STATUS.Items(iIndex).DisplayText = "All"
                    sqlFilter = String.Empty
                Case "O"
                    optGROUP_STATUS.Items(iIndex).DisplayText = "Open"
                    sqlFilter = "ORDR_CNT_OPEN <> 0"
                Case "P"
                    optGROUP_STATUS.Items(iIndex).DisplayText = "In Pick"
                    sqlFilter = "ORDR_CNT_PICK <> 0"
                Case "OP"
                    optGROUP_STATUS.Items(iIndex).DisplayText = "Open or Pick"
                    sqlFilter = "ORDR_CNT_OPEN <> 0 OR ORDR_CNT_PICK <> 0"
                Case "OAP"
                    optGROUP_STATUS.Items(iIndex).DisplayText = "Open and Pick"
                    sqlFilter = "ORDR_CNT_OPEN <> 0 AND ORDR_CNT_PICK <> 0"
                Case "C"
                    optGROUP_STATUS.Items(iIndex).DisplayText = "Cancelled"
                    sqlFilter = "ORDR_QTY_SHIP = 0"
                Case "E"
                    optGROUP_STATUS.Items(iIndex).DisplayText = "Error"
                    sqlFilter = "ISNULL(ERRORS, '') <> ''"
                Case "H"
                    optGROUP_STATUS.Items(iIndex).DisplayText = "Hold"
                    sqlFilter = "ORDR_HOLD = '1'"
            End Select

            optGROUP_STATUS.Items(iIndex).DisplayText &= $" ({dst.Tables("SOTORDR0").Select(sqlFilter).Length})"
        Next

        optGROUP_STATUS_ValueChanged(Nothing, Nothing)

        Toggle_ChgShipCancel()

        If optOrders.Value = "S" Or optOrders.Value = "C" Or optOrders.Value = "A" Then
            Sort_grdColumns(grdSOTORDR0, "ORDR_GROUP_NO".ToLower)
        Else
            Sort_grdColumns(grdSOTORDR0, "ORDR_CANCEL_DATE")
        End If

        grdSOTORDR0.Visible = True
        grdSOTORDR0.Rows.ColumnFilters.ClearAllFilters()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")

    End Sub

    Sub Toggle_ChgShipCancel()
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Orders with Changed Ship/Cancel"), UltraWinToolbars.StateButtonTool)

        Dim dvw As DataView = dst.Tables("SOTORDR0").DefaultView

        If tlb_sbt.Checked Then
            dvw.RowFilter = "ORDR_ORIG_SHIP_DATE <> ORDR_SHIP_DATE or ORDR_ORIG_CANCEL_DATE <> ORDR_CANCEL_DATE"
            grdSOTORDR0.Text &= " (Orders with Changes to Ship or Cancel Dates)"
        Else
            dvw.RowFilter = ""
            grdSOTORDR0.Text = Replace(grdSOTORDR0.Text, " (Orders with Changes to Ship or Cancel Dates)", "")
        End If

        Toggle_ShowShipCancel()
    End Sub

    Sub Toggle_ShowShipCancel()
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Original Ship/Cancel"), UltraWinToolbars.StateButtonTool)

        With grdSOTORDR0.DisplayLayout.Bands(0)
            .Columns("ORDR_ORIG_SHIP_DATE").Hidden = Not tlb_sbt.Checked
            .Columns("ORDR_ORIG_CANCEL_DATE").Hidden = Not tlb_sbt.Checked
        End With
    End Sub

    Sub Setup_SOTORDR0()

        If grdSOTORDR0.ActiveRow Is Nothing OrElse Not grdSOTORDR0.ActiveRow.IsDataRow Then
            tabDetails.Visible = False
        Else
            tabDetails.Visible = True
            ORDR_GROUP_NO = grdSOTORDR0.ActiveRow.Cells("ORDR_GROUP_NO").Value
            Dim ORDR_CUST_PO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_CUST_PO").Value & ""
            chkShowSelectedOrder.Enabled = True
            EnforceConstraints(False)
            dst.Tables("SOTPICK2").Rows.Clear()
            dst.Tables("SOTSHIP2").Rows.Clear()
            Fill_Records("SOTORDR1", ORDR_GROUP_NO)
            Sort_grdColumns(grdSOTORDR1, "ORDR_NO")
            grdSOTORDR1.Text = "Sales Orders for Order Group " & ORDR_GROUP_NO

            Fill_Records("SOTPICK1", ORDR_GROUP_NO)
            Sort_grdColumns(grdSOTPICK1, "PICK_NO")
            grdSOTPICK1.Text = "Pick Tickets for Order Group " & ORDR_GROUP_NO

            Fill_Records("SOTSHIP1", ORDR_GROUP_NO)
            Sort_grdColumns(grdSOTSHIP1, "SHIP_BOL_NO")
            grdSOTSHIP1.Text = "Shipments for Order Group " & ORDR_GROUP_NO

            EnforceConstraints(True)
            Load_SOTORDRS()
            ' ASCMAIN1.Progress("")
        End If

    End Sub

    Private Sub grdSOTORDR0_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTORDR0.AfterRowActivate
        If ScreenMode Then
            Setup_SOTORDR0()
        End If
    End Sub

    Private Sub grdSOTORDR0_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTORDR0.DoubleClickRow
        If e.Row.IsDataRow Then
            If Not ScreenMode Then
                Absx1.txtFor("CUST_CODE").Text = e.Row.Cells("CUST_CODE").Value
                Click_Command("Select")
            End If
        End If
    End Sub

    Private Sub txtFindBy_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtFindBy.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then

            Dim FIND_BY As String = txtFindBy.Text.ToUpper()
            FIND_BY = Replace(Replace(FIND_BY, ";", ""), "'", "")

            If optFindBy.Value = "T" OrElse optFindBy.Value = "B" Then
                ' Do not alter the Carton code
            ElseIf optFindBy.Value <> "C" Then
                If Len(FIND_BY) > 10 Then
                    FIND_BY = Mid(FIND_BY, 1, 10)
                Else
                    FIND_BY = FIND_BY.PadLeft(10, "0")
                End If
            End If

            Select Case optFindBy.Value
                Case "C"
                    ASCMAIN1.sql = "Select Distinct CUST_CODE from SOTORDR1 where UPPER(ORDR_CUST_PO) = :PARM1 and ORDR_STATUS <> 'D'"
                Case "O"
                    ASCMAIN1.sql = "Select CUST_CODE from SOTORDR1 where UPPER(ORDR_NO) = :PARM1 and ORDR_STATUS <> 'D'"
                Case "I"
                    ASCMAIN1.sql = "Select CUST_CODE from SOTINVH1 where UPPER(INV_NO) = :PARM1"
                Case "P"
                    ASCMAIN1.sql = "Select SOTORDR1.CUST_CODE from SOTPICK1,SOTORDR1 where UPPER(SOTPICK1.PICK_NO) = :PARM1 and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO and SOTORDR1.ORDR_STATUS <> 'D'"
                Case "T"
                    ASCMAIN1.sql = "Select Distinct SOTORDR1.CUST_CODE, SOTORDR1.ORDR_CUST_PO
                        from SOTORDR1, SOTPICK1, SOTCART1
                        WHERE SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO
                        AND SOTCART1.PICK_NO = SOTPICK1.PICK_NO
                        AND UPPER(SOTCART1.CART_NO) = :PARM1"
                Case "B"
                    ASCMAIN1.sql = "Select Distinct SOTORDR1.CUST_CODE, SOTORDR1.ORDR_CUST_PO, SOTSHIP1.SHIPPED_ACTUAL SHIP_DATE
                        from SOTORDR1, SOTPICK1, SOTSHIP1
                        WHERE SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO
                        AND SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO
                        AND UPPER(SOTSHIP1.BILL_OF_LADING_NO) = :PARM1"
            End Select

            Dim rows() As DataRow = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {FIND_BY}).Select()

            ' A Clarins BOL may apoear more than once
            If rows.Count > 1 AndAlso optFindBy.Value = "B" Then

                Dim CUST_CODE As String = String.Empty
                Dim ORDR_CUST_PO As String = String.Empty

                If ASCDATA1.SelectDistinct(rows, New String() {"CUST_CODE"}).Rows.Count = 1 Then
                    CUST_CODE = rows(0).Item("CUST_CODE") & String.Empty
                    ORDR_CUST_PO = rows(0).Item("ORDR_CUST_PO") & String.Empty
                Else
                    With ASCMAIN1.CodeSelector
                        .Caption = "Select Customer / PO"
                        .COLUMN_NAME = String.Empty
                        .VIEW_NAME = String.Empty
                        .SQL = ASCMAIN1.sql
                        .ParamTypes = "V"
                        .Params = New Object() {FIND_BY}
                        .MultipleSelections = False
                        .PreviouslySelectedCodes0 = ""
                    End With

                    Using F As New ASFCODE1
                        F.ShowDialog()
                    End Using

                    If ASCMAIN1.CodeSelector.SelectedRows.Count = 1 Then
                        CUST_CODE = ASCMAIN1.CodeSelector.SelectedRows(0).Item("CUST_CODE") & String.Empty
                        ORDR_CUST_PO = ASCMAIN1.CodeSelector.SelectedRows(0).Item("ORDR_CUST_PO") & String.Empty
                    Else
                        Exit Sub
                    End If

                End If

                ASCMAIN1.sql = "Select Distinct SOTORDR1.CUST_CODE, SOTORDR1.ORDR_CUST_PO, SOTSHIP1.SHIPPED_ACTUAL SHIP_DATE
                        from SOTORDR1, SOTPICK1, SOTSHIP1
                        WHERE SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO
                        AND SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO
                        AND UPPER(SOTSHIP1.BILL_OF_LADING_NO) = :PARM1
                        AND SOTORDR1.CUST_CODE = :PARM2
                        AND SOTORDR1.ORDR_CUST_PO = :PARM3"
                rows = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "VVV", New Object() {FIND_BY, CUST_CODE, ORDR_CUST_PO}).Select()
            End If

            If rows.Length = 0 Then
                MsgBox("No Customer(s) found with " & optFindBy.Text & " " & FIND_BY, MsgBoxStyle.OkOnly, "Could Not Locate a Matching Customer")
            ElseIf rows.Length = 1 Then
                txtFindBy.Text = ""
                Absx1.txtFor("CUST_CODE").Text = rows(0).Item(0)
                Click_Command("Select")

                If optFindBy.Value = "C" Then
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Rows
                        If grow.IsDataRow Then
                            If grow.Cells("ORDR_CUST_PO").Value & "" = FIND_BY Then
                                grdSOTORDR0.ActiveRow = grow
                                grdSOTORDR0.ActiveRowScrollRegion.FirstRow = grow
                                Exit For
                            End If
                        End If
                    Next
                ElseIf optFindBy.Value = "T" Then
                    Dim ORDR_CUST_PO As String = rows(0).Item("ORDR_CUST_PO") & String.Empty
                    If ORDR_CUST_PO.Length > 0 Then
                        Dim found As Boolean = False
                        For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Rows
                            If grow.IsDataRow Then
                                If grow.Cells("ORDR_CUST_PO").Value & "" = ORDR_CUST_PO Then
                                    grdSOTORDR0.ActiveRow = grow
                                    grdSOTORDR0.ActiveRowScrollRegion.FirstRow = grow
                                    found = True
                                    Exit For
                                End If
                            End If
                        Next

                        If found Then
                            For Each grow As UltraWinGrid.UltraGridRow In grdSOTCART1.Rows
                                If grow.IsDataRow Then
                                    If grow.Cells("CART_NO").Value & "" = FIND_BY Then
                                        grdSOTCART1.ActiveRow = grow
                                        grdSOTCART1.ActiveRowScrollRegion.FirstRow = grow
                                        grdSOTCART1.Selected.Rows.Clear()
                                        grdSOTCART1.Selected.Rows.Add(grow)
                                        tabDetails.SelectedTab = tabDetails.Tabs("Cartons")
                                        tabDetails.ActiveTab = tabDetails.Tabs("Cartons")
                                        found = True
                                        Exit For
                                    End If
                                End If
                            Next
                        End If

                        If Not found Then
                            MessageBox.Show($"The P.O. for the provided Carton is {ORDR_CUST_PO}.", "Find Customer By", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    End If
                ElseIf optFindBy.Value = "B" Then
                    Dim BILL_OF_LADING_NO As String = FIND_BY
                    tabDetails.SelectedTab = tabDetails.Tabs("Shipments")
                    tabDetails.ActiveTab = tabDetails.Tabs("Shipments")
                    Clear_All_Filters(grdSOTSHIP1)
                    grdSOTSHIP1.DisplayLayout.Bands(0).ColumnFilters("BILL_OF_LADING_NO").FilterConditions.Add(UltraWinGrid.FilterComparisionOperator.Equals, BILL_OF_LADING_NO)
                    Show_Filter(grdSOTSHIP1)
                End If
            Else
                grdSOTORDR0.Tag = "ORDR_CUST_PO"
                Load_SOTORDR0(FIND_BY)
            End If
        End If
    End Sub

    Private Sub cmbSALES_DIVISION_CODE_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles cmbSALES_DIVISION_CODE.KeyDown
        If e.KeyCode = Keys.Delete Then
            cmbSALES_DIVISION_CODE.Value = ""
        End If
    End Sub

    Private Sub grdSOTORDR0_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDR0.InitializeRow

        ' 02/03/2023
        ' Code in location where SOTORDR0 is loaded to set a status in an new column GROUP_STATUS
        If e.Row.IsDataRow Then
            If Val(e.Row.Cells("ORDR_CNT_OPEN").Value & "") <> 0 Then
                e.Row.Cells("ORDR_GROUP_NO").ToolTipText = "Some or All Orders are still Open"
                e.Row.Cells("ORDR_GROUP_NO").Appearance.ForeColor = Drawing.Color.Green
            ElseIf Val(e.Row.Cells("ORDR_CNT_PICK").Value & "") <> 0 Then
                e.Row.Cells("ORDR_GROUP_NO").ToolTipText = "Some or All Orders are In Pick"
                e.Row.Cells("ORDR_GROUP_NO").Appearance.ForeColor = Drawing.Color.Blue
            ElseIf Val(e.Row.Cells("ORDR_QTY_SHIP").Value & "") = 0 Then
                e.Row.Cells("ORDR_GROUP_NO").ToolTipText = "Cancelled Order"
                e.Row.Cells("ORDR_GROUP_NO").Appearance.ForeColor = Drawing.Color.Orange
            End If

            If e.Row.Cells("ORDR_HOLD").Value & "" = "1" Then
                e.Row.Cells("ORDR_CUST_PO").Appearance.BackColor = Drawing.Color.Yellow
                e.Row.Cells("ORDR_CUST_PO").ToolTipText = "Order on Hold"
            End If

            If e.Row.Cells("ERRORS").Value & "" <> "" Then
                e.Row.Cells("ORDR_GROUP_NO").ToolTipText = "Errors Reported from 3PL" & vbCrLf & e.Row.Cells("ERRORS").Value
                e.Row.Cells("ORDR_GROUP_NO").Appearance.ForeColor = Drawing.Color.Red
            End If

            Dim isOrange As Boolean = False
            If e.Row.Cells("ORDR_HIGH_PRIORITY").Value & "" = "1" Then
                e.Row.Appearance.BackColor = Drawing.Color.Orange
                isOrange = True
            End If

            Dim isReleased As Boolean = False
            If ASCMAIN1.USER_ID <> "wjz" AndAlso (ASCMAIN1.CLIENT = "INT" Or ASCMAIN1.CLIENT = "AHA") Then
                If dst.Tables("SOTSHIPX").Rows.Find(e.Row.Cells("ORDR_GROUP_NO").Value) IsNot Nothing Then
                    e.Row.Appearance.BackColor = Drawing.Color.LightGreen
                    isReleased = True
                End If
            End If

            If Not isReleased AndAlso e.Row.Cells("REVERSE_PO").Value & "" = "1" Then
                e.Row.Appearance.BackColor = Drawing.Color.LightBlue
            End If
        End If
    End Sub


    Private Sub Generate_History()
        dst.Tables("SOTCORDD").Rows.Clear()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Collecting History")
        Dim RYP As String = cmb12Months.Value
        RYP = Mid(RYP, 1, 4) & Mid(RYP, 6, 2)
        Dim YPs(12) As String
        For i As Integer = 1 To 12
            YPs(i) = ASCMAIN1.Period_Calc(RYP, i - 12)
            Dim LEGEND = ASCMAIN1.Get_Legend(YPs(i))
            grdSOTCORDX.DisplayLayout.Bands(0).Columns("V" & Format(i, "00")).Header.Caption = Mid(LEGEND, 10, 6)
            grdSOTCORDD.DisplayLayout.Bands(0).Columns("V" & Format(i, "00")).Header.Caption = Mid(LEGEND, 10, 6)
            grdSOTCORDD.DisplayLayout.Bands(0).Columns("V" & Format(i, "00")).Tag = YPs(i)
        Next i

        ' this ought to be processed by a datareader with 1 pass thru all of styles

        For Each SORT_SEQ As String In New String() {"1", "2", "4", "9"}
            Dim exp As String = ""
            Dim sqlexpw As String = ""
            If SORT_SEQ = "1" Then
                exp = "DECODE(SOTINVH2.INV_TYPE,'I',(SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE),0)"
                sqlexpw = " and SOTINVH2.INV_TYPE = 'I' and SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE <> 0"
            ElseIf SORT_SEQ = "2" Then
                exp = "DECODE(SOTINVH2.INV_TYPE,'C',(SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE),0)"
                sqlexpw = " and SOTINVH2.INV_TYPE = 'C' and SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE <> 0"
            ElseIf SORT_SEQ = "9" Then
                exp = "ORDR_QTY_SHIP"
                sqlexpw = " and SOTINVH2.ORDR_QTY_SHIP <> 0"
            ElseIf SORT_SEQ = "4" Then
                exp = "SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ITEM_UNIT_COST"
                sqlexpw = " and SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ITEM_UNIT_COST <> 0"
            End If
            ASCMAIN1.sql = "Select '" & SORT_SEQ & "' SORT_SEQ, ITEM_CODE CODE_VALUE" & vbCrLf
            For i As Integer = 1 To 12
                ASCMAIN1.sql &= ", SUM (DECODE (ORDR_YYYYPP_UPDATED,'" & YPs(i) & "'," & exp & ",0)) V" & Format(i, "00") & vbCrLf
            Next i
            ASCMAIN1.sql &= " from SOTINVH2" & vbCrLf
            ASCMAIN1.sql &= " where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf
            ASCMAIN1.sql &= "   and ORDR_YYYYPP_UPDATED >= '" & ASCMAIN1.Period_Calc(RYP, -11) & "' AND ORDR_YYYYPP_UPDATED <= '" & RYP & "'" & vbCrLf
            ASCMAIN1.sql &= sqlexpw & vbCrLf
            ASCMAIN1.sql &= " group by ITEM_CODE" & vbCrLf
            If SORT_SEQ <> "4" Or ASCMAIN1.USER_SECURITY_CODEs.Contains("X2") Then
                dst.Tables("SOTCORDD").Merge(ASCDATA1.GetDataTable)
                'Fill_Records("SOTCORDD", "", False, ASCMAIN1.sql)
            End If
        Next

        ASCMAIN1.sql = "Select '7' SORT_SEQ, ARTPYMT5.REASON_CODE CODE_VALUE" & vbCrLf
        For i As Integer = 1 To 12
            ASCMAIN1.sql &= " , SUM (DECODE (ARTPYMT1.OPS_YYYYPP,'" & YPs(i) & "',ARTPYMT5.GL_DIST_AMT,0)) V" & Format(i, "00") & vbCrLf
        Next i
        ASCMAIN1.sql &= "" _
        & " from ARTPYMT1, ARTPYMT2, ARTPYMT5" & vbCrLf _
        & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
        & "   and ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
        & "   and ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf _
        & "   and Decode(ARTPYMT5.CUST_CODE_SO,Null,ARTPYMT2.CUST_CODE,ARTPYMT5.CUST_CODE_SO) = '" & CUST_CODE & "'" & vbCrLf _
        & "   and NVL(ARTPYMT5.CHARGEBACK_IND,'0') = '0'" & vbCrLf _
        & "   and ARTPYMT1.OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(RYP, -11) & "' and ARTPYMT1.OPS_YYYYPP <= '" & RYP & "'" & vbCrLf _
        & " group by ARTPYMT5.REASON_CODE" & vbCrLf
        dst.Tables("SOTCORDD").Merge(ASCDATA1.GetDataTable)
        ' Fill_Records("SOTORDDD", "", False, ASCMAIN1.sql)


        ASCMAIN1.sql = "Select '6' SORT_SEQ, NVL(SOTINVH1.REASON_CODE,'?') CODE_VALUE" & vbCrLf
        For i As Integer = 1 To 12
            ASCMAIN1.sql &= " , SUM (DECODE (SOTINVH1.ORDR_YYYYPP_UPDATED,'" & YPs(i) & "',NVL(SOTINVH1.INV_TOTAL_AMOUNT,0),0)) V" & Format(i, "00") & vbCrLf
        Next i
        ASCMAIN1.sql &= "" _
            & " from SOTINVH1" & vbCrLf _
            & " where SOTINVH1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and (SOTINVH1.ORDR_TYPE_CODE = 'TOP' or SOTINVH1.ORDR_TYPE_CODE = 'DIF')" & vbCrLf _
            & "   and SOTINVH1.ORDR_YYYYPP_UPDATED >= '" & ASCMAIN1.Period_Calc(RYP, -11) & "' and SOTINVH1.ORDR_YYYYPP_UPDATED <= '" & RYP & "'" & vbCrLf _
            & "   and SOTINVH1.REASON_CODE is Not Null" & vbCrLf _
            & " group by NVL(SOTINVH1.REASON_CODE,'?')" & vbCrLf
        dst.Tables("SOTCORDD").Merge(ASCDATA1.GetDataTable)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Set_History_Summary()

        spl12Months.Visible = True
    End Sub

    Sub Set_History_Summary()
        dst.Tables("SOTCORDX").Rows.Clear()

        Dim V(10, 12) As Decimal
        For i As Integer = 1 To 12
            Dim PP As String = Format(i, "00")
            V(1, i) = Val(dst.Tables("SOTCORDD").Compute("SUM(V" & PP & ")", "SORT_SEQ = '1'") & "")
            V(2, i) = Val(dst.Tables("SOTCORDD").Compute("SUM(V" & PP & ")", "SORT_SEQ = '2'") & "")
            V(3, i) = V(1, i) + V(2, i)
            V(4, i) = Val(dst.Tables("SOTCORDD").Compute("SUM(V" & PP & ")", "SORT_SEQ = '4'") & "")
            V(5, i) = V(3, i) - V(4, i)
            V(6, i) = Val(dst.Tables("SOTCORDD").Compute("SUM(V" & PP & ")", "SORT_SEQ = '6'") & "")
            V(7, i) = Val(dst.Tables("SOTCORDD").Compute("SUM(V" & PP & ")", "SORT_SEQ = '7'") & "")
            V(8, i) = V(5, i) + V(6, i) - V(7, i)
            V(9, i) = Val(dst.Tables("SOTCORDD").Compute("SUM(V" & PP & ")", "SORT_SEQ = '9'") & "")
            If V(9, i) <> 0 Then V(10, i) = V(3, i) / V(9, i)
        Next i

        Dim S(10) As String
        S(1) = "Gross Sales"
        S(2) = "Returns"
        S(3) = "Net Sales"
        S(4) = "CGS"
        S(5) = "GP"
        S(6) = "Credits"
        S(7) = "Deductions"
        S(8) = "Net Profit"
        S(9) = "Units"
        S(10) = "Price"

        Dim rowSOTCORDX As DataRow

        For j As Integer = 1 To 9
            If ASCMAIN1.USER_SECURITY_CODEs.Contains("X2") Or j = 1 Or j = 2 Or j = 3 Or j = 7 Then
                rowSOTCORDX = dst.Tables("SOTCORDX").NewRow
                rowSOTCORDX.Item("SORT_SEQ") = Format(j, "0")
                rowSOTCORDX.Item("CODE_VALUE") = S(j)
                For i As Integer = 1 To 12
                    Dim PP As String = Format(i, "00")
                    rowSOTCORDX.Item("V" & PP) = V(j, i)
                Next i
                dst.Tables("SOTCORDX").Rows.Add(rowSOTCORDX)
            End If
        Next j

        rowSOTCORDX = dst.Tables("SOTCORDX").Rows.Find("1")
        Dim TOT_SALES = Val(rowSOTCORDX.Item("TOT") & "")
        Dim YTD_SALES = Val(rowSOTCORDX.Item("YTD") & "")
        dst.Tables("SOTCORDX").Columns("TOTPCT").Expression = IIf(TOT_SALES = 0, "0", "100 * TOT / " & CStr(TOT_SALES))
        dst.Tables("SOTCORDX").Columns("YTDPCT").Expression = IIf(YTD_SALES = 0, "0", "100 * YTD / " & CStr(YTD_SALES))
        CreateGraph_SATCSLS1_X()
    End Sub


    Private Sub grdSOTPICK1_BeforeRowExpanded(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTPICK1.BeforeRowExpanded
        If e.Row.IsDataRow Then
            Dim PICK_NO As String = e.Row.Cells("PICK_NO").Value & ""
            Fill_Records("SOTPICK2", PICK_NO)
            Sort_grdColumns(grdSOTPICK1, "PICK_LNO", False, 1)
        End If
    End Sub

    Private Sub grdSOTPICK1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTPICK1.InitializeRow
        If e.Row.Band.Key = "SOTPICK1" Then
            If e.Row.Cells("PICK_FORCED").Value & "" = "1" Then
                e.Row.Cells("PICK_STATUS").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("PICK_STATUS").ToolTipText = "Force Picked"
            End If
            If e.Row.Cells("PICK_NO_REV").Value & "" <> "" Then
                e.Row.Cells("PICK_NO").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("PICK_STATUS").ToolTipText = "Reversed"
            End If
        End If
    End Sub

    Private Sub grdSOTSHIP1_BeforeRowExpanded(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTSHIP1.BeforeRowExpanded
        If e.Row.IsGroupByRow Then
        Else


            Dim SHIP_BOL_NO As String = e.Row.Cells("SHIP_BOL_NO").Value & ""
            Fill_Records("SOTSHIP2", SHIP_BOL_NO)
            Sort_grdColumns(grdSOTSHIP1, "ITEM_CODE", False, 1)
        End If
    End Sub

    Private Sub grdSOTSHIP1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTSHIP1.InitializeRow
        If e.Row.Band.Key = "SOTSHIP1" Then
            If e.Row.Cells("SHIP_BOL_NO_REV").Value & "" <> "" Then
                e.Row.Cells("SHIP_BOL_NO").Appearance.ForeColor = Drawing.Color.Red
            End If
        End If
    End Sub

    Private Sub grdSOTCORDX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTCORDX.AfterRowActivate
        Setup_SOTCORDD()
        If tabMonth.SelectedTab.Key = "Details" Then tabMonth.SelectedTab = tabMonth.Tabs("Summary")
    End Sub

    Sub Setup_SOTCORDD()
        If grdSOTCORDX.ActiveRow Is Nothing Then
            tabMonth.Visible = False
        Else
            Dim TOT_SALES = Val(grdSOTCORDX.ActiveRow.Cells("TOT").Value & "")
            Dim YTD_SALES = Val(grdSOTCORDX.ActiveRow.Cells("YTD").Value & "")
            dst.Tables("SOTCORDD").Columns("TOTPCT").Expression = IIf(TOT_SALES = 0, "0", "100 * TOT / " & CStr(TOT_SALES))
            dst.Tables("SOTCORDD").Columns("YTDPCT").Expression = IIf(YTD_SALES = 0, "0", "100 * YTD / " & CStr(YTD_SALES))

            Dim dvw As DataView = DirectCast(grdSOTCORDD.DataSource, DataTable).DefaultView
            Dim SORT_SEQ As String = grdSOTCORDX.ActiveRow.Cells("SORT_SEQ").Value
            dvw.RowFilter = "SORT_SEQ = '" & SORT_SEQ & "'"
            Sort_grdColumns(grdSOTCORDD, "CODE_VALUE")
            tabMonth.Visible = True
        End If
    End Sub

    Private Sub cmdGenerateHistory_Click(sender As System.Object, e As System.EventArgs) Handles cmdGenerateHistory.Click
        Generate_History()
    End Sub

    Private Sub chkShowSelectedOrder_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShowSelectedOrder.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_SOTORDRS()
        If tabItems.SelectedTab.Key = "Summary" Then tabItems.SelectedTab = tabItems.Tabs("Detail")
        tabItems.Tabs("Summary").Visible = Not chkShowSelectedOrder.Checked
    End Sub

    Private Sub grdSOTORDRX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTORDRX.AfterRowActivate
        Setup_SOTORDRS()
    End Sub

    Sub Setup_Summary_SOTORDRM(ORDR_GROUP_NO As String)
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Setting up Summary by Store")

        Dim COLUMN_NAME As String = optQTY.Value

        ASCMAIN1.sql = "Select SOTORDR2.ITEM_CODE"
        ASCMAIN1.sql &= ",Sum (" & COLUMN_NAME & ") QTY"
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTORDR1"), New String() {"CUST_STORE_NO"}).Select("", "CUST_STORE_NO")
            Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
            ASCMAIN1.sql &= ", Sum (Decode(SOTORDR1.CUST_STORE_NO,'" & CUST_STORE_NO & "'," & COLUMN_NAME & ",0)) QTY_" & CUST_STORE_NO
        Next
        'For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("", "CUST_STORE_NO")
        '    Dim CUST_STORE_NO As String = rowSOTORDR1.Item("CUST_STORE_NO")
        '    ASCMAIN1.sql &= ", Sum (Decode(SOTORDR1.CUST_STORE_NO,'" & CUST_STORE_NO & "'," & COLUMN_NAME & ",0)) QTY_" & CUST_STORE_NO
        'Next
        ASCMAIN1.sql &= " from SOTORDR1,SOTORDR2" _
            & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
            & "   and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" _
            & "   and SOTORDR1.ORDR_STATUS <> 'D'" _
            & " group by SOTORDR2.ITEM_CODE"
        grdSOTORDRM.DataSource = Nothing
        grdSOTORDRM.DisplayLayout.Bands(0).Summaries.Clear()
        grdSOTORDRM.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
        dst.Tables.Remove("SOTORDRM")
        Dim t As DataTable = ASCDATA1.GetDataTable
        t.TableName = "SOTORDRM"
        dst.Tables.Add(t)
        grdSOTORDRM.DataSource = t
        For Each gcol As UltraWinGrid.UltraGridColumn In grdSOTORDRM.DisplayLayout.Bands(0).Columns
            If gcol.Key = "ITEM_CODE" Then
                gcol.Width = 90
                gcol.Header.Caption = "Item"
                Create_Summary(grdSOTORDRM, "ITEM_CODE", "Count")
            ElseIf gcol.Key = "QTY" Then
                gcol.Width = 70
                gcol.Header.Caption = "Total"
                gcol.Format = "#,##0"
                Create_Summary(grdSOTORDRM, "QTY")
            Else
                gcol.Width = 70
                gcol.Header.Caption = Mid(gcol.Key, 5)
                gcol.Format = "#,##0"
                Create_Summary(grdSOTORDRM, gcol.Key)
            End If
        Next

        grdSOTORDRM.Text = "Order Group " & ORDR_GROUP_NO & ", Item Summary by Store, " & optQTY.Text

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Setup_Summary()
        If SELECTION_NO = 0 Then Exit Sub
        If tabDetails.SelectedTab.Key = "Items" AndAlso tabItems.SelectedTab.Key = "Summary" AndAlso grdSOTORDR0.ActiveRow IsNot Nothing Then
            Dim ORDR_GROUP_NO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_GROUP_NO").Value
            If Not chkShowSelectedOrder.Checked Then
                Setup_Summary_SOTORDRM(ORDR_GROUP_NO)
                grdSOTORDRM.Visible = True
            Else
                grdSOTORDRM.Visible = False
            End If
        End If
    End Sub

    Sub Load_SOTORDRS()
        If Not chkShowSelectedOrder.Checked And grdSOTORDR0.ActiveRow Is Nothing Or Not grdSOTORDR0.ActiveRow.IsDataRow Then
            tabDetails.Visible = False
        ElseIf chkShowSelectedOrder.Checked And grdSOTORDRX.ActiveRow Is Nothing Then
            tabDetails.Visible = False
        Else
            Setup_SOTORDRS()

            Dim ORDR_CUST_PO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_CUST_PO").Value & ""
            Dim ORDR_GROUP_NO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_GROUP_NO").Value

            With grdSOTORDRS.DisplayLayout.Bands(0)
                .Columns("ORDR_QTY_SHIP").Header.Caption = "#Ship"
            End With
            Setup_Summary()
        End If
        If tabDetails.SelectedTab.Key = "Cartons" Then
            Setup_Cartons()
        End If
    End Sub

    Sub Setup_SOTORDRS()
        If Not grdSOTORDR0.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Dim ORDR_CUST_PO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_CUST_PO").Value & ""
        Dim ORDR_GROUP_NO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_GROUP_NO").Value

        ASCMAIN1.Progress("Now Getting Item Details")

        Dim sql As String = ""
        If Not chkShowSelectedOrder.Checked Then
            sql = Replace(sqlSOTORDRS, " group by ", " and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' group by ")
            grdSOTORDRS.Text = "Item Summary for Order Group " & ORDR_GROUP_NO & ", Customer PO " & ORDR_CUST_PO
        Else
            Dim ORDR_NO As String = grdSOTORDRX.ActiveRow.Cells("ORDR_NO").Value
            Dim CUST_STORE_NO As String = grdSOTORDRX.ActiveRow.Cells("CUST_STORE_NO").Value
            sql = Replace(sqlSOTORDRS, " group by ", " and SOTORDR1.ORDR_NO = '" & ORDR_NO & "' group by ")
            grdSOTORDRS.Text = "Item Details for Order No " & ORDR_NO & ", Customer PO " & ORDR_CUST_PO & ", Store No " & CUST_STORE_NO
        End If

        Fill_Records("SOTORDRS", "", True, sql)
        Sort_grdColumns(grdSOTORDRS, "ITEM_CODE")

        ASCMAIN1.Progress("")

        tabDetails.Visible = True
    End Sub

    Private Sub tabMain_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        If e.Tab.Key = "Orders" Then
        ElseIf e.Tab.Key = "Status && Tracking" Then
            Set_Summary_Splitter()
        Else
            If Not spl12Months.Visible Then
                Generate_History()
            End If
        End If

        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()
        With UltraExplorerBar1
            .Groups("12 Month History").Visible = (tabMain.SelectedTab.Key = "12 Mos")
            .Groups("Show Orders").Visible = (tabMain.SelectedTab.Key = "Orders")
            .Groups("Items").Visible = (tabMain.SelectedTab.Key = "Orders") And (tabDetails.SelectedTab.Key = "Items")
            .Groups("Status && Tracking").Visible = ScreenMode And (tabMain.SelectedTab.Key = "Status && Tracking")
        End With

    End Sub

    Private Sub tabDetails_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabDetails.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabMain()
        Setup_Summary()

        If tabDetails.SelectedTab.Key = "Cartons" Then
            Setup_Cartons()
        End If

    End Sub

    Sub Setup_Cartons()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now loading Carton Details")

        EnforceConstraints(False)
        Fill_Records("SOTCART1", ORDR_GROUP_NO)
        Fill_Records("SOTCART2", ORDR_GROUP_NO)
        Fill_Records("SOTCART3", ORDR_GROUP_NO)
        EnforceConstraints(True)
        Sort_grdColumns(grdSOTCART1, "CART_NO")

        grdSOTCART1.Text = "Cartons on All Shipments for Order Group " & ORDR_GROUP_NO

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub optOrders_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optOrders.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_SOTORDR0("", CUST_CODE)
    End Sub

    Private Sub chkReservations_CheckedChanged(sender As System.Object, e As System.EventArgs)
        If SELECTION_NO = 0 Then Exit Sub
        Load_SOTORDR0("", CUST_CODE)
    End Sub

    Private Sub cmbSALES_DIVISION_CODE_ValueChanged(sender As Object, e As System.EventArgs) Handles cmbSALES_DIVISION_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_SOTORDR0("", CUST_CODE)
    End Sub

    Private Sub grdSOTCORDD_DoubleClickCell(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdSOTCORDD.DoubleClickCell
        If e.Cell.Row.IsDataRow Then
            Dim SORT_SEQ As String = e.Cell.Row.Cells("SORT_SEQ").Value
            If SORT_SEQ = "1" Or SORT_SEQ = "9" Then
                Dim ITEM_CODE As String = e.Cell.Row.Cells("CODE_VALUE").Value
                Dim YP As String = e.Cell.Column.Tag
                Fill_Records("SOTCORDY", New String() {CUST_CODE, ITEM_CODE, YP})
                grdSOTCORDY.Text = grdSOTCORDX.ActiveRow.Cells("CODE_VALUE").Value & " Invoice Details for Customer " & CUST_CODE & ", Item " & ITEM_CODE & ", in " & e.Cell.Column.Header.Caption
                tabMonth.Tabs("Details").Visible = True
                tabMonth.SelectedTab = tabMonth.Tabs("Details")
            End If
        End If
    End Sub

    Sub CreateGraph_SATCSLS1_X()

        Dim chtIsVisible As Boolean = chtSATCSLS1_X.Visible
        chtSATCSLS1_X.Visible = False

        chtSATCSLS1_X.DataSource = Nothing

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Charting Data")

        Me.SuspendLayout()

        Dim RL() As String
        Dim CL() As String
        Dim Periods As Int32 = 12
        ReDim CL(Periods)

        'this will be necessary for line graph
        'For i As Integer = MOSMAX To 0 Step -1
        '    Dim L As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * i))
        '    CL(MOSMAX - i) = Mid(L, 10, 6)
        '    grdSATCSLS1.DisplayLayout.Bands(0).Columns("M" & Format(i, "00")).Header.Caption = Mid(L, 10, 3)
        'Next
        For i As Integer = 1 To Periods
            'Dim L As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * i))
            CL(i - 1) = grdSOTCORDX.DisplayLayout.Bands(0).Columns("V" & Format(i, "00")).Header.Caption
            'grdSATCSLS1.DisplayLayout.Bands(0).Columns("M" & Format(i, "00")).Header.Caption = Mid(L, 10, 3)
        Next

        'chtICTINVAT.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.LabelPlusDataValue
        'chtICTINVAT.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom

        chtSATCSLS1_X.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtSATCSLS1_X.LabelHash = labelHash

        chtSATCSLS1_X.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtSATCSLS1_X.Tooltips.FormatString = "<HIGHLOW>"

        Dim DT As New DataTable
        DT.Columns.Add("CODE_VALUE")
        DT.Columns.Add("DESC_VALUE")
        For P As Integer = 1 To Periods
            DT.Columns.Add("V" & Format(P, "00"), GetType(System.Decimal))
        Next

        Dim RLi As Integer = 0
        ReDim RL(dst.Tables("SOTCORDX").Rows.Count - 1)
        For Each row As DataRow In dst.Tables("SOTCORDX").Select("", "CODE_VALUE")
            RL(RLi) = row("CODE_VALUE") ' & ":" & row("DESC_VALUE")
            RLi += 1

            Dim rowDT As DataRow = DT.NewRow
            rowDT.Item("CODE_VALUE") = row("CODE_VALUE")
            'rowDT.Item("DESC_VALUE") = row("DESC_VALUE")
            For P As Integer = 1 To Periods
                rowDT.Item("V" & Format(P, "00")) = row("V" & Format(P, "00"))
            Next
            DT.Rows.Add(rowDT)
        Next
        chtSATCSLS1_X.Data.SetRowLabels(RL)
        chtSATCSLS1_X.Data.SetColumnLabels(CL)

        chtSATCSLS1_X.DataSource = DT
        'chtSATCSLS1_X.Data.IncludeColumn("CODE_VALUE", False)
        'chtSATCSLS1_X.Data.IncludeColumn("DESC_VALUE", False)
        'chtSATCSLS1_X.Data.IncludeColumn("P00", False)

        chtSATCSLS1_X.DataBind()

        chtSATCSLS1_X.Visible = True ' chtIsVisible

        Me.ResumeLayout()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()
    End Sub

    Private Sub tabItems_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabItems.SelectedTabChanged
        grdSOTORDRX.Visible = Not (tabItems.SelectedTab.Key = "Summary")
        optQTY.Visible = (tabItems.SelectedTab.Key = "Summary")
        Setup_Summary()
    End Sub

    Private Sub optQTY_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optQTY.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_Summary()
    End Sub

    Sub Store_Configuration_Report(ByVal ORDR_GROUP_NO As String)
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Store Configuration Report")

        Dim RPT As String = "SORCONF1"
        Dim sqlw As String = " AND SOTORDR0.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
        If Not REPORTS.ContainsKey(RPT) Then
            REPORTS.Add(RPT, Load_rptClass(RPT))
            REPORTS(RPT).Prepare_dst(True, sqlw)

        Else
            REPORTS(RPT).Fill_Records_RPT(sqlw)
        End If


        Dim FILENAME As String = ""
        With REPORTS(RPT).clsASCBASE1
            .Print_Report_Begin()
            .CR_params.Add("SUBT", "")
            'Dim REPORT_NO As String = .Generate_Report(RPT, "", "", False, False, "", "PDF", ORDR_GROUP_NO, False)
            'FILENAME = .F.REPORT_FILENAMES(REPORT_NO)
            .Generate_Report(RPT, "Store Configuration Report", "", True)
            .Print_Report_End()
            ' .Print_Report_End(, True)
        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        ' Return FILENAME
    End Sub

    Function Create_Invoice(
                           SHIP_BOL_NO As String,
                           PICK_NO As String,
                           Optional make_pdf As Boolean = False,
                           Optional pro_forma As Boolean = False) As String

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Invoice")

        Dim REPORTFILE As String = "SORINVP1"
        Dim RPT As String = ROWs("ARTPARM1").Item("AR_PARM_INVOICE_RPT") & ""
        If RPT = "" Then RPT = REPORTFILE

        If Not REPORTS.ContainsKey(REPORTFILE) Then
            REPORTS.Add(REPORTFILE, Load_rptClass(REPORTFILE))
            REPORTS(REPORTFILE).Prepare_dst(False, "")
        End If

        'To fill the report's dataset with data from Oracle, 
        ' set the parameter array to values that the Fill_Records_RPT method expects, and then call it

        Dim sqlw As String = ""
        If SHIP_BOL_NO <> "" Then
            sqlw = " and SOTINVH1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        Else
            sqlw = " and SOTINVH1.PICK_NO = '" & PICK_NO & "'"
        End If
        REPORTS(REPORTFILE).Fill_Records_RPT(New String() {sqlw, IIf(pro_forma, "1", "0")})

        'To fill the report's dataset with data from this form's dataset:
        'With REPORTS(REPORTFILE).clsASCBASE1
        '    .EnforceConstraints(False)
        '    For Each TABLE_NAME As String In New String() {"SOTPPDI1", "SOTPPDI2", "SOTPPDI3", "SOTINVH1", "SOTSVIA1"}
        '        .dst.Tables(TABLE_NAME).Rows.Clear()
        '        Dim SQL As String = ""
        '        If TABLE_NAME = "SOTINVH1" Then
        '            SQL = "ORDR_NO = '" & ORDR_NO & "'"
        '        End If

        '        For Each row As DataRow In dst.Tables(TABLE_NAME).Select(SQL)
        '            Dim rowr As DataRow = .dst.Tables(TABLE_NAME).NewRow
        '            If TABLE_NAME = "SOTPPDI2" Or TABLE_NAME = "SOTPPDI3" Or TABLE_NAME = "SOTINVH1" Then

        '                For I As Integer = 0 To .dst.Tables(TABLE_NAME).Columns.Count - 1
        '                    Dim COLUMN_NAME As String = .dst.Tables(TABLE_NAME).Columns(I).ColumnName
        '                    rowr.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
        '                Next
        '            Else
        '                rowr.ItemArray = row.ItemArray
        '            End If
        '            .dst.Tables(TABLE_NAME).Rows.Add(rowr)
        '        Next
        '    Next
        '    .EnforceConstraints(True)
        'End With

        Dim FILENAME As String = ""
        With REPORTS(REPORTFILE).clsASCBASE1
            .Print_Report_Begin()
            .CR_params.Add("SUBT", "")
            .CR_params.Add("CONS_INV", "0")
            FILENAME = .Generate_Report(RPT, "Sales Invoice", , Not make_pdf, , , IIf(make_pdf, "PDF", ""), , False)
            .Print_Report_End(make_pdf, make_pdf)
        End With

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("")

        Return FILENAME

    End Function

    Private Sub grdSOTORDR1_ClickCell(sender As Object, e As UltraWinGrid.ClickCellEventArgs) Handles grdSOTORDR1.ClickCell
        If e.Cell.Column.Key = "ORDR_REL_HOLD_CODES" Then
            Dim ORDR_REL_HOLD_CODES As String = e.Cell.Value & ""
            If ORDR_REL_HOLD_CODES <> "" Then
                Dim T As String = ""
                For I As Integer = 1 To ORDR_REL_HOLD_CODES.Length
                    Dim C As String = Mid(ORDR_REL_HOLD_CODES, I, 1)
                    T &= vbCrLf & ORDR_REL_HOLD_CODES_list(C)
                Next
                e.Cell.ToolTipText = Mid(T, 3)
            End If

        End If
    End Sub

    Private Sub chkOSS_CheckedChanged(sender As Object, e As EventArgs) Handles chkOSS.CheckedChanged
        dteOSFrom.Visible = chkOSS.Checked
        dteOSTo.Visible = chkOSS.Checked
        lblOSFrom.Visible = chkOSS.Checked
        lblOSTo.Visible = chkOSS.Checked
    End Sub

    Private Sub grdSOTORDR0_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTORDR0.InitializeLayout

    End Sub

    Private Sub cmdFetchOS_Click(sender As Object, e As EventArgs) Handles cmdFetchOS.Click

        ASCMAIN1.Progress("Now Fetching Order Information")

        If Not chkOSO.Checked And Not chkOSS.Checked Then Exit Sub
        If chkOSS.Checked Then
            If Format(dteOSFrom.Value, "yyyyMMdd") > Format(dteOSTo.Value, "yyyyMMdd") Then
                MsgBox("From and To Dates not chronological", MsgBoxStyle.OkOnly, "Cannot Process Request")
                Exit Sub
            End If
        End If

        ASCDATA1.ExecuteSQL("Truncate Table " & SOTORDRT)
        If chkOSS.Checked Then
            ASCMAIN1.sql = "Insert into " & SOTORDRT & vbCrLf _
                & " Select SOTCART2.*,ICTITEM1.ITEM_DESC, SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_QTY_CANC" & vbCrLf _
                & " from SOTCART2,ICTITEM1,SOTORDR2" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = SOTCART2.ITEM_CODE and SOTCART2.ORDR_NO in " & vbCrLf _
                & "(Select ORDR_NO from SOTINVH1 where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                & "  and INV_DATE >= '" & Format(dteOSFrom.Value, "dd-MMM-yyyy") & "'" & vbCrLf _
                & "  and INV_DATE <= '" & Format(dteOSTo.Value, "dd-MMM-yyyy") & "')" & vbCrLf _
                & "  and SOTORDR2.ORDR_NO  = SOTCART2.ORDR_NO" & vbCrLf _
                & "  and SOTORDR2.ORDR_LNO = SOTCART2.ORDR_LNO"
            ASCDATA1.ExecuteSQL()
            If chkOSC.Checked Then
                ASCMAIN1.sql = "Insert into " & SOTORDRT & vbCrLf _
                    & " Select SOTORDR2.ORDR_NO CART_NO, SOTORDR2.ORDR_LNO CART_LNO, SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO" & vbCrLf _
                    & ", 0 QTY_PACKED" & vbCrLf _
                    & ", ICTITEM1.ITEM_UPC_CODE, ICTITEM1.ITEM_EAN_CODE, SOTORDR2.ITEM_CODE" & vbCrLf _
                    & ", 0 ITEM_TOTAL_WGT_CALC, 0 ITEM_TOTAL_VOL_CALC, ICTITEM1.ITEM_DESC, SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_QTY_CANC" & vbCrLf _
                    & " from SOTORDR2,ICTITEM1 where ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE" & vbCrLf _
                    & " and SOTORDR2.ORDR_NO in " & vbCrLf _
                    & "(Select ORDR_NO from SOTORDR1 where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                    & "  and ORDR_STATUS >= 'C' and ORDR_STATUS <= 'F'" & vbCrLf _
                    & "  and ORDR_DATE_CLOSED >= '" & Format(dteOSFrom.Value, "dd-MMM-yyyy") & "'" & vbCrLf _
                    & "  and ORDR_DATE_CLOSED <= '" & Format(dteOSTo.Value, "dd-MMM-yyyy") & "')" & vbCrLf _
                    & "  and SOTORDR2.ORDR_QTY_SHIP = 0"
                ASCDATA1.ExecuteSQL()
            End If

        End If

        If chkOSO.Checked Then
            ASCMAIN1.sql = "Insert into " & SOTORDRT & vbCrLf _
                & " Select SOTORDR2.ORDR_NO CART_NO, SOTORDR2.ORDR_LNO CART_LNO, SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO" & vbCrLf _
                & ", DECODE(SOTORDR2.ORDR_STATUS,'O',SOTORDR2.ORDR_QTY_OPEN,'P',SOTORDR2.ORDR_QTY_PICK,0) QTY_PACKED" & vbCrLf _
                & ", ICTITEM1.ITEM_UPC_CODE, ICTITEM1.ITEM_EAN_CODE, SOTORDR2.ITEM_CODE" & vbCrLf _
                & ", 0 ITEM_TOTAL_WGT_CALC, 0 ITEM_TOTAL_VOL_CALC, ICTITEM1.ITEM_DESC, SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_QTY_CANC" & vbCrLf _
                & " from SOTORDR2,ICTITEM1 where ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE" & vbCrLf _
                & " and SOTORDR2.ORDR_NO in " & vbCrLf _
                & "(Select ORDR_NO from SOTORDR1 where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                & "  and ORDR_STATUS >= 'O' and ORDR_STATUS <= 'P')"
            ASCDATA1.ExecuteSQL()
        End If

        EnforceConstraints(False)
        dst.Tables("SOTORDRT").Rows.Clear()
        Fill_Records("SOTORDRT")
        For Each row As DataRow In dst.Tables("SOTORDRT").Select _
            ("CART_TRACKING_NO IS NULL AND SHIP_REF IS NOT NULL AND QTY_PACKED <> 0")
            row.Item("CART_TRACKING_NO") = row.Item("SHIP_REF")
        Next

        dst.Tables("SOTORDRU").Rows.Clear()
        Fill_Records("SOTORDRU")

        EnforceConstraints(True)

        Sort_grdColumns(grdSOTORDRT, "ORDR_NO,ORDR_LNO")
        grdSOTORDRT.DisplayLayout.Bands(0).Columns("SHIP_REF").Hidden = True
        grdSOTORDRT.DisplayLayout.Bands(0).Columns("ORDR_QTY").Hidden = Not chkOSC.Checked
        grdSOTORDRT.DisplayLayout.Bands(0).Columns("ORDR_QTY_CANC").Hidden = Not chkOSC.Checked

        Sort_grdColumns(grdSOTORDRU, "ORDR_NO")
        grdSOTORDRU.DisplayLayout.Bands(0).Columns("SHIP_REF").Hidden = True
        grdSOTORDRU.DisplayLayout.Bands(0).Columns("ORDR_QTY").Hidden = Not chkOSC.Checked
        grdSOTORDRU.DisplayLayout.Bands(0).Columns("ORDR_QTY_CANC").Hidden = Not chkOSC.Checked

        ASCMAIN1.Progress("Order Summary")
        dst.Tables("SOTORDCC").Rows.Clear()
        For Each rowSOTORDRU As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTORDRU"), New String() {"ORDR_NO"}).Select()
            Dim ORDR_NO As String = rowSOTORDRU.Item("ORDR_NO")
            Dim TOTAL_CTNS As Integer = dst.Tables("SOTORDRU").Select("ORDR_NO = '" & ORDR_NO & "'").Count
            Dim TOTAL_WGT As Decimal = Val(dst.Tables("SOTORDRU").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "ORDR_NO = '" & ORDR_NO & "'") & "")
            Dim rowSOTORDCC As DataRow = dst.Tables("SOTORDCC").NewRow
            rowSOTORDCC.Item("ORDR_NO") = ORDR_NO
            rowSOTORDCC.Item("TOT_CTNS") = TOTAL_CTNS
            rowSOTORDCC.Item("TOT_WGT") = TOTAL_WGT
            dst.Tables("SOTORDCC").Rows.Add(rowSOTORDCC)
        Next
        Sort_grdColumns(grdSOTORDCC, "ORDR_NO")
        Set_Summary_Splitter()

        ASCMAIN1.Progress("")
    End Sub

    Private Sub chkOrdrSum_CheckedChanged(sender As Object, e As EventArgs) Handles chkOrdrSum.CheckedChanged
        Set_Summary_Splitter()
    End Sub
    Sub Set_Summary_Splitter()
        splCartonSummary.SplitterDistance = Calc_Splitter_Min_Width(grdSOTORDCC)
        splCartonSummary.Panel1Collapsed = Not chkOrdrSum.Checked
    End Sub
    Sub Export_Status_And_Tracking()
        Me.Cursor = Cursors.WaitCursor
        Dim exportDict As New Dictionary(Of UltraWinGrid.UltraGrid, String)
        exportDict.Add(grdSOTORDCC, "Summary by Order")
        exportDict.Add(grdSOTORDRT, "Sales Order Details")
        exportDict.Add(grdSOTORDRU, "Summary by Carton")
        Export_Grids_To_Excel(exportDict)
    End Sub
    Function Calc_Splitter_Min_Width(grd As UltraWinGrid.UltraGrid) As Integer
        Dim minWidth As Decimal = grd.DisplayLayout.ScrollBarLook.VerticalScrollBarWidth + grd.DisplayLayout.Override.RowSelectorWidth
        grd.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        For Each col As Infragistics.Win.UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
            minWidth += col.Width
        Next
        Return CInt(minWidth) + 2
    End Function

#Region "EDI 855"

    Private Sub Generate855s(ByVal lstOrderGroups As List(Of String))

        Try
            Dim numRecs As Int32 = 0
            For Each ORDR_GROUP_NO As String In lstOrderGroups
                If ORDR_GROUP_NO.Length > 0 Then
                    ' Lock the Order Group
                    If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO, , False, 1) Then
                        ASCMAIN1.MultiTask_Release(, , 1)
                        MessageBox.Show($"Could Not Lock Order Group {ORDR_GROUP_NO} it will be skipped.", "Generate 855s", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Continue For
                    End If

                    TAC.EDC855O1.Generate_855(clsASCBASE1, ORDR_GROUP_NO, String.Empty)
                    If dst.Tables("EDT855O1").Rows.Count > 0 Then
                        ASCMAIN1.sql = "SELECT * FROM SOTORDR1 WHERE ORDR_GROUP_NO = :PARM1 AND ORDR_STATUS = 'O'"
                        Dim tblSOTORDR1 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTORDR1", "V", New Object() {ORDR_GROUP_NO})

                        For Each rowSOTORDR1 As DataRow In tblSOTORDR1.Select("")
                            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO") & String.Empty
                            Try
                                TAC.SOCMAIN1.Record_Event_SOTORDR1(ORDR_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "E855", "Manual EDI 855")
                            Catch ex As Exception
                            End Try
                        Next
                    End If
                    numRecs += dst.Tables("EDT855O1").Rows.Count
                End If
            Next

            MessageBox.Show($"{numRecs} EDI 855 records generated.", "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"Error Generating EDI 855s: {ex.Message}", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Finally
            ASCMAIN1.MultiTask_Release(, , 1)
        End Try
    End Sub


    Private Sub optGROUP_STATUS_ValueChanged(sender As Object, e As EventArgs) Handles optGROUP_STATUS.ValueChanged

        If Not dst.Tables.Contains("SOTORDR0") Then
            Exit Sub
        End If

        Dim sqlFilter As String = String.Empty

        Select Case optGROUP_STATUS.Value
            Case "A"
                sqlFilter = String.Empty
                grdSOTORDR0.Text = "All Orders"
            Case "O"
                sqlFilter = "ORDR_CNT_OPEN <> 0"
                grdSOTORDR0.Text = "Open Orders"
            Case "P"
                sqlFilter = "ORDR_CNT_PICK <> 0"
                grdSOTORDR0.Text = "Orders In Pick"
            Case "OP"
                sqlFilter = "ORDR_CNT_OPEN <> 0 OR ORDR_CNT_PICK <> 0"
                grdSOTORDR0.Text = "Open or Pick Orders"
            Case "OAP"
                sqlFilter = "ORDR_CNT_OPEN <> 0 AND ORDR_CNT_PICK <> 0"
                grdSOTORDR0.Text = "Open and Pick Orders"
            Case "R"
                sqlFilter = "ORDR_TYPE_CODE = 'R'"
                grdSOTORDR0.Text = "Reserved Orders"
            Case "C"
                sqlFilter = "ORDR_QTY_SHIP = 0"
                grdSOTORDR0.Text = "Cancelled Orders"
            Case "E"
                sqlFilter = "ISNULL(ERRORS, '') <> ''"
                grdSOTORDR0.Text = "Orders with Errors"
            Case "H"
                sqlFilter = "ORDR_HOLD = '1'"
                grdSOTORDR0.Text = "Orders on Hold"
        End Select

        Dim dvw As DataView = dst.Tables("SOTORDR0").DefaultView
        dvw.RowFilter = sqlFilter

    End Sub

    Private Sub chkEditInternalNotes_CheckedChanged(sender As Object, e As EventArgs) Handles chkEditInternalNotes.CheckedChanged
        If optOrders.Value <> "O" Then
            If chkEditInternalNotes.Checked Then
                MsgBox("You may Edit Internal Comments for Open Orders only", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                chkEditInternalNotes.Checked = False
                Exit Sub
            End If
        End If

        If chkEditInternalNotes.Checked Then
            chkEditInternalNotes.Tag = "?"
            If Not ASCMAIN1.Logical_Lock("SOTORDR0", CUST_CODE) Then
                chkEditInternalNotes.Checked = False
                Exit Sub
            End If

            dst.Tables("SOTORDR0").AcceptChanges()
            grdSOTORDR0.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdSOTORDR0.DisplayLayout.Bands(0).Columns("ORDR_INTERNAL_NOTES").CellAppearance.BackColor = System.Drawing.Color.Yellow

            UltraExplorerBar1.Groups("Screen Control").Visible = False
            UltraExplorerBar1.Groups("Show Orders").Visible = False

            chkEditInternalNotes.Tag = ""
            UltraExplorerBar1.Groups("Screen Control").Items("Done").Visible = False
            'Set_ScreenMode_Base(True)
        Else
            If chkEditInternalNotes.Tag & "" = "?" Then
                chkEditInternalNotes.Tag = ""
                Exit Sub
            End If
            grdSOTORDR0.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grdSOTORDR0.DisplayLayout.Bands(0).Columns("ORDR_INTERNAL_NOTES").CellAppearance.BackColor = System.Drawing.Color.Empty

            UltraExplorerBar1.Groups("Screen Control").Visible = True
            UltraExplorerBar1.Groups("Show Orders").Visible = True

            Try

            Catch ex As Exception

            End Try

            BeginTrans()

            If grdSOTORDR0.ActiveRow IsNot Nothing AndAlso grdSOTORDR0.ActiveRow.DataChanged Then
                grdSOTORDR0.ActiveRow.Update()
            End If
            For Each rowSOTORDR0 As DataRow In dst.Tables("SOTORDR0").Select("", "", DataViewRowState.ModifiedCurrent)
                Dim ORDR_GROUP_NO As String = rowSOTORDR0.Item("ORDR_GROUP_NO")
                Dim ORDR_INTERNAL_NOTES As String = rowSOTORDR0.Item("ORDR_INTERNAL_NOTES") & ""
                'ASCMAIN1.sql = "Update SOTORDR0 Set ORDR_INTERNAL_NOTES = :PARM1 where ORDR_GROUP_NO = :PARM2"
                'ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {ORDR_INTERNAL_NOTES, ORDR_GROUP_NO})
                ASCMAIN1.sql = "Update SOTORDR1 Set ORDR_INTERNAL_NOTES = :PARM1 where ORDR_GROUP_NO = :PARM2"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {ORDR_INTERNAL_NOTES, ORDR_GROUP_NO})
            Next

            Update_Record_TDA("SOTORDR0")

            CommitTrans()

            UltraExplorerBar1.Groups("Screen Control").Items("Done").Visible = True
            'Set_ScreenMode_Base(False)
        End If
    End Sub

    Private Sub SOFCORD1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

    End Sub

    Private Sub SOFCORD1_Deactivate(sender As Object, e As EventArgs) Handles Me.Deactivate

    End Sub

#End Region
End Class


Public Class MyCustomTooltip
    Implements IRenderLabel

    Public Sub New()

    End Sub 'New

    Public Overloads Function ToString(ByVal Context As System.Collections.Hashtable) As String Implements IRenderLabel.ToString
        'Return Context("ITEM_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))
        'Return Context("SERIES_LABEL") & vbCrLf & Context("ITEM_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))
        Return Context("SERIES_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))

    End Function 'ToString 
End Class 'MyCustomTooltip
