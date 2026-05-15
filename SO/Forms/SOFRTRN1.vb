Imports System.Drawing
Imports System.Math

Public Class SOFRTRN1
    Private rowSOTRTRN1 As DataRow
    Private location_support As Boolean = False
    Private rowICTWHSE1 As DataRow
    Private whse_is_a_3PL As Boolean

    Private RECORD_INDEXs As List(Of Int32)
    Private INV_NO_RETURNED As String
    Private KEY_3PL_RECORD As String
    Private PRD_END_DATE As Date

    ' Used for Returns Authorization.
    Private RA_NO As String = String.Empty
    Private RTRN_NO As String = String.Empty
    Private rowARTCUST2 As DataRow = Nothing

    Private processing3PL As Boolean = False
    Private EDI_DOC_SEQ_NO As String = String.Empty
    Private SOT3PLF1 As String = String.Empty
    Private rowEDTRTRN1 As DataRow = Nothing
    Private tblEDTRTRN1 As DataTable = Nothing
    Private pl3Cust_store_no As String = String.Empty

    Private viewSOT3PLF1 As DataView

    Private EDI_DOC_SEQ_NO_List As New List(Of String)
    Private IC_PARM_REASON_ADJ As String = String.Empty

    Private validDates() As Date = TAC.SOCMAIN1.Validate_Invoice_Date(Nothing, 0, 1, Nothing)

    Private IC_PARM_WHSE_CODE_RFB As String = String.Empty
    Private IC_PARM_WHSE_CODE_DST As String = String.Empty
    Private IC_PARM_WHSE_CODE_DISC As String = String.Empty
    Private IC_PARM_WHSE_CODE_RTN As String = String.Empty
    Private IC_PARM_WHSE_CODE As String = String.Empty

    Private rowARTCUST1_BT As DataRow = Nothing
    Private PRICE_CLASS_CODE As String = String.Empty
    Private PRICE_BASE_DPCT As Decimal = 0
    Private PRICE_BASIS As String = String.Empty
    Private PRICE_LIST_CODE As String = String.Empty
    Private CUST_BILL_TO_CUST As String = String.Empty
    Private PRICE_LIST_CODE_ALLO As String = String.Empty
    Private CUST_CODE_ALLO As String = String.Empty
    Private ITEM_RETAIL_PRICE As Decimal = 0

    Private sotrtrn2RtrnNo As String = String.Empty

    ' 08/10/2020 - Initially done for IPLB for Kate Spade items.
    Private tblSpecialReturns As DataTable
    Private Const INTSpecialReturnsAdjustmentReasonCode As String = "KSP"

    Private rowWHTTPLP1 As DataRow = Nothing

#Region "ABS Standard Routines"

    Private Sub SOFRTRN1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        'If ScannerInUse AndAlso Not txtItemCode.Focused Then
        '    txtItemCode.Focus()
        'End If
    End Sub

    ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "SOFRTRNI" Then
            InquiryMode = True
        End If

        Get_PARM("ICTPARM1")
        Get_PARM("WHTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("SOTPARM1")

        ' 08/10/2020 - Initially done for IPLB for Kate Spade items.
        Select Case ASCMAIN1.CLIENT
            Case "INT"
                ASCMAIN1.sql = "SELECT ITEM_CODE, ITEM_DESC, COLLECTION_CODE, ITEM_BASIC_PROMO, ITEM_SNU_CODE, ITEM_CLASS_CODE, ITEM_UPC_CODE, ITEM_EAN_CODE
                                     FROM ICTITEM1 
                                     WHERE COLLECTION_CODE IN (SELECT COLLECTION_CODE FROM ICTCOLL1 WHERE BRAND_CODE  in ('KSP','LCS')) 
                                     AND ITEM_DESC LIKE 'XXX %'"
            Case Else
                ASCMAIN1.sql = "SELECT ITEM_CODE, ITEM_DESC, COLLECTION_CODE, ITEM_BASIC_PROMO, ITEM_SNU_CODE, ITEM_CLASS_CODE, ITEM_UPC_CODE, ITEM_EAN_CODE
                                     FROM ICTITEM1 
                                     WHERE ROWNUM < 1"
        End Select

        ' 08/10/2020 - Initially done for IPLB for Kate Spade items.
        tblSpecialReturns = ASCDATA1.GetDataTable(ASCMAIN1.sql)
        tblSpecialReturns.PrimaryKey = New DataColumn() {tblSpecialReturns.Columns("ITEM_CODE")}

        IC_PARM_REASON_ADJ = ROWs("ICTPARM1").Item("IC_PARM_REASON_ADJ") & String.Empty
        IC_PARM_REASON_ADJ = IC_PARM_REASON_ADJ.Trim
        If IC_PARM_REASON_ADJ.Length > 0 Then
            If LookUp("ICTREAS1", IC_PARM_REASON_ADJ) Is Nothing Then
                IC_PARM_REASON_ADJ = String.Empty
            End If
        End If

        ' Default Values
        IC_PARM_WHSE_CODE_RFB = ROWs("ICTPARM1").Item("IC_PARM_WHSE_CODE_RFB") & String.Empty
        IC_PARM_WHSE_CODE_DST = ROWs("ICTPARM1").Item("IC_PARM_WHSE_CODE_DST") & String.Empty
        IC_PARM_WHSE_CODE_DISC = ROWs("ICTPARM1").Item("IC_PARM_WHSE_CODE_DISC") & String.Empty
        IC_PARM_WHSE_CODE_RTN = ROWs("ICTPARM1").Item("IC_PARM_WHSE_CODE_RTN") & String.Empty
        IC_PARM_WHSE_CODE = ROWs("ICTPARM1").Item("IC_PARM_WHSE_CODE") & String.Empty

        With dst
            ASCMAIN1.sql = "Select SOTRTRN1.*" _
            & " from SOTRTRN1 where SOTRTRN1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "SOTRTRNX", "**", 0, False, "V")

            ASCMAIN1.sql = "Select SOTRTRN1.*, SOTRTRN2.RTRN_LNO, SOTRTRN2.ITEM_CODE, SOTRTRN2.RTRN_QTY, SOTRTRN2.RTRN_QTY_1, SOTRTRN2.RTRN_QTY_2, SOTRTRN2.RTRN_QTY_3, 
                            SOTRTRN2.RTRN_PRICE, SOTRTRN2.ITEM_COST_STD, SOTRTRN2.COST_CATGY_CODE, SOTRTRN2.PROD_CODE" _
            & " from SOTRTRN1, SOTRTRN2 where SOTRTRN1.RTRN_NO = SOTRTRN2.RTRN_NO AND SOTRTRN1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "SOTRTRNX_D", "**", 0, False, "V")
            .Tables("SOTRTRNX_D").Columns.Add("LINE_SALES", GetType(System.Decimal), "ISNULL(RTRN_QTY,0) * ISNULL(RTRN_PRICE,0)")
            .Tables("SOTRTRNX_D").Columns.Add("LINE_COSTS", GetType(System.Decimal), "ISNULL(RTRN_QTY,0) * ISNULL(ITEM_COST_STD,0)")

            ASCMAIN1.sql = "Select SOTRTRN2.RTRN_NO, SUM(SOTRTRN2.RTRN_QTY) RTRN_QTY_TOTAL" _
                & " from SOTRTRN1, SOTRTRN2 where SOTRTRN1.RTRN_NO = SOTRTRN2.RTRN_NO AND SOTRTRN1.OPS_YYYYPP = :PARM1" _
                & " GROUP BY SOTRTRN2.RTRN_NO"
            Create_TDA(.Tables.Add, "SOTRTRNX2", "**", 0, False, "V")

            ASCMAIN1.sql = "Select SOTRTRN3.*, GLTACCT1.ACCT_DESC" _
            & ", SOTRTRN1.RTRN_DATE, SOTRTRN1.WHSE_CODE, SOTRTRN1.REASON_CODE" _
            & ", SOTRTRN1.RTRN_NOTE, SOTRTRN1.INIT_OPER, SOTRTRN1.INIT_DATE" _
            & ", SOTRTRN1.RTRN_SOURCE, SOTRTRN1.OPS_YYYYPP" _
            & " from SOTRTRN1,SOTRTRN3,GLTACCT1 where SOTRTRN1.OPS_YYYYPP = :PARM1" _
            & " and GLTACCT1.ACCT_CODE = SOTRTRN3.ACCT_CODE" _
            & " and SOTRTRN3.RTRN_NO = SOTRTRN1.RTRN_NO"
            Create_TDA(.Tables.Add, "SOTRTRNG", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "SOTRTRN1", "*")
            Create_TDA(.Tables.Add("SOTRTRNB"), "SOTRTRN1", "*")
            .Tables("SOTRTRN1").Columns.Add("EDI_DOC_SEQ_NO", GetType(System.String))
            .Tables("SOTRTRNB").Columns.Add("EDI_DOC_SEQ_NO", GetType(System.String))

            ASCMAIN1.sql = "Select SOTRTRN2.*, ICTITEM1.ITEM_DESC" _
            & " from SOTRTRN2, ICTITEM1 where SOTRTRN2.ITEM_CODE = ICTITEM1.ITEM_CODE"
            Create_TDA(.Tables.Add, "SOTRTRN2", "**", 1)
            If Not .Tables("SOTRTRN2").Columns.Contains("RTRN_QTY_4") Then
                .Tables("SOTRTRN2").Columns.Add("RTRN_QTY_4", GetType(System.Int32))
                .Tables("SOTRTRN2").Columns.Add("RTRN_AS_PO_REC", GetType(System.String))
            End If

            .Tables("SOTRTRN2").Columns.Add("RECORD_INDEX", GetType(System.Int32))
            .Tables("SOTRTRN2").Columns.Add("LINE_SALES", GetType(System.Decimal), "ISNULL(RTRN_QTY,0) * ISNULL(RTRN_PRICE,0)")
            .Tables("SOTRTRN2").Columns.Add("LINE_COSTS", GetType(System.Decimal), "ISNULL(RTRN_QTY,0) * ISNULL(ITEM_COST_STD,0)")
            .Tables("SOTRTRN2").Columns.Add("RTRN_QTY_TOTAL", GetType(System.Decimal), "ISNULL(RTRN_QTY_1,0) + ISNULL(RTRN_QTY_2,0) + ISNULL(RTRN_QTY_3,0) + ISNULL(RTRN_QTY_4,0)")
            .Tables("SOTRTRN2").Columns.Add("RTRN_QTY_ORIG", GetType(System.Int32))
            .Tables("SOTRTRN2").Columns.Add("EDI_DOC_SEQ_NO", GetType(System.String))

            ' ISSUE-7369 Return To Stock needs to look at ICTWHSE2
            .Tables("SOTRTRN2").Columns.Add("WHSE_CODE_STOCK", GetType(System.String))

            ASCMAIN1.sql = "Select SOTRTRN3.*, GLTACCT1.ACCT_DESC" _
            & " from SOTRTRN3,GLTACCT1 where GLTACCT1.ACCT_CODE = SOTRTRN3.ACCT_CODE"
            Create_TDA(.Tables.Add, "SOTRTRN3", "**", 1)

            ASCMAIN1.sql = "Select ICTSTAT2.*" _
            & " from ICTSTAT2 where ITEM_CODE = :PARM1 and WHSE_CODE = :PARM2"
            Create_TDA(.Tables.Add, "ICTSTAT2", "**", 0, False, "VV")

            .Tables.Add("SOTRTRN0")
            .Tables("SOTRTRN0").Columns.Add("KEY")
            .Tables("SOTRTRN0").Columns.Add("DESCRIPTION")

            Create_TDA(.Tables.Add, "ARTOPEN1", "*")
            Create_TDA(.Tables.Add, "SOTINVH1", "*")
            Create_TDA(.Tables.Add, "SOTINVH2", "*")
            Create_TDA(.Tables.Add, "SOTORDR0", "*")
            Create_TDA(.Tables.Add, "SOTORDR1", "*")
            Create_TDA(.Tables.Add, "SOTORDR2", "*")
            Create_TDA(.Tables.Add, "SOTORDR5", "*")

            ASCMAIN1.sql = "Select SOTINVH1.INV_TYPE, SOTINVH1.INV_NO" & vbCrLf _
                & ", SOTINVH1.CUST_CODE, SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_CUST_PO" & vbCrLf _
                & ", SOTINVH1.ORDR_NO, SOTINVH1.WHSE_CODE" & vbCrLf _
                & ", SOTINVH1.INV_SALES, SOTINVH1.INV_FREIGHT, SOTINVH1.INV_TOTAL_AMOUNT" & vbCrLf _
                & ", SOTINVH1.INV_DATE, SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
                & ", SOTINVH1.ORDR_DEPT, SOTINVH1.ORDR_TYPE_CODE, SOTINVH1.CUST_BILL_TO_CUST" & vbCrLf _
                & " from SOTINVH1" & vbCrLf _
                & " where SOTINVH1.INV_TYPE = 'I' and SOTINVH1.CUST_CODE = :PARM1 and SOTINVH1.ORDR_YYYYPP_UPDATED >= :PARM2"
            ' CHANGING QUERY BELOW TO BRING IN CREDITS 
            ASCMAIN1.sql = "Select SOTINVH1.INV_TYPE, SOTINVH1.INV_NO" & vbCrLf _
                & ", SOTINVH1.CUST_CODE, SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_CUST_PO" & vbCrLf _
                & ", SOTINVH1.ORDR_NO, SOTINVH1.WHSE_CODE" & vbCrLf _
                & ", SOTINVH1.INV_SALES, SOTINVH1.INV_FREIGHT, SOTINVH1.INV_TOTAL_AMOUNT" & vbCrLf _
                & ", SOTINVH1.INV_DATE, SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
                & ", SOTINVH1.ORDR_DEPT, SOTINVH1.ORDR_TYPE_CODE, SOTINVH1.CUST_BILL_TO_CUST" & vbCrLf _
                & " from SOTINVH1" & vbCrLf _
                & " where (SOTINVH1.INV_TYPE = 'I' OR (SOTINVH1.INV_TYPE = 'C' AND SOTINVH1.ORDR_TYPE_CODE = 'RTN')) and SOTINVH1.CUST_CODE = :PARM1 and SOTINVH1.ORDR_YYYYPP_UPDATED >= :PARM2"
            Create_TDA(.Tables.Add, "SOTINVHH", "**", 0, False, "VV", 2)

            ASCMAIN1.sql = "Select SOTINVH2.INV_TYPE, SOTINVH2.INV_NO, SOTINVH2.INV_LNO" & vbCrLf _
                & ", SOTINVH2.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
                & ", SOTINVH2.ORDR_UNIT_PRICE, SOTINVH2.ORDR_QTY_SHIP" & vbCrLf _
                & ", SOTINVH1.INV_DATE, SOTINVH1.ORDR_CUST_PO" & vbCrLf _
                & " from SOTINVH2, ICTITEM1, SOTINVH1" & vbCrLf _
                & " where SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
                & "  and SOTINVH2.CUST_CODE = :PARM1" & vbCrLf _
                & "  and SOTINVH2.ORDR_YYYYPP_UPDATED > :PARM2" & vbCrLf _
                & "  and SOTINVH2.ITEM_CODE = :PARM3 " & vbCrLf _
                & "  and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
                & "  and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE and SOTINVH1.INV_NO = SOTINVH2.INV_NO"
            Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False, "VVV", 3)

            Create_TDA(.Tables.Add, "SOTRMAF1", "*", 1)

            ASCMAIN1.sql = "Select SOTRMAF1.* from SOTRMAF1 where RA_STATUS = 'O'"
            Create_TDA(.Tables.Add, "SOTRMAFX", "**", 0, False, , 1)

            ASCMAIN1.sql = "Select SOTRMAF1.*, ARTCUST1.CUST_NAME from SOTRMAF1, ARTCUST1 where ARTCUST1.CUST_CODE = SOTRMAF1.CUST_CODE AND  SOTRMAF1.RA_STATUS = 'O'"
            Create_TDA(.Tables.Add, "SOT3PLF1", "**", 0, False, , 1)

            ASCMAIN1.sql = "Select * from SOTRMAF2"
            Create_TDA(.Tables.Add, "SOT3PLF2", "**", 0, False, , 2)
            .Relations.Add("SOT3PLF1_SOT3PLF2", dst.Tables("SOT3PLF1").Columns("RA_NO"), dst.Tables("SOT3PLF2").Columns("RA_NO"))

            ASCMAIN1.sql = "Select SOTRMAF2.*, ICTITEM1.ITEM_DESC" _
                & " from SOTRMAF2,ICTITEM1" _
                & " where ICTITEM1.ITEM_CODE (+) = SOTRMAF2.ITEM_CODE"
            Create_TDA(.Tables.Add, "SOTRMAF2", "**", 1)

            Create_TDA(.Tables.Add, "ICTIXFR1", "*", 1)
            Create_TDA(.Tables.Add, "ICTIXFR2", "*", 2)

            ASCMAIN1.sql = "Select EDTRTRN1.*, ARTCUST1.CUST_NAME, SOTRMAF1.CUST_CODE RA_CUST_CODE" _
                    & " from EDTRTRN1, ARTCUST1, SOTRMAF1" _
                    & " where ARTCUST1.CUST_CODE (+) = EDTRTRN1.CUST_CODE" _
                    & " and SOTRMAF1.RA_NO (+) = EDTRTRN1.EDI_RA_NO" _
                    & " and NVL(PROCESS_IND, '0') = '0'"
            Create_TDA(.Tables.Add, "EDTRTRN1", "**", 0, False)

            ASCMAIN1.sql = "Select EDTRTRN2.*, ICTITEM1.ITEM_DESC" _
                & " from EDTRTRN1, EDTRTRN2, ICTITEM1" _
                & " where EDTRTRN1.EDI_DOC_SEQ_NO = EDTRTRN2.EDI_DOC_SEQ_NO" _
                & " and ICTITEM1.ITEM_CODE (+) = EDTRTRN2.EDI_ITEM_CODE" _
                & " and NVL(EDTRTRN1.PROCESS_IND, '0') = '0'"
            Create_TDA(.Tables.Add, "EDTRTRN2", "**", 0, False)

            .Relations.Add("EDTRTRN1_EDTRTRN2", dst.Tables("EDTRTRN1").Columns("EDI_DOC_SEQ_NO"), dst.Tables("EDTRTRN2").Columns("EDI_DOC_SEQ_NO"))

            SOT3PLF1 = ASCMAIN1.Temp_Table("select ra_no from sotrmaf1 where rownum < 1")

            Create_TDA(.Tables.Add, "ICTIADJ1", "*")
            Create_TDA(.Tables.Add, "ICTIADJ2", "*")

            ' 07/28/2020
            ' INT Kate Spade changes
            ' *************************************************************
            Create_TDA(.Tables.Add, "ICTIREC1", "*")
            Create_TDA(.Tables.Add, "ICTIREC2", "*")
            Create_TDA(.Tables.Add, "ICTIREC4", "*")
            ' SR-6549 - Lot Numbers on Shipments and Receipts
            Create_TDA(.Tables.Add, "ICTIRECL", "*")

            Create_TDA(.Tables.Add, "ICTPINV1", "*")
            Create_TDA(.Tables.Add, "ICTPINV2", "*")
            With .Tables("ICTPINV2").Columns
                .Add("QTY_REC", GetType(System.Int64))
                .Add("QTY_INV", GetType(System.Int64))
                .Add("AMT_INV", GetType(System.Decimal), "ISNULL(PINV_COST,0) * ISNULL(QTY_INV,0)")
                .Add("QTY_REC_NOT_INV", GetType(System.Int64))
            End With

            Create_TDA(.Tables.Add, "APTINVH1", "*")
            Create_TDA(.Tables.Add, "APTINVH2", "*")

            ASCMAIN1.sql = "Select APTINVH5.*" & vbCrLf _
                & ", ICTITEM1.ITEM_DESC, ICTITEM1.COLLECTION_CODE, ICTCOST1.ACCT_CODE_PPV" & vbCrLf _
                & ", ICTIREC2.QTY_REC, ICTIREC2.QTY_INV, ICTIREC2.PO_COST" & vbCrLf _
                & ", ICTIREC2.ITEM_CODE, ICTIREC2.ITEM_UOM" & vbCrLf _
                & " from ICTIREC2,APTINVH5,ICTITEM1,ICTCOST1" & vbCrLf _
                & " where APTINVH5.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
                & "   and APTINVH5.RECEIPT_LNO = ICTIREC2.RECEIPT_LNO" & vbCrLf _
                & "   and APTINVH5.RECEIPT_NO = :PARM1" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = ICTIREC2.ITEM_CODE" _
                & "   and ICTCOST1.COST_CATGY_CODE (+) = ICTITEM1.COST_CATGY_CODE"
            Create_TDA(.Tables.Add, "APTINVH5", "**", 0, True, "V", 2)

            With .Tables("APTINVH5").Columns
                .Add("AMT_REC", GetType(System.Decimal), "QTY_REC * PO_COST")
                .Add("AMT_INV", GetType(System.Decimal), "INV_QTY * INV_COST")
                .Add("AMT_REC_NOT_INV", GetType(System.Decimal), "QTY_REC_NOT_INV * PO_COST")
                .Add("AMT_REC_NOT_INV_OFFSET", GetType(System.Decimal), "IIF(CLOSE_LINE='1',QTY_REC_NOT_INV * PO_COST,INV_QTY * PO_COST)")
                .Add("QTY_VAR", GetType(System.Int64), "IIF(CLOSE_LINE='0',0,ISNULL(INV_QTY,0) - ISNULL(QTY_REC_NOT_INV,0))")
                .Add("AMT_VAR", GetType(System.Decimal), "AMT_INV - IIF(CLOSE_LINE='0', ISNULL(INV_QTY,0) * ISNULL(PO_COST,0), ISNULL(QTY_REC_NOT_INV,0) * ISNULL(PO_COST,0))")
            End With

            Create_TDA(.Tables.Add, "POTORDR1", "*")
            Create_TDA(.Tables.Add, "POTORDR2", "*")

            Create_TDA(.Tables.Add, "APTVEND5", "*")

            .Tables.Add("APTINVH5_VAR")
            With .Tables("APTINVH5_VAR")
                .Columns.Add("COST_CATGY_CODE")
                .Columns.Add("COLLECTION_CODE")
                .Columns.Add("ACCT_CODE_PPV")
                .Columns.Add("SEG2_CODE")
                .Columns.Add("SEG3_CODE")
                .Columns.Add("SEG4_CODE")
                .Columns.Add("AMT_REC", GetType(System.Decimal))
                .Columns.Add("AMT_INV", GetType(System.Decimal))
                .Columns.Add("AMT_REC_NOT_INV", GetType(System.Decimal))
                .Columns.Add("AMT_REC_NOT_INV_OFFSET", GetType(System.Decimal))
                .Columns.Add("AMT_VAR", GetType(System.Decimal))
                .Columns.Add("AMT_VAR_CB", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() { .Columns("COST_CATGY_CODE"), .Columns("COLLECTION_CODE")}
            End With

            Get_PARM("APTPARM1")
            Get_PARM("GLTPARM1")
            ' *************************************************************

            If ASCMAIN1.CLIENT = "INT" Then
                ASCMAIN1.sql = " SELECT 'A' CONTROL_CODE, CUST_CODE, CUST_STORE_NO, CUST_NO_3PL, CUST_STORE_NO_3PL" _
                    & " FROM ARTCUST2" _
                    & " WHERE CUST_NO_3PL IS NOT NULL AND CUST_STORE_NO_3PL IS NOT NULL AND CUST_CODE = :PARM1" _
                    & " UNION" _
                    & " SELECT 'X' CONTROL_CODE, CUST_CODE, CUST_STORE_NO, TO_CHAR(CSCUS1) CUST_NO_3PL,  TO_CHAR(CSCUS2) CUST_STORE_NO_3PL" _
                    & " FROM TATXREFX" _
                    & " WHERE CSCUS1 IS NOT NULL AND CSCUS2 IS NOT NULL AND CUST_CODE = :PARM1"
                Create_TDA(.Tables.Add, "ARTCUST3PL", ASCMAIN1.sql, 0, False, "V", 5)
            End If

            Create_TDA(.Tables.Add, "WHTTPLP1", "*")
            Fill_Records("WHTTPLP1", "", True, "SELECT * FROM WHTTPLP1")

            ASCMAIN1.sql = "SELECT X.*, ICTITEM1.ITEM_DESC FROM ICTITEM1, (
                                SELECT SOTRTRN2.ITEM_CODE, SOTRTRN2.RTRN_PRICE, SUM (RTRN_QTY) QTY, SUM (RTRN_QTY * RTRN_PRICE) AMT FROM SOTRTRN1,SOTRTRN2
                                WHERE SOTRTRN1.CUST_CODE = :PARM1 and SOTRTRN1.CUST_CLAIM_NO = :PARM2
                                AND SOTRTRN1.RTRN_NO = SOTRTRN2.RTRN_NO
                                GROUP BY SOTRTRN2.ITEM_CODE, SOTRTRN2.RTRN_PRICE) X
                                WHERE X.ITEM_CODE = ICTITEM1.ITEM_CODE"
            Create_TDA(.Tables.Add, "SOTRTRNX_I", ASCMAIN1.sql, 0, False, "VV", 0)

            ' ISSUE-7369 Return To Stock needs to look at ICTWHSE2
            ASCMAIN1.sql = "SELECT *
                            FROM ICTWHSE2 T2
                            WHERE EXISTS (
                                SELECT 1
                                FROM ICTWHSE1 W1
                                JOIN ICTWHSE1 W2 
                                    ON W1.LP_CODE = W2.LP_CODE
                                WHERE W2.WHSE_CODE = :PARM1
                                    AND W1.WHSE_CODE = T2.WHSE_CODE)"
            Create_TDA(.Tables.Add, "ICTWHSE2", ASCMAIN1.sql, 0, False, "V")
        End With

        ' 2020-07-29
        Create_Relation("SOTRTRNX", "SOTRTRNX2", "RTRN_NO")
        dst.Tables("SOTRTRNX").Columns.Add("RTRN_QTY_TOTAL", GetType(System.Int32), "SUM(CHILD.RTRN_QTY_TOTAL)")
        'grdSOTRTRNX.DisplayLayout.Bands(1).Hidden = True

        Show_Filter(grdSOTINVHH, True)

        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)

        cbeInvoiceHistory.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeInvoiceHistory.SelectedItem = cbeInvoiceHistory.Items(3)

        grdSOTRTRN0.DataSource = dst.Tables("SOTRTRN0")
        grdSOTRTRN2.DataSource = dst.Tables("SOTRTRN2")
        grdSOTRTRN3.DataSource = dst.Tables("SOTRTRN3")
        grdSOTRTRNX.DataSource = dst.Tables("SOTRTRNX")
        grdSOTRTRNX_D.DataSource = dst.Tables("SOTRTRNX_D")
        grdSOTRTRNX_I.DataSource = dst.Tables("SOTRTRNX_I")
        grdSOTRTRNG.DataSource = dst.Tables("SOTRTRNG")
        grdEDTRTRN1.DataSource = dst.Tables("EDTRTRN1")

        viewSOT3PLF1 = New DataView(dst.Tables("SOT3PLF1"))
        grdSOT3PLF1.DataSource = viewSOT3PLF1

        grdSOTINVHH.DataSource = dst.Tables("SOTINVHH")
        grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")

        Create_Summary(grdSOTRTRNX, "RTRN_NO", "Count")
        Create_Summary(grdSOTRTRNX, New String() {"RTRN_COSTS", "RTRN_SALES", "RTRN_AMOUNT"})

        Create_Summary(grdSOTRTRNX_D, "RTRN_NO", "Count")
        Create_Summary(grdSOTRTRNX_D, New String() {"RTRN_COSTS", "RTRN_SALES", "RTRN_AMOUNT", "RTRN_QTY", "RTRN_QTY_1", "RTRN_QTY_2", "RTRN_QTY_3", "LINE_SALES", "LINE_COSTS"})


        Create_Summary(grdSOTRTRNX_I, "ITEM_CODE", "Count")
        Create_Summary(grdSOTRTRNX_I, New String() {"QTY", "AMT"})


        Create_Summary(grdSOTRTRNG, "RTRN_NO", "Count")
        Create_Summary(grdSOTRTRNG, "DIST_AMT")

        Create_Summary(grdSOTRTRN2, "RTRN_LNO", "Count")
        Create_Summary(grdSOTRTRN2, New String() {"RTRN_QTY", "RTRN_QTY_1", "RTRN_QTY_2", "RTRN_QTY_3", "RTRN_QTY_4", "LINE_SALES", "LINE_COSTS"})

        Create_Summary(grdSOTRTRN3, "RTRN_GNO", "Count")
        Create_Summary(grdSOTRTRN3, "DIST_AMT")

        Create_Summary(grdEDTRTRN1, "EDI_RA_NO", "Count")
        Create_Summary(grdEDTRTRN1, "EDI_QTY_RETURNED", "Sum", "EDTRTRN1_EDTRTRN2")
        Create_Summary(grdEDTRTRN1, "EDI_QTY_BACK_TO_STOCK", "Sum", "EDTRTRN1_EDTRTRN2")
        Create_Summary(grdEDTRTRN1, "EDI_QTY_IN_REPAIR", "Sum", "EDTRTRN1_EDTRTRN2")
        Create_Summary(grdEDTRTRN1, "EDI_QTY_DAMAGED", "Sum", "EDTRTRN1_EDTRTRN2")

        With grdSOTRTRNX.DisplayLayout.Bands("SOTRTRNX")
            .Columns("RTRN_NO").Header.Fixed = True
        End With

        With grdSOTRTRNX_D.DisplayLayout.Bands(0)
            .Columns("RTRN_NO").Header.Fixed = True
            .Columns("RTRN_LNO").Header.Fixed = True
        End With

        With grdSOTRTRNG.DisplayLayout.Bands("SOTRTRNG")
            .Columns("RTRN_NO").Header.Fixed = True
        End With

        'ASCMAIN1.Add_Value_List(grdSOTRTRNX, "REASON_CODE", Nothing, New String() {":", "D:Damaged", "X:Destroyed", "O:Overstock", "Z:Other"})
        'ASCMAIN1.Add_Value_List(grdSOTRTRNG, "REASON_CODE", Nothing, New String() {":", "D:Damaged", "X:Destroyed", "O:Overstock", "Z:Other"})

        'ASCMAIN1.Add_Value_List(grdSOTRTRNX, "REASON_CODE", "Select REASON_CODE, REASON_DESC from ARTREAS1 order by REASON_DESC")
        'ASCMAIN1.Add_Value_List(grdSOTRTRNG, "REASON_CODE", "Select REASON_CODE, REASON_DESC from ARTREAS1 order by REASON_DESC")

        grdSOTRTRN0.DisplayLayout.Bands(0).ColHeadersVisible = False
        Set_SEGS(grdSOTRTRN3, "SOTRTRN3")

        grdSOTRMAFX.DataSource = dst.Tables("SOTRMAFX")
        grdSOTRMAFX.DisplayLayout.UseFixedHeaders = True
        With grdSOTRMAFX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"RA_NO", "CUST_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        Bind_Controls(grpTotals, "SOTRTRN1")
        'Set_Read_Only(grpTotals, True)
        If InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") = 0 Then
            'grpTotals.Visible = False
            lblRTRN_COSTS.Visible = False
            numRTRN_COSTS.Visible = False
            With grdSOTRTRN2.DisplayLayout.Bands(0)
                .Columns("ITEM_COST_STD").Hidden = True
                .Columns("LINE_COSTS").Hidden = True
                .Columns("COST_CATGY_CODE").Hidden = True
                .Columns("PROD_CODE").Hidden = True
            End With
        End If

        With grdSOTRTRN2.DisplayLayout.Bands(0)
            .Columns("RTRN_QTY_1").Header.Caption = "Stock"
            .Columns("RTRN_QTY_2").Header.Caption = "Refurb"
            .Columns("RTRN_QTY_3").Header.Caption = "Destroy"
            .Columns("RTRN_QTY_4").Header.Caption = "Discount"

            If Not ASCMAIN1.CLIENT = "AHA" Then
                .Columns("RTRN_QTY_4").Hidden = True
            End If

            ' Keep option set in sysn with captions
            For Each optLoc As Infragistics.Win.ValueListItem In optLocation.Items
                optLoc.DisplayText = .Columns(optLoc.DataValue).Header.Caption
            Next

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"RTRN_LNO", "ITEM_CODE", "ITEM_DESC", "COST_CATGY_CODE", "PROD_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                End If
                If New String() {"RTRN_QTY", "RTRN_QTY_1", "RTRN_QTY_2", "RTRN_QTY_3", "RTRN_QTY_4"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                    If gcol.Key = "RTRN_QTY" Then
                        gcol.CellAppearance.BackColor = Color.LightBlue
                    End If
                    gcol.Width = 70
                End If
                If New String() {"RTRN_PRICE", "ITEM_COST_STD", "LINE_SALES", "LINE_COSTS"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                    If gcol.Key.StartsWith("LINE") Then
                        gcol.CellAppearance.BackColor = Color.LightGreen
                        gcol.Width = 90
                    Else
                        gcol.Width = 70
                    End If
                End If
            Next
        End With

        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", ASCMAIN1.CYP)
        PRD_END_DATE = rowGLTPARM2.Item("PRD_END_DATE")

        grpHeader.Visible = False
        Set_SEGS(grdSOTRTRNG, "SOTRTRNG")
        Bind_Controls(grpHeader, "SOTRTRN1")

        'If InquiryMode Then
        '    MyBase.Absx1.dteFor("RTRN_DATE").MinDate = CDate("01/01/2013")
        '    MyBase.Absx1.dteFor("RTRN_DATE").MaxDate = PRD_END_DATE
        'Else
        '    MyBase.Absx1.dteFor("RTRN_DATE").MinDate = validDates(0)
        '    MyBase.Absx1.dteFor("RTRN_DATE").MaxDate = PRD_END_DATE
        'End If

        ' Update EDTRTRN1 with the customer Code
        Try
            ASCMAIN1.sql = "Update EDTRTRN1 SET CUST_CODE = (SELECT CUST_CODE FROM SOTRMAF1 WHERE RA_NO = EDTRTRN1.EDI_RA_NO) WHERE CUST_CODE IS NULL AND NVL(PROCESS_IND, '0') = '0'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        Catch ex As Exception

        End Try

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"

                RA_NO = String.Empty
                RTRN_NO = String.Empty

                Absx1.txtFor("RA_NO").Text = Absx1.txtFor("RA_NO").Text.Trim

                ' Returns Authorizartion Trumps Cust_Code / Whse_Code
                If Absx1.txtFor("RA_NO").TextLength > 0 Then

                    Dim rowSOTRMAF1 As DataRow = Nothing

                    If IsNumeric(Absx1.txtFor("RA_NO").Text) Then
                        Absx1.txtFor("RA_NO").Text = ASCMAIN1.Format_Field(Absx1.txtFor("RA_NO").Text, "RA_NO")
                    End If

                    rowSOTRMAF1 = LookUp("SOTRMAF1", Absx1.txtFor("RA_NO").Text)
                    If rowSOTRMAF1 Is Nothing AndAlso ASCMAIN1.CLIENT = "INT" Then
                        RA_NO = Absx1.txtFor("RA_NO").Text
                    ElseIf rowSOTRMAF1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Return Authorization no: " & Absx1.txtFor("RA_NO").Text
                    ElseIf rowSOTRMAF1.Item("RA_STATUS") & String.Empty <> "O" Then
                        EMsg &= vbCr & "Return Authorization no: " & Absx1.txtFor("RA_NO").Text & " is not Open"
                    Else
                        RA_NO = Absx1.txtFor("RA_NO").Text
                    End If

                    If EMsg = "" Then
                        If Not ASCMAIN1.Logical_Lock("SOTRMAF1", RA_NO) Then Exit Sub
                        If rowSOTRMAF1 Is Nothing Then
                            If Not ASCMAIN1.Logical_Lock("SOTRMAF1", "CUST_CODE_" & MyBase.Absx1.txtFor("CUST_CODE").Text) Then Exit Sub
                        Else
                            If Not ASCMAIN1.Logical_Lock("SOTRMAF1", "CUST_CODE_" & rowSOTRMAF1.Item("CUST_CODE")) Then Exit Sub
                        End If
                    Else
                        Exit Select
                    End If

                    If rowSOTRMAF1 IsNot Nothing Then
                        Absx1.txtFor("CUST_CODE").Text = rowSOTRMAF1.Item("CUST_CODE") & String.Empty
                        Absx1.txtFor("WHSE_CODE").Text = rowSOTRMAF1.Item("WHSE_CODE") & String.Empty
                    End If

                    If Not IsDate(MyBase.Absx1.dteFor("RTRN_DATE").Value) Then
                        MyBase.Absx1.dteFor("RTRN_DATE").Value = DateTime.Now
                    End If

                Else
                    Validate_Code("CUST_CODE")
                    Validate_Code("WHSE_CODE")

                    If Absx1.dteFor("RTRN_DATE").Value & "" = "" Then
                        EMsg &= vbCr & "Invalid Date Specified for Entry"
                    End If

                    If Absx1.txtFor("WHSE_CODE").Text.Length = 0 Then
                        EMsg &= vbCr & "You must supply a Valid Warehouse"
                    Else
                        rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                        If IsNothing(rowICTWHSE1) Then
                            EMsg &= vbCr & "Warehouse Entered Is Not Valid"
                        Else
                            If rowICTWHSE1.Item("WHSE_STATUS").ToString <> "A" Then
                                EMsg &= vbCr & "Warehouse Entered Is Not Active"
                            Else
                                'disabling this message for now since we do want to put back the inventory into the ADS warhouse
                                ' re-enabling WJZ 02/01/2024
                                If rowICTWHSE1.Item("LP_CODE") & "" <> "" And KEY_3PL_RECORD = "" Then
                                    MsgBox("You are entering a Customer Credit involving a 3PL warehouse" _
                                           & vbCrLf & vbCrLf & "You must choose a reason code that does NOT impact inventory",
                                           MsgBoxStyle.OkOnly, "Verification")
                                    '    EMsg &= vbCr & "Warehouse Entered Is A 3PL.  No Adjustments Allowed"
                                End If
                            End If
                        End If
                    End If
                End If

                If EMsg.Length = 0 AndAlso processing3PL Then
                    ' See if the Edi Doc No is stale.
                    Dim sql As String = "Select * from EDTRTRN1 where EDI_DOC_SEQ_NO in ('" & Join(EDI_DOC_SEQ_NO_List.ToArray, "', '") & "')"
                    sql &= " and NVL(PROCESS_IND, '0') = '0'"
                    tblEDTRTRN1 = ASCDATA1.GetDataTable(sql)

                    If tblEDTRTRN1 Is Nothing OrElse tblEDTRTRN1.Rows.Count = 0 Then
                        EMsg &= vbCr & "Cannot locate the selected 3PL Entry or it has already been processed."
                    Else
                        EDI_DOC_SEQ_NO_List.Clear()
                        For Each row As DataRow In tblEDTRTRN1.Select("")
                            EDI_DOC_SEQ_NO_List.Add(row.Item("EDI_DOC_SEQ_NO"))
                            If Not ASCMAIN1.Logical_Lock("SOTRMAF1", "E" & row.Item("EDI_DOC_SEQ_NO"), , , True, 9) Then
                                EDI_DOC_SEQ_NO_List.Clear()
                                Exit Sub
                            End If
                        Next
                        rowEDTRTRN1 = tblEDTRTRN1.Rows(0)
                        EDI_DOC_SEQ_NO = rowEDTRTRN1.Item("EDI_DOC_SEQ_NO")
                    End If

                    'rowEDTRTRN1 = ASCDATA1.GetDataRow("Select * from EDTRTRN1 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
                    'If rowEDTRTRN1 Is Nothing Then
                    '    EMsg &= vbCr & "Cannot locate the selected 3PL Entry."
                    'ElseIf Val(rowEDTRTRN1.Item("PROCESS_IND") & String.Empty) <> 0 Then
                    '    EMsg &= vbCr & "The selected 3PL Entry is no longer available to be processed."
                    'ElseIf Not ASCMAIN1.Logical_Lock("SOTRMAF1", "E" & EDI_DOC_SEQ_NO) Then
                    '    Exit Sub
                    'End If
                End If

                If EMsg.Length > 0 Then
                    processing3PL = False
                    EDI_DOC_SEQ_NO = String.Empty
                    EDI_DOC_SEQ_NO_List.Clear()
                End If

            Case "View"
                If Absx1.txtFor("RTRN_NO").Text = "" Then
                    EMsg &= vbCr & "You must specify a Document No to View"
                Else
                    rowSOTRTRN1 = LookUp("SOTRTRN1", Absx1.txtFor("RTRN_NO").Text)
                    If rowSOTRTRN1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Document " & Absx1.txtFor("RTRN_NO").Text & " on File"
                    End If
                End If

            Case "Update"

                If Absx1.optFor("REASON_CODE").Value = "" Then
                    EMsg &= vbCr & "Reason is required"
                Else
                    For Each rowSOTRTRNB As DataRow In dst.Tables("SOTRTRNB").Select("")
                        rowSOTRTRNB.Item("REASON_CODE") = rowSOTRTRN1.Item("REASON_CODE")
                    Next
                End If

                Absx1.txtFor("CUST_STORE_NO").Text = Absx1.txtFor("CUST_STORE_NO").Text.Trim
                If Absx1.txtFor("CUST_STORE_NO").TextLength > 0 Then
                    rowARTCUST2 = LookUp("ARTCUST2", New String() {Absx1.txtFor("CUST_CODE").Text, Absx1.txtFor("CUST_STORE_NO").Text})
                    If rowARTCUST2 Is Nothing Then
                        EMsg &= vbCr & "Invalid Value Specified for Customer Store"
                    Else
                        rowSOTRTRN1.Item("CUST_STORE_NAME") = rowARTCUST2.Item("CUST_STORE_NAME") & String.Empty
                    End If
                Else
                    rowSOTRTRN1.Item("CUST_STORE_NAME") = ""
                End If

                If grdSOTRTRN2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Details Entered"
                Else
                    For Each rowSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Select("", "", DataViewRowState.CurrentRows)
                        If rowSOTRTRN2.Item("COST_CATGY_CODE") & "" = "" Then
                            EMsg &= vbCr & "Unable to determine Cost Catgy Code for " & rowSOTRTRN2.Item("ITEM_CODE") & ""
                        End If
                        If rowSOTRTRN2.Item("PROD_CODE") & "" = "" Then
                            EMsg &= vbCr & "Unable to determine Product Code for " & rowSOTRTRN2.Item("ITEM_CODE") & ""
                        End If

                        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", rowSOTRTRN2.Item("ITEM_CODE"))
                        If rowICTITEM1.Item("ITEM_COST_STATUS") & "" <> "" Then
                            EMsg &= vbCr & "Item " & rowSOTRTRN2.Item("ITEM_CODE") & " does not have a Standard Cost"
                        End If
                    Next
                End If

                If dst.Tables("SOTRTRN2").Select("RTRN_PRICE IS NULL").Length <> 0 Then
                    EMsg &= vbCr & "Some lines do not have Price"
                End If

                If Val(Absx1.numFor("RTRN_AMOUNT").Value & "") < 0 Then
                    If ASCMAIN1.USER_ID = "wjz" And ASCMAIN1.Running_in_VS Then
                        Stop ' PERMIT REVERSAL OF A CREDIT FOR RETURNS
                    End If
                    EMsg &= vbCr & "Total Amount is not a Credit to the Customer"
                End If

                If Absx1.dteFor("RTRN_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "Return Date is Mandatory"
                ElseIf Format(Absx1.dteFor("RTRN_DATE").Value, "yyyyMMdd") < Format(validDates(0), "yyyyMMdd") Then
                    EMsg &= vbCr & "Return Date may not be Earlier than " & Format(validDates(0), "MM/dd/yyyy")
                ElseIf Format(Absx1.dteFor("RTRN_DATE").Value, "yyyyMMdd") > Format(PRD_END_DATE, "yyyyMMdd") Then
                    EMsg &= vbCr & "Return Date may not be Later than " & Format(PRD_END_DATE, "MM/dd/yyyy")
                End If


                If dst.Tables("SOTRTRN2").Select("ISNULL(RTRN_QTY_TOTAL,0) > 0").Length = 0 Then
                    If ASCMAIN1.USER_ID = "wjz" And ASCMAIN1.Running_in_VS Then
                        Stop ' PERMIT REVERSAL OF A CREDIT FOR RETURNS
                    End If
                    EMsg &= vbCr & "No items have been returned"
                End If


                If EntryMode = "N" AndAlso EMsg.Length = 0 Then

                    Select Case ASCMAIN1.CLIENT
                        Case "INT"
                            If dst.Tables("SOTRTRN2").Select("ISNULL(RTRN_QTY_3, 0) > 0").Length > 0 Then
                                If IC_PARM_REASON_ADJ.Length = 0 Then
                                    EMsg &= vbCr & "You must have an Adjustments Default Reason Code in the Inventory Parameters to process Back To Stock Quantities."
                                End If
                            End If

                        Case "AHA"
                            'If dst.Tables("SOTRTRN2").Select("ISNULL(RTRN_QTY_2, 0) > 0").Length > 0 _
                            '    OrElse dst.Tables("SOTRTRN2").Select("ISNULL(RTRN_QTY_3, 0) > 0").Length > 0 Then
                            '    If IC_PARM_REASON_ADJ.Length = 0 Then
                            '        EMsg &= vbCr & "You must have an Adjustments Default Reason Code in the Inventory Parameters to process Ajdustments."
                            '    End If
                            'End If

                    End Select
                End If

                ' 08/10/2020 - Initially done for IPLB for Kate Spade items.
                If EntryMode = "N" AndAlso EMsg.Length = 0 Then
                    If dst.Tables("SOTRTRN2").Select("RTRN_AS_PO_REC = '1' and RTRN_QTY_1 <> 0").Length > 0 Then
                        EMsg &= vbCr & "All Returned Items marked for PO Receipts (Pink Background) must be placed in the 'Refurb' or 'Destroy' columns."
                    End If
                End If

                If EntryMode = "N" AndAlso EMsg.Length = 0 Then
                    If dst.Tables("SOTRTRN2").Select("ISNULL(RTRN_QTY,0) <> ISNULL(RTRN_QTY_TOTAL,0)").Length <> 0 Then
                        If MessageBox.Show("Some lines are out of Balance." & Environment.NewLine & "Do you want to continue?" _
                                            , "Out of Balance", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                            Exit Sub
                        End If

                        For Each rowSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Select("", "", DataViewRowState.CurrentRows)
                            rowSOTRTRN2.Item("RTRN_QTY") = Val(rowSOTRTRN2.Item("RTRN_QTY_TOTAL") & String.Empty)
                        Next
                        DisplayTotals()
                    End If
                End If

                ' Verify detail total matches Inv_sales
                Dim RTRN_SALES As Decimal = Val(dst.Tables("SOTRTRN2").Compute("SUM(LINE_SALES)", "") & String.Empty)
                Dim INV_SALES As Decimal = Val(dst.Tables("SOTRTRN1").Rows(0).Item("RTRN_SALES") & String.Empty) _
                                           + Val(dst.Tables("SOTRTRNB").Compute("SUM(RTRN_SALES)", "") & String.Empty)

                'If RA_NO <> "" AndAlso RTRN_SALES <> INV_SALES Then
                If RTRN_SALES <> INV_SALES Then
                    EMsg &= vbCr & "Total Detail Sales unequal Returns Total."
                End If

                If EMsg.Length = 0 AndAlso chkNoImpact.Checked Then
                    Dim msg As String = "You selected 'No Impact to Qty On Hand' this means the returned items will not be put back into inventory"
                    If dst.Tables("SOTRTRN2").Select("RTRN_QTY_1 > 0 or RTRN_QTY_2 > 0 or RTRN_QTY_4 > 0 ").Length > 0 Then
                        msg &= " and all returned quantities will be placed in Destroy"
                    End If
                    msg &= "." & Environment.NewLine & "Do you want to continue with the Update?"

                    If MessageBox.Show(msg, "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If

                    For Each rowSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Select("", "", DataViewRowState.CurrentRows)
                        rowSOTRTRN2.Item("RTRN_QTY_3") += Val(rowSOTRTRN2.Item("RTRN_QTY_1") & String.Empty) _
                                                        + Val(rowSOTRTRN2.Item("RTRN_QTY_2") & String.Empty) _
                                                        + Val(rowSOTRTRN2.Item("RTRN_QTY_3") & String.Empty) _
                                                        + Val(rowSOTRTRN2.Item("RTRN_QTY_4") & String.Empty)
                        rowSOTRTRN2.Item("RTRN_QTY_1") = 0
                        rowSOTRTRN2.Item("RTRN_QTY_2") = 0
                        rowSOTRTRN2.Item("RTRN_QTY_4") = 0
                    Next
                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Refresh"

            Case "Reverse"

                If ASCMAIN1.CLIENT = "INT" AndAlso Not ASCMAIN1.Running_in_VS Then ' INTX
                    MessageBox.Show("You are not permitted to perform reversals.", "Reverse", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub

                Else
                    If MessageBox.Show("Are you sure you want to reverse this Entry?", "Confirm Reversal",
                                       MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
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

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Reverse"
                If ASCMAIN1.CLIENT <> "INT" Then
                    EntryMode = "R"
                    Set_Up_Reversal()
                    Update_Record()
                End If

                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "Cancel", "Done", "Refresh"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("Refresh").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode

                    .Items("Reverse").Visible = (ScreenMode AndAlso EntryMode = "V") _
                        AndAlso Not InquiryMode _
                        AndAlso rowSOTRTRN1 IsNot Nothing _
                        AndAlso rowSOTRTRN1.Item("REVERSED_BY_RTRN_NO") & String.Empty = String.Empty _
                        AndAlso rowSOTRTRN1.Item("REVERSED_RTRN_NO") & String.Empty = String.Empty _
                        AndAlso Val(rowSOTRTRN1.Item("RTRN_AS_PO_REC") & String.Empty) <> 1 _
                        AndAlso ASCMAIN1.Running_in_VS '(ASCMAIN1.CLIENT <> "INT" AndAlso Not ASCMAIN1.Running_in_VS)

                    .Items("New").Visible = Not InquiryMode
                    .Items("Done").Visible = (EntryMode = "V" And ScreenMode) Or InquiryMode
                    .Items("Print").Visible = (EntryMode = "V" And ScreenMode) Or InquiryMode
                    .Items("Update").Visible = (Not (EntryMode = "V") Or Not ScreenMode) And Not InquiryMode
                    .Items("Cancel").Visible = (Not (EntryMode = "V") Or Not ScreenMode) And Not InquiryMode
                End With

                .Groups("GL Distribution").Visible = ScreenMode And (EntryMode = "V") And InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") <> 0
                .Groups("Totals").Visible = ScreenMode
                .Groups("Events").Visible = ScreenMode And (EntryMode <> "N")
                .Groups("Item Scan").Visible = ScreenMode AndAlso Not InquiryMode And (EntryMode <> "V")

            End With
        End If

        Set_Read_Only(grpTotals, (EntryMode = "V"))
        Set_Read_Only(UltraGroupBox1, ScreenMode)
        splHeader.Visible = ScreenMode
        grpHeader.Visible = ScreenMode

        tab0.Tabs("3PL").Visible = Not InquiryMode
        tab0.Tabs("Invoice History").Visible = Not InquiryMode AndAlso ASCMAIN1.DBS_COMPANY <> "AHA"

        tab0.Visible = Not ScreenMode
        Setup_tab0()

        If ScreenMode Then
            If InquiryMode Then
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("New").Visible = False
                    .Items("Update").Visible = False
                    .Items("Cancel").Visible = False
                End With
            End If

            tabDetails.Tabs("Sales History").Visible = (EntryMode = "N")
            tabDetails.SelectedTab = tabDetails.Tabs("Sales History")
            tabDetails.Tabs("GL Distribution").Visible = (EntryMode = "V") And ASCMAIN1.USER_SECURITY_CODEs.Contains("X5")

            Set_Read_Only(grpHeader, (EntryMode = "V"))
            'Set_Read_Only(splGL, (EntryMode = "V"))
            'Set_Read_Only_for_ctl(chkNoImpact, True)

            If dst.Tables("SOTRMAF1").Rows.Count > 0 Then
                optB.Enabled = False
                ' MyBase.Absx1.txtFor("CUST_STORE_NO").ReadOnly = True
            Else
                optB.Enabled = True
                ' MyBase.Absx1.txtFor("CUST_STORE_NO").ReadOnly = False
            End If

            If EntryMode = "N" Then
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTRTRN2}
                    With grd.DisplayLayout.Override
                        If whse_is_a_3PL Then
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            .AllowDelete = DefaultableBoolean.False
                            .AllowUpdate = DefaultableBoolean.True
                        Else
                            If INV_NO_RETURNED = "" AndAlso RA_NO = "" Then
                                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                            Else
                                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            End If
                            .AllowDelete = DefaultableBoolean.True
                            .AllowUpdate = DefaultableBoolean.True
                        End If
                    End With
                Next
                With grdSOTRTRN2.DisplayLayout.Bands(0)
                    If whse_is_a_3PL Then
                        .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                        .Columns("RTRN_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                    Else
                        .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                        '.Columns("RTRN_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                    End If
                End With
            Else
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTRTRN2, grdSOTRTRN3}
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    End With
                Next
                With grdSOTRTRN2.DisplayLayout.Bands(0)
                    .Columns("ITEM_CODE").CellAppearance.BackColor = Color.Empty
                    .Columns("RTRN_QTY").CellAppearance.BackColor = Color.Empty
                End With
            End If

        Else
            Clear_Record()

            Show_Filter(grdSOTRMAFX, True)
            grdSOTRMAFX.DisplayLayout.GroupByBox.Hidden = False
        End If

        If UltraExplorerBar1.Groups("Item Scan").Visible Then
            timItemCode.Start()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        ' SR-6549 - Lot Numbers on Shipments and Receipts
        For Each TABLE_NAME As String In New String() {"SOTRTRN0", "SOTRTRN1", "SOTRTRN2", "SOTRTRN3", "SOTINVHH", "SOTINVHX", "SOTRMAF1", "SOTRMAF2",
                                                       "SOTINVH1", "SOTINVH2", "ARTOPEN1", "ICTIXFR1", "ICTIXFR2", "EDTRTRN1", "EDTRTRN2",
                                                       "ICTIADJ1", "ICTIADJ2", "SOTRTRNB",
                                                       "ICTIREC1", "ICTIREC2", "ICTIREC4", "ICTIRECL",
                                                       "ICTPINV1", "ICTPINV2", "APTINVH5_VAR", "POTORDR1", "POTORDR2",
                                                       "APTINVH1", "APTINVH2", "APTINVH5", "APTVEND5", "SOTORDR0", "SOTORDR1", "SOTORDR2", "SOTORDR5"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        If dst.Tables.Contains("ARTCUST3PL") Then
            dst.Tables("ARTCUST3PL").Rows.Clear()
        End If

        EnforceConstraints(True)

        If chkGL.Checked Then
            chkGL.Checked = False
        Else
            Refresh_Documents()
        End If

        Setup_tab0_GL()

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            Setup_3PL_VAN()
        Else
            Setup_3PL()
            'tab0.Tabs("3PL").Visible = False
        End If

        INV_NO_RETURNED = ""
        KEY_3PL_RECORD = ""
        Absx1.txtFor("WHSE_CODE").Clear()
        Absx1.dteFor("RTRN_DATE").Value = IIf(Now > PRD_END_DATE, PRD_END_DATE, Format(Now, "MM/dd/yyyy"))
        Absx1.txtFor("RTRN_NO").Clear()
        Absx1.txtFor("EDI_COMMENTS").Clear()
        chkNoImpact.Checked = False
        rowEDTRTRN1 = Nothing
        tblEDTRTRN1 = Nothing
        pl3Cust_store_no = String.Empty

        Load_SOTRMAFX()
        Load_SOT3PLF1()

        RA_NO = String.Empty
        RTRN_NO = String.Empty

        EDI_DOC_SEQ_NO = String.Empty

        optGL.Tag = ""

        If processing3PL Then
            tab0.SelectedTab = tab0.Tabs("3PL")
        End If
        processing3PL = False
        EDI_DOC_SEQ_NO_List.Clear()

        Absx1.numFor("RTRN_AMOUNT").Value = 0
        Absx1.numFor("RTRN_SALES").Value = 0
        Absx1.numFor("RTRN_FREIGHT").Value = 0
        Absx1.numFor("RTRN_HANDLING").Value = 0
        Absx1.numFor("RTRN_COSTS").Value = 0

        rowWHTTPLP1 = Nothing

        ' Default Values
        IC_PARM_WHSE_CODE_RFB = ROWs("ICTPARM1").Item("IC_PARM_WHSE_CODE_RFB") & String.Empty
        IC_PARM_WHSE_CODE_DST = ROWs("ICTPARM1").Item("IC_PARM_WHSE_CODE_DST") & String.Empty
        IC_PARM_WHSE_CODE_DISC = ROWs("ICTPARM1").Item("IC_PARM_WHSE_CODE_DISC") & String.Empty
        IC_PARM_WHSE_CODE_RTN = ROWs("ICTPARM1").Item("IC_PARM_WHSE_CODE_RTN") & String.Empty
        IC_PARM_WHSE_CODE = ROWs("ICTPARM1").Item("IC_PARM_WHSE_CODE") & String.Empty


    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        Dim rowSOTRMAF1 As DataRow = Nothing
        Dim rowARTCUST1 As DataRow = Nothing
        RTRN_NO = Absx1.txtFor("RTRN_NO").Text

        If EntryMode = "N" Then

            If EDI_DOC_SEQ_NO_List.Count = 0 Then
                EDI_DOC_SEQ_NO_List.Add("1")
            End If

            ' Grab RA NO
            rowARTCUST1 = LookUp("ARTCUST1", HFs("CUST_CODE"))
            If RA_NO.Length > 0 Then
                Fill_Records("SOTRMAF1", RA_NO)
                Fill_Records("SOTRMAF2", RA_NO)
                HFs("CUST_NAME") = rowARTCUST1.Item("CUST_NAME") & String.Empty
                If dst.Tables("SOTRMAF1").Rows.Count > 0 Then
                    rowSOTRMAF1 = dst.Tables("SOTRMAF1").Rows(0)
                End If
            End If

            If ASCMAIN1.CLIENT = "INT" Then
                CUST_BILL_TO_CUST = rowARTCUST1.Item("CUST_BILL_TO_CUST") & ""
                If CUST_BILL_TO_CUST = "" Then CUST_BILL_TO_CUST = rowARTCUST1.Item("CUST_CODE")
                rowARTCUST1_BT = LookUp("ARTCUST1", CUST_BILL_TO_CUST)

                PRICE_CLASS_CODE = rowARTCUST1.Item("PRICE_CLASS_CODE")
                Dim rowSOTPCLS1 As DataRow = LookUp("SOTPCLS1", PRICE_CLASS_CODE)
                PRICE_BASE_DPCT = Val(rowSOTPCLS1.Item("PRICE_BASE_DPCT") & "")
                PRICE_BASIS = rowSOTPCLS1.Item("PRICE_BASIS") & ""
                PRICE_LIST_CODE = rowARTCUST1.Item("PRICE_LIST_CODE") & ""
                CUST_CODE_ALLO = rowARTCUST1.Item("CUST_CODE_ALLO") & ""
                PRICE_LIST_CODE_ALLO = ""
                If CUST_CODE_ALLO <> "" Then
                    Dim rowARTCUST1_ALLO As DataRow = LookUp("ARTCUST1", CUST_CODE_ALLO)
                    If rowARTCUST1_ALLO IsNot Nothing Then
                        PRICE_LIST_CODE_ALLO = rowARTCUST1_ALLO.Item("PRICE_LIST_CODE") & ""
                    End If
                End If
            End If

            Dim rowEDTRTRN1loop As DataRow = Nothing
            Dim tableName As String = String.Empty

            For Each EDI_DOC_SEQ_NO As String In EDI_DOC_SEQ_NO_List

                rowEDTRTRN1loop = Nothing

                If tableName = String.Empty Then
                    tableName = "SOTRTRN1"
                Else
                    tableName = "SOTRTRNB"
                End If

                rowSOTRTRN1 = dst.Tables(tableName).NewRow

                If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                    RTRN_NO = ASCMAIN1.Next_Control_No("TRAN_NO_C")
                Else
                    RTRN_NO = ASCMAIN1.Next_Control_No("SOTRTRN1.RTRN_NO")
                End If
                rowSOTRTRN1.Item("RTRN_NO") = RTRN_NO
                ' No credits prior to todays date. If period issue then it wil be fixed below
                If HFs("RTRN_DATE") < DateTime.Now Then
                    HFs("RTRN_DATE") = DateTime.Now
                End If

                rowSOTRTRN1.Item("CUST_CODE") = HFs("CUST_CODE")
                rowSOTRTRN1.Item("CUST_NAME") = HFs("CUST_NAME")
                rowSOTRTRN1.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO

                If processing3PL Then
                    rowEDTRTRN1loop = dst.Tables("EDTRTRN1").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")(0)
                    rowSOTRTRN1.Item("WHSE_CODE") = rowEDTRTRN1loop.Item("WHSE_CODE") & String.Empty
                Else
                    rowSOTRTRN1.Item("WHSE_CODE") = HFs("WHSE_CODE")
                End If

                If processing3PL AndAlso ASCMAIN1.CLIENT = "INT" Then

                    Dim EDI_CUSTOMER_NO As String = rowEDTRTRN1loop.Item("EDI_CUSTOMER_NO") & String.Empty
                    Dim EDI_CUST_SHIP_TO As String = rowEDTRTRN1loop.Item("EDI_CUST_SHIP_TO") & String.Empty
                    Dim CUST_CODE As String = String.Empty
                    Dim CUST_STORE_NO As String = String.Empty

                    Select Case rowSOTRTRN1.Item("WHSE_CODE") & String.Empty
                        Case "ADSRTN"
                            CUST_CODE = EDI_CUSTOMER_NO
                            CUST_STORE_NO = EDI_CUST_SHIP_TO

                        Case Else
                            Convert3PLDoor(EDI_CUSTOMER_NO, EDI_CUST_SHIP_TO, CUST_CODE, CUST_STORE_NO)
                    End Select

                    Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
                    If rowARTCUST2 IsNot Nothing Then
                        rowSOTRTRN1.Item("CUST_STORE_NO") = rowARTCUST2.Item("CUST_STORE_NO")
                        rowSOTRTRN1.Item("CUST_STORE_NAME") = rowARTCUST2.Item("CUST_STORE_NAME")
                    End If
                End If

                rowSOTRTRN1.Item("RTRN_DATE") = IIf(CDate(HFs("RTRN_DATE")) > PRD_END_DATE, PRD_END_DATE, HFs("RTRN_DATE"))
                rowSOTRTRN1.Item("RTRN_DATE") = CDate(rowSOTRTRN1.Item("RTRN_DATE")).ToShortDateString
                rowSOTRTRN1.Item("RTRN_SOURCE") = "E"
                rowSOTRTRN1.Item("RTRN_STATUS") = "O"
                rowSOTRTRN1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
                rowSOTRTRN1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowSOTRTRN1.Item("INIT_DATE") = DATETIME_STAMP
                rowSOTRTRN1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowSOTRTRN1.Item("LAST_DATE") = DATETIME_STAMP
                rowSOTRTRN1.Item("REGISTER_IND") = "0"
                rowSOTRTRN1.Item("RTRN_FREIGHT") = 0
                rowSOTRTRN1.Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE") & String.Empty
                rowSOTRTRN1.Item("POST_CODE") = rowARTCUST1.Item("POST_CODE")

                If rowSOTRMAF1 IsNot Nothing Then
                    rowSOTRTRN1.Item("RA_NO") = rowSOTRMAF1.Item("RA_NO")
                    rowSOTRTRN1.Item("CUST_CLAIM_NO") = rowSOTRMAF1.Item("CUST_CLAIM_NO")
                    If rowSOTRTRN1.Item("CUST_STORE_NO") & String.Empty = String.Empty Then
                        rowSOTRTRN1.Item("CUST_STORE_NO") = rowSOTRMAF1.Item("CUST_STORE_NO")
                    End If
                    rowSOTRTRN1.Item("SALES_DIVISION_CODE") = rowSOTRMAF1.Item("SALES_DIVISION_CODE")
                    rowSOTRTRN1.Item("SREP_CODE") = rowSOTRMAF1.Item("SREP_CODE")
                    rowSOTRTRN1.Item("REASON_CODE") = rowSOTRMAF1.Item("RA_REASON_CODE")

                    Dim rowEDTRTRN1 As DataRow = Nothing
                    If dst.Tables("EDTRTRN1").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'").Length > 0 Then
                        rowEDTRTRN1 = dst.Tables("EDTRTRN1").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")(0)
                    End If

                    If rowEDTRTRN1 IsNot Nothing AndAlso rowEDTRTRN1.Item("CUST_CLAIM_NO") & String.Empty <> String.Empty Then
                        Select Case ASCMAIN1.CLIENT
                            Case "INT"
                                'Dim tblEDTRTRN2 As DataTable = ASCDATA1.GetDataTable($"Select * from EDTRTRN2 WHERE EDI_DOC_SEQ_NO = '{EDI_DOC_SEQ_NO}'")
                                'If tblEDTRTRN2.Select("EDI_REASON_CODE = 'C15'").Length > 0 Then
                                '    rowSOTRTRN1.Item("CUST_CLAIM_NO") = rowEDTRTRN1.Item("CUST_CLAIM_NO")
                                'End If

                                ' Change requested by Petra on 08/10/2010
                                If rowSOTRTRN1.Item("CUST_CLAIM_NO") & String.Empty = String.Empty Then
                                    rowSOTRTRN1.Item("CUST_CLAIM_NO") = rowEDTRTRN1.Item("CUST_CLAIM_NO")
                                End If

                            Case Else
                                rowSOTRTRN1.Item("CUST_CLAIM_NO") = rowEDTRTRN1.Item("CUST_CLAIM_NO")
                        End Select
                    End If

                    optB.Enabled = False
                ElseIf EDI_DOC_SEQ_NO.Length > 0 AndAlso EDI_DOC_SEQ_NO <> "1" AndAlso ASCMAIN1.CLIENT = "INT" Then
                    Dim rowEDTRTRN1 As DataRow = dst.Tables("EDTRTRN1").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")(0)
                    rowSOTRTRN1.Item("RA_NO") = RA_NO
                    rowSOTRTRN1.Item("CUST_CLAIM_NO") = rowEDTRTRN1.Item("CUST_CLAIM_NO")

                    Dim rowARTCUST2 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST2 WHERE CUST_NO_3PL = '" & rowEDTRTRN1.Item("EDI_CUSTOMER_NO") & "' AND CUST_STORE_NO_3PL = '" & rowEDTRTRN1.Item("EDI_CUST_SHIP_TO") & "'")
                    If rowARTCUST2 IsNot Nothing AndAlso rowARTCUST1.Item("CUST_CODE") = MyBase.Absx1.txtFor("CUST_CODE").Text Then
                        rowSOTRTRN1.Item("CUST_STORE_NO") = rowARTCUST2.Item("CUST_STORE_NO")
                        rowSOTRTRN1.Item("SREP_CODE") = rowARTCUST2.Item("SREP_CODE")
                        If rowSOTRTRN1.Item("SREP_CODE") & String.Empty = String.Empty Then
                            rowSOTRTRN1.Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE")
                        End If
                    End If

                    ASCMAIN1.sql = " SELECT SALES_DIVISION_CODE, COUNT(*) NUM_RECS" _
                        & " FROM ICTITEM1, EDTRTRN2" _
                        & " WHERE ICTITEM1.ITEM_CODE = EDTRTRN2.EDI_ITEM_CODE" _
                        & " AND EDTRTRN2.EDI_DOC_SEQ_NO IN ('" & String.Join("', '", EDI_DOC_SEQ_NO_List) & "')" _
                        & " GROUP BY SALES_DIVISION_CODE"
                    Dim tblICTITEM1 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
                    If tblICTITEM1.Rows.Count > 0 Then
                        rowSOTRTRN1.Item("SALES_DIVISION_CODE") = tblICTITEM1.Select("", "NUM_RECS DESC")(0).Item("SALES_DIVISION_CODE") & String.Empty
                    End If
                    'rowSOTRTRN1.Item("REASON_CODE") = rowSOTRMAF1.Item("RA_REASON_CODE")
                    optB.Enabled = False
                Else
                    optB.Enabled = True
                End If

                dst.Tables(tableName).Rows.Add(rowSOTRTRN1)
            Next
        Else

            rowSOTRTRN1 = Fill_Record("SOTRTRN1", RTRN_NO)
            'rowSOTRTRN1 = dst.Tables("SOTRTRN1").Rows(0)
            dst.AcceptChanges()

            RA_NO = (rowSOTRTRN1.Item("RA_NO") & String.Empty).ToString.Trim
            If RA_NO.Length > 0 Then
                RA_NO = rowSOTRTRN1.Item("RA_NO") & String.Empty
                Fill_Records("SOTRMAF1", RA_NO)
                Fill_Records("SOTRMAF2", RA_NO)
            End If

            dst.Tables("SOTRTRN0").Rows.Add(New String() {"Entered", Format(rowSOTRTRN1.Item("INIT_DATE"), "MM/dd/yy hh:mm tt")})
            dst.Tables("SOTRTRN0").Rows.Add(New String() {"By", rowSOTRTRN1.Item("INIT_OPER")})
            dst.Tables("SOTRTRN0").Rows.Add(New String() {"Source", rowSOTRTRN1.Item("RTRN_SOURCE")})
            Dim RTRN_AS_PO_REC As String = rowSOTRTRN1.Item("RTRN_AS_PO_REC") & ""
            If RTRN_AS_PO_REC = "1" Then
                dst.Tables("SOTRTRN0").Rows.Add(New String() {"Note", "Rtn -> PO Rec"})
            End If


            If rowSOTRTRN1.Item("REVERSED_BY_RTRN_NO") & "" <> "" Then
                Dim row As DataRow = LookUp("SOTRTRN1", rowSOTRTRN1.Item("REVERSED_BY_RTRN_NO"))
                dst.Tables("SOTRTRN0").Rows.Add(New String() {"Reversed", Format(row.Item("INIT_DATE"), "MM/dd/yy hh:mm tt")})
                dst.Tables("SOTRTRN0").Rows.Add(New String() {"By", row.Item("INIT_OPER")})
                dst.Tables("SOTRTRN0").Rows.Add(New String() {"using", rowSOTRTRN1.Item("REVERSED_BY_RTRN_NO")})
            ElseIf rowSOTRTRN1.Item("REVERSED_RTRN_NO") & "" <> "" Then
                dst.Tables("SOTRTRN0").Rows.Add(New String() {"Reverses", rowSOTRTRN1.Item("REVERSED_RTRN_NO")})
            End If
        End If

        rowICTWHSE1 = LookUp("ICTWHSE1", rowSOTRTRN1.Item("WHSE_CODE"))
        location_support = (rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1")
        whse_is_a_3PL = (rowICTWHSE1.Item("LP_CODE") & "" <> "")
        ' whse_is_a_3PL = True
        'With grdSOTRTRN2.DisplayLayout.Bands(0)
        '    .Columns("BAR_CODE").Hidden = True ' Not location_support
        '    .Columns("LOCATION_CODE").Hidden = Not location_support
        'End With

        ' Make sure we are pointing to the record in SOTRTRN1
        rowSOTRTRN1 = dst.Tables("SOTRTRN1").Rows(0)

        If EntryMode = "N" AndAlso whse_is_a_3PL AndAlso 1 = 2 Then

            ASCMAIN1.sql = "Select RCPTHDR.TRANS_SEQ, RCPTHDR.ARRDTE, RCPTHDR.PO_SHIPMENT_NO, RCPTHDR.CONTAINER_NO " _
                & ", RCPTDTL.ITEM_CODE, RCPTDTL.RCVQTY" _
                & " from ADS.RCPTHDR@ADSIIS,ADS.RCPTDTL@ADSIIS where RCPTHDR.STATUS in ('0','V')" _
                & " and RCPTHDR.INVTYP = 'R'" _
                & "AND RCPTHDR.TRANS_SEQ = RCPTDTL.TRANS_SEQ " _
                & "AND RCPTHDR.LP_CODE = RCPTDTL.LP_CODE " _
                & "AND RCPTHDR.WHSE_CODE = RCPTDTL.WHSE_CODE " _
                & "AND RCPTHDR.WHSE_CODE = '" & rowICTWHSE1.Item("WHSE_CODE") & "'"
        End If

        Fill_Records("SOTRTRN2", RTRN_NO)
        Fill_Records("SOTRTRN3", RTRN_NO)

        If INV_NO_RETURNED <> "" Then
            rowSOTRTRN1.Item("INV_NO_RETURNED") = INV_NO_RETURNED
            sotrtrn2RtrnNo = rowSOTRTRN1.Item("RTRN_NO")
            ASCMAIN1.sql = "Select * from SOTINVH2 where INV_TYPE = 'I' and INV_NO = '" & INV_NO_RETURNED & "'"
            ' PERMIT REVERSING ENTIRE CREDITS FOR RETURNS AS WELL AS INVOICES
            ASCMAIN1.sql = "Select * from SOTINVH2 where INV_NO = '" & INV_NO_RETURNED & "'"
            grdSOTRTRN2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "INV_LNO")
                grdSOTRTRN2.DisplayLayout.Bands(0).AddNew()
                With grdSOTRTRN2.ActiveRow
                    .Cells("ITEM_CODE").Value = row.Item("ITEM_CODE")
                    .Cells("RTRN_QTY").Value = row.Item("ORDR_QTY_SHIP")
                    .Cells("RTRN_QTY_1").Value = row.Item("ORDR_QTY_SHIP")
                    .Cells("RTRN_PRICE").Value = row.Item("ORDR_UNIT_PRICE")
                    .Update()
                End With
            Next
            grdSOTRTRN2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            Sort_grdColumns(grdSOTRTRN2, "RTRN_NO,RTRN_LNO")
            Absx1.txtFor("CUST_CLAIM_NO").Text = grdSOTINVHH.ActiveRow.Cells("ORDR_CUST_PO").Value
            Absx1.txtFor("CUST_STORE_NO").Text = grdSOTINVHH.ActiveRow.Cells("CUST_STORE_NO").Value
            Absx1.numFor("RTRN_FREIGHT").Value = grdSOTINVHH.ActiveRow.Cells("INV_FREIGHT").Value
            rowSOTRTRN1.Item("CUST_STORE_NO") = grdSOTINVHH.ActiveRow.Cells("CUST_STORE_NO").Value

        ElseIf (EntryMode = "N" AndAlso RA_NO.Length > 0 AndAlso processing3PL = True) Then

            If dst.Tables("SOTRMAF1").Rows.Count > 0 Then
                rowSOTRTRN1.Item("RA_NO") = RA_NO
                rowSOTRTRN1.Item("REASON_CODE") = dst.Tables("SOTRMAF1").Rows(0).Item("RA_REASON_CODE") & String.Empty
                rowSOTRTRN1.Item("RTRN_NOTE") = dst.Tables("SOTRMAF1").Rows(0).Item("RA_NOTES") & String.Empty

                For Each rowSOTRTRNB As DataRow In dst.Tables("SOTRTRNB").Select()
                    EDI_DOC_SEQ_NO = rowSOTRTRNB.Item("EDI_DOC_SEQ_NO") & String.Empty
                    rowEDTRTRN1 = dst.Tables("EDTRTRN1").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")(0)
                    rowSOTRTRNB.Item("RA_NO") = rowEDTRTRN1.Item("EDI_RA_NO") & String.Empty
                    rowSOTRTRNB.Item("REASON_CODE") = dst.Tables("SOTRMAF1").Rows(0).Item("RA_REASON_CODE") & String.Empty
                    rowSOTRTRNB.Item("RTRN_NOTE") = dst.Tables("SOTRMAF1").Rows(0).Item("RA_NOTES") & String.Empty
                    rowSOTRTRNB.Item("CUST_CLAIM_NO") = rowEDTRTRN1.Item("CUST_CLAIM_NO") & String.Empty
                    rowSOTRTRNB.Item("KEY_3PL_RECORD") = EDI_DOC_SEQ_NO
                Next
            End If

            EDI_DOC_SEQ_NO = dst.Tables("SOTRTRN1").Rows(0).Item("EDI_DOC_SEQ_NO") & String.Empty
            rowEDTRTRN1 = dst.Tables("EDTRTRN1").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")(0)
            If rowEDTRTRN1 IsNot Nothing Then
                Absx1.txtFor("EDI_COMMENTS").Text = (rowEDTRTRN1.Item("EDI_COMMENTS_1") & " " & rowEDTRTRN1.Item("EDI_COMMENTS_2")).ToString.Replace("  ", " ")
                Absx1.txtFor("CUST_CLAIM_NO").Text = rowEDTRTRN1.Item("CUST_CLAIM_NO") & String.Empty
            End If
            Absx1.txtFor("KEY_3PL_RECORD").Text = EDI_DOC_SEQ_NO
            rowSOTRTRN1.Item("KEY_3PL_RECORD") = EDI_DOC_SEQ_NO


            grdSOTRTRN2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes

            ASCMAIN1.sql = " SELECT SOTINVH2.ITEM_CODE, SUBSTR(SOTINVH1.ORDR_YYYYPP_UPDATED, 1, 4) PERIOD, MIN(SOTINVH2.ORDR_UNIT_PRICE) ORDR_UNIT_PRICE" _
                & " FROM SOTINVH1, SOTINVH2, EDTRTRN1, EDTRTRN2" _
                & " WHERE SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" _
                & " AND SOTINVH1.INV_NO = SOTINVH2.INV_NO" _
                & " AND SOTINVH2.ITEM_CODE = EDTRTRN2.EDI_ITEM_CODE " _
                & " AND EDTRTRN1.EDI_DOC_SEQ_NO = EDTRTRN2.EDI_DOC_SEQ_NO" _
                & " AND SOTINVH1.INV_DATE < EDTRTRN1.EDI_RETURN_DATE" _
                & " AND SOTINVH1.CUST_CODE = '" & HFs("CUST_CODE") & "'" _
                & " AND EDTRTRN2.EDI_DOC_SEQ_NO IN ('" & String.Join("', '", EDI_DOC_SEQ_NO_List) & "')" _
                & " GROUP BY SOTINVH2.ITEM_CODE, SUBSTR(SOTINVH1.ORDR_YYYYPP_UPDATED, 1 ,4)"
            Dim tblPrices As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

            sotrtrn2RtrnNo = String.Empty

            For Each EDI_DOC_SEQ_NO As String In EDI_DOC_SEQ_NO_List

                Dim EDI_RA_NO_WK As String = String.Empty

                sotrtrn2RtrnNo = String.Empty
                If dst.Tables("SOTRTRN1").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'").Length = 1 Then
                    sotrtrn2RtrnNo = dst.Tables("SOTRTRN1").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")(0).Item("RTRN_NO") & String.Empty
                    EDI_RA_NO_WK = dst.Tables("SOTRTRN1").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")(0).Item("RA_NO") & String.Empty
                ElseIf dst.Tables("SOTRTRNB").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'").Length = 1 Then
                    EDI_RA_NO_WK = dst.Tables("SOTRTRNB").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")(0).Item("RA_NO") & String.Empty
                    sotrtrn2RtrnNo = dst.Tables("SOTRTRNB").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")(0).Item("RTRN_NO") & String.Empty
                End If

                If sotrtrn2RtrnNo = String.Empty Then
                    Throw New Exception("Unable to resolve EDI Doc Sequence No: " & EDI_DOC_SEQ_NO)
                End If

                For Each row As DataRow In dst.Tables("EDTRTRN2").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
                    Dim EDI_ITEM_CODE As String = row.Item("EDI_ITEM_CODE")
                    ' Total Quantity returned
                    Dim EDI_QTY_RETURNED As Int32 = Val(row.Item("EDI_QTY_RETURNED") & String.Empty)
                    ' Breakdown of what was returned
                    Dim EDI_QTY_BACK_TO_STOCK As Int32 = Val(row.Item("EDI_QTY_BACK_TO_STOCK") & String.Empty)
                    Dim EDI_QTY_IN_REPAIR As Int32 = Val(row.Item("EDI_QTY_IN_REPAIR") & String.Empty)
                    Dim EDI_QTY_DAMAGED As Int32 = Val(row.Item("EDI_QTY_DAMAGED") & String.Empty)
                    Dim EDI_QTY_AS_IS As Int32 = Val(row.Item("EDI_QTY_AS_IS") & String.Empty)

                    ' See if we need to add the item to the RA
                    If ASCMAIN1.CLIENT = "INT" Then
                        If dst.Tables("SOTRMAF2").Select("ITEM_CODE = '" & EDI_ITEM_CODE & "'").Length = 0 Then
                            Dim RA_LNO As Int16 = dst.Tables("SOTRMAF2").Compute("MAX(RA_LNO)", "") + 1

                            Dim rowSOTRMAF2 As DataRow = dst.Tables("SOTRMAF2").NewRow
                            rowSOTRMAF2.Item("RA_NO") = RA_NO
                            rowSOTRMAF2.Item("RA_LNO") = RA_LNO
                            rowSOTRMAF2.Item("ITEM_CODE") = EDI_ITEM_CODE
                            rowSOTRMAF2.Item("RA_QTY") = 0
                            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", EDI_ITEM_CODE)
                            If rowICTITEM1 IsNot Nothing Then
                                rowSOTRMAF2.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")

                                ITEM_RETAIL_PRICE = Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & String.Empty)
                                Dim ORDR_UNIT_PRICE As Decimal = TAC.SOCMAIN1.Get_Price _
                                                             (Me,
                                                              PRICE_LIST_CODE,
                                                              PRICE_LIST_CODE_ALLO,
                                                              PRICE_BASIS,
                                                              PRICE_BASE_DPCT,
                                                              EDI_ITEM_CODE,
                                                              rowICTITEM1,
                                                              rowSOTRMAF1.Item("RA_DATE"), ITEM_RETAIL_PRICE) ' MAYBE SHOULD USE RA_DATE - 60

                                rowSOTRMAF2.Item("RA_NET_PRICE") = ORDR_UNIT_PRICE
                                rowSOTRMAF2.Item("RA_RETAIL") = ITEM_RETAIL_PRICE
                                'rowSOTRMAF2.Item("RA_LINE_AMT") = ""
                                rowSOTRMAF2.Item("RA_QTY_OPEN") = 0
                                'rowSOTRMAF2.Item("RA_QTY_USED") = ""
                                'rowSOTRMAF2.Item("RA_QTY_CANC") = ""
                                'rowSOTRMAF2.Item("EDI_PRICE") = ""
                                'rowSOTRMAF2.Item("NET_PRICE") = ""
                                dst.Tables("SOTRMAF2").Rows.Add(rowSOTRMAF2)
                            Else
                                MessageBox.Show($"Invalid Item Code {EDI_ITEM_CODE}, it can not be added to the Return", "Item Code", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            End If
                        End If
                    End If

                    ' If many returns are selected then create one return for each EDI Doc Selected
                    ' Consolidate item to one row by Return No
                    If dst.Tables("SOTRTRN2").Select("ITEM_CODE = '" & EDI_ITEM_CODE & "' AND RTRN_NO = '" & sotrtrn2RtrnNo & "'").Length > 0 Then
                        Dim rowSOTRTRN2 As DataRow = dst.Tables("SOTRTRN2").Select("ITEM_CODE = '" & EDI_ITEM_CODE & "' AND RTRN_NO = '" & sotrtrn2RtrnNo & "'")(0)
                        With rowSOTRTRN2
                            .Item("RTRN_QTY") = Val(.Item("RTRN_QTY") & String.Empty) + EDI_QTY_RETURNED
                            .Item("RTRN_QTY_1") = Val(.Item("RTRN_QTY_1") & String.Empty) + EDI_QTY_BACK_TO_STOCK
                            .Item("RTRN_QTY_2") = Val(.Item("RTRN_QTY_2") & String.Empty) + EDI_QTY_IN_REPAIR
                            .Item("RTRN_QTY_3") = Val(.Item("RTRN_QTY_3") & String.Empty) + EDI_QTY_DAMAGED
                            .Item("RTRN_QTY_4") = Val(.Item("RTRN_QTY_4") & String.Empty) + EDI_QTY_AS_IS
                        End With
                        Continue For
                    End If

                    grdSOTRTRN2.DisplayLayout.Bands(0).AddNew()
                    With grdSOTRTRN2.ActiveRow
                        .Cells("ITEM_CODE").Value = EDI_ITEM_CODE
                        .Cells("EDI_DOC_SEQ_NO").Value = EDI_DOC_SEQ_NO

                        .Cells("RTRN_QTY").Value = EDI_QTY_RETURNED
                        .Cells("RTRN_QTY_1").Value = EDI_QTY_BACK_TO_STOCK
                        .Cells("RTRN_QTY_2").Value = EDI_QTY_IN_REPAIR
                        .Cells("RTRN_QTY_3").Value = EDI_QTY_DAMAGED
                        .Cells("RTRN_QTY_4").Value = EDI_QTY_AS_IS
                        .Cells("KEY_3PL_RECORD").Value = EDI_DOC_SEQ_NO
                        .Cells("OPS_YYYYPP").Value = ASCMAIN1.CYP

                        If .Cells("RTRN_QTY").Value > 0 Then
                            If dst.Tables("SOTRMAF2").Select("ITEM_CODE = '" & EDI_ITEM_CODE & "'").Length > 0 Then
                                .Cells("RTRN_PRICE").Value = dst.Tables("SOTRMAF2").Select("ITEM_CODE = '" & EDI_ITEM_CODE & "'")(0).Item("RA_NET_PRICE")
                            ElseIf ASCMAIN1.CLIENT = "INT" AndAlso tblPrices.Select("ITEM_CODE = '" & EDI_ITEM_CODE & "'").Length > 0 Then
                                .Cells("RTRN_PRICE").Value = tblPrices.Select("ITEM_CODE = '" & EDI_ITEM_CODE & "'", "PERIOD DESC")(0).Item("ORDR_UNIT_PRICE")
                            Else
                                .Cells("RTRN_PRICE").Value = 0
                            End If
                        Else
                            .Cells("RTRN_PRICE").Value = 0
                        End If
                        .Update()
                    End With
                Next
            Next
            grdSOTRTRN2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            Sort_grdColumns(grdSOTRTRN2, "RTRN_NO,RTRN_LNO")

        ElseIf EntryMode = "N" AndAlso RA_NO.Length > 0 Then

            rowSOTRTRN1.Item("RA_NO") = RA_NO
            rowSOTRTRN1.Item("REASON_CODE") = dst.Tables("SOTRMAF1").Rows(0).Item("RA_REASON_CODE") & String.Empty
            rowSOTRTRN1.Item("RTRN_NOTE") = dst.Tables("SOTRMAF1").Rows(0).Item("RA_NOTES") & String.Empty

            grdSOTRTRN2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes
            For Each row As DataRow In dst.Tables("SOTRMAF2").Select("RA_QTY_OPEN > 0")
                grdSOTRTRN2.DisplayLayout.Bands(0).AddNew()
                With grdSOTRTRN2.ActiveRow
                    .Cells("RTRN_NO").Value = RTRN_NO
                    .Cells("RTRN_LNO").Value = row.Item("RA_LNO")
                    .Cells("ITEM_CODE").Value = row.Item("ITEM_CODE")
                    .Cells("RTRN_QTY").Value = row.Item("RA_QTY_OPEN")

                    If .Cells("RTRN_QTY").Value > 0 Then
                        .Cells("RTRN_PRICE").Value = row.Item("RA_NET_PRICE")
                    Else
                        .Cells("RTRN_PRICE").Value = 0
                    End If

                    .Update()
                End With
            Next
            grdSOTRTRN2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            Sort_grdColumns(grdSOTRTRN2, "RTRN_NO,RTRN_LNO")
        End If

        For Each rowSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Rows
            rowSOTRTRN2.Item("RTRN_QTY_ORIG") = rowSOTRTRN2.Item("RTRN_QTY")
        Next

        Select Case ASCMAIN1.CLIENT
            Case "INT"
                Dim rowSOTRTRNB_SpecialReturns As DataRow = Nothing
                Dim RTRN_NO_SpecialReturns As String = String.Empty
                Dim RTRN_LNO As Int32 = Val(dst.Tables("SOTRTRN2").Compute("MAX(RTRN_LNO)", "") & String.Empty)

                ' Mark Items as Specail Returns Items
                For Each rowSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Select("", "ITEM_CODE")
                    rowSOTRTRN2.Item("RTRN_AS_PO_REC") = "0"
                    If tblSpecialReturns.Rows.Find(rowSOTRTRN2.Item("ITEM_CODE") & String.Empty) IsNot Nothing Then
                        rowSOTRTRN2.Item("RTRN_AS_PO_REC") = "1"
                        'MsgBox("We have a LCS item with XXX in the description - " & rowSOTRTRN2.Item("ITEM_CODE") & vbCrLf & "", vbOKOnly, "Please send Screenshot to ABS before clicking OK")
                        'Throw New Exception("Lacost Item " & rowSOTRTRN2.Item("ITEM_CODE") & " has been returned")
                        'Me.Close()
                        'Exit Sub
                    End If
                Next

                For Each tableName As String In New String() {"SOTRTRN1", "SOTRTRNB"}
                    For Each row As DataRow In dst.Tables(tableName).Select("")
                        Dim RTRN_NO As String = row.Item("RTRN_NO")
                        row.Item("RTRN_AS_PO_REC") = "0"

                        ' check to see if this return has nothing in it except for items that will be PO Received
                        If dst.Tables("SOTRTRN2").Select($"ISNULL(RTRN_AS_PO_REC,'0') = '0' AND RTRN_NO = '{RTRN_NO}'").Length = 0 Then ' IE, we have no non-PO_REC items in this return
                            row.Item("RTRN_AS_PO_REC") = "1" ' set the header to indicate that the entire return is PO Receipt items
                            ' just a heads up, the new row SOTRTRNB_SpecialReturns will also wind up here
                            Continue For
                        End If

                        ' if you get to this point, you either have a mix or all non-PO REC
                        ' so if you have any PO Rec items, make yourself a Return Record

                        If dst.Tables("SOTRTRN2").Select($"ISNULL(RTRN_AS_PO_REC,'0') = '1' AND RTRN_NO = '{RTRN_NO}'").Length <> 0 Then ' IE, we have at least 1 PO_REC item in this return
                            If rowSOTRTRNB_SpecialReturns Is Nothing Then
                                RTRN_NO_SpecialReturns = ASCMAIN1.Next_Control_No("SOTRTRN1.RTRN_NO")
                                rowSOTRTRNB_SpecialReturns = dst.Tables("SOTRTRNB").NewRow
                                rowSOTRTRNB_SpecialReturns.ItemArray = row.ItemArray ' dst.Tables("SOTRTRN1").Rows(0).ItemArray
                                rowSOTRTRNB_SpecialReturns.Item("RTRN_NO") = RTRN_NO_SpecialReturns
                                rowSOTRTRNB_SpecialReturns.Item("RTRN_AS_PO_REC") = "1"
                                dst.Tables("SOTRTRNB").Rows.Add(rowSOTRTRNB_SpecialReturns)
                            End If

                            For Each rowSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Select($"RTRN_AS_PO_REC = '1' and RTRN_NO = '{RTRN_NO}'")
                                ' Special items get placed into a separate return; however, if a return contians all special items wethat entire return is identified as having all special item
                                ' and is not split up wher the special items go to the Special Rerturn.
                                RTRN_LNO += 1
                                rowSOTRTRN2.Item("RTRN_LNO") = RTRN_LNO
                                rowSOTRTRN2.Item("RTRN_NO") = RTRN_NO_SpecialReturns
                            Next
                        End If
                    Next
                Next
        End Select

        'If EntryMode = "N" AndAlso Not InquiryMode Then
        DisplayTotals()
        Absx1.numFor("RTRN_AMOUNT").Value _
            = Val(Absx1.numFor("RTRN_SALES").Value & "") _
            + Val(Absx1.numFor("RTRN_FREIGHT").Value & "") _
            + Val(Absx1.numFor("RTRN_HANDLING").Value & "")
        'End If

        'For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSOTRTRN2.Rows
        '    row.Refresh(Infragistics.Win.UltraWinGrid.RefreshRow.FireInitializeRow)
        'Next

        grdSOTRTRN2.DisplayLayout.Bands(0).Columns("RTRN_NO").Hidden = ASCDATA1.SelectDistinct(dst.Tables("SOTRTRN2"), New String() {"RTRN_NO"}).Rows.Count <= 1

        If dst.Tables("SOTRTRN1").Rows(0).Item("NO_INV_IMPACT") & String.Empty = "1" Then
            chkNoImpact.Checked = True
        End If

        'ASCMAIN1.Progress("NO_INV_IMPACT")
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Dim inTransaction As Boolean = False
        Dim rowARTPOST1 As DataRow = LookUp("ARTPOST1", rowSOTRTRN1.Item("POST_CODE") & String.Empty)
        Dim issueCreditCardCredit As Boolean = False

        rowSOTRTRN1 = dst.Tables("SOTRTRN1").Rows(0)

        Try

            ' Only done for the main record in SOTRTRN1
            rowSOTRTRN1.Item("CUST_STORE_NO") = MyBase.Absx1.txtFor("CUST_STORE_NO").Text.Trim
            rowSOTRTRN1.Item("CUST_CLAIM_NO") = MyBase.Absx1.txtFor("CUST_CLAIM_NO").Text.Trim

            ' Update the fields on all the Returns
            For Each tableName As String In New String() {"SOTRTRN1", "SOTRTRNB"}
                For Each rowSOTRTRNX As DataRow In dst.Tables(tableName).Rows

                    If chkNoImpact.Checked Then
                        rowSOTRTRNX.Item("NO_INV_IMPACT") = "1"
                    Else
                        rowSOTRTRNX.Item("NO_INV_IMPACT") = "0"
                    End If

                    rowSOTRTRNX.Item("RTRN_AMOUNT") _
                        = Val(rowSOTRTRNX.Item("RTRN_SALES") & "") _
                        + Val(rowSOTRTRNX.Item("RTRN_STAX") & "") _
                        + Val(rowSOTRTRNX.Item("RTRN_FREIGHT") & "") _
                        + Val(rowSOTRTRNX.Item("RTRN_HANDLING") & "")

                    rowSOTRTRNX.Item("RTRN_NOTE") = MyBase.Absx1.txtFor("RTRN_NOTE").Text.Trim

                    If rowSOTRTRNX.Item("CUST_STORE_NO") & "" = "" Then
                        rowSOTRTRNX.Item("CUST_STORE_NAME") = ""
                    Else
                        rowARTCUST2 = LookUp("ARTCUST2", New String() {rowSOTRTRNX.Item("CUST_CODE") & String.Empty, rowSOTRTRNX.Item("CUST_STORE_NO") & String.Empty})
                        If rowARTCUST2 IsNot Nothing Then
                            rowSOTRTRNX.Item("CUST_STORE_NAME") = rowARTCUST2.Item("CUST_STORE_NAME")
                        End If
                    End If

                    rowSOTRTRNX.Item("RTRN_SOURCE_DOC_NO") = MyBase.Absx1.txtFor("RTRN_SOURCE_DOC_NO").Text.Trim

                Next
            Next

            For Each row As DataRow In dst.Tables("SOTRTRNB").Rows
                dst.Tables("SOTRTRN1").ImportRow(row)
            Next

            dst.Tables("SOTRTRNB").Clear()

            For Each row As DataRow In dst.Tables("SOTRTRN1").Rows
                row.Item("INV_NO") = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
            Next


            ' RTRN_AS_PO_REC
            Dim returnCredit As New SOCINVH1(dst.Tables("SOTINVH1"),
                                             dst.Tables("SOTINVH2"),
                                             dst.Tables("SOTRTRN1"),
                                             dst.Tables("SOTRTRN2"),
                                             dst.Tables("ARTOPEN1"))

            returnCredit.CreateReturnsCredit("RTN")

            Dim CUST_CODE As String = HFs("CUST_CODE")
            MyBase.Absx1.txtFor("CUST_CLAIM_NO").Text = MyBase.Absx1.txtFor("CUST_CLAIM_NO").Text.Trim
            For Each row As DataRow In dst.Tables("SOTRTRN1").Select("", "RTRN_NO")
                Dim INV_NO As String = row.Item("INV_NO")
                dst.Tables("SOTINVH1").Select("INV_NO = '" & INV_NO & "'", "")(0).Item("PARTNER_ORDR_NO") = row.Item("KEY_3PL_RECORD")
                dst.Tables("ARTOPEN1").Select("INV_NUM = '" & INV_NO & "'", "")(0).Item("PARTNER_ORDR_NO") = row.Item("KEY_3PL_RECORD")

                Dim ORDR_CUST_PO As String = row.Item("CUST_CLAIM_NO") & String.Empty

                If ASCMAIN1.CLIENT = "AHA" AndAlso ORDR_CUST_PO.Length > 0 Then
                    Dim INV_NO_CR As String = String.Empty
                    Dim CC_SALE_TRANS_ID As String = String.Empty
                    SOCMAIN1.GetCreditCardSaleTransaction(CUST_CODE, ORDR_CUST_PO, INV_NO_CR, CC_SALE_TRANS_ID)
                    dst.Tables("SOTINVH1").Select("INV_NO = '" & INV_NO & "'", "")(0).Item("CC_SALE_TRANS_ID") = CC_SALE_TRANS_ID
                    dst.Tables("SOTINVH1").Select("INV_NO = '" & INV_NO & "'", "")(0).Item("INV_NO_CR") = INV_NO_CR

                    issueCreditCardCredit = issueCreditCardCredit OrElse (CC_SALE_TRANS_ID <> String.Empty)
                End If
            Next

            BeginTrans()
            inTransaction = True

            If processing3PL And EDI_DOC_SEQ_NO_List.Count > 0 Then
                For Each EDI_DOC_SEQ_NO As String In EDI_DOC_SEQ_NO_List
                    Dim rowSOTRTRNX As DataRow = dst.Tables("SOTRTRN1").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")(0)
                    Dim RTRN_NOx As String = rowSOTRTRNX.Item("RTRN_NO")
                    ASCDATA1.ExecuteSQL("Update EDTRTRN1 SET PROCESS_IND = '1', RTRN_NO = '" & RTRN_NOx & "' WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
                Next
            End If

            Update_Record_TDA("SOTRTRN1")
            Update_Record_TDA("SOTRTRN2")
            Update_Record_TDA("SOTRTRN3")

            Update_Record_TDA("SOTRMAF2")

            Update_Record_TDA("SOTINVH1")
            Update_Record_TDA("SOTINVH2")
            Update_Record_TDA("ARTOPEN1")

            If EntryMode = "R" Then
                Dim row As DataRow = dst.Tables("SOTRTRN1").Rows(0)
                Dim REVERSED_BY_RTRN_NO As String = row.Item("RTRN_NO")
                Dim REVERSED_RTRN_NO As String = row.Item("REVERSED_RTRN_NO")
                KEY_3PL_RECORD = row.Item("KEY_3PL_RECORD") & ""
                ASCMAIN1.sql = "" _
                    & "Begin Declare Cursor C1 is Select * from SOTRTRN1" _
                    & " where RTRN_NO = '" & REVERSED_BY_RTRN_NO & "';" _
                    & " Begin For R1 in C1 Loop" _
                    & "  Update SOTRTRN1 Set REVERSED_BY_RTRN_NO = R1.RTRN_NO" _
                    & ", LAST_DATE = R1.INIT_DATE" _
                    & ", LAST_OPER = R1.INIT_OPER" _
                    & " where RTRN_NO = '" & REVERSED_RTRN_NO & "';" _
                    & " End Loop; End; " _
                    & "End;"
                ASCDATA1.ExecuteSQL()
            ElseIf EntryMode = "N" And RA_NO.Length > 0 Then
                ASCMAIN1.sql = "" _
                    & "Begin Declare Cursor C1 is SELECT SOTRTRN2.ITEM_CODE, SUM(SOTRTRN2.RTRN_QTY) RTRN_QTY" _
                    & " FROM SOTRTRN1, SOTRTRN2" _
                    & " WHERE SOTRTRN1.RTRN_NO = SOTRTRN2.RTRN_NO" _
                    & " AND SOTRTRN1.RA_NO = '" & RA_NO & "' " _
                    & " GROUP BY SOTRTRN2.ITEM_CODE;" _
                    & " Begin For R1 in C1 Loop" _
                    & "     Update SOTRMAF2 " _
                    & "     SET RA_QTY_OPEN = GREATEST(0, RA_QTY - R1.RTRN_QTY)" _
                    & "     , RA_QTY_USED = DECODE(RA_QTY, 0, R1.RTRN_QTY, LEAST(RA_QTY, R1.RTRN_QTY))" _
                    & "     where RA_NO = '" & RA_NO & "' AND ITEM_CODE = R1.ITEM_CODE;" _
                    & " Update SOTRMAF1 SET LAST_OPER = '" & ASCMAIN1.USER_ID & "', LAST_DATE = SYSDATE WHERE RA_NO = '" & RA_NO & "';" _
                    & " End Loop; End; " _
                    & "End;"
                ASCDATA1.ExecuteSQL()

                Dim RA_QTY_OPEN As Int32 = Val(ASCDATA1.GetDataValue("SELECT SUM(RA_QTY_OPEN) FROM SOTRMAF2 WHERE RA_NO = :PARM1", "V", New Object() {RA_NO}) & String.Empty)
                If RA_QTY_OPEN = 0 Then
                    ASCDATA1.ExecuteSQL("UPDATE SOTRMAF1 SET RA_STATUS = 'F' WHERE RA_NO = '" & RA_NO & "'")
                End If

                For Each row As DataRow In dst.Tables("SOTINVH1").Rows
                    Dim INV_NO As String = row.Item("INV_NO")
                    Dim INV_TYPE As String = row.Item("INV_TYPE")
                    ASCDATA1.ExecuteSP("ARPCUST6_IC", "VV",
                        New Object() {INV_TYPE, INV_NO},
                        New String() {"INV_TYPE_IN", "INV_NO_IN"})
                Next
            End If

            ' See if we need to impact inventory
            If Not chkNoImpact.Checked Then
                For Each rowSOTRTRN1 As DataRow In dst.Tables("SOTRTRN1").Select("") ' ISNULL(RTRN_AS_PO_REC, '0') = '0'
                    Dim INV_NO As String = rowSOTRTRN1.Item("INV_NO")
                    Dim RTRN_NO As String = rowSOTRTRN1.Item("RTRN_NO")

                    Dim RTRN_AS_PO_REC As String = rowSOTRTRN1.Item("RTRN_AS_PO_REC") & String.Empty

                    If RTRN_AS_PO_REC <> "1" Then

                        Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").Select("INV_NO = '" & INV_NO & "'")(0)

                        ASCMAIN1.sql = "BEGIN SOPSTAT1('" & rowSOTINVH1.Item("INV_TYPE") & "','" & rowSOTINVH1.Item("INV_NO") & "'); END;"
                        ASCDATA1.ExecuteSQL()

                        ' Call this since it is a Credit.
                        ASCMAIN1.sql = "BEGIN SOPSTAT2('" & rowSOTINVH1.Item("INV_TYPE") & "','" & rowSOTINVH1.Item("INV_NO") & "'); END;"
                        ASCDATA1.ExecuteSQL()

                        ' Do the following if the warehouse use a locator system
                        If location_support Then
                            TAC.ICCMAIN1.Update_WHTLOCBX("N", dst.Tables("SOTRTRN1").Rows(0).Item("RTRN_NO"))
                        End If
                    End If

                    CreateTransferOrAdjustment(RTRN_NO)
                Next

                TAC.ICCMAIN1.Update_Adjustment(Me)
                TAC.ICCMAIN1.Update_Transfer(Me)
                If location_support Then
                    Update_WHTLOCBX()
                End If

            End If

            ' Call Rob's routine for Kate Spade Items.
            If dst.Tables("SOTRTRN1").Select("RTRN_AS_PO_REC = '1'").Length > 0 Then
                TAC.SOCMAIN1.Create_PO_Rec_From_Return(Me)
            End If

            CommitTrans("Update Complete")

            ' See if we need to issue credit card credit.
            If issueCreditCardCredit AndAlso rowSOTRTRN1.Item("RTRN_AMOUNT") <> 0 Then
                ' This is done since paypal transaction IDs cannt be used for Authorize.net
                If MessageBox.Show("Do you want to refund the Credit Card?", "Refund", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    ASCMAIN1.Progress("Processing CC Credit", "")
                    Dim errorMessage As String = String.Empty
                    If Not SOCMAIN1.IssueCredit(rowSOTRTRN1.Item("INV_NO"), errorMessage) Then
                        MessageBox.Show("Error Processing Credit Card Refund: " & errorMessage, "CC Credit", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                    ASCMAIN1.Progress("", "")
                End If
            End If

        Catch ex As Exception
            If inTransaction Then Rollback()
            inTransaction = False
            MessageBox.Show(ex.Message)

        Finally
            ASCMAIN1.Progress("", "")
        End Try

    End Sub

    Sub Print_Record()
        Me.Cursor = Cursors.WaitCursor

        ASCMAIN1.Progress("Now Preparing for Printing")

        Dim REPORT_NAME As String = "SORINVP1"
        If Not REPORTS.ContainsKey(REPORT_NAME) Then
            REPORTS.Add(REPORT_NAME, Load_rptClass(REPORT_NAME))
            REPORTS(REPORT_NAME).Prepare_dst(False, "")
        End If

        Dim INV_NO As String = rowSOTRTRN1.Item("INV_NO")
        Dim sqlw As String = " and SOTINVH1.INV_TYPE = 'C' and SOTINVH1.INV_NO = '" & INV_NO & "'"
        REPORTS(REPORT_NAME).Fill_Records_RPT(New Object() {sqlw, True, "C"})
        Dim FILENAME As String = ""
        With REPORTS(REPORT_NAME).clsASCBASE1
            .Print_Report_Begin()
            .CR_params.Add("SUBT", "")
            .CR_params.Add("CONS_INV", "0")
            Dim REPORT_NO As String = .Generate_Report(REPORT_NAME, "", "", False, False, "", "PDF", INV_NO, False)
            FILENAME = .F.REPORT_FILENAMES(REPORT_NO)
            .Print_Report_End(, True)
        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Show_Document(FILENAME)
        '  Return FILENAME
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control,
         ByVal COLUMN_NAME As String,
         Optional ByRef sql_where As String = "",
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "REASON_CODE"
                sql_where &= " AND RETURN_IND = '1'"

                'Case "RSRV_NO"

                '    If Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("ORDR_CUST_PO").Text = "" Then
                '        MsgBox("You must enter a Customer Code or a PO No", vbOKOnly, "Cannot Perform Requested Action")
                '        Cancel = True
                '        Exit Sub
                '    End If
                '    sql_where = ""

                '    If InquiryMode Then
                '    Else
                '        sql_where &= " and SOTRSRV1.RSRV_STATUS = 'O' "
                '    End If

                '    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                '        sql_where &= " and SOTRSRV1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                '    End If
                '    If Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                '        sql_where &= " and SOTRSRV1.ORDR_CUST_PO = '" & Absx1.txtFor("ORDR_CUST_PO").Text & "'"
                '    End If
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

            Case "View", "Edit"
                'If ScreenMode Then
                '    Click_Command("Done")
                'End If

                Absx1.txtFor("RTRN_NO").Text = key
                Click_Command("View")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "SOTRTRN1"
            E.COLUMN_NAME = "RTRN_NO"
            E.CODE_VALUE = Absx1.txtFor("RTRN_NO").Text
            E.DESC_VALUE = "Return"
            E.ATTACHMENT_NOTES = ""
            'If rowSOTRSRV1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTRTRNX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTRTRNX_D, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTRTRNX_I, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTRTRN2, "B", "Item Status Inquiry", "Copy Price to All Lines", "Copy All Lines to Negate Inventory Impact")
        Load_Popup_Menu(grdSOTRTRN3, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTINVHH, "SS", "Show Filter", "Show GroupBox", "Credit Entire Invoice")
        Load_Popup_Menu(grdSOTINVHX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTRMAFX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        If 1 = 2 Then
            Load_Popup_Menu(grdEDTRTRN1, "SSPBBBBBB", "Show Filter", "Show GroupBox", "Modify RA Number", "3PL Returns Report", "Remove From List", "Change Customer", "Change Ship To", "Reship to AE")
        Else
            Load_Popup_Menu(grdEDTRTRN1, "SSPBBBBB", "Show Filter", "Show GroupBox", "Modify RA Number", "3PL Returns Report", "Remove From List", "Change Customer", "Change Ship To")
        End If
        Load_Popup_Menu(grdSOT3PLF1, "SSPB", "Show Filter", "Show GroupBox", "Set RA Number")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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

        Select Case grd.Name
            Case "grdSOTINVHH"
                tlb_btn = DirectCast(tlb_pop.Tools("Credit Entire Invoice"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not InquiryMode

            Case "grdSOT3PLF1"
                tlb_btn = DirectCast(tlb_pop.Tools("Set RA Number"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = grdEDTRTRN1.ActiveRow IsNot Nothing

            Case "grdEDTRTRN1"
                tlb_btn = DirectCast(tlb_pop.Tools("Change Customer"), UltraWinToolbars.ButtonTool)

                If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.Band.Index <> 0 Then
                    tlb_btn.SharedProps.Visible = False
                Else
                    If ASCMAIN1.CLIENT = "AHA" Then
                        tlb_btn.SharedProps.Visible = grd.ActiveRow.Cells("CUST_CODE").Value & String.Empty = "CONSUMER"

                    ElseIf ASCMAIN1.CLIENT = "INT" Then
                        tlb_btn.SharedProps.Visible = True OrElse ((grd.ActiveRow.Cells("CUST_CODE").Value & String.Empty <> grd.ActiveRow.Cells("RA_CUST_CODE").Value & String.Empty) _
                                                        AndAlso grd.ActiveRow.Cells("RA_CUST_CODE").Value & String.Empty <> String.Empty)
                    Else
                        tlb_btn.SharedProps.Visible = False
                    End If

                    tlb_btn = DirectCast(tlb_pop.Tools("Change Ship To"), UltraWinToolbars.ButtonTool)

                    If ASCMAIN1.CLIENT = "AHA" Then
                        tlb_btn.SharedProps.Visible = False 'grd.ActiveRow.Cells("CUST_CODE").Value & String.Empty = "CONSUMER"

                    ElseIf ASCMAIN1.CLIENT = "INT" Then
                        tlb_btn.SharedProps.Visible = ((grd.ActiveRow.Cells("CUST_CODE").Value & String.Empty = grd.ActiveRow.Cells("RA_CUST_CODE").Value & String.Empty) _
                                                        AndAlso grd.ActiveRow.Cells("RA_CUST_CODE").Value & String.Empty <> String.Empty)
                    Else
                        tlb_btn.SharedProps.Visible = False
                    End If

                    If 1 = 2 Then
                        tlb_btn = DirectCast(tlb_pop.Tools("Reship to AE"), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Visible = Not InquiryMode AndAlso ((grd.ActiveRow.Cells("CUST_CODE").Value & "") = "IPLBAE")
                    End If

                End If

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Refresh"
                If grd.Name = grdEDTRTRN1.Name Then
                    EnforceConstraints(False)
                    Setup_3PL()
                    EnforceConstraints(True)
                End If

            Case "3PL Returns Report"
                Print_Report_Begin()
                Generate_Report("WHR3PLRR", "3PL Returns Report", , , , , False)
                Print_Report_End()

            Case "Copy All Lines to Negate Inventory Impact"
                Dim TBL As DataTable = dst.Tables("SOTRTRN2").Clone
                Dim RTRN_LNO As Int32 = Val(dst.Tables("SOTRTRN2").Compute("MAX (RTRN_LNO)", "") & "")
                For Each row As DataRow In dst.Tables("SOTRTRN2").Select("", "RTRN_LNO")
                    TBL.Rows.Add(row.ItemArray)
                Next
                For Each row As DataRow In TBL.Select("", "RTRN_LNO")
                    RTRN_LNO += 1
                    Dim rowSOTRTRN2 As DataRow = dst.Tables("SOTRTRN2").NewRow
                    rowSOTRTRN2.ItemArray = row.ItemArray
                    rowSOTRTRN2.Item("RTRN_LNO") = RTRN_LNO
                    rowSOTRTRN2.Item("RTRN_QTY") = -1 * Val(rowSOTRTRN2.Item("RTRN_QTY") & "")
                    rowSOTRTRN2.Item("RTRN_QTY_1") = -1 * Val(rowSOTRTRN2.Item("RTRN_QTY_1") & "")
                    rowSOTRTRN2.Item("RTRN_QTY_2") = -1 * Val(rowSOTRTRN2.Item("RTRN_QTY_2") & "")
                    rowSOTRTRN2.Item("RTRN_QTY_3") = -1 * Val(rowSOTRTRN2.Item("RTRN_QTY_3") & "")
                    rowSOTRTRN2.Item("RTRN_QTY_4") = -1 * Val(rowSOTRTRN2.Item("RTRN_QTY_4") & "")
                    dst.Tables("SOTRTRN2").Rows.Add(rowSOTRTRN2)
                Next

                DisplayTotals()
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Change Ship To"

                Select Case ASCMAIN1.CLIENT
                    Case "INT"
                        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("CUST_STORE_NO", "ARTCUST2",
                                                          "CUST_CODE = '" & grd.ActiveRow.Cells("RA_CUST_CODE").Value & String.Empty & "'" _
                                                          & " and CUST_STORE_NO_3PL IS NOT NULL and NVL(CUST_STORE_STATUS, 'A') = 'A'")

                        If ASCMAIN1.CodeSelector.SQL <> "" Then
                            ASCMAIN1.CodeSelector.MultipleSelections = False
                            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                            ' Show the Customer Code
                            ASCMAIN1.CodeSelector.grdColumns(0).Item("COLUMN_HIDDEN") = "0"
                            Using F As New ASFCODE1
                                F.ShowDialog()
                            End Using

                            Select Case ASCMAIN1.CodeSelector.SelectedRows.Count
                                Case 0
                                    ' Nothing to do
                                Case 1
                                    Dim CUST_CODE As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("CUST_CODE") & String.Empty
                                    Dim CUST_STORE_NO As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("CUST_STORE_NO") & String.Empty

                                    Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
                                    If rowARTCUST2 Is Nothing Then
                                        MessageBox.Show("Could not locate the selected Store", "Change Customer", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                        Exit Sub
                                    End If

                                    Dim CUST_STORE_NO_3PL As String = rowARTCUST2.Item("CUST_STORE_NO_3PL") & String.Empty
                                    Dim CUST_NO_3PL As String = rowARTCUST2.Item("CUST_NO_3PL") & String.Empty
                                    'grd.ActiveRow.Cells("CUST_CODE").Value = CUST_CODE
                                    'grd.ActiveRow.Cells("CUST_NAME").Value = rowARTCUST2.Item("CUST_STORE_NAME") & String.Empty
                                    'grd.ActiveRow.Cells("EDI_CUSTOMER_NO").Value = CUST_NO_3PL
                                    grd.ActiveRow.Cells("EDI_CUST_SHIP_TO").Value = CUST_STORE_NO_3PL

                                    ASCMAIN1.sql = "UPDATE EDTRTRN1 SET EDI_CUST_SHIP_TO = '" & CUST_STORE_NO_3PL & "'" _
                                        & " WHERE EDI_DOC_SEQ_NO = '" & grd.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & String.Empty & "'"
                                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                                Case Else
                                    MessageBox.Show("Multiple Cross References selected for Customer on Return.")
                            End Select
                        End If


                    Case Else
                        Exit Sub
                End Select


            Case "Change Customer"
                ChangeCustomer(grd)

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Credit Entire Invoice"
                INV_NO_RETURNED = grd.ActiveRow.Cells("INV_NO").Value
                Absx1.txtFor("CUST_CODE").Text = grd.ActiveRow.Cells("CUST_CODE").Value
                Absx1.txtFor("WHSE_CODE").Text = grd.ActiveRow.Cells("WHSE_CODE").Value
                Click_Command("New")
                If Not ScreenMode Then
                    INV_NO_RETURNED = ""
                Else
                    'Dim rowSOTINVH1 As DataRow = LookUp("SOTINVH1", New String() {"U", INV_NO_RETURNED})
                    'Stop

                    'Absx1.txtFor("CUST_STORE_NO").Text = grd.ActiveRow.Cells("CUST_STORE_NO").Value
                    'Absx1.txtFor("CUST_CLAIM_NO").Text = grd.ActiveRow.Cells("ORDR_CUST_PO").Value
                End If

            Case "Copy Price to All Lines"
                Dim RTRN_PRICE As Decimal = Val(grd.ActiveRow.Cells("RTRN_PRICE").Value & "")
                If MsgBox("OK to copy price " & Format(RTRN_PRICE, "$#.00") & " to all lines?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Copying Price")

                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    grow.Cells("RTRN_PRICE").Value = RTRN_PRICE
                    grow.Update()
                Next

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

            Case "Remove From List"
                If grdEDTRTRN1.ActiveRow Is Nothing Then
                    MessageBox.Show("You must select an entry in the " & grdEDTRTRN1.Text & " grid.", e.Tool.Key, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim EDI_DOC_SEQ_NO As String = grdEDTRTRN1.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value
                Dim rowEDTRTRN1 As DataRow = dst.Tables("EDTRTRN1").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")(0)
                Dim RA_NO As String = rowEDTRTRN1.Item("EDI_RA_NO") & String.Empty
                Dim CUST_CODE As String = rowEDTRTRN1.Item("CUST_CODE") & String.Empty
                Dim RA_NO_NEW As String = String.Empty

                If MessageBox.Show("Do you want to remove the selected Return from the list?", "Remove From List", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                If Not ASCMAIN1.Logical_Lock("SOTRMAF1", "E" & EDI_DOC_SEQ_NO) Then
                    Exit Sub
                End If

                Try
                    BeginTrans()
                    ASCMAIN1.sql = "Update EDTRTRN1 set PROCESS_IND = 'D' WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                    CommitTrans("Return Removed from List")

                    EnforceConstraints(False)
                    Setup_3PL()
                    EnforceConstraints(True)

                Catch ex As Exception
                    Rollback(ex.Message)
                Finally
                    ASCMAIN1.MultiTask_Release()
                End Try

            Case "Modify RA Number", "Set RA Number"

                If grdEDTRTRN1.Selected.Rows.Count = 0 Then
                    MessageBox.Show("You must select at least one Return in the " & grdEDTRTRN1.Text & " grid.", e.Tool.Key, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim EDI_DOC_SEQ_NO As String = grdEDTRTRN1.Selected.Rows(0).Cells("EDI_DOC_SEQ_NO").Value
                Dim rowEDTRTRN1 As DataRow = dst.Tables("EDTRTRN1").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")(0)
                'Dim RA_NO As String = rowEDTRTRN1.Item("EDI_RA_NO") & String.Empty
                Dim CUST_CODE As String = rowEDTRTRN1.Item("CUST_CODE") & String.Empty
                Dim RA_NO_NEW As String = String.Empty
                Dim lstEDI_DOC_SEQ_NO As New List(Of String)

                If e.Tool.Key = "Set RA Number" Then
                    If grdSOT3PLF1.ActiveRow Is Nothing Then
                        MessageBox.Show("You must select one RA from the Open RAs list.", "Set RA Number", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If
                    RA_NO_NEW = grdSOT3PLF1.ActiveRow.Cells("RA_NO").Value
                Else
                    ' Verify all selected Rows have the same customer.
                    For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdEDTRTRN1.Selected.Rows
                        If grdRow.Cells("CUST_CODE").Value & String.Empty <> CUST_CODE Then
                            MessageBox.Show("All the selected Returns must have the same Customer.", "Modify RA Number", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End If
                    Next

                    If MessageBox.Show("Do you want to change RA Number for the (" & grdEDTRTRN1.Selected.Rows.Count & ") selected Returns?", "Modify RA Number", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If

                    'RA_NO_NEW = InputBox("Enter the New RA Number.", "Modify RA Number", String.Empty)

                    ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("RA_NO", "SOTRMAF1", "NVL(RA_STATUS, 'O') = 'O'")
                    If ASCMAIN1.CodeSelector.SQL <> "" Then
                        ASCMAIN1.CodeSelector.MultipleSelections = False
                        ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                        Using F As New ASFCODE1
                            F.ShowDialog()
                        End Using

                        If ASCMAIN1.CodeSelector.SelectedRows.Count <> 1 Then
                            MessageBox.Show("Multiple RA Nos Selected for the Returns.", "Modify RA Number", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End If

                        Dim row As DataRow = ASCMAIN1.CodeSelector.SelectedRows(0)
                        RA_NO_NEW = row.Item("RA_NO")

                    End If
                End If

                If RA_NO_NEW.Length = 0 Then
                    Exit Sub
                End If

                RA_NO_NEW = ASCMAIN1.Format_Field(RA_NO_NEW, "RA_NO")

                ' Validate the New RA_NO
                Dim rowSOTRMAF1 As DataRow = ASCDATA1.GetDataRow("Select * From SOTRMAF1 where RA_NO = '" & RA_NO_NEW & "'")

                If rowSOTRMAF1 Is Nothing Then
                    MessageBox.Show("The provided RA No (" & RA_NO_NEW & ") is an invalid value.", "Modify RA Number", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                ElseIf rowSOTRMAF1.Item("RA_STATUS") <> "O" Then
                    MessageBox.Show("The provided RA No (" & RA_NO_NEW & ") does not have an 'Open' status.", "Modify RA Number", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim changeOfCustomer As Boolean = False
                If CUST_CODE.Length > 0 AndAlso rowSOTRMAF1.Item("CUST_CODE") & String.Empty <> CUST_CODE Then
                    'MessageBox.Show("The provided RA No (" & RA_NO_NEW & ") is for " & rowSOTRMAF1.Item("CUST_CODE") & "; however, the selected RA is for " & CUST_CODE & ".", "Modify RA Number", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    'Exit Sub
                    changeOfCustomer = True
                End If

                If MessageBox.Show($"Do you want to change RA Number {RA_NO} to {RA_NO_NEW}, Customer {rowSOTRMAF1.Item("CUST_CODE")}?", "Modify RA Number", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                ' Verify all selected Rows have the same customer.
                For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdEDTRTRN1.Selected.Rows
                    EDI_DOC_SEQ_NO = grdRow.Cells("EDI_DOC_SEQ_NO").Value
                    If Not ASCMAIN1.Logical_Lock("SOTRMAF1", "E" & EDI_DOC_SEQ_NO) Then
                        Exit Sub
                    End If
                    lstEDI_DOC_SEQ_NO.Add(EDI_DOC_SEQ_NO)
                Next

                Try
                    EnforceConstraints(False)

                    For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdEDTRTRN1.Selected.Rows
                        EDI_DOC_SEQ_NO = grdRow.Cells("EDI_DOC_SEQ_NO").Value
                        rowEDTRTRN1 = dst.Tables("EDTRTRN1").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")(0)
                        rowEDTRTRN1.Item("EDI_RA_NO") = RA_NO_NEW
                        rowEDTRTRN1.Item("CUST_CODE") = rowSOTRMAF1.Item("CUST_CODE")
                        rowEDTRTRN1.Item("RA_CUST_CODE") = rowSOTRMAF1.Item("CUST_CODE")
                        ASCDATA1.ExecuteSQL($"UPDATE EDTRTRN1 SET EDI_RA_NO = '{RA_NO_NEW}', CUST_CODE = '{rowSOTRMAF1.Item("CUST_CODE")}' WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
                    Next
                    dst.Tables("EDTRTRN1").AcceptChanges()
                    EnforceConstraints(True)
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                Finally
                    ASCMAIN1.MultiTask_Release()
                End Try

                If changeOfCustomer Then
                    MessageBox.Show("The Customer has changed; therefore, you must select a valid Ship To for the RA.", e.Tool.Key, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ChangeCustomer(grd)
                End If

            Case "Reship to AE"
                If grdEDTRTRN1.ActiveRow Is Nothing OrElse grdEDTRTRN1.ActiveRow.Band.Index <> 0 Then Exit Sub

                Dim CUST_CODE As String = grdEDTRTRN1.ActiveRow.Cells("CUST_CODE").Value & ""
                If CUST_CODE <> "IPLBAE" Then Exit Sub

                Dim EDI_DOC_SEQ_NO As String = grdEDTRTRN1.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & ""

                If MessageBox.Show("Create a reship Sales Order to AE for this return?",
                                   "Reship to AE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                    Exit Sub
                End If

                RESHIP_TO_AE(EDI_DOC_SEQ_NO)
        End Select
    End Sub
    Private Sub RESHIP_TO_AE(EDI_DOC_SEQ_NO As String)
        'TODO
        'GO INTO EDTRTRN1/EDTRTRN2
        'GET ORIGINAL PO NUMBER
        'FIND ORIGINAL SHIP TO FOR THAT SO/PO AND USE THAT AE SHIP-TO
        'CREATE NEW SO UNDER THE SAME GROUP AS ORIGINAL PO
        If grdEDTRTRN1.ActiveRow Is Nothing Then Exit Sub
        If grdEDTRTRN1.ActiveRow.Band.Index <> 0 Then Exit Sub

        Dim CUST_CODE As String = grdEDTRTRN1.ActiveRow.Cells("CUST_CODE").Value & ""
        If CUST_CODE <> "IPLBAE" Then
            MessageBox.Show("Reship to AE is only available for IPLBAE.", "Reship to AE", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim TRACKING_NO As String = ""
        If grdEDTRTRN1.ActiveRow.Cells.Exists("TRACKING_NO") Then
            TRACKING_NO = grdEDTRTRN1.ActiveRow.Cells("TRACKING_NO").Value & ""
        End If

        TRACKING_NO = TRACKING_NO.Trim().Replace(" ", "")

        If TRACKING_NO = "" Then
            MessageBox.Show("Tracking # is blank on this return row.", "Reship to AE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ASCMAIN1.sql =
        "SELECT SOTCSTO1.CSO_NO, SOTCSTO1.SELL_CODE, " & vbCrLf &
        "       SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_DATE_SHIPPED, " & vbCrLf &
        "       SOTORDR1.ORDR_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.CUST_CODE, SOTORDR1.ORDR_GROUP_NO," & vbCrLf &
        "       SOTCART1.CART_NO, SOTCART1.CART_TRACKING_NO," & vbCrLf &
        "       SOTPICK1.PICK_NO, SOTPICK1.SHIP_BOL_NO, SOTPICK1.INV_NO" & vbCrLf &
        "  FROM SOTPICK1, SOTCART1, SOTORDR1, SOTCSTO1" & vbCrLf &
        " WHERE SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf &
        "   AND SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf &
        "   AND SOTCSTO1.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO" & vbCrLf &
        "   AND SOTORDR1.ORDR_DATE_SHIPPED >= TRUNC(SYSDATE) - 120" & vbCrLf &
        "   AND REPLACE(SOTCART1.CART_TRACKING_NO, ' ', '') = '" & TRACKING_NO.Replace("'", "''") & "'" & vbCrLf &
        " ORDER BY SOTORDR1.ORDR_DATE_SHIPPED DESC, SOTORDR1.ORDR_NO DESC"

        Dim dtRef As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
        If dtRef.Rows.Count = 0 Then
            MessageBox.Show("Could not find an original order for Tracking #: " & TRACKING_NO, "Reship to AE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If dtRef.Rows.Count > 1 Then
            Dim msg As New System.Text.StringBuilder()
            msg.AppendLine("Multiple shipped orders were found for Tracking #: " & TRACKING_NO)
            msg.AppendLine("The most recently shipped one will be used unless you cancel.")
            msg.AppendLine("")

            For Each r As DataRow In dtRef.Rows
                Dim shipped As String = ""
                If Not IsDBNull(r("ORDR_DATE_SHIPPED")) Then
                    shipped = CDate(r("ORDR_DATE_SHIPPED")).ToString("MM/dd/yyyy")
                End If

                msg.AppendLine("ORDR_NO=" & (r("ORDR_NO") & "") &
                       ", Shipped=" & shipped &
                       ", INV_NO=" & (r("INV_NO") & "") &
                       ", PICK_NO=" & (r("PICK_NO") & ""))
            Next

            Dim resp = MessageBox.Show(msg.ToString(), "Reship to AE",
                               MessageBoxButtons.OKCancel, MessageBoxIcon.Warning)

            If resp = DialogResult.Cancel Then Exit Sub
        End If

        Dim rowRef As DataRow = dtRef.Rows(0)
        Dim ORDR_NO_ORIG As String = rowRef("ORDR_NO") & ""
        Dim ORDR_GROUP_NO_ORIG As String = rowRef("ORDR_GROUP_NO") & ""
        Dim ORDR_CUST_PO_ORIG As String = rowRef("ORDR_CUST_PO") & ""
        Dim SELL_CODE As String = (rowRef("SELL_CODE") & "").ToString().Trim()
        If SELL_CODE = "" Then
            MessageBox.Show("SELL_CODE is blank from tracking lookup (SOTCSTO1).", "Reship to AE",
                    MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim CUST_STORE_NO As String = SELL_CODE.PadLeft(6, "0"c)

        ASCMAIN1.sql =
    "SELECT * FROM ARTCUST2 " &
    " WHERE CUST_CODE = 'IPLBAE' " &
    "   AND CUST_STORE_NO = '" & CUST_STORE_NO.Replace("'", "''") & "'"

        Dim rowARTCUST2_AE As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)
        If rowARTCUST2_AE Is Nothing Then
            MessageBox.Show("Could not find AE store in ARTCUST2 for IPLBAE / " & CUST_STORE_NO, "Reship to AE",
                    MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If ORDR_NO_ORIG.Trim() = "" OrElse ORDR_GROUP_NO_ORIG.Trim() = "" Then
            MessageBox.Show("Tracking lookup did not return ORDR_NO / ORDR_GROUP_NO.", "Reship to AE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ASCMAIN1.sql = "SELECT * FROM SOTORDR0 WHERE ORDR_GROUP_NO = '" & ORDR_GROUP_NO_ORIG.Replace("'", "''") & "'"
        Fill_Records("SOTORDR0", , , ASCMAIN1.sql)
        Dim rowSOTORDR0_orig As DataRow = dst.Tables("SOTORDR0").Rows.Find(ORDR_GROUP_NO_ORIG)
        If rowSOTORDR0_orig Is Nothing Then
            MessageBox.Show("Could not load SOTORDR0 for ORDR_GROUP_NO " & ORDR_GROUP_NO_ORIG, "Reship to AE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ASCMAIN1.sql = "SELECT * FROM SOTORDR1 WHERE ORDR_NO = '" & ORDR_NO_ORIG.Replace("'", "''") & "'"
        Fill_Records("SOTORDR1", , , ASCMAIN1.sql)
        Dim rowSOTORDR1_orig As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO_ORIG)
        If rowSOTORDR1_orig Is Nothing Then
            MessageBox.Show("Could not load SOTORDR1 for ORDR_NO " & ORDR_NO_ORIG, "Reship to AE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ASCMAIN1.sql = "SELECT * FROM SOTORDR2 WHERE ORDR_NO = '" & ORDR_NO_ORIG.Replace("'", "''") & "'"
        Fill_Records("SOTORDR2", , , ASCMAIN1.sql)
        Dim rowsSOTORDR2_orig() As DataRow = dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO_ORIG.Replace("'", "''") & "'")
        If rowsSOTORDR2_orig Is Nothing OrElse rowsSOTORDR2_orig.Length = 0 Then
            MessageBox.Show("Original order has no SOTORDR2 lines.", "Reship to AE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ASCMAIN1.sql = "SELECT * FROM SOTORDR5 WHERE ORDR_NO = '" & ORDR_NO_ORIG.Replace("'", "''") & "'"
        Fill_Records("SOTORDR5", , , ASCMAIN1.sql)
        Dim rowsSOTORDR5_orig() As DataRow = dst.Tables("SOTORDR5").Select("ORDR_NO = '" & ORDR_NO_ORIG.Replace("'", "''") & "'")

        Dim rowsEDTRTRN2() As DataRow = dst.Tables("EDTRTRN2").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO.Replace("'", "''") & "'")
        If rowsEDTRTRN2 Is Nothing OrElse rowsEDTRTRN2.Length = 0 Then
            MessageBox.Show("No EDTRTRN2 rows found for EDI_DOC_SEQ_NO " & EDI_DOC_SEQ_NO, "Reship to AE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim sqlDup As String =
    "SELECT ORDR_NO FROM SOTORDR1 " &
    " WHERE CUST_CODE = 'IPLBAE' " &
    "   AND ORDR_INTERNAL_NOTES LIKE '%Reship to AE:%' " &
    "   AND ORDR_INTERNAL_NOTES LIKE '%EDI_DOC_SEQ_NO=" & EDI_DOC_SEQ_NO.Replace("'", "''") & "%' " &
    "   AND ORDR_INTERNAL_NOTES LIKE '%OrigOrd=%' " &
    " ORDER BY ORDR_NO DESC"


        Dim rowDup As DataRow = ASCDATA1.GetDataRow(sqlDup)
        If rowDup IsNot Nothing Then
            MessageBox.Show("A Reship to AE order already exists for this return." & vbCrLf &
                    "Sales Order: " & (rowDup("ORDR_NO") & ""), "Reship to AE",
                    MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim ORDR_NO_NEW As String = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")

        'SOTORDR1
        Dim rowSOTORDR1_new As DataRow = dst.Tables("SOTORDR1").NewRow()
        rowSOTORDR1_new.ItemArray = rowSOTORDR1_orig.ItemArray

        rowSOTORDR1_new("ORDR_NO") = ORDR_NO_NEW 'NEW ORDR NO
        rowSOTORDR1_new("ORDR_GROUP_NO") = ORDR_GROUP_NO_ORIG  'SAME ORDER GROUP
        rowSOTORDR1_new("ORDR_CUST_PO") = ORDR_CUST_PO_ORIG  'SAME PO

        rowSOTORDR1_new("ORDR_DATE") = Now
        rowSOTORDR1_new("ORDR_DATE_RECD") = Now
        rowSOTORDR1_new("ORDR_DATE_SHIPPED") = DBNull.Value
        rowSOTORDR1_new("ORDR_DATE_CLOSED") = DBNull.Value
        rowSOTORDR1_new("ORDR_YYYYPP_CLOSED") = DBNull.Value
        rowSOTORDR1_new("ORDR_INVOICED") = "0"
        rowSOTORDR1_new("ORDR_BATCHED") = "0"
        rowSOTORDR1_new("ORDR_PICK_SEQ") = 0

        rowSOTORDR1_new("INIT_OPER") = ASCMAIN1.USER_ID
        rowSOTORDR1_new("LAST_OPER") = ASCMAIN1.USER_ID
        rowSOTORDR1_new("INIT_DATE") = Now
        rowSOTORDR1_new("LAST_DATE") = Now

        rowSOTORDR1_new("ORDR_REL_BATCH_NO") = DBNull.Value
        rowSOTORDR1_new("ORDR_DATE_REL") = DBNull.Value
        rowSOTORDR1_new("ORDR_STATUS") = "O"
        rowSOTORDR1_new("ORDR_DATE_BOOKED") = Now

        rowSOTORDR1_new("CUST_STORE_LOCATION") = (rowARTCUST2_AE("CUST_STORE_NAME") & "").ToString().Trim()

        rowSOTORDR1_new("ORDR_OVERRIDE_NOT_ALLOCATED") = "1"

        'rowSOTORDR1_new("ORDR_STATUS") = "0"
        'rowSOTORDR1_new("ORDR_INVOICED") = "0"
        'rowSOTORDR1_new("ORDR_DATE_SHIPPED") = DBNull.Value
        'rowSOTORDR1_new("ORDR_DATE_CLOSED") = DBNull.Value
        'rowSOTORDR1_new("ORDR_YYYYPP_CLOSED") = DBNull.Value

        Dim addNote As String = $"Reship to AE: EDI_DOC_SEQ_NO={EDI_DOC_SEQ_NO}, Tracking={TRACKING_NO}, OrigOrd={ORDR_NO_ORIG}"
        rowSOTORDR1_new("ORDR_INTERNAL_NOTES") = addNote

        dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1_new)

        'SOTORDR2
        Dim nextLno As Integer = 1
        Dim createdLines As Integer = 0
        Dim warnMsg As New System.Text.StringBuilder()

        For Each rowRet2 As DataRow In rowsEDTRTRN2

            Dim ITEM_CODE As String = (rowRet2("EDI_ITEM_CODE") & "").ToString().Trim()
            Dim QTY_RETURNED As Integer = Val(rowRet2("EDI_QTY_RETURNED") & "")

            If ITEM_CODE = "" OrElse QTY_RETURNED <= 0 Then Continue For

            Dim rowLineOrig As DataRow = Nothing
            For Each r As DataRow In rowsSOTORDR2_orig
                If (r("ITEM_CODE") & "").ToString().Trim() = ITEM_CODE Then
                    rowLineOrig = r
                    Exit For
                End If
            Next

            'Item on return but not on original shipped order
            If rowLineOrig Is Nothing Then
                warnMsg.AppendLine("Item " & ITEM_CODE & " was found on the return, but was not found on the original order. It was not included on the reship order.")
                Continue For
            End If

            Dim QTY_SHIPPED_ORIG As Integer = Val(rowLineOrig("ORDR_QTY_SHIP") & "")
            Dim QTY_TO_USE As Integer = QTY_RETURNED

            'Compare return qty to original shipped qty
            If QTY_RETURNED > QTY_SHIPPED_ORIG Then
                warnMsg.AppendLine("Item " & ITEM_CODE & " has returned qty " & QTY_RETURNED &
                           " which is greater than original shipped qty " & QTY_SHIPPED_ORIG &
                           ". The reship order will use shipped qty " & QTY_SHIPPED_ORIG & ".")
                QTY_TO_USE = QTY_SHIPPED_ORIG
            Else
                ' If returned < shipped, use returned qty
                QTY_TO_USE = QTY_RETURNED
            End If

            If QTY_TO_USE <= 0 Then
                warnMsg.AppendLine("Item " & ITEM_CODE & " has no valid quantity to reship and was skipped.")
                Continue For
            End If

            Dim rowSOTORDR2_new As DataRow = dst.Tables("SOTORDR2").NewRow()
            rowSOTORDR2_new.ItemArray = rowLineOrig.ItemArray

            rowSOTORDR2_new("ORDR_NO") = ORDR_NO_NEW
            rowSOTORDR2_new("ORDR_LNO") = nextLno : nextLno += 1

            rowSOTORDR2_new("ORDR_QTY") = QTY_TO_USE
            rowSOTORDR2_new("ORDR_QTY_OPEN") = QTY_TO_USE
            rowSOTORDR2_new("ORDR_QTY_PICK") = 0
            rowSOTORDR2_new("ORDR_QTY_SHIP") = 0
            rowSOTORDR2_new("ORDR_QTY_CANC") = 0
            rowSOTORDR2_new("ORDR_QTY_ORIG") = QTY_TO_USE
            rowSOTORDR2_new("ORDR_QTY_BACK") = 0

            rowSOTORDR2_new("ORDR_STATUS") = "O"
            'rowSOTORDR2_new("ORDR_RELEASE") = "0"

            dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2_new)
            createdLines += 1
        Next

        For Each rowOrig As DataRow In rowsSOTORDR2_orig
            Dim ITEM_CODE_ORIG As String = (rowOrig("ITEM_CODE") & "").ToString().Trim()
            Dim QTY_SHIPPED_ORIG As Integer = Val(rowOrig("ORDR_QTY_SHIP") & "")

            If ITEM_CODE_ORIG = "" OrElse QTY_SHIPPED_ORIG <= 0 Then Continue For

            Dim foundOnReturn As Boolean = False
            For Each rowRet2 As DataRow In rowsEDTRTRN2
                If (rowRet2("EDI_ITEM_CODE") & "").ToString().Trim() = ITEM_CODE_ORIG Then
                    foundOnReturn = True
                    Exit For
                End If
            Next

            If Not foundOnReturn Then
                warnMsg.AppendLine("Item " & ITEM_CODE_ORIG & " was shipped on the original order but was not found on the return. It was not included on the reship order.")
            End If
        Next

        If createdLines = 0 Then
            MessageBox.Show("No reship lines were created (no matching items vs original shipped order)." &
                    If(warnMsg.Length > 0, vbCrLf & vbCrLf & warnMsg.ToString(), ""),
                    "Reship to AE",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
            Exit Sub
        End If

        If warnMsg.Length > 0 Then
            MessageBox.Show(warnMsg.ToString(), "Reship to AE Warnings",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

        'SOTORDR5
        Dim rowBT_orig As DataRow = Nothing
        For Each r As DataRow In rowsSOTORDR5_orig
            If (r("CUST_ADDR_TYPE") & "").ToString().Trim() = "BT" Then
                rowBT_orig = r
                Exit For
            End If
        Next

        If rowBT_orig IsNot Nothing Then
            Dim rBT_new As DataRow = dst.Tables("SOTORDR5").NewRow()
            rBT_new.ItemArray = rowBT_orig.ItemArray
            rBT_new("ORDR_NO") = ORDR_NO_NEW
            rBT_new("CUST_ADDR_TYPE") = "BT"
            dst.Tables("SOTORDR5").Rows.Add(rBT_new)
        End If

        Dim rST_new As DataRow = dst.Tables("SOTORDR5").NewRow()
        Dim rowST_orig As DataRow = Nothing
        For Each r As DataRow In rowsSOTORDR5_orig
            If (r("CUST_ADDR_TYPE") & "").ToString().Trim() = "ST" Then
                rowST_orig = r
                Exit For
            End If
        Next
        If rowST_orig IsNot Nothing Then
            rST_new.ItemArray = rowST_orig.ItemArray
        End If

        rST_new("ORDR_NO") = ORDR_NO_NEW
        rST_new("CUST_ADDR_TYPE") = "ST"

        rST_new("CUST_NAME") = (rowARTCUST2_AE("CUST_STORE_NAME") & "").ToString().Trim()
        rST_new("CUST_ADDR1") = (rowARTCUST2_AE("CUST_STORE_ADDR1") & "").ToString().Trim()
        rST_new("CUST_ADDR2") = (rowARTCUST2_AE("CUST_STORE_ADDR2") & "").ToString().Trim()
        rST_new("CUST_CITY") = (rowARTCUST2_AE("CUST_STORE_CITY") & "").ToString().Trim()
        rST_new("CUST_STATE") = (rowARTCUST2_AE("CUST_STORE_STATE") & "").ToString().Trim()
        rST_new("CUST_ZIP_CODE") = (rowARTCUST2_AE("CUST_STORE_ZIP_CODE") & "").ToString().Trim()
        rST_new("CUST_PHONE") = (rowARTCUST2_AE("CUST_STORE_PHONE") & "").ToString().Trim()

        dst.Tables("SOTORDR5").Rows.Add(rST_new)


        Dim inTransaction As Boolean = False
        Try
            Dim evtOld As String =
                $"Reship to AE created:  New ORDR_NO {ORDR_NO_NEW} is a reship for ORDR_NO {ORDR_NO_ORIG}. "

            Dim evtNew As String =
                $"Reship to AE order created from ORDR_NO {ORDR_NO_ORIG}. "

            BeginTrans()
            inTransaction = True

            Update_Record_TDA("SOTORDR1")
            Update_Record_TDA("SOTORDR2")
            Update_Record_TDA("SOTORDR5")

            ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO_ORIG}, New String() {"ORDR_GROUP_NO_IN"})

            CommitTrans("Reship Sales Order Created: " & ORDR_NO_NEW)
            inTransaction = False

            Try
                'Write_Event_Log("SOTORDR1", ORDR_NO_ORIG, evtOld, "RESHIP")
                'Write_Event_Log("SOTORDR1", ORDR_NO_NEW, evtNew, "RESHIP")
                TAC.TACMAIN1.Record_Event("SOTORDR1", ORDR_NO_ORIG, Now, ASCMAIN1.USER_ID, "RESHIP", evtOld, ORDR_NO_NEW, Me.Name)
                TAC.TACMAIN1.Record_Event("SOTORDR1", ORDR_NO_NEW, Now, ASCMAIN1.USER_ID, "RESHIP", evtNew, ORDR_NO_ORIG, Me.Name)
            Catch ex As Exception
                MessageBox.Show("Reship order was created, but event log was Not written." & vbCrLf &
                                ex.Message, "Reship to AE",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try

            MessageBox.Show("Created Reship Sales Order: " & ORDR_NO_NEW, "Reship to AE", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            If inTransaction Then Rollback(ex.Message)
            MessageBox.Show(ex.Message, "Reship to AE", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress("", "")
        End Try

        'MAKE SURE WE BYPASS ALLOCATION
        'IN SO ENTRY, MARISSA/NATHAN SHOULD BE ABLE TO EDIT THE ADDRESS
        'ANY OTHER CHANGES TO ALLOW ERIN TO RELEASE LIKE A BACKORDER
    End Sub

    Private Sub ChangeCustomer(ByRef grd As UltraWinGrid.UltraGrid)

        Select Case ASCMAIN1.CLIENT
            Case "AHA"
                If grd.ActiveRow.Cells("CUST_CODE").Value & String.Empty = "CONSUMER" Then
                    ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("CUST_CODE", "ARTCUST1", "CUST_CODE IN (SELECT CUST_CODE FROM WBTPART1 WHERE CUST_CODE <> 'CONSUMER')")
                    If ASCMAIN1.CodeSelector.SQL <> "" Then
                        ASCMAIN1.CodeSelector.MultipleSelections = False
                        ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                        ' Show the Customer Code
                        ASCMAIN1.CodeSelector.grdColumns(0).Item("COLUMN_HIDDEN") = "0"
                        Using F As New ASFCODE1
                            F.ShowDialog()
                        End Using

                        Select Case ASCMAIN1.CodeSelector.SelectedRows.Count
                            Case 0
                                        ' Nothing to do
                            Case 1
                                grdEDTRTRN1.ActiveRow.Cells("CUST_CODE").Value = ASCMAIN1.CodeSelector.SelectedRows(0).Item("CUST_CODE")
                            Case Else
                                MessageBox.Show("Multiple Cross References selected for Customer on Return.")
                        End Select
                    End If
                End If

            Case "INT"
                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("CUST_STORE_NO", "ARTCUST2",
                                                                          "CUST_CODE = '" & grd.ActiveRow.Cells("RA_CUST_CODE").Value & String.Empty & "'" _
                                                                          & " and CUST_STORE_NO_3PL IS NOT NULL and NVL(CUST_STORE_STATUS, 'A') = 'A'")
                If ASCMAIN1.CodeSelector.SQL <> "" Then
                    ASCMAIN1.CodeSelector.MultipleSelections = False
                    ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                    ' Show the Customer Code
                    ASCMAIN1.CodeSelector.grdColumns(0).Item("COLUMN_HIDDEN") = "0"
                    Using F As New ASFCODE1
                        F.ShowDialog()
                    End Using

                    Select Case ASCMAIN1.CodeSelector.SelectedRows.Count
                        Case 0
                                    ' Nothing to do
                        Case 1
                            Dim CUST_CODE As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("CUST_CODE") & String.Empty
                            Dim CUST_STORE_NO As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("CUST_STORE_NO") & String.Empty

                            Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
                            If rowARTCUST2 Is Nothing Then
                                MessageBox.Show("Could not locate the selected Store", "Change Customer", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                Exit Sub
                            End If

                            Dim CUST_STORE_NO_3PL As String = rowARTCUST2.Item("CUST_STORE_NO_3PL") & String.Empty
                            Dim CUST_NO_3PL As String = rowARTCUST2.Item("CUST_NO_3PL") & String.Empty
                            grd.ActiveRow.Cells("CUST_CODE").Value = CUST_CODE
                            grd.ActiveRow.Cells("CUST_NAME").Value = rowARTCUST2.Item("CUST_STORE_NAME") & String.Empty
                            grd.ActiveRow.Cells("EDI_CUSTOMER_NO").Value = CUST_NO_3PL
                            grd.ActiveRow.Cells("EDI_CUST_SHIP_TO").Value = CUST_STORE_NO_3PL

                            ASCMAIN1.sql = "UPDATE EDTRTRN1 SET CUST_CODE = '" & CUST_CODE & "'" _
                                & ", EDI_CUSTOMER_NO = '" & CUST_NO_3PL & "'" _
                                & ", EDI_CUST_SHIP_TO = '" & CUST_STORE_NO_3PL & "'" _
                                & " WHERE EDI_DOC_SEQ_NO = '" & grd.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & String.Empty & "'"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                        Case Else
                            MessageBox.Show("Multiple Cross References selected for Customer on Return.")
                    End Select
                End If

        End Select

    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)

            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode And Absx1.txtFor("WHSE_CODE").Text <> "" Then
                        ' Click_Command("New", e)
                    End If
                End If

            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode And Absx1.txtFor("CUST_CODE").Text <> "" Then
                        ' Click_Command("New", e)
                    End If
                End If

            Case "RTRN_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If

            Case "ITEM_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Me.ProcessEnterKeyStroke(Absx1.txtFor("ITEM_CODE").Text.Trim)
                    timItemCode.Start()
                End If

        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                If Not InquiryMode And Absx1.txtFor("WHSE_CODE").Text <> "" Then
                    ' Click_Command("New")
                End If
            Case "WHSE_CODE"
                If Not InquiryMode And Absx1.txtFor("CUST_CODE").Text <> "" Then
                    ' Click_Command("New")
                End If
            Case "RTRN_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

            Case "CUST_CODE"
                dst.Tables("SOTINVHH").Rows.Clear()
                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                If CUST_CODE <> "" Then
                    Dim row As DataRow = LookUp("ARTCUST1", CUST_CODE)
                    If row IsNot Nothing Then
                        Load_SOTINVHH()
                    End If
                End If

            Case "WHSE_CODE"
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "REASON_CODE"
                Dim REASON_CODE As String = Absx1.txtFor("REASON_CODE").Text
                chkNoImpact.Visible = False
                If REASON_CODE <> "" Then

                    Dim rowARTREAS1 As DataRow = LookUp("ARTREAS1", REASON_CODE)
                    If rowARTREAS1 IsNot Nothing Then
                        chkNoImpact.Visible = True
                        chkNoImpact.Checked = (rowARTREAS1.Item("RETURN_NO_STOCK_IND") & "" = "1")
                    End If
                End If
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        If Not ScreenMode Then Exit Sub

        Select Case Absx1.GetABSColumnName(sender)
            Case "RTRN_SALES", "RTRN_FREIGHT", "RTRN_HANDLING"
                Absx1.numFor("RTRN_AMOUNT").Value _
                    = Val(Absx1.numFor("RTRN_SALES").Value & "") _
                    + Val(Absx1.numFor("RTRN_FREIGHT").Value & "") _
                    + Val(Absx1.numFor("RTRN_HANDLING").Value & "")
        End Select
    End Sub

#End Region

#Region "grdSOTRTRN2"

    Private Sub grdSOTRTRN2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTRTRN2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"

                grdCodeDesc(grdSOTRTRN2, "ICTITEM1", "ITEM_CODE", "ITEM_DESC")
                If cdr IsNot Nothing Then
                    Dim ITEM_CODE As String = e.Cell.Value
                    Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                    e.Cell.Row.Cells("PROD_CODE").Value = cdr.Item("PROD_CODE")
                    Dim COST_CATGY_CODE As String = cdr.Item("COST_CATGY_CODE") & ""
                    Dim PROD_CODE As String = cdr.Item("PROD_CODE") & ""
                    Dim ITEM_COST_STD As Decimal = Val(cdr.Item("ITEM_COST_STD") & "")
                    e.Cell.Row.Cells("COST_CATGY_CODE").Value = COST_CATGY_CODE
                    e.Cell.Row.Cells("PROD_CODE").Value = PROD_CODE
                    e.Cell.Row.Cells("ITEM_COST_STD").Value = ITEM_COST_STD
                    'If location_support Then
                    '    e.Cell.Row.Cells("LOCATION_CODE").Value = ROWICTWHSE1.Item("WHSE_LOC_RTN")
                    '    ' USE ITEM_BIN AS A DEFAULT FOR AHA
                    'End If

                    If ScreenMode And Not IsLoading Then
                        Load_SOTINVHX(ITEM_CODE)
                    End If
                Else
                    grdSOTRTRN2.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If

            Case "RTRN_QTY"

        End Select
    End Sub

    Private Sub grdSOTRTRN2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSOTRTRN2.AfterExitEditMode
        Select Case grdSOTRTRN2.ActiveCell.Column.Key
            'Case "ACCT_CODE"
            '    Dim ACCT_CODE As String = grdICTIXFR2.ActiveCell.Text
            '    If ACCT_CODE <> "" Then
            '        grdICTIXFR2.ActiveCell.Value = ASCMAIN1.Format_Field(ACCT_CODE, grdGLTJRNL2.ActiveCell.Column.Key)
            '    End If
        End Select
    End Sub

    Private Sub grdSOTRTRN2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSOTRTRN2.AfterRowActivate
        With grdSOTRTRN2.DisplayLayout.Bands(0)
            If grdSOTRTRN2.ActiveRow.IsAddRow And Not whse_is_a_3PL Then
                .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdSOTRTRN2.ActiveCell = grdSOTRTRN2.ActiveRow.Cells("ITEM_CODE")
                grdSOTRTRN2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

        If EntryMode = "N" And Not grdSOTRTRN2.ActiveRow.IsAddRow Then
            Load_SOTINVHX(grdSOTRTRN2.ActiveRow.Cells("ITEM_CODE").Value)
        End If

        If EntryMode = "V" Then
            Show_GL()
        End If
    End Sub

    Private Sub grdSOTRTRN2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSOTRTRN2.AfterRowsDeleted
        DisplayTotals()
    End Sub

    Private Sub grdSOTRTRN2_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTRTRN2.AfterRowUpdate
        DisplayTotals()
    End Sub

    Private Sub grdSOTRTRN2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTRTRN2.BeforeExitEditMode
        If grdSOTRTRN2.ActiveCell Is Nothing Then Exit Sub
        With grdSOTRTRN2.ActiveCell
            Select Case .Column.Key
                Case "ITEM_CODE"
                    If .Text <> "" Then
                        If .Value IsNot Nothing Then
                            .Value = .Text.ToUpper
                        End If

                    End If
                    If .Text <> "" Then
                        cdr = LookUp("ICTITEM1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid Item Code (" & .Text & ")")
                            If .Value IsNot Nothing Then
                                .Value = ""
                            End If
                            e.Cancel = True
                        Else

                            If cdr.Item("ITEM_COST_STATUS") & "" <> "" Then
                                EMsg &= vbCr & "Item " & .Text & " does not have a Standard Cost"
                            End If

                        End If
                    End If

                Case "LOCATION_CODE"
                    If location_support Then
                        If .Text <> "" Then
                            If .Value IsNot Nothing Then
                                .Value = .Text.ToUpper
                            End If

                        End If
                        If .Text <> "" Then
                            cdr = LookUp("WHTLOCM1", New String() {Absx1.txtFor("WHSE_CODE").Text, .Text})
                            If cdr Is Nothing Then
                                ASCMAIN1.Progress("Invalid Location Code (" & .Text & ")")
                                If .Value IsNot Nothing Then
                                    .Value = ""
                                End If
                                e.Cancel = True
                            End If
                        End If
                    End If

            End Select
        End With
    End Sub

    Private Sub grdSOTRTRN2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTRTRN2.BeforeRowsDeleted

        'If whse_is_a_3PL Then
        '    RECORD_INDEXs = New List(Of Int32)
        '    For Each grow As UltraWinGrid.UltraGridRow In e.Rows
        '        RECORD_INDEXs.Add(grow.Cells("RECORD_INDEX").Value)
        '    Next
        'End If

    End Sub

    Private Sub grdSOTRTRN2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTRTRN2.BeforeRowUpdate
        With grdSOTRTRN2

            If e.Row.Cells("ITEM_CODE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("ICTITEM1", e.Row.Cells("ITEM_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Item Code (" & e.Row.Cells("ITEM_CODE").Text & ")",
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If location_support Then
                'If e.Row.Cells("BAR_CODE").Text = "" Then
                '    e.Cancel = True
                'Else
                '    LookUp("WHTBARC1", e.Row.Cells("BAR_CODE").Text)
                '    If cdr Is Nothing Then
                '        MsgBox("Invalid Value entered for Bar Code (" & e.Row.Cells("BAR_CODE").Text & ")", _
                '               MsgBoxStyle.OkOnly, "Cannot Update Row")
                '        e.Cancel = True
                '    End If
                'End If

                'If e.Row.Cells("LOCATION_CODE").Text = "" Then
                '    e.Cancel = True
                'Else
                '    LookUp("WHTLOCM1", New String() {Absx1.txtFor("WHSE_CODE").Text, e.Row.Cells("LOCATION_CODE").Text})
                '    If cdr Is Nothing Then
                '        MsgBox("Invalid Value entered for Location Code (" & e.Row.Cells("LOCATION_CODE").Text & ")", _
                '               MsgBoxStyle.OkOnly, "Cannot Update Row")
                '        e.Cancel = True
                '    End If
                'End If
            End If

            If Val(e.Row.Cells("RTRN_QTY_1").Value & String.Empty) < 0 _
                OrElse Val(e.Row.Cells("RTRN_QTY_2").Value & String.Empty) < 0 _
                OrElse Val(e.Row.Cells("RTRN_QTY_3").Value & String.Empty) < 0 _
                OrElse Val(e.Row.Cells("RTRN_QTY_4").Value & String.Empty) < 0 Then

                If ASCMAIN1.USER_ID = "wjz" And ASCMAIN1.Running_in_VS Then
                    Stop ' PERMIT REVERSAL OF A CREDIT FOR RETURNS
                Else
                    MsgBox("Invalid Value entered for Qty, qty may not be negative", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If Val(e.Row.Cells("RTRN_QTY").Value & "") = 0 Then
                MsgBox("Invalid Value entered for Qty (" & e.Row.Cells("RTRN_QTY").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            End If


            If Val(e.Row.Cells("RTRN_QTY").Value & "") <
                    (Val(e.Row.Cells("RTRN_QTY_1").Value & String.Empty) _
                    + Val(e.Row.Cells("RTRN_QTY_2").Value & String.Empty) _
                    + Val(e.Row.Cells("RTRN_QTY_3").Value & String.Empty) _
                    + Val(e.Row.Cells("RTRN_QTY_4").Value & String.Empty)) _
                AndAlso Val(e.Row.Cells("RTRN_PRICE").Text) > 0 Then
                MsgBox("Total quantity may not be more than Return quantity for items with a Non Zero price", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            End If


            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("RTRN_NO").Text = "" Then
                    '.ActiveRow.Cells("RTRN_NO").Value = RTRN_NO '  sotrtrn2RtrnNo ' Absx1.CtlFor("RTRN_NO").Text
                    '.ActiveRow.Cells("RTRN_LNO").Value = Val(dst.Tables("SOTRTRN2").Compute("Max(RTRN_LNO)", "RTRN_NO = '" & RTRN_NO & "'") & "") + 1

                    .ActiveRow.Cells("RTRN_NO").Value = sotrtrn2RtrnNo ' Absx1.CtlFor("RTRN_NO").Text
                    .ActiveRow.Cells("RTRN_LNO").Value = Val(dst.Tables("SOTRTRN2").Compute("Max(RTRN_LNO)", "RTRN_NO = '" & sotrtrn2RtrnNo & "'") & "") + 1

                End If
            End If
        End With
    End Sub

    Private Sub grdSOTRTRN2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTRTRN2.ClickCellButton

        If grdSOTRTRN2.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
            Case "LOCATION_CODE"
                sql_where = "WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'"
        End Select
        grdClickCellButton(grdSOTRTRN2, sql_where, False)

    End Sub

    Private Sub grdSOTRTRN2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdSOTRTRN2.Error
        grdSOTRTRN2.ActiveRow.CancelUpdate()
    End Sub

#End Region

    Sub DisplayTotals()
        Dim RTRN_SALES As Decimal = Val(dst.Tables("SOTRTRN2").Compute("SUM(LINE_SALES)", "") & "")
        Absx1.numFor("RTRN_SALES").Value = RTRN_SALES
        'dst.Tables("SOTRTRN1").Rows(0).Item("RTRN_SALES") = RTRN_SALES

        Dim RTRN_COSTS As Decimal = Val(dst.Tables("SOTRTRN2").Compute("SUM(LINE_COSTS)", "") & "")
        Absx1.numFor("RTRN_COSTS").Value = RTRN_COSTS
        'dst.Tables("SOTRTRN1").Rows(0).Item("RTRN_COSTS") = RTRN_COSTS

        For Each tableName As String In New String() {"SOTRTRN1", "SOTRTRNB"}
            For Each rowSOTRTRNX As DataRow In dst.Tables(tableName).Rows
                Dim RTRN_NO As String = rowSOTRTRNX.Item("RTRN_NO") & String.Empty
                RTRN_SALES = Val(dst.Tables("SOTRTRN2").Compute("SUM(LINE_SALES)", "RTRN_NO = '" & RTRN_NO & "'") & "")
                RTRN_COSTS = Val(dst.Tables("SOTRTRN2").Compute("SUM(LINE_COSTS)", "RTRN_NO = '" & RTRN_NO & "'") & "")
                dst.Tables(tableName).Select("RTRN_NO = '" & RTRN_NO & "'")(0).Item("RTRN_SALES") = RTRN_SALES
                Dim rowX() As DataRow = dst.Tables(tableName).Select("RTRN_NO = '" & RTRN_NO & "'")
                If rowX.Length <> 1 Then Stop
                dst.Tables(tableName).Select("RTRN_NO = '" & RTRN_NO & "'")(0).Item("RTRN_COSTS") = RTRN_COSTS
            Next
        Next
    End Sub

    Private Sub grdSOTRTRNX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTRTRNX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("RTRN_NO").Text = e.Row.Cells("RTRN_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub grdSOTRTRNG_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTRTRN3.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("RTRN_NO").Text = e.Row.Cells("RTRN_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub optGL_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optGL.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_GL()
    End Sub

    Sub Show_GL()

        If optGL.Tag <> optGL.Value Or optGL.Value = "L" Then
            optGL.Tag = optGL.Value
            If optGL.Value = "A" Then
                grdSOTRTRN3.DataSource = dst.Tables("SOTRTRN3")
                Dim dvw As DataView = dst.Tables("SOTRTRN3").DefaultView
                dvw.RowFilter = ""
            ElseIf optGL.Value = "L" Then
                grdSOTRTRN3.DataSource = dst.Tables("SOTRTRN3")
                Dim dvw As DataView = dst.Tables("SOTRTRN3").DefaultView
                Dim RTRN_LNO As Integer = 0
                If grdSOTRTRN2.ActiveRow IsNot Nothing Then
                    RTRN_LNO = Val(grdSOTRTRN2.ActiveRow.Cells("RTRN_LNO").Text)
                End If
                dvw.RowFilter = "RTRN_LNO = " & CStr(RTRN_LNO)
            ElseIf optGL.Value = "S" Then
                Dim tbl As DataTable = dst.Tables("SOTRTRN3").Clone
                Dim RTRN_GNO As Integer = 0
                For Each rowA234 As DataRow In ASCDATA1.SelectDistinct _
                ("SOTRTRN3", New String() {"ACCT_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE", "ACCT_DESC"}).Rows
                    Dim DIST_AMT As Decimal = dst.Tables("SOTRTRN3").Compute _
                    ("SUM(DIST_AMT)",
                     "ACCT_CODE = '" & rowA234.Item("ACCT_CODE") & "' and SEG2_CODE = '" & rowA234.Item("SEG2_CODE") & "' and SEG3_CODE = '" & rowA234.Item("SEG3_CODE") & "' and SEG4_CODE = '" & rowA234.Item("SEG4_CODE") & "'")
                    Dim row As DataRow = tbl.NewRow
                    row.Item("RTRN_NO") = Absx1.txtFor("RTRN_NO").Text
                    row.Item("RTRN_LNO") = 0
                    RTRN_GNO += 1
                    row.Item("RTRN_GNO") = RTRN_GNO
                    row.Item("ACCT_CODE") = rowA234.Item("ACCT_CODE")
                    row.Item("SEG2_CODE") = rowA234.Item("SEG2_CODE")
                    row.Item("SEG3_CODE") = rowA234.Item("SEG3_CODE")
                    row.Item("SEG4_CODE") = rowA234.Item("SEG4_CODE")
                    row.Item("ACCT_DESC") = rowA234.Item("ACCT_DESC")
                    row.Item("DIST_AMT") = DIST_AMT
                    tbl.Rows.Add(row)
                Next

                grdSOTRTRN3.DataSource = tbl
            End If
        End If
    End Sub

    Private Sub cbeYP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeYP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Refresh_Documents()
    End Sub

    Private Sub cbeInvoiceHistory_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeInvoiceHistory.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Load_SOTINVHH()
    End Sub

    Sub Refresh_Documents()

        If Not dst.Tables.Contains("SOTRTRNX") Then
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        Static YP As String = String.Empty

        EnforceConstraints(False)

        Dim RTRN_AS_PO_REC As String = String.Empty
        Select Case optRetOrPO.Value
            Case "A"
                RTRN_AS_PO_REC = ""
            Case "R"
                RTRN_AS_PO_REC = "ISNULL(RTRN_AS_PO_REC, '0') = '0'"
            Case "P"
                RTRN_AS_PO_REC = "ISNULL(RTRN_AS_PO_REC, '0') = '1'"
        End Select

        If YP <> cbeYP.Value Then
            YP = cbeYP.Value
            Fill_Records("SOTRTRNX", YP)
            Fill_Records("SOTRTRNX2", YP)
            Fill_Records("SOTRTRNX_D", YP)
        End If

        Dim vSOTRTRNX As DataView = dst.Tables("SOTRTRNX").DefaultView
        vSOTRTRNX.RowFilter = RTRN_AS_PO_REC

        Dim vSOTRTRNX_D As DataView = dst.Tables("SOTRTRNX_D").DefaultView
        vSOTRTRNX_D.RowFilter = RTRN_AS_PO_REC

        EnforceConstraints(True)

        grdSOTRTRNX.Text = "Entered in " & cbeYP.Text
        grdSOTRTRNX_D.Text = "Entered in " & cbeYP.Text

        grdSOTRTRNX.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        grdSOTRTRNX_D.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        grdSOTRTRNX.DisplayLayout.Bands(0).Columns("RTRN_NOTE").Width = 250
        grdSOTRTRNX_D.DisplayLayout.Bands(0).Columns("RTRN_NOTE").Width = 250

        If chkGL.Checked Then
            Fill_Records("SOTRTRNG", YP)
            grdSOTRTRN3.Text = "Entered in " & cbeYP.Text
        End If
    End Sub

    Private Sub chkGL_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkGL.CheckedChanged
        Setup_tab0_GL()
    End Sub

    Sub Setup_tab0_GL()
        If Not chkGL.Checked Then
            tab0.Tabs(0).Selected = True
        Else
            Refresh_Documents()
        End If
        tab0.Tabs("GL").Visible = chkGL.Checked

        If chkGL.Checked Then
            tab0.Tabs("GL").Selected = True
        End If
    End Sub

    Sub Set_Up_Reversal()

        Dim REVERSED_BY_RTRN_NO As String = ""
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            REVERSED_BY_RTRN_NO = ASCMAIN1.Next_Control_No("TRAN_NO_C")
        Else
            REVERSED_BY_RTRN_NO = ASCMAIN1.Next_Control_No("SOTRTRN1.RTRN_NO")
        End If

        rowSOTRTRN1 = dst.Tables("SOTRTRN1").Rows(0)
        rowSOTRTRN1.AcceptChanges()
        rowSOTRTRN1.SetAdded()

        With rowSOTRTRN1
            .Item("REVERSED_RTRN_NO") = .Item("RTRN_NO")
            .Item("RTRN_NO") = REVERSED_BY_RTRN_NO
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("RTRN_DATE") = .Item("RTRN_DATE") ' DATETIME_STAMP.Date
            .Item("RTRN_SALES") = -1 * Val(.Item("RTRN_SALES") & "")
            .Item("RTRN_COSTS") = -1 * Val(.Item("RTRN_COSTS") & "")
            .Item("RTRN_STAX") = -1 * Val(.Item("RTRN_STAX") & "")
            .Item("RTRN_FREIGHT") = -1 * Val(.Item("RTRN_FREIGHT") & "")
            .Item("RTRN_HANDLING") = -1 * Val(.Item("RTRN_HANDLING") & "")
            .Item("RTRN_AMOUNT") = -1 * Val(.Item("RTRN_AMOUNT") & "")
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("REGISTER_IND") = "0"
            .Item("REGISTER_XNO") = DBNull.Value
        End With

        'Set new RTRN_NO and reverse all quantities for this return.
        For Each row As DataRow In dst.Tables("SOTRTRN2").Rows
            row.Item("RTRN_NO") = REVERSED_BY_RTRN_NO
            If row.Item("RTRN_QTY") IsNot DBNull.Value Then
                row.Item("RTRN_QTY") = -1 * Val(row.Item("RTRN_QTY") & "")
            End If
            If row.Item("RTRN_QTY_1") IsNot DBNull.Value Then
                row.Item("RTRN_QTY_1") = -1 * Val(row.Item("RTRN_QTY_1") & "")
            End If
            If row.Item("RTRN_QTY_2") IsNot DBNull.Value Then
                row.Item("RTRN_QTY_2") = -1 * Val(row.Item("RTRN_QTY_2") & "")
            End If
            If row.Item("RTRN_QTY_3") IsNot DBNull.Value Then
                row.Item("RTRN_QTY_3") = -1 * Val(row.Item("RTRN_QTY_3") & "")
            End If
            If row.Item("RTRN_QTY_4") IsNot DBNull.Value Then
                row.Item("RTRN_QTY_4") = -1 * Val(row.Item("RTRN_QTY_4") & "")
            End If
            If row.Item("OPS_YYYYPP") IsNot DBNull.Value Then
                row.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            End If

            row.AcceptChanges()
            row.SetAdded()
        Next
    End Sub

    Private Sub grdSOTRTRN2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTRTRN2.InitializeRow
        With e.Row.Cells("RTRN_QTY")
            If Val(e.Row.Cells("RTRN_QTY_TOTAL").Value & "") <> Val(.Value & "") Then
                .Appearance.ForeColor = Color.Red
                .ToolTipText = "Total Return does not balance with Sum of Stock + Refurb + Destroy"
            Else
                .Appearance.ForeColor = Color.Empty
                .ToolTipText = ""
            End If
        End With

        If e.Row.Cells("RTRN_AS_PO_REC").Value & String.Empty = "1" Then
            e.Row.Appearance.BackColor = Color.Pink
        End If

    End Sub

    Private Sub Setup_3PL()
        Fill_Records("EDTRTRN1")
        Fill_Records("EDTRTRN2")
    End Sub

    Sub Setup_3PL_VAN()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        'ASCMAIN1.sql = "Select RCPTHDR.TRANS_SEQ, RCPTHDR.ARRDTE, RCPTHDR.PO_SHIPMENT_NO, RCPTHDR.CONTAINER_NO " _
        '    & ", RCPTDTL.ITEM_CODE, RCPTDTL.RCVQTY" _
        '    & " from ADS.RCPTHDR@ADSIIS,ADS.RCPTDTL@ADSIIS where RCPTHDR.STATUS in ('0','V')" _
        '    & " and RCPTHDR.INVTYP = 'R'" _
        '    & "AND RCPTHDR.TRANS_SEQ = RCPTDTL.TRANS_SEQ " _
        '    & "AND RCPTHDR.LP_CODE = RCPTDTL.LP_CODE " _
        '    & "AND RCPTHDR.WHSE_CODE = RCPTDTL.WHSE_CODE "

        If Not dst.Tables.Contains("RCPTHDR") Then
            ASCMAIN1.sql = "Select RCPTHDR.*" _
                & " from ADS.RCPTHDR@ADSIIS where RCPTHDR.STATUS in ('0','V')" _
                & " and RCPTHDR.INVTYP = 'R'"
            Create_TDA(dst.Tables.Add, "RCPTHDR", "**", 0, False, "", 0)
            'Dim RCPTHDR As DataTable = ASCDATA1.GetDataTable

            ASCMAIN1.sql = "Select RCPTDTL.*" _
                & " from ADS.RCPTHDR@ADSIIS,ADS.RCPTDTL@ADSIIS where RCPTHDR.STATUS in ('0','V')" _
                & " and RCPTHDR.INVTYP = 'R'" _
                & " and RCPTHDR.TRANS_SEQ = RCPTDTL.TRANS_SEQ " _
                & " and RCPTHDR.LP_CODE = RCPTDTL.LP_CODE " _
                & " and RCPTHDR.WHSE_CODE = RCPTDTL.WHSE_CODE "
            Create_TDA(dst.Tables.Add, "RCPTDTL", "**", 0, False, "", 0)
            dst.Tables("RCPTDTL").Columns.Add("ITEM_CODE")
            dst.Tables("RCPTDTL").Columns.Add("ITEM_DESC")
            'Dim RCPTDTL As DataTable = ASCDATA1.GetDataTable
            dst.Relations.Add(dst.Tables("RCPTHDR").Columns("TRANS_SEQ"), dst.Tables("RCPTDTL").Columns("TRANS_SEQ"))
        End If

        EnforceConstraints(False)
        Fill_Records("RCPTHDR")
        Fill_Records("RCPTDTL")

        Dim RCPTDTL2 As DataTable = dst.Tables("RCPTDTL").Clone

        For Each rowRCPTDTL As DataRow In dst.Tables("RCPTDTL").Select("")
            Dim ITEM_CODE As String = rowRCPTDTL.Item("ITEM_CODE")
            If ITEM_CODE.EndsWith("PPK") Then
                ASCMAIN1.sql = "Select Sum (PPK_QTY) from WHTPPKM2 where PPK_CODE = '" & ITEM_CODE & "'"
                Dim PPK_QTY As Int64 = Val(ASCDATA1.GetDataValue)
                ASCMAIN1.sql = "Select * from WHTPPKM2 where PPK_CODE = '" & ITEM_CODE & "'"
                For Each rowWHTPPKM2 As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim rowRCPTDTL2 As DataRow = RCPTDTL2.NewRow
                    rowRCPTDTL2.ItemArray = rowRCPTDTL.ItemArray
                    rowRCPTDTL2.Item("ITEM_CODE") = rowWHTPPKM2.Item("ITEM_CODE")
                    Dim QTY As Int64 = Val(rowWHTPPKM2.Item("PPK_QTY") & "") * Val(rowRCPTDTL.Item("RCVQTY") & "") / PPK_QTY
                    rowRCPTDTL2.Item("RCVQTY") = QTY
                    RCPTDTL2.Rows.Add(rowRCPTDTL2)
                Next
                rowRCPTDTL.Delete()
            End If
        Next

        For Each row As DataRow In RCPTDTL2.Select("")
            dst.Tables("RCPTDTL").Rows.Add(row.ItemArray)
        Next


        EnforceConstraints(True)

        For Each row As DataRow In dst.Tables("RCPTDTL").Select("")
            Dim ITEM_CODE As String = row.Item("ITEM_CODE") & ""

            Dim rowICTITEM1 = LookUp("ICTITEM1", ITEM_CODE)
            If rowICTITEM1 IsNot Nothing Then
                row.Item("ITEM_DESC") = rowICTITEM1.ITEM("ITEM_DESC") & ""
            End If
        Next
        grdEDTRTRN1.DataSource = dst.Tables("RCPTHDR")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Load_SOTINVHH()
        If SELECTION_NO = 0 Then Exit Sub
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Dim YP As String = cbeInvoiceHistory.Value  ' ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -3)
        Fill_Records("SOTINVHH", New String() {CUST_CODE, YP})
        Sort_grdColumns(grdSOTINVHH, "INV_NO".ToLower)
    End Sub

    Private Sub grdSOTINVHH_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTINVHH.DoubleClickRow
        If Not InquiryMode Then

        End If
    End Sub

    Sub Load_SOTINVHX(ITEM_CODE As String)
        Dim YP As String = cbeInvoiceHistory.Value '  ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -3)
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Fill_Records("SOTINVHX", New String() {CUST_CODE, YP, ITEM_CODE})
        Sort_grdColumns(grdSOTINVHX, "INV_DATE".ToLower)
        grdSOTINVHX.Text = "Recent Sales of item " & ITEM_CODE & " To " & Absx1.txtFor("CUST_CODE").Text
    End Sub

    Private Sub grdEDTRTRN1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdEDTRTRN1.AfterRowActivate
        If grdEDTRTRN1.ActiveRow.Band.Key <> grdEDTRTRN1.DisplayLayout.Bands(0).Key Then
            Exit Sub
        End If

        Dim CUST_CODE As String = String.Empty
        Try
            CUST_CODE = grdEDTRTRN1.ActiveRow.Cells("CUST_CODE").Value & String.Empty
        Catch ex As Exception

        End Try

        If CUST_CODE.Length > 0 Then
            viewSOT3PLF1.RowFilter = "CUST_CODE = '" & CUST_CODE & "'"
            viewSOT3PLF1.Sort = "RA_NO"
        Else
            viewSOT3PLF1.RowFilter = String.Empty
            viewSOT3PLF1.Sort = "CUST_CODE, RA_NO"
        End If

    End Sub

    Private Sub grdEDTRTRN1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdEDTRTRN1.DoubleClickRow

        If InquiryMode Then Exit Sub
        If grdEDTRTRN1.ActiveRow Is Nothing Then Exit Sub
        If grdEDTRTRN1.ActiveRow.Band.Index <> 0 Then Exit Sub

        ' Automated Service Lock
        If Not ASCMAIN1.Logical_Lock("SERVICE", "SOCRTRN1",, False) Then
            MessageBox.Show("The automated service is currently processing 3PL Returns. Please try again later.", "Load Return", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        EDI_DOC_SEQ_NO_List.Clear()
        EDI_DOC_SEQ_NO = grdEDTRTRN1.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & String.Empty
        EDI_DOC_SEQ_NO_List.Add(EDI_DOC_SEQ_NO)

        Dim EDI_RA_NO As String = grdEDTRTRN1.ActiveRow.Cells("EDI_RA_NO").Value & String.Empty
        Dim CUST_CODE As String = grdEDTRTRN1.ActiveRow.Cells("CUST_CODE").Value & String.Empty
        Dim WHSE_CODE As String = grdEDTRTRN1.ActiveRow.Cells("WHSE_CODE").Value & String.Empty
        Dim EDI_CUST_SHIP_TO As String = grdEDTRTRN1.ActiveRow.Cells("EDI_CUST_SHIP_TO").Value & String.Empty
        Dim EDI_CUSTOMER_NO As String = grdEDTRTRN1.ActiveRow.Cells("EDI_CUSTOMER_NO").Value & String.Empty

        If WHSE_CODE.Length = 0 Then
            MessageBox.Show("The selected Return does not have an assigned warehouse.", "Load Return", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If dst.Tables("WHTTPLP1").Select($"WHSE_CODE_RTN = '{WHSE_CODE}'").Length = 0 Then
            MessageBox.Show("The selected return's warehouse is not assigned to a 3PL.", "Load Return", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        ElseIf dst.Tables("WHTTPLP1").Select($"WHSE_CODE_RTN = '{WHSE_CODE}'").Length > 1 Then
            MessageBox.Show("The selected return's warehouse is assigned to more than one 3PL.", "Load Return", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        rowWHTTPLP1 = dst.Tables("WHTTPLP1").Select($"WHSE_CODE_RTN = '{WHSE_CODE}'")(0)

        If rowWHTTPLP1.Item("WHSE_CODE") & String.Empty = String.Empty Then
            MessageBox.Show("The selected return's 3PL Code is not assigned a Main warehouse.", "Load Return", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        ElseIf rowWHTTPLP1.Item("WHSE_CODE_RFB") & String.Empty = String.Empty Then
            MessageBox.Show("The selected return's 3PL Code is not assigned a Refurbish warehouse.", "Load Return", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        ElseIf rowWHTTPLP1.Item("WHSE_CODE_RTN") & String.Empty = String.Empty Then
            MessageBox.Show("The selected return's 3PL Code is not assigned a Returns warehouse.", "Load Return", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ' 03/08/2024 - Added WHSE_ODE since retirns will come from multiple warehouses
        Dim sql As String = $"EDI_RA_NO = '{EDI_RA_NO}' AND CUST_CODE = '{CUST_CODE}' AND WHSE_CODE = '{WHSE_CODE}'"
        pl3Cust_store_no = String.Empty

        ' Let user know the Customer is not set up
        If ASCMAIN1.CLIENT = "INT" AndAlso CUST_CODE = String.Empty Then
            ASCMAIN1.sql = "SELECT * FROM ARTCUST2 WHERE CUST_NO_3PL = '" & EDI_CUSTOMER_NO & "' AND CUST_STORE_NO_3PL = '" & EDI_CUST_SHIP_TO & "'"
            Dim tblCustomer As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ARTCUST2")

            If tblCustomer.Rows.Count = 0 Then
                ASCMAIN1.sql = "SELECT * FROM ARTCUST1 WHERE CUST_NO_3PL = '" & EDI_CUSTOMER_NO & "' AND CUST_STORE_NO_3PL = '" & EDI_CUST_SHIP_TO & "'"
                tblCustomer = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ARTCUST1")
            End If

            ' Check Store Map Exceptions
            If tblCustomer.Rows.Count = 0 Then
                ASCMAIN1.sql = "SELECT * FROM TATXREFX WHERE CSCUS1 = '" & EDI_CUSTOMER_NO & "' AND CSCUS2 = '" & EDI_CUST_SHIP_TO & "'"
                tblCustomer = ASCDATA1.GetDataTable(ASCMAIN1.sql, "TATXREFX")

                If tblCustomer.Rows.Count > 0 Then
                    Dim CUST_CODEx As String = tblCustomer.Rows(0).Item("CUST_CODE")
                    Dim CUST_STORE_NO As String = tblCustomer.Rows(0).Item("CUST_STORE_NO")
                    ASCMAIN1.sql = "SELECT * FROM ARTCUST2 WHERE CUST_CODE = '" & CUST_CODEx & "' AND CUST_STORE_NO = '" & CUST_STORE_NO & "'"
                    tblCustomer = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ARTCUST2")
                End If
            End If

            Dim numRecords As Int16 = tblCustomer.Rows.Count

            Select Case numRecords
                Case 0
                    MessageBox.Show("No Cross Reference for Customer/Ship To (" & EDI_CUSTOMER_NO & "/" & EDI_CUST_SHIP_TO & ") on Return.")
                    pl3Cust_store_no = String.Empty
                    Exit Sub

                Case 1
                    grdEDTRTRN1.ActiveRow.Cells("CUST_CODE").Value = tblCustomer.Rows(0).Item("CUST_CODE") & String.Empty
                    CUST_CODE = grdEDTRTRN1.ActiveRow.Cells("CUST_CODE").Value
                    pl3Cust_store_no = tblCustomer.Rows(0).Item("CUST_STORE_NO") & String.Empty

                Case Else
                    Dim cust_codes As String = String.Empty
                    For Each row As DataRow In tblCustomer.Select("", "CUST_CODE,CUST_STORE_NO")
                        cust_codes &= ", ('" & row.Item("CUST_CODE") & "', '" & row.Item("CUST_STORE_NO") & "')"
                    Next

                    ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("CUST_STORE_NO", "ARTCUST2", "(CUST_CODE, CUST_STORE_NO) IN ( " & cust_codes.Substring(1) & ")")
                    If ASCMAIN1.CodeSelector.SQL <> "" Then
                        ASCMAIN1.CodeSelector.MultipleSelections = False
                        ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                        ' Show the Customer Code
                        ASCMAIN1.CodeSelector.grdColumns(0).Item("COLUMN_HIDDEN") = "0"
                        Using F As New ASFCODE1
                            F.ShowDialog()
                        End Using

                        If ASCMAIN1.CodeSelector.SelectedRows.Count <> 1 Then
                            MessageBox.Show("Multiple Cross References for Customer/Ship To (" & EDI_CUSTOMER_NO & "/" & EDI_CUST_SHIP_TO & ") on Return." & cust_codes)
                            Exit Sub
                        End If

                        Dim row As DataRow = ASCMAIN1.CodeSelector.SelectedRows(0)
                        grdEDTRTRN1.ActiveRow.Cells("CUST_CODE").Value = row.Item("CUST_CODE")
                        pl3Cust_store_no = row.Item("CUST_STORE_NO")
                        CUST_CODE = grdEDTRTRN1.ActiveRow.Cells("CUST_CODE").Value
                    End If
            End Select

        ElseIf ASCMAIN1.CLIENT = "INT" Then
            ASCMAIN1.sql = "SELECT * FROM ARTCUST2 WHERE CUST_NO_3PL = '" & EDI_CUSTOMER_NO & "' AND CUST_STORE_NO_3PL = '" & EDI_CUST_SHIP_TO & "'"
            Dim tblCustomer As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ARTCUST2")

            If tblCustomer.Rows.Count = 1 Then
                pl3Cust_store_no = tblCustomer.Rows(0).Item("CUST_STORE_NO") & String.Empty
            End If
        End If


        Select Case ASCMAIN1.CLIENT

            Case "AHA"

            Case "INT"

                ASCMAIN1.sql = "SELECT * FROM SOTRMAF1 WHERE RA_NO = :PARM2"
                Dim rowSOTRMAF1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {EDI_RA_NO})
                If rowSOTRMAF1 Is Nothing Then
                    MessageBox.Show(EDI_RA_NO & " is NOT a valid Returns Authorization Number.", "Returns Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If rowSOTRMAF1.Item("CUST_CODE") & String.Empty <> CUST_CODE Then
                    MessageBox.Show("RA Number (" & EDI_RA_NO & ") belongs to Customer " & rowSOTRMAF1.Item("CUST_CODE") & ".", "Returns Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Fill_Records("ARTCUST3PL", CUST_CODE)

        End Select

        Dim numRecs As Int32 = dst.Tables("EDTRTRN1").Select(sql).Length
        If numRecs > 1 Then
            Select Case MessageBox.Show($"Do you want to load all ({numRecs}) Warehouse Returns For the RA NO: {EDI_RA_NO} from warehouse {WHSE_CODE}?", "Load RA", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
                Case Windows.Forms.DialogResult.Yes
                    EDI_DOC_SEQ_NO_List.Clear()
                    For Each row As DataRow In dst.Tables("EDTRTRN1").Select(sql)
                        If Not EDI_DOC_SEQ_NO_List.Contains(row.Item("EDI_DOC_SEQ_NO")) Then
                            EDI_DOC_SEQ_NO_List.Add(row.Item("EDI_DOC_SEQ_NO"))
                        End If
                    Next

                Case Windows.Forms.DialogResult.No

                Case Windows.Forms.DialogResult.Cancel
                    Exit Sub

            End Select
        End If

        ' Additional Processing
        Select Case ASCMAIN1.CLIENT
            Case "AHA"

            Case "INT"
                ' If multi-selections verify all 3pl customer / stores are set up

                Select Case WHSE_CODE
                    Case "CLARTN"
                        CUST_CODE = String.Empty
                        Dim CUST_STORE_NO As String = String.Empty

                        EDI_CUST_SHIP_TO = String.Empty
                        EDI_CUSTOMER_NO = String.Empty

                        For Each EDI_DOC_SEQ_NO As String In EDI_DOC_SEQ_NO_List
                            rowEDTRTRN1 = dst.Tables("EDTRTRN1").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")(0)
                            CUST_CODE = String.Empty
                            CUST_STORE_NO = String.Empty

                            EDI_CUST_SHIP_TO = rowEDTRTRN1.Item("EDI_CUST_SHIP_TO") & String.Empty
                            EDI_CUSTOMER_NO = rowEDTRTRN1.Item("EDI_CUSTOMER_NO") & String.Empty
                            Convert3PLDoor(EDI_CUSTOMER_NO, EDI_CUST_SHIP_TO, CUST_CODE, CUST_STORE_NO)

                            If CUST_CODE.Length = 0 OrElse CUST_STORE_NO.Length = 0 Then
                                MessageBox.Show("Could Not determine Cross Referencesfor Customer/Ship To (" & EDI_CUSTOMER_NO & "/" & EDI_CUST_SHIP_TO & ") on Return.")
                                Exit Sub
                            End If
                        Next

                    Case "ADSRTN"
                        ' Currently ADS uses the same codes as in ABSolution.
                End Select

        End Select

        processing3PL = True

        MyBase.Absx1.txtFor("RA_NO").Text = (grdEDTRTRN1.ActiveRow.Cells("EDI_RA_NO").Value & String.Empty).ToString.Trim

        If ASCMAIN1.CLIENT = "INT" AndAlso MyBase.Absx1.txtFor("RA_NO").TextLength = 0 Then
            MyBase.Absx1.txtFor("RA_NO").Text = grdEDTRTRN1.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & String.Empty
        End If

        MyBase.Absx1.txtFor("CUST_CODE").Text = CUST_CODE
        MyBase.Absx1.txtFor("WHSE_CODE").Text = WHSE_CODE
        Click_Command("New")
    End Sub

    Private Sub Convert3PLDoor(ByVal EDI_CUSTOMER_NO As String,
                               ByVal EDI_CUST_SHIP_TO As String,
                               ByRef CUST_CODE As String,
                               ByRef CUST_STORE_NO As String)

        CUST_CODE = String.Empty
        CUST_STORE_NO = String.Empty
        Dim rowARTCUST3PL As DataRow = Nothing

        'artcust3pl CONTROL_CODE, CUST_CODE, CUST_STORE_NO, CUST_NO_3PL, CUST_STORE_NO_3PL
        If dst.Tables("ARTCUST3PL").Select("CONTROL_CODE = 'A' AND CUST_NO_3PL = '" & EDI_CUSTOMER_NO & "' AND CUST_STORE_NO_3PL = '" & EDI_CUST_SHIP_TO & "'").Length = 1 Then
            rowARTCUST3PL = dst.Tables("ARTCUST3PL").Select("CONTROL_CODE = 'A' AND CUST_NO_3PL = '" & EDI_CUSTOMER_NO & "' AND CUST_STORE_NO_3PL = '" & EDI_CUST_SHIP_TO & "'")(0)
        ElseIf dst.Tables("ARTCUST3PL").Select("CONTROL_CODE = 'X' AND CUST_NO_3PL = '" & EDI_CUSTOMER_NO & "' AND CUST_STORE_NO_3PL = '" & EDI_CUST_SHIP_TO & "'").Length = 1 Then
            rowARTCUST3PL = dst.Tables("ARTCUST3PL").Select("CONTROL_CODE = 'X' AND CUST_NO_3PL = '" & EDI_CUSTOMER_NO & "' AND CUST_STORE_NO_3PL = '" & EDI_CUST_SHIP_TO & "'")(0)
        End If

        If rowARTCUST3PL IsNot Nothing Then
            CUST_CODE = rowARTCUST3PL.Item("CUST_CODE") & String.Empty
            CUST_STORE_NO = rowARTCUST3PL.Item("CUST_STORE_NO") & String.Empty
        End If

    End Sub

    Private Sub tab0_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        Setup_tab0()
    End Sub

    Sub Setup_tab0()
        If SELECTION_NO = 0 Then Exit Sub
        UltraExplorerBar1.Groups("Invoice History Since").Visible = (tab0.SelectedTab.Key = "Invoice History") And Not ScreenMode
        UltraExplorerBar1.Groups("Show if Entered in").Visible = (tab0.SelectedTab.Key = "Returns" OrElse tab0.SelectedTab.Key = "Return Details" OrElse tab0.SelectedTab.Key = "GL") And Not ScreenMode

        If (tab0.SelectedTab.Key = "Returns Summary") Then
            dst.Tables("SOTRTRNX2").Rows.Clear()
            dst.Tables("SOTRTRNX").Rows.Clear()
            dst.Tables("SOTRTRNX_D").Rows.Clear()
            dst.Tables("SOTRTRNX_I").Rows.Clear()
            grdSOTRTRNX.Parent = splSOTRTRNX.Panel1
            grdSOTRTRNX_D.Parent = splSOTRTRNX.Panel2

            grdSOTRTRNX.Text = "Returns Header"
            grdSOTRTRNX_D.Text = "Returns Details"
        Else
            If grdSOTRTRNX.Parent Is splSOTRTRNX.Panel1 Then
                EnforceConstraints(False)
                Dim YP As String = cbeYP.Value
                Fill_Records("SOTRTRNX", YP)
                Fill_Records("SOTRTRNX2", YP)
                Fill_Records("SOTRTRNX_D", YP)
                EnforceConstraints(True)

                grdSOTRTRNX.Parent = tab0.Tabs("Returns").TabPage
                grdSOTRTRNX_D.Parent = tab0.Tabs("Return Details").TabPage

                grdSOTRTRNX.Text = "Entered in " & cbeYP.Text
                grdSOTRTRNX_D.Text = "Entered in " & cbeYP.Text
            End If

        End If

    End Sub

    Private Sub grdSOTRTRNX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTRTRNX.InitializeRow
        If e.Row.Band.Key = "SOTRTRNX_SOTRTRNX2" Then
        Else
            If e.Row.Cells("REVERSED_BY_RTRN_NO").Value & "" <> "" Then
                e.Row.Cells("RTRN_NO").Appearance.ForeColor = Color.Red
                e.Row.Cells("RTRN_NO").ToolTipText = "Reversed by Return No " & e.Row.Cells("REVERSED_BY_RTRN_NO").Value
            ElseIf e.Row.Cells("REVERSED_RTRN_NO").Value & "" <> "" Then
                e.Row.Cells("RTRN_NO").Appearance.ForeColor = Color.Red
                e.Row.Cells("RTRN_NO").ToolTipText = "Reverses Return No " & e.Row.Cells("REVERSED_RTRN_NO").Value
            Else
                e.Row.Cells("RTRN_NO").Appearance.ForeColor = Color.Empty
                e.Row.Cells("RTRN_NO").ToolTipText = ""
            End If
        End If
    End Sub

    Private Sub grdSOTRMAFX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTRMAFX.DoubleClickRow
        If grdSOTRMAFX.ActiveRow Is Nothing Then Exit Sub

        ' This code needs to act like the Invoice double click grid
        MyBase.Absx1.txtFor("RA_NO").Text = grdSOTRMAFX.ActiveRow.Cells("RA_NO").Value & String.Empty
        Click_Command("New")
    End Sub

    Private Sub Load_SOTRMAFX(Optional FillByDate As Boolean = False)

        If Not FillByDate Then
            ASCMAIN1.sql = "Select * from SOTRMAF1 where RA_STATUS = 'O'"
            Fill_Records("SOTRMAFX", "", , ASCMAIN1.sql)
            grdSOTRMAFX.Text = "Open Returns Authorizations"
            Sort_grdColumns(grdSOTRMAFX, "RA_NO".ToLower)
            grdSOTRMAFX.Visible = True
        Else

        End If
    End Sub

    Private Sub Load_SOT3PLF1()

        EnforceConstraints(False)
        ASCDATA1.ExecuteSQL("truncate table " & SOT3PLF1)
        ASCMAIN1.sql = "Insert into " & SOT3PLF1 & " Select ra_no from SOTRMAF1 where RA_STATUS = 'O'"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        'ASCMAIN1.sql = "Select * from SOTRMAF1 where ra_no in (select ra_no from " & SOT3PLF1 & ")"
        ASCMAIN1.sql = "Select SOTRMAF1.*, ARTCUST1.CUST_NAME from SOTRMAF1, ARTCUST1 where ARTCUST1.CUST_CODE = SOTRMAF1.CUST_CODE AND  SOTRMAF1.RA_STATUS = 'O'" _
            & " AND SOTRMAF1.ra_no in (select ra_no from " & SOT3PLF1 & ")"

        Fill_Records("SOT3PLF1", "", , ASCMAIN1.sql)

        ASCMAIN1.sql = "Select * from SOTRMAF2 where ra_no in (select ra_no from " & SOT3PLF1 & ")"
        Fill_Records("SOT3PLF2", "", , ASCMAIN1.sql)
        EnforceConstraints(True)

        Clear_All_Filters(grdSOT3PLF1)
        Sort_grdColumns(grdSOT3PLF1, "CUST_CODE,RA_NO")
        Sort_grdColumns(grdSOT3PLF1, "RA_LNO", False, 1)

    End Sub

    Private Sub grdSOTRTRN2_DoubleClickCell(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdSOTRTRN2.DoubleClickCell
        If e Is Nothing Then Exit Sub
        If Not e.Cell.Column.Key = "ITEM_CODE" Then Exit Sub
        txtItemCode.Text = e.Cell.Value & String.Empty
        ProcessEnterKeyStroke(txtItemCode.Text)
    End Sub

    Private Sub timItemCode_Tick(sender As System.Object, e As System.EventArgs) Handles timItemCode.Tick
        txtItemCode.Focus()
        timItemCode.Stop()
    End Sub

    Private Sub ProcessEnterKeyStroke(ByVal scannedData As String)

        Dim ITEM_CODE As String = scannedData.Trim
        Dim rowSOTRTRN2 As DataRow = Nothing
        Dim itemFound As Boolean = True
        Dim tblICTITEM1 As DataTable = Nothing

        If dst.Tables("SOTRTRN2").Select("ITEM_CODE = '" & ITEM_CODE & "'").Length = 0 Then
            ' VALIDATE AND ADD ITEM
            tblICTITEM1 = ASCDATA1.GetDataTable("Select * from ICTITEM1 WHERE ITEM_CODE = :PARM1", "", "V", ITEM_CODE)

            If tblICTITEM1.Rows.Count = 0 Then
                tblICTITEM1 = ASCDATA1.GetDataTable("Select * from ICTITEM1 WHERE ITEM_UPC_CODE = :PARM1", "", "V", ITEM_CODE)
                If tblICTITEM1.Rows.Count = 0 Then
                    MessageBox.Show("Invalid Item (" & ITEM_CODE & ").", "Validate Item", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    itemFound = False
                ElseIf tblICTITEM1.Rows.Count > 1 Then
                    MessageBox.Show("Multiple Items found for the provided UPC Code(" & ITEM_CODE & ").", "Validate Item", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    itemFound = False
                Else
                    ITEM_CODE = tblICTITEM1.Rows(0).Item("ITEM_CODE")
                End If
            End If

            If itemFound Then
                If dst.Tables("SOTRTRN2").Select("ITEM_CODE = '" & ITEM_CODE & "'").Length = 0 Then
                    grdSOTRTRN2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes
                    grdSOTRTRN2.DisplayLayout.Bands(0).AddNew()
                    With grdSOTRTRN2.ActiveRow
                        .Cells("ITEM_CODE").Value = ITEM_CODE
                        .Cells("RTRN_QTY").Value = 1
                        .Cells("RTRN_PRICE").Value = 0
                        .Update()
                    End With
                    grdSOTRTRN2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                End If
            End If
        End If

        If itemFound Then
            Dim COLUMN_NAME As String = optLocation.Value
            rowSOTRTRN2 = dst.Tables("SOTRTRN2").Select("ITEM_CODE = '" & ITEM_CODE & "'")(0)
            If Val(rowSOTRTRN2.Item("RTRN_QTY_TOTAL") & String.Empty) >= Val(rowSOTRTRN2.Item("RTRN_QTY") & String.Empty) _
                AndAlso Val(rowSOTRTRN2.Item("RTRN_PRICE") & String.Empty) > 0 Then
                MsgBox("Total quantity may not be more than Return quantity for items with a Non Zero price", MsgBoxStyle.OkOnly, "Cannot Update Row")
            Else
                rowSOTRTRN2.Item(COLUMN_NAME) = Val(rowSOTRTRN2.Item(COLUMN_NAME) & String.Empty) + 1
            End If
        End If

        txtItemCode.Clear()
    End Sub

    Sub CreateTransferOrAdjustment(ByVal RTRN_NO As String)

        ' Point to the return gettin processed.
        Dim rowSOTRTRN1 As DataRow = dst.Tables("SOTRTRN1").Rows.Find(RTRN_NO)

        If processing3PL Then
            IC_PARM_WHSE_CODE_RFB = rowWHTTPLP1.Item("WHSE_CODE_RFB") & String.Empty
            IC_PARM_WHSE_CODE_DST = rowWHTTPLP1.Item("WHSE_CODE_DST") & String.Empty
            IC_PARM_WHSE_CODE_DISC = rowWHTTPLP1.Item("WHSE_CODE_DISC") & String.Empty
            IC_PARM_WHSE_CODE_RTN = rowWHTTPLP1.Item("WHSE_CODE_RTN") & String.Empty
            IC_PARM_WHSE_CODE = rowWHTTPLP1.Item("WHSE_CODE") & String.Empty
        Else
            ' Default Values
            IC_PARM_WHSE_CODE_RFB = ROWs("ICTPARM1").Item("IC_PARM_WHSE_CODE_RFB") & String.Empty
            IC_PARM_WHSE_CODE_DST = ROWs("ICTPARM1").Item("IC_PARM_WHSE_CODE_DST") & String.Empty
            IC_PARM_WHSE_CODE_DISC = ROWs("ICTPARM1").Item("IC_PARM_WHSE_CODE_DISC") & String.Empty
            IC_PARM_WHSE_CODE_RTN = ROWs("ICTPARM1").Item("IC_PARM_WHSE_CODE_RTN") & String.Empty
            IC_PARM_WHSE_CODE = ROWs("ICTPARM1").Item("IC_PARM_WHSE_CODE") & String.Empty
        End If

        ' ISSUE-7369 Return To Stock needs to look at ICTWHSE2
        Dim WHSE_CODE As String = rowSOTRTRN1.Item("WHSE_CODE") & String.Empty
        Fill_Records("ICTWHSE2", {WHSE_CODE})

        ' ISSUE-7369 Return To Stock needs to look at ICTWHSE2
        For Each drSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Select("")
            drSOTRTRN2.Item("WHSE_CODE_STOCK") = IC_PARM_WHSE_CODE
            If dst.Tables("ICTWHSE2").Select($"ITEM_CODE = '{drSOTRTRN2.Item("ITEM_CODE")}'").Length > 0 Then
                drSOTRTRN2.Item("WHSE_CODE_STOCK") = dst.Tables("ICTWHSE2").Select($"ITEM_CODE = '{drSOTRTRN2.Item("ITEM_CODE")}'")(0).Item("WHSE_CODE")
            End If
        Next

        ' ISSUE-7369 Return To Stock needs to look at ICTWHSE2
        Dim lstWhseCodes As New List(Of String)

        For BUCKET_NO As Integer = 1 To 4

            Dim COLUMN_NAME As String = "RTRN_QTY_" & CStr(BUCKET_NO)
            Dim sql As String = "RTRN_NO = '" & RTRN_NO & "' AND " & COLUMN_NAME & " <> 0"
            Dim TOTAL_QTY As Integer = Val(dst.Tables("SOTRTRN2").Compute("SUM(" & COLUMN_NAME & ")", "") & "RTRN_NO = '" & RTRN_NO & "'")

            If TOTAL_QTY = 0 Then
                Continue For
            End If

            Dim WHSE_CODE_TO As String = String.Empty 'IIf(R = 2, rowICTWHSE1.Item("WHSE_CODE_RFB") & "", rowICTWHSE1.Item("WHSE_CODE_DST") & "")
            Dim LOCATION_CODE As String = String.Empty 'IIf(R = 2, rowICTWHSE1.Item("WHSE_LOC_RFB") & "", rowICTWHSE1.Item("WHSE_LOC_DST") & "")

            Select Case BUCKET_NO
                Case 1 ' Stock 
                    WHSE_CODE_TO = IC_PARM_WHSE_CODE
                    LOCATION_CODE = rowICTWHSE1.Item("WHSE_LOC_SHP") & String.Empty

                    ' ISSUE-7369 Return To Stock needs to look at ICTWHSE2
                    lstWhseCodes.Add(WHSE_CODE_TO)
                    For Each drSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Select(sql, "")
                        Dim WHSE_CODE_STOCK As String = drSOTRTRN2.Item("WHSE_CODE_STOCK") & ""
                        If WHSE_CODE_STOCK.Length > 0 Then
                            If Not lstWhseCodes.Contains(WHSE_CODE_STOCK) Then
                                lstWhseCodes.Add(WHSE_CODE_STOCK)
                            End If
                        End If
                    Next

                Case 2 ' Refurbish
                    WHSE_CODE_TO = IC_PARM_WHSE_CODE_RFB
                    LOCATION_CODE = rowICTWHSE1.Item("WHSE_LOC_RFB") & String.Empty
                    ' ISSUE-7369 Return To Stock needs to look at ICTWHSE2
                    lstWhseCodes.Add(WHSE_CODE_TO)

                Case 3 ' Destroy
                    WHSE_CODE_TO = IC_PARM_WHSE_CODE_DST
                    LOCATION_CODE = rowICTWHSE1.Item("WHSE_LOC_DST") & String.Empty
                    ' ISSUE-7369 Return To Stock needs to look at ICTWHSE2
                    lstWhseCodes.Add(WHSE_CODE_TO)

                Case 4 ' Discontinued
                    WHSE_CODE_TO = IC_PARM_WHSE_CODE_DISC
                    LOCATION_CODE = String.Empty ' rowICTWHSE1.Item("WHSE_LOC_DISC") & String.Empty
                    ' ISSUE-7369 Return To Stock needs to look at ICTWHSE2
                    lstWhseCodes.Add(WHSE_CODE_TO)

            End Select

            Select Case ASCMAIN1.CLIENT
                Case "INT"
                    Select Case BUCKET_NO
                        Case 1 ' Stock - Transfer to 3PL (Shipping Warehouse)

                        Case 2 ' Refurbish - Transfer to Refurb Warehouse.
                            '  As per Lauren on 7/8/2016
                            Continue For

                        Case 3 ' Destroy - Adjust out of Stock
                            ' 08/10/2020 Initially done for IPLB for Kate Spade returns
                            Dim AdjReasonCodeOverride As String = String.Empty
                            Select Case ASCMAIN1.CLIENT
                                Case "INT"
                                    If rowSOTRTRN1.Item("RTRN_AS_PO_REC") & String.Empty = "1" Then
                                        AdjReasonCodeOverride = INTSpecialReturnsAdjustmentReasonCode
                                    End If
                            End Select

                            CreateWarehouseAdjustment(rowSOTRTRN1.Item("WHSE_CODE"), LOCATION_CODE, BUCKET_NO, RTRN_NO, AdjReasonCodeOverride)
                            Continue For

                        Case 4 ' Discontinue
                            Continue For
                    End Select

                Case "AHA"
                    Select Case BUCKET_NO
                        Case 1 ' Stock
                            Continue For

                        Case 2 ' Discount

                        Case 3 ' Destroy

                        Case 4 ' Discontinue


                    End Select

                Case Else
                    ' All Other Customers
                    Select Case BUCKET_NO

                        Case 1 ' Stock
                            Continue For

                        Case 2 ' Refurbish

                        Case 3 ' Destroy

                        Case 4 ' Discontinue
                            Continue For
                    End Select
            End Select

            If WHSE_CODE_TO <> "" And TOTAL_QTY <> 0 Then
                ' ISSUE-7369 Return To Stock needs to look at ICTWHSE2
                For Each WHSE_CODE_STOCK As String In lstWhseCodes
                    If dst.Tables("SOTRTRN2").Select(sql & $" AND WHSE_CODE_STOCK = '{WHSE_CODE_STOCK}'", "RTRN_LNO").Length = 0 Then
                        Continue For
                    End If

                    Dim XFR_NO As String = ASCMAIN1.Next_Control_No("ICTIXFR1.XFR_NO")
                    Dim TOTAL_COSTS As Decimal = 0

                    For Each rowSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Select(sql & $" AND WHSE_CODE_STOCK = '{WHSE_CODE_STOCK}'", "RTRN_LNO")
                        Dim rowICTIXFR2 As DataRow = dst.Tables("ICTIXFR2").NewRow
                        With rowICTIXFR2
                            .Item("XFR_NO") = XFR_NO
                            .Item("XFR_LNO") = rowSOTRTRN2.Item("RTRN_LNO")
                            .Item("ITEM_CODE") = rowSOTRTRN2.Item("ITEM_CODE")
                            .Item("XFR_QTY") = rowSOTRTRN2.Item(COLUMN_NAME)

                            .Item("ITEM_COST_STD") = rowSOTRTRN2.Item("ITEM_COST_STD")
                            .Item("COST_CATGY_CODE") = rowSOTRTRN2.Item("COST_CATGY_CODE")
                            .Item("PROD_CODE") = rowSOTRTRN2.Item("PROD_CODE")
                            .Item("OPS_YYYYPP") = rowSOTRTRN2.Item("OPS_YYYYPP")
                            .Item("LOCATION_CODE") = LOCATION_CODE

                            TOTAL_COSTS += Val(.Item("XFR_QTY") & "") * Val(.Item("ITEM_COST_STD") & "")
                        End With
                        dst.Tables("ICTIXFR2").Rows.Add(rowICTIXFR2)
                    Next

                    Dim rowICTIXFR1 As DataRow = dst.Tables("ICTIXFR1").NewRow
                    With rowICTIXFR1
                        .Item("XFR_NO") = XFR_NO
                        .Item("XFR_DATE") = rowSOTRTRN1.Item("RTRN_DATE")
                        .Item("WHSE_CODE") = rowSOTRTRN1.Item("WHSE_CODE")
                        .Item("WHSE_CODE_TO") = WHSE_CODE_STOCK
                        '.Item("XFR_NOTE") =""
                        .Item("INIT_OPER") = rowSOTRTRN1.Item("INIT_OPER")
                        .Item("INIT_DATE") = rowSOTRTRN1.Item("INIT_DATE")
                        .Item("REGISTER_IND") = "0"
                        .Item("JOURNAL_IND") = "0"

                        .Item("XFR_SOURCE") = "R"
                        .Item("OPS_YYYYPP") = rowSOTRTRN1.Item("OPS_YYYYPP")
                        .Item("JOURNAL_IND") = "0"
                        .Item("TOTAL_COSTS") = TOTAL_COSTS
                        .Item("RTRN_NO") = rowSOTRTRN1.Item("RTRN_NO")
                    End With
                    dst.Tables("ICTIXFR1").Rows.Add(rowICTIXFR1)
                Next
            End If
        Next
    End Sub

    Private Sub CreateWarehouseAdjustment(ByVal WHSE_CODE As String, ByVal LOCATION_CODE As String, ByVal BUCKET_NO As Int16, ByVal RTRN_NO As String, ByVal AdjReasonCodeOverride As String)

        Dim returnField As String = "RTRN_QTY_" & BUCKET_NO.ToString.Trim
        Dim rowSOTRTRN1 As DataRow = dst.Tables("SOTRTRN1").Rows.Find(RTRN_NO)

        Dim totalReturnToStock As Integer = Val(dst.Tables("SOTRTRN2").Compute("SUM(" & returnField & ")", "RTRN_NO = '" & RTRN_NO & "'") & String.Empty)
        If totalReturnToStock = 0 Then
            Exit Sub
        End If

        Select Case ASCMAIN1.CLIENT
            Case "INT"

            Case "AHA"

            Case Else

        End Select

        Dim rowICTIADJ1 As DataRow = dst.Tables("ICTIADJ1").NewRow
        rowICTIADJ1.Item("ADJ_NO") = ASCMAIN1.Next_Control_No("ICTIADJ1.ADJ_NO")
        rowICTIADJ1.Item("WHSE_CODE") = WHSE_CODE
        rowICTIADJ1.Item("ADJ_DATE") = CDate(DateTime.Now.ToShortDateString)
        rowICTIADJ1.Item("ADJ_SOURCE") = "E"
        rowICTIADJ1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        rowICTIADJ1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowICTIADJ1.Item("INIT_DATE") = DATETIME_STAMP
        rowICTIADJ1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowICTIADJ1.Item("LAST_DATE") = DATETIME_STAMP
        rowICTIADJ1.Item("REGISTER_IND") = "0"
        rowICTIADJ1.Item("JOURNAL_IND") = "0"

        If AdjReasonCodeOverride.Length > 0 Then
            rowICTIADJ1.Item("REASON_CODE") = AdjReasonCodeOverride
        Else
            rowICTIADJ1.Item("REASON_CODE") = IC_PARM_REASON_ADJ
        End If

        rowICTIADJ1.Item("ADJ_REF") = rowSOTRTRN1.Item("KEY_3PL_RECORD") & String.Empty

        Dim ADJ_NOTE As String = String.Empty
        Select Case ASCMAIN1.CLIENT
            Case "INT"
                ADJ_NOTE = $"{WHSE_CODE} return to Stock."
            Case "AHA"
                ADJ_NOTE = "ADS Return."
        End Select

        If rowSOTRTRN1.Item("KEY_3PL_RECORD") & String.Empty <> String.Empty Then
            ADJ_NOTE &= " 3PL EDI Ref: " & rowSOTRTRN1.Item("KEY_3PL_RECORD") & String.Empty
        End If

        If rowSOTRTRN1.Item("RA_NO") & String.Empty <> String.Empty Then
            ADJ_NOTE &= " RA No:" & rowSOTRTRN1.Item("RA_NO") & String.Empty
        End If

        rowICTIADJ1.Item("ADJ_NOTE") = ADJ_NOTE

        dst.Tables("ICTIADJ1").Rows.Add(rowICTIADJ1)

        Dim ADJ_LNO As Int16 = 0

        Dim TOTAL_COSTS As Decimal = 0
        For Each rowSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Select("RTRN_NO = '" & RTRN_NO & "' AND " & returnField & " > 0", "ITEM_CODE")
            Dim rowICTIADJ2 As DataRow = dst.Tables("ICTIADJ2").NewRow

            rowICTIADJ2.Item("ADJ_NO") = rowICTIADJ1.Item("ADJ_NO")
            ADJ_LNO += 1
            rowICTIADJ2.Item("ADJ_LNO") = ADJ_LNO
            rowICTIADJ2.Item("ITEM_CODE") = rowSOTRTRN2.Item("ITEM_CODE") & String.Empty
            rowICTIADJ2.Item("ADJ_QTY") = rowSOTRTRN2.Item(returnField) * IIf(ASCMAIN1.CLIENT = "INT", -1, 1)
            rowICTIADJ2.Item("ITEM_COST_STD") = rowSOTRTRN2.Item("ITEM_COST_STD") & String.Empty
            rowICTIADJ2.Item("COST_CATGY_CODE") = rowSOTRTRN2.Item("COST_CATGY_CODE") & String.Empty
            rowICTIADJ2.Item("PROD_CODE") = rowSOTRTRN2.Item("PROD_CODE") & String.Empty
            rowICTIADJ2.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowICTIADJ2.Item("LOCATION_CODE") = LOCATION_CODE
            'rowICTIADJ2.Item("BAR_CODE") = String.Empty
            rowICTIADJ2.Item("ADJ_REF") = rowSOTRTRN1.Item("KEY_3PL_RECORD") & String.Empty
            dst.Tables("ICTIADJ2").Rows.Add(rowICTIADJ2)

            TOTAL_COSTS += Val(rowICTIADJ2.Item("ITEM_COST_STD") & String.Empty) * rowICTIADJ2.Item("ADJ_QTY")
        Next

        rowICTIADJ1.Item("TOTAL_COSTS") = TOTAL_COSTS

    End Sub

    Sub Update_WHTLOCBX()
        For Each rowICTIXFR1 As DataRow In dst.Tables("ICTIXFR1").Rows
            TAC.ICCMAIN1.Update_WHTLOCBX("T", rowICTIXFR1.Item("XFR_NO"))
        Next
    End Sub

    Private Sub grdEDTRTRN1_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdEDTRTRN1.InitializeLayout
        e.Layout.Bands(1).SummaryFooterCaption = "Totals:"
    End Sub

    Private Sub grdSOT3PLF1_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOT3PLF1.InitializeLayout
        e.Layout.Bands(1).SummaryFooterCaption = "Totals:"
    End Sub

    Private Sub grdEDTRTRN1_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdEDTRTRN1.InitializeRow
        If ASCMAIN1.CLIENT = "INT" And e.Row.Band.Index = 0 AndAlso e.Row.Cells("CUST_CODE").Value & String.Empty = String.Empty Then
            e.Row.Appearance.BackColor = Color.Red
        End If
    End Sub

    Private Sub optRetOrPO_ValueChanged(sender As Object, e As EventArgs) Handles optRetOrPO.ValueChanged
        Refresh_Documents()
    End Sub

    Private Sub cmdKSP_Click(sender As Object, e As EventArgs) Handles cmdKSP.Click

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now creating work tables")

        Dim YPX As String = cbeYP.Value

        ASCMAIN1.sql = "" _
            & "Select" & vbCrLf _
            & "SOTRTRN1.OPS_YYYYPP, SOTRTRN1.INV_NO, SOTRTRN1.RTRN_NO, SOTRTRN1.RTRN_DATE" & vbCrLf _
            & ", SOTRTRN2.ITEM_CODE, SOTRTRN2.RTRN_QTY, SOTRTRN2.RTRN_PRICE, SOTRTRN2.ITEM_COST_STD" & vbCrLf _
            & ", SOTRTRN2.RTRN_QTY * SOTRTRN2.RTRN_PRICE AMT" & vbCrLf _
            & ", SOTRTRN2.RTRN_QTY * SOTRTRN2.ITEM_COST_STD CGR" & vbCrLf _
            & ", SOTRTRN1.CUST_NAME, SOTRTRN1.CUST_STORE_NO, SOTRTRN1.CUST_CLAIM_NO" & vbCrLf _
            & " from SOTRTRN1,SOTRTRN2" & vbCrLf _
            & " where SOTRTRN2.RTRN_NO = SOTRTRN1.RTRN_NO" & vbCrLf _
            & "   and SOTRTRN1.OPS_YYYYPP >= '" & YPX & "' AND SOTRTRN1.OPS_YYYYPP <= '" & YPX & "'" & vbCrLf _
            & "   and SOTRTRN1.RTRN_AS_PO_REC = '1'"

        Dim tbl As DataTable = ASCDATA1.GetDataTable
        Dim frmmsg As New ASFMSGBF
        frmmsg.Show_grd(tbl, Me, "KSP/LCS Returns entered as a PO in " & cbeYP.Text)

        'EXCEL_SHEET = frmmsg.grow.Cells("TABLE_NAME").Text & "$"

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub txtCUST_CLAIM_NO_ValueChanged(sender As Object, e As EventArgs) Handles txtCUST_CLAIM_NO.ValueChanged


    End Sub

    Private Sub txtCUST_CLAIM_NO_KeyDown(sender As Object, e As KeyEventArgs) Handles txtCUST_CLAIM_NO.KeyDown
        If e.KeyValue = Keys.Enter Then
            If Absx1.txtFor("CUST_CODE").Text = "" Then
                MsgBox("You must first enter a Customer", MsgBoxStyle.OkOnly, "Cannot Fetch Claims Data without a Customer")
                Exit Sub
            End If

            Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
            Dim CUST_CLAIM_NO As String = txtCUST_CLAIM_NO.Text
            Dim sfx As String = $"for Customer {CUST_CODE} Claim No {CUST_CLAIM_NO}"
            grdSOTRTRNX_I.Text = $"Item Summary {sfx}"
            Fill_Records("SOTRTRNX_I", New String() {CUST_CODE, CUST_CLAIM_NO})
            Sort_grdColumns(grdSOTRTRNX_I, "ITEM_CODE")

            EnforceConstraints(False)

            Dim sqlw As String = $"SOTRTRN1.CUST_CODE = '{CUST_CODE}' and SOTRTRN1.CUST_CLAIM_NO = '{CUST_CLAIM_NO}'"
            ASCMAIN1.sql = $"Select SOTRTRN1.* from SOTRTRN1 where {sqlw}"
            Fill_Records("SOTRTRNX",,, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select SOTRTRN1.*, SOTRTRN2.RTRN_LNO, SOTRTRN2.ITEM_CODE, SOTRTRN2.RTRN_QTY, SOTRTRN2.RTRN_QTY_1, SOTRTRN2.RTRN_QTY_2, SOTRTRN2.RTRN_QTY_3, 
                            SOTRTRN2.RTRN_PRICE, SOTRTRN2.ITEM_COST_STD, SOTRTRN2.COST_CATGY_CODE, SOTRTRN2.PROD_CODE" _
            & $" from SOTRTRN1, SOTRTRN2 where SOTRTRN1.RTRN_NO = SOTRTRN2.RTRN_NO AND {sqlw}"
            Fill_Records("SOTRTRNX_D",,, ASCMAIN1.sql)

            ASCMAIN1.sql = $"Select SOTRTRN2.RTRN_NO, SUM(SOTRTRN2.RTRN_QTY) RTRN_QTY_TOTAL 
                from SOTRTRN1, SOTRTRN2 where SOTRTRN1.RTRN_NO = SOTRTRN2.RTRN_NO AND {sqlw}
                GROUP BY SOTRTRN2.RTRN_NO"
            Fill_Records("SOTRTRNX2",,, ASCMAIN1.sql)

            EnforceConstraints(False)

            grdSOTRTRNX.Text = "Entered in " & cbeYP.Text
            grdSOTRTRNX_D.Text = "Entered in " & cbeYP.Text

        End If

    End Sub

End Class