Imports System.Drawing
Imports System.Math
Imports Infragistics.Win.UltraWinGrid

Public Class ICFIREC1
    Dim rowICTIREC1 As DataRow
    Dim location_support As Boolean = False
    Dim rowICTWHSE1 As DataRow
    Dim PO_ORDER_NO_RECEIVED As String
    Dim reversal_update As Boolean = False

    Private PO_ORDER_NO_3PL As String = String.Empty
    Private TRANS_NUM_3PL As String = String.Empty
    Private InvoiceNum As String = String.Empty

    Dim rowICT3PLTX As DataRow

    Dim rowAPTINVH1 As DataRow
    Dim rowAPTVEND5 As DataRow

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "ICFIRECI" Then
            InquiryMode = True
        End If

        Get_PARM("ICTPARM1")
        Get_PARM("WHTPARM1")
        Get_PARM("APTPARM1")
        Get_PARM("GLTPARM1")

        With dst
            ASCMAIN1.sql = "Select ICTIREC1.*" _
            & " from ICTIREC1 where ICTIREC1.OPS_YYYYPP between :PARM1 and :PARM2"
            Create_TDA(.Tables.Add, "ICTIRECX", "**", 0, False, "VV")

            ASCMAIN1.sql = "Select ICTIREC2.*, ICTIREC1.VEND_CODE, ICTIREC1.WHSE_CODE" & vbCrLf _
                & ", ICTIREC1.RECEIPT_DATE, ICTIREC1.REVERSED_BY_RECEIPT_NO, ICTIREC1.REVERSES_RECEIPT_NO, ICTIREC1.SOURCE_DOC_NO" & vbCrLf _
                & ", ICTIREC2.QTY_REC * ICTIREC2.ITEM_COST_STD AMT_REC" & vbCrLf _
                & ", ICTIREC2.QTY_REC * ICTIREC2.PO_COST EXT_PO_COST" & vbCrLf _
                & ", ICTCOSTA.COLLECTION_CODE, ICTCOSTA.ITEM_COST_VCOST STD_VCOST" & vbCrLf _
                & ", (NVL(ICTCOSTA.ITEM_COST_MATLS,0) + NVL(ICTCOSTA.ITEM_COST_LANDGI,0) + NVL(ICTCOSTA.ITEM_COST_TOOLGI,0) + NVL(ICTCOSTA.ITEM_COST_OVRHDI,0)) STD_MATLS" & vbCrLf _
                & " from ICTIREC2,ICTIREC1,ICTCOSTA" & vbCrLf _
                & " where ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
                & "   and ICTCOSTA.ITEM_CODE (+) = ICTIREC2.ITEM_CODE" & vbCrLf _
                & "   and ICTCOSTA.OPS_YYYYPP (+) = ICTIREC2.OPS_YYYYPP" & vbCrLf _
                & "   and ICTIREC2.OPS_YYYYPP between :PARM1 and :PARM2"
            Create_TDA(.Tables.Add, "ICTIRECY", "**", 0, False, "VV")
            .Tables("ICTIRECY").Columns.Add("ORIG_PV", GetType(System.Decimal), "ISNULL(QTY_REC,0)*(ISNULL(PO_COST,0)-ISNULL(STD_VCOST,0))")
            .Tables("ICTIRECY").Columns.Add("ORIG_MV", GetType(System.Decimal), "ISNULL(EXT_COST_MATLS,0) - ISNULL(QTY_REC,0)*ISNULL(STD_MATLS,0)")
            .Tables("ICTIRECY").Columns.Add("AMT_PUR", GetType(System.Decimal), "ISNULL(QTY_REC,0)*ISNULL(PO_COST,0)")


            Create_Relation("ICTIRECX", "ICTIRECY", "RECEIPT_NO")
            .Tables("ICTIRECX").Columns.Add("AMT_PUR", GetType(System.Decimal), "SUM (CHILD.AMT_PUR)")

            ASCMAIN1.sql = "Select ICTIREC3.*, GLTACCT1.ACCT_DESC" _
            & ", ICTIREC1.RECEIPT_DATE, ICTIREC1.WHSE_CODE, ICTIREC1.VEND_CODE" _
            & ", ICTIREC1.SOURCE_DOC_NO, ICTIREC1.INIT_OPER, ICTIREC1.INIT_DATE" _
            & ", ICTIREC1.RECEIPT_SOURCE, ICTIREC1.OPS_YYYYPP, ICTIREC1.PO_ORDER_NO" _
            & " from ICTIREC1,ICTIREC3,GLTACCT1 where ICTIREC1.OPS_YYYYPP between :PARM1 and :PARM2" _
            & " and GLTACCT1.ACCT_CODE = ICTIREC3.ACCT_CODE" _
            & " and ICTIREC3.RECEIPT_NO = ICTIREC1.RECEIPT_NO"
            Create_TDA(.Tables.Add, "ICTIRECG", "**", 0, False, "VV")

            ASCMAIN1.sql = "Select POTORDR1.PO_ORDER_NO, POTORDR1.VEND_CODE, POTORDR1.VEND_NAME" & vbCrLf _
                & ", POTORDR1.PO_REFERENCE, POTORDR1.PO_TYPE, POTORDR1.WHSE_CODE" & vbCrLf _
                & ", POTORDR1.PO_SHIP_VESSEL, POTORDR1.CONTAINER_NO, POTORDR1.PO_DATE_SHIPPED, POTORDR1.PO_DATE_ETA" & vbCrLf _
                & " from POTORDR1 where POTORDR1.PO_STATUS = 'O'"
            Create_TDA(.Tables.Add, "POTORDRO", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO" & vbCrLf _
                & ", POTORDR2.ITEM_CODE, POTORDR2.ITEM_DESC, POTORDR2.ITEM_UOM" _
                & ", POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_OPN, POTORDR2.PO_DATE_REQUIRED, POTORDR2.WHSE_CODE, POTORDR2.PO_COST" _
                & " from POTORDR1,POTORDR2 where POTORDR2.PO_STATUS = 'O'" & vbCrLf _
                & " and POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO"
            Create_TDA(.Tables.Add, "POTORDRD", "**", 0, False, "", 2)

            Create_Relation("POTORDRO", "POTORDRD", "PO_ORDER_NO")

            ASCMAIN1.sql = "Select POTORDR2.*" & vbCrLf _
                & " from POTORDR1,POTORDR2 where POTORDR2.PO_STATUS = 'O' and POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO"
            Create_TDA(.Tables.Add, "POTORDRX", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "POTORDR1", "*")
            Create_TDA(.Tables.Add, "POTORDR2", "*")

            Create_TDA(.Tables.Add("POTORDR1_3PL"), "POTORDR1", "*")
            Create_TDA(.Tables.Add("POTORDR2_3PL"), "POTORDR2", "*")
            ' SR-6549 - Lot Numbers on Shipments and Receipts
            With dst.Tables("POTORDR2_3PL")
                .Columns.Add("TRX_NO", GetType(String))
                .Columns.Add("TRX_LNO", GetType(Int32))
                .Columns.Add("LOT_NO", GetType(String))
            End With

            Create_TDA(.Tables.Add(""), "POTORDR2_3PL_ERR", "Select PO_QTY_REC Qty, ITEM_CODE Item, ITEM_DESC Description from POTORDR2", 0, False, "", 0)

            ASCMAIN1.sql = "Select BM_ISSUE_NO, BM_ISSUE_DATE, BM_ISSUE_COMMENT" _
                    & " from BMTMAIN2 where BM_PROD_ITEM = :PARM1" _
                    & " and BM_ISSUE_NO <> '00'"
            Create_TDA(.Tables.Add, "BMTMAIN2", "**", 0, False, "V")

            ASCMAIN1.sql = "Select POTORDR9.*, ICTITEM1.ITEM_DESC" _
                & " from POTORDR9,ICTITEM1 where ICTITEM1.ITEM_CODE = POTORDR9.ITEM_CODE"
            Create_TDA(.Tables.Add, "POTORDR9", "**", 3)

            Create_TDA(.Tables.Add, "ICTIREC1", "*")
            '.Tables("").Columns.Add("")
            ASCMAIN1.sql = " Select EDTTRXN1.TRANS_NUM, EDTTRXN1.PO_ORDER_NO, EDTTRXN1.TRANS_DATE" & vbCrLf _
                & ", POTORDR1.VEND_CODE, POTORDR1.VEND_NAME, EDTTRXN1.WHSE_CODE" & vbCrLf _
                & ", SUM(NVL(TRAN_QTY, 0)) TRAN_QTY, MIN(LOCATION) LOCATION" & vbCrLf _
                & " from EDTTRXN1, POTORDR1" & vbCrLf _
                & " where EDTTRXN1.PO_ORDER_NO = POTORDR1.PO_ORDER_NO (+)" & vbCrLf _
                & "   and NVL(EDTTRXN1.PROCESS_IND, '0') = '0' AND EDTTRXN1.TRANS_TYPE = 'REC'" & vbCrLf _
                & " group by EDTTRXN1.TRANS_NUM, EDTTRXN1.PO_ORDER_NO, EDTTRXN1.TRANS_DATE, POTORDR1.VEND_CODE, POTORDR1.VEND_NAME, EDTTRXN1.WHSE_CODE"
            Create_TDA(.Tables.Add, "EDTTRXNX", "**", 0, False, String.Empty, 0)
            .Tables("EDTTRXNX").Columns("TRANS_NUM").DataType = GetType(System.String)

            ASCMAIN1.sql = "Select EDTTRXN1.*, ICTITEM1.ITEM_DESC " & vbCrLf _
                & " from EDTTRXN1, ICTITEM1" & vbCrLf _
                & " where EDTTRXN1.ITEM_CODE = ICTITEM1.ITEM_CODE (+)" & vbCrLf _
                & "   and NVL(EDTTRXN1.PROCESS_IND, '0') = '0' AND EDTTRXN1.TRANS_TYPE = 'REC'"
            Create_TDA(.Tables.Add, "EDTTRXNZ", "**", 0, False, String.Empty, 0)
            .Tables("EDTTRXNZ").Columns("TRANS_NUM").DataType = GetType(System.String)

            .Relations.Add("EDTTRXNX_EDTTRXNZ",
                          New DataColumn() { .Tables("EDTTRXNX").Columns("TRANS_NUM"), .Tables("EDTTRXNX").Columns("PO_ORDER_NO"), .Tables("EDTTRXNX").Columns("TRANS_DATE")},
                          New DataColumn() { .Tables("EDTTRXNZ").Columns("TRANS_NUM"), .Tables("EDTTRXNZ").Columns("PO_ORDER_NO"), .Tables("EDTTRXNZ").Columns("TRANS_DATE")})

            Create_TDA(.Tables.Add, "EDTTRXN1", "*")

            ASCMAIN1.sql = "Select ICTIREC2.*, ICTITEM1.ITEM_DESC, POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_OPN" _
            & " from ICTIREC2,ICTITEM1,POTORDR2 where ICTITEM1.ITEM_CODE = ICTIREC2.ITEM_CODE" _
            & " and POTORDR2.PO_ORDER_NO (+) = ICTIREC2.PO_ORDER_NO" _
            & " and POTORDR2.PO_ORDER_LNO (+) = ICTIREC2.PO_ORDER_LNO"
            Create_TDA(.Tables.Add, "ICTIREC2", "**", 1)
            With .Tables("ICTIREC2")
                .Columns.Add("EXT_COST_VCOST", GetType(System.Decimal), "ISNULL(QTY_REC,0) * ISNULL(PO_COST,0)")
                .Columns.Add("EXT_COST_TOTAL", GetType(System.Decimal), "ISNULL(QTY_REC,0) * ISNULL(ITEM_COST_STD,0)")
                .Columns.Add("QTY_REC_NOT_INV", GetType(System.Int64), "ISNULL(QTY_REC,0) - ISNULL(QTY_INV,0)")
                .Columns("ACCRUAL_STATUS").DefaultValue = "0"
                ' SR-6549 - Lot Numbers on Shipments and Receipts
                .Columns.Add("TRX_NO", GetType(String))
                .Columns.Add("TRX_LNO", GetType(Int32))
                If Not InquiryMode Then
                    .Columns.Add("LOT_NO", GetType(String))
                End If
            End With

            ASCMAIN1.sql = "Select ICTIREC3.*, GLTACCT1.ACCT_DESC" _
                & " from ICTIREC3,GLTACCT1 where GLTACCT1.ACCT_CODE (+) = ICTIREC3.ACCT_CODE"
            Create_TDA(.Tables.Add, "ICTIREC3", "**", 1)

            ASCMAIN1.sql = "Select ICTIREC4.*, ICTITEM1.ITEM_DESC" _
                & " from ICTIREC4,ICTITEM1 where ICTITEM1.ITEM_CODE (+) = ICTIREC4.ITEM_CODE"
            Create_TDA(.Tables.Add, "ICTIREC4", "**", 1)
            .Tables("ICTIREC4").Columns.Add("EXT_COST_MATLS", GetType(System.Decimal), "ISNULL(QTY_CON,0) * ISNULL(ITEM_COST_STD,0)")

            ' SR-6549 - Lot Numbers on Shipments and Receipts
            Create_TDA(.Tables.Add, "ICTIRECL", "*", 1)
            Create_Relation("ICTIREC2", "ICTIRECL", "RECEIPT_NO,RECEIPT_LNO")
            With .Tables("ICTIREC2")
                .Columns.Add("QTY_REC_LOT", GetType(System.Int32), "SUM(CHILD.REC_QTY)")
            End With

            ASCMAIN1.sql = "Select ICTSTAT2.*" _
                & " from ICTSTAT2 where ITEM_CODE = :PARM1 and WHSE_CODE = :PARM2"
            Create_TDA(.Tables.Add, "ICTSTAT2", "**", 0, False, "VV")

            .Tables.Add("ICTIREC0")
            .Tables("ICTIREC0").Columns.Add("KEY")
            .Tables("ICTIREC0").Columns.Add("DESCRIPTION")

            ASCMAIN1.sql = "Select * from ICTREAS1"
            Create_TDA(.Tables.Add, "ICTREAS1", "**", 0, False)

            ASCMAIN1.sql = "Select * from ICTCLAS1"
            Create_TDA(.Tables.Add, "ICTCLAS1", "**", 0, False)

            Create_TDA(.Tables.Add, "ICTPINV1", "*", 1)
            Create_TDA(.Tables.Add, "ICTPINV2", "*", 1)
            With .Tables("ICTPINV2").Columns
                .Add("QTY_REC", GetType(System.Int64))
                .Add("QTY_INV", GetType(System.Int64))
                .Add("AMT_INV", GetType(System.Decimal), "ISNULL(PINV_COST,0) * ISNULL(QTY_INV,0)")
                .Add("QTY_REC_NOT_INV", GetType(System.Int64))
            End With

            Create_TDA(.Tables.Add, "ICTIADJ1", "*")
            Create_TDA(.Tables.Add, "ICTIADJ2", "*")

            Create_TDA(.Tables.Add, "APTINVH1", "*", 1)
            Create_TDA(.Tables.Add, "APTINVH2", "*", 1)
            ' Create_TDA(.Tables.Add, "APTINVH5", "*", 1)

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


            Create_TDA(.Tables.Add, "APTVEND5", "*")


            If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then

                ASCMAIN1.sql = "Select X.*, ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC from ICTITEM1, (" & vbCrLf _
                    & "Select DTDATE, DTTIME, DTUSER, DTADJC, COUNT (*) RECORDS" & vbCrLf _
                    & ", MAX (CASE WHEN DTTQTY > 0 THEN DTITEM ELSE NULL END) DTITEM" & vbCrLf _
                    & ", SUM (CASE WHEN DTTQTY > 0 THEN 1 ELSE 0 END) DTITEMS_POS" & vbCrLf _
                    & ", SUM (CASE WHEN DTTQTY < 0 THEN 1 ELSE 0 END) DTITEMS_NEG" & vbCrLf _
                    & ", SUM (CASE WHEN DTTQTY > 0 THEN DTTQTY ELSE 0 END) DTTQTY_POS" & vbCrLf _
                    & ", SUM (CASE WHEN DTTQTY < 0 THEN DTTQTY ELSE 0 END) DTTQTY_NEG" & vbCrLf _
                    & " from EDTTRXNA where DTTTYP = 'A' and DTADJC in ('A03') and NVL(PROCESSED_IND,'0') = '0'" & vbCrLf _
                    & " group by DTDATE, DTTIME, DTUSER, DTADJC ) X where ICTITEM1.ITEM_ALT_SORT (+) = X.DTITEM"
                ' WJZ: THE SQL BLOCK BELOW IS TO SUPPORT THE DIS-ASSEMBLY METHOD DISCUSSED ON 03/11/25
                ' PROBABLY WON'T WORK THE NEXT TIME WHEN WE DISASSEMBLE PRODUCT THAT WAS PREVIOUSLY ASSEMBLED
                ' SEE EMAILS SENT FOR DETAILS
                ASCMAIN1.sql &= vbCrLf _
                    & " UNION " & vbCrLf _
                    & "Select X.*, ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC from ICTITEM1, (" & vbCrLf _
                    & "Select DTDATE, DTTIME, DTUSER, DTADJC, COUNT (*) RECORDS" & vbCrLf _
                    & ", MAX (CASE WHEN DTTQTY < 0 THEN DTITEM ELSE NULL END) DTITEM" & vbCrLf _
                    & ", SUM (CASE WHEN DTTQTY > 0 THEN 1 ELSE 0 END) DTITEMS_POS" & vbCrLf _
                    & ", SUM (CASE WHEN DTTQTY < 0 THEN 1 ELSE 0 END) DTITEMS_NEG" & vbCrLf _
                    & ", SUM (CASE WHEN DTTQTY > 0 THEN DTTQTY ELSE 0 END) DTTQTY_POS" & vbCrLf _
                    & ", SUM (CASE WHEN DTTQTY < 0 THEN DTTQTY ELSE 0 END) DTTQTY_NEG" & vbCrLf _
                    & " from EDTTRXNA where DTTTYP = 'A' and DTADJC in ('A21') and NVL(PROCESSED_IND,'0') = '0'" & vbCrLf _
                    & " group by DTDATE, DTTIME, DTUSER, DTADJC ) X where ICTITEM1.ITEM_ALT_SORT (+) = X.DTITEM"

                Create_TDA(.Tables.Add, "ICT3PLTX", "**", 0, False, String.Empty, 3)

                .Tables("ICT3PLTX").Columns("DTITEMS_POS").DataType = GetType(System.Int64)
                .Tables("ICT3PLTX").Columns("DTITEMS_NEG").DataType = GetType(System.Int64)
                .Tables("ICT3PLTX").Columns("RECORDS").DataType = GetType(System.Int64)
                .Tables("ICT3PLTX").Columns("DTTQTY_POS").DataType = GetType(System.Int64)
                .Tables("ICT3PLTX").Columns("DTTQTY_NEG").DataType = GetType(System.Int64)

                .Tables("ICT3PLTX").Columns("DTDATE").DataType = GetType(System.String)
                .Tables("ICT3PLTX").Columns("DTTIME").DataType = GetType(System.String)


                ' Get POs for items with Positive Qty

                ASCMAIN1.sql = "Select X.DTDATE, X.DTTIME, X.DTUSER, ICTITEM1.ITEM_CODE" & vbCrLf _
                    & ", POTORDR1.PO_ORDER_NO, POTORDR1.PO_TYPE, POTORDR1.VEND_WHSE_CODE, POTORDR1.VEND_CODE" & vbCrLf _
                    & ", POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_OPN, POTORDR2.PO_DATE_REQUIRED, POTORDR2.WHSE_CODE" & vbCrLf _
                    & " from POTORDR1,POTORDR2,ICTITEM1, " & vbCrLf _
                    & "(Select * from EDTTRXNA where DTTTYP = 'A' and DTADJC in ('A03') and NVL(PROCESSED_IND,'0') = '0' and EDTTRXNA.DTTQTY > 0) X" & vbCrLf _
                    & " where ICTITEM1.ITEM_ALT_SORT (+) = X.DTITEM" & vbCrLf _
                    & "   and POTORDR2.ITEM_CODE = ICTITEM1.ITEM_CODE" & vbCrLf _
                    & "   and POTORDR2.PO_STATUS = 'O'" & vbCrLf _
                    & "   and POTORDR1.PO_TYPE = 'M'" & vbCrLf _
                    & "   and NVL(POTORDR1.PO_DISASSEMBLY_IND,'0') = '0'" & vbCrLf _
                    & "   and POTORDR1.VEND_CODE = 'CLARINSUSA'" & vbCrLf _
                    & "   and POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO"
                ' WJZ: THE SQL BLOCK BELOW IS TO SUPPORT THE DIS-ASSEMBLY METHOD DISCUSSED ON 03/11/25
                ' PROBABLY WON'T WORK THE NEXT TIME WHEN WE DISASSEMBLE PRODUCT THAT WAS PREVIOUSLY ASSEMBLED
                ' SEE EMAILS SENT FOR DETAILS
                ASCMAIN1.sql &= vbCrLf _
                    & " UNION " & vbCrLf _
                    & "Select X.DTDATE, X.DTTIME, X.DTUSER, ICTITEM1.ITEM_CODE" & vbCrLf _
                    & ", POTORDR1.PO_ORDER_NO, POTORDR1.PO_TYPE, POTORDR1.VEND_WHSE_CODE, POTORDR1.VEND_CODE" & vbCrLf _
                    & ", POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_OPN, POTORDR2.PO_DATE_REQUIRED, POTORDR2.WHSE_CODE" & vbCrLf _
                    & " from POTORDR1,POTORDR2,ICTITEM1, " & vbCrLf _
                    & "(Select * from EDTTRXNA where DTTTYP = 'A' and DTADJC in ('A21') and NVL(PROCESSED_IND,'0') = '0' and EDTTRXNA.DTTQTY > 0) X" & vbCrLf _
                    & " where ICTITEM1.ITEM_ALT_SORT (+) = X.DTITEM" & vbCrLf _
                    & "   and POTORDR2.ITEM_CODE = ICTITEM1.ITEM_CODE" & vbCrLf _
                    & "   and POTORDR2.PO_STATUS = 'O'" & vbCrLf _
                    & "   and POTORDR1.PO_TYPE = 'B'" & vbCrLf _
                    & "   and NVL(POTORDR1.PO_DISASSEMBLY_IND,'0') = '1'" & vbCrLf _
                    & "   and POTORDR1.VEND_CODE = 'CLARINSUSA'" & vbCrLf _
                    & "   and POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO"
                Create_TDA(.Tables.Add, "ICT3PLTZ", "**", 0, False, String.Empty, 0)
                .Tables("ICT3PLTZ").Columns("DTDATE").DataType = GetType(System.String)
                .Tables("ICT3PLTZ").Columns("DTTIME").DataType = GetType(System.String)

                Create_Relation("ICT3PLTX", "ICT3PLTZ", "DTDATE,DTTIME,DTUSER")



                ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, X.* from ICTITEM1, (" & vbCrLf _
                    & "Select * from EDTTRXNA where DTTTYP = 'A' and DTADJC in ('A03','A21') and NVL(PROCESSED_IND,'0') = '0') X" & vbCrLf _
                    & " where ICTITEM1.ITEM_ALT_SORT (+) = X.DTITEM"
                Create_TDA(.Tables.Add, "ICT3PLTY", "**", 0, False, String.Empty, 0)
                .Tables("ICT3PLTY").Columns("DTDATE").DataType = GetType(System.String)
                .Tables("ICT3PLTY").Columns("DTTIME").DataType = GetType(System.String)

                Create_Relation("ICT3PLTX", "ICT3PLTY", "DTDATE,DTTIME,DTUSER")

                Create_TDA(.Tables.Add, "ICTWHSE1", "*")
                Fill_Records("ICTWHSE1", "", True, "SELECT * FROM ICTWHSE1")
            End If

            ASCMAIN1.sql = "Select * from ICTFRTC1"
            Create_TDA(.Tables.Add, "ICTFRTC1", "**", 0, False)
            ASCMAIN1.sql = "Select * from ICTTRFC1"
            Create_TDA(.Tables.Add, "ICTTRFC1", "**", 0, False)

            Create_TDA(.Tables.Add, "TATALRT1", "*")

        End With

        Set_Read_Only(grpTotals, True)

        Fill_Records("ICTREAS1")
        Fill_Records("ICTCLAS1")
        Fill_Records("ICTFRTC1")
        Fill_Records("ICTTRFC1")

        Show_Filter(grdPOTORDRO, True)

        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)
        cbeYP0.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP0.SelectedItem = cbeYP0.Items(0)
        ' cbe.DataSource = ASCDATA1.GetDataTable("Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")

        grdICTIREC0.DataSource = dst.Tables("ICTIREC0")
        grdICTIREC2.DataSource = dst.Tables("ICTIREC2")
        grdICTIREC3.DataSource = dst.Tables("ICTIREC3")
        grdICTIREC4.DataSource = dst.Tables("ICTIREC4")
        grdICTIRECX.DataSource = dst.Tables("ICTIRECX")
        grdICTIRECY.DataSource = dst.Tables("ICTIRECY")
        grdPOTORDRO.DataSource = dst.Tables("POTORDRO")
        grdPOTORDRX.DataSource = dst.Tables("POTORDRX")
        grdICTIRECG.DataSource = dst.Tables("ICTIRECG")
        grdEDTTRXNX.DataSource = dst.Tables("EDTTRXNX")

        If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
            grdICT3PLTX.DataSource = dst.Tables("ICT3PLTX")
        Else
            tab0.Tabs("3PL Assy").Visible = False
        End If

        Create_Summary(grdICTIRECX, "RECEIPT_NO", "Count")
        Create_Summary(grdICTIRECX, New String() {"QTY_REC", "AMT_REC", "QTY_INV", "AMT_INV", "AMT_PUR"})

        Create_Summary(grdICTIRECY, "RECEIPT_NO", "Count")
        Create_Summary(grdICTIRECY, New String() {"QTY_REC", "AMT_REC", "TRAN_PV", "TRAN_MV", "ORIG_PV", "ORIG_MV", "EXT_COST_MATLS", "EXT_PO_COST", "AMT_PUR"})

        Create_Summary(grdICTIRECG, "RECEIPT_NO", "Count")
        Create_Summary(grdICTIRECG, "DIST_AMT")

        Create_Summary(grdPOTORDRO, "PO_ORDER_NO", "Count")

        Create_Summary(grdICTIREC2, "RECEIPT_LNO", "Count")
        Create_Summary(grdICTIREC2, New String() {"QTY_REC", "EXT_COST_VCOST", "EXT_COST_MATLS", "EXT_COST_TOTAL", "TRAN_PV", "TRAN_MV", "TRAN_CV", "TRAN_FV", "TRAN_TV"})
        ' SR-6549 - Lot Numbers on Shipments and Receipts        
        'Create_Summary(grdICTIREC2, "REC_QTY", "Sum", "ICTIREC2_ICTIRECL")
        'Create_Summary(grdICTIREC2, "RECEIPT_LNO_SEQ", "Count", "ICTIREC2_ICTIRECL")

        Create_Summary(grdICTIREC3, "RECEIPT_GNO", "Count")
        Create_Summary(grdICTIREC3, "DIST_AMT")

        Create_Summary(grdICTIREC4, "ITEM_CODE", "Count")
        Create_Summary(grdICTIREC4, "EXT_COST_MATLS")

        With grdICTIRECX.DisplayLayout.Bands("ICTIRECX")
            .Columns("RECEIPT_NO").Header.Fixed = True
        End With

        With grdICTIRECY.DisplayLayout.Bands("ICTIRECY")
            .Columns("RECEIPT_NO").Header.Fixed = True
            .Columns("RECEIPT_LNO").Header.Fixed = True
        End With


        With grdPOTORDRX.DisplayLayout.Bands("POTORDRX")
            .Columns("PO_ORDER_NO").Header.Fixed = True
        End With

        With grdICTIRECG.DisplayLayout.Bands("ICTIRECG")
            .Columns("RECEIPT_NO").Header.Fixed = True
        End With

        'ASCMAIN1.Add_Value_List(grdICTIRECX, "REASON_CODE", "Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")
        'ASCMAIN1.Add_Value_List(grdPOTORDRX, "REASON_CODE", "Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")

        grdICTIREC0.DisplayLayout.Bands(0).ColHeadersVisible = False
        Set_SEGS(grdICTIREC3, "ICTIREC3")
        Set_SEGS(grdICTIRECG, "ICTIRECG")

        Set_Read_Only(grpTotals, True)
        'If InStr(ASCMAIN1.USER_SECURITY_CODEs, "WS") = 0 Then
        '    grpTotals.Visible = False
        '    With grdICTIREC2.DisplayLayout.Bands(0)
        '        .Columns("ITEM_COST_STD").Hidden = True
        '        .Columns("LINE_COSTS").Hidden = True
        '        .Columns("COST_CATGY_CODE").Hidden = True
        '        .Columns("PROD_CODE").Hidden = True
        '    End With
        'End If

        With grdICTIREC2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray

                If New String() {"PINV_QTY", "PINV_COST", "PO_COST_VAR_AMT", "PO_COST_VAR_PCT", "PO_COST_VAR_AMT_IND", "PO_COST_VAR_PCT_IND"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                    gcol.CellActivation = Activation.NoEdit
                End If
            Next
            .Columns("RECEIPT_LNO").Header.Fixed = True
            .Columns("ITEM_CODE").Header.Fixed = True
            .Columns("ITEM_DESC").Header.Fixed = True
        End With


        With grdICTIREC3.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            Next
        End With

        MakeTransparent(chkReceiptsSummaryOnly)

        grpHeader.Visible = False
        'Set_SEGS(grdPOTORDRX, "POTORDRX")
        SplitContainer1.Panel1Collapsed = True ' until we need more header data

        If ASCMAIN1.CLIENT = "INT" Then
            grdEDTTRXNX.DisplayLayout.Bands(0).Columns("LOCATION").Header.Caption = "Invoice"
        Else
            grdEDTTRXNX.DisplayLayout.Bands(0).Columns("LOCATION").Hidden = True
        End If

        ASCMAIN1.Add_Value_List(grdICTIREC3, "DIST_TYPE", , New String() {":", "TOOLG:TARIFF"})

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Done"

                If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" AndAlso Format(Now, "yyyyMMdd") = "20250909" Then
                    Dim r As Integer = 0
                    ASCMAIN1.sql = "SELECT DISTINCT RECEIPT_NO FROM (" & vbCrLf _
                                            & "SELECT X.*, Y.COST_ACC, Y.COST_ACT, Y.COST_ORIG, Y.CTL_NO, Y.CTL_STATUS, Y.CTL_DATE" & vbCrLf _
                                            & ", ROUND (CASE WHEN Y.COST_ACC IS NULL THEN NULL ELSE NVL(X.TRF,0) - NVL(Y.COST_ACC,0) END,2) DIFF FROM (" & vbCrLf _
                                            & "SELECT ICTIREC1.RECEIPT_DATE, ICTIREC1.INIT_OPER, ICTIREC1.LAST_OPER, ICTIREC1.INIT_DATE, ICTIREC1.LAST_DATE" & vbCrLf _
                                            & ", ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO, ROUND (ICTIREC2.PO_COST_TRF  * ICTIREC2.QTY_REC,2)  TRF" & vbCrLf _
                                            & "FROM ICTIREC1,ICTIREC2" & vbCrLf _
                                            & "WHERE ICTIREC1.OPS_YYYYPP = '202509' AND ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
                                            & "AND ICTIREC2.PO_COST_TRF <> 0) X," & vbCrLf _
                                            & "(SELECT CTL_NO, RECEIPT_NO, RECEIPT_LNO, COST_ACC, COST_ACT, COST_ORIG, CTL_STATUS, CTL_DATE" & vbCrLf _
                                            & " FROM APTACRC1 WHERE OPS_YYYYPP = '202509' AND ACCRUAL_CODE = 'TRF' AND NVL(PPD_IND,'0') = '0') Y" & vbCrLf _
                                            & "WHERE Y.RECEIPT_NO (+) = X.RECEIPT_NO AND Y.RECEIPT_LNO (+) = X.RECEIPT_LNO" & vbCrLf _
                                            & "ORDER BY CTL_DATE DESC, RECEIPT_DATE" & vbCrLf _
                                            & ") WHERE CTL_NO IS NULL"
                    For Each ROW As DataRow In ASCDATA1.GetDataTable.Select("")

                        Dim RECEIPT_NO_in As String = ROW.Item("RECEIPT_NO") ' Absx1.txtFor("RECEIPT_NO").Text
                        Debug.Print(RECEIPT_NO_in)
                        r += 1
                        '     TAC.ICCMAIN1.Create_Accrual_FRT(RECEIPT_NO_in)
                        TAC.ICCMAIN1.Create_Accrual_TRF(RECEIPT_NO_in)
                    Next

                    EMsg &= vbCr & $"{CStr(r)} Records added to TRF Subsidiary - Check Accruals"

                End If

            Case "New"
                Validate_Code("WHSE_CODE")

                Dim DT As Date = Absx1.dteFor("RECEIPT_DATE").Value
                If DT & "" = "" Then
                    EMsg &= vbCr & "Document Date is Mandatory"
                Else
                    TAC.SOCMAIN1.Validate_Invoice_Date(DT, 0, 0, EMsg)
                End If

                If Absx1.txtFor("VEND_CODE").Text = "" Then
                    EMsg &= vbCr & "You must supply a Valid Supplier"
                Else
                    Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", Absx1.txtFor("VEND_CODE").Text)
                    If IsNothing(rowAPTVEND1) Then
                        EMsg &= vbCr & "Supplier Entered Is Not Valid"
                    Else
                        ASCMAIN1.sql = "Select Count (*) from POTORDR1 where VEND_CODE = '" & Absx1.txtFor("VEND_CODE").Text & "' and PO_STATUS = 'O'"
                        If Val(ASCDATA1.GetDataValue) = 0 Then
                            EMsg &= vbCr & "No Open POs on file with Supplier " & Absx1.txtFor("VEND_CODE").Text
                        End If
                    End If
                End If

                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCr & "You must supply a Valid Warehouse"
                Else
                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If IsNothing(rowICTWHSE1) Then
                        EMsg &= vbCr & "Warehouse Entered Is Not Valid"
                    Else
                        If rowICTWHSE1.Item("WHSE_STATUS").ToString <> "A" Then
                            EMsg &= vbCr & "Warehouse Entered Is Not Active"
                        Else
                            If rowICTWHSE1.Item("LP_CODE") & "" <> "" AndAlso PO_ORDER_NO_3PL.Length = 0 Then
                                If MsgBox("Warehouse Entered Is A 3PL." & vbCrLf & vbCrLf & "Do you want to Manually Receive anyway?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                                Else
                                    EMsg &= vbCr & "Warehouse Entered Is A 3PL.  No Manual Receipts Allowed"
                                End If
                            End If
                        End If
                    End If
                End If

                EMsg &= TAC.ICCMAIN1.Check_Standard_Cost_Initialization(Me, "ICTIREC2")

                If PO_ORDER_NO_RECEIVED <> "" Then
                    If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO_RECEIVED) Then Exit Sub
                End If

                If EMsg.Length = 0 AndAlso PO_ORDER_NO_3PL <> "" Then
                    If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO_3PL) Then
                        PO_ORDER_NO_3PL = String.Empty
                        Exit Sub
                    End If
                End If

                If EMsg = "" Then
                    Dim PO_ORDER_NO As String = PO_ORDER_NO_RECEIVED
                    If PO_ORDER_NO = "" Then PO_ORDER_NO = PO_ORDER_NO_3PL

                    If PO_ORDER_NO = "" Then
                        EMsg &= vbCr & $"You must select a PO to receive from one of the tabs below"
                    Else
                        Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)

                        Dim FRT_CLASS_CODE As String = rowPOTORDR1.Item("FRT_CLASS_CODE") & ""
                        Dim TRF_CLASS_CODE As String = rowPOTORDR1.Item("TRF_CLASS_CODE") & ""

                        If FRT_CLASS_CODE = "" Then
                            EMsg &= vbCr & $"There is no Freight Class Code specifed for PO {PO_ORDER_NO}"
                        End If
                        If TRF_CLASS_CODE = "" Then
                            EMsg &= vbCr & $"There is no Tariff Class Code specifed for PO {PO_ORDER_NO}"
                        End If
                    End If
                End If

                If EMsg.Length > 0 Then
                    PO_ORDER_NO_3PL = String.Empty
                    InvoiceNum = String.Empty
                End If

            Case "View"
                If Absx1.txtFor("RECEIPT_NO").Text = "" Then
                    EMsg &= vbCr & "You must specify a Document No to View"
                Else
                    rowICTIREC1 = LookUp("ICTIREC1", Absx1.txtFor("RECEIPT_NO").Text)
                    If rowICTIREC1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Document " & Absx1.txtFor("RECEIPT_NO").Text & " on File"
                    End If
                End If

            Case "Update"
                'If Absx1.txtFor("REASON_CODE").Text = "" Then
                '    EMsg &= vbCr & "You Must Specify a Reason"
                'Else
                '    Dim rowICTREAS1 As DataRow = LookUp("ICTREAS1", Absx1.txtFor("REASON_CODE").Text)
                '    If rowICTREAS1 Is Nothing Then
                '        EMsg &= vbCr & "Invalid Value Specified for Reason"
                '    End If
                'End If


                Dim DT As Date = Absx1.dteFor("RECEIPT_DATE").Value
                If DT & "" = "" Then
                    EMsg &= vbCr & "Document Date is Mandatory"
                Else
                    TAC.SOCMAIN1.Validate_Invoice_Date(DT, 0, 0, EMsg)
                End If

                If grdICTIREC2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Details Entered"
                Else
                    Dim ITEM_CODEs As New List(Of String)
                    Dim ITEM_CODE_with_duplicate_PO_ORDER_LNO As String = ""
                    For Each rowICTIREC2 As DataRow In dst.Tables("ICTIREC2").Select("", "", DataViewRowState.CurrentRows)

                        Dim ITEM_CODE As String = rowICTIREC2.Item("ITEM_CODE")
                        If Not ITEM_CODEs.Contains(ITEM_CODE) Then ITEM_CODEs.Add(ITEM_CODE)

                        If rowICTIREC2.Item("COST_CATGY_CODE") & "" = "" Then
                            EMsg &= vbCr & $"Unable to determine Cost Category for Item {ITEM_CODE}"
                        End If
                        If rowICTIREC2.Item("PROD_CODE") & "" = "" Then
                            EMsg &= vbCr & $"Unable to determine Product Code for tem {ITEM_CODE}"
                        End If

                        Dim sqlw As String = $"ITEM_CODE = '{ITEM_CODE}' and PO_ORDER_NO = '" & rowICTIREC2.Item("PO_ORDER_NO") & "' and PO_ORDER_LNO = " & rowICTIREC2.Item("PO_ORDER_LNO") & " and RECEIPT_LNO <> " & rowICTIREC2.Item("RECEIPT_LNO")
                        If dst.Tables("ICTIREC2").Select(sqlw).Length > 0 Then
                            ITEM_CODE_with_duplicate_PO_ORDER_LNO &= "," & rowICTIREC2.Item("ITEM_CODE")
                        End If
                    Next
                    If ITEM_CODE_with_duplicate_PO_ORDER_LNO <> "" Then
                        If MsgBox("Some Items on this Receipt were found" _
                                  & vbCrLf & " on more than one receipt line" _
                                  & vbCrLf & " in connection with the same PO Order and PO Line." _
                                  & vbCrLf & vbCrLf & "OK to Continue?",
                                  MsgBoxStyle.YesNo, "Items found received more than once against Same PO/Line") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If


                    If EMsg = "" Then

                        Dim PO_ORDER_NO As String = rowICTIREC1.Item("PO_ORDER_NO")
                        Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
                        Dim PO_ORDER_TYPE As String = rowPOTORDR1.Item("PO_ORDER_TYPE") & ""
                        If PO_ORDER_TYPE = "R" Then
                            ' OK AND EXPECTED TO SEE VARIANCES
                        Else

                            Dim FRT_CLASS_CODE As String = rowICTIREC1.Item("FRT_CLASS_CODE")
                            Dim TRF_CLASS_CODE As String = rowICTIREC1.Item("TRF_CLASS_CODE")

                            Dim iTEMS_WITH_ISSUES = ""
                            For Each ITEM_CODE As String In ITEM_CODEs
                                Dim rowSTD As DataRow = LookUp("ICTCOSTF", ITEM_CODE)
                                If rowSTD Is Nothing Then
                                    rowSTD = LookUp("ICTCOSTC", ITEM_CODE)
                                End If
                                If rowSTD Is Nothing Then
                                    iTEMS_WITH_ISSUES &= $"{vbCrLf}{ITEM_CODE} No Item Cost"
                                Else
                                    Dim ITEM_COST_FRT_CLASS As String = rowSTD.Item("ITEM_COST_FRT_CLASS") & ""
                                    If ITEM_COST_FRT_CLASS = "" Then ITEM_COST_FRT_CLASS = "Z"
                                    If ITEM_COST_FRT_CLASS <> FRT_CLASS_CODE Then
                                        iTEMS_WITH_ISSUES &= $"{vbCrLf}{ITEM_CODE} Cost Frt = {ITEM_COST_FRT_CLASS}"
                                    End If

                                    Dim ITEM_COST_TRF_CLASS As String = rowSTD.Item("ITEM_COST_TRF_CLASS") & ""
                                    If ITEM_COST_TRF_CLASS = "" Then ITEM_COST_TRF_CLASS = "Z"
                                    If ITEM_COST_TRF_CLASS <> TRF_CLASS_CODE Then
                                        iTEMS_WITH_ISSUES &= $"{vbCrLf}{ITEM_CODE} Cost Trf = {ITEM_COST_TRF_CLASS}"
                                    End If
                                End If
                            Next

                            If iTEMS_WITH_ISSUES <> "" Then
                                If MsgBox($"PO Codes: Freight = {FRT_CLASS_CODE}, Tariff = {TRF_CLASS_CODE}" & vbCrLf & vbCrLf & "Items with Discrepanancies:" & iTEMS_WITH_ISSUES & vbCrLf & vbCrLf & "OK to Continue with Update?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "There are Items with Freight and/or Tariff Discrepancies") = MsgBoxResult.No Then
                                    EMsg &= vbCr & "Fix Freight and/or Tariff Discrepancies where PO Codes <> Cost Codes"
                                End If
                            End If

                            ' Modeled after logic above, check cost against pending, then future, then standard
                            Dim iTEMS_WITH_COST_ISSUES = ""

                            For Each rowICTIREC2 As DataRow In dst.Tables("ICTIREC2").Select("", "", DataViewRowState.CurrentRows)
                                Dim ITEM_CODE As String = rowICTIREC2.Item("ITEM_CODE")
                                Dim PO_COST As String = rowICTIREC2.Item("PO_COST")
                                Dim ITEM_COST_VCURR As Decimal = 0

                                ' change first row to look for ictcostp for current period (receipt is going against the current period)
                                ' possible scenario: PO could be made expecting to be received in may, but be received in april, in which case we will use the april cost (a good thing, but confirm with nathan)

                                ASCMAIN1.sql = $"select * from ICTCOSTP where OPS_YYYYPP<='{ASCMAIN1.CYP}' and item_code='{ITEM_CODE}' ORDER BY OPS_YYYYPP DESC FETCH FIRST 1 ROW ONLY"
                                Dim rowICTCOSTP As DataRow = ASCDATA1.GetDataRow()
                                If rowICTCOSTP IsNot Nothing Then
                                    ITEM_COST_VCURR = Val(rowICTCOSTP.Item("ITEM_COST_VCOST") & "")
                                    ' when 0 rows, move on to costf
                                Else
                                    Dim rowICTCOSTF As DataRow = LookUp("ICTCOSTF", ITEM_CODE)
                                    If rowICTCOSTF Is Nothing Then
                                        Dim rowICTCOSTC As DataRow = LookUp("ICTCOSTC", ITEM_CODE, True)
                                        If rowICTCOSTC Is Nothing Then
                                            iTEMS_WITH_COST_ISSUES &= $"{vbCrLf}{ITEM_CODE} No Item Cost"
                                        Else
                                            ITEM_COST_VCURR = Val(rowICTCOSTC.Item("ITEM_COST_VCURR") & "")
                                        End If
                                    Else
                                        ITEM_COST_VCURR = Val(rowICTCOSTF.Item("ITEM_COST_VCURR") & "")
                                    End If
                                End If

                                If ITEM_COST_VCURR <> PO_COST Then
                                    iTEMS_WITH_COST_ISSUES &= $"{vbCrLf}{ITEM_CODE} PO Cost {PO_COST} <> VCost {ITEM_COST_VCURR}"
                                End If

                            Next

                            If iTEMS_WITH_COST_ISSUES <> "" Then
                                If MsgBox($"Items with Discrepanancies:" & iTEMS_WITH_COST_ISSUES & vbCrLf & vbCrLf & "OK to Continue with Update?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "There are Items with Cost Discrepancies") = MsgBoxResult.No Then
                                    EMsg &= vbCr & "Fix Cost Discrepancies where PO Cost <> VCost"
                                End If
                            End If

                        End If
                    End If

                End If

                Dim SOURCE_DOC_NO As String = Absx1.txtFor("SOURCE_DOC_NO").Text
                ASCMAIN1.sql = $"Select * from ICTIREC1 where PO_ORDER_NO = '{PO_ORDER_NO_3PL}' and SOURCE_DOC_NO = '{SOURCE_DOC_NO}'"
                ASCMAIN1.sql &= " and REVERSED_BY_RECEIPT_NO IS NULL and REVERSES_RECEIPT_NO IS NULL"
                Dim rowICTIREC1s() As DataRow = ASCDATA1.GetDataTable.Select("")
                If rowICTIREC1s.Length > 0 Then
                    If MsgBox("OK to Continue Anyway?", MsgBoxStyle.YesNo, $"Invoice {SOURCE_DOC_NO} has already been received (see Receipt {rowICTIREC1s(0).Item("RECEIPT_NO")})") = MsgBoxResult.No Then
                        EMsg &= vbCr & $"Invoice {SOURCE_DOC_NO} has already been received (see Receipt {rowICTIREC1s(0).Item("RECEIPT_NO") })"
                    End If

                End If

                EMsg &= TAC.ICCMAIN1.Check_Standard_Cost_Initialization(Me, "ICTIREC2")
                ' CHECK FOR ITEM BEING RECEIVED TWICE ON SAME RECEIPT?

                If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                    Check_for_IPSA_Invoice(False)
                End If

                If EMsg = "" Then
                    Dim msg As String = Check_Qty("ICTIREC2", Absx1.txtFor("WHSE_CODE").Text, "QTY_REC", 1)
                    If msg <> "" Then
                        If MsgBox(msg & vbCr & vbCr & "OK To Continue Anyway?",
                                  MsgBoxStyle.YesNo,
                                  "The following Items Do Not have Sufficent Qty Open On PO For this Transaction") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If

                    Dim rowPO As DataRow = LookUp("POTORDR1", rowICTIREC1.Item("PO_ORDER_NO"))
                    If Absx1.txtFor("WHSE_CODE").Text <> rowPO.Item("WHSE_CODE") Then
                        msg = $"Receiving Whse { Absx1.txtFor("WHSE_CODE").Text } is not the same as PO Whse {rowPO.Item("WHSE_CODE")}"
                        If MsgBox(msg & vbCr & vbCr & "OK To Continue Anyway?",
                              MsgBoxStyle.YesNo,
                              "Receiving Warehouse has been Changed") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If

                    ' what if there are other lines on the PO
                    ' how to change this answer after receipt - on either direction - to re-open PO and send to ADS, or close PO and have ADS close PO
                    'If dst.Tables("ICTIREC2").Select("ISNULL(QTY_REC,0) <> ISNULL(PO_QTY_OPN,0)").Length > 0 Then
                    '    msg = $"Some Lines on this PO Receipt will leave a PO Qty Open." & vbCrLf & "Do you want to leave the Balances Open?"
                    '    If MsgBox(msg & vbCr & vbCr & "Yes to Leave Open, No to Close the PO",
                    '        MsgBoxStyle.YesNoCancel,
                    '        "What to do with Balance Still Open on PO after Receipt is Updated") = MsgBoxResult.Cancel Then
                    '        Exit Sub
                    '    End If
                    'End If
                End If

            Case "Reverse"

                ASCMAIN1.sql = "Select Distinct PO_ORDER_NO from ICTIREC2 where RECEIPT_NO = '" & Absx1.txtFor("RECEIPT_NO").Text & "'"
                For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
                    If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then Exit Sub
                Next

                If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                    Check_for_IPSA_Invoice(True)
                End If

                If MessageBox.Show("Are you sure you want to reverse this Entry?", "Confirm Reversal",
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
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

                ' Issue-7044 - balance POs/ASNs
                If EntryMode = "N" AndAlso dst.Tables("ICTIREC1").Rows.Count = 1 Then
                    ' Get Receipt No, check vendor. No IPSA or ADS
                    Select Case dst.Tables("ICTIREC1").Rows(0).Item("VEND_CODE") & String.Empty
                        Case "IPSA", "ADS3PL"
                            ' Currently do not send
                        Case Else
                            ResendPOsWithOpenQuantity(dst.Tables("ICTIREC1").Rows(0).Item("PO_ORDER_NO") & String.Empty)
                    End Select
                End If
                Mode_Settings(False)

            Case "Reverse"
                reversal_update = True
                Update_Record()
                reversal_update = False

                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Refresh"
                Refresh_Documents()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Refresh").Settings.Enabled = not_iScreenMode

                    If InquiryMode Then
                        .Items("New").Visible = False
                        '.Items("Refresh").Visible = False
                        .Items("Update").Visible = False
                        .Items("Cancel").Visible = False
                    End If

                    If ScreenMode And EntryMode <> "N" Then
                        .Items("Update").Settings.Enabled = not_iScreenMode
                        .Items("Cancel").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                    End If

                    If ScreenMode And EntryMode <> "V" Then
                        .Items("Done").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Done").Settings.Enabled = iScreenMode
                    End If

                    .Items("Reverse").Visible = (ScreenMode AndAlso EntryMode = "V") And Not InquiryMode _
                        AndAlso rowICTIREC1 IsNot Nothing _
                        AndAlso rowICTIREC1.Item("REVERSED_BY_RECEIPT_NO") Is DBNull.Value _
                        AndAlso rowICTIREC1.Item("REVERSES_RECEIPT_NO") Is DBNull.Value
                End With

                .Groups("GL Distribution").Visible = False ' ScreenMode And (EntryMode = "V") And InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") <> 0
                .Groups("Show if Entered in").Visible = Not ScreenMode ' And InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") <> 0
                .Groups("Totals").Visible = False ' ScreenMode
                .Groups("Events").Visible = ScreenMode And (EntryMode <> "N")
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = ScreenMode
        grpHeader.Visible = ScreenMode

        tab0.Visible = Not ScreenMode

        'With grdICTIREC2.DisplayLayout.Bands(0)
        '    .Columns("PO_QTY_ORD").Hidden = (EntryMode <> "N")
        '    .Columns("PO_QTY_OPN").Hidden = (EntryMode <> "N")
        'End With

        If ScreenMode Then
            'If InquiryMode Then
            '    With UltraExplorerBar1.Groups("Screen Control")
            '        .Items("New").Visible = False
            '        .Items("Update").Visible = False
            '        .Items("Cancel").Visible = False
            '    End With
            'End If

            ' SplitContainer2.Panel2Collapsed = (EntryMode <> "V") Or InStr(ASCMAIN1.USER_SECURITY_CODEs, "WS") = 0
            If EntryMode <> "V" Then
                SplitContainer2.Panel2Collapsed = True
            Else
                SplitContainer2.Panel2Collapsed = False
                tabDetails.Tabs("GL Distribution").Visible = (EntryMode = "V")
                tabDetails.Tabs("Components Consumed").Visible = (EntryMode = "V")
            End If

            tabDetails.Tabs("Open POs").Visible = (EntryMode = "N") And (rowICTIREC1.Item("PO_ORDER_NO") & "" <> "")

            Set_Read_Only(grpHeader, (EntryMode = "V"))

            ' Nathan does not want this option - he thinks it is a possible control issue
            'If EntryMode = "N" Then
            '    Set_Read_Only_for_ctl(Absx1.txtFor("WHSE_CODE"), False)
            'End If

            Set_Read_Only(SplitContainer2, (EntryMode = "V"))
            If EntryMode = "N" Then
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTIREC2}
                    With grd.DisplayLayout.Override
                        '.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.True
                        .AllowUpdate = DefaultableBoolean.True
                    End With
                Next
                With grdICTIREC2.DisplayLayout.Bands(0)
                    .Columns("LOCATION_CODE").CellAppearance.BackColor = Color.Yellow
                    .Columns("QTY_REC").CellAppearance.BackColor = Color.Yellow
                    .Columns("REC_REF").CellAppearance.BackColor = Color.Yellow

                    .Columns("TRAN_PV").Hidden = True
                    .Columns("TRAN_MV").Hidden = True
                    .Columns("TRAN_CV").Hidden = True
                    .Columns("TRAN_FV").Hidden = True
                    .Columns("TRAN_TV").Hidden = True

                    .Columns("ACCRUAL_STATUS").Hidden = True
                    .Columns("AMT_INV").Hidden = True
                    .Columns("QTY_REC_NOT_INV").Hidden = True
                End With

            Else
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTIREC2}
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    End With
                Next
                With grdICTIREC2.DisplayLayout.Bands(0)
                    .Columns("LOCATION_CODE").CellAppearance.BackColor = Color.Empty
                    .Columns("QTY_REC").CellAppearance.BackColor = Color.Empty
                    .Columns("REC_REF").CellAppearance.BackColor = Color.Empty

                    .Columns("TRAN_PV").Hidden = False
                    .Columns("TRAN_MV").Hidden = False
                    .Columns("TRAN_CV").Hidden = False
                    .Columns("TRAN_FV").Hidden = False
                    .Columns("TRAN_TV").Hidden = False

                    .Columns("ACCRUAL_STATUS").Hidden = False
                    .Columns("AMT_INV").Hidden = False
                    .Columns("QTY_REC_NOT_INV").Hidden = False
                End With
            End If
            If EntryMode = "N" Then
                Set_Read_Only_for_ctl(Absx1.txtFor("SOURCE_DOC_NO"), False)
                Set_Read_Only_for_ctl(Absx1.dteFor("RECEIPT_DATE"), False)
            End If

            If Absx1.txtFor("VEND_CODE").Text = "IPSA" And (EntryMode <> "N") Then
                Variance_Column_Visibility(True)
            End If

        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        ' SR-6549 - Lot Numbers on Shipments and Receipts
        For Each TABLE_NAME As String In New String() {"ICTIREC0", "ICTIREC1", "ICTIREC2", "ICTIREC3", "ICTIREC4", "ICTIRECL",
                                                       "POTORDR1", "POTORDR2", "POTORDR9", "ICTSTAT2", "EDTTRXN1",
                                                       "POTORDR1_3PL", "POTORDR2_3PL", "ICTPINV1", "ICTPINV2",
                                                       "ICTIADJ1", "ICTIADJ2", "TATALRT1",
                                                       "APTINVH1", "APTINVH2", "APTINVH5", "APTINVH5_VAR", "APTVEND5"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Documents()

        Absx1.txtFor("WHSE_CODE").Text = ""
        Absx1.dteFor("RECEIPT_DATE").Value = Format(Now, "MM/dd/yyyy")
        Absx1.txtFor("RECEIPT_NO").Text = ""

        lblVARIANCE_EXPLANATION.Visible = False
        txtVARIANCE_EXPLANATION.Visible = False
        Variance_Column_Visibility(False)

        PO_ORDER_NO_RECEIVED = String.Empty
        PO_ORDER_NO_3PL = String.Empty
        InvoiceNum = String.Empty
        TRANS_NUM_3PL = ""
        rowICT3PLTX = Nothing

        optGL.Tag = ""
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            rowICTIREC1 = dst.Tables("ICTIREC1").NewRow

            With rowICTIREC1
                If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                    .Item("RECEIPT_NO") = ASCMAIN1.Next_Control_No("TRAN_NO_A")
                Else
                    .Item("RECEIPT_NO") = ASCMAIN1.Next_Control_No("ICTIREC1.RECEIPT_NO")
                End If
                .Item("WHSE_CODE") = HFs("WHSE_CODE")
                .Item("VEND_CODE") = HFs("VEND_CODE")
                .Item("RECEIPT_DATE") = HFs("RECEIPT_DATE")

                '.Item("RECEIPT_SOURCE") = "E"
                If TRANS_NUM_3PL <> "" Then
                    rowICTIREC1.Item("RECEIPT_SOURCE") = "T"
                    rowICTIREC1.Item("TRX_NO") = TRANS_NUM_3PL
                End If

                .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_DATE") = DATETIME_STAMP
                .Item("REGISTER_IND") = "0"
                .Item("JOURNAL_IND") = "0"
                .Item("ACCRUAL_STATUS") = "0"

            End With

            dst.Tables("ICTIREC1").Rows.Add(rowICTIREC1)
        Else
            rowICTIREC1 = Fill_Record("ICTIREC1", Absx1.txtFor("RECEIPT_NO").Text)
            dst.AcceptChanges()

            With dst.Tables("ICTIREC0").Rows
                .Add(New String() {"Entered", Format(rowICTIREC1.Item("INIT_DATE"), "MM/dd/yy hh:mm tt")})
                .Add(New String() {"By", rowICTIREC1.Item("INIT_OPER")})
                .Add(New String() {"Source", rowICTIREC1.Item("SOURCE_DOC_NO") & ""})
                If rowICTIREC1.Item("ACCRUAL_STATUS") & "" = "1" Then
                    .Add(New String() {"Status", "Invoiced"})
                    ' .Add(New String() {"Voucher", rowICTIREC1.Item("VOUCHER_NO") & ""})
                    .Add(New String() {"PO No", rowICTIREC1.Item("PO_ORDER_NO") & ""})
                Else
                    .Add(New String() {"Status", "Open"})
                    .Add(New String() {"PO No", rowICTIREC1.Item("PO_ORDER_NO") & ""})
                End If
                If rowICTIREC1.Item("REVERSED_BY_RECEIPT_NO") & "" <> "" Then
                    Dim row As DataRow = LookUp("ICTIREC1", rowICTIREC1.Item("REVERSED_BY_RECEIPT_NO"))
                    .Add(New String() {"Reversed", Format(rowICTIREC1.Item("LAST_DATE"), "MM/dd/yy hh:mm tt")})
                    .Add(New String() {"By", rowICTIREC1.Item("LAST_OPER")})
                    .Add(New String() {"using", rowICTIREC1.Item("REVERSED_BY_RECEIPT_NO")})
                ElseIf rowICTIREC1.Item("REVERSES_RECEIPT_NO") & "" <> "" Then
                    .Add(New String() {"Reverses", rowICTIREC1.Item("REVERSES_RECEIPT_NO")})
                End If
            End With
        End If

        rowICTWHSE1 = LookUp("ICTWHSE1", rowICTIREC1.Item("WHSE_CODE"))
        location_support = (rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1")
        With grdICTIREC2.DisplayLayout.Bands(0)
            .Columns("BAR_CODE").Hidden = True ' Not location_support
            .Columns("LOCATION_CODE").Hidden = Not location_support
        End With
        '   dst.Tables("ICTIREC2").Rows.Clear()

        Fill_Records("ICTIREC2", Absx1.txtFor("RECEIPT_NO").Text)
        Fill_Records("ICTIREC3", Absx1.txtFor("RECEIPT_NO").Text)
        Fill_Records("ICTIREC4", Absx1.txtFor("RECEIPT_NO").Text)
        ' SR-6549 - Lot Numbers on Shipments and Receipts
        Fill_Records("ICTIRECL", Absx1.txtFor("RECEIPT_NO").Text)

        ' PO_ORDER_NO_RECEIVED has data if you use the grid to receive entire PO

        If PO_ORDER_NO_RECEIVED <> "" Then
            rowICTIREC1.Item("PO_ORDER_NO") = PO_ORDER_NO_RECEIVED
            If TRANS_NUM_3PL.StartsWith("CLARINS_ASSY") Then
                rowICTIREC1.Item("RECEIPT_SOURCE") = "A"

                Dim DTADJC As String = rowICT3PLTX.Item("DTADJC")

                If DTADJC = "A03" Then ' Assembly
                    ASCMAIN1.sql = "Select * from POTORDR2" & vbCrLf _
                        & " where PO_ORDER_NO = '" & PO_ORDER_NO_RECEIVED & "'" & vbCrLf _
                        & "   and PO_QTY_OPN <> 0" & vbCrLf _
                        & "   and ITEM_CODE = '" & rowICT3PLTX.Item("ITEM_CODE") & "'" & vbCrLf _
                        & "   and ROWNUM <=1"
                    For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "PO_ORDER_LNO")
                        Receive_PO_Line(row, Val(rowICT3PLTX.Item("DTTQTY_POS") & ""))
                    Next

                ElseIf DTADJC = "A21" Then ' Dis-Assembly
                    For Each rowICT3PLTY As DataRow In rowICT3PLTX.GetChildRows("ICT3PLTX_ICT3PLTY")
                        Dim ITEM_CODE As String = rowICT3PLTY.Item("ITEM_CODE")

                        ASCMAIN1.sql = "Select * from POTORDR2" & vbCrLf _
                            & " where PO_ORDER_NO = '" & PO_ORDER_NO_RECEIVED & "'" & vbCrLf _
                            & "   and PO_QTY_OPN <> 0" & vbCrLf _
                            & $"   and ITEM_CODE = '{ITEM_CODE}'" & vbCrLf _
                            & "   and ROWNUM <=1"

                        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "PO_ORDER_LNO")
                            Debug.Print(ITEM_CODE & ":" & CStr(rowICT3PLTY.Item("DTTQTY")))
                            Receive_PO_Line(row, Val(rowICT3PLTY.Item("DTTQTY") & ""))
                        Next

                    Next

                Else
                    ' UNKNOWN REASON CODE

                End If
                ASCMAIN1.sql = "Select * from POTORDR2" & vbCrLf _
                    & " where PO_ORDER_NO = '" & PO_ORDER_NO_RECEIVED & "'" & vbCrLf _
                    & "   and PO_QTY_OPN <> 0" & vbCrLf _
                    & "   and ITEM_CODE = '" & rowICT3PLTX.Item("ITEM_CODE") & "'" & vbCrLf _
                    & "   and ROWNUM <=1"
                For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "PO_ORDER_LNO")
                    Receive_PO_Line(row, Val(rowICT3PLTX.Item("DTTQTY_POS") & ""))
                Next
                Sort_grdColumns(grdICTIREC2, "RECEIPT_LNO")
            Else
                ASCMAIN1.sql = "Select * from POTORDR2 where PO_ORDER_NO = '" & PO_ORDER_NO_RECEIVED & "' and PO_QTY_OPN <> 0"
                For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "PO_ORDER_LNO")
                    Receive_PO_Line(row)
                Next
                Sort_grdColumns(grdICTIREC2, "RECEIPT_LNO")
                'Absx1.txtFor("VEND_CODE").Text = grdPOTORDRO.ActiveRow.Cells("VEND_CODE").Value
                'Absx1.txtFor("WHSE_CODE").Text = grdPOTORDRO.ActiveRow.Cells("WHSE_CODE").Value
            End If
        End If

        ' PO_ORDER_NO_3PL has a value if you double click grdEDTTRXNX to receive from a 3PL receipt batch of data

        If PO_ORDER_NO_3PL <> "" Then
            dst.Tables("POTORDR2_3PL_ERR").Rows.Clear()
            dst.Tables("POTORDR1_3PL").Rows.Clear()
            dst.Tables("POTORDR2_3PL").Rows.Clear()
            rowICTIREC1.Item("PO_ORDER_NO") = PO_ORDER_NO_3PL

            If InvoiceNum.Length = 0 Then
                InvoiceNum = "?"
            End If

            ASCMAIN1.sql = $"Select * from EDTTRXN1 where PO_ORDER_NO = '{PO_ORDER_NO_3PL}' and TRANS_NUM = '{TRANS_NUM_3PL}' and TRAN_QTY <> 0 and NVL(PROCESS_IND, '0') = '0' AND NVL(INV_NUM, '?') = '{InvoiceNum}'"
            Fill_Records("EDTTRXN1", String.Empty, True, ASCMAIN1.sql)

            If InvoiceNum = "?" Then
                InvoiceNum = String.Empty
            End If

            ASCMAIN1.sql = "Select * from POTORDR1 where PO_ORDER_NO = '" & PO_ORDER_NO_3PL & "'"
            Fill_Records("POTORDR1_3PL", "", True, ASCMAIN1.sql)
            Dim rowPOTORDR1_3PL As DataRow = dst.Tables("POTORDR1_3PL").Rows(0)

            ASCMAIN1.sql = "Select * from POTORDR2 where PO_ORDER_NO = '" & PO_ORDER_NO_3PL & "'"
            Fill_Records("POTORDR2_3PL", "", True, ASCMAIN1.sql)

            Dim LOCATION As String = "" ' FOR INT, THIS FIELD CONTAINS THE IPSA INVOICE NO
            Dim PINV_NO As String = ""
            dst.Tables("ICTPINV2").Rows.Clear()

            ' Loop through all items in the Receipt provided by the Warehouse
            For Each rowEDTTRXN1 As DataRow In dst.Tables("EDTTRXN1").Select("", "ITEM_CODE")
                Dim TRAN_QTY As Int32 = rowEDTTRXN1.Item("TRAN_QTY")
                Dim ITEM_CODE As String = rowEDTTRXN1.Item("ITEM_CODE")
                Dim entriesFound As Boolean = False
                Dim numRows As Int32 = 0
                Dim processedRows As Int32 = 1
                rowEDTTRXN1.Item("PROCESS_IND") = "1"

                If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                    rowICTIREC1.Item("SOURCE_DOC_NO") = rowEDTTRXN1.Item("LOCATION")
                    If LOCATION = "" Then
                        LOCATION = rowEDTTRXN1.Item("LOCATION") & ""
                        ASCMAIN1.sql = "SELECT * FROM ICTPINV1 WHERE INV_NUM LIKE '%" & LOCATION & "'"
                        Dim rowI As DataRow = ASCDATA1.GetDataRow
                        If rowI IsNot Nothing Then
                            PINV_NO = rowI.Item("PINV_NO")
                            Fill_Records("ICTPINV2", PINV_NO)
                        End If
                    End If
                Else
                    rowICTIREC1.Item("SOURCE_DOC_NO") = rowEDTTRXN1.Item("TRANS_NUM")
                End If

                If TRAN_QTY <= 0 Then
                    Continue For
                End If

                ' Distribute Quantity evenly. There may be multiple lines for one Item Code
                processedRows = 0
                ' Get number of PO LInes to spread the Trans Quantity
                Dim sql_LNO As String = ""
                If ASCMAIN1.CLIENT = "INT" Then
                    'If ITEM_CODE = "MB013A03" Then Stop
                    Dim sqlIPSA As String = "ITEM_CODE = '" & ITEM_CODE & "' and ISNULL(PINV_QTY,0) - ISNULL(QTY_REC,0) = " & CStr(TRAN_QTY)
                    Dim rowICTPINV2s() As DataRow = dst.Tables("ICTPINV2").Select(sqlIPSA)
                    If rowICTPINV2s.Length = 0 Then
                        sqlIPSA = "ITEM_CODE = '" & ITEM_CODE & "' and ISNULL(PINV_QTY,0) - ISNULL(QTY_REC,0) >= " & CStr(TRAN_QTY)
                        rowICTPINV2s = dst.Tables("ICTPINV2").Select(sqlIPSA)
                    End If
                    If rowICTPINV2s.Length >= 1 Then
                        sql_LNO = " and PO_ORDER_LNO = " & rowICTPINV2s(0).Item("PO_ORDER_LNO")
                        rowICTPINV2s(0).Item("QTY_REC") = Val(rowICTPINV2s(0).Item("QTY_REC") & "") + TRAN_QTY
                    End If
                End If

                Dim sqlF As String = "ITEM_CODE = '" & ITEM_CODE & "' and PO_QTY_OPN <> 0" & sql_LNO
                numRows = dst.Tables("POTORDR2_3PL").Select(sqlF).Length

                For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2_3PL").Select(sqlF)
                    entriesFound = True
                    processedRows += 1

                    Dim PO_QTY_OPN As Int32 = rowPOTORDR2.Item("PO_QTY_OPN")

                    ' SR-6549 - Lot Numbers on Shipments and Receipts
                    rowPOTORDR2.Item("TRX_NO") = rowEDTTRXN1.Item("TRX_NO")
                    rowPOTORDR2.Item("TRX_LNO") = rowEDTTRXN1.Item("TRX_LNO")
                    rowPOTORDR2.Item("LOT_NO") = rowEDTTRXN1.Item("LOT_NO")

                    ' If we are processing the last/only row then apply entire quantity to this row
                    If numRows = processedRows Then
                        ' Everything received on this line
                        Receive_PO_Line(rowPOTORDR2, TRAN_QTY)
                        TRAN_QTY = 0
                    Else
                        ' Receive only the Quantity Open
                        Receive_PO_Line(rowPOTORDR2)
                        TRAN_QTY -= PO_QTY_OPN
                    End If

                    If TRAN_QTY <= 0 Then Exit For
                Next

                ' Do not Add additional items
                If TRAN_QTY > 0 Then
                    Dim ITEM_DESC As String = String.Empty
                    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                    If rowICTITEM1 IsNot Nothing Then
                        ITEM_DESC = rowICTITEM1.Item("ITEM_DESC") & String.Empty
                    End If
                    dst.Tables("POTORDR2_3PL_ERR").Rows.Add(New Object() {TRAN_QTY, ITEM_CODE, ITEM_DESC})
                End If

            Next
            Sort_grdColumns(grdICTIREC2, "RECEIPT_LNO")
            If dst.Tables("POTORDR2_3PL_ERR").Rows.Count > 0 Then
                Dim f As New ASFMSGBF
                f.Show_grd(dst.Tables("POTORDR2_3PL_ERR"), Me, "Items Closed or Not on PO", "")
                f = Nothing
            End If

            dst.Tables("ICTPINV2").Rows.Clear()

            Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO_3PL)
            Dim WHSE_CODE_PO As String = rowPOTORDR1.Item("WHSE_CODE")
            Dim WHSE_CODE_SOURCE As String = "PO"
            If PINV_NO <> "" Then
                Dim rowICTPINV1 As DataRow = LookUp("ICTPINV1", PINV_NO)
                WHSE_CODE_PO = rowICTPINV1.Item("WHSE_CODE")
                WHSE_CODE_SOURCE = "Invoice"
            End If
            If rowICTIREC1.Item("WHSE_CODE") & "" <> WHSE_CODE_PO Then
                MsgBox($"3PL Record indicates Whse {rowICTIREC1.Item("WHSE_CODE")}, but {WHSE_CODE_SOURCE} indicates Whse {WHSE_CODE_PO}" & vbCrLf & vbCrLf & $"Switching to {WHSE_CODE_PO}", MsgBoxStyle.OkOnly, "Receiving Warehouse Code was Changed")
                rowICTIREC1.Item("WHSE_CODE") = WHSE_CODE_PO
            End If

        End If

        If PO_ORDER_NO_3PL.Length > 0 Then
            If InvoiceNum.Length > 0 Then
                rowICTIREC1.Item("SOURCE_DOC_NO") = InvoiceNum
            End If
        End If

        If EntryMode = "N" Then
            ' I am not sure of the best place for this code, as there are several paths with EntryMode = N
            ' so I am electing to do this after all the smoke clears
            With rowICTIREC1
                Dim PO_ORDER_NO As String = .Item("PO_ORDER_NO")
                Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)

                Dim FRT_CLASS_CODE As String = rowPOTORDR1.Item("FRT_CLASS_CODE")
                Dim TRF_CLASS_CODE As String = rowPOTORDR1.Item("TRF_CLASS_CODE")

                .Item("FRT_CLASS_CODE") = FRT_CLASS_CODE
                .Item("TRF_CLASS_CODE") = TRF_CLASS_CODE

                Dim rowICTFRTC1 As DataRow = dst.Tables("ICTFRTC1").Rows.Find(FRT_CLASS_CODE)
                Dim FRT_CLASS_PCT As Decimal = Val(rowICTFRTC1.Item("FRT_CLASS_PCT_CUR") & "")
                .Item("FRT_CLASS_PCT") = FRT_CLASS_PCT
                Dim rowICTTRFC1 As DataRow = dst.Tables("ICTTRFC1").Rows.Find(TRF_CLASS_CODE)
                Dim TRF_CLASS_PCT As Decimal = Val(rowICTTRFC1.Item("TRF_CLASS_PCT_CUR") & "")
                .Item("TRF_CLASS_PCT") = TRF_CLASS_PCT

                For Each rowICTIREC2 As DataRow In dst.Tables("ICTIREC2").Select("", "RECEIPT_LNO")
                    With rowICTIREC2
                        Dim PO_COST As Decimal = Val(.Item("PO_COST") & "")
                        .Item("PO_COST_FRT") = PO_COST * FRT_CLASS_PCT / 100
                        .Item("PO_COST_TRF") = PO_COST * TRF_CLASS_PCT / 100
                    End With
                Next
            End With
        End If


        ' As Per Nathan - 05/08/2025 Provide a message when there are missing lot codes
        ' ADS to send Lot Nos starting June 1, 2025

        ' ISSUE-7230 ADS as the defaUlt warehouse
        Dim drICTWHSE1 As DataRow = dst.Tables("ICTWHSE1").Rows.Find(HFs("WHSE_CODE"))
        Dim isADS3PL As Boolean = drICTWHSE1 IsNot Nothing AndAlso drICTWHSE1.Item("LP_CODE") & String.Empty = "ADS"

        If EntryMode = "N" AndAlso isADS3PL AndAlso Val(DateTime.Now.ToString("yyyyMMdd")) >= Val("20250601") Then
            Dim lotMsg As New List(Of String)
            For Each rowICTIREC2 As DataRow In dst.Tables("ICTIREC2").Select("ISNULL(QTY_REC,0) <> ISNULL(QTY_REC_LOT,0)", "RECEIPT_LNO")
                Dim QTY_REC As Int32 = Val(rowICTIREC2.Item("QTY_REC") & String.Empty)
                Dim QTY_REC_LOT As Int32 = Val(rowICTIREC2.Item("QTY_REC_LOT") & String.Empty)
                Dim ITEM_CODE As String = rowICTIREC2.Item("ITEM_CODE") & String.Empty
                lotMsg.Add($"Item {ITEM_CODE} has a received quantity of {QTY_REC}; however, {QTY_REC_LOT} have Lot numbers.")
            Next

            If lotMsg.Count > 0 Then
                MessageBox.Show(String.Join(Environment.NewLine, lotMsg.ToArray), "Lot Numbers", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Try
            BeginTrans()

            If PO_ORDER_NO_3PL <> "" Then

                If TRANS_NUM_3PL.StartsWith("CLARINS_ASSY") Then
                    Dim DTDATE As String = Split(TRANS_NUM_3PL, vbTab)(1)
                    Dim DTTIME As String = Split(TRANS_NUM_3PL, vbTab)(2)
                    Dim DTUSER As String = Split(TRANS_NUM_3PL, vbTab)(3)
                    If Not HFs.ContainsKey("CLARINS_ASSY") Then
                        HFs.Add("CLARINS_ASSY", "")
                    End If
                    HFs("CLARINS_ASSY") = TRANS_NUM_3PL
                    ' Dim PROC_KEY As String = Split(TRANS_NUM_3PL, vbTab)(4)
                    ASCMAIN1.sql = "Update EDTTRXNA" & vbCrLf _
                    & " Set PROCESSED_IND = '1', PROCESSED_DATE = SYSDATE" & vbCrLf _
                    & " where DTDATE = '" & DTDATE & "' and DTTIME = '" & DTTIME & "' and DTUSER = '" & DTUSER & "'"
                    ASCDATA1.ExecuteSQL()

                Else
                    For Each row As DataRow In dst.Tables("POTORDR2_3PL").Select("")
                        If row.RowState <> DataRowState.Added Then
                            row.Delete()
                        End If
                    Next

                    dst.Tables("POTORDR2_3PL").AcceptChanges()
                    For Each row As DataRow In dst.Tables("POTORDR2_3PL").Select("")
                        row.SetAdded()
                    Next

                    Update_Record_TDA("EDTTRXN1")
                    Update_Record_TDA("POTORDR2_3PL")
                End If

            End If

            If reversal_update Then
                Set_Up_Reversal()

                If rowICTIREC1.Item("RECEIPT_SOURCE") & "" = "T" And rowICTIREC1.Item("TRX_NO") & "" <> "" Then
                    ASCMAIN1.sql = "Update EDTTRXN1 SET PROCESS_IND = '0' where TRX_NO = '" & rowICTIREC1.Item("TRX_NO") & "'"
                    ASCDATA1.ExecuteSQL()
                Else
                    ' Assy A03, Dis-assembly A21

                    Dim TRX_NO As String = rowICTIREC1.Item("TRX_NO") & ""
                    If TRX_NO <> "" Then
                        Dim DTDATE As String = Split(TRX_NO, vbTab)(1)
                        Dim DTTIME As String = Split(TRX_NO, vbTab)(2)
                        Dim DTUSER As String = Split(TRX_NO, vbTab)(3)
                        If Not HFs.ContainsKey("CLARINS_ASSY") Then
                            HFs.Add("CLARINS_ASSY", "")
                        End If
                        HFs("CLARINS_ASSY") = TRX_NO
                        ' Dim PROC_KEY As String = Split(TRANS_NUM_3PL, vbTab)(4)
                        ASCMAIN1.sql = "Update EDTTRXNA" & vbCrLf _
                        & " Set PROCESSED_IND = NULL" & vbCrLf _
                        & " where DTDATE = '" & DTDATE & "' and DTTIME = '" & DTTIME & "' and DTUSER = '" & DTUSER & "'"
                        ASCDATA1.ExecuteSQL()
                    End If
                End If
            End If

            ICCMAIN1.Update_Receipt(Me, reversal_update)

            ' POTENTIAL ISSUE OF THE VEND_WHSE_CODE NEEDS LOCATION SUPPORT BUT WHSE_CODE DOES NOT
            If location_support Then
                Update_WHTLOCBX()
            End If

            If dst.Tables("ICTPINV1").Rows.Count > 0 Then
                Record_AP_Invoice()
                'SOCMAIN1.Record_AP_Invoice(Me, rowICTIREC1, reversal_update)
            End If

            ' Record Accruals - these calls were moved out of ICCMAIN1.Update Receipt
            '  because there is a dependency on ICTPINV1.BOL_NO  which is updated in Record_AP_Invoice
            For Each rowICTIREC1 As DataRow In dst.Tables("ICTIREC1").Select("")
                With rowICTIREC1
                    Dim RECEIPT_NO_in As String = .Item("RECEIPT_NO")
                    TAC.ICCMAIN1.Create_Accrual_FRT(RECEIPT_NO_in)
                    TAC.ICCMAIN1.Create_Accrual_TRF(RECEIPT_NO_in)
                End With
            Next

            CommitTrans("Update Complete")

        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

    Sub Record_AP_Invoice()

        Dim VOUCHER_NO As String = ASCMAIN1.Next_Control_No("APTINVH1.VOUCHER_NO")
        Dim RECEIPT_NO As String = rowICTIREC1.Item("RECEIPT_NO")
        Dim VEND_CODE As String = rowICTIREC1.Item("VEND_CODE")
        Dim rowICTPINV1 As DataRow = dst.Tables("ICTPINV1").Rows(0)
        Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", VEND_CODE)

        If reversal_update Then
            rowICTPINV1.Item("PINV_STATUS") = "O"
            rowICTPINV1.Item("VOUCHER_NO") = DBNull.Value
            rowICTPINV1.Item("RECEIPT_NO") = DBNull.Value
        Else
            rowICTPINV1.Item("PINV_STATUS") = "I"
            rowICTPINV1.Item("VOUCHER_NO") = VOUCHER_NO
            rowICTPINV1.Item("RECEIPT_NO") = RECEIPT_NO
        End If


        Dim INV_AMT As Decimal = Val(dst.Tables("ICTPINV2").Compute("SUM(AMT_INV)", "") & "")
        If reversal_update Then INV_AMT = -1 * INV_AMT

        rowAPTINVH1 = dst.Tables("APTINVH1").NewRow
        With rowAPTINVH1
            .Item("VOUCHER_NO") = VOUCHER_NO
            .Item("VEND_CODE") = VEND_CODE
            .Item("INV_TYPE") = "I"
            .Item("INV_NUM") = rowICTPINV1.Item("INV_NUM")
            .Item("INV_DATE") = rowICTPINV1.Item("INV_DATE")

            .Item("INV_AMT") = INV_AMT
            .Item("INV_REF") = rowICTPINV1.Item("PINV_SOURCE_DOC")

            .Item("VEND_CODE_AP") = rowAPTVEND1.Item("VEND_CODE_AP")
            If rowAPTVEND1.Item("VEND_PYMT_ADDR") & "" = "" Then
                .Item("VEND_ALT_CODE") = ""
                .Item("INV_REMIT_TO") = "V"
            Else
                .Item("VEND_ALT_CODE") = rowAPTVEND1.Item("VEND_PYMT_ADDR")
                .Item("INV_REMIT_TO") = "A"
            End If

            .Item("INV_SEP_CHECK") = rowAPTVEND1.Item("VEND_SEP_CHECKS")
            .Item("TERM_CODE") = rowAPTVEND1.Item("TERM_CODE")

            If rowAPTVEND1.Item("BANK_CODE") & "" = "" Then
                .Item("BANK_CODE") = ROWs("APTPARM1").Item("AP_PARM_BANK_CODE")
            Else
                .Item("BANK_CODE") = rowAPTVEND1.Item("BANK_CODE")
            End If

            If rowAPTVEND1.Item("VEND_PYMT_METHOD") & "" = "" Then
                If .Item("BANK_CODE") & "" <> "" Then
                    Dim rowGLTBANK1 As DataRow = LookUp("GLTBANK1", .Item("BANK_CODE"))
                    .Item("INV_PYMT_METHOD") = rowGLTBANK1.Item("BANK_PYMT_METHOD")
                End If
            Else
                .Item("INV_PYMT_METHOD") = rowAPTVEND1.Item("VEND_PYMT_METHOD")
            End If

            .Item("INV_PYMT_CYCLE") = rowAPTVEND1.Item("VEND_PYMT_CYCLE")

            If rowAPTVEND1.Item("POST_CODE") & "" <> "" Then
                .Item("POST_CODE") = rowAPTVEND1.Item("POST_CODE")
            Else
                .Item("POST_CODE") = ROWs("APTPARM1").Item("AP_PARM_POST_CODE")
            End If

            .Item("INV_STATUS") = "O"
            .Item("INV_PYMT_CYCLE") = DBNull.Value
            .Item("INV_DUE_DATE") = TAC.TACMAIN1.Calculate_INV_DUE_DATE(Me, rowAPTVEND1.Item("TERM_CODE") & "", Nothing, .Item("INV_DATE"))
            .Item("INV_BALANCE") = INV_AMT
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
            .Item("CURR_EXCH_RATE") = 1

            .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")

            .Item("REGISTER_IND") = "0"
            .Item("INV_BL_DATE") = .Item("INV_DATE")
            .Item("INV_AMT_VEND") = .Item("INV_AMT")

            If rowAPTVEND1.Item("VEND_AUTO_APPROVE") & "" = "1" Then
                .Item("INV_APPR_STATUS") = "A"
                Write_Event_Log("APTINVH1", VOUCHER_NO, "Auto Approved")
            Else
                .Item("INV_APPR_STATUS") = "P"
            End If

            .Item("VEND_BUYER_CODE") = rowAPTVEND1.Item("VEND_BUYER_CODE")
        End With

        dst.Tables("APTINVH1").Rows.Add(rowAPTINVH1)

        ASCMAIN1.Record_Event("APTINVH1", VOUCHER_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "AUTO", "Auto Rec/Inv", RECEIPT_NO)

        ' ADO.NET BEHAVIOR USING CUSTOM SQL STATEMENTS:
        ' NOTE THAT THE FIELDS VAR_QTY AND VAR_AMT ARE NOT ACCOUNTED FOR IN THE SQL
        ' THEY SHOULD BE RIGHT BEFORE '0' VB
        ' YET THE FILL MATCHES BASED ON NAME SO THAT CB AND COST_CATGY_CODE GET THE VALUES COMING IN FROM THOSE COLUMN NAMES
        ' MEANING NAME MAPPING TRUMPS POSITIONAL MAPPING

        '            & ", DECODE(ICTIREC2.ACCRUAL_STATUS,'1',0,NVL(ICTIREC2.QTY_REC,0) - NVL(ICTIREC2.QTY_INV,0)) INV_QTY" & vbCrLf _
        '            & ", DECODE(ICTIREC2.ACCRUAL_STATUS,'1',0,NVL(ICTIREC2.QTY_REC,0) - NVL(ICTIREC2.QTY_INV,0)) QTY_REC_NOT_INV" & vbCrLf _

        Dim sql As String = "Select '" & VOUCHER_NO & "' VOUCHER_NO " & vbCrLf _
            & ", ICTIREC2.RECEIPT_LNO VOUCHER_DLNO " & vbCrLf _
            & ", ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO" & vbCrLf _
            & ", DECODE(ICTIREC2.ACCRUAL_STATUS,'1',0,0) INV_QTY" & vbCrLf _
            & ", ICTIREC2.PO_COST INV_COST" & vbCrLf _
            & ", '0' CB" & vbCrLf _
            & ", ICTITEM1.COST_CATGY_CODE, '0' CLOSE_LINE" & vbCrLf _
            & ", DECODE(ICTIREC2.ACCRUAL_STATUS,'1',0,NVL(ICTIREC2.QTY_REC,0) - NVL(ICTIREC2.QTY_INV,0)) QTY_REC_NOT_INV" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC, ICTITEM1.COLLECTION_CODE, ICTCOST1.ACCT_CODE_PPV" & vbCrLf _
            & ", ICTIREC2.QTY_REC, ICTIREC2.QTY_INV, ICTIREC2.PO_COST " & vbCrLf _
            & ", ICTIREC2.ITEM_CODE, ICTIREC2.ITEM_UOM" & vbCrLf _
            & " from ICTIREC2,ICTITEM1,ICTCOST1" & vbCrLf _
            & " where ICTIREC2.RECEIPT_NO = '" & RECEIPT_NO & "'" & vbCrLf _
            & " and ICTITEM1.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf _
            & " and ICTCOST1.COST_CATGY_CODE (+) = ICTITEM1.COST_CATGY_CODE"
        Fill_Records("APTINVH5", , False, sql)

        dst.Tables("APTINVH5").AcceptChanges()
        For Each rowAPTINVH5 As DataRow In dst.Tables("APTINVH5").Select("")
            rowAPTINVH5.SetAdded()
            Dim RECEIPT_LNO As Int64 = Val(rowAPTINVH5.Item("RECEIPT_LNO") & "")
            Dim rowICTIREC2 As DataRow = dst.Tables("ICTIREC2").Rows.Find(New Object() {RECEIPT_NO, RECEIPT_LNO})
            Dim rowICTPINV2 As DataRow = dst.Tables("ICTPINV2").Select("RECEIPT_LNO = " & CStr(RECEIPT_LNO))(0)
            Dim QTY_INV As Int64 = Val(rowICTPINV2.Item("QTY_INV") & "")

            If reversal_update Then
                QTY_INV = -1 * QTY_INV
            End If

            rowAPTINVH5.Item("INV_QTY") = QTY_INV
            rowAPTINVH5.Item("INV_COST") = rowICTPINV2.Item("PINV_COST")
            ' NEXT 2 LINES ADDED BELOW TO SUPPORT THEM - THEY WERE NULL IN ORACLE
            ' VARIANCE WAS BEING CALCULATED And TRANSACTED TO GL PROPERLY 
            ' BECAUSE EXTENDED FIELDS QTY_VAR AND AMT_VAR WERE BEING CALCULATED AND USED
            rowAPTINVH5.Item("VAR_QTY") = 0 ' Val(rowAPTINVH5.Item("INV_QTY") & "") - Val(rowICTIREC2.Item("QTY_REC") & "")
            rowAPTINVH5.Item("VAR_AMT") = (Val(rowAPTINVH5.Item("INV_COST") & "") - Val(rowICTIREC2.Item("PO_COST") & "")) * Val(rowAPTINVH5.Item("INV_QTY") & "")
        Next

        Create_APTINVH5_VAR()
        Create_APTINVH2_P()
        rowAPTVEND5 = Fill_Record("APTVEND5", VEND_CODE, True)

        Update_Record_TDA("APTINVH1")
        Update_Record_TDA("APTINVH2")
        Update_Record_TDA("APTINVH5")
        Update_Record_TDA("ICTPINV1")
        Update_Record_TDA("ICTPINV2")

        Dependent_Updates(VOUCHER_NO, False)

    End Sub

    Sub Dependent_Updates(ByVal VOUCHER_NO As String, ByVal reverse As Boolean)

        ASCMAIN1.sql = "" _
            & "Begin " & vbCrLf _
            & " Declare " & vbCrLf _
            & "  Cursor C1 is Select * from APTINVH5 where VOUCHER_NO = '" & VOUCHER_NO & "';" & vbCrLf _
            & " Begin " & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update ICTIREC2 Set " & vbCrLf _
            & IIf(reverse,
                  "QTY_INV = NVL(QTY_INV,0) - NVL(R1.INV_QTY,0)",
                  "QTY_INV = NVL(QTY_INV,0) + NVL(R1.INV_QTY,0)") & vbCrLf _
            & IIf(reverse,
                  ",AMT_INV = NVL(AMT_INV,0) - NVL(R1.INV_QTY,0) * NVL(R1.INV_COST,0)",
                  ",AMT_INV = NVL(AMT_INV,0) + NVL(R1.INV_QTY,0) * NVL(R1.INV_COST,0)") & vbCrLf _
            & IIf(reverse,
                  ", ACCRUAL_STATUS = '0'",
                  ", ACCRUAL_STATUS = NVL(R1.CLOSE_LINE,'0')") & vbCrLf _
            & "    where RECEIPT_NO = R1.RECEIPT_NO and RECEIPT_LNO = R1.RECEIPT_LNO;" & vbCrLf _
            & "   Update ICTIREC2 Set ACCRUAL_STATUS = '1'" & vbCrLf _
            & "    where RECEIPT_NO = R1.RECEIPT_NO and RECEIPT_LNO = R1.RECEIPT_LNO and NVL(QTY_REC,0) = NVL(QTY_INV,0);" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select Distinct RECEIPT_NO from APTINVH5 where VOUCHER_NO = '" & VOUCHER_NO & "'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select()

            Dim RECEIPT_NO As String = row.Item("RECEIPT_NO")

            ASCMAIN1.sql = "" _
                & "Begin " & vbCrLf _
                & " Declare Cursor C1 is " & vbCrLf _
                & "  Select RECEIPT_NO" & vbCrLf _
                & " , Min (ACCRUAL_STATUS) ACCRUAL_STATUS" & vbCrLf _
                & " , Sum (QTY_INV) QTY_INV, Sum (AMT_INV) AMT_INV" & vbCrLf _
                & "  from ICTIREC2 where RECEIPT_NO = '" & RECEIPT_NO & "' group by RECEIPT_NO;" & vbCrLf _
                & " Begin " & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update ICTIREC1 Set " & vbCrLf _
                & "    ACCRUAL_STATUS = R1.ACCRUAL_STATUS" & vbCrLf _
                & "    , QTY_INV = R1.QTY_INV" & vbCrLf _
                & "    , AMT_INV = R1.AMT_INV" & vbCrLf _
                & "    where RECEIPT_NO = R1.RECEIPT_NO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

            ASCDATA1.ExecuteSP("ICPIRECV", "V", New Object() {RECEIPT_NO}, New String() {"RECEIPT_NO_in"})
        Next

        Dim INV_AMT As Decimal = Val(rowAPTINVH1("INV_AMT") & "")
        If reverse Then
            Dim row As DataRow = LookUp("APTINVH1", VOUCHER_NO)
            INV_AMT = -1 * Val(row("INV_AMT") & "")
        End If
        ' rowAPTVEND5 = Fill_Record("APTVEND5", VEND_CODE, True)

        With rowAPTVEND5
            .Item("VEND_PURCHASES_MTD") = Val(.Item("VEND_PURCHASES_MTD") & "") + INV_AMT
            .Item("VEND_PURCHASES_YTD") = Val(.Item("VEND_PURCHASES_YTD") & "") + INV_AMT
            .Item("VEND_NUM_INV_MTD") = Val(.Item("VEND_NUM_INV_MTD") & "") + IIf(reverse, -1, 1)
            .Item("VEND_NUM_INV_YTD") = Val(.Item("VEND_NUM_INV_YTD") & "") + IIf(reverse, -1, 1)
            If Not reverse Then
                .Item("VEND_LAST_INV_DATE") = rowAPTINVH1("INV_DATE")
                .Item("VEND_LAST_INV_AMT") = rowAPTINVH1("INV_AMT")
                .Item("VEND_LAST_INV_NUM") = rowAPTINVH1("INV_NUM")
                If .Item("VEND_1ST_PURCH_DATE") & "" = "" Then
                    .Item("VEND_1ST_PURCH_DATE") = rowAPTINVH1("INV_DATE")
                End If
            End If

        End With
    End Sub

    Sub Create_APTINVH5_VAR()

        dst.Tables("APTINVH5_VAR").Rows.Clear()

        For Each rowAPTINVH5 As DataRow In dst.Tables("APTINVH5") _
        .Select("", "", DataViewRowState.CurrentRows)
            Dim COST_CATGY_CODE As String = rowAPTINVH5.Item("COST_CATGY_CODE")
            Dim rowICTCOST1 As DataRow = LookUp("ICTCOST1", COST_CATGY_CODE)

            Dim COLLECTION_CODE As String = rowAPTINVH5.Item("COLLECTION_CODE")
            Dim rowICTCOLL1 As DataRow = LookUp("ICTCOLL1", COLLECTION_CODE) '  dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE) ' LookUp("ICTCOLL1", COLLECTION_CODE)

            Dim SEG2_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            Dim SEG3_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            Dim SEG4_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")

            If ROWs("ICTPARM1").Item("IC_PARM_EXP_SEG4") & "" = "1" Then
                If rowICTCOLL1.Item("SEG4_CODE") & "" <> "" Then
                    SEG4_CODE = rowICTCOLL1.Item("SEG4_CODE")
                Else
                    SEG4_CODE = COLLECTION_CODE
                End If
            End If

            Dim rowAPTINVH5_VAR As DataRow = dst.Tables("APTINVH5_VAR").Rows.Find(New String() {COST_CATGY_CODE, COLLECTION_CODE})
            If rowAPTINVH5_VAR Is Nothing Then
                rowAPTINVH5_VAR = dst.Tables("APTINVH5_VAR").NewRow
                rowAPTINVH5_VAR.Item("COST_CATGY_CODE") = COST_CATGY_CODE
                rowAPTINVH5_VAR.Item("COLLECTION_CODE") = COLLECTION_CODE
                rowAPTINVH5_VAR.Item("ACCT_CODE_PPV") = rowAPTINVH5.Item("ACCT_CODE_PPV")
                Dim rowICTIREC1 As DataRow = dst.Tables("ICTIREC1").Rows.Find(rowAPTINVH5.Item("RECEIPT_NO"))
                rowAPTINVH5_VAR.Item("SEG2_CODE") = SEG2_CODE ' rowICTIREC1.Item("SEG2_CODE")
                rowAPTINVH5_VAR.Item("SEG3_CODE") = SEG3_CODE
                rowAPTINVH5_VAR.Item("SEG4_CODE") = SEG4_CODE
                dst.Tables("APTINVH5_VAR").Rows.Add(rowAPTINVH5_VAR)
            End If
            rowAPTINVH5_VAR.Item("AMT_REC") = Val(rowAPTINVH5_VAR.Item("AMT_REC") & "") + Val(rowAPTINVH5.Item("AMT_REC") & "")
            rowAPTINVH5_VAR.Item("AMT_INV") = Val(rowAPTINVH5_VAR.Item("AMT_INV") & "") + Val(rowAPTINVH5.Item("AMT_INV") & "")
            rowAPTINVH5_VAR.Item("AMT_REC_NOT_INV") = Val(rowAPTINVH5_VAR.Item("AMT_REC_NOT_INV") & "") + Val(rowAPTINVH5.Item("AMT_REC_NOT_INV") & "") ' Val(rowAPTINVH5.Item("QTY_REC_NOT_INV") & "") * Val(rowAPTINVH5.Item("PO_COST") & "")
            rowAPTINVH5_VAR.Item("AMT_REC_NOT_INV_OFFSET") = Val(rowAPTINVH5_VAR.Item("AMT_REC_NOT_INV_OFFSET") & "") + Val(rowAPTINVH5.Item("AMT_REC_NOT_INV_OFFSET") & "")
            rowAPTINVH5_VAR.Item("AMT_VAR") = Val(rowAPTINVH5_VAR.Item("AMT_VAR") & "") + Val(rowAPTINVH5.Item("AMT_VAR") & "")
            If rowAPTINVH5.Item("CB") & "" = "1" Then
                rowAPTINVH5_VAR.Item("AMT_VAR_CB") = Val(rowAPTINVH5_VAR.Item("AMT_VAR_CB") & "") + Val(rowAPTINVH5.Item("AMT_VAR") & "")
            End If
        Next

    End Sub

    Sub Create_APTINVH2_P()

        Delete_Rows("APTINVH2", "INV_LTYP = 'P'")

        Dim VOUCHER_NO As String = rowAPTINVH1.Item("VOUCHER_NO")
        Dim VOUCHER_LNO_ctr As Int32 = Val(dst.Tables("APTINVH2").Compute("MAX(VOUCHER_LNO)", "VOUCHER_NO = '" & VOUCHER_NO & "'") & "")

        Dim INV_LINE_AMT As Decimal = Val(dst.Tables("APTINVH5_VAR").Compute("SUM(AMT_REC_NOT_INV_OFFSET)", "") & "")
        Dim rowAPTINVH2 As DataRow

        If INV_LINE_AMT <> 0 Then
            rowAPTINVH2 = dst.Tables("APTINVH2").NewRow
            With rowAPTINVH2
                .Item("VOUCHER_NO") = VOUCHER_NO
                VOUCHER_LNO_ctr += 1
                .Item("VOUCHER_LNO") = VOUCHER_LNO_ctr
                .Item("ACCT_CODE") = ROWs("ICTPARM1").Item("IC_PARM_ACCT_CODE_ACCR_PURCH")
                .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                .Item("INV_LINE_AMT") = INV_LINE_AMT
                .Item("INV_LTYP") = "P"
            End With
            dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)
        End If

        For Each rowAPTINVH5_VAR As DataRow In dst.Tables("APTINVH5_VAR").Select("ISNULL(AMT_VAR,0) - ISNULL(AMT_VAR_CB,0) <> 0")
            ' INV_LINE_AMT = -1 * (Val(rowAPTINVH5_VAR.Item("AMT_VAR") & "") - Val(rowAPTINVH5_VAR.Item("AMT_VAR_CB") & ""))
            INV_LINE_AMT = (Val(rowAPTINVH5_VAR.Item("AMT_VAR") & "") - Val(rowAPTINVH5_VAR.Item("AMT_VAR_CB") & ""))
            If INV_LINE_AMT <> 0 Then
                rowAPTINVH2 = dst.Tables("APTINVH2").NewRow
                With rowAPTINVH2
                    .Item("VOUCHER_NO") = VOUCHER_NO
                    VOUCHER_LNO_ctr += 1
                    .Item("VOUCHER_LNO") = VOUCHER_LNO_ctr
                    .Item("ACCT_CODE") = rowAPTINVH5_VAR.Item("ACCT_CODE_PPV")
                    .Item("SEG2_CODE") = rowAPTINVH5_VAR.Item("SEG2_CODE")
                    .Item("SEG3_CODE") = rowAPTINVH5_VAR.Item("SEG3_CODE")
                    .Item("SEG4_CODE") = rowAPTINVH5_VAR.Item("SEG4_CODE")
                    .Item("INV_LINE_AMT") = INV_LINE_AMT
                    .Item("INV_LTYP") = "P"
                End With
                dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)
            End If
        Next
    End Sub

    Sub Update_WHTLOCBX()
        Dim rowICTIREC1 As DataRow = dst.Tables("ICTIREC1").Rows(0)
        TAC.ICCMAIN1.Update_WHTLOCBX("R", rowICTIREC1.Item("RECEIPT_NO"))
    End Sub

    Sub Update_WHTLOCBX_OLD()
        If dst.Tables.Contains("WHTLOCB1") Then
            dst.Tables("WHTLOCB1").Rows.Clear()
            dst.Tables("WHTLOCB2").Rows.Clear()
        Else
            Create_TDA(dst.Tables.Add, "WHTLOCB1", "*")
            Create_TDA(dst.Tables.Add, "WHTLOCB2", "*")
        End If

        Dim rowICTIREC1 As DataRow = dst.Tables("ICTIREC1").Rows(0)

        For Each row As DataRow In dst.Tables("ICTIREC2").Select("")
            Dim TRAN_NO As String = row.Item("RECEIPT_NO")
            Dim TRAN_LNO As Integer = row.Item("RECEIPT_LNO")
            Dim WHSE_CODE As String = rowICTIREC1.Item("WHSE_CODE")
            Dim BAR_CODE As String = "0000000000" ' row.Item("BAR_CODE")
            Dim LOCATION_CODE As String = row.Item("LOCATION_CODE")
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim QTY_REC As Int64 = Val(row.Item("QTY_REC") & "")

            Dim rowWHTLOCB1 As DataRow = dst.Tables("WHTLOCB1").Rows.Find(New Object() _
                                         {WHSE_CODE, LOCATION_CODE, BAR_CODE, ITEM_CODE})
            If rowWHTLOCB1 Is Nothing Then
                Fill_Records("WHTLOCB1", New String() {WHSE_CODE, LOCATION_CODE, BAR_CODE, ITEM_CODE}, False)
                rowWHTLOCB1 = dst.Tables("WHTLOCB1").Rows.Find(New Object() _
                                         {WHSE_CODE, LOCATION_CODE, BAR_CODE, ITEM_CODE})
            End If

            If rowWHTLOCB1 Is Nothing Then
                rowWHTLOCB1 = dst.Tables("WHTLOCB1").NewRow
                With rowWHTLOCB1
                    .Item("WHSE_CODE") = WHSE_CODE
                    .Item("LOCATION_CODE") = LOCATION_CODE
                    .Item("BAR_CODE") = BAR_CODE
                    .Item("ITEM_CODE") = ITEM_CODE
                    .Item("LOCATION_QTY") = QTY_REC
                End With
                dst.Tables("WHTLOCB1").Rows.Add(rowWHTLOCB1)
            Else
                rowWHTLOCB1.Item("LOCATION_QTY") = Val(rowWHTLOCB1.Item("LOCATION_QTY") & "") + QTY_REC
            End If

            Dim rowWHTLOCB2 As DataRow = dst.Tables("WHTLOCB2").NewRow
            With rowWHTLOCB2
                .Item("WHSE_CODE") = WHSE_CODE
                .Item("LOCATION_CODE") = LOCATION_CODE
                .Item("BAR_CODE") = BAR_CODE
                .Item("ITEM_CODE") = ITEM_CODE
                .Item("WHSE_TRAN_QTY") = QTY_REC
                .Item("WHSE_TRAN_TYPE") = "A"
                .Item("WHSE_TRAN_NO") = TRAN_NO
                .Item("WHSE_TRAN_LNO") = TRAN_LNO
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LOCATION_CODE_OTHER") = ""
                .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
            End With
            dst.Tables("WHTLOCB2").Rows.Add(rowWHTLOCB2)
        Next

        Update_Record_TDA("WHTLOCB1")
        Update_Record_TDA("WHTLOCB2")

        dst.Tables("WHTLOCB1").Rows.Clear()
        dst.Tables("WHTLOCB2").Rows.Clear()
    End Sub

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

    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

            Case "View"
                Absx1.txtFor("RECEIPT_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            'Case "VEND_CODE"
            '    sql_where = "VEND_TYPE = 'S'"
        End Select
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTIRECX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdICTIRECY, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdICTIREC2, "BB", "Item Status Inquiry", "PO Inquiry")
        Load_Popup_Menu(grdPOTORDRO, "SSB", "Show Filter", "Show GroupBox", "Receive Entire PO")
        Load_Popup_Menu(grdICTIRECG, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdICTIREC4, "B", "Item Status Inquiry")
        Load_Popup_Menu(grdEDTTRXNX, "BB", "Modify PO Number", "Remove PO Number", "Delete 3PL Receipt")
        Load_Popup_Menu(grdICT3PLTX, "BB", "Delete Transaction", "Change Date/Time")

    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Or Not GRDs.Keys.Contains(Mid(e.SourceControl.Name, 4)) Then
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
            Case "grdPOTORDRO"
                tlb_btn = DirectCast(tlb_pop.Tools("Receive Entire PO"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not InquiryMode

            Case "grdEDTTRXNX"
                tlb_btn = DirectCast(tlb_pop.Tools("Modify PO Number"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not InquiryMode
                tlb_btn = DirectCast(tlb_pop.Tools("Remove PO Number"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not InquiryMode

            Case "grdICT3PLTX"
                tlb_btn = DirectCast(tlb_pop.Tools("Delete Transaction"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.Band.Key = "ICT3PLTX"
                tlb_btn = DirectCast(tlb_pop.Tools("Change Date/Time"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.Band.Key = "ICT3PLTX"

            Case "grdEDTTRXNX"
                tlb_btn = DirectCast(tlb_pop.Tools("Delete 3PL Receipt"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.Band.Key = "EDTTRXNX"

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

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Acknowledge w/Notes"
            '    Log_SetMode(True, True)

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI")

            Case "Receive Entire PO"
                PO_ORDER_NO_RECEIVED = grd.ActiveRow.Cells("PO_ORDER_NO").Value
                Absx1.txtFor("VEND_CODE").Text = grd.ActiveRow.Cells("VEND_CODE").Value & ""
                Absx1.txtFor("WHSE_CODE").Text = grd.ActiveRow.Cells("WHSE_CODE").Value & ""
                Click_Command("New")
                If Not ScreenMode Then
                    PO_ORDER_NO_RECEIVED = ""
                End If

            Case "Modify PO Number"

                If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                    Exit Sub
                End If

                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Dim TRANS_NUM As String = grd.ActiveRow.Cells("TRANS_NUM").Text.Replace(",", "")
                If MessageBox.Show("Do you want to change PO Number " & PO_ORDER_NO & " to another Number?", "Modify PO Number", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                Dim PO_ORDER_NO_NEW As String = InputBox("Enter the New PO Number.", "Modify PO Number", String.Empty)
                If PO_ORDER_NO_NEW.Length = 0 Then
                    Exit Sub
                End If

                PO_ORDER_NO_NEW = ASCMAIN1.Format_Field(PO_ORDER_NO_NEW, "PO_ORDER_NO")

                If MessageBox.Show("Do you want to change PO Number " & PO_ORDER_NO & " to " & PO_ORDER_NO_NEW & "?", "Modify PO Number", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                EnforceConstraints(False)
                For Each tableName As String In New String() {"EDTTRXNZ", "EDTTRXNX"}
                    For Each row As DataRow In dst.Tables(tableName).Select("TRANS_NUM = " & TRANS_NUM & " and ISNULL(PO_ORDER_NO, '') = '" & PO_ORDER_NO & "'")
                        row.Item("PO_ORDER_NO") = PO_ORDER_NO_NEW
                    Next
                    dst.Tables(tableName).AcceptChanges()
                Next
                EnforceConstraints(True)

                Try
                    If PO_ORDER_NO.Length = 0 Then PO_ORDER_NO = "*"
                    ASCDATA1.ExecuteSQL("UPDATE EDTTRXN1 SET PO_ORDER_NO = '" & PO_ORDER_NO_NEW & "' WHERE NVL(PO_ORDER_NO, '*') = '" & PO_ORDER_NO & "' and TRANS_NUM = " & TRANS_NUM & " and TRANS_TYPE = 'REC'")
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try


            Case "Remove PO Number"

                If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                    Exit Sub
                End If

                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Dim TRANS_NUM As String = grd.ActiveRow.Cells("TRANS_NUM").Text.Replace(",", "")
                If MessageBox.Show("Do you want to Remove PO Number " & PO_ORDER_NO & " from the list?", "Remove PO Number", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                EnforceConstraints(False)
                For Each tableName As String In New String() {"EDTTRXNZ", "EDTTRXNX"}
                    For Each row As DataRow In dst.Tables(tableName).Select("TRANS_NUM = " & TRANS_NUM & " and ISNULL(PO_ORDER_NO, '') = '" & PO_ORDER_NO & "'")
                        row.Delete()
                    Next
                    dst.Tables(tableName).AcceptChanges()
                Next
                EnforceConstraints(True)

                Try
                    If PO_ORDER_NO.Length = 0 Then
                        ASCDATA1.ExecuteSQL("UPDATE EDTTRXN1 SET PROCESS_IND = 'R' WHERE PO_ORDER_NO IS NULL and TRANS_NUM = " & TRANS_NUM & " and TRANS_TYPE = 'REC'")
                    Else
                        ASCDATA1.ExecuteSQL("UPDATE EDTTRXN1 SET PROCESS_IND = 'R' WHERE PO_ORDER_NO = '" & PO_ORDER_NO & "' and TRANS_NUM = " & TRANS_NUM & " and TRANS_TYPE = 'REC'")
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try


            Case "Change Date/Time"

                If ASCMAIN1.CLIENT <> "INT" Then
                    Exit Sub
                End If

                With grdICT3PLTX.DisplayLayout.Bands(0)
                    .Override.AllowUpdate = DefaultableBoolean.True
                    .Override.CellClickAction = CellClickAction.EditAndSelectText
                    .Columns("DTDATE").CellActivation = Activation.AllowEdit
                    .Columns("DTTIME").CellActivation = Activation.AllowEdit
                End With

                MsgBox("Grid will now permit Update")


            Case "Delete Transaction"

                If ASCMAIN1.CLIENT <> "INT" Then
                    Exit Sub
                End If

                Dim DTDATE As String = grd.ActiveRow.Cells("DTDATE").Text
                Dim DTTIME As String = grd.ActiveRow.Cells("DTTIME").Text
                Dim DTUSER As String = grd.ActiveRow.Cells("DTUSER").Text
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                If MessageBox.Show("Do you want to Delete Transaction " & DTDATE & "-" & DTTIME & " for Item " & ITEM_CODE & " from the list?", "Remove Transaction", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                EnforceConstraints(False)
                grd.ActiveRow.Delete(False)
                EnforceConstraints(True)

                Try
                    ASCMAIN1.sql = "Update EDTTRXNA set PROCESSED_IND = 'D'" & vbCrLf _
                        & " where NVL(PROCESSED_IND,'0') = '0' and DTTTYP = 'A' and DTADJC in ('A03','A21')" & vbCrLf _
                        & "   and DTDATE = '" & DTDATE & "' AND DTTIME = '" & DTTIME & "' and DTUSER = '" & DTUSER & "'"

                    ASCDATA1.ExecuteSQL()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try


            Case "Delete 3PL Receipt"

                If ASCMAIN1.CLIENT <> "INT" Then
                    Exit Sub
                End If

                Dim TRANS_NUM As String = grd.ActiveRow.Cells("TRANS_NUM").Text
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                If MessageBox.Show("Do you want to Delete 3PL Receipt " & TRANS_NUM & " for PO " & PO_ORDER_NO & " from the list?", "Remove Transaction", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                EnforceConstraints(False)
                grd.ActiveRow.Delete(False)
                EnforceConstraints(True)

                Try
                    ASCMAIN1.sql = "Update EDTTRXN1 set PROCESS_IND = 'D'" & vbCrLf _
                        & " where NVL(PROCESS_IND,'0') = '0' and TRANS_NUM = '" & TRANS_NUM & "'" & vbCrLf _
                        & IIf(PO_ORDER_NO = "", " and PO_ORDER_NO is Null", $"   and PO_ORDER_NO = '{PO_ORDER_NO}' ") & vbCrLf _
                        & "   and TRANS_TYPE = 'REC'"

                    ASCDATA1.ExecuteSQL()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode And Not ScreenMode Then
                        Click_Command("New", e)
                    End If
                End If
            Case "RECEIPT_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                If Not InquiryMode And Not ScreenMode Then
                    Click_Command("New")
                End If
            Case "RECEIPT_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "WHSE_CODE"
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "INV_FREIGHT"
        End Select
    End Sub

#End Region

#Region "grdICTIREC2"

    Private Sub grdICTIREC2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTIREC2.AfterCellUpdate

        ' SR-6549 - Lot Numbers on Shipments and Receipts
        If e.Cell.Row.Band.Key <> grdICTIREC2.DisplayLayout.Bands(0).Key Then
            Exit Sub
        End If

        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"

                grdCodeDesc(grdICTIREC2, "ICTITEM1", "ITEM_CODE", "ITEM_DESC")
                If cdr IsNot Nothing Then
                    Dim ITEM_CODE As String = e.Cell.Value
                    Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                    e.Cell.Row.Cells("PROD_CODE").Value = cdr.Item("PROD_CODE")

                    'Dim rowICTSTAT2 = Fill_Record("ICTSTAT2", New String() {ITEM_CODE, WHSE_CODE}, True)
                    Dim COST_CATGY_CODE As String = cdr.Item("COST_CATGY_CODE") & ""
                    Dim PROD_CODE As String = cdr.Item("PROD_CODE") & ""
                    Dim ITEM_COST_STD As Decimal = Val(cdr.Item("ITEM_COST_STD") & "")
                    e.Cell.Row.Cells("COST_CATGY_CODE").Value = COST_CATGY_CODE
                    e.Cell.Row.Cells("PROD_CODE").Value = PROD_CODE
                    e.Cell.Row.Cells("ITEM_COST_STD").Value = ITEM_COST_STD
                Else
                    grdICTIREC2.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If

            Case "RECEIPT_QTY"

        End Select
    End Sub

    Private Sub grdICTIREC2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTIREC2.AfterExitEditMode
        'Select Case grdICTIREC2.ActiveCell.Column.Key

        'End Select
    End Sub

    Private Sub grdICTIREC2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTIREC2.AfterRowActivate
        'With grdICTIREC2.DisplayLayout.Bands(0)
        '    If grdICTIREC2.ActiveRow.IsAddRow Then
        '        .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        '        grdICTIREC2.ActiveCell = grdICTIREC2.ActiveRow.Cells("ITEM_CODE")
        '        grdICTIREC2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
        '    Else
        '        .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        '    End If
        'End With

        If EntryMode = "V" Then
            Show_GL()
            Show_Components()
        End If

    End Sub

    Private Sub grdICTIREC2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTIREC2.AfterRowsDeleted
        DisplayTotals()
    End Sub

    Private Sub grdICTIREC2_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTIREC2.AfterRowUpdate
        DisplayTotals()
    End Sub


    Private Sub grdICTIREC2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdICTIREC2.BeforeExitEditMode
        If grdICTIREC2.ActiveCell Is Nothing Then Exit Sub
        ' SR-6549 - Lot Numbers on Shipments and Receipts
        If grdICTIREC2.ActiveCell.Row.Band.Key <> grdICTIREC2.DisplayLayout.Bands(0).Key Then
            Exit Sub
        End If

        With grdICTIREC2.ActiveCell
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
                        End If
                    End If
                    'Case "BAR_CODE"
                    '    If location_support Then
                    '        If .Text <> "" Then
                    '            If .Value IsNot Nothing Then
                    '                .Value = .Text.ToUpper
                    '            End If

                    '        End If
                    '        If .Text <> "" Then
                    '            cdr = LookUp("WHTBARC1", .Text)
                    '            If cdr Is Nothing Then
                    '                ASCMAIN1.Progress("Invalid Bar Code (" & .Text & ")")
                    '                If .Value IsNot Nothing Then
                    '                    .Value = ""
                    '                End If
                    '                e.Cancel = True
                    '            End If
                    '        End If
                    '    End If

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

    Private Sub grdICTIREC2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTIREC2.BeforeRowUpdate
        With grdICTIREC2

            ' SR-6549 - Lot Numbers on Shipments and Receipts
            Select Case e.Row.Band.Key
                Case grdICTIREC2.DisplayLayout.Bands(0).Key

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

                        If e.Row.Cells("LOCATION_CODE").Text = "" Then
                            e.Cancel = True
                        Else
                            LookUp("WHTLOCM1", New String() {Absx1.txtFor("WHSE_CODE").Text, e.Row.Cells("LOCATION_CODE").Text})
                            If cdr Is Nothing Then
                                MsgBox("Invalid Value entered for Location Code (" & e.Row.Cells("LOCATION_CODE").Text & ")",
                                       MsgBoxStyle.OkOnly, "Cannot Update Row")
                                e.Cancel = True
                            End If
                        End If

                    End If

                    If Val(e.Row.Cells("QTY_REC").Text) = 0 Then
                        MsgBox("Invalid Value entered for Qty (" & e.Row.Cells("QTY_REC").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If

                    If e.Cancel Then
                        e.Row.CancelUpdate()
                    End If

                    If Not e.Cancel Then
                        If e.Row.Cells("RECEIPT_NO").Text = "" Then
                            .ActiveRow.Cells("RECEIPT_NO").Value = Absx1.CtlFor("RECEIPT_NO").Text
                            .ActiveRow.Cells("RECEIPT_LNO").Value = Val(dst.Tables("ICTIREC2").Compute("Max(RECEIPT_LNO)", "") & "") + 1
                        End If
                    End If

                Case grdICTIREC2.DisplayLayout.Bands(1).Key

            End Select

        End With
    End Sub

    Private Sub grdICTIREC2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTIREC2.ClickCellButton

        If grdICTIREC2.ActiveRow Is Nothing Then Exit Sub

        ' SR-6549 - Lot Numbers on Shipments and Receipts
        If grdICTIREC2.ActiveRow.Band.Key <> grdICTIREC2.DisplayLayout.Bands(0).Key Then
            Exit Sub
        End If

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
            Case "LOCATION_CODE"
                sql_where = "WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'"
        End Select
        grdClickCellButton(grdICTIREC2, sql_where, False)

    End Sub

    Private Sub grdICTIREC2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdICTIREC2.Error
        grdICTIREC2.ActiveRow.CancelUpdate()
    End Sub

#End Region

    Private Sub ResendPOsWithOpenQuantity(ByVal PO_ORDER_NO As String)
        ' ISSUE-7044 - balance POs/ASNs
        Try
            Dim lstRetransmitPOs As New List(Of String)
            Dim lstLeaveOpenPOs As New List(Of String)
            Dim lstClosePOs As New List(Of String)

            Dim LP_CODE As String = "ADS"

            ASCMAIN1.sql = $"SELECT DISTINCT PO_ORDER_NO FROM
                            (
                                SELECT DISTINCT POTORDR1.PO_ORDER_NO, ICTPINV1.PO_ORDER_NO PO_ORDER_NO_I
                                FROM POTORDR1, POTORDR2, ICTWHSE1, ICTPINV1
                                WHERE POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO
                                AND POTORDR1.PO_ORDER_NO = :PARM1
                                AND POTORDR1.WHSE_CODE = ICTWHSE1.WHSE_CODE
                                AND ICTWHSE1.LP_CODE = '{LP_CODE}'
                                AND POTORDR1.PO_STATUS = 'O'
                                AND POTORDR2.PO_STATUS = 'O'
                                AND POTORDR1.VEND_CODE NOT IN ('IPSA', 'ADS3PL')
                                AND NVL(POTORDR2.PO_QTY_OPN, 0) > 0
                                and POTORDR1.PO_ORDER_NO = ICTPINV1.PO_ORDER_NO (+)
                            )
                            WHERE PO_ORDER_NO_I IS NULL"

            Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", {PO_ORDER_NO})
            If tbl.Rows.Count = 0 Then
                Exit Sub
            End If

            For Each dr As DataRow In tbl.Select("")
                PO_ORDER_NO = dr.Item("PO_ORDER_NO")
                ' Nathan turned this off on 12/15/2025. Chelsea (ADS) confimred they will make another flight
                ' "Leave PO Open and Retransmit Now"
                Using frmASFMSGBF As New ASFMSGBF
                    Dim frtOption As Integer = frmASFMSGBF.Get_opt_from_User($"Options for PO No: {PO_ORDER_NO}", New String() {"Leave PO Open", "Close PO"}, 0, "PO Procesing.")
                    Select Case frtOption
                        Case 0
                            Exit Sub
                            'lstLeaveOpenPOs.Add(PO_ORDER_NO)
                        Case 1
                            lstClosePOs.Add(PO_ORDER_NO)
                            'Case 2
                            '    lstRetransmitPOs.Add(PO_ORDER_NO)
                        Case Else
                            Exit Sub
                    End Select
                End Using
            Next

            ' Nathan turned this off on 12/15/2025. Chelsea (ADS) confimred they will make another flight
            ' "Leave PO Open and Retransmit Now"
            'If lstRetransmitPOs.Count > 0 Then
            '    ' VerifY all items are updated
            '    Dim XMIT_NO As String = String.Empty
            '    Try
            '        ASCMAIN1.Progress("Sending Item Master to ADS", "")
            '        XMIT_NO = TAC.ICCMAIN1.Transmit_Document("WHC", "WHC888O1", "AC", "", LP_CODE)
            '    Catch ex As Exception
            '        MessageBox.Show(ex.Message, "Update Item Master", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '    End Try

            '    Try
            '        ASCMAIN1.Progress("Transmitting Purchase Order(s)", "")
            '        XMIT_NO = String.Empty
            '        XMIT_NO = TAC.ICCMAIN1.Transmit_Document("WHC", "WHC943O1", "SPO", String.Join(",", lstRetransmitPOs.ToArray), LP_CODE)

            '        Dim rowWHT3PLX1 As DataRow
            '        rowWHT3PLX1 = LookUp("WHT3PLX1", XMIT_NO)
            '        If Val(rowWHT3PLX1.Item("XMIT_RECORDS") & "") <> 0 Then
            '            For Each PO_ORDER_NO In lstRetransmitPOs
            '                ASCMAIN1.sql = $"Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY) 
            '                            Select 'POTORDR1', PO_ORDER_NO, SYSDATE, '{ASCMAIN1.USER_ID}', 'PO_943I','PO/Invoice Transmitted to {LP_CODE}', '{XMIT_NO}'
            '                            From POTORDR1 where POTORDR1.PO_ORDER_NO = '{PO_ORDER_NO}'"
            '                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            '            Next
            '        End If
            '    Catch ex As Exception
            '        MessageBox.Show(ex.Message, "Transfer PO Balances", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '    End Try
            'End If

            For Each PO_ORDER_NO In lstLeaveOpenPOs
                Try
                    BeginTrans()
                    ASCMAIN1.sql = "UPDATE POTORDR1 SET PO_PRINTED_IND = NULL, PO_HDR_CTR_REV = NVL(PO_HDR_CTR_REV, 0) + 1, PO_XMIT_IND = NULL WHERE PO_ORDER_NO = :PARM1"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {PO_ORDER_NO})
                    CommitTrans($"PO {PO_ORDER_NO} set for retransmission")
                Catch ex As Exception
                    Rollback($"Error marking PO {PO_ORDER_NO} to permit retransmission: " & ex.Message)
                End Try
            Next

            For Each PO_ORDER_NO In lstClosePOs
                Try
                    BeginTrans()
                    CancelPOOpenQuantities(PO_ORDER_NO)
                    CommitTrans($"PO {PO_ORDER_NO} has been Cancelled")
                Catch ex As Exception
                    Rollback($"Error Cancelling Open Quantities on PO {PO_ORDER_NO}: " & ex.Message)
                End Try
            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Resend PO Open Quantity", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress("", "")
        End Try
    End Sub

    Sub CancelPOOpenQuantities(PO_ORDER_NO As String)
        ' ISSUE-7044 - balance POs/ASNs

        ' Dependent_Updates(-1, PO_ORDER_NO)
        TAC.POCMAIN1.Production_Commit(-1, PO_ORDER_NO)
        TAC.POCMAIN1.ICTSTAT2_PO(-1, PO_ORDER_NO)

        Dim sqlw As String = $" where PO_ORDER_NO = '{PO_ORDER_NO}'"
        Dim sqlw2 As String = " and PO_QTY_OPN <> 0"
        ASCDATA1.ExecuteSQL("Update POTORDR2 set PO_QTY_CXL = NVL(PO_QTY_CXL,0) + NVL(PO_QTY_OPN,0) " & sqlw & sqlw2)
        ASCDATA1.ExecuteSQL("Update POTORDR2 set PO_QTY_OPN = 0, PO_STATUS = 'C' " & sqlw & sqlw2)
        ASCDATA1.ExecuteSQL("Update POTORDR5 set PO_NINV_QTY_OPN = 0, PO_NINV_AMT_OPN = 0, PO_NINV_STATUS = 'C' " & sqlw)

        Fill_Records("POTORDR1", PO_ORDER_NO)
        Dim rowPOTORDR1 As DataRow = dst.Tables("POTORDR1").Rows.Find(PO_ORDER_NO)

        If rowPOTORDR1.Item("PO_PRINTED_IND") & "" = "1" Then
            rowPOTORDR1.Item("PO_HDR_CTR_REV") = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "") + 1
            rowPOTORDR1.Item("PO_PRINTED_IND") = DBNull.Value
        End If
        rowPOTORDR1.Item("PO_STATUS") = "C"
        rowPOTORDR1.Item("PO_DATE_CANCELLED") = DATETIME_STAMP.Date
        Update_Record_TDA("POTORDR1")

        TAC.TACMAIN1.Record_Event("POTORDR1", PO_ORDER_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "PO-CXL", "PO Cancelled")
    End Sub


    Sub DisplayTotals()
        'Dim TOTAL_COSTS As Decimal = Val(dst.Tables("ICTIREC2").Compute("SUM(LINE_COSTS)", "") & "")
        'Absx1.numFor("TOTAL_COSTS").Value = TOTAL_COSTS
    End Sub

    Private Sub grdICTIRECX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTIRECX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("RECEIPT_NO").Text = e.Row.Cells("RECEIPT_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub grdICTIRECY_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTIRECY.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("RECEIPT_NO").Text = e.Row.Cells("RECEIPT_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub grdICTIRECG_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTIRECG.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("RECEIPT_NO").Text = e.Row.Cells("RECEIPT_NO").Text
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
                grdICTIREC3.DataSource = dst.Tables("ICTIREC3")
                Dim dvw As DataView = dst.Tables("ICTIREC3").DefaultView
                dvw.RowFilter = ""
            ElseIf optGL.Value = "L" Then
                grdICTIREC3.DataSource = dst.Tables("ICTIREC3")
                Dim dvw As DataView = dst.Tables("ICTIREC3").DefaultView
                Dim RECEIPT_LNO As Integer = 0
                If grdICTIREC2.ActiveRow IsNot Nothing Then
                    RECEIPT_LNO = Val(grdICTIREC2.ActiveRow.Cells("RECEIPT_LNO").Text)
                End If
                dvw.RowFilter = "RECEIPT_LNO = " & CStr(RECEIPT_LNO)
            ElseIf optGL.Value = "S" Then
                Dim tbl As DataTable = dst.Tables("ICTIREC3").Clone
                Dim RECEIPT_GNO As Integer = 0
                For Each rowA234 As DataRow In ASCDATA1.SelectDistinct _
                ("ICTIREC3", New String() {"ACCT_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE", "ACCT_DESC"}).Rows
                    Dim DIST_AMT As Decimal = Val(dst.Tables("ICTIREC3").Compute _
                    ("SUM(DIST_AMT)",
                     "ACCT_CODE = '" & rowA234.Item("ACCT_CODE") & "' and SEG2_CODE = '" & rowA234.Item("SEG2_CODE") & "' and SEG3_CODE = '" & rowA234.Item("SEG3_CODE") & "' and SEG4_CODE = '" & rowA234.Item("SEG4_CODE") & "'") & "")
                    Dim row As DataRow = tbl.NewRow
                    row.Item("RECEIPT_NO") = Absx1.txtFor("RECEIPT_NO").Text
                    row.Item("RECEIPT_LNO") = 0
                    RECEIPT_GNO += 1
                    row.Item("RECEIPT_GNO") = RECEIPT_GNO
                    row.Item("ACCT_CODE") = rowA234.Item("ACCT_CODE")
                    row.Item("SEG2_CODE") = rowA234.Item("SEG2_CODE")
                    row.Item("SEG3_CODE") = rowA234.Item("SEG3_CODE")
                    row.Item("SEG4_CODE") = rowA234.Item("SEG4_CODE")
                    row.Item("ACCT_DESC") = rowA234.Item("ACCT_DESC")
                    row.Item("DIST_AMT") = DIST_AMT
                    tbl.Rows.Add(row)
                Next

                grdICTIREC3.DataSource = tbl
            End If
        End If
    End Sub

    Sub Show_Components()
        Dim dvw As DataView = dst.Tables("ICTIREC4").DefaultView
        Dim RECEIPT_LNO As Integer = 0
        If grdICTIREC2.ActiveRow IsNot Nothing Then
            RECEIPT_LNO = Val(grdICTIREC2.ActiveRow.Cells("RECEIPT_LNO").Text)
        End If
        dvw.RowFilter = "RECEIPT_LNO = " & CStr(RECEIPT_LNO)
    End Sub

    Private Sub cbeYP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeYP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Refresh_Documents()
    End Sub

    Private Sub cbeYP0_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeYP0.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Refresh_Documents()
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Refreshing Documents")

        EnforceConstraints(False)

        Dim YP0 As String = cbeYP0.Value
        Dim YP As String = cbeYP.Value

        Fill_Records("ICTIRECX", New String() {YP0, YP})
        Sort_grdColumns(grdICTIRECX, "RECEIPT_NO".ToLower)
        'grdICTIRECX.Text = "Entered in " & cbeYP.Text
        grdICTIRECX.Text = "Entered from " & cbeYP.Text & " thru " & cbeYP0.Text

        Fill_Records("ICTIRECY", New String() {YP0, YP})
        Sort_grdColumns(grdICTIRECY, "RECEIPT_NO".ToLower)
        'grdICTIRECY.Text = "Entered in " & cbeYP.Text
        grdICTIRECY.Text = "Entered from " & cbeYP.Text & " thru " & cbeYP0.Text

        Fill_Records("POTORDRO")
        Sort_grdColumns(grdPOTORDRO, "PO_ORDER_NO".ToLower)
        grdPOTORDRO.Text = "Open Purchase Orders"

        If chkGL.Checked Then
            Fill_Records("ICTIRECG", New String() {YP0, YP})
            grdICTIRECG.Text = "Entered in " & cbeYP.Text
            grdICTIRECG.Text = "Entered from " & cbeYP.Text & " thru " & cbeYP0.Text
        End If


        Fill_Records("EDTTRXNX")
        Fill_Records("EDTTRXNZ")
        If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
            Fill_Records("ICT3PLTX")
            Fill_Records("ICT3PLTY")
            Fill_Records("ICT3PLTZ")
        End If

        EnforceConstraints(True)

        grdICT3PLTX.Rows.ExpandAll(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Function Check_Qty(ByVal TABLE_NAME As String,
                       ByVal WHSE_CODE As String,
                       ByVal QTY_FIELD As String,
                       ByVal S As Integer) As String

        Dim msg As String = ""

        For Each row As DataRow In dst.Tables(TABLE_NAME).Rows
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
            Dim PO_ORDER_LNO As Integer = Val(row.Item("PO_ORDER_LNO") & "")
            Dim QTY As Integer = row.Item(QTY_FIELD)
            ASCMAIN1.sql = "Select * from POTORDR2 where ITEM_CODE = '" & ITEM_CODE & "' and PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)
            Dim rowPOTORDR2 As DataRow = ASCDATA1.GetDataRow
            Dim PO_QTY_OPN As Integer = 0
            If rowPOTORDR2 IsNot Nothing Then
                PO_QTY_OPN = Val(rowPOTORDR2.Item("PO_QTY_OPN") & "")
            End If
            If PO_QTY_OPN + S * QTY < 0 Then
                msg &= vbCr & Format("Item " & ITEM_CODE & " has only " & CStr(PO_QTY_OPN) & " Open On Order")
            End If
        Next

        Return msg
    End Function

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

        Dim REVERSED_BY_RECEIPT_NO As String = ASCMAIN1.Next_Control_No("ICTIREC1.RECEIPT_NO")

        Dim rowICTIREC1_orig As DataRow = dst.Tables("ICTIREC1").NewRow
        rowICTIREC1_orig.ItemArray = rowICTIREC1.ItemArray

        rowICTIREC1 = dst.Tables("ICTIREC1").Rows(0)
        rowICTIREC1.Item("REVERSED_BY_RECEIPT_NO") = REVERSED_BY_RECEIPT_NO
        rowICTIREC1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowICTIREC1.Item("LAST_DATE") = DATETIME_STAMP
        Update_Record_TDA("ICTIREC1")

        rowICTIREC1.ItemArray = rowICTIREC1_orig.ItemArray
        rowICTIREC1.AcceptChanges()
        rowICTIREC1.SetAdded()

        With rowICTIREC1
            .Item("REVERSES_RECEIPT_NO") = .Item("RECEIPT_NO")
            .Item("RECEIPT_NO") = REVERSED_BY_RECEIPT_NO
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP

            Dim DT As Date = DATETIME_STAMP.Date
            Dim dts() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
            If Format(DT, "yyyyMMdd") > Format(dts(dts.Length - 1), "yyyyMMdd") Then
                DT = dts(dts.Length - 1)
            End If

            .Item("RECEIPT_DATE") = DT
            '  .Item("TOTAL_COSTS") *= -1

            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("REGISTER_IND") = "0"
            .Item("REGISTER_XNO") = DBNull.Value
            .Item("JOURNAL_IND") = "0"
            .Item("JOURNAL_XNO") = DBNull.Value
            '.Item("RECEIPT_SOURCE") = "R" 'Reversal Source?
        End With

        For Each row As DataRow In dst.Tables("ICTIREC2").Rows
            row.Item("RECEIPT_NO") = REVERSED_BY_RECEIPT_NO
            Dim QTY_REC As Int64 = Val(row.Item("QTY_REC") & "")
            Dim QTY_INV As Int64 = Val(row.Item("QTY_INV") & "")
            Dim PO_COST As Decimal = Val(row.Item("PO_COST") & "")
            Dim AMT_INV As Decimal = Val(row.Item("AMT_INV") & "")
            Dim TRAN_PV As Decimal = Val(row.Item("TRAN_PV") & "")

            row.Item("QTY_REC") = -1 * QTY_REC
            row.Item("QTY_INV") = 0
            row.Item("AMT_INV") = 0
            row.Item("TRAN_PV") = TRAN_PV - (AMT_INV - PO_COST * QTY_INV) ' REDUCE TRAN_PV BY PV INCURRED AT AP
            row.Item("ACCRUAL_STATUS") = "0"

            Dim EXT_COST_MATLS As Decimal = Val(row.Item("EXT_COST_MATLS") & "")
            row.Item("EXT_COST_MATLS") = -1 * EXT_COST_MATLS

            row.Item("OPS_YYYYPP") = ASCMAIN1.CYP

            row.AcceptChanges()
            row.SetAdded()
        Next

        For Each row As DataRow In dst.Tables("ICTIREC4").Rows
            row.Item("RECEIPT_NO") = REVERSED_BY_RECEIPT_NO
            If row.Item("QTY_CON") IsNot DBNull.Value Then
                row.Item("QTY_CON") *= -1
            End If
            row.AcceptChanges()
            row.SetAdded()
        Next
    End Sub

    Private Sub grdPOTORDRX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTORDRX.DoubleClickRow

        Dim row As DataRow = dst.Tables("POTORDRX").Rows.Find(New Object() {e.Row.Cells("PO_ORDER_NO").Value, e.Row.Cells("PO_ORDER_LNO").Value})
        If Not ASCMAIN1.Logical_Lock("POTORDR1", e.Row.Cells("PO_ORDER_NO").Value, False, False, False, 1) Then Exit Sub
        Receive_PO_Line(row)
    End Sub

    Sub Receive_PO_Line(row As DataRow, Optional QTY_REC As Int64 = 0)

        ' SR-6549 - Lot Numbers on Shipments and Receipts
        Dim LOT_NO As String = String.Empty
        If row.Table.Columns.Contains("LOT_NO") Then
            LOT_NO = row.Item("LOT_NO") & String.Empty
        End If

        Dim TRX_NO As String = String.Empty
        If row.Table.Columns.Contains("TRX_NO") Then
            TRX_NO = row.Item("TRX_NO") & String.Empty
        End If

        Dim TRX_LNO As String = 0
        If row.Table.Columns.Contains("TRX_LNO") Then
            TRX_LNO = Val(row.Item("TRX_LNO") & String.Empty)
        End If

        If ASCMAIN1.CLIENT = "INT" Then
            Dim rowICTIREC2s() As DataRow = dst.Tables("ICTIREC2").Select("ITEM_CODE = '" & row.Item("ITEM_CODE") & "' and PO_ORDER_LNO = " & row.Item("PO_ORDER_LNO"))
            If rowICTIREC2s.Length = 1 Then
                rowICTIREC2s(0).Item("QTY_REC") = Val(rowICTIREC2s(0).Item("QTY_REC") & "") + QTY_REC

                ' SR-6549 - Lot Numbers on Shipments and Receipts
                If LOT_NO <> String.Empty Then
                    Dim rowICTIRECL As DataRow = dst.Tables("ICTIRECL").NewRow
                    rowICTIRECL.Item("RECEIPT_NO") = rowICTIREC2s(0).Item("RECEIPT_NO")
                    rowICTIRECL.Item("RECEIPT_LNO") = rowICTIREC2s(0).Item("RECEIPT_LNO")
                    rowICTIRECL.Item("RECEIPT_LNO_SEQ") = Val(dst.Tables("ICTIRECL").Compute("MAX(RECEIPT_LNO_SEQ)", $"RECEIPT_NO = '{rowICTIRECL.Item("RECEIPT_NO")}' AND RECEIPT_LNO = {rowICTIRECL.Item("RECEIPT_LNO")}") & String.Empty) + 1
                    rowICTIRECL.Item("LOT_NO") = row.Item("LOT_NO")
                    rowICTIRECL.Item("REC_QTY") = QTY_REC
                    dst.Tables("ICTIRECL").Rows.Add(rowICTIRECL)
                End If
                Exit Sub
            End If
        End If

        grdICTIREC2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes
        Dim ITEM_CODE As String = row.Item("ITEM_CODE")
        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
        grdICTIREC2.DisplayLayout.Bands(0).AddNew()
        With grdICTIREC2.ActiveRow
            .Cells("ITEM_CODE").Value = row.Item("ITEM_CODE")

            If QTY_REC <> 0 Then
                .Cells("QTY_REC").Value = QTY_REC
            Else
                .Cells("QTY_REC").Value = row.Item("PO_QTY_OPN")
            End If
            .Cells("PO_ORDER_NO").Value = row.Item("PO_ORDER_NO")
            .Cells("PO_ORDER_LNO").Value = row.Item("PO_ORDER_LNO")
            .Cells("PO_QTY_ORD").Value = row.Item("PO_QTY_ORD")
            .Cells("PO_QTY_OPN").Value = row.Item("PO_QTY_OPN")
            .Cells("ITEM_COST_STD").Value = rowICTITEM1.Item("ITEM_COST_STD")
            .Cells("PO_COST").Value = row.Item("PO_COST")
            .Cells("ITEM_UOM").Value = row.Item("ITEM_UOM")
            Dim rowPOTORDR1 As DataRow = dst.Tables("POTORDR1").Rows.Find(New Object() {row.Item("PO_ORDER_NO")})
            If rowPOTORDR1 Is Nothing Then rowPOTORDR1 = Fill_Record("POTORDR1", New Object() {row.Item("PO_ORDER_NO")}, False, False)
            Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find(New Object() {row.Item("PO_ORDER_NO"), row.Item("PO_ORDER_LNO")})
            If rowPOTORDR2 Is Nothing Then rowPOTORDR2 = Fill_Record("POTORDR2", New Object() {row.Item("PO_ORDER_NO"), Val(row.Item("PO_ORDER_LNO"))}, False, False)

            If rowPOTORDR2.Item("BM_ISSUE_SEL") & "" = "1" Or rowPOTORDR2.Item("BM_ISSUE_NO") & "" <> "" Then
                .Cells("VEND_WHSE_CODE").Value = rowPOTORDR1.Item("VEND_WHSE_CODE")
                .Cells("BM_ISSUE_SEL").Value = rowPOTORDR2.Item("BM_ISSUE_SEL")
                .Cells("BM_ISSUE_NO").Value = rowPOTORDR2.Item("BM_ISSUE_NO")
            End If

            If location_support Then
                .Cells("LOCATION_CODE").Value = rowICTWHSE1.Item("WHSE_LOC_REC")
            End If

            If LOT_NO.Length > 0 AndAlso TRX_NO.Length > 0 Then
                .Cells("LOT_NO").Value = LOT_NO
                .Cells("TRX_NO").Value = TRX_NO
                .Cells("TRX_LNO").Value = TRX_LNO
            End If

            .Update()

            ' SR-6549 - Lot Numbers on Shipments and Receipts
            If LOT_NO.Length > 0 Then
                Dim sql As String = $"TRX_NO = '{TRX_NO}' and TRX_LNO = {TRX_LNO} and LOT_NO = '{LOT_NO}'"
                If dst.Tables("ICTIREC2").Select(sql).Length > 0 Then
                    Dim rowICITREC2 As DataRow = dst.Tables("ICTIREC2").Select(sql)(0)
                    Dim rowICTIRECL As DataRow = dst.Tables("ICTIRECL").NewRow
                    rowICTIRECL.Item("RECEIPT_NO") = rowICITREC2.Item("RECEIPT_NO")
                    rowICTIRECL.Item("RECEIPT_LNO") = rowICITREC2.Item("RECEIPT_LNO")
                    rowICTIRECL.Item("RECEIPT_LNO_SEQ") = Val(dst.Tables("ICTIRECL").Compute("MAX(RECEIPT_LNO_SEQ)", $"RECEIPT_NO = '{rowICTIRECL.Item("RECEIPT_NO")}' AND RECEIPT_LNO = {rowICTIRECL.Item("RECEIPT_LNO")}") & String.Empty) + 1
                    rowICTIRECL.Item("LOT_NO") = LOT_NO
                    rowICTIRECL.Item("REC_QTY") = Val(rowICITREC2.Item("QTY_REC") & String.Empty)
                    dst.Tables("ICTIRECL").Rows.Add(rowICTIRECL)
                End If
            End If
        End With
        grdICTIREC2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No

    End Sub

    Private Sub grdICTIRECX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTIRECX.InitializeRow
        If e.Row.IsDataRow Then
            If e.Row.Cells("REVERSED_BY_RECEIPT_NO").Value & "" <> "" Then
                e.Row.Cells("REVERSED_BY_RECEIPT_NO").Appearance.ForeColor = Color.Red
                e.Row.Cells("RECEIPT_NO").Appearance.ForeColor = Color.Red
                e.Row.Cells("RECEIPT_NO").ToolTipText = "This receipt was Reversed by " & e.Row.Cells("REVERSED_BY_RECEIPT_NO").Value
                e.Row.Cells("REVERSED_BY_RECEIPT_NO").ToolTipText = "This receipt was Reversed by " & e.Row.Cells("REVERSED_BY_RECEIPT_NO").Value
            ElseIf e.Row.Cells("REVERSES_RECEIPT_NO").Value & "" <> "" Then
                e.Row.Cells("REVERSES_RECEIPT_NO").Appearance.ForeColor = Color.Red
                e.Row.Cells("REVERSES_RECEIPT_NO").ToolTipText = "This receipt Reverses Receipt " & e.Row.Cells("REVERSES_RECEIPT_NO").Value
                e.Row.Appearance.BackColor = Color.Yellow
            End If
        End If
    End Sub

    Private Sub grdEDTTRXNX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdEDTTRXNX.DoubleClickRow
        If InquiryMode Then Exit Sub

        PO_ORDER_NO_3PL = e.Row.Cells("PO_ORDER_NO").Value & String.Empty
        TRANS_NUM_3PL = e.Row.Cells("TRANS_NUM").Value & ""

        If PO_ORDER_NO_3PL.Length = 0 Then
            Exit Sub
        End If

        If dst.Tables("EDTTRXNX").Select("PO_ORDER_NO = '" & PO_ORDER_NO_3PL & "' and TRANS_NUM = '" & TRANS_NUM_3PL & "'").Length = 0 Then
            Exit Sub
        End If

        If e.Row.Band.Key = grdEDTTRXNX.DisplayLayout.Bands(1).Key Then
            InvoiceNum = e.Row.Cells("INV_NUM").Value & ""
        Else
            ' See if this is an ASN with multiple Invoice Numbers
            Dim tblInvNos As DataTable = ASCDATA1.SelectDistinct(dst.Tables("EDTTRXNz").Select("PO_ORDER_NO = '" & PO_ORDER_NO_3PL & "' and TRANS_NUM = '" & TRANS_NUM_3PL & "'"), "INV_NUM")
            Select Case tblInvNos.Rows.Count
                Case 0
                    InvoiceNum = String.Empty
                Case 1
                    InvoiceNum = tblInvNos.Rows(0).Item("INV_NUM") & String.Empty
                Case Else
                    MessageBox.Show("The selected PO contains more than 1 Invoice No. You must double-click one of the detail lines and only that Invoice No will be received.", "Select PO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
            End Select
        End If

        Dim row As DataRow = dst.Tables("EDTTRXNX").Select("PO_ORDER_NO = '" & PO_ORDER_NO_3PL & "' and TRANS_NUM = '" & TRANS_NUM_3PL & "'")(0)

        '  PO_ORDER_NO_3PL = grdEDTTRXNX.ActiveRow.Cells("PO_ORDER_NO").Value
        Absx1.txtFor("VEND_CODE").Text = row.Item("VEND_CODE") & ""
        Absx1.txtFor("WHSE_CODE").Text = row.Item("WHSE_CODE") & ""

        If Absx1.dteFor("RECEIPT_DATE").Value & "" <> "" AndAlso Format(Absx1.dteFor("RECEIPT_DATE").Value, "yyyyMMdd") = Format(Now, "yyyyMMdd") Then
            Absx1.dteFor("RECEIPT_DATE").Value = row.Item("TRANS_DATE")
        End If


        Click_Command("New")
    End Sub

    Private Sub grdICTIREC2_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdICTIREC2.InitializeRow

        ' SR-6549 - Lot Numbers on Shipments and Receipts
        Select Case e.Row.Band.Key
            Case grdICTIREC2.DisplayLayout.Bands(0).Key
                Dim QTY_REC As Int64 = Val(e.Row.Cells("QTY_REC").Value & "")
                Dim PO_QTY_OPN As Int64 = Val(e.Row.Cells("PO_QTY_OPN").Value & "")
                e.Row.Cells("QTY_REC").ToolTipText = ""
                If EntryMode = "N" Then
                    With e.Row.Cells("QTY_REC").Appearance
                        If QTY_REC = PO_QTY_OPN Then
                            .ForeColor = Color.Empty
                        ElseIf QTY_REC < PO_QTY_OPN Then
                            e.Row.Cells("QTY_REC").ToolTipText = "Qty Rec < Open PO"
                            .ForeColor = Color.Red
                        ElseIf QTY_REC > PO_QTY_OPN Then
                            e.Row.Cells("QTY_REC").ToolTipText = "Qty Rec > Open PO"
                            .ForeColor = Color.Green
                        End If
                    End With
                End If

                'If HFs("WHSE_CODE") = "ADS" Then
                Dim QTY_REC_LOT As Int64 = Val(e.Row.Cells("QTY_REC_LOT").Value & "")
                If QTY_REC <> QTY_REC_LOT Then
                    e.Row.Cells("QTY_REC_LOT").Appearance.BackColor = Color.Red
                    e.Row.Cells("QTY_REC_LOT").Appearance.ForeColor = Color.White
                Else
                    e.Row.Cells("QTY_REC_LOT").Appearance.BackColor = Color.LightBlue
                    e.Row.Cells("QTY_REC_LOT").Appearance.ForeColor = Color.Black
                End If
                'End If
        End Select

    End Sub

    Private Sub grdICT3PLTX_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdICT3PLTX.DoubleClickRow
        If e.Row.Band.Key = "ICT3PLTX_ICT3PLTZ" Then

            Dim PO_ORDER_NO As String = e.Row.Cells("PO_ORDER_NO").Value
            Dim PO_TYPE As String = e.Row.Cells("PO_TYPE").Value
            Dim WHSE_CODE As String = e.Row.Cells("WHSE_CODE").Value
            Dim VEND_WHSE_CODE As String = e.Row.Cells("VEND_WHSE_CODE").Value & ""
            Dim VEND_CODE As String = e.Row.Cells("VEND_CODE").Value
            Dim DTADJC As String = e.Row.ParentRow.Cells("DTADJC").Value

            ' smz 12/17/25: Clarins only code removed as part of warehouse transition
            'If DTADJC = "A03" Then
            '    If PO_TYPE <> "M" Or WHSE_CODE <> "CLA" Or VEND_WHSE_CODE <> "CLA" Or VEND_CODE <> "CLARINSUSA" Then
            '        MsgBox("PO " & PO_ORDER_NO & " must be a Make PO with CLARINS, assembled at CLA, to be received at CLA, and it is not", MsgBoxStyle.OkOnly, "Cannot Receive Selected PO as an Assembly Receipt")
            '        Exit Sub
            '    End If
            'Else
            '    If PO_TYPE <> "B" Or WHSE_CODE <> "CLA" Or VEND_CODE <> "CLARINSUSA" Then
            '        MsgBox("PO " & PO_ORDER_NO & " must be a Buy PO with CLARINS, dis-assembled at CLA, to be received at CLA, and it is not", MsgBoxStyle.OkOnly, "Cannot Receive Selected PO as an Assembly Receipt")
            '        Exit Sub
            '    End If
            'End If


            Dim DTDATE As String = e.Row.Cells("DTDATE").Value
            Dim DTTIME As String = e.Row.Cells("DTTIME").Value
            Dim DTUSER As String = e.Row.Cells("DTUSER").Value

            PO_ORDER_NO_RECEIVED = PO_ORDER_NO
            PO_ORDER_NO_3PL = PO_ORDER_NO
            TRANS_NUM_3PL = "CLARINS_ASSY" & vbTab & DTDATE & vbTab & DTTIME & vbTab & DTUSER
            rowICT3PLTX = dst.Tables("ICT3PLTX").Rows.Find(New String() {DTDATE, DTTIME, DTUSER})

            Absx1.txtFor("VEND_CODE").Text = VEND_CODE
            Absx1.txtFor("WHSE_CODE").Text = WHSE_CODE
            Click_Command("New")
            If Not ScreenMode Then
                PO_ORDER_NO_RECEIVED = ""
                rowICT3PLTX = Nothing
            Else

            End If

        End If
    End Sub

    Sub Check_for_IPSA_Invoice(reverse As Boolean)

        If Absx1.txtFor("VEND_CODE").Text = "IPSA" Then
            Dim SOURCE_DOC_NO As String = Absx1.txtFor("SOURCE_DOC_NO").Text
            ' not sure about protecting against receipt no
            If SOURCE_DOC_NO = "NONE" And ASCMAIN1.USER_ID = "lmarinelli" Then
                ' disable integration with IDOC Invoice
            Else

                If reverse Then
                    ASCMAIN1.sql = "Select * from ICTPINV1" & vbCrLf _
                        & " where VEND_CODE = 'IPSA'" & vbCrLf _
                        & "   and INV_NUM LIKE '%' || :PARM1" & vbCrLf _
                        & "   and PINV_STATUS = 'I'" & vbCrLf _
                        & "   and RECEIPT_NO = '" & Absx1.txtFor("RECEIPT_NO").Text & "'"
                Else
                    ASCMAIN1.sql = "Select * from ICTPINV1" & vbCrLf _
                        & " where VEND_CODE = 'IPSA'" & vbCrLf _
                        & "   and INV_NUM LIKE '%' || :PARM1" & vbCrLf _
                        & "   and PINV_STATUS = 'O'" & vbCrLf _
                        & "   and RECEIPT_NO is Null" & vbCrLf _
                        & "   and VOUCHER_NO is Null"
                End If

                Dim sqlICTPINV1 As String = ASCMAIN1.sql
                Dim rows() As DataRow = ASCDATA1.GetDataTable(sqlICTPINV1, "", "V", New Object() {SOURCE_DOC_NO}).Select("", "PINV_NO")
                If rows.Length = 2 Then
                    Dim PINV_NO_1 As String = rows(0).Item("PINV_NO")
                    Dim PINV_NO_2 As String = rows(1).Item("PINV_NO")
                    ASCMAIN1.sql = $"Select PINV_NO, COUNT (*) from ICTPINV2 where PINV_NO IN ('{PINV_NO_1}','{PINV_NO_2}') group by PINV_NO"
                    Dim rowsPINV2() As DataRow = ASCDATA1.GetDataTable().Select("")
                    If rowsPINV2.Length = 1 AndAlso rowsPINV2(0).Item("PINV_NO") = PINV_NO_1 Then
                        Dim row1 As DataRow = rows(0)
                        If row1.Item("VESSEL_NAME") & "" <> "" And row1.Item("BOL_NO") & "" <> "" And row1.Item("CONTAINER_NO") & "" <> "" Then

                        End If
                        If MsgBox($"Found Multiple IPSA Invoices ending with {SOURCE_DOC_NO}" & vbCrLf & vbCrLf & "Would you like to correct this Duplication?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "IMPORTANT - Please look into the cause of this Duplication") = MsgBoxResult.Yes Then

                            BeginTrans()

                            ' execute fix
                            ASCMAIN1.sql = $"
begin declare 
P1 VARCHAR2(6);
P2 VARCHAR2(6);
begin 
P1 := '{PINV_NO_1}';
P2 := '{PINV_NO_2}';
BEGIN DECLARE cursor c1 is
SELECT * FROM ICTPINV1 WHERE PINV_NO = P1;
begin for r1 in c1 loop
update ICTPINV1 set VESSEL_NAME = r1.vessel_name, BOL_NO = r1.bol_no
, CONTAINER_NO= r1.container_no, SHIP_DATE = r1.ship_date, ETA_DATE = r1.eta_date
, last_oper = r1.last_oper, last_date = r1.last_date
WHERE PINV_NO = P2;
delete from ICTPINV1 where PINV_NO = P1;
update ICTPINV1 set PINV_NO = P1 WHERE PINV_NO = P2;
end loop; end; end;
END; END;
"
                            ASCDATA1.ExecuteSQL()

                            ' record event
                            ASCMAIN1.Record_Event("ICTPINV1", PINV_NO_1, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "DUP1", $"Dup {PINV_NO_2} Resolved in Favor of {PINV_NO_1}", PINV_NO_1)

                            ' RE-ESTABLISH rows
                            rows = ASCDATA1.GetDataTable(sqlICTPINV1, "", "V", New Object() {SOURCE_DOC_NO}).Select("")

                            CommitTrans("Duplication Resolved")

                        End If
                    End If
                End If

                If rows.Length = 0 Then
                    EMsg &= vbCr & "Unable to find IPSA Invoice (" & SOURCE_DOC_NO & ")"
                ElseIf rows.Length > 1 Then
                    EMsg &= vbCr & "Found Multiple IPSA Invoices ending with (" & SOURCE_DOC_NO & ")"
                Else

                    Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", Absx1.txtFor("VEND_CODE").Text)
                    Dim PO_COST_VAR_TOL_PCT As Decimal = Val(rowAPTVEND1.Item("PO_COST_VAR_TOL_PCT") & "")
                    Dim PO_COST_VAR_TOL_AMT As Decimal = Val(rowAPTVEND1.Item("PO_COST_VAR_TOL_AMT") & "")

                    Dim PINV_NO As String = rows(0).Item("PINV_NO")
                    Dim rowICTPINV1 As DataRow = Fill_Record("ICTPINV1", PINV_NO)
                    Fill_Records("ICTPINV2", PINV_NO)

                    For Each rowICTIREC2 As DataRow In dst.Tables("ICTIREC2").Select("")
                        Dim PO_ORDER_NO As String = rowICTIREC2.Item("PO_ORDER_NO")
                        Dim PO_ORDER_LNO As Int64 = Val(rowICTIREC2.Item("PO_ORDER_LNO") & "")

                        Dim rowPOTORDR2 As DataRow = LookUp("POTORDR2", New String() {PO_ORDER_NO, PO_ORDER_LNO})
                        Dim ITEM_CODE_ALT As String = rowPOTORDR2.Item("ITEM_CODE_ALT") & ""
                        Dim ITEM_CODE As String = rowICTIREC2.Item("ITEM_CODE")

                        With rowICTIREC2
                            .Item("PINV_QTY") = DBNull.Value
                            .Item("PINV_COST") = DBNull.Value
                            .Item("PO_COST_VAR_AMT") = DBNull.Value
                            .Item("PO_COST_VAR_PCT") = DBNull.Value
                            .Item("PO_COST_VAR_AMT_IND") = DBNull.Value
                            .Item("PO_COST_VAR_PCT_IND") = DBNull.Value
                        End With

                        Dim ITEM_CODE_X As String = ITEM_CODE_ALT
                        If ITEM_CODE_X = "" Then ITEM_CODE_X = ITEM_CODE

                        ITEM_CODE_X = ITEM_CODE  ' change made 09/01 to use the real item code in ICTPINV2

                        Dim SQLW As String = "PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & CStr(PO_ORDER_LNO) & " and ITEM_CODE = '" & ITEM_CODE_X & "'"
                        Dim row() As DataRow = dst.Tables("ICTPINV2").Select(SQLW)
                        If row.Length = 0 Then
                            EMsg &= vbCr & "Unable to match PO-Line " & PO_ORDER_NO & ":" & CStr(PO_ORDER_LNO) & ":" & ITEM_CODE_X
                        ElseIf row.Length > 1 Then
                            EMsg &= vbCr & "Multiple matches for PO-Line " & PO_ORDER_NO & ":" & CStr(PO_ORDER_LNO) & ":" & ITEM_CODE_X
                        Else
                            row(0).Item("QTY_REC") = rowICTIREC2.Item("QTY_REC")
                            row(0).Item("QTY_INV") = row(0).Item("PINV_QTY")
                            row(0).Item("RECEIPT_NO") = rowICTIREC2.Item("RECEIPT_NO")
                            row(0).Item("RECEIPT_LNO") = rowICTIREC2.Item("RECEIPT_LNO")
                            row(0).Item("COST_CATGY_CODE") = rowICTIREC2.Item("COST_CATGY_CODE")
                            ' NOTE THAT SINCE ICTIREC2.QTY_INV ALWAYS STARTS WITH 0, THAT THIS FIELD IS ALWAYS THE SAME AS QTY_REC
                            row(0).Item("QTY_REC_NOT_INV") = Val(rowICTIREC2.Item("QTY_REC") & "") - Val(rowICTIREC2.Item("QTY_INV") & "")

                            Dim PO_COST As Decimal = Val(rowICTIREC2.Item("PO_COST") & "")
                            Dim PINV_COST As Decimal = Val(row(0).Item("PINV_COST") & "")
                            Dim PINV_QTY As Int32 = Val(row(0).Item("PINV_QTY") & "")
                            Dim PO_COST_VAR_AMT As Decimal = (PO_COST - PINV_COST) * PINV_QTY
                            Dim PO_COST_VAR_PCT As Decimal = 0
                            If PO_COST_VAR_AMT <> 0 Then
                                If PO_COST = 0 Then
                                    PO_COST_VAR_PCT = 100
                                Else
                                    PO_COST_VAR_PCT = 100 * ((PO_COST_VAR_AMT / PINV_QTY) / PO_COST)
                                End If
                            End If

                            If Abs(PO_COST_VAR_AMT) >= 1 Then
                                With rowICTIREC2
                                    .Item("PINV_QTY") = PINV_QTY
                                    .Item("PINV_COST") = PINV_COST
                                    .Item("PO_COST_VAR_AMT") = PO_COST_VAR_AMT
                                    .Item("PO_COST_VAR_PCT") = PO_COST_VAR_PCT
                                    If Abs(PO_COST_VAR_AMT) > PO_COST_VAR_TOL_AMT Or Abs(PO_COST_VAR_PCT) > PO_COST_VAR_TOL_PCT Then
                                        If Abs(PO_COST_VAR_AMT) > PO_COST_VAR_TOL_AMT Then
                                            .Item("PO_COST_VAR_AMT_IND") = "1"
                                        End If
                                        If Abs(PO_COST_VAR_PCT) > PO_COST_VAR_TOL_PCT Then
                                            .Item("PO_COST_VAR_PCT_IND") = "1"
                                        End If
                                    End If
                                End With

                            End If
                        End If

                    Next

                    If reverse Then
                    Else
                        If dst.Tables("ICTIREC2").Select("PO_COST_VAR_AMT_IND = '1'").Length > 0 _
                        Or dst.Tables("ICTIREC2").Select("PO_COST_VAR_PCT_IND = '1'").Length > 0 Then
                            EMsg &= vbCr & "PO Cost (to Invoice) Variances exceed Tolerances"
                            lblVARIANCE_EXPLANATION.Visible = False
                            txtVARIANCE_EXPLANATION.Visible = False
                            Variance_Column_Visibility(True)

                            'Generate_Variance_Email() ' use this to generate an email for demo
                            'Stop
                        Else
                            'If dst.Tables("ICTIREC2").Select("ABS(ISNULL(PO_COST_VAR_AMT,0)) >=1").Length > 0 _
                            'Or dst.Tables("ICTIREC2").Select("ABS(ISNULL(PO_COST_VAR_PCT,0)) >=1").Length > 0 Then

                            If dst.Tables("ICTIREC2").Select("ISNULL(PO_COST_VAR_AMT,0) <> 0").Length > 0 _
                            Or dst.Tables("ICTIREC2").Select("ISNULL(PO_COST_VAR_PCT,0) <> 0").Length > 0 Then
                                If Absx1.txtFor("VARIANCE_EXPLANATION").Text = "" Then
                                    EMsg &= vbCr & "There are PO Cost (to Invoice) Variances - Explanation Required"
                                    lblVARIANCE_EXPLANATION.Visible = True
                                    txtVARIANCE_EXPLANATION.Visible = True
                                    Set_Read_Only_for_ctl(txtVARIANCE_EXPLANATION, False)
                                    Variance_Column_Visibility(True)
                                Else
                                    Generate_Variance_Email()
                                End If
                            End If
                        End If
                    End If

                    If EMsg = "" Then
                        If reverse Then
                        Else
                            If dst.Tables("ICTPINV2").Select("ISNULL(PINV_QTY,0) <> ISNULL(QTY_REC,0)").Length > 0 Then
                                If MsgBox("Qty Mis-Match between Receipt and Invoice" & vbCrLf & vbCrLf & "OK to Continue Anyway?",
                                  MsgBoxStyle.YesNo,
                                  "Some Items on this Receipt have qtys that do not match the qtys invoiced by IPSA") = MsgBoxResult.No Then
                                    EMsg &= vbCr & "Qty Mis-Match between Receipt and Invoice"
                                End If
                            End If
                        End If
                    End If
                End If

            End If


        End If
    End Sub

    Private Sub grdICT3PLTX_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdICT3PLTX.AfterRowUpdate
        'If e.Row.Band.Key = "ICT3PLTX" Then
        '    ASCMAIN1.sql = "Update ICT3PLTX Set DTDATE = :PARM1, DTTIME = :PARM2 where PROC_KEY = :PARM3"
        '    ASCDATA1.ExecuteSQL()
        '    MsgBox("Click Refresh")
        'End If
    End Sub

    Private Sub grdICT3PLTX_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdICT3PLTX.BeforeRowUpdate
        If e.Row.Band.Key = "ICT3PLTX" Then
            Dim DTDATE As String = e.Row.Cells("DTDATE").Value
            Dim DTTIME As String = e.Row.Cells("DTTIME").Value
            Dim DTUSER As String = e.Row.Cells("DTUSER").Value

            Dim DTDATE_ORIG As String = e.Row.Cells("DTDATE").OriginalValue
            Dim DTTIME_ORIG As String = e.Row.Cells("DTTIME").OriginalValue

            ASCMAIN1.sql = "Update EDTTRXNA Set DTDATE = :PARM1, DTTIME = :PARM2 where DTDATE = :PARM3 and DTTIME = :PARM4 and DTUSER = :PARM5"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVVV", New String() {DTDATE, DTTIME, DTDATE_ORIG, DTTIME_ORIG, DTUSER})

            EnforceConstraints(False)
            Fill_Records("ICT3PLTX")
            Fill_Records("ICT3PLTY")
            Fill_Records("ICT3PLTZ")
            EnforceConstraints(True)

            e.Cancel = True
        End If
    End Sub

    Private Sub chkReceiptsSummaryOnly_CheckedChanged(sender As Object, e As EventArgs) Handles chkReceiptsSummaryOnly.CheckedChanged
        If chkReceiptsSummaryOnly.Checked Then
            grdICTIRECX.DisplayLayout.ViewStyle = ViewStyle.SingleBand
        Else
            grdICTIRECX.DisplayLayout.ViewStyle = ViewStyle.MultiBand
        End If
    End Sub

    Private Sub grdICT3PLTX_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdICT3PLTX.InitializeRow
        If e.Row.Band.Index = 0 Then
            If e.Row.Cells("DTADJC").Value = "A03" Then
                e.Row.Cells("DTADJC").Appearance.ForeColor = System.Drawing.Color.Blue
            ElseIf e.Row.Cells("DTADJC").Value = "A21" Then
                e.Row.Cells("DTADJC").Appearance.ForeColor = System.Drawing.Color.Red
            Else
                e.Row.Cells("DTADJC").Appearance.ForeColor = System.Drawing.Color.Empty
            End If
        End If
    End Sub

    Sub Variance_Column_Visibility(visibility As Boolean)
        For Each C As String In New String() _
            {"PINV_QTY", "PINV_COST", "PO_COST_VAR_AMT", "PO_COST_VAR_PCT", "PO_COST_VAR_AMT_IND", "PO_COST_VAR_PCT_IND"}
            grdICTIREC2.DisplayLayout.Bands(0).Columns(C).Hidden = Not visibility
        Next

    End Sub

    Sub Generate_Variance_Email()

        Dim RECEIPT_NO As String = Absx1.txtFor("RECEIPT_NO").Text
        Dim VEND_CODE As String = Absx1.txtFor("VEND_CODE").Text


        Dim COLs As New List(Of String)
        For Each gcol As UltraWinGrid.UltraGridColumn In grdICTIREC2.DisplayLayout.Bands(0).Columns
            If New String() {"PINV_QTY", "PINV_COST", "PO_COST_VAR_AMT", "PO_COST_VAR_PCT", "PO_COST_VAR_AMT_IND", "PO_COST_VAR_PCT_IND"}.Contains(gcol.Key) Then
                ' do not hide
            ElseIf New String() {"RECEIPT_LNO", "ITEM_CODE", "ITEM_DESC", "PO_ORDER_NO", "PO_ORDER_LNO", "PO_COST"}.Contains(gcol.Key) Then
                ' do not hide
            Else
                If Not gcol.Hidden Then
                    gcol.Hidden = True
                    COLs.Add(gcol.Key)
                End If
            End If
        Next

        ' EXPORT GRID TO EXCEL

        'ASCMAIN1.JOB_STREAM_CODE = "x" ' dangerous cluge to avoid showing the excel document
        ' - we really need to add an optional parameter to avoid showing the document,
        '   but this would require a full rebuild and deployment - DONE, now using the optional parameter below
        Dim FILENAME As String = Excel_Export_to_SSG(grdICTIREC2, False)
        'ASCMAIN1.JOB_STREAM_CODE = "" ' dangerous cluge to avoid showing the excel document

        Dim attachments As New Dictionary(Of String, String)
        attachments.Add(FILENAME, FILENAME)

        For Each gcol As UltraWinGrid.UltraGridColumn In grdICTIREC2.DisplayLayout.Bands(0).Columns
            If COLs.Contains(gcol.Key) Then
                gcol.Hidden = False
            End If
        Next

        Dim ALERT_MESSAGE As String = "PO Cost Variance Explanation: " & Absx1.txtFor("VARIANCE_EXPLANATION").Text & vbCrLf

        For Each rowICTIREC2 As DataRow In dst.Tables("ICTIREC2").Select()
            With rowICTIREC2
                Dim ITEM_CODE As String = .Item("ITEM_CODE")
                Dim ITEM_DESC As String = .Item("ITEM_DESC")

                Dim QTY_REC As Int32 = Val(.Item("QTY_REC") & "")
                Dim PO_COST As Decimal = Val(.Item("PO_COST") & "")

                Dim PINV_QTY As Int32 = Val(.Item("PINV_QTY") & "")
                Dim PINV_COST As Decimal = Val(.Item("PINV_COST") & "")
                Dim PO_COST_VAR_AMT As Decimal = Val(.Item("PO_COST_VAR_AMT") & "")
                Dim PO_COST_VAR_PCT As Decimal = Val(.Item("PO_COST_VAR_PCT") & "")

                If PO_COST_VAR_AMT <> 0 Then

                    'PO_COST_VAR_AMT_IND
                    'PO_COST_VAR_PCT_IND

                    'Dim COST_LIST_CODE As String = rowPOTORDR1.ITEM("COST_LIST_CODE") & ""
                    'Dim COST_CLASS_CODE As String = rowPOTORDR1.ITEM("COST_LIST_CODE") & ""

                    ALERT_MESSAGE &= vbCrLf _
                        & vbCrLf & "PO Cost Variance " _
                        & vbCrLf & " Item: " & ITEM_CODE & ":" & ITEM_DESC _
                        & vbCrLf & " - PO Cost: " & vbTab & Format(PO_COST, "$#.0000") _
                        & vbCrLf & " - Inv Cost: " & vbTab & Format(PINV_COST, "$#.0000") _
                        & vbCrLf & " - Qty Rec: " & vbTab & Format(QTY_REC, "#,##0") _
                        & vbCrLf & " - Qty Inv: " & vbTab & Format(PINV_QTY, "#,##0") _
                        & vbCrLf & " - Var Amt: " & vbTab & Format(PO_COST_VAR_AMT, "$#,##0") _
                        & vbCrLf & " - Var Pct: " & vbTab & Format(PO_COST_VAR_PCT, "#,##0%")

                End If
            End With

        Next

        Dim ALERT_SUBJECT As String = ""
        Dim ALERT_EMAIL As String = ROWs("ICTPARM1").Item("IC_PARM_EMAIL_ALERT_COST") & ""

        Dim rowTATALRT1 As DataRow = dst.Tables("TATALRT1").NewRow
        With rowTATALRT1
            Dim ALERT_NO As String = ASCMAIN1.Next_Control_No("TATALRT1.ALERT_NO")
            .Item("ALERT_NO") = ALERT_NO
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("FORM_NAME") = "ICTIREC1"
            .Item("FORM_KEY") = ALERT_NO
            .Item("ALERT_EMAIL") = ALERT_EMAIL
            .Item("ALERT_EML") = "1"

            .Item("ALERT_EML_DATE") = DATETIME_STAMP
            ALERT_SUBJECT = "PO Cost Variance Alert for Receipt No: " & RECEIPT_NO
            .Item("ALERT_SUBJECT") = Mid(ALERT_SUBJECT, 1, 200)
            'ALERT_MESSAGE = "Control No: " & ALERT_NO & vbCrLf & "Price Change: " & RECEIPT_NO & ALERT_MESSAGE
            .Item("ALERT_MESSAGE") = Mid(ALERT_MESSAGE, 1, 2000)
        End With
        dst.Tables("TATALRT1").Rows.Add(rowTATALRT1)

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        EMAIL_ADDRESSs.Add(ALERT_EMAIL, "PO Cost Variance Auditor")

        ALERT_MESSAGE = "Variance Explanation: " & Absx1.txtFor("VARIANCE_EXPLANATION").Text

        Dim SEND_NO As String = ""
        If ASCMAIN1.Running_in_VS Then
            SEND_NO = "TESTING"
            '   Stop
        Else
            SEND_NO = ASCMAIN1.TACMAIN1.Send_email _
                    (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, attachments,
                    ALERT_SUBJECT, "PO_CPIEXC", True, False, RECEIPT_NO, VEND_CODE, "PO Receipt", ALERT_MESSAGE)
        End If

        rowTATALRT1.Item("SEND_NO") = SEND_NO
        Update_Record_TDA("TATALRT1")

        TAC.TACMAIN1.Record_Event("ICTIREC1", RECEIPT_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "CPIEXC", "PO Cost to Inv Cost Exception Alert emailed", SEND_NO, "ICTIREC1")

    End Sub


End Class