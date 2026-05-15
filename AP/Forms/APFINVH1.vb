Imports System.Math
Imports Infragistics.Win.UltraWinGrid

Public Class APFINVH1
    Dim rowAPTVEND1 As DataRow
    Dim rowAPTINVH1 As DataRow
    Dim rowAPTVEND5 As DataRow
    Dim ICTIREC1 As String
    Dim sql_APTINVR2 As String
    Dim sql_APTVEND1 As String
    Dim batch_update As Boolean = False
    Dim BANK_LAST_CHECK_NO As String
    Dim BANK_NEXT_CHECK_NO As String
    Dim auto_next_check As Boolean = False
    Private discrepancies_only As Boolean = False
    Dim ApprovalMode As Boolean = False
    Dim sqlAPTSUBM1 As String
    Dim INVOICE_FROM_EMAIL As String

    ' this is where we get emails from apabs
    ' TAC.TACMAIN1.GetEmails(Me)

    'THE APPROVER OF AN INVOICE DEFAULTS TO THE BUYER ASSOCIATED WITH THE VENDOR
    'THIS MAY BE CHANGED IN INVOICE ENTRY
    'ONCE AN INVOICE IS ENTERED, IT IS VISIBLE ON THE APPROVAL SCREEN
    'APPROVAL SCREEN SHOWS ALL INVOICES
    'THAT THE LOGGED IN PERSON IS NOTED AS THE APPROVER FOR
    'THAT ARE OPEN AND PENDING APPROVAL
    'APPROVER MAY APPROVE
    'AN APPROVER WILL HAVE AN INVOICE LIMIT
    'IF APPROVAL IS BEYOND THE APPROVER'S INVOICE LIMIT,
    'THE APPROVER MUST RE-ROUTE TO SOMEONE OF HIGHER AUTHORITY
    'APPROVER MAY FORWARD TO SOME OTHER APPROVER
    'WHILE AN INVOICE IS PENDING FINAL APPROVAL, IT MAY NOT BE SELECTED FOR PYMT

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If MENU_ITEM_OBJECT = "APFINVHI" Then
            InquiryMode = True
        End If
        If MENU_ITEM_OBJECT = "APFINVHA" Then
            ApprovalMode = True
        End If

        Get_PARM("GLTPARM1")
        Get_PARM("APTPARM1")
        Get_PARM("ICTPARM1")

        With UltraExplorerBar1.Groups("Screen Control")
            .Items("Load").Visible = InquiryMode Or ApprovalMode
            .Items("Done").Visible = InquiryMode
            .Items("New").Visible = Not InquiryMode And Not ApprovalMode
            .Items("Edit").Visible = Not InquiryMode And Not ApprovalMode
            .Items("Update").Visible = Not InquiryMode
            .Items("Delete").Visible = Not InquiryMode And Not ApprovalMode
            .Items("Cancel").Visible = Not InquiryMode
            .Items("Print Edit").Visible = Not InquiryMode And Not ApprovalMode
            .Items("New Batch").Visible = Not InquiryMode And Not ApprovalMode
            .Items("Multi-Invoice Edit").Visible = Not InquiryMode And Not ApprovalMode
        End With


        With dst
            ASCMAIN1.sql = "Select * from ASTAUDT1" _
                    & " where ASTAUDT1.TABLE_NAME = 'APTINVH1'" _
                    & "   and ASTAUDT1.KEY_VALUE = :PARM1"
            Create_TDA(.Tables.Add, "ASTAUDTX", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "APTVEND1", "*")
            Create_TDA(.Tables.Add, "APTVEND2", "*")
            Create_TDA(.Tables.Add, "APTVEND5", "*")
            Create_TDA(.Tables.Add, "APTINVH1", "*")

            Create_TDA(.Tables.Add, "APTCHCK1", "*")
            Create_TDA(.Tables.Add, "APTCHCK2", "*")
            Create_TDA(.Tables.Add, "GLTBANK1", "*")

            ASCMAIN1.sql = "Select APTINVH2.*, GLTACCT1.ACCT_DESC from APTINVH2,GLTACCT1 where APTINVH2.ACCT_CODE = GLTACCT1.ACCT_CODE"
            Create_TDA(.Tables.Add, "APTINVH2", "**", 1)

            Create_TDA(.Tables.Add, "APTINVH8", "*", 1)

            ASCMAIN1.sql = "Select APTINVH7.*,APTACRC1.ACCRUAL_CODE,APTACRM1.ACCRUAL_DESC, ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & ", APTACRC1.PO_ORDER_NO, APTACRC1.RECEIPT_NO, APTACRC1.ITEM_CODE, APTACRC1.CTL_TYPE, APTACRC1.COST_CATGY_CODE, APTACRC1.CTL_DATE, APTACRC1.CTL_NOTE" & vbCrLf _
                & " from APTINVH7,APTACRC1,APTACRM1,ICTITEM1" & vbCrLf _
                & " where APTACRC1.CTL_NO = APTINVH7.CTL_NO and APTINVH7.VOUCHER_NO = :PARM1" & vbCrLf _
                & "   and APTACRM1.ACCRUAL_CODE = APTACRC1.ACCRUAL_CODE" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE (+) = APTACRC1.ITEM_CODE" & vbCrLf _
                & "   and APTACRC1.VEND_CODE_ACC = :PARM2"
            Create_TDA(.Tables.Add, "APTINVH7", "**", 0, True, "Vv", 2)
            With .Tables("APTINVH7").Columns
                '.Add("AMT_REC_NOT_INV", GetType(System.Decimal), "QTY_REC_NOT_INV * PO_COST")
                .Add("TOTAL_VAR", GetType(System.Decimal), "IIF(ISNULL(CTL_TYPE,'?') = 'M',0,ISNULL(TOTAL_INV,0) - ISNULL(TOTAL_ACC,0))")
            End With

            ASCMAIN1.sql = "Select APTACRC1.*" & vbCrLf _
                & " from APTACRC1" & vbCrLf _
                & " where (APTACRC1.VOUCHER_NO = :PARM1" & vbCrLf _
                & " or (APTACRC1.VEND_CODE_ACC = :PARM2 and (APTACRC1.CTL_STATUS = '0' OR (APTACRC1.CTL_STATUS = '1' and NVL(APTACRC1.PPD_IND,'0') = '1' and NVL(APTACRC1.PPD_MATCHED,'0') = '0'))))"
            '& " where (APTACRC1.VOUCHER_NO = :PARM1 or (APTACRC1.VOUCHER_NO is Null and NVL(APTACRC1.COST_ACC,0) <> 0))"
            Create_TDA(.Tables.Add, "APTACRC1", "**", 0, True, "VV", 1)

            ASCMAIN1.sql = "Select APTINVH1.VOUCHER_NO, APTINVH1.INV_NUM" _
            & ", APTINVH1.INV_DATE" _
            & ", NVL(APTINVH1.INV_BALANCE,0) INV_BALANCE" _
            & ", NVL(APTINVH1.INV_DISC_AMT,0) INV_DISC_AMT" _
            & " from APTINVH1 " _
            & " where VEND_CODE = :PARM1 and VOUCHER_NO <> :PARM2 " _
            & " and INV_STATUS = 'O' and INV_TYPE in ('A','I','C')"
            Create_TDA(.Tables.Add, "APTINVHX", "**", 0, False, "VV", 1)
            .Tables("APTINVHX").Columns.Add("INV_PAYMENTS", GetType(System.Decimal), "INV_BALANCE - INV_DISC_AMT")
            .Tables("APTINVHX").Columns.Add("SELECTED", GetType(System.String))
            .Tables("APTINVHX").Columns("SELECTED").DefaultValue = "0"

            ASCMAIN1.sql = "Select APTINVH1.VOUCHER_NO" _
            & ", APTINVH1.INV_NUM, APTINVH1.INV_DATE, APTINVH1.INV_AMT" _
            & ", APTINVH1.INV_TYPE, APTINVH1.TERM_CODE" _
            & ", APTINVH1.INV_PYMT_METHOD, APTINVH1.INV_PYMT_CYCLE" _
            & ", APTINVH1.BANK_CODE, APTINVH1.INV_REF" _
            & ", APTINVH1.VEND_ALT_CODE" _
            & " from APTINVH1 where APTINVH1.INV_STATUS in ('H','O') and APTINVH1.VEND_CODE = :PARM1"
            Create_TDA(.Tables.Add, "APTINVHM", "**", 0, False, "V")
            .Tables("APTINVHM").Columns.Add("UPDATE_STATUS", GetType(System.String))
            .Tables("APTINVHM").Columns.Add("UPDATE_MESSAGE", GetType(System.String))

            ASCMAIN1.sql = "Select APTINVH1.VOUCHER_NO" _
            & ", APTINVH1.VEND_CODE, APTVEND1.VEND_NAME" _
            & ", APTINVH1.INV_NUM, APTINVH1.INV_DATE, APTINVH1.INV_AMT" _
            & ", APTINVH1.INV_TYPE, APTINVH1.INV_STATUS, APTINVH1.TERM_CODE" _
            & ", APTINVH1.INV_PYMT_METHOD, APTINVH1.INV_PYMT_CYCLE" _
            & ", APTINVH1.BANK_CODE, APTINVH1.INV_REF" _
            & ", APTINVH1.CHECK_NUM, APTINVH1.CHECK_DATE, APTINVH1.VEND_ALT_CODE, APTINVH1.POST_CODE" _
            & " from APTINVH1,APTVEND1 where APTVEND1.VEND_CODE = APTINVH1.VEND_CODE"
            Create_TDA(.Tables.Add, "APTINVHB", "**", 0, False)
            .Tables("APTINVHB").Columns.Add("UPDATE_STATUS", GetType(System.String))
            .Tables("APTINVHB").Columns.Add("UPDATE_MESSAGE", GetType(System.String))

            ASCMAIN1.sql = "Select APTINVH1.* from APTINVH1"
            Create_TDA(.Tables.Add, "APTINVR1", "**", 0, False)
            .Tables("APTINVR1").Columns.Add("CHECK_AMT", GetType(System.Decimal))
            .Tables("APTINVR1").Columns.Add("CHECK_AMT_OTHERS", GetType(System.Decimal))
            .Tables("APTINVR1").Columns.Add("INV_AMT_GL", GetType(System.Decimal))

            ASCMAIN1.sql = "Select APTINVH2.*, DECODE(APTINVH2.INV_LTYP,NULL,APTINVH2.INV_LINE_AMT,0) INV_LINE_AMT_GL, GLTACCT1.ACCT_DESC from APTINVH2,GLTACCT1 where APTINVH2.ACCT_CODE = GLTACCT1.ACCT_CODE"
            Create_TDA(.Tables.Add, "APTINVR2", "**", 0, False)
            .Tables("APTINVR2").Columns.Add("OPS_YYYYPP", GetType(System.String))

            .Relations.Add("APTINVR2",
            .Tables("APTINVR1").Columns("VOUCHER_NO"),
            .Tables("APTINVR2").Columns("VOUCHER_NO"))

            .Tables("APTINVR1").Columns("INV_AMT_GL").Expression = "SUM(CHILD(APTINVR2).INV_LINE_AMT_GL)"
            .Tables("APTINVR2").Columns("OPS_YYYYPP").Expression = "PARENT(APTINVR2).OPS_YYYYPP"

            ASCMAIN1.sql = "Select APTINVH5.*" & vbCrLf _
                & ", ICTITEM1.ITEM_DESC, ICTITEM1.COLLECTION_CODE, ICTCOST1.ACCT_CODE_PPV" & vbCrLf _
                & ", ICTIREC2.QTY_REC, ICTIREC2.QTY_INV, ICTIREC2.PO_COST" & vbCrLf _
                & ", ICTIREC2.ITEM_CODE, ICTIREC2.ITEM_UOM" & vbCrLf _
                & ", NVL(ICTIREC2.QTY_REC,0) * NVL(ICTCOSTA.ITEM_COST_VCOST,0) VCOST" & vbCrLf _
                & ", NVL(ICTIREC2.QTY_REC,0) * NVL(ICTCOSTA.ITEM_COST_LANDG,0) LANDG" & vbCrLf _
                & ", NVL(ICTIREC2.QTY_REC,0) * NVL(ICTCOSTA.ITEM_COST_TOOLG,0) TOOLG" & vbCrLf _
                & ", NVL(ICTIREC2.QTY_REC,0) * NVL(ICTCOSTA.ITEM_COST_OVRHD,0) OVRHD" & vbCrLf _
                & " from ICTIREC2,APTINVH5,ICTITEM1,ICTCOST1,ICTCOSTA" & vbCrLf _
                & " where APTINVH5.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
                & "   and APTINVH5.RECEIPT_LNO = ICTIREC2.RECEIPT_LNO" & vbCrLf _
                & "   and APTINVH5.VOUCHER_NO = :PARM1" & vbCrLf _
                & "   and ICTCOSTA.OPS_YYYYPP = ICTIREC2.OPS_YYYYPP" & vbCrLf _
                & "   and ICTCOSTA.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = ICTIREC2.ITEM_CODE" _
                & "   and ICTCOST1.COST_CATGY_CODE (+) = ICTITEM1.COST_CATGY_CODE"

            Create_TDA(.Tables.Add, "APTINVH5", "**", 0, True, "V", 2)
            With .Tables("APTINVH5").Columns
                .Add("AMT_REC", GetType(System.Decimal), "QTY_REC * PO_COST")
                .Add("AMT_INV", GetType(System.Decimal), "INV_QTY * INV_COST")
                .Add("AMT_REC_NOT_INV", GetType(System.Decimal), "QTY_REC_NOT_INV * PO_COST")
                .Add("AMT_REC_NOT_INV_OFFSET", GetType(System.Decimal), "IIF(CLOSE_LINE='1',QTY_REC_NOT_INV * PO_COST,INV_QTY * PO_COST)")
                .Add("QTY_VAR", GetType(System.Int64), "IIF(CLOSE_LINE='0',0,ISNULL(INV_QTY,0) - ISNULL(QTY_REC_NOT_INV,0))")
                '.Add("AMT_VAR", GetType(System.Decimal), "(INV_QTY * INV_COST) - (QTY_REC * PO_COST)")
                '.Add("AMT_VAR", GetType(System.Decimal), "AMT_INV - AMT_REC")
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

            ASCMAIN1.sql = "Select * from TATEVNT1 where TABLE_NAME = 'APTINVH1' and TABLE_KEY = :PARM1"
            Create_TDA(.Tables.Add, "TATEVNT1", "**", 0, False, "V", 0)

            .Tables.Add("APTINVH5_SUM")
            With .Tables("APTINVH5_SUM")
                .Columns.Add("RECEIPT_NO", GetType(System.String))
                .Columns.Add("RECEIPT_DATE", GetType(System.DateTime))
                .Columns.Add("PO_ORDER_NO", GetType(System.String))
                .PrimaryKey = New DataColumn() { .Columns("RECEIPT_NO")}
            End With

            .Relations.Add("APTINVH5", .Tables("APTINVH5_SUM").Columns("RECEIPT_NO"), .Tables("APTINVH5").Columns("RECEIPT_NO"))

            With .Tables("APTINVH5_SUM")
                .Columns.Add("QTY_REC", GetType(System.Int32), "SUM(Child(APTINVH5).QTY_REC)")
                .Columns.Add("AMT_REC", GetType(System.Decimal), "SUM(Child(APTINVH5).AMT_REC)")
                .Columns.Add("QTY_INV", GetType(System.Int32), "SUM(Child(APTINVH5).INV_QTY)")
                .Columns.Add("AMT_INV", GetType(System.Decimal), "SUM(Child(APTINVH5).AMT_INV)")
                .Columns.Add("QTY_REC_NOT_INV", GetType(System.Int64), "SUM(Child(APTINVH5).QTY_REC_NOT_INV)")
                .Columns.Add("AMT_REC_NOT_INV", GetType(System.Decimal), "SUM(Child(APTINVH5).AMT_REC_NOT_INV)")
                .Columns.Add("AMT_REC_NOT_INV_OFFSET", GetType(System.Decimal), "SUM(Child(APTINVH5).AMT_REC_NOT_INV_OFFSET)")
                .Columns.Add("QTY_VAR", GetType(System.Int32), "SUM(Child(APTINVH5).QTY_VAR)")
                .Columns.Add("AMT_VAR", GetType(System.Decimal), "SUM(Child(APTINVH5).AMT_VAR)")
            End With

            Dim sql As String = "Select ICTIREC1.* from ICTIREC1 where ROWNUM < 1"
            ICTIREC1 = ASCMAIN1.Temp_Table(sql)
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIREC1 & " Add Primary Key (RECEIPT_NO)")
            'ASCDATA1.ExecuteSQL("Alter Table " & ICTIREC1 & " Add EXT_PO_COST NUMBER (13,2)")

            ASCMAIN1.sql = "Select ICTIREC1.*, X.EXT_PO_COST, X.QTY_REC_NOT_INV" & vbCrLf _
                & " from " & ICTIREC1 & " ICTIREC1" & vbCrLf _
                & ", (Select ICTIREC2.RECEIPT_NO, SUM (ICTIREC2.QTY_REC * ICTIREC2.PO_COST) EXT_PO_COST" & vbCrLf _
                & ", Sum (NVL(ICTIREC2.QTY_REC,0) - NVL(ICTIREC2.QTY_INV,0)) QTY_REC_NOT_INV" & vbCrLf _
                & " from ICTIREC2," & ICTIREC1 & " ICTIREC1" & vbCrLf _
                & " where ICTIREC2.RECEIPT_NO = ICTIREC1.RECEIPT_NO group by ICTIREC2.RECEIPT_NO) X" & vbCrLf _
                & " where X.RECEIPT_NO (+) = ICTIREC1.RECEIPT_NO"
            Create_TDA(.Tables.Add, "ICTIREC1", "**", 0, False, "", 1)
            .Tables("ICTIREC1").Columns("QTY_REC_NOT_INV").DataType = GetType(System.Int64)

            ASCMAIN1.sql = "Select ICTIREC2.*, ICTITEM1.ITEM_DESC" & vbCrLf _
                & ", ICTIREC2.PO_COST * ICTIREC2.QTY_REC EXT_PO_COST " & vbCrLf _
                & ", NVL(ICTIREC2.QTY_REC,0) - NVL(ICTIREC2.QTY_INV,0) QTY_REC_NOT_INV " & vbCrLf _
                & ", NVL(ICTIREC2.QTY_REC,0) * NVL(ICTCOSTA.ITEM_COST_VCOST,0) VCOST" & vbCrLf _
                & ", NVL(ICTIREC2.QTY_REC,0) * NVL(ICTCOSTA.ITEM_COST_LANDG,0) LANDG" & vbCrLf _
                & ", NVL(ICTIREC2.QTY_REC,0) * NVL(ICTCOSTA.ITEM_COST_TOOLG,0) TOOLG" & vbCrLf _
                & ", NVL(ICTIREC2.QTY_REC,0) * NVL(ICTCOSTA.ITEM_COST_OVRHD,0) OVRHD" & vbCrLf _
                & " from ICTIREC2,ICTITEM1,ICTCOSTA " & vbCrLf _
                & " where ICTIREC2.ITEM_CODE = ICTITEM1.ITEM_CODE " & vbCrLf _
                & "   and ICTCOSTA.OPS_YYYYPP = ICTIREC2.OPS_YYYYPP" & vbCrLf _
                & "   and ICTCOSTA.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf _
                & "   and ICTIREC2.RECEIPT_NO = :PARM1"
            Create_TDA(.Tables.Add, "ICTIREC2", "**", 0, False, "V", 2)
            .Tables("ICTIREC2").Columns("QTY_REC_NOT_INV").DataType = GetType(System.Int64)

            Create_Relation("ICTIREC1", "ICTIREC2", "RECEIPT_NO")

            '.Tables("ICTIREC1").Columns.Add("EXT_PO_COST", GetType(System.Decimal), "SUM (CHILD.EXT_PO_COST)")
            '.Tables("ICTIREC1").Columns.Add("QTY_REC_NOT_INV", GetType(System.Int64), "SUM (CHILD.QTY_REC_NOT_INV)")


            'ASCMAIN1.sql = "" _
            '    & "Select ICTIREC2.ITEM_CODE, ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO" _
            '    & ", ICTIREC2.QTY_REC * ICTCOSTA.ITEM_COST_VCOST VCOST" _
            '    & ", ICTIREC2.QTY_REC * ICTCOSTA.ITEM_COST_TOOLG LANDG" _
            '    & ", ICTIREC2.QTY_REC * ICTCOSTA.ITEM_COST_TOOLG TOOLG" _
            '    & ", ICTIREC2.QTY_REC * ICTCOSTA.ITEM_COST_TOOLG OVRHD" _
            '    & " from ICTCOSTA,ICTIREC2,ICTIREC1" _
            '    & " where ICTCOSTA.OPS_YYYYPP = ICTIREC1.OPS_YYYYPP AND ICTCOSTA.ITEM_CODE = ICTIREC2.ITEM_CODE" _
            '    & "   and ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO" _
            '    & "   and ICTIREC2.RECEIPT_NO in (Select RECEIPT_NO from " & ICTIREC1 & ")"
            'Create_TDA(.Tables.Add, "ICTIRECD", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select GLTPARM2.* " _
             & " from GLTPARM2 " _
             & " where OPS_YYYYPP = " _
             & " (Select Min(OPS_YYYYPP) from GLTPARM2 " _
             & "  where GLTPARM2.PRD_END_DATE >= :PARM1)"
            Create_TDA(.Tables.Add, "GLTPARM2", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select APTINVH1.VOUCHER_NO, APTINVH1.INV_DATE, APTINVH1.INV_AMT, APTINVH1.INV_RECUR_GEN " _
             & " from APTINVH1 " _
             & " where VOUCHER_NO_RECUR = :PARM1"
            Create_TDA(.Tables.Add, "APTINVHR", "**", 0, False, "V", 1)

            Create_TDA(.Tables.Add, "GLTDIST1", "*", 0, False)

            Create_TDA(.Tables.Add, "ICTCOLL1", "*", 0, False)
            Fill_Records("ICTCOLL1")

            ASCMAIN1.sql = "Select * from ASTATTA2" & vbCrLf _
                & " where TABLE_NAME = 'APTINVH1' AND COLUMN_NAME = 'VOUCHER_NO'" & vbCrLf _
                & "   and CODE_VALUE = :PARM1"
            Create_TDA(.Tables.Add, "ASTATTA2", "**", 0, True, "V")

            Create_TDA(.Tables.Add, "APTACRM1", "*", 0)

            'ASCMAIN1.sql = "Select * from APTSUBM1" & vbCrLf
            'Create_TDA(.Tables.Add, "APTSUBM1", "**", 0, True, "V")
            'Fill_Records("APTSUBM1")


            sqlAPTSUBM1 = "Select APTSUBM1.*, APTINVH1.INV_STATUS from APTSUBM1,APTINVH1 where APTINVH1.VOUCHER_NO (+) = APTSUBM1.VOUCHER_NO" & vbCrLf
            ASCMAIN1.sql = sqlAPTSUBM1 ' & "  and POTPACK1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "APTSUBM1", "**", 0, True, "")
            '.Tables("APTSUBM1").Columns.Add("SEL")
            '.Tables("APTSUBM1").Columns("SEL").DefaultValue = "0"

            Create_TDA(.Tables.Add, "ICTCOST1", "*", 0, False)
            Fill_Records("ICTCOST1")

            Create_TDA(.Tables.Add, "APTCHCK5", "*")
            .Tables("APTCHCK5").Columns.Add("VEND_BANK_ACCT_ID_DECRYPTED")

        End With

        Fill_Records("APTACRM1")

        grdAPTINVHB.Dock = DockStyle.Fill
        grdAPTINVR1.Dock = DockStyle.Fill
        tabMain.Dock = DockStyle.Fill

        ASCMAIN1.sql = "Update ICTIREC1 set ACCRUAL_STATUS = '1', VOUCHER_NO = :PARM2 where RECEIPT_NO = :PARM1"
        Create_Update_Command("ICTIREC1", "VV")

        grdAPTINVH2.DisplayLayout.NewColumnLoadStyle = NewColumnLoadStyle.Show
        grdAPTINVH2.DataSource = dst.Tables("APTINVH2")
        With grdAPTINVH2.DisplayLayout.Bands(0).Columns("COST_CTR_CODE")
            .Width = 100
            .Header.Caption = "Cost Ctr"
            .Header.VisiblePosition = grdAPTINVH2.DisplayLayout.Bands(0).Columns("INV_COMMENT_DTL").Header.VisiblePosition
        End With
        grdAPTINVH2.DisplayLayout.Bands(0).Columns("DETL_CTL_TYPE").Hidden = True
        grdAPTINVH2.DisplayLayout.Bands(0).Columns("DETL_CTL_NO").Hidden = True
        ASCMAIN1.Add_Value_List(grdAPTINVH2, "COST_CTR_CODE", "Select COST_CTR_CODE, COST_CTR_DESC from GLTCCTR1")

        Dim dvw As DataView = DirectCast(grdAPTINVH2.DataSource, DataTable).DefaultView
        dvw.RowFilter = "INV_LTYP is Null or INV_LTYP = 'A'"

        grdICTIREC1.DataSource = dst.Tables("ICTIREC1")
        grdAPTINVH5.DataSource = dst.Tables("APTINVH5")
        grdAPTINVR1.DataSource = dst.Tables("APTINVR1")
        grdAPTINVH5_SUM.DataSource = dst.Tables("APTINVH5_SUM")
        grdASTAUDTX.DataSource = dst.Tables("ASTAUDTX")
        grdAPTINVHX.DataSource = dst.Tables("APTINVHX")
        grdTATEVNT1.DataSource = dst.Tables("TATEVNT1")
        grdAPTINVHR.DataSource = dst.Tables("APTINVHR")
        grdAPTINVHB.DataSource = dst.Tables("APTINVHB")
        grdAPTINVHM.DataSource = dst.Tables("APTINVHM")
        grdAPTINVH8.DataSource = dst.Tables("APTINVH8")
        grdAPTINVH7.DataSource = dst.Tables("APTINVH7")
        grdAPTACRC1.DataSource = dst.Tables("APTACRC1")
        grdAPTSUBM1.DataSource = dst.Tables("APTSUBM1")


        Set_SEGS(grdAPTINVH2, "APTINVH2")
        Set_SEGS(grdAPTINVR1, "APTINVR2")

        Bind_Controls(Me, "APTVEND1")
        Bind_Controls(Me, "APTVEND2")
        Bind_Controls(Me, "APTINVH1")

        Create_Summary(grdAPTINVH2, "VOUCHER_LNO", "Count")
        Create_Summary(grdAPTINVH2, "INV_LINE_AMT")

        Create_Summary(grdAPTINVH8, "VOUCHER_ANO", "Count")
        Create_Summary(grdAPTINVH8, "VOUCHER_ADJ_AMT")

        Create_Summary(grdAPTINVH7, "VOUCHER_CLNO", "Count")
        Create_Summary(grdAPTINVH7, New String() {"TOTAL_INV", "TOTAL_ACC", "TOTAL_VAR"})

        Create_Summary(grdAPTACRC1, "CTL_NO", "Count")
        Create_Summary(grdAPTACRC1, New String() {"COST_ACC", "COST_ACT"})

        Create_Summary(grdAPTINVR1, "VOUCHER_NO", "Count")
        Create_Summary(grdAPTINVR1, "INV_AMT")

        Create_Summary(grdAPTINVHB, "VOUCHER_NO", "Count")
        Create_Summary(grdAPTINVHB, "INV_AMT")

        Create_Summary(grdAPTINVHM, "VOUCHER_NO", "Count")
        Create_Summary(grdAPTINVHM, "INV_AMT")

        Create_Summary(grdAPTINVHR, "VOUCHER_NO", "Count")
        Create_Summary(grdAPTINVHR, "INV_AMT")

        Create_Summary(grdICTIREC1, "RECEIPT_NO", "Count", "ICTIREC1")
        Create_Summary(grdICTIREC1, New String() {"QTY_REC", "AMT_REC", "QTY_INV", "AMT_INV", "QTY_REC_NOT_INV", "EXT_PO_COST"}, , "ICTIREC1")

        Create_Summary(grdAPTINVH5_SUM, "RECEIPT_NO", "Count")
        Create_Summary(grdAPTINVH5_SUM, New String() {"QTY_REC", "AMT_REC", "QTY_INV", "AMT_INV", "QTY_REC_NOT_INV", "AMT_REC_NOT_INV", "AMT_REC_NOT_INV_OFFSET", "QTY_VAR", "AMT_VAR"})

        Create_Summary(grdAPTINVH5, "VOUCHER_DLNO", "Count")
        Create_Summary(grdAPTINVH5, New String() {"QTY_REC", "AMT_REC", "INV_QTY", "AMT_INV", "QTY_REC_NOT_INV", "AMT_REC_NOT_INV", "AMT_REC_NOT_INV_OFFSET", "QTY_VAR", "AMT_VAR"})
        Create_Summary(grdAPTINVH5, New String() {"VCOST", "LANDG", "TOOLG", "OVRHD"})

        Create_Summary(grdAPTINVHX, "VOUCHER_NO", "Count")
        Create_Summary(grdAPTINVHX, "SELECTED")
        Create_Summary(grdAPTINVHX, "INV_BALANCE")
        Create_Summary(grdAPTINVHX, "INV_DISC_AMT")
        Create_Summary(grdAPTINVHX, "INV_PAYMENTS")

        Create_Summary(grdAPTSUBM1, "SUBMIT_CTL_NO", "Count")
        '  Create_Summary(grdAPTSUBM1, New String() {"SEL"})

        With grdICTIREC1.DisplayLayout.Bands("ICTIREC1")
            .Columns("RECEIPT_NO").Header.Fixed = True
        End With
        With grdAPTINVH5_SUM.DisplayLayout.Bands("APTINVH5_SUM")
            .Columns("RECEIPT_NO").Header.Fixed = True
        End With
        With grdAPTINVH5.DisplayLayout.Bands("APTINVH5")
            .Columns("VOUCHER_DLNO").Header.Fixed = True
            .Columns("ITEM_CODE").Header.Fixed = True
            If ASCMAIN1.CLIENT = "INT" Then
                grdAPTINVH5.DisplayLayout.Bands(0).Columns("TOOLG").Header.Caption = "Tariff"
            End If
        End With

        With grdAPTINVH5_SUM.DisplayLayout.Bands("APTINVH5_SUM")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
            .Columns("QTY_REC").Header.Appearance.BackColor2 = Drawing.Color.LightPink
            .Columns("AMT_REC").Header.Appearance.BackColor2 = Drawing.Color.LightPink
            .Columns("QTY_REC_NOT_INV").Header.Appearance.BackColor2 = Drawing.Color.LightPink
            .Columns("QTY_INV").Header.Appearance.BackColor2 = Drawing.Color.Yellow
            .Columns("AMT_INV").Header.Appearance.BackColor2 = Drawing.Color.Yellow
            .Columns("QTY_VAR").Header.Appearance.BackColor2 = Drawing.Color.LightSkyBlue
            .Columns("AMT_VAR").Header.Appearance.BackColor2 = Drawing.Color.LightSkyBlue
        End With

        With grdAPTINVH5.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
            .Columns("INV_QTY").Header.Appearance.BackColor2 = Drawing.Color.Yellow
            .Columns("CLOSE_LINE").Header.Appearance.BackColor2 = Drawing.Color.Yellow
            .Columns("INV_COST").Header.Appearance.BackColor2 = Drawing.Color.Yellow
            .Columns("QTY_REC").Header.Appearance.BackColor2 = Drawing.Color.LightPink
            .Columns("PO_COST").Header.Appearance.BackColor2 = Drawing.Color.LightPink
            .Columns("AMT_REC").Header.Appearance.BackColor2 = Drawing.Color.LightPink
            .Columns("QTY_REC_NOT_INV").Header.Appearance.BackColor2 = Drawing.Color.LightPink
            .Columns("AMT_INV").Header.Appearance.BackColor2 = Drawing.Color.Yellow
            .Columns("QTY_VAR").Header.Appearance.BackColor2 = Drawing.Color.LightSkyBlue
            .Columns("AMT_VAR").Header.Appearance.BackColor2 = Drawing.Color.LightSkyBlue
        End With

        grdAPTINVH2.DisplayLayout.Bands("APTINVH2").SummaryFooterCaption = "Voucher Totals"
        grdAPTINVH5_SUM.DisplayLayout.Bands("APTINVH5_SUM").SummaryFooterCaption = "Totals for All Receipts in the Voucher"
        grdICTIREC1.DisplayLayout.Bands("ICTIREC1").SummaryFooterCaption = "Totals for All Receipts Not Vouchered"

        Set_Read_Only(grpAPTVEND1, True)
        Set_Read_Only(grpAPTVEND2, True)
        Set_Read_Only(grpOtherVendorInfo, True)

        Absx1.txtFor("VEND_ALT_CODE").ReadOnly = False Or InquiryMode Or ApprovalMode

        grpShowDistribution.Visible = InquiryMode
        tabReceipts.Tabs("Open Accrued PO Receipts").Visible = Not InquiryMode

        With grdAPTINVR1.DisplayLayout.Bands("APTINVR1")
            For i As Integer = 0 To .Columns.Count - 1
                If .Columns(i).Key = .Columns(i).Header.Caption Then
                    .Columns(i).Hidden = True
                End If
            Next
        End With

        Sort_grdColumns(grdAPTINVH8, "VOUCHER_ANO")
        Sort_grdColumns(grdAPTINVH7, "VOUCHER_CLNO")
        Sort_grdColumns(grdAPTACRC1, "CTL_NO")
        Sort_grdColumns(grdAPTSUBM1, "SUBMIT_CTL_NO".ToLower)



        grdAPTINVHB.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdAPTINVHB.DisplayLayout.Bands(0).SortedColumns.Add("VOUCHER_NO", False)

        grdAPTINVHM.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdAPTINVHM.DisplayLayout.Bands(0).SortedColumns.Add("VOUCHER_NO", False)

        ' THIS SHOULD ALL HAPPEN IN A CALL 
        grdAPTINVR1.DisplayLayout.UseFixedHeaders = True
        With grdAPTINVR1.DisplayLayout.Bands("APTINVR1")
            .Columns("VOUCHER_NO").Header.Fixed = True
            .Columns("VEND_CODE").Header.Fixed = True
            .Columns("INV_TYPE").Header.Fixed = True
            .Columns("INV_NUM").Header.Fixed = True
        End With

        With grdAPTINVH2.DisplayLayout.Bands(0)
            .Columns("VOUCHER_LNO").Header.Fixed = True
        End With

        grdAPTINVHX.DisplayLayout.UseFixedHeaders = True
        With grdAPTINVHX.DisplayLayout.Bands("APTINVHX")
            .Columns("SELECTED").Header.Fixed = True
        End With

        grdAPTINVHB.DisplayLayout.UseFixedHeaders = True
        With grdAPTINVHB.DisplayLayout.Bands("APTINVHB")
            .Columns("VOUCHER_NO").Header.Fixed = True
            .Columns("VEND_CODE").Header.Fixed = True
            .Columns("VEND_NAME").Header.Fixed = True
        End With

        grdAPTINVHM.DisplayLayout.UseFixedHeaders = True
        With grdAPTINVHM.DisplayLayout.Bands("APTINVHM")
            .Columns("VOUCHER_NO").Header.Fixed = True
            .Columns("INV_NUM").Header.Fixed = True
            .Columns("INV_DATE").Header.Fixed = True
            .Columns("INV_AMT").Header.Fixed = True
        End With

        grdAPTSUBM1.DisplayLayout.UseFixedHeaders = True
        With grdAPTSUBM1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                'If gcol.Key = "SEL" Then
                '    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                'Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                'End If
                If gcol.Key = "EMAIL" Then
                    .Columns("EMAIL").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                End If

            Next
            .Override.AllowUpdate = DefaultableBoolean.True
        End With
        grdAPTSUBM1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single


        If ROWs("APTPARM1").Item("AP_PARM_ALLOW_ACCRUAL") & "" <> "1" Then
            chkACCRUE_PRIOR.Visible = False
        End If

        For Each gcol As UltraWinGrid.UltraGridColumn In grdAPTINVH7.DisplayLayout.Bands(0).Columns
            If gcol.Key = "TOTAL_INV" Or gcol.Key = "CTL_NOTE" Then
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                gcol.CellAppearance.BackColor = Drawing.Color.Yellow
            Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next

        If InquiryMode Or ApprovalMode Then
            optINV_STATUS.ValueList.ValueListItems.Add("D", "Deleted")

            With grdAPTINVH2.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowDelete = DefaultableBoolean.False
                .AllowUpdate = DefaultableBoolean.False
            End With

            With grdAPTINVH7.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowDelete = DefaultableBoolean.False
                .AllowUpdate = DefaultableBoolean.False
            End With

            Set_Read_Only(tabMain, True)

            'grdAPTINVHX.Visible = (rowAPTINVH1.Item("INV_STATUS") & "" = "P")
            cmdNextCheckNo.Visible = False
        End If

        'tabMain.Tabs("PO Receipts").Visible = False
        tabMain.Tabs("Other Accruals").Visible = False

        Fill_Records("GLTDIST1")
        If dst.Tables("GLTDIST1").Rows.Count = 0 Then
            grdAPTINVH2.DisplayLayout.Bands(0).Columns("DIST_CODE").Hidden = False
        Else
            'Dim VL As New ValueList
            'For Each rowGLTDIST1 As DataRow In dst.Tables("GLTDIST1").Select("")
            '    Dim VLI As New ValueListItem(rowGLTDIST1.Item("DIST_CODE"))
            '    VL.ValueListItems.Add(VLI)
            'Next
            'grdAPTINVH2.DisplayLayout.Bands(0).Columns("DIST_CODE").ValueList = VL
        End If

        ASCMAIN1.Add_Value_List(grdAPTINVR1, "INV_APPR_STATUS", Nothing, New String() {":", "P:Pending", "A:Approved"})

        ' cmdCheck.Visible = (ASCMAIN1.Running_in_VS)

        chkQuickEntry.Checked = Not InquiryMode And Not ApprovalMode

        If ApprovalMode Then
            ASCMAIN1.sql = "Select * from (Select VEND_BUYER_CODE, VEND_BUYER_NAME from POTBUYR1 union Select '*' VEND_BUYER_CODE, 'All Approvers' VEND_BUYER_NAME from DUAL) order by VEND_BUYER_CODE"
            Dim tbl As DataTable = ASCDATA1.GetDataTable
            cmbVEND_BUYER_CODE.DataSource = tbl
            If tbl.Select("VEND_BUYER_CODE = '" & ASCMAIN1.USER_ID & "'").Length <> 0 Then
                cmbVEND_BUYER_CODE.Value = ASCMAIN1.USER_ID
            End If
            grpApprovalStatus.Top = grpBuyer.Top
            grpApprovalStatus.Left = grpBuyer.Left

            ASCMAIN1.sql = "Select POTBUYR1.VEND_BUYER_CODE, POTBUYR1.VEND_BUYER_NAME from POTBUYR1,ASTUSER1 where ASTUSER1.USER_ID = POTBUYR1.VEND_BUYER_CODE and ASTUSER1.USER_STATUS = 'A' order by VEND_BUYER_CODE"
            cbeAPPR_ROUTE_TO.DataSource = ASCDATA1.GetDataTable

        End If
        If ASCMAIN1.CLIENT = "INT" Then
            grdAPTINVH2.DisplayLayout.Bands(0).Columns("COST_CTR_CODE").Hidden = True
        End If

        'If ASCMAIN1.CLIENT = "INT" Then
        tabMain.Tabs("Other Accruals").Visible = False
        'End If



        dteStart.Value = DateAdd(DateInterval.Day, -365, DateTime.Now)
        dteEnd.Value = DateAdd(DateInterval.Day, +1, DateTime.Now)


        Absx1.dteFor("CTL_DATE").Value = Now.Date
        ASCMAIN1.Add_Value_List(grdAPTACRC1, "CTL_STATUS", Nothing, New String() {":", "0:Open", "1:Closed"})
        ASCMAIN1.Add_Value_List(grdAPTSUBM1, "SUBMIT_STATUS", Nothing, New String() {":", "U:Pending", "P:Processed", "D:Deleted"})
        ASCMAIN1.Add_Value_List(grdAPTSUBM1, "INV_STATUS", Nothing, New String() {":", "O:Open", "P:Paid", "D:Deleted"})

        'UltraTabControl1.Tabs("Submitted Invoices").Visible = (MENU_ITEM_OBJECT = "APFINVH1") And Not InquiryMode

        Set_Read_Only(grpBankingInfo, True)
        Bind_Controls(grpBankingInfo, "APTCHCK5")
        tabHeader.Tabs("Banking").Visible = (ASCMAIN1.USER_SECURITY_CODEs.Contains("P3") Or ASCMAIN1.USER_SECURITY_CODEs.Contains("P4"))

        cmdShowBankingInfo.Visible = Not InquiryMode

        MakeTransparent(lblAddSupplier)

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("VEND_CODE")
                If cdr IsNot Nothing AndAlso cdr.Item("VEND_STATUS") & "" <> "A" Then
                    EMsg &= vbCr & "Vendor is Inactive"
                End If
                If Absx1.txtFor("INV_NUM").Text = "" Then
                    EMsg &= vbCr & "Invoice No Required"
                End If

                If Absx1.cbeFor("INV_TYPE").Text = "" Then
                    EMsg &= vbCr & "Invoice Type Required"
                End If

                '   If .Groups("Submitted Invoices").Visible = True Then
                If UltraTabControl1.ActiveTab.Key = "Submitted Invoices" Then
                    If grdAPTSUBM1.Selected.Rows.Count <> 1 Then
                        EMsg &= vbCr & "Submitted Pending Invoice Line must be selected for new Voucher"
                    Else
                        Dim SUBMIT_CTL_NO As String = grdAPTSUBM1.Selected.Rows(0).Cells("SUBMIT_CTL_NO").Value & ""
                        If Not ASCMAIN1.Logical_Lock("APTSUBM1", SUBMIT_CTL_NO) Then Exit Sub
                    End If

                End If


                If EMsg = "" Then
                    If Check_Invoice() <> "YES" Then
                        Exit Sub
                    End If
                    If Not ASCMAIN1.Logical_Lock("APTVEND1", Absx1.txtFor("VEND_CODE").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Edit", "Load"

                ''**************************************** REMOVE THIS
                'If ASCMAIN1.Running_in_VS Then
                '    Stop
                '    Clear_Out_Accrued_AP()
                '    Exit Sub
                'End If

                If Validate_Code("VOUCHER_NO") Then
                    If Not InquiryMode And Not ApprovalMode Then
                        If cdr.Item("INV_STATUS") & "" = "D" Then
                            EMsg &= vbCr & "Voucher " & Absx1.txtFor("VOUCHER_NO").Text & " has been Deleted"
                        End If
                        If cdr.Item("INV_STATUS") & "" = "P" Then
                            EMsg &= vbCr & "Voucher " & Absx1.txtFor("VOUCHER_NO").Text & " has been Paid"
                        End If
                        If cdr.Item("BATCH_NO_PYMT") & "" <> "" Then
                            EMsg &= vbCr & "Voucher " & Absx1.txtFor("VOUCHER_NO").Text & " has been Selected for Payment in Batch " & cdr.Item("BATCH_NO_PYMT")
                        End If
                    End If
                End If

                If EMsg = "" Then
                    If Not InquiryMode Then
                        If Not ASCMAIN1.Logical_Lock(TABLE_NAME, Absx1.txtFor("VOUCHER_NO").Text) Then Exit Sub
                        If Not ASCMAIN1.Logical_Lock("APTVEND1", cdr.Item("VEND_CODE") & "") Then Exit Sub
                    End If
                End If

                If ApprovalMode Then
                    Dim VOUCHER_NO As String = Absx1.txtFor("VOUCHER_NO").Text
                    If VOUCHER_NO = "" Then
                        EMsg &= vbCr & "No Voucher Selected - cannot load"
                    Else
                        Dim row As DataRow = LookUp("APTINVH1", VOUCHER_NO)
                        If row Is Nothing Then
                            EMsg &= vbCr & "Voucher " & VOUCHER_NO & " is not available any longer - refresh screen"
                        Else
                            If row.Item("INV_APPR_STATUS") & "" <> "P" Then
                                EMsg &= vbCr & "Voucher " & Absx1.txtFor("VOUCHER_NO").Text & " is No Longer Pending Approval"
                                'ASCMAIN1.MultiTask_Release()
                                Show_Batch()
                            ElseIf row.Item("VEND_BUYER_CODE") & "" <> ASCMAIN1.USER_ID Then
                                EMsg &= vbCr & "Voucher " & Absx1.txtFor("VOUCHER_NO").Text & " is not assigned to you for Approval." & vbCr & "It is assigned to " & row.Item("VEND_BUYER_CODE") & vbCr & "See AP to transfer Buyer Approval"
                                'ASCMAIN1.MultiTask_Release()
                                Show_Batch()
                            End If
                        End If
                    End If

                    If EMsg <> "" Then
                        ASCMAIN1.MultiTask_Release()
                    End If
                End If

            Case "Multi-Invoice Edit"
                Validate_Code("VEND_CODE")

                If ASCMAIN1.USER_ID <> "mattinam" And ASCMAIN1.USER_ID <> "wjzz" Then
                    EMsg &= vbCr & "This Option is NOT Ready yet (See Maria/Walter)"
                End If

            Case "Delete"
                If MsgBox("OK to Delete Voucher?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                Else
                    Exit Sub
                End If

            Case "Update"
                If EntryMode = "B" Then
                    ' validate things
                ElseIf EntryMode = "M" Then
                    ' validate things
                ElseIf ApprovalMode Then

                    If optAPPR_ACTION.Value & "" = "" Then
                        EMsg &= vbCr & "You Must Select an Approval Action"
                    Else
                        If optAPPR_ACTION.Value = "P" Or optAPPR_ACTION.Value = "R" Then
                            If cbeAPPR_ROUTE_TO.Value & "" = "" Then
                                EMsg &= vbCr & "You Must Select an Approver to Route To"
                            Else
                                If LookUp("POTBUYR1", cbeAPPR_ROUTE_TO.Value) Is Nothing Then
                                    EMsg &= vbCr & "Invalid Value for Approver to Route To"
                                End If
                            End If
                        Else
                            Dim row As DataRow = LookUp("POTBUYR1", ASCMAIN1.USER_ID)
                            If row Is Nothing Then
                                EMsg &= vbCr & "You are not Authorized to Approve"
                            Else
                                Dim INV_AMT As Decimal = Val(rowAPTINVH1.Item("INV_AMT") & "")
                                Dim VEND_BUYER_APPR_LIMIT As Decimal = Val(row.Item("VEND_BUYER_APPR_LIMIT") & "")
                                If INV_AMT > VEND_BUYER_APPR_LIMIT Then
                                    EMsg &= vbCr & "You are not Authorized to Finally Approve over " & Format(VEND_BUYER_APPR_LIMIT, "$#,##0")
                                End If
                            End If
                        End If
                    End If
                Else

                    If Absx1.txtFor("INV_NUM").Text = "" Then
                        EMsg &= vbCr & "You Must Enter an Invoice No"
                    End If
                    If Absx1.dteFor("INV_DATE").Value & "" = "" Then
                        EMsg &= vbCr & "You Must Enter a Document Date"
                    End If
                    If Absx1.dteFor("INV_DUE_DATE").Value & "" = "" Then
                        EMsg &= vbCr & "You Must Enter a Valid Terms Code and Document Date so that a Due Date may be calculated"
                    End If

                    Dim DIST_OOBAL As Decimal = Val(Absx1.numFor("DIST_OOBAL").Value & "")
                    If System.Math.Round(DIST_OOBAL, 2) <> 0 Then
                        EMsg &= vbCr & "Distribution is Out of Balance"
                    End If

                    If ROWs("APTPARM1").Item("AP_PARM_APPR_REQD") & "" = "1" Then
                        If rowAPTVEND1.Item("VEND_AUTO_APPROVE") & "" = "1" Then
                        Else
                            If Val(rowAPTINVH1.Item("INV_AMT") & "") > Val(ROWs("APTPARM1").Item("AP_PARM_APPR_LIMIT") & "") Then
                                Dim VEND_BUYER_CODE As String = rowAPTINVH1.Item("VEND_BUYER_CODE") & ""
                                If VEND_BUYER_CODE = "" Then
                                    EMsg &= vbCr & "Buyer is necessary to handle Approvals"
                                Else
                                    Dim row As DataRow = LookUp("POTBUYR1", VEND_BUYER_CODE)
                                    If row Is Nothing Then
                                        EMsg &= vbCr & "Invalid Value specified for Buyer"
                                    Else
                                        Dim rowASTUSER1 As DataRow = LookUp("ASTUSER1", VEND_BUYER_CODE)
                                        If rowASTUSER1 Is Nothing Then
                                            EMsg &= vbCr & "Invalid Value specified for Buyer"
                                        Else
                                            If rowASTUSER1.Item("USER_STATUS") & "" <> "A" Then
                                                EMsg &= vbCr & "Invalid Value specified for Buyer"
                                            End If
                                        End If
                                    End If
                                End If
                                If rowAPTINVH1("INV_STATUS") = "P" Then
                                    If rowAPTINVH1.Item("INV_APPR_STATUS") & "" <> "A" Then
                                        EMsg &= vbCr & "Approval is required before Paying"
                                    End If
                                End If
                            End If
                        End If

                        For Each row As DataRow In dst.Tables("APTINVHX").Select("SELECTED = '1'")
                            Dim VOUCHER_NO As String = row.Item("VOUCHER_NO")
                            Dim rowAPTINVH1_other As DataRow = LookUp("APTINVH1", VOUCHER_NO)
                            If rowAPTINVH1_other.Item("INV_APPR_STATUS") & "" <> "A" Then
                                EMsg &= vbCr & "Approval is required before Paying (Voucher " & VOUCHER_NO & ")"
                            End If
                        Next
                    End If

                    If Not TAC.APCMAIN1.check_Bank_Payment_Method(Me, Absx1.txtFor("BANK_CODE").Text, Absx1.txtFor("INV_PYMT_METHOD").Text) Then
                        EMsg &= vbCr & $"Bank {Absx1.txtFor("BANK_CODE").Text} does NOT support Payment Method { Absx1.txtFor("INV_PYMT_METHOD").Text}"
                    End If

                    Select Case Absx1.cbeFor("INV_TYPE").Value
                        Case "I", "D"
                            Dim INV_AMT As Decimal = Val(Absx1.numFor("INV_AMT").Value & "")
                            Dim INV_AMT_VEND As Decimal = Val(Absx1.numFor("INV_AMT_VEND").Value & "")
                            If Val(Absx1.numFor("INV_AMT").Value & "") < 0 _
                            Or Val(Absx1.numFor("INV_AMT_VEND").Value & "") < 0 Then
                                EMsg &= vbCr & "Amount must be Postive (i.e., an amount owed TO the Vendor) for this type of Document"
                            End If
                        Case "C", "R"
                            If Val(Absx1.numFor("INV_AMT").Value & "") > 0 _
                            Or Val(Absx1.numFor("INV_AMT_VEND").Value & "") > 0 Then
                                EMsg &= vbCr & "Amount must be Negative (i.e., an amount owed FROM the Vendor) for this type of Document"
                            End If
                        Case "A"
                            If EntryMode = "N" Then
                                If Val(Absx1.numFor("INV_AMT").Value & "") < 0 _
                                Or Val(Absx1.numFor("INV_AMT_VEND").Value & "") < 0 Then
                                    EMsg &= vbCr & "Amount must be Postive (i.e., an amount owed TO the Vendor) for this type of Document"
                                    EMsg &= vbCr & "- An offsetting document with the same (but Negative) Amount will be automatically created"
                                End If
                            End If
                    End Select

                    If Absx1.optFor("INV_REMIT_TO").Value = "N" Then
                        If Absx1.txtFor("VEND_ALT_CODE").Text = "" _
                        Or Absx1.txtFor("VEND_ALT_NAME").Text = "" _
                        Or Absx1.txtFor("VEND_ALT_CITY").Text = "" _
                        Or Absx1.txtFor("VEND_ALT_STATE").Text = "" _
                        Or Absx1.txtFor("VEND_ALT_ZIP_CODE").Text = "" Then
                            EMsg &= vbCr & "Address Code, Vendor Name, City, State & Zip are Required for a New Payment Address"
                        Else
                            LookUp("APTVEND2", New String() {HFs("VEND_CODE"), Absx1.txtFor("VEND_ALT_CODE").Text})
                            If cdr IsNot Nothing Then
                                EMsg &= vbCr & "Vendor Address Code " & Absx1.txtFor("VEND_ALT_CODE").Text & " is Already on File"
                            End If
                        End If
                    Else
                        If Absx1.optFor("INV_REMIT_TO").Value = "V" _
                        Or Absx1.optFor("INV_REMIT_TO").Value = "P" Then
                            If Absx1.txtFor("VEND_ALT_CODE").Text <> "" Then
                                Absx1.txtFor("VEND_ALT_CODE").Text = ""
                            End If
                        End If
                        If Absx1.optFor("INV_REMIT_TO").Value = "A" Then
                            If Absx1.txtFor("VEND_ALT_CODE").Text = "" Then
                                EMsg &= vbCr & "You Must Specify a Valid (Alternate) Payment Address Code"
                            Else
                                LookUp("APTVEND2", New String() {HFs("VEND_CODE"), Absx1.txtFor("VEND_ALT_CODE").Text})
                                If cdr Is Nothing Then
                                    EMsg &= vbCr & "Vendor Address Code " & Absx1.txtFor("VEND_ALT_CODE").Text & " is not on File"
                                End If
                            End If
                        End If
                    End If

                    If ROWs("APTPARM1").Item("AP_PARM_BANK_METHOD") & "" = "1" Then


                        If Absx1.txtFor("INV_PYMT_METHOD").Text = "ACH" Or Absx1.txtFor("INV_PYMT_METHOD").Text = "WIRE" Then
                            Dim colpfx As String = "VEND_"

                            Dim rowP As DataRow = Nothing
                            If Absx1.optFor("INV_REMIT_TO").Value = "V" Then
                                rowP = LookUp("APTVEND1", HFs("VEND_CODE"))
                            ElseIf Absx1.optFor("INV_REMIT_TO").Value = "P" Then
                                rowP = LookUp("APTVEND1", Absx1.txtFor("VEND_CODE_AP").Text)
                            ElseIf Absx1.optFor("INV_REMIT_TO").Value = "A" Then
                                rowP = LookUp("APTVEND2", New String() {HFs("VEND_CODE"), Absx1.txtFor("VEND_ALT_CODE").Text})
                                colpfx = "VEND_ALT_"
                            End If

                            Dim PYMT_METHOD As String = Absx1.txtFor("INV_PYMT_METHOD").Text
                            If rowP Is Nothing Then
                                EMsg &= vbCr & "Cannot Validate Payment Address for ACH"
                            Else
                                If rowP.Item($"{colpfx}BANK_ACCT_ID") & "" = "" Then
                                    EMsg &= vbCr & $"Cannot Validate Payment Address for {PYMT_METHOD} - Bank Account ID not Specified"
                                End If
                                If rowP.Item($"{colpfx}BANK_COUNTRY") & "" = "USA" Then
                                    If rowP.Item($"{colpfx}BANK_ROUTING_NO") & "" = "" Then
                                        EMsg &= vbCr & $"Cannot Validate Payment Address for {PYMT_METHOD} - Bank Account Routing No not Specified"
                                    End If
                                Else
                                    If rowP.Item($"{colpfx}BANK_SWIFT_NO") & "" = "" Then
                                        EMsg &= vbCr & $"Cannot Validate Payment Address for {PYMT_METHOD} - Bank Account Swift No not Specified"
                                    End If
                                    If PYMT_METHOD <> "WIRE" Then
                                        EMsg &= vbCr & $"Cannot use {PYMT_METHOD} as a Payment Method for Non-US Banks"
                                    End If
                                End If
                                If rowP.Item($"{colpfx}BANK_ACCT_CLASS") & "" = "" Then
                                    EMsg &= vbCr & $"Cannot Validate Payment Address for {PYMT_METHOD} - Bank Account Class not Specified"
                                End If
                                If rowP.Item($"{colpfx}BANK_ACCT_TYPE") & "" = "" Then
                                    EMsg &= vbCr & $"Cannot Validate Payment Address for {PYMT_METHOD} - Bank Account Type not Specified"
                                End If
                            End If
                        End If

                    End If

                    If rowAPTINVH1("INV_STATUS") = "R" Then
                        If Absx1.cbeFor("INV_TYPE").Value <> "I" Then
                            EMsg &= vbCr & "Recurring Feature applies to Invoices Only"
                        End If
                        If Val(Absx1.numFor("INV_AMT").Value & "") <= 0 Then
                            EMsg &= vbCr & "Recurring Invoice Template must have a positive, non-zero Amount"
                        Else
                            If Val(Absx1.numFor("INV_AMT_VEND").Value & "") <= 0 Then
                                EMsg &= vbCr & "Recurring Invoice Template must have a positive, non-zero Amount"
                            End If
                        End If
                        If Absx1.optFor("INV_RECUR_CYCLE").Value & "" = "" Then
                            EMsg &= vbCr & "Recurring Invoice Template must have an Recurring Cycle"
                        End If

                        If dst.Tables("APTINVH5").Rows.Count <> 0 Then
                            EMsg &= vbCr & "Recurring Invoice Templates may not have entries for Accrued Purchases"
                        End If
                        If dst.Tables("APTINVH8").Rows.Count <> 0 Then
                            EMsg &= vbCr & "Recurring Invoice Templates may not have entries for Adjustments"
                        End If
                        If dst.Tables("APTINVH7").Rows.Count <> 0 Then
                            EMsg &= vbCr & "Recurring Invoice Templates may not have entries for Other Accruals"
                        End If

                        If Absx1.txtFor("INV_RECUR_OPS_YYYYPP_BEGIN").Text = "" Then
                            EMsg &= vbCr & "Recurring Invoice Template must have a Starting Period"
                        Else
                            If Not Validate_Code("INV_RECUR_OPS_YYYYPP_BEGIN") Then
                                EMsg &= vbCr & "Invalid Starting Period"
                            End If
                        End If
                    End If

                    If rowAPTINVH1("INV_STATUS") = "P" Then
                        If Absx1.txtFor("BANK_CODE").Text = "" Then
                            EMsg &= vbCr & "You Must Specify a Bank Code to Pay upon Entry"
                        End If
                        If Absx1.txtFor("INV_PYMT_METHOD").Text = "" Then
                            EMsg &= vbCr & "You Must Specify a Payment Method to Pay upon Entry"
                        Else
                            If ROWs("APTPARM1").Item("AP_PARM_BANK_METHOD") & "" = "1" Then
                                Dim INV_PYMT_METHOD As String = Absx1.txtFor("INV_PYMT_METHOD").Text
                                If INV_PYMT_METHOD = "ACH" Or INV_PYMT_METHOD = "WIRE" Then
                                    EMsg &= vbCr & $"Invalid Payment Method ({INV_PYMT_METHOD}) for an invoice that is Paid upon Entry"
                                    EMsg &= vbCr & " - use CHECK or MANUAL"
                                End If
                            End If
                        End If
                        If Absx1.txtFor("CHECK_NUM").Text = "" Or Absx1.dteFor("CHECK_DATE").Value & "" = "" Then
                            If batch_update Then
                                ' generate the check number
                            End If
                            EMsg &= vbCr & "You Must Specify a Check Number and Date to Pay upon Entry"
                        End If
                        If EMsg = "" Then
                            Dim rowAPTCHCK1 As DataRow = LookUp("APTCHCK1", New String() {Absx1.txtFor("BANK_CODE").Text, Absx1.txtFor("CHECK_NUM").Text})
                            If rowAPTCHCK1 IsNot Nothing Then
                                EMsg &= vbCr & "Check No " & Absx1.txtFor("CHECK_NUM").Text & " has already been Posted"
                            End If
                        End If
                        'If Val(Absx1.numFor("CHECK_AMT").Value) = 0 And (Absx1.txtFor("BANK_CODE").Text <> "Z") Then
                        '    EMsg &= vbCr & "You Must Use Bank Code Z for Zero Checks"
                        'End If
                        'If Val(Absx1.numFor("CHECK_AMT").Value) <> 0 And (Absx1.txtFor("BANK_CODE").Text = "Z") Then
                        '    EMsg &= vbCr & "You May NOT Use Bank Code Z for Non-Zero Checks"
                        'End If
                    Else
                        If EMsg = "" Then
                            Absx1.txtFor("CHECK_NUM").Text = ""
                            Absx1.dteFor("CHECK_DATE").Value = ""
                            If Val(Absx1.numFor("INV_AMT").Value & "") = 0 Then
                                If vbNo = MsgBox("You have not clicked 'Paid'," &
                                                 vbCr & "  which means this Invoice will be updated as 'Open'" &
                                                 vbCr & "  and will need to be Selected for Payment" &
                                                 vbCr & "  to remove it from the Open AP Items Report" _
                                                 & vbCr & vbCr & "Continue Anyway?", vbQuestion + vbYesNo,
                                                 "Normally, a $0 Invoice is entered as 'Paid' on a $0 Check") Then
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If

                    Validate_Code("POST_CODE")
                    Validate_Code("TERM_CODE")
                    Validate_Code("INV_PYMT_METHOD")
                    Validate_Code("BANK_CODE", , True)
                    Validate_Code("INV_PYMT_CYCLE", , True)
                    Validate_Code("CURR_CODE")
                    Validate_Code("REASON_CODE", , True) ' TRUE/FALSE SHOULD BE RELATED TO MEMO

                    For Each row As DataRow In ASCDATA1.SelectDistinct _
                        (dst.Tables("APTINVH2").Select("INV_LTYP is Null and DIST_CODE is Not Null"),
                         New String() {"DIST_CODE"}).Select("")
                        Dim DIST_CODE As String = row.Item("DIST_CODE")
                        If LookUp("GLTDIST1", DIST_CODE) Is Nothing Then
                            EMsg &= vbCr & "Invalid Distribution Code " & DIST_CODE
                        End If
                    Next


                    EMsg &= Validate_Accounts_and_Segments_EMsg(dst.Tables("APTINVH2"), False)

                    'For Each row As DataRow In dst.Tables("APTINVH2").Select("")
                    '    Dim ACCT_CODE As String = row.Item("ACCT_CODE") & ""
                    '    If LookUp("GLTACCT1", ACCT_CODE) Is Nothing Then
                    '        EMsg &= vbCr & "Invalid Account Code " & ACCT_CODE
                    '    Else
                    '        If cdr.Item("ACCT_STATUS") & "" <> "A" Then
                    '            EMsg &= vbCr & "Acct Code " & ACCT_CODE & " is not Active"
                    '        End If
                    '        If cdr.Item("ACCT_SUB_CTL") & "" = "1" Then
                    '            EMsg &= vbCr & "Acct Code " & ACCT_CODE & " is a Control Account - no Manual J/E permitted"
                    '        End If
                    '    End If
                    'Next

                    If EMsg = "" Then
                        If Val(Absx1.numFor("INV_AMT_VEND").Value & "") <> Val(Absx1.numFor("INV_AMT").Value & "") Then
                            If MsgBox("Please Verify the Following Information:" & vbCr & vbCr & "Vendor Invoice Amount: " & Format(Val(Absx1.numFor("INV_AMT_VEND").Value & ""), "#,##0.00") & vbCr & "Invoice Payable: " & Format(Val(Absx1.numFor("INV_AMT").Value & ""), "#,##0.00") & vbCr & vbCr & "OK To Continue with Update?", vbQuestion + vbYesNo, "Verification: Invoice will be Booked with Adjustments") = vbNo Then
                                Exit Sub
                            End If
                        End If
                    End If

                    If EMsg = "" Then
                        If chkINV_1099_IND.Checked And Val(Absx1.numFor("INV_1099_AMT").Value & "") = 0 Then
                            If MsgBox("1099 Amount Option is checked, but there is no 1099 Amount Entered." _
                                       & vbCrLf & vbCrLf & "Proceed with Update?",
                                       MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub
                        End If
                    End If




                    If EMsg = "" And rowAPTINVH1("INV_STATUS") = "P" Then
                        Dim OTHERS As Decimal = Val(dst.Tables("APTINVHX").Compute("SUM(INV_PAYMENTS)", "SELECTED = '1'") & "")
                        Dim CHECK_AMT As Decimal = Val(Absx1.numFor("CHECK_AMT").Value & "")
                        Dim INV_PAYMENTS As Decimal = Val(Absx1.numFor("INV_AMT").Value & "") - Val(Absx1.numFor("INV_DISC_AMT").Value & "")
                        If CHECK_AMT <> OTHERS + INV_PAYMENTS Then
                            EMsg = EMsg & vbCr & "The payment amount for this Invoice (" & Format(INV_PAYMENTS, "$##,##0.00") & ")" & vbCr & " plus Selected Other AP Items (" & Format(OTHERS, "$##,##0.00") & ")" & vbCr & " does not agree with Check Amount Specified (" & Format(CHECK_AMT, "$##,##0.00") & ")"
                        End If
                        If CHECK_AMT < 0 Then
                            EMsg = EMsg & vbCr & "This Invoice plus Selected Other AP Items has a Net Negative Balance" & vbCr & " - Negative Payment Not Permitted"
                        End If

                        If EMsg = "" And Not batch_update Then
                            ' PAYEENAME IS WRONG

                            '**************************************** REMOVE THIS
                            'If ASCMAIN1.Running_in_VS And Absx1.txtFor("INV_NUM").Text = "2051 CLEANUP 201512" Then
                            'Else
                            If MsgBox("Please Verify the Following Information:" & vbCr & vbCr & "Check No: " & Absx1.txtFor("CHECK_NUM").Text & ", " & Absx1.dteFor("CHECK_DATE").Value & vbCr & "Bank: " & Absx1.txtFor("BANK_DESC").Text & vbCr & "Payee: " & Absx1.txtFor("VEND_NAME").Text & vbCr & "Amount: " & Format(Val(Absx1.numFor("CHECK_AMT").Value & ""), "$###,##0.00") & IIf(Val(Absx1.numFor("INV_AMT").Value & "") = 0, vbCr & vbCr & "*** This Invoice has a $0 Balance ***", "") & vbCr & vbCr & "OK To Continue with Update?", vbQuestion + vbYesNo, "Verification: You are about to Record a Payment") = vbNo Then
                                Exit Sub
                            End If
                            'End If

                        End If
                    End If

                    'If EntryMode = "E" Then
                    '    If Absx1.optFor("INV_STATUS").Value = "R" _
                    '    And rowAPTINVH1("INV_STATUS", DataRowVersion.Original) <> "R" Then
                    '        EMsg &= vbCr & "Cannot Change an invoice into a Recurring Invoice Template"
                    '    End If
                    '    If Absx1.optFor("INV_STATUS").Value <> "R" _
                    '    And rowAPTINVH1("INV_STATUS", DataRowVersion.Original) = "R" Then
                    '        EMsg &= vbCr & "Cannot Use a Recurring Invoice Template for an Actual Invoice Posting"
                    '    End If
                    'End If

                    Dim RNI As New List(Of String)

                    For Each rowAPTINVH5_SUM As DataRow In dst.Tables("APTINVH5_SUM").Select("")
                        Dim RECEIPT_NO As String = rowAPTINVH5_SUM.Item("RECEIPT_NO") & ""
                        Dim QTY_REC As Int32 = Val(rowAPTINVH5_SUM.Item("QTY_REC") & "")
                        Dim QTY_INV As Int32 = Val(rowAPTINVH5_SUM.Item("QTY_INV") & "")
                        If QTY_INV = 0 Then
                            ' DO NOTHING
                        Else
                            'QTY_INV is for this entry, need to get Total QTY_INV from APTINVH5 for this receipt to re-calc R^I
                            Dim QTY_INV_TOTAL = Val(dst.Tables("APTINVH5").Compute("SUM(QTY_INV)", $"RECEIPT_NO = '{RECEIPT_NO}'") & "")
                            Dim QTY_REC_NOT_INV_net As Int32 = QTY_REC - QTY_INV_TOTAL - QTY_INV
                            If QTY_REC > 0 And QTY_INV > 0 And QTY_REC_NOT_INV_net < 0 Then
                                Dim RNI_item As String = $"Receipt {RECEIPT_NO}"
                                RNI.Add(RNI_item)
                                If RNI.Count > 3 Then
                                    If RNI.Count = 4 Then
                                        ' EMsg &= vbCr & "... and others"
                                        RNI.Add("... and others")
                                    End If
                                Else
                                    ' EMsg &= vbCr & $"Rec^Inv will be Negative (See {RNI_item})"
                                End If
                            End If
                        End If
                    Next


                    If RNI.Count > 0 Then
                        If ASCMAIN1.Running_in_VS Then
                            Stop

                            Dim msg As String = "Rec^Inv will be Negative" & vbCrLf & vbCrLf & "OK to Update anyway?" & vbCrLf & vbCrLf & Join(RNI.ToArray, vbCrLf)
                            If MsgBox(msg, MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                EMsg &= vbCr & "Review Qty Invoiced with respect to Qty R^I"
                            End If

                        End If

                    End If


                    Dim COST_CATGY_CODEs As New List(Of String)
                    Dim sqlw As String = "ISNULL(INV_COST,0) < 0 OR (ISNULL(QTY_REC,0)>0 AND ISNULL(INV_QTY,0)<0) OR (ISNULL(QTY_REC,0)<0 AND ISNULL(INV_QTY,0)>0) or (ISNULL(AMT_VAR,0) = 0 AND ISNULL(CB,'0') = '1') or (COST_CATGY_CODE is Null or ACCT_CODE_PPV is Null) "
                    For Each rowAPTINVH5 As DataRow In dst.Tables("APTINVH5").Select("")
                        Dim QTY_REC As Int32 = Val(rowAPTINVH5.Item("QTY_REC") & "")
                        Dim QTY_INV As Int32 = Val(rowAPTINVH5.Item("QTY_INV") & "")
                        Dim INV_QTY As Int32 = Val(rowAPTINVH5.Item("INV_QTY") & "")
                        Dim QTY_REC_NOT_INV As Int32 = QTY_REC - QTY_INV
                        ' If (QTY_REC > 0 And INV_QTY < 0) Or (QTY_REC < 0 And INV_QTY > 0) Then
                        If (QTY_REC_NOT_INV > 0 And INV_QTY < 0) Or (QTY_REC_NOT_INV < 0 And INV_QTY > 0) Then
                            'EMsg &= vbCr & "Invoice Qty cannot be opposite in sign to Receipt Qty (See Receipt " & rowAPTINVH5.Item("RECEIPT_NO") & " Line " & rowAPTINVH5.Item("RECEIPT_LNO") & ")"
                            ' 09/03/19 - entry of invoice permits over paying, but calling up invoice to make manual payment gets caught here.
                            ' disabling for now
                            ' EMsg &= vbCr & "Invoice Qty cannot be opposite in sign to Qty Received NOT Invoiced (See Receipt " & rowAPTINVH5.Item("RECEIPT_NO") & " Line " & rowAPTINVH5.Item("RECEIPT_LNO") & ")"
                        End If

                        'If QTY_REC > 0 And QTY_INV > 0 And QTY_REC_NOT_INV < 0 Then
                        '    Dim RNI_item As String = "Receipt " & rowAPTINVH5.Item("RECEIPT_NO") & " Line " & rowAPTINVH5.Item("RECEIPT_LNO")
                        '    RNI.Add(RNI_item)
                        '    If RNI.Count > 3 Then
                        '        If RNI.Count = 4 Then
                        '            ' EMsg &= vbCr & "... and others"
                        '            RNI.Add("... and others")
                        '        End If
                        '    Else
                        '        ' EMsg &= vbCr & $"Rec^Inv will be Negative (See {RNI_item})"
                        '    End If

                        'End If

                        Dim INV_COST As Decimal = Val(rowAPTINVH5.Item("INV_COST") & "")
                        If INV_COST < 0 Then
                            EMsg &= vbCr & "Invoice Cost cannot be Negative (See Receipt " & rowAPTINVH5.Item("RECEIPT_NO") & " Line " & rowAPTINVH5.Item("RECEIPT_LNO") & ")"
                        End If

                        Dim AMT_VAR As Decimal = Val(rowAPTINVH5.Item("AMT_VAR") & "")
                        Dim CB As Decimal = Val(rowAPTINVH5.Item("CB") & "")
                        If AMT_VAR = 0 And CB = "1" Then
                            EMsg &= vbCr & "Cannot Chargeback If Variance Is 0 (See Receipt " & rowAPTINVH5.Item("RECEIPT_NO") & " Line " & rowAPTINVH5.Item("RECEIPT_LNO") & ")"
                        End If

                        Dim COST_CATGY_CODE As String = rowAPTINVH5.Item("COST_CATGY_CODE") & ""
                        Dim ACCT_CODE_PPV As String = rowAPTINVH5.Item("ACCT_CODE_PPV") & ""
                        If COST_CATGY_CODE = "" Or ACCT_CODE_PPV = "" Then
                            If Not COST_CATGY_CODEs.Contains(COST_CATGY_CODE) Then
                                EMsg &= vbCr & "Price Variance Account For Cost Category " & COST_CATGY_CODE & " Is Not Set up (See Receipt " & rowAPTINVH5.Item("RECEIPT_NO") & " Line " & rowAPTINVH5.Item("RECEIPT_LNO") & ")"
                                COST_CATGY_CODEs.Add(COST_CATGY_CODE)
                            End If
                        End If

                    Next

                    'If RNI.Count > 0 Then
                    '    Dim msg As String = "Rec^Inv will be Negative" & vbCrLf & vbCrLf & "OK to Update anyway?" & vbCrLf & vbCrLf & Join(RNI.ToArray, vbCrLf)
                    '    If MsgBox(msg, MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    '        EMsg &= vbCr & "Review Qty Invoiced with respect to Qty R^I"
                    '    End If
                    'End If


                    If chkACCRUE_PRIOR.Checked Then
                        'If Absx1.cmbFor("OPS_YYYYPP_ACCRUE").Value & "" = "" Then
                        If Absx1.txtFor("OPS_YYYYPP_ACCRUE").Text & "" = "" Then
                            EMsg &= vbCr & "Missing Accrual Period"
                        End If
                    End If

                    If EMsg = "" Then

                        '**************************************** REMOVE THIS
                        'If ASCMAIN1.Running_in_VS And Absx1.txtFor("INV_NUM").Text = "2051 CLEANUP 201512" Then
                        'Else
                        If Val(Absx1.numFor("INV_AMT").Value & "") = 0 Then
                            If MsgBox("Proceed With Entry?", MsgBoxStyle.YesNo, "Invoice Payable Amount Is Zero") <> MsgBoxResult.Yes Then
                                Exit Sub
                            End If
                        End If
                        'End If

                    End If
                End If

        End Select

        If EMsg <> "" Then
            If Not batch_update Then
                MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            End If
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

                lblSUBMITTED_INVOICE1.Visible = False
                lblSUBMITTED_INVOICE2.Visible = False
                lblSUBMITTED_INVOICE3.Visible = False
                INVOICE_FROM_EMAIL = ""

                If UltraTabControl1.ActiveTab.Key = "Submitted Invoices" Then

                    Dim grow As UltraWinGrid.UltraGridRow = grdAPTSUBM1.Selected.Rows(0)

                    lblSUBMITTED_INVOICE1.Text = "From: " & grow.Cells("SUBMIT_EMAIL_FROM").Value
                    lblSUBMITTED_INVOICE1.Visible = True
                    lblSUBMITTED_INVOICE2.Text = "Subj: " & grow.Cells("SUBMIT_SUBJECT").Value
                    lblSUBMITTED_INVOICE2.Visible = True
                    lblSUBMITTED_INVOICE3.Text = "Submitted:  " & grow.Cells("SUBMIT_DATE_RECEIVED").Value
                    lblSUBMITTED_INVOICE3.Visible = True

                    INVOICE_FROM_EMAIL = grow.Cells("SUBMIT_CTL_NO").Value
                    ASCMAIN1.sql = sqlAPTSUBM1 & " and SUBMIT_CTL_NO = '" & INVOICE_FROM_EMAIL & "'" & vbCrLf
                    Fill_Records("APTSUBM1", "", True, ASCMAIN1.sql)
                End If


                If chkQuickEntry.Checked Then
                    Absx1.txtFor("INV_REF").Text = Absx1.txtFor("QE_INV_REF").Text
                    Absx1.dteFor("INV_DATE").Value = Absx1.dteFor("QE_INV_DATE").Value
                    Absx1.numFor("INV_AMT_VEND").Value = Absx1.numFor("QE_INV_AMT").Value
                    'Absx1.numFor("INV_AMT").Value = Absx1.numFor("QE_INV_AMT").Value

                    Calculate_INV_DUE_DATE()
                    If dst.Tables("APTINVH2").Rows.Count = 0 And Val(numINV_AMT.Value & "") <> 0 Then
                        Generate_Pre_Distribution()
                        Calc_Totals()
                    End If
                    tabMain.SelectedTab = tabMain.Tabs(2)

                End If

            Case "New Batch"
                EntryMode = "B"
                Prepare_for_Batch_Entry()
                Mode_Settings(True)

            Case "Multi-Invoice Edit"
                EntryMode = "M"
                Prepare_for_Multi_Invoice_Edit()
                Mode_Settings(True)

            Case "Edit", "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                If EntryMode = "B" Then
                    Update_Batch()
                ElseIf EntryMode = "M" Then
                    Update_Multi()
                ElseIf ApprovalMode Then
                    Update_Approval()
                Else
                    Update_Record()
                End If
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Print Edit"
                Print_Record()

            Case "Export to Excel"
                If EntryMode = "B" Then
                    Export_to_Excel(grdAPTINVHB)
                ElseIf EntryMode = "M" Then
                    Export_to_Excel(grdAPTINVHM)
                End If

            Case "Import from Excel"
                Import_Batch_from_Excel()

            Case "All Vouchers"
                Copy_Change_to("All")

            Case "Selected Vouchers"
                Copy_Change_to("Sel")
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")

                    If InquiryMode Then
                        .Items("Load").Settings.Enabled = not_iScreenMode
                        .Items("Done").Settings.Enabled = iScreenMode
                    ElseIf ApprovalMode Then
                        .Items("Load").Settings.Enabled = not_iScreenMode
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                    Else
                        .Items("New").Settings.Enabled = not_iScreenMode
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode

                        If EntryMode = "B" Then
                            .Items("Delete").Settings.Enabled = DefaultableBoolean.False
                        ElseIf EntryMode = "M" Then
                            .Items("Delete").Settings.Enabled = DefaultableBoolean.False
                        Else
                            .Items("Delete").Settings.Enabled = iScreenMode
                        End If

                        If EntryMode = "N" Then
                            .Items("Delete").Visible = False
                        Else
                            .Items("Delete").Visible = Not InquiryMode And Not ApprovalMode
                        End If

                        .Items("Print Edit").Settings.Enabled = not_iScreenMode
                        .Items("New Batch").Settings.Enabled = not_iScreenMode
                        .Items("Multi-Invoice Edit").Settings.Enabled = not_iScreenMode
                    End If

                    'WJZ DEMO
                    .Items("Print Edit").Visible = False
                    .Items("New Batch").Visible = False
                    .Items("Multi-Invoice Edit").Visible = False
                End With

                If InquiryMode Or ApprovalMode Then
                    .Groups("Distribution Options").Visible = False
                    .Groups("Entry Options").Visible = False
                    .Groups("Copy Last Change to ...").Visible = False
                    .Groups("Batch / Excel Options").Visible = False
                Else
                    .Groups("Copy Last Change to ...").Visible = False
                    .Groups("Distribution Options").Visible = False
                    .Groups("Entry Options").Visible = Not tf

                    If EntryMode = "B" Then
                        .Groups("Batch / Excel Options").Visible = tf
                    ElseIf EntryMode = "M" Then
                        .Groups("Copy Last Change to ...").Visible = tf
                    Else
                        .Groups("Batch / Excel Options").Visible = False
                    End If
                End If

                .Groups("Generate Accrual").Visible = False ' ApprovalMode

                .Groups("Approval Options").Visible = ApprovalMode
                If ApprovalMode Then
                    grpApprovalStatus.Visible = ScreenMode
                    grpBuyer.Visible = Not ScreenMode
                End If
            End With
        End If

        tabMain.Tabs("PO Receipts").Visible = ScreenMode AndAlso (dst.Tables("ICTIREC1").Rows.Count <> 0)

        If ASCMAIN1.CLIENT = "INT" Then
            cmbVEND_BUYER_CODE.ReadOnly = True
        End If

        'WJZ DEMO
        chkQuickEntry.Checked = False
        chkQuickEntry.Visible = False
        chkRecurring.Visible = False

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grpLastChange.Visible = False

        grpProRate.Visible = Not InquiryMode

        lblDIST_OOBAL.Visible = Not InquiryMode
        numDIST_OOBAL.Visible = Not InquiryMode

        If InquiryMode Or ApprovalMode Then
            Absx1.txtFor("INV_NUM").ReadOnly = True
            Absx1.txtFor("VEND_ALT_CODE").ButtonsRight(0).Enabled = False
            Absx1.txtFor("VEND_ALT_CODE").ReadOnly = True

        ElseIf EntryMode = "Z" Then
            Absx1.txtFor("INV_NUM").ReadOnly = True
            Absx1.txtFor("VEND_ALT_CODE").ButtonsRight(0).Enabled = False
            Absx1.txtFor("VEND_ALT_CODE").Enabled = False
        Else
            Absx1.txtFor("INV_NUM").ReadOnly = False
            Absx1.txtFor("VEND_ALT_CODE").ButtonsRight(0).Enabled = True
            Absx1.txtFor("VEND_ALT_CODE").Enabled = True
        End If

        grdAPTINVR1.Visible = Not ScreenMode
        UltraTabControl1.Visible = Not ScreenMode

        tabMain.Visible = tf
        grdAPTINVHB.Visible = False
        grdAPTINVHM.Visible = False

        If ScreenMode And Not InquiryMode And Not ApprovalMode Then
            If EntryMode = "N" Then
                chkINV_1099_IND.Checked = (rowAPTVEND1.Item("VEND_TAX_ID") & "" <> "")
            Else
                chkINV_1099_IND.Checked = (Val(rowAPTINVH1.Item("INV_1099_AMT") & "") <> 0)
            End If

            If EntryMode = "N" Then
                tabReceipts.SelectedTab = tabReceipts.Tabs("Open Accrued PO Receipts")
            End If
        End If

        If ScreenMode Then

            If EntryMode = "B" Then
                tabMain.Visible = False
                grdAPTINVHB.Visible = True
                UltraGroupBox1.Visible = False
            End If

            If EntryMode = "M" Then
                tabMain.Visible = False
                grdAPTINVHM.Visible = True
                lblCOLUMN_NAME.Visible = False
                lblNEW_VALUE.Visible = False
                grpLastChange.Visible = True
            End If

            If ApprovalMode Then
                tabMain.SelectedTab = tabMain.Tabs("Header Data")
                tabMain.SelectedTab = tabMain.Tabs("Header Data")
            End If

            btnAddAccruals.Visible = Not InquiryMode
            Absx1.txtFor("VEND_CODE_ADD").Visible = Not InquiryMode
            lblAddSupplier.Visible = Not InquiryMode

        Else
            UltraGroupBox1.Visible = True
            Clear_Record()

            lblAUTHs.Text = ""
            lblAUTHs.Visible = False
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"APTINVH1", "APTINVH2", "APTCHCK1", "APTCHCK2", "ASTATTA2",
             "TATEVNT1", "APTINVHX", "ASTAUDTX", "APTINVH5_SUM", "APTINVH5", "APTINVH7", "APTINVH8", "ICTIREC1", "ICTIREC2", "APTVEND1", "APTVEND2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        If Not batch_update Then
            dst.Tables("APTINVHB").Rows.Clear()
        End If

        EnforceConstraints(True)

        If HFs.ContainsKey("VEND_CODE") AndAlso HFs("VEND_CODE") <> "" Then
            Absx1.txtFor("VEND_CODE").Text = HFs("VEND_CODE")
        End If

        'Fill_Records("APTINVR1")
        Show_Batch()

        grpBankingInfo.Tag = ""

        Absx1.numFor("DIST_GL").Value = 0
        Absx1.numFor("DIST_PO").Value = 0
        Absx1.numFor("DIST_OTHER").Value = 0

        Absx1.cbeFor("INV_TYPE").Value = "I"

        Absx1.txtFor("QE_INV_REF").Text = ""
        Absx1.dteFor("QE_INV_DATE").Value = Null
        Absx1.numFor("QE_INV_AMT").Value = 0

        Absx1.numFor("CHECK_AMT").Value = 0

        Setup_QE(chkQuickEntry.Checked)

        lblRecurringTemplate.Visible = False
        lblRecurring.Visible = False
        tabHeader.SelectedTab = tabHeader.Tabs("Codes")
        tabHeader.Tabs("Pymt Info").Enabled = False

        optShow.Value = "U"
        OPT_SHOW_REFRESH()

        txtPPDDutyBOL.Text = ""
        'numPPDDuty.Value = 0

        INVOICE_FROM_EMAIL = ""
        UltraTabControl1.SelectedTab = UltraTabControl1.Tabs("A/P Invoices")
        lblSUBMITTED_INVOICE1.Visible = False
        lblSUBMITTED_INVOICE2.Visible = False
        lblSUBMITTED_INVOICE3.Visible = False

        ASCMAIN1.Progress("", "")
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        If chkQuickEntry.Checked Then
            Setup_QE(False)
        End If

        If EntryMode = "N" Then
            HFs("VOUCHER_NO") = ASCMAIN1.Next_Control_No("APTINVH1.VOUCHER_NO")
        End If

        Dim X As CurrencyManager = Me.BindingContext(dst.Tables("APTINVH1"))
        X.SuspendBinding()

        EnforceConstraints(False)

        rowAPTINVH1 = Fill_Record("APTINVH1", New String() {HFs("VOUCHER_NO")}, EntryMode = "N")
        If EntryMode = "E" Then
            HFs("VEND_CODE") = rowAPTINVH1.Item("VEND_CODE")
            HFs("INV_TYPE") = rowAPTINVH1.Item("INV_TYPE")
            'HFs("VEND_NAME") = rowAPTINVH1.Item("VEND_NAME")
            HFs("INV_NUM") = rowAPTINVH1.Item("INV_NUM")
        End If
        rowAPTVEND1 = Fill_Record("APTVEND1", HFs("VEND_CODE"))
        X.ResumeBinding()

        Fill_Records("APTINVH2", New String() {HFs("VOUCHER_NO")})
        Fill_Records("APTINVH5", New String() {HFs("VOUCHER_NO")})
        Fill_Records("APTINVH8", New String() {HFs("VOUCHER_NO")})
        Fill_Records("APTINVH7", New String() {HFs("VOUCHER_NO"), HFs("VEND_CODE")})
        Fill_Records("APTACRC1", New String() {HFs("VOUCHER_NO"), HFs("VEND_CODE")})
        auto_next_check = False
        BANK_LAST_CHECK_NO = ""
        BANK_NEXT_CHECK_NO = ""


        If EntryMode = "N" Then
            With rowAPTINVH1
                .Item("VOUCHER_NO") = HFs("VOUCHER_NO")
                .Item("VEND_CODE") = HFs("VEND_CODE")
                .Item("INV_TYPE") = HFs("INV_TYPE")
                .Item("INV_NUM") = HFs("INV_NUM")

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
                .Item("CURR_CODE") = rowAPTVEND1.Item("CURR_CODE")
                If .Item("CURR_CODE") & "" = "" Then
                    .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
                End If
                If .Item("CURR_CODE") & "" = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") & "" Then
                    .Item("CURR_EXCH_RATE") = 1
                End If

                Set_Recurring(chkRecurring.Checked)
                If chkRecurring.Checked Then
                    .Item("INV_STATUS") = "R"
                Else
                    .Item("INV_STATUS") = "O"
                End If

                .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")

                .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                .Item("REGISTER_IND") = "0"
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID

                .Item("VEND_BUYER_CODE") = rowAPTVEND1.Item("VEND_BUYER_CODE")

            End With

            chkACCRUE_PRIOR.Checked = False
        Else
            Save_Header_Fields(UltraGroupBox1)
            If EntryMode = "E" Then
                If ApprovalMode Then
                    Write_Event_Log(TABLE_NAME, HFs("VOUCHER_NO"), "Called up for Approval")
                Else
                    If InquiryMode Then
                    Else
                        Write_Event_Log(TABLE_NAME, HFs("VOUCHER_NO"), "Called up for Changes")
                    End If
                End If
            End If

            If rowAPTINVH1.Item("OPS_YYYYPP_ACCRUE") & "" <> "" Then
                chkACCRUE_PRIOR.Checked = True
            End If
            Set_Recurring(rowAPTINVH1.Item("INV_STATUS") = "R")


            ASCMAIN1.sql = "Select * from APTSUBM1 where VOUCHER_NO = '" & HFs("VOUCHER_NO") & "'"
            Dim rowAPTSUBM1 As DataRow = ASCDATA1.GetDataRow
            Dim PO_QTY_OPN As Integer = 0
            If rowAPTSUBM1 IsNot Nothing Then
                lblSUBMITTED_INVOICE1.Text = "From: " & rowAPTSUBM1.Item("SUBMIT_EMAIL_FROM") & ""
                lblSUBMITTED_INVOICE1.Visible = True
                lblSUBMITTED_INVOICE2.Text = "Subj: " & rowAPTSUBM1.Item("SUBMIT_SUBJECT") & ""
                lblSUBMITTED_INVOICE2.Visible = True
                lblSUBMITTED_INVOICE3.Text = "Submitted:  " & rowAPTSUBM1.Item("SUBMIT_DATE_RECEIVED") & ""
                lblSUBMITTED_INVOICE3.Visible = True
            End If



            Dim INV_REF As String = rowAPTINVH1.Item("INV_REF") & ""
            If INV_REF.Length = 10 Then

                Dim rowSPTPYMT1 As DataRow = LookUp("SPTPYMT1", INV_REF)
                If rowSPTPYMT1 IsNot Nothing Then
                    ' PROBABLY SHOULD CHECK THE AMT
                    ASCMAIN1.sql = $"Select * from SPTPYMT2 where PYMT_NO = '{INV_REF}'"
                    Dim AUTH_NOs As String = ""
                    For Each rowSPTPYMT2 As DataRow In ASCDATA1.GetDataTable.Select("")
                        Dim AUTH_NO As String = rowSPTPYMT2.Item("AUTH_NO") & ""
                        AUTH_NOs &= vbCrLf & AUTH_NO
                    Next

                    lblAUTHs.Text = "Auths:" & AUTH_NOs
                    lblAUTHs.Visible = True
                End If

            End If

        End If


        cmdCheck.Visible = (rowAPTINVH1.Item("INV_STATUS") = "P")

        If rowAPTINVH1.Item("INV_STATUS") = "R" Then
            chkACCRUE_PRIOR.Checked = False
        End If

        Fill_Records("TATEVNT1", HFs("VOUCHER_NO"))
        Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)

        If EntryMode = "E" And Not InquiryMode And Not ApprovalMode Then
            Delete_Rows("APTINVH2", "INV_LTYP = 'P'")
            Delete_Rows("APTINVH2", "INV_LTYP = 'O'")
            ' MAYBE SHOULD DELETE ALL NON NULL INV_LTYP
        End If

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        Load_OPS_YYYYPP_ACCRUE()
        Load_CURR_EXCH_RATE()

        If Absx1.cbeFor("INV_TYPE").Value = "A" Then
            tabMain.Tabs("PO Receipts").Enabled = False
            tabMain.Tabs("Other Accruals").Enabled = False
            tabMain.Tabs("GL Distribution").Enabled = False
        Else
            tabMain.Tabs("GL Distribution").Enabled = True
            If rowAPTINVH1.Item("INV_STATUS") = "R" Then
                tabMain.Tabs("PO Receipts").Enabled = False
                tabMain.Tabs("Other Accruals").Enabled = False
                Fill_Records("APTINVHR", HFs("VOUCHER_NO"))
            Else
                Load_ICTIREC1()
            End If
        End If

        lblINV_STATUS_NOTE.Visible = False
        If rowAPTINVH1.Item("VOUCHER_NO_ORIG") & "" <> "" Then
            lblINV_STATUS_NOTE.Text = "Original Voucher " & rowAPTINVH1.Item("VOUCHER_NO_ORIG")
            lblINV_STATUS_NOTE.Visible = True
        ElseIf rowAPTINVH1.Item("INV_STATUS") & "" = "D" Then
            ASCMAIN1.sql = "Select Min (VOUCHER_NO) from APTINVH1 where VOUCHER_NO_ORIG = '" & rowAPTINVH1.Item("VOUCHER_NO") & "'"
            Dim VOUCHER_NO_reversing As String = ASCDATA1.GetDataValue
            lblINV_STATUS_NOTE.Text = "Reversed by Voucher " & VOUCHER_NO_reversing
            lblINV_STATUS_NOTE.Visible = (VOUCHER_NO_reversing <> "")
        End If

        lblRecurringTemplate.Visible = (rowAPTINVH1.Item("INV_STATUS") = "R")
        lblRecurring.Text = "from Recurring Template " & rowAPTINVH1.Item("VOUCHER_NO_RECUR")
        lblRecurring.Visible = (rowAPTINVH1.Item("VOUCHER_NO_RECUR") & "" <> "")

        tabHeader.Tabs("Adjustments").Enabled = Not (rowAPTINVH1.Item("INV_STATUS") = "R")
        tabHeader.Tabs("Recurring").Enabled = (rowAPTINVH1.Item("INV_STATUS") = "R")
        tabHeader.Tabs("Recurring").Visible = False
        tabHeader.Tabs("Pymt Info").Enabled = (rowAPTINVH1.Item("INV_STATUS") = "P")

        If EntryMode = "N" Then
            Absx1.dteFor("INV_BL_DATE").ReadOnly = InquiryMode Or ApprovalMode
        Else
            '   Setup_INV_BL_DATE()
        End If

        Absx1.txtFor("INV_PYMT_METHOD").ReadOnly = (rowAPTVEND1.Item("VEND_PYMT_METHOD_FIXED") & "" = "1") Or InquiryMode Or ApprovalMode

        tabMain.SelectedTab = tabMain.Tabs(0)
        grdAPTINVH2.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdAPTINVH2.DisplayLayout.Bands(0).SortedColumns.Add("VOUCHER_LNO", False)

        If EntryMode = "N" Then
            ' CAN'T GET FOCUS TO GO TO A SPECIFIC CONTROL
            tabMain.SelectedTab = tabMain.Tabs(1)
            tabMain.Focus()
            Application.DoEvents()
            For i As Integer = 1 To 5
                If Absx1.CtlFor("INV_DATE").Focused Then
                    Exit For
                End If
                SendKeys.Send(Chr(9))
                Application.DoEvents()
            Next
        End If

        If rowAPTINVH1.Item("INV_STATUS") & "" = "P" Then

            ASCMAIN1.sql = "Select APTCHCK2.VOUCHER_NO, SUBSTR(APTCHCK2.INV_NUM,1,20) INV_NUM" _
            & ", APTCHCK2.INV_DATE" _
            & ", NVL(APTCHCK2.INV_AMT_APPLIED,0) INV_BALANCE" _
            & ", NVL(APTCHCK2.INV_DISC_TAKEN,0) INV_DISC_AMT" _
            & " from APTCHCK2 " _
            & " where BANK_CODE = '" & rowAPTINVH1.Item("BANK_CODE") & "' and CHECK_NUM = '" & rowAPTINVH1.Item("CHECK_NUM") & "'"
            Fill_Records("APTINVHX", "", True, ASCMAIN1.sql)

            Dim BANK_CODE As String = rowAPTINVH1.Item("BANK_CODE") & ""
            Dim CHECK_NUM As String = rowAPTINVH1.Item("CHECK_NUM") & ""

            Dim rowAPTCHCK1 As DataRow = LookUp("APTCHCK1", New String() {BANK_CODE, CHECK_NUM})
            Dim CHECK_AMT As Decimal = 0
            If rowAPTCHCK1 IsNot Nothing Then
                CHECK_AMT = Val(rowAPTCHCK1.Item("CHECK_AMT") & "")
            End If
            Absx1.numFor("CHECK_AMT").Value = CHECK_AMT
        Else
            Fill_Records("APTINVHX", New Object() {HFs("VEND_CODE"), HFs("VOUCHER_NO")})
        End If

        Setup_APTCHCK5()

        Fill_Records("ASTAUDTX", HFs("VOUCHER_NO"))
        Sort_grdColumns(grdASTAUDTX, "INIT_DATE".ToLower)

        EnforceConstraints(True)

        tabHeader.ActiveTab = tabHeader.Tabs("Codes")

        Calc_DIST_GL()
        Calc_DIST_PO()
        Calc_DIST_Adjustments()
        Calc_DIST_Other()

        If ApprovalMode Then
            cbeAPPR_ROUTE_TO.Value = ""
            optAPPR_ACTION.CheckedIndex = -1
        End If
    End Sub

    Sub Update_Record()
        Dim X As CurrencyManager = Me.BindingContext(dst.Tables("APTINVH1"))
        X.EndCurrentEdit()
        X.SuspendBinding()

        Dim VOUCHER_NO_TO_PAY As String = HFs("VOUCHER_NO")

        Try

            BeginTrans()

            Calculate_INV_DISC_AMT()

            Delete_Rows("APTINVH2", "INV_LTYP is Null and (INV_LINE_AMT = 0 or INV_LINE_AMT is Null)")

            If rowAPTINVH1("OPS_YYYYPP_ACCRUE") & "" _
             = rowAPTINVH1("OPS_YYYYPP") & "" Then
                rowAPTINVH1("OPS_YYYYPP_ACCRUE") = ""
            End If

            Dim pay_upon_entry As Boolean = False
            rowAPTINVH1("INV_BALANCE") = rowAPTINVH1("INV_AMT")
            If rowAPTINVH1("INV_STATUS") = "P" Then
                rowAPTINVH1("INV_PAID_UPON_ENTRY") = "1"
                pay_upon_entry = True
            Else
                rowAPTINVH1("INV_PAID_UPON_ENTRY") = Null
            End If

            If rowAPTINVH1("INV_STATUS") = "R" Then
                rowAPTINVH1("REGISTER_IND") = "R"
            End If

            If rowAPTINVH1("INV_REMIT_TO") = "N" Then
                dst.Tables("APTVEND2").AcceptChanges()
                dst.Tables("APTVEND2").Rows(0).SetAdded()
                Update_Record_TDA("APTVEND2")
                rowAPTINVH1("INV_REMIT_TO") = "A"
            End If

            rowAPTVEND5 = Fill_Record("APTVEND5", HFs("VEND_CODE"), True)

            Create_APTINVH5_VAR()
            Create_APTINVH2_P()
            'Create_APTINVH2_R()
            Create_APTINVH2_from_APTINVH7()

            Dim CTL_NOs As New List(Of String)

            For Each row As DataRow In dst.Tables("APTACRC1").Select("", "", DataViewRowState.ModifiedCurrent)
                Dim COST_ACC As Decimal = Val(row.Item("COST_ACC") & "")
                Dim COST_ACC_orig As Decimal = Val(row.Item("COST_ACC", DataRowVersion.Original) & "")
                If COST_ACC <> COST_ACC_orig Then
                    Dim CTL_NO As String = row.Item("CTL_NO")
                    CTL_NOs.Add(CTL_NO)
                    ASCMAIN1.sql = "Update APTACRC1 Set COST_ACC = :PARM1 where CTL_NO = :PARM2"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "NV", New Object() {COST_ACC, CTL_NO})
                End If
                row.AcceptChanges() ' do not update rows which were not newly added - they will get updated by dependent updates
            Next
            For Each row As DataRow In dst.Tables("APTACRC1").Select("", "", DataViewRowState.Added)
                ' how is this not killing the pre-paid records which have nothing but COST_ACT?
                ' 07/01/25 - it was killing them
                'row.Item("COST_ACT") = 0
                Dim CTL_NO As String = row.Item("CTL_NO")
                CTL_NOs.Add(CTL_NO)
            Next
            Update_Record_TDA("APTACRC1") ' TO GET ADDED ROWS OUT THERE - DEPENDENT UPDATES WILL DO THE REST

            ' MAYBE THIS SHOULD BE HANDLED IN LOOPS ABOVE?
            If CTL_NOs.Count > 0 Then
                ASCMAIN1.sql = $"Update APTACRC1 Set SOURCE_DOC_NO = (Select CTL_REF_NO from APTINVH7 where VOUCHER_NO = '{HFs("VOUCHER_NO")}' and CTL_NO = APTACRC1.CTL_NO) where CTL_STATUS = 'P' and CTL_NO IN ('{Join(CTL_NOs.ToArray, "','")}')"
                ASCDATA1.ExecuteSQL()
            End If

            If EntryMode = "N" Then
                If rowAPTINVH1("INV_PAID_UPON_ENTRY") & "" = "1" Then
                    Write_Event_Log(TABLE_NAME, HFs("VOUCHER_NO"), "Entered as Paid")
                Else
                    Write_Event_Log(TABLE_NAME, HFs("VOUCHER_NO"), "Entered")
                End If

                Check_for_Approval()



                Update_Record_TDA("APTINVH1")
                Update_Record_TDA("APTINVH2")
                Update_Record_TDA("APTINVH8")

                'For Each row As DataRow In dst.Tables("APTACRC1").Select("", "", DataViewRowState.ModifiedCurrent)
                '    row.AcceptChanges() ' do not update rows which were not newly added - they will get updated by dependent updates
                'Next
                'Update_Record_TDA("APTACRC1") ' TO GET ADDED ROWS OUT THERE - DEPENDENT UPDATES WILL DO THE REST

                For Each rowAPTINVH5 As DataRow In dst.Tables("APTINVH5").Select("", "", DataViewRowState.CurrentRows)
                    rowAPTINVH5.Item("VAR_QTY") = rowAPTINVH5.Item("QTY_VAR")
                    rowAPTINVH5.Item("VAR_AMT") = rowAPTINVH5.Item("AMT_VAR")
                Next
                Update_Record_TDA("APTINVH5", "VOUCHER_NO = '" & HFs("VOUCHER_NO") & "'")

                Update_Record_TDA("APTINVH7", "VOUCHER_NO = '" & HFs("VOUCHER_NO") & "'")

                dst.Tables("APTINVH5").Rows.Clear()
                dst.Tables("APTINVH8").Rows.Clear()
                dst.Tables("APTINVH7").Rows.Clear()

                Dependent_Updates(HFs("VOUCHER_NO"), False)




                If rowAPTINVH1("INV_TYPE") = "A" Then ' If Absx1.cbeFor("INV_TYPE").Value = "A" Then
                    Dim VOUCHER_NO_ADV As String = ASCMAIN1.Next_Control_No("APTINVH1.VOUCHER_NO")

                    ReNumber_Voucher(HFs("VOUCHER_NO"), VOUCHER_NO_ADV)
                    rowAPTINVH1("INV_PAID_UPON_ENTRY") = Null
                    rowAPTINVH1("INV_STATUS") = "O"
                    rowAPTINVH1("CHECK_NUM") = ""
                    rowAPTINVH1("CHECK_DATE") = Null
                    Negate_Voucher(VOUCHER_NO_ADV)

                    Update_Record_TDA("APTINVH1")
                    Update_Record_TDA("APTINVH2")
                End If
                'Stop
            Else
                Write_Audit_Trail(rowAPTINVH1, Nothing, "E")

                Dim something_GL_related_was_changed As Boolean = False
                If dst.Tables("APTINVH2").GetChanges IsNot Nothing _
                    Or dst.Tables("APTINVH5").GetChanges IsNot Nothing _
                    Or dst.Tables("APTINVH7").GetChanges IsNot Nothing Then
                    something_GL_related_was_changed = True
                End If
                If dst.Tables("APTINVH2").GetChanges IsNot Nothing Then
                    something_GL_related_was_changed = True
                End If
                For Each COLUMN_NAME As String In New String() _
                {"INV_AMT", "INV_DATE", "OPS_YYYYPP_ACCRUE", "POST_CODE", "CURR_CODE", "CURR_EXCH_RATE", "INV_PAID_UPON_ENTRY"}
                    If rowAPTINVH1(COLUMN_NAME, DataRowVersion.Current) & "" _
                    <> rowAPTINVH1(COLUMN_NAME, DataRowVersion.Original) & "" Then
                        something_GL_related_was_changed = True
                    End If
                Next

                If rowAPTINVH1("REGISTER_IND") = "0" _
                Or rowAPTINVH1("REGISTER_IND") = "R" _
                Or Not something_GL_related_was_changed Then
                    Write_Event_Log(TABLE_NAME, HFs("VOUCHER_NO"), "Changed")
                    Dependent_Updates(HFs("VOUCHER_NO"), True)

                    rowAPTINVH1("LAST_DATE") = DATETIME_STAMP
                    rowAPTINVH1("LAST_OPER") = ASCMAIN1.USER_ID

                    If rowAPTINVH1.Item("INV_APPR_STATUS") & "" <> "A" Or something_GL_related_was_changed Then
                        Check_for_Approval()
                    End If

                    Dim sql_delete As String = "VOUCHER_NO = '" & HFs("VOUCHER_NO") & "'"
                    Update_Record_TDA("APTINVH1", sql_delete)
                    Update_Record_TDA("APTINVH2", sql_delete)

                    For Each rowAPTINVH5 As DataRow In dst.Tables("APTINVH5").Select("", "", DataViewRowState.CurrentRows)
                        rowAPTINVH5.Item("VAR_QTY") = rowAPTINVH5.Item("QTY_VAR")
                        rowAPTINVH5.Item("VAR_AMT") = rowAPTINVH5.Item("AMT_VAR")
                    Next
                    Update_Record_TDA("APTINVH5", sql_delete)

                    Update_Record_TDA("APTINVH8", sql_delete)
                    Update_Record_TDA("APTINVH7", sql_delete)

                    Dependent_Updates(HFs("VOUCHER_NO"), False)
                Else
                    Dim VOUCHER_NO_ORIG As String = HFs("VOUCHER_NO")
                    Dim VOUCHER_NO_NEG As String = ASCMAIN1.Next_Control_No("APTINVH1.VOUCHER_NO")
                    Dim VOUCHER_NO_NEW As String = ASCMAIN1.Next_Control_No("APTINVH1.VOUCHER_NO")
                    VOUCHER_NO_TO_PAY = VOUCHER_NO_NEW

                    Write_Event_Log(TABLE_NAME, VOUCHER_NO_ORIG, "Reversed (" & VOUCHER_NO_NEG & ") and Replaced (" & VOUCHER_NO_NEW & ")")

                    Dependent_Updates(VOUCHER_NO_ORIG, True)

                    ReNumber_Voucher(VOUCHER_NO_ORIG, VOUCHER_NO_NEW)
                    rowAPTINVH1("REGISTER_IND") = "0"
                    rowAPTINVH1("REGISTER_XNO") = Null
                    rowAPTINVH1("OPS_YYYYPP") = ASCMAIN1.CYP
                    rowAPTINVH1("LAST_DATE") = DATETIME_STAMP
                    rowAPTINVH1("LAST_OPER") = ASCMAIN1.USER_ID
                    Check_for_Approval()



                    Update_Record_TDA("APTINVH1")
                    Update_Record_TDA("APTINVH2")

                    For Each rowAPTINVH5 As DataRow In dst.Tables("APTINVH5").Select("", "", DataViewRowState.CurrentRows)
                        rowAPTINVH5.Item("VAR_QTY") = rowAPTINVH5.Item("QTY_VAR")
                        rowAPTINVH5.Item("VAR_AMT") = rowAPTINVH5.Item("AMT_VAR")
                    Next
                    Update_Record_TDA("APTINVH5")

                    Update_Record_TDA("APTINVH8")
                    Update_Record_TDA("APTINVH7")

                    Dependent_Updates(VOUCHER_NO_NEW, False)

                    Fill_Records("ASTATTA2", VOUCHER_NO_ORIG)
                    For Each row As DataRow In dst.Tables("ASTATTA2").Select("")
                        row.Item("CODE_VALUE") = VOUCHER_NO_NEW
                        row.AcceptChanges()
                        row.SetAdded()
                    Next
                    Update_Record_TDA("ASTATTA2")

                    ReLoad_Voucher(VOUCHER_NO_ORIG)
                    rowAPTINVH1 = dst.Tables("APTINVH1").Rows(0)
                    rowAPTINVH1("LAST_DATE") = DATETIME_STAMP
                    rowAPTINVH1("LAST_OPER") = ASCMAIN1.USER_ID
                    rowAPTINVH1("INV_STATUS") = "D"
                    Update_Record_TDA("APTINVH1")

                    ReNumber_Voucher(VOUCHER_NO_ORIG, VOUCHER_NO_NEG)
                    Negate_Voucher(VOUCHER_NO_NEG)
                    rowAPTINVH1("REGISTER_IND") = "0"
                    rowAPTINVH1("REGISTER_XNO") = Null
                    rowAPTINVH1("OPS_YYYYPP") = ASCMAIN1.CYP

                    Update_Record_TDA("APTINVH1")
                    Update_Record_TDA("APTINVH2")

                    For Each rowAPTINVH5 As DataRow In dst.Tables("APTINVH5").Select("", "", DataViewRowState.CurrentRows)
                        rowAPTINVH5.Item("VAR_QTY") = rowAPTINVH5.Item("QTY_VAR")
                        rowAPTINVH5.Item("VAR_AMT") = rowAPTINVH5.Item("AMT_VAR")
                    Next
                    Update_Record_TDA("APTINVH5")

                    Update_Record_TDA("APTINVH8")
                    Update_Record_TDA("APTINVH7")
                End If
            End If


            If pay_upon_entry Then ' If rowAPTINVH1("INV_PAID_UPON_ENTRY") & "" = "1" Then

                Update_as_Paid(VOUCHER_NO_TO_PAY)
                Update_Record_TDA("APTCHCK1")
                Update_Record_TDA("APTCHCK2")
                Update_Record_TDA("APTINVH1")
                If auto_next_check Then
                    Update_Record_TDA("GLTBANK1")
                End If
            End If
            Update_Record_TDA("APTVEND5")

            If INVOICE_FROM_EMAIL <> "" Then
                Update_APTSUBM1()
                Update_Record_TDA("APTSUBM1")
            End If



            If batch_update Then
                CommitTrans()
                X.ResumeBinding()
            Else
                X.ResumeBinding()
                CommitTrans("Update Complete")

                If pay_upon_entry Then
                    '******************************** REMOVE THE REM
                    Print_Check("")
                End If
            End If

            ' UPDATE 

        Catch ex As Exception
            X.ResumeBinding()
            Rollback("Error Occurred - Please call ABS", ex)
        End Try

    End Sub

    Sub Update_Approval()

        BeginTrans()

        If optAPPR_ACTION.Value = "A" Then
            rowAPTINVH1.Item("INV_APPR_STATUS") = "A"
            Write_Event_Log(TABLE_NAME, HFs("VOUCHER_NO"), "Approved")
        Else
            rowAPTINVH1.Item("VEND_BUYER_CODE") = cbeAPPR_ROUTE_TO.Value
            If optAPPR_ACTION.Value = "P" Then
                Write_Event_Log(TABLE_NAME, HFs("VOUCHER_NO"), "Routed as Pending")
            ElseIf optAPPR_ACTION.Value = "R" Then
                Write_Event_Log(TABLE_NAME, HFs("VOUCHER_NO"), "Routed as Approved")
            End If
        End If
        Update_Record_TDA("APTINVH1")

        ASCMAIN1.sql = "Select X.*, FATFACL1.ASSET_CLASS_DESC, FATFACL1.ASSET_CLASS_EMAIL" & vbCrLf _
            & " from FATFACL1, (" & vbCrLf _
            & "Select APTINVH1.VOUCHER_NO, APTINVH1.INV_NUM, APTINVH1.INV_AMT" & vbCrLf _
            & ", APTINVH1.VEND_CODE, APTINVH2.INV_LINE_AMT" & vbCrLf _
            & ", FATFACL1.ACCT_CODE_CAP" & vbCrLf _
            & ", MIN (FATFACL1.ASSET_CLASS_CODE) ASSET_CLASS_CODE" & vbCrLf _
            & " from APTINVH1,APTINVH2,FATFACL1" & vbCrLf _
            & " where APTINVH1.VOUCHER_NO = APTINVH2.VOUCHER_NO" & vbCrLf _
            & "   and APTINVH2.ACCT_CODE = FATFACL1.ACCT_CODE_CAP" & vbCrLf _
            & $"   And APTINVH1.VOUCHER_NO = '{HFs("VOUCHER_NO")}'" & vbCrLf _
            & " group by APTINVH1.VOUCHER_NO, APTINVH1.INV_NUM, APTINVH1.INV_AMT" & vbCrLf _
            & ", APTINVH1.VEND_CODE, APTINVH2.INV_LINE_AMT" & vbCrLf _
            & ", FATFACL1.ACCT_CODE_CAP) X" & vbCrLf _
            & " where FATFACL1.ASSET_CLASS_CODE = X.ASSET_CLASS_CODE" & vbCrLf _
            & "   and FATFACL1.ASSET_CLASS_EMAIL IS NOT NULL"


        For Each row As DataRow In ASCDATA1.GetDataTable.Select()
            Dim VOUCHER_NO As String = row.Item("VOUCHER_NO")
            Dim INV_NUM As String = row.Item("INV_NUM")
            Dim INV_AMT As Decimal = Val(row.Item("INV_AMT") & "")
            Dim INV_LINE_AMT As Decimal = Val(row.Item("INV_LINE_AMT") & "")
            Dim VEND_CODE As String = row.Item("VEND_CODE")
            Dim ACCT_CODE_CAP As String = row.Item("ACCT_CODE_CAP")
            Dim ASSET_CLASS_EMAIL As String = row.Item("ASSET_CLASS_EMAIL")
            Dim ASSET_CLASS_CODE As String = row.Item("ASSET_CLASS_CODE")
            Dim ASSET_CLASS_DESC As String = row.Item("ASSET_CLASS_DESC")

            Dim ALERT_SUBJECT As String = ""
            Dim ALERT_MESSAGE As String = ""

            If Not dst.Tables.Contains("TATALRT1") Then
                Create_TDA(dst.Tables.Add, "TATALRT1", "*")
            End If

            Dim rowTATALRT1 As DataRow = dst.Tables("TATALRT1").NewRow
            With rowTATALRT1
                Dim ALERT_NO As String = ASCMAIN1.Next_Control_No("TATALRT1.ALERT_NO")
                .Item("ALERT_NO") = ALERT_NO
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("FORM_NAME") = "APFINVH1"
                .Item("FORM_KEY") = ALERT_NO
                .Item("ALERT_EMAIL") = ASSET_CLASS_EMAIL
                .Item("ALERT_EML") = "1"

                .Item("ALERT_EML_DATE") = DATETIME_STAMP
                ALERT_SUBJECT = "Fixed Asset Added for Class " & ASSET_CLASS_CODE & ":" & ASSET_CLASS_DESC
                .Item("ALERT_SUBJECT") = Mid(ALERT_SUBJECT, 1, 200)
                ALERT_MESSAGE = "Control No: " & ALERT_NO & vbCrLf & $"Fixed Asset Added for Class {ASSET_CLASS_CODE} : {ASSET_CLASS_DESC}"
                ALERT_MESSAGE &= vbCrLf & $"Vendor: {VEND_CODE}, Invoice {INV_NUM}, Amount {Format(INV_LINE_AMT, "$#,##0.00")}"
                .Item("ALERT_MESSAGE") = Mid(ALERT_MESSAGE, 1, 2000)
            End With
            dst.Tables("TATALRT1").Rows.Add(rowTATALRT1)

            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
            EMAIL_ADDRESSs.Add(ASSET_CLASS_EMAIL, "Fixed Asset Admin")

            Dim SEND_NO As String = ""
            If ASCMAIN1.Running_in_VS Then
                SEND_NO = "TESTING"
                Stop
            Else
                SEND_NO = ASCMAIN1.TACMAIN1.Send_email _
                (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, Nothing,
                ALERT_SUBJECT, "FA_ADD", True, False, ASSET_CLASS_CODE, ASSET_CLASS_CODE, "FA Class Code", ALERT_MESSAGE)
            End If

            rowTATALRT1.Item("SEND_NO") = SEND_NO
            Update_Record_TDA("TATALRT1")

        Next

        CommitTrans("Update Complete")

    End Sub

    Sub Check_for_Approval()

        If ROWs("APTPARM1").Item("AP_PARM_APPR_REQD") & "" = "1" Then
            If Val(rowAPTINVH1.Item("INV_AMT") & "") > Val(ROWs("APTPARM1").Item("AP_PARM_APPR_LIMIT") & "") Then
                rowAPTINVH1.Item("INV_APPR_STATUS") = "P"

                If rowAPTVEND1.Item("VEND_AUTO_APPROVE") & "" = "1" Then
                    rowAPTINVH1.Item("INV_APPR_STATUS") = "A"
                    Write_Event_Log(TABLE_NAME, HFs("VOUCHER_NO"), "Auto Approved")
                End If

            Else
                rowAPTINVH1.Item("INV_APPR_STATUS") = "A"
            End If
        Else
            rowAPTINVH1.Item("INV_APPR_STATUS") = DBNull.Value
        End If

    End Sub
    Sub Delete_Record()

        Dim X As CurrencyManager = Me.BindingContext(dst.Tables("APTINVH1"))
        X.EndCurrentEdit()
        X.SuspendBinding()

        BeginTrans()

        If EntryMode = "N" Then
            Write_Event_Log(TABLE_NAME, HFs("VOUCHER_NO"), "Entered And Deleted before Update")
        Else
            Write_Event_Log(TABLE_NAME, HFs("VOUCHER_NO"), "Deleted")
            rowAPTVEND5 = Fill_Record("APTVEND5", HFs("VEND_CODE"), True)

            Dependent_Updates(HFs("VOUCHER_NO"), True)
            Dim VOUCHER_NO_ORIG As String = HFs("VOUCHER_NO")
            ReLoad_Voucher(VOUCHER_NO_ORIG)

            rowAPTINVH1 = dst.Tables("APTINVH1").Rows(0)
            rowAPTINVH1("LAST_DATE") = DATETIME_STAMP
            rowAPTINVH1("LAST_OPER") = ASCMAIN1.USER_ID
            'rowAPTINVH1("INV_BALANCE") = 0 ' NOT NEC AND NOT CONSISTENT WITH OTHER D'S
            rowAPTINVH1("INV_STATUS") = "D"
            If rowAPTINVH1("REGISTER_IND") = "0" Then
                rowAPTINVH1("REGISTER_IND") = "D"
            End If
            Update_Record_TDA("APTINVH1")
            Update_Record_TDA("APTVEND5")

            If rowAPTINVH1("REGISTER_IND") = "1" Then
                Dim VOUCHER_NO_NEG As String = ASCMAIN1.Next_Control_No("APTINVH1.VOUCHER_NO")

                Write_Event_Log(TABLE_NAME, VOUCHER_NO_ORIG, "Reversed (" & VOUCHER_NO_NEG & ")")

                ReNumber_Voucher(VOUCHER_NO_ORIG, VOUCHER_NO_NEG)
                Negate_Voucher(VOUCHER_NO_NEG)
                rowAPTINVH1("REGISTER_IND") = "0"
                rowAPTINVH1("REGISTER_XNO") = Null
                rowAPTINVH1("OPS_YYYYPP") = ASCMAIN1.CYP

                Update_Record_TDA("APTINVH1")
                Update_Record_TDA("APTINVH2")
                Update_Record_TDA("APTINVH5")
                Update_Record_TDA("APTINVH8")
                Update_Record_TDA("APTINVH7")
            End If
        End If

        X.ResumeBinding()
        CommitTrans("Deletion Completed")


    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
            & " where VOUCHER_NO = '" & HFs("VOUCHER_NO") & "'")
    End Sub

    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "Load"
                Absx1.txtFor("VOUCHER_NO").Text = key
                Click_Command("Load")
        End Select

        Return return_key
    End Function

    Public Overrides Function Data_Export_Context() As ABSolution.ASFBASE0.Data_Export_Entity

        Dim E As New Data_Export_Entity
        E.enabled = True
        ASTDATA1s.Clear()
        'ASTDATA1s.Add("APTINVHX", "Vendor Invoices")
        ASTDATA1s.Add("APTINVH5", "Invoiced Purchase Accruals")
        ASTDATA1s.Add("APTINVH8", "Invoice Adjustments")
        ASTDATA1s.Add("APTINVH7", "Invoiced Other PO Accruals")

        Return E
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "APTINVH1"
            E.COLUMN_NAME = "VOUCHER_NO"
            E.CODE_VALUE = Absx1.txtFor("VOUCHER_NO").Text
            E.DESC_VALUE = "Vendor Invoice"
            E.ATTACHMENT_NOTES = ""
            If rowAPTINVH1.Item("INV_STATUS") & "" <> "O" And rowAPTINVH1.Item("INV_STATUS") & "" <> "H" Then
                E.RESTRICTIONS = "D"
            End If
            'E.READ_ONLY = True

            E.OTHER_ENTITIES = New List(Of Dropped_On_Entity_Other)

            Dim E_other As New Dropped_On_Entity_Other
            E_other.TABLE_NAME = "SPTPYMT1"
            E_other.COLUMN_NAME = "PYMT_NO"
            E_other.COLUMN_NAME_linked = "PYMT_CTL_NO"
            E.OTHER_ENTITIES.Add(E_other)
        End If

        Return E
    End Function

    Public Overrides Function Audit_Context() As Audit_Entity

        Dim E As New Audit_Entity
        If ScreenMode Then
            E.TABLE_NAME = "APTINVH1"
            E.KEY_VALUE = Absx1.txtFor("VOUCHER_NO").Text
            E.KEY_DESC = "Vendor Invoice"
        End If
        Return E
    End Function

    Public Overrides Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        E.TABLE_NAME = "APTINVH1"
        E.TABLE_KEY_CAPTION = "AP"
        If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = Absx1.txtFor("VOUCHER_NO").Text
            E.TABLE_KEY_DESC = Absx1.txtFor("VEND_CODE").Text & " " & Absx1.txtFor("VEND_NAME").Text
            E.TABLE_KEY_locked = ScreenMode And (EntryMode = "E" Or EntryMode = "A")
        Else
            E.enabled = False
            E.read_only = True
            E.TABLE_KEY_locked = False
            E.TABLE_KEY = ""
        End If

        Return E
    End Function

#End Region

#Region "Popup_Menus"
    Overrides Sub Load_Popup_Menus()
        If (MENU_ITEM_OBJECT = "APFINVHA") Then
            Load_Popup_Menu(grdAPTINVR1, "SS", "Show Filter", "Show GroupBox", "Refresh")
        End If
        Load_Popup_Menu(grdICTIREC1, "SSBBB", "Show Filter", "Show GroupBox", "PO Inquiry", "Retrieve Closed Accrued PO", "Retrieve Accrued POs from Voucher", "Retrieve Accrued POs from Voucher (w/Prev Inv Cost)")
        Load_Popup_Menu(grdAPTINVH5_SUM, "SSB", "Show Filter", "Show GroupBox", "PO Inquiry")
        Load_Popup_Menu(grdAPTINVH5, "SSSSSBB", "Show Filter", "Show GroupBox", "Show Description", "Show Receipt Qty/Price/Amt", "Show Discrepancies Only", "Copy Price to All Lines", "Chargeback All Variances", "Allow All Variances", "Clone Line", "Clone Line (w/Adj)")
        Load_Popup_Menu(grdAPTSUBM1, "SSBB", "Show Filter", "Show GroupBox", "Delete Submitted Email", "Re-Submit Submitted Email")
        Load_Popup_Menu(grdAPTACRC1, "SSBBB", "Show Filter", "Show GroupBox", "Split Accrual", "PO Inquiry", "PO Receipts Inquiry")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        Try
            grd = GRDs(Mid(e.SourceControl.Name, 4))
        Catch ex As Exception
            e.Cancel = True
            Exit Sub
        End Try

        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If tlb_pop.Tools.Exists("Show Description") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Description"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns("ITEM_DESC").Hidden
        End If

        If tlb_pop.Tools.Exists("Show Receipt Qty/Price/Amt") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Receipt Qty/Price/Amt"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns("QTY_REC").Hidden
        End If

        If tlb_pop.Tools.Exists("Show Discrepancies Only") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Discrepancies Only"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = discrepancies_only
        End If

        If tlb_pop.Tools.Exists("Retrieve Closed Accrued PO") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Retrieve Closed Accrued PO"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = Not InquiryMode
        End If

        If tlb_pop.Tools.Exists("Retrieve Accrued POs from Voucher") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Retrieve Accrued POs from Voucher"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = Not InquiryMode
        End If
        If tlb_pop.Tools.Exists("Retrieve Accrued POs from Voucher (w/Prev Inv Cost)") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Retrieve Accrued POs from Voucher (w/Prev Inv Cost)"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = Not InquiryMode
        End If

        If tlb_pop.Tools.Exists("Copy Price to All Lines") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Copy Price to All Lines"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = Not InquiryMode
        End If
        If tlb_pop.Tools.Exists("Chargeback All Variances") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Chargeback All Variances"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = Not InquiryMode
        End If
        If tlb_pop.Tools.Exists("Allow All Variances") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Allow All Variances"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = Not InquiryMode
        End If
        If tlb_pop.Tools.Exists("Clone Line") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Clone Line"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = Not InquiryMode
        End If
        If tlb_pop.Tools.Exists("Clone Line (w/Adj)") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Clone Line (w/Adj)"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = Not InquiryMode
        End If

        'If tlb_pop.Tools.Exists("Re-Submit") Then
        '    If optShow.Value = "P" Then
        '        tlb_btn = DirectCast(tlb_pop.Tools("Re-Submit"), UltraWinToolbars.ButtonTool)
        '        tlb_btn.SharedProps.Visible = True
        '    Else
        '        tlb_btn = DirectCast(tlb_pop.Tools("Re-Submit"), UltraWinToolbars.ButtonTool)
        '        tlb_btn.SharedProps.Visible = False
        '    End If
        'End If


        'If tlb_pop.Tools.Exists("Delete Submitted Email") Then
        '    tlb_btn = DirectCast(tlb_pop.Tools("Delete Submitted Email"), UltraWinToolbars.ButtonTool)
        '    tlb_btn.SharedProps.Visible = (MENU_ITEM_OBJECT = "APFINVH1") And Not InquiryMode
        'End If

        'If tlb_pop.Tools.Exists("Re-Submit Submitted Email") Then
        '    tlb_btn = DirectCast(tlb_pop.Tools("Re-Submit Submitted Email"), UltraWinToolbars.ButtonTool)
        '    tlb_btn.SharedProps.Visible = (MENU_ITEM_OBJECT = "APFINVH1") And Not InquiryMode
        'End If

        If tlb_pop.Tools.Exists("Split Accrual") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Split Accrual"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = grdAPTACRC1.ActiveRow IsNot Nothing AndAlso grdAPTACRC1.ActiveRow.IsDataRow And Not InquiryMode
        End If


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdAPTSUBM1"
                    tlb_btn = DirectCast(tlb_pop.Tools("Re-Submit Submitted Email"), UltraWinToolbars.ButtonTool)
                    If optShow.Value = "P" Then
                        tlb_btn.SharedProps.Visible = (MENU_ITEM_OBJECT = "APFINVH1") And Not InquiryMode
                    Else
                        tlb_btn.SharedProps.Visible = False
                    End If
                    tlb_btn = DirectCast(tlb_pop.Tools("Delete Submitted Email"), UltraWinToolbars.ButtonTool)
                    If optShow.Value = "U" Then
                        tlb_btn.SharedProps.Visible = (MENU_ITEM_OBJECT = "APFINVH1") And Not InquiryMode
                    Else
                        tlb_btn.SharedProps.Visible = False
                    End If


                    ' 
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Description"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Columns("ITEM_DESC").Hidden = Not tlb_sbt.Checked

            Case "Show Receipt Qty/Price/Amt"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Columns("QTY_REC").Hidden = Not tlb_sbt.Checked
                grd.DisplayLayout.Bands(0).Columns("PO_COST").Hidden = Not tlb_sbt.Checked
                grd.DisplayLayout.Bands(0).Columns("AMT_REC").Hidden = Not tlb_sbt.Checked

            Case "Show Discrepancies Only"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                discrepancies_only = tlb_sbt.Checked
                If grdICTIREC1.ActiveRow IsNot Nothing Then
                    Dim RECEIPT_NO As String = grdICTIREC1.ActiveRow.Cells("RECEIPT_NO").Text
                    Setup_grdAPTINVH5(RECEIPT_NO)
                End If

            Case "Copy Price to All Lines"
                If grd.ActiveCell Is Nothing Then
                Else
                    If grd.ActiveCell.Column.Key = "INV_COST" Then
                        Dim INV_COST As Double = Val(grd.ActiveCell.Value & "")
                        If MsgBox("Copy Invoice Cost of " & Format(INV_COST, "#.00") & " to All Lines Displayed on This Receipt?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                            For Each grow As UltraWinGrid.UltraGridRow In grdAPTINVH5.Rows
                                grow.Cells("INV_COST").Value = INV_COST
                                grow.Update()
                            Next
                        End If

                    End If
                End If

            Case "Chargeback All Variances"
                If MsgBox("Chargeback All Variances on This Receipt?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                    For Each grow As UltraWinGrid.UltraGridRow In grdAPTINVH5.Rows
                        grow.Cells("CB").Value = "1"
                        grow.Update()
                    Next
                End If

            Case "Allow All Variances"
                If MsgBox("Allow All Variances on This Receipt?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                    For Each grow As UltraWinGrid.UltraGridRow In grdAPTINVH5.Rows
                        grow.Cells("CB").Value = "0"
                        grow.Update()
                    Next
                End If

            Case "Retrieve Accrued POs from Voucher", "Retrieve Accrued POs from Voucher (w/Prev Inv Cost)"

                Dim SQLW As String = $"VEND_CODE = '{Absx1.txtFor("VEND_CODE").Text}' and OPS_YYYYPP >= '{ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12)}'"
                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("VOUCHER_NO", , SQLW)
                ASCMAIN1.CodeSelector.DoNotFilterFirst = True
                If ASCMAIN1.CodeSelector.SQL <> "" Then
                    ASCMAIN1.CodeSelector.MultipleSelections = False
                    ASCMAIN1.CodeSelector.ForceFilterFirst = False

                    Dim F As New ASFCODE1
                    F.ShowDialog()
                    F.Dispose()

                    If ASCMAIN1.CodeSelector.Selections <> 0 Then
                        Dim VOUCHER_NO As String = ASCMAIN1.CodeSelector.SelectedCode

                        'Dim VOUCHER_NO As String = "0009382495"

                        Dim RECEIPT_NOs As New List(Of String)
                            ASCMAIN1.sql = $"Select * from APTINVH5 where VOUCHER_NO = '{VOUCHER_NO}'"
                        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "VOUCHER_DLNO")
                            Dim RECEIPT_NO As String = row.Item("RECEIPT_NO")
                            If RECEIPT_NOs.Contains(RECEIPT_NO) Then
                                ' skip
                            Else
                                RECEIPT_NOs.Add(RECEIPT_NO)
                                ASCMAIN1.sql = "Select ICTIREC1.*, X.EXT_PO_COST, X.QTY_REC_NOT_INV" & vbCrLf _
                                & " from ICTIREC1" & vbCrLf _
                                & ", (Select ICTIREC2.RECEIPT_NO, SUM (ICTIREC2.QTY_REC * ICTIREC2.PO_COST) EXT_PO_COST" & vbCrLf _
                                & ", Sum (NVL(ICTIREC2.QTY_REC,0) - NVL(ICTIREC2.QTY_INV,0)) QTY_REC_NOT_INV" & vbCrLf _
                                & " from ICTIREC2" & vbCrLf _
                                & " where ICTIREC2.RECEIPT_NO = '" & RECEIPT_NO & "' group by ICTIREC2.RECEIPT_NO) X" & vbCrLf _
                                & " where X.RECEIPT_NO = ICTIREC1.RECEIPT_NO"
                                Fill_Records("ICTIREC1", "", False, ASCMAIN1.sql)
                                Fill_Records("ICTIREC2", RECEIPT_NO, False)

                                For Each grow As UltraWinGrid.UltraGridRow In grdICTIREC1.Rows
                                    If grow.IsDataRow Then

                                        If grow.Cells("RECEIPT_NO").Value & "" = RECEIPT_NO Then
                                            grow.Activate()
                                            'grow.Expanded = True
                                            grdICTIREC1.DisplayLayout.RowScrollRegions(0).FirstRow = grow
                                            If grow.IsDataRow Then
                                                ProcessDoubleClickedRow()
                                            End If
                                            Exit For
                                        End If
                                    End If
                                Next

                                If e.Tool.Key = "Retrieve Accrued POs from Voucher (w/Prev Inv Cost)" Then
                                    For Each rowAPTINVH5 As DataRow In dst.Tables("APTINVH5").Select($"RECEIPT_NO = '{RECEIPT_NO}'")
                                        Dim RECEIPT_LNO As Integer = Val(rowAPTINVH5.Item("RECEIPT_LNO"))
                                        ASCMAIN1.sql = $"Select MIN(APTINVH5.VOUCHER_NO) VOUCHER_NO from APTINVH5,APTINVH1 where APTINVH1.VOUCHER_NO = APTINVH5.VOUCHER_NO and APTINVH1.INV_STATUS <> 'D' and APTINVH5.RECEIPT_NO = '{RECEIPT_NO}' and APTINVH5.RECEIPT_LNO = {CStr(RECEIPT_LNO)} and APTINVH5.VOUCHER_NO <> '{VOUCHER_NO}'"
                                        ASCMAIN1.sql = $"Select * from APTINVH5,({ASCMAIN1.sql}) X where X.VOUCHER_NO = APTINVH5.VOUCHER_NO and APTINVH5.RECEIPT_NO = '{RECEIPT_NO}' and APTINVH5.RECEIPT_LNO = {CStr(RECEIPT_LNO)}"
                                        Dim rowx As DataRow = ASCDATA1.GetDataRow
                                        If rowx IsNot Nothing Then
                                            rowAPTINVH5.Item("INV_COST") = rowx.Item("INV_COST")
                                        End If
                                    Next
                                End If
                            End If

                        Next

                        Sort_grdColumns(grdAPTINVH5_SUM, "RECEIPT_NO")

                        Exit Sub

                    End If
                End If

            Case "Retrieve Closed Accrued PO"

                Dim SQLW As String = "VEND_CODE = '" & Absx1.txtFor("VEND_CODE").Text & "' and ACCRUAL_STATUS = '1'"
                                                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("RECEIPT_NO", , SQLW)

                If ASCMAIN1.CodeSelector.SQL <> "" Then
                    ASCMAIN1.CodeSelector.MultipleSelections = False

                    Dim F As New ASFCODE1
                    F.ShowDialog()
                    F.Dispose()
                    If ASCMAIN1.CodeSelector.Selections <> 0 Then
                        Dim RECEIPT_NO As String = ASCMAIN1.CodeSelector.SelectedCode

                        If dst.Tables("ICTIREC1").Rows.Find(RECEIPT_NO) Is Nothing Then
                            ASCMAIN1.sql = "Select ICTIREC1.*, X.EXT_PO_COST, X.QTY_REC_NOT_INV" & vbCrLf _
                                & " from ICTIREC1" & vbCrLf _
                                & ", (Select ICTIREC2.RECEIPT_NO, SUM (ICTIREC2.QTY_REC * ICTIREC2.PO_COST) EXT_PO_COST" & vbCrLf _
                                & ", Sum (NVL(ICTIREC2.QTY_REC,0) - NVL(ICTIREC2.QTY_INV,0)) QTY_REC_NOT_INV" & vbCrLf _
                                & " from ICTIREC2" & vbCrLf _
                                & " where ICTIREC2.RECEIPT_NO = '" & RECEIPT_NO & "' group by ICTIREC2.RECEIPT_NO) X" & vbCrLf _
                                & " where X.RECEIPT_NO = ICTIREC1.RECEIPT_NO"
                            Fill_Records("ICTIREC1", "", False, ASCMAIN1.sql)
                            Fill_Records("ICTIREC2", RECEIPT_NO, False)

                            For Each grow As UltraWinGrid.UltraGridRow In grdICTIREC1.Rows
                                If grow.IsDataRow Then

                                    If grow.Cells("RECEIPT_NO").Value & "" = RECEIPT_NO Then
                                        grow.Activate()
                                        grow.Expanded = True
                                        grdICTIREC1.DisplayLayout.RowScrollRegions(0).FirstRow = grow
                                        Exit For
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If

            Case "Re-Submit Submitted Email"
                If optShow.Value = "P" Then
                    '       Dim INV_COST As Double = Val(grd.ActiveCell.Value & "")
                    Dim SUBMIT_CTL_NO As String = grdAPTSUBM1.ActiveRow.Cells("SUBMIT_CTL_NO").Value & ""
                    Dim SUBMIT_NO_ORIG As String = grdAPTSUBM1.ActiveRow.Cells("SUBMIT_NO_ORIG").Value & ""

                    If MsgBox("Do you really want to Re-Submit Submitteed Invoice for Selected Ctl No " & SUBMIT_CTL_NO & "?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                        Dim row As DataRow = dst.Tables("APTSUBM1").Rows.Find(New Object() {SUBMIT_CTL_NO})
                        Dim rowAPTSUBM1 As DataRow = dst.Tables("APTSUBM1").NewRow
                        Dim SUBMIT_CTL_NO_NEW As String = ASCMAIN1.Next_Control_No("APTSUBM1.SUBMIT_CTL_NO")
                        With rowAPTSUBM1
                            .Item("SUBMIT_CTL_NO") = SUBMIT_CTL_NO_NEW
                            .Item("SUBMIT_EMAIL_FROM") = row.Item("SUBMIT_EMAIL_FROM")
                            .Item("SUBMIT_SUBJECT") = row.Item("SUBMIT_SUBJECT")
                            .Item("SUBMIT_DATE_RECEIVED") = row.Item("SUBMIT_DATE_RECEIVED")
                            .Item("SUBMIT_STATUS") = "U"
                            .Item("VOUCHER_NO") = ""
                            .Item("INV_NUM") = ""
                            '   .Item("INV_DATE") = Now
                            ' .Item("INV_AMT") = ""
                            .Item("INIT_DATE") = row.Item("INIT_DATE")
                            .Item("INIT_OPER") = row.Item("INIT_OPER")
                            .Item("LAST_DATE") = DATETIME_STAMP
                            .Item("LAST_OPER") = ASCMAIN1.USER_ID
                            If SUBMIT_NO_ORIG <> "" Then
                                .Item("SUBMIT_NO_ORIG") = SUBMIT_NO_ORIG
                            Else
                                .Item("SUBMIT_NO_ORIG") = SUBMIT_CTL_NO
                            End If
                            .Item("SUBMIT_NO_COPY_FROM") = SUBMIT_CTL_NO
                        End With
                        dst.Tables("APTSUBM1").Rows.Add(rowAPTSUBM1)
                        Update_Record_TDA("APTSUBM1")
                        optShow.Value = "U"
                    End If

                End If

            Case "Delete Submitted Email"
                If optShow.Value = "U" Then
                    Dim SUBMIT_CTL_NO As String = grdAPTSUBM1.ActiveRow.Cells("SUBMIT_CTL_NO").Value

                    If MsgBox("Do you really want to Delete Submitteed Invoice for Selected Ctl No " & SUBMIT_CTL_NO & "?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                        Dim row As DataRow = dst.Tables("APTSUBM1").Rows.Find(New Object() {SUBMIT_CTL_NO})
                        For Each grow As UltraWinGrid.UltraGridRow In grdAPTSUBM1.Selected.Rows
                            grow.Cells("SUBMIT_STATUS").Value = "D"
                            grow.Cells("LAST_DATE").Value = DATETIME_STAMP
                            grow.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
                            grow.Update()
                        Next
                        Update_Record_TDA("APTSUBM1")
                        optShow.Value = "U"
                        OPT_SHOW_REFRESH()
                        '   Update_APTSUBM()
                    End If

                End If

            Case "Refresh"
                Show_Batch()
                '   Fill_Records("APTINVR1")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI")

            Case "PO Receipts Inquiry"
                Dim RECEIPT_NO As String = grd.ActiveRow.Cells("RECEIPT_NO").Text
                Context_Launch("View", RECEIPT_NO, e.Tool.Key, "ICFIRECI")

            Case "Split Accrual"
                Dim COST_ACC As Decimal = Val(grdAPTACRC1.ActiveRow.Cells("COST_ACC").Value & "")
                Dim COST_ACT As Decimal = Val(grdAPTACRC1.ActiveRow.Cells("COST_ACT").Value & "")
                Dim CTL_STATUS As String = grdAPTACRC1.ActiveRow.Cells("CTL_STATUS").Value

                If COST_ACC - COST_ACT <= 0 Or CTL_STATUS <> "0" Then
                    MsgBox("Cannot Split this Accrual", MsgBoxStyle.OkOnly, "No Open Accrued Amount to Split")
                    Exit Sub
                End If

                Using frm As New ASFMSGBF

                    Dim SPLIT_AMT As Decimal = frm.Get_numdec_from_User("Amount to Split", "Enter the Amount to Split", COST_ACC - COST_ACT, 0, 0)
                    If frm.user_option <> -1 Then
                        If SPLIT_AMT <> 0 Then
                            Dim CTL_NO As String = grdAPTACRC1.ActiveRow.Cells("CTL_NO").Value
                            Dim rowAPTACRC1 As DataRow = dst.Tables("APTACRC1").Rows.Find(CTL_NO)
                            rowAPTACRC1.Item("COST_ACC") = COST_ACC - SPLIT_AMT
                            Dim rowAPTACRC1_split As DataRow = dst.Tables("APTACRC1").NewRow
                            CTL_NO = ASCMAIN1.Next_Control_No("APTACRC1.CTL_NO")
                            rowAPTACRC1_split.ItemArray = rowAPTACRC1.ItemArray
                            rowAPTACRC1_split.Item("CTL_NO") = CTL_NO
                            rowAPTACRC1_split.Item("COST_ACC") = SPLIT_AMT
                            rowAPTACRC1_split.Item("COST_ACT") = 0
                            dst.Tables("APTACRC1").Rows.Add(rowAPTACRC1_split)
                        End If
                    End If
                End Using


            Case "Clone Line", "Clone Line (w/Adj)"
                If MsgBox("OK to Clone this Invoice Detail Line?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then

                    Dim adjEntered As String = ""
                    Dim adj As Decimal = 0
                    If e.Tool.Key = "Clone Line (w/Adj)" Then
                        adjEntered = InputBox("Enter Adjustment to Price", 0)
                        adj = Round(Val(adjEntered), 4)
                        If Val(adj) = 0 And adjEntered <> "0" Then Exit Sub
                    End If

                    Dim RECEIPT_NO As String = grdAPTINVH5.ActiveRow.Cells("RECEIPT_NO").Value & ""
                    Dim RECEIPT_LNO As Int32 = Val(grdAPTINVH5.ActiveRow.Cells("RECEIPT_LNO").Value & "")
                    Dim PO_COST As Decimal = Val(grdAPTINVH5.ActiveRow.Cells("PO_COST").Value & "")
                    Dim INV_COST As Decimal = Val(grdAPTINVH5.ActiveRow.Cells("INV_COST").Value & "")

                    'Dim rowICTIREC2 As DataRow = LookUp("ICTIREC2", New Object() {RECEIPT_NO, RECEIPT_LNO})
                    Dim QTY_REC As Int32 = Val(grdAPTINVH5.ActiveRow.Cells("QTY_REC").Value & "")
                    Dim VOUCHER_NO As String = grdAPTINVH5.ActiveRow.Cells("VOUCHER_NO").Value & ""
                    Dim VOUCHER_DLNO As Int32 = Val(grdAPTINVH5.ActiveRow.Cells("VOUCHER_DLNO").Value & "")
                    Dim rowAPTINVH5 As DataRow = dst.Tables("APTINVH5").Rows.Find(New Object() {VOUCHER_NO, VOUCHER_DLNO})
                    'rowAPTINVH5.Item("INV_QTY") = -1 * QTY_REC
                    grdAPTINVH5.ActiveRow.Cells("INV_QTY").Value = -1 * QTY_REC
                    grdAPTINVH5.ActiveRow.Update()

                    Dim VOUCHER_DLNO_max As Int32 = Val(dst.Tables("APTINVH5").Compute("MAX(VOUCHER_DLNO)", $"VOUCHER_NO = '{VOUCHER_NO}'") & "")
                    Dim rowAPTINVH5_new As DataRow = dst.Tables("APTINVH5").NewRow
                    rowAPTINVH5_new.ItemArray = rowAPTINVH5.ItemArray
                    'rowAPTINVH5_new.Item("INV_QTY") = 0

                    rowAPTINVH5_new.Item("INV_QTY") = QTY_REC
                    If e.Tool.Key = "Clone Line (w/Adj)" Then
                        rowAPTINVH5_new.Item("INV_COST") = INV_COST + Val(adj)
                    Else
                        rowAPTINVH5_new.Item("INV_COST") = INV_COST
                    End If

                    rowAPTINVH5_new.Item("VOUCHER_DLNO") = VOUCHER_DLNO_max + 1
                    dst.Tables("APTINVH5").Rows.Add(rowAPTINVH5_new)
                End If

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

            Case "VEND_CODE"

                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    If Not chkQuickEntry.Checked Then
                        'Click_Command("New", e)
                    End If
                End If

            Case "INV_NUM"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    If Not chkQuickEntry.Checked Then
                        Click_Command("New", e)
                    End If
                End If

            Case "VOUCHER_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    optShow.Value = "U"
                    OPT_SHOW_REFRESH()
                    UltraTabControl1.Tabs("A/P Invoices").Selected = True
                    Click_Command("Edit", e)
                End If

            Case "ACCRUAL_CODE"

        End Select

    End Sub

    Overrides Sub num_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.num_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

            Case "QE_INV_AMT"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    Click_Command("New", e)
                End If


        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "TERM_CODE"
                Calculate_INV_DUE_DATE()
            Case "CURR_CODE"
                Load_CURR_EXCH_RATE()
            Case "ACCRUAL_CODE"
                If ScreenMode And EntryMode = "N" Then
                    Dim ACCRUAL_CODE As String = Absx1.txtFor("ACCRUAL_CODE").Text
                    If ACCRUAL_CODE <> "" Then
                        If Absx1.txtFor("VEND_CODE_ACC").Text = "" Then
                            Dim row As DataRow = LookUp("APTACRM1", ACCRUAL_CODE)
                            If row Is Nothing Then
                                Absx1.txtFor("ACCRUAL_CODE").Text = ""
                            Else
                                Absx1.txtFor("VEND_CODE_ACC").Text = row.Item("VEND_CODE") & ""
                            End If

                        End If
                    End If
                End If
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "VEND_CODE"

        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "VOUCHER_NO"
                Click_Command("Edit")
            Case "VEND_ALT_CODE"
                Load_Alternate_Payment_Address()
            Case "OPS_YYYYPP_ACCRUE"

        End Select
    End Sub

#End Region

    Private Sub tabMain_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged

        If EntryMode = "" Then
            Exit Sub
        End If
        With UltraExplorerBar1
            ' .Groups("Entry Options").Visible = False
            .Groups("Distribution Options").Visible = False
            .Groups("Generate Accrual").Visible = False

            Select Case tabMain.ActiveTab.Key
                Case "Vendor Information"
                Case "GL Distribution"
                    .Groups("Distribution Options").Visible = True And Not InquiryMode And Not ApprovalMode
                Case "Header Data"
                    '.Groups("Payment Options").Visible = True and not inquirymode AND NOT ApprovalMode
                Case "Details"

                Case "Other Accruals"
                    .Groups("Generate Accrual").Visible = Not InquiryMode
            End Select
        End With

    End Sub

    Private Sub optINV_REMIT_TO_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optINV_REMIT_TO.ValueChanged
        lblVEND_CODE_AP.Visible = (optINV_REMIT_TO.Value = "P")
        txtVEND_CODE_AP.Visible = (optINV_REMIT_TO.Value = "P")
        Absx1.txtFor("VEND_ALT_CODE").Visible = (optINV_REMIT_TO.Value = "A" Or optINV_REMIT_TO.Value = "N")
        Absx1.txtFor("VEND_ALT_NAME").ReadOnly = (optINV_REMIT_TO.Value <> "N") Or InquiryMode Or ApprovalMode
        For Each COLUMN_NAME As String In New String() {"VEND_ALT_NAME", "VEND_ALT_ADDR1", "VEND_ALT_ADDR2", "VEND_ALT_ADDR3", "VEND_ALT_CITY", "VEND_ALT_STATE", "VEND_ALT_ZIP_CODE", "VEND_ALT_PHONE", "VEND_ALT_EXT", "VEND_ALT_FAX", "VEND_ALT_COUNTRY", "VEND_ALT_CONTACT", "VEND_ALT_EMAIL"}
            If COLUMN_NAME = "VEND_ALT_PHONE" Or COLUMN_NAME = "VEND_ALT_FAX" Then
                Absx1.medFor(COLUMN_NAME).ReadOnly = (optINV_REMIT_TO.Value <> "N") Or InquiryMode Or ApprovalMode
            Else
                Absx1.txtFor(COLUMN_NAME).ReadOnly = (optINV_REMIT_TO.Value <> "N") Or InquiryMode Or ApprovalMode
            End If
        Next

        'If optINV_REMIT_TO.Value <> "N" Then
        Set_Payment_Address()
        'End If
    End Sub

    Private Sub optINV_STATUS_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optINV_STATUS.ValueChanged
        tabHeader.Tabs("Pymt Info").Enabled = (optINV_STATUS.Value = "P")
        If optINV_STATUS.Value = "P" AndAlso (EntryMode = "E" Or EntryMode = "N") Then
            tabHeader.Tabs("Pymt Info").Selected = True
        Else
            If tabHeader.SelectedTab IsNot Nothing AndAlso tabHeader.SelectedTab.Key = "Pymt Info" Then
                tabHeader.Tabs("Codes").Selected = True
            End If
        End If
    End Sub

    Sub Load_ICTIREC1()
        ASCMAIN1.Progress("Now Loading Receipts")
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()

        Dim sql_where As String = " and ICTIREC1.RECEIPT_NO in " _
        & " (Select Distinct RECEIPT_NO from APTINVH5 where VOUCHER_NO = '" & HFs("VOUCHER_NO") & "'"
        If InquiryMode Or ApprovalMode Then
            sql_where &= ")"
        Else
            sql_where &= " union " _
            & "  Select RECEIPT_NO from ICTIREC1 where ICTIREC1.ACCRUAL_STATUS = '0' AND VEND_CODE = '" & HFs("VEND_CODE") & "')"
        End If

        ASCDATA1.ExecuteSQL("Truncate Table " & ICTIREC1)
        If Not batch_update Then
            ASCDATA1.ExecuteSQL("Insert into " & ICTIREC1 & vbCrLf _
                                & " Select ICTIREC1.* from ICTIREC1" & vbCrLf _
                                & " where ICTIREC1.VEND_CODE = '" & HFs("VEND_CODE") & "'" & sql_where)
        End If
        Fill_Records("ICTIREC1")

        Sort_grdColumns(grdICTIREC1, "RECEIPT_NO".ToLower)

        Setup_ICTIREC1()

        For Each rowAPTINVH5 As DataRow In dst.Tables("APTINVH5").Select("")
            Dim QTY_REC_NOT_INV As Int64 = Val(rowAPTINVH5.Item("QTY_REC_NOT_INV") & "")
            Dim QTY_REC As Int64 = Val(rowAPTINVH5.Item("QTY_REC") & "")
            Dim QTY_INV As Int64 = Val(rowAPTINVH5.Item("QTY_INV") & "")

            ' these 2 lines cause the R^I to be doubled
            'QTY_REC_NOT_INV = QTY_REC_NOT_INV + QTY_INV
            'rowAPTINVH5.Item("QTY_REC_NOT_INV") = QTY_REC_NOT_INV

            Dim RECEIPT_NO As String = rowAPTINVH5.Item("RECEIPT_NO")
            Dim rowICTIREC1 As DataRow = dst.Tables("ICTIREC1").Rows.Find(RECEIPT_NO)
            rowICTIREC1.Item("QTY_REC_NOT_INV") = Val(rowICTIREC1.Item("QTY_REC_NOT_INV") & "") + QTY_REC_NOT_INV

            Dim rowAPTINVH5_SUM As DataRow
            rowAPTINVH5_SUM = dst.Tables("APTINVH5_SUM").Rows.Find(RECEIPT_NO)
            If rowAPTINVH5_SUM Is Nothing Then
                rowAPTINVH5_SUM = dst.Tables("APTINVH5_SUM").NewRow
                rowAPTINVH5_SUM.Item("RECEIPT_NO") = RECEIPT_NO
                rowAPTINVH5_SUM.Item("RECEIPT_DATE") = rowICTIREC1.Item("RECEIPT_DATE")
                rowAPTINVH5_SUM.Item("PO_ORDER_NO") = rowICTIREC1.Item("PO_ORDER_NO")
                dst.Tables("APTINVH5_SUM").Rows.Add(rowAPTINVH5_SUM)
            End If
        Next

        If EntryMode <> "" Then
            Setup_Other_Accruals()
            tabMain.Tabs("Other Accruals").Enabled = True
            'If ASCMAIN1.CLIENT = "INT" Then
            'Else
            tabMain.Tabs("Other Accruals").Visible = True
            'End If
        End If


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Setup_Other_Accruals()
        If EntryMode = "" Then
            Exit Sub
        End If
        'If grdICTIREC1.Rows.Count <> 0 Then
        '    grdICTIREC1.ActiveRow = grdICTIREC1.Rows(0)
        '    tabMain.Tabs("PO Receipts").Enabled = True
        'Else
        '    tabMain.Tabs("PO Receipts").Enabled = False
        'End If
    End Sub
    Sub Setup_ICTIREC1()
        If EntryMode = "" Then
            Exit Sub
        End If
        If grdICTIREC1.Rows.Count <> 0 Then
            grdICTIREC1.ActiveRow = grdICTIREC1.Rows(0)
            tabMain.Tabs("PO Receipts").Enabled = True
        Else
            tabMain.Tabs("PO Receipts").Enabled = False
        End If
    End Sub

    Private Sub grdICTIREC1_BeforeRowExpanded(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTIREC1.BeforeRowExpanded
        If grdICTIREC1.ActiveRow Is Nothing Then
        Else
            If grdICTIREC1.ActiveRow.IsGroupByRow Then
            Else
                Me.Cursor = Cursors.WaitCursor
                Dim RECEIPT_NO As String = e.Row.Cells("RECEIPT_NO").Text '  grdICTIREC1.ActiveRow.Cells("RECEIPT_NO").Text
                Fill_Records("ICTIREC2", New String() {RECEIPT_NO})
                grdICTIREC1.DisplayLayout.Bands("ICTIREC1_ICTIREC2").SummaryFooterCaption = "Totals for Receipt " & RECEIPT_NO
                Me.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Private Sub grdICTIREC1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTIREC1.InitializeRow
        If e.Row.Band.Key = "ICTIREC1" Then
            If e.Row.Cells("ACCRUAL_STATUS").Text = "2" Then
                e.Row.Appearance.BackColor = Drawing.Color.Yellow
            ElseIf e.Row.Cells("ACCRUAL_STATUS").Text = "1" Then
                e.Row.Appearance.ForeColor = Drawing.Color.Purple
            Else
                e.Row.Appearance.BackColor = Drawing.Color.FromArgb(0, 0, 0, 0)
            End If
        End If

    End Sub

    Private Sub grdAPTINVH5_SUM_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTINVH5_SUM.AfterRowActivate

        Dim RECEIPT_NO As String = grdAPTINVH5_SUM.ActiveRow.Cells("RECEIPT_NO").Text
        Setup_grdAPTINVH5(RECEIPT_NO)
        Dim rowICTIREC1 As DataRow = dst.Tables("ICTIREC1").Rows.Find(RECEIPT_NO)
        Dim i As Integer = dst.Tables("ICTIREC1").Rows.IndexOf(rowICTIREC1)
        grdICTIREC1.Rows.GetRowWithListIndex(i).Activate()

    End Sub

    Sub Setup_grdAPTINVH5(ByVal RECEIPT_NO As String)

        Dim dvw As DataView = DirectCast(grdAPTINVH5.DataSource, DataTable).DefaultView
        dvw.RowFilter = "RECEIPT_NO = '" & RECEIPT_NO & "'"
        If discrepancies_only Then
            dvw.RowFilter &= "AND ISNULL(QTY_REC,0) <> ISNULL(INV_QTY,0)"
        End If
        Sort_grdColumns(grdAPTINVH5, "VOUCHER_DLNO")

        grdAPTINVH5.DisplayLayout.Bands("APTINVH5").SummaryFooterCaption = "Totals for Receipt " & RECEIPT_NO
        grdAPTINVH5.Visible = True

    End Sub

    Private Sub grdAPTINVH5_SUM_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTINVH5_SUM.AfterRowsDeleted
        grdAPTINVH5.Visible = False

        Dim RECEIPT_NOs As List(Of String) = DirectCast(grdAPTINVH5_SUM.Tag, List(Of String))

        For Each RECEIPT_NO As String In RECEIPT_NOs
            Dim rowICTIREC1 As DataRow = dst.Tables("ICTIREC1").Rows.Find(RECEIPT_NO)
            rowICTIREC1.Item("ACCRUAL_STATUS") = "0"
        Next
        'grdICTIREC1.ActiveRow.Cells("ACCRUAL_STATUS").Value = "0"
        'grdICTIREC1.UpdateData()

        Calc_DIST_PO()
    End Sub

    Private Sub grdAPTINVH5_SUM_BeforeRowsDeleted(sender As Object, e As BeforeRowsDeletedEventArgs) Handles grdAPTINVH5_SUM.BeforeRowsDeleted

        Dim RECEIPT_NOs As New List(Of String)
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            RECEIPT_NOs.Add(grow.Cells("RECEIPT_NO").Value)
        Next

        grdAPTINVH5_SUM.Tag = RECEIPT_NOs

    End Sub

    'Sub Display_Totals()

    '    ' this section is unnec

    '    Dim INV_AMT As Decimal = Val(Absx1.numFor("INV_AMT").Value & "")
    '    Dim DIST_GL As Decimal = Val(dst.Tables("APTINVH2").Compute("SUM(INV_LINE_AMT)", "INV_LTYP is Null or INV_LTYP = 'A'") & "")
    '    Dim DIST_PO As Decimal = Val(dst.Tables("APTINVH5_SUM").Compute("SUM(AMT_INV)", "") & "")

    '    Dim DIST_OTHER As Decimal = Val(dst.Tables("APTINVH7").Compute("SUM(TOTAL_INV)", "") & "")
    '    Dim DIST_OOBAL As Decimal = INV_AMT - DIST_GL - DIST_PO - DIST_OTHER

    '    Absx1.numFor("DIST_GL").Value = DIST_GL
    '    Absx1.numFor("DIST_PO").Value = DIST_PO
    '    Absx1.numFor("DIST_OTHER").Value = DIST_OTHER
    '    Absx1.numFor("DIST_OOBAL").Value = DIST_OOBAL
    'End Sub

    Sub Set_Payment_Address()
        If EntryMode = "" Then
            Exit Sub
        End If
        If optINV_REMIT_TO.Value & "" = "" Then
            Exit Sub
        End If
        Dim rowAPTVEND2 As DataRow

        With dst.Tables("APTVEND2")
            .Rows.Clear()
            rowAPTVEND2 = .NewRow
            rowAPTVEND2.Item("VEND_CODE") = HFs("VEND_CODE")
            'rowAPTVEND2.Item("VEND_CODE") = rowAPTINVH1.Item("VEND_CODE")

            Select Case optINV_REMIT_TO.Value
                Case "V"
                    rowAPTVEND2.Item("VEND_ALT_CODE") = "VENDOR"
                    Absx1.txtFor("VEND_ALT_CODE").Text = ""
                    For i As Integer = 0 To .Columns.Count - 1
                        Dim COLUMN_NAME As String = .Columns(i).ColumnName
                        If COLUMN_NAME <> "VEND_ALT_CODE" Then
                            COLUMN_NAME = Replace(COLUMN_NAME, "_ALT_", "_")
                            rowAPTVEND2.Item(i) = rowAPTVEND1.Item(COLUMN_NAME)
                        End If
                    Next

                Case "P"
                    rowAPTVEND2.Item("VEND_ALT_CODE") = "VENDOR"
                    Absx1.txtFor("VEND_ALT_CODE").Text = ""
                    If txtVEND_CODE_AP.Text = "" Then
                        txtVEND_CODE_AP.Text = rowAPTVEND2.Item("VEND_CODE")
                    End If
                    LookUp("APTVEND1", txtVEND_CODE_AP.Text, True)
                    For i As Integer = 0 To .Columns.Count - 1
                        Dim COLUMN_NAME As String = .Columns(i).ColumnName
                        If COLUMN_NAME <> "VEND_ALT_CODE" Then
                            COLUMN_NAME = Replace(COLUMN_NAME, "_ALT_", "_")
                            rowAPTVEND2.Item(i) = cdr.Item(COLUMN_NAME)
                        End If
                    Next

                Case "A"
                    Absx1.txtFor("VEND_ALT_CODE").ReadOnly = InquiryMode Or ApprovalMode
                    rowAPTVEND2.Item("VEND_ALT_CODE") = rowAPTINVH1.Item("VEND_ALT_CODE")

                    LookUp("APTVEND2", New String() {HFs("VEND_CODE"), rowAPTVEND2.Item("VEND_ALT_CODE") & ""}, True)
                    'LookUp("APTVEND2", New String() {rowAPTVEND2.Item("VEND_CODE") & "", rowAPTVEND2.Item("VEND_ALT_CODE") & ""}, True)
                    If cdr IsNot Nothing Then
                        rowAPTVEND2.ItemArray = cdr.ItemArray
                    End If

                Case "N"
                    Absx1.txtFor("VEND_ALT_CODE").ReadOnly = InquiryMode Or ApprovalMode
                    rowAPTVEND2.Item("VEND_ALT_CODE") = "NEW"
                    Absx1.txtFor("VEND_ALT_CODE").Text = "NEW"
            End Select
            .Rows.Add(rowAPTVEND2)
        End With

    End Sub

    Private Sub grdAPTINVH2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTINVH2.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "ACCT_CODE"
                Dim ACCT_CODE As String = e.Cell.Value & ""

                grdCodeDesc(grdAPTINVH2, "GLTACCT1", "ACCT_CODE", "ACCT_DESC")
                For i As Integer = 2 To 4
                    If e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Text = "" Then
                        e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Value = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i))
                    End If
                Next
        End Select
    End Sub

    Private Sub grdAPTINVH2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTINVH2.AfterExitEditMode
        With grdAPTINVH2
            Select Case .ActiveCell.Column.Key
                Case "ACCT_CODE"
                    Dim ACCT_CODE As String = .ActiveCell.Text
                    If ACCT_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(ACCT_CODE, .ActiveCell.Column.Key)
                    End If
            End Select
        End With
    End Sub

    Private Sub grdAPTINVH2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTINVH2.AfterRowActivate
        With grdAPTINVH2
            If .ActiveRow.IsAddRow Then
                .DisplayLayout.Bands(0).Columns("ACCT_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .ActiveCell = grdAPTINVH2.ActiveRow.Cells("ACCT_CODE")
                .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                '.DisplayLayout.Bands(0).Columns("ACCT_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                ' why cant we edit the acct code?
            End If
        End With
    End Sub

    Private Sub grdAPTINVH2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTINVH2.AfterRowsDeleted
        Calc_DIST_GL()
    End Sub

    Private Sub grdAPTINVH2_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdAPTINVH2.AfterRowUpdate
        Calc_DIST_GL()
    End Sub

    Private Sub grdAPTINVH2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdAPTINVH2.BeforeRowUpdate
        With grdAPTINVH2
            If e.Row.Cells("ACCT_CODE").Text = "" Then
                e.Cancel = True
            Else
                LookUp("GLTACCT1", e.Row.Cells("ACCT_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Acct Code (" & e.Row.Cells("ACCT_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                Else
                    If cdr.Item("ACCT_STATUS") & "" <> "A" Then
                        MsgBox("Acct Code " & e.Row.Cells("ACCT_CODE").Text & " is not Active", MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If
                    If cdr.Item("ACCT_SUB_CTL") & "" = "1" Then
                        MsgBox("Acct Code " & e.Row.Cells("ACCT_CODE").Text & " is a Control Account - no Manual J/E permitted", MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If
                End If

                Dim DIST_APP_CODE As String = "AP"
                If LookUp("GLTDSTR1", DIST_APP_CODE) IsNot Nothing AndAlso cdr.Item("DIST_APP_STATUS") & "" = "A" Then
                    If LookUp("GLTDSTR2", New String() {DIST_APP_CODE, e.Row.Cells("ACCT_CODE").Text}) Is Nothing Then
                        MsgBox("Acct Code " & e.Row.Cells("ACCT_CODE").Text & " is not permitted for Posting in this Application (" & DIST_APP_CODE & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If
                End If
            End If

            If e.Row.Cells("DIST_CODE").Text <> "" Then
                LookUp("GLTDIST1", e.Row.Cells("DIST_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Dist Code (" & e.Row.Cells("DIST_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            Dim COLUMN_NAME As String
            For i As Integer = 2 To 4
                COLUMN_NAME = "SEG" & CStr(i) & "_CODE"
                If Not e.Row.Cells(COLUMN_NAME).Column.Hidden Then
                    If e.Row.Cells(COLUMN_NAME).Text = "" Then
                        e.Cancel = True
                    Else
                        LookUp("GLTSEGM1", New String() {CStr(i), e.Row.Cells(COLUMN_NAME).Text})
                        If cdr Is Nothing Then
                            MsgBox("Invalid Value entered for " & e.Row.Cells(COLUMN_NAME).Column.Header.Caption & " (" & e.Row.Cells(COLUMN_NAME).Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                            e.Cancel = True
                        ElseIf cdr.Item("ACCT_SEG_STATUS") & "" <> "A" Then
                            MsgBox(e.Row.Cells(COLUMN_NAME).Column.Header.Caption & " " & e.Row.Cells(COLUMN_NAME).Text & " is Inactive", MsgBoxStyle.OkOnly, "Cannot Update Row")
                            e.Cancel = True
                        ElseIf cdr.Item("ACCT_SEG_NO_GL") & "" = "1" Then
                            MsgBox(e.Row.Cells(COLUMN_NAME).Column.Header.Caption & " " & e.Row.Cells(COLUMN_NAME).Text & " is Not set up to allow Posting", MsgBoxStyle.OkOnly, "Cannot Update Row")
                            e.Cancel = True
                        End If
                    End If
                End If
            Next

            If Not e.Cancel Then
                If e.Row.Cells("VOUCHER_NO").Text = "" Then
                    .ActiveRow.Cells("VOUCHER_NO").Value = Absx1.CtlFor("VOUCHER_NO").Text
                    .ActiveRow.Cells("VOUCHER_LNO").Value = Val(dst.Tables("APTINVH2").Compute("Max(VOUCHER_LNO)", "") & "") + 1
                End If
            End If
        End With

    End Sub

    Private Sub grdAPTINVH2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTINVH2.ClickCellButton
        Dim sql_where As String = ""
        If LookUp("GLTDSTR1", "AP") IsNot Nothing AndAlso cdr.Item("DIST_APP_STATUS") & "" = "A" Then
            sql_where = "ACCT_CODE in (Select ACCT_CODE from GLTDSTR2 where DIST_APP_CODE = 'AP')"
        End If
        grdClickCellButton(grdAPTINVH2, sql_where, sql_where <> "")
    End Sub

    Private Sub numDIST_GL_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numDIST_GL.ValueChanged
        Calc_Totals()
    End Sub

    Private Sub numDIST_OTHER_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numDIST_OTHER.ValueChanged
        Calc_Totals()
    End Sub

    Private Sub numDIST_PO_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numDIST_PO.ValueChanged
        Calc_Totals()
    End Sub

    Sub Calc_Totals()
        numINV_AMT.Value = Val(numINV_AMT_VEND.Value & "") + Val(Absx1.numFor("INV_ADJUSTMENTS").Value & "") '  - Val(Absx1.numFor("INV_ALLOWANCES").Value & "")

        numDIST_OOBAL.Value = Val(numINV_AMT.Value & "") _
                            - Val(numDIST_GL.Value & "") _
                            - Val(numDIST_PO.Value & "") _
                            + Val(Absx1.numFor("INV_ALLOWANCES").Value & "") _
                            - Val(numDIST_OTHER.Value & "")

        '                            - Val(Absx1.numFor("INV_ADJUSTMENTS").Value & "") _

        If Abs(Val(numINV_AMT.Value & "") - Val(numINV_AMT_VEND.Value & "")) > 0.01 Then
            numINV_AMT.Appearance.ForeColor = Drawing.Color.Red
            Absx1.numFor("INV_ADJUSTMENTS").Appearance.ForeColor = Drawing.Color.Red
        Else
            numINV_AMT.Appearance.ForeColor = Drawing.Color.Empty
            Absx1.numFor("INV_ADJUSTMENTS").Appearance.ForeColor = Drawing.Color.Empty
        End If

    End Sub

    Private Sub numINV_AMT_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles numINV_AMT.Leave
        Automatic_Distribution()
    End Sub

    Private Sub numINV_AMT_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numINV_AMT.ValueChanged
        Calc_Totals()
    End Sub

    Sub Calc_DIST_GL()
        Dim DIST_GL As Decimal = Val(dst.Tables("APTINVH2").Compute("SUM(INV_LINE_AMT)", "INV_LTYP is Null or INV_LTYP = 'A'") & "")
        numDIST_GL.Value = DIST_GL
        Calc_Totals()
    End Sub

    Sub Calc_DIST_PO()

        Dim DISC As Decimal = 0
        For Each rowAPTINVH5 As DataRow In dst.Tables("APTINVH5").Select("")
            Dim INV_QTY As Int64 = Val(rowAPTINVH5.Item("INV_QTY") & "")
            Dim QTY_REC As Int64 = Val(rowAPTINVH5.Item("QTY_REC") & "")
            Dim TOOLG As Decimal = Val(rowAPTINVH5.Item("TOOLG") & "")
            DISC -= TOOLG * INV_QTY / QTY_REC
        Next
        DISC = System.Math.Round(DISC + 0.001, 2)
        DISC = 0 ' THE CALCULATION OF DISC AS SOMETHING PLACED IN TOOLG HAS BEEN IN PLACE SINCE THE BEGINNING
        ' AND ONLY REARS ITS UGLY HEAD WHEN THE INV_QTY IS NOT THE SAME AS THE QTY_REC
        Calc_DIST_Adjustments()
        Dim DIST_PO As Decimal = System.Math.Round(Val(dst.Tables("APTINVH5_SUM").Compute("SUM(AMT_REC_NOT_INV_OFFSET)", "") & ""), 2) - System.Math.Round(DISC, 2)
        Absx1.numFor("DIST_PO").Value = DIST_PO
        Calc_Totals()
    End Sub

    Sub Calc_DIST_Adjustments()
        Dim INV_ADJUSTMENTS As Decimal = -1 * Val(dst.Tables("APTINVH5").Compute("SUM(AMT_VAR)", "ISNULL(CB,'0') = '1'") & "") + 1 * Val(dst.Tables("APTINVH8").Compute("SUM(VOUCHER_ADJ_AMT)", "") & "")
        Absx1.numFor("INV_ADJUSTMENTS").Value = INV_ADJUSTMENTS
        Dim INV_ALLOWANCES As Decimal = -1 * Val(dst.Tables("APTINVH5").Compute("SUM(AMT_VAR)", "ISNULL(CB,'0') = '0'") & "")
        Absx1.numFor("INV_ALLOWANCES").Value = INV_ALLOWANCES
    End Sub

    Sub Calc_DIST_Other()
        Dim DIST_OTHER As Decimal = Val(dst.Tables("APTINVH7").Compute("SUM(TOTAL_INV)", "") & "")
        Absx1.numFor("DIST_OTHER").Value = DIST_OTHER
        Calc_Totals()
    End Sub

    Private Sub chkINV_1099_IND_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkINV_1099_IND.CheckedChanged
        numINV_1099_AMT.ReadOnly = Not chkINV_1099_IND.Checked Or InquiryMode Or ApprovalMode
        If Not chkINV_1099_IND.Checked Then
            numINV_1099_AMT.Value = 0
        End If
    End Sub

    Private Sub chkEXP_CATGY_CODE_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkEXP_CATGY_CODE.CheckedChanged
        grdAPTINVH2.DisplayLayout.Bands("APTINVH2").Columns("EXP_CATGY_CODE").Hidden = Not chkEXP_CATGY_CODE.Checked
    End Sub

    Private Sub chkDIST_CODE_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDIST_CODE.CheckedChanged
        grdAPTINVH2.DisplayLayout.Bands("APTINVH2").Columns("DIST_CODE").Hidden = Not chkDIST_CODE.Checked
    End Sub

    Private Sub optAPTINVH2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAPTINVH2.ValueChanged
        If dst.Tables.Count = 0 Then
            Exit Sub
        End If
        grdAPTINVH2.DisplayLayout.Bands("APTINVH2").Columns("INV_LTYP").Hidden = (optAPTINVH2.CheckedIndex <> 1)
        grdAPTINVH2.DisplayLayout.Bands("APTINVH2").Columns("INV_DLNO").Hidden = (optAPTINVH2.CheckedIndex <> 1)
    End Sub

    Private Sub dteINV_DATE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dteINV_DATE.ValueChanged
        Calculate_INV_DUE_DATE()
    End Sub

    Private Sub dteINV_BL_DATE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dteINV_BL_DATE.ValueChanged
        Calculate_INV_DUE_DATE()
    End Sub

    Sub Calculate_INV_DUE_DATE()
        If EntryMode = "" Then
            Exit Sub
        End If

        Dim INV_DATE As Object = Absx1.dteFor("INV_DATE").Value
        Dim INV_BL_DATE As Object = Absx1.dteFor("INV_BL_DATE").Value
        Dim INV_BASE_DATE As Object = Nothing

        If INV_BL_DATE & "" <> "" And Not Absx1.chkFor("VEND_DUE_FROM_INV_DATE").Checked Then
            INV_BASE_DATE = INV_BL_DATE
        Else
            INV_BASE_DATE = INV_DATE
        End If
        If INV_BASE_DATE Is Nothing Then Exit Sub
        If Absx1.txtFor("TERM_CODE").Text = "" Then
            Absx1.dteFor("INV_DUE_DATE").Value = DBNull.Value
        Else
            Try
                Absx1.dteFor("INV_DUE_DATE").Value = TAC.TACMAIN1.Calculate_INV_DUE_DATE(Me, Absx1.txtFor("TERM_CODE").Text, Nothing, INV_BASE_DATE)

            Catch ex As Exception

            End Try
        End If

        Calculate_INV_DISC_AMT(True)
    End Sub

    Sub Generate_Pre_Distribution()
        Dim VOUCHER_LNO_ctr As Integer = 0
        Dim DIST_AMT As Decimal = 0

        'If dst.Tables("APTINVH5").Rows.Count <> 0 Or dst.Tables("APTINVH7").Rows.Count <> 0 Then
        '    If ASCMAIN1.CLIENT = "AHA" Then Exit Sub
        '    ' NOT TESTED FOR INT
        'End If

        ASCMAIN1.sql = "Select * from APTVEND9 where VEND_CODE = '" & HFs("VEND_CODE") & "'"
        Dim tblAPTVEND9 As DataTable = ASCDATA1.GetDataTable
        If Absx1.cbeFor("INV_TYPE").Value = "A" Then ' Absx1.CtlFor("INV_TYPE").Text = "A" Then
            tblAPTVEND9.Rows.Clear()
            Dim rowAPTVEND9 As DataRow = tblAPTVEND9.NewRow
            rowAPTVEND9.Item("VEND_CODE") = HFs("VEND_CODE")
            rowAPTVEND9.Item("ACCT_CODE") = ROWs("APTPARM1").Item("AP_PARM_ACCT_CODE_ADVANCES")
            rowAPTVEND9.Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            rowAPTVEND9.Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            rowAPTVEND9.Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
            rowAPTVEND9.Item("DIST_AMT") = 100
            tblAPTVEND9.Rows.Add(rowAPTVEND9)
        Else
            If tblAPTVEND9.Rows.Count = 0 Then
                If rowAPTVEND1.Item("ACCT_CODE") & "" <> "" Then
                    Dim rowAPTVEND9 As DataRow = tblAPTVEND9.NewRow
                    rowAPTVEND9.Item("VEND_CODE") = HFs("VEND_CODE")
                    rowAPTVEND9.Item("ACCT_CODE") = rowAPTVEND1.Item("ACCT_CODE")
                    rowAPTVEND9.Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                    rowAPTVEND9.Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                    rowAPTVEND9.Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                    If rowAPTVEND1.Item("VEND_PRE_DIST_TYPE") & "" = "A" Then
                        DIST_AMT = Val(numINV_AMT.Value & "")
                    Else
                        DIST_AMT = 100
                    End If
                    rowAPTVEND9.Item("DIST_AMT") = DIST_AMT
                    tblAPTVEND9.Rows.Add(rowAPTVEND9)
                End If
            End If
        End If


        Dim DIST_PO As Decimal = Val(dst.Tables("APTINVH5_SUM").Compute("SUM(AMT_REC_NOT_INV_OFFSET)", "") & "")
        Dim DIST_OTHER As Decimal = Val(dst.Tables("APTINVH7").Compute("SUM(TOTAL_INV)", "") & "")
        Dim INV_AMT As Decimal = Val(numINV_AMT.Value & "")
        Dim INV_AMT_to_distribute As Decimal = INV_AMT - DIST_PO - DIST_OTHER

        For Each rowAPTVEND9 As DataRow In tblAPTVEND9.Rows
            If rowAPTVEND1.Item("VEND_PRE_DIST_TYPE") & "" = "A" And Absx1.cbeFor("INV_TYPE").Value <> "A" Then 'Absx1.CtlFor("INV_TYPE").Text <> "A" Then
                DIST_AMT = Val(rowAPTVEND9.Item("DIST_AMT") & "")
            Else
                DIST_AMT = Val(rowAPTVEND9.Item("DIST_AMT") & "") * INV_AMT_to_distribute / 100
            End If
            Dim rowAPTINVH2 As DataRow = dst.Tables("APTINVH2").NewRow
            rowAPTINVH2.Item("VOUCHER_NO") = HFs("VOUCHER_NO")
            VOUCHER_LNO_ctr += 1
            rowAPTINVH2.Item("VOUCHER_LNO") = VOUCHER_LNO_ctr
            rowAPTINVH2.Item("ACCT_CODE") = rowAPTVEND9.Item("ACCT_CODE")
            rowAPTINVH2.Item("SEG2_CODE") = rowAPTVEND9.Item("SEG2_CODE")
            rowAPTINVH2.Item("SEG3_CODE") = rowAPTVEND9.Item("SEG3_CODE")
            rowAPTINVH2.Item("SEG4_CODE") = rowAPTVEND9.Item("SEG4_CODE")
            rowAPTINVH2.Item("ACCT_DESC") = LookUp("GLTACCT1", rowAPTINVH2.Item("ACCT_CODE") & "", True).ITEM("ACCT_DESC")
            rowAPTINVH2.Item("INV_LINE_AMT") = DIST_AMT
            If Absx1.cbeFor("INV_TYPE").Value = "A" Then
                rowAPTINVH2.Item("INV_LTYP") = "A"
            End If
            dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)
        Next
        Calc_DIST_GL()
    End Sub

    Private Sub chkACCRUE_PRIOR_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkACCRUE_PRIOR.CheckedChanged
        Setup_OPS_YYYYPP_ACCRUE()
    End Sub

    Sub Setup_OPS_YYYYPP_ACCRUE()
        'Absx1.cmbFor("OPS_YYYYPP_ACCRUE").Visible = True
        Absx1.txtFor("OPS_YYYYPP_ACCRUE").Visible = True
        If Not chkACCRUE_PRIOR.Checked Then
            'Absx1.cmbFor("OPS_YYYYPP_ACCRUE").Text = ""
            Absx1.txtFor("OPS_YYYYPP_ACCRUE").Text = ""
            'Absx1.cmbFor("OPS_YYYYPP_ACCRUE").Visible = False
            Absx1.txtFor("OPS_YYYYPP_ACCRUE").Visible = False
        Else
            If Absx1.txtFor("OPS_YYYYPP_ACCRUE").Text = "" Then
                If rowAPTINVH1.RowState = DataRowState.Modified Then
                    If rowAPTINVH1.Item("OPS_YYYYPP_ACCRUE", DataRowVersion.Original) & "" <> "" Then
                        Absx1.txtFor("OPS_YYYYPP_ACCRUE").Text = rowAPTINVH1.Item("OPS_YYYYPP_ACCRUE", DataRowVersion.Original) & ""
                    End If
                End If
            End If
        End If
    End Sub

    Sub Load_OPS_YYYYPP_ACCRUE()
        If ROWs("APTPARM1").Item("AP_PARM_ALLOW_ACCRUAL") & "" = "0" Then
            If chkACCRUE_PRIOR.Checked = False Then
                Setup_OPS_YYYYPP_ACCRUE()
            Else
                chkACCRUE_PRIOR.Checked = False
            End If
            chkACCRUE_PRIOR.Visible = False
        Else
            chkACCRUE_PRIOR.Visible = True
        End If
    End Sub

    Sub Load_CURR_EXCH_RATE()
        If Absx1.txtFor("CURR_CODE").Text = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
            Absx1.numFor("CURR_EXCH_RATE").Value = 1
            Absx1.numFor("CURR_EXCH_RATE").ReadOnly = True Or InquiryMode Or ApprovalMode
        Else
            ' USE STD RATE? - IFSO - ONLY ON LOADING NEW VOUCHER
            Absx1.numFor("CURR_EXCH_RATE").ReadOnly = True Or InquiryMode Or ApprovalMode ' False
        End If
    End Sub

    Sub Setup_INV_BL_DATE()
        If dst.Tables("APTINVH5_SUM").Select("", "", DataViewRowState.CurrentRows).Length = 0 Then
            Absx1.dteFor("INV_BL_DATE").ReadOnly = False
        Else
            Absx1.dteFor("INV_BL_DATE").ReadOnly = True
            Absx1.dteFor("INV_BL_DATE").Value = dst.Tables("APTINVH5_SUM").Compute("MIN(RECEIPT_DATE)", "")
        End If
    End Sub

    Sub Negate_Voucher(ByVal VOUCHER_NO As String)
        For Each rowAPTINVH1 As DataRow In dst.Tables("APTINVH1") _
            .Select("VOUCHER_NO = '" & VOUCHER_NO & "'", "")
            For Each COLUMN_NAME As String In New String() _
            {"INV_AMT", "INV_AMT_VEND", "INV_DISC_BASED_ON", "INV_DISC_AMT", "INV_BALANCE", "INV_1099_AMT"}
                rowAPTINVH1.Item(COLUMN_NAME) = -1 * Val(rowAPTINVH1.Item(COLUMN_NAME) & "")
            Next
        Next

        For Each rowAPTINVH2 As DataRow In dst.Tables("APTINVH2") _
            .Select("VOUCHER_NO = '" & VOUCHER_NO & "'", "")
            rowAPTINVH2.Item("INV_LINE_AMT") = -1 * Val(rowAPTINVH2.Item("INV_LINE_AMT") & "")
        Next

        For Each rowAPTINVH5 As DataRow In dst.Tables("APTINVH5") _
            .Select("VOUCHER_NO = '" & VOUCHER_NO & "'", "")
            rowAPTINVH5.Item("QTY_REC_NOT_INV") = -1 * Val(rowAPTINVH5.Item("QTY_REC_NOT_INV") & "")
            rowAPTINVH5.Item("INV_QTY") = -1 * Val(rowAPTINVH5.Item("INV_QTY") & "")
            rowAPTINVH5.Item("VAR_QTY") = -1 * Val(rowAPTINVH5.Item("VAR_QTY") & "") ' PRETTY SURE THIS ONE BELONGS HERE TOO
            rowAPTINVH5.Item("VAR_AMT") = -1 * Val(rowAPTINVH5.Item("VAR_AMT") & "")
        Next

        For Each rowAPTINVH8 As DataRow In dst.Tables("APTINVH8") _
            .Select("VOUCHER_NO = '" & VOUCHER_NO & "'", "")
            rowAPTINVH8.Item("VOUCHER_ADJ_AMT") = -1 * Val(rowAPTINVH8.Item("VOUCHER_ADJ_AMT") & "")
        Next

        For Each rowAPTINVH7 As DataRow In dst.Tables("APTINVH7") _
            .Select("VOUCHER_NO = '" & VOUCHER_NO & "'", "")
            rowAPTINVH7.Item("TOTAL_INV") = -1 * Val(rowAPTINVH7.Item("TOTAL_INV") & "")
            rowAPTINVH7.Item("TOTAL_ACC") = -1 * Val(rowAPTINVH7.Item("TOTAL_ACC") & "")
        Next

    End Sub

    Sub ReNumber_Voucher(ByVal VOUCHER_NO_old As String, ByVal VOUCHER_NO_new As String)
        dst.EnforceConstraints = False
        For Each TABLE_NAME As String In New String() _
            {"APTINVH1", "APTINVH2", "APTINVH5", "APTINVH8", "APTINVH7"}
            dst.Tables(TABLE_NAME).AcceptChanges() ' IF YOU DON'T DO THIS, THEN THE DELETED ROWS WILL REMAIN IN THE TABLE AND BE DELETED FROM THE VOUCHER YOU ARE RENUMBERING FROM 
            ReNumber_Voucher_1(TABLE_NAME, VOUCHER_NO_old, VOUCHER_NO_new)
        Next
        dst.EnforceConstraints = True
    End Sub

    Sub ReNumber_Voucher_1(ByVal TABLE_NAME As String,
    ByVal VOUCHER_NO_old As String,
    ByVal VOUCHER_NO_new As String)
        For Each row As DataRow In dst.Tables(TABLE_NAME).Select("VOUCHER_NO = '" & VOUCHER_NO_old & "'", "")
            row.Item("VOUCHER_NO") = VOUCHER_NO_new
            If TABLE_NAME = "APTINVH1" Then
                row.Item("VOUCHER_NO_ORIG") = VOUCHER_NO_old
            End If
            row.AcceptChanges()
            row.SetAdded()
        Next
    End Sub

    Function Check_Invoice() As String
        ASCMAIN1.sql = "Select * from APTINVH1 " _
        & " where VEND_CODE = :PARM1" _
        & "   and INV_NUM = :PARM2" _
        & "   and INV_TYPE = :PARM3" _
        & "   and INV_STATUS IN ('O','H','P')"

        If Absx1.txtFor("VOUCHER_NO").Text <> "" Then
            ASCMAIN1.sql &= "   and VOUCHER_NO <> '" & Absx1.txtFor("VOUCHER_NO").Text & "'"
        End If

        Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVV", New Object() {Absx1.txtFor("VEND_CODE").Text, Absx1.txtFor("INV_NUM").Text, Absx1.cbeFor("INV_TYPE").Value})

        If row IsNot Nothing And Not batch_update Then
            If "YES" <> InputBox("Invoice Number " & Absx1.txtFor("INV_NUM").Text & " has Already been Entered" & vbCrLf & "(Voucher " & row.Item("VOUCHER_NO") & ")" & vbCrLf & vbCrLf & "Enter YES to Proceed", "Enter YES to Proceed") Then
                Check_Invoice = "NO"
                Exit Function
            End If
        Else
            'EXI130591
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "INV_NUM", "REPLACE(INV_NUM,'0','')")
            Dim TBL As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "VVV", New Object() {Absx1.txtFor("VEND_CODE").Text, Replace(Absx1.txtFor("INV_NUM").Text, "0", ""), Absx1.cbeFor("INV_TYPE").Value})
            If TBL.Rows.Count <> 0 And Not batch_update Then
                Using FRM As New ASFMSGBF
                    FRM.Show_grd(TBL, Me, "Invoices with Invoice Numbers like " & Absx1.txtFor("INV_NUM").Text, "")
                    If FRM.user_option = -1 Then
                        Check_Invoice = "NO"
                        Exit Function
                    End If

                End Using
                'If "YES" <> InputBox("Invoice Number " & Absx1.txtFor("INV_NUM").Text & " may have Already been Entered" & vbCrLf & "(Voucher " & row2.Item("VOUCHER_NO") & ")" & vbCrLf & "(Invoice " & row2.Item("INV_NUM") & ")" & vbCrLf & vbCrLf & "Enter YES to Proceed", "Enter YES to Proceed") Then
                '    Check_Invoice = "NO"
                '    Exit Function
                'End If
            End If
        End If
        Check_Invoice = "YES"
    End Function

    Sub Load_Alternate_Payment_Address()
        LookUp("APTVEND2", New String() {HFs("VEND_CODE"), Absx1.txtFor("VEND_ALT_CODE").Text})
        If cdr IsNot Nothing Then
            Dim rowAPTVEND2 As DataRow = dst.Tables("APTVEND2").NewRow
            rowAPTVEND2.ItemArray = cdr.ItemArray
            dst.Tables("APTVEND2").Rows.Clear()
            dst.Tables("APTVEND2").Rows.Add(rowAPTVEND2)

            'For i As Integer = 0 To .Columns.Count - 1
            '    Dim COLUMN_NAME As String = .Columns(i).ColumnName
            '    If COLUMN_NAME <> "VEND_ALT_CODE" Then
            '        COLUMN_NAME = Replace(COLUMN_NAME, "_ALT_", "_")
            '        rowAPTVEND2.Item(i) = rowAPTVEND1.Item(COLUMN_NAME)
            '    End If
            'Next

        End If
    End Sub


    Private Sub numDIST_OOBAL_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numDIST_OOBAL.ValueChanged
        grdAPTINVH2.DisplayLayout.Bands(0).SummaryFooterCaption = "Voucher Totals; Total Un-Distributed Amount = " & Format(numDIST_OOBAL.Value, "$#,##0.00") & "; Double-Click the Amount Cell to Auto-Balance"
    End Sub

    Private Sub grdAPTINVR1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdAPTINVR1.DoubleClickRow
        If grdAPTINVR1.ActiveRow IsNot Nothing Then
            Absx1.txtFor("VOUCHER_NO").Text = grdAPTINVR1.ActiveRow.Cells("VOUCHER_NO").Text
            Click_Command("Edit")
        End If

    End Sub

    Private Sub grdAPTINVH2_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdAPTINVH2.DoubleClickCell
        If e.Cell.Column.Key = "INV_LINE_AMT" Then
            If grdAPTINVH2.ActiveCell Is Nothing Then
            Else
                grdAPTINVH2.ActiveCell.Value = Val(grdAPTINVH2.ActiveCell.Value & "") + Val(numDIST_OOBAL.Value & "")
                grdAPTINVH2.UpdateData()
            End If
        End If
    End Sub

    Private Sub grpMode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub chkQuickEntry_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuickEntry.CheckedChanged
        Setup_QE(chkQuickEntry.Checked)
    End Sub

    Sub Setup_QE(ByVal tf As Boolean)
        lblQE_INV_DATE.Visible = tf
        lblQE_INV_AMT.Visible = tf
        lblQE_INV_REF.Visible = tf
        Absx1.CtlFor("QE_INV_DATE").Visible = tf
        Absx1.CtlFor("QE_INV_AMT").Visible = tf
        Absx1.CtlFor("QE_INV_REF").Visible = tf

        lblINV_TYPE.Visible = Not tf
        Absx1.CtlFor("INV_TYPE").Visible = Not tf
        lblVOUCHER_NO.Visible = Not tf
        Absx1.CtlFor("VOUCHER_NO").Visible = Not tf
    End Sub

    Sub ReLoad_Voucher(ByVal VOUCHER_NO As String)
        EnforceConstraints(False)
        Fill_Records("APTINVH1", VOUCHER_NO)
        Fill_Records("APTINVH2", VOUCHER_NO)
        Fill_Records("APTINVH5", VOUCHER_NO)
        Fill_Records("APTINVH8", VOUCHER_NO)
        Fill_Records("APTINVH7", VOUCHER_NO)

        Load_ICTIREC1()

        EnforceConstraints(True)
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
                  ", AMT_INV = NVL(AMT_INV,0) - NVL(R1.INV_QTY,0) * NVL(R1.INV_COST,0)",
                  ", AMT_INV = NVL(AMT_INV,0) + NVL(R1.INV_QTY,0) * NVL(R1.INV_COST,0)") & vbCrLf _
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

        Dim SQLN As String = ""
        'If EntryMode = "N" Then
        '    SQLN = " and NVL(CTL_NO_CREATED,'0') <> '1'"
        'End If

        'ASCMAIN1.sql = "" _
        '    & "Begin " & vbCrLf _
        '    & " Declare " & vbCrLf _
        '    & "  Cursor C1 is Select APTINVH7.*, APTACRC1.PPD_IND" & vbCrLf _
        '    & " from APTINVH7, APTACRC1" & vbCrLf _
        '    & " where APTACRC1.CTL_NO = APTINVH7.CTL_NO" & vbCrLf _
        '    & "   and APTINVH7.VOUCHER_NO = '" & VOUCHER_NO & "'" & SQLN & ";" & vbCrLf _
        '    & " Begin " & vbCrLf _
        '    & "  For R1 in C1 Loop" & vbCrLf _
        '    & "   If NVL(R1.PPD_IND,'0') = '1' Then " & vbCrLf _
        '    & "    Update APTACRC1 set " & vbCrLf _
        '    & IIf(reverse,
        '            "PPD_MATCHED = '0'",
        '            "PPD_MATCHED = '1'") & vbCrLf _
        '    & "    where CTL_NO = R1.CTL_NO;" & vbCrLf _
        '    & "   Else" & vbCrLf _
        '    & "    Update APTACRC1 set " & vbCrLf _
        '    & IIf(reverse,
        '            "COST_ACT = NVL(COST_ACT,0) - NVL(R1.TOTAL_INV,0), CTL_STATUS = DECODE(CTL_TYPE,'M','X','0'), VOUCHER_NO = NULL",
        '            "COST_ACT = NVL(COST_ACT,0) + NVL(R1.TOTAL_INV,0), CTL_STATUS = '1', VOUCHER_NO = R1.VOUCHER_NO") & vbCrLf _
        '    & "    where CTL_NO = R1.CTL_NO;" & vbCrLf _
        '    & "   End If;" & vbCrLf _
        '    & "  End Loop; " & vbCrLf _
        '    & " End;" & vbCrLf _
        '    & "End;"

        'ASCMAIN1.sql = "" _
        '    & "Begin " & vbCrLf _
        '    & " Declare " & vbCrLf _
        '    & "  Cursor C1 is Select APTINVH7.*, APTACRC1.PPD_IND" & vbCrLf _
        '    & " from APTINVH7, APTACRC1" & vbCrLf _
        '    & " where APTACRC1.CTL_NO = APTINVH7.CTL_NO" & vbCrLf _
        '    & "   and APTINVH7.VOUCHER_NO = '" & VOUCHER_NO & "'" & SQLN & ";" & vbCrLf _
        '    & " Begin " & vbCrLf _
        '    & "  For R1 in C1 Loop" & vbCrLf _
        '    & "   If NVL(R1.PPD_IND,'0') = '0' Then " & vbCrLf _
        '    & "    Update APTACRC1 set " & vbCrLf _
        '    & IIf(reverse,
        '            "COST_ACT = NVL(COST_ACT,0) - NVL(R1.TOTAL_INV,0), CTL_STATUS = DECODE(CTL_TYPE,'M','X','0'), VOUCHER_NO = NULL",
        '            "COST_ACT = NVL(COST_ACT,0) + NVL(R1.TOTAL_INV,0), CTL_STATUS = '1', VOUCHER_NO = R1.VOUCHER_NO") & vbCrLf _
        '    & "    where CTL_NO = R1.CTL_NO;" & vbCrLf _
        '    & "   End If;" & vbCrLf _
        '    & "  End Loop; " & vbCrLf _
        '    & " End;" & vbCrLf _
        '    & "End;"

        ASCMAIN1.sql = "" _
            & "Begin " & vbCrLf _
            & " Declare " & vbCrLf _
            & "  Cursor C1 is Select APTINVH7.*, APTACRC1.PPD_IND" & vbCrLf _
            & " from APTINVH7, APTACRC1" & vbCrLf _
            & " where APTACRC1.CTL_NO = APTINVH7.CTL_NO" & vbCrLf _
            & "   and APTINVH7.VOUCHER_NO = '" & VOUCHER_NO & "'" & SQLN & ";" & vbCrLf _
            & " Begin " & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   If NVL(R1.PPD_IND,'0') = '1' Then " & vbCrLf _
            & "    Update APTACRC1 set " & vbCrLf _
            & IIf(reverse,
                    "CTL_STATUS = DECODE(CTL_TYPE,'M','X','0')",
                    "CTL_STATUS = '1'") & vbCrLf _
            & "    where CTL_NO = R1.CTL_NO;" & vbCrLf _
            & "   Else" & vbCrLf _
            & "    Update APTACRC1 set " & vbCrLf _
            & IIf(reverse,
                    "COST_ACT = NVL(COST_ACT,0) - NVL(R1.TOTAL_INV,0), CTL_STATUS = DECODE(CTL_TYPE,'M','X','0'), VOUCHER_NO = NULL",
                    "COST_ACT = NVL(COST_ACT,0) + NVL(R1.TOTAL_INV,0), CTL_STATUS = '1', VOUCHER_NO = R1.VOUCHER_NO") & vbCrLf _
            & "    where CTL_NO = R1.CTL_NO;" & vbCrLf _
            & "   End If;" & vbCrLf _
            & "  End Loop; " & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"


        ASCDATA1.ExecuteSQL()


        Dim INV_AMT As Decimal = Val(rowAPTINVH1("INV_AMT") & "")
        If reverse Then
            Dim row As DataRow = LookUp("APTINVH1", VOUCHER_NO)
            INV_AMT = -1 * Val(row("INV_AMT") & "")
        End If
        'rowAPTVEND5 = Fill_Record("APTVEND5", HFs("VEND_CODE"), True)

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

    Private Sub chkRecurring_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRecurring.CheckedChanged
        Show_Batch()
        'optstatus will be R only and invisible
    End Sub

    Sub Show_Batch()

        If dst.Tables.Count = 0 Then Exit Sub

        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()

        Dim sql As String = "from APTINVH1,APTVEND1 where APTVEND1.VEND_CODE = APTINVH1.VEND_CODE"
        Dim caption As String = ""

        If chkRecurring.Checked Then
            sql &= " and APTINVH1.INV_STATUS = 'R'"
            caption = "Recurring Invoice Templates"
        Else
            If ApprovalMode Then
                sql &= " and APTINVH1.INV_STATUS = 'O' and INV_APPR_STATUS = 'P'"
                caption = "Invoices Pending Approval"
            Else
                sql &= " and APTINVH1.REGISTER_IND = '0'"
                caption = "Invoices Pending Invoice Register Update"
            End If
        End If

        If ApprovalMode Then
            If cmbVEND_BUYER_CODE.Value = "*" Then
                caption &= ", queued for Approval"
            Else
                sql &= " and APTINVH1.VEND_BUYER_CODE = '" & cmbVEND_BUYER_CODE.Value & "'"
                caption &= ", queued for Approval by " & cmbVEND_BUYER_CODE.Value
            End If
        Else
            If optFilter.Value = "R" Then
                sql &= " and APTINVH1.INIT_OPER = '" & ASCMAIN1.USER_ID & "'"
                caption &= ", entered by " & ASCMAIN1.USER_ID
            ElseIf optFilter.Value = "V" Then
                sql &= " and APTVEND1.PROCESSOR_CODE = '" & ASCMAIN1.USER_ID & "'"
                caption &= ", for processor " & ASCMAIN1.USER_ID
            End If
        End If

        EnforceConstraints(False)
        sql_APTINVR2 = "Select APTINVH2.*,GLTACCT1.ACCT_DESC from APTINVH2,GLTACCT1 where GLTACCT1.ACCT_CODE = APTINVH2.ACCT_CODE and APTINVH2.VOUCHER_NO in (Select APTINVH1.VOUCHER_NO " & sql & ")"
        sql_APTVEND1 = "Select APTVEND1.* from APTVEND1 where APTVEND1.VEND_CODE in (Select DISTINCT APTINVH1.VEND_CODE " & sql & ")"
        sql = "Select APTINVH1.* " & sql
        grdAPTINVR1.Text = caption
        Fill_Records("APTINVR1", "", True, sql)
        Fill_Records("APTINVR2", "", , sql_APTINVR2)

        Sort_grdColumns(grdAPTINVR1, "VOUCHER_NO")
        Sort_grdColumns(grdAPTINVR1, "VOUCHER_LNO", , 1)
        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
    End Sub

    Sub Set_Recurring(ByVal Recurring As Boolean)
        'tabHeader.Tabs("Recurring").Enabled = (optINV_STATUS.Value = "R")
        grpINV_STATUS.Visible = Not Recurring
    End Sub

    Private Sub UltraTabPageControl2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles UltraTabPageControl2.Paint

    End Sub

    Sub Calculate_INV_DISC_AMT(Optional ByVal update_UI As Boolean = False)

        'Me.Validate()
        If Absx1.dteFor("INV_DATE").Value & "" = "" Then Exit Sub

        Dim TERM_CODE As String
        If update_UI Then
            TERM_CODE = Absx1.txtFor("TERM_CODE").Text
        Else
            TERM_CODE = rowAPTINVH1("TERM_CODE")
        End If
        Dim rowTATTERM1 As DataRow = LookUp("TATTERM1", TERM_CODE, True)

        Dim INV_DISC_BASED_ON As Double = Val(rowAPTINVH1("INV_AMT") & "")
        If update_UI Then
            INV_DISC_BASED_ON = Val(Absx1.numFor("INV_AMT").Value & "")
            'Absx1.numFor("INV_DISC_BASED_ON").Value = INV_DISC_BASED_ON
        Else
            rowAPTINVH1("INV_DISC_BASED_ON") = INV_DISC_BASED_ON
        End If

        Dim TERM_DAYS_DISC As Double = Val(rowTATTERM1("TERM_DAYS_DISC") & "")
        Dim TERM_DISC_PERC As Double = Val(rowTATTERM1("TERM_DISC_PERC") & "")

        If TERM_DISC_PERC = 0 Or HFs("INV_TYPE") <> "I" Then
            If update_UI Then
                Absx1.dteFor("INV_DISC_DUE").Value = Null
                Absx1.numFor("INV_DISC_AMT").Value = 0
            Else
                rowAPTINVH1("INV_DISC_DUE") = Null
                rowAPTINVH1("INV_DISC_AMT") = 0
            End If
        Else
            If rowTATTERM1("TERM_DISC_ELIG_DUE") & "" = "1" Then
                If update_UI Then
                    Absx1.dteFor("INV_DISC_DUE").Value = Absx1.dteFor("INV_DUE_DATE").Value
                Else
                    rowAPTINVH1("INV_DISC_DUE") = rowAPTINVH1("INV_DUE_DATE")
                End If
            Else
                If update_UI Then
                    Absx1.dteFor("INV_DISC_DUE").Value = DateValue(Absx1.dteFor("INV_DATE").Value).AddDays(TERM_DAYS_DISC)
                Else
                    rowAPTINVH1("INV_DISC_DUE") = DateValue(rowAPTINVH1("INV_DATE")).AddDays(TERM_DAYS_DISC)
                End If
            End If
            If update_UI Then
                Absx1.numFor("INV_DISC_AMT").Value = Round(INV_DISC_BASED_ON * TERM_DISC_PERC / 100, 2)
            Else
                rowAPTINVH1("INV_DISC_AMT") = Round(INV_DISC_BASED_ON * TERM_DISC_PERC / 100, 2)
            End If
        End If
    End Sub

    Private Sub numQE_INV_AMT_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles numQE_INV_AMT.Enter
        'numQE_INV_AMT.SelectAll()
    End Sub

    Private Sub numQE_INV_AMT_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles numQE_INV_AMT.GotFocus
        numQE_INV_AMT.SelectAll()
    End Sub

    Sub Print_Record()
        Print_Report_Begin()

        'Fill_Records("APTINVR2", "", , sql_APTINVR2)
        Fill_Records("APTVEND1", "", , sql_APTVEND1)

        Dim SUBT As String = "Edit List"
        'CR_params.Add("REPORT_PARAMETER_NAME", "VALUE")
        'RecordSelectionFormula = "{EDTSLSVP.SLS} > 10"
        Generate_Report("APRINVR1", "Vendor Invoice Entry", SUBT)

        Print_Report_End()

        dst.Tables("APTVEND1").Rows.Clear()
    End Sub

    Private Sub UltraNumericEditor4_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraNumericEditor4.ValueChanged

    End Sub

    Private Sub optFilter_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optFilter.ValueChanged
        Show_Batch()
    End Sub

    Sub Update_as_Paid(ByVal VOUCHER_NO As String)

        rowAPTINVH1 = Fill_Record("APTINVH1", VOUCHER_NO)
        Dim rowAPTINVHX As DataRow = dst.Tables("APTINVHX").NewRow
        rowAPTINVHX("VOUCHER_NO") = rowAPTINVH1("VOUCHER_NO")
        rowAPTINVHX("INV_NUM") = rowAPTINVH1("INV_NUM")
        rowAPTINVHX("INV_DATE") = rowAPTINVH1("INV_DATE")
        rowAPTINVHX("INV_BALANCE") = rowAPTINVH1("INV_BALANCE")
        rowAPTINVHX("INV_DISC_AMT") = rowAPTINVH1("INV_DISC_AMT")
        rowAPTINVHX("SELECTED") = "1"
        dst.Tables("APTINVHX").Rows.Add(rowAPTINVHX)

        Dim VEND_CODE As String = rowAPTINVH1("VEND_CODE") & ""
        Dim VEND_CODE_AP As String = rowAPTINVH1("VEND_CODE_AP") & ""
        If VEND_CODE_AP = "" Then
            VEND_CODE_AP = VEND_CODE
        End If
        Dim VEND_ALT_CODE As String = rowAPTINVH1("VEND_ALT_CODE") & ""
        If VEND_ALT_CODE = "" Then
            VEND_ALT_CODE = "VENDOR"
        End If

        Dim VEND_NAME As String
        Dim rowPayee As DataRow
        If VEND_CODE_AP <> "" And VEND_CODE_AP <> VEND_CODE Then
            rowPayee = LookUp("APTVEND1", VEND_CODE_AP)
            VEND_NAME = rowPayee.Item("VEND_NAME")
        Else
            VEND_NAME = HFs("VEND_NAME")
        End If

        Dim BANK_CODE As String = rowAPTINVH1("BANK_CODE")
        Dim PYMT_METHOD As String = rowAPTINVH1("INV_PYMT_METHOD")
        Dim CHECK_NUM As String = rowAPTINVH1("CHECK_NUM")
        Dim CHECK_DATE As Date = rowAPTINVH1("CHECK_DATE")
        Dim CHECK_AMT As Double = Absx1.numFor("CHECK_AMT").Value

        Dim SEQ_NUM As Integer
        SEQ_NUM = 0
        For Each rowAPTINVHX In dst.Tables("APTINVHX").Select("SELECTED = '1'", "")
            If rowAPTINVHX("VOUCHER_NO") = VOUCHER_NO Then
                rowAPTINVH1 = dst.Tables("APTINVH1").Rows.Find(VOUCHER_NO)
            Else
                rowAPTINVH1 = Fill_Record("APTINVH1", rowAPTINVHX("VOUCHER_NO"), , False)
            End If
            rowAPTINVH1("INV_STATUS") = "P"
            rowAPTINVH1("INV_PAYMENTS") = rowAPTINVHX("INV_BALANCE")
            rowAPTINVH1("INV_DISC_TAKEN") = rowAPTINVHX("INV_DISC_AMT")
            rowAPTINVH1("INV_LAST_PMT_DATE") = CHECK_DATE
            rowAPTINVH1("BATCH_NO_PYMT") = ""
            rowAPTINVH1("INV_BALANCE") = 0
            rowAPTINVH1("BATCH_PYMT") = 0
            rowAPTINVH1("BATCH_DISC") = 0
            rowAPTINVH1("BANK_CODE") = BANK_CODE
            rowAPTINVH1("CHECK_NUM") = CHECK_NUM
            rowAPTINVH1("CHECK_DATE") = CHECK_DATE

            Dim rowAPTCHCK2 As DataRow = dst.Tables("APTCHCK2").NewRow
            rowAPTCHCK2("BANK_CODE") = BANK_CODE
            rowAPTCHCK2("CHECK_NUM") = CHECK_NUM
            SEQ_NUM = SEQ_NUM + 1
            rowAPTCHCK2("SEQ_NUM") = SEQ_NUM
            rowAPTCHCK2("VEND_CODE") = rowAPTINVH1("VEND_CODE")
            rowAPTCHCK2("INV_NUM") = rowAPTINVH1("INV_NUM")
            rowAPTCHCK2("INV_DATE") = rowAPTINVH1("INV_DATE")
            rowAPTCHCK2("VOUCHER_NO") = rowAPTINVH1("VOUCHER_NO")
            rowAPTCHCK2("INV_AMT_APPLIED") = rowAPTINVH1("INV_AMT")
            rowAPTCHCK2("INV_DISC_TAKEN") = rowAPTINVH1("INV_DISC_AMT")
            dst.Tables("APTCHCK2").Rows.Add(rowAPTCHCK2)
        Next

        Dim rowAPTCHCK1 As DataRow = dst.Tables("APTCHCK1").NewRow
        rowAPTCHCK1("BANK_CODE") = BANK_CODE
        rowAPTCHCK1("CHECK_NUM") = CHECK_NUM
        rowAPTCHCK1("CHECK_DATE") = CHECK_DATE
        rowAPTCHCK1("CHECK_AMT") = CHECK_AMT
        rowAPTCHCK1("PYMT_METHOD") = PYMT_METHOD
        rowAPTCHCK1("VEND_CODE") = HFs("VEND_CODE")
        rowAPTCHCK1("VEND_CODE_AP") = VEND_CODE_AP
        rowAPTCHCK1("VEND_ALT_CODE") = VEND_ALT_CODE
        rowAPTCHCK1("OPS_YYYYPP") = ASCMAIN1.CYP
        rowAPTCHCK1("CHECK_STATUS") = "I"
        rowAPTCHCK1("VEND_NAME") = VEND_NAME
        rowAPTCHCK1("INIT_DATE") = DATETIME_STAMP
        rowAPTCHCK1("INIT_OPER") = ASCMAIN1.USER_ID
        rowAPTCHCK1("REGISTER_IND") = "0"
        dst.Tables("APTCHCK1").Rows.Add(rowAPTCHCK1)

        Dim INV_PAYMENTS As Double = Val(dst.Tables("APTCHCK2").Compute("SUM(INV_AMT_APPLIED)", "") & "")
        Dim INV_DISC_TAKEN As Double = Val(dst.Tables("APTCHCK2").Compute("SUM(INV_DISC_TAKEN)", "") & "")

        rowAPTVEND5.Item("VEND_PAYMENTS_MTD") = Val(rowAPTVEND5.Item("VEND_PAYMENTS_MTD") & "") + INV_PAYMENTS
        rowAPTVEND5.Item("VEND_PAYMENTS_YTD") = Val(rowAPTVEND5.Item("VEND_PAYMENTS_YTD") & "") + INV_PAYMENTS
        rowAPTVEND5.Item("VEND_DISC_TAKEN_MTD") = Val(rowAPTVEND5.Item("VEND_DISC_TAKEN_MTD") & "") + INV_DISC_TAKEN
        rowAPTVEND5.Item("VEND_DISC_TAKEN_YTD") = Val(rowAPTVEND5.Item("VEND_DISC_TAKEN_YTD") & "") + INV_DISC_TAKEN
        rowAPTVEND5.Item("VEND_NUM_CHKS_MTD") = Val(rowAPTVEND5.Item("VEND_NUM_CHKS_MTD") & "") + 1
        rowAPTVEND5.Item("VEND_NUM_CHKS_YTD") = Val(rowAPTVEND5.Item("VEND_NUM_CHKS_YTD") & "") + 1
        rowAPTVEND5.Item("VEND_LAST_PMT_DATE") = CHECK_DATE
        rowAPTVEND5.Item("VEND_LAST_PMT_AMT") = INV_PAYMENTS

        If auto_next_check Then
            Dim rowGLTBANK1 As DataRow = Fill_Record("GLTBANK1", rowAPTINVH1("BANK_CODE"))
            If rowGLTBANK1("BANK_LAST_CHECK_NO") & "" = BANK_LAST_CHECK_NO Then
                rowGLTBANK1("BANK_LAST_CHECK_NO") = BANK_NEXT_CHECK_NO
            End If
        End If

        'If last_check_no <> "" And LAST_CHECK_NO_bank = datAPWINVH1.Recordset.Fields("BANK_CODE").Value Then
        '    OraD.Parameters("CODE").Value = datAPWINVH1.Recordset.Fields("BANK_CODE").Value
        '    dynGLTBANK1.Refresh()
        '    dynGLTBANK1.Edit()
        '    dynGLTBANK1.Fields("LAST_CHECK_NO").Value = last_check_no
        '    dynGLTBANK1.Update()
        'End If
    End Sub

    Private Sub grdAPTINVHX_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTINVHX.AfterCellUpdate
    End Sub

    Private Sub grdAPTINVHX_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdAPTINVHX.AfterRowUpdate
        grdAPTINVHX.DisplayLayout.Bands(0).SummaryFooterCaption = "Total Amount Selected = " & Format(Val(dst.Tables("APTINVHX").Compute("Sum(INV_PAYMENTS)", "SELECTED = '1'") & ""), "$#,##0.00")
    End Sub

    Private Sub grdAPTINVHX_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTINVHX.CellChange
        grdAPTINVHX.Update()
    End Sub

    'Private Sub chkBatchEntry_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBatchEntry.CheckedChanged
    '    chkQuickEntry.Enabled = Not chkBatchEntry.Checked
    '    chkRecurring.Enabled = Not chkBatchEntry.Checked
    '    optFilter.Enabled = Not chkBatchEntry.Checked
    '    grdAPTINVR1.Visible = Not chkBatchEntry.Checked
    '    UltraGroupBox1.Visible = Not chkBatchEntry.Checked
    '    grdAPTINVHB.Visible = chkBatchEntry.Checked
    'End Sub

    Private Sub UltraTextEditor2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraTextEditor2.ValueChanged

    End Sub

    Sub Prepare_for_Multi_Invoice_Edit()

        Fill_Record("APTINVHM", Absx1.txtFor("VEND_CODE").Text)
        grdAPTINVHM.DisplayLayout.Bands(0).Columns("UPDATE_STATUS").Hidden = True
        grdAPTINVHM.DisplayLayout.Bands(0).Columns("UPDATE_MESSAGE").Hidden = True
        grpLastChange.Dock = DockStyle.None
        grpLastChange.Left = Absx1.txtFor("VEND_NAME").Left + Absx1.txtFor("VEND_NAME").Width + 2
        grpLastChange.Dock = DockStyle.Right
    End Sub

    Sub Prepare_for_Batch_Entry()

    End Sub

    Private Sub grdAPTINVHB_AfterCellUpdate(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTINVHB.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "VEND_CODE"
                Dim VEND_CODE As String = e.Cell.Value & ""
                grdCodeDesc(grdAPTINVHB, "APTVEND1", "VEND_CODE", "VEND_NAME")
        End Select
    End Sub

    Private Sub grdAPTINVHB_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTINVHB.AfterExitEditMode
        With grdAPTINVHB
            Select Case .ActiveCell.Column.Key
                Case "VEND_CODE"
                    Dim VEND_CODE As String = .ActiveCell.Text
                    If VEND_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(VEND_CODE, .ActiveCell.Column.Key)
                    End If
            End Select
        End With
    End Sub

    Private Sub grdAPTINVHB_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTINVHB.AfterRowActivate
        With grdAPTINVHB
            If .ActiveRow.IsAddRow Then
                .DisplayLayout.Bands(0).Columns("VEND_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .ActiveCell = grdAPTINVHB.ActiveRow.Cells("VEND_CODE")
                .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else

            End If
        End With
    End Sub

    Private Sub grdAPTINVHB_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdAPTINVHB.BeforeRowUpdate
        With grdAPTINVHB
            If e.Row.Cells("VEND_CODE").Text = "" Then
                e.Cancel = True
            Else
                LookUp("APTVEND1", e.Row.Cells("VEND_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Vendor Code (" & e.Row.Cells("VEND_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If Not e.Cancel Then
                If e.Row.Cells("VOUCHER_NO").Text = "" Then
                    .ActiveRow.Cells("VOUCHER_NO").Value = Format(Val(dst.Tables("APTINVHB").Compute("Max(VOUCHER_NO)", "") & "") + 1, "0000000000")
                End If
            End If
        End With

    End Sub

    Private Sub grdAPTINVHB_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTINVHB.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdAPTINVHB, sql_where, sql_where <> "")
    End Sub

    Sub Update_Batch()
        ' HAVE A LINE NUMBER SERVE AS THE KEY

        batch_update = True
        Click_Command("Cancel")
        chkQuickEntry.Checked = False
        Dim V As Integer = 0
        Dim X As CurrencyManager = Me.BindingContext(dst.Tables("APTINVH1"))
        X.EndCurrentEdit()

        Application.DoEvents()

        Dim STATS(2) As Integer
        For Each rowAPTINVHB As DataRow In dst.Tables("APTINVHB").Rows
            STATS(0) += 1
            V = V + 1
            rowAPTINVHB("VOUCHER_NO") = "X" & Format(V, "000000000")
            Absx1.txtFor("VEND_CODE").Text = rowAPTINVHB("VEND_CODE") & ""
            Absx1.txtFor("INV_NUM").Text = rowAPTINVHB("INV_NUM") & ""
            If rowAPTINVHB("INV_TYPE") & "" = "" Then
                Absx1.cbeFor("INV_TYPE").Text = "I"
            Else
                'Stop
                Absx1.cbeFor("INV_TYPE").Text = rowAPTINVHB("INV_TYPE")
            End If
            Application.DoEvents()
            X.EndCurrentEdit()

            Click_Command("New")
            Application.DoEvents()
            If EMsg <> "" Then
                rowAPTINVHB("UPDATE_STATUS") = "ERROR"
                rowAPTINVHB("UPDATE_MESSAGE") = EMsg
                STATS(2) += 1
            Else
                rowAPTINVHB("VEND_NAME") = Absx1.txtFor("VEND_NAME").Text
                'rowAPTINVH1("INV_DATE") = rowAPTINVHB("INV_DATE")
                Absx1.dteFor("INV_DATE").Value = rowAPTINVHB("INV_DATE")
                'Absx1.numFor("INV_AMT").Value = rowAPTINVHB("INV_AMT")
                Absx1.numFor("INV_AMT_VEND").Value = rowAPTINVHB("INV_AMT")
                Application.DoEvents()
                Automatic_Distribution()
                If rowAPTINVHB("BANK_CODE") & "" <> "" Then
                    Absx1.txtFor("BANK_CODE").Text = rowAPTINVHB("BANK_CODE") & ""
                End If
                If rowAPTINVHB("TERM_CODE") & "" <> "" Then
                    Absx1.txtFor("TERM_CODE").Text = rowAPTINVHB("TERM_CODE") & ""
                End If
                Absx1.txtFor("INV_REF").Text = rowAPTINVHB("INV_REF") & ""
                If rowAPTINVHB("CHECK_NUM") & "" <> "" Then
                    Absx1.txtFor("CHECK_NUM").Text = rowAPTINVHB("CHECK_NUM") & ""
                    Absx1.dteFor("CHECK_DATE").Value = rowAPTINVHB("CHECK_DATE")
                End If
                If rowAPTINVHB("INV_STATUS") & "" = "" Then
                    Absx1.optFor("INV_STATUS").Value = "O"
                Else
                    Absx1.optFor("INV_STATUS").Value = rowAPTINVHB("INV_STATUS") & ""
                End If
                If Absx1.optFor("INV_STATUS").Value = "P" Then
                    If Absx1.txtFor("CHECK_NUM").Text = "" Then
                        If Absx1.txtFor("BANK_CODE").Text <> "" Then
                            Dim rowGLTBANK1 As DataRow = Fill_Record("GLTBANK1", Absx1.txtFor("BANK_CODE").Text)
                            BANK_LAST_CHECK_NO = rowGLTBANK1("BANK_LAST_CHECK_NO") & ""
                            BANK_NEXT_CHECK_NO = CStr(Val(BANK_LAST_CHECK_NO) + 1)
                            BANK_NEXT_CHECK_NO = ASCMAIN1.Format_Field(BANK_NEXT_CHECK_NO, "CHECK_NUM")
                            Absx1.txtFor("CHECK_NUM").Text = BANK_NEXT_CHECK_NO
                            Absx1.dteFor("CHECK_DATE").Value = rowAPTINVHB("INV_DATE")
                            auto_next_check = True

                            rowAPTINVHB("CHECK_NUM") = BANK_NEXT_CHECK_NO
                            rowAPTINVHB("CHECK_DATE") = rowAPTINVHB("INV_DATE")
                        End If
                    End If
                    Absx1.numFor("CHECK_AMT").Value = Absx1.numFor("INV_AMT").Value
                End If
                If rowAPTINVHB("INV_PYMT_METHOD") & "" <> "" Then
                    Absx1.txtFor("INV_PYMT_METHOD").Text = rowAPTINVHB("INV_PYMT_METHOD") & ""
                End If
                If rowAPTINVHB("INV_PYMT_CYCLE") & "" <> "" Then
                    Absx1.txtFor("INV_PYMT_CYCLE").Text = rowAPTINVHB("INV_PYMT_CYCLE") & ""
                End If
                If rowAPTINVHB("VEND_ALT_CODE") & "" <> "" Then
                    Absx1.txtFor("VEND_ALT_CODE").Text = rowAPTINVHB("VEND_ALT_CODE") & ""
                End If
                If rowAPTINVHB("POST_CODE") & "" <> "" Then
                    Absx1.txtFor("POST_CODE").Text = rowAPTINVHB("POST_CODE") & ""
                End If

                Application.DoEvents()
                tabMain.SelectedTab = tabMain.Tabs("GL Distribution")
                Application.DoEvents()
                X.EndCurrentEdit()

                Click_Command("Update")
                Application.DoEvents()
                If EMsg <> "" Then
                    rowAPTINVHB("UPDATE_STATUS") = "ERROR"
                    rowAPTINVHB("UPDATE_MESSAGE") = EMsg
                    Click_Command("Cancel")
                    STATS(2) += 1
                Else
                    rowAPTINVHB("UPDATE_STATUS") = "UPDATED"
                    STATS(1) += 1
                    rowAPTINVHB("VOUCHER_NO") = HFs("VOUCHER_NO")
                End If
            End If
            Application.DoEvents()
        Next
        batch_update = False
        grdAPTINVHB.DisplayLayout.Bands("APTINVHB").Columns("UPDATE_STATUS").Hidden = False
        grdAPTINVHB.DisplayLayout.Bands("APTINVHB").Columns("UPDATE_MESSAGE").Hidden = False
        Export_to_Excel(grdAPTINVHB)
        grdAPTINVHB.DisplayLayout.Bands("APTINVHB").Columns("UPDATE_STATUS").Hidden = True
        grdAPTINVHB.DisplayLayout.Bands("APTINVHB").Columns("UPDATE_MESSAGE").Hidden = True

        MsgBox("Batch Update Complete." & vbCr & vbCr & STATS(0) & " Records Processed" & vbCr & STATS(1) & " Records Updated Successfully" & vbCr & STATS(2) & " Records Were NOT Updated" & vbCr & vbCr & "Voucher (and Check Numbers, if any) appear in the Workbook Generated" & vbCr & "Processing Errors (if any) also appear in the 'Update Message' Column of the Workbook", MsgBoxStyle.OkOnly, "Verification")
    End Sub

    Sub Import_Batch_from_Excel()
        Dim FILENAME As String = ""

        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.InitialDirectory = "c:\"
            openFileDialog1.Title = "Locate the workbook containing AP Items to Import"
            openFileDialog1.Filter = "txt files (*.xls)|*.xls|All files (*.*)|*.*"
            openFileDialog1.FilterIndex = 2
            openFileDialog1.RestoreDirectory = True

            If openFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        If FILENAME <> "" Then

            Dim GCs As New Dictionary(Of String, String)
            With grdAPTINVHB.DisplayLayout.Bands("APTINVHB")
                For j As Integer = 1 To .Columns.Count
                    GCs.Add(.Columns(j - 1).Header.Caption, .Columns(j - 1).Key)
                Next
            End With

            grdAPTINVHB.DataSource = DirectCast(grdAPTINVHB.DataSource, DataTable).Clone

            Dim xlApp As Object
            Dim xlBook As Object
            Dim xlSheet As Object

            Try
                ASCMAIN1.Progress("Now Examining XLS Workbook")
                Me.Cursor = Cursors.WaitCursor

                Dim XLS As New Infragistics.Documents.Excel.Workbook

                ' Create the Excel App Object 
                xlApp = CreateObject("Excel.Application")
                ' Create the Excel Workbook Object. 
                xlBook = xlApp.Workbooks.Open(FILENAME)
                ' XLS = DirectCast(xlBook, Infragistics.Excel.Workbook)
                xlSheet = xlBook.Sheets(1)
                Dim heading_row As Integer = 0
                Dim columns_to_import As Integer = 0
                Dim COLUMN_NAMEs() As String
                ReDim COLUMN_NAMEs(0)
                Dim found_heading_row As Boolean
                For i As Integer = 1 To 10
                    found_heading_row = False
                    If xlSheet.cells(i, 1).text <> "" Then
                        found_heading_row = True
                        ReDim COLUMN_NAMEs(0)
                        For j As Integer = 1 To GCs.Count
                            Dim CellText As String = xlSheet.cells(i, j).text
                            If j > 1 And CellText = "" Then
                                columns_to_import = j - 1
                                Exit For
                            End If
                            If Not GCs.ContainsKey(CellText) Then
                                found_heading_row = False
                                Exit For
                            Else
                                ReDim Preserve COLUMN_NAMEs(j)
                                COLUMN_NAMEs(j) = GCs(CellText)
                                columns_to_import = j
                            End If
                        Next
                        If found_heading_row Then
                            heading_row = i
                            Exit For
                        End If
                    End If
                Next
                If heading_row = 0 Then
                    MsgBox("Cannot Find Heading Row", MsgBoxStyle.OkOnly, "Problem with Workbook Selected")
                Else
                    ASCMAIN1.Progress("Now Importing Data")
                    dst.Tables("APTINVHB").Rows.Clear()

                    Dim XR As Integer = heading_row + 1
                    Dim XI As Integer = 0
                    Do While xlSheet.cells(XR, 1).text <> "" And xlSheet.cells(XR, 1).text <> "Totals"
                        ASCMAIN1.Progress("-", CStr(XR - heading_row))
                        Dim rowAPTINVHB As DataRow = dst.Tables("APTINVHB").NewRow
                        For XC As Integer = 1 To columns_to_import
                            Dim CellText As String = xlSheet.cells(XR, XC).value & ""
                            COLUMN_NAME = COLUMN_NAMEs(XC)
                            If CellText <> "" Then
                                rowAPTINVHB(COLUMN_NAME) = CellText
                            End If
                        Next
                        If Len(rowAPTINVHB("INV_TYPE") & "") > 1 Then
                            If rowAPTINVHB("INV_TYPE") = "ChargeBack" Then
                                rowAPTINVHB("INV_TYPE") = "B"
                            Else
                                rowAPTINVHB("INV_TYPE") = Mid(rowAPTINVHB("INV_TYPE") & "", 1, 1)
                            End If
                        End If
                        If rowAPTINVHB("INV_STATUS") & "" <> "" Then

                        End If
                        rowAPTINVHB("VOUCHER_NO") = "" ' make sure that we use our own
                        If rowAPTINVHB("VEND_CODE") & "" <> "" Then
                            rowAPTINVHB("VEND_CODE") = ASCMAIN1.Format_Field(rowAPTINVHB("VEND_CODE"), "VEND_CODE")
                        End If
                        If rowAPTINVHB("INV_STATUS") & "" <> "" Then
                            rowAPTINVHB("INV_STATUS") = rowAPTINVHB("INV_STATUS").ToString.ToUpper
                            If rowAPTINVHB("INV_STATUS") <> "P" And rowAPTINVHB("INV_STATUS") <> "O" And rowAPTINVHB("INV_STATUS") <> "H" Then
                                rowAPTINVHB("INV_STATUS") = "O"
                            End If
                        End If

                        If rowAPTINVHB("VOUCHER_NO") & "" = "" Then
                            rowAPTINVHB("VOUCHER_NO") = Format(XR - heading_row, "0000000000")
                        End If
                        If rowAPTINVHB("UPDATE_STATUS") & "" = "UPDATED" Then
                            XI = XI + 1
                        Else
                            dst.Tables("APTINVHB").Rows.Add(rowAPTINVHB)
                        End If
                        XR = XR + 1
                    Loop
                    MsgBox("Import Successful" & vbCr & vbCr & "Records Processed = " & CStr(XR - (heading_row + 1)) & vbCr & "Records Imported = " & CStr(XR - (heading_row + 1) - XI) & vbCr & "Records Ignored (Updated Already) = " & CStr(XI), MsgBoxStyle.OkOnly, "Verification")
                End If

                xlApp.DisplayAlerts = False
                xlApp.Quit()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Exception Occurred")
            Finally

                xlSheet = Nothing
                xlBook = Nothing
                xlApp = Nothing

                ASCMAIN1.Progress("")
                Me.Cursor = Cursors.Default
            End Try

            grdAPTINVHB.DataSource = dst.Tables("APTINVHB")
            grdAPTINVHB.Refresh()

        End If
    End Sub

    Sub Automatic_Distribution()
        If EntryMode = "N" Then
            If Absx1.cbeFor("INV_TYPE").Value = "A" Then
                'If Absx1.CtlFor("INV_TYPE").Text = "A" Then
                dst.Tables("APTINVH2").Rows.Clear()
            End If
            If dst.Tables("APTINVH2").Rows.Count = 0 And Val(numINV_AMT.Value & "") <> 0 Then
                Generate_Pre_Distribution()
                Calc_Totals()
            End If
        End If
        'Calculate_INV_DISC_AMT()
    End Sub

    Private Sub cmdNextCheckNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNextCheckNo.Click


        Dim rowGLTBANK1 As DataRow = Fill_Record("GLTBANK1", Absx1.txtFor("BANK_CODE").Text)
        If rowGLTBANK1 Is Nothing Then
            MsgBox("Invalid Bank Code", MsgBoxStyle.OkOnly, "Could Not Generate the Next Check")
            auto_next_check = False
            BANK_LAST_CHECK_NO = ""
            BANK_NEXT_CHECK_NO = ""
        Else
            BANK_LAST_CHECK_NO = rowGLTBANK1("BANK_LAST_CHECK_NO") & ""
            BANK_NEXT_CHECK_NO = CStr(Val(BANK_LAST_CHECK_NO) + 1)
            BANK_NEXT_CHECK_NO = ASCMAIN1.Format_Field(BANK_NEXT_CHECK_NO, "CHECK_NUM")
            Absx1.txtFor("CHECK_NUM").Text = BANK_NEXT_CHECK_NO
            auto_next_check = True
        End If

    End Sub

    Private Sub grdAPTINVHM_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTINVHM.AfterCellUpdate
        lblCOLUMN_NAME.Visible = True
        lblCOLUMN_NAME.Text = e.Cell.Column.Header.Caption
        lblNEW_VALUE.Visible = True
        lblNEW_VALUE.Text = e.Cell.Text
        lblCOLUMN_NAME.Tag = e.Cell.Column.Key
    End Sub

    Private Sub grdAPTINVHM_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTINVHM.ClickCellButton
        Dim sql_where As String = ""
        If e.Cell.Column.Key = "VEND_ALT_CODE" Then
            sql_where = "VEND_CODE = '" & Absx1.txtFor("VEND_CODE").Text & "'"
        End If
        grdClickCellButton(grdAPTINVHM, sql_where, sql_where <> "")
    End Sub

    Private Sub grdAPTINVHM_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdAPTINVHM.InitializeLayout
        For Each C As UltraWinGrid.UltraGridColumn In grdAPTINVHM.DisplayLayout.Bands(0).Columns
            If C.CellActivation = UltraWinGrid.Activation.NoEdit Then
                C.CellAppearance.BackColor = Drawing.Color.LightGray
            End If
        Next
    End Sub

    Sub Update_Multi()

        batch_update = True
        Click_Command("Cancel")
        chkQuickEntry.Checked = False
        Dim V As Integer = 0
        Dim X As CurrencyManager = Me.BindingContext(dst.Tables("APTINVH1"))
        X.EndCurrentEdit()

        Application.DoEvents()

        Dim STATS(2) As Integer
        For Each rowAPTINVHM As DataRow In dst.Tables("APTINVHM").Select("", "", DataViewRowState.ModifiedCurrent)
            STATS(0) += 1
            V = V + 1
            Absx1.txtFor("VOUCHER_NO").Text = rowAPTINVHM("VOUCHER_NO") & ""
            Application.DoEvents()
            X.EndCurrentEdit()

            Click_Command("Edit")
            Application.DoEvents()
            If EMsg <> "" Then
                rowAPTINVHM("UPDATE_STATUS") = "ERROR"
                rowAPTINVHM("UPDATE_MESSAGE") = EMsg
                STATS(2) += 1
            Else
                For c As Integer = 0 To rowAPTINVHM.ItemArray.Length - 1
                    Dim COLUMN_NAME As String = rowAPTINVHM.Table.Columns(c).ColumnName
                    If grdAPTINVHM.DisplayLayout.Bands(0).Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit Then
                        Absx1.txtFor(COLUMN_NAME).Text = rowAPTINVHM.Item(c) & ""
                    End If

                    'If rowAPTINVHM.Item(c, DataRowVersion.Current) & "" _
                    '<> rowAPTINVHM.Item(c, DataRowVersion.Original) & "" Then
                    '    Stop

                    '    If grdAPTINVHM.DisplayLayout.Bands(0).Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit Then
                    '        Stop
                    '    End If

                    '    Absx1.txtFor(COLUMN_NAME).Text = rowAPTINVHM.Item(c, DataRowVersion.Current) & ""
                    'End If
                Next

                Application.DoEvents()
                tabMain.SelectedTab = tabMain.Tabs("GL Distribution")
                Application.DoEvents()
                X.EndCurrentEdit()

                Click_Command("Update")
                Application.DoEvents()

                If EMsg <> "" Then
                    rowAPTINVHM("UPDATE_STATUS") = "ERROR"
                    rowAPTINVHM("UPDATE_MESSAGE") = EMsg
                    Click_Command("Cancel")
                    STATS(2) += 1
                Else
                    rowAPTINVHM("UPDATE_STATUS") = "UPDATED"
                    STATS(1) += 1
                End If
            End If
            Application.DoEvents()
        Next
        batch_update = False

        grdAPTINVHM.DisplayLayout.Bands("APTINVHM").Columns("UPDATE_STATUS").Hidden = False
        grdAPTINVHM.DisplayLayout.Bands("APTINVHM").Columns("UPDATE_MESSAGE").Hidden = False
        Export_to_Excel(grdAPTINVHM)
        grdAPTINVHM.DisplayLayout.Bands("APTINVHM").Columns("UPDATE_STATUS").Hidden = True
        grdAPTINVHM.DisplayLayout.Bands("APTINVHM").Columns("UPDATE_MESSAGE").Hidden = True

        MsgBox("Multiple-Invoice Edit Complete." & vbCr & vbCr & STATS(0) & " Records Processed" & vbCr & STATS(1) & " Records Updated Successfully" & vbCr & STATS(2) & " Records Were NOT Updated" & vbCr & vbCr & "Processing Errors (if any) will appear in the 'Update Message' Column of the Workbook", MsgBoxStyle.OkOnly, "Verification")

    End Sub

    Sub Copy_Change_to(ByVal Copy_to As String)
        If Copy_to = "All" Then
            For Each row As DataRow In dst.Tables("APTINVHM").Rows
                row.Item(lblCOLUMN_NAME.Tag) = lblNEW_VALUE.Text
            Next
        Else
            For Each C As UltraWinGrid.UltraGridRow In grdAPTINVHM.Selected.Rows
                C.Cells(lblCOLUMN_NAME.Tag).Value = lblNEW_VALUE.Text
                grdAPTINVHM.UpdateData()
            Next
        End If
    End Sub

    Private Sub txtOPS_YYYYPP_ACCRUE_EditorButtonClick(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinEditors.EditorButtonEventArgs) Handles txtOPS_YYYYPP_ACCRUE.EditorButtonClick
        txtOPS_YYYYPP_ACCRUE.ReadOnly = InquiryMode Or ApprovalMode
    End Sub

    Private Sub txtOPS_YYYYPP_ACCRUE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOPS_YYYYPP_ACCRUE.ValueChanged

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(
     ByVal ctl As Control,
     ByVal COLUMN_NAME As String,
     Optional ByRef sql_where As String = "",
     Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "OPS_YYYYPP_ACCRUE"
                If ROWs("APTPARM1").Item("AP_PARM_ALLOW_ACCRUAL") & "" = "1" Then
                    sql_where = "OPS_YYYYPP >= (SELECT GL_PARM_CURRENT_YYYYPP FROM GLTPARM1) and OPS_YYYYPP < '" & ASCMAIN1.CYP & "'"
                Else
                    sql_where = "OPS_YYYYPP > (SELECT GL_PARM_CURRENT_YYYYPP FROM GLTPARM1) and OPS_YYYYPP < '" & ASCMAIN1.CYP & "'"
                End If
                txtOPS_YYYYPP_ACCRUE.ReadOnly = True Or InquiryMode Or ApprovalMode

            Case "VEND_BUYER_CODE"
                sql_where = "VEND_BUYER_CODE in (Select USER_ID from ASTUSER1 where USER_STATUS = 'A')"

        End Select
    End Sub

#Region "grdAPTINVH5"

    Private Sub grdAPTINVH5_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTINVH5.AfterCellUpdate
        If e.Cell.Column.Key = "INV_COST" Or e.Cell.Column.Key = "INV_QTY" Or e.Cell.Column.Key = "CLOSE_LINE" Then

            Dim PO_COST As Decimal = Val(e.Cell.Row.Cells("PO_COST").Value & "")
            Dim INV_COST As Decimal = Val(e.Cell.Row.Cells("INV_COST").Value & "")
            Dim QTY_REC_NOT_INV As Int64 = Val(e.Cell.Row.Cells("QTY_REC_NOT_INV").Value & "")
            Dim INV_QTY As Int64 = Val(e.Cell.Row.Cells("INV_QTY").Value & "")
            Dim CLOSE_LINE As String = e.Cell.Row.Cells("CLOSE_LINE").Value & ""
            Dim AMT_VAR As Decimal = (INV_QTY * INV_COST) - IIf(CLOSE_LINE = "0", (INV_QTY * INV_COST), (QTY_REC_NOT_INV * PO_COST))
            If CLOSE_LINE = "0" Then
                'AMT_VAR = (QTY_REC_NOT_INV * PO_COST) - (INV_QTY * INV_COST)
                AMT_VAR = (INV_QTY * INV_COST) - (INV_QTY * INV_COST)
            Else
                'AMT_VAR = (QTY_REC_NOT_INV * PO_COST) - (INV_QTY * INV_COST)
                AMT_VAR = (INV_QTY * INV_COST) - (QTY_REC_NOT_INV * PO_COST)

            End If
            '  .Tables("APTINVH5").Columns.Add("AMT_VAR", GetType(System.Decimal), "AMT_INV - IIF(CLOSE_LINE='0', ISNULL(INV_QTY,0) * ISNULL(PO_COST,0), ISNULL(QTY_REC_NOT_INV,0) * ISNULL(PO_COST,0))")

            Dim AMT_VAR2 As Decimal = Val(e.Cell.Row.Cells("AMT_VAR").Value & "")

            'If AMT_VAR < 0 Then
            '    e.Cell.Row.Cells("CB").Value = "1"
            'ElseIf AMT_VAR >= 0 Then
            '    e.Cell.Row.Cells("CB").Value = "0"
            'End If

            If AMT_VAR > 0 Then
                If ROWs("APTPARM1").Item("AP_PARM_CB_VAR") & "" = "1" Then
                    e.Cell.Row.Cells("CB").Value = "1"
                Else
                    e.Cell.Row.Cells("CB").Value = "0"
                End If
            ElseIf AMT_VAR <= 0 Then
                e.Cell.Row.Cells("CB").Value = "0"
            End If
        End If

    End Sub

    Private Sub grdAPTINVH5_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTINVH5.AfterRowsDeleted

    End Sub

    Private Sub grdAPTINVH5_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdAPTINVH5.AfterRowUpdate
        Calc_DIST_PO()
    End Sub

    Private Sub grdAPTINVH5_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdAPTINVH5.BeforeRowUpdate

        Dim PO_COST As Decimal = Val(e.Row.Cells("PO_COST").Value & "")
        Dim INV_COST As Decimal = Val(e.Row.Cells("INV_COST").Value & "")
        Dim QTY_REC_NOT_INV As Int64 = Val(e.Row.Cells("QTY_REC_NOT_INV").Value & "")
        Dim INV_QTY As Int64 = Val(e.Row.Cells("INV_QTY").Value & "")
        Dim CLOSE_LINE As String = e.Row.Cells("CLOSE_LINE").Value & ""
        Dim AMT_VAR As Decimal
        If CLOSE_LINE = "0" Then
            'AMT_VAR = (QTY_REC_NOT_INV * PO_COST) - (INV_QTY * INV_COST)
            AMT_VAR = (INV_QTY * INV_COST) - (INV_QTY * INV_COST)
        Else
            'AMT_VAR = (QTY_REC_NOT_INV * PO_COST) - (INV_QTY * INV_COST)
            AMT_VAR = (INV_QTY * INV_COST) - (QTY_REC_NOT_INV * PO_COST)

        End If
        '  .Tables("APTINVH5").Columns.Add("AMT_VAR", GetType(System.Decimal), "AMT_INV - IIF(CLOSE_LINE='0', ISNULL(INV_QTY,0) * ISNULL(PO_COST,0), ISNULL(QTY_REC_NOT_INV,0) * ISNULL(PO_COST,0))")

        Dim AMT_VAR2 As Decimal = Val(e.Row.Cells("AMT_VAR").Value & "")

        'AMT_VAR = (QTY_REC * PO_COST) - (INV_QTY * INV_COST)
        'If AMT_VAR < 0 Then
        '    'e.Row.Cells("CB").Value = "1"
        'ElseIf AMT_VAR >= 0 Then
        '    e.Row.Cells("CB").Value = "0"
        'End If

        If AMT_VAR > 0 Then
            ' e.Row.Cells("CB").Value = "1"
        ElseIf AMT_VAR <= 0 Then
            ' e.Row.Cells("CB").Value = "0"
        End If

    End Sub

    Private Sub grdAPTINVH5_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdAPTINVH5.InitializeRow
        If Val(e.Row.Cells("QTY_REC_NOT_INV").Value & "") - Val(e.Row.Cells("INV_QTY").Value & "") <> 0 Then
            e.Row.Cells("INV_QTY").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("INV_QTY").Appearance.FontData.Bold = DefaultableBoolean.True
        Else
            e.Row.Cells("INV_QTY").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("INV_QTY").Appearance.FontData.Bold = DefaultableBoolean.False
        End If

        If Val(e.Row.Cells("PO_COST").Value & "") <> Val(e.Row.Cells("INV_COST").Value & "") Then
            e.Row.Cells("INV_COST").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("INV_COST").Appearance.FontData.Bold = DefaultableBoolean.True
        Else
            e.Row.Cells("INV_COST").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("INV_COST").Appearance.FontData.Bold = DefaultableBoolean.False
        End If

        Dim AMT_VAR As Decimal = Val(e.Row.Cells("AMT_VAR").Value & "")

        If AMT_VAR > 0 Then
            e.Row.Cells("CB").Appearance.BackColor = Drawing.Color.Red
            e.Row.Cells("CB").ToolTipText = "Unfavorable Variance"
            e.Row.Cells("QTY_VAR").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("AMT_VAR").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("QTY_VAR").Appearance.FontData.Bold = DefaultableBoolean.True
            e.Row.Cells("AMT_VAR").Appearance.FontData.Bold = DefaultableBoolean.True
        ElseIf AMT_VAR < 0 Then
            e.Row.Cells("CB").Appearance.BackColor = Drawing.Color.LightGreen
            e.Row.Cells("CB").ToolTipText = "Favorable Variance"
            e.Row.Cells("QTY_VAR").Appearance.ForeColor = Drawing.Color.DarkGreen
            e.Row.Cells("AMT_VAR").Appearance.ForeColor = Drawing.Color.DarkGreen
            e.Row.Cells("QTY_VAR").Appearance.FontData.Bold = DefaultableBoolean.True
            e.Row.Cells("AMT_VAR").Appearance.FontData.Bold = DefaultableBoolean.True
        Else
            e.Row.Cells("CB").Appearance.BackColor = Drawing.Color.Empty
            e.Row.Cells("CB").ToolTipText = ""
            e.Row.Cells("QTY_VAR").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("AMT_VAR").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("QTY_VAR").Appearance.FontData.Bold = DefaultableBoolean.False
            e.Row.Cells("AMT_VAR").Appearance.FontData.Bold = DefaultableBoolean.False
        End If
    End Sub

    Sub Create_APTINVH5_VAR()

        dst.Tables("APTINVH5_VAR").Rows.Clear()

        For Each rowAPTINVH5 As DataRow In dst.Tables("APTINVH5") _
        .Select("", "", DataViewRowState.CurrentRows)
            Dim COST_CATGY_CODE As String = rowAPTINVH5.Item("COST_CATGY_CODE")
            Dim rowICTCOST1 As DataRow = LookUp("ICTCOST1", COST_CATGY_CODE)

            Dim COLLECTION_CODE As String = rowAPTINVH5.Item("COLLECTION_CODE")
            Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)

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
                With rowAPTINVH5_VAR
                    .Item("COST_CATGY_CODE") = COST_CATGY_CODE
                    .Item("COLLECTION_CODE") = COLLECTION_CODE
                    .Item("ACCT_CODE_PPV") = rowAPTINVH5.Item("ACCT_CODE_PPV")
                    Dim rowICTIREC1 As DataRow = dst.Tables("ICTIREC1").Rows.Find(rowAPTINVH5.Item("RECEIPT_NO"))
                    .Item("SEG2_CODE") = SEG2_CODE
                    .Item("SEG3_CODE") = SEG3_CODE
                    .Item("SEG4_CODE") = SEG4_CODE
                End With
                dst.Tables("APTINVH5_VAR").Rows.Add(rowAPTINVH5_VAR)
            End If
            With rowAPTINVH5_VAR
                For Each C As String In New String() {"AMT_REC", "AMT_INV", "AMT_REC_NOT_INV", "AMT_REC_NOT_INV_OFFSET", "AMT_VAR"}
                    .Item(C) = Val(.Item(C) & "") + Val(rowAPTINVH5.Item(C) & "")
                Next
                If rowAPTINVH5.Item("CB") & "" = "1" Then
                    .Item("AMT_VAR_CB") = Val(.Item("AMT_VAR_CB") & "") + Val(rowAPTINVH5.Item("AMT_VAR") & "")
                End If
            End With
        Next
    End Sub

    Function CreateAPTINVH2(ByRef VOUCHER_LNO_ctr As Int32, INV_LTYP As String, ACCT_CODE As String, INV_LINE_AMT As Decimal, Optional COLLECTION_CODE As String = "") As DataRow

        Dim SEG4_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")

        If COLLECTION_CODE <> "" AndAlso ROWs("ICTPARM1").Item("IC_PARM_EXP_SEG4") & "" = "1" Then
            Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)
            If rowICTCOLL1.Item("SEG4_CODE") & "" <> "" Then
                SEG4_CODE = rowICTCOLL1.Item("SEG4_CODE")
            Else
                SEG4_CODE = COLLECTION_CODE
            End If
        End If

        Dim rowAPTINVH2 As DataRow = dst.Tables("APTINVH2").NewRow
        With rowAPTINVH2
            .Item("VOUCHER_NO") = HFs("VOUCHER_NO")
            VOUCHER_LNO_ctr += 1
            .Item("VOUCHER_LNO") = VOUCHER_LNO_ctr
            .Item("ACCT_CODE") = ACCT_CODE
            .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            .Item("SEG4_CODE") = SEG4_CODE ' ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
            .Item("INV_LINE_AMT") = INV_LINE_AMT
            .Item("INV_LTYP") = INV_LTYP
        End With
        dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)
        Return rowAPTINVH2
    End Function

    Sub Create_APTINVH2_P()

        Delete_Rows("APTINVH2", "INV_LTYP = 'P'")

        Dim VOUCHER_LNO_ctr As Int32 = Val(dst.Tables("APTINVH2").Compute("MAX(VOUCHER_LNO)", "VOUCHER_NO = '" & HFs("VOUCHER_NO") & "'") & "")

        Dim INV_LINE_AMT As Decimal = Val(dst.Tables("APTINVH5_VAR").Compute("SUM(AMT_REC_NOT_INV_OFFSET)", "") & "")
        Dim rowAPTINVH2 As DataRow

        If INV_LINE_AMT <> 0 Then
            rowAPTINVH2 = CreateAPTINVH2(VOUCHER_LNO_ctr, "P", ROWs("ICTPARM1").Item("IC_PARM_ACCT_CODE_ACCR_PURCH"), INV_LINE_AMT)
        End If

        For Each rowAPTINVH5_VAR As DataRow In dst.Tables("APTINVH5_VAR").Select("ISNULL(AMT_VAR,0) - ISNULL(AMT_VAR_CB,0) <> 0")
            Dim COLLECTION_CODE As String = rowAPTINVH5_VAR.Item("COLLECTION_CODE")
            INV_LINE_AMT = (Val(rowAPTINVH5_VAR.Item("AMT_VAR") & "") - Val(rowAPTINVH5_VAR.Item("AMT_VAR_CB") & ""))
            If INV_LINE_AMT <> 0 Then
                rowAPTINVH2 = CreateAPTINVH2(VOUCHER_LNO_ctr, "P", rowAPTINVH5_VAR.Item("ACCT_CODE_PPV"), INV_LINE_AMT, COLLECTION_CODE)
            End If
        Next
    End Sub

    Sub Create_APTINVH2_R()

        Delete_Rows("APTINVH2", "INV_LTYP = 'R'")

        Dim VOUCHER_LNO_ctr As Int32 = Val(dst.Tables("APTINVH2").Compute("MAX(VOUCHER_LNO)", "VOUCHER_NO = '" & HFs("VOUCHER_NO") & "'") & "")

        Dim INV_LINE_AMT As Decimal = -1 * Val(dst.Tables("APTINVH7").Compute("SUM(REBATE_USED)", "") & "")
        Dim rowAPTINVH2 As DataRow

        If INV_LINE_AMT <> 0 Then
            rowAPTINVH2 = CreateAPTINVH2(VOUCHER_LNO_ctr, "R", ROWs("PPTPARM1").Item("PP_PARM_REBATE_ACCRUAL"), INV_LINE_AMT)
        End If

        ' NO VARIANCE LOGIC YET - WE STILL HAVE TO DEAL WITH HOW TO HANDLE VARIANCES BETWEEN WHAT WAS ACCRUED AND WHAT WAS OFFERED BY THE VENDOR
    End Sub

    Sub Create_APTINVH2_from_APTINVH7()

        Delete_Rows("APTINVH2", "INV_LTYP = 'O'")

        Dim VOUCHER_LNO_ctr As Int32 = Val(dst.Tables("APTINVH2").Compute("MAX(VOUCHER_LNO)", "VOUCHER_NO = '" & HFs("VOUCHER_NO") & "'") & "")

        For Each rowAPTINVH7 As DataRow In dst.Tables("APTINVH7").Select("", "VOUCHER_CLNO")
            Dim TOTAL_INV As Decimal = Val(rowAPTINVH7.Item("TOTAL_INV") & "")

            Dim CTL_NO As String = rowAPTINVH7.Item("CTL_NO")
            Dim rowAPTACRC1 As DataRow = dst.Tables("APTACRC1").Rows.Find(CTL_NO)
            Dim COST_ACC As Decimal = Val(rowAPTACRC1.Item("COST_ACC") & "")
            Dim COST_ACT As Decimal = TOTAL_INV ' Val(rowAPTACRC1.Item("COST_ACT") & "")

            Dim ACCRUAL_CODE As String = rowAPTINVH7.Item("ACCRUAL_CODE")
            Dim rowAPTACRM1 As DataRow = dst.Tables("APTACRM1").Rows.Find(ACCRUAL_CODE)
            Dim ACCT_CODE_ACC As String = rowAPTACRM1.Item("ACCT_CODE_ACC") & ""
            Dim ACCT_CODE_EXP As String = rowAPTACRM1.Item("ACCT_CODE_EXP") & ""
            Dim COST_CATGY_CODE As String = rowAPTACRC1.Item("COST_CATGY_CODE") & ""
            If COST_CATGY_CODE <> "" Then
                Dim rowICTCOST1 As DataRow = dst.Tables("ICTCOST1").Rows.Find(COST_CATGY_CODE)
                ACCT_CODE_ACC = ROWs("ICTPARM1").Item("IC_PARM_ACCT_CODE_ACCR_LANDG")
                ACCT_CODE_EXP = rowICTCOST1.Item("ACCT_CODE_FPV")
            End If

            Dim CTL_TYPE As String = rowAPTINVH7.Item("CTL_TYPE") & ""
            Dim CTL_REF_NO As String = rowAPTINVH7.Item("CTL_REF_NO") & ""

            Dim rowAPTINVH2 As DataRow = Nothing

            If COST_ACC <> 0 Then
                rowAPTINVH2 = CreateAPTINVH2(VOUCHER_LNO_ctr, "O", ACCT_CODE_ACC, COST_ACC)
            End If

            If COST_ACT - COST_ACC <> 0 Then
                If CTL_TYPE = "M" Then '"chkPrePayment.Checked Then

                    VOUCHER_LNO_ctr = Val(dst.Tables("APTINVH2").Compute("MAX(VOUCHER_LNO)", "") & "")
                    rowAPTINVH2 = dst.Tables("APTINVH2").NewRow
                    With rowAPTINVH2
                        .Item("VOUCHER_NO") = HFs("VOUCHER_NO")
                        VOUCHER_LNO_ctr += 1
                        .Item("VOUCHER_LNO") = VOUCHER_LNO_ctr
                        .Item("ACCT_CODE") = ROWs("ICTPARM1").Item("IC_PARM_ACCT_CODE_ACCR_TOOLG")
                        .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                        .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                        .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                        .Item("ACCT_DESC") = LookUp("GLTACCT1", rowAPTINVH2.Item("ACCT_CODE") & "", True).ITEM("ACCT_DESC")
                        .Item("INV_LINE_AMT") = COST_ACT
                        .Item("INV_COMMENT_DTL") = CTL_REF_NO ' txtPPDDutyBOL.Text
                        .Item("INV_LTYP") = "O" ' WAS "A" - BUT I DON'T KNOW WHY - A CAUSES EDITS TO GO BONKERS - A IS FOR ADVANCES
                    End With
                    dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)

                Else
                    Dim COLLECTION_CODE As String = rowAPTINVH7.Item("COLLECTION_CODE")
                    rowAPTINVH2 = CreateAPTINVH2(VOUCHER_LNO_ctr, "O", ACCT_CODE_EXP, COST_ACT - COST_ACC, COLLECTION_CODE)
                End If
            End If
        Next
    End Sub
#End Region

    Private Sub grdICTIREC1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTIREC1.ClickCellButton
        If InquiryMode Or ApprovalMode Then Exit Sub

        If e.Cell.Row.IsDataRow Then
            ProcessDoubleClickedRow()
        End If
    End Sub

    Private Sub ProcessDoubleClickedRow()

        If InquiryMode Or ApprovalMode Then Exit Sub

        Dim RECEIPT_NO As String = grdICTIREC1.ActiveRow.Cells("RECEIPT_NO").Text

        If dst.Tables("APTINVH5_SUM").Rows.Find(RECEIPT_NO) Is Nothing Then
            Me.Cursor = Cursors.WaitCursor
            Application.DoEvents()

            Dim APTINVH5_SUM As DataRow = dst.Tables("APTINVH5_SUM").NewRow
            APTINVH5_SUM.Item("RECEIPT_NO") = RECEIPT_NO
            APTINVH5_SUM.Item("RECEIPT_DATE") = grdICTIREC1.ActiveRow.Cells("RECEIPT_DATE").Text
            APTINVH5_SUM.Item("PO_ORDER_NO") = grdICTIREC1.ActiveRow.Cells("PO_ORDER_NO").Text
            APTINVH5_SUM.Item("QTY_REC") = Val(grdICTIREC1.ActiveRow.Cells("QTY_REC").Text)
            APTINVH5_SUM.Item("AMT_REC") = Val(grdICTIREC1.ActiveRow.Cells("AMT_REC").Text)
            APTINVH5_SUM.Item("QTY_INV") = Val(grdICTIREC1.ActiveRow.Cells("QTY_REC").Text)
            APTINVH5_SUM.Item("AMT_INV") = Val(grdICTIREC1.ActiveRow.Cells("AMT_REC").Text)
            dst.Tables("APTINVH5_SUM").Rows.Add(APTINVH5_SUM)

            Dim VOUCHER_DLNO_max As Integer = Val(dst.Tables("APTINVH5").Compute("MAX(VOUCHER_DLNO)", "") & "")


            'Dim sql As String = "Select '" & HFs("VOUCHER_NO") & "' VOUCHER_NO " & vbCrLf _
            '    & ", ICTIREC2.RECEIPT_LNO + " & CStr(VOUCHER_DLNO_max) & " VOUCHER_DLNO " & vbCrLf _
            '    & ", ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO" & vbCrLf _
            '    & ", DECODE(ICTIREC2.ACCRUAL_STATUS,'1',0,NVL(ICTIREC2.QTY_REC,0) - NVL(ICTIREC2.QTY_INV,0)) INV_QTY" & vbCrLf _
            '    & ", ICTIREC2.PO_COST INV_COST" & vbCrLf _
            '    & ", '0' CB" & vbCrLf _
            '    & ", ICTITEM1.COST_CATGY_CODE, '1' CLOSE_LINE" & vbCrLf _
            '    & ", DECODE(ICTIREC2.ACCRUAL_STATUS,'1',0,NVL(ICTIREC2.QTY_REC,0) - NVL(ICTIREC2.QTY_INV,0)) QTY_REC_NOT_INV" & vbCrLf _
            '    & ", ICTITEM1.ITEM_DESC, ICTITEM1.COLLECTION_CODE, ICTCOST1.ACCT_CODE_PPV" & vbCrLf _
            '    & ", ICTIREC2.QTY_REC, ICTIREC2.QTY_INV, ICTIREC2.PO_COST " & vbCrLf _
            '    & ", ICTIREC2.ITEM_CODE, ICTIREC2.ITEM_UOM" & vbCrLf _
            '     & ", NVL(ICTIREC2.QTY_REC,0) * NVL(ICTCOSTA.ITEM_COST_VCOST,0) VCOST" & vbCrLf _
            '     & ", NVL(ICTIREC2.QTY_REC,0) * NVL(ICTCOSTA.ITEM_COST_LANDG,0) LANDG" & vbCrLf _
            '     & ", NVL(ICTIREC2.QTY_REC,0) * NVL(ICTCOSTA.ITEM_COST_TOOLG,0) TOOLG" & vbCrLf _
            '     & ", NVL(ICTIREC2.QTY_REC,0) * NVL(ICTCOSTA.ITEM_COST_OVRHD,0) OVRHD" & vbCrLf _
            '    & " from ICTIREC2,ICTITEM1,ICTCOST1,ICTCOSTA" & vbCrLf _
            '    & " where ICTIREC2.RECEIPT_NO = '" & RECEIPT_NO & "'" & vbCrLf _
            '    & "   and ICTCOSTA.OPS_YYYYPP = ICTIREC2.OPS_YYYYPP" & vbCrLf _
            '    & "   and ICTCOSTA.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf _
            '    & " and ICTITEM1.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf _
            '    & " and ICTCOST1.COST_CATGY_CODE (+) = ICTITEM1.COST_CATGY_CODE"

            Dim sql As String = "Select '" & HFs("VOUCHER_NO") & "' VOUCHER_NO " & vbCrLf _
                & ", ICTIREC2.RECEIPT_LNO + " & CStr(VOUCHER_DLNO_max) & " VOUCHER_DLNO " & vbCrLf _
                & ", ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO" & vbCrLf _
                & ", CASE WHEN APTINVH5.VOUCHER_NO IS NULL THEN DECODE(ICTIREC2.ACCRUAL_STATUS,'1',0,NVL(ICTIREC2.QTY_REC,0) - NVL(ICTIREC2.QTY_INV,0)) ELSE NVL(ICTIREC2.QTY_REC,0) - NVL(ICTIREC2.QTY_INV,0) + NVL(APTINVH5.INV_QTY,0) END INV_QTY" & vbCrLf _
                & ", ICTIREC2.PO_COST INV_COST" & vbCrLf _
                & ", '0' CB" & vbCrLf _
                & ", ICTITEM1.COST_CATGY_CODE, '1' CLOSE_LINE" & vbCrLf _
                & ", CASE WHEN APTINVH5.VOUCHER_NO IS NULL THEN DECODE(ICTIREC2.ACCRUAL_STATUS,'1',0,NVL(ICTIREC2.QTY_REC,0) - NVL(ICTIREC2.QTY_INV,0)) ELSE NVL(ICTIREC2.QTY_REC,0) - NVL(ICTIREC2.QTY_INV,0) + NVL(APTINVH5.INV_QTY,0) END QTY_REC_NOT_INV" & vbCrLf _
                & ", ICTITEM1.ITEM_DESC, ICTITEM1.COLLECTION_CODE, ICTCOST1.ACCT_CODE_PPV" & vbCrLf _
                & ", ICTIREC2.QTY_REC, ICTIREC2.QTY_INV, ICTIREC2.PO_COST " & vbCrLf _
                & ", ICTIREC2.ITEM_CODE, ICTIREC2.ITEM_UOM" & vbCrLf _
                & ", NVL(ICTIREC2.QTY_REC,0) * NVL(ICTCOSTA.ITEM_COST_VCOST,0) VCOST" & vbCrLf _
                & ", NVL(ICTIREC2.QTY_REC,0) * NVL(ICTCOSTA.ITEM_COST_LANDG,0) LANDG" & vbCrLf _
                & ", NVL(ICTIREC2.QTY_REC,0) * NVL(ICTCOSTA.ITEM_COST_TOOLG,0) TOOLG" & vbCrLf _
                & ", NVL(ICTIREC2.QTY_REC,0) * NVL(ICTCOSTA.ITEM_COST_OVRHD,0) OVRHD" & vbCrLf _
                & " from ICTIREC2,ICTITEM1,ICTCOST1,ICTCOSTA,APTINVH5" & vbCrLf _
                & " where ICTIREC2.RECEIPT_NO = '" & RECEIPT_NO & "'" & vbCrLf _
                & "   and ICTCOSTA.OPS_YYYYPP = ICTIREC2.OPS_YYYYPP" & vbCrLf _
                & "   and ICTCOSTA.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf _
                & " and ICTITEM1.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf _
                & " and ICTCOST1.COST_CATGY_CODE (+) = ICTITEM1.COST_CATGY_CODE" & vbCrLf _
                & " and APTINVH5.VOUCHER_NO (+) = '" & HFs("VOUCHER_NO") & "'" & vbCrLf _
                & " and APTINVH5.RECEIPT_NO (+) = ICTIREC2.RECEIPT_NO" & vbCrLf _
                & " and APTINVH5.RECEIPT_LNO (+) = ICTIREC2.RECEIPT_LNO"
            Fill_Records("APTINVH5", , False, sql)

            '                & " and ICTIREC2.ACCRUAL_STATUS = '0'" & vbCrLf _

            grdICTIREC1.ActiveRow.Cells("ACCRUAL_STATUS").Value = "2"
            grdICTIREC1.UpdateData()

            'dst.Tables("APTINVH5_SUM").Rows.Find("RECEIPT_NO = '" & RECEIPT_NO & "'")
            grdAPTINVH5_SUM.Rows(grdAPTINVH5_SUM.Rows.Count - 1).Activate()

            'dteINV_BL_DATE.Value = grdICTIREC1.ActiveRow.Cells("DATE_RECEIVED").Value
            Setup_INV_BL_DATE()
            Calculate_INV_DUE_DATE()
            Calc_DIST_PO()
            'Display_Totals()
            Me.Cursor = Cursors.Default
            grdAPTINVH5.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        Else
            MsgBox("Receipt " & RECEIPT_NO & " is already part of this Voucher", MsgBoxStyle.OkOnly, "Duplicate Selection")
        End If
    End Sub

    Private Sub numINV_AMT_VEND_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles numINV_AMT_VEND.Leave
        Automatic_Distribution()
    End Sub

    Private Sub numINV_AMT_VEND_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numINV_AMT_VEND.ValueChanged
        Calc_Totals()
    End Sub

    Private Sub grdAPTINVH8_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTINVH8.AfterRowsDeleted
        Calc_DIST_Adjustments()
    End Sub

    Private Sub grdAPTINVH8_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdAPTINVH8.AfterRowUpdate
        Calc_DIST_Adjustments()
    End Sub

    Private Sub grdAPTINVH8_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdAPTINVH8.BeforeRowUpdate
        With grdAPTINVH8
            If Not e.Cancel Then
                If e.Row.Cells("VOUCHER_NO").Text = "" Then
                    .ActiveRow.Cells("VOUCHER_NO").Value = Absx1.CtlFor("VOUCHER_NO").Text
                    .ActiveRow.Cells("VOUCHER_ANO").Value = Val(dst.Tables("APTINVH8").Compute("Max(VOUCHER_ANO)", "") & "") + 1
                End If
            End If
        End With
    End Sub

    Private Sub cmdCreateAccrual_Click(sender As System.Object, e As System.EventArgs) Handles cmdCreateAccrual.Click

        ' THIS PROCESS IS BEING OPTIMIZED FOR AHA WHERE A NEG ACCRUAL (IE, A DR TO THE BALANCE SHEET) IS BEING CREATED
        ' NORMALLY THIS BUTTON IS TO CREATE A $0 ACCRUAL, AND THEN AN EXPENSE

        Dim EMsg As String = ""
        Dim rowAPTACRM1 As DataRow = Nothing

        Dim ACCRUAL_CODE As String = Absx1.txtFor("ACCRUAL_CODE").Text
        If ACCRUAL_CODE = "" Then
            EMsg &= vbCr & "No Value Specified for Accrual Code"
        Else
            rowAPTACRM1 = LookUp("APTACRM1", ACCRUAL_CODE)
            If rowAPTACRM1 Is Nothing Then
                EMsg &= vbCr & "Invalid Value Specified for Accrual Code"
            End If
        End If

        If chkPrePayment.Checked Then
            If Absx1.txtFor("VEND_CODE_ACC").Text = "" Then
                Absx1.txtFor("VEND_CODE_ACC").Text = HFs("VEND_CODE")
            End If
        Else
            Dim VEND_CODE_ACC As String = Absx1.txtFor("VEND_CODE_ACC").Text
            If VEND_CODE_ACC = "" Then
                VEND_CODE_ACC = Absx1.txtFor("VEND_CODE").Text
                ' EMsg &= vbCr & "No Value Specified for Vendor Code"
            Else
                rowAPTVEND1 = LookUp("APTVEND1", VEND_CODE_ACC)
                If rowAPTVEND1 Is Nothing Then
                    EMsg &= vbCr & "Invalid Value Specified for Vendor Code"
                End If
            End If
        End If



        ' THERE ARE SEVERAL ISSUES WITH THIS BLOCK OF CODE
        ' AND THAT IS WHY IT IS PROBABLY TRUE THAT YOU CANNOT CREATE $0 ACCRUAL RECORDS UNLESS YOU CHECK PREPAYMENT
        ' NOTE THE COMMENTS WITH ***
        If chkPrePayment.Checked Then
            Absx1.txtFor("CTL_NOTE").Text = "Pre-Paid"
            Absx1.dteFor("CTL_DATE").Value = Now.Date
            Dim CTL_REF_NO As String = Trim(txtPPDDutyBOL.Text) ' Trim(Absx1.txtFor("CTL_NOTE").Text)
            If CTL_REF_NO = "" Then
                EMsg &= vbCr & "No Value Specified for Reference No"
            Else
                If ACCRUAL_CODE = "TRF" Then
                    ' CHECK FOR ALREADY MATCHED RECORDS WITH THIS BOL
                    ASCMAIN1.sql = "Select * From APTACRC0 where SOURCE_DOC_NO = :PARM1"
                    Dim rowAPTACRC0 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, False, "V", New String() {txtPPDDutyBOL.Text})
                    If rowAPTACRC0 IsNot Nothing Then
                        If MsgBox("This Pre-Payment will be Merged with an already Matched Entry" & vbCrLf & vbCrLf & "OK To continue?", MsgBoxStyle.YesNo, "Verification - this BOL has already been Matched") = MsgBoxResult.No Then
                            EMsg &= vbCr & "Duplicate entry with an Already Matched BOL - Merged Declined"
                        End If
                    End If
                End If
            End If
            ' *** CTL_REF_NO IS NOT EVEN USED BELOW THIS LINE
        Else
            Dim CTL_NOTE As String = Trim(Absx1.txtFor("CTL_NOTE").Text)
            If CTL_NOTE = "" Then
                EMsg &= vbCr & "No Value Specified for Note"
            End If

            If txtOA_PO_ORDER_NO.Text = "" Or txtOA_PO_ORDER_LNO.Text = "" Or txtOA_RECEIPT_NO.Text = "" Or txtOA_RECEIPT_LNO.Text = "" Then
                EMsg &= vbCr & "No PO Receipt Specified"
            End If
        End If

            If numCOST_ACT.Value = 0 Then
            EMsg &= vbCr & "Invalid Value Specified for Actual Cost"
        End If

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Create Accrual")
            Exit Sub
        End If

        Dim rowAPTACRC1 As DataRow = dst.Tables("APTACRC1").NewRow
        Dim CTL_NO As String = ASCMAIN1.Next_Control_No("APTACRC1.CTL_NO")
        With rowAPTACRC1
            .Item("CTL_NO") = CTL_NO
            .Item("VEND_CODE_ACC") = Absx1.txtFor("VEND_CODE_ACC").Text
            .Item("ACCRUAL_CODE") = ACCRUAL_CODE
            .Item("CHARGEBACK_IND") = rowAPTACRM1.Item("CHARGEBACK_IND")
            .Item("CTL_DATE") = Absx1.dteFor("CTL_DATE").Value
            .Item("CTL_NOTE") = Absx1.txtFor("CTL_NOTE").Text
            .Item("VOUCHER_NO_ORIG") = Absx1.txtFor("VOUCHER_NO").Text
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("CTL_STATUS") = "0"
            .Item("INV_PRINT_IND") = "0"
            If chkPrePayment.Checked Then
            Else
                .Item("PO_ORDER_NO") = txtOA_PO_ORDER_NO.Text
                .Item("PO_ORDER_LNO") = txtOA_PO_ORDER_LNO.Text
                .Item("RECEIPT_NO") = txtOA_RECEIPT_NO.Text
                .Item("RECEIPT_LNO") = txtOA_RECEIPT_LNO.Text
                .Item("ITEM_CODE") = txtOA_ITEM_CODE.Text

                Dim rowICTIREC1 As DataRow = LookUp("ICTIREC1", .Item("RECEIPT_NO"))
                Dim rowICTIREC2 As DataRow = LookUp("ICTIREC2", New String() { .Item("RECEIPT_NO"), .Item("RECEIPT_LNO")})
                .Item("COST_CATGY_CODE") = rowICTIREC2.Item("COST_CATGY_CODE")
                .Item("SOURCE_DOC_NO") = rowICTIREC1.Item("SOURCE_DOC_NO")
            End If

            .Item("OPS_YYYYPP") = ASCMAIN1.CYP

            ' *** SOME OF THESE FIELDS SHOULD BE SET AS SHOWN BELOW ONLY IF THIS IS A PREPAYMENT
            .Item("COST_ACT") = Val(numCOST_ACT.Value & "")
            .Item("COST_ORIG") = Val(numCOST_ACT.Value & "")
            .Item("CTL_TYPE") = "M" ' Manual
            .Item("PPD_IND") = "1" ' Pre-Paid
            .Item("SOURCE_DOC_NO") = txtPPDDutyBOL.Text.Trim
        End With
        dst.Tables("APTACRC1").Rows.Add(rowAPTACRC1)

        Add_APTINVH7(CTL_NO, True)

        Clear_Other_Accrual_Controls()
    End Sub

    Sub Clear_Other_Accrual_Controls()

        'Absx1.txtFor("ACCRUAL_CODE").Text = ""
        'Absx1.txtFor("VEND_CODE_ACC").Text = ""
        'If Absx1.dteFor("INV_DATE").Value & "" = "" Then
        '    Absx1.dteFor("CTL_DATE").Value = Now.Date
        'Else
        '    Absx1.dteFor("CTL_DATE").Value = Absx1.dteFor("INV_DATE").Value
        'End If

        'Absx1.txtFor("CTL_NOTE").Text = ""

        numCOST_ACT.Value = 0
        txtOA_ACCRUAL_CODE.Text = ""
        dteOA_CTL_DATE.Value = DBNull.Value
        txtOA_VEND_CODE_ACC.Text = ""
        txtOA_PO_ORDER_NO.Text = ""
        txtOA_PO_ORDER_LNO.Text = ""
        txtOA_RECEIPT_NO.Text = ""
        txtOA_RECEIPT_LNO.Text = ""
        txtOA_ITEM_CODE.Text = ""
        txtOA_CTL_NOTE.Text = ""

        txtPPDDutyBOL.Text = ""
    End Sub

    Private Sub ProcessDoubleClickedRow_APTACRC1(CTL_NO As String, Optional RECEIPT_NO As String = "",
                                                 Optional IncludeOthersAutomatically As Boolean = False,
                                                 Optional IncludeOthersAtZero As Boolean = False)
        If InquiryMode Then Exit Sub

        If dst.Tables("APTINVH7").Select("CTL_NO = '" & CTL_NO & "'").Length = 0 Then
            Add_APTINVH7(CTL_NO)
            ' IF THIS IS THE ONLY DETOUTS0 RECORD FOR THIS CTL_NO THAT IS VOUCHERED, AND THERE ARE OTHER UNVOUCHERED DETOUTS0 RECORDS FOR THIS CTL_NO, OFFER THEM UP
            If RECEIPT_NO <> "" Then
                If dst.Tables("APTACRC1").Select("VOUCHER_NO = '" & HFs("VOUCHER_NO") & "' and RECEIPT_NO = '" & RECEIPT_NO & "'").Length = 1 _
                    And dst.Tables("APTACRC1").Select("VOUCHER_NO IS NULL and RECEIPT_NO = '" & RECEIPT_NO & "'").Length > 0 Then
                    If IncludeOthersAutomatically OrElse
                        MsgBox("There are other Freight Accruals on this Receipt.  Do you want them included along with the one that you Selected?", MsgBoxStyle.YesNo, "Verification") = vbYes Then
                        For Each row As DataRow In dst.Tables("APTACRC1").Select("VOUCHER_NO IS NULL and RECEIPT_NO = '" & RECEIPT_NO & "'")
                            CTL_NO = row.Item("CTL_NO")
                            'Dim COST_INV2 As Decimal = Val(row.Item("COST_ACC_OPEN") & "")
                            'If IncludeOthersAtZero Then COST_INV2 = 0
                            ProcessDoubleClickedRow_APTACRC1(CTL_NO)
                        Next
                    End If
                End If
            End If
        Else
            MsgBox("Other Accrual " & CTL_NO & " is already part of this Voucher",
                   MsgBoxStyle.OkOnly, "Duplicate Selection")
        End If
    End Sub


    Sub Add_APTINVH7(CTL_NO As String, Optional created As Boolean = False)

        Dim rowAPTACRC1 As DataRow = dst.Tables("APTACRC1").Rows.Find(CTL_NO)
        Dim rowAPINVH7() As DataRow = dst.Tables("APTINVH7").Select("CTL_NO = '" & CTL_NO & "'")
        If rowAPINVH7.Length <> 0 Then
            MsgBox("Ctl No " & CTL_NO & " is already added to this Voucher", MsgBoxStyle.OkOnly, "Cannot add same records twice")
            Exit Sub
        End If

        rowAPTACRC1.Item("VOUCHER_NO") = HFs("VOUCHER_NO")

        Dim rowAPTINVH7 As DataRow = dst.Tables("APTINVH7").NewRow
        With rowAPTINVH7
            .Item("VOUCHER_NO") = HFs("VOUCHER_NO")
            Dim VOUCHER_CLNO As Integer = Val(dst.Tables("APTINVH7").Compute("MAX(VOUCHER_CLNO)", "") & "") + 1
            .Item("VOUCHER_CLNO") = VOUCHER_CLNO
            .Item("CTL_NO") = CTL_NO

            For Each C As String In New String() {"ACCRUAL_CODE", "CHARGEBACK_IND", "PO_ORDER_NO", "RECEIPT_NO", "ITEM_CODE", "CTL_TYPE", "COST_CATGY_CODE", "CTL_DATE", "CTL_NOTE"}
                .Item(C) = rowAPTACRC1.Item(C)
            Next

            If rowAPTACRC1.Item("CTL_TYPE") & "" = "M" And rowAPTACRC1.Item("PPD_IND") & "" = "1" Then
                ' Set Accrual to -1 * PPD Amount for Matching
                .Item("TOTAL_ACC") = -1 * Val(rowAPTACRC1.Item("COST_ORIG") & "")
            Else
                .Item("TOTAL_ACC") = rowAPTACRC1.Item("COST_ACC")
            End If

            Dim rowAPTACRM1 As DataRow = LookUp("APTACRM1", rowAPTACRC1.Item("ACCRUAL_CODE"))
            .Item("ACCRUAL_DESC") = rowAPTACRM1.Item("ACCRUAL_DESC")

            If created Then
                .Item("TOTAL_INV") = rowAPTACRC1.Item("COST_ACT")
            Else
                If rowAPTACRC1.Item("CTL_TYPE") & "" = "M" And rowAPTACRC1.Item("PPD_IND") & "" = "1" Then
                    .Item("TOTAL_INV") = -1 * Val(rowAPTACRC1.Item("COST_ORIG") & "")
                Else
                    .Item("TOTAL_INV") = Val(rowAPTACRC1.Item("COST_ACC") & "") - Val(rowAPTACRC1.Item("COST_ACT") & "")

                End If
            End If

            If chkPrePayment.Checked Then
                ' rowAPTACRC1.Item("COST_ACT") = rowAPTACRC1.Item("COST_ACT")
            Else
                ' NOT SURE ABOUT THIS NEXT LINE - NEED TO TEST
                rowAPTACRC1.Item("COST_ACT") = rowAPTACRC1.Item("COST_ACC")
            End If


            If created Then
                .Item("CTL_NO_CREATED") = "1"
            End If

            .Item("CTL_REF_NO") = Trim(txtPPDDutyBOL.Text)

            If chkPrePayment.Checked Then

            Else
                Dim ITEM_CODE As String = rowAPTACRC1.Item("ITEM_CODE") & ""
                If ITEM_CODE <> "" Then
                    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                    .Item("COLLECTION_CODE") = rowICTITEM1.Item("COLLECTION_CODE")
                End If
            End If

        End With
        dst.Tables("APTINVH7").Rows.Add(rowAPTINVH7)
        Calc_DIST_Other()
    End Sub

    Private Sub grdAPTACRC1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdAPTACRC1.DoubleClickRow
        If e.Row IsNot Nothing AndAlso e.Row.IsDataRow Then
            ProcessDoubleClickedRow_APTACRC1(e.Row.Cells("CTL_NO").Value & "", e.Row.Cells("RECEIPT_NO").Value & "", False)
            'Add_APTINVH7(e.Row.Cells("CTL_NO").Value)
        End If
    End Sub

    Private Sub grdAPTACRC1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdAPTACRC1.InitializeRow
        If e.Row.Cells("VOUCHER_NO").Value & "" = HFs("VOUCHER_NO") Then
            e.Row.CellAppearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.CellAppearance.BackColor = Drawing.Color.Empty
        End If
        If e.Row.Cells("VEND_CODE_ACC").Value & "" = HFs("VEND_CODE") Then
            e.Row.Cells("VEND_CODE_ACC").Appearance.ForeColor = Drawing.Color.Empty
        Else
            e.Row.Cells("VEND_CODE_ACC").Appearance.ForeColor = Drawing.Color.Red
        End If
    End Sub
#Region "grdAPTINVH7"

    Private Sub grdAPTINVH7_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdAPTINVH7.AfterRowsDeleted

        Dim CTL_NOs As List(Of String) = DirectCast(grdAPTINVH7.Tag, List(Of String))
        For Each CTL_NO As String In CTL_NOs
            Dim rowAPTACRC1 As DataRow = dst.Tables("APTACRC1").Rows.Find(CTL_NO)
            If rowAPTACRC1.RowState = DataRowState.Added Then
                rowAPTACRC1.Delete()
            Else
                rowAPTACRC1.Item("VOUCHER_NO") = ""
                rowAPTACRC1.Item("CTL_STATUS") = "0"
                rowAPTACRC1.Item("COST_ACT") = 0
            End If
        Next
        Calc_DIST_Other()
    End Sub

    Private Sub grdAPTINVH7_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdAPTINVH7.AfterRowUpdate
        Dim CTL_NO As String = e.Row.Cells("CTL_NO").Value
        Dim rowAPTACRC1 As DataRow = dst.Tables("APTACRC1").Rows.Find(CTL_NO)

        If rowAPTACRC1.Item("PPD_IND") & "" = "1" Then Exit Sub

        rowAPTACRC1.Item("COST_ACT") = e.Row.Cells("TOTAL_INV").Value
        Calc_DIST_Other()
    End Sub

    Private Sub grdAPTINVH7_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdAPTINVH7.BeforeRowsDeleted
        Dim CTL_NOs As New List(Of String)
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            Dim CTL_NO As String = grow.Cells("CTL_NO").Value
            CTL_NOs.Add(CTL_NO)
        Next
        grdAPTINVH7.Tag = CTL_NOs
    End Sub

#End Region
    Sub Print_Check(COPY As String)
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Check")

        Dim REPORTFILE As String = "APRCHKP1"
        Dim RPT As String = ""
        Dim rowGLTBANK1 As DataRow = LookUp("GLTBANK1", rowAPTINVH1.Item("BANK_CODE"))
        If rowGLTBANK1.Item("CHECK_REPORT") & "" <> "" Then
            RPT = rowGLTBANK1.Item("CHECK_REPORT")
        End If
        If RPT = "" Then RPT = REPORTFILE

        If Not REPORTS.ContainsKey(REPORTFILE) Then
            REPORTS.Add(REPORTFILE, Load_rptClass(REPORTFILE))
            REPORTS(REPORTFILE).Prepare_dst(False, "")
        End If

        'To fill the report's dataset with data from Oracle, 
        ' set the parameter array to values that the Fill_Records_RPT method expects, and then call it

        REPORTS(REPORTFILE).Fill_Records_RPT(New String() {"", rowAPTINVH1.Item("BANK_CODE"), rowAPTINVH1.Item("CHECK_NUM")})
        Dim REPORT_NO As String = ""

        Dim make_pdf As Boolean = (COPY = "1")
        Dim FILENAME_body As String = ""

        With REPORTS(REPORTFILE).clsASCBASE1
            .Print_Report_Begin()
            .CR_params.Add("SUBT", "")
            .CR_params.Add("COPY", COPY)
            'If make_pdf Then
            'REPORT_NO = .Generate_Report(RPT, "Check", , True, , , "PDF", FILENAME_body, False)
            'Show_Document(FILENAME_body)
            'Else
            REPORT_NO = .Generate_Report(RPT, "Check", , True, , , , , False)
            'End If
            '   .Print_Report_End(make_pdf, make_pdf)
            .Print_Report_End(False, False)
        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub cmdCheck_Click(sender As System.Object, e As System.EventArgs) Handles cmdCheck.Click
        Print_Check("1")
    End Sub

    Private Sub grdAPTINVR1_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdAPTINVR1.InitializeRow
        If e.Row.IsDataRow AndAlso e.Row.Band.Key = "APTINVR1" Then
            If e.Row.Cells("INV_APPR_STATUS").Value & "" = "P" And e.Row.Cells("INV_STATUS").Value = "O" Then
                e.Row.Cells("INV_STATUS").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("INV_STATUS").ToolTipText = "Invoice is Pending Approval"
            Else
                e.Row.Cells("INV_STATUS").Appearance.ForeColor = Drawing.Color.Empty
                e.Row.Cells("INV_STATUS").ToolTipText = ""

            End If
        End If
    End Sub

    Private Sub cmbVEND_BUYER_CODE_ValueChanged(sender As Object, e As EventArgs) Handles cmbVEND_BUYER_CODE.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Show_Batch()
    End Sub

    Private Sub cbeAPPR_ROUTE_TO_KeyDown(sender As Object, e As KeyEventArgs) Handles cbeAPPR_ROUTE_TO.KeyDown
        If e.KeyCode = Keys.Delete Then
            cbeAPPR_ROUTE_TO.Value = DBNull.Value
        End If
    End Sub

    Private Sub optINV_APPR_STATUS_ValueChanged(sender As Object, e As EventArgs) Handles optAPPR_ACTION.ValueChanged
        cbeAPPR_ROUTE_TO.Visible = (optAPPR_ACTION.Value & "" = "P" Or optAPPR_ACTION.Value & "" = "R")
        lblAPPR_ROUTE_TO.Visible = cbeAPPR_ROUTE_TO.Visible
    End Sub

    Sub Clear_Out_Accrued_AP()

        If ASCMAIN1.Running_in_VS Then
        Else
            Stop
        End If

        ASCMAIN1.sql = "Select Distinct VEND_CODE from WJZ_2051"
        For Each rowV As DataRow In ASCDATA1.GetDataTable.Select("", "VEND_CODE")
            Dim VEND_CODE As String = rowV.Item("VEND_CODE")
            Absx1.txtFor("VEND_CODE").Text = VEND_CODE
            Absx1.txtFor("INV_NUM").Text = "2051 CLEANUP 201512"

            Click_Command("New")

            Absx1.dteFor("INV_DATE").Value = CDate("12/31/2015")
            Absx1.numFor("INV_AMT_VEND").Value = 0
            tabMain.Tabs("PO Receipts").Selected = True

            Dim RECEIPT_NOs As New List(Of String)
            ASCMAIN1.sql = "Select Distinct RECEIPT_NO from WJZ_2051 where VEND_CODE = '" & VEND_CODE & "'"
            For Each rowR As DataRow In ASCDATA1.GetDataTable().Select("", "RECEIPT_NO")
                Dim RECEIPT_NO As String = rowR.Item("RECEIPT_NO")
                RECEIPT_NOs.Add(RECEIPT_NO)
            Next

            For Each grow As UltraWinGrid.UltraGridRow In grdICTIREC1.Rows
                If grow.Band.Key = "ICTIREC1" Then
                    Dim RECEIPT_NO As String = grow.Cells("RECEIPT_NO").Value
                    If RECEIPT_NOs.Contains(RECEIPT_NO) Then
                        grow.Activate()
                        ProcessDoubleClickedRow()
                    End If
                End If
            Next

            Dim TOTAL As Decimal = Val(dst.Tables("APTINVH5_SUM").Compute("SUM(AMT_INV)", "") & "")
            Dim rowG As DataRow = dst.Tables("APTINVH2").NewRow
            rowG.Item("VOUCHER_NO") = HFs("VOUCHER_NO")
            rowG.Item("VOUCHER_LNO") = 0
            rowG.Item("ACCT_CODE") = "2051"
            rowG.Item("SEG2_CODE") = "000"
            rowG.Item("SEG3_CODE") = "00"
            rowG.Item("SEG4_CODE") = "000"
            rowG.Item("INV_LINE_AMT") = -1 * TOTAL
            dst.Tables("APTINVH2").Rows.Add(rowG)

            Calc_DIST_GL()

            Absx1.txtFor("BANK_CODE").Text = "BLOP"
            Absx1.optFor("INV_STATUS").Value = "P"

            Absx1.dteFor("CHECK_DATE").Value = CDate("12/31/2015")
            Absx1.numFor("CHECK_AMT").Value = 0

            Dim rowGLTBANK1 As DataRow = Fill_Record("GLTBANK1", Absx1.txtFor("BANK_CODE").Text)
            BANK_LAST_CHECK_NO = rowGLTBANK1("BANK_LAST_CHECK_NO") & ""
            BANK_NEXT_CHECK_NO = CStr(Val(BANK_LAST_CHECK_NO) + 1)
            BANK_NEXT_CHECK_NO = ASCMAIN1.Format_Field(BANK_NEXT_CHECK_NO, "CHECK_NUM")
            Absx1.txtFor("CHECK_NUM").Text = BANK_NEXT_CHECK_NO
            auto_next_check = True

            Click_Command("Update")

            If ScreenMode Then
                Exit For
            End If
            'Click_Command("Cancel")
        Next

    End Sub

    Private Sub grdAPTINVH2_KeyDown(sender As Object, e As KeyEventArgs) Handles grdAPTINVH2.KeyDown
        If e.KeyCode = Keys.Delete Then
            If grdAPTINVH2.ActiveCell IsNot Nothing AndAlso grdAPTINVH2.ActiveCell.Column.Key = "COST_CTR_CODE" Then
                grdAPTINVH2.ActiveCell.Value = DBNull.Value
            End If
        End If
    End Sub

    Private Sub UltraTabControl1_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged
        With UltraExplorerBar1
            ' .Groups("Entry Options").Visible = False

            Select Case UltraTabControl1.ActiveTab.Key
                Case "Submitted Invoices"
                    .Groups("Submitted Invoices").Visible = True
                    .Groups("Screen Control").Visible = True
                    .Groups("Entry Options").Visible = False
                    Try
                        TAC.TACMAIN1.GetEmails(Me)
                    Catch ex As Exception
                        MsgBox("Could not Get New emails", MsgBoxStyle.OkOnly, "Please contact ABS")
                    End Try

                    OPT_SHOW_REFRESH()
                   ' Fill_Records("APTSUBM1")
                Case "A/P Invoices"
                    optShow.Value = "U"
                    OPT_SHOW_REFRESH()
                    .Groups("Submitted Invoices").Visible = False
                    '   .Groups("Screen Control").Visible = True
                    .Groups("Entry Options").Visible = True
            End Select
        End With
    End Sub

    Private Sub optShow_ValueChanged(sender As Object, e As EventArgs) Handles optShow.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        OPT_SHOW_REFRESH()
    End Sub
    Sub Update_APTSUBM1()
        Dim rowAPTSUBM1 As DataRow
        Dim SUBMIT_CTL_NO As String = INVOICE_FROM_EMAIL
        rowAPTSUBM1 = dst.Tables("APTSUBM1").Rows.Find(SUBMIT_CTL_NO)
        If rowAPTSUBM1 Is Nothing Then
        Else
            rowAPTSUBM1.Item("SUBMIT_CTL_NO") = SUBMIT_CTL_NO
            rowAPTSUBM1.Item("VOUCHER_NO") = HFs("VOUCHER_NO")
            rowAPTSUBM1.Item("INV_NUM") = rowAPTINVH1("INV_NUM")
            rowAPTSUBM1.Item("INV_DATE") = rowAPTINVH1("INV_DATE")
            rowAPTSUBM1.Item("INV_AMT") = rowAPTINVH1("INV_AMT")
            rowAPTSUBM1.Item("LAST_DATE") = DATETIME_STAMP
            rowAPTSUBM1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowAPTSUBM1.Item("SUBMIT_STATUS") = "P"
        End If

        Dim SUBMIT_SUBJECT As String = rowAPTSUBM1.Item("SUBMIT_SUBJECT") & ""

        Dim PEND_INVOICE_FILENAME = ASCMAIN1.Folders("Archive") & "SubmittedInvoices\" & SUBMIT_CTL_NO & ".eml"
        If rowAPTSUBM1.Item("SUBMIT_NO_ORIG") & "" <> "" Then
            'PEND_INVOICE_FILENAME = ASCMAIN1.Folders("Archive") & "SubmittedInvoices\" & rowAPTSUBM1.Item("SUBMIT_NO_ORIG") & ".msg"
            PEND_INVOICE_FILENAME = ASCMAIN1.Folders("Archive") & "SubmittedInvoices\" & rowAPTSUBM1.Item("SUBMIT_NO_ORIG") & ".eml"
        End If

        If PEND_INVOICE_FILENAME <> "" Then

            If Not dst.Tables.Contains("ASTATTA2") Then
                Create_TDA(dst.Tables.Add, "ASTATTA2", "*")
            End If
            Dim ATTACHMENT_NO As String = ASCMAIN1.Next_Control_No("ASTATTA2.ATTACHMENT_NO")

            'If My.Computer.FileSystem.FileExists(PEND_INVOICE_FILENAME) Then
            'Dim ATTACH_PATH As String = ASCMAIN1.Folders("Attach") & ATTACHMENT_NO
            'My.Computer.FileSystem.CopyFile(PEND_INVOICE_FILENAME, ATTACH_PATH)


            Dim rowASTATTA2 As DataRow = dst.Tables("ASTATTA2").NewRow
            With rowASTATTA2
                .Item("TABLE_NAME") = "APTINVH1"
                .Item("COLUMN_NAME") = "VOUCHER_NO"
                .Item("CODE_VALUE") = HFs("VOUCHER_NO")
                .Item("ATTACHMENT_NO") = ATTACHMENT_NO
                .Item("ATTACHMENT_DESC") = SUBMIT_SUBJECT
                .Item("ATTACHMENT_FILENAME") = PEND_INVOICE_FILENAME
                .Item("ATTACHMENT_EXT") = "eml"
                .Item("COMPUTER_NAME") = My.Computer.Name
                .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("ATTACHMENT_TYPE") = "EML"
                .Item("ATTACHMENT_ORIGINATOR") = ""
                .Item("ATTACHMENT_DATETIME") = DATETIME_STAMP
                .Item("ATTACHMENT_STATUS") = ""
                .Item("ATTACHMENT_NOTES") = ""
            End With
            dst.Tables("ASTATTA2").Rows.Add(rowASTATTA2)
            Update_Record_TDA("ASTATTA2")
            dst.Tables("ASTATTA2").Rows.Clear()
            'End If

            Try
                Dim ATTA_INVOICE_FILENAME As String = ASCMAIN1.Folders("Attach") & ATTACHMENT_NO
                My.Computer.FileSystem.CopyFile(PEND_INVOICE_FILENAME, ATTA_INVOICE_FILENAME, True)

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error trying to copy email to Attachment")
            End Try

        End If

    End Sub
    Sub PROCESS_DATES_REFRESH()

        If optShow.Value = "P" Then
            ASCMAIN1.sql = sqlAPTSUBM1 & " AND APTSUBM1.SUBMIT_STATUS IN ('P','D')" & vbCrLf _
            & " and APTSUBM1.INIT_DATE between '" & CDate(dteStart.Value).ToString("dd-MMM-yyyy") & "' and '" & CDate(dteEnd.Value).ToString("dd-MMM-yyyy") & "'"
            grdAPTSUBM1.Text = "Processed Submitted Invoices between " & dteStart.Value & " and " & dteStart.Value
        Else
            ASCMAIN1.sql = sqlAPTSUBM1 & " AND SUBMIT_STATUS = 'U'"
        End If
        Fill_Records("APTSUBM1", "", True, ASCMAIN1.sql)

    End Sub

    Private Sub cmdRefresh_Click(sender As Object, e As EventArgs) Handles cmdRefresh.Click
        PROCESS_DATES_REFRESH()
    End Sub


    Private Sub grdAPTSUBM1_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdAPTSUBM1.ClickCellButton
        If grdAPTSUBM1.ActiveCell.Column.Key = "EMAIL" Then
            Dim PEND_INVOICE_FILENAME = grdAPTSUBM1.ActiveRow.Cells("SUBMIT_CTL_NO").Text & ".eml"
            If grdAPTSUBM1.ActiveRow.Cells("SUBMIT_NO_ORIG").Text <> "" Then
                PEND_INVOICE_FILENAME = grdAPTSUBM1.ActiveRow.Cells("SUBMIT_NO_ORIG").Text & ".eml"
            End If
            Show_Document(ASCMAIN1.Folders("Archive") & "SubmittedInvoices\" & PEND_INVOICE_FILENAME)
        End If
    End Sub

    Private Sub grdAPTSUBM1_BeforeRowsDeleted(sender As Object, e As BeforeRowsDeletedEventArgs) Handles grdAPTSUBM1.BeforeRowsDeleted

        'Dim SUBMIT_CTL_NOs As New List(Of String)
        'For Each grow As UltraWinGrid.UltraGridRow In e.Rows
        '    SUBMIT_CTL_NOs.Add(grow.Cells("SUBMIT_CTL_NO").Value)
        'Next
        'grdAPTSUBM1.Tag = SUBMIT_CTL_NOs

    End Sub

    Sub OPT_SHOW_REFRESH()
        grdAPTSUBM1.Selected.Rows.Clear()
        If optShow.Value = "P" Then
            ' grdAPTSUBM1.DisplayLayout.Bands(0).Columns("SEL").Hidden = True
            grpDateRange.Visible = True
            With UltraExplorerBar1
                .Groups("Screen Control").Enabled = False
            End With
            PROCESS_DATES_REFRESH()
            grdAPTSUBM1.Text = "Processed"
        ElseIf optShow.Value = "U" Then
            ' grdAPTSUBM1.DisplayLayout.Bands(0).Columns("SEL").Hidden = False
            grpDateRange.Visible = False
            With UltraExplorerBar1
                .Groups("Screen Control").Enabled = True
            End With
            PROCESS_DATES_REFRESH()
            grdAPTSUBM1.Text = "Pending"
        End If

        With grdAPTSUBM1.DisplayLayout.Bands(0)
            .Columns("VOUCHER_NO").Hidden = (optShow.Value = "U")
            .Columns("INV_NUM").Hidden = (optShow.Value = "U")
            .Columns("INV_DATE").Hidden = (optShow.Value = "U")
            .Columns("INV_AMT").Hidden = (optShow.Value = "U")
            .Columns("INV_STATUS").Hidden = (optShow.Value = "U")
        End With
    End Sub

    Private Sub btnProRateTOTAL_INV_Click(sender As Object, e As EventArgs) Handles btnProRateTOTAL_INV.Click

        Dim TOTAL_INV As Decimal = System.Math.Round(Val(numTOTAL_INV.Value & ""), 2)
        If TOTAL_INV <= 0 Then
            MsgBox("Cannot ProRate $0 Total Invoice", MsgBoxStyle.OkOnly, "Cannot ProRate $0 Total Invoice")
            Exit Sub
        End If

        Dim CTL_NOs As New List(Of String)

        Dim TOTAL_ACC As Decimal = 0
        Dim TOTAL_ACC_max As Decimal = 0
        Dim CTL_NO_max As String = ""
        For Each grow As UltraWinGrid.UltraGridRow In grdAPTINVH7.Selected.Rows
            Dim COST_ACC As Decimal = Val(grow.Cells("TOTAL_ACC").Value)
            If COST_ACC > 0 Then
                Dim CTL_NO As String = grow.Cells("CTL_NO").Value
                CTL_NOs.Add(CTL_NO)
                TOTAL_ACC += COST_ACC
                If COST_ACC > TOTAL_ACC_max Then
                    TOTAL_ACC_max = COST_ACC
                    CTL_NO_max = CTL_NO
                End If
            End If
        Next

        If grdAPTINVH7.Selected.Rows.Count < 2 Or CTL_NOs.Count < 2 Then
            MsgBox("Cannot Pro-Rate unless you select 2 or more Rows with Positive Accrual Amounts", MsgBoxStyle.OkOnly, "Select 2 or more Rows to Prorate the Total Invoice Amount")
            Exit Sub
        End If

        Dim COST_ACT_total As Decimal = 0
        For Each CTL_NO As String In CTL_NOs
            Dim rowAPTINVH7 As DataRow = dst.Tables("APTINVH7").Select($"CTL_NO = '{CTL_NO}'")(0)
            Dim COST_ACC As Decimal = Val(rowAPTINVH7.Item("TOTAL_ACC") & "")
            Dim COST_ACT As Decimal = System.Math.Round(COST_ACC * TOTAL_INV / TOTAL_ACC, 2)
            COST_ACT_total += COST_ACT
            rowAPTINVH7.Item("TOTAL_INV") = COST_ACT

            Dim rowAPTACRC1 As DataRow = dst.Tables("APTACRC1").Rows.Find(CTL_NO)
            rowAPTACRC1.Item("COST_ACT") = COST_ACT
            Calc_DIST_Other()

        Next

        If COST_ACT_total <> TOTAL_INV Then
            Dim rowAPTINVH7 As DataRow = dst.Tables("APTINVH7").Select($"CTL_NO = '{CTL_NO_max}'")(0)
            Dim COST_ACT As Decimal = Val(rowAPTINVH7.Item("TOTAL_ACC") & "")
            COST_ACT += (TOTAL_INV - COST_ACT_total)
            rowAPTINVH7.Item("TOTAL_INV") = COST_ACT

            Dim rowAPTACRC1 As DataRow = dst.Tables("APTACRC1").Rows.Find(CTL_NO_max)
            rowAPTACRC1.Item("COST_ACT") = COST_ACT
            Calc_DIST_Other()
        End If

        grdAPTINVH7.Selected.Rows.Clear()

    End Sub

    Private Sub grdAPTINVH7_ExternalSummaryValueRequested(sender As Object, e As ExternalSummaryValueEventArgs) Handles grdAPTINVH7.ExternalSummaryValueRequested

    End Sub

    Private Sub grdAPTINVH7_BeforeRowRegionSize(sender As Object, e As BeforeRowRegionSizeEventArgs) Handles grdAPTINVH7.BeforeRowRegionSize

    End Sub

    Private Sub cmdFindReceipt_Click(sender As Object, e As EventArgs) Handles cmdFindReceipt.Click
        Dim YP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -3)
        ASCMAIN1.sql = "Select ICTIREC1.VEND_CODE, ICTIREC2.PO_ORDER_NO, ICTIREC2.PO_ORDER_LNO, ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO, ICTIREC2.RECEIPT_LNO, ICTIREC1.RECEIPT_DATE, ICTIREC2.ITEM_CODE" & vbCrLf _
            & " from ICTIREC1, ICTIREC2" & vbCrLf _
            & $" where ICTIREC1.VEND_CODE = '{HFs("VEND_CODE")}' and ICTIREC2.RECEIPT_NO = ICTIREC1.RECEIPT_NO And ICTIREC1.OPS_YYYYPP >= '{YP}'"
        Dim TBL As DataTable = ASCDATA1.GetDataTable
        Using frm As New ASFMSGBF
            frm.Show_grd(TBL, Me, "Select the Receipt for which to create an Accrual", "ROW")
            If frm.user_option = -1 Then
                ' USER CLICKED CANCEL
            Else
                Dim G As UltraWinGrid.UltraGridRow = frm.grow
                txtOA_ACCRUAL_CODE.Text = "FRT"
                dteOA_CTL_DATE.Value = G.Cells("RECEIPT_DATE").Value & ""
                txtOA_VEND_CODE_ACC.Text = G.Cells("VEND_CODE").Value & ""
                txtOA_PO_ORDER_NO.Text = G.Cells("PO_ORDER_NO").Value & ""
                txtOA_PO_ORDER_LNO.Text = G.Cells("PO_ORDER_LNO").Value & ""
                txtOA_RECEIPT_NO.Text = G.Cells("RECEIPT_NO").Value & ""
                txtOA_RECEIPT_LNO.Text = G.Cells("RECEIPT_LNO").Value & ""
                txtOA_ITEM_CODE.Text = G.Cells("ITEM_CODE").Value & ""
            End If
        End Using
    End Sub

    Private Sub cmdCreateAccrual_MouseLeave(sender As Object, e As EventArgs) Handles cmdCreateAccrual.MouseLeave

    End Sub

    Private Sub cmdCreateAccrual_FontChanged(sender As Object, e As EventArgs) Handles cmdCreateAccrual.FontChanged

    End Sub

    Sub Setup_APTCHCK5()

        If Me.SELECTION_NO = 0 Then Exit Sub

        dst.Tables("APTCHCK5").Rows.Clear()

        'If Not ScreenMode Then Exit Sub

        Dim VEND_CODE As String = rowAPTINVH1.Item("VEND_CODE")
        Dim VEND_ACCT_NO As String = ""
        Dim VEND_CODE_AP As String = rowAPTINVH1.Item("VEND_ALT_CODE") & ""
        Dim VEND_ALT_CODE As String = rowAPTINVH1.Item("VEND_ALT_CODE") & ""

        Dim VEND_EMAIL As String

        Dim VEND_CONTACT As String
        Dim VEND_PHONE As String


        Dim BANK_CODE As String = rowAPTINVH1.Item("BANK_CODE") & ""
        Dim CHECK_NUM As String = rowAPTINVH1.Item("CHECK_NUM") & ""
        Dim INV_STATUS As String = rowAPTINVH1.Item("INV_STATUS") & ""
        Dim INV_PYMT_METHOD As String = rowAPTINVH1.Item("INV_PYMT_METHOD") & ""

        Dim TAG_DATA As String = VEND_CODE & ":" & VEND_CODE_AP & ":" & VEND_ALT_CODE & ":" & BANK_CODE & ":" & INV_PYMT_METHOD
        grpBankingInfo.Tag = TAG_DATA

        If INV_PYMT_METHOD = "ECHECK" Or INV_PYMT_METHOD = "ACH" Or INV_PYMT_METHOD = "WIRE" Then
        Else
            Exit Sub
        End If

        Dim rowAPTCHCK5 As DataRow = Fill_Record("APTCHCK5", New String() {BANK_CODE, CHECK_NUM})
        If INV_STATUS = "P" Then
            If rowAPTCHCK5 IsNot Nothing Then
                Dim VEND_BANK_ACCT_ID As String = rowAPTCHCK5.Item("VEND_BANK_ACCT_ID")
                Dim VEND_BANK_ACCT_ID_DECRYPTED As String = ASCMAIN1.DecryptAES(VEND_BANK_ACCT_ID).ToString
                rowAPTCHCK5.Item("VEND_BANK_ACCT_ID") = VEND_BANK_ACCT_ID_DECRYPTED
            End If
            Exit Sub
        End If

        If rowAPTCHCK5 Is Nothing Then

            Dim VEND_BANK_ACCT_ID As String
            Dim VEND_BANK_ROUTING_NO As String
            Dim VEND_BANK_SWIFT_NO As String
            Dim VEND_BANK_ACCT_CLASS As String
            Dim VEND_BANK_ACCT_TYPE As String

            Dim VEND_BANK_NAME As String
            Dim VEND_BANK_ADDR1 As String
            Dim VEND_BANK_ADDR2 As String
            Dim VEND_BANK_ADDR3 As String

            Dim VEND_BANK_CITY As String
            Dim VEND_BANK_STATE As String
            Dim VEND_BANK_ZIP_CODE As String

            Dim VEND_BANK_COUNTRY As String
            Dim VEND_BANK_CONTACT As String
            Dim VEND_REMIT_EMAIL As String

            rowAPTCHCK5 = dst.Tables("APTCHCK5").NewRow
            rowAPTCHCK5.Item("BANK_CODE") = BANK_CODE
            rowAPTCHCK5.Item("CHECK_NUM") = CHECK_NUM
            dst.Tables("APTCHCK5").Rows.Add(rowAPTCHCK5)

            VEND_CODE_AP = rowAPTINVH1.Item("VEND_CODE_AP") & ""
            VEND_ALT_CODE = rowAPTINVH1.Item("VEND_ALT_CODE") & ""

            If VEND_CODE <> VEND_CODE_AP And VEND_CODE_AP <> "" Then
                LookUp("APTVEND1", VEND_CODE_AP)
                VEND_EMAIL = cdr.Item("VEND_EMAIL") & ""
                VEND_ACCT_NO = cdr.Item("VEND_ACCT_NO") & ""
                VEND_CONTACT = cdr.Item("VEND_CONTACT") & ""
                VEND_PHONE = cdr.Item("VEND_PHONE") & ""

                VEND_BANK_ACCT_ID = cdr.Item("VEND_BANK_ACCT_ID") & ""
                VEND_BANK_ROUTING_NO = cdr.Item("VEND_BANK_ROUTING_NO") & ""
                VEND_BANK_SWIFT_NO = cdr.Item("VEND_BANK_SWIFT_NO") & ""
                VEND_BANK_ACCT_CLASS = cdr.Item("VEND_BANK_ACCT_CLASS") & ""
                VEND_BANK_ACCT_TYPE = cdr.Item("VEND_BANK_ACCT_TYPE") & ""

                VEND_BANK_NAME = cdr.Item("VEND_BANK_NAME") & ""
                VEND_BANK_ADDR1 = cdr.Item("VEND_BANK_ADDR1") & ""
                VEND_BANK_ADDR2 = cdr.Item("VEND_BANK_ADDR2") & ""
                VEND_BANK_ADDR3 = cdr.Item("VEND_BANK_ADDR3") & ""

                VEND_BANK_CITY = cdr.Item("VEND_BANK_CITY") & ""
                VEND_BANK_STATE = cdr.Item("VEND_BANK_STATE") & ""
                VEND_BANK_ZIP_CODE = cdr.Item("VEND_BANK_ZIP_CODE") & ""
                VEND_BANK_COUNTRY = cdr.Item("VEND_BANK_COUNTRY") & ""
                VEND_BANK_CONTACT = cdr.Item("VEND_BANK_CONTACT") & ""
                VEND_REMIT_EMAIL = cdr.Item("VEND_REMIT_EMAIL") & ""

            Else
                If VEND_ALT_CODE <> "VENDOR" And VEND_ALT_CODE <> "" Then
                    LookUp("APTVEND2", New String() {IIf(VEND_CODE_AP = "", VEND_CODE, VEND_CODE_AP), VEND_ALT_CODE})
                    VEND_EMAIL = cdr.Item("VEND_ALT_EMAIL") & ""
                    VEND_CONTACT = cdr.Item("VEND_ALT_CONTACT") & ""
                    VEND_PHONE = cdr.Item("VEND_ALT_PHONE") & ""

                    VEND_BANK_ACCT_ID = cdr.Item("VEND_ALT_BANK_ACCT_ID") & ""
                    VEND_BANK_ROUTING_NO = cdr.Item("VEND_ALT_BANK_ROUTING_NO") & ""
                    VEND_BANK_SWIFT_NO = cdr.Item("VEND_ALT_BANK_SWIFT_NO") & ""
                    VEND_BANK_ACCT_CLASS = cdr.Item("VEND_ALT_BANK_ACCT_CLASS") & ""
                    VEND_BANK_ACCT_TYPE = cdr.Item("VEND_ALT_BANK_ACCT_TYPE") & ""

                    VEND_BANK_NAME = cdr.Item("VEND_ALT_BANK_NAME") & ""
                    VEND_BANK_ADDR1 = cdr.Item("VEND_ALT_BANK_ADDR1") & ""
                    VEND_BANK_ADDR2 = cdr.Item("VEND_ALT_BANK_ADDR2") & ""
                    VEND_BANK_ADDR3 = cdr.Item("VEND_ALT_BANK_ADDR3") & ""

                    VEND_BANK_CITY = cdr.Item("VEND_ALT_BANK_CITY") & ""
                    VEND_BANK_STATE = cdr.Item("VEND_ALT_BANK_STATE") & ""
                    VEND_BANK_ZIP_CODE = cdr.Item("VEND_ALT_BANK_ZIP_CODE") & ""
                    VEND_BANK_COUNTRY = cdr.Item("VEND_ALT_BANK_COUNTRY") & ""
                    VEND_BANK_CONTACT = cdr.Item("VEND_ALT_BANK_CONTACT") & ""
                    VEND_REMIT_EMAIL = cdr.Item("VEND_ALT_REMIT_EMAIL") & ""

                Else
                    LookUp("APTVEND1", IIf(VEND_CODE_AP = "", VEND_CODE, VEND_CODE_AP))
                    VEND_EMAIL = cdr.Item("VEND_EMAIL") & ""
                    VEND_ACCT_NO = cdr.Item("VEND_ACCT_NO") & ""
                    VEND_CONTACT = cdr.Item("VEND_CONTACT") & ""
                    VEND_PHONE = cdr.Item("VEND_PHONE") & ""

                    VEND_BANK_ACCT_ID = cdr.Item("VEND_BANK_ACCT_ID") & ""
                    VEND_BANK_ROUTING_NO = cdr.Item("VEND_BANK_ROUTING_NO") & ""
                    VEND_BANK_SWIFT_NO = cdr.Item("VEND_BANK_SWIFT_NO") & ""
                    VEND_BANK_ACCT_CLASS = cdr.Item("VEND_BANK_ACCT_CLASS") & ""
                    VEND_BANK_ACCT_TYPE = cdr.Item("VEND_BANK_ACCT_TYPE") & ""

                    VEND_BANK_NAME = cdr.Item("VEND_BANK_NAME") & ""
                    VEND_BANK_ADDR1 = cdr.Item("VEND_BANK_ADDR1") & ""
                    VEND_BANK_ADDR2 = cdr.Item("VEND_BANK_ADDR2") & ""
                    VEND_BANK_ADDR3 = cdr.Item("VEND_BANK_ADDR3") & ""

                    VEND_BANK_CITY = cdr.Item("VEND_BANK_CITY") & ""
                    VEND_BANK_STATE = cdr.Item("VEND_BANK_STATE") & ""
                    VEND_BANK_ZIP_CODE = cdr.Item("VEND_BANK_ZIP_CODE") & ""
                    VEND_BANK_COUNTRY = cdr.Item("VEND_BANK_COUNTRY") & ""
                    VEND_BANK_CONTACT = cdr.Item("VEND_BANK_CONTACT") & ""
                    VEND_REMIT_EMAIL = cdr.Item("VEND_REMIT_EMAIL") & ""
                End If
            End If

            With rowAPTCHCK5
                .Item("VEND_BANK_ACCT_ID") = VEND_BANK_ACCT_ID
                .Item("VEND_BANK_ROUTING_NO") = VEND_BANK_ROUTING_NO
                .Item("VEND_BANK_SWIFT_NO") = VEND_BANK_SWIFT_NO
                .Item("VEND_BANK_ACCT_CLASS") = VEND_BANK_ACCT_CLASS
                .Item("VEND_BANK_ACCT_TYPE") = VEND_BANK_ACCT_TYPE

                .Item("VEND_BANK_NAME") = VEND_BANK_NAME
                .Item("VEND_BANK_ADDR1") = VEND_BANK_ADDR1
                .Item("VEND_BANK_ADDR2") = VEND_BANK_ADDR2
                .Item("VEND_BANK_ADDR3") = VEND_BANK_ADDR3

                .Item("VEND_BANK_CITY") = VEND_BANK_CITY
                .Item("VEND_BANK_STATE") = VEND_BANK_STATE
                .Item("VEND_BANK_ZIP_CODE") = VEND_BANK_ZIP_CODE
                .Item("VEND_BANK_COUNTRY") = VEND_BANK_COUNTRY
                .Item("VEND_BANK_CONTACT") = VEND_BANK_CONTACT
                .Item("VEND_REMIT_EMAIL") = VEND_REMIT_EMAIL
                If VEND_BANK_ACCT_ID <> "" Then
                    .Item("VEND_BANK_ACCT_ID_DECRYPTED") = ASCMAIN1.DecryptAES(VEND_BANK_ACCT_ID).ToString
                End If
                ' for inquiry purposes
                .Item("VEND_BANK_ACCT_ID") = .Item("VEND_BANK_ACCT_ID_DECRYPTED")
            End With

            If INV_PYMT_METHOD = "ECHECK" Or INV_PYMT_METHOD = "ACH" Or INV_PYMT_METHOD = "WIRE" Then
                VEND_EMAIL = VEND_REMIT_EMAIL
            End If
        End If
    End Sub

    Private Sub cmdShowBankingInfo_Click(sender As Object, e As EventArgs) Handles cmdShowBankingInfo.Click
        Setup_APTCHCK5()

    End Sub

    Private Sub btnAddAccruals_Click(sender As Object, e As EventArgs) Handles btnAddAccruals.Click
        Add_Accrued_Other()
    End Sub

    Sub Add_Accrued_Other()

        If EntryMode <> "N" And EntryMode <> "E" Then
            Exit Sub
        End If

        If InquiryMode Then
            Exit Sub
        End If

        Dim VEND_CODE_ADD = Absx1.txtFor("VEND_CODE_ADD").Text.Trim

        If VEND_CODE_ADD = "" Then
            Exit Sub
        End If

        If LookUp("APTVEND1", VEND_CODE_ADD) Is Nothing Then
            MsgBox("Invalid Supplier Code (" & VEND_CODE_ADD & ")", MsgBoxStyle.OkOnly, "Cannot Add Accrued Other")
            Exit Sub
        End If

        If dst.Tables("APTACRC1").Select("VEND_CODE_ACC = '" & VEND_CODE_ADD & "'").Length > 0 Then
            MsgBox("Accrued Other for Supplier Code (" & VEND_CODE_ADD & ") have already been Added", MsgBoxStyle.OkOnly, "Cannot Add Accrued Other more than Once")
            Exit Sub
        End If


        If Not ASCMAIN1.Logical_Lock("APTVEND1", VEND_CODE_ADD,,, False) Then
            Exit Sub
        End If

        ASCMAIN1.Progress("Now Loading Accrued Other Accruals")
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()

        Dim sql_where As String = " and APTACRC1.CTL_NO in " _
        & " (Select CTL_NO from APTACRC1 where APTACRC1.CTL_STATUS = '0' AND VEND_CODE_ACC = '" & VEND_CODE_ADD & "')"

        ASCMAIN1.sql = "Select APTACRC1.*" & vbCrLf _
                & " from APTACRC1" & vbCrLf _
                & " where APTACRC1.CTL_STATUS = '0'" & vbCrLf _
                & sql_where

        Fill_Records("APTACRC1", , False, ASCMAIN1.sql)

        Sort_grdColumns(grdAPTACRC1, "CTL_NO".ToLower)

        With grdAPTACRC1.DisplayLayout.Bands(0).Columns("VEND_CODE_ACC")
            .Hidden = False
            .Header.Caption = "Supplier"
        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub btnAutoMatch_Click(sender As Object, e As EventArgs) Handles btnAutoMatch.Click

        Dim MATCH_CANDIDATES As New List(Of String)
        For Each rowSOURCE_DOC_NO As DataRow In ASCDATA1.SelectDistinct(dst.Tables("APTACRC1"), "SOURCE_DOC_NO").Select()
            Dim SOURCE_DOC_NO As String = rowSOURCE_DOC_NO.Item("SOURCE_DOC_NO")
            Dim PPDS As Integer = 0
            Dim RECS As Integer = 0
            For Each rowAPTACRC1 As DataRow In dst.Tables("APTACRC1").Select($"SOURCE_DOC_NO = '{SOURCE_DOC_NO}'")
                Dim PPD_IND As String = rowAPTACRC1.Item("PPD_IND") & ""
                If PPD_IND = "1" Then PPDS += 1
                RECS += 1
            Next

            If RECS > 1 And PPDS = 1 Then
                MATCH_CANDIDATES.Add(SOURCE_DOC_NO)
            End If
        Next

        For Each SOURCE_DOC_NO As String In MATCH_CANDIDATES
            For Each rowAPTACRC1 As DataRow In dst.Tables("APTACRC1").Select($"SOURCE_DOC_NO = '{SOURCE_DOC_NO}'")
                Dim CTL_NO As String = rowAPTACRC1.Item("CTL_NO")
                ProcessDoubleClickedRow_APTACRC1(CTL_NO)
            Next
        Next

    End Sub

    Private Sub grdAPTINVH7_BeforeRowActivate(sender As Object, e As RowEventArgs) Handles grdAPTINVH7.BeforeRowActivate

        'If e.Row.IsDataRow AndAlso e.Row.Cells("ACCRUAL_CODE").Value & "" = "TRF" Then
        '    Dim CTL_NO As String = e.Row.Cells("CTL_NO").Value
        '    Dim rowAPTACRC1 As DataRow = dst.Tables("APTACRC1").Rows.Find(CTL_NO)
        'End If

        If e.Row.IsDataRow AndAlso e.Row.Cells("ACCRUAL_CODE").Value & "" = "TRF" AndAlso e.Row.Cells("CTL_TYPE").Value & "" = "M" Then
            e.Row.Cells("TOTAL_INV").Column.CellActivation = Activation.NoEdit
        Else
            e.Row.Cells("TOTAL_INV").Column.CellActivation = Activation.AllowEdit
        End If
    End Sub

    Private Sub grdAPTINVH7_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdAPTINVH7.InitializeRow
        If e.Row.IsDataRow AndAlso e.Row.Cells("ACCRUAL_CODE").Value & "" = "TRF" AndAlso e.Row.Cells("CTL_TYPE").Value & "" = "M" Then
            e.Row.Cells("TOTAL_INV").Appearance.BackColor = System.Drawing.Color.Beige
        Else
            e.Row.Cells("TOTAL_INV").Appearance.BackColor = System.Drawing.Color.Empty
        End If
    End Sub

    Private Sub cmdMultiPORec_Click(sender As Object, e As EventArgs) Handles cmdMultiPORec.Click

        Dim EMsg As String = ""
        Dim rowAPTACRM1 As DataRow = Nothing

        'Dim ACCRUAL_CODE As String = Absx1.txtFor("ACCRUAL_CODE").Text
        'If ACCRUAL_CODE = "" Then
        '    EMsg &= vbCr & "No Value Specified for Accrual Code"
        'Else
        '    rowAPTACRM1 = LookUp("APTACRM1", ACCRUAL_CODE)
        '    If rowAPTACRM1 Is Nothing Then
        '        EMsg &= vbCr & "Invalid Value Specified for Accrual Code"
        '    End If
        'End If

        'Dim VEND_CODE_ACC As String = Absx1.txtFor("VEND_CODE_ACC").Text
        'If VEND_CODE_ACC = "" Then
        '    VEND_CODE_ACC = Absx1.txtFor("VEND_CODE").Text
        '    ' EMsg &= vbCr & "No Value Specified for Vendor Code"
        'Else
        '    rowAPTVEND1 = LookUp("APTVEND1", VEND_CODE_ACC)
        '    If rowAPTVEND1 Is Nothing Then
        '        EMsg &= vbCr & "Invalid Value Specified for Vendor Code"
        '    End If
        'End If

        'Dim CTL_NOTE As String = Trim(Absx1.txtFor("CTL_NOTE").Text)
        'If CTL_NOTE = "" Then
        '    EMsg &= vbCr & "No Value Specified for Note"
        'End If

        'Dim COST_ACT As Decimal = numCOST_ACT.Value
        'If COST_ACT = 0 Then
        '    EMsg &= vbCr & "Invalid Value Specified for Actual Cost"
        'End If

        Dim INV_DATE As Date = Absx1.dteFor("INV_DATE").Value
        If INV_DATE & "" = "" Or Format(INV_DATE, "yyyyMMdd") = "00010101" Then
            EMsg &= vbCr & "Invalid Value Specified for Invoice Date"
        End If

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Create Accrual")
            Exit Sub
        End If

        Using frmAPFINVHR As New APFINVHR(Me, Absx1.txtFor("VEND_CODE").Text, Absx1.txtFor("VOUCHER_NO").Text, Absx1.txtFor("INV_NUM").Text, Absx1.dteFor("INV_DATE").Value)
            frmAPFINVHR.ShowDialog()

            If frmAPFINVHR.updated Then
                If frmAPFINVHR.CTL_NOs.Count > 0 Then
                    For Each CTL_NO As String In frmAPFINVHR.CTL_NOs
                        Add_APTINVH7(CTL_NO, True)
                    Next
                    Clear_Other_Accrual_Controls()
                End If

            End If
        End Using

    End Sub

    'Private Sub btnCreatePPDDuty_Click(sender As Object, e As EventArgs)

    '    Dim INV_LINE_AMT As Decimal = Val(numPPDDuty.Value & "")
    '    If INV_LINE_AMT <= 0 Then
    '        MsgBox("PPD Amount must be a positive non-zero value", MsgBoxStyle.OkCancel, "Cannot Create PPD")
    '        Exit Sub
    '    End If

    '    If txtPPDDutyBOL.Text = "" Or txtPPDDutyBOL.Text.Length < 6 Then
    '        MsgBox("You must enter a value for the Bill of Lading", MsgBoxStyle.OkCancel, "Cannot Create PPD")
    '        Exit Sub
    '    End If

    '    Dim VOUCHER_LNO_ctr = Val(dst.Tables("APTINVH2").Compute("MAX(VOUCHER_LNO)", "") & "")
    '    Dim rowAPTINVH2 As DataRow = dst.Tables("APTINVH2").NewRow
    '    With rowAPTINVH2
    '        .Item("VOUCHER_NO") = HFs("VOUCHER_NO")
    '        VOUCHER_LNO_ctr += 1
    '        .Item("VOUCHER_LNO") = VOUCHER_LNO_ctr
    '        .Item("ACCT_CODE") = ROWs("ICTPARM1").Item("IC_PARM_ACCT_CODE_ACCR_TOOLG")
    '        .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
    '        .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
    '        .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
    '        .Item("ACCT_DESC") = LookUp("GLTACCT1", rowAPTINVH2.Item("ACCT_CODE") & "", True).ITEM("ACCT_DESC")
    '        .Item("INV_LINE_AMT") = INV_LINE_AMT
    '        .Item("INV_COMMENT_DTL") = txtPPDDutyBOL.Text
    '        .Item("INV_LTYP") = "T"
    '    End With
    '    dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)


    '    Dim VOUCHER_CLNO_ctr = Val(dst.Tables("APTINVH7").Compute("MAX(VOUCHER_CLNO)", "") & "")
    '    Dim rowAPTINVH7 As DataRow = dst.Tables("APTINVH7").NewRow
    '    With rowAPTINVH7
    '        .Item("VOUCHER_NO") = HFs("VOUCHER_NO")
    '        VOUCHER_LNO_ctr += 1
    '        .Item("VOUCHER_CLNO") = VOUCHER_LNO_ctr
    '        .Item("CTL_NO") = "0000000000"
    '        .Item("CHARGEBACK_IND") = "0"
    '        .Item("CTL_NO_CREATED") = "0"
    '        .Item("TOTAL_ACC") = 0
    '        .Item("TOTAL_INV") = INV_LINE_AMT
    '        .Item("CTL_REF_NO") = txtPPDDutyBOL.Text
    '    End With
    '    dst.Tables("APTINVH7").Rows.Add(rowAPTINVH7)

    '    Calc_DIST_GL()

    '    txtPPDDutyBOL.Text = ""
    '    numPPDDuty.Value = 0
    'End Sub
End Class