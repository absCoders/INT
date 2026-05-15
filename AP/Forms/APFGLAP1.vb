Public Class APFGLAP1
    Dim APTAPGLV As String
    Dim APTGLAP1_APIN As String
    Dim APTGLAP1_APCD As String
    Dim APTGLAP1_APGL As String
    Dim APTAEXPX As String
    Dim sql_APTAEXPX As String
    Dim RYP As String = ""

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst

            .Tables.Add("APTGLAP1")
            With .Tables("APTGLAP1").Columns
                .Add("LINE_TYPE", GetType(System.String))
                .Add("DESCRIPTION", GetType(System.String))
                .Add("JOURNAL_DATE", GetType(System.DateTime))
                .Add("APIN_AP", GetType(System.Double))
                .Add("APIN_GL", GetType(System.Double))
                .Add("APCD_AP", GetType(System.Double))
                .Add("APCD_GL", GetType(System.Double))
                .Add("OTHER_GL", GetType(System.Double))
                .Add("BEG_AP", GetType(System.Double))
                .Add("BEG_GL", GetType(System.Double))
                .Add("NET_AP", GetType(System.Double), "BEG_AP + APIN_AP - APCD_AP")
                .Add("NET_GL", GetType(System.Double), "BEG_GL + APIN_GL + APCD_GL + OTHER_GL")
            End With

            With .Tables("APTGLAP1")
                For I As Integer = 0 To .Columns.Count - 1
                    If .Columns(I).DataType.ToString = "System.Double" Then
                        .Columns(I).DefaultValue = 0
                    End If
                Next
            End With

            ASCMAIN1.sql = "" _
                & "SELECT GLTJRNL1.JOURNAL_DATE, APTINVH1.REGISTER_XNO" & vbCrLf _
                & ", APTINVH1.POST_CODE, APTINVH1.INV_TYPE, APTVEND1.VEND_TYPE" & vbCrLf _
                & ", SUM (APTINVH1.INV_AMT) INV_AMT" & vbCrLf _
                & " FROM APTINVH1,GLTJRNL1,APTVEND1 " & vbCrLf _
                & "WHERE GLTJRNL1.REGISTER_XNO (+) = APTINVH1.REGISTER_XNO" & vbCrLf _
                & "AND GLTJRNL1.JOURNAL_TYPE (+) = 'APIN'" & vbCrLf _
                & "AND APTINVH1.OPS_YYYYPP = :PARM1" & vbCrLf _
                & "AND NVL(APTINVH1.REGISTER_IND,'0') <> 'D'" & vbCrLf _
                & "AND NVL(APTINVH1.INV_STATUS,'?') <> 'R'" & vbCrLf _
                & "AND APTVEND1.VEND_CODE = APTINVH1.VEND_CODE" & vbCrLf _
                & "GROUP BY GLTJRNL1.JOURNAL_DATE, APTINVH1.REGISTER_XNO" & vbCrLf _
                & ", APTINVH1.POST_CODE, APTINVH1.INV_TYPE, APTVEND1.VEND_TYPE" & vbCr
            Create_TDA(.Tables.Add, "APTGLAP1_APIN", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "" _
            & "SELECT JOURNAL_DATE, REGISTER_XNO, POST_CODE, INV_TYPE, VEND_TYPE, BANK_CODE, CHECK_STATUS" & vbCrLf _
            & ", SUM (PYMT) PYMT" & vbCrLf _
            & ", SUM (APPL) APPL" & vbCrLf _
            & ", SUM (DISC) DISC FROM (" & vbCrLf _
            & "SELECT GLTJRNL1.JOURNAL_DATE, APTCHCK1.REGISTER_XNO" & vbCrLf _
            & ", APTINVH1.POST_CODE, APTINVH1.INV_TYPE, APTVEND1.VEND_TYPE, APTCHCK1.BANK_CODE, APTCHCK1.CHECK_STATUS" & vbCrLf _
            & ", SUM (APTCHCK2.INV_AMT_APPLIED - APTCHCK2.INV_DISC_TAKEN) PYMT" & vbCrLf _
            & ", SUM (APTCHCK2.INV_AMT_APPLIED) APPL" & vbCrLf _
            & ", SUM (APTCHCK2.INV_DISC_TAKEN) DISC" & vbCrLf _
            & " FROM APTCHCK1, APTCHCK2, APTINVH1, GLTJRNL1, APTVEND1 " & vbCrLf _
            & "WHERE GLTJRNL1.REGISTER_XNO (+) = APTCHCK1.REGISTER_XNO" & vbCrLf _
            & "AND GLTJRNL1.JOURNAL_TYPE (+) = 'APCD'" & vbCrLf _
            & "AND APTCHCK1.OPS_YYYYPP = :PARM1" & vbCrLf _
            & "AND APTCHCK2.BANK_CODE = APTCHCK1.BANK_CODE" & vbCrLf _
            & "AND APTCHCK2.CHECK_NUM = APTCHCK1.CHECK_NUM" & vbCrLf _
            & "AND APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO" & vbCrLf _
            & "AND APTVEND1.VEND_CODE = APTCHCK2.VEND_CODE" & vbCrLf _
            & "GROUP BY GLTJRNL1.JOURNAL_DATE, APTCHCK1.REGISTER_XNO" & vbCrLf _
            & ", APTINVH1.POST_CODE, APTINVH1.INV_TYPE, APTVEND1.VEND_TYPE, APTCHCK1.BANK_CODE, APTCHCK1.CHECK_STATUS" & vbCrLf _
            & "UNION" & vbCrLf _
            & "SELECT GLTJRNL1.JOURNAL_DATE, APTCHCK1.REGISTER_XNO" & vbCrLf _
            & ", APTINVH1.POST_CODE, APTINVH1.INV_TYPE, APTVEND1.VEND_TYPE, APTCHCK1.BANK_CODE, APTCHCK1.CHECK_STATUS" & vbCrLf _
            & ", SUM (-1 * (APTCHCK2.INV_AMT_APPLIED - APTCHCK2.INV_DISC_TAKEN)) PYMT" & vbCrLf _
            & ", SUM (-1 * APTCHCK2.INV_AMT_APPLIED) APPL" & vbCrLf _
            & ", SUM (-1 * APTCHCK2.INV_DISC_TAKEN) DISC" & vbCrLf _
            & "FROM APTCHCK1, APTCHCK2, APTINVH1, GLTJRNL1, APTVEND1 " & vbCrLf _
            & "WHERE GLTJRNL1.REGISTER_XNO (+) = APTCHCK1.REGISTER_XNO_F" & vbCrLf _
            & "AND GLTJRNL1.JOURNAL_TYPE (+) = 'APCD'" & vbCrLf _
            & "AND APTCHCK1.OPS_YYYYPP_F = :PARM1" & vbCrLf _
            & "AND APTCHCK2.BANK_CODE = APTCHCK1.BANK_CODE" & vbCrLf _
            & "AND APTCHCK2.CHECK_NUM = APTCHCK1.CHECK_NUM" & vbCrLf _
            & "AND APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO" & vbCrLf _
            & "AND APTVEND1.VEND_CODE = APTCHCK2.VEND_CODE" & vbCrLf _
            & "GROUP BY GLTJRNL1.JOURNAL_DATE, APTCHCK1.REGISTER_XNO" & vbCrLf _
            & ", APTINVH1.POST_CODE, APTINVH1.INV_TYPE, APTVEND1.VEND_TYPE, APTCHCK1.BANK_CODE, APTCHCK1.CHECK_STATUS" & vbCrLf _
            & ") GROUP BY JOURNAL_DATE, REGISTER_XNO, POST_CODE, INV_TYPE, VEND_TYPE, BANK_CODE, CHECK_STATUS" & vbCr
            Create_TDA(.Tables.Add, "APTGLAP1_APCD", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "" _
            & "SELECT GLTJRNL1.JOURNAL_DATE" & vbCrLf _
            & ", GLTJRNL1.REGISTER_XNO REGISTER_XNO" & vbCrLf _
            & ", GLTJRNL1.JOURNAL_TYPE" & vbCrLf _
            & ", (CASE WHEN GLTJRNL1.JOURNAL_TYPE LIKE 'GL%' THEN GLTJRNL1.JOURNAL_NO" & vbCrLf _
            & " ELSE '000000' END) JOURNAL_NO" & vbCrLf _
            & ", (CASE WHEN GLTJRNL1.JOURNAL_TYPE LIKE 'GL%' THEN GLTJRNL1.JOURNAL_DESC" & vbCrLf _
            & " ELSE GLTTYPE1.JOURNAL_TYPE_DESC END) JOURNAL_DESC" & vbCrLf _
            & ", SUM (DETL_POSTING_AMT) AMT" & vbCrLf _
            & "  FROM GLTDETL1, GLTJRNL1, GLTTYPE1" & vbCrLf _
            & " WHERE GLTDETL1.JOURNAL_NO = GLTJRNL1.JOURNAL_NO" & vbCrLf _
            & "AND GLTDETL1.OPS_YYYYPP = :PARM1" & vbCrLf _
            & "AND GLTDETL1.ACCT_CODE IN (SELECT DISTINCT ACCT_CODE FROM APTPOST1)" & vbCrLf _
            & "AND GLTTYPE1.JOURNAL_TYPE (+) = GLTJRNL1.JOURNAL_TYPE" & vbCrLf _
            & "GROUP BY GLTJRNL1.JOURNAL_DATE, GLTJRNL1.REGISTER_XNO, GLTJRNL1.JOURNAL_TYPE" & vbCrLf _
            & ", (CASE WHEN GLTJRNL1.JOURNAL_TYPE LIKE 'GL%' THEN GLTJRNL1.JOURNAL_NO" & vbCrLf _
            & "ELSE '000000' END)" & vbCrLf _
            & ", (CASE WHEN GLTJRNL1.JOURNAL_TYPE LIKE 'GL%' THEN GLTJRNL1.JOURNAL_DESC" & vbCrLf _
            & "ELSE GLTTYPE1.JOURNAL_TYPE_DESC END)" & vbCr
            Create_TDA(.Tables.Add, "APTGLAP1_APGL", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "Select APTVEND1.VEND_CODE, APTVEND1.VEND_NAME, APTVEND1.VEND_TYPE, APTVEND1.VEND_CLASS_CODE" _
            & " from APTVEND1"
            Create_TDA(.Tables.Add, "APTSUMV1", "**", 0, False, "", 1)
            With .Tables("APTSUMV1").Columns
                .Add("BEG_BAL", GetType(System.Decimal))
                .Add("APIN_ALL", GetType(System.Decimal))
                .Add("APCD_PYMT", GetType(System.Decimal))
                .Add("APCD_APPL", GetType(System.Decimal))
                .Add("APCD_DISC", GetType(System.Decimal))
                .Add("END_BAL", GetType(System.Decimal))
                .Add("BEG_BAL_YTD", GetType(System.Decimal))
                .Add("APIN_ALL_YTD", GetType(System.Decimal))
                .Add("APCD_PYMT_YTD", GetType(System.Decimal))
                .Add("APCD_APPL_YTD", GetType(System.Decimal))
                .Add("APCD_DISC_YTD", GetType(System.Decimal))
                .Add("END_BAL_YTD", GetType(System.Decimal))
                .Add("ACCRUED_IN", GetType(System.Decimal))
                .Add("ACCRUED_INTO", GetType(System.Decimal))
                .Add("OOBAL", GetType(System.Decimal), "ISNULL(BEG_BAL,0)+ISNULL(APIN_ALL,0)-ISNULL(APCD_PYMT,0)-ISNULL(END_BAL,0)")
            End With

            ASCMAIN1.sql = "Select APTINVH1.VEND_CODE, APTINVH1.VOUCHER_NO, APTINVH1.INV_TYPE" & vbCrLf _
            & ", APTINVH1.INV_NUM, APTINVH1.INV_DATE, APTINVH1.INV_AMT" & vbCrLf _
            & ", APTINVH1.INV_REF, APTINVH1.PO_ORDER_NO, APTINVH1.TERM_CODE" & vbCrLf _
            & ", APTINVH1.INV_DUE_DATE, APTINVH1.OPS_YYYYPP_ACCRUE, APTINVH1.OPS_YYYYPP" & vbCrLf _
            & ", APTINVH1.POST_CODE, APTINVH1.REGISTER_XNO, GLTJRNL1.JOURNAL_DATE" & vbCrLf _
            & " from APTINVH1,GLTJRNL1" & vbCrLf _
            & " where GLTJRNL1.JOURNAL_TYPE (+) = 'APIN'" & vbCrLf _
            & "   and GLTJRNL1.REGISTER_XNO (+) = APTINVH1.REGISTER_XNO" & vbCrLf _
            & "   and NVL(APTINVH1.REGISTER_IND,'0') <> 'D'" & vbCrLf _
            & "   and NVL(APTINVH1.INV_STATUS,'?') <> 'R'" & vbCrLf _
            & "   and APTINVH1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "APTSUMV2", "**", 0, False, "V", 2)

            .Relations.Add("APTSUMV2", _
            .Tables("APTSUMV1").Columns("VEND_CODE"), _
            .Tables("APTSUMV2").Columns("VEND_CODE"))

            ASCMAIN1.sql _
            = "Select APTCHCK1.VEND_CODE, APTCHCK1.BANK_CODE, APTCHCK1.CHECK_NUM, 'I' CHECK_STATUS, APTCHCK1.CHECK_AMT" _
            & " from APTCHCK1 where OPS_YYYYPP = :PARM1" _
            & " union " _
            & "Select APTCHCK1.VEND_CODE, APTCHCK1.BANK_CODE, APTCHCK1.CHECK_NUM, 'V' CHECK_STATUS, -1 * APTCHCK1.CHECK_AMT" _
            & " from APTCHCK1 where OPS_YYYYPP_F = :PARM1"
            Create_TDA(.Tables.Add, "APTSUMV3", "**", 0, False, "V", 4)
            With .Tables("APTSUMV3").Columns
                .Add("APPL", GetType(System.Double))
                .Add("DISC", GetType(System.Double))
            End With

            .Relations.Add("APTSUMV3", _
            .Tables("APTSUMV1").Columns("VEND_CODE"), _
            .Tables("APTSUMV3").Columns("VEND_CODE"))

            ASCMAIN1.sql _
            = "Select APTCHCK1.BANK_CODE, APTCHCK1.CHECK_NUM, 'I' CHECK_STATUS" & vbCrLf _
            & ", APTCHCK2.VOUCHER_NO, APTINVH1.INV_TYPE, APTINVH1.INV_NUM, APTINVH1.INV_REF" & vbCrLf _
            & ", APTCHCK2.INV_AMT_APPLIED, APTCHCK2.INV_DISC_TAKEN" & vbCrLf _
            & " from APTCHCK1,APTCHCK2,APTINVH1 " & vbCrLf _
            & " where APTCHCK1.BANK_CODE = APTCHCK2.BANK_CODE and APTCHCK1.CHECK_NUM = APTCHCK2.CHECK_NUM " & vbCrLf _
            & "  and APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO" & vbCrLf _
            & "  and APTCHCK1.OPS_YYYYPP = :PARM1" & vbCrLf _
            & " union " & vbCrLf _
            & "Select APTCHCK1.BANK_CODE, APTCHCK1.CHECK_NUM, 'V' CHECK_STATUS" & vbCrLf _
            & ", APTCHCK2.VOUCHER_NO, APTINVH1.INV_TYPE, APTINVH1.INV_NUM, APTINVH1.INV_REF" & vbCrLf _
            & ", -1 * APTCHCK2.INV_AMT_APPLIED, -1 * APTCHCK2.INV_DISC_TAKEN" & vbCrLf _
            & " from APTCHCK1,APTCHCK2,APTINVH1 " & vbCrLf _
            & " where APTCHCK1.BANK_CODE = APTCHCK2.BANK_CODE and APTCHCK1.CHECK_NUM = APTCHCK2.CHECK_NUM " & vbCrLf _
            & "  and APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO" & vbCrLf _
            & "  and APTCHCK1.OPS_YYYYPP_F = :PARM1"
            Create_TDA(.Tables.Add, "APTSUMV4", "**", 0, False, "V", 3)

            .Relations.Add("APTSUMV4", _
            New DataColumn() {.Tables("APTSUMV3").Columns("BANK_CODE"), .Tables("APTSUMV3").Columns("CHECK_NUM"), .Tables("APTSUMV3").Columns("CHECK_STATUS")}, _
            New DataColumn() {.Tables("APTSUMV4").Columns("BANK_CODE"), .Tables("APTSUMV4").Columns("CHECK_NUM"), .Tables("APTSUMV4").Columns("CHECK_STATUS")})

            .Tables("APTSUMV3").Columns("APPL").Expression = "SUM(Child(APTSUMV4).INV_AMT_APPLIED)"
            .Tables("APTSUMV3").Columns("DISC").Expression = "SUM(Child(APTSUMV4).INV_DISC_TAKEN)"


            ASCMAIN1.sql = "" _
            & "SELECT APTINVH1.VOUCHER_NO, APTINVH1.VEND_CODE, APTINVH1.INV_TYPE" & vbCrLf _
            & ", APTINVH1.INV_NUM, APTINVH1.INV_DATE, APTINVH1.INV_AMT" & vbCrLf _
            & ", APTINVH1.INV_REF, APTINVH1.PO_ORDER_NO, APTINVH1.TERM_CODE" & vbCrLf _
            & ", APTINVH1.INV_DUE_DATE, APTINVH1.POST_CODE, APTINVH1.INIT_OPER" & vbCrLf _
            & ", APTINVH1.INV_PAID_UPON_ENTRY, APTINVH1.REGISTER_XNO" & vbCrLf _
            & " from APTINVH1" & vbCrLf _
            & " where APTINVH1.OPS_YYYYPP = :PARM1" & vbCrLf _
            & "   and NVL(APTINVH1.REGISTER_IND,'0') <> 'D'" & vbCrLf _
            & "   and NVL(APTINVH1.INV_STATUS,'?') <> 'R'" & vbCr
            Create_TDA(.Tables.Add, "APTINVH1", "**", 0, False, "V", 1)
            With .Tables("APTINVH1").Columns
                .Add("VEND_NAME", GetType(System.String))
                .Add("VEND_TYPE", GetType(System.String))
                .Add("VEND_CLASS_CODE", GetType(System.String))
            End With


            ASCMAIN1.sql = "" _
            & "SELECT APTINVH2.VOUCHER_NO, APTINVH2.VOUCHER_LNO" & vbCrLf _
            & ", APTINVH2.ACCT_CODE, APTINVH2.SEG2_CODE, APTINVH2.SEG3_CODE, APTINVH2.SEG4_CODE" & vbCrLf _
            & ", APTINVH2.INV_LINE_AMT, APTINVH2.INV_COMMENT_DTL" & vbCrLf _
            & ", APTINVH2.INV_LTYP, APTINVH2.INV_DLNO" & vbCrLf _
            & " from APTINVH1,APTINVH2" & vbCrLf _
            & " where APTINVH1.VOUCHER_NO = APTINVH2.VOUCHER_NO" & vbCrLf _
            & "   and APTINVH1.OPS_YYYYPP = :PARM1" & vbCr
            Create_TDA(.Tables.Add, "APTINVH2", "**", 0, False, "V", 2)

            .Relations.Add("APTINVH2", _
            .Tables("APTINVH1").Columns("VOUCHER_NO"), _
            .Tables("APTINVH2").Columns("VOUCHER_NO"))

            ASCMAIN1.sql = "" _
            & "SELECT 'I' CHECK_STATUS, APTCHCK1.BANK_CODE, APTCHCK1.CHECK_NUM" & vbCrLf _
            & ", APTCHCK1.CHECK_DATE, APTCHCK1.CHECK_AMT, APTCHCK1.PYMT_METHOD" & vbCrLf _
            & ", APTCHCK1.VEND_CODE, APTCHCK1.VEND_NAME, APTCHCK1.BATCH_NO_PYMT" & vbCrLf _
            & ", APTCHCK1.VEND_CODE_AP, APTCHCK1.REGISTER_XNO" & vbCrLf _
            & ", APTVEND1.VEND_TYPE, APTVEND1.VEND_CLASS_CODE" & vbCrLf _
            & " FROM APTCHCK1,APTVEND1 WHERE APTVEND1.VEND_CODE = APTCHCK1.VEND_CODE" & vbCrLf _
            & " AND APTCHCK1.OPS_YYYYPP = :PARM1" & vbCrLf _
            & " UNION " & vbCrLf _
            & "SELECT 'V' CHECK_STATUS, APTCHCK1.BANK_CODE, APTCHCK1.CHECK_NUM" & vbCrLf _
            & ", APTCHCK1.CHECK_DATE, -1 * APTCHCK1.CHECK_AMT, APTCHCK1.PYMT_METHOD" & vbCrLf _
            & ", APTCHCK1.VEND_CODE, APTCHCK1.VEND_NAME, APTCHCK1.BATCH_NO_PYMT" & vbCrLf _
            & ", APTCHCK1.VEND_CODE_AP, APTCHCK1.REGISTER_XNO_F REGISTER_XNO" & vbCrLf _
            & ", APTVEND1.VEND_TYPE, APTVEND1.VEND_CLASS_CODE" & vbCrLf _
            & " FROM APTCHCK1,APTVEND1 WHERE APTVEND1.VEND_CODE = APTCHCK1.VEND_CODE" & vbCrLf _
            & " AND APTCHCK1.OPS_YYYYPP_F = :PARM1" & vbCr
            Create_TDA(.Tables.Add, "APTCHCK1", "**", 0, False, "V", 3)
            With .Tables("APTCHCK1").Columns
                .Add("APPL", GetType(System.Double))
                .Add("DISC", GetType(System.Double))
            End With

            ASCMAIN1.sql = "" _
            & "Select 'I' CHECK_STATUS, APTCHCK2.BANK_CODE, APTCHCK2.CHECK_NUM" & vbCrLf _
            & ", APTCHCK2.SEQ_NUM, APTCHCK2.VEND_CODE" & vbCrLf _
            & ", APTCHCK2.VOUCHER_NO, APTCHCK2.INV_NUM, APTCHCK2.INV_DATE" & vbCrLf _
            & ", APTCHCK2.INV_AMT_APPLIED, APTCHCK2.INV_DISC_TAKEN" & vbCrLf _
            & " from APTCHCK1,APTCHCK2 " & vbCrLf _
            & " where APTCHCK1.BANK_CODE = APTCHCK2.BANK_CODE" & vbCrLf _
            & "   and APTCHCK1.CHECK_NUM = APTCHCK2.CHECK_NUM" & vbCrLf _
            & "   and APTCHCK1.OPS_YYYYPP = :PARM1" & vbCrLf _
            & " union " & vbCrLf _
            & "Select 'V' CHECK_STATUS, APTCHCK2.BANK_CODE, APTCHCK2.CHECK_NUM" & vbCrLf _
            & ", APTCHCK2.SEQ_NUM, APTCHCK2.VEND_CODE" & vbCrLf _
            & ", APTCHCK2.VOUCHER_NO, APTCHCK2.INV_NUM, APTCHCK2.INV_DATE" & vbCrLf _
            & ", -1 * APTCHCK2.INV_AMT_APPLIED, -1 * APTCHCK2.INV_DISC_TAKEN" & vbCrLf _
            & " from APTCHCK1,APTCHCK2 " & vbCrLf _
            & " where APTCHCK1.BANK_CODE = APTCHCK2.BANK_CODE" & vbCrLf _
            & "   and APTCHCK1.CHECK_NUM = APTCHCK2.CHECK_NUM" & vbCrLf _
            & "   and APTCHCK1.OPS_YYYYPP_F = :PARM1" & vbCr
            Create_TDA(.Tables.Add, "APTCHCK2", "**", 0, False, "V", 4)

            .Relations.Add("APTCHCK2", _
            New DataColumn() { _
            .Tables("APTCHCK1").Columns("CHECK_STATUS"), _
            .Tables("APTCHCK1").Columns("BANK_CODE"), _
            .Tables("APTCHCK1").Columns("CHECK_NUM")}, _
            New DataColumn() { _
            .Tables("APTCHCK2").Columns("CHECK_STATUS"), _
            .Tables("APTCHCK2").Columns("BANK_CODE"), _
            .Tables("APTCHCK2").Columns("CHECK_NUM")})

            .Tables("APTCHCK1").Columns("APPL").Expression = "SUM(Child(APTCHCK2).INV_AMT_APPLIED)"
            .Tables("APTCHCK1").Columns("DISC").Expression = "SUM(Child(APTCHCK2).INV_DISC_TAKEN)"



            ASCMAIN1.sql = "Select APTINVH1.VOUCHER_NO" _
            & ", APTINVH1.VEND_CODE, APTVEND1.VEND_NAME, APTVEND1.PROCESSOR_CODE, APTVEND1.VEND_TYPE, APTVEND1.VEND_CLASS_CODE" _
            & ", APTINVH1.INV_TYPE, APTINVH1.INV_NUM, APTINVH1.INV_DATE, GLTCREC3.CREC_AMT INV_AMT" _
            & ", APTINVH1.INV_REF, APTINVH1.PO_ORDER_NO, APTINVH1.TERM_CODE" _
            & ", APTINVH1.INV_DUE_DATE, APTINVH1.OPS_YYYYPP_ACCRUE, APTINVH1.OPS_YYYYPP" _
            & ", APTINVH1.POST_CODE, APTINVH1.INV_PYMT_METHOD, APTINVH1.REGISTER_XNO, GLTJRNL1.JOURNAL_DATE" _
            & " from APTINVH1,GLTJRNL1,APTVEND1,GLTCREC3" _
            & " where GLTJRNL1.JOURNAL_TYPE (+) = 'APIN'" _
            & "   and GLTJRNL1.REGISTER_XNO (+) = APTINVH1.REGISTER_XNO" _
            & "   and APTVEND1.VEND_CODE (+) = APTINVH1.VEND_CODE" _
            & "   and GLTCREC3.OPS_YYYYPP = :PARM1" _
            & "   and GLTCREC3.CREC_TYPE_CODE = 'AP'" _
            & "   and GLTCREC3.DETL_CTL_NO = APTINVH1.VOUCHER_NO"
            Create_TDA(.Tables.Add, "APTINVHB", "**", 0, False, "V", 1)
            Create_TDA(.Tables.Add, "APTINVHE", "**", 0, False, "V", 1)

         
            Dim tbl As DataTable = dst.Tables("APTINVHB").Clone
            tbl.TableName = "APTINVHX"
            dst.Tables.Add(tbl)
            With dst.Tables("APTINVHX")
                .Columns.Add("PURCHASES", GetType(System.Double))
                .Columns.Add("PAYMENTS", GetType(System.Double))
                .Columns.Add("DISCOUNTS", GetType(System.Double))
                .Columns.Add("BALANCE", GetType(System.Double))
            End With


            Dim sql As String = "" _
            & "SELECT APTINVH2.*" & vbCrLf _
            & ", APTINVH1.VEND_CODE, APTINVH1.INV_TYPE" & vbCrLf _
            & ", APTINVH1.INV_NUM, APTINVH1.INV_DATE" & vbCrLf _
            & ", APTINVH1.OPS_YYYYPP, APTINVH1.OPS_YYYYPP_ACCRUE, APTINVH1.REGISTER_XNO" & vbCrLf _
            & " from APTINVH1, APTINVH2" & vbCrLf _
            & " where APTINVH1.VOUCHER_NO = APTINVH2.VOUCHER_NO" & vbCrLf _
            & " AND NVL(APTINVH1.REGISTER_IND,'0') <> 'D'" & vbCrLf _
            & " AND NVL(APTINVH1.INV_STATUS,'?') <> 'R'" & vbCr
            APTAEXPX = ASCMAIN1.Temp_Table(sql & " AND ROWNUM < 1")

            '            & "   and APTINVH1.OPS_YYYYPP <> :PARM1" & vbCrLF _

            sql_APTAEXPX = sql _
            & "   and APTINVH1.OPS_YYYYPP = :PARM1" & vbCrLf _
            & " union " & sql _
            & "   and APTINVH1.OPS_YYYYPP <> :PARM1" & vbCrLf _
            & "   and APTINVH1.OPS_YYYYPP_ACCRUE = :PARM1" & vbCr

            ASCMAIN1.sql = "Select APTVEND1.VEND_CODE, APTVEND1.VEND_NAME, APTVEND1.VEND_TYPE, APTVEND1.VEND_CLASS_CODE" _
            & " from APTVEND1 where APTVEND1.VEND_CODE in (Select Distinct VEND_CODE from " & APTAEXPX & ")"
            Create_TDA(.Tables.Add, "APTVEND1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select GLTACCT1.ACCT_CODE, GLTACCT1.ACCT_DESC" _
            & " from GLTACCT1 where GLTACCT1.ACCT_CODE in (Select Distinct ACCT_CODE from " & APTAEXPX & ")"
            Create_TDA(.Tables.Add, "GLTACCT1", "**", 0, False, "", 1)

            With .Tables("APTSUMV1")
                .Columns("BEG_BAL").DefaultValue = 0
                .Columns("APIN_ALL").DefaultValue = 0
                .Columns("APCD_APPL").DefaultValue = 0
                .Columns("APCD_DISC").DefaultValue = 0
                .Columns("END_BAL").DefaultValue = 0
            End With

            ASCMAIN1.sql = "SELECT APTAEXPX.*" & vbCrLf _
            & " from " & APTAEXPX & " APTAEXPX" & vbCr
            Create_TDA(.Tables.Add, "APTAEXP0", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "SELECT Distinct APTAEXPX.ACCT_CODE" & vbCrLf _
            & ", APTAEXPX.SEG2_CODE, APTAEXPX.SEG3_CODE, APTAEXPX.SEG4_CODE" & vbCrLf _
            & " from " & APTAEXPX & " APTAEXPX" & vbCr
            Create_TDA(.Tables.Add, "APTGEXP0", "**", 0, False, "V", 1)
            With dst.Tables("APTGEXP0")
                .Columns.Add("ACCT_DESC", GetType(System.String))
                .Columns.Add("INV_AMT", GetType(System.Double))
            End With

            ASCMAIN1.sql = "SELECT Distinct APTAEXPX.VEND_CODE" & vbCrLf _
            & " from " & APTAEXPX & " APTAEXPX" & vbCr
            Create_TDA(.Tables.Add, "APTVEXP0", "**", 0, False, "V", 1)
            With dst.Tables("APTVEXP0")
                .Columns.Add("VEND_NAME", GetType(System.String))
                .Columns.Add("INV_AMT", GetType(System.Double))
            End With

            ASCMAIN1.sql = "SELECT Distinct APTAEXPX.ACCT_CODE" & vbCrLf _
            & ", APTAEXPX.SEG2_CODE, APTAEXPX.SEG3_CODE, APTAEXPX.SEG4_CODE" & vbCrLf _
            & ", APTAEXPX.VEND_CODE" & vbCrLf _
            & " from " & APTAEXPX & " APTAEXPX" & vbCr
            Create_TDA(.Tables.Add, "APTHEXP0", "**", 0, False, "V", 1)
            With dst.Tables("APTHEXP0")
                .Columns.Add("VEND_NAME", GetType(System.String))
                .Columns.Add("ACCT_DESC", GetType(System.String))
                .Columns.Add("INV_AMT", GetType(System.Double))
            End With

            .Relations.Add("APTHEXP0_G", _
            New DataColumn() { _
            .Tables("APTGEXP0").Columns("ACCT_CODE"), _
            .Tables("APTGEXP0").Columns("SEG2_CODE"), _
            .Tables("APTGEXP0").Columns("SEG3_CODE"), _
            .Tables("APTGEXP0").Columns("SEG4_CODE")}, _
            New DataColumn() { _
            .Tables("APTHEXP0").Columns("ACCT_CODE"), _
            .Tables("APTHEXP0").Columns("SEG2_CODE"), _
            .Tables("APTHEXP0").Columns("SEG3_CODE"), _
            .Tables("APTHEXP0").Columns("SEG4_CODE")})

            .Relations.Add("APTHEXP0_V", _
            .Tables("APTVEXP0").Columns("VEND_CODE"), _
            .Tables("APTHEXP0").Columns("VEND_CODE"))

            .Relations.Add("APTAEXP0", _
            New DataColumn() { _
            .Tables("APTHEXP0").Columns("ACCT_CODE"), _
            .Tables("APTHEXP0").Columns("SEG2_CODE"), _
            .Tables("APTHEXP0").Columns("SEG3_CODE"), _
            .Tables("APTHEXP0").Columns("SEG4_CODE"), _
            .Tables("APTHEXP0").Columns("VEND_CODE")}, _
            New DataColumn() { _
            .Tables("APTAEXP0").Columns("ACCT_CODE"), _
            .Tables("APTAEXP0").Columns("SEG2_CODE"), _
            .Tables("APTAEXP0").Columns("SEG3_CODE"), _
            .Tables("APTAEXP0").Columns("SEG4_CODE"), _
            .Tables("APTAEXP0").Columns("VEND_CODE")})

            .Relations.Add("APTAEXP0_VEND_CODE", _
            .Tables("APTVEND1").Columns("VEND_CODE"), _
            .Tables("APTAEXP0").Columns("VEND_CODE"))

            .Relations.Add("APTHEXP0_VEND_CODE", _
            .Tables("APTVEND1").Columns("VEND_CODE"), _
            .Tables("APTHEXP0").Columns("VEND_CODE"))

            .Relations.Add("APTHEXP0_ACCT_CODE", _
            .Tables("GLTACCT1").Columns("ACCT_CODE"), _
            .Tables("APTHEXP0").Columns("ACCT_CODE"))

            .Relations.Add("APTVEXP0_VEND_CODE", _
            .Tables("APTVEND1").Columns("VEND_CODE"), _
            .Tables("APTVEXP0").Columns("VEND_CODE"))

            .Relations.Add("APTGEXP0_ACCT_CODE", _
            .Tables("GLTACCT1").Columns("ACCT_CODE"), _
            .Tables("APTGEXP0").Columns("ACCT_CODE"))

            With dst.Tables("APTVEXP0")
                .Columns("VEND_NAME").Expression = "PARENT(APTVEXP0_VEND_CODE).VEND_NAME"
                .Columns("INV_AMT").Expression = "SUM(CHILD.INV_AMT)"
            End With

            With dst.Tables("APTHEXP0")
                .Columns("VEND_NAME").Expression = "PARENT(APTHEXP0_VEND_CODE).VEND_NAME"
                .Columns("ACCT_DESC").Expression = "PARENT(APTHEXP0_ACCT_CODE).ACCT_DESC"
                .Columns("INV_AMT").Expression = "SUM(CHILD.INV_LINE_AMT)"
            End With

            With dst.Tables("APTGEXP0")
                .Columns("ACCT_DESC").Expression = "PARENT(APTGEXP0_ACCT_CODE).ACCT_DESC"
                .Columns("INV_AMT").Expression = "SUM(CHILD.INV_AMT)"
            End With

            .Relations.Add("APTINVH1_VEND_CODE", _
            .Tables("APTVEND1").Columns("VEND_CODE"), _
            .Tables("APTINVH1").Columns("VEND_CODE"))
            With .Tables("APTINVH1")
                .Columns("VEND_NAME").Expression = "PARENT(APTINVH1_VEND_CODE).VEND_NAME"
                .Columns("VEND_NAME").MaxLength = dst.Tables("APTVEND1").Columns("VEND_NAME").MaxLength
                .Columns("VEND_TYPE").Expression = "PARENT(APTINVH1_VEND_CODE).VEND_TYPE"
                .Columns("VEND_CLASS_CODE").Expression = "PARENT(APTINVH1_VEND_CODE).VEND_CLASS_CODE"
            End With


            ASCMAIN1.sql = "Select APTINVH1.VOUCHER_NO,APTINVH1.VEND_CODE" & vbCrLf _
                & ",APTINVH1.INV_TYPE,APTINVH1.INV_NUM,APTINVH1.INV_DATE,APTINVH1.INV_AMT, X.GL, X.PO" & vbCrLf _
                & ",INV_STATUS,APTINVH1.TERM_CODE,APTINVH1.INV_DUE_DATE,APTINVH1.INV_BALANCE" & vbCrLf _
                & ",APTINVH1.OPS_YYYYPP_ACCRUE,APTINVH1.OPS_YYYYPP,APTINVH1.POST_CODE" & vbCrLf _
                & ",APTINVH1.CHECK_NUM,APTINVH1.CHECK_DATE,APTINVH1.VEND_CODE_AP,APTINVH1.BANK_CODE" & vbCrLf _
                & ",APTINVH1.REGISTER_XNO,APTINVH1.INV_NOTE_INT,APTINVH1.VOUCHER_NO_ORIG" & vbCrLf _
                & " from APTINVH1, (" & vbCrLf _
                & "Select APTINVH2.VOUCHER_NO" & vbCrLf _
                & ", SUM (DECODE(APTINVH2.INV_LTYP,'P',0,APTINVH2.INV_LINE_AMT)) GL" & vbCrLf _
                & ", SUM (DECODE(APTINVH2.INV_LTYP,'P',APTINVH2.INV_LINE_AMT,0)) PO" & vbCrLf _
                & " from APTINVH1,APTINVH2 WHERE APTINVH1.VOUCHER_NO = APTINVH2.VOUCHER_NO" & vbCrLf _
                & " and APTINVH1.OPS_YYYYPP >= :PARM1 GROUP BY APTINVH2.VOUCHER_NO" & vbCrLf _
                & ") X where APTINVH1.OPS_YYYYPP_ACCRUE <= :PARM1 " & vbCrLf _
                & " and APTINVH1.OPS_YYYYPP > :PARM1" & vbCrLf _
                & " and REGISTER_XNO IS NOT NULL" & vbCrLf _
                & " and APTINVH1.VOUCHER_NO = X.VOUCHER_NO"
            If ASCMAIN1.CLIENT = "AHA" Then
                ASCMAIN1.sql &= vbCrLf & " UNION " & vbCrLf _
                & "Select APTINVH1.VOUCHER_NO,APTINVH1.VEND_CODE" & vbCrLf _
                & ",APTINVH1.INV_TYPE,APTINVH1.INV_NUM,APTINVH1.INV_DATE,APTINVH1.INV_AMT, -1 * X.GL, -1 * X.PO" & vbCrLf _
                & ",INV_STATUS,APTINVH1.TERM_CODE,APTINVH1.INV_DUE_DATE,APTINVH1.INV_BALANCE" & vbCrLf _
                & ",APTINVH1.OPS_YYYYPP_ACCRUE,APTINVH1.OPS_YYYYPP,APTINVH1.POST_CODE" & vbCrLf _
                & ",APTINVH1.CHECK_NUM,APTINVH1.CHECK_DATE,APTINVH1.VEND_CODE_AP,APTINVH1.BANK_CODE" & vbCrLf _
                & ",APTINVH1.REGISTER_XNO,APTINVH1.INV_NOTE_INT,APTINVH1.VOUCHER_NO_ORIG" & vbCrLf _
                & " from APTINVH1, (" & vbCrLf _
                & "Select APTINVH2.VOUCHER_NO" & vbCrLf _
                & ", SUM (DECODE(APTINVH2.INV_LTYP,'P',0,APTINVH2.INV_LINE_AMT)) GL" & vbCrLf _
                & ", SUM (DECODE(APTINVH2.INV_LTYP,'P',APTINVH2.INV_LINE_AMT,0)) PO" & vbCrLf _
                & " from APTINVH1,APTINVH2 WHERE APTINVH1.VOUCHER_NO = APTINVH2.VOUCHER_NO" & vbCrLf _
                & " and APTINVH1.OPS_YYYYPP = :PARM1 GROUP BY APTINVH2.VOUCHER_NO" & vbCrLf _
                & ") X where APTINVH1.OPS_YYYYPP_ACCRUE <= :PARM1 " & vbCrLf _
                & " and APTINVH1.OPS_YYYYPP = :PARM1" & vbCrLf _
                & " and REGISTER_XNO IS NOT NULL" & vbCrLf _
                & " and APTINVH1.VOUCHER_NO = X.VOUCHER_NO"

            End If
            Create_TDA(.Tables.Add, "APTINVH1_ACCRUED", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "" _
                & "Select APTINVH1.VEND_CODE, APTINVH1.INV_NUM, APTINVH1.INV_DATE, APTINVH1.REGISTER_XNO," & vbCrLf _
                & "APTINVH5.QTY_REC_NOT_INV * ICTIREC2.PO_COST ACCRPO, " & vbCrLf _
                & "APTINVH5.INV_QTY * APTINVH5.INV_COST APCOST," & vbCrLf _
                & "ICTIREC2.PO_ORDER_NO, ICTIREC2.PO_ORDER_LNO, ICTIREC2.ITEM_CODE," & vbCrLf _
                & "APTINVH5.* FROM APTINVH1,APTINVH5,ICTIREC2" & vbCrLf _
                & "WHERE APTINVH1.VOUCHER_NO = APTINVH5.VOUCHER_NO" & vbCrLf _
                & "AND APTINVH1.OPS_YYYYPP = :PARM1" & vbCrLf _
                & "AND ICTIREC2.RECEIPT_NO = APTINVH5.RECEIPT_NO" & vbCrLf _
                & "AND ICTIREC2.RECEIPT_LNO = APTINVH5.RECEIPT_LNO"
            Create_TDA(.Tables.Add, "APTINVH5", "**", 0, False, "V", 0)


        End With

        grdAPTINVH1_ACCRUED.DataSource = dst.Tables("APTINVH1_ACCRUED")
        grdAPTSUMV1.DataSource = dst.Tables("APTSUMV1")
        grdAPTGLAP1.DataSource = dst.Tables("APTGLAP1")
        grdAPTGLAP1_APIN.DataSource = dst.Tables("APTGLAP1_APIN")
        grdAPTGLAP1_APCD.DataSource = dst.Tables("APTGLAP1_APCD")
        grdAPTGLAP1_APGL.DataSource = dst.Tables("APTGLAP1_APGL")

        grdAPTINVH1.DataSource = dst.Tables("APTINVH1")
        grdAPTCHCK1.DataSource = dst.Tables("APTCHCK1")

        Dim tbl2 As DataTable = dst.Tables("APTINVHB").Clone
        grdAPTINVHB.DataSource = tbl2 ' dst.Tables("APTINVHB")
        'grdAPTINVHB.DataSource = dst.Tables("APTINVHB")

        grdAPTGEXP0.DataSource = dst.Tables("APTGEXP0")
        grdAPTVEXP0.DataSource = dst.Tables("APTVEXP0")
        grdAPTAEXP0.DataSource = dst.Tables("APTAEXP0")

        grdAPTINVH5.DataSource = dst.Tables("APTINVH5")

        Get_PARM("GLTPARM1")
        Get_PARM("APTPARM1")

        Set_SEGS(grdAPTGEXP0, "APTGEXP0")
        Set_SEGS(grdAPTVEXP0, "APTHEXP0_V")
        Set_SEGS(grdAPTAEXP0, "APTAEXP0")
        Set_SEGS(grdAPTINVH1, "APTINVH2")

        grdAPTINVH1.DisplayLayout.Bands("APTINVH2").SummaryFooterCaption = "Voucher Distribution Totals"
        grdAPTCHCK1.DisplayLayout.Bands("APTCHCK2").SummaryFooterCaption = "Check Remittance Advice Totals"

        grdAPTSUMV1.DisplayLayout.Bands("APTSUMV1").SummaryFooterCaption = "Totals" '  for Vendor: [VEND_CODE] [VEND_NAME]"
        grdAPTSUMV1.DisplayLayout.Bands("APTSUMV2").SummaryFooterCaption = "Total Invoices"
        grdAPTSUMV1.DisplayLayout.Bands("APTSUMV3").SummaryFooterCaption = "Total Checks"
        grdAPTSUMV1.DisplayLayout.Bands("APTSUMV4").SummaryFooterCaption = "Total Invoices Paid"

        grdAPTGEXP0.DisplayLayout.Bands("APTGEXP0").SummaryFooterCaption = "Grand Totals"
        grdAPTGEXP0.DisplayLayout.Bands("APTHEXP0_G").SummaryFooterCaption = "Account Totals"
        grdAPTGEXP0.DisplayLayout.Bands("APTAEXP0").SummaryFooterCaption = "Vendor Totals"

        grdAPTVEXP0.DisplayLayout.Bands("APTVEXP0").SummaryFooterCaption = "Grand Totals"
        grdAPTVEXP0.DisplayLayout.Bands("APTHEXP0_V").SummaryFooterCaption = "Vendor Totals"
        grdAPTVEXP0.DisplayLayout.Bands("APTAEXP0").SummaryFooterCaption = "Account Totals"

        Create_Lookup("GLTACCT1")

        Create_Summary(grdAPTINVH1_ACCRUED, "VOUCHER_NO", "Count")
        Create_Summary(grdAPTINVH1_ACCRUED, New String() {"INV_AMT", "GL", "PO", "INV_BALANCE"})

        Create_Summary(grdAPTGLAP1, "JOURNAL_DATE", "Count")
        Create_Summary(grdAPTGLAP1, New String() {"APIN_AP", "APIN_GL", "APCD_AP", "APCD_GL", "OTHER_GL", "NET_AP", "NET_GL"})

        Create_Summary(grdAPTGLAP1_APIN, "JOURNAL_DATE", "Count")
        Create_Summary(grdAPTGLAP1_APIN, "INV_AMT")

        Create_Summary(grdAPTGLAP1_APCD, "JOURNAL_DATE", "Count")
        Create_Summary(grdAPTGLAP1_APCD, New String() {"PYMT", "APPL", "DISC"})

        Create_Summary(grdAPTGLAP1_APGL, "JOURNAL_DATE", "Count")
        Create_Summary(grdAPTGLAP1_APGL, "AMT")

        Create_Summary(grdAPTSUMV1, "VEND_CODE", "Count")
        Create_Summary(grdAPTSUMV1, New String() {"BEG_BAL", "APIN_ALL", "APCD_PYMT", "APCD_DISC", "END_BAL"})

        Create_Summary(grdAPTSUMV1, New String() {"BEG_BAL_YTD", "APIN_ALL_YTD", "APCD_PYMT_YTD", "APCD_DISC_YTD", "END_BAL_YTD", "ACCRUED_IN", "ACCRUED_INTO", "OOBAL"})

        Create_Summary(grdAPTSUMV1, "VOUCHER_NO", "Count", "APTSUMV2")
        Create_Summary(grdAPTSUMV1, "INV_AMT", , "APTSUMV2")
        Create_Summary(grdAPTSUMV1, "CHECK_NUM", "Count", "APTSUMV3")
        Create_Summary(grdAPTSUMV1, "CHECK_AMT", , "APTSUMV3")
        Create_Summary(grdAPTSUMV1, "APPL", , "APTSUMV3")
        Create_Summary(grdAPTSUMV1, "DISC", , "APTSUMV3")
        Create_Summary(grdAPTSUMV1, "VOUCHER_NO", "Count", "APTSUMV4")
        Create_Summary(grdAPTSUMV1, "INV_AMT_APPLIED", , "APTSUMV4")
        Create_Summary(grdAPTSUMV1, "INV_DISC_TAKEN", , "APTSUMV4")

        Create_Summary(grdAPTINVH1, "VOUCHER_NO", "Count")
        Create_Summary(grdAPTINVH1, "INV_AMT")

        Create_Summary(grdAPTINVH1, "VOUCHER_LNO", "Count", "APTINVH2")
        Create_Summary(grdAPTINVH1, "INV_LINE_AMT", , "APTINVH2")

        Create_Summary(grdAPTCHCK1, "CHECK_NUM", "Count")
        Create_Summary(grdAPTCHCK1, "CHECK_AMT")
        Create_Summary(grdAPTCHCK1, "APPL")
        Create_Summary(grdAPTCHCK1, "DISC")

        Create_Summary(grdAPTCHCK1, "SEQ_NUM", "Count", "APTCHCK2")
        Create_Summary(grdAPTCHCK1, "INV_AMT_APPLIED", , "APTCHCK2")
        Create_Summary(grdAPTCHCK1, "INV_DISC_TAKEN", , "APTCHCK2")

        Create_Summary(grdAPTINVHB, "VOUCHER_NO", "Count")
        Create_Summary(grdAPTINVHB, "INV_AMT")

        Create_Summary(grdAPTGEXP0, "ACCT_CODE", "Count")
        Create_Summary(grdAPTGEXP0, "INV_AMT")
        Create_Summary(grdAPTGEXP0, "VEND_CODE", "Count", "APTHEXP0_G")
        Create_Summary(grdAPTGEXP0, "INV_AMT", , "APTHEXP0_G")
        Create_Summary(grdAPTGEXP0, "VOUCHER_NO", "Count", "APTAEXP0")
        Create_Summary(grdAPTGEXP0, "INV_LINE_AMT", , "APTAEXP0")

        Create_Summary(grdAPTVEXP0, "VEND_CODE", "Count")
        Create_Summary(grdAPTVEXP0, "INV_AMT")
        Create_Summary(grdAPTVEXP0, "ACCT_CODE", "Count", "APTHEXP0_V")
        Create_Summary(grdAPTVEXP0, "INV_AMT", , "APTHEXP0_V")
        Create_Summary(grdAPTVEXP0, "VOUCHER_NO", "Count", "APTAEXP0")
        Create_Summary(grdAPTVEXP0, "INV_LINE_AMT", , "APTAEXP0")

        Create_Summary(grdAPTAEXP0, "VOUCHER_NO", "Count")
        Create_Summary(grdAPTAEXP0, "INV_LINE_AMT")

        Create_Summary(grdAPTINVH5, "VOUCHER_NO", "Count")
        Create_Summary(grdAPTINVH5, New String() {"ACCRPO", "APCOST", "INV_QTY", "VAR_QTY", "VAR_AMT", "QTY_REC_NOT_INV"})


        With grdAPTGLAP1.DisplayLayout.Bands("APTGLAP1")
            .Columns("DESCRIPTION").Header.Fixed = True
        End With
        With grdAPTGLAP1_APIN.DisplayLayout.Bands("APTGLAP1_APIN")
            .Columns("JOURNAL_DATE").Header.Fixed = True
        End With
        With grdAPTGLAP1_APCD.DisplayLayout.Bands("APTGLAP1_APCD")
            .Columns("JOURNAL_DATE").Header.Fixed = True
        End With
        With grdAPTGLAP1_APGL.DisplayLayout.Bands("APTGLAP1_APGL")
            .Columns("JOURNAL_DATE").Header.Fixed = True
        End With

        With grdAPTSUMV1.DisplayLayout.Bands("APTSUMV1")
            .Columns("VEND_CODE").Header.Fixed = True
            .Columns("VEND_NAME").Header.Fixed = True
            '.Groups("Vendor").Header.Fixed = True
        End With

        With grdAPTINVH1_ACCRUED.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If gcol.Key = "GL" Then
                    gcol.CellAppearance.BackColor = Drawing.Color.LightGreen
                End If
            Next
        End With

        Sort_grdColumns(grdAPTGLAP1, "LINE_TYPE,JOURNAL_DATE")
        Sort_grdColumns(grdAPTGLAP1_APIN, "JOURNAL_DATE")
        Sort_grdColumns(grdAPTGLAP1_APCD, "JOURNAL_DATE")
        Sort_grdColumns(grdAPTGLAP1_APGL, "JOURNAL_DATE")
        Sort_grdColumns(grdAPTSUMV1, "VEND_CODE", , 0)
        Sort_grdColumns(grdAPTSUMV1, "VOUCHER_NO", , 1)

        Sort_grdColumns(grdAPTINVH1, "VOUCHER_NO")
        Sort_grdColumns(grdAPTCHCK1, "BANK_CODE,CHECK_NUM")

        Sort_grdColumns(grdAPTINVH1_ACCRUED, "VEND_CODE")

        Set_cmbYP("OPS_YYYYPP", ASCMAIN1.CYP, -37, 0, -1)
        ASCMAIN1.Add_Value_List(grdAPTINVH1_ACCRUED, "INV_STATUS", Nothing, New String() {":", "O:Open", "P:Paid", "D:Deleted", "H:Hold"})
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                Validate_Code("OPS_YYYYPP")
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "ExcelTEST"
                With grdAPTINVHB
                    .DisplayLayout.UseFixedHeaders = True
                    .DisplayLayout.Bands(0).Columns("VOUCHER_NO").Header.Fixed = True
                    .DisplayLayout.Bands(0).Columns("VOUCHER_NO").Header.FixedHeaderIndicator = UltraWinGrid.FixedHeaderIndicator.Button
                    .ActiveRow = .Rows(0)
                    .DisplayLayout.Override.FixedHeaderIndicator = UltraWinGrid.FixedHeaderIndicator.Button
                    .DisplayLayout.Override.RowSelectorHeaderStyle = UltraWinGrid.RowSelectorHeaderStyle.SeparateElement

                End With

            Case "Excel"
                Select Case tabMain.SelectedTab.Key
                    Case "Activity by Date"
                        Select Case tabActivity.SelectedTab.Key
                            Case "Activity by Date"
                                Set_format_grdAPTGLAP1(False)
                                Export_to_Excel(grdAPTGLAP1)
                                Set_format_grdAPTGLAP1(True)

                            Case "Voucher Activity"
                                Export_to_Excel(grdAPTGLAP1_APIN)
                            Case "Payment Activity"
                                Export_to_Excel(grdAPTGLAP1_APCD)
                            Case "GL Transactions Summary"
                                Export_to_Excel(grdAPTGLAP1_APGL)
                        End Select
                    Case "Summary by Vendor"
                        Export_to_Excel(grdAPTSUMV1)
                    Case "Purchases"
                        Export_to_Excel(grdAPTINVH1)
                    Case "Payments"
                        Export_to_Excel(grdAPTCHCK1)
                    Case "Open AP Items"
                        Export_to_Excel(grdAPTINVHB)
                    Case "GL Dist - Purchases"
                        Export_to_Excel(New UltraWinGrid.UltraGrid() {grdAPTGEXP0, grdAPTVEXP0, grdAPTAEXP0})

                End Select

            Case "Excel - All"
                Export_to_Excel_All()

            Case "Report"
                Print_Report_Begin()
                'Generate_Report("GLRTBAL1")
                Print_Report_End()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Excel").Settings.Enabled = iScreenMode
                    .Items("Excel - All").Settings.Enabled = iScreenMode

                    .Items("Excel").Visible = False
                    .Items("Excel - All").Visible = False
                End With

                If Not tf Then
                    .Groups("Show Activity").Visible = False
                    .Groups("Show Balances").Visible = False
                    .Groups("Show Distributions").Visible = False
                    .Groups("Show Vendor Stats For").Visible = False
                    .Groups("Show Filter").Visible = False
                End If
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabMain.Visible = tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("APTSUMV1").Rows.Clear()
        dst.Tables("APTSUMV2").Rows.Clear()
        dst.Tables("APTSUMV3").Rows.Clear()
        dst.Tables("APTSUMV4").Rows.Clear()

        dst.Tables("APTAEXP0").Rows.Clear()
        dst.Tables("APTGEXP0").Rows.Clear()
        dst.Tables("APTVEXP0").Rows.Clear()
        dst.Tables("APTHEXP0").Rows.Clear()

        dst.Tables("APTINVH1").Rows.Clear()
        dst.Tables("APTINVH2").Rows.Clear()
        dst.Tables("APTCHCK1").Rows.Clear()
        dst.Tables("APTCHCK2").Rows.Clear()
        dst.Tables("APTINVHB").Rows.Clear()
        dst.Tables("APTINVHE").Rows.Clear()

        dst.Tables("APTVEND1").Rows.Clear()
        dst.Tables("GLTACCT1").Rows.Clear()

        dst.Tables("APTGLAP1").Rows.Clear()
        dst.Tables("APTGLAP1_APIN").Rows.Clear()
        dst.Tables("APTGLAP1_APCD").Rows.Clear()
        dst.Tables("APTGLAP1_APGL").Rows.Clear()

        dst.EnforceConstraints = True

        tabMain.SelectedTab = tabMain.Tabs(0)

        Absx1.txtFor("OPS_YYYYPP").Text = ASCMAIN1.CYP
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading AP Account Reconciliation Data")
        Save_Header_Fields(UltraGroupBox1)
        dst.EnforceConstraints = False

        RYP = HFs("OPS_YYYYPP")

        Fill_Records("APTGLAP1_APIN", New Object() {RYP})
        Fill_Records("APTGLAP1_APCD", New Object() {RYP})
        Fill_Records("APTGLAP1_APGL", New Object() {RYP})

        Fill_Records("APTINVH1", New Object() {RYP})
        Fill_Records("APTINVH2", New Object() {RYP})

        ASCDATA1.ExecuteSQL("Truncate Table " & APTAEXPX)
        ASCDATA1.ExecuteSQL("Insert into " & APTAEXPX & " " & sql_APTAEXPX, "V", RYP)

        Fill_Records("APTINVH5", New Object() {RYP})

        Dim LEGEND As String = Absx1.txtFor("LEGEND").Text
        Dim mmm As String = Mid(Split(LEGEND, "(")(1), 1, 3)
        optDistributions.Items("0").DisplayText = "Posted in " & mmm
        optDistributions.Items("1").DisplayText = "Accrued into " & mmm
        optDistributions.Items("2").DisplayText = "Accrued in " & mmm
        optDistributions.Tag = "0"
        optDistributions.Value = "0"
        optDistributions.Tag = ""

        Load_Record_EXP()

        Fill_Records("APTCHCK1", New Object() {RYP})
        Fill_Records("APTCHCK2", New Object() {RYP})
        Fill_Records("APTINVH1_ACCRUED", New Object() {RYP})

        Fill_Records("APTVEND1")
        Fill_Records("GLTACCT1")

        ASCMAIN1.sql = "" _
        & "Select GLTCREC3.DETL_CVX_NO VEND_CODE" & vbCr _
        & ", SUM (GLTCREC3.CREC_AMT) BEG_BAL" & vbCr _
        & ", 0 APIN_ALL" & vbCr _
        & ", 0 APCD_PYMT" & vbCr _
        & ", 0 APCD_APPL" & vbCr _
        & ", 0 APCD_DISC" & vbCr _
        & ", 0 END_BAL" & vbCr _
        & ", 0 BEG_BAL_YTD, 0 APIN_ALL_YTD, 0 APCD_PYMT_YTD, 0 APCD_APPL_YTD, 0 APCD_DISC_YTD, 0 END_BAL_YTD" & vbCr _
        & ", 0 ACCRUED_IN" & vbCr _
        & ", 0 ACCRUED_INTO" & vbCr _
        & " from GLTCREC3 " & vbCr _
        & " where GLTCREC3.OPS_YYYYPP = '" & ASCMAIN1.Period_Calc(RYP, -1) & "'" & vbCr _
        & "   and GLTCREC3.CREC_TYPE_CODE = 'AP'" & vbCrLf _
        & " group by GLTCREC3.DETL_CVX_NO" & vbCr

        ASCMAIN1.sql &= " union " & vbCr _
        & "Select APTINVH1.VEND_CODE" & vbCr _
        & ", 0 BEG_BAL" & vbCr _
        & ", SUM (APTINVH1.INV_AMT) APIN_ALL" & vbCr _
        & ", 0 APCD_PYMT" & vbCr _
        & ", 0 APCD_APPL" & vbCr _
        & ", 0 APCD_DISC" & vbCr _
        & ", 0 END_BAL" & vbCr _
        & ", 0 BEG_BAL_YTD, 0 APIN_ALL_YTD, 0 APCD_PYMT_YTD, 0 APCD_APPL_YTD, 0 APCD_DISC_YTD, 0 END_BAL_YTD" & vbCr _
        & ", 0 ACCRUED_IN" & vbCr _
        & ", 0 ACCRUED_INTO" & vbCr _
        & " from APTINVH1 " & vbCr _
        & " where APTINVH1.OPS_YYYYPP = '" & RYP & "'" & vbCr _
        & "   and NVL(APTINVH1.REGISTER_IND,'0') <> 'D'" & vbCr _
        & "   and NVL(APTINVH1.INV_STATUS,'?') <> 'R'" & vbCr _
        & " group by APTINVH1.VEND_CODE" & vbCr

        ASCMAIN1.sql &= " union " & vbCr _
        & "Select APTCHCK2.VEND_CODE" & vbCr _
        & ", 0 BEG_BAL" & vbCr _
        & ", 0 APIN_ALL" & vbCr _
        & ", SUM (NVL(APTCHCK2.INV_AMT_APPLIED,0) - NVL(APTCHCK2.INV_DISC_TAKEN,0)) APCD_PYMT" & vbCr _
        & ", SUM (APTCHCK2.INV_AMT_APPLIED) APCD_APPL" & vbCr _
        & ", SUM (APTCHCK2.INV_DISC_TAKEN) APCD_DISC" & vbCr _
        & ", 0 END_BAL" & vbCr _
        & ", 0 BEG_BAL_YTD, 0 APIN_ALL_YTD, 0 APCD_PYMT_YTD, 0 APCD_APPL_YTD, 0 APCD_DISC_YTD, 0 END_BAL_YTD" & vbCr _
        & ", 0 ACCRUED_IN" & vbCr _
        & ", 0 ACCRUED_INTO" & vbCr _
        & " from APTCHCK1, APTCHCK2" & vbCr _
        & " where APTCHCK1.BANK_CODE = APTCHCK2.BANK_CODE" & vbCr _
        & "   and APTCHCK1.CHECK_NUM = APTCHCK2.CHECK_NUM" & vbCr _
        & "   and APTCHCK1.OPS_YYYYPP = '" & RYP & "'" & vbCr _
        & " group by APTCHCK2.VEND_CODE" & vbCr

        ASCMAIN1.sql &= " union " & vbCr _
        & "Select APTCHCK2.VEND_CODE" & vbCr _
        & ", 0 BEG_BAL" & vbCr _
        & ", 0 APIN_ALL" & vbCr _
        & ", SUM (-1 * (NVL(APTCHCK2.INV_AMT_APPLIED,0) - NVL(APTCHCK2.INV_DISC_TAKEN,0))) APCD_PYMT" & vbCr _
        & ", SUM (-1 * APTCHCK2.INV_AMT_APPLIED) APCD_APPL" & vbCr _
        & ", SUM (-1 * APTCHCK2.INV_DISC_TAKEN) APCD_DISC" & vbCr _
        & ", 0 END_BAL" & vbCr _
        & ", 0 BEG_BAL_YTD, 0 APIN_ALL_YTD, 0 APCD_PYMT_YTD, 0 APCD_APPL_YTD, 0 APCD_DISC_YTD, 0 END_BAL_YTD" & vbCr _
        & ", 0 ACCRUED_IN" & vbCr _
        & ", 0 ACCRUED_INTO" & vbCr _
        & " from APTCHCK1, APTCHCK2" & vbCr _
        & " where APTCHCK1.BANK_CODE = APTCHCK2.BANK_CODE" & vbCr _
        & "   and APTCHCK1.CHECK_NUM = APTCHCK2.CHECK_NUM" & vbCr _
        & "   and APTCHCK1.OPS_YYYYPP_F = '" & RYP & "'" & vbCr _
        & " group by APTCHCK2.VEND_CODE" & vbCr

        If RYP = ASCMAIN1.CYP Then
            ASCMAIN1.sql &= " union " & vbCr _
            & "Select APTINVH1.VEND_CODE" & vbCr _
            & ", 0 BEG_BAL" & vbCr _
            & ", 0 APIN_ALL" & vbCr _
            & ", 0 APCD_PYMT" & vbCr _
            & ", 0 APCD_APPL" & vbCr _
            & ", 0 APCD_DISC" & vbCr _
            & ", SUM (INV_BALANCE) END_BAL" & vbCr _
            & ", 0 BEG_BAL_YTD, 0 APIN_ALL_YTD, 0 APCD_PYMT_YTD, 0 APCD_APPL_YTD, 0 APCD_DISC_YTD, 0 END_BAL_YTD" & vbCr _
            & ", 0 ACCRUED_IN" & vbCr _
            & ", 0 ACCRUED_INTO" & vbCr _
            & " from APTINVH1 " & vbCr _
            & " where APTINVH1.INV_STATUS IN ('O','H')" & vbCr _
            & " group by APTINVH1.VEND_CODE" & vbCr
        Else
            ASCMAIN1.sql &= " union " & vbCr _
            & "Select GLTCREC3.DETL_CVX_NO VEND_CODE" & vbCr _
            & ", 0 BEG_BAL" & vbCr _
            & ", 0 APIN_ALL" & vbCr _
            & ", 0 APCD_PYMT" & vbCr _
            & ", 0 APCD_APPL" & vbCr _
            & ", 0 APCD_DISC" & vbCr _
            & ", SUM (GLTCREC3.CREC_AMT) END_BAL" & vbCr _
            & ", 0 BEG_BAL_YTD, 0 APIN_ALL_YTD, 0 APCD_PYMT_YTD, 0 APCD_APPL_YTD, 0 APCD_DISC_YTD, 0 END_BAL_YTD" & vbCr _
            & ", 0 ACCRUED_IN" & vbCr _
            & ", 0 ACCRUED_INTO" & vbCr _
            & " from GLTCREC3 " & vbCr _
            & " where GLTCREC3.OPS_YYYYPP = '" & RYP & "'" & vbCr _
            & "   and GLTCREC3.CREC_TYPE_CODE = 'AP'" & vbCrLf _
            & " group by GLTCREC3.DETL_CVX_NO" & vbCr
        End If

        ASCMAIN1.sql &= " union " & vbCr _
        & "Select GLTCREC3.DETL_CVX_NO VEND_CODE" & vbCr _
        & ", 0 BEG_BAL, 0 APIN_ALL, 0 APCD_PYMT, 0 APCD_APPL, 0 APCD_DISC, 0 END_BAL" & vbCr _
        & ", SUM (GLTCREC3.CREC_AMT) BEG_BAL_YTD" & vbCr _
        & ", 0 APIN_ALL_YTD" & vbCr _
        & ", 0 APCD_PYMT_YTD" & vbCr _
        & ", 0 APCD_APPL_YTD" & vbCr _
        & ", 0 APCD_DISC_YTD" & vbCr _
        & ", 0 END_BAL_YTD" & vbCr _
        & ", 0 ACCRUED_IN" & vbCr _
        & ", 0 ACCRUED_INTO" & vbCr _
        & " from GLTCREC3 " & vbCr _
        & " where GLTCREC3.OPS_YYYYPP = '" & Mid(RYP, 1, 4) & "01" & "'" & vbCr _
        & "   and GLTCREC3.CREC_TYPE_CODE = 'AP'" & vbCrLf _
        & " group by GLTCREC3.DETL_CVX_NO" & vbCr

        ASCMAIN1.sql &= " union " & vbCr _
        & "Select APTINVH1.VEND_CODE" & vbCr _
        & ", 0 BEG_BAL, 0 APIN_ALL, 0 APCD_PYMT, 0 APCD_APPL, 0 APCD_DISC, 0 END_BAL" & vbCr _
        & ", 0 BEG_BAL_YTD" & vbCr _
        & ", SUM (APTINVH1.INV_AMT) APIN_ALL_YTD" & vbCr _
        & ", 0 APCD_PYMT_YTD" & vbCr _
        & ", 0 APCD_APPL_YTD" & vbCr _
        & ", 0 APCD_DISC_YTD" & vbCr _
        & ", 0 END_BAL_YTD" & vbCr _
        & ", 0 ACCRUED_IN" & vbCr _
        & ", 0 ACCRUED_INTO" & vbCr _
        & " from APTINVH1 " & vbCr _
        & " where APTINVH1.OPS_YYYYPP >= '" & Mid(RYP, 1, 4) & "01" & "'" & vbCr _
        & "   and APTINVH1.OPS_YYYYPP <= '" & RYP & "'" & vbCr _
        & "   and NVL(APTINVH1.INV_STATUS,'?') <> 'R'" & vbCr _
        & " group by APTINVH1.VEND_CODE" & vbCr

        ASCMAIN1.sql &= " union " & vbCr _
        & "Select APTCHCK2.VEND_CODE" & vbCr _
        & ", 0 BEG_BAL, 0 APIN_ALL, 0 APCD_PYMT, 0 APCD_APPL, 0 APCD_DISC, 0 END_BAL" & vbCr _
        & ", 0 BEG_BAL_YTD" & vbCr _
        & ", 0 APIN_ALL_YTD" & vbCr _
        & ", SUM (NVL(APTCHCK2.INV_AMT_APPLIED,0) - NVL(APTCHCK2.INV_DISC_TAKEN,0)) APCD_PYMT_YTD" & vbCr _
        & ", SUM (APTCHCK2.INV_AMT_APPLIED) APCD_APPL_YTD" & vbCr _
        & ", SUM (APTCHCK2.INV_DISC_TAKEN) APCD_DISC_YTD" & vbCr _
        & ", 0 END_BAL_YTD" & vbCr _
        & ", 0 ACCRUED_IN" & vbCr _
        & ", 0 ACCRUED_INTO" & vbCr _
        & " from APTCHCK1, APTCHCK2" & vbCr _
        & " where APTCHCK1.BANK_CODE = APTCHCK2.BANK_CODE" & vbCr _
        & "   and APTCHCK1.CHECK_NUM = APTCHCK2.CHECK_NUM" & vbCr _
        & "   and APTCHCK1.OPS_YYYYPP >= '" & Mid(RYP, 1, 4) & "01" & "'" & vbCr _
        & "   and APTCHCK1.OPS_YYYYPP <= '" & RYP & "'" & vbCr _
        & " group by APTCHCK2.VEND_CODE" & vbCr

        ASCMAIN1.sql &= " union " & vbCr _
        & "Select APTCHCK2.VEND_CODE" & vbCr _
        & ", 0 BEG_BAL, 0 APIN_ALL, 0 APCD_PYMT, 0 APCD_APPL, 0 APCD_DISC, 0 END_BAL" & vbCr _
        & ", 0 BEG_BAL_YTD" & vbCr _
        & ", 0 APIN_ALL_YTD" & vbCr _
        & ", SUM (-1 * (NVL(APTCHCK2.INV_AMT_APPLIED,0) - NVL(APTCHCK2.INV_DISC_TAKEN,0))) APCD_PYMT_YTD" & vbCr _
        & ", SUM (-1 * APTCHCK2.INV_AMT_APPLIED) APCD_APPL_YTD" & vbCr _
        & ", SUM (-1 * APTCHCK2.INV_DISC_TAKEN) APCD_DISC_YTD" & vbCr _
        & ", 0 END_BAL_YTD" & vbCr _
        & ", 0 ACCRUED_IN" & vbCr _
        & ", 0 ACCRUED_INTO" & vbCr _
        & " from APTCHCK1, APTCHCK2" & vbCr _
        & " where APTCHCK1.BANK_CODE = APTCHCK2.BANK_CODE" & vbCr _
        & "   and APTCHCK1.CHECK_NUM = APTCHCK2.CHECK_NUM" & vbCr _
        & "   and APTCHCK1.OPS_YYYYPP_F >= '" & Mid(RYP, 1, 4) & "01" & "'" & vbCr _
        & "   and APTCHCK1.OPS_YYYYPP_F <= '" & RYP & "'" & vbCr _
        & " group by APTCHCK2.VEND_CODE" & vbCr


        ASCMAIN1.sql &= " union " & vbCr _
        & "Select APTAEXPX.VEND_CODE" & vbCr _
        & ", 0 BEG_BAL, 0 APIN_ALL, 0 APCD_PYMT, 0 APCD_APPL, 0 APCD_DISC, 0 END_BAL" & vbCr _
        & ", 0 BEG_BAL_YTD" & vbCr _
        & ", 0 APIN_ALL_YTD" & vbCr _
        & ", 0 APCD_PYMT_YTD" & vbCr _
        & ", 0 APCD_APPL_YTD" & vbCr _
        & ", 0 APCD_DISC_YTD" & vbCr _
        & ", 0 END_BAL_YTD" & vbCr _
        & ", SUM (DECODE(APTAEXPX.OPS_YYYYPP_ACCRUE,'" & RYP & "',0,APTAEXPX.INV_LINE_AMT)) ACCRUED_IN" & vbCr _
        & ", SUM (DECODE(APTAEXPX.OPS_YYYYPP_ACCRUE,'" & RYP & "',APTAEXPX.INV_LINE_AMT,0)) ACCRUED_INTO" & vbCr _
        & " from " & APTAEXPX & " APTAEXPX" & vbCr _
        & " where APTAEXPX.OPS_YYYYPP_ACCRUE is not Null" & vbCr _
        & " group by APTAEXPX.VEND_CODE" & vbCr

        ASCMAIN1.sql = "Select VEND_CODE" & vbCr _
        & ", Sum (BEG_BAL) BEG_BAL" & vbCr _
        & ", Sum (APIN_ALL) APIN_ALL" & vbCr _
        & ", Sum (APCD_PYMT) APCD_PYMT" & vbCr _
        & ", Sum (APCD_APPL) APCD_APPL" & vbCr _
        & ", Sum (APCD_DISC) APCD_DISC" & vbCr _
        & ", Sum (END_BAL) END_BAL" & vbCr _
        & ", Sum (BEG_BAL_YTD) BEG_BAL_YTD" & vbCr _
        & ", Sum (APIN_ALL_YTD) APIN_ALL_YTD" & vbCr _
        & ", Sum (APCD_PYMT_YTD) APCD_PYMT_YTD" & vbCr _
        & ", Sum (APCD_APPL_YTD) APCD_APPL_YTD" & vbCr _
        & ", Sum (APCD_DISC_YTD) APCD_DISC_YTD" & vbCr _
        & ", Sum (END_BAL) END_BAL_YTD" & vbCr _
        & ", Sum (ACCRUED_IN) ACCRUED_IN" & vbCr _
        & ", Sum (ACCRUED_INTO) ACCRUED_INTO" & vbCr _
        & " from (" & ASCMAIN1.sql & ")" & vbCr _
        & " group by VEND_CODE" & vbCr

        If APTAPGLV = "" Then
            APTAPGLV = ASCMAIN1.Temp_Table
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & APTAPGLV)
            ASCDATA1.ExecuteSQL("Insert into " & APTAPGLV & " " & ASCMAIN1.sql)
        End If

        'ASCMAIN1.sql = "Select APTVEND1.VEND_CODE, APTVEND1.VEND_NAME, APTVEND1.VEND_TYPE, APTVEND1.VEND_CLASS_CODE" _
        '& ", APTAPGLV.BEG_BAL" _
        '& " from APTVEND1, " & APTAPGLV & " APTAPGLV where APTAPGLV.VEND_CODE = APTVEND1.VEND_CODE"
        'Fill_Records("APTSUMV1", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select APTVEND1.VEND_CODE, APTVEND1.VEND_NAME, APTVEND1.VEND_TYPE, APTVEND1.VEND_CLASS_CODE" _
        & ", APTAPGLV.BEG_BAL, APTAPGLV.APIN_ALL, APTAPGLV.APCD_PYMT, APTAPGLV.APCD_APPL, APTAPGLV.APCD_DISC, APTAPGLV.END_BAL" _
        & ", APTAPGLV.BEG_BAL_YTD, APTAPGLV.APIN_ALL_YTD, APTAPGLV.APCD_PYMT_YTD, APTAPGLV.APCD_APPL_YTD, APTAPGLV.APCD_DISC_YTD, APTAPGLV.END_BAL_YTD" _
        & ", APTAPGLV.ACCRUED_IN, APTAPGLV.ACCRUED_INTO" _
        & " from APTVEND1, " & APTAPGLV & " APTAPGLV where APTAPGLV.VEND_CODE = APTVEND1.VEND_CODE"

        Fill_Records("APTSUMV1", "", True, ASCMAIN1.sql)
        Fill_Records("APTSUMV2", RYP)
        Fill_Records("APTSUMV3", RYP)
        Fill_Records("APTSUMV4", RYP)

        'dst.EnforceConstraints = True

        Fill_Records("APTINVHB", ASCMAIN1.Period_Calc(RYP, -1))
        If RYP = ASCMAIN1.CYP Then
            ASCMAIN1.sql = "Select APTINVH1.VOUCHER_NO" _
            & ", APTINVH1.VEND_CODE, APTVEND1.VEND_NAME, APTVEND1.PROCESSOR_CODE, APTVEND1.VEND_TYPE, APTVEND1.VEND_CLASS_CODE" _
            & ", APTINVH1.INV_TYPE, APTINVH1.INV_NUM, APTINVH1.INV_DATE, APTINVH1.INV_BALANCE INV_AMT" _
            & ", APTINVH1.INV_REF, APTINVH1.PO_ORDER_NO, APTINVH1.TERM_CODE" _
            & ", APTINVH1.INV_DUE_DATE, APTINVH1.OPS_YYYYPP_ACCRUE, APTINVH1.OPS_YYYYPP" _
            & ", APTINVH1.POST_CODE, APTINVH1.INV_PYMT_METHOD, APTINVH1.REGISTER_XNO, GLTJRNL1.JOURNAL_DATE" _
            & " from APTINVH1,GLTJRNL1,APTVEND1" _
            & " where GLTJRNL1.JOURNAL_TYPE (+) = 'APIN'" _
            & "   and GLTJRNL1.REGISTER_XNO (+) = APTINVH1.REGISTER_XNO" _
            & "   and APTVEND1.VEND_CODE  (+) = APTINVH1.VEND_CODE" _
            & "   and APTINVH1.INV_STATUS in ('O','H')"
            Fill_Records("APTINVHE", "", , ASCMAIN1.sql)
        Else
            Fill_Records("APTINVHE", RYP)
        End If

        Set_BE()

        'Accrual_Summary_by_Vendor()
        Fill_APTGLAP1()

        Setup_tabMain()
        optShowAPIN.CheckedIndex = 0
        optShowAPCD.CheckedIndex = 0
        optShowGL.CheckedIndex = 0

        Setup_grdAPTSUMV1()

        ASCMAIN1.Progress("")
    End Sub

    'Sub Accrual_Summary_by_Vendor()

    'End Sub

    Sub Load_Record_EXP()
        Load_Record_EXP_1("APTAEXP0")
        Load_Record_EXP_1("APTGEXP0")
        Load_Record_EXP_1("APTVEXP0")
        Load_Record_EXP_1("APTHEXP0")

        'Fill_Records("APTAEXP0", New Object() {RYP})
        'Fill_Records("APTGEXP0", New Object() {RYP})
        'Fill_Records("APTVEXP0", New Object() {RYP})
        'Fill_Records("APTHEXP0", New Object() {RYP})
    End Sub

    Sub Load_Record_EXP_1(ByVal TABLE_NAME As String)
        Dim SQL As String = Get_SelectCommand(TABLE_NAME)

        Dim CODE As String = optDistributions.Value
        Select Case CODE
            Case "0"
                SQL = SQL _
                & " where APTAEXPX.OPS_YYYYPP = :PARM1" & vbCr
            Case "1"
                SQL = SQL _
                & " where APTAEXPX.OPS_YYYYPP_ACCRUE = :PARM1" & vbCr _
                & "   and APTAEXPX.INV_LTYP IS NULL"
            Case "2"
                SQL = SQL _
                & " where APTAEXPX.OPS_YYYYPP = :PARM1" & vbCr _
                & "   and APTAEXPX.OPS_YYYYPP_ACCRUE is Not Null" & vbCr _
                & "   and APTAEXPX.INV_LTYP IS NULL"
        End Select

        Fill_Records(TABLE_NAME, New Object() {RYP}, True, SQL)
    End Sub

    Sub Update_Record()
        BeginTrans()
        CommitTrans("Update Complete")
    End Sub
#End Region


#Region "Popup_Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdAPTAEXP0, "SSB", "Show Filter", "Show Groupbox", "Voucher Inquiry")
        Load_Popup_Menu(grdAPTINVH5, "SSB", "Show Filter", "Show Groupbox", "Voucher Inquiry")
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

        'If tlb_pop.Tools.Exists("Show Description") Then
        '    tlb_sbt = DirectCast(tlb_pop.Tools("Show Description"), UltraWinToolbars.StateButtonTool)
        '    tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns("ITEM_DESC").Hidden
        'End If
 

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

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            'Case "Show Description"
            '    Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
            '    grd.DisplayLayout.Bands(0).Columns("ITEM_DESC").Hidden = Not tlb_sbt.Checked


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Voucher Inquiry"
                Dim VOUCHER_NO As String = grd.ActiveRow.Cells("VOUCHER_NO").Value
                Dim rowAPTINVH1 As DataRow = LookUp("APTINVH1", VOUCHER_NO)
                If rowAPTINVH1 IsNot Nothing Then
                    Context_Launch("Load", VOUCHER_NO, e.Tool.Key, "APFINVHI")
                End If

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI")
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        'If EntryMode <> "" Then
        '    Exit Sub
        'End If

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

            Case "OPS_YYYYPP"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    Click_Command("Load", e)
                End If

        End Select

    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)

        End Select
    End Sub

#End Region

    Sub Fill_APTGLAP1()
        With dst.Tables("APTGLAP1")

            .Rows.Clear()

            Dim rowAPTGLAP1 As DataRow = .NewRow
            rowAPTGLAP1("LINE_TYPE") = "0"
            rowAPTGLAP1("DESCRIPTION") = "Opening Balance"
            rowAPTGLAP1("BEG_AP") = dst.Tables("APTINVHB").Compute("SUM(INV_AMT)", "")

            ASCMAIN1.sql = "Select * from GLTACCT3" _
            & " where ACCT_CODE in (Select ACCT_CODE from APTPOST1)" _
            & "   and ACCT_YEAR <= '" & Mid(RYP, 1, 4) & "'"
            Dim GL_BEG_BAL As Double = 0
            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                'GL_BEG_BAL += Val(row("ACCT_BEG_BAL") & "")
                Dim imax As Integer
                If row("ACCT_YEAR") = Mid(RYP, 1, 4) Then
                    imax = Val(Mid(RYP, 5, 2)) - 1
                Else
                    imax = 12
                End If
                If imax <> 0 Then
                    For i As Integer = 1 To imax
                        GL_BEG_BAL += Val(row("ACCT_ACT_P" & Format(i, "00")) & "")
                    Next
                End If
            Next

            rowAPTGLAP1("BEG_GL") = GL_BEG_BAL
            .Rows.Add(rowAPTGLAP1)
        End With

        Fill_APTGLAP1_by_Date()

    End Sub

    Sub Fill_APTGLAP1_by_Date( _
    Optional ByVal SINGLE_TABLE_NAME As String = "", _
    Optional ByVal COLUMN_NAME As String = "", _
    Optional ByVal COLUMN_VALUE As String = "")

        Dim LINE_TYPE As String = ""
        Dim sql As String

        Dim sfx As String = ""
        If COLUMN_NAME = "" Then
            sfx = ""
        Else
            sfx = "_X" & COLUMN_VALUE
        End If

        Dim rowAPTGLAP1 As DataRow

        With dst.Tables("APTGLAP1")
            For Each TABLE_NAME As String In New String() _
            {"APTGLAP1_APIN", "APTGLAP1_APCD", "APTGLAP1_APGL"}
                If SINGLE_TABLE_NAME = "" Or SINGLE_TABLE_NAME = TABLE_NAME Then
                    If COLUMN_NAME = "" Or (TABLE_NAME = "APTGLAP1_APCD" And COLUMN_VALUE = ".") Then
                        sql = ""
                    Else
                        sql = COLUMN_NAME & " = '" & COLUMN_VALUE & "'"
                    End If
                    For Each row As DataRow In dst.Tables(TABLE_NAME).Select(sql, "")
                        If row("JOURNAL_DATE") & "" = "" Then
                            LINE_TYPE = "2"
                            sql = "JOURNAL_DATE is Null"
                            'sql = "JOURNAL_DATE = '" & Format(Now, "MM/dd/yyyy") & "'"
                        Else
                            LINE_TYPE = "1"
                            sql = "JOURNAL_DATE = '" & Format(row("JOURNAL_DATE"), "MM/dd/yyyy") & "'"
                        End If
                        sql = "LINE_TYPE = '" & LINE_TYPE & "' and " & sql
                        Dim rows() As DataRow = .Select(sql, "")
                        If rows.Length = 0 Then
                            rowAPTGLAP1 = .NewRow
                            rowAPTGLAP1("LINE_TYPE") = LINE_TYPE
                            rowAPTGLAP1("JOURNAL_DATE") = row("JOURNAL_DATE")
                            If LINE_TYPE = "2" Then
                                rowAPTGLAP1("DESCRIPTION") = "Unposted"
                            Else
                                rowAPTGLAP1("DESCRIPTION") = Format(row("JOURNAL_DATE"), "MM/dd/yyyy")
                            End If
                            .Rows.Add(rowAPTGLAP1)
                        Else
                            rowAPTGLAP1 = rows(0)
                        End If

                        Select Case TABLE_NAME
                            Case "APTGLAP1_APIN"
                                rowAPTGLAP1("APIN_AP" & sfx) = Val(rowAPTGLAP1("APIN_AP" & sfx) & "") + Val(row("INV_AMT") & "")
                            Case "APTGLAP1_APCD"
                                If COLUMN_NAME = "" Then
                                    rowAPTGLAP1("APCD_AP" & sfx) = Val(rowAPTGLAP1("APCD_AP" & sfx) & "") + Val(row("APPL") & "")
                                Else
                                    If COLUMN_VALUE = "." Then
                                        rowAPTGLAP1("APCD_AP" & sfx) = Val(rowAPTGLAP1("APCD_AP" & sfx) & "") + Val(row("DISC") & "")
                                    Else
                                        rowAPTGLAP1("APCD_AP" & sfx) = Val(rowAPTGLAP1("APCD_AP" & sfx) & "") + Val(row("PYMT") & "")
                                    End If
                                End If
                            Case "APTGLAP1_APGL"
                                Select Case row("JOURNAL_TYPE")
                                    Case "APIN"
                                        rowAPTGLAP1("APIN_GL" & sfx) = Val(rowAPTGLAP1("APIN_GL" & sfx) & "") + Val(row("AMT") & "")
                                    Case "APCD"
                                        rowAPTGLAP1("APCD_GL" & sfx) = Val(rowAPTGLAP1("APCD_GL" & sfx) & "") + Val(row("AMT") & "")
                                    Case Else
                                        rowAPTGLAP1("OTHER_GL" & sfx) = Val(rowAPTGLAP1("OTHER_GL" & sfx) & "") + Val(row("AMT") & "")
                                End Select
                        End Select
                    Next
                End If
            Next
        End With

    End Sub
    Private Sub optBalances_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optBalances.ValueChanged
        If dst.Tables.Count = 0 Then Exit Sub

        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()
        Set_BE()
        Me.Cursor = Cursors.Default
    End Sub

    Sub Set_BE()

        If dst.Tables.Count = 0 Then Exit Sub

        Dim BE As String = optBalances.Value

        'grdAPTINVHB.DataSource = dst.Tables("APTINVH" & BE)

        Dim tbl2 As DataTable = DirectCast(grdAPTINVHB.DataSource, DataTable)
        tbl2.Rows.Clear()
        tbl2.Merge(dst.Tables("APTINVH" & BE))

        grdAPTINVHB.Text = "Open AP Items - " & optBalances.Text & " Balance for " & Absx1.txtFor("LEGEND").Text

    End Sub

    Private Sub tabMain_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged

        If EntryMode = "" Then
            Exit Sub
        End If
        Setup_tabMain()

    End Sub

    Sub Setup_tabMain()
        With UltraExplorerBar1
            .Groups("Show Activity").Visible = False
            .Groups("Show Balances").Visible = False
            .Groups("Show Distributions").Visible = False
            .Groups("Show Vendor Stats For").Visible = False
            .Groups("Show Filter").Visible = False

            Select Case tabMain.ActiveTab.Key
                Case "Activity by Date"
                    .Groups("Show Activity").Visible = True
                Case "Summary by Vendor"
                    .Groups("Show Vendor Stats For").Visible = True
                    .Groups("Show Filter").Visible = True
                Case "Open AP Items"
                    .Groups("Show Balances").Visible = True
                Case "GL Dist - Purchases"
                    .Groups("Show Distributions").Visible = True
                Case "Purchases", "Payments"
                    .Groups("Show Filter").Visible = True

            End Select
        End With
    End Sub

    Private Sub optDistributions_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optDistributions.ValueChanged
        If dst.Tables.Count = 0 Then Exit Sub
        If optDistributions.Tag = "0" Then Exit Sub

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Distribution Data")
        Application.DoEvents()

        dst.EnforceConstraints = False
        Load_Record_EXP()
        'dst.EnforceConstraints = True
        ' fix why I cannot enable constraints

        grdAPTGEXP0.Text = "GL by Account - " & optDistributions.Text
        grdAPTVEXP0.Text = "GL by Vendor - " & optDistributions.Text
        grdAPTAEXP0.Text = "GL Details - " & optDistributions.Text


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub chkShowFilter_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowFilter.CheckedChanged
        Setup_grd_for_filter(grdAPTINVH1)
        Setup_grd_for_filter(grdAPTCHCK1)
        Setup_grd_for_filter(grdAPTSUMV1)
    End Sub

    Sub Setup_grd_for_filter(ByVal grd As UltraWinGrid.UltraGrid)
        grd.DisplayLayout.Override.FilterUIType = UltraWinGrid.FilterUIType.FilterRow
        grd.DisplayLayout.Override.FilterClearButtonLocation = UltraWinGrid.FilterClearButtonLocation.Row
        grd.DisplayLayout.Override.FilterRowAppearance.BackColor = System.Drawing.Color.AliceBlue
        If chkShowFilter.Checked Then
            grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True
        Else
            grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.False
        End If
    End Sub
 
    Private Sub chkShowGL_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowGL.CheckedChanged
        If dst.Tables.Count = 0 Then Exit Sub

        With grdAPTGLAP1.DisplayLayout.Bands("APTGLAP1")
            .Columns("APIN_GL").Hidden = Not chkShowGL.Checked
            .Columns("APCD_GL").Hidden = Not chkShowGL.Checked
            .Columns("OTHER_GL").Hidden = Not chkShowGL.Checked
            .Columns("NET_GL").Hidden = Not chkShowGL.Checked
        End With

        If chkShowGL.Checked Then
            optShowGL.Visible = True
        Else
            optShowGL.CheckedIndex = 0
            optShowGL.Visible = False
        End If
    End Sub

    Private Sub optShowAPIN_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optShowAPIN.ValueChanged
        If dst.Tables.Count = 0 Then Exit Sub
        Dim COLUMN_NAME As String = optShowAPIN.Value
        Setup_grdAPTGLAP1("APIN_AP", COLUMN_NAME)

    End Sub

    Sub Setup_grdAPTGLAP1( _
    ByVal DATA_TYPE As String, _
    ByVal COLUMN_NAME As String)

        grdAPTGLAP1.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show

        With dst.Tables("APTGLAP1").Columns
            For i As Integer = .Count - 1 To 0 Step -1
                Dim X_COLUMN_NAME As String = .Item(i).ColumnName
                If X_COLUMN_NAME Like DATA_TYPE & "_X*" Then
                    .Remove(X_COLUMN_NAME)
                    With grdAPTGLAP1.DisplayLayout.Bands(0)
                        .Summaries.Remove(.Summaries(X_COLUMN_NAME))
                    End With
                End If
            Next

            If COLUMN_NAME <> "0" Then
                Dim tbl As DataTable = ASCMAIN1.Distinct_Values("APTGLAP1_" & Mid(DATA_TYPE, 1, 4), _
                  dst.Tables("APTGLAP1_" & Mid(DATA_TYPE, 1, 4)), _
                  COLUMN_NAME)
                If DATA_TYPE = "APCD_AP" Then
                    tbl.Rows.Add(New String() {"."})
                End If
                For Each row As DataRow In tbl.Rows
                    Dim X_COLUMN_NAME As String = DATA_TYPE & "_X" & row(0)
                    .Add(X_COLUMN_NAME, GetType(System.Double))
                    With grdAPTGLAP1.DisplayLayout.Bands(0).Columns(X_COLUMN_NAME)
                        If row(0) = "." Then
                            .Header.Caption = "Disc"
                            .Width = grdAPTGLAP1.DisplayLayout.Bands(0).Columns(DATA_TYPE).Width / 2
                        Else
                            .Width = grdAPTGLAP1.DisplayLayout.Bands(0).Columns(DATA_TYPE).Width
                            .Header.Caption = row(0)
                        End If
                        .CellAppearance.BackColor = grdAPTGLAP1.DisplayLayout.Bands(0).Columns(DATA_TYPE).CellAppearance.BackColor

                        .Format = grdAPTGLAP1.DisplayLayout.Bands(0).Columns(DATA_TYPE).Format
                        .Header.Appearance.TextHAlign = HAlign.Right
                        .CellAppearance.TextHAlign = HAlign.Right
                        .Header.VisiblePosition = grdAPTGLAP1.DisplayLayout.Bands(0).Columns(DATA_TYPE).Header.VisiblePosition
                    End With
                    Create_Summary(grdAPTGLAP1, X_COLUMN_NAME)
                    Fill_APTGLAP1_by_Date("APTGLAP1_" & Mid(DATA_TYPE, 1, 4), COLUMN_NAME, row(0))
                Next

            End If
        End With

    End Sub
    Private Sub optShowAPCD_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optShowAPCD.ValueChanged

        If dst.Tables.Count = 0 Then Exit Sub
        Dim COLUMN_NAME As String = optShowAPCD.Value
        Setup_grdAPTGLAP1("APCD_AP", COLUMN_NAME)

    End Sub
 
    Private Sub grdAPTGLAP1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdAPTGLAP1.InitializeRow
        'If e.Row.Cells("LINE_TYPE").Text = "0" Then
        '    e.Row.Cells("DESCRIPTION").Value = "Opening Balance"
        'Else
        '    e.Row.Cells("DESCRIPTION").Value = e.Row.Cells("JOURNAL_DATE").Value
        'End If
    End Sub

    Private Sub optShowGL_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optShowGL.ValueChanged

        If dst.Tables.Count = 0 Then Exit Sub
        Dim COLUMN_NAME As String = optShowGL.Value
        'Setup_grdAPTGLAP1("APGL", COLUMN_NAME)

    End Sub

    Private Sub optXTD_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optXTD.ValueChanged
        If dst.Tables.Count = 0 Then Exit Sub

        Setup_grdAPTSUMV1()
    End Sub

    Sub Setup_grdAPTSUMV1()
        With grdAPTSUMV1.DisplayLayout.Bands("APTSUMV1")
            .Columns("BEG_BAL").Hidden = (optXTD.Value = "Y")
            .Columns("APIN_ALL").Hidden = (optXTD.Value = "Y")
            .Columns("APCD_PYMT").Hidden = (optXTD.Value = "Y")
            .Columns("APCD_APPL").Hidden = True
            .Columns("APCD_DISC").Hidden = (optXTD.Value = "Y")
            .Columns("END_BAL").Hidden = (optXTD.Value = "Y")

            .Columns("BEG_BAL_YTD").Hidden = (optXTD.Value = "M")
            .Columns("APIN_ALL_YTD").Hidden = (optXTD.Value = "M")
            .Columns("APCD_PYMT_YTD").Hidden = (optXTD.Value = "M")
            .Columns("APCD_APPL_YTD").Hidden = True
            .Columns("APCD_DISC_YTD").Hidden = (optXTD.Value = "M")
            .Columns("END_BAL_YTD").Hidden = (optXTD.Value = "M")
        End With

    End Sub

    Private Sub chkSingleBand_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSingleBand.CheckedChanged
        If dst.Tables.Count = 0 Then Exit Sub

        If chkSingleBand.Checked Then
            grdAPTINVH1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand
            grdAPTCHCK1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand
            grdAPTSUMV1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand
        Else
            grdAPTINVH1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.MultiBand
            grdAPTCHCK1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.MultiBand
            grdAPTSUMV1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.MultiBand
        End If

    End Sub

    Sub Export_to_Excel_All()

        Me.Cursor = Cursors.WaitCursor
        Dim myWorkbook As New Infragistics.Documents.Excel.Workbook

        Set_format_grdAPTGLAP1(False)
        Export_to_Excel_Add_grd(myWorkbook, grdAPTGLAP1, False, "Summary Activity by Date")
        Set_format_grdAPTGLAP1(True)

        Export_to_Excel_Add_grd(myWorkbook, grdAPTGLAP1_APIN, False, "Voucher Activity by Date")
        Export_to_Excel_Add_grd(myWorkbook, grdAPTGLAP1_APCD, False, "Payment Activity by Date")
        Export_to_Excel_Add_grd(myWorkbook, grdAPTGLAP1_APGL, False, "GL Activity by Date")

        grdAPTSUMV1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand
        optXTD.Value = "M"
        Export_to_Excel_Add_grd(myWorkbook, grdAPTSUMV1, False, "Summary by Vendor - MTD")
        optXTD.Value = "Y"
        Export_to_Excel_Add_grd(myWorkbook, grdAPTSUMV1, False, "Summary by Vendor - YTD")
        If Not chkSingleBand.Checked Then
            grdAPTSUMV1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.MultiBand
        End If

        Export_to_Excel_Add_grd(myWorkbook, grdAPTINVH1, False, "Purchases")
        Export_to_Excel_Add_grd(myWorkbook, grdAPTCHCK1, False, "Cash Disbursements")

        optBalances.Value = "B"
        Export_to_Excel_Add_grd(myWorkbook, grdAPTINVHB, False)
        optBalances.Value = "E"
        Export_to_Excel_Add_grd(myWorkbook, grdAPTINVHB, False)

        optDistributions.Value = "0"
        Export_to_Excel_Add_grd(myWorkbook, grdAPTGEXP0, False)
        Export_to_Excel_Add_grd(myWorkbook, grdAPTVEXP0, False)
        Export_to_Excel_Add_grd(myWorkbook, grdAPTAEXP0, False)
        optDistributions.Value = "1"
        Export_to_Excel_Add_grd(myWorkbook, grdAPTGEXP0, False)
        Export_to_Excel_Add_grd(myWorkbook, grdAPTVEXP0, False)
        Export_to_Excel_Add_grd(myWorkbook, grdAPTAEXP0, False)
        optDistributions.Value = "2"
        Export_to_Excel_Add_grd(myWorkbook, grdAPTGEXP0, False)
        Export_to_Excel_Add_grd(myWorkbook, grdAPTVEXP0, False)
        Export_to_Excel_Add_grd(myWorkbook, grdAPTAEXP0, False)

        Export_to_Excel_Show(myWorkbook, Me.Text)

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default

    End Sub

    Sub Set_format_grdAPTGLAP1(Optional ByVal accounting As Boolean = False)

        Dim mask As String = ""
        If accounting Then
            mask = "#,##0.00DR;#,##0.00CR;#,##0DR"
        Else
            mask = "#,##0.00"
        End If

        With grdAPTGLAP1.DisplayLayout.Bands(0)
            .Columns("APIN_GL").Format = mask
            .Columns("APCD_GL").Format = mask
            .Columns("OTHER_GL").Format = mask
            .Columns("NET_GL").Format = mask
        End With
    End Sub
End Class