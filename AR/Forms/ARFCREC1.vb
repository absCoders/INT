Public Class ARFCREC1

    Dim RYP As String
    Dim LYP As String
    Dim ARTCUST9 As String
    Dim ARTCREC1 As String
    Dim ARTCREC2 As String
    Dim ARTCRECG As String
    Dim ARTGLARR As String

    Dim ORDR_TYPE_CODEs As String()
    Dim MISC_CHG_CODEs As String()
    Dim PYMT_SOURCEs As String()
    Dim REASON_CODEs As String()

    Dim sql_ACCT_CODE As String
    Dim ACCTS As String
    Dim POST_CODEs As New List(Of String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Load_ARTCREC1(True)

        Get_PARM("GLTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("SOTPARM1")
        Get_PARM("SPTPARM1")

        With dst
            ASCMAIN1.sql = "Select * from " & ARTCREC1
            Create_TDA(.Tables.Add, "ARTCREC1", "**", 0, False)
            .Tables("ARTCREC1").Columns.Add("CALC", GetType(System.Decimal), "BEGBAL + SLSINV + SLSCRM + SLSDRM + REBCRM + FNDCRM + CRMAMT + DRMAMT - CSHREC - DSCAMT - OFFAMT - DEDAMT - GLDAMT")
            .Tables("ARTCREC1").Columns.Add("OOB", GetType(System.Decimal), "CALC - ENDBAL")

            ASCMAIN1.sql = "Select SOTINVH1.*, NVL(ARTOPEN1.INV_BALANCE , 0) INV_BALANCE, NVL(ARTOPEN1.LAST_DATE, ARTOPEN1.INIT_DATE) TRANS_DATE" _
                & " from SOTINVH1, ARTOPEN1" _
                & " where SOTINVH1.INV_TYPE = 'C'" _
                & " and SOTINVH1.ORDR_YYYYPP_UPDATED = :PARM1" _
                & " and SOTINVH1.CUST_CODE = ARTOPEN1.CUST_CODE" _
                & " and SOTINVH1.INV_TYPE = ARTOPEN1.INV_TYPE" _
                & " and SOTINVH1.INV_NO = ARTOPEN1.INV_NUM" _
                & " UNION " _
                & "Select SOTINVH1.*, NVL(ARTOPENX.INV_BALANCE , 0) INV_BALANCE, ARTOPENX.LAST_DATE TRANS_DATE" _
                & " from SOTINVH1, ARTOPENX" _
                & " where SOTINVH1.INV_TYPE = 'C'" _
                & " and SOTINVH1.ORDR_YYYYPP_UPDATED = :PARM2" _
                & " and SOTINVH1.CUST_CODE = ARTOPENX.CUST_CODE" _
                & " and SOTINVH1.INV_TYPE = ARTOPENX.INV_TYPE" _
                & " and SOTINVH1.INV_NO = ARTOPENX.INV_NUM" _
                & " UNION " _
                & "Select SOTINVH1.*, NVL(ARTOPEN1.INV_BALANCE , 0) INV_BALANCE, NVL(ARTOPEN1.LAST_DATE, ARTOPEN1.INIT_DATE) TRANS_DATE" _
                & " from SOTINVH1, ARTOPEN1" _
                & " where SOTINVH1.INV_TYPE = 'C'" _
                & " and SOTINVH1.ORDR_YYYYPP_UPDATED <= :PARM3" _
                & " and SOTINVH1.CUST_CODE = ARTOPEN1.CUST_CODE" _
                & " and SOTINVH1.INV_TYPE = ARTOPEN1.INV_TYPE" _
                & " and SOTINVH1.INV_NO = ARTOPEN1.INV_NUM"

            Create_TDA(.Tables.Add, "SOTINVH1_C", "**", 0, False, "VVV")

            ASCMAIN1.sql = "Select *" _
                & " from SOTINVH1" _
                & " where SOTINVH1.INV_TYPE = 'I'" _
                & " and SOTINVH1.ORDR_YYYYPP_UPDATED = :PARM1"

            Create_TDA(.Tables.Add, "SOTINVH1_I", "**", 0, False, "V")

            ASCMAIN1.sql = "SELECT GLTJRNL1.JOURNAL_TYPE" & vbCrLf _
            & ", SUM (GLTDETL1.DETL_POSTING_AMT) GL_AMT" & vbCrLf _
            & Setup_sqlARTCRECG() _
            & "from GLTDETL1, GLTJRNL1" & vbCrLf
            ASCMAIN1.sql &= "" _
            & " where GLTDETL1.ACCT_CODE IN (" & ACCTS & ")" & vbCrLf _
            & "   and GLTDETL1.JOURNAL_NO = GLTJRNL1.JOURNAL_NO" & vbCrLf _
            & "   and GLTDETL1.OPS_YYYYPP = :PARM1" & vbCrLf _
            & " group by GLTJRNL1.JOURNAL_TYPE"
            Create_TDA(.Tables.Add, "ARTCRECG", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "SELECT X.*, BEGBAL + SLS - PYMTS - DISCS - ENDBAL BAL FROM (" & vbCrLf _
            & " SELECT INV_NO, SUM (BEGBAL) BEGBAL, SUM (SLS) SLS" & vbCrLf _
            & " , SUM (PYMTS) PYMTS, SUM (DISCS) DISCS, SUM (ENDBAL) ENDBAL FROM (" & vbCrLf _
            & " SELECT DETL_CTL_NO INV_NO, CREC_AMT BEGBAL, 0 SLS, 0 PYMTS, 0 DISCS, " & vbCrLf _
            & " 0 ENDBAL FROM GLTCREC3 WHERE OPS_YYYYPP = :PARM2 AND CREC_TYPE_CODE = 'AR'" & vbCrLf _
            & " AND DETL_CVX_NO = :PARM1" & vbCrLf _
            & " UNION" & vbCrLf _
            & " SELECT ORDR_INV_NO INV_NO, 0 BEGBAL, ORDR_TOTAL_AMT SLS, 0 PYMTS, 0 DISCS, 0 ENDBAL " & vbCrLf _
            & " FROM SOTINVH1" & vbCrLf _
            & " WHERE OPS_YYYYPP = :PARM3 AND NVL(CUST_BILL_TO_CUST,CUST_CODE) = :PARM1" & vbCrLf _
            & " UNION" & vbCrLf _
            & " SELECT ARTPYMT3.INV_NUM INV_NO, 0 BEGBAL, 0 SLS" & vbCrLf _
            & " , ARTPYMT3.INV_BALANCE - ARTPYMT3.INV_BALANCE_NEW PYMTS" & vbCrLf _
            & " , NVL(ARTPYMT3.INV_DISC_TAKEN,0) + NVL(ARTPYMT3.INV_WRITE_OFF,0) DISCS" & vbCrLf _
            & " , 0 ENDBAL " & vbCrLf _
            & " FROM ARTPYMT1, ARTPYMT2, ARTPYMT3" & vbCrLf _
            & " WHERE ARTPYMT1.OPS_YYYYPP = :PARM3 " & vbCrLf _
            & " AND ARTPYMT2.CUST_CODE = :PARM1" & vbCrLf _
            & " AND ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" & vbCrLf _
            & " AND ARTPYMT3.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            & " AND ARTPYMT3.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf _
            & " UNION" & vbCrLf _
            & " SELECT DETL_CTL_NO INV_NO, 0 BEGBAL, 0 SLS, 0 PYMTS, 0 DISCS" & vbCrLf _
            & " , CREC_AMT ENDBAL FROM GLTCREC3" & vbCrLf _
            & " WHERE OPS_YYYYPP = :PARM3 AND CREC_TYPE_CODE = 'AR'" & vbCrLf _
            & " AND DETL_CVX_NO = :PARM1" & vbCrLf _
            & " ) GROUP BY INV_NO) X"
            'Create_TDA(.Tables.Add, "ARTCREC2", "**", 0, False, "VVV", 0)

            ' NOTE THAT WE NEED TO CODE EACH PARAMETER INDIVIDUALLY

            ASCMAIN1.sql = "SELECT X.*, BEGBAL + SLS - PYMTS - DISCS - ENDBAL BAL FROM (" & vbCrLf _
            & " SELECT INV_NO, SUM (BEGBAL) BEGBAL, SUM (SLS) SLS" & vbCrLf _
            & " , SUM (PYMTS) PYMTS, SUM (DISCS) DISCS, SUM (ENDBAL) ENDBAL FROM (" & vbCrLf _
            & " SELECT DETL_CTL_NO INV_NO, CREC_AMT BEGBAL, 0 SLS, 0 PYMTS, 0 DISCS, " & vbCrLf _
            & " 0 ENDBAL FROM GLTCREC3 WHERE OPS_YYYYPP = :PARM1 AND CREC_TYPE_CODE = 'AR'" & vbCrLf _
            & " AND DETL_CVX_NO = :PARM2" & vbCrLf _
            & " UNION" & vbCrLf _
            & " SELECT INV_NO, 0 BEGBAL, INV_TOTAL_AMOUNT SLS, 0 PYMTS, 0 DISCS, 0 ENDBAL " & vbCrLf _
            & " FROM SOTINVH1" & vbCrLf _
            & " WHERE ORDR_YYYYPP_UPDATED = :PARM3 AND NVL(CUST_BILL_TO_CUST,CUST_CODE) = :PARM4" & vbCrLf _
            & " UNION" & vbCrLf _
            & " SELECT ARTPYMT3.INV_NUM INV_NO, 0 BEGBAL, 0 SLS" & vbCrLf _
            & " , ARTPYMT3.INV_BALANCE - ARTPYMT3.INV_BALANCE_NEW PYMTS" & vbCrLf _
            & " , NVL(ARTPYMT3.INV_DISC_TAKEN,0) + NVL(ARTPYMT3.INV_WRITE_OFF,0) DISCS" & vbCrLf _
            & " , 0 ENDBAL " & vbCrLf _
            & " FROM ARTPYMT1, ARTPYMT2, ARTPYMT3" & vbCrLf _
            & " WHERE ARTPYMT1.OPS_YYYYPP = :PARM5 " & vbCrLf _
            & " AND ARTPYMT2.CUST_CODE = :PARM6" & vbCrLf _
            & " AND ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" & vbCrLf _
            & " AND ARTPYMT3.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            & " AND ARTPYMT3.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf _
            & " UNION" & vbCrLf _
            & " SELECT DETL_CTL_NO INV_NO, 0 BEGBAL, 0 SLS, 0 PYMTS, 0 DISCS" & vbCrLf _
            & " , CREC_AMT ENDBAL FROM GLTCREC3" & vbCrLf _
            & " WHERE OPS_YYYYPP = :PARM7 AND CREC_TYPE_CODE = 'AR'" & vbCrLf _
            & " AND DETL_CVX_NO = :PARM8" & vbCrLf _
            & " ) GROUP BY INV_NO) X"
            Create_TDA(.Tables.Add, "ARTCREC2", "**", 0, False, "VVVVVVVV", 0)


            ' ARTGLARR

            Dim CUST As String = " and ARTPYMT2.CUST_CODE = '" & "NORDSTROM" & "'"
            CUST = "" ' THIS WAS NOT A GOOD IDEA

            ORDR_TYPE_CODEs = New String() {"REG", "B2C"}
            Dim SLSs As String = ""

            MISC_CHG_CODEs = New String() {"DF", "CO", "FR"}
            Dim MSCs As String = ""

            PYMT_SOURCEs = New String() {"MAN", "CC", "BOX"}
            Dim PMTs As String = ""

            REASON_CODEs = New String() {"DA", "FR", "SH"}
            'REASON_CODEs(0) = ROWs("ARTPARM1").Item("AR_PARM_REASON_CODE_DISC") & ""
            'REASON_CODEs(1) = ROWs("ARTPARM1").Item("AR_PARM_REASON_CODE_WOFF") & ""
            Dim DEDs As String = ""

            Dim sqlSLSX As String = ""
            ASCMAIN1.sql = "Select 'SLS' TYPE, REGISTER_DATE" & vbCrLf
            For Each ORDR_TYPE_CODE As String In ORDR_TYPE_CODEs
                ASCMAIN1.sql &= ", SUM (DECODE(SOTINVH1.ORDR_TYPE_CODE,'" & ORDR_TYPE_CODE & "',NVL(INV_SALES,0))) SLS_" & ORDR_TYPE_CODE & vbCrLf
                SLSs &= ", SUM (SLS_" & ORDR_TYPE_CODE & ") SLS_" & ORDR_TYPE_CODE
                sqlSLSX &= ",'" & ORDR_TYPE_CODE & "'"
            Next
            sqlSLSX = "SUM (CASE WHEN SOTINVH1.ORDR_TYPE_CODE NOT IN (" & Mid(sqlSLSX, 2) & ") THEN NVL(INV_SALES,0) ELSE 0 END)"
            ASCMAIN1.sql &= ", " & sqlSLSX & " SLSX, SUM (NVL(INV_SALES,0)) SLS" & vbCrLf
            ASCMAIN1.sql &= ", SUM (NVL(INV_STAX,0)) TAX" & vbCrLf
            ASCMAIN1.sql &= ", SUM (NVL(INV_FREIGHT,0)) FRT" & vbCrLf
            SLSs &= ", SUM(SLSX) SLSX, SUM (SLS) SLS, SUM (TAX) TAX, SUM (FRT) FRT" & vbCrLf

            Dim sqlMSCX As String = ""
            For Each MISC_CHG_CODE As String In MISC_CHG_CODEs
                ASCMAIN1.sql &= ", SUM (DECODE(SOTINVH1.MISC_CHG_CODE,'" & MISC_CHG_CODE & "',NVL(INV_MISC_CHG,0))) MSC_" & MISC_CHG_CODE & vbCrLf
                MSCs &= ", SUM (MSC_" & MISC_CHG_CODE & ") MSC_" & MISC_CHG_CODE
                sqlMSCX &= ",'" & MISC_CHG_CODE & "'"
            Next
            sqlMSCX = "SUM (CASE WHEN SOTINVH1.MISC_CHG_CODE NOT IN (" & Mid(sqlMSCX, 2) & ") THEN NVL(INV_MISC_CHG,0) ELSE 0 END)"
            ASCMAIN1.sql &= ", " & sqlMSCX & " MSCX, SUM (NVL(INV_MISC_CHG,0)) MSC" & vbCrLf
            MSCs &= ", SUM (MSCX) MSCX, SUM (MSC) MSC" & vbCrLf
            ASCMAIN1.sql &= "" _
            & ", SUM (DECODE(INV_TYPE,'I',NVL(INV_TOTAL_AMOUNT,0),0)) TOTI, SUM (DECODE(INV_TYPE,'C',NVL(INV_TOTAL_AMOUNT,0),0)) TOTC, SUM (NVL(INV_TOTAL_AMOUNT,0)) TOT" & vbCrLf
            ASCMAIN1.sql &= ", 0 PMT_" & Join(PYMT_SOURCEs, ", 0 PMT_") & ", 0 PMTX, 0 PMT" & vbCrLf
            ASCMAIN1.sql &= ", 0 DED_" & Join(REASON_CODEs, ", 0 DED_") & ", 0 DEDX, 0 DED" & vbCrLf
            ASCMAIN1.sql &= ", 0 GL_SJ, 0 GL_AR, 0 GL_XX, 0 GL, 0 NET" & vbCrLf
            ASCMAIN1.sql &= " from SOTINVH1 where ORDR_YYYYPP_UPDATED = :PARM1" & vbCrLf _
                & Replace(CUST, "ARTPYMT2", "SOTINVH1") _
            & " group by SOTINVH1.REGISTER_DATE" & vbCrLf


            Dim rowSOTTYPE1 As DataRow

            ASCMAIN1.sql &= " UNION " & vbCrLf

            Dim MSC_P As String = "MSCX"
            rowSOTTYPE1 = LookUp("SOTTYPE1", ROWs("SPTPARM1").Item("SP_PARM_PROMO_TYPE_CODE") & "")
            If rowSOTTYPE1 IsNot Nothing Then
                MSC_P = rowSOTTYPE1.Item("MISC_CHG_CODE") & ""
                If MSC_P = "" Then
                    MSC_P = rowSOTTYPE1.Item("REASON_CODE") & ""
                End If
                MSC_P = "MSC_" & MSC_P
            End If

            ASCMAIN1.sql &= "Select 'SLS' TYPE, GLTPARM2.PRD_END_DATE" & vbCrLf
            ASCMAIN1.sql &= ", 0 SLS_" & Join(ORDR_TYPE_CODEs, ", 0 SLS_") & ", 0 SLSX, 0 SLS" & vbCrLf
            ASCMAIN1.sql &= ", 0 TAX, 0 FRT" & vbCrLf
            ASCMAIN1.sql &= Replace(", 0 MSC_" & Join(MISC_CHG_CODEs, ", 0 MSC_") & ", 0 MSCX, SUM (-1 * SPTPYMT1.PYMT_REF_AMT) MSC", ", 0 " & MSC_P, ", SUM (-1 * SPTPYMT1.PYMT_REF_AMT) " & MSC_P) & vbCrLf
            ASCMAIN1.sql &= ", 0 TOTI, SUM (-1 * SPTPYMT1.PYMT_REF_AMT) TOTC, SUM (-1 * SPTPYMT1.PYMT_REF_AMT) TOT" & vbCrLf
            ASCMAIN1.sql &= ", 0 PMT_" & Join(PYMT_SOURCEs, ", 0 PMT_") & ", 0 PMTX, 0 PMT" & vbCrLf
            ASCMAIN1.sql &= ", 0 DED_" & Join(REASON_CODEs, ", 0 DED_") & ", 0 DEDX, 0 DED" & vbCrLf
            ASCMAIN1.sql &= ", 0 GL_SJ, 0 GL_AR, 0 GL_XX, 0 GL, 0 NET" & vbCrLf
            ASCMAIN1.sql &= " from SPTPYMT1,ARTPYMT1,GLTPARM2" & vbCrLf _
                & " where ARTPYMT1.PYMT_BATCH_NO = SPTPYMT1.PYMT_BATCH_NO" & vbCrLf _
                & "   and GLTPARM2.OPS_YYYYPP = SPTPYMT1.OPS_YYYYPP" & vbCrLf _
                & "   and SPTPYMT1.PYMT_TYPE = 'R'" & vbCrLf _
                & "   and SPTPYMT1.OPS_YYYYPP = :PARM1" & vbCrLf _
                & CUST _
                & " group by GLTPARM2.PRD_END_DATE" & vbCrLf


            ASCMAIN1.sql &= " UNION " & vbCrLf

            Dim MSC_A As String = "MSCX"
            rowSOTTYPE1 = LookUp("SOTTYPE1", ROWs("SPTPARM1").Item("SP_PARM_ASP_TYPE_CODE") & "")
            If rowSOTTYPE1 IsNot Nothing Then
                MSC_A = rowSOTTYPE1.Item("MISC_CHG_CODE") & ""
                If MSC_A = "" Then
                    MSC_A = rowSOTTYPE1.Item("REASON_CODE") & ""
                End If
                MSC_A = "MSC_" & MSC_A
            End If

            ASCMAIN1.sql &= "Select 'SLS' TYPE, GLTPARM2.PRD_END_DATE" & vbCrLf
            ASCMAIN1.sql &= ", 0 SLS_" & Join(ORDR_TYPE_CODEs, ", 0 SLS_") & ", 0 SLSX, 0 SLS" & vbCrLf
            ASCMAIN1.sql &= ", 0 TAX, 0 FRT" & vbCrLf
            ASCMAIN1.sql &= Replace(", 0 MSC_" & Join(MISC_CHG_CODEs, ", 0 MSC_") & ", 0 MSCX, SUM (-1 * SPTAPMT1.PYMT_REF_AMT) MSC", ", 0 " & MSC_A, ", SUM (-1 * SPTAPMT1.PYMT_REF_AMT) " & MSC_A) & vbCrLf
            ASCMAIN1.sql &= ", 0 TOTI, SUM (-1 * SPTAPMT1.PYMT_REF_AMT) TOTC, SUM (-1 * SPTAPMT1.PYMT_REF_AMT) TOT" & vbCrLf
            ASCMAIN1.sql &= ", 0 PMT_" & Join(PYMT_SOURCEs, ", 0 PMT_") & ", 0 PMTX, 0 PMT" & vbCrLf
            ASCMAIN1.sql &= ", 0 DED_" & Join(REASON_CODEs, ", 0 DED_") & ", 0 DEDX, 0 DED" & vbCrLf
            ASCMAIN1.sql &= ", 0 GL_SJ, 0 GL_AR, 0 GL_XX, 0 GL, 0 NET" & vbCrLf
            ASCMAIN1.sql &= " from SPTAPMT1,ARTPYMT1,GLTPARM2" & vbCrLf _
                & " where ARTPYMT1.PYMT_BATCH_NO = SPTAPMT1.PYMT_BATCH_NO" & vbCrLf _
                & "   and GLTPARM2.OPS_YYYYPP = SPTAPMT1.OPS_YYYYPP" & vbCrLf _
                & "   and SPTAPMT1.OPS_YYYYPP = :PARM1" & vbCrLf _
                & CUST _
                & " group by GLTPARM2.PRD_END_DATE" & vbCrLf

            ASCMAIN1.sql &= " UNION " & vbCrLf

            Dim MSC_D As String = "MSCX"
            rowSOTTYPE1 = LookUp("SOTTYPE1", ROWs("SPTPARM1").Item("SP_PARM_DEMO_TYPE_CODE") & "")
            If rowSOTTYPE1 IsNot Nothing Then
                MSC_D = rowSOTTYPE1.Item("MISC_CHG_CODE") & ""
                If MSC_D = "" Then
                    MSC_D = rowSOTTYPE1.Item("REASON_CODE") & ""
                End If
                MSC_D = "MSC_" & MSC_D
            End If

            ASCMAIN1.sql &= "Select 'SLS' TYPE, GLTPARM2.PRD_END_DATE" & vbCrLf
            ASCMAIN1.sql &= ", 0 SLS_" & Join(ORDR_TYPE_CODEs, ", 0 SLS_") & ", 0 SLSX, 0 SLS" & vbCrLf
            ASCMAIN1.sql &= ", 0 TAX, 0 FRT" & vbCrLf
            ASCMAIN1.sql &= Replace(", 0 MSC_" & Join(MISC_CHG_CODEs, ", 0 MSC_") & ", 0 MSCX, SUM (-1 * SPTDPMT1.PYMT_REF_AMT) MSC", ", 0 " & MSC_D, ", SUM (-1 * SPTDPMT1.PYMT_REF_AMT) " & MSC_D) & vbCrLf
            ASCMAIN1.sql &= ", 0 TOTI, SUM (-1 * SPTDPMT1.PYMT_REF_AMT) TOTC, SUM (-1 * SPTDPMT1.PYMT_REF_AMT) TOT" & vbCrLf
            ASCMAIN1.sql &= ", 0 PMT_" & Join(PYMT_SOURCEs, ", 0 PMT_") & ", 0 PMTX, 0 PMT" & vbCrLf
            ASCMAIN1.sql &= ", 0 DED_" & Join(REASON_CODEs, ", 0 DED_") & ", 0 DEDX, 0 DED" & vbCrLf
            ASCMAIN1.sql &= ", 0 GL_SJ, 0 GL_AR, 0 GL_XX, 0 GL, 0 NET" & vbCrLf
            ASCMAIN1.sql &= " from SPTDPMT1,ARTPYMT1,GLTPARM2" & vbCrLf _
                & " where ARTPYMT1.PYMT_BATCH_NO = SPTDPMT1.PYMT_BATCH_NO" & vbCrLf _
                & "   and GLTPARM2.OPS_YYYYPP = SPTDPMT1.OPS_YYYYPP" & vbCrLf _
                & "   and SPTDPMT1.OPS_YYYYPP = :PARM1" & vbCrLf _
                & CUST _
                & " group by GLTPARM2.PRD_END_DATE" & vbCrLf

            ASCMAIN1.sql &= " UNION " & vbCrLf

            ' ASCMAIN1.sql &= "Select 'PMT' TYPE, ARTPYMT1.REGISTER_DATE" & vbCrLf

            ASCMAIN1.sql &= "Select 'PMT' TYPE, ARTPYMT1.PYMT_BATCH_DATE" & vbCrLf
            ASCMAIN1.sql &= ", 0 SLS_" & Join(ORDR_TYPE_CODEs, ", 0 SLS_") & ", 0 SLSX, 0 SLS" & vbCrLf
            ASCMAIN1.sql &= ", 0 TAX, 0 FRT" & vbCrLf
            ASCMAIN1.sql &= ", 0 MSC_" & Join(MISC_CHG_CODEs, ", 0 MSC_") & ", 0 MSCX, 0 MSC" & vbCrLf
            ASCMAIN1.sql &= ", 0 TOTI, 0 TOTC, 0 TOT" & vbCrLf

            Dim sqlPMTX As String = ""
            For Each PYMT_SOURCE As String In PYMT_SOURCEs
                ASCMAIN1.sql &= ", SUM (DECODE(ARTPYMT1.PYMT_SOURCE,'" & PYMT_SOURCE & "',NVL(ARTPYMT2.CUST_PYMT_AMT,0))) PMT_" & PYMT_SOURCE & vbCrLf
                PMTs &= ", SUM (PMT_" & PYMT_SOURCE & ") PMT_" & PYMT_SOURCE
                sqlPMTX &= ",'" & PYMT_SOURCE & "'"
            Next
            sqlPMTX = "SUM (CASE WHEN ARTPYMT1.PYMT_SOURCE NOT IN (" & Mid(sqlPMTX, 2) & ") THEN NVL(ARTPYMT2.CUST_PYMT_AMT,0) ELSE 0 END)"
            ASCMAIN1.sql &= ", " & sqlPMTX & " MSCX, SUM (NVL(ARTPYMT2.CUST_PYMT_AMT,0)) PMT" & vbCrLf
            PMTs &= ", SUM (PMTX) PMTX, SUM (PMT) PMT" & vbCrLf

            ASCMAIN1.sql &= ", 0 DED_" & Join(REASON_CODEs, ", 0 DED_") & ", 0 DEDX, 0 DED" & vbCrLf
            ASCMAIN1.sql &= ", 0 GL_SJ, 0 GL_AR, 0 GL_XX, 0 GL, 0 NET" & vbCrLf
            ASCMAIN1.sql &= " from ARTPYMT1,ARTPYMT2" & vbCrLf _
                & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
                & " and ARTPYMT1.OPS_YYYYPP = :PARM1" & vbCrLf _
                & " and NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" & vbCrLf _
                & CUST _
                & " group by ARTPYMT1.PYMT_BATCH_DATE" & vbCrLf

            '  & " group by ARTPYMT1.REGISTER_DATE" & vbCrLf

            ASCMAIN1.sql &= " UNION " & vbCrLf

            ' ASCMAIN1.sql &= "Select 'PMT' TYPE, ARTPYMT1.REGISTER_DATE" & vbCrLf

            ASCMAIN1.sql &= "Select 'PMT' TYPE, ARTPYMT1.PYMT_BATCH_DATE" & vbCrLf
            ASCMAIN1.sql &= ", 0 SLS_" & Join(ORDR_TYPE_CODEs, ", 0 SLS_") & ", 0 SLSX, 0 SLS" & vbCrLf
            ASCMAIN1.sql &= ", 0 TAX, 0 FRT" & vbCrLf
            ASCMAIN1.sql &= ", 0 MSC_" & Join(MISC_CHG_CODEs, ", 0 MSC_") & ", 0 MSCX, 0 MSC" & vbCrLf
            ASCMAIN1.sql &= ", 0 TOTI, 0 TOTC, 0 TOT" & vbCrLf
            ASCMAIN1.sql &= ", 0 PMT_" & Join(PYMT_SOURCEs, ", 0 PMT_") & ", 0 PMTX, 0 PMT" & vbCrLf
            ASCMAIN1.sql &= ", 0 DED_" & Join(REASON_CODEs, ", 0 DED_") & ", SUM (-1 * NVL(CUST_PYMT_AMT,0)) DEDX, SUM (-1 * NVL(CUST_PYMT_AMT,0)) DED" & vbCrLf
            ASCMAIN1.sql &= ", 0 GL_SJ, 0 GL_AR, 0 GL_XX, 0 GL, 0 NET" & vbCrLf
            ASCMAIN1.sql &= " from ARTPYMT1,ARTPYMT2" & vbCrLf _
                & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
                & " and ARTPYMT1.OPS_YYYYPP = :PARM1" & vbCrLf _
                & " and NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" & vbCrLf _
                & " and ARTPYMT2.CUST_CODE is Null" & vbCrLf _
                & CUST _
                & " group by ARTPYMT1.PYMT_BATCH_DATE" & vbCrLf

            '& " group by ARTPYMT1.REGISTER_DATE" & vbCrLf

            ASCMAIN1.sql &= " UNION " & vbCrLf

            'ASCMAIN1.sql &= "Select 'D/W' TYPE, ARTPYMT1.REGISTER_DATE" & vbCrLf

            ASCMAIN1.sql &= "Select 'D/W' TYPE, ARTPYMT1.PYMT_BATCH_DATE" & vbCrLf
            ASCMAIN1.sql &= ", 0 SLS_" & Join(ORDR_TYPE_CODEs, ", 0 SLS_") & ", 0 SLSX, 0 SLS" & vbCrLf
            ASCMAIN1.sql &= ", 0 TAX, 0 FRT" & vbCrLf
            ASCMAIN1.sql &= ", 0 MSC_" & Join(MISC_CHG_CODEs, ", 0 MSC_") & ", 0 MSCX, 0 MSC" & vbCrLf
            ASCMAIN1.sql &= ", 0 TOTI, 0 TOTC, 0 TOT" & vbCrLf
            ASCMAIN1.sql &= ", 0 PMT_" & Join(PYMT_SOURCEs, ", 0 PMT_") & ", 0 PMTX, 0 PMT" & vbCrLf
            For Each REASON_CODE As String In REASON_CODEs
                If REASON_CODE = ROWs("ARTPARM1").Item("AR_PARM_REASON_CODE_DISC") & "" Then
                    ASCMAIN1.sql &= ", SUM (NVL(ARTPYMT3.INV_DISC_TAKEN,0)) DED_" & REASON_CODE & vbCrLf
                ElseIf REASON_CODE = ROWs("ARTPARM1").Item("AR_PARM_REASON_CODE_WOFF") & "" Then
                    ASCMAIN1.sql &= ", SUM (NVL(ARTPYMT3.INV_WRITE_OFF,0)) DED_" & REASON_CODE & vbCrLf
                Else
                    ASCMAIN1.sql &= ", 0 DED_" & REASON_CODE & vbCrLf
                End If
                DEDs &= ", SUM (DED_" & REASON_CODE & ") DED_" & REASON_CODE
            Next
            ASCMAIN1.sql &= ", 0 DEDX, SUM (NVL(ARTPYMT3.INV_DISC_TAKEN,0)+NVL(ARTPYMT3.INV_WRITE_OFF,0)) DED" & vbCrLf
            DEDs &= ", SUM (DEDX) DEDX, SUM (DED) DED"

            ASCMAIN1.sql &= ", 0 GL_SJ, 0 GL_AR, 0 GL_XX, 0 GL, 0 NET" & vbCrLf
            ASCMAIN1.sql &= " from ARTPYMT1,ARTPYMT3,ARTPYMT2" & vbCrLf _
            & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT3.PYMT_BATCH_NO" & vbCrLf _
            & "   and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT3.PYMT_BATCH_NO" & vbCrLf _
            & "   and ARTPYMT2.PYMT_BATCH_LNO = ARTPYMT3.PYMT_BATCH_LNO" & vbCrLf _
            & " and ARTPYMT1.OPS_YYYYPP = :PARM1" & vbCrLf _
                & CUST _
            & " group by ARTPYMT1.PYMT_BATCH_DATE" & vbCrLf

            '  & " group by ARTPYMT1.REGISTER_DATE" & vbCrLf

            ASCMAIN1.sql &= " UNION " & vbCrLf

            Dim DEDX_CODES As String = ""

            'ASCMAIN1.sql &= "Select 'DED' TYPE, ARTPYMT1.REGISTER_DATE" & vbCrLf

            ASCMAIN1.sql &= "Select 'DED' TYPE, ARTPYMT1.PYMT_BATCH_DATE" & vbCrLf
            ASCMAIN1.sql &= ", 0 SLS_" & Join(ORDR_TYPE_CODEs, ", 0 SLS_") & ", 0 SLSX, 0 SLS" & vbCrLf
            ASCMAIN1.sql &= ", 0 TAX, 0 FRT" & vbCrLf
            ASCMAIN1.sql &= ", 0 MSC_" & Join(MISC_CHG_CODEs, ", 0 MSC_") & ", 0 MSCX, 0 MSC" & vbCrLf
            ASCMAIN1.sql &= ", 0 TOTI, 0 TOTC, 0 TOT" & vbCrLf
            ASCMAIN1.sql &= ", 0 PMT_" & Join(PYMT_SOURCEs, ", 0 PMT_") & ", 0 PMTX, 0 PMT" & vbCrLf
            For Each REASON_CODE As String In REASON_CODEs
                DEDX_CODES &= ",'" & REASON_CODE & "'"
                ASCMAIN1.sql &= ", SUM (DECODE(ARTPYMT5.REASON_CODE,'" & REASON_CODE & "',NVL(ARTPYMT5.GL_DIST_AMT,0))) DED_" & REASON_CODE & vbCrLf
            Next
            ASCMAIN1.sql &= ", SUM (CASE WHEN ARTPYMT5.REASON_CODE IN (" & Mid(DEDX_CODES, 2) & ") THEN 0 ELSE NVL(ARTPYMT5.GL_DIST_AMT,0) END) DEDX, SUM (NVL(ARTPYMT5.GL_DIST_AMT,0)) DED" & vbCrLf
            ASCMAIN1.sql &= ", 0 GL_SJ, 0 GL_AR, 0 GL_XX, 0 GL, 0 NET" & vbCrLf
            ASCMAIN1.sql &= " from ARTPYMT1,ARTPYMT5,ARTPYMT2" & vbCrLf _
                & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT5.PYMT_BATCH_NO" & vbCrLf _
                & " and NVL(ARTPYMT5.CHARGEBACK_IND,'0') = '0'" & vbCrLf _
                & " and ARTPYMT1.OPS_YYYYPP = :PARM1" & vbCrLf _
                & "   and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT5.PYMT_BATCH_NO" & vbCrLf _
                & "   and ARTPYMT2.PYMT_BATCH_LNO = ARTPYMT5.PYMT_BATCH_LNO" & vbCrLf _
                & CUST _
                & " group by ARTPYMT1.PYMT_BATCH_DATE" & vbCrLf

            ' & " group by ARTPYMT1.REGISTER_DATE" & vbCrLf

            ASCMAIN1.sql &= " UNION " & vbCrLf



            ASCMAIN1.sql &= "Select 'DED' TYPE, ARTPYMT1.PYMT_BATCH_DATE" & vbCrLf
            ASCMAIN1.sql &= ", 0 SLS_" & Join(ORDR_TYPE_CODEs, ", 0 SLS_") & ", 0 SLSX, 0 SLS" & vbCrLf
            ASCMAIN1.sql &= ", 0 TAX, 0 FRT" & vbCrLf
            ASCMAIN1.sql &= ", 0 MSC_" & Join(MISC_CHG_CODEs, ", 0 MSC_") & ", 0 MSCX, 0 MSC" & vbCrLf
            ASCMAIN1.sql &= ", 0 TOTI, 0 TOTC, 0 TOT" & vbCrLf
            ASCMAIN1.sql &= ", 0 PMT_" & Join(PYMT_SOURCEs, ", 0 PMT_") & ", 0 PMTX, 0 PMT" & vbCrLf
            For Each REASON_CODE As String In REASON_CODEs
                DEDX_CODES &= ",'" & REASON_CODE & "'"
                ASCMAIN1.sql &= ", SUM (0) DED_" & REASON_CODE & vbCrLf
            Next
            ASCMAIN1.sql &= ", SUM (NVL(ARTPYMT4.GL_DIST_AMT,0)) DEDX, SUM (NVL(ARTPYMT4.GL_DIST_AMT,0)) DED" & vbCrLf
            ASCMAIN1.sql &= ", 0 GL_SJ, 0 GL_AR, 0 GL_XX, 0 GL, 0 NET" & vbCrLf
            ASCMAIN1.sql &= " from ARTPYMT1,ARTPYMT2,ARTPYMT4" & vbCrLf _
                & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT4.PYMT_BATCH_NO" & vbCrLf _
                & " and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT4.PYMT_BATCH_NO" & vbCrLf _
                & " and ARTPYMT2.PYMT_BATCH_LNO = ARTPYMT4.PYMT_BATCH_LNO" & vbCrLf _
                & " and ARTPYMT2.CUST_CODE is Not Null" & vbCrLf _
                & " and ARTPYMT1.OPS_YYYYPP = :PARM1" & vbCrLf _
                & "   and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT4.PYMT_BATCH_NO" & vbCrLf _
                & "   and ARTPYMT2.PYMT_BATCH_LNO = ARTPYMT4.PYMT_BATCH_LNO" & vbCrLf _
                & CUST _
                & " group by ARTPYMT1.PYMT_BATCH_DATE" & vbCrLf

            ' & " group by ARTPYMT1.REGISTER_DATE" & vbCrLf

            ASCMAIN1.sql &= " UNION " & vbCrLf


            ASCMAIN1.sql &= "Select 'GL' TYPE, GLTDETL1.DETL_CTL_DATE" & vbCrLf _
                & ", 0 SLS_" & Join(ORDR_TYPE_CODEs, ", 0 SLS_") & ", 0 SLSX, 0 SLS" & vbCrLf _
                & ", 0 TAX, 0 FRT" & vbCrLf _
                & ", 0 MSC_" & Join(MISC_CHG_CODEs, ", 0 MSC_") & ", 0 MSCX, 0 MSC" & vbCrLf _
                & ", 0 TOTI, 0 TOTC, 0 TOT" & vbCrLf _
                & ", 0 PMT_" & Join(PYMT_SOURCEs, ", 0 PMT_") & ", 0 PMTX, 0 PMT" & vbCrLf _
                & ", 0 DED_" & Join(REASON_CODEs, ", 0 DED_") & ", 0 DEDX, 0 DED" & vbCrLf _
                & ", SUM (CASE WHEN GLTJRNL1.JOURNAL_TYPE IN ('OPSJ','ARFC','ARFR','ARFS','AREC') " & vbCrLf _
                & "            THEN GLTDETL1.DETL_POSTING_AMT ELSE 0 END) GL_SJ" & vbCrLf _
                & ", SUM (CASE WHEN GLTJRNL1.JOURNAL_TYPE IN ('ARCR','SPCP','SPDP','SPAP') " & vbCrLf _
                & "            THEN GLTDETL1.DETL_POSTING_AMT ELSE 0 END) GL_AR" & vbCrLf _
                & ", SUM (CASE WHEN GLTJRNL1.JOURNAL_TYPE NOT IN ('OPSJ','ARFC','ARFR','ARFS','AREC','ARCR','SPCP','SPDP','SPAP') " & vbCrLf _
                & "            THEN GLTDETL1.DETL_POSTING_AMT ELSE 0 END) GL_XX" & vbCrLf _
                & ", SUM (GLTDETL1.DETL_POSTING_AMT) GL" & vbCrLf _
                & ", 0 NET" & vbCrLf _
                & " from GLTDETL1,GLTJRNL1 where GLTDETL1.JOURNAL_NO = GLTJRNL1.JOURNAL_NO" & vbCrLf _
                & " and GLTDETL1.OPS_YYYYPP = :PARM1" & vbCrLf _
                & " and (GLTDETL1.ACCT_CODE,GLTDETL1.SEG2_CODE,GLTDETL1.SEG3_CODE,GLTDETL1.SEG4_CODE) in " _
                & " (Select ACCT_CODE,NVL(SEG2_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "'), NVL(SEG3_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "'), NVL(SEG4_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "') from ARTPOST1)" & vbCrLf _
                & " group by GLTDETL1.DETL_CTL_DATE" & vbCrLf

            ASCMAIN1.sql = "Select '1' RECORD_TYPE, REGISTER_DATE" & vbCrLf _
                & SLSs & vbCrLf _
                & MSCs & vbCrLf _
                & ", SUM (TOTI) TOTI, SUM (TOTC) TOTC, SUM (TOT) TOT" & vbCrLf _
                & PMTs & vbCrLf _
                & DEDs & vbCrLf _
                & ", SUM (GL_SJ) GL_SJ, SUM (GL_AR) GL_AR, SUM (GL_XX) GL_XX, SUM (GL) GL" & vbCrLf _
                & " from (" & ASCMAIN1.sql & ") X" & vbCrLf _
                & " group by REGISTER_DATE"
            ARTGLARR = ASCMAIN1.Temp_Table(Replace(ASCMAIN1.sql, ":PARM1", "'000000'"))

            Create_TDA(.Tables.Add, "ARTGLARR", "**", 0, False, "V", 0)
            .Tables("ARTGLARR").Columns.Add("RF_GL", GetType(System.Decimal), "")
            .Tables("ARTGLARR").Columns.Add("RF_AR", GetType(System.Decimal), "")
        End With


        grdARTCREC1.DataSource = dst.Tables("ARTCREC1")
        grdARTCREC2.DataSource = dst.Tables("ARTCREC2")
        grdARTCRECG.DataSource = dst.Tables("ARTCRECG")
        grdARTGLARR.DataSource = dst.Tables("ARTGLARR")

        grdSOTINVH1_C.DataSource = dst.Tables("SOTINVH1_C")
        grdSOTINVH1_I.DataSource = dst.Tables("SOTINVH1_I")

        Format_grdARTGLARR()

        Create_Summary(grdARTCREC1, "CUST_CODE", "Count")
        Create_Summary(grdARTCREC1, New String() {"BEGBAL", "SLSINV", "SLSCRM", "SLSDRM", "REBCRM", "FNDCRM", "CRMAMT", "DRMAMT", "CSHREC", "NEWCHG", "NEWONA", "PMTAMT", "DSCAMT", "OFFAMT", "DEDAMT", "GLDAMT", "INVPMT", "CRMPMT", "DRMPMT", "ONAPMT", "RTNPMT", "CHGPMT", "ENDBAL", "CALC", "OOB"})

        Create_Summary(grdARTCREC2, New String() {"BEGBAL", "SLS", "PYMTS", "DISCS", "ENDBAL", "BAL"})

        Create_Summary(grdSOTINVH1_C, "INV_NO", "Count")
        Create_Summary(grdSOTINVH1_C, New String() {"INV_SALES", "INV_STAX", "INV_MISC_CHG", "INV_TOTAL_AMOUNT", "INV_BALANCE"})

        Create_Summary(grdSOTINVH1_I, "INV_NO", "Count")
        Create_Summary(grdSOTINVH1_I, New String() {"INV_SALES", "INV_STAX", "INV_MISC_CHG", "INV_TOTAL_AMOUNT", "SHIP_FRT_AMT_ACCRUED"}) ', "INV_CARTONS", "INV_WEIGHT"

        Create_Summary(grdARTCRECG, "JOURNAL_TYPE", "Count")

        With grdARTCRECG.DisplayLayout.Bands(0).Columns("JOURNAL_TYPE")
            .Header.Caption = "Type"
            .Width = 60
            .CellAppearance.BackColor = Drawing.Color.Beige
        End With

        Create_Summary(grdARTCRECG, "GL_AMT")

        With grdARTCRECG.DisplayLayout.Bands(0).Columns("GL_AMT")
            .Header.Caption = "Total"
            .Width = 140
            .CellAppearance.BackColor = Drawing.Color.LightGray
        End With

        For Each POST_CODE As String In POST_CODEs
            Create_Summary(grdARTCRECG, "GL_AMT_" & POST_CODE)
            With grdARTCRECG.DisplayLayout.Bands(0).Columns("GL_AMT_" & POST_CODE)
                .Header.Caption = POST_CODE
                .Width = 140
            End With
        Next


        grdARTCREC1.DisplayLayout.GroupByBox.Hidden = True
        grdARTCREC2.DisplayLayout.GroupByBox.Hidden = True
        grdARTCRECG.DisplayLayout.GroupByBox.Hidden = True


        With grdARTCREC1.DisplayLayout.Bands(0)
            .Columns("CUST_CODE").Header.Fixed = True
            .Columns("CUST_NAME").Header.Fixed = True

            .Columns("CUST_CODE").Header.Caption = "Code"
            .Columns("CUST_NAME").Header.Caption = "Customer Name"

            .Columns("CUST_CODE").CellAppearance.BackColor = Drawing.Color.Beige
            .Columns("CUST_NAME").CellAppearance.BackColor = Drawing.Color.Beige

            '.Columns("OOB").CellAppearance.BackColor = Drawing.Color.LightPink
            .Columns("ENDBAL").CellAppearance.BackColor = Drawing.Color.LightGray
            .Columns("CALC").CellAppearance.BackColor = Drawing.Color.LightGray
        End With

        With grdSOTINVH1_C.DisplayLayout.Bands(0)
            .Columns("INV_TYPE").Header.Fixed = True
            .Columns("INV_NO").Header.Fixed = True
            .Columns("CUST_CODE").Header.Fixed = True
            .Columns("CUST_STORE_NO").Header.Fixed = True
            .Columns("ORDR_CUST_PO").Header.Fixed = True
            .Columns("INV_BALANCE").Header.Fixed = True
            .Columns("TRANS_DATE").Header.Fixed = True
        End With

        With grdSOTINVH1_I.DisplayLayout.Bands(0)
            .Columns("INV_TYPE").Header.Fixed = True
            .Columns("INV_NO").Header.Fixed = True
            .Columns("CUST_CODE").Header.Fixed = True
            .Columns("CUST_STORE_NO").Header.Fixed = True
            .Columns("ORDR_CUST_PO").Header.Fixed = True
            .Columns("ORDR_DEPT").Header.Fixed = True
        End With

        ' Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        For Each toolkey As String In New String() {
            "Show Misc Charge Codes" _
            , "Show Pymt Sources" _
            , "Show Deduction Codes"}
            ' tlb_sbt = DirectCast(tlb.Tools(toolkey), UltraWinToolbars.StateButtonTool)
            ' tlb_sbt.Tag = ""
            Show_Types(toolkey, False)
            'tlb_sbt = DirectCast(tlb.Tools(toolkey), UltraWinToolbars.StateButtonTool)
            'tlb_sbt.Checked = True
            'tlb_sbt.Checked = False
        Next

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

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

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub


    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

        tabMain.Visible = ScreenMode
        Setup_tabMain()
        Setup_grdARTCREC2()
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ARTGLARR"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        'dst.Tables("ARTGLAR1").Rows.Clear()
        'dst.Tables("ARTGLAR2").Rows.Clear()
        'dst.Tables("ARTGLARC").Rows.Clear()
        'dst.Tables("ARTGLARR").Rows.Clear()
        EnforceConstraints(True)

        Absx1.txtFor("OPS_YYYYPP").Text = ASCMAIN1.CYP
    End Sub

    Sub Load_Record()
        ASCMAIN1.Progress("Now Loading Data")
        Save_Header_Fields(UltraGroupBox1)

        RYP = Absx1.txtFor("OPS_YYYYPP").Text  ' HFs("OPS_YYYYPP")
        LYP = ASCMAIN1.Period_Calc(RYP, -1)

        Fill_Records("SOTINVH1_C", New Object() {RYP, RYP, RYP})

        Sort_grdColumns(grdSOTINVH1_C, "INV_NO".ToLower)
        Clear_All_Filters(grdSOTINVH1_C)
        Show_Filter(grdSOTINVH1_C)
        grdSOTINVH1_C.DisplayLayout.Bands(0).ColumnFilters("INV_BALANCE").FilterConditions.Add(UltraWinGrid.FilterComparisionOperator.NotEquals, 0)

        Fill_Records("SOTINVH1_I", New Object() {RYP})
        Sort_grdColumns(grdSOTINVH1_I, "INV_NO")
        Clear_All_Filters(grdSOTINVH1_I)
        Show_Filter(grdSOTINVH1_I)

        Load_ARTCREC1(False)

        Fill_Records("ARTCREC1")
        Sort_grdColumns(grdARTCREC1, "CUST_CODE")

        Fill_Records("ARTCRECG", RYP)
        Sort_grdColumns(grdARTCRECG, "JOURNAL_TYPE")

        Setup_grdARTCREC1()

        Dim OPS_YYYYPP_E As String = RYP
        Fill_Records("ARTGLARR", OPS_YYYYPP_E)

        ASCMAIN1.sql = "Select SUM (NVL(TOTAL_DUE, 0) + NVL(AGE_0, 0)) from ARTSTMT1 " _
            & " where OPS_YYYYPP = '" & ASCMAIN1.Period_Calc(OPS_YYYYPP_E, -1) & "'"
        ASCMAIN1.sql = "Select SUM (NVL(TOTAL_OPEN_AMT, 0)) from ARTSTMT1 " _
            & " where OPS_YYYYPP = '" & ASCMAIN1.Period_Calc(OPS_YYYYPP_E, -1) & "'"
        Dim NET_BEG As Decimal = Val(ASCDATA1.GetDataValue)

        ASCMAIN1.sql = "Select SUM (INV_BALANCE) from ARTOPEN1"
        Dim NET_END As Decimal = Val(ASCDATA1.GetDataValue)

        Dim P_ALL As String = ""
        Dim P_TD As String = ""

        For P As Integer = 1 To 12
            Dim PZ As String = "+NVL(ACCT_ACT_P" & Format(P, "00") & ",0)"
            P_ALL &= PZ
            If P < Val(Mid(HFs("OPS_YYYYPP"), 5, 2)) Then
                P_TD &= PZ
            End If

        Next

        Dim GL_PY As Decimal = 0
        If Mid(ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP"), 1, 4) < Mid(HFs("OPS_YYYYPP"), 1, 4) Then
            ASCMAIN1.sql = "SELECT SUM (NVL(ACCT_BEG_BAL,0)" & P_ALL & ")" _
            & " from GLTACCT3" _
            & " where ACCT_YEAR >= '" & Mid(ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP"), 1, 4) & "'" _
            & "   and ACCT_YEAR < '" & Mid(HFs("OPS_YYYYPP"), 1, 4) & "'" _
            & "   and (ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE) in " _
            & " (Select ACCT_CODE,NVL(SEG2_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "'), NVL(SEG3_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "'), NVL(SEG4_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "') from ARTPOST1)" & vbCrLf
            GL_PY = Val(ASCDATA1.GetDataValue)
        End If

        ASCMAIN1.sql = "SELECT SUM (NVL(ACCT_BEG_BAL,0)" & P_TD & ")" _
        & " from GLTACCT3" _
        & " where ACCT_YEAR = '" & Mid(HFs("OPS_YYYYPP"), 1, 4) & "'" _
        & "   and (ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE) in " _
        & " (Select ACCT_CODE,NVL(SEG2_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "'), NVL(SEG3_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "'), NVL(SEG4_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "') from ARTPOST1)" & vbCrLf
        Dim GL_BEG As Decimal = Val(ASCDATA1.GetDataValue)


        Dim rowARTGLARR As DataRow = dst.Tables("ARTGLARR").NewRow
        rowARTGLARR.Item("RECORD_TYPE") = "0"
        rowARTGLARR.Item("GL") = GL_BEG + GL_PY
        rowARTGLARR.Item("RF_GL") = GL_BEG + GL_PY
        'rowARTGLARR.Item("TOT") = NET_BEG
        rowARTGLARR.Item("RF_AR") = NET_BEG
        dst.Tables("ARTGLARR").Rows.Add(rowARTGLARR)

        For Each rowARTGLARR_NODATE As DataRow In dst.Tables("ARTGLARR").Select("RECORD_TYPE = '1' and REGISTER_DATE is Null")
            rowARTGLARR_NODATE.Item("RECORD_TYPE") = "2"
        Next

        Dim RF_GL As Decimal = GL_BEG + GL_PY
        Dim RF_AR As Decimal = NET_BEG
        For Each rowARTGLARR In dst.Tables("ARTGLARR").Select("RECORD_TYPE <> '0'", "RECORD_TYPE,REGISTER_DATE")
            RF_GL += Val(rowARTGLARR.Item("GL") & "")
            rowARTGLARR.Item("RF_GL") = RF_GL
            RF_AR += Val(rowARTGLARR.Item("TOT") & "") - Val(rowARTGLARR.Item("PMT") & "") - Val(rowARTGLARR.Item("DED") & "")
            rowARTGLARR.Item("RF_AR") = RF_AR
        Next

        Set_Descriptions("SLS")

        Sort_grdColumns(grdARTGLARR, "RECORD_TYPE,REGISTER_DATE", True)
        grdARTGLARR.Text = "A/R Roll Forward for " & Absx1.txtFor("LEGEND").Text

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        ' do stuff
        Call CommitTrans("Update Complete")
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdARTCREC1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Customer Inquiry")

        Load_Popup_Menu(grdARTGLARR, "SSSSS" _
                           , "Show Sales Order Types" _
                           , "Show Misc Charge Codes" _
                           , "Show Pymt Sources" _
                           , "Show Deduction Codes" _
                           , "Show GL Columns")

        Load_Popup_Menu(grdSOTINVH1_C, "SSSPBBB", "Show Filter", "Show GroupBox", "Show Pins", "Issue Credit", "Issue Credit w/Credit Card", "Enter Trans ID")
        Load_Popup_Menu(grdSOTINVH1_I, "SSS", "Show Filter", "Show GroupBox", "Show Pins")

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


        Select Case e.SourceControl.Name
            Case "grdSOTINVH1_C"

                '"Issue Credit", "Issue Credit w/Credit Card" , "Enter Trans ID"
                If grdSOTINVH1_C.ActiveRow Is Nothing Then
                    tlb_btn = DirectCast(tlb_pop.Tools("Issue Credit"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = False

                    tlb_btn = DirectCast(tlb_pop.Tools("Issue Credit w/Credit Card"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = False

                    tlb_btn = DirectCast(tlb_pop.Tools("Enter Trans ID"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = False

                Else
                    tlb_btn = DirectCast(tlb_pop.Tools("Issue Credit"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = grdSOTINVH1_C.ActiveRow.Cells("CC_SALE_TRANS_ID").Value & String.Empty <> String.Empty _
                        AndAlso grdSOTINVH1_C.ActiveRow.Cells("CC_CRED_TRANS_ID").Value & String.Empty = String.Empty

                    tlb_btn = DirectCast(tlb_pop.Tools("Issue Credit w/Credit Card"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = grdSOTINVH1_C.ActiveRow.Cells("CC_SALE_TRANS_ID").Value & String.Empty <> String.Empty _
                        AndAlso grdSOTINVH1_C.ActiveRow.Cells("CC_CRED_TRANS_ID").Value & String.Empty = String.Empty

                    tlb_btn = DirectCast(tlb_pop.Tools("Enter Trans ID"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = grdSOTINVH1_C.ActiveRow.Cells("CC_SALE_TRANS_ID").Value & String.Empty <> String.Empty _
                        AndAlso grdSOTINVH1_C.ActiveRow.Cells("CC_CRED_TRANS_ID").Value & String.Empty = String.Empty

                End If
        End Select


        If tlb_pop.Tools.Exists("Show Sales Order Types") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Sales Order Types"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (tlb_sbt.Tag = "")
        End If
        If tlb_pop.Tools.Exists("Show Misc Charge Codes") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Misc Charge Codes"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (tlb_sbt.Tag = "")
        End If
        If tlb_pop.Tools.Exists("Show Pymt Sources") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Pymt Sources"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (tlb_sbt.Tag = "")
        End If
        If tlb_pop.Tools.Exists("Show Deduction Codes") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Deduction Codes"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (tlb_sbt.Tag = "")
        End If
        If tlb_pop.Tools.Exists("Show GL Columns") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GL Columns"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (tlb_sbt.Tag = "")
        End If



        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        End If

        Select Case e.SourceControl.Name
            'Case "grdSOTINVHX"
            '    e.Tool.ToolbarsManager.Tools("Sales Order Inquiry").SharedProps.Visible = True
        End Select
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "Show Sales Order Types"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    tlb_sbt.Tag = ""
                Else
                    tlb_sbt.Tag = "X"
                End If
                For Each COLUMN_NAME In ORDR_TYPE_CODEs
                    With grdARTGLARR.DisplayLayout.Bands(0)
                        .Columns("SLS_" & COLUMN_NAME).Hidden = (tlb_sbt.Tag = "X")
                    End With
                Next
                grdARTGLARR.DisplayLayout.Bands(0).Columns("SLSX").Hidden = (tlb_sbt.Tag = "X")

            Case "Show Misc Charge Codes"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    tlb_sbt.Tag = ""
                Else
                    tlb_sbt.Tag = "X"
                End If
                For Each COLUMN_NAME In MISC_CHG_CODEs
                    With grdARTGLARR.DisplayLayout.Bands(0)
                        .Columns("MSC_" & COLUMN_NAME).Hidden = (tlb_sbt.Tag = "X")
                    End With
                Next
                '  grdARTGLARR.DisplayLayout.Bands(0).Columns("MSC_SAM").Hidden = (tlb_sbt.Tag = "X")
                grdARTGLARR.DisplayLayout.Bands(0).Columns("MSCX").Hidden = (tlb_sbt.Tag = "X")

            Case "Show Pymt Sources"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    tlb_sbt.Tag = ""
                Else
                    tlb_sbt.Tag = "X"
                End If
                For Each COLUMN_NAME In PYMT_SOURCEs
                    With grdARTGLARR.DisplayLayout.Bands(0)
                        .Columns("PMT_" & COLUMN_NAME).Hidden = (tlb_sbt.Tag = "X")
                    End With
                Next
                grdARTGLARR.DisplayLayout.Bands(0).Columns("PMTX").Hidden = (tlb_sbt.Tag = "X")

            Case "Show Deduction Codes"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    tlb_sbt.Tag = ""
                Else
                    tlb_sbt.Tag = "X"
                End If
                For Each COLUMN_NAME In REASON_CODEs
                    With grdARTGLARR.DisplayLayout.Bands(0)
                        .Columns("DED_" & COLUMN_NAME).Hidden = (tlb_sbt.Tag = "X")
                    End With
                Next
                grdARTGLARR.DisplayLayout.Bands(0).Columns("DEDX").Hidden = (tlb_sbt.Tag = "X")

            Case "Show GL Columns"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    tlb_sbt.Tag = ""
                Else
                    tlb_sbt.Tag = "X"
                End If
                grdARTGLARR.DisplayLayout.Bands(0).Groups("GL").Hidden = (tlb_sbt.Tag = "X")
                'For Each COLUMN_NAME In New String() {"GL_SJ", "GL_AR"}
                '    With grdARTGLARR.DisplayLayout.Bands(0)
                '        .Columns(COLUMN_NAME).Hidden = (tlb_sbt.Tag = "X")
                '    End With
                'Next

            Case "Mgmt Summary"

                With grdARTGLARR.DisplayLayout.Bands(0)
                    .Groups("SLS").Hidden = True
                    .Groups("TAX").Hidden = True
                    .Groups("FRT").Hidden = True
                    .Groups("MSC").Hidden = True
                    .Groups("GL").Hidden = True
                    '.Groups("R/F").Hidden = True

                    '.Columns("REG").Hidden = True
                    ' .Columns("B2C").Hidden = True
                    ' .Columns("ECP").Hidden = True

                    ' .Columns("REG").Hidden = True
                    ' .Columns("B2C").Hidden = True
                    ' .Columns("ECP").Hidden = True
                    .Columns("RF_GL").Hidden = True

                    For Each COLUMN_NAME In REASON_CODEs
                        If COLUMN_NAME = "SALDIS" Or COLUMN_NAME = "FINCHR" Then
                            .Columns("DED_" & COLUMN_NAME).Hidden = False
                        Else
                            .Columns("DED_" & COLUMN_NAME).Hidden = True
                        End If
                    Next
                    .Columns("DEDX").Hidden = True

                    Export_to_Excel(grdARTGLARR)


                    .Groups("SLS").Hidden = False
                    .Groups("TAX").Hidden = False
                    .Groups("FRT").Hidden = False
                    .Groups("MSC").Hidden = False
                    .Groups("GL").Hidden = False
                    '.Groups("R/F").Hidden = False

                    '.Columns("REG").Hidden = False
                    ' .Columns("B2C").Hidden = False
                    ' .Columns("ECP").Hidden = False
                    .Columns("RF_GL").Hidden = False

                    Dim tlb_sbt As UltraWinToolbars.StateButtonTool = _
                        DirectCast(tlb.Tools("Show Deduction Codes"), UltraWinToolbars.StateButtonTool)
                    For Each COLUMN_NAME In REASON_CODEs
                        .Columns("DED_" & COLUMN_NAME).Hidden = (tlb_sbt.Tag = "X")
                    Next
                    .Columns("DEDX").Hidden = (tlb_sbt.Tag = "X")

                End With

                'For Each COLUMN_NAME In ORDR_TYPE_CODEs
                '    With grdARTGLARR.DisplayLayout.Bands(0)
                '        .Columns("SLS_" & COLUMN_NAME).Hidden = True
                '    End With
                'Next
                'grdARTGLARR.DisplayLayout.Bands(0).Columns("SLSX").Hidden = True

            Case "Issue Credit"
                IssueCredit(grdSOTINVH1_C.ActiveRow.Cells("INV_NO").Value & String.Empty)

            Case "Issue Credit w/Credit Card"
                IssueCreditWithCreditCard(grdSOTINVH1_C.ActiveRow.Cells("INV_NO").Value & String.Empty)

            Case "Enter Trans ID"
                ' This allows AR to process the Credit on Autorize.net and then enter the Transaction ID number in ABSolution
                ' so the Credit can be flagged as processed and match up the data from the Synergy File
                Dim CC_CRED_TRANS_ID As String = InputBox("Provide the Authorize.net Transaction Number.", "Update Auth.Net Trans ID")
                CC_CRED_TRANS_ID = CC_CRED_TRANS_ID.Trim
                If CC_CRED_TRANS_ID.Length > 0 Then
                    Dim INV_NO As String = grd.ActiveRow.Cells("INV_NO").Text
                    If MessageBox.Show("Do you want to update Credit Invoice Number (" & INV_NO & ") with the Authorize.Net Credit Transaction Number " & CC_CRED_TRANS_ID & "?", _
                                        "Update Transaction Number", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                        ASCMAIN1.sql = "Update SOTINVH1 Set CC_CRED_TRANS_ID = '" & CC_CRED_TRANS_ID & "' where INV_NO = '" & INV_NO & "'"
                        Try
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                            MessageBox.Show("Credit Updated", "Update", MessageBoxButtons.OK)
                            grd.ActiveRow.Cells("CC_CRED_TRANS_ID").Value = CC_CRED_TRANS_ID
                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
                        End Try
                    End If
                End If

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Customer Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")
        End Select
    End Sub

#End Region

    Sub Show_Types(toolkey As String, show As Boolean)
        Select Case toolkey
            Case "Show Sales Order Types"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools(toolkey), UltraWinToolbars.StateButtonTool)
                If show Then
                    tlb_sbt.Tag = ""
                Else
                    tlb_sbt.Tag = "X"
                End If
                For Each COLUMN_NAME In ORDR_TYPE_CODEs
                    With grdARTGLARR.DisplayLayout.Bands(0)
                        .Columns("SLS_" & COLUMN_NAME).Hidden = (tlb_sbt.Tag = "X")
                    End With
                Next
                grdARTGLARR.DisplayLayout.Bands(0).Columns("SLSX").Hidden = (tlb_sbt.Tag = "X")

            Case "Show Misc Charge Codes"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools(toolkey), UltraWinToolbars.StateButtonTool)
                If show Then
                    tlb_sbt.Tag = ""
                Else
                    tlb_sbt.Tag = "X"
                End If
                For Each COLUMN_NAME In MISC_CHG_CODEs
                    With grdARTGLARR.DisplayLayout.Bands(0)
                        .Columns("MSC_" & COLUMN_NAME).Hidden = (tlb_sbt.Tag = "X")
                    End With
                Next
                '  grdARTGLARR.DisplayLayout.Bands(0).Columns("MSC_SAM").Hidden = (tlb_sbt.Tag = "X")
                grdARTGLARR.DisplayLayout.Bands(0).Columns("MSCX").Hidden = (tlb_sbt.Tag = "X")

            Case "Show Pymt Sources"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools(toolkey), UltraWinToolbars.StateButtonTool)
                If show Then
                    tlb_sbt.Tag = ""
                Else
                    tlb_sbt.Tag = "X"
                End If
                For Each COLUMN_NAME In PYMT_SOURCEs
                    With grdARTGLARR.DisplayLayout.Bands(0)
                        .Columns("PMT_" & COLUMN_NAME).Hidden = (tlb_sbt.Tag = "X")
                    End With
                Next
                grdARTGLARR.DisplayLayout.Bands(0).Columns("PMTX").Hidden = (tlb_sbt.Tag = "X")

            Case "Show Deduction Codes"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools(toolkey), UltraWinToolbars.StateButtonTool)
                If show Then
                    tlb_sbt.Tag = ""
                Else
                    tlb_sbt.Tag = "X"
                End If
                For Each COLUMN_NAME In REASON_CODEs
                    With grdARTGLARR.DisplayLayout.Bands(0)
                        .Columns("DED_" & COLUMN_NAME).Hidden = (tlb_sbt.Tag = "X")
                    End With
                Next
                grdARTGLARR.DisplayLayout.Bands(0).Columns("DEDX").Hidden = (tlb_sbt.Tag = "X")

            Case "Show GL Columns"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools(toolkey), UltraWinToolbars.StateButtonTool)
                If show Then
                    tlb_sbt.Tag = ""
                Else
                    tlb_sbt.Tag = "X"
                End If
                grdARTGLARR.DisplayLayout.Bands(0).Groups("GL").Hidden = (tlb_sbt.Tag = "X")
                'For Each COLUMN_NAME In New String() {"GL_SJ", "GL_AR"}
                '    With grdARTGLARR.DisplayLayout.Bands(0)
                '        .Columns(COLUMN_NAME).Hidden = (tlb_sbt.Tag = "X")
                '    End With
                'Next

        End Select
    End Sub

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If ScreenMode Then
            Exit Sub
        End If

        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "OPS_YYYYPP"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    e.Handled = True
                    Me.ProcessTabKey(Not e.Shift)
                    Click_Command("Load", e)
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "OPS_YYYYPP"
                Click_Command("Load")
        End Select
    End Sub
#End Region

    Private Sub tabMain_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()
        UltraExplorerBar1.Groups("Activity by Customer").Visible = ScreenMode And (tabMain.SelectedTab IsNot Nothing AndAlso tabMain.SelectedTab.Key = "Activity by Customer")
        UltraExplorerBar1.Groups("Descriptions").Visible = False ' ScreenMode And (tabMain.SelectedTab IsNot Nothing AndAlso tabMain.SelectedTab.Key = "Roll Forward by Date")

    End Sub


    Sub Load_ARTCREC1(ByVal initialize As Boolean)

        ' Create or Fill ARTCUST9

        ASCMAIN1.sql = "Select OPS_YYYYPP, DETL_CVX_NO CUST_CODE, SUM (CREC_AMT) AMT from GLTCREC3 "
        If initialize Then
            ASCMAIN1.sql &= " where ROWNUM < 1"
        Else
            ASCMAIN1.sql &= "" _
            & " where CREC_TYPE_CODE = 'AR' and OPS_YYYYPP >= PERIOD_CALC('" & LYP & "',-1) " & vbCrLf _
            & "   and OPS_YYYYPP <= '" & RYP & "'"
        End If
        ASCMAIN1.sql &= " group by OPS_YYYYPP, DETL_CVX_NO"
        If initialize Then
            ARTCUST9 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST9 & " Add Primary Key (OPS_YYYYPP, CUST_CODE)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & ARTCUST9)
            ASCDATA1.ExecuteSQL("Insert into " & ARTCUST9 & " " & ASCMAIN1.sql)
        End If
        ASCMAIN1.AnalyzeTable(ARTCUST9)

        If RYP = ASCMAIN1.CYP Then
            ASCMAIN1.sql = "Delete from " & ARTCUST9 & " where OPS_YYYYPP = '" & RYP & "'"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Insert into " & ARTCUST9 _
            & " Select '" & RYP & "' OPS_YYYYPP, CUST_CODE, SUM (INV_BALANCE) AMT from ARTOPEN1 group BY CUST_CODE"
            ASCDATA1.ExecuteSQL()
        End If


        ' Create or Fill ARTCREC1

        ASCMAIN1.sql = " SELECT X.CUST_CODE, ARTCUST1.CUST_NAME, SUM (X.BEGBAL) BEGBAL" & vbCrLf _
        & " , SUM (X.SLSINV) SLSINV, SUM (X.SLSCRM) SLSCRM, SUM (X.SLSDRM) SLSDRM" & vbCrLf _
        & " , SUM (X.REBCRM) REBCRM, SUM (X.FNDCRM) FNDCRM" & vbCrLf _
        & " , SUM (X.CRMAMT) CRMAMT, SUM (X.DRMAMT) DRMAMT" & vbCrLf _
        & " , SUM (X.CSHREC) CSHREC, SUM (X.NEWCHG) NEWCHG, SUM (X.NEWONA) NEWONA" & vbCrLf _
        & " , SUM (X.PMTAMT) PMTAMT, SUM (X.DSCAMT) DSCAMT, SUM (X.OFFAMT) OFFAMT" & vbCrLf _
        & " , SUM (X.DEDAMT) DEDAMT, SUM (X.GLDAMT) GLDAMT" & vbCrLf _
        & " , SUM (X.INVPMT) INVPMT, SUM (X.CRMPMT) CRMPMT, SUM (X.DRMPMT) DRMPMT" & vbCrLf _
        & " , SUM (X.ONAPMT) ONAPMT, SUM (X.RTNPMT) RTNPMT, SUM (X.CHGPMT) CHGPMT" & vbCrLf _
        & " , SUM (X.ENDBAL) ENDBAL" & vbCrLf _
        & " FROM ARTCUST1, (" & vbCrLf _
        & " Select ARTCUST1.CUST_CODE" & vbCrLf _
        & " , Sum (AMT) BEGBAL" & vbCrLf _
        & " , 0 SLSINV, 0 SLSCRM, 0 SLSDRM, 0 REBCRM, 0 FNDCRM" & vbCrLf _
        & " , 0 CRMAMT, 0 DRMAMT" & vbCrLf _
        & " , 0 CSHREC, 0 NEWCHG, 0 NEWONA" & vbCrLf _
        & " , 0 PMTAMT, 0 DSCAMT, 0 OFFAMT" & vbCrLf _
        & " , 0 DEDAMT, 0 GLDAMT" & vbCrLf _
        & " , 0 INVPMT, 0 CRMPMT, 0 DRMPMT, 0 ONAPMT, 0 RTNPMT, 0 CHGPMT" & vbCrLf _
        & " , 0 ENDBAL" & vbCrLf _
        & "  from " & ARTCUST9 & " ARTCUST9,ARTCUST1  where ARTCUST1.CUST_CODE = ARTCUST9.CUST_CODE" & vbCrLf _
        & "  and OPS_YYYYPP = '" & LYP & "' group by ARTCUST1.CUST_CODE UNION" & vbCrLf _
        & " Select ARTCUST1.CUST_CODE" & vbCrLf _
        & " , 0 BEGBAL" & vbCrLf _
        & " , 0 SLSINV, 0 SLSCRM, 0 SLSDRM, 0 REBCRM, 0 FNDCRM" & vbCrLf _
        & " , 0 CRMAMT, 0 DRMAMT" & vbCrLf _
        & " , 0 CSHREC, 0 NEWCHG, 0 NEWONA" & vbCrLf _
        & " , 0 PMTAMT, 0 DSCAMT, 0 OFFAMT" & vbCrLf _
        & " , 0 DEDAMT, 0 GLDAMT" & vbCrLf _
        & " , 0 INVPMT, 0 CRMPMT, 0 DRMPMT, 0 ONAPMT, 0 RTNPMT, 0 CHGPMT" & vbCrLf _
        & " , Sum (AMT) ENDBAL " & vbCrLf _
        & "  from " & ARTCUST9 & " ARTCUST9,ARTCUST1  where ARTCUST1.CUST_CODE = ARTCUST9.CUST_CODE " & vbCrLf _
        & "  and OPS_YYYYPP = '" & RYP & "' group by ARTCUST1.CUST_CODE UNION" & vbCrLf _
        & " Select Decode(SOTINVH1.CUST_BILL_TO_CUST,Null,SOTINVH1.CUST_CODE,SOTINVH1.CUST_BILL_TO_CUST) CUST_CODE" & vbCrLf _
        & " , 0 BEGBAL" & vbCrLf _
        & " , Sum (Decode (SOTINVH1.INV_TYPE,'I', INV_TOTAL_AMOUNT,0)) SLSINV" & vbCrLf _
        & " , Sum (Decode (SOTINVH1.INV_TYPE,'C', INV_TOTAL_AMOUNT,0)) SLSCRM" & vbCrLf _
        & " , Sum (Decode (SOTINVH1.INV_TYPE,'D', INV_TOTAL_AMOUNT,0)) SLSDRM" & vbCrLf _
        & " , 0 REBCRM, 0 FNDCRM" & vbCrLf _
        & " , 0 CRMAMT, 0 DRMAMT" & vbCrLf _
        & " , 0 CSHREC, 0 NEWCHG, 0 NEWONA" & vbCrLf _
        & " , 0 PMTAMT, 0 DSCAMT, 0 OFFAMT" & vbCrLf _
        & " , 0 DEDAMT, 0 GLDAMT" & vbCrLf _
        & " , 0 INVPMT, 0 CRMPMT, 0 DRMPMT, 0 ONAPMT, 0 RTNPMT, 0 CHGPMT, 0 ENDBAL " & vbCrLf _
        & "  from SOTINVH1, ARTCUST1  where ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE " & vbCrLf _
        & "  and ORDR_YYYYPP_UPDATED = '" & RYP & "' and ORDR_TYPE_CODE <> 'P' " & vbCrLf _
        & "  group by Decode(SOTINVH1.CUST_BILL_TO_CUST,Null,SOTINVH1.CUST_CODE,SOTINVH1.CUST_BILL_TO_CUST) UNION" & vbCrLf _
        & " Select ARTCUST1.CUST_CODE" & vbCrLf _
        & " , 0 BEGBAL" & vbCrLf _
        & " , 0 SLSINV, 0 SLSCRM, 0 SLSDRM, 0 REBCRM, 0 FNDCRM" & vbCrLf _
        & " , 0 CRMAMT, 0 DRMAMT" & vbCrLf _
        & " , 0 CSHREC" & vbCrLf _
        & " , Sum (Decode (ABS(GL_DIST_AMT),GL_DIST_AMT,GL_DIST_AMT,0)) NEWCHG" & vbCrLf _
        & " , Sum (Decode (ABS(GL_DIST_AMT),GL_DIST_AMT,0,GL_DIST_AMT)) NEWONA" & vbCrLf _
        & " , 0 PMTAMT,  0 DSCAMT,  0 OFFAMT" & vbCrLf _
        & " , 0 DEDAMT, 0 GLDAMT" & vbCrLf _
        & " , 0 INVPMT, 0 CRMPMT, 0 DRMPMT, 0 ONAPMT, 0 RTNPMT, 0 CHGPMT" & vbCrLf _
        & " , 0 ENDBAL" & vbCrLf _
        & "  from ARTPYMT1, ARTPYMT2, ARTPYMT5, ARTCUST1 " & vbCrLf _
        & "  where ARTCUST1.CUST_CODE = ARTPYMT2.CUST_CODE and OPS_YYYYPP = '" & RYP & "'" & vbCrLf _
        & "  and NVL(CHARGEBACK_IND,'0') = '1' and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" & vbCrLf _
        & "  and ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
        & "  and ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO and ARTPYMT2.CUST_CODE is Not Null" & vbCrLf _
        & "  group by ARTCUST1.CUST_CODE UNION" & vbCrLf _
        & " Select ARTCUST1.CUST_CODE" & vbCrLf _
        & " , 0 BEGBAL" & vbCrLf _
        & " , 0 SLSINV, 0 SLSCRM, 0 SLSDRM, 0 REBCRM, 0 FNDCRM" & vbCrLf _
        & " , 0 CRMAMT, 0 DRMAMT" & vbCrLf _
        & " , 0 CSHREC, 0 NEWCHG, 0 NEWONA" & vbCrLf _
        & " , 0 PMTAMT, 0 DSCAMT, 0 OFFAMT" & vbCrLf _
        & " , Sum (GL_DIST_AMT) DEDAMT, 0 GLDAMT" & vbCrLf _
        & " , 0 INVPMT, 0 CRMPMT, 0 DRMPMT, 0 ONAPMT, 0 RTNPMT, 0 CHGPMT" & vbCrLf _
        & " , 0 ENDBAL" & vbCrLf _
        & " from ARTPYMT1, ARTPYMT2, ARTPYMT5, ARTCUST1 " & vbCrLf _
        & "  where ARTCUST1.CUST_CODE = ARTPYMT2.CUST_CODE and OPS_YYYYPP = '" & RYP & "'" & vbCrLf _
        & "  and NVL(CHARGEBACK_IND,'0') = '0'" & vbCrLf _
        & "  and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" & vbCrLf _
        & "  and ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO and ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf _
        & "  and ARTPYMT2.CUST_CODE is Not Null and (ARTPYMT2.PYMT_DELETED is Null OR ARTPYMT2.PYMT_DELETED <> '1')" & vbCrLf _
        & "  group by ARTCUST1.CUST_CODE UNION" & vbCrLf _
        & " Select ARTCUST1.CUST_CODE" & vbCrLf _
        & " , 0 BEGBAL" & vbCrLf _
        & " , 0 SLSINV, 0 SLSCRM, 0 SLSDRM, 0 REBCRM, 0 FNDCRM" & vbCrLf _
        & " , 0 CRMAMT, 0 DRMAMT" & vbCrLf _
        & " , 0 CSHREC, 0 NEWCHG, 0 NEWONA" & vbCrLf _
        & " , Sum (INV_PMT) PMTAMT, Sum (INV_DISC_TAKEN) DSCAMT, Sum (INV_WRITE_OFF) OFFAMT" & vbCrLf _
        & " , 0 DEDAMT, 0 GLDAMT" & vbCrLf _
        & " , 0 INVPMT, 0 CRMPMT, 0 DRMPMT, 0 ONAPMT, 0 RTNPMT, 0 CHGPMT" & vbCrLf _
        & " , 0 ENDBAL" & vbCrLf _
        & "  from ARTPYMT1, ARTPYMT2, ARTPYMT3, ARTCUST1  where ARTCUST1.CUST_CODE = ARTPYMT2.CUST_CODE" & vbCrLf _
        & "  and OPS_YYYYPP = '" & RYP & "' and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" & vbCrLf _
        & "  and ARTPYMT3.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO and ARTPYMT3.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf _
        & "  and (ARTPYMT2.PYMT_DELETED is Null OR ARTPYMT2.PYMT_DELETED <> '1') and ARTPYMT2.CUST_CODE is Not Null" & vbCrLf _
        & "  group by ARTCUST1.CUST_CODE UNION" & vbCrLf _
        & " Select ARTCUST1.CUST_CODE" & vbCrLf _
        & " , 0 BEGBAL" & vbCrLf _
        & " , 0 SLSINV, 0 SLSCRM, 0 SLSDRM, 0 REBCRM, 0 FNDCRM" & vbCrLf _
        & " , 0 CRMAMT, 0 DRMAMT" & vbCrLf _
        & " , 0 CSHREC, 0 NEWCHG, 0 NEWONA" & vbCrLf _
        & " , 0 PMTAMT, 0 DSCAMT, 0 OFFAMT, 0 DEDAMT, Sum (GL_DIST_AMT) GLDAMT" & vbCrLf _
        & " , 0 INVPMT, 0 CRMPMT, 0 DRMPMT, 0 ONAPMT, 0 RTNPMT, 0 CHGPMT" & vbCrLf _
        & " , 0 ENDBAL" & vbCrLf _
        & "  from ARTPYMT1, ARTPYMT2, ARTPYMT4, ARTCUST1 " & vbCrLf _
        & "  where ARTCUST1.CUST_CODE = ARTPYMT2.CUST_CODE and OPS_YYYYPP = '" & RYP & "' " & vbCrLf _
        & "  and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO and ARTPYMT4.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
        & "  and ARTPYMT4.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO and ARTPYMT2.CUST_CODE is Not Null" & vbCrLf _
        & "  and (ARTPYMT2.PYMT_DELETED is Null OR ARTPYMT2.PYMT_DELETED <> '1') group by ARTCUST1.CUST_CODE UNION" & vbCrLf _
        & " Select ARTCUST1.CUST_CODE" & vbCrLf _
        & " , 0 BEGBAL" & vbCrLf _
        & " , 0 SLSINV, 0 SLSCRM, 0 SLSDRM, 0 REBCRM, 0 FNDCRM" & vbCrLf _
        & " , 0 CRMAMT, 0 DRMAMT" & vbCrLf _
        & " , Sum (CUST_PYMT_AMT) CSHREC, 0 NEWCHG, 0 NEWONA" & vbCrLf _
        & " , 0 PMTAMT, 0 DSCAMT, 0 OFFAMT, 0 DEDAMT, 0 GLDAMT" & vbCrLf _
        & " , 0 INVPMT, 0 CRMPMT, 0 DRMPMT, 0 ONAPMT, 0 RTNPMT, 0 CHGPMT" & vbCrLf _
        & " , 0 ENDBAL" & vbCrLf _
        & "  from ARTPYMT1, ARTPYMT2, ARTCUST1 " & vbCrLf _
        & "  where ARTCUST1.CUST_CODE = ARTPYMT2.CUST_CODE and OPS_YYYYPP = '" & RYP & "'" & vbCrLf _
        & "  and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO and ARTPYMT2.CUST_CODE is Not Null" & vbCrLf _
        & "  and (ARTPYMT2.PYMT_DELETED is Null OR ARTPYMT2.PYMT_DELETED <> '1') " & vbCrLf _
        & " group by ARTCUST1.CUST_CODE union" & vbCrLf _
        & " Select ARTCUST1.CUST_CODE" & vbCrLf _
        & " , 0 BEGBAL" & vbCrLf _
        & " , 0 SLSINV, 0 SLSCRM, 0 SLSDRM, 0 REBCRM, 0 FNDCRM" & vbCrLf _
        & " , 0 CRMAMT, 0 DRMAMT" & vbCrLf _
        & " , 0 CSHREC, 0 NEWCHG, 0 NEWONA" & vbCrLf _
        & " , 0 PMTAMT, 0 DSCAMT, 0 OFFAMT, 0 DEDAMT, 0 GLDAMT" & vbCrLf _
        & " , Sum (Decode (INV_TYPE, 'I', INV_PMT, 0)) INVPMT, Sum (Decode (INV_TYPE, 'C', INV_PMT, 0)) CRMPMT" & vbCrLf _
        & " , Sum (Decode (INV_TYPE, 'D', INV_PMT, 0)) DRMPMT, Sum (Decode (INV_TYPE, 'O', INV_PMT, 0)) ONAPMT" & vbCrLf _
        & " , Sum (Decode (INV_TYPE, 'R', INV_PMT, 0)) RTNPMT, Sum (Decode (INV_TYPE, 'B', INV_PMT, 0)) CHGPMT" & vbCrLf _
        & " , 0 ENDBAL" & vbCrLf _
        & "  from ARTPYMT1, ARTPYMT2, ARTPYMT3, ARTCUST1  where ARTCUST1.CUST_CODE = ARTPYMT2.CUST_CODE" & vbCrLf _
        & "  and OPS_YYYYPP = '" & RYP & "' and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" & vbCrLf _
        & "  and ARTPYMT3.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO and ARTPYMT3.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf _
        & "  and ARTPYMT2.CUST_CODE is Not Null group by ARTCUST1.CUST_CODE union" & vbCrLf _
        & " Select ARTCUST1.CUST_CODE" & vbCrLf _
        & " , 0 BEGBAL" & vbCrLf _
        & " , 0 SLSINV, 0 SLSCRM, 0 SLSDRM, 0 REBCRM, 0 FNDCRM" & vbCrLf _
        & " , SUM (-1 * SPTPYMT1.PYMT_REF_AMT) CRMAMT, 0 DRMAMT" & vbCrLf _
        & " , 0 CSHREC, 0 NEWCHG, 0 NEWONA" & vbCrLf _
        & " , 0 PMTAMT, 0 DSCAMT, 0 OFFAMT, 0 DEDAMT, 0 GLDAMT" & vbCrLf _
        & " , 0 INVPMT, 0 CRMPMT, 0 DRMPMT, 0 ONAPMT, 0 RTNPMT, 0 CHGPMT" & vbCrLf _
        & " , 0 ENDBAL" & vbCrLf _
        & "  from SPTPYMT1,ARTCUST1 where ARTCUST1.CUST_CODE = SPTPYMT1.CUST_CODE" & vbCrLf _
        & "  and SPTPYMT1.OPS_YYYYPP = '" & RYP & "'" & vbCrLf _
        & "  and SPTPYMT1.PYMT_TYPE = 'R' group by ARTCUST1.CUST_CODE union" & vbCrLf _
        & " Select ARTCUST1.CUST_CODE" & vbCrLf _
        & " , 0 BEGBAL" & vbCrLf _
        & " , 0 SLSINV, 0 SLSCRM, 0 SLSDRM, 0 REBCRM, 0 FNDCRM" & vbCrLf _
        & " , SUM (-1 * SPTAPMT1.PYMT_REF_AMT) CRMAMT, 0 DRMAMT" & vbCrLf _
        & " , 0 CSHREC, 0 NEWCHG, 0 NEWONA" & vbCrLf _
        & " , 0 PMTAMT, 0 DSCAMT, 0 OFFAMT, 0 DEDAMT, 0 GLDAMT" & vbCrLf _
        & " , 0 INVPMT, 0 CRMPMT, 0 DRMPMT, 0 ONAPMT, 0 RTNPMT, 0 CHGPMT" & vbCrLf _
        & " , 0 ENDBAL" & vbCrLf _
        & "  from SPTAPMT1,ARTCUST1 where ARTCUST1.CUST_CODE = SPTAPMT1.CUST_CODE" & vbCrLf _
        & "  and SPTAPMT1.OPS_YYYYPP = '" & RYP & "'" & vbCrLf _
        & "  group by ARTCUST1.CUST_CODE" & vbCrLf _
        & " ) X WHERE X.CUST_CODE = ARTCUST1.CUST_CODE GROUP BY X.CUST_CODE, ARTCUST1.CUST_NAME"

        If initialize Then
            ARTCREC1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ARTCREC1 & " Add Primary Key (CUST_CODE)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & ARTCREC1)
            ASCDATA1.ExecuteSQL("Insert into " & ARTCREC1 & " " & ASCMAIN1.sql)
        End If
        ASCMAIN1.AnalyzeTable(ARTCREC1)

    End Sub

    Sub Load_ARTCREC2()
        If Not grdARTCREC1.ActiveRow.IsDataRow Then
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Retrieving Invoices")

        Dim CUST_CODE As String = grdARTCREC1.ActiveRow.Cells("CUST_CODE").Value
        'Fill_Records("ARTCREC2", New String() {CUST_CODE, LYP, RYP})
        Fill_Records("ARTCREC2", New String() {LYP, CUST_CODE, RYP, CUST_CODE, RYP, CUST_CODE, RYP, CUST_CODE})
        Sort_grdColumns(grdARTCREC2, "INV_NO")

        grdARTCREC2.Text = "Invoices for Customer " & CUST_CODE & " (Quick View - May Not include All Activity)"

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Function Setup_sqlARTCRECG() As String

        POST_CODEs.Clear()

        Dim sql_ACCT_CODE As String = ""
        ASCMAIN1.sql = "Select POST_CODE, ACCT_CODE from ARTPOST1 order by POST_CODE"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "POST_CODE")
            Dim POST_CODE As String = row.Item("POST_CODE") & ""
            Dim ACCT_CODE As String = row.Item("ACCT_CODE") & ""
            sql_ACCT_CODE &= ", SUM (DECODE(GLTDETL1.ACCT_CODE,'" & ACCT_CODE & "',GLTDETL1.DETL_POSTING_AMT,0)) GL_AMT_" & POST_CODE & vbCrLf
            ACCTS &= ",'" & ACCT_CODE & "'"
            POST_CODEs.Add(POST_CODE)
        Next

        ACCTS = Mid(ACCTS, 2)

        Return sql_ACCT_CODE
    End Function

    Private Sub chkShowInvoices_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowInvoices.CheckedChanged
        Setup_grdARTCREC2()
    End Sub

    Sub Setup_grdARTCREC2()
        SplitContainer1.Panel2Collapsed = Not chkShowInvoices.Checked

        If chkShowInvoices.Checked And ScreenMode Then
            Load_ARTCREC2()
        End If
    End Sub

    Private Sub grdARTCREC1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTCREC1.AfterRowActivate
        If chkShowInvoices.Checked Then
            Load_ARTCREC2()
        End If
    End Sub

    Private Sub grdARTCREC1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCREC1.InitializeRow
        Dim OOB As Decimal = Val(e.Row.Cells("OOB").Value & "")
        If System.Math.Abs(OOB) > 0.001 Then
            e.Row.Cells("CUST_CODE").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("CUST_NAME").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("OOB").Appearance.ForeColor = Drawing.Color.Red
        End If
    End Sub

    Sub Setup_grdARTCREC1()
        Dim dvw As DataView = DirectCast(grdARTCREC1.DataSource, DataTable).DefaultView
        If chkUseThreshold.Checked Then
            dvw.RowFilter = " (OOB > 0 and OOB > " & CStr(numOOBAL.Value) & ") OR (OOB < 0 and OOB < " & CStr(numOOBAL.Value) & ")"
        Else
            dvw.RowFilter = ""
        End If
    End Sub

    Private Sub chkUseThreshold_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseThreshold.CheckedChanged
        Setup_grdARTCREC1()
    End Sub

    Private Sub numOOBAL_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numOOBAL.ValueChanged
        Setup_grdARTCREC1()
    End Sub


    Sub Format_grdARTGLARR()
        With grdARTGLARR.DisplayLayout.Bands(0)
            With .Groups.Add("DATE", "")
                .Header.Fixed = True
            End With
            With .Groups.Add("SLS", "Sales")
                '.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                '.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            End With
            With .Groups.Add("TAX", "Tax")

            End With
            With .Groups.Add("FRT", "Freight")

            End With
            With .Groups.Add("MSC", "Misc")

            End With
            With .Groups.Add("TOT", "Net A/R")

            End With
            With .Groups.Add("PMT", "Pymts")

            End With
            With .Groups.Add("DED", "Deducts")

            End With
            With .Groups.Add("GL", "GL J/E's")

            End With
            With .Groups.Add("R/F", "Roll Fwd")

            End With

            ' Or GC.Key Like "ECP*" Or GC.Key Like "REG*" Or GC.Key Like "B2C*"

            .Columns("REGISTER_DATE").CellAppearance.BackColor = Drawing.Color.LightYellow
            For Each GC As UltraWinGrid.UltraGridColumn In .Columns
                If GC.Key = "REGISTER_DATE" Then
                    GC.Group = .Groups("DATE")
                Else
                    Dim handled As Boolean = False
                    If GC.Key Like "SLS*" Then
                        GC.CellAppearance.BackColor = Drawing.Color.LightGreen
                        GC.Group = .Groups("SLS")
                    ElseIf GC.Key Like "MSC*" Then
                        GC.CellAppearance.BackColor = Drawing.Color.Lavender
                        GC.Group = .Groups("MSC")
                    ElseIf GC.Key Like "TOT*" Then
                        GC.CellAppearance.BackColor = Drawing.Color.AntiqueWhite
                        GC.Group = .Groups("TOT")
                    ElseIf GC.Key Like "PMT*" Then
                        GC.CellAppearance.BackColor = Drawing.Color.LightBlue
                        GC.Group = .Groups("PMT")
                    ElseIf GC.Key Like "DED*" Then
                        GC.CellAppearance.BackColor = Drawing.Color.LightPink
                        GC.Group = .Groups("DED")
                    ElseIf GC.Key Like "GL*" Then
                        GC.CellAppearance.BackColor = Drawing.Color.Yellow
                        GC.Group = .Groups("GL")
                    ElseIf GC.Key = "RF_GL" Or GC.Key = "RF_AR" Then
                        GC.CellAppearance.BackColor = Drawing.Color.Beige
                        GC.Group = .Groups("R/F")
                    Else
                        handled = True
                        If GC.Key = "TAX" Or GC.Key = "FRT" Then
                            GC.Group = .Groups(GC.Key)
                        End If
                    End If

                    If Not handled Then
                        If Len(GC.Key) = 3 Then
                            Select Case GC.Key
                                Case "SLS"
                                    GC.Header.Caption = "Total"
                                Case "MSC"
                                    GC.Header.Caption = "Total"
                                Case "PMT"
                                    GC.Header.Caption = "Total"
                                Case "DED"
                                    GC.Header.Caption = "Total"
                                Case "B2C"
                                    GC.Header.Caption = "B2C"
                                Case "NRSP"
                                    GC.Header.Caption = "NRSP"
                                Case "TOT"
                                    GC.Header.Caption = "Total"
                            End Select
                        Else
                            If Len(GC.Key) = 4 And GC.Key Like "*X" Then
                                GC.Header.Caption = "Other"
                            Else
                                Select Case GC.Key
                                    Case "GL_SJ"
                                        GC.Header.Caption = "Sales"
                                    Case "GL_AR"
                                        GC.Header.Caption = "AR"
                                    Case "GL_XX"
                                        GC.Header.Caption = "Other"
                                    Case "GL"
                                        GC.Header.Caption = "Total"
                                    Case "RF_AR"
                                        GC.Header.Caption = "AR"
                                    Case "RF_GL"
                                        GC.Header.Caption = "GL"
                                    Case "TOTI"
                                        GC.Header.Caption = "Invs"
                                    Case "TOTC"
                                        GC.Header.Caption = "Crds"
                                    Case Else
                                        GC.Header.Caption = Mid(GC.Key, 5)
                                End Select
                            End If
                        End If
                    End If

                    GC.Width = 80

                    If GC.Key = "REG" Or GC.Key = "B2C" Or GC.Key = "TOTI" Or GC.Key = "TOTC" Or GC.Key = "NRSF" Or GC.Key = "TOT" Or GC.Key = "GL" Or GC.Key = "RF_GL" Or GC.Key = "RF_AR" Then
                        GC.Width = 100
                    ElseIf GC.Key = "FRT" Or GC.Key = "TAX" Or GC.Key Like "MSC*" Then
                        GC.Width = 70
                    End If
                    GC.Format = "#,##0"
                    If GC.Key <> "RF_GL" And GC.Key <> "RF_AR" Then
                        Create_Summary(grdARTGLARR, GC.Key, , , "###,##0")
                    End If
                End If
            Next
        End With

        grdARTGLARR.DisplayLayout.UseFixedHeaders = True
        With grdARTGLARR.DisplayLayout.Bands(0)
            .Columns("REGISTER_DATE").Header.Fixed = True
        End With
    End Sub


    Private Sub grdARTGLARR_AfterCellActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTGLARR.AfterCellActivate
        Dim COLUMN_NAME = grdARTGLARR.ActiveCell.Column.Key
        If COLUMN_NAME Like "SLS*" Then
            Set_Descriptions("SLS")
        ElseIf COLUMN_NAME Like "MSC*" Then
            Set_Descriptions("MSC")
        ElseIf COLUMN_NAME Like "PMT*" Then
            Set_Descriptions("PMT")
        ElseIf COLUMN_NAME Like "DED*" Then
            Set_Descriptions("DED")
        End If
    End Sub

    Private Sub grdARTGLARR_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTGLARR.InitializeRow
        Dim RF_GL As Decimal = Val(e.Row.Cells("RF_GL").Value & "")
        Dim RF_AR As Decimal = Val(e.Row.Cells("RF_AR").Value & "")
        If RF_GL <> RF_AR Then
            e.Row.Cells("RF_GL").Appearance.ForeColor = Drawing.Color.Red
        End If

        If e.Row.Cells("RECORD_TYPE").Value & "" = "1" Then
            Dim SLS As Decimal = Val(e.Row.Cells("SLS").Value & "")
            Dim FRT As Decimal = Val(e.Row.Cells("FRT").Value & "")
            Dim TAX As Decimal = Val(e.Row.Cells("TAX").Value & "")
            Dim MSC As Decimal = Val(e.Row.Cells("MSC").Value & "")
            Dim TOT As Decimal = Val(e.Row.Cells("TOT").Value & "")
            If SLS + FRT + TAX + MSC <> TOT Then
                e.Row.Cells("TOT").Appearance.ForeColor = Drawing.Color.Red
            End If
            '  e.Row.Cells("ECP").Appearance.ForeColor = Drawing.Color.Blue
        End If
    End Sub

    Sub Set_Descriptions(ByVal code_type As String)
        Select Case code_type
            Case "SLS"
                grdDescriptions.Text = "Order Types"
                ASCMAIN1.sql = "Select ORDR_TYPE_CODE CODE_VALUE, ORDR_TYPE_DESC DESC_VALUE from SOTTYPE1"
            Case "MSC"
                grdDescriptions.Text = "Misc Charges"
                ASCMAIN1.sql = "Select MISC_CHG_CODE CODE_VALUE, MISC_CHG_DESC DESC_VALUE from SOTMISC1"
            Case "PMT"
                grdDescriptions.Text = "Pymt Sources"
                ASCMAIN1.sql = "Select T_CODE CODE_VALUE, T_DESC DESC_VALUE from ASTCODE1 where TABLE_NAME = 'ARTPYMT1' and COLUMN_NAME = 'PYMT_SOURCE'"
            Case "DED"
                grdDescriptions.Text = "Deductions"
                ASCMAIN1.sql = "Select REASON_CODE CODE_VALUE, REASON_DESC DESC_VALUE from ARTREAS1"
        End Select

        grdDescriptions.DataSource = ASCDATA1.GetDataTable
        With grdDescriptions.DisplayLayout.Bands(0)
            .Columns("CODE_VALUE").Header.Caption = "Code"
            .Columns("CODE_VALUE").Width = 60
            .Columns("DESC_VALUE").Header.Caption = "Description"
            .Columns("DESC_VALUE").Width = 120
        End With
        grdDescriptions.DisplayLayout.Override.RowSelectors = DefaultableBoolean.False
        grdDescriptions.DisplayLayout.Override.RowSelectorHeaderStyle = UltraWinGrid.RowSelectorHeaderStyle.None

        Sort_grdColumns(grdDescriptions, "CODE_VALUE")
    End Sub

    Private Sub IssueCredit(ByVal INV_NO As String)

        Try
            ASCMAIN1.Progress("Processing CC Credit", "")
            Dim errorMessage As String = String.Empty
            If Not SOCMAIN1.IssueCredit(INV_NO, errorMessage) Then
                MessageBox.Show("Error Processing Credit Card Refund: " & errorMessage, "CC Credit", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                'grdSOTINVH1_C.ActiveRow.Cells("INV_NO").Value & String.Empty
                dst.Tables("SOTINVH1_C").Select("INV_NO = '" & INV_NO & "'")(0).Item("CC_CRED_TRANS_ID") = errorMessage
                MessageBox.Show("Credit Card Successfully Refunded. Transaction ID: " & errorMessage, "CC Credit", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
            ASCMAIN1.Progress("", "")

        Catch ex As Exception
            MessageBox.Show("Error Processing Credit Card Refund: " & ex.Message, "Issue Credit", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub IssueCreditWithCreditCard(ByVal INV_NO As String)

        MsgBox("Method not Supported (Need Credit Card Integrator")
        Exit Sub

        ' the code block below was remmed out by WJZ on 06/17 to start implementation of ACH for INT using ePayment Integrator

        'Try

        '    If 1 = 1 Then
        '        MessageBox.Show("Not fpor production yet.")
        '        Exit Sub
        '    End If

        '    Dim CCPA_NO As String = String.Empty

        '    If dst.Tables("SOTINVH1_C").Select("INV_NO = '" & INV_NO & "'").Length = 0 Then
        '        MessageBox.Show("Cannot Locate the supplied Credit Invoice No: " & INV_NO, "Issue Credit", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        Exit Sub
        '    End If

        '    Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1_C").Select("INV_NO = '" & INV_NO & "'")(0)
        '    If rowSOTINVH1.Item("CC_CRED_TRANS_ID") & String.Empty <> String.Empty Then
        '        MessageBox.Show("The credit was already issued a refund using a credit card.", "Issue Credit", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        Exit Sub
        '    End If


        '    Dim InvoiceBalance As Decimal = Math.Abs(Val(rowSOTINVH1.Item("INV_BALANCE") & String.Empty))
        '    If InvoiceBalance = 0 Then
        '        MessageBox.Show("Cannot process Credit for $0.00 balance", "Issue Credit", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        Exit Sub
        '    End If

        '    Dim CreditAmount As Decimal = Math.Abs(Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & String.Empty))
        '    If CreditAmount = 0 Then
        '        MessageBox.Show("Cannot process Credit for $0.00 credit", "Issue Credit", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        Exit Sub
        '    End If

        '    Dim CUST_CODE As String = rowSOTINVH1.Item("CUST_CODE")

        '    Dim frmCCProcessor As New TAC.TAFCARDF(Me)

        '    frmCCProcessor.test_mode = ROWs("SOTPARM1").Item("SO_PARM_CC_TEST_MODE") & String.Empty = "1"
        '    frmCCProcessor.CUST_CODE = CUST_CODE
        '    frmCCProcessor.CCPA_REASON = "M"
        '    frmCCProcessor.TRAN_TYPE = "C"

        '    With frmCCProcessor.rowARTCCPA1
        '        .Item("CUST_CODE") = CUST_CODE
        '        .Item("CCPA_AMT") = CreditAmount
        '        .Item("CCPA_NOTE") = "Credit Card Credit"
        '    End With

        '    Try
        '        frmCCProcessor.ShowDialog()
        '        CCPA_NO = frmCCProcessor.CCPA_NO

        '        If CCPA_NO.Length > 0 Then
        '            ASCMAIN1.sql = "Update SOTINVH1 set CCPA_NO = '" & CCPA_NO & "', CC_CRED_TRANS_ID = '" & frmCCProcessor.MerchantTransID & "' WHERE INV_TYPE = 'C' AND INV_NO = '" & INV_NO & "'"
        '            ASCDATA1.ExecuteSQL()
        '            rowSOTINVH1.Item("CCPA_NO") = CCPA_NO
        '            rowSOTINVH1.Item("CC_CRED_TRANS_ID") = frmCCProcessor.MerchantTransID
        '        Else
        '            MessageBox.Show("Could not process Credit for the following reason: " & frmCCProcessor.responseErrorMessage, "Issue Credit", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        End If

        '    Catch ex As Exception
        '        MessageBox.Show(ex.Message, "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    End Try

        '    Try

        '    Catch ex As Exception
        '        MessageBox.Show("Error Processing Credit: " & ex.Message, "Issue Credit", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    End Try

        '    If CCPA_NO.Length > 0 Then
        '        ASCMAIN1.sql = "Update SOTINVH1 set CCPA_NO = '" & CCPA_NO & "', CC_CRED_TRANS_ID = '" & frmCCProcessor.MerchantTransID & "' WHERE INV_TYPE = 'C' AND INV_NO = '" & INV_NO & "'"
        '        ASCDATA1.ExecuteSQL()
        '        rowSOTINVH1.Item("CCPA_NO") = CCPA_NO
        '        rowSOTINVH1.Item("CC_CRED_TRANS_ID") = frmCCProcessor.MerchantTransID
        '    Else
        '        MessageBox.Show("Could not process Credit for the following reason: " & frmCCProcessor.responseErrorMessage, "Issue Credit", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    End If

        '    frmCCProcessor.Dispose()

        'Catch ex As Exception
        '    MessageBox.Show("Error Processing Credit: " & ex.Message, "Issue Credit", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End Try

    End Sub

    Private Sub grdSOTINVH1_C_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTINVH1_C.InitializeRow

        If (e.Row.Cells("INV_BALANCE").Value & String.Empty) <> 0 Then
            If e.Row.Cells("CC_CRED_TRANS_ID").Value & String.Empty <> String.Empty OrElse e.Row.Cells("CCPA_NO").Value & String.Empty <> String.Empty Then
                e.Row.Appearance.BackColor = Drawing.Color.IndianRed
            ElseIf e.Row.Cells("CC_SALE_TRANS_ID").Value & String.Empty <> String.Empty Then
                e.Row.Appearance.BackColor = Drawing.Color.LightGreen
            End If
        End If
    End Sub

End Class