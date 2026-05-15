Public Class ARFGLAR1

    Dim ARTGLAR2 As String
    Dim sqlARTGLAR2 As String
    Dim ARTGLARC As String
    Dim ARTGLARR As String

    Dim ARTGLARS As String
    Dim sqlARTGLARS As String

    Dim ORDR_TYPE_CODEs As String()
    Dim MISC_CHG_CODEs As String()
    Dim PYMT_SOURCEs As String()
    Dim REASON_CODEs As String()
    Dim PPE As Date

    Dim OPS_YYYYPP_B As String = String.Empty
    Dim OPS_YYYYPP_E As String = String.Empty
    Dim viewARTPYMT1X As DataView
    Const periodDays As Int16 = 35

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Call Get_PARM("GLTPARM1")
        Call Get_PARM("ARTPARM1")

        With dst

            Dim sqlx As String = ""

            ' ARTGLARS

            sqlx = "" _
            & "Select DIVISION_CODE, 0 WEEK, 0 LINE, INV_DATE" & vbCrLf _
            & ", SUM (INV_SLS) INV_SLS, SUM (INV_FRT) INV_FRT, SUM (INV_TAX) INV_TAX" & vbCrLf _
            & ", SUM (MISC) MISC, SUM (RSF) RSF, SUM (NET_AR) NET_AR" & vbCrLf _
            & ", SUM (INVS) INVS, SUM (CRM_SLS) CRM_SLS, SUM (CRM_FRT) CRM_FRT, SUM (CRM_TAX) CRM_TAX" & vbCrLf _
            & ", SUM (WEB_SLS) WEB_SLS, SUM (WEB_FRT) WEB_FRT, SUM (WEB_TAX) WEB_TAX, SUM (WEB_FEE) WEB_FEE, SUM (WEB_NET) WEB_NET" & vbCrLf _
            & ", SUM (DEL_SLS) DEL_SLS" & vbCrLf _
            & ", SUM (DEL_EDG) DEL_EDG" & vbCrLf _
            & ", SUM (DEL_ARC) DEL_ARC" & vbCrLf _
            & ", SUM (DEL_LTR) DEL_LTR" & vbCrLf _
            & ", SUM (DEL_DSC) DEL_DSC" & vbCrLf _
            & ", SUM (DEL_NET) DEL_NET" & vbCrLf _
            & " FROM (" & vbCrLf _
            & " Select DIVISION_CODE, INV_DATE" & vbCrLf _
            & ", SUM (DECODE(INV_TYPE,'I',NVL(INV_SALES,0) - NVL(INV_FREE_AMT,0),0)) INV_SLS" & vbCrLf _
            & ", SUM (DECODE(INV_TYPE,'I',INV_FREIGHT,0)) INV_FRT" & vbCrLf _
            & ", SUM (DECODE(INV_TYPE,'I',INV_STAX,0)) INV_TAX" & vbCrLf _
            & ", SUM (DECODE(MISC_CHG_CODE,'RSF',0,INV_MISC_CHG)) MISC" & vbCrLf _
            & ", SUM (DECODE(MISC_CHG_CODE,'RSF',INV_MISC_CHG,0)) RSF" & vbCrLf _
            & ", SUM (INV_TOTAL_AMOUNT) NET_AR" & vbCrLf _
            & ", SUM (DECODE(INV_TYPE || NVL(INV_RESHIP,'0'),'I0',1,0)) INVS" & vbCrLf _
            & ", SUM (DECODE(INV_TYPE,'C',INV_SALES,0)) CRM_SLS" & vbCrLf _
            & ", SUM (DECODE(INV_TYPE,'C',INV_FREIGHT,0)) CRM_FRT" & vbCrLf _
            & ", SUM (DECODE(INV_TYPE,'C',INV_STAX,0)) CRM_TAX" & vbCrLf _
            & ", SUM (CASE WHEN ORDR_TYPE_CODE IN ('B2B','B2C') THEN NVL(INV_SALES,0) ELSE 0 END) WEB_SLS" & vbCrLf _
            & ", SUM (CASE WHEN ORDR_TYPE_CODE IN ('B2B','B2C') THEN NVL(INV_FREIGHT,0) ELSE 0 END) WEB_FRT" & vbCrLf _
            & ", SUM (CASE WHEN ORDR_TYPE_CODE IN ('B2B','B2C') THEN NVL(INV_STAX,0) ELSE 0 END) WEB_TAX" & vbCrLf _
            & ", SUM (CASE WHEN ORDR_TYPE_CODE IN ('B2B','B2C') THEN NVL(INV_MISC_CHG,0) ELSE 0 END) WEB_FEE" & vbCrLf _
            & ", SUM (CASE WHEN ORDR_TYPE_CODE IN ('B2B','B2C') THEN NVL(INV_TOTAL_AMOUNT,0) ELSE 0 END) WEB_NET" & vbCrLf _
            & ", 0 DEL_SLS, 0 DEL_EDG, 0 DEL_ARC, 0 DEL_LTR, 0 DEL_DSC, 0 DEL_NET" & vbCrLf _
            & "         from SOTINVH1 " & vbCrLf _
            & " where (ORDR_YYYYPP_UPDATED Between '000000' AND '000001') AND DIVISION_CODE <> 'DEL'" & vbCrLf _
            & " group by DIVISION_CODE, INV_DATE" & vbCrLf _
            & " union " & vbCrLf _
            & " Select SOTINVH1.DIVISION_CODE, SOTINVH1.INV_DATE" & vbCrLf _
            & ", SUM (DECODE(SOTINVH1.INV_TYPE,'I',NVL(SOTINVH1.INV_SALES,0) - NVL(SOTINVH1.INV_FREE_AMT,0),0)) INV_SLS" & vbCrLf _
            & ", SUM (DECODE(SOTINVH1.INV_TYPE,'I',SOTINVH1.INV_FREIGHT,0)) INV_FRT" & vbCrLf _
            & ", SUM (DECODE(SOTINVH1.INV_TYPE,'I',SOTINVH1.INV_STAX,0)) INV_TAX" & vbCrLf _
            & ", SUM (DECODE(SOTINVH1.MISC_CHG_CODE,'RSF',0,SOTINVH1.INV_MISC_CHG)) MISC" & vbCrLf _
            & ", SUM (DECODE(SOTINVH1.MISC_CHG_CODE,'RSF',SOTINVH1.INV_MISC_CHG,0)) RSF" & vbCrLf _
            & ", SUM (SOTINVH1.INV_TOTAL_AMOUNT) NET_AR" & vbCrLf _
            & ", SUM (DECODE(SOTINVH1.INV_TYPE || NVL(SOTINVH1.INV_RESHIP,'0'),'I0',1,0)) INVS" & vbCrLf _
            & ", SUM (DECODE(SOTINVH1.INV_TYPE,'C',SOTINVH1.INV_SALES,0)) CRM_SLS" & vbCrLf _
            & ", SUM (DECODE(SOTINVH1.INV_TYPE,'C',SOTINVH1.INV_FREIGHT,0)) CRM_FRT" & vbCrLf _
            & ", SUM (DECODE(SOTINVH1.INV_TYPE,'C',SOTINVH1.INV_STAX,0)) CRM_TAX" & vbCrLf _
            & ", SUM (CASE WHEN NVL(ORDR_SOURCE, '*') IN ('W') THEN NVL(SOTINVH1.INV_SALES,0) ELSE 0 END) WEB_SLS" & vbCrLf _
            & ", SUM (CASE WHEN NVL(ORDR_SOURCE, '*') IN ('W') THEN NVL(SOTINVH1.INV_FREIGHT,0) ELSE 0 END) WEB_FRT" & vbCrLf _
            & ", SUM (CASE WHEN NVL(ORDR_SOURCE, '*') IN ('W') THEN NVL(SOTINVH1.INV_STAX,0) ELSE 0 END) WEB_TAX" & vbCrLf _
            & ", SUM (CASE WHEN NVL(ORDR_SOURCE, '*') IN ('W') THEN NVL(SOTINVH1.INV_MISC_CHG,0) ELSE 0 END) WEB_FEE" & vbCrLf _
            & ", SUM (CASE WHEN NVL(ORDR_SOURCE, '*') IN ('W') THEN NVL(SOTINVH1.INV_TOTAL_AMOUNT,0) ELSE 0 END) WEB_NET" & vbCrLf _
            & ", 0 DEL_SLS, 0 DEL_EDG, 0 DEL_ARC, 0 DEL_LTR, 0 DEL_DSC, 0 DEL_NET" & vbCrLf _
            & "         from SOTINVH1, DETJOBM1 " & vbCrLf _
            & " where SOTINVH1.ORDR_YYYYPP_UPDATED Between '000000' AND '000001' " & vbCrLf _
            & " AND SOTINVH1.DIVISION_CODE = 'DEL'" & vbCrLf _
            & " AND SOTINVH1.INV_NO = DETJOBM1.INV_NO (+)" & vbCrLf _
            & " group by SOTINVH1.DIVISION_CODE, SOTINVH1.INV_DATE" & vbCrLf _
            & " union " & vbCrLf _
            & " SELECT 'DEL' DIVISION_CODE, INV_DATE" & vbCrLf _
            & ", 0 INV_SLS, 0 INV_FRT, 0 INV_TAX, 0 MISC, 0 RSF, 0 NET_AR" & vbCrLf _
            & ", 0 INVS, 0 CRM_SLS, 0 CRM_FRT, 0 CRM_TAX" & vbCrLf _
            & ", 0 WEB_SLS, 0 WEB_FRT, 0 WEB_TAX, 0 WEB_FEE, 0 WEB_NET" & vbCrLf _
            & ", SUM (DEL_SLS) DEL_SLS" & vbCrLf _
            & ", SUM (DEL_EDG) DEL_EDG" & vbCrLf _
            & ", SUM (DEL_ARC) DEL_ARC" & vbCrLf _
            & ", SUM (DEL_LTR) DEL_LTR" & vbCrLf _
            & ", SUM (DEL_DSC) DEL_DSC" & vbCrLf _
            & ", SUM (DEL_NET) DEL_NET" & vbCrLf _
            & " FROM (" & vbCrLf _
            & " SELECT DETJOBM1.INV_DATE,DETJOBM1.INV_NO" & vbCrLf _
            & ", SUM (CASE WHEN DETJOBM2.JOB_CHARGE_TYPE IN ('L','B') THEN DETJOBM2.JOB_PRICE * DETJOBM2.JOB_QTY ELSE 0 END) DEL_SLS" & vbCrLf _
            & ", SUM (CASE WHEN DETJOBM2.JOB_CHARGE_TYPE IN ('E') THEN DETJOBM2.JOB_PRICE * DETJOBM2.JOB_QTY ELSE 0 END) DEL_EDG" & vbCrLf _
            & ", SUM (CASE WHEN DETJOBM2.JOB_CHARGE_TYPE IN ('C') THEN DETJOBM2.JOB_PRICE * DETJOBM2.JOB_QTY ELSE 0 END) DEL_ARC" & vbCrLf _
            & ", SUM (CASE WHEN DETJOBM2.JOB_CHARGE_TYPE NOT IN ('L','B','E','C') THEN DETJOBM2.JOB_PRICE * DETJOBM2.JOB_QTY ELSE 0 END) DEL_LTR" & vbCrLf _
            & ", SUM ((DECODE(NVL(DETJOBM2.LIST_PRICE,0),0,NVL(DETJOBM2.JOB_PRICE,0),NVL(DETJOBM2.LIST_PRICE,0)) - NVL(DETJOBM2.JOB_PRICE,0)) * DETJOBM2.JOB_QTY) DEL_DSC" & vbCrLf _
            & ", DETJOBM1.INV_TOTAL_AMOUNT DEL_NET" & vbCrLf _
            & " FROM DETJOBM1,DETJOBM2 WHERE DETJOBM1.OPS_YYYYPP Between '000000' AND '000001'" & vbCrLf _
            & " AND DETJOBM2.JOB_NO = DETJOBM1.JOB_NO" & vbCrLf _
            & " GROUP BY DETJOBM1.INV_DATE, DETJOBM1.INV_NO,DETJOBM1.INV_TOTAL_AMOUNT" & vbCrLf _
            & " UNION" & vbCrLf _
            & " SELECT DETJOBC1.INV_DATE,DETJOBC1.INV_NO" & vbCrLf _
            & ", SUM (CASE WHEN DETJOBM2.JOB_CHARGE_TYPE IN ('L','B') THEN -1 * DETJOBC2.CREDIT_AMT ELSE 0 END) DEL_SLS" & vbCrLf _
            & ", SUM (CASE WHEN DETJOBM2.JOB_CHARGE_TYPE IN ('E') THEN -1 * DETJOBC2.CREDIT_AMT ELSE 0 END) DEL_EDG" & vbCrLf _
            & ", SUM (CASE WHEN DETJOBM2.JOB_CHARGE_TYPE IN ('C') THEN -1 * DETJOBC2.CREDIT_AMT ELSE 0 END) DEL_ARC" & vbCrLf _
            & ", SUM (CASE WHEN DETJOBM2.JOB_CHARGE_TYPE NOT IN ('L','B','E','C') THEN -1 * DETJOBC2.CREDIT_AMT ELSE 0 END) DEL_LTR" & vbCrLf _
            & ", 0 DEL_DSC" & vbCrLf _
            & ", -1 * DETJOBC1.INV_TOTAL_AMOUNT DEL_NET" & vbCrLf _
            & " FROM DETJOBC1,DETJOBM2,DETJOBC2 WHERE DETJOBC1.OPS_YYYYPP Between '000000' AND '000001'" & vbCrLf _
            & " AND DETJOBM2.JOB_NO = DETJOBC2.JOB_NO" & vbCrLf _
            & " AND DETJOBM2.JOB_LNO = DETJOBC2.JOB_LNO" & vbCrLf _
            & " AND DETJOBC1.INV_NO = DETJOBC2.INV_NO" & vbCrLf _
            & " GROUP BY DETJOBC1.INV_DATE, DETJOBC1.INV_NO,DETJOBC1.INV_TOTAL_AMOUNT" & vbCrLf _
            & ") GROUP BY INV_DATE" & vbCrLf _
            & ") GROUP BY DIVISION_CODE, INV_DATE"

            ASCMAIN1.sql = sqlx
            ARTGLARS = ASCMAIN1.Temp_Table
            sqlARTGLARS = "Insert into " & ARTGLARS _
            & " " & Replace(Replace(ASCMAIN1.sql, "'000000'", ":PARM1"), "'000001'", ":PARM2")

            ASCMAIN1.sql = "Select * from " & ARTGLARS
            Create_TDA(.Tables.Add, "ARTGLARS", "**", 0, False)
            .Tables("ARTGLARS").Columns.Add("GROSS", GetType(System.Decimal), "ISNULL(INV_SLS,0)+ISNULL(INV_FRT,0)+ISNULL(INV_TAX,0)")
            .Tables("ARTGLARS").Columns.Add("CREDITS", GetType(System.Decimal), "ISNULL(CRM_SLS,0)+ISNULL(CRM_FRT,0)+ISNULL(CRM_TAX,0)")
            .Tables("ARTGLARS").Columns.Add("AVG", GetType(System.Decimal), "IIF(ISNULL(INVS,0)=0,0,ISNULL(GROSS,0)/ISNULL(INVS,0))")


            ' ARTGLAR2

            sqlx = "" _
            & "Select INV_DATE, 'T' CODE_TYPE, ORDR_TYPE_CODE CODE_VALUE" _
            & ", SUM (DECODE(INV_TYPE,'I',1,0)) INVS" _
            & ", SUM (DECODE(INV_TYPE,'C',1,0)) CRMS" _
            & ", SUM (NVL(INV_SALES,0) - NVL(INV_FREE_AMT,0)) SLS" _
            & ", SUM (INV_COGS) CGS" _
            & ", SUM (INV_FREIGHT) FRT" _
            & ", SUM (INV_STAX) TAX" _
            & ", SUM (NVL(INV_SAMPLE_SURCHARGE,0) + NVL(INV_MISC_CHG,0)) MSC" _
            & ", SUM (INV_TOTAL_AMOUNT) NET" _
            & ", SUM (DECODE(INV_TYPE,'I',NVL(INV_SALES,0) - NVL(INV_FREE_AMT,0),0)) ISLS" _
            & ", SUM (DECODE(INV_TYPE,'C',INV_SALES,0)) CSLS" _
            & ", SUM (DECODE(INV_TYPE,'I',INV_TOTAL_AMOUNT,0)) INET" _
            & ", SUM (DECODE(INV_TYPE,'C',INV_TOTAL_AMOUNT,0)) CNET" _
            & "         from SOTINVH1 " _
            & " where ORDR_YYYYPP_UPDATED Between '000000' AND '000001'" _
            & " group by INV_DATE, ORDR_TYPE_CODE"

            ASCMAIN1.sql = sqlx _
            & " union " _
            & Replace(Replace(sqlx, "'T' CODE_TYPE", "'C' CODE_TYPE"), "ORDR_TYPE_CODE", "CUST_CODE") _
            & " union " _
            & Replace(Replace(sqlx, "'T' CODE_TYPE", "'S' CODE_TYPE"), "ORDR_TYPE_CODE", "SREP_CODE")
            ARTGLAR2 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ARTGLAR2 & " Add DESC_VALUE VARCHAR2(60)")
            ASCDATA1.ExecuteSQL("Alter Table " & ARTGLAR2 & " Add CODE2_VALUE VARCHAR2(10)")
            sqlARTGLAR2 = "Insert into " & ARTGLAR2 _
            & " (INV_DATE, CODE_TYPE, CODE_VALUE, INVS, CRMS, SLS, CGS, FRT, TAX, MSC, NET, ISLS, CSLS, INET, CNET) " _
            & Replace(Replace(ASCMAIN1.sql, "'000000'", ":PARM1"), "'000001'", ":PARM2")

            ASCMAIN1.sql = "Select * from " & ARTGLAR2 & " where INV_DATE = :PARM1 and CODE_TYPE = :PARM2"
            Create_TDA(.Tables.Add, "ARTGLAR2", "**", 0, False, "DV", 0) ' 3
            .Tables("ARTGLAR2").Columns("INVS").DataType = GetType(System.Int32)
            .Tables("ARTGLAR2").Columns("CRMS").DataType = GetType(System.Int32)
            .Tables("ARTGLAR2").Columns.Add("AVG_INV", GetType(System.Double), "IIF(ISNULL(INVS,0)=0,0,ISNULL(INET,0)/ISNULL(INVS,0))")


            ' ARTGLAR1

            ASCMAIN1.sql = "SELECT INV_DATE, TO_CHAR(INV_DATE,'Day') DESC_VALUE, 'x' CODE2_VALUE" _
            & ", SUM (INVS) INVS" _
            & ", SUM (CRMS) CRMS" _
            & ", SUM (SLS) SLS" _
            & ", SUM (CGS) CGS" _
            & ", SUM (TAX) TAX" _
            & ", SUM (FRT) FRT" _
            & ", SUM (MSC) MSC" _
            & ", SUM (NET) NET" _
            & ", SUM (ISLS) ISLS" _
            & ", SUM (CSLS) CSLS" _
            & ", SUM (INET) INET" _
            & ", SUM (CNET) CNET" _
            & "  FROM " & ARTGLAR2 _
            & " where CODE_TYPE = 'T'" _
            & " GROUP BY INV_DATE"
            Create_TDA(.Tables.Add, "ARTGLAR1", "**", 0, False, "", 1)
            .Tables("ARTGLAR1").Columns("INVS").DataType = GetType(System.Int32)
            .Tables("ARTGLAR1").Columns("CRMS").DataType = GetType(System.Int32)
            .Tables("ARTGLAR1").Columns.Add("AVG_INV", GetType(System.Double), "IIF(ISNULL(INVS,0)=0,0,ISNULL(INET,0)/ISNULL(INVS,0))")



            ' ARTGLARC

            ASCMAIN1.sql = "" _
            & "SELECT ARTSTMT1.CUST_CODE, ARTCUST1.CUST_NAME, ARTCUST1.POST_CODE" _
            & ", ARTSTMT1.TOTAL_DUE BEG_BAL" _
            & ", ARTSTMT1.TOTAL_DUE INV" _
            & ", ARTSTMT1.TOTAL_DUE RTN" _
            & ", ARTSTMT1.TOTAL_DUE CRM" _
            & ", ARTSTMT1.TOTAL_DUE PMT" _
            & ", ARTSTMT1.TOTAL_DUE DED" _
            & ", ARTSTMT1.TOTAL_DUE ECP" _
            & ", ARTSTMT1.TOTAL_DUE END_BAL" _
            & " FROM ARTSTMT1,ARTCUST1 WHERE ROWNUM < 1"
            ARTGLARC = ASCMAIN1.Temp_Table
            ASCMAIN1.sql = "Select * from " & ARTGLARC
            Create_TDA(.Tables.Add, "ARTGLARC", "**", 0, False, "", 1)
            .Tables("ARTGLARC").Columns.Add("OOBAL", GetType(System.Decimal), _
                "ISNULL(BEG_BAL,0)+ISNULL(INV,0)+ISNULL(RTN,0)+ISNULL(CRM,0)-ISNULL(PMT,0)-ISNULL(DED,0)+ISNULL(ECP,0)-ISNULL(END_BAL,0)")

            ' ARTGLARR

            ORDR_TYPE_CODEs = New String() {"REG", "B2B", "B2C", "DEL", "RTN"}
            Dim SLSs As String = ""

            MISC_CHG_CODEs = New String() {"RSF", "FC", "FS", "B2C", "ECP", "NRSF"}
            Dim MSCs As String = ""

            PYMT_SOURCEs = New String() {"MAN", "BOX", "CC", "B2C"}
            Dim PMTs As String = ""

            REASON_CODEs = New String() _
            {"SALDIS", "RESTFE", "BANKFE", "CCSUCH", "FUELSR", "FINCHR"}
            'REASON_CODEs(0) = ROWs("ARTPARM1").Item("AR_PARM_REASON_CODE_DISC") & ""
            'REASON_CODEs(1) = ROWs("ARTPARM1").Item("AR_PARM_REASON_CODE_WOFF") & ""
            Dim DEDs As String = ""

            Dim sqlSLSX As String = ""
            ASCMAIN1.sql = "Select 'SLS' TYPE, REGISTER_DATE" & vbCrLf
            For Each ORDR_TYPE_CODE As String In ORDR_TYPE_CODEs
                ASCMAIN1.sql &= ", SUM (DECODE(SOTINVH1.ORDR_TYPE_CODE,'" & ORDR_TYPE_CODE & "',NVL(INV_SALES,0)-NVL(INV_FREE_AMT,0))) SLS_" & ORDR_TYPE_CODE & vbCrLf
                SLSs &= ", SUM (SLS_" & ORDR_TYPE_CODE & ") SLS_" & ORDR_TYPE_CODE
                sqlSLSX &= ",'" & ORDR_TYPE_CODE & "'"
            Next
            sqlSLSX = "SUM (CASE WHEN SOTINVH1.ORDR_TYPE_CODE NOT IN (" & Mid(sqlSLSX, 2) & ") THEN NVL(INV_SALES,0)-NVL(INV_FREE_AMT,0) ELSE 0 END)"
            ASCMAIN1.sql &= ", " & sqlSLSX & " SLSX, SUM (NVL(INV_SALES,0)-NVL(INV_FREE_AMT,0)) SLS" & vbCrLf
            ASCMAIN1.sql &= ", SUM (NVL(INV_STAX,0)) TAX" & vbCrLf
            ASCMAIN1.sql &= ", SUM (NVL(INV_FREIGHT,0)) FRT" & vbCrLf
            SLSs &= ", SUM(SLSX) SLSX, SUM (SLS) SLS, SUM (TAX) TAX, SUM (FRT) FRT" & vbCrLf

            ASCMAIN1.sql &= ", SUM (NVL(INV_SAMPLE_SURCHARGE,0)) MSC_SAM" & vbCrLf
            MSCs &= ", SUM (MSC_SAM) MSC_SAM" & vbCrLf

            Dim sqlMSCX As String = ""
            For Each MISC_CHG_CODE As String In MISC_CHG_CODEs
                ASCMAIN1.sql &= ", SUM (DECODE(SOTINVH1.MISC_CHG_CODE,'" & MISC_CHG_CODE & "',NVL(INV_MISC_CHG,0))) MSC_" & MISC_CHG_CODE & vbCrLf
                MSCs &= ", SUM (MSC_" & MISC_CHG_CODE & ") MSC_" & MISC_CHG_CODE
                sqlMSCX &= ",'" & MISC_CHG_CODE & "'"
            Next
            sqlMSCX = "SUM (CASE WHEN SOTINVH1.MISC_CHG_CODE NOT IN (" & Mid(sqlMSCX, 2) & ") THEN NVL(INV_MISC_CHG,0) ELSE 0 END)"
            ASCMAIN1.sql &= ", " & sqlMSCX & " MSCX, SUM (NVL(INV_SAMPLE_SURCHARGE,0) + NVL(INV_MISC_CHG,0)) MSC" & vbCrLf
            MSCs &= ", SUM (MSCX) MSCX, SUM (MSC) MSC" & vbCrLf
            ASCMAIN1.sql &= "" _
            & ", SUM (DECODE(ORDR_TYPE_CODE,'B2C',0,NVL(INV_TOTAL_AMOUNT,0))) REG" & vbCrLf _
            & ", SUM (DECODE(ORDR_TYPE_CODE,'B2C',NVL(INV_TOTAL_AMOUNT_B2C,0),0)) B2C" & vbCrLf _
            & ", SUM (DECODE(ORDR_TYPE_CODE,'B2C',NVL(INV_TOTAL_AMOUNT_B2C,0) - NVL(INV_TOTAL_AMOUNT,0),0)) ECP" & vbCrLf _
            & ", SUM (DECODE(ORDR_TYPE_CODE,'B2C',NVL(INV_TOTAL_AMOUNT_B2C,0),NVL(INV_TOTAL_AMOUNT,0))) TOT" & vbCrLf
            ASCMAIN1.sql &= ", 0 PMT_" & Join(PYMT_SOURCEs, ", 0 PMT_") & ", 0 PMTX, 0 PMT" & vbCrLf
            ASCMAIN1.sql &= ", 0 DED_" & Join(REASON_CODEs, ", 0 DED_") & ", 0 DEDX, 0 DED" & vbCrLf
            ASCMAIN1.sql &= ", 0 GL_SJ, 0 GL_AR, 0 GL_XX, 0 GL, 0 NET" & vbCrLf
            ASCMAIN1.sql &= " from SOTINVH1 where ORDR_YYYYPP_UPDATED = :PARM1" & vbCrLf _
            & " group by SOTINVH1.REGISTER_DATE" & vbCrLf

            ASCMAIN1.sql &= " UNION " & vbCrLf

            ASCMAIN1.sql &= "Select 'PMT' TYPE, ARTPYMT1.REGISTER_DATE" & vbCrLf
            ASCMAIN1.sql &= ", 0 SLS_" & Join(ORDR_TYPE_CODEs, ", 0 SLS_") & ", 0 SLSX, 0 SLS" & vbCrLf
            ASCMAIN1.sql &= ", 0 TAX, 0 FRT" & vbCrLf
            ASCMAIN1.sql &= ", 0 MSC_SAM" & vbCrLf
            ASCMAIN1.sql &= ", 0 MSC_" & Join(MISC_CHG_CODEs, ", 0 MSC_") & ", 0 MSCX, 0 MSC" & vbCrLf
            ASCMAIN1.sql &= ", 0 REG, 0 B2C, 0 ECP, 0 TOT" & vbCrLf

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
            & " group by ARTPYMT1.REGISTER_DATE" & vbCrLf

            ASCMAIN1.sql &= " UNION " & vbCrLf

            ASCMAIN1.sql &= "Select 'PMT' TYPE, ARTPYMT1.REGISTER_DATE" & vbCrLf
            ASCMAIN1.sql &= ", 0 SLS_" & Join(ORDR_TYPE_CODEs, ", 0 SLS_") & ", 0 SLSX, 0 SLS" & vbCrLf
            ASCMAIN1.sql &= ", 0 TAX, 0 FRT" & vbCrLf
            ASCMAIN1.sql &= ", 0 MSC_SAM" & vbCrLf
            ASCMAIN1.sql &= ", 0 MSC_" & Join(MISC_CHG_CODEs, ", 0 MSC_") & ", 0 MSCX, 0 MSC" & vbCrLf
            ASCMAIN1.sql &= ", 0 REG, 0 B2C, 0 ECP, 0 TOT" & vbCrLf
            ASCMAIN1.sql &= ", 0 PMT_" & Join(PYMT_SOURCEs, ", 0 PMT_") & ", 0 PMTX, 0 PMT" & vbCrLf
            ASCMAIN1.sql &= ", 0 DED_" & Join(REASON_CODEs, ", 0 DED_") & ", SUM (-1 * NVL(CUST_PYMT_AMT,0)) DEDX, SUM (-1 * NVL(CUST_PYMT_AMT,0)) DED" & vbCrLf
            ASCMAIN1.sql &= ", 0 GL_SJ, 0 GL_AR, 0 GL_XX, 0 GL, 0 NET" & vbCrLf
            ASCMAIN1.sql &= " from ARTPYMT1,ARTPYMT2" & vbCrLf _
            & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            & " and ARTPYMT1.OPS_YYYYPP = :PARM1" & vbCrLf _
            & " and NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" & vbCrLf _
            & " and ARTPYMT2.CUST_CODE is Null" & vbCrLf _
            & " group by ARTPYMT1.REGISTER_DATE" & vbCrLf

            ASCMAIN1.sql &= " UNION " & vbCrLf

            ASCMAIN1.sql &= "Select 'D/W' TYPE, ARTPYMT1.REGISTER_DATE" & vbCrLf
            ASCMAIN1.sql &= ", 0 SLS_" & Join(ORDR_TYPE_CODEs, ", 0 SLS_") & ", 0 SLSX, 0 SLS" & vbCrLf
            ASCMAIN1.sql &= ", 0 TAX, 0 FRT" & vbCrLf
            ASCMAIN1.sql &= ", 0 MSC_SAM" & vbCrLf
            ASCMAIN1.sql &= ", 0 MSC_" & Join(MISC_CHG_CODEs, ", 0 MSC_") & ", 0 MSCX, 0 MSC" & vbCrLf
            ASCMAIN1.sql &= ", 0 REG, 0 B2C, 0 ECP, 0 TOT" & vbCrLf
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
            ASCMAIN1.sql &= " from ARTPYMT1,ARTPYMT3" & vbCrLf _
            & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT3.PYMT_BATCH_NO" & vbCrLf _
            & " and ARTPYMT1.OPS_YYYYPP = :PARM1" & vbCrLf _
            & " group by ARTPYMT1.REGISTER_DATE" & vbCrLf

            ASCMAIN1.sql &= " UNION " & vbCrLf

            Dim DEDX_CODES As String = ""

            ASCMAIN1.sql &= "Select 'DED' TYPE, ARTPYMT1.REGISTER_DATE" & vbCrLf
            ASCMAIN1.sql &= ", 0 SLS_" & Join(ORDR_TYPE_CODEs, ", 0 SLS_") & ", 0 SLSX, 0 SLS" & vbCrLf
            ASCMAIN1.sql &= ", 0 TAX, 0 FRT" & vbCrLf
            ASCMAIN1.sql &= ", 0 MSC_SAM" & vbCrLf
            ASCMAIN1.sql &= ", 0 MSC_" & Join(MISC_CHG_CODEs, ", 0 MSC_") & ", 0 MSCX, 0 MSC" & vbCrLf
            ASCMAIN1.sql &= ", 0 REG, 0 B2C, 0 ECP, 0 TOT" & vbCrLf
            ASCMAIN1.sql &= ", 0 PMT_" & Join(PYMT_SOURCEs, ", 0 PMT_") & ", 0 PMTX, 0 PMT" & vbCrLf
            For Each REASON_CODE As String In REASON_CODEs
                DEDX_CODES &= ",'" & REASON_CODE & "'"
                ASCMAIN1.sql &= ", SUM (DECODE(ARTPYMT5.REASON_CODE,'" & REASON_CODE & "',NVL(ARTPYMT5.GL_DIST_AMT,0))) DED_" & REASON_CODE & vbCrLf
            Next
            ASCMAIN1.sql &= ", SUM (CASE WHEN ARTPYMT5.REASON_CODE IN (" & Mid(DEDX_CODES, 2) & ") THEN 0 ELSE NVL(ARTPYMT5.GL_DIST_AMT,0) END) DEDX, SUM (NVL(ARTPYMT5.GL_DIST_AMT,0)) DED" & vbCrLf
            ASCMAIN1.sql &= ", 0 GL_SJ, 0 GL_AR, 0 GL_XX, 0 GL, 0 NET" & vbCrLf
            ASCMAIN1.sql &= " from ARTPYMT1,ARTPYMT5" & vbCrLf _
            & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT5.PYMT_BATCH_NO" & vbCrLf _
            & " and NVL(ARTPYMT5.CHARGEBACK_IND,'0') = '0'" & vbCrLf _
            & " and ARTPYMT1.OPS_YYYYPP = :PARM1" & vbCrLf _
            & " group by ARTPYMT1.REGISTER_DATE" & vbCrLf

            ASCMAIN1.sql &= " UNION " & vbCrLf

            ASCMAIN1.sql &= "Select 'GL' TYPE, GLTDETL1.DETL_CTL_DATE" & vbCrLf _
            & ", 0 SLS_" & Join(ORDR_TYPE_CODEs, ", 0 SLS_") & ", 0 SLSX, 0 SLS" & vbCrLf _
            & ", 0 TAX, 0 FRT" & vbCrLf _
            & ", 0 MSC_SAM" & vbCrLf _
            & ", 0 MSC_" & Join(MISC_CHG_CODEs, ", 0 MSC_") & ", 0 MSCX, 0 MSC" & vbCrLf _
            & ", 0 REG, 0 B2C, 0 ECP, 0 TOT" & vbCrLf _
            & ", 0 PMT_" & Join(PYMT_SOURCEs, ", 0 PMT_") & ", 0 PMTX, 0 PMT" & vbCrLf _
            & ", 0 DED_" & Join(REASON_CODEs, ", 0 DED_") & ", 0 DEDX, 0 DED" & vbCrLf _
            & ", SUM (CASE WHEN GLTJRNL1.JOURNAL_TYPE IN ('OPSJ','ARFC','ARFR','ARFS','AREC') " & vbCrLf _
            & "            THEN GLTDETL1.DETL_POSTING_AMT ELSE 0 END) GL_SJ" & vbCrLf _
            & ", SUM (CASE WHEN GLTJRNL1.JOURNAL_TYPE IN ('ARCR') " & vbCrLf _
            & "            THEN GLTDETL1.DETL_POSTING_AMT ELSE 0 END) GL_AR" & vbCrLf _
            & ", SUM (CASE WHEN GLTJRNL1.JOURNAL_TYPE NOT IN ('OPSJ','ARFC','ARFR','ARFS','AREC','ARCR') " & vbCrLf _
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
            & ", SUM (REG) REG, SUM (B2C) B2C, SUM (ECP) ECP, SUM (TOT) TOT" & vbCrLf _
            & PMTs & vbCrLf _
            & DEDs & vbCrLf _
            & ", SUM (GL_SJ) GL_SJ, SUM (GL_AR) GL_AR, SUM (GL_XX) GL_XX, SUM (GL) GL" & vbCrLf _
            & " from (" & ASCMAIN1.sql & ") X" & vbCrLf _
            & " group by REGISTER_DATE"
            ARTGLARR = ASCMAIN1.Temp_Table(Replace(ASCMAIN1.sql, ":PARM1", "'000000'"))

            Create_TDA(.Tables.Add, "ARTGLARR", "**", 0, False, "V", 0)
            .Tables("ARTGLARR").Columns.Add("RF_GL", GetType(System.Decimal), "")
            .Tables("ARTGLARR").Columns.Add("RF_AR", GetType(System.Decimal), "")

            'Bank Reconcile
            Dim sql As String
            sql = " SELECT * FROM ARTPYMT1 WHERE OPS_YYYYPP BETWEEN :PARM1 AND :PARM2"
            Create_TDA(.Tables.Add, "ARTPYMT1X", sql, 0, False, "VV", 1)

            sql = " SELECT * FROM ARTPYMT2 WHERE PYMT_BATCH_NO BETWEEN :PARM1 AND :PARM2"
            Create_TDA(.Tables.Add, "ARTPYMT2X", sql, 0, False, "VV", 2)

            .Relations.Add("ARTPYMT1X_ARTPYMT2X" _
            , New DataColumn() {.Tables("ARTPYMT1X").Columns("PYMT_BATCH_NO")} _
            , New DataColumn() {.Tables("ARTPYMT2X").Columns("PYMT_BATCH_NO")})

            .Tables("ARTPYMT1X").Columns.Add("BATCH_AMT", GetType(System.Decimal), "SUM (CHILD.CUST_PYMT_AMT)")
            .Tables("ARTPYMT1X").Columns.Add("NO_TRANS", GetType(System.Int32), "COUNT (CHILD.PYMT_BATCH_LNO)")

            Create_TDA(.Tables.Add, "GLTPARM2", "*")
            sql = "SELECT SOTMISC1.ACCT_CODE, SOTMISC1.MISC_CHG_DESC, SOTINVH1.MISC_CHG_CODE"
            For iCol As Integer = 1 To periodDays
                sql &= ", SOTINVH1.INV_MISC_CHG COL_" & iCol
            Next
            sql &= " FROM SOTINVH1, SOTMISC1"
            sql &= " WHERE ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTMISCX", sql, 0, False, "", 0)

            sql = String.Empty
            For iCol As Integer = 1 To periodDays
                sql &= " + COL_" & iCol
            Next
            sql = sql.Trim
            sql = sql.Substring(1)

            .Tables("SOTMISCX").Columns.Add("COL_TOTAL", GetType(System.Decimal), sql)

            For iCol As Integer = 0 To dst.Tables("SOTMISCX").Columns.Count - 1
                grdSOTMISCX.DisplayLayout.Bands(0).Columns.Add(dst.Tables("SOTMISCX").Columns(iCol).ColumnName)
                Select Case dst.Tables("SOTMISCX").Columns(iCol).ColumnName
                    Case "ACCT_CODE"
                        grdSOTMISCX.DisplayLayout.Bands(0).Columns(dst.Tables("SOTMISCX").Columns(iCol).ColumnName).Width = 100
                        grdSOTMISCX.DisplayLayout.Bands(0).Columns(dst.Tables("SOTMISCX").Columns(iCol).ColumnName).Header.Caption = "Account"
                        Create_Summary(grdSOTMISCX, "ACCT_CODE", "Count")

                    Case "MISC_CHG_DESC"
                        grdSOTMISCX.DisplayLayout.Bands(0).Columns(dst.Tables("SOTMISCX").Columns(iCol).ColumnName).Width = 200
                        grdSOTMISCX.DisplayLayout.Bands(0).Columns(dst.Tables("SOTMISCX").Columns(iCol).ColumnName).Header.Caption = "Description"

                    Case "MISC_CHG_CODE"
                        grdSOTMISCX.DisplayLayout.Bands(0).Columns(dst.Tables("SOTMISCX").Columns(iCol).ColumnName).Width = 120
                        grdSOTMISCX.DisplayLayout.Bands(0).Columns(dst.Tables("SOTMISCX").Columns(iCol).ColumnName).Header.Caption = "Misc Chg Code"

                    Case Else
                        grdSOTMISCX.DisplayLayout.Bands(0).Columns(dst.Tables("SOTMISCX").Columns(iCol).ColumnName).Width = 75
                        grdSOTMISCX.DisplayLayout.Bands(0).Columns(dst.Tables("SOTMISCX").Columns(iCol).ColumnName).CellAppearance.TextHAlign = HAlign.Right
                        Create_Summary(grdSOTMISCX, dst.Tables("SOTMISCX").Columns(iCol).ColumnName, "Sum", "", "#,###0.00")
                End Select
            Next
        End With

        grdARTGLAR1.DataSource = dst.Tables("ARTGLAR1")
        grdARTGLAR2.DataSource = dst.Tables("ARTGLAR2")
        grdARTGLARC.DataSource = dst.Tables("ARTGLARC")
        grdARTGLARR.DataSource = dst.Tables("ARTGLARR")
        grdARTGLARS.DataSource = dst.Tables("ARTGLARS")
        grdSOTMISCX.DataSource = dst.Tables("SOTMISCX")

        viewARTPYMT1X = New DataView(dst.Tables("ARTPYMT1X"))
        grdARTPYMT1.DataSource = viewARTPYMT1X

        Create_Lookup("GLTACCT1")

        Format_grdARTGLAR1_2()
        Format_grdARTGLARC()
        Format_grdARTGLARR()
        Format_grdARTGLARS()

        Create_Summary(grdARTPYMT1, "PYMT_BATCH_NO", "Count")
        Create_Summary(grdARTPYMT1, "BATCH_AMT", "Sum")
        Create_Summary(grdARTPYMT1, "NO_TRANS", "Sum")
        ASCMAIN1.Add_Value_List(grdARTPYMT1, "PYMT_STATUS", Nothing, Nothing, 1)

        Create_Summary(grdARTPYMT1, "PYMT_BATCH_LNO", "Count", "ARTPYMT1X_ARTPYMT2X")
        Create_Summary(grdARTPYMT1, "CUST_PYMT_AMT", "Sum", "ARTPYMT1X_ARTPYMT2X")
        optChoice.CheckedIndex = 0

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                OPS_YYYYPP_B = String.Empty
                OPS_YYYYPP_E = String.Empty

                Select Case optChoice.Value
                    Case "P"
                        Validate_Code("OPS_YYYYPP")
                        OPS_YYYYPP_B = MyBase.Absx1.txtFor("OPS_YYYYPP").Text
                        OPS_YYYYPP_E = OPS_YYYYPP_B

                    Case "Y"
                        OPS_YYYYPP_B = MyBase.Absx1.txtFor("OPS_YYYY").Text
                        If OPS_YYYYPP_B.Length <> 4 OrElse Not IsNumeric(OPS_YYYYPP_B) Then
                            EMsg &= "Invalid Year Specified."
                            Exit Select
                        End If

                        OPS_YYYYPP_B = MyBase.Absx1.txtFor("OPS_YYYY").Text & "01"
                        OPS_YYYYPP_E = MyBase.Absx1.txtFor("OPS_YYYY").Text & "12"

                        If ASCMAIN1.CYP.StartsWith(MyBase.Absx1.txtFor("OPS_YYYY").Text) Then
                            OPS_YYYYPP_E = ASCMAIN1.CYP
                        End If
                End Select


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
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Done"
                Call Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                '.Groups("Screen Control").Items("Excel").Settings.Enabled = iScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        UltraTabControl1.Visible = tf
        Setup_tabMain()
        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("ARTGLAR1").Rows.Clear()
        dst.Tables("ARTGLAR2").Rows.Clear()
        dst.Tables("ARTGLARC").Rows.Clear()
        dst.Tables("ARTGLARR").Rows.Clear()
        dst.Tables("ARTGLARS").Rows.Clear()
        dst.Tables("SOTMISCX").Rows.Clear()
        dst.EnforceConstraints = True

        Absx1.txtFor("OPS_YYYYPP").Text = ASCMAIN1.CYP
        Absx1.txtFor("OPS_YYYY").Text = ASCMAIN1.CYP.Substring(0, 4)

        dst.Tables("ARTPYMT2X").Rows.Clear()
        dst.Tables("ARTPYMT1X").Rows.Clear()
        viewARTPYMT1X.RowFilter = String.Empty

        Show_Filter(grdARTPYMT1, True)
        Clear_All_Filters(grdARTPYMT1)
        grdARTPYMT1.DisplayLayout.GroupByBox.Hidden = True

    End Sub

    Sub Load_Record()
        Call ASCMAIN1.Progress("Now Loading Data")
        Call Save_Header_Fields(UltraGroupBox1)

        LOAD_ARTGLAR1()
        Load_ARTGLARC()
        LOAD_ARTGLARS()

        Fill_Records("ARTGLARR", OPS_YYYYPP_E)

        ASCMAIN1.sql = "Select SUM (NVL(TOTAL_DUE, 0) + NVL(AGE_0, 0)) from ARTSTMT1 " _
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

        Sort_grdColumns(grdARTGLARR, "RECORD_TYPE,REGISTER_DATE", True)
        grdARTGLARR.Text = "A/R Roll Forward for " & Absx1.txtFor("LEGEND").Text

        Setup_tabMain()

        Set_Descriptions("SLS")

        Fill_Records("ARTPYMT1X", New Object() {OPS_YYYYPP_B, OPS_YYYYPP_E})
        Dim minPYMT_BATCH_NO As String = dst.Tables("ARTPYMT1X").Compute("MIN(PYMT_BATCH_NO)", "") & String.Empty
        Dim maxPYMT_BATCH_NO As String = dst.Tables("ARTPYMT1X").Compute("MAX(PYMT_BATCH_NO)", "") & String.Empty
        Fill_Records("ARTPYMT2X", New Object() {minPYMT_BATCH_NO, maxPYMT_BATCH_NO})

        SetUpMiscCharges()

        Call ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        Call BeginTrans()


        Call CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special( _
        ByVal ctl As Control, _
        ByVal COLUMN_NAME As String, _
        Optional ByRef sql_where As String = "", _
        Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            Case "OPS_YYYY"
                sql_where = "SUBSTR(OPS_YYYYPP,1,4) <= '" & ASCMAIN1.CYP.Substring(0, 4) & "'"
            Case "OPS_YYYYPP"
                sql_where = "OPS_YYYYPP <= '" & ASCMAIN1.CYP & "'"

        End Select

    End Sub
#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdARTGLARR, "SSSSSB" _
                             , "Show Sales Order Types" _
                             , "Show Misc Charge Codes" _
                             , "Show Pymt Sources" _
                             , "Show Deduction Codes" _
                             , "Show GL Columns" _
                             , "Mgmt Summary")

        Call Load_Popup_Menu(grdARTGLAR1, "SS" _
                     , "Show Filter" _
                     , "Show GroupBox")

        Call Load_Popup_Menu(grdARTGLAR2, "SS" _
                     , "Show Filter" _
                     , "Show GroupBox")

        Call Load_Popup_Menu(grdARTGLARC, "SS" _
             , "Show Filter" _
             , "Show GroupBox")

        Call Load_Popup_Menu(grdARTPYMT1, "SS" _
             , "Show Filter" _
             , "Show GroupBox")


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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool

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
                grdARTGLARR.DisplayLayout.Bands(0).Columns("MSC_SAM").Hidden = (tlb_sbt.Tag = "X")
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

                    .Columns("REG").Hidden = True
                    .Columns("B2C").Hidden = True
                    .Columns("ECP").Hidden = True

                    .Columns("REG").Hidden = True
                    .Columns("B2C").Hidden = True
                    .Columns("ECP").Hidden = True
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

                    .Columns("REG").Hidden = False
                    .Columns("B2C").Hidden = False
                    .Columns("ECP").Hidden = False
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


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Sales Order Inquiry"
            '    Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Text
            '    Context_Launch("Load", ORDR_NO, e.Tool.Key, "SOFORDRI", "F", "SO")

        End Select
    End Sub

#End Region

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
                    Call Click_Command("Load", e)
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

    Private Sub grdARTGLAR1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTGLAR1.AfterRowActivate
        Setup_ARTGLAR2()
    End Sub

    Sub Setup_ARTGLARS(Optional ByVal initialize As Boolean = False)

        Dim dvw As DataView = dst.Tables("ARTGLARS").DefaultView
        dvw.RowFilter = "DIVISION_CODE = '" & optDIVISION_CODE.Value & "'"
        Sort_grdColumns(grdARTGLARS, "WEEK,LINE,INV_DATE", True)
        grdARTGLARS.DisplayLayout.Bands(0).SortedColumns.Add("WEEK", False, True)
        grdARTGLARS.Rows.ExpandAll(True)
        'e.Layout.GroupByRowDescriptionMaskDefault = _
        '   "[caption] of [value] Min: [min:CustomerID], Max:[max:CustomerID]"
        grdARTGLARS.DisplayLayout.Override.GroupByRowDescriptionMask = "[caption] [value]"

        grdARTGLARS.DisplayLayout.Override.SummaryDisplayArea = UltraWinGrid.SummaryDisplayAreas.BottomFixed + UltraWinGrid.SummaryDisplayAreas.GroupByRowsFooter


        With grdARTGLARS.DisplayLayout.Bands(0)
            '.Columns("WEB_SLS").Hidden = (optDIVISION_CODE.Value = "DEL")
            '.Columns("WEB_FRT").Hidden = (optDIVISION_CODE.Value = "DEL")
            '.Columns("WEB_TAX").Hidden = (optDIVISION_CODE.Value = "DEL")
            '.Columns("WEB_FEE").Hidden = (optDIVISION_CODE.Value = "DEL")
            '.Columns("WEB_NET").Hidden = (optDIVISION_CODE.Value = "DEL")

            .Columns("DEL_SLS").Hidden = (optDIVISION_CODE.Value <> "DEL")
            .Columns("DEL_EDG").Hidden = (optDIVISION_CODE.Value <> "DEL")
            .Columns("DEL_ARC").Hidden = (optDIVISION_CODE.Value <> "DEL")
            .Columns("DEL_LTR").Hidden = (optDIVISION_CODE.Value <> "DEL")
            .Columns("DEL_DSC").Hidden = (optDIVISION_CODE.Value <> "DEL")
            .Columns("DEL_NET").Hidden = (optDIVISION_CODE.Value <> "DEL")
        End With
    End Sub

    Sub Setup_ARTGLAR1(Optional ByVal initialize As Boolean = False)

        If optD.Value = "ByDate" Then
            If grdARTGLAR2.ActiveRow Is Nothing OrElse Not grdARTGLAR2.ActiveRow.IsDataRow Then
                grdARTGLAR1.Visible = False
            Else
                grdARTGLAR1.Visible = True

                ASCMAIN1.sql = "SELECT INV_DATE" _
                & ", TO_CHAR(INV_DATE,'Day') DESC_VALUE, 'x' CODE2_VALUE" _
                & ", SUM (INVS) INVS" _
                & ", SUM (CRMS) CRMS" _
                & ", SUM (SLS) SLS" _
                & ", SUM (CGS) CGS" _
                & ", SUM (TAX) TAX" _
                & ", SUM (FRT) FRT" _
                & ", SUM (MSC) MSC" _
                & ", SUM (NET) NET" _
                & ", SUM (ISLS) ISLS" _
                & ", SUM (CSLS) CSLS" _
                & ", SUM (INET) INET" _
                & ", SUM (CNET) CNET" _
                & "  FROM " & ARTGLAR2 _
                & " where CODE_TYPE = '" & optDateBy.Value & "'" _
                & " and CODE_VALUE = '" & grdARTGLAR2.ActiveRow.Cells("CODE_VALUE").Text & "'" _
                & " GROUP BY INV_DATE"

                Call Fill_Records("ARTGLAR1", "", True, ASCMAIN1.sql)
                Call Sort_grdColumns(grdARTGLAR1, "INV_DATE")
                grdARTGLAR1.Text = optDateBy.Text & " by Date"
            End If
        Else
            If Not initialize Then
                Exit Sub
            End If

            grdARTGLAR1.Visible = True

            Fill_Records("ARTGLAR1")
            Call Sort_grdColumns(grdARTGLAR1, "INV_DATE", True)
            'Setup_ARTGLAR2()
            grdARTGLAR1.Text = "Summary by Date"
        End If

        For Each rowARTGLAR1 As DataRow In dst.Tables("ARTGLAR1").Rows
            Dim INV_DATE As Date = rowARTGLAR1.Item("INV_DATE")
            Dim DAYS As Int32 = INV_DATE.Subtract(PPE).TotalDays
            If DAYS > 0 Then
                rowARTGLAR1.Item("CODE2_VALUE") = 1 + (DAYS - 1) \ 7
            End If
        Next

    End Sub

    Sub Setup_ARTGLAR2(Optional ByVal initialize As Boolean = False)

         If optD.Value = "DateBy" Then
            If grdARTGLAR1.ActiveRow Is Nothing OrElse Not grdARTGLAR1.ActiveRow.IsDataRow Then
                grdARTGLAR2.Visible = False
            Else
                grdARTGLAR2.Visible = True

                Dim INV_DATE As Date = grdARTGLAR1.ActiveRow.Cells("INV_DATE").Text

                Call Fill_Records("ARTGLAR2", New String() {Format(INV_DATE, "dd-MMM-yyyy"), optDateBy.Value})
                Call Sort_grdColumns(grdARTGLAR2, "CODE_VALUE")
                grdARTGLAR2.Text = "Summary Breakdown by " & optDateBy.Text & " for " & Format(INV_DATE, "MM/dd/yyyy")
            End If
        Else
            If Not initialize Then
                Exit Sub
            End If

            grdARTGLAR2.Visible = True



            ASCMAIN1.sql = "SELECT NULL INV_DATE" _
            & ", CODE_TYPE, CODE_VALUE" _
            & ", SUM (INVS) INVS" _
            & ", SUM (CRMS) CRMS" _
            & ", SUM (SLS) SLS" _
            & ", SUM (CGS) CGS" _
            & ", SUM (TAX) TAX" _
            & ", SUM (FRT) FRT" _
            & ", SUM (MSC) MSC" _
            & ", SUM (NET) NET" _
            & ", SUM (ISLS) ISLS" _
            & ", SUM (CSLS) CSLS" _
            & ", SUM (INET) INET" _
            & ", SUM (CNET) CNET" _
            & ", DESC_VALUE, CODE2_VALUE" _
            & "  FROM " & ARTGLAR2 _
            & " where CODE_TYPE = '" & optDateBy.Value & "'" _
            & " GROUP BY CODE_TYPE, CODE_VALUE" _
            & ", DESC_VALUE, CODE2_VALUE"

            'ASCMAIN1.sql = "Select * from " & ARTGLAR2 & " where CODE_TYPE = '" & optDateBy.Value & "'"
            Call Fill_Records("ARTGLAR2", "", True, ASCMAIN1.sql)
            Call Sort_grdColumns(grdARTGLAR2, "CODE_VALUE")
            grdARTGLAR2.Text = "Summary by " & optDateBy.Text

            'Setup_ARTGLAR1()
        End If

        With grdARTGLAR2.DisplayLayout.Bands(0)
            .Columns("CODE_VALUE").Header.Caption = optDateBy.Text
            Select Case optDateBy.Value
                Case "T"
                    .Columns("CODE2_VALUE").Header.Caption = ""
                    .Columns("DESC_VALUE").Header.Caption = "Description"
                Case "S"
                    .Columns("CODE2_VALUE").Header.Caption = "Type"
                    .Columns("DESC_VALUE").Header.Caption = "Name"
                Case "C"
                    .Columns("CODE2_VALUE").Header.Caption = "SRep"
                    .Columns("DESC_VALUE").Header.Caption = "Name"
            End Select
        End With

    End Sub

    Private Sub optDateBy_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optDateBy.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_optD()
        Setup_ARTGLAR2(True)
    End Sub

    Private Sub UltraTabControl1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged
        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()

        If UltraTabControl1.SelectedTab Is Nothing Then
            UltraExplorerBar1.Groups("Summary by Date").Visible = False
            UltraExplorerBar1.Groups("Descriptions").Visible = False
            UltraExplorerBar1.Groups("Sales Analysis").Visible = False
        Else
            UltraExplorerBar1.Groups("Summary by Date").Visible = ScreenMode And (UltraTabControl1.SelectedTab.Key = "Summary by Date")
            UltraExplorerBar1.Groups("Descriptions").Visible = ScreenMode And (UltraTabControl1.SelectedTab.Key = "AR Roll Forward")
            UltraExplorerBar1.Groups("Sales Analysis").Visible = ScreenMode And (UltraTabControl1.SelectedTab.Key = "Sales Analysis")
        End If

    End Sub

    Sub Load_ARTGLAR1()
        ASCDATA1.ExecuteSQL("Truncate Table " & ARTGLAR2)
        ASCDATA1.ExecuteSQL(sqlARTGLAR2, "VV", New String() {OPS_YYYYPP_B, OPS_YYYYPP_E})

        ASCDATA1.ExecuteSQL("Update " & ARTGLAR2 & " Set DESC_VALUE = (SELECT ORDR_TYPE_DESC from SOTTYPE1 where ORDR_TYPE_CODE = CODE_VALUE) where CODE_TYPE = 'T'")

        ASCDATA1.ExecuteSQL("Update " & ARTGLAR2 & " Set DESC_VALUE = (SELECT SREP_NAME from SOTSREP1 where SREP_CODE = CODE_VALUE) where CODE_TYPE = 'S'")
        ASCDATA1.ExecuteSQL("Update " & ARTGLAR2 & " Set CODE2_VALUE = (SELECT SREP_TYPE from SOTSREP1 where SREP_CODE = CODE_VALUE) where CODE_TYPE = 'S'")
        ASCDATA1.ExecuteSQL("Update " & ARTGLAR2 & " Set DESC_VALUE = (SELECT CUST_NAME from ARTCUST1 where CUST_CODE = CODE_VALUE) where CODE_TYPE = 'C'")
        ASCDATA1.ExecuteSQL("Update " & ARTGLAR2 & " Set CODE2_VALUE = (SELECT SREP_CODE from ARTCUST1 where CUST_CODE = CODE_VALUE) where CODE_TYPE = 'C'")

        ASCMAIN1.sql = "Select PRD_END_DATE from GLTPARM2 where OPS_YYYYPP = '" & ASCMAIN1.Period_Calc(OPS_YYYYPP_E, -1) & "'"
        PPE = CDate(ASCDATA1.GetDataValue)

        If optD.Value = "DateBy" Then
            Setup_optD()
            Setup_ARTGLAR1(True)
        Else
            optD.Value = "DateBy"
        End If
    End Sub

    Sub Load_ARTGLARS()

        ' Get 1st and Last Dates of Month
        Dim RYP As String = OPS_YYYYPP_E
        Dim LYP As String = ASCMAIN1.Period_Calc(RYP, -1)


        ASCMAIN1.sql = "Select PRD_END_DATE from GLTPARM2 " _
        & "where OPS_YYYYPP = '" & LYP & "'"
        Dim DATE_FIRST As Date = CDate(ASCDATA1.GetDataValue).AddDays(1)
        ASCMAIN1.sql = "Select PRD_END_DATE from GLTPARM2 " _
        & "where OPS_YYYYPP = '" & RYP & "'"
        Dim DATE_LAST As Date = CDate(ASCDATA1.GetDataValue)

        ' Load Data into Work Table
        ASCDATA1.ExecuteSQL("Truncate Table " & ARTGLARS)
        ASCDATA1.ExecuteSQL(sqlARTGLARS, "VV", New String() {OPS_YYYYPP_B, OPS_YYYYPP_E})

        ' Prepare Summary SQL
        Dim sql_SUM As String = ""
        'Dim GBY As String = ""
        Dim HC() As String = {"DIVISION_CODE", "WEEK", "LINE", "INV_DATE"}
        For i As Int32 = 0 To dst.Tables("ARTGLARS").Columns.Count - 1
            Dim DC As DataColumn = dst.Tables("ARTGLARS").Columns(i)
            If DC.ColumnName = "DIVISION_CODE" Or DC.ColumnName = "WEEK" _
            Or DC.ColumnName = "LINE" Or DC.ColumnName = "INV_DATE" Then
                'GBY &= "," & DC.ColumnName
            ElseIf DC.Expression <> "" Then
            Else
                sql_SUM &= ", Sum (" & DC.ColumnName & ") " & DC.ColumnName
            End If
        Next

        If optChoice.Value = "P" Then
            ' Fix Dates Before and After Month, and on Weekend Dates
            ASCMAIN1.sql = "Update " & ARTGLARS _
            & " Set INV_DATE = '" & Format(DATE_FIRST, "dd-MMM-yyyy") _
            & "' where INV_DATE < '" & Format(DATE_FIRST, "dd-MMM-yyyy") & "'"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update " & ARTGLARS _
            & " Set INV_DATE = '" & Format(DATE_LAST, "dd-MMM-yyyy") _
            & "' where INV_DATE > '" & Format(DATE_LAST, "dd-MMM-yyyy") & "'"
            ASCDATA1.ExecuteSQL()

        End If

        ASCMAIN1.sql = "Update " & ARTGLARS _
        & " Set INV_DATE = INV_DATE + 2 where TO_CHAR(INV_DATE,'D') = 7"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update " & ARTGLARS _
        & " Set INV_DATE = INV_DATE + 1 where TO_CHAR(INV_DATE,'D') = 1"
        ASCDATA1.ExecuteSQL()


        ' Re-Summarize by M-F Dates
        ASCMAIN1.sql = "Insert into " & ARTGLARS _
        & " Select DIVISION_CODE, WEEK, 1 LINE, INV_DATE" _
        & sql_SUM & " from " & ARTGLARS & " group by DIVISION_CODE, WEEK, INV_DATE"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Delete from " & ARTGLARS & " where LINE = 0"
        ASCDATA1.ExecuteSQL()

        ' Establish Relative Week No
        ASCMAIN1.sql = "Update " & ARTGLARS & " ARTGLARS Set WEEK = " _
        & "(Select REL_WEEK from GLTPARM3 where YYYYWW = " _
        & "(SELECT Min (YYYYWW) from GLTPARM3 " _
        & " where WEEK_END_DATE >= ARTGLARS.INV_DATE))"
        ASCDATA1.ExecuteSQL()

        ' Create Consolidated Records
        ASCMAIN1.sql = "Insert into " & ARTGLARS _
        & " Select '*' DIVISION_CODE, WEEK, LINE, INV_DATE" _
        & sql_SUM & " from " & ARTGLARS & " group by WEEK, LINE, INV_DATE"
        ASCDATA1.ExecuteSQL()

        '' Create Weekly Totals
        'ASCMAIN1.sql = "Insert into " & ARTGLARS _
        '& " Select DIVISION_CODE, WEEK, 2 LINE, NULL INV_DATE" _
        '& sql_SUM & " group by DIVISION_CODE, WEEK"
        'ASCDATA1.ExecuteSQL()

        Fill_Records("ARTGLARS")

        If optDIVISION_CODE.Value = "*" Then
            Setup_ARTGLARS(True)
        Else
            optDIVISION_CODE.Value = "*"
        End If
    End Sub

    Sub Load_ARTGLARC()

        ASCMAIN1.sql = "Truncate Table " & ARTGLARC
        ASCDATA1.ExecuteSQL()

        ' & "(CUST_CODE, BEG_BAL, INV, RTN, CRM, PMT, DED, ECP, END_BAL) " & vbCr _

        Dim RYP As String = OPS_YYYYPP_E
        Dim LYP As String = ASCMAIN1.Period_Calc(RYP, -1)

        ASCMAIN1.sql = "Insert into " & ARTGLARC & vbCr _
        & "Select NVL(X.CUST_CODE,'.') CUST_CODE" & vbCr _
        & ", ARTCUST1.CUST_NAME, ARTCUST1.POST_CODE" & vbCr _
        & ", Sum (X.BEG_BAL) BEG_BAL" & vbCr _
        & ", Sum (X.INV) INV" & vbCr _
        & ", Sum (X.RTN) RTN" & vbCr _
        & ", Sum (X.CRM) CRM" & vbCr _
        & ", Sum (X.PMT) PMT" & vbCr _
        & ", Sum (X.DED) DED" & vbCr _
        & ", Sum (X.ECP) ECP" & vbCr _
        & ", Sum (X.END_BAL) END_BAL" & vbCr _
        & "from (" & vbCr _
        & "Select CUST_CODE, SUM (TOTAL_DUE) BEG_BAL" & vbCr _
        & ", 0 INV, 0 RTN, 0 CRM, 0 PMT, 0 DED, 0 ECP, 0 END_BAL" & vbCr _
        & " from ARTSTMT1 WHERE OPS_YYYYPP = '" & LYP & "'" & vbCr _
        & " group by CUST_CODE " & vbCr _
        & " union " & vbCr _
        & "Select SOTINVH1.CUST_CODE, 0 BEG_BAL" & vbCr _
        & ", SUM (DECODE(SOTINVH1.INV_TYPE,'I',SOTINVH1.INV_TOTAL_AMOUNT,0)) INV" & vbCr _
        & ", SUM (CASE WHEN SOTINVH1.INV_TYPE = 'C' AND ORDR_TYPE_CODE = 'RTN' THEN SOTINVH1.INV_TOTAL_AMOUNT ELSE 0 END) RTN" & vbCr _
        & ", SUM (CASE WHEN SOTINVH1.INV_TYPE = 'C' AND ORDR_TYPE_CODE <> 'RTN' THEN SOTINVH1.INV_TOTAL_AMOUNT ELSE 0 END) CRM" & vbCr _
        & ", 0 PMT, 0 DED" & vbCr _
        & ", SUM (CASE WHEN SOTINVH1.ORDR_TYPE_CODE = 'B2C' THEN NVL(SOTINVH1.INV_TOTAL_AMOUNT_B2C,0) - NVL(SOTINVH1.INV_TOTAL_AMOUNT,0) ELSE 0 END) ECP" & vbCr _
        & ", 0 END_BAL" & vbCr _
        & " from SOTINVH1" & vbCr _
        & " where ORDR_YYYYPP_UPDATED = '" & RYP & "'" & vbCr _
        & " group by SOTINVH1.CUST_CODE" & vbCr _
        & " union " & vbCr _
        & "Select ARTPYMT2.CUST_CODE, 0 BEG_BAL, 0 INV, 0 RTN, 0 CRM " & vbCr _
        & ", SUM (ARTPYMT2.CUST_PYMT_AMT) PMT, 0 DED, 0 ECP, 0 END_BAL" & vbCr _
        & " from ARTPYMT1,ARTPYMT2 " & vbCr _
        & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO " & vbCr _
        & " and ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO " & vbCr _
        & " and ARTPYMT1.OPS_YYYYPP = '" & RYP & "'" & vbCr _
        & " and ARTPYMT2.PYMT_STATUS = '2'" & vbCr _
        & " and NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" & vbCr _
        & " group by ARTPYMT2.CUST_CODE" & vbCr _
        & " union " & vbCr _
        & "Select ARTPYMT2.CUST_CODE, 0 BEG_BAL, 0 INV, 0 RTN, 0 CRM " & vbCr _
        & ", 0 PMT, SUM (NVL(ARTPYMT3.INV_DISC_TAKEN,0)+NVL(ARTPYMT3.INV_WRITE_OFF,0)) DED, 0 ECP, 0 END_BAL" & vbCr _
        & " from ARTPYMT1,ARTPYMT2,ARTPYMT3 " & vbCr _
        & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO " & vbCr _
        & " and ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO " & vbCr _
        & " and ARTPYMT3.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO " & vbCr _
        & " and ARTPYMT3.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO " & vbCr _
        & " and ARTPYMT1.OPS_YYYYPP = '" & RYP & "'" & vbCr _
        & " and ARTPYMT2.PYMT_STATUS = '2'" & vbCr _
        & " and NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" & vbCr _
        & " and NVL(ARTPYMT3.INV_DISC_TAKEN,0)+NVL(ARTPYMT3.INV_WRITE_OFF,0) <> 0" & vbCr _
        & " group by ARTPYMT2.CUST_CODE" & vbCr _
        & " union " & vbCr _
        & "Select ARTPYMT2.CUST_CODE, 0 BEG_BAL, 0 INV, 0 RTN, 0 CRM " & vbCr _
        & ", 0 PMT, SUM (NVL(ARTPYMT5.GL_DIST_AMT,0)) DED, 0 ECP, 0 END_BAL" & vbCr _
        & " from ARTPYMT1,ARTPYMT2,ARTPYMT5 " & vbCr _
        & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO " & vbCr _
        & " and ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO " & vbCr _
        & " and ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO " & vbCr _
        & " and ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO " & vbCr _
        & " and ARTPYMT1.OPS_YYYYPP = '" & RYP & "'" & vbCr _
        & " and ARTPYMT2.PYMT_STATUS = '2'" & vbCr _
        & " and NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" & vbCr _
        & " and NVL(ARTPYMT5.GL_DIST_AMT,0) <> 0" & vbCr _
        & " and NVL(ARTPYMT5.CHARGEBACK_IND,'0') <> '1'" & vbCr _
        & " group by ARTPYMT2.CUST_CODE" & vbCr _
        & " union " & vbCr

        If RYP = ASCMAIN1.CYP Then
            ASCMAIN1.sql &= "" _
            & "Select CUST_CODE, 0 BEG_BAL" & vbCr _
            & ", 0 INV, 0 RTN, 0 CRM, 0 PMT, 0 DED, 0 ECP" & vbCr _
            & ", SUM (INV_BALANCE) END_BAL" & vbCr _
            & " from ARTOPEN1" & vbCr _
            & " group by CUST_CODE " & vbCr
        Else
            ASCMAIN1.sql &= "" _
            & "Select CUST_CODE, 0 BEG_BAL" & vbCr _
            & ", 0 INV, 0 RTN, 0 CRM, 0 PMT, 0 DED, 0 ECP" & vbCr _
            & ", SUM (TOTAL_DUE) END_BAL" & vbCr _
            & " from ARTSTMT1 WHERE OPS_YYYYPP = '" & RYP & "'" & vbCr _
            & " group by CUST_CODE " & vbCr
        End If

        ASCMAIN1.sql &= "" _
        & ") X, ARTCUST1 where ARTCUST1.CUST_CODE = X.CUST_CODE" & vbCr _
        & " group by NVL(X.CUST_CODE,'.')" & vbCr _
        & ", ARTCUST1.CUST_NAME, ARTCUST1.POST_CODE" & vbCr

        ASCDATA1.ExecuteSQL()

        Fill_Records("ARTGLARC")
        Sort_grdColumns(grdARTGLARC, "CUST_CODE")

    End Sub

    Private Sub grdARTGLARC_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTGLARC.InitializeRow

        Dim END_BAL As Decimal = Val(e.Row.Cells("END_BAL").Value & "")

        Dim END_BAL_CALC As Decimal _
        = Val(e.Row.Cells("BEG_BAL").Value & "") _
        + Val(e.Row.Cells("INV").Value & "") _
        + Val(e.Row.Cells("RTN").Value & "") _
        + Val(e.Row.Cells("CRM").Value & "") _
        - Val(e.Row.Cells("PMT").Value & "") _
        - Val(e.Row.Cells("DED").Value & "") _
        + Val(e.Row.Cells("ECP").Value & "")

        If END_BAL <> END_BAL_CALC Then
            If System.Math.Round(END_BAL, 2) <> System.Math.Round(END_BAL_CALC, 2) Then
                e.Row.Cells("END_BAL").Appearance.ForeColor = Drawing.Color.Red
            Else
                'e.Row.Cells("END_BAL").Appearance.ForeColor = Drawing.Color.Green
            End If
            'e.Row.Cells("END_BAL").Appearance.ForeColor = Drawing.Color.Red
        End If

    End Sub

    Private Sub optD_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optD.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If optD.Value = "DateBy" Then
            grdARTGLAR1.Parent = SplitContainer1.Panel1
            grdARTGLAR2.Parent = SplitContainer1.Panel2
            Setup_ARTGLAR1(True)
        Else
            grdARTGLAR1.Parent = SplitContainer1.Panel2
            grdARTGLAR2.Parent = SplitContainer1.Panel1
            Setup_ARTGLAR2(True)
        End If
    End Sub

    Private Sub grdARTGLAR2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTGLAR2.AfterRowActivate
        Setup_ARTGLAR1()
    End Sub

    Sub Setup_optD()
        For Each vli As ValueListItem In optD.ValueList.ValueListItems
            If vli.DataValue = "DateBy" Then
                vli.DisplayText = "Date By " & optDateBy.Text
            Else
                vli.DisplayText = optDateBy.Text & " By Date"
            End If
        Next
    End Sub

    Private Sub SetUpMiscCharges()

        Dim fieldName As String = IIf(optChoice.Value = "P", "INV_DATE", "ORDR_YYYYPP_UPDATED")
        Dim selectionValue As String = IIf(optChoice.Value = "P", MyBase.Absx1.txtFor("OPS_YYYYPP").Text.Trim, MyBase.Absx1.txtFor("OPS_YYYY").Text.Trim)

        Dim dts As Integer = 0
        Dim sqlMisc As String = String.Empty
        Dim sqlSample As String = String.Empty

        Dim year As String = String.Empty
        Dim queryInvDate As Date = DateTime.Now
        Dim queryInvDateEnd As Date = DateTime.Now

        ' Get the number of columns to display
        If optChoice.Value = "P" Then
            dts = Date.DaysInMonth(selectionValue.Substring(0, 4), selectionValue.Substring(4, 2))
            year = selectionValue.Substring(0, 4)
            Dim previousPeriod As String = ASCMAIN1.Period_Calc(selectionValue, -1)
            queryInvDate = ASCDATA1.GetDataValue("Select PRD_END_DATE From GLTPARM2 WHERE OPS_YYYYPP = '" & previousPeriod & "'")
            queryInvDateEnd = ASCDATA1.GetDataValue("Select PRD_END_DATE From GLTPARM2 WHERE OPS_YYYYPP = '" & selectionValue & "'")
        Else
            dts = 12
            year = selectionValue
            Fill_Records("GLTPARM2", "", True, "SELECT * FROM GLTPARM2 WHERE SUBSTR(OPS_YYYYPP,1,4) = '" & year & "'")
            If ASCMAIN1.CYP.StartsWith(year) Then
                dts = ASCMAIN1.CYP.Substring(4, 2)
            End If
        End If

        sqlMisc &= "Select SOTMISC1.ACCT_CODE, SOTMISC1.MISC_CHG_DESC, SOTINVH1.MISC_CHG_CODE"
        sqlSample &= "Select SOTMISC1.ACCT_CODE, SOTMISC1.MISC_CHG_DESC, 'SSC' MISC_CHG_CODE"

        ' Column names
        For iCol As Integer = 1 To periodDays
            If optChoice.Value = "P" Then
                queryInvDate = DateAdd(DateInterval.Day, 1, queryInvDate)

                If queryInvDate <= queryInvDateEnd Then
                    grdSOTMISCX.DisplayLayout.Bands(0).Columns("COL_" & iCol).Header.Caption = queryInvDate.ToString("MM/dd/yyyy")
                    sqlMisc &= ",SUM (CASE WHEN " & fieldName & " = '" & queryInvDate.ToString("dd-MMM-yyyy") & "' THEN NVL(INV_MISC_CHG,0) ELSE 0 END) COL_" & iCol
                    sqlSample &= ",SUM (CASE WHEN " & fieldName & " = '" & queryInvDate.ToString("dd-MMM-yyyy") & "' THEN NVL(INV_SAMPLE_SURCHARGE,0) ELSE 0 END) COL_" & iCol
                    grdSOTMISCX.DisplayLayout.Bands(0).Columns("COL_" & iCol).Hidden = False
                Else
                    grdSOTMISCX.DisplayLayout.Bands(0).Columns("COL_" & iCol).Header.Caption = queryInvDate.ToString("MM/dd/yyyy")
                    sqlMisc &= ",SUM (0) COL_" & iCol
                    sqlSample &= ",SUM (0) COL_" & iCol
                    grdSOTMISCX.DisplayLayout.Bands(0).Columns("COL_" & iCol).Hidden = True
                End If

                grdSOTMISCX.DisplayLayout.Bands(0).Columns("COL_" & iCol).Width = 100
            Else
                If dst.Tables("GLTPARM2").Select("OPS_YYYYPP = '" & year & iCol.ToString("00") & "'").Length > 0 Then
                    grdSOTMISCX.DisplayLayout.Bands(0).Columns("COL_" & iCol).Header.Caption = dst.Tables("GLTPARM2").Select("OPS_YYYYPP = '" & year & iCol.ToString("00") & "'")(0).Item("LEGEND") & String.Empty
                Else
                    grdSOTMISCX.DisplayLayout.Bands(0).Columns("COL_" & iCol).Header.Caption = String.Empty
                End If

                If iCol <= dts Then
                    grdSOTMISCX.DisplayLayout.Bands(0).Columns("COL_" & iCol).Hidden = False
                    sqlMisc &= ",SUM (CASE WHEN " & fieldName & " = '" & year & iCol.ToString("00") & "' THEN NVL(INV_MISC_CHG,0) ELSE 0 END) COL_" & iCol
                    sqlSample &= ",SUM (CASE WHEN " & fieldName & " = '" & year & iCol.ToString("00") & "' THEN NVL(INV_SAMPLE_SURCHARGE,0) ELSE 0 END) COL_" & iCol
                Else
                    grdSOTMISCX.DisplayLayout.Bands(0).Columns("COL_" & iCol).Hidden = True
                    sqlMisc &= ",SUM (0) COL_" & iCol
                    sqlSample &= ",SUM (0) COL_" & iCol

                End If

                grdSOTMISCX.DisplayLayout.Bands(0).Columns("COL_" & iCol).Width = 140
            End If

            If iCol Mod 2 = 0 Then
                grdSOTMISCX.DisplayLayout.Bands(0).Columns("COL_" & iCol).CellAppearance.BackColor = Drawing.Color.LightYellow
            Else
                grdSOTMISCX.DisplayLayout.Bands(0).Columns("COL_" & iCol).CellAppearance.BackColor = Drawing.Color.LightSkyBlue
            End If

            grdSOTMISCX.DisplayLayout.Bands(0).Columns("COL_" & iCol).Format = "#,##0.00"
            grdSOTMISCX.DisplayLayout.Bands(0).Columns("COL_" & iCol).CellAppearance.TextHAlign = HAlign.Right
        Next

        grdSOTMISCX.DisplayLayout.Bands(0).Columns("COL_TOTAL").Header.Caption = "Totals"
        grdSOTMISCX.DisplayLayout.Bands(0).Columns("COL_TOTAL").Width = 100
        grdSOTMISCX.DisplayLayout.Bands(0).Columns("COL_TOTAL").Format = "#,##0.00"
        grdSOTMISCX.DisplayLayout.Bands(0).Columns("COL_TOTAL").Hidden = False
        grdSOTMISCX.DisplayLayout.Bands(0).Columns("COL_TOTAL").CellAppearance.BackColor = Drawing.Color.Coral

        grdSOTMISCX.DisplayLayout.UseFixedHeaders = True
        With grdSOTMISCX.DisplayLayout.Bands(0)
            .Columns("ACCT_CODE").Header.Fixed = True
            .Columns("MISC_CHG_DESC").Header.Fixed = True
            .Columns("MISC_CHG_CODE").Header.Fixed = True
        End With

        sqlMisc &= " FROM SOTINVH1,SOTMISC1"
        If optChoice.Value = "P" Then
            sqlMisc &= " WHERE ORDR_YYYYPP_UPDATED = '" & selectionValue & "'"
        Else
            sqlMisc &= " WHERE ORDR_YYYYPP_UPDATED BETWEEN '" & year & "01' AND '" & year & dts.ToString("00") & "'"
        End If
        sqlMisc &= " AND SOTINVH1.MISC_CHG_CODE IS NOT NULL"
        sqlMisc &= " AND SOTINVH1.MISC_CHG_CODE =SOTMISC1.MISC_CHG_CODE"
        sqlMisc &= " GROUP BY SOTMISC1.ACCT_CODE, SOTMISC1.MISC_CHG_DESC, SOTINVH1.MISC_CHG_CODE"


        sqlSample &= " FROM SOTINVH1,SOTMISC1"
        If optChoice.Value = "P" Then
            sqlSample &= " WHERE ORDR_YYYYPP_UPDATED = '" & selectionValue & "'"
        Else
            sqlSample &= " WHERE ORDR_YYYYPP_UPDATED BETWEEN '" & year & "01' AND '" & year & dts.ToString("00") & "'"
        End If
        sqlSample &= " AND SOTMISC1.MISC_CHG_CODE = 'SSC'"
        sqlSample &= " AND NVL(INV_SAMPLE_SURCHARGE,0) <> 0"
        sqlSample &= " GROUP BY SOTMISC1.ACCT_CODE, SOTMISC1.MISC_CHG_DESC)"

        Dim sql As String = "select * from ( "
        sql &= sqlMisc
        sql &= " Union "
        sql &= sqlSample

        Fill_Records("SOTMISCX", String.Empty, True, sql)

        Clear_All_Filters(grdSOTMISCX)
        Show_Filter(grdSOTMISCX, False)
        grdSOTMISCX.DisplayLayout.GroupByBox.Hidden = True

        Sort_grdColumns(grdSOTMISCX, "ACCT_CODE")
    End Sub

#Region "Format grds"
    Sub Format_grdARTGLAR1_2()

        Call Create_Summary(grdARTGLAR1, "INV_DATE", "Count")
        Call Create_Summary(grdARTGLAR1, "INVS")
        Call Create_Summary(grdARTGLAR1, "CRMS")
        Call Create_Summary(grdARTGLAR1, "SLS")
        Call Create_Summary(grdARTGLAR1, "CGS")
        Call Create_Summary(grdARTGLAR1, "FRT")
        Call Create_Summary(grdARTGLAR1, "TAX")
        Call Create_Summary(grdARTGLAR1, "MSC")
        Call Create_Summary(grdARTGLAR1, "NET")
        Call Create_Summary(grdARTGLAR1, "ISLS")
        Call Create_Summary(grdARTGLAR1, "INET")
        Call Create_Summary(grdARTGLAR1, "CSLS")
        Call Create_Summary(grdARTGLAR1, "CNET")

        Call Create_Summary(grdARTGLAR2, "CODE_VALUE", "Count")
        Call Create_Summary(grdARTGLAR2, "INVS")
        Call Create_Summary(grdARTGLAR2, "CRMS")
        Call Create_Summary(grdARTGLAR2, "SLS")
        Call Create_Summary(grdARTGLAR2, "CGS")
        Call Create_Summary(grdARTGLAR2, "FRT")
        Call Create_Summary(grdARTGLAR2, "TAX")
        Call Create_Summary(grdARTGLAR2, "MSC")
        Call Create_Summary(grdARTGLAR2, "NET")
        Call Create_Summary(grdARTGLAR2, "ISLS")
        Call Create_Summary(grdARTGLAR2, "INET")
        Call Create_Summary(grdARTGLAR2, "CSLS")
        Call Create_Summary(grdARTGLAR2, "CNET")

        grdARTGLAR1.DisplayLayout.UseFixedHeaders = True
        With grdARTGLAR1.DisplayLayout.Bands(0)
            .Columns("INV_DATE").Header.Fixed = True
        End With

        grdARTGLAR2.DisplayLayout.UseFixedHeaders = True
        With grdARTGLAR2.DisplayLayout.Bands(0)
            .Columns("CODE_VALUE").Header.Fixed = True
        End With

        grdARTGLARR.DisplayLayout.UseFixedHeaders = True
        With grdARTGLARR.DisplayLayout.Bands(0)
            .Columns("REGISTER_DATE").Header.Fixed = True
        End With

        With grdARTGLAR1.DisplayLayout.Bands(0).Columns("INV_DATE")
            .Width = 100
            .CellAppearance.BackColor = Drawing.Color.Beige
            .Header.Caption = "Inv Date"
        End With

        grdARTGLAR1.DisplayLayout.Bands(0).Columns("CODE2_VALUE").Header.Caption = "Week"
        grdARTGLAR1.DisplayLayout.Bands(0).Columns("DESC_VALUE").Header.Caption = "Day of Week"

        With grdARTGLAR2.DisplayLayout.Bands(0).Columns("CODE_VALUE")
            .Width = 100
            .CellAppearance.BackColor = Drawing.Color.Beige
        End With

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdARTGLAR1, grdARTGLAR2}
            With grd.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"DESC_VALUE", "CODE2_VALUE"}
                    .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
                Next

                .Columns("AVG_INV").Format = "#,##0"

                .Columns("DESC_VALUE").Width = 150
                .Columns("CODE2_VALUE").Width = 80

                .Columns("CRMS").Width = 70
                .Columns("INVS").Width = 70
                .Columns("AVG_INV").Width = 70
                .Columns("AVG_INV").Header.Caption = "AvgInv"

                For Each COLUMN_NAME As String In New String() {"ISLS", "INET", "CSLS", "CNET"}
                    If COLUMN_NAME.StartsWith("I") Then
                        .Columns(COLUMN_NAME).CellAppearance.ForeColor = Drawing.Color.Green
                    Else
                        .Columns(COLUMN_NAME).CellAppearance.ForeColor = Drawing.Color.Red
                    End If
                    .Columns(COLUMN_NAME).Width = 110
                Next

                For Each COLUMN_NAME As String In New String() {"SLS", "CGS", "FRT", "MSC", "TAX", "NET"}
                    .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightBlue
                    If COLUMN_NAME = "FRT" Or COLUMN_NAME = "MSC" Then
                        .Columns(COLUMN_NAME).Width = 90
                    Else
                        .Columns(COLUMN_NAME).Width = 110
                    End If
                Next

            End With
        Next
    End Sub

    Sub Format_grdARTGLARC()

        Call Create_Summary(grdARTGLARC, "CUST_CODE", "Count")
        Call Create_Summary(grdARTGLARC, "BEG_BAL")
        Call Create_Summary(grdARTGLARC, "INV")
        Call Create_Summary(grdARTGLARC, "RTN")
        Call Create_Summary(grdARTGLARC, "CRM")
        Call Create_Summary(grdARTGLARC, "PMT")
        Call Create_Summary(grdARTGLARC, "DED")
        Call Create_Summary(grdARTGLARC, "ECP")
        Call Create_Summary(grdARTGLARC, "END_BAL")
        Call Create_Summary(grdARTGLARC, "OOBAL")

        grdARTGLARC.DisplayLayout.UseFixedHeaders = True
        With grdARTGLARC.DisplayLayout.Bands(0)
            .Columns("CUST_CODE").Header.Fixed = True
        End With

        With grdARTGLARC.DisplayLayout.Bands(0)
            With .Columns("CUST_CODE")
                .Width = 100
                .CellAppearance.BackColor = Drawing.Color.Beige
                .Header.Caption = "Customer"
            End With
            With .Columns("CUST_NAME")
                .Width = 150
                .CellAppearance.BackColor = Drawing.Color.Beige
                .Header.Caption = "Name"
            End With
            With .Columns("POST_CODE")
                .Width = 70
                .CellAppearance.BackColor = Drawing.Color.Beige
                .Header.Caption = "Code"
            End With
        End With

        With grdARTGLARC.DisplayLayout.Bands(0)
            .Columns("BEG_BAL").Width = 110
            .Columns("BEG_BAL").Header.Caption = "Beg Bal"
            With .Columns("INV")
                .Width = 100
                .Header.Caption = "INVs"
                .CellAppearance.BackColor = Drawing.Color.LightBlue
            End With
            With .Columns("RTN")
                .Width = 80
                .Header.Caption = "RTNs"
                .CellAppearance.BackColor = Drawing.Color.LightBlue
            End With
            With .Columns("CRM")
                .Width = 80
                .Header.Caption = "CRMs"
                .CellAppearance.BackColor = Drawing.Color.LightBlue
            End With
            With .Columns("PMT")
                .Width = 100
                .Header.Caption = "PMTs"
                .CellAppearance.BackColor = Drawing.Color.LightGreen
            End With
            With .Columns("DED")
                .Width = 80
                .Header.Caption = "DEDs"
                .CellAppearance.BackColor = Drawing.Color.LightGreen
            End With
            With .Columns("ECP")
                .Width = 80
                .Header.Caption = "ECP"
                .CellAppearance.BackColor = Drawing.Color.LightYellow
            End With
            With .Columns("END_BAL")
                .Width = 110
                .Header.Caption = "End Bal"
            End With

        End With
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
                    ElseIf GC.Key Like "TOT*" Or GC.Key Like "ECP*" Or GC.Key Like "REG*" Or GC.Key Like "B2C*" Then
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
                                Case "REG"
                                    GC.Header.Caption = "Trade"
                                Case "B2C"
                                    GC.Header.Caption = "B2C"
                                Case "ECP"
                                    GC.Header.Caption = "ECP"
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
                                    Case Else
                                        GC.Header.Caption = Mid(GC.Key, 5)
                                End Select
                            End If
                        End If
                    End If

                    GC.Width = 80

                    If GC.Key = "REG" Or GC.Key = "B2C" Or GC.Key = "ECP" Or GC.Key = "NRSF" Or GC.Key = "TOT" Or GC.Key = "GL" Or GC.Key = "RF_GL" Or GC.Key = "RF_AR" Then
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


    Sub Format_grdARTGLARS()

        grdARTGLARS.DisplayLayout.UseFixedHeaders = True
        With grdARTGLARS.DisplayLayout.Bands(0)
            .Columns("INV_DATE").Header.Fixed = True
        End With

        With grdARTGLARS.DisplayLayout.Bands(0)
            With .Columns("INV_DATE")
                .Width = 100
                .CellAppearance.BackColor = Drawing.Color.Beige
            End With

            For Each COLUMN_NAME As String In New String() _
            {"INV_SLS", "INV_FRT", "INV_TAX", "GROSS"}
                With .Columns(COLUMN_NAME)
                    .CellAppearance.BackColor = Drawing.Color.LightBlue
                End With
            Next

            For Each COLUMN_NAME As String In New String() _
            {"CREDITS", "MISC", "RSF"}
                With .Columns(COLUMN_NAME)
                    .CellAppearance.BackColor = Drawing.Color.LightYellow
                End With
            Next

            For Each COLUMN_NAME As String In New String() _
            {"NET_AR", "INVS", "AVG"}
                With .Columns(COLUMN_NAME)
                    .CellAppearance.BackColor = Drawing.Color.LightSeaGreen
                End With
            Next

            For Each COLUMN_NAME As String In New String() _
            {"CRM_SLS", "CRM_FRT", "CRM_TAX"}
                With .Columns(COLUMN_NAME)
                    .CellAppearance.BackColor = Drawing.Color.LightPink
                End With
            Next

            For Each COLUMN_NAME As String In New String() _
            {"WEB_SLS", "WEB_FRT", "WEB_TAX", "WEB_FEE", "WEB_NET"}
                With .Columns(COLUMN_NAME)
                    .CellAppearance.BackColor = Drawing.Color.Orange ' LightSalmon 
                End With
            Next

            For Each COLUMN_NAME As String In New String() _
            {"DEL_SLS", "DEL_EDG", "DEL_ARC", "DEL_LTR", "DEL_DSC", "DEL_NET"}
                With .Columns(COLUMN_NAME)
                    .CellAppearance.BackColor = Drawing.Color.AliceBlue
                End With
            Next

            For Each DC As UltraWinGrid.UltraGridColumn In .Columns
                If DC.Key = "INV_DATE" Then
                    Call Create_Summary(grdARTGLARS, DC.Key, "Count")
                ElseIf DC.Key = "WEEK" Or DC.Key = "LINE" Then
                    DC.Format = "###,##0"
                ElseIf DC.Key = "DIVISION" Then
                ElseIf DC.Key = "AVG" Then
                    'Setup_Calc()
                    Call Create_Summary(grdARTGLARS, DC.Key, "Custom", , "##0.00")
                Else
                    DC.Format = "###,##0.00"
                    If DC.Key = "INVS" Then
                        DC.Format = "###,##0"
                    End If
                    Call Create_Summary(grdARTGLARS, DC.Key)
                End If
            Next
        End With
    End Sub

#End Region

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
            Dim ECP As Decimal = Val(e.Row.Cells("ECP").Value & "")
            If SLS + FRT + TAX + MSC <> TOT - ECP Then
                e.Row.Cells("TOT").Appearance.ForeColor = Drawing.Color.Red
            End If
            e.Row.Cells("ECP").Appearance.ForeColor = Drawing.Color.Blue
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
    End Sub

    Private Sub optDIVISION_CODE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optDIVISION_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub

        Setup_ARTGLARS(True)
    End Sub

    Private Sub cmdMagicButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMagicButton.Click
        Me.Cursor = Cursors.WaitCursor
        Dim myWorkbook As New Infragistics.Documents.Excel.Workbook

        For Each gcol As UltraWinGrid.UltraGridColumn In grdARTGLARS.DisplayLayout.Bands(0).Columns
            gcol.Header.Appearance.BackColor = Drawing.Color.Blue ' .DodgerBlue
            gcol.Header.Appearance.ForeColor = Drawing.Color.White
        Next

        If optDIVISION_CODE.Value = "*" Then
            Setup_ARTGLARS(True)
        Else
            optDIVISION_CODE.Value = "*"
        End If
        Call Export_to_Excel_Add_grd(myWorkbook, grdARTGLARS, False, "COMBINED", , "Sales Analysis")
        optDIVISION_CODE.Value = "ODG"
        Call Export_to_Excel_Add_grd(myWorkbook, grdARTGLARS, False, "ODG", , "Sales Analysis")
        optDIVISION_CODE.Value = "DEL"
        Call Export_to_Excel_Add_grd(myWorkbook, grdARTGLARS, False, "DEL", , "Sales Analysis")

        Call Export_to_Excel_Show(myWorkbook, Me.Text)

        optDIVISION_CODE.Value = "*"

        For Each gcol As UltraWinGrid.UltraGridColumn In grdARTGLARS.DisplayLayout.Bands(0).Columns
            gcol.Header.Appearance.BackColor = Drawing.Color.Empty
            gcol.Header.Appearance.ForeColor = Drawing.Color.Empty
        Next

        Call ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub grdARTGLARS_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTGLARS.InitializeRow

        If Not e.Row.IsDataRow Then
            Stop
        End If
        If e.Row.IsSummaryRow Then
            If Val(e.Row.Cells("INVS").Value & "") = 0 Then
                e.Row.Cells("AVG").Value = 0
            Else
                e.Row.Cells("AVG").Value = Val(e.Row.Cells("GROSS").Value & "") / Val(e.Row.Cells("INVS").Value & "")
            End If
        End If
    End Sub

    Private Sub chkDollars_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDollars.CheckedChanged
        Dim FORMAT As String
        If chkDollars.Checked Then
            FORMAT = "###,##0"
        Else
            FORMAT = "###,##0.00"
        End If
        For Each GC As UltraWinGrid.UltraGridColumn In grdARTGLARS.DisplayLayout.Bands(0).Columns
            Select Case GC.Key
                Case "INV_SLS", "INV_FRT", "INV_TAX", "GROSS", "CREDITS", "MISC", "RSF", "NET_AR", "CRM_SLS", "CRM_FRT", "CRM_TAX", "WEB_SLS", "WEB_FRT", "WEB_TAX", "WEB_FEE", "WEB_NET"
                    GC.Format = FORMAT
                    GC.Band.Summaries(GC.Key).DisplayFormat = "{0:" & FORMAT & "}"
                    GC.Format = FORMAT
            End Select
        Next
    End Sub

    Overrides Sub CustomSummary_DataRows( _
    ByVal summarySettings As UltraWinGrid.SummarySettings, _
    ByVal row As UltraWinGrid.UltraGridRow, _
    ByRef CustomValue As Double, _
    ByVal grd As UltraWinGrid.UltraGrid)

        Select Case grd.Name
            Case "grdARTGLARS"
                If summarySettings.Key = "AVG" Then
                    'Dim GROSS As Object = row.GetCellValue(summarySettings.SourceColumn.Band.Columns("GROSS"))
                End If
        End Select
    End Sub

    Overrides Function CustomSummary_End( _
    ByVal summarySettings As UltraWinGrid.SummarySettings, _
    ByVal rows As UltraWinGrid.RowsCollection, _
    ByVal CustomValue As Double, _
    ByVal grd As UltraWinGrid.UltraGrid) As Double

        Select Case grd.Name
            Case "grdARTGLARS"
                If summarySettings.Key = "AVG" Then
                    Dim COLUMN_NAME As String = summarySettings.SourceColumn.Key
                    Dim INVS As Decimal = Val(rows.SummaryValues("INVS").Value & "")
                    Dim GROSS As Decimal = Val(rows.SummaryValues("GROSS").Value & "")
                    Dim AVG As Decimal = 0
                    If INVS <> 0 Then
                        AVG = GROSS / INVS
                    End If
                    Return AVG
                End If
        End Select
    End Function

    Private Sub UltraOptionSet2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optChoice.ValueChanged

        MyBase.Absx1.txtFor("OPS_YYYYPP").Visible = optChoice.Value = "P"
        MyBase.Absx1.txtFor("LEGEND").Visible = optChoice.Value = "P"
        MyBase.Absx1.txtFor("OPS_YYYY").Visible = optChoice.Value = "Y"
    End Sub

    Private Sub chkPymtApplyOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPymtApplyOnly.CheckedChanged
        If chkPymtApplyOnly.Checked = True Then
            viewARTPYMT1X.RowFilter = "ISNULL(PYMT_APPL_ONLY, '0') = '1'"
        Else
            viewARTPYMT1X.RowFilter = String.Empty
        End If
    End Sub

    Private Sub grdARTPYMT1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdARTPYMT1.InitializeLayout
        grdARTPYMT1.DisplayLayout.Bands(1).SummaryFooterCaption = "Summaries for Batch No: [SCROLLTIPFIELD]"
    End Sub

End Class