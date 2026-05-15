Imports System.Drawing
Imports System.Math
Imports Infragistics.Win.UltraWinGrid

Public Class ICFSTAT1
    Dim ICTSTATX As String
    Dim ICTTRANX As String
    Dim rowICTITEM1 As DataRow
    Dim RYP As String = ASCMAIN1.CYP
    Dim ITEM_CODE As String

    Dim ICTSTATI As String
    Dim sqlICTSTATI As String
    Dim ICTSTATP As String
    Dim sqlICTSTATP As String
    Dim PO_PARM_PINV_LT As Integer
    Dim PO_PARM_PINV_PORT As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")
        Get_PARM("SOTPARM1")
        Get_PARM("POTPARM1")
        splGridContainer.Panel1Collapsed = False
        splGridContainer.Panel2Collapsed = True

        PO_PARM_PINV_LT = Val(ROWs("POTPARM1").Item("PO_PARM_PINV_LT") & "")
        PO_PARM_PINV_PORT = ROWs("POTPARM1").Item("PO_PARM_PINV_PORT") & ""

        With dst

            ASCMAIN1.sql = "SELECT ICTSTAT1.*, " _
            & " WHSE_QTY_ON_HAND," _
            & " WHSE_QTY_ONPO," _
            & " WHSE_QTY_PLAN," _
            & " WHSE_QTY_OPEN," _
            & " WHSE_QTY_PICK," _
            & " WHSE_QTY_COMM," _
            & " WHSE_QTY_HOLD" _
            & " FROM ICTSTAT1,ICTSTAT2 WHERE ROWNUM < 1"
            ICTSTATX = ASCMAIN1.Temp_Table

            ASCMAIN1.sql = "Select ICTSTATX.* from " & ICTSTATX & " ICTSTATX"
            Create_TDA(.Tables.Add, "ICTSTATX", "**", 0, False, "", 3)
            .Tables("ICTSTATX").Columns.Add("WHSE_QTY_NETA", GetType(System.Int64), "ISNULL(WHSE_QTY_ON_HAND,0) - ISNULL(WHSE_QTY_PICK,0)")
            .Tables("ICTSTATX").Columns.Add("WHSE_QTY_ATS", GetType(System.Int64), "ISNULL(WHSE_QTY_NETA,0) - ISNULL(WHSE_QTY_OPEN,0) - ISNULL(WHSE_QTY_COMM,0)")

            ASCMAIN1.sql = "Select ITEM_CODE, WHSE_CODE" & vbCrLf _
                & ", WHSE_QTY_ON_HAND PO_DTL" & vbCrLf _
                & ", WHSE_QTY_ON_HAND PO_SUM" & vbCrLf _
                & ", WHSE_QTY_ON_HAND PP_DTL" & vbCrLf _
                & ", WHSE_QTY_ON_HAND PP_SUM" & vbCrLf _
                & ", WHSE_QTY_ON_HAND SO_DTL" & vbCrLf _
                & ", WHSE_QTY_ON_HAND SO_SUM" & vbCrLf _
                & ", WHSE_QTY_ON_HAND SP_DTL" & vbCrLf _
                & ", WHSE_QTY_ON_HAND SP_SUM" & vbCrLf _
                & ", WHSE_QTY_ON_HAND PC_DTL" & vbCrLf _
                & ", WHSE_QTY_ON_HAND PC_SUM" & vbCrLf _
                & " from ICTSTAT2"
            Create_TDA(.Tables.Add, "ICTSTATO", "**", 0, False, "", 2)

            'Create_TDA(.Tables.Add, "ICTITEM1", "*", -1, False)

            ASCMAIN1.sql = "Select ICTITEM1.*" & vbCrLf _
                & " from ICTITEM1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = :PARM1"
            For Each TABLE_NAME As String In New String() {"ICTITEM1", "ICTITEM1_RECENT", "ICTITEM1_VIEW"}
                Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, False, "V", IIf(TABLE_NAME = "ICTITEM1_RECENT", 1, 0))
                .Tables(TABLE_NAME).Columns.Add("QTY_ONHD", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("QTY_ONPO", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("QTY_PLAN", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("QTY_OPEN", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("QTY_PICK", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("QTY_COMM", GetType(System.Int64))
                '.Tables(TABLE_NAME).Columns.Add("QTY_NETA", GetType(System.Int64), "ISNULL(QTY_ONHD,0) + ISNULL(QTY_ONPO,0) + ISNULL(QTY_PLAN,0) - ISNULL(QTY_OPEN,0) - ISNULL(QTY_PICK,0) - ISNULL(QTY_COMM,0)")
                .Tables(TABLE_NAME).Columns.Add("QTY_NETA", GetType(System.Int64), "ISNULL(QTY_ONHD,0) - ISNULL(QTY_OPEN,0) - ISNULL(QTY_PICK,0) - ISNULL(QTY_COMM,0)")
                .Tables(TABLE_NAME).Columns.Add("MTD_SHP", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("ITEM_FC", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("ITEM_FC_MOS", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("ITEM_POS", GetType(System.Decimal))
                .Tables(TABLE_NAME).Columns.Add("ITEM_FC0", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("ITEM_FC1", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("ITEM_FC2", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("ITEM_FC3", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("ITEM_POS0", GetType(System.Decimal))
                .Tables(TABLE_NAME).Columns.Add("ITEM_POS1", GetType(System.Decimal))
                .Tables(TABLE_NAME).Columns.Add("ITEM_POS2", GetType(System.Decimal))
                .Tables(TABLE_NAME).Columns.Add("ITEM_POS3", GetType(System.Decimal))
                .Tables(TABLE_NAME).Columns.Add("IMAGE", GetType(System.Byte()))
                .Tables(TABLE_NAME).Columns.Add("EXT_COST", GetType(System.Decimal), "ISNULL(QTY_ONHD,0) * ISNULL(ITEM_COST_STD,0)")
            Next

            ASCMAIN1.sql = "Select * from ICTCOLL1"
            Create_TDA(.Tables.Add, "ICTCOLL1", "**", 0, False)
            ASCMAIN1.sql = "Select * from ICTCATG1"
            Create_TDA(.Tables.Add, "ICTCATG1", "**", 0, False)
            ASCMAIN1.sql = "Select * from ICTDEPT1"
            Create_TDA(.Tables.Add, "ICTDEPT1", "**", 0, False)


            ASCMAIN1.sql = "SELECT ICTIADJ2.OPS_YYYYPP, ICTIADJ2.ITEM_CODE, ICTIADJ1.WHSE_CODE" _
            & ", ICTIADJ1.ADJ_NO TRAN_NO, ICTIADJ1.ADJ_SOURCE TRAN_SOURCE" _
            & ", ICTIADJ1.INIT_DATE, ICTIADJ1.INIT_OPER" _
            & ", ICTIADJ1.ADJ_DATE TRAN_DATE, 'X' TRAN_TYPE" _
            & ", ICTIADJ2.ADJ_QTY TRAN_QTY, ICTIADJ1.ADJ_NOTE TRAN_NOTE" _
            & " FROM ICTIADJ1,ICTIADJ2 WHERE ROWNUM < 1"
            ICTTRANX = ASCMAIN1.Temp_Table

            ASCMAIN1.sql = "Select ICTTRANX.* from " & ICTTRANX & " ICTTRANX"
            Create_TDA(.Tables.Add, "ICTTRANX", "**", 0, False)
            With .Tables("ICTTRANX").Columns
                For Each C As String In New String() {"S", "R", "P", "A", "T", "X", "V"}
                    .Add("TRAN_QTY_" & C, GetType(System.Int64), "IIF(TRAN_TYPE='" & C & "',TRAN_QTY,0)")
                Next
            End With

            ASCMAIN1.sql = "Select SOTALLO1.* from SOTALLO1 where ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SOTALLO1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTALLO2.* from SOTALLO2 where ALLO_CTL_NO in (Select ALLO_CTL_NO from SOTALLO1 where ITEM_CODE = :PARM1)"
            Create_TDA(.Tables.Add, "SOTALLO2", "**", 0, False, "V", 2)
            With .Tables("SOTALLO2").Columns
                .Add("ORDR_QTY", GetType(System.Int64))
                .Add("ORDR_QTY_OPEN", GetType(System.Int64))
                .Add("ORDR_QTY_PICK", GetType(System.Int64))
                .Add("ORDR_QTY_SHIP", GetType(System.Int64))
                .Add("ORDR_QTY_CANC", GetType(System.Int64))
                '.Add("QTY_LEFT", GetType(System.Int64), "ISNULL(QTY_ALLO,0)-ISNULL(ORDR_QTY_SHIP,0)-ISNULL(ORDR_QTY_PICK,0)")
                '.Add("QTY_BAL", GetType(System.Int64), "IIF(QTY_LEFT>=0,QTY_LEFT,0)")
                '.Add("QTY_OVER", GetType(System.Int64), "IIF(QTY_LEFT-ISNULL(ORDR_QTY_OPEN,0)>=0,0,ISNULL(ORDR_QTY_OPEN,0)-QTY_LEFT)")
                .Add("QTY_BAL", GetType(System.Int64), "ISNULL(QTY_ALLO,0)-ISNULL(ORDR_QTY_SHIP,0)-ISNULL(ORDR_QTY_PICK,0)")
                .Add("QTY_LEFT", GetType(System.Int64), "ISNULL(QTY_BAL,0)-ISNULL(ORDR_QTY_OPEN,0)")
                .Add("QTY_OVER", GetType(System.Int64), "IIF(QTY_LEFT>=0,0,-1*QTY_LEFT)")
            End With

            ASCMAIN1.sql = "Select SOTORDR2.ALLO_CTL_NO, NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE) CUST_CODE" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
                & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'O',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_OPEN" & vbCrLf _
                & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'P',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_PICK" & vbCrLf _
                & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'F',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_SHIP" & vbCrLf _
                & " from SOTORDR2,SOTORDR1,ARTCUST1" & vbCrLf _
                & "   where SOTORDR2.ALLO_CTL_NO in  (Select ALLO_CTL_NO from SOTALLO1 where ITEM_CODE = :PARM1)" & vbCrLf _
                & "     and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                & "     and ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE" & vbCrLf _
                & "     and SOTORDR2.ORDR_STATUS IN ('O','P','F','C')" & vbCrLf _
                & " group by SOTORDR2.ALLO_CTL_NO, NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE)"
            Create_TDA(.Tables.Add, "SOTALLOZ", "**", 0, False, "V", 2)

            'ASCMAIN1.sql = "Select SOTORDR2.*,SOTORDR1.CUST_NAME" _
            '& " from SOTORDR2,SOTORDR1" _
            '& " where SOTORDR2.ITEM_CODE = :PARM1 and (SOTORDR1.ORDR_STATUS = 'O' or SOTORDR1.ORDR_STATUS = 'P')"
            'Create_TDA(.Tables.Add, "SOTORDRX", "**", 0, False, "V")

            ASCMAIN1.sql = "Select SOTINVH2.INV_TYPE, SOTINVH2.INV_NO, SOTINVH2.INV_LNO" & vbCrLf _
                & ", SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO, SOTINVH1.ORDR_CUST_PO, SOTINVH1.INV_DATE" & vbCrLf _
                & ", ARTCUST1.CUST_NAME, SOTINVH2.ORDR_QTY_SHIP" & vbCrLf _
                & ", DECODE(SOTINVH2.INV_TYPE,'I',SOTINVH2.ORDR_QTY_SHIP,0) SHP" & vbCrLf _
                & ", CASE WHEN SOTINVH2.INV_TYPE = 'C' AND NVL(SOTINVH1.ORDR_TYPE_CODE, 'X') <> 'DIF' THEN SOTINVH2.ORDR_QTY_SHIP ELSE 0 END RTN" & vbCrLf _
                & ", DECODE(SOTINVH1.ORDR_TYPE_CODE,'DIF',SOTINVH2.ORDR_QTY_SHIP, 0) DIF" & vbCrLf _
                & ", SOTINVH2.ORDR_UNIT_PRICE" & vbCrLf _
                & " from SOTINVH2,ARTCUST1,SOTINVH1 " & vbCrLf _
                & " where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
                & "  and SOTINVH2.ITEM_CODE = :PARM1 and SOTINVH2.ORDR_YYYYPP_UPDATED = :PARM2" & vbCrLf _
                & "  and ARTCUST1.CUST_CODE (+) = SOTINVH2.CUST_CODE"
            Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False, "VV", 3)
            .Tables("SOTINVHX").Columns("SHP").DataType = GetType(System.Int64)
            .Tables("SOTINVHX").Columns("RTN").DataType = GetType(System.Int64)
            .Tables("SOTINVHX").Columns("DIF").DataType = GetType(System.Int64)

            sqlICTSTATI = "SELECT " & vbCrLf _
            & "  ICTSTAT2.ITEM_CODE" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_CATGY_CODE" & vbCrLf _
            & ", ICTITEM1.DEPT_CODE" & vbCrLf _
            & ", ICTSTAT2.WHSE_CODE" & vbCrLf _
            & ", SUM (NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0)) QOH" & vbCrLf _
            & ", SUM (NVL(ICTSTAT2.WHSE_QTY_ONPO,0)) ONPO" & vbCrLf _
            & ", SUM (NVL(ICTSTAT2.WHSE_QTY_PLAN,0)) PLAN" & vbCrLf _
            & ", SUM (NVL(ICTSTAT2.WHSE_QTY_OPEN,0)) OPEN" & vbCrLf _
            & ", SUM (NVL(ICTSTAT2.WHSE_QTY_PICK,0)) PICK" & vbCrLf _
            & ", SUM (NVL(ICTSTAT2.WHSE_QTY_COMM,0)) COMM" & vbCrLf _
            & ", SUM (NVL(ICTSTAT2.WHSE_QTY_HOLD,0)) HOLD" & vbCrLf _
            & " from ICTSTAT2,ICTITEM1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = ICTSTAT2.ITEM_CODE" & vbCrLf _
            & " group by ICTSTAT2.ITEM_CODE" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE, ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.DEPT_CODE, ICTSTAT2.WHSE_CODE"
            ICTSTATI = ASCMAIN1.Temp_Table(Replace(sqlICTSTATI, "group by", " and ROWNUM < 1 group by"))
            ASCDATA1.ExecuteSQL("Create Index I_" & ICTSTATI & "_COLLECTION_CODE ON " & ICTSTATI & " (COLLECTION_CODE)")
            ASCDATA1.ExecuteSQL("Create Index I_" & ICTSTATI & "_ITEM_CATGY_CODE ON " & ICTSTATI & " (ITEM_CATGY_CODE)")
            ASCDATA1.ExecuteSQL("Create Index I_" & ICTSTATI & "_DEPT_CODE ON " & ICTSTATI & " (DEPT_CODE)")
            ASCDATA1.ExecuteSQL("Create Index I_" & ICTSTATI & "_WHSE_CODE ON " & ICTSTATI & " (WHSE_CODE)")

            sqlICTSTATP = "SELECT " _
            & "  ICTSTATI.COLLECTION_CODE" & vbCrLf _
            & ", ICTSTATI.ITEM_CATGY_CODE" & vbCrLf _
            & ", ICTSTATI.DEPT_CODE" & vbCrLf _
            & ", ICTSTATI.WHSE_CODE" & vbCrLf _
            & ", COUNT (*) ITEMS" _
            & ", SUM (ICTSTATI.QOH) QOH" _
            & ", SUM (ICTSTATI.ONPO) ONPO" _
            & ", SUM (ICTSTATI.PLAN) PLAN" _
            & ", SUM (ICTSTATI.OPEN) OPEN" _
            & ", SUM (ICTSTATI.PICK) PICK" _
            & ", SUM (ICTSTATI.COMM) COMM" _
            & ", SUM (ICTSTATI.HOLD) HOLD" _
            & " from " & ICTSTATI & " ICTSTATI" _
            & " group by COLLECTION_CODE,ITEM_CATGY_CODE,DEPT_CODE,WHSE_CODE"
            ICTSTATP = ASCMAIN1.Temp_Table(Replace(sqlICTSTATP, "group by", "where ROWNUM < 1 group by"))
            'ASCDATA1.ExecuteSQL("Alter Table " & ICTSTATP & " Add CODE_DESC VARCHAR2(60)")

            ASCMAIN1.sql = "Select ICTSTATP.* from " & ICTSTATP & " ICTSTATP"
            Create_TDA(.Tables.Add, "ICTSTATP", "**", 0, False, "", 0)
            .Tables("ICTSTATP").Columns("ITEMS").DataType = GetType(System.Int32)
            .Tables("ICTSTATP").Columns.Add("COLLECTION_DESC", GetType(System.String))
            .Tables("ICTSTATP").Columns.Add("ITEM_CATGY_DESC", GetType(System.String))
            .Tables("ICTSTATP").Columns.Add("DEPT_DESC", GetType(System.String))
            .Tables("ICTSTATP").Columns.Add("WHSE_DESC", GetType(System.String))
            .Tables("ICTSTATP").Columns.Add("WHSE_TYPE", GetType(System.String))


            ASCMAIN1.sql = "Select ICTSTATI.*,ICTITEM1.ITEM_DESC" & vbCrLf _
            & " from " & ICTSTATI & " ICTSTATI,ICTITEM1 " & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = ICTSTATI.ITEM_CODE" & vbCrLf _
            & " and (ICTSTATI.COLLECTION_CODE = :PARM1)" & vbCrLf _
            & " and (ICTSTATI.ITEM_CATGY_CODE = :PARM2)" & vbCrLf _
            & " and (ICTSTATI.DEPT_CODE = :PARM3)" & vbCrLf _
            & " and (ICTSTATI.WHSE_CODE = :PARM4)" & vbCrLf
            Create_TDA(.Tables.Add, "ICTSTATI", "**", 0, False, "VVVV", 1)

            ASCMAIN1.sql = "SELECT POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO" & vbCrLf _
            & ", POTORDR1.VEND_CODE, POTORDR2.WHSE_CODE" & vbCrLf _
            & ", POTORDR1.PO_DATE_ORDERED, POTORDR2.PO_DATE_REQUIRED" & vbCrLf _
            & ", POTORDR2.PO_COST" & vbCrLf _
            & ", POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_REC, POTORDR2.PO_QTY_OPN" & vbCrLf _
            & ", POTORDR2.PO_DATE_REQUESTED" & vbCrLf _
            & ", POTORDR2.PO_DATE_ETD, POTORDR2.PO_DATE_ETD_NOTES" & vbCrLf _
            & " from POTORDR2,POTORDR1 " & vbCrLf _
            & " where POTORDR2.ITEM_CODE = :PARM1" & vbCrLf _
            & "   and POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & "   and (:PARM2 = 'A' or POTORDR2.PO_STATUS = 'O')"
            Create_TDA(.Tables.Add, "POTORDRX", "**", 0, False, "VV", 2)
            With .Tables("POTORDRX")
                .Columns.Add("INV_QTY", GetType(System.Int32))
                .Columns.Add("BALANCE", GetType(System.Int32), "ISNULL(PO_QTY_OPN,0) - ISNULL(INV_QTY,0)")
            End With

            ASCMAIN1.sql = "Select POTORDR8.* from POTORDR8, POTORDR2, POTORDR1
                where POTORDR2.PO_ORDER_NO = POTORDR8.PO_ORDER_NO 
                and POTORDR2.PO_ORDER_LNO = POTORDR8.PO_ORDER_LNO and POTORDR2.ITEM_CODE = :PARM1
                and POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO
                and (:PARM2 = 'A' or POTORDR2.PO_STATUS = 'O')"
            Create_TDA(.Tables.Add, "POTORDR8", "**", 0, True, "VV")

            Create_Relation("POTORDRX", "POTORDR8", "PO_ORDER_NO,PO_ORDER_LNO")

            ASCMAIN1.sql = "SELECT SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO" & vbCrLf _
            & ", SOTORDR1.WHSE_CODE, SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_QTY_OPEN, SOTORDR2.ORDR_QTY_PICK, SOTORDR2.ORDR_QTY_CANC" & vbCrLf _
            & ", SOTORDR1.CUST_STORE_NO " & vbCrLf _
            & ", SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE " & vbCrLf _
            & ", SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_ALLO_DATE, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.CUST_DC_NO " & vbCrLf _
            & " FROM SOTORDR1,SOTORDR2" & vbCrLf _
            & " WHERE SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & " AND SOTORDR2.ITEM_CODE = :PARM1" & vbCrLf _
            & " AND (SOTORDR2.ORDR_STATUS IN ('O','P','B'))"
            Create_TDA(.Tables.Add, "SOTORDRX", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "SELECT DISTINCT BMTMAIN1.BM_PROD_ITEM, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_UPC_CODE, ICTITEM1.ITEM_UOM" & vbCrLf _
                 & " FROM ICTITEM1, BMTMAIN1, BMTMAIN3" & vbCrLf _
                 & " WHERE BMTMAIN1.BM_PROD_ITEM = ICTITEM1.ITEM_CODE" & vbCrLf _
                 & " AND BMTMAIN1.BM_PROD_ITEM = BMTMAIN3.BM_PROD_ITEM" & vbCrLf _
                 & " AND BMTMAIN3.BM_COMP_ITEM = :PARM1"
            Create_TDA(.Tables.Add, "BMTMAIN1", ASCMAIN1.sql, 0, False, "V", 0)

            'ASCMAIN1.sql = " SELECT BMTMAIN3.*, ICTITEM1.ITEM_DESC" & vbCrLf _
            '    & " FROM BMTMAIN3, ICTITEM1 " & vbCrLf _
            '    & " WHERE BMTMAIN3.BM_COMP_ITEM = ICTITEM1.ITEM_CODE" & vbCrLf _
            '    & " AND (BM_PROD_ITEM, BM_ISSUE_NO) IN" & vbCrLf _
            '    & " (" & vbCrLf _
            '    & " SELECT BM_PROD_ITEM, MAX(BM_ISSUE_NO) BM_ISSUE_NO" & vbCrLf _
            '    & " FROM BMTMAIN3" & vbCrLf _
            '    & " WHERE BM_COMP_ITEM = :PARM1" & vbCrLf _
            '    & " GROUP BY BM_PROD_ITEM" & vbCrLf _
            '    & " )"
            'Create_TDA(.Tables.Add, "BMTMAIN3", ASCMAIN1.sql, 0, False, "V", 0)

            ASCMAIN1.sql = "Select BMTMAIN3.*" & vbCrLf _
                & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_UOM" & vbCrLf _
                & ", ICTITEM1.ITEM_COST_STD" & vbCrLf _
                & ", ICTITEM1.ITEM_COST_WASTE_PCT" & vbCrLf _
                & ", ICTITEM1.ITEM_PLAN_WASTE_PCT" & vbCrLf _
                & ", ICTITEM1.VEND_CODE" & vbCrLf _
                & ", ICTITEM1.VEND_ITEM_CODE" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_VCOST" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_LANDG" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_TOOLG" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_OVRHD" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_TOTAL" & vbCrLf _
                & " from BMTMAIN3,ICTITEM1,ICTCOSTC" & vbCrLf _
                & " where BMTMAIN3.BM_COMP_ITEM = ICTITEM1.ITEM_CODE" & vbCrLf _
                & " and ICTCOSTC.ITEM_CODE (+) = BMTMAIN3.BM_COMP_ITEM " & vbCrLf _
                & " AND (BMTMAIN3.BM_PROD_ITEM, BMTMAIN3.BM_ISSUE_NO) IN" & vbCrLf _
                & " (" & vbCrLf _
                & " SELECT BM_PROD_ITEM, MAX(BM_ISSUE_NO) BM_ISSUE_NO" & vbCrLf _
                & " FROM BMTMAIN3" & vbCrLf _
                & " WHERE BM_COMP_ITEM = :PARM1" & vbCrLf _
                & " GROUP BY BM_PROD_ITEM" & vbCrLf _
                & " )"
            Create_TDA(.Tables.Add, "BMTMAIN3", ASCMAIN1.sql, 0, False, "V", 0)
            Dim CALC As String = "ISNULL(BM_QTY_PER_ASSY,0) * ISNULL(?,0) * (1 + ISNULL(ITEM_COST_WASTE_PCT,0)/100)"
            With .Tables("BMTMAIN3").Columns
                .Add("EXT_COST", GetType(System.Decimal), Replace(CALC, "?", "ITEM_COST_STD"))
                .Add("VCOST", GetType(System.Decimal), Replace(CALC, "?", "ITEM_COST_VCOST"))
                .Add("LANDG", GetType(System.Decimal), Replace(CALC, "?", "ITEM_COST_LANDG"))
                .Add("TOOLG", GetType(System.Decimal), Replace(CALC, "?", "ITEM_COST_TOOLG"))
                .Add("OVRHD", GetType(System.Decimal), Replace(CALC, "?", "ITEM_COST_OVRHD"))
                .Add("TOTAL", GetType(System.Decimal), Replace(CALC, "?", "ITEM_COST_TOTAL"))
                .Add("QTY_ON_HAND", GetType(System.Int32))
                .Add("QTY_ONPO", GetType(System.Int32))
                .Add("QTY_PLAN", GetType(System.Int32))
                .Add("QTY_OPEN", GetType(System.Int32))
                .Add("QTY_PICK", GetType(System.Int32))
                .Add("QTY_COMM", GetType(System.Int32))
                .Add("QTY_OPEN_PICK", GetType(System.Int32), "ISNULL(QTY_OPEN,0)+ISNULL(QTY_PICK,0)")
                .Add("QTY_AVA", GetType(System.Int32), "ISNULL(QTY_ON_HAND,0)+ISNULL(QTY_OPEN,0)+ISNULL(QTY_PLAN,0)-ISNULL(QTY_COMM,0)-ISNULL(QTY_OPEN,0)-ISNULL(QTY_PICK,0)")
                .Add("BM_COMPONENT_SORT")
            End With

            .Relations.Add("BMTMAIN1_BMTMAIN3", dst.Tables("BMTMAIN1").Columns("BM_PROD_ITEM"), dst.Tables("BMTMAIN3").Columns("BM_PROD_ITEM"))

            ASCMAIN1.sql = "" _
                & "Select POTORDR9.*, POTORDR1.VEND_CODE, POTORDR1.VEND_WHSE_CODE" _
                & ", POTORDR2.ITEM_CODE ITEM_CODE_PO, POTORDR2.ITEM_DESC ITEM_DESC_PO, POTORDR2.PO_DATE_COMPSDUE" _
                & " from POTORDR9,POTORDR1,POTORDR2 where POTORDR9.ITEM_CODE = :PARM1" _
                & " and POTORDR2.PO_ORDER_NO = POTORDR9.PO_ORDER_NO" _
                & " and POTORDR2.PO_ORDER_LNO = POTORDR9.PO_ORDER_LNO" _
                & " and POTORDR1.PO_ORDER_NO = POTORDR9.PO_ORDER_NO" _
                & " and POTORDR9.PO_ORDER_LNO <> 0" _
                & " union " _
                & "Select POTORDR9.*, DPTPLAN1.VEND_CODE, DPTPLAN1.AT_WHSE VEND_WHSE_CODE" _
                & ", DPTPLAN1.ITEM_CODE ITEM_CODE_PO, ICTITEM1.ITEM_DESC ITEM_DESC_PO, DPTPLAN1.DATE_COMPSDUE PO_DATE_COMPSDUE" _
                & " from POTORDR9,DPTPLAN1,ICTITEM1 where POTORDR9.ITEM_CODE = :PARM1" _
                & " and DPTPLAN1.PLAN_NO = POTORDR9.PO_ORDER_NO" _
                & " and ICTITEM1.ITEM_CODE = DPTPLAN1.ITEM_CODE" _
                & " and POTORDR9.PO_ORDER_LNO = 0"
            Create_TDA(.Tables.Add, "POTORDR9", "**", 0, True, "V", 3)

            If ASCMAIN1.Running_in_VS And ASCMAIN1.CLIENT = "XXX" Then
                ASCMAIN1.sql = "Select ICTSTAT2.ITEM_CODE,ICTSTAT2.WHSE_CODE,ICTCOLL1.BRAND_CODE, ICTBRAN1.BRAND_NAME,ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_UPC_CODE," & vbCrLf _
                    & "ICTSTAT2.WHSE_QTY_ON_HAND, ICTSTAT2.WHSE_QTY_ONPO, GABATS.PO_30DAY, GABATS.LOW, GABATS.DEAD_NET, GABATS.PO_COST, GABATS.AVG_COST, GABATS.SOLD_YTD, GABATS.AVG_SELL" & vbCrLf _
                    & " from ICTSTAT2, ICTITEM1, ICTBRAN1, ICTCOLL1, GABATS" & vbCrLf _
                    & " where ICTITEM1.ITEM_CODE = ICTSTAT2.ITEM_CODE" & vbCrLf _
                    & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                    & "   and GABATS.ITEM_CODE (+) = ICTSTAT2.ITEM_CODE" & vbCrLf _
                    & "   and GABATS.WHSE_CODE (+) = ICTSTAT2.WHSE_CODE" & vbCrLf _
                    & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf

                Create_TDA(.Tables.Add, "ICTSTATS", "**", 0, False, "", 2)
            End If

            ASCMAIN1.sql = "Select ICTPINV1.WHSE_CODE, ICTPINV1.INV_NUM, ICTPINV1.INV_DATE, ICTPINV2.PO_ORDER_NO, ICTPINV2.PO_ORDER_LNO" & vbCrLf _
                & ", ICTPINV1.VESSEL_NAME, ICTPINV1.BOL_NO, ICTPINV1.CONTAINER_NO, ICTPINV1.SHIP_DATE, ICTPINV1.ETA_DATE" & vbCrLf _
                & ", SUM(ICTPINV2.PINV_QTY) INV_QTY, ICTITEM1.PORT_CODE" & vbCrLf _
                & " from ICTPINV1, ICTPINV2, ICTITEM1" & vbCrLf _
                & " where ICTPINV1.PINV_NO = ICTPINV2.PINV_NO" & vbCrLf _
                & " And ICTPINV1.PINV_STATUS = 'O'" & vbCrLf _
                & " And ICTITEM1.ITEM_CODE = ICTPINV2.ITEM_CODE" & vbCrLf _
                & " AND ICTPINV2.ITEM_CODE = :PARM1" & vbCrLf _
                & " group by ICTPINV1.WHSE_CODE, ICTPINV1.INV_NUM, ICTPINV1.INV_DATE, ICTPINV2.PO_ORDER_NO, ICTPINV2.PO_ORDER_LNO" & vbCrLf _
                & ", ICTPINV1.VESSEL_NAME, ICTPINV1.BOL_NO, ICTPINV1.CONTAINER_NO, ICTPINV1.SHIP_DATE, ICTPINV1.ETA_DATE, ICTITEM1.PORT_CODE"
            Create_TDA(.Tables.Add, "ICTPINV1", "**", 0, False, "V", 5)
            With .Tables("ICTPINV1")
                .Columns.Add("ETA_DATE_DC", GetType(System.DateTime))
                .Columns("VESSEL_NAME").AllowDBNull = True
                .Columns("INV_DATE").AllowDBNull = True
                .Columns.Add("OPO_QTY", GetType(System.Int32))
                .Columns("INV_QTY").DataType = GetType(System.Int32)
            End With

            Create_TDA(.Tables.Add, "ICTPORT2", "*", 0, False)
            Fill_Records("ICTPORT2")

            Create_TDA(.Tables.Add, "ICTWHSE1", "*", 0, False)
            With .Tables("ICTWHSE1").Columns
                .Add("SEL")
            End With
            .Tables("ICTWHSE1").Columns("SEL").DefaultValue = "0"


            Create_TDA(.Tables.Add, "ICTPROD1", "*", 0, False)
            With .Tables("ICTPROD1").Columns
                .Add("SEL")
            End With
            .Tables("ICTPROD1").Columns("SEL").DefaultValue = "0"
            Fill_Records("ICTPROD1")

            Create_TDA(.Tables.Add, "ICTEXCL1", "*", 0, False)
            Fill_Records("ICTEXCL1")
        End With

        grdICTITEM1_Recent.DataSource = dst.Tables("ICTITEM1_RECENT")
        grdICTSTATX.DataSource = dst.Tables("ICTSTATX")
        grdICTTRANX.DataSource = dst.Tables("ICTTRANX")
        If ASCMAIN1.Running_in_VS And ASCMAIN1.CLIENT = "XXX" Then
            grdICTSTATS.DataSource = dst.Tables("ICTSTATS")
        End If


        For Each C As String In New String() {"S", "R", "P", "A", "T", "X", "V"}
            Dim COLUMN_NAME As String = "TRAN_QTY_" & C
            With grdICTTRANX.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                .Width = 70
                Dim DX As String = ""
                Select Case C
                    Case "S"
                        DX = "Shp"
                    Case "R"
                        DX = "Rtn"
                    Case "P"
                        DX = "Rec"
                    Case "A"
                        DX = "Adj"
                    Case "T"
                        DX = "Xfr"
                    Case "X"
                        DX = "Con"
                    Case "V"
                        DX = "RTV"
                End Select
                .Header.Caption = DX
            End With
            'grdICTTRANX.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption = IIf()
            Create_Summary(grdICTTRANX, COLUMN_NAME)
        Next

        grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")

        grdICTSTATP.DataSource = dst.Tables("ICTSTATP")
        grdICTSTATI.DataSource = dst.Tables("ICTSTATI")

        grdPOTORDRX.DataSource = dst.Tables("POTORDRX")
        grdSOTORDRX.DataSource = dst.Tables("SOTORDRX")

        grdSOTALLO1.DataSource = dst.Tables("SOTALLO1")
        grdSOTALLO2.DataSource = dst.Tables("SOTALLO2")

        grdBMTMAIN1.DataSource = dst.Tables("BMTMAIN1")
        grdPOTORDR9.DataSource = dst.Tables("POTORDR9")

        grdICTSTATO.DataSource = dst.Tables("ICTSTATO")
        grdICTPINV1.DataSource = dst.Tables("ICTPINV1")

        grdICTPROD1.DataSource = dst.Tables("ICTPROD1")
        grdICTWHSE1.DataSource = dst.Tables("ICTWHSE1")

        Create_Summary(grdICTSTATO, "ITEM_CODE", "Count")

        Show_Filter(grdICTITEM1_Recent, True)
        grdICTITEM1_Recent.DisplayLayout.GroupByBox.Hidden = False
        Create_Summary(grdICTITEM1_Recent, "ITEM_CODE", "Count")

        Bind_Controls(grpICTITEM1, "ICTITEM1")

        With grdICTITEM1_Recent.DisplayLayout.Bands(0)

            For Each COLUMN_NAME As String In New String() _
         {"ITEM_CODE", "ITEM_STATUS", "ITEM_DESC", "COLLECTION_CODE", "ITEM_CATGY_CODE",
          "ITEM_BASIC_PROMO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                If New String() {"QTY_ONHD", "QTY_ONPO", "QTY_PLAN", "QTY_OPEN", "QTY_PICK", "QTY_COMM", "QTY_NETA"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    If gcol.Key = "QTY_NETA" Then
                        'gcol.CellAppearance.ForeColor = Color.Purple
                    ElseIf New String() {"QTY_ONHD", "QTY_ONPO", "QTY_PLAN"}.Contains(gcol.Key) Then
                        gcol.CellAppearance.ForeColor = Color.Green
                    Else
                        gcol.CellAppearance.ForeColor = Color.Red
                    End If
                ElseIf gcol.Key = "MTD_SHP" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf gcol.Key = "ITEM_FC_MOS" Or (gcol.Key.Length <= 8 And gcol.Key.StartsWith("ITEM_FC")) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                ElseIf gcol.Key.Length <= 9 And gcol.Key.StartsWith("ITEM_POS") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Violet
                ElseIf gcol.Key.StartsWith("ITEM_PALLET") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Azure
                ElseIf New String() {"CARTON_PACK_QTY", "ITEM_SO_QTY_MULT", "ITEM_SO_QTY_MIN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"ITEM_BUFFER_QTY", "ITEM_BUFFER_PCT", "ITEM_ABC_PARMS_LOCKED"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                ElseIf New String() {"ITEM_CODE", "ITEM_STATUS", "ITEM_DESC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    gcol.Header.Fixed = True
                ElseIf New String() {"VEND_CODE", "VEND_ITEM_CODE", "ITEM_COST_MAKE_BUY", "ITEM_PLAN_MAKE_BUY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"PROD_CODE", "ITEM_CLASS_CODE", "COLLECTION_CODE", "ITEM_TYPE_CODE", "COST_CATGY_CODE",
                                     "ITEM_SNU_CODE", "ITEM_PWP_GWP_CODE", "ITEM_BASIC_PROMO", "ITEM_NOT_ALLOCATED", "ITEM_CATGY_CODE",
                                     "SEASON_CODE", "DEPT_CODE", "ACTUAL_LAUNCH", "ITEM_ORDR_REL_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                ElseIf New String() {"ITEM_UPC_CODE", "ITEM_EAN_CODE", "ITEM_UOM", "ITEM_RETAIL_PRICE", "ITEM_VALUE", "ITEM_BIN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Yellow
                ElseIf New String() {"ITEM_AEROSOL", "ITEM_CRITICAL_TO_SHIP", "ITEM_CRITICAL_TO_SHIP",
                    "ITEM_WEIGHT_CHECK", "ITEM_APPR_1ST_REC", "ITEM_ALLOW_HALF_PACK", "ITEM_LOT_CONTROL"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LimeGreen
                ElseIf New String() {"ITEM_ABC_CODE", "ITEM_POS_MAX", "ITEM_POS_MIN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Magenta
                ElseIf New String() {"ITEM_COST_STD", "EXT_COST"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                End If
            Next

            For I As Integer = 0 To 3
                Dim mmm As String = "PD"
                If I = 0 Then
                Else
                    Dim YP_Legend As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, I - 1))
                    mmm = Mid(YP_Legend, 10, 3)
                End If

                With .Columns("ITEM_FC" & CStr(I))
                    .Header.Caption = "FC " & mmm
                    .Width = grdICTITEM1_Recent.DisplayLayout.Bands(0).Columns("ITEM_FC").Width
                    .Format = grdICTITEM1_Recent.DisplayLayout.Bands(0).Columns("ITEM_FC").Format
                    If I = 0 Then ' lm email 04/08/21
                        .Header.Caption = "FC Over CM"
                        .Width = grdICTITEM1_Recent.DisplayLayout.Bands(0).Columns("ITEM_FC").Width * 1.2
                    End If
                End With
                If I = 0 Then
                    ' .Columns("ITEM_POS" & CStr(I)).Hidden = False
                    .Columns("ITEM_POS" & CStr(I)).Header.Caption = "OH Pos"
                Else
                    .Columns("ITEM_POS" & CStr(I)).Header.Caption = "Pos " & mmm
                End If
                With .Columns("ITEM_POS" & CStr(I))
                    ' .Header.Caption = "Pos " & mmm
                    .Width = grdICTITEM1_Recent.DisplayLayout.Bands(0).Columns("ITEM_POS").Width
                    .Format = grdICTITEM1_Recent.DisplayLayout.Bands(0).Columns("ITEM_POS").Format
                End With
            Next

        End With

        With grdICTITEM1_optP.DisplayLayout.Bands(0)

            For Each COLUMN_NAME As String In New String() _
         {"ITEM_CODE", "ITEM_DESC", "SALES_DIVISION_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                If New String() {"QTY_ONHD", "QTY_ONPO", "QTY_PLAN", "QTY_OPEN", "QTY_PICK", "QTY_COMM", "QTY_NETA"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    If gcol.Key = "QTY_NETA" Then
                        'gcol.CellAppearance.ForeColor = Color.Purple
                    ElseIf New String() {"QTY_ONHD", "QTY_ONPO", "QTY_PLAN"}.Contains(gcol.Key) Then
                        gcol.CellAppearance.ForeColor = Color.Green
                    Else
                        gcol.CellAppearance.ForeColor = Color.Red
                    End If
                ElseIf gcol.Key = "MTD_SHP" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf gcol.Key = "ITEM_FC_MOS" Or (gcol.Key.Length <= 8 And gcol.Key.StartsWith("ITEM_FC")) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                ElseIf gcol.Key.Length <= 9 And gcol.Key.StartsWith("ITEM_POS") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Violet
                ElseIf gcol.Key.StartsWith("ITEM_PALLET") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Azure
                ElseIf New String() {"CARTON_PACK_QTY", "ITEM_SO_QTY_MULT", "ITEM_SO_QTY_MIN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"ITEM_BUFFER_QTY", "ITEM_BUFFER_PCT", "ITEM_ABC_PARMS_LOCKED"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                ElseIf New String() {"ITEM_CODE", "ITEM_STATUS", "ITEM_DESC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    gcol.Header.Fixed = True
                ElseIf New String() {"VEND_CODE", "VEND_ITEM_CODE", "ITEM_COST_MAKE_BUY", "ITEM_PLAN_MAKE_BUY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"PROD_CODE", "ITEM_CLASS_CODE", "COLLECTION_CODE", "ITEM_TYPE_CODE", "COST_CATGY_CODE",
                                     "ITEM_SNU_CODE", "ITEM_PWP_GWP_CODE", "ITEM_BASIC_PROMO", "ITEM_NOT_ALLOCATED", "ITEM_CATGY_CODE",
                                     "SEASON_CODE", "DEPT_CODE", "ACTUAL_LAUNCH", "ITEM_ORDR_REL_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                ElseIf New String() {"ITEM_UPC_CODE", "ITEM_EAN_CODE", "ITEM_UOM", "ITEM_RETAIL_PRICE", "ITEM_VALUE", "ITEM_BIN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Yellow
                ElseIf New String() {"ITEM_AEROSOL", "ITEM_CRITICAL_TO_SHIP", "ITEM_CRITICAL_TO_SHIP",
                    "ITEM_WEIGHT_CHECK", "ITEM_APPR_1ST_REC", "ITEM_ALLOW_HALF_PACK", "ITEM_LOT_CONTROL"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LimeGreen
                ElseIf New String() {"ITEM_ABC_CODE", "ITEM_POS_MAX", "ITEM_POS_MIN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Magenta
                ElseIf New String() {"ITEM_COST_STD", "EXT_COST"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                End If
            Next

            For I As Integer = 0 To 3
                Dim mmm As String = "PD"
                If I = 0 Then
                Else
                    Dim YP_Legend As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, I - 1))
                    mmm = Mid(YP_Legend, 10, 3)
                End If

                With .Columns("ITEM_FC" & CStr(I))
                    .Header.Caption = "FC " & mmm
                    .Width = grdICTITEM1_Recent.DisplayLayout.Bands(0).Columns("ITEM_FC").Width
                    .Format = grdICTITEM1_Recent.DisplayLayout.Bands(0).Columns("ITEM_FC").Format
                    If I = 0 Then ' lm email 04/08/21
                        .Header.Caption = "FC Over CM"
                        .Width = grdICTITEM1_Recent.DisplayLayout.Bands(0).Columns("ITEM_FC").Width * 1.2
                    End If
                End With
                If I = 0 Then
                    ' .Columns("ITEM_POS" & CStr(I)).Hidden = False
                    .Columns("ITEM_POS" & CStr(I)).Header.Caption = "OH Pos"
                Else
                    .Columns("ITEM_POS" & CStr(I)).Header.Caption = "Pos " & mmm
                End If
                With .Columns("ITEM_POS" & CStr(I))
                    ' .Header.Caption = "Pos " & mmm
                    .Width = grdICTITEM1_Recent.DisplayLayout.Bands(0).Columns("ITEM_POS").Width
                    .Format = grdICTITEM1_Recent.DisplayLayout.Bands(0).Columns("ITEM_POS").Format
                End With
            Next

        End With

        With grdICTSTATX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() _
            {"WHSE_QTY_BEG", "WHSE_QTY_SHP", "WHSE_QTY_RTN", "WHSE_QTY_REC", "WHSE_QTY_ADJ",
             "WHSE_QTY_XFR", "WHSE_QTY_CON", "WHSE_QTY_RTV", "WHSE_QTY_PHY"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Color.Yellow
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() _
            {"WHSE_QTY_ON_HAND", "WHSE_QTY_ONPO", "WHSE_QTY_PLAN", "WHSE_QTY_OPEN",
             "WHSE_QTY_PICK", "WHSE_QTY_COMM", "WHSE_QTY_HOLD", "WHSE_QTY_NETA", "WHSE_QTY_ATS"}
                .Columns(COLUMN_NAME).Header.Appearance.ForeColor = Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Color.DarkGreen
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
            COLUMN_NAME = "WHSE_CODE"
            '  .Columns(COLUMN_NAME).Header.Appearance.ForeColor = Color.White
            .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Color.LightBlue
            .Columns(COLUMN_NAME).Header.Appearance.BackColor = Color.White
            .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            COLUMN_NAME = "OPS_YYYYPP"
            ' .Columns(COLUMN_NAME).Header.Appearance.ForeColor = Color.White
            .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Color.Orange
            .Columns(COLUMN_NAME).Header.Appearance.BackColor = Color.White
            .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            .Columns("OPS_YYYYPP").Width = .Columns("WHSE_CODE").Width


            .Columns("WHSE_QTY_NETA").CellAppearance.BackColor = Color.LightBlue
            .Columns("WHSE_QTY_ATS").CellAppearance.BackColor = Color.LightBlue
        End With

        With grdSOTALLO2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If gcol.Key = "QTY_BAL" Or gcol.Key = "QTY_LEFT" Or gcol.Key = "QTY_OVER" Then
                    gcol.CellAppearance.BackColor = Color.LightGreen
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                ElseIf gcol.Key = "ORDR_QTY" Or gcol.Key = "ORDR_QTY_CANC" Then
                    gcol.CellAppearance.BackColor = Color.LightGray
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                ElseIf gcol.Key = "ORDR_QTY_SHIP" Or gcol.Key = "ORDR_QTY_PICK" Or gcol.Key = "ORDR_QTY_OPEN" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                ElseIf gcol.Key = "CUST_CODE" Or gcol.Key = "QTY_ALLO" Or gcol.Key = "ALLO_NOTES" Then
                    gcol.Header.Appearance.BackColor2 = Color.Orange
                End If
            Next
        End With

        If ASCMAIN1.Running_in_VS And ASCMAIN1.CLIENT = "XXX" Then
            With grdICTSTATS.DisplayLayout.Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    gcol.Header.Appearance.BackColor = Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    If gcol.Key.StartsWith("QTY") Then
                        '  gcol.CellAppearance.BackColor = Color.LightBlue
                        gcol.Header.Appearance.BackColor2 = Color.LightBlue
                    ElseIf gcol.Key.Contains("PRICE") Then
                        ' gcol.CellAppearance.BackColor = Color.LightGray
                        gcol.Header.Appearance.BackColor2 = Color.LightGreen
                        'ElseIf gcol.Key = "ORDR_QTY_SHIP" Or gcol.Key = "ORDR_QTY_PICK" Or gcol.Key = "ORDR_QTY_OPEN" Then
                        '    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                    ElseIf gcol.Key = "PO_30DAY" Or gcol.Key = "DEAD_NET" Or gcol.Key = "LOW" Or gcol.Key = "PO_COST" Or gcol.Key = "AVG_COST" Or gcol.Key = "AVG_SELL" Then
                        gcol.Header.Appearance.BackColor2 = Color.Orange
                    End If

                    If gcol.Key = "ITEM_CODE" Or gcol.Key = "BRAND_NAME" Or gcol.Key = "ITEM_DESC" Then
                        gcol.Header.Fixed = True
                    End If
                Next
            End With
        End If

        With grdPOTORDRX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Color.LightBlue
            Next
        End With
        With grdPOTORDRX.DisplayLayout.Bands(1)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Color.LightGreen
            Next
        End With
        grdPOTORDRX.DisplayLayout.Override.ExpansionIndicator = ShowExpansionIndicator.CheckOnDisplay

        grdICTPROD1.DisplayLayout.UseFixedHeaders = True
        grdICTPROD1.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
        grdICTPROD1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdICTPROD1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        With grdICTPROD1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SEL"}
                .Columns(COLUMN_NAME).Header.Fixed = True
                .Columns(COLUMN_NAME).CellActivation = Activation.AllowEdit
            Next
            For Each COLUMN_NAME As String In New String() {"PROD_CODE", "PROD_DESC"}
                .Columns(COLUMN_NAME).CellActivation = Activation.NoEdit
            Next
        End With

        grdICTWHSE1.DisplayLayout.UseFixedHeaders = True
        grdICTWHSE1.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
        grdICTWHSE1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdICTWHSE1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        With grdICTWHSE1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SEL"}
                .Columns(COLUMN_NAME).Header.Fixed = True
                .Columns(COLUMN_NAME).CellActivation = Activation.AllowEdit
            Next
            For Each COLUMN_NAME As String In New String() {"WHSE_CODE", "WHSE_DESC"}
                .Columns(COLUMN_NAME).CellActivation = Activation.NoEdit
            Next
        End With

        Create_Summary(grdICTSTATX, "OPS_YYYYPP", "Count")
        Create_Summary(grdICTSTATX, "WHSE_CODE", "Count")
        Create_Summary(grdICTSTATX, New String() _
                       {"WHSE_QTY_BEG", "WHSE_QTY_SHP", "WHSE_QTY_RTN", "WHSE_QTY_REC", "WHSE_QTY_ADJ",
                        "WHSE_QTY_XFR", "WHSE_QTY_CON", "WHSE_QTY_RTV", "WHSE_QTY_PHY",
                        "WHSE_QTY_ON_HAND", "WHSE_QTY_ONPO", "WHSE_QTY_PLAN", "WHSE_QTY_OPEN", "WHSE_QTY_PICK", "WHSE_QTY_COMM", "WHSE_QTY_HOLD", "WHSE_QTY_NETA", "WHSE_QTY_ATS"})

        Create_Summary(grdICTTRANX, "TRAN_DATE", "Count")
        Create_Summary(grdICTTRANX, "TRAN_QTY")

        Create_Summary(grdSOTINVHX, "INV_NO", "Count")
        Create_Summary(grdSOTINVHX, New String() {"ORDR_QTY_SHIP", "SHP", "RTN", "DIF"})

        Create_Summary(grdPOTORDRX, "PO_ORDER_NO", "Count")
        Create_Summary(grdPOTORDRX, New String() {"PO_QTY_ORD", "PO_QTY_REC", "PO_QTY_OPN", "INV_QTY", "BALANCE"})

        Create_Summary(grdSOTORDRX, "ORDR_NO", "Count")
        Create_Summary(grdSOTORDRX, New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_CANC"})

        Create_Summary(grdICTSTATP, "COLLECTION_CODE", "Count")
        Create_Summary(grdICTSTATP, "ITEMS")
        Create_Summary(grdICTSTATI, "ITEM_CODE", "Count")

        Create_Summary(grdSOTALLO2, New String() {"QTY_ALLO", "ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "QTY_LEFT", "QTY_BAL", "QTY_OVER"})

        Create_Summary(grdPOTORDR9, "PO_ORDER_NO", "Count")
        Create_Summary(grdPOTORDR9, New String() {"PO_QTY_COM"})

        For Each COLUMN_NAME In New String() _
        {"QOH", "ONPO", "OPEN", "PICK", "COMM"}
            Call Create_Summary(grdICTSTATP, COLUMN_NAME, , , "###,##0")
            Call Create_Summary(grdICTSTATI, COLUMN_NAME, , , "###,##0")
            If COLUMN_NAME = "QOH" Or COLUMN_NAME = "ONPO" Then
                grdICTSTATP.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Width = 80
                grdICTSTATI.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Width = 80
            Else
                grdICTSTATP.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Width = 60
                grdICTSTATI.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Width = 60
            End If
            grdICTSTATI.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Format = "###,##0"
            grdICTSTATP.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Format = "###,##0"
        Next

        For Each COLUMN_NAME As String In New String() {"COLLECTION_CODE", "ITEM_CATGY_CODE", "DEPT_CODE", "WHSE_CODE", "ITEMS"}
            grdICTSTATP.DisplayLayout.Bands(0).Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
        Next
        For Each COLUMN_NAME As String In New String() {"ITEM_CODE", "ITEM_DESC"}
            grdICTSTATI.DisplayLayout.Bands(0).Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
        Next

        With grdICTSTATX.DisplayLayout.Bands("ICTSTATX")
            .Columns("OPS_YYYYPP").Header.Fixed = True
            .Columns("WHSE_CODE").Header.Fixed = True
        End With

        With grdICTTRANX.DisplayLayout.Bands("ICTTRANX")
            .Columns("TRAN_DATE").Header.Fixed = True
        End With

        grdICTSTATP.DisplayLayout.UseFixedHeaders = True
        With grdICTSTATP.DisplayLayout.Bands(0)
            .Columns("COLLECTION_CODE").Header.Fixed = True
            .Columns("ITEM_CATGY_CODE").Header.Fixed = True
            .Columns("DEPT_CODE").Header.Fixed = True
            .Columns("WHSE_CODE").Header.Fixed = True
            .Columns("COLLECTION_DESC").Header.Fixed = True
            .Columns("ITEM_CATGY_DESC").Header.Fixed = True
            .Columns("DEPT_DESC").Header.Fixed = True
            .Columns("WHSE_TYPE").Header.Fixed = True
        End With
        grdICTSTATI.DisplayLayout.UseFixedHeaders = True
        With grdICTSTATI.DisplayLayout.Bands(0)
            .Columns("ITEM_CODE").Header.Fixed = True
        End With

        Create_Summary(grdICTPINV1, "PO_ORDER_NO", "Count")
        Create_Summary(grdICTPINV1, New String() {"INV_QTY", "OPO_QTY"})
        With grdICTPINV1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Color.LightGray
                If gcol.Key = "INV_QTY" Or gcol.Key = "OPO_QTY" Then
                    gcol.Format = "###,##0"
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                Else

                End If
            Next

        End With

        ASCMAIN1.Add_Value_List(grdICTTRANX, "TRAN_TYPE")
        ASCMAIN1.Add_Value_List(grdICTTRANX, "TRAN_SOURCE")

        ASCMAIN1.Add_Value_List(grdICTSTATP, "ITEM_CATGY_CODE")
        ASCMAIN1.Add_Value_List(grdICTSTATP, "WHSE_TYPE")


        ASCMAIN1.Add_Value_List(grdICTSTATI, "ITEM_CATGY_CODE")
        ASCMAIN1.Add_Value_List(grdICTTRANX, "TRAN_TYPE")

        ASCMAIN1.Add_Value_List(grdICTITEM1_optP, "FC_MONTH", , New String() {":", "x1:Cur Mo", "x2:Cur Mo + 1", "x3:Cur Mo + 2"})
        Set_Read_Only(grpItemMasterData, True)

        Fill_Records("ICTCOLL1")
        Fill_Records("ICTCATG1")
        Fill_Records("ICTDEPT1")
        Fill_Records("ICTWHSE1")

        tab.Visible = False
        grdICTSTATI.Visible = False
        grdICTSTATP.Visible = False

        Setup_Recent()
        ReParent_Tabs(tabViewItems)

        Set_Read_Only(grpICTITEM1, True)
        ASCMAIN1.Add_Value_List(Absx1.cbeFor("ITEM_ORDR_REL_CODE"), "ICTITEM1.ITEM_ORDR_REL_CODE")


        'Create_Summary(grdBMTMAIN1, "BM_COMP_ITEM", "Count", Nothing, "BMTMAIN1_BMTMAIN3")
        Create_Summary(grdBMTMAIN1, "EXT_COST", Nothing, "BMTMAIN1_BMTMAIN3")
        Create_Summary(grdBMTMAIN1, New String() {"VCOST", "LANDG", "TOOLG", "OVRHD", "TOTAL"}, Nothing, "BMTMAIN1_BMTMAIN3")

        grdBMTMAIN1.DisplayLayout.UseFixedHeaders = True
        With grdBMTMAIN1.DisplayLayout.Bands(1)
            .Columns("BM_COMP_ITEM").Header.Fixed = True
            .Columns("ITEM_DESC").Header.Fixed = True
            .Columns("BM_SEQ").Header.Fixed = True

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                Dim COLUMN_NAME As String = gcol.Key
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"ITEM_DESC", "ITEM_UOM", "ITEM_PLAN_WASTE_PCT"}.Contains(gcol.Key) Then
                    .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit

                ElseIf New String() {"VEND_CODE", "VEND_ITEM_CODE"}.Contains(gcol.Key) Then
                    '.Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightPink
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit

                ElseIf New String() {"ITEM_COST_STD", "ITEM_COST_WASTE_PCT", "EXT_COST"}.Contains(gcol.Key) Then
                    '.Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightPink
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightPink
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit

                ElseIf New String() {"QTY_ON_HAND", "QTY_ONPO", "QTY_PLAN", "QTY_OPEN", "QTY_PICK", "QTY_COMM", "QTY_OPEN_PICK", "QTY_AVA"}.Contains(gcol.Key) Then
                    '.Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightBlue
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit

                ElseIf New String() {"VCOST", "LANDG", "TOOLG", "OVRHD", "TOTAL"}.Contains(gcol.Key) Then
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        'lblITEM_EAN_CODE.Visible = (ROWs("ICTPARM1").Item("IC_PARM_UPC_OR_EAN") & "" = "EN")
        'txtITEM_EAN_CODE.Visible = (ROWs("ICTPARM1").Item("IC_PARM_UPC_OR_EAN") & "" = "EN")
        'lblITEM_UPC_CODE.Visible = (ROWs("ICTPARM1").Item("IC_PARM_UPC_OR_EAN") & "" <> "EN")
        'txtITEM_UPC_CODE.Visible = (ROWs("ICTPARM1").Item("IC_PARM_UPC_OR_EAN") & "" <> "EN")

        splMain.SplitterDistance = splMain.Height * 0.33
        splSOTALLO1.SplitterDistance = splSOTALLO1.Height / 2

        If ASCMAIN1.CLIENT = "XXX" Then

        Else
            optViewItems.ValueList.ValueListItems.Remove(4)
            btnATSRefresh.Visible = False
        End If

        ASCMAIN1.sql = "Select WHSE_CODE, WHSE_DESC from ICTWHSE1 where WHSE_STATUS = 'A' order by WHSE_CODE"
        cbeWHSE_CODE.DataSource = ASCDATA1.GetDataTable
        Dim defWhse As String = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")
        cbeWHSE_CODE.Value = defWhse

        For Each row As UltraGridRow In grdICTWHSE1.Rows
            If row.Cells("WHSE_CODE").Value.ToString() = defWhse Then
                row.Cells("SEL").Value = 1
                row.Update()
            End If
        Next

        For Each row As UltraGridRow In grdICTPROD1.Rows
            If row.Cells("PROD_CODE").Value.ToString() = "SB" Then
                row.Cells("SEL").Value = 1
                row.Update()
            End If
        Next
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"

                If Absx1.txtFor("ITEM_UPC_CODE").Text <> "" Then
                    ASCMAIN1.sql = "Select ITEM_CODE from ICTITEM1 where ITEM_UPC_CODE LIKE '%' || :PARM1"
                    Dim ITEM_CODE As String = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {Absx1.txtFor("ITEM_UPC_CODE").Text})
                    If ITEM_CODE = "" Then
                        ASCMAIN1.sql = "Select ITEM_CODE from ICTITEM1 where ITEM_EAN_CODE LIKE '%' || :PARM1"
                        ITEM_CODE = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {Absx1.txtFor("ITEM_UPC_CODE").Text})
                    End If
                    If ITEM_CODE <> "" Then
                        Absx1.txtFor("ITEM_CODE").Text = ITEM_CODE
                    End If
                End If

                If Absx1.txtFor("ITEM_EAN_CODE").Text <> "" Then
                    ASCMAIN1.sql = "Select ITEM_CODE from ICTITEM1 where ITEM_EAN_CODE LIKE '%' || :PARM1"
                    Dim ITEM_CODE As String = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {Absx1.txtFor("ITEM_EAN_CODE").Text})
                    If ITEM_CODE = "" Then
                        ASCMAIN1.sql = "Select ITEM_CODE from ICTITEM1 where ITEM_UPC_CODE LIKE '%' || :PARM1"
                        ITEM_CODE = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {Absx1.txtFor("ITEM_EAN_CODE").Text})
                    End If
                    If ITEM_CODE <> "" Then
                        Absx1.txtFor("ITEM_CODE").Text = ITEM_CODE
                    End If
                End If

                If Absx1.txtFor("ITEM_CODE").Text = "" Then
                    EMsg &= vbCr & "You must specify an Item Code to View"
                Else
                    rowICTITEM1 = LookUp("ICTITEM1", Absx1.txtFor("ITEM_CODE").Text)
                    If rowICTITEM1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Item " & Absx1.txtFor("ITEM_CODE").Text & " on File"
                    End If
                End If

            Case "Refresh Summary"
                If Not chkCOLLECTION_CODE.Checked And
                   Not chkITEM_CATGY_CODE.Checked And
                   Not chkDEPT_CODE.Checked And
                   Not chkWHSE_CODE.Checked Then
                    EMsg &= vbCr & "You Must Select at least 1 Field to Summarize By"
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

            Case "View"
                EntryMode = "V"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Refresh Summary"
                Call Refresh_Summary()

            Case "Cancel", "Done"
                Call Mode_Settings(False)

            Case "Integrity Check"
                Integrity_Check()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Refresh Summary").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Integrity Check").Visible = Not ScreenMode And ASCMAIN1.Running_in_VS
                End With

                .Groups("Status/Activity").Visible = ScreenMode
                .Groups("View Items").Visible = Not ScreenMode
                .Groups("Item Image").Visible = ScreenMode
                .Groups("Excluded Items").Visible = optViewItems.Value = "P" And Not ScreenMode
                .Groups("Warehouses").Visible = optViewItems.Value = "P" And Not ScreenMode
                .Groups("Prod Codes").Visible = optViewItems.Value = "P" And Not ScreenMode
            End With

        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        splMain.Visible = ScreenMode
        Setup_ViewItems()

        grpItemMasterData.Visible = ScreenMode
        Setup_tabDetails()

        If ScreenMode Then

            lblITEM_EAN_CODE.Visible = txtITEM_EAN_CODE.Text <> ""
            txtITEM_EAN_CODE.Visible = txtITEM_EAN_CODE.Text <> ""
            lblITEM_UPC_CODE.Visible = txtITEM_EAN_CODE.Text = ""
            txtITEM_UPC_CODE.Visible = txtITEM_EAN_CODE.Text = ""

        Else

            grdICTSTATO.Visible = False
            chkHistory.Checked = False

            Clear_Record()

            lblITEM_EAN_CODE.Visible = (ROWs("ICTPARM1").Item("IC_PARM_UPC_OR_EAN") & "" = "EN")
            txtITEM_EAN_CODE.Visible = (ROWs("ICTPARM1").Item("IC_PARM_UPC_OR_EAN") & "" = "EN")
            lblITEM_UPC_CODE.Visible = (ROWs("ICTPARM1").Item("IC_PARM_UPC_OR_EAN") & "" <> "EN")
            txtITEM_UPC_CODE.Visible = (ROWs("ICTPARM1").Item("IC_PARM_UPC_OR_EAN") & "" <> "EN")

        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ICTITEM1", "ICTSTATX", "SOTINVHX", "ICTTRANX", "SOTALLO1",
                                                       "SOTALLO2", "BMTMAIN1", "BMTMAIN3", "POTORDR9", "ICTPINV1",
                                                       "SOTORDRX", "POTORDRX", "POTORDR8"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
        RYP = ASCMAIN1.CYP
        Absx1.txtFor("ITEM_CODE").Text = ""
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)
        ITEM_CODE = Absx1.txtFor("ITEM_CODE").Text

        If EntryMode = "N" Then
        Else
            rowICTITEM1 = Fill_Record("ICTITEM1", ITEM_CODE)
            dst.AcceptChanges()
        End If


        Dim rowICTITEM1_RECENT As DataRow = dst.Tables("ICTITEM1_RECENT").Rows.Find(ITEM_CODE)
        If rowICTITEM1_RECENT Is Nothing Then
            rowICTITEM1_RECENT = dst.Tables("ICTITEM1_RECENT").NewRow
            rowICTITEM1_RECENT.ItemArray = rowICTITEM1.ItemArray
            dst.Tables("ICTITEM1_RECENT").Rows.Add(rowICTITEM1_RECENT)
        End If

        EnforceConstraints(False)

        grdPOTORDRX.Text = "Open Purchase Orders"
        grdSOTORDRX.Text = "Open Sales Orders"
        grdSOTINVHX.Text = "Sales Shipments / Returns"
        grdICTTRANX.Text = "Transaction Details"
        grdPOTORDR9.Text = "Component Commitments"

        Load_ICTSTATX()
        Setup_grdICTSTATX()

        Fill_Records("POTORDRX", New String() {ITEM_CODE, "O"})
        Fill_Records("POTORDR8", {ITEM_CODE, "O"})
        Sort_grdColumns(grdPOTORDRX, "INIT_DATE".ToLower,, 1)

        Fill_Records("SOTORDRX", ITEM_CODE)

        grdPOTORDRX.Text = "Open Purchase Orders for Item " & ITEM_CODE
        grdSOTORDRX.Text = "Open Sales Orders for Item " & ITEM_CODE

        Fill_Records("SOTALLO1", ITEM_CODE)
        Fill_Records("SOTALLO2", ITEM_CODE)
        Fill_Records("SOTALLOZ", ITEM_CODE)

        For Each rowSOTALLOZ As DataRow In dst.Tables("SOTALLOZ").Select("")
            Dim ALLO_CTL_NO As String = rowSOTALLOZ.Item("ALLO_CTL_NO")
            Dim CUST_CODE As String = rowSOTALLOZ.Item("CUST_CODE")
            Dim rowSOTALLO2 As DataRow = dst.Tables("SOTALLO2").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE})
            If rowSOTALLO2 Is Nothing Then
                rowSOTALLO2 = dst.Tables("SOTALLO2").NewRow
                rowSOTALLO2.Item("ALLO_CTL_NO") = ALLO_CTL_NO
                rowSOTALLO2.Item("CUST_CODE") = CUST_CODE
                dst.Tables("SOTALLO2").Rows.Add(rowSOTALLO2)
            End If
            For Each COLUMN_NAME As String In New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}
                rowSOTALLO2.Item(COLUMN_NAME) = rowSOTALLOZ.Item(COLUMN_NAME)
            Next
        Next

        Dim found_one As Boolean = False
        For Each GROW As UltraWinGrid.UltraGridRow In grdSOTALLO1.Rows
            If Format(GROW.Cells("DATE_START").Value, "yyyyMMdd") < Format(Now, "yyyyMMdd") And
                Format(GROW.Cells("DATE_END").Value, "yyyyMMdd") < Format(Now, "yyyyMMdd") Then
                grdSOTALLO1.ActiveRow = GROW
                found_one = True
                Exit For
            End If
        Next
        If Not found_one Then Sort_grdColumns(grdSOTALLO1, "DATE_START".ToLower)

        grdSOTALLO2.Text = "Allocations for Item " & ITEM_CODE
        Sort_grdColumns(grdSOTALLO1, "DATE_START".ToLower)

        Fill_Records("BMTMAIN1", ITEM_CODE)
        Fill_Records("BMTMAIN3", ITEM_CODE)
        Sort_grdColumns(grdBMTMAIN1, "BM_PROD_ITEM", False, 0)
        Sort_grdColumns(grdBMTMAIN1, "BM_COMP_ITEM", False, 1)

        Fill_Records("POTORDR9", ITEM_CODE)
        Sort_grdColumns(grdPOTORDR9, "ITEM_CODE")
        Setup_grdPOTORDR9()

        Fill_Records("ICTPINV1", ITEM_CODE)

        For Each row As DataRow In dst.Tables("POTORDRX").Select("")
            Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
            Dim PO_ORDER_LNO As Integer = Val(row.Item("PO_ORDER_LNO") & "")

            Dim INV_QTY As Integer = Val(dst.Tables("ICTPINV1").Compute("SUM(INV_QTY)", $"PO_ORDER_NO = '{PO_ORDER_NO}' AND PO_ORDER_LNO = {CStr(PO_ORDER_LNO)}") & "")
            Dim PO_QTY_OPN As Integer = Val(row.Item("PO_QTY_OPN") & "")

            Dim rowPOTORDRX As DataRow = dst.Tables("POTORDRX").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})
            rowPOTORDRX.Item("INV_QTY") = Val(rowPOTORDRX.Item("INV_QTY") & "") + INV_QTY

            If PO_QTY_OPN <> INV_QTY Then
                Dim rowICTPINV1 As DataRow = dst.Tables("ICTPINV1").NewRow
                With rowICTPINV1
                    .Item("PO_ORDER_NO") = row.Item("PO_ORDER_NO")
                    .Item("PO_ORDER_LNO") = row.Item("PO_ORDER_LNO")
                    .Item("CONTAINER_NO") = "Qty Open"
                    .Item("INV_NUM") = "Not Inv"
                    .Item("ETA_DATE") = row.Item("PO_DATE_REQUIRED")
                    .Item("OPO_QTY") = PO_QTY_OPN - INV_QTY
                    .Item("WHSE_CODE") = row.Item("WHSE_CODE")
                End With
                dst.Tables("ICTPINV1").Rows.Add(rowICTPINV1)
            End If
        Next

        For Each row As DataRow In dst.Tables("ICTPINV1").Select("CONTAINER_NO IS NULL")
            With row
                Dim WHSE_CODE As String = row.Item("WHSE_CODE") & ""
                ' ISSUE-7230 Clarins to ADS
                'If ASCMAIN1.CLIENT = "INT" Then
                '    If WHSE_CODE <> "CLA" Or WHSE_CODE <> "ADS" Then
                '        WHSE_CODE = "CLA"
                '    End If
                'End If

                If WHSE_CODE.Length = 0 Then
                    WHSE_CODE = ROWs("POTPARM1").Item("PO_PARM_WHSE_CODE")
                End If

                Dim PORT_CODE As String = row.Item("PORT_CODE") & ""
                If PORT_CODE = "" Then PORT_CODE = PO_PARM_PINV_PORT
                Dim rowICTPORT2 As DataRow = dst.Tables("ICTPORT2").Rows.Find(New String() {PORT_CODE, WHSE_CODE})
                Dim PINV_LT As Int32 = PO_PARM_PINV_LT
                If rowICTPORT2 IsNot Nothing Then
                    PINV_LT = Val(rowICTPORT2.Item("ETD_TO_ETA") & "")
                End If
                Dim INV_DATE As Date = CDate(row.Item("INV_DATE"))
                .Item("CONTAINER_NO") = $"Inv Date +{CStr(PINV_LT)}"
                .Item("ETA_DATE") = INV_DATE.AddDays(PINV_LT)
            End With
        Next

        For Each ROW As DataRow In dst.Tables("ICTPINV1").Select("")
            If ROW.Item("ETA_DATE") & "" <> "" Then
                Dim ETA_DATE As Date = ROW.Item("ETA_DATE")
                Dim ETA_DATE_DC As Date = ETA_DATE
                For I As Integer = 1 To 5
                    ETA_DATE_DC = ETA_DATE_DC.AddDays(1)
                    If ETA_DATE_DC.DayOfWeek = DayOfWeek.Saturday Or ETA_DATE_DC.DayOfWeek = DayOfWeek.Sunday Then
                        I = I - 1
                    End If
                Next
                ROW.Item("ETA_DATE_DC") = ETA_DATE_DC
            End If
        Next

        ' GS does not want to see green rows in this grid
        Dim dvw As DataView = DirectCast(grdICTPINV1.DataSource, DataTable).DefaultView
        dvw.RowFilter = "INV_NUM <> 'Not Inv'"

        Sort_grdColumns(grdICTPINV1, "PO_ORDER_NO".ToLower & "," & "INV_DATE".ToLower)
        EnforceConstraints(True)

        Dim IMAGE_NAME As String = ITEM_CODE

        Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
        Dim imgba() As Byte = Nothing
        picItemImage.Image = ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, False, , , imgba)
        UltraExplorerBar1.Groups("Item Image").Text = "Item Image " & ITEM_CODE

        Call ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Call BeginTrans()
        Stop
        Call CommitTrans("Update Complete")

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

    Public Overrides Function Remote_Control(
    ByVal command As String,
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
        Load_Popup_Menu(grdICTTRANX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdICTSTATP, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTORDRX, "SS", "Show Filter", "Show GroupBox", "Sales Order Inquiry", "Sales Order Entry")

        Load_Popup_Menu(grdICTSTATI, "SSB", "Show Filter", "Show GroupBox", "Demand Planning")
        Load_Popup_Menu(grdPOTORDRX, "BB", "PO Inquiry", "Show All POs")
        Load_Popup_Menu(grdICTITEM1_Recent, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Item Status Inquiry")
        Load_Popup_Menu(grdICTITEM1_optP, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Item Status Inquiry")
        Load_Popup_Menu(grdPOTORDR9, "B", "PO Inquiry")
        Load_Popup_Menu(grdSOTINVHX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdICTPINV1, "SSB", "Show Filter", "Show GroupBox", "PO Inquiry")
        If ASCMAIN1.Running_in_VS And ASCMAIN1.CLIENT = "GAB" Then
            Load_Popup_Menu(grdICTSTATS, "SS", "Show Filter", "Show GroupBox")
        End If
        Load_Popup_Menu(grdICTPROD1, "BBBB", "Select All", "Deselect All", "Select Selected", "Deselect Selected")
        Load_Popup_Menu(grdICTWHSE1, "BBBB", "Select All", "Deselect All", "Select Selected", "Deselect Selected")

    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        If Not GRDs.Keys.Contains(Mid(e.SourceControl.Name, 4)) Then
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

            Case "Show All POs"
                EnforceConstraints(False)
                Fill_Records("POTORDRX", New String() {ITEM_CODE, "A"})
                Fill_Records("POTORDR8", {ITEM_CODE, "A"})
                EnforceConstraints(True)
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Track Shipment"
            '    If grd.ActiveRow.Cells("SHIP_REF").Text <> "" Then
            '        Me.Cursor = Cursors.WaitCursor
            '        Call ASCMAIN1.Progress("Now Locating DHL POD")
            '        System.Diagnostics.Process.Start("http://track.dhl-usa.com/TrackByNbr.asp?ShipmentNumber=" & grd.ActiveRow.Cells("SHIP_REF").Text)
            '        Me.Cursor = Cursors.Default
            '        Call ASCMAIN1.Progress("")
            '    End If

            Case "Demand Planning"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                Dim STYLE_CODE As String = rowICTITEM1.Item("STYLE_CODE") & ""
                If STYLE_CODE <> "" Then
                    Context_Launch("Load", STYLE_CODE, e.Tool.Key, "DPFPLAN1")
                End If

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Value & ""
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI")


            Case "Sales Order Inquiry"
                If grd.ActiveRow Is Nothing _
                    OrElse grd.ActiveRow.IsAddRow _
                    OrElse Not grd.ActiveRow.IsDataRow _
                    OrElse grd.ActiveRow.Cells Is Nothing _
                    OrElse Not grd.ActiveRow.Cells.Exists("ORDR_NO") Then Exit Sub

                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value & ""
                If ORDR_NO = "" Then Exit Sub

                Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")

            Case "Sales Order Entry"
                If grd.ActiveRow Is Nothing _
                    OrElse grd.ActiveRow.IsAddRow _
                    OrElse Not grd.ActiveRow.IsDataRow _
                    OrElse grd.ActiveRow.Cells Is Nothing _
                    OrElse Not grd.ActiveRow.Cells.Exists("ORDR_NO") Then Exit Sub

                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value & ""
                If ORDR_NO = "" Then Exit Sub

                Context_Launch("Edit", ORDR_NO, e.Tool.Key, "SOFORDR1")

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Value & ""
                Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")

            Case "Select All", "Deselect All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsDataRow Then
                        grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", 1, 0)
                        grow.Update()
                    End If
                Next

            Case "Select Selected", "Deselect Selected"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsDataRow AndAlso grow.Selected Then
                        grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select Selected", 1, 0)
                        grow.Update()
                    End If
                Next
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "ITEM_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If

            Case "ITEM_UPC_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If

            Case "ITEM_EAN_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "ITEM_CODE"
                Click_Command("View")
            Case "ITEM_UPC_CODE"
                Click_Command("View")
            Case "ITEM_EAN_CODE"
                Click_Command("View")
        End Select
    End Sub
#End Region

    Sub Setup_grdICTSTATX()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")
        optTD.Visible = Not chkHistory.Checked

        grdICTSTATX.DisplayLayout.Bands(0).Columns("OPS_YYYYPP").Hidden = Not chkHistory.Checked
        grdICTSTATX.DisplayLayout.Bands(0).Columns("WHSE_CODE").Hidden = chkHistory.Checked

        Dim sqlF As String = ""
        Dim sqlW As String = ""
        Dim sqlG As String = ""
        If chkHistory.Checked Then
            sqlF = "OPS_YYYYPP, ITEM_CODE, '000' WHSE_CODE"
            sqlW = ""
            sqlG = "OPS_YYYYPP, ITEM_CODE"
            grdICTSTATX.Text = "Item Status / Activity Summary by Period"
            splMain.Panel2Collapsed = True
        Else
            If optTD.Value = "YTD" Then
                sqlF = "'000000' OPS_YYYYPP, ITEM_CODE, WHSE_CODE"
                sqlW = " where OPS_YYYYPP >= '" & Mid(RYP, 1, 4) & "01'" & " and OPS_YYYYPP <= '" & RYP & "'"
                sqlG = "ITEM_CODE, WHSE_CODE"
                grdICTSTATX.Text = "Item Status / Activity by Warehouse for " & ASCMAIN1.Get_Legend(RYP) & " (YTD)"
            Else
                sqlF = "OPS_YYYYPP, ITEM_CODE, WHSE_CODE"
                sqlW = " where OPS_YYYYPP = '" & RYP & "'"
                sqlG = "OPS_YYYYPP, ITEM_CODE, WHSE_CODE"
                grdICTSTATX.Text = "Item Status / Activity by Warehouse for " & ASCMAIN1.Get_Legend(RYP)
            End If

            splMain.Panel2Collapsed = False

            Load_ICTTRANX()
        End If

        Dim sql_Stats As String = ""
        If optTD.Value = "YTD" Then
            sql_Stats = "" _
                & ", SUM (DECODE(OPS_YYYYPP,'" & RYP & "', WHSE_QTY_ON_HAND,0)) WHSE_QTY_ON_HAND" & vbCrLf _
                & ", SUM (DECODE(OPS_YYYYPP,'" & RYP & "', WHSE_QTY_ONPO,0)) WHSE_QTY_ONPO" & vbCrLf _
                & ", SUM (DECODE(OPS_YYYYPP,'" & RYP & "', WHSE_QTY_PLAN,0)) WHSE_QTY_PLAN" & vbCrLf _
                & ", SUM (DECODE(OPS_YYYYPP,'" & RYP & "', WHSE_QTY_OPEN,0)) WHSE_QTY_OPEN" & vbCrLf _
                & ", SUM (DECODE(OPS_YYYYPP,'" & RYP & "', WHSE_QTY_PICK,0)) WHSE_QTY_PICK" & vbCrLf _
                & ", SUM (DECODE(OPS_YYYYPP,'" & RYP & "', WHSE_QTY_COMM,0)) WHSE_QTY_COMM" & vbCrLf _
                & ", SUM (DECODE(OPS_YYYYPP,'" & RYP & "', WHSE_QTY_HOLD,0)) WHSE_QTY_HOLD" & vbCrLf
        Else
            sql_Stats = "" _
                & ", SUM (WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND" & vbCrLf _
                & ", SUM (WHSE_QTY_ONPO) WHSE_QTY_ONPO" & vbCrLf _
                & ", SUM (WHSE_QTY_PLAN) WHSE_QTY_PLAN" & vbCrLf _
                & ", SUM (WHSE_QTY_OPEN) WHSE_QTY_OPEN" & vbCrLf _
                & ", SUM (WHSE_QTY_PICK) WHSE_QTY_PICK" & vbCrLf _
                & ", SUM (WHSE_QTY_COMM) WHSE_QTY_COMM" & vbCrLf _
                & ", SUM (WHSE_QTY_HOLD) WHSE_QTY_HOLD" & vbCrLf
        End If

        Dim SQL As String = "" _
            & " Select " & sqlF & vbCrLf _
            & IIf(optTD.Value = "YTD" And Not chkHistory.Checked,
                  ", SUM (DECODE(OPS_YYYYPP,'" & Mid(RYP, 1, 4) & "01" & "', WHSE_QTY_BEG,0)) WHSE_QTY_BEG",
                  ", SUM (WHSE_QTY_BEG) WHSE_QTY_BEG") & vbCrLf _
            & ", SUM (WHSE_QTY_SHP) WHSE_QTY_SHP" & vbCrLf _
            & ", SUM (WHSE_QTY_RTN) WHSE_QTY_RTN" & vbCrLf _
            & ", SUM (WHSE_QTY_REC) WHSE_QTY_REC" & vbCrLf _
            & ", SUM (WHSE_QTY_ADJ) WHSE_QTY_ADJ" & vbCrLf _
            & ", SUM (WHSE_QTY_XFR) WHSE_QTY_XFR" & vbCrLf _
            & ", SUM (WHSE_QTY_CON) WHSE_QTY_CON" & vbCrLf _
            & ", SUM (WHSE_QTY_RTV) WHSE_QTY_RTV" & vbCrLf _
            & ", SUM (WHSE_QTY_PHY) WHSE_QTY_PHY" & vbCrLf _
            & sql_Stats _
            & " from " & ICTSTATX & vbCrLf _
            & sqlW & vbCrLf _
            & " group by " & sqlG
        Fill_Records("ICTSTATX", "", True, SQL)
        Sort_grdColumns(grdICTSTATX, "OPS_YYYYPP".ToLower & ",WHSE_CODE")


        tabDetails.Tabs("Purchase Orders").Visible = (RYP = ASCMAIN1.CYP)
        tabDetails.Tabs("Sales Orders").Visible = (RYP = ASCMAIN1.CYP)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Load_ICTSTATX()

        ASCDATA1.ExecuteSQL("Truncate Table " & ICTSTATX)

        ASCMAIN1.sql = "" _
        & "Insert into " & ICTSTATX _
        & " Select OPS_YYYYPP, ITEM_CODE, WHSE_CODE" _
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
        & "SELECT OPS_YYYYPP, ITEM_CODE, WHSE_CODE" _
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
        & " FROM ICTSTAT1 WHERE ITEM_CODE = '" & HFs("ITEM_CODE") & "'" _
        & " GROUP BY OPS_YYYYPP, ITEM_CODE, WHSE_CODE" _
        & " UNION " _
        & "SELECT OPS_YYYYPP, ITEM_CODE, WHSE_CODE" _
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
        & " FROM ICTSTAT5 WHERE ITEM_CODE = '" & HFs("ITEM_CODE") & "'" _
        & " GROUP BY OPS_YYYYPP, ITEM_CODE, WHSE_CODE" _
        & " UNION " _
        & "SELECT '" & ASCMAIN1.CYP & "' OPS_YYYYPP, ITEM_CODE, WHSE_CODE" _
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
        & " FROM ICTSTAT2 WHERE ITEM_CODE = '" & HFs("ITEM_CODE") & "'" _
        & " GROUP BY ITEM_CODE, WHSE_CODE" _
        & ") group by OPS_YYYYPP, ITEM_CODE, WHSE_CODE"
        ASCDATA1.ExecuteSQL()

    End Sub

    Sub Load_ICTTRANX()

        Try
            ASCDATA1.ExecuteSQL("Truncate Table " & ICTTRANX)
            Setup_tabDetails()
            'Exit Sub

            ASCMAIN1.sql = "Insert into " & ICTTRANX _
            & " SELECT T2.OPS_YYYYPP, T2.ITEM_CODE, T1.WHSE_CODE" _
            & ", T1.ADJ_NO TRAN_NO, T1.ADJ_SOURCE TRAN_SOURCE" _
            & ", T1.INIT_DATE, T1.INIT_OPER" _
            & ", T1.ADJ_DATE TRAN_DATE, 'A' TRAN_TYPE" _
            & ", T2.ADJ_QTY TRAN_QTY, ICTREAS1.REASON_DESC TRAN_NOTE" _
            & " FROM ICTIADJ1 T1,ICTIADJ2 T2, ICTREAS1" _
            & " WHERE T1.ADJ_NO = T2.ADJ_NO" _
            & " and T2.ITEM_CODE = '" & HFs("ITEM_CODE") & "'" _
            & IIf(optTD.Value = "YTD",
                  " and T2.OPS_YYYYPP >= '" & Mid(RYP, 1, 4) & "01' and T2.OPS_YYYYPP <= '" & RYP & "'",
                  " and T2.OPS_YYYYPP = '" & RYP & "'") _
            & " and ICTREAS1.REASON_CODE = T1.REASON_CODE"
            ASCDATA1.ExecuteSQL()


            ASCMAIN1.sql = "Insert into " & ICTTRANX _
                & " SELECT T1.OPS_YYYYPP, T4.ITEM_CODE, T2.VEND_WHSE_CODE WHSE_CODE" _
                & ", T1.RECEIPT_NO TRAN_NO, 'X' TRAN_SOURCE" _
                & ", T1.INIT_DATE, T1.INIT_OPER" _
                & ", T1.RECEIPT_DATE TRAN_DATE, 'X' TRAN_TYPE" _
                & ", -1 * T4.QTY_CON TRAN_QTY, 'PO ' || T2.PO_ORDER_NO || ', BM# ' || T2.BM_ISSUE_NO TRAN_NOTE" _
                & " FROM ICTIREC1 T1,ICTIREC2 T2,ICTIREC4 T4" _
                & " WHERE T1.RECEIPT_NO = T2.RECEIPT_NO" _
                & " and T4.ITEM_CODE = '" & HFs("ITEM_CODE") & "'" _
                & IIf(optTD.Value = "YTD",
                  " and T1.OPS_YYYYPP >= '" & Mid(RYP, 1, 4) & "01' and T2.OPS_YYYYPP <= '" & RYP & "'",
                  " and T1.OPS_YYYYPP = '" & RYP & "'") _
                & " and T2.RECEIPT_NO = T4.RECEIPT_NO" _
                & " and T2.RECEIPT_LNO = T4.RECEIPT_LNO"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Insert into " & ICTTRANX _
            & " SELECT T2.OPS_YYYYPP, T2.ITEM_CODE, T1.WHSE_CODE" _
            & ", T1.XFR_NO TRAN_NO, T1.XFR_SOURCE TRAN_SOURCE" _
            & ", T1.INIT_DATE, T1.INIT_OPER" _
            & ", T1.XFR_DATE TRAN_DATE, 'T' TRAN_TYPE" _
            & ", -1 * T2.XFR_QTY TRAN_QTY, 'XFR to ' || T1.WHSE_CODE_TO TRAN_NOTE" _
            & " FROM ICTIXFR1 T1,ICTIXFR2 T2" _
            & " WHERE T1.XFR_NO = T2.XFR_NO" _
            & " and T2.ITEM_CODE = '" & HFs("ITEM_CODE") & "'" _
            & IIf(optTD.Value = "YTD",
                  " and T2.OPS_YYYYPP >= '" & Mid(RYP, 1, 4) & "01' and T2.OPS_YYYYPP <= '" & RYP & "'",
                  " and T2.OPS_YYYYPP = '" & RYP & "'")
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Insert into " & ICTTRANX _
            & " SELECT T1.OPS_YYYYPP, T2.ITEM_CODE, T1.WHSE_CODE" _
            & ", T1.RECEIPT_NO TRAN_NO, 'P' TRAN_SOURCE" _
            & ", T1.INIT_DATE, T1.INIT_OPER" _
            & ", T1.RECEIPT_DATE TRAN_DATE, 'P' TRAN_TYPE" _
            & ", T2.QTY_REC TRAN_QTY, 'PO ' || T1.PO_ORDER_NO TRAN_NOTE" _
            & " FROM ICTIREC1 T1,ICTIREC2 T2" _
            & " WHERE T1.RECEIPT_NO = T2.RECEIPT_NO" _
            & " and T2.ITEM_CODE = '" & HFs("ITEM_CODE") & "'" _
            & IIf(optTD.Value = "YTD",
                  " and T2.OPS_YYYYPP >= '" & Mid(RYP, 1, 4) & "01' and T2.OPS_YYYYPP <= '" & RYP & "'",
                  " and T2.OPS_YYYYPP = '" & RYP & "'")
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Insert into " & ICTTRANX _
            & " SELECT T2.OPS_YYYYPP, T2.ITEM_CODE, T1.WHSE_CODE_TO" _
            & ", T1.XFR_NO TRAN_NO, T1.XFR_SOURCE TRAN_SOURCE" _
            & ", T1.INIT_DATE, T1.INIT_OPER" _
            & ", T1.XFR_DATE TRAN_DATE, 'T' TRAN_TYPE" _
            & ", T2.XFR_QTY TRAN_QTY, 'XFR from ' || T1.WHSE_CODE TRAN_NOTE" _
            & " FROM ICTIXFR1 T1,ICTIXFR2 T2" _
            & " WHERE T1.XFR_NO = T2.XFR_NO" _
            & " and T2.ITEM_CODE = '" & HFs("ITEM_CODE") & "'" _
            & IIf(optTD.Value = "YTD",
                  " and T2.OPS_YYYYPP >= '" & Mid(RYP, 1, 4) & "01' and T2.OPS_YYYYPP <= '" & RYP & "'",
                  " and T2.OPS_YYYYPP = '" & RYP & "'")
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Insert into " & ICTTRANX _
            & " SELECT T2.ORDR_YYYYPP_UPDATED, T2.ITEM_CODE, T1.WHSE_CODE" _
            & ", Decode(T2.INV_TYPE,'I','Shp','C', DECODE(S1.RA_NO, NULL, 'Rtn', 'RA ' || S1.RA_NO),'?') TRAN_NO, 'S' TRAN_SOURCE" _
            & ", NULL INIT_DATE, NULL INIT_OPER" _
            & ", T1.INV_DATE TRAN_DATE, DECODE(T2.INV_TYPE,'I','S','C','R',NULL) TRAN_TYPE" _
            & ", SUM (-1 * T2.ORDR_QTY_SHIP) TRAN_QTY, ' Line Items:' || Count (*) TRAN_NOTE" _
            & " FROM SOTINVH1 T1,SOTINVH2 T2, SOTRTRN1 S1" _
            & " WHERE T1.INV_NO = T2.INV_NO" _
            & " and T1.INV_TYPE = T2.INV_TYPE" _
            & " and T2.ITEM_CODE = '" & HFs("ITEM_CODE") & "'" _
            & " and T1.INV_NO = S1.INV_NO (+)" _
            & IIf(optTD.Value = "YTD",
                  " and T2.ORDR_YYYYPP_UPDATED >= '" & Mid(RYP, 1, 4) & "01' and T2.ORDR_YYYYPP_UPDATED <= '" & RYP & "'",
                  " and T2.ORDR_YYYYPP_UPDATED = '" & RYP & "'") _
            & " GROUP BY T2.ORDR_YYYYPP_UPDATED, T2.ITEM_CODE, T1.WHSE_CODE, Decode(T2.INV_TYPE,'I','Shp','C', DECODE(S1.RA_NO, NULL, 'Rtn', 'RA ' || S1.RA_NO),'?'), T1.INV_DATE, DECODE(T2.INV_TYPE,'I','S','C','R',NULL)"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Insert into " & ICTTRANX _
            & " SELECT T2.ORDR_YYYYPP_UPDATED, T2.ITEM_CODE, T1.WHSE_CODE" _
            & ", T1.INV_NO TRAN_NO, 'S' TRAN_SOURCE" _
            & ", T1.INIT_DATE, T1.INIT_OPER" _
            & ", T1.INV_DATE TRAN_DATE, 'V' TRAN_TYPE" _
            & ", -1 * T2.ORDR_QTY_SHIP TRAN_QTY, T1.ORDR_CUST_PO TRAN_NOTE" _
            & " from SOTINVT1 T1,SOTINVT2 T2" _
            & " where T1.INV_NO = T2.INV_NO" _
            & " and T1.INV_TYPE = T2.INV_TYPE" _
            & " and T1.ORDR_TYPE_CODE = 'RTV'" _
            & " and T2.ITEM_CODE = '" & HFs("ITEM_CODE") & "'" _
            & IIf(optTD.Value = "YTD",
                  " and T2.ORDR_YYYYPP_UPDATED >= '" & Mid(RYP, 1, 4) & "01' and T2.ORDR_YYYYPP_UPDATED <= '" & RYP & "'",
                  " and T2.ORDR_YYYYPP_UPDATED = '" & RYP & "'")
            ASCDATA1.ExecuteSQL()

            dst.Tables("ICTTRANX").Rows.Clear()
            grdICTTRANX.Text = "No Transactions"

            Setup_tabDetails()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Please Call ABS - Problem mapping Transactions")
        End Try

    End Sub

    Sub Setup_grdICTTRANX()

        Dim WHSE_CODE As String = grdICTSTATX.ActiveRow.Cells("WHSE_CODE").Text

        Dim sql As String = "Select * from " & ICTTRANX _
        & " where WHSE_CODE = '" & WHSE_CODE & "'"

        Fill_Records("ICTTRANX", "", True, sql)
        Sort_grdColumns(grdICTTRANX, "INIT_DATE,TRAN_DATE,TRAN_TYPE,TRAN_NO")

        grdICTTRANX.Text = "Transaction Details for Whse " & WHSE_CODE
    End Sub

    Sub Setup_grdSOTINVHX()
        Fill_Records("SOTINVHX", New String() {ITEM_CODE, RYP})
        Sort_grdColumns(grdSOTINVHX, "INV_NO")
        grdSOTINVHX.Text = "Sales Shipments / Returns for Item " & ITEM_CODE & " for " & RYP
    End Sub

    Sub Setup_grdPOTORDR9()
        If grdICTSTATX.ActiveRow Is Nothing OrElse (Not grdICTSTATX.ActiveRow.IsDataRow Or grdICTSTATX.ActiveRow.IsAddRow) Then
            grdPOTORDR9.Visible = False
        Else
            Dim dvw As DataView = DirectCast(grdPOTORDR9.DataSource, DataTable).DefaultView
            Dim VEND_WHSE_CODE As String = grdICTSTATX.ActiveRow.Cells("WHSE_CODE").Value
            dvw.RowFilter = "VEND_WHSE_CODE = '" & VEND_WHSE_CODE & "'"
            Sort_grdColumns(grdPOTORDR9, "ITEM_CODE")
            grdPOTORDR9.Text = "Component Commitments in Whse " & VEND_WHSE_CODE
            grdPOTORDR9.Visible = True
        End If
    End Sub

    Private Sub grdICTSTATX_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTSTATX.AfterRowActivate
        If chkHistory.Checked Then Exit Sub
        Setup_grdICTTRANX()
        Setup_grdSOTINVHX()
        Setup_grdPOTORDR9()
    End Sub

    Private Sub grdICTSTATX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTSTATX.DoubleClickRow

        If chkHistory.Checked Then
            RYP = e.Row.Cells("OPS_YYYYPP").Text
            chkHistory.Checked = False
            Setup_grdSOTINVHX()
        End If

    End Sub

    Private Sub chkHistory_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkHistory.CheckedChanged
        Setup_grdICTSTATX()
    End Sub


    Private Sub grdICTTRANX_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTTRANX.InitializeRow
        Dim TRAN_TYPE As String = e.Row.Cells("TRAN_TYPE").Text
        Dim C As System.Drawing.Color
        Select Case TRAN_TYPE
            Case "A"
                C = Color.LightBlue
            Case "T"
                C = Color.LightPink
            Case "R"
                C = Color.LightSalmon
        End Select

        e.Row.Cells("TRAN_TYPE").Appearance.BackColor = C
    End Sub

    Private Sub tabDetails_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabDetails.SelectedTabChanged
        Setup_tabDetails()
    End Sub

    Sub Setup_tabDetails()
        If SELECTION_NO = 0 Then Exit Sub
        'UltraExplorerBar1.Groups("Shipments / Returns").Visible = (tabDetails.SelectedTab.Key = "Shipments / Returns") And ScreenMode
    End Sub

    Sub Refresh_Summary()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Fetching Inventory Summary")

        ASCDATA1.ExecuteSQL("Truncate Table " & ICTSTATI)

        Dim sql As String = sqlICTSTATI
        If Not chkCOLLECTION_CODE.Checked Then
            sql = Replace(sql, ", ICTITEM1.COLLECTION_CODE", ", NULL COLLECTION_CODE", 1, 1)
            sql = Replace(sql, ", ICTITEM1.COLLECTION_CODE", "")

        End If
        If Not chkITEM_CATGY_CODE.Checked Then
            sql = Replace(sql, ", ICTITEM1.ITEM_CATGY_CODE", ", NULL ITEM_CATGY_CODE", 1, 1)
            sql = Replace(sql, ", ICTITEM1.ITEM_CATGY_CODE", "")
        End If
        If Not chkDEPT_CODE.Checked Then
            sql = Replace(sql, ", ICTITEM1.DEPT_CODE", ", NULL DEPT_CODE", 1, 1)
            sql = Replace(sql, ", ICTITEM1.DEPT_CODE", "")
        End If
        If Not chkWHSE_CODE.Checked Then
            sql = Replace(sql, ", ICTSTAT2.WHSE_CODE", ", NULL WHSE_CODE", 1, 1)
            sql = Replace(sql, ", ICTSTAT2.WHSE_CODE", "")
        End If

        With grdICTSTATP.DisplayLayout.Bands(0)
            .Columns("COLLECTION_CODE").Hidden = Not chkCOLLECTION_CODE.Checked
            .Columns("ITEM_CATGY_CODE").Hidden = Not chkITEM_CATGY_CODE.Checked
            .Columns("DEPT_CODE").Hidden = Not chkDEPT_CODE.Checked
            .Columns("WHSE_CODE").Hidden = Not chkWHSE_CODE.Checked

            .Columns("COLLECTION_DESC").Hidden = Not chkCOLLECTION_CODE.Checked
            .Columns("ITEM_CATGY_DESC").Hidden = Not chkITEM_CATGY_CODE.Checked
            .Columns("DEPT_DESC").Hidden = Not chkDEPT_CODE.Checked

            .Columns("WHSE_DESC").Hidden = Not chkWHSE_CODE.Checked
            .Columns("WHSE_TYPE").Hidden = Not chkWHSE_CODE.Checked
        End With

        ASCDATA1.ExecuteSQL("Insert into " & ICTSTATI & " " & sql)
        ASCDATA1.ExecuteSQL("Truncate Table " & ICTSTATP)
        ASCDATA1.ExecuteSQL("Insert into " & ICTSTATP & " (COLLECTION_CODE,ITEM_CATGY_CODE,DEPT_CODE,WHSE_CODE,ITEMS,QOH,ONPO,PLAN,OPEN,PICK,COMM,HOLD) " & sqlICTSTATP)

        'If chkWHSE_CODE.Checked Then
        '    ASCMAIN1.sql = "Update " & ICTSTATP & " ICTSTATP " _
        '    & " Set CODE_DESC = (Select WHSE_DESC from ICTWHSE1 where WHSE_CODE = ICTSTATP.WHSE_CODE)"
        '    ASCMAIN1.sql = "Update " & ICTSTATP & " ICTSTATP " _
        '    & " Set CODE_TYPE = (Select WHSE_TYPE from ICTWHSE1 where WHSE_CODE = ICTSTATP.WHSE_CODE)"
        '    ASCDATA1.ExecuteSQL()
        'End If

        EnforceConstraints(False)

        dst.Tables("ICTSTATP").Columns("COLLECTION_DESC").Expression = ""
        dst.Tables("ICTSTATP").Columns("ITEM_CATGY_DESC").Expression = ""
        dst.Tables("ICTSTATP").Columns("DEPT_DESC").Expression = ""
        dst.Tables("ICTSTATP").Columns("WHSE_DESC").Expression = ""
        dst.Tables("ICTSTATP").Columns("WHSE_TYPE").Expression = ""
        For Each RELATION_NAME As String In New String() _
        {"ICTCOLL1_ICTSTATP", "ICTCATG1_ICTSTATP", "ICTDEPT1_ICTSTATP", "ICTWHSE1_ICTSTATP"}
            If dst.Relations.Contains(RELATION_NAME) Then
                dst.Relations.Remove(RELATION_NAME)
            End If
        Next

        Fill_Records("ICTSTATP")
        Sort_grdColumns(grdICTSTATP, "COLLECTION_CODE,ITEM_CATGY_CODE,DEPT_CODE,WHSE_CODE")

        If chkCOLLECTION_CODE.Checked Then
            dst.Relations.Add("ICTCOLL1_ICTSTATP" _
                              , New DataColumn() {dst.Tables("ICTCOLL1").Columns("COLLECTION_CODE")} _
                              , New DataColumn() {dst.Tables("ICTSTATP").Columns("COLLECTION_CODE")})
            dst.Tables("ICTSTATP").Columns("COLLECTION_DESC").Expression = "PARENT(ICTCOLL1_ICTSTATP).COLLECTION_NAME"
        End If

        If chkITEM_CATGY_CODE.Checked Then
            dst.Relations.Add("ICTCATG1_ICTSTATP" _
                              , New DataColumn() {dst.Tables("ICTCATG1").Columns("ITEM_CATGY_CODE")} _
                              , New DataColumn() {dst.Tables("ICTSTATP").Columns("ITEM_CATGY_CODE")})
            dst.Tables("ICTSTATP").Columns("ITEM_CATGY_DESC").Expression = "PARENT(ICTCATG1_ICTSTATP).ITEM_CATGY_DESC"
        End If

        If chkDEPT_CODE.Checked Then
            dst.Relations.Add("ICTDEPT1_ICTSTATP" _
                              , New DataColumn() {dst.Tables("ICTDEPT1").Columns("DEPT_CODE")} _
                              , New DataColumn() {dst.Tables("ICTSTATP").Columns("DEPT_CODE")})
            dst.Tables("ICTSTATP").Columns("DEPT_DESC").Expression = "PARENT(ICTDEPT1_ICTSTATP).DEPT_DESC"
        End If

        If chkWHSE_CODE.Checked Then
            dst.Relations.Add("ICTWHSE1_ICTSTATP" _
                              , New DataColumn() {dst.Tables("ICTWHSE1").Columns("WHSE_CODE")} _
                              , New DataColumn() {dst.Tables("ICTSTATP").Columns("WHSE_CODE")})
            dst.Tables("ICTSTATP").Columns("WHSE_DESC").Expression = "PARENT(ICTWHSE1_ICTSTATP).WHSE_DESC"
            dst.Tables("ICTSTATP").Columns("WHSE_TYPE").Expression = "PARENT(ICTWHSE1_ICTSTATP).WHSE_TYPE"
        End If

        EnforceConstraints(True)

        grdICTSTATI.Visible = True
        grdICTSTATP.Visible = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub


    Private Sub grdICTSTATP_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTSTATP.AfterRowActivate
        Setup_grdICTSTATI()
    End Sub

    Sub Setup_grdICTSTATI()
        With grdICTSTATP
            If .ActiveRow Is Nothing OrElse Not .ActiveRow.IsDataRow Then
                grdICTSTATI.Visible = False
                Exit Sub
            End If

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Fetching Items")

            grdICTSTATI.Visible = True

            Dim COLLECTION_CODE As String = .ActiveRow.Cells("COLLECTION_CODE").Value & ""
            Dim ITEM_CATGY_CODE As String = .ActiveRow.Cells("ITEM_CATGY_CODE").Value & ""
            Dim DEPT_CODE As String = .ActiveRow.Cells("DEPT_CODE").Value & ""
            Dim WHSE_CODE As String = .ActiveRow.Cells("WHSE_CODE").Value & ""
            'Fill_Records("ICTSTATI", New String() {COLLECTION_CODE, ITEM_CATGY_CODE, DEPT_CODE, WHSE_CODE})

            ASCMAIN1.sql = "Select ICTSTATI.*,ICTITEM1.ITEM_DESC" & vbCrLf _
            & " from " & ICTSTATI & " ICTSTATI,ICTITEM1 " & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = ICTSTATI.ITEM_CODE" & vbCrLf _
            & IIf(chkCOLLECTION_CODE.Checked, " and (ICTSTATI.COLLECTION_CODE " & IIf(COLLECTION_CODE = "", " is Null", "= '" & COLLECTION_CODE & "'") & ")" & vbCrLf, "") _
            & IIf(chkITEM_CATGY_CODE.Checked, " and (ICTSTATI.ITEM_CATGY_CODE " & IIf(ITEM_CATGY_CODE = "", " is Null", "= '" & ITEM_CATGY_CODE & "'") & ")" & vbCrLf, "") _
            & IIf(chkDEPT_CODE.Checked, " and (ICTSTATI.DEPT_CODE " & IIf(DEPT_CODE = "", " is Null", "= '" & DEPT_CODE & "'") & ")" & vbCrLf, "") _
            & IIf(chkWHSE_CODE.Checked, " and (ICTSTATI.WHSE_CODE " & IIf(WHSE_CODE = "", " is Null", "= '" & WHSE_CODE & "'") & ")" & vbCrLf, "")
            Fill_Records("ICTSTATI", , , ASCMAIN1.sql)

            Dim xdesc As String = ""
            If chkCOLLECTION_CODE.Checked Then xdesc &= ",Collection " & COLLECTION_CODE
            If chkITEM_CATGY_CODE.Checked Then xdesc &= ",Category " & ITEM_CATGY_CODE
            If chkDEPT_CODE.Checked Then xdesc &= ",Dept " & DEPT_CODE
            If chkWHSE_CODE.Checked Then xdesc &= ",Whse " & WHSE_CODE

            grdICTSTATI.Text = "Items in " & Mid(xdesc, 2)
            Sort_grdColumns(grdICTSTATI, "ITEM_CODE")

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End With

    End Sub

    Private Sub grdICTSTATI_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTSTATI.DoubleClickRow
        If grdICTSTATI.ActiveRow IsNot Nothing AndAlso grdICTSTATI.ActiveRow.IsDataRow Then
            Absx1.txtFor("ITEM_CODE").Text = grdICTSTATI.ActiveRow.Cells("ITEM_CODE").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub optViewItems_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optViewItems.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_cbeWHSE_CODE()
        If optViewItems.Value <> "S" Then Setup_Recent()
        Setup_ViewItems()
        If optViewItems.Value = "P" Then
            UltraExplorerBar1.Groups("Excluded Items").Visible = True
            UltraExplorerBar1.Groups("Warehouses").Visible = True
            UltraExplorerBar1.Groups("Prod Codes").Visible = True
            Dim sqlExcl As String = "SELECT ITEM_CODE FROM ICTEXCL1 ORDER BY ITEM_CODE"
            Dim DT As DataTable = ASCDATA1.GetDataTable(sqlExcl)

            If DT.Rows.Count > 0 Then
                Dim excludedItems As String = String.Join(vbCrLf, DT.AsEnumerable().Select(Function(row) row.Field(Of String)("ITEM_CODE")))
                txtExcluded.Text = excludedItems
            Else
                txtExcluded.Text = ""
            End If
            splGridContainer.Panel1Collapsed = True
            splGridContainer.Panel2Collapsed = False
            For Each row As UltraGridRow In grdICTWHSE1.Rows
                If row.Cells("WHSE_CODE").Value.ToString() = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE") Then
                    row.Cells("SEL").Value = 1
                    row.Update()
                End If
            Next

            For Each row As UltraGridRow In grdICTPROD1.Rows
                If row.Cells("PROD_CODE").Value.ToString() = "SB" Then
                    row.Cells("SEL").Value = 1
                    row.Update()
                End If
            Next

        Else
            UltraExplorerBar1.Groups("Excluded Items").Visible = False
            UltraExplorerBar1.Groups("Warehouses").Visible = False
            UltraExplorerBar1.Groups("Prod Codes").Visible = False
            grdICTITEM1_Recent.Visible = True
            grdICTITEM1_optP.Visible = False
            grdICTITEM1_Recent.BringToFront()
            splGridContainer.Panel1Collapsed = False
            splGridContainer.Panel2Collapsed = True
        End If
    End Sub

    Sub Setup_ViewItems()
        UltraExplorerBar1.Groups("Summary By").Visible = Not ScreenMode And (optViewItems.Value = "S")
        splSummary.Visible = (optViewItems.Value = "S")
        grdICTITEM1_Recent.Visible = Not (optViewItems.Value = "S") And Not (optViewItems.Value = "A") And Not (optViewItems.Value = "P")
        grdICTITEM1_optP.Visible = (optViewItems.Value = "P")

        If ASCMAIN1.Running_in_VS And ASCMAIN1.CLIENT = "GAB" Then
            grdICTSTATS.Visible = (optViewItems.Value = "A")
            btnATSRefresh.Visible = (optViewItems.Value = "A")
            If optViewItems.Value = "A" Then
                Refresh_ICTSTATS()
            End If
        End If

    End Sub

    Sub Setup_Recent()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Getting " & optViewItems.Text, "")

        Dim sqlWHSE As String = ""
        If optWHSE_CODE.Value = "I" Then
            sqlWHSE = $" where WHSE_CODE = '{cbeWHSE_CODE.Value}'" & vbCrLf
        End If

        If optViewItems.Value = "V" Then
            grdICTITEM1_Recent.DataSource = dst.Tables("ICTITEM1_RECENT")
        Else
            Dim sqlx As String = "Select ITEM_CODE" _
                                 & ", SUM (WHSE_QTY_ON_HAND) QTY_ONHD" & vbCrLf _
                                 & ", SUM (WHSE_QTY_ONPO) QTY_ONPO" & vbCrLf _
                                 & ", SUM (WHSE_QTY_PLAN) QTY_PLAN" & vbCrLf _
                                 & ", SUM (WHSE_QTY_OPEN) QTY_OPEN" & vbCrLf _
                                 & ", SUM (WHSE_QTY_PICK) QTY_PICK" & vbCrLf _
                                 & ", SUM (WHSE_QTY_COMM) QTY_COMM" & vbCrLf _
                                 & " from ICTSTAT2" & vbCrLf _
                                 & sqlWHSE _
                                 & " group by ITEM_CODE"

            Dim sqlx1 As String = "Select ITEM_CODE" _
                                 & ", SUM (WHSE_QTY_SHP) MTD_SHP" & vbCrLf _
                                 & " from ICTSTAT1" & vbCrLf _
                                 & $" where OPS_YYYYPP = '{ASCMAIN1.CYP}'" & vbCrLf _
                                 & Replace(sqlWHSE, " where ", " and ") _
                                 & " group by ITEM_CODE"

            Dim sqlF As String = "Select DPTITMF1.ITEM_CODE, SUM (DPTITMF1.FORECAST) ITEM_FC from DPTITMF1,SOTMKTC1" & vbCrLf _
                                 & "where SOTMKTC1.MARKET_CODE = DPTITMF1.MARKET_CODE" & vbCrLf _
                                 & $"   and DPTITMF1.OPS_YYYYPP = '{ASCMAIN1.CYP}'" & vbCrLf _
                                 & Replace(sqlWHSE, "where ", "and ") _
                                 & " group by DPTITMF1.ITEM_CODE"

            'Dim sqlP As String = "Select DPTMRPG1.ITEM_CODE, DPTMRPG1.QTY_00 ITEM_POS from DPTMRPG1" & vbCrLf _
            '                     & " where DPTITMF1.MRP_TYPE = '6'" & vbCrLf _
            '                     & " group by DPTITMF1.ITEM_CODE"

            Dim sqlFC_MOS As String = ""
            For I As Integer = 1 To 24
                sqlFC_MOS &= $"+SIGN(ABS(NVL(DPTMRPG1_FC.QTY_{Format(I, "00")}, 0)))"
            Next
            sqlFC_MOS = Mid(sqlFC_MOS, 2)

            Dim SQLW As String = ""
            If optViewItems.Value = "Q" Then
                SQLW = " And (NVL(QTY_ONHD,0) <> 0 Or NVL(QTY_ONPO,0) <> 0 Or NVL(QTY_PLAN,0) <> 0 Or NVL(QTY_OPEN,0) <> 0 Or NVL(QTY_PICK,0) <> 0 Or NVL(QTY_COMM,0) <> 0)"
            Else
                SQLW = " And (NVL(QTY_ONHD,0) + NVL(QTY_ONPO,0) + NVL(QTY_PLAN,0) - NVL(QTY_OPEN,0) - NVL(QTY_PICK,0) - NVL(QTY_COMM,0) < 0)"
            End If
            If optViewItems.Value <> "P" Then
                ASCMAIN1.sql = "Select ICTITEM1.*" & vbCrLf _
                & ", X.QTY_ONHD, X.QTY_ONPO, X.QTY_PLAN, X.QTY_OPEN, X.QTY_PICK, X.QTY_COMM, F.ITEM_FC, DPTMRPG1_POS.QTY_01 ITEM_POS" & vbCrLf _
                & ", X1.MTD_SHP" & vbCrLf _
                & ", DPTMRPG1_POS.QTY_00 ITEM_POS0, DPTMRPG1_POS.QTY_01 ITEM_POS1, DPTMRPG1_POS.QTY_02 ITEM_POS2, DPTMRPG1_POS.QTY_03 ITEM_POS3" & vbCrLf _
                & ", DPTMRPG1_FC.QTY_00 ITEM_FC0, DPTMRPG1_FC.QTY_01 ITEM_FC1, DPTMRPG1_FC.QTY_02 ITEM_FC2, DPTMRPG1_FC.QTY_03 ITEM_FC3" & vbCrLf _
                & ", " & sqlFC_MOS & " ITEM_FC_MOS" & vbCrLf _
                & " from ICTITEM1,(" & sqlx & ") X,(" & sqlx1 & ") X1, (" & sqlF & ") F, DPTMRPG1 DPTMRPG1_POS, DPTMRPG1 DPTMRPG1_FC" & vbCrLf _
                & " where X.ITEM_CODE = ICTITEM1.ITEM_CODE" & vbCrLf _
                & "   and X1.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
                & "   And F.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
                & "   And DPTMRPG1_POS.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
                & "   And DPTMRPG1_POS.MRP_TYPE (+) = '6'" & vbCrLf _
                & "   And DPTMRPG1_FC.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
                & "   And DPTMRPG1_FC.MRP_TYPE (+) = '1'" & vbCrLf _
                & SQLW
                Fill_Records("ICTITEM1_VIEW", "", True, ASCMAIN1.sql)
                grdICTITEM1_Recent.DataSource = dst.Tables("ICTITEM1_VIEW")
                Sort_grdColumns(grdICTITEM1_Recent, "ITEM_CODE")
            End If
        End If
        grdICTITEM1_Recent.Text = optViewItems.Text

        If optWHSE_CODE.Visible Then
            Dim WHSEs As String = ""
            If optWHSE_CODE.Value = "A" Then
                WHSEs = "All Warehouses Combined"
            Else
                WHSEs = "Warehouse " & cbeWHSE_CODE.Value
            End If

            grdICTITEM1_Recent.Text = optViewItems.Text & " - " & WHSEs
        End If

        For Each COLUMN_NAME As String In New String() {"QTY_ONHD", "QTY_ONPO", "QTY_PLAN", "QTY_OPEN", "QTY_PICK", "QTY_COMM", "QTY_NETA"}
            grdICTITEM1_Recent.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = (optViewItems.Value = "V")
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Private Sub grdICTITEM1_Recent_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTITEM1_Recent.InitializeRow
        If e.Row.IsDataRow Then
            If e.Row.Cells("QTY_NETA").Value >= 0 Then
                e.Row.Cells("QTY_NETA").Appearance.ForeColor = Color.Green
            Else
                e.Row.Cells("QTY_NETA").Appearance.ForeColor = Color.Red
            End If

            For Each C As String In New String() {"ITEM_POS", "ITEM_POS0", "ITEM_POS1", "ITEM_POS2", "ITEM_POS3"}
                If Val(e.Row.Cells(C).Value & "") < Val(e.Row.Cells("ITEM_POS_MIN").Value & "") Then
                    e.Row.Cells(C).Appearance.ForeColor = Color.Red
                Else
                    e.Row.Cells(C).Appearance.ForeColor = Color.Empty
                End If
            Next

        End If
    End Sub

    Private Sub grdICTITEM1_Recent_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTITEM1_Recent.DoubleClickRow
        If e.Row.IsDataRow Then
            Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Value
            Absx1.txtFor("ITEM_CODE").Text = ITEM_CODE
            Click_Command("View")
        End If
    End Sub

    Private Sub grdSOTALLO1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTALLO1.AfterRowActivate
        Setup_SOTALLO2()
    End Sub

    Sub Setup_SOTALLO2()
        If grdSOTALLO1.ActiveRow Is Nothing OrElse Not grdSOTALLO1.ActiveRow.IsDataRow Then
            grdSOTALLO2.Visible = False
        Else
            grdSOTALLO2.Visible = True
            Dim dvw As DataView = DirectCast(grdSOTALLO2.DataSource, DataTable).DefaultView
            Dim ALLO_CTL_NO As String = grdSOTALLO1.ActiveRow.Cells("ALLO_CTL_NO").Value
            dvw.RowFilter = "ALLO_CTL_NO = '" & ALLO_CTL_NO & "'"
            grdSOTALLO2.Text = "Allocation " & ALLO_CTL_NO
            Sort_grdColumns(grdSOTALLO2, "CUST_CODE")
        End If
    End Sub


    Private Sub grdBMTMAIN1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdBMTMAIN1.InitializeRow

        If e.Row.Band.Index = 1 Then
            If e.Row.Cells("BM_COMP_ITEM").Value = MyBase.Absx1.txtFor("ITEM_CODE").Text Then
                e.Row.Appearance.BackColor = Color.LightBlue
            End If
        End If
    End Sub

    Private Sub optTD_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optTD.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_grdICTSTATX()
    End Sub

    Sub Integrity_Check()
        ASCMAIN1.sql = "Select ITEM_CODE, WHSE_CODE" & vbCrLf _
            & ", SUM (PO_DTL) PO_DTL, SUM (PO_SUM) PO_SUM" & vbCrLf _
            & ", SUM (PP_DTL) PP_DTL, SUM (PP_SUM) PP_SUM" & vbCrLf _
            & ", SUM (SO_DTL) SO_DTL, SUM (SO_SUM) SO_SUM" & vbCrLf _
            & ", SUM (SP_DTL) SP_DTL, SUM (SP_SUM) SP_SUM" & vbCrLf _
            & ", SUM (PC_DTL) PC_DTL, SUM (PC_SUM) PC_SUM" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select 'IC' TYPE, ITEM_CODE, WHSE_CODE" & vbCrLf _
            & ", 0 PO_DTL, SUM (WHSE_QTY_ONPO) PO_SUM" & vbCrLf _
            & ", 0 PP_DTL, SUM (WHSE_QTY_PLAN) PP_SUM" & vbCrLf _
            & ", 0 SO_DTL, SUM (WHSE_QTY_OPEN) SO_SUM" & vbCrLf _
            & ", 0 SP_DTL, SUM (WHSE_QTY_PICK) SP_SUM" & vbCrLf _
            & ", 0 PC_DTL, SUM (WHSE_QTY_COMM) PC_SUM" & vbCrLf _
            & " from ICTSTAT2" & vbCrLf _
            & " group by ITEM_CODE, WHSE_CODE" & vbCrLf _
            & " union" & vbCrLf _
            & "Select 'PO' TYPE, POTORDR2.ITEM_CODE, POTORDR2.WHSE_CODE" & vbCrLf _
            & ", SUM (POTORDR2.PO_QTY_OPN) PO_DTL, 0 PO_SUM" & vbCrLf _
            & ", 0 PP_DTL, 0 PP_SUM" & vbCrLf _
            & ", 0 SO_DTL, 0 SO_SUM" & vbCrLf _
            & ", 0 SP_DTL, 0 SP_SUM" & vbCrLf _
            & ", 0 PC_DTL, 0 PC_SUM" & vbCrLf _
            & " from POTORDR2,POTORDR1 where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & " group by POTORDR2.ITEM_CODE, POTORDR2.WHSE_CODE" & vbCrLf _
            & " union" & vbCrLf _
            & "Select 'PP' TYPE, DPTPLAN1.ITEM_CODE, DPTPLAN1.TO_WHSE WHSE_CODE" & vbCrLf _
            & ", 0 PO_DTL, 0 PO_SUM" & vbCrLf _
            & ", SUM (0 * DPTPLAN1.QTY_PLANNED) PP_DTL, 0 PP_SUM" & vbCrLf _
            & ", 0 SO_DTL, 0 SO_SUM" & vbCrLf _
            & ", 0 SP_DTL, 0 SP_SUM" & vbCrLf _
            & ", 0 PC_DTL, 0 PC_SUM" & vbCrLf _
            & " from DPTPLAN1" & vbCrLf _
            & " group by DPTPLAN1.ITEM_CODE, DPTPLAN1.TO_WHSE" & vbCrLf _
            & " union" & vbCrLf _
            & "Select 'SO' TYPE, SOTORDR2.ITEM_CODE, SOTORDR1.WHSE_CODE" & vbCrLf _
            & ", 0 PO_DTL, 0 PO_SUM" & vbCrLf _
            & ", 0 PP_DTL, 0 PP_SUM" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) SO_DTL, 0 SO_SUM" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK) SP_DTL, 0 SP_SUM" & vbCrLf _
            & ", 0 PC_DTL, 0 PC_SUM" & vbCrLf _
            & " from SOTORDR2,SOTORDR1 where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "  and SOTORDR2.ORDR_STATUS <> 'D' and SOTORDR2.ORDR_STATUS <> 'C'" & vbCrLf _
            & " group by SOTORDR2.ITEM_CODE, SOTORDR1.WHSE_CODE" & vbCrLf _
            & " union" & vbCrLf _
            & "Select 'PC' TYPE, POTORDR9.ITEM_CODE, DECODE(POTORDR9.PO_ORDER_LNO,0,DPTPLAN1.AT_WHSE,POTORDR1.VEND_WHSE_CODE) WHSE_CODE" & vbCrLf _
            & ", 0 PO_DTL, 0 PO_SUM" & vbCrLf _
            & ", 0 PP_DTL, 0 PP_SUM" & vbCrLf _
            & ", 0 SO_DTL, 0 SO_SUM" & vbCrLf _
            & ", 0 SP_DTL, 0 SP_SUM" & vbCrLf _
            & ", SUM (POTORDR9.PO_QTY_COM) PP_DTL, 0 PP_SUM" & vbCrLf _
            & " from POTORDR9,POTORDR1,DPTPLAN1" & vbCrLf _
            & " where POTORDR1.PO_ORDER_NO (+) = POTORDR9.PO_ORDER_NO and DPTPLAN1.PLAN_NO (+) = POTORDR9.PO_ORDER_NO" & vbCrLf _
            & " group by POTORDR9.ITEM_CODE, DECODE(POTORDR9.PO_ORDER_LNO,0,DPTPLAN1.AT_WHSE,POTORDR1.VEND_WHSE_CODE)" & vbCrLf _
            & ")" & vbCrLf _
            & " group by ITEM_CODE, WHSE_CODE" & vbCrLf _
            & "having SUM (PO_DTL) <> SUM (PO_SUM) or SUM (PP_DTL) <> SUM (PP_SUM) or SUM (SO_DTL) <> SUM (SO_SUM) or SUM (SP_DTL) <> SUM (SP_SUM) or SUM (PC_DTL) <> SUM (PC_SUM)"

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Checking Status")

        Dim ICTSTATO As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ICTSTATO & " Add Primary Key (ITEM_CODE, WHSE_CODE)")

        'ASCMAIN1.sql = "Update " & ICTSTATO & " Set PP_SUM = PP_DTL"
        'ASCDATA1.ExecuteSQL()

        Fill_Records("ICTSTATO", "", True, "Select * from " & ICTSTATO)
        Sort_grdColumns(grdICTSTATO, "ITEM_CODE,WHSE_CODE")

        Dim TBL As DataTable = ASCDATA1.GetDataTable

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")


        If dst.Tables("ICTSTATO").Rows.Count = 0 Then
            MsgBox("All Items are in Balance", MsgBoxStyle.OkOnly, "Success")
            grdICTSTATO.Visible = False
            tabViewItems.Visible = True
        Else
            grdICTSTATO.Visible = True
            tabViewItems.Visible = False
            'SplitContainer1.Visible = False
            'frmPreAllocate.Visible = False

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")

            If MsgBox("Reset Status Summary Positions to Detail Qtys?", MsgBoxStyle.YesNo, "Option to Fix Summary Status Table ICTSTAT2") = MsgBoxResult.Yes Then
                ASCMAIN1.sql = "" _
                    & "Begin" & vbCrLf _
                    & " Declare" & vbCrLf _
                    & "  Cursor C1 is Select * from " & ICTSTATO & ";" & vbCrLf _
                    & " Begin" & vbCrLf _
                    & "  For R1 in C1 Loop" & vbCrLf _
                    & "   Update ICTSTAT2 Set " & vbCrLf _
                    & "    WHSE_QTY_ONPO = R1.PO_DTL" & vbCrLf _
                    & "  , WHSE_QTY_OPEN = R1.SO_DTL" & vbCrLf _
                    & "  , WHSE_QTY_PICK = R1.SP_DTL" & vbCrLf _
                    & "  , WHSE_QTY_COMM = R1.PC_DTL" & vbCrLf _
                    & "    where WHSE_CODE = R1.WHSE_CODE and ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
                    & "   If SQL%NOTFOUND Then" & vbCrLf _
                    & "    Insert into ICTSTAT2" & vbCrLf _
                    & "     (WHSE_CODE, ITEM_CODE, WHSE_QTY_ONPO, WHSE_QTY_PLAN, WHSE_QTY_OPEN, WHSE_QTY_PICK, WHSE_QTY_COMM)" & vbCrLf _
                    & "    Values (R1.WHSE_CODE, R1.ITEM_CODE, R1.PO_DTL, R1.PP_DTL, R1.SO_DTL, R1.SP_DTL, R1.PC_DTL);" & vbCrLf _
                    & "   End If;" & vbCrLf _
                    & "  End Loop;" & vbCrLf _
                    & " End;" & vbCrLf _
                    & "End;"
                ASCDATA1.ExecuteSQL()
                '                    & "  , WHSE_QTY_PLAN = R1.PP_DTL" & vbCrLf _

                MsgBox("Status Summary Qtys have been Reset to match Detail Qty", MsgBoxStyle.OkOnly, "Success")
            End If
        End If
    End Sub

    Private Sub grdICTSTATO_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTSTATO.DoubleClickRow
        Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Value & ""
        Absx1.txtFor("ITEM_CODE").Text = ITEM_CODE
        Click_Command("View")
    End Sub

    Private Sub grdICTSTATO_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTSTATO.InitializeRow
        For Each STAT As String In New String() {"PO", "PP", "SO", "SP", "PC"}
            If Val(e.Row.Cells(STAT & "_DTL").Value & "") <> Val(e.Row.Cells(STAT & "_SUM").Value & "") Then
                e.Row.Cells(STAT & "_SUM").Appearance.BackColor = Color.Yellow
            End If
        Next
    End Sub

    Private Sub grdICTTRANX_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdICTTRANX.InitializeLayout

    End Sub

    Private Sub btnATSRefresh_Click(sender As Object, e As EventArgs) Handles btnATSRefresh.Click
        Refresh_ICTSTATS()

    End Sub
    Sub Refresh_ICTSTATS()

        Fill_Records("ICTSTATS")
    End Sub

    Private Sub grdICTSTATS_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTSTATS.DoubleClickRow
        If grdICTSTATS.ActiveRow IsNot Nothing AndAlso grdICTSTATS.ActiveRow.IsDataRow Then
            Absx1.txtFor("ITEM_CODE").Text = grdICTSTATS.ActiveRow.Cells("ITEM_CODE").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub optWHSE_CODE_ValueChanged(sender As Object, e As EventArgs) Handles optWHSE_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_cbeWHSE_CODE()
        Setup_Recent()
    End Sub

    Sub Setup_cbeWHSE_CODE()

        cbeWHSE_CODE.Visible = (optViewItems.Value = "Q" Or optViewItems.Value = "N")
        optWHSE_CODE.Visible = cbeWHSE_CODE.Visible

        cbeWHSE_CODE.Enabled = (optWHSE_CODE.Value = "I")

        Dim WHSEs As String = ""

        If optWHSE_CODE.Value = "A" Then
            WHSEs = "All Warehouses Combined"
        Else
            WHSEs = "Warehouse " & cbeWHSE_CODE.Value
        End If

        grdICTITEM1_Recent.Text = "Items w/Status Qtys - " & WHSEs

    End Sub

    Private Sub cbeWHSE_CODE_ValueChanged(sender As Object, e As EventArgs) Handles cbeWHSE_CODE.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Setup_Recent()
    End Sub


    Private Sub grdICTPINV1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdICTPINV1.InitializeRow
        Dim CONTAINER_NO As String = e.Row.Cells("CONTAINER_NO").Value & ""
        Dim ETA As String = "ETA_DATE_DC"
        If CONTAINER_NO.StartsWith("Inv Date +LT") Then
            e.Row.Cells(ETA).Appearance.ForeColor = System.Drawing.Color.Red
        ElseIf CONTAINER_NO = "Qty Open" Then
            e.Row.Appearance.BackColor = System.Drawing.Color.LightGreen
        Else
            e.Row.Appearance.BackColor = System.Drawing.Color.Empty
        End If

        Dim OPO_QTY As Int32 = Val(e.Row.Cells("OPO_QTY").Value & "")
        If OPO_QTY < 0 Then
            e.Row.Cells(ETA).Appearance.BackColor = System.Drawing.Color.Red
            e.Row.Cells(ETA).Appearance.ForeColor = System.Drawing.Color.White
        End If
    End Sub

    Private Sub btnProceed_Click(sender As Object, e As EventArgs) Handles btnProceed.Click

        Dim PROD_CODEs As New List(Of String)
        For Each row As UltraGridRow In grdICTPROD1.Rows
            If row.Cells("SEL").Value = 1 Then
                PROD_CODEs.Add($"'{row.Cells("PROD_CODE").Value}'")
            End If
        Next

        Dim WHSE_CODEs As New List(Of String)
        For Each row As UltraGridRow In grdICTWHSE1.Rows
            If row.Cells("SEL").Value = 1 Then
                WHSE_CODEs.Add($"'{row.Cells("WHSE_CODE").Value}'")
            End If
        Next

        Dim sqlProdFilter As String = If(PROD_CODEs.Count > 0, " AND ICTITEM1.PROD_CODE IN (" & String.Join(",", PROD_CODEs) & ")", "")
        Dim sqlWhseFilter As String = If(WHSE_CODEs.Count > 0, " where WHSE_CODE IN (" & String.Join(",", WHSE_CODEs) & ")", "")

        Dim EXCLUDED_ITEMS As String = txtExcluded.Text.Trim()
        Dim sqlExcl As String = ""

        If Not String.IsNullOrEmpty(EXCLUDED_ITEMS) Then
            Dim ITEM_LIST As List(Of String) = EXCLUDED_ITEMS.Split({",", vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries).
                                      Select(Function(x) x.Trim()).Distinct().ToList()
            Dim sqlITEM_LIST As String = "'" & String.Join("','", ITEM_LIST) & "'"

            Dim sqlValidation As String = "SELECT ITEM_CODE FROM ICTITEM1 WHERE ITEM_CODE IN (" & sqlITEM_LIST & ")"
            Dim VALID_ITEMS As DataTable = ASCDATA1.GetDataTable(sqlValidation)

            Dim INVALID_ITEMS As New List(Of String)
            For Each item As String In ITEM_LIST
                If Not VALID_ITEMS.AsEnumerable().Any(Function(row) row.Field(Of String)("ITEM_CODE") = item) Then
                    INVALID_ITEMS.Add(item)
                End If
            Next

            If INVALID_ITEMS.Count > 0 Then
                MessageBox.Show("The following item codes are invalid: " & vbCrLf & String.Join(", ", INVALID_ITEMS),
                            "Invalid Items", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            sqlExcl = " AND ICTITEM1.ITEM_CODE NOT IN (" & sqlITEM_LIST & ")"
        End If

        Dim sqlx As String = "Select ITEM_CODE" _
                                 & ", SUM (WHSE_QTY_ON_HAND) QTY_ONHD" & vbCrLf _
                                 & ", SUM (WHSE_QTY_ONPO) QTY_ONPO" & vbCrLf _
                                 & ", SUM (WHSE_QTY_PLAN) QTY_PLAN" & vbCrLf _
                                 & ", SUM (WHSE_QTY_OPEN) QTY_OPEN" & vbCrLf _
                                 & ", SUM (WHSE_QTY_PICK) QTY_PICK" & vbCrLf _
                                 & ", SUM (WHSE_QTY_COMM) QTY_COMM" & vbCrLf _
                                 & " from ICTSTAT2" & vbCrLf _
                                 & sqlWhseFilter _
                                 & " group by ITEM_CODE"

        Dim sqlx1 As String = "Select ITEM_CODE" _
                                 & ", SUM (WHSE_QTY_SHP) MTD_SHP" & vbCrLf _
                                 & " from ICTSTAT1" & vbCrLf _
                                 & $" where OPS_YYYYPP = '{ASCMAIN1.CYP}'" & vbCrLf _
                                 & Replace(sqlWhseFilter, " where ", " and ") _
                                 & " group by ITEM_CODE"

        Dim sqlF As String = "Select DPTITMF1.ITEM_CODE, SUM (DPTITMF1.FORECAST) ITEM_FC from DPTITMF1,SOTMKTC1" & vbCrLf _
                                 & "where SOTMKTC1.MARKET_CODE = DPTITMF1.MARKET_CODE" & vbCrLf _
                                 & $"   and DPTITMF1.OPS_YYYYPP = '{ASCMAIN1.CYP}'" & vbCrLf _
                                 & Replace(sqlWhseFilter, "where ", "and ") _
                                 & " group by DPTITMF1.ITEM_CODE"

        Dim sqlFC_MOS As String = ""
        For I As Integer = 1 To 24
            sqlFC_MOS &= $"+SIGN(ABS(NVL(DPTMRPG1_FC.QTY_{Format(I, "00")}, 0)))"
        Next
        sqlFC_MOS = Mid(sqlFC_MOS, 2)

        Dim SQLW As String = " And (NVL(QTY_ONHD,0) <> 0 Or NVL(QTY_ONPO,0) <> 0 Or NVL(QTY_PLAN,0) <> 0 Or NVL(QTY_OPEN,0) <> 0 Or NVL(QTY_PICK,0) <> 0 Or NVL(QTY_COMM,0) <> 0)"

        SQLW &= " AND (" & vbCrLf _
         & "  (DPTMRPG1_POS.QTY_00 < 1 AND DPTMRPG1_FC.QTY_01 > 0)" & vbCrLf _
         & "  OR (DPTMRPG1_POS.QTY_01 < 1 AND DPTMRPG1_FC.QTY_02 > 0)" & vbCrLf _
         & "  OR (DPTMRPG1_POS.QTY_02 < 1 AND DPTMRPG1_FC.QTY_03 > 0)" & vbCrLf _
         & ")"

        ASCMAIN1.sql = "Select ICTITEM1.SALES_DIVISION_CODE, ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_POS_MIN, ICTITEM1.PROD_CODE" & vbCrLf _
                & ", X.QTY_ONHD, X.QTY_ONPO, X.QTY_PLAN, X.QTY_OPEN, X.QTY_PICK, X.QTY_COMM, F.ITEM_FC, DPTMRPG1_POS.QTY_01 ITEM_POS" & vbCrLf _
                & ", X1.MTD_SHP" & vbCrLf _
                & ", DPTMRPG1_POS.QTY_00 ITEM_POS0, DPTMRPG1_POS.QTY_01 ITEM_POS1, DPTMRPG1_POS.QTY_02 ITEM_POS2, DPTMRPG1_POS.QTY_03 ITEM_POS3" & vbCrLf _
                & ", DPTMRPG1_FC.QTY_00 ITEM_FC0, DPTMRPG1_FC.QTY_01 ITEM_FC1, DPTMRPG1_FC.QTY_02 ITEM_FC2, DPTMRPG1_FC.QTY_03 ITEM_FC3" & vbCrLf _
                & ", " & sqlFC_MOS & " ITEM_FC_MOS" & vbCrLf _
                & ", CASE " & vbCrLf _
                & "     WHEN DPTMRPG1_POS.QTY_00 < 1 AND DPTMRPG1_FC.QTY_01 > 0 THEN 'x1' " & vbCrLf _
                & "     WHEN DPTMRPG1_POS.QTY_01 < 1 AND DPTMRPG1_FC.QTY_02 > 0 THEN 'x2' " & vbCrLf _
                & "     WHEN DPTMRPG1_POS.QTY_02 < 1 AND DPTMRPG1_FC.QTY_03 > 0 THEN 'x3' " & vbCrLf _
                & "   END AS FC_MONTH " & vbCrLf _
                & " from ICTITEM1,(" & sqlx & ") X,(" & sqlx1 & ") X1, (" & sqlF & ") F, DPTMRPG1 DPTMRPG1_POS, DPTMRPG1 DPTMRPG1_FC" & vbCrLf _
                & " where X.ITEM_CODE = ICTITEM1.ITEM_CODE" & vbCrLf _
                & "   and X1.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
                & "   And F.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
                & "   And DPTMRPG1_POS.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
                & "   And DPTMRPG1_POS.MRP_TYPE (+) = '6'" & vbCrLf _
                & "   And DPTMRPG1_FC.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
                & "   And DPTMRPG1_FC.MRP_TYPE (+) = '1'" & vbCrLf _
                & SQLW & sqlExcl & sqlProdFilter
        Dim DT As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
        If Not DT.Columns.Contains("QTY_NETA") Then
            DT.Columns.Add("QTY_NETA", GetType(System.Int64), "ISNULL(QTY_ONHD,0) - ISNULL(QTY_OPEN,0) - ISNULL(QTY_PICK,0) - ISNULL(QTY_COMM,0)")
        End If
        grdICTITEM1_optP.DataSource = DT
        Sort_grdColumns(grdICTITEM1_optP, "ITEM_CODE")
        If Not grdICTITEM1_optP.DisplayLayout.Bands(0).Summaries.Exists("ITEM_CODE") Then
            Create_Summary(grdICTITEM1_optP, "ITEM_CODE", "Count")
        End If

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim EXCLUDED_ITEMS As String = txtExcluded.Text.Trim()
        If String.IsNullOrEmpty(EXCLUDED_ITEMS) Then
            Exit Sub
        End If

        Dim ITEM_LIST As List(Of String) = EXCLUDED_ITEMS.Split({",", vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries).
                                  Select(Function(x) x.Trim()).Distinct().ToList() 'GETTING ALL DISTINCT ITEM CODES FROM THE TEXTBOX, SEPARATED BY A COMMA
        Dim sqlITEM_LIST As String = "'" & String.Join("','", ITEM_LIST) & "'"

        Dim sqlValidation As String = "SELECT ITEM_CODE FROM ICTITEM1 WHERE ITEM_CODE IN (" & sqlITEM_LIST & ")"
        Dim VALID_ITEMS As DataTable = ASCDATA1.GetDataTable(sqlValidation)

        Dim INVALID_ITEMS As New List(Of String)
        For Each item As String In ITEM_LIST
            If Not VALID_ITEMS.AsEnumerable().Any(Function(row) row.Field(Of String)("ITEM_CODE") = item) Then
                INVALID_ITEMS.Add(item)
            End If
        Next

        If INVALID_ITEMS.Count > 0 Then
            MessageBox.Show("The following item codes are invalid and cannot be saved: " & vbCrLf & String.Join(", ", INVALID_ITEMS),
                            "Invalid Items", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim result As DialogResult = MessageBox.Show("You are about to overwrite the existing list of excluded items. Would you like to proceed?",
                                                     "Confirm Overwrite", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If result = DialogResult.No Then Exit Sub

        Dim tbl As DataTable = dst.Tables("ICTEXCL1")
        For Each ITEM_CODE As String In ITEM_LIST
            If tbl.Select($"ITEM_CODE = '{ITEM_CODE}'").Length = 0 Then
                Dim newRow As DataRow = tbl.NewRow()
                newRow("ITEM_CODE") = ITEM_CODE
                tbl.Rows.Add(newRow)
            End If
        Next

        For Each row As DataRow In tbl.Select()
            Dim code As String = row("ITEM_CODE").ToString()
            If Not ITEM_LIST.Contains(code) Then
                row.Delete()
            End If
        Next

        For Each row As DataRow In tbl.Rows
            If row.RowState <> DataRowState.Unchanged Then
                Write_Audit_Trail(row, "E")
            End If
        Next

        ASCDATA1.ExecuteSQL("DELETE FROM ICTEXCL1")
        For Each ITEM_CODE As String In ITEM_LIST
            ASCDATA1.ExecuteSQL($"INSERT INTO ICTEXCL1 (ITEM_CODE) VALUES ('{ITEM_CODE}')")
        Next

        MessageBox.Show("Excluded items list has been successfully updated.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        'TAC.TACMAIN1.Record_Event("ICTEXCL1", ITEM_CODE, DATETIME_STAMP, ASCMAIN1.USER_ID, "ITEMEX", "Excluded Items List Changed", ITEM_CODE, "ICFSTAT1")
        'Write_Audit_Trail(rowICTITEM1, "E")
    End Sub

    Private Sub grdICTITEM1_optP_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdICTITEM1_optP.InitializeRow
        If e.Row.IsDataRow Then
            If e.Row.Cells("QTY_NETA").Value >= 0 Then
                e.Row.Cells("QTY_NETA").Appearance.ForeColor = Color.Green
            Else
                e.Row.Cells("QTY_NETA").Appearance.ForeColor = Color.Red
            End If

            For Each C As String In New String() {"ITEM_POS", "ITEM_POS0", "ITEM_POS1", "ITEM_POS2", "ITEM_POS3"}
                If Val(e.Row.Cells(C).Value & "") < Val(e.Row.Cells("ITEM_POS_MIN").Value & "") Then
                    e.Row.Cells(C).Appearance.ForeColor = Color.Red
                Else
                    e.Row.Cells(C).Appearance.ForeColor = Color.Empty
                End If
            Next

        End If
    End Sub
    Public Overrides Function Audit_Context() As Audit_Entity
        Dim E As New Audit_Entity



        E.TABLE_NAME = “ICTEXCL1”

        E.TABLE_DESC = “Excluded Items”



        Return E

    End Function
End Class