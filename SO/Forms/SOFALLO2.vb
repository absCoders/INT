Imports ABSolution

Public Class SOFALLO2

    Const maxAllocations As Integer = 99
    Dim iColumn As Integer
    Dim SINGLE_ITEM_ictr As Integer
    Dim ALLO_CTL_NOs As String
    Dim ALLO_CTL_NOs_to_Delete As New List(Of String)
    Dim ALLO_CTL_NOi() As String
    Dim COLLECTION_CODE As String
    Dim ALLO_CTL_NO_to_copy As String
    Dim CUST_CODE_to_copy As String
    Dim CUST_STORE_NO_to_copy As String
    Dim sql_ICTITEM1 As String
    Dim col_ICTITEM1 As New List(Of String)
    Dim rowARTCUST1 As DataRow
    Dim ALLO_CTL_NO_new As New List(Of String)
    Dim ITEM_CODE_new As New List(Of String)
    Dim iCol As New Dictionary(Of String, Int64)
    Dim sqlSOTORDR2 As String
    Dim sqlSOTORDRA As String
    Dim sqlSOTORDRA_groupBy As String
    Dim bln_Maintain_Sales As Boolean = False
    Dim rowSOTALLO1_Maintain_Sales As DataRow
    Dim HIDDEN_COLS As New List(Of Integer)
    Dim SOTALLOA As String = ""
    Dim SOTALLOB As String = ""
    Dim SOTALLOD As String = ""

    Dim SOTALLO1S As String
    Dim SOTALLO2S As String
    Dim SOTALLO3S As String
    Dim DT As String


    Dim COLUMN_NAMEs As New ArrayList
    Dim COLUMN_CAPTIONs As New ArrayList
    Dim COLUMN_NAME_by_Lvl() As String
    Dim COLUMN_CAPTION_by_Lvl() As String
    Dim G_by_Lvl() As Integer
    Dim SCOPE() As String
    Dim QCOLS As New Dictionary(Of String, String)
    Dim LVL As Int16

    Dim HC_CODE_lead_item As String = ""

    Dim SOTORDRX As String = ""

    Dim ARTCUSF2 As String = ""
    Dim sqlwhere_F As String = ""

    Dim sqlWHSE_CODEs As String = ""

    Dim lastStart As Date = Nothing
    Dim lastEnd As Date = Nothing
    Dim datesSet As Boolean = False

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("ICTPARM1")
        Get_PARM("SOTPARM1")

        AUDIT.Add("ICTITEM1", "E")

        If MENU_ITEM_OBJECT = "SOFALLOI" Then
            InquiryMode = True
        End If

        sqlWHSE_CODEs = ""
        ASCMAIN1.sql = "Select WHSE_CODE from ICTWHSE1 WHERE NVL(WHSE_MRP_EXC_IND,'0') <> '1' AND LP_CODE IS NOT NULL"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim WHSE_CODE As String = row.Item("WHSE_CODE")
            sqlWHSE_CODEs &= $",'{WHSE_CODE}'"
        Next
        sqlWHSE_CODEs = Mid(sqlWHSE_CODEs, 2)


        sql_ICTITEM1 = ", ICTITEM1.ITEM_DESC" & vbCrLf _
                & ", ICTITEM1.ITEM_SNU_CODE, ICTITEM1.ITEM_BASIC_PROMO" & vbCrLf _
                & ", ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.ITEM_STATUS" & vbCrLf _
                & ", ICTITEM1.PROD_CODE, ICTITEM1.ITEM_NOT_ALLOCATED, ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_SO_QTY_MULT, ICTITEM1.ITEM_SO_QTY_MIN" & vbCrLf _
                & ", ICTITEM1.ITEM_DATE_TO_SHIP, ICTITEM1.ITEM_CODE_COMPARE_TO LIKE_AS_ITEM" & vbCrLf
        For Each COLUMN_NAME In Split(Replace(Replace(sql_ICTITEM1, " ICTITEM1.", ""), vbCrLf, ""), ",")
            col_ICTITEM1.Add(COLUMN_NAME)
        Next

        Create_Allocation_Status_Tables(True)

        With dst

            SOTORDRX = ASCMAIN1.Temp_Table("Select ORDR_NO, ALLO_CTL_NO, ITEM_CODE, CUST_CODE, NVL(ORDR_QTY_OPEN,0)+NVL(ORDR_QTY_PICK,0)+NVL(ORDR_QTY_SHIP,0) QTY from SOTORDR2 where ROWNUM < 1")

            ASCMAIN1.sql = "Select ITEM_CODE, CUST_CODE, ALLO_CTL_NO" & vbCrLf _
                & ", SUM (DECODE(IO,'ADD',QTY,0)) ADDS" & vbCrLf _
                & ", SUM (DECODE(IO,'DED',QTY,0)) DEDS" & vbCrLf _
                & " from (" & vbCrLf _
                & "Select 'DED' IO, X.* from (" & vbCrLf _
                & "Select ORDR_NO, ITEM_CODE, CUST_CODE, ALLO_CTL_NO, QTY from " & SOTORDRX & " SOTORDRX" & vbCrLf _
                & " minus " & vbCrLf _
                & "Select ORDR_NO, ITEM_CODE, CUST_CODE, ALLO_CTL_NO" & vbCrLf _
                & ", NVL(ORDR_QTY_OPEN,0)+NVL(ORDR_QTY_PICK,0)+NVL(ORDR_QTY_SHIP,0) QTY from SOTORDR2" & vbCrLf _
                & " where ALLO_CTL_NO = :PARM1) X" & vbCrLf _
                & " union " & vbCrLf _
                & "Select 'ADD' IO, Y.* FROM (" & vbCrLf _
                & "Select ORDR_NO, ITEM_CODE, CUST_CODE, ALLO_CTL_NO" & vbCrLf _
                & ", NVL(ORDR_QTY_OPEN,0)+NVL(ORDR_QTY_PICK,0)+NVL(ORDR_QTY_SHIP,0) QTY from SOTORDR2" & vbCrLf _
                & " where ALLO_CTL_NO = :PARM1" & vbCrLf _
                & " minus " & vbCrLf _
                & "Select ORDR_NO, ITEM_CODE, CUST_CODE, ALLO_CTL_NO, QTY from " & SOTORDRX & " SOTORDRX) Y" & vbCrLf _
                & ") group by ITEM_CODE, CUST_CODE, ALLO_CTL_NO"

            MyBase.Create_TDA(.Tables.Add, "SOTORDRD", "**", 0, False, "V", 0)
            With .Tables("SOTORDRD")
                .Columns("ADDS").DataType = GetType(Int64)
                .Columns("DEDS").DataType = GetType(Int64)
            End With

            ASCMAIN1.sql = "Select ORDR_NO, ITEM_CODE, CUST_CODE, ALLO_CTL_NO" & vbCrLf _
                & ", SUM (DECODE(IO,'ADD',QTY,0)) ADDS" & vbCrLf _
                & ", SUM (DECODE(IO,'DED',QTY,0)) DEDS" & vbCrLf _
                & " from (" & vbCrLf _
                & "Select 'DED' IO, X.* from (" & vbCrLf _
                & "Select ORDR_NO, ITEM_CODE, CUST_CODE, ALLO_CTL_NO, QTY from " & SOTORDRX & " SOTORDRX" & vbCrLf _
                & " minus " & vbCrLf _
                & "Select ORDR_NO, ITEM_CODE, CUST_CODE, ALLO_CTL_NO" & vbCrLf _
                & ", NVL(ORDR_QTY_OPEN,0)+NVL(ORDR_QTY_PICK,0)+NVL(ORDR_QTY_SHIP,0) QTY from SOTORDR2" & vbCrLf _
                & " where ALLO_CTL_NO = :PARM1) X" & vbCrLf _
                & " union " & vbCrLf _
                & "Select 'ADD' IO, Y.* FROM (" & vbCrLf _
                & "Select ORDR_NO, ITEM_CODE, CUST_CODE, ALLO_CTL_NO" & vbCrLf _
                & ", NVL(ORDR_QTY_OPEN,0)+NVL(ORDR_QTY_PICK,0)+NVL(ORDR_QTY_SHIP,0) QTY from SOTORDR2" & vbCrLf _
                & " where ALLO_CTL_NO = :PARM1" & vbCrLf _
                & " minus " & vbCrLf _
                & "Select ORDR_NO, ITEM_CODE, CUST_CODE, ALLO_CTL_NO, QTY from " & SOTORDRX & " SOTORDRX) Y" & vbCrLf _
                & ") group by ORDR_NO, ITEM_CODE, CUST_CODE, ALLO_CTL_NO"
            MyBase.Create_TDA(dst.Tables.Add, "SOTORDRN", "**", 0, False, "V", 0)
            With .Tables("SOTORDRN")
                .Columns("ADDS").DataType = GetType(Int64)
                .Columns("DEDS").DataType = GetType(Int64)
            End With
            Create_Relation("SOTORDRD", "SOTORDRN", "ALLO_CTL_NO,ITEM_CODE,CUST_CODE")

            ASCMAIN1.sql = $"SELECT * FROM SOTALLH1 WHERE ALLO_CTL_NO = :PARM1"
            MyBase.Create_TDA(dst.Tables.Add, "SOTALLH1", "**", 0, False, "V", 0)

            ASCMAIN1.sql = $"SELECT * FROM SOTALLH2 WHERE ALLO_CTL_NO = :PARM1"
            MyBase.Create_TDA(dst.Tables.Add, "SOTALLH2", "**", 0, False, "V", 0)

            ASCMAIN1.sql = $"SELECT * FROM SOTALLH3 WHERE ALLO_CTL_NO = :PARM1"
            MyBase.Create_TDA(dst.Tables.Add, "SOTALLH3", "**", 0, False, "V", 0)

            'SOTALLH1_SOTALLH2 => CONNECT USING XNO AND ALLO CTL NO
            Dim relation1 As New DataRelation("SOTALLH1_SOTALLH2", New DataColumn() {dst.Tables("SOTALLH1").Columns("XNO"), dst.Tables("SOTALLH1").Columns("ALLO_CTL_NO")}, New DataColumn() {dst.Tables("SOTALLH2").Columns("XNO"), dst.Tables("SOTALLH2").Columns("ALLO_CTL_NO")})
            dst.Relations.Add(relation1)

            'SOTALLH2_SOTALLH3 => CONNECT USING XNO,ALLO CTL NO, AND CUST CODE
            Dim relation2 As New DataRelation("SOTALLH2_SOTALLH3", New DataColumn() {dst.Tables("SOTALLH2").Columns("XNO"), dst.Tables("SOTALLH2").Columns("ALLO_CTL_NO"), dst.Tables("SOTALLH2").Columns("CUST_CODE")}, New DataColumn() {dst.Tables("SOTALLH3").Columns("XNO"), dst.Tables("SOTALLH3").Columns("ALLO_CTL_NO"), dst.Tables("SOTALLH3").Columns("CUST_CODE")})
            dst.Relations.Add(relation2)

            ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE" & sql_ICTITEM1 _
                & ", SOTALLO1.ALLO_CTL_NO, SOTALLO1.DATE_START, SOTALLO1.DATE_END" & vbCrLf _
                & ", SOTALLO1.INIT_OPER, SOTALLO1.INIT_DATE, SOTALLO1.LAST_OPER, SOTALLO1.LAST_DATE" & vbCrLf _
                & " from SOTALLO1, ICTITEM1" & vbCrLf _
                & " where SOTALLO1.ITEM_CODE = ICTITEM1.ITEM_CODE"
            'MyBase.Create_TDA(.Tables.Add, "ICTITEM1", "**", 0, False, String.Empty, 1)
            MyBase.Create_TDA(.Tables.Add, "ICTITEM1", "**", 1, True, String.Empty, 1, "ITEM_NOT_ALLOCATED")
            With .Tables("ICTITEM1")
                .Columns.Add("SELECTED")
                .Columns("SELECTED").DefaultValue = "0"
                .Columns("ALLO_CTL_NO").AllowDBNull = True
                .Columns.Add("ITEM_IMAGE", GetType(System.Byte()))
            End With

            ASCMAIN1.sql = "Select SOTALLO1.*" & sql_ICTITEM1 _
                & ", X.WHSE_QTY_ON_HAND" & vbCrLf _
                & " from SOTALLO1, ICTITEM1" & vbCrLf _
                & ", (Select ITEM_CODE, Sum (WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND from ICTSTAT2" & vbCrLf _
                & $" where WHSE_CODE IN ({sqlWHSE_CODEs}) group by ITEM_CODE) X" & vbCrLf _
                & " where SOTALLO1.ITEM_CODE = ICTITEM1.ITEM_CODE" & vbCrLf _
                & "   And X.ITEM_CODE (+) = SOTALLO1.ITEM_CODE"
            MyBase.Create_TDA(.Tables.Add, "SOTALLOX", "**", 0, False, String.Empty, 1)
            With .Tables("SOTALLOX")
                .Columns.Add("SELECTED")
                .Columns("SELECTED").DefaultValue = "0"
                .Columns("ALLO_CTL_NO").AllowDBNull = True
                '   .Columns.Add("WHSE_QTY_ON_HAND", GetType(System.Int32))
            End With

            ASCMAIN1.sql = "Select SOTALLO1.*" & sql_ICTITEM1 _
                & ", X.WHSE_QTY_ON_HAND" & vbCrLf _
                & " from SOTALLO1,ICTITEM1" & vbCrLf _
                & ", (Select ITEM_CODE, Sum (WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND from ICTSTAT2" & vbCrLf _
                & $" where WHSE_CODE IN ({sqlWHSE_CODEs}) group by ITEM_CODE) X" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = SOTALLO1.ITEM_CODE" & vbCrLf _
                & "   and X.ITEM_CODE (+) = SOTALLO1.ITEM_CODE"
            Create_TDA(.Tables.Add, "SOTALLO1", "**", 0)
            With .Tables("SOTALLO1")
                .Columns.Add("NEW")
                .Columns("NEW").DefaultValue = "0"
                .Columns.Add("QTY_ALLO_TOTAL", GetType(System.Int64))
                .Columns.Add("QTY_ALLO_BALANCE", GetType(System.Int64), "IIF(ISNULL(QTY_ALLO_PLAN,0)=0,NULL,ISNULL(QTY_ALLO_PLAN,0) - ISNULL(QTY_ALLO_TOTAL,0))")
            End With

            ASCMAIN1.sql = "Select SOTALLO2.*, ARTCUST1.CUST_NAME, ARTCUST1.SREP_CODE, ARTCUST1.TRADE_CLASS_CODE, ARTCUST1.CUST_ALLOCATE_BY_STORE" & vbCrLf _
                & " from SOTALLO2, ARTCUST1" & vbCrLf _
                & " where ARTCUST1.CUST_CODE = SOTALLO2.CUST_CODE"
            Create_TDA(.Tables.Add, "SOTALLO2", "**", 1)


            ASCMAIN1.sql = Construct_sqlSOTALLO3()
            Create_TDA(.Tables.Add, "SOTALLO3", "**", 1)

            ASCMAIN1.sql = Construct_sqlSOTALLO4()
            Create_TDA(.Tables.Add, "SOTALLO4", "**", 1)

            Create_TDA(dst.Tables.Add, "SOTNGMSG", "*", 0)
            Fill_Records("SOTNGMSG")

            ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, SREP_CODE, TRADE_CLASS_CODE, CUST_ALLOCATE_BY_STORE from ARTCUST1"
            MyBase.Create_TDA(.Tables.Add, "SOTALLOC", "**", 0, False, String.Empty, 1)
            With .Tables("SOTALLOC").Columns
                .Add("STORE_COUNT_HC", GetType(System.Int64))
            End With

            ASCMAIN1.sql = "Select CUST_CODE, CUST_STORE_NO, NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_NAME) CUST_STORE_LOCATION, SELL_CODE, DMA_CODE, CUST_STORE_STATUS from ARTCUST2"
            MyBase.Create_TDA(.Tables.Add, "SOTALLOS", "**", 0, False, String.Empty, 2)

            ASCMAIN1.sql = "Select CUST_CODE, CUST_STORE_NO, NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_NAME) CUST_STORE_LOCATION, SELL_CODE from ARTCUST2"
            MyBase.Create_TDA(.Tables.Add, "SOTALLOT", "**", 0, False, String.Empty, 2)
            With .Tables("SOTALLOT")
                For iCtr As Integer = 1 To 10
                    Dim colName As String = "EVENT_" & Format(iCtr, "00")
                    If Not .Columns.Contains(colName) Then
                        .Columns.Add(colName, GetType(Int64))
                    End If
                Next
            End With

            For Each TABLE_NAME As String In New String() {"SOTALLOC", "SOTALLOS"}
                With .Tables(TABLE_NAME)
                    .Columns.Add("USED", GetType(System.Int64))
                    .Columns.Add("BALANCE", GetType(System.Int64)) ', "ISNULL(ALLO_01,0)-ISNULL(USED,0)")
                End With

                For iCtr As Integer = 1 To maxAllocations
                    .Tables(TABLE_NAME).Columns.Add("ALLO_" & Format(iCtr, "00"), GetType(System.Int64))
                    .Tables(TABLE_NAME).Columns.Add("ALLO_NOTES_" & Format(iCtr, "00"), GetType(System.String))
                    .Tables(TABLE_NAME).Columns("ALLO_NOTES_" & Format(iCtr, "00")).MaxLength = 50
                Next
                With .Tables(TABLE_NAME).Columns
                    .Add("ORDR_QTY", GetType(System.Int64))
                    .Add("ORDR_QTY_OPEN", GetType(System.Int64))
                    .Add("ORDR_QTY_PICK", GetType(System.Int64))
                    .Add("ORDR_QTY_SHIP", GetType(System.Int64))
                    .Add("ORDR_QTY_CANC", GetType(System.Int64))
                    .Add("QTY_LEFT", GetType(System.Int64), "ISNULL(ALLO_01,0)-ISNULL(ORDR_QTY_SHIP,0)-ISNULL(ORDR_QTY_PICK,0)")
                    .Add("QTY_BAL", GetType(System.Int64), "IIF(QTY_LEFT>=0,QTY_LEFT,0)")
                    .Add("QTY_OVER", GetType(System.Int64), "IIF(QTY_LEFT-ISNULL(ORDR_QTY_OPEN,0)>=0,0,ISNULL(ORDR_QTY_OPEN,0)-QTY_LEFT)")

                    .Add("SI_PLAN", GetType(System.Int64))
                    .Add("ST_PLAN", GetType(System.Int64))
                    .Add("SI_HIST", GetType(System.Int64))
                    .Add("ST_HIST", GetType(System.Int64))

                    .Add("LY_QTY_SELL_IN", GetType(System.Int64))
                    .Add("LY_QTY_SELL_THRU", GetType(System.Int64))
                    .Add("LY_QTY_SELL_IN_THRU_PCT", GetType(System.Decimal), "IIF(ISNULL(LY_QTY_SELL_IN,0)=0,0,100 * ISNULL(LY_QTY_SELL_THRU,0)/ISNULL(LY_QTY_SELL_IN,0))")
                    .Add("TY_VS_LY_PCT", GetType(System.Decimal), "IIF(ISNULL(LY_QTY_SELL_IN,0)=0,0,100 * ISNULL(ORDR_QTY_SHIP,0)/ISNULL(LY_QTY_SELL_IN,0))")

                    .Add("STORES", GetType(System.Int64))
                    .Add("RETAIL_SALES", GetType(System.Decimal))
                    .Add("RETAIL_SALES_PCT", GetType(System.Decimal))
                    .Add("ALLO_SPREAD", GetType(System.Int64))
                End With
            Next

            '            & ", MIN (SOTORDR1.CUST_DC_NO) CUST_DC_NO1" & vbCrLf _
            '& ", MAX (SOTORDR1.CUST_DC_NO) CUST_DC_NO2" & vbCrLf _


            ASCMAIN1.sql = "Select SOTORDR2.ITEM_CODE, SOTORDR2.ALLO_CTL_NO, SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
                & ", COUNT (*) LINES" & vbCrLf _
                & ", MIN (SOTORDR1.CUST_STORE_NO) CUST_STORE_NO1" & vbCrLf _
                & ", MAX (SOTORDR1.CUST_STORE_NO) CUST_STORE_NO2" & vbCrLf _
                & ", SOTORDR1.CUST_DC_NO, SOTORDR1.ORDR_STATUS" & vbCrLf _
                & ", SOTORDR1.CUST_CODE, NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE) CUST_CODE_ALLO" & vbCrLf _
                & ", SOTORDR1.ORDR_YYYYPP_UPDATED" & vbCrLf _
                & ", MIN (SOTORDR1.ORDR_DATE_SHIPPED) ORDR_DATE_SHIPPED" & vbCrLf _
                & ", SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE, SOTORDR1.ORDR_ALLO_DATE" & vbCrLf _
                & ", SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_DATE_REL" & vbCrLf _
                & " from SOTORDR1,SOTORDR2,ARTCUST1" & vbCrLf _
                & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE" & vbCrLf _
                & "   And SOTORDR1.ORDR_STATUS in ('O','F','P')" & vbCrLf
            sqlSOTORDRA = ASCMAIN1.sql
            ASCMAIN1.sql &= "" _
                & "  and SOTORDR2.ITEM_CODE = :PARM1 and (SOTORDR2.ALLO_CTL_NO = :PARM2 or SOTORDR1.ORDR_SHIP_DATE >= :PARM3 and SOTORDR1.ORDR_SHIP_DATE <= :PARM4)" & vbCrLf
            sqlSOTORDRA_groupBy = " group by SOTORDR2.ITEM_CODE, SOTORDR2.ALLO_CTL_NO, SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
                & ", SOTORDR1.CUST_DC_NO, SOTORDR1.ORDR_STATUS, SOTORDR1.CUST_CODE, NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE), SOTORDR1.ORDR_YYYYPP_UPDATED" & vbCrLf _
                & ", SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE, SOTORDR1.ORDR_ALLO_DATE" & vbCrLf _
                & ", SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_DATE_REL"
            ASCMAIN1.sql &= sqlSOTORDRA_groupBy

            Create_TDA(.Tables.Add, "SOTORDRA", "**", 0, True, "VVDD", 0)
            With .Tables("SOTORDRA")
                .Columns.Add("SELECTED")
                .Columns("SELECTED").DefaultValue = "0"
                .Columns.Add("USED", GetType(System.Int64), "IIF(ISNULL(SELECTED,'0')='1',ISNULL(ORDR_QTY_OPEN,0)+ISNULL(ORDR_QTY_PICK,0)+ISNULL(ORDR_QTY_SHIP,0),NULL)")
                .Columns("LINES").DataType = GetType(System.Int64)
                .Columns("ORDR_QTY").DataType = GetType(System.Int64)
                .Columns("ORDR_QTY_OPEN").DataType = GetType(System.Int64)
                .Columns("ORDR_QTY_PICK").DataType = GetType(System.Int64)
                .Columns("ORDR_QTY_SHIP").DataType = GetType(System.Int64)
                .Columns("ORDR_QTY_CANC").DataType = GetType(System.Int64)
            End With

            ASCMAIN1.sql = "Select SOTORDR2.ITEM_CODE, SOTORDR2.ALLO_CTL_NO, SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO" & vbCrLf _
                & ", NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE) CUST_CODE_ALLO" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
                & ", MIN (SOTORDR1.ORDR_DATE_SHIPPED) ORDR_DATE_SHIPPED" & vbCrLf _
                & " from SOTORDR1,SOTORDR2,ARTCUST1" & vbCrLf _
                & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   And ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE" & vbCrLf _
                & "   and SOTORDR1.ORDR_STATUS in ('O','F','P')" & vbCrLf
            ASCMAIN1.sql &= "" _
                & "  and SOTORDR2.ITEM_CODE = :PARM1 and (SOTORDR2.ALLO_CTL_NO = :PARM2 or SOTORDR1.ORDR_SHIP_DATE >= :PARM3 and SOTORDR1.ORDR_SHIP_DATE <= :PARM4)" & vbCrLf
            ASCMAIN1.sql &= " group by SOTORDR2.ITEM_CODE, SOTORDR2.ALLO_CTL_NO, SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE)"

            Create_TDA(.Tables.Add, "SOTORDRQ", "**", 0, False, "VVDD", 0)


            ASCMAIN1.sql = "Select SOTORDR2.*" & vbCrLf _
                & ", SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE, SOTORDR1.ORDR_ALLO_DATE" & vbCrLf _
                & ", SOTORDR1.CUST_DC_NO, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                & ", SOTORDR1.ORDR_DATE_REL, NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE) CUST_CODE_ALLO" & vbCrLf _
                & " from SOTORDR1,SOTORDR2,ARTCUST1" & vbCrLf _
                & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   And ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE" & vbCrLf _
                & "   And SOTORDR1.ORDR_STATUS in ('O','F','P')" & vbCrLf
            sqlSOTORDR2 = ASCMAIN1.sql
            ASCMAIN1.sql &= "" _
                & "  and SOTORDR2.ITEM_CODE = :PARM1 and (SOTORDR2.ALLO_CTL_NO = :PARM2 or SOTORDR1.ORDR_SHIP_DATE >= :PARM3 and SOTORDR1.ORDR_SHIP_DATE <= :PARM4)"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, True, "VVDD", 2)
            With .Tables("SOTORDR2")
                .Columns.Add("SELECTED")
                .Columns("SELECTED").DefaultValue = "0"
                .Columns.Add("USED", GetType(System.Int64), "IIF(ISNULL(SELECTED,'0')='1',ISNULL(ORDR_QTY_OPEN,0)+ISNULL(ORDR_QTY_PICK,0)+ISNULL(ORDR_QTY_SHIP,0),NULL)")
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
                & "   where SOTORDR2.ALLO_CTL_NO = :PARM1" & vbCrLf _
                & "     and ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE" & vbCrLf _
                & "     and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                & " group by SOTORDR2.ALLO_CTL_NO, NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE)"
            Create_TDA(.Tables.Add, "SOTALLOZ", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select ICTSTAT2.ITEM_CODE" & vbCrLf _
                 & ", SUM (ICTSTAT2.WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND" & vbCrLf _
                 & ", SUM (ICTSTAT2.WHSE_QTY_ONPO) WHSE_QTY_ONPO" & vbCrLf _
                 & ", SUM (ICTSTAT2.WHSE_QTY_PLAN) WHSE_QTY_PLAN" & vbCrLf _
                 & ", SUM (ICTSTAT2.WHSE_QTY_OPEN) WHSE_QTY_OPEN" & vbCrLf _
                 & ", SUM (ICTSTAT2.WHSE_QTY_PICK) WHSE_QTY_PICK" & vbCrLf _
                 & ", SUM (ICTSTAT2.WHSE_QTY_COMM) WHSE_QTY_COMM" & vbCrLf _
                 & "  from ICTSTAT2" & vbCrLf _
                 & " where ICTSTAT2.ITEM_CODE = :PARM1" & vbCrLf _
                 & " group by ICTSTAT2.ITEM_CODE"
            Create_TDA(.Tables.Add, "ICTSTAT2", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, CUST_CODE_ALLO, CUST_ALLOCATE_BY_STORE from ARTCUST1" _
                & " where CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select ALLO_CTL_NO, CUST_CODE" & vbCrLf _
                & ", ORDR_QTY" & vbCrLf _
                & ", ORDR_QTY_OPEN, ORDR_QTY_PICK, ORDR_QTY_SHIP, ORDR_QTY_CANC" & vbCrLf _
                & ", ORDR_QTY_SHIP LY_QTY_SELL_IN1" & vbCrLf _
                & ", ORDR_QTY_SHIP LY_QTY_SELL_IN2" & vbCrLf _
                & ", ORDR_QTY_SHIP LY_QTY_SELL_THRU1" & vbCrLf _
                & ", ORDR_QTY_SHIP LY_QTY_SELL_THRU2" & vbCrLf _
                & " from SOTORDR2"
            SOTALLOB = ASCMAIN1.Temp_Table(ASCMAIN1.sql & " where ROWNUM < 1")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTALLOB & " Add Primary Key (ALLO_CTL_NO,CUST_CODE)")
            Create_TDA(.Tables.Add("SOTALLOB"), SOTALLOB, "*")
            With .Tables("SOTALLOB")
                .Columns.Add("LY_QTY_SELL_IN", GetType(System.Int64), "IIF(ISNULL(LY_QTY_SELL_IN1,0)<>0,LY_QTY_SELL_IN1,LY_QTY_SELL_IN2)")
                .Columns.Add("LY_QTY_SELL_THRU", GetType(System.Int64), "IIF(ISNULL(LY_QTY_SELL_THRU1,0)<>0,LY_QTY_SELL_THRU1,LY_QTY_SELL_THRU2)")
            End With

            ASCMAIN1.sql = "Select ALLO_CTL_NO, CUST_CODE, CUST_STORE_NO" & vbCrLf _
                & ", ORDR_QTY" & vbCrLf _
                & ", ORDR_QTY_OPEN, ORDR_QTY_PICK, ORDR_QTY_SHIP, ORDR_QTY_CANC" & vbCrLf _
                & ", ORDR_QTY_SHIP LY_QTY_SELL_IN1" & vbCrLf _
                & ", ORDR_QTY_SHIP LY_QTY_SELL_IN2" & vbCrLf _
                & ", ORDR_QTY_SHIP LY_QTY_SELL_THRU1" & vbCrLf _
                & ", ORDR_QTY_SHIP LY_QTY_SELL_THRU2" & vbCrLf _
                & " from SOTORDR2"
            SOTALLOD = ASCMAIN1.Temp_Table(ASCMAIN1.sql & " where ROWNUM < 1")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTALLOD & " Add Primary Key (ALLO_CTL_NO,CUST_CODE,CUST_STORE_NO)")
            Create_TDA(.Tables.Add("SOTALLOD"), SOTALLOD, "*")
            With .Tables("SOTALLOD")
                .Columns.Add("LY_QTY_SELL_IN", GetType(System.Int64), "IIF(ISNULL(LY_QTY_SELL_IN1,0)<>0,LY_QTY_SELL_IN1,LY_QTY_SELL_IN2)")
                .Columns.Add("LY_QTY_SELL_THRU", GetType(System.Int64), "IIF(ISNULL(LY_QTY_SELL_THRU1,0)<>0,LY_QTY_SELL_THRU1,LY_QTY_SELL_THRU2)")
            End With

            ASCMAIN1.sql = "Select ALLO_CTL_NO, ITEM_CODE" & vbCrLf _
                & ", ITEM_CODE ITEM_CODE_COMPARE_TO, ITEM_CODE ITEM_CODE_COMPARE_TO_ALT" & vbCrLf _
                & ", ORDR_YYYYPP_UPDATED TYP_START, ORDR_YYYYPP_UPDATED TYP_END" & vbCrLf _
                & ", ORDR_YYYYPP_UPDATED LYP_START, ORDR_YYYYPP_UPDATED LYP_END" & vbCrLf _
                & " from SOTORDR2"
            SOTALLOA = ASCMAIN1.Temp_Table(ASCMAIN1.sql & " where ROWNUM < 1")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTALLOA & " Add Primary Key (ALLO_CTL_NO)")
            Create_TDA(.Tables.Add("SOTALLOA"), SOTALLOA, "*")

            ASCMAIN1.sql = "Select * from " & SOTALLO1S
            Create_TDA(.Tables.Add, "SOTALLO1S", "**", 0, False)
            ASCMAIN1.sql = "Select * from " & SOTALLO2S
            Create_TDA(.Tables.Add, "SOTALLO2S", "**", 0, False)
            Create_Relation("SOTALLO1S", "SOTALLO2S", "ALLO_CTL_NO")
            With .Tables("SOTALLO1S").Columns
                .Add("QTY_ALLO", GetType(System.Int64), "SUM(CHILD.QTY_ALLO)")
                .Add("QTY_ALLO_RSRV", GetType(System.Int64), "IIF(ISNULL(QTY_ALLO_PLAN,0)=0,0,QTY_ALLO_PLAN - QTY_ALLO_TOTAL)")
            End With
            With .Tables("SOTALLO2S").Columns
                .Add("QTY_ALLO_BAL", GetType(System.Int64), "IIF(ISNULL(QTY_ALLO,0)<=ISNULL(ORDR_QTY_SHIP,0)+ISNULL(ORDR_QTY_PICK,0),0,ISNULL(QTY_ALLO,0) - (ISNULL(ORDR_QTY_SHIP,0)+ISNULL(ORDR_QTY_PICK,0)))")
                ' .Add("QTY_ALLO_USED", GetType(System.Decimal), "IIF(ISNULL(QTY_ALLO,0)=0,0,100*(1-(ISNULL(QTY_ALLO_BAL,0)/ISNULL(QTY_ALLO,0))))")
                .Add("QTY_ALLO_USED", GetType(System.Decimal), "IIF(ISNULL(QTY_ALLO,0)=0,0,100*((ISNULL(ORDR_QTY_SHIP,0)+ISNULL(ORDR_QTY_PICK,0))/ISNULL(QTY_ALLO,0)))")
            End With



            ASCMAIN1.sql = "Select SOTALLO1S.*" & vbCrLf _
                & " from " & SOTALLO1S & " SOTALLO1S"
            Create_TDA(.Tables.Add, "SOTALLOCS", "**", 0, False)
            With .Tables("SOTALLOCS").Columns
                .Add("QTY_ALLO", GetType(System.Int64))
                .Add("ORDR_QTY", GetType(System.Int64))
                .Add("ORDR_QTY_OPEN", GetType(System.Int64))
                .Add("ORDR_QTY_PICK", GetType(System.Int64))
                .Add("ORDR_QTY_SHIP", GetType(System.Int64))
                .Add("ORDR_QTY_CANC", GetType(System.Int64))
                .Add("LY_QTY_SELL_IN", GetType(System.Int64))
                .Add("LY_QTY_SELL_THRU", GetType(System.Int64))

                .Add("QTY_LEFT", GetType(System.Int64), "ISNULL(QTY_ALLO,0)-ISNULL(ORDR_QTY_SHIP,0)-ISNULL(ORDR_QTY_PICK,0)")
                .Add("QTY_BAL", GetType(System.Int64), "IIF(QTY_LEFT>=0,QTY_LEFT,0)")
                .Add("QTY_OVER", GetType(System.Int64), "IIF(QTY_LEFT-ISNULL(ORDR_QTY_OPEN,0)>=0,0,ISNULL(ORDR_QTY_OPEN,0)-QTY_LEFT)")

                .Add("SI_PLAN", GetType(System.Int64))
                .Add("ST_PLAN", GetType(System.Int64))
                .Add("SI_HIST", GetType(System.Int64))
                .Add("ST_HIST", GetType(System.Int64))

                .Add("LY_QTY_SELL_IN_THRU_PCT", GetType(System.Decimal), "IIF(ISNULL(LY_QTY_SELL_IN,0)=0,0,100 * ISNULL(LY_QTY_SELL_THRU,0)/ISNULL(LY_QTY_SELL_IN,0))")
                .Add("TY_VS_LY_PCT", GetType(System.Decimal), "IIF(ISNULL(LY_QTY_SELL_IN,0)=0,0,100 * ISNULL(ORDR_QTY_SHIP,0)/ISNULL(LY_QTY_SELL_IN,0))")
                .Add("QTY_ALLO_RSRV", GetType(System.Int64), "IIF(ISNULL(QTY_ALLO_PLAN,0)=0,0,QTY_ALLO_PLAN - QTY_ALLO_TOTAL)")

                .Add("STORES", GetType(System.Int64))
                .Add("RETAIL_SALES", GetType(System.Decimal))
                .Add("RETAIL_SALES_PCT", GetType(System.Decimal))
                .Add("ALLO_SPREAD", GetType(System.Int64))

            End With

            With .Tables.Add("SATANALC")
                .Columns.Add("COLUMN_NAME")
                .Columns.Add("COLUMN_NAME_CODE")
                .Columns.Add("COLUMN_NAME_DESC")
                .Columns.Add("TABLE_NAME_LOOKUP")
                .PrimaryKey = New DataColumn() { .Columns("COLUMN_NAME")}
            End With

            ASCMAIN1.sql = "Select CTL_NO_TEXT COLUMN_NAME, CTL_NO_TEXT CODE_VALUE, CTL_NO_TEXT DESC_VALUE from TATCTLN1"
            Create_TDA(.Tables.Add, "SATANALD", "**", 0, False, "", 2)


            .Tables("SATANALD").Columns("DESC_VALUE").MaxLength = -1 ' 100


            ASCMAIN1.sql = "Select X.*, X1.DATE_START DATE_START_F, X2.DATE_START DATE_START_R" & vbCrLf _
                & " from SOTALLO1 X1, SOTALLO1 X2, (" & vbCrLf _
                & "Select SOTORDR1.CUST_CODE, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                & ", SOTORDR2.ALLO_CTL_NO, SOTORDR2.ALLO_CTL_NO_REL, SOTORDR2.ITEM_CODE" & vbCrLf _
                & ", COUNT (*) RECS, SUM (SOTORDR2.ORDR_QTY_SHIP) QTY" & vbCrLf _
                & ", MIN (SOTORDR1.ORDR_NO) OMIN, MAX (SOTORDR1.ORDR_NO) OMAX from SOTORDR2,SOTORDR1" & vbCrLf _
                & "where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO and SOTORDR2.ORDR_YYYYPP_UPDATED = :PARM1" & vbCrLf _
                & "and SOTORDR2.ALLO_CTL_NO_REL is not null and SOTORDR2.ALLO_CTL_NO is not null " & vbCrLf _
                & "and SOTORDR2.ALLO_CTL_NO <> SOTORDR2.ALLO_CTL_NO_REL" & vbCrLf _
                & "group by SOTORDR1.CUST_CODE, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                & ", SOTORDR2.ALLO_CTL_NO, SOTORDR2.ALLO_CTL_NO_REL, SOTORDR2.ITEM_CODE) X" & vbCrLf _
                & "where X1.ALLO_CTL_NO = X.ALLO_CTL_NO" & vbCrLf _
                & "  and X2.ALLO_CTL_NO = X.ALLO_CTL_NO_REL"
            Create_TDA(.Tables.Add, "SOTORDRU", "**", 0, False, "V", 0)
        End With




        dst.Tables("SATANALC").Rows.Clear()
        With dst.Tables("SATANALC").Rows
            For Each COLUMN_NAME As String In New String() {"TRADE_CLASS_CODE", "CUST_CODE", "CUST_STORE_NO"}
                Dim COLUMN_NAME2 As String = COLUMN_NAME
                Select Case COLUMN_NAME
                    Case "TRADE_CLASS_CODE"
                        ASCMAIN1.sql = "Select TRADE_CLASS_CODE, TRADE_CLASS_DESC from SOTTCLS1"
                    Case "CUST_CODE"
                        ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME from ARTCUST1"
                    Case "CUST_STORE_NO"
                        ASCMAIN1.sql = "Select CUST_STORE_NO, CUST_STORE_NAME from ARTCUST2 where CUST_CODE = 'IPLBAE'"
                End Select

                Dim DT As DataTable = ASCDATA1.GetDataTable
                .Add(New String() {COLUMN_NAME2, COLUMN_NAME2, DT.Columns(1).ColumnName, DT.TableName})

                For Each row As DataRow In DT.Rows
                    Dim rowSATANALD As DataRow = dst.Tables("SATANALD").Rows.Find(New String() {COLUMN_NAME2, row.Item(0)})
                    If rowSATANALD Is Nothing Then
                        rowSATANALD = dst.Tables("SATANALD").NewRow
                        rowSATANALD.Item("COLUMN_NAME") = COLUMN_NAME2
                        rowSATANALD.Item("CODE_VALUE") = row.Item(0)
                        rowSATANALD.Item("DESC_VALUE") = row.Item(1)
                        dst.Tables("SATANALD").Rows.Add(rowSATANALD)
                    End If
                Next

            Next

        End With


        grdSOTALLO1S.DataSource = dst.Tables("SOTALLO1S")
        grdSOTALLOCS.DataSource = dst.Tables("SOTALLOCS")

        grdSOTALLOX.DataSource = dst.Tables("SOTALLOX")
        grdICTITEM1.DataSource = dst.Tables("ICTITEM1")
        grdSOTORDR2.DataSource = dst.Tables("SOTORDR2")
        grdSOTORDRA.DataSource = dst.Tables("SOTORDRA")

        grdSOTALLO1.DataSource = dst.Tables("SOTALLO1")
        grdSOTALLOC.DataSource = dst.Tables("SOTALLOC")
        grdSOTALLOS.DataSource = dst.Tables("SOTALLOS")
        grdSOTALLOT.DataSource = dst.Tables("SOTALLOT")

        grdSOTORDRD.DataSource = dst.Tables("SOTORDRD")
        grdSOTORDRU.DataSource = dst.Tables("SOTORDRU")

        grdSOTNGMSG.DataSource = dst.Tables("SOTNGMSG")
        Sort_grdColumns(grdSOTNGMSG, "init_date")

        Create_Summary(grdSOTORDRD, "ITEM_CODE", "Count")
        Create_Summary(grdSOTORDRD, "ADDS")
        Create_Summary(grdSOTORDRD, "DEDS")

        grdSOTALLH1.DataSource = dst.Tables("SOTALLH1")
        Create_Summary(grdSOTALLH1, "XNO", "Count")
        Sort_grdColumns(grdSOTALLH1, "last_date")

        Create_Summary(grdSOTALLOX, "ITEM_CODE", "Count")
        Create_Summary(grdSOTALLOX, "SELECTED", "Sum")

        Create_Summary(grdICTITEM1, "ITEM_CODE", "Count")
        Create_Summary(grdICTITEM1, "SELECTED", "Sum")

        Create_Summary(grdSOTORDR2, "SELECTED", "Sum")
        Create_Summary(grdSOTORDR2, "ORDR_NO", "Count")
        Create_Summary(grdSOTORDR2, New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "USED"})

        Create_Summary(grdSOTORDRA, "SELECTED", "Sum")
        Create_Summary(grdSOTORDRA, "ORDR_GROUP_NO", "Count")
        Create_Summary(grdSOTORDRA, New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "USED"})

        grdSOTALLO1S.DisplayLayout.Bands(1).Override.SummaryValueAppearance.BackColor = Drawing.Color.LightGray
        Create_Summary(grdSOTALLO1S, "CUST_CODE", "Count", "SOTALLO1S_SOTALLO2S")
        Create_Summary(grdSOTALLO1S, New String() {"QTY_ALLO", "ORDR_QTY_SHIP", "ORDR_QTY_PICK", "ORDR_QTY_OPEN", "QTY_ALLO_BAL"}, "Sum", "SOTALLO1S_SOTALLO2S")
        Create_Summary(grdSOTALLO1S, "TRADE_CLASS_CODE", "CustomString", "SOTALLO1S_SOTALLO2S")
        Create_Summary(grdSOTALLO1S, "QTY_ALLO_USED", "Custom", "SOTALLO1S_SOTALLO2S")


        For iCtr As Integer = 1 To 10
            Dim colKey As String = "EVENT_" & Format(iCtr, "00")
            If grdSOTALLOT.DisplayLayout.Bands(0).Columns.Exists(colKey) Then
                Create_Summary(grdSOTALLOT, colKey, "Sum")
            End If
        Next

        For Each GCOL As UltraWinGrid.UltraGridColumn In grdSOTORDR2.DisplayLayout.Bands(0).Columns
            If GCOL.Key = "SELECTED" Then
                GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                GCOL.CellAppearance.BackColor = System.Drawing.Color.Yellow
            Else
                GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next

        For Each GCOL As UltraWinGrid.UltraGridColumn In grdSOTORDRA.DisplayLayout.Bands(0).Columns
            If GCOL.Key = "SELECTED" Then
                GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                GCOL.CellAppearance.BackColor = System.Drawing.Color.Yellow
            Else
                GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next

        grdSOTALLOX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdSOTALLOX.DisplayLayout.Override.ExpansionIndicator = UltraWinGrid.ShowExpansionIndicator.CheckOnDisplay
        grdSOTALLOX.DisplayLayout.Bands(0).Override.CellClickAction = UltraWinGrid.CellClickAction.CellSelect
        grdSOTALLOX.DisplayLayout.Bands(0).Columns("SELECTED").CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        With grdSOTALLOX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SELECTED", "ITEM_CODE", "ITEM_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key = "SELECTED" Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                    GCOL.CellAppearance.BackColor = System.Drawing.Color.GhostWhite
                End If
                GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If GCOL.Key = "ITEM_CODE" Or GCOL.Key = "ALLO_GROUP_CODE" Or col_ICTITEM1.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                ElseIf New String() {"SELECTED"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Gold
                ElseIf New String() {"DATE_START", "DATE_END", "ALLO_CTL_NO", "ALLOW_OVER", "ITEM_CODE_COMPARE_TO", "ITEM_CODE_COMPARE_TO_ALT"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                End If
            Next
        End With

        grdICTITEM1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdICTITEM1.DisplayLayout.Override.ExpansionIndicator = UltraWinGrid.ShowExpansionIndicator.CheckOnDisplay
        grdICTITEM1.DisplayLayout.Bands(0).Override.CellClickAction = UltraWinGrid.CellClickAction.CellSelect
        grdICTITEM1.DisplayLayout.Bands(0).Columns("SELECTED").CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        With grdICTITEM1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SELECTED", "ITEM_CODE", "ITEM_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key = "SELECTED" Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                    GCOL.CellAppearance.BackColor = System.Drawing.Color.GhostWhite
                End If
                GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If GCOL.Key = "ITEM_CODE" Or col_ICTITEM1.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                ElseIf New String() {"SELECTED"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Gold
                ElseIf New String() {"DATE_START", "DATE_END", "ALLO_CTL_NO"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                End If
            Next
        End With


        With grdSOTALLO1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ITEM_CODE", "ITEM_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key = "DATE_START" Or GCOL.Key = "DATE_END" _
                    Or GCOL.Key = "ALLOW_OVER" Or GCOL.Key = "QTY_ALLO_PLAN" _
                    Or GCOL.Key = "ITEM_CODE_COMPARE_TO" Or GCOL.Key = "ITEM_CODE_COMPARE_TO_ALT" Or GCOL.Key = "ALLO_GROUP_CODE" Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                    GCOL.CellAppearance.BackColor = System.Drawing.Color.GhostWhite ' WhiteSmoke
                End If
                GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If GCOL.Key = "ITEM_CODE" Or col_ICTITEM1.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                ElseIf New String() {"DATE_START", "DATE_END", "ALLO_CTL_NO", "ALLOW_OVER", "ITEM_CODE_COMPARE_TO", "ITEM_CODE_COMPARE_TO_ALT", "ALLO_GROUP_CODE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                End If
            Next
        End With

        With grdSOTALLOC.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_CODE", "CUST_NAME", "SREP_CODE", "TRADE_CLASS_CODE", "STORE_COUNT_HC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = System.Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
            Next
            Add_ALLO_Columns(grdSOTALLOC)
        End With
        grdSOTALLOC.DisplayLayout.Override.ActiveRowAppearance.BackColor = System.Drawing.Color.PaleGreen
        Create_Summary(grdSOTALLOC, "STORE_COUNT_HC")

        With grdSOTALLOS.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_STORE_NO", "CUST_STORE_LOCATION", "SELL_CODE", "DMA_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = System.Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
            Next
            .Columns("CUST_STORE_NO").Width = grdSOTALLOC.DisplayLayout.Bands(0).Columns("CUST_CODE").Width
            .Columns("CUST_STORE_LOCATION").Width = grdSOTALLOC.DisplayLayout.Bands(0).Columns("CUST_NAME").Width
            .Columns("SELL_CODE").Width = grdSOTALLOC.DisplayLayout.Bands(0).Columns("SREP_CODE").Width
            .Columns("DMA_CODE").Width = grdSOTALLOC.DisplayLayout.Bands(0).Columns("TRADE_CLASS_CODE").Width
            Add_ALLO_Columns(grdSOTALLOS)
        End With

        With grdSOTALLOT.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_STORE_NO", "CUST_STORE_LOCATION", "SELL_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = System.Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
            Next
            .Columns("CUST_STORE_NO").Width = grdSOTALLOC.DisplayLayout.Bands(0).Columns("CUST_CODE").Width
            .Columns("CUST_STORE_LOCATION").Width = grdSOTALLOC.DisplayLayout.Bands(0).Columns("CUST_NAME").Width
            For iCtr As Integer = 1 To 10
                Dim colKey As String = "EVENT_" & Format(iCtr, "00")
                If .Columns.Exists(colKey) Then
                    .Columns(colKey).Hidden = True
                End If
            Next
        End With


        With grdSOTALLO1S.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                With gcol.Header.Appearance
                    .BackColor = System.Drawing.Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                    .BackColor2 = System.Drawing.Color.LightBlue
                End With
            Next
            .Columns("QTY_ALLO").CellAppearance.BackColor = Drawing.Color.PaleGreen
        End With
        With grdSOTALLO1S.DisplayLayout.Bands(1)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                With gcol.Header.Appearance
                    .BackColor = System.Drawing.Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                    .BackColor2 = System.Drawing.Color.LightGray
                End With
            Next
            .Columns("QTY_ALLO").CellAppearance.BackColor = Drawing.Color.PaleGreen
        End With


        With grdSOTALLOCS.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                With gcol.Header.Appearance
                    .BackColor = System.Drawing.Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                    .BackColor2 = System.Drawing.Color.LightGray
                End With
            Next
            .Columns("QTY_ALLO").CellAppearance.BackColor = Drawing.Color.PaleGreen
            Add_ALLO_Columns2(grdSOTALLOCS)
        End With


        ASCMAIN1.Add_Value_List(grdSOTALLOX, "ITEM_CATGY_CODE")
        ASCMAIN1.Add_Value_List(grdSOTALLOX, "ITEM_STATUS")
        ASCMAIN1.Add_Value_List(grdICTITEM1, "ITEM_CATGY_CODE")
        ASCMAIN1.Add_Value_List(grdICTITEM1, "ITEM_STATUS")

        ASCMAIN1.Add_Value_List(grdSOTALLO1, "ITEM_CATGY_CODE")
        ASCMAIN1.Add_Value_List(grdSOTALLO1, "ITEM_STATUS")

        ASCMAIN1.Add_Value_List(grdSOTALLO1, "ALLO_GROUP_CODE", "Select ALLO_GROUP_CODE, ALLO_GROUP_CODE from SOTALLOG where ALLO_GROUP_STATUS = 'A'")

        '  grdSOTALLO1S.DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.OutlookGroupBy
        ' Show_Filter(grdSOTALLO1S, True)
        ' grdSOTALLO1S.DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.Horizontal

        If ASCMAIN1.CLIENT = "INT" Then
            ' If today is 07/01 thru 12/31, then use 06/01 of the current year,
            ' If today is 01/01 thru 06/30, then use 12/01 of the prior year.
            If Now.Date.Month >= 7 Then
                dteEndDate.DateTime = Format(Now.Date, "06/01/yyyy")
            Else
                dteEndDate.DateTime = Format(Now.Date.AddYears(-1), "12/01/yyyy")
            End If
        Else
            If Now.Date.Month >= 7 Then
                dteEndDate.DateTime = Format(Now.Date, "07/01/yyyy")
            Else
                dteEndDate.DateTime = Format(Now.Date, "01/01/yyyy")
            End If
        End If

        ' dteEndDate.DateTime = Now.Date.AddDays(-90)
        dteAllocations.DateTime = Now.Date.AddDays(0)

        chkUseSSG.Visible = False
        tabSOTALLOX.Tabs("Net Changes").Visible = False

        cmbYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cmbYP.SelectedRow = cmbYP.Rows(0)

        If InquiryMode Then
            tabSOTALLOX.Tabs("Items").Visible = False
            tabSOTALLOX.Tabs("Changes at Shipment").Visible = False

            With grdSOTALLO1.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With

            With grdSOTALLOC.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With

            With grdSOTALLOS.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With

            With grdSOTALLOT.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With

            'With grdSOTALLOX.DisplayLayout.Override
            '    .AllowAddNew = UltraWinGrid.AllowAddNew.No
            '    .AllowUpdate = DefaultableBoolean.False
            '    .AllowDelete = DefaultableBoolean.False
            'End With
        End If

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                If tabSOTALLOX.SelectedTab.Key = "Allocations" Then
                    Dim SELs As Integer = dst.Tables("SOTALLOX").Select("SELECTED='1'").Length
                    If SELs = 0 Then
                        EMsg &= "You clicked 'New' but you have not chosen specific Allocations to Clone"
                    Else

                        'For Each row As DataRow In dst.Tables("SOTALLOX").Select("SELECTED='1'")
                        '    Dim ITEM_CODE As String = row.Item("ITEM_CODE")
                        '    If dst.Tables("SOTALLOX").Select("ITEM_CODE = '" & ITEM_CODE & "' and SELECTED='1'").Length > 1 Then
                        '        EMsg &= "Item " & ITEM_CODE & " appears in multiple Allocations Selected to Clone"
                        '        Exit For
                        '    End If
                        'Next

                        If EMsg = "" Then
                            If MsgBox("You have chosen " & CStr(SELs) & " Allocations to Clone." _
                                      & vbCrLf & vbCrLf & "OK to Proceed?",
                                      MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                EMsg &= "No action taken"
                            End If
                        End If

                    End If

                ElseIf tabSOTALLOX.SelectedTab.Key = "Items" Then
                    If dst.Tables("ICTITEM1").Select("SELECTED = '1'").Length = 0 Then
                        EMsg = "There are no Items Selected"
                    End If

                    If dst.Tables("ICTITEM1").Select("SELECTED = '1'").Length > maxAllocations Then
                        EMsg = "You may load a maximum of " & CStr(maxAllocations) & " Items"
                    End If

                    If EMsg = "" Then
                        For Each row As DataRow In dst.Tables("ICTITEM1").Select("SELECTED = '1'")
                            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
                            If Not ASCMAIN1.Logical_Lock("SOTALLO1", "ITEM:" & ITEM_CODE) Then
                                Exit Sub
                            End If
                        Next
                    End If

                Else
                    EMsg &= "You must choose Items from the Items tab to start a New Allocation"
                End If

            Case "Edit", "View"
                If tabSOTALLOX.SelectedTab.Key <> "Allocations" Then
                    EMsg &= "You must choose Allocations from the Allocations tab to " & eItemKey & " an Allocation"
                Else

                    If dst.Tables("SOTALLOX").Select("SELECTED = '1'").Length = 0 Then
                        EMsg = "There are no Allocations Selected"
                    End If

                    If dst.Tables("SOTALLOX").Select("SELECTED = '1'").Length > maxAllocations Then
                        EMsg = "You may load a maximum of " & CStr(maxAllocations) & " Allocations"
                    End If

                    If EMsg = "" Then
                        If eItemKey = "Edit" Then
                            For Each row As DataRow In dst.Tables("SOTALLOX").Select("SELECTED = '1'")
                                Dim ALLO_CTL_NO As String = row.Item("ALLO_CTL_NO")
                                If Not ASCMAIN1.Logical_Lock("SOTALLO1", ALLO_CTL_NO) Then
                                    Exit Sub
                                End If
                                Dim ITEM_CODE As String = row.Item("ITEM_CODE")
                                If Not ASCMAIN1.Logical_Lock("SOTALLO1", "ITEM:" & ITEM_CODE) Then
                                    Exit Sub
                                End If
                            Next
                        End If
                    End If
                End If

            Case "Cancel"
                If MsgBox("Are you sure you want to Cancel your changes?",
                            MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Update"

                grdSOTALLO1.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
                grdSOTALLOC.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
                If tabSOTALLOC.Tabs("Store Allocations").Visible Then
                    grdSOTALLOS.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
                End If

                Dim ITEM_SO_QTY_MULTs As New Dictionary(Of String, Int32)

                Dim ITEM_NOT_ALLOCATEDs As New List(Of String)
                ASCMAIN1.MultiTask_Release(,, 1)

                Dim dateError As Boolean = False
                For Each rowSOTALLO1 As DataRow In dst.Tables("SOTALLO1").Select("", "ITEM_CODE,ALLO_CTL_NO")
                    Dim ALLO_CTL_NO As String = rowSOTALLO1.Item("ALLO_CTL_NO")
                    Dim ITEM_CODE As String = rowSOTALLO1.Item("ITEM_CODE")

                    If Not ASCMAIN1.Logical_Lock("ICTITEM1", ITEM_CODE, ,,, 1) Then Exit Sub

                    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                    If rowICTITEM1.Item("ITEM_NOT_ALLOCATED") & "" = "1" Then
                        ITEM_NOT_ALLOCATEDs.Add(ITEM_CODE)
                    End If

                    Dim ITEM_SO_QTY_MULT As Int32 = Val(rowSOTALLO1.Item("ITEM_SO_QTY_MULT") & "")
                    ITEM_SO_QTY_MULTs.Add(ALLO_CTL_NO, ITEM_SO_QTY_MULT)

                    Dim ITEM_CODE_COMPARE_TO As String = rowSOTALLO1.Item("ITEM_CODE_COMPARE_TO") & ""
                    Dim ITEM_CODE_COMPARE_TO_ALT As String = rowSOTALLO1.Item("ITEM_CODE_COMPARE_TO_ALT") & ""

                    If ITEM_CODE_COMPARE_TO <> "" Then If LookUp("ICTITEM1", ITEM_CODE_COMPARE_TO) Is Nothing Then EMsg &= vbCr & "Compare to Item is Invalid for Allocation " & ALLO_CTL_NO
                    If ITEM_CODE_COMPARE_TO_ALT <> "" Then If LookUp("ICTITEM1", ITEM_CODE_COMPARE_TO_ALT) Is Nothing Then EMsg &= vbCr & "Compare to Alternate Item is Invalid for Allocation " & ALLO_CTL_NO

                    dateError = False
                    If Not IsDate(rowSOTALLO1.Item("DATE_START") & String.Empty) Then
                        EMsg &= Environment.NewLine & "Missing or Invalid Start Date for Allocation: " & rowSOTALLO1.Item("ALLO_CTL_NO")
                        dateError = True
                    End If

                    If Not IsDate(rowSOTALLO1.Item("DATE_END") & String.Empty) Then
                        EMsg &= Environment.NewLine & "Missing or Invalid End Date for Allocation: " & rowSOTALLO1.Item("ALLO_CTL_NO")
                        dateError = True
                    End If

                    If Not dateError Then

                        Dim DATE_START As DateTime = rowSOTALLO1.Item("DATE_START")
                        Dim DATE_END As DateTime = rowSOTALLO1.Item("DATE_END")

                        If DateDiff(DateInterval.Day, rowSOTALLO1.Item("DATE_START"), rowSOTALLO1.Item("DATE_END")) < 0 Then
                            EMsg &= Environment.NewLine & "End Date is prior to Start Date for Allocation: " & rowSOTALLO1.Item("ALLO_CTL_NO")
                        Else
                            ASCMAIN1.sql = "Select * from SOTALLO1" & vbCrLf _
                                & " where ITEM_CODE = :PARM1" & vbCrLf _
                                & " and ALLO_CTL_NO <> :PARM2" & vbCrLf _
                                & " and (:PARM3 between DATE_START and DATE_END or :PARM4 between DATE_START and DATE_END)"
                            Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVDD",
                                                                     New Object() {ITEM_CODE,
                                                                                   ALLO_CTL_NO,
                                                                                   rowSOTALLO1.Item("DATE_START"),
                                                                                   rowSOTALLO1.Item("DATE_END")})

                            If row Is Nothing Then
                                ' THIS CHECK WILL FIND IF OUR ALLOCATION DATES TOTALLY ENCOMPASS START/END DATES OF OTHER ALLOCAITONS
                                ASCMAIN1.sql = "Select * from SOTALLO1" & vbCrLf _
                                    & " where ITEM_CODE = :PARM1" & vbCrLf _
                                    & " and ALLO_CTL_NO <> :PARM2" & vbCrLf _
                                    & " and (:PARM3 < DATE_START and :PARM4 > DATE_END)"
                                row = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVDD",
                                                                         New Object() {ITEM_CODE,
                                                                                       ALLO_CTL_NO,
                                                                                       rowSOTALLO1.Item("DATE_START"),
                                                                                       rowSOTALLO1.Item("DATE_END")})
                            End If

                            If row IsNot Nothing Then
                                EMsg &= vbCr & "Allocation " & ALLO_CTL_NO & " for Item " & ITEM_CODE & vbCrLf _
                                    & " (Date Range " & Format(rowSOTALLO1.Item("DATE_START"), "MM/dd/yyyy") & " thru " & Format(rowSOTALLO1.Item("DATE_END"), "MM/dd/yyyy") & ")" & vbCrLf _
                                    & " is within Date Range of Allocation " & row.Item("ALLO_CTL_NO") & vbCrLf _
                                    & " (Date Range " & Format(row.Item("DATE_START"), "MM/dd/yyyy") & " thru " & Format(row.Item("DATE_END"), "MM/dd/yyyy") & ")"
                            Else
                                For Each row2 As DataRow In dst.Tables("SOTALLO1").Select("ITEM_CODE = '" & ITEM_CODE & "' and ALLO_CTL_NO <> '" & ALLO_CTL_NO & "' and DATE_START is Not Null and DATE_END is Not Null")
                                    Dim DATE_START_ymd As String = Format(rowSOTALLO1.Item("DATE_START"), "yyyyMMdd")
                                    Dim DATE_END_ymd As String = Format(rowSOTALLO1.Item("DATE_END"), "yyyyMMdd")
                                    If DATE_START_ymd >= Format(row2.Item("DATE_START"), "yyyyMMdd") And DATE_START_ymd <= Format(row2.Item("DATE_END"), "yyyyMMdd") _
                                    Or DATE_END_ymd >= Format(row2.Item("DATE_START"), "yyyyMMdd") And DATE_END_ymd <= Format(row2.Item("DATE_END"), "yyyyMMdd") Then
                                        EMsg &= vbCr & "Allocation " & ALLO_CTL_NO & " for Item " & ITEM_CODE & vbCrLf _
                                              & " (Date Range " & Format(rowSOTALLO1.Item("DATE_START"), "MM/dd/yyyy") & " thru " & Format(rowSOTALLO1.Item("DATE_END"), "MM/dd/yyyy") & ")" & vbCrLf _
                                              & " is within Date Range of Allocation " & row2.Item("ALLO_CTL_NO") & vbCrLf _
                                              & " (Date Range " & Format(row2.Item("DATE_START"), "MM/dd/yyyy") & " thru " & Format(row2.Item("DATE_END"), "MM/dd/yyyy") & ")"

                                    End If
                                Next

                                Dim ictr As Int64 = 0
                                For i As Integer = 1 To ALLO_CTL_NOi.Count
                                    If ALLO_CTL_NOi(i) = ALLO_CTL_NO Then
                                        ictr = i
                                        Exit For
                                    End If
                                Next

                                ' check to see if over allocation
                                Dim QTY_ALLO_orig As Int64 = 0
                                Dim QTY As Int64 = 0
                                Fill_Records("SOTORDRQ", New Object() {ITEM_CODE, ALLO_CTL_NO, DATE_START, DATE_END})
                                Dim sqlw As String = $"ITEM_CODE = '{ITEM_CODE}' and ALLO_CTL_NO = '{ALLO_CTL_NO}'"
                                For Each rowSOTALLOC As DataRow In dst.Tables("SOTALLOC").Select("")
                                    Dim CUST_CODE As String = rowSOTALLOC.Item("CUST_CODE")
                                    Dim rowSOTALLO2 As DataRow = dst.Tables("SOTALLO2").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE})
                                    If rowSOTALLO2 Is Nothing Then
                                        QTY_ALLO_orig = 0
                                    Else
                                        QTY_ALLO_orig = Val(rowSOTALLO2.Item("QTY_ALLO") & "")
                                    End If

                                    QTY = Val(rowSOTALLOC.Item("ALLO_" & Format(ictr, "00")) & "")

                                    Dim sqlx As String = sqlw & $" and CUST_CODE_ALLO = '{CUST_CODE}'"
                                    Dim ORDR_QTY_SHIP As Int64 = Val(dst.Tables("SOTORDRQ").Compute("SUM(ORDR_QTY_SHIP)", sqlx) & "")
                                    Dim ORDR_QTY_PICK As Int64 = Val(dst.Tables("SOTORDRQ").Compute("SUM(ORDR_QTY_PICK)", sqlx) & "")
                                    If QTY < ORDR_QTY_PICK + ORDR_QTY_SHIP And QTY <> QTY_ALLO_orig Then
                                        EMsg &= vbCr & $"Item {ITEM_CODE} Allocation Qty for {CUST_CODE} ({QTY})" & vbCrLf _
                                            & $" is less than Qty Shipped ({ORDR_QTY_SHIP}) + Qty In Pick ({ORDR_QTY_PICK})"
                                    End If

                                    If rowSOTALLOC.Item("CUST_ALLOCATE_BY_STORE") & "" = "1" Then
                                        For Each rowSOTALLOS As DataRow In dst.Tables("SOTALLOS").Select("CUST_CODE = '" & CUST_CODE & "'")
                                            Dim CUST_STORE_NO As String = rowSOTALLOS.Item("CUST_STORE_NO")
                                            Dim rowSOTALLO3 As DataRow = dst.Tables("SOTALLO3").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE, CUST_STORE_NO})
                                            If rowSOTALLO3 Is Nothing Then
                                                QTY_ALLO_orig = 0
                                            Else
                                                QTY_ALLO_orig = Val(rowSOTALLO3.Item("QTY_ALLO") & "")
                                            End If

                                            QTY = Val(rowSOTALLOS.Item("ALLO_" & Format(ictr, "00")) & "")

                                            sqlx = sqlw & $" and CUST_CODE = '{CUST_CODE}' and CUST_STORE_NO = '{CUST_STORE_NO}'"
                                            ORDR_QTY_SHIP = Val(dst.Tables("SOTORDRQ").Compute("SUM(ORDR_QTY_SHIP)", sqlx) & "")
                                            ORDR_QTY_PICK = Val(dst.Tables("SOTORDRQ").Compute("SUM(ORDR_QTY_PICK)", sqlx) & "")
                                            If QTY < ORDR_QTY_PICK + ORDR_QTY_SHIP And QTY <> QTY_ALLO_orig Then
                                                EMsg &= vbCr & $"Item {ITEM_CODE} Allocation Qty for {CUST_CODE} {CUST_STORE_NO} ({QTY})" & vbCrLf _
                                                    & $" is less than Qty Shipped ({ORDR_QTY_SHIP}) + Qty In Pick ({ORDR_QTY_PICK})"
                                            End If

                                        Next
                                    End If
                                Next

                            End If
                        End If
                    End If
                Next

                For ictr As Integer = 1 To iColumn


                    If ALLO_CTL_NOi(ictr) <> "" Then

                        Dim ITEM_SO_QTY_MULT As Int32 = Val(ITEM_SO_QTY_MULTs(ALLO_CTL_NOi(ictr)))

                        Dim sql As String = "ALLO_" & Format(ictr, "00") & " < 0"
                        Dim rows() As DataRow
                        rows = dst.Tables("SOTALLOC").Select(sql)
                        If rows.Length <> 0 Then
                            EMsg &= vbCr & "Negative Allocation for " & rows(0).Item("CUST_CODE")
                            Exit For
                        End If

                        If ITEM_SO_QTY_MULT > 1 Then
                            For Each row As DataRow In dst.Tables("SOTALLOC").Select("")
                                Dim QTY As Int64 = Val(row.Item("ALLO_" & Format(ictr, "00")) & "")
                                If QTY Mod ITEM_SO_QTY_MULT <> 0 Then
                                    EMsg &= vbCr & "Qty Allocated Is Not evenly divisible by Order Multiple (" & CStr(ITEM_SO_QTY_MULT) & ") for " & row.Item("CUST_CODE")
                                    Exit For
                                End If
                            Next
                        End If

                        rows = dst.Tables("SOTALLOS").Select(sql)
                        If rows.Length <> 0 Then
                            EMsg &= vbCr & "Negative Allocation for " & rows(0).Item("CUST_CODE") & " - " & rows(0).Item("CUST_STORE_NO")
                            Exit For
                        End If

                        If ITEM_SO_QTY_MULT > 1 Then
                            For Each row As DataRow In dst.Tables("SOTALLOS").Select("")
                                Dim QTY As Int64 = Val(row.Item("ALLO_" & Format(ictr, "00")) & "")
                                If QTY Mod ITEM_SO_QTY_MULT <> 0 Then
                                    EMsg &= vbCr & "Qty Allocated Is Not evenly divisible by Order Multiple (" & CStr(ITEM_SO_QTY_MULT) & ") for " & row.Item("CUST_CODE") & " - " & row.Item("CUST_STORE_NO")
                                    Exit For
                                End If
                            Next
                        End If

                    End If
                Next

                For Each rowSOTALLOC As DataRow In dst.Tables("SOTALLOC").Select
                    If LookUp("ARTCUST1", rowSOTALLOC.Item("CUST_CODE")) Is Nothing Then
                        EMsg &= Environment.NewLine & "Invalid Customer Code: " & rowSOTALLOC.Item("CUST_CODE")
                    End If
                Next

                If ITEM_NOT_ALLOCATEDs.Count > 0 Then
                    If EMsg = "" Then
                        If MsgBox("The following items are flagged as Not Allocated:" & vbCrLf & Join(ITEM_NOT_ALLOCATEDs.ToArray, ",") _
                                  & vbCrLf & vbCrLf & "If you continue with this update," & vbCrLf & " the 'Not Allocated' flags for these items will be cleared" _
                                  & vbCrLf & vbCrLf & "Continue?", MsgBoxStyle.YesNo, "Items are Flagged as Not Allocated") = MsgBoxResult.No Then
                            EMsg &= vbCr & "Items are flagged as Not Allocated: " & Join(ITEM_NOT_ALLOCATEDs.ToArray, ",")
                        End If
                    Else
                        EMsg &= vbCr & "Items are flagged as Not Allocated: " & Join(ITEM_NOT_ALLOCATEDs.ToArray, ",")
                    End If
                End If

                If EMsg = "" Then
                    If ALLO_CTL_NOs_to_Delete.Count <> 0 Then
                        Dim sqlw As String = " where ALLO_CTL_NO in ('" & Join(ALLO_CTL_NOs_to_Delete.ToArray, "','") & "')"
                        ASCMAIN1.sql = "Select * from SOTALLO1" & sqlw
                        Dim dt As DataTable = ASCDATA1.GetDataTable
                        Using frmmsg As New ASFMSGBF
                            frmmsg.Show_grd(dt, Me, "As part of this Update, the following Allocations will be Deleted")
                            ' frmmsg = Nothing
                        End Using
                        If MsgBox("As part of this Update," & vbCrLf & " there will be " & CStr(ALLO_CTL_NOs_to_Delete.Count) & " Allocations Deleted" & vbCrLf & vbCrLf & "OK to Proceed?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Maintain Sales"
                If Not grdSOTORDR2.Visible Then
                    Exit Sub
                End If

        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
                MyBase.EntryMode = "N"
                Me.Load_Record()
                Me.Mode_Settings(True)

            Case "Edit", "View"
                If eItemKey = "Edit" Then
                    MyBase.EntryMode = "E"
                Else
                    MyBase.EntryMode = "V"
                End If

                Me.Load_Record()
                Me.Mode_Settings(True)

            Case "Cancel", "Done"
                Me.Mode_Settings(False)
                tabSOTALLOX.Tabs("Net Changes").Visible = False

            Case "Print"
                Me.Print_Record()

            Case "Update"
                If dst.Tables("SOTALLO1").Rows.Count > 0 Then
                    Dim lastRow As DataRow = Nothing
                    For Each row As DataRow In dst.Tables("SOTALLO1").Rows
                        If row.RowState <> DataRowState.Deleted Then
                            lastRow = row
                        End If
                    Next

                    If lastRow IsNot Nothing Then
                        lastStart = lastRow("DATE_START")
                        lastEnd = lastRow("DATE_END")
                        datesSet = True
                        btnCopyDates.Text = $"Copy Dates: {lastStart:MM/dd/yyyy} - {lastEnd:MM/dd/yyyy}"
                    End If
                End If
                Me.Update_Record()
                Me.Mode_Settings(False)

                If dst.Tables("SOTORDRD").Rows.Count <> 0 Then
                    tabSOTALLOX.Tabs("Net Changes").Visible = True
                    tabSOTALLOX.SelectedTab = tabSOTALLOX.Tabs("Net Changes")
                    MsgBox("Please Note Net Changes" _
                           & vbCrLf & "The contents of this grid will be cleared with next screen activity" _
                           & vbCrLf & "So export to Excel if you wish to save the contents",
                           MsgBoxStyle.OkOnly, "Verification")
                Else
                    tabSOTALLOX.Tabs("Net Changes").Visible = False
                End If


            Case "Maintain Sales"
                Toggle_Maintain_Sales(True)

            Case "Update Changes"
                dst.Tables("SOTORDR2").AcceptChanges()
                Dim sqlw As String = "(ISNULL(SELECTED,'0') = '0' and ISNULL(ALLO_CTL_NO,'') <> '') or (ISNULL(SELECTED,'0') = '1' and ISNULL(ALLO_CTL_NO,'') = '')"
                For Each ROW As DataRow In dst.Tables("SOTORDR2").Select(sqlw)
                    If ROW.Item("SELECTED") & "" = "1" Then
                        ROW.Item("ALLO_CTL_NO") = rowSOTALLO1_Maintain_Sales.Item("ALLO_CTL_NO")
                    Else
                        ROW.Item("ALLO_CTL_NO") = DBNull.Value
                    End If
                Next

                BeginTrans()
                Update_Record_TDA("SOTORDR2")
                CommitTrans()

                Toggle_Maintain_Sales(False)

            Case "Cancel Changes"
                Toggle_Maintain_Sales(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode
                    .Items("Save Orig").Settings.Enabled = iScreenMode

                    .Items("New").Visible = Not InquiryMode

                    .Items("Edit").Visible = Not InquiryMode
                    .Items("View").Visible = InquiryMode

                    .Items("Update").Visible = Not InquiryMode
                    .Items("Cancel").Visible = Not InquiryMode
                    .Items("Done").Visible = InquiryMode

                    .Items("Save Orig").Visible = False
                    .Items("Print").Visible = False ' sp had an error clicking on this - was wondering why it was even an option in the screen
                End With

                '.Groups("Items").Visible = Not ScreenMode
                lblITEM_CODE.Visible = Not ScreenMode
                txtITEM_CODE.Visible = Not ScreenMode

                .Groups("Customers").Visible = ScreenMode And Not InquiryMode
                .Groups("Items").Visible = Not ScreenMode
                .Groups("Display Options").Visible = ScreenMode
                .Groups("Sales").Visible = False
                .Groups("Add Event Tags").Visible = ScreenMode
            End With
        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)

        splSOTALLO1.Visible = ScreenMode
        splSOTALLO1.SplitterDistance = splSOTALLO1.Height * 2 / 3
        tabSOTALLOX.Visible = Not ScreenMode
        btnCopyDates.Visible = ScreenMode And datesSet

        If ScreenMode Then
            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTALLO1, grdSOTALLOC, grdSOTALLOS}
                With grd.DisplayLayout.Override
                    If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
                        If grd.Name = "grdSOTALLOC" Or grd.Name = "grdSOTALLOS" Then
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            '.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                        Else
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        End If
                        '  .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        If Not InquiryMode Then
                            .AllowUpdate = DefaultableBoolean.True
                        End If

                        .AllowDelete = DefaultableBoolean.False
                    Else
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowUpdate = DefaultableBoolean.False
                        .AllowDelete = DefaultableBoolean.False
                    End If
                End With
            Next
        End If

        spl.Panel1Collapsed = ScreenMode
        Toggle_Display_Options()

        If ScreenMode Then
            grdSOTALLOX.Parent = tabSOTALLOC.Tabs("Other Allocations").TabPage
            grdICTITEM1.Parent = tabSOTALLOC.Tabs("Other Items").TabPage
            grdSOTALLOX.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = True
            grdICTITEM1.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = True

            tabSOTALLOC.SelectedTab = tabSOTALLOC.Tabs(0)
            Setup_tabSOTALLOC()
            Toggle_Maintain_Sales(False)

            'grdSOTALLO1.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText

            grdSOTALLOC.DisplayLayout.Bands(0).Columns("STORE_COUNT_HC").Header.Caption = "#Strs" & vbCrLf & HC_CODE_lead_item
        Else
            Me.Clear_Record()

            chkSingleItemMode.Checked = False
            grdSOTALLOX.Parent = tabSOTALLOX.Tabs("Allocations").TabPage
            grdICTITEM1.Parent = tabSOTALLOX.Tabs("Items").TabPage
            grdSOTALLOX.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = False
            grdICTITEM1.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = False

            'grdSOTALLO1.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.RowSelect

            Setup_tabSOTALLOX()

            If tabSOTALLOX.Tabs("Status by Item").Tag & "" <> "" Or tabSOTALLOX.Tabs("Status by Customer").Tag & "" <> "" Then
                Dim ALLO_CTL_NO As String = tabSOTALLOX.Tabs("Status by Item").Tag & tabSOTALLOX.Tabs("Status by Customer").Tag
                For Each row As DataRow In dst.Tables("SOTALLO2S").Select("ALLO_CTL_NO = '" & ALLO_CTL_NO & "'")
                    row.Item("QTY_ALLO") = 0
                Next
                ASCMAIN1.sql = "Select * from SOTALLO2 where ALLO_CTL_NO = '" & ALLO_CTL_NO & "'"
                For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim CUST_CODE As String = row.Item("CUST_CODE")
                    Dim rowSOTALLO2S As DataRow = dst.Tables("SOTALLO2S").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE})
                    If rowSOTALLO2S Is Nothing Then
                        rowSOTALLO2S = dst.Tables("SOTALLO2S").NewRow
                        rowSOTALLO2S.Item("ALLO_CTL_NO") = ALLO_CTL_NO
                        rowSOTALLO2S.Item("CUST_CODE") = CUST_CODE
                        dst.Tables("SOTALLO2S").Rows.Add(rowSOTALLO2S)
                    End If
                    rowSOTALLO2S.Item("QTY_ALLO") = row.Item("QTY_ALLO")
                Next

                ASCMAIN1.sql = "Delete from " & SOTALLO2S & " where ALLO_CTL_NO = '" & ALLO_CTL_NO & "'"
                ASCDATA1.ExecuteSQL()
                ASCMAIN1.sql = "Delete from " & SOTALLO3S & " where ALLO_CTL_NO = '" & ALLO_CTL_NO & "'"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Select SOTALLO2.*, ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                    & ", 0 ORDR_QTY, 0 ORDR_QTY_OPEN, 0 ORDR_QTY_PICK, 0 ORDR_QTY_SHIP, 0 ORDR_QTY_CANC" & vbCrLf _
                    & ", 0 LY_QTY_SELL_IN1, 0 LY_QTY_SELL_IN2" & vbCrLf _
                    & ", 0 LY_QTY_SELL_THRU1, 0 LY_QTY_SELL_THRU2" & vbCrLf _
                    & " from " & SOTALLO1S & " SOTALLO1,SOTALLO2,ARTCUST1" & vbCrLf _
                    & " where SOTALLO2.ALLO_CTL_NO = SOTALLO1.ALLO_CTL_NO" & vbCrLf _
                    & "   and ARTCUST1.CUST_CODE = SOTALLO2.CUST_CODE" & vbCrLf _
                    & "   and SOTALLO1.ALLO_CTL_NO = '" & ALLO_CTL_NO & "'"
                ASCDATA1.ExecuteSQL("Insert into " & SOTALLO2S & " " & ASCMAIN1.sql)

                ASCMAIN1.sql = "Select SOTALLO3.*, ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                    & ", 0 ORDR_QTY, 0 ORDR_QTY_OPEN, 0 ORDR_QTY_PICK, 0 ORDR_QTY_SHIP, 0 ORDR_QTY_CANC" & vbCrLf _
                    & ", 0 LY_QTY_SELL_IN1, 0 LY_QTY_SELL_IN2" & vbCrLf _
                    & ", 0 LY_QTY_SELL_THRU1, 0 LY_QTY_SELL_THRU2" & vbCrLf _
                    & " from " & SOTALLO1S & " SOTALLO1,SOTALLO3,ARTCUST1,ARTCUST2" & vbCrLf _
                    & " where SOTALLO3.ALLO_CTL_NO = SOTALLO1.ALLO_CTL_NO" & vbCrLf _
                    & "   and ARTCUST1.CUST_CODE = SOTALLO3.CUST_CODE" & vbCrLf _
                    & "   and ARTCUST2.CUST_CODE = SOTALLO3.CUST_CODE" & vbCrLf _
                    & "   and ARTCUST2.CUST_STORE_NO = SOTALLO3.CUST_STORE_NO" & vbCrLf _
                    & "   and SOTALLO1.ALLO_CTL_NO = '" & ALLO_CTL_NO & "'"
                ASCDATA1.ExecuteSQL("Insert into " & SOTALLO3S & " " & ASCMAIN1.sql)

                Get_Sales_STATS(True, ALLO_CTL_NO)

                Click_Node(tvwDQ.ActiveNode)

                tabSOTALLOX.Tabs("Status by Item").Tag = ""
                tabSOTALLOX.Tabs("Status by Customer").Tag = ""
            End If

        End If

    End Sub

    Private Sub Clear_Record()

        MyBase.EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTALLO1", "SOTALLO2", "SOTALLOB", "SOTALLOC", "SOTALLOD", "SOTALLO3", "SOTALLO4", "SOTALLOS", "SOTORDR2", "SOTORDRA", "SOTORDRU", "SOTORDRD", "SOTORDRN", "SOTALLH1", "SOTALLH2", "SOTALLH3", "SOTALLOT"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        ARTCUSF2 = "ARTCUST2"
        sqlwhere_F = ""

        If grdSOTALLOS.DisplayLayout.Bands(0).Columns.Exists("TOTAL_EVENT_QTY") Then
            grdSOTALLOS.DisplayLayout.Bands(0).Columns("TOTAL_EVENT_QTY").Hidden = True

        End If

        'ALLO_CTL_NOs_to_Delete.Clear()
        'ALLO_CTL_NO_to_copy = ""
        'CUST_CODE_to_copy = ""
        'ALLO_CTL_NO_new.Clear()
        'ITEM_CODE_new.Clear()
        'iCol.Clear()
        'ReDim ALLO_CTL_NOi(maxAllocations)
        'ALLO_CTL_NOs = ""
        cmbEvent.Text = ""


        Load_SOTALLOX()
        MyBase.EnforceConstraints(True)

        For i As Integer = 1 To maxAllocations
            With grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_" & Format(i, "00"))
                .Header.Caption = ""
                .Hidden = True
                .Tag = ""
            End With
            With grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_NOTES_" & Format(i, "00"))
                .Header.Caption = ""
                .Hidden = True
                .Tag = ""
            End With
        Next

        For i As Integer = 1 To maxAllocations
            With grdSOTALLOS.DisplayLayout.Bands(0).Columns("ALLO_" & Format(i, "00"))
                .Header.Caption = ""
                .Hidden = True
                .Tag = ""
            End With
            With grdSOTALLOS.DisplayLayout.Bands(0).Columns("ALLO_NOTES_" & Format(i, "00"))
                .Header.Caption = ""
                .Hidden = True
                .Tag = ""
            End With
        Next

        '  txtCOLLECTION_CODE.Clear()
        Absx1.txtFor("COLLECTION_CODE").Text = COLLECTION_CODE

        grdSOTORDRU.Visible = False

        Clear_All_Filters(grdICTITEM1)
        Clear_All_Filters(grdSOTALLOX)
    End Sub

    Private Sub Load_Record()

        MyBase.EnforceConstraints(False)

        Absx1.txtFor("COLLECTION_CODE").Text = COLLECTION_CODE

        Dim rowICTCOLL1 As DataRow = LookUp("ICTCOLL1", COLLECTION_CODE)
        HC_CODE_lead_item = rowICTCOLL1.Item("HC_CODE") & ""

        ALLO_CTL_NOs_to_Delete.Clear()

        ALLO_CTL_NO_to_copy = ""
        CUST_CODE_to_copy = ""

        ALLO_CTL_NO_new.Clear()
        ITEM_CODE_new.Clear()
        iCol.Clear()
        ReDim ALLO_CTL_NOi(maxAllocations)

        ALLO_CTL_NOs = ""

        If EntryMode = "N" Then
            If tabSOTALLOX.SelectedTab.Key = "Allocations" Then ' Cloning existing Allocations to New Allocations from Allocations Tab
                For Each row As DataRow In dst.Tables("SOTALLOX").Select("SELECTED = '1'")
                    Dim ALLO_CTL_NO_old As String = row.Item("ALLO_CTL_NO")

                    Dim rowSOTALLO1 As DataRow = Add_Item(row.Item("ITEM_CODE"))
                    Dim ALLO_CTL_NO_new As String = rowSOTALLO1.Item("ALLO_CTL_NO")

                    Fill_Records("SOTALLO2", ALLO_CTL_NO_old, False)
                    For Each rowSOTALLO2 As DataRow In dst.Tables("SOTALLO2").Select("ALLO_CTL_NO = '" & ALLO_CTL_NO_old & "'")
                        rowSOTALLO2.Item("ALLO_CTL_NO") = ALLO_CTL_NO_new
                        rowSOTALLO2.AcceptChanges()
                        rowSOTALLO2.SetAdded()
                    Next

                    Fill_Records("SOTALLO3", ALLO_CTL_NO_old, False)
                    For Each rowSOTALLO3 As DataRow In dst.Tables("SOTALLO3").Select("ALLO_CTL_NO = '" & ALLO_CTL_NO_old & "'")
                        rowSOTALLO3.Item("ALLO_CTL_NO") = ALLO_CTL_NO_new
                        rowSOTALLO3.AcceptChanges()
                        rowSOTALLO3.SetAdded()
                    Next

                    Fill_Records("SOTALLO4", ALLO_CTL_NO_old, False)
                    For Each rowSOTALLO4 As DataRow In dst.Tables("SOTALLO4").Select("ALLO_CTL_NO = '" & ALLO_CTL_NO_old & "'")
                        rowSOTALLO4.Item("ALLO_CTL_NO") = ALLO_CTL_NO_new
                        rowSOTALLO4.AcceptChanges()
                        rowSOTALLO4.SetAdded()
                    Next
                Next
            Else ' New Allocations for Items Selected in Items Tab
                For Each row As DataRow In dst.Tables("ICTITEM1").Select("SELECTED = '1'")
                    Add_Item(row.Item("ITEM_CODE"))
                Next
            End If

        Else
            For Each row As DataRow In dst.Tables("SOTALLOX").Select("SELECTED = '1'")
                ALLO_CTL_NOs &= ", '" & row.Item("ALLO_CTL_NO") & "'"
            Next
            If ALLO_CTL_NOs <> "" Then
                Add_Allocations(ALLO_CTL_NOs, True)
            End If
        End If

        iColumn = 0
        ARTCUSF2 = "ARTCUST2"
        sqlwhere_F = ""
        For Each rowSOTALLO1 As DataRow In dst.Tables("SOTALLO1").Select

            If iColumn = 0 Then
                If rowSOTALLO1.Item("DATE_START") & "" <> "" Then
                    ' THERE SHOULD ONLY BE 1 RECORD IN THIS TABLE IF THERE ARE ANY FUTURE REALIGNMENTS QUEUED
                    Dim rowARTCUSF1 As DataRow = LookUp("ARTCUSF1", "000000")
                    If rowARTCUSF1 IsNot Nothing Then
                        If Format(rowSOTALLO1.Item("DATE_START"), "yyyyMMdd") >= Format(rowARTCUSF1.Item("REALIGN_DATE"), "yyyyMMdd") Then
                            ARTCUSF2 = "ARTCUSF2"
                            sqlwhere_F = " and ARTCUSF2.REALIGN_NO (+) = '000000'" & vbCrLf
                        End If
                    End If
                End If
            End If

            Add_Allocation_to_Grid(rowSOTALLO1)
        Next

        If ARTCUSF2 = "ARTCUSF2" Then
            lblFutureAE.Visible = True
        Else
            lblFutureAE.Visible = False
        End If

        MyBase.EnforceConstraints(True)

        Sort_grdColumns(grdSOTALLOC, "CUST_CODE")
        Setup_SOTALLOC()
        Update_Totals()

        chkSingleItemMode.Checked = False
        chkSI.Checked = False
        chkST.Checked = False
        For Each row As DataRow In dst.Tables("SOTALLO1").Select("")
            If row.Item("ITEM_SNU_CODE") & "" = "N" Then chkSI.Checked = True
            If row.Item("ITEM_SNU_CODE") & "" = "S" Then chkST.Checked = True
        Next

        ASCMAIN1.Progress(String.Empty, String.Empty)
    End Sub

    Sub Print_Record()
        Create_Report()
    End Sub
    Function Create_Report() As String

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Report")

        Dim REPORT_NAME As String = "SORALLO1"
        Dim RPT As String = REPORT_NAME

        If Not REPORTS.ContainsKey(REPORT_NAME) Then
            REPORTS.Add(REPORT_NAME, Load_rptClass(REPORT_NAME))
            REPORTS(REPORT_NAME).Prepare_dst(False, "")
        End If


        'REPORTS(REPORT_NAME).Fill_Records_RPT("")

        dst.Tables("SOTALLOZ").Rows.Clear()
        dst.Tables("ICTSTAT2").Rows.Clear()
        dst.Tables("ARTCUST1").Rows.Clear()
        dst.Tables("ICTITEM1").Rows.Clear()

        REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("ARTCUST1").Rows.Clear()
        REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("SOTALLO2").Rows.Clear()
        REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("SOTALLO1").Rows.Clear()
        REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("SOTALLOZ").Rows.Clear()
        REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("ICTSTAT2").Rows.Clear()
        REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("ICTITEM1").Rows.Clear()

        Dim ITEM_CODEs As New List(Of String)
        Dim CUST_CODEs As New List(Of String)

        For Each row As DataRow In dst.Tables("SOTALLO1").Select("")
            With REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("SOTALLO1")
                Dim rowR As DataRow = .NewRow
                For Each COLUMN_NAME As String In New String() _
                    {"ALLO_CTL_NO", "ITEM_CODE", "DATE_START", "DATE_END", "INIT_OPER", "INIT_DATE", "LAST_OPER", "LAST_DATE", "ALLOW_OVER",
                     "ITEM_DESC", "COLLECTION_CODE", "BRAND_CODE", "ITEM_BASIC_PROMO", "ITEM_SNU_CODE", "QTY_ALLO_PLAN", "QTY_ALLO_TOTAL", "ITEM_DATE_TO_SHIP"}
                    If COLUMN_NAME = "BRAND_CODE" Then
                    Else
                        rowR.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                    End If
                Next
                .Rows.Add(rowR)
            End With

            Dim ALLO_CTL_NO As String = row.Item("ALLO_CTL_NO")

            Fill_Records("SOTALLOZ", ALLO_CTL_NO, False)


            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            If Not ITEM_CODEs.Contains(ITEM_CODE) Then
                Fill_Records("ICTSTAT2", ITEM_CODE, False)


                ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE" & sql_ICTITEM1 _
                    & " from ICTITEM1 where ITEM_CODE = '" & ITEM_CODE & "'"
                Fill_Records("ICTITEM1", "", False, ASCMAIN1.sql)
                Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)

                Dim rowR As DataRow = REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("ICTITEM1").NewRow
                For Each DC As DataColumn In dst.Tables("ICTITEM1").Columns
                    Dim COLUMN_NAME As String = DC.ColumnName
                    If REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("ICTITEM1").Columns.Contains(COLUMN_NAME) Then
                        rowR.Item(COLUMN_NAME) = rowICTITEM1.Item(COLUMN_NAME)
                    End If
                Next
                REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("ICTITEM1").Rows.Add(rowR)

                Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
                Dim imgba() As Byte = Nothing
                Dim IMAGE_FILENAME As String = FOLDER_NAME & "\" & ITEM_CODE & ".JPG"
                If My.Computer.FileSystem.FileExists(IMAGE_FILENAME) Then
                    rowR.Item("ITEM_IMAGE") = ASCMAIN1.GetImageData(IMAGE_FILENAME)
                Else
                    IMAGE_FILENAME = FOLDER_NAME & "\" & ITEM_CODE & ".PNG"
                    If My.Computer.FileSystem.FileExists(IMAGE_FILENAME) Then
                        rowR.Item("ITEM_IMAGE") = ASCMAIN1.GetImageData(IMAGE_FILENAME)
                    End If
                End If
            End If
        Next

        Fill_SOTALLO2()
        For Each row As DataRow In dst.Tables("SOTALLO2").Select("")
            With REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("SOTALLO2")
                Dim rowR As DataRow = .NewRow
                For Each COLUMN_NAME As String In New String() {"ALLO_CTL_NO", "CUST_CODE", "QTY_ALLO"}
                    rowR.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                Next
                .Rows.Add(rowR)
            End With
            Dim ALLO_CTL_NO As String = row.Item("ALLO_CTL_NO")
            Dim CUST_CODE As String = row.Item("CUST_CODE")

            If dst.Tables("SOTALLOZ").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE}) Is Nothing Then
                dst.Tables("SOTALLOZ").Rows.Add(New String() {ALLO_CTL_NO, CUST_CODE})
            End If

            If Not CUST_CODEs.Contains(CUST_CODE) Then
                Fill_Records("ARTCUST1", CUST_CODE, False)
            End If
        Next

        For Each row As DataRow In dst.Tables("ARTCUST1").Select("")
            With REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("ARTCUST1")
                Dim rowR As DataRow = .NewRow
                For Each COLUMN_NAME As String In New String() {"CUST_CODE", "CUST_NAME"}
                    rowR.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                Next
                .Rows.Add(rowR)
            End With
        Next

        For Each row As DataRow In dst.Tables("SOTALLOZ").Select("")
            With REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("SOTALLOZ")
                Dim rowR As DataRow = .NewRow
                For Each DC As DataColumn In dst.Tables("SOTALLOZ").Columns
                    Dim COLUMN_NAME As String = DC.ColumnName
                    rowR.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                Next
                .Rows.Add(rowR)
            End With
        Next

        For Each row As DataRow In dst.Tables("ICTSTAT2").Select("")
            With REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("ICTSTAT2")
                Dim rowR As DataRow = .NewRow
                For Each DC As DataColumn In dst.Tables("ICTSTAT2").Columns
                    Dim COLUMN_NAME As String = DC.ColumnName
                    rowR.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                Next
                .Rows.Add(rowR)
            End With
        Next


        With REPORTS(REPORT_NAME).clsASCBASE1
            .Print_Report_Begin()
            '.CR_params.Add("SUBT", "")
            '.CR_params.Add("CONS_INV", "")
            'Dim REPORT_NO As String = .Generate_Report(RPT, "", "", False, False, "", "PDF", , False)

            Dim SUBT As String = "Allocations by Item/Customer (Screen Report)"
            .CR_params.Add("SUBT", SUBT) ' "")
            .CR_params.Add("PAGE_EJECT", "0")
            .CR_params.Add("EXC_ONLY", "0")
            .CR_params.Add("SUMMARY", "0")
            .Generate_Report(RPT, Me.Text, SUBT)
            .Print_Report_End()
            '.Print_Report_End(,  True)

            'CR_params.Add("SUBT", "Customer/Item")
            'CR_params.Add("PAGE_EJECT", IIf(Absx1.chkFor("CHKPAGE_EJECT").Checked, "1", "0"))
            'CR_params.Add("EXC_ONLY", IIf(Absx1.chkFor("CHKEXC_ONLY").Checked, "1", "0"))
            'RPT = "SORALLO2"
            'Generate_Report(RPT, , SUBT)

        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Return ""
    End Function

    Private Sub Update_Record()

        ASCMAIN1.sql = "Truncate Table " & SOTORDRX
        ASCDATA1.ExecuteSQL()

        dst.Tables("SOTORDRN").Rows.Clear()
        dst.Tables("SOTORDRD").Rows.Clear()

        Try
            MyBase.BeginTrans()

            dst.Tables("ICTITEM1").Rows.Clear()

            Dim ITEM_NOT_ALLOCATEDs As New List(Of String)

            For Each rowSOTALLO1 As DataRow In dst.Tables("SOTALLO1").Select("", "ITEM_CODE,ALLO_CTL_NO")
                Dim ALLO_CTL_NO As String = rowSOTALLO1.Item("ALLO_CTL_NO")
                Dim ITEM_CODE As String = rowSOTALLO1.Item("ITEM_CODE")
                Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
                If rowICTITEM1 Is Nothing Then

                    ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE" & sql_ICTITEM1 _
                    & $" from ICTITEM1 where ITEM_CODE = '{ITEM_CODE}'"

                    Fill_Records("ICTITEM1",, False, ASCMAIN1.sql)
                    rowICTITEM1 = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)

                    If rowICTITEM1.Item("ITEM_NOT_ALLOCATED") & "" = "1" Then
                        ITEM_NOT_ALLOCATEDs.Add(ITEM_CODE)
                        rowICTITEM1.Item("ITEM_NOT_ALLOCATED") = "0"
                    End If
                End If
            Next

            If ITEM_NOT_ALLOCATEDs.Count > 0 Then
                Update_Record_TDA("ICTITEM1")
            End If

            Dim sqlA As String = "ALLO_CTL_NO in (" & Mid(ALLO_CTL_NOs, 2) & ")"
            Update_Record_TDA("SOTALLO1", sqlA)
            ASCDATA1.ExecuteSQL("UPDATE SOTALLO1 SET LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "' WHERE " & sqlA)

            Fill_SOTALLO2()
            Update_Record_TDA("SOTALLO2", sqlA)

            Fill_SOTALLO3()
            Update_Record_TDA("SOTALLO3", sqlA)

            Fill_SOTALLO4()
            Update_Record_TDA("SOTALLO4", sqlA)

            ASCMAIN1.sql = "Insert into " & SOTORDRX & vbCrLf _
                & "Select ORDR_NO, ALLO_CTL_NO, ITEM_CODE, CUST_CODE" & vbCrLf _
                & ", NVL(ORDR_QTY_OPEN,0)+NVL(ORDR_QTY_PICK,0)+NVL(ORDR_QTY_SHIP,0) QTY" & vbCrLf _
                & " from SOTORDR2 where " & sqlA
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update SOTORDR2 Set ALLO_CTL_NO = Null where " & sqlA
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C0 is" & vbCrLf _
                & "  Select ALLO_CTL_NO, DATE_START, DATE_END, ITEM_CODE from SOTALLO1 where " & sqlA & ";" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R0 in C0 Loop" & vbCrLf _
                & "   Begin" & vbCrLf _
                & "    Declare Cursor C1 is" & vbCrLf _
                & "     Select Distinct SOTORDR1.ORDR_GROUP_NO from SOTORDR2,SOTORDR1" & vbCrLf _
                & "       where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO and SOTORDR2.ITEM_CODE = R0.ITEM_CODE" & vbCrLf _
                & "         and R0.DATE_END >= NVL(SOTORDR1.ORDR_DATE_SHIPPED,NVL(SOTORDR1.ORDR_ALLO_DATE,SOTORDR1.ORDR_SHIP_DATE))" & vbCrLf _
                & "         and R0.DATE_START <= NVL(SOTORDR1.ORDR_DATE_SHIPPED,NVL(SOTORDR1.ORDR_ALLO_DATE,SOTORDR1.ORDR_SHIP_DATE))" & vbCrLf _
                & "         and SOTORDR2.ORDR_STATUS IN ('O','P','F') and SOTORDR2.ALLO_CTL_NO IS NULL" & vbCrLf _
                & "      union " & vbCrLf _
                & "     Select Distinct SOTORDR1.ORDR_GROUP_NO from SOTORDR2,SOTORDR1" & vbCrLf _
                & "       where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO and SOTORDR2.ITEM_CODE = R0.ITEM_CODE" & vbCrLf _
                & "         and SOTORDR2.ALLO_CTL_NO_REL = R0.ALLO_CTL_NO" & vbCrLf _
                & "         and SOTORDR2.ORDR_STATUS IN ('O','P','F') and SOTORDR2.ALLO_CTL_NO IS NULL;" & vbCrLf _
                & "    Begin" & vbCrLf _
                & "     For R1 in C1 Loop" & vbCrLf _
                & "      SOPORDR0_GA(R1.ORDR_GROUP_NO);" & vbCrLf _
                & "     End Loop;" & vbCrLf _
                & "    End;" & vbCrLf _
                & "   End;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

            For Each row As DataRow In dst.Tables("SOTALLO1").Select("")
                Dim ALLO_CTL_NO As String = row.Item("ALLO_CTL_NO")
                Fill_Records("SOTORDRD", ALLO_CTL_NO, False)
                Fill_Records("SOTORDRN", ALLO_CTL_NO, False)
            Next

            ASCMAIN1.sql = "Insert into SOTALLH1 Select '" & XNO & "' XNO, SOTALLO1.* from SOTALLO1 where " & sqlA
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Insert into SOTALLH2 Select '" & XNO & "' XNO, SOTALLO2.* from SOTALLO2 where " & sqlA
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Insert into SOTALLH3 Select '" & XNO & "' XNO, SOTALLO3.* from SOTALLO3 where " & sqlA
            ASCDATA1.ExecuteSQL()

            If ALLO_CTL_NOs_to_Delete.Count <> 0 Then
                Dim sqlw As String = " where ALLO_CTL_NO in ('" & Join(ALLO_CTL_NOs_to_Delete.ToArray, "','") & "')"
                ASCMAIN1.sql = "Delete from SOTALLO1" & sqlw
                ASCDATA1.ExecuteSQL()
                ASCMAIN1.sql = "Delete from SOTALLO2" & sqlw
                ASCDATA1.ExecuteSQL()
                ASCMAIN1.sql = "Delete from SOTALLO3" & sqlw
                ASCDATA1.ExecuteSQL()
                ASCMAIN1.sql = "Update SOTORDR2 Set ALLO_CTL_NO = Null" & sqlw
                ASCDATA1.ExecuteSQL()
            End If

            MyBase.CommitTrans("Update Complete")

        Catch ex As Exception
            MyBase.Rollback(ex.Message)
        End Try

    End Sub
#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTALLO1, "BBBBBBB",
                        "Item Status Inquiry",
                        "Copy Allocation Dates & Qtys",
                        "Paste Allocation Dates to Selected Rows",
                        "Paste Allocation Qtys to Selected Rows",
                        "Remove from List (No Update)",
                        "Clear Allocation Qtys",
                        "Delete Allocation")
        Load_Popup_Menu(grdSOTALLOX, "SBBBBBBB", "Show Filter", "Select All", "De-Select All", "Select All in Group", "Item Status Inquiry", "Select Selected", "De-Select Selected", "Select All PROD_CODE")
        Load_Popup_Menu(grdICTITEM1, "SBBBBBB", "Show Filter", "Select All", "De-Select All", "Item Status Inquiry", "Select Selected", "De-Select Selected", "Select All PROD_CODE")

        Load_Popup_Menu(grdSOTALLOC, "SSBBBB", "Show Filter", "Show Pins", "Add Customers", "Copy Customer Qtys", "Paste Customer Qtys to Selected Rows", "Load from XLS")
        Load_Popup_Menu(grdSOTALLOS, "SSBBBBBB", "Show Filter", "Show Pins", "Add Stores", "Copy Store Qtys", "Paste Store Qtys to Selected Rows", "Paste Qtys from Excel", "Clear Quantities", "Copy Event Qtys")

        Load_Popup_Menu(grdSOTALLO1S, "SBB", "Show Filter", "Expand All", "Collapse All")
        Load_Popup_Menu(grdSOTALLOCS, "SSBB", "Show Filter", "Show All Levels", "Expand All", "Collapse All")

        Load_Popup_Menu(grdSOTORDRD, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTORDR2, "SS", "Show Filter", "Show GroupBox")

        Load_Popup_Menu(grdSOTALLH1, "S", "Show Filter")
        Load_Popup_Menu(grdSOTALLOT, "B", "Remove Event")
        ' Load_Popup_Menu(grdSOTORDRA, "SB", "Show Filter", "Move Allocation")
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
            Case "grdICTITEM1"
                tlb_btn = DirectCast(tlb_pop.Tools("Select All"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not ScreenMode And Not InquiryMode
                tlb_btn = DirectCast(tlb_pop.Tools("De-Select All"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not ScreenMode And Not InquiryMode
                tlb_btn = DirectCast(tlb_pop.Tools("Select Selected"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not ScreenMode And (grdICTITEM1.Selected.Rows.Count <> 0) And Not InquiryMode
                tlb_btn = DirectCast(tlb_pop.Tools("De-Select Selected"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not ScreenMode And (grdICTITEM1.Selected.Rows.Count <> 0) And Not InquiryMode

                If InquiryMode Then
                    tlb_btn = DirectCast(tlb_pop.Tools("Select All PROD_CODE"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = False
                End If

            Case "grdSOTALLOX"
                tlb_btn = DirectCast(tlb_pop.Tools("Select All"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not ScreenMode And Not InquiryMode
                tlb_btn = DirectCast(tlb_pop.Tools("De-Select All"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not ScreenMode And Not InquiryMode
                tlb_btn = DirectCast(tlb_pop.Tools("Select All in Group"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not ScreenMode And Not InquiryMode
                tlb_btn = DirectCast(tlb_pop.Tools("Select Selected"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not ScreenMode And (grdSOTALLOX.Selected.Rows.Count <> 0) And Not InquiryMode
                tlb_btn = DirectCast(tlb_pop.Tools("De-Select Selected"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not ScreenMode And (grdSOTALLOX.Selected.Rows.Count <> 0) And Not InquiryMode


                If InquiryMode Then
                    tlb_btn = DirectCast(tlb_pop.Tools("Select All PROD_CODE"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = False
                End If

            Case "grdSOTALLO1"

                tlb_btn = DirectCast(tlb_pop.Tools("Paste Allocation Qtys to Selected Rows"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (ALLO_CTL_NO_to_copy <> "") And grdSOTALLO1.Selected.Rows.Count > 0 And Not InquiryMode
                tlb_btn.SharedProps.Caption = "Paste Allocation " & ALLO_CTL_NO_to_copy & " Qtys to Selected Rows"

                tlb_btn = DirectCast(tlb_pop.Tools("Paste Allocation Dates to Selected Rows"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (ALLO_CTL_NO_to_copy <> "") And grdSOTALLO1.Selected.Rows.Count > 0 And Not InquiryMode
                tlb_btn.SharedProps.Caption = "Paste Allocation " & ALLO_CTL_NO_to_copy & " Dates to Selected Rows"


                If InquiryMode Then
                    tlb_btn = DirectCast(tlb_pop.Tools("Copy Allocation Dates & Qtys"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = False
                    tlb_btn = DirectCast(tlb_pop.Tools("Remove from List (No Update)"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = False
                    tlb_btn = DirectCast(tlb_pop.Tools("Clear Allocation Qtys"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = False
                    tlb_btn = DirectCast(tlb_pop.Tools("Delete Allocation"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = False
                End If


            Case "grdSOTALLOC"

                tlb_btn = DirectCast(tlb_pop.Tools("Paste Customer Qtys to Selected Rows"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (CUST_CODE_to_copy <> "") And grdSOTALLOC.Selected.Rows.Count > 0 And Not InquiryMode
                tlb_btn.SharedProps.Caption = "Paste Customer " & CUST_CODE_to_copy & " Qtys to Selected Rows"

                tlb_btn = DirectCast(tlb_pop.Tools("Load from XLS"), UltraWinToolbars.ButtonTool)
                'tlb_btn.SharedProps.Visible = Not (ASCMAIN1.CLIENT = "INT")

                'tlb_btn = DirectCast(tlb_pop.Tools("Prorate"), UltraWinToolbars.ButtonTool)

                'Me.Cursor = New Cursor(Cursor.Current.Handle)
                ''Cursor.Position = New System.Drawing.Point(Cursor.Position.X, Cursor.Position.Y)
                'Dim elementPoint As UIElement = grd.DisplayLayout.UIElement.ElementFromPoint(New System.Drawing.Point(Cursor.Position.X, Cursor.Position.Y))
                'Dim columnHeaderUIElement As UltraWinGrid.ColumnHeader = elementPoint.GetContext(GetType(Infragistics.Win.UltraWinGrid.ColumnHeader))
                'Dim cellUIElement As UltraWinGrid.RowCellAreaUIElement = elementPoint.GetContext(GetType(Infragistics.Win.UltraWinGrid.RowCellAreaUIElement))

                'tlb_btn.Tag = ""
                'If grd.ActiveCell IsNot Nothing Then
                '    If grd.ActiveCell.Column.Key = "SI_PLAN" Or grd.ActiveCell.Column.Key = "ST_PLAN" Or grd.ActiveCell.Column.Key = "SI_HIST" Or grd.ActiveCell.Column.Key = "ST_HIST" Then
                '        tlb_btn.Tag = grd.ActiveCell.Column.Key
                '    End If
                'End If
                'If Not chkSingleItemMode.Checked Or grdSOTALLO1.ActiveRow Is Nothing Or tlb_btn.Tag = "" Then
                '    tlb_btn.SharedProps.Visible = False
                'Else
                '    ' Dim ITEM_CODE As String = grdSOTALLO1.ActiveRow.Cells("ITEM_CODE").Value
                '    ' Dim QTY_ALLO_PLAN As Int64 = Val(grdSOTALLO1.ActiveRow.Cells("QTY_ALLO_PLAN").Value & "")
                '    tlb_btn.SharedProps.Visible = True
                '    tlb_btn.SharedProps.Caption = "Set Allocation to " & tlb_btn.Tag
                'End If

                If InquiryMode Then
                    tlb_btn = DirectCast(tlb_pop.Tools("Add Customers"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = False
                    tlb_btn = DirectCast(tlb_pop.Tools("Copy Customer Qtys"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = False
                    tlb_btn = DirectCast(tlb_pop.Tools("Load from XLS"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = False
                End If


            Case "grdSOTALLOS"

                tlb_btn = DirectCast(tlb_pop.Tools("Paste Store Qtys to Selected Rows"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (CUST_STORE_NO_to_copy <> "") And grdSOTALLOC.Selected.Rows.Count > 0 And Not InquiryMode
                tlb_btn.SharedProps.Caption = "Paste Store " & CUST_STORE_NO_to_copy & " Qtys to Selected Rows"

                tlb_btn = DirectCast(tlb_pop.Tools("Clear Quantities"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = grdSOTALLOS.ActiveCell IsNot Nothing AndAlso Not InquiryMode
                'tlb_btn = DirectCast(tlb_pop.Tools("Prorate"), UltraWinToolbars.ButtonTool)
                'If Not chkSingleItemMode.Checked Or grdSOTALLO1.ActiveRow Is Nothing Then
                '    tlb_btn.SharedProps.Visible = False
                'Else
                '    Dim ITEM_CODE As String = grdSOTALLO1.ActiveRow.Cells("ITEM_CODE").Value
                '    Dim ALLO_CTL_NO As String = grdSOTALLO1.ActiveRow.Cells("ALLO_CTL_NO").Value
                '    Dim iCtr As Integer = iCol(ALLO_CTL_NO)
                '    Dim Z As String = "ALLO_" & Format(iCtr, "00")

                '    Dim QTY_ALLO_PLAN As Int64 = Val(grdSOTALLOC.ActiveRow.Cells(Z).Value & "")
                '    tlb_btn.SharedProps.Visible = chkSingleItemMode.Checked And QTY_ALLO_PLAN <> 0
                '    tlb_btn.SharedProps.Caption = "Prorate Allocation Qty of " & Format(QTY_ALLO_PLAN, "#,##0") & " for Item " & ITEM_CODE & " to Stores of " & grdSOTALLOC.ActiveRow.Cells("CUST_CODE").Value & " using LY Sell-In"
                'End If

                If InquiryMode Then
                    tlb_btn = DirectCast(tlb_pop.Tools("Add Stores"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = False
                    tlb_btn = DirectCast(tlb_pop.Tools("Copy Store Qtys"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = False
                End If

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            '   e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
                Case "grdSOTALLOX", "grdICTITEM1"
                    tlb_pop = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
                    tlb_btn = DirectCast(tlb_pop.Tools("Select All PROD_CODE"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = False
                    If grd.ActiveRow IsNot Nothing And grd.ActiveRow.IsDataRow Then
                        Dim PROD_CODE As String = grd.ActiveRow.Cells("PROD_CODE").Value & ""
                        If PROD_CODE <> "" Then
                            tlb_btn.SharedProps.Caption = "Select All " & PROD_CODE
                            tlb_btn.Tag = PROD_CODE
                            tlb_btn.SharedProps.Visible = True And Not InquiryMode
                        End If
                    End If
                    '    Dim ORDR_TYPE As String = ""
                    '    If grd.ActiveRow IsNot Nothing And grd.ActiveRow.IsDataRow Then
                    '        ORDR_TYPE = grd.ActiveRow.Cells("ORDR_TYPE").Value & ""
                    '    End If
                    '    tlb_btn.SharedProps.Visible = (ORDR_TYPE = "O")
                    '    tlb_btn = DirectCast(tlb_pop.Tools("Customer Order Summary"), UltraWinToolbars.ButtonTool)
                    '    tlb_btn.SharedProps.Visible = ScreenMode


            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If e.Tool.Key <> "Show All Levels" Then grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    grow.Cells("SELECTED").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                    grow.Update()
                Next

            Case "Add Customers"
                Add_Codes(grdSOTALLOC, "ARTCUST1", "CUST_CODE", "Customers")

            Case "Add Stores"
                Dim CUST_CODE As String = grdSOTALLOC.ActiveRow.Cells("CUST_CODE").Value & ""
                Me.Add_CUST_STOREs(CUST_CODE)

            Case "Load from XLS"
                Load_from_XLS()

            Case "Paste Allocation Dates to Selected Rows"
                If ALLO_CTL_NO_to_copy = "" Or grdSOTALLO1.Selected.Rows.Count = 0 Then
                Else
                    Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO_to_copy)
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTALLO1.Selected.Rows
                        grow.Cells("DATE_START").Value = rowSOTALLO1.Item("DATE_START")
                        grow.Cells("DATE_END").Value = rowSOTALLO1.Item("DATE_END")
                        grow.Update()
                    Next
                End If

            Case "Paste Allocation Qtys to Selected Rows"
                If ALLO_CTL_NO_to_copy = "" Or grdSOTALLO1.Selected.Rows.Count = 0 Then
                Else
                    Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO_to_copy)
                    Dim ALLO_CTL_NOs_to_copy_to As New List(Of String)
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTALLO1.Selected.Rows
                        If grow.Cells("ALLO_CTL_NO").Value & "" <> ALLO_CTL_NO_to_copy Then
                            ALLO_CTL_NOs_to_copy_to.Add(grow.Cells("ALLO_CTL_NO").Value)
                        End If
                    Next

                    If ALLO_CTL_NOs_to_copy_to.Count <> 0 Then
                        Dim i_to_copy As Integer = iCol(ALLO_CTL_NO_to_copy)
                        For Each rowSOTALLOC As DataRow In dst.Tables("SOTALLOC").Select("")
                            For Each ALLO_CTL_NO As String In ALLO_CTL_NOs_to_copy_to
                                Dim i As Integer = iCol(ALLO_CTL_NO)
                                rowSOTALLOC.Item("ALLO_" & Format(i, "00")) = rowSOTALLOC.Item("ALLO_" & Format(i_to_copy, "00"))
                            Next
                        Next
                    End If
                End If

            Case "Paste Customer Qtys to Selected Rows"
                If CUST_CODE_to_copy = "" Or grdSOTALLOC.Selected.Rows.Count = 0 Then
                Else
                    Dim rowSOTALLOC As DataRow = dst.Tables("SOTALLOC").Rows.Find(CUST_CODE_to_copy)
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTALLOC.Selected.Rows
                        For I As Integer = 1 To iColumn
                            grow.Cells("ALLO_" & Format(I, "00")).Value = rowSOTALLOC.Item("ALLO_" & Format(I, "00"))
                        Next
                        grow.Update()
                    Next
                End If

            Case "Paste Store Qtys to Selected Rows"
                If CUST_STORE_NO_to_copy = "" Or grdSOTALLOS.Selected.Rows.Count = 0 Then
                Else
                    Dim rowSOTALLOS As DataRow = dst.Tables("SOTALLOS").Rows.Find(CUST_CODE_to_copy)
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTALLOS.Selected.Rows
                        For I As Integer = 1 To iColumn
                            grow.Cells("ALLO_" & Format(I, "00")).Value = rowSOTALLOS.Item("ALLO_" & Format(I, "00"))
                        Next
                        grow.Update()
                    Next
                End If


            Case "Show All Levels"
                'tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
                'grd.DisplayLayout.Bands(0).ColHeadersVisible = tlb_sbt.Checked

                'Click_Node(tvwDQ.ActiveNode)
                'Exit Sub

            Case "Expand All"

                grd.Rows.ExpandAll(True)

            Case "Collapse All"

                grd.Rows.CollapseAll(True)

            Case "Select All PROD_CODE"
                tlb_btn = DirectCast(e.Tool, UltraWinToolbars.ButtonTool)
                Dim PROD_CODE As String = tlb_btn.Tag
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.Cells("PROD_CODE").Value = PROD_CODE Then
                        grow.Cells("SELECTED").Value = "1"
                        grow.Update()
                    End If
                Next

        End Select

        If grd Is Nothing OrElse (grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow) Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Select All in Group"
                Dim ALLO_GROUP_CODE As String = grd.ActiveRow.Cells("ALLO_GROUP_CODE").Value & ""
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.Cells("ALLO_GROUP_CODE").Value & "" = ALLO_GROUP_CODE Then
                        grow.Cells("SELECTED").Value = "1"
                        grow.Update()
                    End If
                Next

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("Load", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Value
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Remove from List (No Update)"
                Dim ALLO_CTL_NO = grd.ActiveRow.Cells("ALLO_CTL_NO").Value
                Dim I As Integer = iCol(ALLO_CTL_NO)
                grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_" & Format(I, "00")).Hidden = True
                grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_NOTES_" & Format(I, "00")).Hidden = True
                iCol.Remove(ALLO_CTL_NO)
                ALLO_CTL_NOi(I) = ""
                ALLO_CTL_NOs = Replace(ALLO_CTL_NOs, ", '" & ALLO_CTL_NO & "'", "")
                Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO)
                rowSOTALLO1.Delete()

            Case "Clear Allocation Qtys"
                Dim ALLO_CTL_NO As String = grd.ActiveRow.Cells("ALLO_CTL_NO").Value
                Dim i As Integer = iCol(ALLO_CTL_NO)
                For Each rowSOTALLOC As DataRow In dst.Tables("SOTALLOC").Select("")
                    rowSOTALLOC.Item("ALLO_" & Format(i, "00")) = DBNull.Value
                    rowSOTALLOC.Item("ALLO_NOTES_" & Format(i, "00")) = DBNull.Value
                Next

            Case "Delete Allocation"
                Dim ALLO_CTL_NO As String = grd.ActiveRow.Cells("ALLO_CTL_NO").Value

                ASCMAIN1.sql = "Select Count (*) from SOTORDR2 where ALLO_CTL_NO = '" & ALLO_CTL_NO & "'"
                Dim USES As Integer = Val(ASCDATA1.GetDataValue)

                If USES <> 0 Then
                    Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Value
                    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                    If rowICTITEM1.Item("ITEM_NOT_ALLOCATED") & "" <> "1" Then
                        If ASCMAIN1.CLIENT = "INT" Then
                            'SP AND CM WANT TO DELETE ALLOCATION HISTORY, REGARDLESS IF THE ALLOCATION WAS USED, AT THE END OF THE SEASON.
                        Else
                            MsgBox("You Cannot Delete an Allocation if it has been used unless the Item is flagged as Not-Allocated", MsgBoxStyle.OkOnly, "Cannot Peform Update Requested")
                            Exit Sub
                        End If
                    End If
                End If

                Dim i As Integer = iCol(ALLO_CTL_NO)
                For Each rowSOTALLOC As DataRow In dst.Tables("SOTALLOC").Select("")
                    rowSOTALLOC.Item("ALLO_" & Format(i, "00")) = DBNull.Value
                    rowSOTALLOC.Item("ALLO_NOTES_" & Format(i, "00")) = DBNull.Value
                Next
                If Not ALLO_CTL_NOs_to_Delete.Contains(ALLO_CTL_NO) Then
                    ALLO_CTL_NOs_to_Delete.Add(ALLO_CTL_NO)
                End If
                grd.ActiveRow.Delete(False)
                For ictr As Integer = 1 To iColumn
                    If ALLO_CTL_NOi(ictr) = ALLO_CTL_NO Then
                        ALLO_CTL_NOi(ictr) = ""
                        Exit For
                    End If
                Next

                grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_" & Format(i, "00")).Hidden = True
                grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_NOTES_" & Format(i, "00")).Hidden = True
                iCol.Remove(ALLO_CTL_NO)



            Case "Copy Allocation Dates & Qtys"
                ALLO_CTL_NO_to_copy = grd.ActiveRow.Cells("ALLO_CTL_NO").Value
                grdSOTALLO1.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

            Case "Copy Customer Qtys"
                CUST_CODE_to_copy = grd.ActiveRow.Cells("CUST_CODE").Value
                grdSOTALLOC.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

            Case "Copy Store Qtys"
                CUST_STORE_NO_to_copy = grd.ActiveRow.Cells("CUST_STORE_NO").Value
                grdSOTALLOS.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

            Case "Paste Qtys from Excel"
                Dim clipboardText As String = Clipboard.GetText().Trim()
                If String.IsNullOrEmpty(clipboardText) Then Exit Sub

                Dim rows() As String = clipboardText.Split({vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
                If rows.Length = 0 Then Exit Sub

                Dim activeRow As UltraWinGrid.UltraGridRow = grdSOTALLOS.ActiveRow
                If activeRow Is Nothing AndAlso grdSOTALLOS.Selected.Rows.Count > 0 Then
                    activeRow = grdSOTALLOS.Selected.Rows(0)
                End If

                Dim activeColumn As UltraWinGrid.UltraGridColumn = grdSOTALLOS.ActiveCell?.Column
                If activeRow Is Nothing OrElse activeColumn Is Nothing Then Exit Sub

                For Each rowValue As String In rows
                    If Not activeRow.Cells.Exists(activeColumn.Key) Then Continue For
                    activeRow.Cells(activeColumn.Key).Value = rowValue.Trim()
                    activeRow.Update()
                    activeRow = activeRow.GetSibling(UltraWinGrid.SiblingRow.Next)

                    If activeRow Is Nothing Then Exit For
                Next

            Case "Clear Quantities"
                If grdSOTALLOS.ActiveCell Is Nothing Then Exit Sub

                Dim columnKey As String = grdSOTALLOS.ActiveCell.Column.Key

                For Each row As UltraWinGrid.UltraGridRow In grdSOTALLOS.Rows
                    If row.IsDataRow Then
                        row.Cells(columnKey).Value = DBNull.Value
                        row.Update()
                    End If
                Next

            Case "Copy Event Qtys"
                If grdSOTALLOS.Rows.Count = 0 Then Exit Sub

                For Each row As UltraWinGrid.UltraGridRow In grdSOTALLOS.Rows
                    row.Cells("ALLO_01").Value = row.Cells("TOTAL_EVENT_QTY").Value
                    row.Update()
                Next

            Case "Remove Event"
                If grdSOTALLOT.ActiveCell Is Nothing Then Exit Select

                Dim col As UltraWinGrid.UltraGridColumn = grdSOTALLOT.ActiveCell.Column
                If col Is Nothing OrElse Not col.Key.StartsWith("EVENT_") Then
                    MsgBox("Please right-click on an EVENT column to remove it.", MsgBoxStyle.Exclamation)
                    Exit Select
                End If

                If MsgBox("Are you sure you want to remove column '" & col.Header.Caption & "'?", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                    col.Hidden = True

                    If dst.Tables("SOTALLOT").Columns.Contains(col.Key) Then
                        For Each row As DataRow In dst.Tables("SOTALLOT").Rows
                            row(col.Key) = DBNull.Value
                        Next
                    End If

                    For Each rowSOTALLOT As UltraWinGrid.UltraGridRow In grdSOTALLOT.Rows
                        Dim CUST_CODE As String = rowSOTALLOT.Cells("CUST_CODE").Value.ToString()
                        Dim CUST_STORE_NO As String = rowSOTALLOT.Cells("CUST_STORE_NO").Value.ToString()

                        Dim QTY As Int64 = 0
                        For iCtr As Integer = 1 To 10
                            Dim colKey As String = "EVENT_" & Format(iCtr, "00")
                            If grdSOTALLOT.DisplayLayout.Bands(0).Columns.Exists(colKey) Then
                                If Not grdSOTALLOT.DisplayLayout.Bands(0).Columns(colKey).Hidden Then
                                    QTY += Val(rowSOTALLOT.Cells(colKey).Value & "")
                                End If
                            End If
                        Next

                        For Each rowSOTALLOS As UltraWinGrid.UltraGridRow In grdSOTALLOS.Rows
                            If rowSOTALLOS.Cells("CUST_CODE").Value.ToString() = CUST_CODE AndAlso
                               rowSOTALLOS.Cells("CUST_STORE_NO").Value.ToString() = CUST_STORE_NO Then

                                If rowSOTALLOS.Cells.Exists("TOTAL_EVENT_QTY") Then
                                    rowSOTALLOS.Cells("TOTAL_EVENT_QTY").Value = QTY
                                End If
                                Exit For
                            End If
                        Next
                    Next
                End If
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "ITEM_CODE"
                If Not ScreenMode Then
                    If e.KeyCode = System.Windows.Forms.Keys.Enter Then
                        Me.Find_ITEM_CODE()
                    End If
                End If

            Case "COLLECTION_CODE"
                If Not ScreenMode Then
                    If e.KeyCode = System.Windows.Forms.Keys.Enter Then
                        Load_SOTALLOX()
                    End If
                End If

            Case "CUST_CODE"
                If e.KeyCode = System.Windows.Forms.Keys.Enter Then
                    Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                    If CUST_CODE <> "" Then
                        Dim EMsg As String = ""
                        If Not Add_CUST_CODE(CUST_CODE, EMsg) Then
                            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Add Customer")
                        End If
                        Absx1.txtFor("CUST_CODE").Text = ""
                        Application.DoEvents()
                        Absx1.txtFor("CUST_CODE").Focus()
                    End If
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "ITEM_CODE"
                Find_ITEM_CODE()

            Case "COLLECTION_CODE"
                Load_SOTALLOX()

            Case "CUST_CODE"
                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                If CUST_CODE <> "" Then
                    rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                    If rowARTCUST1 Is Nothing Then
                        MsgBox("Invalid Value Specified for Customer Code (" & CUST_CODE & ")", MsgBoxStyle.OkOnly, "Cannot Add Customer")
                    Else
                        Dim rowSOTALLOC As DataRow = dst.Tables("SOTALLOC").Rows.Find(CUST_CODE)
                        If rowSOTALLOC IsNot Nothing Then
                            MsgBox("Customer " & CUST_CODE & " is already in Allocation List", MsgBoxStyle.OkOnly, "Cannot Add Customer")
                        Else
                            rowSOTALLOC = Get_SOTALLOC(rowARTCUST1)
                        End If
                        Absx1.txtFor("CUST_CODE").Text = ""
                    End If
                    Application.DoEvents()
                    Absx1.txtFor("CUST_CODE").Focus()
                End If
        End Select
    End Sub
#End Region

#Region "grdSOTALLOC"
    Private Sub grdSOTALLOC_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTALLOC.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "CUST_CODE"
                Dim CUST_CODE As String = Validate_Customer(e.Cell.Value & "") ' grdSOTORDR2.ActiveRow.Cells("ITEM_CODE").Value)
                If CUST_CODE <> "" Then
                    e.Cell.Row.Cells("CUST_NAME").Value = rowARTCUST1.Item("CUST_NAME") & ""
                    e.Cell.Row.Cells("SREP_CODE").Value = rowARTCUST1.Item("SREP_CODE") & ""
                    e.Cell.Row.Cells("TRADE_CLASS_CODE").Value = rowARTCUST1.Item("TRADE_CLASS_CODE")
                    e.Cell.Row.Cells("CUST_ALLOCATE_BY_STORE").Value = rowARTCUST1.Item("CUST_ALLOCATE_BY_STORE")
                End If

                'Case "ORDR_QTY"
                '    grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_OPEN").Value = grdSOTORDR2.ActiveRow.Cells("ORDR_QTY").Value

                'Case "ORDR_QTY_OPEN"
                '    grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Value = Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY").Value & "") _
                '        - Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_SHIP").Value & "") _
                '        - Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_OPEN").Value & "") _
                '        - Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_PICK").Value & "")
                '    If Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Text) < 0 Then
                '        grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Value = 0
                '    End If

        End Select
    End Sub

    Private Sub grdSOTALLOC_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTALLOC.AfterRowActivate

        For ictr As Integer = 1 To iColumn
            If ALLO_CTL_NOi(ictr) <> "" Then
                With grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_" & Format(ictr, "00"))
                    If grdSOTALLOC.ActiveRow.Cells("CUST_ALLOCATE_BY_STORE").Value & "" = "1" Then
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                    Else
                        .CellActivation = UltraWinGrid.Activation.AllowEdit
                    End If
                End With
            End If
        Next

        Setup_SOTALLOC()
        If tabSOTALLOC.SelectedTab.Key <> "Allocation Items" Then
            If tabSOTALLOC.Tabs("Store Allocations").Visible Then
                tabSOTALLOC.SelectedTab = tabSOTALLOC.Tabs("Store Allocations")
            End If
        End If

        Dim activeRow As UltraWinGrid.UltraGridRow = grdSOTALLOC.ActiveRow
        If activeRow IsNot Nothing AndAlso activeRow.Cells.Exists("CUST_CODE") Then
            Dim CUST_CODE As String = activeRow.Cells("CUST_CODE").Text.Trim()
            tabSOTALLOC.Tabs("Event Tags").Visible = (CUST_CODE = "IPLBAE")
            grdSOTALLOT.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.TemplateOnTop
            UltraExplorerBar1.Groups("Add Event Tags").Visible = (CUST_CODE = "IPLBAE")
        Else
            tabSOTALLOC.Tabs("Event Tags").Visible = False
            grdSOTALLOT.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            UltraExplorerBar1.Groups("Add Event Tags").Visible = False
        End If
    End Sub

    Private Sub grdSOTALLOC_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSOTALLOC.AfterRowUpdate
        If grdSOTALLOC.Tag & "" = "S" Then
        Else
            If grdSOTALLOC.ActiveRow.Cells("CUST_ALLOCATE_BY_STORE").Value & "" = "1" Then
                If grdSOTALLOS.ActiveRow IsNot Nothing Then
                    grdSOTALLOS.Tag = "C"
                    For ictr As Integer = 1 To iColumn
                        If ALLO_CTL_NOi(ictr) <> "" Then
                            Dim QTY_ALLO As Int64 = Val(dst.Tables("SOTALLOS").Compute("SUM(ALLO_" & Format(ictr, "00") & ")", "CUST_CODE = '" & grdSOTALLOC.ActiveRow.Cells("CUST_CODE").Value & "'") & "")
                            Dim QTY_ALLOC As Int64 = Val(e.Row.Cells("ALLO_" & Format(ictr, "00")).Value & "")
                            grdSOTALLOS.ActiveRow.Cells("ALLO_" & Format(ictr, "00")).Value = QTY_ALLOC - QTY_ALLO
                        End If
                    Next
                    grdSOTALLOC.ActiveRow.Update()
                    grdSOTALLOS.Tag = ""
                End If
            End If
        End If
        Update_Totals()
    End Sub

    Private Sub grdSOTALLOC_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTALLOC.BeforeCellUpdate

        Select Case e.Cell.Column.Key

        End Select

    End Sub

    Private Sub grdSOTALLOC_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTALLOC.BeforeExitEditMode

        ' similar code in grdSOTALLOCS

        If grdSOTALLOC.ActiveCell IsNot Nothing Then
            With grdSOTALLOC.ActiveCell
                Select Case .Column.Key
                    Case "CUST_CODE"
                        .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)

                End Select

                If .Column.Key.StartsWith("ALLO_") And Not .Column.Key.StartsWith("ALLO_NOTES") Then
                    Dim QTY_ALLO As Integer = Val(.EditorResolved.Value & "")
                    Dim i As Integer = Val(Mid(.Column.Key, 6))
                    Dim ALLO_CTL_NO As String = ALLO_CTL_NOi(i)
                    Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO)
                    Dim ITEM_CODE As String = rowSOTALLO1.Item("ITEM_CODE")
                    Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)

                    ' SO Min Qty
                    Dim ITEM_SO_QTY_MIN As Int32 = Val(rowICTITEM1.Item("ITEM_SO_QTY_MIN") & String.Empty)
                    ' SO Multiple
                    Dim ITEM_SO_QTY_MULT As Int32 = Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & String.Empty)
                    '' Inner Pack
                    'Dim ITEM_STD_PACK_SLS As Int32 = Val(rowICTITEM1.Item("ITEM_STD_PACK_SLS") & String.Empty)
                    '' Allow Half Pack
                    'Dim ITEM_ALLOW_HALF_PACK As Boolean = (Val(rowICTITEM1.Item("ITEM_ALLOW_HALF_PACK") & String.Empty) = 1)

                    'Order Quantity meets Min Qty And Order Multiple restictions
                    If QTY_ALLO > 0 And QTY_ALLO < ITEM_SO_QTY_MIN Then
                        QTY_ALLO = (ITEM_SO_QTY_MIN)
                        .EditorResolved.Value = QTY_ALLO
                    End If
                    If ITEM_SO_QTY_MULT <> 0 AndAlso QTY_ALLO Mod ITEM_SO_QTY_MULT <> 0 Then
                        QTY_ALLO += (ITEM_SO_QTY_MULT - (QTY_ALLO Mod ITEM_SO_QTY_MULT))
                        .EditorResolved.Value = QTY_ALLO
                    End If
                End If

            End With
        End If
    End Sub

    Private Sub grdSOTALLOC_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTALLOC.BeforeRowUpdate

        ' Validate_Columns("CUST_CODE", e.Cancel)
        'If Not e.Cancel Then
        '    Validate_Columns("ORDR_QTY", e.Cancel) 'THIS IS DONE IN BEFORECOLUPDATE
        'End If

        If e.Cancel = True Then
            Exit Sub
        End If

        ' ITEM_CODE_last_entry = e.Row.Cells("ITEM_CODE").Value & ""

        'If e.Row.IsAddRow Then
        '    e.Row.Cells("ORDR_NO").Value = ORDR_NO

        'End If
    End Sub

    Private Sub grdSOTALLOC_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTALLOC.ClickCellButton

        Dim COLUMN_NAME As String = e.Cell.Column.Key

        Select Case COLUMN_NAME
            Case "CUST_CODE"

                Dim sql_where As String = ""
                grdClickCellButton(grdSOTALLOC, sql_where)

        End Select
    End Sub

#End Region


    Sub Validate_Columns(COLUMN_NAME As String, ByRef Cancel As Boolean)
        With grdSOTALLOC.ActiveRow
            Select Case COLUMN_NAME

                Case "CUST_CODE"
                    If .Cells("CUST_CODE").Text <> "" Then
                        Dim CUST_CODE As String = Validate_Customer(.Cells("CUST_CODE").Value & "")
                        Cancel = (CUST_CODE = "")
                    End If

            End Select
        End With
    End Sub

    Function Validate_Customer(CUST_CODE_z As String) As String
        Dim E As String = ""

        Dim CUST_CODE As String = ""
        rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE_z)

        If rowARTCUST1 Is Nothing Then
            E = "Customer is Not on File" & vbCrLf
        Else
            If rowARTCUST1.Item("CUST_STATUS") & "" <> "A" Then
                E = "Customer Status is not Active" & vbCrLf
            End If
            If rowARTCUST1.Item("SREP_CODE") & "" = "" Then
                E = "Customer does not have a valid Sales Rep" & vbCrLf
            End If
            If rowARTCUST1.Item("TRADE_CLASS_CODE") & "" = "" Then
                E = "Customer does not have a valid Trade Class" & vbCrLf
            End If
            'If rowICTITEM1.Item("SALES_DIVISION_CODE") & "" = "" Then
            '    E = "Item does not have a valid Division Code" & vbCrLf
            'End If
        End If

        If E <> "" And grdSOTALLOC.ActiveRow IsNot Nothing AndAlso grdSOTALLOC.ActiveRow.IsAddRow Then
            MsgBox(E, MsgBoxStyle.OkOnly, "Customer Code Entered is Invalid because ...")
        Else
            If E = "" Then
                CUST_CODE = rowARTCUST1.Item(0)
            End If
        End If
        Return CUST_CODE
    End Function

    Sub Setup_tabSOTALLOX()
        UltraExplorerBar1.Groups("Items").Visible = (tabSOTALLOX.SelectedTab.Key = "Items" Or tabSOTALLOX.SelectedTab.Key = "Allocations")
        UltraExplorerBar1.Groups("Allocation Status").Visible = Not ScreenMode And (tabSOTALLOX.SelectedTab.Key = "Status by Item" Or tabSOTALLOX.SelectedTab.Key = "Status by Customer")
        UltraExplorerBar1.Groups("Changes at Shipment").Visible = Not ScreenMode And (tabSOTALLOX.SelectedTab.Key = "Changes at Shipment")

        If tabSOTALLOX.SelectedTab.Key = "Changes at Shipment" Then
            If Not grdSOTORDRU.Visible Then
                Find_Changes()
            End If
        End If

        If (tabSOTALLOX.SelectedTab.Key = "Status by Item" Or tabSOTALLOX.SelectedTab.Key = "Status by Customer") Then
            If tvwDQ.Nodes.Count = 0 Then
                Setup_Allocation_Status()
            End If
        End If
    End Sub
    Sub Load_SOTALLOX()

        dst.Tables("SOTALLOX").Rows.Clear()
        dst.Tables("ICTITEM1").Rows.Clear()

        COLLECTION_CODE = Absx1.txtFor("COLLECTION_CODE").Text
        If LookUp("ICTCOLL1", COLLECTION_CODE) Is Nothing Then
            COLLECTION_CODE = ""
        End If

        If COLLECTION_CODE = "" Then

            grdSOTALLOX.Text = "Allocations"
            grdICTITEM1.Text = "Items"

        Else
            ASCMAIN1.sql = "Select SOTALLO1.*" & sql_ICTITEM1 _
                & ", X.WHSE_QTY_ON_HAND" & vbCrLf _
                & " from ICTITEM1,SOTALLO1" & vbCrLf _
                & ", (Select ITEM_CODE, Sum (WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND from ICTSTAT2" & vbCrLf _
                & $" where WHSE_CODE IN ({sqlWHSE_CODEs}) group by ITEM_CODE) X" & vbCrLf _
                & " where ICTITEM1.COLLECTION_CODE = '" & COLLECTION_CODE & "'" & vbCrLf _
                & "   and SOTALLO1.ITEM_CODE = ICTITEM1.ITEM_CODE" & vbCrLf _
                & "   and X.ITEM_CODE (+) = SOTALLO1.ITEM_CODE"
            Fill_Records("SOTALLOX", "", True, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE" & sql_ICTITEM1 _
                    & " from ICTITEM1 where COLLECTION_CODE = '" & COLLECTION_CODE & "'"
            Fill_Records("ICTITEM1", "", True, ASCMAIN1.sql)

            Set_SOTALLOX()

            ASCMAIN1.sql = "Select * from SOTALLO1 where ITEM_CODE in (Select ITEM_CODE from (" & ASCMAIN1.sql & "))"

            For Each row As DataRow In dst.Tables("SOTALLOX").Select("", "ITEM_CODE,ALLO_CTL_NO DESC")
                Dim ITEM_CODE As String = row.Item("ITEM_CODE")
                Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
                If rowICTITEM1.Item("ALLO_CTL_NO") & "" = "" Then
                    rowICTITEM1.Item("ALLO_CTL_NO") = row.Item("ALLO_CTL_NO")
                    rowICTITEM1.Item("DATE_START") = row.Item("DATE_START")
                    rowICTITEM1.Item("DATE_END") = row.Item("DATE_END")
                End If
            Next
        End If
    End Sub

    Sub Set_SOTALLOX()

        grdSOTALLOX.Text = "Allocations for Items in Collection " & COLLECTION_CODE
        grdICTITEM1.Text = "Items in Collection " & COLLECTION_CODE

        Dim sqlw As String = ""
        If optSN.Value <> "*" Then
            sqlw &= " and ITEM_SNU_CODE = '" & optSN.Value & "'"
            grdSOTALLOX.Text &= ", " & optSN.Text
            grdICTITEM1.Text &= ", " & optSN.Text
        End If
        If optBP.Value <> "*" Then
            sqlw &= " and ITEM_BASIC_PROMO = '" & optBP.Value & "'"
            grdSOTALLOX.Text &= ", " & optBP.Text
            grdICTITEM1.Text &= ", " & optBP.Text
        End If
        If chkActiveOnly.Checked Then
            sqlw &= " and ITEM_STATUS = 'A'"
            grdSOTALLOX.Text &= ", " & chkActiveOnly.Text
            grdICTITEM1.Text &= ", " & chkActiveOnly.Text
        End If

        Dim sqld As String = " and DATE_END >= '" & Format(dteEndDate.DateTime, "MM/dd/yyyy") & "'"

        Dim dvw As DataView
        dvw = DirectCast(grdSOTALLOX.DataSource, DataTable).DefaultView
        dvw.RowFilter = Mid(sqlw & sqld, 5)
        dvw = DirectCast(grdICTITEM1.DataSource, DataTable).DefaultView
        dvw.RowFilter = Mid(sqlw, 5)

        Sort_grdColumns(grdSOTALLOX, "ITEM_CODE")
        Sort_grdColumns(grdICTITEM1, "ITEM_CODE")
    End Sub

    Private Sub optSN_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optSN.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Set_SOTALLOX()
    End Sub

    Private Sub optBP_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optBP.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Set_SOTALLOX()
    End Sub

    Private Sub chkActiveOnly_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkActiveOnly.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Set_SOTALLOX()
    End Sub

    Sub Find_ITEM_CODE()
        Dim ITEM_CODE As String = Absx1.txtFor("ITEM_CODE").Text
        If ITEM_CODE <> "" Then
            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            If rowICTITEM1 Is Nothing Then
                MsgBox("Invalid Item Code Specified (" & ITEM_CODE & ")")
                Exit Sub
            Else
                Dim COLLECTION_CODE As String = rowICTITEM1.Item("COLLECTION_CODE")
                If COLLECTION_CODE = "" Then
                    MsgBox("No Collection Code Specified for Item " & ITEM_CODE)
                    Exit Sub
                Else
                    Absx1.txtFor("COLLECTION_CODE").Text = COLLECTION_CODE
                    Load_SOTALLOX()
                    If InquiryMode Then
                        tabSOTALLOX.SelectedTab = tabSOTALLOX.Tabs("Allocations")
                    Else
                        tabSOTALLOX.SelectedTab = tabSOTALLOX.Tabs("Items")
                    End If
                    'tabSOTALLOX.SelectedTab = tabSOTALLOX.Tabs("Items")
                    rowICTITEM1 = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
                    rowICTITEM1.Item("SELECTED") = "1"

                    optBP.CheckedIndex = 0
                    optSN.CheckedIndex = 0

                    For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTITEM1, grdSOTALLOX}
                        For Each row As UltraWinGrid.UltraGridRow In grd.Rows
                            If row.Cells("ITEM_CODE").Value = ITEM_CODE Then
                                row.Selected = True
                                grd.DisplayLayout.RowScrollRegions(0).ScrollRowIntoView(row)
                            End If
                        Next
                    Next

                End If
            End If
        End If
    End Sub

#Region "grdSOTALLOX"

    Private Sub grdSOTALLOX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTALLOX.DoubleClickRow
        If ScreenMode Then
            Dim ALLO_CTL_NO As String = e.Row.Cells("ALLO_CTL_NO").Value
            Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO)
            If rowSOTALLO1 IsNot Nothing Then
                If MsgBox("Allocation " & ALLO_CTL_NO & " is already in the Allocation List.", MsgBoxStyle.OkOnly, "Cannot Add an Allocation Twice") Then
                    Exit Sub
                End If
            End If

            If Add_Allocations(",'" & ALLO_CTL_NO & "'", False, True) Then
                rowSOTALLO1 = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO)
                Add_Allocation_to_Grid(rowSOTALLO1)
                MsgBox("Allocation " & ALLO_CTL_NO & " for " & rowSOTALLO1.Item("ITEM_CODE") & " has been Added",
                       MsgBoxStyle.OkOnly, "Verification")
            End If
        End If
    End Sub

    Private Sub grdSOTALLOX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTALLOX.InitializeRow
        With e.Row.Cells("SELECTED")
            If .Value & "" = "1" Then
                .Appearance.BackColor = System.Drawing.Color.Red
            Else
                .Appearance.BackColor = System.Drawing.Color.Empty
            End If
        End With
    End Sub

#End Region

#Region "grdICTITEM1"

    Private Sub grdICTITEM1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdICTITEM1.AfterRowActivate

    End Sub

    Private Sub grdICTITEM1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTITEM1.DoubleClickRow
        If ScreenMode Then
            Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Value
            If ITEM_CODE_new.Contains(ITEM_CODE) Then
                If MsgBox("Item " & ITEM_CODE & " is already in the Allocation List with a new allocation." _
                          & vbCrLf & vbCrLf & "Do you really want to add it again?",
                          MsgBoxStyle.YesNo, "Item has already been selected to create a New Allocation") = MsgBoxResult.No Then
                    Exit Sub
                End If
            End If
            Dim rowSOTALLO1 As DataRow = Add_Item(ITEM_CODE, True)
            If rowSOTALLO1 IsNot Nothing Then
                Add_Allocation_to_Grid(rowSOTALLO1)
                MsgBox("A New Allocation record has been added for Item " & ITEM_CODE, MsgBoxStyle.OkOnly, "Verification")
            End If
        End If
    End Sub

    Private Sub grdICTITEM1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTITEM1.InitializeRow
        With e.Row.Cells("SELECTED")
            If .Value & "" = "1" Then
                .Appearance.BackColor = System.Drawing.Color.Red
            Else
                .Appearance.BackColor = System.Drawing.Color.Empty
            End If
        End With
    End Sub
#End Region

    Function Add_Allocations(sqlALLO_CTL_NOs As String, clear_table As Boolean, Optional multi_task As Boolean = False) As Boolean

        If multi_task Then
            Dim ALLO_CTL_NO As String = Replace(Mid(sqlALLO_CTL_NOs, 2), "'", "")
            If Not ASCMAIN1.Logical_Lock("SOTALLO1", ALLO_CTL_NO, False, True, False) Then
                Return False
            End If
            Dim row As DataRow = LookUp("SOTALLO1", ALLO_CTL_NO)
            If row IsNot Nothing Then
                Dim ITEM_CODE As String = row.Item("ITEM_CODE")
                If Not ASCMAIN1.Logical_Lock("SOTALLO1", "ITEM:" & ITEM_CODE) Then
                    Exit Function
                End If
            End If
        End If

        ASCMAIN1.sql = "Select SOTALLO1.*" & sql_ICTITEM1 _
            & ", X.WHSE_QTY_ON_HAND" & vbCrLf _
            & " from SOTALLO1,ICTITEM1" & vbCrLf _
            & ", (Select ITEM_CODE, Sum (WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND from ICTSTAT2" & vbCrLf _
            & $" where WHSE_CODE IN ({sqlWHSE_CODEs}) group by ITEM_CODE) X" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = SOTALLO1.ITEM_CODE" & vbCrLf _
            & "   and SOTALLO1.ALLO_CTL_NO IN (" & Mid(sqlALLO_CTL_NOs, 2) & ")" & vbCrLf _
            & "   and X.ITEM_CODE (+) = SOTALLO1.ITEM_CODE"
        Fill_Records("SOTALLO1", "", clear_table, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select SOTALLO2.*, ARTCUST1.CUST_NAME, ARTCUST1.SREP_CODE, ARTCUST1.TRADE_CLASS_CODE, ARTCUST1.CUST_ALLOCATE_BY_STORE" & vbCrLf _
            & " from SOTALLO2, ARTCUST1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = SOTALLO2.CUST_CODE" & vbCrLf _
            & " and SOTALLO2.ALLO_CTL_NO IN (" & Mid(sqlALLO_CTL_NOs, 2) & ")"
        Fill_Records("SOTALLO2", "", clear_table, ASCMAIN1.sql)

        ASCMAIN1.sql = Construct_sqlSOTALLO3() _
            & " and SOTALLO3.ALLO_CTL_NO IN (" & Mid(sqlALLO_CTL_NOs, 2) & ")"
        Fill_Records("SOTALLO3", "", clear_table, ASCMAIN1.sql)

        ASCMAIN1.sql = Construct_sqlSOTALLO4() _
            & " and SOTALLO4.ALLO_CTL_NO IN (" & Mid(sqlALLO_CTL_NOs, 2) & ")"
        Fill_Records("SOTALLO4", "", clear_table, ASCMAIN1.sql)
        Return True
    End Function

    Function Construct_sqlSOTALLO3() As String
        ASCMAIN1.sql = "Select SOTALLO3.*" & vbCrLf _
            & ", NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_NAME) CUST_STORE_LOCATION" & vbCrLf _
            & ", DECODE(ARTCUSF2.CUST_CODE,NULL,ARTCUST2.SELL_CODE,ARTCUSF2.SELL_CODE) SELL_CODE" & vbCrLf _
            & " from SOTALLO3, ARTCUST2, " & ARTCUSF2 & " ARTCUSF2" & vbCrLf _
            & " where ARTCUST2.CUST_CODE = SOTALLO3.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_STORE_NO = SOTALLO3.CUST_STORE_NO" & vbCrLf _
            & sqlwhere_F _
            & "   and ARTCUSF2.CUST_CODE (+) = SOTALLO3.CUST_CODE" & vbCrLf _
            & "   and ARTCUSF2.CUST_STORE_NO (+) = SOTALLO3.CUST_STORE_NO" & vbCrLf
        Return ASCMAIN1.sql
    End Function
    Function Construct_sqlSOTALLO4() As String
        ASCMAIN1.sql = "Select SOTALLO4.*" & vbCrLf _
        & ", NVL(ARTCUST2.CUST_STORE_LOCATION, ARTCUST2.CUST_STORE_NAME) CUST_STORE_LOCATION" & vbCrLf _
        & ", DECODE(ARTCUSF2.CUST_CODE, NULL, ARTCUST2.SELL_CODE, ARTCUSF2.SELL_CODE) SELL_CODE" & vbCrLf _
        & " from SOTALLO4, ARTCUST2, " & ARTCUSF2 & " ARTCUSF2" & vbCrLf _
        & " where ARTCUST2.CUST_CODE = SOTALLO4.CUST_CODE" & vbCrLf _
        & "   and ARTCUST2.CUST_STORE_NO = SOTALLO4.CUST_STORE_NO" & vbCrLf _
        & sqlwhere_F _
        & "   and ARTCUSF2.CUST_CODE (+) = SOTALLO4.CUST_CODE" & vbCrLf _
        & "   and ARTCUSF2.CUST_STORE_NO (+) = SOTALLO4.CUST_STORE_NO" & vbCrLf
        Return ASCMAIN1.sql
    End Function

    Function Add_Allocation_to_Grid(rowSOTALLO1 As DataRow) As Int64

        iColumn += 1

        Dim ALLO_CTL_NO As String = rowSOTALLO1.Item("ALLO_CTL_NO")

        With grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_" & Format(iColumn, "00"))
            .Hidden = False
            .Width = 100
            .Tag = ALLO_CTL_NO
        End With
        With grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_NOTES_" & Format(iColumn, "00"))
            .Hidden = True
            .Width = 100
            .Tag = ALLO_CTL_NO
        End With
        With grdSOTALLOS.DisplayLayout.Bands(0).Columns("ALLO_" & Format(iColumn, "00"))
            .Hidden = False
            .Width = 100
            .Tag = ALLO_CTL_NO
        End With
        With grdSOTALLOS.DisplayLayout.Bands(0).Columns("ALLO_NOTES_" & Format(iColumn, "00"))
            .Hidden = True
            .Width = 100
            .Tag = ALLO_CTL_NO
        End With

        dst.Tables("SOTALLOC").Columns("ALLO_" & Format(iColumn, "00")).ExtendedProperties("ALLO_CTL_NO") = ALLO_CTL_NO
        dst.Tables("SOTALLOS").Columns("ALLO_" & Format(iColumn, "00")).ExtendedProperties("ALLO_CTL_NO") = ALLO_CTL_NO

        dst.Tables("SOTALLOC").Columns("ALLO_NOTES_" & Format(iColumn, "00")).ExtendedProperties("ALLO_CTL_NO") = ALLO_CTL_NO
        dst.Tables("SOTALLOS").Columns("ALLO_NOTES_" & Format(iColumn, "00")).ExtendedProperties("ALLO_CTL_NO") = ALLO_CTL_NO

        For Each rowSOTALLO2 As DataRow In dst.Tables("SOTALLO2").Select("ALLO_CTL_NO = '" & ALLO_CTL_NO & "'")
            'Dim CUST_CODE As String = rowSOTALLO2.Item("CUST_CODE")
            Dim rowSOTALLOC As DataRow = Get_SOTALLOC(rowSOTALLO2)
            rowSOTALLOC.Item("ALLO_" & Format(iColumn, "00")) = Val(rowSOTALLO2.Item("QTY_ALLO") & "")
            rowSOTALLOC.Item("ALLO_NOTES_" & Format(iColumn, "00")) = rowSOTALLO2.Item("ALLO_NOTES") & ""
        Next


        For Each rowSOTALLO3 As DataRow In dst.Tables("SOTALLO3").Select("ALLO_CTL_NO = '" & ALLO_CTL_NO & "'")
            Dim CUST_CODE As String = rowSOTALLO3.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = rowSOTALLO3.Item("CUST_STORE_NO")
            Dim rowSOTALLOS As DataRow = dst.Tables("SOTALLOS").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
            If rowSOTALLOS Is Nothing Then
                rowSOTALLOS = dst.Tables("SOTALLOS").NewRow
                rowSOTALLOS.Item("CUST_CODE") = rowSOTALLO3.Item("CUST_CODE")
                rowSOTALLOS.Item("CUST_STORE_NO") = rowSOTALLO3.Item("CUST_STORE_NO")
                rowSOTALLOS.Item("CUST_STORE_LOCATION") = rowSOTALLO3.Item("CUST_STORE_LOCATION")
                rowSOTALLOS.Item("SELL_CODE") = rowSOTALLO3.Item("SELL_CODE")
                'rowSOTALLOS.Item("DMA_CODE") = rowSOTALLO3.Item("DMA_CODE")
                'rowSOTALLOS.Item("CUST_STORE_STATUS") = rowSOTALLO3.Item("CUST_STORE_STATUS")
                dst.Tables("SOTALLOS").Rows.Add(rowSOTALLOS)
            End If
            rowSOTALLOS.Item("ALLO_" & Format(iColumn, "00")) = Val(rowSOTALLO3.Item("QTY_ALLO") & "")
            rowSOTALLOS.Item("ALLO_NOTES_" & Format(iColumn, "00")) = rowSOTALLO3.Item("ALLO_NOTES") & ""
        Next

        For Each rowSOTALLOS As DataRow In dst.Tables("SOTALLOS").Select("CUST_STORE_STATUS = 'A'")
            Dim CUST_CODE As String = rowSOTALLOS.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = rowSOTALLOS.Item("CUST_STORE_NO")

            Dim rowSOTALLOT As DataRow = dst.Tables("SOTALLOT").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
            If rowSOTALLOT Is Nothing Then
                rowSOTALLOT = dst.Tables("SOTALLOT").NewRow
                rowSOTALLOT.Item("CUST_CODE") = CUST_CODE
                rowSOTALLOT.Item("CUST_STORE_NO") = CUST_STORE_NO
                rowSOTALLOT.Item("CUST_STORE_LOCATION") = rowSOTALLOS.Item("CUST_STORE_LOCATION")
                rowSOTALLOT.Item("SELL_CODE") = rowSOTALLOS.Item("SELL_CODE")
                dst.Tables("SOTALLOT").Rows.Add(rowSOTALLOT)
            End If
        Next

        For Each rowSOTALLO4 As DataRow In dst.Tables("SOTALLO4").Select("ALLO_CTL_NO = '" & ALLO_CTL_NO & "'")
            Dim CUST_CODE As String = rowSOTALLO4.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = rowSOTALLO4.Item("CUST_STORE_NO")
            Dim EVENT_NAME As String = rowSOTALLO4.Item("EVENT")
            Dim QTY_ALLO As Int64 = rowSOTALLO4.Item("QTY_ALLO")

            Dim eventColumn As UltraWinGrid.UltraGridColumn = Nothing
            For i As Integer = 1 To 10
                Dim colKey As String = "EVENT_" & Format(i, "00")
                If grdSOTALLOT.DisplayLayout.Bands(0).Columns.Exists(colKey) Then
                    Dim col As UltraWinGrid.UltraGridColumn = grdSOTALLOT.DisplayLayout.Bands(0).Columns(colKey)
                    If col.Hidden Then
                        eventColumn = col
                        Exit For
                    ElseIf col.Header.Caption = EVENT_NAME Then
                        eventColumn = col
                        Exit For
                    End If
                End If
            Next

            If eventColumn IsNot Nothing Then
                eventColumn.Hidden = False
                eventColumn.Header.Caption = EVENT_NAME

                Dim rowSOTALLOT As DataRow = dst.Tables("SOTALLOT").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
                If rowSOTALLOT IsNot Nothing Then
                    rowSOTALLOT.Item(eventColumn.Key) = QTY_ALLO
                End If
            End If
        Next



        iCol.Add(ALLO_CTL_NO, iColumn)
        ALLO_CTL_NOi(iColumn) = ALLO_CTL_NO
        Set_Header(ALLO_CTL_NO)
        Return iColumn
    End Function

    Sub Set_Header(ALLO_CTL_NO As String)
        Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO)
        Dim ITEM_CODE As String = rowSOTALLO1.Item("ITEM_CODE")
        Dim ITEM_DESC As String = rowSOTALLO1.Item("ITEM_DESC")
        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)

        Dim DATE_START As String = ""
        If rowSOTALLO1.Item("DATE_START") & "" <> "" Then DATE_START = Format(rowSOTALLO1.Item("DATE_START"), "MM/dd/yy")

        Dim DATE_END As String = ""
        If rowSOTALLO1.Item("DATE_END") & "" <> "" Then DATE_END = Format(rowSOTALLO1.Item("DATE_END"), "MM/dd/yy")

        Dim header As String = "" _
            & ITEM_CODE & vbCrLf _
            & DATE_START & vbCrLf _
            & DATE_END & vbCrLf

        Dim DATES As String = ""
        If DATE_START <> "" And DATE_END <> "" Then
            DATES = Format(rowSOTALLO1.Item("DATE_START"), "MM/dd") & "-" & Format(rowSOTALLO1.Item("DATE_END"), "MM/dd")
        End If

        header = "" _
            & ITEM_CODE & vbCrLf _
            & ITEM_DESC & vbCrLf _
            & DATES & vbCrLf

        Dim I As Integer = iCol(ALLO_CTL_NO)
        With grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_" & Format(I, "00"))
            If Not chkSingleItemMode.Checked Then .Header.Caption = header
            .Header.Tag = header
            .Header.ToolTipText = rowICTITEM1.Item("ITEM_DESC") & ""
        End With
        With grdSOTALLOS.DisplayLayout.Bands(0).Columns("ALLO_" & Format(I, "00"))
            If Not chkSingleItemMode.Checked Then .Header.Caption = ITEM_CODE
            .Header.Tag = ITEM_CODE
        End With
    End Sub

    Function Add_Item(ITEM_CODE As String, Optional multi_task As Boolean = False) As DataRow

        If multi_task Then
            If Not ASCMAIN1.Logical_Lock("SOTALLO1", "ITEM:" & ITEM_CODE, False, True, False) Then
                Return Nothing
            End If
        End If

        Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
        Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").NewRow

        Dim ALLO_CTL_NO As String = ASCMAIN1.Next_Control_No("SOTALLO1.ALLO_CTL_NO")
        ALLO_CTL_NO_new.Add(ALLO_CTL_NO)
        If Not ITEM_CODE_new.Contains(ITEM_CODE) Then ITEM_CODE_new.Add(ITEM_CODE)

        rowSOTALLO1.Item("ALLO_CTL_NO") = ALLO_CTL_NO
        rowSOTALLO1.Item("ITEM_CODE") = ITEM_CODE
        rowSOTALLO1.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
        rowSOTALLO1.Item("ITEM_BASIC_PROMO") = rowICTITEM1.Item("ITEM_BASIC_PROMO")
        rowSOTALLO1.Item("ITEM_SNU_CODE") = rowICTITEM1.Item("ITEM_SNU_CODE")
        rowSOTALLO1.Item("ITEM_CATGY_CODE") = rowICTITEM1.Item("ITEM_CATGY_CODE")
        rowSOTALLO1.Item("ITEM_STATUS") = rowICTITEM1.Item("ITEM_STATUS")
        rowSOTALLO1.Item("PROD_CODE") = rowICTITEM1.Item("PROD_CODE")
        rowSOTALLO1.Item("ITEM_NOT_ALLOCATED") = rowICTITEM1.Item("ITEM_NOT_ALLOCATED")
        rowSOTALLO1.Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
        rowSOTALLO1.Item("ITEM_SO_QTY_MULT") = rowICTITEM1.Item("ITEM_SO_QTY_MULT")
        rowSOTALLO1.Item("ITEM_SO_QTY_MIN") = rowICTITEM1.Item("ITEM_SO_QTY_MIN")
        rowSOTALLO1.Item("ITEM_CODE_COMPARE_TO") = rowICTITEM1.Item("LIKE_AS_ITEM")

        ASCMAIN1.sql = $"Select Sum (WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND from ICTSTAT2 where WHSE_CODE IN ({sqlWHSE_CODEs}) and ITEM_CODE = :PARM1"
        Dim WHSE_QTY_ON_HAND As Int32 = Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New String() {ITEM_CODE}))

        rowSOTALLO1.Item("WHSE_QTY_ON_HAND") = WHSE_QTY_ON_HAND

        'rowSOTALLO1.Item("DATE_START") = String.Empty
        'rowSOTALLO1.Item("DATE_END") = String.Empty

        rowSOTALLO1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowSOTALLO1.Item("INIT_DATE") = DATETIME_STAMP
        rowSOTALLO1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowSOTALLO1.Item("LAST_DATE") = DATETIME_STAMP

        rowSOTALLO1.Item("ALLOW_OVER") = "0"
        rowSOTALLO1.Item("NEW") = "1"
        dst.Tables("SOTALLO1").Rows.Add(rowSOTALLO1)

        ALLO_CTL_NOs &= ",'" & ALLO_CTL_NO & "'"
        Return rowSOTALLO1
    End Function

    Private Sub tabSOTALLOC_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabSOTALLOC.SelectedTabChanged
        If SELECTION_NO = 0 OrElse tabSOTALLOC.SelectedTab Is Nothing Then Exit Sub
        Setup_tabSOTALLOC()
    End Sub

    Sub Setup_tabSOTALLOC()
        If tabSOTALLOC.SelectedTab.Key = "Order Details" Or tabSOTALLOC.SelectedTab.Key = "Order Summary" Then
            Get_Sales()
            If grdSOTALLOC.ActiveRow IsNot Nothing AndAlso grdSOTALLOC.ActiveRow.IsDataRow Then
                Setup_SOTALLOC()
            End If
        End If

        If tabSOTALLOC.SelectedTab.Key = "Store Allocations" Then
            Setup_ALLOC_SPREAD_options()
        End If

        ' the problem with this is that if we ever run sopordr0_g on the order group, it will pick up the order again
        ' UltraExplorerBar1.Groups("Sales").Visible = (tabSOTALLOC.SelectedTab.Key = "Order Details")
    End Sub

    Sub Get_Sales_NOT_IN_USE()
        Dim COL As String = ""
        If grdSOTALLOC.ActiveCell IsNot Nothing Then
            COL = grdSOTALLOC.ActiveCell.Column.Key
        End If
        If Not COL.StartsWith("ALLO_") Then
            grdSOTORDR2.Visible = False
            grdSOTORDRA.Visible = False
        Else
            Dim ALLO_CTL_NO As String = grdSOTALLOC.ActiveCell.Column.Tag
            Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO)
            If rowSOTALLO1.Item("DATE_START") & "" = "" _
            Or rowSOTALLO1.Item("DATE_END") & "" = "" Then
                dst.Tables("SOTORDR2").Rows.Clear()
                dst.Tables("SOTORDRA").Rows.Clear()
            Else
                Dim ITEM_CODE As String = rowSOTALLO1.Item("ITEM_CODE")
                Dim DATE_START As Date = rowSOTALLO1.Item("DATE_START")
                Dim DATE_END As Date = rowSOTALLO1.Item("DATE_END")
                Fill_Records("SOTORDR2", New Object() {ITEM_CODE, ALLO_CTL_NO, DATE_START, DATE_END})
                Fill_Records("SOTORDRA", New Object() {ITEM_CODE, ALLO_CTL_NO, DATE_START, DATE_END})
            End If

            grdSOTORDR2.Visible = True
            grdSOTORDRA.Visible = True
        End If
    End Sub

    Sub Get_Sales()
        If grdSOTALLO1.ActiveRow Is Nothing Then
            grdSOTORDR2.Visible = False
            grdSOTORDRA.Visible = False
        Else

            ' THIS IS NOT WORKING WHEN YOU CALL UP MULTIPLE ITEMS
            ' NEED TO GET THE ALLO_CTL_NO AND ITEM AND CUST FROM THE ACTIVECELL
            ' AND NEED TO WIRE IN GRDSOTALLOC.AFTERCELLACTIVATE
            Dim ALLO_CTL_NO As String = grdSOTALLO1.ActiveRow.Cells("ALLO_CTL_NO").Value & "" '  grdSOTALLOC.ActiveCell.Column.Tag
            rowSOTALLO1_Maintain_Sales = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO)
            Dim ITEM_CODE As String = rowSOTALLO1_Maintain_Sales.Item("ITEM_CODE")
            grdSOTORDR2.Text = "Allocation " & ALLO_CTL_NO & " for Item " & ITEM_CODE
            If rowSOTALLO1_Maintain_Sales.Item("DATE_START") & "" = "" _
            Or rowSOTALLO1_Maintain_Sales.Item("DATE_END") & "" = "" Then
                dst.Tables("SOTORDR2").Rows.Clear()
                grdSOTORDR2.Text &= " - No Start or End Date"
            Else
                Dim DATE_START As Date = rowSOTALLO1_Maintain_Sales.Item("DATE_START")
                Dim DATE_END As Date = rowSOTALLO1_Maintain_Sales.Item("DATE_END")
                Fill_Records("SOTORDR2", New Object() {ITEM_CODE, ALLO_CTL_NO, DATE_START, DATE_END})
            End If
            grdSOTORDR2.Visible = True

            grdSOTORDRA.Text = "Allocation " & ALLO_CTL_NO & " for Item " & ITEM_CODE
            If rowSOTALLO1_Maintain_Sales.Item("DATE_START") & "" = "" _
            Or rowSOTALLO1_Maintain_Sales.Item("DATE_END") & "" = "" Then
                dst.Tables("SOTORDRA").Rows.Clear()
                grdSOTORDRA.Text &= " - No Start or End Date"
            Else
                Dim DATE_START As Date = rowSOTALLO1_Maintain_Sales.Item("DATE_START")
                Dim DATE_END As Date = rowSOTALLO1_Maintain_Sales.Item("DATE_END")
                Fill_Records("SOTORDRA", New Object() {ITEM_CODE, ALLO_CTL_NO, DATE_START, DATE_END})
            End If
            grdSOTORDRA.Visible = True

        End If
    End Sub

#Region "grdSOTALLO1"

    Private Sub grdSOTALLO1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTALLO1.AfterRowActivate
        If chkSingleItemMode.Checked Then
            With dst.Tables("SOTALLOC")
                .Columns("SI_PLAN").Expression = ""
                .Columns("ST_PLAN").Expression = ""
                .Columns("SI_HIST").Expression = ""
                .Columns("ST_HIST").Expression = ""
            End With
            With dst.Tables("SOTALLOS")
                .Columns("SI_PLAN").Expression = ""
                .Columns("ST_PLAN").Expression = ""
                .Columns("SI_HIST").Expression = ""
                .Columns("ST_HIST").Expression = ""
            End With

            Setup_Single_Item()
            Setup_Plan_Hist()
        End If

        Dim ITEM_CODE As String = grdSOTALLO1.ActiveRow.Cells("ITEM_CODE").Value
        chkSingleItemMode.Text = ITEM_CODE

        If grdSOTALLO1.ActiveRow IsNot Nothing AndAlso grdSOTALLO1.ActiveRow.Cells.Exists("ALLO_CTL_NO") Then
            Dim ALLO_CTL_NO As String = grdSOTALLO1.ActiveRow.Cells("ALLO_CTL_NO").Value.ToString()
            LoadAllocationChangeHistory(ALLO_CTL_NO)
        End If
    End Sub

    Private Sub grdSOTALLO1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTALLO1.AfterRowUpdate

        Dim ALLO_CTL_NO As String = e.Row.Cells("ALLO_CTL_NO").Value & ""
        Set_Header(ALLO_CTL_NO)
        Setup_Plan_Hist()
        grdSOTALLOC.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)
    End Sub

    Private Sub grdSOTALLO1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTALLO1.InitializeRow
        Dim DATE_START As String = ""
        If e.Row.Cells("DATE_START").Value & "" <> "" Then DATE_START = Format(e.Row.Cells("DATE_START").Value, "yyyyMMdd")
        Dim DATE_END As String = ""
        If e.Row.Cells("DATE_END").Value & "" <> "" Then DATE_END = Format(e.Row.Cells("DATE_END").Value, "yyyyMMdd")

        If DATE_START = "" Then
            e.Row.Cells("DATE_START").Appearance.BackColor = System.Drawing.Color.Yellow
        Else
            e.Row.Cells("DATE_START").Appearance.BackColor = System.Drawing.Color.Empty
        End If
        If DATE_END = "" Then
            e.Row.Cells("DATE_END").Appearance.BackColor = System.Drawing.Color.Yellow
        Else
            e.Row.Cells("DATE_END").Appearance.BackColor = System.Drawing.Color.Empty
        End If
        If DATE_START <> "" And DATE_END <> "" Then
            If DATE_START > DATE_END Then
                e.Row.Cells("DATE_END").Appearance.ForeColor = System.Drawing.Color.Red
            Else
                e.Row.Cells("DATE_END").Appearance.ForeColor = System.Drawing.Color.Empty
            End If
        End If

        If e.Row.Cells("ALLO_CTL_NO").Value & "" = ALLO_CTL_NO_to_copy Then
            e.Row.Cells("ALLO_CTL_NO").Appearance.ForeColor = System.Drawing.Color.Green
        Else
            e.Row.Cells("ALLO_CTL_NO").Appearance.ForeColor = System.Drawing.Color.Empty
        End If

        If e.Row.Cells("NEW").Value & "" = "1" Then
            e.Row.Cells("ALLO_CTL_NO").Appearance.ForeColor = System.Drawing.Color.Green
            e.Row.Cells("ALLO_CTL_NO").ToolTipText = "New Allocation"
        End If
    End Sub

    Private Sub grdSOTALLO1_KeyDown(sender As Object, e As KeyEventArgs) Handles grdSOTALLO1.KeyDown
        If e.KeyCode = Keys.Enter Then
            If grdSOTALLO1.ActiveRow IsNot Nothing AndAlso grdSOTALLO1.ActiveRow.DataChanged Then
                grdSOTALLO1.ActiveRow.Update()
            End If
        ElseIf e.KeyCode = Keys.Delete Or e.KeyCode = Keys.Space Then
            If grdSOTALLO1.ActiveRow IsNot Nothing AndAlso grdSOTALLO1.ActiveCell IsNot Nothing AndAlso grdSOTALLO1.ActiveCell.Column.Key = "ALLO_GROUP_CODE" Then
                grdSOTALLO1.ActiveRow.Cells("ALLO_GROUP_CODE").Value = DBNull.Value
            End If
        End If
    End Sub
#End Region

    Private Sub grdSOTALLOC_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTALLOC.InitializeRow
        If e.Row.Cells("CUST_CODE").Value & "" = CUST_CODE_to_copy Then
            e.Row.Cells("CUST_CODE").Appearance.ForeColor = System.Drawing.Color.Green
        Else
            e.Row.Cells("CUST_CODE").Appearance.ForeColor = System.Drawing.Color.Empty
        End If
        Dim BALANCE As Int64 = Val(e.Row.Cells("BALANCE").Value & "")
        With e.Row.Cells("BALANCE").Appearance
            If BALANCE < 0 Then
                .BackColor = System.Drawing.Color.Red
                .ForeColor = System.Drawing.Color.White
            Else
                .BackColor = System.Drawing.Color.Empty
                .ForeColor = System.Drawing.Color.Empty
            End If
        End With

        e.Row.Cells("SI_PLAN").ToolTipText = "Double-Click to Copy All Sell-In Plan Qtys to Allocation"
        e.Row.Cells("ST_PLAN").ToolTipText = "Double-Click to Copy All Sell-Thru Plan Qtys to Allocation"
        e.Row.Cells("SI_HIST").ToolTipText = "Double-Click to Copy All Sell-In History Qtys to Allocation"
        e.Row.Cells("ST_HIST").ToolTipText = "Double-Click to Copy All Sell-Thru History Qtys to Allocation"

    End Sub

    Private Sub dteEndDate_ValueChanged(sender As Object, e As EventArgs) Handles dteEndDate.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_SOTALLOX()
    End Sub

    Sub Toggle_Maintain_Sales(tf As Boolean)

        Dim ALLO_CTL_NO As String = ""
        Dim ITEM_CODE As String = ""
        If tf Then
            ALLO_CTL_NO = rowSOTALLO1_Maintain_Sales.Item("ALLO_CTL_NO")
            ITEM_CODE = rowSOTALLO1_Maintain_Sales.Item("ITEM_CODE")
        End If

        With UltraExplorerBar1
            .Groups("Screen Control").Enabled = Not tf
            .Groups("Customers").Enabled = Not tf
            With .Groups("Sales")
                If tf Then
                    .Items("Maintain Sales").Settings.Enabled = DefaultableBoolean.False
                    .Items("Update Changes").Settings.Enabled = DefaultableBoolean.True
                    .Items("Cancel Changes").Settings.Enabled = DefaultableBoolean.True
                Else
                    .Items("Maintain Sales").Settings.Enabled = DefaultableBoolean.True
                    .Items("Update Changes").Settings.Enabled = DefaultableBoolean.False
                    .Items("Cancel Changes").Settings.Enabled = DefaultableBoolean.False
                End If
            End With
        End With

        tabSOTALLOC.Tabs("Allocation Items").Enabled = Not tf
        tabSOTALLOC.Tabs("Other Items").Enabled = Not tf
        tabSOTALLOC.Tabs("Other Allocations").Enabled = Not tf

        ' grdSOTORDR2.DisplayLayout.Bands(0).Columns("USED").Hidden = Not tf
        grdSOTORDR2.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = Not tf
        grdSOTORDRA.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = Not tf

        ' grdSOTALLOC.DisplayLayout.Bands(0).Columns("USED").Hidden = Not tf
        ' grdSOTALLOC.DisplayLayout.Bands(0).Columns("BALANCE").Hidden = Not tf

        grdSOTORDR2.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdSOTORDRA.DisplayLayout.Bands(0).SortedColumns.Clear()

        bln_Maintain_Sales = tf

        If tf Then
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Setting up for Maintenance")

            With dst.Tables("SOTORDR2")
                .Columns("USED").Expression = "IIF(ISNULL(SELECTED,'0')='1',ISNULL(ORDR_QTY_OPEN,0)+ISNULL(ORDR_QTY_PICK,0)+ISNULL(ORDR_QTY_SHIP,0),NULL)"
            End With

            dst.Tables("ARTCUST1").Rows.Clear()

            If Not InquiryMode Then
                grdSOTORDR2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            End If
            grdSOTORDR2.DisplayLayout.Bands(0).SortedColumns.Add("CUST_CODE", False, True)

            If Not InquiryMode Then
                grdSOTORDRA.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            End If
            grdSOTORDRA.DisplayLayout.Bands(0).SortedColumns.Add("CUST_CODE", False, True)

            grdSOTALLOC.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            HIDDEN_COLS.Clear()
            Dim iMaintained As Integer = 0
            For i As Integer = 1 To maxAllocations
                With grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_" & Format(i, "00"))
                    If Not .Hidden Then
                        If .Tag = ALLO_CTL_NO Then
                            iMaintained = i
                        Else
                            .Hidden = True
                            HIDDEN_COLS.Add(i)
                        End If
                    End If
                End With
            Next
            dst.Tables("SOTALLOC").Columns("BALANCE").Expression = "ISNULL(ALLO_" & Format(iMaintained, "00") & ",0) - ISNULL(USED,0)"
            '"ISNULL(ALLO_01,0)-ISNULL(USED,0)"
            'dst.Tables("SOTALLOS").Columns("BALANCE").Expression = "ISNULL(ALLO_" & Format(iMaintained, "00") & ",0) - ISNULL(USED,0)"
            For Each rowSOTALLOC As DataRow In dst.Tables("SOTALLOC").Select("")
                rowSOTALLOC.Item("USED") = DBNull.Value
            Next

            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
                If rowSOTORDR2.Item("ALLO_CTL_NO") & "" = "" Then
                    rowSOTORDR2.Item("SELECTED") = "0"
                Else
                    rowSOTORDR2.Item("SELECTED") = "1"
                End If
            Next

            For Each rowCUST_CODE As DataRow In ASCDATA1.SelectDistinct _
                (dst.Tables("SOTORDR2"), New String() {"CUST_CODE"}).Select("")
                Dim CUST_CODE As String = rowCUST_CODE.Item(0)
                Update_USED(CUST_CODE)
            Next

            Sort_grdColumns(grdSOTALLOC, "CUST_CODE")

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        Else

            With dst.Tables("SOTORDR2")
                .Columns("USED").Expression = "IIF(ISNULL(ALLO_CTL_NO,'')<>'',ISNULL(ORDR_QTY_OPEN,0)+ISNULL(ORDR_QTY_PICK,0)+ISNULL(ORDR_QTY_SHIP,0),NULL)"
            End With

            If Not InquiryMode Then
                grdSOTALLOC.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            End If
            For Each i As Integer In HIDDEN_COLS
                grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_" & Format(i, "00")).Hidden = False
            Next
            grdSOTORDR2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grdSOTORDRA.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        End If

    End Sub

    Sub Update_USED(CUST_CODE As String)

        Dim row As DataRow = dst.Tables("ARTCUST1").Rows.Find(CUST_CODE)
        If row Is Nothing Then
            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
            row = dst.Tables("ARTCUST1").NewRow
            row.Item("CUST_CODE") = rowARTCUST1.Item("CUST_CODE")
            row.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME")
            row.Item("CUST_CODE_ALLO") = rowARTCUST1.Item("CUST_CODE_ALLO")
            row.Item("CUST_ALLOCATE_BY_STORE") = rowARTCUST1.Item("CUST_ALLOCATE_BY_STORE")
            dst.Tables("ARTCUST1").Rows.Add(row)
        End If
        Dim CUST_CODE_ALLO As String = CUST_CODE
        If row.Item("CUST_CODE_ALLO") & "" <> "" Then
            CUST_CODE_ALLO = row.Item("CUST_CODE_ALLO")
        End If

        Dim rowSOTALLOC As DataRow = dst.Tables("SOTALLOC").Rows.Find(CUST_CODE_ALLO)
        If rowSOTALLOC Is Nothing Then
            rowSOTALLOC = dst.Tables("SOTALLOC").Rows.Add(New String() {CUST_CODE_ALLO})
        End If

        rowSOTALLOC.Item("USED") = 0

        Dim sqlw As String = "CUST_CODE = '" & CUST_CODE & "' or CUST_CODE_ALLO = '" & CUST_CODE_ALLO & "'"
        For Each rowARTCUST1 As DataRow In dst.Tables("ARTCUST1").Select(sqlw)
            Dim CUST_CODE_SALES As String = rowARTCUST1.Item("CUST_CODE")
            Dim USED As Int64 = Val(dst.Tables("SOTORDR2").Compute("SUM(USED)", "CUST_CODE = '" & CUST_CODE_SALES & "'") & "")
            rowSOTALLOC.Item("USED") = Val(rowSOTALLOC.Item("USED") & "") + USED
        Next
    End Sub

    Private Sub grdSOTORDR2_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSOTORDR2.AfterRowUpdate
        Update_USED(e.Row.Cells("CUST_CODE").Value)
    End Sub

    Private Sub grdSOTORDRA_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSOTORDRA.AfterRowUpdate
        Update_USED(e.Row.Cells("CUST_CODE").Value)
    End Sub

    Sub Setup_SOTALLOC()

        CUST_STORE_NO_to_copy = ""

        If grdSOTALLOC.ActiveRow Is Nothing OrElse (grdSOTALLOC.ActiveRow.IsFilterRow Or Not grdSOTALLOC.ActiveRow.IsDataRow) Then
            tabSOTALLOC.Tabs("Store Allocations").Visible = False
        Else
            Dim CUST_CODE As String = grdSOTALLOC.ActiveRow.Cells("CUST_CODE").Value & ""
            Dim rowSOTALLOC As DataRow = dst.Tables("SOTALLOC").Rows.Find(CUST_CODE)
            If rowSOTALLOC.Item("CUST_ALLOCATE_BY_STORE") & "" = "1" Then
                Dim sqlw As String = " and (CUST_STORE_STATUS = 'A' or ISNULL(ORDR_QTY,0) <> 0 or ISNULL(QTY_LEFT,0) <> 0 or ISNULL(QTY_BAL,0) <> 0 or ISNULL(SI_HIST,0) <> 0 or ISNULL(ST_HIST,0) <> 0 or ISNULL(LY_QTY_SELL_IN,0) <> 0 or ISNULL(LY_QTY_SELL_THRU,0) <> 0)"
                Dim dvw As DataView = DirectCast(grdSOTALLOS.DataSource, DataTable).DefaultView
                dvw.RowFilter = "CUST_CODE = '" & CUST_CODE & "'" & sqlw
                Sort_grdColumns(grdSOTALLOS, "CUST_STORE_NO")
                grdSOTALLOS.Text = "Store Allocations for " & CUST_CODE
                tabSOTALLOC.Tabs("Store Allocations").Visible = True
            Else
                tabSOTALLOC.Tabs("Store Allocations").Visible = False
            End If

        End If

        Setup_ALLOC_SPREAD_options()

        If grdSOTALLOC.ActiveRow Is Nothing OrElse (grdSOTALLOC.ActiveRow.IsFilterRow Or Not grdSOTALLOC.ActiveRow.IsDataRow) Then
        Else
            Dim CUST_CODE_X As String = grdSOTALLOC.ActiveRow.Cells("CUST_CODE").Value & ""
            Dim dvwSOTORDR2 As DataView = DirectCast(grdSOTORDR2.DataSource, DataTable).DefaultView
            Dim C As String = "Sales Orders for " & CUST_CODE_X
            Dim sqlw As String = "CUST_CODE_ALLO = '" & CUST_CODE_X & "'"
            If chkSingleItemMode.Checked Then
                sqlw &= " and ALLO_CTL_NO = '" & ALLO_CTL_NOi(SINGLE_ITEM_ictr) & "'"
                C &= ", Allocation " & ALLO_CTL_NOi(SINGLE_ITEM_ictr) & ", " & grdSOTALLO1.ActiveRow.Cells("DATE_START").Text & "-" & grdSOTALLO1.ActiveRow.Cells("DATE_END").Text
            End If
            dvwSOTORDR2.RowFilter = sqlw
            Sort_grdColumns(grdSOTORDR2, "ORDR_NO")
            grdSOTORDR2.Text = C

            Dim dvwSOTORDRA As DataView = DirectCast(grdSOTORDRA.DataSource, DataTable).DefaultView
            'sqlw = Replace(sqlw, "CUST_CODE_ALLO", "CUST_CODE")
            dvwSOTORDRA.RowFilter = sqlw
            Sort_grdColumns(grdSOTORDRA, "ORDR_GROUP_NO")
            grdSOTORDRA.Text = C

        End If
    End Sub

    Sub Setup_ALLOC_SPREAD_options()

        Dim tf As Boolean = tabSOTALLOC.Tabs("Store Allocations").Visible And tabSOTALLOC.SelectedTab.Key = "Store Allocations" And chkSingleItemMode.Checked
        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTALLOC, grdSOTALLOS}
            With grd.DisplayLayout.Bands(0)
                .Columns("STORES").Hidden = Not tf
                .Columns("RETAIL_SALES").Hidden = Not tf
                .Columns("RETAIL_SALES_PCT").Hidden = Not tf
                .Columns("ALLO_SPREAD").Hidden = Not tf
            End With
        Next

        btnGet6Mos.Visible = tf

    End Sub

    Sub Add_CUST_STOREs(CUST_CODE As String)
        'ASCMAIN1.sql = "Select ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO" & vbCrLf _
        '    & ", NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_NAME) CUST_STORE_LOCATION" & vbCrLf _
        '    & ", DECODE(ARTCUSF2.CUST_CODE,NULL,ARTCUST2.SELL_CODE,ARTCUSF2.SELL_CODE) SELL_CODE, ARTCUST2.DMA_CODE, ARTCUST2.CUST_STORE_STATUS" & vbCrLf _
        '    & " from ARTCUST2, " & ARTCUSF2 & " ARTCUSF2 where ARTCUST2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
        '    & sqlwhere_F _
        '    & "   and ARTCUSF2.CUST_CODE (+) = ARTCUST2.CUST_CODE" & vbCrLf _
        '    & "   and ARTCUSF2.CUST_STORE_NO (+) = ARTCUST2.CUST_STORE_NO" & vbCrLf

        ASCMAIN1.sql = "Select ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO" & vbCrLf _
            & ", NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_NAME) CUST_STORE_LOCATION" & vbCrLf _
            & ", DECODE(ARTCUSF2.CUST_CODE,NULL,NVL(ARTCUST2.SELL_CODE_AC,ARTCUST2.SELL_CODE),NVL(ARTCUSF2.SELL_CODE_AC,ARTCUSF2.SELL_CODE)) SELL_CODE" & vbCrLf _
            & ", ARTCUST2.DMA_CODE, ARTCUST2.CUST_STORE_STATUS" & vbCrLf _
            & " from ARTCUST2, " & ARTCUSF2 & " ARTCUSF2 where ARTCUST2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & sqlwhere_F _
            & "   and ARTCUSF2.CUST_CODE (+) = ARTCUST2.CUST_CODE" & vbCrLf _
            & "   and ARTCUSF2.CUST_STORE_NO (+) = ARTCUST2.CUST_STORE_NO"

        For Each rowARTCUST2 As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim CUST_STORE_NO As String = rowARTCUST2.Item("CUST_STORE_NO")
            Dim rowSOTALLOS As DataRow = dst.Tables("SOTALLOS").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
            rowSOTALLOS = dst.Tables("SOTALLOS").NewRow
            rowSOTALLOS.Item("CUST_CODE") = CUST_CODE
            rowSOTALLOS.Item("CUST_STORE_NO") = CUST_STORE_NO
            rowSOTALLOS.Item("CUST_STORE_LOCATION") = rowARTCUST2.Item("CUST_STORE_LOCATION")
            rowSOTALLOS.Item("SELL_CODE") = rowARTCUST2.Item("SELL_CODE")
            rowSOTALLOS.Item("DMA_CODE") = rowARTCUST2.Item("DMA_CODE")
            rowSOTALLOS.Item("CUST_STORE_STATUS") = rowARTCUST2.Item("CUST_STORE_STATUS")
            dst.Tables("SOTALLOS").Rows.Add(rowSOTALLOS)
        Next

        Sort_grdColumns(grdSOTALLOS, "CUST_STORE_NO")
    End Sub

    Private Sub grdSOTALLOS_AfterColPosChanged(sender As Object, e As UltraWinGrid.AfterColPosChangedEventArgs) Handles grdSOTALLOS.AfterColPosChanged
        Dim COLUMN_NAME As String = e.ColumnHeaders(0).Column.Key
        Dim COLUMN_NAME2 As String = COLUMN_NAME
        If COLUMN_NAME = "CUST_STORE_NO" Then COLUMN_NAME2 = "CUST_CODE"
        If COLUMN_NAME = "CUST_STORE_LOCATION" Then COLUMN_NAME2 = "CUST_NAME"
        If COLUMN_NAME = "SELL_CODE" Then COLUMN_NAME2 = "SREP_CODE"
        If COLUMN_NAME = "DMA_CODE" Then COLUMN_NAME2 = "TRADE_CLASS_CODE"

        If e.PosChanged = UltraWinGrid.PosChanged.Moved Then
            grdSOTALLOC.DisplayLayout.Bands(0).Columns(COLUMN_NAME2).Header.VisiblePosition = grdSOTALLOS.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.VisiblePosition
        End If

        If e.PosChanged = UltraWinGrid.PosChanged.Sized And COLUMN_NAME & "" <> "TOTAL_EVENT_QTY" Then
            grdSOTALLOC.DisplayLayout.Bands(0).Columns(COLUMN_NAME2).Width = grdSOTALLOS.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Width
        End If
    End Sub

    Private Sub grdSOTALLOS_AfterColRegionScroll(sender As Object, e As UltraWinGrid.ColScrollRegionEventArgs) Handles grdSOTALLOS.AfterColRegionScroll
        grdSOTALLOC.DisplayLayout.ColScrollRegions(0).Position = e.ColScrollRegion.Position
    End Sub

    Private Sub grdSOTALLOS_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSOTALLOS.AfterRowUpdate
        If grdSOTALLOS.Tag = "C" Then
        Else
            grdSOTALLOC.Tag = "S"
            For ictr As Integer = 1 To iColumn
                If ALLO_CTL_NOi(ictr) <> "" Then
                    Dim QTY_ALLO As Int64 = Val(dst.Tables("SOTALLOS").Compute("SUM(ALLO_" & Format(ictr, "00") & ")", "CUST_CODE = '" & grdSOTALLOC.ActiveRow.Cells("CUST_CODE").Value & "'") & "")
                    grdSOTALLOC.ActiveRow.Cells("ALLO_" & Format(ictr, "00")).Value = QTY_ALLO
                End If
            Next
            grdSOTALLOC.ActiveRow.Update()
            grdSOTALLOC.Tag = ""
        End If
    End Sub

    Private Sub chkSingleItemMode_CheckedChanged(sender As Object, e As EventArgs) Handles chkSingleItemMode.CheckedChanged
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Setting up Screen Options")

        chkPriorYear.Enabled = chkSingleItemMode.Checked
        chkBalances.Enabled = chkSingleItemMode.Checked

        'grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_NOTES").Hidden = Not chkSingleItemMode.Checked
        'grdSOTALLOS.DisplayLayout.Bands(0).Columns("ALLO_NOTES").Hidden = Not chkSingleItemMode.Checked

        If chkSingleItemMode.Checked Then
            chkPriorYear.Checked = True
            tabSOTALLOC.SelectedTab = tabSOTALLOC.Tabs("Allocation Items")
            If grdSOTALLO1.ActiveRow Is Nothing Then
                grdSOTALLOC.Text = "Allocations by Customer"
                grdSOTALLOC.DisplayLayout.Bands(0).ColHeaderLines = 3
            Else
                Setup_Single_Item()
            End If
            Setup_Plan_Hist()

            'chkBalances.Checked = False
            'chkPriorYear.Checked = False
            For Each row As DataRow In dst.Tables("SOTALLO1").Select("")
                If row.Item("DATE_START") & "" <> "" AndAlso Format(row.Item("DATE_START"), "yyyyMMdd") < Format(Now, "yyyyMMdd") Then chkBalances.Checked = True
                If row.Item("DATE_START") & "" <> "" AndAlso Format(row.Item("DATE_START"), "yyyyMMdd") >= Format(Now, "yyyyMMdd") Then chkPriorYear.Checked = True
            Next

        Else
            chkPriorYear.Checked = False
            chkBalances.Checked = False
            grdSOTALLOC.Text = "Allocations by Customer"
            grdSOTALLOC.DisplayLayout.Bands(0).ColHeaderLines = 3
            For ictr As Integer = 1 To iColumn
                If ALLO_CTL_NOi(ictr) <> "" And ScreenMode Then ' THIS BLOCK WAS CAUSING THE INFAMOUS EXTRA COLUMNS TO APPEAR ON NEXT ALLOCATION
                    grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_" & Format(ictr, "00")).Header.Caption = grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_" & Format(ictr, "00")).Header.Tag
                    grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_" & Format(ictr, "00")).Hidden = False
                    grdSOTALLOS.DisplayLayout.Bands(0).Columns("ALLO_" & Format(ictr, "00")).Header.Caption = grdSOTALLOS.DisplayLayout.Bands(0).Columns("ALLO_" & Format(ictr, "00")).Header.Tag
                    grdSOTALLOS.DisplayLayout.Bands(0).Columns("ALLO_" & Format(ictr, "00")).Hidden = False

                    grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_NOTES_" & Format(ictr, "00")).Hidden = True
                    grdSOTALLOS.DisplayLayout.Bands(0).Columns("ALLO_NOTES_" & Format(ictr, "00")).Hidden = True
                End If
            Next
            If grdSOTALLOS.DisplayLayout.Bands(0).Columns.Exists("TOTAL_EVENT_QTY") Then
                grdSOTALLOS.DisplayLayout.Bands(0).Columns("TOTAL_EVENT_QTY").Hidden = True
            End If
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub cmdAddMultipleCustomers_Click(sender As Object, e As EventArgs) Handles cmdAddMultipleCustomers.Click

        Dim CUST_CODEsX As String = ""
        Dim CUST_CODEs As New List(Of String)
        For Each row As DataRow In dst.Tables("SOTALLOC").Select("")
            CUST_CODEs.Add(row.Item("CUST_CODE"))
        Next
        If CUST_CODEs.Count <> 0 Then CUST_CODEsX = "CUST_CODE NOT in ('" & Join(CUST_CODEs.ToArray, "','") & "')"

        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("CUST_CODE", "", CUST_CODEsX)

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            'Dim CUST_CODEs As New List(Of String)
            'For Each row As DataRow In dst.Tables("SOTALLOC").Select("")
            '    CUST_CODEs.Add(row.Item("CUST_CODE"))
            'Next
            'ASCMAIN1.CodeSelector.Custom_sql_where = " and CUST_CODE NOT in ('" & Join(CUST_CODEs.ToArray, "','") & "')"
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading")

                For Each CUST_CODE As String In ASCMAIN1.CodeSelector.SelectedCodes
                    Add_CUST_CODE(CUST_CODE, "")
                Next

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
            End If

            Sort_grdColumns(grdSOTALLOC, "CUST_CODE")
        End If
    End Sub

    Function Add_CUST_CODE(CUST_CODE As String, ByRef EMsg As String) As Boolean

        rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
        If rowARTCUST1 Is Nothing Then
            EMsg = "Invalid Value Specified for Customer Code (" & CUST_CODE & ")"
            Return False
        Else
            Dim rowSOTALLOC As DataRow = dst.Tables("SOTALLOC").Rows.Find(CUST_CODE)
            If rowSOTALLOC IsNot Nothing Then
                EMsg = "Customer " & CUST_CODE & " is already in Allocation List"
                Return False
            Else
                rowSOTALLOC = Get_SOTALLOC(rowARTCUST1)
                'If rowARTCUST1.Item("CUST_ALLOCATE_BY_STORE") & "" = "1" Then
                '    Add_CUST_STOREs(CUST_CODE)
                'End If
                Return True
            End If
        End If
    End Function

    Sub Update_Totals()

        For ictr As Integer = 1 To iColumn
            If ALLO_CTL_NOi(ictr) <> "" Then
                Dim QTY_ALLO As Int64 = Val(dst.Tables("SOTALLOC").Compute("SUM(ALLO_" & Format(ictr, "00") & ")", "") & "")
                Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NOi(ictr))
                'If rowSOTALLO1 IsNot Nothing Then
                rowSOTALLO1.Item("QTY_ALLO_TOTAL") = QTY_ALLO
                'End If
            End If
        Next

    End Sub

    Sub Setup_Plan_Hist()
        If Me.SELECTION_NO = 0 Then Exit Sub
        If grdSOTALLO1.ActiveRow Is Nothing Then
            If grdSOTALLO1.Rows.Count > 0 Then
                grdSOTALLO1.ActiveRow = grdSOTALLO1.Rows(0)
            Else
                MsgBox("No Allocations", MsgBoxStyle.OkOnly, "Cannot Setup Plan & History")
                Exit Sub
            End If
        End If
        Dim QTY_ALLO_PLAN As Int64 = Val(grdSOTALLO1.ActiveRow.Cells("QTY_ALLO_PLAN").Value & "")
        Dim LY_QTY_SELL_IN As Int64 = Val(dst.Tables("SOTALLOC").Compute("SUM(LY_QTY_SELL_IN)", "") & "")
        Dim LY_QTY_SELL_THRU As Int64 = Val(dst.Tables("SOTALLOC").Compute("SUM(LY_QTY_SELL_THRU)", "") & "")

        With dst.Tables("SOTALLOC")
            .Columns("SI_PLAN").Expression = IIf(LY_QTY_SELL_IN = 0, "0", "ISNULL(LY_QTY_SELL_IN,0) / " & CStr(LY_QTY_SELL_IN) & " * " & QTY_ALLO_PLAN)
            .Columns("ST_PLAN").Expression = IIf(LY_QTY_SELL_THRU = 0, "0", "ISNULL(LY_QTY_SELL_THRU,0) / " & CStr(LY_QTY_SELL_THRU) & " * " & QTY_ALLO_PLAN)

            Dim x As Decimal = Val(numX.Value & "")
            .Columns("SI_HIST").Expression = "ISNULL(LY_QTY_SELL_IN,0) * " & CStr(x)
            .Columns("ST_HIST").Expression = "ISNULL(LY_QTY_SELL_THRU,0) * " & CStr(x)
        End With

        grdSOTALLOC.Rows.Refresh(UltraWinGrid.RefreshRow.RefreshDisplay)

        With dst.Tables("SOTALLOS")
            .Columns("SI_PLAN").Expression = IIf(LY_QTY_SELL_IN = 0, "0", "ISNULL(LY_QTY_SELL_IN,0) / " & CStr(LY_QTY_SELL_IN) & " * " & QTY_ALLO_PLAN)
            .Columns("ST_PLAN").Expression = IIf(LY_QTY_SELL_THRU = 0, "0", "ISNULL(LY_QTY_SELL_THRU,0) / " & CStr(LY_QTY_SELL_THRU) & " * " & QTY_ALLO_PLAN)

            Dim x As Decimal = Val(numX.Value & "")
            .Columns("SI_HIST").Expression = "ISNULL(LY_QTY_SELL_IN,0) * " & CStr(x)
            .Columns("ST_HIST").Expression = "ISNULL(LY_QTY_SELL_THRU,0) * " & CStr(x)
        End With

        grdSOTALLOS.Rows.Refresh(UltraWinGrid.RefreshRow.RefreshDisplay)

    End Sub
    Sub Toggle_Display_Options()

        If Me.SELECTION_NO = 0 Then Exit Sub
        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTALLOC, grdSOTALLOS}
            With grd.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "QTY_LEFT", "QTY_BAL", "QTY_OVER"}
                    .Columns(COLUMN_NAME).Hidden = Not chkSingleItemMode.Checked Or Not chkBalances.Checked
                Next
                For Each COLUMN_NAME As String In New String() {"LY_QTY_SELL_IN", "LY_QTY_SELL_THRU", "LY_QTY_SELL_IN_THRU_PCT", "TY_VS_LY_PCT"}
                    .Columns(COLUMN_NAME).Hidden = Not chkSingleItemMode.Checked Or Not chkPriorYear.Checked
                Next

                .Columns("SI_PLAN").Hidden = Not chkSingleItemMode.Checked Or Not chkSI.Checked Or Not chkPlan.Checked
                .Columns("ST_PLAN").Hidden = Not chkSingleItemMode.Checked Or Not chkST.Checked Or Not chkPlan.Checked
                .Columns("SI_HIST").Hidden = Not chkSingleItemMode.Checked Or Not chkSI.Checked Or Not chkHist.Checked
                .Columns("ST_HIST").Hidden = Not chkSingleItemMode.Checked Or Not chkST.Checked Or Not chkHist.Checked
            End With
        Next

        Setup_ALLOC_SPREAD_options()
    End Sub

    Private Sub chkPriorYear_CheckedChanged(sender As Object, e As EventArgs) Handles chkPriorYear.CheckedChanged
        Toggle_Display_Options()
    End Sub

    Private Sub chkBalances_CheckedChanged(sender As Object, e As EventArgs) Handles chkBalances.CheckedChanged
        Toggle_Display_Options()
    End Sub

    Sub Add_ALLO_Columns2(grd As UltraWinGrid.UltraGrid)
        With grd.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() _
              {"SI_PLAN", "ST_PLAN", "SI_HIST", "ST_HIST"}
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = System.Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = System.Drawing.Color.Gold
                .Columns(COLUMN_NAME).Width = 80
                Create_Summary(grd, COLUMN_NAME)
            Next
            .Columns("SI_PLAN").Header.Caption = "S-In Plan"
            .Columns("ST_PLAN").Header.Caption = "S-Th Plan"
            .Columns("SI_HIST").Header.Caption = "S-In Hist"
            .Columns("ST_HIST").Header.Caption = "S-Th Hist"

            With .Columns("STORES")
                .Header.Caption = "Stores"
                .Width = 60
                .Format = "#,##0"
            End With

            With .Columns("RETAIL_SALES")
                .Header.Caption = "6Mos Rtl"
                .Width = 80
                .Format = "#,##0"
            End With

            With .Columns("RETAIL_SALES_PCT")
                .Header.Caption = "%Ttl"
                .Width = 50
                .Format = "#,##0.0"
            End With

            With .Columns("ALLO_SPREAD")
                .Header.Caption = "Rtl%Plan"
                .Width = 80
                .Format = "#,##0"
            End With

            For Each COLUMN_NAME As String In New String() _
              {"STORES", "RETAIL_SALES", "RETAIL_SALES_PCT", "ALLO_SPREAD"}
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = System.Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = System.Drawing.Color.Tomato
                Create_Summary(grd, COLUMN_NAME)
            Next

            For Each COLUMN_NAME As String In New String() _
                {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC",
                 "QTY_LEFT", "QTY_BAL", "QTY_OVER"}
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = System.Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                .Columns(COLUMN_NAME).Width = 80
                Create_Summary(grd, COLUMN_NAME)
            Next
            .Columns("ORDR_QTY").Header.Caption = "Order"
            .Columns("ORDR_QTY_OPEN").Header.Caption = "Open"
            .Columns("ORDR_QTY_PICK").Header.Caption = "Pick"
            .Columns("ORDR_QTY_SHIP").Header.Caption = "Ship"
            .Columns("ORDR_QTY_CANC").Header.Caption = "Canc"
            .Columns("QTY_LEFT").Header.Caption = "Left"
            .Columns("QTY_BAL").Header.Caption = "Balance"
            .Columns("QTY_OVER").Header.Caption = "Over"

            For Each COLUMN_NAME As String In New String() _
               {"LY_QTY_SELL_IN", "LY_QTY_SELL_THRU", "LY_QTY_SELL_IN_THRU_PCT", "TY_VS_LY_PCT"}
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = System.Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = System.Drawing.Color.PaleVioletRed
                .Columns(COLUMN_NAME).Width = 80
                If COLUMN_NAME = "LY_QTY_SELL_IN_THRU_PCT" Or COLUMN_NAME = "TY_VS_LY_PCT" Then
                    'Create_Summary(grd, COLUMN_NAME)
                Else
                    Create_Summary(grd, COLUMN_NAME)
                End If
            Next
            .Columns("LY_QTY_SELL_IN").Header.Caption = "LY S-In"
            .Columns("LY_QTY_SELL_THRU").Header.Caption = "LY Thru"
            .Columns("LY_QTY_SELL_IN_THRU_PCT").Header.Caption = "In/Thru%"
            .Columns("TY_VS_LY_PCT").Header.Caption = "TY/LY%"
        End With
    End Sub
    Sub Add_ALLO_Columns(grd As UltraWinGrid.UltraGrid)
        With grd.DisplayLayout.Bands(0)
            For iCtr As Integer = 1 To maxAllocations
                Dim COLUMN_NAME As String = "ALLO_" & Format(iCtr, "00")
                With .Columns(COLUMN_NAME)
                    .CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Header.Appearance.BackColor = System.Drawing.Color.White
                    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    If grdSOTALLOS.Name = "grdSOTALLOS" Then
                        .Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                    Else
                        .Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    End If
                End With
                Create_Summary(grd, COLUMN_NAME)

                COLUMN_NAME = "ALLO_NOTES_" & Format(iCtr, "00")
                With .Columns(COLUMN_NAME)
                    .CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Header.Appearance.BackColor = System.Drawing.Color.White
                    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    If grdSOTALLOS.Name = "grdSOTALLOS" Then
                        .Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                    Else
                        .Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    End If
                    .Hidden = True
                End With
            Next

            Add_ALLO_Columns2(grd)

        End With
    End Sub

    Private Sub grdSOTALLOC_AfterColRegionScroll(sender As Object, e As UltraWinGrid.ColScrollRegionEventArgs) Handles grdSOTALLOC.AfterColRegionScroll
        grdSOTALLOS.DisplayLayout.ColScrollRegions(0).Position = e.ColScrollRegion.Position
    End Sub

    Private Sub grdSOTALLOC_AfterColPosChanged(sender As Object, e As UltraWinGrid.AfterColPosChangedEventArgs) Handles grdSOTALLOC.AfterColPosChanged

        If e Is Nothing OrElse e.ColumnHeaders Is Nothing OrElse e.ColumnHeaders.Count = 0 Then Exit Sub
        If grdSOTALLOC.DisplayLayout.Bands.Count = 0 Then Exit Sub
        If grdSOTALLOS.DisplayLayout.Bands.Count = 0 Then Exit Sub

        Dim sourceColumnName As String = e.ColumnHeaders(0).Column.Key
        Dim targetColumnName As String = String.Empty

        Select Case sourceColumnName
            Case "CUST_CODE"
                targetColumnName = "CUST_STORE_NO"
            Case "CUST_NAME"
                targetColumnName = "CUST_STORE_LOCATION"
            Case "SREP_CODE"
                targetColumnName = "SELL_CODE"
            Case "TRADE_CLASS_CODE"
                targetColumnName = "DMA_CODE"
            Case Else
                Exit Sub
        End Select

        If Not grdSOTALLOC.DisplayLayout.Bands(0).Columns.Exists(sourceColumnName) Then Exit Sub
        If Not grdSOTALLOS.DisplayLayout.Bands(0).Columns.Exists(targetColumnName) Then Exit Sub

        Select Case e.PosChanged
            Case UltraWinGrid.PosChanged.Moved
                grdSOTALLOS.DisplayLayout.Bands(0).Columns(targetColumnName).Header.VisiblePosition =
                grdSOTALLOC.DisplayLayout.Bands(0).Columns(sourceColumnName).Header.VisiblePosition

            Case UltraWinGrid.PosChanged.Sized
                grdSOTALLOS.DisplayLayout.Bands(0).Columns(targetColumnName).Width =
                grdSOTALLOC.DisplayLayout.Bands(0).Columns(sourceColumnName).Width
        End Select

    End Sub
    Sub Get_Sales_STATS(Optional summary As Boolean = False, Optional ALLO_CTL_NO_to_refresh As String = "")

        If Not summary Then
            dst.Tables("SOTALLOA").Rows.Clear()

            For Each rowSOTALLO1 As DataRow In dst.Tables("SOTALLO1").Select("")
                Dim ALLO_CTL_NO As String = rowSOTALLO1.Item("ALLO_CTL_NO")
                Dim ITEM_CODE = rowSOTALLO1.Item("ITEM_CODE")
                Dim ITEM_CODE_COMPARE_TO = rowSOTALLO1.Item("ITEM_CODE_COMPARE_TO") & ""
                Dim ITEM_CODE_COMPARE_TO_ALT = rowSOTALLO1.Item("ITEM_CODE_COMPARE_TO_ALT") & ""

                If rowSOTALLO1.Item("DATE_START") & "" <> "" And rowSOTALLO1.Item("DATE_END") & "" <> "" Then
                    Dim DATE_START As Date = rowSOTALLO1.Item("DATE_START")
                    Dim DATE_END As Date = rowSOTALLO1.Item("DATE_END")

                    ASCMAIN1.sql = "Select Min (OPS_YYYYPP) from GLTPARM2 where PRD_END_DATE > '" & Format(DATE_START, "dd-MMM-yyyy") & "'"
                    Dim TYP_START As String = ASCDATA1.GetDataValue
                    ASCMAIN1.sql = "Select Min (OPS_YYYYPP) from GLTPARM2 where PRD_END_DATE > '" & Format(DATE_END, "dd-MMM-yyyy") & "'"
                    Dim TYP_END As String = ASCDATA1.GetDataValue

                    Dim LYP_START As String = ASCMAIN1.Period_Calc(TYP_START, -12)
                    Dim LYP_END As String = ASCMAIN1.Period_Calc(TYP_END, -12)

                    dst.Tables("SOTALLOA").Rows.Add(New Object() {ALLO_CTL_NO, ITEM_CODE, ITEM_CODE_COMPARE_TO, ITEM_CODE_COMPARE_TO_ALT, TYP_START, TYP_END, LYP_START, LYP_END})
                End If
            Next

            Update_Record_TDA("SOTALLOA", "1=1")

        End If

        Dim TABLE_NAME As String = SOTALLO1S
        If Not summary Then
            TABLE_NAME = SOTALLOA
        End If


        Dim sql_STATS As String = "" _
            & "Select SOTALLOA.ALLO_CTL_NO, SOTORDR2.CUST_CODE, SOTORDR2.CUST_STORE_NO" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN, SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP, SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & ", 0 LY_QTY_SELL_IN1, 0 LY_QTY_SELL_IN2, 0 LY_QTY_SELL_THRU1, 0 LY_QTY_SELL_THRU2" & vbCrLf _
            & " from SOTORDR2," & TABLE_NAME & " SOTALLOA" & vbCrLf _
            & " where SOTORDR2.ITEM_CODE = SOTALLOA.ITEM_CODE" & vbCrLf _
            & "   and SOTORDR2.ALLO_CTL_NO = SOTALLOA.ALLO_CTL_NO" & vbCrLf _
            & IIf(ALLO_CTL_NO_to_refresh = "", "", "  and SOTALLOA.ALLO_CTL_NO = '" & ALLO_CTL_NO_to_refresh & "'" & vbCrLf) _
            & " group by SOTALLOA.ALLO_CTL_NO, SOTORDR2.CUST_CODE, SOTORDR2.CUST_STORE_NO" & vbCrLf _
            & " union " & vbCrLf _
            & "Select SOTALLOA.ALLO_CTL_NO, SOTORDR2.CUST_CODE, SOTORDR2.CUST_STORE_NO, 0 ORDR_QTY" & vbCrLf _
            & ", 0 ORDR_QTY_OPEN, 0 ORDR_QTY_PICK, 0 ORDR_QTY_SHIP, 0 ORDR_QTY_CANC" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_SHIP) LY_QTY_SELL_IN1, 0 LY_QTY_SELL_IN2, 0 LY_QTY_SELL_THRU1, 0 LY_QTY_SELL_THRU2" & vbCrLf _
            & " from SOTORDR2, " & TABLE_NAME & " SOTALLOA" & vbCrLf _
            & " where SOTORDR2.ITEM_CODE = NVL(SOTALLOA.ITEM_CODE_COMPARE_TO,SOTALLOA.ITEM_CODE)" & vbCrLf _
            & "   and SOTORDR2.ORDR_YYYYPP_UPDATED Between SOTALLOA.LYP_START and SOTALLOA.LYP_END" & vbCrLf _
            & IIf(ALLO_CTL_NO_to_refresh = "", "", "  and SOTALLOA.ALLO_CTL_NO = '" & ALLO_CTL_NO_to_refresh & "'" & vbCrLf) _
            & " group by SOTALLOA.ALLO_CTL_NO, SOTORDR2.CUST_CODE, SOTORDR2.CUST_STORE_NO" & vbCrLf _
            & " union " & vbCrLf _
            & "Select SOTALLOA.ALLO_CTL_NO, SOTORDR2.CUST_CODE, SOTORDR2.CUST_STORE_NO, 0 ORDR_QTY" & vbCrLf _
            & ", 0 ORDR_QTY_OPEN, 0 ORDR_QTY_PICK, 0 ORDR_QTY_SHIP, 0 ORDR_QTY_CANC" & vbCrLf _
            & ", 0 LY_QTY_SELL_IN1, SUM (SOTORDR2.ORDR_QTY_SHIP) LY_QTY_SELL_IN2, 0 LY_QTY_SELL_THRU1, 0 LY_QTY_SELL_THRU2" & vbCrLf _
            & " from SOTORDR2, " & TABLE_NAME & " SOTALLOA" & vbCrLf _
            & " where SOTORDR2.ITEM_CODE = SOTALLOA.ITEM_CODE_COMPARE_TO_ALT" & vbCrLf _
            & "   and SOTORDR2.ORDR_YYYYPP_UPDATED Between SOTALLOA.LYP_START and SOTALLOA.LYP_END" & vbCrLf _
            & IIf(ALLO_CTL_NO_to_refresh = "", "", "  and SOTALLOA.ALLO_CTL_NO = '" & ALLO_CTL_NO_to_refresh & "'" & vbCrLf) _
            & " group by SOTALLOA.ALLO_CTL_NO, SOTORDR2.CUST_CODE, SOTORDR2.CUST_STORE_NO" & vbCrLf _
            & " union " & vbCrLf _
            & "Select SOTALLOA.ALLO_CTL_NO, RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, 0 ORDR_QTY" & vbCrLf _
            & ", 0 ORDR_QTY_OPEN, 0 ORDR_QTY_PICK, 0 ORDR_QTY_SHIP, 0 ORDR_QTY_CANC" & vbCrLf _
            & ", 0 LY_QTY_SELL_IN1, 0 LY_QTY_SELL_IN2, SUM (QTY_SOLD) LY_QTY_SELL_THRU1, 0 LY_QTY_SELL_THRU2" & vbCrLf _
            & " from RSTRETL1, " & SOTALLOA & " SOTALLOA" & vbCrLf _
            & " where RSTRETL1.ITEM_CODE = NVL(SOTALLOA.ITEM_CODE_COMPARE_TO,SOTALLOA.ITEM_CODE)" & vbCrLf _
            & "   and OPS_YYYYPP  Between SOTALLOA.LYP_START and SOTALLOA.LYP_END" & vbCrLf _
            & IIf(ALLO_CTL_NO_to_refresh = "", "", "  and SOTALLOA.ALLO_CTL_NO = '" & ALLO_CTL_NO_to_refresh & "'" & vbCrLf) _
            & " group by SOTALLOA.ALLO_CTL_NO, RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO" & vbCrLf _
            & " union " & vbCrLf _
            & "Select SOTALLOA.ALLO_CTL_NO, RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, 0 ORDR_QTY" & vbCrLf _
            & ", 0 ORDR_QTY_OPEN, 0 ORDR_QTY_PICK, 0 ORDR_QTY_SHIP, 0 ORDR_QTY_CANC" & vbCrLf _
            & ", 0 LY_QTY_SELL_IN1, 0 LY_QTY_SELL_IN2, 0 LY_QTY_SELL_THRU1, SUM (QTY_SOLD) LY_QTY_SELL_THRU2" & vbCrLf _
            & " from RSTRETL1, " & SOTALLOA & " SOTALLOA" & vbCrLf _
            & " where RSTRETL1.ITEM_CODE = SOTALLOA.ITEM_CODE_COMPARE_TO_ALT" & vbCrLf _
            & "   and OPS_YYYYPP  Between SOTALLOA.LYP_START and SOTALLOA.LYP_END" & vbCrLf _
            & IIf(ALLO_CTL_NO_to_refresh = "", "", "  and SOTALLOA.ALLO_CTL_NO = '" & ALLO_CTL_NO_to_refresh & "'" & vbCrLf) _
            & " group by SOTALLOA.ALLO_CTL_NO, RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO"


        ASCMAIN1.sql = "Select ALLO_CTL_NO, CUST_CODE, SUM (ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM (ORDR_QTY_OPEN) ORDR_QTY_OPEN, SUM (ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (ORDR_QTY_SHIP) ORDR_QTY_SHIP, SUM (ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & ", SUM (LY_QTY_SELL_IN1) LY_QTY_SELL_IN1, SUM (LY_QTY_SELL_IN2) LY_QTY_SELL_IN2" & vbCrLf _
            & ", SUM (LY_QTY_SELL_THRU1) LY_QTY_SELL_THRU1, SUM (LY_QTY_SELL_THRU2) LY_QTY_SELL_THRU2" & vbCrLf _
            & " from (" & vbCrLf _
            & sql_STATS & vbCrLf _
            & ") group by ALLO_CTL_NO, CUST_CODE"

        If Not summary Then
            Fill_Records("SOTALLOB", "", True, ASCMAIN1.sql)
        Else

            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is" & vbCrLf _
                & ASCMAIN1.sql & ";" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & SOTALLO2S & " Set " & vbCrLf _
                & "      ORDR_QTY_OPEN = R1.ORDR_QTY_OPEN" & vbCrLf _
                & "     ,ORDR_QTY_PICK = R1.ORDR_QTY_PICK" & vbCrLf _
                & "     ,ORDR_QTY_SHIP = R1.ORDR_QTY_SHIP" & vbCrLf _
                & "     ,LY_QTY_SELL_IN1 = R1.LY_QTY_SELL_IN1" & vbCrLf _
                & "     ,LY_QTY_SELL_IN2 = R1.LY_QTY_SELL_IN2" & vbCrLf _
                & "     ,LY_QTY_SELL_THRU1 = R1.LY_QTY_SELL_THRU1" & vbCrLf _
                & "     ,LY_QTY_SELL_THRU2 = R1.LY_QTY_SELL_THRU2" & vbCrLf _
                & "    where ALLO_CTL_NO = R1.ALLO_CTL_NO and CUST_CODE = R1.CUST_CODE;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

        End If


        ASCMAIN1.sql = "Select ALLO_CTL_NO, CUST_CODE, CUST_STORE_NO, SUM (ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM (ORDR_QTY_OPEN) ORDR_QTY_OPEN, SUM (ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (ORDR_QTY_SHIP) ORDR_QTY_SHIP, SUM (ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & ", SUM (LY_QTY_SELL_IN1) LY_QTY_SELL_IN1" & vbCrLf _
            & ", SUM (LY_QTY_SELL_IN2) LY_QTY_SELL_IN2" & vbCrLf _
            & ", SUM (LY_QTY_SELL_THRU1) LY_QTY_SELL_THRU1" & vbCrLf _
            & ", SUM (LY_QTY_SELL_THRU2) LY_QTY_SELL_THRU2" & vbCrLf _
            & " from (" & vbCrLf _
            & sql_STATS & vbCrLf _
            & ") group by ALLO_CTL_NO, CUST_CODE, CUST_STORE_NO"

        If Not summary Then
            Fill_Records("SOTALLOD", "", True, ASCMAIN1.sql)
        Else

            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is" & vbCrLf _
                & ASCMAIN1.sql & ";" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & SOTALLO3S & " Set " & vbCrLf _
                & "      ORDR_QTY_OPEN = R1.ORDR_QTY_OPEN" & vbCrLf _
                & "     ,ORDR_QTY_PICK = R1.ORDR_QTY_PICK" & vbCrLf _
                & "     ,ORDR_QTY_SHIP = R1.ORDR_QTY_SHIP" & vbCrLf _
                & "     ,LY_QTY_SELL_IN1 = R1.LY_QTY_SELL_IN1" & vbCrLf _
                & "     ,LY_QTY_SELL_IN2 = R1.LY_QTY_SELL_IN2" & vbCrLf _
                & "     ,LY_QTY_SELL_THRU1 = R1.LY_QTY_SELL_THRU1" & vbCrLf _
                & "     ,LY_QTY_SELL_THRU2 = R1.LY_QTY_SELL_THRU2" & vbCrLf _
                & "    where ALLO_CTL_NO = R1.ALLO_CTL_NO and CUST_CODE = R1.CUST_CODE and CUST_STORE_NO = R1.CUST_STORE_NO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

        End If


        If Not summary Then
            Set_Sales_Stats_for_Item()
            Sort_grdColumns(grdSOTALLOC, "CUST_CODE")
        End If

    End Sub

    Sub Set_Sales_Stats_for_Item()
        ASCMAIN1.Progress("Now Getting Sales History", "")

        Dim ALLO_CTL_NO As String = grdSOTALLO1.ActiveRow.Cells("ALLO_CTL_NO").Value

        For Each row As DataRow In dst.Tables("SOTALLOC").Select("")
            For Each COLUMN_NAME As String In New String() _
                {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "LY_QTY_SELL_IN", "LY_QTY_SELL_THRU"}
                row.Item(COLUMN_NAME) = DBNull.Value
            Next
        Next

        For Each row As DataRow In dst.Tables("SOTALLOB").Select("ALLO_CTL_NO = '" & ALLO_CTL_NO & "'")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim rowSOTALLOC As DataRow = dst.Tables("SOTALLOC").Rows.Find(CUST_CODE)
            If rowSOTALLOC Is Nothing Then
                Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE, True)
                rowSOTALLOC = Get_SOTALLOC(rowARTCUST1)
            End If
            If rowSOTALLOC IsNot Nothing Then
                For Each COLUMN_NAME As String In New String() _
                    {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "LY_QTY_SELL_IN", "LY_QTY_SELL_THRU"}
                    rowSOTALLOC.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                Next
            End If
        Next


        For Each row As DataRow In dst.Tables("SOTALLOS").Select("")
            For Each COLUMN_NAME As String In New String() _
                {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "LY_QTY_SELL_IN", "LY_QTY_SELL_THRU"}
                row.Item(COLUMN_NAME) = DBNull.Value
            Next
        Next

        For Each row As DataRow In dst.Tables("SOTALLOD").Select("ALLO_CTL_NO = '" & ALLO_CTL_NO & "'")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
            Dim rowSOTALLOC As DataRow = dst.Tables("SOTALLOC").Rows.Find(CUST_CODE)
            If rowSOTALLOC.Item("CUST_ALLOCATE_BY_STORE") & "" = "1" Then
                Dim rowSOTALLOS As DataRow = dst.Tables("SOTALLOS").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
                If rowSOTALLOS Is Nothing Then
                    Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE, True)
                    rowSOTALLOS = Get_SOTALLOC(rowARTCUST1)
                End If
                If rowSOTALLOS IsNot Nothing Then
                    For Each COLUMN_NAME As String In New String() _
                        {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "LY_QTY_SELL_IN", "LY_QTY_SELL_THRU"}
                        rowSOTALLOS.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                    Next
                End If
            End If
        Next

        ASCMAIN1.Progress("", "")
    End Sub
    Private Sub cmdGetSalesStats_Click(sender As Object, e As EventArgs) Handles cmdGetSalesStats.Click
        For Each row As DataRow In dst.Tables("SOTALLO1").Select("")
            If row.Item("DATE_START") & "" = "" Or row.Item("DATE_END") & "" = "" Then
                MsgBox("You must first provide Start and End Dates for all Allocations", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                Exit Sub
            End If
        Next
        Get_Sales_STATS()
    End Sub

    Sub Setup_Single_Item()
        Dim ITEM_CODE As String = grdSOTALLO1.ActiveRow.Cells("ITEM_CODE").Value
        Dim ITEM_DESC As String = grdSOTALLO1.ActiveRow.Cells("ITEM_DESC").Value
        Dim ALLO_CTL_NO As String = grdSOTALLO1.ActiveRow.Cells("ALLO_CTL_NO").Value
        Dim SE As String = ""
        If grdSOTALLO1.ActiveRow.Cells("DATE_START").Value & "" <> "" And grdSOTALLO1.ActiveRow.Cells("DATE_END").Value & "" <> "" Then
            SE = Format(grdSOTALLO1.ActiveRow.Cells("DATE_START").Value, "MM/dd/yyyy") & " thru " & Format(grdSOTALLO1.ActiveRow.Cells("DATE_END").Value, "MM/dd/yyyy")
        End If
        grdSOTALLOC.Text = "Allocations by Customer, Item " & ITEM_CODE & ":" & ITEM_DESC & ", Allocation " & ALLO_CTL_NO & " " & SE
        grdSOTALLOC.DisplayLayout.Bands(0).ColHeaderLines = 1

        For ictr As Integer = 1 To iColumn
            If ALLO_CTL_NOi(ictr) <> "" Then
                If ALLO_CTL_NOi(ictr) = ALLO_CTL_NO Then
                    grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_" & Format(ictr, "00")).Header.Caption = "Allo"
                    grdSOTALLOS.DisplayLayout.Bands(0).Columns("ALLO_" & Format(ictr, "00")).Header.Caption = "Allo"
                    grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_NOTES_" & Format(ictr, "00")).Header.Caption = "Notes"
                    grdSOTALLOS.DisplayLayout.Bands(0).Columns("ALLO_NOTES_" & Format(ictr, "00")).Header.Caption = "Notes"
                    SINGLE_ITEM_ictr = ictr
                    dst.Tables("SOTALLOC").Columns("QTY_LEFT").Expression = "ISNULL(ALLO_" & Format(ictr, "00") & ",0)-ISNULL(ORDR_QTY_SHIP,0)-ISNULL(ORDR_QTY_PICK,0)"
                    dst.Tables("SOTALLOS").Columns("QTY_LEFT").Expression = "ISNULL(ALLO_" & Format(ictr, "00") & ",0)-ISNULL(ORDR_QTY_SHIP,0)-ISNULL(ORDR_QTY_PICK,0)"
                End If
                grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_" & Format(ictr, "00")).Hidden = Not (ALLO_CTL_NOi(ictr) = ALLO_CTL_NO)
                grdSOTALLOS.DisplayLayout.Bands(0).Columns("ALLO_" & Format(ictr, "00")).Hidden = Not (ALLO_CTL_NOi(ictr) = ALLO_CTL_NO)
                grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_NOTES_" & Format(ictr, "00")).Hidden = Not chkSingleItemMode.Checked Or Not (ALLO_CTL_NOi(ictr) = ALLO_CTL_NO)
                grdSOTALLOS.DisplayLayout.Bands(0).Columns("ALLO_NOTES_" & Format(ictr, "00")).Hidden = Not chkSingleItemMode.Checked Or Not (ALLO_CTL_NOi(ictr) = ALLO_CTL_NO)
            End If
        Next

        Dim eventCount As Integer = 0
        For iCtr As Integer = 1 To 10
            Dim colKey As String = "EVENT_" & Format(iCtr, "00")
            If grdSOTALLOT.DisplayLayout.Bands(0).Columns.Exists(colKey) AndAlso Not grdSOTALLOT.DisplayLayout.Bands(0).Columns(colKey).Hidden Then
                eventCount += 1
            End If
        Next

        If eventCount >= 1 Then
            Dim insertIndex As Integer = grdSOTALLOS.DisplayLayout.Bands(0).Columns.Count
            If grdSOTALLOS.DisplayLayout.Bands(0).Columns.Exists("ALLO_NOTES_01") Then
                insertIndex = grdSOTALLOS.DisplayLayout.Bands(0).Columns("ALLO_NOTES_01").Index + 1
            End If
            Dim COLUMN_NAME As String = "TOTAL_EVENT_QTY"
            If Not grdSOTALLOS.DisplayLayout.Bands(0).Columns.Exists(COLUMN_NAME) Then
                Dim newColumn As UltraWinGrid.UltraGridColumn = grdSOTALLOS.DisplayLayout.Bands(0).Columns.Add(COLUMN_NAME, COLUMN_NAME)
                newColumn.DataType = GetType(Int64)
                newColumn.Width = 120
                newColumn.CellActivation = UltraWinGrid.Activation.NoEdit
                newColumn.Header.VisiblePosition = insertIndex
                newColumn.Header.Caption = "Total Event Qty"
                Create_Summary(grdSOTALLOS, "TOTAL_EVENT_QTY", "Sum")
            End If
            grdSOTALLOS.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = Not chkSingleItemMode.Checked
            For Each rowSOTALLOS As UltraWinGrid.UltraGridRow In grdSOTALLOS.Rows
                Dim CUST_CODE As String = rowSOTALLOS.Cells("CUST_CODE").Value.ToString()
                Dim CUST_STORE_NO As String = rowSOTALLOS.Cells("CUST_STORE_NO").Value.ToString()
                Dim QTY As Int64 = 0
                For Each rowSOTALLOT As UltraWinGrid.UltraGridRow In grdSOTALLOT.Rows
                    If rowSOTALLOT.Cells("CUST_CODE").Value.ToString() = CUST_CODE AndAlso rowSOTALLOT.Cells("CUST_STORE_NO").Value.ToString() = CUST_STORE_NO Then
                        For iCtr As Integer = 1 To 10
                            Dim colKey As String = "EVENT_" & Format(iCtr, "00")
                            If grdSOTALLOT.DisplayLayout.Bands(0).Columns.Exists(colKey) Then
                                QTY += Val(rowSOTALLOT.Cells(colKey).Value & "")
                            End If
                        Next
                        Exit For
                    End If
                Next
                rowSOTALLOS.Cells(COLUMN_NAME).Value = QTY
            Next
        Else
        End If

        Set_Sales_Stats_for_Item()

    End Sub

    Function Get_SOTALLOC(row As DataRow) As DataRow
        Dim CUST_CODE As String = row.Item("CUST_CODE")
        Dim rowSOTALLOC As DataRow = dst.Tables("SOTALLOC").Rows.Find(CUST_CODE)
        If rowSOTALLOC Is Nothing Then
            rowSOTALLOC = dst.Tables("SOTALLOC").NewRow
            rowSOTALLOC.Item("CUST_CODE") = CUST_CODE
            rowSOTALLOC.Item("CUST_NAME") = row.Item("CUST_NAME")
            rowSOTALLOC.Item("SREP_CODE") = row.Item("SREP_CODE")
            rowSOTALLOC.Item("TRADE_CLASS_CODE") = row.Item("TRADE_CLASS_CODE")
            rowSOTALLOC.Item("CUST_ALLOCATE_BY_STORE") = row.Item("CUST_ALLOCATE_BY_STORE")
            dst.Tables("SOTALLOC").Rows.Add(rowSOTALLOC)

            ASCMAIN1.sql = "Select Count (*) STORE_COUNT_HC" & vbCrLf _
                & " from SATAUTH1" & vbCrLf _
                & " where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                & "   and HC_CODE = '" & HC_CODE_lead_item & "'" & vbCrLf _
                & "   and OPS_YYYYPP_OPENED <= '" & ASCMAIN1.CYP & "'" & vbCrLf _
                & "   and OPS_YYYYPP_CLOSED is Null"
            Dim STORE_COUNT_HC As Int32 = Val(ASCDATA1.GetDataValue)
            rowSOTALLOC.Item("STORE_COUNT_HC") = STORE_COUNT_HC

            If rowSOTALLOC.Item("CUST_ALLOCATE_BY_STORE") & "" = "1" Then
                Add_CUST_STOREs(CUST_CODE)
            End If
        End If

        'If rowSOTALLOC.Item("CUST_ALLOCATE_BY_STORE") & "" = "1" Then
        '    Add_CUST_STOREs(CUST_CODE)
        'End If

        Return rowSOTALLOC
    End Function

    Sub Fill_SOTALLO2()
        dst.Tables("SOTALLO2").Rows.Clear()
        For Each row As DataRow In dst.Tables("SOTALLOC").Select
            For ictr As Integer = 1 To iColumn
                If ALLO_CTL_NOi(ictr) <> "" Then
                    Dim QTY_ALLO As Int64 = Val(row.Item("ALLO_" & Format(ictr, "00")) & "")
                    Dim ALLO_NOTES As String = row.Item("ALLO_NOTES_" & Format(ictr, "00")) & ""
                    If QTY_ALLO <> 0 Then
                        Dim rowSOTALLO2 As DataRow = dst.Tables("SOTALLO2").NewRow
                        rowSOTALLO2.Item("ALLO_CTL_NO") = ALLO_CTL_NOi(ictr)
                        rowSOTALLO2.Item("CUST_CODE") = row.Item("CUST_CODE")
                        rowSOTALLO2.Item("QTY_ALLO") = QTY_ALLO
                        rowSOTALLO2.Item("ALLO_NOTES") = ALLO_NOTES
                        dst.Tables("SOTALLO2").Rows.Add(rowSOTALLO2)
                    End If
                End If
            Next
        Next
    End Sub

    Sub Fill_SOTALLO3()
        dst.Tables("SOTALLO3").Rows.Clear()
        For Each row As DataRow In dst.Tables("SOTALLOS").Select
            For ictr As Integer = 1 To iColumn
                If ALLO_CTL_NOi(ictr) <> "" Then
                    Dim QTY_ALLO As Int64 = Val(row.Item("ALLO_" & Format(ictr, "00")) & "")
                    Dim ALLO_NOTES As String = row.Item("ALLO_NOTES_" & Format(ictr, "00")) & ""
                    If QTY_ALLO <> 0 Then
                        Dim rowSOTALLO3 As DataRow = dst.Tables("SOTALLO3").NewRow
                        rowSOTALLO3.Item("ALLO_CTL_NO") = ALLO_CTL_NOi(ictr)
                        rowSOTALLO3.Item("CUST_CODE") = row.Item("CUST_CODE")
                        rowSOTALLO3.Item("CUST_STORE_NO") = row.Item("CUST_STORE_NO")
                        rowSOTALLO3.Item("QTY_ALLO") = QTY_ALLO
                        rowSOTALLO3.Item("ALLO_NOTES") = ALLO_NOTES
                        dst.Tables("SOTALLO3").Rows.Add(rowSOTALLO3)
                    End If
                End If
            Next
        Next
    End Sub

    Sub Fill_SOTALLO4()
        dst.Tables("SOTALLO4").Rows.Clear()

        Dim eventColumns As New List(Of String)
        For Each col As UltraWinGrid.UltraGridColumn In grdSOTALLOT.DisplayLayout.Bands(0).Columns
            If col.Key.StartsWith("EVENT_") AndAlso Not col.Hidden Then
                eventColumns.Add(col.Key)
            End If
        Next

        If eventColumns.Count = 0 Then Exit Sub

        For Each row As UltraWinGrid.UltraGridRow In grdSOTALLOT.Rows
            Dim CUST_CODE As String = row.Cells("CUST_CODE").Value.ToString()
            Dim CUST_STORE_NO As String = row.Cells("CUST_STORE_NO").Value.ToString()

            Dim ALLO_CTL_NO As String = ALLO_CTL_NOi(1)

            For Each colKey As String In eventColumns
                If grdSOTALLOT.DisplayLayout.Bands(0).Columns.Exists(colKey) Then
                    Dim EVENT_QTY As Int64 = Val(row.Cells(colKey).Value & "")
                    Dim EVENT_NAME As String = grdSOTALLOT.DisplayLayout.Bands(0).Columns(colKey).Header.Caption

                    If EVENT_QTY <> 0 Then
                        Dim rowSOTALLO4 As DataRow = dst.Tables("SOTALLO4").NewRow()
                        rowSOTALLO4.Item("ALLO_CTL_NO") = ALLO_CTL_NO
                        rowSOTALLO4.Item("CUST_CODE") = CUST_CODE
                        rowSOTALLO4.Item("CUST_STORE_NO") = CUST_STORE_NO
                        rowSOTALLO4.Item("QTY_ALLO") = EVENT_QTY
                        rowSOTALLO4.Item("EVENT") = EVENT_NAME

                        dst.Tables("SOTALLO4").Rows.Add(rowSOTALLO4)
                    End If
                End If
            Next
        Next
    End Sub



    Private Sub chkSI_CheckedChanged(sender As Object, e As EventArgs) Handles chkSI.CheckedChanged
        Toggle_Display_Options()
    End Sub

    Private Sub chkST_CheckedChanged(sender As Object, e As EventArgs) Handles chkST.CheckedChanged
        Toggle_Display_Options()
    End Sub

    Private Sub chkPlan_CheckedChanged(sender As Object, e As EventArgs) Handles chkPlan.CheckedChanged
        Toggle_Display_Options()
    End Sub

    Private Sub chkHist_CheckedChanged(sender As Object, e As EventArgs) Handles chkHist.CheckedChanged
        Toggle_Display_Options()
    End Sub

    Private Sub numX_ValueChanged(sender As Object, e As EventArgs) Handles numX.ValueChanged
        Setup_Plan_Hist()
    End Sub

    Private Sub grdSOTALLOC_DoubleClickCell(sender As Object, e As UltraWinGrid.DoubleClickCellEventArgs) Handles grdSOTALLOC.DoubleClickCell
        If New String() {"SI_PLAN", "ST_PLAN", "SI_HIST", "ST_HIST"}.Contains(e.Cell.Column.Key) Then
            If MsgBox("OK to " & Mid(e.Cell.ToolTipText, "Double-Click to ".Length + 1), MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Copying " & Mid(e.Cell.ToolTipText, "Double-Click to ".Length + 1))
                For Each grow As UltraWinGrid.UltraGridRow In grdSOTALLOC.Rows
                    grow.Cells("ALLO_" & Format(SINGLE_ITEM_ictr, "00")).Value = grow.Cells(e.Cell.Column.Key).Value
                    grow.Update()
                Next
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
            End If
        End If
    End Sub

    Sub Create_Allocation_Status_Tables(Optional initialize As Boolean = False)

        If initialize Then
        Else
            ASCMAIN1.Progress("Now Creating Allocations Status Tables")
        End If

        Dim DT As String = Format(dteAllocations.Value, "dd-MMM-yyyy")

        ASCMAIN1.sql = "Select SOTALLO1.*" & vbCrLf _
            & ", '000000' TYP_START, '000000' TYP_END" & vbCrLf _
            & ", '000000' LYP_START, '000000' LYP_END" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE" & vbCrLf _
            & sql_ICTITEM1 & vbCrLf _
            & ", 0 WHSE_QTY_ON_HAND, 0 WHSE_QTY_PICK" & vbCrLf _
            & ", 0 QTY_ALLO_TOTAL" & vbCrLf _
            & " from SOTALLO1,ICTITEM1,ICTCOLL1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = SOTALLO1.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and SOTALLO1.DATE_START <= '" & DT & "'" & vbCrLf _
            & "   and SOTALLO1.DATE_END   >= '" & DT & "'"


        If initialize Then
            ASCMAIN1.sql &= " and ROWNUM < 1"
            SOTALLO1S = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTALLO1S & " Add Primary Key (ALLO_CTL_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTALLO1S)
            ASCDATA1.ExecuteSQL("Insert into " & SOTALLO1S & " " & ASCMAIN1.sql)


            ASCMAIN1.sql = "Update " & SOTALLO1S & " Set TYP_START = (Select Min (OPS_YYYYPP) from GLTPARM2 where PRD_END_DATE > DATE_START)"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Update " & SOTALLO1S & " Set TYP_END = (Select Min (OPS_YYYYPP) from GLTPARM2 where PRD_END_DATE > DATE_END)"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Update " & SOTALLO1S & " Set LYP_START = PERIOD_CALC(TYP_START,-12)"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Update " & SOTALLO1S & " Set LYP_END = PERIOD_CALC(TYP_END,-12)"
            ASCDATA1.ExecuteSQL()



            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is" & vbCrLf _
                & "  Select ITEM_CODE" & vbCrLf _
                & "  , Sum (WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND" & vbCrLf _
                & "  , Sum (WHSE_QTY_PICK) WHSE_QTY_PICK" & vbCrLf _
                & "   from ICTSTAT2" & vbCrLf _
                & "   where ITEM_CODE in (Select ITEM_CODE from " & SOTALLO1S & ")" & vbCrLf _
                & "  group by ITEM_CODE;" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & SOTALLO1S & " Set " & vbCrLf _
                & "      WHSE_QTY_ON_HAND = R1.WHSE_QTY_ON_HAND" & vbCrLf _
                & "     ,WHSE_QTY_PICK = R1.WHSE_QTY_PICK" & vbCrLf _
                & "    where ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()
        End If


        ASCMAIN1.sql = "Select SOTALLO2.*, ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & ", 0 ORDR_QTY, 0 ORDR_QTY_OPEN, 0 ORDR_QTY_PICK, 0 ORDR_QTY_SHIP, 0 ORDR_QTY_CANC" & vbCrLf _
            & ", 0 LY_QTY_SELL_IN1, 0 LY_QTY_SELL_IN2, 0 LY_QTY_SELL_THRU1, 0 LY_QTY_SELL_THRU2" & vbCrLf _
            & " from " & SOTALLO1S & " SOTALLO1,SOTALLO2,ARTCUST1" & vbCrLf _
            & " where SOTALLO2.ALLO_CTL_NO = SOTALLO1.ALLO_CTL_NO" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTALLO2.CUST_CODE"
        If initialize Then
            ASCMAIN1.sql &= " and ROWNUM < 1"
            SOTALLO2S = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTALLO2S & " Add Primary Key (ALLO_CTL_NO,CUST_CODE)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTALLO2S)
            ASCDATA1.ExecuteSQL("Insert into " & SOTALLO2S & " " & ASCMAIN1.sql)
        End If



        ASCMAIN1.sql = "Select SOTALLO3.*, ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & ", 0 ORDR_QTY, 0 ORDR_QTY_OPEN, 0 ORDR_QTY_PICK, 0 ORDR_QTY_SHIP, 0 ORDR_QTY_CANC" & vbCrLf _
            & ", 0 LY_QTY_SELL_IN1, 0 LY_QTY_SELL_IN2, 0 LY_QTY_SELL_THRU1, 0 LY_QTY_SELL_THRU2" & vbCrLf _
            & " from " & SOTALLO1S & " SOTALLO1,SOTALLO3,ARTCUST1,ARTCUST2" & vbCrLf _
            & " where SOTALLO3.ALLO_CTL_NO = SOTALLO1.ALLO_CTL_NO" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTALLO3.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_CODE = SOTALLO3.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_STORE_NO = SOTALLO3.CUST_STORE_NO"
        If initialize Then
            ASCMAIN1.sql &= " and ROWNUM < 1"
            SOTALLO3S = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTALLO3S & " Add Primary Key (ALLO_CTL_NO,CUST_CODE,CUST_STORE_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTALLO3S)
            ASCDATA1.ExecuteSQL("Insert into " & SOTALLO3S & " " & ASCMAIN1.sql)

        End If

        If Not initialize Then
            Get_Sales_STATS(True)

            ASCMAIN1.sql = "Update " & SOTALLO1S & " SOTALLO1S Set QTY_ALLO_TOTAL = (Select Sum (QTY_ALLO) from " & SOTALLO2S & " SOTALLO2S where SOTALLO2S.ALLO_CTL_NO = SOTALLO1S.ALLO_CTL_NO)"
            ASCDATA1.ExecuteSQL()
        End If

        ASCMAIN1.Progress("")

    End Sub
    Private Sub cmdGeneratePivot_Click(sender As Object, e As EventArgs) Handles cmdGeneratePivot.Click

        Dim useSSG As Boolean = chkUseSSG.Checked

        Create_Allocation_Status_Tables()

        ASCMAIN1.Progress("Now Creating Allocations Status Workbook")

        Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & ".xlsm"
        Dim DataTable As DataTable
        Dim r As Integer = 0

        Dim excel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlSourceRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim xlDestRange As Microsoft.Office.Interop.Excel.Range = Nothing

        Dim workbook As SpreadsheetGear.IWorkbook = Nothing
        Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing

        If useSSG Then
            workbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
            worksheet = workbook.Sheets("VLOOKUP")
        Else
            excel = New Microsoft.Office.Interop.Excel.Application
            wb = excel.Workbooks.Open(FILENAME)
            ws = wb.Worksheets("VLOOKUP")
        End If


        ASCMAIN1.sql = "Select ICTCOLL1.COLLECTION_CODE, ICTCOLL1.COLLECTION_NAME" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_GENDER, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf _
            & " from ICTCOLL1,ICTBRAN1 " & vbCrLf _
            & " where ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & " order by ICTCOLL1.COLLECTION_CODE"
        DataTable = ASCDATA1.GetDataTable

        If useSSG Then
            range = worksheet.Cells("K4")
            range.CopyFromDataTable(DataTable, SpreadsheetGear.Data.SetDataFlags.NoColumnHeaders)
            workbook.Names.Add("Brand_List", "=VLOOKUP!$K$3:$O$" & CStr(3 + DataTable.Rows.Count))
        Else
            r = 0
            For Each row As DataRow In DataTable.Select("", "COLLECTION_CODE")
                r += 1
                ws.Range("K" & CStr(3 + r) & ":O" & CStr(3 + r)).Value2 = row.ItemArray
            Next
            wb.Names.Add("Brand_List", "=VLOOKUP!$K$3:$O$" & CStr(3 + DataTable.Rows.Count))
        End If

        'DT = Format(dteAllocations.DateTime.Date, "dd-MMM-yyyy")
        DT = dteAllocations.DateTime.ToString("dd-MMM-yyyy")

        'D4=VLOOKUP(C4,$K$4:$L$97,2,FALSE)
        'D4=VLOOKUP(C4,Brand_List,2,FALSE)
        ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE, '=VLOOKUP(C' || TRIM(TO_CHAR(3 + ROWNUM)) || ',Brand_List,2,FALSE)' COLLECTION_NAME" & vbCrLf _
            & ", CASE WHEN NVL(ICTITEM1.COST_CATGY_CODE,'?') = 'S' THEN 'RETAIL' ELSE 'COLLATERAL' END RC_TYPE" & vbCrLf _
            & ", ICTITEM1.ITEM_DATE_TO_SHIP" & vbCrLf _
            & ", ICTITEM1.PROD_CODE, 'H1' HALF, A.QTY_ALLO" & vbCrLf _
            & " from ICTITEM1, " & vbCrLf _
            & "(Select SOTALLO1.ITEM_CODE, Sum (SOTALLO2.QTY_ALLO) QTY_ALLO" & vbCrLf _
            & " from SOTALLO1,SOTALLO2" & vbCrLf _
            & " where SOTALLO1.ALLO_CTL_NO = SOTALLO2.ALLO_CTL_NO" & vbCrLf _
            & "   and SOTALLO1.DATE_START <= '" & DT & "'" & vbCrLf _
            & "   and SOTALLO1.DATE_END   >= '" & DT & "'" & vbCrLf _
            & "   group by SOTALLO1.ITEM_CODE" & vbCrLf _
            & "   order by SOTALLO1.ITEM_CODE) A" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = A.ITEM_CODE"
        DataTable = ASCDATA1.GetDataTable

        If useSSG Then
            range = worksheet.Cells("A4")
            range.CopyFromDataTable(DataTable, SpreadsheetGear.Data.SetDataFlags.NoColumnHeaders)
            workbook.Names.Add("Item_List", "=VLOOKUP!$A$3:$I$" & CStr(3 + DataTable.Rows.Count))
            worksheet.Cells("D4:D4").Copy(worksheet.Cells("D5:D" & CStr(3 + DataTable.Rows.Count)))
        Else
            r = 0
            For Each row As DataRow In DataTable.Select("", "ITEM_CODE")
                r += 1
                ws.Range("A" & CStr(3 + r) & ":I" & CStr(3 + r)).Value2 = row.ItemArray
            Next
            wb.Names.Add("Item_List", "=VLOOKUP!$A$3:$I$" & CStr(3 + DataTable.Rows.Count))

            xlSourceRange = ws.Range("D4:D4")
            xlDestRange = ws.Range("D5:D" & CStr(3 + DataTable.Rows.Count))
            xlSourceRange.Copy(xlDestRange)
        End If





        Dim CUST_CODE_AE As String = "IPLBAE" ' Probably should be Parameterized

        ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, TRADE_CLASS_CODE from (" & vbCrLf _
            & "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME, ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & " from ARTCUST1" & vbCrLf _
            & " where CUST_CODE IN (Select Distinct CUST_CODE from " & SOTALLO2S & ")" & vbCrLf _
            & " union " & vbCrLf _
            & "Select ARTCUST2.CUST_CODE || '-' || ARTCUST2.CUST_STORE_NO CUST_CODE," & vbCrLf _
            & " ARTCUST2.CUST_STORE_NAME, ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & " from ARTCUST1,ARTCUST2" & vbCrLf _
            & " where ARTCUST2.CUST_CODE = ARTCUST1.CUST_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = '" & CUST_CODE_AE & "'" & vbCrLf _
            & ") order by CUST_CODE"
        DataTable = ASCDATA1.GetDataTable

        If useSSG Then
            range = worksheet.Cells("Q4")
            range.CopyFromDataTable(DataTable, SpreadsheetGear.Data.SetDataFlags.NoColumnHeaders)
            workbook.Names.Add("Chain_List", "=VLOOKUP!$Q$3:$S$" & CStr(3 + DataTable.Rows.Count))
        Else
            r = 0
            For Each row As DataRow In DataTable.Select("", "CUST_CODE")
                r += 1
                ws.Range("Q" & CStr(3 + r) & ":S" & CStr(3 + r)).Value2 = row.ItemArray
            Next
            wb.Names.Add("Chain_List", "=VLOOKUP!$Q$3:$S$" & CStr(3 + DataTable.Rows.Count))
        End If


        If useSSG Then
            worksheet.Visible = False
        Else
            ws.Visible = False
        End If

        If useSSG Then
            worksheet = workbook.Sheets("DATA")
        Else
            ws = wb.Worksheets("DATA")
        End If

        ASCMAIN1.sql = "Select ARTCUST1.CUST_NAME, SOTALLO2.CUST_CODE" & vbCrLf _
            & ", TO_CHAR(SOTALLO1.DATE_START,'MMDDYY') DATE_START" & vbCrLf _
            & ", SOTALLO1.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.ITEM_EAN_CODE, ICTITEM1.ITEM_RETAIL_PRICE, SOTALLO2.QTY_ALLO" & vbCrLf _
            & ", SOTALLO2.ORDR_QTY_SHIP, SOTALLO2.ORDR_QTY_OPEN, ICTITEM1.ITEM_SO_QTY_MULT" & vbCrLf _
            & ", SOTALLO1.WHSE_QTY_ON_HAND, SOTALLO1.WHSE_QTY_PICK" & vbCrLf _
            & " from " & SOTALLO1S & " SOTALLO1," & SOTALLO2S & " SOTALLO2,ARTCUST1,ICTITEM1" & vbCrLf _
            & "where ICTITEM1.ITEM_CODE = SOTALLO1.ITEM_CODE " & vbCrLf _
            & "  and ARTCUST1.CUST_CODE = SOTALLO2.CUST_CODE" & vbCrLf _
            & "  and SOTALLO2.ALLO_CTL_NO = SOTALLO1.ALLO_CTL_NO"
        DataTable = ASCDATA1.GetDataTable

        If useSSG Then
            range = worksheet.Cells("R4")
            range.CopyFromDataTable(DataTable, SpreadsheetGear.Data.SetDataFlags.NoColumnHeaders)
            workbook.Names.Add("PivotBase", "=DATA!$R$3:$AD$" & CStr(3 + DataTable.Rows.Count))
        Else
            r = 0
            For Each row As DataRow In DataTable.Select("")
                r += 1
                ws.Range("R" & CStr(3 + r) & ":AD" & CStr(3 + r)).Value2 = row.ItemArray
            Next
            wb.Names.Add("PivotBase", "=DATA!$R$3:$AD$" & CStr(3 + DataTable.Rows.Count))
        End If

        If useSSG Then
            worksheet.Cells("A4:Q4").Copy(worksheet.Cells("A4:Q" & CStr(3 + DataTable.Rows.Count)))
            worksheet.Cells("C1").Value = Now
        Else
            xlSourceRange = ws.Range("A4:Q4")
            xlDestRange = ws.Range("A4:Q" & CStr(3 + DataTable.Rows.Count))
            xlSourceRange.Copy(xlDestRange)
            ws.Cells(1, 3).Value = Now
        End If

        If useSSG Then
            worksheet.Visible = False
        Else
            ws.Visible = False
        End If

        If useSSG Then

        Else
            'excel.Run("ResetData")
        End If


        ASCMAIN1.Progress("Now Saving Workbook")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = "Allocations"
                XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsm"

                If useSSG Then
                    workbook.SaveAs(ASCMAIN1.Folders("Work") & XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbookMacroEnabled) ' SpreadsheetGear.FileFormat.OpenXMLWorkbook) ' SpreadsheetGear.FileFormat.OpenXMLWorkbookMacroEnabled)
                Else
                    Dim objOpt As Object = Nothing ' Missing.Value
                    wb.SaveAs(ASCMAIN1.Folders("Work") & XLS_FILENAME _
                              , Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)
                    wb.Close(False, objOpt, objOpt)
                End If

                success = True

            Catch ex As Exception
                Stop
            End Try
        Loop

        If useSSG Then
        Else
            excel.Quit()
            ws = Nothing
            wb = Nothing
            excel = Nothing
            xlSourceRange = Nothing
            xlDestRange = Nothing

            ReleaseCOMObject(xlDestRange)
            ReleaseCOMObject(xlSourceRange)
            ReleaseCOMObject(ws)
            ReleaseCOMObject(wb)
            ReleaseCOMObject(excel)
        End If

        Show_Document(ASCMAIN1.Folders("Work") & XLS_FILENAME)

        ASCMAIN1.Progress("")
    End Sub

    Private Sub tabSOTALLOX_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabSOTALLOX.SelectedTabChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Setup_tabSOTALLOX()
    End Sub

    Sub Setup_Allocation_Status()
        Create_Allocation_Status_Tables()

        EnforceConstraints(False)
        Fill_Records("SOTALLO1S")
        Fill_Records("SOTALLO2S")
        EnforceConstraints(True)

        ' grdSOTALLO1S.DisplayLayout.GroupByBox.Hidden = False
        With grdSOTALLO1S.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            For Each COLUMN_NAME As String In New String() {"BRAND_CODE", "COLLECTION_CODE"}
                .Columns(COLUMN_NAME).HiddenWhenGroupBy = DefaultableBoolean.True
                .SortedColumns.Add(COLUMN_NAME, False, True)
            Next
            .SortedColumns.Add("PROD_CODE", False, False)
            .SortedColumns.Add("ITEM_CODE", False, False)
        End With
        With grdSOTALLO1S.DisplayLayout.Bands(1)
            .SortedColumns.Clear()
            .SortedColumns.Add("TRADE_CLASS_CODE", False, False)
            .SortedColumns.Add("CUST_CODE", False, False)
        End With
        grdSOTALLO1S.Rows.ExpandAll(True)

        ' grdSOTALLO1S.DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.Horizontal
        grdSOTALLO1S.DisplayLayout.GroupByBox.Hidden = True
        grdSOTALLO1S.DisplayLayout.Bands(0).Override.HeaderPlacement = UltraWinGrid.HeaderPlacement.FixedOnTop
        grdSOTALLO1S.DisplayLayout.Bands(1).Override.HeaderPlacement = UltraWinGrid.HeaderPlacement.FixedOnTop
        grdSOTALLO1S.DisplayLayout.Bands(1).Override.SummaryFooterCaptionVisible = DefaultableBoolean.False

        'grdSOTALLO1S.DisplayLayout.Bands(1).Override.SummaryFooterAppearance.BackColor = Drawing.Color.LightGray
        'grdSOTALLO1S.DisplayLayout.Bands(1).Override.SummaryValueAppearance.BackColor = Drawing.Color.LightGreen
        'grdSOTALLO1S.DisplayLayout.Bands(1).Override.SummaryValueAppearance.ForeColor = Drawing.Color.Red
        grdSOTALLO1S.DisplayLayout.Override.SummaryValueAppearance.BackColor2 = Drawing.Color.LightPink
        grdSOTALLO1S.DisplayLayout.Bands(1).Override.SummaryValueAppearance.BackColor = Drawing.Color.LightGray

        For i As Integer = 0 To grdSOTALLO1S.DisplayLayout.Bands(1).Summaries.Count - 1
            grdSOTALLO1S.DisplayLayout.Bands(1).Summaries(i).Appearance.BackColor = Drawing.Color.WhiteSmoke
        Next

        Setup_Tree()
        dst.Tables("SOTALLOCS").Rows.Clear()
        tvwDQ.ActiveNode = tvwDQ.Nodes(0)
    End Sub
    Private Sub cmdAllocationStatus_Click(sender As Object, e As EventArgs) Handles cmdAllocationStatus.Click
        Setup_Allocation_Status()
    End Sub

    Sub Edit_Allocation_SOTALLO1S(ITEM_CODE As String, ALLO_CTL_NO As String)
        tabSOTALLOX.SelectedTab = tabSOTALLOX.Tabs("Items")
        Absx1.txtFor("ITEM_CODE").Text = ITEM_CODE
        Find_ITEM_CODE()
        tabSOTALLOX.SelectedTab = tabSOTALLOX.Tabs("Allocations")
        For Each grow As UltraWinGrid.UltraGridRow In grdSOTALLOX.Rows
            If grow.Cells("ALLO_CTL_NO").Value = ALLO_CTL_NO Then
                grow.Cells("SELECTED").Value = "1"
                grow.Update()
                Click_Command("Edit")
                tabSOTALLOX.SelectedTab = tabSOTALLOX.Tabs("Status by Item")
                tabSOTALLOX.Tabs("Status by Item").Tag = ALLO_CTL_NO
            End If
        Next
    End Sub
    Private Sub grdSOTALLO1S_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTALLO1S.DoubleClickRow
        If e.Row.Band.Key = "SOTALLO1S" Then
            Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Value
            Dim ALLO_CTL_NO As String = e.Row.Cells("ALLO_CTL_NO").Value
            Edit_Allocation_SOTALLO1S(ITEM_CODE, ALLO_CTL_NO)
        ElseIf e.Row.Band.Key = "SOTALLO1S_SOTALLO2S" Then
            Dim ITEM_CODE As String = e.Row.ParentRow.Cells("ITEM_CODE").Value
            Dim ALLO_CTL_NO As String = e.Row.ParentRow.Cells("ALLO_CTL_NO").Value
            Edit_Allocation_SOTALLO1S(ITEM_CODE, ALLO_CTL_NO)
        End If
    End Sub

    Public Overrides Function CustomSummary_End(
        ByVal summarySettings As UltraWinGrid.SummarySettings,
        ByVal rows As UltraWinGrid.RowsCollection,
        ByVal CustomValue As Double,
        ByVal grd As UltraWinGrid.UltraGrid) As Double

        CustomValue = 0
        Dim TOTALS As New Dictionary(Of String, Decimal)

        '.Item("TYRBVD_P" & Format(I, "00")).Expression = Replace("ISNULL(TY_P00,0) - ISNULL(RB_P00,0)", "P00", "P" & Format(I, "00"))
        '.Item("TYRBVP_P" & Format(I, "00")).Expression = Replace("IIF(ISNULL(RB_P00,0)=0,0,100*ISNULL(TYRBVD_P00,0)/ISNULL(RB_P00,0))", "P00", "P" & Format(I, "00"))
        '.Item("TYWBVP_P" & Format(I, "00")).Expression = Replace("IIF(ISNULL(WB_P00,0)=0,0,100*(ISNULL(TY_P00,0)-ISNULL(WB_P00,0))/ISNULL(WB_P00,0))", "P00", "P" & Format(I, "00"))
        '.Item("TYLYVP_P" & Format(I, "00")).Expression = Replace("IIF(ISNULL(LY_P00,0)=0,0,100*(ISNULL(TY_P00,0)-ISNULL(LY_P00,0))/ISNULL(LY_P00,0))", "P00", "P" & Format(I, "00"))


        Select Case grd.Name
            Case "grdSOTALLO1S"
                Dim KEY As String = summarySettings.Key
                If KEY.StartsWith("TYRBVP") Then
                    Dim RB As String = "RB" & Mid(KEY, 7)
                    Dim D As String = "TYRBVD" & Mid(KEY, 7)
                    TOTALS.Add(RB, 0)
                    TOTALS.Add(D, 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS(RB) <> 0 Then CustomValue = 100 * TOTALS(D) / TOTALS(RB)

                ElseIf KEY = "QTY_ALLO_USED" Then
                    TOTALS.Add("QTY_ALLO", 0)
                    TOTALS.Add("ORDR_QTY_SHIP", 0)
                    TOTALS.Add("ORDR_QTY_PICK", 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS("QTY_ALLO") <> 0 Then CustomValue = 100 * (TOTALS("ORDR_QTY_SHIP") + TOTALS("ORDR_QTY_PICK")) / TOTALS("QTY_ALLO")

                ElseIf KEY = "TRADE_CLASS_CODE" Then
                    Stop
                End If

            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Public Overrides Function CustomStringSummary_End(
        ByVal summarySettings As UltraWinGrid.SummarySettings,
        ByVal rows As UltraWinGrid.RowsCollection,
        ByVal CustomValue As String,
        ByVal grd As UltraWinGrid.UltraGrid) As String

        Select Case grd.Name
            Case "grdSOTALLO1S"
                Dim KEY As String = summarySettings.Key
                CustomValue = "Totals"
            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Sub CustomSummary_Calculate_Totals(
       ByVal rows As UltraWinGrid.RowsCollection,
       ByRef TOTALS As Dictionary(Of String, Decimal),
       ByVal KEY As String)

        For Each grow2 As UltraWinGrid.UltraGridRow In rows
            If grow2.IsGroupByRow Then
                Dim gbrow As UltraWinGrid.UltraGridGroupByRow = DirectCast(grow2, UltraWinGrid.UltraGridGroupByRow)
                CustomSummary_Calculate_Totals(gbrow.Rows, TOTALS, KEY)
            Else
                If KEY.StartsWith("TYRBVP") Then
                    Dim RB As String = "RB" & Mid(KEY, 7)
                    Dim D As String = "TYRBVD" & Mid(KEY, 7)
                    TOTALS(RB) += Val(grow2.Cells(RB).Value & "")
                    TOTALS(D) += Val(grow2.Cells(D).Value & "")
                ElseIf KEY = "QTY_ALLO_USED" Then
                    TOTALS("QTY_ALLO") += Val(grow2.Cells("QTY_ALLO").Value & "")
                    TOTALS("ORDR_QTY_SHIP") += Val(grow2.Cells("ORDR_QTY_SHIP").Value & "")
                    TOTALS("ORDR_QTY_PICK") += Val(grow2.Cells("ORDR_QTY_PICK").Value & "")
                ElseIf KEY = "TRADE_CLASS_CODE" Then
                    '  TOTALS(KEY) = "Totals"
                End If
            End If
        Next
    End Sub

    Private Sub grdSOTALLO1S_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSOTALLO1S.InitializeRow
        If e.Row.IsDataRow Then
            If e.Row.Band.Key = "SOTALLO1S_SOTALLO2S" Then
                If Val(e.Row.Cells("QTY_ALLO_USED").Value & "") > 100 Then
                    e.Row.Cells("QTY_ALLO_USED").Appearance.ForeColor = Drawing.Color.Red
                Else
                    e.Row.Cells("QTY_ALLO_USED").Appearance.ForeColor = Drawing.Color.Empty
                End If
            End If
        End If

    End Sub

    Sub Setup_Tree()

        Application.DoEvents()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Collections into Selection Tree")

        COLUMN_NAMEs.Clear()
        COLUMN_NAMEs.Add("TRADE_CLASS_CODE")
        COLUMN_NAMEs.Add("CUST_CODE")
        COLUMN_NAMEs.Add("CUST_STORE_NO")

        ReDim COLUMN_NAME_by_Lvl(COLUMN_NAMEs.Count)
        ReDim COLUMN_CAPTION_by_Lvl(COLUMN_NAMEs.Count)
        ReDim G_by_Lvl(COLUMN_NAMEs.Count)
        ReDim SCOPE(COLUMN_NAMEs.Count)
        For G As Integer = 1 To COLUMN_NAMEs.Count
            Dim Lvl As Integer = G
            COLUMN_NAME_by_Lvl(Lvl) = COLUMN_NAMEs(G - 1)

            Dim COLUMN_CAPTION As String = ""
            Select Case Lvl
                Case 1
                    COLUMN_CAPTION = "Trade Class"
                Case 2
                    COLUMN_CAPTION = "Customer"
                Case 3
                    COLUMN_CAPTION = "Store"
            End Select
            COLUMN_CAPTION_by_Lvl(Lvl) = COLUMN_CAPTION
            G_by_Lvl(Lvl) = G
        Next

        With tvwDQ
            Dim rootColumnSet As UltraWinTree.UltraTreeColumnSet = .ColumnSettings.RootColumnSet
            rootColumnSet.Columns.Clear()
            For Lvl As Integer = 1 To COLUMN_NAMEs.Count
                Dim column As UltraWinTree.UltraTreeNodeColumn = rootColumnSet.Columns.Add(COLUMN_NAME_by_Lvl(Lvl))
            Next
        End With

        Dim COLUMN_NAMEs_ordered As String = ""
        Dim CODE_COLUMNs_ordered As String = ""
        For Lvl As Integer = 1 To COLUMN_NAMEs.Count
            COLUMN_NAMEs_ordered &= "," & COLUMN_NAME_by_Lvl(Lvl)
            CODE_COLUMNs_ordered &= ",CODE" & CStr(G_by_Lvl(Lvl))
        Next
        COLUMN_NAMEs_ordered = Mid(COLUMN_NAMEs_ordered, 2)
        CODE_COLUMNs_ordered = Mid(CODE_COLUMNs_ordered, 2)

        Dim aNode As New Infragistics.Win.UltraWinTree.UltraTreeNode
        Dim CODE_VALUE_at_Lvl() As String = Nothing
        ReDim CODE_VALUE_at_Lvl(COLUMN_NAMEs.Count)

        Dim IMAGE_FOLDER As String = ASCMAIN1.Folders("Images") & "ABS\Menu\Tree\"

        tvwDQ.Nodes.Clear()

        Dim cur_Node_at_Lvl() As Infragistics.Win.UltraWinTree.UltraTreeNode
        ReDim cur_Node_at_Lvl(COLUMN_NAMEs.Count)
        If COLUMN_CAPTION_by_Lvl.Length = 1 Then
            aNode = tvwDQ.Nodes.Add("*", "All")
        Else
            'aNode = tvwDQ.Nodes.Add("*", "All (" & COLUMN_CAPTION_by_Lvl(1) & ")")
            aNode = tvwDQ.Nodes.Add("*", "All " & COLUMN_CAPTION_by_Lvl(1) & "s")
        End If

        cur_Node_at_Lvl(0) = aNode

        ASCMAIN1.sql = "Select Distinct " & CODE_COLUMNs_ordered & " from (" & vbCrLf _
            & "Select ARTCUST1.TRADE_CLASS_CODE CODE1, ARTCUST1.CUST_CODE CODE2, NULL CODE3" & vbCrLf _
            & " from ARTCUST1 where ARTCUST1.CUST_CODE in (Select Distinct CUST_CODE from " & SOTALLO2S & ")" & vbCrLf _
            & "  and NVL(ARTCUST1.CUST_ALLOCATE_BY_STORE,'0') = '0'" & vbCrLf _
            & " union " & vbCrLf _
            & "Select ARTCUST1.TRADE_CLASS_CODE CODE1, ARTCUST2.CUST_CODE CODE2, ARTCUST2.CUST_STORE_NO CODE3" & vbCrLf _
            & " from ARTCUST1,ARTCUST2 where ARTCUST1.CUST_CODE in (Select Distinct CUST_CODE from " & SOTALLO2S & ")" & vbCrLf _
            & "  and NVL(ARTCUST1.CUST_ALLOCATE_BY_STORE,'0') = '1'" & vbCrLf _
            & "  and ARTCUST2.CUST_CODE = ARTCUST1.CUST_CODE" & vbCrLf _
            & ")"
        Dim TBL As DataTable = ASCDATA1.GetDataTable
        Dim last_level_set As Integer = 0

        Dim show_codes As Boolean = False
        'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Codes"), UltraWinToolbars.StateButtonTool)
        show_codes = False ' tlb_sbt.Checked

        Dim images As New Dictionary(Of String, System.Drawing.Bitmap)
        images.Add("LEAF", ASCMAIN1.Get_Image(IMAGE_FOLDER, "ITEM_green"))
        images.Add("M", ASCMAIN1.Get_Image(IMAGE_FOLDER, "M"))
        images.Add("M_OPEN", ASCMAIN1.Get_Image(IMAGE_FOLDER, "M_OPEN"))

        For Each row As DataRow In TBL.Select("", CODE_COLUMNs_ordered)
            For Lvl As Integer = 1 To COLUMN_NAMEs.Count '- 1
                If CODE_VALUE_at_Lvl(Lvl) <> row.Item(Lvl - 1) & "" Or last_level_set < Lvl Then
                    last_level_set = Lvl

                    If Lvl > 1 And row.Item(Lvl - 1) & "" = "" Then
                        ' DO NOTHING
                    Else
                        If Lvl = 1 Then
                            aNode = tvwDQ.Nodes.Add
                        Else
                            aNode = cur_Node_at_Lvl(Lvl - 1).Nodes.Add
                        End If

                        cur_Node_at_Lvl(Lvl) = aNode

                        If Lvl = COLUMN_NAMEs.Count Then
                            ' Dim KEY As String = row.Item("ITEM_CATGY_CODE") & "/" & row.Item("COLLECTION_CODE") & "/" & row.Item("ITEM_CLASS_CODE")
                            ' aNode.Key = KEY
                            ' IF WE EVER EXPAND UPON WHAT COLUMNS TO PLACE INTO THE KEY, WE NEED TO ALSO LOOK AT TXTFINDITEM_CODE
                        End If

                        Dim CAPTION As String = "?"
                        'Dim COLUMN_NAME_CODE As String = COLUMN_NAME_by_Lvl(Lvl) ' Gs(Lvl - 1)
                        'Dim rowSATANALC As DataRow = dst.Tables("SATANALC").Rows.Find(COLUMN_NAME_CODE)
                        'If rowSATANALC Is Nothing Then
                        '    CAPTION = "?"
                        'Else
                        '    Dim COLUMN_NAME_DESC As String = rowSATANALC.Item("COLUMN_NAME_DESC")
                        '    Dim TABLE_NAME_LOOKUP As String = rowSATANALC.Item("TABLE_NAME_LOOKUP")
                        '    CAPTION = LookUp(TABLE_NAME_LOOKUP, row.Item(Lvl - 1) & "", True).Item(COLUMN_NAME_DESC) & ""
                        '    If CAPTION = "" Then
                        '        CAPTION = "?"
                        '    End If
                        'End If

                        Dim CODE_VALUE As String = row.Item(Lvl - 1)
                        '  If Lvl = 3 Then CODE_VALUE = row.Item(Lvl - 2) & ":" & row.Item(Lvl - 1)
                        ' WILL NEED THIS EVENTUALLY, SINCE STORE NOS ARE NOT PREFIXED BY CUST_CODE RIGHT NOW IN SATANALD
                        CAPTION = Get_Description(COLUMN_NAME_by_Lvl(Lvl), CODE_VALUE)
                        'If Lvl = 3 Then Stop

                        If show_codes Then
                            aNode.Text = row.Item(Lvl - 1) & ":" & CAPTION
                        Else
                            aNode.Text = CAPTION
                        End If

                        aNode.Tag = row.Item(Lvl - 1) & ":" & CAPTION
                        aNode.Expanded = False

                        CODE_VALUE_at_Lvl(Lvl) = row.Item(Lvl - 1) & ""
                        If (last_level_set = COLUMN_NAMEs.Count - 1 And row.Item(COLUMN_NAMEs.Count - 1) & "" = "") _
                        Or (last_level_set = COLUMN_NAMEs.Count) Then
                            aNode.LeftImages.Add(images("LEAF"))
                        Else
                            aNode.Override.NodeAppearance.Image = images("M")
                            aNode.Override.ExpandedNodeAppearance.Image = images("M_OPEN")
                        End If

                        For iLvl As Integer = 1 To Lvl
                            aNode.Cells(iLvl - 1).Value = CODE_VALUE_at_Lvl(iLvl)
                        Next
                    End If

                End If
            Next
        Next

        'Dim rows() As DataRow = dst.Tables("SATANALR").Select("SEL1 = '1' and SEL2 = '1' and SEL = '1'", "SEQ")
        'grdSATANALR.Tag = rows(0)("DATA_CODE1") & ":" & rows(0)("DATA_CODE2")


        'Setup_View()
        'Setup_tabDetails()

        'If tvwDQ.Nodes.Count > 0 Then
        '    tvwDQ.ActiveNode = tvwDQ.Nodes(0)
        '    tvwDQ.Nodes(0).Selected = True
        '    Click_Node(tvwDQ.Nodes(0))
        '    SortGrid("CODES", False)
        '    grdSATANALR.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        '    Setup_Layout_Option()
        'End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

#Region "tvwDQ"
    Private Sub tvwDQ_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tvwDQ.Click

        Try
            Dim xx As System.Windows.Forms.MouseEventArgs = DirectCast(e, System.Windows.Forms.MouseEventArgs)
            Dim tt As UltraWinTree.UltraTree = DirectCast(sender, UltraWinTree.UltraTree)
            Dim tnode As UltraWinTree.UltraTreeNode = tt.GetNodeFromPoint(xx.X, xx.Y)

            If tnode IsNot Nothing Then
                Click_Node(tvwDQ.ActiveNode)
                tvwDQ.SelectedNodes.Clear()
                tvwDQ.ActiveNode.Selected = True
            End If


        Catch ex As Exception

        End Try

    End Sub
#End Region

    Sub Click_Node(ByVal tnode As UltraWinTree.UltraTreeNode)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Summary")

        If tnode IsNot Nothing Then
            Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show All Levels"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = False
            LVL = tnode.Level + 1
            If tnode.Key = "*" Then
                LVL = 0
            End If
            Dim COLS_Select As String = ""
            Dim COLS_Group_By As String = ""
            Dim COLS_Order_By As String = ""
            Dim sqlW As String = ""
            Dim CAPTION As String = ""


            ' grdSOTALLOCS.DisplayLayout.Bands(0).Columns("DESCRIPTION").Header.Caption = COLUMN_CAPTION_by_Lvl(LVL + 1)

            Dim sqlA As String = ""

            For G As Integer = 1 To COLUMN_NAME_by_Lvl.Count - 1
                If G <= LVL Then

                    Dim COLUMN_NAME As String = COLUMN_NAME_by_Lvl(G)
                    If tnode.Cells(G - 1).Text = "" Then
                        sqlA &= " and " & COLUMN_NAME & " is Null"
                    Else
                        sqlA &= " and " & COLUMN_NAME & " = '" & tnode.Cells(G - 1).Text & "'"
                    End If

                    Dim CODE_VALUE As String = tnode.Cells(G - 1).Text
                    Dim DESC_VALUE As String = Get_Description(COLUMN_NAME_by_Lvl(G), CODE_VALUE)
                    CAPTION &= ", " & COLUMN_CAPTION_by_Lvl(G) & " " & CODE_VALUE & ":" & DESC_VALUE
                End If
            Next

            If LVL = COLUMN_NAMEs.Count Then
                ASCMAIN1.sql = "Select ALLO_CTL_NO" & vbCrLf _
                    & ", SUM (QTY_ALLO) QTY_ALLO" & vbCrLf _
                    & ", SUM (ORDR_QTY) ORDR_QTY" _
                    & ", SUM (ORDR_QTY_OPEN) ORDR_QTY_OPEN" _
                    & ", SUM (ORDR_QTY_PICK) ORDR_QTY_PICK" _
                    & ", SUM (ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
                    & ", SUM (ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
                    & ", SUM (CASE WHEN NVL(LY_QTY_SELL_IN1,0) <> 0 THEN NVL(LY_QTY_SELL_IN1,0) ELSE NVL(LY_QTY_SELL_IN2,0) END) LY_QTY_SELL_IN" & vbCrLf _
                    & ", SUM (CASE WHEN NVL(LY_QTY_SELL_THRU1,0) <> 0 THEN NVL(LY_QTY_SELL_THRU1,0) ELSE NVL(LY_QTY_SELL_THRU2,0) END) LY_QTY_SELL_THRU" & vbCrLf _
                    & " from " & SOTALLO3S & vbCrLf _
                    & IIf(sqlA = "", "", ASCMAIN1.SQL_Add_WHERE(sqlA)) & vbCrLf _
                    & " group by ALLO_CTL_NO"
            Else
                ASCMAIN1.sql = "Select ALLO_CTL_NO" & vbCrLf _
                    & ", SUM (QTY_ALLO) QTY_ALLO" & vbCrLf _
                    & ", SUM (ORDR_QTY) ORDR_QTY" _
                    & ", SUM (ORDR_QTY_OPEN) ORDR_QTY_OPEN" _
                    & ", SUM (ORDR_QTY_PICK) ORDR_QTY_PICK" _
                    & ", SUM (ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
                    & ", SUM (ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
                    & ", SUM (CASE WHEN NVL(LY_QTY_SELL_IN1,0) <> 0 THEN NVL(LY_QTY_SELL_IN1,0) ELSE NVL(LY_QTY_SELL_IN2,0) END) LY_QTY_SELL_IN" & vbCrLf _
                    & ", SUM (CASE WHEN NVL(LY_QTY_SELL_THRU1,0) <> 0 THEN NVL(LY_QTY_SELL_THRU1,0) ELSE NVL(LY_QTY_SELL_THRU2,0) END) LY_QTY_SELL_THRU" & vbCrLf _
                    & " from " & SOTALLO2S & vbCrLf _
                    & IIf(sqlA = "", "", ASCMAIN1.SQL_Add_WHERE(sqlA)) & vbCrLf _
                    & " group by ALLO_CTL_NO"
            End If

            If LVL = 0 Then
                grdSOTALLOCS.Text = "All" ' tvwDQ.Nodes(0).Text
            Else
                grdSOTALLOCS.Text = Mid(CAPTION, 3)
            End If

            ASCMAIN1.sql = "Select SOTALLO1S.*" & vbCrLf _
                & ", X.QTY_ALLO" & vbCrLf _
                & ", X.ORDR_QTY, X.ORDR_QTY_OPEN, X.ORDR_QTY_PICK, X.ORDR_QTY_SHIP, X.ORDR_QTY_CANC" & vbCrLf _
                & ", X.LY_QTY_SELL_IN, X.LY_QTY_SELL_THRU" & vbCrLf _
                & " from " & SOTALLO1S & " SOTALLO1S, (" & ASCMAIN1.sql & ") X" & vbCrLf _
                & " where SOTALLO1S.ALLO_CTL_NO = X.ALLO_CTL_NO"
            Fill_Records("SOTALLOCS", "", True, ASCMAIN1.sql)

            grdSOTALLOCS.DisplayLayout.GroupByBox.Hidden = False
            grdSOTALLOCS.DisplayLayout.Bands(0).SortedColumns.Clear()
            For Each COLUMN_NAME As String In New String() {"BRAND_CODE", "COLLECTION_CODE", "PROD_CODE"}
                '                grdSOTALLOCS.DisplayLayout.Bands(0).Columns(COLUMN_NAME).HiddenWhenGroupBy = DefaultableBoolean.True
                grdSOTALLOCS.DisplayLayout.Bands(0).SortedColumns.Add(COLUMN_NAME, False, True)
            Next
            grdSOTALLOCS.DisplayLayout.Bands(0).SortedColumns.Add("ITEM_CODE", False, False)
            grdSOTALLOCS.DisplayLayout.GroupByBox.Hidden = True
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Function Get_Description(
    ByVal COLUMN_NAME As String,
    ByVal CODE_VALUE As String,
    Optional ByVal use_code_as_default_value As Boolean = False)
        Dim DESC_VALUE As String = IIf(use_code_as_default_value, CODE_VALUE, "")
        Dim rowSATANALD As DataRow = dst.Tables("SATANALD").Rows.Find _
                       (New String() {COLUMN_NAME, CODE_VALUE})
        If rowSATANALD IsNot Nothing Then
            DESC_VALUE = rowSATANALD.Item("DESC_VALUE") & ""
        Else
            DESC_VALUE = "?"
        End If

        Return DESC_VALUE
    End Function

    Private Sub grdSOTALLOCS_BeforeExitEditMode(sender As Object, e As UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTALLOCS.BeforeExitEditMode

    End Sub

    Private Sub grdSOTALLOCS_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTALLOCS.DoubleClickRow
        If e.Row.IsDataRow Then

            If e.Row.Band.Key = "SOTALLOCS" Then
                Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Value
                Dim ALLO_CTL_NO As String = e.Row.Cells("ALLO_CTL_NO").Value
                Edit_Allocation_SOTALLOCS(ITEM_CODE, ALLO_CTL_NO)
                'ElseIf e.Row.Band.Key = "SOTALLO1S_SOTALLO2S" Then
                '    Dim ITEM_CODE As String = e.Row.ParentRow.Cells("ITEM_CODE").Value
                '    Dim ALLO_CTL_NO As String = e.Row.ParentRow.Cells("ALLO_CTL_NO").Value
                '    Edit_Allocation_SOTALLOCS(ITEM_CODE, ALLO_CTL_NO)
            End If

        End If

    End Sub

    Sub Edit_Allocation_SOTALLOCS(ITEM_CODE As String, ALLO_CTL_NO As String)
        tabSOTALLOX.SelectedTab = tabSOTALLOX.Tabs("Items")
        Absx1.txtFor("ITEM_CODE").Text = ITEM_CODE
        Find_ITEM_CODE()
        tabSOTALLOX.SelectedTab = tabSOTALLOX.Tabs("Allocations")
        For Each grow As UltraWinGrid.UltraGridRow In grdSOTALLOX.Rows
            If grow.Cells("ALLO_CTL_NO").Value = ALLO_CTL_NO Then
                grow.Cells("SELECTED").Value = "1"
                grow.Update()
                Click_Command("Edit")
                tabSOTALLOX.SelectedTab = tabSOTALLOX.Tabs("Status by Customer")
                tabSOTALLOX.Tabs("Status by Customer").Tag = ALLO_CTL_NO
            End If
        Next
    End Sub

    Private Sub grdSOTALLOS_BeforeExitEditMode(sender As Object, e As UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTALLOS.BeforeExitEditMode

        If grdSOTALLOS.ActiveCell IsNot Nothing Then
            With grdSOTALLOS.ActiveCell
                If .Column.Key.StartsWith("ALLO_") And Not .Column.Key.StartsWith("ALLO_NOTES") Then
                    Dim QTY_ALLO As Integer = Val(.EditorResolved.Value & "")
                    Dim i As Integer = Val(Mid(.Column.Key, 6))
                    Dim ALLO_CTL_NO As String = ALLO_CTL_NOi(i)
                    Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO)
                    Dim ITEM_CODE As String = rowSOTALLO1.Item("ITEM_CODE")
                    Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)

                    Dim ITEM_SO_QTY_MULT As Int32 = Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & String.Empty)
                    If ITEM_SO_QTY_MULT <> 0 AndAlso QTY_ALLO Mod ITEM_SO_QTY_MULT <> 0 Then
                        QTY_ALLO += (ITEM_SO_QTY_MULT - (QTY_ALLO Mod ITEM_SO_QTY_MULT))
                        .EditorResolved.Value = QTY_ALLO
                    End If
                End If
            End With
        End If
    End Sub

    Private Sub grdSOTALLOS_DoubleClickCell(sender As Object, e As UltraWinGrid.DoubleClickCellEventArgs) Handles grdSOTALLOS.DoubleClickCell
        If New String() {"SI_PLAN", "ST_PLAN", "SI_HIST", "ST_HIST", "ALLO_SPREAD"}.Contains(e.Cell.Column.Key) Then
            If MsgBox("OK to " & Mid(e.Cell.ToolTipText, "Double-Click to ".Length + 1), MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Copying " & Mid(e.Cell.ToolTipText, "Double-Click to ".Length + 1))
                For Each grow As UltraWinGrid.UltraGridRow In grdSOTALLOS.Rows
                    grow.Cells("ALLO_" & Format(SINGLE_ITEM_ictr, "00")).Value = grow.Cells(e.Cell.Column.Key).Value
                    grow.Update()
                Next
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
            End If
        End If
    End Sub

    Private Sub grdSOTALLOS_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSOTALLOS.InitializeRow
        e.Row.Cells("SI_PLAN").ToolTipText = "Double-Click to Copy All Sell-In Plan Qtys to Allocation"
        e.Row.Cells("ST_PLAN").ToolTipText = "Double-Click to Copy All Sell-Thru Plan Qtys to Allocation"
        e.Row.Cells("SI_HIST").ToolTipText = "Double-Click to Copy All Sell-In History Qtys to Allocation"
        e.Row.Cells("ST_HIST").ToolTipText = "Double-Click to Copy All Sell-Thru History Qtys to Allocation"
        e.Row.Cells("ALLO_SPREAD").ToolTipText = "Double-Click to Copy All Retail%Plan Qtys to Allocation"
    End Sub

    Private Sub grdSOTALLO1_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSOTALLO1.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdSOTALLO1, sql_where, , , "ITEM_CODE")
    End Sub

    Sub Get_HC_Data()

        Dim rowSOTALLOC As DataRow = dst.Tables("SOTALLOC").Rows.Find("IPLBAE")

        Dim COLLECTION_CODEs As New List(Of String)
        Dim HC_CODEs As New List(Of String)
        Dim rowICTITEM1 As DataRow
        For Each rowICTITEM1 In dst.Tables("ICTITEM1").Select("")
            Dim COLLECTION_CODE As String = rowICTITEM1.Item("COLLECTION_CODE")
            If Not COLLECTION_CODEs.Contains(COLLECTION_CODE) Then
                COLLECTION_CODEs.Add(COLLECTION_CODE)
                Dim rowICTCOLL1 As DataRow = LookUp("ICTCOLL1", COLLECTION_CODE)
                Dim HC_CODE As String = rowICTCOLL1.Item("HC_CODE")
                If Not HC_CODEs.Contains(HC_CODE) Then
                    HC_CODEs.Add(HC_CODE)
                End If
            End If
        Next

        Dim ITEM_CODE As String = grdSOTALLO1.ActiveRow.Cells("ITEM_CODE").Value
        rowICTITEM1 = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
        Dim ITEM_SO_QTY_MULT As Integer = Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT"))
        If ITEM_SO_QTY_MULT <= 0 Then ITEM_SO_QTY_MULT = 1

        'ASCMAIN1.sql = "Select DECODE(ARTCUSF2.CUST_CODE,NULL,ARTCUST2.SELL_CODE,ARTCUSF2.SELL_CODE) SELL_CODE" & vbCrLf _
        '    & ", Sum (RSTRETL2.RETAIL_SALES) RETAIL_SALES, Min (ARTCUST2.CUST_CODE || '-' || ARTCUST2.CUST_STORE_NO) CS" & vbCrLf _
        '    & " from RSTRETL2,ARTCUST2,ICTCOLL1," & ARTCUSF2 & " ARTCUSF2" & vbCrLf _
        '    & " where ARTCUST2.CUST_CODE = RSTRETL2.CUST_CODE" & vbCrLf _
        '    & "   and ARTCUST2.CUST_STORE_NO = RSTRETL2.CUST_STORE_NO" & vbCrLf _
        '    & "   and ICTCOLL1.COLLECTION_CODE = RSTRETL2.COLLECTION_CODE" & vbCrLf _
        '    & "   and RSTRETL2.OPS_YYYYPP between '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -5) & "' and '" & ASCMAIN1.CYP & "'" & vbCrLf _
        '    & "   and ICTCOLL1.HC_CODE in ('" & Join(HC_CODEs.ToArray, "','") & "')" & vbCrLf _
        '    & "   and DECODE(ARTCUSF2.CUST_CODE,NULL,ARTCUST2.SELL_CODE,ARTCUSF2.SELL_CODE) is Not Null" & vbCrLf _
        '    & sqlwhere_F _
        '    & "   and ARTCUSF2.CUST_CODE (+) = ARTCUST2.CUST_CODE" & vbCrLf _
        '    & "   and ARTCUSF2.CUST_STORE_NO (+) = ARTCUST2.CUST_STORE_NO" & vbCrLf _
        '    & " group by DECODE(ARTCUSF2.CUST_CODE,NULL,ARTCUST2.SELL_CODE,ARTCUSF2.SELL_CODE)"

        ASCMAIN1.sql = "Select DECODE(ARTCUSF2.CUST_CODE,NULL,NVL(ARTCUST2.SELL_CODE_AC,ARTCUST2.SELL_CODE),NVL(ARTCUSF2.SELL_CODE_AC,ARTCUSF2.SELL_CODE)) SELL_CODE" & vbCrLf _
            & ", Sum (RSTRETL2.RETAIL_SALES) RETAIL_SALES, Min (ARTCUST2.CUST_CODE || '-' || ARTCUST2.CUST_STORE_NO) CS" & vbCrLf _
            & " from RSTRETL2,ARTCUST2,ICTCOLL1," & ARTCUSF2 & " ARTCUSF2" & vbCrLf _
            & " where ARTCUST2.CUST_CODE = RSTRETL2.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_STORE_NO = RSTRETL2.CUST_STORE_NO" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = RSTRETL2.COLLECTION_CODE" & vbCrLf _
            & "   and RSTRETL2.OPS_YYYYPP between '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -5) & "' and '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "   and ICTCOLL1.HC_CODE in ('" & Join(HC_CODEs.ToArray, "','") & "')" & vbCrLf _
            & "   and DECODE(ARTCUSF2.CUST_CODE,NULL,ARTCUST2.SELL_CODE,ARTCUSF2.SELL_CODE) is Not Null" & vbCrLf _
            & sqlwhere_F _
            & "   and ARTCUSF2.CUST_CODE (+) = ARTCUST2.CUST_CODE" & vbCrLf _
            & "   and ARTCUSF2.CUST_STORE_NO (+) = ARTCUST2.CUST_STORE_NO" & vbCrLf _
            & " group by DECODE(ARTCUSF2.CUST_CODE,NULL,NVL(ARTCUST2.SELL_CODE_AC,ARTCUST2.SELL_CODE),NVL(ARTCUSF2.SELL_CODE_AC,ARTCUSF2.SELL_CODE))"

        Dim RETAIL_SALES_TOTAL As Decimal = 0
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim SELL_CODE As String = row.Item("SELL_CODE") & ""
            Dim RETAIL_SALES As Decimal = Val(row.Item("RETAIL_SALES") & "")
            RETAIL_SALES_TOTAL += RETAIL_SALES

            Dim rowSOTALLOS As DataRow = dst.Tables("SOTALLOS").Rows.Find(New String() {"IPLBAE", "000" & SELL_CODE})
            If rowSOTALLOS IsNot Nothing Then
                rowSOTALLOS.Item("RETAIL_SALES") = RETAIL_SALES
            Else
                'MsgBox("Warning - some stores with retail sales did not have an AE associated")
            End If
            rowSOTALLOC.Item("RETAIL_SALES") = Val(rowSOTALLOC.Item("RETAIL_SALES") & "") + RETAIL_SALES
        Next

        dst.Tables("SOTALLOS").Columns("RETAIL_SALES_PCT").Expression = IIf(RETAIL_SALES_TOTAL = 0, "0", "100 * ISNULL(RETAIL_SALES,0) / " & CStr(RETAIL_SALES_TOTAL))
        'dst.Tables("SOTALLOS").Columns("ALLO_SPREAD").Expression = "ISNULL(RETAIL_SALES_PCT,0) * PARENT.QTY_ALLO / 100"

        Dim ALLO_QTY As Int64 = Val(rowSOTALLOC.Item("ALLO_" & Format(SINGLE_ITEM_ictr, "00")) & "")
        Dim T As Int64 = 0
        For Each row As DataRow In dst.Tables("SOTALLOS").Select("")
            Dim RETAIL_SALES_PCT As Decimal = Val(row.Item("RETAIL_SALES_PCT") & "")
            Dim ALLO_SPREAD As Int64 = System.Math.Round(ALLO_QTY * RETAIL_SALES_PCT / 100, 0)
            If ALLO_SPREAD < 0 Then ALLO_SPREAD = 0
            If ITEM_SO_QTY_MULT > 1 Then
                ALLO_SPREAD = Math.Floor(ALLO_SPREAD / ITEM_SO_QTY_MULT) * ITEM_SO_QTY_MULT
            End If
            row.Item("ALLO_SPREAD") = ALLO_SPREAD
            T += ALLO_SPREAD
        Next
        'dst.Tables("SOTALLOS").Columns("ALLO_SPREAD").Expression = "ISNULL(RETAIL_SALES_PCT,0)  * " & CStr(Val(rowSOTALLOC.Item("ALLO_01") & "")) & " / 100"

        If T <> ALLO_QTY And Math.Abs(T - ALLO_QTY) > ITEM_SO_QTY_MULT Then
            For Each ROW As DataRow In dst.Tables("SOTALLOS").Select("ALLO_SPREAD <> 0", "ALLO_SPREAD DESC")
                Dim ALLO_SPREAD As Int64 = Val(ROW.Item("ALLO_SPREAD") & "")
                ALLO_SPREAD += System.Math.Sign(ALLO_QTY - T) * ITEM_SO_QTY_MULT
                ROW.Item("ALLO_SPREAD") = ALLO_SPREAD
                T += System.Math.Sign(ALLO_QTY - T) * ITEM_SO_QTY_MULT
                If T >= ALLO_QTY Or Math.Abs(T - ALLO_QTY) < ITEM_SO_QTY_MULT Then
                    Exit For
                End If
            Next
        End If



        'ASCMAIN1.sql = "Select DECODE(ARTCUSF2.CUST_CODE,NULL,ARTCUST2.SELL_CODE,ARTCUSF2.SELL_CODE) SELL_CODE" & vbCrLf _
        '    & ", Count (*) STORES" & vbCrLf _
        '    & " from SATAUTH1,ARTCUST2," & ARTCUSF2 & " ARTCUSF2" & vbCrLf _
        '    & " where ARTCUST2.CUST_CODE = SATAUTH1.CUST_CODE" & vbCrLf _
        '    & "   and ARTCUST2.CUST_STORE_NO = SATAUTH1.CUST_STORE_NO" & vbCrLf _
        '    & "   and SATAUTH1.OPS_YYYYPP_OPENED <= '" & ASCMAIN1.CYP & "'" & vbCrLf _
        '    & "   and SATAUTH1.OPS_YYYYPP_CLOSED is Null" & vbCrLf _
        '    & "   and SATAUTH1.HC_CODE in ('" & Join(HC_CODEs.ToArray, "','") & "')" & vbCrLf _
        '    & sqlwhere_F _
        '    & "   and ARTCUSF2.CUST_CODE (+) = ARTCUST2.CUST_CODE" & vbCrLf _
        '    & "   and ARTCUSF2.CUST_STORE_NO (+) = ARTCUST2.CUST_STORE_NO" & vbCrLf _
        '    & " group by DECODE(ARTCUSF2.CUST_CODE,NULL,ARTCUST2.SELL_CODE,ARTCUSF2.SELL_CODE)"

        ASCMAIN1.sql = "Select DECODE(ARTCUSF2.CUST_CODE,NULL,NVL(ARTCUST2.SELL_CODE_AC,ARTCUST2.SELL_CODE),NVL(ARTCUSF2.SELL_CODE_AC,ARTCUSF2.SELL_CODE)) SELL_CODE" & vbCrLf _
            & ", Count (*) STORES" & vbCrLf _
            & " from SATAUTH1,ARTCUST2," & ARTCUSF2 & " ARTCUSF2" & vbCrLf _
            & " where ARTCUST2.CUST_CODE = SATAUTH1.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_STORE_NO = SATAUTH1.CUST_STORE_NO" & vbCrLf _
            & "   and SATAUTH1.OPS_YYYYPP_OPENED <= '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "   and SATAUTH1.OPS_YYYYPP_CLOSED is Null" & vbCrLf _
            & "   and SATAUTH1.HC_CODE in ('" & Join(HC_CODEs.ToArray, "','") & "')" & vbCrLf _
            & sqlwhere_F _
            & "   and ARTCUSF2.CUST_CODE (+) = ARTCUST2.CUST_CODE" & vbCrLf _
            & "   and ARTCUSF2.CUST_STORE_NO (+) = ARTCUST2.CUST_STORE_NO" & vbCrLf _
            & " group by DECODE(ARTCUSF2.CUST_CODE,NULL,NVL(ARTCUST2.SELL_CODE_AC,ARTCUST2.SELL_CODE),NVL(ARTCUSF2.SELL_CODE_AC,ARTCUSF2.SELL_CODE))"


        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim SELL_CODE As String = row.Item("SELL_CODE") & ""
            Dim STORES As Int64 = Val(row.Item("STORES") & "")

            Dim rowSOTALLOS As DataRow = dst.Tables("SOTALLOS").Rows.Find(New String() {"IPLBAE", "000" & SELL_CODE})
            If rowSOTALLOS IsNot Nothing Then
                rowSOTALLOS.Item("STORES") = STORES
                rowSOTALLOC.Item("STORES") = Val(rowSOTALLOC.Item("STORES") & "") + STORES
            End If
        Next
    End Sub

    Private Sub btnGet6Mos_Click(sender As Object, e As EventArgs) Handles btnGet6Mos.Click
        Get_HC_Data()
        chkPriorYear.Checked = False
        chkBalances.Checked = False
        chkSI.Checked = False
        chkST.Checked = False
    End Sub

    Private Sub txtITEM_CODE_ValueChanged(sender As Object, e As EventArgs) Handles txtITEM_CODE.ValueChanged

    End Sub

    Sub Load_from_XLS()

        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
            openFileDialog1.RestoreDirectory = True

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        If FILENAME <> "" Then

            Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
            Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
            Dim range As SpreadsheetGear.IRange = Nothing

            Dim ITEM_CODEs As New Dictionary(Of Integer, Integer)
            Dim c As Integer = 1
            Do While oSheet.Cells(0, c).Value & "" <> ""
                Dim ITEM_CODE As String = Split(oSheet.Cells(0, c).Value & "", vbLf)(0)
                Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then

                    Dim rowSOTALLO1 As DataRow = Nothing
                    Dim rowSOTALLO1s() As DataRow = dst.Tables("SOTALLO1").Select("ITEM_CODE = '" & ITEM_CODE & "'")
                    If rowSOTALLO1s.Length > 0 Then
                        rowSOTALLO1 = rowSOTALLO1s(0)
                    Else
                        rowSOTALLO1 = Add_Item(ITEM_CODE, True)
                        Add_Allocation_to_Grid(rowSOTALLO1)
                    End If

                    For i As Integer = 1 To ALLO_CTL_NOi.Length - 1
                        If ALLO_CTL_NOi(i) = rowSOTALLO1.Item("ALLO_CTL_NO") Then
                            ITEM_CODEs.Add(c, i)
                            Exit For
                        End If
                    Next

                End If

                c += 1
            Loop

            Dim r As Integer = 1
            Do While oSheet.Cells(r, 0).Value & "" <> ""
                Dim CUST_CODE As String = oSheet.Cells(r, 0).Value & ""
                Add_CUST_CODE(CUST_CODE, "")

                For Each grow As UltraWinGrid.UltraGridRow In grdSOTALLOC.Rows
                    If grow.IsDataRow AndAlso grow.Cells("CUST_CODE").Value = CUST_CODE Then
                        grow.Activate()
                        For Each c In ITEM_CODEs.Keys
                            Dim QTY As Int64 = Val(oSheet.Cells(r, c).Value & "")
                            If QTY <> 0 Then
                                grow.Cells("ALLO_" & Format(ITEM_CODEs(c), "00")).Value = QTY
                            End If
                        Next
                        grow.Update()
                    End If
                Next
                r += 1
            Loop
            Sort_grdColumns(grdSOTALLOC, "CUST_CODE")

            MsgBox("XLS has been Loaded", MsgBoxStyle.OkOnly, "Success")
        End If

    End Sub

    Private Sub grdSOTORDRA_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDRA.InitializeRow
        If e.Row.IsDataRow And chkSingleItemMode.Checked Then
            Dim ORDR_DATE_SHIPPED As Date = IIf(e.Row.Cells("ORDR_DATE_SHIPPED").Value & "" = "", Nothing, e.Row.Cells("ORDR_DATE_SHIPPED").Value)
            Dim ORDR_SHIP_DATE As Date = IIf(e.Row.Cells("ORDR_SHIP_DATE").Value & "" = "", Nothing, e.Row.Cells("ORDR_SHIP_DATE").Value)
            Dim ORDR_ALLO_DATE As Date = IIf(e.Row.Cells("ORDR_ALLO_DATE").Value & "" = "", Nothing, e.Row.Cells("ORDR_ALLO_DATE").Value)

            Dim ALLO_CTL_NO As String = e.Row.Cells("ALLO_CTL_NO").Value
            Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO)

            If rowSOTALLO1 IsNot Nothing Then
                Dim DATE_START As Date = rowSOTALLO1.Item("DATE_START")
                Dim DATE_END As Date = rowSOTALLO1.Item("DATE_END")

                If DATE_START & "" <> "" And DATE_END & "" <> "" Then
                    If ORDR_DATE_SHIPPED & "" <> "" Then
                        If Format(ORDR_DATE_SHIPPED, "yyyyMMdd") >= Format(DATE_START, "yyyyMMdd") And Format(ORDR_DATE_SHIPPED, "yyyyMMdd") <= Format(DATE_END, "yyyyMMdd") Then
                            e.Row.Cells("ORDR_DATE_SHIPPED").Appearance.ForeColor = Drawing.Color.Empty
                        Else
                            e.Row.Cells("ORDR_DATE_SHIPPED").Appearance.ForeColor = Drawing.Color.Red
                        End If
                    End If
                    If ORDR_ALLO_DATE & "" <> "" Then
                        If Format(ORDR_ALLO_DATE, "yyyyMMdd") >= Format(DATE_START, "yyyyMMdd") And Format(ORDR_ALLO_DATE, "yyyyMMdd") <= Format(DATE_END, "yyyyMMdd") Then
                            e.Row.Cells("ORDR_ALLO_DATE").Appearance.ForeColor = Drawing.Color.Empty
                        Else
                            e.Row.Cells("ORDR_ALLO_DATE").Appearance.ForeColor = Drawing.Color.Red
                        End If
                    End If
                    If ORDR_SHIP_DATE & "" <> "" Then
                        If Format(ORDR_SHIP_DATE, "yyyyMMdd") >= Format(DATE_START, "yyyyMMdd") And Format(ORDR_SHIP_DATE, "yyyyMMdd") <= Format(DATE_END, "yyyyMMdd") Then
                            e.Row.Cells("ORDR_ALLO_DATE").Appearance.ForeColor = Drawing.Color.Empty
                        Else
                            e.Row.Cells("ORDR_ALLO_DATE").Appearance.ForeColor = Drawing.Color.Red
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub btnUseShipBy_Click(sender As Object, e As EventArgs)
        If grdSOTORDRA.Selected.Rows.Count = 0 Then
            MsgBox("No Orders Selected", MsgBoxStyle.OkOnly, "Cannot Move Allocation")
        Else
            For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDRA.Selected.Rows
                Dim ORDR_GROUP_NO As String = grow.Cells("ORDR_GROUP_NO").Value
                Dim CUST_DC_NO As String = grow.Cells("CUST_DC_NO").Value
                Dim USE_DATE As Date = grow.Cells("ORDR_SHIP_DATE").Value
                ASCMAIN1.sql = "Update SOTORDR1 Set ORDR_ALLO_DATE = '" & Format(USE_DATE, "dd-MMM-yyyy") & "'" & vbCrLf _
                    & " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and CUST_DC_NO = '" & CUST_DC_NO & "'"

                ASCMAIN1.sql = "Update SOTORDR1 Set ORDR_ALLO_DATE = NULL" & vbCrLf _
                    & " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and CUST_DC_NO = '" & CUST_DC_NO & "'"
                ASCDATA1.ExecuteSQL()
            Next
        End If
    End Sub

    Private Sub btnUseShipped_Click(sender As Object, e As EventArgs)
        If grdSOTORDRA.Selected.Rows.Count = 0 Then
            MsgBox("No Orders Selected", MsgBoxStyle.OkOnly, "Cannot Move Allocation")
        Else
            For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDRA.Selected.Rows
                Dim ORDR_GROUP_NO As String = grow.Cells("ORDR_GROUP_NO").Value
                Dim CUST_DC_NO As String = grow.Cells("CUST_DC_NO").Value
                Dim USE_DATE As Date = grow.Cells("ORDR_DATE_SHIPPED").Value
                ASCMAIN1.sql = "Update SOTORDR1 Set ORDR_ALLO_DATE = '" & Format(USE_DATE, "dd-MMM-yyyy") & "'" & vbCrLf _
                    & " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and CUST_DC_NO = '" & CUST_DC_NO & "'"
                ASCDATA1.ExecuteSQL()
            Next
        End If
    End Sub

    Private Sub grdSOTALLO1_BeforeExitEditMode(sender As Object, e As UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTALLO1.BeforeExitEditMode
        Try
            If grdSOTALLO1.ActiveCell IsNot Nothing Then
                With grdSOTALLO1.ActiveCell
                    Select Case .Column.Key
                        Case "ITEM_CODE_COMPARE_TO"
                            .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)

                    End Select
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cmdFindChanges_Click(sender As Object, e As EventArgs) Handles cmdFindChanges.Click
        Find_Changes()
    End Sub
    Sub Find_Changes()
        Dim YPx As String = cmbYP.Value
        Dim YP As String = Mid(YPx, 1, 4) & Mid(YPx, 6, 2)
        Fill_Records("SOTORDRU", YP)
        Sort_grdColumns(grdSOTORDRU, "CUST_CODE,ITEM_CODE")
        grdSOTORDRU.Text = "Changes to Allocations upon Shipment Update for Shipments in " & cmbYP.Text
        grdSOTORDRU.Visible = True
    End Sub

    Private Sub grdSOTALLO1_BeforeRowRegionScroll(sender As Object, e As UltraWinGrid.BeforeRowRegionScrollEventArgs) Handles grdSOTALLO1.BeforeRowRegionScroll

    End Sub

    Private Sub btnCopyDates_Click(sender As Object, e As EventArgs) Handles btnCopyDates.Click
        If datesSet Then
            For Each row As DataRow In dst.Tables("SOTALLO1").Rows
                row("DATE_START") = lastStart
                row("DATE_END") = lastEnd
            Next
        End If
    End Sub

    Private Sub LoadAllocationChangeHistory(ByVal ALLO_CTL_NO As String)
        dst.Tables("SOTALLH3").Rows.Clear()
        dst.Tables("SOTALLH2").Rows.Clear()
        dst.Tables("SOTALLH1").Rows.Clear()


        Fill_Records("SOTALLH1", New String() {ALLO_CTL_NO})
        Fill_Records("SOTALLH2", New String() {ALLO_CTL_NO})
        Fill_Records("SOTALLH3", New String() {ALLO_CTL_NO})
    End Sub
    Private Sub grdSOTORDRD_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDRD.InitializeRow
        If e.Row.Band.Index = 1 Then
            Dim ADDS As Integer = Convert.ToInt32(e.Row.Cells("ADDS").Value)
            Dim DEDS As Integer = Convert.ToInt32(e.Row.Cells("DEDS").Value)
            If ADDS = 0 AndAlso DEDS = 0 Then
                e.Row.Hidden = True
            End If
        End If
    End Sub

    Private Sub grdSOTORDRD_AfterRowActivate(ByVal sender As Object, ByVal e As EventArgs) Handles grdSOTORDRD.AfterRowActivate
        Me.grdSOTORDRD.Rows.ExpandAll(True)
    End Sub

    Private Sub btnAddEvent_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAddEvent.Click
        Dim newEventName As String = cmbEvent.Text.Trim()

        If String.IsNullOrEmpty(newEventName) Then
            MessageBox.Show("Please enter a valid event name.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        For i As Integer = 1 To 10
            Dim colKey As String = "EVENT_" & i.ToString("00")
            If grdSOTALLOT.DisplayLayout.Bands(0).Columns.Exists(colKey) Then
                Dim col As UltraWinGrid.UltraGridColumn = grdSOTALLOT.DisplayLayout.Bands(0).Columns(colKey)
                If Not col.Hidden AndAlso col.Header.Caption = newEventName Then
                    MessageBox.Show("This event already exists.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If
            End If
        Next

        Dim availableColumn As UltraWinGrid.UltraGridColumn = Nothing
        For i As Integer = 1 To 10
            Dim colKey As String = "EVENT_" & i.ToString("00")
            If grdSOTALLOT.DisplayLayout.Bands(0).Columns.Exists(colKey) Then
                Dim col As UltraWinGrid.UltraGridColumn = grdSOTALLOT.DisplayLayout.Bands(0).Columns(colKey)
                If col.Hidden Then
                    availableColumn = col
                    Exit For
                End If
            End If
        Next

        availableColumn.Header.Caption = newEventName
        availableColumn.Hidden = False
        grdSOTALLOT.DataBind()
    End Sub

    Private Sub grdSOTALLOT_AfterCellUpdate(ByVal sender As Object, ByVal e As UltraWinGrid.CellEventArgs) Handles grdSOTALLOT.AfterCellUpdate
        Dim rowSOTALLOT As UltraWinGrid.UltraGridRow = TryCast(e.Cell.Row, UltraWinGrid.UltraGridRow)
        If rowSOTALLOT Is Nothing Then Exit Sub

        Dim CUST_CODE As String = rowSOTALLOT.Cells("CUST_CODE").Value.ToString()
        Dim CUST_STORE_NO As String = rowSOTALLOT.Cells("CUST_STORE_NO").Value.ToString()

        Dim QTY As Int64 = 0
        For iCtr As Integer = 1 To 10
            Dim colKey As String = "EVENT_" & Format(iCtr, "00")
            If grdSOTALLOT.DisplayLayout.Bands(0).Columns.Exists(colKey) Then
                QTY += Val(rowSOTALLOT.Cells(colKey).Value & "") ' Prevents null errors
            End If
        Next

        For Each rowSOTALLOS As UltraWinGrid.UltraGridRow In grdSOTALLOS.Rows
            If rowSOTALLOS.Cells("CUST_CODE").Value.ToString() = CUST_CODE AndAlso
               rowSOTALLOS.Cells("CUST_STORE_NO").Value.ToString() = CUST_STORE_NO Then

                If rowSOTALLOS.Cells.Exists("TOTAL_EVENT_QTY") Then
                    rowSOTALLOS.Cells("TOTAL_EVENT_QTY").Value = QTY
                End If
                Exit For
            End If
        Next
    End Sub


End Class