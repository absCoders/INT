Imports Infragistics.Win.UltraWinGrid

Public Class DPFITMFX
    Dim YP() As String
    Dim ICTITEMS As String
    Dim YPF(,) As String
    Dim YPFD() As Date
    Dim YPP(,) As String
    Dim YPPD() As Date
    Dim FCMAX As Int32 = 26 ' 25 Future Periods - currently supporting only 18 for forecasting
    Dim SHMAX As Int32 = 26
    Dim TRMAX As Int32 = 6
    Dim CSMAX As Int32 = 24
    Dim IPMAX As Int32 = 24

    Dim SATSLSCT As String
    Dim sqlSATSLSCT As String
    Dim SATSLSCI As String
    Dim sqlSATSLSCI As String
    Dim SATSLSIT As String
    Dim sqlSATSLSIT As String
    Dim SATSLSII As String
    Dim sqlSATSLSII As String
    Dim DPTITMFH As String
    Dim sqlDPTITMFH As String
    Dim DPTITMFO As String
    Dim sqlDPTITMFO As String
    Dim rowICTITEM1 As DataRow


    Dim Ps As New List(Of String)
    Dim AllocPs As New List(Of String)
    Dim PerBs As New List(Of String)
    Dim sqlPlans As String
    Dim CUST_DISC_PCT As Decimal = 0
    Dim PRICE_CLASS_CODE As String
    Dim PRICE_LIST_CODE As String
    Dim CUST_CODE As String
    Dim sqlDPTITMFX_where As String
    Dim DPTITMFP_PROCESS As Boolean = False
    Dim BASE_IP_PERC_CHANGE As Boolean = False
    Dim PREV_COLL_CMB As String = ""


    Dim DP_PARM_ALLOC_MARKET As String = "DPT" ' only for DPT will the Allocations Tab appear

    Dim ITEM_CODE_F As String
    Dim modes_Edit As Boolean = False
    Dim cellHasNotes As New Infragistics.Win.Appearance

    Dim sqlDPTITFX_filter As String = ""
    Private QuarterlyUploadYear As String = ""


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")
        Get_PARM("DPTPARM1")
        Setup_Periods()


        With dst

            ' ICTITEMS - All Items on Screen

            ASCMAIN1.sql = "Select ITEM_CODE from ICTITEM1 where ROWNUM < 1"
            ICTITEMS = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add Primary Key (ITEM_CODE)")


            Dim sqlU As String
            Dim sqlS As String

            Dim SATSLSC1cols As String = ""

            ' SATSLSCI

            sqlU = ""
            sqlS = ""
            sqlSATSLSCI = "Select SATSSUMI.CUST_CODE"
            For P As Integer = 0 To CSMAX
                Dim YP As String = YPP(P, 0)
                COLUMN_NAME = "U" & Format(P, "00")
                sqlU &= ", SUM (DECODE(SATSSUMI.OPS_YYYYPP,'" & YP & "',SATSSUMI.ORDR_QTY_SHIP,0)) " & COLUMN_NAME
                SATSLSC1cols &= ", " & COLUMN_NAME
                COLUMN_NAME = "S" & Format(P, "00")
                sqlS &= ", SUM (DECODE(SATSSUMI.OPS_YYYYPP,'" & YP & "',SATSSUMI.ORDR_AMT_SHIP,0)) " & COLUMN_NAME
                SATSLSC1cols &= ", " & COLUMN_NAME
            Next
            sqlSATSLSCI &= sqlU & sqlS
            sqlSATSLSCI &= " from SATSSUMI" _
            & " where SATSSUMI.ITEM_CODE = :PARM1" _
            & "   and SATSSUMI.INV_TYPE = 'I'" _
            & "   and SATSSUMI.OPS_YYYYPP BETWEEN '" & YPP(CSMAX, 0) & "' AND '" & YPP(0, 0) & "'" _
            & " group by SATSSUMI.CUST_CODE"
            SATSLSCI = ASCMAIN1.Temp_Table(Replace(sqlSATSLSCI, ":PARM1", "null"))

            ' SATSLSCT

            sqlU = ""
            sqlS = ""
            sqlSATSLSCT = "Select RSTRETL1.CUST_CODE"
            For P As Integer = 0 To CSMAX
                Dim YP As String = YPP(P, 0)
                COLUMN_NAME = "RU" & Format(P, "00")
                sqlU &= ", SUM (DECODE(RSTRETL1.OPS_YYYYPP,'" & YP & "',RSTRETL1.QTY_SOLD,0)) " & COLUMN_NAME
                SATSLSC1cols &= ", " & COLUMN_NAME
                COLUMN_NAME = "RS" & Format(P, "00")
                sqlS &= ", SUM (DECODE(RSTRETL1.OPS_YYYYPP,'" & YP & "',RSTRETL1.AMT_SOLD,0)) " & COLUMN_NAME
                SATSLSC1cols &= ", " & COLUMN_NAME
            Next
            sqlSATSLSCT &= sqlU & sqlS
            sqlSATSLSCT &= " from RSTRETL1" _
            & " where RSTRETL1.ITEM_CODE = :PARM1" _
            & "   and RSTRETL1.OPS_YYYYPP BETWEEN '" & YPP(CSMAX, 0) & "' AND '" & YPP(0, 0) & "'" _
            & " group by RSTRETL1.CUST_CODE"
            SATSLSCT = ASCMAIN1.Temp_Table(Replace(sqlSATSLSCT, ":PARM1", "null"))


            ASCMAIN1.sql = "Select SATSLSCI.CUST_CODE, ARTCUST1.CUST_NAME" _
            & SATSLSC1cols & " from ARTCUST1" _
            & ", " & SATSLSCI & " SATSLSCI" _
            & ", " & SATSLSCT & " SATSLSCT" _
            & " where SATSLSCT.CUST_CODE (+) = SATSLSCI.CUST_CODE " _
            & " and ARTCUST1.CUST_CODE = SATSLSCI.CUST_CODE"
            Create_TDA(.Tables.Add, "SATSLSC1", "**", 0, False, "", 1)
            For P = 0 To CSMAX
                .Tables("SATSLSC1").Columns("U" & Format(P, "00")).DataType = GetType(System.Int32)
                .Tables("SATSLSC1").Columns("S" & Format(P, "00")).DataType = GetType(System.Int32)
                .Tables("SATSLSC1").Columns("RU" & Format(P, "00")).DataType = GetType(System.Int32)
                .Tables("SATSLSC1").Columns("RS" & Format(P, "00")).DataType = GetType(System.Int32)
            Next

            ' DPTITMFH

            sqlDPTITMFH = "Select DPTITMF1.ITEM_CODE"
            For P As Integer = 6 To 0 Step -1
                sqlDPTITMFH &= ", SUM (DECODE(DPTITMF1.OPS_YYYYPP,'" & YPP(P, 0) & "',DPTITMF1.FORECAST,0)) HUFC" & Format(P, "00")
            Next
            sqlDPTITMFH &= " FROM DPTITMF1, " & ICTITEMS & " ICTITEMS" _
            & " WHERE DPTITMF1.OPS_YYYYPP BETWEEN '" & YPP(6, 0) & "' AND '" & YPP(0, 0) & "'" _
            & " AND DPTITMF1.OPS_YYYYPP_FC = DPTITMF1.OPS_YYYYPP" _
            & " AND ICTITEMS.ITEM_CODE = DPTITMF1.ITEM_CODE" _
            & " group by DPTITMF1.ITEM_CODE"
            DPTITMFH = ASCMAIN1.Temp_Table(sqlDPTITMFH)
            ASCDATA1.ExecuteSQL("Alter Table " & DPTITMFH & " Add Primary Key (ITEM_CODE)")
            ASCMAIN1.sql = "Select * from " & DPTITMFH
            Create_TDA(.Tables.Add, "DPTITMFH", "**", 0, False, "", 1)


            ' DPTITMFO - THIS ONE NEEDS TO BE FIXED

            sqlDPTITMFO = "Select DPTITMF1.ITEM_CODE"
            For P As Integer = -6 To 0
                Dim YP As String
                Dim CSFX As String
                YP = YPP(-1 * P, 0)
                CSFX = "P" & Format(-1 * P, "00")
                sqlDPTITMFO &= ", SUM (DECODE(DPTITMF1.OPS_YYYYPP,'" & YP & "',DPTITMF1.FORECAST,0)) OUFC" & CSFX
            Next
            For P As Integer = 0 To FCMAX
                Dim YP As String
                Dim CSFX As String
                YP = YPF(P, 0)
                CSFX = "F" & Format(P, "00")
                sqlDPTITMFO &= ", SUM (DECODE(DPTITMF1.OPS_YYYYPP,'" & YP & "',DPTITMF1.FORECAST,0)) OUFC" & CSFX
            Next
            sqlDPTITMFO &= " FROM DPTITMF1, " & ICTITEMS & " ICTITEMS" _
            & " WHERE DPTITMF1.OPS_YYYYPP BETWEEN '" & YPP(6, 0) & "' AND '" & YPP(FCMAX, 0) & "'" _
            & " AND DPTITMF1.OPS_YYYYPP_FC = DPTITMF1.OPS_YYYYPP" _
            & " AND ICTITEMS.ITEM_CODE = DPTITMF1.ITEM_CODE" _
            & " group by DPTITMF1.ITEM_CODE"
            DPTITMFO = ASCMAIN1.Temp_Table(sqlDPTITMFO)
            ASCDATA1.ExecuteSQL("Alter Table " & DPTITMFO & " Add Primary Key (ITEM_CODE)")
            ASCMAIN1.sql = "Select * from " & DPTITMFO
            Create_TDA(.Tables.Add, "DPTITMFO", "**", 0, False, "", 1)

            ' DPTITMFX

            ASCMAIN1.sql = "Select DPTITMF1.ITEM_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_CLASS_CODE" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.PROD_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_BASIC_PROMO, ICTITEM1.COST_CATGY_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_RETAIL_PRICE"
            For P As Integer = -1 To FCMAX
                Dim COLUMN_NAME As String = "UFCPD"
                Dim YP As String = "000000"
                If P >= 0 Then
                    COLUMN_NAME = "UFC" & Format(P, "00")
                    YP = YPF(P, 0)
                End If
                ASCMAIN1.sql &= ", SUM (DECODE(DPTITMF1.OPS_YYYYPP_FC,'" & YP & "',DPTITMF1.FORECAST,0)) " & COLUMN_NAME
            Next
            ASCMAIN1.sql &= " from DPTITMF1, " & ICTITEMS & " ICTITEMS, ICTITEM1, ICTCOLL1" & vbCrLf _
            & " where DPTITMF1.ITEM_CODE = ICTITEMS.ITEM_CODE" & vbCrLf _
            & " and DPTITMF1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & " and ICTITEM1.ITEM_CODE = ICTITEMS.ITEM_CODE" & vbCrLf _
            & " and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " and DPTITMF1.MARKET_CODE = :PARM1" & vbCrLf _
            & " group by DPTITMF1.ITEM_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_CLASS_CODE" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.PROD_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_BASIC_PROMO, ICTITEM1.COST_CATGY_CODE" _
            & ", ICTITEM1.ITEM_RETAIL_PRICE, .50 * ICTITEM1.ITEM_PRICE"
            '& ", ICTITEM1.ITEM_RETAIL_PRICE, .60 * ICTITEM1.ITEM_PRICE"
            Create_TDA(.Tables.Add, "DPTITMFX", "**", 0, False, "V", 1)
            .Tables("DPTITMFX").Columns.Add("ITEM_WHOLESALE_PRICE", GetType(System.Decimal), "ITEM_RETAIL_PRICE * 50 / 100")
            .Tables("DPTITMFX").Columns.Add("ITEM_PRICE", GetType(System.Decimal), "ITEM_WHOLESALE_PRICE")
            Dim TOTAL_EXP As String = ""
            For P = -1 To FCMAX
                Dim COLUMN_NAME As String = IIf(P = -1, "UFCPD", "UFC" & Format(P, "00"))
                .Tables("DPTITMFX").Columns(COLUMN_NAME).DataType = GetType(System.Int32)
                TOTAL_EXP &= "+ISNULL(" & COLUMN_NAME & ",0)"
            Next
            .Tables("DPTITMFX").Columns.Add("UFCTOT", GetType(System.Int32), Mid(TOTAL_EXP, 2))

            TOTAL_EXP = ""
            For P = -1 To FCMAX
                Dim COLUMN_NAME As String = IIf(P = -1, "SFCPD", "SFC" & Format(P, "00"))
                .Tables("DPTITMFX").Columns.Add(COLUMN_NAME, GetType(System.Int32), "U" & Mid(COLUMN_NAME, 2) & "*ISNULL(ITEM_PRICE,0)")
                TOTAL_EXP &= "+ISNULL(" & COLUMN_NAME & ",0)"
                ' temp = should be decimal for sales
            Next
            .Tables("DPTITMFX").Columns.Add("SFCTOT", GetType(System.Int32), Mid(TOTAL_EXP, 2))

            .Tables("DPTITMFX").Columns.Add("EXCL")

            ' DGJ
            For P = -1 To FCMAX
                If P = -1 Then
                    .Tables("DPTITMFX").Columns.Add("BASE_IP", GetType(System.Decimal))
                Else
                    Dim COLUMN_NAME As String = IIf(P = -1, "IFCPD", "IFC" & Format(P, "00"))
                    .Tables("DPTITMFX").Columns.Add(COLUMN_NAME, GetType(System.Double))
                    COLUMN_NAME = IIf(P = -1, "IFCPD", "IFCA" & Format(P, "00"))
                    .Tables("DPTITMFX").Columns.Add(COLUMN_NAME, GetType(System.Double))
                End If
            Next

            Dim TBL1 As DataTable = .Tables("DPTITMFX").Copy()
            TBL1.TableName = "DPTITMFY"
            .Tables.Add(TBL1)

            With dst.Tables("DPTITMFY")
                .Columns.Add("RS09", GetType(System.Double))
                .Columns.Add("RS10", GetType(System.Double))
                .Columns.Add("RS11", GetType(System.Double))
                .Columns.Add("RS12", GetType(System.Double))

                .PrimaryKey = New DataColumn() { .Columns("ITEM_CODE")}
            End With


            With cellHasNotes
                .BackColor2 = Color.White
                .BackColor = Color.LightSkyBlue
                .BackGradientStyle = GradientStyle.ForwardDiagonal
            End With


            ASCMAIN1.sql = "Select DPTITMF2.*" & vbCrLf _
            & " from DPTITMF2," & ICTITEMS & " ICTITEMS" & vbCrLf _
            & " where DPTITMF2.ITEM_CODE = ICTITEMS.ITEM_CODE " & vbCrLf _
            & "   And DPTITMF2.OPS_YYYYPP_FC >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -6) & "'"
            Create_TDA(.Tables.Add, "DPTITMF2", "**", 0, False, "", 0)


            Dim SATSLSI1cols As String = ""

            ' SATSLSII

            sqlU = ""
            sqlS = ""
            sqlSATSLSII = "Select SATSSUMI.ITEM_CODE" & vbCrLf
            For P As Integer = 0 To SHMAX
                Dim YP As String = YPP(P, 0)
                COLUMN_NAME = "U" & Format(P, "00")
                sqlU &= ", SUM (DECODE(SATSSUMI.OPS_YYYYPP,'" & YP & "',SATSSUMI.ORDR_QTY_SHIP,0)) " & COLUMN_NAME & vbCrLf
                SATSLSI1cols &= ", " & COLUMN_NAME
                COLUMN_NAME = "S" & Format(P, "00")
                sqlS &= ", SUM (DECODE(SATSSUMI.OPS_YYYYPP,'" & YP & "',SATSSUMI.ORDR_AMT_SHIP,0)) " & COLUMN_NAME & vbCrLf
                SATSLSI1cols &= ", " & COLUMN_NAME
            Next
            sqlSATSLSII &= sqlU & sqlS
            sqlSATSLSII &= " from SATSSUMI,ARTCUST1,SOTTCLS1,SOTMKTC1" & vbCrLf _
            & " where SATSSUMI.ITEM_CODE in (Select Distinct ITEM_CODE from " & ICTITEMS & ")" & vbCrLf _
            & "   and SATSSUMI.INV_TYPE = 'I'" & vbCrLf _
            & "   and SATSSUMI.OPS_YYYYPP BETWEEN '" & YPP(SHMAX, 0) & "' AND '" & YPP(0, 0) & "'" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SATSSUMI.CUST_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and SOTMKTC1.CUST_CODE (+) = SATSSUMI.CUST_CODE" & vbCrLf _
            & "   and NVL(SOTMKTC1.MARKET_CODE,SOTTCLS1.MARKET_CODE) = ''" & vbCrLf _
            & " group by SATSSUMI.ITEM_CODE"
            SATSLSII = ASCMAIN1.Temp_Table(sqlSATSLSII)

            ' SATSLSIT

            sqlU = ""
            sqlS = ""
            sqlSATSLSIT = "Select RSTRETL1.ITEM_CODE"
            For P As Integer = 0 To SHMAX
                Dim YP As String = YPP(P, 0)
                COLUMN_NAME = "RU" & Format(P, "00")
                sqlU &= ", SUM (DECODE(RSTRETL1.OPS_YYYYPP,'" & YP & "',RSTRETL1.QTY_SOLD,0)) " & COLUMN_NAME
                SATSLSI1cols &= ", " & COLUMN_NAME
                COLUMN_NAME = "RS" & Format(P, "00")
                sqlS &= ", SUM (DECODE(RSTRETL1.OPS_YYYYPP,'" & YP & "',RSTRETL1.AMT_SOLD,0)) " & COLUMN_NAME
                SATSLSI1cols &= ", " & COLUMN_NAME
            Next
            sqlSATSLSIT &= sqlU & sqlS
            sqlSATSLSIT &= " from RSTRETL1,ARTCUST1,SOTTCLS1,SOTMKTC1" _
            & " where RSTRETL1.ITEM_CODE in (Select Distinct ITEM_CODE from " & ICTITEMS & ")" _
            & "   and RSTRETL1.OPS_YYYYPP BETWEEN '" & YPP(SHMAX, 0) & "' AND '" & YPP(0, 0) & "'" _
            & "   and ARTCUST1.CUST_CODE = RSTRETL1.CUST_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and SOTMKTC1.CUST_CODE (+) = RSTRETL1.CUST_CODE" & vbCrLf _
            & "   and NVL(SOTMKTC1.MARKET_CODE,SOTTCLS1.MARKET_CODE) = ''" & vbCrLf _
            & " group by RSTRETL1.ITEM_CODE"
            SATSLSIT = ASCMAIN1.Temp_Table(sqlSATSLSIT)


            ' SATSLSI1

            ASCMAIN1.sql = "Select SATSLSII.ITEM_CODE, ICTITEM1.ITEM_DESC" _
            & SATSLSI1cols & " from ICTITEM1" _
            & ", " & SATSLSII & " SATSLSII" _
            & ", " & SATSLSIT & " SATSLSIT" _
            & " where SATSLSIT.ITEM_CODE (+) = SATSLSII.ITEM_CODE " _
            & " and ICTITEM1.ITEM_CODE = SATSLSII.ITEM_CODE"
            Create_TDA(.Tables.Add, "SATSLSI1", "**", 0, False, "", 1)
            For P = 0 To SHMAX
                .Tables("SATSLSI1").Columns("U" & Format(P, "00")).DataType = GetType(System.Int32)
                .Tables("SATSLSI1").Columns("S" & Format(P, "00")).DataType = GetType(System.Int32)
                .Tables("SATSLSI1").Columns("RU" & Format(P, "00")).DataType = GetType(System.Int32)
                .Tables("SATSLSI1").Columns("RS" & Format(P, "00")).DataType = GetType(System.Int32)
            Next

            Dim TBL As DataTable = .Tables("SATSLSI1").Copy
            TBL.TableName = "SATSLSI1_ALL"
            .Tables.Add(TBL)


            .Relations.Add("DPTITMFX_SATSLSI1",
            New DataColumn() { .Tables("DPTITMFX").Columns("ITEM_CODE")},
            New DataColumn() { .Tables("SATSLSI1").Columns("ITEM_CODE")})

            .Relations.Add("DPTITMFX_DPTITMFH",
            New DataColumn() { .Tables("DPTITMFX").Columns("ITEM_CODE")},
            New DataColumn() { .Tables("DPTITMFH").Columns("ITEM_CODE")})

            .Relations.Add("DPTITMFX_DPTITMFO",
            New DataColumn() { .Tables("DPTITMFX").Columns("ITEM_CODE")},
            New DataColumn() { .Tables("DPTITMFO").Columns("ITEM_CODE")})


            For P = 0 To 6
                With .Tables("DPTITMFX").Columns
                    .Add("TOU" & Format(P, "00"), GetType(System.Int32), "SUM (CHILD(DPTITMFX_DPTITMFO).OUFCP" & Format(P, "00") & ")")
                    .Add("TOS" & Format(P, "00"), GetType(System.Int32), "ISNULL(ITEM_PRICE,0) * SUM (CHILD(DPTITMFX_DPTITMFO).OUFCP" & Format(P, "00") & ")")
                    .Add("TPU" & Format(P, "00"), GetType(System.Int32), "SUM (CHILD(DPTITMFX_DPTITMFH).HUFC" & Format(P, "00") & ")")
                    .Add("TPS" & Format(P, "00"), GetType(System.Int32), "ISNULL(ITEM_PRICE,0) * SUM (CHILD(DPTITMFX_DPTITMFH).HUFC" & Format(P, "00") & ")")

                    .Add("TU" & Format(P, "00"), GetType(System.Int32), "SUM (CHILD(DPTITMFX_SATSLSI1).U" & Format(P, "00") & ")")
                    .Add("TS" & Format(P, "00"), GetType(System.Int32), "SUM (CHILD(DPTITMFX_SATSLSI1).S" & Format(P, "00") & ")")
                    .Add("TU" & Format(P + 12, "00"), GetType(System.Int32), "SUM (CHILD(DPTITMFX_SATSLSI1).U" & Format(P + 12, "00") & ")")
                    .Add("TS" & Format(P + 12, "00"), GetType(System.Int32), "SUM (CHILD(DPTITMFX_SATSLSI1).S" & Format(P + 12, "00") & ")")

                    .Add("TAPU" & Format(P, "00"), GetType(System.Double), "IIF (ISNULL(" & "TPU" & Format(P, "00") & ",0)=0,0,100*(ISNULL(" & "TU" & Format(P, "00") & ",0)-ISNULL(" & "TPU" & Format(P, "00") & ",0))/ISNULL(" & "TPU" & Format(P, "00") & ",0))")
                    .Add("TAPS" & Format(P, "00"), GetType(System.Double), "IIF (ISNULL(" & "TPS" & Format(P, "00") & ",0)=0,0,100*(ISNULL(" & "TS" & Format(P, "00") & ",0)-ISNULL(" & "TPS" & Format(P, "00") & ",0))/ISNULL(" & "TPS" & Format(P, "00") & ",0))")
                    .Add("TALU" & Format(P, "00"), GetType(System.Double), "IIF (ISNULL(" & "TU" & Format(P + 12, "00") & ",0)=0,0,100*(ISNULL(" & "TU" & Format(P, "00") & ",0)-ISNULL(" & "TU" & Format(P + 12, "00") & ",0))/ISNULL(" & "TU" & Format(P + 12, "00") & ",0))")
                    .Add("TALS" & Format(P, "00"), GetType(System.Double), "IIF (ISNULL(" & "TS" & Format(P + 12, "00") & ",0)=0,0,100*(ISNULL(" & "TS" & Format(P, "00") & ",0)-ISNULL(" & "TS" & Format(P + 12, "00") & ",0))/ISNULL(" & "TS" & Format(P + 12, "00") & ",0))")
                End With
            Next
            For P = 0 To 12
                Dim SP As Int32 = 12 - P
                With .Tables("DPTITMFX").Columns
                    .Add("OU" & Format(P, "00"), GetType(System.Int32), "SUM (CHILD(DPTITMFX_DPTITMFO).OUFCF" & Format(P, "00") & ")")
                    .Add("OS" & Format(P, "00"), GetType(System.Int32), "ISNULL(ITEM_PRICE,0) * SUM (CHILD(DPTITMFX_DPTITMFO).OUFCF" & Format(P, "00") & ")")

                    .Add("U" & Format(P, "00"), GetType(System.Int32), "SUM (CHILD(DPTITMFX_SATSLSI1).U" & Format(SP, "00") & ")")
                    .Add("S" & Format(P, "00"), GetType(System.Int32), "SUM (CHILD(DPTITMFX_SATSLSI1).S" & Format(SP, "00") & ")")
                    .Add("RU" & Format(P, "00"), GetType(System.Int32), "SUM (CHILD(DPTITMFX_SATSLSI1).RU" & Format(SP, "00") & ")")
                    .Add("RS" & Format(P, "00"), GetType(System.Int32), "SUM (CHILD(DPTITMFX_SATSLSI1).RS" & Format(SP, "00") & ")")
                    .Add("VU" & Format(P, "00"), GetType(System.Int32), "UFC" & Format(P, "00") & "-" & "U" & Format(P, "00"))
                    .Add("VS" & Format(P, "00"), GetType(System.Int32), "SFC" & Format(P, "00") & "-" & "S" & Format(P, "00"))
                    .Add("PU" & Format(P, "00"), GetType(System.Double), "IIF(ISNULL(" & "U" & Format(P, "00") & ",0)=0,0," & "100*ISNULL(" & "VU" & Format(P, "00") & ",0)/" & "ISNULL(" & "U" & Format(P, "00") & ",0))")
                    .Add("PS" & Format(P, "00"), GetType(System.Double), "IIF(ISNULL(" & "S" & Format(P, "00") & ",0)=0,0," & "100*ISNULL(" & "VS" & Format(P, "00") & ",0)/" & "ISNULL(" & "S" & Format(P, "00") & ",0))")
                End With
            Next

            Create_TDA(.Tables.Add, "DPTITMF1", "*")

            Create_TDA(.Tables.Add, "DPTITMB1", "*")

            Create_TDA(.Tables.Add, "DPTXLSX1", "*")

            With .Tables.Add("DPTITMFP")
                .Columns.Add("DATA_TYPE")
                .Columns.Add("DATA_DESC")
                '.Columns.Add("DATA_TYPE1")
                '.Columns.Add("DATA_TYPE2")

                Dim T As String = ""
                For I As Integer = 0 To FCMAX
                    .Columns.Add("P" & Format(I, "00"), GetType(System.Decimal))
                    .Columns.Add("A" & Format(I, "00"), GetType(System.Decimal))
                    T &= "+ISNULL(P" & Format(I, "00") & ",0)"
                Next
                .Columns.Add("P_TOTAL", GetType(System.Decimal), Mid(T, 2))
                .Columns.Add("A_TOTAL", GetType(System.Decimal), Replace(Mid(T, 2), "P", "A"))

                .PrimaryKey = New DataColumn() { .Columns("DATA_TYPE")}
            End With


            With .Tables.Add("DPTITMFZ")
                .Columns.Add("DATA_DESC")
                For I As Integer = 0 To FCMAX
                    .Columns.Add("P" & Format(I, "00"), GetType(System.Decimal))
                Next
                .PrimaryKey = New DataColumn() { .Columns("DATA_TYPE")}
            End With

            Dim rowDPTITMFZ As DataRow = dst.Tables("DPTITMFZ").NewRow
            rowDPTITMFZ.Item("DATA_DESC") = "%Change"
            .Tables("DPTITMFZ").Rows.Add(rowDPTITMFZ)


            With .Tables.Add("SOTALLOX")
                .Columns.Add("ITEM_CODE")
                .Columns.Add("ITEM_DESC")
                .Columns.Add("SEL")
                .Columns("SEL").DefaultValue = "0"

                Dim T As String = ""
                For I As Integer = 0 To FCMAX
                    .Columns.Add("P" & Format(I, "00"), GetType(System.Int64))
                    .Columns.Add("A" & Format(I, "00"), GetType(System.Int64))
                    T &= "+ISNULL(P" & Format(I, "00") & ",0)"
                Next
                .Columns.Add("P_TOTAL", GetType(System.Decimal), Mid(T, 2))
                .Columns.Add("A_TOTAL", GetType(System.Decimal), Replace(Mid(T, 2), "P", "A"))

                .PrimaryKey = New DataColumn() { .Columns("ITEM_CODE")}
            End With

            With .Tables.Add("DPTITMFM")
                .Columns.Add("P")
                .Columns.Add("LEGEND")
            End With

            With .Tables.Add("SOTALLOM")
                .Columns.Add("P")
                .Columns.Add("LEGEND")
            End With

            Dim sqlPlan As String = ""
            Dim sqlFC As String = ""
            Dim sqlPO As String = ""
            Dim sqlPP As String = ""
            Dim PRD_END_DATE_LMO As String = Format(Now.Date.AddYears(-1), "dd-MMM-yyyy")
            For i As Integer = 0 To 26 ' USED TO BE 1 TO 12 - LM THEN WANTED CURRENT MONTH BROKEN OUT
                Dim YP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, i)
                sqlFC &= ", Sum (Decode(OPS_YYYYPP_FC,'" & YP & "',FORECAST,0)) FC" & Format(i, "00") & vbCrLf
                'If i <= 12 Then
                Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", YP)
                'Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", ASCMAIN1.Period_Calc(YP, 1))
                Dim PRD_END_DATE As String = Format(rowGLTPARM2.Item("PRD_END_DATE"), "dd-MMM-yyyy")
                sqlPlan &= ", Sum (Case when DATE_REQUIRED > '" & PRD_END_DATE_LMO & "'" _
                    & " and DATE_REQUIRED <= '" & PRD_END_DATE & "' then QTY_PLANNED else 0 End) FC" & Format(i, "00") & vbCrLf
                sqlPO &= ", Sum (Case when POTORDR2.PO_DATE_REQUIRED > '" & PRD_END_DATE_LMO & "'" _
                    & " and POTORDR2.PO_DATE_REQUIRED <= '" & PRD_END_DATE & "' then POTORDR2.PO_QTY_OPN else 0 End) FC" & Format(i, "00") & vbCrLf
                sqlPP &= ", Sum (FC" & Format(i, "00") & ") FC" & Format(i, "00")
                PRD_END_DATE_LMO = PRD_END_DATE
                'End If
            Next

            sqlFC = "Select ITEM_CODE" & vbCrLf & sqlFC & " from DPTITMF1 where OPS_YYYYPP = '" & ASCMAIN1.CYP & "' group by ITEM_CODE"
            sqlPlan = "Select ITEM_CODE" & vbCrLf & sqlPP & " from (" & vbCrLf _
                & "Select ITEM_CODE" & vbCrLf & sqlPlan & " from DPTPLAN1 where VEND_CODE = '" & "VEND_CODE" & "' group by ITEM_CODE" & vbCrLf _
                & " union " & vbCrLf _
                & "Select POTORDR2.ITEM_CODE" & vbCrLf & sqlPO & " from POTORDR1,POTORDR2 where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO and VEND_CODE = '" & "VEND_CODE" & "' and POTORDR2.PO_STATUS = 'O' group by POTORDR2.ITEM_CODE" & vbCrLf _
                & ") group by ITEM_CODE"

            Dim LYP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)
            Dim LYP12 As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12)

            Dim LYP03 As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -3)

            ASCMAIN1.sql = "Select X.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.CARTON_PACK_QTY" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_VCOST" & vbCrLf _
                & ", ICTITEM1.COLLECTION_CODE, ICTCOLL1.COLLECTION_NAME, ICTCOLL1.HC_CODE, ICTCOLL0.HC_NAME" & vbCrLf _
                & ", X.FC00, X.FC01, X.FC02, X.FC03, X.FC04, X.FC05, X.FC06, X.FC07, X.FC08, X.FC09, X.FC10, X.FC11, X.FC12" & vbCrLf _
                & ", X.FC13, X.FC14, X.FC15, X.FC16, X.FC17, X.FC18, X.FC19, X.FC20, X.FC21, X.FC22, X.FC23, X.FC24, X.FC25, X.FC26" & vbCrLf _
                & ", Y.SHP_L03, Y.SHP_L12" & vbCrLf _
                & " from ICTITEM1, ICTCOLL1, ICTCOLL0, ICTCOSTC" & vbCrLf _
                & ", (" & sqlFC & ") X" & vbCrLf _
                & ", (Select ITEM_CODE" & vbCrLf _
                & ", SUM (CASE WHEN ORDR_YYYYPP_UPDATED >= '" & LYP03 & "' AND ORDR_YYYYPP_UPDATED <= '" & LYP & "' THEN ORDR_QTY_SHIP ELSE 0 END) SHP_L03" & vbCrLf _
                & ", SUM (CASE WHEN ORDR_YYYYPP_UPDATED >= '" & LYP12 & "' AND ORDR_YYYYPP_UPDATED <= '" & LYP & "' THEN ORDR_QTY_SHIP ELSE 0 END) SHP_L12" & vbCrLf _
                & " from SOTINVH2 where INV_TYPE = 'I' and ORDR_YYYYPP_UPDATED >= '" & LYP12 & "' group by ITEM_CODE) Y" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & "   and ICTCOLL0.HC_CODE (+) = ICTCOLL1.HC_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & "   and ICTCOSTC.ITEM_CODE (+) = X.ITEM_CODE" & vbCrLf _
                & "   and Y.ITEM_CODE (+) = X.ITEM_CODE"

            sqlPlans = Replace(ASCMAIN1.sql, sqlFC, sqlPlan)

            Create_TDA(.Tables.Add, "DPTITMFE", "**", 0, False)
            With .Tables("DPTITMFE").Columns
                .Add("CUSTNAME")
                .Add("CUSTDES")
                .Add("CUSTOMER_TYPE")
                .Add("CUSTOMER_TYPE_DES")
                .Add("YEAR")
                '.Add("SALES_LY", GetType(System.Decimal))
                '.Add("SALES_L3MO", GetType(System.Decimal))
            End With
            With .Tables("DPTITMFE")
                .Columns("CUSTNAME").DefaultValue = "512000011"
                .Columns("CUSTDES").DefaultValue = "AHAVA NORTH AMERICA LLC"
                .Columns("CUSTOMER_TYPE").DefaultValue = ""
                .Columns("CUSTOMER_TYPE_DES").DefaultValue = ""
                .Columns("YEAR").DefaultValue = Mid(ASCMAIN1.CYM, 1, 4)
            End With


            ' SATAUTHX

            ASCMAIN1.sql = "Select X.*, ARTCUST1.CUST_NAME from ARTCUST1, " & vbCrLf _
                & " (Select CUST_CODE, COUNT (*) STORE_COUNT from (" & vbCrLf _
                & " Select Distinct SATAUTH1.CUST_CODE, SATAUTH1.CUST_STORE_NO" & vbCrLf _
                & "  from SATAUTH1,ICTCOLL0" & vbCrLf _
                & " where ICTCOLL0.HC_CODE = SATAUTH1.HC_CODE" & vbCrLf _
                & "   and ICTCOLL0.BRAND_CODE = :PARM1" & vbCrLf _
                & "   and SATAUTH1.OPS_YYYYPP_CLOSED is Null" & vbCrLf _
                & "   and SATAUTH1.OPS_YYYYPP_OPENED is Not Null)" & vbCrLf _
                & " group by CUST_CODE) X" & vbCrLf _
                & " where ARTCUST1.CUST_CODE = X.CUST_CODE"
            Create_TDA(.Tables.Add, "SATAUTHX", "**", 0, False, "V", 1)
            dst.Tables("SATAUTHX").Columns("STORE_COUNT").DataType = GetType(System.Int32)

            Dim SQLB As String = ""
            Dim BUDY_TY As String = Mid(ASCMAIN1.CYP, 1, 4)
            Dim BUDY As String = BUDY_TY
            Dim BUDM As Integer = Val(Mid(ASCMAIN1.CYP, 5, 2)) - 1
            For i As Integer = 0 To 12
                BUDM += 1
                If BUDM > 12 Then
                    BUDM = 1
                    BUDY = Format(Val(BUDY_TY) + 1, "0000")
                End If
                SQLB &= ", SUM (DECODE(SATBUDW1.OPS_YYYY,'" & BUDY & "',SATBUDW1.WB_P" & Format(BUDM, "00") & ",0)) BUD_P" & Format(i, "00")
            Next

            If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then

                Dim SATBUDW1 As String = TAC.SOCMAIN1.Setup_Budgets_by_Customer

                ASCMAIN1.sql = "Select SATBUDW1.CUST_CODE" & SQLB & vbCrLf _
                    & " from " & SATBUDW1 & " SATBUDW1,ARTCUST1,SOTTCLS1,SOTMKTC1" & vbCrLf _
                    & " where SATBUDW1.OPS_YYYY >= '" & BUDY_TY & "'" & vbCrLf _
                    & "   and SATBUDW1.OPS_YYYY <= '" & BUDY & "'" & vbCrLf _
                    & "   and ARTCUST1.CUST_CODE = SATBUDW1.CUST_CODE" & vbCrLf _
                    & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                    & "   and SOTMKTC1.CUST_CODE (+) = SATBUDW1.CUST_CODE" & vbCrLf _
                    & "   and NVL(SOTMKTC1.MARKET_CODE,SOTTCLS1.MARKET_CODE) = :PARM1" & vbCrLf _
                    & "   and SATBUDW1.BRAND_CODE = :PARM2" & vbCrLf _
                    & " group by SATBUDW1.CUST_CODE"

            Else
                ASCMAIN1.sql = "Select SATBUDW1.CUST_CODE, ICTCOLL1.HC_CODE, SATBUDW1.ITEM_BASIC_PROMO" & SQLB & vbCrLf _
                    & " from SATBUDW1,ARTCUST1,SOTTCLS1,ICTCOLL1,SOTMKTC1" & vbCrLf _
                    & " where SATBUDW1.OPS_YYYY >= '" & BUDY_TY & "'" & vbCrLf _
                    & "   and SATBUDW1.OPS_YYYY <= '" & BUDY & "'" & vbCrLf _
                    & "   and ARTCUST1.CUST_CODE = SATBUDW1.CUST_CODE" & vbCrLf _
                    & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                    & "   and SOTMKTC1.CUST_CODE (+) = SATBUDW1.CUST_CODE" & vbCrLf _
                    & "   and NVL(SOTMKTC1.MARKET_CODE,SOTTCLS1.MARKET_CODE) = :PARM1" & vbCrLf _
                    & "   and ICTCOLL1.COLLECTION_CODE = SATBUDW1.COLLECTION_CODE" & vbCrLf _
                    & "   and ICTCOLL1.BRAND_CODE = :PARM2" & vbCrLf _
                    & "   and SATBUDW1.CUST_CODE not in (Select CUST_CODE from SOTCHAN1)" & vbCrLf _
                    & " group by SATBUDW1.CUST_CODE, ICTCOLL1.HC_CODE, SATBUDW1.ITEM_BASIC_PROMO"

            End If


            Create_TDA(.Tables.Add, "ICTCOLL1", "*", 0, False)
            With dst.Tables("ICTCOLL1")
                .Columns.Add("SEL")
                .Columns("SEL").DefaultValue = "1"
            End With
            Create_TDA(.Tables.Add, "SATBUDWX", "**", 0, False, "VV", 3)
            Create_TDA(.Tables.Add, "SOTPRIC2", "*", 1, False, , 2)



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
                .Add("QTY_LEFT", GetType(System.Int64), "ISNULL(QTY_ALLO,0)-ISNULL(ORDR_QTY_SHIP,0)-ISNULL(ORDR_QTY_PICK,0)")
                .Add("QTY_BAL", GetType(System.Int64), "IIF(QTY_LEFT>=0,QTY_LEFT,0)")
                .Add("QTY_OVER", GetType(System.Int64), "IIF(QTY_LEFT-ISNULL(ORDR_QTY_OPEN,0)>=0,0,ISNULL(ORDR_QTY_OPEN,0)-QTY_LEFT)")
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


        End With

        Fill_Records("ICTCOLL1")

        grdDPTITMFX.DataSource = dst.Tables("DPTITMFX")
        grdDPTITMFY.DataSource = dst.Tables("DPTITMFY")
        grdSATSLSC1.DataSource = dst.Tables("SATSLSC1")
        grdSATSLSI1.DataSource = dst.Tables("SATSLSI1_ALL")
        grdDPTITMFE.DataSource = dst.Tables("DPTITMFE")
        grdSATAUTHX.DataSource = dst.Tables("SATAUTHX")

        grdDPTITMFZ.DataSource = dst.Tables("DPTITMFZ")
        grdDPTITMFP.DataSource = dst.Tables("DPTITMFP")
        grdDPTITMFM.DataSource = dst.Tables("DPTITMFM")
        grdSATBUDWX.DataSource = dst.Tables("SATBUDWX")

        grdSOTALLOX.DataSource = dst.Tables("SOTALLOX")
        grdSOTALLOM.DataSource = dst.Tables("SOTALLOM")

        grdSOTALLO1.DataSource = dst.Tables("SOTALLO1")
        grdSOTALLO2.DataSource = dst.Tables("SOTALLO2")

        grdICTCOLL1.DataSource = dst.Tables("ICTCOLL1")
        grdDPTXLSX1.DataSource = dst.Tables("DPTXLSX1")

        Create_Summary(grdSATAUTHX, "CUST_CODE", "Count")
        Create_Summary(grdSATAUTHX, "STORE_COUNT")

        Create_Summary(grdSATBUDWX, "CUST_CODE", "Count")
        grdSATBUDWX.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        grdSATBUDWX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        grdSATBUDWX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        With grdSATBUDWX.DisplayLayout.Bands(0)
            For I As Integer = 0 To 12
                Dim COLUMN_NAME As String = "BUD_P" & Format(I, "00")
                .Columns(COLUMN_NAME).Width = 70
                .Columns(COLUMN_NAME).Format = "###,##0"
                Create_Summary(grdSATBUDWX, COLUMN_NAME)
                Dim YP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, I)
                Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", YP)
                .Columns(COLUMN_NAME).Header.Caption = Mid(rowGLTPARM2.Item("LEGEND"), 10, 6)

                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Color.LightGreen
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

            Next
            With .Columns("CUST_CODE")
                .Width = 100
                .Header.Caption = "Customer"
                .Header.Fixed = True
                .Header.Appearance.BackColor = Color.White
                .Header.Appearance.BackColor2 = Color.LightBlue
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            End With
            If ASCMAIN1.CLIENT = "INT" Then
                With .Columns("ITEM_BASIC_PROMO")
                    .Width = 40
                    .Header.Caption = "BP"
                    .Header.Fixed = True
                    .Header.Appearance.BackColor = Color.White
                    .Header.Appearance.BackColor2 = Color.LightBlue
                    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                End With

                With .Columns("HC_CODE")
                    .Width = 80
                    .Header.Caption = "HC"
                    .Header.Fixed = True
                    .Header.Appearance.BackColor = Color.White
                    .Header.Appearance.BackColor2 = Color.LightBlue
                    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                End With
            End If

        End With

        Create_Summary(grdDPTITMFE, "ITEM_CODE", "Count")
        With grdDPTITMFE.DisplayLayout.Bands(0)
            For I As Integer = 0 To 26 ' used to be 1 to 12
                Dim COLUMN_NAME As String = "FC" & Format(I, "00")
                .Columns(COLUMN_NAME).Width = 100
                .Columns(COLUMN_NAME).Format = "###,##0"
                Create_Summary(grdDPTITMFE, COLUMN_NAME)
                Dim YM As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYM, I)
                Dim MDY As String = Mid(YM, 5, 2) & "/01/" & Mid(YM, 1, 4)
                '.Columns("FC" & Format(I, "00")).Header.Caption = Format(CDate(MDY), "MMMM") & " -Rolling For."
                .Columns("FC" & Format(I, "00")).Header.Caption = Format(CDate(MDY), "MMM") & "'" & Format(CDate(MDY), "yy")
            Next
            For Each COLUMN_NAME As String In New String() {"SHP_L03", "SHP_L12"}
                .Columns(COLUMN_NAME).Width = 100
                .Columns(COLUMN_NAME).Format = "#####0"
                Create_Summary(grdDPTITMFE, COLUMN_NAME)
            Next
            .Columns("HC_CODE").Hidden = True
            .Columns("HC_NAME").Header.Caption = "Category"
            .Columns("COLLECTION_CODE").Hidden = True
            .Columns("COLLECTION_NAME").Header.Caption = "Subcategory"
            .Columns("ITEM_CODE").Header.Caption = "Cat. No"
            .Columns("ITEM_DESC").Header.Caption = "Part Description"
            .Columns("CARTON_PACK_QTY").Header.Caption = "MasterPack"
            .Columns("ITEM_COST_VCOST").Header.Caption = "FOBPrice"
            .Columns("SHP_L12").Header.Caption = "Tot. Sales Last Year"
            .Columns("SHP_L03").Header.Caption = "Tot. Sales Last 3 M."
            .Columns("ITEM_COST_VCOST").Header.Caption = "FOBPrice"

            .Columns("CUSTNAME").Header.Caption = "custname"
            .Columns("CUSTDES").Header.Caption = "custdes"
            .Columns("CUSTOMER_TYPE").Header.Caption = "customer type"
            .Columns("CUSTOMER_TYPE_DES").Header.Caption = "customer type des"
            .Columns("YEAR").Header.Caption = "Year"

            'For I As Integer = 1 To 12
            'Next
        End With

        With grdDPTITMFX.DisplayLayout.Bands(0)
            Dim G As UltraWinGrid.UltraGridGroup
            For Each COLUMN_NAME As String In New String() {"ITEM_CODE", "ITEM_DESC", "EXCL", "ITEM_CLASS_CODE", "COLLECTION_CODE", "ITEM_CATGY_CODE", "PROD_CODE", "ITEM_BASIC_PROMO", "COST_CATGY_CODE", "ITEM_RETAIL_PRICE", "ITEM_PRICE"}
                If COLUMN_NAME <> "ITEM_CODE" Then
                    .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.Beige
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                End If
                G = .Groups.Add(COLUMN_NAME, .Columns(COLUMN_NAME).Header.Caption)
                G.Width = .Columns(COLUMN_NAME).Width
                .Columns(COLUMN_NAME).Group = G
                If COLUMN_NAME <> "ITEM_DESC" And COLUMN_NAME <> "ITEM_CODE" Then
                    G.Hidden = True
                End If
                If ASCMAIN1.CLIENT = "INT" Then
                Else
                    If COLUMN_NAME = "EXCL" Then
                        G.Hidden = False
                        .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.Empty
                        .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                    End If
                End If
            Next
            .LevelCount = 6

            Dim GC As UltraWinGrid.UltraGridColumn

            'Dim i As Int32 = 0
            'G = .Groups.Add("TREND", "Trend")
            'For Each COLUMN_NAME As String In New String() _
            '{"TREND_ORIG", "TREND_FC", "TREND_TY", "TREND_PCT_TY", "TREND_LY", "TREND_PCT_LY"}
            '    GC = .Columns.Add(COLUMN_NAME)
            '    GC.Group = G
            '    GC.Level = i
            '    i += 1
            '    GC.CellAppearance.BackColor = Color.LightBlue
            'Next
            'G.Width = 30

            G = .Groups.Add("TREND", "Trend ->")
            For j As Integer = 0 To 5
                Dim COLUMN_NAME As String = "DATA_TREND" & Format(j, "0")
                GC = .Columns.Add(COLUMN_NAME)
                GC.Group = G
                GC.Level = j
                GC.CellAppearance.BackColor = Color.LightBlue
            Next
            G.Width = 70


            For P As Integer = 6 To 0 Step -1
                G = .Groups.Add("T" & Format(P, "00"))

                .Columns("TOU" & Format(P, "00")).Group = G
                .Columns("TOU" & Format(P, "00")).Level = 1
                .Columns("TOS" & Format(P, "00")).Group = G
                .Columns("TOS" & Format(P, "00")).Level = 1
                .Columns("TPU" & Format(P, "00")).Group = G
                .Columns("TPU" & Format(P, "00")).Level = 0
                .Columns("TPS" & Format(P, "00")).Group = G
                .Columns("TPS" & Format(P, "00")).Level = 0

                .Columns("TOU" & Format(P, "00")).CellAppearance.ForeColor = Drawing.Color.OrangeRed
                .Columns("TOS" & Format(P, "00")).CellAppearance.ForeColor = Drawing.Color.OrangeRed

                .Columns("TPU" & Format(P, "00")).CellAppearance.ForeColor = Drawing.Color.HotPink
                .Columns("TPS" & Format(P, "00")).CellAppearance.ForeColor = Drawing.Color.HotPink

                .Columns("TU" & Format(P, "00")).Group = G
                .Columns("TU" & Format(P, "00")).Level = 2
                .Columns("TS" & Format(P, "00")).Group = G
                .Columns("TS" & Format(P, "00")).Level = 2
                .Columns("TU" & Format(P + 12, "00")).Group = G
                .Columns("TU" & Format(P + 12, "00")).Level = 4
                .Columns("TS" & Format(P + 12, "00")).Group = G
                .Columns("TS" & Format(P + 12, "00")).Level = 4

                .Columns("TU" & Format(P, "00")).CellAppearance.BackColor = Color.Yellow
                .Columns("TS" & Format(P, "00")).CellAppearance.BackColor = Color.Yellow
                .Columns("TU" & Format(P + 12, "00")).CellAppearance.BackColor = Color.LightCyan
                .Columns("TS" & Format(P + 12, "00")).CellAppearance.BackColor = Color.LightCyan

                .Columns("TAPU" & Format(P, "00")).Group = G
                .Columns("TAPU" & Format(P, "00")).Level = 3
                .Columns("TAPU" & Format(P, "00")).Format = "##0.0"
                .Columns("TAPS" & Format(P, "00")).Group = G
                .Columns("TAPS" & Format(P, "00")).Level = 3
                .Columns("TAPS" & Format(P, "00")).Format = "##0.0"
                .Columns("TALU" & Format(P, "00")).Group = G
                .Columns("TALU" & Format(P, "00")).Level = 5
                .Columns("TALU" & Format(P, "00")).Format = "##0.0"
                .Columns("TALS" & Format(P, "00")).Group = G
                .Columns("TALS" & Format(P, "00")).Level = 5
                .Columns("TALS" & Format(P, "00")).Format = "##0.0"

                G.Header.Caption = YPP(P, 1)
                G.Width = 70
            Next

            G = .Groups.Add("DATA", "Data ->")
            For j As Integer = 0 To 5
                Dim COLUMN_NAME As String = "DATA" & Format(j, "0")
                GC = .Columns.Add(COLUMN_NAME)
                GC.CellActivation = UltraWinGrid.Activation.NoEdit
                GC.Group = G
                GC.Level = j
                GC.CellAppearance.BackColor = Color.LightBlue
            Next
            G.Width = 70


            G = .Groups.Add("TOTAL")
            G.Header.Caption = "Total"
            .Columns("UFCTOT").Group = G
            .Columns("UFCTOT").Level = 0
            .Columns("SFCTOT").Group = G
            .Columns("SFCTOT").Level = 0

            For P As Integer = -1 To FCMAX
                Dim SFX = "PD"
                If P > -1 Then SFX = Format(P, "00")
                G = .Groups.Add("FC" & SFX)
                .Columns("UFC" & SFX).Group = G
                .Columns("UFC" & SFX).Level = 0
                .Columns("SFC" & SFX).Group = G
                .Columns("SFC" & SFX).Level = 0
                If P > -1 Then

                    If P <= 12 Then
                        .Columns("S" & SFX).Group = G
                        .Columns("S" & SFX).Level = 2
                        .Columns("U" & SFX).Group = G
                        .Columns("U" & SFX).Level = 2
                        .Columns("RS" & SFX).Group = G
                        .Columns("RS" & SFX).Level = 3
                        .Columns("RU" & SFX).Group = G
                        .Columns("RU" & SFX).Level = 3

                        .Columns("S" & SFX).CellAppearance.BackColor = Color.LightCyan
                        .Columns("U" & SFX).CellAppearance.BackColor = Color.LightCyan

                        .Columns("RS" & SFX).CellAppearance.BackColor = Color.LightGreen
                        .Columns("RU" & SFX).CellAppearance.BackColor = Color.LightGreen

                        .Columns("VS" & SFX).Group = G
                        .Columns("VS" & SFX).Level = 4
                        .Columns("VU" & SFX).Group = G
                        .Columns("VU" & SFX).Level = 4
                        .Columns("PS" & SFX).Group = G
                        .Columns("PS" & SFX).Level = 5
                        .Columns("PS" & SFX).Format = "##0.0"
                        .Columns("PU" & SFX).Group = G
                        .Columns("PU" & SFX).Level = 5
                        .Columns("PU" & SFX).Format = "##0.0"

                    End If

                    '.Columns("U" & SFX).Format = "###,##0"
                    '.Columns("S" & SFX).Format = "###,##0"


                    G.Header.Caption = YPF(P, 1)
                Else
                    G.Header.Caption = "Past Due"
                End If
                G.Width = 70
                If P = FCMAX Then
                    G.Hidden = True
                End If
            Next
            .ColHeadersVisible = False
            '.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            '.Override.AllowUpdate = DefaultableBoolean.False
            '.Override.AllowDelete = DefaultableBoolean.False
            '      .Groups(1).Hidden = True
        End With


        grdDPTITMFY.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.Select
        grdDPTITMFY.DisplayLayout.Override.RowSelectors = DefaultableBoolean.Default
        grdDPTITMFY.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement


        With grdDPTITMFY.DisplayLayout.Bands(0)
            .Columns("ITEM_CODE").Width = 80
            .Columns("ITEM_DESC").Width = 170

            For I As Integer = 0 To 26 ' used to be 1 to 12
                Dim COLUMN_NAME As String = "IFC" & Format(I, "00")
                .Columns(COLUMN_NAME).Width = 70
                .Columns(COLUMN_NAME).Format = "##0.00"
                .Columns(COLUMN_NAME).Hidden = True
                Create_Summary(grdDPTITMFY, COLUMN_NAME)
                Dim YM As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYM, I)
                Dim MDY As String = Mid(YM, 5, 2) & "/01/" & Mid(YM, 1, 4)

                .Columns("IFC" & Format(I, "00")).Header.Caption = Format(CDate(MDY), "MMM") & "'" & Format(CDate(MDY), "yy")

                COLUMN_NAME = "IFCA" & Format(I, "00")
                .Columns(COLUMN_NAME).Width = 70
                .Columns(COLUMN_NAME).Format = "##0.00"
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                Create_Summary(grdDPTITMFY, COLUMN_NAME)


                .Columns("IFCA" & Format(I, "00")).Header.Caption = Format(CDate(MDY), "MMM") & "'" & Format(CDate(MDY), "yy")

            Next
            For I As Integer = 12 To 9 Step -1 ' Retail Sales Last 3 Months
                Dim COLUMN_NAME As String = "RS" & Format(I, "00")
                .Columns(COLUMN_NAME).Width = 95
                .Columns(COLUMN_NAME).Format = "##0.00"
                .Columns(COLUMN_NAME).Hidden = True
                Create_Summary(grdDPTITMFY, COLUMN_NAME)
                Dim YM As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYM, I - 12)
                Dim MDY As String = Mid(YM, 5, 2) & "/01/" & Mid(YM, 1, 4)
                If I = 12 Then
                    .Columns("RS12").Header.Caption = "RS% Avg, Past 3 Mth"
                Else
                    .Columns("RS" & Format(I, "00")).Header.Caption = "RS% " & Format(CDate(MDY), "MMM") & "'" & Format(CDate(MDY), "yy")
                End If
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Color.LightGreen
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next

            .Columns("SFCTOT").Header.Caption = "Total Sales"
            .Columns("BASE_IP").Header.Caption = "BaseLine %"
            Create_Summary(grdDPTITMFY, "SFCTOT")
            Create_Summary(grdDPTITMFY, "BASE_IP")

            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No

        End With


        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSATSLSC1, grdSATSLSI1}

            With grd.DisplayLayout.Bands(0)
                Dim G As UltraWinGrid.UltraGridGroup
                If grd.Name = "grdSATSLSC1" Then
                    G = .Groups.Add("CUSTOMER", "Customer")
                    .Columns("CUST_CODE").Group = G
                    .Columns("CUST_NAME").Group = G
                Else
                    G = .Groups.Add("ITEM", "Item")
                    .Columns("ITEM_CODE").Group = G
                    .Columns("ITEM_DESC").Group = G
                End If
                G.Header.Fixed = True

                .LevelCount = 2

                G = .Groups.Add("DATA", "Data")
                With .Columns.Add("DATA0")
                    .Group = G
                    .Level = 0
                End With
                With .Columns.Add("DATA1")
                    .Group = G
                    .Level = 1
                End With
                G.Width = 70
                G.CellAppearance.BackColor = Color.LightBlue
                G.Header.Fixed = True

                For P As Integer = CSMAX To 0 Step -1
                    G = .Groups.Add("G" & Format(P, "00"))
                    .Columns("U" & Format(P, "00")).Group = G
                    .Columns("U" & Format(P, "00")).Level = 0
                    .Columns("S" & Format(P, "00")).Group = G
                    .Columns("S" & Format(P, "00")).Level = 0
                    .Columns("RU" & Format(P, "00")).Group = G
                    .Columns("RU" & Format(P, "00")).Level = 1
                    .Columns("RS" & Format(P, "00")).Group = G
                    .Columns("RS" & Format(P, "00")).Level = 1

                    G.Header.Caption = YPP(P, 1)
                    G.Header.Appearance.TextHAlign = HAlign.Right
                    G.Width = 70
                Next
                .ColHeadersVisible = False
                .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                .Override.AllowUpdate = DefaultableBoolean.False
                .Override.AllowDelete = DefaultableBoolean.False
            End With
        Next

        Create_Summary(grdDPTITMFX, "ITEM_CODE", "Count")
        Create_Summary(grdDPTITMFX, "EXCL")

        Create_Summary(grdDPTITMFY, "ITEM_CODE", "Count")


        Create_Summary(grdSOTALLO2, New String() {"QTY_ALLO", "ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "QTY_LEFT", "QTY_BAL", "QTY_OVER"})


        For P = -2 To FCMAX
            Dim COLUMN_NAME As String = IIf(P = -2, "UFCTOT", IIf(P = -1, "UFCPD", "UFC" & Format(P, "00")))
            Create_Summary(grdDPTITMFX, COLUMN_NAME)
            COLUMN_NAME = Replace(COLUMN_NAME, "UFC", "SFC")
            grdDPTITMFX.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Level = 0
            Create_Summary(grdDPTITMFX, COLUMN_NAME, , , "###,##0")
            grdDPTITMFX.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Level = 0
        Next

        'Call Create_Summary(grdSATSLSC1, "CUST_CODE", "Count")
        'For P = 0 To CSMAX
        '    Call Create_Summary(grdSATSLSC1, "U" & Format(P, "00"), , , "###,##0")
        '    Call Create_Summary(grdSATSLSC1, "S" & Format(P, "00"), , , "###,##0")
        'Next

        With grdDPTITMFX.DisplayLayout.Bands("DPTITMFX")
            .Groups("ITEM_CODE").Header.Fixed = True
        End With



        grdDPTITMFY.DisplayLayout.UseFixedHeaders = True
        With grdDPTITMFY.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ITEM_CODE", "ITEM_DESC", "BASE_IP"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        ' DGJ








        optITEM_CATGY_CODE.ValueList = ASCMAIN1.ValueListFor("ITEM_CATGY_CODE", , New String() {":", "ALL:All"})

        For Each COLUMN_NAME As String In New String() _
        {"ITEM_CODE", "ITEM_DESC", "COLLECTION_CODE", "COLLECTION_CODE"}
            With grdDPTITMFX.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                .CellAppearance.BackColor = Drawing.Color.Beige
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End With
        Next

        For Each COLUMN_NAME As String In New String() _
        {"ITEM_CODE", "ITEM_DESC", "RS09", "RS10", "RS11", "RS12"}
            With grdDPTITMFY.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                .CellAppearance.BackColor = Drawing.Color.Beige
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End With
        Next




        With grdDPTITMFZ.DisplayLayout.Bands(0)

            With .Columns("DATA_DESC")
                .CellActivation = UltraWinGrid.Activation.NoEdit
                .Width = 100
                With .Header.Appearance
                    .BackColor2 = Color.LightBlue
                    .BackColor = Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                End With
                .Header.Caption = ""
            End With


            For P As Integer = 0 To FCMAX
                With .Columns("P" & Format(P, "00"))

                    With .Header.Appearance
                        .BackColor2 = Color.Orange
                        .BackColor = Color.White
                        .BackGradientStyle = GradientStyle.ForwardDiagonal
                    End With
                    .Format = "#.00"
                    .CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Header.Appearance.TextHAlign = HAlign.Right
                    .Width = 80

                    .Header.Caption = YPF(P, 1)
                End With
            Next

            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False
        End With



        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdDPTITMFP, grdSOTALLOX}
            With grd.DisplayLayout.Bands(0)

                .LevelCount = 2

                Dim G As UltraWinGrid.UltraGridGroup
                G = .Groups.Add("DATA_DESC", "Data")
                If grd.Name = "grdDPTITMFP" Then
                    .Columns("DATA_DESC").Group = G
                    .Columns("DATA_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    With .Columns("ITEM_CODE")
                        .Group = G
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                        .Level = 0
                        .Width = 120
                    End With

                    With .Columns("SEL")
                        .Group = G
                        .Style = UltraWinGrid.ColumnStyle.CheckBox
                        .CellAppearance.TextHAlign = HAlign.Center
                        .Header.Appearance.TextHAlign = HAlign.Center
                        .CellActivation = UltraWinGrid.Activation.AllowEdit
                        .Level = 0
                        .Width = 30
                    End With

                    With .Columns("ITEM_DESC")
                        .Group = G
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                        .Level = 1
                    End With
                End If

                If grd.Name = "grdSOTALLOX" Then
                    G.Width = 150
                    .Columns("ITEM_CODE").Width = 120
                Else
                    G.Width = 100
                End If
                With G.Header.Appearance
                    .BackColor2 = Color.LightBlue
                    .BackColor = Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                End With

                G = .Groups.Add("DATA_TYPE", "Type")
                With .Columns.Add("DATA_TYPE1")
                    .Group = G
                    .Level = 0
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                End With
                With .Columns.Add("DATA_TYPE2")
                    .Group = G
                    .Level = 1
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                End With
                G.Width = 40
                With G.Header.Appearance
                    .BackColor2 = Color.LightBlue
                    .BackColor = Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                End With

                G = .Groups.Add("TOTALS", "Totals")
                With .Columns("A_TOTAL")
                    .Group = G
                    .Level = 0
                    .Format = "#,##0"
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                End With
                With .Columns("P_TOTAL")
                    .Group = G
                    .Level = 1
                    .Format = IIf(grd.Name = "grdDPTITMFP", "#.00", "#,##0")
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                End With
                G.Width = 90
                'G.CellAppearance.BackColor = Color.LightGreen
                With G.Header.Appearance
                    .BackColor2 = Color.LightGreen
                    .BackColor = Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                End With

                For P As Integer = 0 To FCMAX
                    G = .Groups.Add("M" & Format(P, "00"), YPF(P, 1))
                    With G.Header.Appearance
                        .BackColor2 = Color.Orange
                        .BackColor = Color.White
                        .BackGradientStyle = GradientStyle.ForwardDiagonal
                    End With
                    With .Columns("A" & Format(P, "00"))
                        .Group = G
                        .Level = 0
                        .Format = "#,##0"
                        .CellActivation = IIf(grd.Name = "grdDPTITMFP", UltraWinGrid.Activation.AllowEdit, UltraWinGrid.Activation.NoEdit)
                    End With
                    With .Columns("P" & Format(P, "00"))
                        .Group = G
                        .Level = 1
                        .Format = IIf(grd.Name = "grdDPTITMFP", "#.00", "#,##0")
                        .CellActivation = IIf(grd.Name = "grdDPTITMFP", UltraWinGrid.Activation.AllowEdit, UltraWinGrid.Activation.NoEdit)
                    End With
                    'G.Header.Caption = YPF(P, 1)
                    G.Header.Appearance.TextHAlign = HAlign.Right
                    G.Width = 80
                Next
                .ColHeadersVisible = False
                .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                .Override.AllowUpdate = DefaultableBoolean.True
                .Override.AllowDelete = DefaultableBoolean.False
            End With
        Next

        ASCMAIN1.grdInitializeLayout(grdSOTALLOX)
        grdSOTALLOX.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText

        Dim c As UltraWinGrid.UltraGridColumn

        With grdDPTITMFX.DisplayLayout.Bands(0)

            c = .Columns("ITEM_CODE")
            c.Header.Caption = "Item"
            c.Hidden = False
            c.Width = 80
            c.ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
            c.Style = UltraWinGrid.ColumnStyle.EditButton
            'c.CellActivation = UltraWinGrid.Activation.NoEdit
            c.CellAppearance.BackColor = Color.Beige


            For P = -1 To FCMAX
                Dim COLUMN_NAME As String = "UFCPD"
                Dim COLUMN_CAPTION As String = "PastDue"
                c = .Columns(COLUMN_NAME)
                c.Hidden = False
                c.Width = 60
                c.CellActivation = UltraWinGrid.Activation.AllowEdit
                c.Header.Caption = COLUMN_CAPTION
                c.CellAppearance.ForeColor = Drawing.Color.HotPink ' Gold ' CornflowerBlue ' RosyBrown ' CadetBlue
                'c.CellAppearance.BackColor = Drawing.Color.HotPink ' Gold ' CornflowerBlue ' RosyBrown ' CadetBlue
                'c.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If ROWs("DPTPARM1").Item("DP_PARM_FC_FREEZE") & "" = "1" Then
                    If P <= Val(ROWs("DPTPARM1").Item("DP_PARM_FC_FREEZE_MOS") & "") Then
                        c.CellAppearance.ForeColor = Color.Red
                        c.CellAppearance.BackColor = Color.Beige
                        c.CellActivation = UltraWinGrid.Activation.NoEdit
                    End If
                End If

                If P >= 0 And c.CellActivation = UltraWinGrid.Activation.AllowEdit Then
                    COLUMN_NAME = "UFC" & Format(P, "00")
                    COLUMN_CAPTION = YPF(P, 1)
                    dst.Tables("DPTITMFM").Rows.Add(New String() {Format(P, "00"), COLUMN_CAPTION})
                    dst.Tables("SOTALLOM").Rows.Add(New String() {Format(P, "00"), COLUMN_CAPTION})

                End If
            Next
            ASCMAIN1.grdInitializeLayout(grdDPTITMFX)
        End With

        ASCMAIN1.Add_Value_List(grdDPTITMFX, "ITEM_CATGY_CODE")

        SplitContainer3.Panel1Collapsed = True
        MakeTransparent(chk3MRS)
        MakeTransparent(chk3MRSA)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("MARKET_CODE")
                Validate_Code("BRAND_CODE")

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("DPTITMFX", Absx1.txtFor("MARKET_CODE").Text & ":" & Absx1.txtFor("BRAND_CODE").Text) Then
                        Exit Sub
                    End If

                    If Not ASCMAIN1.Logical_Open("DPTITMFX", "*" & ":" & Absx1.txtFor("BRAND_CODE").Text) Then
                        Exit Sub
                    End If

                End If

            Case "Update"


                If BASE_IP_PERC_CHANGE = True Then
                    If MsgBox("Baseline % Modified, Do you want to disregard these changes?" & vbCr & vbCr & "Base IP has not been applied to Forecast Grid", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        '   BASE_IP_PERC_CHANGE = False
                        Exit Sub
                    End If
                    BASE_IP_PERC_CHANGE = False
                End If


                'If Absx1.txtFor("STAX_CODE").Text = "" And Val(Absx1.numFor("INV_STAX").Value & "") <> 0 Then
                '    EMsg &= vbCr & "You Must Specify a Tax Code"
                'End If

            Case "Clear Past Due"
                If MsgBox("This will clear the Past Due column for all Items on display" & vbCr & vbCr & "Continue with this Update?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
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
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Import from Excel"
                Import_from_Excel()

            Case "Update"
                'If ASCMAIN1.CLIENT = "AHA" Then
                '    Export_to_Excel(grdDPTITMFX, True)
                'End If
                Call Update_Record()
                Call Mode_Settings(False)

            Case "Cancel", "Done"
                Call Mode_Settings(False)

            Case "Clear Past Due"
                Call Clear_Past_Due()

            Case "Export Forecasts"
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Building Worktable")
                Fill_Records("DPTITMFE")
                grdDPTITMFE.Visible = True
                grdDPTITMFE.Text = "Forecasts and Shipments History"
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

                Sort_grdColumns(grdDPTITMFE, "ITEM_CODE")




            Case "Export Plans"
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Building Worktable")
                Dim VEND_CODE As String = "DEADSEALAB"
                ASCMAIN1.sql = Replace(sqlPlans, "'VEND_CODE'", "'" & VEND_CODE & "'")
                Fill_Records("DPTITMFE", "", True, ASCMAIN1.sql)
                grdDPTITMFE.Visible = True
                grdDPTITMFE.Text = "Generated Plans for " & VEND_CODE
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

                Sort_grdColumns(grdDPTITMFE, "ITEM_CODE")

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Import from Excel").Settings.Enabled = iScreenMode
                    .Items("Clear Past Due").Settings.Enabled = iScreenMode

                    If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
                        .Items("Export Forecasts").Visible = Not ScreenMode
                        .Items("Export Plans").Visible = Not ScreenMode
                    Else
                        .Items("Export Forecasts").Visible = False
                        .Items("Export Plans").Visible = False
                    End If

                End With

                .Groups("Display Options").Visible = ScreenMode

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdDPTITMFE.Visible = False
        SplitContainer1.Visible = ScreenMode

        If ScreenMode Then
            optITEM_CATGY_CODE.Visible = True
            optITEM_CATGY_CODE.Visible = False

            tabDetails.Tabs("Allocations").Visible = (Absx1.txtFor("MARKET_CODE").Text = DP_PARM_ALLOC_MARKET)
            pnlUpload.Visible = False
        Else
            Clear_Record()
            tab.Tabs("XLS Upload").Visible = True
            pnlUpload.Visible = True
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"DPTITMFX", "DPTITMFH", "DPTITMF1", "SATSLSC1", "SATSLSI1", "SOTPRIC2",
            "SOTALLO1", "SOTALLO2", "DPTXLSX1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Absx1.txtFor("MARKET_CODE").Text = ""
        Absx1.txtFor("BRAND_CODE").Text = ""
        SplitContainer6.Visible = False

        Fill_Records("ICTCOLL1")

        optUS.Value = "U"
        BASE_IP_PERC_CHANGE = False
        chkHistory.Enabled = False

        Set_US()
        Set_History()
        Set_Trend()
        Set_FullScreen()

        Toggle_Prorate(False)
        Toggle_Allocate(False)

        grdDPTITMFM.Selected.Rows.Clear()
        grdSOTALLOM.Selected.Rows.Clear()
    End Sub

    Sub Load_Record()

        Dim dt As Date = Now

        Debug.Print("Start: " & Now.Subtract(dt).Seconds)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        Debug.Print("Truncating: " & Now.Subtract(dt).Seconds)
        ASCDATA1.ExecuteSQL("Truncate Table " & ICTITEMS)

        Debug.Print("Inserting: " & Now.Subtract(dt).Seconds)
        ASCMAIN1.sql = "Insert into " & ICTITEMS _
        & " Select ITEM_CODE from ICTITEM1 where COLLECTION_CODE in " _
        & "(Select COLLECTION_CODE FROM ICTCOLL1 " _
            & " where BRAND_CODE = '" & Absx1.txtFor("BRAND_CODE").Text & "')"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        Debug.Print("Disable Constraints: " & Now.Subtract(dt).Seconds)
        EnforceConstraints(False)

        Dim MARKET_CODE As String = HFs("MARKET_CODE")
        Dim rowSOTMKTC1 As DataRow = LookUp("SOTMKTC1", MARKET_CODE)
        PRICE_CLASS_CODE = rowSOTMKTC1.Item("PRICE_CLASS_CODE") & ""
        CUST_CODE = rowSOTMKTC1.Item("CUST_CODE") & ""

        Debug.Print("Setting up Memory: " & Now.Subtract(dt).Seconds)
        If CUST_CODE <> "" Then
            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
            PRICE_CLASS_CODE = rowARTCUST1.Item("PRICE_CLASS_CODE") & ""
            PRICE_LIST_CODE = rowARTCUST1.Item("PRICE_LIST_CODE") & ""

            If PRICE_LIST_CODE <> "" Then
                Fill_Records("SOTPRIC2", PRICE_LIST_CODE)
            End If

            If PRICE_CLASS_CODE = "" Then
                CUST_DISC_PCT = 0
                If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
                    CUST_DISC_PCT = 0 ' 50
                End If
            Else
                Dim rowSOTPCLS1 As DataRow = LookUp("SOTPCLS1", PRICE_CLASS_CODE)
                CUST_DISC_PCT = Val(rowSOTPCLS1.Item("PRICE_BASE_DPCT") & "")
            End If
        Else
            If PRICE_CLASS_CODE = "" Then
                CUST_DISC_PCT = 40
                If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
                    CUST_DISC_PCT = 50
                End If
            Else
                Dim rowSOTPCLS1 As DataRow = LookUp("SOTPCLS1", PRICE_CLASS_CODE)
                CUST_DISC_PCT = Val(rowSOTPCLS1.Item("PRICE_BASE_DPCT") & "")
            End If
        End If

        dst.Tables("DPTITMFX").Columns("ITEM_WHOLESALE_PRICE").Expression = "ITEM_RETAIL_PRICE * " & CStr(100 - CUST_DISC_PCT) & " / 100"
        dst.Tables("DPTITMFY").Columns("ITEM_WHOLESALE_PRICE").Expression = "ITEM_RETAIL_PRICE * " & CStr(100 - CUST_DISC_PCT) & " / 100"

        Debug.Print("Filling SATAUTHX: " & Now.Subtract(dt).Seconds)
        Fill_Records("SATAUTHX", Absx1.txtFor("BRAND_CODE").Text)
        Sort_grdColumns(grdSATAUTHX, "CUST_CODE")

        '& "   and NVL(SOTMKTC1.MARKET_CODE,SOTTCLS1.MARKET_CODE) = ''" & vbCrLf _

        Debug.Print("Building SATSLSII: " & Now.Subtract(dt).Seconds)
        ASCDATA1.ExecuteSQL("Truncate Table " & SATSLSII)
        ASCDATA1.ExecuteSQL("Insert into " & SATSLSII & " " & Replace(sqlSATSLSII, "NVL(SOTMKTC1.MARKET_CODE,SOTTCLS1.MARKET_CODE) = ''", "NVL(SOTMKTC1.MARKET_CODE,SOTTCLS1.MARKET_CODE) = '" & MARKET_CODE & "'"))

        Debug.Print("Building SATSLSIT: " & Now.Subtract(dt).Seconds)
        ASCDATA1.ExecuteSQL("Truncate Table " & SATSLSIT)

        If chkIncludeRetailData.Checked Then
            ASCDATA1.ExecuteSQL("Insert into " & SATSLSIT & " " & Replace(sqlSATSLSIT, "NVL(SOTMKTC1.MARKET_CODE,SOTTCLS1.MARKET_CODE) = ''", "NVL(SOTMKTC1.MARKET_CODE,SOTTCLS1.MARKET_CODE) = '" & MARKET_CODE & "'"))
        End If

        Debug.Print("Filling SATSLSII: " & Now.Subtract(dt).Seconds)
        ASCDATA1.ExecuteSQL("Insert into " & SATSLSII & " (ITEM_CODE) Select ITEM_CODE FROM " & SATSLSIT & " Minus Select ITEM_CODE from " & SATSLSII)
        Fill_Records("SATSLSI1")

        Debug.Print("Merge Tables: " & Now.Subtract(dt).Seconds)
        Sort_grdColumns(grdSATSLSI1, "ITEM_CODE")
        'Fill_Records("RSTRETL1")
        dst.Tables("SATSLSI1_ALL").Rows.Clear()
        dst.Tables("SATSLSI1_ALL").Merge(dst.Tables("SATSLSI1"))

        Debug.Print("Building DPTITMFH: " & Now.Subtract(dt).Seconds)
        ASCDATA1.ExecuteSQL("Truncate Table " & DPTITMFH)
        ASCDATA1.ExecuteSQL("Insert into " & DPTITMFH & " " & sqlDPTITMFH)
        Fill_Records("DPTITMFH")

        Debug.Print("Building DPTITMFO: " & Now.Subtract(dt).Seconds)
        ASCDATA1.ExecuteSQL("Truncate Table " & DPTITMFO)
        ASCDATA1.ExecuteSQL("Insert into " & DPTITMFO & " " & sqlDPTITMFO)
        Fill_Records("DPTITMFO")

        Fill_Records("DPTITMF2")

        Debug.Print("Building DPTITMFX: " & Now.Subtract(dt).Seconds)
        Fill_Records("DPTITMFX", HFs("MARKET_CODE"))
        Sort_grdColumns(grdDPTITMFX, "ITEM_CODE")

        Debug.Print("Loop DPTITMFX -> DPTITMFH: " & Now.Subtract(dt).Seconds)
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("DPTITMFH"), "ITEM_CODE").Select("")
            Dim ITEM_CODE As String = row.Item(0)
            If dst.Tables("DPTITMFX").Rows.Find(ITEM_CODE) Is Nothing Then
                ASCDATA1.DeleteRows(dst.Tables("DPTITMFH"), "ITEM_CODE = '" & ITEM_CODE & "'")
            End If
        Next

        Debug.Print("Loop DPTITMFX -> DPTITMFO: " & Now.Subtract(dt).Seconds)
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("DPTITMFO"), "ITEM_CODE").Select("")
            Dim ITEM_CODE As String = row.Item(0)
            If dst.Tables("DPTITMFX").Rows.Find(ITEM_CODE) Is Nothing Then
                ASCDATA1.DeleteRows(dst.Tables("DPTITMFO"), "ITEM_CODE = '" & ITEM_CODE & "'")
            Else
                '  Fill_Records("DPTITMF2")
            End If

            ' REM ADD DPTITMF2 DGJ
        Next



        Debug.Print("Loop SATSLSI1 -> DPTITMFX: " & Now.Subtract(dt).Seconds)
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SATSLSI1"), "ITEM_CODE").Select("")
            Dim ITEM_CODE As String = row.Item(0)
            If dst.Tables("DPTITMFX").Rows.Find(ITEM_CODE) Is Nothing Then
                ASCDATA1.DeleteRows(dst.Tables("SATSLSI1"), "ITEM_CODE = '" & ITEM_CODE & "'")
            End If
        Next

        Debug.Print("Loop Price List: " & Now.Subtract(dt).Seconds)
        If PRICE_LIST_CODE <> "" Then
            dst.Tables("DPTITMFX").Columns("ITEM_WHOLESALE_PRICE").Expression = ""
            dst.Tables("DPTITMFX").Columns("ITEM_WHOLESALE_PRICE").ReadOnly = False
            For Each rowDPTITMFX As DataRow In dst.Tables("DPTITMFX").Select("")
                Dim ITEM_CODE As String = rowDPTITMFX.Item("ITEM_CODE")
                Dim rowSOTPRIC2 As DataRow = dst.Tables("SOTPRIC2").Rows.Find(New String() {PRICE_LIST_CODE, ITEM_CODE})
                If rowSOTPRIC2 IsNot Nothing Then
                    rowDPTITMFX.Item("ITEM_WHOLESALE_PRICE") = rowSOTPRIC2.Item("ITEM_PRICE")
                Else
                    rowDPTITMFX.Item("ITEM_WHOLESALE_PRICE") = Val(rowDPTITMFX.Item("ITEM_RETAIL_PRICE") & "") * (100 - CUST_DISC_PCT) / 100
                End If
            Next
        End If
        ASCMAIN1.sql = "Select * from DPTITMB1 where OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
        Fill_Records("DPTITMB1", "", False, ASCMAIN1.sql)


        ' Fill_Records("DPTITMB1", ASCMAIN1.CYP)
        For Each rowDPTITMFX As DataRow In dst.Tables("DPTITMFX").Select("")
            Dim ITEM_CODE As String = rowDPTITMFX.Item("ITEM_CODE")
            Dim PERIOD As String = ASCMAIN1.CYP
            Dim rowDPTITMB1 As DataRow = dst.Tables("DPTITMB1").Rows.Find(New String() {PERIOD, ITEM_CODE})
            If rowDPTITMB1 IsNot Nothing Then
                rowDPTITMFX.Item("BASE_IP") = Val(rowDPTITMB1.Item("BASE_IP") & "")
            Else
                rowDPTITMFX.Item("BASE_IP") = Null
            End If
        Next
        Debug.Print("Fill SATBUDWX: " & Now.Subtract(dt).Seconds)
        Fill_Records("SATBUDWX", New String() {HFs("MARKET_CODE"), HFs("BRAND_CODE")})
        Sort_grdColumns(grdSATBUDWX, "CUST_CODE")

        EnforceConstraints(True)

        Dim BRAND_CODE As String = Absx1.txtFor("BRAND_CODE").Text
        Dim dvw As DataView = DirectCast(grdICTCOLL1.DataSource, DataTable).DefaultView
        dvw.RowFilter = $"BRAND_CODE = '{BRAND_CODE}'"
        Sort_grdColumns(grdICTCOLL1, "COLLECTION_CODE")



        optITEM_CATGY_CODE.Value = "E"

        Debug.Print("Set up Drop Downs: " & Now.Subtract(dt).Seconds)
        ASCMAIN1.sql = "Select COLLECTION_CODE, COLLECTION_NAME from ICTCOLL1 where BRAND_CODE = '" & Absx1.txtFor("BRAND_CODE").Text & "' ORDER BY COLLECTION_CODE"
        cbeCOLLECTION_CODE.DataSource = ASCDATA1.GetDataTable
        cbeCOLLECTION_CODE.Value = cbeCOLLECTION_CODE.Items(0)

        Setup_grdDPTITMFX()

        optCOLLECTION_CODE.Value = "A"
        optITEM_CATGY_CODE.Value = "ALL"

        optBP.Value = "*"
        optSN.Value = "*"

        'With grdDPTITMFX.DisplayLayout.Bands(0)
        '    .Columns("DATA0").Hidden = True
        'End With


        Debug.Print("Finish: " & Now.Subtract(dt).Seconds)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Call BeginTrans()

        Dim sql_Delete As String = "Delete from DPTITMF1" _
            & " where MARKET_CODE = '" & Absx1.txtFor("MARKET_CODE").Text & "'" _
            & "   and OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" _
            & " and ITEM_CODE in " _
            & " (Select ITEM_CODE from " & ICTITEMS & ")"
        ASCDATA1.ExecuteSQL(sql_Delete)

        Dim sql_Delete_DPTITMB1 As String = "Delete from DPTITMB1" _
            & " where  OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" _
            & " and ITEM_CODE in " _
            & " (Select ITEM_CODE from " & ICTITEMS & ")"
        ASCDATA1.ExecuteSQL(sql_Delete_DPTITMB1)


        Write_Event_Log("DPTITMF1", Absx1.txtFor("MARKET_CODE").Text, "Forecast Updated")

        dst.Tables("DPTITMF1").Rows.Clear()
        dst.Tables("DPTITMB1").Rows.Clear()

        dst.Tables("DPTITMFX").AcceptChanges()
        For Each rowDPTITMFX As DataRow In dst.Tables("DPTITMFX").Rows
            For P As Integer = -1 To FCMAX ' 25
                Dim COLUMN_NAME As String = IIf(P = -1, "FCPD", "FC" & Format(P, "00"))
                Dim FORECAST As Integer = Val(rowDPTITMFX.Item("U" & COLUMN_NAME) & "")
                If FORECAST <> 0 Then
                    Dim rowDPTITMF1 As DataRow = dst.Tables("DPTITMF1").NewRow
                    rowDPTITMF1.Item("MARKET_CODE") = HFs("MARKET_CODE")
                    rowDPTITMF1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
                    rowDPTITMF1.Item("ITEM_CODE") = rowDPTITMFX.Item("ITEM_CODE")
                    Dim YP As String = ""
                    If P > -1 Then YP = YPF(P, 0)
                    Dim OPS_YYYYPP_FC As String = IIf(P = -1, "000000", YP)
                    rowDPTITMF1.Item("OPS_YYYYPP_FC") = OPS_YYYYPP_FC
                    rowDPTITMF1.Item("FORECAST") = FORECAST
                    dst.Tables("DPTITMF1").Rows.Add(rowDPTITMF1)
                End If
            Next
            If rowDPTITMFX.Item("BASE_IP") & "" <> "" Then
                Dim rowDPTITMB1 As DataRow = dst.Tables("DPTITMB1").NewRow
                rowDPTITMB1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
                rowDPTITMB1.Item("ITEM_CODE") = rowDPTITMFX.Item("ITEM_CODE")
                rowDPTITMB1.Item("BASE_IP") = rowDPTITMFX.Item("BASE_IP")
                dst.Tables("DPTITMB1").Rows.Add(rowDPTITMB1)
            End If
        Next

        Call Update_Record_TDA("DPTITMF1", sql_Delete)

        Call Update_Record_TDA("DPTITMB1", sql_Delete_DPTITMB1)
        Call CommitTrans("Update Complete")

    End Sub

    Sub Delete_Record()
        'Call BeginTrans()
        'Call Delete_Records("DPTITMF1")
        'Call CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where MARKET_CODE = '" & HFs("MARKET_CODE") & "'" _
        '    & "   and OPS_YYYYPP = '" & ASCMAIN1.CYP & "'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdDPTITMFX, "SSSSSSSSSSS", "Show Filter", "Show GroupBox",
                             "Show Description", "Show Class", "Show Collection",
                             "Show Item Catgy", "Show Product", "Show Basic/Promo", "Show Cost Catgy", "Show Retail", "Show Trade")

        Load_Popup_Menu(grdSATSLSI1, "BB", "Load Selected", "Load All") ' , "Load All Active Items", "Load Items with History")

        Load_Popup_Menu(grdSOTALLOX, "BB", "Select All", "De-Select All")
        Load_Popup_Menu(grdICTCOLL1, "BB", "Select All", "De-Select All")
        Load_Popup_Menu(grdDPTITMFY, "SSBBBBB", "Show Filter", "Show GroupBox",
                             "Establish Initial BaseLine %", "Rounding of Selected Column", "Distribute Base % To Selected Periods", "Distribute Base % To Selected Periods (Include FC 0's)", "Copy 3 Mth Avg to Base %")
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

        If tlb_pop.Tools.Exists("Show Description") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Description"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Groups("ITEM_DESC").Hidden
        End If
        If tlb_pop.Tools.Exists("Show Class") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Class"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Groups("ITEM_CLASS_CODE").Hidden
        End If
        If tlb_pop.Tools.Exists("Show Collection") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Collection"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Groups("COLLECTION_CODE").Hidden
        End If
        If tlb_pop.Tools.Exists("Show Item Catgy") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Item Catgy"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Groups("ITEM_CATGY_CODE").Hidden
        End If
        If tlb_pop.Tools.Exists("Show Product") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Product"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Groups("PROD_CODE").Hidden
        End If
        If tlb_pop.Tools.Exists("Show Basic/Promo") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Basic/Promo"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Groups("ITEM_BASIC_PROMO").Hidden
        End If
        If tlb_pop.Tools.Exists("Show Cost Catgy") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Cost Catgy"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Groups("COST_CATGY_CODE").Hidden
        End If
        If tlb_pop.Tools.Exists("Show Retail") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Retail"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Groups("ITEM_RETAIL_PRICE").Hidden
        End If
        If tlb_pop.Tools.Exists("Show Trade") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Trade"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Groups("ITEM_PRICE").Hidden
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

            Case "Show Description"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Groups("ITEM_DESC").Hidden = Not tlb_sbt.Checked
            Case "Show Class"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Groups("ITEM_CLASS_CODE").Hidden = Not tlb_sbt.Checked
            Case "Show Collection"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Groups("COLLECTION_CODE").Hidden = Not tlb_sbt.Checked
            Case "Show Item Catgy"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Groups("ITEM_CATGY_CODE").Hidden = Not tlb_sbt.Checked
            Case "Show Product"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Groups("PROD_CODE").Hidden = Not tlb_sbt.Checked
            Case "Show Basic/Promo"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Groups("ITEM_BASIC_PROMO").Hidden = Not tlb_sbt.Checked
            Case "Show Cost Catgy"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Groups("COST_CATGY_CODE").Hidden = Not tlb_sbt.Checked
            Case "Show Retail"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Groups("ITEM_RETAIL_PRICE").Hidden = Not tlb_sbt.Checked
            Case "Show Trade"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Groups("ITEM_PRICE").Hidden = Not tlb_sbt.Checked

            Case "Load Selected", "Load All"

                If e.Tool.Key = "Load All" Then
                    grdSATSLSI1.Selected.Rows.Clear()
                    For Each grow As UltraWinGrid.UltraGridRow In grdSATSLSI1.Rows
                        grow.Selected = True
                    Next
                Else
                    If grdSATSLSI1.Selected.Rows.Count = 0 Then
                        If grdSATSLSI1.ActiveRow IsNot Nothing Then
                            grdSATSLSI1.ActiveRow.Selected = True
                        End If
                    End If
                End If

                If grdSATSLSI1.Selected.Rows.Count = 0 Then
                    MsgBox("No Items Selected", MsgBoxStyle.OkOnly, "Cannot Add Selected Items")
                Else
                    Dim bln As Boolean = True
                    If grdDPTITMFX.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No Then
                        bln = False
                    End If
                    grdDPTITMFX.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop

                    Dim c As Integer = 0

                    For Each grow As UltraWinGrid.UltraGridRow In grdSATSLSI1.Selected.Rows
                        Dim ITEM_CODE As String = grow.Cells("ITEM_CODE").Value
                        Dim row As DataRow = dst.Tables("DPTITMFX").Rows.Find(ITEM_CODE)
                        If row Is Nothing Then
                            ' Dim grow2 As UltraWinGrid.UltraGridRow = grdDPTITMFX 
                            With grdDPTITMFX
                                If .ActiveRow IsNot Nothing AndAlso .ActiveRow.IsAddRow Then
                                    .ActiveRow = Nothing
                                End If
                                .DisplayLayout.Bands(0).AddNew()
                                With .ActiveRow
                                    .Cells("ITEM_CODE").Value = ITEM_CODE
                                    .Update()
                                    c += 1
                                End With
                            End With
                        End If
                    Next
                    If Not bln Then
                        grdDPTITMFX.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                    End If
                    grdSATSLSI1.Selected.Rows.Clear()
                    grdSATSLSI1.Refresh() ' .Refresh(UltraWinGrid.RefreshRow.ReloadData)
                    MsgBox(CStr(c) & " Items Added to Forecast", MsgBoxStyle.OkOnly, "Verification")
                End If
                ' "Load All Active Items", "Load Items with History")

            Case "Select All", "De-Select All"
                Dim TBL As String = "SOTALLOX"
                If grd.Name = "grdICTCOLL1" Then
                    grd.Tag = "X"
                    TBL = "ICTCOLL1"
                End If
                For Each row As DataRow In dst.Tables(TBL).Select("")
                    row.Item("SEL") = IIf(e.Tool.Key = "Select All", "1", "0")
                Next
                If grd.Name = "grdICTCOLL1" Then
                    grd.Tag = ""
                    Set_Collection_Inclusions()
                End If
            Case "Establish Initial BaseLine %"
                GENERATE_BASELINE_PERC()
            Case "Rounding of Selected Column"
                ROUND_IFCA_PERIOD()
            Case "Distribute Base % To Selected Periods"

                '  Dim BASETOTAL As Double = Val(dst.Tables("DPTITMFY").Compute("SUM(BASE_IP)", ""))
                Dim BASETOTAL As Decimal = Val(dst.Tables("DPTITMFY").Compute("SUM(BASE_IP)", Mid(sqlDPTITMFX_where, 5)) & "")

                If Math.Round(BASETOTAL, 2) <> 100 And Math.Round(BASETOTAL, 2) <> 0 Then
                    MsgBox("BaseLine % not distributed 100%", MsgBoxStyle.OkOnly, "Cannot Distribute Base % To Selected Periods")
                    Exit Sub
                End If

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Updating")


                PerBs.Clear()
                For Each grow As UltraWinGrid.ColumnHeader In grdDPTITMFY.Selected.Columns
                    Dim KEYSELECT As String = grow.Column.ToString
                    If Mid(KEYSELECT, 1, 4) = "IFCA" Then
                        Dim PX As String = Mid(KEYSELECT, 5, 2)
                        PerBs.Add(PX)
                    End If
                Next

                If PerBs.Count = 0 Then
                    MsgBox("No Range Of Months has been Selected", MsgBoxStyle.OkOnly, "Cannot Proceed With Distribute Base % To Selected Periods")
                    Exit Sub
                Else
                    For Each rowDPTITMFY As DataRow In dst.Tables("DPTITMFY").Select()
                        If rowDPTITMFY.Item("BASE_IP") & "" = "" Then
                        Else
                            For p As Integer = 0 To FCMAX
                                Dim Per_Used As String = Format(p, "00")
                                If PerBs.Contains(Format(p, "00")) Then
                                    If rowDPTITMFY.Item("IFCA" & Per_Used) = 0 Then
                                    Else
                                        rowDPTITMFY.Item("IFCA" & Per_Used) = rowDPTITMFY.Item("BASE_IP")
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If

                For p As Integer = 0 To FCMAX
                    Dim Per_Used As String = Format(p, "00")
                    If PerBs.Contains(Format(p, "00")) Then
                        Dim COLTOTAL As Double = Val(dst.Tables("DPTITMFY").Compute("SUM(IFCA" & Per_Used & ")", ""))
                        Dim ROUND_AMOUNT As Double = 100 - COLTOTAL
                        If ROUND_AMOUNT <> 0 Then
                            For Each rowDPTITMFY As DataRow In dst.Tables("DPTITMFY").Select()
                                If Val(rowDPTITMFY.Item("IFC" & Per_Used)) <> 0 Then
                                    rowDPTITMFY.Item("IFCA" & Per_Used) = rowDPTITMFY.Item("IFCA" & Per_Used) + ROUND_AMOUNT * Val(rowDPTITMFY.Item("IFC" & Per_Used)) / 100
                                Else
                                End If
                            Next
                        End If
                    End If
                Next
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
            Case "Copy 3 Mth Avg to Base %"
                GENERATE_BASELINE_PERC_FROM_AVG()

            Case "Distribute Base % To Selected Periods (Include FC 0's)"
                '  Dim BASETOTAL As Double = Val(dst.Tables("DPTITMFY").Compute("SUM(BASE_IP)", ""))
                Dim BASETOTAL As Decimal = Val(dst.Tables("DPTITMFY").Compute("SUM(BASE_IP)", Mid(sqlDPTITMFX_where, 5)) & "")

                If Math.Round(BASETOTAL, 2) <> 100 And Math.Round(BASETOTAL, 2) <> 0 Then
                    MsgBox("BaseLine % not distributed 100%", MsgBoxStyle.OkOnly, "Cannot Distribute Base % To Selected Periods(Inc 0's)")
                    Exit Sub
                End If

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Updating")


                PerBs.Clear()
                For Each grow As UltraWinGrid.ColumnHeader In grdDPTITMFY.Selected.Columns
                    Dim KEYSELECT As String = grow.Column.ToString
                    If Mid(KEYSELECT, 1, 4) = "IFCA" Then
                        Dim PX As String = Mid(KEYSELECT, 5, 2)
                        PerBs.Add(PX)
                    End If
                Next

                If PerBs.Count = 0 Then
                    MsgBox("No Range Of Months has been Selected", MsgBoxStyle.OkOnly, "Cannot Proceed With Distribute Base % To Selected Periods (Inc 0's)")
                    Exit Sub
                Else
                    For Each rowDPTITMFY As DataRow In dst.Tables("DPTITMFY").Select()
                        If rowDPTITMFY.Item("BASE_IP") & "" = "" Then
                        Else
                            For p As Integer = 0 To FCMAX
                                Dim Per_Used As String = Format(p, "00")
                                If PerBs.Contains(Format(p, "00")) Then
                                    rowDPTITMFY.Item("IFCA" & Per_Used) = rowDPTITMFY.Item("BASE_IP")
                                End If
                            Next
                        End If
                    Next
                End If

                For p As Integer = 0 To FCMAX
                    Dim Per_Used As String = Format(p, "00")
                    If PerBs.Contains(Format(p, "00")) Then
                        Dim COLTOTAL As Double = Val(dst.Tables("DPTITMFY").Compute("SUM(IFCA" & Per_Used & ")", ""))
                        Dim ROUND_AMOUNT As Double = 100 - COLTOTAL
                        If ROUND_AMOUNT <> 0 Then
                            For Each rowDPTITMFY As DataRow In dst.Tables("DPTITMFY").Select()
                                If Val(rowDPTITMFY.Item("IFC" & Per_Used)) <> 0 Then
                                    rowDPTITMFY.Item("IFCA" & Per_Used) = rowDPTITMFY.Item("IFCA" & Per_Used) + ROUND_AMOUNT * Val(rowDPTITMFY.Item("IFC" & Per_Used)) / 100
                                Else
                                End If
                            Next
                        End If
                    End If
                Next
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

                    'Case "Track Shipment"
                    '    If grd.ActiveRow.Cells("SHIP_REF").Text <> "" Then
                    '        Me.Cursor = Cursors.WaitCursor
                    '        Call ASCMAIN1.Progress("Now Locating DHL POD")
                    '        System.Diagnostics.Process.Start("http:   //track.dhl-usa.com/TrackByNbr.asp?ShipmentNumber=" & grd.ActiveRow.Cells("SHIP_REF").Text)
                    '        Me.Cursor = Cursors.Default
                    '        Call ASCMAIN1.Progress("")
                    '    End If

                    'Case "Job Order Inquiry"
                    '    Dim JOB_NO As String = grd.ActiveRow.Cells("JOB_NO").Text
                    '    Context_Launch("Load", JOB_NO, e.Tool.Key, "DEFJOBMI")

            Case "Clear Column"
                Dim COLUMN_NAME As String = grdDPTITMFX.Tag
                If COLUMN_NAME = "" Then Exit Sub
                If COLUMN_NAME = "CUST_STORE_NO" Then Exit Sub
                For Each row As DataRow In dst.Tables("ARTCUST2").Select("", "", DataViewRowState.CurrentRows)
                    row.Item(COLUMN_NAME) = DBNull.Value
                Next
            Case "Copy Value"
                Dim COLUMN_NAME As String = grdDPTITMFX.Tag
                If COLUMN_NAME = "" Then Exit Sub
                If grdDPTITMFX.ActiveRow Is Nothing OrElse grdDPTITMFX.ActiveRow.IsAddRow OrElse Not grdDPTITMFX.ActiveRow.IsDataRow Then Exit Sub
                Dim COPY_VALUE As String = grdDPTITMFX.ActiveRow.Cells(COLUMN_NAME).Value
                For Each row As DataRow In dst.Tables("ARTCUST2").Select("", "", DataViewRowState.CurrentRows)
                    row.Item(COLUMN_NAME) = COPY_VALUE
                Next

            Case Else
                'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select
    End Sub

    Public Overrides Sub tlb_ToolValueChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolEventArgs)
        MyBase.tlb_ToolValueChanged(sender, e)
        'Select Case e.Tool.Key
        '    Case "Best"
        '        Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
        '        = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
        '        Me.UltraChart1.ColorModel.ColorEnd = tlb_cpt.SelectedColor
        '        UltraChart1.DataBind()
        '        'grdSATCSLSS.DataBind()
        '        Application.DoEvents()
        '        grdSATCSLSS.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)

        '    Case "Worst"
        '        Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
        '        = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
        '        Me.UltraChart1.ColorModel.ColorBegin = tlb_cpt.SelectedColor
        '        UltraChart1.DataBind()
        '        'grdSATCSLSS.DataBind()
        '        Application.DoEvents()
        '        grdSATCSLSS.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)
        'End Select

    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "MARKET_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("Load", e)
                End If
            Case "BRAND_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("Load", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "MARKET_CODE"
            '    Call Click_Command("Load")
            'Case "BRAND_CODE"
            '    Call Click_Command("Load")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "MARKET_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("MARKET_CODE").Text <> "" Then
                        Call LookUp("SOTMKTC1", Absx1.txtFor("MARKET_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

            Case "BRAND_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("BRAND_CODE").Text <> "" Then
                        Call LookUp("ICTBRAN1", Absx1.txtFor("BRAND_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

        End Select
    End Sub

#End Region

#Region "grdDPTITMFX"

    Private Sub grdDPTITMFX_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdDPTITMFX.AfterCellUpdate
        With grdDPTITMFX.ActiveRow
            Select Case e.Cell.Column.Key
                Case "ITEM_CODE"
                    Dim ITEM_CODE As String = Validate_Item(.Cells("ITEM_CODE").Value)
                    If ITEM_CODE <> "" Then
                        .Cells("ITEM_DESC").Value = rowICTITEM1.Item("ITEM_DESC")

                        .Cells("ITEM_CLASS_CODE").Value = rowICTITEM1.Item("ITEM_CLASS_CODE")
                        .Cells("COLLECTION_CODE").Value = rowICTITEM1.Item("COLLECTION_CODE")
                        .Cells("ITEM_CATGY_CODE").Value = rowICTITEM1.Item("ITEM_CATGY_CODE")
                        .Cells("PROD_CODE").Value = rowICTITEM1.Item("PROD_CODE")
                        .Cells("ITEM_BASIC_PROMO").Value = rowICTITEM1.Item("ITEM_BASIC_PROMO")
                        .Cells("COST_CATGY_CODE").Value = rowICTITEM1.Item("COST_CATGY_CODE")
                        .Cells("ITEM_RETAIL_PRICE").Value = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
                        '.Cells("ITEM_RETAIL_PRICE").Value = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
                        '.Cells("ITEM_PRICE").Value = 0.6 * Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & "")
                        '.Cells("ITEM_PRICE").Value = 0.5 * Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & "")



                        If PRICE_LIST_CODE <> "" Then
                            Dim rowSOTPRIC2 As DataRow = dst.Tables("SOTPRIC2").Rows.Find(New String() {PRICE_LIST_CODE, ITEM_CODE})
                            If rowSOTPRIC2 IsNot Nothing Then
                                .Cells("ITEM_WHOLESALE_PRICE").Value = rowSOTPRIC2.Item("ITEM_PRICE")
                            Else
                                .Cells("ITEM_WHOLESALE_PRICE").Value = Val(.Cells("ITEM_RETAIL_PRICE").Value & "") * (100 - CUST_DISC_PCT) / 100
                            End If
                        End If

                        For P As Integer = -1 To FCMAX
                            ' ARE WE SUPPOSED TO BE GETTING HISTORY HERE?
                            ' MAYBE SHOULD USE THE SQL STATEMENT FROM FORM LOAD
                        Next
                    End If

                    'Case "RA_QTY"
                    '    .Cells("RA_QTY_OPEN").Value = .Cells("RA_QTY").Value

                    'Case "RA_QTY_OPEN"
                    '    .Cells("RA_QTY_CANC").Value _
                    '        = Val(.Cells("RA_QTY").Value & "") _
                    '        - Val(.Cells("RA_QTY_USED").Value & "") _
                    '        - Val(.Cells("RA_QTY_OPEN").Value & "")
                    '    If Val(.Cells("RA_QTY_CANC").Value) < 0 Then
                    '        .Cells("RA_QTY_CANC").Value = 0
                    '    End If
            End Select
        End With
    End Sub

    Private Sub grdDPTITMFX_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdDPTITMFX.AfterRowActivate


        If Trim(grdDPTITMFX.ActiveRow.Cells("ITEM_CODE").Value & "") = "" And
                (grdDPTITMFX.ActiveCell Is Nothing OrElse
                (grdDPTITMFX.ActiveCell.Column.Key <> "ITEM_CODE")) _
        Then
            grdDPTITMFX.ActiveCell = grdDPTITMFX.ActiveRow.Cells("ITEM_CODE")
            ' Exit Sub
        End If

        With grdDPTITMFX.DisplayLayout.Bands(0)
            If grdDPTITMFX.ActiveRow.IsAddRow Then
                .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
        If grdDPTITMFX.ActiveRow.IsAddRow Then
            grdSATSLSC1.Visible = False
        Else
            If Not chkFullScreen.Checked And tabDetails.SelectedTab.Key = "Item by Customer" Then
                grdSATSLSC1.Visible = True
                Dim ITEM_CODE As String = grdDPTITMFX.ActiveRow.Cells("ITEM_CODE").Text
                Setup_grdSATSLSC1(ITEM_CODE)
            Else
                grdSATSLSC1.Visible = False
            End If
        End If

        If grdDPTITMFX.ActiveRow.IsAddRow Then
            UltraPictureBox1.Visible = False
        Else
            UltraPictureBox1.Visible = True
            Dim ITEM_CODE As String = grdDPTITMFX.ActiveRow.Cells("ITEM_CODE").Text
            Dim IMAGE_NAME As String = ITEM_CODE

            Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
            Dim imgba() As Byte = Nothing
            UltraPictureBox1.Image = ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, False, , , imgba)
            ' UltraExplorerBar1.Groups("Item Image").Text = "Item Image " & ITEM_CODE
        End If

        If grdDPTITMFX.ActiveRow.IsAddRow Then
            splSOTALLO1.Visible = False
        Else
            splSOTALLO1.Visible = True
            Dim ITEM_CODE As String = grdDPTITMFX.ActiveRow.Cells("ITEM_CODE").Text

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
        End If

    End Sub

    Private Sub grdDPTITMFX_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdDPTITMFX.AfterRowsDeleted
        ' Display_Totals
    End Sub

    Private Sub grdDPTITMFX_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdDPTITMFX.AfterRowUpdate
        ' Display_Totals
        Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Value
        Dim sqlw As String = "ITEM_CODE = '" & ITEM_CODE & "'"
        If dst.Tables("DPTITMFH").Select(sqlw).Length = 0 Then
            ASCMAIN1.sql = "Select * from (" & sqlDPTITMFH & ") where " & sqlw
            Fill_Records("DPTITMFH", "", False, ASCMAIN1.sql)
        End If
        If dst.Tables("DPTITMFO").Select(sqlw).Length = 0 Then
            ASCMAIN1.sql = "Select * from (" & sqlDPTITMFO & ") where " & sqlw
            Fill_Records("DPTITMFO", "", False, ASCMAIN1.sql)
        End If
        If dst.Tables("SATSLSI1").Select(sqlw).Length = 0 Then
            Dim sql As String = Get_SelectCommand("SATSLSI1")
            ASCMAIN1.sql = "Select * from (" & sql & ") where " & sqlw
            Fill_Records("SATSLSI1", "", False, ASCMAIN1.sql)
        End If
    End Sub

    Private Sub grdDPTITMFX_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdDPTITMFX.BeforeCellUpdate

        If grdDPTITMFX.ActiveCell IsNot Nothing Then
            With grdDPTITMFX.ActiveCell
                Select Case .Column.Key
                    Case "ITEM_CODE"
                        If .Value & "" <> "" Then
                            Dim ITEM_CODE As String = Validate_Item(.Value)
                            If ITEM_CODE <> "" Then
                            Else
                                e.Cancel = True
                            End If
                        End If

                        'Case "RA_QTY2"
                        '    If Val(grdSOWRMAF2.Columns("QTY_RESERVED").Text) < 0 Then
                        '        MsgBox("Qty May Not be Negative", vbOKOnly, "Invalid Quantity")
                        '        Cancel = True
                        '    End If
                        '    If Val(grdSOWRMAF2.Columns("QTY_RESERVED").Text) = 0 Then
                        '        MsgBox("Qty May Not be Zero", 0, "Cannot Update Record")
                        '        Cancel = True
                        '        Exit Sub
                        '    End If
                End Select
            End With
        End If

    End Sub

    Private Sub grdDPTITMFX_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdDPTITMFX.BeforeExitEditMode
        If grdDPTITMFX.ActiveCell IsNot Nothing Then
            With grdDPTITMFX.ActiveCell
                Select Case .Column.Key
                    Case "ITEM_CODE"
                        If .EditorResolved IsNot Nothing AndAlso .EditorResolved.IsValid AndAlso .EditorResolved.Value IsNot Nothing Then
                            .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                        End If
                End Select
            End With
        End If
    End Sub

    Private Sub grdDPTITMFX_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdDPTITMFX.BeforeRowUpdate
        Validate_Columns("ITEM_CODE", e.Cancel)
        'If Not e.Cancel Then
        '    Validate_Columns("RA_QTY", e.Cancel) 'THIS IS DONE IN BEFORECOLUPDATE
        'End If

        If e.Cancel = True Then
            Exit Sub
        End If

        If e.Row.IsAddRow Then
            'e.Row.Cells("RA_NO").Value = RA_NO
            'Dim RA_LNO As Int64 = Val(dst.Tables("SOTRMAF2").Compute("MAX(RA_LNO)", "") & "") + 1
            'e.Row.Cells("RA_LNO").Value = RA_LNO
        End If
    End Sub

    Private Sub grdDPTITMFX_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdDPTITMFX.ClickCellButton
        With e.Cell.Row
            Select Case e.Cell.Column.Key
                Case "ITEM_CODE"
                    Dim sql_where As String = ""
                    grdClickCellButton(grdDPTITMFX, sql_where, False)
                    grdDPTITMFX.PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab)

            End Select
        End With

    End Sub

    Private Sub grdDPTITMFX_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdDPTITMFX.InitializeLayout

    End Sub
#End Region

    Sub Validate_Columns(COLUMN_NAME As String, ByRef Cancel As Boolean)
        With grdDPTITMFX.ActiveRow
            Select Case COLUMN_NAME
                Case "ITEM_CODE"
                    Dim ITEM_CODE As String = ""
                    If Trim(.Cells("ITEM_CODE").Value & "") <> "" Then
                        ITEM_CODE = Validate_Item(.Cells("ITEM_CODE").Value & "")
                    End If
                    Cancel = (ITEM_CODE = "")

                    'Case "RA_QTY"
                    '    If Trim(.Cells("ITEM_CODE").Value & "") = "" Then
                    '        Cancel = True
                    '        Exit Sub
                    '    End If
                    '    If Val(.Cells("RA_QTY").Value & "") = 0 Then
                    '        MsgBox("Qty Not Specified", vbOKOnly, "Cannot Update Record")
                    '        Cancel = True
                    '        grdSOTRMAF2.ActiveCell = grdSOTRMAF2.ActiveRow.Cells("RA_QTY")
                    '        Exit Sub
                    '    End If
                    '    If Val(.Cells("RA_QTY").Value & "") < 0 Then
                    '        MsgBox("Qty May Not be Negative", vbOKOnly, "Invalid Quantity")
                    '        Cancel = True
                    '    End If
            End Select
        End With
    End Sub

    Function Validate_Item(ITEM_CODE_z As String) As String
        Dim EMsg As String = ""
        If ITEM_CODE_z = "" Then Return ""

        Dim ITEM_CODE As String = ""
        rowICTITEM1 = LookUp("ICTITEM1", ITEM_CODE_z)

        If rowICTITEM1 Is Nothing Then
            EMsg = "Item is Not on File" & vbCrLf
        Else
            'If rowICTITEM1.Item("ITEM_STATUS") & "" <> "A" Then
            '    EMsg = "Item Status is not Active" & vbCrLf
            'End If
            If rowICTITEM1.Item("ITEM_UOM") & "" = "" Then
                EMsg = "Item does not have a valid Unit of Measure" & vbCrLf
            End If
            'If rowICTITEM1.Item("SALES_DIVISION_CODE") & "" = "" Then
            '    EMsg = "Item does not have a valid Division Code" & vbCrLf
            'End If

            Dim COLLECTION_CODE As String = rowICTITEM1.Item("COLLECTION_CODE") & ""
            Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)
            If rowICTCOLL1 Is Nothing Then
                EMsg = "Item does not have a valid Collection"
                'e.Cancel = True
            Else
                If rowICTCOLL1.Item("BRAND_CODE") & "" <> Absx1.txtFor("BRAND_CODE").Text Then
                    EMsg = "Item's Collection (" & COLLECTION_CODE & ") does not belong to Brand " & Absx1.txtFor("BRAND_CODE").Text
                    ' e.Cancel = True
                End If
            End If
        End If

        If EMsg <> "" And grdDPTITMFX.ActiveRow.IsAddRow Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Item Code Entered is Invalid because ...")
        Else
            If EMsg = "" Then
                ITEM_CODE = rowICTITEM1.Item(0)
            End If
        End If
        Return ITEM_CODE
    End Function

    Private Sub optCOLLECTION_CODE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCOLLECTION_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If DPTITMFP_PROCESS = True Then
            DPTITMFP_PROCESS = False
            Exit Sub
        End If

        If BASE_IP_PERC_CHANGE = True Then
            If MsgBox("Baseline % Modified, Do you want to disregard these changes?" & vbCr & vbCr & "Base IP has not been applied to Forecast Grid", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                DPTITMFP_PROCESS = True
                If optCOLLECTION_CODE.Value = "A" Then
                    optCOLLECTION_CODE.Value = "I"
                ElseIf optCOLLECTION_CODE.Value = "I" Then
                    optCOLLECTION_CODE.Value = "A"
                End If
                Exit Sub
                BASE_IP_PERC_CHANGE = False

                '  MsgBox("Baseline % has already been established for this High Collection", MsgBoxStyle.OkOnly, "Cannot Generate Baseline %")
            End If '
            DPTITMFP_PROCESS = False

            BASE_IP_PERC_CHANGE = False
        End If
        DPTITMFP_PROCESS = False
        'If BASE_IP_PERC_CHANGE = True Then
        'Else
        '   
        'End IfI
        Setup_grdDPTITMFX()

    End Sub

    Private Sub optITEM_CATGY_CODE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optITEM_CATGY_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_grdDPTITMFX()
        grdDPTITMFX.DisplayLayout.Bands(0).Groups("ITEM_CATGY_CODE").Hidden = (optITEM_CATGY_CODE.Value <> "A")
    End Sub

    Sub Setup_grdDPTITMFX()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now setting up Items")

        cbeCOLLECTION_CODE.Enabled = (optCOLLECTION_CODE.Value = "I")
        'grdDPTITMFX.DisplayLayout.GroupByBox.Hidden = (optCOLLECTION_CODE.Value = "A")

        With grdDPTITMFX.DisplayLayout.Bands(0)
            .Groups("COLLECTION_CODE").Hidden = (optCOLLECTION_CODE.Value <> "A")
            .SortedColumns.Clear()
            If optCOLLECTION_CODE.Value = "A" Then
                'grdDPTITMFX.DisplayLayout.Bands(0).SortedColumns.Add("COLLECTION_CODE", False, True)
            End If
            .SortedColumns.Add("ITEM_CODE", False)

            .Groups("ITEM_CATGY_CODE").Hidden = (optITEM_CATGY_CODE.Value <> "A")
        End With

        Dim COLLS As String = ""
        Dim rowICTCOLL1 As DataRow = Nothing

        Dim HC_CODE As String = ""


        ASCMAIN1.Progress("Setting up view")

        Dim DVW As DataView = DirectCast(grdDPTITMFX.DataSource, DataTable).DefaultView
        sqlDPTITMFX_where = ""
        If optCOLLECTION_CODE.Value = "A" Then
            COLLS = "All Collections"
        Else
            If chkUseHC.Checked Then
                If cbeCOLLECTION_CODE.Value & "" <> "" Then
                    rowICTCOLL1 = dst.Tables("ICTCOLL1").Rows.Find(cbeCOLLECTION_CODE.Value)
                    HC_CODE = rowICTCOLL1.Item("HC_CODE")
                    sqlDPTITMFX_where &= " and HC_CODE = '" & HC_CODE & "'"
                    COLLS = "High Collection " & HC_CODE
                Else
                    COLLS = "No High Collection Selected"
                End If
            Else
                sqlDPTITMFX_where &= " and COLLECTION_CODE = '" & cbeCOLLECTION_CODE.Value & "'"
                COLLS = cbeCOLLECTION_CODE.Value
            End If
        End If

        If optBP.Value = "*" Then
        Else
            COLLS &= ", " & optBP.Text
            sqlDPTITMFX_where &= " and ITEM_BASIC_PROMO = '" & optBP.Value & "'"
        End If

        If optSN.Value = "*" Then
        Else
            COLLS &= ", " & optSN.Text
            sqlDPTITMFX_where &= " and COST_CATGY_CODE = '" & optSN.Value & "'"
        End If

        If optITEM_CATGY_CODE.Value = "ALL" Then
            ' no changes when all collections are on display
        Else
            sqlDPTITMFX_where &= " and ITEM_CATGY_CODE = '" & optITEM_CATGY_CODE.Value & "'"
        End If

        sqlDPTITFX_filter = Mid(sqlDPTITMFX_where, 5)
        DVW.RowFilter = sqlDPTITFX_filter
        grdDPTITMFX.Text = "Sales Forecasts, by Item/Month, for Market " & Absx1.txtFor("MARKET_CODE").Text & " - " & COLLS

        ASCMAIN1.Progress("Toggle_Maintenance")
        Toggle_Maintenance()

        ASCMAIN1.Progress("Toggle_Prorate")
        Toggle_Prorate(False)

        ASCMAIN1.Progress("Toggle_Prorate")
        Toggle_Allocate(False)

        ASCMAIN1.Progress("Setup_Budgets")
        Setup_Budgets()

        DVW = DirectCast(grdDPTITMFY.DataSource, DataTable).DefaultView
        DVW.RowFilter = sqlDPTITFX_filter

        Dim BRAND_CODE As String = Absx1.txtFor("BRAND_CODE").Text
        DVW = DirectCast(grdICTCOLL1.DataSource, DataTable).DefaultView
        Dim sqlC As String = $"BRAND_CODE = '{BRAND_CODE}'"
        If HC_CODE <> "" Then
            sqlC &= $" and HC_CODE = '{HC_CODE}'"
        End If
        DVW.RowFilter = sqlC
        Sort_grdColumns(grdICTCOLL1, "COLLECTION_CODE")

        Set_Collection_Inclusions()

        If optBP.Value = "B" And optSN.Value = "S" Then
            CALC_FROM_DPTITMFX()
            tabDetails.Tabs("Item FC% to Total Sales By Month").Visible = True
        Else
            tabDetails.Tabs("Item FC% to Total Sales By Month").Visible = False
        End If
        Sort_grdColumns(grdDPTITMFY, "ITEM_CODE")
        grdDPTITMFY.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.Select
        grdDPTITMFY.DisplayLayout.Override.RowSelectors = DefaultableBoolean.Default
        grdDPTITMFY.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub cbeCOLLECTION_CODE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeCOLLECTION_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If DPTITMFP_PROCESS = True Then
            DPTITMFP_PROCESS = False
            Exit Sub
        End If

        If BASE_IP_PERC_CHANGE = True Then
            If MsgBox("Baseline % Modified, Do you want to disregard these changes?" & vbCr & vbCr & "Base IP has not been applied to Forecast Grid", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                DPTITMFP_PROCESS = True
                cbeCOLLECTION_CODE.Text = PREV_COLL_CMB
                Exit Sub
            End If
            BASE_IP_PERC_CHANGE = False
        End If
        DPTITMFP_PROCESS = False
        Setup_grdDPTITMFX()
        PREV_COLL_CMB = cbeCOLLECTION_CODE.Text

    End Sub

    Sub Import_from_Excel()
        Dim openFileDialog1 As New OpenFileDialog
        'openFileDialog1.InitialDirectory = "C:\ABS\icons\iconexperience\48x48\plain\"
        openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
        openFileDialog1.Filter = "xls files (*.xls)|*.xls"
        'openFileDialog1.FilterIndex = 1
        openFileDialog1.RestoreDirectory = True

        If openFileDialog1.ShowDialog() = DialogResult.OK Then

            Dim FILENAME As String = openFileDialog1.FileName
            Try
                Dim strConnection As String = "Provider=Microsoft.Jet.OleDb.4.0;" &
                "data source=" & FILENAME & ";" &
                "Extended Properties=Excel 8.0;"
                Dim objConnection As New System.Data.OleDb.OleDbConnection(strConnection)
                objConnection.Open()
                Dim dbSchema As DataTable = objConnection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, Nothing)
                If dbSchema.Rows.Count = 0 Then
                    MsgBox("No Sheets Found")
                    Exit Sub
                End If
                Dim strSQL As String = "SELECT * FROM [" & dbSchema.Rows(0).Item("TABLE_NAME") & "]"
                Dim objCommand As New System.Data.OleDb.OleDbCommand(strSQL, objConnection)
                Dim objAdapter As New System.Data.OleDb.OleDbDataAdapter(strSQL, objConnection)
                Dim dt As New DataTable
                objAdapter.Fill(dt)
                objConnection.Close()

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading Forecast Data from XLS")

                Dim COLs As Int32 = dt.Columns.Count
                Dim PRDmax As Int32 = COLs - 3

                If COLs < 2 Then
                    MsgBox("There appear to be no Forecasts to Import", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                Else

                End If

                EnforceConstraints(False)

                dst.Tables("DPTITMFX").Rows.Clear()

                For Each row As DataRow In dt.Rows
                    Dim ITEM_CODE As String = Trim(row.Item(0) & "")
                    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                    If rowICTITEM1 Is Nothing Then
                        ' LOG ERROR
                    Else
                        Try
                            Dim rowDPTITMFX As DataRow = Add_ITEM_CODE_to_DPTITMFX(ITEM_CODE)


                            For P As Integer = -1 To FCMAX ' PRDmax
                                Dim COLUMN_NAME As String = IIf(P = -1, "FCPD", "FC" & Format(P, "00"))
                                If P + 2 < row.Table.Columns.Count Then
                                    Dim FC As Int32 = Val(row.Item(P + 2) & "")
                                    rowDPTITMFX.Item("U" & COLUMN_NAME) = FC
                                End If
                            Next

                            dst.Tables("DPTITMFX").Rows.Add(rowDPTITMFX)
                            row.Delete()

                        Catch ex As Exception
                            'Stop
                        End Try
                    End If
                Next






                Sort_grdColumns(grdDPTITMFX, "ITEM_CODE")

                For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("DPTITMFH"), "ITEM_CODE").Rows
                    Dim ITEM_CODE As String = row.Item(0)
                    If dst.Tables("DPTITMFX").Rows.Find(ITEM_CODE) Is Nothing Then
                        ASCDATA1.DeleteRows(dst.Tables("DPTITMFH"), "ITEM_CODE = '" & ITEM_CODE & "'")
                    End If
                Next

                For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("DPTITMFO"), "ITEM_CODE").Rows
                    Dim ITEM_CODE As String = row.Item(0)
                    If dst.Tables("DPTITMFX").Rows.Find(ITEM_CODE) Is Nothing Then
                        ASCDATA1.DeleteRows(dst.Tables("DPTITMFO"), "ITEM_CODE = '" & ITEM_CODE & "'")
                    End If
                Next

                For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SATSLSI1"), "ITEM_CODE").Rows
                    Dim ITEM_CODE As String = row.Item(0)
                    If dst.Tables("DPTITMFX").Rows.Find(ITEM_CODE) Is Nothing Then
                        ASCDATA1.DeleteRows(dst.Tables("SATSLSI1"), "ITEM_CODE = '" & ITEM_CODE & "'")
                    End If
                Next


                EnforceConstraints(True)






                If dt.Rows.Count <> 0 Then
                    Dim frmASFMSGBF As New ASFMSGBF

                    frmASFMSGBF.Show_grd(dt, Me, "Records which Failed to Load")

                End If

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Cannot Import Forecast Data")
            End Try


            optCOLLECTION_CODE.Value = "A"
            optITEM_CATGY_CODE.Value = "ALL"
            optBP.Value = "*"
            optSN.Value = "*"

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")


        End If

    End Sub

    Sub Setup_Periods()
        'ReDim YPF(FCMAX, 1)
        'ReDim YPFD(FCMAX)

        'Dim P As Integer

        'ASCMAIN1.sql = "Select * from GLTPARM2 " _
        '& " where OPS_YYYYPP between '" & ASCMAIN1.CYP & "' and '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, FCMAX) & "'"
        'P = 0
        'For Each rowGLTPARM2 As DataRow In ASCDATA1.GetDataTable.Select("", "OPS_YYYYPP")
        '    YPF(P, 0) = rowGLTPARM2.Item("OPS_YYYYPP")
        '    YPF(P, 1) = Mid(rowGLTPARM2.Item("LEGEND"), 10, 6)
        '    YPFD(P) = rowGLTPARM2.Item("PRD_END_DATE")
        '    P += 1
        'Next

        'ReDim YPP(SHMAX, 1)
        'ReDim YPPD(SHMAX)

        'ASCMAIN1.sql = "Select * from GLTPARM2 " _
        '& " where OPS_YYYYPP between '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * SHMAX) & "' and '" & ASCMAIN1.CYP & "'"
        'P = 0
        'For Each rowGLTPARM2 As DataRow In ASCDATA1.GetDataTable.Select("", "OPS_YYYYPP DESC")
        '    YPP(P, 0) = rowGLTPARM2.Item("OPS_YYYYPP")
        '    YPP(P, 1) = Mid(rowGLTPARM2.Item("LEGEND"), 10, 6)
        '    YPPD(P) = rowGLTPARM2.Item("PRD_END_DATE")
        '    P += 1
        'Next

        ASCMAIN1.Get_Period_Range(-1 * SHMAX, YPPD, YPP)
        ASCMAIN1.Get_Period_Range(FCMAX, YPFD, YPF)

    End Sub

    Sub Clear_Past_Due()
        For Each grow As UltraWinGrid.UltraGridRow In grdDPTITMFX.Rows
            grow.Cells("UFCPD").Value = 0
            grow.Update()
        Next

    End Sub

    Sub Setup_grdSATSLSC1(ByVal ITEM_CODE As String)
        '  Exit Sub
        ASCDATA1.ExecuteSQL("Truncate Table " & SATSLSCI)
        ASCDATA1.ExecuteSQL("Insert into " & SATSLSCI & " " & sqlSATSLSCI, "V", New Object() {ITEM_CODE})
        ASCDATA1.ExecuteSQL("Truncate Table " & SATSLSCT)
        ASCDATA1.ExecuteSQL("Insert into " & SATSLSCT & " " & sqlSATSLSCT, "V", New Object() {ITEM_CODE})
        ASCDATA1.ExecuteSQL("Insert into " & SATSLSCI & " (CUST_CODE) Select CUST_CODE FROM " & SATSLSCT & " Minus Select CUST_CODE from " & SATSLSCI)

        Fill_Records("SATSLSC1")
        Sort_grdColumns(grdSATSLSC1, "CUST_CODE")
    End Sub

    Private Sub chkFullScreen_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFullScreen.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_FullScreen()
    End Sub

    Sub Set_FullScreen()
        SplitContainer1.Panel2Collapsed = chkFullScreen.Checked
    End Sub

    Private Sub optUS_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optUS.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If chkHistory.Checked Then
            If optUS.Value = "R" Then
                optUS.Value = "U"
            End If
        End If
        Set_US()
    End Sub

    Private Sub chkHistory_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkHistory.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_History()
    End Sub

    Sub Set_History()
        With grdDPTITMFX.DisplayLayout.Bands(0)
            For P As Int32 = 0 To 12
                .Columns("OU" & Format(P, "00")).Hidden = Not chkHistory.Checked Or optUS.Value = "S" Or Not chkShowOriginal.Checked
                .Columns("OS" & Format(P, "00")).Hidden = Not chkHistory.Checked Or optUS.Value = "U" Or Not chkShowOriginal.Checked
                .Columns("U" & Format(P, "00")).Hidden = Not chkHistory.Checked Or optUS.Value = "S"
                .Columns("S" & Format(P, "00")).Hidden = Not chkHistory.Checked Or optUS.Value = "U"
                .Columns("RU" & Format(P, "00")).Hidden = Not chkHistory.Checked Or optUS.Value = "S"
                .Columns("RS" & Format(P, "00")).Hidden = Not chkHistory.Checked Or optUS.Value = "U"
                .Columns("VU" & Format(P, "00")).Hidden = Not chkHistory.Checked Or optUS.Value = "S"
                .Columns("VS" & Format(P, "00")).Hidden = Not chkHistory.Checked Or optUS.Value = "U"
                .Columns("PU" & Format(P, "00")).Hidden = Not chkHistory.Checked Or optUS.Value = "S"
                .Columns("PS" & Format(P, "00")).Hidden = Not chkHistory.Checked Or optUS.Value = "U"
            Next

            .Columns("DATA2").Hidden = Not chkHistory.Checked
            .Columns("DATA3").Hidden = Not chkHistory.Checked
            .Columns("DATA4").Hidden = Not chkHistory.Checked
            .Columns("DATA5").Hidden = Not chkHistory.Checked
            .Columns("DATA1").Hidden = True

            If chkHistory.Checked Then
                .LevelCount = 6
                optUS.ValueList.ValueListItems(2).Appearance.ForeColor = Color.Gray
            Else
                .LevelCount = 1
                optUS.ValueList.ValueListItems(2).Appearance.ForeColor = Color.Empty
            End If

        End With

        Toggle_Maintenance()

        If Not chkHistory.Checked Then
            optUS.Value = "U"
            'optUS.Enabled = False - WHY - I MIGHT WANT TO SEE THE PROJ IN $
        Else
            If optUS.Value = "R" Then optUS.Value = "U"
            'optUS.Enabled = True
        End If
    End Sub

    Sub Toggle_Maintenance()
        Dim allow_maintenance As Boolean = False
        With grdDPTITMFX.DisplayLayout.Override
            If chkHistory.Checked Or chkTrend.Checked Then ' Or optCOLLECTION_CODE.Value = "A" Then
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                grdDPTITMFX.DisplayLayout.Bands(0).Groups("TOTAL").Hidden = True
            Else
                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop ' FixedAddRowOnBottom
                .AllowUpdate = DefaultableBoolean.True
                allow_maintenance = True
                grdDPTITMFX.DisplayLayout.Bands(0).Groups("TOTAL").Hidden = False
            End If
        End With

        With grdDPTITMFX.DisplayLayout.Bands(0)
            For i As Integer = -1 To FCMAX
                Dim P As String = "PD"
                If i > -1 Then P = Format(i, "00")
                If allow_maintenance Then
                    .Columns("UFC" & P).CellAppearance.BackColor = Color.Yellow
                    .Columns("UFC" & P).CellAppearance.ForeColor = Color.Empty
                Else
                    .Columns("UFC" & P).CellAppearance.BackColor = Color.Empty
                    .Columns("UFC" & P).CellAppearance.ForeColor = Color.HotPink
                End If
            Next
        End With

    End Sub

    Private Sub chkTrend_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTrend.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        If chkTrend.Checked Then
            chkHistory.Checked = True
            chkHistory.Enabled = False
        Else
            chkHistory.Enabled = True
        End If

        Set_Trend()
        Toggle_Maintenance()
    End Sub

    Sub Set_Trend()
        With grdDPTITMFX.DisplayLayout.Bands(0)
            For P As Int32 = 0 To 6
                .Groups("T" & Format(P, "00")).Hidden = Not chkTrend.Checked
            Next
            .Groups("TREND").Hidden = Not chkTrend.Checked

            .Columns("DATA_TREND4").Hidden = Not chkTrend.Checked
            .Columns("DATA_TREND5").Hidden = Not chkTrend.Checked

            .Columns("DATA_TREND1").Hidden = True
            'If chkTrend.Checked Then
            '    .LevelCount = 6
            'Else
            '    .LevelCount = 4
            'End If
        End With

    End Sub

    Sub Set_US()

        Dim COLUMN_NAME_price As String = "ITEM_WHOLESALE_PRICE"
        If optUS.Value = "R" Then
            COLUMN_NAME_price = "ITEM_RETAIL_PRICE"
        End If
        dst.Tables("DPTITMFX").Columns("ITEM_PRICE").Expression = COLUMN_NAME_price

        With grdDPTITMFX.DisplayLayout.Bands(0)
            .Columns("UFC" & "TOT").Hidden = (optUS.Value = "S") Or (optUS.Value = "R")
            .Columns("SFC" & "TOT").Hidden = (optUS.Value = "U")
            For P As Int32 = -1 To FCMAX
                Dim SFX As String = "PD"
                If P > -1 Then SFX = Format(P, "00")
                .Columns("UFC" & SFX).Hidden = (optUS.Value = "S") Or (optUS.Value = "R")
                .Columns("SFC" & SFX).Hidden = (optUS.Value = "U")
                .Columns("SFC" & SFX).Width = .Columns("UFC" & SFX).Width
                If P > -1 And P < 13 Then
                    .Columns("OU" & SFX).Hidden = (Not chkHistory.Checked Or Not chkShowOriginal.Checked Or optUS.Value = "S" Or (optUS.Value = "R"))
                    .Columns("OS" & SFX).Hidden = (Not chkHistory.Checked Or Not chkShowOriginal.Checked Or optUS.Value = "U")

                    .Columns("U" & SFX).Hidden = (Not chkHistory.Checked Or optUS.Value = "S" Or (optUS.Value = "R"))
                    .Columns("S" & SFX).Hidden = (Not chkHistory.Checked Or optUS.Value = "U")
                    .Columns("RU" & SFX).Hidden = (Not chkHistory.Checked Or optUS.Value = "S" Or (optUS.Value = "R"))
                    .Columns("RS" & SFX).Hidden = (Not chkHistory.Checked Or optUS.Value = "U")
                    .Columns("VU" & SFX).Hidden = (Not chkHistory.Checked Or optUS.Value = "S" Or (optUS.Value = "R"))
                    .Columns("VS" & SFX).Hidden = (Not chkHistory.Checked Or optUS.Value = "U")
                    .Columns("PU" & SFX).Hidden = (Not chkHistory.Checked Or optUS.Value = "S" Or (optUS.Value = "R"))
                    .Columns("PS" & SFX).Hidden = (Not chkHistory.Checked Or optUS.Value = "U")

                    '.Columns("OU" & Format(P, "00")).Hidden = Not chkHistory.Checked Or Not chkShowOriginal.Checked Or optUS.Value = "S" 
                    '.Columns("OS" & Format(P, "00")).Hidden = Not chkHistory.Checked Or optUS.Value = "U" Or Not chkShowOriginal.Checked


                End If
                .Groups("FC" & SFX).Width = 60
            Next
            For P As Int32 = 0 To 6
                Dim SFX As String = Format(P, "00")
                .Columns("TU" & Format(P, "00")).Hidden = (optUS.Value = "S" Or (optUS.Value = "R"))
                .Columns("TS" & Format(P, "00")).Hidden = (optUS.Value = "U")
                .Columns("TS" & Format(P, "00")).Width = .Columns("TU" & Format(P, "00")).Width

                .Columns("TOU" & Format(P, "00")).Hidden = (optUS.Value = "S" Or (optUS.Value = "R"))
                .Columns("TOS" & Format(P, "00")).Hidden = (optUS.Value = "U")

                .Columns("TPU" & Format(P, "00")).Hidden = (optUS.Value = "S" Or (optUS.Value = "R"))
                .Columns("TPS" & Format(P, "00")).Hidden = (optUS.Value = "U")

                .Columns("TAPU" & Format(P, "00")).Hidden = (optUS.Value = "S" Or (optUS.Value = "R"))
                .Columns("TAPS" & Format(P, "00")).Hidden = (optUS.Value = "U")

                .Columns("TALU" & Format(P, "00")).Hidden = (optUS.Value = "S" Or (optUS.Value = "R"))
                .Columns("TALS" & Format(P, "00")).Hidden = (optUS.Value = "U")

                .Columns("TU" & Format(P + 12, "00")).Hidden = (optUS.Value = "S" Or (optUS.Value = "R"))
                .Columns("TS" & Format(P + 12, "00")).Hidden = (optUS.Value = "U")
                .Groups("T" & Format(P, "00")).Width = 60
            Next
        End With

        With grdDPTITMFY.DisplayLayout.Bands(0)
            .Columns("UFC" & "TOT").Hidden = True
            .Columns("SFC" & "TOT").Hidden = False
            For P As Int32 = -1 To FCMAX
                Dim SFX As String = "PD"
                If P > -1 Then SFX = Format(P, "00")
                .Columns("UFC" & SFX).Hidden = True
                .Columns("SFC" & SFX).Hidden = True
                .Columns("SFC" & SFX).Width = .Columns("UFC" & SFX).Width

            Next

        End With
        'grdDPTITMFX.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)
        'grdDPTITMFX.DisplayLayout.Bands(0).LevelCount = 4

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSATSLSC1, grdSATSLSI1}
            With grd.DisplayLayout.Bands(0)
                For P As Int32 = 0 To CSMAX
                    Dim SFX As String = Format(P, "00")
                    .Columns("U" & Format(P, "00")).Hidden = (optUS.Value = "S" Or (optUS.Value = "R"))
                    .Columns("S" & Format(P, "00")).Hidden = (optUS.Value = "U")
                    .Columns("RU" & Format(P, "00")).Hidden = (optUS.Value = "S" Or (optUS.Value = "R"))
                    .Columns("RS" & Format(P, "00")).Hidden = (optUS.Value = "U")
                    .Groups("G" & Format(P, "00")).Width = 70
                Next
            End With
        Next

        grdDPTITMFX.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
    End Sub

    Private Sub grdDPTITMFX_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdDPTITMFX.InitializeRow
        e.Row.Cells("DATA_TREND0").Value = "FC"
        e.Row.Cells("DATA_TREND1").Value = ""
        e.Row.Cells("DATA_TREND2").Value = "Actual"
        e.Row.Cells("DATA_TREND3").Value = "Act%Prj"
        e.Row.Cells("DATA_TREND4").Value = "LY"
        e.Row.Cells("DATA_TREND5").Value = "Act%LY"

        e.Row.Cells("DATA0").Value = "FC"
        e.Row.Cells("DATA1").Value = ""
        e.Row.Cells("DATA2").Value = "LY-In"
        e.Row.Cells("DATA3").Value = "LY-Thru"
        e.Row.Cells("DATA4").Value = "Var"
        e.Row.Cells("DATA5").Value = "Var%"
        e.Row.Appearance.BackColor = Color.White

        For P As Integer = 6 To 0 Step -1
            If Val(e.Row.Cells("TAPU" & Format(P, "00")).Value & "") < 0 Then
                e.Row.Cells("TAPU" & Format(P, "00")).Appearance.ForeColor = Color.Red
            ElseIf Val(e.Row.Cells("TAPU" & Format(P, "00")).Value & "") > 0 Then
                e.Row.Cells("TAPU" & Format(P, "00")).Appearance.ForeColor = Color.Green
            End If

        Next



        If e.Row Is Nothing OrElse e.Row.IsFilterRow OrElse Not e.Row.IsDataRow Then
            Exit Sub
        End If

        If optUS.Value = "U" Then


            Dim MARKET_CODE As String = Absx1.txtFor("MARKET_CODE").Text
            ITEM_CODE_F = e.Row.Cells("ITEM_CODE").Value & ""


            For P = 0 To 26
                Dim sqlw As String = "ITEM_CODE = '" & ITEM_CODE_F & "' and MARKET_CODE = '" & MARKET_CODE & "' and OPS_YYYYPP_FC = '" & YPF(P, 0) & "'"
                Dim rows() As DataRow = dst.Tables("DPTITMF2").Select(sqlw, "INIT_DATE")
                If rows.Length <> 0 Then
                    Dim TT As String = ""
                    For Each row As DataRow In rows
                        If row("STATUS") & "" <> "D" Then
                            TT &= Format(row.Item("INIT_DATE"), "MM/dd/yyyy") & vbTab &
                          Format(Val(row.Item("FORECAST") & ""), "#,##0") & vbTab &
                          row.Item("FORECAST_NOTE") & vbCrLf
                        End If
                    Next
                    If TT <> "" Then
                        With e.Row.Cells("UFC" & Format(P, "00"))
                            .Appearance = cellHasNotes
                            .ToolTipText = TT
                        End With
                    End If
                Else
                    'e.Row.Cells("").Appearance = nothing
                End If
            Next
            For P = 0 To 6
                Dim sqlw As String = "ITEM_CODE = '" & ITEM_CODE_F & "' and MARKET_CODE = '" & MARKET_CODE & "' and OPS_YYYYPP_FC =  '" & YPP(P, 0) & "'"
                Dim rows() As DataRow = dst.Tables("DPTITMF2").Select(sqlw, "INIT_DATE")
                If rows.Length <> 0 Then
                    Dim TT As String = ""
                    For Each row As DataRow In rows
                        If row("STATUS") & "" <> "D" Then
                            TT &= Format(row.Item("INIT_DATE"), "MM/dd/yyyy") & vbTab &
                          Format(Val(row.Item("FORECAST") & ""), "#,##0") & vbTab &
                          row.Item("FORECAST_NOTE") & vbCrLf
                        End If
                    Next
                    If TT <> "" Then
                        With e.Row.Cells("UFC" & Format(P, "00"))
                            .Appearance = cellHasNotes
                            .ToolTipText = TT
                        End With
                    End If
                Else
                    'e.Row.Cells("").Appearance = nothing
                End If
            Next
        End If
        ' GET SOFTER COLOR FINSIHED FIRST

        ' UFCPD, UFC00-UFC25 (PDF FC & FORWARD MONTHS) - calc P from COLUMN_NAME, and use YPF(P,0) TO OBTAIN OPS_YYYYPP_FC

        ' TPU00, TPU01, THRU TPU06 (PAST MONTHS) - calc P from COLUMN_NAME, and use YPP(P,0) TO OBTAIN OPS_YYYYPP_FC
        ' in init row - set cell appearances, set tooltips

        ' also configure double click row to bring up modal

    End Sub

    Private Sub chkShowOriginal_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowOriginal.CheckedChanged

    End Sub

    Private Sub grdSATSLSI1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSATSLSI1.InitializeRow
        e.Row.Cells("DATA0").Value = "Sell-In"
        e.Row.Cells("DATA1").Value = "Sell-Thru"

        Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Value
        Dim row As DataRow = dst.Tables("DPTITMFX").Rows.Find(ITEM_CODE)
        If row Is Nothing Then
            e.Row.Cells("ITEM_CODE").Appearance.ForeColor = Color.Silver
        Else
            e.Row.Cells("ITEM_CODE").Appearance.ForeColor = Color.Empty
        End If
    End Sub

    Private Sub grdSATSLSC1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSATSLSC1.InitializeRow
        e.Row.Cells("DATA0").Value = "Sell-In"
        e.Row.Cells("DATA1").Value = "Sell-Thru"
    End Sub

    Private Sub cmdProRate_Click(sender As System.Object, e As System.EventArgs) Handles cmdProRate.Click

        If Ps.Count = 0 Then
            MsgBox("No Range of Months has been Selected", MsgBoxStyle.OkOnly, "Cannot Proceed with Sales Proration")
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Prorating")
        ' DGJ
        Dim rowDPTITMFP_FC As DataRow = dst.Tables("DPTITMFP").Rows.Find("3")
        Dim rowDPTITMFP_PR As DataRow = dst.Tables("DPTITMFP").Rows.Find("4")
        Dim FC(FCMAX) As Decimal
        Dim PR(FCMAX) As Decimal

        Dim F(FCMAX) As Decimal
        For Each P As String In Ps
            Dim M As Integer = Val(P)
            FC(M) = Val(rowDPTITMFP_FC.Item("A" & P) & "")
            PR(M) = Val(rowDPTITMFP_PR.Item("A" & P) & "")
            If PR(M) = 0 Or FC(M) = 0 Then
            Else
                F(M) = PR(M) / FC(M)
            End If
        Next

        Dim sqlSel As String = Get_Selected_Items()

        For Each rowDPTITMFX As DataRow In dst.Tables("DPTITMFX").Select(Mid(sqlDPTITMFX_where, 5) & sqlSel)
            For Each P As String In Ps
                Dim M As Integer = Val(P)
                Dim NEWFC As Decimal = Val(rowDPTITMFX.Item("UFC" & P) & "") * F(M)
                rowDPTITMFX.Item("UFC" & P) = CInt(NEWFC + 0.5)
            Next
        Next


        Dim FC_TOTAL As Decimal = 0
        For Each P As String In Ps
            FC(Val(P)) = Val(dst.Tables("DPTITMFX").Compute("SUM(SFC" & P & ")", Mid(sqlDPTITMFX_where, 5) & sqlSel) & "")
            rowDPTITMFP_FC.Item("A" & P) = FC(Val(P))
            FC_TOTAL += FC(Val(P))
        Next
        For Each P As String In Ps
            If FC_TOTAL <> 0 Then rowDPTITMFP_FC.Item("P" & P) = 100 * FC(Val(P)) / FC_TOTAL
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub cmdSelectRange_Click(sender As System.Object, e As System.EventArgs) Handles cmdSelectRange.Click
        'grdDPTITMFM.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.RowSelect
        Toggle_Prorate(False)

        Ps.Clear()
        For Each grow As UltraWinGrid.UltraGridRow In grdDPTITMFM.Selected.Rows
            Ps.Add(grow.Cells("P").Value)
        Next

        If Ps.Count = 0 Then
            MsgBox("No Range of Months has been Selected", MsgBoxStyle.OkOnly, "Cannot Proceed with Sales Pro-Ration")
            Exit Sub
        End If

        If optCOLLECTION_CODE.Value = "I" And Not chkUseHC.Checked Then
            MsgBox("Cannot Pro-rate Individual Collection without using High Collection", MsgBoxStyle.OkOnly, "Cannot Proceed with Sales Pro-Ration")
            Exit Sub
        End If


        With grdDPTITMFZ.DisplayLayout.Bands(0)
            For p As Integer = 0 To FCMAX
                .Columns("P" & Format(p, "00")).Hidden = Not Ps.Contains(Format(p, "00"))
            Next
        End With

        With grdDPTITMFP.DisplayLayout.Bands(0)
            For p As Integer = 0 To FCMAX
                .Groups("M" & Format(p, "00")).Hidden = Not Ps.Contains(Format(p, "00"))
            Next
        End With

        dst.Tables("DPTITMFP").Rows.Clear()

        Dim FC_TOTAL As Decimal = 0
        Dim LY_TOTAL As Decimal = 0
        Dim B_TOTAL As Decimal = 0
        Dim FC(FCMAX) As Decimal
        Dim LY(FCMAX) As Decimal
        Dim B(FCMAX) As Decimal
        Dim rowDPTITMFP_FC As DataRow = dst.Tables("DPTITMFP").NewRow()
        Dim rowDPTITMFP_LY As DataRow = dst.Tables("DPTITMFP").NewRow()

        Dim sqlSel As String = Get_Selected_Items()



        Dim rowDPTITMFP_B As DataRow = dst.Tables("DPTITMFP").NewRow()
        ' DGJ
        Dim BUDSQL_where As String = sqlDPTITMFX_where
        BUDSQL_where = Replace(BUDSQL_where, "and COST_CATGY_CODE = 'S'", "")
        BUDSQL_where = Replace(BUDSQL_where, "and COST_CATGY_CODE = 'N'", "")
        For Each P As String In Ps
            FC(Val(P)) = Val(dst.Tables("DPTITMFX").Compute("SUM(SFC" & P & ")", Mid(sqlDPTITMFX_where, 5) & sqlSel) & "")
            rowDPTITMFP_FC.Item("A" & P) = FC(Val(P))
            FC_TOTAL += FC(Val(P))
            If P <= 12 Then
                LY(Val(P)) = Val(dst.Tables("DPTITMFX").Compute("SUM(S" & P & ")", Mid(sqlDPTITMFX_where, 5) & sqlSel) & "")
                rowDPTITMFP_LY.Item("A" & P) = LY(Val(P))

                B(Val(P)) = Val(dst.Tables("SATBUDWX").Compute("SUM(BUD_P" & P & ")", Mid(BUDSQL_where, 5) & sqlSel) & "")
                rowDPTITMFP_B.Item("A" & P) = B(Val(P))
            End If
            LY_TOTAL += LY(Val(P))
            B_TOTAL += B(Val(P))
        Next
        For Each P As String In Ps
            If FC_TOTAL <> 0 Then rowDPTITMFP_FC.Item("P" & P) = 100 * FC(Val(P)) / FC_TOTAL
            If LY_TOTAL <> 0 Then rowDPTITMFP_LY.Item("P" & P) = 100 * LY(Val(P)) / LY_TOTAL
            If B_TOTAL <> 0 Then rowDPTITMFP_B.Item("P" & P) = 100 * B(Val(P)) / B_TOTAL
        Next

        numSales.Value = FC_TOTAL

        rowDPTITMFP_LY.Item("DATA_TYPE") = "1"
        rowDPTITMFP_LY.Item("DATA_DESC") = "LY Sales"
        dst.Tables("DPTITMFP").Rows.Add(rowDPTITMFP_LY)
        rowDPTITMFP_B.Item("DATA_TYPE") = "2"
        rowDPTITMFP_B.Item("DATA_DESC") = "Budget"
        dst.Tables("DPTITMFP").Rows.Add(rowDPTITMFP_B)

        rowDPTITMFP_FC.Item("DATA_TYPE") = "3"
        rowDPTITMFP_FC.Item("DATA_DESC") = "FC Sales"
        dst.Tables("DPTITMFP").Rows.Add(rowDPTITMFP_FC)

        Dim rowDPTITMFP As DataRow = dst.Tables("DPTITMFP").NewRow
        rowDPTITMFP.ItemArray = rowDPTITMFP_FC.ItemArray
        rowDPTITMFP.Item("DATA_TYPE") = "4"
        rowDPTITMFP.Item("DATA_DESC") = "Prorate"
        rowDPTITMFP = dst.Tables("DPTITMFP").Rows.Add(rowDPTITMFP.ItemArray)

        Sort_grdColumns(grdDPTITMFP, "DATA_TYPE")
        Toggle_Prorate(True)

    End Sub

    Sub Toggle_Prorate(tf As Boolean)
        grdDPTITMFP.Visible = tf
        lblSales.Visible = tf
        numSales.Visible = tf
        cmdProRate.Visible = tf


        grdDPTITMFZ.Visible = tf
        cmdChangePct.Visible = tf
        optSP.Visible = tf

        If tf Then
            Setup_Change_or_Prorate()
        End If


    End Sub

    Private Sub grdDPTITMFP_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdDPTITMFP.AfterCellUpdate
        Dim COLUMN_NAME As String = e.Cell.Column.Key
        Dim P As String = Mid(COLUMN_NAME, 2)
        If DPTITMFP_PROCESS = True Then Exit Sub

        If COLUMN_NAME.StartsWith("P") Then
            DPTITMFP_PROCESS = True
            e.Cell.Row.Cells("A" & P).Value = Val(numSales.Value & "") * Val(e.Cell.Value & "") / 100
        ElseIf COLUMN_NAME.StartsWith("A") Then
            If Val(numSales.Value & "") <> 0 Then
                DPTITMFP_PROCESS = True

                e.Cell.Row.Cells("P" & P).Value = Val(e.Cell.Value & "") / Val(numSales.Value & "") * 100
            End If
        End If
        DPTITMFP_PROCESS = False
    End Sub

    Private Sub grdDPTITMFP_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdDPTITMFP.AfterRowActivate
        With grdDPTITMFP.DisplayLayout.Bands(0)
            For P As Integer = 0 To FCMAX
                If grdDPTITMFP.ActiveRow.Cells("DATA_TYPE").Value <> "4" Then
                    .Columns("P" & Format(P, "00")).CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    .Columns("P" & Format(P, "00")).CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Columns("A" & Format(P, "00")).CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            Next
        End With

    End Sub

    Private Sub grdDPTITMFP_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdDPTITMFP.InitializeRow
        e.Row.Cells("DATA_TYPE1").Value = "$"
        e.Row.Cells("DATA_TYPE2").Value = "%"

        If e.Row.Cells("DATA_TYPE").Value = "4" Then
            For P As Integer = 0 To FCMAX
                e.Row.Cells("P" & Format(P, "00")).Appearance.BackColor = Color.Yellow
                e.Row.Cells("A" & Format(P, "00")).Appearance.BackColor = Color.Yellow
            Next

            If System.Math.Round(Val(e.Row.Cells("P_TOTAL").Value & ""), 2) <> 100 Then
                e.Row.Cells("P_TOTAL").Appearance.ForeColor = Color.Red
            Else
                e.Row.Cells("P_TOTAL").Appearance.ForeColor = Color.Empty
            End If
        End If
    End Sub

    Private Sub numSales_ValueChanged(sender As System.Object, e As System.EventArgs) Handles numSales.ValueChanged
        Dim rowDPTITMFP As DataRow = dst.Tables("DPTITMFP").Rows.Find("4")
        If rowDPTITMFP IsNot Nothing Then
            Dim SALES As Decimal = Val(numSales.Value & "")
            For Each P As String In Ps
                Dim PCT As Decimal = Val(rowDPTITMFP.Item("P" & P) & "")
                rowDPTITMFP.Item("A" & P) = SALES * PCT / 100
            Next
        End If
    End Sub

    Private Sub optBP_ValueChanged(sender As Object, e As EventArgs) Handles optBP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_grdDPTITMFX()
    End Sub

    Sub Setup_Budgets()
        Dim dvw As DataView = DirectCast(grdSATBUDWX.DataSource, DataTable).DefaultView
        Dim sqlw As String = ""

        If ASCMAIN1.CLIENT = "INT" Then
            If optBP.Value = "B" Or optBP.Value = "P" Then
                sqlw &= "and ITEM_BASIC_PROMO = '" & optBP.Value & "'"
            End If
            If optCOLLECTION_CODE.Value = "I" Then
                Dim COLLECTION_CODE As String = cbeCOLLECTION_CODE.Value
                If COLLECTION_CODE <> "" Then
                    'If chkUseHC.Checked Then
                    ' int budgets come in at HC
                    Dim rowICTCOLL1 As DataRow = LookUp("ICTCOLL1", COLLECTION_CODE)
                    sqlw &= "and HC_CODE = '" & rowICTCOLL1.Item("HC_CODE") & "'"
                    'Else
                    '    sqlw &= "and COLLECTION_CODE = '" & COLLECTION_CODE & "'"
                    'End If
                End If
            End If
        End If

        dvw.RowFilter = Mid(sqlw, 5)
    End Sub
    Private Sub optSN_ValueChanged(sender As Object, e As EventArgs) Handles optSN.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_grdDPTITMFX()
    End Sub

    Private Sub tabDetails_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabDetails.SelectedTabChanged

    End Sub

    Private Sub cmdSelectRangeAllo_Click(sender As Object, e As EventArgs) Handles cmdSelectRangeAllo.Click
        Toggle_Allocate(False)

        AllocPs.Clear()
        Dim pMin As String = ""
        Dim pMax As String = ""
        For Each grow As UltraWinGrid.UltraGridRow In grdSOTALLOM.Selected.Rows
            Dim PX As String = grow.Cells("P").Value
            If pMin = "" Or PX < pMin Then pMin = PX
            If pMax = "" Or PX > pMin Then pMax = PX
            AllocPs.Add(PX)
        Next

        If AllocPs.Count = 0 Then
            MsgBox("No Range of Months has been Selected", MsgBoxStyle.OkOnly, "Cannot Proceed with Allocation Application")
            Exit Sub
        End If

        With grdSOTALLOX.DisplayLayout.Bands(0)
            For p As Integer = 0 To FCMAX
                .Groups("M" & Format(p, "00")).Hidden = Not AllocPs.Contains(Format(p, "00"))
            Next
        End With

        dst.Tables("SOTALLOX").Rows.Clear()

        Dim YP1 As String = YPF(Val(pMin), 0)
        Dim YP2 As String = YPF(Val(pMax), 0)

        Dim DTE1 As Date = CDate(LookUp("GLTPARM2", YP1).Item("PRD_END_DATE")).AddDays(1).AddMonths(-1)
        Dim DTE2 As Date = CDate(LookUp("GLTPARM2", YP2).Item("PRD_END_DATE")).AddDays(1)

        Dim sqlA As String = ""
        For Each AllocP As String In AllocPs
            Dim i As Integer = Val(AllocP)
            sqlA &= ", Sum (Decode(TO_CHAR(SOTALLO1.DATE_START,'YYYYMM'),'" & YPF(i, 0) & "',SOTALLO1.QTY_ALLO_PLAN,0)) A" & Format(i, "00")
        Next

        ASCMAIN1.sql = "Select SOTALLO1.ITEM_CODE" & sqlA & vbCrLf _
            & " from SOTALLO1,ICTITEM1,ICTCOLL1" _
            & " where SOTALLO1.DATE_START >= '" & Format(DTE1, "dd-MMM-yyyy") & "'" & vbCrLf _
            & "   and SOTALLO1.DATE_START < '" & Format(DTE2, "dd-MMM-yyyy") & "'" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = SOTALLO1.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and SOTALLO1.QTY_ALLO_PLAN <> 0" & vbCrLf _
            & "   and ICTCOLL1.BRAND_CODE = '" & Absx1.txtFor("BRAND_CODE").Text & "'" & vbCrLf _
            & " group by SOTALLO1.ITEM_CODE"

        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            Dim rowSOTALLOX As DataRow = dst.Tables("SOTALLOX").NewRow
            rowSOTALLOX.Item("ITEM_CODE") = ITEM_CODE
            rowSOTALLOX.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
            rowSOTALLOX.Item("SEL") = "1"
            Dim rowDPTITMFX As DataRow = dst.Tables("DPTITMFX").Rows.Find(ITEM_CODE)
            For Each AllocP As String In AllocPs
                Dim i As Integer = Val(AllocP)
                rowSOTALLOX.Item("A" & AllocP) = row.Item("A" & AllocP)
                If rowDPTITMFX IsNot Nothing Then
                    rowSOTALLOX.Item("P" & AllocP) = rowDPTITMFX.Item("UFC" & AllocP)
                End If
            Next
            dst.Tables("SOTALLOX").Rows.Add(rowSOTALLOX)
        Next

        Sort_grdColumns(grdSOTALLOX, "ITEM_CODE")

        Toggle_Allocate(True)
    End Sub

    Sub Toggle_Allocate(tf As Boolean)
        grdSOTALLOX.Visible = tf
        cmdApplyAllocations.Visible = tf
    End Sub

    Private Sub grdSOTALLOX_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSOTALLOX.InitializeRow
        e.Row.Cells("DATA_TYPE1").Value = "Allo"
        e.Row.Cells("DATA_TYPE2").Value = "FC"

        'If e.Row.Cells("DATA_TYPE").Value = "3" Then
        '    For P As Integer = 0 To FCMAX
        '        e.Row.Cells("P" & Format(P, "00")).Appearance.BackColor = Color.Yellow
        '    Next

        '    If System.Math.Round(Val(e.Row.Cells("P_TOTAL").Value & ""), 2) <> 100 Then
        '        e.Row.Cells("P_TOTAL").Appearance.ForeColor = Color.Red
        '    Else
        '        e.Row.Cells("P_TOTAL").Appearance.ForeColor = Color.Empty
        '    End If
        'End If
    End Sub

    Private Sub cmdApplyAllocations_Click(sender As Object, e As EventArgs) Handles cmdApplyAllocations.Click

        AllocPs.Clear()
        For Each grow As UltraWinGrid.UltraGridRow In grdSOTALLOM.Selected.Rows
            Dim PX As String = grow.Cells("P").Value
            AllocPs.Add(PX)
        Next

        For Each rowSOTALLOX As DataRow In dst.Tables("SOTALLOX").Select("SEL='1'")
            Dim ITEM_CODE As String = rowSOTALLOX.Item("ITEM_CODE")
            Dim rowDPTITEMFX As DataRow = dst.Tables("DPTITMFX").Rows.Find(ITEM_CODE)
            If rowDPTITEMFX Is Nothing Then
                rowDPTITEMFX = Add_ITEM_CODE_to_DPTITMFX(ITEM_CODE)
                dst.Tables("DPTITMFX").Rows.Add(rowDPTITEMFX)
            End If
            For Each AllocP As String In AllocPs
                rowDPTITEMFX.Item("UFC" & AllocP) = rowSOTALLOX.Item("A" & AllocP)
            Next
        Next
        Toggle_Allocate(False)
    End Sub

    Private Sub grdDPTITMFM_AfterSelectChange(sender As Object, e As UltraWinGrid.AfterSelectChangeEventArgs) Handles grdDPTITMFM.AfterSelectChange
        Toggle_Prorate(False)
    End Sub

    Private Sub grdSOTALLOM_AfterSelectChange(sender As Object, e As UltraWinGrid.AfterSelectChangeEventArgs) Handles grdSOTALLOM.AfterSelectChange
        Toggle_Allocate(False)
    End Sub

    Function Add_ITEM_CODE_to_DPTITMFX(ITEM_CODE As String) As DataRow

        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)

        Dim rowDPTITMFX As DataRow = dst.Tables("DPTITMFX").NewRow
        rowDPTITMFX.Item("ITEM_CODE") = ITEM_CODE

        rowDPTITMFX.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
        rowDPTITMFX.Item("COLLECTION_CODE") = rowICTITEM1.Item("COLLECTION_CODE")
        rowDPTITMFX.Item("ITEM_CATGY_CODE") = rowICTITEM1.Item("ITEM_CATGY_CODE")

        rowDPTITMFX.Item("ITEM_CLASS_CODE") = rowICTITEM1.Item("ITEM_CLASS_CODE")
        rowDPTITMFX.Item("PROD_CODE") = rowICTITEM1.Item("PROD_CODE")
        rowDPTITMFX.Item("ITEM_BASIC_PROMO") = rowICTITEM1.Item("ITEM_BASIC_PROMO")
        rowDPTITMFX.Item("COST_CATGY_CODE") = rowICTITEM1.Item("COST_CATGY_CODE")
        rowDPTITMFX.Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
        ' rowDPTITMFX.Item("ITEM_WHOLESALE_PRICE") = 0.5 * Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & "")

        If PRICE_LIST_CODE <> "" Then
            Dim rowSOTPRIC2 As DataRow = dst.Tables("SOTPRIC2").Rows.Find(New String() {PRICE_LIST_CODE, ITEM_CODE})
            If rowSOTPRIC2 IsNot Nothing Then
                rowDPTITMFX.Item("ITEM_WHOLESALE_PRICE") = rowSOTPRIC2.Item("ITEM_PRICE")
            Else
                rowDPTITMFX.Item("ITEM_WHOLESALE_PRICE") = Val(rowDPTITMFX.Item("ITEM_RETAIL_PRICE") & "") * (100 - CUST_DISC_PCT) / 100
            End If
        End If

        Return rowDPTITMFX
    End Function

    Private Sub chkUseHC_CheckedChanged(sender As Object, e As EventArgs) Handles chkUseHC.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_grdDPTITMFX()
    End Sub

    Private Sub optSP_ValueChanged(sender As Object, e As EventArgs) Handles optSP.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Setup_Change_or_Prorate()
    End Sub

    Sub Setup_Change_or_Prorate()
        lblSales.Visible = (optSP.Value = "$")
        numSales.Visible = (optSP.Value = "$")
        grdDPTITMFP.Visible = (optSP.Value = "$")
        grdDPTITMFZ.Visible = Not (optSP.Value = "$")

        cmdChangePct.Visible = Not (optSP.Value = "$")
        cmdProRate.Visible = (optSP.Value = "$")
    End Sub

    Private Sub cmdChangePct_Click(sender As Object, e As EventArgs) Handles cmdChangePct.Click

        If Ps.Count = 0 Then
            MsgBox("No Range of Months has been Selected", MsgBoxStyle.OkOnly, "Cannot Proceed with % Change")
            Exit Sub
        End If


        If MsgBox("OK to change the Unit Forecasts by the %'ages shown?" _
                  & vbCrLf & vbCrLf & "Note - Items marked as Excluded will NOT be changed",
                  MsgBoxStyle.YesNo, "Verificaton") = MsgBoxResult.No Then Exit Sub

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Prorating")

        Dim rowDPTITMFZ As DataRow = dst.Tables("DPTITMFZ").Rows(0)

        Dim FC(FCMAX) As Decimal
        Dim PR(FCMAX) As Decimal

        Dim F(FCMAX) As Decimal
        For Each P As String In Ps
            Dim M As Integer = Val(P)
            F(M) = Val(rowDPTITMFZ.Item("P" & P) & "")
        Next


        Dim sqlSel As String = Get_Selected_Items()

        For Each rowDPTITMFX As DataRow In dst.Tables("DPTITMFX").Select(Mid(sqlDPTITMFX_where, 5) & sqlSel)
            Dim EXCL As String = rowDPTITMFX.Item("EXCL") & ""
            If EXCL <> "1" Then
                For Each P As String In Ps
                    Dim M As Integer = Val(P)
                    Dim NEWFC As Decimal = Val(rowDPTITMFX.Item("UFC" & P) & "") * (100 + F(M)) / 100
                    Dim Q As Int64 = CInt(NEWFC + 0.5)
                    If Q < 0 Then Q = 0
                    rowDPTITMFX.Item("UFC" & P) = Q
                Next
            End If
        Next


        'Dim FC_TOTAL As Decimal = 0
        'For Each P As String In Ps
        '    FC(Val(P)) = Val(dst.Tables("DPTITMFX").Compute("SUM(SFC" & P & ")", Mid(sqlDPTITMFX_where, 5)) & "")
        '    rowDPTITMFP_FC.Item("A" & P) = FC(Val(P))
        '    FC_TOTAL += FC(Val(P))
        'Next
        'For Each P As String In Ps
        '    If FC_TOTAL <> 0 Then rowDPTITMFP_FC.Item("P" & P) = 100 * FC(Val(P)) / FC_TOTAL
        'Next

        MsgBox("% Changes have been Applied", MsgBoxStyle.OkOnly, "Verificaton")


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdDPTITMFE_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdDPTITMFE.InitializeLayout

    End Sub

    Private Sub UltraGroupBox1_Click(sender As Object, e As EventArgs) Handles UltraGroupBox1.Click

    End Sub

    Private Sub DPFITMFX_PaddingChanged(sender As Object, e As EventArgs) Handles Me.PaddingChanged

    End Sub

    Private Sub grdDPTITMFX_DoubleClickCell(sender As Object, e As DoubleClickCellEventArgs) Handles grdDPTITMFX.DoubleClickCell
        If e.Cell Is Nothing OrElse e.Cell.Row Is Nothing OrElse e.Cell.Row.IsFilterRow OrElse Not e.Cell.Row.IsDataRow Then
            Exit Sub
        End If

        Dim COLUMN_NAME As String = e.Cell.Column.Key
        Dim P As Integer = 0
        If COLUMN_NAME.Length = 5 AndAlso (COLUMN_NAME.StartsWith("UFC") Or COLUMN_NAME.StartsWith("TPU")) Then
            P = Val(Mid(COLUMN_NAME, 4))
        Else
            Exit Sub
        End If

        'OPTUS.Value


        ITEM_CODE_F = e.Cell.Row.Cells("ITEM_CODE").Value

        Dim MARKET_CODE As String = Absx1.txtFor("MARKET_CODE").Text
        Dim FORECAST As Int64 = Val(e.Cell.Value & "")

        Using F As New TAC.DPFITMF2
            F.ITEM_CODE = ITEM_CODE_F
            F.MARKET_CODE = MARKET_CODE
            If COLUMN_NAME.StartsWith("UFC") Then
                F.OPS_YYYYPP_FC = YPF(P, 0)
            Else
                F.OPS_YYYYPP_FC = YPP(P, 0)
            End If

            ' F.OPS_YYYYPP_FC = YPP(P, 0)
            F.FORECAST = FORECAST
            F.allow_new_notes = True ' False

            F.ShowDialog()

            If F.update_was_clicked Then
                Fill_Records("DPTITMF2")
                grdDPTITMFX.Rows.Refresh(RefreshRow.FireInitializeRow)
            End If

        End Using
    End Sub

    Private Sub grdSOTALLO1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTALLO1.AfterRowActivate
        Setup_SOTALLO2()
    End Sub

    Sub Setup_SOTALLO2()
        If grdSOTALLO1.ActiveRow Is Nothing OrElse Not grdSOTALLO1.ActiveRow.IsDataRow Then
            grdSOTALLO2.Visible = False
        Else
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Loading Allocations")
            grdSOTALLO2.Visible = True
            Dim dvw As DataView = DirectCast(grdSOTALLO2.DataSource, DataTable).DefaultView
            Dim ALLO_CTL_NO As String = grdSOTALLO1.ActiveRow.Cells("ALLO_CTL_NO").Value
            dvw.RowFilter = "ALLO_CTL_NO = '" & ALLO_CTL_NO & "'"
            grdSOTALLO2.Text = "Allocation " & ALLO_CTL_NO
            Sort_grdColumns(grdSOTALLO2, "CUST_CODE")
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End If
    End Sub

    Sub Set_Collection_Inclusions()
        Dim COLLECTION_CODEs As New List(Of String)
        Dim COLLECTION_CODEs_EXCL As New List(Of String)
        For Each grow As UltraWinGrid.UltraGridRow In grdICTCOLL1.Rows
            Dim COLLECTION_CODE As String = grow.Cells("COLLECTION_CODE").Value
            If grow.Cells("SEL").Value = "1" Then

                COLLECTION_CODEs.Add(COLLECTION_CODE)
            Else
                COLLECTION_CODEs_EXCL.Add(COLLECTION_CODE)
            End If
        Next

        Dim dvw As DataView = DirectCast(grdDPTITMFX.DataSource, DataTable).DefaultView

        If COLLECTION_CODEs_EXCL.Count = 0 Then
            dvw.RowFilter = sqlDPTITFX_filter
        Else
            Dim SQL = ""

            If sqlDPTITFX_filter <> "" Then
                SQL = sqlDPTITFX_filter & $" and COLLECTION_CODE in ('{Join(COLLECTION_CODEs.ToArray, "','")}')"
            Else
                SQL = $"COLLECTION_CODE in ('{Join(COLLECTION_CODEs.ToArray, "','")}')"
            End If

            dvw.RowFilter = SQL
        End If
    End Sub
    Private Sub grdICTCOLL1_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdICTCOLL1.AfterRowUpdate
        If grdICTCOLL1.Tag & "" <> "X" Then
            Set_Collection_Inclusions()
        End If
    End Sub


    Sub GENERATE_BASELINE_PERC()
        Dim SFCTOT As Decimal = Val(dst.Tables("DPTITMFY").Compute("SUM(SFCTOT)", Mid(sqlDPTITMFX_where, 5)) & "")
        '     Dim PTOTAL As Decimal = Val(dst.Tables("DPTITMFX").Compute("SUM(SFC" & Format(p, "00") & ")", Mid(sqlDPTITMFX_where, 5)) & "")


        Dim BASE_IP_TOT As Decimal = Val(dst.Tables("DPTITMFY").Compute("SUM(BASE_IP)", Mid(sqlDPTITMFX_where, 5)) & "")
        Dim SSPD As String = "SFCTOT"
        If Val(SFCTOT) = 0 Then
            MsgBox("There are No Sales to Generate Baseline %", MsgBoxStyle.OkOnly, "Cannot Generate Baseline %")
            Exit Sub
        ElseIf Val(BASE_IP_TOT) <> 0 Then
            If MsgBox("Baseline % has already been established for this High Collection, Do you want to Re-build Item FC% Grid From Sales/Forecast Grid" & vbCr & vbCr & "Rebuild Item FC%", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                Exit Sub
                '  MsgBox("Baseline % has already been established for this High Collection", MsgBoxStyle.OkOnly, "Cannot Generate Baseline %")
            End If
        End If
        For Each rowDPTITMFY As DataRow In dst.Tables("DPTITMFY").Select(Mid(sqlDPTITMFX_where, 5))
            rowDPTITMFY.Item("BASE_IP") = Val(rowDPTITMFY.Item(SSPD)) / SFCTOT * 100
        Next
        BASE_IP_PERC_CHANGE = True
        '  Create_Summary(grdDPTITMFY, "BASE_IP")



    End Sub
    Sub GENERATE_BASELINE_PERC_FROM_REFRESH()
        Dim SFCTOT As Decimal = Val(dst.Tables("DPTITMFY").Compute("SUM(SFCTOT)", Mid(sqlDPTITMFX_where, 5)) & "")
        '     Dim PTOTAL As Decimal = Val(dst.Tables("DPTITMFX").Compute("SUM(SFC" & Format(p, "00") & ")", Mid(sqlDPTITMFX_where, 5)) & "")


        Dim BASE_IP_TOT As Decimal = Val(dst.Tables("DPTITMFY").Compute("SUM(BASE_IP)", Mid(sqlDPTITMFX_where, 5)) & "")
        Dim SSPD As String = "SFCTOT"
        'If Val(SFCTOT) = 0 Then
        '    MsgBox("There are No Sales to Generate Baseline %", MsgBoxStyle.OkOnly, "Cannot Generate Baseline %")
        '    Exit Sub
        'ElseIf Val(BASE_IP_TOT) <> 0 Then
        '    If MsgBox("Baseline % has already been established for this High Collection, Do you want to Re-build Item FC% Grid From Sales/Forecast Grid" & vbCr & vbCr & "Rebuild Item FC%", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
        '        Exit Sub
        '        '  MsgBox("Baseline % has already been established for this High Collection", MsgBoxStyle.OkOnly, "Cannot Generate Baseline %")
        '    End If
        'End If
        For Each rowDPTITMFY As DataRow In dst.Tables("DPTITMFY").Select(Mid(sqlDPTITMFX_where, 5))
            rowDPTITMFY.Item("BASE_IP") = Val(rowDPTITMFY.Item(SSPD)) / SFCTOT * 100
        Next
        BASE_IP_PERC_CHANGE = True
        '  Create_Summary(grdDPTITMFY, "BASE_IP")



    End Sub
    Sub GENERATE_BASELINE_PERC_FROM_AVG()
        Dim RS12 As Decimal = Val(dst.Tables("DPTITMFY").Compute("SUM(RS12)", Mid(sqlDPTITMFX_where, 5)) & "")
        '     Dim PTOTAL As Decimal = Val(dst.Tables("DPTITMFX").Compute("SUM(SFC" & Format(p, "00") & ")", Mid(sqlDPTITMFX_where, 5)) & "")


        Dim BASE_IP_TOT As Decimal = Val(dst.Tables("DPTITMFY").Compute("SUM(BASE_IP)", Mid(sqlDPTITMFX_where, 5)) & "")
        '  Dim RS12 As String = "RS12"
        If Val(RS12) = 0 Then
            MsgBox("There are No 3 Mth Avg to Copy to Baseline %", MsgBoxStyle.OkOnly, "Cannot Copy to Baseline %")
            Exit Sub
        ElseIf Val(BASE_IP_TOT) <> 0 Then
            If MsgBox("Baseline % has already been established for this High Collection, Do you want to Override Base % from  3 Mth Avg " & vbCr & vbCr & "Copy 3 Mth Avg to Base %P", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                Exit Sub
                '  MsgBox("Baseline % has already been established for this High Collection", MsgBoxStyle.OkOnly, "Cannot Generate Baseline %")
            End If
        End If
        For Each rowDPTITMFY As DataRow In dst.Tables("DPTITMFY").Select(Mid(sqlDPTITMFX_where, 5))
            rowDPTITMFY.Item("BASE_IP") = Val(rowDPTITMFY.Item("RS12"))
        Next
        BASE_IP_PERC_CHANGE = True
        '  Create_Summary(grdDPTITMFY, "BASE_IP")



    End Sub
    Sub RECALC_BASELINE_PERC()
        Dim DGJ_where = " and ITEM_BASIC_PROMO = 'B' and COST_CATGY_CODE = 'S'"
        Dim SFCTOTBS As Decimal = Val(dst.Tables("DPTITMFX").Compute("SUM(SFCTOT)", Mid(DGJ_where, 5)) & "")
        Dim SFCTOT As Decimal = Val(dst.Tables("DPTITMFY").Compute("SUM(SFCTOT)", Mid(sqlDPTITMFX_where, 5)) & "")
        '     Dim PTOTAL As Decimal = Val(dst.Tables("DPTITMFX").Compute("SUM(SFC" & Format(p, "00") & ")", Mid(sqlDPTITMFX_where, 5)) & "")


        Dim BASE_IP_TOT As Decimal = Val(dst.Tables("DPTITMFY").Compute("SUM(BASE_IP)", Mid(sqlDPTITMFX_where, 5)) & "")
        Dim SSPD As String = "SFCTOT"
        If Val(SFCTOT) = 0 Then
            MsgBox("There are No Sales to Generate Baseline %", MsgBoxStyle.OkOnly, "Cannot Generate Baseline %")
            Exit Sub
        ElseIf Val(BASE_IP_TOT) <> 0 Then
        End If
        For Each rowDPTITMFY As DataRow In dst.Tables("DPTITMFY").Select(Mid(sqlDPTITMFX_where, 5))
            rowDPTITMFY.Item("BASE_IP") = Val(rowDPTITMFY.Item(SSPD)) / SFCTOT * 100
        Next

    End Sub
    Private Sub cmdApplyCancel_Click(sender As Object, e As EventArgs) Handles cmdApplyCancel.Click
        If MsgBox("This will revert changes made in the Item FC% Grid to load FC%" & vbCr & vbCr & "Cancel changes to Item FC%", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
            Exit Sub
        End If
        CALC_fROM_ORIG()
    End Sub
    Sub CALC_FROM_DPTITMFX()
        Dim DGJ_where = " and ITEM_BASIC_PROMO = 'B' and COST_CATGY_CODE = 'S'"
        Dim SFCTOTBS As Decimal = 0
        Dim SFCTOT As Decimal = 0
        If optCOLLECTION_CODE.Value <> "A" Then
            SFCTOTBS = Val(dst.Tables("DPTITMFX").Compute("SUM(SFCTOT)", Mid(DGJ_where, 5)) & "")
            SFCTOT = Val(dst.Tables("DPTITMFX").Compute("SUM(SFCTOT)", Mid(sqlDPTITMFX_where, 5)) & "")
        End If


        dst.Tables("DPTITMFY").Rows.Clear()
        For Each ROW As DataRow In dst.Tables("DPTITMFX").Select(Mid(sqlDPTITMFX_where, 5))
            Dim ROWY As DataRow = dst.Tables("DPTITMFY").NewRow
            For Each DCOL As DataColumn In dst.Tables("DPTITMFY").Columns
                ROWY.Item(DCOL.ColumnName) = ROW.Item(DCOL.ColumnName)

                If DCOL.ColumnName = "BASE_IP" And SFCTOT <> 0 Then
                    If optCOLLECTION_CODE.Value <> "A" And Val(ROW.Item("BASE_IP") & "") <> 0 Then
                        ROWY.Item("BASE_IP") = Val(ROW.Item("BASE_IP")) * SFCTOTBS / SFCTOT
                    Else
                        ROWY.Item("BASE_IP") = ROW.Item("BASE_IP")
                    End If

                End If





            Next
            dst.Tables("DPTITMFY").Rows.Add(ROWY)
            ' dst.Tables("DPTITMFY").Rows.Add(ROW.ItemArray)
        Next

        For Each rowDPTITMFY As DataRow In dst.Tables("DPTITMFY").Select("")
            For P = 0 To FCMAX
                Dim PTOTAL As Decimal = Val(dst.Tables("DPTITMFX").Compute("SUM(SFC" & Format(P, "00") & ")", Mid(sqlDPTITMFX_where, 5)) & "")
                Dim IFC As String = "IFC" & Format(P, "00")
                Dim SSPD As String = "SFC" & Format(P, "00")
                Dim IFCA As String = "IFCA" & Format(P, "00")
                If Val(PTOTAL) = 0 Then
                    rowDPTITMFY.Item(IFC) = 0
                    rowDPTITMFY.Item(IFCA) = 0
                Else
                    rowDPTITMFY.Item(IFC) = Val(rowDPTITMFY.Item(SSPD)) / PTOTAL * 100
                    rowDPTITMFY.Item(IFCA) = rowDPTITMFY.Item(IFC)
                End If
            Next
            Dim FC_TOTAL As Decimal = 0
            Dim RSTOT As Decimal = 0
            Dim RS3MTOT As Decimal = 0
            Dim WORKTOT As Decimal = 0
            For Z = 9 To 12
                If Z = 12 And RS3MTOT <> 0 Then
                    rowDPTITMFY.Item("RS" & Format(Z, "00")) = WORKTOT / RS3MTOT * 100
                Else
                    RSTOT = Val(dst.Tables("DPTITMFX").Compute("SUM(RS" & Format(Z, "00") & ")", Mid(sqlDPTITMFX_where, 5)) & "")
                    If RSTOT <> 0 Then
                        RS3MTOT = RS3MTOT + RSTOT
                        WORKTOT = WORKTOT + Val(rowDPTITMFY.Item("RS" & Format(Z, "00")) & "")
                        rowDPTITMFY.Item("RS" & Format(Z, "00")) = Val(rowDPTITMFY.Item("RS" & Format(Z, "00")) & "") / RSTOT * 100
                    End If

                End If
                '        FC_TOTAL += fc(Val(p))
            Next

        Next

        '    RECALC_BASELINE_PERC()



    End Sub
    Sub CALC_fROM_ORIG()
        ' NEED FASTER AND CLICK OFF SET FOREGROUND COLOR FROM CELL CHANGE EVENT 

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating")

        For Each rowDPTITMFY As DataRow In dst.Tables("DPTITMFY").Select()
            For P = 0 To FCMAX
                Dim IFC As String = "IFC" & Format(P, "00")
                Dim IFCA As String = "IFCA" & Format(P, "00")
                rowDPTITMFY.Item(IFCA) = rowDPTITMFY.Item(IFC)
            Next
        Next
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub cmdApply1_Click(sender As Object, e As EventArgs) Handles cmdApply1.Click
        If MsgBox("This will Re-build Item FC% Grid From Sales/Forecast Grid" & vbCr & vbCr & "Rebuild Item FC%", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
            Exit Sub
        End If
        CALC_FROM_DPTITMFX()
        GENERATE_BASELINE_PERC_FROM_REFRESH()

    End Sub

    Private Sub grdDPTITMFY_MouseUp(sender As Object, e As MouseEventArgs) Handles grdDPTITMFY.MouseUp
        ' Exit Sub
        Dim G As UltraGrid = TryCast(sender, UltraGrid)
        Dim CE As UIElement = G.DisplayLayout.UIElement
        Dim MME As UIElement = If(CE IsNot Nothing, CE.ElementFromPoint(e.Location), Nothing)
        Dim COL As UltraGridColumn = Nothing
        While MME IsNot Nothing
            Dim HEADER_E As HeaderUIElement = TryCast(MME, HeaderUIElement)
            If HEADER_E IsNot Nothing AndAlso TypeOf HEADER_E.Header Is Infragistics.Win.UltraWinGrid.ColumnHeader Then
                COL = TryCast(HEADER_E.GetContext(GetType(UltraGridColumn)), UltraGridColumn)
                Exit While
            End If
            MME = MME.Parent

        End While
        If grdDPTITMFY.Selected.Columns.Count > 0 And COL IsNot Nothing Then
            If COL.Key = "ITEM_DESC" Or COL.Key = "ITEM_CODE" Or COL.Key = "BASE_IP" Or COL.Key = "SFCTOT" Or COL.Key = "RS09" Or COL.Key = "RS10" Or COL.Key = "RS11" Or COL.Key = "RS12" Then
                ' grdDPTITMFY.DisplayLayout.Bands(0).Columns("ITEM_CODE").
                For C As Integer = 0 To grdDPTITMFY.Selected.Columns.Count - 1
                    If grdDPTITMFY.Selected.Columns.Item(C).Column.Key = COL.Key Then
                        grdDPTITMFY.Selected.Columns.Item(C).Selected = False
                        Exit For
                    End If
                Next
            End If

        End If

    End Sub

    Private Sub cmdApply2_Click(sender As Object, e As EventArgs) Handles cmdApply2.Click
        If MsgBox("Are you sure you want to Update the Sales/Forecast Grid based on FC% Grid" & vbCr & "Update Sales/Forecast Grid", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
            Exit Sub
        End If
        For P = 0 To FCMAX
            Dim COLTOTAL As Double = Val(dst.Tables("DPTITMFY").Compute("SUM(IFCA" & Format(P, "00") & "" & ")", ""))
            If Math.Round(COLTOTAL, 2) <> 100 And Math.Round(COLTOTAL, 2) <> 0 Then
                MsgBox("All Periods are not distributed 100%", MsgBoxStyle.OkOnly, "Rounding Issues")
                Exit Sub
            End If
        Next

        ''Dim BASETOTAL As Double = Val(dst.Tables("DPTITMFY").Compute("SUM(BASE_IP)", ""))
        ''If Math.Round(BASETOTAL, 2) <> 100 And Math.Round(BASETOTAL, 2) <> 0 Then
        ''    MsgBox("BaseLine % not distributed 100%", MsgBoxStyle.OkOnly, "Incorrect Baseline Totals")
        ''    Exit Sub
        ''End If

        Dim DGJ_where = " and ITEM_BASIC_PROMO = 'B' and COST_CATGY_CODE = 'S'"
        Dim SFCTOTBS As Decimal = 0
        Dim SFCTOT As Decimal = 0
        If optCOLLECTION_CODE.Value <> "A" Then
            SFCTOTBS = Val(dst.Tables("DPTITMFX").Compute("SUM(SFCTOT)", Mid(DGJ_where, 5)) & "")
            SFCTOT = Val(dst.Tables("DPTITMFX").Compute("SUM(SFCTOT)", Mid(sqlDPTITMFX_where, 5)) & "")
        End If


        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating")

        Dim PTOTAL(FCMAX) As Decimal

        For P = 0 To FCMAX
            PTOTAL(P) = Val(dst.Tables("DPTITMFX").Compute("SUM(SFC" & Format(P, "00") & ")", Mid(sqlDPTITMFX_where, 5)) & "")
        Next


        For Each rowDPTITMFX As DataRow In dst.Tables("DPTITMFX").Select(Mid(sqlDPTITMFX_where, 5))
            Dim ITEM_CODE As String = rowDPTITMFX.Item("ITEM_CODE") & ""
            Dim FC(FCMAX) As Decimal
            Dim PR(FCMAX) As Decimal
            Dim F(FCMAX) As Decimal
            Dim SL(FCMAX) As Decimal
            Dim rowDPTITMFY As DataRow = dst.Tables("DPTITMFY").Rows.Find(New String() {ITEM_CODE})
            If rowDPTITMFY Is Nothing Then
            Else
                If optCOLLECTION_CODE.Value <> "A" And Val(rowDPTITMFY.Item("BASE_IP") & "") <> 0 And SFCTOTBS <> 0 Then
                    rowDPTITMFX.Item("BASE_IP") = Val(rowDPTITMFY.Item("BASE_IP")) * SFCTOT / SFCTOTBS
                Else
                    rowDPTITMFX.Item("BASE_IP") = rowDPTITMFY.Item("BASE_IP")
                End If

                For P = 0 To FCMAX

                    Dim IFCA As String = "IFCA" & Format(P, "00")
                    Dim IFC As String = "IFC" & Format(P, "00")
                    Dim M As Integer = Val(P)

                    If Val(rowDPTITMFY.Item(IFC) & "") <> Val(rowDPTITMFY.Item(IFCA) & "") Then
                        SL(M) = PTOTAL(P) * Val(rowDPTITMFY.Item(IFCA) & "") / 100
                        PR(M) = Val(rowDPTITMFX.Item("ITEM_PRICE"))
                        If PR(M) = 0 Or SL(M) = 0 Then
                        Else
                            F(M) = SL(M) / PR(M)
                        End If
                        Dim NEWFC As Decimal = F(M)
                        rowDPTITMFX.Item("UFC" & Format(P, "00")) = CInt(NEWFC)
                        'rowDPTITMFX.Item("UFC" & Format(P, "00")) = CInt(NEWFC + 0.5)
                    End If

                Next
            End If

        Next

        Me.Cursor = Cursors.Default
        CALC_FROM_DPTITMFX()
        BASE_IP_PERC_CHANGE = False
        ASCMAIN1.Progress("")

    End Sub

    Private Sub grdDPTITMFY_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdDPTITMFY.InitializeRow
        For P As Integer = 0 To FCMAX
            If Val(e.Row.Cells("IFCA" & Format(P, "00")).Value & "") <> Val(e.Row.Cells("IFC" & Format(P, "00")).Value & "") Then
                e.Row.Cells("IFCA" & Format(P, "00")).Appearance.ForeColor = Color.Red
            ElseIf Val(e.Row.Cells("IFCA" & Format(P, "00")).Value & "") <> Val(e.Row.Cells("BASE_IP").Value & "") And Val(e.Row.Cells("BASE_IP").Value & "") <> 0 Then
                e.Row.Cells("IFCA" & Format(P, "00")).Appearance.ForeColor = Color.BlueViolet
            Else
                e.Row.Cells("IFCA" & Format(P, "00")).Appearance.ForeColor = Color.Black
            End If
            If P = 9 Or P = 10 Or P = 11 Or P = 12 Then
                e.Row.Cells("RS" & Format(P, "00")).Appearance.ForeColor = Color.DarkGreen
            End If
        Next
    End Sub

    Private Sub chk3MRS_CheckedChanged(sender As Object, e As EventArgs) Handles chk3MRS.CheckedChanged
        With grdDPTITMFY.DisplayLayout.Bands(0)
            .Columns("RS09").Hidden = Not chk3MRS.Checked
            .Columns("RS10").Hidden = Not chk3MRS.Checked
            .Columns("RS11").Hidden = Not chk3MRS.Checked
        End With
    End Sub
    Private Sub ROUND_IFCA_PERIOD()
        PerBs.Clear()
        For Each grow As UltraWinGrid.ColumnHeader In grdDPTITMFY.Selected.Columns
            Dim KEYSELECT As String = grow.Column.ToString
            If Mid(KEYSELECT, 1, 4) = "IFCA" Then
                Dim PX As String = Mid(KEYSELECT, 5, 2)
                PerBs.Add(PX)
            End If
        Next
        If PerBs.Count = 0 Then
            MsgBox("No Range of Months has been Selected", MsgBoxStyle.OkOnly, "Cannot Proceed with Distribute Base % To Selected Periods")
            Exit Sub
        Else
            For p As Integer = 0 To FCMAX
                Dim Per_Used As String = Format(p, "00")
                If PerBs.Contains(Format(p, "00")) Then
                    Dim COLTOTAL As Double = Val(dst.Tables("DPTITMFY").Compute("SUM(IFCA" & Per_Used & ")", ""))
                    Dim ROUND_AMOUNT As Double = 100 - COLTOTAL
                    If ROUND_AMOUNT <> 0 Then
                        For Each row As DataRow In dst.Tables("DPTITMFY").Select("", "IFCA" & Per_Used & " DESC")
                            row.Item("IFCA" & Per_Used) = row.Item("IFCA" & Per_Used) + ROUND_AMOUNT * Val(row.Item("IFCA" & Per_Used)) / 100
                        Next
                        COLTOTAL = Val(dst.Tables("DPTITMFY").Compute("SUM(IFCA" & Per_Used & ")", ""))
                        ROUND_AMOUNT = 100 - COLTOTAL
                        If ROUND_AMOUNT <> 0 Then
                            For Each row As DataRow In dst.Tables("DPTITMFY").Select("", "IFCA" & Per_Used & " DESC")
                                row.Item("IFCA" & Per_Used) = row.Item("IFCA" & Per_Used) + ROUND_AMOUNT * Val(row.Item("IFCA" & Per_Used)) / 100
                            Next
                        End If
                    End If
                End If
            Next
        End If
    End Sub


    Function Get_Selected_Items() As String

        ' need to do selected items with a checkbox and not with rowselection
        ' also need to ckeck a box to use selected items rather than inferentially because some rows were selected

        Dim sqlSel As String = ""

        If ASCMAIN1.DBS_COMPANY = "SLP" AndAlso grdDPTITMFX.Selected.Rows.Count > 0 Then
            For Each grow As UltraWinGrid.UltraGridRow In grdDPTITMFX.Selected.Rows
                Dim ITEM_CODE As String = grow.Cells("ITEM_CODE").Value
                sqlSel &= $" or ITEM_CODE = '{ITEM_CODE}'"
            Next
            If sqlSel <> "" Then
                sqlSel = " (" & Mid(sqlSel, 5) & ")"
            End If
            If sqlDPTITMFX_where <> "" Then sqlSel = " and " & sqlSel
        End If

        Return sqlSel
    End Function

    Private Sub chk3MRSA_CheckedChanged(sender As Object, e As EventArgs) Handles chk3MRSA.CheckedChanged
        With grdDPTITMFY.DisplayLayout.Bands(0)
            .Columns("RS12").Hidden = Not chk3MRSA.Checked
        End With
    End Sub

    Private Sub Upload_Quarterly_XLS(filePath As String)
        Dim yearForm As New Form With {
        .Text = "Select Year for Quarterly Upload",
        .FormBorderStyle = FormBorderStyle.FixedDialog,
        .StartPosition = FormStartPosition.CenterParent,
        .MinimizeBox = False,
        .MaximizeBox = False,
        .Size = New Size(250, 150)
    }

        Dim cmbYear As New ComboBox With {
        .DropDownStyle = ComboBoxStyle.DropDownList,
        .Width = 100
    }
        Dim currentYear As Integer = CInt(ASCMAIN1.CYP.Substring(0, 4))
        For y As Integer = currentYear - 2 To currentYear + 2
            cmbYear.Items.Add(y.ToString())
        Next
        cmbYear.SelectedItem = currentYear.ToString()
        cmbYear.Location = New Point((yearForm.ClientSize.Width - cmbYear.Width) \ 2, 45)

        Dim btnOK As New Button With {.Text = "OK", .DialogResult = DialogResult.OK, .Width = 80}
        btnOK.Location = New Point((yearForm.ClientSize.Width - btnOK.Width) \ 2, 80)

        yearForm.Controls.Add(cmbYear)
        yearForm.Controls.Add(btnOK)
        yearForm.AcceptButton = btnOK

        If yearForm.ShowDialog() <> DialogResult.OK Then
            MessageBox.Show("Upload cancelled. No year selected.", "Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        QuarterlyUploadYear = cmbYear.SelectedItem.ToString()

        grdDPTXLSX1.Text = "XLS Upload: " & txtComment.Text
        Try
            ssgUPLOAD_XLS.GetLock()
            ssgUPLOAD_XLS.ActiveWorkbook = SpreadsheetGear.Factory.GetWorkbook(filePath)

            Dim DT As DataTable = dst.Tables("DPTXLSX1")
            DT.Rows.Clear()

            For Each sheet As SpreadsheetGear.IWorksheet In ssgUPLOAD_XLS.ActiveWorkbook.Worksheets
                Dim cells = sheet.Cells
                Dim lastRow As Integer = sheet.UsedRange.Row + sheet.UsedRange.RowCount - 1

                For rowIndex As Integer = 3 To lastRow
                    Dim ITEM_CODE As String = cells(rowIndex, 1).Text.Trim()

                    If String.IsNullOrWhiteSpace(ITEM_CODE) _
                OrElse ITEM_CODE.Equals("Item", StringComparison.OrdinalIgnoreCase) _
                OrElse ITEM_CODE.Equals("Description", StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If

                    Dim newRow As DataRow = DT.NewRow()
                    newRow("ITEM_CODE") = ITEM_CODE

                    Dim quarterCols As Dictionary(Of String, List(Of Integer)) = New Dictionary(Of String, List(Of Integer)) From {
                    {"FC01", New List(Of Integer) From {9, 10, 11, 12}},  ' J–M
                    {"FC04", New List(Of Integer) From {13, 14, 15, 16}}, ' N–Q
                    {"FC07", New List(Of Integer) From {17, 18, 19, 20}}, ' R–U
                    {"FC10", New List(Of Integer) From {21, 22, 23, 24}}  ' V–Y
                }

                    For Each fcKey In quarterCols.Keys
                        Dim sum As Decimal = 0
                        For Each colIndex In quarterCols(fcKey)
                            Dim val As Decimal = 0
                            Decimal.TryParse(cells(rowIndex, colIndex).Text.Replace(",", ""), val)
                            sum += val
                        Next
                        newRow(fcKey) = sum
                    Next

                    DT.Rows.Add(newRow)
                Next
            Next

        Catch ex As Exception
            MessageBox.Show("Error reading quarterly XLS: " & ex.Message, "Upload Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ssgUPLOAD_XLS.ReleaseLock()
        End Try

        SplitContainer6.Visible = True
        btnUpload.Enabled = True
        Set_Forecast_Column_Captions()
    End Sub


    Private Sub Upload_Monthly_XLS(filePath As String)
        grdDPTXLSX1.Text = "XLS Upload: " & txtComment.Text
        Try
            ssgUPLOAD_XLS.GetLock()
            ssgUPLOAD_XLS.ActiveWorkbook = SpreadsheetGear.Factory.GetWorkbook(filePath)

            Dim worksheet As SpreadsheetGear.IWorksheet = ssgUPLOAD_XLS.ActiveWorkbook.Worksheets(0)
            Dim cells = worksheet.Cells

            Dim headerRow As Integer = 2
            Dim startRow As Integer = 3
            Dim startCol As Integer = 2

            Dim DT As DataTable = dst.Tables("DPTXLSX1")
            DT.Rows.Clear()

            Dim forecastColumns As New List(Of Integer)
            Dim forecastCaptions As New List(Of String)

            ' Capture up to 12 month columns
            For col = startCol To cells.ColumnCount - 1
                Dim headerText As String = cells(headerRow, col).Text.Trim()
                If String.IsNullOrWhiteSpace(headerText) Then Exit For
                forecastColumns.Add(col)
                forecastCaptions.Add(headerText)
                If forecastColumns.Count = 12 Then Exit For
            Next

            ' Update FC01–FC12 column captions
            For i = 0 To forecastCaptions.Count - 1
                Dim fcKey As String = "FC" & (i + 1).ToString("00")
                If grdDPTXLSX1.DisplayLayout.Bands(0).Columns.Exists(fcKey) Then
                    grdDPTXLSX1.DisplayLayout.Bands(0).Columns(fcKey).Header.Caption = forecastCaptions(i)
                End If
            Next

            Dim lastRow As Integer = worksheet.UsedRange.Row + worksheet.UsedRange.RowCount - 1
            For rowIndex As Integer = startRow To lastRow
                Dim ITEM_CODE As String = cells(rowIndex, 0).Text.Trim()
                Dim DESCRIPTION As String = cells(rowIndex, 1).Text.Trim()

                If String.IsNullOrWhiteSpace(ITEM_CODE) _
               OrElse ITEM_CODE.Equals("Item", StringComparison.OrdinalIgnoreCase) _
               OrElse ITEM_CODE.Equals("Totals", StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                Dim newRow As DataRow = DT.NewRow()
                newRow("ITEM_CODE") = ITEM_CODE

                For i = 0 To forecastColumns.Count - 1
                    Dim colIndex = forecastColumns(i)
                    Dim val As Decimal = 0
                    Decimal.TryParse(cells(rowIndex, colIndex).Text.Replace(",", ""), val)
                    newRow("FC" & (i + 1).ToString("00")) = val
                Next

                DT.Rows.Add(newRow)
            Next
        Catch ex As Exception
            MessageBox.Show("Error reading monthly XLS: " & ex.Message, "Upload Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ssgUPLOAD_XLS.ReleaseLock()
        End Try

        SplitContainer6.Visible = True
        btnUpload.Enabled = True
    End Sub


    Private Function Excel_ColumnName_To_Index(columnName As String) As Integer
        Dim index As Integer = 0
        For Each c As Char In columnName.ToUpper()
            index = index * 26 + (Asc(c) - Asc("A"c) + 1)
        Next
        Return index - 1
    End Function
    Private Sub Set_Forecast_Column_Captions()
        Try
            Dim FC_Year As String = If(String.IsNullOrEmpty(QuarterlyUploadYear), ASCMAIN1.CYP.Substring(2, 2), QuarterlyUploadYear.Substring(2, 2))
            Dim MONTHS As String() = {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}

            For i As Integer = 1 To 12
                Dim colKey As String = "FC" & i.ToString("00")
                If grdDPTXLSX1.DisplayLayout.Bands(0).Columns.Exists(colKey) Then
                    grdDPTXLSX1.DisplayLayout.Bands(0).Columns(colKey).Header.Caption = $"{MONTHS(i - 1)}'{FC_Year}"
                End If
            Next
        Catch ex As Exception
            Debug.Print("Error setting forecast captions: " & ex.Message)
        End Try
    End Sub

    Private Function Caption_ToOPS_YYYYPP(caption As String) As String
        Dim parts() As String = caption.Split("'"c)
        If parts.Length = 2 Then
            Dim monthName As String = parts(0).Trim()
            Dim yearSuffix As String = parts(1).Trim()
            Dim month As Integer = DateTime.ParseExact(monthName, "MMM", Globalization.CultureInfo.InvariantCulture).Month
            Return $"20{yearSuffix}{month.ToString("00")}"
        End If
        Return ""
    End Function
    Private Function Get_Quarter_Start_FC_Indexes() As List(Of Integer)
        ' FC01 = Jan, FC04 = Apr, FC07 = Jul, FC10 = Oct
        Return New List(Of Integer) From {1, 4, 7, 10}
    End Function

    Private Sub btnForecast_Click_1(sender As Object, e As EventArgs) Handles btnForecast.Click
        If UltraTextEditor4.Text.Trim() = "" Then
            MessageBox.Show("Please select a Market.", "Missing Market", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim comment As String = txtComment.Text.Trim()
        If comment = "" Then
            Dim proceedWithoutComment = MessageBox.Show("You haven't entered a comment. Are you sure you want to continue without one?",
                                                    "No Comment Entered",
                                                    MessageBoxButtons.YesNo,
                                                    MessageBoxIcon.Question)
            If proceedWithoutComment = DialogResult.No Then Exit Sub
        End If

        Dim ofd As New OpenFileDialog With {
        .Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx",
        .Title = "Select Forecast XLS File"
    }

        If ofd.ShowDialog() = DialogResult.OK Then
            Dim filePath As String = ofd.FileName

            Dim uploadType As String = If(optUpload.Value?.ToString() = "M", "Monthly", "Quarterly")
            Dim confirmText As String = $"You are about to upload {uploadType} Forecast data for Market '{UltraTextEditor4.Text}'." & vbCrLf &
                                    $"Comment: {txtComment.Text}" & vbCrLf & vbCrLf &
                                    "Do you want to continue?"

            If MessageBox.Show(confirmText, "Confirm Upload", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Select Case optUpload.Value?.ToString()
                    Case "M"
                        Upload_Monthly_XLS(filePath)
                    Case "Q"
                        Upload_Quarterly_XLS(filePath)
                End Select
            End If
        End If
    End Sub

    Private Sub btnUpload_Click(sender As Object, e As EventArgs) Handles btnUpload.Click
        Dim FC_CAPTIONS As New List(Of String)
        For Each col As UltraWinGrid.UltraGridColumn In grdDPTXLSX1.DisplayLayout.Bands(0).Columns
            If col.Key.StartsWith("FC") Then
                FC_CAPTIONS.Add(col.Header.Caption)
            End If
        Next

        Dim MONTH_SELECTOR As New Form With {
        .Text = "Select Months to Upload",
        .Size = New Size(350, 400),
        .StartPosition = FormStartPosition.CenterParent
    }

        Dim clb As New CheckedListBox With {
        .Dock = DockStyle.Top,
        .Height = 300
    }

        For Each caption In FC_CAPTIONS
            clb.Items.Add(caption, False)
        Next

        Dim btnOK As New Button With {.Text = "OK", .DialogResult = DialogResult.OK, .Dock = DockStyle.Bottom}
        Dim btnCancel As New Button With {.Text = "Cancel", .DialogResult = DialogResult.Cancel, .Dock = DockStyle.Bottom}

        MONTH_SELECTOR.Controls.Add(clb)
        MONTH_SELECTOR.Controls.Add(btnCancel)
        MONTH_SELECTOR.Controls.Add(btnOK)
        MONTH_SELECTOR.AcceptButton = btnOK
        MONTH_SELECTOR.CancelButton = btnCancel

        If MONTH_SELECTOR.ShowDialog() = DialogResult.OK Then
            Dim SELECTED_CAPTIONS As New List(Of String)
            For Each item In clb.CheckedItems
                SELECTED_CAPTIONS.Add(item.ToString())
            Next

            If SELECTED_CAPTIONS.Count = 0 Then
                MessageBox.Show("No months selected.", "Upload Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            Dim UPLOAD_TYPE As String = optUpload.Value?.ToString()
            Dim SELECTED_FC_KEYS As New List(Of String)
            Dim FC_CAPTION_TO_YYYYPP As New Dictionary(Of String, String)

            For Each col As UltraWinGrid.UltraGridColumn In grdDPTXLSX1.DisplayLayout.Bands(0).Columns
                If col.Key.StartsWith("FC") Then
                    Dim OPS_YYYYPP As String = Caption_ToOPS_YYYYPP(col.Header.Caption)
                    FC_CAPTION_TO_YYYYPP(col.Key) = OPS_YYYYPP

                    If SELECTED_CAPTIONS.Contains(col.Header.Caption) Then
                        SELECTED_FC_KEYS.Add(col.Key)
                    End If
                End If
            Next

            Dim DT As DataTable = dst.Tables("DPTXLSX1")
            Dim QUARTER_STARTS As List(Of Integer) = Get_Quarter_Start_FC_Indexes()

            For Each row As DataRow In DT.Rows
                Dim ITEM_CODE As String = row("ITEM_CODE").ToString()
                Dim MARKET_CODE As String = UltraTextEditor4.Text
                ASCMAIN1.sql = $"SELECT COUNT(*) FROM ICTITEM1 WHERE ITEM_CODE = '{ITEM_CODE}'"
                Dim ITEM_EXISTS As Boolean = Convert.ToInt32(ASCDATA1.GetDataValue(ASCMAIN1.sql)) > 0
                If Not ITEM_EXISTS Then Continue For

                If UPLOAD_TYPE = "M" Then
                    For Each FC_KEY In SELECTED_FC_KEYS
                        Dim OPS_YYYYPP_FC As String = FC_CAPTION_TO_YYYYPP(FC_KEY)
                        Dim FORECAST As Decimal = If(IsDBNull(row(FC_KEY)), 0D, Convert.ToDecimal(row(FC_KEY)))
                        Dim CYP As String = ASCMAIN1.CYP

                        ASCMAIN1.sql = $"SELECT COUNT(*) FROM DPTITMF1 " &
                                   $"WHERE ITEM_CODE = '{ITEM_CODE}' AND MARKET_CODE = '{MARKET_CODE}' " &
                                   $"AND OPS_YYYYPP = '{CYP}' AND OPS_YYYYPP_FC = '{OPS_YYYYPP_FC}'"
                        Dim EXISTS As Boolean = Convert.ToInt32(ASCDATA1.GetDataValue(ASCMAIN1.sql)) > 0

                        If EXISTS Then
                            ASCMAIN1.sql = $"UPDATE DPTITMF1 SET FORECAST = {FORECAST} WHERE ITEM_CODE = '{ITEM_CODE}' AND MARKET_CODE = '{MARKET_CODE}' AND OPS_YYYYPP = '{CYP}' AND OPS_YYYYPP_FC = '{OPS_YYYYPP_FC}'"
                        Else
                            ASCMAIN1.sql = $"INSERT INTO DPTITMF1 (OPS_YYYYPP, ITEM_CODE, MARKET_CODE, OPS_YYYYPP_FC, FORECAST) VALUES ('{CYP}', '{ITEM_CODE}', '{MARKET_CODE}', '{OPS_YYYYPP_FC}', {FORECAST})"
                        End If
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                        'Dim RECORD_NO As String = ASCMAIN1.Next_Control_No("DPTITMF2.RECORD_NO")
                        'Dim REVISION = 1
                        'Dim FORECAST_NOTE As String = txtComment.Text & ""
                        'ASCMAIN1.sql = "INSERT INTO DPTITMF2 (" &
                        '           "OPS_YYYYPP, ITEM_CODE, MARKET_CODE, OPS_YYYYPP_FC, FORECAST, " &
                        '           "INIT_OPER, INIT_DATE, FORECAST_NOTE, LAST_OPER, LAST_DATE, STATUS, RECORD_NO, REVISION) " &
                        '           "VALUES (" &
                        '           $"'{CYP}', '{ITEM_CODE}', '{MARKET_CODE}', '{OPS_YYYYPP_FC}', {FORECAST}, " &
                        '           $"'{ASCMAIN1.USER_ID}', SYSDATE, '{FORECAST_NOTE}', '{ASCMAIN1.USER_ID}', SYSDATE, NULL, '{RECORD_NO}', {REVISION})"
                        'ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                    Next

                ElseIf UPLOAD_TYPE = "Q" Then
                    For Each START_INDEX In QUARTER_STARTS
                        Dim FC_KEY As String = "FC" & START_INDEX.ToString("00")
                        If SELECTED_FC_KEYS.Contains(FC_KEY) Then
                            Dim LEAD_MONTH As String = FC_CAPTION_TO_YYYYPP(FC_KEY)
                            Dim FORECAST As Decimal = If(IsDBNull(row(FC_KEY)), 0D, Convert.ToDecimal(row(FC_KEY)))
                            Dim CYP As String = ASCMAIN1.CYP

                            ' Zero out other months in quarter (only if they belong to current period)
                            For i = 0 To 2
                                Dim Q_KEY As String = "FC" & (START_INDEX + i).ToString("00")
                                If FC_CAPTION_TO_YYYYPP.ContainsKey(Q_KEY) Then
                                    Dim OPS_YYYYPP_FC As String = FC_CAPTION_TO_YYYYPP(Q_KEY)

                                    If OPS_YYYYPP_FC <> LEAD_MONTH Then
                                        ASCMAIN1.sql = $"SELECT COUNT(*) FROM DPTITMF1 " &
                                                   $"WHERE ITEM_CODE = '{ITEM_CODE}' AND MARKET_CODE = '{MARKET_CODE}' " &
                                                   $"AND OPS_YYYYPP = '{CYP}' AND OPS_YYYYPP_FC = '{OPS_YYYYPP_FC}'"
                                        Dim EXISTS As Boolean = Convert.ToInt32(ASCDATA1.GetDataValue(ASCMAIN1.sql)) > 0

                                        If EXISTS Then
                                            ASCMAIN1.sql = $"UPDATE DPTITMF1 SET FORECAST = 0 WHERE ITEM_CODE = '{ITEM_CODE}' AND MARKET_CODE = '{MARKET_CODE}' AND OPS_YYYYPP = '{CYP}' AND OPS_YYYYPP_FC = '{OPS_YYYYPP_FC}'"
                                        Else
                                            ASCMAIN1.sql = $"INSERT INTO DPTITMF1 (OPS_YYYYPP, ITEM_CODE, MARKET_CODE, OPS_YYYYPP_FC, FORECAST) VALUES ('{CYP}', '{ITEM_CODE}', '{MARKET_CODE}', '{OPS_YYYYPP_FC}', 0)"
                                        End If
                                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                                    End If
                                    End If
                            Next


                            ASCMAIN1.sql = $"SELECT COUNT(*) FROM DPTITMF1 " &
                                       $"WHERE ITEM_CODE = '{ITEM_CODE}' AND MARKET_CODE = '{MARKET_CODE}' " &
                                       $"AND OPS_YYYYPP = '{CYP}' AND OPS_YYYYPP_FC = '{LEAD_MONTH}'"
                                Dim EXISTS_LEAD As Boolean = Convert.ToInt32(ASCDATA1.GetDataValue(ASCMAIN1.sql)) > 0

                            If EXISTS_LEAD Then
                                ASCMAIN1.sql = $"UPDATE DPTITMF1 SET FORECAST = {FORECAST} WHERE ITEM_CODE = '{ITEM_CODE}' AND MARKET_CODE = '{MARKET_CODE}' AND OPS_YYYYPP = '{CYP}' AND OPS_YYYYPP_FC = '{LEAD_MONTH}'"
                            Else
                                ASCMAIN1.sql = $"INSERT INTO DPTITMF1 (OPS_YYYYPP, ITEM_CODE, MARKET_CODE, OPS_YYYYPP_FC, FORECAST) VALUES ('{CYP}', '{ITEM_CODE}', '{MARKET_CODE}', '{LEAD_MONTH}', {FORECAST})"
                            End If
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                            Dim RECORD_NO As String = ASCMAIN1.Next_Control_No("DPTITMF2.RECORD_NO")
                            Dim REVISION = 1
                            Dim FORECAST_NOTE As String = txtComment.Text & ""
                            ASCMAIN1.sql = "INSERT INTO DPTITMF2 (" &
                                       "OPS_YYYYPP, ITEM_CODE, MARKET_CODE, OPS_YYYYPP_FC, FORECAST, " &
                                       "INIT_OPER, INIT_DATE, FORECAST_NOTE, LAST_OPER, LAST_DATE, STATUS, RECORD_NO, REVISION) " &
                                       "VALUES (" &
                                       $"'{CYP}', '{ITEM_CODE}', '{MARKET_CODE}', '{LEAD_MONTH}', {FORECAST}, " &
                                       $"'{ASCMAIN1.USER_ID}', SYSDATE, '{FORECAST_NOTE}', '{ASCMAIN1.USER_ID}', SYSDATE, NULL, '{RECORD_NO}', {REVISION})"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                        End If
                    Next
                End If
            Next

            MessageBox.Show("Upload complete!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
End Class