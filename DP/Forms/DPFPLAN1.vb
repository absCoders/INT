Imports Infragistics.Win
Imports Infragistics.UltraChart.Resources
Imports Infragistics.Win.UltraWinGrid
Imports System.ComponentModel


Public Class DPFPLAN1

    Dim DPTPLANX As String
    Dim ICTITEMS As String
    Dim DPTITMFS As String
    Dim ICTSTATX As String
    Dim ICTTRANX As String
    Dim ICTPINVX As String
    Dim sqlICTPINVX As String
    Dim sqlPOTORDRM As String

    Dim COLUMN_NAMEs As New ArrayList
    Dim COLUMN_CAPTIONs As New ArrayList
    Dim COLUMN_NAME_by_Lvl() As String
    Dim COLUMN_CAPTION_by_Lvl() As String
    Dim G_by_Lvl() As Integer

    Dim sqlICTITEMY As String
    Dim ICTITEMY As String
    Dim sqlDPTPLANS(2) As String

    Dim YPP(,) As String
    Dim YPPD() As Date
    Dim YPF(,) As String
    Dim YPFD() As Date
    Dim YMF() As String

    Dim SN_curr As String
    Dim SN_next As String

    Dim loading_items As Boolean = False
    Dim generating_plans As Boolean = True

    ' the tooltip that we will use when the cursor is over a cell of the grid
    Dim tooltip As New System.Windows.Forms.ToolTip()

    ' this allows our tooltips to have a delay before appearing
    Dim timer As New Timer()

    ' the message that will be put in the tooltip
    Dim tooltip_msg As String
    Dim tooltip_title As String
    Dim byItem As Boolean

    Dim ITEM_CODE As String

    Dim modes_Edit As Boolean = False
    Dim modes_Edit_Plans As Boolean = False

    Dim cellHasNotes As New Appearance
    Dim PO_PARM_PINV_LT As Integer
    Dim PO_PARM_PINV_PORT As String

    Dim sqlDPTMRPGO As String
    Dim DPTMRPGO As String

    Dim RYP_desc As String
    Dim RYP_end As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "DPFPLANI" Then
            InquiryMode = True
            chkEditForecasts.Visible = False
        End If

        Get_PARM("DPTPARM1")
        Get_PARM("ICTPARM1")
        Get_PARM("POTPARM1")

        PO_PARM_PINV_LT = Val(ROWs("POTPARM1").Item("PO_PARM_PINV_LT") & "")
        PO_PARM_PINV_PORT = ROWs("POTPARM1").Item("PO_PARM_PINV_PORT") & ""

        Create_Worktables(True)
        'Create_Worktables_DPTMRPGO(True)
        TAC.DPCMAIN1.Create_Worktables_DPTMRPGO(True, Me, sqlDPTMRPGO, DPTMRPGO, RYP_end, chkShowAllMonths.Checked)

        Dim LYM As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYM, -1)
        If Mid(LYM, 5, 2) >= "07" Then
            SN_curr = "F" & Mid(LYM, 1, 4)
            SN_next = "S" & Format(Val(Mid(LYM, 1, 4)) + 1, "0000")
            chkUseNextSN.Checked = Mid(LYM, 5, 2) >= "11"
        Else
            SN_curr = "S" & Mid(LYM, 1, 4)
            SN_next = "F" & Mid(LYM, 1, 4)
            chkUseNextSN.Checked = Mid(LYM, 5, 2) >= "05"
        End If
        chkUseNextSN.Text = Replace(chkUseNextSN.Text, "SYY", Mid(SN_next, 1, 1) & Mid(SN_next, 4, 2))

        chkUseNextSN.Checked = False ' this feature is disabled for AHA until we understand better what we are doing with this concept

        Dim P As Integer = 0

        ASCMAIN1.Get_Period_Range(-25, YPPD, YPP)
        ASCMAIN1.Get_Period_Range(25, YPFD, YPF)
        ReDim YMF(25)
        For i As Int32 = 0 To 25
            YMF(i) = ASCMAIN1.Get_YYYYMM(YPF(i, 0), 0)
        Next

        With dst

            ' ICTITEMS - All Items on Screen

            ASCMAIN1.sql = "Select ITEM_CODE from ICTITEM1 where ROWNUM < 1"
            ICTITEMS = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add Primary Key (ITEM_CODE)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add ONH NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add COM NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add COMMTD NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add AVA NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add MTD NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add RTN NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add MNQ NUMBER (8,0)")

            Create_TDA(dst.Tables.Add, "APTVEND1", "*", -1, False, "")

            ASCMAIN1.sql = "Select DPTVNDI1.* from DPTVNDI1 where DPTVNDI1.ITEM_CODE = :PARM1"
            Create_TDA(dst.Tables.Add, "DPTVNDI1", "*", 0, False, "V")


            ' DPTPLAN1

            ASCMAIN1.sql = "Select DPTPLAN1.* from DPTPLAN1," & ICTITEMS & " ICTITEMS where DPTPLAN1.ITEM_CODE = ICTITEMS.ITEM_CODE"
            Create_TDA(.Tables.Add, "DPTPLAN1", "**", 0, True, "", 1)


            ' DPTPLAND

            With .Tables.Add("DPTPLAND")
                .Columns.Add("ITEM_CODE")
                .Columns.Add("DATE_BALANCE", GetType(System.DateTime))
                .Columns.Add("ACTIVITY", GetType(System.Int32))
                .Columns.Add("BALANCE", GetType(System.Int32))
                .Columns.Add("SOURCE")
                .Columns.Add("WHSE_CODE")
            End With


            ' DPTDBAL1

            With .Tables.Add("DPTDBAL1")
                .Columns.Add("ITEM_CODE")
                .Columns.Add("DATE_BALANCE", GetType(System.DateTime))
                .Columns.Add("ACTIVITY", GetType(System.Int32))
                .Columns.Add("BALANCE", GetType(System.Int32))
                .PrimaryKey = New DataColumn() { .Columns("ITEM_CODE"), .Columns("DATE_BALANCE")}
            End With

            ' SATAUTHX

            ASCMAIN1.sql = "Select X.*, ARTCUST1.CUST_NAME from ARTCUST1, " & vbCrLf _
                & " (Select CUST_CODE, COUNT (*) STORE_COUNT from (" & vbCrLf _
                & " Select Distinct SATAUTH1.CUST_CODE, SATAUTH1.CUST_STORE_NO" & vbCrLf _
                & " from SATAUTH1,ICTCOLL1" & vbCrLf _
                & " where ICTCOLL1.COLLECTION_CODE = :PARM1" & vbCrLf _
                & "   and SATAUTH1.HC_CODE = ICTCOLL1.HC_CODE" & vbCrLf _
                & "   and SATAUTH1.OPS_YYYYPP_CLOSED is Null" & vbCrLf _
                & "   and SATAUTH1.OPS_YYYYPP_OPENED is Not Null)" & vbCrLf _
                & " group by CUST_CODE) X" & vbCrLf _
                & " where ARTCUST1.CUST_CODE = X.CUST_CODE"
            Create_TDA(.Tables.Add, "SATAUTHX", "**", 0, False, "V", 1)
            dst.Tables("SATAUTHX").Columns("STORE_COUNT").DataType = GetType(System.Int32)

            ' ICTITEM1

            ASCMAIN1.sql = "Select ICTITEM1.* from ICTITEM1," & ICTITEMS & " ICTITEMS where ICTITEM1.ITEM_CODE = ICTITEMS.ITEM_CODE"
            Create_TDA(.Tables.Add, "ICTITEM1", "**", 0, True, "", 1, "ITEM_LEAD_TIME_DAYS,ITEM_SAFETY_STOCK,ITEM_SAFETY_DAYS,ITEM_FC_DAY_REQ,ITEM_PO_QTY_MIN,ITEM_PO_QTY_MULT,ITEM_POS_MAX,ITEM_POS_MIN,ITEM_PLAN_QUIET_ZONE_TYPE,ITEM_PLAN_QUIET_ZONE_DAYS,ITEM_PLAN_QUIET_ZONE_DATE,ITEM_BUFFER_QTY,ITEM_BUFFER_PCT,ITEM_ABC_PARMS_LOCKED")


            ' ICTSTATH

            ASCMAIN1.sql = "Select ICTSTAT5.ITEM_CODE" _
            & ", SUM (DECODE(ICTSTAT5.OPS_YYYYPP,'" & YPP(3, 0) & "',ICTSTAT5.WHSE_QTY_ON_HAND,0)) OH3 " _
            & ", SUM (DECODE(ICTSTAT5.OPS_YYYYPP,'" & YPP(2, 0) & "',ICTSTAT5.WHSE_QTY_ON_HAND,0)) OH2 " _
            & ", SUM (DECODE(ICTSTAT5.OPS_YYYYPP,'" & YPP(1, 0) & "',ICTSTAT5.WHSE_QTY_ON_HAND,0)) OH1 " _
            & " from ICTSTAT5,ICTWHSE1," & ICTITEMS & " ICTITEMS" _
            & " where ICTSTAT5.ITEM_CODE = ICTITEMS.ITEM_CODE" _
            & "   and ICTWHSE1.WHSE_CODE = ICTSTAT5.WHSE_CODE" _
            & "   and NVL(ICTWHSE1.WHSE_MRP_EXC_IND,'0') <> '1'" _
            & " group by ICTSTAT5.ITEM_CODE"
            Create_TDA(.Tables.Add, "ICTSTATH", "**", 0, False, "", 1)


            ' DPTITMFM

            ASCMAIN1.sql = "Select DPTITMF1.ITEM_CODE" & vbCrLf _
            & ", SUM (DECODE(DPTITMF1.OPS_YYYYPP_FC,'" & "000000" & "',DPTITMF1.FORECAST,0)) FCPD" & vbCrLf
            For P = 0 To 25
                ASCMAIN1.sql &= ", SUM (DECODE(DPTITMF1.OPS_YYYYPP_FC,'" & YPF(P, 0) & "',DPTITMF1.FORECAST,0)) FC" & Format(P, "00") & vbCrLf
            Next
            ASCMAIN1.sql &= " from DPTITMF1," & ICTITEMS & " ICTITEMS" & vbCrLf _
            & " where DPTITMF1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "   and DPTITMF1.ITEM_CODE = ICTITEMS.ITEM_CODE " & vbCrLf _
            & " group by DPTITMF1.ITEM_CODE"
            Create_TDA(.Tables.Add, "DPTITMFM", "**", 0, False, "", 1)


            With cellHasNotes
                .BackColor2 = Color.White
                .BackColor = Color.LightSkyBlue
                .BackGradientStyle = GradientStyle.ForwardDiagonal
            End With


            ASCMAIN1.sql = "Select DPTITMF2.*" & vbCrLf _
            & " from DPTITMF2," & ICTITEMS & " ICTITEMS" & vbCrLf _
            & " where DPTITMF2.ITEM_CODE = ICTITEMS.ITEM_CODE " & vbCrLf _
            & "   and DPTITMF2.OPS_YYYYPP_FC >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12) & "'"
            Create_TDA(.Tables.Add, "DPTITMF2", "**", 0, False, "", 0)

            ' POTORDRM

            Dim LM As Date = YPPD(1)
            ASCMAIN1.sql = "Select POTORDR2.ITEM_CODE" & vbCrLf _
            & ", SUM (CASE WHEN POTORDR2.PO_DATE_REQUIRED <= '" & Format(LM, "dd-MMM-yyyy") & "' THEN POTORDR2.PO_QTY_OPN ELSE 0 END) POPD" & vbCrLf
            For P = 0 To 25
                Dim TM As Date = YPFD(P)
                ASCMAIN1.sql &= ", SUM (CASE WHEN POTORDR2.PO_DATE_REQUIRED " _
                & " > '" & Format(LM, "dd-MMM-yyyy") & "'" _
                & " AND POTORDR2.PO_DATE_REQUIRED <= '" & Format(TM, "dd-MMM-yyyy") & "'" _
                & " THEN POTORDR2.PO_QTY_OPN ELSE 0 END) PO" & Format(P, "00") & vbCrLf
                LM = TM
            Next
            ASCMAIN1.sql &= " from POTORDR2," & ICTITEMS & " ICTITEMS" & vbCrLf _
            & " where POTORDR2.ITEM_CODE = ICTITEMS.ITEM_CODE " & vbCrLf _
            & " group by POTORDR2.ITEM_CODE"
            'sqlPOTORDRM = Replace(Replace(Replace(ASCMAIN1.sql, "POTORDR2", ICTPINVX), "PO_DATE_REQUIRED", "ETA_DATE"), "PO_QTY_OPN", "PINV_QTY")
            sqlPOTORDRM = Replace(Replace(Replace(ASCMAIN1.sql, "POTORDR2", ICTPINVX), "PO_DATE_REQUIRED", "ETA_DATE_DC"), "PO_QTY_OPN", "PINV_QTY")
            Create_TDA(.Tables.Add, "POTORDRM", "**", 0, False, "", 1)


            ' DPTPLANM

            '  Dim LM As Date = YPPD(1)
            LM = YPPD(1)
            ASCMAIN1.sql = "Select DPTPLAN1.ITEM_CODE" & vbCrLf _
            & ", SUM (CASE WHEN DPTPLAN1.DATE_REQUIRED <= '" & Format(LM, "dd-MMM-yyyy") & "' THEN DPTPLAN1.QTY_PLANNED ELSE 0 END) PLPD" & vbCrLf
            For P = 0 To 25
                Dim TM As Date = YPFD(P)
                ASCMAIN1.sql &= ", SUM (CASE WHEN DPTPLAN1.DATE_REQUIRED " _
                & " > '" & Format(LM, "dd-MMM-yyyy") & "'" _
                & " AND DPTPLAN1.DATE_REQUIRED <= '" & Format(TM, "dd-MMM-yyyy") & "'" _
                & " THEN DPTPLAN1.QTY_PLANNED ELSE 0 END) PL" & Format(P, "00") & vbCrLf
                LM = TM
            Next
            ASCMAIN1.sql &= " from DPTPLAN1," & ICTITEMS & " ICTITEMS" & vbCrLf _
            & " where DPTPLAN1.ITEM_CODE = ICTITEMS.ITEM_CODE " & vbCrLf _
            & " group by DPTPLAN1.ITEM_CODE"
            Create_TDA(.Tables.Add, "DPTPLANM", "**", 0, False, "", 1)


            ' POTORDCM

            '  Dim LM As Date = YPPD(1)
            LM = YPPD(1)
            ASCMAIN1.sql = "Select POTORDR9.ITEM_CODE" & vbCrLf _
            & ", SUM (CASE WHEN NVL(DPTPLAN1.DATE_REQUIRED,POTORDR2.PO_DATE_REQUIRED) <= '" & Format(LM, "dd-MMM-yyyy") & "' THEN POTORDR9.PO_QTY_COM ELSE 0 END) PCPD" & vbCrLf
            For P = 0 To 25
                Dim TM As Date = YPFD(P)
                ASCMAIN1.sql &= ", SUM (CASE WHEN NVL(DPTPLAN1.DATE_REQUIRED,POTORDR2.PO_DATE_REQUIRED) " _
                & " > '" & Format(LM, "dd-MMM-yyyy") & "'" _
                & " AND NVL(DPTPLAN1.DATE_REQUIRED,POTORDR2.PO_DATE_REQUIRED) <= '" & Format(TM, "dd-MMM-yyyy") & "'" _
                & " THEN POTORDR9.PO_QTY_COM ELSE 0 END) PC" & Format(P, "00") & vbCrLf
                LM = TM
            Next
            ASCMAIN1.sql &= " from POTORDR9,POTORDR2,DPTPLAN1," & ICTITEMS & " ICTITEMS" & vbCrLf _
            & " where POTORDR9.ITEM_CODE = ICTITEMS.ITEM_CODE " & vbCrLf _
            & "   and DPTPLAN1.PLAN_NO (+) = (CASE WHEN POTORDR9.PO_ORDER_LNO = 0 THEN POTORDR9.PO_ORDER_NO ELSE 'X' END)" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_NO (+) = POTORDR9.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_LNO (+) = POTORDR9.PO_ORDER_LNO" & vbCrLf _
            & " group by POTORDR9.ITEM_CODE"
            Create_TDA(.Tables.Add, "POTORDCM", "**", 0, False, "", 1)


            ' POTORDR9

            ASCMAIN1.sql = "Select POTORDR9.*" & vbCrLf _
                & ", NVL(DPTPLAN1.ITEM_CODE,POTORDR2.ITEM_CODE) ITEM_CODE_PROD" & vbCrLf _
                & ", NVL(ICTITEM1.ITEM_DESC,POTORDR2.ITEM_DESC) ITEM_DESC_PROD" & vbCrLf _
                & ", NVL(DPTPLAN1.DATE_REQUIRED,POTORDR2.PO_DATE_REQUIRED) DATE_REQUIRED" & vbCrLf _
                & ", NVL(DPTPLAN1.DATE_COMPSDUE,POTORDR2.PO_DATE_COMPSDUE) DATE_COMPSDUE" & vbCrLf _
                & ", NVL(DPTPLAN1.AT_WHSE,POTORDR1.VEND_WHSE_CODE) AT_WHSE" & vbCrLf _
                & ", NVL(DPTPLAN1.TO_WHSE,POTORDR2.WHSE_CODE) TO_WHSE" & vbCrLf _
                & " from POTORDR9,POTORDR2,POTORDR1,DPTPLAN1," & ICTITEMS & " ICTITEMS,ICTITEM1" & vbCrLf _
                & " where POTORDR9.ITEM_CODE = ICTITEMS.ITEM_CODE " & vbCrLf _
                & "   and DPTPLAN1.PLAN_NO (+) = (CASE WHEN POTORDR9.PO_ORDER_LNO = 0 THEN POTORDR9.PO_ORDER_NO ELSE 'X' END)" & vbCrLf _
                & "   and POTORDR2.PO_ORDER_NO (+) = POTORDR9.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR2.PO_ORDER_LNO (+) = POTORDR9.PO_ORDER_LNO" & vbCrLf _
                & "   and POTORDR1.PO_ORDER_NO (+) = POTORDR9.PO_ORDER_NO" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE (+) = DPTPLAN1.ITEM_CODE"
            Create_TDA(.Tables.Add, "POTORDR9", "**", 0, False, "", 3)



            ' SOTORDRM

            LM = YPPD(1)
            ASCMAIN1.sql = "Select NVL(SOTMKTC1_CUST_CODE.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,NVL(SOTTCLS1.MARKET_CODE,'?'))) MARKET_CODE, SOTORDR2.ITEM_CODE" & vbCrLf _
            & ", SUM (CASE WHEN SOTORDR1.ORDR_SHIP_DATE <= '" & Format(LM, "dd-MMM-yyyy") & "' THEN " _
            & " NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0) ELSE 0 END) SOPD" & vbCrLf
            For P = 0 To 25
                Dim TM As Date = YPFD(P)
                ASCMAIN1.sql &= ", SUM (CASE WHEN SOTORDR1.ORDR_SHIP_DATE " _
                & " > '" & Format(LM, "dd-MMM-yyyy") & "'" _
                & " AND SOTORDR1.ORDR_SHIP_DATE <= '" & Format(TM, "dd-MMM-yyyy") & "'" _
                & " THEN NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0) ELSE 0 END) SO" & Format(P, "00") & vbCrLf
                LM = TM
            Next
            ASCMAIN1.sql &= " from SOTORDR2,SOTORDR1,ARTCUST1,SOTTCLS1," & ICTITEMS & " ICTITEMS,SOTMKTC1, SOTMKTC1 SOTMKTC1_CUST_CODE" & vbCrLf _
            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO " & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE " & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE " & vbCrLf _
            & "   and SOTMKTC1_CUST_CODE.CUST_CODE (+) = SOTORDR2.CUST_CODE" & vbCrLf _
            & "   and SOTMKTC1.MARKET_CODE (+) = NVL(SOTTCLS1.MARKET_CODE,'?')" & vbCrLf _
            & "   and SOTORDR2.ITEM_CODE = ICTITEMS.ITEM_CODE " & vbCrLf _
            & "   and SOTORDR2.ORDR_STATUS >= 'O' AND SOTORDR2.ORDR_STATUS <= 'P' " & vbCrLf _
            & "   and SOTORDR2.WHSE_CODE in (Select WHSE_CODE from ICTWHSE1 where NVL(WHSE_MRP_EXC_IND,'0') <> '1')" & vbCrLf _
            & " group by NVL(SOTMKTC1_CUST_CODE.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,NVL(SOTTCLS1.MARKET_CODE,'?'))), SOTORDR2.ITEM_CODE"
            Create_TDA(.Tables.Add, "SOTORDRM", "**", 0, False, "", 2)


            ' DPTPLANS

            ASCMAIN1.sql = ""
            For S As Int16 = 1 To 2
                Dim DATA_COLUMN As String = ""
                If S = 1 Then
                    DATA_COLUMN = "ORDR_QTY_SHIP"
                Else
                    DATA_COLUMN = "QTY_SOLD"
                End If

                ASCMAIN1.sql = "Select ITEM_CODE, '" & IIf(S = 1, "I", "T") & "' DATA_TYPE"
                For PRD As Integer = 0 To 12

                    Dim YP As String = YPP(12 - PRD, 0)
                    'CS &= "+ISNULL(S" & Format(P, "00") & ",0)"
                    ASCMAIN1.sql &= ", SUM (DECODE(OPS_YYYYPP,'" & YP & "'," & DATA_COLUMN & ",0)) S" & Format(PRD, "00")
                Next

                Dim TABLE_NAME As String = ""
                If S = 1 Then
                    TABLE_NAME = "SATSSUMI"
                Else
                    TABLE_NAME = "RSTRETL1"
                End If
                ASCMAIN1.sql &= " from " & TABLE_NAME _
                & " where OPS_YYYYPP between '" & YPP(12, 0) & "' and '" & YPP(0, 0) & "'" _
                & " and ITEM_CODE = :PARM1" _
                & " group by ITEM_CODE"

                'If S = 1 Then
                'ASCMAIN1.sql &= " UNION "
                ' End If
                sqlDPTPLANS(S) = ASCMAIN1.sql

            Next

            Create_TDA(.Tables.Add, "DPTPLANS", "**", 0, False, "V", 2)



            ' DPTPLANO

            ASCMAIN1.sql = "Select ICTITEMS.ITEM_CODE from " & ICTITEMS & " ICTITEMS"
            Create_TDA(.Tables.Add, "DPTPLANO", "**", 0, False, "", 1)
            Dim T As String
            For Each COL_TYPE As String In New String() {"OS", "GP"} ' {"OS", "GP", "SO", "SI", "ST"}
                T = ""
                For P = -1 To 25
                    Dim COLUMN_NAME As String = COL_TYPE & IIf(P = -1, "PD", Format(P, "00"))
                    Dim DC As DataColumn =
                    .Tables("DPTPLANO").Columns.Add(COLUMN_NAME, GetType(System.Int32))
                    T = "+ISNULL(" & COLUMN_NAME & ",0)"
                Next
                .Tables("DPTPLANO").Columns.Add(COL_TYPE, GetType(System.Int32), Mid(T, 2))
            Next

            ' DPTPLANX

            ASCMAIN1.sql = "Select ICTITEMS.* from " & ICTITEMS & " ICTITEMS"
            Create_TDA(.Tables.Add, "DPTPLANX", "**", 0, False)

            .Relations.Add("ICTITEM1_DPTPLANX",
                           .Tables("ICTITEM1").Columns("ITEM_CODE"),
                           .Tables("DPTPLANX").Columns("ITEM_CODE"))
            For Each COLUMN_NAME As String In New String() _
            {"ITEM_DESC", "ITEM_ABC_CODE", "DEPT_CODE", "MATL_CODE", "STYLE_CODE", "COLOR_CODE", "SIZE_CODE"}
                .Tables("DPTPLANX").Columns.Add(COLUMN_NAME, GetType(System.String), "PARENT(ICTITEM1_DPTPLANX)." & COLUMN_NAME)
            Next
            .Tables("DPTPLANX").Columns.Add("SEL")

            Create_Relation("DPTITMFM", "DPTPLANX", "ITEM_CODE")
            Create_Relation("POTORDRM", "DPTPLANX", "ITEM_CODE")
            ' Create_Relation("SOTORDRM", "DPTPLANX", "ITEM_CODE")
            Create_Relation("DPTPLANX", "SOTORDRM", "ITEM_CODE")
            Create_Relation("DPTPLANM", "DPTPLANX", "ITEM_CODE")
            Create_Relation("POTORDCM", "DPTPLANX", "ITEM_CODE")

            .Tables("DPTPLANX").Columns.Add("POS_STATUS_CODES", GetType(System.String))
            .Tables("DPTPLANX").Columns.Add("LAST_FC_P", GetType(System.Int32))
            .Tables("DPTPLANX").Columns.Add("ITEM_POS_MAX", GetType(System.Decimal))
            .Tables("DPTPLANX").Columns.Add("ITEM_POS_MIN", GetType(System.Decimal))

            For Each DTYP As String In New String() {"DEM", "SUP"}
                Dim DTs() As String = IIf(DTYP = "DEM", New String() {"FC", "PC", "SO"}, New String() {"PO", "PL"})
                For Each DT As String In DTs
                    For P = -1 To 25
                        Dim COLUMN_NAME As String = DT & IIf(P = -1, "PD", Format(P, "00"))
                        Dim DC As DataColumn =
                        .Tables("DPTPLANX").Columns.Add(COLUMN_NAME, GetType(System.Int32))
                        If DT = "FC" Or DT = "PO" Then
                            DC.Expression = "PARENT(" & IIf(DTYP = "DEM", "DPTITMFM", "POTORDRM") & "_DPTPLANX)." & COLUMN_NAME
                        End If
                        If DT = "SO" Then
                            'DC.Expression = "PARENT(SOTORDRM_DPTPLANX)." & COLUMN_NAME
                            DC.Expression = "SUM(CHILD." & COLUMN_NAME & ")"
                        End If
                        If DT = "PC" Or DT = "PL" Then
                            DC.Expression = "PARENT(" & IIf(DTYP = "DEM", "POTORDCM", "DPTPLANM") & "_DPTPLANX)." & COLUMN_NAME
                        End If
                    Next
                Next
                Dim strTOTAL As String = ""
                For P = -1 To 25
                    Dim SFX As String = IIf(P = -1, "PD", Format(P, "00"))
                    .Tables("DPTPLANX").Columns.Add(DTYP & SFX, GetType(System.Int32),
                        "ISNULL(" & DTs(0) & SFX & ",0)+ISNULL(" & DTs(1) & SFX & ",0)")
                    strTOTAL &= "+ISNULL(" & DTYP & SFX & ",0)"
                Next
                .Tables("DPTPLANX").Columns.Add("TOTAL_" & DTYP, GetType(System.Int32), strTOTAL)
            Next


            For Each DTYP As String In New String() {"PO", "SO"}

                Dim strTOTAL As String = ""
                For P = -1 To 25
                    Dim SFX As String = IIf(P = -1, "PD", Format(P, "00"))
                    strTOTAL &= "+ISNULL(" & DTYP & SFX & ",0)"
                Next
                .Tables("DPTPLANX").Columns.Add("TOTAL_" & DTYP, GetType(System.Int32), strTOTAL)
            Next

            ' NOTE THAT ONHPD IS ADDING BACK MTD SHIPMENTS AND THEN SUBTRACTING CURRENT MONTH FULL FORECAST 
            ' (OR MTD + COMMTD IF GREATER), BUT DOES NOT CONSIDER THE IMPACT OF NEG PD FORECAST RULES
            Dim ONHPD_CALC As String = "ISNULL(ONH,0)+ISNULL(MTD,0)+ISNULL(SUPPD,0)-ISNULL(DEMPD,0)"
            .Tables("DPTPLANX").Columns.Add("ONH" & "PD", GetType(System.Int32), ONHPD_CALC)
            .Tables("DPTPLANX").Columns.Add("POS" & "PD", GetType(System.Decimal))

            For P = 0 To 25
                Dim SOH As String = "ISNULL(ONH" & IIf(P = 0, "", Format(P - 1, "00")) & ",0)"
                'Dim DEM_ACT As String = "ISNULL(SO00,0)"
                If P = 0 Then
                    SOH &= "+ISNULL(MTD,0)+ISNULL(SUPPD,0)-ISNULL(DEMPD,0)"
                    'DEM_ACT &= "+ISNULL(MTD,0)+ISNULL(SOPD,0)"
                End If
                'SOH &= Replace(Replace("+ISNULL(SUP00,0)-ISNULL(DEM00,0)-IIF(ISNULL(DEM00,0)<{DEM_ACT},{DEM_ACT}-ISNULL(DEM00,0),0)", "{DEM_ACT}", DEM_ACT), "00", Format(P, "00"))
                SOH &= Replace("+ISNULL(SUP00,0)-ISNULL(DEM00,0)-ISNULL(DEMADJ00,0)", "00", Format(P, "00"))
                '  IIF(ISNULL(DEM00,0)<{DEM_ACT},{DEM_ACT}-ISNULL(DEM00,0),0)  
                .Tables("DPTPLANX").Columns.Add("DEMADJ" & Format(P, "00"), GetType(System.Int32))
                .Tables("DPTPLANX").Columns.Add("ONH" & Format(P, "00"), GetType(System.Int32), SOH)
            Next
            .Tables("DPTPLANX").Columns.Add("TOTAL_ONH", GetType(System.Int32))

            For P = 0 To 25
                .Tables("DPTPLANX").Columns.Add("POS" & Format(P, "00"), GetType(System.Decimal))
            Next
            .Tables("DPTPLANX").Columns.Add("TOTAL_POS", GetType(System.String))

            For P = -1 To 25
                Dim SFX As String = IIf(P = -1, "PD", Format(P, "00"))
                .Tables("DPTPLANX").Columns.Add("OS" & SFX, GetType(System.Int32))
            Next
            .Tables("DPTPLANX").Columns.Add("TOTAL_OS", GetType(System.Int32))

            For Each TREND_TYPE As String In New String() {"SHP", "DEM", "ONH", "PCT"}
                For P = 3 To 1 Step -1
                    .Tables("DPTPLANX").Columns.Add("TREND_" & TREND_TYPE & "_" & Format(P, "0"), GetType(System.Decimal))
                Next
            Next

            Create_TDA(.Tables.Add, "DPTITMF1", "*")
            Create_TDA(.Tables.Add, "ICTBRAN1", "*", -1, False)

            Create_TDA(.Tables.Add, "ICTDEPT1", "*", 0, False)
            Create_TDA(.Tables.Add, "ICTCOLL1", "*", 0, False)
            Create_TDA(.Tables.Add, "ICTMATL1", "*", 0, False)
            Create_TDA(.Tables.Add, "ICTCLAS1", "*", 0, False)
            Create_TDA(.Tables.Add, "ICTCATG1", "*", 0, False)


            ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE,ICTITEM1.DEPT_CODE,ICTITEM1.ITEM_CATGY_CODE" _
            & ", ICTITEM1.COLLECTION_CODE,ICTITEM1.ITEM_CLASS_CODE,ICTITEM1.STYLE_CODE,ICTCOLL1.HC_CODE,ICTITEM1.ITEM_BASIC_PROMO,ICTITEM1.ITEM_SNU_CODE" _
            & " from ICTITEM1,ICTCOLL1,DPTPROJ0 where ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" _
            & " and DPTPROJ0.OPS_YYYY (+) = '" & Mid(SN_next, 2, 4) & "'" _
            & " and DPTPROJ0.SEASON (+) = '" & Mid(SN_next, 1, 1) & "'" _
            & " and DPTPROJ0.ITEM_CODE (+) = ICTITEM1.ITEM_CODE"
            sqlICTITEMY = ASCMAIN1.sql
            ICTITEMY = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMY & " Add Primary Key (ITEM_CODE)")
            ASCDATA1.ExecuteSQL("Create Index I_" & ICTITEMY & "_1 on " & ICTITEMY & " (DEPT_CODE,ITEM_CATGY_CODE,COLLECTION_CODE,ITEM_CLASS_CODE,ITEM_CODE)")

            ASCMAIN1.sql = "Select Distinct DEPT_CODE,ITEM_CATGY_CODE,COLLECTION_CODE,ITEM_CLASS_CODE from " & ICTITEMY
            ASCMAIN1.sql = "Select Distinct HC_CODE,COLLECTION_CODE,ITEM_CATGY_CODE,ITEM_CLASS_CODE from " & ICTITEMY
            Create_TDA(.Tables.Add, "ICTITEMX", "**", 0, False, "", 0)

            ASCMAIN1.sql = "Select * from DPTMUPD0 where ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "DPTMUPD0", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "Select DPTMUPD1.*,DPTCRDM1.CRDM_DESC,ICTITEM1.ITEM_DESC,ICTITEM1.PROD_CODE" & vbCrLf _
                & " from DPTMUPD1,DPTCRDM1,ICTITEM1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = DPTMUPD1.ITEM_CODE and DPTCRDM1.CRDM_CODE = DPTMUPD1.CRDM_CODE"
            Create_TDA(.Tables.Add, "DPTMUPD1", "**", 0, False)

            ASCMAIN1.sql = "Select DPTMUPD2.*,DPTEXCM1.EXC_MSG_DESC,ICTITEM1.ITEM_DESC,ICTITEM1.PROD_CODE" & vbCrLf _
                & " from DPTMUPD2,DPTEXCM1,ICTITEM1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = DPTMUPD2.ITEM_CODE and DPTEXCM1.EXC_MSG_CODE = DPTMUPD2.EXC_MSG_CODE"
            Create_TDA(.Tables.Add, "DPTMUPD2", "**", 0, False)

            ASCMAIN1.sql = "Select GLTPARM2.* from GLTPARM2 " _
            & " where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -36) & "'" _
            & " and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "'"
            Create_TDA(.Tables.Add, "GLTPARM2", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select GLTPARM3.YYYYWW OPS_YYYYPP, GLTPARM3.LEGEND from GLTPARM3 " _
            & " where GLTPARM3.YYYYWW >= '" & ASCMAIN1.Week_Calc(ASCMAIN1.CYW, -110) & "'" _
            & " and GLTPARM3.YYYYWW <= '" & ASCMAIN1.CYW & "'"
            Create_TDA(.Tables.Add, "GLTPARM3", "**", 0, False, "", 1)


            ' DPTITMFX

            ASCMAIN1.sql = "Select ITEM_CODE, MARKET_CODE"
            For P = -1 To 25
                Dim COLUMN_NAME As String = "FCPD"
                Dim YP As String = "000000"
                If P >= 0 Then
                    COLUMN_NAME = "FC" & Format(P, "00")
                    YP = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, P)
                End If
                ASCMAIN1.sql &= ", SUM (DECODE(OPS_YYYYPP_FC,'" & YP & "',DPTITMF1.FORECAST,0)) " & COLUMN_NAME
            Next
            ASCMAIN1.sql &= " from DPTITMF1 where ITEM_CODE = :PARM1" _
            & " and OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" _
            & " group by ITEM_CODE, MARKET_CODE"
            ASCMAIN1.sql = $"Select X.*, SOTMKTC1.MARKET_DESC from SOTMKTC1, ({ASCMAIN1.sql}) X where SOTMKTC1.MARKET_CODE (+) = X.MARKET_CODE"
            Create_TDA(.Tables.Add, "DPTITMFX", "**", 0, False, "V", 2)
            For P = -1 To 25
                Dim COLUMN_NAME As String = IIf(P = -1, "FCPD", "FC" & Format(P, "00"))
                .Tables("DPTITMFX").Columns(COLUMN_NAME).DataType = GetType(System.Int32)
            Next


            ' DPTITMFS

            ASCMAIN1.sql = "Select ITEM_CODE, MARKET_CODE, 'X' DATA_TYPE"
            For P = 0 To 12
                ASCMAIN1.sql &= ", FORECAST Q" & Format(P, "00")
            Next
            ASCMAIN1.sql &= " from DPTITMF1 where ROWNUM < 1"
            DPTITMFS = ASCMAIN1.Temp_Table

            ASCMAIN1.sql = "Select ITEM_CODE, NVL(MARKET_CODE,'?') MARKET_CODE"
            For P = 0 To 12
                ASCMAIN1.sql &= ", SUM (DECODE(DATA_TYPE,'F',Q" & Format(P, "00") & ",0)) F" & Format(P, "00")
            Next
            For P = 0 To 12
                ASCMAIN1.sql &= ", SUM (DECODE(DATA_TYPE,'S',Q" & Format(P, "00") & ",0)) S" & Format(P, "00")
            Next
            ASCMAIN1.sql &= " from " & DPTITMFS & " DPTITMFS group by ITEM_CODE, NVL(MARKET_CODE,'?')"
            ASCMAIN1.sql = $"Select X.*, DECODE(X.MARKET_CODE,'*','All Markets',SOTMKTC1.MARKET_DESC) MARKET_DESC from SOTMKTC1, ({ASCMAIN1.sql}) X where SOTMKTC1.MARKET_CODE (+) = X.MARKET_CODE"
            Call Create_TDA(.Tables.Add, "DPTITMFS", "**", 0, False, "", 2)
            For P = 0 To 12
                .Tables("DPTITMFS").Columns("F" & Format(P, "00")).DataType = GetType(System.Int32)
                .Tables("DPTITMFS").Columns("S" & Format(P, "00")).DataType = GetType(System.Int32)
            Next


            ' ICTSTATX

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
            .Tables("ICTSTATX").Columns.Add("WHSE_TYPE")
            .Tables("ICTSTATX").Columns.Add("WHSE_MRP_EXC_IND")


            'ICTTRANX

            ASCMAIN1.sql = "SELECT ICTIADJ2.OPS_YYYYPP, ICTIADJ2.ITEM_CODE, ICTIADJ1.WHSE_CODE" _
            & ", ICTIADJ1.ADJ_NO TRAN_NO, ICTIADJ1.ADJ_SOURCE TRAN_SOURCE" _
            & ", ICTIADJ1.INIT_DATE, ICTIADJ1.INIT_OPER" _
            & ", ICTIADJ1.ADJ_DATE TRAN_DATE, 'X' TRAN_TYPE" _
            & ", ICTIADJ2.ADJ_QTY TRAN_QTY, ICTIADJ1.ADJ_NOTE TRAN_NOTE" _
            & " FROM ICTIADJ1,ICTIADJ2 WHERE ROWNUM < 1"
            ICTTRANX = ASCMAIN1.Temp_Table

            ASCMAIN1.sql = "Select ICTTRANX.* from " & ICTTRANX & " ICTTRANX"
            Call Create_TDA(.Tables.Add, "ICTTRANX", "**", 0, False)


            ' POTORDRX 

            ASCMAIN1.sql = "SELECT POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO" _
            & ", POTORDR2.PO_DATE_REQUIRED, POTORDR2.PO_QTY_OPN, POTORDR2.WHSE_CODE" _
            & ", POTORDR1.VEND_CODE, POTORDR1.PO_DATE_ORDERED, POTORDR1.PO_ORDER_TYPE, POTORDR1.PO_REFERENCE" _
            & ", POTORDR2.PO_DATE_REQUESTED" _
            & ", POTORDR2.PO_DATE_REQUIRED_MRP" _
            & ", POTORDR2.PO_DATE_ETD, POTORDR2.PO_DATE_ETD_NOTES, POTORDR2.ITEM_CODE" & vbCrLf _
            & " from POTORDR2,POTORDR1" _
            & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" _
            & "   and POTORDR2.PO_QTY_OPN <> 0" _
            & "   and POTORDR2.ITEM_CODE = :PARM1"
            Call Create_TDA(.Tables.Add, "POTORDRX", "**", 0, False, "V", 3)
            With .Tables("POTORDRX")
                .Columns.Add("OPO_QTY", GetType(System.Int32))
                .Columns.Add("INV_QTY", GetType(System.Int32))
            End With

            ASCMAIN1.sql = "Select POTORDR8.* from POTORDR8, POTORDR2, POTORDR1
                where POTORDR2.PO_ORDER_NO = POTORDR8.PO_ORDER_NO 
                and POTORDR2.PO_ORDER_LNO = POTORDR8.PO_ORDER_LNO and POTORDR2.ITEM_CODE = :PARM1
                and POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO
                and (:PARM2 = 'A' or (POTORDR2.PO_QTY_OPN > 0 AND POTORDR2.PO_STATUS = 'O'))"
            Create_TDA(.Tables.Add, "POTORDR8", "**", 0, True, "VV")


            ' ICTPINVX 

            ASCMAIN1.sql = $"Select * from {ICTPINVX} where ITEM_CODE = :PARM1"
            'Create_TDA(.Tables.Add("ICTPINVX"), ICTPINVX, "*")
            Create_TDA(.Tables.Add("ICTPINVX"), ICTPINVX, "**", ,, "V")
            With .Tables("ICTPINVX")
                .Columns.Add("OPO_QTY", GetType(System.Int32))
                .Columns.Add("INV_QTY", GetType(System.Int32), "IIF(CONTAINER_NO='Qty Open',0,PINV_QTY)")
            End With


            Create_Relation("POTORDRX", "ICTPINVX", "PO_ORDER_NO, PO_ORDER_LNO")
            With .Tables("POTORDRX")
                .Columns("INV_QTY").Expression = "SUM(CHILD(POTORDRX_ICTPINVX).INV_QTY)"
                .Columns("OPO_QTY").Expression = "SUM(CHILD(POTORDRX_ICTPINVX).OPO_QTY)"
            End With

            Create_Relation("POTORDRX", "POTORDR8", "PO_ORDER_NO,PO_ORDER_LNO")

            ASCMAIN1.sql = "Select '0' SEL, DPTPOSS1.* from DPTPOSS1"
            Create_TDA(.Tables.Add, "DPTPOSS1", "**", 0)


            ASCMAIN1.sql = "Select SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & ", SOTORDR1.WHSE_CODE, SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN, SUM(SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", NVL(SOTMKTC1_CUST_CODE.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,NVL(SOTTCLS1.MARKET_CODE,'?'))) MARKET_CODE" & vbCrLf _
            & ", SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
            & " from SOTORDR1,SOTORDR2,ARTCUST1,SOTTCLS1,SOTMKTC1, SOTMKTC1 SOTMKTC1_CUST_CODE" & vbCrLf _
            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE " & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE " & vbCrLf _
            & "   and SOTMKTC1_CUST_CODE.CUST_CODE (+) = SOTORDR1.CUST_CODE" & vbCrLf _
            & "   and SOTMKTC1.MARKET_CODE (+) = NVL(SOTTCLS1.MARKET_CODE,'?')" & vbCrLf _
            & "   and SOTORDR2.ITEM_CODE = :PARM1" & vbCrLf _
            & "   and SOTORDR2.ORDR_STATUS >= 'O' AND SOTORDR2.ORDR_STATUS <= 'P'" & vbCrLf _
            & " group by SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & ", SOTORDR1.WHSE_CODE, SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME" & vbCrLf _
            & ", NVL(SOTMKTC1_CUST_CODE.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,NVL(SOTTCLS1.MARKET_CODE,'?')))" & vbCrLf _
            & ", SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE"
            Create_TDA(.Tables.Add, "SOTORDRX", "**", 0, False, "V", 0)
            With .Tables("SOTORDRX")
                .Columns("ORDR_QTY_OPEN").DataType = GetType(System.Int32)
                .Columns("ORDR_QTY_PICK").DataType = GetType(System.Int32)
            End With


            ' NOTE A SIMILAR TABLE IS CALLED SOTINVHX IN MRP UPDATE
            ASCMAIN1.sql = "Select NVL(SOTMKTC1_CUST_CODE.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,NVL(SOTTCLS1.MARKET_CODE,'?'))) MARKET_CODE, SOTINVH2.ITEM_CODE" & vbCrLf _
                & ", SUM (DECODE(SOTINVH2.INV_TYPE,'I',SOTINVH2.ORDR_QTY_SHIP)) QTYI" & vbCrLf _
                & ", SUM (DECODE(SOTINVH2.INV_TYPE,'C',SOTINVH2.ORDR_QTY_SHIP)) QTYC" & vbCrLf _
                & " from SOTINVH2,ARTCUST1,SOTTCLS1," & ICTITEMS & " ICTITEMS,SOTMKTC1,ICTWHSE1,SOTMKTC1 SOTMKTC1_CUST_CODE" & vbCrLf _
                & " where SOTINVH2.ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "'" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
                & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                & "   and ICTWHSE1.WHSE_CODE = SOTINVH2.WHSE_CODE" & vbCrLf _
                & "   and NVL(ICTWHSE1.WHSE_MRP_EXC_IND,'0') <> '1'" & vbCrLf _
                & "   and SOTMKTC1_CUST_CODE.CUST_CODE (+) = SOTINVH2.CUST_CODE" & vbCrLf _
                & "   and SOTMKTC1.MARKET_CODE (+) = NVL(SOTTCLS1.MARKET_CODE,'?')" & vbCrLf _
                & "   and SOTINVH2.ITEM_CODE = ICTITEMS.ITEM_CODE " & vbCrLf _
                & " group by NVL(SOTMKTC1_CUST_CODE.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,NVL(SOTTCLS1.MARKET_CODE,'?'))), SOTINVH2.ITEM_CODE"
            Create_TDA(dst.Tables.Add, "SOTINVHM", "**", 0, False, "", 2)


            ASCMAIN1.sql = "Select * from DPTPROJ0 where OPS_YYYY = :PARM1 and SEASON = :PARM2"
            Create_TDA(.Tables.Add, "DPTPROJ0", "**", 0, False, "VV", 3)

            ASCMAIN1.sql = "Select * from DPTABCP0 where OPS_YYYY = :PARM1 and SEASON = :PARM2"
            Create_TDA(.Tables.Add, "DPTABCP0", "**", 0, False, "VV", 3)

            Create_TDA(.Tables.Add, "DPTABCP2", "*", 0)

            ASCMAIN1.sql = "Select * from ICTWHSE1"
            Create_TDA(.Tables.Add, "ICTWHSE1", "**", 0, False)

            ASCMAIN1.sql = "Select * from SOTMKTC1"
            Create_TDA(.Tables.Add, "SOTMKTC1", "**", 0, False)

            ASCMAIN1.sql = "Select COLLECTION_CODE, COLLECTION_NAME, '1' SEL" & vbCrLf _
                & " from ICTCOLL1 where BRAND_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTCOLLX", "**", 0, False, "V")


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

            Create_TDA(.Tables.Add, "ICTPORT2", "*", 0, False)
            Fill_Records("ICTPORT2")

            'ASCMAIN1.sql = $"Select * from {DPTMRPGO}"
            'Create_TDA(.Tables.Add, "DPTMRPGO", "**", 0, False, "", 2)
            'With .Tables("DPTMRPGO").Columns
            '    .Add("OVER_QTY", GetType(System.Int64), "ISNULL(EOM,0)-ISNULL(FC,0)")
            '    .Add("OVER_EXT_COST", GetType(System.Decimal), "ISNULL(OVER_QTY,0)*ISNULL(ITEM_COST_STD,0)")
            '    .Add("FC_CUR", GetType(System.Int64))
            '    .Add("FC_FUT", GetType(System.Int64))
            '    .Add("PO_CUR", GetType(System.Int64))
            '    .Add("PO_FUT", GetType(System.Int64))
            '    .Add("PP_CUR", GetType(System.Int64))
            '    .Add("PP_FUT", GetType(System.Int64))
            '    .Add("FC_EXT_COST", GetType(System.Decimal), "ISNULL(FC,0)*ISNULL(ITEM_COST_STD,0)")
            '    .Add("FCTM", GetType(System.Int64))
            '    .Add("FCTM_EXT_COST", GetType(System.Decimal), "ISNULL(FCTM,0)*ISNULL(ITEM_COST_STD,0)")
            '    .Add("POTM", GetType(System.Int64))
            '    .Add("POTM_EXT_COST", GetType(System.Decimal), "ISNULL(POTM,0)*ISNULL(ITEM_COST_STD,0)")
            '    .Add("PPTM", GetType(System.Int64))
            '    .Add("PPTM_EXT_COST", GetType(System.Decimal), "ISNULL(PPTM,0)*ISNULL(ITEM_COST_STD,0)")

            '    .Add("ZERO", GetType(System.String), "IIF(ISNULL(EOM,0)=0 AND ISNULL(FC_CUR,0)=0 AND ISNULL(FC_FUT,0)=0 AND ISNULL(PO_CUR,0)=0 AND ISNULL(PO_FUT,0)=0 AND ISNULL(OVER_QTY,0)=0 AND ISNULL(FC,0)=0,'0','1')")
            'End With

        End With

        Set_cmbYP("RYP", ASCMAIN1.CYP, 0, 18, 0)

        grdDPTPOSS1.DataSource = dst.Tables("DPTPOSS1")
        Setup_grdDPTPOSS1()

        Fill_Records("ICTWHSE1")
        Fill_Records("SOTMKTC1")
        Fill_Records("ICTMATL1")
        Fill_Records("ICTCLAS1")
        Fill_Records("ICTCATG1")
        Fill_Records("ICTBRAN1")
        Fill_Records("ICTCOLL1")
        Fill_Records("DPTABCP2")

        Fill_Records("GLTPARM2")
        Fill_Records("GLTPARM3")

        Fill_Records("DPTMUPD1")
        Fill_Records("DPTMUPD2")

        grdDPTPLANX.DataSource = dst.Tables("DPTPLANX")
        grdDPTITMFX.DataSource = dst.Tables("DPTITMFX")
        grdDPTITMFS.DataSource = dst.Tables("DPTITMFS")
        grdPOTORDR9.DataSource = dst.Tables("POTORDR9")

        grdDPTPLAND.DataSource = dst.Tables("DPTPLAND")
        grdDPTPLAN1.DataSource = dst.Tables("DPTPLAN1")
        grdDPTDBAL1.DataSource = dst.Tables("DPTDBAL1")

        grdGLTPARM2.DataSource = dst.Tables("GLTPARM2")
        Sort_grdColumns(grdGLTPARM2, "OPS_YYYYPP".ToLower)

        'grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")
        grdICTSTATX.DataSource = dst.Tables("ICTSTATX")
        grdICTTRANX.DataSource = dst.Tables("ICTTRANX")
        grdPOTORDRX.DataSource = dst.Tables("POTORDRX")
        grdSOTORDRX.DataSource = dst.Tables("SOTORDRX")

        grdDPTMUPD0.DataSource = dst.Tables("DPTMUPD0")
        grdDPTMUPD1.DataSource = dst.Tables("DPTMUPD1")
        grdDPTMUPD2.DataSource = dst.Tables("DPTMUPD2")
        grdSATAUTHX.DataSource = dst.Tables("SATAUTHX")

        grdICTCOLLX.DataSource = dst.Tables("ICTCOLLX")


        grdSOTALLO1.DataSource = dst.Tables("SOTALLO1")
        grdSOTALLO2.DataSource = dst.Tables("SOTALLO2")

        grdDPTMRPGO.DataSource = dst.Tables("DPTMRPGO")

        Setup_grdDPTPLANX()
        Setup_grdDPTITMFX()
        Setup_grdDPTITMFS()

        Create_Summary(grdDPTPLANX, "ITEM_CODE", "Count")

        Create_Summary(grdPOTORDR9, "PO_ORDER_NO", "Count")
        Create_Summary(grdPOTORDR9, "PO_QTY_COM")

        Create_Summary(grdDPTPLAN1, "PLAN_NO", "Count")
        Create_Summary(grdDPTPLAN1, "QTY_PLANNED")

        Create_Summary(grdDPTPLAND, "DATE_BALANCE", "Count")
        Create_Summary(grdDPTPLAND, "ACTIVITY")

        Create_Summary(grdPOTORDRX, "PO_ORDER_NO", "Count")
        Create_Summary(grdPOTORDRX, "PO_QTY_OPN")

        Create_Summary(grdSOTALLO2, New String() {"QTY_ALLO", "ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "QTY_LEFT", "QTY_BAL", "QTY_OVER"})

        Create_Summary(grdDPTMRPGO, "ITEM_CODE", "Count")
        Create_Summary(grdDPTMRPGO, New String() {"OVER_QTY", "OVER_EXT_COST", "EOM_EXT_COST"})
        Create_Summary(grdDPTMRPGO, New String() {"QTY_BEG", "AMT_BEG", "QTY_REC", "AMT_REC", "QTY_RTN", "AMT_RTN", "QTY_ADJ", "AMT_ADJ"})

        Create_Summary(grdDPTITMFX, "MARKET_CODE", "Count")
        'Create_Summary(grdDPTITMFX, "ITEM_CODE", "Count")
        For P = -1 To 25
            Dim COLUMN_NAME As String = IIf(P = -1, "FCPD", "FC" & Format(P, "00"))
            Create_Summary(grdDPTITMFX, COLUMN_NAME)
        Next

        With grdICTSTATX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() _
            {"WHSE_QTY_BEG", "WHSE_QTY_SHP", "WHSE_QTY_RTN", "WHSE_QTY_REC", "WHSE_QTY_ADJ",
             "WHSE_QTY_XFR", "WHSE_QTY_CON", "WHSE_QTY_RTV", "WHSE_QTY_PHY"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Color.Yellow
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() _
            {"WHSE_QTY_ON_HAND", "WHSE_QTY_ONPO", "WHSE_QTY_PLAN", "WHSE_QTY_OPEN",
             "WHSE_QTY_PICK", "WHSE_QTY_COMM", "WHSE_QTY_HOLD"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Color.Green
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next

        End With


        With grdDPTPLAN1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackColor2 = Color.LightBlue
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If gcol.Key = "DATE_REQUIRED" Or gcol.Key = "QTY_PLANNED" Then
                    gcol.CellActivation = Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                Else
                    gcol.CellActivation = Activation.NoEdit
                End If
            Next
        End With


        With grdDPTMRPGO.DisplayLayout.Bands("DPTMRPGO")
            .Columns("ITEM_CODE").Header.Fixed = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackColor2 = Color.LightGray
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If gcol.Key = "FC_CUR" Or gcol.Key = "FC_FUT" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                ElseIf gcol.Key = "PO_CUR" Or gcol.Key = "PO_FUT" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                ElseIf gcol.Key = "OVER_QTY" Or gcol.Key = "OVER_EXT_COST" Then
                    gcol.Header.Appearance.BackColor2 = Color.PaleVioletRed
                ElseIf gcol.Key = "QTY_BEG" Or gcol.Key = "AMT_BEG" Then
                    gcol.Header.Appearance.BackColor2 = Color.LimeGreen
                ElseIf gcol.Key = "QTY_REC" Or gcol.Key = "AMT_REC" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightSeaGreen
                ElseIf gcol.Key = "QTY_RTN" Or gcol.Key = "AMT_RTN" Then
                    gcol.Header.Appearance.BackColor2 = Color.Gold
                ElseIf gcol.Key = "QTY_ADJ" Or gcol.Key = "AMT_ADJ" Then
                    gcol.Header.Appearance.BackColor2 = Color.Orange
                Else
                    gcol.CellActivation = Activation.NoEdit
                End If
            Next
        End With
        Show_Filter(grdDPTMRPGO, True)
        grdDPTMRPGO.DisplayLayout.GroupByBox.Hidden = False

        Create_Summary(grdICTSTATX, "OPS_YYYYPP", "Count")
        Create_Summary(grdICTSTATX, "WHSE_CODE", "Count")
        Create_Summary(grdICTSTATX, "WHSE_QTY_BEG")
        Create_Summary(grdICTSTATX, "WHSE_QTY_SHP")
        Create_Summary(grdICTSTATX, "WHSE_QTY_RTN")
        Create_Summary(grdICTSTATX, "WHSE_QTY_REC")
        Create_Summary(grdICTSTATX, "WHSE_QTY_ADJ")
        Create_Summary(grdICTSTATX, "WHSE_QTY_XFR")
        Create_Summary(grdICTSTATX, "WHSE_QTY_CON")
        Create_Summary(grdICTSTATX, "WHSE_QTY_RTV")
        Create_Summary(grdICTSTATX, "WHSE_QTY_PHY")
        Create_Summary(grdICTSTATX, "WHSE_QTY_ON_HAND")
        Create_Summary(grdICTSTATX, "WHSE_QTY_ONPO")
        Create_Summary(grdICTSTATX, "WHSE_QTY_PLAN")
        Create_Summary(grdICTSTATX, "WHSE_QTY_OPEN")
        Create_Summary(grdICTSTATX, "WHSE_QTY_PICK")
        Create_Summary(grdICTSTATX, "WHSE_QTY_COMM")
        Create_Summary(grdICTSTATX, "WHSE_QTY_HOLD")

        Create_Summary(grdICTTRANX, "TRAN_DATE", "Count")
        Create_Summary(grdICTTRANX, "TRAN_QTY")

        Create_Summary(grdSOTORDRX, "ORDR_GROUP_NO", "Count")
        Create_Summary(grdSOTORDRX, "ORDR_QTY_OPEN")
        Create_Summary(grdSOTORDRX, "ORDR_QTY_PICK")

        Create_Summary(grdSATAUTHX, "CUST_CODE", "Count")
        Create_Summary(grdSATAUTHX, "STORE_COUNT")

        With grdDPTPLANX.DisplayLayout.Bands("DPTPLANX")
            .Groups(0).Header.Fixed = True
        End With

        'With grdDPTPLANX.DisplayLayout.Bands("DPTPLANX")
        '    .Columns("ITEM_CODE").Header.Fixed = True
        'End With
        With grdDPTITMFX.DisplayLayout.Bands("DPTITMFX")
            .Columns("MARKET_CODE").Header.Fixed = True
            .Columns("MARKET_DESC").Header.Fixed = True
        End With
        With grdDPTITMFS.DisplayLayout.Bands("DPTITMFS")
            .Columns("MARKET_CODE").Header.Fixed = True
            .Columns("MARKET_DESC").Header.Fixed = True
        End With

        With grdICTSTATX.DisplayLayout.Bands("ICTSTATX")
            .Columns("OPS_YYYYPP").Header.Fixed = True
            .Columns("WHSE_CODE").Header.Fixed = True
        End With

        With grdICTTRANX.DisplayLayout.Bands("ICTTRANX")
            .Columns("TRAN_DATE").Header.Fixed = True
        End With

        optBy.Value = "I"

        grdPOTORDRX.DisplayLayout.Bands(0).Columns("PO_QTY_OPN").CellAppearance.BackColor = Color.LightGreen

        optIT.Items(0).Appearance.ForeColor = Color.DarkViolet
        optIT.Items(0).Appearance.BackColor = Color.Violet
        optIT.Items(0).Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
        optIT.Items(1).Appearance.ForeColor = Color.Green
        optIT.Items(1).Appearance.BackColor = Color.Green
        optIT.Items(1).Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

        Bind_Controls(grpICTITEM1, "ICTITEM1")

        Set_Read_Only(grpICTITEM1, True)
        Set_Read_Only(grpLT, True)

        splMain.Visible = False

        Setup_DQ()


        'grdDPTPLANX.DisplayLayout.Override.TipStyleCell = UltraWinGrid.TipStyle.Hide
        '' set this value to however many milliseconds the tooltip delay should be
        'timer.Interval = 500
        '' when the timer ticks we want our method to be called
        'AddHandler timer.Tick, AddressOf OnTimerTick
        'tooltip_msg = String.Empty


        With chtTotals
            .Axis.X.ScrollScale.Visible = True
            .Axis.Y.ScrollScale.Visible = True

            .Axis.X.ScrollScale.Scale = 1 ' 0.25
            .Axis.Y.ScrollScale.Scale = 1 ' 0.25
            'Me.trkbrXAxis.Value = .Axis.X.ScrollScale.Scale * 100
            'Me.trkbrYAxis.Value = .Axis.Y.ScrollScale.Scale * 100
            .EnableCrossHair = True

            '.ColorModel.ModelStyle = ColorModels.CustomLinear '  CType(System.Enum.Parse(GetType(ColorModels), System.Enum.GetNames(GetType(ColorModels))(0)), ColorModels)
        End With

        ' splChart.Visible = False

        'ASCMAIN1.Add_Value_List(grdDPTITMFX, "MARKET_CODE", Nothing, New String() {":", "*: Total"})
        'ASCMAIN1.Add_Value_List(grdDPTITMFS, "MARKET_CODE", Nothing, New String() {":", "*: Total"})

        Dim DP_PARM_FC_FREEZE As String = Val(ROWs("DPTPARM1").Item("DP_PARM_FC_FREEZE") & "")
        Dim DP_PARM_FC_FREEZE_MOS As Int32 = Val(ROWs("DPTPARM1").Item("DP_PARM_FC_FREEZE_MOS") & "")
        If DP_PARM_FC_FREEZE = "1" Then
            For P = -1 To DP_PARM_FC_FREEZE_MOS
                Dim COLUMN_NAME As String = IIf(P = -1, "FCPD", "FC" & Format(P, "00"))
                With grdDPTITMFX.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                    .CellAppearance.BackColor = Color.Beige
                End With
            Next
        End If

        tabDetails.Tabs("grd").Visible = False
        Set_LYVF_Tab_Visible()

        'Set_Read_Only(grpITEM_CATGY_CODE, True)
        With grdDPTPLANX.DisplayLayout.Bands(0).Columns("ITEM_CODE")
            .Style = UltraWinGrid.ColumnStyle.EditButton
            .ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
            .CellButtonAppearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "CAMERA2")
        End With

        chkEditForecasts.Enabled = (ROWs("DPTPARM1").Item("DP_PARM_FC_MAINT_IN_PLAN") & "" = "1")

        grdDPTPLANX.DisplayLayout.Bands(0).Groups("PPD").Hidden = True
        Setup_Conditions()

        ' ASCMAIN1.Add_Value_List(grdPOTORDRX, "PO_STAGE")
        ASCMAIN1.Add_Value_List(grdICTSTATX, "WHSE_TYPE", Nothing, New String() {":", "F:Refurb", "R:Return", "S:Ship", "I:Intl", "X:InXit", "X:Unknown"})

        ASCMAIN1.Add_Value_List(grdDPTMRPGO, "ITEM_STATUS")

        AUDIT.Add("ICTITEM1", "*")

        ' MakeTransparent(chkShowAllItems)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                If Absx1.txtFor("ITEM_CODE").Text <> "" Then
                    Validate_Code("ITEM_CODE")
                Else
                    Validate_Code("BRAND_CODE")
                End If

                If EMsg = "" Then

                    If Not InquiryMode Then
                        If Absx1.txtFor("ITEM_CODE").Text <> "" Then
                            If Not ASCMAIN1.Logical_Lock("ICTITEM1", Absx1.txtFor("ITEM_CODE").Text) Then Exit Sub
                        Else
                            If Not ASCMAIN1.Logical_Lock("ICTBRAN1", Absx1.txtFor("BRAND_CODE").Text) Then Exit Sub
                        End If
                    End If

                    byItem = (Absx1.txtFor("ITEM_CODE").Text <> "")
                    If byItem Then
                        optBy.Value = "I"

                    Else
                        optBy.Value = "B"
                        chkHideTree.Checked = False
                    End If

                End If

            Case "Update"
                If ASCMAIN1.CLIENT = "INT" Then
                    If Absx1.dteFor("ITEM_PLAN_QUIET_ZONE_DATE").Value & "" <> "" Then
                        Absx1.optFor("ITEM_PLAN_QUIET_ZONE_TYPE").Value = "2"
                        ' rowict.Item("ITEM_PLAN_QUIET_ZONE_TYPE") = "2"
                    Else
                        If Absx1.optFor("ITEM_PLAN_QUIET_ZONE_TYPE").Value & "" = "2" Then
                            Absx1.optFor("ITEM_PLAN_QUIET_ZONE_TYPE").Value = "0"
                        End If
                    End If

                    Dim ITEM_PLAN_QUIET_ZONE_TYPE As String = Absx1.optFor("ITEM_PLAN_QUIET_ZONE_TYPE").Value & "" ' rowASFBASE1.Item("ITEM_PLAN_QUIET_ZONE_TYPE") & "" '  Absx1.optFor("ITEM_PLAN_QUIET_ZONE_TYPE").Value & ""
                    If ITEM_PLAN_QUIET_ZONE_TYPE = "2" Then
                        If Absx1.dteFor("ITEM_PLAN_QUIET_ZONE_DATE").Value & "" = "" Then
                            EMsg &= vbCr & "You Must Specify a real date if using the Do Not Plan Before feature"
                        End If
                    End If
                End If

                Dim ITEM_ABC_CODE As String = txtITEM_ABC_CODE.Text
                If ITEM_ABC_CODE <> "" Then
                    Dim rowDPTABCP1 As DataRow = LookUp("DPTABCP1", ITEM_ABC_CODE)
                    If rowDPTABCP1 Is Nothing Then
                        EMsg &= vbCr & "Invalid value for ABC Code"
                    End If
                End If

                Dim ITEM_BUFFER_QTY As Int32 = Val(Absx1.numFor("ITEM_BUFFER_QTY").Value & "")
                Dim ITEM_BUFFER_PCT As Int32 = Val(Absx1.numFor("ITEM_BUFFER_PCT").Value & "")
                If ITEM_BUFFER_QTY <> 0 And ITEM_BUFFER_PCT <> 0 Then
                    EMsg &= vbCr & "Cannot Specify both Buffer Qty and Pct"
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
                Load_Record()
                Mode_Settings(True)

            Case "Report"
                Print_Report()

            Case "Generate Plans"
                Generate_Plans()

            Case "Done"
                'Absx1.txtFor("BRAND_NAME").Text = "Burberry"
                Mode_Settings(False)


                ' Item Master

            Case "Edit"
                Mode_Settings_Edit(True)

            Case "Update"
                Update_Record()
                Mode_Settings_Edit(False)

            Case "Cancel"
                Mode_Settings_Edit(False)


            Case "Edit Plans"
                Mode_Settings_Edit_Plans(True)

            Case "Update Plans"
                Update_Record_Plans()
                Reset_Plans()
                Mode_Settings_Edit_Plans(False)

            Case "Cancel Plans"
                Reset_Plans()
                Mode_Settings_Edit_Plans(False)
        End Select

    End Sub

    Sub Mode_Settings_Edit(tf)

        modes_Edit = tf

        Set_Read_Only(grpICTITEM1, Not tf)
        Set_Read_Only(grpLT, True)

        If tf Then
            tabDetails.SelectedTab = tabDetails.Tabs("Item Master")

            ' WE NEED TO REFRESH THE ITEM FROM THE ITEM MASTER
            ' CONSTRAIN THE UPDATE TO JUST THOSE FIELDS THAT ARE REPRESENTED IN THE TAB
            ' Do NOT PERMIT UPDATES TO SPECIFIC FIELDS, LIKE BASIC/PROMO AND ITEM SNU

            Set_Read_Only_for_ctl(Absx1.optFor("ITEM_BASIC_PROMO"), True)
            Set_Read_Only_for_ctl(Absx1.optFor("ITEM_SNU_CODE"), True)

            Set_ABC_Parameters_ReadOnly()

        End If



        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Item Master")
                    .Items("Edit").Settings.Enabled = IIf(modes_Edit, DefaultableBoolean.False, DefaultableBoolean.True)
                    .Items("Update").Settings.Enabled = IIf(modes_Edit, DefaultableBoolean.True, DefaultableBoolean.False)
                    .Items("Cancel").Settings.Enabled = IIf(modes_Edit, DefaultableBoolean.True, DefaultableBoolean.False)
                End With
            End With
        End If

        If modes_Edit Then

        End If
    End Sub

    Sub Mode_Settings_Edit_Plans(tf)

        modes_Edit_Plans = tf

        'Set_Read_Only(grpICTITEM1, Not tf)
        With grdDPTPLAN1.DisplayLayout.Override
            If tf Then
                .AllowAddNew = AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.True
            Else
                .AllowAddNew = AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End If
        End With


        With grdDPTPLAN1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.CellActivation = Activation.AllowEdit Then
                    If tf Then
                        gcol.CellAppearance.BackColor = System.Drawing.Color.Yellow
                    Else
                        gcol.CellAppearance.BackColor = System.Drawing.Color.Empty
                    End If
                End If
            Next

        End With

        If tf Then
            tabDetails.SelectedTab = tabDetails.Tabs("Plans")


            ' WE NEED TO REFRESH THE ITEM FROM THE ITEM MASTER
            ' CONSTRAIN THE UPDATE TO JUST THOSE FIELDS THAT ARE REPRESENTED IN THE TAB
            ' Do NOT PERMIT UPDATES TO SPECIFIC FIELDS, LIKE BASIC/PROMO AND ITEM SNU

            'Set_Read_Only_for_ctl(Absx1.optFor("ITEM_BASIC_PROMO"), True)
            'Set_Read_Only_for_ctl(Absx1.optFor("ITEM_SNU_CODE"), True)
        End If



        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Plans")
                    .Items("Edit Plans").Settings.Enabled = IIf(modes_Edit_Plans, DefaultableBoolean.False, DefaultableBoolean.True)
                    .Items("Update Plans").Settings.Enabled = IIf(modes_Edit_Plans, DefaultableBoolean.True, DefaultableBoolean.False)
                    .Items("Cancel Plans").Settings.Enabled = IIf(modes_Edit_Plans, DefaultableBoolean.True, DefaultableBoolean.False)
                End With
            End With
        End If

        If modes_Edit_Plans Then

        End If
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    '.Items("Update").Settings.Enabled = iScreenMode
                    '.Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Generate Plans").Settings.Enabled = iScreenMode
                    .Items("Generate Plans").Visible = False
                    .Items("Report").Settings.Enabled = iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                End With
                .Groups("Display Options").Visible = ScreenMode
                .Groups("Position Legend").Visible = ScreenMode
                .Groups("Report Options").Visible = ScreenMode
                .Groups("Picture").Visible = ScreenMode

                .Groups("Item Master").Visible = False
                .Groups("Plans").Visible = False

                If ASCMAIN1.DBS_COMPANY = "INT" Then
                    .Groups("Item Master").Visible = ScreenMode AndAlso byItem AndAlso Not InquiryMode
                End If

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        splMain.Visible = tf

        lblFindITEM_CODE.Visible = tf And Not byItem
        txtFindITEM_CODE.Visible = tf And Not byItem
        Set_Read_Only(txtFindITEM_CODE, False)
        Setup_Conditions()

        With grdDPTMUPD1.DisplayLayout.Bands(0)
            .Columns("ITEM_CODE").Hidden = ScreenMode
            .Columns("ITEM_DESC").Hidden = ScreenMode
        End With
        With grdDPTMUPD2.DisplayLayout.Bands(0)
            .Columns("ITEM_CODE").Hidden = ScreenMode
            .Columns("ITEM_DESC").Hidden = ScreenMode
        End With

        grdDPTMUPD1.DisplayLayout.GroupByBox.Hidden = ScreenMode
        grdDPTMUPD2.DisplayLayout.GroupByBox.Hidden = ScreenMode
        grdDPTMUPD1.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdDPTMUPD2.DisplayLayout.Bands(0).SortedColumns.Clear()

        Set_BI()

        If ScreenMode Then
            splMessages1.Parent = tabDetails.Tabs("Messages").TabPage ' splMessages1.Panel1
            tabTG.SelectedTab = tabTG.Tabs("Grid")

            Set_Read_Only_for_ctl(Absx1.optFor("ITEM_SNU_CODE"), True)
            Set_Read_Only_for_ctl(Absx1.optFor("ITEM_BASIC_PROMO"), True)

        Else
            splMessages1.Parent = spl.Panel2
            Clear_Record()
            chkShowTab.Checked = False
        End If

        splMessages1.SplitterDistance = splMessages1.Width / 2
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
        {"ICTBRAN1", "ICTCOLL1", "DPTPLANX", "DPTPLANO", "ICTITEM1", "POTORDRX", "ICTPINVX", "POTORDR8",
         "ICTSTATX", "ICTTRANX", "DPTITMFX", "DPTITMFS", "DPTPLANS", "SOTALLO1", "SOTALLO2",
         "SOTORDRX", "SOTORDRM", "SOTINVHM"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        PictureBox1.Image = Nothing
        UltraExplorerBar1.Groups("Picture").Expanded = False

        'grdDPTMUPD1.DataSource = Nothing
        'grdDPTMUPD2.DataSource = Nothing

        'grdDPTMUPD1.DataSource = dst.Tables("DPTMUPD1")
        'grdDPTMUPD2.DataSource = dst.Tables("DPTMUPD2")

        Dim dvw As DataView = Nothing
        dvw = DirectCast(grdDPTMUPD1.DataSource, DataTable).DefaultView
        'dvw.RowFilter = "1<>1"
        grdDPTMUPD1.DisplayLayout.Bands(0).SortedColumns.Clear()
        'grdDPTMUPD1.DisplayLayout.Bands(0).SortedColumns.Add("CRDM_DESC", False, True)
        grdDPTMUPD1.DisplayLayout.Bands(0).SortedColumns.Add("ITEM_CODE", False)
        dvw.RowFilter = ""

        tvwDQ.Tag = ""

        dvw = DirectCast(grdDPTMUPD2.DataSource, DataTable).DefaultView
        'dvw.RowFilter = "1<>1"
        grdDPTMUPD2.DisplayLayout.Bands(0).SortedColumns.Clear()
        ' grdDPTMUPD2.DisplayLayout.Bands(0).SortedColumns.Add("EXC_MSG_DESC", False, True)
        grdDPTMUPD2.DisplayLayout.Bands(0).SortedColumns.Add("ITEM_CODE", False)
        dvw.RowFilter = ""

        Absx1.txtFor("BRAND_CODE").Text = ""

        ITEM_CODE = ""

        Mode_Settings_Edit(False)
        Mode_Settings_Edit_Plans(False)

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Dim sql As String = sqlICTITEMY
        If byItem Then
            Absx1.txtFor("BRAND_CODE").Text = ""
            sql &= " and ICTITEM1.ITEM_CODE = '" & Absx1.txtFor("ITEM_CODE").Text & "'"
        Else
            sql &= " and ICTCOLL1.BRAND_CODE = '" & Absx1.txtFor("BRAND_CODE").Text & "'"
            If chkActiveOnly.Checked Then
                If chkUseNextSN.Checked Then
                    sql &= " and NVL(DPTPROJ0.ITEM_CATGY_CODE,'I') <> 'I'"
                Else
                    sql &= " and ICTITEM1.ITEM_STATUS = 'A'"
                End If
            End If

            Fill_Records("ICTCOLLX", Absx1.txtFor("BRAND_CODE").Text)
        End If


        If chkUseNextSN.Checked Then
            sql = Replace(sql, "ICTITEM1.ITEM_CATGY_CODE", "NVL(DPTPROJ0.ITEM_CATGY_CODE,'I') ITEM_CATGY_CODE")
        End If

        ASCDATA1.ExecuteSQL("Truncate Table " & ICTITEMY)
        ASCDATA1.ExecuteSQL("Insert into " & ICTITEMY & " " & sql)

        Fill_Records("ICTITEMX", , , sql)
        'Dim X As String = Absx1.txtFor("BRAND_CODE").Text

        If chkUseNextSN.Checked Then
            Fill_Records("DPTPROJ0", New String() {Mid(SN_next, 2, 4), Mid(SN_next, 1, 1)})
            Fill_Records("DPTABCP0", New String() {Mid(SN_next, 2, 4), Mid(SN_next, 1, 1)})
            'For Each rowICTITEMX As DataRow In dst.Tables("ICTITEMX").Rows
            '    rowICTITEMX.Item("ITEM_CATGY_CODE") = "I"
            'Next
            'For Each rowDPTPROJ0 As DataRow In dst.Tables("DPTPROJ0").Rows
            '    Dim ITEM_CODE As String = rowDPTPROJ0.Item("ITEM_CODE")
            '    Dim rowICTITEMX As DataRow = dst.Tables("ICTITEMX").Rows.Find(ITEM_CODE)
            '    If rowICTITEMX IsNot Nothing Then
            '        rowICTITEMX.Item("ITEM_CATGY_CODE") = rowDPTPROJ0.Item("ITEM_CATGY_CODE")
            '    End If
            'Next
        End If

        EnforceConstraints(True)

        Sort_grdColumns(grdDPTPLANX, "ITEM_CODE")

        Generate_Inquiry()

        If byItem Then
            tabTG.SelectedTab = tabTG.Tabs("Tree")
        End If

        If Absx1.txtFor("ITEM_CODE").Text <> "" Then
            chkItemDetails.Checked = True
            If tabTG.SelectedTab.Key = "Tree" Then
                tvwDQ.ExpandAll()
                If tvwDQ.Nodes.Count <> 0 Then
                    Dim tnode As UltraWinTree.UltraTreeNode = tvwDQ.Nodes(0)
                    Do While tnode.HasNodes
                        tnode = tnode.Nodes(0)
                    Loop
                    Click_Node(tnode)
                End If
            Else
                Dim tnode As UltraWinTree.UltraTreeNode = tvwDQ.Nodes(0)
                Click_Node(Nothing)
            End If
        End If
        tabProjections.Tabs("Forecasts by Market").Selected = True
        tabProjections.Tabs("Forecasts vs Actual").Selected = True

        '  If byItem Then Setup_Item()

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()
        Update_Record_TDA("ICTITEM1")
        CommitTrans("Update Complete")

    End Sub

    Sub Update_Record_Plans()

        Dim ITEM_CODE As String = grdDPTPLANX.ActiveRow.Cells("ITEM_CODE").Text

        BeginTrans()
        Update_Record_TDA("DPTPLAN1")

        ASCMAIN1.sql = "DELETE FROM DPTMUPD0 where ITEM_CODE = :PARM1"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New String() {ITEM_CODE})

        ASCMAIN1.sql = "INSERT INTO DPTMUPD0 (ITEM_CODE, DATE_REQ, QTY_PLN, PO_ORDER_NO, VEND_CODE, WHSE_CODE, NOTES, SD, P_INDEX)" & vbCrLf _
            & "Select ITEM_CODE, DATE_REQUIRED, QTY_PLANNED, PLAN_NO, VEND_CODE, TO_WHSE, 'System Plan', 'P', ROWNUM" & vbCrLf _
            & "FROM DPTPLAN1 WHERE ITEM_CODE = :PARM1"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New String() {ITEM_CODE})

        CommitTrans("Update Complete")

    End Sub

    Sub Reset_Plans()

        EnforceConstraints(False)
        Fill_Records("DPTPLAN1")
        Fill_Records("DPTPLANM")
        EnforceConstraints(True)

        Dim ITEM_CODE As String = grdDPTPLANX.ActiveRow.Cells("ITEM_CODE").Text



        With dst.Tables("DPTDBAL1")
            For Each rowDPTPLAND As DataRow In dst.Tables("DPTPLAND").Select("")
                Dim SOURCE As String = rowDPTPLAND.Item("SOURCE") & ""
                If SOURCE.StartsWith("Plan ") Then
                    rowDPTPLAND.Delete()
                End If
            Next
            ' ASCDATA1.DeleteRows("DPTPLAND", "SOURCE LIKE 'Plan *'")

            '.Rows.Clear()
            'grdDPTDBAL1.DataSource = Nothing
            'If .Columns.Count > 4 Then
            '    For I = .Columns.Count - 1 To 4 Step -1
            '        .Columns.Remove(.Columns(I).ColumnName)
            '    Next
            'End If

            Dim WHSE_CODEs As New List(Of String)
            For Each ROW As DataRow In ASCDATA1.SelectDistinct _
            (dst.Tables("DPTPLAND").Select("ITEM_CODE = '" & ITEM_CODE & "'"), "WHSE_CODE").Rows
                Dim WHSE_CODE As String = ROW.Item("WHSE_CODE")
                WHSE_CODEs.Add(WHSE_CODE)
                'Dim COLUMN_NAME As String = "WHSE_" & WHSE_CODE
                '.Columns.Add(COLUMN_NAME, GetType(System.Int32))
            Next


            For Each ROW As DataRow In dst.Tables("DPTPLAND").Select("ITEM_CODE = '" & ITEM_CODE & "'")
                Dim WHSE_CODE As String = ROW.Item("WHSE_CODE")
                Dim COLUMN_NAME As String = "WHSE_" & WHSE_CODE
                Dim DATE_BALANCE As Date = ROW.Item("DATE_BALANCE")
                Dim ACTIVITY As Int32 = ROW.Item("ACTIVITY")
                Dim rowDPTDBAL1 As DataRow = .Rows.Find(New Object() {ITEM_CODE, DATE_BALANCE})
                If rowDPTDBAL1 Is Nothing Then
                    rowDPTDBAL1 = dst.Tables("DPTDBAL1").NewRow
                    rowDPTDBAL1.Item("ITEM_CODE") = ITEM_CODE
                    rowDPTDBAL1.Item("DATE_BALANCE") = DATE_BALANCE
                    dst.Tables("DPTDBAL1").Rows.Add(rowDPTDBAL1)
                End If
                rowDPTDBAL1.Item("ACTIVITY") = Val(rowDPTDBAL1.Item("ACTIVITY") & "") + ACTIVITY
                ' rowDPTDBAL1.Item("BALANCE") = Val(rowDPTDBAL1.Item("BALANCE") & "") + ACTIVITY
                rowDPTDBAL1.Item(COLUMN_NAME) = Val(rowDPTDBAL1.Item(COLUMN_NAME) & "") + ACTIVITY
            Next

            Dim BALANCE As Int32 = 0
            For Each rowDPTDBAL1 As DataRow In dst.Tables("DPTDBAL1").Select("ITEM_CODE = '" & ITEM_CODE & "'", "DATE_BALANCE")
                BALANCE += Val(rowDPTDBAL1.Item("ACTIVITY") & "")
                rowDPTDBAL1.Item("BALANCE") = BALANCE
            Next

            For Each WHSE_CODE As String In WHSE_CODEs
                Dim COLUMN_NAME As String = "WHSE_" & WHSE_CODE
                BALANCE = 0
                For Each rowDPTDBAL1 As DataRow In dst.Tables("DPTDBAL1").Select("ITEM_CODE = '" & ITEM_CODE & "'", "DATE_BALANCE")
                    BALANCE += Val(rowDPTDBAL1.Item(COLUMN_NAME) & "")
                    rowDPTDBAL1.Item(COLUMN_NAME) = BALANCE
                Next
            Next
        End With
        Sort_grdColumns(grdDPTDBAL1, "DATE_BALANCE")


        Fill_Records("DPTMUPD0", ITEM_CODE)
        Dim QTY_BAL As Int64 = 0
        For Each row As DataRow In dst.Tables("DPTMUPD0").Select("", "DATE_REQ")
            QTY_BAL += Val(row.Item("QTY_OH") & "") _
                - Val(row.Item("QTY_REQ") & "") _
                + Val(row.Item("QTY_ORD") & "") _
                + Val(row.Item("QTY_PLN") & "")
            row.Item("QTY_BAL") = QTY_BAL
        Next
        Sort_grdColumns(grdDPTMUPD0, "DATE_REQ")



        For Each rowDPTPLAN1 As DataRow In dst.Tables("DPTPLAN1").Select("", "")

            For P As Integer = -1 To 25
                'Dim COLUMN_NAME As String = IIf(P = -1, "FCPD", "FC" & Format(P, "00"))
                'Dim QTY_PLANNED As Integer = Val(rowDPTITMFX.Item(COLUMN_NAME) & "")
                'If FORECAST <> 0 Then

                'End If
            Next
        Next

        Calculate_Position(ITEM_CODE)
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        '  Delete_Records("table")
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

            Case "Load"
                Absx1.txtFor("ITEM_CODE").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdDPTPLANX, "SSBBS", "Show Trend", "Show Past Due", "Show Image", "Item Status Inquiry", "Items with Forecasts Only")
        Load_Popup_Menu(grdDPTPOSS1, "CC", "BackColor", "ForeColor")
        Load_Popup_Menu(grdPOTORDRX, "B", "PO Inquiry")
        Load_Popup_Menu(chtTotals, "B", "Export to JPG")
        Load_Popup_Menu(chtTrend, "B", "Export to JPG")
        Load_Popup_Menu(grdDPTMUPD1, "SS", "Show Filter", "Show Groupbox")
        Load_Popup_Menu(grdDPTMUPD2, "SS", "Show Filter", "Show Groupbox")
        Load_Popup_Menu(grdPOTORDR9, "B", "PO Inquiry")
        Load_Popup_Menu(grdICTCOLLX, "BB", "Select All", "De-Select All")
        Load_Popup_Menu(grdDPTMRPGO, "SSB", "Show Filter", "Show Groupbox", "Demand Planning")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        If e.SourceControl.GetType.Equals(GetType(UltraWinChart.UltraChart)) Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            If e.Tool.Key <> "grdDPTPLANX" And e.Tool.Key <> "grdDPTPOSS1" Then
                'Dim tlb_sbt As UltraWinToolbars.StateButtonTool
                'tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
                'tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
                'tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
                'tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
            End If

            Select Case e.SourceControl.Name
                Case "grdDPTPLANX"

                Case "grdDPTPOSS1"
                    Dim tlb_cpt As UltraWinToolbars.PopupColorPickerTool

                    tlb_cpt = DirectCast(e.Tool.ToolbarsManager.Tools("BackColor"), UltraWinToolbars.PopupColorPickerTool)
                    tlb_cpt.SelectedColor = Color.FromArgb(Val(grd.ActiveRow.Cells("POS_RBG_BACKCOLOR").Value & ""))
                    tlb_cpt = DirectCast(e.Tool.ToolbarsManager.Tools("ForeColor"), UltraWinToolbars.PopupColorPickerTool)
                    tlb_cpt.SelectedColor = Color.FromArgb(Val(grd.ActiveRow.Cells("POS_RBG_FORECOLOR").Value & ""))
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Select Case e.Tool.Key
            Case "Select All", "De-Select All"
                For Each row As DataRow In dst.Tables("ICTCOLLX").Select("")
                    row.Item("SEL") = IIf(e.Tool.Key = "Select All", "1", "0")
                Next
                Exit Sub

            Case "Export to JPG"
                Dim IMAGE_FILENAME = ASCMAIN1.Next_Control_No("DPFPLAN1.IMAGE_FILENAME") & ".JPG"
                chtTotals.SaveTo(IMAGE_FILENAME, System.Drawing.Imaging.ImageFormat.Jpeg)
                Show_Document(IMAGE_FILENAME)
                Exit Sub

        End Select

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Show Trend"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grdDPTPLANX.DisplayLayout.Bands(0).Groups("TREND_L").Hidden = Not tlb_sbt.Checked
                grdDPTPLANX.DisplayLayout.Bands(0).Groups("TREND_3").Hidden = Not tlb_sbt.Checked
                grdDPTPLANX.DisplayLayout.Bands(0).Groups("TREND_2").Hidden = Not tlb_sbt.Checked
                grdDPTPLANX.DisplayLayout.Bands(0).Groups("TREND_1").Hidden = Not tlb_sbt.Checked
            Case "Show Past Due"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
                grdDPTPLANX.DisplayLayout.Bands(0).Groups("PPD").Hidden = Not tlb_sbt.Checked

            Case "Show Image"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Dim FOLDERNAME As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
                'If ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_SERVER = "" Then FOLDERNAME = "C:\Documents and Settings\Walter\Desktop\JHI\Images\"
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                Dim FILENAME As String = FOLDERNAME & rowICTITEM1.Item("ITEM_PICTURE_FILENAME") & ""

                If My.Computer.FileSystem.FileExists(FILENAME) Then
                    Me.Cursor = Cursors.WaitCursor
                    Call ASCMAIN1.Progress("Now Loading Image Viewer")
                    System.Diagnostics.Process.Start(FILENAME)
                    Me.Cursor = Cursors.Default
                    Call ASCMAIN1.Progress("")
                Else
                    ASCMAIN1.Notify("Cannot Find " & FILENAME, 1)
                    'Call ASCMAIN1.Progress("Cannot Find " & FILENAME)
                End If

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI")

            Case "Demand Planning"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                optBy.Value = "I"
                Absx1.txtFor("ITEM_CODE").Text = ITEM_CODE
                Click_Command("Load")

            Case "Items with Forecasts Only"
                Set_View_grdDPTPLANX()
        End Select
    End Sub

    Public Overrides Sub tlb_ToolValueChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolEventArgs)
        MyBase.tlb_ToolValueChanged(sender, e)

        If e.Tool.OwningMenu Is Nothing Then Exit Sub

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "BackColor"
                Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
                = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
                grd.ActiveRow.Cells("POS_RBG_BACKCOLOR").Value = tlb_cpt.SelectedColor.ToArgb
                grd.UpdateData()
                'Application.DoEvents()
                grdDPTPLANX.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)
                Update_Record_TDA("DPTPOSS1")

            Case "ForeColor"
                Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
                = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
                grd.ActiveRow.Cells("POS_RBG_FORECOLOR").Value = tlb_cpt.SelectedColor.ToArgb
                grd.UpdateData()
                'Application.DoEvents()
                grdDPTPLANX.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)
                Update_Record_TDA("DPTPOSS1")
        End Select

    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BRAND_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Absx1.txtFor("ITEM_CODE").Text = ""
                    Click_Command("Load", e)
                End If
            Case "ITEM_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Absx1.txtFor("BRAND_CODE").Text = ""
                    Click_Command("Load", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "BRAND_CODE"
                Absx1.txtFor("ITEM_CODE").Text = ""
                Click_Command("Load")
            Case "ITEM_CODE"
                Absx1.txtFor("BRAND_CODE").Text = ""
                Click_Command("Load")
            Case "ITEM_CODE_FIND"
                Find_Item()
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "BRAND_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("BRAND_CODE").Text <> "" Then
                        Absx1.txtFor("ITEM_CODE").Text = ""
                        LookUp("ICTBRAN1", Absx1.txtFor("BRAND_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

            Case "ITEM_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("ITEM_CODE").Text <> "" Then
                        Absx1.txtFor("BRAND_CODE").Text = ""
                        LookUp("ICTITEM1", Absx1.txtFor("ITEM_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If
        End Select
    End Sub


    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)

        If SELECTION_NO = 0 Then Exit Sub
        If grdDPTPLANX.ActiveRow Is Nothing Then Exit Sub

        If loading_items Then
            Exit Sub
        End If

        Dim numctl As UltraWinEditors.UltraNumericEditor = DirectCast(sender, UltraWinEditors.UltraNumericEditor)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(numctl)

        Dim ITEM_CODE As String = grdDPTPLANX.ActiveRow.Cells("ITEM_CODE").Text

        Select Case COLUMN_NAME
            Case "ITEM_POS_MAX"
                Calculate_Position(ITEM_CODE)
                grdDPTPLANX.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)

            Case "ITEM_POS_MIN"
                Calculate_Position(ITEM_CODE)
                grdDPTPLANX.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)

            Case "ITEM_LEAD_TIME"
                Calculate_Position(ITEM_CODE)
                grdDPTPLANX.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)

        End Select
    End Sub


#End Region

    Sub Setup_DQ()
        ' IF CHANGING THE TREE, YOU NEED ALSO TO PAY ATTENTION TO ICTITEMX AND ICTITEMY IN FORM LOAD
        COLUMN_NAMEs.Add("HC_CODE")
        COLUMN_NAMEs.Add("COLLECTION_CODE")
        'COLUMN_NAMEs.Add("ITEM_CATGY_CODE")
        'COLUMN_NAMEs.Add("ITEM_CLASS_CODE")

        COLUMN_CAPTIONs.Add("High Coll")
        COLUMN_CAPTIONs.Add("Collection")
        'COLUMN_CAPTIONs.Add("Category")
        'COLUMN_CAPTIONs.Add("Class")


        'COLUMN_NAMEs.Add("DEPT_CODE")
        'COLUMN_NAMEs.Add("ITEM_CATGY_CODE")
        'COLUMN_NAMEs.Add("COLLECTION_CODE")
        'COLUMN_NAMEs.Add("ITEM_CLASS_CODE")
        ''COLUMN_NAMEs.add("STYLE_CODE")

        'COLUMN_CAPTIONs.Add("Department")
        'COLUMN_CAPTIONs.Add("Category")
        'COLUMN_CAPTIONs.Add("Collection")
        'COLUMN_CAPTIONs.Add("Class")
        ''COLUMN_CAPTIONs.add("Style")

        ' Setup DQ Column Sequencing Control

        Dim dt As New DataTable
        For i As Integer = 0 To COLUMN_NAMEs.Count - 1
            dt.Columns.Add(COLUMN_NAMEs(i))
            dt.Columns(i).Caption = COLUMN_CAPTIONs(i)
        Next
        Dim row As DataRow = dt.NewRow
        For i As Integer = 0 To COLUMN_NAMEs.Count - 1
            row.Item(i) = i + 1
        Next
        dt.Rows.Add(row)
        grdDQseq.DataSource = dt
        With grdDQseq.DisplayLayout.Bands(0)
            .CardView = True
            .CardSettings.LabelWidth = 100
            .CardSettings.ShowCaption = False
            .CardSettings.Width = 1
        End With


        'With tvwDQ
        '    Dim rootColumnSet As UltraWinTree.UltraTreeColumnSet = .ColumnSettings.RootColumnSet
        '    rootColumnSet.Columns.Clear()
        '    For i As Integer = 1 To COLUMN_NAMEs.Count
        '        Dim column As UltraWinTree.UltraTreeNodeColumn = rootColumnSet.Columns.Add(COLUMN_NAMEs(i - 1))
        '    Next
        'End With

    End Sub


    Private Sub grdDQseq_AfterColPosChanged(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterColPosChangedEventArgs) Handles grdDQseq.AfterColPosChanged
        Generate_Inquiry()
    End Sub


    Sub Generate_Inquiry()

        If COLUMN_NAMEs.Count = 0 Then
            Exit Sub
        End If

        Application.DoEvents()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Collections into Selection Tree")

        ReDim COLUMN_NAME_by_Lvl(COLUMN_NAMEs.Count)
        ReDim COLUMN_CAPTION_by_Lvl(COLUMN_NAMEs.Count)
        ReDim G_by_Lvl(COLUMN_NAMEs.Count)
        For G As Integer = 1 To COLUMN_NAMEs.Count
            Dim GC As UltraWinGrid.UltraGridColumn = grdDQseq.DisplayLayout.Bands(0).Columns(COLUMN_NAMEs(G - 1))
            Dim Lvl As Integer = GC.Header.VisiblePosition + 1
            COLUMN_NAME_by_Lvl(Lvl) = COLUMN_NAMEs(G - 1)
            COLUMN_CAPTION_by_Lvl(Lvl) = GC.Header.Caption
            G_by_Lvl(Lvl) = G
        Next

        With tvwDQ
            Dim rootColumnSet As UltraWinTree.UltraTreeColumnSet = .ColumnSettings.RootColumnSet
            rootColumnSet.Columns.Clear()
            For Lvl As Integer = 1 To COLUMN_NAMEs.Count
                Dim column As UltraWinTree.UltraTreeNodeColumn = rootColumnSet.Columns.Add(COLUMN_NAME_by_Lvl(Lvl))
            Next
        End With

        Dim Gs() As String = Nothing
        ReDim Gs(COLUMN_NAMEs.Count - 1)
        Dim orderby As String = ""
        For Lvl As Integer = 1 To COLUMN_NAMEs.Count
            Gs(Lvl - 1) = COLUMN_NAMEs(G_by_Lvl(Lvl) - 1) ' "G" & CStr(G_by_Lvl(Lvl))
            orderby &= "," & COLUMN_NAMEs(G_by_Lvl(Lvl) - 1) '",G" & CStr(G_by_Lvl(Lvl))
        Next

        Dim aNode As New Infragistics.Win.UltraWinTree.UltraTreeNode
        Dim CODE_VALUE_at_Lvl() As String = Nothing
        ReDim CODE_VALUE_at_Lvl(COLUMN_NAMEs.Count)

        Dim IMAGE_FOLDER As String = ASCMAIN1.Folders("Images") & "ABS\Menu\Tree\"

        tvwDQ.Nodes.Clear()

        Dim cur_Node_at_Lvl() As Infragistics.Win.UltraWinTree.UltraTreeNode
        ReDim cur_Node_at_Lvl(COLUMN_NAMEs.Count)
        ' do not need an All node
        'If COLUMN_CAPTION_by_Lvl.Length = 1 Then
        '    aNode = tvwDQ.Nodes.Add("*", "All")
        'Else
        '    aNode = tvwDQ.Nodes.Add("*", "All (" & COLUMN_CAPTION_by_Lvl(1) & ")")
        'End If
        'cur_Node_at_Lvl(0) = aNode
        Dim TBL As DataTable = ASCDATA1.SelectDistinct("ICTITEMX", Gs)
        Dim last_level_set As Integer = 0
        'If COLUMN_NAMEs.Count > 1 Then ' no nodes (other than All) when there is only 1 level
        For Each row As DataRow In TBL.Select("", Mid(orderby, 2))
            For Lvl As Integer = 1 To COLUMN_NAMEs.Count ' - 1
                If CODE_VALUE_at_Lvl(Lvl) <> row.Item(Lvl - 1) & "" Or last_level_set < Lvl Then
                    last_level_set = Lvl
                    If Lvl = 1 Then
                        aNode = tvwDQ.Nodes.Add
                    Else
                        aNode = cur_Node_at_Lvl(Lvl - 1).Nodes.Add
                    End If
                    cur_Node_at_Lvl(Lvl) = aNode

                    If Lvl = COLUMN_NAMEs.Count Then
                        Dim KEY As String = ""
                        For Each COLUMN_NAME As String In COLUMN_NAMEs
                            KEY &= "/" & row.Item(COLUMN_NAME)
                        Next
                        aNode.Key = Mid(KEY, 2)
                    End If
                    Dim CAPTION As String = "?" ' Split(row.Item(Lvl - 1) & "", ":")(1)

                    Select Case Gs(Lvl - 1)
                        Case "COLLECTION_CODE"
                            Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(row.Item(Lvl - 1))
                            If rowICTCOLL1 Is Nothing Then
                                CAPTION = "?"
                            Else
                                CAPTION = rowICTCOLL1.Item("COLLECTION_NAME") & ""
                            End If

                        Case "ITEM_CLASS_CODE"
                            Dim rowICTCLAS1 As DataRow = dst.Tables("ICTCLAS1").Rows.Find(row.Item(Lvl - 1))
                            If rowICTCLAS1 Is Nothing Then
                                CAPTION = "?"
                            Else
                                CAPTION = rowICTCLAS1.Item("ITEM_CLASS_DESC") & ""
                            End If

                        Case "ITEM_CATGY_CODE"
                            Dim rowICTCATG1 As DataRow = dst.Tables("ICTCATG1").Rows.Find(row.Item(Lvl - 1))
                            If rowICTCATG1 Is Nothing Then
                                CAPTION = "?"
                            Else
                                CAPTION = rowICTCATG1.Item("ITEM_CATGY_DESC") & ""
                            End If

                    End Select

                    If CAPTION = "?" Then
                        aNode.Text = row.Item(Lvl - 1) & ""
                    Else
                        aNode.Text = row.Item(Lvl - 1) & ":" & CAPTION
                    End If
                    'aNode.Tag = row.Item("MENU_ID") & Chr(1) & KEY
                    aNode.Expanded = False
                    CODE_VALUE_at_Lvl(Lvl) = row.Item(Lvl - 1) & ""
                    If last_level_set = COLUMN_NAMEs.Count Then ' - 1 Then
                        aNode.LeftImages.Add(ASCMAIN1.Get_Image(IMAGE_FOLDER, "ITEM_green")) ' "graph_node"))
                    Else
                        aNode.Override.NodeAppearance.Image = ASCMAIN1.Get_Image(IMAGE_FOLDER, "M")
                        aNode.Override.ExpandedNodeAppearance.Image = ASCMAIN1.Get_Image(IMAGE_FOLDER, "M_OPEN")
                    End If


                    For iLvl As Integer = 1 To Lvl
                        aNode.Cells(iLvl - 1).Value = CODE_VALUE_at_Lvl(iLvl)
                    Next
                End If
            Next
        Next
        'End If

        cur_Node_at_Lvl(1).Expanded = True
        'Click_Node(cur_Node_at_Lvl(1))

        'cur_Node_at_Lvl(0).Expanded = True
        'Click_Node(cur_Node_at_Lvl(0))

        'Sort_grdColumns(grdASTDSQL1, "SORT_VALUE,CODE_VALUE")
        ItemDetails()
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Click_Node(ByVal tnode As UltraWinTree.UltraTreeNode)

        Dim D As String = ""
        Dim sqlw As String = ""

        Dim byTree As Boolean = (tabTG.SelectedTab.Key = "Tree")
        If byTree Then

            If (tnode Is Nothing OrElse tvwDQ.Tag & "" = tnode.Key) Then
                Exit Sub
            End If

            Dim L As Integer = tnode.Level
            If L <> COLUMN_NAMEs.Count - 1 And Not chkAllInNode.Checked Then
                Exit Sub
            End If

            Dim COLUMN_VALUEs As New Dictionary(Of String, String)
            For Each COLUMN_NAME As String In COLUMN_NAMEs
                Dim CODE_VALUE As String = tnode.Cells(COLUMN_NAME).Text
                COLUMN_VALUEs.Add(COLUMN_NAME, CODE_VALUE)
            Next

            For i As Integer = 1 To L + 1
                Dim COLUMN_NAME As String = COLUMN_NAME_by_Lvl(i) ' tnode.Cells(i).Column.Key
                sqlw &= " and " & IIf(COLUMN_VALUEs(COLUMN_NAME) = "", COLUMN_NAME & " is Null", COLUMN_NAME & " = '" & COLUMN_VALUEs(COLUMN_NAME) & "'")
                D &= ", " & COLUMN_CAPTION_by_Lvl(i) & ":" & COLUMN_VALUEs(COLUMN_NAME)
            Next

            D = "Items in" & Mid(D, 2)

        Else
            Dim COLLECTION_CODEs As String = ""
            For Each row As DataRow In dst.Tables("ICTCOLLX").Select("")
                If row.Item("SEL") = "1" Then
                    COLLECTION_CODEs &= ",'" & row.Item("COLLECTION_CODE") & "'"
                End If
            Next
            If COLLECTION_CODEs = "" Then
                COLLECTION_CODEs = ",'?'"
            End If

            If COLLECTION_CODEs <> "" Then sqlw &= " and COLLECTION_CODE IN (" & Mid(COLLECTION_CODEs, 2) & ")"
            If optBP.Value <> "*" Then sqlw &= " and ITEM_BASIC_PROMO = '" & optBP.Value & "'"
            If optSNU.Value <> "*" Then sqlw &= " and ITEM_SNU_CODE = '" & optSNU.Value & "'"
            If txtFindITEM_CODE.Text <> "" Then sqlw &= " and ITEM_CODE = '" & txtFindITEM_CODE.Text & "'"

            If txtFindITEM_CODE.Text <> "" Or byItem Then
                D = "Item Selected"
            Else
                D = IIf(optBP.Value = "*", "", ", " & optBP.Text) _
                    & IIf(optSNU.Value = "*", "", ", " & optSNU.Text) _
                    & IIf(optBP.Value = "*" And optSNU.Value = "*", "", " ") _
                    & "Items in " & Replace(Mid(COLLECTION_CODEs, 2), "'", "")
                If D.StartsWith(", ") Then D = Mid(D, 3)
            End If
        End If

        loading_items = True

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Item Planning Information")

        ASCMAIN1.sql = "Truncate Table " & ICTITEMS
        ASCDATA1.ExecuteSQL()

        If byItem Then
            sqlw &= " and ITEM_CODE = '" & Absx1.txtFor("ITEM_CODE").Text & "'"
        End If

        ASCMAIN1.sql = "Insert into " & ICTITEMS & " (ITEM_CODE) " _
        & " Select ITEM_CODE from " & ICTITEMY _
        & ASCMAIN1.SQL_Add_WHERE(sqlw)

        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
        & "Begin Declare Cursor C1 is " _
        & "  Select ICTSTAT2.ITEM_CODE" _
        & "       , SUM (NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0)) ONH" _
        & "       , SUM (NVL(ICTSTAT2.WHSE_QTY_OPEN,0)) COM" _
        & "       , SUM (NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) - NVL(ICTSTAT2.WHSE_QTY_HOLD,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0)) AVA" _
        & "    from ICTSTAT2,ICTWHSE1," & ICTITEMS & " ICTITEMS" _
        & "   where ICTSTAT2.ITEM_CODE = ICTITEMS.ITEM_CODE" _
        & "     and ICTWHSE1.WHSE_CODE = ICTSTAT2.WHSE_CODE" _
        & "     and NVL(ICTWHSE1.WHSE_MRP_EXC_IND,'0') <> '1'" _
        & "   group by ICTSTAT2.ITEM_CODE;" _
        & " Begin For R1 in C1 Loop" _
        & "  Update " & ICTITEMS & " Set ONH = R1.ONH, COM = R1.COM, AVA = R1.AVA" _
        & "   where ITEM_CODE = R1.ITEM_CODE;" _
        & "End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

        '        & "     and NVL(ICTWHSE1.WHSE_TYPE,'N') = 'R'" & vbCrLf _

        ASCMAIN1.sql = "" _
        & "Begin Declare Cursor C1 is " & vbCrLf _
        & "  Select ICTSTAT2.ITEM_CODE" & vbCrLf _
        & "       , SUM (NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0)) RTN" & vbCrLf _
        & "    from ICTSTAT2,ICTWHSE1," & ICTITEMS & " ICTITEMS" & vbCrLf _
        & "   where ICTSTAT2.ITEM_CODE = ICTITEMS.ITEM_CODE" & vbCrLf _
        & "     and ICTWHSE1.WHSE_CODE = ICTSTAT2.WHSE_CODE" & vbCrLf _
        & "     and NVL(ICTWHSE1.WHSE_MRP_EXC_IND,'0') = '1'" & vbCrLf _
        & "   group by ICTSTAT2.ITEM_CODE;" & vbCrLf _
        & " Begin For R1 in C1 Loop" & vbCrLf _
        & "  Update " & ICTITEMS & " Set RTN = R1.RTN" & vbCrLf _
        & "   where ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
        & "End Loop; End; End;"
        If ASCMAIN1.CLIENT = "INT" Then
            ' LM DOES NOT WANT TO SEE THIS VALUE
        Else
            ASCDATA1.ExecuteSQL()
        End If

        ASCMAIN1.sql = "Update " & ICTITEMS & " X" _
        & " Set MNQ = (Select ITEM_SAFETY_STOCK from ICTITEM1 " _
        & " where ITEM_CODE = X.ITEM_CODE)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
        & "Begin Declare Cursor C1 is " _
        & " SELECT SOTORDR2.ITEM_CODE" _
        & ", SUM (NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) COMMTD" _
        & " FROM SOTORDR2,SOTORDR1," & ICTITEMS & " ICTITEMS" _
        & " WHERE SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" _
        & " AND ICTITEMS.ITEM_CODE = SOTORDR2.ITEM_CODE" _
        & " AND SOTORDR2.ORDR_STATUS >= 'O' AND SOTORDR2.ORDR_STATUS <= 'P'" _
        & " AND ((SOTORDR2.ORDR_STATUS = 'O' AND SOTORDR1.ORDR_SHIP_DATE <= '" & Format(YPFD(0), "dd-MMM-yyyy") & "') OR SOTORDR2.ORDR_STATUS = 'P')" _
        & " GROUP BY SOTORDR2.ITEM_CODE;" _
        & " Begin For R1 in C1 Loop" _
        & "  Update " & ICTITEMS & " Set COMMTD = R1.COMMTD" _
        & "   where ITEM_CODE = R1.ITEM_CODE;" _
        & "End Loop; End; End;"

        'If ASCMAIN1.DBS_COMPANY = "JHI" Then
        '    ASCMAIN1.sql = Replace(ASCMAIN1.sql, _
        '                           "AND SOTORDR2.ORDR_STATUS >= 'O' AND SOTORDR2.ORDR_STATUS <= 'P'", _
        '                           "AND SOTORDR1.ORDR_SHIP_DATE <= '" & Format(YPFD(0), "dd-MMM-yyyy") & "'")
        'End If
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is " _
            & "  Select SOTINVH2.ITEM_CODE" _
            & "       , SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0)) MTD" _
            & "    from SOTINVH2," & ICTITEMS & " ICTITEMS,ICTWHSE1" _
            & "   where SOTINVH2.ITEM_CODE = ICTITEMS.ITEM_CODE" _
            & "     and SOTINVH2.ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "'" _
            & "     and SOTINVH2.INV_TYPE = 'I'" _
            & "   and ICTWHSE1.WHSE_CODE = SOTINVH2.WHSE_CODE" & vbCrLf _
            & "   and NVL(ICTWHSE1.WHSE_MRP_EXC_IND,'0') <> '1'" & vbCrLf _
            & "   group by SOTINVH2.ITEM_CODE;" _
            & " Begin For R1 in C1 Loop" _
            & "  Update " & ICTITEMS & " Set MTD = R1.MTD" _
            & "   where ITEM_CODE = R1.ITEM_CODE;" _
            & "End Loop; End; End;"

        ASCDATA1.ExecuteSQL()

        grdDPTPLANX.Tag = "X"

        ' grdDPTPLANX.Text = "Items in" & Mid(D, 2)
        grdDPTPLANX.Text = D

        EnforceConstraints(False)

        Fill_Records("ICTITEM1")

        Fill_Records("DPTITMF2")

        If chkUseNextSN.Checked Then

            For Each rowICTITEM1 As DataRow In dst.Tables("ICTITEM1").Rows
                rowICTITEM1.Item("ITEM_CATGY_CODE") = "I"
            Next

            For Each rowDPTPROJ0 As DataRow In dst.Tables("DPTPROJ0").Rows
                Dim ITEM_CODE As String = rowDPTPROJ0.Item("ITEM_CODE")
                Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    rowICTITEM1.Item("ITEM_CATGY_CODE") = rowDPTPROJ0.Item("ITEM_CATGY_CODE")
                End If
            Next

            For Each rowDPTABCP0 As DataRow In dst.Tables("DPTABCP0").Rows
                Dim ITEM_CODE As String = rowDPTABCP0.Item("ITEM_CODE")
                Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    rowICTITEM1.Item("ITEM_ABC_CODE") = rowDPTABCP0.Item("ITEM_ABC_CODE")

                    Dim ITEM_CATGY_CODE As String = rowICTITEM1.Item("ITEM_CATGY_CODE") & ""
                    Dim ITEM_ABC_CODE As String = rowICTITEM1.Item("ITEM_ABC_CODE") & ""

                    Dim rowDPTABCP2 As DataRow = dst.Tables("DPTABCP2").Rows.Find(New String() {ITEM_CATGY_CODE, ITEM_ABC_CODE})
                    If rowDPTABCP2 IsNot Nothing Then
                        rowICTITEM1.Item("ITEM_POS_MAX") = rowDPTABCP2.Item("ABC_MAX_POS")
                        rowICTITEM1.Item("ITEM_POS_MIN") = rowDPTABCP2.Item("ABC_MIN_POS")
                        rowICTITEM1.Item("ITEM_MIN_DAYS_SUPPLY") = rowDPTABCP2.Item("ABC_MIN_DAYS_SUPPLY")
                    End If
                End If
            Next

        End If


        dst.Tables("APTVEND1").Rows.Clear()
        dst.Tables("DPTVNDI1").Rows.Clear()

        For Each rowICTITEM1 As DataRow In dst.Tables("ICTITEM1").Select("")
            Dim ITEM_CODE As String = rowICTITEM1.Item("ITEM_CODE") & ""
            Dim VEND_CODE As String = rowICTITEM1.Item("VEND_CODE") & ""
            Fill_Records("APTVEND1", New String() {VEND_CODE}, False)
            Fill_Records("DPTVNDI1", New String() {ITEM_CODE}, False)
        Next

        Fill_Records("ICTSTATH")

        Sales_Forecasts()
        Fill_Records("DPTPLANX")

        Fill_Records("DPTITMFM")
        Fill_Records("POTORDRM")
        Fill_Records("DPTPLANM")
        Fill_Records("POTORDCM")

        Fill_Records("SOTORDRM")
        Fill_Records("SOTINVHM")

        ' put a placeholder record in SOTORDRM for every row in SOTINVHM that is not already represented in SOTORDRM
        For Each row As DataRow In dst.Tables("SOTINVHM").Select("")
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim MARKET_CODE As String = row.Item("MARKET_CODE")
            If dst.Tables("SOTORDRM").Rows.Find(New String() {MARKET_CODE, ITEM_CODE}) Is Nothing Then
                dst.Tables("SOTORDRM").Rows.Add(New Object() {MARKET_CODE, ITEM_CODE, 0})
            End If
        Next

        For Each rowDPTPLANX As DataRow In dst.Tables("DPTPLANX").Rows
            Dim ITEM_CODE As String = rowDPTPLANX.Item("ITEM_CODE")
            For Each TABLE_NAME As String In New String() {"DPTITMFM", "POTORDRM", "DPTPLANM", "POTORDCM"}
                If dst.Tables(TABLE_NAME).Rows.Find(ITEM_CODE) Is Nothing Then
                    dst.Tables(TABLE_NAME).Rows.Add(New String() {ITEM_CODE})
                End If
            Next
        Next

        'Dim ITEM_CODE_BUFFER As String = ""
        'Dim rowICTITEM1_BUFFER As DataRow = Nothing

        'Dim ITEM_BUFFER_QTY As Int32 = 0
        'Dim ITEM_BUFFER_PCT As Decimal = 0
        'For Each rowDPTITMFM As DataRow In dst.Tables("DPTITMFM").Select("", "ITEM_CODE")
        '    If rowDPTITMFM.Item("ITEM_CODE") <> ITEM_CODE_BUFFER Then
        '        ITEM_CODE_BUFFER = rowDPTITMFM.Item("ITEM_CODE")
        '        rowICTITEM1_BUFFER = LookUp("ICTITEM1", ITEM_CODE_BUFFER)
        '        ITEM_BUFFER_QTY = Val(rowICTITEM1_BUFFER.Item("ITEM_BUFFER_QTY") & "")
        '        ITEM_BUFFER_PCT = Val(rowICTITEM1_BUFFER.Item("ITEM_BUFFER_PCT") & "")
        '    End If

        '    'If ASCMAIN1.Running_in_VS AndAlso ITEM_CODE_BUFFER = "CC004P80" Then Stop

        '    If ITEM_BUFFER_QTY > 0 Or ITEM_BUFFER_PCT > 0 Then
        '        For FC As Integer = 0 To 25
        '            Dim FCQTY As Int32 = Val(rowDPTITMFM.Item("FC" & Format(FC, "00")) & "")
        '            If FCQTY > 0 Then
        '                If ITEM_BUFFER_QTY > 0 Then
        '                    FCQTY += ITEM_BUFFER_QTY
        '                ElseIf ITEM_BUFFER_PCT > 0 Then
        '                    FCQTY = System.Math.Round(FCQTY * (100 + ITEM_BUFFER_PCT) / 100, 0)
        '                End If

        '                rowDPTITMFM.Item("FC" & Format(FC, "00")) = FCQTY
        '            End If
        '        Next
        '    End If
        'Next


        If grdSOTINVHX.DataSource IsNot Nothing Then
            Try
                DirectCast(grdSOTINVHX.DataSource, DataTable).Rows.Clear()
            Catch ex As Exception

            End Try
        End If

        Fill_Records("DPTPLAN1")
        Fill_Records("POTORDR9")

        'For Each T As String In New String() {"POTORDRX", "ICTPINVX", "SATAUTHX"}
        '    dst.Tables(T).Rows.Clear()
        'Next

        EnforceConstraints(True)

        dst.Tables("DPTPLAND").Rows.Clear()

        Dim WHSE_CODE_proj = ROWs("DPTPARM1").Item("DP_PARM_DEF_PLAN_WHSE") & ""

        Dim truncate As Boolean = True
        Dim gotDPTPLANX As String = ""

        For Each rowDPTPLANX As DataRow In dst.Tables("DPTPLANX").Rows
            Dim ITEM_CODE As String = rowDPTPLANX.Item("ITEM_CODE")
            ASCMAIN1.Progress("-", ITEM_CODE)
            'Fill_Records("POTORDRX", ITEM_CODE) ' THIS IS BEING DONE IN SETUP ITEM AS WELL - NEED TO RETHINK IT
            ' also being done 5 lines below, so remming it out for now

            dst.Tables("ICTPINVX").Rows.Clear()

            Fill_Records("DPTITMFX", ITEM_CODE, IIf(optBy.Value = "B", False, True))
            Sort_grdColumns(grdDPTITMFX, "MARKET_CODE")

            Fill_Records("POTORDRX", ITEM_CODE, IIf(optBy.Value = "B", False, True))
            Sort_grdColumns(grdPOTORDRX, "PO_ORDER_NO")

            Fill_Records("POTORDR8", {ITEM_CODE, "O"}, IIf(optBy.Value = "B", False, True))
            Sort_grdColumns(grdPOTORDRX, "INIT_DATE".ToLower,, 2)

            Create_Worktables(False, ITEM_CODE, truncate)
            truncate = False
            Fill_Records("ICTPINVX", ITEM_CODE, IIf(optBy.Value = "B", False, True))
            Sort_grdColumns(grdPOTORDRX, "ETA_DATE_DC", , 1)

            Set_DC_Date(ITEM_CODE)

            If gotDPTPLANX = "" Then
                gotDPTPLANX = "1"

                EnforceConstraints(False)
                Dim sqlPOTORDRMx As String = sqlPOTORDRM
                'If optBy.Value = "B" Then
                '    sqlPOTORDRMx = Replace(sqlPOTORDRMx, " group by", $"and ICTITEMS.ITEM_CODE = '{ITEM_CODE}' group by")
                'End If
                Fill_Records("POTORDRM",, IIf(optBy.Value = "B", False, True), sqlPOTORDRMx)

                Dim sqlwx As String = $"ITEM_CODE = '{ITEM_CODE}'"
                sqlwx = ""
                For Each row As DataRow In dst.Tables("DPTPLANX").Select(sqlwx)
                    Dim ITEM_CODE_PLANX As String = row.Item("ITEM_CODE")
                    If dst.Tables("POTORDRM").Rows.Find(ITEM_CODE_PLANX) Is Nothing Then
                        dst.Tables("POTORDRM").Rows.Add(New String() {ITEM_CODE_PLANX})
                    End If
                Next
                EnforceConstraints(True)
            End If

            Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
            '  Calculate_Position(ITEM_CODE)

            Fill_Records("SATAUTHX", rowICTITEM1.Item("COLLECTION_CODE") & "")
            Sort_grdColumns(grdSATAUTHX, "CUST_CODE")

            Dim rowDPTITMFS As DataRow = dst.Tables("DPTITMFS").Rows.Find(New String() {ITEM_CODE, "*"})
            If rowDPTITMFS IsNot Nothing Then
                For P = 1 To 3
                    Dim SHP As Int32 = Val(rowDPTITMFS.Item("S" & Format(P, "00")) & "")
                    Dim DEM As Int32 = Val(rowDPTITMFS.Item("F" & Format(P, "00")) & "")
                    rowDPTPLANX.Item("TREND_SHP_" & Format(P, "0")) = SHP
                    rowDPTPLANX.Item("TREND_DEM_" & Format(P, "0")) = DEM
                    Dim PCT As Decimal = 0
                    If DEM <> 0 Then PCT = 100 * (SHP - DEM) / DEM
                    rowDPTPLANX.Item("TREND_PCT_" & Format(P, "0")) = PCT
                Next
            End If

            Dim ITEM_FC_DAY_REQ As Int32 = Val(rowICTITEM1.Item("ITEM_FC_DAY_REQ") & "")
            If ITEM_FC_DAY_REQ = 0 Then
                ITEM_FC_DAY_REQ = Val(ROWs("DPTPARM1").Item("DP_PARM_FC_REQ_DAY") & "")
            End If
            If ITEM_FC_DAY_REQ = 0 Then ITEM_FC_DAY_REQ = 1

            Dim ONH As Int32 = Val(rowDPTPLANX.Item("ONH") & "")
            dst.Tables("DPTPLAND").Rows.Add(New Object() {ITEM_CODE, CDate("01/01/0001"), ONH, 0, "Current On Hand", WHSE_CODE_proj})
            Dim MTD As Int32 = Val(rowDPTPLANX.Item("MTD") & "")
            dst.Tables("DPTPLAND").Rows.Add(New Object() {ITEM_CODE, CDate("01/01/0001"), MTD, 0, "MTD Shipments", WHSE_CODE_proj})
            Dim COM As Int32 = Val(rowDPTPLANX.Item("COM") & "")
            Dim COMMTD As Int32 = Val(rowDPTPLANX.Item("COMMTD") & "")
            Dim DEM00 As Int32 = Val(rowDPTPLANX.Item("DEM00") & "")

            For P = -1 To 25
                Dim DEM As Int64 = Val(rowDPTPLANX.Item("FC" & IIf(P = -1, "PD", Format(P, "00"))) & "")
                Dim DEMDTX As String = ""
                Dim DEMDTX_LBL As String = ""
                If P = -1 Then
                    DEMDTX = YMF(0) & Format(ITEM_FC_DAY_REQ, "00")
                    DEMDTX_LBL = YPF(0, 1) & " PD"
                Else
                    DEMDTX = YMF(P) & Format(ITEM_FC_DAY_REQ, "00")
                    DEMDTX_LBL = YPF(P, 1)
                End If
                Dim DEMDT As Date = CDate(Mid(DEMDTX, 5, 2) & "/" & Mid(DEMDTX, 7, 2) & "/" & Mid(DEMDTX, 1, 4))
                If DEM <> 0 Then
                    dst.Tables("DPTPLAND").Rows.Add(New Object() {ITEM_CODE, DEMDT, -1 * DEM, 0, "Forecast " & DEMDTX_LBL, WHSE_CODE_proj})
                End If
            Next


            Dim SNU As String = rowICTITEM1.Item("ITEM_SNU_CODE") & ""
            Dim BP As String = rowICTITEM1.Item("ITEM_BASIC_PROMO") & ""

            If SNU = "" Then SNU = "S"
            If BP = "" Then BP = "B"


            For Each rowSOTORDRM As DataRow In dst.Tables("SOTORDRM").Select("ITEM_CODE = '" & ITEM_CODE & "'", "MARKET_CODE")
                Dim MARKET_CODE As String = rowSOTORDRM.Item("MARKET_CODE")

                Dim rowSOTMKTC1 As DataRow = dst.Tables("SOTMKTC1").Rows.Find(MARKET_CODE)

                Dim PAST_DUE_FC As String = "" ' THIS SECTION MUST BE CONSISTENT IN MRP UPDATE AND DP INQUIRY
                If rowSOTMKTC1 IsNot Nothing Then
                    PAST_DUE_FC = rowSOTMKTC1.Item("PAST_DUE_FC_" & SNU & BP) & ""
                    If PAST_DUE_FC = "" Then
                        PAST_DUE_FC = "0"
                    End If
                Else
                    'PAST_DUE_FC = "A"
                    PAST_DUE_FC = "0" ' PER NF 06/27/25
                End If

                If "0P".Contains(PAST_DUE_FC) Then

                    For P As Integer = 0 To 25
                        Dim DEMDTX As String = YMF(P) & Format(ITEM_FC_DAY_REQ, "00")
                        Dim DEMDTX_LBL As String = YPF(P, 1)
                        Dim DEMDT As Date = CDate(Mid(DEMDTX, 5, 2) & "/" & Mid(DEMDTX, 7, 2) & "/" & Mid(DEMDTX, 1, 4))
                        Dim SOXX As Int32 = Val(rowSOTORDRM.Item("SO" & Format(P, "00")) & "")

                        Dim rowDPTITMFX As DataRow = dst.Tables("DPTITMFX").Rows.Find(New String() {ITEM_CODE, MARKET_CODE})

                        Dim FCXX As Int32 = 0
                        If rowDPTITMFX IsNot Nothing Then
                            FCXX = Val(rowDPTITMFX.Item("FC" & Format(P, "00")))
                        End If

                        Dim DEMADJ As Int32 = FCXX - SOXX
                        If P = 0 Then
                            Dim rowSOTINVHM As DataRow = dst.Tables("SOTINVHM").Rows.Find(New String() {MARKET_CODE, ITEM_CODE})
                            Dim MTDSHP As Int32 = 0
                            If rowSOTINVHM IsNot Nothing Then MTDSHP = Val(rowSOTINVHM.Item("QTYI") & "")
                            Dim SOPD As Int32 = Val(rowSOTORDRM.Item("SOPD"))
                            Dim FCPD As Int32 = 0
                            If rowDPTITMFX IsNot Nothing Then FCPD = Val(rowDPTITMFX.Item("FCPD"))
                            DEMADJ = FCPD + FCXX - (MTDSHP + SOPD + SOXX)
                        End If

                        If DEMADJ < 0 Then
                            If P = 0 Then
                                dst.Tables("DPTPLAND").Rows.Add(New Object() {ITEM_CODE, YPPD(0), DEMADJ, 0, "MTD+Open+Pick>FC " & MARKET_CODE, WHSE_CODE_proj})
                            Else
                                dst.Tables("DPTPLAND").Rows.Add(New Object() {ITEM_CODE, DEMDT, DEMADJ, 0, "Open>FC " & MARKET_CODE, WHSE_CODE_proj})
                            End If
                            rowDPTPLANX.Item("DEMADJ" & Format(P, "00")) = Val(rowDPTPLANX.Item("DEMADJ" & Format(P, "00")) & "") - 1 * DEMADJ
                        End If
                    Next
                End If
            Next


            For Each rowPOTORDRX As DataRow In dst.Tables("POTORDRX").Select("ITEM_CODE = '" & ITEM_CODE & "'")
                For Each rowICTPINVX As DataRow In rowPOTORDRX.GetChildRows("POTORDRX_ICTPINVX")
                    Dim SUP As Int64 = Val(rowICTPINVX.Item("PINV_QTY") & "") 'Note: PINV_QTY = INV_QTY + OPO_QTY
                    'Dim SUPDT As Date = rowICTPINVX.Item("ETA_DATE")
                    Dim SUPDT As Date = rowICTPINVX.Item("ETA_DATE_DC")
                    Dim WHSE_CODE_rel As String = rowICTPINVX.Item("WHSE_CODE") & ""
                    dst.Tables("DPTPLAND").Rows.Add(New Object() {ITEM_CODE, SUPDT, SUP, 0, "PO " & rowICTPINVX.Item("PO_ORDER_NO") & ", " & rowICTPINVX.Item("CONTAINER_NO"), WHSE_CODE_rel})
                Next
                'Dim SUP As Int64 = Val(rowPOTORDRX.Item("PO_QTY_OPN") & "")
                'Dim SUPDT As Date = rowPOTORDRX.Item("PO_DATE_REQUIRED")
                'Dim WHSE_CODE_rel As String = rowPOTORDRX.Item("WHSE_CODE") & ""
                'dst.Tables("DPTPLAND").Rows.Add(New Object() {ITEM_CODE, SUPDT, SUP, 0, "PO " & rowPOTORDRX.Item("PO_ORDER_NO"), WHSE_CODE_rel})
            Next

            For Each rowDPTPLAN1 As DataRow In dst.Tables("DPTPLAN1").Select
                Dim SUP As Int64 = Val(rowDPTPLAN1.Item("QTY_PLANNED") & "")
                Dim SUPDT As Date = rowDPTPLAN1.Item("DATE_REQUIRED")
                Dim PLAN_NO As String = rowDPTPLAN1.Item("PLAN_NO")
                Dim WHSE_CODE_rel As String = rowDPTPLAN1.Item("TO_WHSE") & ""
                dst.Tables("DPTPLAND").Rows.Add(New Object() {ITEM_CODE, SUPDT, SUP, 0, "Plan " & PLAN_NO, WHSE_CODE_rel})
            Next

            For Each rowPOTORDR9 As DataRow In dst.Tables("POTORDR9").Select("ITEM_CODE = '" & ITEM_CODE & "'")
                Dim DEM As Int64 = Val(rowPOTORDR9.Item("PO_QTY_COM") & "")
                Dim DEMDT As Date = rowPOTORDR9.Item("DATE_REQUIRED")
                Dim PO_ORDER_NO As String = rowPOTORDR9.Item("PO_ORDER_NO")
                Dim PO_ORDER_LNO As Integer = rowPOTORDR9.Item("PO_ORDER_LNO")
                Dim WHSE_CODE_rel As String = rowPOTORDR9.Item("AT_WHSE") & ""
                dst.Tables("DPTPLAND").Rows.Add(New Object() {ITEM_CODE, DEMDT, -1 * DEM, 0, IIf(PO_ORDER_LNO = 0, "PCom Plan " & PO_ORDER_NO, "PCom PO " & PO_ORDER_NO), WHSE_CODE_rel})
            Next

            Dim BALANCE As Int32 = 0
            For Each rowDPTPLAND As DataRow In dst.Tables("DPTPLAND").Select("ITEM_CODE = '" & ITEM_CODE & "'", "DATE_BALANCE")
                BALANCE += Val(rowDPTPLAND.Item("ACTIVITY") & "")
                rowDPTPLAND.Item("BALANCE") = BALANCE
            Next

            Sort_grdColumns(grdDPTPLAND, "DATE_BALANCE")

            Dim rowICTSTATH As DataRow = dst.Tables("ICTSTATH").Rows.Find(New String() {ITEM_CODE})
            If rowICTSTATH IsNot Nothing Then
                For P = 1 To 3
                    Dim ONHXX As Int32 = Val(rowICTSTATH.Item("OH" & Format(P, "0")) & "")
                    rowDPTPLANX.Item("TREND_ONH_" & Format(P, "0")) = ONHXX
                Next
            End If

            Calculate_Position(ITEM_CODE)

        Next

        grdDPTPLANX.Tag = ""
        Sort_grdColumns(grdDPTPLANX, "ITEM_CODE")
        If grdDPTPLANX.Rows.Count <> 0 Then
            grdDPTPLANX.ActiveRow = grdDPTPLANX.Rows(0)
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        If byTree Then tvwDQ.Tag = tnode.Key

        loading_items = False

    End Sub

    Private Sub tvwDQ_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tvwDQ.Click

        Try
            Dim xx As System.Windows.Forms.MouseEventArgs = DirectCast(e, System.Windows.Forms.MouseEventArgs)
            Dim tt As UltraWinTree.UltraTree = DirectCast(sender, UltraWinTree.UltraTree)
            Dim tnode As UltraWinTree.UltraTreeNode = tt.GetNodeFromPoint(xx.X, xx.Y)

            If tnode IsNot Nothing Then
                Click_Node(tvwDQ.ActiveNode)
            End If


        Catch ex As Exception

        End Try

    End Sub

    Private Sub grdDPTPLANX_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdDPTPLANX.AfterRowActivate
        If modes_Edit Then
            Mode_Settings_Edit(False)
        End If
        If modes_Edit_Plans Then
            Mode_Settings_Edit_Plans(False)
        End If
        Setup_Item()
    End Sub

    Private Sub grdDPTPLANX_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdDPTPLANX.InitializeRow

        If grdDPTPLANX.Tag = "X" Then Exit Sub

        ' e.Row.Cells("DATA_SUP").Value = "POs"
        e.Row.Cells("DATA_PO").Value = "POs"
        e.Row.Cells("DATA_DEM").Value = "FC"
        e.Row.Cells("DATA_ONH").Value = "EOM"
        e.Row.Cells("DATA_POS").Value = "Pos"
        e.Row.Cells("DATA_OS").Value = "+/-"
        e.Row.Cells("DATA_GP").Value = "Plans"
        e.Row.Cells("DATA_SO").Value = "S/O"
        e.Row.Cells("DATA_SI").Value = "S-In"
        e.Row.Cells("DATA_ST").Value = "S-Th"

        e.Row.Cells("STALBL_MTD").Value = "Mtd Sls"
        e.Row.Cells("STALBL_ONH").Value = "On Hand"
        e.Row.Cells("STALBL_COM").Value = "Opn Ord"
        e.Row.Cells("STALBL_AVA").Value = "Ava"
        e.Row.Cells("STALBL_RTN").Value = IIf(ASCMAIN1.CLIENT = "INT", "", "ExclOH")
        e.Row.Cells("STALBL_MNQ").Value = "MinQty"

        e.Row.Cells("TREND_SHP_L").Value = "Actual"
        e.Row.Cells("TREND_DEM_L").Value = "FC"
        e.Row.Cells("TREND_ONH_L").Value = "EOM"
        e.Row.Cells("TREND_PCT_L").Value = "+/- %"

        e.Row.Appearance.BackColor = Color.Beige

        Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Value & ""
        Dim rowDPTPLANO As DataRow = dst.Tables("DPTPLANO").Rows.Find(ITEM_CODE)

        Dim GP As Int32 = 0
        Dim OS As Int32 = 0


        If chkShowOverShort.Checked Then
            'sqlDPTPLANS(S)
            Dim rowDPTPLANS_SI As DataRow = dst.Tables("DPTPLANS").Rows.Find(New String() {ITEM_CODE, "I"})
            If rowDPTPLANS_SI Is Nothing Then
                Fill_Records("DPTPLANS", ITEM_CODE, False)
                Fill_Records("DPTPLANS", , False, Replace(sqlDPTPLANS(1), ":PARM1", "'" & ITEM_CODE & "'"))
                rowDPTPLANS_SI = dst.Tables("DPTPLANS").Rows.Find(New String() {ITEM_CODE, "I"})
            End If
            Dim rowDPTPLANS_ST As DataRow = dst.Tables("DPTPLANS").Rows.Find(New String() {ITEM_CODE, "T"})
            For P As Integer = 0 To 12
                If rowDPTPLANS_SI IsNot Nothing Then
                    e.Row.Cells("SI" & Format(P, "00")).Value = rowDPTPLANS_SI.Item("S" & Format(P, "00"))
                End If
                If rowDPTPLANS_ST IsNot Nothing Then
                    e.Row.Cells("ST" & Format(P, "00")).Value = rowDPTPLANS_ST.Item("S" & Format(P, "00"))
                End If
            Next
        End If


        Dim POS_STATUS_CODES As String = e.Row.Cells("POS_STATUS_CODES").Text
        Dim GP_TOTAL As Int64 = 0
        For P As Integer = -1 To 24
            Dim PX As String = IIf(P = -1, "PD", Format(P, "00"))
            If POS_STATUS_CODES.Length >= (P + 2) Then
                Dim POS_STATUS_CODE As String = Mid(POS_STATUS_CODES, P + 2, 1)
                Dim rowDPTPOSS1 As DataRow = dst.Tables("DPTPOSS1").Rows.Find(POS_STATUS_CODE)

                Dim C As String = "POS" & PX
                If Val(rowDPTPOSS1.Item("POS_RBG_FORECOLOR") & "") = 0 Then
                    e.Row.Cells(C).Appearance.ForeColor = Color.Empty
                Else
                    e.Row.Cells(C).Appearance.ForeColor = Color.FromArgb(Val(rowDPTPOSS1.Item("POS_RBG_FORECOLOR") & ""))
                End If
                If Val(rowDPTPOSS1.Item("POS_RBG_BACKCOLOR") & "") = 0 Then
                    e.Row.Cells(C).Appearance.BackColor = Color.Empty
                Else
                    e.Row.Cells(C).Appearance.BackColor = Color.FromArgb(Val(rowDPTPOSS1.Item("POS_RBG_BACKCOLOR") & ""))
                End If
            End If

            e.Row.Cells("GP" & PX).Value = e.Row.Cells("PL" & PX).Value
            GP_TOTAL += Val(e.Row.Cells("GP" & PX).Value & "")
            If rowDPTPLANO IsNot Nothing Then
                If (e.Row.Cells("GP" & PX).Value & "") <> "" Then
                    OS = Val(rowDPTPLANO.Item("OS" & PX) & "")
                    If OS <> 0 Then
                        e.Row.Cells("OS" & PX).Value = OS
                        If OS < 0 Then
                            e.Row.Cells("OS" & PX).Appearance.ForeColor = Color.Red
                        Else
                            e.Row.Cells("OS" & PX).Appearance.ForeColor = Color.Green
                        End If
                    End If
                End If

                'GP = Val(rowDPTPLANO.Item("GP" & PX) & "")
                ''e.Row.Cells("GP" & PX).Appearance.BackColor = Color.LightCyan
                ''e.Row.Cells("GP" & PX).Appearance.BorderColor = Color.Crimson
                'If GP <> 0 Then
                '    e.Row.Cells("GP" & PX).Value = GP
                '    'If GP < 0 Then
                '    '    e.Row.Cells("GP" & PX).Appearance.ForeColor = Color.White
                '    '    e.Row.Cells("GP" & PX).Appearance.BackColor = Color.OrangeRed
                '    'Else
                '    '    e.Row.Cells("GP" & PX).Appearance.ForeColor = Color.White
                '    '    e.Row.Cells("GP" & PX).Appearance.BackColor = Color.DodgerBlue
                '    'End If
                'End If
            End If
        Next

        If ASCMAIN1.Running_in_VS Then
            If rowDPTPLANO Is Nothing Then
                ' Stop
                Exit Sub
            End If
        End If



        GP = GP_TOTAL ' Val(rowDPTPLANO.Item("GP") & "")
        ' OS = Val(rowDPTPLANO.Item("OS") & "") ' loop above will leave OS with the last non-null value

        e.Row.Cells("TOTAL_OS").Value = OS
        e.Row.Cells("TOTAL_GP").Value = GP

        Dim PLT As Integer
        Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
        Dim ITEM_PLAN_MAKE_BUY As String = rowICTITEM1.Item("ITEM_PLAN_MAKE_BUY") & ""
        Dim VEND_CODE As String = rowICTITEM1.Item("VEND_CODE") & ""

        Dim rowDPTVNDI1 As DataRow = dst.Tables("DPTVNDI1").Rows.Find(New String() {VEND_CODE, ITEM_CODE})
        Dim rowAPTVEND1 As DataRow = dst.Tables("APTVEND1").Rows.Find(VEND_CODE)
        If rowAPTVEND1 Is Nothing Then rowAPTVEND1 = dst.Tables("APTVEND1").NewRow



        Dim PO_LEAD_TIME As Int64 = 0
        If rowDPTVNDI1 IsNot Nothing AndAlso Val(rowDPTVNDI1.Item("PO_LEAD_TIME") & "") <> 0 Then
            PO_LEAD_TIME = Val(rowDPTVNDI1.Item("PO_LEAD_TIME") & "")
        Else
            If Val(rowICTITEM1.Item("ITEM_LEAD_TIME_DAYS") & "") <> 0 Then
                ' INTRODUCED ICTITEM1.ITEM_LEAD_TIME_DAYS AND REVERSED ORDER ITEM VS VENDOR
                PO_LEAD_TIME = Val(rowICTITEM1.Item("ITEM_LEAD_TIME_DAYS") & "")
            Else
                PO_LEAD_TIME = Val(rowAPTVEND1.Item("PO_LEAD_TIME") & "")
            End If
            'If Val(rowAPTVEND1.Item("PO_LEAD_TIME") & "") <> 0 Then
            '    PO_LEAD_TIME = Val(rowAPTVEND1.Item("PO_LEAD_TIME") & "")
            'Else
            '    'PO_LEAD_TIME = 0 ' Val(rowICTITEM1.Item("ITEM_DAYS_PO_LT") & "")
            'End If
        End If

        Dim PO_SCH_DAYS As Int64 = Get_PO_SCH_DAYS(ITEM_PLAN_MAKE_BUY, rowDPTVNDI1, rowAPTVEND1)

        Dim PO_XIT_DAYS As Int64 = 0
        If rowDPTVNDI1 IsNot Nothing AndAlso Val(rowDPTVNDI1.Item("PO_XIT_DAYS") & "") <> 0 Then
            PO_XIT_DAYS = Val(rowDPTVNDI1.Item("PO_XIT_DAYS") & "")
        Else
            If Val(rowAPTVEND1.Item("PO_XIT_DAYS") & "") <> 0 Then
                PO_XIT_DAYS = Val(rowAPTVEND1.Item("PO_XIT_DAYS") & "")
            Else
                PO_XIT_DAYS = Val(ROWs("DPTPARM1").Item("DP_PARM_DEF_PO_XIT_DAYS") & "")
            End If
        End If

        Dim ITEM_LEAD_TIME_DAYS As Integer = Val(rowICTITEM1.Item("ITEM_LEAD_TIME_DAYS") & "")

        Dim CRITICAL_LEAD_TIME_days As Integer = PO_LEAD_TIME + PO_SCH_DAYS + PO_XIT_DAYS ' + mfg
        If ASCMAIN1.DBS_COMPANY = "INT" Then
            CRITICAL_LEAD_TIME_days = ITEM_LEAD_TIME_DAYS
        End If

        Absx1.numFor("PO_LEAD_TIME").Value = PO_LEAD_TIME
        Absx1.numFor("PO_SCH_DAYS").Value = PO_SCH_DAYS
        Absx1.numFor("PO_XIT_DAYS").Value = PO_XIT_DAYS
        Absx1.numFor("CRITICAL_LEAD_TIME").Value = CRITICAL_LEAD_TIME_days
        Dim LTD As Date = Now.AddDays(CRITICAL_LEAD_TIME_days)

        lblCRITICAL_LEAD_TIME_date.Text = Format(LTD, "MM/dd/yyyy")
        For P As Integer = 0 To 25
            PLT = P
            'e.Row.Cells("SUP" & Format(P, "00")).Appearance.ForeColor = Color.Red
            e.Row.Cells("PO" & Format(P, "00")).Appearance.ForeColor = Color.Red
            If Format(YPFD(P), "yyyyMMdd") > Format(LTD, "yyyyMMdd") Then
                If P < 25 Then
                    For P2 As Integer = P + 1 To 25
                        e.Row.Cells("PO" & Format(P2, "00")).Appearance.ForeColor = Color.Empty
                        e.Row.Cells("PO" & Format(P2, "00")).Appearance.BackGradientStyle = GradientStyle.None
                        e.Row.Cells("PO" & Format(P2, "00")).Appearance.BackColor = Color.Empty
                        e.Row.Cells("PO" & Format(P2, "00")).Appearance.BackColor2 = Color.Empty
                    Next
                End If
                Exit For
            End If
        Next
        e.Row.Cells("PO" & Format(PLT, "00")).Appearance.BackColor = Color.MistyRose
        e.Row.Cells("PO" & Format(PLT, "00")).Appearance.BackColor2 = Color.White
        e.Row.Cells("PO" & Format(PLT, "00")).Appearance.BackGradientStyle = GradientStyle.GlassRight20
        For P = 1 To 3
            Dim COLUMN_NAME As String = "TREND_PCT_" & Format(P, "0")
            Dim PCT As Decimal = Val(e.Row.Cells(COLUMN_NAME).Value & "")
            If PCT > 0 Then
                If PCT > 10 Then
                    e.Row.Cells(COLUMN_NAME).Appearance.BackColor = Color.Green
                    e.Row.Cells(COLUMN_NAME).Appearance.ForeColor = Color.White
                Else
                    e.Row.Cells(COLUMN_NAME).Appearance.BackColor = Color.LightGreen
                End If
            ElseIf PCT < 0 Then
                If PCT < -10 Then
                    e.Row.Cells(COLUMN_NAME).Appearance.BackColor = Color.Red
                    'e.Row.Cells(COLUMN_NAME).Appearance.ForeColor = Color.White
                Else
                    e.Row.Cells(COLUMN_NAME).Appearance.BackColor = Color.MistyRose
                End If
            End If
        Next

        'e.Row.Appearance.BackColor = Color.Azure

        Dim MNQ As Integer = Val(e.Row.Cells("MNQ").Value & "")
        Dim ONH As Integer = Val(e.Row.Cells("ONH").Value & "")
        If ONH < MNQ Then
            e.Row.Cells("ONH").Appearance.ForeColor = Color.Red
        End If

        If Val(e.Row.Cells("MTD").Value & "") + Val(e.Row.Cells("SO00").Value & "") > Val(e.Row.Cells("DEM00").Value & "") Then
            'If Val(e.Row.Cells("MTD").Value & "") + Val(e.Row.Cells("COMMTD").Value & "") > Val(e.Row.Cells("DEM00").Value & "") Then
            ' USE TOTAL_SO OR SO00 INSTEAD OF COMMTD

            e.Row.Cells("DEM00").Appearance.BackColor = Color.LimeGreen
            e.Row.Cells("DEM00").ToolTipText = "Total MTD Shipped (" & e.Row.Cells("MTD").Value & ") + Open Curr Mo (" & e.Row.Cells("SO00").Value & ") is > Forecasted Demand (" & e.Row.Cells("DEM00").Value & ")"
        Else
            e.Row.Cells("DEM00").Appearance.BackColor = Color.Empty
        End If

        For i As Integer = 1 To 3 ' need to know how many months out we need to look
            Dim SO As String = "SO" & Format(i, "00")
            Dim DEM As String = "DEM" & Format(i, "00")
            If Val(e.Row.Cells(SO).Value & "") > Val(e.Row.Cells(DEM).Value & "") Then
                e.Row.Cells(DEM).Appearance.BackColor = Color.LimeGreen
                e.Row.Cells(DEM).ToolTipText = "Open SO (" & e.Row.Cells(SO).Value & ") is > Forecasted Demand (" & e.Row.Cells(DEM).Value & ")"
            Else
                e.Row.Cells(DEM).Appearance.BackColor = Color.Empty
            End If
        Next
    End Sub


    Function Get_PO_SCH_DAYS(ITEM_PLAN_MAKE_BUY As String, rowDPTVNDI1 As DataRow, rowAPTVEND1 As DataRow) As Int64
        Dim PO_SCH_DAYS As Int64 = 0

        If ITEM_PLAN_MAKE_BUY = "M" Then
            If rowDPTVNDI1 Is Nothing Then
                PO_SCH_DAYS = 0
            Else
                If Val(rowDPTVNDI1.Item("PO_SCH_DAYS") & "") <> 0 Then
                    PO_SCH_DAYS = Val(rowDPTVNDI1.Item("PO_SCH_DAYS") & "")
                Else
                    If Val(rowAPTVEND1.Item("PO_SCH_DAYS") & "") <> 0 Then
                        PO_SCH_DAYS = Val(rowAPTVEND1.Item("PO_SCH_DAYS") & "")
                    Else
                        PO_SCH_DAYS = Val(ROWs("DPTPARM1").Item("DP_PARM_DEF_PO_SCH_DAYS") & "")
                    End If
                End If
            End If
        End If

        Return PO_SCH_DAYS
    End Function

    Private Sub grdDPTPLANX_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdDPTPLANX.InitializeLayout

    End Sub

    Sub Set_Item_Master_Column(
    ByVal COLUMN_NAME As String,
    ByVal G As UltraWinGrid.UltraGridGroup,
    ByVal Level As Integer,
    Optional ByVal ColSpan As Integer = 0)

        With grdDPTPLANX.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
            .Group = G
            .Level = Level
            .CellActivation = UltraWinGrid.Activation.NoEdit
            If COLUMN_NAME = "ITEM_CODE" Then
                .CellAppearance.BackColor = Color.WhiteSmoke
            Else
                .CellAppearance.BackColor = Color.White
            End If
            If ColSpan <> 0 Then
                .ColSpan = ColSpan
            End If
            .Width = 80 * ColSpan

        End With
    End Sub

    Sub ItemDetails()
        splDPTPLANX.Panel2Collapsed = Not (chkItemDetails.Checked)
        Setup_tabDetails()
    End Sub


    Private Sub optGC_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optGC.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        'splChart.Visible = (optGC.Value = "C")
        'splGrid.Visible = (optGC.Value = "G")
        'splChart.Visible = True
        'splGrid.Visible = True
        'If optGC.Value = "C" Then
        '    chtTotals.Visible = True
        '    chtTrend.Visible = True
        'End If
        If optGC.Value = "C" Then
            SplitContainer1.Panel1Collapsed = True
            SplitContainer1.Panel2Collapsed = False
        ElseIf optGC.Value = "G" Then
            SplitContainer1.Panel1Collapsed = False
            SplitContainer1.Panel2Collapsed = True
        ElseIf optGC.Value = "B" Then
            SplitContainer1.Panel1Collapsed = False
            SplitContainer1.Panel2Collapsed = False
        End If
    End Sub

    Private Sub optIT_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optIT.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If grdGLTPARM2.Selected.Rows.Count > 0 Then
            cmdFetchSales.PerformClick()
        End If

    End Sub

    Private Sub optMW_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optMW.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If optMW.Value = "M" Then
            grdGLTPARM2.DataSource = dst.Tables("GLTPARM2")
        Else
            grdGLTPARM2.DataSource = dst.Tables("GLTPARM3")
        End If
        Sort_grdColumns(grdGLTPARM2, "OPS_YYYYPP".ToLower)
        grdSOTINVHX.Visible = False
        optUD.Visible = False
        optCM.Visible = False
        optCM.Value = "C"
        optGC.Visible = False
        chtTotals.Visible = False
        chtTrend.Visible = False


        optGC.Value = "G"
    End Sub

    Private Sub cmdFetchSales_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFetchSales.Click

        If grdDPTPLANX.ActiveRow Is Nothing Then
            MsgBox("You must first select an Item", MsgBoxStyle.OkOnly, "Cannot Fetch Sales for No Item")
            Exit Sub
        End If

        Dim PRDS As Integer = grdGLTPARM2.Selected.Rows.Count
        If PRDS = 0 Then
            MsgBox("You Must First Select a Time Frame", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        If optIT.Value = "I" And optMW.Value = "W" Then
            MsgBox("Weekly Time Frames not supported for Sell-In", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        ASCMAIN1.Progress("Now Fetching Sales")

        grdSOTINVHX.DisplayLayout.Bands(0).Summaries.Clear()

        grdSOTINVHX.DataSource = Nothing

        Dim YP1 As String = "999999" ' grdGLTPARM2.Selected.Rows(0).Cells("OPS_YYYYPP").Text
        Dim YP2 As String = "000000" '  grdGLTPARM2.Selected.Rows(PRDS - 1).Cells("OPS_YYYYPP").Text

        Dim ITEM_CODE As String = grdDPTPLANX.ActiveRow.Cells("ITEM_CODE").Text

        Dim YPC As String = ""
        If optMW.Value = "M" Then
            If optIT.Value = "I" Then
                YPC = "ORDR_YYYYPP_UPDATED"
            Else
                YPC = "OPS_YYYYPP"
            End If

        Else
            YPC = "OPS_YYYYWW"
        End If

        Dim CS As String = ""
        Dim YPs() As String
        Dim LEGENDs() As String

        ReDim YPs(PRDS)
        ReDim LEGENDs(PRDS)
        If optCM.Value = "C" Then
            ASCMAIN1.sql = "Select CUST_CODE CODE_VALUE"
        ElseIf optCM.Value = "M" Then
            ASCMAIN1.sql = "Select SOTTCLS1.MARKET_cODE CODE_VALUE"

        End If
        For P As Integer = 1 To PRDS ' PRDS To 1 Step -1

            Dim YP As String = grdGLTPARM2.Selected.Rows(P - 1).Cells("OPS_YYYYPP").Text
            Dim LEGEND As String = ""
            If optMW.Value = "M" Then
                LEGEND = ASCMAIN1.Get_Legend(YP, False, True)
            Else
                LEGEND = ASCMAIN1.Get_Legend_Wk(YP, True)
            End If
            YPs(P) = YP
            If YP < YP1 Then YP1 = YP
            If YP > YP2 Then YP2 = YP
            LEGENDs(P) = LEGEND
            Dim DATA_COLUMN As String = ""
            If optIT.Value = "I" Then
                If optUD.Value = "U" Then
                    DATA_COLUMN = "ORDR_QTY_SHIP"
                Else
                    DATA_COLUMN = "ORDR_AMT_SHIP"
                End If
            Else
                If optUD.Value = "U" Then
                    DATA_COLUMN = "QTY_SOLD"
                Else
                    DATA_COLUMN = "AMT_SOLD"
                End If
            End If
            CS &= "+ISNULL(S" & Format(P, "00") & ",0)"
            ASCMAIN1.sql &= ", SUM (DECODE(" & YPC & ",'" & YP & "'," & DATA_COLUMN & ",0)) S" & Format(P, "00")
        Next

        Dim TABLE_NAME As String = ""
        If optIT.Value = "I" Then
            TABLE_NAME = "SOTINVH2" ' "SATSSUMI"
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "ORDR_AMT_SHIP", "ORDR_QTY_SHIP * ORDR_UNIT_PRICE")
        Else
            TABLE_NAME = "RSTRETL1"
        End If

        '   Dim GROUPBY As String = ""
        If optCM.Value = "C" Then
            ASCMAIN1.sql &= " from " & TABLE_NAME _
            & " where " & YPC & " between '" & YP1 & "' and '" & YP2 & "'" _
            & " and ITEM_CODE = '" & ITEM_CODE & "'" _
            & " group by CUST_CODE"
        ElseIf optCM.Value = "M" Then
            ASCMAIN1.sql &= " from " & TABLE_NAME & ", ARTCUST1, SOTTCLS1" _
            & " where " & YPC & " between '" & YP1 & "' and '" & YP2 & "'" _
            & " and ITEM_CODE = '" & ITEM_CODE & "'" _
            & " and  ARTCUST1.CUST_CODE = " & TABLE_NAME & ".CUST_CODE AND SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" _
            & " group by SOTTCLS1.MARKET_CODE"
        End If


        If optCM.Value = "C" Then
            ASCMAIN1.sql = "Select X.*, ARTCUST1.CUST_NAME DESC_VALUE from (" & ASCMAIN1.sql & ") X,ARTCUST1" _
            & " where X.CODE_VALUE = ARTCUST1.CUST_CODE"
        ElseIf optCM.Value = "M" Then
            ASCMAIN1.sql = "Select X.*, SOTMKTC1.MARKET_DESC DESC_VALUE from(" & ASCMAIN1.sql & ") X,SOTMKTC1" _
        & " where X.CODE_VALUE = SOTMKTC1.MARKET_CODE"
        End If

        grdSOTINVHX.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
        Dim DT As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTINVHX")
        grdSOTINVHX.DataSource = DT
        DT.Columns.Add("S00", GetType(System.Decimal), Mid(CS, 2))


        With grdSOTINVHX.DisplayLayout.Bands(0)
            If optCM.Value = "C" Then
                .Columns("CODE_VALUE").Header.Caption = "Customer"
                .Columns("CODE_VALUE").Width = 100
                .Columns("DESC_VALUE").Header.Caption = "Customer Name"
                .Columns("DESC_VALUE").Header.VisiblePosition = 1
                grdSOTINVHX.Text = "Sales By Customer"
            Else
                .Columns("CODE_VALUE").Header.Caption = "Market"
                .Columns("CODE_VALUE").Width = 100
                .Columns("DESC_VALUE").Header.Caption = "Market Desc"
                .Columns("DESC_VALUE").Header.VisiblePosition = 1
                grdSOTINVHX.Text = "Sales By Market"
            End If

            For P As Integer = 1 To PRDS
                With .Columns("S" & Format(P, "00"))
                    .Header.Appearance.TextHAlign = HAlign.Right
                    .Header.Caption = LEGENDs(P)
                    .Format = "###,##0"
                    .Width = 65
                    .Header.Appearance.TextHAlign = HAlign.Right
                    Create_Summary(grdSOTINVHX, "S" & Format(P, "00"))
                End With

            Next
            With .Columns("S00")
                .Header.VisiblePosition = 2
                .Header.Caption = "Total"
                .CellAppearance.BackColor = Color.WhiteSmoke
                .Hidden = False
                .Format = "###,##0"
                .Width = 65
                .Header.Appearance.TextHAlign = HAlign.Right
                .CellAppearance.TextHAlign = HAlign.Right
                Create_Summary(grdSOTINVHX, "S00")
            End With

            Create_Summary(grdSOTINVHX, "CODE_VALUE", "Count")

            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                If optIT.Value = "I" Then
                    grdCol.Header.Appearance.BackColor = Color.Violet
                Else
                    grdCol.Header.Appearance.BackColor = Color.LimeGreen
                End If
                grdCol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next

        End With

        CreateGraph_Trend(PRDS)
        CreateGraph_Totals()

        grdSOTINVHX.Visible = True
        optUD.Visible = True
        optGC.Visible = True
        optCM.Visible = True
        If grdSOTINVHX.Rows.Count = 0 Then
            chtTotals.Visible = False
            chtTrend.Visible = False
        Else
            chtTotals.Visible = True
            chtTrend.Visible = True


        End If


        ASCMAIN1.Progress("")

    End Sub

    Sub Sales_Forecasts()

        ASCDATA1.ExecuteSQL("Truncate Table " & DPTITMFS)

        ASCMAIN1.sql = "Select DPTITMF1.ITEM_CODE, DPTITMF1.MARKET_CODE, 'F' DATA_TYPE"
        For P = 0 To 12
            Dim YP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * P)
            ASCMAIN1.sql &= ", SUM (DECODE(DPTITMF1.OPS_YYYYPP_FC,'" & YP & "',DPTITMF1.FORECAST,0)) Q" & Format(P, "00")
        Next
        ASCMAIN1.sql &= " from DPTITMF1," & ICTITEMS & " ICTITEMS" _
        & " where DPTITMF1.ITEM_CODE = ICTITEMS.ITEM_CODE" _
        & " and DPTITMF1.OPS_YYYYPP = OPS_YYYYPP_FC" _
        & " and DPTITMF1.OPS_YYYYPP " _
        & " between '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12) & "'" _
        & "     and '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -0) & "'" _
        & " group by DPTITMF1.ITEM_CODE, DPTITMF1.MARKET_CODE"
        ASCDATA1.ExecuteSQL("Insert into " & DPTITMFS & " " & ASCMAIN1.sql)

        'ASCMAIN1.sql = "Select SATSSUMI.ITEM_CODE, NVL(SOTMKTC1.MARKET_CODE,SOTTCLS1.MARKET_CODE), 'S' DATA_TYPE"
        'For P = 1 To 12
        '    Dim YP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * P)
        '    ASCMAIN1.sql &= ", SUM (DECODE(SATSSUMI.OPS_YYYYPP,'" & YP & "',SATSSUMI.ORDR_QTY_SHIP,0)) Q" & Format(P, "00")
        'Next
        'ASCMAIN1.sql &= " from SATSSUMI," & ICTITEMS & " ICTITEMS,ARTCUST1,SOTTCLS1,SOTMKTC1" _
        '& " where SATSSUMI.ITEM_CODE = ICTITEMS.ITEM_CODE" _
        '& " and ARTCUST1.CUST_CODE (+) = SATSSUMI.CUST_CODE" _
        '& " and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" _
        '& " and SOTMKTC1.CUST_CODE (+) = SOTINVH2.CUST_CODE" & vbCrLf _
        '& " and SATSSUMI.OPS_YYYYPP " _
        '& " between '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12) & "'" _
        '& "     and '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1) & "'" _
        '& " group by SATSSUMI.ITEM_CODE, NVL(SOTMKTC1.MARKET_CODE,SOTTCLS1.MARKET_CODE)"

        ASCMAIN1.sql = "Select SOTINVH2.ITEM_CODE" & vbCrLf _
            & ", NVL(SOTMKTC1_CUST_CODE.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,NVL(SOTTCLS1.MARKET_CODE,'?'))) MARKET_CODE" & vbCrLf _
            & ", 'S' DATA_TYPE"
        For P = 0 To 12
            Dim YP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * P)
            ASCMAIN1.sql &= ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & YP & "',SOTINVH2.ORDR_QTY_SHIP,0)) Q" & Format(P, "00")
        Next
        ASCMAIN1.sql &= " from SOTINVH2," & ICTITEMS & " ICTITEMS,ARTCUST1,SOTTCLS1,SOTMKTC1,ICTWHSE1, SOTMKTC1 SOTMKTC1_CUST_CODE" & vbCrLf _
        & " where SOTINVH2.ITEM_CODE = ICTITEMS.ITEM_CODE" & vbCrLf _
        & " and ARTCUST1.CUST_CODE (+) = SOTINVH2.CUST_CODE" & vbCrLf _
        & " and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
        & "   and ICTWHSE1.WHSE_CODE = SOTINVH2.WHSE_CODE" & vbCrLf _
        & "   and NVL(ICTWHSE1.WHSE_MRP_EXC_IND,'0') <> '1'" & vbCrLf _
        & " and SOTMKTC1_CUST_CODE.CUST_CODE (+) = SOTINVH2.CUST_CODE" & vbCrLf _
        & " and SOTMKTC1.MARKET_CODE (+) = NVL(SOTTCLS1.MARKET_CODE,'?')" & vbCrLf _
        & " and SOTINVH2.INV_TYPE = 'I' " & vbCrLf _
        & " and SOTINVH2.ORDR_YYYYPP_UPDATED " & vbCrLf _
        & " between '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12) & "'" & vbCrLf _
        & "     and '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 0) & "'" & vbCrLf _
        & " group by SOTINVH2.ITEM_CODE, NVL(SOTMKTC1_CUST_CODE.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,NVL(SOTTCLS1.MARKET_CODE,'?')))"

        ASCDATA1.ExecuteSQL("Insert into " & DPTITMFS & " " & ASCMAIN1.sql)

        Fill_Records("DPTITMFS")

        ASCDATA1.ExecuteSQL("Update " & DPTITMFS & " Set MARKET_CODE = '*'")
        Fill_Records("DPTITMFS", "", False)

    End Sub

    Sub Setup_Item()

        ' Dim ITEM_CODE As String
        If byItem Then
            ITEM_CODE = Absx1.txtFor("ITEM_CODE").Text
        Else
            ITEM_CODE = grdDPTPLANX.ActiveRow.Cells("ITEM_CODE").Text
        End If
        tabDetails.Tabs(0).Text = ITEM_CODE & " - Forecasts" ' "Forecasts - " & ITEM_CODE

        chkHideTree.Checked = byItem
        chkHideTree.Enabled = Not byItem

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Setting Up Item " & ITEM_CODE)

        Setup_Item_Views("DPTITMFS", ITEM_CODE)
        Setup_Item_Views("ICTITEM1", ITEM_CODE)
        Sort_grdColumns(grdDPTITMFS, "MARKET_CODE")

        Fill_Records("DPTITMFX", ITEM_CODE)
        Sort_grdColumns(grdDPTITMFX, "MARKET_CODE")

        Fill_Records("POTORDRX", ITEM_CODE)
        Sort_grdColumns(grdPOTORDRX, "PO_ORDER_NO")
        grdPOTORDRX.Text = "Open POs Orders for Item " & ITEM_CODE
        Fill_Records("ICTPINVX", ITEM_CODE)
        Sort_grdColumns(grdPOTORDRX, "ETA_DATE_DC", , 1)
        Set_DC_Date(ITEM_CODE)

        Fill_Records("SOTORDRX", ITEM_CODE)
        'Sort_grdColumns(grdSOTORDRX, "ORDR_NO")
        Sort_grdColumns(grdSOTORDRX, "CUST_CODE")
        grdSOTORDRX.Text = "Open Sales Orders for Item " & ITEM_CODE

        Load_ICTSTATX(ITEM_CODE)
        Setup_grdICTSTATX(ITEM_CODE)

        Dim dvw As DataView = Nothing

        dvw = dst.Tables("DPTPLAND").DefaultView
        dvw.RowFilter = "ITEM_CODE = '" & ITEM_CODE & "'"
        ' Sort_grdColumns(grdDPTDBAL1, "DATE_BALANCE")

        ' THE FOLLOWING 2 SETS OF CODE ARE REQUIRED SO THAT THE GRIDS UNDER THE ITEMS SHOW JUST THE PLANS AND JUST THE COMMITMENTS FOR THE ITEMS ON DISPLAY
        dvw = dst.Tables("DPTPLAN1").DefaultView
        dvw.RowFilter = "ITEM_CODE = '" & ITEM_CODE & "'"
        Sort_grdColumns(grdDPTPLAN1, "DATE_REQUIRED")
        dvw = dst.Tables("POTORDR9").DefaultView
        dvw.RowFilter = "ITEM_CODE = '" & ITEM_CODE & "'"
        Sort_grdColumns(grdPOTORDR9, "DATE_REQUIRED")
        grdPOTORDR9.Text = "Open Production Commitments for Item " & ITEM_CODE

        If chkItemDetails.Checked And tabDetails.SelectedTab IsNot Nothing AndAlso tabDetails.SelectedTab.Key = "Sales" Then
            cmdFetchSales.PerformClick()
        Else
            If grdSOTINVHX.DataSource IsNot Nothing Then
                Try
                    DirectCast(grdSOTINVHX.DataSource, DataTable).Rows.Clear()
                Catch ex As Exception

                End Try

            End If
            optGC.Value = "B"
        End If


        With dst.Tables("DPTDBAL1")
            .Rows.Clear()
            grdDPTDBAL1.DataSource = Nothing
            If .Columns.Count > 4 Then
                For I = .Columns.Count - 1 To 4 Step -1
                    .Columns.Remove(.Columns(I).ColumnName)
                Next
            End If

            Dim WHSE_CODEs As New List(Of String)
            For Each ROW As DataRow In ASCDATA1.SelectDistinct _
            (dst.Tables("DPTPLAND").Select("ITEM_CODE = '" & ITEM_CODE & "'"), "WHSE_CODE").Rows
                Dim WHSE_CODE As String = ROW.Item("WHSE_CODE")
                WHSE_CODEs.Add(WHSE_CODE)
                Dim COLUMN_NAME As String = "WHSE_" & WHSE_CODE
                .Columns.Add(COLUMN_NAME, GetType(System.Int32))
            Next

            For Each ROW As DataRow In dst.Tables("DPTPLAND").Select("ITEM_CODE = '" & ITEM_CODE & "'")
                Dim WHSE_CODE As String = ROW.Item("WHSE_CODE")
                Dim COLUMN_NAME As String = "WHSE_" & WHSE_CODE
                Dim DATE_BALANCE As Date = ROW.Item("DATE_BALANCE")
                Dim ACTIVITY As Int32 = ROW.Item("ACTIVITY")
                Dim rowDPTDBAL1 As DataRow = .Rows.Find(New Object() {ITEM_CODE, DATE_BALANCE})
                If rowDPTDBAL1 Is Nothing Then
                    rowDPTDBAL1 = dst.Tables("DPTDBAL1").NewRow
                    rowDPTDBAL1.Item("ITEM_CODE") = ITEM_CODE
                    rowDPTDBAL1.Item("DATE_BALANCE") = DATE_BALANCE
                    dst.Tables("DPTDBAL1").Rows.Add(rowDPTDBAL1)
                End If
                rowDPTDBAL1.Item("ACTIVITY") = Val(rowDPTDBAL1.Item("ACTIVITY") & "") + ACTIVITY
                ' rowDPTDBAL1.Item("BALANCE") = Val(rowDPTDBAL1.Item("BALANCE") & "") + ACTIVITY
                rowDPTDBAL1.Item(COLUMN_NAME) = Val(rowDPTDBAL1.Item(COLUMN_NAME) & "") + ACTIVITY
            Next

            Dim BALANCE As Int32 = 0
            For Each rowDPTDBAL1 As DataRow In dst.Tables("DPTDBAL1").Select("ITEM_CODE = '" & ITEM_CODE & "'", "DATE_BALANCE")
                BALANCE += Val(rowDPTDBAL1.Item("ACTIVITY") & "")
                rowDPTDBAL1.Item("BALANCE") = BALANCE
            Next

            For Each WHSE_CODE As String In WHSE_CODEs
                Dim COLUMN_NAME As String = "WHSE_" & WHSE_CODE
                BALANCE = 0
                For Each rowDPTDBAL1 As DataRow In dst.Tables("DPTDBAL1").Select("ITEM_CODE = '" & ITEM_CODE & "'", "DATE_BALANCE")
                    BALANCE += Val(rowDPTDBAL1.Item(COLUMN_NAME) & "")
                    rowDPTDBAL1.Item(COLUMN_NAME) = BALANCE
                Next
            Next

            grdDPTDBAL1.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
            grdDPTDBAL1.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
            grdDPTDBAL1.DataSource = dst.Tables("DPTDBAL1")
            With grdDPTDBAL1.DisplayLayout.Bands(0)
                .Columns("ITEM_CODE").Hidden = True
                .Columns("DATE_BALANCE").Header.Caption = "Date"
                .Columns("BALANCE").Header.Caption = "Total"
                .Columns("ACTIVITY").Header.Caption = "Activity"

                .Columns("DATE_BALANCE").CellAppearance.BackColor = Color.Beige
                .Columns("BALANCE").CellAppearance.BackColor = Color.Beige
                .Columns("ACTIVITY").CellAppearance.BackColor = Color.Beige

                For Each WHSE_CODE As String In WHSE_CODEs
                    COLUMN_NAME = "WHSE_" & WHSE_CODE
                    .Columns(COLUMN_NAME).Header.Caption = WHSE_CODE
                Next
            End With
        End With
        Sort_grdColumns(grdDPTDBAL1, "DATE_BALANCE")

        Fill_Records("DPTMUPD0", ITEM_CODE)
        Dim QTY_BAL As Int64 = 0
        For Each row As DataRow In dst.Tables("DPTMUPD0").Select("", "DATE_REQ")
            QTY_BAL += Val(row.Item("QTY_OH") & "") _
                - Val(row.Item("QTY_REQ") & "") _
                + Val(row.Item("QTY_ORD") & "") _
                + Val(row.Item("QTY_PLN") & "")
            row.Item("QTY_BAL") = QTY_BAL
        Next
        Sort_grdColumns(grdDPTMUPD0, "DATE_REQ")

        '     Dim dvw As DataView = Nothing
        dvw = DirectCast(grdDPTMUPD1.DataSource, DataTable).DefaultView
        dvw.RowFilter = "ITEM_CODE = '" & ITEM_CODE & "'"
        Sort_grdColumns(grdDPTMUPD1, "CRDM_CODE")
        dvw = DirectCast(grdDPTMUPD2.DataSource, DataTable).DefaultView
        dvw.RowFilter = "ITEM_CODE = '" & ITEM_CODE & "'"
        Sort_grdColumns(grdDPTMUPD2, "EXC_MSG_CODE")


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


        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Setup_Item_Views(ByVal TABLE_NAME As String, ByVal ITEM_CODE As String)
        Dim dvw As DataView = dst.Tables(TABLE_NAME).DefaultView
        dvw.RowFilter = "ITEM_CODE = '" & ITEM_CODE & "'"
    End Sub

    Sub Setup_grdDPTITMFX()

        Dim c As UltraWinGrid.UltraGridColumn

        With grdDPTITMFX.DisplayLayout.Bands(0)

            c = .Columns("MARKET_CODE")
            c.Header.Caption = "Market"
            c.Hidden = False
            c.Width = 70
            c.ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
            c.Style = UltraWinGrid.ColumnStyle.EditButton
            'c.CellActivation = UltraWinGrid.Activation.NoEdit
            c.CellAppearance.BackColor = Color.Beige

            c = .Columns("MARKET_DESC")
            c.Header.Caption = "Description"
            c.Hidden = False
            c.Width = 150
            c.CellActivation = Activation.NoEdit
            'c.Header.VisiblePosition = 1
            'c.ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
            'c.Style = UltraWinGrid.ColumnStyle.EditButton
            'c.CellActivation = UltraWinGrid.Activation.NoEdit
            c.CellAppearance.BackColor = Color.Beige

            For P = -1 To 25
                Dim COLUMN_NAME As String = "FCPD"
                Dim COLUMN_CAPTION As String = "PastDue"
                If P >= 0 Then
                    COLUMN_NAME = "FC" & Format(P, "00")
                    COLUMN_CAPTION = YPF(P, 1)
                End If
                c = .Columns(COLUMN_NAME)
                c.Hidden = False
                c.Width = 60
                c.Header.Caption = COLUMN_CAPTION
                c.Header.Appearance.BackColor = Drawing.Color.HotPink ' Gold ' CornflowerBlue ' RosyBrown ' CadetBlue
                c.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
        End With

        ASCMAIN1.grdInitializeLayout(grdDPTITMFX)


    End Sub

    Sub Setup_grdDPTITMFS()

        Dim Gs As String()
        Dim G As UltraWinGrid.UltraGridGroup
        With grdDPTITMFS.DisplayLayout.Bands(0)
            .LevelCount = 2
            G = .Groups.Add("MARKET", "Market")
            G.Width = 100

            With .Columns("MARKET_CODE")
                .Group = G
                .Level = 0
                .CellActivation = UltraWinGrid.Activation.NoEdit
                .CellAppearance.BackColor = Color.WhiteSmoke
                .Width = 108
            End With

            With .Columns("MARKET_DESC")
                .Group = G
                .Level = 1
                .CellActivation = UltraWinGrid.Activation.NoEdit
                .CellAppearance.BackColor = Color.WhiteSmoke
                .Width = 70
            End With

            G = .Groups.Add("FS", "")
            G.Width = 20
            Gs = New String() {"F", "S"}
            For i As Integer = 0 To Gs.Length - 1
                With .Columns.Add(Gs(i))
                    .Group = G
                    .Level = i
                    .Width = 20
                    .Header.Appearance.TextHAlign = HAlign.Center
                    .CellAppearance.TextHAlign = HAlign.Center
                    .CellAppearance.BackColor = Color.White
                    If Gs(i) = "F" Then
                        .CellAppearance.ForeColor = Color.Green
                    Else
                        .CellAppearance.ForeColor = Color.DarkViolet
                    End If
                End With
            Next i

            For P = 12 To 0 Step -1
                Dim PX As String = Format(P, "00")
                G = .Groups.Add("P" & PX, YPP(P, 1))
                G.Header.Appearance.BackColor = Drawing.Color.Cornsilk ' BlueViolet ' .BlanchedAlmond
                G.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                G.Width = 60
                Gs = New String() {"F", "S"}
                For i As Integer = 0 To Gs.Length - 1
                    With .Columns(Gs(i) & PX)
                        .Group = G
                        .Level = i
                        .Width = 60
                        .CellAppearance.TextHAlign = HAlign.Right
                        .CellAppearance.BackColor = Color.White
                        If Gs(i) = "F" Then
                            .CellAppearance.ForeColor = Color.Green
                        Else
                            .CellAppearance.ForeColor = Color.DarkViolet
                        End If
                    End With
                Next
            Next

            .ColHeadersVisible = False
            For Each c As UltraWinGrid.UltraGridColumn In .Columns
                c.Hidden = False
                c.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
        End With
    End Sub

    Sub Setup_grdDPTPLANX()
        Dim Gs As String()
        Dim G As UltraWinGrid.UltraGridGroup
        With grdDPTPLANX.DisplayLayout.Bands(0)
            .LevelCount = 9
            G = .Groups.Add("CODE", "Code & Description")
            G.Width = 240
            Set_Item_Master_Column("ITEM_CODE", G, 0, 3)
            Set_Item_Master_Column("ITEM_DESC", G, 1, 3)
            Set_Item_Master_Column("DEPT_CODE", G, 2, 1)
            Set_Item_Master_Column("ITEM_ABC_CODE", G, 2, 1)
            Set_Item_Master_Column("MATL_CODE", G, 2, 1)
            Set_Item_Master_Column("STYLE_CODE", G, 3, 1)
            Set_Item_Master_Column("COLOR_CODE", G, 3, 1)
            Set_Item_Master_Column("SIZE_CODE", G, 3, 1)


            G = .Groups.Add("STATUS", "Status")
            G.Width = 130
            G.Header.Appearance.BackColor = Drawing.Color.Lime ' Cornsilk '.LimeGreen ' . Honeydew  ' Orange ' LightBlue
            G.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Gs = New String() {"MTD", "ONH", "COM", "AVA", "RTN", "MNQ"}
            For i As Integer = 0 To Gs.Length - 1
                With .Columns.Add("STALBL_" & Gs(i))
                    .Group = G
                    .Level = i
                    .CellAppearance.BackColor = Color.Beige
                    .CellAppearance.ForeColor = Color.DarkViolet
                    .Width = 65
                End With
                With .Columns(Gs(i))
                    .Group = G
                    .Level = i
                    .CellAppearance.BackColor = Color.White
                    .Width = 65
                End With
            Next

            For Each TLBL As String In New String() {"L", "3", "2", "1"}
                Dim CAPTION As String = "Trend"
                If TLBL <> "L" Then CAPTION = YPP(Val(TLBL), 1)
                G = .Groups.Add("TREND_" & TLBL, CAPTION)
                G.Header.Appearance.BackColor = Drawing.Color.Purple
                G.Header.Appearance.ForeColor = Drawing.Color.White
                G.Header.Appearance.BackGradientStyle = GradientStyle.GlassRight20

                If TLBL = "L" Then
                    G.Header.Appearance.TextHAlign = HAlign.Left
                Else
                    G.Header.Appearance.TextHAlign = HAlign.Right
                End If

                Gs = New String() {"SHP", "DEM", "ONH", "PCT"}
                For i As Integer = 0 To Gs.Length - 1
                    If TLBL = "L" Then .Columns.Add("TREND_" & Gs(i) & "_" & TLBL)
                    With .Columns("TREND_" & Gs(i) & "_" & TLBL)
                        .Group = G
                        .Level = i
                        .Width = 70
                        If TLBL = "L" Then
                            .CellAppearance.BackColor = Color.Beige
                        Else
                            If Gs(i) = "PCT" Then
                                .Format = "##0.0"
                            Else
                                .Format = "#,##0"
                            End If
                        End If
                    End With
                Next
                G.Width = 70
                G.Hidden = True
            Next

            G = .Groups.Add("TOTALS", "Totals")
            G.Header.Appearance.BackColor = Drawing.Color.Yellow
            G.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            G.Header.Appearance.TextHAlign = HAlign.Right
            Gs = New String() {"PO", "DEM", "ONH", "POS", "OS", "GP", "SO", "SI", "ST"}
            For i As Integer = 0 To Gs.Length - 1
                If Gs(i) = "GP" Or Gs(i) = "SI" Or Gs(i) = "ST" Then
                    ' If Gs(i) = "OS" Or Gs(i) = "GP" Or Gs(i) = "SI" Or Gs(i) = "ST" Then
                    With .Columns.Add("TOTAL_" & Gs(i))
                        .DataType = GetType(System.Int32)
                        .Format = "###,##0"
                        .CellAppearance.TextHAlign = HAlign.Right
                    End With
                End If

                With .Columns("TOTAL_" & Gs(i))
                    .Group = G
                    .Level = i
                    .Width = 70

                    If Gs(i) = "POS" Then
                        .CellAppearance.TextHAlign = HAlign.Center
                    End If
                End With
            Next
            G.Width = 70


            G = .Groups.Add("DATA", "")
            G.Header.Appearance.BackColor = Drawing.Color.Yellow
            G.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            'Gs = New String() {"SUP", "DEM", "ONH", "POS", "OS", "GP", "SO", "SI", "ST"}
            Gs = New String() {"PO", "DEM", "ONH", "POS", "OS", "GP", "SO", "SI", "ST"}
            For i As Integer = 0 To Gs.Length - 1
                With .Columns.Add("DATA_" & Gs(i))
                    .Group = G
                    .Level = i
                    .CellAppearance.BackColor = Color.Beige
                    .CellAppearance.ForeColor = Color.DarkViolet
                End With
            Next
            G.Width = 40

            For P = -1 To 25
                Dim PX As String = IIf(P = -1, "PD", Format(P, "00"))
                Dim CAPTION As String = "PastDue"
                If P > -1 Then CAPTION = YPF(P, 1)
                G = .Groups.Add("P" & PX, CAPTION)
                G.Header.Appearance.BackColor = Drawing.Color.Yellow
                If P = -1 Then G.Header.Appearance.ForeColor = Color.Red
                G.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                G.Header.Appearance.TextHAlign = HAlign.Right
                G.Width = 60
                Gs = New String() {"PO", "DEM", "ONH", "POS", "OS", "GP", "SO", "SI", "ST"}
                For i As Integer = 0 To Gs.Length - 1
                    If Gs(i) = "GP" Or Gs(i) = "SI" Or Gs(i) = "ST" Then
                        'If Gs(i) = "OS" Or Gs(i) = "GP" Or Gs(i) = "SI" Or Gs(i) = "ST" Then
                        With .Columns.Add(Gs(i) & PX)
                            .DataType = GetType(System.Int32)
                            .Format = "###,##0"
                            .CellAppearance.TextHAlign = HAlign.Right
                        End With
                    End If
                    With .Columns(Gs(i) & PX)
                        .Group = G
                        .Level = i
                        .Width = 60
                        If Gs(i) = "POS" Then
                            .CellAppearance.BorderColor = Color.DarkViolet
                            .CellAppearance.BackColor = Color.WhiteSmoke
                            .Format = "#,##0.0"
                        Else
                            .CellAppearance.BackColor = Color.White
                        End If
                    End With
                Next


                ' If P = -1 Then G.Hidden = True ' this line hides the PD column
            Next
            .ColHeadersVisible = False
            For Each c As UltraWinGrid.UltraGridColumn In .Columns
                If c.Key.StartsWith("DPTPLANX_SOTORDRM") Then
                Else
                    c.Hidden = False
                    c.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        Setup_OverShort()
    End Sub

#Region "grdDPTITMFX"

    Private Sub grdDPTITMFX_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdDPTITMFX.AfterCellUpdate
        'grdSOTORDR2.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
    End Sub

    Private Sub grdDPTITMFX_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdDPTITMFX.BeforeExitEditMode
        If e.CancellingEditOperation Then Exit Sub

        With grdDPTITMFX.ActiveCell
            Select Case .Column.Key
                Case "MARKET_CODE"
                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                        cdr = LookUp("SOTMKTC1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid " & .Column.Header.Caption & "(" & .Text & ")")
                            .Value = ""
                            e.Cancel = True
                        Else
                            .Row.Cells("MARKET_DESC").Value = cdr.Item("MARKET_DESC") & ""
                        End If
                    End If
            End Select
        End With

    End Sub

    Private Sub grdDPTITMFX_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdDPTITMFX.BeforeRowUpdate
        If e.Row.IsAddRow Then
            e.Row.Cells("ITEM_CODE").Value = grdDPTPLANX.ActiveRow.Cells("ITEM_CODE").Text
        End If
    End Sub

    Private Sub grdDPTITMFX_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdDPTITMFX.ClickCellButton
        Dim sql_where As String = ""
        Select Case grdDPTITMFX.ActiveCell.Column.Key
            Case "MARKET_CODE"
        End Select

        Call grdClickCellButton(grdDPTITMFX, sql_where, False)
    End Sub

    Private Sub grdDPTITMFX_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdDPTITMFX.AfterRowActivate
        If grdDPTITMFX.ActiveRow.IsAddRow Then
            grdDPTITMFX.DisplayLayout.Bands(0).Columns("MARKET_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            grdDPTITMFX.DisplayLayout.Bands(0).Columns("MARKET_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        End If
    End Sub

    Private Sub grdDPTITMFX_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdDPTITMFX.AfterExitEditMode
        With grdDPTITMFX
            Select Case .ActiveCell.Column.Key
                Case "MARKET_CODE"
                    If .ActiveCell.Text <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(.ActiveCell.Text, .ActiveCell.Column.Key)
                        '.ActiveCell.Row.Cells("MARKET_DESC").Value = MARKET_DESC
                    End If
            End Select
        End With
    End Sub

    Private Sub grdDPTITMFX_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdDPTITMFX.Error
        grdDPTITMFX.ActiveRow.CancelUpdate()
    End Sub
#End Region

    Private Sub chkEditForecasts_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkEditForecasts.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub

        If chkEditForecasts.Checked Then
            ASCMAIN1.sql = $"Select ICTITEM1.COLLECTION_CODE, ICTCOLL1.BRAND_CODE from ICTITEM1,ICTCOLL1 where ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTiON_CODE and ICTITEM1.ITEM_CODE = '{ITEM_CODE}'"
            Dim rowCB As DataRow = ASCDATA1.GetDataRow
            Dim BRAND_CODE As String = rowCB.Item("BRAND_CODE")

            If Not ASCMAIN1.Logical_Lock("DPTITMFX", $"*:{BRAND_CODE}",,,, 1) Then
                chkEditForecasts.Checked = False
                Exit Sub
            End If

            Fill_Records("DPTITMFX", ITEM_CODE, True)
            ReCalculate_FC()
        Else
            ASCMAIN1.MultiTask_Release(,, 1)
        End If

        With grdDPTITMFX.DisplayLayout.Override
            If chkEditForecasts.Checked Then
                MsgBox("Warning: Changes made to the Forecasts will be written to the Database", MsgBoxStyle.OkOnly, "Verification")
                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                .AllowDelete = DefaultableBoolean.True
                .AllowUpdate = DefaultableBoolean.True

                tabDetails.Tabs("Forecasts").Selected = True
                tabProjections.Tabs("Forecasts by Market").Selected = True
            Else
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowDelete = DefaultableBoolean.False
                .AllowUpdate = DefaultableBoolean.False

            End If
        End With

    End Sub

    Private Sub grdDPTITMFX_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdDPTITMFX.AfterRowUpdate
        Dim ITEM_CODE As String = grdDPTPLANX.ActiveRow.Cells("ITEM_CODE").Text
        dst.Tables("DPTITMF1").Rows.Clear()

        BeginTrans()
        ASCMAIN1.sql = "Delete from DPTITMF1 " _
        & " where ITEM_CODE = '" & ITEM_CODE & "'" _
        & " and OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
        ASCDATA1.ExecuteSQL()
        For Each rowDPTITMFX As DataRow In dst.Tables("DPTITMFX").Select("", "")
            Dim MARKET_CODE As String = rowDPTITMFX.Item("MARKET_CODE")
            For P As Integer = -1 To 25
                Dim COLUMN_NAME As String = IIf(P = -1, "FCPD", "FC" & Format(P, "00"))
                Dim FORECAST As Integer = Val(rowDPTITMFX.Item(COLUMN_NAME) & "")
                If FORECAST <> 0 Then
                    Dim rowDPTITMF1 As DataRow = dst.Tables("DPTITMF1").NewRow
                    rowDPTITMF1.Item("ITEM_CODE") = ITEM_CODE
                    rowDPTITMF1.Item("MARKET_CODE") = MARKET_CODE
                    rowDPTITMF1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
                    If P = -1 Then
                        rowDPTITMF1.Item("OPS_YYYYPP_FC") = "000000"
                    Else
                        rowDPTITMF1.Item("OPS_YYYYPP_FC") = YPF(P, 0)
                    End If
                    rowDPTITMF1.Item("FORECAST") = FORECAST
                    dst.Tables("DPTITMF1").Rows.Add(rowDPTITMF1)
                End If
            Next
        Next
        Update_Record_TDA("DPTITMF1")
        CommitTrans()

        ReCalculate_FC()
    End Sub

    Private Sub grdDPTITMFX_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdDPTITMFX.AfterRowsDeleted
        Dim ITEM_CODE As String = grdDPTPLANX.ActiveRow.Cells("ITEM_CODE").Text
        If Not BeforeRowsDeletedRows Is Nothing Then
            BeginTrans()
            For Each C() As VariantType In BeforeRowsDeletedRows
                Dim MARKET_CODE As String = C(1)
                ASCMAIN1.sql = "Delete from DPTITMF1 " _
            & " where ITEM_CODE = '" & ITEM_CODE & "'" _
            & " and OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" _
            & " and MARKET_CODE = '" & MARKET_CODE & "'"
                ASCDATA1.ExecuteSQL()
            Next
            CommitTrans()
        End If

        ReCalculate_FC()
    End Sub

    Sub ReCalculate_FC()
        Dim rowDPTITMFM As DataRow = dst.Tables("DPTITMFM").Rows.Find(grdDPTPLANX.ActiveRow.Cells("ITEM_CODE").Text)
        For P As Integer = -1 To 25
            Dim COLUMN_NAME As String = IIf(P = -1, "FCPD", "FC" & Format(P, "00"))
            rowDPTITMFM.Item(COLUMN_NAME) = dst.Tables("DPTITMFX").Compute("SUM(" & COLUMN_NAME & ")", "")
        Next
        Calculate_Position(ITEM_CODE)
    End Sub

    Sub Setup_grdICTSTATX(ByVal ITEM_CODE As String)

        grdICTSTATX.DisplayLayout.Bands(0).Columns("OPS_YYYYPP").Hidden = True
        Dim RYP As String = ASCMAIN1.CYP

        Dim sqlF As String = ""
        Dim sqlW As String = ""
        Dim sqlG As String = ""
        sqlF = "OPS_YYYYPP, ITEM_CODE, WHSE_CODE"
        sqlW = " where OPS_YYYYPP = '" & RYP & "'"
        sqlG = "OPS_YYYYPP, ITEM_CODE, WHSE_CODE"
        grdICTSTATX.Text = "Item Status / Activity by Warehouse for " & ASCMAIN1.Get_Legend(RYP)
        splMain.Panel2Collapsed = False

        Load_ICTTRANX(ITEM_CODE)

        Dim SQL As String = "" _
        & " Select " & sqlF & vbCrLf _
        & ", SUM (WHSE_QTY_BEG) WHSE_QTY_BEG" & vbCrLf _
        & ", SUM (WHSE_QTY_SHP) WHSE_QTY_SHP" & vbCrLf _
        & ", SUM (WHSE_QTY_RTN) WHSE_QTY_RTN" & vbCrLf _
        & ", SUM (WHSE_QTY_REC) WHSE_QTY_REC" & vbCrLf _
        & ", SUM (WHSE_QTY_ADJ) WHSE_QTY_ADJ" & vbCrLf _
        & ", SUM (WHSE_QTY_XFR) WHSE_QTY_XFR" & vbCrLf _
        & ", SUM (WHSE_QTY_CON) WHSE_QTY_CON" & vbCrLf _
        & ", SUM (WHSE_QTY_RTV) WHSE_QTY_RTV" & vbCrLf _
        & ", SUM (WHSE_QTY_PHY) WHSE_QTY_PHY" & vbCrLf _
        & ", SUM (WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND" & vbCrLf _
        & ", SUM (WHSE_QTY_ONPO) WHSE_QTY_ONPO" & vbCrLf _
        & ", SUM (WHSE_QTY_PLAN) WHSE_QTY_PLAN" & vbCrLf _
        & ", SUM (WHSE_QTY_OPEN) WHSE_QTY_OPEN" & vbCrLf _
        & ", SUM (WHSE_QTY_PICK) WHSE_QTY_PICK" & vbCrLf _
        & ", SUM (WHSE_QTY_COMM) WHSE_QTY_COMM" & vbCrLf _
        & ", SUM (WHSE_QTY_HOLD) WHSE_QTY_HOLD" & vbCrLf _
        & " from " & ICTSTATX & vbCrLf _
        & sqlW & vbCrLf _
        & " group by " & sqlG

        SQL = "Select X.*,ICTWHSE1.WHSE_TYPE,ICTWHSE1.WHSE_MRP_EXC_IND " _
        & " from ICTWHSE1,(" & SQL & ") X where ICTWHSE1.WHSE_CODE (+) = X.WHSE_CODE"
        Fill_Records("ICTSTATX", "", True, SQL)

        Sort_grdColumns(grdICTSTATX, "OPS_YYYYPP".ToLower & ",WHSE_CODE")
    End Sub

    Sub Load_ICTSTATX(ByVal ITEM_CODE As String)

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
        & " FROM ICTSTAT1 WHERE ITEM_CODE = '" & ITEM_CODE & "'" _
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
        & " FROM ICTSTAT5 WHERE ITEM_CODE = '" & ITEM_CODE & "'" _
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
        & " FROM ICTSTAT2 WHERE ITEM_CODE = '" & ITEM_CODE & "'" _
        & " GROUP BY ITEM_CODE, WHSE_CODE" _
        & ") group by OPS_YYYYPP, ITEM_CODE, WHSE_CODE"
        ASCDATA1.ExecuteSQL()

    End Sub

    Sub Load_ICTTRANX(ByVal ITEM_CODE As String)
        ASCDATA1.ExecuteSQL("Truncate Table " & ICTTRANX)
        Dim RYP As String = ASCMAIN1.CYP

        ASCMAIN1.sql = "Insert into " & ICTTRANX _
        & " SELECT T2.OPS_YYYYPP, T2.ITEM_CODE, T1.WHSE_CODE" _
        & ", T1.ADJ_NO TRAN_NO, T1.ADJ_SOURCE TRAN_SOURCE" _
        & ", T1.INIT_DATE, T1.INIT_OPER" _
        & ", T1.ADJ_DATE TRAN_DATE, 'A' TRAN_TYPE" _
        & ", T2.ADJ_QTY TRAN_QTY, ICTREAS1.REASON_DESC TRAN_NOTE" _
        & " FROM ICTIADJ1 T1,ICTIADJ2 T2, ICTREAS1" _
        & " WHERE T1.ADJ_NO = T2.ADJ_NO" _
        & " and T2.ITEM_CODE = '" & ITEM_CODE & "'" _
        & " and T2.OPS_YYYYPP = '" & RYP & "'" _
        & " and ICTREAS1.REASON_CODE = T1.REASON_CODE"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into " & ICTTRANX _
        & " SELECT T2.OPS_YYYYPP, T2.ITEM_CODE, T1.WHSE_CODE" _
        & ", T1.XFR_NO TRAN_NO, T1.XFR_SOURCE TRAN_SOURCE" _
        & ", T1.INIT_DATE, T1.INIT_OPER" _
        & ", T1.XFR_DATE TRAN_DATE, 'T' TRAN_TYPE" _
        & ", -1 * T2.XFR_QTY TRAN_QTY, 'XFR to ' || T1.WHSE_CODE_TO TRAN_NOTE" _
        & " FROM ICTIXFR1 T1,ICTIXFR2 T2" _
        & " WHERE T1.XFR_NO = T2.XFR_NO" _
        & " and T2.ITEM_CODE = '" & ITEM_CODE & "'" _
        & " and T2.OPS_YYYYPP = '" & RYP & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into " & ICTTRANX _
        & " SELECT T2.OPS_YYYYPP, T2.ITEM_CODE, T1.WHSE_CODE_TO" _
        & ", T1.XFR_NO TRAN_NO, T1.XFR_SOURCE TRAN_SOURCE" _
        & ", T1.INIT_DATE, T1.INIT_OPER" _
        & ", T1.XFR_DATE TRAN_DATE, 'T' TRAN_TYPE" _
        & ", T2.XFR_QTY TRAN_QTY, 'XFR from ' || T1.WHSE_CODE TRAN_NOTE" _
        & " FROM ICTIXFR1 T1,ICTIXFR2 T2" _
        & " WHERE T1.XFR_NO = T2.XFR_NO" _
        & " and T2.ITEM_CODE = '" & ITEM_CODE & "'" _
        & " and T2.OPS_YYYYPP = '" & RYP & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into " & ICTTRANX _
        & " SELECT T2.ORDR_YYYYPP_UPDATED, T2.ITEM_CODE, T1.WHSE_CODE" _
        & ", NULL TRAN_NO, 'S' TRAN_SOURCE" _
        & ", NULL INIT_DATE, NULL INIT_OPER" _
        & ", T1.INV_DATE TRAN_DATE, DECODE(T2.INV_TYPE,'I','S','C','R',NULL) TRAN_TYPE" _
        & ", SUM (T2.ORDR_QTY_SHIP) TRAN_QTY, ' Line Items:' || Count (*) TRAN_NOTE" _
        & " FROM SOTINVH1 T1,SOTINVH2 T2" _
        & " WHERE T1.INV_NO = T2.INV_NO" _
        & " and T1.INV_TYPE = T2.INV_TYPE" _
        & " and T2.ITEM_CODE = '" & ITEM_CODE & "'" _
        & " and T2.ORDR_YYYYPP_UPDATED = '" & RYP & "'" _
        & " GROUP BY T2.ORDR_YYYYPP_UPDATED, T2.ITEM_CODE, T1.WHSE_CODE, T2.INV_DATE, DECODE(T2.INV_TYPE,'I','S','C','R',NULL)"
        'ASCDATA1.ExecuteSQL()
    End Sub

    Sub Setup_grdICTTRANX()

        'Dim WHSE_CODE As String = grdICTSTATX.ActiveRow.Cells("WHSE_CODE").Text

        'Dim sql As String = "Select * from " & ICTTRANX _
        '& " where WHSE_CODE = '" & WHSE_CODE & "'"

        'Fill_Records("ICTTRANX", "", True, Sql)
        'Sort_grdColumns(grdICTTRANX, "INIT_DATE,TRAN_DATE,TRAN_TYPE,TRAN_NO")

        'grdICTTRANX.Text = "Transaction Details for Whse " & WHSE_CODE
    End Sub

    Private Sub grdICTSTATX_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTSTATX.AfterRowActivate
        'SETUP_grdICTTRANX()
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

    Sub Calculate_Position(ByVal ITEM_CODE As String)

        Dim POS_STATUS_CODES As String = ""

        grdDPTPLANX.Tag = "X"

        Dim rowDPTPLANX As DataRow = dst.Tables("DPTPLANX").Rows.Find(ITEM_CODE)

        Dim ITEM_POS_MAX As Decimal = Val(rowDPTPLANX.GetParentRow("ICTITEM1_DPTPLANX").Item("ITEM_POS_MAX") & "")
        Dim ITEM_POS_MIN As Decimal = Val(rowDPTPLANX.GetParentRow("ICTITEM1_DPTPLANX").Item("ITEM_POS_MIN") & "")
        Dim TOTAL_POS As String = Format(ITEM_POS_MAX, "0.0") & "/" & Format(ITEM_POS_MIN, "0.0")
        If ITEM_POS_MAX = 0 And ITEM_POS_MIN = 0 Then ' TOTAL_POS = "0.0/0.0" Then
            TOTAL_POS = Format(Val(ROWs("DPTPARM1").Item("DP_PARM_POS_MAX") & ""), "0.0") & "/" & Format(Val(ROWs("DPTPARM1").Item("DP_PARM_POS_MIN") & ""), "0.0")
        End If
        rowDPTPLANX.Item("TOTAL_POS") = TOTAL_POS

        Dim LAST_FC_P As Int32 = 0
        Dim DEM(25) As Int32
        For P As Integer = 0 To 25
            DEM(P) = Val(rowDPTPLANX.Item("DEM" & Format(P, "00")) & "")
            If DEM(P) <> 0 Then
                LAST_FC_P = P
            End If
        Next

        Dim EOM_at_MAXP As Int32 = Val(rowDPTPLANX.Item("ONH" & Format(LAST_FC_P, "00")) & "")

        Dim rowDPTPLANO As DataRow = dst.Tables("DPTPLANO").Rows.Find(ITEM_CODE)
        If rowDPTPLANO Is Nothing Then
            rowDPTPLANO = dst.Tables("DPTPLANO").NewRow
            rowDPTPLANO.Item("ITEM_CODE") = ITEM_CODE
            dst.Tables("DPTPLANO").Rows.Add(rowDPTPLANO)
        End If

        Dim PLAN_QTY_CUM As Int32 = 0
        'If ITEM_CODE = "EB2029" Then Stop ' WHY DO I HIT THIS ROUTINE 3 TIMES FOR A SINGLE ITEM?

        For P As Integer = -1 To 24

            Dim PX As String = IIf(P = -1, "PD", Format(P, "00"))

            Dim EOM As Int32 = Val(rowDPTPLANX.Item("ONH" & PX) & "")
            Dim EOM_BAL As Int32 = EOM

            Dim POS As Decimal = 0

            Dim EOM_MAX As Int32 = 0
            Dim EOM_MIN As Int32 = 0

            Dim calc_position As Boolean = True
            For F As Integer = P + 1 To 25
                If DEM(F) = 0 Then
                    Exit For
                End If
                If EOM_BAL <= 0 Then
                    calc_position = False
                    'Exit For
                End If
                If calc_position Then
                    If EOM_BAL > DEM(F) Then
                        POS += 1
                        EOM_BAL -= DEM(F)
                    Else
                        POS += EOM_BAL / DEM(F)
                        calc_position = False
                        'Exit For
                    End If
                End If
                If F - P <= ITEM_POS_MAX Then
                    EOM_MAX += DEM(F)
                ElseIf (F - P) - 1 <= ITEM_POS_MAX Then
                    EOM_MAX += DEM(F) * (ITEM_POS_MAX - ((F - P) - 1))
                End If
                If F - P <= ITEM_POS_MIN Then
                    EOM_MIN += DEM(F)
                ElseIf (F - P) - 1 <= ITEM_POS_MIN Then
                    EOM_MIN += DEM(F) * (ITEM_POS_MIN - ((F - P) - 1))
                End If
                If Not calc_position And F - P > ITEM_POS_MAX Then
                    Exit For
                End If
            Next

            rowDPTPLANX.Item("POS" & PX) = POS

            'If P > 0 Then
            Dim POS_STATUS_CODE As String = ""
            Dim C As String = "POS" & PX
            'G Past Due PO Rec
            'H Negative OH    
            'If EOM_BAL < 0 Then
            If EOM < 0 Then
                POS_STATUS_CODE = "A" 'A EOM Qty OH < 0 
            Else
                If POS > ITEM_POS_MAX Then
                    POS_STATUS_CODE = "D" 'D Pos > Max      
                    'ElseIf EOM_BAL <> 0 And DEM(P + 1) = 0 Then
                ElseIf EOM <> 0 And DEM(P + 1) = 0 Then
                    POS_STATUS_CODE = "F" 'F Qty w/No Demand
                ElseIf (POS < ITEM_POS_MIN And P + ITEM_POS_MIN <= LAST_FC_P) Then
                    POS_STATUS_CODE = "B" 'B Pos < Min   
                ElseIf POS = LAST_FC_P - P And EOM_at_MAXP <> 0 Then
                    POS_STATUS_CODE = "E" 'E Qty > Demand  
                Else
                    POS_STATUS_CODE = "C" 'C Min < Pos < Max
                End If
            End If

            POS_STATUS_CODES &= POS_STATUS_CODE
            'End If

            Dim OS As Int32 = 0
            rowDPTPLANO.Item("OS" & PX) = DBNull.Value
            If POS_STATUS_CODE = "A" Or POS_STATUS_CODE = "B" Then
                OS = EOM - EOM_MIN
                rowDPTPLANO.Item("OS" & PX) = OS
            End If
            If POS_STATUS_CODE = "D" Or POS_STATUS_CODE = "E" Or POS_STATUS_CODE = "F" Then
                OS = EOM - EOM_MAX
                rowDPTPLANO.Item("OS" & PX) = OS
            End If

            rowDPTPLANX.Item("OS" & PX) = rowDPTPLANO.Item("OS" & PX)

            generating_plans = False
            If generating_plans Then
                Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
                Dim PLAN_QTY As Int32 = 0
                Dim OS_PLUS_PLANS As Int32 = OS + PLAN_QTY_CUM

                If OS_PLUS_PLANS > 0 Then ' we have an oversupply
                    Dim PO_QTY As Int32 = Val(rowDPTPLANX.Item("PO" & PX) & "")
                    If PO_QTY <> 0 Then
                        If PO_QTY > OS_PLUS_PLANS Then
                            PLAN_QTY = -1 * OS_PLUS_PLANS
                        Else
                            PLAN_QTY = -1 * PO_QTY
                        End If
                    End If
                ElseIf OS_PLUS_PLANS < 0 Then ' we have a shortfall
                    Dim ITEM_PO_QTY_MIN As Int32 = Val(rowICTITEM1.Item("ITEM_PO_QTY_MIN") & "")
                    Dim ITEM_PO_QTY_MULT As Int32 = Val(rowICTITEM1.Item("ITEM_PO_QTY_MULT") & "")

                    PLAN_QTY = EOM_MAX - (EOM + PLAN_QTY_CUM)
                    If ITEM_PO_QTY_MIN > 0 Then
                        If PLAN_QTY < ITEM_PO_QTY_MIN Then
                            PLAN_QTY = ITEM_PO_QTY_MIN
                        End If
                    End If

                    If ITEM_PO_QTY_MULT > 0 Then
                        If PLAN_QTY Mod ITEM_PO_QTY_MULT <> 0 Then
                            PLAN_QTY += ITEM_PO_QTY_MULT - (PLAN_QTY Mod ITEM_PO_QTY_MULT)
                        End If
                    End If
                End If

                rowDPTPLANO.Item("GP" & PX) = PLAN_QTY
                PLAN_QTY_CUM += PLAN_QTY

            Else

                Dim PLAN_QTY As Int32 = 0

                rowDPTPLANO.Item("GP" & PX) = PLAN_QTY
                PLAN_QTY_CUM += PLAN_QTY

            End If
        Next
        rowDPTPLANX.Item("LAST_FC_P") = LAST_FC_P
        rowDPTPLANX.Item("ITEM_POS_MAX") = ITEM_POS_MAX
        rowDPTPLANX.Item("ITEM_POS_MIN") = ITEM_POS_MIN

        grdDPTPLANX.Tag = ""
        rowDPTPLANX.Item("POS_STATUS_CODES") = POS_STATUS_CODES

    End Sub

    Private Sub grdDPTITMFS_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdDPTITMFS.InitializeRow
        e.Row.Cells("F").Value = "FC"
        e.Row.Cells("S").Value = "Ship"
        If e.Row.Cells("MARKET_CODE").Value = "*" Then
            e.Row.CellAppearance.BackColor = Color.WhiteSmoke
        End If


        If e.Row Is Nothing OrElse e.Row.IsFilterRow OrElse Not e.Row.IsDataRow Then
            Exit Sub
        End If

        Dim MARKET_CODE As String = e.Row.Cells("MARKET_CODE").Value & ""

        If MARKET_CODE = "*" Then
            Exit Sub
        End If

        For P = 0 To 6
            Dim sqlw As String = "ITEM_CODE = '" & ITEM_CODE & "' and MARKET_CODE = '" & MARKET_CODE & "' and OPS_YYYYPP_FC = '" & YPP(P, 0) & "' and STATUS IS NULL"
            Dim rows() As DataRow = dst.Tables("DPTITMF2").Select(sqlw, "INIT_DATE")
            If rows.Length <> 0 Then
                Dim TT As String = ""
                For Each row As DataRow In rows
                    TT &= Format(row.Item("INIT_DATE"), "MM/dd/yyyy") & vbTab & Format(Val(row.Item("FORECAST") & ""), "#,##0") & vbTab & row.Item("FORECAST_NOTE")
                    TT &= vbCrLf
                Next
                With e.Row.Cells("F" & Format(P, "00"))
                    .Appearance = cellHasNotes
                    .ToolTipText = TT
                End With
            Else
                'e.Row.Cells("").Appearance = nothing
            End If
        Next

    End Sub

    Private Sub CreateToolTipMessage(ByVal cell As UltraWinGrid.UltraGridCell)
        tooltip_title = ""
        Select Case cell.Column.Key
            Case "MATL_CODE"
                tooltip_msg = "Material"
            Case "DEPT_CODE"
                tooltip_msg = "Department"
            Case "ITEM_ABC_CODE"
                tooltip_title = "ABC Code"
                tooltip_msg = "Calculated using the ABC Classification Function." & vbCrLf & "The ABC code sets the Min/Max Position, and the Minumum Days of Supply (which may then be modified on an individual item by item basis)."
            Case "STYLE_CODE"
                tooltip_msg = "Style"
            Case "COLOR_CODE"
                tooltip_msg = "Color"
            Case "SIZE_CODE"
                tooltip_msg = "Size"
            Case "SUP00"
                tooltip_title = "Supply - Purchase Orders & Plans."
                tooltip_msg = "Background shows up in Pink in month of Critical Lead Time Date." & vbCrLf & "Qtys show up in Red if they are within the Critical Lead Time."
            Case "TREND_PCT_3", "TREND_PCT_2", "TREND_PCT_1"
                tooltip_title = "Trend Pct = (Actual - Forecast) / Forecast"
                tooltip_msg = "Shown in Light Red if negative, and Dark Red if negative by more than 10%." & vbCrLf & "Shown in Light Green if positive, and Dark Green if positive by more than 10%."
            Case "POS00", "POS01", "POS02", "POS03"
                tooltip_title = "Months of Supply"
                tooltip_msg = "Shown in " & vbCrLf _
                & " - Red Backcolor if Projected Qty On Hand is negative" & vbCrLf _
                & " - Green Backcolor if > Max Pos" & vbCrLf _
                & " - Yellow Backcolor if < Min Pos (and there is demand > Min Pos)" & vbCrLf _
                & " - Green Forecolor if there is Qty in Excess of Demand" & vbCrLf _
                & " - White Backcolor if Pos is within Max and Min" & vbCrLf _
                & " - and Steel Blue Backcolor there is no Position and no Demand"

            Case "ONH00", "ONH01", "ONH02", "ONH03"
                tooltip_title = "End of Month (EOM) Qty on Hand"
                tooltip_msg = "Calculated by taking the Prior Period EOM + Current Period Supply - Current Period Demand." & vbCrLf & "1st Month's value is calculated as Qty On Hand (now), plus MTD Shipments, less Total Current Month's Forecasted Shipments." & vbCrLf & " - This is how we account for over shipments in the current month." & vbCrLf & " - however, if MTD Shipments + Open Orders (with Ship Date by End of Current Month) is > than Forecasted Shipments for the Current Month," & vbCrLf & "  MTD + Open (w/Ship Date <= Cur EOM) is used in place of the Current Month's Forecast."
            Case "POS00", "POS01", "POS02", "POS03"
                tooltip_msg = "Calculated as the number of months (forward) of Forecasted shipments that the current month's EOM can satisfy."
            Case "DEMPD"
                tooltip_title = "Past Due Forecasts"
                tooltip_msg = "Calculated and carried forward from the previous month only if the EOM on hand from the previous period is 0 (i.e., we were out of stock)." & vbCrLf & "This value may be maintained or zeroed out in the Forecasted Shipments Maintenance screens."
            Case "DEM00"
                tooltip_title = "Current Month Forecast"
                tooltip_msg = "If the MTD Shipped + Open (w/Ship Date <= Cur EOM) is already greater than the Forecast for the Current Month," & vbCrLf & " then the sum of MTD + Open (w/Ship Date <= Cur EOM) is used in place of the Current Month's Forecast" & vbCrLf & " in the calculation of the EOM position for the Current Month, and the Current Month Forecast is shown with a Green background."
            Case "SUPPD"
                tooltip_msg = "Past Due Purchase Order Receipts or Production Orders" & vbCrLf & " - all Open Supply Records with a Due Date of prior to the 1st of the current month."
            Case "MTD"
                tooltip_msg = "MTD Shipments"
            Case "COM"
                tooltip_msg = "Open Sales Orders"
                'THIS SYSTEM IS NOT WORKING FOR TOOLTIPS IN THE MAIN GRID
                'Case "DEM01", "DEM02", "DEM03"
                '    ' e.Row.Cells(DEM).ToolTipText = "Open SO (" & e.Row.Cells(SO).Value & ") is > Forecasted Demand (" & e.Row.Cells(DEM).Value & ")"
                '    tooltip_msg = "Open SO (" & cell.Value & ") is > Forecasted Demand (" & cell.Value & ")"
            Case Else
                tooltip_msg = ""
        End Select
    End Sub

    Private Sub grdDPTPLANX_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdDPTPLANX.MouseLeave
    End Sub

    Private Sub grdDPTPLANX_MouseLeaveElement(ByVal sender As Object, ByVal e As Infragistics.Win.UIElementEventArgs) Handles grdDPTPLANX.MouseLeaveElement
        ' if we are not leaving a cell, then don't anything
        If Not e.Element.GetType().Equals(GetType(UltraWinGrid.CellUIElement)) Then
            Exit Sub
        End If

        '' prevent the timer from ticking again
        'timer.Stop()

        '' destroy the tooltip
        'If Not tooltip Is Nothing Then
        '    tooltip.SetToolTip(Me, String.Empty)
        '    tooltip.Dispose()
        '    tooltip = Nothing
        'End If
    End Sub

    Private Sub grdDPTPLANX_MouseEnterElement(ByVal sender As Object, ByVal e As Infragistics.Win.UIElementEventArgs) Handles grdDPTPLANX.MouseEnterElement
        ' if we are not entering a cell, then don't anything
        If Not e.Element.GetType().Equals(GetType(UltraWinGrid.CellUIElement)) Then
            Exit Sub
        End If

        If ASCMAIN1.DBS_COMPANY = "SLP" Then
            '' find the cell that the cursor is over, if any
            'Dim cell As UltraWinGrid.UltraGridCell = e.Element.GetContext(GetType(UltraWinGrid.UltraGridCell))
            'If Not cell Is Nothing Then
            '    CreateToolTipMessage(cell)
            '    timer.Stop()
            '    timer.Start()
            'End If
        End If

    End Sub

    Private Sub OnTimerTick(ByVal sender As Object, ByVal e As EventArgs)
        tooltip = New System.Windows.Forms.ToolTip()
        tooltip.SetToolTip(grdDPTPLANX, tooltip_msg)
        tooltip.ToolTipTitle = tooltip_title
        tooltip.AutoPopDelay = 12000

        ' once the timer has ticked, stop it
        timer.Stop()
    End Sub


    Sub CreateGraph_Totals()

        Dim chtIsVisible As Boolean = chtTotals.Visible
        chtTotals.Visible = False

        chtTotals.DataSource = Nothing

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Charting Data")

        Me.SuspendLayout()

        Dim RL() As String

        chtTotals.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtTotals.LabelHash = labelHash

        chtTotals.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtTotals.Tooltips.FormatString = "<HIGHLOW>"


        Dim RLi As Integer = 0

        Dim DTY As New DataTable
        With DTY
            .Columns.Add("CODE")
            .Columns.Add("VALUE", GetType(System.Decimal))
        End With


        Dim DTX As DataTable = DirectCast(grdSOTINVHX.DataSource, DataTable)
        ReDim RL(DTX.Rows.Count - 1)
        For Each row As DataRow In DTX.Select("", "CODE_VALUE")
            RL(RLi) = row("CODE_VALUE") ' & ":" & row("DESC_VALUE")
            RLi += 1
            DTY.Rows.Add(New Object() {row.Item("CODE_VALUE"), row.Item("S00")})
        Next
        'chtTotals.Data.SetRowLabels(RL)
        'chtTotals.Data.SetColumnLabels(CL)


        chtTotals.DataSource = DTY
        chtTotals.PieChart.ColumnIndex = -1
        chtTotals.PieChart.OthersCategoryPercent = 2
        'chtTotals.Data.IncludeColumn("CODE_VALUE", False)
        'chtTotals.Data.IncludeColumn("DESC_VALUE", False)
        chtTotals.DataBind()

        'chtTotals.Data.IncludeColumn("S00", True)
        'chtTotals.Data.excludedColumns.Add("S01")


        chtTotals.Visible = chtIsVisible

        Me.ResumeLayout()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()
    End Sub

    Sub CreateGraph_Trend(ByVal Periods As Integer)

        Dim chtIsVisible As Boolean = chtTrend.Visible
        chtTrend.Visible = False

        chtTrend.DataSource = Nothing

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Charting Data")

        Me.SuspendLayout()

        Dim RL() As String
        Dim CL() As String
        ReDim CL(Periods)

        'this will be necessary for line graph
        'For i As Integer = MOSMAX To 0 Step -1
        '    Dim L As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * i))
        '    CL(MOSMAX - i) = Mid(L, 10, 6)
        '    grdSATCSLS1.DisplayLayout.Bands(0).Columns("M" & Format(i, "00")).Header.Caption = Mid(L, 10, 3)
        'Next
        For i As Integer = 1 To Periods
            'Dim L As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * i))
            CL(i - 1) = grdSOTINVHX.DisplayLayout.Bands(0).Columns("S" & Format(i, "00")).Header.Caption
            'grdSATCSLS1.DisplayLayout.Bands(0).Columns("M" & Format(i, "00")).Header.Caption = Mid(L, 10, 3)
        Next

        'chtICTINVAT.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.LabelPlusDataValue
        'chtICTINVAT.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom

        chtTrend.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtTrend.LabelHash = labelHash

        chtTrend.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtTrend.Tooltips.FormatString = "<HIGHLOW>"

        Dim DT As New DataTable
        DT.Columns.Add("CODE_VALUE")
        DT.Columns.Add("DESC_VALUE")
        For P As Integer = 1 To Periods
            DT.Columns.Add("P" & Format(P, "00"), GetType(System.Decimal))
        Next

        Dim RLi As Integer = 0
        Dim DTX As DataTable = DirectCast(grdSOTINVHX.DataSource, DataTable)
        ReDim RL(DTX.Rows.Count - 1)
        For Each row As DataRow In DTX.Select("", "CODE_VALUE")
            RL(RLi) = row("CODE_VALUE") ' & ":" & row("DESC_VALUE")
            RLi += 1

            Dim rowDT As DataRow = DT.NewRow
            rowDT.Item("CODE_VALUE") = row("CODE_VALUE")
            rowDT.Item("DESC_VALUE") = row("DESC_VALUE")
            For P As Integer = 1 To Periods
                rowDT.Item("P" & Format(P, "00")) = row("S" & Format(P, "00"))
            Next
            DT.Rows.Add(rowDT)
        Next
        chtTrend.Data.SetRowLabels(RL)
        chtTrend.Data.SetColumnLabels(CL)

        chtTrend.DataSource = DT
        'chtTrend.Data.IncludeColumn("CODE_VALUE", False)
        'chtTrend.Data.IncludeColumn("DESC_VALUE", False)
        'chtTrend.Data.IncludeColumn("P00", False)

        chtTrend.DataBind()

        chtTrend.Visible = chtIsVisible

        Me.ResumeLayout()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()
    End Sub


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

    Sub Print_Report()

        Dim POS_STATUS_CODES_sel As New List(Of String)
        For Each rowDPTPOSS1 As DataRow In dst.Tables("DPTPOSS1").Select("SEL = '1'")
            Dim POS_STATUS_CODE As String = rowDPTPOSS1.Item("POS_STATUS_CODE")
            POS_STATUS_CODES_sel.Add(POS_STATUS_CODE)
        Next

        Dim MOS As Int32 = numMOS.Value
        For Each rowDPTPLANX As DataRow In dst.Tables("DPTPLANX").Rows
            rowDPTPLANX.Item("SEL") = "0"
            If optCONDITIONS.Value = "A" Then
                rowDPTPLANX.Item("SEL") = "1"
            Else
                If POS_STATUS_CODES_sel.Contains("G") And Val(rowDPTPLANX.Item("POPD") & "") > 0 _
                Or POS_STATUS_CODES_sel.Contains("H") And Val(rowDPTPLANX.Item("ONH") & "") < 0 Then
                    rowDPTPLANX.Item("SEL") = "1"
                Else
                    'Dim ITEM_CODE As String = rowDPTPLANX.Item("ITEM_CODE")
                    'Dim rowDPTPLANO As DataRow = dst.Tables("DPTPLANO").Rows.Find(ITEM_CODE)
                    Dim POS_STATUS_CODES As String = rowDPTPLANX.Item("POS_STATUS_CODES") & ""
                    For P As Int32 = 0 To MOS
                        If POS_STATUS_CODES_sel.Contains(Mid(POS_STATUS_CODES, P + 1, 1)) Then
                            rowDPTPLANX.Item("SEL") = "1"
                            Exit For
                        End If
                    Next
                End If
            End If
        Next


        Call Print_Report_Begin()

        For P = 3 To 1 Step -1
            CR_params.Add("PP" & Format(P, "00"), Format(YPP(P, 1)))
        Next

        For P = 0 To 9
            CR_params.Add("PF" & Format(P, "00"), Format(YPF(P, 1)))
        Next
        CR_params.Add("PF25", Format(YPF(25, 1)))
        Generate_Report("DPRPLAN1", "Demand Planning", grdDPTPLANX.Text)

        Call Print_Report_End()

    End Sub

    Private Sub optUD_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optUD.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If grdGLTPARM2.Selected.Rows.Count > 0 Then
            cmdFetchSales.PerformClick()
        End If
    End Sub

    Sub Generate_Plans()
        For Each rowDPTPLANX As DataRow In dst.Tables("DPTPLANX").Rows
            Dim ITEM_POS_MIN As Decimal = Val(Absx1.numFor("ITEM_POS_MIN").Value & "")

            For P = 0 To 25
                Dim POS As Decimal = Val(rowDPTPLANX.Item("POS" & Format(P, "00")) & "")
                If POS < ITEM_POS_MIN Then ' WE HAVE A SHORTFALL
                    If 1 <> 1 Then
                        Dim rowDPTPLAN1 As DataRow = dst.Tables("DPTPLAN1").NewRow
                        With rowDPTPLAN1
                            Dim PLAN_NO As String = ASCMAIN1.Next_Control_No("DPTPLAN1.PLAN_NO")
                            .Item("PLAN_NO") = PLAN_NO
                            .Item("ITEM_CODE") = rowDPTPLANX.Item("ITEM_CODE")
                            .Item("DATE_ENTERED") = Now.Date
                            .Item("VEND_CODE") = PLAN_NO
                            .Item("PLAN_MB") = PLAN_NO
                            .Item("AT_WHSE") = PLAN_NO
                            .Item("TO_WHSE") = PLAN_NO
                            .Item("DATE_REQUIRED") = PLAN_NO
                            .Item("DATE_COMPSDUE") = PLAN_NO
                            .Item("QTY_PLANNED") = PLAN_NO
                            .Item("DATE_DELETE") = PLAN_NO
                            .Item("DATE_PLAN_ACTION") = PLAN_NO
                            .Item("DATE_PO_ISSUE") = PLAN_NO
                            .Item("PLAN_TYPE") = PLAN_NO
                            .Item("BM_ISSUE_NO") = PLAN_NO
                            .Item("ACT_MSG_FLAG") = PLAN_NO
                            .Item("ACT_MSG_DATE") = PLAN_NO
                            .Item("DATE_VEND_SHIP") = PLAN_NO
                        End With
                        dst.Tables("DPTPLAN1").Rows.Add(rowDPTPLAN1)

                    End If
                End If
            Next
        Next
    End Sub

    Private Sub tabDetails_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabDetails.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabDetails()
    End Sub

    Sub Setup_tabDetails()
        chkEditForecasts.Visible = (chkItemDetails.Checked) And (tabDetails.SelectedTab.Key = "Forecasts") And Not InquiryMode
        UltraExplorerBar1.Groups("Item Master").Visible = (tabDetails.SelectedTab.Key = "Item Master") And Not InquiryMode
        UltraExplorerBar1.Groups("Plans").Visible = (tabDetails.SelectedTab.Key = "Plans") And Not InquiryMode And ASCMAIN1.DBS_COMPANY <> "INT"

        UltraExplorerBar1.Groups("Display Options").Visible = Not ((tabDetails.SelectedTab.Key = "Item Master") And (tabDetails.SelectedTab.Key = "Plans"))
        UltraExplorerBar1.Groups("Report Options").Visible = Not ((tabDetails.SelectedTab.Key = "Item Master") And (tabDetails.SelectedTab.Key = "Plans"))
    End Sub

    Private Sub grdLegend_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdDPTPOSS1.InitializeRow

        Dim C As String = "POS_STATUS_DESC"
        If Val(e.Row.Cells("POS_RBG_BACKCOLOR").Value & "") = 0 Then
            e.Row.Cells(C).Appearance.BackColor = Color.Empty
        Else
            e.Row.Cells(C).Appearance.BackColor = Color.FromArgb(Val(e.Row.Cells("POS_RBG_BACKCOLOR").Value & ""))
        End If
        If Val(e.Row.Cells("POS_RBG_FORECOLOR").Value & "") = 0 Then
            e.Row.Cells(C).Appearance.ForeColor = Color.Empty
        Else
            e.Row.Cells(C).Appearance.ForeColor = Color.FromArgb(Val(e.Row.Cells("POS_RBG_FORECOLOR").Value & ""))
        End If
    End Sub

    Sub Setup_grdDPTPOSS1()
        Fill_Records("DPTPOSS1")

        With grdDPTPOSS1.DisplayLayout.Bands(0)
            For Each GC As UltraWinGrid.UltraGridColumn In .Columns
                GC.Hidden = (GC.Key <> "POS_STATUS_DESC" And GC.Key <> "SEL")
            Next
            '.Columns("SEL").Header.VisiblePosition = 0
            '.Columns("SEL").Style = UltraWinGrid.ColumnStyle.CheckBox
            .ColHeadersVisible = False
            .Columns("SEL").Width = 20
            .Columns("SEL").Header.Caption = "Sel"
            .Columns("SEL").CellActivation = UltraWinGrid.Activation.AllowEdit

            .Columns("POS_STATUS_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
            .Columns("POS_STATUS_DESC").Width = 100
        End With
        grdDPTPOSS1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        'grdDPTPOSS1.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText

        grdDPTPOSS1.DisplayLayout.Bands(0).Columns("SEL").CellClickAction = UltraWinGrid.CellClickAction.CellSelect

        For Each rowDPTPOSS1 As DataRow In dst.Tables("DPTPOSS1").Rows
            Dim POS_STATUS_CODE As String = rowDPTPOSS1.Item("POS_STATUS_CODE")
            If InStr("ABCDEF", POS_STATUS_CODE) <> 0 Then
                rowDPTPOSS1.Item("SEL") = "1"
            Else
                rowDPTPOSS1.Item("SEL") = "0"
            End If
        Next

        Sort_grdColumns(grdDPTPOSS1, "POS_STATUS_CODE")
    End Sub

    Private Sub chkHideTree_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkHideTree.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        splMain.Panel1Collapsed = chkHideTree.Checked
    End Sub

    Private Sub chkItemDetails_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkItemDetails.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        ItemDetails()
    End Sub

    Overrides Function Excel_Export(ByVal grd As UltraWinGrid.UltraGrid) As GemBox.Spreadsheet.ExcelFile
        If grd.Name = "grdDPTPLANXX" Then

            grdDPTPLANY.Text = grdDPTPLANX.Text
            If dst.Tables.Contains("DPTPLANY") Then
                grdDPTPLANY.DataSource = Nothing
                dst.Tables.Remove("DPTPLANY")
            End If
            Dim CL() As Int32 = Nothing
            Dim CW() As Int32 = Nothing
            Dim CG() As Int32 = Nothing
            Dim CF() As String = Nothing
            Dim CA() As Appearance = Nothing

            Dim RL() As String = Nothing
            ReDim RL(grd.DisplayLayout.Bands(0).LevelCount - 1)
            With dst.Tables.Add("DPTPLANY")
                .Columns.Add("ROWNUM", GetType(System.Int32))

                For Each G As UltraWinGrid.UltraGridGroup In grd.DisplayLayout.Bands(0).Groups
                    ReDim CL(grd.DisplayLayout.Bands(0).LevelCount - 1)
                    Dim CLmax As Int32 = 0
                    Dim dt() As System.Type = Nothing
                    For Each C As UltraWinGrid.UltraGridColumn In G.Columns
                        If Not C.Hidden Then
                            CL(C.Level) += 1
                            If CL(C.Level) > CLmax Then
                                CLmax = CL(C.Level)
                                If CW Is Nothing Then
                                    ReDim Preserve CW(0)
                                Else
                                    ReDim Preserve CW(CW.Length)
                                End If
                                CW(CW.Length - 1) = C.Width
                                If CG Is Nothing Then
                                    ReDim Preserve CG(0)
                                Else
                                    ReDim Preserve CG(CG.Length)
                                End If
                                CG(CG.Length - 1) = G.Index
                                If CF Is Nothing Then
                                    ReDim Preserve CF(0)
                                Else
                                    ReDim Preserve CF(CF.Length)
                                End If
                                CF(CF.Length - 1) = C.Format
                                If CA Is Nothing Then
                                    ReDim Preserve CA(0)
                                Else
                                    ReDim Preserve CA(CA.Length)
                                End If
                                CA(CA.Length - 1) = C.CellAppearance
                            End If
                            If dt Is Nothing OrElse dt.Length < CLmax Then
                                ReDim Preserve dt(CLmax - 1)
                                dt(CLmax - 1) = C.DataType ' dst.Tables("DPTPLANX").Columns(C.Key).DataType
                            End If

                            If C.DataType.ToString <> dt(CL(C.Level) - 1).ToString Then
                                If dt(CL(C.Level) - 1).ToString <> "System.String" Then
                                    If C.DataType.ToString = "System.String" Then
                                        dt(CL(C.Level) - 1) = C.DataType
                                    ElseIf C.DataType.ToString = "System.Decimal" And dt(CL(C.Level) - 1).ToString Like "System.Int*" Then
                                        dt(CL(C.Level) - 1) = C.DataType
                                    End If
                                End If
                                'dt(CL(C.Level) - 1) = GetType(System.String)
                            End If

                            Dim COLUMN_NAME As String = "G" & CStr(G.Index) & "_" & CStr(CL(C.Level))
                            'If C.IsBound Then
                            RL(C.Level) &= "," & C.Key & ":" & COLUMN_NAME

                            'End If

                        End If
                    Next
                    For i As Int32 = 1 To CLmax
                        Dim COLUMN_NAME As String = "G" & CStr(G.Index) & "_" & CStr(i)
                        With .Columns.Add(COLUMN_NAME, dt(i - 1))

                        End With
                    Next
                Next

                Dim ROWNUM As Int32 = -1
                grdDPTPLANY.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
                grdDPTPLANY.DataSource = dst.Tables("DPTPLANY")
                grdDPTPLANY.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes
                grdDPTPLANY.DisplayLayout.Override.RowSizing = UltraWinGrid.RowSizing.Fixed

                With grdDPTPLANY.DisplayLayout.Bands(0)
                    Dim LAST_G As Int32 = -1
                    For I As Int32 = 1 To .Columns.Count - 1
                        .Columns(I).Width = CW(I - 1)
                        .Columns(I).Format = CF(I - 1)
                        .Columns(I).CellAppearance = CA(I - 1)
                        If LAST_G <> CG(I - 1) Then
                            LAST_G = CG(I - 1)
                            .Columns(I).Header.Caption = grdDPTPLANX.DisplayLayout.Bands(0).Groups(CG(I - 1)).Header.Caption
                        Else
                            .Columns(I).Header.Caption = ""
                        End If
                        .Columns(I).Header.Appearance = grdDPTPLANX.DisplayLayout.Bands(0).Groups(CG(I - 1)).Header.Appearance
                    Next
                End With

                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    Dim grow2 As UltraWinGrid.UltraGridRow
                    For R As Int32 = 0 To RL.Length - 1
                        grdDPTPLANY.DataSource = dst.Tables("DPTPLANY")
                        grow2 = grdDPTPLANY.DisplayLayout.Bands(0).AddNew

                        For Each CXY As String In Split(Mid(RL(R), 2), ",")
                            Dim COLUMN_NAME_X As String = Split(CXY, ":")(0)
                            Dim COLUMN_NAME_Y As String = Split(CXY, ":")(1)
                            'If COLUMN_NAME_X = "TOTAL_POS" Then
                            'Else
                            If grow.Cells(COLUMN_NAME_X).Column.DataType.ToString = "System.String" Then
                                grow2.Cells(COLUMN_NAME_Y).Value = grow.Cells(COLUMN_NAME_X).Value & ""
                            Else
                                grow2.Cells(COLUMN_NAME_Y).Value = Val(grow.Cells(COLUMN_NAME_X).Value & "")
                            End If
                            grow2.Cells(COLUMN_NAME_Y).Appearance = grow.Cells(COLUMN_NAME_X).Appearance
                            'End If
                        Next

                        ROWNUM += 1
                        grow2.Cells("ROWNUM").Value = ROWNUM

                        grow2.Update()
                    Next
                    'grow2 = grdDPTPLANY.DisplayLayout.Bands(0).AddNew
                    'ROWNUM += 1
                    'grow2.Cells("ROWNUM").Value = ROWNUM
                    'grow2.Update()
                    'grow2.Height = grow2.Height / 5
                Next

                Sort_grdColumns(grdDPTPLANY, "ROWNUM")

                grdDPTPLANY.DisplayLayout.Bands(0).Columns("ROWNUM").Hidden = True

                'For Each row As DataRow In dst.Tables("DPTPLANX").Rows
                '    For R As Int32 = 0 To RL.Length - 1
                '        Dim row2 As DataRow = dst.Tables("DPTPLANY").NewRow
                '        For Each CXY As String In Split(Mid(RL(R), 2), ",")
                '            Dim COLUMN_NAME_X As String = Split(CXY, ":")(0)
                '            Dim COLUMN_NAME_Y As String = Split(CXY, ":")(1)
                '            If COLUMN_NAME_X = "TOTAL_POS" Then
                '            Else
                '                row2.Item(COLUMN_NAME_Y) = row.Item(COLUMN_NAME_X)
                '            End If
                '        Next
                '        dst.Tables("DPTPLANY").Rows.Add(row2)
                '    Next
                'Next
            End With

            'grdDPTPLANY.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
            'grdDPTPLANY.DataSource = dst.Tables("DPTPLANY")
            tabDetails.Tabs("grd").Visible = True
            MyBase.Excel_Export(grdDPTPLANY)
            grdDPTPLANY.DataSource = Nothing
            tabDetails.Tabs("grd").Visible = False
        Else
            '  grd.DisplayLayout.Bands(0).ColHeadersVisible = False
            MyBase.Excel_Export(grd)
        End If
        Return Nothing
    End Function

    Private Sub chkShowOverShort_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowOverShort.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_OverShort()
    End Sub

    Sub Setup_OverShort()
        With grdDPTPLANX.DisplayLayout.Bands(0)
            If chkShowOverShort.Checked Then
                .LevelCount = 9
                For p As Int32 = -1 To 25
                    Dim PX As String = IIf(p = -2, "", IIf(p = -1, "PD", Format(p, "00")))
                    .Columns("OS" & PX).Level = 4
                    .Columns("OS" & PX).Hidden = False
                    '.Columns("GP" & PX).Level = 5
                    '.Columns("GP" & PX).Hidden = False
                    .Columns("SO" & PX).Level = 6
                    .Columns("SO" & PX).Hidden = False
                    .Columns("SI" & PX).Level = 7
                    .Columns("SI" & PX).Hidden = False
                    .Columns("ST" & PX).Level = 8
                    .Columns("ST" & PX).Hidden = False
                Next
                .Columns("DATA_OS").Level = 4
                .Columns("DATA_OS").Hidden = False
                '.Columns("DATA_GP").Level = 5
                '.Columns("DATA_GP").Hidden = False
                .Columns("DATA_SO").Level = 6
                .Columns("DATA_SO").Hidden = False
                .Columns("DATA_SI").Level = 7
                .Columns("DATA_SI").Hidden = False
                .Columns("DATA_ST").Level = 8
                .Columns("DATA_ST").Hidden = False

                .Columns("TOTAL_OS").Level = 4
                .Columns("TOTAL_OS").Hidden = False
                '.Columns("TOTAL_GP").Level = 5
                '.Columns("TOTAL_GP").Hidden = False
                .Columns("TOTAL_SO").Level = 6
                .Columns("TOTAL_SO").Hidden = False
                .Columns("TOTAL_SI").Level = 7
                .Columns("TOTAL_SI").Hidden = False
                .Columns("TOTAL_ST").Level = 8
                .Columns("TOTAL_ST").Hidden = False
            Else
                .Columns("DATA_OS").Level = 3
                .Columns("DATA_OS").Hidden = True
                '.Columns("DATA_GP").Level = 3
                '.Columns("DATA_GP").Hidden = True
                .Columns("DATA_SO").Level = 3
                .Columns("DATA_SO").Hidden = True
                .Columns("DATA_SI").Level = 3
                .Columns("DATA_SI").Hidden = True
                .Columns("DATA_ST").Level = 3
                .Columns("DATA_ST").Hidden = True
                .Columns("TOTAL_OS").Level = 3
                .Columns("TOTAL_OS").Hidden = True
                '.Columns("TOTAL_GP").Level = 3
                '.Columns("TOTAL_GP").Hidden = True
                .Columns("TOTAL_SO").Level = 3
                .Columns("TOTAL_SO").Hidden = True
                .Columns("TOTAL_SI").Level = 3
                .Columns("TOTAL_SI").Hidden = True
                .Columns("TOTAL_ST").Level = 3
                .Columns("TOTAL_ST").Hidden = True
                For p As Int32 = -1 To 25
                    Dim PX As String = IIf(p = -2, "", IIf(p = -1, "PD", Format(p, "00")))
                    .Columns("OS" & PX).Level = 3
                    .Columns("OS" & PX).Hidden = True
                    '.Columns("GP" & PX).Level = 3
                    '.Columns("GP" & PX).Hidden = True
                    .Columns("SO" & PX).Level = 3
                    .Columns("SO" & PX).Hidden = True
                    .Columns("SI" & PX).Level = 3
                    .Columns("SI" & PX).Hidden = True
                    .Columns("ST" & PX).Level = 3
                    .Columns("ST" & PX).Hidden = True
                Next
                .LevelCount = 6
            End If
        End With
    End Sub

    Private Sub grdDPTPLANX_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdDPTPLANX.ClickCellButton
        Dim ITEM_CODE As String = e.Cell.Value & ""
        Dim FOLDERNAME As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
        'If ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_SERVER = "" Then FOLDERNAME = "C:\Documents and Settings\wjz\Desktop\JHI\Images\"
        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
        'Dim FILENAME As String = e.Cell.Row.Cells("ITEM_PICTURE_FILENAME").Text ' "AB301.jpg"
        Dim FILENAME As String = rowICTITEM1.Item("ITEM_PICTURE_FILENAME") & ""
        PictureBox1.Image = ASCMAIN1.Get_Image(FOLDERNAME, FILENAME)

        With UltraExplorerBar1
            '.Groups("Forecast Status").Expanded = False
            '.Groups("Prorate Sales Using").Expanded = False
            '.Groups("Display Options").Expanded = False
            .Groups("Picture").Expanded = True
            .Groups("Picture").Text = "Item " & ITEM_CODE
        End With

        'Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
        'Dim FILENAME As String = "C:\Documents and Settings\wjz\Desktop\JHI\Images\" & ITEM_CODE & ".JPG"
        'If My.Computer.FileSystem.FileExists(FOLDERNAME & "\" & FILENAME) Then
        '    Me.Cursor = Cursors.WaitCursor
        '    Call ASCMAIN1.Progress("Now Loading Image Viewer")
        '    System.Diagnostics.Process.Start(FOLDERNAME & "\" & FILENAME)
        '    Me.Cursor = Cursors.Default
        '    Call ASCMAIN1.Progress("")
        'End If
    End Sub

    Private Sub optCONDITIONS_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCONDITIONS.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_Conditions()
    End Sub

    Sub Setup_Conditions()
        lblMOS.Visible = (optCONDITIONS.Value = "S")
        numMOS.Visible = (optCONDITIONS.Value = "S")
        grdDPTPOSS1.DisplayLayout.Bands(0).Columns("SEL").Hidden = Not (optCONDITIONS.Value = "S")
    End Sub

    Sub Find_Item()
        Dim ITEM_CODE As String = txtFindITEM_CODE.Text

        If ITEM_CODE <> "" Then
            Dim rowICTITEMY As DataRow = LookUp(ICTITEMY, ITEM_CODE)
            If rowICTITEMY IsNot Nothing Then
                If tabTG.SelectedTab.Key = "Tree" Then
                    Dim KEY As String = ""
                    For Each COLUMN_NAME As String In COLUMN_NAMEs
                        KEY &= "/" & rowICTITEMY.Item(COLUMN_NAME)
                    Next
                    Dim tnode As UltraWinTree.UltraTreeNode = tvwDQ.GetNodeByKey(Mid(KEY, 2))
                    If tnode IsNot Nothing Then
                        tvwDQ.ActiveNode = tnode
                        Click_Node(tvwDQ.ActiveNode)
                        For Each grow As UltraWinGrid.UltraGridRow In grdDPTPLANX.Rows
                            If grow.Cells("ITEM_CODE").Text = ITEM_CODE Then
                                grdDPTPLANX.ActiveRow = grow
                                chkItemDetails.Checked = True
                                grdDPTPLANX.ActiveRowScrollRegion.FirstRow = grow
                                '  grdDPTPLANX.ActiveRowScrollRegion.ScrollRowIntoView(grow)
                                Exit Sub
                            End If
                        Next
                        txtFindITEM_CODE.Text = ""
                    End If
                Else
                    Click_Node(Nothing)
                    txtFindITEM_CODE.Text = ""
                End If
            End If
        End If
    End Sub

    Sub Set_View_grdDPTPLANX()

        Dim sql As String = ""
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Items with Forecasts Only"), UltraWinToolbars.StateButtonTool)
        If tlb_sbt.Checked Then
            sql = "ISNULL(TOTAL_DEM,0)<>0"
        Else
            sql = ""
        End If
        Dim dvw As DataView = DirectCast(grdDPTPLANX.DataSource, DataTable).DefaultView
        dvw.RowFilter = sql

    End Sub

    Private Sub grdICTSTATX_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdICTSTATX.InitializeLayout

    End Sub

    Private Sub grdICTSTATX_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTSTATX.InitializeRow
        'Dim rowICTWHSE1 As DataRow = dst.Tables("ICTWHSE1").Rows.Find(e.Row.Cells("WHSE_CODE").Value & "")
        'If rowICTWHSE1 IsNot Nothing Then
        '    If rowICTWHSE1.Item("WHSE_MRP_EXC_IND") & "" = "1" Then
        '        e.Row.Appearance.BackColor = Color.Pink
        '    End If
        'End If
        If e.Row.Cells("WHSE_MRP_EXC_IND").Value & "" = "1" Then
            e.Row.Appearance.BackColor = Color.Pink
        End If
        If e.Row.Cells("WHSE_TYPE").Value & "" = "R" Then
            e.Row.Appearance.ForeColor = Color.Red
        End If

    End Sub

    Private Sub grdPOTORDRX_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTORDRX.InitializeRow
        ' Debug.Print(e.Row.Index & ":" & e.Row.Cells.Contains("CONTAINER_NO") & ":" & e.Row.Cells(7).Column.Key)
        If e.Row.Band.Index = 1 Then ' e.Row.Cells(7).Column.Key = "CONTAINER_NO" Then ' e.Row.Cells.Contains("CONTAINER_NO") Then ' e.Row.Index = 1 Then
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
        End If
    End Sub

    Private Sub grdSOTORDRX_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDRX.InitializeRow
        Dim ORDR_SHIP_DATE As Date = CDate(e.Row.Cells("ORDR_SHIP_DATE").Value)
        If Format(ORDR_SHIP_DATE, "yyyyMMdd") <= Format(YPPD(1), "yyyyMMdd") Then
            e.Row.Appearance.BackColor = Color.Pink
        End If
    End Sub

    Private Sub grdDPTMUPD1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdDPTMUPD1.DoubleClickRow
        If Not ScreenMode And e.Row.IsDataRow Then
            Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Value
            Absx1.txtFor("BRAND_CODE").Text = ""
            Absx1.txtFor("ITEM_CODE").Text = ITEM_CODE
            Click_Command("Load")
        End If
    End Sub

    Private Sub grdDPTMUPD2_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdDPTMUPD2.DoubleClickRow
        If Not ScreenMode And e.Row.IsDataRow Then
            Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Value
            Absx1.txtFor("BRAND_CODE").Text = ""
            Absx1.txtFor("ITEM_CODE").Text = ITEM_CODE
            Click_Command("Load")
        End If
    End Sub

    Private Sub txtFindITEM_CODE_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtFindITEM_CODE.KeyDown
        Find_Item()
    End Sub

    Private Sub txtFindITEM_CODE_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtFindITEM_CODE.ValueChanged

    End Sub

    Private Sub optBP_ValueChanged(sender As Object, e As EventArgs) Handles optBP.ValueChanged
        If Me.SELECTION_NO = 0 Or Not ScreenMode Then Exit Sub
        '  Click_Node(Nothing)
    End Sub

    Private Sub optSNU_ValueChanged(sender As Object, e As EventArgs) Handles optSNU.ValueChanged
        If Me.SELECTION_NO = 0 Or Not ScreenMode Then Exit Sub
        '  Click_Node(Nothing)
    End Sub

    Private Sub grdICTCOLL1_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdICTCOLLX.AfterRowUpdate
        If Me.SELECTION_NO = 0 Or Not ScreenMode Then Exit Sub
        ' Click_Node(Nothing)
    End Sub

    Sub Set_BI()
        lblITEM_CODE.Visible = ScreenMode And optBy.Value = "I"
        txtITEM_CODE.Visible = optBy.Value = "I"
        txtITEM_DESC.Visible = optBy.Value = "I"
        lblBRAND_CODE.Visible = ScreenMode And optBy.Value = "B"
        txtBRAND_CODE.Visible = optBy.Value = "B"
        txtBRAND_NAME.Visible = optBy.Value = "B"

        optBy.Visible = Not ScreenMode

        If optBy.Value = "B" Then
            txtBRAND_CODE.Focus()
            '    splDPTPLANX.SplitterDistance = 450

            splDPTPLANX.SplitterDistance = splDPTPLANX.Height * 3 / 5
            chkItemDetails.Checked = False
            If tabDetails.SelectedTab IsNot Nothing Then
                tabDetails.SelectedTab = tabDetails.Tabs("Forecasts")
            End If
        Else '
            txtITEM_CODE.Focus()
            ' splDPTPLANX.SplitterDistance = 250
            splDPTPLANX.SplitterDistance = splDPTPLANX.Height * 2 / 5
        End If
    End Sub
    Private Sub optBy_ValueChanged(sender As Object, e As EventArgs) Handles optBy.ValueChanged
        Set_BI()
    End Sub

    Private Sub tabTG_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabTG.SelectedTabChanged
        chkAllInNode.Visible = (tabTG.SelectedTab.Key = "Tree")
        If tabTG.SelectedTab.Key = "Tree" Then
            chkHideTree.Text = "Hide Tree"
        Else
            chkHideTree.Text = "Hide Collections"
        End If
    End Sub

    Private Sub btnGetItems_Click(sender As Object, e As EventArgs) Handles btnGetItems.Click
        Click_Node(Nothing)
    End Sub

    Private Sub grdDPTITMFX_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdDPTITMFX.InitializeRow

        If e.Row Is Nothing OrElse e.Row.IsFilterRow OrElse Not e.Row.IsDataRow Then
            Exit Sub
        End If

        Dim MARKET_CODE As String = e.Row.Cells("MARKET_CODE").Value & ""

        For P = 0 To 25
            Dim sqlw As String = "ITEM_CODE = '" & ITEM_CODE & "' and MARKET_CODE = '" & MARKET_CODE & "' and OPS_YYYYPP_FC = '" & YPF(P, 0) & "' and STATUS IS NULL"
            Dim rows() As DataRow = dst.Tables("DPTITMF2").Select(sqlw, "INIT_DATE")
            If rows.Length <> 0 Then
                Dim TT As String = ""
                For Each row As DataRow In rows
                    TT &= Format(row.Item("INIT_DATE"), "MM/dd/yyyy") & vbTab & Format(Val(row.Item("FORECAST") & ""), "#,##0") & vbTab & row.Item("FORECAST_NOTE")
                    TT &= vbCrLf
                Next
                With e.Row.Cells("FC" & Format(P, "00"))
                    .Appearance = cellHasNotes
                    .ToolTipText = TT
                End With
            Else
                '  e.Row.Cells("").Appearance = Nothing
            End If
        Next

    End Sub

    Private Sub grdDPTITMFX_DoubleClickCell(sender As Object, e As DoubleClickCellEventArgs) Handles grdDPTITMFX.DoubleClickCell
        If e.Cell Is Nothing OrElse e.Cell.Row Is Nothing OrElse e.Cell.Row.IsFilterRow OrElse Not e.Cell.Row.IsDataRow Then
            Exit Sub
        End If

        Dim COLUMN_NAME As String = e.Cell.Column.Key
        Dim P As Integer = 0
        If COLUMN_NAME.Length = 4 AndAlso COLUMN_NAME.StartsWith("FC") And COLUMN_NAME <> "FCPD" Then
            P = Val(Mid(COLUMN_NAME, 3))
        Else
            Exit Sub
        End If

        Dim MARKET_CODE As String = e.Cell.Row.Cells("MARKET_CODE").Value & ""
        Dim FORECAST As Int64 = Val(e.Cell.Value & "")

        Using F As New TAC.DPFITMF2
            F.ITEM_CODE = ITEM_CODE
            F.MARKET_CODE = MARKET_CODE
            F.OPS_YYYYPP_FC = YPF(P, 0)
            F.FORECAST = FORECAST
            F.allow_new_notes = True

            F.ShowDialog()

            If F.update_was_clicked Then
                Fill_Records("DPTITMF2")
                grdDPTITMFX.Rows.Refresh(RefreshRow.FireInitializeRow)
            End If

        End Using

    End Sub

    Private Sub grdDPTITMFS_DoubleClickCell(sender As Object, e As DoubleClickCellEventArgs) Handles grdDPTITMFS.DoubleClickCell
        If e.Cell Is Nothing OrElse e.Cell.Row Is Nothing OrElse e.Cell.Row.IsFilterRow OrElse Not e.Cell.Row.IsDataRow Then
            Exit Sub
        End If

        Dim COLUMN_NAME As String = e.Cell.Column.Key
        Dim P As Integer = 0
        If COLUMN_NAME.Length = 3 AndAlso COLUMN_NAME.StartsWith("F") And COLUMN_NAME <> "FPD" Then
            P = Val(Mid(COLUMN_NAME, 2))
        Else
            Exit Sub
        End If

        Dim MARKET_CODE As String = e.Cell.Row.Cells("MARKET_CODE").Value & ""

        If MARKET_CODE = "*" Then
            Exit Sub
        End If


        If e.Cell.ToolTipText = "" Then
            MsgBox("No Notes for this Month", MsgBoxStyle.OkOnly, "Cannot Show Notes")
            Exit Sub
        End If


        'Dim FORECAST As Int64 = Val(e.Cell.Value & "")

        Using F As New TAC.DPFITMF2
            F.ITEM_CODE = ITEM_CODE
            F.MARKET_CODE = MARKET_CODE
            F.OPS_YYYYPP_FC = YPP(P, 0)
            'F.FORECAST = FORECAST
            'F.allow_new_notes = False

            F.ShowDialog()

            'If F.update_was_clicked Then
            '    Fill_Records("DPTITMF2")
            '    grdDPTITMFX.Rows.Refresh(RefreshRow.FireInitializeRow)
            'End If

        End Using

    End Sub

    Sub Set_DC_Date(ITEM_CODE As String)

        ASCDATA1.DeleteRows("ICTPINVX", "INV_NUM = 'Not Inv'")
        dst.Tables("ICTPINVX").AcceptChanges()

        For Each row As DataRow In dst.Tables("POTORDRX").Select("") ' $"ITEM_CODE = '{ITEM_CODE}'")
            Dim INV_QTY As Integer = Val(dst.Tables("ICTPINVX").Compute("SUM(PINV_QTY)", $"PO_ORDER_NO = '{row.Item("PO_ORDER_NO")}' AND PO_ORDER_LNO = {row.Item("PO_ORDER_LNO")}") & "")
            Dim PO_QTY_OPN As Integer = Val(row.Item("PO_QTY_OPN") & "")
            If PO_QTY_OPN <> INV_QTY Then
                Dim rowICTPINVX As DataRow = dst.Tables("ICTPINVX").NewRow
                With rowICTPINVX
                    .Item("PINV_NO") = "000000"
                    '.Item("PINV_LNO") = 1
                    .Item("PO_ORDER_NO") = row.Item("PO_ORDER_NO")
                    .Item("PO_ORDER_LNO") = row.Item("PO_ORDER_LNO")
                    .Item("CONTAINER_NO") = "Qty Open"
                    .Item("INV_NUM") = "Not Inv"
                    .Item("ITEM_CODE") = ITEM_CODE
                    .Item("ETA_DATE") = row.Item("PO_DATE_REQUIRED")
                    .Item("OPO_QTY") = PO_QTY_OPN - INV_QTY
                    .Item("PINV_QTY") = PO_QTY_OPN - INV_QTY
                    '.Item("WHSE_CODE") = row.Item("WHSE_CODE")
                End With
                dst.Tables("ICTPINVX").Rows.Add(rowICTPINVX)
            End If
        Next

        For Each row As DataRow In dst.Tables("ICTPINVX").Select($"ITEM_CODE = '{ITEM_CODE}' AND CONTAINER_NO IS NULL")
            With row
                Dim WHSE_CODE As String = row.Item("WHSE_CODE") & ""
                ' ISSUE-7230 Clarins to ADS
                'If ASCMAIN1.CLIENT = "INT" Then
                '    If WHSE_CODE <> "CLA" Or WHSE_CODE <> "ADS" Then
                '        WHSE_CODE = "CLA"
                '    End If
                'End If

                ' DPTPARM1 DP_PARM_DEF_PLAN_WHSE
                If WHSE_CODE.Length = 0 Then
                    WHSE_CODE = ROWs("DPTPARM1").Item("DP_PARM_DEF_PLAN_WHSE") & String.Empty
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

        For Each ROW As DataRow In dst.Tables("ICTPINVX").Select($"ITEM_CODE = '{ITEM_CODE}'")
            If ROW.Item("ETA_DATE") & "" <> "" Then
                Dim ETA_DATE As Date = ROW.Item("ETA_DATE")
                Dim ETA_DATE_DC As Date = ETA_DATE
                ' ISSUE-6931 9/18/25 - if not invoiced, ETA date is ETA to DC, so do not add all this stuff
                If ROW.Item("INV_NUM" & "") <> "Not Inv" Then
                    For I As Integer = 1 To 5
                        ETA_DATE_DC = ETA_DATE_DC.AddDays(1)
                        If ETA_DATE_DC.DayOfWeek = DayOfWeek.Saturday Or ETA_DATE_DC.DayOfWeek = DayOfWeek.Sunday Then
                            I = I - 1
                        End If
                    Next
                End If
                ' NEXT LINE TO APPROXIMATE ADDING 5 WEEK_DAYS LIKE WE DO IN MRP UPDATE USING AN ORACLE FUNCTION - BUT THIS IS BEING DONE ABOVE
                ROW.Item("ETA_DATE_DC") = ETA_DATE_DC ' .AddDays(7)
                End If
                ROW.AcceptChanges()
            ROW.SetAdded()
        Next

        'ASCDATA1.ExecuteSQL($"Delete from {ICTPINVX} where ITEM_CODE = '{ITEM_CODE}' and CONTAINER_NO is Null")
        ASCDATA1.ExecuteSQL($"Delete from {ICTPINVX} where ITEM_CODE = '{ITEM_CODE}'")
        'Update_Record_TDA("ICTPINVX", $"ITEM_CODE = '{ITEM_CODE}'")
        Update_Record_TDA("ICTPINVX")
    End Sub

    Sub Create_Worktables(initialize As Boolean, Optional ITEM_CODE As String = "", Optional truncate As Boolean = False)

        If initialize Then

            sqlICTPINVX = "Select ICTPINV2.PO_ORDER_NO, ICTPINV2.PO_ORDER_LNO, ICTPINV2.PINV_NO, ICTPINV2.PINV_LNO
                , ICTPINV2.ITEM_CODE, ICTPINV1.INV_NUM, ICTPINV1.INV_DATE, ICTPINV1.WHSE_CODE
                , ICTPINV1.VESSEL_NAME, ICTPINV1.BOL_NO, ICTPINV1.CONTAINER_NO, ICTPINV1.SHIP_DATE
                , ICTPINV1.ETA_DATE, ICTPINV2.PINV_QTY, ICTPINV1.ETA_DATE ETA_DATE_DC, ICTITEM1.PORT_CODE
                 from ICTPINV1,ICTPINV2, ICTITEM1
                 where ICTPINV1.PO_ORDER_NO In (
                Select POTORDR2.PO_ORDER_NO
                             from POTORDR2,POTORDR1
                             where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO
                               And POTORDR2.PO_QTY_OPN <> 0
                               And POTORDR2.ITEM_CODE = :PARM1
                ) And ICTPINV1.PINV_STATUS = 'O' 
And ICTITEM1.ITEM_CODE = ICTPINV2.ITEM_CODE
AND ICTPINV1.PINV_NO = ICTPINV2.PINV_NO AND ICTPINV2.ITEM_CODE = :PARM1"

            ICTPINVX = ASCMAIN1.Temp_Table(Replace(sqlICTPINVX, ":PARM1", "''"))
            'ASCDATA1.ExecuteSQL($"Alter Table {ICTPINVX} Add Primary Key (PO_ORDER_NO, PO_ORDER_LNO, PINV_NO, PINV_LNO)")
            ASCDATA1.ExecuteSQL($"Create Index I_{ICTPINVX}_1 on {ICTPINVX} (ITEM_CODE)")
            'ASCDATA1.ExecuteSQL(sqlICTPINVX, "V", New String() {""})
        Else

            If truncate Then
                ASCMAIN1.sql = $"Truncate Table {ICTPINVX}"
                ASCDATA1.ExecuteSQL()
            Else
                ASCMAIN1.sql = $"Delete from {ICTPINVX} where ITEM_CODE = '{ITEM_CODE}'"
                ASCDATA1.ExecuteSQL()
            End If

            ASCMAIN1.sql = $"Insert into {ICTPINVX} " & sqlICTPINVX
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New String() {ITEM_CODE})
            'ASCDATA1.ExecuteSQL(Replace(ASCMAIN1.sql, ":PARM1", $"'{ITEM_CODE}'"))
        End If
    End Sub

    '    Function Set_sqlYP(RYP As String)

    '        Dim sqlYP As String = sqlDPTMRPGO

    '        Dim RYP_diff As Integer = ASCMAIN1.Period_Diff(ASCMAIN1.CYP, RYP) + 1
    '        Dim mx As String = Format(RYP_diff, "00")
    '        sqlYP = Replace(sqlYP, $"QTY_00 EOM", $"QTY_{mx} EOM")
    '        sqlYP = Replace(sqlYP, $"QTY_00 POS", $"QTY_{mx} POS")
    '        sqlYP = Replace(sqlYP, $"'000000' OPS_YYYYPP", $"'{RYP}' OPS_YYYYPP")


    '        'For m As Integer = 25 To 1 Step -1
    '        '    Dim m0 As String = Format(m, "00")
    '        '    Dim mx As String = Format(m + X, "00")
    '        '    sqlYP = Replace(sqlYP, $"QTY_{m0}", $"QTY_{mx}")
    '        'Next

    '        Return sqlYP
    '    End Function

    '    Sub Create_Worktables_DPTMRPGO(initialize As Boolean)
    '        'Create_Worktables_DPTMRPGO(initialize As Boolean, Optional sqlYP As String = "")


    '        If initialize Then

    '            sqlDPTMRPGO = "SELECT '000000' OPS_YYYYPP, ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_COST_STD
    ', ICTITEM1.COLLECTION_CODE, ICTCOLL1.BRAND_CODE, ICTITEM1.VEND_CODE, ICTITEM1.ITEM_STATUS
    ', ICTITEM1.PROD_CODE, ICTITEM1.ITEM_SNU_CODE, ICTITEM1.ITEM_BASIC_PROMO
    ', ICTITEM1.ITEM_POS_MAX, ICTITEM1.ITEM_POS_MIN
    ', 0 FC
    ', DEM00, DEM01, DEM02, DEM03, DEM04, DEM05, DEM06, DEM07, DEM08, DEM09, DEM10, DEM11, DEM12
    ', DEM13, DEM14, DEM15, DEM16, DEM17, DEM18, DEM19, DEM20, DEM21, DEM22, DEM23, DEM24, DEM25
    ', OPO00, OPO01, OPO02, OPO03, OPO04, OPO05, OPO06, OPO07, OPO08, OPO09, OPO10, OPO11, OPO12
    ', OPO13, OPO14, OPO15, OPO16, OPO17, OPO18, OPO19, OPO20, OPO21, OPO22, OPO23, OPO24, OPO25
    ', EOM.EOM, EOM.EOM * ICTITEM1.ITEM_COST_STD EOM_EXT_COST, POS.POS
    ' FROM ICTITEM1, ICTCOLL1
    ', (Select ITEM_CODE, QTY_00 EOM from DPTMRPG1 WHERE MRP_TYPE = '5') EOM
    ', (Select ITEM_CODE, QTY_00 POS from DPTMRPG1 WHERE MRP_TYPE = '6') POS
    ', (Select ITEM_CODE, QTY_00 DEM00
    ', QTY_01 DEM01, QTY_02 DEM02, QTY_03 DEM03, QTY_04 DEM04, QTY_05 DEM05, QTY_06 DEM06
    ', QTY_07 DEM07, QTY_08 DEM08, QTY_09 DEM09, QTY_10 DEM10, QTY_11 DEM11, QTY_12 DEM12
    ', QTY_13 DEM13, QTY_14 DEM14, QTY_15 DEM15, QTY_16 DEM16, QTY_17 DEM17, QTY_18 DEM18
    ', QTY_19 DEM19, QTY_20 DEM20, QTY_21 DEM21, QTY_22 DEM22, QTY_23 DEM23, QTY_24 DEM24, QTY_25 DEM25
    ' from DPTMRPG1 WHERE MRP_TYPE = '1') FCS
    ', (Select ITEM_CODE, QTY_00 OPO00
    ', QTY_01 OPO01, QTY_02 OPO02, QTY_03 OPO03, QTY_04 OPO04, QTY_05 OPO05, QTY_06 OPO06
    ', QTY_07 OPO07, QTY_08 OPO08, QTY_09 OPO09, QTY_10 OPO10, QTY_11 OPO11, QTY_12 OPO12
    ', QTY_13 OPO13, QTY_14 OPO14, QTY_15 OPO15, QTY_16 OPO16, QTY_17 OPO17, QTY_18 OPO18
    ', QTY_19 OPO19, QTY_20 OPO20, QTY_21 OPO21, QTY_22 OPO22, QTY_23 OPO23, QTY_24 OPO24, QTY_25 OPO25
    ' from DPTMRPG1 WHERE MRP_TYPE = '3') OPO
    'where EOM.ITEM_CODE = ICTITEM1.ITEM_CODE
    '  AND POS.ITEM_CODE = ICTITEM1.ITEM_CODE
    '  AND FCS.ITEM_CODE = ICTITEM1.ITEM_CODE
    '  AND OPO.ITEM_CODE = ICTITEM1.ITEM_CODE
    '  AND ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE"
    '            DPTMRPGO = ASCMAIN1.Temp_Table(sqlDPTMRPGO)

    '        Else

    '            'ASCMAIN1.sql = $"Truncate Table {DPTMRPGO}"
    '            'ASCDATA1.ExecuteSQL()

    '            'Dim sqlYP As String = Set_sqlYP()

    '            If chkShowAllMonths.Checked Then

    '                'Dim yp As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 1)
    '                Dim yp As String = ASCMAIN1.CYP
    '                Do While yp <= RYP_end

    '                    ASCMAIN1.sql = $"Truncate Table {DPTMRPGO}"
    '                    ASCDATA1.ExecuteSQL()

    '                    Dim sqlYP As String = Set_sqlYP(yp)
    '                    ASCMAIN1.sql = $"Insert into {DPTMRPGO} " & sqlYP
    '                    ASCDATA1.ExecuteSQL()
    '                    Fill_DPTMRPGO(yp)

    '                    yp = ASCMAIN1.Period_Calc(yp, 1)
    '                Loop


    '            Else
    '                ASCMAIN1.sql = $"Truncate Table {DPTMRPGO}"
    '                ASCDATA1.ExecuteSQL()

    '                Dim sqlYP As String = Set_sqlYP(RYP_end)
    '                ASCMAIN1.sql = $"Insert into {DPTMRPGO} " & sqlYP
    '                ASCDATA1.ExecuteSQL()
    '                Fill_DPTMRPGO(RYP_end)
    '            End If

    '        End If
    '    End Sub

    '    Sub Fill_DPTMRPGO(RYPx As String)

    '        Dim X As Integer = ASCMAIN1.Period_Diff(ASCMAIN1.CYP, RYPx) + 1

    '        Fill_Records("DPTMRPGO", , False)

    '        For Each rowDPTMRPGO As DataRow In dst.Tables("DPTMRPGO").Select($"OPS_YYYYPP = '{RYPx}'")
    '            Dim ITEM_CODE As String = rowDPTMRPGO.Item("ITEM_CODE") & ""
    '            Dim ITEM_POS_MAX As Decimal = Val(rowDPTMRPGO.Item("ITEM_POS_MAX") & "")
    '            If ITEM_POS_MAX = 0 Then
    '                rowDPTMRPGO.Item("FC") = 0
    '            Else
    '                Dim FC As Int32 = 0 ' Val(rowDPTMRPGO.Item($"DEM{Format(0, "00")}") & "") ' PD
    '                Dim P As Decimal = Math.Truncate(ITEM_POS_MAX)
    '                'If ASCMAIN1.Running_in_VS AndAlso ITEM_CODE = "BN001A01" Then Stop
    '                If P >= 1 Then
    '                    For I As Integer = 1 To P
    '                        If I + X < 25 Then
    '                            FC += Val(rowDPTMRPGO.Item($"DEM{Format(I + X, "00")}") & "")
    '                        End If
    '                    Next
    '                End If
    '                'If ASCMAIN1.Running_in_VS AndAlso ITEM_CODE = "CH017A06" Then Stop
    '                If ITEM_POS_MAX - P > 0 Then
    '                    If P + X + 1 < 25 Then
    '                        FC += Val(rowDPTMRPGO.Item($"DEM{Format(P + X + 1, "00")}") & "") * (ITEM_POS_MAX - P)
    '                    End If
    '                End If
    '                rowDPTMRPGO.Item("FC") = FC
    '            End If

    '            Dim FC_CUR As Int32 = 0 ' Val(rowDPTMRPGO.Item($"DEM{Format(0, "00")}") & "") ' PD
    '            'If ASCMAIN1.Running_in_VS AndAlso ITEM_CODE = "CH017A06" Then Stop
    '            Dim FC_FUT As Int32 = 0
    '            For I As Integer = 0 To 24
    '                Dim DEM As Int32 = Val(rowDPTMRPGO.Item($"DEM{Format(I, "00")}") & "")
    '                If I <= X Then
    '                    FC_CUR += DEM
    '                Else
    '                    FC_FUT += DEM
    '                    'Debug.Print(DEM)
    '                End If
    '            Next
    '            rowDPTMRPGO.Item("FC_CUR") = FC_CUR
    '            rowDPTMRPGO.Item("FC_FUT") = FC_FUT

    '            Dim PO_CUR As Int32 = 0
    '            Dim PO_FUT As Int32 = 0
    '            For I As Integer = 0 To 24
    '                Dim OPO As Int32 = Val(rowDPTMRPGO.Item($"OPO{Format(I, "00")}") & "")
    '                If I <= X Then
    '                    PO_CUR += OPO
    '                Else
    '                    PO_FUT += OPO
    '                End If
    '            Next
    '            rowDPTMRPGO.Item("PO_CUR") = PO_CUR
    '            rowDPTMRPGO.Item("PO_FUT") = PO_FUT
    '        Next

    '    End Sub

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

    Private Sub optCM_ValueChanged(sender As Object, e As EventArgs) Handles optCM.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If grdGLTPARM2.Selected.Rows.Count > 0 Then
            cmdFetchSales.PerformClick()
        End If
    End Sub

    Private Sub chkEditForecasts_HandleDestroyed(sender As Object, e As EventArgs) Handles chkEditForecasts.HandleDestroyed

    End Sub

    Private Sub tabProjections_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabProjections.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub

        If e.Tab.Text = "Forecasts vs Actual" Then
            grdDPTITMFX.Parent = SplitContainer2.Panel2
            '      chkEditForecasts.Checked = False
            With grdDPTITMFX.DisplayLayout.Bands(0)
                .Columns("MARKET_DESC").Hidden = True
            End With
        ElseIf e.Tab.Text = "Forecasts by Market" Then
            grdDPTITMFX.Parent = UltraTabPageControl8
            With grdDPTITMFX.DisplayLayout.Bands(0)
                .Columns("MARKET_DESC").Hidden = False
            End With

        End If

        'If e.Tab.Text = "Forecasts vs Actual" Then
        'ElseIf e.Tab.Text = "Forecasts by Market" Then
        'End If

    End Sub

    Private Sub btnPosOverMax_Click(sender As Object, e As EventArgs) Handles btnPosOverMax.Click

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Generating Report")

        RYP_desc = Absx1.cmbFor("RYP").Value
        RYP_end = Mid(RYP_desc, 1, 4) & Mid(RYP_desc, 6, 2)

        dst.Tables("DPTMRPGO").Rows.Clear()

        TAC.DPCMAIN1.Create_Worktables_DPTMRPGO(False, Me, sqlDPTMRPGO, DPTMRPGO, RYP_end, chkShowAllMonths.Checked)
        'Create_Worktables_DPTMRPGO(False)

        Set_Filter_for_Excess()

        Sort_grdColumns(grdDPTMRPGO, "OPS_YYYYPP,BRAND_CODE,COLLECTION_CODE,ITEM_CODE")

        splMessages1.Visible = False
        grdDPTMRPGO.Visible = True

        btnPosOverMaxHide.Visible = True

        chkShowAllItems.Visible = True
        chkShowAllItems.Enabled = False
        btnPosOverMaxHide.Visible = True

        grdDPTMRPGO.DisplayLayout.Bands(0).SortedColumns.Clear()

        If chkShowAllMonths.Checked Then
            grdDPTMRPGO.DisplayLayout.Bands(0).Columns("OPS_YYYYPP").Hidden = False
            grdDPTMRPGO.DisplayLayout.Bands(0).SortedColumns.Add("OPS_YYYYPP", False, True)

        Else
            grdDPTMRPGO.DisplayLayout.Bands(0).Columns("OPS_YYYYPP").Hidden = True

        End If

        Set_Format()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Set_Filter_for_Excess()

        Dim dvw As DataView = DirectCast(grdDPTMRPGO.DataSource, DataTable).DefaultView

        Dim sqlw As String = ""
        If chkShowAllItems.Checked Then
            grdDPTMRPGO.Text = "All Items"
            sqlw = ""
        Else
            grdDPTMRPGO.Text = "Items with EOM On Hand over Max Position in " & RYP_desc
            sqlw = " and OVER_QTY > 0"
        End If

        If chkExcludeZeros.Checked Then
            grdDPTMRPGO.Text &= " (Excluding Items with All 0's)"
            sqlw &= " and ZERO <> '0'"
        End If

        dvw.RowFilter = Mid(sqlw, 6)

    End Sub

    Private Sub btnPosOverMaxHide_Click(sender As Object, e As EventArgs) Handles btnPosOverMaxHide.Click

        splMessages1.Visible = True
        grdDPTMRPGO.Visible = False
        btnPosOverMaxHide.Visible = False
        chkShowAllItems.Visible = False
        btnPosOverMaxHide.Visible = False
    End Sub

    Private Sub chkShowAllItems_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowAllItems.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Set_Filter_for_Excess()
    End Sub

    Private Sub chkExcludeZeros_CheckedChanged(sender As Object, e As EventArgs) Handles chkExcludeZeros.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Set_Filter_for_Excess()
    End Sub

    Private Sub chkExportFormat_CheckedChanged(sender As Object, e As EventArgs) Handles chkExportFormat.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Set_Format()
    End Sub

    Sub Set_Format()

        ' per NF email - see wjz response 10/25/24
        Dim tf As Boolean = chkExportFormat.Checked
        With grdDPTMRPGO.DisplayLayout.Bands(0)
            .Columns("ITEM_POS_MAX").Hidden = tf
            .Columns("ITEM_POS_MIN").Hidden = tf

            .Columns("OVER_QTY").Hidden = tf
            .Columns("OVER_EXT_COST").Hidden = tf

            .Columns("FC_CUR").Hidden = tf
            .Columns("FC_FUT").Hidden = tf
            .Columns("PO_CUR").Hidden = tf
            .Columns("PO_FUT").Hidden = tf
            .Columns("PP_CUR").Hidden = tf
            .Columns("PP_FUT").Hidden = tf

            .Columns("FC").Hidden = tf
            .Columns("FC_EXT_COST").Hidden = tf
        End With
    End Sub

    Private Sub chkITEM_ABC_PARMS_LOCKED_CheckedChanged(sender As Object, e As EventArgs) Handles chkITEM_ABC_PARMS_LOCKED.CheckedChanged
        Set_ABC_Parameters_ReadOnly()
        If ScreenMode Then
            Set_ABC_Parameter_Values(txtITEM_ABC_CODE.Text)
        End If
    End Sub

    Sub Set_ABC_Parameters_ReadOnly()
        Dim TF As Boolean = Not Absx1.chkFor("ITEM_ABC_PARMS_LOCKED").Checked
        Absx1.numFor("ITEM_POS_MAX").ReadOnly = TF And ScreenMode
        Absx1.numFor("ITEM_POS_MIN").ReadOnly = TF And ScreenMode
        Absx1.numFor("ITEM_MIN_DAYS_SUPPLY").ReadOnly = TF And ScreenMode
    End Sub

    Sub Set_ABC_Parameter_Values(ITEM_ABC_CODE As String)
        If Not chkITEM_ABC_PARMS_LOCKED.Checked Then
            If ITEM_ABC_CODE <> "" Then
                Dim row As DataRow = LookUp("DPTABCP1", ITEM_ABC_CODE)
                If row IsNot Nothing Then
                    Dim ABC_MAX_POS As Decimal = Val(row.Item("ABC_MAX_POS") & "")
                    Dim ABC_MIN_POS As Decimal = Val(row.Item("ABC_MIN_POS") & "")
                    Dim ABC_MIN_DAYS_SUPPLY As Integer = Val(row.Item("ABC_MIN_DAYS_SUPPLY") & "")
                    Absx1.numFor("ITEM_POS_MAX").Value = ABC_MAX_POS
                    Absx1.numFor("ITEM_POS_MIN").Value = ABC_MIN_POS
                    Absx1.numFor("ITEM_MIN_DAYS_SUPPLY").Value = ABC_MIN_DAYS_SUPPLY
                End If
            End If
        End If
    End Sub

    Private Sub txtITEM_ABC_CODE_ValueChanged(sender As Object, e As EventArgs) Handles txtITEM_ABC_CODE.ValueChanged
        If ScreenMode Then
            Set_ABC_Parameter_Values(txtITEM_ABC_CODE.Text)
        End If
    End Sub
    Private Sub Set_LYVF_Tab_Visible()
        Try
            Dim tabKey As String = "Last Year Sales vs Forecast"
            Dim showIt As Boolean = chkShowTab.Checked

            tabDetails.Tabs(tabKey).Visible = showIt

            If showIt Then
                tabDetails.SelectedTab = tabDetails.Tabs(tabKey)
            End If

        Catch ex As Exception
        End Try
    End Sub
    Private Sub chkShowTab_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowTab.CheckedChanged
        Set_LYVF_Tab_Visible()

        If chkShowTab.Checked Then
            Refresh_LYVF_Chart()
        End If
    End Sub
    Private Function Build_LYVF_Data() As DataTable

        Dim dt As New DataTable("LYVF")
        dt.Columns.Add("MONTH_LABEL", GetType(String))
        dt.Columns.Add("MONTH_YEAR", GetType(String))
        dt.Columns.Add("TY Forecast", GetType(Integer))
        dt.Columns.Add("LY Sales", GetType(Integer))

        Dim rptYP As String = ASCMAIN1.CYP

        Dim fc(12) As Integer
        fc = Get_Forecast_13(rptYP)

        Dim ly(12) As Integer
        ly = Get_LY_Sales_13(rptYP)

        For i As Integer = 0 To 12
            Dim yp As String = ASCMAIN1.Period_Calc(rptYP, i)  ' YYYYPP for TY range
            Dim label As String = Period_To_MonthLabel(yp)     ' "Mar"
            Dim labelYear As String = Period_To_MonthYearLabel(yp) ' "Mar 2026"

            Dim r As DataRow = dt.NewRow()
            r("MONTH_LABEL") = label
            r("MONTH_YEAR") = labelYear
            r("TY Forecast") = fc(i)
            r("LY Sales") = ly(i)
            dt.Rows.Add(r)
        Next

        Return dt

    End Function
    Private Function Period_To_MonthLabel(ByVal opsYYYYPP As String) As String
        If opsYYYYPP Is Nothing OrElse opsYYYYPP.Length < 6 Then Return opsYYYYPP

        Dim yyyy As Integer = Val(Mid(opsYYYYPP, 1, 4))
        Dim mm As Integer = Val(Mid(opsYYYYPP, 5, 2))

        Dim d As New Date(yyyy, Math.Max(1, Math.Min(12, mm)), 1)
        Return d.ToString("MMM") ' Feb/Mar etc
    End Function
    Private Function Period_To_MonthYearLabel(ByVal opsYYYYPP As String) As String
        If opsYYYYPP Is Nothing OrElse opsYYYYPP.Length < 6 Then Return opsYYYYPP

        Dim yyyy As Integer = Val(Mid(opsYYYYPP, 1, 4))
        Dim mm As Integer = Val(Mid(opsYYYYPP, 5, 2))

        Dim d As New Date(yyyy, Math.Max(1, Math.Min(12, mm)), 1)
        Return d.ToString("MMM yyyy") ' "Mar 2026"
    End Function
    Private Function Get_Forecast_13(ByVal rptYP As String) As Integer()

        Dim result(12) As Integer
        Dim t As DataTable = dst.Tables("DPTITMFX")

        For i As Integer = 0 To 12
            Dim col As String = "FC" & Format(i, "00")

            Dim sum As Long = 0
            For Each r As DataRow In t.Rows
                sum += Val(r(col) & "")
            Next

            result(i) = CInt(Math.Min(Integer.MaxValue, Math.Max(Integer.MinValue, sum)))
        Next

        Return result

    End Function
    Private Function Get_LY_Sales_13(ByVal rptYP As String) As Integer()
        Dim result(12) As Integer

        Try
            If grdDPTPLANX.ActiveRow Is Nothing Then
                Return result
            End If

            Dim ITEM_CODE As String = grdDPTPLANX.ActiveRow.Cells("ITEM_CODE").Text

            Dim ypStart As String = ASCMAIN1.Period_Calc(rptYP, -12)
            Dim ypEnd As String = rptYP

            Dim sql As String = Build_LY_Sales_SQL_ItemUnits_ByCustomer(ITEM_CODE, ypStart, ypEnd)

            Dim dt As DataTable = ASCDATA1.GetDataTable(sql, "LY_SOTINVHX")

            For i As Integer = 0 To 12
                Dim colName As String = "S" & (i + 1).ToString("00")
                If dt Is Nothing OrElse Not dt.Columns.Contains(colName) Then
                    result(i) = 0
                    Continue For
                End If

                Dim sum As Long = 0
                For Each r As DataRow In dt.Rows
                    If r.RowState = DataRowState.Deleted Then Continue For
                    If IsDBNull(r(colName)) Then Continue For
                    sum += CLng(Val(r(colName).ToString()))
                Next

                result(i) = CInt(Math.Min(Integer.MaxValue, Math.Max(Integer.MinValue, sum)))
            Next

        Catch ex As Exception
        End Try

        Return result
    End Function
    Private Function Build_LY_Sales_SQL_ItemUnits_ByCustomer(ByVal itemCode As String, ByVal ypStart As String, ByVal ypEnd As String) As String

        Dim inner As New System.Text.StringBuilder()
        inner.Append("Select CUST_CODE CODE_VALUE")

        For i As Integer = 0 To 12
            Dim yp As String = ASCMAIN1.Period_Calc(ypStart, i)
            Dim sCol As String = "S" & (i + 1).ToString("00")
            inner.Append(", SUM (DECODE(ORDR_YYYYPP_UPDATED,'")
            inner.Append(yp)
            inner.Append("',ORDR_QTY_SHIP,0)) ")
            inner.Append(sCol)
        Next

        inner.Append(" from SOTINVH2")
        inner.Append(" where ORDR_YYYYPP_UPDATED between '")
        inner.Append(ypStart)
        inner.Append("' and '")
        inner.Append(ypEnd)
        inner.Append("'")
        inner.Append(" and ITEM_CODE = '")
        inner.Append(itemCode.Replace("'", "''"))
        inner.Append("'")
        inner.Append(" group by CUST_CODE")

        Dim sql As String =
        "Select X.*, ARTCUST1.CUST_NAME DESC_VALUE " &
        "from (" & inner.ToString() & ") X, ARTCUST1 " &
        "where X.CODE_VALUE = ARTCUST1.CUST_CODE"

        Return sql
    End Function
    Private Sub Refresh_LYVF_Chart()
        Try
            Dim dt As DataTable = Build_LYVF_Data()

            chtLYVF.DataSource = Nothing

            chtLYVF.Data.SwapRowsAndColumns = True

            chtLYVF.Data.UseRowLabelsColumn = True
            chtLYVF.Data.RowLabelsColumn = 0

            ' Bind
            chtLYVF.DataSource = dt
            chtLYVF.DataBind()

            Try
                Dim ypTYStart As String = ASCMAIN1.CYP
                Dim ypTYEnd As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 12)

                Dim ypLYStart As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12)
                Dim ypLYEnd As String = ASCMAIN1.CYP

                Dim tyRange As String = Period_To_MonthYearLabel(ypTYStart) & "–" & Period_To_MonthYearLabel(ypTYEnd)
                Dim lyRange As String = Period_To_MonthYearLabel(ypLYStart) & "–" & Period_To_MonthYearLabel(ypLYEnd)

                If chtLYVF.Series IsNot Nothing Then
                    If chtLYVF.Series.Count >= 1 Then chtLYVF.Series(0).Label = "TY Forecast (" & tyRange & ")"
                    If chtLYVF.Series.Count >= 2 Then chtLYVF.Series(1).Label = "LY Sales (" & lyRange & ")"
                End If
            Catch
            End Try

            Try
                chtLYVF.Axis.X.Extent = 45
            Catch
            End Try

            Try
                Dim sumFC As Long = 0
                Dim sumLY As Long = 0

                For Each r As DataRow In dt.Rows
                    sumFC += CLng(Val(r("TY Forecast") & ""))
                    sumLY += CLng(Val(r("LY Sales") & ""))
                Next

                Dim diff As Long = sumFC - sumLY
                Dim sign As String = If(diff >= 0, "+", "-")

                chtLYVF.TitleTop.Text =
                "Last Year Sales vs Forecast (Units)  " &
                "TY=" & Format(sumFC, "#,0") &
                "  LY=" & Format(sumLY, "#,0") &
                "  (" & sign & Format(Math.Abs(diff), "#,0") & " vs LY)"

                Try
                    chtLYVF.TitleTop.HorizontalAlign = StringAlignment.Center
                Catch
                End Try
            Catch
            End Try

            Try
                chtLYVF.Axis.Y.RangeMin = 0
            Catch
            End Try

            Format_LYVF_Chart_For_Clarity()

            chtLYVF.Legend.Visible = False
            chtLYVF.Legend.Visible = True


            chtLYVF.Refresh()

        Catch ex As Exception
        End Try
    End Sub
    Private Sub Format_LYVF_Chart_For_Clarity()

        Try
            chtLYVF.Axis.X.Labels.ItemFormatString = "<ITEM_LABEL>"
            chtLYVF.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:0,0>"

            chtLYVF.Axis.Y.MajorGridLines.Visible = True
            chtLYVF.Axis.X.MajorGridLines.Visible = False


            chtLYVF.Axis.X.Labels.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            chtLYVF.Axis.X.Labels.Orientation = Infragistics.UltraChart.Shared.Styles.TextOrientation.Horizontal
            chtLYVF.Axis.X.Labels.HorizontalAlign = StringAlignment.Center

            chtLYVF.Axis.X.Extent = 70

            chtLYVF.Padding = New Padding(10, 10, 60, 35)
            chtLYVF.ColorModel.ModelStyle = Infragistics.UltraChart.Shared.Styles.ColorModels.CustomLinear
            chtLYVF.ColorModel.AlphaLevel = 255

            Dim pal(1) As Color
            pal(0) = Color.FromArgb(0, 102, 204)  ' TY Forecast
            pal(1) = Color.FromArgb(255, 102, 0)  ' LY Sales
            chtLYVF.ColorModel.CustomPalette = pal

            If chtLYVF.Series IsNot Nothing Then
                If chtLYVF.Series.Count >= 1 Then Force_Series_Style(chtLYVF.Series(0), pal(0))
                If chtLYVF.Series.Count >= 2 Then Force_Series_Style(chtLYVF.Series(1), pal(1))
            End If
            chtLYVF.Legend.Visible = True

            Try
                chtLYVF.Legend.Location = Infragistics.UltraChart.Shared.Styles.LegendLocation.Right
            Catch
            End Try

            Try
                chtLYVF.Legend.SpanPercentage = 12
            Catch
            End Try

            Try
                chtLYVF.Legend.BorderThickness = 0
            Catch
            End Try

            Try
                chtLYVF.Tooltips.FormatString =
        "<SERIES_LABEL>" & vbCrLf &
        "<ITEM_LABEL>: <DATA_VALUE:#,##0>"
            Catch
            End Try

            Try
                chtLYVF.Tooltips.Font = New Font("Segoe UI", 10, FontStyle.Bold)
            Catch
            End Try


        Catch ex As Exception
        End Try

    End Sub
    Private Sub Force_Series_Style(ByVal s As Object, ByVal c As System.Drawing.Color)

        Try : s.PE.Fill = c : Catch : End Try
        Try : s.PE.Stroke = c : Catch : End Try
        Try : s.PE.StrokeColor = c : Catch : End Try
        Try : s.PE.ElementColor = c : Catch : End Try
        Try : s.PE.LineColor = c : Catch : End Try

        Try : s.Appearance.Color = c : Catch : End Try
        Try : s.Appearance.ForeColor = c : Catch : End Try
        Try : s.Appearance.BorderColor = c : Catch : End Try


        Try : s.PE.StrokeWidth = 5 : Catch : End Try
        Try : s.PE.BorderWidth = 3 : Catch : End Try
        Try : s.PE.Width = 3 : Catch : End Try

        Try : s.PE.MarkerSize = 7 : Catch : End Try
        Try : s.MarkerSize = 7 : Catch : End Try
        Try : s.ShowMarkers = True : Catch : End Try

    End Sub
End Class