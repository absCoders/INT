Public Class SAFCSLS1
    Dim SATCSLS1 As String
    Dim SATCSLS2 As String
    Dim sqlSATCSLS1 As String
    Dim sqlSATCSLS1_sum As String
    Dim sqlSATCSLS2 As String
    Dim sqlSOTINVHX As String
    Dim sqlSATCSLS1_STORES As String
    Dim STORES As List(Of String)
    Dim STORES_XX As String
    Dim PERIODS_XX As String
    Dim summary_loaded As Boolean

    Dim sqlSPTCOOPX As String
    Dim APPR_STATUS_CODE_BackColors As New Dictionary(Of String, System.Drawing.Color)
    Dim APPR_STATUS_CODE_ForeColors As New Dictionary(Of String, System.Drawing.Color)

    Dim rowARTCUST1 As DataRow

    Dim Ps As New List(Of String)
    Dim WPs As New List(Of String)
    Dim Ls As New List(Of String)
    Dim PE_DATEs() As Date

    Dim PsLY As New List(Of String)
    Dim WPsLY As New List(Of String)
    Dim LsLY As New List(Of String)
    Dim PE_DATEsLY() As Date

    Dim SATSIST1 As String = ""
    Dim SATSIST2 As String = ""
    Dim SATSIST3 As String = ""
    Dim sqlSATSIST2 As String

    Dim RYP0 As String
    Dim RYP1 As String
    Dim RYW0 As String
    Dim RYW1 As String
    Dim Periods As Integer
    Dim CUST_CODE As String
    Dim ITEM_CODE_LY As String
    Dim US_STATES() As String
    Dim USmap As MapLayer
    Dim Stores_Max As Integer = 120
    Dim SIST_PMAX As Integer = 53

    Dim QTY_ALLO_PLAN As Int64

    Dim sqlSATSISTI As String
    Dim sqlSATSISTC As String

    Dim COLUMN_NAMEs As New ArrayList
    Dim COLUMN_CAPTIONs As New ArrayList
    Dim COLUMN_NAME_by_Lvl() As String
    Dim COLUMN_CAPTION_by_Lvl() As String
    Dim G_by_Lvl() As Integer
    Dim SCOPE() As String
    Dim QCOLS As New Dictionary(Of String, String)
    Dim LVL As Int16
    Dim XLS_FILENAME_multi As String = ""

    Dim MAX_SALES As Decimal

    Dim needsEmptyWeek As Boolean = False


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Set_cmbYP("RYP0", ASCMAIN1.CYP, -36, 0, -11)
        'Set_cmbYP("RYP1", ASCMAIN1.CYP, -36, 0, 0)

        'Set_cmbYW("RYW0", ASCMAIN1.CYW, -3 * 52, 0, -13)
        'Set_cmbYW("RYW1", ASCMAIN1.CYW, -3 * 52, 0, 0)
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 12, -11) ' -11 - 12)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -60, 12, 0) '  0 - 12)

        Set_cmbYW("RYW0", ASCMAIN1.CYW, -5 * 52, 1 * 52, -13) '  -13 - 1 * 52)
        Set_cmbYW("RYW1", ASCMAIN1.CYW, -5 * 52, 1 * 52, 0) ' 0 - 1 * 52)

        With dst
            ASCMAIN1.sql = "Select ITEM_CODE CODE_VALUE, ITEM_DESC DESC_VALUE from ICTITEM1"
            Create_TDA(.Tables.Add, "SATCSLS1", "**", 0, False)
            With .Tables("SATCSLS1")
                .Columns.Add("SUB_CODE_VALUE1")
                .Columns.Add("SUB_CODE_VALUE2")
                .Columns.Add("SUB_CODE_VALUE3")
                .Columns.Add("SUB_CODE_VALUE4")
                .Columns.Add("SUB_CODE_VALUE5")
                .Columns.Add("RTL_PRICE", GetType(System.Decimal))
                .Columns.Add("WSL_PRICE", GetType(System.Decimal))
                For P As Integer = 0 To Stores_Max
                    .Columns.Add("P" & Format(P, "00"), GetType(System.Int32))
                Next
                .Columns.Add("PXX", GetType(System.Int32))
                .Columns.Add("PYY", GetType(System.Int32))
                .Columns.Add("PYY00PCT", GetType(System.Decimal), "IIF(ISNULL(PYY,0)=0,0,100*P00/PYY)")
            End With

            ASCMAIN1.sql = "Select ITEM_CODE CODE_VALUE, '0' YEAR, 'XX' DATA_TYPE from ICTITEM1"
            Create_TDA(.Tables.Add, "SATCSLS1_DTL", "**", 0, False)
            With .Tables("SATCSLS1_DTL")
                For P As Integer = 0 To Stores_Max
                    .Columns.Add("P" & Format(P, "00"), GetType(System.Int32))
                Next
                .Columns.Add("PXX", GetType(System.Int32))
                .Columns.Add("PYY", GetType(System.Int32))
                .Columns.Add("PYY00PCT", GetType(System.Decimal), "IIF(ISNULL(PYY,0)=0,0,100*P00/PYY)")
            End With

            .Relations.Add("SATCSLS1_SATCSLS1_DTL" _
                           , New DataColumn() {.Tables("SATCSLS1").Columns("CODE_VALUE")} _
                           , New DataColumn() {.Tables("SATCSLS1_DTL").Columns("CODE_VALUE")})

            ASCMAIN1.sql = "Select ITEM_CODE CODE_VALUE_PARENT, ITEM_CODE CODE_VALUE, ITEM_DESC DESC_VALUE from ICTITEM1"
            Create_TDA(.Tables.Add, "SATCSLS2", "**", 0, False)
            With .Tables("SATCSLS2")
                .Columns.Add("SUB_CODE_VALUE1")
                .Columns.Add("SUB_CODE_VALUE2")
                .Columns.Add("SUB_CODE_VALUE3")
                .Columns.Add("SUB_CODE_VALUE4")
                .Columns.Add("SUB_CODE_VALUE5")
                .Columns.Add("RTL_PRICE", GetType(System.Decimal))
                .Columns.Add("WSL_PRICE", GetType(System.Decimal))
                For P As Integer = 0 To Stores_Max
                    .Columns.Add("P" & Format(P, "00"), GetType(System.Int32))
                Next
                .Columns.Add("PXX", GetType(System.Int32))
                .Columns.Add("PYY", GetType(System.Int32))
                .Columns.Add("PYY00PCT", GetType(System.Decimal), "IIF(ISNULL(PYY,0)=0,0,100*P00/PYY)")
            End With

            Create_Relation("SATCSLS1", "SATCSLS2", "CODE_VALUE", "CODE_VALUE_PARENT")

            ' & ", NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_NAME) || DECODE(CUST_STORE_MARK_FOR,NULL,'',' (' || CUST_STORE_MARK_FOR || ')') CUST_STORE_LOCATION" _
            ' RC WANTS THE TRADITIONAL STORE NAME, FORMATTED THE WAY EVERYONE IS USED TO SEEING IT

            ASCMAIN1.sql = "Select SOTINVH2.INV_TYPE, SOTINVH2.INV_NO" _
            & ", SOTINVH1.INV_DATE, SOTINVH1.ORDR_TYPE_CODE, SOTINVH1.ORDR_CUST_PO" _
            & ", SOTINVH2.ORDR_YYYYPP_UPDATED OPS_YYYYPP, SOTINVH2.OPS_YYYYWW" _
            & ", SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO" _
            & ", ARTCUST2.CUST_STORE_NAME CUST_STORE_LOCATION" _
            & ", SOTINVH2.ITEM_CODE, ICTITEM1.ITEM_DESC" _
            & ", SOTINVH2.ORDR_QTY_SHIP, SOTINVH2.ORDR_UNIT_PRICE" _
            & ", SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE ORDR_AMT_SHIP" _
            & " from SOTINVH2,ICTITEM1,ARTCUST2,SOTINVH1 " _
            & " where ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE " _
            & " and ARTCUST2.CUST_CODE (+) = SOTINVH2.CUST_CODE " _
            & " and ARTCUST2.CUST_STORE_NO (+) = SOTINVH2.CUST_STORE_NO " _
            & " and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE " _
            & " and SOTINVH1.INV_NO = SOTINVH2.INV_NO"
            sqlSOTINVHX = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False)

            Create_TDA(.Tables.Add, "EDT852T1", "*", 0, False)

            Create_TDA(.Tables.Add, "TATSTATE", "*", 0, False)

            With .Tables.Add("SATCSLSS")
                .Columns.Add("STATE_CODE")
                .Columns.Add("STATE_NAME")
                .Columns.Add("SALES", GetType(System.Int32))
            End With

            sqlSATSISTI = "Select ICTITEM1.ITEM_CODE" _
            & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_RETAIL_PRICE" _
            & ", ICTITEM1.COLLECTION_CODE, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.ITEM_ALT_SORT" _
            & " from ICTITEM1"
            ASCMAIN1.sql = sqlSATSISTI _
                & " where ICTITEM1.ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SATSISTI", "**", 0, False, "V")

            sqlSATSISTC = "Select ARTCUST1.CUST_CODE" _
                & ", ARTCUST1.CUST_NAME" _
                & ", ARTCUST1.SREP_CODE, ARTCUST1.TRADE_CLASS_CODE" _
                & " from ARTCUST1"
            ASCMAIN1.sql = sqlSATSISTC _
                & " where ARTCUST1.CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SATSISTC", "**", 0, False, "V")

            For Each TABLE_NAME As String In New String() {"SATSIST1", "SATSIST2", "SATSIST3"}
                ASCMAIN1.sql = "Select ITEM_CODE" & vbCrLf _
                    & IIf(TABLE_NAME = "SATSIST2", ", CUST_STORE_NO", "") & vbCrLf _
                    & ", ITEM_DESC from " & IIf(TABLE_NAME = "SATSIST2", "ICTITEM1", "ICTITEM1")
                If TABLE_NAME = "SATSIST2" Then
                    ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ARTCUST2.CUST_STORE_NO, ARTCUST2.CUST_STORE_NAME from ICTITEM1,ARTCUST2"
                Else
                    ASCMAIN1.sql = "Select ITEM_CODE, ITEM_DESC from ICTITEM1"
                End If
                Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, False)

                If TABLE_NAME = "SATSIST2" Then
                    With dst.Tables("SATSIST2")
                        .PrimaryKey = New DataColumn() {.Columns(0), .Columns(1)}
                    End With
                End If

                With .Tables(TABLE_NAME)
                    ' .Columns("ITEM_DESC").MaxLength = -1
                    .Columns.Add("INV_DATE_SHIPPED", GetType(System.DateTime))
                    Dim T As String = ""
                    For P As Integer = 0 To SIST_PMAX
                        .Columns.Add("SIQTY_P" & Format(P, "00"), GetType(System.Int32))
                        .Columns.Add("STQTY_P" & Format(P, "00"), GetType(System.Int32))
                        .Columns.Add("OHQTY_P" & Format(P, "00"), GetType(System.Int32))
                        .Columns.Add("STPCT_P" & Format(P, "00"), GetType(System.Decimal), Replace("IIF(ISNULL(SIQTY_PXX,0)=0,0,100*STQTY_PXX/SIQTY_PXX)", "PXX", "P" & Format(P, "00")))
                        .Columns.Add("STAMT_P" & Format(P, "00"), GetType(System.Decimal))
                        If P > 0 Then T &= " + ISNULL(SIQTY_P" & Format(P, "00") & ",0)"
                    Next
                    .Columns("SIQTY_P00").Expression = "ISNULL(SIQTY_P" & Format(SIST_PMAX, "00") & ",0)"
                    .Columns("STQTY_P00").Expression = Mid(Replace(T, "SIQTY", "STQTY"), 4)
                    .Columns("OHQTY_P00").Expression = "ISNULL(OHQTY_P" & Format(1, "00") & ",0)"
                    .Columns("STAMT_P00").Expression = Mid(Replace(T, "SIQTY", "STAMT"), 4)
                End With
            Next

            sqlSPTCOOPX = "Select SPTCOOP1.*, SPTCOOP3.COLLECTION_CODE, SPTCOOP3.DIST_AMT, SPTCOOP3.AUTH_LNO, SPTCOOP3.FEATURE_DESC" & vbCrLf _
                & ", ICTCOLL1.COLLECTION_NAME, ICTCOLL1.COLLECTION_GENDER, ICTCOLL1.BRAND_CODE, ICTBRAN1.BRAND_NAME, ICTCOLL1.HC_CODE, ICTCOLL0.HC_NAME" & vbCrLf _
                & " from SPTCOOP1,SPTCOOP3,ICTCOLL1,ICTBRAN1,ICTCOLL0,SPTTYPE1" & vbCrLf _
                & " where ICTCOLL1.COLLECTION_CODE = SPTCOOP3.COLLECTION_CODE" & vbCrLf _
                & "   and SPTCOOP1.APPR_STATUS_CODE in ('A','P','G')" & vbCrLf _
                & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
                & "   and ICTCOLL0.HC_CODE = ICTCOLL1.HC_CODE" & vbCrLf _
                & "   and SPTTYPE1.EXPENSE_TYPE_CODE = SPTCOOP1.EXPENSE_TYPE_CODE" & vbCrLf _
                & "   and NVL(SPTTYPE1.EXPENSE_TYPE_INCL_SIST,'0') = '1'" & vbCrLf _
                & "   and SPTCOOP3.AUTH_NO = SPTCOOP1.AUTH_NO"
            ASCMAIN1.sql = sqlSPTCOOPX & "  and SPTCOOP1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "SPTCOOPX", "**", 0, False, "V")

            With .Tables.Add("SATANALC")
                .Columns.Add("COLUMN_NAME")
                .Columns.Add("COLUMN_NAME_CODE")
                .Columns.Add("COLUMN_NAME_DESC")
                .Columns.Add("TABLE_NAME_LOOKUP")
                .PrimaryKey = New DataColumn() {.Columns("COLUMN_NAME")}
            End With

            ASCMAIN1.sql = "Select CTL_NO_TEXT COLUMN_NAME, CTL_NO_TEXT CODE_VALUE, CTL_NO_TEXT DESC_VALUE from TATCTLN1"
            Create_TDA(.Tables.Add, "SATANALD", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select ASTLIST1.* from ASTLIST1 where COLUMN_NAME = 'SATISLS1.CUST_CODE' and LINKED_LIST_CODE is Not Null"
            Create_TDA(.Tables.Add, "ASTLIST1", "**", 0)

            Create_TDA(.Tables.Add, "ASTLIST2", "*", 1)

            .Tables("SATANALD").Columns("DESC_VALUE").MaxLength = -1 ' 100







            ASCMAIN1.sql = "Select INV_DATE, SUM (INV_SALES) INV_SALES, COUNT (*) INVS" & vbCrLf _
                & " from SOTINVH1 where CUST_CODE = :PARM1 group by INV_DATE"
            Create_TDA(.Tables.Add, "SATWKSL1", "**", 0, False, "V", 1)
            ASCMAIN1.sql = "Select NVL(CUST_SHIP_TO_STATE,'??') CUST_SHIP_TO_STATE, SUM (INV_SALES) INV_SALES" & vbCrLf _
                & " from SOTINVH1 where CUST_CODE = :PARM1 and INV_DATE > SYSDATE -365 group by NVL(CUST_SHIP_TO_STATE,'??')"
            Create_TDA(.Tables.Add, "SATWKSLS", "**", 0, False, "V", 1)


            ASCMAIN1.sql = "Select 0 LINE_NO, 'X' MONTH, '99' WEEK_NO"
            For i As Integer = 1 To 5   ' 5 years
                For j As Integer = 1 To 7  ' 7 days in a week
                    ASCMAIN1.sql &= ", 0 DATE_" & Format$(i, "0") & Format$(j, "0")
                Next j
            Next i
            ASCMAIN1.sql &= " from GLTPARM3"

            Create_TDA(.Tables.Add, "ASTWCAL1", "**", 0, False, "", 1)
            With .Tables("ASTWCAL1").Columns
                For I As Integer = 1 To 5
                    .Add("DATES" & CStr(I))
                Next
            End With
        End With

        dst.Tables("SATANALC").Rows.Clear()
        With dst.Tables("SATANALC").Rows
            For Each COLUMN_NAME As String In New String() {"SREP_CODE", "MARKET_CODE", "TRADE_CLASS_CODE"}
                Dim COLUMN_NAME2 As String = COLUMN_NAME
                Select Case COLUMN_NAME
                    Case "SREP_CODE"
                        ASCMAIN1.sql = "Select SREP_CODE, SREP_NAME from SOTSREP1"
                    Case "MARKET_CODE"
                        ASCMAIN1.sql = "Select MARKET_CODE, MARKET_DESC from SOTMKTC1"
                    Case "TRADE_CLASS_CODE"
                        ASCMAIN1.sql = "Select TRADE_CLASS_CODE, TRADE_CLASS_DESC from SOTTCLS1"
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

        Fill_Records("TATSTATE")

        With grdSATCSLS1.DisplayLayout
            .NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
            .NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
            .ViewStyle = UltraWinGrid.ViewStyle.MultiBand
            .MaxBandDepth = 2

            grdSATCSLS1.DataSource = dst.Tables("SATCSLS1")
            .Bands(1).Hidden = True
            .Bands(2).Hidden = True
        End With

        grdASTLIST1.DataSource = New DataView(dst.Tables("ASTLIST1"), "COLUMN_NAME = 'SATISLS1.CUST_CODE'", "INIT_DATE DESC", DataViewRowState.CurrentRows)

        grdSATCSLS2.DataSource = dst.Tables("SATCSLS2")
        grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")
        grdEDT852T1.DataSource = dst.Tables("EDT852T1")
        grdSATSISTI.DataSource = dst.Tables("SATSISTI")
        grdSATSISTC.DataSource = dst.Tables("SATSISTC")
        grdSATSIST1.DataSource = dst.Tables("SATSIST1")
        grdSATSIST2.DataSource = dst.Tables("SATSIST2")
        grdSATSIST3.DataSource = dst.Tables("SATSIST3")
        grdSPTCOOPX.DataSource = dst.Tables("SPTCOOPX")

        Dim dvw As DataView = dst.Tables("SATCSLSS").DefaultView
        dvw.RowFilter = "SALES <> 0"
        grdSATCSLSS.DataSource = dvw

        Create_Summary(grdSOTINVHX, "INV_NO", "Count")
        Create_Summary(grdSOTINVHX, "ORDR_QTY_SHIP")
        Create_Summary(grdSOTINVHX, "ORDR_AMT_SHIP")

        Create_Summary(grdSATCSLSS, "STATE_CODE", "Count")
        Create_Summary(grdSATCSLSS, "SALES")

        Create_Summary(grdSPTCOOPX, "AUTH_NO", "Count")
        Create_Summary(grdSPTCOOPX, New String() {"DIST_AMT", "OPEN_AMT", "PAID_AMT"})

        With grdSATSISTI.DisplayLayout.Bands("SATSISTI")
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key <> "ITEM_CODE" Then GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Violet
            Next
        End With
        With grdSATSISTC.DisplayLayout.Bands("SATSISTC")
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key <> "CUST_CODE" Then GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Violet
            Next
        End With

        With grdSPTCOOPX.DisplayLayout.Bands("SPTCOOPX")
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

                If New String() {"OPEN_AMT", "PAID_AMT"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
            .Columns("AUTH_NO").Header.Fixed = True
        End With



        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSATCSLS1, grdSATCSLS2}
            Create_Summary(grd, "CODE_VALUE", "Count")
            For P As Integer = 0 To Stores_Max
                Create_Summary(grd, "P" & Format(P, "00"))
            Next
            Create_Summary(grd, "PXX")
            Create_Summary(grd, "PYY")
            Create_Summary(grd, "PYY00PCT", "Custom", , "##0.0")
            With grd.DisplayLayout.Bands(0)
                .Columns("CODE_VALUE").Header.Fixed = True
                .Columns("DESC_VALUE").Header.Fixed = True
                .Columns("SUB_CODE_VALUE1").Header.Fixed = True
                .Columns("SUB_CODE_VALUE2").Header.Fixed = True
                .Columns("SUB_CODE_VALUE3").Header.Fixed = True
                .Columns("SUB_CODE_VALUE4").Header.Fixed = True
                .Columns("SUB_CODE_VALUE5").Header.Fixed = True
                .Columns("RTL_PRICE").Header.Fixed = True
                .Columns("WSL_PRICE").Header.Fixed = True
                .Columns("P00").Header.Fixed = True
                .Columns("PXX").Header.Fixed = True
                .Columns("PYY").Header.Fixed = True
                .Columns("PYY00PCT").Header.Fixed = True
            End With
        Next

        Create_Summary(grdEDT852T1, "EDI_DOC_SEQ_NO", "Count")
        With grdEDT852T1.DisplayLayout.Bands("EDT852T1")
            .Columns("EDI_DOC_SEQ_NO").Header.Fixed = True
        End With

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSATSIST1, grdSATSIST2, grdSATSIST3}
            With grd.DisplayLayout.Bands(0)
                Dim G As UltraWinGrid.UltraGridGroup = .Groups.Add
                ' G.Header.Caption = "Customer & 1st Date Shipped"
                For Each COLUMN_NAME As String In New String() {"ITEM_CODE", "ITEM_DESC", "INV_DATE_SHIPPED"}
                    If COLUMN_NAME = "ITEM_CODE" Then
                        If grd.Name = "grdSATSIST2" Then
                            COLUMN_NAME = "CUST_STORE_NO"
                        End If
                        Create_Summary(grd, COLUMN_NAME, "Count")
                    End If
                    If COLUMN_NAME = "ITEM_DESC" Then
                        If grd.Name = "grdSATSIST2" Then
                            COLUMN_NAME = "CUST_STORE_NAME"
                        End If
                    End If
                    With .Columns(COLUMN_NAME)
                        .Group = G
                        If grd.Name = "grdSATSIST2" Then
                            .Hidden = IIf(COLUMN_NAME = "CUST_STORE_NO", True, False) ' False
                        Else
                            .Hidden = IIf(COLUMN_NAME = "ITEM_DESC", True, False) ' False
                        End If
                        .Width = 100 '  IIf(COLUMN_NAME = "CUST_NAME", 200, 100)

                        .Header.Appearance.BackColor = Drawing.Color.White
                        .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                        .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    End With

                Next

                G.Header.Fixed = True
                G.Header.Appearance.BackColor = Drawing.Color.White
                G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                G.Header.Appearance.BackColor2 = Drawing.Color.LightGray

                .Columns("ITEM_CODE").CellAppearance.BackColor = Drawing.Color.LightGray

                If grd.Name = "grdSATSIST2" Then
                    .Columns("CUST_STORE_NO").CellAppearance.BackColor = Drawing.Color.LightGray
                    .Columns("CUST_STORE_NO").Header.Caption = "Store"
                End If

                .Columns("ITEM_CODE").Header.Caption = "Item" ' "Code"

                If grd.Name = "grdSATSIST2" Then
                    .Columns("CUST_STORE_NAME").CellAppearance.BackColor = Drawing.Color.LightGray
                    .Columns("CUST_STORE_NAME").Header.Caption = "Store Name"
                Else
                    .Columns("ITEM_DESC").Header.Caption = "Item Description"
                End If
                .Columns("INV_DATE_SHIPPED").Header.Caption = "Shipped"

                For P As Integer = 0 To SIST_PMAX
                    G = .Groups.Add
                    If P = 0 Then G.Header.Fixed = True
                    G.Header.Appearance.TextHAlign = HAlign.Center
                    G.Header.Appearance.BackColor = Drawing.Color.White
                    G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    If P = 0 Then
                        G.Header.Caption = "Totals"
                        G.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    Else
                        G.Header.Caption = "MM/dd - MM/dd Dec Week 1/5"
                        G.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    End If

                    For Each COLUMN_NAME As String In New String() {"SIQTY", "STQTY", "OHQTY", "STPCT", "STAMT"}
                        COLUMN_NAME &= "_P" & Format(P, "00")
                        .Columns(COLUMN_NAME).Group = G
                        .Columns(COLUMN_NAME).Hidden = False
                        If COLUMN_NAME.StartsWith("STPCT") Then
                            .Columns(COLUMN_NAME).Format = "##0.0"
                        Else
                            .Columns(COLUMN_NAME).Format = "#,##0"
                        End If

                        If COLUMN_NAME.StartsWith("STPCT") Then
                            Create_Summary(grd, COLUMN_NAME, "Custom")
                        Else
                            Create_Summary(grd, COLUMN_NAME)
                        End If
                        With .Columns(COLUMN_NAME).Header.Appearance
                            .BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                            .BackColor = Drawing.Color.White
                            If COLUMN_NAME.StartsWith("STAMT") Then
                                .BackColor2 = Drawing.Color.LightGreen
                            ElseIf COLUMN_NAME.StartsWith("STPCT") Then
                                .BackColor2 = Drawing.Color.SteelBlue
                            Else
                                .BackColor2 = Drawing.Color.LightBlue
                            End If
                        End With

                        With .Columns(COLUMN_NAME)
                            .Width = 60
                            If COLUMN_NAME.StartsWith("SIQTY") Then .Header.Caption = "In"
                            If COLUMN_NAME.StartsWith("OHQTY") Then .Header.Caption = "OH"
                            If COLUMN_NAME.StartsWith("STQTY") Then .Header.Caption = "Thru"
                            If COLUMN_NAME.StartsWith("STPCT") Then .Header.Caption = "SiSt%"
                            If COLUMN_NAME.StartsWith("STPCT") Then .Width = 50
                            If COLUMN_NAME.StartsWith("STAMT") Then .Header.Caption = "$Thru"
                            If COLUMN_NAME.StartsWith("STAMT") Then .CellAppearance.BackColor = Drawing.Color.LightGreen

                            If P = 0 Then
                                If COLUMN_NAME.StartsWith("SIQTY") Then .Header.ToolTipText = "Sell-In Units"
                                If COLUMN_NAME.StartsWith("OHQTY") Then .Header.ToolTipText = "End of Week On Hand Units"
                                If COLUMN_NAME.StartsWith("STQTY") Then .Header.ToolTipText = "Sell-Thru Units"
                                If COLUMN_NAME.StartsWith("STPCT") Then .Header.ToolTipText = "Sell-In / Sell-Thru %"
                                If COLUMN_NAME.StartsWith("STAMT") Then .Header.ToolTipText = "Sell-Thru $Retail"
                            End If
                        End With

                    Next
                Next

            End With
        Next
        'grdSATSIST2.DisplayLayout.Bands(0).ColHeadersVisible = False
        grdSATSIST2.DisplayLayout.Bands(0).GroupHeadersVisible = False

        With chtSATCSLS1
            .Axis.X.ScrollScale.Visible = True
            .Axis.Y.ScrollScale.Visible = True

            .Axis.X.ScrollScale.Scale = 1 ' 0.25
            .Axis.Y.ScrollScale.Scale = 1 ' 0.25
            Me.trkbrXAxis.Value = .Axis.X.ScrollScale.Scale * 100
            Me.trkbrYAxis.Value = .Axis.Y.ScrollScale.Scale * 100
            .EnableCrossHair = True

            '.ColorModel.ModelStyle = ColorModels.CustomLinear '  CType(System.Enum.Parse(GetType(ColorModels), System.Enum.GetNames(GetType(ColorModels))(0)), ColorModels)
        End With

        grpWEEK_RANGE.Top = grpPERIOD_RANGE.Top
        grpWEEK_RANGE.Left = grpPERIOD_RANGE.Left

        'CType(System.Enum.Parse(GetType(ColorModels), System.Enum.GetNames(GetType(ColorModels))(0)), ColorModels)
        'Dim modelStyle As String() = System.Enum.GetNames(GetType(ColorModels))
        'Dim s As String
        'For Each s In modelStyle
        '    Me.comboBox1.Items.Add(s)
        'Next s

        'Me.comboBox1.SelectedItem = Me.comboBox1.Items(Me.comboBox1.FindString(System.Enum.GetName(GetType(ColorModels), chtICTINVAT.ColorModel.ModelStyle), 0))

        'chtICTINVAT.ColorModel.ModelStyle = CType(System.Enum.Parse(GetType(ColorModels), Me.comboBox1.SelectedItem.ToString()), ColorModels)

        'Dim colors As Array = System.Enum.GetValues(GetType(Infragistics.UltraChart.Shared.Styles.ColorModels))
        'For i As Integer = 0 To colors.Length
        '    colors(i).ToString()
        'Next
        'cbeColor.DataSource = colors

        Dim modelStyle As String() = System.Enum.GetNames(GetType(ColorModels))
        cbeColor.DataSource = modelStyle
        cbeColor.SelectedItem = cbeColor.Items(cbeColor.FindString(System.Enum.GetName(GetType(ColorModels), chtSATCSLS1.ColorModel.ModelStyle), 0))

        'cbeColorBest.DataSource = System.Enum.GetNames(GetType(System.Drawing.Color))
        'cbeColorBest.SelectedItem = cbeColorBest.Items(cbeColor.FindString("Yellow", 0))
        'cbeColorWorst.DataSource = System.Enum.GetNames(GetType(System.Drawing.Color))
        'cbeColorWorst.SelectedItem = cbeColorBest.Items(cbeColor.FindString("Red", 0))

        Setup_Map()

        optSI.Tag = "*"

        chkExtendedData.Visible = False ' CLICK THIS OPTION, SELECT VONMAUR, THEN LOAD, THEN GET AN ERROR
        ASCMAIN1.Add_Value_List(grdSPTCOOPX, "APPR_STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'APPR_STATUS_CODE'")
        ASCMAIN1.Add_Value_List(grdSPTCOOPX, "STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'STATUS_CODE'")

        APPR_STATUS_CODE_BackColors.Add("A", System.Drawing.Color.Empty)
        APPR_STATUS_CODE_BackColors.Add("P", System.Drawing.Color.Empty)
        APPR_STATUS_CODE_BackColors.Add("G", System.Drawing.Color.Empty)
        APPR_STATUS_CODE_BackColors.Add("X", System.Drawing.Color.Empty)

        APPR_STATUS_CODE_ForeColors.Add("A", System.Drawing.Color.Green)
        APPR_STATUS_CODE_ForeColors.Add("P", System.Drawing.Color.Purple)
        APPR_STATUS_CODE_ForeColors.Add("G", System.Drawing.Color.Blue)
        APPR_STATUS_CODE_ForeColors.Add("X", System.Drawing.Color.Red)


        Dim YY As String = ""
        If Mid(ASCMAIN1.CYM, 5, 2) >= "02" And Mid(ASCMAIN1.CYM, 5, 2) <= "07" Then
            YY = Mid(ASCMAIN1.CYM, 3, 2)
            cmdS0.Text = "S" & YY
            cmdS1.Text = "F" & Format(Val(YY) - 1, "00")
            cmdS2.Text = "S" & Format(Val(YY) - 1, "00")
        Else
            YY = Mid(ASCMAIN1.CYM, 3, 2)
            If Mid(ASCMAIN1.CYM, 5, 2) = "01" Then YY = Format(Val(YY) - 1, "00")
            cmdS0.Text = "F" & YY
            cmdS1.Text = "S" & YY
            cmdS2.Text = "F" & Format(Val(YY) - 1, "00")
        End If



        grdASTWCAL1.DataSource = dst.Tables("ASTWCAL1")

        Dim CW As Integer = 25
        Dim WEEKDAYS As String = "SMTWRFS"

        With grdASTWCAL1.DisplayLayout.Bands("ASTWCAL1")

            .Groups("WEEK").Header.Appearance.TextHAlign = HAlign.Center
            .Groups("WEEK").Header.Caption = ""
            .Groups("WEEK").Width = CW * 2
            .Groups("WEEK").Header.Appearance.BackColor = Drawing.Color.DodgerBlue
            .Groups("WEEK").Header.Appearance.BackColor2 = Drawing.Color.Empty
            .Groups("WEEK").Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.None
            .Groups("WEEK").Header.Appearance.ForeColor = Drawing.Color.White

            .Columns("LINE_NO").Hidden = True

            .Columns("MONTH").Width = CW
            .Columns("MONTH").Header.Caption = "M"
            .Columns("MONTH").Header.Appearance.TextHAlign = HAlign.Center
            .Columns("MONTH").CellAppearance.TextHAlign = HAlign.Center
            .Columns("MONTH").Header.Appearance.ForeColor = Drawing.Color.DodgerBlue
            .Columns("MONTH").CellAppearance.ForeColor = Drawing.Color.DarkBlue

            .Columns("WEEK_NO").Width = CW
            .Columns("WEEK_NO").Header.Caption = "W"
            .Columns("WEEK_NO").Header.Appearance.TextHAlign = HAlign.Center
            .Columns("WEEK_NO").CellAppearance.TextHAlign = HAlign.Center

            .Columns("WEEK_NO").Header.Appearance.ForeColor = Drawing.Color.DodgerBlue
            .Columns("WEEK_NO").CellAppearance.BackColor = Drawing.Color.AliceBlue

            .Columns("WEEK_NO").CellAppearance.ForeColor = Drawing.Color.DarkBlue
            '.Columns("WEEK_NO").CellAppearance.BackColor = Drawing.Color.DodgerBlue
            '.Columns("WEEK_NO").CellAppearance.ForeColor = Drawing.Color.White

            For i As Integer = 1 To 5   ' 5 years
                Dim G As String = "Y" & Format(i, "0")
                .Groups(G).Header.Appearance.TextHAlign = HAlign.Center
                .Groups(G).Width = CW * 8.5
                .Groups(G).Header.Appearance.BackColor = Drawing.Color.DodgerBlue
                .Groups(G).Header.Appearance.BackColor2 = Drawing.Color.Empty
                .Groups(G).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.None
                .Groups(G).Header.Appearance.ForeColor = Drawing.Color.White


                With .Columns("DATES" & CStr(i))
                    .Width = CW * 2.5
                    .Header.Caption = "W/E"
                    .Header.Appearance.TextHAlign = HAlign.Center
                    .CellAppearance.TextHAlign = HAlign.Center
                    .Header.Appearance.BackColor2 = Drawing.Color.Gold
                    .Header.Appearance.BackColor = Drawing.Color.White
                    .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

                    .CellAppearance.BackColor = Drawing.Color.Yellow
                    .CellAppearance.BackColor2 = Drawing.Color.White
                    .CellAppearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassRight20
                End With


                For j As Integer = 1 To 7  ' 7 days in a week
                    Dim C As String = "DATE_" & Format(i, "0") & Format(j, "0")
                    .Columns(C).Width = CW
                    .Columns(C).Header.Caption = Mid(WEEKDAYS, j, 1)
                    .Columns(C).Header.Appearance.TextHAlign = HAlign.Center
                    .Columns(C).CellAppearance.TextHAlign = HAlign.Center
                    'If j = 7 Then
                    '    .Columns(C).CellAppearance.BackColor = Drawing.Color.Yellow
                    '    .Columns(C).CellAppearance.BackColor2 = Drawing.Color.White
                    '    .Columns(C).CellAppearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassRight20
                    'End If
                Next j
            Next i
        End With


        If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then
            cmdABS.Visible = True
        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Weekly Sales"
                Validate_Code("CUST_CODE")

            Case "View"
                Validate_Code("CUST_CODE")

                If EMsg = "" Then

                    If Trim(ASCMAIN1.USER_CODES) = "FS" Then
                        If cdr.Item("SREP_CODE") & "" <> TAC.TACMAIN1.SREP_CODE _
                                And Not TAC.TACMAIN1.SREP_CODEs.Contains(cdr.Item("SREP_CODE") & "") Then


                            Dim found_store As Boolean = False
                            ASCMAIN1.sql = "Select Distinct SREP_CODE from ARTCUST2 where CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                            ASCMAIN1.sql &= " UNION "
                            ASCMAIN1.sql &= "Select Distinct SELL_CODE from ARTCUST2 where CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                            For Each rowARTCUST2_SREP As DataRow In ASCDATA1.GetDataTable.Select("")
                                If rowARTCUST2_SREP.Item("SREP_CODE") & "" <> TAC.TACMAIN1.SREP_CODE _
                                    And Not TAC.TACMAIN1.SREP_CODEs.Contains(rowARTCUST2_SREP.Item("SREP_CODE") & "") Then
                                Else
                                    found_store = True
                                End If
                            Next

                            If Not found_store Then
                                EMsg &= vbCr & "Customer " & Absx1.txtFor("CUST_CODE").Text & " is not connected to Sales Rep code " & TAC.TACMAIN1.SREP_CODE
                            End If
                        End If
                    End If

                    Validate_Range(EMsg)


                    If EMsg = "" Then
                        If Absx1.optFor("RANGE").Value = "P" Then
                            If Absx1.cmbFor("RYP0").Value & "" = "" Then
                                EMsg &= vbCr & "You must Specify a Starting Period"
                            End If
                            If Absx1.cmbFor("RYP1").Value & "" = "" Then
                                EMsg &= vbCr & "You must Specify an Ending Period"
                            End If

                            If EMsg = "" Then
                                RYP0 = Absx1.cmbFor("RYP0").Value
                                RYP1 = Absx1.cmbFor("RYP1").Value
                                Periods = ASCMAIN1.Period_Diff(RYP0, RYP1) + 1

                                If Periods < 1 Then
                                    EMsg &= vbCr & "Invalid Period Range"
                                Else
                                    WPs.Clear()
                                    Ps.Clear()
                                    Ls.Clear()
                                    WPsLY.Clear()
                                    PsLY.Clear()
                                    LsLY.Clear()
                                    ReDim PE_DATEs(Periods)
                                    ReDim PE_DATEsLY(Periods)
                                    Dim rowGLTPARM2 As DataRow

                                    For P As Integer = 1 To Periods
                                        Dim YP As String = ASCMAIN1.Period_Calc(RYP0, P - 1)
                                        Ps.Add(YP)
                                        ASCMAIN1.sql = "Select Max (YYYYWW) from GLTPARM3 where YYYYPP = '" & YP & "'"
                                        Dim WYP As String = ASCDATA1.GetDataValue
                                        WPs.Add(WYP)
                                        Dim LEGEND As String = ASCMAIN1.Get_Legend(YP)
                                        Ls.Add(Mid(LEGEND, 1, 16))

                                        rowGLTPARM2 = LookUp("GLTPARM2", YP)
                                        PE_DATEs(P) = rowGLTPARM2.Item("PRD_END_DATE")

                                        Dim YPLY As String = ASCMAIN1.Period_Calc(YP, -12)
                                        PsLY.Add(YPLY)
                                        ASCMAIN1.sql = "Select Max (YYYYWW) from GLTPARM3 where YYYYPP = '" & YPLY & "'"
                                        Dim WYPLY As String = ASCDATA1.GetDataValue
                                        WPsLY.Add(WYPLY)
                                        Dim LEGEND_LY As String = ASCMAIN1.Get_Legend(YPLY)
                                        LsLY.Add(Mid(LEGEND_LY, 1, 16))

                                        rowGLTPARM2 = LookUp("GLTPARM2", YPLY)
                                        PE_DATEsLY(P) = rowGLTPARM2.Item("PRD_END_DATE")
                                    Next

                                    PE_DATEs(0) = LookUp("GLTPARM2", ASCMAIN1.Period_Calc(RYP0, -1)).Item("PRD_END_DATE")
                                    PE_DATEsLY(0) = LookUp("GLTPARM2", ASCMAIN1.Period_Calc(RYP0, -1 - 12)).Item("PRD_END_DATE")

                                End If


                            End If
                        Else
                            If Absx1.cmbFor("RYW0").Value & "" = "" Then
                                EMsg &= vbCr & "You must Specify a Starting Week"
                            End If
                            If Absx1.cmbFor("RYW1").Value & "" = "" Then
                                EMsg &= vbCr & "You must Specify an Ending Week"
                            End If

                            If EMsg = "" Then
                                needsEmptyWeek = False
                                Dim addedWeek As Boolean = False

                                Dim pOffset As Integer = 0

                                RYW0 = Absx1.cmbFor("RYW0").Value
                                RYW1 = Absx1.cmbFor("RYW1").Value
                                Periods = ASCMAIN1.Week_Diff(RYW0, RYW1) + 1
                                Dim RYW0_ly As String = Format(Val(Mid(RYW0, 1, 4)) - 1, "0000") & Mid(RYW0, 5, 2)
                                Dim RYW1_ly As String = Format(Val(Mid(RYW1, 1, 4)) - 1, "0000") & Mid(RYW1, 5, 2)
                                Dim Periods_LY As Integer = ASCMAIN1.Week_Diff(RYW0_ly, RYW1_ly) + 1

                                If Periods_LY > Periods Then
                                    needsEmptyWeek = True
                                    Periods = Periods_LY
                                End If

                                WPs.Clear()
                                Ps.Clear()
                                Ls.Clear()
                                WPsLY.Clear()
                                PsLY.Clear()
                                LsLY.Clear()
                                ReDim PE_DATEs(Periods)
                                ReDim PE_DATEsLY(Periods)
                                Dim rowGLTPARM3 As DataRow

                                For P As Integer = 1 To Periods
                                    Dim emptyWeek As Boolean = False

                                    Dim YW As String = ASCMAIN1.Week_Calc(RYW0, P - (1 + pOffset))
                                    Dim YW_orig As String = YW

                                    If ASCMAIN1.CLIENT = "INT" Then
                                        If needsEmptyWeek And pOffset = 0 Then
                                            If Mid(YW, 5, 2) = "01" Then
                                                YW = Format(Val(Mid(YW, 1, 4)) - 1, "0000") & "53"
                                                pOffset = 1
                                                emptyWeek = True
                                            End If
                                        End If
                                    End If


                                    Ps.Add(YW)
                                    WPs.Add(YW)

                                    rowGLTPARM3 = LookUp("GLTPARM3", YW_orig)

                                    If Not emptyWeek Then
                                        Dim WEEK_END_DATE As Date = rowGLTPARM3.Item("WEEK_END_DATE")
                                        Dim LEGEND As String = rowGLTPARM3.Item("LEGEND") & " (" & Format(WEEK_END_DATE.AddDays(-6), "MM/dd") & "-" & Format(WEEK_END_DATE, "MM/dd") & ")"
                                        Ls.Add(LEGEND)
                                    Else
                                        Dim LEGEND As String = "" ' rowGLTPARM3.Item("LEGEND") & " (" & Format(WEEK_END_DATE_LY.AddDays(-6), "MM/dd") & "-" & Format(WEEK_END_DATE_LY, "MM/dd") & ")"
                                        Ls.Add(LEGEND)
                                    End If

                                    PE_DATEs(P) = rowGLTPARM3.Item("WEEK_END_DATE")

                                    Dim YWLY As String = ASCMAIN1.Week_Calc(YW_orig, -52)

                                    If ASCMAIN1.CLIENT = "INT" Then
                                        YWLY = Format(Val(Mid(YW, 1, 4)) - 1, "0000") & Mid(YW, 5, 2)

                                    End If
                                    PsLY.Add(YWLY)
                                    WPsLY.Add(YWLY)
                                    If ASCMAIN1.CLIENT = "INT" And Mid(YW, 5, 2) = "53" Then
                                        'rowGLTPARM3 = LookUp("GLTPARM3", YWLY)
                                        'Dim WEEK_END_DATE_LY As Date = rowGLTPARM3.Item("WEEK_END_DATE")
                                        If Not emptyWeek Then
                                            Dim LEGEND_LY As String = "" ' rowGLTPARM3.Item("LEGEND") & " (" & Format(WEEK_END_DATE_LY.AddDays(-6), "MM/dd") & "-" & Format(WEEK_END_DATE_LY, "MM/dd") & ")"
                                            LsLY.Add(LEGEND_LY)
                                        Else
                                            rowGLTPARM3 = LookUp("GLTPARM3", YWLY)
                                            Dim WEEK_END_DATE_LY As Date = rowGLTPARM3.Item("WEEK_END_DATE")
                                            Dim LEGEND_LY As String = rowGLTPARM3.Item("LEGEND") & " (" & Format(WEEK_END_DATE_LY.AddDays(-6), "MM/dd") & "-" & Format(WEEK_END_DATE_LY, "MM/dd") & ")"
                                            LsLY.Add(LEGEND_LY)
                                        End If
                                    Else
                                        rowGLTPARM3 = LookUp("GLTPARM3", YWLY)
                                        Dim WEEK_END_DATE_LY As Date = rowGLTPARM3.Item("WEEK_END_DATE")
                                        Dim LEGEND_LY As String = rowGLTPARM3.Item("LEGEND") & " (" & Format(WEEK_END_DATE_LY.AddDays(-6), "MM/dd") & "-" & Format(WEEK_END_DATE_LY, "MM/dd") & ")"
                                        LsLY.Add(LEGEND_LY)
                                    End If

                                    PE_DATEsLY(P) = rowGLTPARM3.Item("WEEK_END_DATE")
                                    If emptyWeek Then
                                        emptyWeek = False
                                    End If
                                Next

                                PE_DATEs(0) = PE_DATEs(1).AddDays(-7)
                                PE_DATEsLY(0) = PE_DATEsLY(1).AddDays(-7)
                            End If
                        End If

                        If Periods < 1 Or Periods > Stores_Max Or (chkSIST.Checked And Periods > 52) Then
                            EMsg &= vbCr & "Total number of Periods must be between 1 and " & CStr(IIf(chkSIST.Checked, 52, Stores_Max))
                        End If

                    End If


                End If

            Case "Load Summary"

                Validate_Range(EMsg)

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Weekly Sales"
                EntryMode = "W"
                Weekly_Sales_Heat_Map()
                Mode_Settings(True)

            Case "View"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Print Report"
                Print_Report()

            Case "Load Summary"

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading Sales Summary by Customer")

                Create_SATCSLS1()
                optXP.Items(0).DisplayText = optRANGE.Items(optRANGE.CheckedIndex).DisplayText
                Load_Data()
                grdSATCSLS1.Parent = grpSummary

                grdSATCSLS1.Text = "Sales Summary by Customer"
                With grdSATCSLS1.DisplayLayout.Bands(0)
                    .Columns("CODE_VALUE").Header.Caption = "Customer"
                    .Columns("DESC_VALUE").Header.Caption = "Customer Name"
                End With

                UltraExplorerBar1.Groups("Data Options").Visible = True
                optSI.Visible = False
                optXP.Visible = False
                chkShowDetails.Visible = False
                summary_loaded = True

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

            Case "Multi-Band Export"

                dst.Tables("SATCSLS2").Rows.Clear()
                dst.Tables("SOTINVHX").Rows.Clear()

                Dim DATA_TYPE As String = optType1.Value & optType2.Value
                Dim TOPN As Integer = Val(numN.Value & "")
                Dim TOPX As Integer = 0
                For Each grow As UltraWinGrid.UltraGridRow In grdSATCSLS1.Rows
                    TOPX += 1
                    If TOPN = 0 Or (TOPX <= TOPN Or TOPX > (grdSATCSLS1.Rows.Count - TOPN)) Then
                        Dim CODE_VALUE_PARENT As String = grow.Cells("CODE_VALUE").Value
                        Load_SATCSLS2(DATA_TYPE, CODE_VALUE_PARENT, False, False)
                    Else
                        grow.Hidden = True
                    End If
                Next

                grdSATCSLS1.DisplayLayout.Bands("SATCSLS1_SATCSLS1_DTL").Hidden = True
                grdSATCSLS1.DisplayLayout.Bands("SATCSLS1_SATCSLS2").Hidden = False

                Dim b As UltraWinGrid.UltraGridBand = grdSATCSLS1.DisplayLayout.Bands("SATCSLS1_SATCSLS2")
                With grdSATCSLS2.DisplayLayout.Bands(0)
                    For Each grcol As UltraWinGrid.UltraGridColumn In b.Columns
                        Dim COLUMN_NAME As String = grcol.Key
                        grcol.Header.Caption = .Columns(COLUMN_NAME).Header.Caption
                        grcol.Hidden = .Columns(COLUMN_NAME).Hidden
                    Next
                    For I As Integer = 5 To 1 Step -1
                        Dim COLUMN_NAME As String = "SUB_CODE_VALUE" & Format(I, "0")
                        If Not .Columns(COLUMN_NAME).Hidden Then
                            b.Columns(COLUMN_NAME).Header.VisiblePosition = b.Columns.Count - 1
                            Exit For
                        End If
                    Next
                End With

                Try
                    Export_to_Excel(grdSATCSLS1)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error with Multi-Band Export")
                End Try

                For Each grow As UltraWinGrid.UltraGridRow In grdSATCSLS1.Rows
                    grow.Hidden = False
                Next

                grdSATCSLS1.DisplayLayout.Bands("SATCSLS1_SATCSLS2").Hidden = True
                grdSATCSLS1.DisplayLayout.Bands("SATCSLS1_SATCSLS1_DTL").Hidden = Not chkExtendedData.Checked
                dst.Tables("SATCSLS2").Rows.Clear()

                Setup_grdSATCSLS2()

            Case "Export XLS"
                XLS_FILENAME_multi = ""
                Excel_Extract()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Load Summary").Settings.Enabled = not_iScreenMode
                    .Items("Load Summary").Visible = Not (Trim(ASCMAIN1.USER_CODES) = "FS") And (EntryMode <> "W")
                    .Items("Multi-Band Export").Visible = tf And Not chkSIST.Checked And (EntryMode <> "W")
                    .Items("Export XLS").Visible = tf And chkSIST.Checked

                    .Items("Weekly Sales").Visible = Not ScreenMode

                End With
                .Groups("Data Options").Visible = tf And Not chkSIST.Checked And (EntryMode <> "W")
                .Groups("Options").Visible = Not tf
                .Groups("Charts").Visible = False
                .Groups("Sell-In / Sell-Thru Options").Visible = Not tf And chkSIST.Checked
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        SplitContainer1.Visible = tf And Not (chkSIST.Checked) And (EntryMode <> "W")
        grpSummary.Visible = Not tf
        splSATSIST0.Visible = Not tf
        splSATSIST1.Visible = tf And chkSIST.Checked And (EntryMode <> "W")

        lblN.Visible = tf And (EntryMode <> "W")
        numN.Visible = tf And (EntryMode <> "W")

        cmdS0.Visible = Not ScreenMode
        cmdS1.Visible = Not ScreenMode
        cmdS2.Visible = Not ScreenMode

        grdASTWCAL1.Visible = tf And (EntryMode = "W")

        'lblInfo.Visible = ScreenMode
        'lblInfo2.Visible = ScreenMode

        If ScreenMode Then
            If EntryMode = "W" Then
            Else
                grdSATCSLS1.Parent = SplitContainer1.Panel1
                optSI.Visible = True
                optXP.Visible = True
                chkShowDetails.Visible = True
            End If


            If (EntryMode = "W") Then
                splTrend.Parent = splCharts.Panel1
                splMap.Parent = splCharts.Panel2
            Else
                splTrend.Parent = tabDetails.Tabs("Charts").TabPage
                splMap.Parent = tabDetails.Tabs("Map").TabPage
            End If

        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SATCSLS1", "SATCSLS1_DTL", "SATCSLS2", "EDT852T1", "SOTINVHX", "SATSIST1", "SATSIST2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)


        If SATSIST1 <> "" Then
            ' each query may generate a different number of columns
            ASCDATA1.ExecuteSQL("Drop Table " & SATSIST1)
            ASCDATA1.ExecuteSQL("Drop Table " & SATSIST2)
            ASCDATA1.ExecuteSQL("Drop Table " & SATSIST3)
            SATSIST1 = ""
            SATSIST2 = ""
            SATSIST3 = ""
        End If

        'Absx1.txtFor("CUST_CODE").Text = ""

        tabDetails.SelectedTab = tabDetails.Tabs("Details")

        Setup_SM()
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Customer Sales Data")
        Save_Header_Fields(UltraGroupBox1)

        RYP0 = Absx1.cmbFor("RYP0").Value
        RYP1 = Absx1.cmbFor("RYP1").Value
        RYW0 = Absx1.cmbFor("RYW0").Value
        RYW1 = Absx1.cmbFor("RYW1").Value

        CUST_CODE = HFs("CUST_CODE")

        rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
        ' COLLECTION_CODE = rowICTITEM1.Item("COLLECTION_CODE") & ""


        'ITEM_CODE_LY = ITEM_CODE
        'ASCMAIN1.sql = "Select * from SOTALLO1" & vbCrLf _
        '    & " where ITEM_CODE = '" & ITEM_CODE & "'" & vbCrLf _
        '    & "   and ITEM_CODE_COMPARE_TO is Not Null" & vbCrLf _
        '    & "   and DATE_START between '" & Format(PE_DATEs(0).AddDays(1), "dd-MMM-yyyy") & "' and '" & Format(PE_DATEs(Periods), "dd-MMM-yyyy") & "'"

        'QTY_ALLO_PLAN = 0
        'Dim tblAllo As DataTable = ASCDATA1.GetDataTable
        'If tblAllo.Rows.Count <> 0 Then
        '    ITEM_CODE_LY = tblAllo.Rows(0).Item("ITEM_CODE_COMPARE_TO")
        '    QTY_ALLO_PLAN = Val(tblAllo.Rows(0).Item("QTY_ALLO_PLAN") & "")
        'End If


        ' lblInfo.Text = "Collection: " & COLLECTION_CODE

        ' WHAT ABOUT LAUNCH_DATE? =
        'Dim ITEM_DATE_TO_SHIP As String = ""
        'If rowICTITEM1.Item("ITEM_DATE_TO_SHIP") & "" <> "" Then
        '    ITEM_DATE_TO_SHIP = Format(rowICTITEM1.Item("ITEM_DATE_TO_SHIP"), "MM/dd/yyyy")
        'End If
        'Dim ITEM_RETAIL_PRICE As Decimal = Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & "")

        'lblInfo2.Text = "Retail: " & Format(ITEM_RETAIL_PRICE, "$#,##0.00") & IIf(ITEM_DATE_TO_SHIP = "", "", "; Ship: " & ITEM_DATE_TO_SHIP)
        'If QTY_ALLO_PLAN <> 0 Then lblInfo2.Text &= vbCrLf & "Plan: " & Format(QTY_ALLO_PLAN, "#,##0") & " units, " & Format(QTY_ALLO_PLAN * ITEM_RETAIL_PRICE, "$#,##0")

        If chkSIST.Checked Then
            Load_SIST()

            ' get 1st item in dataset, then get that item's collection
            ' probably ought to get all collections 

            Dim ITEM_CODE As String = ""
            Dim rowICTITEM1 As DataRow = Nothing
            Dim COLLECTION_CODE As String = ""

            If dst.Tables("SATSIST1").Rows.Count > 0 Then
                Dim row As DataRow = dst.Tables("SATSIST1").Rows(0)
                ITEM_CODE = row.Item("ITEM_CODE")
                rowICTITEM1 = LookUp("ICTITEM1", ITEM_CODE)
                COLLECTION_CODE = rowICTITEM1.Item("COLLECTION_CODE")
            End If

            Dim sqlw As String = ""
            Dim sqlf As String = ""
            Dim SYP(1) As String
            If optRANGE.Value = "P" Then
                sqlw = " and SPTCOOP1.OPS_YYYYPP >= '" & Ps(0) & "' and SPTCOOP1.OPS_YYYYPP <= '" & Ps(Periods - 1) & "'"
                SYP(0) = ASCMAIN1.Get_Legend(RYP0)
                SYP(1) = ASCMAIN1.Get_Legend(RYP1)
                sqlf = "OPS_YYYYPP >= '" & RYP0 & "'"
            Else
                sqlw = " and SPTCOOP1.OPS_YYYYWW >= '" & Ps(0) & "' and SPTCOOP1.OPS_YYYYWW <= '" & Ps(Periods - 1) & "'"
                SYP(0) = ASCMAIN1.Get_Legend_Wk(RYW0)
                SYP(1) = ASCMAIN1.Get_Legend_Wk(RYW1)
                sqlf = "OPS_YYYYWW >= '" & RYW0 & "'"
            End If
            ASCMAIN1.sql = sqlSPTCOOPX & sqlw & " and SPTCOOP1.CUST_CODE = '" & CUST_CODE & "'"
            Fill_Records("SPTCOOPX", "", True, ASCMAIN1.sql)

            If optRANGE.Value = "P" Then
                sqlw = " and SPTCOOP1.OPS_YYYYPP >= '" & PsLY(0) & "' and SPTCOOP1.OPS_YYYYPP <= '" & PsLY(Periods - 1) & "'"
            Else
                sqlw = " and SPTCOOP1.OPS_YYYYWW >= '" & PsLY(0) & "' and SPTCOOP1.OPS_YYYYWW <= '" & PsLY(Periods - 1) & "'"
            End If
            ASCMAIN1.sql = sqlSPTCOOPX & sqlw & " and SPTCOOP1.CUST_CODE = '" & CUST_CODE & "'"
            Fill_Records("SPTCOOPX", "", False, ASCMAIN1.sql)


            dst.Tables("SATSIST1").Columns("OHQTY_P00").Expression = "ISNULL(OHQTY_P" & Format(Periods, "00") & ",0)"
            dst.Tables("SATSIST2").Columns("OHQTY_P00").Expression = "ISNULL(OHQTY_P" & Format(Periods, "00") & ",0)"
            dst.Tables("SATSIST3").Columns("OHQTY_P00").Expression = "ISNULL(OHQTY_P" & Format(Periods, "00") & ",0)"

            Dim dvw As DataView = DirectCast(grdSPTCOOPX.DataSource, DataTable).DefaultView
            dvw.RowFilter = sqlf

            Sort_grdColumns(grdSPTCOOPX, "AUTH_NO".ToLower)
            grdSPTCOOPX.Text = "Sales Promotions for " & CUST_CODE & " Starting between " & SYP(0) & " and " & SYP(1)

        Else
            Create_SATCSLS1()
            optXP.Items(0).DisplayText = optRANGE.Items(optRANGE.CheckedIndex).DisplayText
            Load_Data()
        End If

        ASCMAIN1.Progress("")

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "CUST_CODE"

                'If Trim(ASCMAIN1.USER_CODES) = "FS" Then
                '    If TAC.TACMAIN1.SREP_CODEs.Count > 1 Then
                '        sql_where = " and ARTCUST1.SREP_CODE in ('" & Join(TAC.TACMAIN1.SREP_CODEs.ToArray, "','") & "')"
                '    Else
                '        sql_where = " and ARTCUST1.SREP_CODE = '" & TAC.TACMAIN1.SREP_CODE & "'"
                '    End If
                'End If
                If Trim(ASCMAIN1.USER_CODES) = "FS" Then
                    If TAC.TACMAIN1.SREP_CODEs.Count > 1 Then
                        sql_where = " and (ARTCUST1.SREP_CODE in ('" & Join(TAC.TACMAIN1.SREP_CODEs.ToArray, "','") & "')" _
                            & " or ARTCUST1.CUST_CODE in (Select Distinct CUST_CODE from ARTCUST2 where ARTCUST2.SREP_CODE in ('" & Join(TAC.TACMAIN1.SREP_CODEs.ToArray, "','") & "')))"
                    Else
                        sql_where = " and (ARTCUST1.SREP_CODE = '" & TAC.TACMAIN1.SREP_CODE & "'" _
                            & " or ARTCUST1.CUST_CODE in (Select Distinct CUST_CODE from ARTCUST2 where ARTCUST2.SREP_CODE = '" & TAC.TACMAIN1.SREP_CODE & "'))"
                    End If
                End If

        End Select
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSATCSLS1, "SSSSSSSSS", "Show Filter", "Show GroupBox", "Show SUB_CODE_VALUE1", "Show SUB_CODE_VALUE2", "Show SUB_CODE_VALUE3", "Show SUB_CODE_VALUE4", "Show SUB_CODE_VALUE5", "Show RTL_PRICE", "Show WSL_PRICE")
        Load_Popup_Menu(grdSATCSLS2, "SSSSSSSSS", "Show Filter", "Show GroupBox", "Show SUB_CODE_VALUE1", "Show SUB_CODE_VALUE2", "Show SUB_CODE_VALUE3", "Show SUB_CODE_VALUE4", "Show SUB_CODE_VALUE5", "Show RTL_PRICE", "Show WSL_PRICE")
        Load_Popup_Menu(grdEDT852T1, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTINVHX, "SSBB", "Show Filter", "Show GroupBox", "Sales Order Inquiry", "Show Invoice")
        Load_Popup_Menu(grdSATCSLSS, "CC", "Best", "Worst")
        Load_Popup_Menu(grdSATSIST1, "SBBB", "Show Filter", "Customer Order Inquiry", "Customer Sales Summary", "Export to Excel")
        Load_Popup_Menu(grdSPTCOOPX, "SB", "Show Filter", "Sales Promotion Inquiry")
        Load_Popup_Menu(grdSATSISTI, "BB", "Clear All", "Add Items")
        Load_Popup_Menu(grdSATSISTC, "BB", "Clear All", "Add Customers")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool

        If tlb_pop.Tools.Exists("Show SUB_CODE_VALUE1") Then
            For I As Integer = 1 To 5
                Dim COLUMN_NAME As String = "SUB_CODE_VALUE" & CStr(I)
                tlb_sbt = DirectCast(tlb_pop.Tools("Show " & COLUMN_NAME), UltraWinToolbars.StateButtonTool)
                tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden
                tlb_sbt.SharedProps.Caption = "Show " & grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption
            Next

            COLUMN_NAME = "RTL_PRICE"
            tlb_sbt = DirectCast(tlb_pop.Tools("Show " & COLUMN_NAME), UltraWinToolbars.StateButtonTool)
            tlb_sbt.SharedProps.Visible = (EntryMode = "E")
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden
            tlb_sbt.SharedProps.Caption = "Show " & grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption
            COLUMN_NAME = "WSL_PRICE"
            tlb_sbt = DirectCast(tlb_pop.Tools("Show " & COLUMN_NAME), UltraWinToolbars.StateButtonTool)
            tlb_sbt.SharedProps.Visible = (EntryMode = "E")
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden
            tlb_sbt.SharedProps.Caption = "Show " & grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption
        End If


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            '  e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdSATCSLS1"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "Clear All"
                Dim tbl As DataTable = DirectCast(grd.DataSource, DataTable)
                tbl.Rows.Clear()

            Case "Add Items"
                add_items_Select()


            Case "Add Customers"
                Add_Customers_Select()

            Case "Show SUB_CODE_VALUE1", "Show SUB_CODE_VALUE2", "Show SUB_CODE_VALUE3", "Show SUB_CODE_VALUE4", "Show SUB_CODE_VALUE5", "Show RTL_PRICE", "Show WSL_PRICE"

                Dim COLUMN_NAME As String = Mid(e.Tool.Key, 6)
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = Not tlb_sbt.Checked
                If grdSATCSLS1.DisplayLayout.Bands.Count > 1 Then
                    With grdSATCSLS1.DisplayLayout.Bands(1).Columns("DATA_TYPE")
                        If tlb_sbt.Checked Then
                            .ColSpan += 1
                        Else
                            .ColSpan -= 1
                        End If
                    End With
                End If

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Best"
                Dim tlb_cpt As UltraWinToolbars.PopupColorPickerTool = DirectCast(e.Tool, UltraWinToolbars.PopupColorPickerTool)
                tlb_cpt.ReplaceableColor = Me.UltraChart1.ColorModel.ColorEnd

            Case "Worst"
                Dim tlb_cpt As UltraWinToolbars.PopupColorPickerTool = DirectCast(e.Tool, UltraWinToolbars.PopupColorPickerTool)
                tlb_cpt.ReplaceableColor = Me.UltraChart1.ColorModel.ColorBegin


            Case "Sales Order Inquiry"
                Dim INV_NO As String = grd.ActiveRow.Cells("INV_NO").Value
                Dim rowSOTINVH1 As DataRow = LookUp("SOTINVH1", New String() {"I", INV_NO})
                If rowSOTINVH1 IsNot Nothing Then
                    Dim ORDR_NO As String = rowSOTINVH1.Item("ORDR_NO") & ""
                    Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                    If rowSOTORDR1 IsNot Nothing Then
                        Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                    End If
                End If

            Case "Sales Promotion Inquiry"
                Dim AUTH_NO As String = grd.ActiveRow.Cells("AUTH_NO").Value
                Dim rowSPTCOOP1 As DataRow = LookUp("SPTCOOP1", AUTH_NO)
                If rowSPTCOOP1 IsNot Nothing Then
                    Context_Launch("View", AUTH_NO, e.Tool.Key, "SPFCOOPI")
                End If

            Case "Customer Order Inquiry"
                Dim CUST_CODE As String = ""
                If grd.Name = "grdSATISLS1" Then
                    CUST_CODE = grd.ActiveRow.Cells("CODE_VALUE").Value
                ElseIf grd.Name = "grdSATISLS2" Then
                    CUST_CODE = grd.ActiveRow.Cells("CODE_VALUE_PARENT").Value
                ElseIf grd.Name = "grdSOTINVHX" Or grd.Name = "grdSATSIST1" Then
                    CUST_CODE = grd.ActiveRow.Cells("CUST_CODE").Value
                End If
                Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                If rowARTCUST1 IsNot Nothing Then
                    Context_Launch("Select", CUST_CODE, e.Tool.Key, "SOFCORD1")
                End If

            Case "Show Invoice"
                Dim FILENAME As String = ""
                If grd.ActiveRow IsNot Nothing Then
                    If Not grd.ActiveRow.Selected Then
                        grd.Selected.Rows.Clear()
                        grd.ActiveRow.Selected = True
                    End If
                End If

                If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.Selected Then
                    Exit Sub
                End If

                'Dim INV_TYPE As String = grd.ActiveRow.Cells("INV_TYPE").Value & ""
                Dim INV_NO As String = grd.ActiveRow.Cells("INV_NO").Value & ""

                'If INV_TYPE <> "I" And INV_TYPE <> "C" Then
                '    Exit Sub
                'End If
                'FILENAME = TAC.SOCMAIN1.Create_Invoice(Me, INV_NO)

                'Show_Document(FILENAME)

        End Select
    End Sub

    Public Overrides Sub tlb_ToolValueChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolEventArgs)
        MyBase.tlb_ToolValueChanged(sender, e)

        Select Case e.Tool.Key
            Case "Best"
                Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
                = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
                Me.UltraChart1.ColorModel.ColorEnd = tlb_cpt.SelectedColor
                UltraChart1.DataBind()
                'grdSATCSLSS.DataBind()
                Application.DoEvents()
                grdSATCSLSS.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)

            Case "Worst"
                Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
                = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
                Me.UltraChart1.ColorModel.ColorBegin = tlb_cpt.SelectedColor
                UltraChart1.DataBind()
                'grdSATCSLSS.DataBind()
                Application.DoEvents()
                grdSATCSLSS.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)
        End Select

    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    '    Click_Command("View", e)
                End If
        End Select


    End Sub

    Public Overrides Sub cmb_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.cmb_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        If COLUMN_NAME = "RYP0" Or COLUMN_NAME = "RYP1" Then
            If Absx1.cmbFor("RYP0").Value & "" <> "" And Absx1.cmbFor("RYP1").Value & "" <> "" Then
                Dim C As Integer = 1 + ASCMAIN1.Period_Diff(Absx1.cmbFor("RYP0").Value, Absx1.cmbFor("RYP1").Value)
                lblMonths.Text = CStr(C) & " Mos"
            Else
                lblMonths.Text = ""
            End If
        Else
            If Absx1.cmbFor("RYW0").Value & "" <> "" And Absx1.cmbFor("RYW1").Value & "" <> "" Then
                Dim C As Integer = 1 + ASCMAIN1.Week_Diff(Absx1.cmbFor("RYW0").Value, Absx1.cmbFor("RYW1").Value)
                lblWeeks.Text = CStr(C) & " Wks"
            Else
                lblWeeks.Text = ""
            End If
        End If
    End Sub
#End Region

    Sub Create_SATCSLS1()

        If SATCSLS2 = "" Then
            SATCSLS1 = ASCMAIN1.Temp_Table("Select CUST_CODE from ARTCUST1 where ROWNUM < 1")
            SATCSLS2 = ASCMAIN1.Temp_Table("Select CUST_CODE from ARTCUST1 where ROWNUM < 1")
        End If
        ASCDATA1.ExecuteSQL("Drop Table " & SATCSLS1)
        ASCDATA1.ExecuteSQL("Drop Table " & SATCSLS2)

        Dim PX As String = IIf(optRANGE.Value = "P", "OPS_YYYYPP", "OPS_YYYYWW")
        Dim PX0 As String = IIf(optRANGE.Value = "P", RYP0, RYW0)
        Dim PX1 As String = IIf(optRANGE.Value = "P", RYP1, RYW1)

        If EntryMode = "E" Then
            ASCMAIN1.sql = "Select EDT852T1.* from EDT852T1" _
            & " where EDT852T1.EDI_STATUS = '1' " _
            & " and EDT852T1." & PX & " between '" & PX0 & "' and '" & PX1 & "'" _
            & " and EDT852T1.CUST_CODE = '" & CUST_CODE & "'"
            Fill_Records("EDT852T1", "", True, ASCMAIN1.sql)
            grdEDT852T1.Text = "EDI Documents Imported and Loaded for " & CUST_CODE
            grdEDT852T1.DisplayLayout.CaptionVisible = DefaultableBoolean.True
        End If

        sqlSATCSLS1 = ""
        sqlSATCSLS1_sum = ""
        sqlSATCSLS2 = ""
        Dim SQL() As String = New String() {"", ""}
        Dim P As Integer
        PERIODS_XX = ""

        'SQL = ""
        For P = 1 To Periods
            Dim PXP As String = IIf(optRANGE.Value = "P", ASCMAIN1.Period_Calc(RYP0, P - 1), ASCMAIN1.Week_Calc(RYW0, P - 1))
            SQL(0) &= ", Sum (Decode(" & PX & ",'" & PXP & "',RSTRETL1.QTY_SOLD,0)) P" & Format(P, "00") & vbCrLf
            SQL(1) &= ", Sum (Decode(" & PX & ",'" & ASCMAIN1.Period_Calc(PXP, -12) & "',RSTRETL1.QTY_SOLD,0)) P" & Format(P, "00") & vbCrLf
            sqlSATCSLS1 &= ", P" & Format(P, "00")
            sqlSATCSLS1_sum &= ", Sum (P" & Format(P, "00") & ") P" & Format(P, "00")
            sqlSATCSLS2 &= ", SUM (P" & Format(P, "00") & ") P" & Format(P, "00")
            PERIODS_XX &= "+P" & Format(P, "00")
        Next

        Dim YEAR_max As Int32 = 0
        If EntryMode = "E" Then
            If chkPriorYear.Checked Then
                YEAR_max = 1
            End If
        End If

        For YEAR As Int32 = 0 To YEAR_max

            PX = IIf(optRANGE.Value = "P", "OPS_YYYYPP", "OPS_YYYYWW")

            PX0 = IIf(optRANGE.Value = "P", RYP0, RYW0)
            PX1 = IIf(optRANGE.Value = "P", RYP1, RYW1)
            If YEAR = 1 Then
                If optRANGE.Value = "P" Then
                    PX0 = ASCMAIN1.Period_Calc(PX0, -12)
                    PX1 = ASCMAIN1.Period_Calc(PX1, -12)
                Else
                    PX0 = ASCMAIN1.Week_Calc(PX0, -52)
                    PX1 = ASCMAIN1.Week_Calc(PX1, -52)

                    If ASCMAIN1.CLIENT = "INT" Then
                        PX0 = Format(Val(Mid(PX0, 1, 4)) - 1, "0000") & Mid(PX0, 5, 2)
                        PX1 = Format(Val(Mid(PX1, 1, 4)) - 1, "0000") & Mid(PX1, 5, 2)
                    End If
                End If
            End If


            Dim sqla As String = "Select 'TU' DATA_TYPE, '" & CStr(YEAR) & "' YEAR" & vbCrLf _
            & ", RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, RSTRETL1.ITEM_CODE" & vbCrLf _
            & SQL(YEAR) & vbCrLf _
            & " from RSTRETL1,GLTPARM3  " & vbCrLf _
            & " where GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW" & vbCrLf _
            & IIf(EntryMode = "E", " and RSTRETL1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf, "") _
            & " and RSTRETL1." & PX & " Between '" & PX0 & "' and '" & PX1 & "'" & vbCrLf _
            & " group by RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, RSTRETL1.ITEM_CODE"

            Dim SQL_ORIG As String = sqla

            If YEAR = 0 Then
                ASCDATA1.ExecuteSQL("Create Table " & SATCSLS2 & " as " & sqla)
                ASCDATA1.ExecuteSQL("Alter Table " & SATCSLS2 _
                & " Add Primary Key (DATA_TYPE, YEAR, CUST_CODE, CUST_STORE_NO, ITEM_CODE)")
            Else
                ASCDATA1.ExecuteSQL("Insert into " & SATCSLS2 & " " & sqla)
            End If
            sqla = Replace(sqla, "'TU' DATA_TYPE", "'TD' DATA_TYPE")
            sqla = Replace(sqla, "QTY_SOLD", "AMT_SOLD")
            ASCDATA1.ExecuteSQL("Insert into " & SATCSLS2 & " " & sqla)

            sqla = Replace(SQL_ORIG, "'TU' DATA_TYPE", "'HU' DATA_TYPE")
            sqla = Replace(sqla, " and GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW", " and GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW" & vbCrLf & IIf(optRANGE.Value = "P", " and GLTPARM3.REL_WEEK = GLTPARM3.MAX_WEEK", ""))
            If optRANGE.Value = "P" Then
                sqla = Replace(sqla, "RSTRETL1.QTY_SOLD", "CASE WHEN GLTPARM3.REL_WEEK = GLTPARM3.MAX_WEEK THEN RSTRETL1.QTY_EOW ELSE 0 END")
            Else
                sqla = Replace(sqla, "QTY_SOLD", "QTY_EOW")
            End If
            ASCDATA1.ExecuteSQL("Insert into " & SATCSLS2 & " " & sqla)
            '     & ", SUM(DECODE(OPS_YYYYWW,'" & WPs(P - 1) & "',QTY_EOW,0)) OHQTY" & PX _


            sqla = Replace(SQL_ORIG, "'TU' DATA_TYPE", "'HD' DATA_TYPE")
            sqla = Replace(sqla, " and GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW", " and GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW" & vbCrLf & IIf(optRANGE.Value = "P", " and GLTPARM3.REL_WEEK = GLTPARM3.MAX_WEEK", ""))
            sqla = Replace(sqla, "from RSTRETL1", "from RSTRETL1,ICTITEM1")
            sqla = Replace(sqla, "group by", " and ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE group by")
            If optRANGE.Value = "P" Then
                sqla = Replace(sqla, "RSTRETL1.QTY_SOLD", "CASE WHEN GLTPARM3.REL_WEEK = GLTPARM3.MAX_WEEK THEN RSTRETL1.QTY_EOW * ICTITEM1.ITEM_RETAIL_PRICE ELSE 0 END") ' SB USING ICTRETLA
            Else
                sqla = Replace(sqla, "QTY_SOLD", "QTY_EOW * ICTITEM1.ITEM_RETAIL_PRICE") ' SB USING ICTRETLA
            End If
            ASCDATA1.ExecuteSQL("Insert into " & SATCSLS2 & " " & sqla)

            PX = IIf(optRANGE.Value = "P", "ORDR_YYYYPP_UPDATED", "OPS_YYYYWW")
            sqla = ""
            For P = 1 To Periods
                Dim PXP As String = IIf(optRANGE.Value = "P", _
                                       ASCMAIN1.Period_Calc(RYP0, P - 1 - 12 * YEAR), _
                                       ASCMAIN1.Week_Calc(RYW0, P - 1 - 52 * YEAR))
                sqla &= ", Sum (Decode(" & PX & ",'" & PXP & "',SOTINVH2.ORDR_QTY_SHIP,0)) P" & Format(P, "00") & vbCrLf
            Next
            sqla = "Select DECODE(INV_TYPE,'I','S','C','R') || 'U' DATA_TYPE" & vbCrLf _
            & ", '" & CStr(YEAR) & "' YEAR" & vbCrLf _
            & ", CUST_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
            & sqla & vbCrLf _
            & " from SOTINVH2 " & vbCrLf _
            & " where " & PX & " Between '" & PX0 & "' and '" & PX1 & "'" & vbCrLf _
            & IIf(EntryMode = "E", " and SOTINVH2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf, "") _
            & " group by DECODE(INV_TYPE,'I','S','C','R') || 'U', CUST_CODE, CUST_STORE_NO, ITEM_CODE"

            ASCDATA1.ExecuteSQL("Insert into " & SATCSLS2 & " " & sqla)
            sqla = Replace(sqla, "'U' DATA_TYPE", "'D' DATA_TYPE")
            sqla = Replace(sqla, " || 'U',", " || 'D',")
            sqla = Replace(sqla, "ORDR_QTY_SHIP", "ORDR_QTY_SHIP * ORDR_UNIT_PRICE")
            ASCDATA1.ExecuteSQL("Insert into " & SATCSLS2 & " " & sqla)

            sqla = "Select 'I' SI, DATA_TYPE, '" & CStr(YEAR) & "' YEAR" _
            & ", CUST_CODE, ITEM_CODE CODE_VALUE" _
            & sqlSATCSLS2 & " from " & SATCSLS2 _
            & " group by DATA_TYPE, CUST_CODE, ITEM_CODE"
            If YEAR = 0 Then
                ASCDATA1.ExecuteSQL("Create Table " & SATCSLS1 & " as " & sqla)
                ASCDATA1.ExecuteSQL("Alter Table " & SATCSLS1 & " Add Primary Key (SI, DATA_TYPE, YEAR, CUST_CODE, CODE_VALUE)")
            Else
                ASCDATA1.ExecuteSQL("Insert into " & SATCSLS1 & " " & sqla)
            End If

            sqla = "Select 'S' SI, DATA_TYPE, '" & CStr(YEAR) & "' YEAR" _
            & ", CUST_CODE, CUST_STORE_NO CODE_VALUE" & sqlSATCSLS2 _
            & " from " & SATCSLS2 & " group by DATA_TYPE, CUST_CODE, CUST_STORE_NO"
            ASCDATA1.ExecuteSQL("Insert into " & SATCSLS1 & " " & sqla)
        Next YEAR

        ASCDATA1.ExecuteSQL("Alter Table " & SATCSLS1 & " Add PXX NUMBER (10,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SATCSLS2 & " Add PXX NUMBER (10,0)")

        ASCMAIN1.sql = "Update " & SATCSLS1 & " X SET PXX = (Select P" & Format(Periods, "00") & " from " & SATCSLS1 & " " _
        & " where SI = X.SI and DATA_TYPE = 'H' || SUBSTR(X.DATA_TYPE,2,1)" _
        & "   and YEAR = X.YEAR and CUST_CODE = X.CUST_CODE and CODE_VALUE = X.CODE_VALUE)" _
        & " where DATA_TYPE IN ('TU','TD')"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update " & SATCSLS2 & " X SET PXX = (Select P" & Format(Periods, "00") & " from " & SATCSLS2 & " " _
        & " where DATA_TYPE = 'H' || SUBSTR(X.DATA_TYPE,2,1)" _
        & "   and YEAR = X.YEAR and CUST_CODE = X.CUST_CODE and CUST_STORE_NO = X.CUST_STORE_NO AND ITEM_CODE = X.ITEM_CODE)" _
        & " where DATA_TYPE IN ('TU','TD')"
        ASCDATA1.ExecuteSQL()

        ASCDATA1.ExecuteSQL("Alter Table " & SATCSLS1 & " Add PYY NUMBER (10,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SATCSLS2 & " Add PYY NUMBER (10,0)")

        Dim P00 As String = ""
        For P = 1 To Periods
            P00 &= "+NVL(P" & Format(P, "00") & ",0)"
        Next

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is" & vbCrLf _
            & " Select SI, DATA_TYPE, YEAR from " & SATCSLS1 & " where DATA_TYPE in ('TD','TU');" & vbCrLf _
            & "Begin For R1 in C1 Loop" & vbCrLf _
            & "Insert into " & SATCSLS1 & " (SI, DATA_TYPE, YEAR, CUST_CODE, CODE_VALUE)" & vbCrLf _
            & "Select SI, R1.DATA_TYPE, YEAR, CUST_CODE, CODE_VALUE from (" & vbCrLf _
            & "Select SI, YEAR, CUST_CODE, CODE_VALUE FROM " & SATCSLS1 & vbCrLf _
            & " where SI = R1.SI and DATA_TYPE = 'S' || SUBSTR(R1.DATA_TYPE,2) and YEAR = R1.YEAR" & vbCrLf _
            & "minus" & vbCrLf _
            & "Select SI, YEAR, CUST_CODE, CODE_VALUE FROM " & SATCSLS1 & vbCrLf _
            & " where SI = R1.SI and DATA_TYPE = R1.DATA_TYPE and YEAR = R1.YEAR" & vbCrLf _
            & ");" & vbCrLf _
            & "End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is" & vbCrLf _
            & " Select Distinct DATA_TYPE, YEAR from " & SATCSLS2 & " where DATA_TYPE in ('TD','TU');" & vbCrLf _
            & "Begin For R1 in C1 Loop" & vbCrLf _
            & "Insert into " & SATCSLS2 & " (DATA_TYPE, YEAR, CUST_CODE, CUST_STORE_NO, ITEM_CODE)" & vbCrLf _
            & "Select R1.DATA_TYPE, YEAR, CUST_CODE, CUST_STORE_NO, ITEM_CODE from (" & vbCrLf _
            & "Select YEAR, CUST_CODE, CUST_STORE_NO, ITEM_CODE FROM " & SATCSLS2 & vbCrLf _
            & " where DATA_TYPE = 'S' || SUBSTR(R1.DATA_TYPE,2) and YEAR = R1.YEAR" & vbCrLf _
            & "minus" & vbCrLf _
            & "Select YEAR, CUST_CODE, CUST_STORE_NO, ITEM_CODE FROM " & SATCSLS2 & vbCrLf _
            & " where DATA_TYPE = R1.DATA_TYPE and YEAR = R1.YEAR" & vbCrLf _
            & ");" & vbCrLf _
            & "End Loop; End; End;"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "Update " & SATCSLS1 & " X SET PYY = (Select " & Mid(P00, 2) & " from " & SATCSLS1 & " " _
        & " where SI = X.SI and DATA_TYPE = 'S' || SUBSTR(X.DATA_TYPE,2,1)" _
        & "   and YEAR = X.YEAR and CUST_CODE = X.CUST_CODE and CODE_VALUE = X.CODE_VALUE)" _
        & " where DATA_TYPE IN ('TU','TD')"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update " & SATCSLS2 & " X SET PYY = (Select " & Mid(P00, 2) & " from " & SATCSLS2 & " " _
        & " where DATA_TYPE = 'S' || SUBSTR(X.DATA_TYPE,2,1)" _
        & "   and YEAR = X.YEAR and CUST_CODE = X.CUST_CODE and CUST_STORE_NO = X.CUST_STORE_NO AND ITEM_CODE = X.ITEM_CODE)" _
        & " where DATA_TYPE IN ('TU','TD')"
        ASCDATA1.ExecuteSQL()


        STORES = New List(Of String)
        STORES_XX = ""
        Dim SQLX As String = ""

        If EntryMode = "E" Then
            SQLX = "Select Distinct CUST_STORE_NO from " & SATCSLS2 & " order by CUST_STORE_NO"
            For Each row As DataRow In ASCDATA1.GetDataTable(SQLX).Rows
                Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
                STORES.Add(CUST_STORE_NO)
            Next

        End If

        SQLX = ""
        If EntryMode = "E" Then
            For S As Integer = 1 To STORES.Count
                SQLX &= ", Sum (Decode(CUST_STORE_NO,'" & STORES(S - 1) & "'," & Mid(PERIODS_XX, 2) & ",0)) P" & Format(S, "00") & vbCrLf
                STORES_XX &= "+P" & Format(S, "00")
            Next
        End If
        sqlSATCSLS1_STORES = SQLX

    End Sub

    Sub Print_Report()
        Call Print_Report_Begin()

        Dim SUBT As String = ""
        Dim RecordSelectionFormula As String = ""
        Generate_Report("SARCSLS1", "", SUBT, RecordSelectionFormula)

        Call Print_Report_End()
    End Sub

    Private Sub optRANGE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRANGE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        grpPERIOD_RANGE.Visible = (optRANGE.Value = "P")
        grpWEEK_RANGE.Visible = (optRANGE.Value = "W")
    End Sub

    Private Sub optSI_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optSI.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        optSI.Tag = "*"
        Load_Data()
    End Sub

    Sub Setup_grd()

        Dim CAPTION As String = optType1.Text & " (" & optType2.Text & ") for " & CUST_CODE
        If optSI.Value = "S" Then
            CAPTION &= ", by Store"
        Else
            CAPTION &= ", by Item"
        End If
        grdSATCSLS1.Text = CAPTION

        Dim g1 As UltraWinGrid.UltraGrid
        Dim g2 As UltraWinGrid.UltraGrid
        If optSI.Value = "S" Then
            g1 = grdSATCSLS1
            g2 = grdSATCSLS2
        Else
            g1 = grdSATCSLS2
            g2 = grdSATCSLS1
        End If

        If optSI.Tag = "*" Then

            optSI.Tag = ""

            With g1.DisplayLayout.Bands(0)
                .Columns("CODE_VALUE").Header.Caption = "Store"
                .Columns("DESC_VALUE").Header.Caption = "Store Name"
                .Columns("SUB_CODE_VALUE1").Header.Caption = "Rep"
                .Columns("SUB_CODE_VALUE2").Header.Caption = "State"
                .Columns("SUB_CODE_VALUE3").Header.Caption = "City"
                .Columns("SUB_CODE_VALUE4").Header.Caption = "Zip"
                .Columns("SUB_CODE_VALUE5").Header.Caption = "Group"
                .Columns("RTL_PRICE").Header.Caption = "Retail"
                .Columns("WSL_PRICE").Header.Caption = "WhSale"

                .Columns("CODE_VALUE").Width = 60
                .Columns("DESC_VALUE").Width = 140
                .Columns("SUB_CODE_VALUE1").Width = 40
                .Columns("SUB_CODE_VALUE2").Width = 50
                .Columns("SUB_CODE_VALUE3").Width = 50
                .Columns("SUB_CODE_VALUE4").Width = 50
                .Columns("SUB_CODE_VALUE5").Width = 50
                .Columns("RTL_PRICE").Width = 65
                .Columns("WSL_PRICE").Width = 65
            End With

            With g2.DisplayLayout.Bands(0)
                .Columns("CODE_VALUE").Header.Caption = "Item"
                .Columns("DESC_VALUE").Header.Caption = "Description"
                .Columns("SUB_CODE_VALUE1").Header.Caption = "Collection"
                .Columns("SUB_CODE_VALUE2").Header.Caption = "Catgy"
                .Columns("SUB_CODE_VALUE3").Header.Caption = "Class"
                .Columns("SUB_CODE_VALUE4").Header.Caption = "Style"
                .Columns("SUB_CODE_VALUE5").Header.Caption = "Dept"
                .Columns("RTL_PRICE").Header.Caption = "Retail"
                .Columns("WSL_PRICE").Header.Caption = "WhSale"

                .Columns("CODE_VALUE").Width = 120
                .Columns("DESC_VALUE").Width = 180
                .Columns("SUB_CODE_VALUE1").Width = 80
                .Columns("SUB_CODE_VALUE2").Width = 60
                .Columns("SUB_CODE_VALUE3").Width = 60
                .Columns("SUB_CODE_VALUE4").Width = 60
                .Columns("SUB_CODE_VALUE5").Width = 60
                .Columns("RTL_PRICE").Width = 65
                .Columns("WSL_PRICE").Width = 65
            End With

            For Each G As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
            {grdSATCSLS1, grdSATCSLS2}
                With G.DisplayLayout.Bands(0)
                    If G.Name = "grdSATSLSC2" Then
                        .Columns("CODE_VALUE_PARENT").Hidden = True
                    End If
                    .Columns("CODE_VALUE").Hidden = False
                    .Columns("DESC_VALUE").Hidden = False
                    .Columns("SUB_CODE_VALUE1").Hidden = False
                    .Columns("SUB_CODE_VALUE2").Hidden = False
                    .Columns("SUB_CODE_VALUE3").Hidden = True
                    .Columns("SUB_CODE_VALUE4").Hidden = True
                    .Columns("SUB_CODE_VALUE5").Hidden = True
                    .Columns("RTL_PRICE").Hidden = True
                    .Columns("WSL_PRICE").Hidden = True
                End With
            Next

        End If

        If g1.DisplayLayout.Bands.Contains("SATCSLS1_SATCSLS1_DTL") Then
            g1.DisplayLayout.Bands("SATCSLS1_SATCSLS1_DTL").Hidden = Not chkExtendedData.Checked
        End If

        For Each G As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {g1, g2}
            Dim BMAX As Int32 = 0
            If G Is g1 And chkExtendedData.Checked Then
                BMAX = 1
            End If
            For B As Int32 = 0 To BMAX
                With G.DisplayLayout.Bands(B)
                    If B = 1 Then
                        .Columns("DATA_TYPE").Hidden = False
                        .Columns("DATA_TYPE").ColSpan = 3
                        .Columns("DATA_TYPE").Header.Caption = "Data Type"
                        .Columns("YEAR").Hidden = False
                        .Columns("YEAR").Header.Caption = "Year"
                        .Columns("YEAR").Width = 100
                        .RowLayoutStyle = UltraWinGrid.RowLayoutStyle.None
                        .Override.AllowColSizing = UltraWinGrid.AllowColSizing.Synchronized
                    End If
                    For P As Integer = 0 To Stores_Max
                        COLUMN_NAME = "P" & Format(P, "00")
                        If optSI.Value = "I" And optXP.Value = "S" And G.Name = grdSATCSLS1.Name Then
                            .Columns(COLUMN_NAME).Hidden = (P > STORES.Count)
                            If P <= STORES.Count Then
                                Dim LEGEND As String
                                If P = 0 Then
                                    LEGEND = "Total"
                                    .Columns(COLUMN_NAME).Width = 80
                                Else
                                    .Columns(COLUMN_NAME).Width = 70
                                    LEGEND = STORES(P - 1)
                                End If
                                .Columns(COLUMN_NAME).Header.Caption = LEGEND
                            End If
                        Else
                            .Columns(COLUMN_NAME).Hidden = (P > Periods)
                            If P <= Periods Then
                                Dim LEGEND As String
                                If P = 0 Then
                                    LEGEND = "Total"
                                    .Columns(COLUMN_NAME).Width = 80
                                Else
                                    .Columns(COLUMN_NAME).Width = 70
                                    If optRANGE.Value = "P" Then
                                        LEGEND = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(RYP0, P - 1))
                                        LEGEND = Mid(LEGEND, 10, 6)
                                    Else
                                        LEGEND = ASCMAIN1.Get_Legend_Wk(ASCMAIN1.Week_Calc(RYW0, P - 1))
                                        LEGEND = Mid(LEGEND, 10, 7)
                                    End If
                                End If
                                .Columns(COLUMN_NAME).Header.Caption = LEGEND
                            End If
                        End If
                    Next

                    .Columns("PXX").Hidden = Not (optType1.Value = "T")
                    .Columns("PXX").Header.Caption = "O/H"
                    .Columns("PYY").Hidden = Not (optType1.Value = "T")
                    .Columns("PYY").Header.Caption = "Sell-In"
                    .Columns("PYY").CellAppearance.BackColor = Drawing.Color.LightGreen
                    .Columns("PYY00PCT").Hidden = Not (optType1.Value = "T")
                    .Columns("PYY00PCT").Header.Caption = "SiSt%"
                    .Columns("PYY00PCT").Format = "#,##0.0"
                    .Columns("PYY00PCT").Width = 60
                    .Columns("PYY00PCT").CellAppearance.BackColor = Drawing.Color.LightGreen
                End With
            Next
        Next

    End Sub

    Private Sub grdSATCSLS1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSATCSLS1.AfterRowActivate
        If EntryMode = "E" Then
            Setup_grdSATCSLS2()
        End If
    End Sub

    Sub Setup_grdSATCSLS2()

        If grdSATCSLS1.ActiveRow Is Nothing OrElse Not grdSATCSLS1.ActiveRow.IsDataRow Then
            chkShowDetails.Checked = False
            chkShowDetails.Enabled = False
            Exit Sub
        Else
            chkShowDetails.Enabled = True
        End If

        Dim DATA_TYPE As String = optType1.Value & optType2.Value
        Dim CODE_VALUE_PARENT As String = grdSATCSLS1.ActiveRow.Cells("CODE_VALUE").Text

        Load_SATCSLS2(DATA_TYPE, CODE_VALUE_PARENT, False)
        Sort_grdColumns(grdSATCSLS2, "CODE_VALUE")

        Dim CAPTION As String = optType1.Text & " (" & optType2.Text & ") for " & CUST_CODE
        If optSI.Value = "S" Then
            CAPTION &= " - Store " & CODE_VALUE_PARENT & ", by Item"
        Else
            CAPTION &= " - Item " & CODE_VALUE_PARENT & ", by Store"
        End If
        grdSATCSLS2.Text = CAPTION

        Dim sql As String = ""
        sql = sqlSOTINVHX & " and SOTINVH2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & " and SOTINVH2.ORDR_YYYYPP_UPDATED between '" & RYP0 & "' and '" & RYP1 & "'" & vbCrLf
        If optSI.Value = "I" Then
            sql &= " and SOTINVH2.ITEM_CODE = '" & CODE_VALUE_PARENT & "'" & vbCrLf
        Else
            sql &= " and SOTINVH2.CUST_STORE_NO = '" & CODE_VALUE_PARENT & "'" & vbCrLf
        End If
        With grdSOTINVHX.DisplayLayout.Bands(0)
            .Columns("CUST_STORE_NO").Hidden = (optSI.Value = "S")
            .Columns("CUST_STORE_LOCATION").Hidden = (optSI.Value = "S")
            .Columns("ITEM_CODE").Hidden = (optSI.Value = "I")
            .Columns("ITEM_DESC").Hidden = (optSI.Value = "I")
        End With
        Fill_Records("SOTINVHX", "", True, sql)
        grdSOTINVHX.Text = "Sales Documents for " & CUST_CODE & IIf(optSI.Value = "S", " - Store ", " - Item ") & CODE_VALUE_PARENT
        grdSOTINVHX.DisplayLayout.CaptionVisible = DefaultableBoolean.True

    End Sub

    Sub Load_SATCSLS2( _
                     ByVal DATA_TYPE As String, _
                     ByVal CODE_VALUE_PARENT As String, _
                     ByVal all_parents As Boolean, _
                     Optional clear_before_filling As Boolean = True)
        Dim sql As String = ""

        ' & ", NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_NAME) || DECODE(CUST_STORE_MARK_FOR,NULL,'',' (' || CUST_STORE_MARK_FOR || ')') DESC_VALUE" & vbCrLf _

        If optSI.Value = "I" Then
            sql = "Select SATCSLS2.ITEM_CODE CODE_VALUE_PARENT " & vbCrLf _
            & ", SATCSLS2.CUST_STORE_NO CODE_VALUE" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_NAME DESC_VALUE" & vbCrLf _
            & ", ARTCUST2.SELL_CODE SUB_CODE_VALUE1" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_STATE SUB_CODE_VALUE2" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_CITY SUB_CODE_VALUE3" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_ZIP_CODE SUB_CODE_VALUE4" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_GROUP SUB_CODE_VALUE5" & vbCrLf _
            & sqlSATCSLS1 & ",PXX,PYY from ARTCUST2," & SATCSLS2 & " SATCSLS2 " & vbCrLf _
            & " where ARTCUST2.CUST_CODE (+) = SATCSLS2.CUST_CODE " & vbCrLf _
            & " and ARTCUST2.CUST_STORE_NO (+) = SATCSLS2.CUST_STORE_NO" & vbCrLf _
            & " and DATA_TYPE = '" & DATA_TYPE & "'" & vbCrLf _
            & IIf(all_parents, "", " and SATCSLS2.ITEM_CODE = '" & CODE_VALUE_PARENT & "'")
        Else
            sql = "Select SATCSLS2.CUST_STORE_NO CODE_VALUE_PARENT " & vbCrLf _
            & ", SATCSLS2.ITEM_CODE CODE_VALUE" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC DESC_VALUE" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE SUB_CODE_VALUE1" & vbCrLf _
            & ", ICTITEM1.ITEM_CATGY_CODE SUB_CODE_VALUE2" & vbCrLf _
            & ", ICTITEM1.ITEM_CLASS_CODE SUB_CODE_VALUE3" & vbCrLf _
            & ", ICTITEM1.STYLE_CODE SUB_CODE_VALUE4" & vbCrLf _
            & ", ICTITEM1.DEPT_CODE SUB_CODE_VALUE5" & vbCrLf _
            & ", ICTITEM1.ITEM_RETAIL_PRICE RTL_PRICE" & vbCrLf _
            & ", ICTITEM1.ITEM_PRICE WSL_PRICE" & vbCrLf _
            & sqlSATCSLS1 & ",PXX,PYY  from ICTITEM1," & SATCSLS2 & " SATCSLS2 " & vbCrLf _
            & " where ICTITEM1.ITEM_CODE (+) = SATCSLS2.ITEM_CODE " & vbCrLf _
            & " and DATA_TYPE = '" & DATA_TYPE & "'" & vbCrLf _
            & IIf(all_parents, "", " and SATCSLS2.CUST_STORE_NO = '" & CODE_VALUE_PARENT & "'")
        End If
        Fill_Records("SATCSLS2", "", clear_before_filling, sql)

        Dim dvw As DataView = dst.Tables("SATCSLS2").DefaultView
        dvw.RowFilter = "CODE_VALUE_PARENT = '" & CODE_VALUE_PARENT & "'"

    End Sub

    Private Sub chkNoDetails_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowDetails.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        SplitContainer1.Panel2Collapsed = Not chkShowDetails.Checked
    End Sub

    Sub Load_Data()

        dst.EnforceConstraints = False

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Dim DATA_TYPE As String = optType1.Value & optType2.Value

        optXP.Visible = Not (optSI.Value = "S")

        Dim sql As String = ""

        If optSI.Value = "S" Then
            If EntryMode = "E" Then
                ' & ", NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_NAME) || DECODE(CUST_STORE_MARK_FOR,NULL,'',' (' || CUST_STORE_MARK_FOR || ')') DESC_VALUE" & vbCrLf _

                sql = "Select SATCSLS1.CODE_VALUE CODE_VALUE" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_NAME DESC_VALUE" & vbCrLf _
                & ", ARTCUST2.SELL_CODE SUB_CODE_VALUE1" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_STATE SUB_CODE_VALUE2" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_CITY SUB_CODE_VALUE3" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_ZIP_CODE SUB_CODE_VALUE4" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_GROUP SUB_CODE_VALUE5" & vbCrLf _
                & sqlSATCSLS1 & ",PXX,PYY from ARTCUST2," & SATCSLS1 & " SATCSLS1 " & vbCrLf _
                & " where ARTCUST2.CUST_CODE (+) = SATCSLS1.CUST_CODE " & vbCrLf _
                & " and ARTCUST2.CUST_STORE_NO (+) = SATCSLS1.CODE_VALUE" & vbCrLf _
                & " and SATCSLS1.SI = '" & optSI.Value & "' and SATCSLS1.DATA_TYPE = '" & DATA_TYPE & "'" & vbCrLf _
                & " and SATCSLS1.YEAR = '0'"
            Else
                sql = "Select SATCSLS1.CUST_CODE CODE_VALUE" & vbCrLf _
                & ", ARTCUST1.CUST_NAME DESC_VALUE" & vbCrLf _
                & ", ARTCUST1.SREP_CODE SUB_CODE_VALUE1" & vbCrLf _
                & ", ARTCUST1.CUST_STATE SUB_CODE_VALUE2" & vbCrLf _
                & ", ARTCUST1.CUST_CITY SUB_CODE_VALUE3" & vbCrLf _
                & ", ARTCUST1.CUST_ZIP_CODE SUB_CODE_VALUE4" & vbCrLf _
                & ", ARTCUST1.TRADE_CLASS_CODE SUB_CODE_VALUE5" & vbCrLf _
                & sqlSATCSLS1_sum & ", Sum (PXX) PXX, Sum (PYY) PYY from ARTCUST1," & SATCSLS1 & " SATCSLS1 " & vbCrLf _
                & " where ARTCUST1.CUST_CODE (+) = SATCSLS1.CUST_CODE " & vbCrLf _
                & " and SATCSLS1.SI = '" & optSI.Value & "' and SATCSLS1.DATA_TYPE = '" & DATA_TYPE & "'" & vbCrLf _
                & " and SATCSLS1.YEAR = '0'" & vbCrLf _
                & " group by SATCSLS1.CUST_CODE" & vbCrLf _
                & ", ARTCUST1.CUST_NAME" & vbCrLf _
                & ", ARTCUST1.SREP_CODE" & vbCrLf _
                & ", ARTCUST1.CUST_STATE" & vbCrLf _
                & ", ARTCUST1.CUST_CITY" & vbCrLf _
                & ", ARTCUST1.CUST_ZIP_CODE" & vbCrLf _
                & ", ARTCUST1.TRADE_CLASS_CODE"
            End If
        Else
            If optXP.Value = "S" Then

                Dim SQLX As String = sqlSATCSLS1_STORES
                If Mid(DATA_TYPE, 1, 1) = "H" Then
                    SQLX = Replace(sqlSATCSLS1_STORES, Mid(PERIODS_XX, 2), "P" & Format(Periods, "00"))
                End If

                sql = "Select SATCSLS2.ITEM_CODE CODE_VALUE" & vbCrLf _
                & ", ICTITEM1.ITEM_DESC DESC_VALUE" & vbCrLf _
                & ", ICTITEM1.COLLECTION_CODE SUB_CODE_VALUE1" & vbCrLf _
                & ", ICTITEM1.ITEM_CATGY_CODE SUB_CODE_VALUE2" & vbCrLf _
                & ", ICTITEM1.ITEM_CLASS_CODE SUB_CODE_VALUE3" & vbCrLf _
                & ", ICTITEM1.STYLE_CODE SUB_CODE_VALUE4" & vbCrLf _
                & ", ICTITEM1.DEPT_CODE SUB_CODE_VALUE5" & vbCrLf _
                & ", ICTITEM1.ITEM_RETAIL_PRICE RTL_PRICE" & vbCrLf _
                & ", ICTITEM1.ITEM_PRICE WSL_PRICE" & vbCrLf _
                & SQLX & "  from ICTITEM1," & SATCSLS2 & " SATCSLS2 " & vbCrLf _
                & " where ICTITEM1.ITEM_CODE (+) = SATCSLS2.ITEM_CODE " & vbCrLf _
                & " and SATCSLS2.DATA_TYPE = '" & DATA_TYPE & "'" & vbCrLf _
                & " and SATCSLS2.YEAR = '0'" & vbCrLf _
                & " group by SATCSLS2.ITEM_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_DESC" & vbCrLf _
                & ", ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_CATGY_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_CLASS_CODE" & vbCrLf _
                & ", ICTITEM1.STYLE_CODE" & vbCrLf _
                & ", ICTITEM1.DEPT_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_RETAIL_PRICE" & vbCrLf _
                & ", ICTITEM1.ITEM_PRICE" & vbCrLf
            Else
                sql = "Select SATCSLS1.CODE_VALUE CODE_VALUE" & vbCrLf _
                & ", ICTITEM1.ITEM_DESC DESC_VALUE" & vbCrLf _
                & ", ICTITEM1.COLLECTION_CODE SUB_CODE_VALUE1" & vbCrLf _
                & ", ICTITEM1.ITEM_CATGY_CODE SUB_CODE_VALUE2" & vbCrLf _
                & ", ICTITEM1.ITEM_CLASS_CODE SUB_CODE_VALUE3" & vbCrLf _
                & ", ICTITEM1.STYLE_CODE SUB_CODE_VALUE4" & vbCrLf _
                & ", ICTITEM1.DEPT_CODE SUB_CODE_VALUE5" & vbCrLf _
                & ", ICTITEM1.ITEM_RETAIL_PRICE RTL_PRICE" & vbCrLf _
                & ", ICTITEM1.ITEM_PRICE WSL_PRICE" & vbCrLf _
                & sqlSATCSLS1 & ",PXX,PYY from ICTITEM1," & SATCSLS1 & " SATCSLS1 " & vbCrLf _
                & " where ICTITEM1.ITEM_CODE (+) = SATCSLS1.CODE_VALUE " & vbCrLf _
                & " and SATCSLS1.SI = '" & optSI.Value & "'" _
                & " and SATCSLS1.DATA_TYPE = '" & DATA_TYPE & "'" _
                & " and SATCSLS1.YEAR = '0'" & vbCrLf
            End If
        End If

        dst.Tables("SATCSLS1").Rows.Clear()
        dst.Tables("SATCSLS1_DTL").Rows.Clear()
        dst.Tables("SATCSLS2").Rows.Clear()

        If optSI.Value = "I" And optXP.Value = "S" Then
            dst.Tables("SATCSLS1").Columns("P00").Expression = Mid(STORES_XX, 2)
            dst.Tables("SATCSLS1_DTL").Columns("P00").Expression = Mid(STORES_XX, 2)
        Else
            If optType1.Value = "H" Then
                ' NOTE THAT SAFCSLS1 HAS A 2 DIGIT PERIOD AND SAFISLS1 HAS A 3 DIGIT PERIOD
                '  dst.Tables("SATCSLS1").Columns("P00").Expression = "P" & Mid(PERIODS_XX, Len(PERIODS_XX) - 2, 3)
                dst.Tables("SATCSLS1").Columns("P00").Expression = Mid(PERIODS_XX, Len(PERIODS_XX) - 2, 3)
            Else
                dst.Tables("SATCSLS1").Columns("P00").Expression = Mid(PERIODS_XX, 2)
            End If
            dst.Tables("SATCSLS1_DTL").Columns("P00").Expression = "IIF(DATA_TYPE='HU' OR DATA_TYPE='HD'," & Mid(PERIODS_XX, Len(PERIODS_XX) - 2, 3) & "," & Mid(PERIODS_XX, 2) & ")"
        End If

        If optType1.Value = "H" Then
            ' NOTE THAT SAFCSLS1 HAS A 2 DIGIT PERIOD AND SAFISLS1 HAS A 3 DIGIT PERIOD
            ' dst.Tables("SATCSLS2").Columns("P00").Expression = "P" & Mid(PERIODS_XX, Len(PERIODS_XX) - 2, 3)
            dst.Tables("SATCSLS2").Columns("P00").Expression = Mid(PERIODS_XX, Len(PERIODS_XX) - 2, 3)
        Else
            dst.Tables("SATCSLS2").Columns("P00").Expression = Mid(PERIODS_XX, 2)
        End If

        Fill_Records("SATCSLS1", "", True, sql)
        Sort_grdColumns(grdSATCSLS1, "CODE_VALUE")

        If chkExtendedData.Checked Then
            sql = "Select SATCSLS1.CODE_VALUE CODE_VALUE, SATCSLS1.YEAR, SATCSLS1.DATA_TYPE" _
            & sqlSATCSLS1 _
            & " from " & SATCSLS1 & " SATCSLS1 " _
            & " where SATCSLS1.SI = '" & optSI.Value & "'" _
            & " and SATCSLS1.DATA_TYPE LIKE '%" & Mid(DATA_TYPE, 2, 1) & "'" _
            & " and (YEAR = '1' OR SATCSLS1.DATA_TYPE <> '" & DATA_TYPE & "')"

            Fill_Records("SATCSLS1_DTL", "", True, sql)
            Sort_grdColumns(grdSATCSLS1, "DATA_TYPE", , 1)
        End If

        If grdSATCSLS1.Rows.Count = 0 Then
            chkShowDetails.Checked = False
            chkShowDetails.Enabled = False
        Else
            chkShowDetails.Enabled = True
        End If

        'dst.EnforceConstraints = True

        Setup_grd()


        tabDetails.Tabs("EDI Documents").Visible = (optType1.Value = "T")
        tabDetails.Tabs("Sales Documents").Visible = (optType1.Value <> "T")

        If optSI.Value = "I" Then
            If tabDetails.SelectedTab.Key = "Map" Then
                tabDetails.SelectedTab = tabDetails.Tabs("Details")
            End If
        Else
            'dst.Tables("SATCSLSS").Rows.Clear()
            For Each ROW As DataRow In dst.Tables("SATCSLSS").Rows
                ROW.Item("SALES") = 0
            Next
            For Each row As DataRow In ASCDATA1.SelectDistinct("SATCSLS1", "SUB_CODE_VALUE2").Rows
                Dim rowTATSTATE As DataRow = dst.Tables("TATSTATE").Rows.Find(row(0))
                If rowTATSTATE IsNot Nothing Then
                    Dim rows() As DataRow = dst.Tables("SATCSLSS").Select _
                    ("STATE_NAME = '" & rowTATSTATE("STATE_NAME") & "'")
                    If rows.Length = 1 Then
                        Dim SALES As Decimal = Val(dst.Tables("SATCSLS1").Compute("SUM (P00)", "SUB_CODE_VALUE2 = '" & row(0) & "'") & "")
                        rows(0).Item("STATE_CODE") = rowTATSTATE("STATE_CODE")
                        rows(0).Item("SALES") = SALES
                    End If
                End If

                'dst.Tables("SATCSLSS").Rows.Add(New Object() _
                '    {row("STATE_CODE"), row("STATE_NAME"), Val(row("SALES") & "")})
            Next
            Me.UltraChart1.Data.DataSource = StatesData()
            Me.UltraChart1.Data.DataBind()
        End If
        tabDetails.Tabs("Map").Visible = (optSI.Value = "S")

        CreateGraph_SATCSLS1()
        CreateGraph_SATCSLS1_X()
        chtSATCSLS1.Visible = True
        chtSATCSLS1_X.Visible = True

        If optSI.Value = "S" Then
            If grdSATCSLS1.DisplayLayout.ValueLists.Exists("SUB_CODE_VALUE2") Then
                grdSATCSLS1.DisplayLayout.ValueLists.Remove("SUB_CODE_VALUE2")
            End If
            If grdSATCSLS2.DisplayLayout.ValueLists.Exists("SUB_CODE_VALUE1") Then
                grdSATCSLS2.DisplayLayout.ValueLists.Remove("SUB_CODE_VALUE1")
            End If

            ASCMAIN1.Add_Value_List(grdSATCSLS1, "SUB_CODE_VALUE1", , , , "Select SELL_CODE, SELL_NAME from SOTSELL1")
            ASCMAIN1.Add_Value_List(grdSATCSLS2, "SUB_CODE_VALUE2", , , , "Select ITEM_CATGY_CODE, ITEM_CATGY_DESC from ICTCATG1")
        Else
            If grdSATCSLS1.DisplayLayout.ValueLists.Exists("SUB_CODE_VALUE1") Then
                grdSATCSLS1.DisplayLayout.ValueLists.Remove("SUB_CODE_VALUE1")
            End If
            If grdSATCSLS2.DisplayLayout.ValueLists.Exists("SUB_CODE_VALUE2") Then
                grdSATCSLS2.DisplayLayout.ValueLists.Remove("SUB_CODE_VALUE2")
            End If

            ASCMAIN1.Add_Value_List(grdSATCSLS1, "SUB_CODE_VALUE2", , , , "Select ITEM_CATGY_CODE, ITEM_CATGY_DESC from ICTCATG1")
            ASCMAIN1.Add_Value_List(grdSATCSLS2, "SUB_CODE_VALUE1", , , , "Select SELL_CODE, SELL_NAME from SOTSELL1")
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub optType1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optType1.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_Data()
    End Sub

    Private Sub optType2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optType2.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If PERIODS_XX = "" Then Exit Sub
        Load_Data()
    End Sub

    Sub CreateGraph_SATCSLS1()

        Dim chtIsVisible As Boolean = chtSATCSLS1.Visible
        chtSATCSLS1.Visible = False

        chtSATCSLS1.DataSource = Nothing

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Charting Data")

        Me.SuspendLayout()

        Dim RL() As String

        chtSATCSLS1.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtSATCSLS1.LabelHash = labelHash

        chtSATCSLS1.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtSATCSLS1.Tooltips.FormatString = "<HIGHLOW>"

        Dim DTY As New DataTable
        With DTY
            .Columns.Add("CODE")
            .Columns.Add("VALUE", GetType(System.Decimal))
        End With

        Dim RLi As Integer = 0
        ReDim RL(dst.Tables("SATCSLS1").Rows.Count - 1)
        For Each row As DataRow In dst.Tables("SATCSLS1").Select("", "CODE_VALUE")
            RL(RLi) = row("CODE_VALUE") ' & ":" & row("DESC_VALUE")
            RLi += 1
            DTY.Rows.Add(New Object() {row.Item("CODE_VALUE"), row.Item("P00")})
        Next
        'chtSATCSLS1.Data.SetRowLabels(RL)
        'chtSATCSLS1.Data.SetColumnLabels(CL)

        'chtSATCSLS1.DataSource = dst.Tables("SATCSLS1")
        chtSATCSLS1.DataSource = DTY
        chtSATCSLS1.PieChart.ColumnIndex = -1
        chtSATCSLS1.PieChart.OthersCategoryPercent = 2
        'chtSATCSLS1.Data.IncludeColumn("CODE_VALUE", False)
        'chtSATCSLS1.Data.IncludeColumn("DESC_VALUE", False)
        'chtSATCSLS1.Data.IncludeColumn("P00", True)


        chtSATCSLS1.DataBind()

        chtSATCSLS1.Visible = chtIsVisible

        Me.ResumeLayout()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()
    End Sub

    Sub CreateGraph_SATCSLS1_X()

        Dim chtIsVisible As Boolean = chtSATCSLS1_X.Visible
        chtSATCSLS1_X.Visible = False

        chtSATCSLS1_X.DataSource = Nothing

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
            CL(i - 1) = grdSATCSLS1.DisplayLayout.Bands(0).Columns("P" & Format(i, "00")).Header.Caption
            'grdSATCSLS1.DisplayLayout.Bands(0).Columns("M" & Format(i, "00")).Header.Caption = Mid(L, 10, 3)
        Next

        'chtICTINVAT.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.LabelPlusDataValue
        'chtICTINVAT.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom

        chtSATCSLS1_X.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtSATCSLS1_X.LabelHash = labelHash

        chtSATCSLS1_X.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtSATCSLS1_X.Tooltips.FormatString = "<HIGHLOW>"

        Dim DT As New DataTable
        DT.Columns.Add("CODE_VALUE")
        DT.Columns.Add("DESC_VALUE")
        For P As Integer = 1 To Periods
            DT.Columns.Add("P" & Format(P, "00"), GetType(System.Decimal))
        Next

        Dim RLi As Integer = 0
        ReDim RL(dst.Tables("SATCSLS1").Rows.Count - 1)
        For Each row As DataRow In dst.Tables("SATCSLS1").Select("", "CODE_VALUE")
            RL(RLi) = row("CODE_VALUE") ' & ":" & row("DESC_VALUE")
            RLi += 1

            Dim rowDT As DataRow = DT.NewRow
            rowDT.Item("CODE_VALUE") = row("CODE_VALUE")
            rowDT.Item("DESC_VALUE") = row("DESC_VALUE")
            For P As Integer = 1 To Periods
                rowDT.Item("P" & Format(P, "00")) = row("P" & Format(P, "00"))
            Next
            DT.Rows.Add(rowDT)
        Next
        chtSATCSLS1_X.Data.SetRowLabels(RL)
        chtSATCSLS1_X.Data.SetColumnLabels(CL)

        chtSATCSLS1_X.DataSource = DT
        'chtSATCSLS1_X.Data.IncludeColumn("CODE_VALUE", False)
        'chtSATCSLS1_X.Data.IncludeColumn("DESC_VALUE", False)
        'chtSATCSLS1_X.Data.IncludeColumn("P00", False)

        chtSATCSLS1_X.DataBind()

        chtSATCSLS1_X.Visible = chtIsVisible

        Me.ResumeLayout()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()
    End Sub

    Private Sub trkbrXAxis_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles trkbrXAxis.Scroll
        chtSATCSLS1_X.Axis.X.ScrollScale.Scale = Me.trkbrXAxis.Value / 100.0
    End Sub

    Private Sub trkbrYAxis_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles trkbrYAxis.Scroll
        chtSATCSLS1_X.Axis.Y.ScrollScale.Scale = Me.trkbrYAxis.Value / 100.0
    End Sub

    Private Sub tabDetails_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabDetails.SelectedTabChanged
        Setup_tabDetails()
    End Sub

    Sub Setup_tabDetails()
        If SELECTION_NO = 0 Then Exit Sub
        UltraExplorerBar1.Groups("Charts").Visible = (tabDetails.SelectedTab.Key = "Charts")
    End Sub

    Private Sub optTotalsChartType_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTotalsChartType.ValueChanged
        Set_Totals_ChartType()
    End Sub

    Private Sub chkTotalsChart3D_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTotalsChart3D.CheckedChanged
        Set_Totals_ChartType()
    End Sub

    Sub Set_Totals_ChartType()
        If Not chkTotalsChart3D.Checked Then
            chtSATCSLS1_X.ChartType = ChartType.LineChart
            Select Case optTotalsChartType.Value
                Case "PieChart"
                    chtSATCSLS1.ChartType = ChartType.PieChart
                Case "DoughnutChart"
                    chtSATCSLS1.ChartType = ChartType.DoughnutChart
            End Select
        Else
            chtSATCSLS1_X.ChartType = ChartType.LineChart3D
            Select Case optTotalsChartType.Value
                Case "PieChart"
                    chtSATCSLS1.ChartType = ChartType.PieChart3D
                Case "DoughnutChart"
                    chtSATCSLS1.ChartType = ChartType.DoughnutChart3D
            End Select

        End If
    End Sub

    Private Sub cbeColor_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeColor.ValueChanged
        'chtSATCSLS1.ColorModel.ModelStyle = cbeColor.ValueMember
        'chtSATCSLS1_X.ColorModel.ModelStyle = Infragistics.UltraChart.Shared.Styles.ColorModels.PureRandom
        chtSATCSLS1.ColorModel.ModelStyle = CType(System.Enum.Parse(GetType(ColorModels), cbeColor.SelectedItem.ToString()), ColorModels)
        chtSATCSLS1_X.ColorModel.ModelStyle = CType(System.Enum.Parse(GetType(ColorModels), cbeColor.SelectedItem.ToString()), ColorModels)

    End Sub

    Sub Setup_Map()
        '' create the layer
        Dim points As String = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), ASCMAIN1.Folders("Images") & "ABS\UsMap\US_STATES.xml")
        'Dim USmap As New MapLayer(points)
        USmap = New MapLayer(points)

        dst.Tables("SATCSLSS").Rows.Clear()
        US_STATES = USmap.STATES
        For i As Integer = 0 To USmap.STATES.Length - 1
            dst.Tables("SATCSLSS").Rows.Add(New Object() {"", USmap.STATES(i), 0})
        Next

        '' set the layer
        Me.UltraChart1.ChartType = ChartType.Composite
        Me.UltraChart1.CompositeChart.ChartAreas.Add(New Infragistics.UltraChart.Resources.Appearance.ChartArea())
        Me.UltraChart1.UserLayerIndex = New String() {"USMap"}
        Me.UltraChart1.Layer.Add("USMap", USmap)

        '' set the tooltip.
        Dim labelRenderers As New Hashtable()
        labelRenderers.Add("USMap", New USMapLabelRenderer(dst.Tables("SATCSLSS")))
        Me.UltraChart1.LabelHash = labelRenderers
        Me.UltraChart1.Tooltips.FormatString = "<USMap>"

        ''set border
        Me.UltraChart1.Border.CornerRadius = 20
        Me.UltraChart1.Border.Thickness = 0
        Me.UltraChart1.BackColor = System.Drawing.Color.White

        '' set color model
        'Me.UltraChart1.ColorModel.ColorBegin = System.Drawing.Color.AliceBlue
        Me.UltraChart1.ColorModel.ColorBegin = System.Drawing.Color.Red
        Me.UltraChart1.ColorModel.ColorEnd = System.Drawing.Color.Blue '  Color.Yellow ' Color.FromArgb(24, 89, 165)
        Me.UltraChart1.ColorModel.AlphaLevel = 255
        Me.UltraChart1.ColorModel.ModelStyle = ColorModels.DataValueLinearRange

        '' legend
        Me.UltraChart1.Legend.Visible = True
        Me.UltraChart1.Axis.X.Extent = 10
        Me.UltraChart1.Legend.SpanPercentage = 10
        Me.UltraChart1.Legend.Location = LegendLocation.Right

        '' set the data
        Me.UltraChart1.Data.DataSource = StatesData()
        Me.UltraChart1.Data.DataBind()
    End Sub

#Region "Create StateDataView Data"

    Private Function StatesData() As StateDataInfo()
        Dim StatesDataFromDataSource() As StateDataInfo
        ReDim StatesDataFromDataSource(49)
        If SELECTION_NO <> 0 Then
            For I As Integer = 0 To US_STATES.Length - 1
                Dim rows() As DataRow = dst.Tables("SATCSLSS").Select("STATE_NAME = '" & US_STATES(I) & "'")
                Dim SALES As Int32 = 0
                If rows.Length = 1 Then
                    SALES = Val(rows(0).Item("SALES") & "")
                End If
                StatesDataFromDataSource(I) = New StateDataInfo(US_STATES(I), SALES, "")
            Next
        End If
        'StatesDataFromDataSource(0) = New StateExpenseViewInfo("Alabama", 1915560.96, "")
        Return StatesDataFromDataSource
    End Function
#End Region

    Private Sub grdSATCSLSS_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSATCSLSS.DoubleClickRow
        Show_Filter(grdSATCSLS1, True)
        grdSATCSLS1.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        grdSATCSLS1.Rows.ColumnFilters("SUB_CODE_VALUE2").FilterConditions.Add _
        (Infragistics.Win.UltraWinGrid.FilterComparisionOperator.Equals, e.Row.Cells("STATE_CODE").Text)
        chkShowDetails.Checked = True
    End Sub

    Private Sub grdSATCSLSS_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSATCSLSS.InitializeRow
        If USmap.COLORS.ContainsKey(e.Row.Cells("STATE_NAME").Text) Then
            e.Row.Cells("SALES").Appearance.ForeColor = USmap.COLORS(e.Row.Cells("STATE_NAME").Text)
        End If
    End Sub

    Private Sub optXP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optXP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If optXP.Value = "S" Then
            If STORES.Count > Stores_Max Then
                MsgBox("Too Many Stores (" & STORES.Count & ") for this option.  Max is " & CStr(Stores_Max))
                optXP.Value = "P"
                Exit Sub
            End If
        End If

        Load_Data()
    End Sub

    Sub Load_SIST()
        Load_SATSIST1()
        Setup_tabSIST()
        grdSATSIST1.Text = "Sell-In / Sell-Thru"
        'grdSATSIST3.Text = "Sell-In / Sell-Thru LY" & IIf(ITEM_CODE_LY = ITEM_CODE, "", " (" & ITEM_CODE_LY & ")")
    End Sub

    Sub Setup_SIST()
        If Periods <= SIST_PMAX Then
        Else
            MsgBox("Cannot view Sell-In / Sell-Thru with more than 53 " & optRANGE.Items(optRANGE.CheckedIndex).DisplayText & "s")
            chkSIST.Checked = False
            Exit Sub
        End If
    End Sub

    Private Sub chkExtendedData_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkExtendedData.CheckedChanged
        If Not chkExtendedData.Checked Then
            chkPriorYear.Checked = False
        End If
        chkPriorYear.Enabled = chkExtendedData.Checked
    End Sub

    Sub Validate_Range(ByRef EMsg As String)

        If Absx1.optFor("RANGE").Value = "P" Then
            If Absx1.cmbFor("RYP0").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify a Starting Period"
            End If
            If Absx1.cmbFor("RYP1").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify an Ending Period"
            End If

            If EMsg = "" Then
                RYP0 = Absx1.cmbFor("RYP0").Value
                RYP1 = Absx1.cmbFor("RYP1").Value
                Periods = ASCMAIN1.Period_Diff(RYP0, RYP1) + 1
            End If
        Else
            If Absx1.cmbFor("RYW0").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify a Starting Week"
            End If
            If Absx1.cmbFor("RYW1").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify an Ending Week"
            End If
            'If Absx1.dteFor("RYW0").Value > Absx1.dteFor("RYW1").Value Then
            '    EMsg &= vbCr & "Starting Week cannot be later than Ending Week"
            'End If

            If EMsg = "" Then
                RYW0 = Absx1.cmbFor("RYW0").Value
                RYW1 = Absx1.cmbFor("RYW1").Value
                Periods = ASCMAIN1.Week_Diff(RYW0, RYW1) + 1
            End If
        End If

        If EMsg = "" Then
            If Periods < 1 Or Periods > Stores_Max Then
                EMsg &= vbCr & "Total number of Periods must be between 1 and " & CStr(Stores_Max)
            End If
        End If

    End Sub

    Private Sub grdSATCSLS1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSATCSLS1.DoubleClickRow
        If EntryMode = "" Then
            Absx1.txtFor("CUST_CODE").Text = e.Row.Cells("CODE_VALUE").Value
            Click_Command("View")
        End If
    End Sub

    Private Sub cmdS0_Click(sender As System.Object, e As System.EventArgs) Handles cmdS0.Click
        Set_Period(cmdS0.Text)
    End Sub

    Private Sub cmdS1_Click(sender As System.Object, e As System.EventArgs) Handles cmdS1.Click
        Set_Period(cmdS1.Text)
    End Sub

    Private Sub cmdS2_Click(sender As System.Object, e As System.EventArgs) Handles cmdS2.Click
        Set_Period(cmdS2.Text)
    End Sub

    Sub Set_Period(SYY As String)
        Dim YY As String = Mid(SYY, 2, 2)

        Dim P0 As String = ""
        Dim P1 As String = ""

        If Mid(SYY, 1, 1) = "S" Then
            P0 = "20" & YY & "02"
            P1 = "20" & YY & "07"
        Else
            P0 = "20" & YY & "08"
            P1 = "20" & Format(Val(YY) + 1, "00") & "01"
        End If

        ASCMAIN1.sql = "Select OPS_YYYYPP from GLTPARM2 where OPS_YYYYMM = '" & P0 & "'"
        Absx1.cmbFor("RYP0").Value = ASCDATA1.GetDataValue
        ASCMAIN1.sql = "Select OPS_YYYYPP from GLTPARM2 where OPS_YYYYMM = '" & P1 & "'"
        Absx1.cmbFor("RYP1").Value = ASCDATA1.GetDataValue

        ASCMAIN1.sql = "Select MIN (YYYYWW) from GLTPARM3 where YYYYMM = '" & P0 & "'"
        Absx1.cmbFor("RYW0").Value = ASCDATA1.GetDataValue
        ASCMAIN1.sql = "Select MAX (YYYYWW) from GLTPARM3 where YYYYMM = '" & P1 & "'"
        Absx1.cmbFor("RYW1").Value = ASCDATA1.GetDataValue
    End Sub

    Sub Load_SATSIST1()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Sell-In & Sell-Thru")

        Dim sqlp As String = ""
        Dim P_SI As String = ""
        Dim P_ST As String = ""
        If optRANGE.Value = "P" Then
            sqlp = " and ORDR_YYYYPP_UPDATED between '" & RYP0 & "' and '" & RYP1 & "'"
            P_SI = "SOTINVH2.ORDR_YYYYPP_UPDATED"
            P_ST = "OPS_YYYYPP"
        Else
            sqlp = " and OPS_YYYYWW between '" & RYW0 & "' and '" & RYW1 & "'"
            P_SI = "SOTINVH2.OPS_YYYYWW"
            P_ST = "OPS_YYYYWW"
        End If

        '& ", SUM(DECODE(" & P_ST & ",'" & Ps(P - 1) & "',QTY_EOW,0)) OHQTY" & PX _
        Dim sql_SI As String = ""
        Dim sql_ST As String = ""
        Dim sql_SIST As String = ""
        Dim sql_cols As String = ""
        For P As Integer = 1 To Periods
            Dim PX As String = "_P" & Format(P, "00")
            sql_SI &= ", SUM(DECODE(" & P_SI & ",'" & Ps(P - 1) & "',SOTINVH2.ORDR_QTY_SHIP,0)) SIQTY" & PX & ", 0 OHQTY" & PX & ", 0 STQTY" & PX & ", 0 STAMT" & PX & vbCrLf
            sql_ST &= ", 0 SIQTY" & PX _
                    & ", SUM(DECODE(OPS_YYYYWW,'" & WPs(P - 1) & "',QTY_EOW,0)) OHQTY" & PX _
                    & ", SUM(DECODE(" & P_ST & ",'" & Ps(P - 1) & "',QTY_SOLD,0)) STQTY" & PX _
                    & ", SUM(DECODE(" & P_ST & ",'" & Ps(P - 1) & "',AMT_SOLD,0)) STAMT" & PX & vbCrLf
            sql_SIST &= ", SUM (SIQTY" & PX & ") SIQTY" & PX _
                      & ", SUM (OHQTY" & PX & ") OHQTY" & PX _
                      & ", SUM (STQTY" & PX & ") STQTY" & PX _
                      & ", SUM (STAMT" & PX & ") STAMT" & PX & vbCrLf
            sql_cols &= ", SIQTY" & PX _
                      & ", OHQTY" & PX _
                      & ", STQTY" & PX _
                      & ", STAMT" & PX & vbCrLf
            grdSATSIST1.DisplayLayout.Bands(0).Groups(P + 1).Header.Caption = Ls(P - 1)
            grdSATSIST2.DisplayLayout.Bands(0).Groups(P + 1).Header.Caption = Ls(P - 1)
            grdSATSIST1.DisplayLayout.Bands(0).Groups(P + 1).Hidden = False
            grdSATSIST1.DisplayLayout.Bands(0).Groups(P + 1).Header.ToolTipText = ""
        Next

        If Periods < SIST_PMAX Then
            For P As Integer = Periods + 1 To SIST_PMAX
                grdSATSIST1.DisplayLayout.Bands(0).Groups(P + 1).Hidden = True
                grdSATSIST2.DisplayLayout.Bands(0).Groups(P + 1).Hidden = True
                grdSATSIST3.DisplayLayout.Bands(0).Groups(P + 1).Hidden = True
            Next
        End If

        Dim CODEs As String = Get_Multi_CODEs("SATSISTI")

        For Each TABLE_NAME As String In New String() {"SATSIST1", "SATSIST2", "SATSIST3"}

            Dim grd As UltraWinGrid.UltraGrid = IIf(TABLE_NAME = "SATSIST1", grdSATSIST1, grdSATSIST2)

            Dim CUST_CODE_data As String = CUST_CODE

            If TABLE_NAME = "SATSIST3" Then
                grd = grdSATSIST3

                '  CUST_CODE_data = ITEM_CODE_LY

                For P As Integer = 1 To Periods
                    Dim PX As String = "_P" & Format(P, "00")
                    sql_SI = Replace(sql_SI, "'" & Ps(P - 1) & "'", "'" & PsLY(P - 1) & "'")
                    sql_ST = Replace(sql_ST, "'" & Ps(P - 1) & "'", "'" & PsLY(P - 1) & "'")
                    grdSATSIST3.DisplayLayout.Bands(0).Groups(P + 1).Header.Caption = LsLY(P - 1)
                    grdSATSIST3.DisplayLayout.Bands(0).Groups(P + 1).Hidden = False
                    grdSATSIST1.DisplayLayout.Bands(0).Groups(P + 1).Header.ToolTipText = ""
                Next

                If optRANGE.Value = "P" Then
                    sqlp = " and ORDR_YYYYPP_UPDATED between '" & ASCMAIN1.Period_Calc(RYP0, -12) & "' and '" & ASCMAIN1.Period_Calc(RYP1, -12) & "'"
                Else
                    sqlp = " and OPS_YYYYWW between '" & ASCMAIN1.Week_Calc(RYW0, -52) & "' and '" & ASCMAIN1.Week_Calc(RYW1, -52) & "'"
                    If ASCMAIN1.CLIENT = "INT" Then
                        Dim LRYW0 As String = Format(Val(Mid(RYW0, 1, 4)) - 1, "0000") & Mid(RYW0, 5, 2)
                        Dim LRYW1 As String = Format(Val(Mid(RYW1, 1, 4)) - 1, "0000") & Mid(RYW1, 5, 2)
                        sqlp = " and OPS_YYYYWW between '" & LRYW0 & "' and '" & LRYW1 & "'"
                    End If
                End If

            End If

            ASCMAIN1.sql = "Select X.ITEM_CODE" & vbCrLf _
                & IIf(TABLE_NAME = "SATSIST2", ",X.CUST_STORE_NO", "") & vbCrLf _
                & IIf(TABLE_NAME = "SATSIST2", ",ARTCUST2.CUST_STORE_NAME", ",ICTITEM1.ITEM_DESC") & vbCrLf _
                & ", Min (X.INV_DATE_SHIPPED) INV_DATE_SHIPPED" & vbCrLf _
                & sql_SIST & " from (" & vbCrLf _
                & "Select SOTINVH2.ITEM_CODE" & vbCrLf _
                & IIf(TABLE_NAME = "SATSIST2", ",SOTINVH1.CUST_STORE_NO", "") & vbCrLf _
                & ", Min (SOTINVH1.INV_DATE_SHIPPED) INV_DATE_SHIPPED" & vbCrLf _
                & sql_SI _
                & " from SOTINVH1,SOTINVH2" & vbCrLf _
                & " where SOTINVH2.CUST_CODE = '" & CUST_CODE_data & "'" & vbCrLf _
                & IIf(chkSIST_Multi.Checked And CODEs <> "", " and SOTINVH2.ITEM_CODE in (" & CODEs & ")" & vbCrLf, "") _
                & IIf(TABLE_NAME = "SATSIST2", " and SOTINVH2.ITEM_CODE = 'ITEM_CODE'", "") & vbCrLf _
                & Replace(Replace(sqlp, "ORDR_YYYYPP_UPDATED", "SOTINVH2.ORDR_YYYYPP_UPDATED"), "OPS_YYYYWW", "SOTINVH2.OPS_YYYYWW") & vbCrLf _
                & " and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
                & " and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
                & " and SOTINVH1.INV_TYPE = 'I'" & vbCrLf _
                & " group by SOTINVH2.ITEM_CODE" & vbCrLf _
                & IIf(TABLE_NAME = "SATSIST2", ",SOTINVH1.CUST_STORE_NO", "") & vbCrLf _
                & " union" & vbCrLf _
                & "Select ITEM_CODE" & vbCrLf _
                & IIf(TABLE_NAME = "SATSIST2", ",CUST_STORE_NO", "") & vbCrLf _
                & ", NULL INV_DATE_SHIPPED" & vbCrLf _
                & sql_ST _
                & " from RSTRETL1" & vbCrLf _
                & " where CUST_CODE = '" & CUST_CODE_data & "'" & vbCrLf _
                & IIf(chkSIST_Multi.Checked And CODEs <> "", " and ITEM_CODE in (" & CODEs & ")" & vbCrLf, "") _
                & IIf(TABLE_NAME = "SATSIST2", " and ITEM_CODE = 'ITEM_CODE'", "") & vbCrLf _
                & Replace(sqlp, "ORDR_YYYYPP_UPDATED", "OPS_YYYYPP") & vbCrLf _
                & " group by ITEM_CODE" & vbCrLf _
                & IIf(TABLE_NAME = "SATSIST2", ",CUST_STORE_NO", "") & vbCrLf _
                & ") X" & vbCrLf _
                & IIf(TABLE_NAME = "SATSIST2", _
                      ",ARTCUST2 where ARTCUST2.CUST_CODE (+) = '" & CUST_CODE & "' and ARTCUST2.CUST_STORE_NO (+) = X.CUST_STORE_NO", _
                      ",ICTITEM1 where ICTITEM1.ITEM_CODE (+) = X.ITEM_CODE") & vbCrLf _
                & " group by X.ITEM_CODE" & vbCrLf _
                & IIf(TABLE_NAME = "SATSIST2", ",X.CUST_STORE_NO", "") & vbCrLf _
                & IIf(TABLE_NAME = "SATSIST2", ",ARTCUST2.CUST_STORE_NAME", ",ICTITEM1.ITEM_DESC")

            If TABLE_NAME = "SATSIST1" Then
                If SATSIST1 = "" Then
                    SATSIST1 = ASCMAIN1.Temp_Table
                    ASCDATA1.ExecuteSQL("Alter Table " & SATSIST1 & " Add Primary Key (ITEM_CODE)")
                Else
                    ASCDATA1.ExecuteSQL("Truncate Table " & SATSIST1)
                    ASCDATA1.ExecuteSQL("Insert into " & SATSIST1 & " (ITEM_CODE, ITEM_DESC, INV_DATE_SHIPPED" & vbCrLf & sql_cols & ") " & ASCMAIN1.sql)
                End If

                Fill_and_Cum(SATSIST1, "SATSIST1")

            ElseIf TABLE_NAME = "SATSIST3" Then
                If SATSIST3 = "" Then
                    SATSIST3 = ASCMAIN1.Temp_Table
                    ASCDATA1.ExecuteSQL("Alter Table " & SATSIST3 & " Add Primary Key (ITEM_CODE)")
                Else
                    ASCDATA1.ExecuteSQL("Truncate Table " & SATSIST3)
                    ASCDATA1.ExecuteSQL("Insert into " & SATSIST3 & " (ITEM_CODE, ITEM_DESC, INV_DATE_SHIPPED" & vbCrLf & sql_cols & ") " & ASCMAIN1.sql)
                End If

                Fill_and_Cum(SATSIST3, "SATSIST3")

            Else
                If SATSIST2 = "" Then
                    SATSIST2 = ASCMAIN1.Temp_Table
                    ASCDATA1.ExecuteSQL("Alter Table " & SATSIST2 & " Add Primary Key (ITEM_CODE,CUST_STORE_NO)")

                    'SINCE WE ARE WIPING OUT NAME OF TABLE SATSIST2 FOR EACH ITEM WE NEVER HIT THE ELSE
                    sqlSATSIST2 = "Insert into " & SATSIST2 & " (ITEM_CODE, CUST_STORE_NO, CUST_STORE_NAME, INV_DATE_SHIPPED" & vbCrLf & sql_cols & ") " & ASCMAIN1.sql

                Else
                    ASCDATA1.ExecuteSQL("Truncate Table " & SATSIST2)

                    sqlSATSIST2 = "Insert into " & SATSIST2 & " (ITEM_CODE, CUST_STORE_NO, INV_DATE_SHIPPED" & vbCrLf & sql_cols & ") " & ASCMAIN1.sql
                    'ASCDATA1.ExecuteSQL(sqlSATSIST2)

                    'ASCDATA1.ExecuteSQL("Insert into " & SATSIST2 & " (CUST_CODE, CUST_STORE_NO, INV_DATE_SHIPPED" & vbCrLf & sql_cols & ") " & ASCMAIN1.sql)
                End If

                Fill_and_Cum(SATSIST2, "SATSIST2")

            End If

            dst.Tables(TABLE_NAME).Columns("SIQTY_P00").Expression = "ISNULL(SIQTY_P" & Format(Periods, "00") & ",0)"
        Next
        Sort_grdColumns(grdSATSIST1, "ITEM_CODE")
        Setup_grdSATSIST2()
        Sort_grdColumns(grdSATSIST3, "ITEM_CODE")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub grdSATSIST1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSATSIST1.AfterRowActivate
        Setup_grdSATSIST2()
    End Sub

    Private Sub grdSATSIST1_AfterColRegionScroll(sender As Object, e As UltraWinGrid.ColScrollRegionEventArgs) Handles grdSATSIST1.AfterColRegionScroll
        grdSATSIST2.DisplayLayout.ColScrollRegions(0).Position = e.ColScrollRegion.Position
        grdSATSIST3.DisplayLayout.ColScrollRegions(0).Position = e.ColScrollRegion.Position
    End Sub

    Private Sub grdSATSIST2_AfterColRegionScroll(sender As Object, e As UltraWinGrid.ColScrollRegionEventArgs) Handles grdSATSIST2.AfterColRegionScroll
        grdSATSIST1.DisplayLayout.ColScrollRegions(0).Position = e.ColScrollRegion.Position
    End Sub

    Private Sub grdSATSIST3_AfterColRegionScroll(sender As Object, e As UltraWinGrid.ColScrollRegionEventArgs) Handles grdSATSIST3.AfterColRegionScroll
        grdSATSIST1.DisplayLayout.ColScrollRegions(0).Position = e.ColScrollRegion.Position
    End Sub
    Sub Setup_grdSATSIST2()
        'Exit Sub
        If grdSATSIST1.ActiveRow Is Nothing OrElse Not grdSATSIST1.ActiveRow.IsDataRow Then
            grdSATSIST2.Visible = False
        Else
            Dim ITEM_CODE As String = grdSATSIST1.ActiveRow.Cells("ITEM_CODE").Value & ""

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Loading Store Level Data")

            ASCDATA1.ExecuteSQL("Truncate Table " & SATSIST2)
            ASCDATA1.ExecuteSQL(Replace(sqlSATSIST2, "'ITEM_CODE'", "'" & ITEM_CODE & "'"))

            'Dim DVW As DataView = DirectCast(grdSATSIST2.DataSource, DataTable).DefaultView
            'DVW.RowFilter = "CUST_CODE = '" & CUST_CODE & "'"

            Fill_and_Cum(SATSIST2, "SATSIST2")

            grdSATSIST2.Visible = True
            grdSATSIST2.Text = "Activity by Store for " & ITEM_CODE
            'Sort_grdColumns(grdSATSIST2, "CUST_STORE_NO")
            Sort_grdColumns(grdSATSIST2, "CUST_STORE_NO")

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")

        End If
    End Sub

    Sub Fill_and_Cum(TABLE_NAME_ora As String, TABLE_NAME_ado As String)
        ASCMAIN1.sql = "Select * from " & TABLE_NAME_ora
        Fill_Records(TABLE_NAME_ado, "", True, ASCMAIN1.sql)
        For Each row As DataRow In dst.Tables(TABLE_NAME_ado).Select("")
            'Dim CUST_CODE As String = row.Item("CUST_CODE")
            For P As Integer = 2 To Periods
                row.Item("SIQTY_P" & Format(P, "00")) = Val(row.Item("SIQTY_P" & Format(P, "00")) & "") + Val(row.Item("SIQTY_P" & Format(P - 1, "00")) & "")
            Next
        Next
    End Sub
    Private Sub grdSATCSLS1_AfterColRegionScroll(sender As Object, e As UltraWinGrid.ColScrollRegionEventArgs) Handles grdSATCSLS1.AfterColRegionScroll
        grdSATCSLS2.DisplayLayout.ColScrollRegions(0).Position = e.ColScrollRegion.Position
    End Sub

    Private Sub grdSATCSLS2_AfterColRegionScroll(sender As Object, e As UltraWinGrid.ColScrollRegionEventArgs) Handles grdSATCSLS2.AfterColRegionScroll
        grdSATCSLS1.DisplayLayout.ColScrollRegions(0).Position = e.ColScrollRegion.Position
    End Sub


    Public Overrides Function CustomSummary_End( _
        ByVal summarySettings As UltraWinGrid.SummarySettings, _
        ByVal rows As UltraWinGrid.RowsCollection, _
        ByVal CustomValue As Double, _
        ByVal grd As UltraWinGrid.UltraGrid) As Double

        CustomValue = 0
        Dim TOTALS As New Dictionary(Of String, Decimal)
        Dim KEY As String = summarySettings.Key

        Select Case grd.Name
            Case "grdSATSIST1", "grdSATSIST2", "grdSATSIST3"
                If KEY.StartsWith("STPCT") Then
                    Dim SI As String = "SIQTY_P" & Mid(KEY, 8)
                    Dim ST As String = "STQTY_P" & Mid(KEY, 8)
                    TOTALS.Add(SI, 0)
                    TOTALS.Add(ST, 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS(SI) <> 0 Then CustomValue = 100 * TOTALS(ST) / TOTALS(SI)
                Else
                    Stop
                End If

            Case "grdSATCSLS1", "grdSATCSLS2", "grdSATCSLS3"
                If KEY = "PYY00PCT" Then
                    TOTALS.Add("PYY", 0)
                    TOTALS.Add("P00", 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS("PYY") <> 0 Then CustomValue = 100 * TOTALS("P00") / TOTALS("PYY")
                End If
            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Public Overrides Function CustomStringSummary_End( _
        ByVal summarySettings As UltraWinGrid.SummarySettings, _
        ByVal rows As UltraWinGrid.RowsCollection, _
        ByVal CustomValue As String, _
        ByVal grd As UltraWinGrid.UltraGrid) As String

        Select Case grd.Name
            Case "grdSATSIST1"
                Dim KEY As String = summarySettings.Key
                CustomValue = "Totals"
            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Sub CustomSummary_Calculate_Totals( _
       ByVal rows As UltraWinGrid.RowsCollection, _
       ByRef TOTALS As Dictionary(Of String, Decimal), _
       ByVal KEY As String)

        For Each grow2 As UltraWinGrid.UltraGridRow In rows
            If grow2.IsGroupByRow Then
                Dim gbrow As UltraWinGrid.UltraGridGroupByRow = DirectCast(grow2, UltraWinGrid.UltraGridGroupByRow)
                CustomSummary_Calculate_Totals(gbrow.Rows, TOTALS, KEY)
            Else
                If KEY.StartsWith("STPCT") Then
                    Dim SI As String = "SIQTY_P" & Mid(KEY, 8)
                    Dim ST As String = "STQTY_P" & Mid(KEY, 8)
                    TOTALS(SI) += Val(grow2.Cells(SI).Value & "")
                    TOTALS(ST) += Val(grow2.Cells(ST).Value & "")
                ElseIf KEY = "TRADE_CLASS_CODE" Then
                    '  TOTALS(KEY) = "Totals"
                ElseIf KEY = "PYY00PCT" Then
                    TOTALS("PYY") += Val(grow2.Cells("PYY").Value & "")
                    TOTALS("P00") += Val(grow2.Cells("P00").Value & "")
                Else
                End If
            End If
        Next
    End Sub


    Private Sub grdSPTCOOPX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs)
        With e.Row.Cells("APPR_STATUS_CODE")
            Select Case .Value & ""
                Case "A"
                    .Appearance.ForeColor = System.Drawing.Color.Green
                Case "P"
                    .Appearance.ForeColor = System.Drawing.Color.Purple
                Case "G"
                    .Appearance.ForeColor = System.Drawing.Color.Blue
                Case "X"
                    .Appearance.ForeColor = System.Drawing.Color.Red
            End Select

        End With
    End Sub

    Sub Excel_Extract( _
                         Optional ByRef workbook As SpreadsheetGear.IWorkbook = Nothing, _
                         Optional ByRef XLS_FILENAME As String = "", _
                         Optional ByRef Start_Row As Integer = 0)

        Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing

        'Dim XLS_FILENAME As String = ""
        'Dim workbook As SpreadsheetGear.IWorkbook

        If XLS_FILENAME_multi = "" Or XLS_FILENAME_multi = "*" Then
            XLS_FILENAME = ASCMAIN1.Folders("Work") & ASCMAIN1.Next_Control_No(Me.Name & "_SIST") & "SIST.xlsX"
            workbook = SpreadsheetGear.Factory.GetWorkbook()
            If XLS_FILENAME_multi = "*" Then
                XLS_FILENAME_multi = XLS_FILENAME
            End If
            worksheet = workbook.Worksheets(0)
        Else
            If XLS_FILENAME = "" Then
                XLS_FILENAME = XLS_FILENAME_multi
                workbook = SpreadsheetGear.Factory.GetWorkbook(XLS_FILENAME_multi)
                worksheet = workbook.Worksheets.Add
            Else
                worksheet = workbook.Worksheets(0)
            End If
        End If

        Dim CUST_CODE As String = Me.CUST_CODE
        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)

        Dim Total_Cols As Integer = 10

        If XLS_FILENAME_multi = "" Or chkSIST_SingleSheet.Checked Then
            worksheet.Name = "Sell-In Sell-Thru"
        Else
            worksheet.Name = CUST_CODE
        End If

        If Start_Row = 0 Then
            worksheet.Cells(Start_Row, 0).Value = "'" & Format(Now, "MM/dd/yy")
            worksheet.Cells(Start_Row, 1).Value = ASCMAIN1.USER_ID
        End If

        Dim Rx As Integer = Start_Row + 3

        With worksheet.Cells(Rx, 0, Rx + 3, 5)
            .Interior.Color = SpreadsheetGear.Colors.WhiteSmoke
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
        End With

        With worksheet.Cells(Rx, 0)
            .Value = "'" & CUST_CODE
            .Font.Underline = SpreadsheetGear.UnderlineStyle.Single
            .Font.Size = 14
        End With
        With worksheet.Cells(Rx + 1, 0)
            .Value = "'" & rowARTCUST1.Item("CUST_NAME")
            .Font.Color = SpreadsheetGear.Colors.Purple
            .Font.Name = "Times New Roman"
            .Font.Size = 20
        End With
        'With worksheet.Cells(Rx + 2, 0)
        '    .Value = "Retail: " & Format(Val(rowARTCUST1.Item("ITEM_RETAIL_PRICE") & ""), "$#.00")
        '    .Font.Size = 12
        'End With
        'With worksheet.Cells(Rx + 2, 1)
        '    .Value = Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & "")
        '    .NumberFormat = "$#0.00"
        '    .Font.Size = 12
        'End With
        'With worksheet.Cells(Rx + 3, 0)
        '    .Value = "Plan " & Format(QTY_ALLO_PLAN, "#,##0")
        '    .Font.Size = 12
        'End With
        'With worksheet.Cells(Rx + 3, 1)
        '    .Value = 1000
        '    .NumberFormat = "#,##0"
        '    .Font.Size = 12
        'End With
        'If rowARTCUST1.Item("ITEM_DATE_TO_SHIP") & "" <> "" Then
        '    With worksheet.Cells(Rx + 4, 0)
        '        .Value = "Ship Date " & Format(rowARTCUST1.Item("ITEM_DATE_TO_SHIP"), "MM/dd/yy")
        '        .Font.Size = 12
        '    End With
        '    'With worksheet.Cells(Rx + 4, 1)
        '    '    .Value = Now.Date
        '    '    .NumberFormat = "MM/DD/YY"
        '    '    .Font.Size = 12
        '    'End With
        'End If



        'Dim IMAGE_NAME As String = CUST_CODE
        'Dim IMAGE_FOLDER As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
        ''Dim imgba() As Byte = Nothing
        ''ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, False, , , imgba)

        'Dim imageFileStyle As String = IMAGE_FOLDER & "\" & IMAGE_NAME & ".jpg"

        'If My.Computer.FileSystem.FileExists(imageFileStyle) Then
        '    'Dim SourceImagePath As String = IMAGE_FOLDER & "\" & IMAGE_NAME
        '    'Dim imageFileStyle As String = IMAGE_FOLDER & "\resized\" & IMAGE_NAME
        '    'Dim original As Image = Image.FromFile(SourceImagePath)
        '    'Dim resized As Image = ResizeImage(original, New Size(1024, 1024))
        '    'SaveImageWithQuality(resized, imageFileStyle, 75L)

        '    'Dim ImageRows As Integer = 0
        '    'Dim ImageRowsBig As Integer = 0

        '    Dim widthStyle As Double
        '    Dim heightStyle As Double

        '    Dim imageStyle As System.Drawing.Image = System.Drawing.Image.FromFile(imageFileStyle)
        '    Try
        '        widthStyle = imageStyle.Width * 72.0 / imageStyle.HorizontalResolution * 0.75
        '        heightStyle = imageStyle.Height * 72.0 / imageStyle.VerticalResolution * 0.75
        '    Finally
        '        imageStyle.Dispose()
        '    End Try

        '    ' Calculate the left and top placement of the picture by converting 
        '    ' row and column coordinates to points.  Use fractional values to 
        '    ' get coordinates anywhere in between row and column boundaries.
        '    Dim windowInfoStyle As SpreadsheetGear.IWorksheetWindowInfo = worksheet.WindowInfo

        '    'Dim col_adj As Decimal = 0

        '    Dim leftStyle As Double = windowInfoStyle.ColumnToPoints(7) ' + col_adj
        '    Dim topStyle As Double = windowInfoStyle.RowToPoints(Rx) ' + 0.1 ' 1.5)

        '    '   ImageRows = windowInfoStyle.PointsToRow(heightStyle)

        '    ' Add the picture from file.
        '    worksheet.Shapes.AddPicture(imageFileStyle, leftStyle, topStyle, widthStyle, heightStyle)
        'End If


        Dim COLs_per_Group As Integer = 5
        If ASCMAIN1.CLIENT = "AHA" Then
            COLs_per_Group = 6
        End If

        Rx += 5

        Dim ITEM_CODEs As New Dictionary(Of String, Integer)

        For YR As Integer = 0 To 1
            If YR = 1 Then
                If dst.Tables("SATSIST3").Select("").Length = 0 Then
                    Exit For
                End If
            End If

            Dim RxG As Integer = Rx
            Dim CxD As Integer = 0

            CxD = 0
            With worksheet.Cells(Rx + 1, 0 + CxD)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Left
                .ColumnWidth = 12
                .Value = "Item"
            End With
            CxD += 1
            With worksheet.Cells(Rx + 1, 0 + CxD)
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
                .ColumnWidth = 24
                .Value = "Description"
            End With
            CxD += 1
            With worksheet.Cells(Rx + 1, 0 + CxD)
                .EntireColumn.NumberFormat = "@"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
                .ColumnWidth = 8
                .Value = "3PL"
            End With
            CxD += 1
            With worksheet.Cells(Rx + 1, 0 + CxD)
                .EntireColumn.NumberFormat = "###.00"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .ColumnWidth = 8
                .Value = "Retail"
            End With
            CxD += 1
            With worksheet.Cells(Rx + 1, 0 + CxD)
                .EntireColumn.NumberFormat = "MM/DD/YY"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
                .ColumnWidth = 12
                .Value = "Shipped"
            End With

            For P As Integer = 0 To Periods
                With worksheet.Cells(Rx, CxD + P * COLs_per_Group + 1, Rx, CxD + P * COLs_per_Group + COLs_per_Group)
                    .Merge()
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center

                    worksheet.Cells(RxG - 1, CxD + P * COLs_per_Group + 1, RxG - 1, CxD + P * COLs_per_Group + COLs_per_Group).Merge()
                    worksheet.Cells(RxG - 1, CxD + P * COLs_per_Group + 1, RxG - 1, CxD + P * COLs_per_Group + COLs_per_Group).Font.Italic = True

                    If P = 0 Then
                        .Value = "Totals"
                    Else
                        If YR = 1 Then
                            .Value = LsLY(P - 1)
                        Else
                            .Value = Ls(P - 1)
                        End If
                    End If

                    With .Interior
                        .Color = SpreadsheetGear.Colors.Orange
                    End With

                    .Borders.LineStyle = SpreadsheetGear.LineStyle.Continuous
                End With

                worksheet.Cells(Rx, 0).Value = IIf(YR = 0, "This Year", "Last Year") ' & IIf(ITEM_CODE_LY = ITEM_CODE, "", " (" & ITEM_CODE_LY & ")")

                With worksheet.Cells(Rx + 1, 0, Rx + 1, CxD + Periods * COLs_per_Group + COLs_per_Group)
                    .Interior.Color = SpreadsheetGear.Colors.LightBlue
                    .Borders.LineStyle = SpreadsheetGear.LineStyle.Continuous
                End With

                Dim iCol As Integer = 0

                iCol += 1
                With worksheet.Cells(Rx + 1, CxD + P * COLs_per_Group + iCol)
                    .EntireColumn.NumberFormat = "#,##0"
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    .Value = "Sell-In"
                End With
                iCol += 1
                With worksheet.Cells(Rx + 1, CxD + P * COLs_per_Group + iCol)
                    .EntireColumn.NumberFormat = "#,##0"
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    .Value = "Sell-Thru"
                End With
                If COLs_per_Group > 5 Then
                    iCol += 1
                    With worksheet.Cells(Rx + 1, CxD + P * COLs_per_Group + iCol)
                        .EntireColumn.NumberFormat = "#,##0.0%"
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .Value = "Y/Y%"
                    End With
                End If
                iCol += 1
                With worksheet.Cells(Rx + 1, CxD + P * COLs_per_Group + iCol)
                    .EntireColumn.NumberFormat = "#,##0"
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    .Value = "On Hand"
                End With
                iCol += 1
                With worksheet.Cells(Rx + 1, CxD + P * COLs_per_Group + iCol)
                    .EntireColumn.NumberFormat = "#,##0.0"
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    .Value = "SIST%"
                End With
                iCol += 1
                With worksheet.Cells(Rx + 1, CxD + P * COLs_per_Group + iCol)
                    .EntireColumn.NumberFormat = "#,##0"
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    .Value = "$Thru"
                End With
            Next

            Rx += 2
            Dim iRx As Integer = Rx

            Dim SIST As String = "SATSIST1"
            If YR = 1 Then SIST = "SATSIST3"

            If dst.Tables(SIST).Select("").Length > 0 Then
                For Each row As DataRow In dst.Tables(SIST).Select("", "ITEM_CODE")
                    Dim ITEM_CODE As String = row.Item("ITEM_CODE")
                    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE, True)
                    Dim COLLECTION_CODE As String = rowICTITEM1.Item("COLLECTION_CODE") & ""
                    Dim rowICTCOLL1 As DataRow = LookUp("ICTCOLL1", COLLECTION_CODE)
                    Dim HC_CODE As String = ""
                    If rowICTCOLL1 IsNot Nothing Then HC_CODE = rowICTCOLL1.Item("HC_CODE") & ""

                    Dim CxDi As Integer = 0
                    worksheet.Cells(Rx, 0 + CxDi).Value = ITEM_CODE
                    CxDi += 1 : worksheet.Cells(Rx, 0 + CxDi).Value = rowICTITEM1.Item("ITEM_DESC")
                    CxDi += 1 : worksheet.Cells(Rx, 0 + CxDi).Value = rowICTITEM1.Item("ITEM_ALT_SORT")
                    CxDi += 1 : worksheet.Cells(Rx, 0 + CxDi).Value = Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & "")
                    CxDi += 1 : worksheet.Cells(Rx, 0 + CxDi).Value = row.Item("INV_DATE_SHIPPED")

                    If YR = 0 Then
                        ITEM_CODEs.Add(ITEM_CODE, Rx)
                    End If
                    For P As Integer = 0 To Periods
                        Dim iCol As Integer = 0
                        iCol += 1 : worksheet.Cells(Rx, CxD + P * COLs_per_Group + iCol).Formula = row.Item("SIQTY_P" & Format(P, "00")) & ""
                        iCol += 1 : worksheet.Cells(Rx, CxD + P * COLs_per_Group + iCol).Formula = row.Item("STQTY_P" & Format(P, "00")) & ""
                        If COLs_per_Group > 5 Then
                            If YR = 0 Then
                                iCol += 1
                            Else
                                iCol += 1
                                If ITEM_CODEs.ContainsKey(ITEM_CODE) Then
                                    Dim RxTY As Integer = ITEM_CODEs(ITEM_CODE)
                                    Dim LY As String = Excel_Cell0(Rx, CxD + P * COLs_per_Group + iCol - 1)
                                    Dim TY As String = Excel_Cell0(RxTY, CxD + P * COLs_per_Group + iCol - 1)
                                    worksheet.Cells(RxTY, CxD + P * COLs_per_Group + iCol).Formula = "=IFERROR((" & TY & "-" & LY & ")/" & LY & ",0)"
                                End If
                            End If
                        End If
                        iCol += 1 : worksheet.Cells(Rx, CxD + P * COLs_per_Group + iCol).Formula = row.Item("OHQTY_P" & Format(P, "00")) & ""
                        iCol += 1 : worksheet.Cells(Rx, CxD + P * COLs_per_Group + iCol).Formula = row.Item("STPCT_P" & Format(P, "00")) & ""
                        iCol += 1 : worksheet.Cells(Rx, CxD + P * COLs_per_Group + iCol).Formula = row.Item("STAMT_P" & Format(P, "00")) & ""
                    Next

                    Dim sqlw As String = ""
                    If optRANGE.Value = "P" Then
                        If YR = 0 Then
                            sqlw = " and OPS_YYYYPP >= '" & Ps(0) & "' and OPS_YYYYPP <= '" & Ps(Periods - 1) & "'"
                        Else
                            sqlw = " and OPS_YYYYWW >= '" & PsLY(0) & "' and OPS_YYYYWW <= '" & PsLY(Periods - 1) & "'"
                        End If
                    Else
                        If YR = 0 Then
                            sqlw = " and OPS_YYYYWW >= '" & Ps(0) & "' and OPS_YYYYWW <= '" & Ps(Periods - 1) & "'"
                        Else
                            sqlw = " and OPS_YYYYWW >= '" & PsLY(0) & "' and OPS_YYYYWW <= '" & PsLY(Periods - 1) & "'"
                        End If
                    End If

                    'Dim HC_CODEs As String = ""
                    'For Each rowSATSIST1 As DataRow In dst.Tables("SATSIST1").Select("")
                    '    Dim ITEM_CODE As String = rowSATSIST1.Item("ITEM_CODE")

                    'Next

                    If HC_CODE <> "" Then sqlw &= " and HC_CODE = '" & HC_CODE & "'"
                    Dim rows() As DataRow = dst.Tables("SPTCOOPX").Select("CUST_CODE = '" & CUST_CODE & "'" & sqlw)
                    For Each rowSP As DataRow In rows
                        Dim DATE_START As Date = rowSP.Item("DATE_START")
                        Dim DATE_END As Date = rowSP.Item("DATE_END")
                        Dim EVENT_TYPE_CODE As String = rowSP.Item("EVENT_TYPE_CODE") & ""
                        For P As Integer = 1 To Periods
                            Dim PDT1 As Date = IIf(YR = 0, PE_DATEs(P - 1), PE_DATEsLY(P - 1))
                            Dim PDT2 As Date = IIf(YR = 0, PE_DATEs(P - 0), PE_DATEsLY(P - 0))
                            If Format(DATE_START, "yyyyMMdd") > Format(PDT1, "yyyyMMdd") And _
                               Format(DATE_START, "yyyyMMdd") <= Format(PDT2, "yyyyMMdd") Then

                                Dim T As String = Format(DATE_START, "MM/dd") & " " & CUST_CODE & " " & rowSP.Item("BOOKING_NAME") & " " & rowSP.Item("FEATURE_DESC")
                                If ASCMAIN1.CLIENT = "AHA" Then
                                    T = Format(DATE_START, "MM/dd") & "-" & Format(DATE_END, "MM/dd") & " " & EVENT_TYPE_CODE & " " & CUST_CODE & " " & rowSP.Item("BOOKING_NAME") & " " & rowSP.Item("FEATURE_DESC")
                                End If

                                If ASCMAIN1.CLIENT = "AHA" Then
                                Else
                                    With worksheet.Cells(Rx, CxD + P * COLs_per_Group + 2)
                                        .Interior.Color = SpreadsheetGear.Colors.Violet
                                    End With
                                End If

                                With worksheet.Cells(RxG - 1, CxD + P * COLs_per_Group + 1)
                                    If .Value & "" = "" Then
                                        .Value = T
                                    Else
                                        If Not .Value.ToString.Contains(T) Then
                                            .Value &= vbCrLf & T
                                        End If
                                    End If
                                End With
                            End If
                        Next
                    Next

                    Rx += 1
                Next


                For P As Integer = 0 To Periods
                    Dim Cx As Integer = CxD + P * COLs_per_Group

                    With worksheet.Cells(iRx, Cx + COLs_per_Group, Rx, Cx + COLs_per_Group)
                        .Interior.Color = SpreadsheetGear.Colors.LightGreen
                    End With

                    With worksheet.Cells(iRx, Cx + 1, Rx, Cx + COLs_per_Group)
                        .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    End With

                    Cx += 1 : worksheet.Cells(Rx, Cx).Formula = "=SUM(" & Excel_Cell0(iRx, Cx) & ":" & Excel_Cell0(Rx - 1, Cx) & ")"
                    Cx += 1 : worksheet.Cells(Rx, Cx).Formula = "=SUM(" & Excel_Cell0(iRx, Cx) & ":" & Excel_Cell0(Rx - 1, Cx) & ")"
                    If ASCMAIN1.CLIENT = "AHA" Then
                        worksheet.Range(Excel_Cell0(iRx, Cx) & ":" & Excel_Cell0(Rx - 1, Cx)).Interior.Color = SpreadsheetGear.Colors.Yellow
                    End If
                    If COLs_per_Group > 5 Then
                        Cx += 1 ' : worksheet.Cells(Rx, Cx).Formula = "=SUM(" & Excel_Cell0(iRx, Cx) & ":" & Excel_Cell0(Rx - 1, Cx) & ")"
                    End If
                    Cx += 1 : worksheet.Cells(Rx, Cx).Formula = "=SUM(" & Excel_Cell0(iRx, Cx) & ":" & Excel_Cell0(Rx - 1, Cx) & ")"
                    Cx += 1 : worksheet.Cells(Rx, Cx).Formula = String.Format("=IF({0}=0,0,100*{1}/{0})", Excel_Cell0(Rx, Cx - 3), Excel_Cell0(Rx, Cx - 2))
                    Cx += 1 : worksheet.Cells(Rx, Cx).Formula = "=SUM(" & Excel_Cell0(iRx, Cx) & ":" & Excel_Cell0(Rx - 1, Cx) & ")"
                Next
            End If

            With worksheet.Cells(Rx, 0, Rx, CxD + Periods * COLs_per_Group + COLs_per_Group)
                .Borders.LineStyle = SpreadsheetGear.LineStyle.Continuous
                .Interior.Color = SpreadsheetGear.Colors.LightGray
            End With
            worksheet.Cells(Rx, 0).Value = "Totals"

            worksheet.Cells(RxG - 1, 0).Rows.AutoFit()
            worksheet.Cells(RxG - 1, 0).Rows.RowHeight = 60
            ' worksheet.Cells(RxG - 1, 0).Rows.Font.Italic = True

            Rx += 3
        Next

        Start_Row = Rx

        workbook.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        If XLS_FILENAME_multi = "" Then
            ' Show_Document(XLS_FILENAME)
            Add_Document_to_ASTSPRF1(XLS_FILENAME)
        End If

    End Sub

    Private Sub grdSATSIST1_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSATSIST1.InitializeRow
        If 1 = 1 Then Exit Sub
        If e.Row.IsDataRow Then

            Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Value

            Dim sqlw As String = ""
            If optRANGE.Value = "P" Then
                sqlw = " and OPS_YYYYPP >= '" & Ps(0) & "' and OPS_YYYYPP <= '" & Ps(Periods - 1) & "'"
            Else
                sqlw = " and OPS_YYYYWW >= '" & Ps(0) & "' and OPS_YYYYWW <= '" & Ps(Periods - 1) & "'"
            End If

            Dim rows() As DataRow = dst.Tables("SPTCOOPX").Select("CUST_CODE = '" & CUST_CODE & "'" & sqlw)
            For Each row As DataRow In rows
                Dim DATE_START As Date = row.Item("DATE_START")
                For P As Integer = 1 To Periods
                    If Format(DATE_START, "yyyyMMdd") > Format(PE_DATEs(P - 1), "yyyyMMdd") And _
                       Format(DATE_START, "yyyyMMdd") <= Format(PE_DATEs(P), "yyyyMMdd") Then
                        Dim C As String = "STQTY_P" & Format(P, "00")
                        Dim T As String = Format(DATE_START, "MM/dd") & " " & CUST_CODE & " " & row.Item("BOOKING_NAME") ' & " " & row.Item("VEHICLE_CODE")
                        Dim TT As String = e.Row.Cells(C).Column.Group.Header.ToolTipText & ""
                        If Not TT.Contains(T & vbCrLf) Then
                            e.Row.Cells(C).Appearance.BackColor = Drawing.Color.Violet ' .LightGreen
                            e.Row.Cells(C).ToolTipText &= T & vbCrLf
                            e.Row.Cells(C).Column.Group.Header.ToolTipText &= T & vbCrLf
                        End If
                    End If
                Next
            Next
        End If
    End Sub

    Sub Add_Customers()
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("ITEM_CODE")

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            ASCMAIN1.CodeSelector.Custom_sql_where = ""
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading")

                For Each ITEM_CODE As String In ASCMAIN1.CodeSelector.SelectedCodes
                    Add_Customer(ITEM_CODE)
                Next
                grdSATSISTC.PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
            End If

            Sort_grdColumns(grdSATSISTC, "CUST_CODE")
        End If
    End Sub

    Function Add_Item(ITEM_CODE As String) As DataRow
        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
        If rowICTITEM1 Is Nothing Then
            Return Nothing
        Else
            Dim row As DataRow = dst.Tables("SATSISTI").Rows.Find(ITEM_CODE)
            If row Is Nothing Then
                grdSATSISTI.ActiveRow = grdSATSISTI.DisplayLayout.Bands(0).AddNew
                grdSATSISTI.ActiveRow.Cells("ITEM_CODE").Value = ITEM_CODE
                grdSATSISTI.ActiveRow.Update()

                'row = dst.Tables("SATSISTI").NewRow
                'row.Item("ITEM_CODE") = rowICTITEM1.Item("ITEM_CODE")
                'row.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
                'row.Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
                'row.Item("COLLECTION_CODE") = rowICTITEM1.Item("COLLECTION_CODE")
                'row.Item("PROD_CODE") = rowICTITEM1.Item("PROD_CODE")
                'row.Item("ITEM_CATGY_CODE") = rowICTITEM1.Item("ITEM_CATGY_CODE")
                'row.Item("ITEM_ALT_SORT") = rowICTITEM1.Item("ITEM_ALT_SORT")
                'dst.Tables("SATSISTI").Rows.Add(row)
            End If
            Return row
        End If
    End Function

    Function Add_Customer(CUST_CODE As String) As DataRow
        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
        If rowARTCUST1 Is Nothing Then
            Return Nothing
        Else
            Dim row As DataRow = dst.Tables("SATSISTC").Rows.Find(CUST_CODE)
            If row Is Nothing Then
                grdSATSISTC.ActiveRow = grdSATSISTC.DisplayLayout.Bands(0).AddNew
                grdSATSISTC.ActiveRow.Cells("CUST_CODE").Value = CUST_CODE
                grdSATSISTC.ActiveRow.Update()

                'row = dst.Tables("SATSISTC").NewRow
                'row.Item("CUST_CODE") = rowARTCUST1.Item("CUST_CODE")
                'row.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME")
                'row.Item("TRADE_CLASS_CODE") = rowARTCUST1.Item("TRADE_CLASS_CODE")
                'row.Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE")
                'dst.Tables("SATSISTC").Rows.Add(row)
            End If
            Return rowARTCUST1
        End If
    End Function

    Private Sub grdSATSISTI_AfterCellUpdate(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSATSISTI.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
                Dim ITEM_CODE As String = e.Cell.Value & ""

                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    e.Cell.Row.Cells("ITEM_DESC").Value = rowICTITEM1.Item("ITEM_DESC")
                    e.Cell.Row.Cells("ITEM_RETAIL_PRICE").Value = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
                    e.Cell.Row.Cells("COLLECTION_CODE").Value = rowICTITEM1.Item("COLLECTION_CODE")
                    e.Cell.Row.Cells("PROD_CODE").Value = rowICTITEM1.Item("PROD_CODE")
                    e.Cell.Row.Cells("ITEM_CATGY_CODE").Value = rowICTITEM1.Item("ITEM_CATGY_CODE")
                    e.Cell.Row.Cells("ITEM_ALT_SORT").Value = rowICTITEM1.Item("ITEM_ALT_SORT")
                End If

        End Select
    End Sub

    Private Sub grdSATSISTI_AfterExitEditMode(sender As Object, e As EventArgs) Handles grdSATSISTI.AfterExitEditMode
        With grdSATSISTI
            Select Case .ActiveCell.Column.Key
                Case "ITEM_CODE"
                    Dim ITEM_CODE As String = .ActiveCell.Text
                    If ITEM_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(ITEM_CODE, .ActiveCell.Column.Key)
                    End If
            End Select
        End With
    End Sub

    Private Sub grdSATSISTI_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSATSISTI.AfterRowActivate
        'If e.Row.IsAddRow Then
        '    grdSATSISTI.DisplayLayout.Bands(0).Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        'Else
        '    grdSATSISTI.DisplayLayout.Bands(0).Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        'End If
    End Sub

    Private Sub grdSATSISTI_BeforeRowActivate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSATSISTI.BeforeRowActivate
        If e.Row.IsAddRow Then
            grdSATSISTI.DisplayLayout.Bands(0).Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            grdSATSISTI.DisplayLayout.Bands(0).Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        End If
    End Sub

    Private Sub grdSATSISTI_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSATSISTI.InitializeRow

    End Sub

    Private Sub grdSATSISTC_AfterCellUpdate(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSATSISTC.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "CUST_CODE"
                Dim CUST_CODE As String = e.Cell.Value & ""

                Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                If rowARTCUST1 IsNot Nothing Then
                    e.Cell.Row.Cells("CUST_NAME").Value = rowARTCUST1.Item("CUST_NAME")
                    e.Cell.Row.Cells("SREP_CODE").Value = rowARTCUST1.Item("SREP_CODE")
                    e.Cell.Row.Cells("TRADE_CLASS_CODE").Value = rowARTCUST1.Item("TRADE_CLASS_CODE")
                End If
        End Select
    End Sub

    Private Sub grdSATSISTC_AfterExitEditMode(sender As Object, e As EventArgs) Handles grdSATSISTC.AfterExitEditMode
        With grdSATSISTC
            Select Case .ActiveCell.Column.Key
                Case "CUST_CODE"
                    Dim CUST_CODE As String = .ActiveCell.Text
                    If CUST_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(CUST_CODE, .ActiveCell.Column.Key)
                    End If
            End Select
        End With
    End Sub

    Private Sub grdSATSISTC_BeforeRowActivate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSATSISTC.BeforeRowActivate
        If e.Row.IsAddRow Then
            grdSATSISTC.DisplayLayout.Bands(0).Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            grdSATSISTC.DisplayLayout.Bands(0).Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        End If
    End Sub

    Private Sub grdSATSISTC_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSATSISTC.InitializeRow

    End Sub

    Private Sub chkSIST_Multi_CheckedChanged(sender As Object, e As EventArgs) Handles chkSIST_Multi.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Setup_SM()
    End Sub

    Sub Setup_SM()
        splSATSIST0.Visible = Not ScreenMode And chkSIST_Multi.Checked
        lblCUST_CODE.Visible = Not chkSIST_Multi.Checked
        lblCUST_NAME.Visible = Not chkSIST_Multi.Checked
        txtCUST_CODE.Visible = Not chkSIST_Multi.Checked
        txtCUST_NAME.Visible = Not chkSIST_Multi.Checked
        cmdMultiXLS.Visible = Not ScreenMode And chkSIST_Multi.Checked
        chkSIST_Filter.Visible = chkSIST_Multi.Checked
        chkSIST_SingleSheet.Visible = chkSIST_Multi.Checked

        UltraExplorerBar1.Groups("Screen Control").Visible = (Not chkSIST_Multi.Checked)
        UltraExplorerBar1.Groups("Data Options").Visible = summary_loaded And (Not chkSIST_Multi.Checked)
        UltraExplorerBar1.Groups("Saved Lists").Visible = chkSIST_Multi.Checked

        If chkSIST_Multi.Checked Then
            Fill_Records("ASTLIST1")
            Sort_grdColumns(grdASTLIST1, "INIT_DATE".ToLower)
        End If
    End Sub

    Private Sub cmdMultiXLS_Click(sender As Object, e As EventArgs) Handles cmdMultiXLS.Click

        If chkSIST_Filter.Checked Then

            If dst.Tables("SATSISTC").Rows.Count <> 0 Or dst.Tables("SATSISTI").Rows.Count = 0 Then
                MsgBox("When using Customers from Items" _
                       & vbCrLf & " - A list of Customers will be created from the History of the Items Selected" _
                       & vbCrLf & " - Do NOT specify Customers" _
                       & vbCrLf & " - You MUST specify Items", _
                       MsgBoxStyle.OkOnly, "Do Not Specify Items when you use the Items from Customers option")
                Exit Sub
            End If

            Auto_Select()

            If dst.Tables("SATSISTC").Select("").Length = 0 Then
                MsgBox("There were no Customers with History in the Items Selected", MsgBoxStyle.OkOnly, "Cannot Generate XLS")
                Exit Sub
            End If
        End If

        If dst.Tables("SATSISTC").Select("").Length = 0 Then
            MsgBox("You must first add customers to the grid", MsgBoxStyle.OkOnly, "Cannot Generate XLS until you Select Items")
            Exit Sub
        End If

        XLS_FILENAME_multi = "*"

        Dim XLS_FILENAME As String = ""
        Dim workbook As SpreadsheetGear.IWorkbook = Nothing
        Dim Start_Row As Integer = 0

        lblCUST_CODE.Visible = True
        lblCUST_NAME.Visible = True
        txtCUST_CODE.Visible = True
        txtCUST_NAME.Visible = True

        For Each row As DataRow In dst.Tables("SATSISTC").Select("", "CUST_CODE")
            CUST_CODE = row.Item("CUST_CODE")
            Absx1.txtFor("CUST_CODE").Text = CUST_CODE
            Click_Command("View")

            lblCUST_CODE.Visible = True
            lblCUST_NAME.Visible = True
            txtCUST_CODE.Visible = True
            txtCUST_NAME.Visible = True

            If chkSIST_SingleSheet.Checked Then
                Excel_Extract(workbook, XLS_FILENAME, Start_Row)
            Else
                Excel_Extract()
            End If

            Click_Command("Done")
        Next

        lblCUST_CODE.Visible = False
        lblCUST_NAME.Visible = False
        txtCUST_CODE.Visible = False
        txtCUST_NAME.Visible = False

        'Show_Document(XLS_FILENAME_multi)
        Add_Document_to_ASTSPRF1(XLS_FILENAME_multi)
        XLS_FILENAME_multi = ""

        MsgBox("Multiple Sheet XLS has been Generated", MsgBoxStyle.OkOnly, "Success")

    End Sub

    Private Sub grdSATSIST3_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSATSIST3.InitializeRow
        If 1 = 1 Then Exit Sub
        If e.Row.IsDataRow Then
            Dim CUST_CODE As String = e.Row.Cells("CUST_CODE").Value

            Dim sqlw As String = ""
            If optRANGE.Value = "P" Then
                sqlw = " and OPS_YYYYPP >= '" & PsLY(0) & "' and OPS_YYYYPP <= '" & PsLY(Periods - 1) & "'"
            Else
                sqlw = " and OPS_YYYYWW >= '" & PsLY(0) & "' and OPS_YYYYWW <= '" & PsLY(Periods - 1) & "'"
            End If

            Dim rows() As DataRow = dst.Tables("SPTCOOPX").Select("CUST_CODE = '" & CUST_CODE & "'" & sqlw)
            For Each row As DataRow In rows
                Dim DATE_START As Date = row.Item("DATE_START")
                For P As Integer = 1 To Periods
                    If Format(DATE_START, "yyyyMMdd") > Format(PE_DATEsLY(P - 1), "yyyyMMdd") And _
                       Format(DATE_START, "yyyyMMdd") <= Format(PE_DATEsLY(P), "yyyyMMdd") Then
                        Dim C As String = "STQTY_P" & Format(P, "00")
                        Dim T As String = Format(DATE_START, "MM/dd") & " " & CUST_CODE & " " & row.Item("BOOKING_NAME") ' & " " & row.Item("VEHICLE_CODE")
                        Dim TT As String = e.Row.Cells(C).Column.Group.Header.ToolTipText & ""
                        If Not TT.Contains(T & vbCrLf) Then
                            e.Row.Cells(C).Appearance.BackColor = Drawing.Color.Violet ' .LightGreen
                            e.Row.Cells(C).ToolTipText &= T & vbCrLf
                            e.Row.Cells(C).Column.Group.Header.ToolTipText &= T & vbCrLf
                        End If
                    End If
                Next
            Next
        End If
    End Sub

    Private Sub tabSIST_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabSIST.SelectedTabChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Setup_tabSIST()
    End Sub

    Sub Setup_tabSIST()
        Dim W As Integer = 80
        If tabSIST.SelectedTab.Key = "Stores" Then
            W = 150
        End If
        grdSATSIST1.DisplayLayout.Bands(0).Columns("ITEM_CODE").Width = W
        grdSATSIST2.DisplayLayout.Bands(0).Columns("CUST_STORE_NAME").Width = W
    End Sub
    Private Sub chkSIST_CheckedChanged(sender As Object, e As EventArgs) Handles chkSIST.CheckedChanged
        UltraExplorerBar1.Groups("Sell-In / Sell-Thru Options").Visible = (chkSIST.Checked)
        chkSIST_Multi.Checked = False
        chkExtendedData.Enabled = Not chkSIST.Checked
        If chkSIST.Checked Then
            chkExtendedData.Checked = False
        End If
    End Sub

    Sub Auto_Select()

        dst.Tables("SATSISTC").Rows.Clear()

        Dim sqlp As String = ""
        If optRANGE.Value = "P" Then
            RYP0 = Absx1.cmbFor("RYP0").Value
            RYP1 = Absx1.cmbFor("RYP1").Value
            sqlp = " and ORDR_YYYYPP_UPDATED between '" & RYP0 & "' and '" & RYP1 & "'"
        Else
            RYW0 = Absx1.cmbFor("RYW0").Value
            RYW1 = Absx1.cmbFor("RYW1").Value
            sqlp = " and OPS_YYYYWW between '" & RYW0 & "' and '" & RYW1 & "'"
        End If

        Dim CODEs As String = Get_Multi_CODEs("SATSISTI")

        If CODEs <> "" Then
            ASCMAIN1.sql = "Select Distinct CUST_CODE from SOTINVH2" & vbCrLf _
                & " where ITEM_CODE in (" & CODEs & ")" & sqlp
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Add_Customer(row.Item(0))
            Next
        End If
    End Sub

    Function Get_Multi_CODEs(TABLE_NAME As String) As String

        Dim CODEs As String = ""
        For Each row As DataRow In dst.Tables(TABLE_NAME).Select("")
            CODEs &= ",'" & row.Item(0) & "'"
        Next

        Return Mid(CODEs, 2)

    End Function

    Sub Add_Items_Select()
        Dim sql_where As String = "" ' Get_List_of_Customers("ARTCUST1.CUST_CODE not in")
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("ITEM_CODE", , sql_where)

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading")

                grdSATSISTI.Visible = False
                For Each ITEM_CODE As String In ASCMAIN1.CodeSelector.SelectedCodes
                    Add_Item(ITEM_CODE)
                Next
                grdSATSISTI.Visible = True

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
            End If
        End If

    End Sub

    Sub Add_Customers_Select()
        Dim sql_where As String = "" ' Get_List_of_Customers("ARTCUST1.CUST_CODE not in")
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("CUST_CODE", , sql_where)

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading Customers")

                grdSATSISTC.Visible = False
                For Each CUST_CODE As String In ASCMAIN1.CodeSelector.SelectedCodes
                    Add_Customer(CUST_CODE)
                Next
                grdSATSISTC.Visible = True

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
            End If
        End If

    End Sub
    Private Sub grdSATSISTI_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSATSISTI.ClickCellButton
        Add_Items_Select()
    End Sub
    Private Sub grdSATSISTC_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSATSISTC.ClickCellButton
        Add_Customers_Select()
    End Sub

    Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
        If grdSATSISTI.Rows.Count = 0 And grdSATSISTC.Rows.Count = 0 Then
            MsgBox("No Items or Customers to Save", MsgBoxStyle.OkOnly, "Cannot Save Lists")
            Exit Sub
        End If

        Dim DATETIME_STAMP As Date = Now + ASCMAIN1.NowTSD

        dst.Tables("ASTLIST2").Rows.Clear()
        dst.Tables("ASTLIST1").AcceptChanges()
        dst.Tables("ASTLIST2").AcceptChanges()

        If Trim(txtLIST_DESC.Text) = "" Then
            MsgBox("You Must Provide a Description for this List", MsgBoxStyle.OkOnly, "Cannot Save Lists")
            Exit Sub
        End If
        Dim rowASTLIST1 As DataRow = Nothing
        Dim LIST_CODE As String = ""

        LIST_CODE = ASCMAIN1.Next_Control_No("ASTLIST1.LIST_CODE")
        rowASTLIST1 = dst.Tables("ASTLIST1").NewRow
        With rowASTLIST1
            .Item("LIST_CODE") = LIST_CODE
            .Item("LIST_DESC") = txtLIST_DESC.Text
            .Item("COLUMN_NAME") = "SATISLS1.CUST_CODE"
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("LIST_SHAREABLE") = "0"
            .Item("LIST_MODIFIABLE") = "0"
        End With
        dst.Tables("ASTLIST1").Rows.Add(rowASTLIST1)

        For Each row As DataRow In dst.Tables("SATSISTC").Select("")
            Dim rowASTLIST2 As DataRow = dst.Tables("ASTLIST2").NewRow
            rowASTLIST2.Item("LIST_CODE") = LIST_CODE
            rowASTLIST2.Item("CODE_VALUE") = row.Item("CUST_CODE")
            dst.Tables("ASTLIST2").Rows.Add(rowASTLIST2)
        Next


        LIST_CODE = ASCMAIN1.Next_Control_No("ASTLIST1.LIST_CODE")
        rowASTLIST1.Item("LINKED_LIST_CODE") = LIST_CODE

        rowASTLIST1 = dst.Tables("ASTLIST1").NewRow
        With rowASTLIST1
            .Item("LIST_CODE") = LIST_CODE
            .Item("LIST_DESC") = txtLIST_DESC.Text
            .Item("COLUMN_NAME") = "SATISLS1.ITEM_CODE"
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("LIST_SHAREABLE") = "0"
            .Item("LIST_MODIFIABLE") = "0"
        End With
        dst.Tables("ASTLIST1").Rows.Add(rowASTLIST1)

        For Each row As DataRow In dst.Tables("SATSISTI").Select("")
            Dim rowASTLIST2 As DataRow = dst.Tables("ASTLIST2").NewRow
            rowASTLIST2.Item("LIST_CODE") = LIST_CODE
            rowASTLIST2.Item("CODE_VALUE") = row.Item("ITEM_CODE")
            dst.Tables("ASTLIST2").Rows.Add(rowASTLIST2)
        Next

        Dim sqld As String = ""
        Update_Record_TDA("ASTLIST1", sqld)
        Update_Record_TDA("ASTLIST2", sqld)

        Sort_grdColumns(grdASTLIST1, "INIT_DATE".ToLower)

        MsgBox("Distribution List has been Saved", MsgBoxStyle.OkOnly, "Verification")
    End Sub

    Private Sub grdASTLIST1_BeforeRowsDeleted(sender As Object, e As UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdASTLIST1.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            Dim LIST_CODE As String = grow.Cells("LIST_CODE").Value
            Dim LINKED_LIST_CODE As String = grow.Cells("LINKED_LIST_CODE").Value
            ASCDATA1.ExecuteSQL("Delete from ASTLIST1 where LIST_CODE = '" & LIST_CODE & "'")
            ASCDATA1.ExecuteSQL("Delete from ASTLIST2 where LIST_CODE = '" & LIST_CODE & "'")
            ASCDATA1.ExecuteSQL("Delete from ASTLIST1 where LIST_CODE = '" & LINKED_LIST_CODE & "'")
            ASCDATA1.ExecuteSQL("Delete from ASTLIST2 where LIST_CODE = '" & LINKED_LIST_CODE & "'")
        Next
    End Sub

    Private Sub grdASTLIST1_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdASTLIST1.DoubleClickRow

        Dim LIST_CODE As String = e.Row.Cells("LIST_CODE").Value
        Dim LINKED_LIST_CODE As String = e.Row.Cells("LINKED_LIST_CODE").Value
        Dim rowASTLIST1 As DataRow = dst.Tables("ASTLIST1").Rows.Find(LIST_CODE)

        dst.Tables("SATSISTC").Rows.Clear()
        Fill_Records("ASTLIST2", LIST_CODE)
        For Each row As DataRow In dst.Tables("ASTLIST2").Select("")
            Dim CODE_VALUE As String = row.Item("CODE_VALUE")
            If grdSATSISTC.ActiveRow IsNot Nothing Then grdSATSISTC.ActiveRow.CancelUpdate()
            grdSATSISTC.ActiveRow = grdSATSISTC.DisplayLayout.Bands(0).AddNew
            grdSATSISTC.ActiveRow.Cells("CUST_CODE").Value = CODE_VALUE
            grdSATSISTC.ActiveRow.Update()
        Next
        grdSATSISTC.PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)

        dst.Tables("SATSISTI").Rows.Clear()
        Fill_Records("ASTLIST2", LINKED_LIST_CODE)
        For Each row As DataRow In dst.Tables("ASTLIST2").Select("")
            Dim CODE_VALUE As String = row.Item("CODE_VALUE")
            If grdSATSISTI.ActiveRow IsNot Nothing Then grdSATSISTI.ActiveRow.CancelUpdate()
            grdSATSISTI.ActiveRow = grdSATSISTI.DisplayLayout.Bands(0).AddNew
            grdSATSISTI.ActiveRow.Cells("ITEM_CODE").Value = CODE_VALUE
            grdSATSISTI.ActiveRow.Update()
        Next
        grdSATSISTI.PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)

    End Sub


    Sub Weekly_Sales_Heat_Map()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Summarizing Sales")

        Dim Y1 As String = Format(Val(Mid(ASCMAIN1.CYP, 1, 4)) - 4, "000")
        Dim Y5 As String = Format(Val(Y1) + 4, "0000")

        Fill_Records("SATWKSL1", Absx1.txtFor("CUST_CODE").Text)
        MAX_SALES = Val(dst.Tables("SATWKSL1").Compute("MAX(INV_SALES)", "INV_SALES> 0") & "")

        dst.Tables("ASTWCAL1").Rows.Clear()

        Dim Sql As String = ""

        Sql = "Select SUBSTR(YYYYWW,5,2) WW, SUBSTR(YYYYMM,5,2) MM" _
            & ", REL_WEEK, MAX(MAX_WEEK) MAX_WEEK"

        For Y As Integer = 1 To 5
            Sql = Sql & " , MIN(DECODE (SUBSTR(YYYYWW,1,4),'" & Format$(Val(Y1) + Y - 1, "0000") & "',WEEK_END_DATE,NULL)) Y" & Format$(Y, "0")
        Next Y

        Sql &= " From GLTPARM3" _
            & " where YYYYWW >= '" & Y1 & "01'" _
            & "   and YYYYWW <= '" & Y5 & "53'" _
            & " GROUP BY SUBSTR(YYYYWW,5,2), SUBSTR(YYYYMM,5,2), REL_WEEK" _
            & " ORDER BY SUBSTR(YYYYWW,5,2), SUBSTR(YYYYMM,5,2), REL_WEEK"

        Dim LINE_NO As Integer = 0

        For Each row As DataRow In ASCDATA1.GetDataTable(Sql).Select("", "WW,MM,REL_WEEK")
            Dim REL_WEEK As Integer = Val(row.Item("REL_WEEK") & "")
            Dim MAX_WEEK As Integer = Val(row.Item("MAX_WEEK") & "")
            Dim MM As String = row.Item("MM")
            Dim WW As String = row.Item("WW")
            Dim MONTH_X As String = Mid$(" " & Format(DateValue(MM & "/01/2000"), "MMM") & " ", REL_WEEK, 1)

            Dim rowASTWCAL1 As DataRow = dst.Tables("ASTWCAL1").NewRow
            LINE_NO += 1
            rowASTWCAL1.Item("LINE_NO") = LINE_NO
            rowASTWCAL1.Item("MONTH") = MONTH_X
            rowASTWCAL1.Item("WEEK_NO") = WW

            For Y As Integer = 1 To 5
                Dim DT1 As Date
                If row.Item("Y" & Format$(Y, "0")) & "" <> "" Then
                    Dim dt As Date = row.Item("Y" & Format$(Y, "0"))
                    DT1 = dt

                    For d As Integer = 1 To 7
                        Dim rowS As DataRow = dst.Tables("SATWKSL1").Rows.Find(dt.AddDays(-7 + d))
                        If rowS IsNot Nothing Then rowASTWCAL1.Item("DATE_" & Format(Y, "0") & Format(d, "0")) = Val(rowS.Item("INV_SALES") & "") / 1000
                        'rowASTWCAL1.Item("DATE_" & Format(Y, "0") & Format(d, "0")) = Format(dt.AddDays(-7 + d), "dd")
                    Next d


                End If
                rowASTWCAL1.Item("DATES" & Format(Y, "0")) = Format(DT1, "MM/dd")

            Next Y
            dst.Tables("ASTWCAL1").Rows.Add(rowASTWCAL1)

            If REL_WEEK = MAX_WEEK Then
                rowASTWCAL1 = dst.Tables("ASTWCAL1").NewRow
                LINE_NO += 1
                rowASTWCAL1.Item("LINE_NO") = LINE_NO
                dst.Tables("ASTWCAL1").Rows.Add(rowASTWCAL1)
            End If
            'Stop

        Next

        For i As Integer = 1 To 5   ' 5 years
            Dim G As String = "Y" & Format(i, "0")
            grdASTWCAL1.DisplayLayout.Bands("ASTWCAL1").Groups(G).Header.Caption = Format(Y1 + i - 1, "0000")
        Next i


        With grdASTWCAL1.DisplayLayout.Bands("ASTWCAL1").SortedColumns
            .Clear()
            .Add("LINE_NO", False)
        End With

        dst.Tables("SATCSLS1").Rows.Clear()

        Dim YW As String = ASCMAIN1.CYW
        For w As Integer = 12 To 1 Step -1
            YW = ASCMAIN1.Week_Calc(YW, -1)
            Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", YW)
            Dim DTE As Date = rowGLTPARM3.Item("WEEK_END_DATE")
            For i As Integer = 1 To 5   ' 5 years
                Dim YYYY As String = Format(Val(Mid(YW, 1, 4)) - (i - 1), "0000")
                Dim CODE_VALUE As String = YYYY
                Dim row As DataRow = dst.Tables("SATCSLS1").Rows.Find(New String() {CODE_VALUE})
                If row Is Nothing Then
                    row = dst.Tables("SATCSLS1").NewRow
                    row.Item("CODE_VALUE") = CODE_VALUE
                    dst.Tables("SATCSLS1").Rows.Add(row)
                End If
                Dim sqlDates As String = "INV_DATE >= #" & Format(DTE.AddDays(-6), "MM/dd/yyyy") & "# and INV_DATE <= #" & Format(DTE, "MM/dd/yyyy") & "#"
                Dim SALES As Decimal = Val(dst.Tables("SATWKSL1").Compute("SUM(INV_SALES)", sqlDates) & "")
                row.Item("P" & Format(w, "00")) = SALES
                row.Item("P00") = Val(row.Item("P00") & "") + SALES
                DTE = DTE.AddDays(-1 * 52)
            Next i
        Next

        Setup_Weekly_Charts()


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Setup_Weekly_Charts()
 
        Fill_Records("SATWKSLS", Absx1.txtFor("CUST_CODE").Text)

        For Each ROW As DataRow In dst.Tables("SATCSLSS").Rows
            ROW.Item("SALES") = 0
        Next
        For Each row As DataRow In dst.Tables("SATCSLSS").Select("")
            Dim rowTATSTATE() As DataRow = dst.Tables("TATSTATE").Select("STATE_NAME = '" & row.Item("STATE_NAME") & "'")
            If rowTATSTATE.Length > 0 Then
                Dim STATE_CODE As String = rowTATSTATE(0).Item("STATE_CODE")
                Dim rowSATWKSLS As DataRow = dst.Tables("SATWKSLS").Rows.Find(STATE_CODE)
                If rowSATWKSLS IsNot Nothing Then
                    row.Item("STATE_CODE") = rowSATWKSLS(0)
                    row.Item("SALES") = rowSATWKSLS(1)
                End If
            End If
        Next
        grdSATCSLSS.Text = "Sales by State last 365 days"
        Me.UltraChart1.Data.DataSource = StatesData()
        Me.UltraChart1.Data.DataBind()

        Periods = 12
        CreateGraph_SATCSLS1()
        CreateGraph_SATCSLS1_X()
        chtSATCSLS1.Visible = True
        chtSATCSLS1_X.Visible = True
    End Sub

    Private Sub grdASTWCAL1_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdASTWCAL1.InitializeRow

        For Y As Integer = 1 To 5
            For d As Integer = 1 To 7
                Dim C As String = "DATE_" & Format(Y, "0") & Format(d, "0")
                Dim V As Decimal = Val(e.Row.Cells(C).Value & "")
                Dim R As Integer = 0
                Dim G As Integer = 255
                Dim B As Integer = 0
                Dim PCT As Decimal = 0
                If V < 0 Then
                    PCT = 0
                Else
                    PCT = CInt(100 * V * 1000 / MAX_SALES)
                End If

                Dim GG As System.Drawing.Color = Drawing.Color.Empty
                If PCT > 75 Then
                    GG = System.Drawing.Color.FromArgb(255, 0, 192, 0)
                ElseIf PCT > 50 Then
                    GG = System.Drawing.Color.Lime
                ElseIf PCT > 25 Then
                    GG = System.Drawing.Color.FromArgb(128, 255, 128)
                ElseIf PCT > 10 Then
                    GG = System.Drawing.Color.FromArgb(255, 192, 255, 192)
                End If
                 
                'Dim GG2 As System.Drawing.Color = Drawing.Color.LightGreen
                e.Row.Cells(C).Appearance.BackColor = GG ' System.Drawing.Color.FromArgb(255, R, G, B)
            Next d
        Next Y
    End Sub

    Private Sub cmdABS_Click(sender As Object, e As EventArgs) Handles cmdABS.Click
        'xls_BELK
        'xls_LORD()
        xls_LORDOCT()
    End Sub

    Sub xls_LORDOCT()
        Dim RTL_X As String = "RTL_LOCT"
        Dim workbook As SpreadsheetGear.IWorkbook = Nothing
        Dim XLS_FOLDER As String = "C:\Users\wjz\Desktop\Interparfums\IPUS_Retail\LORDOCT"

        Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing

        If dst.Tables.Contains(RTL_X) Then
            dst.Tables(RTL_X).Rows.Clear()
        Else
            Create_TDA(dst.Tables.Add, RTL_X, "*")
        End If
        ASCDATA1.ExecuteSQL("Truncate Table " & RTL_X)



        For Each XLS_FILENAME As String In My.Computer.FileSystem.GetFiles(XLS_FOLDER)

            Dim TBL As DataTable = ASCDATA1.GetDataTable("Select TO_CHAR(WEEK_END_DATE,'YYYYMMDD') WEDATE, YYYYWW, YYYYPP from GLTPARM3", "GLTPARM3", 1)

            workbook = SpreadsheetGear.Factory.GetWorkbook(XLS_FILENAME)

            For Each worksheet In workbook.Worksheets

                ASCMAIN1.Progress(XLS_FILENAME, worksheet.Name)
                If worksheet.Name = "Total" Then
                Else


                    Dim W As String = worksheet.Cells(5, 0).Value & ""


                    Dim W1 As String = Split(W, ":")(0)
                    Dim W2 As String = Split(W, ":")(1)
                    Dim W3 As String = Split(W, ":")(2)
                    'Week: 38:16 Oct 16 - 22 Oct 16
                    Dim WEEK_NO As Integer = Val(W2) '  Val(Mid(W, 7, 2))
                    Dim DATE1 As Date = CDate(Mid(W3, 1, 9)) ' CDate(Mid(W, 10, 9))
                    Dim DATE2 As Date = CDate(Mid(W3, 13, 9)) ' CDate(Mid(W, 22, 9))

                    If Format(DATE2, "yyyyMMdd") < Format(DATE1, "yyyyMMdd") Then Stop

                    Dim rowGLTPARM3 As DataRow = TBL.Rows.Find(Format(DATE2, "yyyyMMdd"))

                    If rowGLTPARM3 Is Nothing Then Stop

                    Dim R As Integer = 8
                    If worksheet.Cells(R - 1, 0).Value & "" <> "Buyer" Then Stop
                    If worksheet.Cells(R - 0, 0).Value & "" <> "Total" Then Stop
                    R += 1

                    Do While worksheet.Cells(R, 0).Value & "" <> ""
                        Dim row As DataRow = dst.Tables(RTL_X).NewRow
                        row.Item("WEEK_NO") = WEEK_NO
                        row.Item("DATE1") = DATE1
                        row.Item("DATE2") = DATE2
                        For C As Integer = 0 To row.Table.Columns.Count - 9 - 1
                            row.Item(C) = worksheet.Cells(R, C).Value
                        Next
                        dst.Tables(RTL_X).Rows.Add(row)
                        R += 1
                    Loop
                End If
            Next
        Next

        ASCMAIN1.Progress("Now Uploading Data to Oracle", "")
        Update_Record_TDA(RTL_X)
        ASCMAIN1.Progress("", "")
        MsgBox("Done")
        'ASCDATA1.ExecuteSQL("Update RTL_B Set TYW = (Select YYYYWW from GLTPARM3 where WEEK_END_DATE >= RTL_B.WEEK_END_DATE and WEEK_END_DATE < RTL_B.WEEK_END_DATE +7)")
        'ASCDATA1.ExecuteSQL("Update RTL_B Set TYP = (Select YYYYPP from GLTPARM3 where YYYYWW = RTL_B.TYW)")
        'ASCDATA1.ExecuteSQL("Update RTL_B Set LYW = (Select YYYYWW from GLTPARM3 where WEEK_END_DATE >= RTL_B.WEEK_END_DATE - 52*7 and WEEK_END_DATE < RTL_B.WEEK_END_DATE -52*7 +7)")
        'ASCDATA1.ExecuteSQL("Update RTL_B Set LYP = (Select YYYYPP from GLTPARM3 where YYYYWW = RTL_B.LYW)")
        'ASCDATA1.ExecuteSQL("UPDATE RTL_B SET ITEM_CODE = (SELECT ITEM_CODE FROM ICTITEM1 WHERE ITEM_CODE = TRIM(RTL_B.ITEM))")
        'ASCDATA1.ExecuteSQL("UPDATE RTL_B SET ITEM_CODE = (SELECT ITEM_CODE FROM ICTITEM1 WHERE ITEM_CODE = REPLACE(TRIM(RTL_B.ITEM),'L00','')) WHERE ITEM_CODE IS NULL AND ITEM LIKE 'L00%'")
    End Sub

    Sub xls_LORD()
        Dim workbook As SpreadsheetGear.IWorkbook = Nothing
        Dim XLS_FOLDER As String = "C:\Users\wjz\Desktop\Interparfums\IPUS_Retail\LORD"

        Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing

        If dst.Tables.Contains("RTL_L") Then '
            dst.Tables("RTL_L").Rows.Clear()
        Else
            Create_TDA(dst.Tables.Add, "RTL_L", "*")
        End If
        ASCDATA1.ExecuteSQL("Truncate Table RTL_L")

        Dim WS As Integer = 0

        For Each XLS_FILENAME As String In My.Computer.FileSystem.GetFiles(XLS_FOLDER)
            WS += 1
            ASCMAIN1.Progress(XLS_FILENAME, CStr(WS))
            workbook = SpreadsheetGear.Factory.GetWorkbook(XLS_FILENAME)
            worksheet = workbook.Worksheets(0)

            Dim R As Integer = 8
            If worksheet.Cells(R - 1, 0).Value & "" <> "Store/DC" Then Stop
            Dim W As String = worksheet.Cells(5, 0).Value & ""
            'Week: 38:16 Oct 16 - 22 Oct 16
            Dim WEEK_NO As Integer = Val(Mid(W, 7, 2))
            Dim DATE1 As Date = CDate(Mid(W, 10, 9))
            Dim DATE2 As Date = CDate(Mid(W, 22, 9))

            Do While worksheet.Cells(R, 0).Value & "" <> ""
                Dim row As DataRow = dst.Tables("RTL_L").NewRow
                row.Item("WEEK_NO") = WEEK_NO
                row.Item("DATE1") = DATE1
                row.Item("DATE2") = DATE2
                For C As Integer = 0 To (12 + 16) - 1
                    row.Item(C) = worksheet.Cells(R, C).Value
                Next
                dst.Tables("RTL_L").Rows.Add(row)
                R += 1
            Loop
        Next

        ASCMAIN1.Progress("Now Uploading Data to Oracle", "")
        Update_Record_TDA("RTL_L")
        ASCMAIN1.Progress("", "")
        MsgBox("Done")
        'ASCDATA1.ExecuteSQL("Update RTL_B Set TYW = (Select YYYYWW from GLTPARM3 where WEEK_END_DATE >= RTL_B.WEEK_END_DATE and WEEK_END_DATE < RTL_B.WEEK_END_DATE +7)")
        'ASCDATA1.ExecuteSQL("Update RTL_B Set TYP = (Select YYYYPP from GLTPARM3 where YYYYWW = RTL_B.TYW)")
        'ASCDATA1.ExecuteSQL("Update RTL_B Set LYW = (Select YYYYWW from GLTPARM3 where WEEK_END_DATE >= RTL_B.WEEK_END_DATE - 52*7 and WEEK_END_DATE < RTL_B.WEEK_END_DATE -52*7 +7)")
        'ASCDATA1.ExecuteSQL("Update RTL_B Set LYP = (Select YYYYPP from GLTPARM3 where YYYYWW = RTL_B.LYW)")
        'ASCDATA1.ExecuteSQL("UPDATE RTL_B SET ITEM_CODE = (SELECT ITEM_CODE FROM ICTITEM1 WHERE ITEM_CODE = TRIM(RTL_B.ITEM))")
        'ASCDATA1.ExecuteSQL("UPDATE RTL_B SET ITEM_CODE = (SELECT ITEM_CODE FROM ICTITEM1 WHERE ITEM_CODE = REPLACE(TRIM(RTL_B.ITEM),'L00','')) WHERE ITEM_CODE IS NULL AND ITEM LIKE 'L00%'")
    End Sub

    Sub xls_BELK()
        Dim workbook As SpreadsheetGear.IWorkbook = Nothing
        Dim XLS_FILENAME As String = "Oscar Sales by Style Weekly Recap 10.19.16.xlsx"
        Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing

        workbook = SpreadsheetGear.Factory.GetWorkbook("c:\users\wjz\desktop\" & XLS_FILENAME)

        If dst.Tables.Contains("RTL_B") Then '
            dst.Tables("RTL_B").Rows.Clear()
        Else
            Create_TDA(dst.Tables.Add, "RTL_B", "*")
        End If
        ASCDATA1.ExecuteSQL("Truncate Table RTL_B")

        For Each worksheet In workbook.Worksheets

            Dim ITEMS As New List(Of String)
            Dim i As Integer = 4
            Do While worksheet.Cells(2, i).Value & "" <> ""
                Dim ITEM As String = worksheet.Cells(2, i).Value & ""
                ITEMS.Add(ITEM)

                Dim r As Integer = 6
                Do While worksheet.Cells(r, 0).Value & "" <> "" Or worksheet.Cells(r, 2).Value & "" <> ""
                    If worksheet.Cells(r, 2).Value & "" <> "" Then 'wkend 08/08/2015
                        Dim D As Date = CDate(Replace(worksheet.Cells(r, 2).Value & "", "wkend ", ""))
                        Dim row As DataRow = dst.Tables("RTL_B").NewRow
                        row.Item("RETAILER") = "BELK"
                        row.Item("WEEK_END_DATE") = D
                        row.Item("ITEM") = ITEM
                        row.Item("SLSQ_TY") = Val(worksheet.Cells(r, i + 0).Value & "")
                        row.Item("SLSQ_LY") = Val(worksheet.Cells(r, i + 1).Value & "")
                        row.Item("SLSA_TY") = Val(worksheet.Cells(r, i + 3).Value & "")
                        row.Item("SLSA_LY") = Val(worksheet.Cells(r, i + 4).Value & "")
                        dst.Tables("RTL_B").Rows.Add(row)
                    End If
                    r += 1
                Loop
                i += 6
            Loop

            Debug.Print(worksheet.Name & ":" & Join(ITEMS.ToArray, ","))
        Next

        Update_Record_TDA("RTL_B")
        ASCDATA1.ExecuteSQL("Update RTL_B Set TYW = (Select YYYYWW from GLTPARM3 where WEEK_END_DATE >= RTL_B.WEEK_END_DATE and WEEK_END_DATE < RTL_B.WEEK_END_DATE +7)")
        ASCDATA1.ExecuteSQL("Update RTL_B Set TYP = (Select YYYYPP from GLTPARM3 where YYYYWW = RTL_B.TYW)")
        ASCDATA1.ExecuteSQL("Update RTL_B Set LYW = (Select YYYYWW from GLTPARM3 where WEEK_END_DATE >= RTL_B.WEEK_END_DATE - 52*7 and WEEK_END_DATE < RTL_B.WEEK_END_DATE -52*7 +7)")
        ASCDATA1.ExecuteSQL("Update RTL_B Set LYP = (Select YYYYPP from GLTPARM3 where YYYYWW = RTL_B.LYW)")
        ASCDATA1.ExecuteSQL("UPDATE RTL_B SET ITEM_CODE = (SELECT ITEM_CODE FROM ICTITEM1 WHERE ITEM_CODE = TRIM(RTL_B.ITEM))")
        ASCDATA1.ExecuteSQL("UPDATE RTL_B SET ITEM_CODE = (SELECT ITEM_CODE FROM ICTITEM1 WHERE ITEM_CODE = REPLACE(TRIM(RTL_B.ITEM),'L00','')) WHERE ITEM_CODE IS NULL AND ITEM LIKE 'L00%'")

    End Sub
End Class