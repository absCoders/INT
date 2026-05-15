Imports Infragistics.Win.UltraWinGrid

Public Class SAFISLS1
    Dim SATISLS1 As String
    Dim SATISLS2 As String
    Dim sqlSATISLS1 As String
    Dim sqlSATISLS2 As String
    Dim sqlSOTINVHX As String
    Dim sqlSATISLS1_STORES As String
    Dim STORES As List(Of String)
    Dim STORES_XX As String
    Dim PERIODS_XX As String
    Dim summary_loaded As Boolean

    Dim sqlSPTCOOPX As String
    Dim APPR_STATUS_CODE_BackColors As New Dictionary(Of String, System.Drawing.Color)
    Dim APPR_STATUS_CODE_ForeColors As New Dictionary(Of String, System.Drawing.Color)

    Dim rowICTITEM1 As DataRow

    Dim Ps As New List(Of String)
    Dim WPs As New List(Of String) ' Ending Weeks of Ps
    Dim Ls As New List(Of String)
    Dim PE_DATEs() As Date

    Dim PsLY As New List(Of String)
    Dim WPsLY As New List(Of String) ' Ending Weeks of PsLY
    Dim LsLY As New List(Of String)
    Dim PE_DATEsLY() As Date

    Dim Ps2LY As New List(Of String)
    Dim WPs2LY As New List(Of String) ' Ending Weeks of PsLY
    Dim Ls2LY As New List(Of String)
    Dim PE_DATEs2LY() As Date

    Dim SATSIST1 As String = ""
    Dim SATSIST2 As String = ""
    Dim SATSIST3 As String = ""
    Dim sqlSATSIST2 As String

    Dim RYP0 As String
    Dim RYP1 As String
    Dim RYW0 As String
    Dim RYW1 As String
    Dim Periods As Integer
    Dim ITEM_CODE As String
    Dim ITEM_CODE_LY As String
    Dim COLLECTION_CODE As String
    Dim HC_CODE As String
    Dim US_STATES() As String
    Dim USmap As MapLayer
    Dim Stores_Max As Integer = 120
    Dim SIST_PMAX As Integer = 53

    Dim QTY_ALLO_PLAN As Int64

    Dim sqlSATSISTI As String
    Dim sqlSATSISTC As String
    Dim sqlSATSISTX As String
    Dim sqlSATSISTY As String

    Dim COLUMN_NAMEs As New ArrayList
    Dim COLUMN_CAPTIONs As New ArrayList
    Dim COLUMN_NAME_by_Lvl() As String
    Dim COLUMN_CAPTION_by_Lvl() As String
    Dim G_by_Lvl() As Integer
    Dim SCOPE() As String
    Dim QCOLS As New Dictionary(Of String, String)
    Dim LVL As Int16
    Dim XLS_FILENAME_multi As String = ""

    Dim needsEmptyWeek As Boolean = False

    Dim XLS_TITLE As String
    Dim XLS_DATE_RANGE As String
    Dim DTES(1) As Date

    Dim CODEs2EXCLUDE As String

    Dim VL_YW As New ValueList
    Dim VL_YP As New ValueList
    Dim VLY_YW As New ValueList
    Dim VLY_YP As New ValueList
    Dim LIST_CODEs_To_Delete As New Dictionary(Of String, String)
    Dim LIST_CODE_To_Update As String = ""
    Dim LINKED_LIST_CODE_To_Update As String = ""

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 12, -11) ' -11 - 12)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -60, 12, 0) '  0 - 12)

        Set_cmbYW("RYW0", ASCMAIN1.CYW, -5 * 52, 1 * 52, -13) '  -13 - 1 * 52)
        Set_cmbYW("RYW1", ASCMAIN1.CYW, -5 * 52, 1 * 52, 0) ' 0 - 1 * 52)

        With dst
            ASCMAIN1.sql = "Select ITEM_CODE CODE_VALUE, ITEM_DESC DESC_VALUE from ICTITEM1"
            Create_TDA(.Tables.Add, "SATISLS1", "**", 0, False)
            With .Tables("SATISLS1")
                .Columns.Add("SUB_CODE_VALUE1")
                .Columns.Add("SUB_CODE_VALUE2")
                .Columns.Add("SUB_CODE_VALUE3")
                .Columns.Add("SUB_CODE_VALUE4")
                .Columns.Add("SUB_CODE_VALUE5")
                .Columns.Add("RTL_PRICE", GetType(System.Decimal))
                .Columns.Add("WSL_PRICE", GetType(System.Decimal))
                For P As Integer = 0 To Stores_Max
                    .Columns.Add("P" & Format(P, "000"), GetType(System.Int32))
                Next
                .Columns.Add("PXX", GetType(System.Int32))
            End With

            ASCMAIN1.sql = "Select ITEM_CODE CODE_VALUE, '0' YEAR, 'XX' DATA_TYPE from ICTITEM1"
            Create_TDA(.Tables.Add, "SATISLS1_DTL", "**", 0, False)
            With .Tables("SATISLS1_DTL")
                For P As Integer = 0 To Stores_Max
                    .Columns.Add("P" & Format(P, "000"), GetType(System.Int32))
                Next
                .Columns.Add("PXX", GetType(System.Int32))
            End With

            .Relations.Add("SATISLS1_SATISLS1_DTL" _
                           , New DataColumn() { .Tables("SATISLS1").Columns("CODE_VALUE")} _
                           , New DataColumn() { .Tables("SATISLS1_DTL").Columns("CODE_VALUE")})

            ASCMAIN1.sql = "Select ITEM_CODE CODE_VALUE_PARENT, ITEM_CODE CODE_VALUE, ITEM_DESC DESC_VALUE from ICTITEM1"
            Create_TDA(.Tables.Add, "SATISLS2", "**", 0, False)
            With .Tables("SATISLS2")
                .Columns.Add("SUB_CODE_VALUE1")
                .Columns.Add("SUB_CODE_VALUE2")
                .Columns.Add("SUB_CODE_VALUE3")
                .Columns.Add("SUB_CODE_VALUE4")
                .Columns.Add("SUB_CODE_VALUE5")
                .Columns.Add("RTL_PRICE", GetType(System.Decimal))
                .Columns.Add("WSL_PRICE", GetType(System.Decimal))
                For P As Integer = 0 To Stores_Max
                    .Columns.Add("P" & Format(P, "000"), GetType(System.Int32))
                Next
                .Columns.Add("PXX", GetType(System.Int32))
            End With

            .Relations.Add("SATISLS1_SATISLS2" _
           , New DataColumn() { .Tables("SATISLS1").Columns("CODE_VALUE")} _
           , New DataColumn() { .Tables("SATISLS2").Columns("CODE_VALUE_PARENT")})

            ' & ", NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_NAME) || DECODE(CUST_STORE_MARK_FOR,NULL,'',' (' || CUST_STORE_MARK_FOR || ')') CUST_STORE_LOCATION" _

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

            Create_TDA(.Tables.Add, "TATSTATE", "*", 0, False)

            With .Tables.Add("SATISLSS")
                .Columns.Add("STATE_CODE")
                .Columns.Add("STATE_NAME")
                .Columns.Add("SALES", GetType(System.Int32))
            End With

            sqlSATSISTI = "Select ICTITEM1.ITEM_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_RETAIL_PRICE" & vbCrLf _
                & ", ICTITEM1.COLLECTION_CODE, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.ITEM_ALT_SORT" & vbCrLf _
                & ", SATSISTZ.YP_IN1, SATSISTZ.YP_IN2, SATSISTZ.YP_TH1, SATSISTZ.YP_TH2" & vbCrLf _
                & ", SATSISTZ.YW_IN1, SATSISTZ.YW_IN2, SATSISTZ.YW_TH1, SATSISTZ.YW_TH2" & vbCrLf _
                & ", SATSISTZ.ALLO_START, SATSISTZ.ALLO_END" & vbCrLf _
                & " from ICTITEM1,SATSISTZ where SATSISTZ.ITEM_CODE (+) = ICTITEM1.ITEM_CODE"
            ASCMAIN1.sql = sqlSATSISTI _
                & " and ICTITEM1.ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SATSISTI", "**", 0, False, "V", 1)
            With .Tables("SATSISTI")
                .Columns.Add("YX_IN1")
                .Columns.Add("YX_IN2")
                .Columns.Add("YX_TH1")
                .Columns.Add("YX_TH2")
            End With

            sqlSATSISTY = "Select SATSISTY.*" _
                & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_RETAIL_PRICE" _
                & ", ICTITEM1.COLLECTION_CODE, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.ITEM_ALT_SORT" _
                & " from ICTITEM1,SATSISTY where ICTITEM1.ITEM_CODE = SATSISTY.ITEM_CODE_COMPARE_TO"
            ASCMAIN1.sql = sqlSATSISTY _
                & " and SATSISTY.ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SATSISTY", "**", 0, True, "V")
            With .Tables("SATSISTY")
                .Columns.Add("YX_IN1")
                .Columns.Add("YX_IN2")
                .Columns.Add("YX_TH1")
                .Columns.Add("YX_TH2")
            End With

            Create_TDA(.Tables.Add, "SATSISTZ", "*")

            sqlSATSISTC = "Select ARTCUST1.CUST_CODE" _
                & ", ARTCUST1.CUST_NAME" _
                & ", ARTCUST1.SREP_CODE, ARTCUST1.TRADE_CLASS_CODE" _
                & " from ARTCUST1"
            ASCMAIN1.sql = sqlSATSISTC _
                & " where ARTCUST1.CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SATSISTC", "**", 0, False, "V")

            sqlSATSISTX = "Select ARTCUST1.CUST_CODE" _
                & ", ARTCUST1.CUST_NAME" _
                & " from ARTCUST1"
            ASCMAIN1.sql = sqlSATSISTX _
                & " where ARTCUST1.CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SATSISTX", "**", 0, False, "V")

            For Each TABLE_NAME As String In New String() {"SATSIST1", "SATSIST2", "SATSIST3"}
                'ASCMAIN1.sql = "Select CUST_CODE" & vbCrLf _
                '    & IIf(TABLE_NAME = "SATSIST2", ", CUST_STORE_NO", "") & vbCrLf _
                '    & ", CUST_NAME from " & IIf(TABLE_NAME = "SATSIST2", "ARTCUST2", "ARTCUST1")
                If TABLE_NAME = "SATSIST2" Then
                    ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO, ARTCUST2.CUST_STORE_NAME CUST_NAME from ARTCUST2, ICTITEM1"
                Else
                    ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME from ARTCUST1, ICTITEM1"
                End If
                Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, False)
                With .Tables(TABLE_NAME)
                    .Columns("CUST_NAME").MaxLength = -1
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
                .Tables(TABLE_NAME).Columns.Add("EOW", GetType(System.Int32))
            Next

            Dim tbl As DataTable = dst.Tables("SATSIST3").Copy
            tbl.TableName = "SATSIST4"
            dst.Tables.Add(TBL)
            .Tables("SATSIST4").PrimaryKey = New DataColumn() { .Tables("SATSIST4").Columns("ITEM_CODE"), .Tables("SATSIST4").Columns("CUST_CODE")}


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
                .PrimaryKey = New DataColumn() { .Columns("COLUMN_NAME")}
            End With

            ASCMAIN1.sql = "Select CTL_NO_TEXT COLUMN_NAME, CTL_NO_TEXT CODE_VALUE, CTL_NO_TEXT DESC_VALUE from TATCTLN1"
            Create_TDA(.Tables.Add, "SATANALD", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select ASTLIST1.* from ASTLIST1 where COLUMN_NAME = 'SATISLS1.ITEM_CODE' and LINKED_LIST_CODE is Not Null"
            Create_TDA(.Tables.Add, "ASTLIST1", "**", 0)

            Create_TDA(.Tables.Add, "ASTLIST2", "*", 1)

            .Tables("SATANALD").Columns("DESC_VALUE").MaxLength = -1 ' 100


            With .Tables.Add("SATSISTU")
                .Columns.Add("YEAR")
                .Columns.Add("ITEM_CODE")
                .Columns.Add("COLLECTION_CODE")
                .Columns.Add("HC_CODE")
                .Columns.Add("BRAND_CODE")
                .Columns.Add("SHEET_NO", GetType(System.Int32))
                .Columns.Add("SHEET_NAME")
                .Columns.Add("ROW_NO", GetType(System.Int32))
                .PrimaryKey = New DataColumn() { .Columns("YEAR"), .Columns("ITEM_CODE")}
            End With


            ASCMAIN1.sql = "Select SOTALLO1.* from SOTALLO1" & vbCrLf _
                & " where SOTALLO1.ITEM_CODE = :PARM1" & vbCrLf _
                & "   and SOTALLO1.DATE_START <= :PARM2 and SOTALLO1.DATE_END >= :PARM3"
            Create_TDA(.Tables.Add, "SOTALLO1", "**", 0, False, "VDD")
            ASCMAIN1.sql = "Select SOTALLO2.* from SOTALLO1,SOTALLO2" & vbCrLf _
                & " where SOTALLO1.ITEM_CODE = :PARM1" & vbCrLf _
                & "   and SOTALLO1.DATE_START <= :PARM2 and SOTALLO1.DATE_END >= :PARM3" & vbCrLf _
                & "   and SOTALLO2.ALLO_CTL_NO = SOTALLO1.ALLO_CTL_NO"
            Create_TDA(.Tables.Add, "SOTALLO2", "**", 0, False, "VDD")

        End With

        dst.Tables("SATANALC").Rows.Clear()
        With dst.Tables("SATANALC").Rows
            For Each COLUMN_NAME As String In New String() {"BRAND_CODE", "COLLECTION_CODE", "PROD_CODE"}
                Dim COLUMN_NAME2 As String = COLUMN_NAME
                Select Case COLUMN_NAME
                    Case "BRAND_CODE"
                        ASCMAIN1.sql = "Select BRAND_CODE, BRAND_NAME from ICTBRAN1"
                    Case "COLLECTION_CODE"
                        ASCMAIN1.sql = "Select COLLECTION_CODE, COLLECTION_NAME from ICTCOLL1"
                    Case "PROD_CODE"
                        ASCMAIN1.sql = "Select PROD_CODE, PROD_DESC from ICTPROD1"
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

        With grdSATISLS1.DisplayLayout
            .NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
            .NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
            .ViewStyle = UltraWinGrid.ViewStyle.MultiBand
            .MaxBandDepth = 2

            grdSATISLS1.DataSource = dst.Tables("SATISLS1")
            .Bands(1).Hidden = True
            .Bands(2).Hidden = True
        End With

        grdASTLIST1.DataSource = New DataView(dst.Tables("ASTLIST1"), "COLUMN_NAME = 'SATISLS1.ITEM_CODE'", "INIT_DATE DESC", DataViewRowState.CurrentRows)

        grdSATISLS2.DataSource = dst.Tables("SATISLS2")
        grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")
        grdSATSISTI.DataSource = dst.Tables("SATSISTI")
        grdSATSISTC.DataSource = dst.Tables("SATSISTC")
        grdSATSISTX.DataSource = dst.Tables("SATSISTX")
        grdSATSIST1.DataSource = dst.Tables("SATSIST1")
        grdSATSIST2.DataSource = dst.Tables("SATSIST2")
        grdSATSIST3.DataSource = dst.Tables("SATSIST3")

        grdSATSISTY.DataSource = dst.Tables("SATSISTY")

        grdSPTCOOPX.DataSource = dst.Tables("SPTCOOPX")

        Dim dvw As DataView = dst.Tables("SATISLSS").DefaultView
        dvw.RowFilter = "SALES <> 0"
        grdSATISLSS.DataSource = dvw

        Create_Summary(grdSOTINVHX, "INV_NO", "Count")
        Create_Summary(grdSOTINVHX, "ORDR_QTY_SHIP")
        Create_Summary(grdSOTINVHX, "ORDR_AMT_SHIP")

        Create_Summary(grdSATISLSS, "STATE_CODE", "Count")
        Create_Summary(grdSATISLSS, "SALES")

        Create_Summary(grdSPTCOOPX, "AUTH_NO", "Count")
        Create_Summary(grdSPTCOOPX, New String() {"DIST_AMT", "OPEN_AMT", "PAID_AMT"})

        Dim YW_MIN As String = ASCMAIN1.Week_Calc(ASCMAIN1.CYW, -5 * 52)
        Dim YW_MAX As String = ASCMAIN1.Week_Calc(ASCMAIN1.CYW, 1 * 52)
        ASCMAIN1.sql = $"Select YYYYWW, LEGEND from GLTPARM3 where YYYYWW BETWEEN '{YW_MIN}' and '{YW_MAX}'"
        For Each ROW As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("", "YYYYWW")
            VL_YW.ValueListItems.Add(ROW.Item("YYYYWW"), ROW.Item("LEGEND"))
            VLY_YW.ValueListItems.Add(ROW.Item("YYYYWW"), ROW.Item("LEGEND"))
        Next

        Dim YP_MIN As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -60)
        Dim YP_MAX As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 12)
        ASCMAIN1.sql = $"Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP BETWEEN '{YP_MIN}' and '{YP_MAX}'"
        For Each ROW As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("", "OPS_YYYYPP")
            VL_YP.ValueListItems.Add(ROW.Item("OPS_YYYYPP"), ROW.Item("LEGEND"))
            VLY_YP.ValueListItems.Add(ROW.Item("OPS_YYYYPP"), ROW.Item("LEGEND"))
        Next

        With grdSATSISTI.DisplayLayout.Bands("SATSISTI")
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key <> "ITEM_CODE" Then GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Violet

                If GCOL.Key = "YX_IN1" Or GCOL.Key = "YX_IN2" Or GCOL.Key = "YX_TH1" Or GCOL.Key = "YX_TH2" Or GCOL.Key = "ALLO_START" Or GCOL.Key = "ALLO_END" Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                End If

            Next

        End With

        With grdSATSISTY.DisplayLayout.Bands("SATSISTY")
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key <> "ITEM_CODE" Then GCOL.CellActivation = UltraWinGrid.Activation.NoEdit

                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Violet

                If GCOL.Key = "LY" Then
                    ' GCOL.Header.Caption = "Years Ago"
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                End If

                If GCOL.Key = "YX_IN1" Or GCOL.Key = "YX_IN2" Or GCOL.Key = "YX_TH1" Or GCOL.Key = "YX_TH2" Or GCOL.Key = "ALLO_START" Or GCOL.Key = "ALLO_END" Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                End If

                If GCOL.Key = "ALLO_START" Or GCOL.Key = "ALLO_END" Then
                    GCOL.Hidden = True ' NOT SURE WHAT WE ARE DOING ABOUT LY YET
                End If
            Next
        End With

        Setup_VL()


        With grdSATSISTC.DisplayLayout.Bands("SATSISTC")
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key <> "CUST_CODE" Then GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Violet
            Next
        End With
        With grdSATSISTX.DisplayLayout.Bands("SATSISTX")
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

        For Each gcol As UltraWinGrid.UltraGridColumn In grdSOTINVHX.DisplayLayout.Bands(0).Columns
            gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            gcol.Header.Appearance.BackColor = Drawing.Color.White
            If New String() {"ORDR_QTY_SHIP", "ORDR_UNIT_PRICE", "ORDR_AMT_SHIP"}.Contains(gcol.Key) Then
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            ElseIf New String() {"CUST_STORE_NO", "CUST_STORE_LOCATION"}.Contains(gcol.Key) Then
                gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
            Else
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
            End If

        Next

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSATISLS1, grdSATISLS2}
            Create_Summary(grd, "CODE_VALUE", "Count")
            For P As Integer = 0 To Stores_Max
                Create_Summary(grd, "P" & Format(P, "000"))
            Next
            Create_Summary(grd, "PXX")

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
                .Columns("P000").Header.Fixed = True
                .Columns("PXX").Header.Fixed = True
            End With
        Next

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSATSIST1, grdSATSIST2, grdSATSIST3}
            With grd.DisplayLayout.Bands(0)
                Dim G As UltraWinGrid.UltraGridGroup = .Groups.Add
                ' G.Header.Caption = "Customer & 1st Date Shipped"
                For Each COLUMN_NAME As String In New String() {"CUST_CODE", "CUST_NAME", "INV_DATE_SHIPPED"}
                    If COLUMN_NAME = "CUST_CODE" Then
                        If grd.Name = "grdSATSIST2" Then
                            COLUMN_NAME = "CUST_STORE_NO"
                        End If
                        Create_Summary(grd, COLUMN_NAME, "Count")
                    End If
                    With .Columns(COLUMN_NAME)
                        .Group = G
                        If grd.Name = "grdSATSIST2" Then
                            .Hidden = IIf(COLUMN_NAME = "CUST_STORE_NO", True, False) ' False
                        Else
                            .Hidden = IIf(COLUMN_NAME = "CUST_NAME", True, False) ' False
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

                .Columns("CUST_CODE").CellAppearance.BackColor = Drawing.Color.LightGray

                If grd.Name = "grdSATSIST2" Then
                    .Columns("CUST_STORE_NO").CellAppearance.BackColor = Drawing.Color.LightGray
                    .Columns("CUST_STORE_NO").Header.Caption = "Store"
                End If

                .Columns("CUST_CODE").Header.Caption = "Customer" ' "Code"

                If grd.Name = "grdSATSIST2" Then
                    .Columns("CUST_NAME").CellAppearance.BackColor = Drawing.Color.LightGray
                    .Columns("CUST_NAME").Header.Caption = "Store Name"
                Else
                    .Columns("CUST_NAME").Header.Caption = "Customer Name"
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

        With chtSATISLS1
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
        cbeColor.SelectedItem = cbeColor.Items(cbeColor.FindString(System.Enum.GetName(GetType(ColorModels), chtSATISLS1.ColorModel.ModelStyle), 0))

        'cbeColorBest.DataSource = System.Enum.GetNames(GetType(System.Drawing.Color))
        'cbeColorBest.SelectedItem = cbeColorBest.Items(cbeColor.FindString("Yellow", 0))
        'cbeColorWorst.DataSource = System.Enum.GetNames(GetType(System.Drawing.Color))
        'cbeColorWorst.SelectedItem = cbeColorBest.Items(cbeColor.FindString("Red", 0))

        Setup_Map()



        ASCMAIN1.Add_Value_List(grdSPTCOOPX, "APPR_STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'APPR_STATUS_CODE'")
        ASCMAIN1.Add_Value_List(grdSPTCOOPX, "STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'STATUS_CODE'")

        ASCMAIN1.Add_Value_List(grdSATSISTY, "LY", Nothing, New String() {":", "1:1 ago", "2:2 ago"})

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

        MakeTransparent(lblTitle)
        MakeTransparent(chkExcludeCustomers)
        MakeTransparent(chkConsolidateLYs)

        dst.Tables("SATSISTX").Rows.Clear()
        Fill_Records("ASTLIST2", Me.Name)
        For Each row As DataRow In dst.Tables("ASTLIST2").Select("")
            Dim CODE_VALUE As String = row.Item("CODE_VALUE")
            If grdSATSISTX.ActiveRow IsNot Nothing Then grdSATSISTX.ActiveRow.CancelUpdate()
            grdSATSISTX.ActiveRow = grdSATSISTX.DisplayLayout.Bands(0).AddNew
            grdSATSISTX.ActiveRow.Cells("CUST_CODE").Value = CODE_VALUE
            grdSATSISTX.ActiveRow.Update()
        Next
        grdSATSISTX.PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"
                Validate_Code("ITEM_CODE")

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
                                WPs2LY.Clear()
                                PsLY.Clear()
                                Ps2LY.Clear()
                                LsLY.Clear()
                                Ls2LY.Clear()
                                ReDim PE_DATEs(Periods)
                                ReDim PE_DATEsLY(Periods)
                                ReDim PE_DATEs2LY(Periods)
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

                                    Dim YP2LY As String = ASCMAIN1.Period_Calc(YP, -24)
                                    Ps2LY.Add(YP2LY)
                                    ASCMAIN1.sql = "Select Max (YYYYWW) from GLTPARM3 where YYYYPP = '" & YP2LY & "'"
                                    Dim WYP2LY As String = ASCDATA1.GetDataValue
                                    WPs2LY.Add(WYP2LY)
                                    Dim LEGEND_2LY As String = ASCMAIN1.Get_Legend(YP2LY)
                                    Ls2LY.Add(Mid(LEGEND_2LY, 1, 16))

                                    rowGLTPARM2 = LookUp("GLTPARM2", YP2LY)
                                    PE_DATEs2LY(P) = rowGLTPARM2.Item("PRD_END_DATE")

                                Next

                                PE_DATEs(0) = LookUp("GLTPARM2", ASCMAIN1.Period_Calc(RYP0, -1)).Item("PRD_END_DATE")
                                PE_DATEsLY(0) = LookUp("GLTPARM2", ASCMAIN1.Period_Calc(RYP0, -1 - 12)).Item("PRD_END_DATE")
                                PE_DATEs2LY(0) = LookUp("GLTPARM2", ASCMAIN1.Period_Calc(RYP0, -1 - 24)).Item("PRD_END_DATE")
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

                            If Periods < 1 Or Periods > Stores_Max Or (chkSIST.Checked And Periods > 52) Then
                                EMsg &= vbCr & "Total number of Periods must be between 1 and " & CStr(IIf(chkSIST.Checked, 52, Stores_Max))
                            Else

                                WPs.Clear()
                                Ps.Clear()
                                Ls.Clear()
                                WPsLY.Clear()
                                PsLY.Clear()
                                LsLY.Clear()
                                WPs2LY.Clear()
                                Ps2LY.Clear()
                                Ls2LY.Clear()
                                ReDim PE_DATEs(Periods)
                                ReDim PE_DATEsLY(Periods)
                                ReDim PE_DATEs2LY(Periods)
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






                                    Dim YW2LY As String = ASCMAIN1.Week_Calc(YW_orig, -52 * 2)

                                    If ASCMAIN1.CLIENT = "INT" Then
                                        YW2LY = Format(Val(Mid(YW, 1, 4)) - 2, "0000") & Mid(YW, 5, 2)
                                    End If
                                    Ps2LY.Add(YW2LY)
                                    WPs2LY.Add(YW2LY)
                                    If ASCMAIN1.CLIENT = "INT" And Mid(YW, 5, 2) = "53" Then
                                        If Not emptyWeek Then
                                            Dim LEGEND_2LY As String = ""
                                            Ls2LY.Add(LEGEND_2LY)
                                        Else
                                            rowGLTPARM3 = LookUp("GLTPARM3", YW2LY)
                                            Dim WEEK_END_DATE_2LY As Date = rowGLTPARM3.Item("WEEK_END_DATE")
                                            Dim LEGEND_2LY As String = rowGLTPARM3.Item("LEGEND") & " (" & Format(WEEK_END_DATE_2LY.AddDays(-6), "MM/dd") & "-" & Format(WEEK_END_DATE_2LY, "MM/dd") & ")"
                                            Ls2LY.Add(LEGEND_2LY)
                                        End If
                                    Else
                                        rowGLTPARM3 = LookUp("GLTPARM3", YW2LY)
                                        Dim WEEK_END_DATE_2LY As Date = rowGLTPARM3.Item("WEEK_END_DATE")
                                        Dim LEGEND_2LY As String = rowGLTPARM3.Item("LEGEND") & " (" & Format(WEEK_END_DATE_2LY.AddDays(-6), "MM/dd") & "-" & Format(WEEK_END_DATE_2LY, "MM/dd") & ")"
                                        Ls2LY.Add(LEGEND_2LY)
                                    End If

                                    PE_DATEs2LY(P) = rowGLTPARM3.Item("WEEK_END_DATE")
                                    If emptyWeek Then
                                        emptyWeek = False
                                    End If
                                Next

                                PE_DATEs(0) = PE_DATEs(1).AddDays(-7)
                                PE_DATEsLY(0) = PE_DATEsLY(1).AddDays(-7)
                                PE_DATEs2LY(0) = PE_DATEs2LY(1).AddDays(-7)
                            End If
                        End If
                    End If

                    If Periods < 1 Or Periods > Stores_Max Or (chkSIST.Checked And Periods > 52) Then
                        EMsg &= vbCr & "Total number of Periods must be between 1 and " & CStr(IIf(chkSIST.Checked, 52, Stores_Max))
                    End If

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

            Case "View"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Print Report"
                Print_Report()

            Case "Multi-Band Export"

                dst.Tables("SATISLS2").Rows.Clear()
                dst.Tables("SOTINVHX").Rows.Clear()

                Dim DATA_TYPE As String = optType1.Value & optType2.Value
                Dim TOPN As Integer = Val(numN.Value & "")
                Dim TOPX As Integer = 0
                For Each grow As UltraWinGrid.UltraGridRow In grdSATISLS1.Rows
                    TOPX += 1
                    If TOPN = 0 Or (TOPX <= TOPN Or TOPX > (grdSATISLS1.Rows.Count - TOPN)) Then
                        Dim CODE_VALUE_PARENT As String = grow.Cells("CODE_VALUE").Value
                        Load_SATISLS2(DATA_TYPE, CODE_VALUE_PARENT, False, False)
                    Else
                        grow.Hidden = True
                    End If
                Next

                grdSATISLS1.DisplayLayout.Bands("SATISLS1_SATISLS1_DTL").Hidden = True
                grdSATISLS1.DisplayLayout.Bands("SATISLS1_SATISLS2").Hidden = False

                Dim b As UltraWinGrid.UltraGridBand = grdSATISLS1.DisplayLayout.Bands("SATISLS1_SATISLS2")
                With grdSATISLS2.DisplayLayout.Bands(0)
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
                    Export_to_Excel(grdSATISLS1)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error with Multi-Band Export")
                End Try

                For Each grow As UltraWinGrid.UltraGridRow In grdSATISLS1.Rows
                    grow.Hidden = False
                Next

                grdSATISLS1.DisplayLayout.Bands("SATISLS1_SATISLS2").Hidden = True
                grdSATISLS1.DisplayLayout.Bands("SATISLS1_SATISLS1_DTL").Hidden = Not chkExtendedData.Checked
                dst.Tables("SATISLS2").Rows.Clear()

                Setup_grdSATISLS2()

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
                    .Items("Multi-Band Export").Visible = tf And Not chkSIST.Checked
                    .Items("Export XLS").Visible = tf And chkSIST.Checked
                End With
                .Groups("Data Options").Visible = tf And Not chkSIST.Checked
                .Groups("Options").Visible = Not tf
                .Groups("Item Image").Visible = tf

                .Groups("Charts").Visible = False
                .Groups("Sell-In / Sell-Thru Options").Visible = Not tf And chkSIST.Checked
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = tf And Not (chkSIST.Checked)
        ' splSATSIST0.Visible = Not tf
        splMulti.Visible = Not tf
        splSATSIST1.Visible = tf And chkSIST.Checked

        lblN.Visible = tf
        numN.Visible = tf

        cmdS0.Visible = Not ScreenMode
        cmdS1.Visible = Not ScreenMode
        cmdS2.Visible = Not ScreenMode

        lblInfo.Visible = ScreenMode
        lblInfo2.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SATISLS1", "SATISLS1_DTL", "SATISLS2", "SOTINVHX", "SATISLSS", "SATSIST1", "SATSIST2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)


        If SATSIST1 <> "" Then
            ' each query may generate a different number of columns
            If SATSIST1 <> "" Then ASCDATA1.ExecuteSQL("Drop Table " & SATSIST1)
            If SATSIST2 <> "" Then ASCDATA1.ExecuteSQL("Drop Table " & SATSIST2)
            If SATSIST3 <> "" Then ASCDATA1.ExecuteSQL("Drop Table " & SATSIST3)
            SATSIST1 = ""
            SATSIST2 = ""
            SATSIST3 = ""
        End If

        optXP.Value = "P"
        Absx1.txtFor("ITEM_CODE").Text = ""
        tabDetails.SelectedTab = tabDetails.Tabs("Details")

        Setup_SM()
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Sales Data")
        Save_Header_Fields(UltraGroupBox1)

        RYP0 = Absx1.cmbFor("RYP0").Value
        RYP1 = Absx1.cmbFor("RYP1").Value
        RYW0 = Absx1.cmbFor("RYW0").Value
        RYW1 = Absx1.cmbFor("RYW1").Value

        Dim RYX0 As String = ""
        Dim RYX1 As String = ""

        Dim rowSATSISTI As DataRow = dst.Tables("SATSISTI").Rows.Find(ITEM_CODE)
        If rowSATSISTI IsNot Nothing Then
            'RYX0 = rowSATSISTI.Item("YX_TH1") & ""
            'RYX1 = rowSATSISTI.Item("YX_TH1") & ""
            RYX0 = rowSATSISTI.Item("YX_IN1") & ""
            RYX1 = rowSATSISTI.Item("YX_IN2") & ""
        End If


        If optRANGE.Value = "P" Then

            If RYX0 = "" Then RYX0 = RYP0
            If RYX1 = "" Then RYX1 = RYP1

            Dim DTES0() As Date = ASCMAIN1.Get_Dates(RYX0)
            DTES(0) = DTES0(1)
            Dim DTES1() As Date = ASCMAIN1.Get_Dates(RYX1)
            DTES(1) = DTES1(DTES1.Length - 1)
        Else

            If RYX0 = "" Then RYX0 = RYW0
            If RYX1 = "" Then RYX1 = RYW1

            Dim rowGLTPARM3 As DataRow = Nothing
            rowGLTPARM3 = LookUp("GLTPARM3", RYX0)
            DTES(0) = CDate(rowGLTPARM3.Item("WEEK_END_DATE")).AddDays(-6)
            rowGLTPARM3 = LookUp("GLTPARM3", RYX1)
            DTES(1) = CDate(rowGLTPARM3.Item("WEEK_END_DATE")) ' .AddDays(-6)
        End If

        Dim DTE0 As Date = DTES(0)
        Dim DTE1 As Date = DTES(1)

        If chkSIST.Checked And chkSIST_Multi.Checked Then
            If rowSATSISTI.Item("ALLO_START") & "" <> "" Then
                DTE0 = rowSATSISTI.Item("ALLO_START")
            End If
            If rowSATSISTI.Item("ALLO_END") & "" <> "" Then
                DTE1 = rowSATSISTI.Item("ALLO_END")
            End If
        End If

        Fill_Records("SOTALLO1", New Object() {ITEM_CODE, DTE1, DTE0})
        Fill_Records("SOTALLO2", New Object() {ITEM_CODE, DTE1, DTE0})


        ITEM_CODE = HFs("ITEM_CODE")
        rowICTITEM1 = LookUp("ICTITEM1", ITEM_CODE)
        COLLECTION_CODE = rowICTITEM1.Item("COLLECTION_CODE") & ""
        Dim rowICTCOLL1 As DataRow = LookUp("ICTCOLL1", COLLECTION_CODE)
        HC_CODE = rowICTCOLL1.Item("HC_CODE") & ""


        ' NOTE THAT THERE ARE MULTIPLE LY AND 2LY ITEMS NOW FOR EACH ITEM IN SATSISTY
        ' THE DATA ACCUMULATED IN THE SECTION BELOW MAY NOT BE VALID ANY LONGER
        ' ITEM_CODE_LY SHOULD ONLY DERIVE FROM A RECORD IN SATSISTY
        ' SEE COMMENT BELOW FOR SATSISTY

        ITEM_CODE_LY = rowICTITEM1.Item("ITEM_CODE_COMPARE_TO") & ""
        If ITEM_CODE_LY = "" Then
            ITEM_CODE_LY = ITEM_CODE
        End If

        Dim DTE1_LY As Date = PE_DATEs(PE_DATEs.Count - 1)
        Dim DTE0_LY As Date = PE_DATEs(0).AddDays(1)

        'ASCMAIN1.sql = "Select * from SOTALLO1" & vbCrLf _
        '    & " where ITEM_CODE = '" & ITEM_CODE & "'" & vbCrLf _
        '    & "   and DATE_START <= '" & Format(PE_DATEs(0).AddDays(1), "dd-MMM-yyyy") & "' and DATE_END >= '" & Format(PE_DATEs(0), "dd-MMM-yyyy") & "'"
        ASCMAIN1.sql = "Select * from SOTALLO1" & vbCrLf _
            & " where ITEM_CODE = '" & ITEM_CODE & "'" & vbCrLf _
            & "   and DATE_START <= '" & Format(DTE1_LY, "dd-MMM-yyyy") & "' and DATE_END >= '" & Format(DTE0_LY, "dd-MMM-yyyy") & "'"

        QTY_ALLO_PLAN = 0
        Dim tblAllo As DataTable = ASCDATA1.GetDataTable
        'If tblAllo.Rows.Count <> 0 Then
        '    If tblAllo.Rows(0).Item("ITEM_CODE_COMPARE_TO") & "" <> "" Then
        '        ITEM_CODE_LY = tblAllo.Rows(0).Item("ITEM_CODE_COMPARE_TO")
        '    End If
        '    QTY_ALLO_PLAN = Val(tblAllo.Rows(0).Item("QTY_ALLO_PLAN") & "")
        'End If
        If tblAllo.Rows.Count <> 0 Then
            For Each rowAllo As DataRow In tblAllo.Select("", "DATE_START")
                If rowAllo.Item("ITEM_CODE_COMPARE_TO") & "" <> "" Then
                    ITEM_CODE_LY = rowAllo.Item("ITEM_CODE_COMPARE_TO")
                End If
                QTY_ALLO_PLAN += Val(rowAllo.Item("QTY_ALLO_PLAN") & "")
            Next
        End If


        ' SATSISTY

        If chkSIST_Multi.Checked Then
        Else
            Fill_Records("SATSISTY", ITEM_CODE)
        End If

        Dim rowSATSISTYs() As DataRow = dst.Tables("SATSISTY").Select($"ITEM_CODE = '{ITEM_CODE}' and LY = '1'")
        If rowSATSISTYs.Length >= 1 Then
            ITEM_CODE_LY = rowSATSISTYs(0).Item("ITEM_CODE_COMPARE_TO")
        End If

        QTY_ALLO_PLAN = Val(dst.Tables("SOTALLO1").Compute("SUM(QTY_ALLO_PLAN)", "") & "")
        'QTY_ALLO_PLAN = Val(dst.Tables("SOTALLO2").Compute("SUM(QTY_ALLO)", "") & "")

        lblInfo.Text = "Collection: " & COLLECTION_CODE

        ' WHAT ABOUT LAUNCH_DATE? =
        Dim ITEM_DATE_TO_SHIP As String = ""
        If rowICTITEM1.Item("ITEM_DATE_TO_SHIP") & "" <> "" Then
            ITEM_DATE_TO_SHIP = Format(rowICTITEM1.Item("ITEM_DATE_TO_SHIP"), "MM/dd/yyyy")
        End If
        Dim ITEM_RETAIL_PRICE As Decimal = Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & "")

        lblInfo2.Text = "Retail: " & Format(ITEM_RETAIL_PRICE, "$#,##0.00") & IIf(ITEM_DATE_TO_SHIP = "", "", "; Ship: " & ITEM_DATE_TO_SHIP)
        If QTY_ALLO_PLAN <> 0 Then lblInfo2.Text &= vbCrLf & "Plan: " & Format(QTY_ALLO_PLAN, "#,##0") & " units, " & Format(QTY_ALLO_PLAN * ITEM_RETAIL_PRICE, "$#,##0")


        Dim IMAGE_NAME As String = ITEM_CODE
        Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
        Dim imgba() As Byte = Nothing
        picItemImage.Image = ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, False, , , imgba)
        UltraExplorerBar1.Groups("Item Image").Text = "Item Image " & ITEM_CODE

        If chkSIST.Checked Then
            Load_SIST()

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
            ASCMAIN1.sql = sqlSPTCOOPX & sqlw & " and ICTCOLL1.HC_CODE = '" & HC_CODE & "'"
            Fill_Records("SPTCOOPX", "", True, ASCMAIN1.sql)

            If optRANGE.Value = "P" Then
                sqlw = " and SPTCOOP1.OPS_YYYYPP >= '" & PsLY(0) & "' and SPTCOOP1.OPS_YYYYPP <= '" & PsLY(Periods - 1) & "'"
            Else
                sqlw = " and SPTCOOP1.OPS_YYYYWW >= '" & PsLY(0) & "' and SPTCOOP1.OPS_YYYYWW <= '" & PsLY(Periods - 1) & "'"
            End If
            ASCMAIN1.sql = sqlSPTCOOPX & sqlw & " and ICTCOLL1.HC_CODE = '" & HC_CODE & "'"
            Fill_Records("SPTCOOPX", "", False, ASCMAIN1.sql)

            dst.Tables("SATSIST1").Columns("OHQTY_P00").Expression = "ISNULL(OHQTY_P" & Format(Periods, "00") & ",0)"
            dst.Tables("SATSIST2").Columns("OHQTY_P00").Expression = "ISNULL(OHQTY_P" & Format(Periods, "00") & ",0)"
            dst.Tables("SATSIST3").Columns("OHQTY_P00").Expression = "ISNULL(OHQTY_P" & Format(Periods, "00") & ",0)"

            Dim dvw As DataView = DirectCast(grdSPTCOOPX.DataSource, DataTable).DefaultView
            dvw.RowFilter = sqlf

            Sort_grdColumns(grdSPTCOOPX, "AUTH_NO".ToLower)
            grdSPTCOOPX.Text = "Sales Promotions for HC " & HC_CODE & " Starting between " & SYP(0) & " and " & SYP(1)
        Else
            Create_SATISLS1()
            optXP.Items(0).DisplayText = optRANGE.Items(optRANGE.CheckedIndex).DisplayText
            Load_Data()
        End If

        ASCMAIN1.Progress("")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSATISLS1, "SSSSSSSBBB", "Show Filter", "Show GroupBox", "Show SUB_CODE_VALUE1", "Show SUB_CODE_VALUE2", "Show SUB_CODE_VALUE3", "Show SUB_CODE_VALUE4", "Show SUB_CODE_VALUE5", "Item Status Inquiry", "Customer Order Inquiry", "Customer Sales Summary")
        Load_Popup_Menu(grdSATISLS2, "SSSSSSSBBB", "Show Filter", "Show GroupBox", "Show SUB_CODE_VALUE1", "Show SUB_CODE_VALUE2", "Show SUB_CODE_VALUE3", "Show SUB_CODE_VALUE4", "Show SUB_CODE_VALUE5")
        Load_Popup_Menu(grdSOTINVHX, "SSBBBB", "Show Filter", "Show GroupBox", "Item Status Inquiry", "Customer Order Inquiry", "Customer Sales Summary", "Sales Order Inquiry")
        Load_Popup_Menu(grdSATISLSS, "CC", "Best", "Worst")
        Load_Popup_Menu(grdSATSIST1, "SBBB", "Show Filter", "Customer Order Inquiry", "Customer Sales Summary", "Export to Excel")
        Load_Popup_Menu(grdSPTCOOPX, "SB", "Show Filter", "Sales Promotion Inquiry")
        Load_Popup_Menu(grdSATSISTI, "BB", "Clear All", "Add Items")
        Load_Popup_Menu(grdSATSISTC, "BB", "Clear All", "Add Customers")
        Load_Popup_Menu(grdSATSISTX, "BB", "Clear All", "Add Customers")
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
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else

            Select Case e.SourceControl.Name
                Case "grdSATISLS1"

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
                Add_Items_Select()

            Case "Add Customers"
                Add_Customers_Select()

            Case "Show SUB_CODE_VALUE1", "Show SUB_CODE_VALUE2", "Show SUB_CODE_VALUE3", "Show SUB_CODE_VALUE4", "Show SUB_CODE_VALUE5" ', "Show RTL_PRICE", "Show WSL_PRICE"

                Dim COLUMN_NAME As String = Mid(e.Tool.Key, 6)
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = Not tlb_sbt.Checked


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = Absx1.txtFor("ITEM_CODE").Text
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If
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

            Case "Customer Sales Summary"

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
                    Context_Launch("View", CUST_CODE, e.Tool.Key, "SOFCSLS1")
                End If

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
                'grdSATISLSS.DataBind()
                Application.DoEvents()
                grdSATISLSS.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)

            Case "Worst"
                Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
                = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
                Me.UltraChart1.ColorModel.ColorBegin = tlb_cpt.SelectedColor
                UltraChart1.DataBind()
                'grdSATISLSS.DataBind()
                Application.DoEvents()
                grdSATISLSS.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "ITEM_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    '  Click_Command("View", e)
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

    Sub Create_SATISLS1()

        If SATISLS2 = "" Then
            SATISLS1 = ASCMAIN1.Temp_Table("Select ITEM_CODE from ICTITEM1 where ROWNUM < 1")
            SATISLS2 = ASCMAIN1.Temp_Table("Select ITEM_CODE from ICTITEM1 where ROWNUM < 1")
        End If
        ASCDATA1.ExecuteSQL("Drop Table " & SATISLS1)
        ASCDATA1.ExecuteSQL("Drop Table " & SATISLS2)

        Dim PX As String = IIf(optRANGE.Value = "P", "OPS_YYYYPP", "OPS_YYYYWW")
        Dim PX0 As String = IIf(optRANGE.Value = "P", RYP0, RYW0)
        Dim PX1 As String = IIf(optRANGE.Value = "P", RYP1, RYW1)

        sqlSATISLS1 = ""
        sqlSATISLS2 = ""
        Dim SQL() As String = New String() {"", ""}
        Dim P As Integer
        PERIODS_XX = ""

        'SQL = ""
        For P = 1 To Periods
            Dim PXP As String = IIf(optRANGE.Value = "P", ASCMAIN1.Period_Calc(RYP0, P - 1), ASCMAIN1.Week_Calc(RYW0, P - 1))
            SQL(0) &= ", Sum (Decode(" & PX & ",'" & PXP & "',RSTRETL1.QTY_SOLD,0)) P" & Format(P, "000") & vbCrLf
            SQL(1) &= ", Sum (Decode(" & PX & ",'" & ASCMAIN1.Period_Calc(PXP, -12) & "',RSTRETL1.QTY_SOLD,0)) P" & Format(P, "000") & vbCrLf
            sqlSATISLS1 &= ", P" & Format(P, "000")
            sqlSATISLS2 &= ", SUM (P" & Format(P, "000") & ") P" & Format(P, "000")
            PERIODS_XX &= "+P" & Format(P, "000")
        Next

        Dim YEAR_max As Int32 = 0
        If chkPriorYear.Checked Then
            YEAR_max = 1
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
            & ", RSTRETL1.ITEM_CODE, RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO" & vbCrLf _
            & SQL(YEAR) & vbCrLf _
            & " from RSTRETL1,GLTPARM3  " & vbCrLf _
            & " where RSTRETL1.ITEM_CODE = '" & ITEM_CODE & "'" & vbCrLf _
            & " and GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW" & vbCrLf _
            & " and RSTRETL1." & PX & " Between '" & PX0 & "' and '" & PX1 & "'" & vbCrLf _
            & " group by RSTRETL1.ITEM_CODE, RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO"

            Dim SQL_ORIG As String = sqla

            If YEAR = 0 Then
                ASCDATA1.ExecuteSQL("Create Table " & SATISLS2 & " as " & sqla)
                ASCDATA1.ExecuteSQL("Alter Table " & SATISLS2 _
                & " Add Primary Key (DATA_TYPE, YEAR, ITEM_CODE, CUST_CODE, CUST_STORE_NO)")
            Else
                ASCDATA1.ExecuteSQL("Insert into " & SATISLS2 & " " & sqla)
            End If
            sqla = Replace(sqla, "'TU' DATA_TYPE", "'TD' DATA_TYPE")
            sqla = Replace(sqla, "QTY_SOLD", "AMT_SOLD")
            ASCDATA1.ExecuteSQL("Insert into " & SATISLS2 & " " & sqla)

            sqla = Replace(SQL_ORIG, "'TU' DATA_TYPE", "'HU' DATA_TYPE")
            sqla = Replace(sqla, " and GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW", " and GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW" & vbCrLf & IIf(optRANGE.Value = "P", " and GLTPARM3.REL_WEEK = GLTPARM3.MAX_WEEK", ""))
            sqla = Replace(sqla, "QTY_SOLD", "QTY_EOW")
            ASCDATA1.ExecuteSQL("Insert into " & SATISLS2 & " " & sqla)

            sqla = Replace(SQL_ORIG, "'TU' DATA_TYPE", "'HD' DATA_TYPE")
            sqla = Replace(sqla, " and GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW", " and GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW" & vbCrLf & IIf(optRANGE.Value = "P", " and GLTPARM3.REL_WEEK = GLTPARM3.MAX_WEEK", ""))
            sqla = Replace(sqla, "from RSTRETL1", "from RSTRETL1,ICTITEM1")
            sqla = Replace(sqla, "group by", " and ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE group by")
            sqla = Replace(sqla, "QTY_SOLD", "QTY_EOW * ICTITEM1.ITEM_RETAIL_PRICE") ' SB USING ICTRETLA
            ASCDATA1.ExecuteSQL("Insert into " & SATISLS2 & " " & sqla)

            PX = IIf(optRANGE.Value = "P", "ORDR_YYYYPP_UPDATED", "OPS_YYYYWW")
            sqla = ""
            For P = 1 To Periods
                Dim PXP As String = IIf(optRANGE.Value = "P",
                                       ASCMAIN1.Period_Calc(RYP0, P - 1 - 12 * YEAR),
                                       ASCMAIN1.Week_Calc(RYW0, P - 1 - 52 * YEAR))
                sqla &= ", Sum (Decode(" & PX & ",'" & PXP & "',SOTINVH2.ORDR_QTY_SHIP,0)) P" & Format(P, "000") & vbCrLf
            Next
            sqla = "Select DECODE(INV_TYPE,'I','S','C','R') || 'U' DATA_TYPE" & vbCrLf _
            & ", '" & CStr(YEAR) & "' YEAR" & vbCrLf _
            & ", ITEM_CODE, CUST_CODE, CUST_STORE_NO" & vbCrLf _
            & sqla & vbCrLf _
            & " from SOTINVH2 " & vbCrLf _
            & " where ITEM_CODE = '" & ITEM_CODE & "'" & vbCrLf _
            & " and " & PX & " Between '" & PX0 & "' and '" & PX1 & "'" & vbCrLf _
            & " group by DECODE(INV_TYPE,'I','S','C','R') || 'U', ITEM_CODE, CUST_CODE, CUST_STORE_NO"

            ASCDATA1.ExecuteSQL("Insert into " & SATISLS2 & " " & sqla)
            sqla = Replace(sqla, "'U' DATA_TYPE", "'D' DATA_TYPE")
            sqla = Replace(sqla, " || 'U',", " || 'D',")
            sqla = Replace(sqla, "ORDR_QTY_SHIP", "ORDR_QTY_SHIP * ORDR_UNIT_PRICE")
            ASCDATA1.ExecuteSQL("Insert into " & SATISLS2 & " " & sqla)

            sqla = "Select 'I' SI, DATA_TYPE, '" & CStr(YEAR) & "' YEAR" _
            & ", ITEM_CODE, CUST_CODE CODE_VALUE" _
            & sqlSATISLS2 & " from " & SATISLS2 _
            & " group by DATA_TYPE, ITEM_CODE, CUST_CODE"
            If YEAR = 0 Then
                ASCDATA1.ExecuteSQL("Create Table " & SATISLS1 & " as " & sqla)
                ASCDATA1.ExecuteSQL("Alter Table " & SATISLS1 & " Add Primary Key (SI, DATA_TYPE, YEAR, ITEM_CODE, CODE_VALUE)")
            Else
                ASCDATA1.ExecuteSQL("Insert into " & SATISLS1 & " " & sqla)
            End If

            sqla = "Select 'S' SI, DATA_TYPE, '" & CStr(YEAR) & "' YEAR" _
            & ", ITEM_CODE, CUST_CODE CODE_VALUE" & sqlSATISLS2 _
            & " from " & SATISLS2 & " group by DATA_TYPE, ITEM_CODE, CUST_CODE"
            ASCDATA1.ExecuteSQL("Insert into " & SATISLS1 & " " & sqla)
        Next YEAR

        ASCDATA1.ExecuteSQL("Alter Table " & SATISLS1 & " Add PXX NUMBER (10,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SATISLS2 & " Add PXX NUMBER (10,0)")

        ASCMAIN1.sql = "Update " & SATISLS1 & " X SET PXX = (Select P" & Format(Periods, "000") & " from " & SATISLS1 & " " _
        & " where SI = X.SI and DATA_TYPE = 'H' || SUBSTR(X.DATA_TYPE,2,1)" _
        & "   and YEAR = X.YEAR and ITEM_CODE = X.ITEM_CODE and CODE_VALUE = X.CODE_VALUE)" _
        & " where DATA_TYPE IN ('TU','TD')"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update " & SATISLS2 & " X SET PXX = (Select P" & Format(Periods, "000") & " from " & SATISLS2 & " " _
        & " where DATA_TYPE = 'H' || SUBSTR(X.DATA_TYPE,2,1)" _
        & "   and YEAR = X.YEAR and ITEM_CODE = X.ITEM_CODE and CUST_STORE_NO = X.CUST_STORE_NO AND CUST_CODE = X.CUST_CODE)" _
        & " where DATA_TYPE IN ('TU','TD')"
        ASCDATA1.ExecuteSQL()

        STORES = New List(Of String)
        STORES_XX = ""
        Dim SQLX As String = ""

        SQLX = "Select Distinct CUST_STORE_NO from " & SATISLS2 & " order by CUST_STORE_NO"
        For Each row As DataRow In ASCDATA1.GetDataTable(SQLX).Rows
            Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
            STORES.Add(CUST_STORE_NO)
        Next

        SQLX = ""
        For S As Integer = 1 To STORES.Count
            SQLX &= ", Sum (Decode(CUST_STORE_NO,'" & STORES(S - 1) & "'," & Mid(PERIODS_XX, 2) & ",0)) P" & Format(S, "000") & vbCrLf
            STORES_XX &= "+P" & Format(S, "000")
        Next
        sqlSATISLS1_STORES = SQLX

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

        Setup_YX("")

        Setup_VL()
    End Sub

    Sub Setup_YX(ITEM_CODE As String)

        Dim sqlw As String = ""
        If ITEM_CODE <> "" Then
            sqlw = $"ITEM_CODE = '{ITEM_CODE}"
        End If

        For Each rowSATSISTI As DataRow In dst.Tables("SATSISTI").Select(sqlw)
            For Each C As String In New String() {"YX_IN1", "YX_IN2", "YX_TH1", "YX_TH2"}
                With rowSATSISTI
                    .Item(C) = .Item(Replace(C, "YX", "Y" & optRANGE.Value))
                End With
            Next
        Next

        For Each rowSATSISTY As DataRow In dst.Tables("SATSISTY").Select(sqlw)
            For Each C As String In New String() {"YX_IN1", "YX_IN2", "YX_TH1", "YX_TH2"}
                With rowSATSISTY
                    .Item(C) = .Item(Replace(C, "YX", "Y" & optRANGE.Value))
                End With
            Next
        Next
    End Sub

    Sub Setup_VL()

        Dim VL As ValueList
        If (optRANGE.Value = "P") Then
            VL = VL_YP
        Else
            VL = VL_YW
        End If

        With grdSATSISTI.DisplayLayout.Bands("SATSISTI")
            .Columns("YX_IN1").ValueList = VL
            .Columns("YX_IN2").ValueList = VL
            .Columns("YX_TH1").ValueList = VL
            .Columns("YX_TH2").ValueList = VL
        End With

        Dim VLY As ValueList
        If (optRANGE.Value = "P") Then
            VLY = VLY_YP
        Else
            VLY = VLY_YW
        End If

        With grdSATSISTY.DisplayLayout.Bands("SATSISTY")
            .Columns("YX_IN1").ValueList = VLY
            .Columns("YX_IN2").ValueList = VLY
            .Columns("YX_TH1").ValueList = VLY
            .Columns("YX_TH2").ValueList = VLY
        End With
    End Sub
    Private Sub optSI_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If SELECTION_NO = 0 Then Exit Sub
        Load_Data()
    End Sub

    Sub Setup_grd()

        Dim CAPTION As String = optType1.Text & " (" & optType2.Text & ") for " & ITEM_CODE
        CAPTION &= ", by Customer"
        grdSATISLS1.Text = CAPTION

        Dim g1 As UltraWinGrid.UltraGrid = grdSATISLS1
        Dim g2 As UltraWinGrid.UltraGrid = grdSATISLS2

        For Each G As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSATISLS1, grdSATISLS2}
            With G.DisplayLayout.Bands(0)
                If G.Name = "grdSATSLSC2" Then
                    .Columns("CODE_VALUE_PARENT").Hidden = True
                End If

                For Each COLUMN_NAME As String In New String() _
                    {"CODE_VALUE", "DESC_VALUE",
                     "SUB_CODE_VALUE1", "SUB_CODE_VALUE2", "SUB_CODE_VALUE3", "SUB_CODE_VALUE4", "SUB_CODE_VALUE5",
                     "RTL_PRICE", "WSL_PRICE"}
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                    .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    If New String() {"CODE_VALUE", "DESC_VALUE"}.Contains(COLUMN_NAME) Then
                        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Gold
                        .Columns(COLUMN_NAME).Hidden = False
                    ElseIf New String() {"SUB_CODE_VALUE1", "SUB_CODE_VALUE2", "SUB_CODE_VALUE3", "SUB_CODE_VALUE4", "SUB_CODE_VALUE5"}.Contains(COLUMN_NAME) Then
                        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
                        .Columns(COLUMN_NAME).Width = 50
                        .Columns("SUB_CODE_VALUE5").Hidden = (COLUMN_NAME.EndsWith("3") Or COLUMN_NAME.EndsWith("4") Or COLUMN_NAME.EndsWith("5"))
                    ElseIf New String() {"RTL_PRICE", "WSL_PRICE"}.Contains(COLUMN_NAME) Then
                        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Lime
                        .Columns(COLUMN_NAME).Width = 65
                        .Columns(COLUMN_NAME).Hidden = True
                    End If
                Next

                .Columns("CODE_VALUE").Header.Caption = IIf(G.Name = grdSATISLS1.Name, "Customer", "Store")
                .Columns("DESC_VALUE").Header.Caption = IIf(G.Name = grdSATISLS1.Name, "Name", "Store Name")
                .Columns("SUB_CODE_VALUE1").Header.Caption = "Rep"
                .Columns("SUB_CODE_VALUE2").Header.Caption = "State"
                .Columns("SUB_CODE_VALUE3").Header.Caption = IIf(G.Name = grdSATISLS1.Name, "Class", "Group")
                .Columns("SUB_CODE_VALUE4").Header.Caption = "City"
                .Columns("SUB_CODE_VALUE5").Header.Caption = "Zip"
                .Columns("RTL_PRICE").Header.Caption = "Retail"
                .Columns("WSL_PRICE").Header.Caption = "WhSale"

                .Columns("CODE_VALUE").Width = 80
                .Columns("DESC_VALUE").Width = 140
            End With
        Next

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
                        COLUMN_NAME = "P" & Format(P, "000")

                        .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                        .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Gold

                        If optXP.Value = "S" And G.Name = grdSATISLS1.Name Then
                            .Columns(COLUMN_NAME).Hidden = (P > STORES.Count)
                            If P <= STORES.Count Then
                                Dim LEGEND As String
                                If P = 0 Then
                                    LEGEND = "Total"
                                    .Columns(COLUMN_NAME).Width = 80
                                    .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightGreen
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
                                    .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightGreen
                                Else
                                    .Columns(COLUMN_NAME).Width = 70
                                    If optRANGE.Value = "P" Then
                                        LEGEND = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(RYP0, P - 1))
                                        LEGEND = Mid(LEGEND, 10, 6)
                                        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                                        .Columns(COLUMN_NAME).CellAppearance.BackColor2 = Drawing.Color.Empty
                                    Else
                                        LEGEND = ASCMAIN1.Get_Legend_Wk(ASCMAIN1.Week_Calc(RYW0, P - 1))
                                        LEGEND = Mid(LEGEND, 10, 7)
                                        If New String() {"Jan", "Mar", "May", "Jul", "Sep", "Nov"}.Contains(Mid(LEGEND, 1, 3)) Then
                                            .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                                            .Columns(COLUMN_NAME).CellAppearance.BackColor2 = Drawing.Color.LightGreen
                                        Else
                                            .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                                            .Columns(COLUMN_NAME).CellAppearance.BackColor2 = Drawing.Color.LightBlue
                                        End If
                                    End If
                                End If
                                .Columns(COLUMN_NAME).Header.Caption = LEGEND
                            End If
                        End If
                    Next

                    With .Columns("PXX")
                        .Header.Appearance.BackColor = Drawing.Color.White
                        .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                        .Header.Appearance.BackColor2 = Drawing.Color.Yellow
                        .Hidden = Not (optType1.Value = "T")
                        .Header.Caption = "O/H"
                    End With
                End With
            Next
        Next
    End Sub

    Private Sub grdSATISLS1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSATISLS1.AfterRowActivate
        Setup_grdSATISLS2()
    End Sub

    Sub Setup_grdSATISLS2()

        If grdSATISLS1.ActiveRow Is Nothing OrElse Not grdSATISLS1.ActiveRow.IsDataRow Then
            chkShowDetails.Checked = False
            chkShowDetails.Enabled = False
            Exit Sub
        Else
            chkShowDetails.Enabled = True
        End If

        Dim DATA_TYPE As String = optType1.Value & optType2.Value
        Dim CODE_VALUE_PARENT As String = grdSATISLS1.ActiveRow.Cells("CODE_VALUE").Text

        Load_SATISLS2(DATA_TYPE, CODE_VALUE_PARENT, False)
        Sort_grdColumns(grdSATISLS2, "CODE_VALUE")

        Dim CAPTION As String = optType1.Text & " (" & optType2.Text & ") for " & ITEM_CODE
        CAPTION &= ", by Customer"
        grdSATISLS2.Text = CAPTION

        Dim sql As String = ""
        sql = sqlSOTINVHX & " and SOTINVH2.ITEM_CODE = '" & ITEM_CODE & "'" & vbCrLf _
            & " and SOTINVH2.ORDR_YYYYPP_UPDATED between '" & RYP0 & "' and '" & RYP1 & "'" & vbCrLf
        sql &= " and SOTINVH2.CUST_CODE = '" & CODE_VALUE_PARENT & "'" & vbCrLf
        With grdSOTINVHX.DisplayLayout.Bands(0)
            .Columns("CUST_STORE_NO").Hidden = False
            .Columns("CUST_STORE_LOCATION").Hidden = False
            .Columns("ITEM_CODE").Hidden = True
            .Columns("ITEM_DESC").Hidden = True
        End With
        Fill_Records("SOTINVHX", "", True, sql)
        grdSOTINVHX.Text = "Sales Documents for " & ITEM_CODE & " - Customer " & CODE_VALUE_PARENT
        grdSOTINVHX.DisplayLayout.CaptionVisible = DefaultableBoolean.True

    End Sub

    Sub Load_SATISLS2(
                     ByVal DATA_TYPE As String,
                     ByVal CODE_VALUE_PARENT As String,
                     ByVal all_parents As Boolean,
                     Optional clear_before_filling As Boolean = True)

        Dim sql As String = ""
        ' & ", NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_NAME) || DECODE(CUST_STORE_MARK_FOR,NULL,'',' (' || CUST_STORE_MARK_FOR || ')') DESC_VALUE" & vbCrLf _

        sql = "Select SATISLS2.CUST_CODE CODE_VALUE_PARENT " & vbCrLf _
          & ", SATISLS2.CUST_STORE_NO CODE_VALUE" & vbCrLf _
          & ", ARTCUST2.CUST_STORE_NAME DESC_VALUE" & vbCrLf _
          & ", ARTCUST2.SELL_CODE SUB_CODE_VALUE1" & vbCrLf _
          & ", ARTCUST2.CUST_STORE_STATE SUB_CODE_VALUE2" & vbCrLf _
          & ", ARTCUST2.CUST_STORE_GROUP SUB_CODE_VALUE3" & vbCrLf _
          & ", ARTCUST2.CUST_STORE_CITY SUB_CODE_VALUE4" & vbCrLf _
          & ", ARTCUST2.CUST_STORE_ZIP_CODE SUB_CODE_VALUE5" & vbCrLf _
          & sqlSATISLS1 & ",PXX  from ARTCUST2," & SATISLS2 & " SATISLS2 " & vbCrLf _
          & " where ARTCUST2.CUST_CODE (+) = SATISLS2.CUST_CODE " & vbCrLf _
          & " and ARTCUST2.CUST_STORE_NO (+) = SATISLS2.CUST_STORE_NO" & vbCrLf _
          & " and DATA_TYPE = '" & DATA_TYPE & "'" & vbCrLf _
          & IIf(all_parents, "", " and SATISLS2.CUST_CODE = '" & CODE_VALUE_PARENT & "'")
        'dst.Tables("SATISLS1").Rows.Clear()
        Fill_Records("SATISLS2", "", clear_before_filling, sql)
    End Sub

    Private Sub chkNoDetails_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowDetails.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        SplitContainer1.Panel2Collapsed = Not chkShowDetails.Checked
    End Sub

    Sub Load_Data()

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Dim DATA_TYPE As String = optType1.Value & optType2.Value

        Dim sql As String = ""

        If optXP.Value = "S" Then

            Dim SQLX As String = sqlSATISLS1_STORES
            If Mid(DATA_TYPE, 1, 1) = "H" Then
                SQLX = Replace(sqlSATISLS1_STORES, Mid(PERIODS_XX, 2), "P" & Format(Periods, "000"))
            End If

            sql = "Select SATISLS2.ITEM_CODE CODE_VALUE" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC DESC_VALUE" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE SUB_CODE_VALUE1" & vbCrLf _
            & ", ICTITEM1.ITEM_CATGY_CODE SUB_CODE_VALUE2" & vbCrLf _
            & ", ICTITEM1.ITEM_CLASS_CODE SUB_CODE_VALUE3" & vbCrLf _
            & ", ICTITEM1.STYLE_CODE SUB_CODE_VALUE4" & vbCrLf _
            & ", ICTITEM1.DEPT_CODE SUB_CODE_VALUE5" & vbCrLf _
            & ", ICTITEM1.ITEM_RETAIL_PRICE RTL_PRICE" & vbCrLf _
            & ", ICTITEM1.ITEM_PRICE WSL_PRICE" & vbCrLf _
            & SQLX & "  from ICTITEM1," & SATISLS2 & " SATISLS2 " & vbCrLf _
            & " where ICTITEM1.ITEM_CODE (+) = SATISLS2.ITEM_CODE " & vbCrLf _
            & " and SATISLS2.DATA_TYPE = '" & DATA_TYPE & "'" & vbCrLf _
            & " and SATISLS2.YEAR = '0'" & vbCrLf _
            & " group by SATISLS2.ITEM_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_CATGY_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_CLASS_CODE" & vbCrLf _
            & ", ICTITEM1.STYLE_CODE" & vbCrLf _
            & ", ICTITEM1.DEPT_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_RETAIL_PRICE" & vbCrLf _
            & ", ICTITEM1.ITEM_PRICE" & vbCrLf
        Else
            sql = "Select SATISLS1.CODE_VALUE CODE_VALUE" & vbCrLf _
            & ", ARTCUST1.CUST_NAME DESC_VALUE" & vbCrLf _
            & ", ARTCUST1.SREP_CODE SUB_CODE_VALUE1" & vbCrLf _
            & ", ARTCUST1.CUST_STATE SUB_CODE_VALUE2" & vbCrLf _
            & ", ARTCUST1.TRADE_CLASS_CODE SUB_CODE_VALUE3" & vbCrLf _
            & ", ARTCUST1.CUST_CITY SUB_CODE_VALUE4" & vbCrLf _
            & ", ARTCUST1.CUST_ZIP_CODE SUB_CODE_VALUE5" & vbCrLf _
            & ", 0 RTL_PRICE" & vbCrLf _
            & ", 0 WSL_PRICE" & vbCrLf _
            & sqlSATISLS1 & ",PXX from ARTCUST1," & SATISLS1 & " SATISLS1 " & vbCrLf _
            & " where ARTCUST1.CUST_CODE (+) = SATISLS1.CODE_VALUE " & vbCrLf _
            & " and SATISLS1.SI = '" & "I" & "'" _
            & " and SATISLS1.DATA_TYPE = '" & DATA_TYPE & "'" _
            & " and SATISLS1.YEAR = '0'" & vbCrLf
        End If

        dst.Tables("SATISLS1").Rows.Clear()
        dst.Tables("SATISLS1_DTL").Rows.Clear()
        dst.Tables("SATISLS2").Rows.Clear()

        If optXP.Value = "S" Then
            dst.Tables("SATISLS1").Columns("P000").Expression = Mid(STORES_XX, 2)
            dst.Tables("SATISLS1_DTL").Columns("P000").Expression = Mid(STORES_XX, 2)
        Else
            If optType1.Value = "H" Then
                ' NOTE THAT SAFCSLS1 HAS A 2 DIGIT PERIOD AND SAFISLS1 HAS A 3 DIGIT PERIOD
                dst.Tables("SATISLS1").Columns("P000").Expression = "P" & Mid(PERIODS_XX, Len(PERIODS_XX) - 2, 3)
            Else
                dst.Tables("SATISLS1").Columns("P000").Expression = Mid(PERIODS_XX, 2)
            End If
            dst.Tables("SATISLS1_DTL").Columns("P000").Expression = "IIF(DATA_TYPE='HU' OR DATA_TYPE='HD'," & Mid(PERIODS_XX, Len(PERIODS_XX) - 2, 3) & "," & Mid(PERIODS_XX, 2) & ")"
        End If

        If optType1.Value = "H" Then
            ' NOTE THAT SAFCSLS1 HAS A 2 DIGIT PERIOD AND SAFISLS1 HAS A 3 DIGIT PERIOD
            dst.Tables("SATISLS2").Columns("P000").Expression = "P" & Mid(PERIODS_XX, Len(PERIODS_XX) - 2, 3)
        Else
            dst.Tables("SATISLS2").Columns("P000").Expression = Mid(PERIODS_XX, 2)
        End If

        Fill_Records("SATISLS1", "", True, sql)
        Sort_grdColumns(grdSATISLS1, "CODE_VALUE")

        If chkExtendedData.Checked Then
            sql = "Select SATISLS1.CODE_VALUE CODE_VALUE, SATISLS1.YEAR, SATISLS1.DATA_TYPE" _
            & sqlSATISLS1 _
            & " from " & SATISLS1 & " SATISLS1 " _
            & " where SATISLS1.SI = '" & "I" & "'" _
            & " and SATISLS1.DATA_TYPE = '" & DATA_TYPE & "'" _
            & " and SATISLS1.YEAR = '0'" & vbCrLf

            ' these filters caused constraint enforcement issues 
            ' for 56100, extended Data, by Week, 2017-08 thru 2017-21
            '& " and SATISLS1.DATA_TYPE LIKE '%" & Mid(DATA_TYPE, 2, 1) & "'" _
            '& " and (YEAR = '1' OR SATISLS1.DATA_TYPE <> '" & DATA_TYPE & "')"

            Fill_Records("SATISLS1_DTL", "", True, sql)
            Sort_grdColumns(grdSATISLS1, "DATA_TYPE", , 1)
        End If

        If grdSATISLS1.Rows.Count = 0 Then
            chkShowDetails.Checked = False
            chkShowDetails.Enabled = False
        Else
            chkShowDetails.Enabled = True
        End If

        'For Each rowDTL As DataRow In ASCMAIN1.Distinct_Values("", dst.Tables("SATISLS1_DTL"), New String() {"CODE_VALUE"}).Select()
        '    Dim CODE_VALUE As String = rowDTL.Item("CODE_VALUE")
        '    If dst.Tables("SATISLS1").Rows.Find(CODE_VALUE) Is Nothing Then
        '        dst.Tables("SATISLS1").Rows.Add(New String() {CODE_VALUE})
        '    End If
        'Next

        'dst.EnforceConstraints = True
        EnforceConstraints(True)

        Setup_grd()

        tabDetails.Tabs("Sales Documents").Visible = (optType1.Value <> "T")

        If tabDetails.SelectedTab.Key = "Map" Then
            tabDetails.SelectedTab = tabDetails.Tabs("Details")
        End If
        tabDetails.Tabs("Map").Visible = False

        CreateGraph_SATISLS1()
        CreateGraph_SATISLS1_X()
        chtSATISLS1.Visible = True
        chtSATISLS1_X.Visible = True

        If grdSATISLS1.DisplayLayout.ValueLists.Exists("SUB_CODE_VALUE1") Then
            grdSATISLS1.DisplayLayout.ValueLists.Remove("SUB_CODE_VALUE1")
        End If
        If grdSATISLS2.DisplayLayout.ValueLists.Exists("SUB_CODE_VALUE2") Then
            grdSATISLS2.DisplayLayout.ValueLists.Remove("SUB_CODE_VALUE2")
        End If

        ASCMAIN1.Add_Value_List(grdSATISLS1, "SUB_CODE_VALUE2", , , , "Select ITEM_CATGY_CODE, ITEM_CATGY_DESC from ICTCATG1")
        ASCMAIN1.Add_Value_List(grdSATISLS2, "SUB_CODE_VALUE1", , , , "Select SELL_CODE, SELL_NAME from SOTSELL1")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub optType1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optType1.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_Data()
    End Sub

    Private Sub optType2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optType2.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_Data()
    End Sub

    Sub CreateGraph_SATISLS1()

        Dim chtIsVisible As Boolean = chtSATISLS1.Visible
        chtSATISLS1.Visible = False

        chtSATISLS1.DataSource = Nothing

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Charting Data")

        Me.SuspendLayout()

        Dim RL() As String

        chtSATISLS1.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtSATISLS1.LabelHash = labelHash

        chtSATISLS1.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtSATISLS1.Tooltips.FormatString = "<HIGHLOW>"

        Dim DTY As New DataTable
        With DTY
            .Columns.Add("CODE")
            .Columns.Add("VALUE", GetType(System.Decimal))
        End With

        Dim RLi As Integer = 0
        ReDim RL(dst.Tables("SATISLS1").Rows.Count - 1)
        For Each row As DataRow In dst.Tables("SATISLS1").Select("", "CODE_VALUE")
            RL(RLi) = row("CODE_VALUE") ' & ":" & row("DESC_VALUE")
            RLi += 1
            DTY.Rows.Add(New Object() {row.Item("CODE_VALUE"), row.Item("P000")})
        Next
        'chtSATISLS1.Data.SetRowLabels(RL)
        'chtSATISLS1.Data.SetColumnLabels(CL)

        'chtSATISLS1.DataSource = dst.Tables("SATISLS1")
        chtSATISLS1.DataSource = DTY
        chtSATISLS1.PieChart.ColumnIndex = -1
        chtSATISLS1.PieChart.OthersCategoryPercent = 2
        'chtSATISLS1.Data.IncludeColumn("CODE_VALUE", False)
        'chtSATISLS1.Data.IncludeColumn("DESC_VALUE", False)
        'chtSATISLS1.Data.IncludeColumn("P00", True)


        chtSATISLS1.DataBind()

        chtSATISLS1.Visible = chtIsVisible

        Me.ResumeLayout()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()
    End Sub

    Sub CreateGraph_SATISLS1_X()

        Dim chtIsVisible As Boolean = chtSATISLS1_X.Visible
        chtSATISLS1_X.Visible = False

        chtSATISLS1_X.DataSource = Nothing

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
        '    grdSATISLS1.DisplayLayout.Bands(0).Columns("M" & Format(i, "000")).Header.Caption = Mid(L, 10, 3)
        'Next
        For i As Integer = 1 To Periods
            'Dim L As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * i))
            CL(i - 1) = grdSATISLS1.DisplayLayout.Bands(0).Columns("P" & Format(i, "000")).Header.Caption
            'grdSATISLS1.DisplayLayout.Bands(0).Columns("M" & Format(i, "000")).Header.Caption = Mid(L, 10, 3)
        Next

        'chtICTINVAT.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.LabelPlusDataValue
        'chtICTINVAT.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom

        chtSATISLS1_X.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtSATISLS1_X.LabelHash = labelHash

        chtSATISLS1_X.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtSATISLS1_X.Tooltips.FormatString = "<HIGHLOW>"

        Dim DT As New DataTable
        DT.Columns.Add("CODE_VALUE")
        DT.Columns.Add("DESC_VALUE")
        For P As Integer = 1 To Periods
            DT.Columns.Add("P" & Format(P, "000"), GetType(System.Decimal))
        Next

        Dim RLi As Integer = 0
        ReDim RL(dst.Tables("SATISLS1").Rows.Count - 1)
        For Each row As DataRow In dst.Tables("SATISLS1").Select("", "CODE_VALUE")
            RL(RLi) = row("CODE_VALUE") ' & ":" & row("DESC_VALUE")
            RLi += 1

            Dim rowDT As DataRow = DT.NewRow
            rowDT.Item("CODE_VALUE") = row("CODE_VALUE")
            rowDT.Item("DESC_VALUE") = row("DESC_VALUE")
            For P As Integer = 1 To Periods
                rowDT.Item("P" & Format(P, "000")) = row("P" & Format(P, "000"))
            Next
            DT.Rows.Add(rowDT)
        Next
        chtSATISLS1_X.Data.SetRowLabels(RL)
        chtSATISLS1_X.Data.SetColumnLabels(CL)

        chtSATISLS1_X.DataSource = DT
        'chtSATISLS1_X.Data.IncludeColumn("CODE_VALUE", False)
        'chtSATISLS1_X.Data.IncludeColumn("DESC_VALUE", False)
        'chtSATISLS1_X.Data.IncludeColumn("P00", False)

        chtSATISLS1_X.DataBind()

        chtSATISLS1_X.Visible = chtIsVisible

        Me.ResumeLayout()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()
    End Sub

    Private Sub trkbrXAxis_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles trkbrXAxis.Scroll
        chtSATISLS1_X.Axis.X.ScrollScale.Scale = Me.trkbrXAxis.Value / 100.0
    End Sub

    Private Sub trkbrYAxis_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles trkbrYAxis.Scroll
        chtSATISLS1_X.Axis.Y.ScrollScale.Scale = Me.trkbrYAxis.Value / 100.0
    End Sub

    Private Sub tabDetails_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabDetails.SelectedTabChanged
        Setup_tabDetails()
    End Sub

    Sub Setup_tabDetails()
        If SELECTION_NO = 0 Then Exit Sub
        UltraExplorerBar1.Groups("Charts").Visible = (tabDetails.SelectedTab.Key = "Charts") And ScreenMode
        UltraExplorerBar1.Groups("Item Image").Visible = Not (tabDetails.SelectedTab.Key = "Charts") And ScreenMode
    End Sub

    Private Sub optTotalsChartType_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTotalsChartType.ValueChanged
        Set_Totals_ChartType()
    End Sub

    Private Sub chkTotalsChart3D_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTotalsChart3D.CheckedChanged
        Set_Totals_ChartType()
    End Sub

    Sub Set_Totals_ChartType()
        If Not chkTotalsChart3D.Checked Then
            chtSATISLS1_X.ChartType = ChartType.LineChart
            Select Case optTotalsChartType.Value
                Case "PieChart"
                    chtSATISLS1.ChartType = ChartType.PieChart
                Case "DoughnutChart"
                    chtSATISLS1.ChartType = ChartType.DoughnutChart
            End Select
        Else
            chtSATISLS1_X.ChartType = ChartType.LineChart3D
            Select Case optTotalsChartType.Value
                Case "PieChart"
                    chtSATISLS1.ChartType = ChartType.PieChart3D
                Case "DoughnutChart"
                    chtSATISLS1.ChartType = ChartType.DoughnutChart3D
            End Select

        End If
    End Sub

    Private Sub cbeColor_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeColor.ValueChanged
        'chtSATISLS1.ColorModel.ModelStyle = cbeColor.ValueMember
        'chtSATISLS1_X.ColorModel.ModelStyle = Infragistics.UltraChart.Shared.Styles.ColorModels.PureRandom
        chtSATISLS1.ColorModel.ModelStyle = CType(System.Enum.Parse(GetType(ColorModels), cbeColor.SelectedItem.ToString()), ColorModels)
        chtSATISLS1_X.ColorModel.ModelStyle = CType(System.Enum.Parse(GetType(ColorModels), cbeColor.SelectedItem.ToString()), ColorModels)

    End Sub

    Sub Setup_Map()
        '' create the layer
        Dim points As String = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), ASCMAIN1.Folders("Images") & "ABS\UsMap\US_STATES.xml")
        'Dim USmap As New MapLayer(points)
        USmap = New MapLayer(points)

        dst.Tables("SATISLSS").Rows.Clear()
        US_STATES = USmap.STATES
        For i As Integer = 0 To USmap.STATES.Length - 1
            dst.Tables("SATISLSS").Rows.Add(New Object() {"", USmap.STATES(i), 0})
        Next

        '' set the layer
        Me.UltraChart1.ChartType = ChartType.Composite
        Me.UltraChart1.CompositeChart.ChartAreas.Add(New Infragistics.UltraChart.Resources.Appearance.ChartArea())
        Me.UltraChart1.UserLayerIndex = New String() {"USMap"}
        Me.UltraChart1.Layer.Add("USMap", USmap)

        '' set the tooltip.
        Dim labelRenderers As New Hashtable()
        labelRenderers.Add("USMap", New USMapLabelRenderer(dst.Tables("SATISLSS")))
        Me.UltraChart1.LabelHash = labelRenderers
        Me.UltraChart1.Tooltips.FormatString = "<USMap>"

        ''set border
        Me.UltraChart1.Border.CornerRadius = 20
        Me.UltraChart1.Border.Thickness = 0
        Me.UltraChart1.BackColor = System.Drawing.Color.White

        '' set color model
        'Me.UltraChart1.ColorModel.ColorBegin = Color.AliceBlue
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
                Dim rows() As DataRow = dst.Tables("SATISLSS").Select("STATE_NAME = '" & US_STATES(I) & "'")
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

    Private Sub grdSATISLSS_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSATISLSS.DoubleClickRow
        Show_Filter(grdSATISLS1, True)
        grdSATISLS1.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        grdSATISLS1.Rows.ColumnFilters("SUB_CODE_VALUE2").FilterConditions.Add _
        (Infragistics.Win.UltraWinGrid.FilterComparisionOperator.Equals, e.Row.Cells("STATE_CODE").Text)
        chkShowDetails.Checked = True
    End Sub

    Private Sub grdSATISLSS_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSATISLSS.InitializeRow
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
        grdSATSIST1.Text = "Sell-In / Sell-Thru"
        grdSATSIST3.Text = "Sell-In / Sell-Thru LY" & IIf(ITEM_CODE_LY = ITEM_CODE, "", " (" & ITEM_CODE_LY & ")")

        Load_Data()
    End Sub

    Sub Load_SIST()
        Load_SATSIST1()
        Setup_tabSIST()
        grdSATSIST1.Text = "Sell-In / Sell-Thru"
        grdSATSIST3.Text = "Sell-In / Sell-Thru LY" & IIf(ITEM_CODE_LY = ITEM_CODE, "", " (" & ITEM_CODE_LY & ")")
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

        ASCMAIN1.sql = "Select OPS_YYYYPP from GLTPARM2 where OPS_YYYYPP = '" & P0 & "'"
        Absx1.cmbFor("RYP0").Value = ASCDATA1.GetDataValue
        ASCMAIN1.sql = "Select OPS_YYYYPP from GLTPARM2 where OPS_YYYYPP = '" & P1 & "'"
        Absx1.cmbFor("RYP1").Value = ASCDATA1.GetDataValue

        ASCMAIN1.sql = "Select MIN (YYYYWW) from GLTPARM3 where YYYYMM = '" & P0 & "'"
        Absx1.cmbFor("RYW0").Value = ASCDATA1.GetDataValue
        ASCMAIN1.sql = "Select MAX (YYYYWW) from GLTPARM3 where YYYYMM = '" & P1 & "'"
        Absx1.cmbFor("RYW1").Value = ASCDATA1.GetDataValue
    End Sub

    Sub Load_SATSIST1()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Sell-In & Sell-Thru")

        Dim sqlp_IN As String = ""
        Dim sqlp_TH As String = ""
        Dim sqlp_THOH As String = ""

        Dim YX_IN1 As String = ""
        Dim YX_IN2 As String = ""
        Dim YX_TH1 As String = ""
        Dim YX_TH2 As String = ""

        If chkSIST_Multi.Checked Then
            Dim rowSATSISTI As DataRow = dst.Tables("SATSISTI").Rows.Find(ITEM_CODE)
            If rowSATSISTI IsNot Nothing Then
                YX_IN1 = rowSATSISTI.Item("YX_IN1") & ""
                YX_IN2 = rowSATSISTI.Item("YX_IN2") & ""
                YX_TH1 = rowSATSISTI.Item("YX_TH1") & ""
                YX_TH2 = rowSATSISTI.Item("YX_TH2") & ""
            End If
        End If

        Dim PYX As String = IIf(YX_TH1 <> "", YX_TH1, WPs(0))
        If optRANGE.Value = "P" Then
            PYX = ASCMAIN1.Period_Calc(PYX, -1)
        Else
            PYX = ASCMAIN1.Week_Calc(PYX, -1)
        End If

        'Dim sqlp As String = ""
        Dim P_SI As String = ""
        Dim P_ST As String = ""
        If optRANGE.Value = "P" Then
            sqlp_IN = " and ORDR_YYYYPP_UPDATED between '" & IIf(YX_IN1 <> "", YX_IN1, RYP0) & "' and '" & IIf(YX_IN2 <> "", YX_IN2, RYP1) & "'"
            sqlp_TH = " and ORDR_YYYYPP_UPDATED between '" & IIf(YX_TH1 <> "", YX_TH1, RYP0) & "' and '" & IIf(YX_TH2 <> "", YX_TH2, RYP1) & "'"
            sqlp_THOH = " and ORDR_YYYYPP_UPDATED = '" & PYX & "'"
            'sqlp = " and ORDR_YYYYPP_UPDATED between '" & RYP0 & "' and '" & RYP1 & "'"
            P_SI = "SOTINVH2.ORDR_YYYYPP_UPDATED"
            P_ST = "OPS_YYYYPP"
        Else
            sqlp_IN = " and OPS_YYYYWW between '" & IIf(YX_IN1 <> "", YX_IN1, RYW0) & "' and '" & IIf(YX_IN2 <> "", YX_IN2, RYW1) & "'"
            sqlp_TH = " and OPS_YYYYWW between '" & IIf(YX_TH1 <> "", YX_TH1, RYW0) & "' and '" & IIf(YX_TH2 <> "", YX_TH2, RYW1) & "'"
            sqlp_THOH = " and OPS_YYYYWW = '" & PYX & "'"
            'sqlp = " and OPS_YYYYWW between '" & RYW0 & "' and '" & RYW1 & "'"
            P_SI = "SOTINVH2.OPS_YYYYWW"
            P_ST = "OPS_YYYYWW"
        End If


        Dim sql_SI As String = ", 0 EOW"
        Dim sql_ST As String = ", 0 EOW"
        Dim sql_STOH As String = ", SUM(QTY_EOW) EOW"
        Dim sql_SIST As String = ", SUM (EOW) EOW"
        Dim sql_cols As String = ", EOW"

        For P As Integer = 1 To Periods
            Dim PX As String = "_P" & Format(P, "00")
            sql_SI &= ", SUM(DECODE(" & P_SI & ",'" & Ps(P - 1) & "',SOTINVH2.ORDR_QTY_SHIP,0)) SIQTY" & PX & ", 0 OHQTY" & PX & ", 0 STQTY" & PX & ", 0 STAMT" & PX & vbCrLf
            sql_ST &= ", 0 SIQTY" & PX _
                    & ", SUM(DECODE(OPS_YYYYWW,'" & WPs(P - 1) & "',QTY_EOW,0)) OHQTY" & PX _
                    & ", SUM(DECODE(" & P_ST & ",'" & Ps(P - 1) & "',QTY_SOLD,0)) STQTY" & PX _
                    & ", SUM(DECODE(" & P_ST & ",'" & Ps(P - 1) & "',AMT_SOLD,0)) STAMT" & PX & vbCrLf
            sql_STOH &= ", 0 SIQTY" & PX _
                    & ", 0 OHQTY" & PX _
                    & ", 0 STQTY" & PX _
                    & ", 0 STAMT" & PX & vbCrLf
            sql_SIST &= ", SUM (SIQTY" & PX & ") SIQTY" & PX _
                      & ", SUM (OHQTY" & PX & ") OHQTY" & PX _
                      & ", SUM (STQTY" & PX & ") STQTY" & PX _
                      & ", SUM (STAMT" & PX & ") STAMT" & PX & vbCrLf
            sql_cols &= ", SIQTY" & PX _
                      & ", OHQTY" & PX _
                      & ", STQTY" & PX _
                      & ", STAMT" & PX & vbCrLf

            grdSATSIST1.DisplayLayout.Bands(0).Groups(P + 1).Header.Caption = Ls(P - 1)
            grdSATSIST1.DisplayLayout.Bands(0).Groups(P + 1).Hidden = False
            grdSATSIST1.DisplayLayout.Bands(0).Groups(P + 1).Header.ToolTipText = ""

            grdSATSIST2.DisplayLayout.Bands(0).Groups(P + 1).Header.Caption = Ls(P - 1)
            grdSATSIST2.DisplayLayout.Bands(0).Groups(P + 1).Hidden = False
            grdSATSIST2.DisplayLayout.Bands(0).Groups(P + 1).Header.ToolTipText = ""

            grdSATSIST3.DisplayLayout.Bands(0).Groups(P + 1).Header.Caption = Ls(P - 1)
            grdSATSIST3.DisplayLayout.Bands(0).Groups(P + 1).Hidden = False
            grdSATSIST3.DisplayLayout.Bands(0).Groups(P + 1).Header.ToolTipText = ""
        Next

        Dim sql_SI_orig As String = sql_SI
        Dim sql_ST_orig As String = sql_ST
        Dim sql_STOH_orig As String = sql_STOH

        If Periods < SIST_PMAX Then
            For P As Integer = Periods + 1 To SIST_PMAX
                grdSATSIST1.DisplayLayout.Bands(0).Groups(P + 1).Hidden = True
                grdSATSIST2.DisplayLayout.Bands(0).Groups(P + 1).Hidden = True
                grdSATSIST3.DisplayLayout.Bands(0).Groups(P + 1).Hidden = True
            Next
        End If

        Dim CODEs As String = Get_Multi_CODEs("SATSISTC")


        CODEs2EXCLUDE = Get_Multi_CODEs("SATSISTX")
        If Not chkExcludeCustomers.Checked Then
            CODEs2EXCLUDE = ""
        End If

        'Fill_Records("SATSISTY", ITEM_CODE)
        'Dim ITEM_CODE_XLYs As New Dictionary(Of String, String)
        'For Each ROW As DataRow In dst.Tables("SATSISTY").Select("", "LY")
        '    Dim ITEM_CODE_COMPARE_TO As String = ROW.Item("ITEM_CODE_COMPARE_TO")
        '    Dim LY As String = ROW.Item("LY")
        '    ITEM_CODE_XLYs.Add(ITEM_CODE_COMPARE_TO, LY)
        'Next

        If SATSIST3 <> "" Then
            ASCDATA1.ExecuteSQL("Truncate Table " & SATSIST3)
        End If

        For Each TABLE_NAME As String In New String() {"SATSIST1", "SATSIST2", "SATSIST3"}
            Dim sqlITEM_CODE As String = "ITEM_CODE = '" & ITEM_CODE & "'"

            Dim xLYs As New List(Of String)
            If TABLE_NAME = "SATSIST3" Then
                For Each rowSATSISTY As DataRow In dst.Tables("SATSISTY").Select($"ITEM_CODE = '{ITEM_CODE}'")
                    Dim ITEM_CODE_COMPARE_TO As String = rowSATSISTY.Item("ITEM_CODE_COMPARE_TO")
                    Dim LY As String = rowSATSISTY.Item("LY")
                    xLYs.Add(LY & ":" & ITEM_CODE_COMPARE_TO)
                Next
            Else
                xLYs.Add("0" & ":" & ITEM_CODE)
            End If

            For Each xLYi As String In xLYs

                Dim xLY As String = Split(xLYi, ":")(0)

                If TABLE_NAME = "SATSIST3" Then

                    Dim ITEM_CODE_COMPARE_TO As String = Split(xLYi, ":")(1)

                    If ASCMAIN1.Running_in_VS And ITEM_CODE_COMPARE_TO = "CC007C08USA" Then Stop

                    sql_SI = sql_SI_orig
                    sql_ST = sql_ST_orig
                    sql_STOH = sql_STOH_orig
                    For P As Integer = 1 To Periods
                        Dim PX As String = "_P" & Format(P, "00")
                        If xLY = "1" Then
                            sql_SI = Replace(sql_SI, "'" & Ps(P - 1) & "'", "'" & PsLY(P - 1) & "'")
                            sql_ST = Replace(sql_ST, "'" & Ps(P - 1) & "'", "'" & PsLY(P - 1) & "'")
                            sql_STOH = Replace(sql_STOH, "'" & Ps(P - 1) & "'", "'" & PsLY(P - 1) & "'")
                        ElseIf xLY = "2" Then
                            sql_SI = Replace(sql_SI, "'" & Ps(P - 1) & "'", "'" & Ps2LY(P - 1) & "'")
                            sql_ST = Replace(sql_ST, "'" & Ps(P - 1) & "'", "'" & Ps2LY(P - 1) & "'")
                            sql_STOH = Replace(sql_STOH, "'" & Ps(P - 1) & "'", "'" & Ps2LY(P - 1) & "'")
                        End If

                        'unremming this so that we see LY headings on the LY grid
                        If xLY = "1" Then grdSATSIST3.DisplayLayout.Bands(0).Groups(P + 1).Header.Caption = LsLY(P - 1)
                        'grdSATSIST3.DisplayLayout.Bands(0).Groups(P + 1).Hidden = False
                        'grdSATSIST1.DisplayLayout.Bands(0).Groups(P + 1).Header.ToolTipText = ""
                    Next

                    sqlITEM_CODE = "ITEM_CODE = '" & ITEM_CODE_COMPARE_TO & "'"
                    Dim rowSATSISTY As DataRow = dst.Tables("SATSISTY").Rows.Find(New String() {ITEM_CODE, xLY, ITEM_CODE_COMPARE_TO})

                    YX_IN1 = rowSATSISTY.Item("YX_IN1") & ""
                    YX_IN2 = rowSATSISTY.Item("YX_IN2") & ""
                    YX_TH1 = rowSATSISTY.Item("YX_TH1") & ""
                    YX_TH2 = rowSATSISTY.Item("YX_TH2") & ""

                    Dim LY0 As String = ""
                    Dim LY1 As String = ""

                    Dim pAdj As Integer = 0
                    If optRANGE.Value = "P" Then
                        pAdj = IIf(xLY = "1", -12 * 1, -12 * 2)
                        LY0 = ASCMAIN1.Period_Calc(RYP0, pAdj)
                        LY1 = ASCMAIN1.Period_Calc(RYP1, pAdj)
                        If YX_IN1 = "" Then YX_IN1 = LY0
                        If YX_IN2 = "" Then YX_IN2 = LY1
                        If YX_TH1 = "" Then YX_TH1 = LY0
                        If YX_TH2 = "" Then YX_TH2 = LY1
                        sqlp_IN = $" and ORDR_YYYYPP_UPDATED between '{YX_IN1}' and '{YX_IN2}'"
                        sqlp_TH = $" and ORDR_YYYYPP_UPDATED between '{YX_TH1}' and '{YX_TH2}'"
                        sqlp_THOH = $" and ORDR_YYYYPP_UPDATED = '{ASCMAIN1.Period_Calc(YX_TH1, -1)}'"
                    Else
                        pAdj = IIf(xLY = "1", -52 * 1, -52 * 2)
                        LY0 = ASCMAIN1.Week_Calc(RYW0, pAdj)
                        LY1 = ASCMAIN1.Week_Calc(RYW1, pAdj)
                        If ASCMAIN1.CLIENT = "INT" Then
                            pAdj = IIf(xLY = "1", 1, 2)
                            LY0 = Format(Val(Mid(RYW0, 1, 4)) - pAdj, "0000") & Mid(RYW0, 5, 2)
                            LY1 = Format(Val(Mid(RYW1, 1, 4)) - pAdj, "0000") & Mid(RYW1, 5, 2)
                        End If
                        If YX_IN1 = "" Then YX_IN1 = LY0
                        If YX_IN2 = "" Then YX_IN2 = LY1
                        If YX_TH1 = "" Then YX_TH1 = LY0
                        If YX_TH2 = "" Then YX_TH2 = LY1
                        sqlp_IN = $" and OPS_YYYYWW between '{YX_IN1}' and '{YX_IN2}'"
                        sqlp_TH = $" and OPS_YYYYWW between '{YX_TH1}' and '{YX_TH2}'"
                        sqlp_THOH = " and OPS_YYYYWW = '" & ASCMAIN1.Week_Calc(YX_TH1, -1) & "'"
                    End If
                End If

                ASCMAIN1.sql = "Select X.ITEM_CODE, X.CUST_CODE" & vbCrLf _
                    & IIf(TABLE_NAME = "SATSIST2", ",X.CUST_STORE_NO", "") & vbCrLf _
                    & IIf(TABLE_NAME = "SATSIST2", ",ARTCUST2.CUST_STORE_NAME CUST_NAME", ",ARTCUST1.CUST_NAME") & vbCrLf _
                    & ", Min (X.INV_DATE_SHIPPED) INV_DATE_SHIPPED" & vbCrLf _
                    & sql_SIST & " from (" & vbCrLf _
                    & "Select SOTINVH2.ITEM_CODE, SOTINVH1.CUST_CODE" & vbCrLf _
                    & IIf(TABLE_NAME = "SATSIST2", ",SOTINVH1.CUST_STORE_NO", "") & vbCrLf _
                    & ", Min (SOTINVH1.INV_DATE_SHIPPED) INV_DATE_SHIPPED" & vbCrLf _
                    & sql_SI _
                    & " from SOTINVH1,SOTINVH2" & vbCrLf _
                    & " where SOTINVH2." & sqlITEM_CODE & vbCrLf _
                    & IIf(chkSIST_Multi.Checked And CODEs <> "", " and SOTINVH2.CUST_CODE in (" & CODEs & ")" & vbCrLf, "") _
                    & IIf(TABLE_NAME = "SATSIST2", " and SOTINVH2.CUST_CODE = 'CUST_CODE'", "") & vbCrLf _
                    & Replace(Replace(sqlp_IN, "ORDR_YYYYPP_UPDATED", "SOTINVH2.ORDR_YYYYPP_UPDATED"), "OPS_YYYYWW", "SOTINVH2.OPS_YYYYWW") & vbCrLf _
                    & " and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
                    & " and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
                    & " and SOTINVH1.INV_TYPE = 'I'" & vbCrLf _
                    & " group by SOTINVH2.ITEM_CODE, SOTINVH1.CUST_CODE" & vbCrLf _
                    & IIf(TABLE_NAME = "SATSIST2", ",SOTINVH1.CUST_STORE_NO", "") & vbCrLf _
                    & " union" & vbCrLf _
                    & "Select ITEM_CODE, CUST_CODE" & vbCrLf _
                    & IIf(TABLE_NAME = "SATSIST2", ",CUST_STORE_NO", "") & vbCrLf _
                    & ", NULL INV_DATE_SHIPPED" & vbCrLf _
                    & sql_ST _
                    & " from RSTRETL1" & vbCrLf _
                    & " where " & sqlITEM_CODE & vbCrLf _
                    & IIf(chkSIST_Multi.Checked And CODEs <> "", " and CUST_CODE in (" & CODEs & ")" & vbCrLf, "") _
                    & IIf(TABLE_NAME = "SATSIST2", " and CUST_CODE = 'CUST_CODE'", "") & vbCrLf _
                    & Replace(sqlp_TH, "ORDR_YYYYPP_UPDATED", "OPS_YYYYPP") & vbCrLf _
                    & " group by ITEM_CODE, CUST_CODE" & vbCrLf _
                    & IIf(TABLE_NAME = "SATSIST2", ",CUST_STORE_NO", "") & vbCrLf _
                    & " union" & vbCrLf _
                    & "Select ITEM_CODE, CUST_CODE" & vbCrLf _
                    & IIf(TABLE_NAME = "SATSIST2", ",CUST_STORE_NO", "") & vbCrLf _
                    & ", NULL INV_DATE_SHIPPED" & vbCrLf _
                    & sql_STOH _
                    & " from RSTRETL1" & vbCrLf _
                    & " where " & sqlITEM_CODE & vbCrLf _
                    & IIf(chkSIST_Multi.Checked And CODEs <> "", " and CUST_CODE in (" & CODEs & ")" & vbCrLf, "") _
                    & IIf(TABLE_NAME = "SATSIST2", " and CUST_CODE = 'CUST_CODE'", "") & vbCrLf _
                    & Replace(sqlp_THOH, "ORDR_YYYYPP_UPDATED", "OPS_YYYYPP") & vbCrLf _
                    & " group by ITEM_CODE, CUST_CODE" & vbCrLf _
                    & IIf(TABLE_NAME = "SATSIST2", ",CUST_STORE_NO", "") & vbCrLf _
                    & ") X" & vbCrLf _
                    & IIf(TABLE_NAME = "SATSIST2",
                          ",ARTCUST2 where ARTCUST2.CUST_CODE (+) = X.CUST_CODE and ARTCUST2.CUST_STORE_NO (+) = X.CUST_STORE_NO",
                          ",ARTCUST1 where ARTCUST1.CUST_CODE (+) = X.CUST_CODE") & vbCrLf _
                    & " group by X.ITEM_CODE, X.CUST_CODE" & vbCrLf _
                    & IIf(TABLE_NAME = "SATSIST2", ",X.CUST_STORE_NO", "") & vbCrLf _
                    & IIf(TABLE_NAME = "SATSIST2", ",ARTCUST2.CUST_STORE_NAME", ",ARTCUST1.CUST_NAME")
                '                & IIf(TABLE_NAME = "SATSIST2", ",DECODE(ARTCUST2.CUST_STORE_LOCATION,NULL,ARTCUST2.CUST_STORE_NAME,ARTCUST2.CUST_STORE_NO || ' ' || ARTCUST2.CUST_STORE_LOCATION)", ",ARTCUST1.CUST_NAME")

                If TABLE_NAME = "SATSIST1" Then
                    ' If ASCMAIN1.Running_in_VS And ITEM_CODE = "CH001C95USA" Then Stop
                    If SATSIST1 = "" Then
                        SATSIST1 = ASCMAIN1.Temp_Table
                        ASCDATA1.ExecuteSQL("Alter Table " & SATSIST1 & " Add Primary Key (ITEM_CODE, CUST_CODE)")
                    Else
                        ASCDATA1.ExecuteSQL("Truncate Table " & SATSIST1)
                        ASCDATA1.ExecuteSQL("Insert into " & SATSIST1 & " (ITEM_CODE, CUST_CODE, CUST_NAME, INV_DATE_SHIPPED" & vbCrLf & sql_cols & ") " & ASCMAIN1.sql)
                    End If

                    Fill_and_Cum(SATSIST1, "SATSIST1")

                ElseIf TABLE_NAME = "SATSIST3" Then
                    If SATSIST3 = "" Then
                        SATSIST3 = ASCMAIN1.Temp_Table
                        ASCDATA1.ExecuteSQL("Alter Table " & SATSIST3 & " Add Primary Key (ITEM_CODE, CUST_CODE)")
                    Else
                        'ASCDATA1.ExecuteSQL("Truncate Table " & SATSIST3)
                        ASCDATA1.ExecuteSQL("Insert into " & SATSIST3 & " (ITEM_CODE, CUST_CODE, CUST_NAME, INV_DATE_SHIPPED" & vbCrLf & sql_cols & ") " & ASCMAIN1.sql)
                    End If

                    Fill_and_Cum(SATSIST3, "SATSIST3")
                Else
                    If SATSIST2 = "" Then
                        SATSIST2 = ASCMAIN1.Temp_Table
                        ASCDATA1.ExecuteSQL("Alter Table " & SATSIST2 & " Add Primary Key (ITEM_CODE,CUST_CODE,CUST_STORE_NO)")
                        'SINCE WE ARE WIPING OUT NAME OF TABLE SATSIST2 FOR EACH ITEM WE NEVER HIT THE ELSE
                        sqlSATSIST2 = "Insert into " & SATSIST2 & " (ITEM_CODE, CUST_CODE, CUST_STORE_NO, CUST_NAME, INV_DATE_SHIPPED" & vbCrLf & sql_cols & ") " & ASCMAIN1.sql
                    Else
                        ASCDATA1.ExecuteSQL("Truncate Table " & SATSIST2)
                        sqlSATSIST2 = "Insert into " & SATSIST2 & " (ITEM_CODE, CUST_CODE, CUST_STORE_NO, INV_DATE_SHIPPED" & vbCrLf & sql_cols & ") " & ASCMAIN1.sql
                    End If

                    Fill_and_Cum(SATSIST2, "SATSIST2")
                End If
            Next

            dst.Tables(TABLE_NAME).Columns("SIQTY_P00").Expression = "ISNULL(SIQTY_P" & Format(Periods, "00") & ",0)"
        Next

        dst.Tables("SATSIST4").Columns("SIQTY_P00").Expression = "ISNULL(SIQTY_P" & Format(Periods, "00") & ",0)"

        Dim dvw As DataView = DirectCast(grdSATSIST3.DataSource, DataTable).DefaultView
        dvw.RowFilter = "ITEM_CODE = '" & ITEM_CODE_LY & "'"

        Sort_grdColumns(grdSATSIST1, "CUST_CODE")
        Setup_grdSATSIST2()
        Sort_grdColumns(grdSATSIST3, "CUST_CODE")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub grdSATSIST1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSATSIST1.AfterRowActivate
        If chkSIST_Multi.Checked Then
        Else
            Setup_grdSATSIST2()
        End If

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
        If grdSATSIST1.ActiveRow Is Nothing Then
            grdSATSIST2.Visible = False
            chkConsolidateLYs.Visible = False
        Else
            Dim CUST_CODE As String = grdSATSIST1.ActiveRow.Cells("CUST_CODE").Value

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Loading Store Level Data")

            ASCDATA1.ExecuteSQL("Truncate Table " & SATSIST2)
            ASCDATA1.ExecuteSQL(Replace(sqlSATSIST2, "'CUST_CODE'", "'" & CUST_CODE & "'"))

            'Dim DVW As DataView = DirectCast(grdSATSIST2.DataSource, DataTable).DefaultView
            'DVW.RowFilter = "CUST_CODE = '" & CUST_CODE & "'"

            Fill_and_Cum(SATSIST2, "SATSIST2")

            grdSATSIST2.Visible = True
            chkConsolidateLYs.Visible = True
            grdSATSIST2.Text = "Activity by Store for " & CUST_CODE
            'Sort_grdColumns(grdSATSIST2, "CUST_STORE_NO")
            Sort_grdColumns(grdSATSIST2, "CUST_NAME")

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")

        End If
    End Sub

    Sub Fill_and_Cum(TABLE_NAME_ora As String, TABLE_NAME_ado As String)

        EnforceConstraints(False)

        If CODEs2EXCLUDE <> "" Then
            ASCMAIN1.sql = "Delete from " & TABLE_NAME_ora & " where CUST_CODE in (" & CODEs2EXCLUDE & ")"
            ASCDATA1.ExecuteSQL()
        End If


        ASCMAIN1.sql = "Select * from " & TABLE_NAME_ora
        Fill_Records(TABLE_NAME_ado, "", True, ASCMAIN1.sql)
        For Each row As DataRow In dst.Tables(TABLE_NAME_ado).Select("")
            'Dim CUST_CODE As String = row.Item("CUST_CODE")
            For P As Integer = 2 To Periods
                row.Item("SIQTY_P" & Format(P, "00")) = Val(row.Item("SIQTY_P" & Format(P, "00")) & "") + Val(row.Item("SIQTY_P" & Format(P - 1, "00")) & "")
            Next
        Next
        EnforceConstraints(True)
    End Sub
    Private Sub grdSATISLS1_AfterColRegionScroll(sender As Object, e As UltraWinGrid.ColScrollRegionEventArgs) Handles grdSATISLS1.AfterColRegionScroll
        grdSATISLS2.DisplayLayout.ColScrollRegions(0).Position = e.ColScrollRegion.Position
    End Sub

    Private Sub grdSATISLS2_AfterColRegionScroll(sender As Object, e As UltraWinGrid.ColScrollRegionEventArgs) Handles grdSATISLS2.AfterColRegionScroll
        grdSATISLS1.DisplayLayout.ColScrollRegions(0).Position = e.ColScrollRegion.Position
    End Sub

    Public Overrides Function CustomSummary_End(
        ByVal summarySettings As UltraWinGrid.SummarySettings,
        ByVal rows As UltraWinGrid.RowsCollection,
        ByVal CustomValue As Double,
        ByVal grd As UltraWinGrid.UltraGrid) As Double

        CustomValue = 0
        Dim TOTALS As New Dictionary(Of String, Decimal)

        Select Case grd.Name
            Case "grdSATSIST1", "grdSATSIST2", "grdSATSIST3"
                Dim KEY As String = summarySettings.Key
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
            Case "grdSATSIST1"
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
                If KEY.StartsWith("STPCT") Then
                    Dim SI As String = "SIQTY_P" & Mid(KEY, 8)
                    Dim ST As String = "STQTY_P" & Mid(KEY, 8)
                    TOTALS(SI) += Val(grow2.Cells(SI).Value & "")
                    TOTALS(ST) += Val(grow2.Cells(ST).Value & "")
                ElseIf KEY = "TRADE_CLASS_CODE" Then
                    '  TOTALS(KEY) = "Totals"
                End If
            End If
        Next
    End Sub


    Private Sub grdSPTCOOPX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSPTCOOPX.InitializeRow
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

    Function Excel_Extract_Headings_Periods(
        worksheet As SpreadsheetGear.IWorksheet,
        LY As String,
        ITEM_CODE_THIS_HEADING As String,
        RxG As Integer,
        Rx As Integer,
        handlePromoCells As Boolean,
        handleDataHeadings As Boolean, Optional isSummary As Boolean = False) As Integer

        Dim CxD As Integer = 0

        CxD = 0
        If handleDataHeadings Then
            With worksheet.Cells(Rx + 1, 0 + CxD)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Left
                .ColumnWidth = 12
                If handlePromoCells Then .Value = "Customer"
            End With
        End If

        CxD += 1
        If handleDataHeadings Then
            With worksheet.Cells(Rx + 1, 0 + CxD)
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
                .ColumnWidth = 30
                If handlePromoCells Then .Value = "Name"
            End With
        End If

        CxD += 1
        If handleDataHeadings Then
            With worksheet.Cells(Rx + 1, 0 + CxD)
                .EntireColumn.NumberFormat = "@"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
                .ColumnWidth = 8
                .Columns.Hidden = True
                If handlePromoCells Then .Value = "SRep"
            End With
        End If

        CxD += 1
        If handleDataHeadings Then
            With worksheet.Cells(Rx + 1, 0 + CxD)
                .EntireColumn.NumberFormat = "###.00"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .ColumnWidth = 8
                If handlePromoCells Then .Value = "Disc"
                .Columns.Hidden = True
            End With
        End If

        CxD += 1
        If handleDataHeadings Then
            With worksheet.Cells(Rx + 1, 0 + CxD)
                .EntireColumn.NumberFormat = "MM/DD/YY"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
                .ColumnWidth = 12
                If handlePromoCells Then .Value = "Shipped"
                If isSummary Then
                    .Columns.Hidden = True
                End If
            End With '
        End If

        CxD += 1
        If handleDataHeadings Then
            With worksheet.Cells(Rx + 1, 0 + CxD)
                .EntireColumn.NumberFormat = "#,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .ColumnWidth = 8
                .Value = "Alloc"
                .Columns.Hidden = False
            End With
        End If

        CxD += 1
        If handleDataHeadings Then
            With worksheet.Cells(Rx + 1, 0 + CxD)
                .EntireColumn.NumberFormat = "#,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .ColumnWidth = 8
                .Value = "OH EOW"
                .Columns.Hidden = False
            End With
        End If


        For P As Integer = 0 To Periods
            With worksheet.Cells(Rx, CxD + P * 5 + 1, Rx, CxD + P * 5 + 5)
                .Merge()
                .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                If P <> 0 And (P + 4 <= Periods) Then
                    .Columns.Hidden = True
                End If

                If handlePromoCells Then
                    ' Merge and italicize the Comments block 1 row above period headings
                    worksheet.Cells(RxG - 1, CxD + P * 5 + 1, RxG - 1, CxD + P * 5 + 5).Merge()
                    worksheet.Cells(RxG - 1, CxD + P * 5 + 1, RxG - 1, CxD + P * 5 + 5).Font.Italic = True
                End If


                If P = 0 Then
                    .Value = "Totals"
                Else
                    If LY = "2" Then
                        .Value = Ls2LY(P - 1)
                    ElseIf LY = "1" Then
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

            worksheet.Cells(Rx, 0).Value = IIf(LY = "0", "This Year", IIf(LY = "1", "Last Year", "2 Years Ago") & IIf(ITEM_CODE_THIS_HEADING = ITEM_CODE Or ITEM_CODE_THIS_HEADING = "", "", " (" & ITEM_CODE_THIS_HEADING & ")"))

            If handleDataHeadings Then
                Excel_Extract_Headings_Data(worksheet, Rx, CxD, P)
            End If

        Next

        Return CxD
    End Function

    Sub Excel_Extract_Headings_Data(worksheet As SpreadsheetGear.IWorksheet, Rx As Integer, CxD As Integer, p As Integer)
        With worksheet.Cells(Rx + 1, 0, Rx + 1, CxD + Periods * 5 + 5)
            .Interior.Color = SpreadsheetGear.Colors.LightBlue
            .Borders.LineStyle = SpreadsheetGear.LineStyle.Continuous
        End With

        With worksheet.Cells(Rx + 1, CxD + p * 5 + 1)
            .EntireColumn.NumberFormat = "#,##0"
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            .Value = "Sell-In"
        End With
        With worksheet.Cells(Rx + 1, CxD + p * 5 + 2)
            .EntireColumn.NumberFormat = "#,##0"
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            .Value = "Sell-Thru"
        End With
        With worksheet.Cells(Rx + 1, CxD + p * 5 + 3)
            .EntireColumn.NumberFormat = "#,##0"
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            .Value = "On Hand"
        End With
        With worksheet.Cells(Rx + 1, CxD + p * 5 + 4)
            .EntireColumn.NumberFormat = "#,##0.0"
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            .Value = "SIST%"
        End With
        With worksheet.Cells(Rx + 1, CxD + p * 5 + 5)
            .EntireColumn.NumberFormat = "#,##0"
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            .Value = "$Thru"
        End With
    End Sub

    Sub Excel_Extract_Summary()

        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(XLS_FILENAME_multi)

        Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing
        worksheet = workbook.Worksheets.AddBefore(workbook.Worksheets(0))
        worksheet.Name = "Summary"

        Dim Rx As Integer = 0

        worksheet.Cells(Rx, 0).Value = "'" & Format(Now, "MM/dd/yyyy")
        worksheet.Cells(Rx, 1).Value = ASCMAIN1.USER_ID

        Rx += 1
        With worksheet.Cells(Rx, 0)
            .RowHeight = 20
            .Font.Bold = True
            .Font.Size = 16
            .Value = XLS_TITLE
        End With
        Rx += 1
        worksheet.Cells(Rx, 0).Value = XLS_DATE_RANGE


        Rx += 2

        With dst.Tables("SATSISTU")

            Dim T_YEAR As New List(Of Integer)
            Dim T_HC_CODE As New List(Of Integer)
            Dim T_COLLECTION_CODE As New List(Of Integer)

            For Each LY As String In New String() {"0", "1", "2"}

                Dim sqlW As String = "YEAR = '" & LY & "'"
                If LY = "2" And .Select(sqlW).Length = 0 Then Exit For

                Rx += 3
                Dim RxG As Integer = Rx
                ' Dim YR As Integer = Val(LY)
                ' PRINT HEADINGS
                Dim CxD As Integer = Excel_Extract_Headings_Periods(worksheet, LY, "", RxG, Rx, False, True, True)
                worksheet.Cells(Rx + 1, 0).Value = "Picture"
                worksheet.Cells(Rx + 1, 1).Value = "Item"
                worksheet.Cells(Rx + 1, CxD - 1).Value = "Alloc"
                worksheet.Cells(Rx + 1, CxD - 0).Value = "OH EOW"
                Rx += 1
                worksheet.Cells(Rx + 1, 0).EntireRow.RowHeight = 5
                Rx += 1

                Dim firstCollectionWithinPeriod As Boolean = True

                For Each rowHC_CODE As DataRow In ASCDATA1.SelectDistinct(.Select(sqlW), New String() {"HC_CODE"}).Select("")

                    Dim HC_CODE As String = rowHC_CODE.Item(0)
                    Dim rowICTCOLL0 As DataRow = LookUp("ICTCOLL0", HC_CODE)
                    Dim sqlw2 As String = sqlW & " and HC_CODE = '" & HC_CODE & "'"

                    For Each rowCOLLECTION_CODE As DataRow In ASCDATA1.SelectDistinct(.Select(sqlw2), New String() {"COLLECTION_CODE"}).Select("")
                        Dim COLLECTION_CODE As String = rowCOLLECTION_CODE.Item(0)
                        Dim rowICTCOLL1 As DataRow = LookUp("ICTCOLL1", COLLECTION_CODE)
                        Dim sqlw3 As String = sqlW & " and COLLECTION_CODE = '" & COLLECTION_CODE & "'"

                        If firstCollectionWithinPeriod Then
                            firstCollectionWithinPeriod = False
                        Else
                            Rx += 1
                            For P As Integer = 0 To Periods
                                Excel_Extract_Headings_Data(worksheet, Rx, CxD, P)
                            Next
                            worksheet.Cells(Rx + 1, 0).Value = "Picture"
                            worksheet.Cells(Rx + 1, 1).Value = "Item"
                            worksheet.Cells(Rx + 1, CxD - 1).Value = "Alloc"
                            worksheet.Cells(Rx + 1, CxD - 0).Value = "OH EOW"
                            Rx += 1
                            worksheet.Cells(Rx + 1, 0).EntireRow.RowHeight = 5
                            Rx += 1
                        End If

                        For Each rowSATSISTU As DataRow In .Select(sqlw3)
                            Rx += 1

                            T_COLLECTION_CODE.Add(Rx)

                            With worksheet.Cells(Rx, 2).EntireRow
                                .RowHeight = 50
                                .VerticalAlignment = SpreadsheetGear.VAlign.Center
                            End With


                            Dim ITEM_CODE As String = rowSATSISTU.Item("ITEM_CODE")
                            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                            '   worksheet.Cells(Rx, 2).Value = ITEM_CODE
                            worksheet.Cells(Rx, 1).Value = ITEM_CODE & vbCrLf & rowICTITEM1.Item("ITEM_DESC")

                            Dim IMAGE_NAME As String = ITEM_CODE
                            Dim IMAGE_FOLDER As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
                            Dim imageFileStyle As String = IMAGE_FOLDER & "\" & IMAGE_NAME & ".jpg"

                            If My.Computer.FileSystem.FileExists(imageFileStyle) Then
                                Dim widthStyle As Double
                                Dim heightStyle As Double

                                Dim imageStyle As System.Drawing.Image = System.Drawing.Image.FromFile(imageFileStyle)
                                Try
                                    widthStyle = imageStyle.Width * 72.0 / imageStyle.HorizontalResolution * 0.25
                                    heightStyle = imageStyle.Height * 72.0 / imageStyle.VerticalResolution * 0.25
                                    heightStyle = 50
                                    widthStyle = 65
                                Finally
                                    imageStyle.Dispose()
                                End Try


                                ' Calculate the left and top placement of the picture by converting 
                                ' row and column coordinates to points.  Use fractional values to 
                                ' get coordinates anywhere in between row and column boundaries.
                                Dim windowInfoStyle As SpreadsheetGear.IWorksheetWindowInfo = worksheet.WindowInfo
                                Dim leftStyle As Double = windowInfoStyle.ColumnToPoints(0)
                                Dim topStyle As Double = windowInfoStyle.RowToPoints(Rx)
                                '  Dim topStyle2 As Double = windowInfoStyle.RowToPoints(Rx + 1)

                                ' Add the picture from file.
                                worksheet.Shapes.AddPicture(imageFileStyle, leftStyle, topStyle, widthStyle, heightStyle)
                            End If

                            Dim SHEET_NAME As String = rowSATSISTU.Item("SHEET_NAME") & ""
                            Dim ROW_NO As Int32 = Val(rowSATSISTU.Item("ROW_NO") & "")

                            For P As Integer = 0 To Periods
                                Dim Cx0 As Integer = CxD + P * 5
                                For Cx As Integer = Cx0 + 1 To Cx0 + 5
                                    If P = 0 And Cx = Cx0 + 1 Then
                                        worksheet.Cells(Rx, Cx - 2).Formula = "='" & SHEET_NAME & "'!" & Excel_Cell0(ROW_NO, Cx - 2)
                                        worksheet.Cells(Rx, Cx - 1).Formula = "='" & SHEET_NAME & "'!" & Excel_Cell0(ROW_NO, Cx - 1)
                                    End If
                                    If Cx = Cx0 + 4 Then
                                        ' Dim RxCx As String = Excel_Cell0(ROW_NO, Cx)
                                        ' worksheet.Cells(Rx, Cx).Copy(workbook.Worksheets(SHEET_NAME).Range(RxCx & ":" & RxCx))
                                        ' workbook.Worksheets(SHEET_NAME).Cells(RxCx).Copy(worksheet.Cells(Rx, Cx))
                                        worksheet.Cells(Rx, Cx).Formula = String.Format("=IF({0}=0,0,100*{1}/{0})", Excel_Cell0(Rx, Cx - 3), Excel_Cell0(Rx, Cx - 2))
                                    Else
                                        worksheet.Cells(Rx, Cx).Formula = "='" & SHEET_NAME & "'!" & Excel_Cell0(ROW_NO, Cx)
                                    End If
                                Next
                            Next
                        Next

                        Rx += 1
                        'COLLECTION TOTALS
                        worksheet.Cells(Rx, 0).Value = "Collection"
                        worksheet.Cells(Rx, 1).Value = "Totals" & " " & COLLECTION_CODE & " " & rowICTCOLL1.Item("COLLECTION_NAME")
                        With worksheet.Cells(Rx, 0, Rx, CxD + Periods * 5 + 5)
                            .Borders.LineStyle = SpreadsheetGear.LineStyle.Continuous
                            .Interior.Color = SpreadsheetGear.Colors.LightGray
                        End With

                        Excel_Extract_Totals(worksheet, CxD, T_COLLECTION_CODE, Rx)
                        T_COLLECTION_CODE.Clear()
                        T_HC_CODE.Add(Rx)


                    Next
                    Rx += 1
                    ' HC TOTALS
                    worksheet.Cells(Rx, 0).Value = "High-Coll"
                    'rowICTCOLL0 = LookUp("ICTCOLL0", HC_CODE)
                    worksheet.Cells(Rx, 1).Value = "Totals" & " " & HC_CODE & " " & rowICTCOLL0.Item("HC_NAME")
                    With worksheet.Cells(Rx, 0, Rx, CxD + Periods * 5 + 5)
                        .Borders.LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Interior.Color = SpreadsheetGear.Colors.DarkGray
                        .Font.Color = SpreadsheetGear.Colors.White
                    End With

                    Excel_Extract_Totals(worksheet, CxD, T_HC_CODE, Rx)
                    T_HC_CODE.Clear()
                    T_YEAR.Add(Rx)

                Next
                Rx += 1
                ' YEAR TOTALS
                worksheet.Cells(Rx, 0).Value = IIf(LY = "0", "This Year", IIf(LY = "1", "Last Year", "2 Years Ago"))
                worksheet.Cells(Rx, 1).Value = "Totals"
                With worksheet.Cells(Rx, 0, Rx, CxD + Periods * 5 + 5)
                    .Borders.LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Interior.Color = SpreadsheetGear.Colors.Black
                    .Font.Color = SpreadsheetGear.Colors.White
                End With

                Excel_Extract_Totals(worksheet, CxD, T_YEAR, Rx)
                T_YEAR.Clear()

            Next
        End With

        workbook.SaveAs(XLS_FILENAME_multi, SpreadsheetGear.FileFormat.OpenXMLWorkbook)

    End Sub

    Sub Excel_Extract_Totals(worksheet As SpreadsheetGear.IWorksheet, CxD As Integer, T As List(Of Integer), Rx As Integer)
        If T.Count = 0 Then Exit Sub

        For P As Integer = 0 To Periods
            Dim Cx0 As Integer = CxD + P * 5
            For Cx As Integer = Cx0 + 1 To Cx0 + 5

                If P = 0 And Cx = Cx0 + 1 Then
                    Dim F As String = ""

                    F = ""
                    For Each TC As Integer In T
                        Dim DC As String = Excel_Cell0(TC, Cx - 2)
                        F &= "+" & DC
                    Next
                    worksheet.Cells(Rx, Cx - 2).Formula = "=" & Mid(F, 2)

                    F = ""
                    For Each TC As Integer In T
                        Dim DC As String = Excel_Cell0(TC, Cx - 1)
                        F &= "+" & DC
                    Next
                    worksheet.Cells(Rx, Cx - 1).Formula = "=" & Mid(F, 2)
                End If


                If Cx = Cx0 + 4 Then
                    ' Dim RxCx As String = Excel_Cell0(Rx - 1, Cx)
                    ' worksheet.Cells(Rx, Cx).Copy(worksheet.Range(RxCx & ":" & RxCx))
                    ' worksheet.Cells(RxCx).Copy(worksheet.Cells(Rx, Cx))
                    worksheet.Cells(Rx, Cx).Formula = String.Format("=IF({0}=0,0,100*{1}/{0})", Excel_Cell0(Rx, Cx - 3), Excel_Cell0(Rx, Cx - 2))
                Else

                    Dim F As String = ""
                    F = ""
                    For Each TC As Integer In T
                        Dim DC As String = Excel_Cell0(TC, Cx)
                        F &= "+" & DC
                    Next
                    worksheet.Cells(Rx, Cx).Formula = "=" & Mid(F, 2)
                End If
            Next
        Next
    End Sub

    Sub Excel_Extract(
                     Optional ByRef workbook As SpreadsheetGear.IWorkbook = Nothing,
                     Optional ByRef XLS_FILENAME As String = "",
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
                dst.Tables("SATSISTU").Rows.Clear()
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

        Dim ITEM_CODE As String = Me.ITEM_CODE
        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)

        Dim Total_Cols As Integer = 10

        If XLS_FILENAME_multi = "" Or chkSIST_SingleSheet.Checked Then
            worksheet = workbook.Worksheets(0)
            worksheet.Name = "Sell-In Sell-Thru"
        Else
            worksheet.Name = ITEM_CODE
        End If

        If Start_Row = 0 Then
            worksheet.Cells(Start_Row, 0).Value = "'" & Format(Now, "MM/dd/yy")
            worksheet.Cells(Start_Row, 1).Value = ASCMAIN1.USER_ID
        End If
        With worksheet.Cells(Start_Row + 1, 0)
            .RowHeight = 20
            .Font.Bold = True
            .Font.Size = 16
            .Value = "'" & XLS_TITLE
        End With
        worksheet.Cells(Start_Row + 2, 0).Value = XLS_DATE_RANGE


        Dim Rx As Integer = Start_Row + 4

        ' Fill_Records("SATSISTY", ITEM_CODE)
        'Dim ITEM_CODE_XLYs As New Dictionary(Of String, String)
        Dim ITEM_CODE_XLYlist As New List(Of String)
        Dim LYs As New Dictionary(Of String, List(Of String))

        ' ITEM_CODE_XLYs.Add(ITEM_CODE, "0")
        ITEM_CODE_XLYlist.Add(ITEM_CODE & ":" & "0")
        LYs.Add("0", New List(Of String))
        LYs("0").Add(ITEM_CODE)

        For Each rowSATSISTY As DataRow In dst.Tables("SATSISTY").Select($"ITEM_CODE = '{ITEM_CODE}'", "LY")
            Dim ITEM_CODE_COMPARE_TO As String = rowSATSISTY.Item("ITEM_CODE_COMPARE_TO")
            Dim LY As String = rowSATSISTY.Item("LY")
            If Not LYs.ContainsKey(LY) Then LYs.Add(LY, New List(Of String))
            LYs(LY).Add(ITEM_CODE_COMPARE_TO)
            If Not chkConsolidateLYs.Checked Or LYs(LY).Count = 1 Then
                ' add items to ITEM_CODE_XLYs if printing Individual, or if this is the 1st item encountered for that LY
                ' for consolidated, the item loaded into ITEM_CODE_XLYs is only a token item - all items for that LY are in LYs
                'ITEM_CODE_XLYs.Add(ITEM_CODE_COMPARE_TO, LY)
                ITEM_CODE_XLYlist.Add(ITEM_CODE_COMPARE_TO & ":" & LY)
            End If
        Next

        Dim rowICTITEM1s As New List(Of DataRow)

        Dim DTE0 As Date = DTES(0)
        Dim DTE1 As Date = DTES(1)

        For iITEM_CODE As Integer = 0 To ITEM_CODE_XLYlist.Count - 1 '  ITEM_CODE_XLYs.Count - 1

            'Dim LY As String = ITEM_CODE_XLYs(ITEM_CODE)
            Dim ITEM_CODE_XLYlistItem As String = ITEM_CODE_XLYlist(iITEM_CODE)
            Dim LY As String = Split(ITEM_CODE_XLYlistItem, ":")(1)

            Dim sqlITEM_CODEs As String = $"'{ITEM_CODE}'"

            If iITEM_CODE >= 1 Then ' EITHER LY OR 2LY
                'ITEM_CODE = ITEM_CODE_XLYs.Keys(iITEM_CODE)
                'LY = ITEM_CODE_XLYs(ITEM_CODE)

                ITEM_CODE = Split(ITEM_CODE_XLYlistItem, ":")(0)

                rowICTITEM1s.Clear()

                If chkConsolidateLYs.Checked Then
                    sqlITEM_CODEs = ""
                    For Each ITEM_CODE_COMPARE_TO As String In LYs(LY)
                        sqlITEM_CODEs &= $" or ITEM_CODE = '{ITEM_CODE_COMPARE_TO}'"
                        rowICTITEM1s.Add(LookUp("ICTITEM1", ITEM_CODE_COMPARE_TO))
                    Next
                    sqlITEM_CODEs = Mid(sqlITEM_CODEs, 5)
                    If dst.Tables("SATSIST3").Select(sqlITEM_CODEs).Length = 0 Then
                        Exit For
                    End If
                Else
                    If dst.Tables("SATSIST3").Select($"ITEM_CODE = '{ITEM_CODE}'").Length = 0 Then
                        ' Exit For
                        Continue For
                    End If
                End If

                rowICTITEM1 = LookUp("ICTITEM1", ITEM_CODE)
            End If

            'Dim DTE0 As Date = DTES(0)
            'Dim DTE1 As Date = DTES(1)

            If LY = "0" Then

                If chkSIST.Checked Then
                    Dim rowSATSISTI As DataRow = dst.Tables("SATSISTI").Rows.Find(ITEM_CODE)
                    If rowSATSISTI Is Nothing Then
                    Else
                        If rowSATSISTI.Item("ALLO_START") & "" <> "" Then
                            DTE0 = rowSATSISTI.Item("ALLO_START")
                        End If
                        If rowSATSISTI.Item("ALLO_END") & "" <> "" Then
                            DTE1 = rowSATSISTI.Item("ALLO_END")
                        End If
                    End If
                End If
            End If

            If LY <> "0" Then
                DTE0 = DTE0.AddYears(-1 * Val(LY))
                DTE1 = DTE1.AddYears(-1 * Val(LY))
            End If

            If LY <> "0" And chkConsolidateLYs.Checked Then
                dst.Tables("SOTALLO1").Rows.Clear()
                dst.Tables("SOTALLO2").Rows.Clear()
                For Each ITEM_CODE_COMPARE_TO As String In LYs(LY)
                    Fill_Records("SOTALLO1", New Object() {ITEM_CODE_COMPARE_TO, DTE1, DTE0}, False)
                    Fill_Records("SOTALLO2", New Object() {ITEM_CODE_COMPARE_TO, DTE1, DTE0}, False)
                Next
            Else
                Fill_Records("SOTALLO1", New Object() {ITEM_CODE, DTE1, DTE0})
                Fill_Records("SOTALLO2", New Object() {ITEM_CODE, DTE1, DTE0})
            End If

            With worksheet.Cells(Rx, 0, Rx + 3, 5)
                If LY = "0" Then
                    .Interior.Color = SpreadsheetGear.Colors.WhiteSmoke
                ElseIf LY = "1" Then
                    .Interior.Color = SpreadsheetGear.Colors.LightGreen
                ElseIf LY = "2" Then
                    .Interior.Color = SpreadsheetGear.Colors.Yellow
                End If

                .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
                .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
            End With

            With worksheet.Cells(Rx, 0)
                .Value = "'" & IIf(LY <> "0" And chkConsolidateLYs.Checked, Get_ICTITEM1_Values("ITEM_CODE", rowICTITEM1s), ITEM_CODE)
                .Font.Underline = SpreadsheetGear.UnderlineStyle.Single
                .Font.Size = 14
            End With
            With worksheet.Cells(Rx + 1, 0)
                .Value = "'" & IIf(LY <> "0" And chkConsolidateLYs.Checked, Get_ICTITEM1_Values("ITEM_DESC", rowICTITEM1s), rowICTITEM1.Item("ITEM_DESC"))
                .Font.Color = SpreadsheetGear.Colors.Purple
                .Font.Name = "Times New Roman"
                .Font.Size = 20
            End With
            With worksheet.Cells(Rx + 2, 0)

                .Value = "Retail: " & IIf(LY <> "0" And chkConsolidateLYs.Checked,
                                          Get_ICTITEM1_Values("ITEM_RETAIL_PRICE", rowICTITEM1s),
                                          Format(Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & ""), "$#.00"))
                .Font.Size = 12
            End With


            With worksheet.Cells(Rx + 3, 0)
                .Value = "Plan " & Format(QTY_ALLO_PLAN, "#,##0")
                .Font.Size = 12
            End With

            If LY <> "0" And chkConsolidateLYs.Checked Then

                With worksheet.Cells(Rx + 4, 0)
                    .Value = "Ship Date " & Get_ICTITEM1_Values("ITEM_DATE_TO_SHIP", rowICTITEM1s)
                    .Font.Size = 12
                End With
            Else
                If rowICTITEM1.Item("ITEM_DATE_TO_SHIP") & "" <> "" Then
                    With worksheet.Cells(Rx + 4, 0)
                        .Value = "Ship Date " & Format(rowICTITEM1.Item("ITEM_DATE_TO_SHIP"), "MM/dd/yy")
                        .Font.Size = 12
                    End With
                End If
            End If

            If LY <> "0" And chkConsolidateLYs.Checked Then
                Dim CX As Integer = 7
                For Each ITEM_CODE_COMPARE_TO As String In LYs(LY)
                    Add_Image_to_Worksheet(worksheet, ITEM_CODE_COMPARE_TO, CX, Rx)
                    CX += 3
                Next
            Else
                Add_Image_to_Worksheet(worksheet, ITEM_CODE, 7, Rx)
            End If


            Rx += 5


            Dim RxG As Integer = Rx

            Dim ITEM_CODE_THIS_HEADING As String = ITEM_CODE
            If LY <> "0" And chkConsolidateLYs.Checked Then
                ITEM_CODE_THIS_HEADING = Join(LYs(LY).ToArray, ",")
            End If
            Dim CxD As Integer = Excel_Extract_Headings_Periods(worksheet, LY, ITEM_CODE_THIS_HEADING, RxG, Rx, True, True)

            Rx += 2
            Dim iRx As Integer = Rx

            Dim sqlITEM As String = "ITEM_CODE = '" & ITEM_CODE & "'"

            Dim SIST As String = "SATSIST1"
            If LY <> "0" Then SIST = "SATSIST3"

            If LY <> "0" And chkConsolidateLYs.Checked Then
                dst.Tables("SATSIST4").Rows.Clear()

                For Each rowSATSIST3 As DataRow In dst.Tables("SATSIST3").Select(sqlITEM_CODEs, "CUST_CODE")
                    Dim CUST_CODE As String = rowSATSIST3.Item("CUST_CODE")
                    Dim rowSATSIST4 As DataRow = dst.Tables("SATSIST4").Rows.Find(New String() {"*", CUST_CODE})
                    If rowSATSIST4 Is Nothing Then
                        rowSATSIST4 = dst.Tables("SATSIST4").NewRow
                        With rowSATSIST4
                            .Item("ITEM_CODE") = "*"
                            .Item("CUST_CODE") = CUST_CODE
                            .Item("INV_DATE_SHIPPED") = rowSATSIST3.Item("INV_DATE_SHIPPED")
                        End With
                        dst.Tables("SATSIST4").Rows.Add(rowSATSIST4)
                    End If

                    If rowSATSIST3.Item("INV_DATE_SHIPPED") & "" <> "" Then
                        If rowSATSIST4.Item("INV_DATE_SHIPPED") & "" = "" Then
                            rowSATSIST4.Item("INV_DATE_SHIPPED") = rowSATSIST3.Item("INV_DATE_SHIPPED")
                        Else
                            If Format(rowSATSIST4.Item("INV_DATE_SHIPPED"), "yyyyMMdd") > Format(rowSATSIST3.Item("INV_DATE_SHIPPED"), "yyyyMMdd") Then
                                rowSATSIST4.Item("INV_DATE_SHIPPED") = rowSATSIST3.Item("INV_DATE_SHIPPED")
                            End If
                        End If
                    End If

                    For P As Integer = 1 To Periods
                        Dim P00 As String = Format(P, "00")
                        With rowSATSIST4
                            For Each C As String In New String() {"SIQTY_P", "STQTY_P", "OHQTY_P", "STAMT_P"}
                                .Item(C & P00) = Val(.Item(C & P00) & "") + Val(rowSATSIST3.Item(C & P00) & "")
                            Next
                        End With
                    Next
                    With rowSATSIST4
                        .Item("EOW") = Val(.Item("EOW") & "") + Val(rowSATSIST3.Item("EOW") & "")
                    End With
                Next

                SIST = "SATSIST4"
                sqlITEM = "ITEM_CODE = '*'"
            End If

            If dst.Tables(SIST).Select(sqlITEM).Length > 0 Then
                For Each row As DataRow In dst.Tables(SIST).Select(sqlITEM, "CUST_CODE")
                    Dim CUST_CODE As String = row.Item("CUST_CODE")
                    Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE, True)
                    Dim rowSOTPCLS1 As DataRow = LookUp("SOTPCLS1", rowARTCUST1.Item("PRICE_CLASS_CODE") & "", True)

                    Dim CxDi As Integer = 0
                    worksheet.Cells(Rx, 0 + CxDi).Value = CUST_CODE
                    CxDi += 1 : worksheet.Cells(Rx, 0 + CxDi).Value = rowARTCUST1.Item("CUST_NAME")
                    CxDi += 1 : worksheet.Cells(Rx, 0 + CxDi).Value = rowARTCUST1.Item("SREP_CODE")
                    CxDi += 1 : worksheet.Cells(Rx, 0 + CxDi).Value = Val(rowSOTPCLS1.Item("PRICE_BASE_DPCT") & "")
                    CxDi += 1 : worksheet.Cells(Rx, 0 + CxDi).Value = row.Item("INV_DATE_SHIPPED")

                    Dim QTY_ALLO As Int32 = Val(dst.Tables("SOTALLO2").Compute("SUM(QTY_ALLO)", "CUST_CODE = '" & CUST_CODE & "'") & "")
                    CxDi += 1 : worksheet.Cells(Rx, 0 + CxDi).Value = QTY_ALLO
                    CxDi += 1 : worksheet.Cells(Rx, 0 + CxDi).Value = row.Item("EOW")

                    For P As Integer = 0 To Periods
                        worksheet.Cells(Rx, CxD + P * 5 + 1).Formula = row.Item("SIQTY_P" & Format(P, "00")) & ""
                        worksheet.Cells(Rx, CxD + P * 5 + 2).Formula = row.Item("STQTY_P" & Format(P, "00")) & ""
                        worksheet.Cells(Rx, CxD + P * 5 + 3).Formula = row.Item("OHQTY_P" & Format(P, "00")) & ""
                        worksheet.Cells(Rx, CxD + P * 5 + 4).Formula = row.Item("STPCT_P" & Format(P, "00")) & ""
                        worksheet.Cells(Rx, CxD + P * 5 + 5).Formula = row.Item("STAMT_P" & Format(P, "00")) & ""
                    Next

                    Dim sqlw As String = ""
                    If optRANGE.Value = "P" Then
                        If LY = "0" Then
                            sqlw = " and OPS_YYYYPP >= '" & Ps(0) & "' and OPS_YYYYPP <='" & Ps(Periods - 1) & "'"
                        ElseIf LY = "1" Then
                            sqlw = " and OPS_YYYYPP >= '" & PsLY(0) & "' and OPS_YYYYPP <='" & PsLY(Periods - 1) & "'"
                        ElseIf LY = "2" Then
                            sqlw = " and OPS_YYYYPP >= '" & Ps2LY(0) & "' and OPS_YYYYPP <='" & Ps2LY(Periods - 1) & "'"
                        End If
                    Else
                        If LY = "0" Then
                            sqlw = " and OPS_YYYYWW >= '" & Ps(0) & "' and OPS_YYYYWW <='" & Ps(Periods - 1) & "'"
                        ElseIf LY = "1" Then
                            sqlw = " and OPS_YYYYWW >= '" & PsLY(0) & "' and OPS_YYYYWW <='" & PsLY(Periods - 1) & "'"
                        ElseIf LY = "2" Then
                            sqlw = " and OPS_YYYYWW >= '" & Ps2LY(0) & "' and OPS_YYYYWW <='" & Ps2LY(Periods - 1) & "'"
                        End If
                    End If

                    Dim rows() As DataRow = dst.Tables("SPTCOOPX").Select("CUST_CODE='" & CUST_CODE & "'" & sqlw)
                    For Each rowSP As DataRow In rows
                        Dim DATE_START As Date = rowSP.Item("DATE_START")
                        Dim DATE_END As Date = rowSP.Item("DATE_END")
                        Dim EVENT_TYPE_CODE As String = rowSP.Item("EVENT_TYPE_CODE") & ""
                        For P As Integer = 1 To Periods
                            Dim PDT1 As Date = IIf(LY = "0", PE_DATEs(P - 1), PE_DATEsLY(P - 1))
                            Dim PDT2 As Date = IIf(LY = "0", PE_DATEs(P - 0), PE_DATEsLY(P - 0))
                            If Format(DATE_START, "yyyyMMdd") > Format(PDT1, "yyyyMMdd") And
                               Format(DATE_START, "yyyyMMdd") <= Format(PDT2, "yyyyMMdd") Then

                                Dim T As String = Format(DATE_START, "MM/dd") & " " & CUST_CODE & " " & rowSP.Item("BOOKING_NAME") ' & " " & rowSP.Item("VEHICLE_CODE")
                                If ASCMAIN1.CLIENT = "AHA" Then
                                    T = Format(DATE_START, "MM/dd") & "-" & Format(DATE_END, "MM/dd") & " " & EVENT_TYPE_CODE & " " & CUST_CODE & " " & rowSP.Item("BOOKING_NAME") & " " & rowSP.Item("FEATURE_DESC")
                                End If
                                If ASCMAIN1.CLIENT = "AHA" Then
                                Else
                                    With worksheet.Cells(Rx, CxD + P * 5 + 2)
                                        .Interior.Color = SpreadsheetGear.Colors.Violet
                                    End With
                                End If


                                With worksheet.Cells(RxG - 1, CxD + P * 5 + 1)
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


                'If XLS_FILENAME_multi And chkSIST.Checked Then
                '    ' ADD A ROW
                'End If

                For P As Integer = 0 To Periods
                    Dim Cx As Integer = CxD + P * 5

                    If P = 0 Then
                        worksheet.Cells(Rx, Cx - 1).Formula = "=SUM(" & Excel_Cell0(iRx, Cx - 1) & ":" & Excel_Cell0(Rx - 1, Cx - 1) & ")"
                        worksheet.Cells(Rx, Cx - 0).Formula = "=SUM(" & Excel_Cell0(iRx, Cx - 0) & ":" & Excel_Cell0(Rx - 1, Cx - 0) & ")"
                    End If

                    With worksheet.Cells(iRx, Cx + 5, Rx, Cx + 5)
                        .Interior.Color = SpreadsheetGear.Colors.LightGreen
                    End With

                    With worksheet.Cells(iRx, Cx + 1, Rx, Cx + 5)
                        .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    End With

                    Cx += 1 : worksheet.Cells(Rx, Cx).Formula = "=SUM(" & Excel_Cell0(iRx, Cx) & ":" & Excel_Cell0(Rx - 1, Cx) & ")"
                    Cx += 1 : worksheet.Cells(Rx, Cx).Formula = "=SUM(" & Excel_Cell0(iRx, Cx) & ":" & Excel_Cell0(Rx - 1, Cx) & ")"
                    Cx += 1 : worksheet.Cells(Rx, Cx).Formula = "=SUM(" & Excel_Cell0(iRx, Cx) & ":" & Excel_Cell0(Rx - 1, Cx) & ")"
                    Cx += 1 : worksheet.Cells(Rx, Cx).Formula = String.Format("=IF({0}=0,0,100*{1}/{0})", Excel_Cell0(Rx, Cx - 3), Excel_Cell0(Rx, Cx - 2))
                    Cx += 1 : worksheet.Cells(Rx, Cx).Formula = "=SUM(" & Excel_Cell0(iRx, Cx) & ":" & Excel_Cell0(Rx - 1, Cx) & ")"
                Next
            End If

            With worksheet.Cells(Rx, 0, Rx, CxD + Periods * 5 + 5)
                .Borders.LineStyle = SpreadsheetGear.LineStyle.Continuous
                .Interior.Color = SpreadsheetGear.Colors.LightGray
            End With
            worksheet.Cells(Rx, 0).Value = "Totals"

            If XLS_FILENAME_multi <> "" Then
                Dim rowSATSISTU As DataRow
                rowSATSISTU = dst.Tables("SATSISTU").Rows.Find(New String() {LY, ITEM_CODE})
                If rowSATSISTU Is Nothing Then

                    rowSATSISTU = dst.Tables("SATSISTU").NewRow
                    With rowSATSISTU
                        .Item("YEAR") = LY
                        .Item("ITEM_CODE") = ITEM_CODE
                        Dim COLLECTION_CODE As String = rowICTITEM1.Item("COLLECTION_CODE")
                        .Item("COLLECTION_CODE") = COLLECTION_CODE
                        Dim rowICTCOLL1 As DataRow = LookUp("ICTCOLL1", COLLECTION_CODE)
                        .Item("HC_CODE") = rowICTCOLL1.Item("HC_CODE")
                        .Item("BRAND_CODE") = rowICTCOLL1.Item("BRAND_CODE")
                        .Item("SHEET_NO") = worksheet.Index
                        .Item("SHEET_NAME") = worksheet.Name
                        .Item("ROW_NO") = Rx
                    End With

                    dst.Tables("SATSISTU").Rows.Add(rowSATSISTU)
                End If
            End If

            worksheet.Cells(RxG - 1, 0).Rows.AutoFit()
            worksheet.Cells(RxG - 1, 0).Rows.RowHeight = 60

            Rx += 3
        Next

        Start_Row = Rx

        workbook.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        If XLS_FILENAME_multi = "" Then
            ' Show_Document(XLS_FILENAME)
            Add_Document_to_ASTSPRF1(XLS_FILENAME)
        End If

    End Sub

    Sub Add_Image_to_Worksheet(worksheet As SpreadsheetGear.IWorksheet, ITEM_CODE As String, cx As Integer, rx As Integer)

        Dim IMAGE_NAME As String = ITEM_CODE
        Dim IMAGE_FOLDER As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
        Dim imageFileStyle As String = IMAGE_FOLDER & "\" & IMAGE_NAME & ".jpg"
        If Not My.Computer.FileSystem.FileExists(imageFileStyle) Then
            imageFileStyle = IMAGE_FOLDER & "\" & IMAGE_NAME & ".PNG"
        End If

        Dim imageStyle As System.Drawing.Image = Nothing
        If My.Computer.FileSystem.FileExists(imageFileStyle) Then
            Dim widthStyle As Double
            Dim heightStyle As Double

            imageStyle = System.Drawing.Image.FromFile(imageFileStyle)
            Try
                widthStyle = imageStyle.Width * 72.0 / imageStyle.HorizontalResolution * 0.25
                heightStyle = imageStyle.Height * 72.0 / imageStyle.VerticalResolution * 0.25
            Finally
                imageStyle.Dispose()
            End Try

            ' Calculate the left and top placement of the picture by converting 
            ' row and column coordinates to points.  Use fractional values to 
            ' get coordinates anywhere in between row and column boundaries.
            Dim windowInfoStyle As SpreadsheetGear.IWorksheetWindowInfo = worksheet.WindowInfo
            Dim leftStyle As Double = windowInfoStyle.ColumnToPoints(cx)
            Dim topStyle As Double = windowInfoStyle.RowToPoints(rx)

            ' Add the picture from file.
            worksheet.Shapes.AddPicture(imageFileStyle, leftStyle, topStyle, widthStyle, heightStyle)
        End If

    End Sub

    Function Get_ICTITEM1_Values(COLUMN_NAME As String, rowICTITEM1s As List(Of DataRow)) As String
        Dim VALUES As String = ""
        For Each row As DataRow In rowICTITEM1s
            Dim VALUE As String = ""
            If COLUMN_NAME = "ITEM_RETAIL_PRICE" Then
                VALUE = Format(Val(row.Item(COLUMN_NAME) & ""), "$0.00")
            ElseIf COLUMN_NAME = "ITEM_DATE_TO_SHIP" Then
                If row.Item(COLUMN_NAME) & "" <> "" Then
                    VALUE = Format(row.Item(COLUMN_NAME) & "", "MM/dd/yyyy")
                Else
                    VALUE = "00/00/0000"
                End If
            Else
                VALUE = row.Item(COLUMN_NAME) & ""
            End If
            VALUES &= "," & VALUE
        Next
        VALUES = Mid(VALUES, 2)
        Return VALUES
    End Function

    Private Sub grdSATSIST1_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSATSIST1.InitializeRow
        If e.Row.IsDataRow Then
            Dim CUST_CODE As String = e.Row.Cells("CUST_CODE").Value

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
                    If Format(DATE_START, "yyyyMMdd") > Format(PE_DATEs(P - 1), "yyyyMMdd") And
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

    Sub Add_Items()
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
                    Add_Item(ITEM_CODE)
                Next
                grdSATSISTI.PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
            End If

            Sort_grdColumns(grdSATSISTI, "ITEM_CODE")
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
            End If
            Return rowARTCUST1
        End If
    End Function

#Region "grdSATSISTI"
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

                    Dim rowSATSISTZ As DataRow = Fill_Record("SATSISTZ", ITEM_CODE)
                    If rowSATSISTZ IsNot Nothing Then
                        For Each C As String In New String() {"YP_IN1", "YP_IN2", "YP_TH1", "YP_TH2", "YW_IN1", "YW_IN2", "YW_TH1", "YW_TH2"}
                            e.Cell.Row.Cells(C).Value = rowSATSISTZ.Item(C)

                            If rowSATSISTZ(C) & "" <> "" And C.StartsWith("Y" & optRANGE.Value) Then
                                e.Cell.Row.Cells(Replace(C, "Y" & optRANGE.Value, "YX")).Value = rowSATSISTZ.Item(C)
                            End If
                        Next
                    End If
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

                        ASCDATA1.DeleteRows("SATSISTY", $"ITEM_CODE = '{ITEM_CODE}'")
                        Fill_Records("SATSISTY", ITEM_CODE, False)

                        For Each ROW As DataRow In dst.Tables("SATSISTY").Select($"ITEM_CODE = '{ITEM_CODE}'")
                            If optRANGE.Value = "P" Then
                                ROW.Item("YX_IN1") = ROW.Item("YP_IN1")
                                ROW.Item("YX_IN2") = ROW.Item("YP_IN2")
                                ROW.Item("YX_TH1") = ROW.Item("YP_TH1")
                                ROW.Item("YX_TH2") = ROW.Item("YP_TH2")
                            Else
                                ROW.Item("YX_IN1") = ROW.Item("YW_IN1")
                                ROW.Item("YX_IN2") = ROW.Item("YW_IN2")
                                ROW.Item("YX_TH1") = ROW.Item("YW_TH1")
                                ROW.Item("YX_TH2") = ROW.Item("YW_TH2")
                            End If
                        Next

                    End If
            End Select
        End With
    End Sub

    Private Sub grdSATSISTI_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSATSISTI.AfterRowActivate
        Setup_grdSATSISTY()

    End Sub

    Sub Setup_grdSATSISTY()
        If grdSATSISTI.ActiveRow Is Nothing OrElse Not grdSATSISTI.ActiveRow.IsDataRow OrElse grdSATSISTI.ActiveRow.IsFilterRow OrElse grdSATSISTI.ActiveRow.IsAddRow Then
            grdSATSISTY.Visible = False
        Else
            Dim dvw As DataView = DirectCast(grdSATSISTY.DataSource, DataTable).DefaultView
            Dim ITEM_CODE As String = grdSATSISTI.ActiveRow.Cells("ITEM_CODE").Value & ""
            dvw.RowFilter = $"ITEM_CODE = '{ITEM_CODE}'"
            grdSATSISTY.Visible = True
            grdSATSISTY.Text = "LY & 2LY Items for " & ITEM_CODE
        End If
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

#End Region

#Region "grdSATSISTY"

    Private Sub grdSATSISTY_AfterCellListCloseUp(sender As Object, e As CellEventArgs) Handles grdSATSISTY.AfterCellListCloseUp
        grdSATSISTY.ActiveRow.Update()
    End Sub
    Private Sub grdSATSISTY_AfterCellUpdate(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSATSISTY.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "ITEM_CODE_COMPARE_TO"
                Dim ITEM_CODE As String = grdSATSISTI.ActiveRow.Cells("ITEM_CODE").Value & ""
                Dim ITEM_CODE_COMPARE_TO As String = e.Cell.Value & ""
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE_COMPARE_TO)
                If rowICTITEM1 IsNot Nothing Then
                    e.Cell.Row.Cells("ITEM_CODE").Value = ITEM_CODE
                    'e.Cell.Row.Cells("ITEM_CODE_COMPARE_TO").Value = ITEM_CODE_COMPARE_TO
                    e.Cell.Row.Cells("ITEM_DESC").Value = rowICTITEM1.Item("ITEM_DESC")
                    e.Cell.Row.Cells("ITEM_RETAIL_PRICE").Value = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
                    e.Cell.Row.Cells("COLLECTION_CODE").Value = rowICTITEM1.Item("COLLECTION_CODE")
                    e.Cell.Row.Cells("PROD_CODE").Value = rowICTITEM1.Item("PROD_CODE")
                    e.Cell.Row.Cells("ITEM_CATGY_CODE").Value = rowICTITEM1.Item("ITEM_CATGY_CODE")
                    e.Cell.Row.Cells("ITEM_ALT_SORT").Value = rowICTITEM1.Item("ITEM_ALT_SORT")
                End If

        End Select
    End Sub

    Private Sub grdSATSISTY_AfterExitEditMode(sender As Object, e As EventArgs) Handles grdSATSISTY.AfterExitEditMode
        With grdSATSISTY
            Select Case .ActiveCell.Column.Key
                Case "ITEM_CODE_COMPARE_TO"
                    Dim ITEM_CODE As String = .ActiveCell.Text
                    If ITEM_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(ITEM_CODE, .ActiveCell.Column.Key)
                    End If
            End Select
        End With
    End Sub

    Private Sub grdSATSISTY_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSATSISTY.AfterRowActivate

    End Sub

    Private Sub grdSATSISTY_BeforeRowActivate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSATSISTY.BeforeRowActivate
        If e.Row.IsAddRow Then
            grdSATSISTY.DisplayLayout.Bands(0).Columns("ITEM_CODE_COMPARE_TO").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            grdSATSISTY.DisplayLayout.Bands(0).Columns("ITEM_CODE_COMPARE_TO").CellActivation = UltraWinGrid.Activation.NoEdit
        End If
    End Sub

    Private Sub grdSATSISTY_DoubleClickCell(sender As Object, e As DoubleClickCellEventArgs) Handles grdSATSISTY.DoubleClickCell
        If grdSATSISTY.ActiveRow Is Nothing OrElse grdSATSISTY.ActiveRow.IsAddRow OrElse Not grdSATSISTY.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Dim LY As String = grdSATSISTY.ActiveRow.Cells("LY").Value
        Dim VLY As Integer = Val(LY)

        Dim C As String = e.Cell.Column.Key
        If New String() {"YX_IN1", "YX_IN2", "YX_TH1", "YX_TH2"}.Contains(C) Then
            If optRANGE.Value = "W" Then
                Dim RX0 As String = Absx1.cmbFor("RYW0").Value
                Dim RX1 As String = Absx1.cmbFor("RYW1").Value
                RX0 = Format(Val(Mid(RX0, 1, 4)) - VLY, "0000") & Mid(RX0, 5, 2)
                RX1 = Format(Val(Mid(RX1, 1, 4)) - VLY, "0000") & Mid(RX1, 5, 2)
                If C.EndsWith("1") Then
                    e.Cell.Value = RX0
                Else
                    e.Cell.Value = RX1
                End If
            Else
                Dim RX0 As String = Absx1.cmbFor("RYP0").Value
                Dim RX1 As String = Absx1.cmbFor("RYP1").Value
                RX0 = Format(Val(Mid(RX0, 1, 4)) - VLY, "0000") & Mid(RX0, 5, 2)
                RX1 = Format(Val(Mid(RX1, 1, 4)) - VLY, "0000") & Mid(RX1, 5, 2)
                If C.EndsWith("1") Then
                    e.Cell.Value = RX0
                Else
                    e.Cell.Value = RX1
                End If
            End If
        End If
    End Sub

    Private Sub grdSATSISTY_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSATSISTY.InitializeRow
        Dim LY As String = e.Row.Cells("LY").Value & ""
        If LY = "1" Then
            e.Row.Appearance.BackColor = Drawing.Color.LightGreen
        ElseIf LY = "2" Then
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.Appearance.BackColor = Drawing.Color.LightGray
        End If
    End Sub


#End Region
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
        ' splSATSIST0.Visible = Not ScreenMode And chkSIST_Multi.Checked
        splMulti.Visible = Not ScreenMode And chkSIST_Multi.Checked
        lblITEM_CODE.Visible = Not chkSIST_Multi.Checked
        txtITEM_CODE.Visible = Not chkSIST_Multi.Checked
        txtITEM_DESC.Visible = Not chkSIST_Multi.Checked
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

            If dst.Tables("SATSISTI").Rows.Count <> 0 Or dst.Tables("SATSISTC").Rows.Count = 0 Then
                MsgBox("When using Items from Customers" _
                       & vbCrLf & " - A list of Items will be created from the History of the Customers Selected" _
                       & vbCrLf & " - Do NOT specify Items" _
                       & vbCrLf & " - You MUST specify Customers",
                       MsgBoxStyle.OkOnly, "Do Not Specify Items when you use the Items from Customers option")
                Exit Sub
            End If

            Auto_Select()

            If dst.Tables("SATSISTI").Select("").Length = 0 Then
                MsgBox("There were no Items with History in the Customers Selected", MsgBoxStyle.OkOnly, "Cannot Generate XLS")
                Exit Sub
            End If
        End If

        For Each row As DataRow In dst.Tables("SATSISTI").Select("")
            Dim ITEM_CODE_TY As String = row.Item("ITEM_CODE")
            Dim sqlY As String = $"ITEM_CODE = '{ITEM_CODE_TY}' and LY = '1'"
            For Each rowSATSISTY As DataRow In dst.Tables("SATSISTY").Select(sqlY)
                Dim ITEM_CODE_COMPARE_TO As String = rowSATSISTY.Item("ITEM_CODE_COMPARE_TO")
                'If ITEM_CODE_COMPARE_TO = "CH012C15USA" Then Stop
                Dim sqlY2 As String = $"ITEM_CODE = '{ITEM_CODE_TY}' and LY = '2' and ITEM_CODE_COMPARE_TO = '{ITEM_CODE_COMPARE_TO}'"
                If dst.Tables("SATSISTY").Select(sqlY2).Length > 0 Then
                    MsgBox("You cannot use the same item as a Compare-To for both TY and LY" & vbCrLf & $" See Item {ITEM_CODE_TY} which uses {ITEM_CODE_COMPARE_TO} for both LY and 2LY comparisons", MsgBoxStyle.OkOnly, "Delete the Item from either TY or LY before Proceeding")
                    Exit Sub
                End If
            Next
        Next

        If dst.Tables("SATSISTI").Select("").Length = 0 Then
            MsgBox("You must first add items to the grid", MsgBoxStyle.OkOnly, "Cannot Generate XLS until you Select Items")
            Exit Sub
        End If

        Dim RX As String = IIf(optRANGE.Value = "W", "RYW", "RYP")
        For Each T As String In New String() {"SATSISTI", "SATSISTY"}
            For Each row As DataRow In dst.Tables(T).Select("")
                Dim YX0 As String = Absx1.cmbFor(RX & "0").Value
                Dim YX1 As String = Absx1.cmbFor(RX & "1").Value

                Dim YX0_Text As String = Absx1.cmbFor(RX & "0").Text
                Dim YX1_Text As String = Absx1.cmbFor(RX & "1").Text

                If T = "SATSISTY" Then
                    Dim LY As String = row.Item("LY")
                    Dim LYV As Integer = Val(LY)

                    YX0 = Format(Val(Mid(YX0, 1, 4)) - LYV, "0000") & Mid(YX0, 5, 2)
                    YX1 = Format(Val(Mid(YX1, 1, 4)) - LYV, "0000") & Mid(YX1, 5, 2)

                    If optRANGE.Value = "W" Then
                        YX0_Text = ASCMAIN1.Get_Legend_Wk(YX0)
                        YX1_Text = ASCMAIN1.Get_Legend_Wk(YX1)
                    Else
                        YX0_Text = ASCMAIN1.Get_Legend(YX0)
                        YX1_Text = ASCMAIN1.Get_Legend(YX1)
                    End If
                End If

                Dim errormsg As String = "Cannot Provide Custom Range outside Report Range"


                For Each IT As String In New String() {"IN", "TH"}
                    Dim ITD As String = IIf(IT = "IN", "Sell-In", "Sell-Thru")
                    For Each x As String In New String() {"1", "2"}
                        Dim xd As String = IIf(x = "1", "From", "To")
                        Dim YX As String = $"YX_{IT}{x}"

                        Dim YXD As String = row.Item(YX) & ""
                        If YXD <> "" Then
                            If YXD < YX0 Then
                                MsgBox($"{ITD} {xd} {optRANGE.Text} {YXD} must not be Prior to " & YX0_Text, MsgBoxStyle.OkOnly, errormsg)
                                Exit Sub
                            End If
                            If YXD > YX1 Then
                                MsgBox($"{ITD} {xd} {optRANGE.Text} {YXD} must not be Later than " & YX1_Text, MsgBoxStyle.OkOnly, errormsg)
                                Exit Sub
                            End If
                        End If
                    Next
                Next

                'Dim YX_IN1 As String = row.Item("YX_IN1") & ""
                'If YX_IN1 <> "" Then
                '    If YX_IN1 < YX0 Then
                '        MsgBox("Sell-In From " & optRANGE.Text & " must not be Prior to " & YX0_Text, MsgBoxStyle.OkOnly, errormsg)
                '        Exit Sub
                '    End If
                '    If YX_IN1 > YX1 Then
                '        MsgBox("Sell-In From " & optRANGE.Text & " must not be Later than" & YX1_Text, MsgBoxStyle.OkOnly, errormsg)
                '        Exit Sub
                '    End If
                'End If

                'Dim YX_TH1 As String = row.Item("YX_TH1") & ""
                'If YX_TH1 <> "" Then
                '    If YX_TH1 < YX0 Then
                '        MsgBox("Sell-Thru From " & optRANGE.Text & " must not be Prior to " & YX0_Text, MsgBoxStyle.OkOnly, errormsg)
                '        Exit Sub
                '    End If
                '    If YX_TH1 > YX1 Then
                '        MsgBox("Sell-Thru From " & optRANGE.Text & " must not be Later than " & YX1_Text, MsgBoxStyle.OkOnly, errormsg)
                '        Exit Sub
                '    End If
                'End If

                'Dim YX_IN2 As String = row.Item("YX_IN2") & ""
                'If YX_IN2 <> "" Then
                '    If YX_IN2 < YX0 Then
                '        MsgBox("Sell-In To " & optRANGE.Text & " must not be Prior to " & YX0_Text, MsgBoxStyle.OkOnly, errormsg)
                '        Exit Sub
                '    End If
                '    If YX_IN2 > YX1 Then
                '        MsgBox("Sell-In To " & optRANGE.Text & " must not be Later than " & YX1_Text, MsgBoxStyle.OkOnly, errormsg)
                '        Exit Sub
                '    End If
                'End If

                'Dim YX_TH2 As String = row.Item("YX_TH2") & ""
                'If YX_TH2 <> "" Then
                '    If YX_TH2 < YX0 Then
                '        MsgBox("Sell-Thru To " & optRANGE.Text & " must not be Prior to " & YX0_Text, MsgBoxStyle.OkOnly, errormsg)
                '        Exit Sub
                '    End If
                '    If YX_TH2 > YX1 Then
                '        MsgBox("Sell-Thru To " & optRANGE.Text & " must not be Later than " & YX1_Text, MsgBoxStyle.OkOnly, errormsg)
                '        Exit Sub
                '    End If
                'End If

            Next
        Next

        XLS_FILENAME_multi = "*"

        XLS_TITLE = txtTitle.Text
        If XLS_TITLE = "" Then
            For Each row As DataRow In dst.Tables("SATSISTI").Select("", "")
                XLS_TITLE &= "," & row.Item(0)
            Next
            XLS_TITLE = Mid(XLS_TITLE, 2)
        End If

        XLS_DATE_RANGE = ""
        If optRANGE.Value = "P" Then
            XLS_DATE_RANGE = ASCMAIN1.Get_Legend(Absx1.cmbFor("RYP0").Value) & " - " & ASCMAIN1.Get_Legend(Absx1.cmbFor("RYP1").Value)
        Else
            XLS_DATE_RANGE = ASCMAIN1.Get_Legend_Wk(Absx1.cmbFor("RYW0").Value) & " - " & ASCMAIN1.Get_Legend_Wk(Absx1.cmbFor("RYW1").Value)

        End If

        Dim XLS_FILENAME As String = ""
        Dim workbook As SpreadsheetGear.IWorkbook = Nothing
        Dim Start_Row As Integer = 0

        lblITEM_CODE.Visible = True
        'lblITEM_DESC.Visible = True
        txtITEM_CODE.Visible = True
        txtITEM_DESC.Visible = True

        For Each row As DataRow In dst.Tables("SATSISTI").Select("", "ITEM_CODE")
            ITEM_CODE = row.Item("ITEM_CODE")
            Absx1.txtFor("ITEM_CODE").Text = ITEM_CODE
            Click_Command("View")

            lblITEM_CODE.Visible = True
            'lblITEM_DESC.Visible = True
            txtITEM_CODE.Visible = True
            txtITEM_DESC.Visible = True

            Application.DoEvents()

            If chkSIST_SingleSheet.Checked Then
                Excel_Extract(workbook, XLS_FILENAME, Start_Row)
            Else
                Excel_Extract()
            End If

            Click_Command("Done")
        Next

        If Not chkSIST_SingleSheet.Checked Then
            Excel_Extract_Summary()
        End If


        lblITEM_CODE.Visible = False
        'lblITEM_DESC.Visible = False
        txtITEM_CODE.Visible = False
        txtITEM_DESC.Visible = False

        '  Show_Document(XLS_FILENAME_multi)
        Add_Document_to_ASTSPRF1(XLS_FILENAME_multi)
        XLS_FILENAME_multi = ""

        MsgBox("Multiple Sheet XLS has been Generated", MsgBoxStyle.OkOnly, "Success")
    End Sub

    Private Sub grdSATSIST3_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSATSIST3.InitializeRow
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
                    If Format(DATE_START, "yyyyMMdd") > Format(PE_DATEsLY(P - 1), "yyyyMMdd") And
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
        grdSATSIST1.DisplayLayout.Bands(0).Columns("CUST_CODE").Width = W
        grdSATSIST2.DisplayLayout.Bands(0).Columns("CUST_NAME").Width = W
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

        dst.Tables("SATSISTI").Rows.Clear()

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

        Dim CODEs As String = Get_Multi_CODEs("SATSISTC")

        If CODEs <> "" Then
            ASCMAIN1.sql = "Select Distinct ITEM_CODE from SOTINVH2" & vbCrLf _
                & " where CUST_CODE in (" & CODEs & ")" & sqlp
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Add_Item(row.Item(0))
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
        Dim blnUpdateExisting As Boolean = False

        dst.Tables("ASTLIST2").Rows.Clear()
        dst.Tables("ASTLIST1").AcceptChanges()
        dst.Tables("ASTLIST2").AcceptChanges()


        If Trim(txtLIST_DESC.Text) = "" Then
            If LIST_CODE_To_Update <> "" And txtTitle.Text <> "" Then
                If MsgBox("Update existing list with settings for " & txtTitle.Text, MsgBoxStyle.YesNo, $"Update List {LIST_CODE_To_Update}") = MsgBoxResult.No Then
                    Exit Sub
                End If
                blnUpdateExisting = True

            Else
                MsgBox("You Must Provide a Description for this List", MsgBoxStyle.OkOnly, "Cannot Save Lists")
                Exit Sub
            End If
        End If

        BeginTrans()

        Dim rowASTLIST1 As DataRow = Nothing
        Dim LIST_CODE As String = ""
        Dim LIST_DESC As String = IIf(blnUpdateExisting, txtTitle.Text, txtLIST_DESC.Text)

        If blnUpdateExisting Then
            LIST_CODE = LIST_CODE_To_Update
            rowASTLIST1 = dst.Tables("ASTLIST1").Rows.Find(LIST_CODE)
        Else
            LIST_CODE = ASCMAIN1.Next_Control_No("ASTLIST1.LIST_CODE")
            LIST_CODE_To_Update = LIST_CODE
            rowASTLIST1 = dst.Tables("ASTLIST1").NewRow
            With rowASTLIST1
                .Item("LIST_CODE") = LIST_CODE
                .Item("COLUMN_NAME") = "SATISLS1.ITEM_CODE"
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LIST_SHAREABLE") = "0"
                .Item("LIST_MODIFIABLE") = "0"
            End With
        End If

        With rowASTLIST1
            .Item("LIST_DESC") = LIST_DESC
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
        End With
        If Not blnUpdateExisting Then
            dst.Tables("ASTLIST1").Rows.Add(rowASTLIST1)
        End If


        For Each row As DataRow In dst.Tables("SATSISTI").Select("")
            Dim rowASTLIST2 As DataRow = dst.Tables("ASTLIST2").NewRow
            rowASTLIST2.Item("LIST_CODE") = LIST_CODE
            rowASTLIST2.Item("CODE_VALUE") = row.Item("ITEM_CODE")
            dst.Tables("ASTLIST2").Rows.Add(rowASTLIST2)
        Next

        If blnUpdateExisting Then
            LIST_CODE = LINKED_LIST_CODE_To_Update
        Else
            LIST_CODE = ASCMAIN1.Next_Control_No("ASTLIST1.LIST_CODE")
            LINKED_LIST_CODE_To_Update = LIST_CODE
        End If

        rowASTLIST1.Item("LINKED_LIST_CODE") = LIST_CODE ' this is rowASTLIST1 from the Item List above

        If blnUpdateExisting Then
            rowASTLIST1 = dst.Tables("ASTLIST1").Rows.Find(LIST_CODE)
            If rowASTLIST1 Is Nothing Then
                Dim row1 As DataRow = LookUp("ASTLIST1", LIST_CODE)

                dst.Tables("ASTLIST1").Rows.Add(row1.ItemArray)
                rowASTLIST1 = dst.Tables("ASTLIST1").Rows.Find(LIST_CODE)
                rowASTLIST1.AcceptChanges()
            End If
        Else
            rowASTLIST1 = dst.Tables("ASTLIST1").NewRow
            With rowASTLIST1
                .Item("LIST_CODE") = LIST_CODE
                .Item("COLUMN_NAME") = "SATISLS1.CUST_CODE"
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LIST_SHAREABLE") = "0"
                .Item("LIST_MODIFIABLE") = "0"
            End With
        End If

        With rowASTLIST1
            .Item("LIST_DESC") = LIST_DESC
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
        End With
        If Not blnUpdateExisting Then
            dst.Tables("ASTLIST1").Rows.Add(rowASTLIST1)
        End If

        For Each row As DataRow In dst.Tables("SATSISTC").Select("")
            Dim rowASTLIST2 As DataRow = dst.Tables("ASTLIST2").NewRow
            rowASTLIST2.Item("LIST_CODE") = LIST_CODE
            rowASTLIST2.Item("CODE_VALUE") = row.Item("CUST_CODE")
            dst.Tables("ASTLIST2").Rows.Add(rowASTLIST2)
        Next


        dst.Tables("SATSISTZ").Rows.Clear()

        For Each rowSATSISTI As DataRow In dst.Tables("SATSISTI").Select("")
            Dim ITEM_CODE As String = rowSATSISTI.Item("ITEM_CODE")
            Dim blnHasValue As Boolean = False
            Dim rowSATSISTZ As DataRow = Fill_Record("SATSISTZ", ITEM_CODE, False, False)
            If rowSATSISTZ Is Nothing Then rowSATSISTZ = dst.Tables("SATSISTZ").NewRow

            With rowSATSISTZ
                .Item("ITEM_CODE") = ITEM_CODE
                For Each C As String In New String() {"YP_IN1", "YP_IN2", "YP_TH1", "YP_TH2", "YW_IN1", "YW_IN2", "YW_TH1", "YW_TH2"}
                    blnHasValue = blnHasValue Or (rowSATSISTI.Item(C) & "" <> "")
                    .Item(C) = rowSATSISTI.Item(C)
                    If C.StartsWith("Y" & optRANGE.Value) Then
                        If rowSATSISTI.Item(Replace(C, "Y" & optRANGE.Value, "YX")) & "" <> "" Then
                            .Item(C) = rowSATSISTI.Item(Replace(C, "Y" & optRANGE.Value, "YX"))
                            blnHasValue = True
                        Else
                            .Item(C) = DBNull.Value
                        End If
                    End If
                Next
            End With

            If blnHasValue Then
                If rowSATSISTZ.RowState = DataRowState.Detached Then
                    dst.Tables("SATSISTZ").Rows.Add(rowSATSISTZ)
                End If
            Else
                If rowSATSISTZ.RowState <> DataRowState.Detached Then
                    rowSATSISTZ.Delete()
                End If
            End If
        Next

        'If blnUpdateExisting Then
        '    Delete_List(LIST_CODE_To_Update, LINKED_LIST_CODE_To_Update)
        'End If

        Dim sqld As String = ""
        Update_Record_TDA("ASTLIST1", "")
        Update_Record_TDA("ASTLIST2", IIf(blnUpdateExisting, $"LIST_CODE = '{LINKED_LIST_CODE_To_Update}' or LIST_CODE = '{LIST_CODE_To_Update}'", ""))

        For Each rowSATSISTY As DataRow In dst.Tables("SATSISTY").Select("")
            With rowSATSISTY
                For Each C As String In New String() {"YP_IN1", "YP_IN2", "YP_TH1", "YP_TH2", "YW_IN1", "YW_IN2", "YW_TH1", "YW_TH2"}
                    If C.StartsWith("Y" & optRANGE.Value) Then
                        If rowSATSISTY.Item(Replace(C, "Y" & optRANGE.Value, "YX")) & "" <> "" Then
                            .Item(C) = rowSATSISTY.Item(Replace(C, "Y" & optRANGE.Value, "YX"))
                        Else
                            .Item(C) = DBNull.Value
                        End If
                    End If
                Next

                If .RowState = DataRowState.Added Then
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                End If
                If .RowState = DataRowState.Added Or .RowState = DataRowState.Modified Then
                    .Item("LAST_OPER") = ASCMAIN1.USER_ID
                    .Item("LAST_DATE") = DATETIME_STAMP
                End If
            End With
        Next

        'sqld = ""
        'For Each rowITEM_CODE As DataRow In ASCDATA1.SelectDistinct("SATSISTY", "ITEM_CODE").Select
        '    Dim ITEM_CODE = rowITEM_CODE.Item("ITEM_CODE")
        '    sqld &= $",'{ITEM_CODE}'"
        'Next
        'If sqld <> "" Then sqld = "ITEM_CODE in (" & Mid(sqld, 2) & ")"

        'Update_Record_TDA("SATSISTY", sqld)

        Update_Record_TDA("SATSISTY")

        Update_Record_TDA("SATSISTZ")

        Sort_grdColumns(grdASTLIST1, "INIT_DATE".ToLower)

        CommitTrans("Item & Customer Lists have been Saved")

    End Sub

    Private Sub grdASTLIST1_AfterRowsDeleted(sender As Object, e As EventArgs) Handles grdASTLIST1.AfterRowsDeleted
        For Each LIST_CODE As String In LIST_CODEs_To_Delete.Keys
            Dim LINKED_LIST_CODE As String = LIST_CODEs_To_Delete(LIST_CODE)
            Delete_List(LIST_CODE, LINKED_LIST_CODE)
        Next
    End Sub

    Private Sub grdASTLIST1_BeforeRowsDeleted(sender As Object, e As UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdASTLIST1.BeforeRowsDeleted
        LIST_CODEs_To_Delete.Clear()

        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            Dim LIST_CODE As String = grow.Cells("LIST_CODE").Value
            Dim LINKED_LIST_CODE As String = grow.Cells("LINKED_LIST_CODE").Value
            LIST_CODEs_To_Delete.Add(LIST_CODE, LINKED_LIST_CODE)
            'ASCDATA1.ExecuteSQL("Delete from ASTLIST1 where LIST_CODE = '" & LIST_CODE & "'")
            'ASCDATA1.ExecuteSQL("Delete from ASTLIST2 where LIST_CODE = '" & LIST_CODE & "'")
            'ASCDATA1.ExecuteSQL("Delete from ASTLIST1 where LIST_CODE = '" & LINKED_LIST_CODE & "'")
            'ASCDATA1.ExecuteSQL("Delete from ASTLIST2 where LIST_CODE = '" & LINKED_LIST_CODE & "'")
        Next
    End Sub

    Sub Delete_List(LIST_CODE As String, LINKED_LIST_CODE As String)
        ASCDATA1.ExecuteSQL("Delete from ASTLIST1 where LIST_CODE = '" & LIST_CODE & "'")
        ASCDATA1.ExecuteSQL("Delete from ASTLIST2 where LIST_CODE = '" & LIST_CODE & "'")
        ASCDATA1.ExecuteSQL("Delete from ASTLIST1 where LIST_CODE = '" & LINKED_LIST_CODE & "'")
        ASCDATA1.ExecuteSQL("Delete from ASTLIST2 where LIST_CODE = '" & LINKED_LIST_CODE & "'")
    End Sub

    Private Sub grdASTLIST1_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdASTLIST1.DoubleClickRow

        Dim LIST_CODE As String = e.Row.Cells("LIST_CODE").Value
        Dim LINKED_LIST_CODE As String = e.Row.Cells("LINKED_LIST_CODE").Value
        Dim rowASTLIST1 As DataRow = dst.Tables("ASTLIST1").Rows.Find(LIST_CODE)

        LIST_CODE_To_Update = LIST_CODE
        LINKED_LIST_CODE_To_Update = LINKED_LIST_CODE

        txtTitle.Text = rowASTLIST1.Item("LIST_DESC") & ""

        dst.Tables("SATSISTY").Rows.Clear()
        'grdSATSISTY.Visible = False

        dst.Tables("SATSISTI").Rows.Clear()
        Fill_Records("ASTLIST2", LIST_CODE)
        For Each row As DataRow In dst.Tables("ASTLIST2").Select("")
            Dim CODE_VALUE As String = row.Item("CODE_VALUE")
            If grdSATSISTI.ActiveRow IsNot Nothing Then grdSATSISTI.ActiveRow.CancelUpdate()
            grdSATSISTI.ActiveRow = grdSATSISTI.DisplayLayout.Bands(0).AddNew
            grdSATSISTI.ActiveRow.Cells("ITEM_CODE").Value = CODE_VALUE
            grdSATSISTI.ActiveRow.Update()

        Next

        grdSATSISTI.PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)

        Setup_grdSATSISTY()


        dst.Tables("SATSISTC").Rows.Clear()
        Fill_Records("ASTLIST2", LINKED_LIST_CODE)
        For Each row As DataRow In dst.Tables("ASTLIST2").Select("")
            Dim CODE_VALUE As String = row.Item("CODE_VALUE")
            If grdSATSISTC.ActiveRow IsNot Nothing Then grdSATSISTC.ActiveRow.CancelUpdate()
            grdSATSISTC.ActiveRow = grdSATSISTC.DisplayLayout.Bands(0).AddNew
            grdSATSISTC.ActiveRow.Cells("CUST_CODE").Value = CODE_VALUE
            grdSATSISTC.ActiveRow.Update()
        Next
        grdSATSISTC.PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)

    End Sub

    Private Sub cmdSaveExcludeCustomers_Click(sender As Object, e As EventArgs) Handles cmdSaveExcludeCustomers.Click
        If grdSATSISTX.Rows.Count = 0 Then
            MsgBox("No Customers to Save", MsgBoxStyle.OkOnly, "Cannot Save List")
            Exit Sub
        End If

        Dim DATETIME_STAMP As Date = Now + ASCMAIN1.NowTSD

        dst.Tables("ASTLIST2").Rows.Clear()
        dst.Tables("ASTLIST1").AcceptChanges()
        dst.Tables("ASTLIST2").AcceptChanges()

        Dim rowASTLIST1 As DataRow = Nothing
        Dim LIST_CODE As String = ""

        LIST_CODE = Me.Name
        rowASTLIST1 = dst.Tables("ASTLIST1").Rows.Find(LIST_CODE)
        If rowASTLIST1 Is Nothing Then
            rowASTLIST1 = dst.Tables("ASTLIST1").NewRow
            With rowASTLIST1
                .Item("LIST_CODE") = LIST_CODE
                .Item("LIST_DESC") = grdSATSISTX.Text
                .Item("COLUMN_NAME") = "SATISLS2.CUST_CODE"

                .Item("LIST_SHAREABLE") = "0"
                .Item("LIST_MODIFIABLE") = "0"
            End With
            dst.Tables("ASTLIST1").Rows.Add(rowASTLIST1)
        Else
            rowASTLIST1.SetAdded()
        End If

        With rowASTLIST1
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
        End With

        For Each row As DataRow In dst.Tables("SATSISTX").Select("")
            Dim rowASTLIST2 As DataRow = dst.Tables("ASTLIST2").NewRow
            rowASTLIST2.Item("LIST_CODE") = LIST_CODE
            rowASTLIST2.Item("CODE_VALUE") = row.Item("CUST_CODE")
            dst.Tables("ASTLIST2").Rows.Add(rowASTLIST2)
        Next

        ASCDATA1.ExecuteSQL("Delete from ASTLIST1 where LIST_CODE = '" & LIST_CODE & "'")
        ASCDATA1.ExecuteSQL("Delete from ASTLIST2 where LIST_CODE = '" & LIST_CODE & "'")

        Dim sqld As String = ""
        Update_Record_TDA("ASTLIST1", sqld)
        Update_Record_TDA("ASTLIST2", sqld)

        rowASTLIST1.Delete()
        dst.Tables("ASTLIST1").AcceptChanges()

        MsgBox("List of Customers to Exclude has been Saved", MsgBoxStyle.OkOnly, "Verification")
    End Sub

    Private Sub grdSATSISTI_KeyPress(sender As Object, e As KeyPressEventArgs) Handles grdSATSISTI.KeyPress
        'If e.KeyChar.Equals(Keys.Delete) Then
        '    Dim C As String = ""gotom
        'End If
    End Sub

    Private Sub grdSATSISTI_KeyDown(sender As Object, e As KeyEventArgs) Handles grdSATSISTI.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Delete Then
            If grdSATSISTI.ActiveCell Is Nothing Then
                Exit Sub
            Else
                If grdSATSISTI.ActiveCell.Value & "" <> "" Then
                    grdSATSISTI.ActiveCell.Value = ""
                    grdSATSISTI.ActiveRow.Update()
                Else
                    grdSATSISTI.ActiveCell.CancelUpdate()
                End If
            End If
        End If
    End Sub

    Private Sub grdSATSISTI_DoubleClickCell(sender As Object, e As DoubleClickCellEventArgs) Handles grdSATSISTI.DoubleClickCell
        Dim C As String = e.Cell.Column.Key
        If New String() {"YX_IN1", "YX_IN2", "YX_TH1", "YX_TH2"}.Contains(C) Then
            If optRANGE.Value = "W" Then
                If C.EndsWith("1") Then
                    e.Cell.Value = Absx1.cmbFor("RYW0").Value
                Else
                    e.Cell.Value = Absx1.cmbFor("RYW1").Value
                End If
            Else
                If C.EndsWith("1") Then
                    e.Cell.Value = Absx1.cmbFor("RYP0").Value
                Else
                    e.Cell.Value = Absx1.cmbFor("RYP1").Value
                End If
            End If
        End If
    End Sub

    Private Sub grdSATSISTX_KeyDown(sender As Object, e As KeyEventArgs) Handles grdSATSISTX.KeyDown

    End Sub

    Private Sub grdSATSISTX_KeyPress(sender As Object, e As KeyPressEventArgs) Handles grdSATSISTX.KeyPress

    End Sub

    Private Sub grdSATSISTY_KeyDown(sender As Object, e As KeyEventArgs) Handles grdSATSISTY.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Delete Then
            If grdSATSISTY.ActiveCell Is Nothing Then
                Exit Sub
            End If
            Dim COLUMN_NAME As String = grdSATSISTY.ActiveCell.Column.Key
            If Not COLUMN_NAME.StartsWith("YX") Then ' If Not New String() {"", ""}.Contains(COLUMN_NAME) Then
                Exit Sub
            End If
            If grdSATSISTY.ActiveCell.Value & "" <> "" Then
                grdSATSISTY.ActiveCell.Value = ""
                grdSATSISTY.ActiveRow.Update()
            Else
                grdSATSISTY.ActiveCell.CancelUpdate()
            End If

        End If
    End Sub

    Private Sub grdSATSISTY_KeyPress(sender As Object, e As KeyPressEventArgs) Handles grdSATSISTY.KeyPress
        'If e.KeyChar.Equals(Keys.Delete) Then
        '    Dim C As String = ""
        'End If
    End Sub
End Class