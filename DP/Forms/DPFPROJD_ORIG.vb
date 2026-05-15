Public Class DPFPROJD_ORIG

    Dim ICTITEMS As String
    Dim YP(,) As String
    Dim YP_LY(,) As String
    Dim OPS_YYYY As String
    Dim DPTPROJI As String
    Dim DPTFCSTD As String
    Dim sqlDPTPROJI As String
    Dim grdDPTFCSTD_HCs As String()
    Dim SEASON As String
    Dim rowDPTPROJ1 As DataRow
    Dim TOTAL_SALES() As Decimal
    Dim VERSION_STATUS As String
    Dim VERSION_to_copy As String
    Dim sqlICTITEMS As String
    Dim negatives_only As Boolean = False

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")

        With dst
            Call Create_TDA(.Tables.Add, "ICTBRAN1", "*")

            ' ICTITEMS - All Items on Screen

            sqlICTITEMS = "Select ITEM_CODE, ITEM_DESC, ITEM_CATGY_CODE, ITEM_CATGY_CODE ITEM_CATGY_CODE_CURR" _
            & ", ITEM_CLASS_CODE, COLLECTION_CODE, DEPT_CODE, VEND_CODE, VEND_CODE VEND_CODE_2" _
            & ", ITEM_RETAIL_PRICE, ITEM_PRICE, ITEM_COST_STD, ITEM_PICTURE_FILENAME"
            ASCMAIN1.sql = sqlICTITEMS & " from ICTITEM1 where ROWNUM < 1"
            ICTITEMS = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add Primary Key (ITEM_CODE)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add PROJ_QOH_BEG NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add QTY_OPO NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add QTY_FC NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add AMT_FC NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add PROJ_QOH_END NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add QTY_FC1 NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add QTY_FC2 NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add QTY_FC3 NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add QTY_FC4 NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add QTY_FC5 NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add QTY_FC6 NUMBER (8,0)")

            ASCMAIN1.sql = "SELECT ITEM_CODE, 'XXXXXXXXXX' DATA_TYPE" _
            & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'YP1',SOTINVH2.ORDR_QTY_SHIP,0)) Q1" _
            & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'YP2',SOTINVH2.ORDR_QTY_SHIP,0)) Q2" _
            & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'YP3',SOTINVH2.ORDR_QTY_SHIP,0)) Q3" _
            & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'YP4',SOTINVH2.ORDR_QTY_SHIP,0)) Q4" _
            & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'YP5',SOTINVH2.ORDR_QTY_SHIP,0)) Q5" _
            & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'YP6',SOTINVH2.ORDR_QTY_SHIP,0)) Q6" _
            & " from SOTINVH2" _
            & " where SOTINVH2.ORDR_YYYYPP_UPDATED BETWEEN 'YP1' AND 'YP6'" _
            & " group by SOTINVH2.ITEM_CODE"
            DPTFCSTD = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & DPTFCSTD & " Add Primary Key (ITEM_CODE, DATA_TYPE)")

            Dim SQLQ As String = ", DPTFCSTD_TY_IN.Q1 TY_IN1, DPTFCSTD_TY_IN.Q2 TY_IN2, DPTFCSTD_TY_IN.Q3 TY_IN3, DPTFCSTD_TY_IN.Q4 TY_IN4, DPTFCSTD_TY_IN.Q5 TY_IN5, DPTFCSTD_TY_IN.Q6 TY_IN6"

            ASCMAIN1.sql = "Select ICTITEMS.*" _
            & ", NVL(DPTPROJ2.PROJ_P1,0) PROJ_P1" _
            & ", NVL(DPTPROJ2.PROJ_P2,0) PROJ_P2" _
            & ", NVL(DPTPROJ2.PROJ_P3,0) PROJ_P3" _
            & ", NVL(DPTPROJ2.PROJ_P4,0) PROJ_P4" _
            & ", NVL(DPTPROJ2.PROJ_P5,0) PROJ_P5" _
            & ", NVL(DPTPROJ2.PROJ_P6,0) PROJ_P6" _
            & SQLQ _
            & Replace(SQLQ, "TY_IN", "LY_IN") _
            & Replace(SQLQ, "TY_IN", "TY_ST") _
            & Replace(SQLQ, "TY_IN", "LY_ST") _
            & " from " & ICTITEMS & " ICTITEMS, DPTPROJ2, DPTPROJ2 DPTPROJ2_O" _
            & "," & DPTFCSTD & " DPTFCSTD_TY_IN" _
            & "," & DPTFCSTD & " DPTFCSTD_LY_IN" _
            & "," & DPTFCSTD & " DPTFCSTD_TY_ST" _
            & "," & DPTFCSTD & " DPTFCSTD_LY_ST" _
            & " where DPTPROJ2.BRAND_CODE (+) = :PARM1" _
            & "   and DPTPROJ2.MARKET_CODE (+) = :PARM2" _
            & "   and DPTPROJ2.OPS_YYYY (+) = :PARM3" _
            & "   and DPTPROJ2.SEASON (+) = :PARM4" _
            & "   and DPTPROJ2.VERSION (+) = :PARM5" _
            & "   and DPTPROJ2_O.VERSION (+) = :PARM6" _
            & "   and DPTPROJ2.ITEM_CODE (+) = ICTITEMS.ITEM_CODE" _
            & "   and DPTPROJ2_O.BRAND_CODE (+) = :PARM1" _
            & "   and DPTPROJ2_O.MARKET_CODE (+) = :PARM2" _
            & "   and DPTPROJ2_O.OPS_YYYY (+) = :PARM3" _
            & "   and DPTPROJ2_O.SEASON (+) = :PARM4" _
            & "   and DPTPROJ2_O.VERSION (+) = :PARM5" _
            & "   and DPTPROJ2_O.ITEM_CODE (+) = ICTITEMS.ITEM_CODE" _
            & "   and ICTITEMS.ITEM_CATGY_CODE <> 'I'" _
            & "   and DPTFCSTD_TY_IN.ITEM_CODE (+) = ICTITEMS.ITEM_CODE" _
            & "   and DPTFCSTD_TY_IN.DATA_TYPE (+) = 'TY_IN'" _
            & "   and DPTFCSTD_LY_IN.ITEM_CODE (+) = ICTITEMS.ITEM_CODE" _
            & "   and DPTFCSTD_LY_IN.DATA_TYPE (+) = 'LY_IN'" _
            & "   and DPTFCSTD_TY_ST.ITEM_CODE (+) = ICTITEMS.ITEM_CODE" _
            & "   and DPTFCSTD_TY_ST.DATA_TYPE (+) = 'TY_ST'" _
            & "   and DPTFCSTD_LY_ST.ITEM_CODE (+) = ICTITEMS.ITEM_CODE" _
            & "   and DPTFCSTD_LY_ST.DATA_TYPE (+) = 'LY_ST'"

            Create_TDA(.Tables.Add, "DPTFCSTD", "**", 0, False, "VVVVVV", 1)
            With .Tables("DPTFCSTD")
                Dim SQL0 As String = "'"
                For M As Integer = 1 To 6
                    Dim T As System.Type
                    If M = 0 Then
                        T = GetType(System.String)
                    Else
                        T = GetType(System.Decimal)
                    End If

                    .Columns.Add("FC" & Format(M, "0"), T)
                    '.Columns.Add("FCO" & Format(M, "0"), T)
                    '.Columns.Add("TY_IN" & Format(M, "0"), T)
                    '.Columns.Add("LY_IN" & Format(M, "0"), T)
                    '.Columns.Add("TY_THRU" & Format(M, "0"), T)
                    '.Columns.Add("LY_THRU" & Format(M, "0"), T)
                    SQL0 &= "+ISNULL(FC" & Format(M, "0") & ",0)"
                Next
                SQL0 = Mid(SQL0, 2)
                .Columns.Add("FC0", GetType(System.Decimal), SQL0)
                .Columns.Add("PROJ_P0", GetType(System.Decimal), Replace(SQL0, "FC", "PROJ_P"))
                .Columns.Add("TY_IN0", GetType(System.Decimal), Replace(SQL0, "FC", "TY_IN"))
                .Columns.Add("LY_IN0", GetType(System.Decimal), Replace(SQL0, "FC", "LY_IN"))
                .Columns.Add("TY_ST0", GetType(System.Decimal), Replace(SQL0, "FC", "TY_ST"))
                .Columns.Add("LY_ST0", GetType(System.Decimal), Replace(SQL0, "FC", "LY_ST"))

            End With
            '.Tables("DPTFCSTD").Columns("QTY_FC").Expression = "ISNULL(QTY_FC1,0)+ISNULL(QTY_FC2,0)+ISNULL(QTY_FC3,0)+ISNULL(QTY_FC4,0)+ISNULL(QTY_FC5,0)+ISNULL(QTY_FC6,0)"
            ' .Tables("DPTFCSTD").Columns("").Expression = ""

            With .Tables.Add("DPTPROJA")
                .Columns.Add("PCT", GetType(System.Int32))
                .Columns.Add("LINE_DESC")
                For Each ITEM_CATGY_CODE As String In New String() {"N", "C", "E", "I"}
                    .Columns.Add(ITEM_CATGY_CODE & "_ITEMS", GetType(System.Int32))
                    .Columns.Add(ITEM_CATGY_CODE & "_PCT", GetType(System.Double))
                    .Columns.Add(ITEM_CATGY_CODE & "_QTY_PROJ", GetType(System.Int32))
                    .Columns.Add(ITEM_CATGY_CODE & "_QTY_SHIP", GetType(System.Int32))
                    .Columns.Add(ITEM_CATGY_CODE & "_AMT_PROJ", GetType(System.Double))
                    .Columns.Add(ITEM_CATGY_CODE & "_AMT_SHIP", GetType(System.Double))
                    .Columns.Add(ITEM_CATGY_CODE & "_QTY_PROJ_PCT", GetType(System.Int32))
                    .Columns.Add(ITEM_CATGY_CODE & "_QTY_SHIP_PCT", GetType(System.Int32))
                    .Columns.Add(ITEM_CATGY_CODE & "_AMT_PROJ_PCT", GetType(System.Double))
                    .Columns.Add(ITEM_CATGY_CODE & "_AMT_SHIP_PCT", GetType(System.Double))
                Next
            End With



            ASCMAIN1.sql = "" _
            & "SELECT X.ITEM_CODE, ICTITEM1.ITEM_DESC" _
            & ", ICTRETLA.ITEM_PRICE, ICTRETLA.ITEM_CATGY_CODE" _
            & ", SUM (QTY_PROJ) QTY_PROJ, SUM (QTY_PROJ * ICTRETLA.ITEM_PRICE) AMT_PROJ" _
            & ", SUM (QTY_SHIP) QTY_SHIP, SUM (AMT_SHIP) AMT_SHIP" _
            & ", SUM (F01) F01, SUM (F02) F02, SUM (F03) F03, SUM (F04) F04, SUM (F05) F05, SUM (F06) F06" _
            & ", SUM (A01) A01, SUM (A02) A02, SUM (A03) A03, SUM (A04) A04, SUM (A05) A05, SUM (A06) A06" _
            & " FROM ICTITEM1,ICTRETLA,(" _
            & " SELECT ITEM_CODE, SUM (FORECAST) QTY_PROJ, 0 QTY_SHIP, 0 AMT_SHIP" _
            & ", SUM (DECODE(OPS_YYYYPP_FC,:PARM1,FORECAST,0)) F01" _
            & ", SUM (DECODE(OPS_YYYYPP_FC,:PARM2,FORECAST,0)) F02" _
            & ", SUM (DECODE(OPS_YYYYPP_FC,:PARM3,FORECAST,0)) F03" _
            & ", SUM (DECODE(OPS_YYYYPP_FC,:PARM4,FORECAST,0)) F04" _
            & ", SUM (DECODE(OPS_YYYYPP_FC,:PARM5,FORECAST,0)) F05" _
            & ", SUM (DECODE(OPS_YYYYPP_FC,:PARM6,FORECAST,0)) F06" _
            & ", 0 A01, 0 A02, 0 A03, 0 A04, 0 A05, 0 A06" _
            & " FROM DPTITMF1" _
            & " WHERE OPS_YYYYPP BETWEEN :PARM1 AND :PARM6" _
            & " AND OPS_YYYYPP_FC = OPS_YYYYPP" _
            & " GROUP BY ITEM_CODE" _
            & " UNION" _
            & " SELECT ITEM_CODE, 0 PROJ" _
            & ", SUM (ORDR_QTY_SHIP) QTY_SHIP, SUM (ORDR_AMT_SHIP) AMT_SHIP" _
            & ", 0 F01, 0 F02, 0 F03, 0 F04, 0 F05, 0 F06" _
            & ", SUM (DECODE(OPS_YYYYPP,:PARM1,ORDR_QTY_SHIP,0)) A01" _
            & ", SUM (DECODE(OPS_YYYYPP,:PARM2,ORDR_QTY_SHIP,0)) A02" _
            & ", SUM (DECODE(OPS_YYYYPP,:PARM3,ORDR_QTY_SHIP,0)) A03" _
            & ", SUM (DECODE(OPS_YYYYPP,:PARM4,ORDR_QTY_SHIP,0)) A04" _
            & ", SUM (DECODE(OPS_YYYYPP,:PARM5,ORDR_QTY_SHIP,0)) A05" _
            & ", SUM (DECODE(OPS_YYYYPP,:PARM6,ORDR_QTY_SHIP,0)) A06" _
            & " FROM SATSSUMI" _
            & " WHERE OPS_YYYYPP BETWEEN :PARM1 AND :PARM6" _
            & " AND INV_TYPE = 'I'" _
            & " GROUP BY ITEM_CODE" _
            & ") X WHERE X.ITEM_CODE = ICTITEM1.ITEM_CODE" _
            & " AND ICTRETLA.OPS_YYYYPP (+) = :PARM1" _
            & " AND ICTRETLA.ITEM_CODE (+) = X.ITEM_CODE" _
            & " AND ICTRETLA.ITEM_CATGY_CODE IN ('C','N','E','I')" _
            & " GROUP BY X.ITEM_CODE, ICTITEM1.ITEM_DESC" _
            & ", ICTRETLA.ITEM_PRICE, ICTRETLA.ITEM_CATGY_CODE" _
            & " HAVING SUM (QTY_PROJ) <> 0 OR SUM (QTY_SHIP) <> 0"
            sqlDPTPROJI = ASCMAIN1.sql
            DPTPROJI = ASCMAIN1.Temp_Table(Replace(Replace(Replace(Replace(Replace(Replace(sqlDPTPROJI, ":PARM1", "''"), ":PARM2", "''"), ":PARM3", "''"), ":PARM4", "''"), ":PARM5", "''"), ":PARM6", "''"))
            'ASCDATA1.ExecuteSQL("Alter Table " & DPTPROJI & " Add PCT NUMBER (10,4)")

            ASCMAIN1.sql = "Select * from " & DPTPROJI
            Create_TDA(.Tables.Add, "DPTPROJI", "**", 0, False, "", 1)
            .Tables("DPTPROJI").Columns.Add("VAR_PCT", GetType(System.Double), "IIF(ISNULL(QTY_PROJ,0)=0,100,100 * ISNULL(QTY_SHIP,0)/ISNULL(QTY_PROJ,0)-100)")
            .Tables("DPTPROJI").Columns.Add("ABS_VAR_PCT", GetType(System.Double), "IIF(VAR_PCT<0,-1*VAR_PCT,VAR_PCT)")
            .Tables("DPTPROJI").Columns.Add("PCT", GetType(System.Int32))
            '.Tables("DPTPROJI").Columns.Add("QTY_PROJ_PCT", GetType(System.Double), "100 * ISNULL(QTY_PROJ,0) / 1")
            '.Tables("DPTPROJI").Columns.Add("AMT_PROJ_PCT", GetType(System.Double), "100 * ISNULL(AMT_PROJ,0) / 1")
            '.Tables("DPTPROJI").Columns.Add("QTY_SHIP_PCT", GetType(System.Double), "100 * ISNULL(QTY_SHIP,0) / 1")
            '.Tables("DPTPROJI").Columns.Add("AMT_SHIP_PCT", GetType(System.Double), "100 * ISNULL(AMT_SHIP,0) / 1")

            With .Tables.Add("DPTFCSTV")
                .Columns.Add("VERSION")
                .Columns.Add("VERSION_DESC")
            End With

            ASCMAIN1.sql = "Select * from DPTPROJ1" _
            & " where DPTPROJ1.BRAND_CODE (+) = :PARM1" _
            & "   and DPTPROJ1.MARKET_CODE (+) = :PARM2" _
            & "   and DPTPROJ1.OPS_YYYY (+) = :PARM3" _
            & "   and DPTPROJ1.SEASON (+) = :PARM4" _
            & "   and DPTPROJ1.VERSION (+) = :PARM5"
            Create_TDA(.Tables.Add, "DPTPROJ1", "**", 0, True, "VVVVV")
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "DPTPROJ1", "DPTPROJ2")
            Create_TDA(.Tables.Add, "DPTPROJ2", "**", 0, True, "VVVVV")

            ASCMAIN1.sql = "Select ITEM_CODE, VERSION" _
            & ", PROJ_P1, PROJ_P2, PROJ_P3, PROJ_P4, PROJ_P5, PROJ_P6" _
            & " from DPTPROJ2 " _
            & " where BRAND_CODE = :PARM1 " _
            & " and OPS_YYYY = :PARM2 " _
            & " and SEASON = :PARM3 " _
            & " and ITEM_CODE = :PARM4"
            Create_TDA(.Tables.Add, "DPTFCSTI", "**", 0, False, "VVVV", 2)
            .Tables("DPTFCSTI").Columns.Add("PROJ_P0", GetType(System.Int32), "ISNULL(PROJ_P1,0)+ISNULL(PROJ_P2,0)+ISNULL(PROJ_P3,0)+ISNULL(PROJ_P4,0)+ISNULL(PROJ_P5,0)+ISNULL(PROJ_P6,0)")

            Create_TDA(.Tables.Add, "DPTPROJ0", "*")
            Create_TDA(.Tables.Add, "DPTITMF1", "*")

            ASCMAIN1.sql = "Select ITEM_CODE, ITEM_DESC, ITEM_CATGY_CODE from ICTITEM1"
            Create_TDA(.Tables.Add, "ICTITEMI", "**", 0, False, "", 1)
            .Tables("ICTITEMI").Columns.Add("ITEM_CATGY_CODE_CURR")
            .Tables("ICTITEMI").Columns.Add("BAD_ITEM")
        End With

        For PCT As Integer = 5 To 100 Step 5
            Dim LINE_DESC As String = Format((100 - PCT) / 100, "##0%") & " to " & Format((100 + PCT) / 100, "##0%")
            dst.Tables("DPTPROJA").Rows.Add(New Object() {PCT, LINE_DESC})
        Next

        grdDPTFCSTD.DataSource = dst.Tables("DPTFCSTD")
        grdDPTPROJA.DataSource = dst.Tables("DPTPROJA")
        grdDPTPROJI.DataSource = dst.Tables("DPTPROJI")
        grdDPTFCSTI.DataSource = dst.Tables("DPTFCSTI")
        grdICTITEMI.DataSource = dst.Tables("ICTITEMI")

        Sort_grdColumns(grdDPTPROJA, "PCT", True)

        Dim G As UltraWinGrid.UltraGridGroup
        With grdDPTPROJA.DisplayLayout.Bands(0)
            .Columns("PCT").Hidden = True
            G = .Groups.Add("LINE_DESC", "")
            G.Header.Appearance.TextHAlign = HAlign.Center
            .Columns("LINE_DESC").Group = G
            .Columns("LINE_DESC").Header.Caption = "Accuracy"
            .Columns("LINE_DESC").CellAppearance.BackColor = Color.Beige
            .Columns("LINE_DESC").Header.Appearance.TextHAlign = HAlign.Center
            .Columns("LINE_DESC").CellAppearance.TextHAlign = HAlign.Center

            For Each ITEM_CATGY_DESC In New String() {"New", "Core", "Existing", "Inactive"}
                Dim ITEM_CATGY_CODE = Mid(ITEM_CATGY_DESC, 1, 1)
                G = .Groups.Add(ITEM_CATGY_CODE, ITEM_CATGY_DESC)
                G.Header.Appearance.TextHAlign = HAlign.Center
                G.Header.Appearance.BackColor = IIf(ITEM_CATGY_CODE = "N", Color.Yellow, (IIf(ITEM_CATGY_CODE = "E", Color.LightGreen, IIf(ITEM_CATGY_CODE = "C", Color.HotPink, Color.LightGray))))
                G.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                G.Width = 210
                .Columns(ITEM_CATGY_CODE & "_ITEMS").Group = G
                .Columns(ITEM_CATGY_CODE & "_ITEMS").Width = 70
                .Columns(ITEM_CATGY_CODE & "_ITEMS").Format = "###,##0"
                .Columns(ITEM_CATGY_CODE & "_ITEMS").Header.Caption = "#Items"
                .Columns(ITEM_CATGY_CODE & "_PCT").Group = G
                .Columns(ITEM_CATGY_CODE & "_PCT").Width = 60
                .Columns(ITEM_CATGY_CODE & "_PCT").Format = "##.0"
                .Columns(ITEM_CATGY_CODE & "_PCT").Header.Caption = "%"
                .Columns(ITEM_CATGY_CODE & "_PCT").CellAppearance.BackColor = Color.Beige
                For Each QA As String In New String() {"QTY", "AMT"}
                    Dim QA_desc As String = IIf(QA = "QTY", "#", "$")
                    For Each PS As String In New String() {"PROJ", "SHIP"}
                        Dim PS_desc As String = IIf(PS = "PROJ", "Proj", "Ship")
                        Dim DC As String = "_" & QA & "_" & PS
                        .Columns(ITEM_CATGY_CODE & DC).Group = G
                        .Columns(ITEM_CATGY_CODE & DC).Width = 80
                        .Columns(ITEM_CATGY_CODE & DC).Format = "###,##0"
                        .Columns(ITEM_CATGY_CODE & DC).Header.Caption = QA_desc & PS_desc
                        .Columns(ITEM_CATGY_CODE & DC & "_PCT").Group = G
                        .Columns(ITEM_CATGY_CODE & DC & "_PCT").Width = 70
                        .Columns(ITEM_CATGY_CODE & DC & "_PCT").Format = "##0.0"
                        .Columns(ITEM_CATGY_CODE & DC & "_PCT").Header.Caption = QA_desc & PS_desc & "%"
                        .Columns(ITEM_CATGY_CODE & DC & "_PCT").CellAppearance.BackColor = Color.Beige
                        Create_Summary(grdDPTPROJA, ITEM_CATGY_CODE & DC, "Max", , "###,##0")
                    Next
                Next
                Create_Summary(grdDPTPROJA, ITEM_CATGY_CODE & "_ITEMS", "Max", , "###,##0")
                'Create_Summary(grdDPTPROJA, ITEM_CATGY_CODE & "_PCT", , , "##.0")
            Next

        End With


        With grdDPTPROJI.DisplayLayout.Bands(0)
            .Columns("ITEM_CODE").Header.Caption = "Item Code"
            .Columns("ITEM_CODE").Width = 100
            .Columns("ITEM_DESC").Header.Caption = "Description"
            .Columns("ITEM_DESC").Width = 200
            .Columns("ITEM_PRICE").Header.Caption = "Price"
            .Columns("ITEM_PRICE").Width = 90
            .Columns("ITEM_PRICE").Format = "###,##0.00"
            .Columns("ITEM_CATGY_CODE").Header.Caption = "Category"
            .Columns("ITEM_CATGY_CODE").Width = 90
            .Columns("QTY_PROJ").Header.Caption = "STD #Proj"
            .Columns("QTY_PROJ").Width = 90
            .Columns("QTY_PROJ").Format = "###,##0"
            .Columns("AMT_PROJ").Header.Caption = "STD $Proj"
            .Columns("AMT_PROJ").Width = 90
            .Columns("AMT_PROJ").Format = "###,##0"
            .Columns("QTY_SHIP").Header.Caption = "STD #Ship"
            .Columns("QTY_SHIP").Width = 90
            .Columns("QTY_SHIP").Format = "###,##0"
            .Columns("AMT_SHIP").Header.Caption = "STD $Ship"
            .Columns("AMT_SHIP").Width = 90
            .Columns("AMT_SHIP").Format = "###,##0"
            .Columns("VAR_PCT").Header.Caption = "Var%"
            .Columns("VAR_PCT").Width = 70
            .Columns("VAR_PCT").Format = "##.0"
            .Columns("ABS_VAR_PCT").Header.Caption = "(Var)%"
            .Columns("ABS_VAR_PCT").Width = 70
            .Columns("ABS_VAR_PCT").Format = "##.0"
            .Columns("PCT").Header.Caption = "%"
            .Columns("PCT").Width = 70
            .Columns("PCT").Format = "##.0"

            For Each T As String In New String() {"F", "A"}
                Dim TC As System.Drawing.Color = IIf(T = "A", Color.LightGreen, Color.HotPink)

                For I As Integer = 1 To 6
                    Dim COLUMN_NAME As String = T & Format(I, "00")
                    .Columns(COLUMN_NAME).Header.Caption = "STD #Ship"
                    .Columns(COLUMN_NAME).CellAppearance.BackColor = TC
                    .Columns(COLUMN_NAME).Width = 70
                    .Columns(COLUMN_NAME).Format = "###,##0"
                    Create_Summary(grdDPTPROJI, COLUMN_NAME, , , "###,##0")
                Next
            Next
        End With
        '.Columns(COLUMN_NAME).Header.Caption = IIf(T = "A", "Act", "Prj") & Mid(YP(I, 1), 1, 3)

        Create_Summary(grdICTITEMI, "ITEM_CODE", "Count")


        Create_Summary(grdDPTPROJI, "ITEM_CODE", "Count")
        Create_Summary(grdDPTPROJI, "QTY_PROJ")
        Create_Summary(grdDPTPROJI, "AMT_PROJ")
        Create_Summary(grdDPTPROJI, "QTY_SHIP")
        Create_Summary(grdDPTPROJI, "AMT_SHIP")

        With grdDPTPROJA.DisplayLayout.Bands("DPTPROJA")
            .Groups("LINE_DESC").Header.Fixed = True
        End With

        With grdDPTPROJI.DisplayLayout.Bands("DPTPROJI")
            .Columns("ITEM_CODE").Header.Fixed = True
            .Columns("ITEM_DESC").Header.Fixed = True
        End With

        Dim YEARs As New List(Of String)
        For Y As Integer = Val(Now.Year) - 1 To Val(Now.Year + 1)
            YEARs.Add(Format(Y, "0000"))
        Next
        Absx1.cbeFor("OPS_YYYY").DataSource = YEARs

        cmbVersion.DataSource = dst.Tables("DPTFCSTV")

        ASCMAIN1.sql = "Select ITEM_CATGY_CODE, ITEM_CATGY_DESC from ICTCATG1 where ITEM_CATGY_CODE in ('C','E','N','I')"
        optITEM_CATGY_CODE.ValueList = ASCMAIN1.ValueListFor("ITEM_CATGY_CODE", , New String() {":", "*:All", "$:All xInactive"}, ASCMAIN1.sql)

        For Each COLUMN_NAME As String In New String() _
        {"ITEM_CODE", "ITEM_DESC", "ITEM_CATGY_CODE" _
        , "ITEM_CLASS_CODE", "COLLECTION_CODE", "VEND_CODE", "VEND_CODE_2" _
        , "PROJ_QOH_BEG", "QTY_OPO", "PROJ_QOH_END"}
            With grdDPTFCSTD.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                If COLUMN_NAME = "PROJ_QOH_BEG" Or COLUMN_NAME = "PROJ_QOH_END" Then
                    .CellAppearance.BackColor = Drawing.Color.LightBlue
                ElseIf COLUMN_NAME = "QTY_OPO" Then
                    .CellAppearance.BackColor = Drawing.Color.LightGreen
                Else
                    .CellAppearance.BackColor = Drawing.Color.Beige
                End If

                .CellActivation = UltraWinGrid.Activation.NoEdit
            End With
        Next

        grdDPTFCSTD.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        grdDPTFCSTD.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

        grdDPTFCSTD.DisplayLayout.Override.FilterOperatorLocation = UltraWinGrid.FilterOperatorLocation.WithOperand

        Format_grdDPTFCSTD()

        ASCMAIN1.Add_Value_List(grdDPTFCSTD, "ITEM_CATGY_CODE")
        ASCMAIN1.Add_Value_List(grdDPTPROJI, "ITEM_CATGY_CODE")
        ASCMAIN1.Add_Value_List(grdICTITEMI, "ITEM_CATGY_CODE")
        ASCMAIN1.Add_Value_List(grdICTITEMI, "ITEM_CATGY_CODE_CURR", "Select ITEM_CATGY_CODE, ITEM_CATGY_DESC from ICTCATG1")

        grdDPTFCSTD.DisplayLayout.Bands(0).Columns("ITEM_CODE").CellButtonAppearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "CAMERA2")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                Validate_Code("BRAND_CODE")
                Validate_Code("MARKET_CODE")

                If cmbOPS_YYYY.Text = "" Then
                    EMsg &= vbCr & "You Must Specify a Year"
                End If

                If cmbVersion.Text = "" Then
                    EMsg &= vbCr & "You Must Specify a Version"
                End If

                VERSION_to_copy = ""
                If EMsg = "" Then
                    If Absx1.cbeFor("VERSION").Value = "L" Then
                        Load_Versions(True)

                        rowDPTPROJ1 = Fill_Record("DPTPROJ1", New String() _
                        {Absx1.txtFor("BRAND_CODE").Text _
                         , Absx1.txtFor("MARKET_CODE").Text _
                         , cmbOPS_YYYY.Text _
                         , optSEASON.Value _
                         , cmbVersion.Value}, True)

                        If rowDPTPROJ1.Item("VERSION_STATUS") & "" = "L" Then
                            If MsgBox("You are calling up the Latest Version, which is Locked" _
                                      & vbCrLf & vbCrLf & "Do you want to create the next version to make revisions?" _
                                       , MsgBoxStyle.YesNo, "Option to Create Next Version") = MsgBoxResult.Yes Then
                                Dim VERSION As String = cmbVersion.Value
                                VERSION_to_copy = VERSION
                                If VERSION = "$$" Then
                                    VERSION = "00"
                                    dst.Tables("DPTFCSTV").Rows.Add(New String() {VERSION, "Original"})
                                Else
                                    VERSION = Format(Val(VERSION) + 1, "00")
                                    dst.Tables("DPTFCSTV").Rows.Add(New String() {VERSION, "Rev #" & VERSION})
                                End If
                                cmbVersion.Value = VERSION
                            End If
                        End If
                    End If
                End If


                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("DPTFCSTD", Absx1.txtFor("BRAND_CODE").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Update"

                Dim SQL As String = "ISNULL(ITEM_CATGY_CODE,'?') <> 'C' AND ISNULL(ITEM_CATGY_CODE,'?') <> 'N' AND ISNULL(ITEM_CATGY_CODE,'?') <> 'E' AND PROJ_P0 <> 0"
                Dim PROJ_INACTIVE As Int32 = Val(dst.Tables("DPTFCSTD").Compute("SUM(PROJ_P0)", SQL) & "")
                If PROJ_INACTIVE <> 0 Then
                    EMsg &= vbCr & "There are Inactive Items with Non-Zero Projection Qty's"
                End If

                SQL = "PROJ_P1 < 0 or PROJ_P2 < 0 or PROJ_P3 < 0 or PROJ_P4 < 0 or PROJ_P5 < 0 or PROJ_P6 < 0"
                Dim NEG_PROJS As Int32 = Val(dst.Tables("DPTFCSTD").Compute("Count(ITEM_CODE)", SQL) & "")
                If NEG_PROJS <> 0 Then
                    EMsg &= vbCr & "There are " & CStr(NEG_PROJS) & " Items with Negative Projection Qty's"
                End If

                Dim WARNING As String = ""
                For Each ITEM_CATGY_CODE As String In New String() {"E", "C", "N"}
                    Dim T As Int64 = Val(dst.Tables("DPTFCSTD").Compute("SUM (PROJ_P0)", "") & "")
                    If T = 0 Then
                        WARNING &= vbCr & "No Projections loaded for Category " & ITEM_CATGY_CODE
                    End If
                Next
                If WARNING <> "" Then
                    If MsgBox(Mid(WARNING, 2) & vbCr & vbCr & "OK to Continue with Update?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

                'If ASCMAIN1.Running_in_VS Then
                '    'Stop
                'Else
                '    EMsg &= vbCr & "Not Yet"
                'End If

                If chkLockStatus.Checked Then
                    If MsgBox("You are updating a Projection AND Locking it." _
                           & vbCrLf & vbCrLf & "This means:" _
                           & vbCrLf & "1) You may not make any further changes to this Version" _
                           & vbCrLf & "2) All Categories (Core, New & Existing) will be Locked for this Version" _
                           & vbCrLf & vbCrLf & "OK to Continue with Locking upon Update?", _
                           MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                Else
                    If MsgBox("You are updating a Projection and NOT Locking it." _
                           & vbCrLf & vbCrLf & "This means:" _
                           & vbCrLf & "1) You may still make changes to this Version (before Locking it)" _
                           & vbCrLf & "2) None of the changes made to this Version will appear in Demand Planning (until you Lock it)" _
                           & vbCrLf & vbCrLf & "OK to Continue with Updating and NOT Locking?", _
                           MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If



            Case "Delete"

                If MsgBox("You are about to DELETE projections for ALL Items in this version" _
                          & vbCrLf & vbCrLf & "Are you sure that you want to Proceed?" _
                          , MsgBoxStyle.YesNo, "Verification - Deleting Entire Projection") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Save to Projections"

                For Each rowICTITEMI As DataRow In dst.Tables("ICTITEMI").Select("", "", DataViewRowState.CurrentRows)
                    Dim ITEM_CODE As String = rowICTITEMI.Item("ITEM_CODE")
                    Dim ITEM_CATGY_CODE As String = rowICTITEMI.Item("ITEM_CATGY_CODE")
                    Dim rowDPTFCSTD As DataRow = dst.Tables("DPTFCSTD").Rows.Find(ITEM_CODE)
                    If rowDPTFCSTD Is Nothing Then
                        EMsg &= "Invalid Item (" & ITEM_CODE & ")"
                    End If
                    If ITEM_CATGY_CODE <> "C" And ITEM_CATGY_CODE <> "N" And ITEM_CATGY_CODE <> "E" And ITEM_CATGY_CODE <> "I" Then
                        EMsg &= "Invalid Category Code (" & ITEM_CATGY_CODE & ") for Item " & ITEM_CODE
                    End If
                Next


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

            Case "Update"
                Call Update_Record()
                Call Mode_Settings(False)

            Case "Delete"
                Call Delete_Record()
                Call Mode_Settings(False)

            Case "Cancel", "Done"
                Call Mode_Settings(False)

            Case "Load from Excel"

                dst.Tables("ICTITEMI").Rows.Clear()

                If Excel_Import(grdICTITEMI) <> -1 Then
                    For Each rowICTITEMI As DataRow In dst.Tables("ICTITEMI").Rows
                        Dim ITEM_CODE As String = rowICTITEMI.Item("ITEM_CODE")
                        If ITEM_CODE = "X" Then Stop
                        Dim rowDPTFCSTD As DataRow = dst.Tables("DPTFCSTD").Rows.Find(ITEM_CODE)
                        If rowDPTFCSTD Is Nothing Then
                            rowICTITEMI.Item("ITEM_CATGY_CODE_CURR") = ""
                        Else
                            rowICTITEMI.Item("ITEM_CATGY_CODE_CURR") = rowDPTFCSTD.Item("ITEM_CATGY_CODE")
                        End If
                    Next
                    grdICTITEMI.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
                End If

            Case "Save to Projections"

                For Each rowICTITEMI As DataRow In dst.Tables("ICTITEMI").Select("ISNULL(BAD_ITEM,'0') <> '1'", "", DataViewRowState.CurrentRows)
                    Dim ITEM_CODE As String = rowICTITEMI.Item("ITEM_CODE")
                    Dim rowDPTFCSTD As DataRow = dst.Tables("DPTFCSTD").Rows.Find(ITEM_CODE)
                    If rowDPTFCSTD IsNot Nothing Then
                        rowDPTFCSTD.Item("ITEM_CATGY_CODE") = rowICTITEMI.Item("ITEM_CATGY_CODE")
                    Else
                        ' ADD ROW?
                    End If
                Next

                MsgBox("The Category Codes from this grid have been loaded into the Projections Table." & vbCr & vbCr & "However, you must Update the Projections in order to make these changes Permanent", MsgBoxStyle.OkOnly, "Save Complete")

            Case "Proj vs Item Master"

                With grdICTITEMI.DisplayLayout.Bands(0)
                    .Columns("ITEM_CATGY_CODE").Header.Caption = "Catgy (Proj)"
                    .Columns("ITEM_CATGY_CODE").CellAppearance.BackColor = Color.LightBlue
                    .Columns("ITEM_CATGY_CODE_CURR").Header.Caption = "Catgy (Item)"
                    .Columns("ITEM_CATGY_CODE_CURR").CellActivation = UltraWinGrid.Activation.NoEdit
                End With

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading Items")
                dst.Tables("ICTITEMI").Rows.Clear()
                For Each rowDPTFCSTD As DataRow In dst.Tables("DPTFCSTD").Rows
                    Dim rowICTITEMI As DataRow = dst.Tables("ICTITEMI").NewRow
                    rowICTITEMI.Item("ITEM_CODE") = rowDPTFCSTD.Item("ITEM_CODE")
                    rowICTITEMI.Item("ITEM_DESC") = rowDPTFCSTD.Item("ITEM_DESC")
                    rowICTITEMI.Item("ITEM_CATGY_CODE") = rowDPTFCSTD.Item("ITEM_CATGY_CODE")
                    'Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", rowDPTFCSTD.Item("ITEM_CODE"))
                    'rowICTITEMI.Item("ITEM_CATGY_CODE_CURR") = rowICTITEM1.Item("ITEM_CATGY_CODE")
                    rowICTITEMI.Item("ITEM_CATGY_CODE_CURR") = rowDPTFCSTD.Item("ITEM_CATGY_CODE_CURR")
                    dst.Tables("ICTITEMI").Rows.Add(rowICTITEMI)
                Next
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                If ScreenMode Then
                    If rowDPTPROJ1.Item("VERSION_STATUS") & "" <> "L" Then
                        .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                        If rowDPTPROJ1.RowState <> DataRowState.Added Then
                            .Groups("Screen Control").Items("Delete").Settings.Enabled = iScreenMode
                        End If
                    End If
                Else
                    .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                    .Groups("Screen Control").Items("Delete").Settings.Enabled = iScreenMode
                End If
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                Setup_tabMain()
                '.Groups("Display Options").Visible = ScreenMode
                '.Groups("Projection Status").Visible = ScreenMode
                '.Groups("Prorate Sales Using").Visible = ScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)


        chkLockStatus.Checked = False

        grdDPTFCSTD.Visible = ScreenMode
        tabMain.Visible = ScreenMode
        Setup_tabMain()
        If ScreenMode Then
            chkLockStatus.Enabled = (VERSION_STATUS <> "L")
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        dst.Tables("DPTFCSTD").Rows.Clear()
        dst.Tables("DPTFCSTI").Rows.Clear()
        dst.Tables("DPTPROJI").Rows.Clear()
        'dst.Tables("DPTPROJA").Rows.Clear()
        dst.Tables("DPTPROJ0").Rows.Clear()
        dst.Tables("DPTPROJ1").Rows.Clear()
        dst.Tables("DPTPROJ2").Rows.Clear()
        dst.Tables("DPTITMF1").Rows.Clear()
        dst.Tables("ICTITEMI").Rows.Clear()
        EnforceConstraints(True)

        Absx1.txtFor("BRAND_CODE").Text = ""
        Absx1.txtFor("MARKET_CODE").Text = ""
        cmbVersion.Value = DBNull.Value
        numPRORATE.Value = 0
        If Absx1.cbeFor("OPS_YYYY").Value & "" = "" Then
            Absx1.cbeFor("OPS_YYYY").Value = Now.Year
        End If

        negatives_only = False
        tabMain.Tabs("grd").Visible = False

        numPRORATE.Value = 0
        PictureBox1.Image = Nothing
        UltraExplorerBar1.Groups("Picture").Text = "Picture"
    End Sub

    Sub Load_Record()

        Call ASCMAIN1.Progress("Now Loading Data ...")

        Call Save_Header_Fields(UltraGroupBox1)

        OPS_YYYY = Absx1.cbeFor("OPS_YYYY").Text

        ReDim YP(6, 1)
        ReDim YP_LY(6, 1)
        For P As Integer = 1 To 6
            Dim YPX As String = ASCMAIN1.Period_Calc(OPS_YYYY & Format(P + IIf(optSEASON.Value = "F", 6, 0), "00"), -1 * ASCMAIN1.PCO + 1)
            YP(P, 0) = YPX
            YP(P, 1) = ASCMAIN1.Get_Legend(YPX, False, True)

            Dim YPX_LY As String = ASCMAIN1.Period_Calc(YPX, -12)
            YP_LY(P, 0) = YPX_LY
            YP_LY(P, 1) = ASCMAIN1.Get_Legend(YPX_LY, False, True)

            With grdDPTFCSTD.DisplayLayout.Bands(0).Groups("P" & Format(P, "0"))
                .Header.Caption = Mid(YP(P, 1), 1, 3)
                '.Width = 60
                '.CellAppearance.BackColor = Color.Yellow
            End With

            With grdDPTFCSTI.DisplayLayout.Bands(0).Columns("PROJ_P" & Format(P, "0"))
                .Header.Caption = Mid(YP(P, 1), 1, 3)
                .Width = 60
                '.CellAppearance.BackColor = Color.Yellow
            End With

            With grdDPTPROJI.DisplayLayout.Bands(0)
                .Columns("F" & Format(P, "00")).Header.Caption = "Prj " & Mid(YP(P, 1), 1, 3)
                .Columns("A" & Format(P, "00")).Header.Caption = "Act " & Mid(YP(P, 1), 1, 3)
            End With

            With grdDPTFCSTD.DisplayLayout.Bands(0)
                .Columns("QTY_FC" & Format(P, "0")).Header.Caption = YP(P, 1)
                With .Columns("PROJ_P" & Format(P, "0"))
                    If YPX >= ASCMAIN1.CYP Then
                        .CellActivation = UltraWinGrid.Activation.AllowEdit
                        .CellAppearance.BackColor = Color.AliceBlue
                    Else
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                        .CellAppearance.BackColor = Color.Beige
                    End If
                End With
            End With
        Next

        With grdDPTFCSTI.DisplayLayout.Bands(0).Columns("PROJ_P0")
            .Header.Caption = "Total"
            .Width = 60
            .CellAppearance.BackColor = Color.Beige
        End With


        ASCDATA1.ExecuteSQL("Truncate Table " & DPTFCSTD)

        Dim SQLX As String = "SELECT X.ITEM_CODE, 'XXXXXXXXXX' DATA_TYPE" _
        & ", SUM (DECODE(X.ORDR_YYYYPP_UPDATED,'YP1',X.ORDR_QTY_SHIP,0)) Q1" _
        & ", SUM (DECODE(X.ORDR_YYYYPP_UPDATED,'YP2',X.ORDR_QTY_SHIP,0)) Q2" _
        & ", SUM (DECODE(X.ORDR_YYYYPP_UPDATED,'YP3',X.ORDR_QTY_SHIP,0)) Q3" _
        & ", SUM (DECODE(X.ORDR_YYYYPP_UPDATED,'YP4',X.ORDR_QTY_SHIP,0)) Q4" _
        & ", SUM (DECODE(X.ORDR_YYYYPP_UPDATED,'YP5',X.ORDR_QTY_SHIP,0)) Q5" _
        & ", SUM (DECODE(X.ORDR_YYYYPP_UPDATED,'YP6',X.ORDR_QTY_SHIP,0)) Q6" _
        & " from SOTINVH2 X, ICTITEM1, ARTCUST1" _
        & " where X.ORDR_YYYYPP_UPDATED BETWEEN 'YP1' AND 'YP6'" _
        & " and ICTITEM1.ITEM_CODE = X.ITEM_CODE" _
        & " and ICTITEM1.COLLECTION_CODE in " _
        & " (Select COLLECTION_CODE from ICTCOLL1 " _
        & "  where BRAND_CODE = '" & Absx1.txtFor("BRAND_CODE").Text & "')" _
        & " and ARTCUST1.CUST_CODE = X.CUST_CODE" _
        & " and ARTCUST1.TRADE_CLASS_CODE in " _
        & " (Select TRADE_CLASS_CODE from SOTTCLS1 " _
        & "  where MARKET_CODE = '" & Absx1.txtFor("MARKET_CODE").Text & "')" _
        & " group by X.ITEM_CODE"

        For Each DATA_TYPE As String In New String() {"TY_IN", "LY_IN", "TY_ST", "LY_ST"}
            Dim SQL As String = Replace(SQLX, "XXXXXXXXXX", DATA_TYPE)
            For P As Integer = 1 To 6
                Dim YPX As String = ""
                If DATA_TYPE Like "TY*" Then
                    YPX = YP(P, 0)
                Else
                    YPX = YP_LY(P, 0)
                End If
                If DATA_TYPE Like "*ST" Then
                    SQL = Replace(SQL, "SOTINVH2", "RSTRETL1")
                    SQL = Replace(SQL, "ORDR_YYYYPP_UPDATED", "OPS_YYYYPP")
                    SQL = Replace(SQL, "ORDR_QTY_SHIP", "QTY_SOLD")
                End If
                SQL = Replace(SQL, "YP" & Format(P, "0"), YPX)
            Next
            ASCDATA1.ExecuteSQL("Insert into " & DPTFCSTD & " " & SQL)
        Next

        ASCDATA1.ExecuteSQL("Truncate Table " & ICTITEMS)
        ASCMAIN1.sql = "Insert into " & ICTITEMS _
        & " (ITEM_CODE, ITEM_DESC, ITEM_CATGY_CODE, ITEM_CATGY_CODE_CURR" _
        & ", ITEM_CLASS_CODE, COLLECTION_CODE" _
        & ", DEPT_CODE, VEND_CODE, VEND_CODE_2" _
        & ", ITEM_RETAIL_PRICE, ITEM_PRICE, ITEM_COST_STD, ITEM_PICTURE_FILENAME) " _
        & " Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC" _
        & ", NVL(DPTPROJ0.ITEM_CATGY_CODE,ICTITEM1.ITEM_CATGY_CODE) ITEM_CATGY_CODE" _
        & ", ICTITEM1.ITEM_CATGY_CODE ITEM_CATGY_CODE_CURR " _
        & ", ICTITEM1.ITEM_CLASS_CODE, ICTITEM1.COLLECTION_CODE, ICTITEM1.DEPT_CODE" _
        & ", ICTITEM1.VEND_CODE, ICTITEM1.VEND_CODE VEND_CODE_2" _
        & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_PRICE, ICTITEM1.ITEM_COST_STD " _
        & ", ICTITEM1.ITEM_PICTURE_FILENAME" _
        & " from ICTITEM1,DPTPROJ0 " _
        & " where NVL(DPTPROJ0.ITEM_CATGY_CODE,ICTITEM1.ITEM_CATGY_CODE) in ('C','N','E','I')" _
        & "   and DPTPROJ0.OPS_YYYY (+) = '" & HFs("OPS_YYYY") & "'" _
        & "   and DPTPROJ0.SEASON (+) = '" & HFs("SEASON") & "'" _
        & "   and DPTPROJ0.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" _
        & "   and ICTITEM1.COLLECTION_CODE in " _
        & " (Select COLLECTION_CODE from ICTCOLL1 " _
        & " where BRAND_CODE = '" & Absx1.txtFor("BRAND_CODE").Text & "')"
        ASCDATA1.ExecuteSQL()

        EnforceConstraints(False)

        rowDPTPROJ1 = Fill_Record("DPTPROJ1", New String() _
        {HFs("BRAND_CODE"), HFs("MARKET_CODE"), HFs("OPS_YYYY"), optSEASON.Value, cmbVersion.Value}, True)

        If VERSION_to_copy = "" Then
            Fill_Records("DPTPROJ2", New String() _
            {HFs("BRAND_CODE"), HFs("MARKET_CODE"), HFs("OPS_YYYY"), optSEASON.Value, cmbVersion.Value})
        Else
            Fill_Records("DPTPROJ2", New String() _
            {HFs("BRAND_CODE"), HFs("MARKET_CODE"), HFs("OPS_YYYY"), optSEASON.Value, VERSION_to_copy})
            For Each rowDPTPROJ2 As DataRow In dst.Tables("DPTPROJ2").Rows
                rowDPTPROJ2.Item("VERSION") = cmbVersion.Value
            Next
        End If

        If rowDPTPROJ1.RowState = DataRowState.Added Then
            If rowDPTPROJ1.Item("VERSION") = "$$" Then
                rowDPTPROJ1.Item("VERSION_DESC") = "Sales"
            ElseIf rowDPTPROJ1.Item("VERSION") = "00" Then
                rowDPTPROJ1.Item("VERSION_DESC") = "Original"
            Else
                rowDPTPROJ1.Item("VERSION_DESC") = "Rev #" & rowDPTPROJ1.Item("VERSION")
            End If
        End If

        VERSION_STATUS = rowDPTPROJ1.Item("VERSION_STATUS") & ""

        dst.Tables("DPTFCSTD").Columns("QTY_FC").Expression = ""
        dst.Tables("DPTFCSTD").Columns("AMT_FC").Expression = ""

        Fill_Records("DPTFCSTD", New String() _
        {HFs("BRAND_CODE"), HFs("MARKET_CODE"), HFs("OPS_YYYY"), optSEASON.Value, _
         IIf(VERSION_to_copy <> "", VERSION_to_copy, _
              cmbVersion.Value)})

        dst.Tables("DPTFCSTD").Columns("QTY_FC").Expression = "PROJ_P0"
        dst.Tables("DPTFCSTD").Columns("AMT_FC").Expression = "PROJ_P0 * ITEM_PRICE"


        ASCDATA1.ExecuteSQL("Truncate Table " & DPTPROJI)
        ASCDATA1.ExecuteSQL("Insert into " & DPTPROJI & " " & sqlDPTPROJI, "VVVVVV", New String() {YP(1, 0), YP(2, 0), YP(3, 0), YP(4, 0), YP(5, 0), YP(6, 0)})

        Fill_Records("DPTPROJI")

        For Each rowDPTPROJA As DataRow In dst.Tables("DPTPROJA").Rows
            Dim PCT As Int32 = Val(rowDPTPROJA.Item("PCT") & "")
            For Each ITEM_CATGY_CODE As String In New String() {"N", "C", "E", "I"}
                Dim sqlw As String = "ITEM_CATGY_CODE = '" & ITEM_CATGY_CODE & "'"
                If PCT = 100 Then
                Else
                    sqlw &= "and VAR_PCT >= " & CStr(-1 * PCT) & " AND VAR_PCT <= " & CStr(PCT)
                End If

                sqlw &= " and QTY_PROJ >= 25"

                For Each rowDPTPROJI As DataRow In dst.Tables("DPTPROJI").Select(sqlw & " AND ISNULL(PCT,0) = 0")
                    rowDPTPROJI.Item("PCT") = PCT
                Next
                Dim ITEMS = dst.Tables("DPTPROJI").Compute("COUNT(ITEM_CODE)", sqlw)
                Dim QTY_PROJ = dst.Tables("DPTPROJI").Compute("SUM(QTY_PROJ)", sqlw)
                Dim QTY_SHIP = dst.Tables("DPTPROJI").Compute("SUM(QTY_SHIP)", sqlw)
                Dim AMT_PROJ = dst.Tables("DPTPROJI").Compute("SUM(AMT_PROJ)", sqlw)
                Dim AMT_SHIP = dst.Tables("DPTPROJI").Compute("SUM(AMT_SHIP)", sqlw)
                rowDPTPROJA.Item(ITEM_CATGY_CODE & "_ITEMS") = ITEMS
                rowDPTPROJA.Item(ITEM_CATGY_CODE & "_QTY_PROJ") = QTY_PROJ
                rowDPTPROJA.Item(ITEM_CATGY_CODE & "_QTY_SHIP") = QTY_SHIP
                rowDPTPROJA.Item(ITEM_CATGY_CODE & "_AMT_PROJ") = AMT_PROJ
                rowDPTPROJA.Item(ITEM_CATGY_CODE & "_AMT_SHIP") = AMT_SHIP
            Next
        Next

        For Each ITEM_CATGY_CODE As String In New String() {"N", "C", "E", "I"}
            Dim TOTAL As Int32 = Val(dst.Tables("DPTPROJA").Compute("MAX(" & ITEM_CATGY_CODE & "_ITEMS)", "") & "")
            Dim EXPRESSION As String = ""
            If TOTAL = 0 Then
            Else
                EXPRESSION = "100 * ISNULL(" & ITEM_CATGY_CODE & "_ITEMS,0) / " & CStr(TOTAL)
            End If
            dst.Tables("DPTPROJA").Columns(ITEM_CATGY_CODE & "_PCT").Expression = EXPRESSION

            For Each QA As String In New String() {"QTY", "AMT"}
                For Each PS As String In New String() {"PROJ", "SHIP"}
                    Dim DC As String = "_" & QA & "_" & PS
                    TOTAL = Val(dst.Tables("DPTPROJA").Compute("MAX(" & ITEM_CATGY_CODE & DC & ")", "") & "")
                    EXPRESSION = ""
                    If TOTAL = 0 Then
                    Else
                        EXPRESSION = "100 * ISNULL(" & ITEM_CATGY_CODE & DC & ",0) / " & CStr(TOTAL)
                    End If
                    dst.Tables("DPTPROJA").Columns(ITEM_CATGY_CODE & DC & "_PCT").Expression = EXPRESSION
                Next
            Next
        Next

        EnforceConstraints(True)

        Sort_grdColumns(grdDPTFCSTD, "ITEM_CODE")
        optITEM_CATGY_CODE.Value = "E"

        optCOLLECTION_CODE.Value = "A"

        ASCMAIN1.sql = "Select COLLECTION_CODE, COLLECTION_NAME from ICTCOLL1 where BRAND_CODE = '" & Absx1.txtFor("BRAND_CODE").Text & "' order BY COLLECTION_CODE"
        cbeCOLLECTION_CODE.DataSource = ASCDATA1.GetDataTable
        cbeCOLLECTION_CODE.Value = cbeCOLLECTION_CODE.Items(0)

        Setup_grdDPTFCSTD()
        lblStatus.Text = "Status: " & IIf(VERSION_STATUS = "L", "Locked", "In-Process")

        splDPTFCSTD.Panel2Collapsed = Not chkVersion.Checked


        If cmbVersion.Value = "$$" Then
            chkVersion.Enabled = False
            chkVersion.Text = "None Prior"
        Else
            chkVersion.Enabled = True
            chkVersion.Text = "Show Prior Versions"
        End If

        ASCMAIN1.Add_Value_List(grdDPTFCSTI, "VERSION", "Select VERSION, VERSION || '-' || VERSION_DESC VERSION_DESC from DPTPROJ1 where BRAND_CODE = '" & HFs("BRAND_CODE") & "' and MARKET_CODE = '" & HFs("MARKET_CODE") & "' and OPS_YYYY = '" & cmbOPS_YYYY.Value & "' and SEASON = '" & optSEASON.Value & "'")


        Call ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Call BeginTrans()

        Dim sql_Delete As String = ""
        'sql_Delete = "Delete from DPTPROJ1" _
        '    & " where BRAND_CODE = '" & Absx1.txtFor("BRAND_CODE").Text & "'" _
        '    & " and MARKET_CODE = '" & Absx1.txtFor("MARKET_CODE").Text & "'" _
        '    & " and OPS_YYYY = '" & Absx1.cbeFor("OPS_YYYY").Value & "'" _
        '    & " and SEASON = '" & Absx1.cbeFor("SEASON").Value & "'" _
        '    & " and VERSION = '" & Absx1.cbeFor("VERSION").Value & "'"

        'Call Update_Record_TDA("DPTPROJ1", sql_Delete)

        dst.Tables("DPTPROJ0").Rows.Clear()
        dst.Tables("DPTPROJ2").Rows.Clear()
        dst.Tables("DPTITMF1").Rows.Clear()

        For Each rowDPTFCSTD As DataRow In dst.Tables("DPTFCSTD").Select _
        ("(ITEM_CATGY_CODE = 'C' OR ITEM_CATGY_CODE = 'E' OR ITEM_CATGY_CODE = 'N') AND PROJ_P0 <> 0", "", DataViewRowState.CurrentRows)
            Dim rowDPTPROJ0 As DataRow = dst.Tables("DPTPROJ0").NewRow
            rowDPTPROJ0.Item("OPS_YYYY") = HFs("OPS_YYYY")
            rowDPTPROJ0.Item("SEASON") = HFs("SEASON")
            rowDPTPROJ0.Item("ITEM_CODE") = rowDPTFCSTD.Item("ITEM_CODE")
            rowDPTPROJ0.Item("ITEM_CATGY_CODE") = rowDPTFCSTD.Item("ITEM_CATGY_CODE")
            dst.Tables("DPTPROJ0").Rows.Add(rowDPTPROJ0)

            Dim rowDPTPROJ2 As DataRow = dst.Tables("DPTPROJ2").NewRow
            rowDPTPROJ2.Item("BRAND_CODE") = HFs("BRAND_CODE")
            rowDPTPROJ2.Item("MARKET_CODE") = HFs("MARKET_CODE")
            rowDPTPROJ2.Item("OPS_YYYY") = HFs("OPS_YYYY")
            rowDPTPROJ2.Item("SEASON") = HFs("SEASON")
            rowDPTPROJ2.Item("VERSION") = HFs("VERSION")
            rowDPTPROJ2.Item("ITEM_CODE") = rowDPTFCSTD.Item("ITEM_CODE")
            rowDPTPROJ2.Item("PROJ_P1") = rowDPTFCSTD.Item("PROJ_P1")
            rowDPTPROJ2.Item("PROJ_P2") = rowDPTFCSTD.Item("PROJ_P2")
            rowDPTPROJ2.Item("PROJ_P3") = rowDPTFCSTD.Item("PROJ_P3")
            rowDPTPROJ2.Item("PROJ_P4") = rowDPTFCSTD.Item("PROJ_P4")
            rowDPTPROJ2.Item("PROJ_P5") = rowDPTFCSTD.Item("PROJ_P5")
            rowDPTPROJ2.Item("PROJ_P6") = rowDPTFCSTD.Item("PROJ_P6")
            dst.Tables("DPTPROJ2").Rows.Add(rowDPTPROJ2)

            If chkLockStatus.Checked Then
                For P As Integer = 1 To 6
                    Dim rowDPTITMF1 As DataRow = dst.Tables("DPTITMF1").NewRow
                    rowDPTITMF1.Item("ITEM_CODE") = rowDPTFCSTD.Item("ITEM_CODE")
                    rowDPTITMF1.Item("MARKET_CODE") = HFs("MARKET_CODE")
                    rowDPTITMF1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
                    rowDPTITMF1.Item("OPS_YYYYPP_FC") = YP(P, 0)
                    rowDPTITMF1.Item("FORECAST") = rowDPTFCSTD.Item("PROJ_P" & CStr(P))
                    dst.Tables("DPTITMF1").Rows.Add(rowDPTITMF1)
                Next
            End If
        Next

        sql_Delete = "Delete from DPTPROJ0" _
            & " where OPS_YYYY = '" & HFs("OPS_YYYY") & "'" _
            & " and SEASON = '" & HFs("SEASON") & "'"
        Call Update_Record_TDA("DPTPROJ0", sql_Delete)

        sql_Delete = "Delete from DPTPROJ2" _
            & " where BRAND_CODE = '" & HFs("BRAND_CODE") & "'" _
            & " and MARKET_CODE = '" & HFs("MARKET_CODE") & "'" _
            & " and OPS_YYYY = '" & HFs("OPS_YYYY") & "'" _
            & " and SEASON = '" & HFs("SEASON") & "'" _
            & " and VERSION = '" & HFs("VERSION") & "'"
        Call Update_Record_TDA("DPTPROJ2", sql_Delete)

        If chkLockStatus.Checked Then
            rowDPTPROJ1.Item("VERSION_STATUS") = "L"

            sql_Delete = "Delete from DPTITMF1" _
                & " where MARKET_CODE = '" & HFs("MARKET_CODE") & "'" _
                & " and OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" _
                & " and OPS_YYYYPP_FC between '" & YP(1, 0) & "' and '" & YP(6, 0) & "'"
            Call Update_Record_TDA("DPTITMF1", sql_Delete)

        End If
        Call INIT_LAST("DPTPROJ1", True)
        Call Update_Record_TDA("DPTPROJ1")


        Call CommitTrans("Update Complete")

    End Sub

    Sub Delete_Record()

        Call BeginTrans()

        ASCMAIN1.sql = "Delete from DPTPROJ1" _
            & " where BRAND_CODE = '" & HFs("BRAND_CODE") & "'" _
            & " and MARKET_CODE = '" & HFs("MARKET_CODE") & "'" _
            & " and OPS_YYYY = '" & HFs("OPS_YYYY") & "'" _
            & " and SEASON = '" & HFs("SEASON") & "'" _
            & " and VERSION = '" & HFs("VERSION") & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Delete from DPTPROJ2" _
            & " where BRAND_CODE = '" & HFs("BRAND_CODE") & "'" _
            & " and MARKET_CODE = '" & HFs("MARKET_CODE") & "'" _
            & " and OPS_YYYY = '" & HFs("OPS_YYYY") & "'" _
            & " and SEASON = '" & HFs("SEASON") & "'" _
            & " and VERSION = '" & HFs("VERSION") & "'"
        ASCDATA1.ExecuteSQL()

        Call CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdDPTPROJI, "SS", "Show Filter", "Show GroupBox")
        Call Load_Popup_Menu(grdDPTPROJA, "S", "Show Extended Data")

        Call Load_Popup_Menu(grdICTITEMI, "SSSB", "Show Filter", "Show GroupBox", "Show Differences Only", "Load from Spreadsheet")

        grdDPTFCSTD_HCs = New String() {"ITEM_DESC" _
                             , "COLLECTION_CODE" _
                             , "ITEM_CATGY_CODE" _
                             , "ITEM_CLASS_CODE" _
                             , "DEPT_CODE" _
                             , "VEND_CODE" _
                             , "VEND_CODE_2" _
                             , "ITEM_RETAIL_PRICE" _
                             , "ITEM_PRICE" _
                             , "ITEM_COST_STD"}


        Call Load_Popup_Menu(grdDPTFCSTD, "SSSSBBBS" & "SSSSSSSSSS" _
                             , "Show Filter", "Group by Category", "Group by Collection", "Show Season Totals", "Load from Spreadsheet" _
                             , "Clear All Projections", "Clear Projections for XXX", "Show Negatives Only" _
                             , "ITEM_DESC|Show Description" _
                             , "COLLECTION_CODE|Show Collection" _
                             , "ITEM_CATGY_CODE|Show Catgy" _
                             , "ITEM_CLASS_CODE|Show Class" _
                             , "DEPT_CODE|Show Dept" _
                             , "VEND_CODE|Show Vendor" _
                             , "VEND_CODE_2|Show Vendor2" _
                             , "ITEM_RETAIL_PRICE|Show Retail" _
                             , "ITEM_PRICE|Show Price" _
                             , "ITEM_COST_STD|Show Cost")

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

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
        End If
        If tlb_pop.Tools.Exists("Group by Category") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Group by Category"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Tag = "X"
            tlb_sbt.Checked = grdDPTFCSTD.DisplayLayout.Bands(0).SortedColumns.Contains(grdDPTFCSTD.DisplayLayout.Bands(0).Columns("ITEM_CATGY_CODE"))
            tlb_sbt.Tag = ""
        End If
        If tlb_pop.Tools.Exists("Group by Collection") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Group by Collection"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Tag = "X"
            tlb_sbt.Checked = grdDPTFCSTD.DisplayLayout.Bands(0).SortedColumns.Contains(grdDPTFCSTD.DisplayLayout.Bands(0).Columns("COLLECTION_CODE"))
            tlb_sbt.Tag = ""
        End If
        If tlb_pop.Tools.Exists("Show Season Totals") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Season Totals"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Groups("PROJ_QOH_BEG").Hidden
        End If
        If tlb_pop.Tools.Exists("Show Extended Data") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Extended Data"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns("N_QTY_PROJ").Hidden
        End If
        If tlb_pop.Tools.Exists("Show Negatives Only") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Negatives Only"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = negatives_only
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdDPTFCSTD"
                    For Each c As String In grdDPTFCSTD_HCs
                        Dim COL As String = Split(c, "|")(0)
                        tlb_sbt = DirectCast(tlb_pop.Tools(COL), UltraWinToolbars.StateButtonTool)
                        tlb_sbt.Tag = "X"
                        If grdDPTFCSTD.DisplayLayout.Bands(0).Columns(COL).Group.Hidden Then
                            tlb_sbt.Checked = False
                        Else
                            tlb_sbt.Checked = True
                        End If
                        tlb_sbt.Tag = ""
                    Next

                    Dim tlb_btn As UltraWinToolbars.ButtonTool
                    tlb_btn = DirectCast(tlb_pop.Tools("Clear Projections for XXX"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = False
                    If (grd.ActiveRow Is Nothing Or grd.ActiveCell Is Nothing) OrElse grd.ActiveCell.Column.CellActivation = UltraWinGrid.Activation.NoEdit Then
                    Else
                        If New String() {"PROJ_P1", "PROJ_P2", "PROJ_P3", "PROJ_P4", "PROJ_P5", "PROJ_P6"}.Contains(grd.ActiveCell.Column.Key) Then
                            tlb_btn.SharedProps.Visible = True
                            tlb_btn.SharedProps.Caption = "Clear Projections for " & grd.ActiveCell.Column.Group.Header.Caption
                        End If
                    End If
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked

            Case "Group by Category"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag = "X" Then Exit Sub
                If tlb_sbt.Checked Then
                    grdDPTFCSTD.DisplayLayout.Bands(0).SortedColumns.Add("ITEM_CATGY_CODE", False, True)
                Else
                    grdDPTFCSTD.DisplayLayout.Bands(0).SortedColumns.Remove("ITEM_CATGY_CODE")
                End If

            Case "Group by Collection"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag = "X" Then Exit Sub
                If tlb_sbt.Checked Then
                    grdDPTFCSTD.DisplayLayout.Bands(0).SortedColumns.Add("COLLECTION_CODE", False, True)
                Else
                    grdDPTFCSTD.DisplayLayout.Bands(0).SortedColumns.Remove("COLLECTION_CODE")
                End If

            Case "Show Season Totals"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Groups("PROJ_QOH_BEG").Hidden = Not tlb_sbt.Checked
                grd.DisplayLayout.Bands(0).Groups("QTY_OPO").Hidden = Not tlb_sbt.Checked
                grd.DisplayLayout.Bands(0).Groups("QTY_FC").Hidden = Not tlb_sbt.Checked
                grd.DisplayLayout.Bands(0).Groups("AMT_FC").Hidden = Not tlb_sbt.Checked
                grd.DisplayLayout.Bands(0).Groups("PROJ_QOH_END").Hidden = Not tlb_sbt.Checked

            Case "Show Extended Data"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                For Each ITEM_CATGY_CODE As String In New String() {"N", "C", "E", "I"}
                    With grd.DisplayLayout.Bands(0)
                        For Each QA As String In New String() {"QTY", "AMT"}
                            For Each PS As String In New String() {"PROJ", "SHIP"}
                                Dim DC As String = "_" & QA & "_" & PS
                                .Columns(ITEM_CATGY_CODE & DC).Hidden = Not tlb_sbt.Checked
                                .Columns(ITEM_CATGY_CODE & DC & "_PCT").Hidden = Not tlb_sbt.Checked
                            Next
                        Next
                    End With

                Next

            Case "Load from Spreadsheet"
                Excel_Import(grd)

            Case "Show Differences Only"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Dim dvw As DataView = DirectCast(grdICTITEMI.DataSource, DataTable).DefaultView
                If tlb_sbt.Checked Then
                    dvw.RowFilter = "ISNULL(ITEM_CATGY_CODE,'?') <> ISNULL(ITEM_CATGY_CODE_CURR,'?')"
                    ASCMAIN1.Notify("Now Showing Only those Items with Differences")
                Else
                    dvw.RowFilter = ""
                    ASCMAIN1.Notify("Now Showing All Items")
                End If

            Case "Show Negatives Only"
                negatives_only = Not negatives_only
                Setup_grdDPTFCSTD()

            Case "Clear Projections for XXX"
                Dim COLUMN_NAME As String = grd.ActiveCell.Column.Key
                Dim I_START As Int16 = Val(Mid(COLUMN_NAME, Len(COLUMN_NAME), 1))
                Dim XXX As String = grd.DisplayLayout.Bands(0).Columns("PROJ_P" & CStr(I_START)).Group.Header.Caption
                If MsgBox("Clear Projections for " & XXX & "?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Clearing Projections (" & XXX & ")")
                For Each rowDPTFCSTD As DataRow In dst.Tables("DPTFCSTD").Select
                    ASCMAIN1.Progress("-", rowDPTFCSTD.Item("ITEM_CODE"))
                    rowDPTFCSTD.Item("PROJ_P" & CStr(I_START)) = 0
                Next
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
                MsgBox("Projections have been cleared for " & XXX, MsgBoxStyle.OkOnly, "Verification")

            Case "Clear All Projections"
                Dim I_START As Int16 = -1
                For i As Int16 = 1 To 6
                    If grdDPTFCSTD.DisplayLayout.Bands(0).Columns("PROJ_P" & CStr(i)).CellActivation = UltraWinGrid.Activation.AllowEdit Then
                        I_START = i
                        Exit For
                    End If
                Next
                If I_START = -1 Then
                    MsgBox("No Periods are open for Maintenance", MsgBoxStyle.OkOnly, "Clearing Projections Not Permitted")
                Else
                    If MsgBox("Clear Projections for " & grd.DisplayLayout.Bands(0).Columns("PROJ_P" & CStr(I_START)).Group.Header.Caption & " thru " & grd.DisplayLayout.Bands(0).Columns("PROJ_P6").Group.Header.Caption & "?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Now Clearing Projections (All Periods)")
                    For Each rowDPTFCSTD As DataRow In dst.Tables("DPTFCSTD").Select
                        ASCMAIN1.Progress("-", rowDPTFCSTD.Item("ITEM_CODE"))
                        For i As Int16 = 1 To 6
                            rowDPTFCSTD.Item("PROJ_P" & CStr(i)) = 0
                        Next
                    Next
                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("")
                    MsgBox("Projections have been cleared from All Periods (permitting maintenance)", MsgBoxStyle.OkOnly, "Verification")
                End If

        End Select

        Select Case grd.Name
            Case "grdDPTFCSTD"
                If Join(grdDPTFCSTD_HCs, ",").Contains(e.Tool.Key) Then
                    Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                    If tlb_sbt.Tag <> "" Then Exit Sub
                    Dim COL As String = tlb_sbt.Key
                    grd.DisplayLayout.Bands(0).Columns(COL).Group.Hidden = Not tlb_sbt.Checked
                    grd.DisplayLayout.Bands(0).Columns(COL).Hidden = Not tlb_sbt.Checked
                    Exit Sub
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case Else
                'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BRAND_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    'Call Click_Command("Load", e)
                End If
            Case "MARKET_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    'Call Click_Command("Load", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "BRAND_CODE"
                'Call Click_Command("Load")
            Case "MARKET_CODE"
                'Call Click_Command("Load")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

            Case "BRAND_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("BRAND_CODE").Text <> "" Then
                        Call LookUp("ICTBRAN1", Absx1.txtFor("BRAND_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

            Case "MARKET_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("MARKET_CODE").Text <> "" Then
                        Call LookUp("SOTMKTC1", Absx1.txtFor("MARKET_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

        End Select
    End Sub

#End Region

#Region "Excel Upload"

    Overrides Sub Excel_Import_Pre_Process _
    (ByVal grd As UltraWinGrid.UltraGrid, _
     Optional ByRef load_by_table As Boolean = False, _
     Optional ByRef load_handled As Boolean = False, _
     Optional ByRef F As ASFEXCL1 = Nothing)

        If grd.Name = "grdICTITEMI" Then
            load_by_table = True
        End If

        If grd.Name = "grdDPTFCSTD" Then
            'For Each row As DataRow In F.dt.Select
            '    Stop
            'Next

            Dim COLs As New Dictionary(Of String, String)
            For Each gcol As UltraWinGrid.UltraGridColumn In F.grdExcel.DisplayLayout.Bands(0).Columns
                If New String() {"ITEM_CODE", "PROJ_P1", "PROJ_P2", "PROJ_P3", "PROJ_P4", "PROJ_P5", "PROJ_P6"}.Contains(gcol.Tag) Then
                    COLs.Add(gcol.Tag, gcol.Key)
                End If
            Next

            If F.grd.Rows.Count = 0 Then
                MsgBox("Nothing to Upload")
                Exit Sub
            Else
                If COLs.Count <> 7 Then
                    MsgBox("Missing Columns")
                    Exit Sub
                End If
            End If

            Dim tblMissing As New DataTable
            tblMissing.Columns.Add("ITEM_CODE")

            For Each grow As UltraWinGrid.UltraGridRow In F.grdExcel.Rows
                Dim ITEM_CODE As String = grow.Cells(COLs("ITEM_CODE")).Text
                Dim rowDPTFCSTD As DataRow = dst.Tables("DPTFCSTD").Rows.Find(ITEM_CODE)
                If rowDPTFCSTD Is Nothing Then
                    If Val(grow.Cells(COLs("PROJ_P1")).Value & "") <> 0 _
                    Or Val(grow.Cells(COLs("PROJ_P2")).Value & "") <> 0 _
                    Or Val(grow.Cells(COLs("PROJ_P3")).Value & "") <> 0 _
                    Or Val(grow.Cells(COLs("PROJ_P4")).Value & "") <> 0 _
                    Or Val(grow.Cells(COLs("PROJ_P5")).Value & "") <> 0 _
                    Or Val(grow.Cells(COLs("PROJ_P6")).Value & "") <> 0 Then
                        tblMissing.Rows.Add(New String() {ITEM_CODE})
                    End If
                Else
                    For I As Int16 = 1 To 6
                        Dim C As String = "PROJ_P" & CStr(I)
                        If grdDPTFCSTD.DisplayLayout.Bands(0).Columns(C).CellActivation = UltraWinGrid.Activation.AllowEdit Then
                            rowDPTFCSTD.Item(C) = grow.Cells(COLs(C)).Value
                        End If
                    Next
                End If
            Next

            If tblMissing.Rows.Count > 0 Then
                Dim Fmsg As New ASFMSGBF
                Fmsg.Show_grd(tblMissing, ASCMAIN1.ActiveForm, "Items Missing from Current List")
            End If

            MsgBox("Load from Spreadsheet Complete")
            'Load_Projections()
            load_handled = True
        End If

    End Sub

    Overrides Sub Excel_Import_Post_Process(ByVal grd As UltraWinGrid.UltraGrid, F As ASFEXCL1)

        Dim BAD_ITEM_count As Int64 = dst.Tables("ICTITEMI").Select("BAD_ITEM = '1'").Length
        If BAD_ITEM_count <> 0 Then
            MsgBox("There were " & CStr(BAD_ITEM_count) & " Bad Items Loaded from Spreadsheet", MsgBoxStyle.OkOnly, CStr(dst.Tables("ICTITEMI").Rows.Count) & " Records Loaded")
        Else
            MsgBox("All Items Loaded Successfully", MsgBoxStyle.OkOnly, CStr(dst.Tables("ICTITEMI").Rows.Count) & " Records Loaded")
        End If

    End Sub

    Overrides Sub Excel_Import_Custom_Processing_row _
    (ByVal row As DataRow, ByVal grow As UltraWinGrid.UltraGridRow, _
     Optional ByVal TBL As DataTable = Nothing)
        'If optCI.Value = "I" Then
        'Else
        '    row.Item("CUST_CODE") = Absx1.txtFor("CUST_CODE").Text
        'End If
        If Len(row.Item("ITEM_CATGY_CODE") & "") > 1 Then
            row.Item("ITEM_CATGY_CODE") = Mid(row.Item("ITEM_CATGY_CODE"), 1, 1)
        End If

        Dim ITEM_CODE As String = row.Item("ITEM_CODE") & ""
        row.Item("ITEM_DESC") = ""
        row.Item("BAD_ITEM") = ""
        Dim rowDPTFCSTD As DataRow = dst.Tables("DPTFCSTD").Rows.Find(ITEM_CODE)
        If rowDPTFCSTD IsNot Nothing Then
            row.Item("ITEM_DESC") = rowDPTFCSTD.Item("ITEM_DESC")
            row.Item("ITEM_CATGY_CODE_CURR") = rowDPTFCSTD.Item("ITEM_CATGY_CODE_CURR")
        Else
            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            If rowICTITEM1 IsNot Nothing Then
                row.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
                row.Item("ITEM_CATGY_CODE_CURR") = rowICTITEM1.Item("ITEM_CATGY_CODE")
            Else
                row.Item("BAD_ITEM") = "1"
            End If
        End If
    End Sub
#End Region

#Region "grdDPTFCSTD"

    Private Sub grdDPTFCSTD_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdDPTFCSTD.AfterExitEditMode

    End Sub

    Private Sub grdDPTFCSTD_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdDPTFCSTD.AfterRowActivate
        If grdDPTFCSTD.ActiveRow.IsDataRow Then
            Load_DPTFCSTI()
        End If
    End Sub

    Private Sub grdDPTFCSTD_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdDPTFCSTD.BeforeRowUpdate
        With grdDPTFCSTD

        End With
    End Sub


    Private Sub grdDPTFCSTD_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdDPTFCSTD.ClickCellButton
        Dim ITEM_CODE As String = e.Cell.Value & ""
        Dim FOLDERNAME As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
        If ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_SERVER = "" Then FOLDERNAME = "C:\Documents and Settings\wjz\Desktop\JHI\Images\"
        Dim FILENAME As String = e.Cell.Row.Cells("ITEM_PICTURE_FILENAME").Text ' "AB301.jpg"
        PictureBox1.Image = ASCMAIN1.Get_Image(FOLDERNAME, FILENAME)

        With UltraExplorerBar1
            .Groups("Projection Status").Expanded = False
            .Groups("Prorate Sales Using").Expanded = False
            '.Groups("Display Options").Expanded = False
            .Groups("Picture").Expanded = True
            .Groups("Picture").Text = "Item " & ITEM_CODE
        End With

        'Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
        'Dim FILENAME As String = "C:\Documents and Settings\wjz\Desktop\JHI\Images\" & ITEM_CODE & ".JPG"
        If My.Computer.FileSystem.FileExists(FOLDERNAME & "\" & FILENAME) Then
            Me.Cursor = Cursors.WaitCursor
            Call ASCMAIN1.Progress("Now Loading Image Viewer")
            System.Diagnostics.Process.Start(FOLDERNAME & "\" & FILENAME)
            Me.Cursor = Cursors.Default
            Call ASCMAIN1.Progress("")
        End If

    End Sub

    Private Sub grdDPTFCSTD_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdDPTFCSTD.InitializeRow
        e.Row.Cells("FCX").Value = "Proj"
        e.Row.Cells("PROJ_PX").Value = "Proj"
        e.Row.Cells("DTY_INX").Value = "TY-In"
        e.Row.Cells("DLY_INX").Value = "LY-In"
        e.Row.Cells("DTY_STX").Value = "TY-ST"
        e.Row.Cells("DLY_STX").Value = "LY-ST"

        Dim F As Decimal = 1
        If optUSC.Value = "S" Then
            F = Val(e.Row.Cells("ITEM_PRICE").Value & "")
        ElseIf optUSC.Value = "C" Then
            F = Val(e.Row.Cells("ITEM_COST_STD").Value & "")
        End If

        Dim negative_projections As Boolean = False
        For p As Integer = 0 To 6
            If Val(e.Row.Cells("PROJ_P" & CStr(p)).Value & "") < 0 Then
                negative_projections = True
                e.Row.Cells("PROJ_P" & CStr(p)).Appearance.ForeColor = Color.Red
            Else
                e.Row.Cells("PROJ_P" & CStr(p)).Appearance.ForeColor = Color.Empty
            End If
            'If e.Row.Cells("PROJ_P" & CStr(p)).Column.CellActivation = UltraWinGrid.Activation.AllowEdit Then
            '    e.Row.Cells("PROJ_P" & CStr(p)).Appearance.BackColor = Color.AliceBlue
            'End If
            e.Row.Cells("DTY_IN" & CStr(p)).Value = Val(e.Row.Cells("TY_IN" & CStr(p)).Value & "") * F
            e.Row.Cells("DLY_IN" & CStr(p)).Value = Val(e.Row.Cells("LY_IN" & CStr(p)).Value & "") * F
            e.Row.Cells("DTY_ST" & CStr(p)).Value = Val(e.Row.Cells("TY_ST" & CStr(p)).Value & "") * F
            e.Row.Cells("DLY_ST" & CStr(p)).Value = Val(e.Row.Cells("LY_ST" & CStr(p)).Value & "") * F
        Next

        If negative_projections Then
            e.Row.Cells("ITEM_CODE").Appearance.ForeColor = Color.Red
        Else
            e.Row.Cells("ITEM_CODE").Appearance.ForeColor = Color.Empty
        End If
    End Sub
#End Region

#Region "grdICTITEMI"

    Private Sub grdICTITEMI_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTITEMI.AfterCellUpdate

        If e.Cell.Column.Key = "ITEM_CODE" Then
            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", e.Cell.Text)
            If rowICTITEM1 IsNot Nothing Then
                e.Cell.Row.Cells("ITEM_DESC").Value = rowICTITEM1.Item("ITEM_DESC") & ""
                e.Cell.Row.Cells("BAD_ITEM").Value = ""
            Else
                e.Cell.Row.Cells("ITEM_DESC").Value = ""
                e.Cell.Row.Cells("BAD_ITEM").Value = "1"
            End If
        End If

    End Sub

    Private Sub grdICTITEMI_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTITEMI.AfterRowActivate

        If grdICTITEMI.ActiveRow Is Nothing Then Exit Sub

        With grdICTITEMI.DisplayLayout.Bands(0)
        End With
    End Sub

    Private Sub grdICTITEMI_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTITEMI.AfterRowUpdate
    End Sub

    Private Sub grdICTITEMI_BeforeCellActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles grdICTITEMI.BeforeCellActivate
    End Sub

    Private Sub grdICTITEMI_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdICTITEMI.BeforeCellUpdate
    End Sub

    Private Sub grdICTITEMI_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdICTITEMI.BeforeExitEditMode
        Call grdFieldFormat(grdICTITEMI)
    End Sub

    Private Sub grdICTITEMI_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdICTITEMI.BeforeRowsDeleted

    End Sub

    Private Sub grdICTITEMI_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTITEMI.BeforeRowUpdate
        With grdICTITEMI
            If LookUp("ICTITEM1", e.Row.Cells("ITEM_CODE").Text) Is Nothing Then
                MsgBox("Invalid Value entered for Item Code (" & e.Row.Cells("ITEM_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
                Exit Sub
            End If
        End With
    End Sub

    Private Sub grdICTITEMI_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTITEMI.ClickCellButton
        Dim sql_where As String = ""
        Call grdClickCellButton(grdICTITEMI, sql_where, False)
    End Sub

    Private Sub grdICTITEMI_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdICTITEMI.Error
        grdICTITEMI.ActiveRow.CancelUpdate()
    End Sub

    Private Sub grdICTITEMI_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTITEMI.InitializeRow
        If e.Row.Cells("ITEM_CATGY_CODE").Value & "" _
        <> e.Row.Cells("ITEM_CATGY_CODE_CURR").Value & "" And e.Row.Cells("ITEM_CATGY_CODE_CURR").Value & "" <> "" Then
            e.Row.Cells("ITEM_CATGY_CODE").Appearance.BackColor = Color.Red
        Else
            e.Row.Cells("ITEM_CATGY_CODE").Appearance.BackColor = Color.Empty
        End If
    End Sub
#End Region


    Sub Load_DPTFCSTI()
        If chkVersion.Checked Then
            If grdDPTFCSTD.ActiveRow Is Nothing Then
                grdDPTFCSTI.Visible = False
            Else
                grdDPTFCSTI.Visible = True
                Dim ITEM_CODE As String = grdDPTFCSTD.ActiveRow.Cells("ITEM_CODE").Text
                Fill_Records("DPTFCSTI", New String() {HFs("BRAND_CODE"), HFs("OPS_YYYY"), HFs("SEASON"), ITEM_CODE})
                Sort_grdColumns(grdDPTFCSTI, "VERSION")
                grdDPTFCSTI.Text = "Projections from Previous Versions for Item " & ITEM_CODE
            End If
        End If
    End Sub

    Private Sub optCOLLECTION_CODE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCOLLECTION_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_grdDPTFCSTD()
    End Sub

    Private Sub optBP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optITEM_CATGY_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        negatives_only = False
        Setup_grdDPTFCSTD()
    End Sub

    Sub Setup_grdDPTFCSTD()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")
        Application.DoEvents()

        cbeCOLLECTION_CODE.Enabled = (optCOLLECTION_CODE.Value = "I")
        'grdDPTFCSTD.DisplayLayout.GroupByBox.Hidden = (optCOLLECTION_CODE.Value = "A")

        With grdDPTFCSTD.DisplayLayout.Bands(0)
            If optCOLLECTION_CODE.Value = "A" Then
                .Columns("COLLECTION_CODE").Hidden = False
            End If
            .SortedColumns.Clear()
            If optCOLLECTION_CODE.Value = "*" Then
                'grdDPTFCSTD.DisplayLayout.Bands(0).SortedColumns.Add("COLLECTION_CODE", False, True)
            End If
            .SortedColumns.Add("ITEM_CODE", False)

            If optITEM_CATGY_CODE.Value = "*" Then
                .Columns("ITEM_CATGY_CODE").Hidden = False
            End If
        End With

        Dim COLLS As String = ""
        Dim allow_modifications As Boolean = True

        Dim DVW As DataView = DirectCast(grdDPTFCSTD.DataSource, DataTable).DefaultView
        Dim sql As String = ""
        If optCOLLECTION_CODE.Value = "A" Then
            COLLS = "All Collections"
        Else
            sql = "and COLLECTION_CODE = '" & cbeCOLLECTION_CODE.Value & "'"
            COLLS = cbeCOLLECTION_CODE.Value
        End If

        If optITEM_CATGY_CODE.Value = "*" Or optITEM_CATGY_CODE.Value = "$" Then
            'allow_modifications = False ' DON'T KNOW WHY i SHOULDNT BE ABLE TO MODIFY PROJECTIONS WHILE LOOKING AT ALL ITEMS
            If optITEM_CATGY_CODE.Value = "$" Then
                sql &= " and (ITEM_CATGY_CODE = 'C' or ITEM_CATGY_CODE = 'N' or ITEM_CATGY_CODE = 'E')"
            End If
        Else
            sql &= " and ITEM_CATGY_CODE = '" & optITEM_CATGY_CODE.Value & "'"
        End If

        If negatives_only Then
            sql &= " and (ISNULL(PROJ_P1,0) < 0 or ISNULL(PROJ_P2,0) < 0 or ISNULL(PROJ_P3,0) < 0 or ISNULL(PROJ_P4,0) < 0 or ISNULL(PROJ_P5,0) < 0 or ISNULL(PROJ_P6,0) < 0)"
        End If

        DVW.RowFilter = Mid(sql, 5)
        grdDPTFCSTD.Text = "Projections by Item/Month, for " & optSEASON.Text & " " & Absx1.cbeFor("OPS_YYYY").Value & " - " & COLLS & ", " & optITEM_CATGY_CODE.Text
        If allow_modifications Then
            grdDPTFCSTD.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        Else
            grdDPTFCSTD.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        End If

        ReDim TOTAL_SALES(6)
        For Each row As DataRow In dst.Tables("DPTFCSTD").Select(Mid(sql, 5))
            Dim ITEM_PRICE As Decimal = Val(row.Item("ITEM_PRICE") & "")
            For P As Integer = 1 To 6
                TOTAL_SALES(P) += Val(row.Item("LY_IN" & CStr(P)) & "") * ITEM_PRICE
            Next
        Next

        For P As Integer = 1 To 6
            TOTAL_SALES(0) += TOTAL_SALES(P)
        Next

        optProRate.ValueList.ValueListItems(1).DisplayText = "LY Actuals " & Format(TOTAL_SALES(0), "$###,##0")

        Setup_tabMain()
        'With UltraExplorerBar1
        '    .Groups("Projection Status").Visible = (tabMain.SelectedTab.Key = "Projections")
        '    .Groups("Prorate Sales Using").Visible = (tabMain.SelectedTab.Key = "Projections")
        'End With

        Sort_grdColumns(grdDPTFCSTD, "ITEM_CODE")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub cbeCOLLECTION_CODE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeCOLLECTION_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_grdDPTFCSTD()
    End Sub

    Private Sub UltraTabControl1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()
        If SELECTION_NO = 0 Then Exit Sub
        With UltraExplorerBar1
            .Groups("Display Options").Visible = (tabMain.SelectedTab.Key = "Projections") And ScreenMode
            .Groups("Projection Status").Visible = (tabMain.SelectedTab.Key = "Projections") And ScreenMode
            .Groups("Prorate Sales Using").Visible = (tabMain.SelectedTab.Key = "Projections") And ScreenMode And VERSION_STATUS <> "L"
            .Groups("Load Category").Visible = (tabMain.SelectedTab.Key = "Import Item Catgy from Excel") And ScreenMode
            .Groups("Picture").Visible = (tabMain.SelectedTab.Key = "Projections") And ScreenMode
        End With
    End Sub

    Sub Load_Versions(Optional ByVal set_latest As Boolean = False)

        Dim VERSION_save As String = cmbVersion.Value & ""
        cmbVersion.Value = DBNull.Value

        dst.Tables("DPTFCSTV").Rows.Clear()

        dst.Tables("DPTFCSTV").Rows.Add(New String() {"L", "Latest"})
        dst.Tables("DPTFCSTV").Rows.Add(New String() {"$$", "Sales"})

        Dim LATEST_VERSION As String = "$$"

        ASCMAIN1.sql = "Select VERSION, VERSION_DESC from DPTPROJ1 " _
        & " where BRAND_CODE = '" & Absx1.txtFor("BRAND_CODE").Text & "'" _
        & "   and MARKET_CODE = '" & Absx1.txtFor("MARKET_CODE").Text & "'" _
        & "   and OPS_YYYY = '" & cmbOPS_YYYY.Text & "'" _
        & "   and SEASON = '" & optSEASON.Value & "' order by VERSION"

        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            Dim VERSION As String = row.Item("VERSION")
            Dim VERSION_DESC As String = row.Item("VERSION_DESC") & ""
            If VERSION <> "$$" Then
                LATEST_VERSION = VERSION
                dst.Tables("DPTFCSTV").Rows.Add(New String() {VERSION, VERSION_DESC})
            End If
        Next

        If set_latest Then
            Absx1.cbeFor("VERSION").Value = LATEST_VERSION
        End If
    End Sub

    Private Sub cmbVersion_BeforeDropDown(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cmbVersion.BeforeDropDown
        If ScreenMode Then Exit Sub
        Load_Versions()
        Load_DPTFCSTI()
    End Sub

    Sub Format_grdDPTFCSTD()

        For M As Integer = 1 To 6
            Call Create_Summary(grdDPTFCSTD, "QTY_FC" & Format(M, "0"), , , "###,##0")
        Next

        Dim G As UltraWinGrid.UltraGridGroup

        With grdDPTFCSTD.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() _
                {"ITEM_CODE", "ITEM_DESC", "ITEM_CATGY_CODE", "ITEM_CLASS_CODE", _
                 "COLLECTION_CODE", "DEPT_CODE", "VEND_CODE", "VEND_CODE_2", _
                 "ITEM_RETAIL_PRICE", "ITEM_PRICE", "ITEM_COST_STD"}
                G = .Groups.Add(COLUMN_NAME, .Columns(COLUMN_NAME).Header.Caption)
                G.Header.Appearance.BackColor = Color.CornflowerBlue
                G.Header.Appearance.BackGradientStyle = GradientStyle.None
                G.Header.Appearance.ForeColor = Color.White
                G.Header.Fixed = True
                With .Columns(COLUMN_NAME)
                    .Group = G
                    If COLUMN_NAME = "ITEM_CATGY_CODE" Then
                        .CellAppearance.BackColor = Drawing.Color.Azure
                        .CellActivation = UltraWinGrid.Activation.AllowEdit
                    Else
                        .CellAppearance.BackColor = Drawing.Color.Beige
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                    End If
                End With
                G.Hidden = .Columns(COLUMN_NAME).Hidden
            Next
            Create_Summary(grdDPTFCSTD, "ITEM_CODE", "Count")

            For Each COLUMN_NAME As String In New String() _
                {"PROJ_QOH_BEG", "QTY_OPO", "QTY_FC", "AMT_FC", "PROJ_QOH_END"}
                G = .Groups.Add(COLUMN_NAME, .Columns(COLUMN_NAME).Header.Caption)
                G.Header.Appearance.BackColor = Color.Goldenrod
                G.Header.Appearance.BackGradientStyle = GradientStyle.None
                G.Header.Appearance.ForeColor = Color.White
                G.Header.Fixed = True

                With .Columns(COLUMN_NAME)
                    .Group = G
                    .CellAppearance.BackColor = Drawing.Color.Beige
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                    .Width = 55
                End With
                Create_Summary(grdDPTFCSTD, COLUMN_NAME, , , "#,##0")
            Next

            .LevelCount = 6

            For P As Integer = -1 To 6
                G = .Groups.Add("P" & IIf(P = -1, "X", Format(P, "0")), IIf(P = -1, "Data", IIf(P = 0, "Totals", "P" & Format(P, "0"))))
                G.Header.Appearance.BackColor = IIf(P = 0, Color.LightGreen, Color.LightYellow)
                G.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If P <> -1 Then
                    G.Header.Appearance.TextHAlign = HAlign.Right
                End If

                Dim LVL As Integer = 0

                For Each COLUMN_PFX As String In New String() _
                {"PROJ_P", "FC", "TY_IN", "LY_IN", "TY_ST", "LY_ST"}
                    Dim COLUMN_NAME As String = COLUMN_PFX & IIf(P = -1, "X", Format(P, "0"))
                    Dim display_column As Boolean = False
                    If COLUMN_PFX = "PROJ_P" Or COLUMN_PFX = "FC" Then
                    Else
                        display_column = True
                    End If
                    If display_column Then
                        COLUMN_NAME = "D" & COLUMN_NAME
                    End If
                    If P = -1 Then
                        .Columns.Add(COLUMN_NAME)
                    ElseIf display_column Then
                        With .Columns.Add(COLUMN_NAME)
                            .Format = "###,##0"
                        End With
                    End If

                    With .Columns(COLUMN_NAME)
                        .Group = G
                        .Level = LVL
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                        .Format = "#,##0"
                        .Width = 55
                        .CellAppearance.BackColor = Color.LightYellow
                        If P <> -1 Then
                            .CellAppearance.TextHAlign = HAlign.Right
                            .Header.Appearance.TextHAlign = HAlign.Right
                        End If

                        If P >= 0 And (COLUMN_PFX = "FC" Or COLUMN_PFX = "PROJ_P") And VERSION_STATUS <> "L" Then
                            Create_Summary(grdDPTFCSTD, COLUMN_NAME, , , "#,##0")
                            If P > 0 Then .CellActivation = UltraWinGrid.Activation.AllowEdit
                        End If

                    End With

                    If COLUMN_PFX <> "PROJ_P" Then
                        LVL += 1
                    End If
                Next
            Next

            .ColHeadersVisible = False
        End With

        Set_Levels()
    End Sub

    Private Sub chkLY_IN_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLY_IN.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_Levels()
    End Sub

    Private Sub chkTY_IN_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTY_IN.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_Levels()
    End Sub

    Private Sub chkTY_ST_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTY_ST.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_Levels()
    End Sub

    Private Sub chkLY_ST_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLY_ST.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_Levels()
    End Sub

    Sub Set_Levels()

        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()
        ASCMAIN1.Progress("Now Setting Up Grid")
        Application.DoEvents()

        Dim LVL As Int32 = 0

        Dim CHK As New Dictionary(Of String, Boolean)
        CHK.Add("TY_IN", chkTY_IN.Checked)
        CHK.Add("LY_IN", chkLY_IN.Checked)
        CHK.Add("TY_ST", chkTY_ST.Checked)
        CHK.Add("LY_ST", chkLY_ST.Checked)

        Dim LVLS As Int32 = 1 + IIf(CHK("TY_IN"), 1, 0) + IIf(CHK("LY_IN"), 1, 0) + IIf(CHK("TY_ST"), 1, 0) + IIf(CHK("LY_ST"), 1, 0)

        With grdDPTFCSTD.DisplayLayout.Bands(0)

            For P As Integer = -1 To 6
                .Columns("PROJ_P" & IIf(P = -1, "X", Format(P, "0"))).Hidden = (optUSC.Value <> "U")
                .Columns("FC" & IIf(P = -1, "X", Format(P, "0"))).Hidden = (optUSC.Value = "U")
                If P > 0 And optUSC.Value <> "U" Then
                    If optUSC.Value = "S" Then
                        dst.Tables("DPTFCSTD").Columns("FC" & Format(P, "0")).Expression = "ISNULL(PROJ_P" & Format(P, "0") & ",0) * ISNULL(ITEM_PRICE,0)"
                    ElseIf optUSC.Value = "C" Then
                        dst.Tables("DPTFCSTD").Columns("FC" & Format(P, "0")).Expression = "ISNULL(PROJ_P" & Format(P, "0") & ",0) * ISNULL(ITEM_COST_STD,0)"
                    Else

                    End If
                End If
            Next

            If .LevelCount < LVLS Then
                .LevelCount = LVLS
            End If

            For Each COLUMN_PFX As String In New String() _
            {"TY_IN", "LY_IN", "TY_ST", "LY_ST"}
                If CHK(COLUMN_PFX) Then
                    LVL += 1
                End If
                For P As Integer = -1 To 6
                    Dim COLUMN_NAME As String = COLUMN_PFX & IIf(P = -1, "X", Format(P, "0"))

                    'Dim display_column As Boolean = False
                    'If COLUMN_PFX = "TY_IN" Or COLUMN_PFX = "LY_IN" Or COLUMN_PFX = "TY_ST" Or COLUMN_PFX = "LY_ST" Then
                    '    display_column = True
                    '    COLUMN_NAME = "D" & COLUMN_NAME
                    'End If
                    COLUMN_NAME = "D" & COLUMN_NAME

                    If CHK(COLUMN_PFX) Then
                        'If P = -1 Then Stop
                        .Columns(COLUMN_NAME).Level = LVL
                        .Columns(COLUMN_NAME).Hidden = False
                    Else
                        .Columns(COLUMN_NAME).Level = 0
                        .Columns(COLUMN_NAME).Hidden = True
                    End If
                Next
            Next

            .LevelCount = LVLS

            .Groups("PX").Hidden = (LVLS = 1)
        End With

        grdDPTFCSTD.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub optUSC_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optUSC.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_Levels()
    End Sub

    Private Sub optProRate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optProRate.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
    End Sub

    Private Sub btnProrate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProrate.Click

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Prorating")

        If TOTAL_SALES(0) = 0 Then
            MsgBox("Cannot ProRate when Sales Basis is 0", MsgBoxStyle.OkOnly, "Cannot Perform Requesetd Action")
            Exit Sub
        End If


        With grdDPTFCSTD.DisplayLayout.Bands(0)
            For P As Int32 = 1 To 6
                If .Columns("PROJ_P" & CStr(P)).CellActivation = UltraWinGrid.Activation.NoEdit Then
                    MsgBox("Some months do not permit editing", MsgBoxStyle.OkOnly, "Proration option is not permitted")
                    Exit Sub
                End If
            Next
        End With

        Dim FACTOR As Decimal = Val(numPRORATE.Value & "") / TOTAL_SALES(0)
        Dim SORTED As String = "1"
        For P As Int32 = 2 To 6
            For S As Int32 = 1 To SORTED.Length
                If TOTAL_SALES(P) < TOTAL_SALES(Val(Mid(SORTED, S, 1))) Then
                    SORTED = Mid(SORTED, 1, S - 1) & CStr(P) & Mid(SORTED, S)
                    Exit For
                End If
            Next
            If InStr(SORTED, CStr(P)) = 0 Then SORTED &= CStr(P)
        Next

        For Each row As DataRow In dst.Tables("DPTFCSTD").Select(dst.Tables("DPTFCSTD").DefaultView.RowFilter)
            Dim LY(6) As Int32
            For P As Int32 = 0 To 6
                LY(P) = Val(row.Item("LY_IN" & CStr(P)) & "")
            Next

            Dim ITEM_PRICE As Decimal = Val(row.Item("ITEM_PRICE") & "")
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")

            'If ITEM_CODE = "NB94634BLX18" Then Stop

            Dim UNITS(6) As Int32
            UNITS(0) = System.Math.Round(LY(0) * FACTOR, 0)
            For P As Int32 = 1 To 6
                UNITS(P) = System.Math.Round(FACTOR * LY(P), 0)
                UNITS(0) -= UNITS(P)
            Next

            If UNITS(0) <> 0 Then
                Dim LI As Int32 = IIf(UNITS(0) < 0, 6, 1)
                Dim UNITS_REMAINING As Int32 = System.Math.Abs(UNITS(0))
                Do While UNITS_REMAINING <> 0
                    Dim P As Int32 = Val(Mid(SORTED, LI, 1))
                    UNITS(P) += System.Math.Sign(UNITS(0))
                    LI += System.Math.Sign(UNITS(0))
                    UNITS_REMAINING -= 1
                    If LI = 0 Then LI = 6
                    If LI = 7 Then LI = 1
                Loop
            End If

            For P As Int32 = 1 To 6
                'If UNITS(P) < 0 Then UNITS(P) = 0
                row.Item("PROJ_P" & CStr(P)) = UNITS(P)
            Next
        Next

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("")

    End Sub

    Private Sub chkVersion_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkVersion.CheckedChanged
        splDPTFCSTD.Panel2Collapsed = Not chkVersion.Checked
    End Sub

    Overrides Function Excel_Export(ByVal grd As UltraWinGrid.UltraGrid) As GemBox.Spreadsheet.ExcelFile
        If grd.Name = "grdDPTFCSTD" Then

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Preparing Export")

            grdDPTFCSTY.Text = grdDPTFCSTY.Text
            If dst.Tables.Contains("DPTFCSTY") Then
                grdDPTFCSTY.DataSource = Nothing
                dst.Tables.Remove("DPTFCSTY")
            End If
            Dim CL() As Int32 = Nothing
            Dim CW() As Int32 = Nothing
            Dim CG() As Int32 = Nothing
            Dim CF() As String = Nothing
            Dim CA() As Infragistics.Win.Appearance = Nothing

            Dim RL() As String = Nothing
            ReDim RL(grd.DisplayLayout.Bands(0).LevelCount - 1)
            With dst.Tables.Add("DPTFCSTY")
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
                                dt(CLmax - 1) = C.DataType
                            End If

                            If C.DataType.ToString <> dt(CL(C.Level) - 1).ToString Then
                                If dt(CL(C.Level) - 1).ToString <> "System.String" Then
                                    If C.DataType.ToString = "System.String" Then
                                        dt(CL(C.Level) - 1) = C.DataType
                                    ElseIf C.DataType.ToString = "System.Decimal" And dt(CL(C.Level) - 1).ToString Like "System.Int*" Then
                                        dt(CL(C.Level) - 1) = C.DataType
                                    End If
                                End If
                            End If

                            Dim COLUMN_NAME As String = "G" & CStr(G.Index) & "_" & CStr(CL(C.Level))
                            RL(C.Level) &= "," & C.Key & ":" & COLUMN_NAME
                        End If
                    Next
                    For i As Int32 = 1 To CLmax
                        Dim COLUMN_NAME As String = "G" & CStr(G.Index) & "_" & CStr(i)
                        With .Columns.Add(COLUMN_NAME, dt(i - 1))

                        End With
                    Next
                Next

                Dim ROWNUM As Int32 = -1
                grdDPTFCSTY.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
                grdDPTFCSTY.DataSource = dst.Tables("DPTFCSTY")
                grdDPTFCSTY.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes
                grdDPTFCSTY.DisplayLayout.Override.RowSizing = UltraWinGrid.RowSizing.Fixed

                With grdDPTFCSTY.DisplayLayout.Bands(0)
                    Dim LAST_G As Int32 = -1
                    For I As Int32 = 1 To .Columns.Count - 1
                        .Columns(I).Width = CW(I - 1)
                        .Columns(I).Format = CF(I - 1)
                        .Columns(I).CellAppearance = CA(I - 1)
                        If LAST_G <> CG(I - 1) Then
                            LAST_G = CG(I - 1)
                            .Columns(I).Header.Caption = grdDPTFCSTD.DisplayLayout.Bands(0).Groups(CG(I - 1)).Header.Caption
                        Else
                            .Columns(I).Header.Caption = ""
                        End If
                        .Columns(I).Header.Appearance = grdDPTFCSTD.DisplayLayout.Bands(0).Groups(CG(I - 1)).Header.Appearance
                    Next
                End With

                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    ASCMAIN1.Progress(grow.Cells("ITEM_CODE").Text)

                    Dim grow2 As UltraWinGrid.UltraGridRow
                    For R As Int32 = 0 To RL.Length - 1
                        grdDPTFCSTY.DataSource = dst.Tables("DPTFCSTY")
                        grow2 = grdDPTFCSTY.DisplayLayout.Bands(0).AddNew

                        For Each CXY As String In Split(Mid(RL(R), 2), ",")
                            Dim COLUMN_NAME_X As String = Split(CXY, ":")(0)
                            Dim COLUMN_NAME_Y As String = Split(CXY, ":")(1)
                            If grow.Cells(COLUMN_NAME_X).Column.DataType.ToString = "System.String" Then
                                grow2.Cells(COLUMN_NAME_Y).Value = grow.Cells(COLUMN_NAME_X).Value & ""
                            Else
                                grow2.Cells(COLUMN_NAME_Y).Value = Val(grow.Cells(COLUMN_NAME_X).Value & "")
                            End If
                            grow2.Cells(COLUMN_NAME_Y).Appearance = grow.Cells(COLUMN_NAME_X).Appearance
                        Next

                        ROWNUM += 1
                        grow2.Cells("ROWNUM").Value = ROWNUM

                        grow2.Update()
                    Next
                Next

                Sort_grdColumns(grdDPTFCSTY, "ROWNUM")

                grdDPTFCSTY.DisplayLayout.Bands(0).Columns("ROWNUM").Hidden = True

            End With

            tabMain.Tabs("grd").Visible = True
            MyBase.Excel_Export(grdDPTFCSTY)
            grdDPTFCSTY.DataSource = Nothing
            tabMain.Tabs("grd").Visible = False

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")

        Else
            MyBase.Excel_Export(grd)
        End If
        Return Nothing
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class