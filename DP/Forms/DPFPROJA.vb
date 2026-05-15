Public Class DPFPROJA

    Dim YP(,) As String
    Dim YP_LY(,) As String
    Dim OPS_YYYY As String
    Dim SEASON As String

    Dim DPTPROJI As String
    Dim sqlDPTPROJI As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")

        With dst
            Call Create_TDA(.Tables.Add, "ICTBRAN1", "*")

            With .Tables.Add("DPTPROJA")
                .Columns.Add("PCT", GetType(System.Int32))
                .Columns.Add("LINE_DESC")
                '.Columns.Add("G1")
                '.Columns.Add("G2")
                '.Columns.Add("G3")
                Dim SQLT As String = "ISNULL(N_X,0) + ISNULL(C_X,0) + ISNULL(E_X,0)"
                For Each ITEM_CATGY_CODE As String In New String() {"N", "C", "E", "I", "T"}
                    .Columns.Add(ITEM_CATGY_CODE & "_ITEMS", GetType(System.Int32), IIf(ITEM_CATGY_CODE = "T", Replace(SQLT, "_X", "_ITEMS"), ""))
                    .Columns.Add(ITEM_CATGY_CODE & "_PCT", GetType(System.Double), IIf(ITEM_CATGY_CODE = "T", Replace(SQLT, "_X", "_PCT"), ""))
                    .Columns.Add(ITEM_CATGY_CODE & "_QTY_PROJ", GetType(System.Int32), IIf(ITEM_CATGY_CODE = "T", Replace(SQLT, "_X", "_QTY_PROJ"), ""))
                    .Columns.Add(ITEM_CATGY_CODE & "_QTY_SHIP", GetType(System.Int32), IIf(ITEM_CATGY_CODE = "T", Replace(SQLT, "_X", "_QTY_SHIP"), ""))
                    .Columns.Add(ITEM_CATGY_CODE & "_AMT_PROJ", GetType(System.Double), IIf(ITEM_CATGY_CODE = "T", Replace(SQLT, "_X", "_AMT_PROJ"), ""))
                    .Columns.Add(ITEM_CATGY_CODE & "_AMT_SHIP", GetType(System.Double), IIf(ITEM_CATGY_CODE = "T", Replace(SQLT, "_X", "_AMT_SHIP"), ""))
                    .Columns.Add(ITEM_CATGY_CODE & "_QTY_PROJ_PCT", GetType(System.Int32), IIf(ITEM_CATGY_CODE = "T", Replace(SQLT, "_X", "_QTY_PROJ_PCT"), ""))
                    .Columns.Add(ITEM_CATGY_CODE & "_QTY_SHIP_PCT", GetType(System.Int32), IIf(ITEM_CATGY_CODE = "T", Replace(SQLT, "_X", "_QTY_SHIP_PCT"), ""))
                    .Columns.Add(ITEM_CATGY_CODE & "_AMT_PROJ_PCT", GetType(System.Double), IIf(ITEM_CATGY_CODE = "T", Replace(SQLT, "_X", "_AMT_PROJ_PCT"), ""))
                    .Columns.Add(ITEM_CATGY_CODE & "_AMT_SHIP_PCT", GetType(System.Double), IIf(ITEM_CATGY_CODE = "T", Replace(SQLT, "_X", "_AMT_SHIP_PCT"), ""))
                Next
            End With

            With .Tables.Add("DPTPROJB")
                .Columns.Add("PCT", GetType(System.Int32))
                .Columns.Add("LINE_DESC")
                .Columns.Add("G1")
                .Columns.Add("G2")
                .Columns.Add("G3")
                Dim SQLT As String = "ISNULL(N_X,0) + ISNULL(C_X,0) + ISNULL(E_X,0)"
                For Each ITEM_CATGY_CODE As String In New String() {"N", "C", "E", "I", "T"}
                    .Columns.Add(ITEM_CATGY_CODE & "_ITEMS", GetType(System.Int32), IIf(ITEM_CATGY_CODE = "T", Replace(SQLT, "_X", "_ITEMS"), ""))
                    .Columns.Add(ITEM_CATGY_CODE & "_PCT", GetType(System.Double), IIf(ITEM_CATGY_CODE = "T", Replace(SQLT, "_X", "_PCT"), ""))
                    .Columns.Add(ITEM_CATGY_CODE & "_QTY_PROJ", GetType(System.Int32), IIf(ITEM_CATGY_CODE = "T", Replace(SQLT, "_X", "_QTY_PROJ"), ""))
                    .Columns.Add(ITEM_CATGY_CODE & "_QTY_SHIP", GetType(System.Int32), IIf(ITEM_CATGY_CODE = "T", Replace(SQLT, "_X", "_QTY_SHIP"), ""))
                    .Columns.Add(ITEM_CATGY_CODE & "_AMT_PROJ", GetType(System.Double), IIf(ITEM_CATGY_CODE = "T", Replace(SQLT, "_X", "_AMT_PROJ"), ""))
                    .Columns.Add(ITEM_CATGY_CODE & "_AMT_SHIP", GetType(System.Double), IIf(ITEM_CATGY_CODE = "T", Replace(SQLT, "_X", "_AMT_SHIP"), ""))
                    .Columns.Add(ITEM_CATGY_CODE & "_QTY_PROJ_PCT", GetType(System.Int32), IIf(ITEM_CATGY_CODE = "T", Replace(SQLT, "_X", "_QTY_PROJ_PCT"), ""))
                    .Columns.Add(ITEM_CATGY_CODE & "_QTY_SHIP_PCT", GetType(System.Int32), IIf(ITEM_CATGY_CODE = "T", Replace(SQLT, "_X", "_QTY_SHIP_PCT"), ""))
                    .Columns.Add(ITEM_CATGY_CODE & "_AMT_PROJ_PCT", GetType(System.Double), IIf(ITEM_CATGY_CODE = "T", Replace(SQLT, "_X", "_AMT_PROJ_PCT"), ""))
                    .Columns.Add(ITEM_CATGY_CODE & "_AMT_SHIP_PCT", GetType(System.Double), IIf(ITEM_CATGY_CODE = "T", Replace(SQLT, "_X", "_AMT_SHIP_PCT"), ""))
                Next
            End With

            Dim DT As DataTable = .Tables("DPTPROJB").Clone
            DT.TableName = "DPTPROJB_ORIG"
            .Tables.Add(DT)

            ASCMAIN1.sql = "" _
            & "SELECT X.ITEM_CODE, ICTITEM1.ITEM_DESC" _
            & ", ICTRETLA.ITEM_PRICE, ICTRETLA.ITEM_CATGY_CODE" _
            & ", ICTITEM1.COLLECTION_CODE, ICTITEM1.DEPT_CODE, ICTITEM1.ITEM_ABC_CODE" _
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
            & ", ICTITEM1.COLLECTION_CODE, ICTITEM1.DEPT_CODE, ICTITEM1.ITEM_ABC_CODE" _
            & " HAVING SUM (QTY_PROJ) <> 0 OR SUM (QTY_SHIP) <> 0"
            sqlDPTPROJI = ASCMAIN1.sql
            DPTPROJI = ASCMAIN1.Temp_Table(Replace(Replace(Replace(Replace(Replace(Replace(sqlDPTPROJI, ":PARM1", "''"), ":PARM2", "''"), ":PARM3", "''"), ":PARM4", "''"), ":PARM5", "''"), ":PARM6", "''"))
            'ASCDATA1.ExecuteSQL("Alter Table " & DPTPROJI & " Add PCT NUMBER (10,4)")

            ASCMAIN1.sql = "Select * from " & DPTPROJI
            Create_TDA(.Tables.Add, "DPTPROJI", "**", 0, False, "", 1)
            .Tables("DPTPROJI").Columns.Add("VAR_PCT", GetType(System.Double), "IIF(ISNULL(QTY_PROJ,0)=0,100,100 * ISNULL(QTY_SHIP,0)/ISNULL(QTY_PROJ,0)-100)")
            .Tables("DPTPROJI").Columns.Add("ABS_VAR_PCT", GetType(System.Double), "IIF(VAR_PCT<0,-1*VAR_PCT,VAR_PCT)")
            .Tables("DPTPROJI").Columns.Add("PCT", GetType(System.Int32))
            .Tables("DPTPROJI").Columns.Add("PCTB", GetType(System.Int32))
        End With

        For PCT As Integer = 5 To 100 Step 5
            Dim LINE_DESC As String = Format((100 - PCT) / 100, "##0%") & " to " & Format((100 + PCT) / 100, "##0%")
            dst.Tables("DPTPROJA").Rows.Add(New Object() {PCT, LINE_DESC})
        Next

        For PCT As Integer = 10 To 100 Step 10
            Dim LINE_DESC As String = Format((1 + 100 - PCT) / 100, "##0%") & " to " & Format((100 - (PCT - 10)) / 100, "##0%")
            If PCT = 100 Then LINE_DESC = "10% and Under"
            dst.Tables("DPTPROJB_ORIG").Rows.Add(New Object() {PCT, LINE_DESC})
        Next

        grdDPTPROJA.DataSource = dst.Tables("DPTPROJA")
        grdDPTPROJB.DataSource = dst.Tables("DPTPROJB")
        grdDPTPROJI.DataSource = dst.Tables("DPTPROJI")

        Sort_grdColumns(grdDPTPROJA, "PCT", True)
        Sort_grdColumns(grdDPTPROJB, "PCT", True)

        Dim G As UltraWinGrid.UltraGridGroup
        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdDPTPROJA, grdDPTPROJB}
            With grd.DisplayLayout.Bands(0)
                .Columns("PCT").Hidden = True

                If grd.Name = "grdDPTPROJB" Then
                    G = .Groups.Add("GROUPS", "")
                    G.Header.Appearance.TextHAlign = HAlign.Center
                    .Columns("G1").Group = G
                    .Columns("G2").Group = G
                    .Columns("G3").Group = G
                    .Columns("G1").CellAppearance.BackColor = Color.PaleTurquoise
                    .Columns("G2").CellAppearance.BackColor = Color.PaleTurquoise
                    .Columns("G3").CellAppearance.BackColor = Color.PaleTurquoise
                End If


                G = .Groups.Add("LINE_DESC", "")
                G.Header.Appearance.TextHAlign = HAlign.Center
                .Columns("LINE_DESC").Group = G
                .Columns("LINE_DESC").Header.Caption = "Accuracy"
                .Columns("LINE_DESC").CellAppearance.BackColor = Color.Beige
                .Columns("LINE_DESC").Header.Appearance.TextHAlign = HAlign.Center
                .Columns("LINE_DESC").CellAppearance.TextHAlign = HAlign.Center

                For Each ITEM_CATGY_DESC In New String() {"New", "Core", "Existing", "Inactive", "Total NCE"}
                    Dim ITEM_CATGY_CODE = Mid(ITEM_CATGY_DESC, 1, 1)

                    G = .Groups.Add(ITEM_CATGY_CODE, ITEM_CATGY_DESC)
                    G.Header.Appearance.TextHAlign = HAlign.Center
                    G.Header.Appearance.BackColor = IIf(ITEM_CATGY_CODE = "N", Color.Yellow, (IIf(ITEM_CATGY_CODE = "E", Color.LightGreen, IIf(ITEM_CATGY_CODE = "C", Color.HotPink, IIf(ITEM_CATGY_CODE = "I", Color.LightGray, Color.LightBlue)))))
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
                            Create_Summary(grd, ITEM_CATGY_CODE & DC, "Max", , "###,##0")
                        Next
                    Next
                    Create_Summary(grd, ITEM_CATGY_CODE & "_ITEMS", "Max", , "###,##0")
                    'Create_Summary(grd, ITEM_CATGY_CODE & "_PCT", , , "##.0")            
                Next
                .Groups("LINE_DESC").Header.Fixed = True
            End With
        Next

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
            .Columns("COLLECTION_CODE").Header.Caption = "Collection"
            .Columns("COLLECTION_CODE").Width = 90
            .Columns("DEPT_CODE").Header.Caption = "Dept"
            .Columns("DEPT_CODE").Width = 60
            .Columns("ITEM_ABC_CODE").Header.Caption = "ABC"
            .Columns("ITEM_ABC_CODE").Width = 40
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

        Create_Summary(grdDPTPROJI, "ITEM_CODE", "Count")
        Create_Summary(grdDPTPROJI, "QTY_PROJ")
        Create_Summary(grdDPTPROJI, "AMT_PROJ")
        Create_Summary(grdDPTPROJI, "QTY_SHIP")
        Create_Summary(grdDPTPROJI, "AMT_SHIP")

        With grdDPTPROJI.DisplayLayout.Bands("DPTPROJI")
            .Columns("ITEM_CODE").Header.Fixed = True
            .Columns("ITEM_DESC").Header.Fixed = True
        End With

        Dim YEARs As New List(Of String)
        For Y As Integer = Val(Now.Year) - 1 To Val(Now.Year + 1)
            YEARs.Add(Format(Y, "0000"))
        Next
        Absx1.cbeFor("OPS_YYYY").DataSource = YEARs

        ASCMAIN1.Add_Value_List(grdDPTPROJI, "ITEM_CATGY_CODE")
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
                Setup_tabMain()
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabMain.Visible = ScreenMode
        SplitContainer1.Visible = Not ScreenMode

        Setup_tabMain()
        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        dst.Tables("DPTPROJI").Rows.Clear()
        EnforceConstraints(True)

        'Absx1.txtFor("BRAND_CODE").Text = ""
        'Absx1.txtFor("MARKET_CODE").Text = ""
        If Absx1.cbeFor("OPS_YYYY").Value & "" = "" Then
            Absx1.cbeFor("OPS_YYYY").Value = Now.Year
        End If

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

            With grdDPTPROJI.DisplayLayout.Bands(0)
                .Columns("F" & Format(P, "00")).Header.Caption = "Prj " & Mid(YP(P, 1), 1, 3)
                .Columns("A" & Format(P, "00")).Header.Caption = "Act " & Mid(YP(P, 1), 1, 3)
            End With
        Next

        EnforceConstraints(False)

        ASCDATA1.ExecuteSQL("Truncate Table " & DPTPROJI)

        Absc1.Get_SQL("*")
        Dim sql_where As String = Absc1.sql_WHERE & " "
        ASCDATA1.ExecuteSQL("Insert into " & DPTPROJI & " " & Replace(sqlDPTPROJI, "GROUP BY X.ITEM_CODE", sql_where & "GROUP BY X.ITEM_CODE"), "VVVVVV", New String() {YP(1, 0), YP(2, 0), YP(3, 0), YP(4, 0), YP(5, 0), YP(6, 0)})
        Dim FILTER As String = ""
        For Each row As DataRow In Absc1.tblASTDSQLA.Select("CODE_VALUES IS NOT NULL")
            FILTER &= ";" & row.Item("COLUMN_CAPTION") & ":" & row.Item("CODE_VALUES")
        Next
        FILTER = Mid(FILTER, 2)

        Fill_Records("DPTPROJI")


        dst.Tables("DPTPROJB").Rows.Clear()
        If Absc1.SEQs = 0 Then
            dst.Tables("DPTPROJB").Merge(dst.Tables("DPTPROJB_ORIG"))
            grdDPTPROJB.DisplayLayout.Bands(0).Groups("GROUPS").Hidden = True
        Else
            grdDPTPROJB.DisplayLayout.Bands(0).Groups("GROUPS").Hidden = True
            For I As Integer = 1 To 3
                With grdDPTPROJB.DisplayLayout.Bands(0).Columns("G" & CStr(I))
                    If I > Absc1.SEQs Then
                        .Hidden = True
                    Else
                        .Hidden = False
                        .Header.Caption = Absc1.COLUMN_CAPTIONs(I - 1)
                    End If
                End With
            Next

            'Dim c3 As String = Join(Absc1.COLUMN_NAMEs.ToArray, ",")
            'Dim c3() As String = Absc1.COLUMN_NAMEs.ToArray
            Dim c3() As String = Split(Join(Absc1.COLUMN_NAMEs.ToArray, ","), ",")
            Dim sqlw As String = "QTY_PROJ >= " & CStr(Val(numMinProjQty.Value & ""))
            Dim dt As DataTable = ASCDATA1.SelectDistinct(dst.Tables("DPTPROJI").Select(sqlw), c3)
            For Each row As DataRow In dt.Rows
                For Each rowDPTPROJB_ORIG As DataRow In dst.Tables("DPTPROJB_ORIG").Rows
                    Dim rowDPTPROJB As DataRow = dst.Tables("DPTPROJB").NewRow
                    rowDPTPROJB.Item("PCT") = rowDPTPROJB_ORIG.Item("PCT")
                    rowDPTPROJB.Item("LINE_DESC") = rowDPTPROJB_ORIG.Item("LINE_DESC")
                    For i As Integer = 1 To Absc1.SEQs
                        rowDPTPROJB.Item("G" & CStr(i)) = row.Item(i - 1)
                    Next
                    dst.Tables("DPTPROJB").Rows.Add(rowDPTPROJB)
                Next
            Next
        End If

        For Each AB As String In New String() {"A", "B"}
            For Each row As DataRow In dst.Tables("DPTPROJ" & AB).Rows
                Dim PCT As Int32 = Val(row.Item("PCT") & "")
                For Each ITEM_CATGY_CODE As String In New String() {"N", "C", "E", "I"}
                    Dim sqlw As String = "ITEM_CATGY_CODE = '" & ITEM_CATGY_CODE & "'"
                    If AB = "B" Then
                        If PCT = 100 Then
                        Else
                            sqlw &= " and ABS_VAR_PCT < " & CStr(PCT)
                        End If
                    Else
                        If PCT = 100 Then
                        Else
                            sqlw &= " and VAR_PCT >= " & CStr(-1 * PCT) & " AND VAR_PCT <= " & CStr(PCT)
                        End If
                    End If

                    sqlw &= " and QTY_PROJ >= " & CStr(Val(numMinProjQty.Value & ""))

                    Dim sqlg As String = ""
                    If AB = "B" And Absc1.SEQs > 0 Then
                        For i As Integer = 1 To Absc1.SEQs
                            sqlg &= " and " & Absc1.COLUMN_NAMEs(i - 1) & " = '" & row.Item("G" & CStr(i)) & "'"
                        Next
                    End If
                    sqlw &= sqlg

                    Dim sqlw2 As String = " AND ISNULL(PCT,0) = 0"
                    If AB = "B" Then sqlw2 = " AND ISNULL(PCTB,0) = 0"
                    For Each rowDPTPROJI As DataRow In dst.Tables("DPTPROJI").Select(sqlw & sqlw2)
                        If AB = "B" Then
                            rowDPTPROJI.Item("PCTB") = PCT
                        Else
                            rowDPTPROJI.Item("PCT") = PCT
                        End If
                    Next
                    Dim ITEMS = dst.Tables("DPTPROJI").Compute("COUNT(ITEM_CODE)", sqlw)
                    Dim QTY_PROJ As Int32 = Val(dst.Tables("DPTPROJI").Compute("SUM(QTY_PROJ)", sqlw) & "")
                    Dim QTY_SHIP As Int32 = Val(dst.Tables("DPTPROJI").Compute("SUM(QTY_SHIP)", sqlw) & "")
                    Dim AMT_PROJ As Decimal = Val(dst.Tables("DPTPROJI").Compute("SUM(AMT_PROJ)", sqlw) & "")
                    Dim AMT_SHIP As Decimal = Val(dst.Tables("DPTPROJI").Compute("SUM(AMT_SHIP)", sqlw) & "")
                    row.Item(ITEM_CATGY_CODE & "_ITEMS") = ITEMS
                    row.Item(ITEM_CATGY_CODE & "_QTY_PROJ") = QTY_PROJ
                    row.Item(ITEM_CATGY_CODE & "_QTY_SHIP") = QTY_SHIP
                    row.Item(ITEM_CATGY_CODE & "_AMT_PROJ") = AMT_PROJ
                    row.Item(ITEM_CATGY_CODE & "_AMT_SHIP") = AMT_SHIP
                Next
            Next


            If Absc1.SEQs = 0 Or AB = "A" Then
                For Each ITEM_CATGY_CODE As String In New String() {"N", "C", "E", "I"}
                    Dim TOTAL As Int32 = Val(dst.Tables("DPTPROJ" & AB).Compute("MAX(" & ITEM_CATGY_CODE & "_ITEMS)", "") & "")
                    Dim EXPRESSION As String = ""
                    If TOTAL = 0 Then
                    Else
                        EXPRESSION = "100 * ISNULL(" & ITEM_CATGY_CODE & "_ITEMS,0) / " & CStr(TOTAL)
                    End If
                    dst.Tables("DPTPROJ" & AB).Columns(ITEM_CATGY_CODE & "_PCT").Expression = EXPRESSION

                    For Each QA As String In New String() {"QTY", "AMT"}
                        For Each PS As String In New String() {"PROJ", "SHIP"}
                            Dim DC As String = "_" & QA & "_" & PS
                            TOTAL = Val(dst.Tables("DPTPROJ" & AB).Compute("MAX(" & ITEM_CATGY_CODE & DC & ")", "") & "")
                            EXPRESSION = ""
                            If TOTAL = 0 Then
                            Else
                                EXPRESSION = "100 * ISNULL(" & ITEM_CATGY_CODE & DC & ",0) / " & CStr(TOTAL)
                            End If
                            dst.Tables("DPTPROJ" & AB).Columns(ITEM_CATGY_CODE & DC & "_PCT").Expression = EXPRESSION
                        Next
                    Next
                Next

            Else

                For Each ITEM_CATGY_CODE As String In New String() {"N", "C", "E", "I"}
                    dst.Tables("DPTPROJ" & AB).Columns(ITEM_CATGY_CODE & "_PCT").Expression = ""
                    dst.Tables("DPTPROJ" & AB).Columns(ITEM_CATGY_CODE & "_PCT").ReadOnly = False
                    For Each QA As String In New String() {"QTY", "AMT"}
                        For Each PS As String In New String() {"PROJ", "SHIP"}
                            Dim DC As String = "_" & QA & "_" & PS
                            dst.Tables("DPTPROJ" & AB).Columns(ITEM_CATGY_CODE & DC & "_PCT").Expression = ""
                            dst.Tables("DPTPROJ" & AB).Columns(ITEM_CATGY_CODE & DC & "_PCT").ReadOnly = False
                        Next
                    Next
                Next

                Dim c3() As String = Split(Join(Absc1.COLUMN_NAMEs.ToArray, ","), ",")
                Dim sqlw As String = "QTY_PROJ >= " & CStr(Val(numMinProjQty.Value & ""))
                Dim dt As DataTable = ASCDATA1.SelectDistinct(dst.Tables("DPTPROJI").Select(sqlw), c3)

                For Each row As DataRow In dt.Rows
                    Dim sqlg As String = ""
                    For i As Integer = 1 To Absc1.SEQs
                        sqlg &= " and G" & CStr(i) & " = '" & row.Item(Absc1.COLUMN_NAMEs(i - 1)) & "'"
                    Next
                    sqlg = Mid(sqlg, 5)

                    For Each ITEM_CATGY_CODE As String In New String() {"N", "C", "E", "I"}

                        Dim TOTAL As Int32 = Val(dst.Tables("DPTPROJ" & AB).Compute("MAX(" & ITEM_CATGY_CODE & "_ITEMS)", sqlg) & "")
                        If TOTAL <> 0 Then
                            For Each rowDPTPROJB As DataRow In dst.Tables("DPTPROJB").Select(sqlg)
                                rowDPTPROJB.Item(ITEM_CATGY_CODE & "_PCT") = 100 * Val(rowDPTPROJB.Item(ITEM_CATGY_CODE & "_ITEMS") & "") / TOTAL
                            Next
                        End If

                        For Each QA As String In New String() {"QTY", "AMT"}
                            For Each PS As String In New String() {"PROJ", "SHIP"}
                                Dim DC As String = "_" & QA & "_" & PS
                                TOTAL = Val(dst.Tables("DPTPROJ" & AB).Compute("MAX(" & ITEM_CATGY_CODE & DC & ")", sqlg) & "")
                                If TOTAL <> 0 Then
                                    For Each rowDPTPROJB As DataRow In dst.Tables("DPTPROJB").Select(sqlg)
                                        rowDPTPROJB.Item(ITEM_CATGY_CODE & DC & "_PCT") = 100 * Val(rowDPTPROJB.Item(ITEM_CATGY_CODE & DC) & "") / TOTAL
                                    Next
                                End If
                            Next
                        Next
                    Next
                Next
            End If
        Next

        EnforceConstraints(True)


        grdDPTPROJB.Text = "Projection Accuracy in 10% Increments - Using Most Recently Locked Version of Projections"
        Sort_grdColumns(grdDPTPROJB, "PCT")
        If Absc1.SEQs <> 0 Then
            Dim BY As String = ""
            For I As Integer = 1 To Absc1.SEQs
                grdDPTPROJB.DisplayLayout.Bands(0).SortedColumns.Add("G" & CStr(I), False, True)
                BY &= "," & grdDPTPROJB.DisplayLayout.Bands(0).Columns("G" & CStr(I)).Header.Caption
            Next
            grdDPTPROJB.Text &= ", By " & Mid(BY, 2)
            grdDPTPROJB.Rows.ExpandAll(True)
        End If
        If FILTER <> "" Then
            grdDPTPROJB.Text &= ", " & FILTER
        End If


        Call ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        CommitTrans("Update Complete")

    End Sub
#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdDPTPROJA, "S", "Show Extended Data")
        Call Load_Popup_Menu(grdDPTPROJI, "SS", "Show Filter", "Show GroupBox")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdDPTFCSTD"
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        Select Case grd.Name
            'Case "grdDPTFCSTD"
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

    Private Sub tabMain_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()
        If SELECTION_NO = 0 Then Exit Sub
        With UltraExplorerBar1
        End With
    End Sub
End Class