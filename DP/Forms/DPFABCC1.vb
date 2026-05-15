Public Class DPFABCC1
    Dim YP() As String
    Dim ICTITEMS As String
    Dim DPTITMFX As String
    Dim sqlDPTITMFX As String
    Dim YPF(,) As String
    Dim YPFD() As Date
    Dim OPS_YYYY As String
    Dim generated As Boolean = False

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        With dst

            ' ICTITEMS - All Items on Screen

            ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC" _
            & ", ICTITEM1.COLLECTION_CODE, ICTITEM1.ITEM_CATGY_CODE" _
            & ", ICTITEM1.ITEM_ABC_CODE ITEM_ABC_CODE_PREV" _
            & ", ICTITEM1.ITEM_ABC_CODE" _
            & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_COST_STD" _
            & ", ICTITEM1.ITEM_PRICE" _
            & " from ICTITEM1 where ROWNUM < 1"
            ICTITEMS = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add Primary Key (ITEM_CODE)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add ABC_GROUP VARCHAR2(1)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add DEMAND_UNITS NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add DEMAND_WEIGHTED NUMBER (13,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add RUN_TOTAL NUMBER (15,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add RANK NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add RUN_PCT NUMBER (15,2)")


            ' DPTITMFX

            ASCMAIN1.sql = "Select DPTITMF1.ITEM_CODE" & vbCrLf
            For M As Integer = 0 To 6
                Dim YP As String = Format(M, "000000")
                'If M > 0 Then YP = YPF(M - 1, 0)
                ASCMAIN1.sql &= ", SUM (CASE WHEN DPTITMF1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "' THEN DECODE(DPTITMF1.OPS_YYYYPP_FC,'" & YP & "',DPTITMF1.FORECAST,0) ELSE DECODE(DPTITMF1.OPS_YYYYPP,'" & YP & "',DPTITMF1.FORECAST,0) END) FC" & Format(M, "00") & vbCrLf
            Next
            ASCMAIN1.sql &= " from " & ICTITEMS & " ICTITEMS,DPTITMF1" & vbCrLf _
            & " where DPTITMF1.ITEM_CODE = ICTITEMS.ITEM_CODE" & vbCrLf _
            & "   and DPTITMF1.OPS_YYYYPP BETWEEN 'YP_MIN' and 'YP_MAX'" _
            & " group by DPTITMF1.ITEM_CODE"
            sqlDPTITMFX = ASCMAIN1.sql
            DPTITMFX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & DPTITMFX & " Add Primary Key (ITEM_CODE)")


            ' DPTABCCY

            ASCMAIN1.sql = "Select ITEM_CATGY_CODE, ABC_GROUP" _
            & ", SUM (DEMAND_WEIGHTED) DEMAND_WEIGHTED, COUNT (*) ITEMS" _
            & " from " & ICTITEMS & " ICTITEMS" _
            & " group by ITEM_CATGY_CODE, ABC_GROUP"
            Create_TDA(.Tables.Add, "DPTABCCY", "**", 0, False, "", 2)

            ' ICTITEMS

            ASCMAIN1.sql = "Select ICTITEMS.* from " & ICTITEMS & " ICTITEMS"
            Create_TDA(.Tables.Add("ICTITEMS"), ICTITEMS, "**", 0, True, "", 1)


            .Relations.Add("DPTABCCY_ICTITEMS" _
            , New DataColumn() _
            {.Tables("DPTABCCY").Columns("ITEM_CATGY_CODE") _
             , .Tables("DPTABCCY").Columns("ABC_GROUP")} _
            , New DataColumn() _
            {.Tables("ICTITEMS").Columns("ITEM_CATGY_CODE") _
             , .Tables("ICTITEMS").Columns("ABC_GROUP")})


            ASCMAIN1.sql = "Select ICTCOLL1.COLLECTION_CODE, ICTCOLL1.COLLECTION_NAME" _
            & ", X.COUNT_C, X.COUNT_N, X.COUNT_E, X.COUNT_F, X.COUNT_P, X.COUNT_I, X.COUNT_X" _
            & " from ICTCOLL1," _
            & "(" _
            & " Select ICTITEM1.COLLECTION_CODE" _
            & ", Sum (DECODE(NVL(DPTPROJ0.ITEM_CATGY_CODE,ICTITEM1.ITEM_CATGY_CODE),'C',1,0)) COUNT_C" _
            & ", Sum (DECODE(NVL(DPTPROJ0.ITEM_CATGY_CODE,ICTITEM1.ITEM_CATGY_CODE),'N',1,0)) COUNT_N" _
            & ", Sum (DECODE(NVL(DPTPROJ0.ITEM_CATGY_CODE,ICTITEM1.ITEM_CATGY_CODE),'E',1,0)) COUNT_E" _
            & ", Sum (DECODE(NVL(DPTPROJ0.ITEM_CATGY_CODE,ICTITEM1.ITEM_CATGY_CODE),'F',1,0)) COUNT_F" _
            & ", Sum (DECODE(NVL(DPTPROJ0.ITEM_CATGY_CODE,ICTITEM1.ITEM_CATGY_CODE),'P',1,0)) COUNT_P" _
            & ", Sum (DECODE(NVL(DPTPROJ0.ITEM_CATGY_CODE,ICTITEM1.ITEM_CATGY_CODE),'I',1,0)) COUNT_I" _
            & ", Sum (DECODE(NVL(DPTPROJ0.ITEM_CATGY_CODE,ICTITEM1.ITEM_CATGY_CODE),'C',0,'N',0,'E',0,'F',0,'P',0,'I',0,1)) COUNT_X" _
            & " from ICTITEM1, DPTPROJ0, " & ICTITEMS & " ICTITEMS " _
            & " where ICTITEMS.ITEM_CODE = ICTITEM1.ITEM_CODE " _
            & "   and DPTPROJ0.OPS_YYYY (+) = :PARM1" _
            & "   and DPTPROJ0.SEASON (+) = :PARM2" _
            & "   and DPTPROJ0.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" _
            & " group by ICTITEM1.COLLECTION_CODE" _
            & ") X where ICTCOLL1.BRAND_CODE = :PARM3" _
            & " and X.COLLECTION_CODE (+) = ICTCOLL1.COLLECTION_CODE"
            Create_TDA(.Tables.Add, "ICTCOLLX", "**", 0, False, "VVV", 1)
            .Tables("ICTCOLLX").Columns.Add("COUNT_T", GetType(System.Int32), "ISNULL(COUNT_C,0)+ISNULL(COUNT_N,0)+ISNULL(COUNT_E,0)+ISNULL(COUNT_I,0)+ISNULL(COUNT_X,0)")
            .Tables("ICTCOLLX").Columns.Add("SELECTED")
            .Tables("ICTCOLLX").Columns.Add("ABC_GROUP")

            .Tables("ICTCOLLX").Columns("COUNT_C").DataType = GetType(System.Int32)
            .Tables("ICTCOLLX").Columns("COUNT_N").DataType = GetType(System.Int32)
            .Tables("ICTCOLLX").Columns("COUNT_E").DataType = GetType(System.Int32)
            .Tables("ICTCOLLX").Columns("COUNT_F").DataType = GetType(System.Int32)
            .Tables("ICTCOLLX").Columns("COUNT_P").DataType = GetType(System.Int32)
            .Tables("ICTCOLLX").Columns("COUNT_I").DataType = GetType(System.Int32)
            .Tables("ICTCOLLX").Columns("COUNT_X").DataType = GetType(System.Int32)

            Create_TDA(.Tables.Add, "DPTABCP2", "*", 0)

            ASCMAIN1.sql = "Select ITEM_CATGY_CODE, ITEM_CATGY_DESC, 'R' WEIGHTING from ICTCATG1"
            Create_TDA(.Tables.Add, "DPTABCC1", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "DPTITMF1", "*")
        End With

        Fill_Records("DPTABCC1")
        Sort_grdColumns(grdDPTABCC1, "ITEM_CATGY_CODE")

        grdDPTABCC1.DataSource = dst.Tables("DPTABCC1")
        grdDPTABCP2.DataSource = dst.Tables("DPTABCP2")
        grdICTCOLLX.DataSource = dst.Tables("ICTCOLLX")

        'grdICTITEMS.DataMember = "DPTABCCY"
        'grdICTITEMS.DataSource = dst
        grdICTITEMS.DataSource = dst.Tables("DPTABCCY")

        Call Create_Summary(grdICTITEMS, "ITEM_CATGY_CODE", "Count")
        Call Create_Summary(grdICTITEMS, "DEMAND_WEIGHTED")
        Call Create_Summary(grdICTITEMS, "ITEMS")

        Call Create_Summary(grdICTITEMS, "ITEM_CODE", "Count", "DPTABCCY_ICTITEMS")
        Call Create_Summary(grdICTITEMS, "DEMAND_UNITS", , "DPTABCCY_ICTITEMS")
        Call Create_Summary(grdICTITEMS, "DEMAND_WEIGHTED", , "DPTABCCY_ICTITEMS")
        Call Create_Summary(grdICTITEMS, "RUN_TOTAL", , "DPTABCCY_ICTITEMS")

        grdICTITEMS.DisplayLayout.Bands("DPTABCCY_ICTITEMS").SummaryFooterCaption = "Totals for [SCROLLTIPFIELD]"

        Create_Summary(grdDPTABCP2, "ABC_CODE", "Count")
        Create_Summary(grdDPTABCP2, "ABC_PCT_RANGE")

        Create_Summary(grdICTCOLLX, "COLLECTION_CODE", "Count")
        Create_Summary(grdICTCOLLX, "SELECTED")
        Create_Summary(grdICTCOLLX, "COUNT_C")
        Create_Summary(grdICTCOLLX, "COUNT_N")
        Create_Summary(grdICTCOLLX, "COUNT_E")
        Create_Summary(grdICTCOLLX, "COUNT_F")
        Create_Summary(grdICTCOLLX, "COUNT_P")
        Create_Summary(grdICTCOLLX, "COUNT_I")
        Create_Summary(grdICTCOLLX, "COUNT_X")
        Create_Summary(grdICTCOLLX, "COUNT_T")

        With grdDPTABCP2.DisplayLayout.Bands("DPTABCP2")
            .Columns("ABC_CODE").Header.Fixed = True
        End With
        With grdDPTABCC1.DisplayLayout.Bands("DPTABCC1")
            .Columns("ITEM_CATGY_DESC").Header.Fixed = True
        End With

        With grdICTITEMS.DisplayLayout.Bands("DPTABCCY_ICTITEMS")
            .Columns("RANK").Header.Fixed = True
            .Columns("ITEM_CODE").Header.Fixed = True
            .Columns("ITEM_DESC").Header.Fixed = True
        End With

        For Each COLUMN_NAME As String In New String() _
        {"ITEM_CODE", "ITEM_DESC", "COLLECTION_CODE", "ITEM_CATGY_CODE"}
            With grdICTITEMS.DisplayLayout.Bands("DPTABCCY_ICTITEMS").Columns(COLUMN_NAME)
                .CellAppearance.BackColor = Drawing.Color.Beige
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End With
        Next

        With grdICTITEMS.DisplayLayout.Bands("DPTABCCY_ICTITEMS")
            .Columns("DEMAND_UNITS").CellAppearance.BackColor = Drawing.Color.LightGoldenrodYellow
            .Columns("DEMAND_WEIGHTED").CellAppearance.BackColor = Drawing.Color.LightGoldenrodYellow
            .Columns("RUN_TOTAL").CellAppearance.BackColor = Drawing.Color.LightYellow
            .Columns("RUN_PCT").CellAppearance.BackColor = Drawing.Color.LightYellow
        End With



        For Each COLUMN_NAME As String In New String() _
            {"ITEM_CATGY_CODE", "ITEM_CATGY_DESC"}
            With grdDPTABCC1.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                .CellAppearance.BackColor = Drawing.Color.Beige
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End With
        Next

        For Each COLUMN_NAME As String In New String() _
            {"ABC_CODE", "ABC_DESC"}
            With grdDPTABCP2.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                .CellAppearance.BackColor = Drawing.Color.Beige
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End With
        Next

        For Each COLUMN_NAME As String In New String() _
            {"COLLECTION_CODE", "COLLECTION_NAME", "COUNT_T"}
            With grdICTCOLLX.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                .CellAppearance.BackColor = Drawing.Color.Beige
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End With
        Next

        COLUMN_NAME = "WEIGHTING"
        ASCMAIN1.Add_Value_List(grdDPTABCC1, COLUMN_NAME, , New String() {":", "R:Retail", "W:Wholesale", "C:Cost", "U:Units"})
        With grdDPTABCC1.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
            .Header.Caption = "Weight" & vbCrLf
            .Style = UltraWinGrid.ColumnStyle.DropDownList
        End With

        With grdDPTABCP2.DisplayLayout.Bands(0)
            .Columns("ABC_MAX_POS").Header.Caption = "Max" & vbCrLf & "Pos"
            .Columns("ABC_MIN_POS").Header.Caption = "Min" & vbCrLf & "Pos"
            .Columns("ABC_PCT_RANGE").Header.Caption = "%" & vbCrLf & "Range"
            .Columns("ABC_MIN_DAYS_SUPPLY").Header.Caption = "Min Days" & vbCrLf & "Supply"
        End With

        Sort_grdColumns(grdDPTABCP2, "ABC_CODE", True)
        Sort_grdColumns(grdDPTABCC1, "ITEM_CATGY_CODE", True)

        Dim ABCs() As String
        ReDim ABCs(26)
        ABCs(0) = ":"
        For L As Integer = Asc("A") To Asc("Z")
            ABCs(L - Asc("A") + 1) = Chr(L) & ":" & Chr(L)
        Next
        ASCMAIN1.Add_Value_List(grdICTCOLLX, "ABC_GROUP", , ABCs)
        grdICTCOLLX.DisplayLayout.Bands(0).Columns("ABC_GROUP").Style = UltraWinGrid.ColumnStyle.DropDownList

        ReDim ABCs(0)
        ABCs(0) = ":"
        For Each rowDPTABCP2 As DataRow In dst.Tables("DPTABCP2").Select("", "ABC_CODE")
            ReDim Preserve ABCs(ABCs.Length)
            ABCs(ABCs.Length - 1) = rowDPTABCP2.Item("ABC_CODE") & ":" & rowDPTABCP2.Item("ABC_CODE")
        Next
        ASCMAIN1.Add_Value_List(grdICTITEMS, "ITEM_ABC_CODE", , ABCs, 1)
        grdICTITEMS.DisplayLayout.Bands(1).Columns("ITEM_ABC_CODE").Style = UltraWinGrid.ColumnStyle.DropDownList

        ASCMAIN1.Add_Value_List(grdICTITEMS, "ITEM_CATGY_CODE")

        Dim YEARs As New List(Of String)
        For Y As Integer = Val(Now.Year) - 1 To Val(Now.Year + 1)
            YEARs.Add(Format(Y, "0000"))
        Next
        Absx1.cbeFor("OPS_YYYY").DataSource = YEARs

        'Dim DPs() As String
        'ReDim DPs(12)
        'DPs(0) = ":"
        'For L As Integer = 1 To 12
        '    DPs(L) = CStr(L) & ":" & CStr(L)
        'Next
        'ASCMAIN1.Add_Value_List(grdDPTABCC1, "PERIODS", , DPs)
        'grdDPTABCC1.DisplayLayout.Bands(0).Columns("PERIODS").Style = UltraWinGrid.ColumnStyle.DropDownList

        'ASCMAIN1.sql = "Select ABC_CODE, ABC_DESC from DPTABCP1"
        'Dim cbeABC_CODE As New UltraWinEditors.UltraComboEditor
        'cbeABC_CODE.DropDownStyle = DropDownStyle.DropDownList
        'cbeABC_CODE.ValueMember = "ABC_CODE"
        'cbeABC_CODE.Visible = True
        'cbeABC_CODE.DataSource = ASCDATA1.GetDataTable
        'cbeABC_CODE.Value = cbeABC_CODE.Items(0)
        'grdDPTABCP2.DisplayLayout.Bands(0).Columns("ABC_CODE").EditorControl = cbeABC_CODE

        ASCMAIN1.Add_Value_List(grdDPTABCP2, "ABC_CODE", "Select ABC_CODE, ABC_CODE from DPTABCP1")
        ASCMAIN1.Add_Value_List(grdICTITEMS, "ITEM_ABC_CODE", , , 1, "Select ABC_CODE, ABC_CODE from DPTABCP1")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("BRAND_CODE")

                If cmbOPS_YYYY.Text = "" Then
                    EMsg &= vbCr & "You Must Specify a Year"
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("DPTABCPP", Absx1.txtFor("BRAND_CODE").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Generate"
                If Val(dst.Tables("ICTCOLLX").Compute("COUNT(SELECTED)", "SELECTED='1'") & "") = 0 Then
                    EMsg &= vbCr & "No Collections were Selected"
                End If

                For Each rowDPTABCC1 As DataRow In dst.Tables("DPTABCC1").Rows
                    Dim ITEM_CATGY_CODE As String = rowDPTABCC1.Item("ITEM_CATGY_CODE")
                    Dim PCT As Decimal = Val(dst.Tables("DPTABCP2").Compute("SUM(ABC_PCT_RANGE)", "ITEM_CATGY_CODE = '" & ITEM_CATGY_CODE & "'") & "")
                    Dim ABCs As Int32 = Val(dst.Tables("DPTABCP2").Compute("COUNT(ABC_CODE)", "ITEM_CATGY_CODE = '" & ITEM_CATGY_CODE & "'") & "")

                    If ABCs <> 0 And PCT <> 100 Then
                        EMsg &= vbCr & "See Category " & ITEM_CATGY_CODE & " - ABC % Ranges do not add up to 100%"
                    End If
                Next

            Case "Update"

                'If Not generated Then
                '    EMsg &= vbcr & "ABC's have not been Generated Yet"
                'End If

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

            Case "Generate"
                Generate()

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Generate").Settings.Enabled = iScreenMode
                '.Groups("Display Options").Visible = ScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)
        UltraTabControl1.Visible = ScreenMode

        grdICTITEMS.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)

        dst.Tables("DPTABCCY").Rows.Clear()
        dst.Tables("ICTITEMS").Rows.Clear()
        dst.Tables("DPTABCP2").Rows.Clear()
        dst.Tables("ICTCOLLX").Rows.Clear()
        EnforceConstraints(True)

        Absx1.txtFor("BRAND_CODE").Text = ""

        Set_generated(False)
    End Sub

    Sub Load_Record()

        Call ASCMAIN1.Progress("Now Loading Data ...")

        Call Save_Header_Fields(UltraGroupBox1)

        Fill_Records("DPTABCP2")

        OPS_YYYY = Absx1.cbeFor("OPS_YYYY").Text

        ReDim YPF(6, 1)
        ReDim YPFD(6)

        For P As Integer = 1 To 6
            Dim YPX As String = ASCMAIN1.Period_Calc(OPS_YYYY & Format(P + IIf(optSEASON.Value = "F", 6, 0), "00"), -1 * ASCMAIN1.PCO + 1)
            YPF(P, 0) = YPX
            YPF(P, 1) = ASCMAIN1.Get_Legend(YPX, False, True)
        Next

        ASCDATA1.ExecuteSQL("Update DPTABCP0 Set ITEM_CATGY_CODE = 'I'")
        ASCDATA1.ExecuteSQL("Update DPTABCP0 Set ITEM_CATGY_CODE = (SELECT ITEM_CATGY_CODE FROM DPTPROJ0 where OPS_YYYY = DPTABCP0.OPS_YYYY and SEASON = DPTABCP0.SEASON and ITEM_CODE = DPTABCP0.ITEM_CODE)")
        ASCDATA1.ExecuteSQL("Update DPTABCP0 Set ITEM_CATGY_CODE = NVL(ITEM_CATGY_CODE,'I')")

        ASCDATA1.ExecuteSQL("Truncate Table " & ICTITEMS)

        ' WARNING - VERY SIMILAR SQL IN Generate_ICTITEMS

        ASCMAIN1.sql = "Insert into " & ICTITEMS & vbCrLf _
        & " (ITEM_CODE, ITEM_DESC, COLLECTION_CODE, ITEM_CATGY_CODE " & vbCrLf _
        & ", ITEM_ABC_CODE_PREV, ITEM_ABC_CODE, ITEM_RETAIL_PRICE, ITEM_COST_STD, ITEM_PRICE" & vbCrLf _
        & ", ABC_GROUP, DEMAND_UNITS, DEMAND_WEIGHTED, RUN_TOTAL, RANK, RUN_PCT)" & vbCrLf _
        & " Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.COLLECTION_CODE" & vbCrLf _
        & ", NVL(DPTPROJ0.ITEM_CATGY_CODE,'I') ITEM_CATGY_CODE" & vbCrLf _
        & ", NVL(DPTABCP0.ITEM_ABC_CODE,ICTITEM1.ITEM_ABC_CODE) ITEM_ABC_CODE_PREV, DPTABCP0.ITEM_ABC_CODE" & vbCrLf _
        & ", DPTABCP0.ITEM_RETAIL_PRICE, DPTABCP0.ITEM_COST_STD" & vbCrLf _
        & ", DPTABCP0.ITEM_PRICE" & vbCrLf _
        & ", NVL(DPTABCP0.ABC_GROUP,'0') ABC_GROUP, DPTABCP0.DEMAND_UNITS, DPTABCP0.DEMAND_WEIGHTED" & vbCrLf _
        & ", DPTABCP0.RUN_TOTAL, DPTABCP0.RANK, DPTABCP0.RUN_PCT" & vbCrLf _
        & " from ICTITEM1,DPTPROJ0,DPTABCP0" & vbCrLf _
        & " where ICTITEM1.COLLECTION_CODE in " & vbCrLf _
        & "(Select COLLECTION_CODE FROM ICTCOLL1 " & vbCrLf _
        & " where BRAND_CODE = '" & Absx1.txtFor("BRAND_CODE").Text & "')" & vbCrLf _
        & "   and DPTPROJ0.OPS_YYYY (+) = '" & OPS_YYYY & "'" & vbCrLf _
        & "   and DPTPROJ0.SEASON (+) = '" & optSEASON.Value & "'" & vbCrLf _
        & "   and DPTPROJ0.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
        & "   and DPTABCP0.OPS_YYYY (+) = '" & OPS_YYYY & "'" & vbCrLf _
        & "   and DPTABCP0.SEASON (+) = '" & optSEASON.Value & "'" & vbCrLf _
        & "   and DPTABCP0.ITEM_CODE (+) = ICTITEM1.ITEM_CODE"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        


        ASCMAIN1.sql = "Select Count (*) from " & ICTITEMS
        Dim RECORDS As Int32 = ASCDATA1.GetDataValue
        If RECORDS = 0 Then
            Generate_ICTITEMS()
            ASCDATA1.ExecuteSQL("Update " & ICTITEMS & " Set ABC_GROUP = '0'")
        End If

        EnforceConstraints(False)
        Fill_Records("DPTABCCY")
        Fill_Records("ICTITEMS")
        Fill_Records("ICTCOLLX", New String() {OPS_YYYY, optSEASON.Value, HFs("BRAND_CODE")})
        EnforceConstraints(True)

        With grdICTITEMS.DisplayLayout.Bands(1)
            .Columns("DEMAND_UNITS").Hidden = (RECORDS = 0)
            .Columns("DEMAND_WEIGHTED").Hidden = (RECORDS = 0)
            .Columns("RUN_TOTAL").Hidden = (RECORDS = 0)
            .Columns("RUN_PCT").Hidden = (RECORDS = 0)
            .Columns("RANK").Hidden = (RECORDS = 0)
        End With

        Setup_DPTABCP2()
        Sort_grdColumns(grdICTITEMS, "RANK,ITEM_ABC_CODE,ITEM_CODE", , 1)

        Call ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        ASCDATA1.ExecuteSQL("Truncate Table " & ICTITEMS)


        Call BeginTrans()

        Update_Record_TDA("DPTABCP2")

        'If generated Then
        Update_Record_TDA("ICTITEMS", "Delete from " & ICTITEMS) ' needed to setadded to all rows

        'ASCDATA1.ExecuteSQL("Delete from DPTABCP0 " _
        '                    & " where OPS_YYYY = '" & OPS_YYYY & "'" _
        '                    & " and SEASON = '" & optSEASON.Value & "'" _
        '                    & " and COLLECTION_CODE in " _
        '                    & " (Select COLLECTION_CODE from ICTCOLL1 " _
        '                    & " where BRAND_CODE = '" & HFs("BRAND_CODE") & "')")

        ASCDATA1.ExecuteSQL("Delete from DPTABCP0 " _
                           & " where OPS_YYYY = '" & OPS_YYYY & "'" _
                           & " and SEASON = '" & optSEASON.Value & "'" _
                           & " and ITEM_CODE in (Select ITEM_CODE from ICTITEM1 where COLLECTION_CODE in " _
                           & " (Select COLLECTION_CODE from ICTCOLL1 " _
                           & " where BRAND_CODE = '" & HFs("BRAND_CODE") & "'))")

        ' THE LINE BELOW IS BECAUSE SOME ITEMS WIND UP WITH NO COLLECTION, OR THE COLLECTION CODE IS DELETED OR CHANGED
        ASCDATA1.ExecuteSQL("Delete from DPTABCP0 " _
                           & " where OPS_YYYY = '" & OPS_YYYY & "'" _
                           & " and SEASON = '" & optSEASON.Value & "'" _
                           & " and ITEM_CODE in (Select ITEM_CODE from " & ICTITEMS & ")")

        ASCDATA1.ExecuteSQL("Insert into DPTABCP0 Select '" & OPS_YYYY & "' OPS_YYYY" _
                             & ", '" & optSEASON.Value & "' SEASON, ITEM_CODE" _
                             & ", COLLECTION_CODE, ITEM_CATGY_CODE" _
                             & ", ITEM_ABC_CODE_PREV, ITEM_ABC_CODE" _
                             & ", ITEM_RETAIL_PRICE, ITEM_COST_STD, ITEM_PRICE" _
                             & ", ABC_GROUP, DEMAND_UNITS, DEMAND_WEIGHTED, RUN_TOTAL" _
                             & ", RANK, RUN_PCT" _
                             & " from " & ICTITEMS)

        ' a version of this stmt appears in TARPEND1
        If ASCMAIN1.CYP >= YPF(1, 0) And ASCMAIN1.CYP <= YPF(6, 0) Then
            ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is Select ICTITEMS.ITEM_CODE, ICTITEMS.ITEM_ABC_CODE" _
            & ", DPTABCP2.ABC_MAX_POS, DPTABCP2.ABC_MIN_POS, DPTABCP2.ABC_MIN_DAYS_SUPPLY" _
            & " from " & ICTITEMS & " ICTITEMS, DPTABCP2" _
            & " where DPTABCP2.ABC_CODE (+) = ICTITEMS.ITEM_ABC_CODE" _
            & "   and DPTABCP2.ITEM_CATGY_CODE (+) = ICTITEMS.ITEM_CATGY_CODE;" _
            & " Begin For R1 in C1 Loop" _
            & "  Update ICTITEM1 Set ITEM_ABC_CODE = R1.ITEM_ABC_CODE" _
            & ", ITEM_POS_MAX = R1.ABC_MAX_POS" _
            & ", ITEM_POS_MIN = R1.ABC_MIN_POS" _
            & ", ITEM_MIN_DAYS_SUPPLY = R1.ABC_MIN_DAYS_SUPPLY" _
            & "   where ITEM_CODE = R1.ITEM_CODE;" _
            & " End Loop; End; End;"
            ASCDATA1.ExecuteSQL()
        End If
        'End If

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

        Load_Popup_Menu(grdICTCOLLX, "BB", "Select All", "Clear All")
        Load_Popup_Menu(grdICTITEMS, "SSB", "Show Filter", "Show GroupBox", "Load from Spreadsheet")

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
                Case "grdDPTITMFX"
                    If grdICTITEMS.Tag = "" Then
                        e.Cancel = True
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


            Case "Load from Spreadsheet"
                If grd.Name = "grdICTITEMS" Then
                    Excel_Import(grd, 1)
                Else
                    Excel_Import(grd)
                End If

        End Select

        Select Case grd.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Select All"
                For Each rowICTCOLLX As DataRow In dst.Tables("ICTCOLLX").Rows
                    rowICTCOLLX.Item("SELECTED") = "1"
                Next

            Case "Clear All"
                For Each rowICTCOLLX As DataRow In dst.Tables("ICTCOLLX").Rows
                    rowICTCOLLX.Item("SELECTED") = "0"
                Next

            Case Else
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
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "BRAND_CODE"
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

        End Select
    End Sub

#End Region

#Region "Excel Upload"

    Overrides Sub Excel_Import_Pre_Process _
    (ByVal grd As UltraWinGrid.UltraGrid, _
     Optional ByRef load_by_table As Boolean = False, _
     Optional ByRef load_handled As Boolean = False, _
     Optional ByRef F As ASFEXCL1 = Nothing)

        If grd.Name = "grdICTITEMS" Then
            'For Each row As DataRow In F.dt.Select
            '    Stop
            'Next

            Dim COLs As New Dictionary(Of String, String)
            For Each gcol As UltraWinGrid.UltraGridColumn In F.grdExcel.DisplayLayout.Bands(0).Columns
                If New String() {"ITEM_CODE", "ITEM_ABC_CODE"}.Contains(gcol.Tag) Then
                    COLs.Add(gcol.Tag, gcol.Key)
                End If
            Next

            If F.grd.Rows.Count = 0 Then
                MsgBox("Nothing to Upload")
                Exit Sub
            Else
                If COLs.Count <> 2 Then
                    MsgBox("Missing Columns")
                    Exit Sub
                End If
            End If

            Dim tblMissing As New DataTable
            tblMissing.Columns.Add("ITEM_CODE")

            For Each grow As UltraWinGrid.UltraGridRow In F.grdExcel.Rows
                Dim ITEM_CODE As String = grow.Cells(COLs("ITEM_CODE")).Text
                Dim rowICTITEMS As DataRow = dst.Tables("ICTITEMS").Rows.Find(ITEM_CODE)
                If rowICTITEMS Is Nothing Then
                    tblMissing.Rows.Add(New String() {ITEM_CODE})
                Else
                    rowICTITEMS.Item("ITEM_ABC_CODE") = grow.Cells(COLs("ITEM_ABC_CODE")).Value
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

    End Sub

    Overrides Sub Excel_Import_Custom_Processing_row _
    (ByVal row As DataRow, ByVal grow As UltraWinGrid.UltraGridRow, _
     Optional ByVal TBL As DataTable = Nothing)

    End Sub
#End Region

    Sub Set_generated(ByVal tf As Boolean)

        generated = tf

        With grdICTITEMS.DisplayLayout.Bands(1)
            .Columns("DEMAND_UNITS").Hidden = Not generated
            .Columns("DEMAND_WEIGHTED").Hidden = Not generated
            .Columns("RUN_TOTAL").Hidden = Not generated
            .Columns("RUN_PCT").Hidden = Not generated
            .Columns("RANK").Hidden = Not generated
        End With

        With UltraExplorerBar1.Groups("Screen Control").Items("Generate")
            If tf Then
                '.Settings.Enabled = DefaultableBoolean.False
            Else
                '.Settings.Enabled = DefaultableBoolean.True
            End If
        End With

        With UltraTabControl1
            If tf Then
            Else
                .SelectedTab = .Tabs("ABC Parameters")
            End If
            '.Tabs("ABC by Item").Enabled = True ' tf
        End With
    End Sub

    Sub Generate()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Generating ABC's")

        Generate_ICTITEMS()

        For Each rowICTCOLLX As DataRow In dst.Tables("ICTCOLLX").Rows
            Dim COLLECTION_CODE As String = rowICTCOLLX.Item("COLLECTION_CODE")
            If rowICTCOLLX.Item("SELECTED") & "" <> "1" Then
                ASCDATA1.ExecuteSQL("Delete from " & ICTITEMS & " ICTITEMS where COLLECTION_CODE = '" & COLLECTION_CODE & "'")
            Else
                Dim ABC_GROUP As String = rowICTCOLLX.Item("ABC_GROUP") & ""
                If ABC_GROUP = "" Then ABC_GROUP = "0"
                ASCDATA1.ExecuteSQL("Update " & ICTITEMS & " ICTITEMS set ABC_GROUP = '" & ABC_GROUP & "' where COLLECTION_CODE = '" & COLLECTION_CODE & "'")
            End If
        Next

        For Each rowDPTABCC1 As DataRow In dst.Tables("DPTABCC1").Rows
            Dim PERIODS As Integer = 6
            Dim ITEM_CATGY_CODE As String = rowDPTABCC1.Item("ITEM_CATGY_CODE")
            If PERIODS = 0 Then
                ASCDATA1.ExecuteSQL("Delete from " & ICTITEMS & " ICTITEMS where ITEM_CATGY_CODE = '" & ITEM_CATGY_CODE & "'")
            Else
                Dim WEIGHTING As String = rowDPTABCC1.Item("WEIGHTING") & ""
                If WEIGHTING = "" Then
                    ASCDATA1.ExecuteSQL("Delete from " & ICTITEMS & " ICTITEMS" _
                            & " where ITEM_CATGY_CODE = '" & ITEM_CATGY_CODE & "'")
                End If
            End If
        Next

        ASCDATA1.ExecuteSQL("Truncate Table " & DPTITMFX)
        Dim sql As String = sqlDPTITMFX
        For P As Int32 = 1 To 6
            sql = Replace(sql, "'" & Format(P, "000000") & "'", "'" & YPF(P, 0) & "'")
        Next
        Dim YP_MIN As String = IIf(YPF(1, 0) < ASCMAIN1.CYP, YPF(1, 0), ASCMAIN1.CYP)
        Dim YP_MAX As String = IIf(YPF(6, 0) < ASCMAIN1.CYP, YPF(6, 0), ASCMAIN1.CYP)
        sql = Replace(sql, "'YP_MIN'", "'" & YP_MIN & "'")
        sql = Replace(sql, "'YP_MAX'", "'" & YP_MAX & "'")
        ASCDATA1.ExecuteSQL("Insert into " & DPTITMFX & " " & sql)

        For Each rowDPTABCC1 As DataRow In dst.Tables("DPTABCC1").Select("WEIGHTING <> ''", "")
            Dim PERIODS As Integer = 6
            Dim ITEM_CATGY_CODE As String = rowDPTABCC1.Item("ITEM_CATGY_CODE")

            Dim DEMAND_UNITS As String = ""
            For P As Integer = 0 To PERIODS
                DEMAND_UNITS &= "+NVL(FC" & Format(P, "00") & ",0)"
            Next
            DEMAND_UNITS = "(" & Mid(DEMAND_UNITS, 2) & ")"

            Dim WEIGHTING As String = rowDPTABCC1.Item("WEIGHTING") & ""
            If WEIGHTING <> "" Then
                Select Case WEIGHTING
                    Case "R"
                        WEIGHTING = "NVL(ITEM_RETAIL_PRICE,0)"
                    Case "W"
                        WEIGHTING = "NVL(ITEM_PRICE,0)"
                    Case "C"
                        WEIGHTING = "NVL(ITEM_COST_STD,0)"
                    Case "U"
                        WEIGHTING = "1"
                End Select

                ASCMAIN1.sql = "" _
                & "BEGIN DECLARE UNITS NUMBER (8,0); " _
                & " CURSOR C1 IS  " _
                & " SELECT ICTITEMS.* " _
                & "  from " & ICTITEMS & " ICTITEMS " _
                & "   where ICTITEMS.ITEM_CATGY_CODE = '" & ITEM_CATGY_CODE & "'" _
                & " For Update; " _
                & "BEGIN FOR R1 IN C1 LOOP " _
                & " Begin " _
                & "Select " & DEMAND_UNITS & " Into UNITS " _
                & " from " & DPTITMFX _
                & " where ITEM_CODE = R1.ITEM_CODE; " _
                & "If SQL%FOUND then " _
                & " Update " & ICTITEMS & " Set DEMAND_UNITS = UNITS" _
                & ", DEMAND_WEIGHTED = NVL(UNITS,0) * " & WEIGHTING _
                & " where Current of C1; " _
                & " End If;" _
                & " Exception " _
                & " when others then null;" _
                & " End;" _
                & "END LOOP; END; END;"
                ASCDATA1.ExecuteSQL()

            End If
        Next

        ASCDATA1.ExecuteSQL("Update " & ICTITEMS & " Set DEMAND_WEIGHTED = NVL(DEMAND_WEIGHTED,0)")

        ASCMAIN1.sql = "" _
        & "BEGIN DECLARE CURSOR C1 IS" _
        & "  SELECT DISTINCT ITEM_CATGY_CODE, ABC_GROUP FROM " & ICTITEMS & ";" _
        & " BEGIN FOR R1 IN C1 LOOP" _
        & "  BEGIN DECLARE " _
        & "   TOTAL NUMBER (15,2);" _
        & "   RNO NUMBER (8,0);" _
        & "   CURSOR C2 IS" _
        & "    SELECT * from " & ICTITEMS & " " _
        & "     where ITEM_CATGY_CODE = R1.ITEM_CATGY_CODE" _
        & "       and ABC_GROUP = R1.ABC_GROUP" _
        & "     order by DEMAND_WEIGHTED DESC;" _
        & "   BEGIN FOR R2 IN C2 LOOP" _
        & "    RNO := NVL(RNO,0) + 1;" _
        & "    TOTAL := NVL(TOTAL,0) + NVL(R2.DEMAND_WEIGHTED,0);" _
        & "    UPDATE " & ICTITEMS & " SET RUN_TOTAL = TOTAL, RANK = RNO" _
        & "   WHERE ITEM_CODE = R2.ITEM_CODE;" _
        & "  END LOOP; END; END;" _
        & "END LOOP; END; END;"
        ASCDATA1.ExecuteSQL()


        EnforceConstraints(False)
        Fill_Records("DPTABCCY")
        Fill_Records("ICTITEMS")
        EnforceConstraints(True)

        Dim ABCs As String = ""
        Dim ABC_pcts() As Decimal

        For Each rowDPTABCCY As DataRow In dst.Tables("DPTABCCY").Rows
            Dim DEMAND_WEIGHTED As Decimal = Val(rowDPTABCCY.Item("DEMAND_WEIGHTED") & "")
            Dim ABC_level As Integer = 0
            Dim ABC_pct As Decimal = 0
            Dim ITEM_CATGY_CODE As String = rowDPTABCCY.Item("ITEM_CATGY_CODE")
            Dim ABC_GROUP As String = rowDPTABCCY.Item("ABC_GROUP")

            ReDim ABC_pcts(0)
            ABCs = ""
            For Each rowDPTABCP2 As DataRow In dst.Tables("DPTABCP2") _
                .Select("ITEM_CATGY_CODE = '" & ITEM_CATGY_CODE & "'", "ABC_CODE")
                ABCs &= rowDPTABCP2.Item("ABC_CODE")
                ReDim Preserve ABC_pcts(Len(ABCs))
                ABC_pcts(Len(ABCs)) = Val(rowDPTABCP2.Item("ABC_PCT_RANGE") & "")
            Next

            For Each rowICTITEMS As DataRow In dst.Tables("ICTITEMS").Select( _
            "ITEM_CATGY_CODE = '" & ITEM_CATGY_CODE & "'" _
            & " and ABC_GROUP = '" & ABC_GROUP & "'", "RANK")
                Dim RUN_TOTAL As Decimal = Val(rowICTITEMS.Item("RUN_TOTAL") & "")
                Dim RUN_PCT As Decimal = 0
                If DEMAND_WEIGHTED <> 0 Then RUN_PCT = 100 * RUN_TOTAL / DEMAND_WEIGHTED
                If RUN_TOTAL > DEMAND_WEIGHTED * ABC_pct / 100 And ABC_level < Len(ABCs) Then
                    ABC_level += 1
                    ABC_pct += ABC_pcts(ABC_level)
                End If
                If ABC_level > 0 Then
                    rowICTITEMS.Item("ITEM_ABC_CODE") = Mid(ABCs, ABC_level, 1)
                End If
                rowICTITEMS.Item("RUN_PCT") = RUN_PCT
            Next
        Next

        Set_generated(True)
        UltraTabControl1.SelectedTab = UltraTabControl1.Tabs("ABC by Item")

        dst.Tables("ICTITEMS").AcceptChanges()
        Sort_grdColumns(grdICTITEMS, "RANK,ITEM_ABC_CODE,ITEM_CODE", , 1)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdICTITEMS_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdICTITEMS.InitializeLayout

    End Sub

    Private Sub grdICTITEMS_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTITEMS.InitializeRow

        If e.Row.Band.Key = "DPTABCCY" Then
        Else
            If e.Row.Cells("ITEM_ABC_CODE").Value & "" = "A" Then
                e.Row.Cells("ITEM_ABC_CODE").Appearance.BackColor = Color.LightGreen
            ElseIf e.Row.Cells("ITEM_ABC_CODE").Value & "" = "B" Then
                e.Row.Cells("ITEM_ABC_CODE").Appearance.BackColor = Color.LightBlue
            ElseIf e.Row.Cells("ITEM_ABC_CODE").Value & "" = "C" Then
                e.Row.Cells("ITEM_ABC_CODE").Appearance.BackColor = Color.Yellow
            Else
                e.Row.Cells("ITEM_ABC_CODE").Appearance.BackColor = Color.Empty
            End If
        End If
    End Sub

    Private Sub grdDPTABCC1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdDPTABCC1.AfterRowActivate
        Setup_DPTABCP2()
    End Sub

    Sub Setup_DPTABCP2()
        Dim ITEM_CATGY_CODE As String = grdDPTABCC1.ActiveRow.Cells("ITEM_CATGY_CODE").Text
        Dim dvw As DataView = DirectCast(grdDPTABCP2.DataSource, DataTable).DefaultView
        dvw.RowFilter = "ITEM_CATGY_CODE = '" & ITEM_CATGY_CODE & "'"
        grdDPTABCP2.Text = "ABC Codes for " & grdDPTABCC1.ActiveRow.Cells("ITEM_CATGY_DESC").Text & " Items"
    End Sub

    Private Sub grdDPTABCP2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdDPTABCP2.AfterRowActivate
        If grdDPTABCP2.ActiveRow.IsAddRow Then
            grdDPTABCP2.ActiveRow.Cells("ABC_CODE").Column.CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            grdDPTABCP2.ActiveRow.Cells("ABC_CODE").Column.CellActivation = UltraWinGrid.Activation.NoEdit
        End If
    End Sub

    Private Sub grdDPTABCP2_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdDPTABCP2.InitializeLayout

    End Sub

#Region "grdDPTABCP2"

    Private Sub grdDPTABCP2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdDPTABCP2.BeforeExitEditMode
        If e.CancellingEditOperation Then Exit Sub
        With grdDPTABCP2.ActiveCell
            Select Case .Column.Key
                Case "ABC_CODE"
                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                        cdr = LookUp("DPTABCP1", New String() {.Text})
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid " & .Column.Header.Caption & "(" & .Text & ")")
                            .Value = ""
                            e.Cancel = True
                        Else
                            .Row.Cells("ABC_DESC").Value = cdr.Item("ABC_DESC")
                            .Row.Cells("ABC_MAX_POS").Value = cdr.Item("ABC_MAX_POS")
                            .Row.Cells("ABC_MIN_POS").Value = cdr.Item("ABC_MIN_POS")
                            .Row.Cells("ABC_PCT_RANGE").Value = cdr.Item("ABC_PCT_RANGE")
                            .Row.Cells("ABC_MIN_DAYS_SUPPLY").Value = cdr.Item("ABC_MIN_DAYS_SUPPLY")
                        End If
                    End If
            End Select
        End With
    End Sub

    Private Sub grdDPTABCP2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdDPTABCP2.BeforeRowUpdate

        If e.Row.Cells("ABC_CODE").Value & "" = "" Then
            e.Cancel = True
        End If
        If Not e.Cancel Then
            If e.Row.Cells("ITEM_CATGY_CODE").Value & "" = "" Then
                e.Row.Cells("ITEM_CATGY_CODE").Value = grdDPTABCC1.ActiveRow.Cells("ITEM_CATGY_CODE").Text
            End If
        End If
    End Sub

    Private Sub grdDPTABCP2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdDPTABCP2.ClickCellButton
        'Dim sql_where As String = ""
        'Select Case grdDPTABCP2.ActiveCell.Column.Key
        '    Case "CUST_STORE_NO"
        '        grdClickCellButton(grdDPTABCP2, "CUST_CODE = '" & HFs("CUST_CODE") & "'")
        'End Select
    End Sub

    Private Sub grdDPTABCP2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdDPTABCP2.AfterExitEditMode
        With grdDPTABCP2
            Select Case .ActiveCell.Column.Key
                Case "ABC_CODE"
                    If .ActiveCell.Text <> "" Then
                        If .ActiveRow.IsAddRow Then
                            '.ActiveRow.Cells("ABC_DESC").Value = .ActiveCell.Text & " Items"
                        End If
                    End If
            End Select
        End With
    End Sub
#End Region

    Sub Generate_ICTITEMS()
        ASCDATA1.ExecuteSQL("Truncate Table " & ICTITEMS)

        ' WARNING - VERY SIMILAR SQL IN Load_Record

        ASCMAIN1.sql = "Insert into " & ICTITEMS & vbCrLf _
        & " (ITEM_CODE, ITEM_DESC, COLLECTION_CODE, ITEM_CATGY_CODE " & vbCrLf _
        & ", ITEM_ABC_CODE_PREV, ITEM_ABC_CODE, ITEM_RETAIL_PRICE, ITEM_COST_STD, ITEM_PRICE)" & vbCrLf _
        & " Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.COLLECTION_CODE" & vbCrLf _
        & ", NVL(DPTPROJ0.ITEM_CATGY_CODE,'I') ITEM_CATGY_CODE" & vbCrLf _
        & ", NVL(DPTABCP0.ITEM_ABC_CODE,ICTITEM1.ITEM_ABC_CODE) ITEM_ABC_CODE_PREV, ICTITEM1.ITEM_ABC_CODE" & vbCrLf _
        & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_COST_STD" & vbCrLf _
        & ", ICTITEM1.ITEM_PRICE" & vbCrLf _
        & " from ICTITEM1,DPTPROJ0,DPTABCP0" & vbCrLf _
        & " where ICTITEM1.COLLECTION_CODE in " & vbCrLf _
        & "(Select COLLECTION_CODE FROM ICTCOLL1 " & vbCrLf _
        & " where BRAND_CODE = '" & Absx1.txtFor("BRAND_CODE").Text & "')" & vbCrLf _
        & "   and DPTPROJ0.OPS_YYYY (+) = '" & OPS_YYYY & "'" & vbCrLf _
        & "   and DPTPROJ0.SEASON (+) = '" & optSEASON.Value & "'" & vbCrLf _
        & "   and DPTPROJ0.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
        & "   and DPTABCP0.OPS_YYYY (+) = '" & OPS_YYYY & "'" & vbCrLf _
        & "   and DPTABCP0.SEASON (+) = '" & optSEASON.Value & "'" & vbCrLf _
        & "   and DPTABCP0.ITEM_CODE (+) = ICTITEM1.ITEM_CODE"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
    End Sub
End Class