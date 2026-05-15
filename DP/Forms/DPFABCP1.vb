Public Class DPFABCP1
    Dim YP() As String
    Dim ICTITEMS As String
    Dim DPTITMFX As String
    Dim sqlDPTITMFX As String
    Dim YPF(,) As String
    Dim YPFD() As Date
    Dim generated As Boolean = False

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        ReDim YPF(25, 1) ' 25 Future Periods
        ReDim YPFD(25)

        Dim P As Integer

        ASCMAIN1.sql = "Select * from GLTPARM2 " _
        & " where OPS_YYYYPP between '" & ASCMAIN1.CYP & "' and '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 25) & "'"
        P = 0
        For Each rowGLTPARM2 As DataRow In ASCDATA1.GetDataTable.Select("", "OPS_YYYYPP")
            YPF(P, 0) = rowGLTPARM2.Item("OPS_YYYYPP")
            YPF(P, 1) = Mid(rowGLTPARM2.Item("LEGEND"), 10, 6)
            YPFD(P) = rowGLTPARM2.Item("PRD_END_DATE")
            P += 1
        Next


        With dst

            ' ICTITEMS - All Items on Screen

            ASCMAIN1.sql = "Select ITEM_CODE, ITEM_DESC, COLLECTION_CODE, ITEM_CATGY_CODE, ITEM_SNU_CODE" _
            & ", ICTITEM1.ITEM_ABC_CODE ITEM_ABC_CODE_PREV, ICTITEM1.ITEM_ABC_CODE, ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_COST_STD" _
            & ", ICTITEM1.ITEM_RETAIL_PRICE * .6 WHOLESALE_PRICE" _
            & " from ICTITEM1 where ROWNUM < 1"
            ICTITEMS = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add Primary Key (ITEM_CODE)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add ABC_GROUP VARCHAR2(1)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add DEMAND_UNITS NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add DEMAND_WEIGHTED NUMBER (13,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add RUN_TOTAL NUMBER (15,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMS & " Add RANK NUMBER (8,0)")


            ' DPTITMFX

            ASCMAIN1.sql = "Select DPTITMF1.ITEM_CODE" & vbCrLf
            For M As Integer = 0 To 12
                Dim YP As String = "000000"
                If M > 0 Then YP = YPF(M - 1, 0)
                ASCMAIN1.sql &= ", SUM (DECODE(DPTITMF1.OPS_YYYYPP_FC,'" & YP & "',DPTITMF1.FORECAST,0)) FC" & Format(M, "00") & vbCrLf
            Next
            ASCMAIN1.sql &= " from " & ICTITEMS & " ICTITEMS,DPTITMF1" & vbCrLf _
            & " where DPTITMF1.ITEM_CODE = ICTITEMS.ITEM_CODE" & vbCrLf _
            & "   and DPTITMF1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" _
            & " group by DPTITMF1.ITEM_CODE"
            sqlDPTITMFX = ASCMAIN1.sql
            DPTITMFX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & DPTITMFX & " Add Primary Key (ITEM_CODE)")


            ' DPTABCCY

            ASCMAIN1.sql = "Select ITEM_CATGY_CODE, ITEM_SNU_CODE, ABC_GROUP" _
            & ", SUM (DEMAND_WEIGHTED) DEMAND_WEIGHTED, COUNT (*) ITEMS" _
            & " from " & ICTITEMS & " ICTITEMS" _
            & " group by ITEM_CATGY_CODE, ITEM_SNU_CODE, ABC_GROUP"
            Create_TDA(.Tables.Add, "DPTABCCY", "**", 0, False, "", 4)

            ' ICTITEMS

            ASCMAIN1.sql = "Select ICTITEMS.* from " & ICTITEMS & " ICTITEMS"
            Create_TDA(.Tables.Add("ICTITEMS"), ICTITEMS, "**", 0, True, "", 1)


            .Relations.Add("DPTABCCY_ICTITEMS" _
            , New DataColumn() _
            {.Tables("DPTABCCY").Columns("ITEM_CATGY_CODE") _
             , .Tables("DPTABCCY").Columns("ITEM_SNU_CODE") _
             , .Tables("DPTABCCY").Columns("ABC_GROUP")} _
            , New DataColumn() _
            {.Tables("ICTITEMS").Columns("ITEM_CATGY_CODE") _
             , .Tables("ICTITEMS").Columns("ITEM_SNU_CODE") _
             , .Tables("ICTITEMS").Columns("ABC_GROUP")})






            ASCMAIN1.sql = "Select ICTCOLL1.COLLECTION_CODE, ICTCOLL1.COLLECTION_NAME" _
            & ", X.COUNT_C, X.COUNT_N, X.COUNT_E, X.COUNT_I, X.COUNT_X" _
            & " from ICTCOLL1," _
            & "(" _
            & " Select ICTITEM1.COLLECTION_CODE" _
            & ", Sum (DECODE(ICTITEM1.ITEM_CATGY_CODE,'C',1,0)) COUNT_C" _
            & ", Sum (DECODE(ICTITEM1.ITEM_CATGY_CODE,'N',1,0)) COUNT_N" _
            & ", Sum (DECODE(ICTITEM1.ITEM_CATGY_CODE,'E',1,0)) COUNT_E" _
            & ", Sum (DECODE(ICTITEM1.ITEM_CATGY_CODE,'I',1,0)) COUNT_I" _
            & ", Sum (DECODE(ICTITEM1.ITEM_CATGY_CODE,'C',0,'N',0,'E',0,'I',0,1)) COUNT_X" _
            & " from ICTITEM1, " & ICTITEMS & " ICTITEMS " _
            & " where ICTITEMS.ITEM_CODE = ICTITEM1.ITEM_CODE " _
            & " group by ICTITEM1.COLLECTION_CODE" _
            & ") X where ICTCOLL1.BRAND_CODE = :PARM1" _
            & " and X.COLLECTION_CODE (+) = ICTCOLL1.COLLECTION_CODE"
            Create_TDA(.Tables.Add, "ICTCOLLX", "**", 0, False, "V", 1)
            .Tables("ICTCOLLX").Columns.Add("COUNT_T", GetType(System.Int32), "ISNULL(COUNT_C,0)+ISNULL(COUNT_N,0)+ISNULL(COUNT_E,0)+ISNULL(COUNT_I,0)+ISNULL(COUNT_X,0)")
            .Tables("ICTCOLLX").Columns.Add("SELECTED")
            .Tables("ICTCOLLX").Columns.Add("ABC_GROUP")

            .Tables("ICTCOLLX").Columns("COUNT_C").DataType = GetType(System.Int32)
            .Tables("ICTCOLLX").Columns("COUNT_N").DataType = GetType(System.Int32)
            .Tables("ICTCOLLX").Columns("COUNT_E").DataType = GetType(System.Int32)
            .Tables("ICTCOLLX").Columns("COUNT_I").DataType = GetType(System.Int32)
            .Tables("ICTCOLLX").Columns("COUNT_X").DataType = GetType(System.Int32)

            Create_TDA(.Tables.Add, "DPTABCP1", "*", 0)

            With .Tables.Add("DPTABCC1")
                .Columns.Add("ITEM_CATGY_CODE")
                .Columns.Add("ITEM_CATGY_DESC")
                .Columns.Add("PERIODS", GetType(System.Int32))
                .Columns.Add("WEIGHTING_S")
                .Columns.Add("WEIGHTING_N")
                .Columns.Add("WEIGHTING_U")
            End With


            Create_TDA(.Tables.Add, "DPTITMF1", "*")

        End With


        With dst.Tables("DPTABCC1").Rows
            .Add(New String() {"C", "Core", "6", "R", "", ""})
            .Add(New String() {"N", "New", "6", "R", "", ""})
            .Add(New String() {"E", "Existing", "6", "R", "", ""})
            .Add(New String() {"I", "Inactive", "0", "R", "", ""})
        End With

        Fill_Records("DPTABCP1")

        grdDPTABCC1.DataSource = dst.Tables("DPTABCC1")
        grdDPTABCP1.DataSource = dst.Tables("DPTABCP1")
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

        Call Create_Summary(grdDPTABCP1, "ABC_CODE", "Count")
        Call Create_Summary(grdDPTABCP1, "ABC_PCT_RANGE")

        Call Create_Summary(grdICTCOLLX, "COLLECTION_CODE", "Count")
        Call Create_Summary(grdICTCOLLX, "SELECTED")
        Call Create_Summary(grdICTCOLLX, "COUNT_C")
        Call Create_Summary(grdICTCOLLX, "COUNT_N")
        Call Create_Summary(grdICTCOLLX, "COUNT_E")
        Call Create_Summary(grdICTCOLLX, "COUNT_I")
        Call Create_Summary(grdICTCOLLX, "COUNT_X")
        Call Create_Summary(grdICTCOLLX, "COUNT_T")

        With grdDPTABCP1.DisplayLayout.Bands("DPTABCP1")
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
            With grdDPTABCP1.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
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

        For Each W As String In New String() {"S", "N", "U"}
            Dim COLUMN_NAME As String = "WEIGHTING_" & W
            ASCMAIN1.Add_Value_List(grdDPTABCC1, COLUMN_NAME, , New String() {":", "R:Retail", "W:Wholesale", "C:Cost", "U:Units"})
            With grdDPTABCC1.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                .Header.Caption = "Weight" & vbCrLf & "By (" & W & ")"
                .Style = UltraWinGrid.ColumnStyle.DropDownList
            End With

        Next
        grdDPTABCC1.DisplayLayout.Bands(0).Columns("PERIODS").Header.Caption = "Demand" & vbCrLf & "Periods"

        With grdDPTABCP1.DisplayLayout.Bands(0)
            .Columns("ABC_MAX_POS").Header.Caption = "Max" & vbCrLf & "Pos"
            .Columns("ABC_MIN_POS").Header.Caption = "Min" & vbCrLf & "Pos"
            .Columns("ABC_PCT_RANGE").Header.Caption = "%" & vbCrLf & "Range"
            .Columns("ABC_MIN_DAYS_SUPPLY").Header.Caption = "Min Days" & vbCrLf & "Supply"
        End With

        Sort_grdColumns(grdDPTABCP1, "ABC_CODE", True)
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
        For Each rowDPTABCP1 As DataRow In dst.Tables("DPTABCP1").Select("", "ABC_CODE")
            ReDim Preserve ABCs(ABCs.Length)
            ABCs(ABCs.Length - 1) = rowDPTABCP1.Item("ABC_CODE") & ":" & rowDPTABCP1.Item("ABC_CODE")
        Next
        ASCMAIN1.Add_Value_List(grdICTITEMS, "ITEM_ABC_CODE", , ABCs, 1)
        grdICTITEMS.DisplayLayout.Bands(1).Columns("ITEM_ABC_CODE").Style = UltraWinGrid.ColumnStyle.DropDownList

        ASCMAIN1.Add_Value_List(grdICTITEMS, "ITEM_CATGY_CODE")

        Dim DPs() As String
        ReDim DPs(12)
        DPs(0) = ":"
        For L As Integer = 1 To 12
            DPs(L) = CStr(L) & ":" & CStr(L)
        Next
        ASCMAIN1.Add_Value_List(grdDPTABCC1, "PERIODS", , DPs)
        grdDPTABCC1.DisplayLayout.Bands(0).Columns("PERIODS").Style = UltraWinGrid.ColumnStyle.DropDownList

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("BRAND_CODE")

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("DPTABCPP", Absx1.txtFor("BRAND_CODE").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Generate"
                If Val(dst.Tables("ICTCOLLX").Compute("COUNT(SELECTED)", "SELECTED='1'") & "") = 0 Then
                    EMsg &= "No Collections were Selected"
                End If

                If Val(dst.Tables("DPTABCP1").Compute("SUM(ABC_PCT_RANGE)", "") & "") <> 100 Then
                    EMsg &= "ABC % Ranges must add up to 100%"
                End If

            Case "Update"

                If Not generated Then
                    EMsg &= "ABC's have not been Generated Yet"
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

            Case "Generate"
                Call Generate()

            Case "Update"
                Call Update_Record()
                Call Mode_Settings(False)

            Case "Cancel", "Done"
                Call Mode_Settings(False)

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

        dst.EnforceConstraints = False
        dst.Tables("ICTITEMS").Rows.Clear()
        dst.EnforceConstraints = True

        Absx1.txtFor("BRAND_CODE").Text = ""

        Set_generated(False)
    End Sub

    Sub Load_Record()

        Call ASCMAIN1.Progress("Now Loading Data ...")

        Call Save_Header_Fields(UltraGroupBox1)

        ASCDATA1.ExecuteSQL("Truncate Table " & ICTITEMS)


        ASCMAIN1.sql = "Insert into " & ICTITEMS & vbCrLf _
        & " (ITEM_CODE, ITEM_DESC, COLLECTION_CODE, ITEM_CATGY_CODE, ITEM_SNU_CODE " & vbCrLf _
        & ", ITEM_ABC_CODE_PREV, ITEM_ABC_CODE, ITEM_RETAIL_PRICE, ITEM_COST_STD, WHOLESALE_PRICE)" & vbCrLf _
        & " Select ITEM_CODE, ITEM_DESC, COLLECTION_CODE, ITEM_CATGY_CODE, ITEM_SNU_CODE" & vbCrLf _
        & ", ITEM_ABC_CODE ITEM_ABC_CODE_PREV, ITEM_ABC_CODE, ITEM_RETAIL_PRICE, ITEM_COST_STD" & vbCrLf _
        & ", (0.6 * ITEM_RETAIL_PRICE) WHOLESALE_PRICE" & vbCrLf _
        & " from ICTITEM1 where COLLECTION_CODE in " & vbCrLf _
        & "(Select COLLECTION_CODE FROM ICTCOLL1 " & vbCrLf _
        & " where BRAND_CODE = '" & Absx1.txtFor("BRAND_CODE").Text & "')"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        Fill_Records("ICTCOLLX", HFs("BRAND_CODE"))


        'dst.EnforceConstraints = False
        'Fill_Records("ICTITEMS")
        'Sort_grdColumns(grdICTITEMS, "ITEM_CODE")
        'dst.EnforceConstraints = True

        Call ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Call BeginTrans()

        'For Each rowICTITEMS As DataRow In dst.Tables("ICTITEMS") _
        '.Select("", "", DataViewRowState.ModifiedCurrent)
        '    Dim ITEM_CODE As String = rowICTITEMS.Item("ITEM_CODE")
        '    Dim ITEM_ABC_CODE As String = rowICTITEMS.Item("ITEM_ABC_CODE")
        '    ASCDATA1.ExecuteSQL("Update " & ICTITEMS & " set ITEM_ABC_CODE = '" & ITEM_ABC_CODE & "' where ITEM_CODE = '" & ITEM_CODE & "'")
        'Next

        Update_Record_TDA("ICTITEMS", "Delete from " & ICTITEMS)

        Update_Record_TDA("DPTABCP1")

        ASCMAIN1.sql = "" _
        & "Begin Declare Cursor C1 is Select ICTITEMS.ITEM_CODE, ICTITEMS.ITEM_ABC_CODE" _
        & ", DPTABCP1.ABC_MAX_POS, DPTABCP1.ABC_MIN_POS, DPTABCP1.ABC_MIN_DAYS_SUPPLY" _
        & " from " & ICTITEMS & " ICTITEMS, DPTABCP1" _
        & " where DPTABCP1.ABC_CODE (+) = ICTITEMS.ITEM_ABC_CODE;" _
        & " Begin For R1 in C1 Loop" _
        & "  Update ICTITEM1 Set ITEM_ABC_CODE = R1.ITEM_ABC_CODE" _
        & ", ITEM_POS_MAX = R1.ABC_MAX_POS" _
        & ", ITEM_POS_MIN = R1.ABC_MIN_POS" _
        & ", ITEM_MIN_DAYS_SUPPLY = R1.ABC_MIN_DAYS_SUPPLY" _
        & "   where ITEM_CODE = R1.ITEM_CODE;" _
        & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

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

    Overrides Sub Load_Popup_Menus()
        'Call Load_Popup_Menu(grdDPTITMFX, "BB", "Clear Column", "Copy Value")
        Call Load_Popup_Menu(grdICTCOLLX, "BB", "Select All", "Clear All")

    End Sub

    Private Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs) Handles tlb.BeforeToolDropdown

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If e.Tool.Key <> "grdSATCSLSS" Then
            '    Dim tlb_sbt As UltraWinToolbars.StateButtonTool
            '    tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            '    tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
            '    tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            '    tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
            'End If

            Select Case e.SourceControl.Name
                Case "grdDPTITMFX"
                    If grdICTITEMS.Tag = "" Then
                        e.Cancel = True
                    End If
            End Select

        End If
    End Sub

    Private Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs) Handles tlb.ToolClick
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

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
                'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select
    End Sub

    Private Sub tlb_ToolValueChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolEventArgs) Handles tlb.ToolValueChanged
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
            Case "BRAND_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("Load", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "BRAND_CODE"
                Call Click_Command("Load")
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

    Sub Set_generated(ByVal tf As Boolean)

        generated = tf

        With grdICTITEMS.DisplayLayout.Bands(1)
            .Columns("DEMAND_UNITS").Hidden = Not generated
            .Columns("DEMAND_WEIGHTED").Hidden = Not generated
            .Columns("RUN_TOTAL").Hidden = Not generated
            .Columns("ABC_GROUP").Hidden = Not generated
            .Columns("RANK").Hidden = Not generated
        End With

        With UltraExplorerBar1.Groups("Screen Control").Items("Generate")
            If tf Then
                .Settings.Enabled = DefaultableBoolean.False
            Else
                '.Settings.Enabled = DefaultableBoolean.True
            End If
        End With

        With UltraTabControl1
            If tf Then
            Else
                .SelectedTab = .Tabs("ABC Parameters")
            End If
            .Tabs("ABC by Item").Enabled = tf
        End With
    End Sub


    Sub Generate()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Generating ABC's")

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
            Dim PERIODS As Integer = Val(rowDPTABCC1.Item("PERIODS") & "")
            Dim ITEM_CATGY_CODE As String = rowDPTABCC1.Item("ITEM_CATGY_CODE")
            If PERIODS = 0 Then
                ASCDATA1.ExecuteSQL("Delete from " & ICTITEMS & " ICTITEMS where ITEM_CATGY_CODE = '" & ITEM_CATGY_CODE & "'")
            Else
                For Each ITEM_SNU_CODE In New String() {"S", "N", "U"}
                    Dim WEIGHTING As String = rowDPTABCC1.Item("WEIGHTING_" & ITEM_SNU_CODE) & ""
                    If WEIGHTING = "" Then
                        ASCDATA1.ExecuteSQL("Delete from " & ICTITEMS & " ICTITEMS where ITEM_CATGY_CODE = '" & ITEM_CATGY_CODE & "' and ITEM_SNU_CODE = '" & ITEM_SNU_CODE & "'")
                    End If
                Next
            End If
        Next

        ASCDATA1.ExecuteSQL("Truncate Table " & DPTITMFX)
        ASCDATA1.ExecuteSQL("Insert into " & DPTITMFX & " " & sqlDPTITMFX)

        For Each rowDPTABCC1 As DataRow In dst.Tables("DPTABCC1").Select("PERIODS > 0", "")
            Dim PERIODS As Integer = Val(rowDPTABCC1.Item("PERIODS") & "")
            Dim ITEM_CATGY_CODE As String = rowDPTABCC1.Item("ITEM_CATGY_CODE")

            Dim DEMAND_UNITS As String = ""
            For P As Integer = 0 To PERIODS
                DEMAND_UNITS &= "+NVL(FC" & Format(P, "00") & ",0)"
            Next
            DEMAND_UNITS = "(" & Mid(DEMAND_UNITS, 2) & ")"

            For Each ITEM_SNU_CODE In New String() {"S", "N", "U"}
                Dim WEIGHTING As String = rowDPTABCC1.Item("WEIGHTING_" & ITEM_SNU_CODE) & ""
                If WEIGHTING <> "" Then
                    Select Case WEIGHTING
                        Case "R"
                            WEIGHTING = "NVL(ITEM_RETAIL_PRICE,0)"
                        Case "W"
                            WEIGHTING = "NVL(WHOLESALE_PRICE,0)"
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
                    & "     and ICTITEMS.ITEM_SNU_CODE = '" & ITEM_SNU_CODE & "'" _
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
        Next

        ASCDATA1.ExecuteSQL("Update " & ICTITEMS & " Set DEMAND_WEIGHTED = NVL(DEMAND_WEIGHTED,0)")

        ASCMAIN1.sql = "" _
        & "BEGIN DECLARE CURSOR C1 IS" _
        & "  SELECT DISTINCT ITEM_CATGY_CODE, ITEM_SNU_CODE, ABC_GROUP FROM " & ICTITEMS & ";" _
        & " BEGIN FOR R1 IN C1 LOOP" _
        & "  BEGIN DECLARE " _
        & "   TOTAL NUMBER (15,2);" _
        & "   RNO NUMBER (8,0);" _
        & "   CURSOR C2 IS" _
        & "    SELECT * from " & ICTITEMS & " " _
        & "     where ITEM_CATGY_CODE = R1.ITEM_CATGY_CODE" _
        & "       and ITEM_SNU_CODE = R1.ITEM_SNU_CODE" _
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

        ReDim ABC_pcts(dst.Tables("DPTABCP1").Rows.Count)
        For Each rowDPTABCP1 As DataRow In dst.Tables("DPTABCP1").Select("", "ABC_CODE")
            ABCs &= rowDPTABCP1.Item("ABC_CODE")
            ABC_pcts(Len(ABCs)) = rowDPTABCP1.Item("ABC_PCT_RANGE")
        Next
        For Each rowDPTABCCY As DataRow In dst.Tables("DPTABCCY").Rows
            Dim DEMAND_WEIGHTED As Decimal = Val(rowDPTABCCY.Item("DEMAND_WEIGHTED") & "")
            Dim ABC_level As Integer = 0
            Dim ABC_pct As Decimal = 0
            Dim ITEM_CATGY_CODE As String = rowDPTABCCY.Item("ITEM_CATGY_CODE")
            Dim ITEM_SNU_CODE As String = rowDPTABCCY.Item("ITEM_SNU_CODE")
            Dim ABC_GROUP As String = rowDPTABCCY.Item("ABC_GROUP")
            For Each rowICTITEMS As DataRow In dst.Tables("ICTITEMS").Select( _
            "ITEM_CATGY_CODE = '" & ITEM_CATGY_CODE & "'" _
            & " and ITEM_SNU_CODE = '" & ITEM_SNU_CODE & "'" _
            & " and ABC_GROUP = '" & ABC_GROUP & "'", "RANK")
                Dim RUN_TOTAL As Decimal = Val(rowICTITEMS.Item("RUN_TOTAL") & "")
                If RUN_TOTAL > DEMAND_WEIGHTED * ABC_pct / 100 And ABC_level < Len(ABCs) Then
                    ABC_level += 1
                    ABC_pct += ABC_pcts(ABC_level)
                End If
                If ABC_level > 0 Then
                    rowICTITEMS.Item("ITEM_ABC_CODE") = Mid(ABCs, ABC_level, 1)
                End If
            Next
        Next

        Set_generated(True)
        UltraTabControl1.SelectedTab = UltraTabControl1.Tabs("ABC by Item")

        grdICTITEMS.DisplayLayout.Bands(1).Columns("ABC_GROUP").Hidden = True

        dst.Tables("ICTITEMS").AcceptChanges()

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
End Class