Public Class SAFBUDG1
    Dim rowARTCUST1 As DataRow
    Dim CUST_CODE As String
    Dim BRAND_CODE As String
    Dim SREP_CODE As String
    Dim SATBUDGX As String
    Dim MOS As Integer
    Dim YPs() As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            ASCMAIN1.sql = "Select *" & vbCrLf _
                & " from ARTCUST1 where CUST_CODE = :PARM1 or :PARM1 is Null"
            Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select *" & vbCrLf _
                & " from ICTCOLL1 where BRAND_CODE = :PARM1 or :PARM1 is Null"
            Create_TDA(.Tables.Add, "ICTCOLL1", "**", 0, False, "V", 1)

            Create_TDA(.Tables.Add, "SATBUDG1", "*")
            Create_BAs("SATBUDG1")

            Create_Work_Tables()

            ASCMAIN1.sql = "Select * from " & SATBUDGX
            Create_TDA(.Tables.Add, "SATBUDGX", "**", 0, False, "", 3)

            Create_Relation("ARTCUST1", "SATBUDGX", "CUST_CODE")
            With .Tables("SATBUDGX").Columns
                .Add("TOTAL", GetType(System.Decimal), _
                      "ISNULL(BUDGET_P01,0)+ISNULL(BUDGET_P02,0)+ISNULL(BUDGET_P03,0)+" _
                    & "ISNULL(BUDGET_P04,0)+ISNULL(BUDGET_P05,0)+ISNULL(BUDGET_P06,0)+" _
                    & "ISNULL(BUDGET_P07,0)+ISNULL(BUDGET_P08,0)+ISNULL(BUDGET_P09,0)+" _
                    & "ISNULL(BUDGET_P10,0)+ISNULL(BUDGET_P11,0)+ISNULL(BUDGET_P12,0)")
                .Add("CUST_NAME", GetType(System.String), "PARENT.CUST_NAME")
                .Add("SREP_CODE", GetType(System.String), "PARENT.SREP_CODE")
            End With
        End With

        grdSATBUDGX.DataSource = dst.Tables("SATBUDGX")

        'Create_Summary(grdSAFBUDGX, "CUST_CODE", "Count")
        For M As Integer = 1 To 12
            Create_Summary(grdSATBUDGX, "BUDGET_P" & Format(M, "00"), , , "###,##0")
        Next
        Create_Summary(grdSATBUDGX, "TOTAL", , , "###,##0")

        'Show_Filter(grdSATBUDGX)
        With grdSATBUDGX.DisplayLayout.Bands("SATBUDGX")
            For Each COLUMN_NAME As String In New String() _
                {"COLLECTION_CODE", "CUST_CODE", "ITEM_BASIC_PROMO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next

            For i As Integer = 1 To 12
                Dim COLUMN_NAME As String = "BUDGET_P" & Format(i, "00")
                .Columns(COLUMN_NAME).Format = "#,##0"
            Next
            .Columns("TOTAL").Format = "#,##0"

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = System.Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.Gold
                If New String() {"CUST_CODE", "ITEM_BASIC_PROMO", "COLLECTION_CODE", "TOTAL", "CUST_NAME", "SREP_CODE"}.Contains(gcol.Key) Then
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    If gcol.Key <> "TOTAL" Then gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
        End With

        Dim YY As String = Mid(ASCMAIN1.CYM, 3, 2)
        If Mid(ASCMAIN1.CYM, 5, 2) >= "02" And Mid(ASCMAIN1.CYM, 5, 2) <= "07" Then
            cmdS0.Text = "S" & YY
            cmdS1.Text = "F" & YY
            cmdS2.Text = "S" & Format(Val(YY) + 1, "00")
        Else
            If Mid(ASCMAIN1.CYM, 5, 2) = "01" Then YY = Format(Val(YY) - 1, "00")
            cmdS0.Text = "F" & YY
            cmdS1.Text = "S" & Format(Val(YY) + 1, "00")
            cmdS2.Text = "F" & Format(Val(YY) + 1, "00")
        End If

        optCalendar.Value = "O"
        ASCMAIN1.Add_Value_List(grdSATBUDGX, "ITEM_BASIC_PROMO")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                If MOS < 1 Or MOS > 12 Then
                    EMsg &= vbCr & "Period Range must span between 1 and 12 month"
                Else
                    Validate_Code("CUST_CODE", , True)
                    Validate_Code("BRAND_CODE", , True)
                    Validate_Code("SREP_CODE", , True)
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("SATBUDG1", "*") Then
                        Exit Sub
                    End If
                End If

            Case "Update"
                For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SATBUDGX"), New String() {"CUST_CODE"}).Select("")
                    Dim CUST_CODE As String = row.Item(0)
                    Dim rowARTCUST1 As DataRow = dst.Tables("ARTCUST1").Rows.Find(New String() {CUST_CODE})
                    If rowARTCUST1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Customer (" & CUST_CODE & ")"
                    Else
                        If SREP_CODE <> "" Then
                            If rowARTCUST1.Item("SREP_CODE") <> SREP_CODE Then
                                EMsg &= vbCr & "Customer (" & CUST_CODE & ") does not belong to Sales Rep " & SREP_CODE
                            End If
                        End If
                    End If
                Next

                For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SATBUDGX"), New String() {"COLLECTION_CODE"}).Select("")
                    Dim COLLECTION_CODE As String = row.Item(0)
                    Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(New String() {COLLECTION_CODE})
                    If rowICTCOLL1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Collection (" & COLLECTION_CODE & ")"
                    Else
                        If BRAND_CODE <> "" Then
                            If rowICTCOLL1.Item("BRAND_CODE") & "" <> BRAND_CODE Then
                                EMsg &= vbCr & "Collection (" & COLLECTION_CODE & ") does not belong to Brand " & BRAND_CODE
                            End If
                        End If
                    End If
                Next

            Case "Import from XLS"

                If MsgBox("This function will Import Gross Shipment Budget data" _
                & vbCrLf & " from a specifically formatted spreadsheet" _
                & vbCrLf & " and use that data to replace the data currently on file" _
                & vbCrLf _
                & vbCrLf & " for the Period Range from " & Absx1.cmbFor("RYP0").Text & " thru " & Absx1.cmbFor("RYP1").Text _
                & vbCrLf _
                & vbCrLf & IIf(BRAND_CODE = "", " for All Brands", " for Brand " & BRAND_CODE & " only") _
                & vbCrLf & IIf(CUST_CODE = "", " for All Customers", " for Customer " & CUST_CODE & " only") _
                & vbCrLf & IIf(SREP_CODE = "", " for All Sales Reps", " for Customers connected to Sales Rep " & SREP_CODE & " only") _
                & vbCrLf _
                & vbCrLf & "Once you click 'Yes' to proceed," _
                & vbCrLf & " you will be asked for the location of the spreadsheet, " _
                & vbCrLf & " and the data will be imported and displayed in the grid below." _
                & vbCrLf _
                & vbCrLf & "You will have an opportunity to review it before clicking 'Update'." _
                & vbCrLf _
                & vbCrLf & "Proceed with the Import?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
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

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Import from XLS"

                With grdSATBUDGX.DisplayLayout.Bands(0)
                    For Each COLUMN_NAME As String In New String() {"COLLECTION_CODE", "CUST_CODE", "ITEM_BASIC_PROMO"}
                        .Columns(COLUMN_NAME).Hidden = False
                        .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                    Next
                End With

                Excel_Import_SG(grdSATBUDGX)
                Sort_grdColumns(grdSATBUDGX, "COLLECTION_CODE,CUST_CODE,ITEM_BASIC_PROMO")
                Setup_grd()
                optCOLLECTION_CODE.CheckedIndex = 0
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
                    .Items("Import from XLS").Visible = ScreenMode '  Not tf

                    .Items("Show All").Visible = False ' Not tf

                End With

                .Groups("Display Options").Visible = ScreenMode And (EntryMode = "L")
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        grdSATBUDGX.Visible = ScreenMode

        With grdSATBUDGX.DisplayLayout.Bands(0)
            .Columns("CUST_CODE").Hidden = ScreenMode And (EntryMode = "L")
            .Columns("COLLECTION_CODE").Hidden = ScreenMode And (EntryMode = "L")
            .Columns("SREP_CODE").Hidden = ScreenMode And (EntryMode = "L")
        End With

        If ScreenMode Then
            Setup_grd()
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ARTCUST1", "SATBUDGX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        CUST_CODE = ""
        BRAND_CODE = ""
        SREP_CODE = ""

        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("BRAND_CODE").Text = ""
        Absx1.txtFor("SREP_CODE").Text = ""

        If Mid(ASCMAIN1.CYM, 5, 2) >= "01" And Mid(ASCMAIN1.CYM, 5, 2) <= "06" Then
            Set_Period("S" & Mid(ASCMAIN1.CYM, 3, 2))
        Else
            Set_Period("F" & Mid(ASCMAIN1.CYM, 3, 2))
        End If

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)
        CUST_CODE = HFs("CUST_CODE")
        BRAND_CODE = HFs("BRAND_CODE")
        SREP_CODE = HFs("SREP_CODE")

        EnforceConstraints(False)

        If CUST_CODE <> "" Then
            rowARTCUST1 = Fill_Record("ARTCUST1", CUST_CODE)
        Else
            Fill_Records("ARTCUST1", CUST_CODE)
        End If

        Create_Work_Tables()
        Fill_Records("SATBUDGX")

        Fill_Records("ICTCOLL1", BRAND_CODE)

        EnforceConstraints(True)

        Sort_grdColumns(grdSATBUDGX, "CUST_CODE,COLLECTION_CODE,ITEM_BASIC_PROMO")

        ASCMAIN1.sql = "Select ICTCOLL1.COLLECTION_CODE, ICTCOLL1.COLLECTION_NAME" & vbCrLf _
            & " from ICTCOLL1,ICTBRAN1" & vbCrLf _
            & " where ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" _
            & "   and ICTCOLL1.COLLECTION_STATUS = 'A'" & vbCrLf _
            & "   and ICTBRAN1.BRAND_STATUS = 'A'" & vbCrLf _
            & IIf(BRAND_CODE = "", "", " and ICTCOLL1.BRAND_CODE = '" & BRAND_CODE & "'" & vbCrLf) _
            & " order by ICTCOLL1.COLLECTION_CODE"
        cbeCOLLECTION_CODE.DataSource = ASCDATA1.GetDataTable
        'cbeCOLLECTION_CODE.DataSource = dst.Tables("ICTCOLL1")
        cbeCOLLECTION_CODE.Value = cbeCOLLECTION_CODE.Items(0)

        Set_Month_Headings()

        Setup_grd()

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        Dim RYP0 As String = Absx1.cmbFor("RYP0").Value
        Dim RYP1 As String = Absx1.cmbFor("RYP1").Value

        Dim sql_Delete As String = "Delete from SATBUDG1" _
            & " where OPS_YYYYPP between '" & RYP0 & "' and '" & RYP1 & "'" _
            & IIf(BRAND_CODE = "", "", " and COLLECTION_CODE in (Select COLLECTION_CODE from ICTCOLL1 where BRAND_CODE = '" & BRAND_CODE & "')" & vbCrLf) _
            & IIf(CUST_CODE = "", "", " and CUST_CODE = '" & CUST_CODE & "'" & vbCrLf) _
            & IIf(SREP_CODE = "", "", " and (CUST_CODE) in (Select CUST_CODE from ARTCUST1 where SREP_CODE = '" & SREP_CODE & "')" & vbCrLf)

        dst.Tables("SATBUDG1").Rows.Clear()

        For Each rowSATBUDGX As DataRow In dst.Tables("SATBUDGX").Select("")
            For I As Integer = 1 To MOS
                Dim BUDGET As Decimal = Val(rowSATBUDGX.Item("BUDGET_P" & Format(I, "00")) & "")
                If BUDGET <> 0 Then
                    Dim rowSATBUDG1 As DataRow = dst.Tables("SATBUDG1").NewRow
                    rowSATBUDG1.Item("COLLECTION_CODE") = rowSATBUDGX.Item("COLLECTION_CODE")
                    rowSATBUDG1.Item("CUST_CODE") = rowSATBUDGX.Item("CUST_CODE")
                    rowSATBUDG1.Item("ITEM_BASIC_PROMO") = rowSATBUDGX.Item("ITEM_BASIC_PROMO")
                    rowSATBUDG1.Item("OPS_YYYYPP") = YPs(I)
                    rowSATBUDG1.Item("ITEM_CATGY_CODE") = "E"
                    rowSATBUDG1.Item("BUDGET") = BUDGET
                    dst.Tables("SATBUDG1").Rows.Add(rowSATBUDG1)
                End If
            Next
        Next

        ' Update_Record_TDA("SATBUDG1", sql_Delete)

        ASCMAIN1.sql = sql_Delete
        ASCDATA1.ExecuteSQL()
        Update_BAs("SATBUDG1")

        If ASCMAIN1.CLIENT = "INT" Then

            Dim rowSOTCHAN1 As DataRow = LookUp("SOTCHAN1", "1")
            Dim CUST_CODE As String = rowSOTCHAN1.Item("CUST_CODE")

            ASCMAIN1.sql = "Delete from SATBUDG1" & vbCrLf _
                & " where OPS_YYYYPP >= '" & YPs(1) & "' and OPS_YYYYPP <= '" & YPs(MOS) & "'" & vbCrLf _
                & "   and CUST_CODE = '" & CUST_CODE & "'"
            ASCDATA1.ExecuteSQL()

            Stop ' NEXT LINE NEEDS TO PULL SALES OUT OF GLTACCT2 AND NOT OUT OF RSTBUDF1
            ASCMAIN1.sql = "Insert into SATBUDG1" & vbCrLf _
                & "Select OPS_YYYYPP, '" & CUST_CODE & "' CUST_CODE, 'B' ITEM_BASIC_PROMO" & vbCrLf _
                & ", COLLECTION_CODE, 'E' ITEM_CATGY_CODE, FIN - SLS BUDGET from (" & vbCrLf _
                & "Select COLLECTION_CODE, OPS_YYYYPP, SUM (SLS) SLS, SUM (FIN) FIN from (" & vbCrLf _
                & "Select COLLECTION_CODE, OPS_YYYYPP, SUM (BUDGET) SLS, 0 FIN from SATBUDG1 " & vbCrLf _
                & " where OPS_YYYYPP >= '" & YPs(1) & "' AND OPS_YYYYPP <= '" & YPs(MOS) & "'" & vbCrLf _
                & " group by COLLECTION_CODE, OPS_YYYYPP" & vbCrLf _
                & " union " & vbCrLf _
                & "Select COLLECTION_CODE, OPS_YYYYPP, 0 SLS, SUM (BUDGET) FIN FROM RSTBUDF1 " & vbCrLf _
                & " where OPS_YYYYPP >= '" & YPs(1) & "' AND OPS_YYYYPP <= '" & YPs(MOS) & "'" & vbCrLf _
                & " group by COLLECTION_CODE, OPS_YYYYPP" & vbCrLf _
                & ") group by COLLECTION_CODE, OPS_YYYYPP" & vbCrLf _
                & ")"
            ASCDATA1.ExecuteSQL()

        End If

        CommitTrans("Update Complete")
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup_Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSATBUDGX, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdSATBUDGX"
                    'tlb_btn = DirectCast(tlb_pop.Tools("Clear Column"), UltraWinToolbars.ButtonTool)
                    'tlb_btn.SharedProps.Visible = ScreenMode And (EntryMode = "L")
                    'tlb_btn = DirectCast(tlb_pop.Tools("Copy Value"), UltraWinToolbars.ButtonTool)
                    'tlb_btn.SharedProps.Visible = ScreenMode And (EntryMode = "L")

                    If grdSATBUDGX.Tag = "" Then
                        'e.Cancel = True
                    End If
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case Else
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"
    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Load", e)
                End If
            Case "BRAND_CODE"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Load", e)
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "CUST_CODE"
            '    Click_Command("Load")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                        LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

            Case "BRAND_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("BRAND_CODE").Text <> "" Then
                        LookUp("ICTBRAN1", Absx1.txtFor("BRAND_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

            Case "SREP_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("SREP_CODE").Text <> "" Then
                        LookUp("SOTSREP1", Absx1.txtFor("SREP_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

        End Select
    End Sub

    Public Overrides Sub cmb_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.cmb_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        If Absx1.cmbFor("RYP0").Value & "" <> "" And Absx1.cmbFor("RYP1").Value & "" <> "" Then
            MOS = 1 + ASCMAIN1.Period_Diff(Absx1.cmbFor("RYP0").Value, Absx1.cmbFor("RYP1").Value)
            lblMonths.Text = CStr(MOS) & " Mos"
        Else
            lblMonths.Text = ""
        End If

        Dim RYP0 As String = Absx1.cmbFor("RYP0").Value
        If MOS >= 1 And MOS <= 12 Then
            ReDim YPs(MOS)
            For i As Integer = 1 To MOS
                YPs(i) = ASCMAIN1.Period_Calc(RYP0, (i - 1))
            Next
        End If

    End Sub
#End Region

#Region "grdSATBUDGX"

    Private Sub grdSATBUDGX_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSATBUDGX.AfterExitEditMode
        With grdSATBUDGX
            Select Case .ActiveCell.Column.Key
                Case "COLLECTION_CODE"
                    Dim COLLECTION_CODE As String = .ActiveCell.Text
                    If COLLECTION_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(COLLECTION_CODE, .ActiveCell.Column.Key)
                    End If

                Case "CUST_CODE"
                    Dim CUST_CODE As String = .ActiveCell.Text
                    If CUST_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(CUST_CODE, .ActiveCell.Column.Key)
                    End If

                Case "ITEM_BASIC_PROMO"
                    Dim ITEM_BASIC_PROMO As String = .ActiveCell.Text
                    If ITEM_BASIC_PROMO <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(ITEM_BASIC_PROMO, .ActiveCell.Column.Key)
                    End If
            End Select
        End With
    End Sub

    Private Sub grdSATBUDGX_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSATBUDGX.AfterRowActivate
        With grdSATBUDGX.DisplayLayout.Bands(0)
            If grdSATBUDGX.ActiveRow.IsAddRow Then
                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("ITEM_BASIC_PROMO").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("COLLECTION_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("ITEM_BASIC_PROMO").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("COLLECTION_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdSATBUDGX_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSATBUDGX.BeforeRowUpdate
        With grdSATBUDGX

            If Not e.Cancel Then
                If e.Row.Cells("CUST_CODE").Text = "" And ScreenMode Then
                    .ActiveRow.Cells("CUST_CODE").Value = Absx1.CtlFor("CUST_CODE").Text
                End If
                If e.Row.Cells("COLLECTION_CODE").Text = "" And ScreenMode Then
                    .ActiveRow.Cells("COLLECTION_CODE").Value = cbeCOLLECTION_CODE.Value
                End If

            End If
        End With
    End Sub

    Private Sub grdSATBUDGX_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSATBUDGX.ClickCellButton
        Select Case grdSATBUDGX.ActiveCell.Column.Key
            Case "COLLECTION_CODE"
                grdClickCellButton(grdSATBUDGX)
            Case "CUST_CODE"
                grdClickCellButton(grdSATBUDGX)
        End Select
    End Sub
#End Region

    Private Sub optCOLLECTION_CODE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCOLLECTION_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_grd()
    End Sub

    Sub Setup_grd()

        cbeCOLLECTION_CODE.Enabled = (optCOLLECTION_CODE.Value = "I")
        'grdSATBUDGX.DisplayLayout.GroupByBox.Hidden = (optCOLLECTION_CODE.Value = "A")

        With grdSATBUDGX.DisplayLayout.Bands(0)
            .Columns("COLLECTION_CODE").Hidden = (optCOLLECTION_CODE.Value <> "A")
            .Columns("CUST_CODE").Hidden = (Absx1.txtFor("CUST_CODE").Text <> "")
            .Columns("SREP_CODE").Hidden = (Absx1.txtFor("SREP_CODE").Text <> "")

            .SortedColumns.Clear()
            If optCOLLECTION_CODE.Value = "A" Then
                'grdSATBUDGX.DisplayLayout.Bands(0).SortedColumns.Add("COLLECTION_CODE", False, True)
            End If
            .SortedColumns.Add("CUST_CODE", False)
            .SortedColumns.Add("COLLECTION_CODE", False)
            .SortedColumns.Add("ITEM_BASIC_PROMO", False)


        End With

        Dim COLLS As String = ""
        Dim allow_modifications As Boolean = True

        Dim DVW As DataView = DirectCast(grdSATBUDGX.DataSource, DataTable).DefaultView
        Dim sql As String = ""
        If optCOLLECTION_CODE.Value = "A" Then
            allow_modifications = False
            COLLS = "All Collections"
        Else
            sql = "and COLLECTION_CODE = '" & cbeCOLLECTION_CODE.Value & "'"
            COLLS = cbeCOLLECTION_CODE.Value
        End If

        Dim RYP_LEGENDS As String = Absx1.cmbFor("RYP0").Text & " thru " & Absx1.cmbFor("RYP1").Text

        DVW.RowFilter = Mid(sql, 5)
        grdSATBUDGX.Text = "Retail Sales Budgets, by Store/Month, for " & RYP_LEGENDS & " - " & COLLS
        If allow_modifications Then
            grdSATBUDGX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdSATBUDGX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            grdSATBUDGX.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
        Else
            grdSATBUDGX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdSATBUDGX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True

            'grdSATBUDGX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            'grdSATBUDGX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            grdSATBUDGX.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        End If
        'grdSATBUDGX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

    End Sub

    Private Sub cbeCOLLECTION_CODE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeCOLLECTION_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_grd()
    End Sub

    Sub Set_Month_Headings()

        Dim RYP0 As String = Absx1.cmbFor("RYP0").Value
        Dim RYP1 As String = Absx1.cmbFor("RYP1").Value

        For M As Integer = 1 To 12
            With grdSATBUDGX.DisplayLayout.Bands(0).Columns("BUDGET_P" & Format(M, "00"))
                Dim YP As String = ASCMAIN1.Period_Calc(RYP0, (M - 1))
                If YP > RYP1 Then
                    .Hidden = True
                Else
                    Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", YP)
                    Dim LEGEND As String = rowGLTPARM2.Item("LEGEND")
                    .Header.Caption = Mid(LEGEND, 10, 6)
                    .Width = 60
                    .Hidden = False
                End If
            End With
        Next
    End Sub

    Overrides Function Excel_Import_Pre_Process_SG _
    (ByVal grd As UltraWinGrid.UltraGrid, dt As DataTable,
     Optional ByRef load_by_table As Boolean = False, _
     Optional ByRef load_handled As Boolean = False, _
     Optional ByRef F As ASFEXCL1 = Nothing) As Int64

        Dim dtbad As DataTable = dt.Clone
        dtbad.Columns.Add("ERROR")

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Budgets from XLS")

        Dim RowsMax As Int64 = dt.Rows.Count
        Dim r As Int64 = 0

        load_handled = True
        If dt.Rows.Count = 0 Then
            MsgBox("No Rows Loaded", MsgBoxStyle.OkOnly, "Import Failed")
        Else
            dst.Tables("SATBUDGX").Rows.Clear()
        End If

        For Each row As DataRow In dt.Select("")
            r += 1
            If r Mod 100 = 0 Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading Budgets from XLS")
                RowsMax = dt.Rows.Count
                ASCMAIN1.Progress("-", CStr(r) & "/" & CStr(RowsMax))
            End If

            Try
                Dim rowSATBUDGX As DataRow = dst.Tables("SATBUDGX").NewRow
                With rowSATBUDGX
                    For Each C As String In New String() {"COLLECTION_CODE", "CUST_CODE", "ITEM_BASIC_PROMO"}
                        .Item(C) = row.Item(C)
                    Next
                    For I As Integer = 1 To MOS
                        Dim C As String = "BUDGET_P" & Format(I, "00")
                        .Item(C) = row.Item(C)
                    Next
                End With

                dst.Tables("SATBUDGX").Rows.Add(rowSATBUDGX)

            Catch ex As Exception
                Dim rowbad As DataRow = dtbad.NewRow
                rowbad.ItemArray = row.ItemArray
                rowbad.Item("ERROR") = ex.Message
                dtbad.Rows.Add(rowbad)
            End Try
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        If dtbad.Rows.Count > 0 Then
            Using fr As New ASFMSGBF
                fr.Show_grd(dtbad, Me, "Some Rows Failed to Update - Please Check Last Column for Messages")
            End Using
        End If

    End Function

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

        If optCalendar.Value = "O" Then
            P0 = ASCMAIN1.Period_Calc(P0, -1)
            P1 = ASCMAIN1.Period_Calc(P1, -1)
        End If

        ASCMAIN1.sql = "Select OPS_YYYYPP from GLTPARM2 where OPS_YYYYMM = '" & P0 & "'"
        Absx1.cmbFor("RYP0").Value = ASCDATA1.GetDataValue
        ASCMAIN1.sql = "Select OPS_YYYYPP from GLTPARM2 where OPS_YYYYMM = '" & P1 & "'"
        Absx1.cmbFor("RYP1").Value = ASCDATA1.GetDataValue
    End Sub

    Private Sub optCalendar_ValueChanged_1(sender As Object, e As EventArgs) Handles optCalendar.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Dim YP As String

        YP = Absx1.cmbFor("RYP0").Value
        If optCalendar.Value = "R" Then
            YP = ASCMAIN1.Period_Calc(YP, 1)
        Else
            YP = ASCMAIN1.Period_Calc(YP, -1)
        End If
        Absx1.cmbFor("RYP0").Value = YP

        YP = Absx1.cmbFor("RYP1").Value
        If optCalendar.Value = "R" Then
            YP = ASCMAIN1.Period_Calc(YP, 1)
        Else
            YP = ASCMAIN1.Period_Calc(YP, -1)
        End If
        Absx1.cmbFor("RYP1").Value = YP

    End Sub

    Sub Create_Work_Tables()

        Dim RYP0 As String = Absx1.cmbFor("RYP0").Value
        Dim RYP1 As String = Absx1.cmbFor("RYP1").Value
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Dim BRAND_CODE As String = Absx1.txtFor("BRAND_CODE").Text
        Dim SREP_CODE As String = Absx1.txtFor("SREP_CODE").Text

        Dim sqlM As String = ""
        For I As Integer = 1 To 12
            Dim YP As String = ASCMAIN1.Period_Calc(RYP0, I - 1)
            If YP > RYP1 Then YP = ""
            sqlM &= ", Sum (Decode(SATBUDG1.OPS_YYYYPP,'" & YP & "',BUDGET,0)) BUDGET_P" & Format(I, "00")
        Next

        ASCMAIN1.sql = "Select SATBUDG1.COLLECTION_CODE, SATBUDG1.CUST_CODE, SATBUDG1.ITEM_BASIC_PROMO" & vbCrLf _
            & sqlM _
            & " from SATBUDG1,ARTCUST1" _
            & " where SATBUDG1.OPS_YYYYPP between '" & RYP0 & "' and '" & RYP1 & "'" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SATBUDG1.CUST_CODE" & vbCrLf _
            & IIf(CUST_CODE = "", "", "   and SATBUDG1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf) _
            & IIf(SREP_CODE = "", "", "   and ARTCUST1.SREP_CODE = '" & SREP_CODE & "'" & vbCrLf) _
            & IIf(BRAND_CODE = "", "", "" _
                  & " and SATBUDG1.COLLECTION_CODE in " _
                  & " (Select COLLECTION_CODE FROM ICTCOLL1 where BRAND_CODE = '" & BRAND_CODE & "')" & vbCrLf) _
            & " group by SATBUDG1.COLLECTION_CODE, SATBUDG1.CUST_CODE, SATBUDG1.ITEM_BASIC_PROMO"

        If SATBUDGX = "" Then
            SATBUDGX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SATBUDGX & " Add Primary Key (COLLECTION_CODE, CUST_CODE, ITEM_BASIC_PROMO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SATBUDGX)
            ASCDATA1.ExecuteSQL("Insert into " & SATBUDGX & " " & ASCMAIN1.sql)
        End If
    End Sub
End Class