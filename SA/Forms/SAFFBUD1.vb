Public Class SAFFBUD1
    Dim rowSATFBUD1 As DataRow
    Dim BUDGET_YEAR As String
    Dim BUDGET_YEAR_LY As String
    Dim BUDGET_VERSION As String
    Dim BUDGET_ACT_THRU As Integer
    Dim BUDGET_FS As String = ""
    Dim BUDGET_ACT_THRU_TY As Integer = 0
    Dim AYP As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select SATFBUD1.*, SATFBUD1.BUDGET_VERSION BUDGET_VERSION_DESC" & vbCrLf _
                & " from SATFBUD1 "
            Create_TDA(.Tables.Add, "SATFBUDX", "**", 0, False, "", 2)

            Create_TDA(.Tables.Add, "SATFBUD1", "*")

            ASCMAIN1.sql = "Select SATFBUD2.* from SATFBUD2" & vbCrLf _
                & " where BUDGET_YEAR = :PARM1 and BUDGET_VERSION = :PARM2"
            Create_TDA(.Tables.Add, "SATFBUD2", "**", 0, True, "VV", 6)

            Dim t As String = ""
            For M As Integer = 1 To 12
                t &= "+ISNULL(BUDGET_P" & Format(M, "00") & ",0)"
            Next
            .Tables("SATFBUD2").Columns.Add("TOTAL", GetType(System.Decimal), t)

            ASCMAIN1.sql = "Select SATFBUD2.* from SATFBUD2"
            Create_TDA(.Tables.Add, "SATFBUD0", "**", 0, False, "", 5)
            .Tables("SATFBUD0").Columns.Add("TOTAL", GetType(System.Decimal), t)

            Create_Relation("SATFBUD0", "SATFBUD2", "BUDGET_YEAR,BUDGET_VERSION,BUDGET_FS,BUDGET_DATA_CODE,CHANNEL_CODE")
            For M As Integer = 0 To 12
                Dim C As String = IIf(M = 0, "TOTAL", "BUDGET_P" & Format(M, "00"))
                dst.Tables("SATFBUD0").Columns(C).Expression = "SUM(CHILD." & C & ")"
            Next

            Create_TDA(.Tables.Add, "ICTCOLL1", "*", 0, False)
            Create_TDA(.Tables.Add, "SATFBUDC", "*", 0, False)
            Create_TDA(.Tables.Add, "SOTCHAN1", "*", 0, False)
        End With

        Fill_Records("ICTCOLL1")
        Fill_Records("SATFBUDC")
        Fill_Records("SOTCHAN1")

        grdSATFBUDX.DataSource = dst.Tables("SATFBUDX")
        grdSATFBUD2.DataSource = dst.Tables("SATFBUD2")
        grdSATFBUD0.DataSource = dst.Tables("SATFBUD0")

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSATFBUD0, grdSATFBUD2}
            Create_Summary(grd, "BUDGET_DATA_CODE", "Count")
            Create_Summary(grd, "LINE_DATA_DESC", "Count")
            For M As Integer = 1 To 12
                Create_Summary(grd, "BUDGET_P" & Format(M, "00"), , , "#,##0")
            Next
            Create_Summary(grd, "TOTAL", , , "#,##0")

            With grd.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() _
                    {"LINE_DATA_DESC", "BUDGET_DATA_CODE", "CHANNEL_CODE", "COLLECTION_CODE"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                Next
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    gcol.Header.Appearance.BackColor = System.Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                    If New String() {"BUDGET_FS", "LINE_DATA_DESC", "BUDGET_DATA_CODE", "CHANNEL_CODE", "COLLECTION_CODE"}.Contains(gcol.Key) Then
                        gcol.CellAppearance.BackColor = System.Drawing.Color.White
                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                        gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    ElseIf gcol.Key = "BUDGET_YEAR" Or gcol.Key = "BUDGET_VERSION" Or gcol.Key = "BUDGET_FS" Then
                        gcol.Hidden = True
                    ElseIf gcol.Key = "TOTAL" Then
                        gcol.Format = "#,##0"
                        gcol.Width = 100
                    Else
                        If gcol.Key.StartsWith("BUDGET_P") Then
                            gcol.Format = "#,##0"
                            gcol.Width = 90
                            gcol.Header.Caption = Format(CDate(Mid(gcol.Key, 9, 2) & "/01/2020"), "MMM") ' & "'" & "20"
                        End If
                    End If
                Next
            End With

            Show_Filter(grd, True)
        Next


        Dim YEARs As New List(Of String)
        For Y As Integer = Val(Now.Year) - 2 To Val(Now.Year + 1)
            YEARs.Add(Format(Y, "0000"))
        Next
        Absx1.cbeFor("BUDGET_YEAR").DataSource = YEARs ' New String() {"2008", "2009", "2010"}
        Absx1.cbeFor("BUDGET_YEAR").Value = Now.Year

        Dim VL As New ValueList
        Dim MONTHs As New List(Of String)
        Dim vl_MONTHs As New List(Of String)
        vl_MONTHs.Add(":")
        For M As Integer = 0 To 12
            Dim VLI_DESC As String
            If M = 0 Then
                VLI_DESC = "None"
            Else
                Dim d As Date = CDate(Format(M, "00") & "/01/2020")
                VLI_DESC = Format(d, "MMM")
            End If
            MONTHs.Add(VLI_DESC)
            Dim VLI As New ValueListItem(M, VLI_DESC)
            VL.ValueListItems.Add(VLI)
            vl_MONTHs.Add(CStr(M) & ":" & VLI_DESC)
        Next
        Absx1.cbeFor("BUDGET_ACT_THRU").DataSource = MONTHs
        Absx1.cbeFor("BUDGET_ACT_THRU").SelectedIndex = 0
        grdSATFBUDX.DisplayLayout.Bands(0).Columns("BUDGET_ACT_THRU").ValueList = VL

        VL.ValueListItems.Clear()
        For I As Integer = 0 To optBUDGET_VERSION.Items.Count - 1
            Dim VLI As New ValueListItem(CStr(I), optBUDGET_VERSION.Items(I).DisplayText)
            VL.ValueListItems.Add(VLI)
        Next
        grdSATFBUDX.DisplayLayout.Bands(0).Columns("BUDGET_VERSION_DESC").ValueList = VL

        ASCMAIN1.Add_Value_List(grdSATFBUD0, "BUDGET_DATA_CODE", "Select BUDGET_DATA_CODE,BUDGET_DATA_DESC from SATFBUDC")
        ASCMAIN1.Add_Value_List(grdSATFBUD2, "BUDGET_DATA_CODE", "Select BUDGET_DATA_CODE,BUDGET_DATA_DESC from SATFBUDC")

        ASCMAIN1.Add_Value_List(grdSATFBUD0, "BUDGET_FS", Nothing, New String() {":", "F:Financial", "S:Sales"})

        ASCMAIN1.Add_Value_List(grdSATFBUD2, "BUDGET_FS", Nothing, New String() {":", "F:Financial", "S:Sales"})
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load Budget"

                BUDGET_YEAR = Absx1.cbeFor("BUDGET_YEAR").Value
                BUDGET_VERSION = Absx1.optFor("BUDGET_VERSION").Value
                BUDGET_FS = optFS.Value
                BUDGET_ACT_THRU = 0

                If BUDGET_VERSION = "0" Then
                    rowSATFBUD1 = LookUp("SATFBUD1", New String() {BUDGET_YEAR, BUDGET_VERSION})
                Else
                    Dim P As Integer = Val(BUDGET_VERSION) - 1
                    Dim rowPrior As DataRow = LookUp("SATFBUD1", New String() {BUDGET_YEAR, CStr(P)})
                    If rowPrior Is Nothing Then
                        EMsg &= vbCr & "Cannot Load " & Absx1.optFor("BUDGET_VERSION").Text & " because " & optBUDGET_VERSION.Items(P).DisplayText & " does not exist yet"
                    Else
                        BUDGET_ACT_THRU = cbeBUDGET_ACT_THRU.SelectedItem.ListIndex

                        If Val(rowPrior.Item("BUDGET_ACT_THRU") & "") > BUDGET_ACT_THRU Then
                            EMsg &= vbCr & "Cannot Load " & Absx1.optFor("BUDGET_VERSION").Text & " with Actuals prior to " & optBUDGET_VERSION.Items(P).DisplayText
                        ElseIf Val(rowPrior.Item("BUDGET_ACT_THRU") & "") = BUDGET_ACT_THRU Then
                            EMsg &= vbCr & "Cannot Load " & Absx1.optFor("BUDGET_VERSION").Text & " with Actuals same as " & optBUDGET_VERSION.Items(P).DisplayText
                        Else
                            rowSATFBUD1 = LookUp("SATFBUD1", New String() {BUDGET_YEAR, BUDGET_VERSION})

                            If rowSATFBUD1 Is Nothing Then
                                If MsgBox("You are about to create a new Version (" & Absx1.optFor("BUDGET_VERSION").Text & ") for Budget Year " & BUDGET_YEAR & "." _
                                          & vbCrLf & vbCrLf & "Please verify that the value provided for Actuals Thru (" & cbeBUDGET_ACT_THRU.Value & ") is correct." _
                                          & vbCrLf & "Actuals Thru is not editable once a new Version has been established." _
                                          & vbCrLf & vbCrLf & "OK to continue with this Version?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                    Exit Sub
                                End If
                            Else
                                With rowSATFBUD1
                                    Dim BUDGET_OPER As String = .Item(IIf(optFS.Value = "F", "BUDGET_FIN_OPER", "BUDGET_SLS_OPER")) & ""
                                    Dim BUDGET_DATE As String = .Item(IIf(optFS.Value = "F", "BUDGET_FIN_DATE", "BUDGET_SLS_DATE")) & ""
                                    If BUDGET_OPER <> "" Then
                                        If MsgBox("You are about to re-load " & optFS.Text & " Budgets for " & BUDGET_YEAR & " " & Absx1.optFor("BUDGET_VERSION").Text & "." _
                                                  & vbCrLf & vbCrLf & "These Budgets were last loaded " & BUDGET_DATE & " by " & BUDGET_OPER & "." _
                                                  & vbCrLf & vbCrLf & "OK to continue?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                End With

                            End If


                        End If

                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("SATFBUD1", BUDGET_YEAR) Then
                        Exit Sub
                    End If
                End If

            Case "Create Report"

                If grdSATFBUDX.ActiveRow Is Nothing Then
                    EMsg &= vbCr & "No Budget Row Selected to Create Report for"
                End If

                If EMsg = "" Then
                    BUDGET_YEAR = grdSATFBUDX.ActiveRow.Cells("BUDGET_YEAR").Value
                    BUDGET_VERSION = grdSATFBUDX.ActiveRow.Cells("BUDGET_VERSION").Value
                    BUDGET_ACT_THRU = Val(grdSATFBUDX.ActiveRow.Cells("BUDGET_ACT_THRU").Value & "")

                    optBUDGET_VERSION.Value = BUDGET_VERSION
                    cbeBUDGET_YEAR.Value = BUDGET_YEAR
                End If

            Case "Delete"

                If grdSATFBUDX.ActiveRow Is Nothing OrElse Not grdSATFBUDX.ActiveRow.IsDataRow OrElse grdSATFBUDX.ActiveRow.IsFilterRow OrElse grdSATFBUDX.Selected.Rows.Count > 1 Then
                    EMsg &= vbCr & "Cannot Delete without Selecting a single Budget Record in the Grid"
                End If

                If EMsg = "" Then

                    BUDGET_YEAR = grdSATFBUDX.ActiveRow.Cells("BUDGET_YEAR").Value
                    BUDGET_VERSION = grdSATFBUDX.ActiveRow.Cells("BUDGET_VERSION").Value

                    If Not ASCMAIN1.Logical_Lock("SATFBUD1", BUDGET_YEAR) Then
                        Exit Sub
                    End If

                    optBUDGET_VERSION.Value = BUDGET_VERSION
                    cbeBUDGET_YEAR.Value = BUDGET_YEAR

                    If MsgBox("OK to Delete all Budget Records for " & BUDGET_YEAR & "-" & optBUDGET_VERSION.Text & "?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    End If
                End If


            Case "Import from XLS"
                Using openFileDialog1 As New OpenFileDialog
                    openFileDialog1.InitialDirectory = "c:\"
                    openFileDialog1.Title = "Select a File to Import Budget Data from"
                    openFileDialog1.Filter = "XLSx files (*.XLSx)|*.XLSx"
                    openFileDialog1.FilterIndex = 2
                    openFileDialog1.RestoreDirectory = True

                    If openFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        If Not Load_from_XLS(openFileDialog1.FileName) Then
                            Me.Cursor = Cursors.Default
                            ASCMAIN1.Progress("")
                            Exit Sub
                        Else
                            Me.Cursor = Cursors.Default
                            ASCMAIN1.Progress("")
                        End If
                    Else
                        Exit Sub
                    End If
                End Using

            Case "Update"
                For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SATFBUD2"), New String() {"COLLECTION_CODE"}).Select("")
                    Dim COLLECTION_CODE As String = row.Item(0)
                    If dst.Tables("ICTCOLL1").Rows.Find(New String() {COLLECTION_CODE}) Is Nothing Then
                        EMsg &= vbCr & "Invalid Collection (" & COLLECTION_CODE & ")"
                    End If
                Next

            Case "Delete"
                If MsgBox("Are you sure that you want to Delete all Budget Records for " & BUDGET_YEAR & " " & optBUDGET_VERSION.Text & "?", MsgBoxStyle.Critical + MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load Budget"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Create Report"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)
                Click_Command("Done")

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load Budget").Settings.Enabled = not_iScreenMode
                    .Items("Create Report").Settings.Enabled = not_iScreenMode
                    .Items("Import from XLS").Settings.Enabled = iScreenMode
                    .Items("Import from XLS").Visible = (EntryMode = "E" AndAlso optFS.Value = "S")

                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Done").Visible = (EntryMode = "V")

                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Update").Visible = (EntryMode = "E")
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Cancel").Visible = (EntryMode = "E")

                    .Items("Delete").Settings.Enabled = not_iScreenMode
                    .Items("Delete").Visible = Not ScreenMode ' (EntryMode = "E")
                    '    .Items("Delete").Visible = False ' before enabling this - are you deleting ALL budget records, or only Finanacial or Sales?

                    ' .Items("Import from XLS").Visible = Not tf
                End With

                .Groups("Load Options").Visible = Not ScreenMode Or (EntryMode = "E")
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        splBUD.Visible = ScreenMode
        grdSATFBUDX.Visible = Not ScreenMode

        Set_Read_Only(grpOptions, ScreenMode)

        If ScreenMode Then

            If EntryMode = "V" Then
                grdSATFBUD0.DisplayLayout.Bands(0).Columns("BUDGET_FS").Hidden = False
                grdSATFBUD2.DisplayLayout.Bands(0).Columns("BUDGET_FS").Hidden = False
            Else
                grdSATFBUD0.DisplayLayout.Bands(0).Columns("BUDGET_FS").Hidden = True
                grdSATFBUD2.DisplayLayout.Bands(0).Columns("BUDGET_FS").Hidden = True
            End If

        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SATFBUD1", "SATFBUD2", "SATFBUD0"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Documents()
        Setup_cbeBUDGET_YEAR()

        BUDGET_YEAR = ""
        BUDGET_VERSION = ""
    End Sub

    Sub Refresh_Documents()
        Fill_Records("SATFBUDX")
        Sort_grdColumns(grdSATFBUDX, "BUDGET_YEAR, BUDGET_VERSION".ToLower)
    End Sub
    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)
        BUDGET_YEAR_LY = Format(Val(BUDGET_YEAR) - 1, "0000")
        AYP = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1) ' last period with completed actuals
        If Mid(AYP, 1, 4) >= BUDGET_YEAR Then
            BUDGET_ACT_THRU_TY = Val(Mid(AYP, 5, 2))
        Else
            BUDGET_ACT_THRU_TY = 0
        End If

        EnforceConstraints(False)

        If EntryMode = "E" Then
            If rowSATFBUD1 Is Nothing Then
                rowSATFBUD1 = dst.Tables("SATFBUD1").NewRow
                With rowSATFBUD1
                    .Item("BUDGET_YEAR") = BUDGET_YEAR
                    .Item("BUDGET_VERSION") = BUDGET_VERSION
                    .Item("BUDGET_ACT_THRU") = BUDGET_ACT_THRU
                End With
                dst.Tables("SATFBUD1").Rows.Add(rowSATFBUD1)
            Else
                rowSATFBUD1 = Fill_Record("SATFBUD1", New String() {BUDGET_YEAR, BUDGET_VERSION})
            End If

            If BUDGET_FS = "F" Then
                Load_Financial_Budgets()
            Else
                'Load_Sales_Admin_Budgets()
            End If

        Else
            rowSATFBUD1 = Fill_Record("SATFBUD1", New String() {BUDGET_YEAR, BUDGET_VERSION})
            Fill_Records("SATFBUD2", New String() {BUDGET_YEAR, BUDGET_VERSION})
        End If

        dst.Tables("SATFBUD0").Rows.Clear()
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SATFBUD2"), New String() {"BUDGET_YEAR", "BUDGET_VERSION", "BUDGET_FS", "BUDGET_DATA_CODE", "CHANNEL_CODE"}).Select()
            'dst.Tables("SATFBUD0").Rows.Add(row.ItemArray)
            dst.Tables("SATFBUD0").Rows.Add(New String() { _
                                            row.Item("BUDGET_YEAR"), _
                                            row.Item("BUDGET_VERSION"), _
                                            row.Item("BUDGET_FS"), _
                                            row.Item("BUDGET_DATA_CODE"), _
                                            row.Item("CHANNEL_CODE"), "*"})
        Next

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSATFBUD0, grdSATFBUD2}
            For M As Integer = 1 To 12
                Dim C As String = "BUDGET_P" & Format(M, "00")
                If M <= BUDGET_ACT_THRU Then
                    grd.DisplayLayout.Bands(0).Columns(C).CellAppearance.BackColor = Drawing.Color.LightGray
                Else
                    grd.DisplayLayout.Bands(0).Columns(C).CellAppearance.BackColor = Drawing.Color.Empty
                End If
            Next
        Next

        EnforceConstraints(True)

        If EntryMode = "V" Then
            grdSATFBUD0.Text = "Budget - " & BUDGET_YEAR & " " & optBUDGET_VERSION.Text
            Create_Report()
        Else
            grdSATFBUD0.Text = optFS.Text & " Budget - " & BUDGET_YEAR & " " & optBUDGET_VERSION.Text
        End If

        Sort_grdColumns(grdSATFBUD0, "BUDGET_DATA_CODE, CHANNEL_CODE, COLLECTION_CODE")

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        Dim sql_Delete As String = "BUDGET_YEAR = '" & BUDGET_YEAR & "' and BUDGET_VERSION = '" & BUDGET_VERSION & "'"

        If BUDGET_FS = "F" Then
            rowSATFBUD1.Item("BUDGET_FIN_OPER") = ASCMAIN1.USER_ID
            rowSATFBUD1.Item("BUDGET_FIN_DATE") = DATETIME_STAMP
        Else
            rowSATFBUD1.Item("BUDGET_SLS_OPER") = ASCMAIN1.USER_ID
            rowSATFBUD1.Item("BUDGET_SLS_DATE") = DATETIME_STAMP
        End If

        Update_Record_TDA("SATFBUD1", sql_Delete)
        Update_Record_TDA("SATFBUD2", sql_Delete & " and BUDGET_FS = '" & BUDGET_FS & "'")

        CommitTrans("Update Complete")
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Dim sql_Delete As String = "BUDGET_YEAR = '" & BUDGET_YEAR & "' and BUDGET_VERSION = '" & BUDGET_VERSION & "'"
        ASCMAIN1.sql = "Delete from SATFBUD1 where " & sql_Delete
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Delete from SATFBUD2 where " & sql_Delete
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.Record_Event("SATFBUD1", BUDGET_YEAR, BUDGET_VERSION, DATETIME_STAMP, ASCMAIN1.USER_ID, "DEL", "Delete Financail Budget", "")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        Dim sql_Delete As String = "BUDGET_YEAR = '" & BUDGET_YEAR & "' and BUDGET_VERSION = '" & BUDGET_VERSION & "'"

        ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
            & " where " & sql_Delete)
    End Sub

#End Region

#Region "Popup_Menus"
    Overrides Sub Load_Popup_Menus()
        '  Load_Popup_Menu(grdSATFBUD2, "BBBB", "Clear Column", "Copy Value", "Export Spring", "Export Fall")
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
                'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select
    End Sub

    Public Overrides Sub tlb_ToolValueChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolEventArgs) Handles tlb.ToolValueChanged

    End Sub

#End Region

#Region "ABSColumn Controls"
    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)

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
            'Case "CUST_CODE"
            '    If EntryMode = "" Then
            '        If Absx1.txtFor("CUST_CODE").Text <> "" Then
            '            LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
            '            If cdr IsNot Nothing Then

            '            End If
            '        End If
            '    End If

        End Select
    End Sub
#End Region

    Sub Create_Report()

        ASCMAIN1.Progress("Now Creating Workbook")

        Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & ".xlsX"

        Dim excel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlSourceRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim xlDestRange As Microsoft.Office.Interop.Excel.Range = Nothing

        excel = New Microsoft.Office.Interop.Excel.Application
        wb = excel.Workbooks.Open(FILENAME)

        ws = wb.Worksheets.Add()
        ws.Name = "Parameters"

        ws.Range(Excel_Cell(1, 1)).EntireColumn.ColumnWidth = 30
        ws.Range(Excel_Cell(1, 2)).EntireColumn.ColumnWidth = 25

        ws.Cells(1, 1).value2 = "Report Creation"
        ws.Range(Excel_Cell(1, 2)).NumberFormat = "@"
        ws.Cells(1, 2).value2 = Now.ToString
        ws.Cells(2, 1).value2 = "Budget Year"
        ws.Range(Excel_Cell(2, 2)).NumberFormat = "@"
        ws.Cells(2, 2).value2 = BUDGET_YEAR
        ws.Cells(3, 1).value2 = "Budget Version"
        ws.Range(Excel_Cell(3, 2)).NumberFormat = "@"
        ws.Cells(3, 2).value2 = BUDGET_VERSION


        Dim BUDGET_ACT_THRU_DESC As String = ""
        If BUDGET_ACT_THRU <> 0 Then
            BUDGET_ACT_THRU_DESC = Format(CDate(Format(BUDGET_ACT_THRU, "00") & "/01/2020"), "MMM")
        End If
        ws.Cells(4, 1).value2 = "Budget lines updated thru " & BUDGET_ACT_THRU_DESC
        ws.Range(Excel_Cell(4, 2)).NumberFormat = "@"
        ws.Cells(4, 2).value2 = BUDGET_ACT_THRU
        ws.Cells(4, 3).value2 = BUDGET_ACT_THRU_DESC

        ' I think we need to have 2 notes…one saying “budget lines updaed thru XXX” and the other saying “actual lines updated thru XXX”

        BUDGET_ACT_THRU_DESC = ""
        If BUDGET_ACT_THRU_TY <> 0 Then
            BUDGET_ACT_THRU_DESC = Format(CDate(Format(BUDGET_ACT_THRU_TY, "00") & "/01/2020"), "MMM")
        End If
        ws.Cells(5, 1).value2 = "TY Actual lines updated thru " & BUDGET_ACT_THRU_DESC
        ws.Range(Excel_Cell(5, 2)).NumberFormat = "@"
        ws.Cells(5, 2).value2 = BUDGET_ACT_THRU_TY
        ws.Cells(5, 3).value2 = BUDGET_ACT_THRU_DESC

        Dim SheetName As String = "Data"
        ws = wb.Worksheets(SheetName)

        BUDGET_FS = "F"
        Dim sql_LY_Actuals As String = Load_Financial_Budgets(True)

        ASCMAIN1.sql = "" _
            & "Select SATFBUD2.*, SATFBUDC.BUDGET_DATA_DESC, 'B' BUD_ACT, SATFBUD2.BUDGET_YEAR YEAR, SATFBUD2.BUDGET_VERSION VERSION" & vbCrLf _
            & " from SATFBUD2, SATFBUDC" & vbCrLf _
            & " where SATFBUD2.BUDGET_YEAR = '" & BUDGET_YEAR & "' and SATFBUD2.BUDGET_VERSION <= '" & BUDGET_VERSION & "'" & vbCrLf _
            & "   and SATFBUDC.BUDGET_DATA_CODE = SATFBUD2.BUDGET_DATA_CODE"
        If sql_LY_Actuals <> "" Then
            Dim LY As String = Format(Val(BUDGET_YEAR) - 1, "0000")
            For y As Integer = 0 To 5
                Dim YYYY As String = Format(Val(BUDGET_YEAR) - y, "0000")
                Dim SQLA As String = Replace(sql_LY_Actuals, "G.ACCT_YEAR = '" & LY & "'", "G.ACCT_YEAR = '" & YYYY & "'")
                ASCMAIN1.sql &= vbCrLf _
                    & " union " & vbCrLf _
                    & "Select X.*, SATFBUDC.BUDGET_DATA_DESC, 'A' BUD_ACT, '" & YYYY & "' YEAR, 'A' VERSION" & vbCrLf _
                    & " from (" & SQLA & ") X, SATFBUDC" & vbCrLf _
                    & " where SATFBUDC.BUDGET_DATA_CODE = X.BUDGET_DATA_CODE"
            Next
        End If

        Dim sql_Retail_Actuals_12 As String = ""
        For i As Integer = 1 To 12
            Dim iz As String = Format(i, "00")
            sql_Retail_Actuals_12 &= ", Sum (CASE WHEN SUBSTR(RSTRETL5.OPS_YYYYPP,5,2) = '" & iz & "' THEN RSTRETL5.AMT_SOLD ELSE 0 END) BUDGET_P" & iz & "" & vbCrLf
        Next
        ASCMAIN1.sql &= vbCrLf _
            & " union " & vbCrLf _
            & "Select '" & BUDGET_YEAR & "' BUDGET_YEAR, '" & BUDGET_VERSION & "' BUDGET_VERSION, 'S' BUDGET_FS, 'S' BUDGET_DATA_CODE" & vbCrLf _
            & ", SOTTCLS1.CHANNEL_CODE, RSTRETL5.COLLECTION_CODE" & vbCrLf _
            & sql_Retail_Actuals_12 _
            & ", SATFBUDC.BUDGET_DATA_DESC, 'A' BUD_ACT, SUBSTR(RSTRETL5.OPS_YYYYPP,1,4) YEAR, 'A' VERSION" & vbCrLf _
            & " from RSTRETL5,ARTCUST1,SOTTCLS1,SATFBUDC" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = RSTRETL5.CUST_CODE" & vbCrLf _
            & "   and RSTRETL5.OPS_YYYYPP BETWEEN '" & Format(Val(BUDGET_YEAR) - 5, "0000") & "01' AND '" & AYP & "'" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and SOTTCLS1.CHANNEL_CODE in ('1')" & vbCrLf _
            & "   and SATFBUDC.BUDGET_DATA_CODE = 'S'" & vbCrLf _
            & " group by SOTTCLS1.CHANNEL_CODE, RSTRETL5.COLLECTION_CODE, SUBSTR(RSTRETL5.OPS_YYYYPP,1,4), SATFBUDC.BUDGET_DATA_DESC"

        '& "   and SOTTCLS1.CHANNEL_CODE in ('1','2','4')" & vbCrLf _

        ASCMAIN1.sql = "Select X.*" & vbCrLf _
            & ", ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & " from (" & ASCMAIN1.sql & ") X, ICTCOLL1" & vbCrLf _
            & " where ICTCOLL1.COLLECTION_CODE (+) = X.COLLECTION_CODE"

        'Dim DataTable As DataTable = dst.Tables("SATFBUD2")
        Dim DataTable As DataTable = ASCDATA1.GetDataTable

        'For Each row As DataRow In DataTable.Select("CHANNEL = '2' D R")

        'Next

        Dim iRx As Integer = 1
        Dim r As Integer = 0 ' since we are using XLS Automation
        Dim c As Integer


        'Dim dc As System.Data.DataColumn
        'Dim colIndex As Integer = 0
        'Dim rowIndex As Integer = 0
        ''Nombre de mesures
        'Dim maxrows As Integer = DataTable.Rows.Count

        ''Ecriture des entêtes de colonne et des mesures
        ''(Write column headers and data)

        'For Each dc In DataTable.Columns
        '    colIndex = colIndex + 1
        '    'Entête de colonnes (column headers)
        '    ws.Cells(1, colIndex) = dc.ColumnName
        '    'Données(data)
        '    'You can use CDbl instead of Cobj If your data is of type Double
        '    ws.Cells(2, colIndex).Resize(maxrows, ).Value = excel.Application.transpose(DataTable.Rows.OfType(Of DataRow)().[Select](Function(k) CObj(k(dc.ColumnName))).ToArray())
        'Next
        ''This worked for me after I set the default values in my data table to be something other than <DBNull>. I got a Type Mismatch otherwise because Excel cells can't handle the null values
        ''https://stackoverflow.com/questions/18388592/fast-export-of-large-datatable-to-excel-spreadsheet-in-vb-net

        Dim MaxCols As Integer = DataTable.Columns.Count

        Dim CHs As New Dictionary(Of String, String)
        For Each gcol As UltraWinGrid.UltraGridColumn In grdSATFBUD0.DisplayLayout.Bands(0).Columns
            CHs.Add(gcol.Key, gcol.Header.Caption)
        Next

        For i As Integer = 0 To MaxCols - 1
            Dim W As Integer = 15
            Dim CH As String = DataTable.Columns(i).ColumnName
            If CHs.ContainsKey(CH) Then
                W = grdSATFBUD0.DisplayLayout.Bands(0).Columns(CH).Width / 8
                CH = CHs(CH)
            End If

            If CH = "BUDGET_YEAR" Then ws.Range(Excel_Cell(1, i + 1)).EntireColumn.ColumnWidth = 0

            If CH = "BUDGET_YEAR" Then CH = "Report Year" : W = 10
            If CH = "BUDGET_VERSION" Then CH = "Report Version" : W = 10
            If CH = "BUDGET_DATA_DESC" Then CH = "Description"
            'If CH = "BUD_ACT_LYR" Then CH = "Bud/Act-Year" : W = 10
            If CH = "BUD_ACT" Then CH = "Bud/Act" : W = 10
            If CH = "YEAR" Then CH = "Year" : W = 10
            If CH = "VERSION" Then CH = "Version" : W = 10
            If CH = "COLLECTION_GENDER" Then CH = "Gender" : W = 10
            If CH = "HC_CODE" Then CH = "HC" : W = 10
            If CH = "BRAND_CODE" Then CH = "Brand" : W = 10

            ws.Cells(iRx + r, i + 1).Value2 = CH

            If DataTable.Columns(i).DataType.ToString = "System.String" Then
                ws.Range(Excel_Cell(iRx + r, i + 1)).EntireColumn.NumberFormat = "@"
            Else
                ws.Range(Excel_Cell(iRx + r, i + 1)).EntireColumn.NumberFormat = "#,##0.00"
            End If
            ws.Range(Excel_Cell(iRx + r, i + 1)).EntireColumn.ColumnWidth = W

        Next

        For Each row As DataRow In DataTable.Select("", "BUDGET_DATA_CODE,COLLECTION_CODE")
            r += 1
            ASCMAIN1.Progress("-", r)
            c = 0
            c = 1

            ws.Range(ws.Cells(iRx + r, c), ws.Cells(iRx + r, c + MaxCols - 1)).Value2 = row.ItemArray
            'ws.Range(ws.Cells(iRx + r, c), ws.Cells(iRx + r, c + MaxCols)).Value = row.ItemArray
            ws.Cells(iRx + r, 1).Value2 = BUDGET_YEAR
            ws.Cells(iRx + r, 2).Value2 = BUDGET_VERSION
            'For i As Integer = 0 To MaxCols - 1
            '    c += 1 : ws.Cells(iRx + r, c).Value2 = row.Item(i)
            'Next
        Next

        ' ws.Range(Excel_Cell(1, 1)).EntireColumn.Hidden = True
        ' ws.Range(Excel_Cell(0, 1)).EntireColumn.Hidden = True
        ws.Range(Excel_Cell(1, 1)).EntireColumn.ColumnWidth = 0
        ws.Range(Excel_Cell(1, 2)).EntireColumn.ColumnWidth = 0
        ws.Range(Excel_Cell(1, 2)).EntireColumn.Delete()
        ws.Range(Excel_Cell(1, 1)).EntireColumn.Delete()

        ASCMAIN1.Progress("-", "Pivot")
        wb.Names.Add("PivotBase", "=" & SheetName & "!" & Excel_Cell(iRx, 1, 3) & ":" & Excel_Cell(iRx + DataTable.Rows.Count, MaxCols, 3))

        ASCMAIN1.Progress("Now Saving Workbook")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = BUDGET_YEAR & "_" & BUDGET_VERSION & "_Budgets"
                XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsX"

                Dim objOpt As Object = Nothing ' Missing.Value
                wb.SaveAs(ASCMAIN1.Folders("Work") & XLS_FILENAME _
                          , Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook)
                wb.Close(False, objOpt, objOpt)

                success = True

            Catch ex As Exception
                Stop
            End Try
        Loop

        excel.Quit()
        ws = Nothing
        wb = Nothing
        excel = Nothing
        xlSourceRange = Nothing
        xlDestRange = Nothing

        'ReleaseCOMObject(xlDestRange)
        'ReleaseCOMObject(xlSourceRange)
        'ReleaseCOMObject(ws)
        'ReleaseCOMObject(wb)
        'ReleaseCOMObject(excel)

        Add_Document_to_ASTSPRF1(ASCMAIN1.Folders("Work") & XLS_FILENAME)

        ASCMAIN1.Progress("")

        'Return ASCMAIN1.Folders("Work") & XLS_FILENAME

    End Sub

    Private Sub grdSATFBUDX_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdSATFBUDX.DoubleClickRow
        If e.Row.IsDataRow AndAlso Not e.Row.IsFilterRow Then
            Click_Command("Create Report")
        End If
    End Sub

    Function Load_Financial_Budgets(Optional sql_for_LY As Boolean = False) As String

        Dim sql_Load_Financial_Budgets As String = ""

        Dim P As Integer = BUDGET_ACT_THRU
        Dim P_TY As Integer = BUDGET_ACT_THRU_TY
        If sql_for_LY Then P = 12

        If Not sql_for_LY And P_TY > BUDGET_ACT_THRU Then P_TY = BUDGET_ACT_THRU ' WITHOUT THIS, BUD + ACT GETS ADDED TOGETHER FOR APR - SEE IM ISSUE 06/11

        Dim sqlB As String = ""
        For I As Integer = 1 To 12
            If I <= P Then
                sqlB &= ", 0"
            Else
                sqlB &= ", Sum (-1 * G.ACCT_BUD_P" & Format(I, "00") & ")"
            End If
            sqlB &= " BUDGET_P" & Format(I, "00") & vbCrLf
        Next

        'Dim sqlAB As String = ""
        'For I As Integer = 1 To 12
        '    If I <= P Then
        '        sqlAB &= ", Sum (-1 * G.ACCT_ACT_P" & Format(I, "00") & ")"
        '    Else
        '        sqlAB &= ", Sum (Case when G.ACCT_YEAR = '" & Mid(AYP, 1, 4) & "' Then 0 Else -1 * G.ACCT_ACT_P" & Format(I, "00") & " End)"
        '        'sqlAB &= ", 0"
        '    End If
        '    sqlAB &= " BUDGET_P" & Format(I, "00") & vbCrLf
        'Next

        Dim sqlA As String = ""
        For I As Integer = 1 To 12
            If I <= P_TY Then
                sqlA &= ", Sum (-1 * G.ACCT_ACT_P" & Format(I, "00") & ")"
            Else
                If sql_for_LY Then
                    sqlA &= ", Sum (Case when G.ACCT_YEAR = '" & BUDGET_YEAR_LY & "' Then -1 * G.ACCT_ACT_P" & Format(I, "00") & " Else 0 End)"

                Else
                    sqlA &= ", Sum (Case when G.ACCT_YEAR = '" & Mid(AYP, 1, 4) & "' Then 0 Else -1 * G.ACCT_ACT_P" & Format(I, "00") & " End)"
                    ' sqlA &= ", Sum (Case when G.ACCT_YEAR = '" & Mid(AYP, 1, 4) & "' Then 0 Else -1 * G.ACCT_ACT_P" & Format(I, "00") & " End)"
                    ' how does this ever work for actuals?
                End If
                'sqlA &= ", 0"
            End If
            sqlA &= " BUDGET_P" & Format(I, "00") & vbCrLf
        Next

        Dim sqlT As String = ""
        For I As Integer = 1 To 12
            sqlT &= ", Sum (BUDGET_P" & Format(I, "00") & ") BUDGET_P" & Format(I, "00") & vbCrLf
        Next

        dst.Tables("SATFBUD2").Rows.Clear()


        Dim ACCT_CODEs As New List(Of String)
        Dim sql_ACCT_CODEs As String = ""
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SATFBUDC"), New String() {"ACCT_CODE"}).Select("ACCT_CODE IS Not Null", "")
            Dim ACCT_CODE As String = row.Item("ACCT_CODE")
            ACCT_CODEs.Add(ACCT_CODE)
            sql_ACCT_CODEs &= ",'" & ACCT_CODE & "'"
        Next

        Dim sql_GLTACCT2 As String = "(Select ACCT_CODE,SEG2_CODE," & vbCrLf _
            & "CASE WHEN SEG3_CODE = 'SEC' AND (ACCT_CODE = '321000' OR ACCT_CODE = '331000') THEN 'DPT' ELSE SEG3_CODE END SEG3_CODE," & vbCrLf _
            & "SEG4_CODE,ACCT_YEAR,ACCT_BEG_BAL," & vbCrLf _
            & "ACCT_BUD_P01,ACCT_BUD_P02,ACCT_BUD_P03,ACCT_BUD_P04,ACCT_BUD_P05,ACCT_BUD_P06," & vbCrLf _
            & "ACCT_BUD_P07,ACCT_BUD_P08,ACCT_BUD_P09,ACCT_BUD_P10,ACCT_BUD_P11,ACCT_BUD_P12,ACCT_BUD_P13" & vbCrLf _
            & " from GLTACCT2)"

        Dim sql_GLTACCT3 As String = "(Select ACCT_CODE,SEG2_CODE," & vbCrLf _
            & "CASE WHEN SEG3_CODE = 'SEC' AND (ACCT_CODE = '321000' OR ACCT_CODE = '331000') THEN 'DPT' ELSE SEG3_CODE END SEG3_CODE," & vbCrLf _
            & "SEG4_CODE,ACCT_YEAR,ACCT_BEG_BAL," & vbCrLf _
            & "ACCT_ACT_P01,ACCT_ACT_P02,ACCT_ACT_P03,ACCT_ACT_P04,ACCT_ACT_P05,ACCT_ACT_P06," & vbCrLf _
            & "ACCT_ACT_P07,ACCT_ACT_P08,ACCT_ACT_P09,ACCT_ACT_P10,ACCT_ACT_P11,ACCT_ACT_P12,ACCT_ACT_P13" & vbCrLf _
            & " from GLTACCT3)"

        For Each row As DataRow In dst.Tables("SATFBUDC").Select("")
            Dim BUDGET_DATA_CODE As String = row.Item("BUDGET_DATA_CODE")
            Dim ACCT_CODE As String = row.Item("ACCT_CODE") & ""
            Dim CHANNEL_CODE As String = row.Item("CHANNEL_CODE") & ""
            ' If BUDGET_DATA_CODE = "G" Then Stop
            If BUDGET_DATA_CODE = "N" Or ACCT_CODE <> "" Then
                Dim sql_ACCT_CODE As String = " where G.ACCT_CODE = '" & ACCT_CODE & "'"
                If BUDGET_DATA_CODE = "N" Then
                    sql_ACCT_CODE = " where G.ACCT_CODE in (" & Mid(sql_ACCT_CODEs, 2) & ")"
                End If

                Dim sqlGL As String = ""
                For Each TYP As String In New String() {"BUD", "ACT"}
                    If TYP = "ACT" And Not sql_for_LY And BUDGET_ACT_THRU = 0 Then Exit For

                    Dim sqlGL_TYP As String = "Select '" & BUDGET_YEAR & "' BUDGET_YEAR, '" & BUDGET_VERSION & "' BUDGET_VERSION, '" & BUDGET_FS & "' BUDGET_FS, '" & BUDGET_DATA_CODE & "' BUDGET_DATA_CODE" & vbCrLf _
                        & ", SOTTCLS1.CHANNEL_CODE, ICTCOLL1.COLLECTION_CODE" & vbCrLf _
                        & IIf(TYP = "BUD", sqlB, sqlA) _
                        & " from " & IIf(TYP = "BUD", sql_GLTACCT2, sql_GLTACCT3) & " G,ICTCOLL1,SOTTCLS1" & vbCrLf _
                        & sql_ACCT_CODE & vbCrLf _
                        & "   and G.ACCT_YEAR = '" & IIf(sql_for_LY, BUDGET_YEAR_LY, BUDGET_YEAR) & "'" & vbCrLf _
                        & "   and SOTTCLS1.TRADE_CLASS_CODE = DECODE(G.SEG3_CODE,'000','DPT',G.SEG3_CODE)" & vbCrLf _
                        & "   and ICTCOLL1.COLLECTION_CODE = G.SEG4_CODE" & vbCrLf _
                        & "   and SOTTCLS1.CHANNEL_CODE in ('1','2','4')" & vbCrLf _
                        & IIf(CHANNEL_CODE = "", "", "   and SOTTCLS1.CHANNEL_CODE = '" & CHANNEL_CODE & "'" & vbCrLf) _
                        & " group by SOTTCLS1.CHANNEL_CODE, ICTCOLL1.COLLECTION_CODE"

                    '& " from " & IIf(TYP = "BUD", "GLTACCT2", "GLTACCT3") & " G,ICTCOLL1,SOTTCLS1" & vbCrLf _

                    If sqlGL <> "" Then sqlGL &= vbCrLf & " union " & vbCrLf
                    sqlGL &= sqlGL_TYP

                    If sql_for_LY And TYP = "BUD" Then
                        sqlGL = ""
                    End If
                Next

                If sql_for_LY Then
                    ASCMAIN1.sql = sqlGL
                    If sql_Load_Financial_Budgets <> "" Then
                        sql_Load_Financial_Budgets &= vbCrLf & " UNION " & vbCrLf
                    End If
                    sql_Load_Financial_Budgets &= sqlGL

                Else
                    'dst.Tables("SATFBUD2").Rows.Clear()

                    ASCMAIN1.sql = "Select BUDGET_YEAR, BUDGET_VERSION, BUDGET_FS, BUDGET_DATA_CODE, CHANNEL_CODE, COLLECTION_CODE" & vbCrLf _
                        & sqlT _
                        & " from (" & sqlGL & ") X" & vbCrLf _
                        & " group by BUDGET_YEAR, BUDGET_VERSION, BUDGET_FS, BUDGET_DATA_CODE, CHANNEL_CODE, COLLECTION_CODE"

                    Fill_Records("SATFBUD2", "", False, ASCMAIN1.sql)
                End If

            ElseIf BUDGET_DATA_CODE = "S" Then ' Financial Retail Sales
                If sql_for_LY Then
                Else
                    Dim sqlS As String = ""
                    For I As Integer = 1 To 12
                        sqlS &= ", Sum (Case When SUBSTR(RSTBUDF1.OPS_YYYYPP,5,2) = '" & Format(I, "00") & "' then RSTBUDF1.BUDGET else 0 END) BUDGET_P" & Format(I, "00") & vbCrLf
                    Next

                    Dim sqlGL As String = "Select '" & BUDGET_YEAR & "' BUDGET_YEAR, '" & BUDGET_VERSION & "' BUDGET_VERSION, '" & BUDGET_FS & "' BUDGET_FS, '" & BUDGET_DATA_CODE & "' BUDGET_DATA_CODE" & vbCrLf _
                            & ", '1' CHANNEL_CODE, RSTBUDF1.COLLECTION_CODE" & vbCrLf _
                            & sqlS _
                            & " from RSTBUDF1" & vbCrLf _
                            & " where RSTBUDF1.OPS_YYYYPP >= '" & BUDGET_YEAR & "01'" & vbCrLf _
                            & "   and RSTBUDF1.OPS_YYYYPP <= '" & BUDGET_YEAR & "12'" & vbCrLf _
                            & " group by RSTBUDF1.COLLECTION_CODE"

                    ASCMAIN1.sql = "Select BUDGET_YEAR, BUDGET_VERSION, BUDGET_FS, BUDGET_DATA_CODE, CHANNEL_CODE, COLLECTION_CODE" & vbCrLf _
                        & sqlT _
                        & " from (" & sqlGL & ") X" & vbCrLf _
                        & " group by BUDGET_YEAR, BUDGET_VERSION, BUDGET_FS, BUDGET_DATA_CODE, CHANNEL_CODE, COLLECTION_CODE"

                    Fill_Records("SATFBUD2", "", False, ASCMAIN1.sql)
                End If
            End If
        Next

        Return sql_Load_Financial_Budgets ' note - this sql does not include Financial Retail Budgets
    End Function

    Private Sub cbeBUDGET_YEAR_ValueChanged(sender As Object, e As EventArgs) Handles cbeBUDGET_YEAR.ValueChanged
        Setup_cbeBUDGET_YEAR()
    End Sub

    Sub Setup_cbeBUDGET_YEAR()
        For i As Integer = 1 To optBUDGET_VERSION.Items.Count - 1
            optBUDGET_VERSION.Items(i).DisplayText = "Estimate " & CStr(i)
            optBUDGET_VERSION.Items(i).Tag = DBNull.Value
        Next

        For Each row As DataRow In dst.Tables("SATFBUDX").Select("BUDGET_YEAR = '" & cbeBUDGET_YEAR.Value & "'")
            Dim BUDGET_VERSION As String = row.Item("BUDGET_VERSION")
            If BUDGET_VERSION <> "0" Then
                Dim BUDGET_ACT_THRU As Integer = Val(row.Item("BUDGET_ACT_THRU") & "")
                Dim i As Integer = Val(BUDGET_VERSION)
                optBUDGET_VERSION.Items(i).DisplayText = "Estimate " & CStr(i) & " (" & Format(CDate(Format(BUDGET_ACT_THRU, "00") & "/01/2020"), "MMM") & ")"
                optBUDGET_VERSION.Items(i).Tag = BUDGET_ACT_THRU
            End If
        Next
        If optBUDGET_VERSION.Value & "" = "" Then optBUDGET_VERSION.Value = "0"
        Setup_cbeBUDGET_ACT_THRU()
    End Sub
    Sub Setup_cbeBUDGET_ACT_THRU()
        Dim BUDGET_YEAR As String = cbeBUDGET_YEAR.Value
        Dim BUDGET_VERSION As String = optBUDGET_VERSION.Value

        Dim rowSATFBUDX As DataRow = dst.Tables("SATFBUDX").Rows.Find(New String() {BUDGET_YEAR, BUDGET_VERSION})

        lblBUDGET_ACT_THRU.Visible = Not (BUDGET_VERSION = "0")
        cbeBUDGET_ACT_THRU.Visible = Not (BUDGET_VERSION = "0")

        cbeBUDGET_ACT_THRU.ReadOnly = (rowSATFBUDX IsNot Nothing)
        If (rowSATFBUDX IsNot Nothing) Then
            cbeBUDGET_ACT_THRU.SelectedIndex = Val(rowSATFBUDX.Item("BUDGET_ACT_THRU") & "")
        Else
            cbeBUDGET_ACT_THRU.SelectedIndex = 0
        End If
    End Sub

    Private Sub optBUDGET_VERSION_ValueChanged(sender As Object, e As EventArgs) Handles optBUDGET_VERSION.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Setup_cbeBUDGET_ACT_THRU()
    End Sub

    Private Sub grdSATFBUD0_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSATFBUD0.AfterRowActivate
        If grdSATFBUD0.ActiveRow IsNot Nothing AndAlso grdSATFBUD0.ActiveRow.IsDataRow AndAlso Not grdSATFBUD0.ActiveRow.IsFilterRow Then
            Dim BUDGET_DATA_CODE As String = grdSATFBUD0.ActiveRow.Cells("BUDGET_DATA_CODE").Value
            Dim rowSATFBUDC As DataRow = dst.Tables("SATFBUDC").Rows.Find(BUDGET_DATA_CODE)
            Dim BUDGET_DATA_DESC As String = rowSATFBUDC.Item("BUDGET_DATA_DESC") & ""
            Dim CHANNEL_CODE As String = grdSATFBUD0.ActiveRow.Cells("CHANNEL_CODE").Value
            Dim dvw As DataView = DirectCast(grdSATFBUD2.DataSource, DataTable).DefaultView
            dvw.RowFilter = "BUDGET_DATA_CODE = '" & BUDGET_DATA_CODE & "' and CHANNEL_CODE = '" & CHANNEL_CODE & "'"
            Sort_grdColumns(grdSATFBUD2, "BUDGET_DATA_CODE, CHANNEL_CODE, COLLECTION_CODE")
            grdSATFBUD2.Text = grdSATFBUD0.Text & " - " & BUDGET_DATA_DESC & " for Channel " & CHANNEL_CODE & ", Detail by Collection"
            grdSATFBUD2.Visible = True
        Else
            grdSATFBUD2.Visible = False
        End If
    End Sub

    Function Load_from_XLS(FILENAME As String) As Boolean
        Dim successful_import As Boolean = False

        Try
            Dim workbook As SpreadsheetGear.IWorkbook = Nothing
            Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
            Dim range As SpreadsheetGear.IRange = Nothing

            workbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
            worksheet = workbook.Worksheets(0)

            Dim Start_Row As Integer = -1
            Do
                Start_Row += 1
                If Start_Row = 10 Then
                    Return False
                End If
            Loop Until worksheet.Cells(Start_Row, 0).Value & "" = "Year"

            dst.Tables("SATFBUD2").Rows.Clear()
            dst.Tables("SATFBUD0").Rows.Clear()

            Dim BUDGET_DATA_CODE As String = ""

            Me.Cursor = Cursors.WaitCursor

            ASCMAIN1.Progress("Now Loading Budgets from XLS")

            Dim i As Integer = Start_Row + 1
            Do While worksheet.Cells(i, 0).Value & "" <> ""
                Dim XL_Row As Integer = i + 1
                Dim WS_YEAR As String = worksheet.Cells(i, 0).Value & ""
                Dim WS_VERSION As String = worksheet.Cells(i, 4).Value & ""
                Dim COLLECTION_CODE As String = worksheet.Cells(i, 1).Value & ""
                Dim CHANNEL_CODE As String = worksheet.Cells(i, 2).Value & ""
                BUDGET_DATA_CODE = worksheet.Cells(i, 3).Value & ""

                Dim E As String = ""
                If dst.Tables("SATFBUDC").Rows.Find(BUDGET_DATA_CODE) Is Nothing Then
                    E &= vbCrLf & "Invalid Budget Data Code (" & BUDGET_DATA_CODE & ") in row " & CStr(XL_Row)
                Else
                    If BUDGET_DATA_CODE <> "S" And BUDGET_DATA_CODE <> "G" Then
                        E &= vbCrLf & "Invalid Budget Data Code (" & BUDGET_DATA_CODE & ") in row " & CStr(XL_Row)
                    End If
                End If
                If dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE) Is Nothing Then
                    E &= vbCrLf & "Invalid Collection Code (" & COLLECTION_CODE & ") in row " & CStr(XL_Row)
                End If
                If dst.Tables("SOTCHAN1").Rows.Find(CHANNEL_CODE) Is Nothing Then
                    E &= vbCrLf & "Invalid Channel Code (" & CHANNEL_CODE & ") in row " & CStr(XL_Row)
                Else
                    If CHANNEL_CODE <> "1" Then
                        E &= vbCrLf & "Invalid Channel Code (" & CHANNEL_CODE & ") in row " & CStr(XL_Row)
                    End If
                End If
                If WS_YEAR <> BUDGET_YEAR Then
                    E &= vbCrLf & "Invalid Budget Year (" & WS_YEAR & ") in row " & CStr(XL_Row)
                End If
                If WS_VERSION <> BUDGET_VERSION Then
                    E &= vbCrLf & "Invalid Budget Version (" & WS_VERSION & ") in row " & CStr(XL_Row)
                End If
                If E <> "" Then
                    MsgBox(E, MsgBoxStyle.OkOnly, "Import will Terminate")

                    dst.Tables("SATFBUD2").Rows.Clear()
                    dst.Tables("SATFBUD0").Rows.Clear()

                    Return False
                End If

                If dst.Tables("SATFBUD0").Rows.Find(New String() {BUDGET_YEAR, BUDGET_VERSION, BUDGET_FS, BUDGET_DATA_CODE, CHANNEL_CODE}) Is Nothing Then
                    dst.Tables("SATFBUD0").Rows.Add(New String() {BUDGET_YEAR, BUDGET_VERSION, BUDGET_FS, BUDGET_DATA_CODE, CHANNEL_CODE, "*"})
                End If

                Dim rowSATFBUD2 As DataRow = dst.Tables("SATFBUD2").NewRow
                With rowSATFBUD2
                    .Item("BUDGET_YEAR") = BUDGET_YEAR
                    .Item("BUDGET_VERSION") = BUDGET_VERSION
                    .Item("BUDGET_FS") = BUDGET_FS
                    .Item("BUDGET_DATA_CODE") = BUDGET_DATA_CODE
                    .Item("CHANNEL_CODE") = CHANNEL_CODE
                    .Item("COLLECTION_CODE") = COLLECTION_CODE
                    For m As Integer = 1 To 12
                        If m > BUDGET_ACT_THRU Then
                            .Item("BUDGET_P" & Format(m, "00")) = Val(worksheet.Cells(i, 5 + m).Value & "") * 1000
                        End If
                    Next
                End With
                dst.Tables("SATFBUD2").Rows.Add(rowSATFBUD2)

                i += 1
            Loop
            ASCMAIN1.Progress("")



            If BUDGET_ACT_THRU > 0 Then

                Dim sqlA As String = ""

                ASCMAIN1.Progress("Now Loading Actual Gross Shipments")
                BUDGET_DATA_CODE = "G"

                sqlA = ""
                For M As Integer = 1 To BUDGET_ACT_THRU
                    sqlA &= ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & BUDGET_YEAR & Format(M, "00") & "',SOTINVH2.ORDR_QTY_SHIP,0) * SOTINVH2.ORDR_UNIT_PRICE) BUDGET_P" & Format(M, "00") & vbCrLf
                Next

                ' & "   and (ICTITEM1.COLLECTION_CODE LIKE 'JCH%' OR ICTITEM1.COLLECTION_CODE LIKE 'MBC%')" & vbCrLf _

                ASCMAIN1.sql = "Select SOTTCLS1.CHANNEL_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
                    & sqlA _
                    & " from SOTINVH2,ICTITEM1,SOTTCLS1,ARTCUST1" & vbCrLf _
                    & " where SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & BUDGET_YEAR & "01' and SOTINVH2.ORDR_YYYYPP_UPDATED <= '" & BUDGET_YEAR & Format(BUDGET_ACT_THRU, "00") & "'" & vbCrLf _
                    & "   and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
                    & "   and ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
                    & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                    & "   and NVL(SOTINVH2.ORDR_UNIT_PRICE,0) <> 0" & vbCrLf _
                    & "   and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
                    & "   and SOTTCLS1.CHANNEL_CODE in ('1','2')" & vbCrLf _
                    & " group by SOTTCLS1.CHANNEL_CODE, ICTITEM1.COLLECTION_CODE"

                For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim CHANNEL_CODE As String = row.Item("CHANNEL_CODE")
                    Dim COLLECTION_CODE As String = row.Item("COLLECTION_CODE")
                    If CHANNEL_CODE = "1" Then
                        BUDGET_DATA_CODE = "G"
                    ElseIf CHANNEL_CODE = "2" Then
                        BUDGET_DATA_CODE = "W"
                    Else
                        BUDGET_DATA_CODE = "?"
                    End If

                    Dim rowSATFBUD2 As DataRow = dst.Tables("SATFBUD2").Rows.Find(New String() {BUDGET_YEAR, BUDGET_VERSION, BUDGET_FS, BUDGET_DATA_CODE, CHANNEL_CODE, COLLECTION_CODE})
                    If rowSATFBUD2 Is Nothing Then
                        If dst.Tables("SATFBUD0").Rows.Find(New String() {BUDGET_YEAR, BUDGET_VERSION, BUDGET_FS, BUDGET_DATA_CODE, CHANNEL_CODE}) Is Nothing Then
                            dst.Tables("SATFBUD0").Rows.Add(New String() {BUDGET_YEAR, BUDGET_VERSION, BUDGET_FS, BUDGET_DATA_CODE, CHANNEL_CODE, "*"})
                        End If
                        rowSATFBUD2 = dst.Tables("SATFBUD2").Rows.Add(New String() {BUDGET_YEAR, BUDGET_VERSION, BUDGET_FS, BUDGET_DATA_CODE, CHANNEL_CODE, COLLECTION_CODE})
                    End If
                    For M As Integer = 1 To BUDGET_ACT_THRU
                        rowSATFBUD2.Item("BUDGET_P" & Format(M, "00")) = row.Item("BUDGET_P" & Format(M, "00"))
                    Next
                Next


                ASCMAIN1.Progress("Now Loading Actual Retail Sales")
                BUDGET_DATA_CODE = "S"

                sqlA = ""
                For M As Integer = 1 To BUDGET_ACT_THRU
                    sqlA &= ", SUM (DECODE(RSTRETL5.OPS_YYYYPP,'" & BUDGET_YEAR & Format(M, "00") & "',RSTRETL5.AMT_SOLD,0)) BUDGET_P" & Format(M, "00") & vbCrLf
                Next

                ' & "   and (ICTITEM1.COLLECTION_CODE LIKE 'JCH%' OR ICTITEM1.COLLECTION_CODE LIKE 'MBC%')" & vbCrLf _

                ASCMAIN1.sql = "Select SOTTCLS1.CHANNEL_CODE, RSTRETL5.COLLECTION_CODE" & vbCrLf _
                    & sqlA _
                    & " from RSTRETL5,SOTTCLS1,ARTCUST1" & vbCrLf _
                    & " where RSTRETL5.OPS_YYYYPP >= '" & BUDGET_YEAR & "01' and RSTRETL5.OPS_YYYYPP <= '" & BUDGET_YEAR & Format(BUDGET_ACT_THRU, "00") & "'" & vbCrLf _
                    & "   and ARTCUST1.CUST_CODE = RSTRETL5.CUST_CODE" & vbCrLf _
                    & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                    & "   and NVL(RSTRETL5.AMT_SOLD,0) <> 0" & vbCrLf _
                    & "   and SOTTCLS1.CHANNEL_CODE in ('1')" & vbCrLf _
                    & " group by SOTTCLS1.CHANNEL_CODE, RSTRETL5.COLLECTION_CODE"

                For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim CHANNEL_CODE As String = row.Item("CHANNEL_CODE")
                    Dim COLLECTION_CODE As String = row.Item("COLLECTION_CODE")

                    Dim rowSATFBUD2 As DataRow = dst.Tables("SATFBUD2").Rows.Find(New String() {BUDGET_YEAR, BUDGET_VERSION, BUDGET_FS, BUDGET_DATA_CODE, CHANNEL_CODE, COLLECTION_CODE})
                    If rowSATFBUD2 Is Nothing Then
                        If dst.Tables("SATFBUD0").Rows.Find(New String() {BUDGET_YEAR, BUDGET_VERSION, BUDGET_FS, BUDGET_DATA_CODE, CHANNEL_CODE}) Is Nothing Then
                            dst.Tables("SATFBUD0").Rows.Add(New String() {BUDGET_YEAR, BUDGET_VERSION, BUDGET_FS, BUDGET_DATA_CODE, CHANNEL_CODE, "*"})
                        End If
                        rowSATFBUD2 = dst.Tables("SATFBUD2").Rows.Add(New String() {BUDGET_YEAR, BUDGET_VERSION, BUDGET_FS, BUDGET_DATA_CODE, CHANNEL_CODE, COLLECTION_CODE})
                    End If
                    For M As Integer = 1 To BUDGET_ACT_THRU
                        rowSATFBUD2.Item("BUDGET_P" & Format(M, "00")) = row.Item("BUDGET_P" & Format(M, "00"))
                    Next
                Next

            End If


            If BUDGET_ACT_THRU < 12 Then

                ASCMAIN1.Progress("Now Loading Wholesale Gross Shipments")
                BUDGET_DATA_CODE = "W"

                Dim sqlS As String = ""
                Dim sqlF As String = ""
                For M As Integer = 1 To 6
                    sqlS &= ", SUM (CASE WHEN SUBSTR(RSTSSPL3.SEASON_CODE,5,1) = 'S' THEN RSTSSPL3.AMT_" & CStr(M) & " ELSE 0 END) BUDGET_P" & Format(M + 0, "00") & vbCrLf
                    sqlF &= ", SUM (CASE WHEN SUBSTR(RSTSSPL3.SEASON_CODE,5,1) = 'S' THEN 0 ELSE RSTSSPL3.AMT_" & CStr(M) & " END) BUDGET_P" & Format(M + 6, "00") & vbCrLf
                Next

                ASCMAIN1.sql = "Select ICTITEM1.COLLECTION_CODE, SOTTCLS1.CHANNEL_CODE" & vbCrLf _
                    & sqlS _
                    & sqlF _
                    & " from RSTSSPL3, ARTCUST1, SOTTCLS1, ICTITEM1" & vbCrLf _
                    & " where ARTCUST1.CUST_CODE = RSTSSPL3.CUST_CODE" & vbCrLf _
                    & "   and ICTITEM1.ITEM_CODE = RSTSSPL3.ITEM_CODE" & vbCrLf _
                    & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" _
                    & "   and SOTTCLS1.CHANNEL_CODE = '2'" & vbCrLf _
                    & "   and RSTSSPL3.SEASON_CODE IN ('" & BUDGET_YEAR & "S','" & BUDGET_YEAR & "F')" & vbCrLf _
                    & " group by ICTITEM1.COLLECTION_CODE, SOTTCLS1.CHANNEL_CODE"

                For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim CHANNEL_CODE As String = row.Item("CHANNEL_CODE")
                    Dim COLLECTION_CODE As String = row.Item("COLLECTION_CODE")

                    Dim rowSATFBUD2 As DataRow = dst.Tables("SATFBUD2").Rows.Find(New String() {BUDGET_YEAR, BUDGET_VERSION, BUDGET_FS, BUDGET_DATA_CODE, CHANNEL_CODE, COLLECTION_CODE})
                    If rowSATFBUD2 Is Nothing Then
                        If dst.Tables("SATFBUD0").Rows.Find(New String() {BUDGET_YEAR, BUDGET_VERSION, BUDGET_FS, BUDGET_DATA_CODE, CHANNEL_CODE}) Is Nothing Then
                            dst.Tables("SATFBUD0").Rows.Add(New String() {BUDGET_YEAR, BUDGET_VERSION, BUDGET_FS, BUDGET_DATA_CODE, CHANNEL_CODE, "*"})
                        End If
                        rowSATFBUD2 = dst.Tables("SATFBUD2").Rows.Add(New String() {BUDGET_YEAR, BUDGET_VERSION, BUDGET_FS, BUDGET_DATA_CODE, CHANNEL_CODE, COLLECTION_CODE})
                    End If
                    For M As Integer = BUDGET_ACT_THRU + 1 To 12
                        rowSATFBUD2.Item("BUDGET_P" & Format(M, "00")) = Val(row.Item("BUDGET_P" & Format(M, "00"))) * 1000
                    Next
                Next
            End If

            'dst.Tables("SATFBUD0").Rows.Clear()
            'For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SATFBUD2"), New String() {"BUDGET_YEAR", "BUDGET_VERSION", "BUDGET_DATA_CODE", "CHANNEL_CODE"}).Select()
            '    'dst.Tables("SATFBUD0").Rows.Add(row.ItemArray)
            '    dst.Tables("SATFBUD0").Rows.Add(New String() {row.Item("BUDGET_YEAR"), row.Item("BUDGET_VERSION"), row.Item("BUDGET_DATA_CODE"), row.Item("CHANNEL_CODE"), "*"})
            'Next

            successful_import = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Cannot continue with Budget Upload")
        End Try

        Return successful_import
    End Function

    Private Sub grdSATFBUD0_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSATFBUD0.InitializeRow
        Dim CHANNEL_CODE As String = e.Row.Cells("CHANNEL_CODE").Value & ""
        Dim rowSOTCHAN1 As DataRow = dst.Tables("SOTCHAN1").Rows.Find(CHANNEL_CODE)
        Dim CHANNEL_DESC As String = ""
        If rowSOTCHAN1 IsNot Nothing Then
            CHANNEL_DESC = rowSOTCHAN1.Item("CHANNEL_DESC") & ""
        End If
        e.Row.Cells("LINE_DATA_DESC").Value = e.Row.Cells("BUDGET_DATA_CODE").Text & " " & CHANNEL_DESC

    End Sub

    Private Sub grdSATFBUD2_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSATFBUD2.InitializeRow
        Dim CHANNEL_CODE As String = e.Row.Cells("CHANNEL_CODE").Value & ""
        Dim rowSOTCHAN1 As DataRow = dst.Tables("SOTCHAN1").Rows.Find(CHANNEL_CODE)
        Dim CHANNEL_DESC As String = ""
        If rowSOTCHAN1 IsNot Nothing Then
            CHANNEL_DESC = rowSOTCHAN1.Item("CHANNEL_DESC") & ""
        End If
        e.Row.Cells("LINE_DATA_DESC").Value = e.Row.Cells("BUDGET_DATA_CODE").Text & " " & CHANNEL_DESC
    End Sub
End Class