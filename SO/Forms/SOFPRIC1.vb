Public Class SOFPRIC1

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If MENU_ITEM_OBJECT = "BMFMAINI" Then
            InquiryMode = True
        End If

        With dst

            ASCMAIN1.sql = "Select BMTMAIN1.*" _
            & " from BMTMAIN1 where BM_PROD_ITEM = :PARM1"
            Create_TDA(.Tables.Add, "BMTMAIN1", "**", 0, True, "V", 1)

            ASCMAIN1.sql = "Select BMTMAIN2.*" _
            & " from BMTMAIN2 where BM_PROD_ITEM = :PARM1"
            Create_TDA(.Tables.Add, "BMTMAINX", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select BMTMAIN2.*" _
            & " from BMTMAIN2 where BM_PROD_ITEM = :PARM1 and BM_ISSUE_NO = :PARM2"
            Create_TDA(.Tables.Add, "BMTMAIN2", "**", 0, True, "VN", 2)

            ASCMAIN1.sql = "Select BMTMAIN3.*" _
            & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_UOM" _
            & ", ICTITEM1.ITEM_COST_STD, ICTITEM1.ITEM_COST_WASTE_PCT" _
            & ", ICTITEM1.VEND_ITEM_CODE" _
            & " from BMTMAIN3,ICTITEM1 " _
            & " where BMTMAIN3.BM_PROD_ITEM = :PARM1" _
            & " and BMTMAIN3.BM_ISSUE_NO = :PARM2" _
            & " and ICTITEM1.ITEM_CODE = BMTMAIN3.BM_COMP_ITEM"
            Create_TDA(.Tables.Add, "BMTMAIN3", "**", 0, True, "VN", 3)
            .Tables("BMTMAIN3").Columns.Add("EXT_COST", GetType(System.Decimal), "ISNULL(BM_QTY_PER_ASSY,0) * ISNULL(ITEM_COST_STD,0) * (1 + ISNULL(ITEM_COST_WASTE_PCT,0)/100)")
            .Tables("BMTMAIN3").Columns.Add("QTY_ONH", GetType(System.Int32))
            .Tables("BMTMAIN3").Columns.Add("QTY_AVA", GetType(System.Int32))
        End With



        grdSOTPRIC1.DataSource = dst.Tables("BMTMAINX")
        grdBMTMAIN3.DataSource = dst.Tables("BMTMAIN3")


        Call Create_Summary(grdSOTPRIC1, "BM_ISSUE_NO", "Count")

        Call Create_Summary(grdBMTMAIN3, "BM_COMP_ITEM", "Count")
        Call Create_Summary(grdBMTMAIN3, "EXT_COST")

        With grdBMTMAIN3.DisplayLayout.Bands("BMTMAIN3")
            .Columns("BM_SEQ").Header.Fixed = True
            .Columns("BM_COMP_ITEM").Header.Fixed = True
            .Columns("ITEM_DESC").Header.Fixed = True
        End With


        For Each COLUMN_NAME As String In New String() {"ITEM_DESC", "ITEM_UOM"}
            With grdBMTMAIN3.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                .CellAppearance.BackColor = Drawing.Color.Beige
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End With
        Next
        For Each COLUMN_NAME As String In New String() {"ITEM_COST_STD", "ITEM_COST_WASTE_PCT", "EXT_COST"}
            With grdBMTMAIN3.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                .CellAppearance.BackColor = Drawing.Color.Yellow
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End With
        Next
        For Each COLUMN_NAME As String In New String() {"QTY_ONH", "QTY_AVA"}
            With grdBMTMAIN3.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                .CellAppearance.BackColor = Drawing.Color.LightBlue
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End With
        Next

        ASCMAIN1.Add_Value_List(grdBMTMAIN3, "BM_WHEN_EXHAUSTED", , New String() {":", "A", "B", "C", "D"})
        ASCMAIN1.Add_Value_List(grdBMTMAIN3, "BM_REPLACE_WITH", , New String() {":", "A", "B", "C", "D"})


        Show_grdBMTMAIN3_Columns()

        Check_Inquiry_Mode()
    End Sub

    Sub Check_Inquiry_Mode()
        With UltraExplorerBar1.Groups("Screen Control")
            .Items("New").Visible = Not InquiryMode
            .Items("Edit").Visible = Not InquiryMode
            .Items("Load").Visible = InquiryMode
            .Items("Update").Visible = Not InquiryMode
            .Items("Cancel").Visible = Not InquiryMode
            .Items("Done").Visible = InquiryMode
        End With

        If InquiryMode Then
            grdBMTMAIN3.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grdBMTMAIN3.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            grdBMTMAIN3.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        Else
            grdBMTMAIN3.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            grdBMTMAIN3.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            grdBMTMAIN3.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New", "Edit", "Load"
                Validate_Code("BM_PROD_ITEM")

                If EMsg = "" And eItemKey <> "Load" Then
                    If Not ASCMAIN1.Logical_Lock("BMTMAIN1", Absx1.txtFor("BM_PROD_ITEM").Text) Then
                        Exit Sub
                    End If
                End If


            Case "Update"
                Stop ' CHECK STUFF
                ' MATCHING XR
                ' IF USE FOR STD THEN ALL COMPS MUST HAVE A STD COST
                ' AT LEAST 1 COMP NOT A VSM

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New", "Edit", "Load"
                EntryMode = "L"
                Call Load_Record()
                Call Mode_Settings(True)

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
                '.Groups("Display Options").Visible = ScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)
        'UltraTabControl1.Visible = ScreenMode
        grdSOTPRIC1.Visible = ScreenMode

        tab.Tabs("BM Issues").Visible = Not ScreenMode
        tab.Tabs("BM Details").Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("BMTMAIN1").Rows.Clear()
        dst.Tables("BMTMAIN2").Rows.Clear()
        dst.Tables("BMTMAIN3").Rows.Clear()
        dst.EnforceConstraints = True

        Absx1.txtFor("BM_PROD_ITEM").Text = ""
        grdSOTPRIC1.Text = "BM Issues"
    End Sub

    Sub Load_Record()

        Call ASCMAIN1.Progress("Now Loading Data ...")

        Call Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)
        Fill_Records("BMTMAIN1", HFs("BM_PROD_ITEM"))
        Fill_Records("BMTMAIN2", HFs("BM_PROD_ITEM"))
        Fill_Records("BMTMAIN3", HFs("BM_PROD_ITEM"))
        Sort_grdColumns(grdBMTMAIN3, "BM_SEQ")
        EnforceConstraints(True)

        Call ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Call BeginTrans()

        Dim sql_delete As String = "BM_PROD_ITEM = '" & HFs("BM_PROD_ITEM") & "'"
        Update_Record_TDA("BMTMAIN1", sql_delete)
        Update_Record_TDA("BMTMAIN2", sql_delete)
        Update_Record_TDA("BMTMAIN3", sql_delete)

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
        Call Load_Popup_Menu(grdBMTMAIN3, "SSS", "Show Costing Data", "Show Qty Data", "Show Misc Data")
    End Sub

    Public Overrides Sub tlb_beforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

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
            '    tlb_sbt = DirectCast(tlb_pop.Tools("Show Costing Data"), UltraWinToolbars.StateButtonTool)
            '    tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
            '    tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            '    tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
            'End If

            Select Case e.SourceControl.Name
                'Case "grdDPTITMFX"
                '    If grdBMTMAIN2.Tag = "" Then
                '        e.Cancel = True
                '    End If
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Select All"
                'For Each rowICTCOLLX As DataRow In dst.Tables("ICTCOLLX").Rows
                '    rowICTCOLLX.Item("SELECTED") = "1"
                'Next

            Case "Clear All"
                'For Each rowICTCOLLX As DataRow In dst.Tables("ICTCOLLX").Rows
                '    rowICTCOLLX.Item("SELECTED") = "0"
                'Next

            Case "Show Costing Data", "Show Qty Data", "Show Misc Data"
                Show_grdBMTMAIN3_Columns()

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
            Case "BM_PROD_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    'Call Click_Command("Load", e)
                    Show_Issues(Absx1.txtFor("BM_PROD_CODE").Text)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "BM_PROD_CODE"
                If InquiryMode Then
                    'Call Click_Command("Load")
                Else
                    'Call Click_Command("New")
                End If
                Show_Issues(Absx1.txtFor("BM_PROD_CODE").Text)
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

            Case "BM_PROD_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("BM_PROD_CODE").Text <> "" Then
                        Call LookUp("ICTITEM1", Absx1.txtFor("BM_PROD_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

        End Select
    End Sub

#End Region

    Sub Show_Issues(ByVal ITEM_CODE As String)
        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE, True)
        Absx1.txtFor("ITEM_DESC").Text = rowICTITEM1.Item("ITEM_DESC")
        Absx1.txtFor("ITEM_UOM").Text = rowICTITEM1.Item("ITEM_UOM")
        Fill_Records("BMTMAIN2", ITEM_CODE)
        grdSOTPRIC1.Text = "BM Issues on file for Item " & ITEM_CODE
        Sort_grdColumns(grdSOTPRIC1, "BM_ISSUE_NO")
    End Sub

    Sub Show_grdBMTMAIN3_Columns()
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool
        With grdBMTMAIN3.DisplayLayout.Bands(0)
            tlb_sbt = DirectCast(tlb.Tools("Show Costing Data"), UltraWinToolbars.StateButtonTool)
            .Columns("ITEM_COST_STD").Hidden = Not tlb_sbt.Checked
            .Columns("ITEM_COST_WASTE_PCT").Hidden = Not tlb_sbt.Checked
            .Columns("EXT_COST").Hidden = Not tlb_sbt.Checked

            tlb_sbt = DirectCast(tlb.Tools("Show Qty Data"), UltraWinToolbars.StateButtonTool)
            .Columns("QTY_ONH").Hidden = Not tlb_sbt.Checked
            .Columns("QTY_AVA").Hidden = Not tlb_sbt.Checked

            tlb_sbt = DirectCast(tlb.Tools("Show Misc Data"), UltraWinToolbars.StateButtonTool)
            .Columns("BM_COMP_COMMENT").Hidden = Not tlb_sbt.Checked
            .Columns("VEND_ITEM_CODE").Hidden = Not tlb_sbt.Checked
        End With
    End Sub

#Region "grdBMTMAIN3"

    Private Sub grdBMTMAIN3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdBMTMAIN3.AfterCellUpdate
        If e.Cell.Column.Key = "BM_COMP_ITEM" Then
            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", New String() {e.Cell.Value})
            If rowICTITEM1 IsNot Nothing Then
                e.Cell.Row.Cells("ITEM_DESC").Value = rowICTITEM1.Item("ITEM_DESC") & ""
            End If
        End If
    End Sub

    Private Sub grdBMTMAIN3_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdBMTMAIN3.AfterExitEditMode
        With grdBMTMAIN3
            Select Case .ActiveCell.Column.Key
                Case "BM_COMP_ITEM"
                    If .ActiveCell.Text <> "" Then
                        '.ActiveCell.Value = ASCMAIN1.Format_Field(.ActiveCell.Text, .ActiveCell.Column.Key)
                        cdr = LookUp("ICTITEM1", .ActiveCell.Text)
                        .ActiveRow.Cells("ITEM_DESC").Value = cdr.Item("ITEM_DESC") & ""
                        .ActiveRow.Cells("ITEM_UOM").Value = cdr.Item("ITEM_UOM") & ""
                        .ActiveRow.Cells("ITEM_COST_STD").Value = Val(cdr.Item("ITEM_COST_STD") & "")
                        .ActiveRow.Cells("ITEM_COST_WASTE_PCT").Value = Val(cdr.Item("ITEM_COST_WASTE_PCT") & "")
                        .ActiveRow.Cells("VEND_ITEM_CODE").Value = cdr.Item("VEND_ITEM_CODE") & ""
                    End If
            End Select
        End With
    End Sub

    Private Sub grdBMTMAIN3_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdBMTMAIN3.BeforeExitEditMode

        If e.CancellingEditOperation Then Exit Sub
        With grdBMTMAIN3.ActiveCell
            Select Case .Column.Key

                Case "BM_COMP_ITEM"

                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                        cdr = LookUp("ICTITEM1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid " & .Column.Header.Caption & "(" & .Text & ")")
                            .Value = ""
                            e.Cancel = True
                        End If
                    End If
            End Select
        End With

    End Sub

    Private Sub grdBMTMAIN3_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdBMTMAIN3.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In grdBMTMAIN3.Selected.Rows
            'If dst.Tables("BMTMAIN3").Rows(grow.ListIndex).RowState = DataRowState.Added Then
            'Else
            '    MsgBox("Cannot Delete Existing Item Records", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            '    e.Cancel = True
            '    Exit For
            'End If
        Next
    End Sub

    Private Sub grdBMTMAIN3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdBMTMAIN3.BeforeRowUpdate

        'If optCOLLECTION_CODE.Value <> "I" Then
        '    cdr = LookUp("ICTCOLL1", e.Row.Cells("COLLECTION_CODE").Text & "")
        '    If cdr Is Nothing Then
        '        'ASCMAIN1.Progress("Invalid Collection Code (" & e.Row.Cells("COLLECTION_CODE").Text & ")")
        '        e.Cancel = True
        '    Else
        '        If cdr.Item("BRAND_CODE") <> HFs("BRAND_CODE") Then
        '            'ASCMAIN1.Progress("Collection does not belong to Brand " & HFs("BRAND_CODE"))
        '            e.Cancel = True
        '        End If
        '    End If
        'End If

        'If Val(e.Row.Cells("QTY_TOTAL").Value & "") = 0 And Val(e.Row.Cells("AMT_TOTAL").Value & "") = 0 Then
        '    e.Cancel = True
        'End If

        If e.Row.IsAddRow And Not e.Cancel Then
            e.Row.Cells("BM_PROD_ITEM").Value = Absx1.txtFor("BM_PROD_ITEM").Text
            e.Row.Cells("BM_ISSUE_NO").Value = Absx1.txtFor("BM_ISSUE_NO").Text
        End If
    End Sub

    Private Sub grdBMTMAIN3_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdBMTMAIN3.ClickCellButton
        Dim sql_where As String = ""
        Select Case grdBMTMAIN3.ActiveCell.Column.Key
            Case "BM_COMP_ITEM"
        End Select

        Call grdClickCellButton(grdBMTMAIN3, sql_where, False)
    End Sub

#End Region
End Class