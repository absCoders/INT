Public Class GLFBUDM1

    Dim ACCT_YEAR As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GLTPARM1")

        With dst
            ASCMAIN1.sql = "Select GLTACCT2.*" _
                & ", GLTACCT1.ACCT_TYPE, GLTACCT1.ACCT_DESC" _
                & " from GLTACCT2,GLTACCT1 " _
                & " where GLTACCT1.ACCT_CODE (+)= GLTACCT2.ACCT_CODE" _
                & "   and GLTACCT2.ACCT_YEAR = :PARM1"
            Create_TDA(.Tables.Add, "GLTACCT2", "**", 0, True, "V")
            .Tables("GLTACCT2").Columns.Add("ACCT_END_BAL", GetType(System.Decimal), "ISNULL(ACCT_BEG_BAL,0) + ISNULL(ACCT_BUD_P01,0) + ISNULL(ACCT_BUD_P02,0) + ISNULL(ACCT_BUD_P03,0) + ISNULL(ACCT_BUD_P04,0) + ISNULL(ACCT_BUD_P05,0) + ISNULL(ACCT_BUD_P06,0) + ISNULL(ACCT_BUD_P07,0) + ISNULL(ACCT_BUD_P08,0) + ISNULL(ACCT_BUD_P09,0) + ISNULL(ACCT_BUD_P10,0) + ISNULL(ACCT_BUD_P11,0) + ISNULL(ACCT_BUD_P12,0) + ISNULL(ACCT_BUD_P13,0)")

            Create_TDA(.Tables.Add, "GLTACCT1", "*", 0, False)
            Fill_Records("GLTACCT1")
            Create_TDA(.Tables.Add, "GLTSEGM1", "*", 0, False)
            Fill_Records("GLTSEGM1")

        End With

        grdGLTACCT2.DataSource = dst.Tables("GLTACCT2")

        'Add_ACCT_TYPEs("GLTACCT2", True)
        Set_SEGS(grdGLTACCT2, "GLTACCT2")

        Create_Summary(grdGLTACCT2, "ACCT_CODE", "Count")
        For Each gcol As UltraWinGrid.UltraGridColumn In grdGLTACCT2.DisplayLayout.Bands(0).Columns
            gcol.Header.Appearance.BackColor = Drawing.Color.White
            gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            If gcol.Key = "ACCT_BEG_BAL" Or gcol.Key = "ACCT_END_BAL" Then
                Create_Summary(grdGLTACCT2, gcol.Key)
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                If gcol.Key = "ACCT_BEG_BAL" Then
                    gcol.Header.Caption = "Beg Bal"
                Else
                    gcol.Header.Caption = "End Bal"
                End If
                gcol.Width = 100
            ElseIf gcol.Key.StartsWith("ACCT_BUD_P") Then
                Create_Summary(grdGLTACCT2, gcol.Key)
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                gcol.Width = 90
            Else
                gcol.Header.Fixed = True
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
            End If
        Next
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Edit"
                ACCT_YEAR = Absx1.txtFor("ACCT_YEAR").Text
                If Not ASCMAIN1.Logical_Lock("GLTACCT2", ACCT_YEAR) Then Exit Sub

            Case "Update"

                For Each row As DataRow In ASCDATA1.SelectDistinct("GLTACCT2", "ACCT_CODE").Select("")
                    Dim ACCT_CODE As String = row.Item(0)
                    If dst.Tables("GLTACCT1").Rows.Find(ACCT_CODE) Is Nothing Then
                        EMsg &= "Invalid Acct Code " & ACCT_CODE
                    End If
                Next
                For i As Integer = 2 To 4
                    For Each row As DataRow In ASCDATA1.SelectDistinct("GLTACCT2", "SEG" & Format(i, "0") & "_CODE").Select("")
                        Dim SEGX_CODE As String = row.Item(0)
                        If dst.Tables("GLTSEGM1").Rows.Find(New String() {Format(i, "0"), SEGX_CODE}) Is Nothing Then
                            EMsg &= "Invalid " & grdGLTACCT2.DisplayLayout.Bands(0).Columns("SEG" & Format(i, "0") & "_CODE").Header.Caption & " Code " & SEGX_CODE
                        End If
                    Next
                Next

                If EMsg.Length > 1000 Then
                    EMsg = Mid(EMsg, 1, 1000) & " ..."
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
            Case "View", "Edit"
                If eItemKey = "Edit" Then
                    EntryMode = "E"
                Else
                    EntryMode = "V"
                End If
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Done", "Cancel"
                Mode_Settings(False)
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    If (EntryMode = "V" And ScreenMode) Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If
                    .Items("Update").Settings.Enabled = iScreenMode

                    .Items("Delete").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    '.Items("Print").Settings.Enabled = iScreenMode

                    .Items("Edit").Visible = Not InquiryMode
                    .Items("Done").Visible = (EntryMode = "V" And ScreenMode)
                    '.Items("Print").Visible = (EntryMode = "V" And ScreenMode) ' False ' ScreenMode
                    .Items("Update").Visible = Not InquiryMode And (Not (EntryMode = "V") Or Not ScreenMode)
                    .Items("Delete").Visible = Not InquiryMode And (EntryMode = "E")
                    .Items("Cancel").Visible = Not InquiryMode And (Not (EntryMode = "V") Or Not ScreenMode)
                End With
            End With

        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdGLTACCT2.Visible = tf



        With grdGLTACCT2.DisplayLayout.Override
            If EntryMode = "E" Then
                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.True
            Else
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End If
        End With

        If ScreenMode Then

            ASCMAIN1.sql = "Select Count (*) from GLTACCT4 where ACCT_YEAR = '" & ACCT_YEAR & "'"
            chkSaveAsOriginal.Visible = (EntryMode = "E")
            Dim ENABLE_SAVE As Boolean = (ASCDATA1.GetDataValue = 0) And (EntryMode = "E")
            Set_Read_Only_for_ctl(chkSaveAsOriginal, Not ENABLE_SAVE)
        Else
            chkSaveAsOriginal.Visible = False
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"GLTACCT2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        chkSaveAsOriginal.Checked = False

        Absx1.txtFor("ACCT_YEAR").Text = Mid(ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP"), 1, 4)
    End Sub

    Sub Load_Record()
        ASCMAIN1.Progress("Now Loading Account Budget Data")
        Save_Header_Fields(UltraGroupBox1)
        ACCT_YEAR = HFs("ACCT_YEAR")

        For p As Integer = 1 To 12
            Dim L As String = ASCMAIN1.Get_Legend(ACCT_YEAR & Format(p, "00"), False, True)
            grdGLTACCT2.DisplayLayout.Bands(0).Columns("ACCT_BUD_P" & Format(p, "00")).Header.Caption = L
        Next

        Fill_Records("GLTACCT2", ACCT_YEAR)
        Sort_grdColumns(grdGLTACCT2, "ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE")
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()
        Update_Record_TDA("GLTACCT2", "ACCT_YEAR = '" & ACCT_YEAR & "'")
        If chkSaveAsOriginal.Checked Then
            ASCMAIN1.sql = "Insert into GLTACCT4 Select * from GLTACCT2 where ACCT_YEAR = '" & ACCT_YEAR & "'"
            ASCDATA1.ExecuteSQL()
        End If
        CommitTrans("Update Complete")
    End Sub
#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdGLTACCT2, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Acct Inquiry", "Load LY Actuals")
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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If tlb_pop.Tools.Exists("Load LY Actuals") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Load LY Actuals"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = (ScreenMode And (EntryMode = "E"))
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            'Case "Show Description"
            '    Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
            '    grd.DisplayLayout.Bands(0).Columns("ITEM_DESC").Hidden = Not tlb_sbt.Checked
            Case "Load LY Actuals"
                If MsgBox("Do you want to clear the current budget data and load LY Actuals as Budgets?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                    dst.Tables("GLTACCT2").Rows.Clear()
                    ASCMAIN1.sql = "Select GLTACCT3.*, GLTACCT1.ACCT_TYPE, GLTACCT1.ACCT_DESC" & vbCrLf _
                        & " from GLTACCT1,GLTACCT3" & vbCrLf _
                        & " where GLTACCT1.ACCT_CODE = GLTACCT3.ACCT_CODE" & vbCrLf _
                        & "   and GLTACCT3.ACCT_YEAR = '" & Format(Val(ACCT_YEAR) - 1, "0000") & "'" & vbCrLf _
                        & "   and GLTACCT1.ACCT_TYPE in ('I','X')"
                    For Each row As DataRow In ASCDATA1.GetDataTable.Select
                        Dim rowGLTACCT2 As DataRow = dst.Tables("GLTACCT2").NewRow
                        With rowGLTACCT2
                            For Each C As String In New String() _
                                {"ACCT_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE", "ACCT_TYPE", "ACCT_DESC"}
                                .Item(C) = row.Item(C)
                            Next

                            For I As Integer = 1 To 13
                                .Item("ACCT_BUD_P" & Format(I, "00")) = row.Item("ACCT_ACT_P" & Format(I, "00"))
                            Next
                        End With
                        dst.Tables("GLTACCT2").Rows.Add(rowGLTACCT2)
                    Next
                    Sort_grdColumns(grdGLTACCT2, "ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE")
                    MsgBox("LY Actuals have been loaded as Budgets", MsgBoxStyle.OkOnly, "Verification")
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Acct Inquiry"
                Dim ACCT_CODE As String = grd.ActiveRow.Cells("ACCT_CODE").Text
                Context_Launch("View", ACCT_CODE, e.Tool.Key, "GLFACTI1")
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

            Case "ACCT_YEAR"

                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Click_Command("View", e)
                End If

        End Select

    End Sub

    Overrides Sub num_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.num_KeyDown(sender, e)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "QE_INV_AMT"

        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "ACCT_YEAR"
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "ACCT_YEAR"

        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "ACCT_YEAR"
                Click_Command("View")
        End Select
    End Sub

#End Region

#Region "grdGLTACCT2"

    Private Sub grdGLTACCT2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTACCT2.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "ACCT_CODE"
                Dim ACCT_CODE As String = e.Cell.Value & ""

                grdCodeDesc(grdGLTACCT2, "GLTACCT1", "ACCT_CODE", "ACCT_DESC")
                grdCodeDesc(grdGLTACCT2, "GLTACCT1", "ACCT_CODE", "ACCT_TYPE")
                For i As Integer = 2 To 4
                    If e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Text = "" Then
                        e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Value = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i))
                    End If
                Next
        End Select
    End Sub

    Private Sub grdGLTACCT2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdGLTACCT2.AfterExitEditMode
        If grdGLTACCT2.ActiveCell Is Nothing Then Exit Sub
        With grdGLTACCT2
            Select Case .ActiveCell.Column.Key
                Case "ACCT_CODE"
                    Dim ACCT_CODE As String = .ActiveCell.Text
                    If ACCT_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(ACCT_CODE, .ActiveCell.Column.Key)
                    End If
            End Select
        End With
    End Sub

    Private Sub grdGLTACCT2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdGLTACCT2.AfterRowActivate
        With grdGLTACCT2
            If .ActiveRow.IsAddRow Then
                .DisplayLayout.Bands(0).Columns("ACCT_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .ActiveCell = grdGLTACCT2.ActiveRow.Cells("ACCT_CODE")
                .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                '.DisplayLayout.Bands(0).Columns("ACCT_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                ' why cant we edit the acct code?
            End If
        End With
    End Sub

    Private Sub grdGLTACCT2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdGLTACCT2.AfterRowUpdate
        grdGLTACCT2.DisplayLayout.ColScrollRegions(0).Scroll(UltraWinGrid.ColScrollAction.Left)
    End Sub

    Private Sub grdGLTACCT2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdGLTACCT2.BeforeRowUpdate
        With grdGLTACCT2
            If e.Row.Cells("ACCT_CODE").Text = "" Then
                EMsg2 = "No value entered for Acct Code"
                e.Cancel = True
            Else
                LookUp("GLTACCT1", e.Row.Cells("ACCT_CODE").Text)
                If cdr Is Nothing Then
                    EMsg2 = "Invalid Value entered for Acct Code (" & e.Row.Cells("ACCT_CODE").Text & ")"
                    If Not loading_grd_from_Excel Then MsgBox(EMsg2, MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            Dim COLUMN_NAME As String
            For i As Integer = 2 To 4
                COLUMN_NAME = "SEG" & CStr(i) & "_CODE"
                If Not e.Row.Cells(COLUMN_NAME).Column.Hidden Then
                    If e.Row.Cells(COLUMN_NAME).Text = "" Then
                        EMsg2 = "No value entered for " & COLUMN_NAME
                        e.Cancel = True
                    Else
                        LookUp("GLTSEGM1", New String() {CStr(i), e.Row.Cells(COLUMN_NAME).Text})
                        If cdr Is Nothing Then
                            EMsg2 = "Invalid Value entered for " & e.Row.Cells(COLUMN_NAME).Column.Header.Caption & " (" & e.Row.Cells(COLUMN_NAME).Text & ")"
                            If Not loading_grd_from_Excel Then MsgBox(EMsg2, MsgBoxStyle.OkOnly, "Cannot Update Row")
                            e.Cancel = True
                        End If
                    End If
                End If
            Next

            If Not e.Cancel Then
                If e.Row.Cells("ACCT_YEAR").Text = "" Then
                    .ActiveRow.Cells("ACCT_YEAR").Value = ACCT_YEAR
                End If
            End If
        End With

    End Sub

    Private Sub grdGLTACCT2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTACCT2.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdGLTACCT2, sql_where, sql_where <> "")
    End Sub
#End Region

    Private Sub chkSaveAsOriginal_CheckedChanged(sender As Object, e As EventArgs) Handles chkSaveAsOriginal.CheckedChanged

    End Sub
End Class