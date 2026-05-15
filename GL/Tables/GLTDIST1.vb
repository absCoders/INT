Public Class GLTDIST1

    Dim sqlGLTDSTR2 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GLTPARM1")

        Dim VL As ValueList = optSegment.ValueList
        VL.ValueListItems(0).DisplayText = ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC")
        VL.ValueListItems(1).DisplayText = ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC")
        VL.ValueListItems(2).DisplayText = ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC")

        With dst
            ASCMAIN1.sql = "Select GLTDIST2.*, GLTSEGM1.ACCT_SEG_DESC" _
                & " from GLTDIST2,GLTSEGM1" _
                & " where GLTDIST2.DIST_CODE = :PARM1" _
                & "   and GLTSEGM1.ACCT_SEG_ID = GLTDIST2.ACCT_SEG_ID" _
                & "   and GLTSEGM1.ACCT_SEG_CODE = GLTDIST2.ACCT_SEG_CODE"
            Create_TDA(.Tables.Add, "GLTDIST2", "**", 0, True, "V", 3)
            '  .Tables("GLTDIST2").Columns("DIST_PCT").DefaultValue = 0

            Create_TDA(.Tables.Add, "GLTSEGM1", "*", 0, False, , 2)
        End With

        Fill_Records("GLTSEGM1")

        grdGLTDIST2.DataSource = dst.Tables("GLTDIST2")

        With grdGLTDIST2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue

                If gcol.Key = "DIST_PCT" Then
                Else
                    gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                End If
            Next
        End With

        Create_Summary(grdGLTDIST2, "ACCT_SEG_CODE", "Count")
        Create_Summary(grdGLTDIST2, "DIST_PCT")

        With grdGLTDIST2.DisplayLayout.Bands(0)
            .Columns("ACCT_SEG_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            .Columns("ACCT_SEG_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
        End With

        With chkShowNonZeroOnly.Appearance
            .ForeColor = System.Drawing.Color.White
            .BackColor = System.Drawing.Color.FromArgb(98, 160, 232)
            .BackColor2 = System.Drawing.Color.FromArgb(83, 115, 191)
            .BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        End With


    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdGLTDIST2, "S", "Show Filter")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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

        Select Case grd.Name

            'Case "grdGLTDSTR2"
            '    tlb_btn = DirectCast(tlb_pop.Tools("Add Codes"), UltraWinToolbars.ButtonTool)
            '    tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Select Case e.Tool.Key
            'Case "Add Codes"
            '    If grd.Name = "grdGLTDSTR2" Then
            '        Add_Codes(grdGLTDIST2, "GLTACCT1", "ACCT_CODE", "Accounts")
            '    End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub
#End Region

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
            Case "Edit"
            Case "Update"
                For Each ACCT_SEG_ID As String In New String() {"2", "3", "4"}
                    If Val(dst.Tables("GLTDIST2").Compute("SUM(DIST_PCT)", "ACCT_SEG_ID = '" & ACCT_SEG_ID & "'") & "") <> 100 Then
                        If dst.Tables("GLTDIST2").Select("ACCT_SEG_ID = '" & ACCT_SEG_ID & "' and ISNULL(DIST_PCT,0) <> 0").Length <> 0 Then
                            EMsg &= vbCr & "Distribution does NOT add up to 100 for " & ROWs("GLTPARM1").Item("GL_PARM_SEG" & ACCT_SEG_ID & "_DESC")
                        End If
                    End If
                Next

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()

        ASCDATA1.DeleteRows("GLTDIST2", "ISNULL(DIST_PCT,0) = 0")
        Dim DIST_CODE As String = Absx1.txtFor("DIST_CODE").Text
        Dim sqlDelete = "DIST_CODE = '" & DIST_CODE & "'"
        Update_Record_TDA("GLTDIST2", sqlDelete)
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Dim DIST_CODE As String = Absx1.txtFor("DIST_CODE").Text
        Fill_Records("GLTDIST2", New String() {DIST_CODE})
        Sort_grdColumns(grdGLTDIST2, "ACCT_SEG_CODE")
        EnforceConstraints(True)

        For Each row As DataRow In dst.Tables("GLTSEGM1").Select()
            Dim ACCT_SEG_ID As String = row.Item("ACCT_SEG_ID")
            Dim ACCT_SEG_CODE As String = row.Item("ACCT_SEG_CODE")
            Dim rowGLTDIST2 As DataRow = dst.Tables("GLTDIST2").Rows.Find(New String() {DIST_CODE, ACCT_SEG_ID, ACCT_SEG_CODE})
            If rowGLTDIST2 Is Nothing Then
                rowGLTDIST2 = dst.Tables("GLTDIST2").NewRow
                rowGLTDIST2.Item("DIST_CODE") = DIST_CODE
                rowGLTDIST2.Item("ACCT_SEG_ID") = ACCT_SEG_ID
                rowGLTDIST2.Item("ACCT_SEG_CODE") = ACCT_SEG_CODE
                rowGLTDIST2.Item("ACCT_SEG_DESC") = row.Item("ACCT_SEG_DESC")
                dst.Tables("GLTDIST2").Rows.Add(rowGLTDIST2)
            End If
        Next

        optSegment.Enabled = True
        optSegment.Value = "4"
        'Set_Read_Only_for_ctl(optSegment, True)
        optSegment.Enabled = False
        optSegment.Visible = False

        chkShowNonZeroOnly.Checked = False
        Set_Segment()
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"GLTDIST2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdGLTDIST2.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdGLTDIST2}
            With grd.DisplayLayout.Override
                If EntryMode = "New" Or EntryMode = "Edit" Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No ' UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.False ' DefaultableBoolean.True
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                End If
            End With
        Next
    End Sub

#End Region

#Region "grdGLTDSTR2"

    Private Sub grdGLTDSTR2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTDIST2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            'Case "ACCT_CODE"
            '    grdCodeDesc(grdGLTDIST2, "GLTACCT1", "ACCT_CODE", "ACCT_DESC")
        End Select
    End Sub

    Private Sub grdGLTDSTR2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdGLTDIST2.AfterRowActivate

        'With grdGLTDIST2.DisplayLayout.Bands(0).Columns("ACCT_SEG_CODE")
        '    If grdGLTDIST2.ActiveRow.IsAddRow Then
        '        .CellActivation = UltraWinGrid.Activation.AllowEdit
        '    Else
        '        .CellActivation = UltraWinGrid.Activation.NoEdit
        '    End If
        'End With
    End Sub

    Private Sub grdGLTDSTR2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdGLTDIST2.AfterRowsDeleted

    End Sub

    Private Sub grdGLTDSTR2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdGLTDIST2.AfterRowUpdate

    End Sub

    Private Sub grdGLTDSTR2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdGLTDIST2.BeforeRowsDeleted
        'For Each grow As UltraWinGrid.UltraGridRow In e.Rows
        '    Dim SREP_CODE As String = grow.Cells("SREP_CODE").Value
        '    Dim CUST_CODE As String = grow.Cells("CUST_CODE").Value
        '    Dim rowGLTDSTR2 As DataRow = dst.Tables("GLTDSTR2").Rows.Find(New String() {SREP_CODE, CUST_CODE})
        '    If Not rowGLTDSTR2.RowState = DataRowState.Added Then
        '        MsgBox("Cannot Delete Previously Updated Colors", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
        '        e.Cancel = True
        '    End If
        'Next
    End Sub

    Private Sub grdGLTDSTR2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdGLTDIST2.BeforeRowUpdate

        'Dim row As DataRow = LookUp("GLTSEGM1", New String() {e.Row.Cells("ACCT_SEG_ID").Text, e.Row.Cells("ACCT_SEG_CODE").Text})
        'If row Is Nothing Then
        '    e.Cancel = True
        'End If

        'If e.Row.IsAddRow Then
        '    e.Row.Cells("DIST_APP_CODE").Value = Absx1.txtFor("DIST_APP_CODE").Text
        'End If

    End Sub

    Private Sub grdGLTDSTR2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTDIST2.ClickCellButton
        'Select Case e.Cell.Column.Key
        '    Case "ACCT_CODE"
        '        Dim sql_where As String = Get_List_of_Codes("GLTACCT1.ACCT_CODE not in", "GLTDSTR2", "ACCT_CODE")
        '        grdClickCellButton(grdGLTDIST2, sql_where, True)
        'End Select
    End Sub

#End Region

    Private Sub optSegment_ValueChanged(sender As Object, e As EventArgs) Handles optSegment.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Set_Segment()
        If optSegment.Value <> "4" And Not Me.IsDone Then
            MsgBox("Only Collections are Supported at this time", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            optSegment.Value = "4"
        End If
    End Sub

    Sub Set_Segment()
        Dim dvw As DataView = DirectCast(grdGLTDIST2.DataSource, DataTable).DefaultView
        Dim SQL As String = "ACCT_SEG_ID = '" & optSegment.Value & "'"
        If chkShowNonZeroOnly.Checked Then
            SQL &= " AND ISNULL(DIST_PCT,0) <> 0"
        End If
        dvw.RowFilter = SQL
        grdGLTDIST2.Text = "Distribution by " & optSegment.Text
        Sort_grdColumns(grdGLTDIST2, "ACCT_SEG_CODE")
    End Sub

    Private Sub chkShowNonZeroOnly_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowNonZeroOnly.CheckedChanged
        Set_Segment()
    End Sub
End Class