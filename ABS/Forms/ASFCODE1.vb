Public Class ASFCODE1
    Dim sql_Straight_List As String
    Dim sql_param_types As String
    Dim sql_params As Object()
    Dim tbl As DataTable
    Public FilterFirst As Boolean   ' Indicating whether ASFCODE1 should ask for a Filter information first before loading any data
    Dim wait_for_filter As Boolean
    Dim menuDelegate As MenuDel
    Dim menuName As String
    Dim COLUMN_NAME_returned As String
    Dim LIST_CODE As String = ""
    Dim use_where As Boolean = True

    Private Sub ASFCODE1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        ASCMAIN1.CodeSelector.Precedent_Keys.Clear()
        ASCMAIN1.CodeSelector.ParamTypes = Nothing
        ASCMAIN1.CodeSelector.Params = Nothing
        ' WHY IS THIS BEING DONE AT THIS POINT- MAKING IT IMPOSSIBLE FOR THE CALLING PROGRAM TO USE PRECEDENT KEYS
    End Sub

    Public Delegate Sub MenuDel(ByVal grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow)

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Call Load_Popup_Menu(grd, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
    End Sub

    Public Sub New(ByVal menuItem As String, ByVal menuFunc As MenuDel)
        InitializeComponent()

        menuName = menuItem
        menuDelegate = menuFunc
        Call Load_Popup_Menu(grd, "SSPB", "Show Filter", "Show GroupBox", menuName)
    End Sub


    Private Sub ASFCODE1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        frmASFBASE1 = ASCMAIN1.ActiveForm

        ASCMAIN1.CodeSelector.SelectedCode = ""
        ASCMAIN1.CodeSelector.Selections = 0
        ASCMAIN1.CodeSelector.SelectedCodes.Clear()
        ASCMAIN1.CodeSelector.SelectedRows.Clear()
        ASCMAIN1.CodeSelector.SelectedCodes0 = ""
        ASCMAIN1.CodeSelector.TABLE_NAME_temp = ""

        chkSelectedOnly.Text = "Selected Only"

        If ASCMAIN1.Running_in_VS Then ' If ASCMAIN1.USER_SECURITY_CODEs.Contains("AS") Then
            cmdSave.Visible = True
        End If

        ASCMAIN1.sql = ASCMAIN1.CodeSelector.SQL
        'If ASCMAIN1.CodeSelector.Custom_sql_where <> "" Then
        '    If use_where Then
        '        use_where = False
        '        ASCMAIN1.sql &= " where " & ASCMAIN1.CodeSelector.Custom_sql_where
        '    Else
        '        ASCMAIN1.sql &= " and " & ASCMAIN1.CodeSelector.Custom_sql_where
        '    End If
        'End If

        If ASCMAIN1.CodeSelector.MultipleSelections Then
            ASCMAIN1.sql = "Select X.*, 0 as SELECTED from (" & ASCMAIN1.sql & ") X"
        End If

        If ASCMAIN1.CodeSelector.VIEW_NAME <> "" Then

            If ASCMAIN1.CodeSelector.DoNotFilterFirst Then
            Else
                If ASCMAIN1.CodeSelector.tblASTVIEW1.Rows(0).Item("FILTER_FIRST") & "" = "1" _
                Or ASCMAIN1.CodeSelector.ForceFilterFirst Then
                    FilterFirst = True
                    wait_for_filter = True
                End If
            End If
        End If

        sql_Straight_List = ASCMAIN1.sql
        sql_param_types = If(ASCMAIN1.CodeSelector.ParamTypes, "")
        sql_params = ASCMAIN1.CodeSelector.Params

        ' only used with GLTSEGM1 (so far) - and now with SOTSVIA1.SHIP_PROD_CODE
        ' need to resolve whether to append "and" or "where"
        ' use_where = True
        If ASCMAIN1.CodeSelector.tblASTVIEW1.Rows.Count > 0 AndAlso ASCMAIN1.CodeSelector.tblASTVIEW1.Rows(0).Item("WHERE_CLAUSE") & "" <> "" Then
            use_where = False
        End If
        If ASCMAIN1.CodeSelector.Custom_sqlkey <> "" Or ASCMAIN1.CodeSelector.Custom_sql_where <> "" Then
            use_where = False
        End If

        With ASCMAIN1.CodeSelector
            If .Precedent_Keys.Count <> 0 Then
                For Each pkey As String In .Precedent_Keys.Keys
                    sql_Straight_List &= IIf(use_where, " where ", " and ") & pkey & " = '" & .Precedent_Keys(pkey) & "'"
                    use_where = False
                Next
            End If
        End With

        If FilterFirst Or wait_for_filter Then
            Show_Filter(grd, True)
        End If

        UltraOptionSet1.Items(0).DisplayText = "Straight List"
        UltraOptionSet1.CheckedIndex = 0

        'Call Setup_Tree(0)

        grd.DisplayLayout.Override.AllowAddNew = False
        grd.DisplayLayout.Override.AllowDelete = False
        If ASCMAIN1.CodeSelector.MultipleSelections Then
        Else
            grd.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True
            SplitContainer3.Panel2Collapsed = True
            SplitContainer2.SplitterDistance = SplitContainer2.SplitterDistance - cmdSelect.Height
            Panel1.Height = Panel1.Height - cmdSelect.Height
            cmdClearAll.Visible = False
            cmdCancel.Left = cmdClearAll.Left
        End If

        'grd.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.CellSelect
        If ASCMAIN1.CodeSelector.MultipleSelections Then
            If ASCMAIN1.USER_ID = "lmarinelli" Or ASCMAIN1.USER_ID = "ealomar" Then
                ' do not change CellClickAction for Lauren
            Else
                grd.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.RowSelect
            End If
        End If


        If ASCMAIN1.CodeSelector.VIEW_NAME <> "" Then
            If ASCMAIN1.CodeSelector.Hierarchal_Views.Count = 0 Or FilterFirst Then
                SplitContainer2.Panel2Collapsed = True
                Panel1.Height = Panel1.Height - UltraGroupBox1.Height
            Else
                For I As Integer = 1 To ASCMAIN1.CodeSelector.Hierarchal_Views.Count
                    UltraOptionSet1.Items.Add(I, ASCMAIN1.CodeSelector.Hierarchal_Views.Item(I).ToString)
                Next
            End If
        End If

        Dim W As Long = 0
        With grd.DisplayLayout.Bands(0)
            For I As Integer = 0 To .Columns.Count - 1
                W = W + .Columns(I).Width
                If .Columns(I).Key <> "SELECTED" Then
                    .Columns(I).CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next I
        End With
        W = W + 50

        If Me.Width < W + 10 Then
            Me.Width = W + 10
        End If

        Call ASCMAIN1.Center(Me)

        If ASCMAIN1.CodeSelector.VIEW_NAME = "" Then
            cmdLoadData.Visible = False
        Else
            cmdLoadData.Visible = ASCMAIN1.CodeSelector.ForceFilterFirst Or ASCMAIN1.CodeSelector.tblASTVIEW1.Rows(0).Item("FILTER_FIRST") & "" = "1"
        End If

        If ASCMAIN1.CodeSelector.Caption <> "" Then
            Me.Text = ASCMAIN1.CodeSelector.Caption
        Else
            If ASCMAIN1.CodeSelector.VIEW_NAME <> "" Then
                If ASCMAIN1.CodeSelector.VIEW_DESC <> "" Then
                    Me.Text = ASCMAIN1.CodeSelector.VIEW_DESC
                End If
            End If
        End If



        COLUMN_NAME_returned = ""
        If ASCMAIN1.CodeSelector.VIEW_NAME = "" Then
            COLUMN_NAME_returned = tbl.Columns(0).ColumnName
        Else
            COLUMN_NAME_returned = _
                ASCMAIN1.CodeSelector.tblASTVIEW1.Rows(0).Item("COLUMN_NAME") & ""
            If COLUMN_NAME_returned = "T_CODE" And tbl.Columns(0).ColumnName <> "T_CODE" Then
                COLUMN_NAME_returned = tbl.Columns(0).ColumnName
                ASCMAIN1.CodeSelector.tblASTVIEW1.Rows(0).Item("COLUMN_NAME") = COLUMN_NAME_returned
            End If
            If COLUMN_NAME_returned = "" Then
                COLUMN_NAME_returned = tbl.Columns(0).ColumnName
            End If
            If Not tbl.Columns.Contains(COLUMN_NAME_returned) Then
                COLUMN_NAME_returned = tbl.Columns(0).ColumnName
            End If
        End If

        grd.ActiveRow = Nothing
        Dim z As String = ASCMAIN1.CodeSelector.PreviouslySelectedCodes0
        ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = "" ' this was necessary to resolve the infamous roger bug, where line 192 would freak if there were non numeric values in the list where it is comparing to PO_ORDER_LNO
        If z = Chr(0) Then
            z = ""
        End If
        If Mid$(z, 1, 1) = Chr(0) Then
            z = Mid$(z, 2)
        End If
        If z <> "" Then
            If ASCMAIN1.CodeSelector.MultipleSelections Then
                For Each dr As DataRow In tbl.Select(tbl.Columns(COLUMN_NAME_returned).ColumnName & " IN ('" & Replace(z, Chr(0), "','") & "')")
                    dr.Item("SELECTED") = True
                Next
                Call Count_Selections("C")
            Else
                Dim sqlwhere As String = tbl.Columns(0).ColumnName & " = '" & z & "'"
                Dim rws As DataRow() = tbl.Select(sqlwhere)
                'Dim rws As DataRow() = tbl.Select(ASCMAIN1.CodeSelector.VIEW_NAME & " = '" & z & "'")

                If UBound(rws) = 0 Then
                    grd.Selected.Rows.Add(grd.Rows.GetRowWithListIndex(tbl.Rows.IndexOf(rws(0))))
                    grd.ActiveRow = grd.Selected.Rows(0)
                Else
                    If FilterFirst Then
                        grd.DisplayLayout.Bands(0).ColumnFilters(tbl.Columns(0).ColumnName).FilterConditions.Add _
                            (UltraWinGrid.FilterComparisionOperator.StartsWith, z)
                    End If
                End If
            End If
        End If

        'Call Load_Popup_Menu(grd, "SSBB", "Show Filter", "Show GroupBox")

        If ASCMAIN1.CodeSelector.PreFilter.Count <> 0 Then
            For Each COLUMN_NAME As String In ASCMAIN1.CodeSelector.PreFilter.Keys
                Dim FILTER_VALUE As Object = ASCMAIN1.CodeSelector.PreFilter(COLUMN_NAME)
                grd.DisplayLayout.Bands(0).ColumnFilters(COLUMN_NAME).FilterConditions.Add(UltraWinGrid.FilterComparisionOperator.StartsWith, FILTER_VALUE)
            Next
        End If

        If Not ASCMAIN1.CodeSelector.MultipleSelections Then
            splHeader.Panel2Collapsed = True
        Else
            With dst
                ASCMAIN1.sql = "Select * from ASTLIST1 where COLUMN_NAME = :PARM1"
                Create_TDA(.Tables.Add, "ASTLIST1", "**", 0, True, "V", 1)

                ASCMAIN1.sql = "Select * from ASTLIST2 where LIST_CODE in (Select LIST_CODE from ASTLIST1 where COLUMN_NAME = :PARM1)"
                Create_TDA(.Tables.Add, "ASTLIST2", "**", 0, True, "V", 2)
            End With

            If EntryMode = "S" Then
                SplitContainer2.Visible = False
                splHeader.Panel1Collapsed = True
                chkSelectedOnly.Checked = True
                Me.ControlBox = True
            Else
                Fill_Records("ASTLIST1", ASCMAIN1.CodeSelector.VIEW_NAME)

                cbeLIST.DisplayMember = "LIST_DESC"
                cbeLIST.ValueMember = "LIST_CODE"
                'cbeLIST.DataMember = "ASTLIST1"
                cbeLIST.DataSource = dst.Tables("ASTLIST1")
            End If
        End If

        txtFind.Focus()
    End Sub

#Region "Popup Menus"

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        'Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case menuName
                menuDelegate(grd.ActiveRow)

        End Select
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        'Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
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
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

            End Select

        End If
    End Sub

#End Region

    Private Sub grd_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grd.AfterRowUpdate
        Count_Selections("C")
    End Sub

    Private Sub cmdClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearAll.Click

        Clear_Selections()

        cbeLIST.SelectedIndex = -1
        cbeLIST.Value = DBNull.Value
        LIST_CODE = ""
        chkSelectedOnly.Checked = False
    End Sub

    Sub Clear_Selections()
        If UltraOptionSet1.CheckedIndex = 0 Then
            For Each grow As Infragistics.Win.UltraWinGrid.UltraGridRow In grd.Rows
                If grow.IsFilteredOut Then
                Else
                    grow.Cells("SELECTED").Value = "0"
                    grow.Update()
                End If
            Next

            'For Each dr As DataRow In tbl.Select("SELECTED = 1")
            '    dr.Item("SELECTED") = "0"
            'Next
        Else
            For Each tbl In dst.Tables
                For Each dr As DataRow In tbl.Select("SELECTED = 1")
                    dr.Item("SELECTED") = "0"
                Next
            Next
        End If
        Count_Selections("C", 0)
    End Sub

    Private Sub cmdSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelectAll.Click
        tbl.Columns("SELECTED").ReadOnly = False
        If UltraOptionSet1.CheckedIndex = 0 Then
            For Each grow As Infragistics.Win.UltraWinGrid.UltraGridRow In grd.Rows
                If grow.IsFilteredOut Then
                Else
                    grow.Cells("SELECTED").Value = "1"
                    grow.Update()
                End If
            Next

            'For Each dr As DataRow In tbl.Rows
            '    dr.Item("SELECTED") = 1
            'Next
        Else
            For Each tbl In dst.Tables
                For Each dr As DataRow In tbl.Rows
                    dr.Item("SELECTED") = "1"
                Next
            Next
        End If

        Count_Selections("C")
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        ASCMAIN1.CodeSelector.Selections = 0
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        cmdOK_Click()
    End Sub

    Sub cmdOK_Click()

        If ASCMAIN1.CodeSelector.MultipleSelections Then
            If ASCMAIN1.CodeSelector.Selections = 0 Then
                Exit Sub
            End If
            If ASCMAIN1.CodeSelector.Selections > 1000 Then
                MsgBox("Too Many Selections (" & CStr(ASCMAIN1.CodeSelector.Selections) & ").  Maximum permitted is 1000")
                Exit Sub
            End If
            Dim VIEW_NO As Integer = UltraOptionSet1.CheckedItem.ValueList.SelectedIndex
            If VIEW_NO <> 0 Then
                Dim TABLE_NAME As String = ""
                For Each dr As DataRow In ASCMAIN1.CodeSelector.Hierarchal_Views_dr(VIEW_NO)
                    TABLE_NAME = dr.Item("HIERARCHAL_TABLE_NAME")
                Next
                For Each dr As DataRow In dst.Tables(TABLE_NAME).Select("SELECTED = 1")
                    ASCMAIN1.CodeSelector.SelectedCodes.Add(dr.Item(0))
                    ASCMAIN1.CodeSelector.SelectedCodes0 = ASCMAIN1.CodeSelector.SelectedCodes0 & Chr(0) & dr.Item(COLUMN_NAME_RETURNED) ' dr.Item(0)
                    ASCMAIN1.CodeSelector.SelectedRows.Add(dr)
                Next
            Else
                For Each dr As DataRow In tbl.Select("SELECTED = 1")
                    ASCMAIN1.CodeSelector.SelectedCodes.Add(dr.Item(COLUMN_NAME_RETURNED))
                    ASCMAIN1.CodeSelector.SelectedCodes0 = ASCMAIN1.CodeSelector.SelectedCodes0 & Chr(0) & dr.Item(COLUMN_NAME_RETURNED) ' dr.Item(0)
                    ASCMAIN1.CodeSelector.SelectedRows.Add(dr)
                Next
            End If
        Else
            If grd.Selected.Rows.Count = 0 Then
                If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                    grd.ActiveRow.Selected = True
                End If
            End If
            If grd.Selected.Rows.Count = 0 Then
                MsgBox("Nothing Selected")
                Exit Sub
            End If
            If grd.ActiveRow.ListIndex = -1 Then
                Exit Sub
            End If
            If grd.Selected.Rows(0).Band.Index <> grd.DisplayLayout.Bands.Count - 1 Then
                Exit Sub
            End If
            If COLUMN_NAME_returned = "" Or COLUMN_NAME_returned = "T_CODE" Then
                ASCMAIN1.CodeSelector.SelectedCode = grd.Selected.Rows(0).Cells(0).Text
            Else
                ASCMAIN1.CodeSelector.SelectedCode = grd.Selected.Rows(0).Cells(COLUMN_NAME_returned).Text
            End If

            ASCMAIN1.CodeSelector.Selections = 1

            ASCMAIN1.CodeSelector.SelectedCodes.Add(ASCMAIN1.CodeSelector.SelectedCode)
            ASCMAIN1.CodeSelector.SelectedCodes0 &= Chr(0) & ASCMAIN1.CodeSelector.SelectedCode
            Dim DR As DataRow = tbl.Rows(grd.ActiveRow.ListIndex)
            ASCMAIN1.CodeSelector.SelectedRows.Add(DR)
        End If

        Me.Close()
    End Sub

    Private Sub grd_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grd.MouseUp
        '        Try

        If grd.ActiveRow Is Nothing Then
            Exit Sub
        End If

        If ASCMAIN1.CodeSelector.MultipleSelections Then

            Dim aUIElement As Infragistics.Win.UIElement
            aUIElement = grd.DisplayLayout.UIElement.ElementFromPoint(New Point(e.X, e.Y))
            If aUIElement Is Nothing Then
                Exit Sub
            End If
            ' declare and retrieve a reference to the Cell
            Dim aCell As Infragistics.Win.UltraWinGrid.UltraGridCell
            aCell = aUIElement.GetContext( _
            GetType(Infragistics.Win.UltraWinGrid.UltraGridCell))
            ' if a cell was found display the band and column key
            If aCell Is Nothing Then
                Exit Sub
            End If

            'If e.Location.X <= grd.ActiveRow.Cells("SELECTED").Column.Width Then
            If aCell.Column.Key = "SELECTED" Then
                'grd.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True

                If Val(grd.ActiveRow.Cells("SELECTED").Value & "") = 0 Then
                    grd.ActiveRow.Cells("SELECTED").Value = "1"
                    'aCell.Value = 1
                Else
                    grd.ActiveRow.Cells("SELECTED").Value = "0"
                End If
                grd.UpdateData()

                'If UltraOptionSet1.CheckedItem.ValueList.SelectedIndex <> 0 Then
                If UltraOptionSet1.CheckedIndex <> 0 Then
                    If Not Nothing Is grd.ActiveRow.ChildBands Then
                        Me.Cursor = Cursors.WaitCursor
                        Call Set_Child_Rows(grd.ActiveRow.ChildBands(0).Rows, grd.ActiveRow.Cells("SELECTED").Value)
                        Call Count_Selections("C")
                        Me.Cursor = Cursors.Default
                    End If
                End If
            End If
        End If
        '       Catch ex As Exception

        '        End Try

    End Sub

    Sub Set_Child_Rows(ByVal grd_rows As Infragistics.Win.UltraWinGrid.RowsCollection, ByVal V As Integer)
        For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In grd_rows
            row.Cells("SELECTED").Value = V
            grd.UpdateData()
            If Not Nothing Is row.ChildBands Then
                Call Set_Child_Rows(row.ChildBands(0).Rows, V)
            End If
        Next
    End Sub
    Sub Count_Selections(ByVal Increment_Set_Calculate As String, Optional ByVal S As Integer = 0)
        If ASCMAIN1.CodeSelector.MultipleSelections = False Then
            Exit Sub
        End If

        'If UltraOptionSet1.CheckedIndex = 0 Then
        Select Case Increment_Set_Calculate
            Case "I"
                ASCMAIN1.CodeSelector.Selections = ASCMAIN1.CodeSelector.Selections + S
            Case "S"
                ASCMAIN1.CodeSelector.Selections = S
            Case "C"
                If UltraOptionSet1.CheckedIndex = 0 Then
                    'If UltraOptionSet1.CheckedItem.ValueList.SelectedIndex = 0 Then
                    ASCMAIN1.CodeSelector.Selections = Convert.ToInt32(tbl.Compute("Count(SELECTED)", "SELECTED = 1"))
                Else
                    ASCMAIN1.CodeSelector.Selections = Convert.ToInt32(dst.Tables(ASCMAIN1.CodeSelector.TABLE_NAME).Compute("Count(SELECTED)", "SELECTED = 1"))
                End If
        End Select
        'End If
        chkSelectedOnly.Text = "Selected Only (" & ASCMAIN1.CodeSelector.Selections & ")"
    End Sub

    Private Sub cmdOK_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles cmdOK.MouseDown
    End Sub

    Private Sub chkSelectedOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSelectedOnly.CheckedChanged

        Dim dvw As DataView = DirectCast(grd.DataSource, DataTable).DefaultView
        If chkSelectedOnly.Checked Then
            dvw.RowFilter = "SELECTED = '1'"
        Else
            dvw.RowFilter = ""
        End If

        'If chkSelectedOnly.CheckState Then
        '    grd.Rows.ColumnFilters("SELECTED").FilterConditions.Clear()
        '    grd.Rows.ColumnFilters("SELECTED").FilterConditions.Add(Infragistics.Win.UltraWinGrid.FilterComparisionOperator.Match, 1)
        'Else
        '    grd.Rows.ColumnFilters("SELECTED").FilterConditions.Clear()
        '    grd.Rows.Refresh(Infragistics.Win.UltraWinGrid.RefreshRow.ReloadData)
        'End If
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Dim sql As String
        If MsgBox("Save Grid Column Widths", vbYesNo + vbQuestion, "Option to Save Designed Column Widths") = vbYes Then
            sql = "Select * from ASTVIEW2 where VIEW_NAME = '" & ASCMAIN1.CodeSelector.VIEW_NAME & "'"
            sql = sql & " and TABLE_NAME = '" & ASCMAIN1.CodeSelector.TABLE_NAME & "'"

            Dim dt As New DataTable
            With ASCDATA1.GetDataAdapter(dt, "ASTVIEW2", sql, True)
                For Each dr As DataRow In dt.Rows
                    dr.Item("COLUMN_WIDTH") = grd.DisplayLayout.Bands(0).Columns(dr.Item("COLUMN_NAME")).Width
                Next
                .Update(dt)
                .Dispose()
            End With

            Call ASCMAIN1.Load_Views()
        End If
    End Sub

    Private Sub cmdSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelect.Click
        For Each r As Infragistics.Win.UltraWinGrid.UltraGridRow In grd.Selected.Rows
            If Val(r.Cells("SELECTED").Value & "") <> 1 Then
                r.Cells("SELECTED").Value = 1
            End If
            grd.UpdateData()
        Next
    End Sub

    Private Sub txtFind_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFind.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            If Not ASCMAIN1.CodeSelector.MultipleSelections Then
                If grd.ActiveRow Is Nothing Then
                    Exit Sub
                Else
                    grd.Selected.Rows.Clear()
                    grd.Selected.Rows.Add(grd.ActiveRow)
                End If
                cmdOK_Click()
            End If
        End If
    End Sub

    Private Sub txtFind_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFind.KeyPress
    End Sub

    Private Sub txtFind_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFind.ValueChanged

        Dim COLUMN_NAME As String = Determine_Sort()

        ' Do not want to blow out the sceen if the user presses an invalid character, for example [
        Try
            If COLUMN_NAME <> "" Then
                Dim gridRow As Infragistics.Win.UltraWinGrid.UltraGridRow
                For Each gridRow In grd.Rows.GetFilteredInNonGroupByRows
                    If gridRow.Cells(COLUMN_NAME).Value.ToString.ToUpper Like txtFind.Text.ToUpper & "*" Then
                        gridRow.Activate()
                        grd.ActiveRowScrollRegion.FirstRow = gridRow
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            ' Nothing 
        End Try

    End Sub

    Private Sub UltraOptionSet1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraOptionSet1.ValueChanged
        Call Setup_Tree(UltraOptionSet1.CheckedItem.ValueList.SelectedIndex)
        Call Count_Selections("C")
    End Sub

    Sub Setup_Tree(ByVal VIEW_NO As Integer)
        Select Case VIEW_NO
            Case 0
                SplitContainer1.Panel1Collapsed = False

                Dim no_data As Boolean = False
                If ASCMAIN1.CodeSelector.UseDataFromTable IsNot Nothing Then
                    no_data = True
                End If

                ' if wait for filter, yet we have previously selected codes, we need to ignore wait for filter
                If wait_for_filter And ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 <> "" Then
                    wait_for_filter = False
                End If

                tbl = ASCDATA1.GetDataTable(sql_Straight_List, , , Not wait_for_filter And Not no_data, , sql_param_types, sql_params)

                If ASCMAIN1.CodeSelector.UseDataFromTable IsNot Nothing Then
                    'tbl.Rows.Clear()
                    For Each row As DataRow In ASCMAIN1.CodeSelector.UseDataFromTable.Select
                        Dim row2 As DataRow = tbl.NewRow
                        If ASCMAIN1.CodeSelector.MultipleSelections Then
                            row2.Item("SELECTED") = 0
                        End If
                        For Each dcol As DataColumn In tbl.Columns
                            If row.Table.Columns.Contains(dcol.ColumnName) Then
                                row2.Item(dcol.ColumnName) = row.Item(dcol.ColumnName)
                            End If
                        Next
                        tbl.Rows.Add(row2)
                    Next
                    tbl.AcceptChanges()
                    ASCMAIN1.CodeSelector.UseDataFromTable = Nothing
                End If


                If ASCMAIN1.CodeSelector.MultipleSelections Then
                    tbl.Columns("SELECTED").ReadOnly = False
                End If

                If ASCMAIN1.CodeSelector.MultipleSelections Then
                    If tbl.Columns.Contains("SELECTED") Then
                        tbl.Columns("SELECTED").DefaultValue = "0"
                    End If
                End If

                grd.DataSource = tbl
                cmdOK.Enabled = Not (tbl.Rows.Count = 0)

                Try
                    ASCMAIN1.grdInitializeLayout(grd)
                    Create_Summary(grd, grd.DisplayLayout.Bands(0).Columns(0).Key, "Count")
                Catch ex As Exception

                End Try


                grd.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ExtendLastColumn
                grd.DisplayLayout.Bands(0).SortedColumns.Clear()

                'Dim ORDER_BY As String = ASCMAIN1.CodeSelector.tblASTVIEW1.Rows(0).Item("ORDER_BY") & ""
                Dim ORDER_BY As String = String.Empty

                If ASCMAIN1.CodeSelector.tblASTVIEW1 IsNot Nothing AndAlso ASCMAIN1.CodeSelector.tblASTVIEW1.Rows.Count > 0 Then
                    ORDER_BY = ASCMAIN1.CodeSelector.tblASTVIEW1.Rows(0).Item("ORDER_BY") & ""
                End If

                If ORDER_BY = "" Then
                    grd.DisplayLayout.Bands(0).SortedColumns.Add(grd.DisplayLayout.Bands(0).Columns(0).Key, False)
                Else
                    If ASCMAIN1.CodeSelector.tblASTVIEW1.Rows(0).Item("ORDER_BY_DESC") & "" = "1" Then
                        ORDER_BY = ORDER_BY.ToLower
                    End If
                End If
                Sort_grdColumns(grd, ORDER_BY)



                Dim COLUMN_NAME_RETURNED As String = ""
                If ASCMAIN1.CodeSelector.VIEW_NAME = "" Then
                    COLUMN_NAME_RETURNED = tbl.Columns(0).ColumnName
                Else

                    COLUMN_NAME_RETURNED = _
                        ASCMAIN1.CodeSelector.tblASTVIEW1.Rows(0).Item("COLUMN_NAME") & ""
                    If COLUMN_NAME_RETURNED = "T_CODE" Then
                        If grd.DisplayLayout.Bands(0).Columns(0).Key <> COLUMN_NAME_RETURNED Then
                            COLUMN_NAME_RETURNED = grd.DisplayLayout.Bands(0).Columns(0).Key
                        End If
                    End If
                    ' ALL LINES BELOW REMMED OUT WJZ 02/07/11 - DON'T SEE THE POINT IF JUST A FEW LINES LOWER WE ARE CALLING SORT_GRDCOLUMNS
                    'If COLUMN_NAME_RETURNED = "" Or COLUMN_NAME_RETURNED = "T_CODE" Then
                    'Else
                    '    If grd.DisplayLayout.Bands(0).Columns(0).Key <> COLUMN_NAME_RETURNED And _
                    '        (grd.DisplayLayout.Bands(0).Columns.Count > 1 AndAlso _
                    '       grd.DisplayLayout.Bands(0).Columns(1).Key = COLUMN_NAME_RETURNED) Then
                    '        grd.DisplayLayout.Bands(0).SortedColumns.Add(grd.DisplayLayout.Bands(0).Columns(1), False)
                    '    End If
                    'End If
                End If
                If ORDER_BY = "" Then
                    Sort_grdColumns(grd, COLUMN_NAME_RETURNED)
                End If

                Set_Screen_Filter()

                If ASCMAIN1.CodeSelector.MultipleSelections Then
                    tbl.Columns("SELECTED").ReadOnly = False
                    grd.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True

                    Call Set_Selected()
                    grd.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
                    'grd.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
                    grd.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended
                Else
                    grd.DisplayLayout.Override.AllowUpdate = False
                    'grd.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
                    grd.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single
                End If

                If ASCMAIN1.CodeSelector.VIEW_NAME = "" Then
                    For I As Integer = 0 To grd.DisplayLayout.Bands(0).Columns.Count - 1
                        grd.DisplayLayout.Bands(0).Columns(I).Header.Caption = ASCMAIN1.Make_Caption(grd.DisplayLayout.Bands(0).Columns(I).Header.Caption)
                        'grd.DisplayLayout.Bands(0).Columns(I - 1).Width = ASCMAIN1.CodeSelector.grdColumns(I - 1).Item("COLUMN_WIDTH")
                    Next

                Else

                    If ASCMAIN1.CodeSelector.grdColumns.Count <> 0 Then
                        For I As Integer = 1 To ASCMAIN1.CodeSelector.grdColumns.Count
                            grd.DisplayLayout.Bands(0).Columns(I - 1).Header.Caption = ASCMAIN1.CodeSelector.grdColumns(I - 1).Item("COLUMN_CAPTION")
                            If Val(ASCMAIN1.CodeSelector.grdColumns(I - 1).Item("COLUMN_WIDTH") & "") <> 0 Then
                                grd.DisplayLayout.Bands(0).Columns(I - 1).Width = ASCMAIN1.CodeSelector.grdColumns(I - 1).Item("COLUMN_WIDTH")
                            End If
                            If ASCMAIN1.CodeSelector.grdColumns(I - 1).Item("COLUMN_HIDDEN") & "" = "1" Then
                                grd.DisplayLayout.Bands(0).Columns(I - 1).Hidden = True
                            End If
                            If ASCMAIN1.CodeSelector.grdColumns(I - 1).Item("COLUMN_CHECKED") & "" = "1" Then
                                grd.DisplayLayout.Bands(0).Columns(I - 1).Style = UltraWinGrid.ColumnStyle.CheckBox
                                grd.DisplayLayout.Bands(0).Columns(I - 1).Header.Appearance.TextHAlign = HAlign.Center
                                grd.DisplayLayout.Bands(0).Columns(I - 1).Editor.DataFilter = New CheckEditorDataFilter
                            End If

                            Dim VL As ValueList = Nothing
                            If ASCMAIN1.CodeSelector.grdColumns(I - 1).Item("VALUE_LIST") & "" = "1" Then
                                VL = ASCMAIN1.Add_Value_List(grd, ASCMAIN1.CodeSelector.grdColumns(I - 1).Item("COLUMN_NAME"))
                            End If

                            Dim DEFAULT_VALUE As String = ASCMAIN1.CodeSelector.grdColumns(I - 1).Item("DEFAULT_VALUE") & ""
                            If DEFAULT_VALUE <> "" Then
                                Show_Filter(grd, True)

                                If VL IsNot Nothing AndAlso VL.ValueListItems.Count > 0 Then
                                    For Each vli As ValueListItem In VL.ValueListItems
                                        If vli.DataValue = DEFAULT_VALUE Then
                                            DEFAULT_VALUE = vli.DisplayText
                                            Exit For
                                        End If
                                    Next
                                End If

                                grd.DisplayLayout.Bands(0).ColumnFilters(I - 1).FilterConditions.Add _
                                    (UltraWinGrid.FilterComparisionOperator.Equals, DEFAULT_VALUE)
                            End If

                        Next
                    End If
                End If

            Case Else
                SplitContainer1.Panel1Collapsed = True

                dst = New DataSet
                grd.DataSource = Nothing
                'grd.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.None
                ' cannot do extendlast column once record_count and selected_count is working
                grd.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ExtendLastColumn


                Dim i As Integer = 0
                Dim TABLE_NAME_prev As String = ""
                Dim COLUMN_NAME_LINK_prev As String = ""

                Dim L As List(Of DataRow) = ASCMAIN1.CodeSelector.Hierarchal_Views_dr(VIEW_NO)

                Dim BT(,) As String
                ReDim BT(2, 0)  ' 2nd dimension is Band, 1st dimension 0 = TABLE_NAME, 1 = COLUMN_NAME_CODE, 2 = COLUMN_NAME_DESC

                For Each dr As DataRow In L
                    If i > 0 Then
                        ReDim Preserve BT(2, i)
                    End If
                    BT(0, i) = dr.Item("HIERARCHAL_TABLE_NAME")
                    BT(1, i) = dr.Item("COLUMN_NAME_CODE")
                    BT(2, i) = dr.Item("COLUMN_NAME_DESC")

                    i = i + 1
                    Dim sql As String
                    sql = "Select " & dr.Item("COLUMN_NAME_CODE") & "," & dr.Item("COLUMN_NAME_DESC")
                    If i <> 1 Then
                        sql = sql & ", " & COLUMN_NAME_LINK_prev
                    End If
                    If i <> L.Count Then
                        sql = sql & ", 0 RECORD_COUNT, 0 SELECTED_COUNT"
                    End If

                    sql = sql & " FROM " & dr.Item("HIERARCHAL_TABLE_NAME")
                    If ASCMAIN1.CodeSelector.MultipleSelections Then
                        sql = "Select X.*, '0' as SELECTED from (" & sql & ")  X"
                    End If

                    dst.Tables.Add(ASCDATA1.GetDataTable(sql, dr.Item("HIERARCHAL_TABLE_NAME")))
                    If ASCMAIN1.CodeSelector.MultipleSelections Then
                        dst.Tables(dr.Item("HIERARCHAL_TABLE_NAME")).Columns("SELECTED").ReadOnly = False
                    End If

                    If i <> 1 Then
                        With dst
                            .Relations.Add(dr.Item("HIERARCHAL_TABLE_NAME") & "_" & TABLE_NAME_prev, _
                            .Tables(TABLE_NAME_prev).Columns(COLUMN_NAME_LINK_prev), _
                            .Tables(dr.Item("HIERARCHAL_TABLE_NAME")).Columns(COLUMN_NAME_LINK_prev))
                        End With
                    End If
                    TABLE_NAME_prev = dr.Item("HIERARCHAL_TABLE_NAME")
                    COLUMN_NAME_LINK_prev = dr.Item("COLUMN_NAME_LINK")
                Next
                grd.DataSource = dst

                Try
                    ASCMAIN1.grdInitializeLayout(grd)
                    Create_Summary(grd, grd.DisplayLayout.Bands(0).Columns(0).Key, "Count")
                Catch ex As Exception

                End Try


                Call Set_Selected()

                i = grd.DisplayLayout.Bands.Count
                grd.DisplayLayout.Bands(i - 1).Columns(0).Width = 100
                grd.DisplayLayout.Bands(i - 1).Columns(1).Width = 300

                For Each b As UltraWinGrid.UltraGridBand In grd.DisplayLayout.Bands

                    For i = 1 To 2
                        Dim rowASTVIEW2() As DataRow = ASCMAIN1.dstASTVIEWS.Tables("ASTVIEW2").Select("VIEW_NAME = '" & BT(1, b.Index) & "' and TABLE_NAME = '" & BT(0, b.Index) & "' and COLUMN_NAME = '" & BT(i, b.Index) & "'")
                        If rowASTVIEW2.GetLength(0) = 1 Then
                            b.Columns(BT(i, b.Index)).Header.Caption = rowASTVIEW2(0).Item("COLUMN_CAPTION")
                        End If
                    Next

                    For i = 0 To b.Columns.Count - 1
                        If b.Columns(i).Key <> "SELECTED" And b.Columns(i).Key <> "RECORD_COUNT" And b.Columns(i).Key <> "SELECTED_COUNT" Then
                            If i > 1 Then
                                b.Columns(i).Hidden = True
                            Else
                                If b.Columns(i).Header.Caption <> "" Then
                                    'b.Columns(i).Header.Caption = ASCMAIN1.Make_Caption(b.Columns(i).Header.Caption)

                                    b.Columns(i).Header.Caption = ASCMAIN1.Make_Caption(b.Columns(i).Key)
                                    If b.Columns(i).Key Like "*_CODE" Then
                                        b.Columns(i).Header.Caption = ASCMAIN1.Make_Caption(Mid$(b.Columns(i).Key, 1, Len(b.Columns(i).Key) - 5))
                                    End If
                                    If b.Columns(i).Key Like "*_DESC" Then
                                        b.Columns(i).Header.Caption = "Description"
                                    End If
                                    If b.Columns(i).Key Like "*_NAME" Then
                                        b.Columns(i).Header.Caption = "Name"
                                    End If
                                End If
                            End If
                        End If
                    Next
                Next
        End Select

    End Sub

    Sub Set_Selected()
        For Each b As Infragistics.Win.UltraWinGrid.UltraGridBand In grd.DisplayLayout.Bands
            If ASCMAIN1.CodeSelector.MultipleSelections Then
                b.Columns("SELECTED").Header.VisiblePosition = 0
                b.Columns("SELECTED").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
                b.Columns("SELECTED").Header.Caption = ""
                b.Columns("SELECTED").Width = 30
                b.Columns("SELECTED").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
                b.Columns("SELECTED").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
            End If

            If b.Index < grd.DisplayLayout.Bands.Count - 1 Then
                b.Columns("RECORD_COUNT").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
                b.Columns("RECORD_COUNT").Header.Caption = "Records"
                b.Columns("RECORD_COUNT").Width = 80
                b.Columns("RECORD_COUNT").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                b.Columns("RECORD_COUNT").Format = "###,##0"
                b.Columns("RECORD_COUNT").Hidden = True

                b.Columns("SELECTED_COUNT").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Right
                b.Columns("SELECTED_COUNT").Header.Caption = "Selected"
                b.Columns("SELECTED_COUNT").Width = 80
                b.Columns("SELECTED_COUNT").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                b.Columns("SELECTED_COUNT").Format = "###,##0"
                b.Columns("SELECTED_COUNT").Hidden = True
            End If

        Next
    End Sub

    Private Sub cmdLoadData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLoadData.Click

        Me.Cursor = Cursors.WaitCursor

        Dim sqlwhere As String = ASCMAIN1.Get_where_from_Filter(grd)

        If FilterFirst And sqlwhere = "" Then
            MsgBox("You Must Provide a Filter before Loading Rows")
            Exit Sub
        End If

        wait_for_filter = False
        Call Set_Screen_Filter()
        grd.DisplayLayout.Override.FilterOperatorLocation = UltraWinGrid.FilterOperatorLocation.AboveOperand

        Dim sql As String = ""
        If InStr(sql_Straight_List.ToUpper, " WHERE ") = 0 And InStr(sql_Straight_List.ToUpper, " LEFT JOIN ") = 0 Then
            sql = sql_Straight_List & ASCMAIN1.SQL_Add_WHERE(sqlwhere & ASCMAIN1.CodeSelector.Custom_sql_where)
        Else
            sql = "Select * from (" & sql_Straight_List & ") " & ASCMAIN1.SQL_Add_WHERE(sqlwhere & ASCMAIN1.CodeSelector.Custom_sql_where)
        End If
        tbl = ASCDATA1.GetDataTable(sql, , , Not wait_for_filter, , sql_param_types, sql_params)
        If ASCMAIN1.CodeSelector.MultipleSelections Then
            tbl.Columns("SELECTED").ReadOnly = False
        End If
        grd.DataSource = tbl
        If grd.Rows.Count <> 0 Then
            grd.ActiveRow = grd.Rows(0)
        End If
        cmdOK.Enabled = Not (tbl.Rows.Count = 0)

        Me.Cursor = Cursors.Default
    End Sub

    Sub Set_Screen_Filter()
        txtFind.Visible = Not wait_for_filter
        'lblFind.Visible = Not wait_for_filter
    End Sub

    Private Sub grd_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grd.DoubleClickRow
        If Not ASCMAIN1.CodeSelector.MultipleSelections Then
            If grd.Selected.Rows.Count = 0 Then
                If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                    grd.ActiveRow.Selected = True
                End If
            End If

            If grd.ActiveRow Is Nothing Or grd.Selected.Rows.Count = 0 Then
                Exit Sub
            End If
            cmdOK_Click()
        End If
    End Sub

    Private Sub grd_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grd.KeyDown

    End Sub

    Private Sub grd_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grd.KeyUp
        Try
            If e.KeyCode = Keys.Enter Then
                If grd.Rows.FilterRow.IsActiveRow Then
                    Dim cmd_e As New System.EventArgs
                    Call cmdLoadData_Click(cmdLoadData, cmd_e)
                End If
            End If

        Catch ex As Exception
            ' Nothing at this point
        End Try

    End Sub

    Private Sub cmdSaveList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveList.Click

        If ASCMAIN1.CodeSelector.Selections = 0 Then
            MsgBox("No Code Values were Selected; Cannot Save", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        If cbeLIST.Text = "" Then
            MsgBox("List Description is Mandatory; Cannot Save", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        Dim sqlDelete As String = ""

        If LIST_CODE <> "" Then
            Using frmASFMSGBF As New ASFMSGBF
                Dim i As Integer = frmASFMSGBF.Get_opt_from_User _
                ("Do you want to replace the list, or create a new one?", _
                 New String() {"Replace the Existing List", "Create a New List"}, _
                 1, _
                 "You are working with an Existing Value List")
                If i = -1 Then
                    Exit Sub
                ElseIf i = 0 Then ' Replace
                    sqlDelete = "LIST_CODE = '" & LIST_CODE & "'"
                Else ' New
                    LIST_CODE = ""
                End If
            End Using
        End If

        Dim rowASTLIST1 As DataRow

        If LIST_CODE = "" Then
            LIST_CODE = ASCMAIN1.Next_Control_No("ASTLIST1.LIST_CODE")
            rowASTLIST1 = dst.Tables("ASTLIST1").NewRow
            With rowASTLIST1
                .Item("LIST_CODE") = LIST_CODE
                .Item("COLUMN_NAME") = COLUMN_NAME_returned
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
            End With
            dst.Tables("ASTLIST1").Rows.Add(rowASTLIST1)
        Else
            rowASTLIST1 = dst.Tables("ASTLIST1").Rows.Find(LIST_CODE)
            With rowASTLIST1
                .Item("LAST_DATE") = DATETIME_STAMP
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
            End With
        End If
        rowASTLIST1.Item("LIST_DESC") = cbeLIST.Text

        dst.Tables("ASTLIST2").Rows.Clear()
        For Each row As DataRow In tbl.Select("SELECTED = '1'")
            Dim rowASTLIST2 As DataRow = dst.Tables("ASTLIST2").NewRow
            rowASTLIST2.Item("LIST_CODE") = LIST_CODE
            rowASTLIST2.Item("CODE_VALUE") = row.Item(COLUMN_NAME_returned)
            dst.Tables("ASTLIST2").Rows.Add(rowASTLIST2)
        Next

        Update_Record_TDA("ASTLIST1")
        Update_Record_TDA("ASTLIST2", sqlDelete)

        If EntryMode = "S" Then
            Me.Close()
        Else
            cmdOK_Click()
        End If

    End Sub

    Private Sub cbeLIST_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeLIST.ValueChanged
        If cbeLIST.SelectedIndex <> -1 Then
            Clear_Selections()
            LIST_CODE = cbeLIST.Value
            ASCMAIN1.sql = "Select CODE_VALUE from ASTLIST2 where LIST_CODE = '" & LIST_CODE & "'"
            For Each rowASTLIST2 As DataRow In ASCDATA1.GetDataTable.Rows
                Dim CODE_VALUE As String = rowASTLIST2.Item("CODE_VALUE")
                Dim row() As DataRow = tbl.Select(COLUMN_NAME_returned & " = '" & CODE_VALUE & "'")
                If row.Length = 1 Then
                    row(0).Item("SELECTED") = "1"
                End If
            Next
            Count_Selections("C")
            chkSelectedOnly.Checked = True
        End If
    End Sub

    'Public Sub Save_List()
    '    SplitContainer2.Visible = False
    '    splHeader.Panel1Collapsed = True
    '    EntryMode = "S"
    '    Me.ControlBox = True
    '    Me.ShowDialog()
    'End Sub

    Private Sub grd_AfterSortChange(sender As Object, e As UltraWinGrid.BandEventArgs) Handles grd.AfterSortChange
        Determine_Sort()
    End Sub

    Function Determine_Sort() As String

        Dim COLUMN_NAME As String = ""
        Dim COLUMN_CAPTION As String = ""
        For Each c As Infragistics.Win.UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
            If c.SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending Or c.SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending Then
                COLUMN_NAME = c.Key
                COLUMN_CAPTION = c.Header.Caption
                Exit For
            End If
        Next

        If COLUMN_CAPTION = "" Then
            UltraGroupBox2.Text = "Find"
        Else
            UltraGroupBox2.Text = "Find " & COLUMN_CAPTION
        End If

        Return COLUMN_NAME
    End Function

    Private Sub ASFCODE1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Determine_Sort()
    End Sub
End Class
