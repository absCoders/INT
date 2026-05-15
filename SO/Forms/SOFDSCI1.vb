Imports Infragistics.Win.UltraWinGrid
Public Class SOFDSCI1

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.
    Dim CUST_CODEs As New List(Of String)
    Private ColumnKey As String = String.Empty
    Dim NON_CUST_COLs As HashSet(Of String) = New HashSet(Of String)({"COLLECTION_CODE", "ITEM_CODE", "BRAND_CODE", "ITEM_STATUS", "ITEM_DESC"})
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        grdSOTDSCIX.Visible = False
        grdASTAUDT1.Visible = False
        UltraExplorerBar1.Groups("Audit Trail").Visible = False
        UltraExplorerBar1.Groups("Add Customer").Visible = False
        UltraExplorerBar1.Groups("Screen Control").Items("Done").Settings.Enabled = True
        UltraExplorerBar1.Groups("Screen Control").Items("Update").Settings.Enabled = True
        With dst
            Create_TDA(.Tables.Add, "SOTDSCI1", "*", 0)
            With .Tables("SOTDSCI1")
                .Columns.Add("NEW_VAL")
                .Columns.Add("OLD_VAL")
                .Columns("OLD_VAL").DefaultValue = "0"
                .Columns("NEW_VAL").DefaultValue = "0"
            End With

            ASCMAIN1.sql = "SELECT ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_STATUS, ICTITEM1.COLLECTION_CODE, ICTCOLL1.BRAND_CODE FROM ICTITEM1,ICTCOLL1 WHERE ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE AND ICTITEM1.ITEM_ORDR_REL_CODE = 'D'"
            Create_TDA(.Tables.Add, "SOTDSCIX", "**", 0, False, "", 1)
        End With
        grdSOTDSCIX.DataSource = dst.Tables("SOTDSCIX")
        grdSOTDSCIX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        Create_Summary(grdSOTDSCIX, "ITEM_CODE", "Count")
        Show_Filter(grdSOTDSCIX)
        For i As Integer = 0 To 4
            grdSOTDSCIX.DisplayLayout.Bands(0).Columns(i).Header.Fixed = True 'freeze item code
        Next
        grdASTAUDT1.DataSource = dst.Tables("ASTAUDT1")
        ASCMAIN1.sql = "SELECT * FROM ASTAUDT1 WHERE TABLE_NAME = 'SOTDSCI1'"
        Fill_Records("ASTAUDT1", "", True, ASCMAIN1.sql)
        Sort_grdColumns(grdASTAUDT1, "init_date")
        Create_Summary(grdASTAUDT1, "KEY_VALUE", "Count")

    End Sub


    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"


            Case "Done"

            Case "Update"
                If CUST_CODEs.Count = 0 Then
                    EMsg &= vbCr & "No Customers Specified"
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
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Update"
                Update_Record()

            Case "Refresh"
                Sort_Customers()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Done").Visible = tf
            .Groups("Screen Control").Items("Update").Visible = tf
            .Groups("Screen Control").Items("Refresh").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Refresh").Visible = tf
            .Groups("Add Customer").Visible = tf
            .Groups("Audit Trail").Visible = tf
        End With

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        grdSOTDSCIX.Visible = ScreenMode
        cmdAddMultipleCustomers.Enabled = ScreenMode
        chkAudit.Enabled = ScreenMode
        grdASTAUDT1.Visible = ScreenMode And chkAudit.Checked

        If Not ScreenMode Then
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTDSCI1", "SOTDSCIX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Clear_All_Filters(grdSOTDSCIX)
    End Sub

    Sub Load_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Fill_Records("SOTDSCIX")
        Fill_Records("SOTDSCI1")
        CUST_CODEs.Clear()

        ' Use SortedSet to collect unique customer codes and automatically sort them
        Dim sortedCustomers = New SortedSet(Of String)()

        For Each row As DataRow In dst.Tables("SOTDSCI1").Rows
            Dim CUST_CODE = row("CUST_CODE").ToString()
            sortedCustomers.Add(CUST_CODE)
            CUST_CODEs.Add(CUST_CODE)
        Next

        ' Add sorted customer columns
        For Each CUST_CODE As String In sortedCustomers
            If Not dst.Tables("SOTDSCIX").Columns.Contains(CUST_CODE) Then
                Add_Column(CUST_CODE)
            End If
        Next

        ' Set OLD_VAL = 1 for rows in SOTDSCI1 and update grid cells
        For Each row As DataRow In dst.Tables("SOTDSCI1").Rows
            row("OLD_VAL") = "1"
            Dim ITEM_CODE As String = row("ITEM_CODE").ToString()
            Dim CUST_CODE As String = row("CUST_CODE").ToString()

            ' Check and update grid cells
            If grdSOTDSCIX.DisplayLayout.Bands(0).Columns.Exists(CUST_CODE) Then
                For Each gridRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSOTDSCIX.Rows
                    If gridRow.Cells("ITEM_CODE").Value.ToString() = ITEM_CODE Then
                        gridRow.Cells(CUST_CODE).Value = "1"
                    End If
                Next
            End If
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub


    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Updating...")

        'go through the grid
        For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSOTDSCIX.Rows
            'get the item code
            Dim ITEM_CODE As String = row.Cells("ITEM_CODE").Value.ToString()
            ' Check each customer column in the grid
            For Each CUST_CODE As String In CUST_CODEs
                If grdSOTDSCIX.DisplayLayout.Bands(0).Columns.Exists(CUST_CODE) Then
                    ' Check if the cell for this customer and item is checked
                    Dim isChecked As String = row.Cells(CUST_CODE).Value & ""
                    'check if this row has already been added 
                    Dim existingRows = dst.Tables("SOTDSCI1").Select($"ITEM_CODE = '{ITEM_CODE}' AND CUST_CODE = '{CUST_CODE}'")
                    If existingRows.Length > 0 Then
                        Dim existingRow = existingRows(0)
                        'if it has, update newval in case they unchecked it
                        existingRow("NEW_VAL") = isChecked
                    Else
                        ' If the row does not exist, add a new row with NEW_VAL set based on isChecked
                        Dim newRow As DataRow = dst.Tables("SOTDSCI1").NewRow()
                        newRow("ITEM_CODE") = ITEM_CODE
                        newRow("CUST_CODE") = CUST_CODE
                        newRow("INIT_DATE") = DATETIME_STAMP
                        newRow("INIT_OPER") = ASCMAIN1.USER_ID
                        newRow("LAST_OPER") = ASCMAIN1.USER_ID
                        newRow("LAST_DATE") = DATETIME_STAMP
                        newRow("NEW_VAL") = isChecked
                        dst.Tables("SOTDSCI1").Rows.Add(newRow)
                    End If
                End If
            Next
        Next
        'loop through 1 table 
        For Each row As DataRow In dst.Tables("SOTDSCI1").Rows
            Dim CUST_CODE = row("CUST_CODE").ToString()
            Dim ITEM_CODE = row("ITEM_CODE").ToString()
            Dim oldVal = row("OLD_VAL").ToString()
            Dim newVal = row("NEW_VAL").ToString()
            'means there was a change 
            If oldVal <> newVal Then
                Dim rowASTAUDT1 As DataRow = dst.Tables("ASTAUDT1").NewRow
                rowASTAUDT1.Item("TABLE_NAME") = "SOTDSCI1"
                rowASTAUDT1.Item("KEY_VALUE") = $"{CUST_CODE}:{ITEM_CODE}"
                rowASTAUDT1.Item("COLUMN_NAME") = "*" ' tblASFBASE1.Columns(i).ColumnName
                rowASTAUDT1.Item("USER_ID") = ASCMAIN1.USER_ID
                rowASTAUDT1.Item("INIT_DATE") = DATETIME_STAMP
                rowASTAUDT1.Item("OLD_VALUE") = oldVal
                rowASTAUDT1.Item("NEW_VALUE") = newVal
                'OLD 0, NEW 1 = ADD
                'OLD 1, NEW 0 = DELETE 
                rowASTAUDT1.Item("FM_MODE") = "E"
                rowASTAUDT1.Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                rowASTAUDT1.Item("SELECTION_NO") = SELECTION_NO
                rowASTAUDT1.Item("XNO") = XNO
                dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)
            End If
        Next
        For Each row As DataRow In dst.Tables("SOTDSCI1").Select("NEW_VAL <> '1'")
            dst.Tables("SOTDSCI1").Rows.Remove(row)
        Next

        Update_Record_TDA("ASTAUDT1")
        Update_Record_TDA("SOTDSCI1", "1=1")


        MsgBox("Update Complete", MsgBoxStyle.OkOnly, "Verification")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTDSCIX, "SSSBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Check All Items", "Un-Check All Items", "Check Item For All Customers", "Un-Check Item For All Customers")
        Load_Popup_Menu(grdASTAUDT1, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
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

        For Each tool As Infragistics.Win.UltraWinToolbars.ToolBase In tlb_pop.Tools
            If tool.Key = "Check All Items" OrElse tool.Key = "Un-Check All Items" Then
                If NON_CUST_COLs.Contains(ColumnKey) Then
                    tool.SharedProps.Visible = False
                Else
                    tool.SharedProps.Visible = True
                End If
            ElseIf tool.Key = "Check Item For All Customers" OrElse tool.Key = "Un-Check Item For All Customers" Then
                If ColumnKey = "ITEM_CODE" Then
                    tool.SharedProps.Visible = True
                Else
                    TOOL.SharedProps.Visible = False
                End If
            End If
        Next

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Check All Items", "Un-Check All Items"
                If Not NON_CUST_COLs.Contains(ColumnKey) Then
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                        If Not grow.IsFilteredOut Then
                            grow.Cells(ColumnKey).Value = If(e.Tool.Key = "Check All Items", "1", "0")
                            grow.Update()
                        End If
                    Next
                End If
            Case "Check Item For All Customers", "Un-Check Item For All Customers"
                If grdSOTDSCIX.ActiveRow IsNot Nothing Then
                    For columnIndex As Integer = 5 To grdSOTDSCIX.ActiveRow.Cells.Count - 1
                        grdSOTDSCIX.ActiveRow.Cells(columnIndex).Value = If(e.Tool.Key = "Check Item For All Customers", "1", "0")
                    Next
                    grdSOTDSCIX.Refresh()
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key


        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

#End Region

    Private Sub cmdAddMultipleCustomers_Click(sender As Object, e As EventArgs) Handles cmdAddMultipleCustomers.Click

        Dim CUST_CODEsX As String = ""

        For Each row As DataRow In dst.Tables("SOTDSCI1").Select("")
            CUST_CODEs.Add(row.Item("CUST_CODE"))
        Next
        If CUST_CODEs.Count <> 0 Then CUST_CODEsX = "CUST_CODE NOT in ('" & Join(CUST_CODEs.ToArray, "','") & "')"

        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("CUST_CODE", "", CUST_CODEsX)

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading")

                For Each CUST_CODE As String In ASCMAIN1.CodeSelector.SelectedCodes
                    'Add_CUST_CODE(CUST_CODE, "")
                    If Not dst.Tables("SOTDSCIX").Columns.Contains(CUST_CODE) Then
                        Add_Column(CUST_CODE)
                        CUST_CODEs.Add(CUST_CODE)
                    End If
                Next

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
            End If

        End If
    End Sub
    Sub Add_Column(COL_NAME As String)

        With dst.Tables("SOTDSCIX")
            .Columns.Add(COL_NAME)
            .Columns(COL_NAME).DefaultValue = "0"

            With grdSOTDSCIX.DisplayLayout.Bands(0).Columns(COL_NAME)
                .Style = UltraWinGrid.ColumnStyle.CheckBox
                .Header.Caption = COL_NAME
                .Header.Appearance.TextHAlign = HAlign.Center
                .CellAppearance.TextHAlign = HAlign.Center
                .Header.Appearance.BackColor = System.Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                .Hidden = False
                .CellActivation = UltraWinGrid.Activation.AllowEdit
                .CellClickAction = UltraWinGrid.CellClickAction.Edit
                Create_Summary(grdSOTDSCIX, COL_NAME)
            End With
            For Each rowSOTDSCIX As DataRow In dst.Tables("SOTDSCIX").Select("")
                rowSOTDSCIX.Item(COL_NAME) = "0"
            Next
        End With
        ASCMAIN1.grdInitializeLayout(grdSOTDSCIX, Me)

    End Sub

    Private Sub chkAudit_CheckedChanged(sender As Object, e As EventArgs) Handles chkAudit.CheckedChanged
        If chkAudit.Checked Then
            grdASTAUDT1.Visible = True
            grdSOTDSCIX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        Else
            grdASTAUDT1.Visible = False
            grdSOTDSCIX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        End If

    End Sub

    Private Sub grdSOTDSCIX_ClickCell(sender As Object, e As Infragistics.Win.UltraWinGrid.ClickCellEventArgs) Handles grdSOTDSCIX.ClickCell
        Dim filterCondition As String = ""
        Dim resetColumns As HashSet(Of String) = New HashSet(Of String)({"Collection", "Brand", "Status", "Desc"})

        If e.Cell IsNot Nothing Then
            If e.Cell.Column.Key = "ITEM_CODE" Then
                Dim ITEM_CODE As String = e.Cell.Value.ToString()
                filterCondition = $"KEY_VALUE LIKE '%{ITEM_CODE}%'"
            ElseIf resetColumns.Contains(e.Cell.Column.Header.Caption) Then
            Else
                Dim CUST_CODE As String = e.Cell.Column.Header.Caption
                Dim KEY_VALUE As String = $"{CUST_CODE}:{e.Cell.Row.Cells("ITEM_CODE").Value.ToString()}"
                filterCondition = $"KEY_VALUE = '{KEY_VALUE.Replace("'", "''")}'"
            End If

            ' Build SQL query based on the determined filter condition
            If Not String.IsNullOrWhiteSpace(filterCondition) Then
                ASCMAIN1.sql = $"SELECT * FROM ASTAUDT1 WHERE TABLE_NAME = 'SOTDSCI1' AND {filterCondition}"
            Else
                ASCMAIN1.sql = "SELECT * FROM ASTAUDT1 WHERE TABLE_NAME = 'SOTDSCI1'"
            End If
            Me.Cursor = Cursors.WaitCursor
            Try
                Fill_Records("ASTAUDT1", "", True, ASCMAIN1.sql)
                Dim dView As New DataView(dst.Tables("ASTAUDT1"))
                If Not String.IsNullOrWhiteSpace(filterCondition) Then
                    dView.RowFilter = filterCondition
                End If
                grdASTAUDT1.DataSource = dView
            Finally
                Me.Cursor = Cursors.Default
            End Try
        End If
    End Sub

    Private Sub grdSOTDSCIX_MouseDown(sender As Object, e As MouseEventArgs) Handles grdSOTDSCIX.MouseDown
        If e.Button = MouseButtons.Right Then
            'element at right click location
            Dim element As Infragistics.Win.UIElement = grdSOTDSCIX.DisplayLayout.UIElement.ElementFromPoint(e.Location)
            If element IsNot Nothing Then
                'use element to get the cell
                Dim cell As Infragistics.Win.UltraWinGrid.UltraGridCell = TryCast(element.GetContext(GetType(Infragistics.Win.UltraWinGrid.UltraGridCell)), Infragistics.Win.UltraWinGrid.UltraGridCell)
                If cell IsNot Nothing Then
                    ColumnKey = cell.Column.Key
                End If
            End If
        End If
    End Sub
    Private Sub Sort_Customers()
        Dim grid = grdSOTDSCIX.DisplayLayout.Bands(0)
        Dim fixedColumns As New List(Of UltraGridColumn)()
        Dim customerColumns As New List(Of UltraGridColumn)()

        For Each col As UltraGridColumn In grid.Columns
            If NON_CUST_COLs.Contains(col.Key) Then
                fixedColumns.Add(col)
            Else
                customerColumns.Add(col)
            End If
        Next

        customerColumns.Sort(Function(x, y) String.Compare(x.Key, y.Key))

        Dim visiblePosition As Integer = 0

        For Each col As UltraGridColumn In fixedColumns
            col.Header.VisiblePosition = visiblePosition
            visiblePosition += 1
        Next

        For Each col As UltraGridColumn In customerColumns
            col.Header.VisiblePosition = visiblePosition
            visiblePosition += 1
        Next
        grdSOTDSCIX.Refresh()
    End Sub
End Class