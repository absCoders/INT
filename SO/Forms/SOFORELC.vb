Imports Infragistics.Win.UltraWinGrid
Imports Infragistics.Win.UltraWinExplorerBar
Imports System.IO
Public Class SOFORELC
    Private selectedRowCount As Integer = 0
    Private selectedRowIds As New HashSet(Of Object)
    Dim FindConnectedReportsItem As New UltraExplorerBarItem()
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        dteREPORT_DATE_TO.Value = DateTime.Today
        dteREPORT_DATE_FROM.Value = DateTime.Today.AddDays(-7)

        Dim fromDate As String = Format(dteREPORT_DATE_FROM.Value, "dd-MMM-yyyy")
        Dim toDate As String = Format(DateValue(dteREPORT_DATE_TO.Value.ToString).AddDays(1), "dd-MMM-yyyy")

        With dst
            'selecting the reports between these dates
            ASCMAIN1.sql = $"Select * From SOTORELC WHERE report_date BETWEEN TO_DATE('{fromDate}', 'DD-MON-YYYY') AND TO_DATE('{toDate}', 'DD-MON-YYYY') "
            Create_TDA(.Tables.Add, "SOTORELX", "**", 0, False)

            ' Add the "SEL" column to the DataTable
            With .Tables("SOTORELX")
                .Columns.Add("SEL") ' Specify the data type as Integer for the 0 or 1 values
                .Columns("SEL").DefaultValue = "0"
                .Columns("SEL").Caption = "Sel"
                .Columns("SEL").SetOrdinal(0)
            End With
        End With

        ' Set the DataTable as the DataSource for the grid
        grdSOTORELC.DataSource = dst.Tables("SOTORELX")
        Create_Summary(grdSOTORELC, "SEL")
        Create_Summary(grdSOTORELC, "XNO", "Count")

        ' Get the "SEL" column from the grid and make it a checkbox
        Dim selectedColumn As UltraWinGrid.UltraGridColumn = grdSOTORELC.DisplayLayout.Bands(0).Columns("SEL")
        selectedColumn.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        selectedColumn.CellClickAction = UltraWinGrid.CellClickAction.Edit

        ' Allow editing only for "SEL" column
        For Each gcol As UltraWinGrid.UltraGridColumn In grdSOTORELC.DisplayLayout.Bands(0).Columns
            gcol.CellActivation = If(gcol.Key = "SEL", UltraWinGrid.Activation.AllowEdit, UltraWinGrid.Activation.NoEdit)
        Next

        ' Enable updates and set the appropriate actions for the "SEL" column
        grdSOTORELC.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdSOTORELC.DisplayLayout.Bands(0).Columns("SEL").CellClickAction = UltraWinGrid.CellClickAction.Edit
        grdSOTORELC.DisplayLayout.Bands(0).Columns("SEL").Hidden = False

        ' Show the filter for the grid
        Show_Filter(grdSOTORELC, True)
        ' Add the "View Reports" button to the UltraExplorerBar
        FindConnectedReportsItem.Key = "FindConnectedReports"
        FindConnectedReportsItem.Text = "View Reports"
        UltraExplorerBar1.Groups("Screen Control").Items.Add(FindConnectedReportsItem)
        FindConnectedReportsItem.Settings.Enabled = False
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load Reports"
                If dteREPORT_DATE_FROM.Value Is Nothing Then
                    EMsg &= vbCr & "No Date Specified"
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

            Case "Load Reports"
                Call Load_Record()
                Call Mode_Settings(True)
                FindConnectedReportsItem.Settings.Enabled = iScreenMode
            Case "Done"
                'If ASCMAIN1.Running_in_VS Then Move_Reports() : Exit Sub
                selectedRowCount = 0
                selectedRowIds.Clear()
                Call Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Load Reports").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            If .Groups("Screen Control").Items.Exists("FindConnectedReports") Then
                ' Enable or disable the "View Reports" button based on the screen mode
                .Groups("Screen Control").Items("FindConnectedReports").Settings.Enabled = iScreenMode
            End If
        End With

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only(grpLoadOptions, ScreenMode)

        grdSOTORELC.Visible = ScreenMode
        If Not ScreenMode Then
            Clear_Record()
        End If


    End Sub
    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTORELX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
        grdSOTORELC.DisplayLayout.Bands("SOTORELX").ColumnFilters.ClearAllFilters()
        ' Dim dvw As DataView = DirectCast(grdICTFAXFR1.DataSource, DataTable).DefaultView
    End Sub
    Sub Load_Record()
        ASCMAIN1.Progress("Now Loading Report from Archive")
        Me.Cursor = Cursors.WaitCursor

        Dim sql As String = ""
        'from date is a week before todate to start
        Dim fromDate As String = Format(dteREPORT_DATE_FROM.Value, "dd-MMM-yyyy")
        Dim toDate As String = Format(DateValue(dteREPORT_DATE_TO.Value.ToString).AddDays(1), "dd-MMM-yyyy")

        sql &= " AND REPORT_DATE >= '" & fromDate & "'"
        sql &= " AND REPORT_DATE <= '" & toDate & "'"

        ASCMAIN1.sql = $"SELECT * FROM SOTORELC" & ASCMAIN1.SQL_Add_WHERE(sql)
        Fill_Records("SOTORELX", "", True, ASCMAIN1.sql)
        For Each row As DataRow In dst.Tables("SOTORELX").Rows
            If Not IsDBNull(row("XNO")) Then
                row("XNO") = row("XNO").ToString().PadLeft(10, "0"c)
            End If
        Next
        Sort_grdColumns(grdSOTORELC, "REPORT_DATE")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
    Sub Update_Record()

    End Sub
#End Region

#Region "Popup_Menus"
    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdSOTORELC, "SBBBBBB", "Show Filter", "Select Selected", "De-Select Selected", "Select All", "De-Select All", "Show Selected Only", "Show All")
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
        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
            End Select

        End If
    End Sub
    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Select All", "De-Select All"
                ' Iterate through all rows and check if they meet the filter criteria
                ' Reset the selectedRowCount to zero when "Select All" is clicked
                selectedRowCount = 0
                selectedRowIds.Clear()

                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If Not grow.IsFilteredOut Then
                        grow.Cells("SEL").Value = If(e.Tool.Key = "Select All", "1", "0")
                        grow.Update()
                        ' Update the counter based on the "Select All" action
                        selectedRowCount += If(CBool(grow.Cells("SEL").Value), 1, 0)
                        ' Track all rows in the "Select All" action
                        If e.Tool.Key = "Select All" Then
                            selectedRowIds.Add(grow.GetCellValue(grow.Band.Columns("XNO")))
                        End If
                    End If
                Next

            Case "Select Selected"
                ASCMAIN1.Progress("Now Selecting Selected Rows")
                ASCMAIN1.Progress("")
                ' Iterate through selected rows and check if they meet the filter criteria
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    If Not grow.IsFilteredOut Then
                        If Not IsRowSelected(grow) Then
                            grow.Cells("SEL").Value = "1"
                            grow.Update()

                            ' Update the counter based on the "Select Selected" action
                            selectedRowCount += 1

                            ' Track the selected row by its unique identifier (XNO)
                            selectedRowIds.Add(grow.GetCellValue(grow.Band.Columns("XNO")))
                        End If
                    End If
                Next
            Case "De-select Selected"
                ASCMAIN1.Progress("Now De-Selecting Selected Rows")
                ASCMAIN1.Progress("")
                ' Iterate through selected rows and check if they meet the filter criteria
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    If Not grow.IsFilteredOut Then
                        If Not IsRowSelected(grow) Then
                            grow.Cells("SEL").Value = "0"
                            grow.Update()

                            ' Update the counter based on the "Select Selected" action
                            selectedRowCount -= 1

                            ' Track the selected row by its unique identifier (XNO)
                            selectedRowIds.Remove(grow.GetCellValue(grow.Band.Columns("XNO")))
                        End If
                    End If
                Next
            Case "Show Selected Only"
                ' Iterate through all rows and filter out those that are not selected
                For Each grow As UltraWinGrid.UltraGridRow In grdSOTORELC.Rows
                    Dim isSelected As Boolean = CBool(grow.Cells("SEL").Value)
                    grow.Hidden = Not isSelected
                Next
            Case "Show All"
                ' Iterate through all rows and reset their visibility
                For Each grow As UltraWinGrid.UltraGridRow In grdSOTORELC.Rows
                    grow.Hidden = False
                Next
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub
#End Region
#Region "grdSOTORELC"
    Private Sub grdSOTORELC_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTORELC.InitializeLayout
        ' Ensure there are columns in the grid
        If e.Layout.Bands.Count > 0 AndAlso e.Layout.Bands(0).Columns.Count > 0 Then
            Dim selectedColumn As UltraWinGrid.UltraGridColumn = e.Layout.Bands(0).Columns("SEL")
            selectedColumn.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
            selectedColumn.Header.VisiblePosition = 0
            selectedColumn.CellClickAction = UltraWinGrid.CellClickAction.Edit ' Set CellClickAction to Edit for checkbox behavior

            ' Specify the order of the remaining columns: REPORT_DATE, XNO, Hold Codes
            Dim columnOrder As String() = {"REPORT_DATE", "XNO", "ORDR_REL_HOLD_CODES"}

            ' Loop through the specified order and set the DisplayIndex for each column
            For i As Integer = 0 To columnOrder.Length - 1
                Dim columnName As String = columnOrder(i)

                ' Find the column by name in the grid's Bands
                Dim column As UltraWinGrid.UltraGridColumn = e.Layout.Bands(0).Columns(columnName)

                ' Set the DisplayIndex to the desired order
                If column IsNot Nothing Then
                    column.Header.VisiblePosition = i + 1 ' Offset by 1 to account for the "SEL" column
                End If

                ' Set column captions if needed
                Select Case columnName
                    Case "REPORT_DATE"
                        column.Header.Caption = "Report Date"
                    Case "XNO"
                        column.Header.Caption = "X-No"
                    Case "ORDR_REL_HOLD_CODES"
                        column.Header.Caption = "Hold Codes"
                    Case "SEL"
                        column.Header.Caption = "Sel"
                End Select
            Next
        End If
    End Sub

    Private Sub grdsotorelc_ClickCell(sender As Object, e As ClickCellEventArgs) Handles grdSOTORELC.ClickCell
        ' Check if the clicked cell is in the "SEL" column and is a data cell
        If e.Cell.Column.Key = "SEL" AndAlso e.Cell.IsDataCell Then
            ' Toggle the value of the "SEL" cell
            e.Cell.Value = Not CBool(e.Cell.Value)

            ' Update the counter based on the "SEL" cell value
            If CBool(e.Cell.Value) Then
                selectedRowCount += 1
                ' Track the selected row by its unique identifier (XNO)
                selectedRowIds.Add(e.Cell.Row.GetCellValue(e.Cell.Row.Band.Columns("XNO")))
            Else
                selectedRowCount -= 1
                ' Remove the row from the tracked selection
                selectedRowIds.Remove(e.Cell.Row.GetCellValue(e.Cell.Row.Band.Columns("XNO")))
            End If

            ' Force a refresh to make sure the UI reflects the changes
            grdSOTORELC.Refresh()

            ' Ensure that the row gets selected or deselected based on the checkbox state
            e.Cell.Row.Selected = CBool(e.Cell.Value)
        End If
    End Sub

    Private Sub grdSOTORELC_AfterCellUpdate(sender As Object, e As CellEventArgs)
        ' Check if the updated cell is in the "SEL" column and is a data cell
        If e.Cell.Column.Key = "SEL" AndAlso e.Cell.IsDataCell Then
            ' Manually update the row selection state based on the "SEL" cell value
            e.Cell.Row.Selected = CBool(e.Cell.Value)
        End If
    End Sub
    Private Sub UltraExplorerBar1_ItemClick(sender As Object, e As UltraWinExplorerBar.ItemEventArgs) Handles UltraExplorerBar1.ItemClick
        ' Check if the clicked item is the "View Reports" button
        Dim lengthOfSelectedRowIds As Integer = selectedRowIds.Count
        If e.Item.Key = "FindConnectedReports" Then
            ' Ensure that at least one row is selected before allowing the action
            If selectedRowCount > 0 Then
                ' Display a confirmation message
                Dim confirmationMessage As String = $"You are about to bring up {lengthOfSelectedRowIds} report(s). Do you want to proceed?"
                Dim confirmationResult As DialogResult = MessageBox.Show(confirmationMessage, "Confirmation", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)

                ' Check the user's choice
                If confirmationResult = DialogResult.OK Then
                    ' User clicked "OK," proceed with displaying reports

                    ' Create a list to store data for all selected rows
                    Dim selectedRowsData As New List(Of DataTable)
                    ' Perform the action related to "View Reports" for each selected row
                    For Each selectedRowId As Object In selectedRowIds
                        ' Retrieve data for the selected row
                        Dim TBLASTSPRF1 As New DataTable
                        'added in the or statement because both rpt titles are seen in the archive
                        ASCMAIN1.sql = $"Select * from ASTSPRF1 where XNO IS NOT NULL and form_name = 'SOROREL1' and RPT = 'SOROREL5' and XNO = '{selectedRowId}'"
                        TBLASTSPRF1 = ASCDATA1.GetDataTable
                        ' Add the data table to the list
                        selectedRowsData.Add(TBLASTSPRF1)
                    Next
                    ' Check if there is any data for the selected rows
                    If selectedRowsData.Count > 0 Then
                        Me.Cursor = Cursors.WaitCursor
                        ASCMAIN1.Progress("Now Retrieving Reports")

                        ' Iterate through the list of data tables and display reports for each
                        For Each dataTable As DataTable In selectedRowsData
                            If dataTable.Rows.Count > 0 Then
                                Dim FORM_NAME As String = "SOROREL1"
                                ' Filter the data table based on the current FORM_NAME and XNO
                                Dim dvw As New DataView(dataTable)
                                dvw.RowFilter = $"FORM_NAME = '{FORM_NAME}'"
                                ' Display the report for the current row
                                Dim f As New ASFSRPTV
                                f.Set_Table(dvw.ToTable)
                            Else
                                Exit Sub
                            End If
                        Next
                        Me.Cursor = Cursors.Default
                        ASCMAIN1.Progress("")
                    End If
                End If
            Else
                ' Inform the user that at least one row should be selected
                MsgBox("Please select at least one report to view.", MsgBoxStyle.Information)
            End If
        End If
    End Sub

    Private Sub grdASTSPRF1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTORELC.DoubleClickRow
        Dim TBLASTSPRF1 As New DataTable
        ASCMAIN1.sql = "Select * from ASTSPRF1 where XNO IS NOT NULL and form_name = 'SOROREL1' and (rpt_title = 'Orders Not Released Report' or rpt_title = 'Un-Releasable Orders Report')"
        TBLASTSPRF1 = ASCDATA1.GetDataTable
        If grdSOTORELC.ActiveRow Is Nothing OrElse Not grdSOTORELC.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Retrieving Report")

        Dim FORM_NAME As String = "SOROREL1"
        Dim XNO As String = grdSOTORELC.ActiveRow.Cells("XNO").Text

        Dim dvw As New DataView(TBLASTSPRF1)
        dvw.RowFilter = "FORM_NAME = '" & FORM_NAME & "' and XNO = '" & XNO & "'"
        Dim f As New ASFSRPTV
        f.Set_Table(dvw.ToTable)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
    Private Function IsRowSelected(row As UltraWinGrid.UltraGridRow) As Boolean
        ' Check if the row is selected based on the "SEL" column value
        Return CBool(row.Cells("SEL").Value)
    End Function
#End Region
End Class