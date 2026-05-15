Public Class ASFAPPD1
    Public grd As UltraWinGrid.UltraGrid
    Public grow As UltraWinGrid.UltraGridRow
    Public gcol As UltraWinGrid.UltraGridColumn
    Private f_Calling_Form As ASFBASE0
    Public user_option As Integer = 0

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        user_option = -1
        Me.Close()
    End Sub

    Private Sub ASFMSGBF_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call ASCMAIN1.Center(Me)
        'grdmsg.DisplayLayout.Override.RowSelectors = DefaultableBoolean.Default
        'grdmsg.DisplayLayout.Override.RowSelectorHeaderStyle = UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        'ASCMAIN1.grdInitializeLayout(grdmsg)
        'ASCMAIN1.MainForm_pgd.SelectedObject = sender

        Dim tbl As New DataTable
        tbl.Columns.Add("BAND")
        tbl.Columns.Add("COLUMN")

        For Each B As UltraWinGrid.UltraGridBand In grd.DisplayLayout.Bands
            For Each C As UltraWinGrid.UltraGridColumn In B.Columns
                If Not C.Hidden Then
                    tbl.Rows.Add(New String() {B.Key, C.Key})
                End If
            Next
        Next
        grdColumns.DataSource = tbl

        With grdColumns.DisplayLayout
            .Override.CellClickAction = UltraWinGrid.CellClickAction.RowSelect
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowDelete = DefaultableBoolean.False
            .Override.AllowUpdate = DefaultableBoolean.False
            .GroupByBox.Hidden = True
        End With
        With grdColumns.DisplayLayout.Bands(0)
            .Columns(0).HiddenWhenGroupBy = DefaultableBoolean.True
            .SortedColumns.Add("BAND", False, True)
        End With
        grdColumns.Rows.ExpandAll(True)

        If gcol Is Nothing Then
            gcol = grd.DisplayLayout.Bands(0).Columns(0)
        End If

        If gcol IsNot Nothing Then
            For Each grow As UltraWinGrid.UltraGridRow In grdColumns.Rows
                If grow.IsDataRow Then
                    If grow.Cells("BAND").Value = gcol.Band.Key And grow.Cells("COLUMN").Value = gcol.Key Then
                        grdColumns.ActiveRow = grow
                        Exit Sub
                    End If
                End If
            Next

            Set_Appearance()
        End If
    End Sub

    Private Sub optCH_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCH.ValueChanged
        Set_Appearance()
    End Sub

    Sub Set_Appearance()
        If grd Is Nothing Then Exit Sub
        If optCH.Value = "C" Then
            pgdExplorerBar.SelectedObject = grd.DisplayLayout.Bands(0).Columns(gcol.Key).CellAppearance
        Else
            pgdExplorerBar.SelectedObject = grd.DisplayLayout.Bands(0).Columns(gcol.Key).Header.Appearance
        End If

        grpAppearance.Text = gcol.Key & "(" & optCH.Text & ")"

    End Sub

    Private Sub grdColumns_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdColumns.AfterRowActivate
        If grdColumns.ActiveRow.IsDataRow Then
            Dim b As String = grdColumns.ActiveRow.Cells("BAND").Value
            Dim C As String = grdColumns.ActiveRow.Cells("COLUMN").Value
            gcol = grd.DisplayLayout.Bands(b).Columns(C)
            Set_Appearance()
        End If
    End Sub

End Class