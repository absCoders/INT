Public Class ASFEXCL1
    Public dt As DataTable
    Public grd As UltraWinGrid.UltraGrid
    Public STATUS As String
    Dim img As Image
    Public xCi As Integer
    Public xRi As Integer
    Public Band As Integer = 0

    Public Sub New(ByVal FF As ASFBASE0)
        frmASFBASE1 = FF
        InitializeComponent()
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        grdExcel.DataSource = dt
        grdExcel.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.Select

        img = grdExcel.DisplayLayout.Override.RowSelectorHeaderAppearance.ImageBackground

        grdExcel.DisplayLayout.Override.AllowColMoving = UltraWinGrid.AllowColMoving.NotAllowed
        initialize_controls()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        STATUS = "Cancel"
        Me.Close()
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        If grdExcel.Rows.VisibleRowCount = 0 Then
            MsgBox("No Rows - You must Cancel")
            Exit Sub
        End If

        With grdExcel.DisplayLayout.Bands(0)
            Dim at_least_1_column_is_visible As Boolean = False
            For i As Integer = 0 To .Columns.Count - 1
                If Not .Columns(i).Hidden Then
                    at_least_1_column_is_visible = True
                    Exit For
                End If
            Next
            If Not at_least_1_column_is_visible Then
                MsgBox("No Columns - You must Cancel")
                Exit Sub
            End If
        End With


        STATUS = "OK"
        Me.Close()
    End Sub

    Private Sub grdExcel_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdExcel.DoubleClickCell

        If Not lblSelect.Visible Then
            Exit Sub
        End If

        Dim xgc As UltraWinGrid.UltraGridCell = grdExcel.ActiveCell
        If xgc Is Nothing Then Exit Sub

        xCi = xgc.Column.Index
        If xci > 0 Then
            For I As Integer = 0 To xci - 1
                grdExcel.DisplayLayout.Bands(0).Columns(I).Hidden = True
            Next
        End If

        xRi = xgc.Row.Index

        If xRi > 0 Then
            For I As Integer = xRi - 1 To 0 Step -1
                grdExcel.Rows(I).Hidden = True
            Next
        End If

        Assign_Headings()

        Dim LR As Integer = grdExcel.Rows.Count - 1
        If LR > 0 And grdExcel.Rows(grdExcel.Rows.Count - 1).Cells(0).Text = "Totals" Then
            LR = LR - 1
        ElseIf LR > 1 AndAlso grdExcel.Rows(grdExcel.Rows.Count - 2).Cells(0).Text = "Totals" Then
            LR = LR - 2
        End If

        If LR < grdExcel.Rows.Count - 1 Then
            For i As Integer = LR + 1 To grdExcel.Rows.Count - 1
                grdExcel.Rows(i).Hidden = True
            Next
        End If

        cmdOK.Enabled = True
        lblSelect.Visible = False
        'grdExcel.DisplayLayout.Override.RowSelectorHeaderAppearance.ImageBackground = img

    End Sub

    Private Sub grdExcel_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grdExcel.MouseDown
        'If Not lblSelect.Visible Then
        Try
            Dim pt As System.Drawing.Point = New System.Drawing.Point(e.X, e.Y)
            Dim elem As Infragistics.Win.UIElement
            elem = grdExcel.DisplayLayout.UIElement.ElementFromPoint(pt)
            If elem.GetType.Equals(GetType(Infragistics.Win.UltraWinGrid.RowSelectorHeaderUIElement)) Then
                If e.Button = Windows.Forms.MouseButtons.Left Then
                    ReSelect()
                End If
            End If
        Catch ex As Exception

        End Try
        'End If
    End Sub

    Sub initialize_controls()
        'grdExcel.DisplayLayout.Override.RowSelectorHeaderAppearance.ImageBackground = Nothing
        cmdOK.Enabled = False
        lblSelect.Visible = True
    End Sub

    Sub ReSelect()
        initialize_controls()

        If grdExcel.Rows.Count > 0 Then
            For i As Integer = 0 To grdExcel.Rows.Count - 1
                If grdExcel.Rows(i).Hidden Then
                    grdExcel.Rows(i).Hidden = False
                Else
                    Exit For
                End If
            Next
            For i As Integer = grdExcel.Rows.Count - 1 To 0 Step -1
                If grdExcel.Rows(i).Hidden Then
                    grdExcel.Rows(i).Hidden = False
                Else
                    Exit For
                End If
            Next
        End If

        dt.RejectChanges()

        With grdExcel.DisplayLayout.Bands(0)
            If .Columns.Count > 0 Then
                For i As Integer = 0 To .Columns.Count - 1
                    If .Columns(i).Hidden Then
                        .Columns(i).Hidden = False
                    End If
                Next
            End If

            If grdExcel.Rows.Count > 0 Then
                grdExcel.ActiveRow = grdExcel.Rows(0)
                grdExcel.ActiveCell = grdExcel.ActiveRow.Cells(0)
            End If
        End With


        For Each GC As UltraWinGrid.UltraGridColumn In grdExcel.DisplayLayout.Bands(0).Columns
            GC.Header.Caption = GC.Key
        Next

    End Sub

    Private Sub grdExcel_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdExcel.InitializeLayout

    End Sub

    Private Sub grdExcel_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grdExcel.KeyDown
        If e.KeyValue = Keys.Delete Then
            If grdExcel.Selected.Columns.Count <> 0 Then
                For Each GC As UltraWinGrid.ColumnHeader In grdExcel.Selected.Columns
                    GC.Column.Hidden = True
                Next
                grdExcel.Selected.Columns.Clear()
                e.Handled = True
                If Not lblSelect.Visible Then
                    Assign_Headings()
                End If
            End If
        End If
    End Sub

    Private Sub grdExcel_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles grdExcel.KeyPress
        'If e.KeyChar = Keys.Delete Then
        '    If grdExcel.Selected.Columns.Count = 0 Then
        '        For Each GC As UltraWinGrid.UltraGridColumn In grdExcel.Selected.Columns.All
        '            GC.Hidden = True
        '        Next
        '        grdExcel.Selected.Columns.Clear()
        '        e.Handled = True
        '    End If
        'End If

    End Sub

    Sub Assign_Headings()

        Dim tbl As New DataTable
        With tbl
            .Columns.Add("GROUP", GetType(System.Int64))
            .Columns.Add("LEVEL", GetType(System.Int64))
            .Columns.Add("COLUMN", GetType(System.Int64))
            .Columns.Add("KEY")
        End With

        Dim SL As New SortedList(Of Integer, String)
        For Each GC As UltraWinGrid.UltraGridColumn In _
        grd.DisplayLayout.Bands(Band).Columns
            If Not GC.Hidden Then
                If grd.DisplayLayout.Bands(Band).Groups.Count = 0 Then
                    tbl.Rows.Add(New Object() {0, 0, GC.Header.VisiblePosition, GC.Key})
                Else
                    If GC.Group IsNot Nothing AndAlso Not GC.Group.Hidden Then
                        tbl.Rows.Add(New Object() {GC.Group.Header.VisiblePosition, GC.Level, GC.Header.VisiblePosition, GC.Key})
                    End If
                End If
            End If
        Next

        Dim c As Int64 = 0
        For Each row As DataRow In tbl.Select("", "GROUP,LEVEL,COLUMN")
            SL.Add(c, row.Item("KEY"))
            c += 1
        Next

        Dim hca As Integer = 0
        For SLi As Integer = 0 To SL.Count - 1
            Dim GC As UltraWinGrid.UltraGridColumn = grd.DisplayLayout.Bands(Band).Columns(SL.Values(SLi))
            With grdExcel.DisplayLayout.Bands(0)
                If SLi + hca < .Columns.Count Then

                    Do While .Columns(SLi + hca).Hidden
                        hca += 1
                        If SLi + hca >= .Columns.Count Then Exit Do
                    Loop
                    If SLi + hca < .Columns.Count Then
                        If GC.Group IsNot Nothing Then
                            If GC.Group.Columns.Count = 1 Or (grd.DisplayLayout.Bands(Band).LevelCount = 1 And GC.Group.Width = GC.Width) Then
                                .Columns(SLi + hca).Header.Caption = GC.Group.Header.Caption
                                .Columns(SLi + hca).Tag = GC.Key
                            Else
                                .Columns(SLi + hca).Header.Caption = GC.Header.Caption
                                .Columns(SLi + hca).Tag = GC.Key
                            End If
                        Else
                            .Columns(SLi + hca).Header.Caption = GC.Header.Caption
                            .Columns(SLi + hca).Tag = GC.Key
                        End If
                    End If
                End If

            End With

        Next

        Exit Sub
    End Sub
End Class