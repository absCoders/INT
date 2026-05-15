<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ABSC
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim Appearance126 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("ASTDSQLA", -1)
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COLUMN_NAME")
        Dim UltraGridColumn2 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COLUMN_CAPTION")
        Dim Appearance127 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance128 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CODE_VALUES")
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EXCLUDE")
        Dim Appearance129 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance130 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridColumn5 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SEQUENCE")
        Dim Appearance131 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance132 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridColumn6 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PAGE_BREAK")
        Dim Appearance133 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance134 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridColumn7 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SORTABLE")
        Dim UltraGridColumn8 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("GROUP_ALL_OTHERS")
        Dim Appearance135 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance136 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridColumn9 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COLUMN_LAST")
        Dim UltraGridColumn10 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SET_SEQ", 0, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Descending, False)
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ABSC))
        Dim Appearance12 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance137 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance138 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance139 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance140 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance141 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance142 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance143 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance145 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance146 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.grdSetup = New Infragistics.Win.UltraWinGrid.UltraGrid()
        CType(Me.grdSetup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grdSetup
        '
        Appearance126.BackColor = System.Drawing.SystemColors.Window
        Appearance126.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdSetup.DisplayLayout.Appearance = Appearance126
        Me.grdSetup.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns
        UltraGridColumn1.Header.VisiblePosition = 0
        UltraGridColumn1.Hidden = True
        UltraGridColumn2.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always
        UltraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        Appearance127.TextHAlignAsString = "Left"
        UltraGridColumn2.CellAppearance = Appearance127
        UltraGridColumn2.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.CellSelect
        Appearance128.TextHAlignAsString = "Left"
        UltraGridColumn2.Header.Appearance = Appearance128
        UltraGridColumn2.Header.Caption = "Code"
        UltraGridColumn2.Header.VisiblePosition = 3
        UltraGridColumn2.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn2.RowLayoutColumnInfo.SpanY = 2
        UltraGridColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.EditButton
        UltraGridColumn2.Width = 140
        Appearance14.TextHAlignAsString = "Left"
        UltraGridColumn3.Header.Appearance = Appearance14
        UltraGridColumn3.Header.Caption = "Values"
        UltraGridColumn3.Header.VisiblePosition = 4
        UltraGridColumn3.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn3.RowLayoutColumnInfo.SpanY = 2
        UltraGridColumn3.Width = 269
        Appearance129.TextHAlignAsString = "Center"
        UltraGridColumn4.CellAppearance = Appearance129
        Appearance130.TextHAlignAsString = "Center"
        UltraGridColumn4.Header.Appearance = Appearance130
        UltraGridColumn4.Header.Caption = "Excl"
        UltraGridColumn4.Header.VisiblePosition = 6
        UltraGridColumn4.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn4.RowLayoutColumnInfo.SpanY = 2
        UltraGridColumn4.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn4.Width = 30
        UltraGridColumn5.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.None
        UltraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly
        Appearance131.TextHAlignAsString = "Center"
        UltraGridColumn5.CellAppearance = Appearance131
        Appearance132.TextHAlignAsString = "Center"
        UltraGridColumn5.Header.Appearance = Appearance132
        UltraGridColumn5.Header.Caption = "Seq"
        UltraGridColumn5.Header.VisiblePosition = 1
        UltraGridColumn5.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn5.RowLayoutColumnInfo.SpanY = 2
        UltraGridColumn5.Width = 14
        Appearance133.TextHAlignAsString = "Center"
        UltraGridColumn6.CellAppearance = Appearance133
        Appearance134.TextHAlignAsString = "Center"
        UltraGridColumn6.Header.Appearance = Appearance134
        UltraGridColumn6.Header.Caption = "Page"
        UltraGridColumn6.Header.VisiblePosition = 5
        UltraGridColumn6.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn6.RowLayoutColumnInfo.SpanY = 2
        UltraGridColumn6.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn6.Width = 39
        UltraGridColumn7.Header.VisiblePosition = 7
        UltraGridColumn7.Hidden = True
        Appearance135.TextHAlignAsString = "Center"
        UltraGridColumn8.CellAppearance = Appearance135
        Appearance136.TextHAlignAsString = "Center"
        UltraGridColumn8.Header.Appearance = Appearance136
        UltraGridColumn8.Header.Caption = "Grp"
        UltraGridColumn8.Header.VisiblePosition = 8
        UltraGridColumn8.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn8.RowLayoutColumnInfo.SpanY = 2
        UltraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn8.Width = 30
        UltraGridColumn9.Header.VisiblePosition = 9
        UltraGridColumn9.Hidden = True
        UltraGridColumn10.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always
        Appearance3.Image = CType(resources.GetObject("Appearance3.Image"), Object)
        Appearance3.ImageHAlign = Infragistics.Win.HAlign.Center
        UltraGridColumn10.CellButtonAppearance = Appearance3
        UltraGridColumn10.Header.Caption = ""
        UltraGridColumn10.Header.VisiblePosition = 2
        UltraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button
        UltraGridColumn10.Width = 14
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn1, UltraGridColumn2, UltraGridColumn3, UltraGridColumn4, UltraGridColumn5, UltraGridColumn6, UltraGridColumn7, UltraGridColumn8, UltraGridColumn9, UltraGridColumn10})
        Me.grdSetup.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdSetup.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance12.TextHAlignAsString = "Left"
        Me.grdSetup.DisplayLayout.CaptionAppearance = Appearance12
        Appearance137.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance137.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance137.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance137.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSetup.DisplayLayout.GroupByBox.Appearance = Appearance137
        Appearance138.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSetup.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance138
        Me.grdSetup.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdSetup.DisplayLayout.GroupByBox.Hidden = True
        Appearance139.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance139.BackColor2 = System.Drawing.SystemColors.Control
        Appearance139.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance139.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSetup.DisplayLayout.GroupByBox.PromptAppearance = Appearance139
        Me.grdSetup.DisplayLayout.MaxColScrollRegions = 1
        Me.grdSetup.DisplayLayout.MaxRowScrollRegions = 1
        Appearance140.BackColor = System.Drawing.SystemColors.Window
        Appearance140.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdSetup.DisplayLayout.Override.ActiveCellAppearance = Appearance140
        Me.grdSetup.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdSetup.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance141.BackColor = System.Drawing.SystemColors.Window
        Me.grdSetup.DisplayLayout.Override.CardAreaAppearance = Appearance141
        Appearance142.BorderColor = System.Drawing.Color.Silver
        Appearance142.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdSetup.DisplayLayout.Override.CellAppearance = Appearance142
        Me.grdSetup.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdSetup.DisplayLayout.Override.CellPadding = 0
        Appearance143.BackColor = System.Drawing.SystemColors.Control
        Appearance143.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance143.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance143.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance143.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSetup.DisplayLayout.Override.GroupByRowAppearance = Appearance143
        Me.grdSetup.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.ExternalSortSingle
        Me.grdSetup.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance145.BackColor = System.Drawing.SystemColors.Window
        Appearance145.BorderColor = System.Drawing.Color.Silver
        Me.grdSetup.DisplayLayout.Override.RowAppearance = Appearance145
        Me.grdSetup.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance146.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdSetup.DisplayLayout.Override.TemplateAddRowAppearance = Appearance146
        Me.grdSetup.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdSetup.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdSetup.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdSetup.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdSetup.Location = New System.Drawing.Point(0, 0)
        Me.grdSetup.Name = "grdSetup"
        Me.grdSetup.Size = New System.Drawing.Size(557, 361)
        Me.grdSetup.TabIndex = 1
        Me.grdSetup.Text = "Click Sort Icon to Sort, Double-Click Sort Heading to Clear Sort"
        '
        'ABSC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.grdSetup)
        Me.Name = "ABSC"
        Me.Size = New System.Drawing.Size(557, 361)
        CType(Me.grdSetup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents grdSetup As Infragistics.Win.UltraWinGrid.UltraGrid

End Class
