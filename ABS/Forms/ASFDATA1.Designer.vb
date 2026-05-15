<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ASFDATA1
    'Inherits System.Windows.Forms.Form
    Inherits ABSolution.ASFBASE2
    'Form overrides dispose to clean up the component list.
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
        Dim Appearance88 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1)
        Dim UltraGridColumn2 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TABLE_NAME", -1, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, False)
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TABLE_DESC")
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ROW_COUNT")
        Dim Appearance89 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance90 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance91 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance92 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance93 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance94 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance95 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance96 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance97 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance98 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance99 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance76 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance77 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance78 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance79 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance80 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance81 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance82 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance83 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance84 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance85 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance86 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance87 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.grdASTDATA1 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.grdData = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.ASFBASE2_Fill_Panel.SuspendLayout()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.grdASTDATA1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.SplitContainer1)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(976, 446)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 446)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(976, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 446)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(976, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 446)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(976, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.grdASTDATA1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.grdData)
        Me.SplitContainer1.Size = New System.Drawing.Size(976, 446)
        Me.SplitContainer1.SplitterDistance = 362
        Me.SplitContainer1.TabIndex = 36
        '
        'grdASTDATA1
        '
        Appearance88.BackColor = System.Drawing.SystemColors.Window
        Appearance88.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdASTDATA1.DisplayLayout.Appearance = Appearance88
        UltraGridColumn2.Header.Caption = "Table Name"
        UltraGridColumn2.Header.VisiblePosition = 0
        UltraGridColumn2.Width = 110
        UltraGridColumn3.Header.Caption = "Description"
        UltraGridColumn3.Header.VisiblePosition = 1
        UltraGridColumn3.Width = 150
        UltraGridColumn4.Header.Caption = "Rows"
        UltraGridColumn4.Header.VisiblePosition = 2
        UltraGridColumn4.Width = 51
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn2, UltraGridColumn3, UltraGridColumn4})
        Me.grdASTDATA1.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdASTDATA1.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance89.TextHAlignAsString = "Left"
        Me.grdASTDATA1.DisplayLayout.CaptionAppearance = Appearance89
        Appearance90.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance90.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance90.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance90.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASTDATA1.DisplayLayout.GroupByBox.Appearance = Appearance90
        Appearance91.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdASTDATA1.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance91
        Me.grdASTDATA1.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASTDATA1.DisplayLayout.GroupByBox.Hidden = True
        Appearance92.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance92.BackColor2 = System.Drawing.SystemColors.Control
        Appearance92.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance92.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdASTDATA1.DisplayLayout.GroupByBox.PromptAppearance = Appearance92
        Me.grdASTDATA1.DisplayLayout.MaxColScrollRegions = 1
        Me.grdASTDATA1.DisplayLayout.MaxRowScrollRegions = 1
        Appearance93.BackColor = System.Drawing.SystemColors.Window
        Appearance93.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdASTDATA1.DisplayLayout.Override.ActiveCellAppearance = Appearance93
        Me.grdASTDATA1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdASTDATA1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdASTDATA1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdASTDATA1.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdASTDATA1.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance94.BackColor = System.Drawing.SystemColors.Window
        Me.grdASTDATA1.DisplayLayout.Override.CardAreaAppearance = Appearance94
        Appearance95.BorderColor = System.Drawing.Color.Silver
        Appearance95.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdASTDATA1.DisplayLayout.Override.CellAppearance = Appearance95
        Me.grdASTDATA1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.CellSelect
        Me.grdASTDATA1.DisplayLayout.Override.CellPadding = 0
        Appearance96.BackColor = System.Drawing.SystemColors.Control
        Appearance96.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance96.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance96.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance96.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASTDATA1.DisplayLayout.Override.GroupByRowAppearance = Appearance96
        Appearance97.TextHAlignAsString = "Left"
        Me.grdASTDATA1.DisplayLayout.Override.HeaderAppearance = Appearance97
        Me.grdASTDATA1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdASTDATA1.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance98.BackColor = System.Drawing.SystemColors.Window
        Appearance98.BorderColor = System.Drawing.Color.Silver
        Me.grdASTDATA1.DisplayLayout.Override.RowAppearance = Appearance98
        Me.grdASTDATA1.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance99.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdASTDATA1.DisplayLayout.Override.TemplateAddRowAppearance = Appearance99
        Me.grdASTDATA1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdASTDATA1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdASTDATA1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdASTDATA1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdASTDATA1.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.grdASTDATA1.Location = New System.Drawing.Point(0, 0)
        Me.grdASTDATA1.Margin = New System.Windows.Forms.Padding(4)
        Me.grdASTDATA1.Name = "grdASTDATA1"
        Me.grdASTDATA1.Size = New System.Drawing.Size(362, 446)
        Me.grdASTDATA1.TabIndex = 0
        Me.grdASTDATA1.Text = "Tables"
        '
        'grdData
        '
        Appearance76.BackColor = System.Drawing.SystemColors.Window
        Appearance76.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdData.DisplayLayout.Appearance = Appearance76
        Me.grdData.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance77.TextHAlignAsString = "Left"
        Me.grdData.DisplayLayout.CaptionAppearance = Appearance77
        Appearance78.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance78.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance78.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance78.BorderColor = System.Drawing.SystemColors.Window
        Me.grdData.DisplayLayout.GroupByBox.Appearance = Appearance78
        Appearance79.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdData.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance79
        Me.grdData.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdData.DisplayLayout.GroupByBox.Hidden = True
        Appearance80.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance80.BackColor2 = System.Drawing.SystemColors.Control
        Appearance80.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance80.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdData.DisplayLayout.GroupByBox.PromptAppearance = Appearance80
        Me.grdData.DisplayLayout.MaxColScrollRegions = 1
        Me.grdData.DisplayLayout.MaxRowScrollRegions = 1
        Appearance81.BackColor = System.Drawing.SystemColors.Window
        Appearance81.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdData.DisplayLayout.Override.ActiveCellAppearance = Appearance81
        Me.grdData.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdData.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdData.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdData.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdData.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance82.BackColor = System.Drawing.SystemColors.Window
        Me.grdData.DisplayLayout.Override.CardAreaAppearance = Appearance82
        Appearance83.BorderColor = System.Drawing.Color.Silver
        Appearance83.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdData.DisplayLayout.Override.CellAppearance = Appearance83
        Me.grdData.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.CellSelect
        Me.grdData.DisplayLayout.Override.CellPadding = 0
        Appearance84.BackColor = System.Drawing.SystemColors.Control
        Appearance84.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance84.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance84.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance84.BorderColor = System.Drawing.SystemColors.Window
        Me.grdData.DisplayLayout.Override.GroupByRowAppearance = Appearance84
        Appearance85.TextHAlignAsString = "Left"
        Me.grdData.DisplayLayout.Override.HeaderAppearance = Appearance85
        Me.grdData.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdData.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance86.BackColor = System.Drawing.SystemColors.Window
        Appearance86.BorderColor = System.Drawing.Color.Silver
        Me.grdData.DisplayLayout.Override.RowAppearance = Appearance86
        Me.grdData.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance87.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdData.DisplayLayout.Override.TemplateAddRowAppearance = Appearance87
        Me.grdData.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdData.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdData.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand
        Me.grdData.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdData.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdData.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.grdData.Location = New System.Drawing.Point(0, 0)
        Me.grdData.Margin = New System.Windows.Forms.Padding(4)
        Me.grdData.Name = "grdData"
        Me.grdData.Size = New System.Drawing.Size(610, 446)
        Me.grdData.TabIndex = 1
        Me.grdData.Text = "Data"
        '
        'ASFDATA1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(976, 446)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "ASFDATA1"
        Me.Text = "ABS DataSet Export Utility"
        Me.ASFBASE2_Fill_Panel.ResumeLayout(False)
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.grdASTDATA1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents grdASTDATA1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents grdData As Infragistics.Win.UltraWinGrid.UltraGrid
End Class
