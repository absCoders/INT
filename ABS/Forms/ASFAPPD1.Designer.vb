<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ASFAPPD1
    'Inherits System.Windows.Forms.Form
    Inherits ABSolution.ASFBASE2

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim ValueListItem6 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem
        Dim ValueListItem7 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance15 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance11 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance13 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance12 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance10 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.cmdOK = New Infragistics.Win.Misc.UltraButton
        Me.cmdCancel = New Infragistics.Win.Misc.UltraButton
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.optCH = New Infragistics.Win.UltraWinEditors.UltraOptionSet
        Me.grdColumns = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pgdExplorerBar = New System.Windows.Forms.PropertyGrid
        Me.UltraColorPicker1 = New Infragistics.Win.UltraWinEditors.UltraColorPicker
        Me.grpAppearance = New Infragistics.Win.Misc.UltraGroupBox
        Me.ASFBASE2_Fill_Panel.SuspendLayout()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.optCH, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdColumns, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraColorPicker1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpAppearance, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpAppearance.SuspendLayout()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.SplitContainer1)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(639, 462)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 462)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(639, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 462)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(639, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 462)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(639, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'cmdOK
        '
        Me.cmdOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Appearance3.BackColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(230, Byte), Integer))
        Appearance3.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(165, Byte), Integer), CType(CType(165, Byte), Integer), CType(CType(150, Byte), Integer))
        Appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Me.cmdOK.Appearance = Appearance3
        Me.cmdOK.Location = New System.Drawing.Point(462, 16)
        Me.cmdOK.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(83, 26)
        Me.cmdOK.TabIndex = 2
        Me.cmdOK.Text = "OK"
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Appearance2.BackColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(230, Byte), Integer))
        Appearance2.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(165, Byte), Integer), CType(CType(165, Byte), Integer), CType(CType(150, Byte), Integer))
        Appearance2.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Me.cmdCancel.Appearance = Appearance2
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(551, 16)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(83, 26)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "Cancel"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel2
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.UltraColorPicker1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.optCH)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cmdOK)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cmdCancel)
        Me.SplitContainer1.Size = New System.Drawing.Size(639, 462)
        Me.SplitContainer1.SplitterDistance = 414
        Me.SplitContainer1.TabIndex = 4
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.grdColumns)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.grpAppearance)
        Me.SplitContainer2.Size = New System.Drawing.Size(639, 414)
        Me.SplitContainer2.SplitterDistance = 213
        Me.SplitContainer2.TabIndex = 13
        '
        'optCH
        '
        Me.optCH.BorderStyle = Infragistics.Win.UIElementBorderStyle.None
        Me.optCH.CheckedIndex = 1
        ValueListItem6.DataValue = "C"
        ValueListItem6.DisplayText = "Cell Appearance"
        ValueListItem7.DataValue = "H"
        ValueListItem7.DisplayText = "Header Appearance"
        Me.optCH.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem6, ValueListItem7})
        Me.optCH.Location = New System.Drawing.Point(3, 4)
        Me.optCH.Name = "optCH"
        Me.optCH.Size = New System.Drawing.Size(182, 40)
        Me.optCH.TabIndex = 12
        Me.optCH.Text = "Header Appearance"
        '
        'grdColumns
        '
        Appearance6.BackColor = System.Drawing.SystemColors.Window
        Appearance6.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdColumns.DisplayLayout.Appearance = Appearance6
        Me.grdColumns.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance15.TextHAlignAsString = "Left"
        Me.grdColumns.DisplayLayout.CaptionAppearance = Appearance15
        Appearance1.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance1.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance1.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance1.BorderColor = System.Drawing.SystemColors.Window
        Me.grdColumns.DisplayLayout.GroupByBox.Appearance = Appearance1
        Appearance4.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdColumns.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance4
        Me.grdColumns.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance5.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance5.BackColor2 = System.Drawing.SystemColors.Control
        Appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance5.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdColumns.DisplayLayout.GroupByBox.PromptAppearance = Appearance5
        Me.grdColumns.DisplayLayout.MaxColScrollRegions = 1
        Me.grdColumns.DisplayLayout.MaxRowScrollRegions = 1
        Appearance14.BackColor = System.Drawing.SystemColors.Window
        Appearance14.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdColumns.DisplayLayout.Override.ActiveCellAppearance = Appearance14
        Appearance9.BackColor = System.Drawing.SystemColors.Highlight
        Appearance9.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.grdColumns.DisplayLayout.Override.ActiveRowAppearance = Appearance9
        Me.grdColumns.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdColumns.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance8.BackColor = System.Drawing.SystemColors.Window
        Me.grdColumns.DisplayLayout.Override.CardAreaAppearance = Appearance8
        Appearance7.BorderColor = System.Drawing.Color.Silver
        Appearance7.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdColumns.DisplayLayout.Override.CellAppearance = Appearance7
        Me.grdColumns.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdColumns.DisplayLayout.Override.CellPadding = 0
        Appearance11.BackColor = System.Drawing.SystemColors.Control
        Appearance11.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance11.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance11.BorderColor = System.Drawing.SystemColors.Window
        Me.grdColumns.DisplayLayout.Override.GroupByRowAppearance = Appearance11
        Appearance13.TextHAlignAsString = "Left"
        Me.grdColumns.DisplayLayout.Override.HeaderAppearance = Appearance13
        Me.grdColumns.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdColumns.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance12.BackColor = System.Drawing.SystemColors.Window
        Appearance12.BorderColor = System.Drawing.Color.Silver
        Me.grdColumns.DisplayLayout.Override.RowAppearance = Appearance12
        Me.grdColumns.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance10.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdColumns.DisplayLayout.Override.TemplateAddRowAppearance = Appearance10
        Me.grdColumns.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdColumns.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdColumns.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdColumns.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdColumns.Location = New System.Drawing.Point(0, 0)
        Me.grdColumns.Name = "grdColumns"
        Me.grdColumns.Size = New System.Drawing.Size(213, 414)
        Me.grdColumns.TabIndex = 3
        Me.grdColumns.Text = "Grid Columns"
        '
        'pgdExplorerBar
        '
        Me.pgdExplorerBar.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pgdExplorerBar.Location = New System.Drawing.Point(3, 20)
        Me.pgdExplorerBar.Name = "pgdExplorerBar"
        Me.pgdExplorerBar.Size = New System.Drawing.Size(416, 391)
        Me.pgdExplorerBar.TabIndex = 2
        '
        'UltraColorPicker1
        '
        Me.UltraColorPicker1.Location = New System.Drawing.Point(454, 12)
        Me.UltraColorPicker1.Name = "UltraColorPicker1"
        Me.UltraColorPicker1.Size = New System.Drawing.Size(144, 25)
        Me.UltraColorPicker1.TabIndex = 0
        Me.UltraColorPicker1.Text = "Control"
        '
        'grpAppearance
        '
        Me.grpAppearance.Controls.Add(Me.pgdExplorerBar)
        Me.grpAppearance.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpAppearance.Location = New System.Drawing.Point(0, 0)
        Me.grpAppearance.Name = "grpAppearance"
        Me.grpAppearance.Size = New System.Drawing.Size(422, 414)
        Me.grpAppearance.TabIndex = 3
        Me.grpAppearance.Text = "UltraGroupBox1"
        '
        'ASFAPPD1
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(206, Byte), Integer))
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(639, 462)
        Me.ControlBox = False
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(349, 138)
        Me.Name = "ASFAPPD1"
        Me.Text = "Grid Appearance Designer"
        Me.ASFBASE2_Fill_Panel.ResumeLayout(False)
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.optCH, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdColumns, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraColorPicker1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpAppearance, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpAppearance.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdCancel As Infragistics.Win.Misc.UltraButton

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Friend WithEvents cmdOK As Infragistics.Win.Misc.UltraButton
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents UltraColorPicker1 As Infragistics.Win.UltraWinEditors.UltraColorPicker
    Public WithEvents pgdExplorerBar As System.Windows.Forms.PropertyGrid
    Friend WithEvents optCH As Infragistics.Win.UltraWinEditors.UltraOptionSet
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents grdColumns As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents grpAppearance As Infragistics.Win.Misc.UltraGroupBox
End Class
