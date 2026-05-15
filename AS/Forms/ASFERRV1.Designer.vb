<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ASFERRV1
    'Inherits System.Windows.Forms.Form
    Inherits ABSolution.ASFBASE1
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
        Dim UltraExplorerBarGroup1 As Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarGroup = New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarGroup
        Dim UltraExplorerBarItem1 As Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem = New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem
        Dim UltraExplorerBarItem2 As Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem = New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance10 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1)
        Dim Appearance11 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance12 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance13 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance15 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance19 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance21 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance22 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.UltraGroupBox1 = New Infragistics.Win.Misc.UltraGroupBox
        Me.grdASTERROR = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.grpERR_TEXT = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtSTACKTRACE = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.splASFBASE1 = New System.Windows.Forms.SplitContainer
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ASFBASE1_Fill_Panel.SuspendLayout()
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdASTERROR, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.grpERR_TEXT, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpERR_TEXT.SuspendLayout()
        CType(Me.txtSTACKTRACE, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splASFBASE1.Panel1.SuspendLayout()
        Me.splASFBASE1.Panel2.SuspendLayout()
        Me.splASFBASE1.SuspendLayout()
        Me.SuspendLayout()
        '
        'UltraExplorerBar1
        '
        UltraExplorerBarItem1.Text = "Load Errors"
        UltraExplorerBarItem2.Text = "Done"
        UltraExplorerBarGroup1.Items.AddRange(New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem() {UltraExplorerBarItem1, UltraExplorerBarItem2})
        UltraExplorerBarGroup1.Key = "Screen Control"
        UltraExplorerBarGroup1.Text = "Screen Control"
        Me.UltraExplorerBar1.Groups.AddRange(New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarGroup() {UltraExplorerBarGroup1})
        Me.UltraExplorerBar1.GroupSettings.UseMnemonics = Infragistics.Win.DefaultableBoolean.[True]
        Me.UltraExplorerBar1.ItemSettings.Style = Infragistics.Win.UltraWinExplorerBar.ItemStyle.Button
        Me.UltraExplorerBar1.Margins.Bottom = 0
        Me.UltraExplorerBar1.Margins.Left = 0
        Me.UltraExplorerBar1.Margins.Right = 0
        Me.UltraExplorerBar1.Margins.Top = 0
        Me.UltraExplorerBar1.ShowDefaultContextMenu = False
        Me.UltraExplorerBar1.Size = New System.Drawing.Size(208, 554)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 574)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(989, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 574)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(989, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 574)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(989, 0)
        '
        'ASFBASE1_Fill_Panel
        '
        Me.ASFBASE1_Fill_Panel.Controls.Add(Me.splASFBASE1)
        Me.ASFBASE1_Fill_Panel.Size = New System.Drawing.Size(776, 574)
        Me.ASFBASE1_Fill_Panel.Controls.SetChildIndex(Me.grdASFBASEX, 0)
        Me.ASFBASE1_Fill_Panel.Controls.SetChildIndex(Me.splASFBASE1, 0)
        '
        'grdASFBASEX
        '
        Appearance1.BackColor = System.Drawing.SystemColors.Window
        Appearance1.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdASFBASEX.DisplayLayout.Appearance = Appearance1
        Me.grdASFBASEX.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdASFBASEX.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.MaxColScrollRegions = 1
        Me.grdASFBASEX.DisplayLayout.MaxRowScrollRegions = 1
        Appearance2.BackColor = System.Drawing.SystemColors.Window
        Appearance2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveCellAppearance = Appearance2
        Appearance3.BackColor = System.Drawing.SystemColors.Highlight
        Appearance3.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveRowAppearance = Appearance3
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance4.BackColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.CardAreaAppearance = Appearance4
        Appearance5.BorderColor = System.Drawing.Color.Silver
        Appearance5.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdASFBASEX.DisplayLayout.Override.CellAppearance = Appearance5
        Me.grdASFBASEX.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdASFBASEX.DisplayLayout.Override.CellPadding = 0
        Appearance6.BackColor = System.Drawing.SystemColors.Control
        Appearance6.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance6.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance6.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.GroupByRowAppearance = Appearance6
        Appearance7.TextHAlignAsString = "Left"
        Me.grdASFBASEX.DisplayLayout.Override.HeaderAppearance = Appearance7
        Me.grdASFBASEX.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdASFBASEX.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance8.BackColor = System.Drawing.SystemColors.Window
        Appearance8.BorderColor = System.Drawing.Color.Silver
        Me.grdASFBASEX.DisplayLayout.Override.RowAppearance = Appearance8
        Me.grdASFBASEX.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance9.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdASFBASEX.DisplayLayout.Override.TemplateAddRowAppearance = Appearance9
        Me.grdASFBASEX.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdASFBASEX.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdASFBASEX.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'UltraGroupBox1
        '
        Me.UltraGroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UltraGroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.UltraGroupBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.UltraGroupBox1.Name = "UltraGroupBox1"
        Me.UltraGroupBox1.Size = New System.Drawing.Size(776, 67)
        Me.UltraGroupBox1.TabIndex = 1
        '
        'grdASTERROR
        '
        Appearance10.BackColor = System.Drawing.SystemColors.Window
        Appearance10.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdASTERROR.DisplayLayout.Appearance = Appearance10
        UltraGridBand1.RowLayoutStyle = UltraWinGrid.RowLayoutStyle.ColumnLayout
        Me.grdASTERROR.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdASTERROR.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance11.TextHAlignAsString = "Left"
        Me.grdASTERROR.DisplayLayout.CaptionAppearance = Appearance11
        Appearance12.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance12.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance12.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASTERROR.DisplayLayout.GroupByBox.Appearance = Appearance12
        Appearance13.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdASTERROR.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance13
        Me.grdASTERROR.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASTERROR.DisplayLayout.GroupByBox.Hidden = True
        Appearance14.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance14.BackColor2 = System.Drawing.SystemColors.Control
        Appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance14.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdASTERROR.DisplayLayout.GroupByBox.PromptAppearance = Appearance14
        Me.grdASTERROR.DisplayLayout.MaxColScrollRegions = 1
        Me.grdASTERROR.DisplayLayout.MaxRowScrollRegions = 1
        Appearance15.BackColor = System.Drawing.SystemColors.Window
        Appearance15.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdASTERROR.DisplayLayout.Override.ActiveCellAppearance = Appearance15
        Me.grdASTERROR.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdASTERROR.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdASTERROR.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdASTERROR.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdASTERROR.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance17.BackColor = System.Drawing.SystemColors.Window
        Me.grdASTERROR.DisplayLayout.Override.CardAreaAppearance = Appearance17
        Appearance18.BorderColor = System.Drawing.Color.Silver
        Appearance18.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdASTERROR.DisplayLayout.Override.CellAppearance = Appearance18
        Me.grdASTERROR.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.CellSelect
        Me.grdASTERROR.DisplayLayout.Override.CellPadding = 0
        Appearance19.BackColor = System.Drawing.SystemColors.Control
        Appearance19.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance19.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance19.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASTERROR.DisplayLayout.Override.GroupByRowAppearance = Appearance19
        Appearance20.TextHAlignAsString = "Left"
        Me.grdASTERROR.DisplayLayout.Override.HeaderAppearance = Appearance20
        Me.grdASTERROR.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdASTERROR.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance21.BackColor = System.Drawing.SystemColors.Window
        Appearance21.BorderColor = System.Drawing.Color.Silver
        Me.grdASTERROR.DisplayLayout.Override.RowAppearance = Appearance21
        Me.grdASTERROR.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance22.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdASTERROR.DisplayLayout.Override.TemplateAddRowAppearance = Appearance22
        Me.grdASTERROR.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdASTERROR.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdASTERROR.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand
        Me.grdASTERROR.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdASTERROR.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdASTERROR.Location = New System.Drawing.Point(0, 0)
        Me.grdASTERROR.Margin = New System.Windows.Forms.Padding(2)
        Me.grdASTERROR.Name = "grdASTERROR"
        Me.grdASTERROR.Size = New System.Drawing.Size(776, 328)
        Me.grdASTERROR.TabIndex = 2
        Me.grdASTERROR.Text = "Application Errors"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.grdASTERROR)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.grpERR_TEXT)
        Me.SplitContainer1.Size = New System.Drawing.Size(776, 503)
        Me.SplitContainer1.SplitterDistance = 328
        Me.SplitContainer1.TabIndex = 3
        '
        'grpERR_TEXT
        '
        Me.grpERR_TEXT.Controls.Add(Me.txtSTACKTRACE)
        Me.grpERR_TEXT.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpERR_TEXT.Location = New System.Drawing.Point(0, 0)
        Me.grpERR_TEXT.Name = "grpERR_TEXT"
        Me.grpERR_TEXT.Size = New System.Drawing.Size(776, 171)
        Me.grpERR_TEXT.TabIndex = 1
        '
        'txtSTACKTRACE
        '
        Me.Absx1.SetABSColumnName(Me.txtSTACKTRACE, "STACKTRACE")
        Me.txtSTACKTRACE.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtSTACKTRACE.Location = New System.Drawing.Point(3, 0)
        Me.txtSTACKTRACE.Multiline = True
        Me.txtSTACKTRACE.Name = "txtSTACKTRACE"
        Me.txtSTACKTRACE.ReadOnly = True
        Me.txtSTACKTRACE.Scrollbars = System.Windows.Forms.ScrollBars.Both
        Me.txtSTACKTRACE.Size = New System.Drawing.Size(770, 168)
        Me.txtSTACKTRACE.TabIndex = 2
        '
        'splASFBASE1
        '
        Me.splASFBASE1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splASFBASE1.Location = New System.Drawing.Point(0, 0)
        Me.splASFBASE1.Name = "splASFBASE1"
        Me.splASFBASE1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'splASFBASE1.Panel1
        '
        Me.splASFBASE1.Panel1.Controls.Add(Me.UltraGroupBox1)
        '
        'splASFBASE1.Panel2
        '
        Me.splASFBASE1.Panel2.Controls.Add(Me.SplitContainer1)
        Me.splASFBASE1.Size = New System.Drawing.Size(776, 574)
        Me.splASFBASE1.SplitterDistance = 67
        Me.splASFBASE1.TabIndex = 4
        '
        'ASFERRV1
        '
        Me.Absx1.SetABSBindToTable(Me, False)
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(989, 574)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "ASFERRV1"
        Me.Text = "ASFERRV1"
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ASFBASE1_Fill_Panel.ResumeLayout(False)
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdASTERROR, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.grpERR_TEXT, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpERR_TEXT.ResumeLayout(False)
        Me.grpERR_TEXT.PerformLayout()
        CType(Me.txtSTACKTRACE, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splASFBASE1.Panel1.ResumeLayout(False)
        Me.splASFBASE1.Panel2.ResumeLayout(False)
        Me.splASFBASE1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraGroupBox1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents grdASTERROR As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents splASFBASE1 As System.Windows.Forms.SplitContainer
    Friend WithEvents grpERR_TEXT As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtSTACKTRACE As Infragistics.Win.UltraWinEditors.UltraTextEditor
End Class
