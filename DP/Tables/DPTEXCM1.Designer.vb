<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DPTEXCM1
    'Inherits System.Windows.Forms.Form
    Inherits ABSolution.ASFCODEM
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
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance25 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance26 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance27 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance28 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance23 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance24 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance29 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("DPTEXCS1", -1)
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ITEM_CODE")
        Dim UltraGridColumn2 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EXC_MSG_CODE")
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EXC_SUPP_UNTIL")
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EXC_SUPP_USED")
        Dim UltraGridColumn5 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ITEM_DESC")
        Dim Appearance30 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance31 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance32 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance33 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance34 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance35 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance36 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance37 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance38 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance39 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance40 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance41 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraTextEditor5 = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraTextEditor1 = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.AbsCheckBox11 = New ABSCS.ABSCheckBox
        Me.grdDPTEXCS1 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Panel1.SuspendLayout()
        CType(Me.tbl, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ASFBASE1_Fill_Panel.SuspendLayout()
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdDPTEXCS1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.grdDPTEXCS1)
        Me.Panel1.Controls.Add(Me.AbsCheckBox11)
        Me.Panel1.Controls.Add(Me.UltraLabel2)
        Me.Panel1.Controls.Add(Me.UltraTextEditor5)
        Me.Panel1.Controls.Add(Me.UltraLabel1)
        Me.Panel1.Controls.Add(Me.UltraTextEditor1)
        Me.Panel1.Controls.SetChildIndex(Me.UltraTextEditor1, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraLabel1, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraTextEditor5, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraLabel2, 0)
        Me.Panel1.Controls.SetChildIndex(Me.AbsCheckBox11, 0)
        Me.Panel1.Controls.SetChildIndex(Me.grdDPTEXCS1, 0)
        '
        'UltraExplorerBar1
        '
        Me.UltraExplorerBar1.GroupSettings.UseMnemonics = Infragistics.Win.DefaultableBoolean.[True]
        Me.UltraExplorerBar1.ItemSettings.Style = Infragistics.Win.UltraWinExplorerBar.ItemStyle.Button
        Me.UltraExplorerBar1.Margins.ForceSerialization = True
        Me.UltraExplorerBar1.ShowDefaultContextMenu = False
        Me.UltraExplorerBar1.Size = New System.Drawing.Size(208, 554)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(989, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(989, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(989, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
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
        Appearance25.BackColor = System.Drawing.SystemColors.Control
        Appearance25.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance25.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance25.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.GroupByRowAppearance = Appearance25
        Appearance26.TextHAlignAsString = "Left"
        Me.grdASFBASEX.DisplayLayout.Override.HeaderAppearance = Appearance26
        Me.grdASFBASEX.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdASFBASEX.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance27.BackColor = System.Drawing.SystemColors.Window
        Appearance27.BorderColor = System.Drawing.Color.Silver
        Me.grdASFBASEX.DisplayLayout.Override.RowAppearance = Appearance27
        Me.grdASFBASEX.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance28.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdASFBASEX.DisplayLayout.Override.TemplateAddRowAppearance = Appearance28
        Me.grdASFBASEX.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdASFBASEX.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdASFBASEX.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        '
        'UltraLabel2
        '
        Appearance23.BackColor = System.Drawing.Color.Transparent
        Me.UltraLabel2.Appearance = Appearance23
        Me.UltraLabel2.AutoSize = True
        Me.UltraLabel2.Location = New System.Drawing.Point(16, 77)
        Me.UltraLabel2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(80, 18)
        Me.UltraLabel2.TabIndex = 32
        Me.UltraLabel2.Text = "Description"
        '
        'UltraTextEditor5
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor5, "EXC_MSG_DESC")
        Me.UltraTextEditor5.Location = New System.Drawing.Point(123, 77)
        Me.UltraTextEditor5.Name = "UltraTextEditor5"
        Me.UltraTextEditor5.Size = New System.Drawing.Size(289, 25)
        Me.UltraTextEditor5.TabIndex = 31
        '
        'UltraLabel1
        '
        Appearance24.BackColor = System.Drawing.Color.Transparent
        Me.UltraLabel1.Appearance = Appearance24
        Me.UltraLabel1.AutoSize = True
        Me.UltraLabel1.Location = New System.Drawing.Point(16, 45)
        Me.UltraLabel1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(39, 18)
        Me.UltraLabel1.TabIndex = 30
        Me.UltraLabel1.Text = "Code"
        '
        'UltraTextEditor1
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor1, "EXC_MSG_CODE")
        Me.Absx1.SetABSHasButton(Me.UltraTextEditor1, True)
        Me.UltraTextEditor1.Location = New System.Drawing.Point(123, 45)
        Me.UltraTextEditor1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UltraTextEditor1.Name = "UltraTextEditor1"
        Me.UltraTextEditor1.Size = New System.Drawing.Size(76, 25)
        Me.UltraTextEditor1.TabIndex = 29
        '
        'AbsCheckBox11
        '
        Me.Absx1.SetABSColumnName(Me.AbsCheckBox11, "EXC_MSG_SUPPRESS")
        Me.AbsCheckBox11.Location = New System.Drawing.Point(19, 108)
        Me.AbsCheckBox11.Name = "AbsCheckBox11"
        Me.AbsCheckBox11.Size = New System.Drawing.Size(95, 21)
        Me.AbsCheckBox11.TabIndex = 97
        Me.AbsCheckBox11.Text = "Suppress"
        '
        'grdDPTEXCS1
        '
        Me.grdDPTEXCS1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Appearance29.BackColor = System.Drawing.SystemColors.Window
        Appearance29.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdDPTEXCS1.DisplayLayout.Appearance = Appearance29
        UltraGridColumn1.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always
        UltraGridColumn1.Header.Caption = "Item Code"
        UltraGridColumn1.Header.VisiblePosition = 0
        UltraGridColumn1.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.EditButton
        UltraGridColumn2.Header.VisiblePosition = 2
        UltraGridColumn2.Hidden = True
        UltraGridColumn3.Header.Caption = "Until"
        UltraGridColumn3.Header.VisiblePosition = 3
        UltraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn4.Header.Caption = "Used"
        UltraGridColumn4.Header.VisiblePosition = 4
        UltraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn5.Header.Caption = "Description"
        UltraGridColumn5.Header.VisiblePosition = 1
        UltraGridColumn5.Width = 247
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn1, UltraGridColumn2, UltraGridColumn3, UltraGridColumn4, UltraGridColumn5})
        Me.grdDPTEXCS1.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdDPTEXCS1.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance30.TextHAlignAsString = "Left"
        Me.grdDPTEXCS1.DisplayLayout.CaptionAppearance = Appearance30
        Appearance31.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance31.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance31.BorderColor = System.Drawing.SystemColors.Window
        Me.grdDPTEXCS1.DisplayLayout.GroupByBox.Appearance = Appearance31
        Appearance32.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdDPTEXCS1.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance32
        Me.grdDPTEXCS1.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdDPTEXCS1.DisplayLayout.GroupByBox.Hidden = True
        Appearance33.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance33.BackColor2 = System.Drawing.SystemColors.Control
        Appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance33.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdDPTEXCS1.DisplayLayout.GroupByBox.PromptAppearance = Appearance33
        Me.grdDPTEXCS1.DisplayLayout.MaxColScrollRegions = 1
        Me.grdDPTEXCS1.DisplayLayout.MaxRowScrollRegions = 1
        Appearance34.BackColor = System.Drawing.SystemColors.Window
        Appearance34.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdDPTEXCS1.DisplayLayout.Override.ActiveCellAppearance = Appearance34
        Appearance35.BackColor = System.Drawing.SystemColors.Highlight
        Appearance35.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.grdDPTEXCS1.DisplayLayout.Override.ActiveRowAppearance = Appearance35
        Me.grdDPTEXCS1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.FixedAddRowOnTop
        Me.grdDPTEXCS1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[True]
        Me.grdDPTEXCS1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[True]
        Me.grdDPTEXCS1.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdDPTEXCS1.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance36.BackColor = System.Drawing.SystemColors.Window
        Me.grdDPTEXCS1.DisplayLayout.Override.CardAreaAppearance = Appearance36
        Appearance37.BorderColor = System.Drawing.Color.Silver
        Appearance37.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdDPTEXCS1.DisplayLayout.Override.CellAppearance = Appearance37
        Me.grdDPTEXCS1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdDPTEXCS1.DisplayLayout.Override.CellPadding = 0
        Appearance38.BackColor = System.Drawing.SystemColors.Control
        Appearance38.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance38.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance38.BorderColor = System.Drawing.SystemColors.Window
        Me.grdDPTEXCS1.DisplayLayout.Override.GroupByRowAppearance = Appearance38
        Appearance39.TextHAlignAsString = "Left"
        Me.grdDPTEXCS1.DisplayLayout.Override.HeaderAppearance = Appearance39
        Me.grdDPTEXCS1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdDPTEXCS1.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance40.BackColor = System.Drawing.SystemColors.Window
        Appearance40.BorderColor = System.Drawing.Color.Silver
        Me.grdDPTEXCS1.DisplayLayout.Override.RowAppearance = Appearance40
        Me.grdDPTEXCS1.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance41.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdDPTEXCS1.DisplayLayout.Override.TemplateAddRowAppearance = Appearance41
        Me.grdDPTEXCS1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdDPTEXCS1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdDPTEXCS1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdDPTEXCS1.Location = New System.Drawing.Point(19, 149)
        Me.grdDPTEXCS1.Name = "grdDPTEXCS1"
        Me.grdDPTEXCS1.Size = New System.Drawing.Size(736, 401)
        Me.grdDPTEXCS1.TabIndex = 98
        Me.grdDPTEXCS1.TabStop = False
        Me.grdDPTEXCS1.Text = "Suppress this Exception for the Following Items"
        '
        'DPTEXCM1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(989, 574)
        Me.Name = "DPTEXCM1"
        Me.Text = "DPTEXCM1"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.tbl, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ASFBASE1_Fill_Panel.ResumeLayout(False)
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdDPTEXCS1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraTextEditor5 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraTextEditor1 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents AbsCheckBox11 As ABSCS.ABSCheckBox
    Friend WithEvents grdDPTEXCS1 As Infragistics.Win.UltraWinGrid.UltraGrid
End Class
