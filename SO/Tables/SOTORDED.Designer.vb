<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class SOTORDED
    'Inherits System.Windows.Forms.Form
    Inherits ABSolution.ASFCODEM
    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance10 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim ValueListItem1 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem2 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.UltraOptionSet1 = New Infragistics.Win.UltraWinEditors.UltraOptionSet()
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraTextEditor1 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.txtCUST_NAME = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.lblCUST_CODE = New Infragistics.Win.Misc.UltraLabel()
        Me.txtCUST_CODE = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.Panel1.SuspendLayout()
        CType(Me.tbl, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ASFBASE1_Fill_Panel.SuspendLayout()
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTLOGX1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraOptionSet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCUST_NAME, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCUST_CODE, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.txtCUST_NAME)
        Me.Panel1.Controls.Add(Me.lblCUST_CODE)
        Me.Panel1.Controls.Add(Me.txtCUST_CODE)
        Me.Panel1.Controls.Add(Me.UltraOptionSet1)
        Me.Panel1.Controls.Add(Me.UltraLabel1)
        Me.Panel1.Controls.Add(Me.UltraTextEditor1)
        Me.Panel1.Size = New System.Drawing.Size(772, 507)
        '
        'UltraExplorerBar1
        '
        Me.UltraExplorerBar1.GroupSettings.UseMnemonics = Infragistics.Win.DefaultableBoolean.[True]
        Me.UltraExplorerBar1.ItemSettings.Style = Infragistics.Win.UltraWinExplorerBar.ItemStyle.Button
        Me.UltraExplorerBar1.Margins.Bottom = 0
        Me.UltraExplorerBar1.Margins.Left = 0
        Me.UltraExplorerBar1.Margins.Right = 0
        Me.UltraExplorerBar1.Margins.Top = 0
        Me.UltraExplorerBar1.Size = New System.Drawing.Size(208, 554)
        '
        'ASFBASE1_Fill_Panel
        '
        Me.ASFBASE1_Fill_Panel.Size = New System.Drawing.Size(776, 574)
        '
        'grdASFBASEX
        '
        Appearance2.BackColor = System.Drawing.SystemColors.Window
        Appearance2.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdASFBASEX.DisplayLayout.Appearance = Appearance2
        Me.grdASFBASEX.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdASFBASEX.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.MaxColScrollRegions = 1
        Me.grdASFBASEX.DisplayLayout.MaxRowScrollRegions = 1
        Appearance3.BackColor = System.Drawing.SystemColors.Window
        Appearance3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveCellAppearance = Appearance3
        Appearance4.BackColor = System.Drawing.SystemColors.Highlight
        Appearance4.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveRowAppearance = Appearance4
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance5.BackColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.CardAreaAppearance = Appearance5
        Appearance6.BorderColor = System.Drawing.Color.Silver
        Appearance6.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdASFBASEX.DisplayLayout.Override.CellAppearance = Appearance6
        Me.grdASFBASEX.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdASFBASEX.DisplayLayout.Override.CellPadding = 0
        Appearance7.BackColor = System.Drawing.SystemColors.Control
        Appearance7.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance7.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance7.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.GroupByRowAppearance = Appearance7
        Appearance8.TextHAlignAsString = "Left"
        Me.grdASFBASEX.DisplayLayout.Override.HeaderAppearance = Appearance8
        Me.grdASFBASEX.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdASFBASEX.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance9.BackColor = System.Drawing.SystemColors.Window
        Appearance9.BorderColor = System.Drawing.Color.Silver
        Me.grdASFBASEX.DisplayLayout.Override.RowAppearance = Appearance9
        Me.grdASFBASEX.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance10.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdASFBASEX.DisplayLayout.Override.TemplateAddRowAppearance = Appearance10
        Me.grdASFBASEX.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdASFBASEX.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdASFBASEX.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
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
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'UltraOptionSet1
        '
        Me.Absx1.SetABSColumnName(Me.UltraOptionSet1, "EDOC_DOMAIN_STATUS")
        Me.UltraOptionSet1.BorderStyle = Infragistics.Win.UIElementBorderStyle.None
        ValueListItem1.DataValue = "A"
        ValueListItem1.DisplayText = "Active"
        ValueListItem2.DataValue = "I"
        ValueListItem2.DisplayText = "In-Active"
        Me.UltraOptionSet1.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem1, ValueListItem2})
        Me.UltraOptionSet1.Location = New System.Drawing.Point(334, 43)
        Me.UltraOptionSet1.Name = "UltraOptionSet1"
        Me.UltraOptionSet1.Size = New System.Drawing.Size(144, 20)
        Me.UltraOptionSet1.TabIndex = 123
        '
        'UltraLabel1
        '
        Appearance1.BackColor = System.Drawing.Color.Transparent
        Me.UltraLabel1.Appearance = Appearance1
        Me.UltraLabel1.AutoSize = True
        Me.UltraLabel1.Location = New System.Drawing.Point(16, 45)
        Me.UltraLabel1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(56, 18)
        Me.UltraLabel1.TabIndex = 104
        Me.UltraLabel1.Text = "Domain"
        '
        'UltraTextEditor1
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor1, "EDOC_DOMAIN_NAME")
        Me.Absx1.SetABSHasButton(Me.UltraTextEditor1, True)
        Me.UltraTextEditor1.Location = New System.Drawing.Point(130, 42)
        Me.UltraTextEditor1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UltraTextEditor1.Name = "UltraTextEditor1"
        Me.UltraTextEditor1.Size = New System.Drawing.Size(150, 25)
        Me.UltraTextEditor1.TabIndex = 102
        '
        'txtCUST_NAME
        '
        Me.Absx1.SetABSColumnName(Me.txtCUST_NAME, "CUST_NAME")
        Me.Absx1.SetABSParentColumnName(Me.txtCUST_NAME, "CUST_CODE")
        Me.txtCUST_NAME.Location = New System.Drawing.Point(232, 120)
        Me.txtCUST_NAME.Name = "txtCUST_NAME"
        Me.txtCUST_NAME.ReadOnly = True
        Me.txtCUST_NAME.Size = New System.Drawing.Size(289, 25)
        Me.txtCUST_NAME.TabIndex = 154
        '
        'lblCUST_CODE
        '
        Me.lblCUST_CODE.Location = New System.Drawing.Point(16, 120)
        Me.lblCUST_CODE.Name = "lblCUST_CODE"
        Me.lblCUST_CODE.Size = New System.Drawing.Size(108, 45)
        Me.lblCUST_CODE.TabIndex = 153
        Me.lblCUST_CODE.Text = "Customer"
        '
        'txtCUST_CODE
        '
        Me.Absx1.SetABSColumnName(Me.txtCUST_CODE, "CUST_CODE")
        Me.Absx1.SetABSHasButton(Me.txtCUST_CODE, True)
        Me.txtCUST_CODE.Location = New System.Drawing.Point(130, 120)
        Me.txtCUST_CODE.Name = "txtCUST_CODE"
        Me.txtCUST_CODE.Size = New System.Drawing.Size(96, 25)
        Me.txtCUST_CODE.TabIndex = 152
        '
        'SOTORDED
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(989, 574)
        Me.Name = "SOTORDED"
        Me.Text = "SOTORDED"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.tbl, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ASFBASE1_Fill_Panel.ResumeLayout(False)
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTLOGX1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraOptionSet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCUST_NAME, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCUST_CODE, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraOptionSet1 As Infragistics.Win.UltraWinEditors.UltraOptionSet
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraTextEditor1 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents txtCUST_NAME As UltraWinEditors.UltraTextEditor
    Friend WithEvents lblCUST_CODE As Misc.UltraLabel
    Friend WithEvents txtCUST_CODE As UltraWinEditors.UltraTextEditor
End Class
