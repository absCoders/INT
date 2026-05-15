<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ASTATTA0
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
        Dim Appearance10 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance11 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance12 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance13 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance15 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.UltraTextEditor4 = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel4 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraTextEditor2 = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraTextEditor1 = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.Panel1.SuspendLayout()
        CType(Me.tbl, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).BeginInit()

        Me.ASFBASE1_Fill_Panel.SuspendLayout()
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).BeginInit()

        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.UltraTextEditor4)
        Me.Panel1.Controls.Add(Me.UltraLabel4)
        Me.Panel1.Controls.Add(Me.UltraTextEditor2)
        Me.Panel1.Controls.Add(Me.UltraLabel2)
        Me.Panel1.Controls.Add(Me.UltraTextEditor1)
        Me.Panel1.Controls.Add(Me.UltraLabel1)
        Me.Panel1.Controls.SetChildIndex(Me.UltraLabel1, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraTextEditor1, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraLabel2, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraTextEditor2, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraLabel4, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraTextEditor4, 0)
        '
        'UltraExplorerBar1
        '
        Me.UltraExplorerBar1.GroupSettings.UseMnemonics = Infragistics.Win.DefaultableBoolean.[True]
        Me.UltraExplorerBar1.ItemSettings.ForceSerialization = True
        Me.UltraExplorerBar1.Margins.ForceSerialization = True
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'grdASFBASEX
        '
        Appearance10.BackColor = System.Drawing.SystemColors.Window
        Appearance10.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdASFBASEX.DisplayLayout.Appearance = Appearance10
        Me.grdASFBASEX.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdASFBASEX.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.MaxColScrollRegions = 1
        Me.grdASFBASEX.DisplayLayout.MaxRowScrollRegions = 1
        Appearance11.BackColor = System.Drawing.SystemColors.Window
        Appearance11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveCellAppearance = Appearance11
        Appearance12.BackColor = System.Drawing.SystemColors.Highlight
        Appearance12.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveRowAppearance = Appearance12
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance13.BackColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.CardAreaAppearance = Appearance13
        Appearance14.BorderColor = System.Drawing.Color.Silver
        Appearance14.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdASFBASEX.DisplayLayout.Override.CellAppearance = Appearance14
        Me.grdASFBASEX.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdASFBASEX.DisplayLayout.Override.CellPadding = 0
        Appearance15.BackColor = System.Drawing.SystemColors.Control
        Appearance15.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance15.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance15.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.GroupByRowAppearance = Appearance15
        Appearance16.TextHAlignAsString = "Left"
        Me.grdASFBASEX.DisplayLayout.Override.HeaderAppearance = Appearance16
        Me.grdASFBASEX.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdASFBASEX.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance17.BackColor = System.Drawing.SystemColors.Window
        Appearance17.BorderColor = System.Drawing.Color.Silver
        Me.grdASFBASEX.DisplayLayout.Override.RowAppearance = Appearance17
        Me.grdASFBASEX.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance18.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdASFBASEX.DisplayLayout.Override.TemplateAddRowAppearance = Appearance18
        Me.grdASFBASEX.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdASFBASEX.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdASFBASEX.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        '
        'UltraTextEditor4
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor4, "COLUMN_NAME_DESC")
        Me.UltraTextEditor4.Location = New System.Drawing.Point(122, 127)
        Me.UltraTextEditor4.Name = "UltraTextEditor4"
        Me.UltraTextEditor4.Size = New System.Drawing.Size(486, 25)
        Me.UltraTextEditor4.TabIndex = 31
        '
        'UltraLabel4
        '
        Me.UltraLabel4.Location = New System.Drawing.Point(16, 130)
        Me.UltraLabel4.Name = "UltraLabel4"
        Me.UltraLabel4.Size = New System.Drawing.Size(100, 23)
        Me.UltraLabel4.TabIndex = 30
        Me.UltraLabel4.Text = "Description"
        '
        'UltraTextEditor2
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor2, "COLUMN_NAME")
        Me.Absx1.SetABSHasButton(Me.UltraTextEditor2, True)
        Me.UltraTextEditor2.Location = New System.Drawing.Point(122, 69)
        Me.UltraTextEditor2.Name = "UltraTextEditor2"
        Me.UltraTextEditor2.Size = New System.Drawing.Size(196, 25)
        Me.UltraTextEditor2.TabIndex = 27
        '
        'UltraLabel2
        '
        Me.UltraLabel2.Location = New System.Drawing.Point(16, 72)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(100, 23)
        Me.UltraLabel2.TabIndex = 26
        Me.UltraLabel2.Text = "Code Column"
        '
        'UltraTextEditor1
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor1, "TABLE_NAME")
        Me.Absx1.SetABSHasButton(Me.UltraTextEditor1, True)
        Me.UltraTextEditor1.Location = New System.Drawing.Point(122, 41)
        Me.UltraTextEditor1.Name = "UltraTextEditor1"
        Me.UltraTextEditor1.Size = New System.Drawing.Size(196, 25)
        Me.UltraTextEditor1.TabIndex = 25
        '
        'UltraLabel1
        '
        Me.UltraLabel1.Location = New System.Drawing.Point(16, 44)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(100, 23)
        Me.UltraLabel1.TabIndex = 24
        Me.UltraLabel1.Text = "Table Name"
        '
        'ASTATTA0
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(989, 574)
        Me.Name = "ASTATTA0"
        Me.Text = "ASTATTA0"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.tbl, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ASFBASE1_Fill_Panel.ResumeLayout(False)
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents UltraTextEditor4 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel4 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraTextEditor2 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraTextEditor1 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
End Class
