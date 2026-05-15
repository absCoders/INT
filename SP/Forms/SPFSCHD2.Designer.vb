<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SPFSCHD2
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
        Dim Appearance22 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance13 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1)
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SCHED_CODE")
        Dim UltraGridColumn2 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SCHED_DESC")
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance15 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance19 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance21 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance23 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance24 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance25 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.cmdUpdate = New Infragistics.Win.Misc.UltraButton
        Me.cmdCancel = New Infragistics.Win.Misc.UltraButton
        Me.txtWH_OPER_ID = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.lblPosition = New Infragistics.Win.Misc.UltraLabel
        Me.txtWH_OPER_NAME = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.dteSCHED_DATE = New Infragistics.Win.UltraWinEditors.UltraDateTimeEditor
        Me.UltraLabel3 = New Infragistics.Win.Misc.UltraLabel
        Me.txtSCHED_NOTE = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraGroupBox8 = New Infragistics.Win.Misc.UltraGroupBox
        Me.cmbSCHED_CODE = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        Me.dteSCHED_DATE_END = New Infragistics.Win.UltraWinEditors.UltraDateTimeEditor
        Me.UltraGroupBox1 = New Infragistics.Win.Misc.UltraGroupBox
        Me.txtWH_OPER_GRP_DESC = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.txtWH_OPER_GRP = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.ASFBASE2_Fill_Panel.SuspendLayout()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtWH_OPER_ID, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtWH_OPER_NAME, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dteSCHED_DATE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSCHED_NOTE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGroupBox8, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox8.SuspendLayout()
        CType(Me.cmbSCHED_CODE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dteSCHED_DATE_END, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox1.SuspendLayout()
        CType(Me.txtWH_OPER_GRP_DESC, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtWH_OPER_GRP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.txtWH_OPER_GRP_DESC)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.txtWH_OPER_GRP)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.dteSCHED_DATE)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.dteSCHED_DATE_END)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.UltraGroupBox1)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.UltraLabel2)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.UltraGroupBox8)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.UltraLabel3)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.txtWH_OPER_NAME)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.lblPosition)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.txtWH_OPER_ID)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.cmdCancel)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.cmdUpdate)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(407, 389)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 389)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(407, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 389)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(407, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 389)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(407, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Location = New System.Drawing.Point(306, 351)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.Size = New System.Drawing.Size(92, 32)
        Me.cmdUpdate.TabIndex = 6
        Me.cmdUpdate.Text = "Update"
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(208, 351)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(92, 32)
        Me.cmdCancel.TabIndex = 7
        Me.cmdCancel.Text = "Cancel"
        '
        'txtWH_OPER_ID
        '
        Me.Absx1.SetABSColumnName(Me.txtWH_OPER_ID, "WH_OPER_ID")
        Me.Absx1.SetABSHasButton(Me.txtWH_OPER_ID, True)
        Me.Absx1.SetABSViewName(Me.txtWH_OPER_ID, "WH_OPER_ID")
        Me.txtWH_OPER_ID.Location = New System.Drawing.Point(12, 33)
        Me.txtWH_OPER_ID.Name = "txtWH_OPER_ID"
        Me.txtWH_OPER_ID.Size = New System.Drawing.Size(81, 25)
        Me.txtWH_OPER_ID.TabIndex = 0
        '
        'lblPosition
        '
        Me.lblPosition.AutoSize = True
        Me.lblPosition.Location = New System.Drawing.Point(12, 15)
        Me.lblPosition.Name = "lblPosition"
        Me.lblPosition.Size = New System.Drawing.Size(65, 18)
        Me.lblPosition.TabIndex = 3
        Me.lblPosition.Text = "Operator"
        '
        'txtWH_OPER_NAME
        '
        Me.Absx1.SetABSColumnName(Me.txtWH_OPER_NAME, "WH_OPER_NAME")
        Me.Absx1.SetABSParentColumnName(Me.txtWH_OPER_NAME, "WH_OPER_ID")
        Me.Absx1.SetABSViewName(Me.txtWH_OPER_NAME, "WH_OPER_ID")
        Me.txtWH_OPER_NAME.Location = New System.Drawing.Point(92, 33)
        Me.txtWH_OPER_NAME.Name = "txtWH_OPER_NAME"
        Me.txtWH_OPER_NAME.ReadOnly = True
        Me.txtWH_OPER_NAME.Size = New System.Drawing.Size(306, 25)
        Me.txtWH_OPER_NAME.TabIndex = 1
        '
        'dteSCHED_DATE
        '
        Me.Absx1.SetABSColumnName(Me.dteSCHED_DATE, "SCHED_DATE")
        Me.dteSCHED_DATE.Location = New System.Drawing.Point(93, 194)
        Me.dteSCHED_DATE.Name = "dteSCHED_DATE"
        Me.dteSCHED_DATE.Size = New System.Drawing.Size(115, 25)
        Me.dteSCHED_DATE.TabIndex = 3
        '
        'UltraLabel3
        '
        Me.UltraLabel3.AutoSize = True
        Me.UltraLabel3.Location = New System.Drawing.Point(12, 201)
        Me.UltraLabel3.Name = "UltraLabel3"
        Me.UltraLabel3.Size = New System.Drawing.Size(75, 18)
        Me.UltraLabel3.TabIndex = 7
        Me.UltraLabel3.Text = "Start Date"
        '
        'txtSCHED_NOTE
        '
        Me.Absx1.SetABSColumnName(Me.txtSCHED_NOTE, "SCHED_NOTE")
        Me.txtSCHED_NOTE.Location = New System.Drawing.Point(5, 20)
        Me.txtSCHED_NOTE.Multiline = True
        Me.txtSCHED_NOTE.Name = "txtSCHED_NOTE"
        Me.txtSCHED_NOTE.Scrollbars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtSCHED_NOTE.Size = New System.Drawing.Size(377, 62)
        Me.txtSCHED_NOTE.TabIndex = 5
        '
        'UltraGroupBox8
        '
        Me.UltraGroupBox8.BorderStyle = Infragistics.Win.Misc.GroupBoxBorderStyle.Rectangular3D
        Me.UltraGroupBox8.Controls.Add(Me.cmbSCHED_CODE)
        Me.UltraGroupBox8.Location = New System.Drawing.Point(12, 97)
        Me.UltraGroupBox8.Name = "UltraGroupBox8"
        Me.UltraGroupBox8.Size = New System.Drawing.Size(386, 65)
        Me.UltraGroupBox8.TabIndex = 73
        Me.UltraGroupBox8.Text = "Event"
        '
        'cmbSCHED_CODE
        '
        Me.Absx1.SetABSBindToTable(Me.cmbSCHED_CODE, False)
        Me.Absx1.SetABSColumnName(Me.cmbSCHED_CODE, "SCHED_CODE")
        Appearance22.BackColor = System.Drawing.Color.Beige
        Appearance22.BackColorDisabled = System.Drawing.Color.Beige
        Me.cmbSCHED_CODE.Appearance = Appearance22
        Me.cmbSCHED_CODE.CheckedListSettings.CheckStateMember = ""
        Appearance13.BackColor = System.Drawing.SystemColors.Window
        Appearance13.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.cmbSCHED_CODE.DisplayLayout.Appearance = Appearance13
        Me.cmbSCHED_CODE.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns
        UltraGridColumn1.Header.VisiblePosition = 0
        UltraGridColumn1.Hidden = True
        UltraGridColumn1.Width = 153
        UltraGridColumn2.Header.Caption = "Description"
        UltraGridColumn2.Header.VisiblePosition = 1
        UltraGridColumn2.Width = 288
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn1, UltraGridColumn2})
        UltraGridBand1.RowLayoutStyle = Infragistics.Win.UltraWinGrid.RowLayoutStyle.ColumnLayout
        Me.cmbSCHED_CODE.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.cmbSCHED_CODE.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.cmbSCHED_CODE.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance14.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance14.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance14.BorderColor = System.Drawing.SystemColors.Window
        Me.cmbSCHED_CODE.DisplayLayout.GroupByBox.Appearance = Appearance14
        Appearance15.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cmbSCHED_CODE.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance15
        Me.cmbSCHED_CODE.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance16.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance16.BackColor2 = System.Drawing.SystemColors.Control
        Appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance16.ForeColor = System.Drawing.SystemColors.GrayText
        Me.cmbSCHED_CODE.DisplayLayout.GroupByBox.PromptAppearance = Appearance16
        Me.cmbSCHED_CODE.DisplayLayout.MaxColScrollRegions = 1
        Me.cmbSCHED_CODE.DisplayLayout.MaxRowScrollRegions = 1
        Me.cmbSCHED_CODE.DisplayLayout.NewBandLoadStyle = Infragistics.Win.UltraWinGrid.NewBandLoadStyle.Hide
        Me.cmbSCHED_CODE.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance17.BackColor = System.Drawing.SystemColors.Window
        Appearance17.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmbSCHED_CODE.DisplayLayout.Override.ActiveCellAppearance = Appearance17
        Appearance18.BackColor = System.Drawing.SystemColors.Highlight
        Appearance18.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.cmbSCHED_CODE.DisplayLayout.Override.ActiveRowAppearance = Appearance18
        Me.cmbSCHED_CODE.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.cmbSCHED_CODE.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance19.BackColor = System.Drawing.SystemColors.Window
        Me.cmbSCHED_CODE.DisplayLayout.Override.CardAreaAppearance = Appearance19
        Appearance20.BorderColor = System.Drawing.Color.Silver
        Appearance20.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.cmbSCHED_CODE.DisplayLayout.Override.CellAppearance = Appearance20
        Me.cmbSCHED_CODE.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.cmbSCHED_CODE.DisplayLayout.Override.CellPadding = 0
        Appearance21.BackColor = System.Drawing.SystemColors.Control
        Appearance21.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance21.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance21.BorderColor = System.Drawing.SystemColors.Window
        Me.cmbSCHED_CODE.DisplayLayout.Override.GroupByRowAppearance = Appearance21
        Appearance23.TextHAlignAsString = "Left"
        Me.cmbSCHED_CODE.DisplayLayout.Override.HeaderAppearance = Appearance23
        Me.cmbSCHED_CODE.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.cmbSCHED_CODE.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance24.BackColor = System.Drawing.SystemColors.Window
        Appearance24.BorderColor = System.Drawing.Color.Silver
        Me.cmbSCHED_CODE.DisplayLayout.Override.RowAppearance = Appearance24
        Me.cmbSCHED_CODE.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance25.BackColor = System.Drawing.SystemColors.ControlLight
        Me.cmbSCHED_CODE.DisplayLayout.Override.TemplateAddRowAppearance = Appearance25
        Me.cmbSCHED_CODE.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.cmbSCHED_CODE.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.cmbSCHED_CODE.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.cmbSCHED_CODE.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList
        Me.cmbSCHED_CODE.Location = New System.Drawing.Point(12, 25)
        Me.cmbSCHED_CODE.Name = "cmbSCHED_CODE"
        Me.cmbSCHED_CODE.Size = New System.Drawing.Size(307, 26)
        Me.cmbSCHED_CODE.TabIndex = 13
        Me.cmbSCHED_CODE.TabStop = False
        '
        'UltraLabel2
        '
        Me.UltraLabel2.AutoSize = True
        Me.UltraLabel2.Location = New System.Drawing.Point(14, 231)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(66, 18)
        Me.UltraLabel2.TabIndex = 75
        Me.UltraLabel2.Text = "End Date"
        '
        'dteSCHED_DATE_END
        '
        Me.Absx1.SetABSColumnName(Me.dteSCHED_DATE_END, "SCHED_DATE_END")
        Me.dteSCHED_DATE_END.Location = New System.Drawing.Point(92, 224)
        Me.dteSCHED_DATE_END.Name = "dteSCHED_DATE_END"
        Me.dteSCHED_DATE_END.Size = New System.Drawing.Size(115, 25)
        Me.dteSCHED_DATE_END.TabIndex = 4
        '
        'UltraGroupBox1
        '
        Me.UltraGroupBox1.BorderStyle = Infragistics.Win.Misc.GroupBoxBorderStyle.Rectangular3D
        Me.UltraGroupBox1.Controls.Add(Me.txtSCHED_NOTE)
        Me.UltraGroupBox1.Location = New System.Drawing.Point(9, 258)
        Me.UltraGroupBox1.Name = "UltraGroupBox1"
        Me.UltraGroupBox1.Size = New System.Drawing.Size(389, 90)
        Me.UltraGroupBox1.TabIndex = 76
        Me.UltraGroupBox1.Text = "Note (Up to 100 characters)"
        '
        'txtWH_OPER_GRP_DESC
        '
        Me.Absx1.SetABSColumnName(Me.txtWH_OPER_GRP_DESC, "WH_OPER_GRP_DESC")
        Me.Absx1.SetABSParentColumnName(Me.txtWH_OPER_GRP_DESC, "WH_OPER_GRP")
        Me.Absx1.SetABSViewName(Me.txtWH_OPER_GRP_DESC, "WH_OPER_GRP")
        Me.txtWH_OPER_GRP_DESC.Location = New System.Drawing.Point(12, 57)
        Me.txtWH_OPER_GRP_DESC.Name = "txtWH_OPER_GRP_DESC"
        Me.txtWH_OPER_GRP_DESC.ReadOnly = True
        Me.txtWH_OPER_GRP_DESC.Size = New System.Drawing.Size(242, 25)
        Me.txtWH_OPER_GRP_DESC.TabIndex = 78
        '
        'txtWH_OPER_GRP
        '
        Me.Absx1.SetABSColumnName(Me.txtWH_OPER_GRP, "WH_OPER_GRP")
        Me.Absx1.SetABSHasButton(Me.txtWH_OPER_GRP, True)
        Me.Absx1.SetABSParentColumnName(Me.txtWH_OPER_GRP, "WH_OPER_ID")
        Me.Absx1.SetABSViewName(Me.txtWH_OPER_GRP, "WH_OPER_GRP")
        Me.txtWH_OPER_GRP.Location = New System.Drawing.Point(12, 57)
        Me.txtWH_OPER_GRP.Name = "txtWH_OPER_GRP"
        Me.txtWH_OPER_GRP.ReadOnly = True
        Me.txtWH_OPER_GRP.Size = New System.Drawing.Size(81, 25)
        Me.txtWH_OPER_GRP.TabIndex = 77
        Me.txtWH_OPER_GRP.Visible = False
        '
        'SHFSCHD2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(407, 389)
        Me.ControlBox = False
        Me.Name = "SHFSCHD2"
        Me.Text = "Schedule an Event"
        Me.ASFBASE2_Fill_Panel.ResumeLayout(False)
        Me.ASFBASE2_Fill_Panel.PerformLayout()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtWH_OPER_ID, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtWH_OPER_NAME, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dteSCHED_DATE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSCHED_NOTE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGroupBox8, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox8.ResumeLayout(False)
        Me.UltraGroupBox8.PerformLayout()
        CType(Me.cmbSCHED_CODE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dteSCHED_DATE_END, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox1.ResumeLayout(False)
        Me.UltraGroupBox1.PerformLayout()
        CType(Me.txtWH_OPER_GRP_DESC, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtWH_OPER_GRP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdUpdate As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents dteSCHED_DATE As Infragistics.Win.UltraWinEditors.UltraDateTimeEditor
    Friend WithEvents txtWH_OPER_NAME As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents lblPosition As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtWH_OPER_ID As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel3 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtSCHED_NOTE As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents dteSCHED_DATE_END As Infragistics.Win.UltraWinEditors.UltraDateTimeEditor
    Friend WithEvents UltraGroupBox8 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraGroupBox1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtWH_OPER_GRP_DESC As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents txtWH_OPER_GRP As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents cmbSCHED_CODE As Infragistics.Win.UltraWinGrid.UltraCombo
End Class
