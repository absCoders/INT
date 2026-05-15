<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TAFLOCM1
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
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("WHTMOVE2", -1)
        Dim UltraGridColumn2 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WHSE_TRAN_NO")
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WHSE_TRAN_LNO")
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LOCATION_CODE_FROM")
        Dim UltraGridColumn9 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LOCATION_CODE_TO", -1, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, False)
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridColumn12 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("BAR_CODE")
        Dim UltraGridColumn13 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WHSE_TRAN_QTY")
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridColumn32 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ITEM_CODE")
        Dim UltraGridColumn14 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_OPER")
        Dim UltraGridColumn15 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_DATE")
        Dim UltraGridColumn16 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_OPER")
        Dim UltraGridColumn17 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_DATE")
        Dim UltraGridColumn18 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STATUS")
        Dim UltraGridColumn20 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ITEM_DESC")
        Dim UltraGridColumn21 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WHSE_TRAN_QTY_ORIG")
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ERROR_CODES")
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance10 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance11 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance12 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance13 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance15 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.splItems = New System.Windows.Forms.SplitContainer()
        Me.grpHeader = New Infragistics.Win.Misc.UltraGroupBox()
        Me.UltraGroupBox2 = New Infragistics.Win.Misc.UltraGroupBox()
        Me.txtWHSE_TRAN_NO = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.grpTo = New Infragistics.Win.Misc.UltraGroupBox()
        Me.UltraTextEditor3 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.grdWHTMOVE2 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.UltraGroupBox1 = New Infragistics.Win.Misc.UltraGroupBox()
        Me.btnMove = New Infragistics.Win.Misc.UltraButton()
        Me.btnCancel = New Infragistics.Win.Misc.UltraButton()
        Me.ASFBASE2_Fill_Panel.SuspendLayout()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.splItems, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splItems.Panel1.SuspendLayout()
        Me.splItems.Panel2.SuspendLayout()
        Me.splItems.SuspendLayout()
        CType(Me.grpHeader, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpHeader.SuspendLayout()
        CType(Me.UltraGroupBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox2.SuspendLayout()
        CType(Me.txtWHSE_TRAN_NO, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpTo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpTo.SuspendLayout()
        CType(Me.UltraTextEditor3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.grdWHTMOVE2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.splItems)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(735, 431)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Margin = New System.Windows.Forms.Padding(4)
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 431)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(735, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Margin = New System.Windows.Forms.Padding(4)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 431)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Margin = New System.Windows.Forms.Padding(4)
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(735, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 431)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Margin = New System.Windows.Forms.Padding(4)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(735, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'splItems
        '
        Me.splItems.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splItems.ImeMode = System.Windows.Forms.ImeMode.[On]
        Me.splItems.IsSplitterFixed = True
        Me.splItems.Location = New System.Drawing.Point(0, 0)
        Me.splItems.Name = "splItems"
        Me.splItems.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'splItems.Panel1
        '
        Me.splItems.Panel1.Controls.Add(Me.grpHeader)
        '
        'splItems.Panel2
        '
        Me.splItems.Panel2.Controls.Add(Me.SplitContainer1)
        Me.splItems.Size = New System.Drawing.Size(735, 431)
        Me.splItems.SplitterDistance = 69
        Me.splItems.TabIndex = 0
        '
        'grpHeader
        '
        Me.grpHeader.Controls.Add(Me.UltraGroupBox2)
        Me.grpHeader.Controls.Add(Me.grpTo)
        Me.grpHeader.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpHeader.Location = New System.Drawing.Point(0, 0)
        Me.grpHeader.Name = "grpHeader"
        Me.grpHeader.Size = New System.Drawing.Size(735, 69)
        Me.grpHeader.TabIndex = 0
        '
        'UltraGroupBox2
        '
        Me.UltraGroupBox2.BorderStyle = Infragistics.Win.Misc.GroupBoxBorderStyle.Rectangular3D
        Me.UltraGroupBox2.Controls.Add(Me.txtWHSE_TRAN_NO)
        Me.UltraGroupBox2.Location = New System.Drawing.Point(6, 3)
        Me.UltraGroupBox2.Name = "UltraGroupBox2"
        Me.UltraGroupBox2.Size = New System.Drawing.Size(142, 55)
        Me.UltraGroupBox2.TabIndex = 122
        Me.UltraGroupBox2.Text = "Trans No"
        '
        'txtWHSE_TRAN_NO
        '
        Me.Absx1.SetABSBindToTable(Me.txtWHSE_TRAN_NO, False)
        Appearance17.BackColor = System.Drawing.Color.White
        Me.txtWHSE_TRAN_NO.Appearance = Appearance17
        Me.txtWHSE_TRAN_NO.BackColor = System.Drawing.Color.White
        Me.txtWHSE_TRAN_NO.Location = New System.Drawing.Point(6, 21)
        Me.txtWHSE_TRAN_NO.Name = "txtWHSE_TRAN_NO"
        Me.txtWHSE_TRAN_NO.ReadOnly = True
        Me.txtWHSE_TRAN_NO.Size = New System.Drawing.Size(130, 25)
        Me.txtWHSE_TRAN_NO.TabIndex = 1
        '
        'grpTo
        '
        Me.grpTo.BorderStyle = Infragistics.Win.Misc.GroupBoxBorderStyle.Rectangular3D
        Me.grpTo.Controls.Add(Me.UltraTextEditor3)
        Me.grpTo.Location = New System.Drawing.Point(567, 3)
        Me.grpTo.Name = "grpTo"
        Me.grpTo.Size = New System.Drawing.Size(132, 55)
        Me.grpTo.TabIndex = 120
        Me.grpTo.Text = "Move To"
        '
        'UltraTextEditor3
        '
        Me.Absx1.SetABSBindToTable(Me.UltraTextEditor3, False)
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor3, "LOCATION_CODE")
        Me.Absx1.SetABSHasButton(Me.UltraTextEditor3, True)
        Me.UltraTextEditor3.Location = New System.Drawing.Point(10, 21)
        Me.UltraTextEditor3.Name = "UltraTextEditor3"
        Me.UltraTextEditor3.Size = New System.Drawing.Size(114, 25)
        Me.UltraTextEditor3.TabIndex = 0
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.grdWHTMOVE2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.UltraGroupBox1)
        Me.SplitContainer1.Size = New System.Drawing.Size(735, 358)
        Me.SplitContainer1.SplitterDistance = 316
        Me.SplitContainer1.TabIndex = 168
        '
        'grdWHTMOVE2
        '
        Appearance2.BackColor = System.Drawing.SystemColors.Window
        Appearance2.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdWHTMOVE2.DisplayLayout.Appearance = Appearance2
        UltraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn2.Header.VisiblePosition = 0
        UltraGridColumn2.Hidden = True
        UltraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn3.Format = "###0"
        UltraGridColumn3.Header.Caption = "Ln"
        UltraGridColumn3.Header.VisiblePosition = 1
        UltraGridColumn3.Width = 33
        UltraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn4.Header.Caption = "From"
        UltraGridColumn4.Header.VisiblePosition = 5
        UltraGridColumn9.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always
        Appearance3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        UltraGridColumn9.CellAppearance = Appearance3
        UltraGridColumn9.Header.Caption = "To"
        UltraGridColumn9.Header.VisiblePosition = 8
        UltraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.EditButton
        UltraGridColumn9.Width = 113
        UltraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn12.Header.VisiblePosition = 2
        UltraGridColumn12.Hidden = True
        Appearance4.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        UltraGridColumn13.CellAppearance = Appearance4
        UltraGridColumn13.Format = ""
        UltraGridColumn13.Header.Caption = "Qty"
        UltraGridColumn13.Header.VisiblePosition = 6
        UltraGridColumn13.Width = 77
        UltraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn32.Header.Caption = "Item"
        UltraGridColumn32.Header.VisiblePosition = 3
        UltraGridColumn32.Width = 101
        UltraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn14.Header.VisiblePosition = 7
        UltraGridColumn14.Hidden = True
        UltraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn15.Header.VisiblePosition = 9
        UltraGridColumn15.Hidden = True
        UltraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn16.Header.VisiblePosition = 10
        UltraGridColumn16.Hidden = True
        UltraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn17.Header.VisiblePosition = 11
        UltraGridColumn17.Hidden = True
        UltraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn18.Header.VisiblePosition = 12
        UltraGridColumn18.Hidden = True
        UltraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn20.Header.Caption = "Description"
        UltraGridColumn20.Header.VisiblePosition = 4
        UltraGridColumn20.Width = 169
        UltraGridColumn21.Header.VisiblePosition = 13
        UltraGridColumn21.Hidden = True
        UltraGridColumn1.Header.VisiblePosition = 14
        UltraGridColumn1.Hidden = True
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn2, UltraGridColumn3, UltraGridColumn4, UltraGridColumn9, UltraGridColumn12, UltraGridColumn13, UltraGridColumn32, UltraGridColumn14, UltraGridColumn15, UltraGridColumn16, UltraGridColumn17, UltraGridColumn18, UltraGridColumn20, UltraGridColumn21, UltraGridColumn1})
        Me.grdWHTMOVE2.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdWHTMOVE2.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance5.TextHAlignAsString = "Left"
        Me.grdWHTMOVE2.DisplayLayout.CaptionAppearance = Appearance5
        Appearance6.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance6.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance6.BorderColor = System.Drawing.SystemColors.Window
        Me.grdWHTMOVE2.DisplayLayout.GroupByBox.Appearance = Appearance6
        Appearance7.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdWHTMOVE2.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance7
        Me.grdWHTMOVE2.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdWHTMOVE2.DisplayLayout.GroupByBox.Hidden = True
        Appearance8.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance8.BackColor2 = System.Drawing.SystemColors.Control
        Appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance8.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdWHTMOVE2.DisplayLayout.GroupByBox.PromptAppearance = Appearance8
        Me.grdWHTMOVE2.DisplayLayout.MaxColScrollRegions = 1
        Me.grdWHTMOVE2.DisplayLayout.MaxRowScrollRegions = 1
        Appearance9.BackColor = System.Drawing.SystemColors.Window
        Appearance9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdWHTMOVE2.DisplayLayout.Override.ActiveCellAppearance = Appearance9
        Me.grdWHTMOVE2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdWHTMOVE2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[True]
        Me.grdWHTMOVE2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[True]
        Me.grdWHTMOVE2.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdWHTMOVE2.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance10.BackColor = System.Drawing.SystemColors.Window
        Me.grdWHTMOVE2.DisplayLayout.Override.CardAreaAppearance = Appearance10
        Appearance11.BorderColor = System.Drawing.Color.Silver
        Appearance11.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdWHTMOVE2.DisplayLayout.Override.CellAppearance = Appearance11
        Me.grdWHTMOVE2.DisplayLayout.Override.CellPadding = 0
        Appearance12.BackColor = System.Drawing.SystemColors.Control
        Appearance12.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance12.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance12.BorderColor = System.Drawing.SystemColors.Window
        Me.grdWHTMOVE2.DisplayLayout.Override.GroupByRowAppearance = Appearance12
        Appearance13.TextHAlignAsString = "Left"
        Me.grdWHTMOVE2.DisplayLayout.Override.HeaderAppearance = Appearance13
        Me.grdWHTMOVE2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdWHTMOVE2.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance14.BackColor = System.Drawing.SystemColors.Window
        Appearance14.BorderColor = System.Drawing.Color.Silver
        Me.grdWHTMOVE2.DisplayLayout.Override.RowAppearance = Appearance14
        Me.grdWHTMOVE2.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Me.grdWHTMOVE2.DisplayLayout.Override.SummaryDisplayArea = Infragistics.Win.UltraWinGrid.SummaryDisplayAreas.BottomFixed
        Appearance15.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdWHTMOVE2.DisplayLayout.Override.TemplateAddRowAppearance = Appearance15
        Me.grdWHTMOVE2.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdWHTMOVE2.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdWHTMOVE2.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdWHTMOVE2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdWHTMOVE2.Location = New System.Drawing.Point(0, 0)
        Me.grdWHTMOVE2.Name = "grdWHTMOVE2"
        Me.grdWHTMOVE2.Size = New System.Drawing.Size(735, 316)
        Me.grdWHTMOVE2.TabIndex = 167
        '
        'UltraGroupBox1
        '
        Me.UltraGroupBox1.Controls.Add(Me.btnMove)
        Me.UltraGroupBox1.Controls.Add(Me.btnCancel)
        Me.UltraGroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UltraGroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.UltraGroupBox1.Name = "UltraGroupBox1"
        Me.UltraGroupBox1.Size = New System.Drawing.Size(735, 38)
        Me.UltraGroupBox1.TabIndex = 0
        '
        'btnMove
        '
        Me.btnMove.Location = New System.Drawing.Point(554, 2)
        Me.btnMove.Name = "btnMove"
        Me.btnMove.Size = New System.Drawing.Size(81, 31)
        Me.btnMove.TabIndex = 117
        Me.btnMove.Text = "Move"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(641, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(81, 31)
        Me.btnCancel.TabIndex = 118
        Me.btnCancel.TabStop = False
        Me.btnCancel.Text = "Cancel"
        '
        'TAFLOCM1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(735, 431)
        Me.ControlBox = False
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "TAFLOCM1"
        Me.Text = "Move to Location"
        Me.ASFBASE2_Fill_Panel.ResumeLayout(False)
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splItems.Panel1.ResumeLayout(False)
        Me.splItems.Panel2.ResumeLayout(False)
        CType(Me.splItems, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splItems.ResumeLayout(False)
        CType(Me.grpHeader, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpHeader.ResumeLayout(False)
        CType(Me.UltraGroupBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox2.ResumeLayout(False)
        Me.UltraGroupBox2.PerformLayout()
        CType(Me.txtWHSE_TRAN_NO, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpTo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpTo.ResumeLayout(False)
        Me.grpTo.PerformLayout()
        CType(Me.UltraTextEditor3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.grdWHTMOVE2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents splItems As System.Windows.Forms.SplitContainer
    Friend WithEvents grdWHTMOVE2 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents grpHeader As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents btnCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents btnMove As Infragistics.Win.Misc.UltraButton
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents UltraGroupBox1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraGroupBox2 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents txtWHSE_TRAN_NO As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents grpTo As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraTextEditor3 As Infragistics.Win.UltraWinEditors.UltraTextEditor
End Class
