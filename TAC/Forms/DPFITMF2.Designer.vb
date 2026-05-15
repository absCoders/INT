<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class DPFITMF2
    Inherits ABSolution.ASFBASE2

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
        Dim ValueListItem30 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem31 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("DPTITMF2", -1)
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("OPS_YYYYPP")
        Dim UltraGridColumn2 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ITEM_CODE")
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("MARKET_CODE")
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("OPS_YYYYPP_FC")
        Dim UltraGridColumn5 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("FORECAST")
        Dim UltraGridColumn6 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_OPER")
        Dim UltraGridColumn7 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_DATE")
        Dim UltraGridColumn8 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("FORECAST_NOTE")
        Dim UltraGridColumn9 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("STATUS", 0, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, False)
        Dim UltraGridColumn10 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_OPER", 1)
        Dim UltraGridColumn11 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_DATE", 2)
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance10 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance11 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance12 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.UltraGroupBox2 = New Infragistics.Win.Misc.UltraGroupBox()
        Me.optCH = New Infragistics.Win.UltraWinEditors.UltraOptionSet()
        Me.UltraTextEditor3 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraTextEditor2 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraTextEditor1 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraTextEditor6 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel4 = New Infragistics.Win.Misc.UltraLabel()
        Me.grdDPTITMF2 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.cmdUpdate = New Infragistics.Win.Misc.UltraButton()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.cmdADD = New Infragistics.Win.Misc.UltraButton()
        Me.lblFORECAST_NOTE = New Infragistics.Win.Misc.UltraLabel()
        Me.txtFORECAST_NOTE = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.cmdCancel = New Infragistics.Win.Misc.UltraButton()
        Me.ASFBASE2_Fill_Panel.SuspendLayout()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTLOGX1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.UltraGroupBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox2.SuspendLayout()
        CType(Me.optCH, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdDPTITMF2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.txtFORECAST_NOTE, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.SplitContainer1)
        Me.ASFBASE2_Fill_Panel.Margin = New System.Windows.Forms.Padding(5, 3, 5, 3)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(1289, 450)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Margin = New System.Windows.Forms.Padding(5, 3, 5, 3)
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 450)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(1289, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Margin = New System.Windows.Forms.Padding(5, 3, 5, 3)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 450)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Margin = New System.Windows.Forms.Padding(5, 3, 5, 3)
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(1289, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 450)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Margin = New System.Windows.Forms.Padding(5, 3, 5, 3)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(1289, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.UltraGroupBox2)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.grdDPTITMF2)
        Me.SplitContainer2.Size = New System.Drawing.Size(1289, 385)
        Me.SplitContainer2.SplitterDistance = 54
        Me.SplitContainer2.TabIndex = 1
        '
        'UltraGroupBox2
        '
        Me.UltraGroupBox2.Controls.Add(Me.optCH)
        Me.UltraGroupBox2.Controls.Add(Me.UltraTextEditor3)
        Me.UltraGroupBox2.Controls.Add(Me.UltraLabel2)
        Me.UltraGroupBox2.Controls.Add(Me.UltraTextEditor2)
        Me.UltraGroupBox2.Controls.Add(Me.UltraLabel1)
        Me.UltraGroupBox2.Controls.Add(Me.UltraTextEditor1)
        Me.UltraGroupBox2.Controls.Add(Me.UltraTextEditor6)
        Me.UltraGroupBox2.Controls.Add(Me.UltraLabel4)
        Me.UltraGroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UltraGroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.UltraGroupBox2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraGroupBox2.Name = "UltraGroupBox2"
        Me.UltraGroupBox2.Size = New System.Drawing.Size(1289, 54)
        Me.UltraGroupBox2.TabIndex = 0
        '
        'optCH
        '
        Me.optCH.BackColor = System.Drawing.Color.Transparent
        Me.optCH.BackColorInternal = System.Drawing.Color.Transparent
        Me.optCH.BorderStyle = Infragistics.Win.UIElementBorderStyle.None
        Me.optCH.CheckedIndex = 0
        ValueListItem30.CheckState = System.Windows.Forms.CheckState.Checked
        ValueListItem30.DataValue = "C"
        ValueListItem30.DisplayText = "Current"
        ValueListItem31.DataValue = "H"
        ValueListItem31.DisplayText = "History"
        Me.optCH.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem30, ValueListItem31})
        Me.optCH.Location = New System.Drawing.Point(982, 26)
        Me.optCH.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.optCH.Name = "optCH"
        Me.optCH.Size = New System.Drawing.Size(202, 31)
        Me.optCH.TabIndex = 5
        Me.optCH.Text = "Current"
        '
        'UltraTextEditor3
        '
        Me.Absx1.SetABSBindToTable(Me.UltraTextEditor3, False)
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor3, "OPS_YYYYPP_FC")
        Me.UltraTextEditor3.Location = New System.Drawing.Point(790, 26)
        Me.UltraTextEditor3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraTextEditor3.Name = "UltraTextEditor3"
        Me.UltraTextEditor3.ReadOnly = True
        Me.UltraTextEditor3.Size = New System.Drawing.Size(174, 29)
        Me.UltraTextEditor3.TabIndex = 142
        '
        'UltraLabel2
        '
        Me.UltraLabel2.AutoSize = True
        Me.UltraLabel2.Location = New System.Drawing.Point(790, 6)
        Me.UltraLabel2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(103, 22)
        Me.UltraLabel2.TabIndex = 143
        Me.UltraLabel2.Text = "Forecast YP"
        '
        'UltraTextEditor2
        '
        Me.Absx1.SetABSBindToTable(Me.UltraTextEditor2, False)
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor2, "MARKET_CODE")
        Me.UltraTextEditor2.Location = New System.Drawing.Point(598, 26)
        Me.UltraTextEditor2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraTextEditor2.Name = "UltraTextEditor2"
        Me.UltraTextEditor2.ReadOnly = True
        Me.UltraTextEditor2.Size = New System.Drawing.Size(174, 29)
        Me.UltraTextEditor2.TabIndex = 140
        '
        'UltraLabel1
        '
        Me.UltraLabel1.AutoSize = True
        Me.UltraLabel1.Location = New System.Drawing.Point(598, 6)
        Me.UltraLabel1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(64, 22)
        Me.UltraLabel1.TabIndex = 141
        Me.UltraLabel1.Text = "Market"
        '
        'UltraTextEditor1
        '
        Me.Absx1.SetABSBindToTable(Me.UltraTextEditor1, False)
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor1, "ITEM_DESC")
        Me.Absx1.SetABSParentColumnName(Me.UltraTextEditor1, "ITEM_CODE")
        Me.UltraTextEditor1.Location = New System.Drawing.Point(189, 26)
        Me.UltraTextEditor1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraTextEditor1.Name = "UltraTextEditor1"
        Me.UltraTextEditor1.ReadOnly = True
        Me.UltraTextEditor1.Size = New System.Drawing.Size(382, 29)
        Me.UltraTextEditor1.TabIndex = 139
        Me.UltraTextEditor1.TabStop = False
        '
        'UltraTextEditor6
        '
        Me.Absx1.SetABSBindToTable(Me.UltraTextEditor6, False)
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor6, "ITEM_CODE")
        Me.UltraTextEditor6.Location = New System.Drawing.Point(8, 26)
        Me.UltraTextEditor6.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraTextEditor6.Name = "UltraTextEditor6"
        Me.UltraTextEditor6.ReadOnly = True
        Me.UltraTextEditor6.Size = New System.Drawing.Size(174, 29)
        Me.UltraTextEditor6.TabIndex = 103
        '
        'UltraLabel4
        '
        Me.UltraLabel4.AutoSize = True
        Me.UltraLabel4.Location = New System.Drawing.Point(8, 6)
        Me.UltraLabel4.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.UltraLabel4.Name = "UltraLabel4"
        Me.UltraLabel4.Size = New System.Drawing.Size(46, 22)
        Me.UltraLabel4.TabIndex = 104
        Me.UltraLabel4.Text = "Item"
        '
        'grdDPTITMF2
        '
        Appearance1.BackColor = System.Drawing.SystemColors.Window
        Appearance1.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdDPTITMF2.DisplayLayout.Appearance = Appearance1
        UltraGridColumn1.Header.Caption = "YP@Note"
        UltraGridColumn1.Header.VisiblePosition = 0
        UltraGridColumn1.Width = 81
        UltraGridColumn2.Header.VisiblePosition = 1
        UltraGridColumn2.Hidden = True
        UltraGridColumn3.Header.VisiblePosition = 2
        UltraGridColumn3.Hidden = True
        UltraGridColumn4.Header.VisiblePosition = 3
        UltraGridColumn4.Hidden = True
        UltraGridColumn5.Header.Caption = "FC"
        UltraGridColumn5.Header.VisiblePosition = 4
        UltraGridColumn5.Width = 94
        UltraGridColumn6.Header.Caption = "By"
        UltraGridColumn6.Header.VisiblePosition = 6
        UltraGridColumn6.Width = 100
        UltraGridColumn7.Format = "MM/dd/yy HH:mm"
        UltraGridColumn7.Header.Caption = "Created"
        UltraGridColumn7.Header.VisiblePosition = 5
        UltraGridColumn7.Width = 127
        UltraGridColumn8.Header.Caption = "Forecast Note"
        UltraGridColumn8.Header.VisiblePosition = 7
        UltraGridColumn8.Width = 352
        UltraGridColumn9.Header.Caption = "Status"
        UltraGridColumn9.Header.VisiblePosition = 8
        UltraGridColumn9.Width = 45
        UltraGridColumn10.Header.Caption = "Last User"
        UltraGridColumn10.Header.VisiblePosition = 9
        UltraGridColumn10.Width = 99
        UltraGridColumn11.Header.Caption = "Last Date"
        UltraGridColumn11.Header.VisiblePosition = 10
        UltraGridColumn11.Width = 97
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn1, UltraGridColumn2, UltraGridColumn3, UltraGridColumn4, UltraGridColumn5, UltraGridColumn6, UltraGridColumn7, UltraGridColumn8, UltraGridColumn9, UltraGridColumn10, UltraGridColumn11})
        Me.grdDPTITMF2.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdDPTITMF2.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance2.TextHAlignAsString = "Left"
        Me.grdDPTITMF2.DisplayLayout.CaptionAppearance = Appearance2
        Me.grdDPTITMF2.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[True]
        Appearance3.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance3.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance3.BorderColor = System.Drawing.SystemColors.Window
        Me.grdDPTITMF2.DisplayLayout.GroupByBox.Appearance = Appearance3
        Appearance4.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdDPTITMF2.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance4
        Me.grdDPTITMF2.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdDPTITMF2.DisplayLayout.GroupByBox.Hidden = True
        Appearance5.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance5.BackColor2 = System.Drawing.SystemColors.Control
        Appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance5.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdDPTITMF2.DisplayLayout.GroupByBox.PromptAppearance = Appearance5
        Me.grdDPTITMF2.DisplayLayout.MaxColScrollRegions = 1
        Me.grdDPTITMF2.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdDPTITMF2.DisplayLayout.NewBandLoadStyle = Infragistics.Win.UltraWinGrid.NewBandLoadStyle.Hide
        Me.grdDPTITMF2.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance6.BackColor = System.Drawing.SystemColors.Window
        Appearance6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdDPTITMF2.DisplayLayout.Override.ActiveCellAppearance = Appearance6
        Me.grdDPTITMF2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdDPTITMF2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdDPTITMF2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdDPTITMF2.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdDPTITMF2.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance7.BackColor = System.Drawing.Color.Transparent
        Me.grdDPTITMF2.DisplayLayout.Override.CardAreaAppearance = Appearance7
        Appearance8.BorderColor = System.Drawing.Color.Silver
        Appearance8.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdDPTITMF2.DisplayLayout.Override.CellAppearance = Appearance8
        Me.grdDPTITMF2.DisplayLayout.Override.CellPadding = 0
        Appearance9.BackColor = System.Drawing.SystemColors.Control
        Appearance9.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance9.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance9.BorderColor = System.Drawing.SystemColors.Window
        Me.grdDPTITMF2.DisplayLayout.Override.GroupByRowAppearance = Appearance9
        Appearance10.TextHAlignAsString = "Left"
        Me.grdDPTITMF2.DisplayLayout.Override.HeaderAppearance = Appearance10
        Me.grdDPTITMF2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdDPTITMF2.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance11.BackColor = System.Drawing.SystemColors.Window
        Appearance11.BorderColor = System.Drawing.Color.Silver
        Me.grdDPTITMF2.DisplayLayout.Override.RowAppearance = Appearance11
        Me.grdDPTITMF2.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance12.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdDPTITMF2.DisplayLayout.Override.TemplateAddRowAppearance = Appearance12
        Me.grdDPTITMF2.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdDPTITMF2.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdDPTITMF2.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand
        Me.grdDPTITMF2.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdDPTITMF2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdDPTITMF2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdDPTITMF2.Location = New System.Drawing.Point(0, 0)
        Me.grdDPTITMF2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.grdDPTITMF2.Name = "grdDPTITMF2"
        Me.grdDPTITMF2.Size = New System.Drawing.Size(1289, 327)
        Me.grdDPTITMF2.TabIndex = 14
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUpdate.Location = New System.Drawing.Point(1068, 12)
        Me.cmdUpdate.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.Size = New System.Drawing.Size(104, 34)
        Me.cmdUpdate.TabIndex = 15
        Me.cmdUpdate.Text = "Update"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.SplitContainer2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.cmdADD)
        Me.SplitContainer1.Panel2.Controls.Add(Me.lblFORECAST_NOTE)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtFORECAST_NOTE)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cmdCancel)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cmdUpdate)
        Me.SplitContainer1.Size = New System.Drawing.Size(1289, 450)
        Me.SplitContainer1.SplitterDistance = 385
        Me.SplitContainer1.TabIndex = 16
        '
        'cmdADD
        '
        Me.cmdADD.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdADD.Location = New System.Drawing.Point(850, 15)
        Me.cmdADD.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cmdADD.Name = "cmdADD"
        Me.cmdADD.Size = New System.Drawing.Size(114, 34)
        Me.cmdADD.TabIndex = 142
        Me.cmdADD.Text = "Add Note"
        '
        'lblFORECAST_NOTE
        '
        Me.lblFORECAST_NOTE.AutoSize = True
        Me.lblFORECAST_NOTE.Location = New System.Drawing.Point(8, 22)
        Me.lblFORECAST_NOTE.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.lblFORECAST_NOTE.Name = "lblFORECAST_NOTE"
        Me.lblFORECAST_NOTE.Size = New System.Drawing.Size(138, 22)
        Me.lblFORECAST_NOTE.TabIndex = 141
        Me.lblFORECAST_NOTE.Text = "Enter New Note"
        '
        'txtFORECAST_NOTE
        '
        Me.Absx1.SetABSBindToTable(Me.txtFORECAST_NOTE, False)
        Me.Absx1.SetABSColumnName(Me.txtFORECAST_NOTE, "FORECAST_NOTE")
        Me.txtFORECAST_NOTE.Location = New System.Drawing.Point(146, 17)
        Me.txtFORECAST_NOTE.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtFORECAST_NOTE.MaxLength = 200
        Me.txtFORECAST_NOTE.Name = "txtFORECAST_NOTE"
        Me.txtFORECAST_NOTE.Size = New System.Drawing.Size(696, 29)
        Me.txtFORECAST_NOTE.TabIndex = 140
        Me.txtFORECAST_NOTE.TabStop = False
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.Location = New System.Drawing.Point(1173, 12)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(104, 34)
        Me.cmdCancel.TabIndex = 16
        Me.cmdCancel.Text = "Cancel"
        '
        'DPFITMF2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(10.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1289, 450)
        Me.ControlBox = False
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.Name = "DPFITMF2"
        Me.Text = "Forecast Notes"
        Me.ASFBASE2_Fill_Panel.ResumeLayout(False)
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTLOGX1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.UltraGroupBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox2.ResumeLayout(False)
        Me.UltraGroupBox2.PerformLayout()
        CType(Me.optCH, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdDPTITMF2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.txtFORECAST_NOTE, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents UltraGroupBox2 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraTextEditor6 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel4 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents cmdUpdate As Infragistics.Win.Misc.UltraButton
    Friend WithEvents grdDPTITMF2 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents cmdCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraTextEditor1 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraTextEditor3 As UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Misc.UltraLabel
    Friend WithEvents UltraTextEditor2 As UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Misc.UltraLabel
    Friend WithEvents lblFORECAST_NOTE As Misc.UltraLabel
    Friend WithEvents txtFORECAST_NOTE As UltraWinEditors.UltraTextEditor
    Friend WithEvents optCH As UltraWinEditors.UltraOptionSet
    Friend WithEvents cmdADD As Misc.UltraButton
End Class
