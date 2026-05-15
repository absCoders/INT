<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SOFORDRC
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
        Dim ValueListItem9 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem10 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem6 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem1 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("SOTORDR1", -1)
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_NO")
        Dim UltraGridColumn2 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_STORE_NO")
        Dim UltraGridColumn14 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_CUST_PO")
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_SHIP_TO")
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_DC_NO")
        Dim UltraGridColumn5 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_SHIP_DATE")
        Dim UltraGridColumn6 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_CANCEL_DATE")
        Dim UltraGridColumn7 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_ALLO_DATE")
        Dim UltraGridColumn8 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_ARRIVAL_DATE", -1, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, False)
        Dim UltraGridColumn9 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WHSE_CODE")
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
        Me.btnFixWhse = New Infragistics.Win.Misc.UltraButton()
        Me.whseFixed = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel()
        Me.optDateType = New Infragistics.Win.UltraWinEditors.UltraOptionSet()
        Me.cmdFixDates = New Infragistics.Win.Misc.UltraButton()
        Me.UltraLabel5 = New Infragistics.Win.Misc.UltraLabel()
        Me.dteFixed = New Infragistics.Win.UltraWinEditors.UltraDateTimeEditor()
        Me.UltraTextEditor1 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraTextEditor6 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel4 = New Infragistics.Win.Misc.UltraLabel()
        Me.grdSOTORDR1 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.cmdUpdate = New Infragistics.Win.Misc.UltraButton()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
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
        CType(Me.whseFixed, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optDateType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dteFixed, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdSOTORDR1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.SplitContainer1)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(1424, 889)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 889)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(1424, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 889)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(1424, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 889)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(1424, 0)
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
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.UltraGroupBox2)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.grdSOTORDR1)
        Me.SplitContainer2.Size = New System.Drawing.Size(1424, 762)
        Me.SplitContainer2.SplitterDistance = 54
        Me.SplitContainer2.TabIndex = 1
        '
        'UltraGroupBox2
        '
        Me.UltraGroupBox2.Controls.Add(Me.btnFixWhse)
        Me.UltraGroupBox2.Controls.Add(Me.whseFixed)
        Me.UltraGroupBox2.Controls.Add(Me.UltraLabel1)
        Me.UltraGroupBox2.Controls.Add(Me.optDateType)
        Me.UltraGroupBox2.Controls.Add(Me.cmdFixDates)
        Me.UltraGroupBox2.Controls.Add(Me.UltraLabel5)
        Me.UltraGroupBox2.Controls.Add(Me.dteFixed)
        Me.UltraGroupBox2.Controls.Add(Me.UltraTextEditor1)
        Me.UltraGroupBox2.Controls.Add(Me.UltraTextEditor6)
        Me.UltraGroupBox2.Controls.Add(Me.UltraLabel4)
        Me.UltraGroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UltraGroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.UltraGroupBox2.Name = "UltraGroupBox2"
        Me.UltraGroupBox2.Size = New System.Drawing.Size(1424, 54)
        Me.UltraGroupBox2.TabIndex = 0
        '
        'btnFixWhse
        '
        Me.btnFixWhse.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFixWhse.Location = New System.Drawing.Point(1247, 17)
        Me.btnFixWhse.Name = "btnFixWhse"
        Me.btnFixWhse.Size = New System.Drawing.Size(122, 30)
        Me.btnFixWhse.TabIndex = 200
        Me.btnFixWhse.Text = "Fix Warehouses"
        '
        'whseFixed
        '
        Me.Absx1.SetABSBindToTable(Me.whseFixed, False)
        Me.Absx1.SetABSColumnName(Me.whseFixed, "WHSE_CODE")
        Me.Absx1.SetABSHasButton(Me.whseFixed, True)
        Me.whseFixed.Location = New System.Drawing.Point(677, 24)
        Me.whseFixed.Name = "whseFixed"
        Me.whseFixed.ReadOnly = True
        Me.whseFixed.Size = New System.Drawing.Size(101, 25)
        Me.whseFixed.TabIndex = 199
        '
        'UltraLabel1
        '
        Me.UltraLabel1.AutoSize = True
        Me.UltraLabel1.Location = New System.Drawing.Point(677, 5)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(75, 18)
        Me.UltraLabel1.TabIndex = 198
        Me.UltraLabel1.Text = "New Whse"
        '
        'optDateType
        '
        Me.Absx1.SetABSBindToTable(Me.optDateType, False)
        Me.optDateType.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.optDateType.BorderStyle = Infragistics.Win.UIElementBorderStyle.None
        Me.optDateType.CheckedIndex = 2
        ValueListItem9.DataValue = "ORDR_SHIP_DATE"
        ValueListItem9.DisplayText = "Ship"
        ValueListItem10.DataValue = "ORDR_CANCEL_DATE"
        ValueListItem10.DisplayText = "Cancel"
        ValueListItem6.CheckState = System.Windows.Forms.CheckState.Checked
        ValueListItem6.DataValue = "ORDR_ALLO_DATE"
        ValueListItem6.DisplayText = "Allocation"
        ValueListItem1.DataValue = "ORDR_ARRIVAL_DATE"
        ValueListItem1.DisplayText = "Arrival"
        Me.optDateType.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem9, ValueListItem10, ValueListItem6, ValueListItem1})
        Me.optDateType.Location = New System.Drawing.Point(986, 16)
        Me.optDateType.Name = "optDateType"
        Me.optDateType.Size = New System.Drawing.Size(151, 36)
        Me.optDateType.TabIndex = 197
        Me.optDateType.Text = "Allocation"
        '
        'cmdFixDates
        '
        Me.cmdFixDates.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdFixDates.Location = New System.Drawing.Point(1156, 18)
        Me.cmdFixDates.Name = "cmdFixDates"
        Me.cmdFixDates.Size = New System.Drawing.Size(83, 30)
        Me.cmdFixDates.TabIndex = 170
        Me.cmdFixDates.Text = "Fix Dates"
        '
        'UltraLabel5
        '
        Me.UltraLabel5.AutoSize = True
        Me.UltraLabel5.Location = New System.Drawing.Point(488, 5)
        Me.UltraLabel5.Name = "UltraLabel5"
        Me.UltraLabel5.Size = New System.Drawing.Size(70, 18)
        Me.UltraLabel5.TabIndex = 169
        Me.UltraLabel5.Text = "New Date"
        '
        'dteFixed
        '
        Me.Absx1.SetABSBindToTable(Me.dteFixed, False)
        Me.dteFixed.DateTime = New Date(2007, 1, 27, 0, 0, 0, 0)
        Me.dteFixed.Location = New System.Drawing.Point(488, 23)
        Me.dteFixed.Name = "dteFixed"
        Me.dteFixed.Size = New System.Drawing.Size(114, 25)
        Me.dteFixed.TabIndex = 168
        Me.dteFixed.Value = New Date(2007, 1, 27, 0, 0, 0, 0)
        '
        'UltraTextEditor1
        '
        Me.Absx1.SetABSBindToTable(Me.UltraTextEditor1, False)
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor1, "CUST_NAME")
        Me.Absx1.SetABSParentColumnName(Me.UltraTextEditor1, "CUST_CODE")
        Me.UltraTextEditor1.Location = New System.Drawing.Point(113, 23)
        Me.UltraTextEditor1.Name = "UltraTextEditor1"
        Me.UltraTextEditor1.ReadOnly = True
        Me.UltraTextEditor1.Size = New System.Drawing.Size(302, 25)
        Me.UltraTextEditor1.TabIndex = 105
        '
        'UltraTextEditor6
        '
        Me.Absx1.SetABSBindToTable(Me.UltraTextEditor6, False)
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor6, "CUST_CODE")
        Me.UltraTextEditor6.Location = New System.Drawing.Point(6, 23)
        Me.UltraTextEditor6.Name = "UltraTextEditor6"
        Me.UltraTextEditor6.ReadOnly = True
        Me.UltraTextEditor6.Size = New System.Drawing.Size(101, 25)
        Me.UltraTextEditor6.TabIndex = 103
        '
        'UltraLabel4
        '
        Me.UltraLabel4.AutoSize = True
        Me.UltraLabel4.Location = New System.Drawing.Point(6, 5)
        Me.UltraLabel4.Name = "UltraLabel4"
        Me.UltraLabel4.Size = New System.Drawing.Size(70, 18)
        Me.UltraLabel4.TabIndex = 104
        Me.UltraLabel4.Text = "Customer"
        '
        'grdSOTORDR1
        '
        Appearance1.BackColor = System.Drawing.SystemColors.Window
        Appearance1.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdSOTORDR1.DisplayLayout.Appearance = Appearance1
        UltraGridColumn1.Header.Caption = "Order No"
        UltraGridColumn1.Header.VisiblePosition = 0
        UltraGridColumn1.Width = 111
        UltraGridColumn2.Header.Caption = "Store"
        UltraGridColumn2.Header.VisiblePosition = 1
        UltraGridColumn2.Width = 81
        UltraGridColumn14.Header.Caption = "Customer PO"
        UltraGridColumn14.Header.VisiblePosition = 2
        UltraGridColumn14.Width = 112
        UltraGridColumn3.Header.Caption = "ST"
        UltraGridColumn3.Header.VisiblePosition = 3
        UltraGridColumn3.Width = 47
        UltraGridColumn4.Header.Caption = "DC"
        UltraGridColumn4.Header.VisiblePosition = 4
        UltraGridColumn4.Width = 67
        UltraGridColumn5.Header.Caption = "Ship"
        UltraGridColumn5.Header.VisiblePosition = 5
        UltraGridColumn5.Width = 110
        UltraGridColumn6.Header.Caption = "Cancel"
        UltraGridColumn6.Header.VisiblePosition = 6
        UltraGridColumn6.Width = 110
        UltraGridColumn7.Header.Caption = "Allocation"
        UltraGridColumn7.Header.VisiblePosition = 7
        UltraGridColumn7.Width = 110
        UltraGridColumn8.Header.Caption = "Arrival"
        UltraGridColumn8.Header.VisiblePosition = 8
        UltraGridColumn8.Width = 110
        UltraGridColumn9.Header.Caption = "Whse"
        UltraGridColumn9.Header.VisiblePosition = 9
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn1, UltraGridColumn2, UltraGridColumn14, UltraGridColumn3, UltraGridColumn4, UltraGridColumn5, UltraGridColumn6, UltraGridColumn7, UltraGridColumn8, UltraGridColumn9})
        Me.grdSOTORDR1.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdSOTORDR1.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance2.TextHAlignAsString = "Left"
        Me.grdSOTORDR1.DisplayLayout.CaptionAppearance = Appearance2
        Me.grdSOTORDR1.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[True]
        Appearance3.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance3.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance3.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTORDR1.DisplayLayout.GroupByBox.Appearance = Appearance3
        Appearance4.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTORDR1.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance4
        Me.grdSOTORDR1.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdSOTORDR1.DisplayLayout.GroupByBox.Hidden = True
        Appearance5.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance5.BackColor2 = System.Drawing.SystemColors.Control
        Appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance5.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTORDR1.DisplayLayout.GroupByBox.PromptAppearance = Appearance5
        Me.grdSOTORDR1.DisplayLayout.MaxColScrollRegions = 1
        Me.grdSOTORDR1.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdSOTORDR1.DisplayLayout.NewBandLoadStyle = Infragistics.Win.UltraWinGrid.NewBandLoadStyle.Hide
        Me.grdSOTORDR1.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance6.BackColor = System.Drawing.SystemColors.Window
        Appearance6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdSOTORDR1.DisplayLayout.Override.ActiveCellAppearance = Appearance6
        Me.grdSOTORDR1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdSOTORDR1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTORDR1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTORDR1.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdSOTORDR1.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance7.BackColor = System.Drawing.Color.Transparent
        Me.grdSOTORDR1.DisplayLayout.Override.CardAreaAppearance = Appearance7
        Appearance8.BorderColor = System.Drawing.Color.Silver
        Appearance8.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdSOTORDR1.DisplayLayout.Override.CellAppearance = Appearance8
        Me.grdSOTORDR1.DisplayLayout.Override.CellPadding = 0
        Appearance9.BackColor = System.Drawing.SystemColors.Control
        Appearance9.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance9.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance9.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTORDR1.DisplayLayout.Override.GroupByRowAppearance = Appearance9
        Appearance10.TextHAlignAsString = "Left"
        Me.grdSOTORDR1.DisplayLayout.Override.HeaderAppearance = Appearance10
        Me.grdSOTORDR1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdSOTORDR1.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance11.BackColor = System.Drawing.SystemColors.Window
        Appearance11.BorderColor = System.Drawing.Color.Silver
        Me.grdSOTORDR1.DisplayLayout.Override.RowAppearance = Appearance11
        Me.grdSOTORDR1.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance12.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdSOTORDR1.DisplayLayout.Override.TemplateAddRowAppearance = Appearance12
        Me.grdSOTORDR1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdSOTORDR1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdSOTORDR1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand
        Me.grdSOTORDR1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdSOTORDR1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdSOTORDR1.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdSOTORDR1.Location = New System.Drawing.Point(0, 0)
        Me.grdSOTORDR1.Name = "grdSOTORDR1"
        Me.grdSOTORDR1.Size = New System.Drawing.Size(1424, 704)
        Me.grdSOTORDR1.TabIndex = 14
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUpdate.Location = New System.Drawing.Point(1247, 11)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.Size = New System.Drawing.Size(83, 30)
        Me.cmdUpdate.TabIndex = 15
        Me.cmdUpdate.Text = "Update"
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.SplitContainer2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.cmdCancel)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cmdUpdate)
        Me.SplitContainer1.Size = New System.Drawing.Size(1424, 889)
        Me.SplitContainer1.SplitterDistance = 762
        Me.SplitContainer1.TabIndex = 16
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.Location = New System.Drawing.Point(1331, 11)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(83, 30)
        Me.cmdCancel.TabIndex = 16
        Me.cmdCancel.Text = "Cancel"
        '
        'SOFORDRC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1424, 889)
        Me.ControlBox = False
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "SOFORDRC"
        Me.Text = "Multi-Order Edit"
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
        CType(Me.whseFixed, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optDateType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dteFixed, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdSOTORDR1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents UltraGroupBox2 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraTextEditor6 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel4 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents cmdUpdate As Infragistics.Win.Misc.UltraButton
    Friend WithEvents grdSOTORDR1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents cmdCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents UltraTextEditor1 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel5 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents dteFixed As Infragistics.Win.UltraWinEditors.UltraDateTimeEditor
    Friend WithEvents cmdFixDates As Infragistics.Win.Misc.UltraButton
    Friend WithEvents optDateType As Infragistics.Win.UltraWinEditors.UltraOptionSet
    Friend WithEvents btnFixWhse As Misc.UltraButton
    Friend WithEvents whseFixed As UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Misc.UltraLabel
End Class
