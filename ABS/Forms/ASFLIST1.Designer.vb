<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ASFLIST1
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
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("ASTLIST1", -1)
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LIST_CODE")
        Dim UltraGridColumn2 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LIST_DESC")
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("COLUMN_NAME")
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LIST_SHAREABLE")
        Dim UltraGridColumn5 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LIST_MODIFIABLE")
        Dim UltraGridColumn6 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_DATE")
        Dim UltraGridColumn7 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_OPER")
        Dim UltraGridColumn8 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_DATE")
        Dim UltraGridColumn22 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_OPER")
        Dim UltraGridColumn23 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CODE_VALUES", -1, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, False)
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
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand2 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("ASTLIST2", -1)
        Dim UltraGridColumn9 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LIST_CODE")
        Dim UltraGridColumn10 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CODE_VALUE")
        Dim UltraGridColumn11 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("DESC_VALUE")
        Dim Appearance15 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance19 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance21 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance22 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance23 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance24 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance25 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraToolTipInfo1 As Infragistics.Win.UltraWinToolTip.UltraToolTipInfo = New Infragistics.Win.UltraWinToolTip.UltraToolTipInfo("Retrieve a List of Code Values", Infragistics.Win.ToolTipImage.[Default], Nothing, Infragistics.Win.DefaultableBoolean.[Default])
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.grdASTLIST1 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.grdASTLIST2 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.lblLIST_DESC = New Infragistics.Win.Misc.UltraLabel()
        Me.cmdUpdate = New Infragistics.Win.Misc.UltraButton()
        Me.txtLIST_DESC = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.cmdCancel = New Infragistics.Win.Misc.UltraButton()
        Me.ASFBASE2_Fill_Panel.SuspendLayout()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.grdASTLIST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdASTLIST2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtLIST_DESC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.SplitContainer1)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(752, 428)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Margin = New System.Windows.Forms.Padding(4)
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 428)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(752, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Margin = New System.Windows.Forms.Padding(4)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 428)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Margin = New System.Windows.Forms.Padding(4)
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(752, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 428)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Margin = New System.Windows.Forms.Padding(4)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(752, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
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
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.lblLIST_DESC)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cmdUpdate)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtLIST_DESC)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cmdCancel)
        Me.SplitContainer1.Size = New System.Drawing.Size(752, 428)
        Me.SplitContainer1.SplitterDistance = 383
        Me.SplitContainer1.TabIndex = 0
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
        Me.SplitContainer2.Panel1.Controls.Add(Me.grdASTLIST1)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.grdASTLIST2)
        Me.SplitContainer2.Size = New System.Drawing.Size(752, 383)
        Me.SplitContainer2.SplitterDistance = 182
        Me.SplitContainer2.TabIndex = 2
        '
        'grdASTLIST1
        '
        Appearance1.BackColor = System.Drawing.SystemColors.Window
        Appearance1.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdASTLIST1.DisplayLayout.Appearance = Appearance1
        UltraGridColumn1.Header.Caption = "List"
        UltraGridColumn1.Header.VisiblePosition = 0
        UltraGridColumn1.Hidden = True
        UltraGridColumn1.Width = 37
        UltraGridColumn2.Header.Caption = "Description"
        UltraGridColumn2.Header.VisiblePosition = 2
        UltraGridColumn2.Width = 189
        UltraGridColumn3.Header.VisiblePosition = 1
        UltraGridColumn3.Hidden = True
        UltraGridColumn4.Header.Caption = "Share"
        UltraGridColumn4.Header.VisiblePosition = 3
        UltraGridColumn4.Hidden = True
        UltraGridColumn4.Width = 47
        UltraGridColumn5.Header.Caption = "Allow Chg"
        UltraGridColumn5.Header.VisiblePosition = 6
        UltraGridColumn5.Hidden = True
        UltraGridColumn5.Width = 55
        UltraGridColumn6.Format = "MM/dd/yy HH:mm"
        UltraGridColumn6.Header.Caption = "Created"
        UltraGridColumn6.Header.VisiblePosition = 5
        UltraGridColumn6.Width = 120
        UltraGridColumn7.Header.Caption = "By"
        UltraGridColumn7.Header.VisiblePosition = 7
        UltraGridColumn7.Width = 60
        UltraGridColumn8.Format = "MM/dd/yy HH:mm"
        UltraGridColumn8.Header.Caption = "Modified"
        UltraGridColumn8.Header.VisiblePosition = 8
        UltraGridColumn8.Width = 120
        UltraGridColumn22.Header.Caption = "By"
        UltraGridColumn22.Header.VisiblePosition = 9
        UltraGridColumn22.Width = 60
        UltraGridColumn23.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always
        UltraGridColumn23.Header.Caption = "Values"
        UltraGridColumn23.Header.VisiblePosition = 4
        UltraGridColumn23.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.EditButton
        UltraGridColumn23.Width = 157
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn1, UltraGridColumn2, UltraGridColumn3, UltraGridColumn4, UltraGridColumn5, UltraGridColumn6, UltraGridColumn7, UltraGridColumn8, UltraGridColumn22, UltraGridColumn23})
        Me.grdASTLIST1.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdASTLIST1.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance2.TextHAlignAsString = "Left"
        Me.grdASTLIST1.DisplayLayout.CaptionAppearance = Appearance2
        Me.grdASTLIST1.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance3.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance3.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance3.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASTLIST1.DisplayLayout.GroupByBox.Appearance = Appearance3
        Appearance4.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdASTLIST1.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance4
        Me.grdASTLIST1.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASTLIST1.DisplayLayout.GroupByBox.Hidden = True
        Appearance5.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance5.BackColor2 = System.Drawing.SystemColors.Control
        Appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance5.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdASTLIST1.DisplayLayout.GroupByBox.PromptAppearance = Appearance5
        Me.grdASTLIST1.DisplayLayout.MaxColScrollRegions = 1
        Me.grdASTLIST1.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdASTLIST1.DisplayLayout.NewBandLoadStyle = Infragistics.Win.UltraWinGrid.NewBandLoadStyle.Hide
        Me.grdASTLIST1.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance6.BackColor = System.Drawing.SystemColors.Window
        Appearance6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdASTLIST1.DisplayLayout.Override.ActiveCellAppearance = Appearance6
        Me.grdASTLIST1.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdASTLIST1.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance7.BackColor = System.Drawing.SystemColors.Window
        Me.grdASTLIST1.DisplayLayout.Override.CardAreaAppearance = Appearance7
        Appearance8.BorderColor = System.Drawing.Color.Silver
        Appearance8.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdASTLIST1.DisplayLayout.Override.CellAppearance = Appearance8
        Me.grdASTLIST1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdASTLIST1.DisplayLayout.Override.CellPadding = 0
        Appearance9.BackColor = System.Drawing.SystemColors.Control
        Appearance9.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance9.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance9.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASTLIST1.DisplayLayout.Override.GroupByRowAppearance = Appearance9
        Appearance10.TextHAlignAsString = "Left"
        Me.grdASTLIST1.DisplayLayout.Override.HeaderAppearance = Appearance10
        Me.grdASTLIST1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.[Select]
        Me.grdASTLIST1.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance11.BackColor = System.Drawing.SystemColors.Window
        Appearance11.BorderColor = System.Drawing.Color.Silver
        Me.grdASTLIST1.DisplayLayout.Override.RowAppearance = Appearance11
        Me.grdASTLIST1.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance12.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdASTLIST1.DisplayLayout.Override.TemplateAddRowAppearance = Appearance12
        Me.grdASTLIST1.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.None
        Me.grdASTLIST1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdASTLIST1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdASTLIST1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand
        Me.grdASTLIST1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdASTLIST1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdASTLIST1.Location = New System.Drawing.Point(0, 0)
        Me.grdASTLIST1.Name = "grdASTLIST1"
        Me.grdASTLIST1.Size = New System.Drawing.Size(752, 182)
        Me.grdASTLIST1.TabIndex = 0
        Me.grdASTLIST1.Text = "Filter"
        '
        'grdASTLIST2
        '
        Appearance14.BackColor = System.Drawing.SystemColors.Window
        Appearance14.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdASTLIST2.DisplayLayout.Appearance = Appearance14
        Me.grdASTLIST2.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns
        UltraGridColumn9.Header.VisiblePosition = 0
        UltraGridColumn9.Hidden = True
        UltraGridColumn9.Width = 88
        UltraGridColumn10.Header.Caption = "Code"
        UltraGridColumn10.Header.VisiblePosition = 1
        UltraGridColumn10.Width = 217
        UltraGridColumn11.Header.Caption = "Description"
        UltraGridColumn11.Header.VisiblePosition = 2
        UltraGridColumn11.Width = 514
        UltraGridBand2.Columns.AddRange(New Object() {UltraGridColumn9, UltraGridColumn10, UltraGridColumn11})
        Me.grdASTLIST2.DisplayLayout.BandsSerializer.Add(UltraGridBand2)
        Me.grdASTLIST2.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance15.TextHAlignAsString = "Left"
        Me.grdASTLIST2.DisplayLayout.CaptionAppearance = Appearance15
        Me.grdASTLIST2.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance16.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance16.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance16.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASTLIST2.DisplayLayout.GroupByBox.Appearance = Appearance16
        Appearance17.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdASTLIST2.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance17
        Me.grdASTLIST2.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASTLIST2.DisplayLayout.GroupByBox.Hidden = True
        Appearance18.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance18.BackColor2 = System.Drawing.SystemColors.Control
        Appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance18.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdASTLIST2.DisplayLayout.GroupByBox.PromptAppearance = Appearance18
        Me.grdASTLIST2.DisplayLayout.MaxColScrollRegions = 1
        Me.grdASTLIST2.DisplayLayout.MaxRowScrollRegions = 1
        Appearance19.BackColor = System.Drawing.SystemColors.Window
        Appearance19.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdASTLIST2.DisplayLayout.Override.ActiveCellAppearance = Appearance19
        Me.grdASTLIST2.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdASTLIST2.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance20.BackColor = System.Drawing.SystemColors.Window
        Me.grdASTLIST2.DisplayLayout.Override.CardAreaAppearance = Appearance20
        Appearance21.BorderColor = System.Drawing.Color.Silver
        Appearance21.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdASTLIST2.DisplayLayout.Override.CellAppearance = Appearance21
        Me.grdASTLIST2.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdASTLIST2.DisplayLayout.Override.CellPadding = 0
        Appearance22.BackColor = System.Drawing.SystemColors.Control
        Appearance22.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance22.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance22.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASTLIST2.DisplayLayout.Override.GroupByRowAppearance = Appearance22
        Appearance23.TextHAlignAsString = "Left"
        Me.grdASTLIST2.DisplayLayout.Override.HeaderAppearance = Appearance23
        Me.grdASTLIST2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.[Select]
        Me.grdASTLIST2.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance24.BackColor = System.Drawing.SystemColors.Window
        Appearance24.BorderColor = System.Drawing.Color.Silver
        Me.grdASTLIST2.DisplayLayout.Override.RowAppearance = Appearance24
        Me.grdASTLIST2.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance25.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdASTLIST2.DisplayLayout.Override.TemplateAddRowAppearance = Appearance25
        Me.grdASTLIST2.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.None
        Me.grdASTLIST2.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdASTLIST2.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdASTLIST2.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdASTLIST2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdASTLIST2.Location = New System.Drawing.Point(0, 0)
        Me.grdASTLIST2.Name = "grdASTLIST2"
        Me.grdASTLIST2.Size = New System.Drawing.Size(752, 197)
        Me.grdASTLIST2.TabIndex = 0
        Me.grdASTLIST2.Text = "Filter"
        '
        'lblLIST_DESC
        '
        Me.lblLIST_DESC.AutoSize = True
        Me.lblLIST_DESC.Location = New System.Drawing.Point(7, 10)
        Me.lblLIST_DESC.Name = "lblLIST_DESC"
        Me.lblLIST_DESC.Size = New System.Drawing.Size(80, 18)
        Me.lblLIST_DESC.TabIndex = 9
        Me.lblLIST_DESC.Text = "Description"
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUpdate.Location = New System.Drawing.Point(605, 3)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.Size = New System.Drawing.Size(69, 33)
        Me.cmdUpdate.TabIndex = 1
        Me.cmdUpdate.Text = "Update"
        '
        'txtLIST_DESC
        '
        Me.txtLIST_DESC.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLIST_DESC.Location = New System.Drawing.Point(93, 6)
        Me.txtLIST_DESC.MaxLength = 20
        Me.txtLIST_DESC.Name = "txtLIST_DESC"
        Me.txtLIST_DESC.Size = New System.Drawing.Size(416, 25)
        Me.txtLIST_DESC.TabIndex = 0
        UltraToolTipInfo1.ToolTipText = "Retrieve a List of Code Values"
        Me.tip.SetUltraToolTip(Me.txtLIST_DESC, UltraToolTipInfo1)
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.Location = New System.Drawing.Point(680, 3)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(69, 33)
        Me.cmdCancel.TabIndex = 2
        Me.cmdCancel.Text = "Cancel"
        '
        'ASFLIST1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(752, 428)
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "ASFLIST1"
        Me.Text = "Code Lists"
        Me.ASFBASE2_Fill_Panel.ResumeLayout(False)
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.grdASTLIST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdASTLIST2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtLIST_DESC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents grdASTLIST1 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents cmdCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents txtLIST_DESC As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents grdASTLIST2 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents cmdUpdate As Infragistics.Win.Misc.UltraButton
    Friend WithEvents lblLIST_DESC As Infragistics.Win.Misc.UltraLabel
End Class
