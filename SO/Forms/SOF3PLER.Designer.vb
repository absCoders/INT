<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SOF3PLER
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
        Dim UltraExplorerBarGroup1 As Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarGroup = New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarGroup()
        Dim UltraExplorerBarItem1 As Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem = New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem()
        Dim UltraExplorerBarItem3 As Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem = New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem()
        Dim Appearance13 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance14 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance15 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance19 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance21 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("SOTSHIPX", -1)
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PROCESS_IND")
        Dim UltraGridColumn2 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CODE")
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_CUST_PO")
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_GROUP_NO")
        Dim UltraGridColumn5 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_DATE")
        Dim UltraGridColumn6 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_SHIP_DATE")
        Dim UltraGridColumn7 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_CANCEL_DATE")
        Dim UltraGridColumn8 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SHIP_BOL_NO")
        Dim UltraGridColumn9 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SOTSHIPX_SOTPICKX")
        Dim UltraGridBand2 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("SOTSHIPX_SOTPICKX", 0)
        Dim UltraGridColumn10 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SHIP_BOL_NO")
        Dim UltraGridColumn11 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PICK_NO")
        Dim UltraGridColumn12 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_NO")
        Dim UltraGridColumn13 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PICK_RELEASED")
        Dim UltraGridColumn14 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("OHORDN")
        Dim UltraGridColumn15 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("OHINVN")
        Dim UltraGridColumn21 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TOT_DTL_UNITS")
        Dim UltraGridColumn22 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TOT_CTN_UNITS")
        Dim UltraGridColumn16 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("SOTPICKX_SHIPDTLX")
        Dim UltraGridBand3 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("SOTPICKX_SHIPDTLX", 1)
        Dim UltraGridColumn17 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_NO")
        Dim UltraGridColumn18 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ITEM_CODE")
        Dim UltraGridColumn19 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("DTL_UNITS")
        Dim UltraGridColumn20 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CTN_UNITS")
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
        Dim Appearance22 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance23 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance24 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance25 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance26 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance27 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance28 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance29 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance30 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance31 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance32 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance33 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.grdSOTSHIP1 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.UltraTabPageControl5 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.UltraGroupBox1 = New Infragistics.Win.Misc.UltraGroupBox()
        Me.UltraCombo1 = New Infragistics.Win.UltraWinGrid.UltraCombo()
        Me.spl = New System.Windows.Forms.SplitContainer()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ASFBASE1_Fill_Panel.SuspendLayout()
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdSOTSHIP1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraCombo1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spl, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.spl.Panel1.SuspendLayout()
        Me.spl.Panel2.SuspendLayout()
        Me.spl.SuspendLayout()
        Me.SuspendLayout()
        '
        'UltraExplorerBar1
        '
        UltraExplorerBarItem1.Key = "Load"
        UltraExplorerBarItem1.Text = "Load"
        UltraExplorerBarItem3.Key = "Done"
        UltraExplorerBarItem3.Text = "Done"
        UltraExplorerBarGroup1.Items.AddRange(New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem() {UltraExplorerBarItem1, UltraExplorerBarItem3})
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
        Me.UltraExplorerBar1.Tag = "CLICK"
        '
        'ASFBASE1_Fill_Panel
        '
        Me.ASFBASE1_Fill_Panel.Controls.Add(Me.spl)
        Me.ASFBASE1_Fill_Panel.Controls.SetChildIndex(Me.grdASFBASEX, 0)
        Me.ASFBASE1_Fill_Panel.Controls.SetChildIndex(Me.spl, 0)
        '
        'grdASFBASEX
        '
        Appearance13.BackColor = System.Drawing.SystemColors.Window
        Appearance13.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdASFBASEX.DisplayLayout.Appearance = Appearance13
        Me.grdASFBASEX.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdASFBASEX.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.MaxColScrollRegions = 1
        Me.grdASFBASEX.DisplayLayout.MaxRowScrollRegions = 1
        Appearance14.BackColor = System.Drawing.SystemColors.Window
        Appearance14.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveCellAppearance = Appearance14
        Appearance15.BackColor = System.Drawing.SystemColors.Highlight
        Appearance15.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveRowAppearance = Appearance15
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance16.BackColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.CardAreaAppearance = Appearance16
        Appearance17.BorderColor = System.Drawing.Color.Silver
        Appearance17.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdASFBASEX.DisplayLayout.Override.CellAppearance = Appearance17
        Me.grdASFBASEX.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdASFBASEX.DisplayLayout.Override.CellPadding = 0
        Appearance18.BackColor = System.Drawing.SystemColors.Control
        Appearance18.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance18.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance18.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.GroupByRowAppearance = Appearance18
        Appearance19.TextHAlignAsString = "Left"
        Me.grdASFBASEX.DisplayLayout.Override.HeaderAppearance = Appearance19
        Me.grdASFBASEX.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdASFBASEX.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance20.BackColor = System.Drawing.SystemColors.Window
        Appearance20.BorderColor = System.Drawing.Color.Silver
        Me.grdASFBASEX.DisplayLayout.Override.RowAppearance = Appearance20
        Me.grdASFBASEX.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance21.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdASFBASEX.DisplayLayout.Override.TemplateAddRowAppearance = Appearance21
        Me.grdASFBASEX.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdASFBASEX.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdASFBASEX.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'grdSOTSHIP1
        '
        Appearance1.BackColor = System.Drawing.SystemColors.Window
        Appearance1.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdSOTSHIP1.DisplayLayout.Appearance = Appearance1
        UltraGridColumn1.Header.Caption = "Error"
        UltraGridColumn1.Header.VisiblePosition = 0
        UltraGridColumn2.Header.Caption = "Customer"
        UltraGridColumn2.Header.VisiblePosition = 1
        UltraGridColumn3.Header.Caption = "Order PO"
        UltraGridColumn3.Header.VisiblePosition = 2
        UltraGridColumn4.Header.Caption = "Group No"
        UltraGridColumn4.Header.VisiblePosition = 3
        UltraGridColumn5.Header.Caption = "Order Date"
        UltraGridColumn5.Header.VisiblePosition = 4
        UltraGridColumn6.Header.Caption = "Ship By Date"
        UltraGridColumn6.Header.VisiblePosition = 5
        UltraGridColumn7.Header.Caption = "Cancel Date"
        UltraGridColumn7.Header.VisiblePosition = 6
        UltraGridColumn8.Header.Caption = "ABS Ship Bol No"
        UltraGridColumn8.Header.VisiblePosition = 7
        UltraGridColumn9.Header.VisiblePosition = 8
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn1, UltraGridColumn2, UltraGridColumn3, UltraGridColumn4, UltraGridColumn5, UltraGridColumn6, UltraGridColumn7, UltraGridColumn8, UltraGridColumn9})
        UltraGridColumn10.Header.VisiblePosition = 0
        UltraGridColumn10.Hidden = True
        UltraGridColumn11.Header.Caption = "Pick No"
        UltraGridColumn11.Header.VisiblePosition = 1
        UltraGridColumn12.Header.Caption = "Order No"
        UltraGridColumn12.Header.VisiblePosition = 2
        UltraGridColumn13.Header.Caption = "Released"
        UltraGridColumn13.Header.VisiblePosition = 3
        UltraGridColumn14.Header.Caption = "Clarins Order No"
        UltraGridColumn14.Header.VisiblePosition = 4
        UltraGridColumn15.Header.Caption = "Clarins Inv No"
        UltraGridColumn15.Header.VisiblePosition = 5
        UltraGridColumn21.Header.Caption = "Total Dtl Units"
        UltraGridColumn21.Header.VisiblePosition = 6
        UltraGridColumn22.Header.Caption = "Total Ctn Units"
        UltraGridColumn22.Header.VisiblePosition = 7
        UltraGridColumn16.Header.VisiblePosition = 8
        UltraGridBand2.Columns.AddRange(New Object() {UltraGridColumn10, UltraGridColumn11, UltraGridColumn12, UltraGridColumn13, UltraGridColumn14, UltraGridColumn15, UltraGridColumn21, UltraGridColumn22, UltraGridColumn16})
        UltraGridColumn17.Header.VisiblePosition = 0
        UltraGridColumn17.Hidden = True
        UltraGridColumn18.Header.Caption = "Item Code"
        UltraGridColumn18.Header.VisiblePosition = 1
        UltraGridColumn19.Format = "#,##0"
        UltraGridColumn19.Header.Caption = "Detail Units"
        UltraGridColumn19.Header.VisiblePosition = 2
        UltraGridColumn20.Format = "#,##0"
        UltraGridColumn20.Header.Caption = "Carton Units"
        UltraGridColumn20.Header.VisiblePosition = 3
        UltraGridBand3.Columns.AddRange(New Object() {UltraGridColumn17, UltraGridColumn18, UltraGridColumn19, UltraGridColumn20})
        Me.grdSOTSHIP1.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdSOTSHIP1.DisplayLayout.BandsSerializer.Add(UltraGridBand2)
        Me.grdSOTSHIP1.DisplayLayout.BandsSerializer.Add(UltraGridBand3)
        Me.grdSOTSHIP1.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance2.TextHAlignAsString = "Left"
        Me.grdSOTSHIP1.DisplayLayout.CaptionAppearance = Appearance2
        Appearance3.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance3.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance3.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTSHIP1.DisplayLayout.GroupByBox.Appearance = Appearance3
        Appearance4.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTSHIP1.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance4
        Me.grdSOTSHIP1.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdSOTSHIP1.DisplayLayout.GroupByBox.Hidden = True
        Appearance5.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance5.BackColor2 = System.Drawing.SystemColors.Control
        Appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance5.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdSOTSHIP1.DisplayLayout.GroupByBox.PromptAppearance = Appearance5
        Me.grdSOTSHIP1.DisplayLayout.MaxColScrollRegions = 1
        Me.grdSOTSHIP1.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdSOTSHIP1.DisplayLayout.NewBandLoadStyle = Infragistics.Win.UltraWinGrid.NewBandLoadStyle.Hide
        Me.grdSOTSHIP1.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance6.BackColor = System.Drawing.SystemColors.Window
        Appearance6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdSOTSHIP1.DisplayLayout.Override.ActiveCellAppearance = Appearance6
        Me.grdSOTSHIP1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdSOTSHIP1.DisplayLayout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
        Me.grdSOTSHIP1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTSHIP1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdSOTSHIP1.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdSOTSHIP1.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance7.BackColor = System.Drawing.SystemColors.Window
        Me.grdSOTSHIP1.DisplayLayout.Override.CardAreaAppearance = Appearance7
        Appearance8.BorderColor = System.Drawing.Color.Silver
        Appearance8.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdSOTSHIP1.DisplayLayout.Override.CellAppearance = Appearance8
        Me.grdSOTSHIP1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        Me.grdSOTSHIP1.DisplayLayout.Override.CellPadding = 0
        Appearance9.BackColor = System.Drawing.SystemColors.Control
        Appearance9.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance9.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance9.BorderColor = System.Drawing.SystemColors.Window
        Me.grdSOTSHIP1.DisplayLayout.Override.GroupByRowAppearance = Appearance9
        Appearance10.TextHAlignAsString = "Left"
        Me.grdSOTSHIP1.DisplayLayout.Override.HeaderAppearance = Appearance10
        Me.grdSOTSHIP1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdSOTSHIP1.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance11.BackColor = System.Drawing.SystemColors.Window
        Appearance11.BorderColor = System.Drawing.Color.Silver
        Me.grdSOTSHIP1.DisplayLayout.Override.RowAppearance = Appearance11
        Me.grdSOTSHIP1.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance12.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdSOTSHIP1.DisplayLayout.Override.TemplateAddRowAppearance = Appearance12
        Me.grdSOTSHIP1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdSOTSHIP1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdSOTSHIP1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdSOTSHIP1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdSOTSHIP1.Location = New System.Drawing.Point(0, 0)
        Me.grdSOTSHIP1.Name = "grdSOTSHIP1"
        Me.grdSOTSHIP1.Size = New System.Drawing.Size(779, 587)
        Me.grdSOTSHIP1.TabIndex = 165
        '
        'UltraTabPageControl5
        '
        Me.UltraTabPageControl5.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabPageControl5.Name = "UltraTabPageControl5"
        Me.UltraTabPageControl5.Size = New System.Drawing.Size(765, 281)
        '
        'UltraGroupBox1
        '
        Me.UltraGroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UltraGroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.UltraGroupBox1.Name = "UltraGroupBox1"
        Me.UltraGroupBox1.Size = New System.Drawing.Size(779, 25)
        Me.UltraGroupBox1.TabIndex = 2
        '
        'UltraCombo1
        '
        Appearance22.BackColor = System.Drawing.SystemColors.Window
        Appearance22.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.UltraCombo1.DisplayLayout.Appearance = Appearance22
        Me.UltraCombo1.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.UltraCombo1.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance23.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance23.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance23.BorderColor = System.Drawing.SystemColors.Window
        Me.UltraCombo1.DisplayLayout.GroupByBox.Appearance = Appearance23
        Appearance24.ForeColor = System.Drawing.SystemColors.GrayText
        Me.UltraCombo1.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance24
        Me.UltraCombo1.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance25.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance25.BackColor2 = System.Drawing.SystemColors.Control
        Appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance25.ForeColor = System.Drawing.SystemColors.GrayText
        Me.UltraCombo1.DisplayLayout.GroupByBox.PromptAppearance = Appearance25
        Me.UltraCombo1.DisplayLayout.MaxColScrollRegions = 1
        Me.UltraCombo1.DisplayLayout.MaxRowScrollRegions = 1
        Appearance26.BackColor = System.Drawing.SystemColors.Window
        Appearance26.ForeColor = System.Drawing.SystemColors.ControlText
        Me.UltraCombo1.DisplayLayout.Override.ActiveCellAppearance = Appearance26
        Appearance27.BackColor = System.Drawing.SystemColors.Highlight
        Appearance27.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.UltraCombo1.DisplayLayout.Override.ActiveRowAppearance = Appearance27
        Me.UltraCombo1.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.UltraCombo1.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance28.BackColor = System.Drawing.SystemColors.Window
        Me.UltraCombo1.DisplayLayout.Override.CardAreaAppearance = Appearance28
        Appearance29.BorderColor = System.Drawing.Color.Silver
        Appearance29.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.UltraCombo1.DisplayLayout.Override.CellAppearance = Appearance29
        Me.UltraCombo1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.UltraCombo1.DisplayLayout.Override.CellPadding = 0
        Appearance30.BackColor = System.Drawing.SystemColors.Control
        Appearance30.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance30.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance30.BorderColor = System.Drawing.SystemColors.Window
        Me.UltraCombo1.DisplayLayout.Override.GroupByRowAppearance = Appearance30
        Appearance31.TextHAlignAsString = "Left"
        Me.UltraCombo1.DisplayLayout.Override.HeaderAppearance = Appearance31
        Me.UltraCombo1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.UltraCombo1.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance32.BackColor = System.Drawing.SystemColors.Window
        Appearance32.BorderColor = System.Drawing.Color.Silver
        Me.UltraCombo1.DisplayLayout.Override.RowAppearance = Appearance32
        Me.UltraCombo1.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance33.BackColor = System.Drawing.SystemColors.ControlLight
        Me.UltraCombo1.DisplayLayout.Override.TemplateAddRowAppearance = Appearance33
        Me.UltraCombo1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.UltraCombo1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.UltraCombo1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.UltraCombo1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList
        Me.UltraCombo1.Location = New System.Drawing.Point(0, 5)
        Me.UltraCombo1.Name = "UltraCombo1"
        Me.UltraCombo1.Size = New System.Drawing.Size(174, 22)
        Me.UltraCombo1.TabIndex = 0
        '
        'spl
        '
        Me.spl.Dock = System.Windows.Forms.DockStyle.Fill
        Me.spl.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.spl.Location = New System.Drawing.Point(0, 0)
        Me.spl.Name = "spl"
        Me.spl.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'spl.Panel1
        '
        Me.spl.Panel1.Controls.Add(Me.UltraGroupBox1)
        '
        'spl.Panel2
        '
        Me.spl.Panel2.Controls.Add(Me.grdSOTSHIP1)
        Me.spl.Size = New System.Drawing.Size(779, 616)
        Me.spl.SplitterDistance = 25
        Me.spl.TabIndex = 168
        '
        'SOF3PLER
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(992, 616)
        Me.Name = "SOF3PLER"
        Me.Text = "SOF3PLER"
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ASFBASE1_Fill_Panel.ResumeLayout(False)
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdSOTSHIP1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraCombo1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.spl.Panel1.ResumeLayout(False)
        Me.spl.Panel2.ResumeLayout(False)
        CType(Me.spl, System.ComponentModel.ISupportInitialize).EndInit()
        Me.spl.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraGroupBox1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraTabPageControl5 As Infragistics.Win.UltraWinTabControl.UltraTabPageControl
    Friend WithEvents UltraCombo1 As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents spl As System.Windows.Forms.SplitContainer
    Friend WithEvents grdSOTSHIP1 As Infragistics.Win.UltraWinGrid.UltraGrid
End Class
