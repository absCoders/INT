<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ICFINVTY
    'Inherits System.Windows.Forms.Form
    Inherits ABSolution.ASFBASE1
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
        Dim UltraExplorerBarGroup1 As Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarGroup = New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarGroup()
        Dim UltraExplorerBarItem1 As Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem = New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem()
        Dim UltraExplorerBarItem2 As Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem = New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarItem()
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance19 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance21 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance22 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance23 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance24 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance25 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("ICTTRANX", -1)
        Dim UltraGridColumn36 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("OPS_YYYYPP")
        Dim UltraGridColumn37 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ITEM_CODE", -1, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, False)
        Dim UltraGridColumn38 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WHSE_CODE")
        Dim UltraGridColumn40 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TRAN_NO")
        Dim UltraGridColumn41 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TRAN_SOURCE")
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridColumn42 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_DATE")
        Dim UltraGridColumn43 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_OPER")
        Dim UltraGridColumn44 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TRAN_DATE")
        Dim UltraGridColumn45 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TRAN_TYPE")
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridColumn46 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TRAN_QTY")
        Dim UltraGridColumn47 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("TRAN_NOTE")
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
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.dteEND_DATE = New Infragistics.Win.UltraWinEditors.UltraDateTimeEditor()
        Me.dteSTART_DATE = New Infragistics.Win.UltraWinEditors.UltraDateTimeEditor()
        Me.UltraGroupBox1 = New Infragistics.Win.Misc.UltraGroupBox()
        Me.spl = New System.Windows.Forms.SplitContainer()
        Me.UltraGroupBox2 = New Infragistics.Win.Misc.UltraGroupBox()
        Me.UltraLabel4 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraLabel3 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraLabel14 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraTextEditor3 = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.txtWHSE_CODE = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel()
        Me.grdICTTRANX = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.grdSOTRMAFX = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.grdEDT812TX = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer()
        Me.grdEDT180T2 = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.grdEDT180TX = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.chk_showDetails180 = New Infragistics.Win.UltraWinEditors.UltraCheckEditor()
        Me.grdSOTRMAFI = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.tabRAMaster = New Infragistics.Win.UltraWinTabControl.UltraTabControl()
        Me.grdICTINVTY = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.UltraTabSharedControlsPage2 = New Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ASFBASE1_Fill_Panel.SuspendLayout()
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTLOGX1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dteEND_DATE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dteSTART_DATE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.spl, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.spl.Panel1.SuspendLayout()
        Me.spl.Panel2.SuspendLayout()
        Me.spl.SuspendLayout()
        CType(Me.UltraGroupBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox2.SuspendLayout()
        CType(Me.UltraTextEditor3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtWHSE_CODE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdICTTRANX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdSOTRMAFX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdEDT812TX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer4.SuspendLayout()
        CType(Me.grdEDT180T2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdEDT180TX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chk_showDetails180, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdSOTRMAFI, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tabRAMaster, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdICTINVTY, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraTabSharedControlsPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'UltraExplorerBar1
        '
        UltraExplorerBarItem1.Key = "Load"
        UltraExplorerBarItem1.Text = "Load"
        UltraExplorerBarItem2.Key = "Done"
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
        Me.UltraExplorerBar1.Size = New System.Drawing.Size(208, 1028)
        Me.UltraExplorerBar1.Tag = "CLICK"
        '
        'ASFBASE1_Fill_Panel
        '
        Me.ASFBASE1_Fill_Panel.Controls.Add(Me.spl)
        Me.ASFBASE1_Fill_Panel.Location = New System.Drawing.Point(0, 10)
        Me.ASFBASE1_Fill_Panel.Size = New System.Drawing.Size(1379, 1038)
        Me.ASFBASE1_Fill_Panel.Controls.SetChildIndex(Me.grdASFBASEX, 0)
        Me.ASFBASE1_Fill_Panel.Controls.SetChildIndex(Me.spl, 0)
        '
        'grdASFBASEX
        '
        Appearance17.BackColor = System.Drawing.SystemColors.Window
        Appearance17.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdASFBASEX.DisplayLayout.Appearance = Appearance17
        Me.grdASFBASEX.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdASFBASEX.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.MaxColScrollRegions = 1
        Me.grdASFBASEX.DisplayLayout.MaxRowScrollRegions = 1
        Appearance18.BackColor = System.Drawing.SystemColors.Window
        Appearance18.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveCellAppearance = Appearance18
        Appearance19.BackColor = System.Drawing.SystemColors.Highlight
        Appearance19.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveRowAppearance = Appearance19
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance20.BackColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.CardAreaAppearance = Appearance20
        Appearance21.BorderColor = System.Drawing.Color.Silver
        Appearance21.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdASFBASEX.DisplayLayout.Override.CellAppearance = Appearance21
        Me.grdASFBASEX.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdASFBASEX.DisplayLayout.Override.CellPadding = 0
        Appearance22.BackColor = System.Drawing.SystemColors.Control
        Appearance22.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance22.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance22.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.GroupByRowAppearance = Appearance22
        Appearance23.TextHAlignAsString = "Left"
        Me.grdASFBASEX.DisplayLayout.Override.HeaderAppearance = Appearance23
        Me.grdASFBASEX.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdASFBASEX.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance24.BackColor = System.Drawing.SystemColors.Window
        Appearance24.BorderColor = System.Drawing.Color.Silver
        Me.grdASFBASEX.DisplayLayout.Override.RowAppearance = Appearance24
        Me.grdASFBASEX.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance25.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdASFBASEX.DisplayLayout.Override.TemplateAddRowAppearance = Appearance25
        Me.grdASFBASEX.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdASFBASEX.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdASFBASEX.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdASFBASEX.Location = New System.Drawing.Point(683, 802)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 1048)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(1592, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 1048)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(1592, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 1048)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(1592, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'dteEND_DATE
        '
        Me.dteEND_DATE.Location = New System.Drawing.Point(540, 35)
        Me.dteEND_DATE.Margin = New System.Windows.Forms.Padding(4)
        Me.dteEND_DATE.Name = "dteEND_DATE"
        Me.dteEND_DATE.Size = New System.Drawing.Size(109, 25)
        Me.dteEND_DATE.TabIndex = 4
        '
        'dteSTART_DATE
        '
        Me.dteSTART_DATE.Location = New System.Drawing.Point(405, 35)
        Me.dteSTART_DATE.Margin = New System.Windows.Forms.Padding(4)
        Me.dteSTART_DATE.Name = "dteSTART_DATE"
        Me.dteSTART_DATE.Size = New System.Drawing.Size(109, 25)
        Me.dteSTART_DATE.TabIndex = 2
        '
        'UltraGroupBox1
        '
        Me.UltraGroupBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.UltraGroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.UltraGroupBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.UltraGroupBox1.Name = "UltraGroupBox1"
        Me.UltraGroupBox1.Size = New System.Drawing.Size(1379, 10)
        Me.UltraGroupBox1.TabIndex = 1
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
        Me.spl.Panel1.Controls.Add(Me.UltraGroupBox2)
        '
        'spl.Panel2
        '
        Me.spl.Panel2.Controls.Add(Me.grdICTTRANX)
        Me.spl.Size = New System.Drawing.Size(1379, 1038)
        Me.spl.SplitterDistance = 68
        Me.spl.TabIndex = 4
        '
        'UltraGroupBox2
        '
        Me.UltraGroupBox2.BorderStyle = Infragistics.Win.Misc.GroupBoxBorderStyle.Rectangular3D
        Me.UltraGroupBox2.Controls.Add(Me.UltraLabel4)
        Me.UltraGroupBox2.Controls.Add(Me.UltraLabel3)
        Me.UltraGroupBox2.Controls.Add(Me.dteEND_DATE)
        Me.UltraGroupBox2.Controls.Add(Me.UltraLabel14)
        Me.UltraGroupBox2.Controls.Add(Me.dteSTART_DATE)
        Me.UltraGroupBox2.Controls.Add(Me.UltraTextEditor3)
        Me.UltraGroupBox2.Controls.Add(Me.txtWHSE_CODE)
        Me.UltraGroupBox2.Controls.Add(Me.UltraLabel1)
        Me.UltraGroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UltraGroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.UltraGroupBox2.Name = "UltraGroupBox2"
        Me.UltraGroupBox2.Size = New System.Drawing.Size(1379, 68)
        Me.UltraGroupBox2.TabIndex = 6
        '
        'UltraLabel4
        '
        Me.UltraLabel4.AutoSize = True
        Me.UltraLabel4.Location = New System.Drawing.Point(405, 12)
        Me.UltraLabel4.Name = "UltraLabel4"
        Me.UltraLabel4.Size = New System.Drawing.Size(75, 18)
        Me.UltraLabel4.TabIndex = 104
        Me.UltraLabel4.Text = "Start Date"
        '
        'UltraLabel3
        '
        Me.UltraLabel3.AutoSize = True
        Me.UltraLabel3.Location = New System.Drawing.Point(540, 12)
        Me.UltraLabel3.Name = "UltraLabel3"
        Me.UltraLabel3.Size = New System.Drawing.Size(66, 18)
        Me.UltraLabel3.TabIndex = 103
        Me.UltraLabel3.Text = "End Date"
        '
        'UltraLabel14
        '
        Me.UltraLabel14.AutoSize = True
        Me.UltraLabel14.Location = New System.Drawing.Point(144, 12)
        Me.UltraLabel14.Name = "UltraLabel14"
        Me.UltraLabel14.Size = New System.Drawing.Size(44, 18)
        Me.UltraLabel14.TabIndex = 102
        Me.UltraLabel14.Text = "Name"
        '
        'UltraTextEditor3
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor3, "WHSE_DESC")
        Me.Absx1.SetABSParentColumnName(Me.UltraTextEditor3, "WHSE_CODE")
        Me.UltraTextEditor3.Location = New System.Drawing.Point(144, 36)
        Me.UltraTextEditor3.Name = "UltraTextEditor3"
        Me.UltraTextEditor3.ReadOnly = True
        Me.UltraTextEditor3.Size = New System.Drawing.Size(230, 25)
        Me.UltraTextEditor3.TabIndex = 5
        Me.UltraTextEditor3.TabStop = False
        '
        'txtWHSE_CODE
        '
        Me.Absx1.SetABSColumnName(Me.txtWHSE_CODE, "WHSE_CODE")
        Me.Absx1.SetABSHasButton(Me.txtWHSE_CODE, True)
        Me.txtWHSE_CODE.Location = New System.Drawing.Point(13, 36)
        Me.txtWHSE_CODE.Name = "txtWHSE_CODE"
        Me.txtWHSE_CODE.Size = New System.Drawing.Size(125, 25)
        Me.txtWHSE_CODE.TabIndex = 0
        '
        'UltraLabel1
        '
        Me.UltraLabel1.AutoSize = True
        Me.UltraLabel1.Location = New System.Drawing.Point(13, 12)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(80, 18)
        Me.UltraLabel1.TabIndex = 3
        Me.UltraLabel1.Text = "Warehouse"
        '
        'grdICTTRANX
        '
        Appearance1.BackColor = System.Drawing.SystemColors.Window
        Appearance1.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdICTTRANX.DisplayLayout.Appearance = Appearance1
        UltraGridColumn36.Header.VisiblePosition = 0
        UltraGridColumn36.Hidden = True
        UltraGridColumn37.Header.Caption = "Item Code"
        UltraGridColumn37.Header.VisiblePosition = 1
        UltraGridColumn38.Header.Caption = "Whse"
        UltraGridColumn38.Header.VisiblePosition = 2
        UltraGridColumn38.Hidden = True
        UltraGridColumn38.Width = 64
        UltraGridColumn40.Header.Caption = "Tran No"
        UltraGridColumn40.Header.VisiblePosition = 3
        UltraGridColumn40.Width = 98
        Appearance2.TextHAlignAsString = "Center"
        UltraGridColumn41.CellAppearance = Appearance2
        Appearance3.TextHAlignAsString = "Center"
        UltraGridColumn41.Header.Appearance = Appearance3
        UltraGridColumn41.Header.Caption = "Src"
        UltraGridColumn41.Header.VisiblePosition = 4
        UltraGridColumn41.Width = 52
        UltraGridColumn42.Format = "MM/dd/yy HH:mm"
        UltraGridColumn42.Header.Caption = "Entered"
        UltraGridColumn42.Header.VisiblePosition = 7
        UltraGridColumn42.Width = 124
        UltraGridColumn43.Header.Caption = "By"
        UltraGridColumn43.Header.VisiblePosition = 8
        UltraGridColumn43.Width = 65
        UltraGridColumn44.Header.Caption = "Date"
        UltraGridColumn44.Header.VisiblePosition = 5
        UltraGridColumn44.Width = 100
        Appearance4.TextHAlignAsString = "Center"
        UltraGridColumn45.CellAppearance = Appearance4
        Appearance5.TextHAlignAsString = "Center"
        UltraGridColumn45.Header.Appearance = Appearance5
        UltraGridColumn45.Header.Caption = "Type"
        UltraGridColumn45.Header.VisiblePosition = 6
        UltraGridColumn45.Width = 54
        UltraGridColumn46.Header.Caption = "Qty"
        UltraGridColumn46.Header.VisiblePosition = 9
        UltraGridColumn46.Width = 84
        UltraGridColumn47.Header.Caption = "Details"
        UltraGridColumn47.Header.VisiblePosition = 10
        UltraGridColumn47.Width = 240
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn36, UltraGridColumn37, UltraGridColumn38, UltraGridColumn40, UltraGridColumn41, UltraGridColumn42, UltraGridColumn43, UltraGridColumn44, UltraGridColumn45, UltraGridColumn46, UltraGridColumn47})
        Me.grdICTTRANX.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdICTTRANX.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance6.TextHAlignAsString = "Left"
        Me.grdICTTRANX.DisplayLayout.CaptionAppearance = Appearance6
        Appearance7.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance7.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance7.BorderColor = System.Drawing.SystemColors.Window
        Me.grdICTTRANX.DisplayLayout.GroupByBox.Appearance = Appearance7
        Appearance8.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdICTTRANX.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance8
        Me.grdICTTRANX.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdICTTRANX.DisplayLayout.GroupByBox.Hidden = True
        Appearance9.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance9.BackColor2 = System.Drawing.SystemColors.Control
        Appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance9.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdICTTRANX.DisplayLayout.GroupByBox.PromptAppearance = Appearance9
        Me.grdICTTRANX.DisplayLayout.MaxColScrollRegions = 1
        Me.grdICTTRANX.DisplayLayout.MaxRowScrollRegions = 1
        Appearance10.BackColor = System.Drawing.SystemColors.Window
        Appearance10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdICTTRANX.DisplayLayout.Override.ActiveCellAppearance = Appearance10
        Me.grdICTTRANX.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdICTTRANX.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdICTTRANX.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdICTTRANX.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdICTTRANX.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance11.BackColor = System.Drawing.SystemColors.Window
        Me.grdICTTRANX.DisplayLayout.Override.CardAreaAppearance = Appearance11
        Appearance12.BorderColor = System.Drawing.Color.Silver
        Appearance12.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdICTTRANX.DisplayLayout.Override.CellAppearance = Appearance12
        Me.grdICTTRANX.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdICTTRANX.DisplayLayout.Override.CellPadding = 0
        Appearance13.BackColor = System.Drawing.SystemColors.Control
        Appearance13.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance13.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance13.BorderColor = System.Drawing.SystemColors.Window
        Me.grdICTTRANX.DisplayLayout.Override.GroupByRowAppearance = Appearance13
        Appearance14.TextHAlignAsString = "Left"
        Me.grdICTTRANX.DisplayLayout.Override.HeaderAppearance = Appearance14
        Me.grdICTTRANX.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdICTTRANX.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance15.BackColor = System.Drawing.SystemColors.Window
        Appearance15.BorderColor = System.Drawing.Color.Silver
        Me.grdICTTRANX.DisplayLayout.Override.RowAppearance = Appearance15
        Me.grdICTTRANX.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance16.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdICTTRANX.DisplayLayout.Override.TemplateAddRowAppearance = Appearance16
        Me.grdICTTRANX.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdICTTRANX.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdICTTRANX.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdICTTRANX.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdICTTRANX.Location = New System.Drawing.Point(0, 0)
        Me.grdICTTRANX.Name = "grdICTTRANX"
        Me.grdICTTRANX.Size = New System.Drawing.Size(1379, 966)
        Me.grdICTTRANX.TabIndex = 166
        Me.grdICTTRANX.Text = "Transactions"
        '
        'grdSOTRMAFX
        '
        Me.grdSOTRMAFX.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdSOTRMAFX.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdSOTRMAFX.Location = New System.Drawing.Point(0, 0)
        Me.grdSOTRMAFX.Name = "grdSOTRMAFX"
        Me.grdSOTRMAFX.Size = New System.Drawing.Size(1371, 910)
        Me.grdSOTRMAFX.TabIndex = 13
        '
        'grdEDT812TX
        '
        Me.grdEDT812TX.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdEDT812TX.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdEDT812TX.Location = New System.Drawing.Point(0, 0)
        Me.grdEDT812TX.Name = "grdEDT812TX"
        Me.grdEDT812TX.Size = New System.Drawing.Size(1371, 910)
        Me.grdEDT812TX.TabIndex = 17
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer4.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer4.Name = "SplitContainer4"
        Me.SplitContainer4.Orientation = System.Windows.Forms.Orientation.Horizontal
        Me.SplitContainer4.Size = New System.Drawing.Size(1371, 910)
        Me.SplitContainer4.SplitterDistance = 439
        Me.SplitContainer4.TabIndex = 0
        '
        'grdEDT180T2
        '
        Me.grdEDT180T2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdEDT180T2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdEDT180T2.Location = New System.Drawing.Point(0, 0)
        Me.grdEDT180T2.Name = "grdEDT180T2"
        Me.grdEDT180T2.Size = New System.Drawing.Size(1371, 466)
        Me.grdEDT180T2.TabIndex = 15
        '
        'grdEDT180TX
        '
        Me.grdEDT180TX.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdEDT180TX.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdEDT180TX.Location = New System.Drawing.Point(0, 0)
        Me.grdEDT180TX.Name = "grdEDT180TX"
        Me.grdEDT180TX.Size = New System.Drawing.Size(1371, 440)
        Me.grdEDT180TX.TabIndex = 15
        '
        'chk_showDetails180
        '
        Me.chk_showDetails180.Location = New System.Drawing.Point(573, 2)
        Me.chk_showDetails180.Name = "chk_showDetails180"
        Me.chk_showDetails180.Size = New System.Drawing.Size(121, 18)
        Me.chk_showDetails180.TabIndex = 195
        '
        'grdSOTRMAFI
        '
        Me.grdSOTRMAFI.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdSOTRMAFI.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdSOTRMAFI.Location = New System.Drawing.Point(0, 0)
        Me.grdSOTRMAFI.Name = "grdSOTRMAFI"
        Me.grdSOTRMAFI.Size = New System.Drawing.Size(1371, 910)
        Me.grdSOTRMAFI.TabIndex = 14
        '
        'tabRAMaster
        '
        Me.tabRAMaster.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabRAMaster.Location = New System.Drawing.Point(0, 0)
        Me.tabRAMaster.Name = "tabRAMaster"
        Me.tabRAMaster.SharedControls.AddRange(New System.Windows.Forms.Control() {Me.grdICTINVTY})
        Me.tabRAMaster.SharedControlsPage = Me.UltraTabSharedControlsPage2
        Me.tabRAMaster.Size = New System.Drawing.Size(1375, 938)
        Me.tabRAMaster.TabIndex = 14
        Me.tabRAMaster.TabOrientation = Infragistics.Win.UltraWinTabs.TabOrientation.BottomLeft
        '
        'grdICTINVTY
        '
        Me.grdICTINVTY.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdICTINVTY.Location = New System.Drawing.Point(0, 0)
        Me.grdICTINVTY.Name = "grdICTINVTY"
        Me.grdICTINVTY.Size = New System.Drawing.Size(1371, 915)
        Me.grdICTINVTY.TabIndex = 167
        '
        'UltraTabSharedControlsPage2
        '
        Me.UltraTabSharedControlsPage2.Controls.Add(Me.grdICTINVTY)
        Me.UltraTabSharedControlsPage2.Location = New System.Drawing.Point(1, 1)
        Me.UltraTabSharedControlsPage2.Name = "UltraTabSharedControlsPage2"
        Me.UltraTabSharedControlsPage2.Size = New System.Drawing.Size(1371, 915)
        '
        'ICFINVTY
        '
        Me.Absx1.SetABSBindToTable(Me, False)
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1592, 1048)
        Me.Controls.Add(Me.UltraGroupBox1)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "ICFINVTY"
        Me.Text = "SOFORELC"
        Me.Controls.SetChildIndex(Me._ASFBASE1_Toolbars_Dock_Area_Top, 0)
        Me.Controls.SetChildIndex(Me._ASFBASE1_Toolbars_Dock_Area_Bottom, 0)
        Me.Controls.SetChildIndex(Me._ASFBASE1_Toolbars_Dock_Area_Right, 0)
        Me.Controls.SetChildIndex(Me._ASFBASE1_Toolbars_Dock_Area_Left, 0)
        Me.Controls.SetChildIndex(Me.UltraGroupBox1, 0)
        Me.Controls.SetChildIndex(Me.ASFBASE1_Fill_Panel, 0)
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ASFBASE1_Fill_Panel.ResumeLayout(False)
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTLOGX1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dteEND_DATE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dteSTART_DATE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.spl.Panel1.ResumeLayout(False)
        Me.spl.Panel2.ResumeLayout(False)
        CType(Me.spl, System.ComponentModel.ISupportInitialize).EndInit()
        Me.spl.ResumeLayout(False)
        CType(Me.UltraGroupBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox2.ResumeLayout(False)
        Me.UltraGroupBox2.PerformLayout()
        CType(Me.UltraTextEditor3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtWHSE_CODE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdICTTRANX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdSOTRMAFX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdEDT812TX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SplitContainer4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer4.ResumeLayout(False)
        CType(Me.grdEDT180T2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdEDT180TX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chk_showDetails180, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdSOTRMAFI, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tabRAMaster, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdICTINVTY, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraTabSharedControlsPage2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraGroupBox1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents dteSTART_DATE As Infragistics.Win.UltraWinEditors.UltraDateTimeEditor
    Friend WithEvents dteEND_DATE As UltraWinEditors.UltraDateTimeEditor
    Friend WithEvents spl As SplitContainer
    Friend WithEvents UltraGroupBox2 As Misc.UltraGroupBox
    Friend WithEvents UltraLabel14 As Misc.UltraLabel
    Friend WithEvents UltraTextEditor3 As UltraWinEditors.UltraTextEditor
    Friend WithEvents txtWHSE_CODE As UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Misc.UltraLabel
    Friend WithEvents grdSOTRMAFX As UltraWinGrid.UltraGrid
    Friend WithEvents grdEDT812TX As UltraWinGrid.UltraGrid
    Friend WithEvents SplitContainer4 As SplitContainer
    Friend WithEvents grdEDT180T2 As UltraWinGrid.UltraGrid
    Friend WithEvents grdEDT180TX As UltraWinGrid.UltraGrid
    Friend WithEvents chk_showDetails180 As UltraWinEditors.UltraCheckEditor
    Friend WithEvents grdSOTRMAFI As UltraWinGrid.UltraGrid
    Friend WithEvents UltraLabel4 As Misc.UltraLabel
    Friend WithEvents UltraLabel3 As Misc.UltraLabel
    Friend WithEvents tabRAMaster As UltraWinTabControl.UltraTabControl
    Friend WithEvents grdICTINVTY As UltraWinGrid.UltraGrid
    Friend WithEvents UltraTabSharedControlsPage2 As UltraWinTabControl.UltraTabSharedControlsPage
    Friend WithEvents grdICTTRANX As UltraWinGrid.UltraGrid
End Class
