<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ASFBASE1
    'Inherits System.Windows.Forms.Form
    Inherits ABSolution.ASFBASE0
    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DockAreaPane1 As Infragistics.Win.UltraWinDock.DockAreaPane = New Infragistics.Win.UltraWinDock.DockAreaPane(Infragistics.Win.UltraWinDock.DockedLocation.DockedRight, New System.Guid("ec7f55a1-a731-409e-ae08-e27066699a58"))
        Dim DockableControlPane1 As Infragistics.Win.UltraWinDock.DockableControlPane = New Infragistics.Win.UltraWinDock.DockableControlPane(New System.Guid("9b9ecd1f-75b5-4c88-9639-bd32f6f1f31c"), New System.Guid("de513fba-e451-4498-9ded-1fc349a589ae"), -1, New System.Guid("ec7f55a1-a731-409e-ae08-e27066699a58"), 0)
        Dim DockAreaPane2 As Infragistics.Win.UltraWinDock.DockAreaPane = New Infragistics.Win.UltraWinDock.DockAreaPane(Infragistics.Win.UltraWinDock.DockedLocation.Floating, New System.Guid("de513fba-e451-4498-9ded-1fc349a589ae"))
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
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
        Me.UltraExplorerBar1 = New Infragistics.Win.UltraWinExplorerBar.UltraExplorerBar()
        Me.UltraDockManager1 = New Infragistics.Win.UltraWinDock.UltraDockManager(Me.components)
        Me._ASFBASE1UnpinnedTabAreaLeft = New Infragistics.Win.UltraWinDock.UnpinnedTabArea()
        Me._ASFBASE1UnpinnedTabAreaRight = New Infragistics.Win.UltraWinDock.UnpinnedTabArea()
        Me._ASFBASE1UnpinnedTabAreaTop = New Infragistics.Win.UltraWinDock.UnpinnedTabArea()
        Me._ASFBASE1UnpinnedTabAreaBottom = New Infragistics.Win.UltraWinDock.UnpinnedTabArea()
        Me._ASFBASE1AutoHideControl = New Infragistics.Win.UltraWinDock.AutoHideControl()
        Me.DockableWindow1 = New Infragistics.Win.UltraWinDock.DockableWindow()
        Me.WindowDockingArea2 = New Infragistics.Win.UltraWinDock.WindowDockingArea()
        Me.WindowDockingArea3 = New Infragistics.Win.UltraWinDock.WindowDockingArea()
        Me.ASFBASE1_Fill_Panel = New System.Windows.Forms.Panel()
        Me.grdASFBASEX = New Infragistics.Win.UltraWinGrid.UltraGrid()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraDockManager1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.DockableWindow1.SuspendLayout()
        Me.WindowDockingArea2.SuspendLayout()
        Me.ASFBASE1_Fill_Panel.SuspendLayout()
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'UltraExplorerBar1
        '
        Me.UltraExplorerBar1.AnimationEnabled = False
        Me.UltraExplorerBar1.AnimationSpeed = Infragistics.Win.UltraWinExplorerBar.AnimationSpeed.Fast
        Me.UltraExplorerBar1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UltraExplorerBar1.GroupSettings.UseMnemonics = Infragistics.Win.DefaultableBoolean.[True]
        Me.UltraExplorerBar1.ItemSettings.Style = Infragistics.Win.UltraWinExplorerBar.ItemStyle.Button
        Me.UltraExplorerBar1.Location = New System.Drawing.Point(0, 20)
        Me.UltraExplorerBar1.Name = "UltraExplorerBar1"
        Me.UltraExplorerBar1.Size = New System.Drawing.Size(208, 596)
        Me.UltraExplorerBar1.Style = Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarStyle.VisualStudio2005Toolbox
        Me.UltraExplorerBar1.TabIndex = 0
        Me.UltraExplorerBar1.TabStop = False
        Me.UltraExplorerBar1.ViewStyle = Infragistics.Win.UltraWinExplorerBar.UltraExplorerBarViewStyle.XPExplorerBar
        '
        'UltraDockManager1
        '
        Me.UltraDockManager1.AnimationEnabled = False
        Me.UltraDockManager1.AutoHideDelay = 100
        Me.UltraDockManager1.CompressUnpinnedTabs = False
        DockAreaPane1.ChildPaneStyle = Infragistics.Win.UltraWinDock.ChildPaneStyle.TabGroup
        DockAreaPane1.DockedBefore = New System.Guid("de513fba-e451-4498-9ded-1fc349a589ae")
        DockAreaPane1.FloatingLocation = New System.Drawing.Point(855, 163)
        DockableControlPane1.Control = Me.UltraExplorerBar1
        DockableControlPane1.OriginalControlBounds = New System.Drawing.Rectangle(757, 0, 233, 574)
        DockableControlPane1.Settings.AllowClose = Infragistics.Win.DefaultableBoolean.[False]
        DockableControlPane1.Settings.AllowDockAsTab = Infragistics.Win.DefaultableBoolean.[True]
        DockableControlPane1.Settings.AllowDockBottom = Infragistics.Win.DefaultableBoolean.[False]
        DockableControlPane1.Settings.AllowDockLeft = Infragistics.Win.DefaultableBoolean.[False]
        DockableControlPane1.Settings.AllowDockRight = Infragistics.Win.DefaultableBoolean.[True]
        DockableControlPane1.Settings.AllowDockTop = Infragistics.Win.DefaultableBoolean.[False]
        DockableControlPane1.Settings.AllowMaximize = Infragistics.Win.DefaultableBoolean.[False]
        DockableControlPane1.Settings.AllowMinimize = Infragistics.Win.DefaultableBoolean.[False]
        DockableControlPane1.Size = New System.Drawing.Size(100, 100)
        DockableControlPane1.Text = "Control Panel"
        DockAreaPane1.Panes.AddRange(New Infragistics.Win.UltraWinDock.DockablePaneBase() {DockableControlPane1})
        DockAreaPane1.Size = New System.Drawing.Size(208, 616)
        DockAreaPane2.ChildPaneStyle = Infragistics.Win.UltraWinDock.ChildPaneStyle.TabGroup
        DockAreaPane2.SelectedTabIndex = -1
        DockAreaPane2.Size = New System.Drawing.Size(100, 100)
        Me.UltraDockManager1.DockAreas.AddRange(New Infragistics.Win.UltraWinDock.DockAreaPane() {DockAreaPane1, DockAreaPane2})
        Me.UltraDockManager1.HostControl = Me
        '
        '_ASFBASE1UnpinnedTabAreaLeft
        '
        Me._ASFBASE1UnpinnedTabAreaLeft.Dock = System.Windows.Forms.DockStyle.Left
        Me._ASFBASE1UnpinnedTabAreaLeft.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._ASFBASE1UnpinnedTabAreaLeft.Location = New System.Drawing.Point(0, 0)
        Me._ASFBASE1UnpinnedTabAreaLeft.Name = "_ASFBASE1UnpinnedTabAreaLeft"
        Me._ASFBASE1UnpinnedTabAreaLeft.Owner = Me.UltraDockManager1
        Me._ASFBASE1UnpinnedTabAreaLeft.Size = New System.Drawing.Size(0, 616)
        Me._ASFBASE1UnpinnedTabAreaLeft.TabIndex = 1
        '
        '_ASFBASE1UnpinnedTabAreaRight
        '
        Me._ASFBASE1UnpinnedTabAreaRight.Dock = System.Windows.Forms.DockStyle.Right
        Me._ASFBASE1UnpinnedTabAreaRight.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._ASFBASE1UnpinnedTabAreaRight.Location = New System.Drawing.Point(992, 0)
        Me._ASFBASE1UnpinnedTabAreaRight.Name = "_ASFBASE1UnpinnedTabAreaRight"
        Me._ASFBASE1UnpinnedTabAreaRight.Owner = Me.UltraDockManager1
        Me._ASFBASE1UnpinnedTabAreaRight.Size = New System.Drawing.Size(0, 616)
        Me._ASFBASE1UnpinnedTabAreaRight.TabIndex = 2
        '
        '_ASFBASE1UnpinnedTabAreaTop
        '
        Me._ASFBASE1UnpinnedTabAreaTop.Dock = System.Windows.Forms.DockStyle.Top
        Me._ASFBASE1UnpinnedTabAreaTop.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._ASFBASE1UnpinnedTabAreaTop.Location = New System.Drawing.Point(0, 0)
        Me._ASFBASE1UnpinnedTabAreaTop.Name = "_ASFBASE1UnpinnedTabAreaTop"
        Me._ASFBASE1UnpinnedTabAreaTop.Owner = Me.UltraDockManager1
        Me._ASFBASE1UnpinnedTabAreaTop.Size = New System.Drawing.Size(992, 0)
        Me._ASFBASE1UnpinnedTabAreaTop.TabIndex = 3
        '
        '_ASFBASE1UnpinnedTabAreaBottom
        '
        Me._ASFBASE1UnpinnedTabAreaBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me._ASFBASE1UnpinnedTabAreaBottom.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._ASFBASE1UnpinnedTabAreaBottom.Location = New System.Drawing.Point(0, 616)
        Me._ASFBASE1UnpinnedTabAreaBottom.Name = "_ASFBASE1UnpinnedTabAreaBottom"
        Me._ASFBASE1UnpinnedTabAreaBottom.Owner = Me.UltraDockManager1
        Me._ASFBASE1UnpinnedTabAreaBottom.Size = New System.Drawing.Size(992, 0)
        Me._ASFBASE1UnpinnedTabAreaBottom.TabIndex = 4
        '
        '_ASFBASE1AutoHideControl
        '
        Me._ASFBASE1AutoHideControl.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._ASFBASE1AutoHideControl.Location = New System.Drawing.Point(0, 0)
        Me._ASFBASE1AutoHideControl.Name = "_ASFBASE1AutoHideControl"
        Me._ASFBASE1AutoHideControl.Owner = Me.UltraDockManager1
        Me._ASFBASE1AutoHideControl.Size = New System.Drawing.Size(0, 0)
        Me._ASFBASE1AutoHideControl.TabIndex = 5
        '
        'DockableWindow1
        '
        Me.DockableWindow1.Controls.Add(Me.UltraExplorerBar1)
        Me.DockableWindow1.Location = New System.Drawing.Point(5, 0)
        Me.DockableWindow1.Name = "DockableWindow1"
        Me.DockableWindow1.Owner = Me.UltraDockManager1
        Me.DockableWindow1.Size = New System.Drawing.Size(208, 616)
        Me.DockableWindow1.TabIndex = 6
        '
        'WindowDockingArea2
        '
        Me.WindowDockingArea2.Controls.Add(Me.DockableWindow1)
        Me.WindowDockingArea2.Dock = System.Windows.Forms.DockStyle.Right
        Me.WindowDockingArea2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.WindowDockingArea2.Location = New System.Drawing.Point(779, 0)
        Me.WindowDockingArea2.Name = "WindowDockingArea2"
        Me.WindowDockingArea2.Owner = Me.UltraDockManager1
        Me.WindowDockingArea2.Size = New System.Drawing.Size(213, 616)
        Me.WindowDockingArea2.TabIndex = 0
        '
        'WindowDockingArea3
        '
        Me.WindowDockingArea3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WindowDockingArea3.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.WindowDockingArea3.Location = New System.Drawing.Point(777, 0)
        Me.WindowDockingArea3.Name = "WindowDockingArea3"
        Me.WindowDockingArea3.Owner = Me.UltraDockManager1
        Me.WindowDockingArea3.Size = New System.Drawing.Size(100, 100)
        Me.WindowDockingArea3.TabIndex = 7
        '
        'ASFBASE1_Fill_Panel
        '
        Me.ASFBASE1_Fill_Panel.Controls.Add(Me.grdASFBASEX)
        Me.ASFBASE1_Fill_Panel.Cursor = System.Windows.Forms.Cursors.Default
        Me.ASFBASE1_Fill_Panel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ASFBASE1_Fill_Panel.Location = New System.Drawing.Point(0, 0)
        Me.ASFBASE1_Fill_Panel.Name = "ASFBASE1_Fill_Panel"
        Me.ASFBASE1_Fill_Panel.Size = New System.Drawing.Size(779, 616)
        Me.ASFBASE1_Fill_Panel.TabIndex = 0
        '
        'grdASFBASEX
        '
        Appearance1.BackColor = System.Drawing.SystemColors.Window
        Appearance1.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdASFBASEX.DisplayLayout.Appearance = Appearance1
        Me.grdASFBASEX.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdASFBASEX.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance2.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance2.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance2.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance2.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.GroupByBox.Appearance = Appearance2
        Appearance3.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdASFBASEX.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance3
        Me.grdASFBASEX.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance4.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance4.BackColor2 = System.Drawing.SystemColors.Control
        Appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance4.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdASFBASEX.DisplayLayout.GroupByBox.PromptAppearance = Appearance4
        Me.grdASFBASEX.DisplayLayout.MaxColScrollRegions = 1
        Me.grdASFBASEX.DisplayLayout.MaxRowScrollRegions = 1
        Appearance5.BackColor = System.Drawing.SystemColors.Window
        Appearance5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveCellAppearance = Appearance5
        Appearance6.BackColor = System.Drawing.SystemColors.Highlight
        Appearance6.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.grdASFBASEX.DisplayLayout.Override.ActiveRowAppearance = Appearance6
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdASFBASEX.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance7.BackColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.CardAreaAppearance = Appearance7
        Appearance8.BorderColor = System.Drawing.Color.Silver
        Appearance8.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdASFBASEX.DisplayLayout.Override.CellAppearance = Appearance8
        Me.grdASFBASEX.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdASFBASEX.DisplayLayout.Override.CellPadding = 0
        Appearance9.BackColor = System.Drawing.SystemColors.Control
        Appearance9.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance9.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance9.BorderColor = System.Drawing.SystemColors.Window
        Me.grdASFBASEX.DisplayLayout.Override.GroupByRowAppearance = Appearance9
        Appearance10.TextHAlignAsString = "Left"
        Me.grdASFBASEX.DisplayLayout.Override.HeaderAppearance = Appearance10
        Me.grdASFBASEX.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdASFBASEX.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance11.BackColor = System.Drawing.SystemColors.Window
        Appearance11.BorderColor = System.Drawing.Color.Silver
        Me.grdASFBASEX.DisplayLayout.Override.RowAppearance = Appearance11
        Me.grdASFBASEX.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[False]
        Appearance12.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdASFBASEX.DisplayLayout.Override.TemplateAddRowAppearance = Appearance12
        Me.grdASFBASEX.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdASFBASEX.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdASFBASEX.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdASFBASEX.Location = New System.Drawing.Point(920, 3)
        Me.grdASFBASEX.Name = "grdASFBASEX"
        Me.grdASFBASEX.Size = New System.Drawing.Size(57, 46)
        Me.grdASFBASEX.TabIndex = 1
        Me.grdASFBASEX.Visible = False
        '
        'ASFBASE1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(206, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(992, 616)
        Me.Controls.Add(Me._ASFBASE1AutoHideControl)
        Me.Controls.Add(Me.ASFBASE1_Fill_Panel)
        Me.Controls.Add(Me.WindowDockingArea2)
        Me.Controls.Add(Me._ASFBASE1UnpinnedTabAreaTop)
        Me.Controls.Add(Me._ASFBASE1UnpinnedTabAreaBottom)
        Me.Controls.Add(Me._ASFBASE1UnpinnedTabAreaLeft)
        Me.Controls.Add(Me._ASFBASE1UnpinnedTabAreaRight)
        Me.Name = "ASFBASE1"
        Me.Text = "ASFBASE1"
        Me.Controls.SetChildIndex(Me._ASFBASE1_Toolbars_Dock_Area_Bottom, 0)
        Me.Controls.SetChildIndex(Me._ASFBASE1_Toolbars_Dock_Area_Top, 0)
        Me.Controls.SetChildIndex(Me._ASFBASE1_Toolbars_Dock_Area_Right, 0)
        Me.Controls.SetChildIndex(Me._ASFBASE1_Toolbars_Dock_Area_Left, 0)
        Me.Controls.SetChildIndex(Me._ASFBASE1UnpinnedTabAreaRight, 0)
        Me.Controls.SetChildIndex(Me._ASFBASE1UnpinnedTabAreaLeft, 0)
        Me.Controls.SetChildIndex(Me._ASFBASE1UnpinnedTabAreaBottom, 0)
        Me.Controls.SetChildIndex(Me._ASFBASE1UnpinnedTabAreaTop, 0)
        Me.Controls.SetChildIndex(Me.WindowDockingArea2, 0)
        Me.Controls.SetChildIndex(Me.ASFBASE1_Fill_Panel, 0)
        Me.Controls.SetChildIndex(Me._ASFBASE1AutoHideControl, 0)
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraDockManager1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.DockableWindow1.ResumeLayout(False)
        Me.WindowDockingArea2.ResumeLayout(False)
        Me.ASFBASE1_Fill_Panel.ResumeLayout(False)
        CType(Me.grdASFBASEX, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Protected Friend WithEvents UltraExplorerBar1 As Infragistics.Win.UltraWinExplorerBar.UltraExplorerBar
    Protected Friend WithEvents ASFBASE1_Fill_Panel As System.Windows.Forms.Panel
    Private WithEvents _ASFBASE1AutoHideControl As Infragistics.Win.UltraWinDock.AutoHideControl
    Private WithEvents UltraDockManager1 As Infragistics.Win.UltraWinDock.UltraDockManager
    Private WithEvents WindowDockingArea3 As Infragistics.Win.UltraWinDock.WindowDockingArea
    Private WithEvents DockableWindow1 As Infragistics.Win.UltraWinDock.DockableWindow
    Private WithEvents _ASFBASE1UnpinnedTabAreaTop As Infragistics.Win.UltraWinDock.UnpinnedTabArea
    Private WithEvents _ASFBASE1UnpinnedTabAreaBottom As Infragistics.Win.UltraWinDock.UnpinnedTabArea
    Private WithEvents _ASFBASE1UnpinnedTabAreaLeft As Infragistics.Win.UltraWinDock.UnpinnedTabArea
    Private WithEvents _ASFBASE1UnpinnedTabAreaRight As Infragistics.Win.UltraWinDock.UnpinnedTabArea
    Private WithEvents WindowDockingArea2 As Infragistics.Win.UltraWinDock.WindowDockingArea
    Public WithEvents grdASFBASEX As Infragistics.Win.UltraWinGrid.UltraGrid
End Class
