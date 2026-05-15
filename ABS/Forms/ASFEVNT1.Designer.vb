<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ASFEVNT1
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
        Dim Appearance22 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("TATEVNT1", -1)
        Dim UltraGridColumn14 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_NO")
        Dim UltraGridColumn107 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_DATE")
        Dim UltraGridColumn108 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_OPER")
        Dim UltraGridColumn109 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EVENT_CODE")
        Dim UltraGridColumn110 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EVENT_DESC")
        Dim UltraGridColumn111 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("EVENT_KEY")
        Dim UltraGridColumn112 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ATTACHMENT_EXT")
        Dim Appearance23 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance38 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance39 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance40 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance41 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance42 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance43 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance44 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance46 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance47 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance48 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.cmdCancel = New Infragistics.Win.Misc.UltraButton()
        Me.cmdOK = New Infragistics.Win.Misc.UltraButton()
        Me.grdTATEVNT1 = New Infragistics.Win.UltraWinGrid.UltraGrid()
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
        CType(Me.grdTATEVNT1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.SplitContainer1)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(1030, 327)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(1030, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(1030, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(1030, 0)
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.grdTATEVNT1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.cmdCancel)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cmdOK)
        Me.SplitContainer1.Size = New System.Drawing.Size(1030, 327)
        Me.SplitContainer1.SplitterDistance = 280
        Me.SplitContainer1.TabIndex = 0
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.Location = New System.Drawing.Point(952, 3)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 34)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "Cancel"
        '
        'cmdOK
        '
        Me.cmdOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOK.Location = New System.Drawing.Point(871, 3)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(75, 34)
        Me.cmdOK.TabIndex = 2
        Me.cmdOK.Text = "OK"
        '
        'grdTATEVNT1
        '
        Appearance22.BackColor = System.Drawing.SystemColors.Window
        Appearance22.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdTATEVNT1.DisplayLayout.Appearance = Appearance22
        UltraGridColumn14.Header.VisiblePosition = 0
        UltraGridColumn14.Hidden = True
        UltraGridColumn107.Format = "MM/dd/yyyy HH:mm"
        UltraGridColumn107.Header.Caption = "When"
        UltraGridColumn107.Header.VisiblePosition = 1
        UltraGridColumn107.Width = 183
        UltraGridColumn108.Header.Caption = "Who"
        UltraGridColumn108.Header.VisiblePosition = 2
        UltraGridColumn109.Header.Caption = "What"
        UltraGridColumn109.Header.VisiblePosition = 3
        UltraGridColumn110.Header.Caption = "Description"
        UltraGridColumn110.Header.VisiblePosition = 4
        UltraGridColumn110.Width = 257
        UltraGridColumn111.Header.Caption = "Ctl No"
        UltraGridColumn111.Header.VisiblePosition = 5
        UltraGridColumn112.Header.VisiblePosition = 6
        UltraGridColumn112.Hidden = True
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn14, UltraGridColumn107, UltraGridColumn108, UltraGridColumn109, UltraGridColumn110, UltraGridColumn111, UltraGridColumn112})
        Me.grdTATEVNT1.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdTATEVNT1.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance23.TextHAlignAsString = "Left"
        Me.grdTATEVNT1.DisplayLayout.CaptionAppearance = Appearance23
        Me.grdTATEVNT1.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[True]
        Appearance38.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance38.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance38.BorderColor = System.Drawing.SystemColors.Window
        Me.grdTATEVNT1.DisplayLayout.GroupByBox.Appearance = Appearance38
        Appearance39.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdTATEVNT1.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance39
        Me.grdTATEVNT1.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdTATEVNT1.DisplayLayout.GroupByBox.Hidden = True
        Appearance40.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance40.BackColor2 = System.Drawing.SystemColors.Control
        Appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance40.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdTATEVNT1.DisplayLayout.GroupByBox.PromptAppearance = Appearance40
        Me.grdTATEVNT1.DisplayLayout.MaxColScrollRegions = 1
        Me.grdTATEVNT1.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdTATEVNT1.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance41.BackColor = System.Drawing.SystemColors.Window
        Appearance41.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdTATEVNT1.DisplayLayout.Override.ActiveCellAppearance = Appearance41
        Me.grdTATEVNT1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdTATEVNT1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdTATEVNT1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdTATEVNT1.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdTATEVNT1.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance42.BackColor = System.Drawing.Color.Transparent
        Me.grdTATEVNT1.DisplayLayout.Override.CardAreaAppearance = Appearance42
        Appearance43.BorderColor = System.Drawing.Color.Silver
        Appearance43.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdTATEVNT1.DisplayLayout.Override.CellAppearance = Appearance43
        Me.grdTATEVNT1.DisplayLayout.Override.CellPadding = 0
        Appearance44.BackColor = System.Drawing.SystemColors.Control
        Appearance44.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance44.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance44.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance44.BorderColor = System.Drawing.SystemColors.Window
        Me.grdTATEVNT1.DisplayLayout.Override.GroupByRowAppearance = Appearance44
        Appearance46.TextHAlignAsString = "Left"
        Me.grdTATEVNT1.DisplayLayout.Override.HeaderAppearance = Appearance46
        Me.grdTATEVNT1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdTATEVNT1.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance47.BackColor = System.Drawing.SystemColors.Window
        Appearance47.BorderColor = System.Drawing.Color.Silver
        Me.grdTATEVNT1.DisplayLayout.Override.RowAppearance = Appearance47
        Me.grdTATEVNT1.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance48.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdTATEVNT1.DisplayLayout.Override.TemplateAddRowAppearance = Appearance48
        Me.grdTATEVNT1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdTATEVNT1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdTATEVNT1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdTATEVNT1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdTATEVNT1.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdTATEVNT1.Location = New System.Drawing.Point(0, 0)
        Me.grdTATEVNT1.Name = "grdTATEVNT1"
        Me.grdTATEVNT1.Size = New System.Drawing.Size(1030, 280)
        Me.grdTATEVNT1.TabIndex = 18
        Me.grdTATEVNT1.Text = "Events"
        '
        'TAFEVNT1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1030, 327)
        Me.ControlBox = False
        Me.Name = "TAFEVNT1"
        Me.Text = "Event Viewer"
        Me.ASFBASE2_Fill_Panel.ResumeLayout(False)
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.grdTATEVNT1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents cmdCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdOK As Infragistics.Win.Misc.UltraButton
    Friend WithEvents grdTATEVNT1 As Infragistics.Win.UltraWinGrid.UltraGrid
End Class
