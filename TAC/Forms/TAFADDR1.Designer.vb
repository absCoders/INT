<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TAFADDR1
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
        Dim Appearance31 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1)
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("Consignee")
        Dim Appearance32 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance33 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim UltraGridColumn2 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("BuildingName")
        Dim Appearance34 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance35 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("AddressLine1")
        Dim Appearance36 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance37 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("AddressLine2")
        Dim Appearance38 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance39 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim UltraGridColumn5 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("AddressLine3")
        Dim Appearance40 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance41 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim UltraGridColumn6 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("City")
        Dim Appearance42 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance43 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim UltraGridColumn7 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("State")
        Dim Appearance44 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance45 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim UltraGridColumn8 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("PostalCode")
        Dim Appearance46 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance47 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim UltraGridColumn9 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CountryCode")
        Dim Appearance48 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance49 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance50 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance51 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance52 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance53 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance54 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance55 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance56 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance57 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance58 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance59 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance60 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.cmdCancel = New Infragistics.Win.Misc.UltraButton
        Me.cmdSelect = New Infragistics.Win.Misc.UltraButton
        Me.grdADDRESS = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.ASFBASE2_Fill_Panel.SuspendLayout()
        CType(Me.grdADDRESS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.grdADDRESS)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.cmdCancel)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.cmdSelect)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(537, 469)
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(453, 425)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(71, 31)
        Me.cmdCancel.TabIndex = 6
        Me.cmdCancel.Text = "Cancel"
        '
        'cmdSelect
        '
        Me.cmdSelect.Location = New System.Drawing.Point(374, 425)
        Me.cmdSelect.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdSelect.Name = "cmdSelect"
        Me.cmdSelect.Size = New System.Drawing.Size(71, 31)
        Me.cmdSelect.TabIndex = 5
        Me.cmdSelect.Text = "Select"
        '
        'grdADDRESS
        '
        Appearance31.BackColor = System.Drawing.SystemColors.Window
        Appearance31.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdADDRESS.DisplayLayout.Appearance = Appearance31
        Appearance32.TextHAlignAsString = "Left"
        UltraGridColumn1.CellAppearance = Appearance32
        Appearance33.ImageHAlign = Infragistics.Win.HAlign.Left
        Appearance33.TextHAlignAsString = "Left"
        UltraGridColumn1.Header.Appearance = Appearance33
        UltraGridColumn1.Header.VisiblePosition = 0
        UltraGridColumn1.Hidden = True
        UltraGridColumn1.RowLayoutColumnInfo.OriginX = 0
        UltraGridColumn1.RowLayoutColumnInfo.OriginY = 0
        UltraGridColumn1.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(157, 0)
        UltraGridColumn1.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn1.RowLayoutColumnInfo.SpanY = 2
        Appearance34.TextHAlignAsString = "Left"
        UltraGridColumn2.CellAppearance = Appearance34
        Appearance35.ImageHAlign = Infragistics.Win.HAlign.Left
        Appearance35.TextHAlignAsString = "Left"
        UltraGridColumn2.Header.Appearance = Appearance35
        UltraGridColumn2.Header.Caption = "Building"
        UltraGridColumn2.Header.VisiblePosition = 1
        UltraGridColumn2.Hidden = True
        UltraGridColumn2.RowLayoutColumnInfo.OriginX = 2
        UltraGridColumn2.RowLayoutColumnInfo.OriginY = 0
        UltraGridColumn2.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn2.RowLayoutColumnInfo.SpanY = 2
        Appearance36.TextHAlignAsString = "Left"
        UltraGridColumn3.CellAppearance = Appearance36
        Appearance37.ImageHAlign = Infragistics.Win.HAlign.Left
        Appearance37.TextHAlignAsString = "Left"
        UltraGridColumn3.Header.Appearance = Appearance37
        UltraGridColumn3.Header.Caption = "Address Line 1"
        UltraGridColumn3.Header.VisiblePosition = 2
        UltraGridColumn3.Hidden = True
        UltraGridColumn3.RowLayoutColumnInfo.OriginX = 6
        UltraGridColumn3.RowLayoutColumnInfo.OriginY = 0
        UltraGridColumn3.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(201, 0)
        UltraGridColumn3.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn3.RowLayoutColumnInfo.SpanY = 2
        Appearance38.TextHAlignAsString = "Left"
        UltraGridColumn4.CellAppearance = Appearance38
        Appearance39.ImageHAlign = Infragistics.Win.HAlign.Left
        Appearance39.TextHAlignAsString = "Left"
        UltraGridColumn4.Header.Appearance = Appearance39
        UltraGridColumn4.Header.Caption = "Address Line 2"
        UltraGridColumn4.Header.VisiblePosition = 3
        UltraGridColumn4.Hidden = True
        UltraGridColumn4.RowLayoutColumnInfo.OriginX = 8
        UltraGridColumn4.RowLayoutColumnInfo.OriginY = 0
        UltraGridColumn4.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(165, 0)
        UltraGridColumn4.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn4.RowLayoutColumnInfo.SpanY = 2
        Appearance40.TextHAlignAsString = "Left"
        UltraGridColumn5.CellAppearance = Appearance40
        Appearance41.ImageHAlign = Infragistics.Win.HAlign.Left
        Appearance41.TextHAlignAsString = "Left"
        UltraGridColumn5.Header.Appearance = Appearance41
        UltraGridColumn5.Header.Caption = "Address Line 3"
        UltraGridColumn5.Header.VisiblePosition = 4
        UltraGridColumn5.Hidden = True
        UltraGridColumn5.RowLayoutColumnInfo.OriginX = 10
        UltraGridColumn5.RowLayoutColumnInfo.OriginY = 0
        UltraGridColumn5.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(175, 0)
        UltraGridColumn5.RowLayoutColumnInfo.SpanX = 1
        UltraGridColumn5.RowLayoutColumnInfo.SpanY = 2
        Appearance42.TextHAlignAsString = "Left"
        UltraGridColumn6.CellAppearance = Appearance42
        Appearance43.ImageHAlign = Infragistics.Win.HAlign.Left
        Appearance43.TextHAlignAsString = "Left"
        UltraGridColumn6.Header.Appearance = Appearance43
        UltraGridColumn6.Header.VisiblePosition = 5
        UltraGridColumn6.RowLayoutColumnInfo.OriginX = 2
        UltraGridColumn6.RowLayoutColumnInfo.OriginY = 0
        UltraGridColumn6.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(250, 0)
        UltraGridColumn6.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn6.RowLayoutColumnInfo.SpanY = 2
        Appearance44.TextHAlignAsString = "Left"
        UltraGridColumn7.CellAppearance = Appearance44
        Appearance45.ImageHAlign = Infragistics.Win.HAlign.Left
        Appearance45.TextHAlignAsString = "Left"
        UltraGridColumn7.Header.Appearance = Appearance45
        UltraGridColumn7.Header.VisiblePosition = 6
        UltraGridColumn7.RowLayoutColumnInfo.OriginX = 0
        UltraGridColumn7.RowLayoutColumnInfo.OriginY = 0
        UltraGridColumn7.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(55, 0)
        UltraGridColumn7.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn7.RowLayoutColumnInfo.SpanY = 2
        Appearance46.TextHAlignAsString = "Left"
        UltraGridColumn8.CellAppearance = Appearance46
        Appearance47.ImageHAlign = Infragistics.Win.HAlign.Left
        Appearance47.TextHAlignAsString = "Left"
        UltraGridColumn8.Header.Appearance = Appearance47
        UltraGridColumn8.Header.Caption = "Postal Code"
        UltraGridColumn8.Header.VisiblePosition = 7
        UltraGridColumn8.RowLayoutColumnInfo.OriginX = 4
        UltraGridColumn8.RowLayoutColumnInfo.OriginY = 0
        UltraGridColumn8.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(100, 0)
        UltraGridColumn8.RowLayoutColumnInfo.SpanX = 2
        UltraGridColumn8.RowLayoutColumnInfo.SpanY = 2
        Appearance48.TextHAlignAsString = "Left"
        UltraGridColumn9.CellAppearance = Appearance48
        Appearance49.ImageHAlign = Infragistics.Win.HAlign.Left
        Appearance49.TextHAlignAsString = "Left"
        UltraGridColumn9.Header.Appearance = Appearance49
        UltraGridColumn9.Header.Caption = "Country"
        UltraGridColumn9.Header.VisiblePosition = 8
        UltraGridColumn9.HiddenWhenGroupBy = Infragistics.Win.DefaultableBoolean.[True]
        UltraGridColumn9.RowLayoutColumnInfo.OriginX = 11
        UltraGridColumn9.RowLayoutColumnInfo.OriginY = 0
        UltraGridColumn9.RowLayoutColumnInfo.PreferredCellSize = New System.Drawing.Size(65, 0)
        UltraGridColumn9.RowLayoutColumnInfo.SpanX = 1
        UltraGridColumn9.RowLayoutColumnInfo.SpanY = 2
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn1, UltraGridColumn2, UltraGridColumn3, UltraGridColumn4, UltraGridColumn5, UltraGridColumn6, UltraGridColumn7, UltraGridColumn8, UltraGridColumn9})
        UltraGridBand1.UseRowLayout = True
        Me.grdADDRESS.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdADDRESS.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance50.TextHAlignAsString = "Left"
        Me.grdADDRESS.DisplayLayout.CaptionAppearance = Appearance50
        Appearance51.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance51.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance51.BorderColor = System.Drawing.SystemColors.Window
        Me.grdADDRESS.DisplayLayout.GroupByBox.Appearance = Appearance51
        Appearance52.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdADDRESS.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance52
        Me.grdADDRESS.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdADDRESS.DisplayLayout.GroupByBox.Hidden = True
        Appearance53.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance53.BackColor2 = System.Drawing.SystemColors.Control
        Appearance53.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance53.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdADDRESS.DisplayLayout.GroupByBox.PromptAppearance = Appearance53
        Me.grdADDRESS.DisplayLayout.MaxColScrollRegions = 1
        Me.grdADDRESS.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdADDRESS.DisplayLayout.NewBandLoadStyle = Infragistics.Win.UltraWinGrid.NewBandLoadStyle.Hide
        Me.grdADDRESS.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance54.BackColor = System.Drawing.SystemColors.Window
        Appearance54.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdADDRESS.DisplayLayout.Override.ActiveCellAppearance = Appearance54
        Me.grdADDRESS.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdADDRESS.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdADDRESS.DisplayLayout.Override.AllowGroupBy = Infragistics.Win.DefaultableBoolean.[True]
        Me.grdADDRESS.DisplayLayout.Override.AllowGroupMoving = Infragistics.Win.UltraWinGrid.AllowGroupMoving.NotAllowed
        Me.grdADDRESS.DisplayLayout.Override.AllowGroupSwapping = Infragistics.Win.UltraWinGrid.AllowGroupSwapping.NotAllowed
        Me.grdADDRESS.DisplayLayout.Override.AllowMultiCellOperations = Infragistics.Win.UltraWinGrid.AllowMultiCellOperation.None
        Me.grdADDRESS.DisplayLayout.Override.AllowRowLayoutCellSizing = Infragistics.Win.UltraWinGrid.RowLayoutSizing.None
        Me.grdADDRESS.DisplayLayout.Override.AllowRowLayoutCellSpanSizing = Infragistics.Win.Layout.GridBagLayoutAllowSpanSizing.None
        Me.grdADDRESS.DisplayLayout.Override.AllowRowLayoutColMoving = Infragistics.Win.Layout.GridBagLayoutAllowMoving.None
        Me.grdADDRESS.DisplayLayout.Override.AllowRowLayoutLabelSpanSizing = Infragistics.Win.Layout.GridBagLayoutAllowSpanSizing.None
        Me.grdADDRESS.DisplayLayout.Override.AllowRowSummaries = Infragistics.Win.UltraWinGrid.AllowRowSummaries.[False]
        Me.grdADDRESS.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdADDRESS.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdADDRESS.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance55.BackColor = System.Drawing.SystemColors.Window
        Me.grdADDRESS.DisplayLayout.Override.CardAreaAppearance = Appearance55
        Appearance56.BorderColor = System.Drawing.Color.Silver
        Appearance56.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdADDRESS.DisplayLayout.Override.CellAppearance = Appearance56
        Me.grdADDRESS.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        Me.grdADDRESS.DisplayLayout.Override.CellPadding = 0
        Appearance57.BackColor = System.Drawing.SystemColors.Control
        Appearance57.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance57.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance57.BorderColor = System.Drawing.SystemColors.Window
        Me.grdADDRESS.DisplayLayout.Override.GroupByRowAppearance = Appearance57
        Appearance58.TextHAlignAsString = "Center"
        Me.grdADDRESS.DisplayLayout.Override.HeaderAppearance = Appearance58
        Me.grdADDRESS.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.[Select]
        Me.grdADDRESS.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance59.BackColor = System.Drawing.SystemColors.Window
        Appearance59.BorderColor = System.Drawing.Color.Silver
        Me.grdADDRESS.DisplayLayout.Override.RowAppearance = Appearance59
        Me.grdADDRESS.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Me.grdADDRESS.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[True]
        Me.grdADDRESS.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.[Single]
        Appearance60.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdADDRESS.DisplayLayout.Override.TemplateAddRowAppearance = Appearance60
        Me.grdADDRESS.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdADDRESS.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdADDRESS.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdADDRESS.Location = New System.Drawing.Point(12, 12)
        Me.grdADDRESS.Name = "grdADDRESS"
        Me.grdADDRESS.Size = New System.Drawing.Size(513, 406)
        Me.grdADDRESS.TabIndex = 23
        Me.grdADDRESS.TabStop = False
        Me.grdADDRESS.Text = "Best Match Addresses"
        '
        'TAFADDR1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(537, 469)
        Me.ControlBox = False
        Me.Name = "TAFADDR1"
        Me.Text = "Address Validation"
        Me.ASFBASE2_Fill_Panel.ResumeLayout(False)
        CType(Me.grdADDRESS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdSelect As Infragistics.Win.Misc.UltraButton
    Friend WithEvents grdADDRESS As Infragistics.Win.UltraWinGrid.UltraGrid
End Class
