<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CreditCardSubmission
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Dim Appearance195 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("ARTCCPA1", -1)
        Dim UltraGridColumn341 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CCPA_NO")
        Dim UltraGridColumn342 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CODE")
        Dim UltraGridColumn343 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CCPA_STATUS")
        Dim UltraGridColumn344 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CCPA_REASON")
        Dim UltraGridColumn345 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CCPA_NOTE")
        Dim UltraGridColumn346 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CCPA_AMT")
        Dim UltraGridColumn347 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CCPA_DATE_AUTH")
        Dim UltraGridColumn348 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CCPA_DATE_SALE")
        Dim UltraGridColumn349 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CCPA_AUTH")
        Dim UltraGridColumn350 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_OPER")
        Dim UltraGridColumn351 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_DATE")
        Dim UltraGridColumn352 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_OPER")
        Dim UltraGridColumn353 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_DATE")
        Dim UltraGridColumn354 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_NO")
        Dim UltraGridColumn355 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_EXP_DATE")
        Dim UltraGridColumn356 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_VER_CODE")
        Dim UltraGridColumn357 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_NAME")
        Dim UltraGridColumn358 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_ADDR1")
        Dim UltraGridColumn359 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_CITY")
        Dim UltraGridColumn360 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_STATE")
        Dim UltraGridColumn361 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_ZIP_CODE")
        Dim UltraGridColumn362 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_LAST4")
        Dim UltraGridColumn363 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("RESPONSE_RETRIEVAL_NO")
        Dim UltraGridColumn364 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("RESPONSE_CODE")
        Dim UltraGridColumn365 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("RESPONSE_BATCH_NO")
        Dim UltraGridColumn366 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("RESPONSE_APPROVAL_CODE")
        Dim UltraGridColumn367 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("RESPONSE_TEXT")
        Dim UltraGridColumn368 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CCPA_TYPE")
        Dim UltraGridColumn369 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ORDR_NO")
        Dim UltraGridColumn370 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INV_NO")
        Dim UltraGridColumn371 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LENS_BANK_INV_NO")
        Dim UltraGridColumn372 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_NAME")
        Dim Appearance196 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance197 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance198 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance199 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance200 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance201 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance202 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance203 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance204 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance205 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance206 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.grdControl = New Infragistics.Win.UltraWinGrid.UltraGrid
        CType(Me.grdControl, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grdControl
        '
        Appearance195.BackColor = System.Drawing.SystemColors.Window
        Appearance195.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdControl.DisplayLayout.Appearance = Appearance195
        UltraGridColumn341.Header.Caption = "Auth No"
        UltraGridColumn341.Header.VisiblePosition = 0
        UltraGridColumn341.Hidden = True
        UltraGridColumn341.Width = 95
        UltraGridColumn342.Header.Caption = "Customer"
        UltraGridColumn342.Header.VisiblePosition = 1
        UltraGridColumn342.Hidden = True
        UltraGridColumn342.Width = 92
        UltraGridColumn343.Header.Caption = "Status"
        UltraGridColumn343.Header.VisiblePosition = 6
        UltraGridColumn343.Width = 103
        UltraGridColumn344.Header.Caption = "Reason"
        UltraGridColumn344.Header.VisiblePosition = 8
        UltraGridColumn344.Width = 79
        UltraGridColumn345.Header.Caption = "Note"
        UltraGridColumn345.Header.VisiblePosition = 21
        UltraGridColumn345.Width = 169
        UltraGridColumn346.Header.Caption = "Amount"
        UltraGridColumn346.Header.VisiblePosition = 7
        UltraGridColumn346.Width = 90
        UltraGridColumn347.Format = "MM/dd/yyyy hh:mm tt"
        UltraGridColumn347.Header.Caption = "Submitted"
        UltraGridColumn347.Header.VisiblePosition = 12
        UltraGridColumn347.Hidden = True
        UltraGridColumn347.Width = 184
        UltraGridColumn348.Format = "MM/dd/yyyy hh:mm tt"
        UltraGridColumn348.Header.Caption = "Authorized"
        UltraGridColumn348.Header.VisiblePosition = 13
        UltraGridColumn348.Hidden = True
        UltraGridColumn348.Width = 165
        UltraGridColumn349.Header.VisiblePosition = 14
        UltraGridColumn349.Hidden = True
        UltraGridColumn350.Header.Caption = "Init Oper"
        UltraGridColumn350.Header.VisiblePosition = 16
        UltraGridColumn351.Header.Caption = "Init Date"
        UltraGridColumn351.Header.VisiblePosition = 15
        UltraGridColumn352.Header.Caption = "Last Oper"
        UltraGridColumn352.Header.VisiblePosition = 17
        UltraGridColumn352.Width = 100
        UltraGridColumn353.Format = "MM/dd/yyyy hh:mm tt"
        UltraGridColumn353.Header.Caption = "Date/Time"
        UltraGridColumn353.Header.VisiblePosition = 5
        UltraGridColumn353.Width = 189
        UltraGridColumn354.Header.Caption = "Credit Card No"
        UltraGridColumn354.Header.VisiblePosition = 18
        UltraGridColumn354.Hidden = True
        UltraGridColumn355.Header.Caption = "Exp"
        UltraGridColumn355.Header.VisiblePosition = 4
        UltraGridColumn355.Width = 52
        UltraGridColumn356.Header.Caption = "Ver"
        UltraGridColumn356.Header.VisiblePosition = 19
        UltraGridColumn356.Hidden = True
        UltraGridColumn356.Width = 63
        UltraGridColumn357.Header.Caption = "Name"
        UltraGridColumn357.Header.VisiblePosition = 20
        UltraGridColumn358.Header.Caption = "Address"
        UltraGridColumn358.Header.VisiblePosition = 22
        UltraGridColumn359.Header.Caption = "City"
        UltraGridColumn359.Header.VisiblePosition = 23
        UltraGridColumn360.Header.Caption = "State"
        UltraGridColumn360.Header.VisiblePosition = 24
        UltraGridColumn360.Width = 63
        UltraGridColumn361.Header.Caption = "Zip Code"
        UltraGridColumn361.Header.VisiblePosition = 27
        UltraGridColumn361.Width = 101
        UltraGridColumn362.Header.Caption = "Last4"
        UltraGridColumn362.Header.VisiblePosition = 3
        UltraGridColumn362.Width = 60
        UltraGridColumn363.Header.VisiblePosition = 25
        UltraGridColumn363.Hidden = True
        UltraGridColumn364.Header.VisiblePosition = 26
        UltraGridColumn364.Hidden = True
        UltraGridColumn365.Header.VisiblePosition = 28
        UltraGridColumn365.Hidden = True
        UltraGridColumn366.Header.VisiblePosition = 29
        UltraGridColumn366.Hidden = True
        UltraGridColumn367.Header.Caption = "Status Text"
        UltraGridColumn367.Header.VisiblePosition = 11
        UltraGridColumn367.Width = 147
        UltraGridColumn368.Header.VisiblePosition = 30
        UltraGridColumn368.Hidden = True
        UltraGridColumn369.Header.Caption = "Order No"
        UltraGridColumn369.Header.VisiblePosition = 9
        UltraGridColumn370.Header.Caption = "Invoice No"
        UltraGridColumn370.Header.VisiblePosition = 10
        UltraGridColumn371.Header.VisiblePosition = 31
        UltraGridColumn371.Hidden = True
        UltraGridColumn372.Header.Caption = "Name"
        UltraGridColumn372.Header.VisiblePosition = 2
        UltraGridColumn372.Hidden = True
        UltraGridColumn372.Width = 160
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn341, UltraGridColumn342, UltraGridColumn343, UltraGridColumn344, UltraGridColumn345, UltraGridColumn346, UltraGridColumn347, UltraGridColumn348, UltraGridColumn349, UltraGridColumn350, UltraGridColumn351, UltraGridColumn352, UltraGridColumn353, UltraGridColumn354, UltraGridColumn355, UltraGridColumn356, UltraGridColumn357, UltraGridColumn358, UltraGridColumn359, UltraGridColumn360, UltraGridColumn361, UltraGridColumn362, UltraGridColumn363, UltraGridColumn364, UltraGridColumn365, UltraGridColumn366, UltraGridColumn367, UltraGridColumn368, UltraGridColumn369, UltraGridColumn370, UltraGridColumn371, UltraGridColumn372})
        Me.grdControl.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdControl.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance196.TextHAlignAsString = "Left"
        Me.grdControl.DisplayLayout.CaptionAppearance = Appearance196
        Appearance197.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance197.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance197.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance197.BorderColor = System.Drawing.SystemColors.Window
        Me.grdControl.DisplayLayout.GroupByBox.Appearance = Appearance197
        Appearance198.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdControl.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance198
        Me.grdControl.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdControl.DisplayLayout.GroupByBox.Hidden = True
        Appearance199.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance199.BackColor2 = System.Drawing.SystemColors.Control
        Appearance199.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance199.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdControl.DisplayLayout.GroupByBox.PromptAppearance = Appearance199
        Me.grdControl.DisplayLayout.MaxColScrollRegions = 1
        Me.grdControl.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdControl.DisplayLayout.NewBandLoadStyle = Infragistics.Win.UltraWinGrid.NewBandLoadStyle.Hide
        Me.grdControl.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance200.BackColor = System.Drawing.SystemColors.Window
        Appearance200.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdControl.DisplayLayout.Override.ActiveCellAppearance = Appearance200
        Me.grdControl.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdControl.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdControl.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdControl.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdControl.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance201.BackColor = System.Drawing.SystemColors.Window
        Me.grdControl.DisplayLayout.Override.CardAreaAppearance = Appearance201
        Appearance202.BorderColor = System.Drawing.Color.Silver
        Appearance202.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdControl.DisplayLayout.Override.CellAppearance = Appearance202
        Me.grdControl.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        Me.grdControl.DisplayLayout.Override.CellPadding = 0
        Appearance203.BackColor = System.Drawing.SystemColors.Control
        Appearance203.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance203.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance203.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance203.BorderColor = System.Drawing.SystemColors.Window
        Me.grdControl.DisplayLayout.Override.GroupByRowAppearance = Appearance203
        Appearance204.TextHAlignAsString = "Left"
        Me.grdControl.DisplayLayout.Override.HeaderAppearance = Appearance204
        Me.grdControl.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdControl.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance205.BackColor = System.Drawing.SystemColors.Window
        Appearance205.BorderColor = System.Drawing.Color.Silver
        Me.grdControl.DisplayLayout.Override.RowAppearance = Appearance205
        Me.grdControl.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance206.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdControl.DisplayLayout.Override.TemplateAddRowAppearance = Appearance206
        Me.grdControl.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdControl.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.grdControl.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdControl.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdControl.Location = New System.Drawing.Point(0, 0)
        Me.grdControl.Name = "grdControl"
        Me.grdControl.Size = New System.Drawing.Size(922, 289)
        Me.grdControl.TabIndex = 4
        Me.grdControl.Text = "Credit Card Submission History"
        '
        'CreditCardSubmission
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.grdControl)
        Me.Name = "CreditCardSubmission"
        Me.Size = New System.Drawing.Size(922, 289)
        CType(Me.grdControl, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grdControl As Infragistics.Win.UltraWinGrid.UltraGrid

End Class
