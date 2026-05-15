<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CreditCardQueue
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
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand1 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("ARTCUSTC", -1)
        Dim UltraGridColumn98 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CODE")
        Dim UltraGridColumn99 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_NO")
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridColumn100 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_EXP_DATE")
        Dim UltraGridColumn101 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_VER_CODE")
        Dim UltraGridColumn102 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_NAME", -1, Nothing, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, False)
        Dim UltraGridColumn103 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_ADDR1")
        Dim UltraGridColumn104 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_CITY")
        Dim UltraGridColumn105 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_STATE")
        Dim UltraGridColumn106 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_ZIP_CODE")
        Dim UltraGridColumn107 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_STATUS", -1, 411696172)
        Dim UltraGridColumn108 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_OPER")
        Dim UltraGridColumn109 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_DATE")
        Dim UltraGridColumn110 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_OPER")
        Dim UltraGridColumn111 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_DATE")
        Dim UltraGridColumn112 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_PREFERRED")
        Dim UltraGridColumn113 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_LAST4")
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridColumn1 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_COUNTRY")
        Dim UltraGridColumn114 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CREDIT_CARD_TYPE", 0)
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
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
        Dim ValueList1 As Infragistics.Win.ValueList = New Infragistics.Win.ValueList(411696172)
        Dim ValueListItem3 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem8 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim Appearance16 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim UltraGridBand2 As Infragistics.Win.UltraWinGrid.UltraGridBand = New Infragistics.Win.UltraWinGrid.UltraGridBand("ARTCUSTC", -1)
        Dim UltraGridColumn2 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("CUST_CODE")
        Dim UltraGridColumn3 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_OPER")
        Dim UltraGridColumn4 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("INIT_DATE")
        Dim UltraGridColumn5 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_OPER")
        Dim UltraGridColumn6 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("LAST_DATE")
        Dim UltraGridColumn7 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ACH_ACCT_ID")
        Dim UltraGridColumn8 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ACH_ACCT_NAME")
        Dim UltraGridColumn9 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ACH_AUTO_PAY_IND")
        Dim UltraGridColumn10 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ACH_DEFAULT_ACCT_IND")
        Dim UltraGridColumn11 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ACH_ACCT_TYPE_ID")
        Dim UltraGridColumn12 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ACH_ROUTING_NO")
        Dim UltraGridColumn13 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ACH_ACCT_NO")
        Dim UltraGridColumn14 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ACH_ACCT_STATUS")
        Dim UltraGridColumn15 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ACH_BANK_NAME")
        Dim UltraGridColumn16 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("WEB_IND")
        Dim UltraGridColumn17 As Infragistics.Win.UltraWinGrid.UltraGridColumn = New Infragistics.Win.UltraWinGrid.UltraGridColumn("ACH_ACCT_NO_LAST4")
        Dim Appearance17 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance18 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance19 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance20 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance21 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance22 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance23 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance24 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance25 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance26 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance27 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance28 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim ValueList2 As Infragistics.Win.ValueList = New Infragistics.Win.ValueList(411696172)
        Dim ValueListItem1 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem2 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim UltraTab2 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim UltraTab1 As Infragistics.Win.UltraWinTabControl.UltraTab = New Infragistics.Win.UltraWinTabControl.UltraTab()
        Dim Appearance29 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim ValueListItem12 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem13 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem10 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem11 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim Appearance30 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim ValueListItem4 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem5 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem6 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Dim ValueListItem7 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem()
        Me.UltraTabPageControl2 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.grdCC = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.UltraTabPageControl1 = New Infragistics.Win.UltraWinTabControl.UltraTabPageControl()
        Me.grdACH = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.grpHeader = New Infragistics.Win.Misc.UltraGroupBox()
        Me.chkCCEdit = New ABSCS.ABSCheckBox()
        Me.SplitContainer24 = New System.Windows.Forms.SplitContainer()
        Me.tabgrids = New Infragistics.Win.UltraWinTabControl.UltraTabControl()
        Me.UltraTabSharedControlsPage1 = New Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage()
        Me.SplitContainer25 = New System.Windows.Forms.SplitContainer()
        Me.grpCCPA = New Infragistics.Win.Misc.UltraGroupBox()
        Me.chkCUST_AUTOQ_WEB = New ABSCS.ABSCheckBox()
        Me.cbeCUST_AUTO_CC_OPER = New Infragistics.Win.UltraWinEditors.UltraComboEditor()
        Me.dteCUST_AUTO_CC_AUTH_DATE = New Infragistics.Win.UltraWinEditors.UltraDateTimeEditor()
        Me.chkCUST_AUTO_CC_AUTH = New ABSCS.ABSCheckBox()
        Me.optCUST_AUTO_CCPA = New Infragistics.Win.UltraWinEditors.UltraOptionSet()
        Me.txtCUST_AUTO_CCPA_NOTE = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel51 = New Infragistics.Win.Misc.UltraLabel()
        Me.grpButtons = New Infragistics.Win.Misc.UltraGroupBox()
        Me.cmdCCUpdate = New Infragistics.Win.Misc.UltraButton()
        Me.cmdCCCancel = New Infragistics.Win.Misc.UltraButton()
        Me.UltraTabPageControl2.SuspendLayout()
        CType(Me.grdCC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraTabPageControl1.SuspendLayout()
        CType(Me.grdACH, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpHeader, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpHeader.SuspendLayout()
        CType(Me.chkCCEdit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer24, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer24.Panel1.SuspendLayout()
        Me.SplitContainer24.Panel2.SuspendLayout()
        Me.SplitContainer24.SuspendLayout()
        CType(Me.tabgrids, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabgrids.SuspendLayout()
        CType(Me.SplitContainer25, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer25.Panel1.SuspendLayout()
        Me.SplitContainer25.Panel2.SuspendLayout()
        Me.SplitContainer25.SuspendLayout()
        CType(Me.grpCCPA, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpCCPA.SuspendLayout()
        CType(Me.chkCUST_AUTOQ_WEB, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cbeCUST_AUTO_CC_OPER, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dteCUST_AUTO_CC_AUTH_DATE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkCUST_AUTO_CC_AUTH, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optCUST_AUTO_CCPA, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCUST_AUTO_CCPA_NOTE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpButtons, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpButtons.SuspendLayout()
        Me.SuspendLayout()
        '
        'UltraTabPageControl2
        '
        Me.UltraTabPageControl2.Controls.Add(Me.grdCC)
        Me.UltraTabPageControl2.Location = New System.Drawing.Point(1, 23)
        Me.UltraTabPageControl2.Name = "UltraTabPageControl2"
        Me.UltraTabPageControl2.Size = New System.Drawing.Size(605, 159)
        '
        'grdCC
        '
        Appearance1.BackColor = System.Drawing.SystemColors.Window
        Appearance1.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdCC.DisplayLayout.Appearance = Appearance1
        UltraGridColumn98.Header.VisiblePosition = 0
        UltraGridColumn98.Hidden = True
        Appearance2.BackColor = System.Drawing.Color.Beige
        UltraGridColumn99.CellAppearance = Appearance2
        UltraGridColumn99.Header.Caption = "Credit Card No"
        UltraGridColumn99.Header.VisiblePosition = 1
        UltraGridColumn99.MaskInput = ""
        UltraGridColumn99.Width = 151
        UltraGridColumn100.Header.Caption = "Exp MMYY"
        UltraGridColumn100.Header.VisiblePosition = 4
        UltraGridColumn100.Width = 85
        UltraGridColumn101.Header.Caption = "CVV2"
        UltraGridColumn101.Header.VisiblePosition = 5
        UltraGridColumn101.Width = 58
        UltraGridColumn102.Header.Caption = "Name on Card"
        UltraGridColumn102.Header.VisiblePosition = 8
        UltraGridColumn103.Header.Caption = "Street Address"
        UltraGridColumn103.Header.VisiblePosition = 9
        UltraGridColumn103.Width = 115
        UltraGridColumn104.Header.Caption = "City"
        UltraGridColumn104.Header.VisiblePosition = 10
        UltraGridColumn104.Width = 106
        UltraGridColumn105.Header.Caption = "State"
        UltraGridColumn105.Header.VisiblePosition = 11
        UltraGridColumn105.Width = 51
        UltraGridColumn106.Header.Caption = "Zip Code"
        UltraGridColumn106.Header.VisiblePosition = 12
        UltraGridColumn106.Width = 73
        UltraGridColumn107.Header.Caption = "Status"
        UltraGridColumn107.Header.VisiblePosition = 7
        UltraGridColumn107.Width = 72
        UltraGridColumn108.Header.VisiblePosition = 13
        UltraGridColumn108.Hidden = True
        UltraGridColumn109.Header.VisiblePosition = 14
        UltraGridColumn109.Hidden = True
        UltraGridColumn110.Header.VisiblePosition = 15
        UltraGridColumn110.Hidden = True
        UltraGridColumn111.Header.VisiblePosition = 16
        UltraGridColumn111.Hidden = True
        UltraGridColumn112.Header.Caption = "Pref"
        UltraGridColumn112.Header.VisiblePosition = 6
        UltraGridColumn112.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn112.Width = 43
        UltraGridColumn113.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        Appearance3.BackColor = System.Drawing.Color.Beige
        UltraGridColumn113.CellAppearance = Appearance3
        UltraGridColumn113.Header.Caption = "Last 4"
        UltraGridColumn113.Header.VisiblePosition = 3
        UltraGridColumn113.Width = 63
        UltraGridColumn1.Header.Caption = "Country Code"
        UltraGridColumn1.Header.VisiblePosition = 17
        UltraGridColumn1.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.EditButton
        UltraGridColumn1.Width = 250
        Appearance4.BackColor = System.Drawing.Color.Beige
        UltraGridColumn114.CellAppearance = Appearance4
        UltraGridColumn114.DataType = GetType(System.Drawing.Bitmap)
        UltraGridColumn114.Header.Caption = "Type"
        UltraGridColumn114.Header.VisiblePosition = 2
        UltraGridColumn114.Width = 88
        UltraGridBand1.Columns.AddRange(New Object() {UltraGridColumn98, UltraGridColumn99, UltraGridColumn100, UltraGridColumn101, UltraGridColumn102, UltraGridColumn103, UltraGridColumn104, UltraGridColumn105, UltraGridColumn106, UltraGridColumn107, UltraGridColumn108, UltraGridColumn109, UltraGridColumn110, UltraGridColumn111, UltraGridColumn112, UltraGridColumn113, UltraGridColumn1, UltraGridColumn114})
        Me.grdCC.DisplayLayout.BandsSerializer.Add(UltraGridBand1)
        Me.grdCC.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance5.TextHAlignAsString = "Left"
        Me.grdCC.DisplayLayout.CaptionAppearance = Appearance5
        Appearance6.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance6.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance6.BorderColor = System.Drawing.SystemColors.Window
        Me.grdCC.DisplayLayout.GroupByBox.Appearance = Appearance6
        Appearance7.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdCC.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance7
        Me.grdCC.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdCC.DisplayLayout.GroupByBox.Hidden = True
        Appearance8.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance8.BackColor2 = System.Drawing.SystemColors.Control
        Appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance8.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdCC.DisplayLayout.GroupByBox.PromptAppearance = Appearance8
        Me.grdCC.DisplayLayout.MaxColScrollRegions = 1
        Me.grdCC.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdCC.DisplayLayout.NewBandLoadStyle = Infragistics.Win.UltraWinGrid.NewBandLoadStyle.Hide
        Me.grdCC.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance9.BackColor = System.Drawing.SystemColors.Window
        Appearance9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdCC.DisplayLayout.Override.ActiveCellAppearance = Appearance9
        Me.grdCC.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.FixedAddRowOnTop
        Me.grdCC.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdCC.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[True]
        Me.grdCC.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdCC.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance10.BackColor = System.Drawing.SystemColors.Window
        Me.grdCC.DisplayLayout.Override.CardAreaAppearance = Appearance10
        Appearance11.BorderColor = System.Drawing.Color.Silver
        Appearance11.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdCC.DisplayLayout.Override.CellAppearance = Appearance11
        Me.grdCC.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.grdCC.DisplayLayout.Override.CellPadding = 0
        Appearance12.BackColor = System.Drawing.SystemColors.Control
        Appearance12.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance12.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance12.BorderColor = System.Drawing.SystemColors.Window
        Me.grdCC.DisplayLayout.Override.GroupByRowAppearance = Appearance12
        Appearance13.TextHAlignAsString = "Left"
        Me.grdCC.DisplayLayout.Override.HeaderAppearance = Appearance13
        Me.grdCC.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdCC.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance14.BackColor = System.Drawing.SystemColors.Window
        Appearance14.BorderColor = System.Drawing.Color.Silver
        Me.grdCC.DisplayLayout.Override.RowAppearance = Appearance14
        Me.grdCC.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance15.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdCC.DisplayLayout.Override.TemplateAddRowAppearance = Appearance15
        Me.grdCC.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdCC.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        ValueList1.Key = "CUST_CREDIT_CARD_STATUS"
        ValueListItem3.DataValue = "A"
        ValueListItem3.DisplayText = "Active"
        ValueListItem8.DataValue = "I"
        ValueListItem8.DisplayText = "Inactive"
        ValueList1.ValueListItems.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem3, ValueListItem8})
        Me.grdCC.DisplayLayout.ValueLists.AddRange(New Infragistics.Win.ValueList() {ValueList1})
        Me.grdCC.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdCC.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdCC.Location = New System.Drawing.Point(0, 0)
        Me.grdCC.Name = "grdCC"
        Me.grdCC.Size = New System.Drawing.Size(605, 159)
        Me.grdCC.TabIndex = 0
        Me.grdCC.TabStop = False
        '
        'UltraTabPageControl1
        '
        Me.UltraTabPageControl1.Controls.Add(Me.grdACH)
        Me.UltraTabPageControl1.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabPageControl1.Name = "UltraTabPageControl1"
        Me.UltraTabPageControl1.Size = New System.Drawing.Size(605, 159)
        '
        'grdACH
        '
        Appearance16.BackColor = System.Drawing.SystemColors.Window
        Appearance16.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.grdACH.DisplayLayout.Appearance = Appearance16
        UltraGridColumn2.Header.VisiblePosition = 0
        UltraGridColumn2.Hidden = True
        UltraGridColumn3.Header.VisiblePosition = 1
        UltraGridColumn3.Hidden = True
        UltraGridColumn4.Header.VisiblePosition = 2
        UltraGridColumn4.Hidden = True
        UltraGridColumn5.Header.VisiblePosition = 3
        UltraGridColumn5.Hidden = True
        UltraGridColumn6.Header.VisiblePosition = 4
        UltraGridColumn6.Hidden = True
        UltraGridColumn7.Header.VisiblePosition = 5
        UltraGridColumn7.Hidden = True
        UltraGridColumn8.Header.Caption = "Acct Name"
        UltraGridColumn8.Header.VisiblePosition = 6
        UltraGridColumn9.Header.Caption = "Auto Pay"
        UltraGridColumn9.Header.VisiblePosition = 12
        UltraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn9.Width = 73
        UltraGridColumn10.Header.Caption = "Default"
        UltraGridColumn10.Header.VisiblePosition = 13
        UltraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn10.Width = 62
        UltraGridColumn11.Header.Caption = "Type"
        UltraGridColumn11.Header.VisiblePosition = 7
        UltraGridColumn12.Header.Caption = "Routing No"
        UltraGridColumn12.Header.VisiblePosition = 9
        UltraGridColumn12.Hidden = True
        UltraGridColumn13.Header.Caption = "Account No"
        UltraGridColumn13.Header.VisiblePosition = 10
        UltraGridColumn13.Hidden = True
        UltraGridColumn14.Header.Caption = "Status"
        UltraGridColumn14.Header.VisiblePosition = 11
        UltraGridColumn15.Header.Caption = "Bank Name"
        UltraGridColumn15.Header.VisiblePosition = 8
        UltraGridColumn15.Width = 144
        UltraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        UltraGridColumn16.Header.Caption = "Web"
        UltraGridColumn16.Header.VisiblePosition = 14
        UltraGridColumn16.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        UltraGridColumn16.Width = 58
        Appearance17.TextHAlignAsString = "Right"
        UltraGridColumn17.CellAppearance = Appearance17
        UltraGridColumn17.Header.Caption = "Acct No Last 4"
        UltraGridColumn17.Header.VisiblePosition = 15
        UltraGridBand2.Columns.AddRange(New Object() {UltraGridColumn2, UltraGridColumn3, UltraGridColumn4, UltraGridColumn5, UltraGridColumn6, UltraGridColumn7, UltraGridColumn8, UltraGridColumn9, UltraGridColumn10, UltraGridColumn11, UltraGridColumn12, UltraGridColumn13, UltraGridColumn14, UltraGridColumn15, UltraGridColumn16, UltraGridColumn17})
        Me.grdACH.DisplayLayout.BandsSerializer.Add(UltraGridBand2)
        Me.grdACH.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance18.TextHAlignAsString = "Left"
        Me.grdACH.DisplayLayout.CaptionAppearance = Appearance18
        Appearance19.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance19.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance19.BorderColor = System.Drawing.SystemColors.Window
        Me.grdACH.DisplayLayout.GroupByBox.Appearance = Appearance19
        Appearance20.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdACH.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance20
        Me.grdACH.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.grdACH.DisplayLayout.GroupByBox.Hidden = True
        Appearance21.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance21.BackColor2 = System.Drawing.SystemColors.Control
        Appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance21.ForeColor = System.Drawing.SystemColors.GrayText
        Me.grdACH.DisplayLayout.GroupByBox.PromptAppearance = Appearance21
        Me.grdACH.DisplayLayout.MaxColScrollRegions = 1
        Me.grdACH.DisplayLayout.MaxRowScrollRegions = 1
        Me.grdACH.DisplayLayout.NewBandLoadStyle = Infragistics.Win.UltraWinGrid.NewBandLoadStyle.Hide
        Me.grdACH.DisplayLayout.NewColumnLoadStyle = Infragistics.Win.UltraWinGrid.NewColumnLoadStyle.Hide
        Appearance22.BackColor = System.Drawing.SystemColors.Window
        Appearance22.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grdACH.DisplayLayout.Override.ActiveCellAppearance = Appearance22
        Me.grdACH.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.grdACH.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdACH.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.[False]
        Me.grdACH.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.grdACH.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance23.BackColor = System.Drawing.SystemColors.Window
        Me.grdACH.DisplayLayout.Override.CardAreaAppearance = Appearance23
        Appearance24.BorderColor = System.Drawing.Color.Silver
        Appearance24.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.grdACH.DisplayLayout.Override.CellAppearance = Appearance24
        Me.grdACH.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        Me.grdACH.DisplayLayout.Override.CellPadding = 0
        Appearance25.BackColor = System.Drawing.SystemColors.Control
        Appearance25.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance25.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance25.BorderColor = System.Drawing.SystemColors.Window
        Me.grdACH.DisplayLayout.Override.GroupByRowAppearance = Appearance25
        Appearance26.TextHAlignAsString = "Left"
        Me.grdACH.DisplayLayout.Override.HeaderAppearance = Appearance26
        Me.grdACH.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.grdACH.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance27.BackColor = System.Drawing.SystemColors.Window
        Appearance27.BorderColor = System.Drawing.Color.Silver
        Me.grdACH.DisplayLayout.Override.RowAppearance = Appearance27
        Me.grdACH.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        Appearance28.BackColor = System.Drawing.SystemColors.ControlLight
        Me.grdACH.DisplayLayout.Override.TemplateAddRowAppearance = Appearance28
        Me.grdACH.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.grdACH.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        ValueList2.Key = "CUST_CREDIT_CARD_STATUS"
        ValueListItem1.DataValue = "A"
        ValueListItem1.DisplayText = "Active"
        ValueListItem2.DataValue = "I"
        ValueListItem2.DisplayText = "Inactive"
        ValueList2.ValueListItems.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem1, ValueListItem2})
        Me.grdACH.DisplayLayout.ValueLists.AddRange(New Infragistics.Win.ValueList() {ValueList2})
        Me.grdACH.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.grdACH.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdACH.Location = New System.Drawing.Point(0, 0)
        Me.grdACH.Name = "grdACH"
        Me.grdACH.Size = New System.Drawing.Size(605, 159)
        Me.grdACH.TabIndex = 1
        Me.grdACH.TabStop = False
        '
        'grpHeader
        '
        Me.grpHeader.Controls.Add(Me.chkCCEdit)
        Me.grpHeader.Controls.Add(Me.SplitContainer24)
        Me.grpHeader.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpHeader.Location = New System.Drawing.Point(0, 0)
        Me.grpHeader.Name = "grpHeader"
        Me.grpHeader.Size = New System.Drawing.Size(615, 339)
        Me.grpHeader.TabIndex = 2
        Me.grpHeader.Text = "                    Accounts"
        '
        'chkCCEdit
        '
        Me.chkCCEdit.Location = New System.Drawing.Point(9, 1)
        Me.chkCCEdit.Name = "chkCCEdit"
        Me.chkCCEdit.Size = New System.Drawing.Size(53, 15)
        Me.chkCCEdit.TabIndex = 1
        Me.chkCCEdit.Text = "Edit"
        '
        'SplitContainer24
        '
        Me.SplitContainer24.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer24.FixedPanel = System.Windows.Forms.FixedPanel.Panel2
        Me.SplitContainer24.IsSplitterFixed = True
        Me.SplitContainer24.Location = New System.Drawing.Point(3, 16)
        Me.SplitContainer24.Name = "SplitContainer24"
        Me.SplitContainer24.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer24.Panel1
        '
        Me.SplitContainer24.Panel1.Controls.Add(Me.tabgrids)
        '
        'SplitContainer24.Panel2
        '
        Me.SplitContainer24.Panel2.Controls.Add(Me.SplitContainer25)
        Me.SplitContainer24.Size = New System.Drawing.Size(609, 320)
        Me.SplitContainer24.SplitterDistance = 185
        Me.SplitContainer24.TabIndex = 152
        '
        'tabgrids
        '
        Me.tabgrids.Controls.Add(Me.UltraTabSharedControlsPage1)
        Me.tabgrids.Controls.Add(Me.UltraTabPageControl1)
        Me.tabgrids.Controls.Add(Me.UltraTabPageControl2)
        Me.tabgrids.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabgrids.Location = New System.Drawing.Point(0, 0)
        Me.tabgrids.Name = "tabgrids"
        Me.tabgrids.SharedControlsPage = Me.UltraTabSharedControlsPage1
        Me.tabgrids.Size = New System.Drawing.Size(609, 185)
        Me.tabgrids.TabIndex = 1
        UltraTab2.TabPage = Me.UltraTabPageControl2
        UltraTab2.Text = "Credit Cards"
        UltraTab1.TabPage = Me.UltraTabPageControl1
        UltraTab1.Text = "ACH"
        Me.tabgrids.Tabs.AddRange(New Infragistics.Win.UltraWinTabControl.UltraTab() {UltraTab2, UltraTab1})
        '
        'UltraTabSharedControlsPage1
        '
        Me.UltraTabSharedControlsPage1.Location = New System.Drawing.Point(-10000, -10000)
        Me.UltraTabSharedControlsPage1.Name = "UltraTabSharedControlsPage1"
        Me.UltraTabSharedControlsPage1.Size = New System.Drawing.Size(605, 159)
        '
        'SplitContainer25
        '
        Me.SplitContainer25.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer25.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer25.Name = "SplitContainer25"
        '
        'SplitContainer25.Panel1
        '
        Me.SplitContainer25.Panel1.Controls.Add(Me.grpCCPA)
        '
        'SplitContainer25.Panel2
        '
        Me.SplitContainer25.Panel2.Controls.Add(Me.grpButtons)
        Me.SplitContainer25.Size = New System.Drawing.Size(609, 131)
        Me.SplitContainer25.SplitterDistance = 504
        Me.SplitContainer25.TabIndex = 153
        '
        'grpCCPA
        '
        Me.grpCCPA.Controls.Add(Me.chkCUST_AUTOQ_WEB)
        Me.grpCCPA.Controls.Add(Me.cbeCUST_AUTO_CC_OPER)
        Me.grpCCPA.Controls.Add(Me.dteCUST_AUTO_CC_AUTH_DATE)
        Me.grpCCPA.Controls.Add(Me.chkCUST_AUTO_CC_AUTH)
        Me.grpCCPA.Controls.Add(Me.optCUST_AUTO_CCPA)
        Me.grpCCPA.Controls.Add(Me.txtCUST_AUTO_CCPA_NOTE)
        Me.grpCCPA.Controls.Add(Me.UltraLabel51)
        Me.grpCCPA.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpCCPA.Location = New System.Drawing.Point(0, 0)
        Me.grpCCPA.Name = "grpCCPA"
        Me.grpCCPA.Size = New System.Drawing.Size(504, 131)
        Me.grpCCPA.TabIndex = 13
        '
        'chkCUST_AUTOQ_WEB
        '
        Me.chkCUST_AUTOQ_WEB.Enabled = False
        Me.chkCUST_AUTOQ_WEB.Location = New System.Drawing.Point(326, 86)
        Me.chkCUST_AUTOQ_WEB.Name = "chkCUST_AUTOQ_WEB"
        Me.chkCUST_AUTOQ_WEB.Size = New System.Drawing.Size(73, 18)
        Me.chkCUST_AUTOQ_WEB.TabIndex = 3
        Me.chkCUST_AUTOQ_WEB.Text = "Web Auth"
        Me.chkCUST_AUTOQ_WEB.Visible = False
        '
        'cbeCUST_AUTO_CC_OPER
        '
        Appearance29.BackColor = System.Drawing.Color.White
        Appearance29.BackColorDisabled = System.Drawing.Color.White
        Appearance29.BackColorDisabled2 = System.Drawing.Color.White
        Me.cbeCUST_AUTO_CC_OPER.Appearance = Appearance29
        Me.cbeCUST_AUTO_CC_OPER.BackColor = System.Drawing.Color.White
        Me.cbeCUST_AUTO_CC_OPER.DisplayMember = "USER_NAME"
        Me.cbeCUST_AUTO_CC_OPER.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList
        ValueListItem12.DataValue = "A"
        ValueListItem12.DisplayText = "Billing Adj No"
        ValueListItem13.DataValue = "D"
        ValueListItem13.DisplayText = "Drop Ship No"
        ValueListItem10.DataValue = "O"
        ValueListItem10.DisplayText = "Order No"
        ValueListItem11.DataValue = "R"
        ValueListItem11.DisplayText = "Return No"
        Me.cbeCUST_AUTO_CC_OPER.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem12, ValueListItem13, ValueListItem10, ValueListItem11})
        Me.cbeCUST_AUTO_CC_OPER.Location = New System.Drawing.Point(212, 5)
        Me.cbeCUST_AUTO_CC_OPER.Name = "cbeCUST_AUTO_CC_OPER"
        Me.cbeCUST_AUTO_CC_OPER.ReadOnly = True
        Me.cbeCUST_AUTO_CC_OPER.Size = New System.Drawing.Size(273, 21)
        Me.cbeCUST_AUTO_CC_OPER.SortStyle = Infragistics.Win.ValueListSortStyle.Ascending
        Me.cbeCUST_AUTO_CC_OPER.TabIndex = 6
        Me.cbeCUST_AUTO_CC_OPER.TabStop = False
        Me.cbeCUST_AUTO_CC_OPER.ValueMember = "USER_ID"
        '
        'dteCUST_AUTO_CC_AUTH_DATE
        '
        Appearance30.BackColor = System.Drawing.Color.White
        Appearance30.BackColorDisabled = System.Drawing.Color.White
        Appearance30.BackColorDisabled2 = System.Drawing.Color.White
        Me.dteCUST_AUTO_CC_AUTH_DATE.Appearance = Appearance30
        Me.dteCUST_AUTO_CC_AUTH_DATE.BackColor = System.Drawing.Color.White
        Me.dteCUST_AUTO_CC_AUTH_DATE.DateTime = New Date(2007, 1, 27, 0, 0, 0, 0)
        Me.dteCUST_AUTO_CC_AUTH_DATE.Location = New System.Drawing.Point(115, 5)
        Me.dteCUST_AUTO_CC_AUTH_DATE.Name = "dteCUST_AUTO_CC_AUTH_DATE"
        Me.dteCUST_AUTO_CC_AUTH_DATE.ReadOnly = True
        Me.dteCUST_AUTO_CC_AUTH_DATE.Size = New System.Drawing.Size(84, 21)
        Me.dteCUST_AUTO_CC_AUTH_DATE.TabIndex = 5
        Me.dteCUST_AUTO_CC_AUTH_DATE.Value = New Date(2007, 1, 27, 0, 0, 0, 0)
        '
        'chkCUST_AUTO_CC_AUTH
        '
        Me.chkCUST_AUTO_CC_AUTH.Location = New System.Drawing.Point(6, 8)
        Me.chkCUST_AUTO_CC_AUTH.Name = "chkCUST_AUTO_CC_AUTH"
        Me.chkCUST_AUTO_CC_AUTH.Size = New System.Drawing.Size(103, 18)
        Me.chkCUST_AUTO_CC_AUTH.TabIndex = 4
        Me.chkCUST_AUTO_CC_AUTH.Text = "Customer Auth"
        '
        'optCUST_AUTO_CCPA
        '
        Me.optCUST_AUTO_CCPA.BorderStyle = Infragistics.Win.UIElementBorderStyle.Rounded1
        Me.optCUST_AUTO_CCPA.Enabled = False
        Me.optCUST_AUTO_CCPA.ItemOrigin = New System.Drawing.Point(1, 2)
        ValueListItem4.DataValue = "0"
        ValueListItem4.DisplayText = "None"
        ValueListItem5.DataValue = "1"
        ValueListItem5.DisplayText = "Reminder Q"
        ValueListItem6.DataValue = "2"
        ValueListItem6.DisplayText = "Auto Chg Q"
        ValueListItem7.DataValue = "3"
        ValueListItem7.DisplayText = "Auto Pay ACH"
        Me.optCUST_AUTO_CCPA.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem4, ValueListItem5, ValueListItem6, ValueListItem7})
        Me.optCUST_AUTO_CCPA.Location = New System.Drawing.Point(6, 80)
        Me.optCUST_AUTO_CCPA.Name = "optCUST_AUTO_CCPA"
        Me.optCUST_AUTO_CCPA.Size = New System.Drawing.Size(314, 24)
        Me.optCUST_AUTO_CCPA.TabIndex = 2
        Me.optCUST_AUTO_CCPA.Visible = False
        '
        'txtCUST_AUTO_CCPA_NOTE
        '
        Me.txtCUST_AUTO_CCPA_NOTE.Location = New System.Drawing.Point(6, 47)
        Me.txtCUST_AUTO_CCPA_NOTE.Name = "txtCUST_AUTO_CCPA_NOTE"
        Me.txtCUST_AUTO_CCPA_NOTE.Size = New System.Drawing.Size(479, 21)
        Me.txtCUST_AUTO_CCPA_NOTE.TabIndex = 7
        '
        'UltraLabel51
        '
        Me.UltraLabel51.AutoSize = True
        Me.UltraLabel51.Location = New System.Drawing.Point(6, 32)
        Me.UltraLabel51.Name = "UltraLabel51"
        Me.UltraLabel51.Size = New System.Drawing.Size(28, 14)
        Me.UltraLabel51.TabIndex = 152
        Me.UltraLabel51.Text = "Note"
        '
        'grpButtons
        '
        Me.grpButtons.Controls.Add(Me.cmdCCUpdate)
        Me.grpButtons.Controls.Add(Me.cmdCCCancel)
        Me.grpButtons.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpButtons.Location = New System.Drawing.Point(0, 0)
        Me.grpButtons.Name = "grpButtons"
        Me.grpButtons.Size = New System.Drawing.Size(101, 131)
        Me.grpButtons.TabIndex = 2
        '
        'cmdCCUpdate
        '
        Me.cmdCCUpdate.Location = New System.Drawing.Point(6, 11)
        Me.cmdCCUpdate.Name = "cmdCCUpdate"
        Me.cmdCCUpdate.Size = New System.Drawing.Size(76, 21)
        Me.cmdCCUpdate.TabIndex = 0
        Me.cmdCCUpdate.Text = "Update"
        '
        'cmdCCCancel
        '
        Me.cmdCCCancel.Location = New System.Drawing.Point(6, 38)
        Me.cmdCCCancel.Name = "cmdCCCancel"
        Me.cmdCCCancel.Size = New System.Drawing.Size(76, 23)
        Me.cmdCCCancel.TabIndex = 1
        Me.cmdCCCancel.Text = "Cancel"
        '
        'CreditCardQueue
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.grpHeader)
        Me.Name = "CreditCardQueue"
        Me.Size = New System.Drawing.Size(615, 339)
        Me.UltraTabPageControl2.ResumeLayout(False)
        CType(Me.grdCC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraTabPageControl1.ResumeLayout(False)
        CType(Me.grdACH, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpHeader, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpHeader.ResumeLayout(False)
        CType(Me.chkCCEdit, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer24.Panel1.ResumeLayout(False)
        Me.SplitContainer24.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer24, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer24.ResumeLayout(False)
        CType(Me.tabgrids, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabgrids.ResumeLayout(False)
        Me.SplitContainer25.Panel1.ResumeLayout(False)
        Me.SplitContainer25.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer25, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer25.ResumeLayout(False)
        CType(Me.grpCCPA, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpCCPA.ResumeLayout(False)
        Me.grpCCPA.PerformLayout()
        CType(Me.chkCUST_AUTOQ_WEB, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cbeCUST_AUTO_CC_OPER, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dteCUST_AUTO_CC_AUTH_DATE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkCUST_AUTO_CC_AUTH, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optCUST_AUTO_CCPA, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCUST_AUTO_CCPA_NOTE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpButtons, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpButtons.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grpHeader As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents chkCCEdit As ABSCS.ABSCheckBox
    Friend WithEvents SplitContainer24 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer25 As System.Windows.Forms.SplitContainer
    Friend WithEvents grpCCPA As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents optCUST_AUTO_CCPA As Infragistics.Win.UltraWinEditors.UltraOptionSet
    Friend WithEvents txtCUST_AUTO_CCPA_NOTE As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel51 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents cmdCCCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdCCUpdate As Infragistics.Win.Misc.UltraButton
    Friend WithEvents grdCC As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents chkCUST_AUTO_CC_AUTH As ABSCS.ABSCheckBox
    Friend WithEvents dteCUST_AUTO_CC_AUTH_DATE As Infragistics.Win.UltraWinEditors.UltraDateTimeEditor
    Friend WithEvents cbeCUST_AUTO_CC_OPER As Infragistics.Win.UltraWinEditors.UltraComboEditor
    Friend WithEvents grpButtons As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents chkCUST_AUTOQ_WEB As ABSCS.ABSCheckBox
    Friend WithEvents tabgrids As Infragistics.Win.UltraWinTabControl.UltraTabControl
    Friend WithEvents UltraTabSharedControlsPage1 As Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage
    Friend WithEvents UltraTabPageControl1 As Infragistics.Win.UltraWinTabControl.UltraTabPageControl
    Friend WithEvents grdACH As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents UltraTabPageControl2 As Infragistics.Win.UltraWinTabControl.UltraTabPageControl

End Class
