Imports nsoftware.InShip
Imports System.Drawing.Printing

Public Class WHFSHIP1

    Private divisionView As DataView
    Private providerView As DataView
    Private shipmethodView As DataView

    Private clsShip As New TAC.WHCSHIP1
    Private SHIP_CNTL_NO As String = String.Empty
    Private MASTER_TRACKING_NO As String = String.Empty

    Private rowWHTSHIP1 As DataRow
    Private tblPayorType As New DataTable
    Private shipPackageDetailList As New List(Of DPayments.DShippingSDK.PackageDetail)
    Private ShippingLabelDirectory As String
 
    Private WithEvents ultraComboPackage As Infragistics.Win.UltraWinGrid.UltraCombo = New Infragistics.Win.UltraWinGrid.UltraCombo

    Private isInternationalShipment As Boolean = False
    Private Const FedexSmartPost As Int16 = 26

    Private printerFound As Boolean = False
    Private registeredWeight As Decimal = 0

    ' Valid values for WHTSHIP1.STATUS
    '   I - Initial Setup before calling request.
    '   P - processed - label printed
    '   C - Cancelled
    '

#Region "ABS Standard Routines"

    Private Sub WHFSHIP1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If ScannerInUse AndAlso Not txtLoadAddress.Focused Then
            txtLoadAddress.Focus()
        End If
    End Sub

    ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim sql As String = String.Empty
        With dst

            Get_PARM("ASTPARM1")
            Get_PARM("SOTPARM1")

            Create_TDA(.Tables.Add, "SOTCARR1", "*")
            Create_TDA(.Tables.Add, "SOTSVIA1", "*")
            Create_TDA(.Tables.Add, "SOTCARR3", "*")
            .Tables("SOTCARR3").Columns.Add("CARRIER_DESC", GetType(System.String))
            .Tables("SOTCARR3").Columns.Add("CARRIER_SHIP_TYPE", GetType(System.String))
            .Tables("SOTCARR3").Columns.Add("PROVIDER_TYPE", GetType(System.String))

            Create_TDA(.Tables.Add, "WHTSHIP1", "*")
            Create_TDA(.Tables.Add, "WHTSHIP2", "*", 1)
            .Tables("WHTSHIP2").Columns.Add("POUNDS", GetType(System.Int16))
            .Tables("WHTSHIP2").Columns.Add("OUNCES", GetType(System.Int16))
            .Tables("WHTSHIP2").Columns.Add("SEL", GetType(System.Int16))
            .Tables("WHTSHIP2").Columns("SEL").DefaultValue = 0

            Create_TDA(.Tables.Add, "WHTSHIPC", "*", 1)
            .Tables("WHTSHIPC").Columns.Add("EXTENDED_PRICE", GetType(System.Decimal), "ISNULL(QUANTITY, 0) * ISNULL(UNIT_PRICE, 0)")
            .Tables("WHTSHIPC").Columns.Add("EXTENDED_WEIGHT", GetType(System.Decimal), "ISNULL(QUANTITY, 0) * ISNULL(WEIGHT, 0)")

            Create_TDA(.Tables.Add, "WHTSHIPS", "*", 1)
            .Tables("WHTSHIPS").Columns.Add("SEL", GetType(System.String))

            Create_TDA(.Tables.Add, "SHTSHIPS", "*")
            .Tables("SHTSHIPS").Columns.Add("SEL", GetType(System.String))
            Fill_Records("SHTSHIPS", String.Empty, True, "SELECT '0' SEL, SHTSHIPS.* FROM SHTSHIPS")

            Create_TDA(.Tables.Add, "WHTSDESC", "*")
            Fill_Records("WHTSDESC", String.Empty, True, "SELECT * FROM WHTSDESC")

            Fill_Records("SOTCARR1", "", True, "SELECT * FROM SOTCARR1")
            Fill_Records("SOTSVIA1", "", True, "SELECT * FROM SOTSVIA1")
            Fill_Records("SOTCARR3", "", True, "Select SOTCARR3.*, SOTCARR1.CARRIER_DESC, SOTCARR1.CARRIER_SHIP_TYPE, SOTCARR1.PROVIDER_TYPE From SOTCARR3, SOTCARR1 Where SOTCARR3.CARRIER_CODE = SOTCARR1.CARRIER_CODE")

            Create_TDA(.Tables.Add("WHTSHIP5_SF"), "WHTSHIP5", "*")
            Create_TDA(.Tables.Add("WHTSHIP5_ST"), "WHTSHIP5", "*")
            Create_TDA(.Tables.Add("WHTSHIP5_HL"), "WHTSHIP5", "*")

            Create_TDA(.Tables.Add("WHTSHIPP_S"), "WHTSHIPP", "*")
            Create_TDA(.Tables.Add("WHTSHIPP_D"), "WHTSHIPP", "*")

            With tblPayorType
                .Columns.Add("P_CODE", GetType(System.String))
                .Columns.Add("P_DESC", GetType(System.String))
            End With

            For Each valuePair As String In New String() {"S:Sender", "R:Recipient", "T:Third Party", "C:Collect"}
                tblPayorType.Rows.Add(New Object() {valuePair.Split(":")(0), valuePair.Split(":")(1)})
            Next

            Create_TDA(.Tables.Add, "ICTWHSE1", "*")
            Fill_Records("ICTWHSE1", "", True, "SELECT * FROM ICTWHSE1")

            Create_Lookup("WHTPKGM1")

        End With

        cmbDutiesPayor.DataSource = tblPayorType
        cmbDutiesPayor.SelectedRow = cmbDutiesPayor.Rows(0)

        cmbShipPayor.DataSource = tblPayorType
        cmbShipPayor.SelectedRow = cmbShipPayor.Rows(0)

        cmbWarehouse.DataSource = dst.Tables("ICTWHSE1")

        divisionView = New DataView(ASCDATA1.SelectDistinct(dst.Tables("SOTCARR3"), New String() {"DIVISION_CODE"}))
        divisionView.Sort = "DIVISION_CODE"
        cmbDivision.DataSource = divisionView
        If cmbDivision.Rows.Count > 0 Then
            cmbDivision.SelectedRow = cmbDivision.Rows(0)
            cmbDivision.Text = cmbDivision.SelectedRow.Cells("DIVISION_CODE").Value
        End If

        cmbDivision_ValueChanged(Nothing, Nothing)

        cmbDivision.DisplayLayout.Bands(0).Columns("DIVISION_CODE").Width = cmbDivision.Width
        cmbProvider.DisplayLayout.Bands(0).Columns("CARRIER_DESC").Width = cmbProvider.Width
        cmbShipMethod.DisplayLayout.Bands(0).Columns("SHIP_VIA_DESC").Width = cmbShipMethod.Width

        Dim SO_PARM_DEF_PICK_WHSE As String = (ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE") & String.Empty).ToString.Trim
        If SO_PARM_DEF_PICK_WHSE.Length > 0 Then
            For Each row As UltraWinGrid.UltraGridRow In cmbWarehouse.Rows
                If row.Cells("WHSE_CODE").Value = SO_PARM_DEF_PICK_WHSE Then
                    cmbWarehouse.SelectedRow = row
                    Exit For
                End If
            Next
        End If

        Bind_Controls(grpShipFrom, "WHTSHIP5_SF")
        Bind_Controls(grpShipTo, "WHTSHIP5_ST")
        Bind_Controls(grpHoldAtLocation, "WHTSHIP5_HL")
        Bind_Controls(grpSmartPost, "WHTSHIP1")

        grdWHTSHIP2.DataSource = dst.Tables("WHTSHIP2")

        Create_Summary(grdWHTSHIP2, "POUNDS", "Sum")
        Create_Summary(grdWHTSHIP2, "OUNCES", "Sum")
        Create_Summary(grdWHTSHIP2, "SHIP_PACKAGE_NO", "Count")

        Create_Summary(grdWHTSHIPC, "COMMODITY_DESC", "Count")
        Create_Summary(grdWHTSHIPC, "WEIGHT", "Sum")
        Create_Summary(grdWHTSHIPC, "EXTENDED_PRICE", "Sum")
        Create_Summary(grdWHTSHIPC, "EXTENDED_WEIGHT", "Sum")

        LoadSmartPostDropDowns()

        ASCMAIN1.Add_Value_List(grdWHTSHIP2, "SIGNATURE_TYPE", Nothing, New String() {":", "0:Default for requested service", "1:Adult", _
                            "2:Direct", "3:Indirect", "4:Not Required"}, 0)

        ASCMAIN1.Add_Value_List(grdWHTSHIP2, "COD_TYPE", Nothing, New String() {":", "0:Any Check", _
                             "1:Cashier's check or money order", "2:None"}, 0)

        With ultraComboPackage.DisplayLayout.Bands(0)

            ultraComboPackage.Font = grdWHTSHIP2.Font
            ultraComboPackage.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Default
            ultraComboPackage.DropDownStyle = UltraWinGrid.UltraComboStyle.DropDownList

            .Columns.Add("PKG_CODE")
            .Columns("PKG_CODE").Header.Caption = "Code"
            .Columns("PKG_CODE").Width = 75

            .Columns.Add("PKG_DESC")
            .Columns("PKG_DESC").Header.Caption = "Desc"
            .Columns("PKG_DESC").Width = 75

            .Columns.Add("PKG_D")
            .Columns("PKG_D").Header.Caption = "L x W x H"
            .Columns("PKG_D").Width = 200

        End With

        ultraComboPackage.DataSource = ASCDATA1.GetDataTable("SELECT PKG_CODE, PKG_DESC, PKG_L || ' x ' ||  PKG_W || ' x ' || PKG_H PKG_D FROM WHTPKGM1")
        ultraComboPackage.ValueMember = "PKG_CODE"
        ultraComboPackage.DisplayMember = "PKG_DESC"
        grdWHTSHIP2.DisplayLayout.Bands(0).Columns("PKG_CODE").EditorComponent = ultraComboPackage


        grdWHTSHIPC.DataSource = dst.Tables("WHTSHIPC")
        grdSHTSHIPS.DataSource = dst.Tables("SHTSHIPS")

        Dim Commodities As String() = GetCommodityUnits()
        ASCMAIN1.Add_Value_List(grdWHTSHIPC, "MANUFACTURER", "SELECT COUNTRY_CODE, COUNTRY_NAME FROM WHTMANF1")
        ASCMAIN1.Add_Value_List(grdWHTSHIPC, "QUANTITY_UOM", Nothing, Commodities, 0)

        dteShipDate.MinDate = CDate("01/01/1980")
        dteShipDate.MaxDate = DateAdd(DateInterval.Day, 10, DateTime.Now)
        dteShipDate.DateTime = DateTime.Now
        dteShipDate.Value = dteShipDate.DateTime

        'Set Max lengths
        txtShipCountry.MaxLength = dst.Tables("WHTSHIPP_S").Columns("PAYOR_COUNTRY").MaxLength
        txtDutiesCountry.MaxLength = dst.Tables("WHTSHIPP_D").Columns("PAYOR_COUNTRY").MaxLength

        txtShipAccountNo.MaxLength = dst.Tables("WHTSHIPP_S").Columns("PAYOR_ACCT_NO").MaxLength
        txtDutiesAccountNo.MaxLength = dst.Tables("WHTSHIPP_D").Columns("PAYOR_ACCT_NO").MaxLength

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = String.Empty

        Select Case eItemKey

            Case "Request Label", "Get Rates"
                isInternationalShipment = False

                If eItemKey = "Request Label" AndAlso Not printerFound Then
                    EMsg &= vbCrLf & "Label printer was not found you may not Request Labels."
                    Exit Select
                End If

                If cmbDivision.SelectedRow Is Nothing Then
                    EMsg &= vbCrLf & "Missing Division"
                    Exit Select
                End If

                If cmbProvider.SelectedRow Is Nothing Then
                    EMsg &= vbCrLf & "Missing Provider"
                    Exit Select
                End If

                If cmbShipMethod.SelectedRow Is Nothing Then
                    EMsg &= vbCrLf & "Missing Shipping Method"
                    Exit Select
                End If

                txtFromCountry.Text = txtFromCountry.Text.Trim.ToUpper
                txtToCountry.Text = txtToCountry.Text.Trim.ToUpper
                txtToState.Text = txtToState.Text.Trim.ToUpper

                If txtFromCountry.TextLength = 0 Then
                    EMsg &= vbCrLf & "Missing Ship From Country"
                End If

                If txtToCountry.TextLength = 0 Then
                    EMsg &= vbCrLf & "Missing Ship To Country"
                End If

                If grdWHTSHIP2.Rows.Count = 0 Then
                    EMsg &= vbCrLf & "At least one package must be entered"
                Else
                    For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdWHTSHIP2.Rows
                        ValidatePacking(grdRow)
                    Next
                End If


                If CDate(dteShipDate.DateTime.ToShortDateString) < CDate(DateTime.Now.ToShortDateString) Then
                    EMsg &= vbCrLf & "Ship date may not be less than today."
                End If

                ' Chnage USA to US.
                txtToCountry.Text = txtToCountry.Text.ToUpper.Trim
                If txtToCountry.Text.StartsWith("US") Then txtToCountry.Text = "US"

                txtFromCountry.Text = txtFromCountry.Text.ToUpper.Trim
                If txtFromCountry.Text.StartsWith("US") Then txtFromCountry.Text = "US"

                txtToPhone.Text = txtToPhone.Text.Trim
                If txtToPhone.TextLength = 0 Then
                    EMsg &= vbCrLf & "Missing Ship From Telephone."
                End If

                If EMsg.Length = 0 Then
                    ' Treat PR as international
                    isInternationalShipment = txtToCountry.Text <> "US" OrElse (txtToCountry.Text = "US" AndAlso txtToState.Text = "PR")

                    If isInternationalShipment Then
                        If grdWHTSHIPC.Rows.Count = 0 Then
                            EMsg &= vbCrLf & "International Shipments require Commodities"
                        Else
                            For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdWHTSHIPC.Rows
                                ValidateCommodity(grdRow)
                            Next
                        End If

                        If Val(numCustomsValue.Value & String.Empty) <= 0 Then
                            EMsg &= vbCrLf & "International shipments require Customs Value."
                        End If
                    End If
                End If

                ' Fedex Smart Post Specific
                If cmbProvider.SelectedRow.Cells("PROVIDER_TYPE").Value = WHCSHIP1.ProviderTypeFedex _
                    AndAlso cmbShipMethod.SelectedRow.Cells("CARRIER_PROD_CODE").Value = 26 Then
                    If cmbSmartPost.SelectedRow Is Nothing Then
                        EMsg &= vbCrLf & "Smart Post shipping methods require a Smart Post Type."
                    End If
                    If cmbSmartPostPackage.SelectedRow Is Nothing Then
                        EMsg &= vbCrLf & "Smart Post shipping methods require a package Type."
                    End If
                    If cmbProvider.SelectedRow.Cells("FEDEX_HUB_ID").Value & String.Empty = String.Empty Then
                        EMsg &= vbCrLf & "Smart Post shipping methods require a Hub Id."
                    End If
                End If

                ' Look for COD amounts
                If eItemKey = "Request Label" AndAlso cmbShipMethod.SelectedRow.Cells("COD_IND").Value & String.Empty = "1" Then
                    For Each rowWHTSHIP2 As DataRow In dst.Tables("WHTSHIP2").Select("", "", DataViewRowState.CurrentRows)
                        Dim COD_AMOUNT As Decimal = Val(rowWHTSHIP2.Item("COD_AMOUNT") & String.Empty)
                        Dim COD_TYPE As String = Val(rowWHTSHIP2.Item("COD_TYPE") & String.Empty)

                        If (COD_TYPE <> "0" AndAlso COD_TYPE <> "1") OrElse COD_AMOUNT < 0 Then
                            EMsg &= vbCrLf & "COD shipments require all packages have a dollar value and proper COD Type setting"
                            Exit For
                        End If

                    Next
                End If

                If EMsg.Length = 0 Then
                    Select Case cmbProvider.SelectedRow.Cells("PROVIDER_TYPE").Value
                        Case WHCSHIP1.ProviderTypeFedex
                            If Not isInternationalShipment Then
                                clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.FederalExpress)
                            Else
                                clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.FederalExpressInternational)
                            End If
                        Case WHCSHIP1.ProviderTypeUPS
                            clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.UPS)
                        Case WHCSHIP1.ProviderTypeUSPS
                            clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.USPS)
                        Case WHCSHIP1.ProviderTypeCanada
                            clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.CanadaPost)
                        Case Else
                            EMsg &= vbCrLf & "Invalid or Missing Provider"
                            Exit Select
                    End Select
                End If

            Case "Clear"
                If MessageBox.Show("Do you want to clear the contents of the screen?", "Clear", _
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

            Case "Cancel Shipment"
                If dst.Tables("WHTSHIP2").Select("SEL = '1'").Length = 0 Then
                    EMsg &= vbCr & "You must select the shipments to cancel."
                    Exit Select
                End If

                If MessageBox.Show("Do you want to cancel the selected packages?", "Cancel Shipment", _
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                Select Case cmbProvider.SelectedRow.Cells("PROVIDER_TYPE").Value
                    Case WHCSHIP1.ProviderTypeFedex
                        If Not isInternationalShipment Then
                            clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.FederalExpress)
                        Else
                            clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.FederalExpressInternational)
                        End If
                    Case WHCSHIP1.ProviderTypeUPS
                        clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.UPS)
                    Case WHCSHIP1.ProviderTypeUSPS
                        clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.USPS)
                    Case WHCSHIP1.ProviderTypeCanada
                        clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.CanadaPost)
                    Case Else
                        EMsg &= vbCrLf & "Invalid or Missing Provider"
                        Exit Select
                End Select

            Case "Reprint Label"
                If Not printerFound AndAlso Not (ASCMAIN1.USER_ID = "edz" AndAlso ASCMAIN1.Running_in_VS) Then
                    EMsg &= vbCrLf & "Label printer was not found you may not Reprint Labels."
                    Exit Select
                End If

                If MessageBox.Show("Do you want to Reprint labels for the shipment displayed on the screen?", "Reprint Label", _
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Request Label"
                EntryMode = "S"
                If RequestShippingLabel() Then
                    Me.Mode_Settings(False)
                End If

            Case "Get Rates"
                EntryMode = "R"
                RequestShippingLabel()

            Case "Clear"
                Me.Mode_Settings(False)

            Case "Cancel Shipment"
                CancelShipment()
                Mode_Settings(False)

            Case "Reprint Label"
                ReprintShippingLabel()
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Request label").Settings.Enabled = DefaultableBoolean.True
                .Groups("Screen Control").Items("Clear").Settings.Enabled = DefaultableBoolean.True
                .Groups("Screen Control").Items("Cancel Shipment").Settings.Enabled = DefaultableBoolean.False
                .Groups("Screen Control").Items("Reprint Label").Settings.Enabled = DefaultableBoolean.False

                If cmbShipMethod.SelectedRow IsNot Nothing AndAlso cmbShipMethod.SelectedRow.Cells("CARRIER_PROD_CODE").Value = FedexSmartPost Then
                    .Groups("Smart Post").Visible = True
                Else
                    .Groups("Smart Post").Visible = False
                End If
            End With
        End If

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Private Sub Clear_Record()
        MyBase.EnforceConstraints(False)

        For Each tableName As String In New String() {"WHTSHIP1", "WHTSHIP2", "WHTSHIP5_SF", _
                                                      "WHTSHIP5_ST", "WHTSHIP5_HL", "WHTSHIPC", _
                                                      "WHTSHIPP_S", "WHTSHIPP_D", "WHTSHIPS"}
            dst.Tables(tableName).Rows.Clear()
        Next

        For Each row As DataRow In dst.Tables("SHTSHIPS").Select
            row.Item("SEL") = "0"
        Next
        dst.Tables("SHTSHIPS").AcceptChanges()

        ' Create Shells for data
        ' Create working header record for WHTSHIP1
        SHIP_CNTL_NO = "XX"

        grdWHTSHIP2.DisplayLayout.Bands(0).Columns("SEL").Hidden = True

        For Each tableName As String In New String() {"WHTSHIP5_SF", "WHTSHIP5_ST", "WHTSHIP5_HL", "WHTSHIPP_S", "WHTSHIPP_D"}
            Dim row As DataRow = dst.Tables(tableName).NewRow
            row.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO

            Select Case tableName
                Case "WHTSHIP5_SF" : row.Item("SHIP_ADDR_TYPE") = "SF"
                Case "WHTSHIP5_ST" : row.Item("SHIP_ADDR_TYPE") = "ST"
                Case "WHTSHIP5_HL" : row.Item("SHIP_ADDR_TYPE") = "HL"
                Case "WHTSHIPP_S" : row.Item("PAYOR_TYPE") = "S"
                Case "WHTSHIPP_D" : row.Item("PAYOR_TYPE") = "D"
            End Select
            dst.Tables(tableName).Rows.Add(row)
        Next

        MyBase.EnforceConstraints(True)

        Show_Filter(grdWHTSHIP2, False)
        grdWHTSHIP2.DisplayLayout.GroupByBox.Hidden = True

        Show_Filter(grdSHTSHIPS, False)
        grdSHTSHIPS.DisplayLayout.GroupByBox.Hidden = True

        chkFromResidential.Checked = False
        chkToResidential.Checked = False

        chkToResidential.Checked = False
        chkToPOBox.Checked = False
        chkFromResidential.Checked = False
        chkSignature.Checked = False

        rowWHTSHIP1 = dst.Tables("WHTSHIP1").NewRow
        rowWHTSHIP1.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO

        ' Keep previously selected options.
        If cmbProvider.SelectedRow IsNot Nothing Then
            cmbDivision_ValueChanged(Nothing, Nothing)
        End If

        If cmbProvider.SelectedRow IsNot Nothing Then
            rowWHTSHIP1.Item("CARRIER_CODE") = cmbProvider.SelectedRow.Cells("CARRIER_CODE").Value
        Else
            rowWHTSHIP1.Item("CARRIER_CODE") = "*"
        End If

        If cmbShipMethod.SelectedRow IsNot Nothing Then
            rowWHTSHIP1.Item("CARRIER_PROD_CODE") = cmbShipMethod.SelectedRow.Cells("CARRIER_PROD_CODE").Value
            rowWHTSHIP1.Item("SHIP_VIA_CODE") = cmbShipMethod.SelectedRow.Cells("SHIP_VIA_CODE").Value
        Else
            rowWHTSHIP1.Item("CARRIER_PROD_CODE") = 0
        End If

        rowWHTSHIP1.Item("CARRIER_ACCOUNT_NO") = String.Empty
        rowWHTSHIP1.Item("STATUS") = "P"
        'rowWHTSHIP1.Item("ERROR_MSG") = String.Empty
        rowWHTSHIP1.Item("SHIP_DATE") = DateTime.Now.ToString("MM/dd/yyyy")
        dteShipDate.Value = DateTime.Now
        rowWHTSHIP1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        rowWHTSHIP1.Item("OPS_YYYYWW") = ASCMAIN1.CYW
        'rowWHTSHIP1.Item("CUST_CODE") = String.Empty
        rowWHTSHIP1.Item("INIT_DATE") = DateTime.Now
        rowWHTSHIP1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowWHTSHIP1.Item("LAST_DATE") = DateTime.Now
        rowWHTSHIP1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        'rowWHTSHIP1.Item("MASTER_TRACKING_NO") = String.Empty
        rowWHTSHIP1.Item("CUSTOMS_VALUE") = 0
        dst.Tables("WHTSHIP1").Rows.Add(rowWHTSHIP1)

        Dim rowWHTSHIP5_SF As DataRow
        rowWHTSHIP5_SF = dst.Tables("WHTSHIP5_SF").Rows(0)
        rowWHTSHIP5_SF.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
        rowWHTSHIP5_SF.Item("SHIP_ADDR_TYPE") = "SF"
        cmbWarehouse_ValueChanged(Nothing, Nothing)

        txtLoadAddress.Clear()
        MASTER_TRACKING_NO = String.Empty

        SetReadOnly(False)
    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("")

        MyBase.EnforceConstraints(False)


        MyBase.EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()
        Try
            BeginTrans()
            Select Case EntryMode

                Case "S" ' Shipping label

                    Update_Record_TDA("WHTSHIP1", "SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "'")
                    Update_Record_TDA("WHTSHIP2", "SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "'")

                    Update_Record_TDA("WHTSHIP5_SF", "DELETE FROM WHTSHIP5 WHERE SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "' AND SHIP_ADDR_TYPE = 'SF'")
                    Update_Record_TDA("WHTSHIP5_ST", "DELETE FROM WHTSHIP5 WHERE SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "' AND SHIP_ADDR_TYPE = 'ST'")
                    Update_Record_TDA("WHTSHIP5_HL", "DELETE FROM WHTSHIP5 WHERE SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "' AND SHIP_ADDR_TYPE = 'HL'")

                    Dim row As DataRow = dst.Tables("WHTSHIPP_S").Select("PAYOR_TYPE = 'S'")(0)
                    row.Item("PAYOR_ACCT_NO") = txtShipAccountNo.Text
                    row.Item("PAYOR_COUNTRY") = txtShipCountry.Text

                    row = dst.Tables("WHTSHIPP_D").Select("PAYOR_TYPE = 'D'")(0)
                    row.Item("PAYOR_ACCT_NO") = txtDutiesAccountNo.Text
                    row.Item("PAYOR_COUNTRY") = txtDutiesCountry.Text

                    Update_Record_TDA("WHTSHIPP_S", "DELETE FROM WHTSHIPP WHERE SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "' AND PAYOR_TYPE = 'S'")
                    Update_Record_TDA("WHTSHIPP_D", "DELETE FROM WHTSHIPP WHERE SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "' AND PAYOR_TYPE = 'D'")

                    For Each rowWHTSHIPC As DataRow In dst.Tables("WHTSHIPC").Select("", "", DataViewRowState.CurrentRows)
                        If rowWHTSHIPC.Item("MANUFACTURER") & String.Empty = String.Empty Then
                            rowWHTSHIPC.Item("MANUFACTURER") = "US"
                        End If
                    Next
                    Update_Record_TDA("WHTSHIPC", "SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "'")

                    dst.Tables("WHTSHIPS").Rows.Clear()
                    For Each row In dst.Tables("SHTSHIPS").Select("SEL = '1'")
                        dst.Tables("WHTSHIPS").Rows.Add(New Object() {SHIP_CNTL_NO, row.Item("SPCL_SHIP_CODE")})
                    Next
                    Update_Record_TDA("WHTSHIPS", "SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "'")

            End Select

            CommitTrans()

        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

#End Region

#Region "Form Controls"

    Private Sub WHFSHIP1_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        SetUpPortsAndPrinters()
    End Sub

    Private Sub cmbDivision_ValueChanged(sender As Object, e As System.EventArgs) Handles cmbDivision.ValueChanged
        providerView = New DataView(dst.Tables("SOTCARR3"))
        providerView.Sort = "CARRIER_DESC"
        providerView.RowFilter = "DIVISION_CODE = ''"
        cmbProvider.DataSource = providerView
        If cmbDivision.SelectedRow IsNot Nothing Then
            providerView.RowFilter = "DIVISION_CODE = '" & cmbDivision.SelectedRow.Cells("DIVISION_CODE").Value & "'"
            cmbProvider.SelectedRow = cmbProvider.Rows(0)
            cmbProvider.Text = cmbProvider.SelectedRow.Cells("CARRIER_DESC").Value
        End If
        cmbProvider_ValueChanged(Nothing, Nothing)
    End Sub

    Private Sub cmbProvider_ValueChanged(sender As Object, e As System.EventArgs) Handles cmbProvider.ValueChanged
        shipmethodView = New DataView(dst.Tables("SOTSVIA1"))
        shipmethodView.Sort = "SHIP_VIA_DESC"
        shipmethodView.RowFilter = "SHIP_VIA_CODE = ''"
        cmbShipMethod.DataSource = shipmethodView
        If cmbProvider.SelectedRow IsNot Nothing Then
            shipmethodView.RowFilter = "CARRIER_CODE = '" & cmbProvider.SelectedRow.Cells("CARRIER_CODE").Value & "' AND SHIP_VIA_STATUS = 'A' AND ISNULL(CARRIER_PROD_CODE , '') <> ''"
            cmbShipMethod.SelectedRow = cmbShipMethod.Rows(0)
            cmbShipMethod.Text = cmbShipMethod.SelectedRow.Cells("SHIP_VIA_DESC").Value
            ASCMAIN1.Add_Value_List(grdWHTSHIP2, "PACKAGING_TYPE", "SELECT PACKAGE_CODE, PACKAGE_DESC FROM SOTCARR4 WHERE CARRIER_CODE = '" & cmbProvider.SelectedRow.Cells("CARRIER_CODE").Value & "'")
            ASCMAIN1.Add_Value_List(grdWHTSHIPC, "QUANTITY_UOM", "SELECT CARRIER_UOM, CARRIER_UOM_DESC FROM SOTCARRU WHERE CARRIER_CODE = '" & cmbProvider.SelectedRow.Cells("CARRIER_CODE").Value & "'")
        Else
            ASCMAIN1.Add_Value_List(grdWHTSHIP2, "PACKAGING_TYPE", "SELECT PACKAGE_CODE, PACKAGE_DESC FROM SOTCARR4 WHERE CARRIER_CODE = ''")
            ASCMAIN1.Add_Value_List(grdWHTSHIPC, "QUANTITY_UOM", "SELECT CARRIER_UOM, CARRIER_UOM_DESC FROM SOTCARRU WHERE CARRIER_CODE = ''")
        End If

    End Sub

    Private Sub cmbShipMethod_ValueChanged(sender As Object, e As System.EventArgs) Handles cmbShipMethod.ValueChanged
        If cmbShipMethod.SelectedRow IsNot Nothing AndAlso dst.Tables("WHTSHIP1").Rows.Count > 0 Then
            dst.Tables("WHTSHIP1").Rows(0).Item("CARRIER_CODE") = cmbShipMethod.SelectedRow.Cells("CARRIER_CODE").Value & String.Empty
        End If

        If cmbShipMethod.SelectedRow IsNot Nothing AndAlso cmbShipMethod.SelectedRow.Cells("CARRIER_PROD_CODE").Value = FedexSmartPost Then
            UltraExplorerBar1.Groups("Smart Post").Visible = True
        Else
            UltraExplorerBar1.Groups("Smart Post").Visible = False
        End If
    End Sub

    Private Sub grdWHTSHIP2_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdWHTSHIP2.BeforeRowUpdate

        EMsg = String.Empty

        ValidatePacking(e.Row)

        If EMsg.Length > 0 Then
            e.Cancel = True
            MessageBox.Show(EMsg, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

    End Sub

    Private Sub grdWHTSHIPC_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdWHTSHIPC.BeforeRowUpdate
        EMsg = String.Empty

        ValidateCommodity(e.Row)

        If EMsg.Length > 0 Then
            e.Cancel = True
            MessageBox.Show(EMsg, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub grdWHTSHIPC_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdWHTSHIPC.InitializeLayout
        Dim UC As New UltraWinGrid.UltraCombo
        UC.DataSource = dst.Tables("WHTSDESC")
        UC.DisplayLayout.Bands(0).ColHeadersVisible = False
        UC.DisplayLayout.Bands(0).Columns("SHIP_DESC").Width = e.Layout.Bands(0).Columns("COMMODITY_DESC").Width
        UC.Font = grdWHTSHIPC.Font

        e.Layout.Bands(0).Columns("COMMODITY_DESC").EditorComponent = UC
    End Sub

    Private Sub ultraComboPackage_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ultraComboPackage.ValueChanged

        Dim vl As IValueList = Me.ultraComboPackage

        If vl.SelectedItemIndex < 0 Then
            ' NOTHING
        Else
            Dim PKG_CODE As String = vl.GetText(vl.SelectedItemIndex)
            Dim rowWHTPKGM1 As DataRow = LookUp("WHTPKGM1", PKG_CODE)

            If rowWHTPKGM1 IsNot Nothing Then
                grdWHTSHIP2.ActiveRow.Cells("LENGTH").Value = Val(rowWHTPKGM1.Item("PKG_L") & String.Empty)
                grdWHTSHIP2.ActiveRow.Cells("WIDTH").Value = Val(rowWHTPKGM1.Item("PKG_W") & String.Empty)
                grdWHTSHIP2.ActiveRow.Cells("HEIGHT").Value = Val(rowWHTPKGM1.Item("PKG_H") & String.Empty)
            End If
        End If

    End Sub

    Private Sub txtLoadAddress_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtLoadAddress.KeyPress

        If Asc(e.KeyChar) = 13 Then

            Dim rowWHTSHIP5_ST As DataRow = dst.Tables("WHTSHIP5_ST").Rows(0)
            Dim value As String = txtLoadAddress.Text.Trim.ToUpper
            If value.Length = 0 Then Exit Sub

            Dim KeyPressEventArgs As New System.Windows.Forms.KeyPressEventArgs(vbCrLf)

            txtLoadAddress.Text = value

            Select Case optLoadAddress.Value

                Case "P" ' Pick Ticket
                    value = ASCMAIN1.Format_Field(value, "PICK_NO")
                    Dim rowSOTPICK1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTPICK1 WHERE PICK_NO = :PARM1", "V", New Object() {value})
                    If rowSOTPICK1 Is Nothing OrElse rowSOTPICK1.Item("INV_NO") & String.Empty = String.Empty Then
                        MessageBox.Show("Invalid or unprocessed Pick Ticket Number", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If
                    optLoadAddress.Value = "I"
                    txtLoadAddress.Text = rowSOTPICK1.Item("INV_NO")
                    txtLoadAddress_KeyPress(Nothing, KeyPressEventArgs)
                    optLoadAddress.Value = "P"
                    txtLoadAddress.Text = value
                    Exit Sub

                Case "W"
                    Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTINVH1 WHERE ORDR_NO_WEB = :PARM1", "V", New Object() {value})
                    If rowSOTINVH1 Is Nothing Then
                        MessageBox.Show("Invalid Web Order Number", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If
                    optLoadAddress.Value = "I"
                    txtLoadAddress.Text = rowSOTINVH1.Item("INV_NO")
                    txtLoadAddress_KeyPress(Nothing, KeyPressEventArgs)
                    Exit Sub

                Case "T" ' Tracking No
                    Dim tbl As DataTable = ASCDATA1.GetDataTable("select * from whtship1 where MASTER_TRACKING_NO = '" & value & "'")

                    If tbl.Rows.Count = 0 Then
                        ' See if it is part of a Milti package
                        tbl = ASCDATA1.GetDataTable("select * from whtship2 where TRACKING_NO = '" & value & "'")
                    End If

                    ' Done for fedex, the scan contains extra data
                    If tbl.Rows.Count = 0 Then
                        ' Scanning in the Tracking Number reads in the additional leading characters.
                        If value.Length > 15 Then
                            value = value.Substring(value.Length - 15)
                            txtLoadAddress.Text = value
                        End If
                        tbl = ASCDATA1.GetDataTable("select * from whtship1 where MASTER_TRACKING_NO = '" & value & "'")
                    End If
 
                    If tbl.Rows.Count = 0 Then
                        MessageBox.Show("The provided Tracking No could not be found.", "Load", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    SHIP_CNTL_NO = tbl.Rows(0).Item("SHIP_CNTL_NO") & String.Empty
                    tbl = ASCDATA1.GetDataTable("select * from whtship1 where SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "'")

                    ' Valid values for WHTSHIP1.STATUS
                    '   I - Initial Setup before calling request.
                    '   P - processed - label printed
                    '   C - Cancelled

                    If tbl.Rows(0).Item("STATUS") & String.Empty = "P" Then
                        UltraExplorerBar1.Groups("Screen Control").Items("Cancel Shipment").Settings.Enabled = DefaultableBoolean.True
                        UltraExplorerBar1.Groups("Screen Control").Items("Reprint Label").Settings.Enabled = DefaultableBoolean.True
                        UltraExplorerBar1.Groups("Screen Control").Items("Request Label").Settings.Enabled = DefaultableBoolean.False
                        grdWHTSHIP2.DisplayLayout.Bands(0).Columns("SEL").Hidden = False
                        grdWHTSHIP2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                    Else
                        UltraExplorerBar1.Groups("Screen Control").Items("Cancel Shipment").Settings.Enabled = DefaultableBoolean.False
                        UltraExplorerBar1.Groups("Screen Control").Items("Reprint Label").Settings.Enabled = DefaultableBoolean.False
                        UltraExplorerBar1.Groups("Screen Control").Items("Request Label").Settings.Enabled = DefaultableBoolean.True
                        grdWHTSHIP2.DisplayLayout.Bands(0).Columns("SEL").Hidden = True
                        grdWHTSHIP2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    End If

                    Fill_Records("WHTSHIP1", SHIP_CNTL_NO)
                    Fill_Records("WHTSHIP2", SHIP_CNTL_NO)
                    Fill_Records("WHTSHIPC", SHIP_CNTL_NO)
                    Fill_Records("WHTSHIPS", SHIP_CNTL_NO)

                    For Each addressType As String In New String() {"SF", "ST", "HL"}
                        Fill_Records("WHTSHIP5_" & addressType, "", True, "SELECT * FROM WHTSHIP5 WHERE SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "' AND SHIP_ADDR_TYPE = '" & addressType & "'")
                    Next

                    For Each payorType As String In New String() {"D", "S"}
                        Fill_Records("WHTSHIPP_" & payorType, "", True, "SELECT * FROM WHTSHIPP WHERE SHIP_CNTL_NO = '" & SHIP_CNTL_NO & "' AND PAYOR_TYPE = '" & payorType & "'")
                    Next

                    Sort_grdColumns(grdWHTSHIP2, "SHIP_PACKAGE_NO")
                    Sort_grdColumns(grdWHTSHIPC, "COMMODITY_DESC")

                    For Each rowWHTSHIP2 As DataRow In dst.Tables("WHTSHIP2").Select()
                        Dim WEIGHT As Int16 = Val(rowWHTSHIP2.Item("WEIGHT") & String.Empty)
                        Dim POUNDS As Int16 = WEIGHT \ 16
                        Dim OUNCES As Int16 = WEIGHT Mod 16

                        rowWHTSHIP2.Item("POUNDS") = POUNDS
                        rowWHTSHIP2.Item("OUNCES") = OUNCES
                    Next

                    Dim SHIP_VIA_CODE As String = dst.Tables("WHTSHIP1").Rows(0).Item("SHIP_VIA_CODE") & String.Empty
                    For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In cmbShipMethod.Rows
                        If row.Cells("SHIP_VIA_CODE").Value = SHIP_VIA_CODE Then
                            cmbShipMethod.SelectedRow = row
                            Exit For
                        End If
                    Next

                    If Not (ASCMAIN1.USER_SECURITY_CODEs.Contains("SY") OrElse ASCMAIN1.USER_SECURITY_CODEs.Contains("WL")) Then
                        SetReadOnly(True)
                    End If

                Case "C" ' Customer
                    value = value.ToUpper.Trim
                    Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", value)
                    If rowARTCUST1 Is Nothing Then
                        MessageBox.Show("Invalid Customer code.", "Error", MessageBoxButtons.OK)
                        Exit Sub
                    End If

                    rowWHTSHIP5_ST.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                    rowWHTSHIP5_ST.Item("SHIP_ADDR_TYPE") = "ST"
                    rowWHTSHIP5_ST.Item("SHIP_FIRST_NAME") = rowARTCUST1.Item("CUST_CONTACT")
                    rowWHTSHIP5_ST.Item("SHIP_COMPANY") = rowARTCUST1.Item("CUST_NAME") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_ADDR1") = rowARTCUST1.Item("CUST_ADDR1") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_ADDR2") = rowARTCUST1.Item("CUST_ADDR2") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_CITY") = rowARTCUST1.Item("CUST_CITY") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_STATE") = rowARTCUST1.Item("CUST_STATE") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_ZIP_CODE") = rowARTCUST1.Item("CUST_ZIP_CODE") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_COUNTRY_CODE") = rowARTCUST1.Item("CUST_COUNTRY") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_PHONE") = rowARTCUST1.Item("CUST_PHONE") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_RESIDENTIAL") = "0"
                    rowWHTSHIP5_ST.Item("SHIP_PO_BOX") = "0"
                    If rowWHTSHIP5_ST.Item("SHIP_COUNTRY_CODE") & String.Empty = String.Empty Then rowWHTSHIP5_ST.Item("SHIP_COUNTRY_CODE") = "US"
                    rowWHTSHIP5_ST.AcceptChanges()
                    txtLoadAddress.Text = value

                Case "V" ' Vendor Address
                    value = value.ToUpper.Trim
                    Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", value)
                    If rowAPTVEND1 Is Nothing Then
                        MessageBox.Show("Invalid Vendor code.", "Error", MessageBoxButtons.OK)
                        Exit Sub
                    End If

                    rowWHTSHIP5_ST.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                    rowWHTSHIP5_ST.Item("SHIP_ADDR_TYPE") = "ST"
                    rowWHTSHIP5_ST.Item("SHIP_FIRST_NAME") = rowAPTVEND1.Item("VEND_CONTACT")
                    rowWHTSHIP5_ST.Item("SHIP_COMPANY") = rowAPTVEND1.Item("VEND_NAME") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_ADDR1") = rowAPTVEND1.Item("VEND_ADDR1") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_ADDR2") = rowAPTVEND1.Item("VEND_ADDR2") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_CITY") = rowAPTVEND1.Item("VEND_CITY") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_STATE") = rowAPTVEND1.Item("VEND_STATE") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_ZIP_CODE") = rowAPTVEND1.Item("VEND_ZIP_CODE") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_COUNTRY_CODE") = rowAPTVEND1.Item("VEND_COUNTRY") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_PHONE") = rowAPTVEND1.Item("VEND_PHONE") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_RESIDENTIAL") = "0"
                    rowWHTSHIP5_ST.Item("SHIP_PO_BOX") = "0"
                    If rowWHTSHIP5_ST.Item("SHIP_COUNTRY_CODE") & String.Empty = String.Empty Then rowWHTSHIP5_ST.Item("SHIP_COUNTRY_CODE") = "US"
                    rowWHTSHIP5_ST.AcceptChanges()
                    txtLoadAddress.Text = value

                Case "I", "S" ' Sales Order 
                    Dim rowSOTINVH1 As DataRow = Nothing
                    Dim SHIP_BOL_NO As String = String.Empty

                    value = value.ToUpper.Trim
                    If optLoadAddress.Value = "I" Then
                        value = ASCMAIN1.Format_Field(value, "INV_NO")
                        rowSOTINVH1 = LookUp("SOTINVH1", New String() {"I", value})
                        If rowSOTINVH1 Is Nothing Then
                            MessageBox.Show("Invalid Invoice Number.", "Error", MessageBoxButtons.OK)
                            Exit Sub
                        End If
                        txtLoadAddress.Text = value
                        value = rowSOTINVH1.Item("ORDR_NO") & String.Empty
                        SHIP_BOL_NO = rowSOTINVH1.Item("SHIP_BOL_NO") & String.Empty
                        If SHIP_BOL_NO.Length > 0 Then
                            Dim rowWHTSIP1X As DataRow = ASCDATA1.GetDataRow("SELECT * FROM WHTSHIP1 WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")
                            If rowWHTSIP1X IsNot Nothing Then
                                Dim MASTER_TRACKING_NO As String = rowWHTSIP1X.Item("MASTER_TRACKING_NO") & String.Empty
                                If MASTER_TRACKING_NO.Length > 0 Then
                                    optLoadAddress.Value = "T"
                                    txtLoadAddress.Text = MASTER_TRACKING_NO
                                    txtLoadAddress_KeyPress(Nothing, KeyPressEventArgs)
                                    optLoadAddress.Value = "I"
                                    txtLoadAddress.Text = value
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If

                    value = ASCMAIN1.Format_Field(value, "ORDR_NO")
                    Dim rowSOTORDR5 As DataRow = LookUp("SOTORDR5", New String() {value, "ST"})
                    If rowSOTORDR5 Is Nothing Then
                        rowSOTORDR5 = LookUp("SOTORDR5", New String() {value, "BT"})
                    End If
                    If rowSOTORDR5 Is Nothing Then
                        MessageBox.Show("Invalid Sales Order.", "Error", MessageBoxButtons.OK)
                        Exit Sub
                    End If

                    rowWHTSHIP5_ST.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                    rowWHTSHIP5_ST.Item("SHIP_ADDR_TYPE") = "ST"
                    rowWHTSHIP5_ST.Item("SHIP_FIRST_NAME") = rowSOTORDR5.Item("CUST_CONTACT")
                    rowWHTSHIP5_ST.Item("SHIP_COMPANY") = rowSOTORDR5.Item("CUST_NAME") & String.Empty
                    If (rowSOTORDR5.Item("CUST_ADDR1") & String.Empty).ToString.Trim.Length = 0 Then
                        rowWHTSHIP5_ST.Item("SHIP_ADDR1") = rowSOTORDR5.Item("CUST_ADDR2") & String.Empty
                        rowWHTSHIP5_ST.Item("SHIP_ADDR2") = rowSOTORDR5.Item("CUST_ADDR3") & String.Empty
                    Else
                        rowWHTSHIP5_ST.Item("SHIP_ADDR1") = rowSOTORDR5.Item("CUST_ADDR1") & String.Empty
                        rowWHTSHIP5_ST.Item("SHIP_ADDR2") = rowSOTORDR5.Item("CUST_ADDR2") & String.Empty
                    End If
                    rowWHTSHIP5_ST.Item("SHIP_CITY") = rowSOTORDR5.Item("CUST_CITY") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_STATE") = rowSOTORDR5.Item("CUST_STATE") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_ZIP_CODE") = rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_COUNTRY_CODE") = rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_PHONE") = rowSOTORDR5.Item("CUST_PHONE") & String.Empty
                    rowWHTSHIP5_ST.Item("SHIP_RESIDENTIAL") = "0"
                    rowWHTSHIP5_ST.Item("SHIP_PO_BOX") = "0"
                    If rowWHTSHIP5_ST.Item("SHIP_COUNTRY_CODE") & String.Empty = String.Empty Then rowWHTSHIP5_ST.Item("SHIP_COUNTRY_CODE") = "US"
                    rowWHTSHIP5_ST.AcceptChanges()

                    If optLoadAddress.Value = "S" Then
                        txtLoadAddress.Text = value
                    End If

                    If rowSOTINVH1 IsNot Nothing AndAlso dst.Tables("WHTSHIPC").Select("", "", DataViewRowState.CurrentRows).Length = 0 Then
                        Dim COMMODITY_LNO As Int16 = 1
                        For Each rowSOTINVH2 As DataRow In ASCDATA1.GetDataTable("SELECT * FROM SOTINVH2 WHERE INV_NO = '" & rowSOTINVH1.Item("INV_NO") & "'").Rows
                            Dim rowWHTSHIPC As DataRow = dst.Tables("WHTSHIPC").NewRow
                            rowWHTSHIPC.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                            rowWHTSHIPC.Item("COMMODITY_LNO") = COMMODITY_LNO
                            COMMODITY_LNO += 1

                            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", rowSOTINVH2.Item("ITEM_CODE"))
                            rowWHTSHIPC.Item("COMMODITY_DESC") = rowICTITEM1.Item("ITEM_DESC") & String.Empty
                            rowWHTSHIPC.Item("NUM_PIECES") = Val(rowSOTINVH2.Item("ORDR_QTY_SHIP") & String.Empty)
                            rowWHTSHIPC.Item("MANUFACTURER") = rowICTITEM1.Item("COUNTRY_CODE") & String.Empty
                            If rowWHTSHIPC.Item("MANUFACTURER") & String.Empty = String.Empty Then
                                rowWHTSHIPC.Item("MANUFACTURER") = "US"
                            End If
                            rowWHTSHIPC.Item("HARMONIZED_CODE") = ""
                            rowWHTSHIPC.Item("WEIGHT") = Val(rowICTITEM1.Item("ITEM_WEIGHT") & String.Empty)
                            rowWHTSHIPC.Item("QUANTITY") = Val(rowSOTINVH2.Item("ORDR_QTY_SHIP") & String.Empty)
                            rowWHTSHIPC.Item("QUANTITY_UOM") = "EA"
                            rowWHTSHIPC.Item("UNIT_PRICE") = Val(rowSOTINVH2.Item("ORDR_UNIT_PRICE") & String.Empty)
                            dst.Tables("WHTSHIPC").Rows.Add(rowWHTSHIPC)
                        Next

                        Dim SHIP_VIA_CODE As String = rowSOTINVH1.Item("SHIP_VIA_CODE") & String.Empty
                        Dim rowSOTSVIA1 As DataRow = dst.Tables("SOTSVIA1").Rows.Find(SHIP_VIA_CODE)
                        If rowSOTSVIA1 IsNot Nothing Then
                            Dim CARRIER_CODE As String = rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty
                            If CARRIER_CODE.Length > 0 Then
                                For Each rowProvider As UltraWinGrid.UltraGridRow In cmbProvider.Rows
                                    If rowProvider.Cells("CARRIER_CODE").Value = CARRIER_CODE Then
                                        cmbProvider.SelectedRow = rowProvider
                                        cmbProvider_ValueChanged(Nothing, Nothing)

                                        For Each rowShipMethod As UltraWinGrid.UltraGridRow In cmbShipMethod.Rows
                                            If rowShipMethod.Cells("SHIP_VIA_CODE").Value = SHIP_VIA_CODE Then
                                                cmbShipMethod.SelectedRow = rowShipMethod
                                                Exit For
                                            End If
                                        Next
                                    End If
                                Next
                            End If
                        End If
                    End If
            End Select
        End If

    End Sub

    Private Sub cmbWarehouse_ValueChanged(sender As Object, e As System.EventArgs) Handles cmbWarehouse.ValueChanged
        Dim WHSE_CODE As String = cmbWarehouse.Text

        Dim rowICTWHSE1 As DataRow = dst.Tables("ICTWHSE1").Rows.Find(WHSE_CODE)
        If rowICTWHSE1 IsNot Nothing Then
            txtFromCity.Text = rowICTWHSE1.Item("WHSE_CITY") & String.Empty
            txtFromCompany.Text = rowICTWHSE1.Item("WHSE_DESC") & String.Empty
            txtFromCountry.Text = rowICTWHSE1.Item("WHSE_COUNTRY") & String.Empty
            txtFromCountry.Text = txtFromCountry.Text.Trim
            If txtFromCountry.TextLength = 0 Then txtFromCountry.Text = "US"
            txtFromFirstName.Text = rowICTWHSE1.Item("WHSE_CONTACT") & String.Empty
            txtFromPhone.Text = rowICTWHSE1.Item("WHSE_PHONE") & String.Empty
            txtFromState.Text = rowICTWHSE1.Item("WHSE_STATE") & String.Empty
            txtFromStreet.Text = rowICTWHSE1.Item("WHSE_ADDR1") & String.Empty
            txtFromSuite.Text = rowICTWHSE1.Item("WHSE_ADDR2") & String.Empty
            txtFromZip.Text = rowICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty
            chkFromResidential.Checked = False
        End If
    End Sub

    Private Sub btnFromClear_Click(sender As System.Object, e As System.EventArgs) Handles btnFromClear.Click
        txtFromCity.Clear()
        txtFromCompany.Clear()
        txtFromCountry.Clear()
        txtFromFirstName.Clear()
        txtFromPhone.Clear()
        txtFromState.Clear()
        txtFromStreet.Clear()
        txtFromSuite.Clear()
        txtFromZip.Clear()
        chkFromResidential.Checked = False
    End Sub

    Private Sub btnToClear_Click(sender As System.Object, e As System.EventArgs) Handles btnToClear.Click
        txtToCity.Clear()
        txtToCompany.Clear()
        txtToCountry.Clear()
        txtToFirstName.Clear()
        txtToPhone.Clear()
        txtToState.Clear()
        txtToStreet.Clear()
        txtToSuite.Clear()
        txtToZip.Clear()
        chkToResidential.Checked = False
        chkSignature.Checked = False
        chkToPOBox.Checked = False
    End Sub

    Private Sub btnUseCommodityTotal_Click(sender As System.Object, e As System.EventArgs) Handles btnUseCommodityTotal.Click
        numCustomsValue.Value = Val(dst.Tables("WHTSHIPC").Compute("SUM(EXTENDED_PRICE)", "") & String.Empty)
    End Sub

#End Region

#Region "Setup Ship Request"

    Private Sub LoadSmartPostDropDowns()

        Dim tblSmartPost As New DataTable
        tblSmartPost.Columns.Add("SMART_POST_TYPE", GetType(System.String))
        tblSmartPost.Columns.Add("SMART_POST_DESC", GetType(System.String))
        tblSmartPost.Rows.Add(New Object() {"0", "Media Mail  (1 to 70 lbs)"})
        tblSmartPost.Rows.Add(New Object() {"1", "Parcel Select (1 to 70 lbs)"})
        tblSmartPost.Rows.Add(New Object() {"2", "Presorted Bound (0.1 to 15 lbs)"})
        tblSmartPost.Rows.Add(New Object() {"3", "Presorted Std (up to 1 lb)"})
        cmbSmartPost.DataSource = tblSmartPost

        Dim tblSmartPostPkg As New DataTable
        tblSmartPostPkg.Columns.Add("SMART_POST_PKG", GetType(System.String))
        tblSmartPostPkg.Columns.Add("SMART_POST_PKG_DESC", GetType(System.String))
        tblSmartPostPkg.Rows.Add(New Object() {"0", "Other"})
        tblSmartPostPkg.Rows.Add(New Object() {"1", "Bag"})
        tblSmartPostPkg.Rows.Add(New Object() {"2", "Barrel"})
        tblSmartPostPkg.Rows.Add(New Object() {"3", "Basket"})
        tblSmartPostPkg.Rows.Add(New Object() {"4", "Box"})
        tblSmartPostPkg.Rows.Add(New Object() {"5", "Bucket"})
        tblSmartPostPkg.Rows.Add(New Object() {"6", "Bundle"})
        tblSmartPostPkg.Rows.Add(New Object() {"7", "Carton"})
        tblSmartPostPkg.Rows.Add(New Object() {"8", "Case"})
        tblSmartPostPkg.Rows.Add(New Object() {"9", "Container"})
        tblSmartPostPkg.Rows.Add(New Object() {"10", "Crate"})
        tblSmartPostPkg.Rows.Add(New Object() {"11", "Cylinder"})
        tblSmartPostPkg.Rows.Add(New Object() {"12", "Drum"})
        tblSmartPostPkg.Rows.Add(New Object() {"13", "Envelope"})
        tblSmartPostPkg.Rows.Add(New Object() {"14", "Hamper"})
        tblSmartPostPkg.Rows.Add(New Object() {"15", "Pail"})
        tblSmartPostPkg.Rows.Add(New Object() {"16", "Pallet"})
        tblSmartPostPkg.Rows.Add(New Object() {"17", "Piece"})
        tblSmartPostPkg.Rows.Add(New Object() {"18", "Reel"})
        tblSmartPostPkg.Rows.Add(New Object() {"19", "Roll"})
        tblSmartPostPkg.Rows.Add(New Object() {"20", "Skid"})
        tblSmartPostPkg.Rows.Add(New Object() {"21", "Tank"})
        tblSmartPostPkg.Rows.Add(New Object() {"22", "Tube"})
        cmbSmartPostPackage.DataSource = tblSmartPostPkg
    End Sub

     Private Function RequestShippingLabel() As Boolean

        Try
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Requesting Shipping label")

            RequestShippingLabel = False

            If EntryMode = "S" Then
                SHIP_CNTL_NO = ASCMAIN1.Next_Control_No("WHTSHIP1.SHIP_CNTL_NO")
                For Each tableName As String In New String() {"WHTSHIP1", "WHTSHIP2", "WHTSHIP5_SF", "WHTSHIP5_ST", "WHTSHIP5_HL", "WHTSHIPC", "WHTSHIPP_S", "WHTSHIPP_D"}
                    For Each row As DataRow In dst.Tables(tableName).Select("", "", DataViewRowState.CurrentRows)
                        row.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                    Next
                Next
            End If

            Dim rowWHTSHIP1 As DataRow = dst.Tables("WHTSHIP1").Rows(0)
            rowWHTSHIP1.Item("CARRIER_PROD_CODE") = cmbShipMethod.SelectedRow.Cells("CARRIER_PROD_CODE").Value
            rowWHTSHIP1.Item("SHIP_VIA_CODE") = cmbShipMethod.SelectedRow.Cells("SHIP_VIA_CODE").Value

            GetSenderInfo()
            GetRecipientInfo()
            BuildPackages()
            GetCredentials()
            GetServiceType()
            GetDropoffType()
            GetShipPayor()
            GetDutiesPayor()
            GetSpecialServices()
            GetHALDetails()
            GetCommodities()
            GetSmartPost()

            With clsShip
                .EzshipLabelImage = DPayments.DShippingSDK.EzShipLabelImageTypes.itEltron
                .ShippingLabelDirectory = ShippingLabelDirectory
                .ShippingLabelPrefix = SHIP_CNTL_NO
                .ShipDate = dteShipDate.DateTime
                clsShip.TotalCustomsValue = numCustomsValue.Value
            End With

            clsShip.PackageDetailList = shipPackageDetailList

            Select Case EntryMode
                Case "S"
                    ' Request label
                    If clsShip.RequestLabel() Then

                        For Each shipPackageDetail As DPayments.DShippingSDK.PackageDetail In shipPackageDetailList
                            Dim SHIP_PACKAGE_NO As String = Val(shipPackageDetail.Id)
                            If dst.Tables("WHTSHIP2").Select("SHIP_PACKAGE_NO = " & SHIP_PACKAGE_NO, "").Length > 0 Then
                                Dim rowWHTSHIP2 As DataRow = dst.Tables("WHTSHIP2").Select("SHIP_PACKAGE_NO = " & SHIP_PACKAGE_NO)(0)
                                rowWHTSHIP2.Item("TRACKING_NO") = shipPackageDetail.TrackingNumber & String.Empty
                                rowWHTSHIP2.Item("BASE_CHARGE") = Val(clsShip.ShipmentBaseCharge(SHIP_PACKAGE_NO) & String.Empty)
                                rowWHTSHIP2.Item("NET_CHARGE") = Val(clsShip.ShipmentNetCharge(SHIP_PACKAGE_NO) & String.Empty)
                                rowWHTSHIP2.Item("TOTAL_DISCOUNT") = Val(clsShip.ShipmentDiscountCharge(SHIP_PACKAGE_NO) & String.Empty)
                                rowWHTSHIP2.Item("TOTAL_SURCHARGES") = Val(clsShip.ShipmentSurCharge(SHIP_PACKAGE_NO) & String.Empty)
                                rowWHTSHIP2.Item("LIST_PRICE") = Val(clsShip.ShipmentListCharge(SHIP_PACKAGE_NO) & String.Empty)
                            End If

                            rowWHTSHIP1.Item("ERROR_MSG") = clsShip.LastError & String.Empty
                            If rowWHTSHIP1 IsNot Nothing AndAlso (rowWHTSHIP1.Item("ERROR_MSG") & String.Empty).ToString.Length > 200 Then
                                rowWHTSHIP1.Item("ERROR_MSG") = rowWHTSHIP1("ERROR_MSG").ToString.Substring(0, 200).Trim
                            End If
                            rowWHTSHIP1.Item("MASTER_TRACKING_NO") = clsShip.MasterTrackingNumber & String.Empty

                            If rowWHTSHIP1.Item("MASTER_TRACKING_NO") & String.Empty = String.Empty Then
                                rowWHTSHIP1.Item("MASTER_TRACKING_NO") = shipPackageDetailList(0).TrackingNumber
                            End If

                            RequestShippingLabel = True
                            Update_Record()

                            If shipPackageDetail.ShippingLabel.Length > 0 Then PrintShipingLabel(shipPackageDetail.ShippingLabel)
                            If shipPackageDetail.CODLabel.Length > 0 Then PrintShipingLabel(shipPackageDetail.CODLabel)
                            If shipPackageDetail.ReturnReceipt.Length > 0 Then PrintShipingLabel(shipPackageDetail.ReturnReceipt)

                        Next
                    Else
                        MessageBox.Show("Shipping Label(s) could not be captured. " & clsShip.LastError)
                    End If

                Case "R"
                    'clsShip.RequestShipmentRates()

                    Dim msg As String = String.Empty
                    'For Each reqSer As ServiceDetail In clsShip.RequestedServicesRates
                    '    msg &= vbCrLf & reqSer.ServiceType.ToString.PadRight(30, "_") _
                    '        & (reqSer.ServiceType & " ").ToString.PadRight(5) _
                    '        & Val(reqSer.AccountNetCharge & String.Empty).ToString("#,##0.00").ToString.PadLeft(8) & reqSer.TransitTime.PadLeft(12)
                    ' Next

                    MessageBox.Show("Available Delivery Options" & vbCrLf & msg, "Requested Rates", MessageBoxButtons.OK, MessageBoxIcon.Information)

            End Select

        Catch ex As Exception
            MessageBox.Show("Error requesting shipping label: " & ex.Message, "Request Label Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If rowWHTSHIP1 IsNot Nothing Then rowWHTSHIP1.Item("ERROR_MSG") = ex.Message
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try
    End Function

    Private Sub ReprintShippingLabel()

        Try
            Dim CARRIER_CODE As String = cmbProvider.SelectedRow.Cells("CARRIER_CODE").Value & String.Empty
            Dim rowSOTCARR1 As DataRow = dst.Tables("SOTCARR1").Select("CARRIER_CODE = '" & CARRIER_CODE & "'")(0)
            Dim rowSOTCARR3 As DataRow = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "'")(0)
            ShippingLabelDirectory = (rowSOTCARR1.Item("CARRIER_ARCHIVE_DIR") & String.Empty).ToString.Trim

            Dim labelFilesFound As Int16 = 0

            If ASCMAIN1.Running_in_VS Then
                'ShippingLabelDirectory = ShippingLabelDirectory.Replace("S:\", "N:\")
                ShippingLabelDirectory = ShippingLabelDirectory.Replace(ASCMAIN1.Folders("SharedRoot"), "N:\")
            End If

            For Each Label As String In My.Computer.FileSystem.GetFiles(ShippingLabelDirectory, FileIO.SearchOption.SearchTopLevelOnly, SHIP_CNTL_NO & "*.*")
                Dim ShippingLabel As String = String.Empty

                Using sr As New IO.StreamReader(Label)
                    ShippingLabel = sr.ReadToEnd
                    sr.Close()
                End Using

                PrintShipingLabel(ShippingLabel)
                labelFilesFound += 1
            Next

            MessageBox.Show("There were (" & labelFilesFound & ") labels sent to the printer.", "Reprint Label", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub GetSmartPost()
        If cmbProvider.SelectedRow.Cells("PROVIDER_TYPE").Value = WHCSHIP1.ProviderTypeFedex _
        AndAlso cmbShipMethod.SelectedRow.Cells("CARRIER_PROD_CODE").Value = FedexSmartPost Then
            clsShip.FedexSmartPost.Indicia = cmbSmartPost.SelectedRow.Cells("SMART_POST_TYPE").Value
            clsShip.FedexSmartPost.PhysicalPackaging = cmbSmartPostPackage.SelectedRow.Cells("SMART_POST_PKG").Value
            clsShip.FedexSmartPost.HubId = cmbProvider.SelectedRow.Cells("FEDEX_HUB_ID").Value
        Else
            dst.Tables("WHTSHIP1").Rows(0).Item("SMART_POST_TYPE") = DBNull.Value
            dst.Tables("WHTSHIP1").Rows(0).Item("SMART_POST_PKG") = DBNull.Value
            dst.Tables("WHTSHIP1").Rows(0).Item("SMART_POST_HUB_ID") = DBNull.Value
        End If
    End Sub

    Private Sub GetSenderInfo()

        With clsShip.Sender
            .FirstName = txtFromFirstName.Text.Trim
            .MiddleInitial = ""
            .LastName = "" 'txtFromLastName.Text.Trim
            .Address1 = txtFromStreet.Text.Trim
            .Address2 = txtFromSuite.Text.Trim
            .City = txtFromCity.Text.Trim
            .State = txtFromState.Text.Trim.ToUpper
            .ZipCode = txtFromZip.Text.Trim
            .CountryCode = txtFromCountry.Text.Trim
            .Company = txtFromCompany.Text.Trim
            .Phone = txtFromPhone.Text

            .IsResidental = chkFromResidential.Checked
            .IsPOBox = False

        End With
    End Sub

    Private Sub GetRecipientInfo()

        With clsShip.Recipient
            .FirstName = txtToFirstName.Text.Trim
            .MiddleInitial = ""
            .LastName = "" 'txtToLastName.Text.Trim
            .Address1 = txtToStreet.Text.Trim
            .Address2 = txtToSuite.Text.Trim
            .City = txtToCity.Text.Trim
            .State = txtToState.Text.Trim
            .ZipCode = txtToZip.Text.Trim
            .CountryCode = txtToCountry.Text.Trim
            .Company = txtToCompany.Text.Trim
            .Phone = txtToPhone.Text

            .IsResidental = chkToResidential.Checked
            .IsPOBox = chkToPOBox.Checked

            If .Company.Length = 0 Then
                .Company = (.FirstName & " " & .LastName).ToString.Trim
            End If

            clsShip.SignatureRequired = chkSignature.Checked

        End With
    End Sub

    Private Sub BuildPackages()

        shipPackageDetailList.Clear()
        Dim idCtr As Int16 = 1
        For Each row As DataRow In dst.Tables("WHTSHIP2").Rows
            Dim shipPackageDetail As New DPayments.DShippingSDK.PackageDetail
            With shipPackageDetail
                .PackagingType = DPayments.DShippingSDK.TPackagingTypes.ptYourPackaging
                .Weight = Convert.ToInt32(row.Item("WEIGHT"))
                .Length = Convert.ToInt32(row.Item("LENGTH"))
                .Width = Convert.ToInt32(row.Item("WIDTH"))
                .Height = Convert.ToInt32(row.Item("HEIGHT"))

                ' Onlt three references per package
                Dim reference As String = String.Empty
                Dim value As String = String.Empty
                Dim iCtr As Int16 = 1
                For Each field As String In New String() {"CR:CUST_REF", "IN:INV_NO", "PO:PO_ORDER_NO", "DN:DEPT_NO"} ' , "BL:INV_BOL_NO"
                    Dim code As String = field.Split(":")(0)
                    Dim fieldName As String = field.Split(":")(1)

                    value = (row.Item(fieldName) & String.Empty).ToString.Trim
                    If value.Length > 0 Then
                        reference &= "; " & code & ":" & value
                        iCtr += 1
                    End If
                    If iCtr > 3 Then Exit For
                Next

                If reference.StartsWith(";") Then
                    reference = reference.Substring(1).Trim
                End If

                .Reference = reference
                .Id = idCtr.ToString.Trim.PadLeft(8, "0")

                If Val(row.Item("INSURED_VALUE") & String.Empty) > 0 Then
                    .InsuredValue = Val(row.Item("INSURED_VALUE") & String.Empty)
                End If

                'COD_TYPE, COD_AMOUNT
                If Val(row.Item("COD_AMOUNT") & String.Empty) > 0 Then
                    .CODAmount = Val(row.Item("COD_AMOUNT") & String.Empty)
                    .CODType = Val(row.Item("COD_TYPE") & String.Empty)
                End If

                idCtr += 1
            End With

            shipPackageDetailList.Add(shipPackageDetail)
        Next

    End Sub

    Private Sub GetCredentials()

        Dim CARRIER_CODE As String = cmbProvider.SelectedRow.Cells("CARRIER_CODE").Value & String.Empty
        Dim rowSOTCARR1 As DataRow = dst.Tables("SOTCARR1").Select("CARRIER_CODE = '" & CARRIER_CODE & "'")(0)
        Dim rowSOTCARR3 As DataRow = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "'")(0)
        ShippingLabelDirectory = (rowSOTCARR1.Item("CARRIER_ARCHIVE_DIR") & String.Empty).ToString.Trim

        Try
            If ASCMAIN1.Running_in_VS Then
                'ShippingLabelDirectory = ShippingLabelDirectory.Replace("S:\", "C:\")
                ShippingLabelDirectory = ShippingLabelDirectory.Replace(ASCMAIN1.Folders("SharedRoot"), "C:\")
            End If
            If ShippingLabelDirectory.Length > 0 Then
                If Not My.Computer.FileSystem.DirectoryExists(ShippingLabelDirectory) Then
                    My.Computer.FileSystem.CreateDirectory(ShippingLabelDirectory)
                End If
            End If
        Catch ex As Exception
            ShippingLabelDirectory = String.Empty
        End Try

        If ShippingLabelDirectory.Length > 0 AndAlso Not ShippingLabelDirectory.EndsWith("\") Then
            ShippingLabelDirectory = ShippingLabelDirectory & "\"
        End If

        ' Credentials
        clsShip.Server = rowSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
        clsShip.UserId = rowSOTCARR3.Item("SHIPPER_ID") & String.Empty
        clsShip.Password = rowSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
        clsShip.AccountNumber = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
        clsShip.UPSAccessKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
        clsShip.FedexMeterNumber = rowSOTCARR3.Item("METER_NUMBER") & String.Empty
        clsShip.FedexDeveloperKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty

        rowWHTSHIP1.Item("CARRIER_ACCOUNT_NO") = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
        rowWHTSHIP1.Item("CARRIER_CODE") = CARRIER_CODE
        clsShip.LabelStockType = (rowSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim

    End Sub

    Private Sub GetServiceType()
        'Service Type
        If isInternationalShipment Then
            clsShip.RequestedServiceType = Val(cmbShipMethod.SelectedRow.Cells("CARRIER_PROD_CODE_INTL").Value & String.Empty)
        Else
            clsShip.RequestedServiceType = Val(cmbShipMethod.SelectedRow.Cells("CARRIER_PROD_CODE").Value & String.Empty)
        End If
    End Sub

    Private Sub GetDropoffType()
        'clsShip.DropOffType = FedexshipintlDropoffTypes.dtRegularPickup
    End Sub

    Private Sub GetShipPayor()
        Select Case cmbShipPayor.Value
            Case "S" : clsShip.Payor = DPayments.DShippingSDK.TPayorTypes.ptSender
            Case "R" : clsShip.Payor = DPayments.DShippingSDK.TPayorTypes.ptRecipient
            Case "T" : clsShip.Payor = DPayments.DShippingSDK.TPayorTypes.ptThirdParty
            Case "C" : clsShip.Payor = DPayments.DShippingSDK.TPayorTypes.ptCollect
        End Select
        clsShip.PayorContact.AccountNumber = txtShipAccountNo.Text
        clsShip.PayorContact.CountryCode = txtShipCountry.Text
    End Sub

    Private Sub GetDutiesPayor()
        Select Case cmbShipPayor.Value
            Case "S" : clsShip.DutiesPayor = DPayments.DShippingSDK.TPayorTypes.ptSender
            Case "R" : clsShip.DutiesPayor = DPayments.DShippingSDK.TPayorTypes.ptRecipient
            Case "T" : clsShip.DutiesPayor = DPayments.DShippingSDK.TPayorTypes.ptThirdParty
            Case "C" : clsShip.DutiesPayor = DPayments.DShippingSDK.TPayorTypes.ptCollect
        End Select
        clsShip.DutiesPayorContact.AccountNumber = txtDutiesAccountNo.Text
        clsShip.DutiesPayorContact.CountryCode = txtDutiesCountry.Text
    End Sub

    Private Sub GetSpecialServices()
        For Each row As DataRow In dst.Tables("SHTSHIPS").Select("SEL = '1'")
            clsShip.ShipmentSpecialServices = clsShip.ShipmentSpecialServices Or Val("&H" & row.Item("SPCL_SHIP_CODE") & "L")
        Next
    End Sub

    Private Sub GetCommodities()
        ' Only used for International
        clsShip.CommodityDetailList.Clear()
        For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdWHTSHIPC.Rows
            Dim CommodityDetail As New DPayments.DShippingSDK.CommodityDetail
            CommodityDetail.Description = grdRow.Cells("COMMODITY_DESC").Value & String.Empty
            CommodityDetail.NumberOfPieces = Val(grdRow.Cells("NUM_PIECES").Value & String.Empty)
            CommodityDetail.Quantity = Val(grdRow.Cells("QUANTITY").Value & String.Empty)
            CommodityDetail.QuantityUnit = grdRow.Cells("QUANTITY_UOM").Value & String.Empty
            CommodityDetail.UnitPrice = grdRow.Cells("UNIT_PRICE").Value & String.Empty
            CommodityDetail.Weight = Val(grdRow.Cells("WEIGHT").Value & String.Empty) ' Leave as pounds
            CommodityDetail.Manufacturer = grdRow.Cells("MANUFACTURER").Value & String.Empty
            clsShip.CommodityDetailList.Add(CommodityDetail)
        Next
    End Sub

    Private Sub GetHALDetails()

        With clsShip.HoldAtLocation
            .AccountNumber = ""
            .Address1 = txtHoldAddress1.Text.Trim
            .Address2 = txtHoldAddress2.Text.Trim
            .City = txtHoldCity.Text.Trim
            .Company = txtHoldCompany.Text.Trim
            .CountryCode = txtHoldCountry.Text.Trim
            .eMail = ""
            .FirstName = txtHoldContact.Text.Trim
            .Fax = ""
            .IsPOBox = False
            .IsResidental = False
            .LastName = ""
            .MiddleInitial = ""
            .Phone = txtHoldPhone.Text.Trim
            .State = txtHoldState.Text.Trim
            .ZipCode = txtHoldZip.Text.Trim
        End With
    End Sub

    Private Sub ValidatePacking(ByRef grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow)

        Dim SHIP_PACKAGE_NO As String = (grdRow.Cells("SHIP_PACKAGE_NO").Value & String.Empty).ToString.Trim
        If SHIP_PACKAGE_NO = String.Empty Then
            SHIP_PACKAGE_NO = Val(dst.Tables("WHTSHIP2").Compute("MAX(SHIP_PACKAGE_NO)", "") & String.Empty) + 1
            grdRow.Cells("SHIP_PACKAGE_NO").Value = SHIP_PACKAGE_NO
        End If

        grdRow.Cells("SHIP_CNTL_NO").Value = SHIP_CNTL_NO

        Dim LENGTH As Int16 = Val(grdRow.Cells("LENGTH").Value & String.Empty)
        Dim WIDTH As Int16 = Val(grdRow.Cells("WIDTH").Value & String.Empty)
        Dim HEIGHT As Int16 = Val(grdRow.Cells("HEIGHT").Value & String.Empty)

        If (grdRow.Cells("PACKAGING_TYPE").Value & String.Empty).ToString.Trim.Length = 0 Then
            EMsg &= vbCrLf & "Package type is required"
        ElseIf Val(grdRow.Cells("PACKAGING_TYPE").Value & String.Empty) <> DPayments.DShippingSDK.TPackagingTypes.ptYourPackaging Then
            If LENGTH <= 0 OrElse WIDTH <= 0 OrElse HEIGHT <= 0 Then
                EMsg &= vbCrLf & "Package Length, Width and Height are required and must be greater than 0"
            End If
        End If

        Dim POUNDS As Int16 = Val(grdRow.Cells("POUNDS").Value & String.Empty)
        Dim OUNCES As Int16 = Val(grdRow.Cells("OUNCES").Value & String.Empty)

        If POUNDS <= 0 AndAlso OUNCES <= 0 Then
            EMsg &= vbCrLf & "Pounds and/or Ounces must be greater than 0"
        ElseIf POUNDS < 0 Then
            EMsg &= vbCrLf & "Pounds must be greater equal than 0"
        ElseIf OUNCES < 0 Then
            EMsg &= vbCrLf & "Ounces must be greater equal than 0"
        Else
            grdRow.Cells("WEIGHT").Value = POUNDS * 16 + OUNCES
        End If

        Dim COD_AMOUNT As Decimal = Val(grdRow.Cells("COD_AMOUNT").Value & String.Empty)
        Dim COD_TYPE As String = grdRow.Cells("COD_TYPE").Value & String.Empty
        If COD_AMOUNT < 0 Then
            EMsg &= vbCrLf & "COD Amount must be greater equal than 0"
        ElseIf COD_AMOUNT > 0 Then
            If COD_TYPE <> "0" AndAlso COD_TYPE <> "1" Then
                EMsg &= vbCrLf & "COD Type is required when providing a COD Amount"
            End If
        ElseIf COD_TYPE = "0" AndAlso COD_TYPE = "1" Then
            If COD_AMOUNT = 0 Then
                EMsg &= vbCrLf & "COD Amount is required when providing a COD Type"
            End If
        End If

        If COD_AMOUNT = 0 And COD_TYPE = String.Empty Then
            grdRow.Cells("COD_TYPE").Value = "2" ' None
        End If
    End Sub

    Private Sub ValidateCommodity(ByRef grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow)

        Dim COMMODITY_LNO As String = (grdRow.Cells("COMMODITY_LNO").Value & String.Empty).ToString.Trim
        If COMMODITY_LNO = String.Empty Then
            COMMODITY_LNO = Val(dst.Tables("WHTSHIPC").Compute("MAX(COMMODITY_LNO)", "") & String.Empty) + 1
            grdRow.Cells("COMMODITY_LNO").Value = COMMODITY_LNO
        End If

        grdRow.Cells("SHIP_CNTL_NO").Value = SHIP_CNTL_NO

        Dim COMMODITY_DESC As String = (grdRow.Cells("COMMODITY_DESC").Value & String.Empty).ToString.Trim
        Dim MANUFACTURER As String = (grdRow.Cells("MANUFACTURER").Value & String.Empty).ToString.Trim
        Dim NUM_PIECES As Int16 = Val(grdRow.Cells("NUM_PIECES").Value & String.Empty)
        Dim UOM As String = Val(grdRow.Cells("QUANTITY_UOM").Value & String.Empty)
        Dim UNIT_PRICE As Decimal = Val(grdRow.Cells("UNIT_PRICE").Value & String.Empty)
        Dim QUANTITY As Decimal = Val(grdRow.Cells("QUANTITY").Value & String.Empty)
        Dim WEIGHT As Decimal = Val(grdRow.Cells("WEIGHT").Value & String.Empty)

        ' Required Fields
        If COMMODITY_DESC.Length = 0 Then
            EMsg &= vbCrLf & "Description is required"
        End If

        If MANUFACTURER.Length = 0 Then
            EMsg &= vbCrLf & "Manufacturer is required"
        End If

        If QUANTITY < 0 Then
            EMsg &= vbCrLf & "Quantity must be greater equal 0"
        End If

        If UOM.Length = 0 Then
            EMsg &= vbCrLf & "Unit of Measure is required"
        End If

        If UNIT_PRICE < 0 Then
            EMsg &= vbCrLf & "Unit Price must be greater equal 0"
        End If

        If WEIGHT < 0 Then
            EMsg &= vbCrLf & "Weight must be greater 0"
        End If

    End Sub

    Private Function GetCommodityUnits() As String()
        Return New String() {":", _
              "AR:Carat" _
             , "CG:Centigram" _
             , "CM:Centimeters" _
             , "CM3:Cubic centimeters" _
             , "CFT:Cubic feet" _
             , "M3:Cubic meters" _
             , "DOZ:Dozen" _
             , "DPR:Dozen pair" _
             , "EA:Each" _
             , "GAL:Gallon" _
             , "G:Grams" _
             , "GR:Gross" _
             , "KG:Kilograms" _
             , "LFT:Linear foot" _
             , "LNM:Linear meters" _
             , "LYD:Linear yard" _
             , "LTR:Liters" _
             , "M:Meters" _
             , "MG:Milligram" _
             , "ML:Milliliter" _
             , "NO:Number" _
             , "OZ:Ounces" _
             , "PRS:Pairs" _
             , "PCS:Pieces" _
             , "LB:Pound" _
             , "CM2:Square centimeters" _
             , "SFT:Square feet" _
             , "SQI:Square inches" _
             , "M2:Square meters" _
             , "SYD:Square yards" _
             , "YD:Yard"}

    End Function


    Private Function CancelShipment() As Boolean
        Try
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Requesting Shipment Cancellation")

            GetCredentials()
            GetServiceType()

            Dim rowWHTSHIP1 As DataRow = dst.Tables("WHTSHIP1").Rows(0)
            'Dim multiShipment As Boolean = dst.Tables("WHTSHIP2").Select("ISNULL(TRACKING_NO, '') <> ''").Length > 1
            Dim rowSOTCARR2 As DataRow = ASCDATA1.GetDataRow("select * from sotcarr2 where CARRIER_CODE = :PARM1" _
                                                             & " and CARRIER_PROD_CODE = :PARM2", "VV", _
                                                             New Object() {rowWHTSHIP1.Item("CARRIER_CODE"), rowWHTSHIP1.Item("CARRIER_PROD_CODE")})

            ' cancel one package at a time
            Dim numCancelled As Int16 = 0
            Dim numNotCancelled As Int16 = 0
            With clsShip
                For Each rowWHTSHIP2 As DataRow In dst.Tables("WHTSHIP2").Select("SEL = '1'")
                    If .CancelShipment(rowWHTSHIP2.Item("TRACKING_NO"), False, Val(rowSOTCARR2.Item("TRACKING_ID_TYPE") & String.Empty)) Then
                        Try
                            BeginTrans()
                            With rowWHTSHIP2
                                .Item("STATUS") = "C"
                                .Item("CANCEL_OPER") = ASCMAIN1.USER_ID
                                .Item("CANCEL_DATE") = DateTime.Now
                                Update_Record_TDA("WHTSHIP2")
                            End With

                            ASCMAIN1.sql = "Update WHTSHIP1 SET STATUS = (SELECT MAX(STATUS) FROM WHTSHIP2 WHERE WHTSHIP1.SHIP_CNTL_NO = WHTSHIP2.SHIP_CNTL_NO)"
                            ASCMAIN1.sql &= " where WHTSHIP1.SHIP_CNTL_NO = '" & rowWHTSHIP1.Item("SHIP_CNTL_NO") & "'"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                            CommitTrans()
                        Catch ex As Exception
                            Rollback()
                        End Try
                        numCancelled += 1
                     Else
                        numNotCancelled += 1
                        If .LastError.Length > 0 Then

                        End If
                    End If
                Next
            End With

            Dim zMsg As String = String.Empty
            zMsg &= "There were " & numCancelled & " packages cancelled"
            zMsg &= Environment.NewLine & "There were " & numNotCancelled & " packages NOT cancelled"
            MessageBox.Show(zMsg, "Cancel Package", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return True

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Cancel Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False

        Finally
            ASCMAIN1.Progress(String.Empty, String.Empty)
            Me.Cursor = Cursors.Default
        End Try
    End Function

    Private Sub SetReadOnly(ByVal readOnlyValue As Boolean)
        Set_Read_Only(grpHeader, readOnlyValue)
        Set_Read_Only(grpShipFrom, readOnlyValue)
        Set_Read_Only(grpShipTo, readOnlyValue)
        Set_Read_Only(grpShipPayor, readOnlyValue)
        Set_Read_Only(grpDutiesPayor, readOnlyValue)
        Set_Read_Only(grpHoldAtLocation, readOnlyValue)
        Set_Read_Only(grpCommodity, readOnlyValue)

        'Set_Read_Only(grdWHTSHIP2, readOnlyValue)
        'Set_Read_Only(grdSHTSHIPS, readOnlyValue)
        'Set_Read_Only(grdWHTSHIPC, readOnlyValue)

        If readOnlyValue Then
            grdWHTSHIP2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grdWHTSHIP2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            grdWHTSHIP2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

            grdSHTSHIPS.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

            grdWHTSHIPC.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grdWHTSHIPC.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            grdWHTSHIPC.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        Else
            grdWHTSHIP2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            grdWHTSHIP2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            grdWHTSHIP2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

            grdSHTSHIPS.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

            grdWHTSHIPC.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            grdWHTSHIPC.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            grdWHTSHIPC.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        End If

        btnFromClear.Visible = Not readOnlyValue
        btnToClear.Visible = Not readOnlyValue
    End Sub

    Public Function PrintShipingLabel(ByVal LabelData As String) As Boolean

        Try
            If ASCMAIN1.USER_ID = "edz" AndAlso ASCMAIN1.Running_in_VS Then
                ' Find Zebra printer
                Dim zebraPrinter As String = FindZebraPrinter()

                Dim vLabelPrinter As New ASCPRINT
                Return vLabelPrinter.SendStringToPrinter(zebraPrinter, LabelData)
            End If

             ASCMAIN1.LabelPrinterSerialPort.WriteLine(LabelData)

        Catch ex As Exception
            MessageBox.Show("Print Shipping Label Error: " & ex.Message)

        End Try

    End Function

    Private Shared Function FindZebraPrinter() As String
        For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
            If printerName.ToUpper.StartsWith("ZEBRA450") Then
                Return printerName
            End If
        Next printerName

        For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
            If printerName.ToUpper.StartsWith("ZEBRA") Then
                Return printerName
            End If
        Next printerName

        Return ""
    End Function

#End Region

#Region "Serial and Com Connections"

    ''' <summary>
    ''' Sets the Printer Settings
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetUpPortsAndPrinters()

        Dim tooltip As New System.Windows.Forms.ToolTip()

        '' Laser Printer
        'Try

        '    txtLaserPrinter.Text = ASCMAIN1.LaserPrinterIpAddress
        '    tooltip.SetToolTip(txtLaserPrinter, ASCMAIN1.LaserPrinterIpAddress)
        '    If ASCMAIN1.LaserPrinterIpAddress.Length = 0 Then
        '        txtLaserPrinter.Appearance.BackColor = Drawing.Color.Red
        '    Else
        '        txtLaserPrinter.Appearance.BackColor = Drawing.Color.Yellow
        '        If Net.IPAddress.TryParse(ASCMAIN1.LaserPrinterIpAddress, Nothing) Then
        '            txtLaserPrinter.Appearance.BackColor = Drawing.Color.Green
        '        End If
        '    End If

        'Catch ex As Exception
        '    txtLaserPrinter.Appearance.BackColor = Drawing.Color.Yellow
        'End Try

        ' Label Printer Port
        Try
            txtLabelPrinter.BackColor = Drawing.Color.Red

            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing Then
                txtlabelPrinter.Text = ASCMAIN1.LabelPrinterSerialPort.PortName
                tooltip.SetToolTip(txtlabelPrinter, txtlabelPrinter.Text)
            Else
                Me.txtLabelPrinter.Text = "No Port"
                tooltip.SetToolTip(txtlabelPrinter, txtlabelPrinter.Text)
            End If

            txtLabelPrinter.BackColor = Drawing.Color.Yellow
            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing AndAlso Not ASCMAIN1.LabelPrinterSerialPort.IsOpen Then
                ASCMAIN1.LabelPrinterSerialPort.Open()
            End If

            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing AndAlso ASCMAIN1.LabelPrinterSerialPort.IsOpen Then
                txtlabelPrinter.BackColor = Drawing.Color.Green
                printerFound = True
            End If

        Catch ex As Exception
            txtLabelPrinter.BackColor = Drawing.Color.Red
        End Try

        '**************************
        '**    Scale Port
        '************************** 
        'ASCMAIN1.ScaleWeightDelegate = AddressOf ProcessScaleData

        'Try
        '    txtScale.BackColor = Drawing.Color.Red

        '    If ASCMAIN1.ScaleSerialPort IsNot Nothing Then
        '        txtScale.Text = ASCMAIN1.ScaleSerialPort.PortName
        '        tooltip.SetToolTip(txtScale, txtScale.Text)
        '    Else
        '        txtScale.Text = "No Port"
        '        tooltip.SetToolTip(txtScale, txtScale.Text)
        '    End If

        '    txtScale.BackColor = Drawing.Color.Yellow
        '    If ASCMAIN1.ScaleSerialPort IsNot Nothing AndAlso Not ASCMAIN1.ScaleSerialPort.IsOpen Then
        '        ASCMAIN1.ScaleSerialPort.Open()
        '    End If

        '    If ASCMAIN1.ScaleSerialPort IsNot Nothing Then
        '        txtScale.BackColor = Drawing.Color.Green
        '    End If

        'Catch ex As Exception
        '    txtScale.BackColor = Drawing.Color.Red
        'End Try

    End Sub

#End Region

    Private Sub grdWHTSHIP2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTSHIP2.ClickCellButton
        If e.Cell.Column.Key = "POUNDS" Then
            registeredWeight = 0
            RequestWeightFromScale()
            e.Cell.Value = Convert.ToInt16(registeredWeight)
            e.Cell.Row.Cells("OUNCES").Value = (registeredWeight * 16) Mod 16
        End If
    End Sub

    ''' <summary>
    ''' Request weight from scale
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub RequestWeightFromScale()

        Try
            registeredWeight = 0

            'If ASCMAIN1.ScaleSerialPort Is Nothing Then Exit Sub

            'If Not ASCMAIN1.ScaleSerialPort.IsOpen Then
            '    ASCMAIN1.ScaleSerialPort.Open()
            'End If

            '' Request the weight from the scale
            'Dim encoding As New System.Text.UTF8Encoding()
            'Dim inBuffer As Byte() = encoding.GetBytes("W")
            ''ASCMAIN1.ScaleSerialPort.Write(inBuffer, 0, inBuffer.Length)
            'ASCMAIN1.ScaleSerialPort.Write("SGW" & vbCrLf)

        Catch ex As Exception
            MessageBox.Show("Scale Weight Error: " & ex.Message, "Scale Weight", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Fires when the weight is requested from the scale
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ProcessScaleData(ByVal scaledata As String)

        If MdiParent.ActiveMdiChild Is Nothing Then Exit Sub
        If MdiParent.ActiveMdiChild.Name <> Me.Name Then Exit Sub

        Try
            'Dim length As Int16 = ASCMAIN1.ScaleSerialPort.BytesToRead
            'If length > 0 Then
            '    Dim numberOfBytesRead As Int16 = 0
            '    Dim readBuffer(length) As Byte
            '    numberOfBytesRead = ASCMAIN1.ScaleSerialPort.Read(readBuffer, 0, length)
            '    MessageBox.Show(readBuffer.ToString)
            '    registeredWeight = Val(readBuffer)
            'End If
        Catch ex As Exception

        End Try
    End Sub

End Class