
Public Class TAFSHIP1

    Private clsShip As New TAC.SCHSHIP1
    Private tblSOTCARR1 As DataTable
    Private tblSOTCARR3 As DataTable

    Private Sub TAFSHIP1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        tblSOTCARR1 = ASCDATA1.GetDataTable("Select * from SOTCARR1")
        tblSOTCARR3 = ASCDATA1.GetDataTable("Select * from SOTCARR3")

        For Each row As DataRow In tblSOTCARR1.Select("", "CARRIER_DESC")
            cmbProvider.Items.Add(row.Item("CARRIER_DESC"))
        Next

        If cmbProvider.Items.Count > 0 Then
            cmbProvider.SelectedIndex = 0
        End If
    End Sub

    Private Sub cmbProvider_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbProvider.SelectedIndexChanged

        cmbShipMethod.Items.Clear()
        cmbShipMethod.Text = String.Empty
        Dim shipmethodprefix As String = String.Empty

        Dim CARRIER_CODE As String = tblSOTCARR1.Select("CARRIER_DESC = '" & cmbProvider.Text & "'", "")(0).Item("CARRIER_CODE")

        Select Case CARRIER_CODE

            Case "FEDEX"
                'Developer Key, Password, Account Number, and Meter Number are required.
                txtDeveloperKey.Enabled = True
                txtPassword.Enabled = True
                txtAccountNumber.Enabled = True
                txtMeterNumber.Enabled = True
                txtUserId.Enabled = False
                txtAccessKey.Enabled = False

                grdService.Text = "Federal Express Authentication"
                shipmethodprefix = "FEDEX"

            Case "UPS"
                'Access Key, User Id, Password, and Account Number are required
                txtDeveloperKey.Enabled = False
                txtPassword.Enabled = True
                txtAccountNumber.Enabled = True
                txtMeterNumber.Enabled = False
                txtUserId.Enabled = True
                txtAccessKey.Enabled = True

                grdService.Text = "UPS Authentication"
                shipmethodprefix = "UPS"

            Case "USPS"
                'UserId, Password, and Account are required.
                txtDeveloperKey.Enabled = False
                txtPassword.Enabled = True
                txtAccountNumber.Enabled = True
                txtMeterNumber.Enabled = False
                txtUserId.Enabled = True
                txtAccessKey.Enabled = False

                grdService.Text = "USPS (Endicia) Authentication"
                shipmethodprefix = "USPS"

            Case "CANADA"
                'UserId, Password, and Account are required.
                txtDeveloperKey.Enabled = False
                txtPassword.Enabled = True
                txtAccountNumber.Enabled = True
                txtMeterNumber.Enabled = False
                txtUserId.Enabled = True
                txtAccessKey.Enabled = False

                grdService.Text = "Canada Post Authentication"
                shipmethodprefix = "CANADA"

            Case Else
                txtDeveloperKey.Enabled = False
                txtPassword.Enabled = False
                txtAccountNumber.Enabled = False
                txtMeterNumber.Enabled = False
                txtUserId.Enabled = False
                txtAccessKey.Enabled = False
                grdService.Text = "Unknown Carrier"
        End Select

        Dim enumValues As Array = System.[Enum].GetValues(GetType(nsoftware.InShip.ServiceTypes))

        For Each resource As nsoftware.InShip.ServiceTypes In enumValues
            If resource.ToString.ToUpper.Contains(shipmethodprefix) Then
                cmbShipMethod.Items.Add(resource.ToString)
            End If
        Next
    End Sub

    Private Sub lblLabelFile_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblLabelFile.DoubleClick
        'System.Diagnostics.Process.Start(ezship1.Packages(0).ShippingLabelFile)
    End Sub

    Private Sub btnGetLabel_Click(sender As System.Object, e As System.EventArgs) Handles btnGetLabel.Click

        If (cmbProvider.SelectedIndex = 0) Then 'FedEx
            clsShip = New TAC.SCHSHIP1(SCHSHIP1.Services.FederalExpress)
        ElseIf (cmbProvider.SelectedIndex = 1) Then 'UPS
            clsShip = New TAC.SCHSHIP1(SCHSHIP1.Services.UPS)
        ElseIf (cmbProvider.SelectedIndex = 2) Then 'Endicia
            clsShip = New TAC.SCHSHIP1(SCHSHIP1.Services.USPS)
        Else 'CanadaPost
            clsShip = New TAC.SCHSHIP1(SCHSHIP1.Services.CanadaPost)
        End If

        Dim CARRIER_CODE As String = tblSOTCARR1.Select("CARRIER_DESC = '" & cmbProvider.Text & "'", "")(0).Item("CARRIER_CODE")
        Dim rowSOTCARR1 As DataRow = tblSOTCARR1.Select("CARRIER_CODE = '" & CARRIER_CODE & "'")(0)
        Dim rowSOTCARR3 As DataRow = tblSOTCARR3.Select("CARRIER_CODE = '" & CARRIER_CODE & "'")(0)

        With clsShip

            clsShip.Server = rowSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
            clsShip.UserId = rowSOTCARR3.Item("SHIPPER_ID") & String.Empty
            clsShip.Password = rowSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
            clsShip.AccountNumber = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty

            clsShip.UPSAccessKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty

            clsShip.FedexMeterNumber = rowSOTCARR3.Item("METER_NUMBER") & String.Empty
            clsShip.FedexDeveloperKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty

            'clsShip.FedExAccount.DeveloperKey = txtDeveloperKey.Text
            'clsShip.FedExAccount.Password = txtPassword.Text
            'clsShip.FedExAccount.AccountNumber = txtAccountNumber.Text
            'clsShip.FedExAccount.MeterNumber = txtMeterNumber.Text


            clsShip.RequestedServiceType = nsoftware.InShip.ServiceTypes.stUnspecified
            Dim enumValues As Array = System.[Enum].GetValues(GetType(nsoftware.InShip.ServiceTypes))

            For Each resource As nsoftware.InShip.ServiceTypes In enumValues
                If resource.ToString = cmbShipMethod.Text Then
                    Dim val = DirectCast([Enum].Parse(GetType(nsoftware.InShip.ServiceTypes), resource.ToString), nsoftware.InShip.ServiceTypes)
                    clsShip.RequestedServiceType = val
                    Exit For
                End If
            Next


            With clsShip.Sender
                .FirstName = txtFromFirstName.Text.Trim
                .MiddleInitial = ""
                .LastName = txtFromLastName.Text.Trim
                .Address1 = txtFromStreet.Text.Trim
                .Address2 = txtFromSuite.Text.Trim
                .City = txtFromCity.Text.Trim
                .State = txtFromState.Text.Trim.ToUpper
                .ZipCode = txtFromZip.Text.Trim
                .CountryCode = txtFromCountry.Text.Trim
                .Company = txtFromCompany.Text.Trim
                .Phone = txtFromPhone.Text
            End With

            With clsShip.Recipient
                .FirstName = txtToFirstName.Text.Trim
                .MiddleInitial = ""
                .LastName = txtToLastName.Text.Trim
                .Address1 = txtToStreet.Text.Trim
                .Address2 = txtToSuite.Text.Trim
                .City = txtToCity.Text.Trim
                .State = txtToState.Text.Trim
                .ZipCode = txtToZip.Text.Trim
                .CountryCode = txtToCountry.Text.Trim
                .Company = txtToCompany.Text.Trim
                .Phone = txtToPhone.Text

                If .Company.Length = 0 Then
                    .Company = (.FirstName & " " & .LastName).ToString.Trim
                End If
            End With

            clsShip.EzshipLabelImage = nsoftware.InShip.EzshipLabelImageTypes.itPNG

            Dim shipPackageDetail As New nsoftware.InShip.PackageDetail
            Dim shipPackageDetailList As New List(Of nsoftware.InShip.PackageDetail)

            With shipPackageDetail
                .PackagingType = nsoftware.InShip.TPackagingTypes.ptYourPackaging
                .Weight = txtWeight.Text.Trim
                .Length = Convert.ToInt32(txtLength.Text.Trim)
                .Width = Convert.ToInt32(txtWidth.Text.Trim)
                .Height = Convert.ToInt32(txtHeight.Text.Trim)
            End With

            shipPackageDetailList.Add(shipPackageDetail)
            If clsShip.RequestLabel(shipPackageDetailList) Then
                lblTrackingNumber.Text = shipPackageDetailList(0).TrackingNumber
                MessageBox.Show("Your tracking number is: " & shipPackageDetailList(0).TrackingNumber)
                picLabel.ImageLocation = shipPackageDetailList(0).ShippingLabelFile
            Else
                MessageBox.Show("Label could not be captured")
            End If

        End With

    End Sub
End Class