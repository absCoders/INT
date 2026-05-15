Imports nsoftware.InShip

Public Class SCHSHIP1

    Private objEzShip As New nsoftware.InShip.Ezship
    Private objFedExInternaltional As New nsoftware.InShip.Fedexshipintl


    Public Enum Services
        FederalExpress = EzshipProviders.pFedEx
        UPS = EzshipProviders.pUPS
        USPS = EzshipProviders.pUSPS
        CanadaPost = EzshipProviders.pCanadaPost
        Unknown = 7
    End Enum

    Private cService As Services = Services.Unknown
    Private cServer As String = String.Empty
    Private cUserId As String = String.Empty
    Private cPassword As String = String.Empty
    Private cRequestedServiceType As ServiceTypes
    Private cAccountNumber As String = String.Empty

    Private cFedexDeveloperKey As String = String.Empty
    Private cFedexMeterNumber As String = String.Empty

    Private cUPSAccessKey As String = String.Empty

    Private cUSPSEndiciaCustomerId As String = String.Empty
    Private cUSPSEndiciaTransactionId As String = String.Empty

    Private cTrackingNumber As String = String.Empty

    Private cSenderContact As New Contact
    Private CRecipientContact As New Contact

    Public PackageDetail As nsoftware.InShip.PackageDetail
    Public EzshipLabelImage As EzshipLabelImageTypes = EzshipLabelImageTypes.itEPL

    '"https://gatewaybeta.fedex.com:443/xml"
    '"https://wwwcie.ups.com/ups.app/xml/ShipConfirm"
    '"https://ct.soa-gw.canadapost.ca" 'development server
    ' endicia Server for PPS

#Region "Instantiate Class"

    Public Sub New()
        InitializeVariables()
    End Sub

    Public Sub New(ByVal ServiceType As Services)
        InitializeVariables()
        cService = ServiceType
    End Sub

    Private Sub InitializeVariables()
        cService = Services.Unknown
        cServer = String.Empty
        cUserId = String.Empty
        cPassword = String.Empty
        cRequestedServiceType = Nothing
        cAccountNumber = String.Empty

        cFedexDeveloperKey = String.Empty
        cFedexMeterNumber = String.Empty

        cUPSAccessKey = String.Empty

        cUSPSEndiciaCustomerId = String.Empty
        cUSPSEndiciaTransactionId = String.Empty
        cTrackingNumber = String.Empty

        cSenderContact = New Contact
        CRecipientContact = New Contact

        PackageDetail = New nsoftware.InShip.PackageDetail
        EzshipLabelImage = EzshipLabelImageTypes.itEPL

    End Sub

#End Region

#Region "Class Properties"

    ''' <summary>
    ''' Get / Set Service Provider
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Service As ServiceTypes
        Get
            Return cService
        End Get
        Set(value As ServiceTypes)
            cService = value
        End Set
    End Property

    ''' <summary>
    ''' Get / Set the Url for Service where requests are to be sent
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Server As String
        Get
            Server = cServer
        End Get
        Set(value As String)
            cServer = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Get / Set User Id for logging into the server. Not Required for Federal Express
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property UserId As String
        Get
            Return cUserId
        End Get
        Set(value As String)
            cUserId = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Get / Set Password for logging into the server
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Password As String
        Get
            Return (cPassword)
        End Get
        Set(value As String)
            cPassword = value
        End Set
    End Property

    ''' <summary>
    ''' Get / Set the domestic service used in the ship request
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property RequestedServiceType As ServiceTypes
        Get
            Return cRequestedServiceType
        End Get
        Set(value As ServiceTypes)
            cRequestedServiceType = value
        End Set
    End Property

    ''' <summary>
    ''' Get / Set the Shippers Account Number
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property AccountNumber As String
        Get
            Return cAccountNumber
        End Get
        Set(value As String)
            cAccountNumber = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Get / Set Identifting part of the authenication key useed for the sender's identity
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FedexDeveloperKey As String
        Get
            Return cFedexDeveloperKey
        End Get
        Set(value As String)
            cFedexDeveloperKey = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Get / Set Meter Number to use for submitting request to the Fedex Server.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FedexMeterNumber As String
        Get
            Return cFedexMeterNumber
        End Get
        Set(value As String)
            cFedexMeterNumber = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Get / Set and identifer required to connect to UPS
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property UPSAccessKey As String
        Get
            Return cUPSAccessKey
        End Get
        Set(value As String)
            cUPSAccessKey = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Get / Set Mandatory Custoder Id for Endicia
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property USPSEndiciaCustomerId As String
        Get
            Return cUSPSEndiciaCustomerId
        End Get
        Set(value As String)
            cUSPSEndiciaCustomerId = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Get / Set Mandatory Transaction ID for Endicia
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property USPSEndiciaTransactionId As String
        Get
            Return cUSPSEndiciaTransactionId
        End Get
        Set(value As String)
            cUSPSEndiciaTransactionId = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Set Sender Contact Information
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public Property Sender As Contact
        Get
            Return cSenderContact
        End Get
        Set(value As Contact)
            cSenderContact = value
        End Set
    End Property

    ''' <summary>
    ''' Set Recipient Information
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public Property Recipient As Contact
        Get
            Return CRecipientContact
        End Get
        Set(value As Contact)
            CRecipientContact = value
        End Set
    End Property

#End Region

#Region "Class Procedures"

    Public Function RequestLabel(ByRef PackageDetailList As List(Of nsoftware.InShip.PackageDetail)) As Boolean

        Try
            objEzShip.Reset()
            objFedExInternaltional.Reset()
            'objEzShip.RuntimeLicense = ASCMAIN1.nSoftwareKeys("")

            If cService = Services.Unknown Then
                Throw New Exception("Unknown Service Type")
                Return False
            End If

            objEzShip.Provider = cService
            objEzShip.Account.Server = cServer
            objEzShip.Account.AccountNumber = cAccountNumber
            objEzShip.Account.UserId = cUserId
            objEzShip.Account.Password = cPassword
            objEzShip.ServiceType = cRequestedServiceType

            If cService = Services.FederalExpress Then
                objEzShip.Account.MeterNumber = cFedexMeterNumber
                objEzShip.Account.DeveloperKey = cFedexDeveloperKey
            Else
                objEzShip.Account.MeterNumber = String.Empty
                objEzShip.Account.DeveloperKey = String.Empty
            End If

            If cService = Services.UPS Then
                objEzShip.Account.AccessKey = cUPSAccessKey
            Else
                objEzShip.Account.AccessKey = String.Empty
            End If

            If cService = Services.USPS Then
                objEzShip.Config("PostageProvider=1") 'Use Endicia instead of USPS directly.
                objEzShip.Config("CustomerId=" & cUSPSEndiciaCustomerId) 'Mandatory for Endicia
                objEzShip.Config("TransactionId=" & cUSPSEndiciaTransactionId) 'Mandatory for Endicia
            End If

            objEzShip.ShipDate = DateTime.Now.ToString("yyyy-MM-dd")


            With cSenderContact
                objEzShip.SenderContact.FirstName = .FirstName
                objEzShip.SenderContact.LastName = .LastName
                objEzShip.SenderContact.MiddleInitial = .MiddleInitial
                objEzShip.SenderContact.Phone = .Phone
                objEzShip.SenderContact.Fax = .Fax
                objEzShip.SenderContact.Email = .eMail

                objEzShip.SenderContact.Company = .Company
                objEzShip.SenderAddress.Address1 = .Address1
                objEzShip.SenderAddress.Address2 = .Address2
                objEzShip.SenderAddress.City = .City
                objEzShip.SenderAddress.ZipCode = .ZipCode
                objEzShip.SenderAddress.State = .State
                objEzShip.SenderAddress.CountryCode = .CountryCode
            End With

            With CRecipientContact
                objEzShip.RecipientContact.FirstName = .FirstName
                objEzShip.RecipientContact.LastName = .LastName
                objEzShip.RecipientContact.MiddleInitial = .MiddleInitial
                objEzShip.RecipientContact.Phone = .Phone
                objEzShip.RecipientContact.Fax = .Fax
                objEzShip.RecipientContact.Email = .eMail

                objEzShip.RecipientContact.Company = .Company
                objEzShip.RecipientAddress.Address1 = .Address1
                objEzShip.RecipientAddress.Address2 = .Address2
                objEzShip.RecipientAddress.City = .City
                objEzShip.RecipientAddress.ZipCode = .ZipCode
                objEzShip.RecipientAddress.State = .State
                objEzShip.RecipientAddress.CountryCode = .CountryCode
            End With

        Catch ex As Exception
            Return False
        End Try

        Try
            Dim ShippingLabelDirectory As String = ASCMAIN1.Folders("Temp")
            If Not ShippingLabelDirectory.EndsWith("\") Then ShippingLabelDirectory &= "\"

            Dim extension As String = EzshipLabelImage.ToString
            If extension.StartsWith("it") Then
                extension = "." & extension.Substring(2)
            Else
                extension = String.Empty
            End If

            objEzShip.LabelImageType = EzshipLabelImageTypes.itPNG

            For Each shippingPackageDetail In PackageDetailList
                shippingPackageDetail.ShippingLabelFile = ShippingLabelDirectory & ASCMAIN1.Next_Control_No("SHIP_REQUEST") & extension
                objEzShip.Packages.Add(shippingPackageDetail)
            Next

            objEzShip.GetShipmentLabels()

            ' Reset the object to have the updated data returned
            PackageDetailList.Clear()
            For ictr As Int16 = 0 To objEzShip.Packages.Count - 1
                PackageDetailList.Add(objEzShip.Packages(ictr))
            Next

        Catch ex As nsoftware.InShip.InShipException
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Catch exc As Exception
            MessageBox.Show("Error: " & exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

        Return True

    End Function

#End Region

#Region "Internal Classes"

    Class Contact
        Public FirstName As String = String.Empty
        Public LastName As String = String.Empty
        Public MiddleInitial As String = String.Empty
        Public Phone As String = String.Empty
        Public eMail As String = String.Empty
        Public Fax As String = String.Empty
        Public Company As String = String.Empty

        ' Address Attributes
        Public Address1 As String = String.Empty
        Public Address2 As String = String.Empty
        Public City As String = String.Empty
        Public State As String = String.Empty
        Public ZipCode As String = String.Empty
        Public CountryCode As String = String.Empty
        Public IsResidental As Boolean = False
    End Class

#End Region

End Class
