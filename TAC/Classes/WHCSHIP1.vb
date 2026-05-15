Imports System.IO
Imports DPayments.DShippingSDK
Imports Newtonsoft.Json

Public Class WHCSHIP1

#Region "Class Variables"

    Private objEzShip As New EzShip
    Private objEzRates As New EzRates

    Private objFedexShipIntl As New FedExShipIntl
    Private objFedexShip As New FedExShip
    Private objFedexRates As New FedExRates

    Private objUpsShip As New UPSShip
    Private objUpsShipIntl As New UPSShipIntl
    Private objUpsRates As New UPSRates
    Private objUpsFreight As New UPSFreightShip

    'Private objUspsShip As New Uspsship
    'Private objUspsShipIntl As New Uspsshipintl

    Public Enum ServiceProviders
        FederalExpress = EzShipProviders.pFedEx
        UPS = EzShipProviders.pUPS
        USPS = EzShipProviders.pUSPS
        CanadaPost = EzShipProviders.pCanadaPost
        FederalExpressInternational = 5
        UPSInternational = 6
        Unknown = 7
        SpeeDee = 9
        RTNY = 10
        Unity = 11
        GenericLabel = 12
    End Enum

    Public EzshipLabelImage As EzShipLabelImageTypes = EzShipLabelImageTypes.itZebra

    Public UPSPickupType As UPSRatesPickupTypes = UPSRatesPickupTypes.ptDailyPickup
    Private cCustomerType As UPSRatesCustomerTypes = UPSRatesCustomerTypes.ccDaily

    Public ShippingLabelDirectory As String = String.Empty
    Public ShippingLabelPrefix As String = String.Empty
    Public PackageDetailList As New List(Of PackageDetail)

    Public ShipDate As Date = DateTime.Now
    Public ShipmentSpecialServices As Long = 0
    Public CommodityDetailList As New List(Of CommodityDetail)
    Public RequestedServicesRates As New List(Of ServiceDetail)
    Public HandlingUnit As String = String.Empty

    Public RTNYLabel As String = String.Empty

    Private cServiceProvider As ServiceProviders = ServiceProviders.Unknown
    Private cServer As String = String.Empty
    Private cUserId As String = String.Empty
    Private cPassword As String = String.Empty
    Private cRequestedServiceType As ServiceTypes
    Private cAccountNumber As String = String.Empty
    Private cTotalCustomsValue As Decimal = 0
    Private cSignatureRequired As Boolean = False

    Private cFedexDeveloperKey As String = String.Empty
    Private cFedexMeterNumber As String = String.Empty
    Private cLabelStockType As String = String.Empty

    Private cUPSAccessKey As String = String.Empty

    Private cUSPSEndiciaCustomerId As String = String.Empty
    Private cUSPSEndiciaTransactionId As String = String.Empty

    Private cSenderContact As New Contact
    Private cRecipientContact As New Contact
    Private cShipFrom As New Contact
    Private cPayorContact As New Contact
    Private cDutiesPayorContact As New Contact
    Private cHoldAtLocation As New Contact
    Private cFedexSmartPost As New SmartPost

    Public Payor As TPayorTypes = TPayorTypes.ptSender
    Public DutiesPayor As TPayorTypes = TPayorTypes.ptRecipient

    Public LastError As String = String.Empty
    Private cMasterTrackingNumber As String = String.Empty
    Private cRawRequest As String = String.Empty
    Private cRawResponse As String = String.Empty

    Private s4DPaymentsShippingSDK As String = TACMAIN1.s4DPaymentsShippingSDK
    Public FedexClose As New CloseDetail

    Public Const ProviderTypeFedex = "F"
    Public Const ProviderTypeUPS = "U"
    Public Const ProviderTypeUSPS = "P"
    Public Const ProviderTypeCanada = "C"

    Public ShipmentBaseCharge As New Dictionary(Of Int32, Decimal)
    Public ShipmentDiscountCharge As New Dictionary(Of Int32, Decimal)
    Public ShipmentListCharge As New Dictionary(Of Int32, Decimal)
    Public ShipmentNetCharge As New Dictionary(Of Int32, Decimal)
    Public ShipmentSurCharge As New Dictionary(Of Int32, Decimal)

    Private cFedexCustomContent As String = String.Empty
    Private Const SSLEnabledProtocols As Int32 = 4032

    Public Enum NotifictaionTypes
        On_Shipment = 1
        On_Exception = 2
        On_Deleivery = 4
        'On_Tender = 8
        'On_Return_UPS = 10 ' Not Used by FedEx
        'HTML_FedEx = 20 ' Not Used by UPS
        'Text_Fedex = 40  ' Not Used by UPS
        'Wireless_Fedex = 80 ' Not Used by UPS
    End Enum

    Public Class Notifications
        Public email As String = String.Empty
        Public NotificationFlags As NotifictaionTypes
        Public Message As String = String.Empty
    End Class

    Public ShipmentNotifications As New List(Of Notifications)

    Public Enum ShipperRates
        List = 0
        Account = 1
    End Enum

    ' Test Regions
    'https://gatewaybeta.fedex.com:443/xml
    'https://wsbeta.fedex.com:443/web-services
    'https://wwwcie.ups.com/ups.app/xml
    'https://ct.soa-gw.canadapost.ca" 'development server
    ' endicia Server for PPS

    ' Production Regions 
    ' https://ws.fedex.com:443/xml
    ' https://onlinetools.ups.com/ups.app/xml/

    Public Structure RateList
        Dim ServiceType As String
        Dim ServiceTypeDescription As String
        Dim TransitTime As String
        Dim DeliveryTime As String
        Dim AccountNetCharge As Decimal
        Dim ListNetCharge As Decimal
        Dim ReferenceIndex As Int16
        Dim OfferID As String
        Dim ServiceCode As String
    End Structure

    Public Structure TransTimeList
        Dim ServiceCode As String
        Dim ServiceDescription As String
        Dim Guaranteed As String
        Dim TransitDays As Int16
    End Structure

    Private Class Address
        Public Name As String
        Public Address1 As String
        Public Address2 As String
        Public Address3 As String
        Public City As String
        Public State As String
        Public Zip As String
        Public Contact As String
        Public Phone As String
        Public Country As String
    End Class

    Public lstShippingLabels As New Dictionary(Of String, String)
    Public lstShippingFiles As New Dictionary(Of String, String)
    Public lstTrackingNumbers As New Dictionary(Of String, String)

    Public GenericLabelShipMethod As String = String.Empty
    Public GenericLabelShipMethodDesc As String = String.Empty
    Public GenericLabelCarrier As String = String.Empty

#End Region

#Region "Instantiate Class"

    Public Sub New()
        InitializeVariables()
    End Sub

    Public Sub New(ByVal ServiceType As ServiceProviders)
        InitializeVariables()
        cServiceProvider = ServiceType
    End Sub

    ''' <summary>
    ''' Set all Objects to the default values
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Reset()
        InitializeVariables()
    End Sub

    ''' <summary>
    ''' Initialize class objects
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitializeVariables()
        objEzShip = New EzShip With {
            .RuntimeLicense = s4DPaymentsShippingSDK
        }

        objFedexShipIntl = New FedExShipIntl With {
            .RuntimeLicense = s4DPaymentsShippingSDK
        }

        objFedexShip = New FedExShip With {
            .RuntimeLicense = s4DPaymentsShippingSDK
        }

        objFedexRates = New FedExRates With {
            .RuntimeLicense = s4DPaymentsShippingSDK
        }

        objEzRates = New EzRates With {
            .RuntimeLicense = s4DPaymentsShippingSDK
        }

        objUpsShip = New UPSShip With {
            .RuntimeLicense = s4DPaymentsShippingSDK
        }

        objUpsShipIntl = New UPSShipIntl With {
            .RuntimeLicense = s4DPaymentsShippingSDK
        }

        objUpsRates = New UPSRates With {
            .RuntimeLicense = s4DPaymentsShippingSDK
        }

        objUpsFreight = New UPSFreightShip With {
            .RuntimeLicense = s4DPaymentsShippingSDK
        }

        FedexClose = New CloseDetail

        EzshipLabelImage = EzShipLabelImageTypes.itZebra
        ShippingLabelDirectory = String.Empty
        ShippingLabelPrefix = String.Empty
        HandlingUnit = String.Empty
        ShipDate = DateTime.Now
        ShipmentSpecialServices = 0
        CommodityDetailList = New List(Of CommodityDetail)

        ShipmentBaseCharge.Clear()
        ShipmentDiscountCharge.Clear()
        ShipmentListCharge.Clear()
        ShipmentNetCharge.Clear()
        ShipmentSurCharge.Clear()

        cServiceProvider = ServiceProviders.Unknown
        cServer = String.Empty
        cUserId = String.Empty
        cPassword = String.Empty
        cRequestedServiceType = New ServiceTypes
        cAccountNumber = String.Empty
        cLabelStockType = String.Empty

        cFedexDeveloperKey = String.Empty
        cFedexMeterNumber = String.Empty

        cUPSAccessKey = String.Empty

        cUSPSEndiciaCustomerId = String.Empty
        cUSPSEndiciaTransactionId = String.Empty
        cTotalCustomsValue = 0
        cMasterTrackingNumber = String.Empty

        cSenderContact = New Contact
        cRecipientContact = New Contact
        cPayorContact = New Contact
        cDutiesPayorContact = New Contact
        cHoldAtLocation = New Contact
        cFedexSmartPost = New SmartPost
        cShipFrom = New Contact

        Payor = TPayorTypes.ptSender
        DutiesPayor = TPayorTypes.ptRecipient

        cRawRequest = String.Empty
        cRawResponse = String.Empty
        cFedexCustomContent = String.Empty

        PackageDetailList = New List(Of PackageDetail)
        cSignatureRequired = False

        ShipmentNotifications = New List(Of Notifications)
        lstShippingLabels = New Dictionary(Of String, String)

        GenericLabelShipMethod = String.Empty
        GenericLabelShipMethodDesc = String.Empty
        GenericLabelCarrier = String.Empty

    End Sub

#End Region

#Region "Class Properties"

    Private cShipmentDescription As String = String.Empty
    Public Property ShipmentDescription As String
        Get
            Return cShipmentDescription
        End Get
        Set(ByVal value As String)
            cShipmentDescription = value
        End Set
    End Property

    Public Property CustomerType() As UPSRatesCustomerTypes
        Get
            Return cCustomerType
        End Get
        Set(ByVal value As UPSRatesCustomerTypes)
            cCustomerType = value
        End Set
    End Property

    ''' <summary>
    '''  Custom Content to place on a Fedex Label
    '''  When using the CustomContent, the LabelStockType must be either 4 (Stock 4x8) or 5 (Stock 4x9 Leading Doc Tab). 
    '''  Also LabelFormatType must be 0 (Common2D) and LabelImageType must 2 (fitEltron), 3 (fitZebra) or 4 (fitUniMark). 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FedexCustomContent As String
        Get
            Return cFedexCustomContent
        End Get
        Set(value As String)
            cFedexCustomContent = value
        End Set
    End Property

    ''' <summary>
    ''' Gets set if a signature is required.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SignatureRequired As Boolean
        Get
            Return cSignatureRequired
        End Get
        Set(value As Boolean)
            cSignatureRequired = value
        End Set
    End Property

    ''' <summary>
    ''' Gets the raw request sent to shipper
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property RawRequest As String
        Get
            Return cRawRequest
        End Get
    End Property

    ''' <summary>
    ''' Gets the Raw Response returned from the shipper
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property RawResponse As String
        Get
            Return cRawResponse
        End Get
    End Property

    ''' <summary>
    ''' Get Shipment Master Tracking Number
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property MasterTrackingNumber
        Get
            Return cMasterTrackingNumber
        End Get
    End Property

    ''' <summary>
    ''' Get / Set Service Provider
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Service As ServiceProviders
        Get
            Return cServiceProvider
        End Get
        Set(value As ServiceProviders)
            cServiceProvider = value
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
            Return cRecipientContact
        End Get
        Set(value As Contact)
            cRecipientContact = value
        End Set
    End Property

    ''' <summary>
    ''' Set Ship From Information
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public Property ShipFrom As Contact
        Get
            Return cShipFrom
        End Get
        Set(value As Contact)
            cShipFrom = value
        End Set
    End Property

    ''' <summary>
    ''' Set Payor Information
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PayorContact As Contact
        Get
            Return cPayorContact
        End Get
        Set(value As Contact)
            cPayorContact = value
        End Set
    End Property

    ''' <summary>
    ''' Set Duties Payor Information
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DutiesPayorContact As Contact
        Get
            Return cDutiesPayorContact
        End Get
        Set(value As Contact)
            cDutiesPayorContact = value
        End Set
    End Property

    ''' <summary>
    ''' Set Hold At Location Information
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property HoldAtLocation As Contact
        Get
            Return cHoldAtLocation
        End Get
        Set(value As Contact)
            cHoldAtLocation = value
        End Set
    End Property

    ''' <summary>
    ''' Set / Get total customs Value
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TotalCustomsValue As Decimal
        Get
            Return cTotalCustomsValue
        End Get
        Set(value As Decimal)
            cTotalCustomsValue = value
        End Set
    End Property

    ''' <summary>
    ''' Get / Set The Indicia type used for a FedEx SmartPost shipment.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FedexSmartPost As SmartPost
        Get
            Return cFedexSmartPost
        End Get
        Set(value As SmartPost)
            cFedexSmartPost = value
        End Set
    End Property

    ''' <summary>
    ''' Get / Set label stock type
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LabelStockType As String
        Get
            Return cLabelStockType
        End Get
        Set(value As String)
            cLabelStockType = value
        End Set
    End Property


#End Region

#Region "Public Class Procedures"

    ''' <summary>
    ''' Request Shipping Label
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function RequestLabel() As Boolean

        Try
            lstShippingLabels = New Dictionary(Of String, String)
            lstShippingFiles = New Dictionary(Of String, String)

            cMasterTrackingNumber = String.Empty
            cRawRequest = String.Empty
            cRawResponse = String.Empty
            LastError = String.Empty
            RequestedServicesRates.Clear()

            If cServiceProvider = ServiceProviders.Unknown Then
                LastError = "Unknown Service Type"
                Return False
            End If

            ShipmentBaseCharge.Clear()
            ShipmentDiscountCharge.Clear()
            ShipmentListCharge.Clear()
            ShipmentNetCharge.Clear()
            ShipmentSurCharge.Clear()

            ValidateInputFieldLengths(Recipient.Address1,
                Recipient.Address2,
                Recipient.Address3,
                Recipient.FirstName,
                Recipient.City,
                Recipient.State,
                Recipient.ZipCode,
                Recipient.Company)

            ValidateInputFieldLengths(Sender.Address1,
                Sender.Address2,
                Sender.Address3,
                Sender.FirstName,
                Sender.City,
                Sender.State,
                Sender.ZipCode,
                Sender.Company)

            Select Case cServiceProvider
                Case ServiceProviders.FederalExpressInternational
                    Return RequestFedexInternaltionalLabel()

                Case ServiceProviders.FederalExpress
                    Return RequestFedexLabel()

                Case ServiceProviders.UPS
                    If HandlingUnit.Length > 0 Then
                        Return RequestUpsFreightlabel()
                    Else
                        Return RequestUPSLabel()
                    End If

                Case ServiceProviders.UPSInternational
                    Return RequestUPSInternaltionalLabel()

                Case ServiceProviders.GenericLabel
                    Return RequestGenericLabel()

                Case Else
                    Return RequestLabelOther()

            End Select

        Catch ex As Exception
            LastError = ex.Message
            Return False
        End Try

    End Function

    Private Sub ValidateInputFieldLengths(ByRef AddressLine1 As String,
            ByRef AddressLine2 As String,
            ByRef AddressLine3 As String,
            ByRef AttentionName As String,
            ByRef City As String,
            ByRef StateProvinceCode As String,
            ByRef PostalCode As String,
            ByRef CompanyName As String)

        AddressLine1 = AddressLine1 & String.Empty
        If AddressLine1.Length > 35 Then
            AddressLine1 = AddressLine1.Substring(0, 35).Trim
        End If

        AddressLine2 = AddressLine2 & String.Empty
        If AddressLine2.Length > 35 Then
            AddressLine2 = AddressLine2.Substring(0, 35).Trim
        End If

        AddressLine3 = AddressLine3 & String.Empty
        If AddressLine3.Length > 35 Then
            AddressLine3 = AddressLine3.Substring(0, 35).Trim
        End If

        AttentionName = AttentionName & String.Empty
        If AttentionName.Length > 35 Then
            AttentionName = AttentionName.Substring(0, 35).Trim
        End If

        City = City & String.Empty
        If City.Length > 30 Then
            City = City.Substring(0, 30).Trim
        End If

        StateProvinceCode = StateProvinceCode & String.Empty
        If StateProvinceCode.Length > 5 Then
            StateProvinceCode = StateProvinceCode.Substring(0, 5).Trim
        End If

        PostalCode = PostalCode & String.Empty
        If PostalCode.Length > 10 Then
            PostalCode = PostalCode.Substring(0, 10).Trim
        End If

        CompanyName = CompanyName & String.Empty
        If CompanyName.Length > 35 Then
            CompanyName = CompanyName.Substring(0, 35).Trim
        End If

    End Sub

    Public Function CancelShipment(ByVal TrackingNumber As String) As Boolean
        Return CancelShipment(TrackingNumber, False, 0)
    End Function

    Public Function CancelShipment(ByVal TrackingNumber As String, ByVal isMultiPackage As Boolean, ByVal FedexTrackingIDType As Int16) As Boolean

        Select Case cServiceProvider
            Case ServiceProviders.CanadaPost
                Return False
            Case ServiceProviders.FederalExpress, ServiceProviders.FederalExpressInternational
                Return CancelFedexShipment(TrackingNumber, isMultiPackage, FedexTrackingIDType)
            Case ServiceProviders.UPS
                Return CancelUpsShipment(TrackingNumber, isMultiPackage)
            Case ServiceProviders.USPS
                Return False
            Case Else
                LastError = "Unknown Carrier"
                Return False
        End Select

    End Function

#End Region

#Region "Federal Express"

    ''' <summary>
    ''' Request Fedex International Shipping label
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RequestFedexInternaltionalLabel() As Boolean

        Try
            LastError = String.Empty
            objFedexShipIntl.RuntimeLicense = s4DPaymentsShippingSDK
            objFedexShipIntl.Reset()
            objFedexShipIntl.RuntimeLicense = s4DPaymentsShippingSDK

            If cServiceProvider <> ServiceProviders.FederalExpressInternational Then
                LastError = "Service type not a Fedex International."
                Return False
            End If
            ' Set credentials
            Dim bearertoken As String = GetAuthorizationToken(Carriers.FedEx, cAccountNumber)
            If inTestMode Then
                objFedexShipIntl.Config("TESTMODE=true")
                cAccountNumber = "740561073"
            End If

            objFedexShip.FedExAccount.AccountNumber = cAccountNumber
            objFedexShip.FedExAccount.AuthorizationToken = bearertoken
            'objFedexShipIntl.FedExAccount.DeveloperKey = cFedexDeveloperKey
            'objFedexShipIntl.FedExAccount.Password = cPassword
            'objFedexShipIntl.FedExAccount.MeterNumber = cFedexMeterNumber
            objFedexShipIntl.ServiceType = cRequestedServiceType

            objFedexShipIntl.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)

            ' Get Sender Information
            With cSenderContact
                objFedexShipIntl.SenderContact.FirstName = .FirstName
                objFedexShipIntl.SenderContact.LastName = .LastName
                objFedexShipIntl.SenderContact.MiddleInitial = .MiddleInitial
                objFedexShipIntl.SenderContact.Phone = .Phone
                objFedexShipIntl.SenderContact.Fax = .Fax
                objFedexShipIntl.SenderContact.Email = .eMail

                objFedexShipIntl.SenderContact.Company = .Company
                objFedexShipIntl.SenderAddress.Address1 = .Address1
                objFedexShipIntl.SenderAddress.Address2 = .Address2
                'objFedexShipIntl.Config("SenderAddress3=" & .Address3)

                objFedexShipIntl.SenderAddress.City = .City
                objFedexShipIntl.SenderAddress.ZipCode = .ZipCode
                objFedexShipIntl.SenderAddress.State = .State
                objFedexShipIntl.SenderAddress.CountryCode = .CountryCode

                If .IsResidental Then
                    objFedexShipIntl.SenderAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objFedexShipIntl.SenderAddress.AddressFlags = &H1 'PO Box
                End If
            End With

            With cRecipientContact
                objFedexShipIntl.RecipientContact.FirstName = .FirstName
                objFedexShipIntl.RecipientContact.LastName = .LastName
                objFedexShipIntl.RecipientContact.MiddleInitial = .MiddleInitial
                objFedexShipIntl.RecipientContact.Phone = .Phone
                objFedexShipIntl.RecipientContact.Fax = .Fax
                objFedexShipIntl.RecipientContact.Email = .eMail

                objFedexShipIntl.RecipientContact.Company = .Company
                objFedexShipIntl.RecipientAddress.Address1 = .Address1
                objFedexShipIntl.RecipientAddress.Address2 = .Address2
                'objFedexShipIntl.Config("RecipientAddress3=" & .Address3)

                objFedexShipIntl.RecipientAddress.City = .City
                objFedexShipIntl.RecipientAddress.ZipCode = .ZipCode
                objFedexShipIntl.RecipientAddress.State = .State
                objFedexShipIntl.RecipientAddress.CountryCode = .CountryCode

                If .IsResidental Then
                    objFedexShipIntl.RecipientAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objFedexShipIntl.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

            Select Case EzshipLabelImage
                Case EzShipLabelImageTypes.itEltron
                    objFedexShipIntl.LabelImageType = FedExShipIntlLabelImageTypes.fitEltron
                Case EzShipLabelImageTypes.itPDF
                    objFedexShipIntl.LabelImageType = FedExShipIntlLabelImageTypes.fitPDF
                Case EzShipLabelImageTypes.itPNG
                    objFedexShipIntl.LabelImageType = FedExShipIntlLabelImageTypes.fitPNG
                Case EzShipLabelImageTypes.itUniMark
                    objFedexShipIntl.LabelImageType = FedExShipIntlLabelImageTypes.fitUniMark
                Case EzShipLabelImageTypes.itZebra
                    objFedexShipIntl.LabelImageType = FedExShipIntlLabelImageTypes.fitZebra
                Case Else ' if not a valid option default to fitEltron
                    objFedexShipIntl.LabelImageType = FedExShipIntlLabelImageTypes.fitEltron
            End Select

            ' 02/06/2024 - Always get ZPL
            objFedexShipIntl.LabelImageType = FedExShipIntlLabelImageTypes.fitZebra

            Dim extension As String = objFedexShipIntl.LabelImageType.ToString
            If extension.StartsWith("fit") Then
                extension = "." & extension.Substring(3)
            Else
                extension = String.Empty
            End If

            If LabelStockType.Length > 0 Then
                objFedexShipIntl.Config("LabelStockType=" & LabelStockType)
            End If

            ' Set Shipping Label File
            Dim idCtr As Int16 = 1
            If ShippingLabelDirectory.Length > 0 AndAlso ShippingLabelPrefix.Length > 0 Then
                If Not ShippingLabelDirectory.EndsWith("\") Then
                    ShippingLabelDirectory &= "\"
                End If

                For Each shippingPackageDetail In PackageDetailList
                    Dim id As String = idCtr.ToString
                    idCtr += 1
                    shippingPackageDetail.ShippingLabelFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & extension
                    shippingPackageDetail.CODFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & "_COD" & extension
                Next
            End If

            Dim totalWeight As Double = 0
            Dim totalInsured As Double = 0

            ' Add packages
            For Each shippingPackageDetail In PackageDetailList
                ' Add packages (package weight is in Ounces - Convert to Pounds)
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight) / 16, "###0.0")
                totalWeight += Val(shippingPackageDetail.Weight)
                totalInsured += Val(shippingPackageDetail.InsuredValue & String.Empty)
                If cSignatureRequired Then
                    shippingPackageDetail.SignatureType = TSignatureTypes.stDirect
                End If
                objFedexShipIntl.Packages.Add(shippingPackageDetail)
            Next

            objFedexShipIntl.TotalWeight = Format(totalWeight, "###0.0")
            objFedexShipIntl.InsuredValue = Format(totalInsured, "###0.00")
            objFedexShipIntl.ShipDate = ShipDate.ToString("yyyy-MM-dd")
            objFedexShipIntl.TotalCustomsValue = Format(TotalCustomsValue, "###0.00")

            ' Service Type
            objFedexShipIntl.ServiceType = cRequestedServiceType
            objFedexShipIntl.Payor.PayorType = Payor

            objFedexShipIntl.Payor.AccountNumber = cSenderContact.AccountNumber
            objFedexShipIntl.Payor.CountryCode = cSenderContact.CountryCode

            objFedexShipIntl.DutiesPayor.PayorType = DutiesPayor
            objFedexShipIntl.DutiesPayor.AccountNumber = cDutiesPayorContact.AccountNumber
            objFedexShipIntl.DutiesPayor.CountryCode = cDutiesPayorContact.CountryCode

            Dim specialService As Long = ShipmentSpecialServices

            If objFedexShipIntl.ShipDate > DateTime.Now.ToString("yyyy-MM-dd") Then
                specialService = specialService Or &H20000000L
            End If

            objFedexShipIntl.ShipmentSpecialServices = specialService

            objFedexShipIntl.HoldAtLocation.Address1 = cHoldAtLocation.Address1
            objFedexShipIntl.HoldAtLocation.Address2 = cHoldAtLocation.Address2
            objFedexShipIntl.HoldAtLocation.City = cHoldAtLocation.City
            objFedexShipIntl.HoldAtLocation.State = cHoldAtLocation.State
            objFedexShipIntl.HoldAtLocation.ZipCode = cHoldAtLocation.ZipCode
            objFedexShipIntl.HoldAtLocationPhone = cHoldAtLocation.Phone

            With objFedexShipIntl.Payor
                .PayorType = Payor
                .AccountNumber = PayorContact.AccountNumber
                .CountryCode = PayorContact.CountryCode
                .ZipCode = PayorContact.ZipCode
            End With

            With objFedexShipIntl.DutiesPayor
                .PayorType = DutiesPayor
                .AccountNumber = DutiesPayorContact.AccountNumber
                .CountryCode = DutiesPayorContact.CountryCode
                .ZipCode = DutiesPayorContact.ZipCode
            End With

            For Each CommDetail As CommodityDetail In CommodityDetailList
                CommDetail.Weight = Format(Val(CommDetail.Weight), "###0.0")
                CommDetail.Description = CommDetail.Description.Replace("&", " ").Replace("<", " ").Replace(">", " ")
                objFedexShipIntl.Commodities.Add(CommDetail)
            Next

        Catch ex As Exception
            Return False
        End Try

        ' Notifications - Not Supported by FedEx International
        Dim notificationsIndex As Int16 = 0
        If ShipmentNotifications.Count > 0 AndAlso Not ASCMAIN1.Running_in_VS Then
            For Each sn As Notifications In ShipmentNotifications
                sn.email = (sn.email & String.Empty).Trim
                If sn.email.Length = 0 Then
                    Continue For
                End If

                Dim notify As New NotifyDetail
                With notify
                    .Email = sn.email
                    .NotificationFlags = CInt(sn.NotificationFlags)
                    .Message = (sn.Message & String.Empty).ToString.Trim
                End With

                notificationsIndex += 1
                If notificationsIndex = 3 Then Exit For
            Next
        End If

        Try

            objFedexShipIntl.RuntimeLicense = s4DPaymentsShippingSDK
            objFedexShipIntl.GetShipmentLabels()

            ' Reset the object to have the updated data returned
            PackageDetailList.Clear()
            For ictr As Int16 = 0 To objFedexShipIntl.Packages.Count - 1
                PackageDetailList.Add(objFedexShipIntl.Packages(ictr))
                GetPackageCosts(objFedexShipIntl.Packages(ictr), objFedexShipIntl)
            Next

            If objFedexShipIntl.Packages.Count = 1 Then
                cMasterTrackingNumber = objFedexShipIntl.Packages(0).TrackingNumber
            Else
                cMasterTrackingNumber = objFedexShipIntl.MasterTrackingNumber
            End If

            Return True

        Catch ex As DShippingSDKException
            LastError = ex.Message
            Return False
        Catch exc As Exception
            LastError = exc.Message
            Return False
        Finally
            cRawRequest = objFedexShipIntl.Config("RawRequest")
            cRawResponse = objFedexShipIntl.Config("RawResponse")
        End Try

        Return True

    End Function

    ''' <summary>
    ''' Request Shipping label. Not used for Fedex Intenational
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RequestFedexLabel() As Boolean

        Try
            LastError = String.Empty
            objFedexShip.RuntimeLicense = s4DPaymentsShippingSDK
            objFedexShip.Reset()
            objFedexShip.RuntimeLicense = s4DPaymentsShippingSDK

            If cServiceProvider = ServiceProviders.Unknown Then
                LastError = "Unknown Service Type"
                Return False
            End If

            Dim bearertoken As String = GetAuthorizationToken(Carriers.FedEx, cAccountNumber)

            If inTestMode Then
                ' FedEx does not have a full test environment
                ' Need tyo hard code Label information.
                cAccountNumber = "740561073"
                objFedexShip.Config("TESTMODE=true") ' no need To Set the server

                With cSenderContact
                    .Company = ""
                    .eMail = ""
                    .FirstName = "Test"
                    .LastName = "Test"
                    .MiddleInitial = ""
                    .Phone = "8889997777"
                    .Address1 = "5007 Southpark Drive"
                    .Address2 = "Suite 240"
                    .City = "Durham"
                    .CountryCode = "US"
                    .State = "NC"
                    .ZipCode = "27713"
                End With

                With cRecipientContact
                    .Company = ""
                    .eMail = ""
                    .FirstName = "Test & Such"
                    .LastName = ""
                    .MiddleInitial = ""
                    .Phone = "0000000000"
                    .Address1 = "8355 Rockville Rd"
                    .Address2 = "Suite B"
                    .City = "Indianapolis"
                    .CountryCode = "US"
                    .State = "IN"
                    .ZipCode = "46234"
                End With

                For Each shippingPackageDetail In PackageDetailList
                    With shippingPackageDetail
                        .Height = 5
                        .Length = 5
                        .Width = 5
                        .Weight = "16" ' 16 OPUNCES
                        .PackagingType = TPackagingTypes.ptYourPackaging
                    End With
                Next

            End If

            objFedexShip.FedExAccount.AccountNumber = cAccountNumber
            objFedexShip.FedExAccount.AuthorizationToken = bearertoken
            objFedexShip.PickupType = FedExRatesPickupTypes.fptUseScheduledPickup

            'objFedexShip.FedExAccount.Password = cPassword
            'objFedexShip.FedExAccount.MeterNumber = cFedexMeterNumber
            'objFedexShip.FedExAccount.DeveloperKey = cFedexDeveloperKey

            objFedexShip.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)

            objFedexShip.ServiceType = cRequestedServiceType

            If objFedexShip.ServiceType = ServiceTypes.stFedExSmartPost Then
                objFedexShip.Config("SmartPostIndicia=" & FedexSmartPost.Indicia)
                objFedexShip.Config("SmartPostHubId=" & FedexSmartPost.HubId)
                objFedexShip.Config("SmartPostPhysicalPackaging=" & FedexSmartPost.PhysicalPackaging)
                objFedexShip.Config("SmartPostAncillaryEndorsement=" & FedexSmartPost.AncillaryEndorsement)
            End If

            objFedexShip.ShipDate = ShipDate.ToString("yyyy-MM-dd")

            With cSenderContact
                objFedexShip.SenderContact.FirstName = .FirstName
                objFedexShip.SenderContact.LastName = .LastName
                objFedexShip.SenderContact.MiddleInitial = .MiddleInitial
                objFedexShip.SenderContact.Phone = .Phone
                objFedexShip.SenderContact.Fax = .Fax
                objFedexShip.SenderContact.Email = .eMail

                objFedexShip.SenderContact.Company = .Company
                objFedexShip.SenderAddress.Address1 = .Address1
                objFedexShip.SenderAddress.Address2 = .Address2
                'objFedexShip.Config("SenderAddress3=" & .Address3)

                objFedexShip.SenderAddress.City = .City
                objFedexShip.SenderAddress.ZipCode = .ZipCode
                objFedexShip.SenderAddress.State = .State
                objFedexShip.SenderAddress.CountryCode = .CountryCode

                If .IsResidental Then
                    'objFedexShip.SenderAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    'objFedexShip.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

            With cRecipientContact
                objFedexShip.RecipientContact.FirstName = .FirstName
                objFedexShip.RecipientContact.LastName = .LastName
                objFedexShip.RecipientContact.MiddleInitial = .MiddleInitial
                objFedexShip.RecipientContact.Phone = .Phone
                objFedexShip.RecipientContact.Fax = .Fax
                objFedexShip.RecipientContact.Email = .eMail

                objFedexShip.RecipientContact.Company = .Company
                objFedexShip.RecipientAddress.Address1 = .Address1
                objFedexShip.RecipientAddress.Address2 = .Address2
                'objFedexShip.Config("RecipientAddress3=" & .Address3)

                objFedexShip.RecipientAddress.City = .City
                objFedexShip.RecipientAddress.ZipCode = .ZipCode
                objFedexShip.RecipientAddress.State = .State
                objFedexShip.RecipientAddress.CountryCode = .CountryCode

                If .IsResidental Then
                    objFedexShip.RecipientAddress.AddressFlags = &H2 'Residential
                    If objFedexShip.ServiceType = ServiceTypes.stFedExGround Then
                        objFedexShip.ServiceType = ServiceTypes.stFedExGroundHomeDelivery
                    End If
                ElseIf .IsPOBox Then
                    objFedexShip.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

        Catch ex As Exception
            LastError = ex.Message
            Return False
        End Try

        Try
            Select Case EzshipLabelImage
                Case EzShipLabelImageTypes.itEltron
                    objFedexShip.LabelImageType = FedExShipLabelImageTypes.fitEltron
                Case EzShipLabelImageTypes.itPDF
                    objFedexShip.LabelImageType = FedExShipLabelImageTypes.fitPDF
                Case EzShipLabelImageTypes.itPNG
                    objFedexShip.LabelImageType = FedExShipLabelImageTypes.fitPNG
                Case EzShipLabelImageTypes.itUniMark
                    objFedexShip.LabelImageType = FedExShipLabelImageTypes.fitUniMark
                Case EzShipLabelImageTypes.itZebra
                    objFedexShip.LabelImageType = FedExShipLabelImageTypes.fitZebra
                Case Else
                    objFedexShip.LabelImageType = FedExShipLabelImageTypes.fitEltron
            End Select

            ' 02/06/2024 - Always get ZPL
            objFedexShip.LabelImageType = FedExShipIntlLabelImageTypes.fitZebra

            Dim extension As String = objFedexShip.LabelImageType.ToString
            If extension.StartsWith("fit") Then
                extension = "." & extension.Substring(3)
            Else
                extension = String.Empty
            End If

            ' Set shipping directory to store the labels
            Dim idCtr As Int16 = 1
            If ShippingLabelDirectory.Length > 0 AndAlso ShippingLabelPrefix.Length > 0 Then
                If My.Computer.FileSystem.DirectoryExists(ShippingLabelDirectory) Then
                    If Not ShippingLabelDirectory.EndsWith("\") Then
                        ShippingLabelDirectory &= "\"
                    End If
                    For Each shippingPackageDetail In PackageDetailList
                        Dim id As String = idCtr.ToString
                        idCtr += 1
                        If Val(shippingPackageDetail.Id) > 0 Then
                            id = Val(shippingPackageDetail.Id)
                        End If

                        shippingPackageDetail.ShippingLabelFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & extension
                        shippingPackageDetail.CODFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & "_COD" & extension
                    Next
                End If
            End If

            ' Add packages (package weight is in Ounces - Convert to Pounds)
            Dim TotalWeight As Decimal = 0
            Dim totalInsured As Decimal = 0
            For Each shippingPackageDetail In PackageDetailList
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight), "###0.0")
                TotalWeight += Val(shippingPackageDetail.Weight & String.Empty)
                totalInsured += Val(shippingPackageDetail.InsuredValue & String.Empty)
                ' fedex require a direct sign if Insured Value > 500, or if signature is required
                If Val(shippingPackageDetail.InsuredValue & String.Empty) > 500 OrElse cSignatureRequired Then
                    shippingPackageDetail.SignatureType = TSignatureTypes.stDirect
                End If
                objFedexShip.Packages.Add(shippingPackageDetail)
            Next

            objFedexShip.TotalWeight = Format(TotalWeight, "###0.0")
            objFedexShip.InsuredValue = Format(totalInsured, "###0.00")

            With objFedexShip.Payor
                .PayorType = Payor
                .AccountNumber = PayorContact.AccountNumber
                .CountryCode = PayorContact.CountryCode
                .ZipCode = PayorContact.ZipCode
            End With

            '  Custom Content to place on a Fedex Label
            '  When using the CustomContent, the LabelStockType must be either 4 (Stock 4x8) or 5 (Stock 4x9 Leading Doc Tab). 
            '  Also LabelFormatType must be 0 (Common2D) and LabelImageType must 2 (fitEltron), 3 (fitZebra) or 4 (fitUniMark). 

            ' NS-HF048860852E
            If cFedexCustomContent.Length > 0 Then
                objFedexShip.Config("CustomContent=" & cFedexCustomContent)

                ' Need to check the labal type
                ' 4 = 4x8, 5=4x9 - has pull tab
                If LabelStockType <> "4" And LabelStockType <> "5" Then
                    LabelStockType = "4"
                End If

                If objFedexShip.LabelImageType <> FedExShipLabelImageTypes.fitEltron _
                    AndAlso objFedexShip.LabelImageType <> FedExShipLabelImageTypes.fitZebra _
                    AndAlso objFedexShip.LabelImageType <> FedExShipLabelImageTypes.fitUniMark Then
                    objFedexShip.LabelImageType = FedExShipLabelImageTypes.fitEltron
                End If
            End If

            If LabelStockType.Length > 0 Then
                'MessageBox.Show("LabelStockType=" & LabelStockType)
                objFedexShip.Config("LabelStockType=" & LabelStockType)
            End If

            Dim specialService As Long = ShipmentSpecialServices

            'If objFedexShipIntl.ShipDate > DateTime.Now.ToString("yyyy-MM-dd") Then
            '    specialService = specialService Or &H20000000L
            'End If

            objFedexShip.ShipmentSpecialServices = specialService

            ' Notifications
            Dim notificationsIndex As Int16 = 0
            If ShipmentNotifications.Count > 0 AndAlso Not ASCMAIN1.Running_in_VS Then
                For Each sn As Notifications In ShipmentNotifications
                    sn.email = (sn.email & String.Empty).Trim
                    If sn.email.Length = 0 Then
                        Continue For
                    End If

                    Dim notify As New NotifyDetail
                    With notify
                        .Email = sn.email
                        .NotificationFlags = CInt(sn.NotificationFlags)
                        .Message = (sn.Message & String.Empty).ToString.Trim
                    End With

                    objFedexShip.Notify.Add(notify)

                    notificationsIndex += 1
                    If notificationsIndex = 3 Then Exit For
                Next
            End If

            objFedexShip.RuntimeLicense = s4DPaymentsShippingSDK
            objFedexShip.GetShipmentLabels()

            ' Reset the object to have the updated data returned
            PackageDetailList.Clear()
            For ictr As Int16 = 0 To objFedexShip.Packages.Count - 1
                PackageDetailList.Add(objFedexShip.Packages(ictr))
                GetPackageCosts(objFedexShip.Packages(ictr), objFedexShip)

                lstShippingLabels.Add(objFedexShip.Packages(ictr).TrackingNumber, objFedexShip.Packages(ictr).ShippingLabel)
            Next

            If objFedexShip.Packages.Count = 1 Then
                cMasterTrackingNumber = objFedexShip.Packages(0).TrackingNumber
            Else
                cMasterTrackingNumber = objFedexShip.MasterTrackingNumber
            End If

            If cMasterTrackingNumber.Length = 0 Then
                LastError = "Shipper did not return a tracking number"
                Return False
            End If

            Dim smartShipTracking As String = objFedexShip.Config("SmartPostTrackingNumbers")
            For Each track As String In smartShipTracking.Split(",")
                If track.Length > 0 Then
                    FedexSmartPost.TrackingNumbers.Add(track)
                End If
            Next

            Return True

        Catch ex As DShippingSDKException
            LastError = ex.Message
            Return False
        Catch exc As Exception
            LastError = exc.Message
            Return False
        Finally
            cRawRequest = objFedexShip.Config("RawRequest")
            cRawResponse = objFedexShip.Config("RawResponse")
        End Try

        Return True

    End Function

    ''' <summary>
    ''' Close Ground Shipment for the day
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FedexCloseGroundShipments() As Boolean
        Try

            LastError = String.Empty
            objFedexShip.RuntimeLicense = s4DPaymentsShippingSDK
            objFedexShip.Reset()
            objFedexShip.RuntimeLicense = s4DPaymentsShippingSDK

            Dim bearertoken As String = GetAuthorizationToken(Carriers.FedEx, cAccountNumber)
            If inTestMode Then
                objFedexShip.Config("TESTMODE=true")
                cAccountNumber = "740561073"
            End If

            objFedexShip.FedExAccount.AccountNumber = cAccountNumber
            objFedexShip.FedExAccount.AuthorizationToken = bearertoken
            'objFedexShip.FedExAccount.MeterNumber = cFedexMeterNumber
            'objFedexShip.FedExAccount.DeveloperKey = cFedexDeveloperKey

            objFedexShip.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)

            With objFedexShip.CloseRequest
                .Date = FedexClose.Date
                .ReportFile = FedexClose.ReportFile
                .ReportType = FedexClose.ReportType
                '.Time = FedexClose.Time
            End With

            objFedexShip.CloseGroundShipments()
            Return True

        Catch ex As Exception
            LastError = ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Cancel / Void Shipment
    ''' </summary>
    ''' <param name="TrackingNumber"></param>
    ''' <param name="isMultiPackage"></param>
    ''' <param name="FedexTrackingIDType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CancelFedexShipment(ByVal TrackingNumber As String, ByVal isMultiPackage As Boolean, ByVal FedexTrackingIDType As Int16) As Boolean

        Try

            LastError = String.Empty
            objFedexShip.RuntimeLicense = s4DPaymentsShippingSDK
            objFedexShip.Reset()
            objFedexShip.RuntimeLicense = s4DPaymentsShippingSDK

            If cServiceProvider <> ServiceProviders.FederalExpress AndAlso cServiceProvider <> ServiceProviders.FederalExpressInternational Then
                LastError = "Invalid Service Type for Fedex shipment cancellation"
                Return False
            End If

            Dim bearertoken As String = GetAuthorizationToken(Carriers.FedEx, cAccountNumber)
            If inTestMode Then
                objFedexShip.Config("TESTMODE=true")
                cAccountNumber = "740561073"
            End If

            objFedexShip.FedExAccount.AccountNumber = cAccountNumber
            objFedexShip.FedExAccount.AuthorizationToken = bearertoken

            'objFedexShip.FedExAccount.Password = cPassword
            'objFedexShip.FedExAccount.MeterNumber = cFedexMeterNumber
            'objFedexShip.FedExAccount.DeveloperKey = cFedexDeveloperKey

            objFedexShip.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)

            If isMultiPackage Then
                objFedexShip.CancelShipment(TrackingNumber)
            Else
                objFedexShip.CancelPackage(TrackingNumber)
            End If

            Return True

        Catch ex As Exception
            LastError = ex.Message
            Return False
        End Try
    End Function

    Public Function GetFedexRates() As Decimal

        Try
            LastError = String.Empty
            objFedexRates.RuntimeLicense = s4DPaymentsShippingSDK
            objFedexRates.Reset()
            objFedexRates.RuntimeLicense = s4DPaymentsShippingSDK

            ' Set credentials
            Dim bearertoken As String = GetAuthorizationToken(Carriers.FedEx, cAccountNumber)

            objFedexRates.FedExAccount.AccountNumber = cAccountNumber
            objFedexRates.FedExAccount.AuthorizationToken = bearertoken

            ' Set credentials
            'objFedexRates.FedExAccount.DeveloperKey = cFedexDeveloperKey
            'objFedexRates.FedExAccount.Password = cPassword
            'objFedexRates.FedExAccount.MeterNumber = cFedexMeterNumber
            objFedexRates.RequestedService = cRequestedServiceType

            objFedexRates.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)

            If inTestMode Then
                With objFedexRates
                    .Config("TESTMODE=true")

                    .FedExAccount.AccountNumber = "740561073"
                    .FedExAccount.AuthorizationToken = bearertoken

                    cSenderContact.ZipCode = "90660" ' "07092"
                    cSenderContact.CountryCode = "US"

                    cRecipientContact.ZipCode = "10007" ' "07081"
                    cRecipientContact.CountryCode = "US"

                    .RequestedService = ServiceTypes.stUnspecified
                    .PickupType = FedExRatesPickupTypes.fptDropoffAtFedexLocation

                    Dim PackageDetail As New PackageDetail
                    With PackageDetail
                        .Weight = "1.5"
                    End With
                    PackageDetailList.Clear()
                    PackageDetailList.Add(PackageDetail)
                End With
            End If

            ' Get Sender Information
            With cSenderContact
                objFedexRates.SenderAddress.State = .State
                objFedexRates.SenderAddress.ZipCode = .ZipCode
                objFedexRates.SenderAddress.CountryCode = .CountryCode
                If objFedexRates.SenderAddress.CountryCode.Length = 0 Then
                    objFedexRates.SenderAddress.CountryCode = "US"
                End If

                If .IsResidental Then
                    objFedexRates.SenderAddress.AddressFlags = &H2 'Residential
                End If
            End With

            With cRecipientContact
                objFedexRates.RecipientAddress.State = .State
                objFedexRates.RecipientAddress.ZipCode = .ZipCode
                objFedexRates.RecipientAddress.CountryCode = .CountryCode
                If objFedexRates.RecipientAddress.CountryCode.Length = 0 Then
                    objFedexRates.RecipientAddress.CountryCode = "US"
                End If

                If .IsResidental Then
                    objFedexRates.RecipientAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objFedexRates.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

            Dim totalWeight As Double = 0
            Dim totalInsured As Decimal = 0
            ' Add packages
            For Each shippingPackageDetail In PackageDetailList
                ' Add packages (package weight is in Ounces - Convert to Pounds)
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight) / 16, "###0.0")
                totalWeight += Val(shippingPackageDetail.Weight)

                shippingPackageDetail.InsuredValue = Math.Abs(Val(shippingPackageDetail.InsuredValue & String.Empty))
                totalInsured += shippingPackageDetail.InsuredValue

                If cSignatureRequired Then
                    shippingPackageDetail.SignatureType = TSignatureTypes.stDirect
                End If

                objFedexRates.Packages.Add(shippingPackageDetail)
            Next

            objFedexRates.TotalWeight = Format(Val(totalWeight), "###0.0")
            objFedexRates.InsuredValue = Format(Val(totalInsured), "###0.00")

            objFedexRates.ShipmentSpecialServices = 0
            objFedexRates.Config("WeightUnit=LB")

            If IsDate(Me.ShipDate) Then
                objFedexRates.ShipDate = CDate(Me.ShipDate).ToString("yyyy-MM-dd")
            Else
                objFedexRates.ShipDate = DateTime.Now.ToString("yyyy-MM-dd")
            End If

            objFedexRates.ShipmentSpecialServices = ShipmentSpecialServices

            objFedexRates.RateType = &H2 Or &H1 ' FedexratesRateTypes.rtList
            objFedexRates.GetRates()

            ShipmentBaseCharge.Clear()
            ShipmentDiscountCharge.Clear()
            ShipmentListCharge.Clear()
            ShipmentNetCharge.Clear()
            ShipmentSurCharge.Clear()

            GetFedexRates = 0

            If objFedexRates.Config("Warning") = String.Empty OrElse objFedexRates.Services.Count > 0 Then
                For i As Integer = 0 To objFedexRates.Services.Count - 1
                    ShipmentBaseCharge.Add(i, Convert.ToDecimal(Val(objFedexRates.Services(i).AccountBaseCharge)))
                    ShipmentDiscountCharge.Add(i, Convert.ToDecimal(Val(objFedexRates.Services(i).AccountTotalDiscount)))
                    ShipmentListCharge.Add(i, Convert.ToDecimal(Val(objFedexRates.Services(i).ListBaseCharge)))
                    ShipmentNetCharge.Add(i, Convert.ToDecimal(Val(objFedexRates.Services(i).AccountNetCharge)))
                    ShipmentSurCharge.Add(i, Convert.ToDecimal(Val(objFedexRates.Services(i).AccountTotalSurcharge)))
                    GetFedexRates += Convert.ToDecimal(Val(objFedexRates.Services(i).ListBaseCharge))
                Next
            End If
        Catch ex As Exception

        End Try
    End Function

    Public Function GetFedExRatesList() As RateList()

        Try

            Dim requestedRateList As RateList()
            ReDim requestedRateList(1)

            LastError = String.Empty

            objFedexRates.RuntimeLicense = s4DPaymentsShippingSDK
            objFedexRates.Reset()
            objFedexRates.RuntimeLicense = s4DPaymentsShippingSDK

            ' Set credentials
            Dim bearertoken As String = GetAuthorizationToken(Carriers.FedEx, cAccountNumber)
            objFedexRates.FedExAccount.AuthorizationToken = bearertoken

            ' Set credentials
            'objFedexRates.FedExAccount.DeveloperKey = cFedexDeveloperKey
            'objFedexRates.FedExAccount.Password = cPassword
            objFedexRates.FedExAccount.AccountNumber = cAccountNumber
            'objFedexRates.FedExAccount.MeterNumber = cFedexMeterNumber
            objFedexRates.RequestedService = cRequestedServiceType

            If inTestMode Then
                With objFedexRates
                    .Config("TESTMODE=true")

                    .FedExAccount.AccountNumber = "740561073"
                    .FedExAccount.AuthorizationToken = bearertoken

                    cSenderContact.ZipCode = "90660" ' "07092"
                    cSenderContact.CountryCode = "US"

                    cRecipientContact.ZipCode = "10007" ' "07081"
                    cRecipientContact.CountryCode = "US"

                    '.RecipientAddress.AddressFlags = 2
                    .RequestedService = ServiceTypes.stUnspecified
                    '.PickupType = FedexratesPickupTypes.fptUseScheduledPickup

                    Dim PackageDetail As New PackageDetail
                    With PackageDetail
                        .Weight = "1.5"
                    End With
                    PackageDetailList.Clear()
                    PackageDetailList.Add(PackageDetail)

                End With
            End If

            objFedexRates.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)

            ' Get Sender Information
            With cSenderContact
                objFedexRates.SenderAddress.State = .State
                objFedexRates.SenderAddress.ZipCode = .ZipCode
                objFedexRates.SenderAddress.CountryCode = .CountryCode
                If objFedexRates.SenderAddress.CountryCode.Length = 0 Then
                    objFedexRates.SenderAddress.CountryCode = "US"
                End If

                If .IsResidental Then
                    objFedexRates.SenderAddress.AddressFlags = &H2 'Residential
                End If
            End With

            With cRecipientContact
                objFedexRates.RecipientAddress.State = .State
                objFedexRates.RecipientAddress.ZipCode = .ZipCode
                objFedexRates.RecipientAddress.CountryCode = .CountryCode
                If objFedexRates.RecipientAddress.CountryCode.Length = 0 Then
                    objFedexRates.RecipientAddress.CountryCode = "US"
                End If

                If .IsResidental Then
                    objFedexRates.RecipientAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objFedexRates.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

            Dim totalWeight As Double = 0
            Dim totalInsured As Decimal = 0
            ' Add packages
            For Each shippingPackageDetail In PackageDetailList
                ' Add packages (package weight is in Ounces - Convert to Pounds)
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight) / 16, "###0.0")
                totalWeight += Val(shippingPackageDetail.Weight)

                shippingPackageDetail.InsuredValue = Math.Abs(Val(shippingPackageDetail.InsuredValue & String.Empty))
                totalInsured += shippingPackageDetail.InsuredValue

                If cSignatureRequired Then
                    shippingPackageDetail.SignatureType = TSignatureTypes.stDirect
                End If

                objFedexRates.Packages.Add(shippingPackageDetail)
            Next

            'objFedexRates.TotalWeight = Format(Val(totalWeight), "###0.0")
            objFedexRates.InsuredValue = Format(Val(totalInsured), "###0.00")

            objFedexRates.ShipmentSpecialServices = 0
            'objFedexRates.Config("WeightUnit=LB")

            If IsDate(Me.ShipDate) Then
                objFedexRates.ShipDate = CDate(Me.ShipDate).ToString("yyyy-MM-dd")
            Else
                objFedexRates.ShipDate = DateTime.Now.ToString("yyyy-MM-dd")
            End If

            ShipmentBaseCharge.Clear()
            ShipmentDiscountCharge.Clear()
            ShipmentListCharge.Clear()
            ShipmentNetCharge.Clear()
            ShipmentSurCharge.Clear()

            objFedexRates.ShipmentSpecialServices = ShipmentSpecialServices
            objFedexRates.RateType = &H2 Or &H1

            ' Geting a resouce error, try it twice. May be the test environment
            Dim success As Boolean = True
            Try
                objFedexRates.GetRates()
            Catch ex As Exception
                success = False
            End Try

            If Not success Then
                objFedexRates.GetRates()
            End If

            ReDim requestedRateList(objFedexRates.Services.Count)
            If objFedexRates.Config("Warning") = String.Empty OrElse objFedexRates.Services.Count > 0 Then

                For iLoop As Integer = 0 To objFedexRates.Services.Count - 1
                    With requestedRateList(iLoop)
                        .ServiceType = objFedexRates.Services(iLoop).ServiceType
                        .ServiceTypeDescription = StrConv(objFedexRates.Services(iLoop).ServiceTypeDescription.Replace("_", " "), VbStrConv.ProperCase)
                        .AccountNetCharge = Val(objFedexRates.Services(iLoop).AccountNetCharge & String.Empty)
                        .DeliveryTime = objFedexRates.Services(iLoop).DeliveryTime
                        .ListNetCharge = Val(objFedexRates.Services(iLoop).ListNetCharge & String.Empty)

                        If inTestMode Then
                            If .AccountNetCharge = 0 Then
                                .AccountNetCharge = .ListNetCharge
                            End If
                        End If

                        Select Case objFedexRates.Services(iLoop).TransitTime
                            Case "ONE_DAY"
                                .TransitTime = "1"
                            Case "TWO_DAYS"
                                .TransitTime = "2"
                            Case "THREE_DAYS"
                                .TransitTime = "3"
                            Case "FOUR_DAYS"
                                .TransitTime = "4"
                            Case "FIVE_DAYS"
                                .TransitTime = "5"
                            Case "SIX_DAYS"
                                .TransitTime = "6"
                            Case "SEVEN_DAYS"
                                .TransitTime = "7"
                            Case "EIGHT_DAYS"
                                .TransitTime = "8"
                            Case "NINE_DAYS"
                                .TransitTime = "9"
                            Case "TEN_DAYS"
                                .TransitTime = "10"
                            Case Else
                                .TransitTime = objFedexRates.Services(iLoop).TransitTime
                        End Select

                        If .AccountNetCharge = 0 Then
                            .AccountNetCharge = .ListNetCharge
                        End If
                    End With
                Next
            End If

            Return requestedRateList

        Catch ex As Exception
            LastError = ex.Message
            Return Nothing
        Finally
            cRawRequest = objFedexRates.Config("RawRequest")
            cRawResponse = objFedexRates.Config("RawResponse")
        End Try

    End Function

#End Region

#Region "UPS"

    Public Enum AddressClassificationTypes
        Unknown = 0
        Commercial = 1
        Residental = 2
    End Enum

    Public Structure AddressValidationResponse
        Dim AddressIndex As Integer
        Dim ClassificationCode As AddressClassificationTypes
        Dim ClassificationDescription As String
        Dim Consignee As String
        Dim BuildingName As String
        Dim AddressLine1 As String
        Dim AddressLine2 As String
        Dim AddressLine3 As String
        Dim City As String
        Dim State As String
        Dim PostalCode As String
        Dim PostalCodeExtended As String
        Dim CountryCode As String
        Dim Selected As Boolean
    End Structure

    Public Function AddressValidation(ByVal AddressLine1 As String,
                                                ByVal AddressLine2 As String,
                                                ByVal AddressLine3 As String,
                                                ByVal AttentionName As String,
                                                ByVal City As String,
                                                ByVal State As String,
                                                ByVal PostalCode As String,
                                                ByVal CompanyName As String,
                                                ByRef avrAddressValidationResponse As List(Of AddressValidationResponse),
                                                ByRef errMsg As String) As Boolean


        Try
            LastError = String.Empty
            ValidateInputFieldLengths(AddressLine1, AddressLine2, AddressLine3, AttentionName, City, State, PostalCode, CompanyName)

            Dim upsAddrVal As New UPSAddress
            upsAddrVal.RuntimeLicense = s4DPaymentsShippingSDK
            objUpsShip.Reset()
            upsAddrVal.RuntimeLicense = s4DPaymentsShippingSDK

            With upsAddrVal
                .Address = New AddressDetail
                With upsAddrVal.Address
                    .Address1 = AddressLine1
                    .Address2 = AddressLine2
                    .City = City
                    .CountryCode = "US"
                    .State = State
                    .ZipCode = PostalCode
                End With
            End With

            upsAddrVal.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)

            Dim bearerToken As String = GetAuthorizationToken(Carriers.UPS, cAccountNumber)
            If inTestMode Then
                upsAddrVal.Config("TESTMODE=true")
            End If

            upsAddrVal.UPSAccount.AccountNumber = cAccountNumber
            upsAddrVal.UPSAccount.AuthorizationToken = bearerToken

            upsAddrVal.ValidateAddress()

            avrAddressValidationResponse = New List(Of AddressValidationResponse)
            For Each match As MatchDetail In upsAddrVal.Matches

                Dim resp As New AddressValidationResponse
                With resp
                    .AddressIndex = avrAddressValidationResponse.Count + 1
                    .AddressLine1 = match.Address1
                    .AddressLine2 = match.Address2
                    .AddressLine3 = ""
                    .City = match.City
                    .CountryCode = match.CountryCode
                    .State = match.State
                    .PostalCode = match.ZipCode
                End With

                avrAddressValidationResponse.Add(resp)

            Next

            Return True

        Catch ex As Exception
            LastError = $"Address Validation Error: {ex.Message}"
            Return False
        End Try

    End Function

    Public Function GetUPSRates() As Decimal

        Try

            LastError = String.Empty
            objUpsRates.RuntimeLicense = s4DPaymentsShippingSDK
            objUpsRates.Reset()
            objUpsRates.RuntimeLicense = s4DPaymentsShippingSDK
            GetUPSRates = 0

            ShipmentBaseCharge.Clear()
            ShipmentDiscountCharge.Clear()
            ShipmentListCharge.Clear()
            ShipmentNetCharge.Clear()
            ShipmentSurCharge.Clear()

            If cServiceProvider = ServiceProviders.Unknown Then
                LastError = "Unknown Service Type"
                Return False
            End If

            objUpsRates.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)

            Dim bearerToken As String = GetAuthorizationToken(Carriers.UPS, cAccountNumber)
            If inTestMode Then
                objUpsRates.Config("TESTMODE=true")
            End If

            objUpsRates.UPSAccount.AccountNumber = cAccountNumber
            objUpsRates.UPSAccount.AuthorizationToken = bearerToken
            objUpsRates.RequestedService = cRequestedServiceType

            objUpsRates.PickupType = UPSRatesPickupTypes.ptDailyPickup
            'objUpsRates.CustomerType = UpsratesCustomerTypes.ccRetail

            ' Insured Value is Positive, Declared Value is Negative
            Dim iPackage As Int16 = 0
            Dim totalWeight As Double = 0
            Dim totalInsured As Decimal = 0

            For Each shippingPackageDetail In PackageDetailList

                If cSignatureRequired Then
                    shippingPackageDetail.SignatureType = TSignatureTypes.stDirect
                End If

                objUpsRates.Packages.Add(shippingPackageDetail)

                shippingPackageDetail.InsuredValue = Format(Val(shippingPackageDetail.InsuredValue), "###0.00")
                If Val(shippingPackageDetail.InsuredValue & String.Empty) < 0 Then
                    shippingPackageDetail.InsuredValue = Math.Abs(Val(shippingPackageDetail.InsuredValue & String.Empty))
                    objUpsRates.Config("PackageDeclaredValueType[" & iPackage & "]=0")
                End If

                ' Format weight 
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight) / 16, "###0.0")
                totalWeight += Val(shippingPackageDetail.Weight)
                totalInsured = Val(shippingPackageDetail.InsuredValue)

                iPackage += 1
            Next

            objUpsRates.TotalWeight = Format(Val(totalWeight), "###0.0")

            With cSenderContact
                objUpsRates.SenderAddress.Address1 = .Address1
                objUpsRates.SenderAddress.Address2 = .Address2
                'objUpsRates.Config("SenderAddress3=" & .Address3)

                objUpsRates.SenderAddress.City = .City
                objUpsRates.SenderAddress.ZipCode = .ZipCode
                objUpsRates.SenderAddress.State = .State
                objUpsRates.SenderAddress.CountryCode = .CountryCode
                If objUpsRates.SenderAddress.CountryCode.ToUpper = "USA" Then
                    objUpsRates.SenderAddress.CountryCode = "US"
                End If

                If .IsResidental Then
                    objUpsRates.SenderAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objUpsRates.RecipientAddress.AddressFlags = &H1 'PO Box
                End If
            End With

            With cRecipientContact
                objUpsRates.RecipientAddress.Address1 = .Address1
                objUpsRates.RecipientAddress.Address2 = .Address2
                'objUpsRates.Config("RecipientAddress3=" & .Address3)

                objUpsRates.RecipientAddress.City = .City
                objUpsRates.RecipientAddress.ZipCode = .ZipCode
                objUpsRates.RecipientAddress.State = .State
                objUpsRates.RecipientAddress.CountryCode = .CountryCode
                If objUpsRates.RecipientAddress.CountryCode.ToUpper = "USA" Then
                    objUpsRates.RecipientAddress.CountryCode = "US"
                End If
            End With

            If IsDate(Me.ShipDate) Then
                objUpsRates.ShipDate = CDate(Me.ShipDate).ToString("yyyyMMdd")
            Else
                objUpsRates.ShipDate = DateTime.Now.ToString("yyyyMMdd")
            End If

            objUpsRates.ShipmentSpecialServices = ShipmentSpecialServices
            objUpsRates.GetRates()

            For i As Integer = 0 To objUpsRates.Services.Count - 1
                ShipmentBaseCharge.Add(i, Convert.ToDecimal(Val(objUpsRates.Services(i).AccountBaseCharge)))
                ShipmentDiscountCharge.Add(i, Convert.ToDecimal(Val(objUpsRates.Services(i).AccountTotalDiscount)))
                ShipmentListCharge.Add(i, Convert.ToDecimal(Val(objUpsRates.Services(i).ListBaseCharge)))
                ShipmentNetCharge.Add(i, Convert.ToDecimal(Val(objUpsRates.Services(i).AccountNetCharge)))
                ShipmentSurCharge.Add(i, Convert.ToDecimal(Val(objUpsRates.Services(i).AccountTotalSurcharge)))
                GetUPSRates += Convert.ToDecimal(Val(objUpsRates.Services(i).ListBaseCharge))
            Next

        Catch ex As Exception
            LastError = ex.Message
        End Try

    End Function

    ''' <summary>
    ''' Request Fedex International Shipping label
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RequestUPSInternaltionalLabel() As Boolean

        Try
            LastError = String.Empty
            objUpsShipIntl.RuntimeLicense = s4DPaymentsShippingSDK
            objUpsShipIntl.Reset()
            objUpsShipIntl.RuntimeLicense = s4DPaymentsShippingSDK

            If cServiceProvider <> ServiceProviders.FederalExpressInternational Then
                LastError = "Service type not a Fedex International."
                Return False
            End If

            Dim bearerToken As String = GetAuthorizationToken(Carriers.UPS, cAccountNumber)
            If inTestMode Then
                objUpsShipIntl.Config("TESTMODE=true")
            End If

            objUpsShipIntl.UPSAccount.AccountNumber = cAccountNumber
            objUpsShipIntl.UPSAccount.AuthorizationToken = bearerToken
            objUpsShipIntl.ServiceType = cRequestedServiceType
            objUpsShipIntl.ShipmentDescription = ShipmentDescription

            objUpsShipIntl.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)

            ' Get Sender Information
            With cSenderContact
                objUpsShipIntl.SenderContact.FirstName = .FirstName
                objUpsShipIntl.SenderContact.LastName = .LastName
                objUpsShipIntl.SenderContact.MiddleInitial = .MiddleInitial
                objUpsShipIntl.SenderContact.Phone = .Phone
                objUpsShipIntl.SenderContact.Fax = .Fax
                objUpsShipIntl.SenderContact.Email = .eMail

                objUpsShipIntl.SenderContact.Company = .Company
                objUpsShipIntl.SenderAddress.Address1 = .Address1
                objUpsShipIntl.SenderAddress.Address2 = .Address2
                objUpsShipIntl.Config("SenderAddress3=" & .Address3)

                objUpsShipIntl.SenderAddress.City = .City
                objUpsShipIntl.SenderAddress.ZipCode = .ZipCode
                objUpsShipIntl.SenderAddress.State = .State
                objUpsShipIntl.SenderAddress.CountryCode = .CountryCode

                If .IsResidental Then
                    objUpsShipIntl.SenderAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objUpsShipIntl.SenderAddress.AddressFlags = &H1 'PO Box
                End If
            End With

            With cRecipientContact
                objUpsShipIntl.RecipientContact.FirstName = .FirstName
                objUpsShipIntl.RecipientContact.LastName = .LastName
                objUpsShipIntl.RecipientContact.MiddleInitial = .MiddleInitial
                objUpsShipIntl.RecipientContact.Phone = .Phone
                objUpsShipIntl.RecipientContact.Fax = .Fax
                objUpsShipIntl.RecipientContact.Email = .eMail

                objUpsShipIntl.RecipientContact.Company = .Company
                objUpsShipIntl.RecipientAddress.Address1 = .Address1
                objUpsShipIntl.RecipientAddress.Address2 = .Address2
                objUpsShipIntl.Config("RecipientAddress3=" & .Address3)

                objUpsShipIntl.RecipientAddress.City = .City
                objUpsShipIntl.RecipientAddress.ZipCode = .ZipCode
                objUpsShipIntl.RecipientAddress.State = .State
                objUpsShipIntl.RecipientAddress.CountryCode = .CountryCode

                If .IsResidental Then
                    objUpsShipIntl.RecipientAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objUpsShipIntl.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

            Select Case EzshipLabelImage
                Case EzShipLabelImageTypes.itEPL
                    objUpsShipIntl.LabelImageType = UPSShipLabelImageTypes.uitEPL
                Case EzShipLabelImageTypes.itGIF
                    objUpsShipIntl.LabelImageType = UPSShipLabelImageTypes.uitGIF
                Case EzShipLabelImageTypes.itSPL
                    objUpsShipIntl.LabelImageType = UPSShipLabelImageTypes.uitSPL
                Case EzShipLabelImageTypes.itZPL
                    objUpsShipIntl.LabelImageType = UPSShipLabelImageTypes.uitZPL
                Case Else
                    objUpsShipIntl.LabelImageType = UPSShipLabelImageTypes.uitEPL
            End Select

            ' 02/06/2024 - Always get ZPL
            objUpsShipIntl.LabelImageType = UPSShipLabelImageTypes.uitZPL

            Dim extension As String = objUpsShipIntl.LabelImageType.ToString
            If extension.StartsWith("uit") Then
                extension = "." & extension.Substring(3)
            Else
                extension = String.Empty
            End If

            ' Set Shipping Label File
            Dim idCtr As Int16 = 1
            If ShippingLabelDirectory.Length > 0 AndAlso ShippingLabelPrefix.Length > 0 Then
                If Not ShippingLabelDirectory.EndsWith("\") Then
                    ShippingLabelDirectory &= "\"
                End If

                For Each shippingPackageDetail In PackageDetailList
                    Dim id As String = idCtr.ToString
                    idCtr += 1
                    shippingPackageDetail.ShippingLabelFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & extension
                    shippingPackageDetail.CODFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & "_COD" & extension
                Next
            End If

            Dim totalWeight As Double = 0
            Dim totalInsured As Double = 0

            ' Add packages
            For Each shippingPackageDetail In PackageDetailList
                ' Add packages (package weight is in Ounces - Convert to Pounds)
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight) / 16, "###0.0")
                totalWeight += Val(shippingPackageDetail.Weight)
                totalInsured += Val(shippingPackageDetail.InsuredValue & String.Empty)
                If cSignatureRequired Then
                    shippingPackageDetail.SignatureType = TSignatureTypes.stDirect
                End If
                objUpsShipIntl.Packages.Add(shippingPackageDetail)
            Next

            'objUpsShipIntl.TotalWeight = Format(totalWeight, "###0.0")
            'objUpsShipIntl.InsuredValue = Format(totalInsured, "###0.00")
            objUpsShipIntl.ShipDate = ShipDate.ToString("yyyy-MM-dd")
            objUpsShipIntl.TotalCustomsValue = Format(TotalCustomsValue, "###0.00")

            ' Service Type
            objUpsShipIntl.ServiceType = cRequestedServiceType
            'objUpsShipIntl.DropoffType = DropOffType
            objUpsShipIntl.Payor.PayorType = Payor
            objUpsShipIntl.Payor.AccountNumber = cSenderContact.AccountNumber
            objUpsShipIntl.Payor.CountryCode = cSenderContact.CountryCode

            objUpsShipIntl.DutiesPayor.PayorType = DutiesPayor
            objUpsShipIntl.DutiesPayor.AccountNumber = cDutiesPayorContact.AccountNumber
            objUpsShipIntl.DutiesPayor.CountryCode = cDutiesPayorContact.CountryCode

            Dim specialService As Long = ShipmentSpecialServices

            With objUpsShipIntl.Payor
                .PayorType = Payor
                .AccountNumber = PayorContact.AccountNumber
                .CountryCode = PayorContact.CountryCode
                .ZipCode = PayorContact.ZipCode
            End With

            With objUpsShipIntl.DutiesPayor
                .PayorType = DutiesPayor
                .AccountNumber = DutiesPayorContact.AccountNumber
                .CountryCode = DutiesPayorContact.CountryCode
                .ZipCode = DutiesPayorContact.ZipCode
            End With

            If objUpsShipIntl.ShipDate > DateTime.Now.ToString("yyyy-MM-dd") Then
                specialService = specialService Or &H20000000L
            End If

            objUpsShipIntl.ShipmentSpecialServices = specialService

            'objUpsShipIntl.HoldAtLocation.Address1 = cHoldAtLocation.Address1
            'objUpsShipIntl.HoldAtLocation.Address2 = cHoldAtLocation.Address2
            'objUpsShipIntl.HoldAtLocation.City = cHoldAtLocation.City
            'objUpsShipIntl.HoldAtLocation.State = cHoldAtLocation.State
            'objUpsShipIntl.HoldAtLocation.ZipCode = cHoldAtLocation.ZipCode
            'objUpsShipIntl.HoldAtLocationPhone = cHoldAtLocation.Phone

            For Each CommDetail As CommodityDetail In CommodityDetailList
                CommDetail.Weight = Format(Val(CommDetail.Weight), "###0.0")
                CommDetail.Description = CommDetail.Description.Replace("&", " ").Replace("<", " ").Replace(">", " ")
                objUpsShipIntl.Commodities.Add(CommDetail)
            Next

        Catch ex As Exception
            LastError = ex.Message
            Return False
        End Try

        ' Notifications
        Dim notificationsIndex As Int16 = 0
        If ShipmentNotifications.Count > 0 AndAlso Not ASCMAIN1.Running_in_VS Then
            For Each sn As Notifications In ShipmentNotifications
                sn.email = (sn.email & String.Empty).Trim
                If sn.email.Length = 0 Then
                    Continue For
                End If

                Dim notify As New NotifyDetail
                With notify
                    .Email = sn.email
                    .NotificationFlags = CInt(sn.NotificationFlags)
                    .Message = (sn.Message & String.Empty).ToString.Trim
                End With

                objUpsShipIntl.Notify.Add(notify)

                notificationsIndex += 1
                If notificationsIndex = 3 Then Exit For
            Next
        End If

        Try
            objUpsShipIntl.RuntimeLicense = s4DPaymentsShippingSDK
            objUpsShipIntl.GetShipmentLabels()

            ' Reset the object to have the updated data returned
            ' For multi UPS package shipments the total cost exists in all packages
            ' so spread the costs
            PackageDetailList.Clear()
            For ictr As Int16 = 0 To objUpsShipIntl.Packages.Count - 1
                PackageDetailList.Add(objUpsShipIntl.Packages(ictr))
                GetPackageCosts(objUpsShipIntl.Packages(ictr), objUpsShipIntl)
                ShipmentBaseCharge(ictr) = ShipmentBaseCharge(ictr) / objUpsShipIntl.Packages.Count
                ShipmentDiscountCharge(ictr) = ShipmentDiscountCharge(ictr) / objUpsShipIntl.Packages.Count
                ShipmentSurCharge(ictr) = ShipmentSurCharge(ictr) / objUpsShipIntl.Packages.Count
                ShipmentNetCharge(ictr) = ShipmentNetCharge(ictr) / objUpsShipIntl.Packages.Count
                ShipmentListCharge(ictr) = ShipmentListCharge(ictr) / objUpsShipIntl.Packages.Count
            Next

            If objUpsShipIntl.Packages.Count = 1 Then
                cMasterTrackingNumber = objUpsShipIntl.Packages(0).TrackingNumber
            Else
                cMasterTrackingNumber = objUpsShipIntl.MasterTrackingNumber
            End If

            Return True

        Catch ex As DShippingSDKException
            LastError = ex.Message
            Return False
        Catch exc As Exception
            LastError = exc.Message
            Return False
        Finally
            cRawRequest = objUpsShipIntl.Config("RawRequest")
            cRawResponse = objUpsShipIntl.Config("RawResponse")
        End Try

        Return True

    End Function

    ''' <summary>
    ''' Request Shipping label. Not used for UPS Intenational
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RequestUPSLabel() As Boolean

        Try
            LastError = String.Empty
            objUpsShip.RuntimeLicense = s4DPaymentsShippingSDK
            objUpsShip.Reset()
            objUpsShip.RuntimeLicense = s4DPaymentsShippingSDK

            If cServiceProvider = ServiceProviders.Unknown Then
                LastError = "Unknown Service Type"
                Return False
            End If

            objUpsShip.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            Dim bearerToken As String = GetAuthorizationToken(Carriers.UPS, cAccountNumber)
            If inTestMode Then
                objUpsShip.Config("TESTMODE=true")
            End If

            objUpsShip.UPSAccount.AccountNumber = cAccountNumber
            objUpsShip.UPSAccount.AuthorizationToken = bearerToken
            objUpsShip.ServiceType = cRequestedServiceType

            objUpsShip.ShipDate = ShipDate.ToString("yyyyMMdd")

            With cSenderContact
                objUpsShip.SenderContact.FirstName = .FirstName
                objUpsShip.SenderContact.LastName = .LastName
                objUpsShip.SenderContact.MiddleInitial = .MiddleInitial
                objUpsShip.SenderContact.Phone = .Phone
                objUpsShip.SenderContact.Fax = .Fax
                objUpsShip.SenderContact.Email = .eMail

                objUpsShip.SenderContact.Company = .Company
                objUpsShip.SenderAddress.Address1 = .Address1
                objUpsShip.SenderAddress.Address2 = .Address2
                objUpsShip.Config("SenderAddress3=" & .Address3)

                objUpsShip.SenderAddress.City = .City
                objUpsShip.SenderAddress.ZipCode = .ZipCode
                objUpsShip.SenderAddress.State = .State
                objUpsShip.SenderAddress.CountryCode = .CountryCode

                objUpsShip.SenderAddress.CountryCode = objUpsShip.SenderAddress.CountryCode.ToUpper
                If objUpsShip.SenderAddress.CountryCode = "" OrElse objUpsShip.SenderAddress.CountryCode = "USA" Then
                    objUpsShip.SenderAddress.CountryCode = "US"
                End If

                If .IsResidental Then
                    objUpsShip.SenderAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objUpsShip.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

            With cRecipientContact
                objUpsShip.RecipientContact.FirstName = .FirstName
                objUpsShip.RecipientContact.LastName = .LastName
                objUpsShip.RecipientContact.MiddleInitial = .MiddleInitial
                objUpsShip.RecipientContact.Phone = .Phone
                objUpsShip.RecipientContact.Fax = .Fax
                objUpsShip.RecipientContact.Email = .eMail

                objUpsShip.RecipientContact.Company = .Company
                objUpsShip.RecipientAddress.Address1 = .Address1
                objUpsShip.RecipientAddress.Address2 = .Address2
                objUpsShip.Config("RecipientAddress3=" & .Address3)

                objUpsShip.RecipientAddress.City = .City
                objUpsShip.RecipientAddress.ZipCode = .ZipCode
                objUpsShip.RecipientAddress.State = .State
                objUpsShip.RecipientAddress.CountryCode = .CountryCode

                objUpsShip.RecipientAddress.CountryCode = objUpsShip.RecipientAddress.CountryCode.ToUpper
                If objUpsShip.RecipientAddress.CountryCode = "" OrElse objUpsShip.RecipientAddress.CountryCode = "USA" Then
                    objUpsShip.RecipientAddress.CountryCode = "US"
                End If


                If .IsResidental Then
                    objUpsShip.RecipientAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objUpsShip.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

        Catch ex As Exception
            LastError = ex.Message
            Return False
        End Try

        Try
            Select Case EzshipLabelImage
                Case EzShipLabelImageTypes.itEPL
                    objUpsShip.LabelImageType = UPSShipLabelImageTypes.uitEPL
                Case EzShipLabelImageTypes.itGIF
                    objUpsShip.LabelImageType = UPSShipLabelImageTypes.uitGIF
                Case EzShipLabelImageTypes.itSPL
                    objUpsShip.LabelImageType = UPSShipLabelImageTypes.uitSPL
                Case EzShipLabelImageTypes.itZPL
                    objUpsShip.LabelImageType = UPSShipLabelImageTypes.uitZPL
                Case Else
                    objUpsShip.LabelImageType = UPSShipLabelImageTypes.uitEPL
            End Select

            ' 02/06/2024 - Always get ZPL
            If objUpsShip.LabelImageType <> UPSShipLabelImageTypes.uitGIF Then
                objUpsShip.LabelImageType = UPSShipLabelImageTypes.uitZPL
            End If

            Dim extension As String = objUpsShip.LabelImageType.ToString
            If extension.StartsWith("uit") Then
                extension = "." & extension.Substring(3)
            Else
                extension = "." & extension
            End If

            ' Set shipping directory to store the labels
            Dim idCtr As Int16 = 1
            If ShippingLabelDirectory.Length > 0 AndAlso ShippingLabelPrefix.Length > 0 Then
                If My.Computer.FileSystem.DirectoryExists(ShippingLabelDirectory) Then
                    If Not ShippingLabelDirectory.EndsWith("\") Then
                        ShippingLabelDirectory &= "\"
                    End If
                    For Each shippingPackageDetail In PackageDetailList
                        Dim id As String = idCtr.ToString
                        idCtr += 1
                        If Val(shippingPackageDetail.Id) > 0 Then
                            id = Val(shippingPackageDetail.Id)
                        End If

                        shippingPackageDetail.ShippingLabelFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & extension
                        shippingPackageDetail.CODFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & "_COD" & extension

                        Try
                            If My.Computer.FileSystem.FileExists(shippingPackageDetail.ShippingLabelFile) Then
                                My.Computer.FileSystem.DeleteFile(shippingPackageDetail.ShippingLabelFile)
                            End If
                        Catch ex As Exception

                        End Try

                        Try
                            If My.Computer.FileSystem.FileExists(shippingPackageDetail.CODFile) Then
                                My.Computer.FileSystem.DeleteFile(shippingPackageDetail.CODFile)
                            End If
                        Catch ex As Exception

                        End Try
                    Next
                End If
            End If

            ' Add packages 
            Dim TotalWeight As Decimal = 0
            Dim totalInsured As Decimal = 0
            For Each shippingPackageDetail In PackageDetailList
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight), "###0.0")
                shippingPackageDetail.InsuredValue = Format(Val(shippingPackageDetail.InsuredValue), "###0.00")
                TotalWeight += shippingPackageDetail.Weight
                totalInsured += Val(shippingPackageDetail.InsuredValue & String.Empty)
                If cSignatureRequired Then
                    shippingPackageDetail.SignatureType = TSignatureTypes.stDirect
                End If
                objUpsShip.Packages.Add(shippingPackageDetail)
            Next

            With objUpsShip.Payor
                .PayorType = Payor
                .AccountNumber = PayorContact.AccountNumber
                .CountryCode = PayorContact.CountryCode
                .ZipCode = PayorContact.ZipCode
            End With

            objUpsShip.ShipmentSpecialServices = ShipmentSpecialServices

            objUpsShip.RuntimeLicense = s4DPaymentsShippingSDK

            ' Notifications
            Dim notificationsIndex As Int16 = 0
            If ShipmentNotifications.Count > 0 AndAlso Not ASCMAIN1.Running_in_VS Then
                For Each sn As Notifications In ShipmentNotifications
                    sn.email = (sn.email & String.Empty).Trim
                    If sn.email.Length = 0 Then
                        Continue For
                    End If

                    Dim notify As New NotifyDetail
                    With notify
                        .Email = sn.email
                        .NotificationFlags = CInt(sn.NotificationFlags)
                        .Message = (sn.Message & String.Empty).ToString.Trim
                    End With

                    objUpsShip.Notify.Add(notify)

                    notificationsIndex += 1
                    If notificationsIndex = 3 Then Exit For
                Next
            End If

            objUpsShip.GetShipmentLabels()

            ' Reset the object to have the updated data returned
            ' For multi UPS package shipments the total cost exists in all packages
            ' so spread the costs
            PackageDetailList.Clear()
            For ictr As Int16 = 0 To objUpsShip.Packages.Count - 1
                PackageDetailList.Add(objUpsShip.Packages(ictr))
                GetPackageCosts(objUpsShip.Packages(ictr), objUpsShip)
                Dim key As Integer = Val(objUpsShip.Packages(ictr).Id)
                ShipmentBaseCharge(key) = Math.Round(ShipmentBaseCharge(key) / objUpsShip.Packages.Count, 2)
                ShipmentDiscountCharge(key) = Math.Round(ShipmentDiscountCharge(key) / objUpsShip.Packages.Count, 2)
                ShipmentSurCharge(key) = Math.Round(ShipmentSurCharge(key) / objUpsShip.Packages.Count, 2)
                ShipmentNetCharge(key) = Math.Round(ShipmentNetCharge(key) / objUpsShip.Packages.Count, 2)
                ShipmentListCharge(key) = Math.Round(ShipmentListCharge(key) / objUpsShip.Packages.Count, 2)

                If Not inTestMode Then
                    lstShippingLabels.Add(objUpsShip.Packages(ictr).TrackingNumber, objUpsShip.Packages(ictr).ShippingLabel)
                    lstShippingLabels.Add(objUpsShip.Packages(ictr).TrackingNumber & "_COD", objUpsShip.Packages(ictr).CODLabel)

                    lstShippingFiles.Add(objUpsShip.Packages(ictr).TrackingNumber, objUpsShip.Packages(ictr).ShippingLabelFile)
                    lstShippingFiles.Add(objUpsShip.Packages(ictr).TrackingNumber & "_COD", objUpsShip.Packages(ictr).CODFile)

                Else
                    lstShippingLabels.Add(objUpsShip.Packages(ictr).TrackingNumber & "_" & ictr, objUpsShip.Packages(ictr).ShippingLabel)
                    lstShippingLabels.Add(objUpsShip.Packages(ictr).TrackingNumber & "_COD" & "_" & ictr, objUpsShip.Packages(ictr).CODLabel)

                    lstShippingFiles.Add(objUpsShip.Packages(ictr).TrackingNumber & "_" & ictr, objUpsShip.Packages(ictr).ShippingLabelFile)
                    lstShippingFiles.Add(objUpsShip.Packages(ictr).TrackingNumber & "_COD" & "_" & ictr, objUpsShip.Packages(ictr).CODFile)

                End If
            Next

            If objUpsShip.Packages.Count = 1 Then
                cMasterTrackingNumber = objUpsShip.Packages(0).TrackingNumber
            Else
                cMasterTrackingNumber = objUpsShip.MasterTrackingNumber
            End If

            If cMasterTrackingNumber.Length = 0 Then
                LastError = "Shipper did not return a tracking number"
                Return False
            End If

            Return True

        Catch ex As DShippingSDKException
            LastError = ex.Message
            Return False
        Catch exc As Exception
            LastError = exc.Message
            Return False
        Finally
            cRawRequest = objUpsShip.Config("RawRequest")
            cRawResponse = objUpsShip.Config("RawResponse")
        End Try

        Return True

    End Function

    ''' <summary>
    ''' Cancel / Void UPS Shipment / Package
    ''' </summary>
    ''' <param name="TrackingNumber"></param>
    ''' <param name="isMultiPackage"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CancelUpsShipment(ByVal TrackingNumber As String, ByVal isMultiPackage As Boolean) As Boolean

        Try

            LastError = String.Empty
            objUpsShip.RuntimeLicense = s4DPaymentsShippingSDK
            objUpsShip.Reset()
            objUpsShip.RuntimeLicense = s4DPaymentsShippingSDK

            If cServiceProvider <> ServiceProviders.UPS Then
                LastError = "Invalid Service Type for UPS shipment cancellation"
                Return False
            End If

            Dim bearerToken As String = GetAuthorizationToken(Carriers.UPS, cAccountNumber)
            If inTestMode Then
                objUpsShip.Config("TESTMODE=true")
            End If

            objUpsShip.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)

            objUpsShip.UPSAccount.AccountNumber = cAccountNumber
            objUpsShip.UPSAccount.AuthorizationToken = bearerToken

            If isMultiPackage Then
                objUpsShip.CancelShipment(TrackingNumber)
            Else
                objUpsShip.CancelPackage(TrackingNumber, TrackingNumber)
            End If

            Return True
        Catch ex As Exception
            LastError = ex.Message
            Return False
        End Try

    End Function

    Private Function RequestUpsFreightlabel() As Boolean

        Try
            LastError = String.Empty
            objUpsFreight.RuntimeLicense = s4DPaymentsShippingSDK
            objUpsFreight.Reset()
            objUpsFreight.RuntimeLicense = s4DPaymentsShippingSDK

            If cServiceProvider = ServiceProviders.Unknown Then
                LastError = "Unknown Service Type"
                Return False
            End If

            Dim bearerToken As String = GetAuthorizationToken(Carriers.UPS, cAccountNumber)
            If inTestMode Then
                objUpsRates.Config("TESTMODE=true")
            End If

            objUpsFreight.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            objUpsFreight.UPSAccount.AccountNumber = cAccountNumber
            objUpsFreight.UPSAccount.AuthorizationToken = bearerToken
            objUpsFreight.ServiceType = cRequestedServiceType

            'objUpsFreight.ShipDate = ShipDate.ToString("yyyy-MM-dd")

            With cSenderContact
                objUpsFreight.SenderContact.Company = .Company
                objUpsFreight.SenderAddress.Address1 = .Address1
                objUpsFreight.SenderAddress.Address2 = .Address2
                objUpsFreight.Config("SenderAddress3=" & .Address3)

                objUpsFreight.SenderAddress.City = .City
                objUpsFreight.SenderAddress.ZipCode = .ZipCode
                objUpsFreight.SenderAddress.State = .State
                objUpsFreight.SenderAddress.CountryCode = .CountryCode
            End With

            With cRecipientContact
                objUpsFreight.RecipientContact.Company = .Company
                objUpsFreight.RecipientAddress.Address1 = .Address1
                objUpsFreight.RecipientAddress.Address2 = .Address2
                objUpsFreight.Config("RecipientAddress3=" & .Address3)

                objUpsFreight.RecipientAddress.City = .City
                objUpsFreight.RecipientAddress.ZipCode = .ZipCode
                objUpsFreight.RecipientAddress.State = .State
                objUpsFreight.RecipientAddress.CountryCode = .CountryCode
            End With

            With objUpsFreight.Payor
                .PayorType = Payor
                .AccountNumber = PayorContact.AccountNumber
                .CountryCode = PayorContact.CountryCode
                .ZipCode = PayorContact.ZipCode
            End With

            For Each commDetail As CommodityDetail In CommodityDetailList
                objUpsFreight.Commodities.Add(commDetail)
            Next

            objUpsFreight.HandlingUnit = HandlingUnit
            Dim docLabelFileName As String = String.Empty
            If ShippingLabelDirectory.Length > 0 AndAlso ShippingLabelPrefix.Length > 0 Then
                If My.Computer.FileSystem.DirectoryExists(ShippingLabelDirectory) Then
                    If Not ShippingLabelDirectory.EndsWith("\") Then
                        ShippingLabelDirectory &= "\"
                    End If

                    docLabelFileName = ShippingLabelDirectory & ShippingLabelPrefix & "_doc" & ".zpl"
                End If
            End If

            Dim docLabel As New DocumentInfo
            docLabel.FileName = docLabelFileName
            docLabel.PrintFormat = TFreightPrintFormats.fpfThermal
            docLabel.PrintSize = TFreightPrintSizes.fpsSize4X6
            docLabel.DocumentType = TFreightDocumentTypes.ftcLabel
            objUpsFreight.Documents.Add(docLabel)

            objUpsFreight.RuntimeLicense = s4DPaymentsShippingSDK
            objUpsFreight.GetShipmentDocuments()

            Return True

        Catch ex As DShippingSDKException
            LastError = ex.Message
            Return False
        Catch exc As Exception
            LastError = exc.Message
            Return False
        Finally
            cRawRequest = objUpsFreight.Config("RawRequest")
            cRawResponse = objUpsFreight.Config("RawResponse")
        End Try

        Return True

    End Function

    Public Function GetUPSShippingTime() As List(Of TransTimeList)
        Try

            LastError = String.Empty
            Dim lstTransTimeList As New List(Of TransTimeList)

            objUpsRates.RuntimeLicense = s4DPaymentsShippingSDK
            objUpsRates.Reset()
            objUpsRates.RuntimeLicense = s4DPaymentsShippingSDK

            Dim bearerToken As String = GetAuthorizationToken(Carriers.UPS, cAccountNumber)
            If inTestMode Then
                objUpsRates.Config("TESTMODE=true")
            End If

            objUpsRates.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)

            objUpsRates.UPSAccount.AccountNumber = cAccountNumber
            objUpsRates.UPSAccount.AuthorizationToken = bearerToken

            objUpsRates.RequestedService = cRequestedServiceType
            objUpsRates.PickupType = UPSPickupType
            objUpsRates.CustomerType = cCustomerType

            ' Insured Value is Positive, Declared Value is Negative
            Dim iPackage As Int16 = 0
            Dim totalWeight As Double = 0
            Dim totalInsured As Decimal = 0

            For Each shippingPackageDetail In PackageDetailList

                If cSignatureRequired Then
                    shippingPackageDetail.SignatureType = TSignatureTypes.stDirect
                End If

                objUpsRates.Packages.Add(shippingPackageDetail)

                shippingPackageDetail.InsuredValue = Format(Val(shippingPackageDetail.InsuredValue), "###0.00")
                If Val(shippingPackageDetail.InsuredValue & String.Empty) < 0 Then
                    shippingPackageDetail.InsuredValue = Math.Abs(Val(shippingPackageDetail.InsuredValue & String.Empty))
                    objUpsRates.Config("PackageDeclaredValueType[" & iPackage & "]=0")
                End If

                ' Format weight 
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight) / 16, "###0.0")
                totalWeight += Val(shippingPackageDetail.Weight)
                totalInsured = Val(shippingPackageDetail.InsuredValue)

                iPackage += 1
            Next

            objUpsRates.TotalWeight = Format(Val(totalWeight), "###0.0")

            With cSenderContact
                objUpsRates.SenderAddress.ZipCode = .ZipCode
                objUpsRates.SenderAddress.State = .State
                objUpsRates.SenderAddress.CountryCode = .CountryCode
            End With

            With cRecipientContact
                objUpsRates.RecipientAddress.ZipCode = .ZipCode

                If .CountryCode = "US" And .State = "PR" Then
                    objUpsRates.RecipientAddress.State = .State
                    objUpsRates.RecipientAddress.CountryCode = .State
                ElseIf .CountryCode = "US" Then
                    objUpsRates.RecipientAddress.State = .State
                    objUpsRates.RecipientAddress.CountryCode = .CountryCode
                ElseIf .CountryCode <> "US" Then
                    objUpsRates.RecipientAddress.State = .State
                    objUpsRates.RecipientAddress.CountryCode = .CountryCode
                End If

                If .IsResidental Then
                    objUpsRates.RecipientAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objUpsRates.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

            If IsDate(Me.ShipDate) Then
                objUpsRates.ShipDate = CDate(Me.ShipDate).ToString("yyyy-MM-dd")
            Else
                objUpsRates.ShipDate = DateTime.Now.ToString("yyyy-MM-dd")
            End If

            objUpsRates.ShipmentSpecialServices = ShipmentSpecialServices

            objUpsRates.GetShippingTime()

            If objUpsRates.Services IsNot Nothing Then
                For Each sd As ServiceDetail In objUpsRates.Services
                    Dim jsonResulttodict = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(sd.Aggregate)
                    Dim item As New TransTimeList
                    item.ServiceCode = jsonResulttodict.Item("serviceLevel")
                    item.ServiceDescription = jsonResulttodict.Item("serviceLevelDescription")
                    item.TransitDays = Val(jsonResulttodict.Item("businessTransitDays") & String.Empty)

                    If jsonResulttodict.Item("guaranteeIndicator") & String.Empty = "1" Then
                        item.Guaranteed = "Yes"
                    Else
                        item.Guaranteed = "No"
                    End If
                    lstTransTimeList.Add(item)
                Next
            End If

            Return lstTransTimeList

        Catch ex As Exception
            LastError = "GetUPSShippingTime Error:" & ex.Message
            Return New List(Of TransTimeList)
        Finally
            cRawRequest = objUpsRates.Config("RawRequest")
            cRawResponse = objUpsRates.Config("RawResponse")
        End Try

    End Function

    Public Function GetUPSRatesList() As RateList()

        Try

            Dim requestedRateList As RateList()
            ReDim requestedRateList(1)
            LastError = String.Empty

            objUpsRates.RuntimeLicense = s4DPaymentsShippingSDK
            objUpsRates.Reset()
            objUpsRates.RuntimeLicense = s4DPaymentsShippingSDK

            Dim bearerToken As String = GetAuthorizationToken(Carriers.UPS, cAccountNumber)
            If inTestMode Then
                objUpsRates.Config("TESTMODE=true")
            End If

            objUpsRates.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)

            objUpsRates.UPSAccount.AccountNumber = cAccountNumber
            objUpsRates.UPSAccount.AuthorizationToken = bearerToken

            objUpsRates.RequestedService = cRequestedServiceType
            objUpsRates.PickupType = UPSPickupType
            objUpsRates.CustomerType = cCustomerType

            ' Insured Value is Positive, Declared Value is Negative
            Dim iPackage As Int16 = 0
            Dim totalWeight As Double = 0
            Dim totalInsured As Decimal = 0

            For Each shippingPackageDetail In PackageDetailList

                If cSignatureRequired Then
                    shippingPackageDetail.SignatureType = TSignatureTypes.stDirect
                End If

                objUpsRates.Packages.Add(shippingPackageDetail)

                shippingPackageDetail.InsuredValue = Format(Val(shippingPackageDetail.InsuredValue), "###0.00")
                If Val(shippingPackageDetail.InsuredValue & String.Empty) < 0 Then
                    shippingPackageDetail.InsuredValue = Math.Abs(Val(shippingPackageDetail.InsuredValue & String.Empty))
                    objUpsRates.Config("PackageDeclaredValueType[" & iPackage & "]=0")
                End If

                ' Format weight 
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight) / 16, "###0.0")
                totalWeight += Val(shippingPackageDetail.Weight)
                totalInsured = Val(shippingPackageDetail.InsuredValue)

                iPackage += 1
            Next

            objUpsRates.TotalWeight = Format(Val(totalWeight), "###0.0")

            With cSenderContact
                objUpsRates.SenderAddress.ZipCode = .ZipCode
                objUpsRates.SenderAddress.State = .State
                objUpsRates.SenderAddress.CountryCode = .CountryCode
            End With

            With cRecipientContact
                objUpsRates.RecipientAddress.ZipCode = .ZipCode

                If .CountryCode = "US" And .State = "PR" Then
                    objUpsRates.RecipientAddress.State = .State
                    objUpsRates.RecipientAddress.CountryCode = .State
                ElseIf .CountryCode = "US" Then
                    objUpsRates.RecipientAddress.State = .State
                    objUpsRates.RecipientAddress.CountryCode = .CountryCode
                ElseIf .CountryCode <> "US" Then
                    objUpsRates.RecipientAddress.State = .State
                    objUpsRates.RecipientAddress.CountryCode = .CountryCode
                End If

                If .IsResidental Then
                    objUpsRates.RecipientAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objUpsRates.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

            If IsDate(Me.ShipDate) Then
                objUpsRates.ShipDate = CDate(Me.ShipDate).ToString("yyyyMMdd")
            Else
                objUpsRates.ShipDate = DateTime.Now.ToString("yyyyMMdd")
            End If

            objUpsRates.ShipmentSpecialServices = ShipmentSpecialServices

            objUpsRates.GetRates()

            ReDim requestedRateList(objUpsRates.Services.Count)
            For iLoop As Integer = 0 To objUpsRates.Services.Count - 1
                With requestedRateList(iLoop)
                    .ServiceType = objUpsRates.Services(iLoop).ServiceType
                    .ServiceTypeDescription = objUpsRates.Services(iLoop).ServiceTypeDescription
                    .AccountNetCharge = Val(objUpsRates.Services(iLoop).AccountNetCharge & String.Empty)
                    .DeliveryTime = objUpsRates.Services(iLoop).DeliveryTime
                    .ListNetCharge = Val(objUpsRates.Services(iLoop).ListNetCharge & String.Empty)
                    .TransitTime = objUpsRates.Services(iLoop).TransitTime
                    .ServiceCode = objUpsRates.Services(iLoop).ServiceType

                    If .AccountNetCharge = 0 Then
                        .AccountNetCharge = .ListNetCharge
                    End If
                End With
            Next

            Return requestedRateList
        Catch ex As Exception
            LastError = ex.Message
            Return Nothing
        Finally
            cRawRequest = objUpsRates.Config("RawRequest")
            cRawResponse = objUpsRates.Config("RawResponse")
        End Try

    End Function

    Public Function TrackUPSPackage(ByVal TrackingNumber As String, Optional ProofOfDeliveryFilename As String = "") As DataTable

        Dim upsTrack As New UPSTrack
        Dim tblTracking As New DataTable

        Try
            With tblTracking
                .Columns.Add("TRACKING_NO", GetType(System.String))
                .Columns.Add("Address1", GetType(System.String))
                .Columns.Add("Address2", GetType(System.String))
                .Columns.Add("City", GetType(System.String))
                .Columns.Add("Company", GetType(System.String))
                .Columns.Add("CountryCode", GetType(System.String))
                .Columns.Add("Date", GetType(System.DateTime))
                .Columns.Add("Exception", GetType(System.String))
                .Columns.Add("Location", GetType(System.String))
                .Columns.Add("Other", GetType(System.String))
                .Columns.Add("State", GetType(System.String))
                .Columns.Add("Status", GetType(System.String))
                .Columns.Add("Time", GetType(System.String))
                .Columns.Add("Zipcode", GetType(System.String))
                .Columns.Add("PackageReceivedBy", GetType(System.String))
                .Columns.Add("PackageWeight", GetType(System.Decimal))
                .Columns.Add("ServiceTypeDescription", GetType(System.String))
                .Columns.Add("DeliveryLocation", GetType(System.String))
                .Columns.Add("DateShipped", GetType(System.String))
                .Columns.Add("PackageCount", GetType(System.Int32))
            End With

            LastError = String.Empty

            upsTrack.RuntimeLicense = s4DPaymentsShippingSDK
            upsTrack.Reset()
            upsTrack.RuntimeLicense = s4DPaymentsShippingSDK

            upsTrack.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            Dim bearertoken As String = GetAuthorizationToken(Carriers.UPS, cAccountNumber)
            upsTrack.Config("TESTMODE=false")
            upsTrack.Config("CustomerTransactionId=100")
            upsTrack.UPSAccount.AuthorizationToken = bearertoken
            upsTrack.TrackShipment(TrackingNumber)

            Dim DeliveryLocation As String = String.Empty
            Dim DateShipped As String = String.Empty
            Dim PackageCount As Int32 = 1

            Try
                Dim dst As New DataSet
                cRawResponse = upsTrack.Config("RawResponse")

                ' Convert Json to XML then convert XML to Dataset
                Dim doc As XmlDocument = JsonConvert.DeserializeXmlNode("{ 'root': " & cRawResponse & "}")
                Dim result As String = "<?xml version=""1.0"" encoding=""UTF-8""?><TrackEvents>" & doc.ChildNodes(0).InnerXml & "</TrackEvents>"
                result = result.Replace("><", $">{Environment.NewLine}<")

                doc.LoadXml(result)
                Dim sr As IO.StringReader = New IO.StringReader(doc.InnerXml)
                Dim xtr As XmlTextReader = New XmlTextReader(sr)
                dst.ReadXml(xtr)
                If dst IsNot Nothing Then
                    If dst.Tables.Contains("deliveryinformation") Then
                        If dst.Tables("deliveryinformation").Rows.Count > 0 Then
                            If dst.Tables("deliveryinformation").Columns.Contains("location") Then
                                DeliveryLocation = dst.Tables("deliveryinformation").Rows(0).Item("location") & String.Empty
                            End If
                        End If
                    End If

                    If dst.Tables.Contains("package") Then
                        If dst.Tables("package").Rows.Count > 0 Then
                            PackageCount = Val(dst.Tables("package").Rows(0).Item("packagecount") & String.Empty)
                            If PackageCount < 1 Then PackageCount = 1
                        End If
                    End If

                    If dst.Tables.Contains("activity") Then
                        For Each rowactivity As DataRow In dst.Tables("activity").Select("", "activity_id desc")
                            Dim aDate As String = rowactivity.Item("date") & String.Empty
                            Dim aTime As String = rowactivity.Item("time") & String.Empty

                            If aDate.Length = 8 Then
                                aDate = aDate.Substring(4, 2) & "/" & aDate.Substring(6, 2) & "/" & aDate.Substring(0, 4)
                                If aTime.Length = 6 Then
                                    aTime = aTime.Substring(0, 2) & ":" & aTime.Substring(2, 2) & ":" & aTime.Substring(4, 2)
                                Else
                                    aTime = String.Empty
                                End If
                            End If
                            DateShipped = (aDate & " " & aTime).Trim
                            Exit For
                        Next
                    End If
                End If
            Catch ex As Exception

            End Try

            If upsTrack.Config("Warning").Length = 0 Then

                For Each trackDetail As TrackDetail In upsTrack.TrackEvents
                    Dim row As DataRow = tblTracking.NewRow
                    row.Item("TRACKING_NO") = upsTrack.PackageTrackingNumber
                    row.Item("Address1") = trackDetail.Address1 & String.Empty
                    row.Item("Address2") = trackDetail.Address2 & String.Empty
                    row.Item("City") = trackDetail.City & String.Empty
                    row.Item("Company") = trackDetail.Company & String.Empty
                    row.Item("CountryCode") = trackDetail.CountryCode & String.Empty
                    If IsDate(trackDetail.Date & String.Empty) Then
                        row.Item("Date") = CDate(trackDetail.Date & String.Empty)
                    Else
                        row.Item("Date") = DBNull.Value
                    End If
                    row.Item("Exception") = trackDetail.Exception & String.Empty
                    row.Item("Location") = trackDetail.Location & String.Empty
                    row.Item("Other") = trackDetail.Other & String.Empty
                    row.Item("State") = trackDetail.State & String.Empty
                    row.Item("Status") = (trackDetail.Status & String.Empty).Trim
                    row.Item("Time") = trackDetail.Time & String.Empty
                    row.Item("ZipCode") = trackDetail.ZipCode & String.Empty
                    row.Item("PackageWeight") = Val(upsTrack.PackageWeight & String.Empty)
                    row.Item("ServiceTypeDescription") = upsTrack.ServiceTypeDescription & String.Empty
                    row.Item("DateShipped") = DateShipped
                    tblTracking.Rows.Add(row)

                    If row.Item("Status") & String.Empty = "DELIVERED" Then
                        row.Item("DeliveryLocation") = DeliveryLocation
                        row.Item("PackageCount") = PackageCount
                        row.Item("PackageReceivedBy") = upsTrack.PackageReceivedBy & String.Empty
                    End If
                Next
            End If

            Return tblTracking

        Catch ex As DShippingSDKException
            LastError = ex.Message
            Return Nothing
        Catch exc As Exception
            LastError = exc.Message
            Return Nothing
        Finally
            cRawRequest = upsTrack.Config("RawRequest")
            cRawResponse = upsTrack.Config("RawResponse")
        End Try

    End Function


    Public Class TRACKINFO
        Public PACKAGE_WEIGHT As Decimal
        Public SERVICE_DESC As String
        Public DELIVERY_DATE As String
        Public RECEIVED_BY As String
        Public DELIVERY_CITY As String
        Public DELIVERY_STATE As String
        Public DELIVERY_ZIPCODE As String
        Public DELIVERY_LOCATION As String
        Public EXPECTED_DELIVERY_DATE As String
        Public PACKAGE_POD As String
    End Class

    ''' <summary>
    ''' Get The UPS Proof of Delivery
    ''' </summary>
    ''' <param name="TrackingNumber">UPS Tracking Nimber</param>
    ''' <param name="AccountNo">UPS Accunt Number</param>
    ''' <param name="clsTRACKINFO">Returns POD HTML</param>
    ''' <returns></returns>
    Public Function UPSProofOfDelivery(ByVal TrackingNumber As String,
                                       ByVal AccountNo As String,
                                       ByRef clsTRACKINFO As TRACKINFO,
                                       ByVal PackagePODFile As String) As Boolean
        Try
            Dim upsTrack As New UPSTrack
            upsTrack.RuntimeLicense = s4DPaymentsShippingSDK
            upsTrack.Reset()
            upsTrack.RuntimeLicense = s4DPaymentsShippingSDK

            upsTrack.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            Dim bearertoken As String = String.Empty ' GetAuthorizationToken(rowSOTCARR3)
            upsTrack.Config("TESTMODE=false")
            upsTrack.Config("CustomerTransactionId=100")

            'Dim appPath As String = System.IO.Path.Combine(System.IO.Path.Combine(ASCMAIN1.Folders("Archive"), "POD"))
            'If Not My.Computer.FileSystem.DirectoryExists(appPath) Then
            '    My.Computer.FileSystem.CreateDirectory(appPath)
            'End If

            bearertoken = GetAuthorizationToken(Carriers.UPS, AccountNo)
            upsTrack.UPSAccount.AuthorizationToken = bearertoken
            upsTrack.TrackShipment(TrackingNumber)

            With clsTRACKINFO
                .PACKAGE_WEIGHT = Val(upsTrack.PackageWeight & String.Empty)
                .SERVICE_DESC = upsTrack.ServiceTypeDescription

                Dim DELIVERY_DATE As String = upsTrack.PackageDeliveredOnDate & String.Empty
                Dim DELIVERY_TIME As String = upsTrack.PackageDeliveryTime.Replace("Delivered Time:-", "").Trim

                If IsDate(DELIVERY_DATE) Then
                    If IsDate(DELIVERY_DATE & " " & DELIVERY_TIME) Then
                        DELIVERY_DATE = DELIVERY_DATE & " " & DELIVERY_TIME
                    End If
                End If

                .DELIVERY_DATE = DELIVERY_DATE
                .RECEIVED_BY = upsTrack.PackageReceivedBy & String.Empty
                .DELIVERY_CITY = upsTrack.RecipientAddress.City & String.Empty
                .DELIVERY_STATE = upsTrack.RecipientAddress.State & String.Empty
                .DELIVERY_ZIPCODE = upsTrack.RecipientAddress.ZipCode & String.Empty
                .DELIVERY_LOCATION = String.Empty
                .EXPECTED_DELIVERY_DATE = upsTrack.PackageScheduledDeliveryDate & String.Empty
                .PACKAGE_POD = ""
            End With

            upsTrack.Config($"PackagePODFile={PackagePODFile}")
            If My.Computer.FileSystem.FileExists(PackagePODFile) Then
                clsTRACKINFO.PACKAGE_POD = PackagePODFile
            End If

            Return True

        Catch ex As Exception
            LastError = $"UPS Proof Of Delivery Error: {ex.Message}"
            Return False
        Finally

        End Try
    End Function


#End Region

#Region "Private Class Procedures"

    Private Sub GetPackageCosts(ByVal package As PackageDetail, ByVal shipObject As Object)

        Dim xdoc As New System.Xml.XmlDocument
        Dim PayorListPackageNetAmount As Decimal = 0
        Dim processingPayorList As Boolean = False
        Dim netFreight As Boolean = False
        Dim SHIP_PACKAGE_NO As String = String.Empty

        Try
            SHIP_PACKAGE_NO = package.Id

            If TypeOf shipObject Is FedExShipIntl Then
                ShipmentBaseCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.TotalNetCharge))
                ShipmentDiscountCharge.Add(SHIP_PACKAGE_NO, 0)
                ShipmentSurCharge.Add(SHIP_PACKAGE_NO, 0)
                ShipmentNetCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.TotalNetCharge))
                ShipmentListCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.TotalNetCharge))
                Exit Try
            ElseIf TypeOf shipObject Is FedExShip Then
                ShipmentBaseCharge.Add(SHIP_PACKAGE_NO, Val(package.BaseCharge))
                ShipmentDiscountCharge.Add(SHIP_PACKAGE_NO, Val(package.TotalDiscount))
                ShipmentSurCharge.Add(SHIP_PACKAGE_NO, Val(package.TotalSurcharges))
                ShipmentNetCharge.Add(SHIP_PACKAGE_NO, Val(package.NetCharge))
                ShipmentListCharge.Add(SHIP_PACKAGE_NO, Val(package.NetCharge))
                If package.RatingAggregate.Length > 0 Then
                    Try
                        Dim packageRatingAggregate As String = "<?xml version=""1.0""?>" & vbCrLf & package.RatingAggregate.Replace("v9:", "").Replace("v12:", "")
                        Dim fedexAggDoc As New System.Xml.XmlDocument
                        fedexAggDoc.LoadXml(packageRatingAggregate)
                        Dim root As XmlNode = fedexAggDoc.DocumentElement
                        Dim PAYOR_LIST_PACKAGE As XmlNode = root.SelectSingleNode("descendant::PackageRateDetails[RateType=""PAYOR_LIST_PACKAGE""]")
                        Dim listNetCharge As Double = Val(PAYOR_LIST_PACKAGE.SelectSingleNode("NetCharge/Amount").InnerText & String.Empty)
                        ShipmentListCharge(SHIP_PACKAGE_NO) = listNetCharge
                    Catch ex As Exception

                    End Try
                    Exit Try
                End If

            ElseIf TypeOf shipObject Is UPSShip Then
                ShipmentBaseCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.TotalBaseCharge))
                ShipmentDiscountCharge.Add(SHIP_PACKAGE_NO, 0)
                ShipmentSurCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.TotalSurcharges))
                ShipmentNetCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.Config("AccountTotalNetCharge") & String.Empty))
                ShipmentListCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.TotalNetCharge))
                Exit Try
            Else
                ShipmentBaseCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.BaseCharge))
                ShipmentDiscountCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.TotalDiscount))
                ShipmentSurCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.TotalSurcharges))
                ShipmentNetCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.NetCharge))
                ShipmentListCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.NetCharge))
                Exit Try
            End If

        Catch ex As Exception
            If Not ShipmentListCharge.ContainsKey(SHIP_PACKAGE_NO) Then
                ShipmentListCharge.Add(SHIP_PACKAGE_NO, PayorListPackageNetAmount)
            End If

        Finally
            If ShipmentListCharge(SHIP_PACKAGE_NO) < ShipmentNetCharge(SHIP_PACKAGE_NO) Then
                ShipmentListCharge(SHIP_PACKAGE_NO) = ShipmentNetCharge(SHIP_PACKAGE_NO)
            End If
        End Try

    End Sub

    Private Function ValidateEmail(ByVal emailAddress As String) As Boolean

        Dim strDomainName As String = String.Empty
        Dim strDomainType As String = String.Empty
        Dim strUserName As String = String.Empty
        Const sInvalidChars As String = "!#$%^&*()=+{}[]|\;:'/?>,< "
        Dim i As Integer

        If Trim(emailAddress) = "" Then
            Return False
        End If

        'Check to see if there is a double quote
        If InStr(1, emailAddress, Chr(34)) > 0 Then Return False

        'Check to see if there are consecutive dots
        If InStr(1, emailAddress, "..") > 0 Then Return False

        ' Check for invalid characters.
        If Len(emailAddress) > Len(sInvalidChars) Then
            For i = 1 To Len(sInvalidChars)
                If InStr(emailAddress, Mid(sInvalidChars, i, 1)) > 0 Then
                    Return False
                End If
            Next
        Else
            For i = 1 To Len(emailAddress)
                If InStr(sInvalidChars, Mid(emailAddress, i, 1)) > 0 Then
                    Return False
                End If
            Next
        End If

        'Check for an @ symbol
        If InStr(1, emailAddress, "@") <= 1 Then
            Return False
        End If

        If emailAddress.EndsWith("@") Then
            Return False
        End If

        strUserName = emailAddress.Substring(0, InStr(1, emailAddress, "@") - 1)
        Dim domain As String = emailAddress.Substring(InStr(1, emailAddress, "@"))

        'Check to see if there are too many @'s
        If InStr(1, domain, "@") > 0 Then
            Return False
        End If

        For Each part As String In domain.Split(".")
            If Trim(part) = "" Then
                Return False
            End If
        Next

        Return True

    End Function

#End Region

#Region "Generic Service Provider"

    Private Function RequestGenericLabel() As Boolean

        Try
            lstShippingLabels.Clear()

            For Each shippingPackageDetail In PackageDetailList

                Dim Weight As Decimal = Val(shippingPackageDetail.Weight & String.Empty)
                Dim Length As Decimal = Val(shippingPackageDetail.Length & String.Empty)
                Dim Width As Decimal = Val(shippingPackageDetail.Width & String.Empty)
                Dim Height As Decimal = Val(shippingPackageDetail.Height & String.Empty)

                Dim labelImage As String = String.Empty
                Dim strText As String = String.Empty

                If AccountNumber.Length > 0 Then
                    cMasterTrackingNumber = AccountNumber & cMasterTrackingNumber
                End If

                labelImage &= "^XA" & Environment.NewLine
                labelImage &= $"^CF0,40" & Environment.NewLine
                labelImage &= $"^FO40,5^FDGeneric Shipping Label^FS" & Environment.NewLine

                labelImage &= $"^CF0,30" & Environment.NewLine
                labelImage &= $"^FO40,50^FD{GenericLabelShipMethod} - {GenericLabelShipMethodDesc}^FS" & Environment.NewLine
                labelImage &= $"^FO40,90^FD{GenericLabelCarrier}^FS" & Environment.NewLine

                labelImage &= "^FO50,150^GB700,3,3^FS" & Environment.NewLine
                labelImage &= $"^FX Ship From Address" & Environment.NewLine
                labelImage &= $"^CFA,30" & Environment.NewLine

                Dim line As Int32 = 215
                labelImage &= $"^FO40,{line}^FD{Sender.Company}^FS" & Environment.NewLine
                line += 40

                If Sender.Address1.Length > 0 Then
                    labelImage &= $"^FO40,{line}^FD{Sender.Address1}^FS" & Environment.NewLine
                    line += 40
                End If

                If Sender.Address2.Length > 0 Then
                    labelImage &= $"^FO40,{line}^FD{Sender.Address2}^FS" & Environment.NewLine
                    line += 40
                End If

                If Sender.Address3.Length > 0 Then
                    labelImage &= $"^FO40,{line}^FD{Sender.Address3}^FS" & Environment.NewLine
                    line += 40
                End If

                strText = $"{Sender.City}, {Sender.State} {Sender.ZipCode }"
                labelImage &= $"^FO40,{line}^FD{strText}^FS" & Environment.NewLine

                'labelImage &= $"^FO50,{line + 60}^GB700,3,3^FS" & Environment.NewLine

                line += 100

                labelImage &= $"^FX Ship To Address" & Environment.NewLine
                labelImage &= $"^FO40,{line}^FDShip To:^FS" & Environment.NewLine
                line += 40

                labelImage &= $"^FO60,{line}^FD{Recipient.Company}^FS" & Environment.NewLine
                line += 40

                If Recipient.Address1.Length > 0 Then
                    labelImage &= $"^FO60,{line}^FD{Recipient.Address1}^FS" & Environment.NewLine
                    line += 40
                End If

                If Recipient.Address2.Length > 0 Then
                    labelImage &= $"^FO60,{line}^FD{Recipient.Address2}^FS" & Environment.NewLine
                    line += 40
                End If

                If Recipient.Address3.Length > 0 Then
                    labelImage &= $"^FO60,{line}^FD{Recipient.Address3}^FS" & Environment.NewLine
                    line += 40
                End If

                strText = $"{Recipient.City}, {Recipient.State} {Recipient.ZipCode}"
                labelImage &= $"^FO60,{line}^FD{strText}^FS" & Environment.NewLine

                'labelImage &= $"^FO50,{line + 60}^GB700,3,3^FS" & Environment.NewLine

                line += 100

                labelImage &= $"^FX Tracking No bar code." & Environment.NewLine
                labelImage &= $"^BY3,2,150" & Environment.NewLine
                labelImage &= $"^FO100,{line}^BC^FD{shippingPackageDetail.Id}^FS" & Environment.NewLine

                'labelImage &= $"^FO50,{line + 200}^GB700,3,3^FS" & Environment.NewLine

                line += 250
                strText = $"Box: {Length}x{Width}x{Height} - Wgt {Weight} lbs." & Environment.NewLine
                labelImage &= $"^FO40,{line}^FD{strText}^FS" & Environment.NewLine
                line += 40

                Dim reference As String = shippingPackageDetail.Reference & String.Empty
                For Each UnityReference As String In reference.Split(";")
                    If UnityReference.Length = 0 Then Continue For
                    labelImage &= $"^FO40,{line}^FD{UnityReference}^FS" & Environment.NewLine
                    line += 40
                Next

                line += 60
                If ASCMAIN1.DBS_SERVER <> "ABO" Then
                    labelImage &= $"^CF0,60" & Environment.NewLine
                    labelImage &= $"^FO60,{line}^FDTEST - Do Not Ship^FS" & Environment.NewLine
                End If

                labelImage &= "^XZ" & Environment.NewLine
                lstShippingLabels.Add(shippingPackageDetail.Id, labelImage)
            Next
            Return True

        Catch ex As Exception
            cMasterTrackingNumber = String.Empty
            lstShippingLabels.Clear()
            LastError = $"Request Unity Label Error: {ex.Message}"
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Request Shipping label. Not used for Fedex Intenational
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' 
    Private Function RequestLabelOther() As Boolean

        Try
            objEzShip.RuntimeLicense = s4DPaymentsShippingSDK
            objEzShip.Reset()
            objEzShip.RuntimeLicense = s4DPaymentsShippingSDK

            If cServiceProvider = ServiceProviders.Unknown Then
                LastError = "Unknown Service Type"
                Return False
            End If

            objEzShip.Provider = cServiceProvider
            objEzShip.Account.AccountNumber = cAccountNumber
            objEzShip.Account.UserId = cUserId
            objEzShip.Account.Password = cPassword
            objEzShip.ServiceType = cRequestedServiceType

            'If cServiceProvider = ServiceProviders.FederalExpress Then
            '    objEzShip.Account.MeterNumber = cFedexMeterNumber
            '    objEzShip.Account.DeveloperKey = cFedexDeveloperKey
            'Else
            '    objEzShip.Account.MeterNumber = String.Empty
            '    objEzShip.Account.DeveloperKey = String.Empty
            'End If

            If cServiceProvider = ServiceProviders.UPS Then
                objEzShip.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)

                Dim bearerToken As String = GetAuthorizationToken(Carriers.UPS, cAccountNumber)
                If inTestMode Then
                    objEzShip.Config("TESTMODE=true")
                End If
                objEzShip.Account.AuthorizationToken = bearerToken
            End If

            If cServiceProvider = ServiceProviders.USPS Then
                objEzShip.Config("PostageProvider=1") 'Use Endicia instead of USPS directly.
                objEzShip.Config("CustomerId=" & cUSPSEndiciaCustomerId) 'Mandatory for Endicia
                objEzShip.Config("TransactionId=" & cUSPSEndiciaTransactionId) 'Mandatory for Endicia
            End If

            objEzShip.ShipDate = ShipDate.ToString("yyyy-MM-dd")

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
                objEzShip.Config("SenderAddress3=" & .Address3)

                objEzShip.SenderAddress.City = .City
                objEzShip.SenderAddress.ZipCode = .ZipCode
                objEzShip.SenderAddress.State = .State
                objEzShip.SenderAddress.CountryCode = .CountryCode
            End With

            With cRecipientContact
                objEzShip.RecipientContact.FirstName = .FirstName
                objEzShip.RecipientContact.LastName = .LastName
                objEzShip.RecipientContact.MiddleInitial = .MiddleInitial
                objEzShip.RecipientContact.Phone = .Phone
                objEzShip.RecipientContact.Fax = .Fax
                objEzShip.RecipientContact.Email = .eMail

                objEzShip.RecipientContact.Company = .Company
                objEzShip.RecipientAddress.Address1 = .Address1
                objEzShip.RecipientAddress.Address2 = .Address2
                objEzShip.Config("RecipientAddress3=" & .Address3)

                objEzShip.RecipientAddress.City = .City
                objEzShip.RecipientAddress.ZipCode = .ZipCode
                objEzShip.RecipientAddress.State = .State
                objEzShip.RecipientAddress.CountryCode = .CountryCode
            End With

        Catch ex As Exception
            LastError = ex.Message
            Return False
        End Try

        Try

            Dim extension As String = EzshipLabelImage.ToString
            If extension.StartsWith("it") Then
                extension = "." & extension.Substring(2)
            Else
                extension = String.Empty
            End If

            ' Fix label type if an error
            Try
                objEzShip.LabelImageType = EzshipLabelImage
            Catch ex As Exception
                Select Case cServiceProvider
                    Case ServiceProviders.CanadaPost
                        objEzShip.LabelImageType = EzShipLabelImageTypes.itZPL
                    Case ServiceProviders.FederalExpress
                        objEzShip.LabelImageType = EzShipLabelImageTypes.itEltron
                    Case ServiceProviders.FederalExpressInternational
                        objEzShip.LabelImageType = EzShipLabelImageTypes.itEltron
                    Case ServiceProviders.UPS
                        objEzShip.LabelImageType = EzShipLabelImageTypes.itZPL
                    Case ServiceProviders.USPS
                        objEzShip.LabelImageType = EzShipLabelImageTypes.itZPL
                End Select
            End Try


            ' Set shipping directory to store the labels
            Dim idCtr As Int16 = 1
            If ShippingLabelDirectory.Length > 0 AndAlso ShippingLabelPrefix.Length > 0 Then
                If My.Computer.FileSystem.DirectoryExists(ShippingLabelDirectory) Then

                    If Not ShippingLabelDirectory.EndsWith("\") Then
                        ShippingLabelDirectory &= "\"
                    End If

                    For Each shippingPackageDetail In PackageDetailList
                        Dim id As String = idCtr.ToString
                        idCtr += 1
                        shippingPackageDetail.ShippingLabelFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & extension
                        shippingPackageDetail.CODFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & "_COD" & extension
                    Next
                End If
            End If

            ' Add packages
            For Each shippingPackageDetail In PackageDetailList
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight), "###0.0")
                objEzShip.Packages.Add(shippingPackageDetail)
            Next

            objEzShip.RuntimeLicense = s4DPaymentsShippingSDK
            objEzShip.GetShipmentLabels()

            ' Reset the object to have the updated data returned
            PackageDetailList.Clear()
            For ictr As Int16 = 0 To objEzShip.Packages.Count - 1
                PackageDetailList.Add(objEzShip.Packages(ictr))
                GetPackageCosts(objEzShip.Packages(ictr), objEzShip)
                lstShippingLabels.Add(objEzShip.Packages(ictr).TrackingNumber, objEzShip.Packages(ictr).ShippingLabel)
            Next

            If objEzShip.Packages.Count = 1 Then
                cMasterTrackingNumber = objEzShip.Packages(0).TrackingNumber
            Else
                cMasterTrackingNumber = objEzShip.MasterTrackingNumber
            End If

            Return True

        Catch ex As DShippingSDKException
            LastError = ex.Message
            Return False
        Catch exc As Exception
            LastError = exc.Message
            Return False
        Finally
            cRawRequest = objEzShip.Config("RawRequest")
            cRawResponse = objEzShip.Config("RawResponse")
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
        Public Address3 As String = String.Empty
        Public City As String = String.Empty
        Public State As String = String.Empty
        Public ZipCode As String = String.Empty
        Public CountryCode As String = String.Empty
        Public IsResidental As Boolean = False
        Public IsPOBox As Boolean = False

        Public AccountNumber As String = String.Empty

        ' Needed for CDL labels
        Public ShipmentReferenceNumber As String = String.Empty
        Public CustomerCode As String = String.Empty
        Public InvoiceNumber As String = String.Empty
        Public Reference As String = String.Empty
    End Class

    Class ABSPayor
        Public PayorType As String = String.Empty
        Public AccountNumber As String = String.Empty
        Public CountryCode As String = String.Empty
    End Class

    Class SmartPost
        Public AncillaryEndorsement As String = "0"
        Public CustomerManifestId As String = String.Empty
        Public HubId As String = "5531"
        Public Indicia As String = "1"
        Public PhysicalPackaging As String = "4"
        Public TrackingNumbers As List(Of String) = New List(Of String)
    End Class

#End Region

#Region "Printing"

    Public Function PrintShippingLabels(ByVal shippinglabels As List(Of String)) As Boolean
        Try

            For Each label As String In shippinglabels
                label = label.Trim
                If label.Length = 0 Then
                    Continue For
                End If
                PrintShippingLabels(label)
            Next

        Catch ex As Exception
            LastError = ex.Message
            Return False
        End Try

    End Function

    Public Function PrintShippingLabels() As Boolean

        Try
            Dim shippinglabels As New List(Of String)
            For Each kvp As KeyValuePair(Of String, String) In lstShippingLabels
                If kvp.Value & String.Empty <> String.Empty Then
                    shippinglabels.Add(kvp.Value)
                End If
            Next

            For Each label As String In shippinglabels
                PrintShippingLabels(label)
            Next

        Catch ex As Exception
            LastError = ex.Message
            Return False
        End Try

    End Function

    Public Function PrintShippingLabels(ByVal shippingLabelFile As String) As Boolean

        Dim Label As String = String.Empty

        If My.Computer.FileSystem.FileExists(shippingLabelFile) Then

            If My.Computer.FileSystem.FileExists(shippingLabelFile) Then
                Using sr As New System.IO.StreamReader(shippingLabelFile)
                    Label = sr.ReadToEnd
                    sr.Close()
                    sr.Dispose()
                End Using

                'clsTACZPLT1.SendLabelToPrinter(ASCMAIN1.LabelPrinterIPAddress, Label)
            End If
        End If

    End Function

    Public Function GetLabelImages() As String

        Try
            Dim LabelImage As String = String.Empty

            Dim shippinglabels As New List(Of String)
            For Each kvp As KeyValuePair(Of String, String) In lstShippingLabels
                If kvp.Value & String.Empty <> String.Empty Then
                    shippinglabels.Add(kvp.Value)
                End If
            Next

            For Each label As String In shippinglabels
                label = label.Trim
                If label.Length = 0 Then
                    Continue For
                End If

                If My.Computer.FileSystem.FileExists(label) Then
                    Using sr As New System.IO.StreamReader(label)
                        label = sr.ReadToEnd
                        sr.Close()
                        sr.Dispose()
                    End Using

                    LabelImage &= label
                End If
            Next
            Return LabelImage

        Catch ex As Exception
            Return String.Empty
        End Try

    End Function

#End Region

#Region "Oauth"

    Private inTestMode As Boolean = False

    Private Enum Carriers
        FedEx
        UPS
    End Enum

    Private Function GetAuthorizationToken(ByVal Carrier As Carriers,
                                           ByVal AccountNo As String) As String

        Static numLoops As Int16 = 0
        Const MT_LEVEL As String = "888"

        If numLoops > 5 Then
            numLoops = 0
            Return String.Empty
        End If

        Dim CARRIER_CODE As String = String.Empty
        Select Case Carrier
            Case Carriers.FedEx
                CARRIER_CODE = "FEDEX"

            Case Carriers.UPS
                CARRIER_CODE = "UPS"

            Case Else
                numLoops = 0
                Return String.Empty
        End Select

        Dim sql As String = "Select * from SOTCARR3 where CARRIER_CODE = :PARM1 AND CARRIER_ACCOUNT_NO = :PARM2"
        Dim rowSOTCARR3 As DataRow = ASCDATA1.GetDataRow(sql, "VV", {CARRIER_CODE, AccountNo})

        If rowSOTCARR3 Is Nothing Then
            numLoops = 0
            Return String.Empty
        End If

        Dim ServerTokenURL As String = rowSOTCARR3.Item("SERVER_TOKEN_URL") & String.Empty
        Dim ClientId As String = rowSOTCARR3.Item("CLIENT_ID") & String.Empty
        Dim ClientSecret As String = rowSOTCARR3.Item("CLIENT_SECRET") & String.Empty
        Dim AUTH_CODE As String = rowSOTCARR3.Item("AUTH_CODE") & String.Empty

        ' Test Server Token Urls.
        'https://wwwcie.ups.com/security/v1/oauth/token
        'https://apis-sandbox.fedex.com/oauth/token

        ' Production Server Token URLs.
        'https://onlinetools.ups.com/security/v1/oauth/token
        'https://apis.fedex.com/oauth/token

        ' Going to use the Server Token URL to determine if we are in test mode.
        Select Case Carrier
            Case Carriers.FedEx
                inTestMode = ServerTokenURL.ToUpper.Contains("apis-sandbox".ToUpper)

            Case Carriers.UPS
                inTestMode = ServerTokenURL.ToUpper.Contains("wwwcie.ups".ToUpper)

            Case Else
                numLoops = 0
                Return String.Empty
        End Select

        ' See if we need to get a new token; otherwise, return the current Authorization Code
        If rowSOTCARR3.Item("TOKEN_EXPIRES") & String.Empty <> String.Empty Then
            If IsDate(rowSOTCARR3.Item("TOKEN_EXPIRES") & String.Empty) Then
                If DateDiff(DateInterval.Minute, CDate(rowSOTCARR3.Item("TOKEN_EXPIRES") & String.Empty), DateTime.Now) < 0 Then
                    If AUTH_CODE.Length > 0 Then
                        numLoops = 0
                        Return AUTH_CODE
                    End If
                End If
            End If
        End If

        ' If we hit this then there is no token or the token has expired.
        ' Only one user can refresh the token
        'If Not ASCMAIN1.Logical_Lock("SOTCARR3", AccountNo,, False, False, MT_LEVEL) Then
        '    System.Threading.Thread.Sleep(2000)
        '    numLoops += 1
        '    Return GetAuthorizationToken(Carrier, AccountNo)
        'End If

        Dim result As Boolean = ASCDATA1.ExecuteSF(Of Boolean)("MULTITASK.LOGICAL_LOCK",
                            {"P_ENTITY_TYPE", "P_ENTITY", "P_SESSION_NO", "P_SELECTION_NO", "P_OPER", "P_MT_LEVEL", "P_REVERSE_PREVIOUS_IF_UNSUCCESSFUL"},
                            {"SOTCARR3", AccountNo, ASCMAIN1.SESSION_NO, ASCMAIN1.ActiveForm.SELECTION_NO, ASCMAIN1.USER_ID, CInt(MT_LEVEL), False})

        If Not result Then
            System.Threading.Thread.Sleep(2000)
            numLoops += 1
            Return GetAuthorizationToken(Carrier, AccountNo)
        End If

        Dim bearerToken As String = String.Empty
        Dim tokenExp As DateTime
        Dim oauth As New OAuth
        Dim errorMessage As String = String.Empty

        oauth.RuntimeLicense = s4DPaymentsShippingSDK

        Try
            Select Case Carrier
                Case Carriers.UPS
                    With oauth
                        .GrantType = OAuthGrantTypes.ogtClientCredentials
                        .ServerTokenURL = ServerTokenURL
                        .ClientId = ClientId
                        .ClientSecret = ClientSecret
                        .ClientProfile = OAuthClientProfiles.ocpApplication
                    End With

                    bearerToken = oauth.GetAuthorization
                    tokenExp = DateAdd(DateInterval.Minute, -5, DateTime.Now.AddSeconds(oauth.AccessTokenExp))

                Case Carriers.FedEx
                    With oauth
                        .GrantType = OAuthGrantTypes.ogtClientCredentials
                        .ServerTokenURL = ServerTokenURL
                        .ClientId = ClientId
                        .ClientSecret = ClientSecret
                        .Config("IncludeClientCredsInBody=true")
                    End With

                    bearerToken = oauth.GetAuthorization
                    tokenExp = DateAdd(DateInterval.Minute, -5, DateTime.Now.AddSeconds(oauth.AccessTokenExp))

            End Select

            sql = "DECLARE PRAGMA AUTONOMOUS_TRANSACTION; 
                    BEGIN UPDATE SOTCARR3 SET AUTH_CODE = :PARM1, TOKEN_EXPIRES = :PARM2 WHERE CARRIER_CODE = :PARM3 AND CARRIER_ACCOUNT_NO = :PARM4;
                    COMMIT; END;"
            ASCDATA1.ExecuteSQL(sql, "VDVV", {bearerToken, tokenExp, CARRIER_CODE, AccountNo})

        Catch ex As Exception
            errorMessage = "Get Authorization Token Error: " & ex.Message
        Finally
            numLoops = 0
            'ASCMAIN1.MultiTask_Release(,, MT_LEVEL)
            ASCDATA1.ExecuteSP("MULTITASK.MULTITASK_RELEASE",
                                    "VVN",
                                    {ASCMAIN1.SESSION_NO, ASCMAIN1.ActiveForm.SELECTION_NO, CInt(MT_LEVEL)},
                                    {"P_SESSION_NO", "P_SELECTION_NO", "P_MT_LEVEL"})
        End Try

        If errorMessage.Length > 0 Then
            Throw New Exception(errorMessage)
        End If

        Return bearerToken

    End Function

#End Region

    Class RootObject(Of T)
        Public Property Table As T
    End Class

End Class
