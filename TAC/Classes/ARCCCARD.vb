Imports nsoftware.InFDMS.Fdmsecommerce
Imports nsoftware.InFDMS.Cardvalidator
Imports nsoftware.InFDMS.Fdmsreversal
Imports nsoftware.InFDMS.Fdmssettle
Imports nsoftware.InFDMS.Fdmsdetailrecord
Imports nsoftware.InFDMS.Fdmsregister
Imports nsoftware.InFDMS.Fdmslevel2
Imports nsoftware.InFDMS.Fdmslevel3
Imports nsoftware.InFDMS.FDMSLineItem
Imports nsoftware.InFDMS.FDMSRecordType

Imports nsoftware.IBizPtech.Ptcharge
Imports nsoftware.IBizPtech.Ptsettle
Imports nsoftware.IBizPtech.PtchargeEntryDataSources
Imports nsoftware.IBizPtech.PtchargeGoodsIndicators

Imports nsoftware.InPay.Icharge
Imports AuthorizeNet

Imports System.Xml.Serialization
Imports System.IO
Imports System.Net

<Serializable()> _
Public Class ARCCCARD

#Region "*************** TABLES NEEDED ***************"

    'DROP TABLE ARTCCPA1 CASCADE CONSTRAINTS ; 

    'CREATE TABLE ARTCCPA1 ( 
    '  CCPA_NO                    VARCHAR2 (10), 
    '  CUST_CODE                  VARCHAR2 (10), 
    '  CCPA_STATUS                VARCHAR2 (1), 
    '  CCPA_REASON                VARCHAR2 (1), 
    '  CCPA_NOTE                  VARCHAR2 (60), 
    '  CCPA_AMT                   NUMBER (13,2), 
    '  CCPA_DATE_AUTH             DATE, 
    '  CCPA_DATE_SALE             DATE, 
    '  CCPA_AUTH                  VARCHAR2 (1), 
    '  INIT_OPER                  VARCHAR2 (8), 
    '  INIT_DATE                  DATE, 
    '  LAST_OPER                  VARCHAR2 (8), 
    '  LAST_DATE                  DATE, 
    '  CUST_CREDIT_CARD_NO        VARCHAR2 (100), 
    '  CUST_CREDIT_CARD_EXP_DATE  VARCHAR2 (80), 
    '  CUST_CREDIT_CARD_VER_CODE  VARCHAR2 (80), 
    '  CUST_CREDIT_CARD_NAME      VARCHAR2 (35), 
    '  CUST_CREDIT_CARD_ADDR1     VARCHAR2 (35), 
    '  CUST_CREDIT_CARD_CITY      VARCHAR2 (25), 
    '  CUST_CREDIT_CARD_STATE     VARCHAR2 (2), 
    '  CUST_CREDIT_CARD_ZIP_CODE  VARCHAR2 (10),
    '  CUST_CREDIT_CARD_COUNTRY   VARCHAR2(2),
    '  CUST_CREDIT_CARD_LAST4     VARCHAR2 (4), 
    '  RESPONSE_RETRIEVAL_NO      VARCHAR2 (8), 
    '  RESPONSE_CODE              VARCHAR2 (1), 
    '  RESPONSE_BATCH_NO          VARCHAR2 (15), 
    '  RESPONSE_APPROVAL_CODE     VARCHAR2 (10), 
    '  RESPONSE_TEXT              VARCHAR2 (100), 
    '  CCPA_TYPE                  VARCHAR2 (1), 
    '  ORDR_NO                    VARCHAR2 (10), 
    '  INV_NO                     VARCHAR2 (10), 
    '  CCPA_DATE_VOID             DATE, 
    '  CCPA_REASON_VOID           VARCHAR2 (60), 
    '  STMT_NO                    VARCHAR2 (10), 
    '  CCPA_NO_CREDITED           VARCHAR2 (10), 
    '  CUST_CREDIT_CARD_TYPE      VARCHAR2 (4), 
    '  OPS_YYYYPP                 VARCHAR2 (6), 
    '  LENS_BANK_INV_NO           VARCHAR2 (10), 
    '  CCPA_NO_AUTH               VARCHAR2 (10), 
    '  CCPA_NO_CAPTURE            VARCHAR2 (10), 
    '  CARD_LEVEL_RESULT          VARCHAR2 (2), 
    '  DATAWIRE_RETURN_CODE       VARCHAR2 (4), 
    '  DATAWIRE_STATUS            VARCHAR2 (20), 
    '  ACI_CODE                   VARCHAR2 (2), 
    '  TRANSACTION_DATE           VARCHAR2 (6), 
    '  TRANS_ID                   VARCHAR2 (30), 
    '  TRANS_NUM                  VARCHAR2 (10), 
    '  VALIDATION_CODE            VARCHAR2 (4), 
    '  WEB_PYMT_ID                NUMBER (20),
    'PRIMARY KEY (CCPA_NO));

    'DROP TABLE ARTCCPA2 CASCADE CONSTRAINTS ; 

    'CREATE TABLE ARTCCPA2 ( 
    '  CCPA_NO                 VARCHAR2 (10), 
    '  RESPONSE_TEXT           VARCHAR2 (100), 
    '  RESPONSE_SEQ_NO         VARCHAR2 (6), 
    '  RESPONSE_RETRIEVAL_NO   VARCHAR2 (8), 
    '  RESPONSE_CODE           VARCHAR2 (1), 
    '  RESPONSE_BATCH_NO       VARCHAR2 (6), 
    '  RESPONSE_APPROVAL_CODE  VARCHAR2 (10), 
    '  RESPONSE_DATA           VARCHAR2 (500), 
    '  RESPONSE_AVS            VARCHAR2 (1), 
    '  RESPONSE_AUTH_SOURCE    VARCHAR2 (1), 
    '  INIT_OPER               VARCHAR2 (8), 
    '  INIT_DATE               DATE, 
    '  CCPA_TYPE               VARCHAR2 (1),
    '  PRIMARY KEY (CCPA_NO));

    'DROP TABLE ARTCCPDA CASCADE CONSTRAINTS ; 

    'CREATE TABLE ARTCCPDA ( 
    '  CCPA_NO             VARCHAR2 (10)  NOT NULL, 
    '  DETAIL_AGGREGATE    VARCHAR2 (2000), 
    '  LEVEL2_AGGREGATE    VARCHAR2 (2000), 
    '  LEVEL3_AGGREGATE    VARCHAR2 (3000), 
    '  LEVEL3_AGGREGATE_2  VARCHAR2 (3000), 
    '  LEVEL3_AGGREGATE_3  VARCHAR2 (3000), 
    '  COMM_CARD_TYPE      NUMBER (3), 
    '  LEVEL3_AGGREGATE_4  VARCHAR2 (3000), 
    '  LEVEL3_AGGREGATE_5  VARCHAR2 (3000),
    '  PRIMARY KEY (CCPA_NO)); 

    'DROP TABLE ARTCCPRC CASCADE CONSTRAINTS ; 

    'CREATE TABLE ARTCCPRC ( 
    '  AR_PARM_ACCOUNT            VARCHAR2 (10)  NOT NULL, 
    '  AR_PARM_SERVER             VARCHAR2 (100), 
    '  AR_PARM_MERCHANT_NO        VARCHAR2 (20), 
    '  AR_PARM_TERMINAL_NO        VARCHAR2 (10), 
    '  AR_PARM_CLIENT_NO          VARCHAR2 (4), 
    '  AR_PARM_USER_ID            VARCHAR2 (15), 
    '  AR_PARM_PASSWORD           VARCHAR2 (25), 
    '  AR_PARM_DATAWIRE_ID        VARCHAR2 (30), 
    '  AR_PARM_FDMS_VISA_ID       VARCHAR2 (20), 
    '  AR_PARM_FDMS_MERCH_TAX_ID  VARCHAR2 (10), 
    '  AR_PARM_FDMS_DATAWIRE_URL  VARCHAR2 (30), 
    '  AR_PARM_FDMS_SERVER_ALT    VARCHAR2 (100), 
    '  AR_PARM_AUTH_TRANS_KEY     VARCHAR2 (25), 
    '  AR_PARM_CC_PROC_TYPE       VARCHAR2 (1), 
    '  PRIMARY KEY ( AR_PARM_ACCOUNT ) ) ; 

    'DROP TABLE ARTCUSTC CASCADE CONSTRAINTS ; 

    'CREATE TABLE ARTCUSTC ( 
    '  CUST_CODE                   VARCHAR2 (10)  NOT NULL, 
    '  CUST_CREDIT_CARD_NO         VARCHAR2 (100)  NOT NULL, 
    '  CUST_CREDIT_CARD_EXP_DATE   VARCHAR2 (80), 
    '  CUST_CREDIT_CARD_VER_CODE   VARCHAR2 (80), 
    '  CUST_CREDIT_CARD_NAME       VARCHAR2 (35), 
    '  CUST_CREDIT_CARD_ADDR1      VARCHAR2 (35), 
    '  CUST_CREDIT_CARD_CITY       VARCHAR2 (25), 
    '  CUST_CREDIT_CARD_STATE      VARCHAR2 (2), 
    '  CUST_CREDIT_CARD_ZIP_CODE   VARCHAR2 (10), 
    '  CUST_CREDIT_CARD_COUNTRY    VARCHAR2 (2),     
    '  CUST_CREDIT_CARD_STATUS     VARCHAR2 (1), 
    '  INIT_OPER                   VARCHAR2 (20), 
    '  INIT_DATE                   DATE, 
    '  LAST_OPER                   VARCHAR2 (20), 
    '  LAST_DATE                   DATE, 
    '  CUST_CREDIT_CARD_PREFERRED  VARCHAR2 (1), 
    '  CUST_CREDIT_CARD_LAST4      VARCHAR2 (4), 

    '  PRIMARY KEY ( CUST_CODE, CUST_CREDIT_CARD_NO ) ) ; 

#End Region

#Region "Constants"

    ' These are live urls - do not have test urls yet
    Public Const DatawireTestServer As String = "https://support.datawire.net/production_expresso/SRS.do"
    Public Const FDMSTestServer As String = "https://vxn.datawire.net/sd"
    Public Const AuthorizenetTestServer As String = "https://test.authorize.net/gateway/transact.dll"
    Public Const AuthorizenetBatchInquiryServer As String = "https://test.authorize.net/gateway/transact.dll"

    Public Const FDMSETestMerchantNumber As String = "000000999990"
    Public Const FDMSETestMerchantTerminalNumber As String = "555555"
    Public Const FDMSETestDatawireId As String = "0000B47FFFFFFFFFFFFF"

    Public Const AuthorizeNetTestAPILoginID As String = "38LzUa6N"
    Public Const AuthorizeNetTestTransactionKey As String = "3Lh4F52uFB99bMDC"
    Public Const AuthorizeNetTestAmericanExpressTestCard As String = "370000000000002"
    Public Const AuthorizeNetTestDiscoverTestCard As String = "6011000000000012"
    Public Const AuthorizeNetTestVisaTestCard As String = "4007000000027"
    Public Const AuthorizeNetTestSecondVisaTesCard As String = "4012888818888"
    Public Const AuthorizeNetTestJCB As String = "3088000000000017"
    Public Const AuthorizeNetTestDinersClubCarteBlanche As String = "38000000000006"

    Public Enum CreditCardNames
        AmericanExpress = 1
        DinersClubCarteBlanche = 2
        Discover = 3
        JCB = 4
        MasterCard = 5
        Visa = 6
    End Enum

    Private Const Approved As String = "A"
    Private Const Declined As String = "E"
    Private Const DatawireError As String = "E"
    Public Const ShipFromZip As String = "10016"

    ' Public Const epaymentRuntimeLicense As String = "42504E354141315355425241533154453345383933333331580000000000000000000000000000004A52344B5057583900004339473856563535455935380000"
    ' Public Const epaymentRuntimeLicense As String = "42504E364141315355425241533154453345383933333331580000000000000000000000000000004D554637445A525A00004B43584D555830395753554E0000"
    ' Public Const epaymentRuntimeLicense As String = "42504E364141315355425241533154453345383933333331580000000000000000000000000000004532594A58424252000032455A374B3543414D5841330000"
    Public Const epaymentRuntimeLicense As String = "42504E3641413153554252415331544533453839333333315800000000000000000000000000000059585246324D544600003358483652314743485732410000"
    ' from wjz machine                               42504E364141315355425241533154453345383933333331580000000000000000000000000000004532594A58424252000032455A374B3543414D5841330000

    Public Const fdmsRuntimeLicense As String = "42444E33414131535542524153315445334538393333333158000000000000000000000000000000413353445233454400004E354B36485A41535A5554370000"
    Public Const paymentechRuntimeLicense As String = "42544E3441414E58524633443848323234335800000000000000000000000000000000000000000041335344523345440000534237574D474E4D573938540000"

    Private Const SSLEnabledProtocols As Int32 = 4032

#End Region

#Region "Enumerations"

    Enum ProcessingTypes
        Paymentech = 1
        FDMS = 2
        AuthorizeNet = 3
    End Enum

    Enum IndustryTypes
        iteDirectMarketing = nsoftware.IBizPtech.PtchargeIndustryTypes.itDirectMarketing
        itECommerce = nsoftware.IBizPtech.PtchargeIndustryTypes.itECommerce
        itRetail = nsoftware.IBizPtech.PtchargeIndustryTypes.itRetail

        iteFDMSDirectMarketing = nsoftware.InFDMS.FdmssettleIndustryTypes.itDirectMarketing
        itFDMSGroceryStore = nsoftware.InFDMS.FdmssettleIndustryTypes.itGroceryStore
        itFDMSRestaurant = nsoftware.InFDMS.FdmssettleIndustryTypes.itRestaurant
        itFDMSRetail = nsoftware.InFDMS.FdmssettleIndustryTypes.itRetail
        itFDMSUnknown = nsoftware.InFDMS.FdmssettleIndustryTypes.itUnknown
    End Enum

    Enum EntryDataSources
        dsManualContactless = nsoftware.IBizPtech.PtchargeEntryDataSources.dsManualContactless
        dsManuallyEntered = nsoftware.IBizPtech.PtchargeEntryDataSources.dsManuallyEntered
        dsSwipeOriginUnknown = nsoftware.IBizPtech.PtchargeEntryDataSources.dsSwipeOriginUnknown
        dsTrack1 = nsoftware.IBizPtech.PtchargeEntryDataSources.dsTrack1
        dsTrack1Contactless = nsoftware.IBizPtech.PtchargeEntryDataSources.dsTrack1Contactless
        dsTrack1FromRFID = nsoftware.IBizPtech.PtchargeEntryDataSources.dsTrack1FromRFID
        dsTrack2 = nsoftware.IBizPtech.PtchargeEntryDataSources.dsTrack2
        dsTrack2Contactless = nsoftware.IBizPtech.PtchargeEntryDataSources.dsTrack2Contactless
        dsTrack2FromRFID = nsoftware.IBizPtech.PtchargeEntryDataSources.dsTrack2FromRFID
    End Enum

    Enum TransactionTypes
        ttECommerce = nsoftware.InFDMS.FdmsecommerceTransactionTypes.ttECommerce
        ttInstallment = nsoftware.InFDMS.FdmsecommerceTransactionTypes.ttInstallment
        ttMOTO = nsoftware.InFDMS.FdmsecommerceTransactionTypes.ttMOTO
        ttRecurring = nsoftware.InFDMS.FdmsecommerceTransactionTypes.ttRecurring
    End Enum

    Enum TaxTypes
        CitySalesTax = nsoftware.InFDMS.FDMSTaxTypes.ittCitySalesTax
        EnergyTax = nsoftware.InFDMS.FDMSTaxTypes.ittEnergyTax
        FederalSalesTax = nsoftware.InFDMS.FDMSTaxTypes.ittFederalSalesTax
        GoodsServicesTax = nsoftware.InFDMS.FDMSTaxTypes.ittGoodsServicesTax
        LocalSalesTax = nsoftware.InFDMS.FDMSTaxTypes.ittLocalSalesTax
        MunicipalSalesTax = nsoftware.InFDMS.FDMSTaxTypes.ittMunicipalSalesTax
        Notsupported = nsoftware.InFDMS.FDMSTaxTypes.ittNotsupported
        OccupancyTax = nsoftware.InFDMS.FDMSTaxTypes.ittOccupancyTax
        OtherTax = nsoftware.InFDMS.FDMSTaxTypes.ittOtherTax
        ProvincialSalesTax = nsoftware.InFDMS.FDMSTaxTypes.ittProvincialSalesTax
        RoomTax = nsoftware.InFDMS.FDMSTaxTypes.ittRoomTax
        StateSalesTax = nsoftware.InFDMS.FDMSTaxTypes.ittStateSalesTax
        Unknown = nsoftware.InFDMS.FDMSTaxTypes.ittUnknown
        ValueAddedTax = nsoftware.InFDMS.FDMSTaxTypes.ittValueAddedTax
    End Enum

    Enum CreditCardTypes
        vctAmex = nsoftware.InFDMS.CardvalidatorCardTypes.vctAmex
        vctBankCard = nsoftware.InFDMS.CardvalidatorCardTypes.vctBankCard
        vctCUP = nsoftware.InFDMS.CardvalidatorCardTypes.vctCUP
        vctDiners = nsoftware.InFDMS.CardvalidatorCardTypes.vctDiners
        vctDiscover = nsoftware.InFDMS.CardvalidatorCardTypes.vctDiscover
        vctJCB = nsoftware.InFDMS.CardvalidatorCardTypes.vctJCB
        vctLaser = nsoftware.InFDMS.CardvalidatorCardTypes.vctLaser
        vctMaestro = nsoftware.InFDMS.CardvalidatorCardTypes.vctMaestro
        vctMasterCard = nsoftware.InFDMS.CardvalidatorCardTypes.vctMasterCard
        vctMCardPurchase = nsoftware.InFDMS.CardvalidatorCardTypes.vctMCardPurchase
        vctSolo = nsoftware.InFDMS.CardvalidatorCardTypes.vctSolo
        vctSwitch = nsoftware.InFDMS.CardvalidatorCardTypes.vctSwitch
        vctTempoPayments = nsoftware.InFDMS.CardvalidatorCardTypes.vctTempoPayments
        vctUnknown = nsoftware.InFDMS.CardvalidatorCardTypes.vctUnknown
        vctVisa = nsoftware.InFDMS.CardvalidatorCardTypes.vctVisa
        vctVisaElectron = nsoftware.InFDMS.CardvalidatorCardTypes.vctVisaElectron
        vctVisaPurchase = nsoftware.InFDMS.CardvalidatorCardTypes.vctVisaPurchase
    End Enum

    Enum eCommerceGoodsIndicatorTypes
        Physicalgoods = nsoftware.IBizPtech.PtchargeGoodsIndicators.giPhysicalGoods
        DigitalGoods = nsoftware.IBizPtech.PtchargeGoodsIndicators.giDigitalGoods
    End Enum

#End Region

#Region "Class Variables"

    ' Authorize.Net Obiect
    Private objAuthorizeNet As nsoftware.InPay.Icharge

    ' FMDS Objects
    Private objCardValidator As nsoftware.InFDMS.Cardvalidator
    Private objFdmseCommerce As nsoftware.InFDMS.Fdmsecommerce
    Private objFdmsReversal As nsoftware.InFDMS.Fdmsreversal
    Private WithEvents objFdmsSettle As nsoftware.InFDMS.Fdmssettle
    Private objFdmsregister As nsoftware.InFDMS.Fdmsregister

    ' Paymentech Objects
    Private objPtcharge As nsoftware.IBizPtech.Ptcharge
    Private objPtsettle As nsoftware.IBizPtech.Ptsettle

    Private cEntryDataSource As EntryDataSources = EntryDataSources.dsManuallyEntered
    Private cIndustryType As IndustryTypes = IndustryTypes.iteDirectMarketing
    Private cTransactionType As TransactionTypes = TransactionTypes.ttECommerce

    Private cTransactionAmount As Double = 0
    Private cTransactionNumber As String = String.Empty
    Private cProcessingType As ProcessingTypes = ProcessingTypes.FDMS
    Private cDetailAggregate As String = String.Empty
    Private cGoodsIndicator As eCommerceGoodsIndicatorTypes
    Private cFDMSDataWireId As String = String.Empty
    Private cBatchNumber As String = String.Empty
    Private cBatchStatus As String = String.Empty

    Private cCreditCardProcessingNo As String = String.Empty
    Private cInternalReference As String = String.Empty

    Private cSettleStatusLog As DataTable

    ' Sub classes Instantiation
    Public MerchantAccount As New Merchant
    Public CustomerCreditCard As New CreditCard
    Public NetworkResponse As New Response
    Public Level2Data As New Level2
    Public Level3Data As New List(Of Level3)
    Public ChargeReversal As New Reversal
    Public Settlement As New Settle
    Private messageCounter As Int32 = 0
    Private cLogFileLocation As String = String.Empty
    Private cXmlFileName As String = String.Empty
    Private cXmlDirectory As String = String.Empty

    Private cTestMode As Boolean = False
    Private clsEncryptionClass As Object = Nothing

    Private clsLastError As String = String.Empty

    Structure FDMSServiceProviderStruct
        Dim Url As String
        Dim ResponseTime As Integer
    End Structure

    Private cFDMSServiceProviders() As FDMSServiceProviderStruct

#End Region

#Region "Class Instantiation"

    ''' <summary>
    ''' Instantiate class, defualt processing type is AuthorizeNet
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        cProcessingType = ProcessingTypes.AuthorizeNet
        cLogFileLocation = String.Empty
        InitializeVariables()
    End Sub

    ''' <summary>
    ''' Instantiate class, defualt processing type is AuthorizeNet
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal slogFileLocation As String)
        cProcessingType = ProcessingTypes.AuthorizeNet
        cLogFileLocation = slogFileLocation
        InitializeVariables()
    End Sub

    ''' <summary>
    ''' Instantiate class, defualt processing type is AuthorizeNet
    ''' </summary>
    ''' <param name="ProcType">Credit card Processing Type</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal slogFileLocation As String, ByVal ProcType As ProcessingTypes)
        ProcessingType = ProcType
        cLogFileLocation = slogFileLocation
        InitializeVariables()
    End Sub

    ''' <summary>
    ''' Instantiate class, defualt processing type is AuthorizeNet
    ''' </summary>
    ''' <param name="slogFileLocation">Location to place serialized object</param>
    ''' <param name="ProcType">Credit card Processing Type</param>
    ''' <param name="encryptionClass">Encryption Class. Must have Functions Ecrypt, Decrypt - each accepting a single string parameter
    ''' and a return value of String</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal slogFileLocation As String, ByVal ProcType As ProcessingTypes, encryptionClass As Type)
        ProcessingType = ProcType
        cLogFileLocation = slogFileLocation
        InitializeVariables()
        clsEncryptionClass = Activator.CreateInstance(encryptionClass)
    End Sub

    ''' <summary>
    ''' Initialize class object and variables
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitializeVariables()

        cEntryDataSource = EntryDataSources.dsManuallyEntered

        Select Case cProcessingType
            Case ProcessingTypes.FDMS
                cIndustryType = IndustryTypes.iteFDMSDirectMarketing
            Case ProcessingTypes.Paymentech
                cIndustryType = IndustryTypes.iteDirectMarketing
        End Select

        cTransactionType = TransactionTypes.ttECommerce
        cGoodsIndicator = eCommerceGoodsIndicatorTypes.Physicalgoods

        cTransactionAmount = 0
        cTransactionNumber = String.Empty
        cDetailAggregate = String.Empty
        cFDMSDataWireId = String.Empty
        cBatchNumber = String.Empty
        cBatchStatus = String.Empty

        ' FDMS Objects
        objCardValidator = New nsoftware.InFDMS.Cardvalidator
        objCardValidator.RuntimeLicense = fdmsRuntimeLicense

        objFdmseCommerce = New nsoftware.InFDMS.Fdmsecommerce
        objFdmseCommerce.RuntimeLicense = fdmsRuntimeLicense

        objFdmsReversal = New nsoftware.InFDMS.Fdmsreversal
        objFdmsReversal.RuntimeLicense = fdmsRuntimeLicense

        objFdmsSettle = New nsoftware.InFDMS.Fdmssettle
        objFdmsSettle.RuntimeLicense = fdmsRuntimeLicense

        objFdmsregister = New nsoftware.InFDMS.Fdmsregister
        objFdmsregister.RuntimeLicense = fdmsRuntimeLicense

        ' Paymentech Objects
        objPtcharge = New nsoftware.IBizPtech.Ptcharge(paymentechRuntimeLicense)
        objPtsettle = New nsoftware.IBizPtech.Ptsettle(paymentechRuntimeLicense)

        ' Authorize.Net Objects
        objAuthorizeNet = New nsoftware.InPay.Icharge
        objAuthorizeNet.RuntimeLicense = epaymentRuntimeLicense
        objAuthorizeNet.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)

        MerchantAccount = New Merchant
        CustomerCreditCard = New CreditCard
        Level2Data = New Level2
        Level3Data = New List(Of Level3)
        ChargeReversal = New Reversal

        NetworkResponse = New Response
        Settlement = New Settle

        cSettleStatusLog = New DataTable
        cSettleStatusLog.TableName = "SettleStatusLog"
        cSettleStatusLog.Columns.Add("Sequence", GetType(System.Int32))
        cSettleStatusLog.Columns.Add("DateTimeStamp", GetType(System.DateTime))
        cSettleStatusLog.Columns.Add("StatusMessage", GetType(System.String))

        cLogFileLocation = cLogFileLocation.Trim
        If cLogFileLocation.Length > 0 AndAlso Not My.Computer.FileSystem.DirectoryExists(cLogFileLocation) Then
            cLogFileLocation = String.Empty
        End If

        clsEncryptionClass = Nothing

    End Sub

#End Region

#Region "Properties"

    Public ReadOnly Property LastError() As String
        Get
            Return clsLastError
        End Get
    End Property

    ''' <summary>
    ''' Settlement Batch Number
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property BatchNumber() As String
        Get
            Return cBatchNumber
        End Get
        Set(ByVal value As String)
            cBatchNumber = value
        End Set
    End Property

    ''' <summary>
    ''' Settlement Batch Status
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property BatchStatus() As String
        Get
            Return cBatchStatus
        End Get
        Set(ByVal value As String)
            cBatchStatus = value
        End Set
    End Property

    ''' <summary>
    ''' Internal Company reference Number. Not used for transaction processing.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CreditCardProcessingNo() As String
        Get
            Return cCreditCardProcessingNo
        End Get
        Set(ByVal value As String)
            cCreditCardProcessingNo = value
        End Set
    End Property

    ''' <summary>
    ''' Indicates the reuslts of the Luhn Check Digit Algorithm
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property CreditCardType() As CreditCardTypes
        Get
            Try
                With objCardValidator
                    ' Since just checking the card type use the current month and year
                    .CardExpMonth = DateTime.Now.Month 'Val(CustomerCreditCard.CardExpMonth & String.Empty)
                    .CardExpYear = DateTime.Now.Year ' Val(CustomerCreditCard.CardExpYear & String.Empty)
                    .CardNumber = CustomerCreditCard.CardNumber
                    .ValidateCard()
                    Return .CardType
                End With
            Catch ex As Exception
                Return CreditCardTypes.vctUnknown
            End Try
        End Get
    End Property

    ''' <summary>
    ''' An aggregate containing details of a transaction, which is then used for settlement (FDMS)
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property DetailAggregate() As String
        Set(ByVal value As String)
            cDetailAggregate = value
        End Set
    End Property

    ''' <summary>
    ''' Determines the manner in which a cardholder's information is being sent to the Host (Paymentech)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EntryDataSource() As EntryDataSources
        Get
            Return cEntryDataSource
        End Get
        Set(ByVal value As EntryDataSources)
            cEntryDataSource = value
        End Set
    End Property

    ''' <summary>
    '''  Id which you must save and send with every FDMS subsequent transaction
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FDMSDataWireId() As String
        Get
            Return cFDMSDataWireId
        End Get
        Set(ByVal value As String)
            cFDMSDataWireId = value
        End Set
    End Property

    ''' <summary>
    ''' FDMS Service Providers used as the URL in the Merchant setup
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FDMSServiceProviders() As FDMSServiceProviderStruct()
        Get
            Return cFDMSServiceProviders
        End Get
        Set(ByVal value As FDMSServiceProviderStruct())
            cFDMSServiceProviders = value
        End Set
    End Property

    ''' <summary>
    ''' Indicates the type of goods being sold by an e-Commerce merchant (Paymentech)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property GoodsIndicatr() As eCommerceGoodsIndicatorTypes
        Get
            Return cGoodsIndicator
        End Get
        Set(ByVal value As eCommerceGoodsIndicatorTypes)
            cGoodsIndicator = value
        End Set
    End Property

    ''' <summary>
    ''' Sets / Gets the Merchant's Industry Type (Paymentech)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property IndustryType() As IndustryTypes
        Get
            Return (cIndustryType)
        End Get
        Set(ByVal value As IndustryTypes)
            cIndustryType = value
        End Set
    End Property

    ''' <summary>
    ''' Internal information. Not used for transaction processing.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property InternalReference() As String
        Get
            Return cInternalReference
        End Get
        Set(ByVal value As String)
            cInternalReference = value
        End Set
    End Property

    ''' <summary>
    ''' Sets / Gets the network to processing the credit card transaction
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ProcessingType() As ProcessingTypes
        Get
            Return cProcessingType
        End Get
        Set(ByVal value As ProcessingTypes)
            cProcessingType = value
        End Set
    End Property

    ''' <summary>
    ''' Set / Set Credit Card Transaction Amount
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TransactionAmount() As Double
        Get
            Return cTransactionAmount
        End Get
        Set(ByVal value As Double)
            cTransactionAmount = value
        End Set
    End Property

    ''' <summary>
    ''' Uniquely identifies the transaction.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TransactionNumber() As String
        Get
            Return cTransactionNumber
        End Get
        Set(ByVal value As String)
            cTransactionNumber = value
        End Set
    End Property

    ''' <summary>
    ''' Specifies the type of transaction to process.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TransactionType() As TransactionTypes
        Get
            Return cTransactionType
        End Get
        Set(ByVal value As TransactionTypes)
            cTransactionType = value
        End Set
    End Property

    ''' <summary>
    ''' Datatable of all status messages.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SettleStatusLog() As DataTable
        Get
            Return cSettleStatusLog
        End Get
        Set(ByVal value As DataTable)
            cSettleStatusLog = value
            cSettleStatusLog.TableName = "SettleStatusLog"
        End Set

    End Property

    ''' <summary>
    ''' Get the Location of the XMl file
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property XmlFileName() As String
        Get
            Return cXmlFileName
        End Get
    End Property

    ''' <summary>
    ''' Gets / Set location to place serialized transaction
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property XmlDirectory As String
        Get
            Return cXmlDirectory
        End Get
        Set(value As String)
            cXmlDirectory = value
        End Set
    End Property


    ''' <summary>
    ''' Get set whether in test mode
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TestMode() As Boolean
        Get
            Return cTestMode
        End Get
        Set(ByVal value As Boolean)
            cTestMode = value
        End Set
    End Property

    ''' <summary>
    ''' Get / Set location of logfile to Serialize Class.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LogFileLocation As String
        Get
            Return cLogFileLocation
        End Get
        Set(value As String)
            cLogFileLocation = value
        End Set
    End Property


#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Sends an AuthOnly transaction to the host
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub AuthOnly()

        clsLastError = String.Empty

        Select Case ProcessingType

            Case ProcessingTypes.FDMS
                FDMSAuthorize("Authorize")
            Case ProcessingTypes.Paymentech
                PaymentechAuthorize()
            Case ProcessingTypes.AuthorizeNet
                AuthorizeNetAuthorize()
            Case Else
                clsLastError = "Unknown Processing Network Type"""
                Throw New System.Exception("Unknown Processing Network Type")
                Exit Sub
        End Select

        ExportSerializedObject()
    End Sub

    ''' <summary>
    ''' Retrieves the current state of the open batch. 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub BatchInquiry()

        clsLastError = String.Empty

        Select Case ProcessingType

            Case ProcessingTypes.FDMS
                FDMSBatchInquiry()
            Case ProcessingTypes.Paymentech
                PaymentechBatchInquiry()
            Case ProcessingTypes.AuthorizeNet
                AuthorizeNetBatchInquiry()
            Case Else
                clsLastError = "Unknown Processing Network Type"
                Throw New System.Exception("Unknown Processing Network Type")
                Exit Sub
        End Select
        ExportSerializedObject()
    End Sub

    ''' <summary>
    ''' Releases the current batch for settlement.
    ''' </summary>
    ''' <param name="NetDeposit"></param>
    ''' <remarks></remarks>
    Public Sub BatchRelease(ByVal NetDeposit As String)

        clsLastError = String.Empty

        Select Case ProcessingType

            Case ProcessingTypes.FDMS
                FDMSBatchRelease()
            Case ProcessingTypes.Paymentech
                PaymentechBatchRelease(NetDeposit)
            Case ProcessingTypes.AuthorizeNet
                AuthorizeNetBatchRelease()
            Case Else
                clsLastError = "Unknown Processing Network Type"
                Throw New System.Exception("Unknown Processing Network Type")
                Exit Sub
        End Select
        ExportSerializedObject()
    End Sub

    ''' <summary>
    ''' Capture (Prior Sales) are typically used when a merchant has previously utilized the AuthOnly method. 
    ''' A Capture transaction adds the transaction to the current open batch, and the transaction will be settled 
    ''' at the next call to the BatchRelease method. 
    ''' </summary>
    ''' <param name="ApprovalCode"></param>
    ''' <remarks></remarks>
    Public Sub Capture(ByVal ApprovalCode As String)

        clsLastError = String.Empty

        Select Case ProcessingType
            Case ProcessingTypes.FDMS
                FDMSCapture(ApprovalCode)
            Case ProcessingTypes.Paymentech
                PaymentechCapture(ApprovalCode)
            Case ProcessingTypes.AuthorizeNet
                AuthorizeNetCapture(ApprovalCode)
            Case Else
                clsLastError = "Unknown Processing Network Type"
                Throw New System.Exception("Unknown Processing Network Type")
                Exit Sub
        End Select
        ExportSerializedObject()
    End Sub

    ''' <summary>
    ''' Credits a Credit Card Transaction
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Credit(ByVal ApprovalCode As String)

        clsLastError = String.Empty

        Select Case ProcessingType

            Case ProcessingTypes.FDMS
                FDMSCredit(ApprovalCode)
            Case ProcessingTypes.Paymentech
                PaymentechCredit()
            Case ProcessingTypes.AuthorizeNet
                AuthorizeNetCredit(ApprovalCode)
            Case Else
                clsLastError = "Unknown Processing Network Type"
                Throw New System.Exception("Unknown Processing Network Type")
                Exit Sub
        End Select
        ExportSerializedObject()
    End Sub

    ''' <summary>
    ''' Indicates whether the card is expired or not.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DateCheckPassed() As Boolean
        Return (objCardValidator.DateCheckPassed)
    End Function

    ''' <summary>
    ''' Indicates the results of the Luhn Digit Check algorithm.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DigitCheckPassed() As Boolean
        Return (objCardValidator.DigitCheckPassed)
    End Function

    ''' <summary>
    ''' Clears all properties to their default values
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Reset()
        cProcessingType = ProcessingTypes.Paymentech
        InitializeVariables()
    End Sub

    ''' <summary>
    ''' Process a Sale on the credit card
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Sale()

        clsLastError = String.Empty

        Select Case ProcessingType
            Case ProcessingTypes.FDMS
                FDMSSale()
            Case ProcessingTypes.Paymentech
                PaymentechSale()
            Case ProcessingTypes.AuthorizeNet
                AuthorizeNetSale()
            Case Else
                clsLastError = "Unknown Processing Network Type"
                Throw New System.Exception("Unknown Processing Network Type")
                Exit Sub
        End Select
        ExportSerializedObject()
    End Sub

    ''' <summary>
    ''' Checks the card number and expiration date for validity.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ValidateCard()

        Try
            With objCardValidator
                If CustomerCreditCard.CardExpMonth.Length > 0 OrElse CustomerCreditCard.CardExpYear.Length > 0 Then
                    .CardExpMonth = CustomerCreditCard.CardExpMonth
                    .CardExpYear = CustomerCreditCard.CardExpYear
                Else
                    .CardExpMonth = DateTime.Now.Month
                    .CardExpYear = DateTime.Now.Year
                End If
                .CardNumber = CustomerCreditCard.CardNumber
                .ValidateCard()
                CustomerCreditCard.CardType = objCardValidator.CardType
            End With

        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    '''  Voids a Credit card transaction
    ''' </summary>
    ''' <param name="RetrievalNumberToVoid">The RetrievalNumber of the transaction you wish to void</param>
    ''' <param name="LastRetrievalNumber">LastRetrievalNumber should be set to the last RetrievalNumber received from the Paymentech Server. If LastRetrievalNumber is left blank, the current contents of the RetrievalNumber property will be used instead. </param>
    ''' <remarks></remarks>
    Public Sub VoidTransaction(ByVal RetrievalNumberToVoid As String, ByVal LastRetrievalNumber As String)

        clsLastError = String.Empty

        Select Case ProcessingType
            Case ProcessingTypes.FDMS
                FDMSVoidReversalTrans()
            Case ProcessingTypes.Paymentech
                PaymentechVoidTrans(RetrievalNumberToVoid, LastRetrievalNumber)
            Case ProcessingTypes.AuthorizeNet
                AuthorizeNetVoidTrans(RetrievalNumberToVoid)
            Case Else
                clsLastError = "Unknown Processing Network Type"
                Throw New System.Exception("Unknown Processing Network Type")
        End Select

        ExportSerializedObject()
    End Sub

    ''' <summary>
    ''' Returns the Level 2 Addendum for Settlement Processing
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetLevel2Addendum(ByVal CommercialCardType As Int16) As String

        Try
            Dim objFDMsLevel2 As New nsoftware.InFDMS.Fdmslevel2()
            objFDMsLevel2.RuntimeLicense = fdmsRuntimeLicense

            Select Case Level2Data.CardType
                Case CreditCardTypes.vctVisa, CreditCardTypes.vctVisaElectron, CreditCardTypes.vctVisaPurchase
                    objFDMsLevel2.CardType = nsoftware.InFDMS.Fdmslevel2CardTypes.ctVisa

                Case CreditCardTypes.vctAmex
                    objFDMsLevel2.CardType = nsoftware.InFDMS.Fdmslevel2CardTypes.ctAmericanExpress

                Case CreditCardTypes.vctMasterCard, CreditCardTypes.vctMCardPurchase
                    objFDMsLevel2.CardType = nsoftware.InFDMS.Fdmslevel2CardTypes.ctMasterCard

                Case Else
                    objFDMsLevel2.CardType = nsoftware.InFDMS.Fdmslevel2CardTypes.ctUnknown

            End Select

            Select Case CommercialCardType
                Case nsoftware.InFDMS.FDMSCommercialCards.cctNotCommercial
                    objFDMsLevel2.CommercialCardType = nsoftware.InFDMS.FDMSCommercialCards.cctNotCommercial
                Case nsoftware.InFDMS.FDMSCommercialCards.cctPurchaseCard
                    objFDMsLevel2.CommercialCardType = nsoftware.InFDMS.FDMSCommercialCards.cctPurchaseCard
                Case nsoftware.InFDMS.FDMSCommercialCards.cctCorporateCard
                    objFDMsLevel2.CommercialCardType = nsoftware.InFDMS.FDMSCommercialCards.cctCorporateCard
                Case nsoftware.InFDMS.FDMSCommercialCards.cctBusinessCard
                    objFDMsLevel2.CommercialCardType = nsoftware.InFDMS.FDMSCommercialCards.cctBusinessCard
                Case Else
                    objFDMsLevel2.CommercialCardType = nsoftware.InFDMS.FDMSCommercialCards.cctUnknown
            End Select

            objFDMsLevel2.DestinationCountry = "US"
            objFDMsLevel2.DestinationState = Level2Data.DestinationState

            If Level2Data.DestinationZip.Length > 0 Then
                objFDMsLevel2.DestinationZip = Level2Data.DestinationZip
            Else
                objFDMsLevel2.DestinationZip = Level2Data.ShipFromZip
            End If

            ' Required for AMEX, to avoid error is settlement
            If objFDMsLevel2.DestinationZip.Length = 0 Then
                objFDMsLevel2.DestinationZip = "10532"
            End If

            If Level2Data.DiscountAmount > 0 Then
                objFDMsLevel2.DiscountAmount = Convert.ToString(Level2Data.DiscountAmount * 100)
            Else
                objFDMsLevel2.DiscountAmount = "000"
            End If

            objFDMsLevel2.DutyAmount = "000"

            If Level2Data.FreightAmount > 0 Then
                objFDMsLevel2.FreightAmount = Convert.ToString(Level2Data.FreightAmount * 100)
            Else
                objFDMsLevel2.FreightAmount = "000"
            End If
            objFDMsLevel2.FreightTaxAmount = "000"
            objFDMsLevel2.FreightTaxRate = "000"

            objFDMsLevel2.InvoiceNumber = Level2Data.InvoiceNumber

            objFDMsLevel2.MerchantReference = Level2Data.InvoiceNumber
            objFDMsLevel2.MerchantTaxId = Level2Data.MerchantTaxId
            If objFDMsLevel2.MerchantTaxId.Length = 0 Then
                objFDMsLevel2.MerchantTaxId = "431995076"
            End If

            objFDMsLevel2.MerchantType = Level2Data.MerchantType

            objFDMsLevel2.OrderDate = String.Empty
            If IsDate(Level2Data.OrderDate & String.Empty) AndAlso Level2Data.OrderDate > DateAdd(DateInterval.Year, -2, DateTime.Now) Then
                objFDMsLevel2.OrderDate = Level2Data.OrderDate.ToString("yyMMdd")
            Else
                objFDMsLevel2.OrderDate = DateTime.Now.ToString("yyMMdd")
            End If

            ' PurchaseIdentifier is required
            If (Level2Data.PurchaseIdentifier & String.Empty).ToString.Length > 25 Then
                objFDMsLevel2.PurchaseIdentifier = (Level2Data.PurchaseIdentifier & String.Empty).ToString.Trim.Substring(0, 25).Trim
            Else
                objFDMsLevel2.PurchaseIdentifier = (Level2Data.PurchaseIdentifier & String.Empty).ToString.Trim
            End If
            If objFDMsLevel2.PurchaseIdentifier.Length = 0 Then objFDMsLevel2.PurchaseIdentifier = Level2Data.InvoiceNumber
            If objFDMsLevel2.PurchaseIdentifier.Length = 0 Then objFDMsLevel2.PurchaseIdentifier = DateTime.Now.ToString("yyMMddhhmm")

            objFDMsLevel2.ShippedFromZip = Level2Data.ShipFromZip

            If Level2Data.TaxAmount <= 0 Then
                objFDMsLevel2.TaxAmount = Convert.ToString(Math.Round(TransactionAmount * 0.0011, 2) * 100)
            Else
                objFDMsLevel2.TaxAmount = Convert.ToString(Math.Round(Level2Data.TaxAmount * 100, 2))
            End If

            Return objFDMsLevel2.GetAddendum

        Catch ex As Exception
            If ASCMAIN1.USER_ID = "edz" Then
                MessageBox.Show(ex.Message)
            End If
            Return (String.Empty)
        End Try

    End Function

    ''' <summary>
    ''' Returns the Level 3 addendum for Settlement Processing
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetLevel3Addendum() As String


        If Level3Data Is Nothing OrElse Level3Data.Count = 0 Then
            Return String.Empty
        End If

        Try
            If Level3Data Is Nothing OrElse Level3Data.Count = 0 Then
                Return String.Empty
            End If

            Dim objFDMSlevel3 As New nsoftware.InFDMS.Fdmslevel3()
            objFDMSlevel3.RuntimeLicense = fdmsRuntimeLicense
            ' Level 3 only applicable for Visa and Mastercard
            Select Case CreditCardType
                Case CreditCardTypes.vctVisa, CreditCardTypes.vctVisaElectron, CreditCardTypes.vctVisaPurchase
                    objFDMSlevel3.CardType = nsoftware.InFDMS.Fdmslevel2CardTypes.ctVisa

                Case CreditCardTypes.vctMasterCard, CreditCardTypes.vctMCardPurchase
                    objFDMSlevel3.CardType = nsoftware.InFDMS.Fdmslevel2CardTypes.ctMasterCard

                Case Else
                    Return String.Empty

            End Select

            objFDMSlevel3.LineItems.Clear()

            For Each level3entry As Level3 In Level3Data

                If objFDMSlevel3.LineItems.Count >= 50 Then
                    Exit For
                End If

                objFDMSlevel3.LineItems.Add(New nsoftware.InFDMS.FDMSLineItem)

                With objFDMSlevel3.LineItems(objFDMSlevel3.LineItems.Count - 1)

                    .CommodityCode = "90013000"

                    .Description = level3entry.Description
                    If .Description.Length = 0 Then
                        .Description = "Eyewear" ' Required from MC
                    End If
                    If .Description.Length > 26 Then
                        .Description = .Description.Substring(0, 26).Trim
                    End If

                    If level3entry.DiscountAmount > 0 Then
                        .DiscountAmount = Convert.ToString(Math.Round(level3entry.DiscountAmount, 2) * 100)
                    Else
                        .DiscountAmount = "000"
                    End If

                    .ProductCode = level3entry.ProductCode
                    If .ProductCode.Length = 0 Then
                        .ProductCode = "PCC" ' required from MC
                    End If
                    If .ProductCode.Length > 12 Then
                        .ProductCode = .ProductCode.Substring(0, 12).Trim
                    End If

                    .Quantity = level3entry.Quantity.ToString
                    If .Quantity = 0 Then
                        .Quantity = 1 ' Required for MC
                    End If

                    If level3entry.TaxAmount > 0 Then
                        .TaxAmount = Convert.ToString(Math.Round(level3entry.TaxAmount * 100, 2))
                    Else
                        .TaxAmount = Convert.ToString(Math.Round(level3entry.Total * 0.0011, 2) * 100)
                    End If


                    .TaxIncluded = True
                    .TaxRate = Convert.ToString(Math.Round(level3entry.TaxRate * 100, 2))
                    .TaxType = level3entry.TaxType
                    .Total = Convert.ToString(Math.Round(level3entry.Total + level3entry.TaxAmount, 2) * 100)

                    .UnitCost = Convert.ToString(Math.Round(level3entry.UnitCost, 2) * 100)
                    .Units = level3entry.Units.ToString

                    Select Case CreditCardType
                        Case CreditCardTypes.vctMasterCard
                            '.TaxType = level3entry.TaxType

                        Case CreditCardTypes.vctVisa
                            '90013000 Contact(lenses)
                            '90014000 Spectacle lenses of glass 
                            '90015000 Spectacle lenses of other materials
                            '.CommodityCode = "90013000"
                    End Select

                End With
            Next

            Return objFDMSlevel3.GetAddendum

        Catch ex As Exception
            If ASCMAIN1.USER_ID = "edz" Then
                MessageBox.Show(ex.Message)
            End If
            Return String.Empty
        End Try


    End Function

#End Region

#Region "Serialization"

    ''' <summary>
    ''' Serial Class to XML
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Serialize() As String

        Try
            Dim objXS As New XmlSerializer(Me.GetType())
            Dim objSW As New StringWriter

            objXS.Serialize(objSW, Me)

            Dim serializedObject As String = objSW.ToString()

            If clsEncryptionClass IsNot Nothing Then
                serializedObject = clsEncryptionClass.Encrypt(serializedObject)
            End If

            Return serializedObject

        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Deserialize Class Object
    ''' </summary>
    ''' <param name="strXML"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeSerialize(ByVal strXML As String) As ARCCCARD

        Try
            Dim objXS As New XmlSerializer(GetType(ARCCCARD))
            Dim objSR As New StringReader(strXML)

            Dim objST As ARCCCARD = CType(objXS.Deserialize(objSR), ARCCCARD)

            Return objST

        Catch ex As Exception
            Return New ARCCCARD(String.Empty)
        End Try
    End Function

    ''' <summary>
    ''' Save contents of object to a log
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ExportSerializedObject()
        cXmlFileName = String.Empty

        If (Me.TransactionNumber & String.Empty).ToString.Trim.Length > 0 Then
            cXmlFileName = Me.TransactionNumber.Trim
        Else ' hopefully this will never fire
            cXmlFileName = "D_" & DateTime.Now.ToString("yyyyMMdd_hhmmss")
        End If

        ExportSerializedObject(cXmlFileName, cXmlDirectory)
    End Sub

    ''' <summary>
    ''' Save contents of object to a log
    ''' </summary>
    ''' <param name="filename"></param>
    ''' <remarks></remarks>
    Public Sub ExportSerializedObject(ByRef filename As String, ByVal filePath As String)
        Try

            Dim RandomClass As New Random()
            Dim RandomNumber As Integer

            If filePath.Length = 0 Then
                filePath = cLogFileLocation.Trim
            End If

            If filePath.Length > 0 AndAlso Not filePath.EndsWith("\") Then
                filePath &= "\"
            End If

            If filePath.Length = 0 OrElse Not My.Computer.FileSystem.DirectoryExists(filePath) Then
                Exit Sub
            End If

            ' Try to get a Unique value that helps identiffy the transaction.
            ' If not available then use the date time
            If filename.Length = 0 Then
                If Me.cCreditCardProcessingNo.Trim.Length > 0 Then
                    filename = "C" & Me.cCreditCardProcessingNo.Trim
                ElseIf Me.TransactionNumber.Trim.Length > 0 Then
                    filename = "T" & Me.TransactionNumber.Trim
                Else
                    RandomNumber = RandomClass.Next(25000)
                    filename = DateTime.Now.ToString("yyyyMMdd_hhmmss") & "_" & RandomNumber.ToString
                End If
            End If

            filename = filename.ToUpper
            filename = filename.Replace(".XML", String.Empty)

            ' Ensure a unique filename
            Dim randonValue As String = String.Empty
            Dim ictr As Integer = 0
            While My.Computer.FileSystem.FileExists(filename & randonValue & ".xml")
                ictr += 1
                RandomNumber = RandomClass.Next(25000)
                randonValue = "_" & RandomNumber.ToString.Trim
                ' Prevent a run away loop
                If ictr > 25 Then Exit While
            End While

            filename &= randonValue & ".xml"
            Dim serializedData As String = Serialize()

            ' Allow append incase for some reason the file uniqueness fails
            Using writer As StreamWriter = New StreamWriter(filePath & filename, True)
                writer.WriteLine(serializedData)
                writer.Close()
            End Using

        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Deserializes an Object stored on disk
    ''' </summary>
    ''' <param name="filePath">Loaction of object on disk</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ImportSerializedObject(ByVal filePath As String) As ARCCCARD

        If Not My.Computer.FileSystem.FileExists(filePath) Then
            Return Nothing
        End If

        ' Read in the serialized class
        Dim strXML As String
        Dim tReader As IO.TextReader = New IO.StreamReader(filePath)
        strXML = tReader.ReadToEnd

        If clsEncryptionClass IsNot Nothing Then
            strXML = strXML.Trim
            strXML = clsEncryptionClass.Decrypt(strXML)
        End If

        Return DeSerialize(strXML)
    End Function

#End Region

#Region "AuthorizeNet Methods"

    Private Sub AuthorizeNetAuthorize()
        AuthorizeNetAuthOnlySale("AUTH_ONLY")
    End Sub

    Private Sub AuthorizeNetBatchInquiry()

        clsLastError = String.Empty

        Try
            AuthorizeNetMerchantSetup()
            Dim reportingGateway As New AuthorizeNet.ReportingGateway(objAuthorizeNet.MerchantLogin, objAuthorizeNet.MerchantPassword)

            Dim responseList As New List(Of AuthorizeNet.Transaction)
            responseList = reportingGateway.GetUnsettledTransactionList()

            Settlement = New Settle(responseList)

        Catch exc As nsoftware.InPay.InPayException
            clsLastError = "AuthorizeNetBatchInquiry: " & exc.Message
            Throw New System.Exception("AuthorizeNetBatchInquiry: " & exc.Message)

        Catch ex As Exception
            clsLastError = "AuthorizeNetBatchInquiry: " & ex.Message
            Throw New System.Exception("AuthorizeNetBatchInquiry: " & ex.Message)
        End Try
    End Sub

    Private Sub AuthorizeNetBatchRelease()

        clsLastError = String.Empty

        Try
            AuthorizeNetMerchantSetup()
            Dim reportingGateway As New AuthorizeNet.ReportingGateway(objAuthorizeNet.MerchantLogin, objAuthorizeNet.MerchantPassword)


        Catch exc As nsoftware.InPay.InPayException
            clsLastError = "AuthorizeNetBatchRelease: " & exc.Message
            Throw New System.Exception("AuthorizeNetBatchRelease: " & exc.Message)

        Catch ex As Exception
            clsLastError = "AuthorizeNetBatchRelease: " & ex.Message
            Throw New System.Exception("AuthorizeNetBatchRelease: " & ex.Message)
        End Try

    End Sub

    Private Sub AuthorizeNetCapture(ByVal TransactionId As String)

        clsLastError = String.Empty

        Try
            AuthorizeNetMerchantSetup()
            objAuthorizeNet.Capture(TransactionId, FormatNumber(TransactionAmount, 2))
            AuthorizeNetResponse(objAuthorizeNet.Response)

        Catch exc As nsoftware.InPay.InPayException
            clsLastError = "AuthorizeNetCapture: Error [" & exc.Code.ToString & "]: " & exc.Message
            Throw New System.Exception("AuthorizeNetCapture: Error [" & exc.Code.ToString & "]: " & exc.Message)

        Catch ex As Exception
            clsLastError = "AuthorizeNetCapture: " & ex.Message
            Throw New Exception("AuthorizeNetCapture: " & ex.Message)
        End Try

    End Sub

    Private Sub AuthorizeNetCredit(ByVal TransactionId As String)

        ' This method refunds a transaction that has already been captured, or settled. 
        ' If the transaction is still outstanding use the VoidTransaction method instead. 
        ' The TransactionId parameter indicates to the Gateway which transaction is to be refunded, and should contain the TransactionId from the original transaction. 
        ' The RefundAmount parameter is the value to be refunded back to the customer, and can be all or part of the original TransactionAmount 

        clsLastError = String.Empty

        Try
            AuthorizeNetMerchantSetup()
            objAuthorizeNet.Card.Number = CustomerCreditCard.CardNumber
            objAuthorizeNet.Refund(TransactionId, FormatNumber(TransactionAmount, 2))
            AuthorizeNetResponse(objAuthorizeNet.Response)
        Catch ex As Exception
            clsLastError = "AuthorizeNetCredit: " & ex.Message
            Throw New Exception("AuthorizeNetCredit: " & ex.Message)
        End Try
    End Sub

    Private Sub AuthorizeNetMerchantSetup()
        objAuthorizeNet.Reset()
        objAuthorizeNet.RuntimeLicense = epaymentRuntimeLicense
        objAuthorizeNet.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
        With objAuthorizeNet
            .Gateway = nsoftware.InPay.IchargeGateways.gwAuthorizeNet
            .GatewayURL = MerchantAccount.Url
            .MerchantLogin = MerchantAccount.UserID
            .MerchantPassword = MerchantAccount.Password()
        End With
    End Sub

    Private Sub AuthorizeNetResponse(response As nsoftware.InPay.EPResponse)
        NetworkResponse = New Response(response)
    End Sub

    Private Sub AuthorizeNetSale()
        AuthorizeNetAuthOnlySale("SALE")
    End Sub

    Private Sub AuthorizeNetVoidTrans(ByVal TransactionId As String)

        clsLastError = String.Empty

        Try
            AuthorizeNetMerchantSetup()
            objAuthorizeNet.VoidTransaction(TransactionId)
            AuthorizeNetResponse(objAuthorizeNet.Response)

        Catch exc As nsoftware.InPay.InPayException
            clsLastError = "AuthorizeNetVoidTrans: " & exc.Message
            Throw New Exception("AuthorizeNetVoidTrans: " & exc.Message)

        Catch ex As Exception
            clsLastError = "AuthorizeNetVoidTrans: " & ex.Message
            Throw New Exception("AuthorizeNetVoidTrans: " & ex.Message)
        End Try
    End Sub

    Private Sub AuthorizeNetAuthOnlySale(ByVal transType As String)

        clsLastError = String.Empty

        Try
            AuthorizeNetMerchantSetup()

            Dim CardExpYear As String = String.Empty
            CardExpYear = CustomerCreditCard.CardExpYear
            If CardExpYear.Length > 2 Then
                CardExpYear = CardExpYear.Substring(CardExpYear.Length - 2)
            End If

            With objAuthorizeNet

                '.AddSpecialField("x_market_type", "2") ' Retail
                .MerchantLogin = MerchantAccount.UserID
                .MerchantPassword = MerchantAccount.Password

                .Card.Number = CustomerCreditCard.CardNumber
                .Card.ExpMonth = CustomerCreditCard.CardExpMonth
                .Card.ExpYear = CardExpYear
                '.Card.CVV2Data = CustomerCreditCard.CardCVVData
                Select Case CustomerCreditCard.CardType
                    Case CreditCardTypes.vctAmex : .Card.CardType = nsoftware.InPay.TCardTypes.ctAMEX
                    Case CreditCardTypes.vctBankCard
                    Case CreditCardTypes.vctCUP
                    Case CreditCardTypes.vctDiners : .Card.CardType = nsoftware.InPay.TCardTypes.ctDiners
                    Case CreditCardTypes.vctDiscover : .Card.CardType = nsoftware.InPay.TCardTypes.ctDiscover
                    Case CreditCardTypes.vctJCB : .Card.CardType = nsoftware.InPay.TCardTypes.ctJCB
                    Case CreditCardTypes.vctLaser : .Card.CardType = nsoftware.InPay.TCardTypes.ctLaser
                    Case CreditCardTypes.vctMaestro : .Card.CardType = nsoftware.InPay.TCardTypes.ctMaestro
                    Case CreditCardTypes.vctMasterCard : .Card.CardType = nsoftware.InPay.TCardTypes.ctMasterCard
                        'Case CreditCardTypes.vctMCardPurchase : .Card.CardType = nsoftware.InPay.TCardTypes.ctMCardPurchase
                        'Case CreditCardTypes.vctSolo : .Card.CardType = nsoftware.InPay.TCardTypes.ctSolo
                        'Case CreditCardTypes.vctSwitch : .Card.CardType = nsoftware.InPay.TCardTypes.ctSwitch
                    Case CreditCardTypes.vctTempoPayments
                    Case CreditCardTypes.vctVisa : .Card.CardType = nsoftware.InPay.TCardTypes.ctVisa
                    Case CreditCardTypes.vctVisaElectron : .Card.CardType = nsoftware.InPay.TCardTypes.ctVisaElectron
                        'Case CreditCardTypes.vctVisaPurchase : .Card.CardType = nsoftware.InPay.TCardTypes.ctVisaPurchase
                End Select

                .Customer.Id = "1"
                .Customer.FirstName = CustomerCreditCard.CardHolderFirstName
                .Customer.LastName = CustomerCreditCard.CardHolderLastName
                .Customer.Address = CustomerCreditCard.CardHolderAddress
                .Customer.City = CustomerCreditCard.CardHolderCity
                .Customer.State = CustomerCreditCard.CardHolderState
                .Customer.Zip = CustomerCreditCard.CardHolderZipCode
                .Customer.Country = CustomerCreditCard.CardHolderCountry
                .Customer.Email = CustomerCreditCard.CardHolderEmail

                .Customer.Phone = CustomerCreditCard.CardHolderTelephone
                .InvoiceNumber = TransactionNumber
                .TransactionAmount = FormatNumber(TransactionAmount, 2).ToString
                .TransactionId = TransactionNumber

                If transType = "AUTH_ONLY" Then
                    .AuthOnly() ' perform Authorization Only
                Else
                    .Sale() ' perform Sale
                End If

                'AuthorizeNetResponse(.Response)

            End With

        Catch exc As nsoftware.InPay.InPayException
            clsLastError = "AuthorizeNetAuthOnlySale: " & exc.Message
            Throw New System.Exception("AuthorizeNetAuthOnlySale: " & exc.Message)
        Catch ex As Exception
            clsLastError = "AuthorizeNetAuthOnlySale: " & ex.Message
            Throw New System.Exception("AuthorizeNetAuthOnlySale: " & ex.Message)
        Finally '
            AuthorizeNetResponse(objAuthorizeNet.Response)
        End Try
    End Sub

    ''' <summary>
    ''' Returns Transaction Attributes
    ''' </summary>
    ''' <param name="TransactionID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AuthorizeNetGetTransactionDetails(ByVal TransactionID As String) As AuthorizeNet.Transaction

        clsLastError = String.Empty

        Try
            AuthorizeNetMerchantSetup()
            Dim reportingGateway As New AuthorizeNet.ReportingGateway(objAuthorizeNet.MerchantLogin, objAuthorizeNet.MerchantPassword)

            Dim response As AuthorizeNet.Transaction
            response = reportingGateway.GetTransactionDetails(TransactionID)
            Return response

        Catch exc As nsoftware.InPay.InPayException
            clsLastError = "AuthorizeNetGetTransactionDetails: " & exc.Message
            Throw New System.Exception("AuthorizeNetGetTransactionDetails: " & exc.Message)

        Catch ex As Exception
            clsLastError = "AuthorizeNetGetTransactionDetails: " & ex.Message
            Throw New System.Exception("AuthorizeNetGetTransactionDetails: " & ex.Message)

        End Try
    End Function

    ''' <summary>
    ''' Returns a list of transactions created within a date range
    ''' </summary>
    ''' <param name="fromDate"></param>
    ''' <param name="toDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AuthorizeNetGetTransactionList(ByVal fromDate As Date, ByVal toDate As Date) As List(Of AuthorizeNet.Transaction)

        clsLastError = String.Empty

        Try
            AuthorizeNetMerchantSetup()
            Dim reportingGateway As New AuthorizeNet.ReportingGateway(objAuthorizeNet.MerchantLogin, objAuthorizeNet.MerchantPassword)

            Dim response As List(Of AuthorizeNet.Transaction)
            response = reportingGateway.GetTransactionList(fromDate, toDate)
            Return response

        Catch exc As nsoftware.InPay.InPayException
            clsLastError = "AuthorizeNetGetTransactionList: " & exc.Message
            Throw New System.Exception("AuthorizeNetGetTransactionList: " & exc.Message)

        Catch ex As Exception
            clsLastError = "AuthorizeNetGetTransactionList: " & ex.Message
            Throw New System.Exception("AuthorizeNetGetTransactionList: " & ex.Message)

        End Try
    End Function

    ''' <summary>
    ''' Get Settled Batch List
    ''' </summary>
    ''' <param name="fromDate"></param>
    ''' <param name="toDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AuthorizeNetGetSettledBatchList(ByVal fromDate As Date, ByVal toDate As Date) As List(Of AuthorizeNet.Batch)

        clsLastError = String.Empty

        Try
            AuthorizeNetMerchantSetup()
            Dim reportingGateway As New AuthorizeNet.ReportingGateway(objAuthorizeNet.MerchantLogin, objAuthorizeNet.MerchantPassword)

            Dim response As List(Of AuthorizeNet.Batch)
            response = reportingGateway.GetSettledBatchList(fromDate, toDate, True)
            Return response

        Catch exc As nsoftware.InPay.InPayException
            clsLastError = "AuthorizeNetGetSettledBatchList: " & exc.Message
            Throw New System.Exception("AuthorizeNetGetSettledBatchList: " & exc.Message)

        Catch ex As Exception
            clsLastError = "AuthorizeNetGetSettledBatchList: " & ex.Message
            Throw New System.Exception("AuthorizeNetGetSettledBatchList: " & ex.Message)

        End Try
    End Function

#End Region

#Region "Paymentech Methods"

    Private Sub PaymentechAuthorize()

        clsLastError = String.Empty

        ' Paymentech offers an Authorization Only option
        Try
            objPtcharge = New nsoftware.IBizPtech.Ptcharge
            PaymentechMerchantSetupPtCharge(objPtcharge)
            NetworkResponse = Nothing

            objPtcharge.CardNumber = CustomerCreditCard.CardNumber
            objPtcharge.CardExpMonth = CustomerCreditCard.CardExpMonth
            objPtcharge.CardExpYear = CustomerCreditCard.CardExpYear
            objPtcharge.IndustryType = Me.IndustryType
            objPtcharge.EntryDataSource = Me.EntryDataSource
            objPtcharge.TransactionAmount = Me.TransactionAmount
            objPtcharge.InvoiceNumber = Me.TransactionNumber
            objPtcharge.Level2SalesTax = Me.Level2Data.TaxAmount.ToString("0.00")
            objPtcharge.Level2PurchaseId = Me.Level2Data.PurchaseIdentifier & String.Empty
            objPtcharge.Level2ShipToZip = Me.Level2Data.DestinationZip & String.Empty
            objPtcharge.CustomerAddress = CustomerCreditCard.CardHolderAddress
            objPtcharge.CustomerZip = CustomerCreditCard.CardHolderZipCode

            objPtcharge.AuthOnly()
            PaymentechResponse(objPtcharge)

        Catch ex As Exception
            clsLastError = "PaymentechAuthorize: " & ex.Message
            Throw New System.Exception("PaymentechAuthorize: " & ex.Message)
        End Try
    End Sub

    Private Sub PaymentechBatchInquiry()

        clsLastError = String.Empty

        Try
            With objPtsettle
                PaymentechMerchantSetupPtCharge(objPtsettle)
                .BatchInquiry()
                Settlement = New Settle(objPtsettle)
            End With

        Catch ex As Exception
            clsLastError = "PaymentechBatchInquiry: " & ex.Message
            Throw New Exception("PaymentechBatchInquiry: " & ex.Message)
        End Try

    End Sub

    Private Sub PaymentechBatchRelease(ByVal NetDeposit As String)
        objPtsettle.BatchRelease(NetDeposit)
    End Sub

    Private Sub PaymentechCapture(ByVal ApprovalCode As String)

        PaymentechMerchantSetupPtCharge(objPtcharge)
        With objPtcharge
            .CardNumber = CustomerCreditCard.CardNumber
            .CardExpMonth = CustomerCreditCard.CardExpMonth
            .CardExpYear = CustomerCreditCard.CardExpYear
            .TransactionAmount = TransactionAmount
            .IndustryType = Me.IndustryType
            .EntryDataSource = Me.EntryDataSource
            .InvoiceNumber = TransactionNumber
            .Level2SalesTax = Me.Level2Data.TaxAmount.ToString("0.00")
            .CustomerAddress = CustomerCreditCard.CardHolderAddress
            .CustomerZip = CustomerCreditCard.CardHolderZipCode
            .Capture(ApprovalCode)
            PaymentechResponse(objPtcharge)
        End With

    End Sub

    Private Sub PaymentechCredit()

        clsLastError = String.Empty

        Try
            PaymentechMerchantSetupPtCharge(objPtcharge)

            With objPtcharge
                .TransactionAmount = ChargeReversal.TransactionAmount
                .InvoiceNumber = ChargeReversal.TransactionNumber
                .Credit()
                PaymentechResponse(objPtcharge)
            End With

        Catch ex As Exception
            clsLastError = "PaymentechCredit: " & ex.Message
            Throw New Exception("PaymentechCredit: " & ex.Message)
        End Try
    End Sub

    Private Sub PaymentechMerchantSetupPtCharge(ByRef obj As Object)

        If TypeOf (obj) Is nsoftware.IBizPtech.Ptsettle Then
            obj.Server = MerchantAccount.Url
            obj.MerchantNumber = MerchantAccount.MerchantNumber()
            obj.TerminalNumber = MerchantAccount.MerchantTerminalNumber()
            obj.ClientNumber = MerchantAccount.ClientNumber()
            obj.UserId = MerchantAccount.UserID()
            obj.Password = MerchantAccount.Password()
        ElseIf TypeOf (obj) Is nsoftware.IBizPtech.Ptcharge Then
            obj.Server = MerchantAccount.Url
            obj.MerchantNumber = MerchantAccount.MerchantNumber()
            obj.TerminalNumber = MerchantAccount.MerchantTerminalNumber()
            obj.ClientNumber = MerchantAccount.ClientNumber()
            obj.UserId = MerchantAccount.UserID()
            obj.Password = MerchantAccount.Password()
        End If

    End Sub

    Private Sub PaymentechMerchantSetupPtSettle(ByRef objPtsettle As nsoftware.IBizPtech.Ptsettle)
        With objPtcharge
            .Server = MerchantAccount.Url
            .MerchantNumber = MerchantAccount.MerchantNumber()
            .TerminalNumber = MerchantAccount.MerchantTerminalNumber()
            .ClientNumber = MerchantAccount.ClientNumber()
            .UserId = MerchantAccount.UserID()
            .Password = MerchantAccount.Password()
        End With
    End Sub

    Private Sub PaymentechResponse(ByRef objPtcharge As nsoftware.IBizPtech.Ptcharge)
        NetworkResponse = New Response(objPtcharge)
    End Sub

    Private Sub PaymentechSale()

        clsLastError = String.Empty

        ' This is for a Sale. This is just the sale, Not the settlement which finalizes the sale
        Try
            PaymentechMerchantSetupPtCharge(objPtcharge)
            With objPtcharge
                .TransactionAmount = TransactionAmount
                .InvoiceNumber = TransactionNumber
                .Level2SalesTax = Me.Level2Data.TaxAmount.ToString("0.00")
                .Level2PurchaseId = Me.Level2Data.PurchaseIdentifier & String.Empty
                .Level2ShipToZip = Me.Level2Data.DestinationZip & String.Empty
                .CustomerAddress = CustomerCreditCard.CardHolderAddress
                .CustomerZip = CustomerCreditCard.CardHolderZipCode
                .Sale()
                PaymentechResponse(objPtcharge)
            End With

        Catch ex As Exception
            clsLastError = "PaymentechSale: " & ex.Message
            Throw New Exception("PaymentechSale: " & ex.Message)
        End Try

    End Sub

    Private Sub PaymentechVoidTrans(ByVal RetrievalNumberToVoid As String, ByVal LastRetrievalNumber As String)
        PaymentechMerchantSetupPtCharge(objPtcharge)
        With objPtcharge
            .VoidTransaction(RetrievalNumberToVoid, LastRetrievalNumber)
        End With
    End Sub

#End Region

#Region "FDMS Methods"

    Private Sub FDMSAuthorize(ByVal text As String)

        clsLastError = String.Empty

        Try
            FDMSMerchantSetup(objFdmseCommerce)
            NetworkResponse = Nothing

            objFdmseCommerce.CustomerAddress = Me.CustomerCreditCard.CardHolderAddress
            objFdmseCommerce.CustomerZip = Me.CustomerCreditCard.CardHolderZipCode

            objFdmseCommerce.TransactionNumber = Me.TransactionNumber
            objFdmseCommerce.TransactionAmount = Convert.ToString(Me.TransactionAmount * 100)
            objFdmseCommerce.TransactionType = Me.TransactionType

            With CustomerCreditCard
                objFdmseCommerce.Authorize(.CardNumber, .CardExpMonth, .CardExpYear, .CardCVVData)
            End With

            FDMSResponse(objFdmseCommerce, text)

        Catch exc As nsoftware.InFDMS.InFDMSException
            clsLastError = "FDMSAuthorize: Error [" & exc.Code.ToString & "]: " & exc.Message
            Throw New System.Exception("FDMSAuthorize: Error [" & exc.Code.ToString & "]: " & exc.Message)

        Catch ex As Exception
            clsLastError = "FDMSAuthorize: " & ex.Message
            Throw New System.Exception("FDMSAuthorize: " & ex.Message)
        End Try
    End Sub

    Private Sub FDMSBatchRelease()

        clsLastError = String.Empty

        Try
            cBatchNumber = String.Empty
            cBatchStatus = String.Empty
            NetworkResponse = Nothing

            cSettleStatusLog.Clear()

            FDMSMerchantSetup(objFdmsSettle)
            objFdmsSettle.IndustryType = cIndustryType '  nsoftware.InFDMS.FdmssettleIndustryTypes.itDirectMarketing
            objFdmsSettle.BatchSequenceNumber = DateTime.Now.ToString("yyyyMMdd")
            objFdmsSettle.DetailRecords.Clear()

            For Each settlementDelegate As TAC.ARCCCARD.Settle.FDMSSettlementDelegates In Settlement.FdmsSettlementDetailRecords
                Dim levelAggregrate As String = String.Empty

                ' Level 3 data may only be sent if there is Level 2 data
                If settlementDelegate.Level2.Length > 0 Then
                    levelAggregrate = settlementDelegate.Level2.Trim & settlementDelegate.level3.Trim
                End If

                objFdmsSettle.DetailRecords.Add(New nsoftware.InFDMS.FDMSRecordType(settlementDelegate.Detail, levelAggregrate))
            Next

            ' Speeds up the settlement
            objFdmsSettle.Config("ReuseSSLSession=true")
            objFdmsSettle.SendSettlement()
            Settlement = New Settle(objFdmsSettle)

            cBatchStatus = objFdmsSettle.Response.BatchStatus & String.Empty
            cBatchNumber = objFdmsSettle.Response.BatchNumber & String.Empty

        Catch exc As nsoftware.InFDMS.InFDMSException
            clsLastError = "FDMSBatchRelease: Error [" & exc.Code.ToString & "]: " & exc.Message
            Throw New System.Exception("FDMSBatchRelease: Error [" & exc.Code.ToString & "]: " & exc.Message)

        Catch ex As Exception
            clsLastError = "FDMSBatchRelease: " & ex.Message
            Throw New Exception("FDMSBatchRelease: " & ex.Message)
        End Try

    End Sub

    Private Sub FDMSBatchInquiry()

        clsLastError = String.Empty

        Try
            FDMSMerchantSetup(objFdmsSettle)
            Settlement = New Settle(objFdmsSettle)

        Catch exc As nsoftware.InFDMS.InFDMSException
            clsLastError = "FDMSBatchInquiry: Error [" & exc.Code.ToString & "]: " & exc.Message
            Throw New System.Exception("FDMSBatchInquiry: Error [" & exc.Code.ToString & "]: " & exc.Message)

        Catch ex As Exception
            clsLastError = "FDMSBatchInquiry: " & ex.Message
            Throw New Exception("FDMSBatchInquiry: " & ex.Message)
        End Try

    End Sub

    Private Sub FDMSCapture(ByVal ApprovalCode As String)

        clsLastError = String.Empty

        Try
            ' FDMS only has Approvals and sales. Assume the Capture is successful for the Previous Approval
            NetworkResponse = New Response(ApprovalCode, Approved)

        Catch exc As nsoftware.InFDMS.InFDMSException
            clsLastError = "FDMSCapture: Error [" & exc.Code.ToString & "]: " & exc.Message
            Throw New System.Exception("FDMSCapture: Error [" & exc.Code.ToString & "]: " & exc.Message)

        Catch ex As Exception
            clsLastError = "FDMSCapture: " & ex.Message
            Throw New Exception("FDMSCapture: " & ex.Message)
        End Try

    End Sub

    Private Sub FDMSCredit(ByVal ApprovalCode As String)

        clsLastError = String.Empty

        ' Credit Transactions are off line transactions. This means there is no authorization of funds.
        ' We must manually add these transactions to the settlement batch
        Try
            NetworkResponse = New Response(ApprovalCode, Approved, CustomerCreditCard, Me.TransactionAmount)

        Catch exc As nsoftware.InFDMS.InFDMSException
            clsLastError = "FDMSCredit: Error [" & exc.Code.ToString & "]: " & exc.Message
            Throw New System.Exception("FDMSCredit: Error [" & exc.Code.ToString & "]: " & exc.Message)

        Catch ex As Exception
            clsLastError = "FDMSCredit: " & ex.Message
            Throw New Exception("FDMSCredit: " & ex.Message)
        End Try
    End Sub

    Private Sub FDMSMerchantSetup(ByRef obj As Object)

        If TypeOf (obj) Is nsoftware.InFDMS.Fdmsecommerce Then
            obj.URL = MerchantAccount.Url()
            obj.MerchantNumber = MerchantAccount.MerchantNumber
            obj.MerchantTerminalNumber = MerchantAccount.MerchantTerminalNumber
            obj.DatawireId = MerchantAccount.DatawireId()
            obj.VisaIdentifier = MerchantAccount.VisaIdentifier()
        ElseIf TypeOf (obj) Is nsoftware.InFDMS.Fdmsreversal Then
            obj.URL = MerchantAccount.Url()
            obj.MerchantNumber = MerchantAccount.MerchantNumber
            obj.MerchantTerminalNumber = MerchantAccount.MerchantTerminalNumber
            obj.DatawireId = MerchantAccount.DatawireId()
        ElseIf TypeOf (obj) Is nsoftware.InFDMS.Fdmssettle Then
            obj.URL = MerchantAccount.Url()
            obj.MerchantNumber = MerchantAccount.MerchantNumber
            obj.MerchantTerminalNumber = MerchantAccount.MerchantTerminalNumber
            obj.DatawireId = MerchantAccount.DatawireId()
        ElseIf TypeOf (obj) Is nsoftware.InFDMS.Fdmsregister Then
            obj.URL = MerchantAccount.Url()
            obj.MerchantNumber = MerchantAccount.MerchantNumber
            obj.MerchantTerminalNumber = MerchantAccount.MerchantTerminalNumber
            obj.DatawireId = MerchantAccount.DatawireId()
        End If

    End Sub

    ''' <summary>
    ''' To use FDMS, FDMS requires that you first register and activate a merchant account with Datawire, 
    ''' which is the real time transaction delivery network that FMDS relies on. 
    ''' It requires a Unique Transaction Number
    ''' </summary>
    ''' <param name="RegistrationTransactionNumber">Unique ID number</param>
    ''' <param name="ActivationTransactionNumber">Unique ID number</param>
    ''' <remarks></remarks>
    Public Sub FDMSDatawireRegistry(ByVal Register As Boolean, ByVal RegistrationTransactionNumber As String, ByVal ActivationTransactionNumber As String, _
                                    ByVal PrimaryDiscoveryURL As String, ByVal SecondaryDiscoveryURL As String)

        clsLastError = String.Empty

        Try

            FDMSMerchantSetup(objFdmsregister)
            objFdmsregister.TransactionNumber = RegistrationTransactionNumber

            If Register Then

                objFdmsregister.Register()

                ' If the registration was successful, the DatawireStatus will contain "OK", 
                ' and the DatawireId will contain the new Id which you must save and send with every subsequent transaction. 
                ' The PrimaryDiscoveryURL and SecondaryDiscoveryURL properties will also be filled after a successful authorization. 
                ' These URLs should be saved for later, and used to perform ServiceDiscovery.
                If objFdmsregister.DatawireStatus = "OK" Then
                    cFDMSDataWireId = objFdmsregister.DatawireId
                End If

                ' After registering, you must immediately Activate the merchant. 
                ' If you wait too long, your registration will time out, and you will have to contact Datawire 
                ' to have them reset your account so that you may register again. When activating the merchant, 
                ' the same properties are required as for the Register method, with the addition of the DatawireId property. 
                ' Since the call to Register filled the DatawireId property with the correct value, all we need to do now is:
                objFdmsregister.TransactionNumber = ActivationTransactionNumber
                objFdmsregister.Activate()

                PrimaryDiscoveryURL = objFdmsregister.PrimaryDiscoveryURL
                SecondaryDiscoveryURL = objFdmsregister.SecondaryDiscoveryURL
            End If

            ' After registering and activating the account, you may now move on to Service Discovery. 
            ' The Register method will return both a PrimaryDiscoveryURL and a SecondaryDiscoveryURL. 
            ' Use these URLs with the ServiceDiscovery method to retrieve a list of valid transaction URLs. For instance:

            Dim numPrimary As Integer = 0
            ' Get the Primary Urls that can be used for transactions
            ' After registering and activating the account, you may now move on to Service Discovery. 
            ' The Register method will return both a PrimaryDiscoveryURL and a SecondaryDiscoveryURL. 
            ' Use these URLs with the ServiceDiscovery method to retrieve a list of valid transaction URLs. For instance:
            ReDim cFDMSServiceProviders(1)

            If PrimaryDiscoveryURL.Length > 0 Then
                objFdmsregister.ServiceDiscovery(PrimaryDiscoveryURL)
                If objFdmsregister.DatawireStatus = "OK" Then
                    numPrimary = objFdmsregister.ServiceProviders.Length
                    ReDim cFDMSServiceProviders(objFdmsregister.ServiceProviders.Length)
                    For iloop As Int16 = 0 To objFdmsregister.ServiceProviders.Length - 1
                        objFdmsregister.Ping(objFdmsregister.ServiceProviders(iloop))
                        cFDMSServiceProviders(iloop).Url = objFdmsregister.ServiceProviders(iloop)
                        cFDMSServiceProviders(iloop).ResponseTime = objFdmsregister.PingResponseTime
                    Next
                End If
            End If

            ' Get the Secondary Urls that can be used for transactions
            If SecondaryDiscoveryURL.Length > 0 Then
                objFdmsregister.ServiceDiscovery(SecondaryDiscoveryURL)
                If objFdmsregister.DatawireStatus = "OK" Then
                    ReDim Preserve cFDMSServiceProviders(numPrimary + objFdmsregister.ServiceProviders.Length)
                    For iloop As Int16 = 0 To objFdmsregister.ServiceProviders.Length - 1
                        objFdmsregister.Ping(objFdmsregister.ServiceProviders(numPrimary + iloop))
                        cFDMSServiceProviders(numPrimary + iloop).Url = objFdmsregister.ServiceProviders(numPrimary + iloop)
                        cFDMSServiceProviders(numPrimary + iloop).ResponseTime = objFdmsregister.PingResponseTime
                    Next
                End If
            End If

        Catch exc As nsoftware.InFDMS.InFDMSException
            clsLastError = "FDMSDatawireRegistry: Error [" & exc.Code.ToString & "]: " & exc.Message
            Throw New System.Exception("FDMSDatawireRegistry: Error [" & exc.Code.ToString & "]: " & exc.Message)

        Catch ex As Exception
            clsLastError = "FDMSDatawireRegistry: " & ex.Message
            Throw New System.Exception("FDMSDatawireRegistry: " & ex.Message)
        End Try
    End Sub

    Private Sub FDMSResponse(ByRef objFdmsecommerce As nsoftware.InFDMS.Fdmsecommerce, ByVal Text As String)
        NetworkResponse = New Response(objFdmsecommerce, Text)
    End Sub

    Private Sub FDMSSale()

        clsLastError = String.Empty

        ' First data sale is when the Authorization is settled.
        Try
            NetworkResponse = Nothing
            FDMSAuthorize("Sale")

        Catch exc As nsoftware.InFDMS.InFDMSException
            clsLastError = "FDMSSale: Error [" & exc.Code.ToString & "]: " & exc.Message
            Throw New System.Exception("FDMSSale: Error [" & exc.Code.ToString & "]: " & exc.Message)

        Catch ex As Exception
            clsLastError = "FDMSSale: " & ex.Message
            Throw New Exception("FDMSSale: " & ex.Message)
        End Try

    End Sub

    Private Sub FDMSVoidReversalTrans()

        clsLastError = String.Empty

        Try
            ' Voids or Reverse part of an Auhtorization
            NetworkResponse = Nothing

            FDMSMerchantSetup(objFdmsReversal)
            objFdmsReversal.Card.Number = CustomerCreditCard.CardNumber
            objFdmsReversal.Card.ExpMonth = CustomerCreditCard.CardExpMonth
            objFdmsReversal.Card.ExpYear = CustomerCreditCard.CardExpYear
            objFdmsReversal.Card.EntryDataSource = nsoftware.InFDMS.FDMSEntryDataSources.edsManuallyEntered

            objFdmsReversal.ApprovalCode = ChargeReversal.ApprovalCode
            objFdmsReversal.TransactionId = ChargeReversal.TransactionId
            objFdmsReversal.AuthorizedAmount = Convert.ToString(ChargeReversal.AuthorizedAmount * 100)
            objFdmsReversal.ReturnedACI = ChargeReversal.ReturnedACI
            objFdmsReversal.ValidationCode = ChargeReversal.ValidationCode

            objFdmsReversal.TransactionNumber = ChargeReversal.TransactionNumber
            objFdmsReversal.SettlementAmount = Convert.ToString(ChargeReversal.SettlementAmount * 100)
            objFdmsReversal.Reverse()

            NetworkResponse = New Response(objFdmsReversal)

        Catch exc As nsoftware.InFDMS.InFDMSException
            clsLastError = "FDMSVoidReversalTrans: Error [" & exc.Code.ToString & "]: " & exc.Message
            Throw New System.Exception("FDMSVoidReversalTrans: Error [" & exc.Code.ToString & "]: " & exc.Message)

        Catch ex As Exception
            clsLastError = "FDMSVoidReversalTrans: " & ex.Message
            Throw New Exception("FDMSVoidReversalTrans: " & ex.Message)
        End Try
    End Sub

#End Region

#Region "Other Classes"

    <Serializable()> _
    Public Class Merchant

        Private cUrl As String = String.Empty
        Private cMerchantNumber As String = String.Empty
        Private cMerchantTerminalNumber As String = String.Empty
        Private cDatawireId As String = String.Empty
        Private cVisaIdentifier As String = String.Empty
        Private cClientNumber As String = String.Empty
        Private cUserID As String = String.Empty
        Private cPassword As String = String.Empty
        Private cMerchantTaxID As String = String.Empty

        Public Sub New()
            cUrl = String.Empty
            cMerchantNumber = String.Empty
            cMerchantTerminalNumber = String.Empty
            cDatawireId = String.Empty
            cVisaIdentifier = String.Empty
            cClientNumber = String.Empty
            cUserID = String.Empty
            cPassword = String.Empty
            cMerchantTaxID = String.Empty
        End Sub

        ''' <summary>
        ''' Location of the Datawire server to which transactions are sent.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Url() As String
            Get
                Return cUrl
            End Get
            Set(ByVal value As String)
                cUrl = value
            End Set
        End Property

        ''' <summary>
        ''' A unique number used to identify the merchant within the system
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property MerchantNumber() As String
            Get
                Return cMerchantNumber
            End Get
            Set(ByVal value As String)
                cMerchantNumber = value
            End Set
        End Property

        ''' <summary>
        ''' Merchant's government Tax ID Number 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property MerchantTaxID() As String
            Get
                Return cMerchantTaxID
            End Get
            Set(ByVal value As String)
                cMerchantTaxID = value
            End Set
        End Property

        ''' <summary>
        ''' Used to identify a unique terminal within a merchant location.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property MerchantTerminalNumber() As String
            Get
                Return cMerchantTerminalNumber
            End Get
            Set(ByVal value As String)
                cMerchantTerminalNumber = value
            End Set
        End Property

        ''' <summary>
        ''' Identifies the merchant to the Datawire System (FDMS)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DatawireId() As String
            Get
                Return cDatawireId
            End Get
            Set(ByVal value As String)
                cDatawireId = value
            End Set
        End Property

        ''' <summary>
        ''' Additional merchant identification field used when authorizing Visa transactions (FDMS)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property VisaIdentifier() As String
            Get
                Return cVisaIdentifier
            End Get
            Set(ByVal value As String)
                cVisaIdentifier = value
            End Set
        End Property

        ''' <summary>
        ''' Merchant configuration property, assigned by (Paymentech)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ClientNumber() As String
            Get
                Return cClientNumber
            End Get
            Set(ByVal value As String)
                cClientNumber = value
            End Set
        End Property

        ''' <summary>
        ''' UserId for authentication with the NetConnect Server (Paymentech)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property UserID() As String
            Get
                Return cUserID
            End Get
            Set(ByVal value As String)
                cUserID = value
            End Set
        End Property

        ''' <summary>
        ''' Password for authentication with the NetConnect Server (Paymentech)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Password() As String
            Get
                Return cPassword
            End Get
            Set(ByVal value As String)
                cPassword = value
            End Set
        End Property

    End Class

    <Serializable()> _
    Public Class CreditCard

        Private cCardNumber As String = String.Empty
        Private cExpMonth As String = String.Empty
        Private cExpYear As String = String.Empty
        Private cCVV As String = String.Empty
        Private cCardType As CreditCardTypes = CreditCardTypes.vctUnknown

        Private cCardHolderFirstName As String = String.Empty
        Private cCardHolderLastName As String = String.Empty
        Private cCardHolderAddress As String = String.Empty
        Private cCardHolderCity As String = String.Empty
        Private cCardHolderState As String = String.Empty
        Private cCardHolderZipCode As String = String.Empty
        Private cCardHolderCountry As String = String.Empty
        Private cCardHolderEmail As String = String.Empty
        Private cCardHolderTelephone As String = String.Empty

        Public Sub New()
            cCardNumber = String.Empty
            cExpMonth = String.Empty
            cExpYear = String.Empty
            cCVV = String.Empty

            cCardHolderFirstName = String.Empty
            cCardHolderLastName = String.Empty
            cCardHolderAddress = String.Empty
            cCardHolderCity = String.Empty
            cCardHolderState = String.Empty
            cCardHolderZipCode = String.Empty
            cCardHolderCountry = String.Empty
            cCardHolderEmail = String.Empty
        End Sub

        ''' <summary>
        ''' The Customer's credit card number
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property CardNumber() As String
            Get
                Return cCardNumber
            End Get
            Set(ByVal value As String)
                cCardNumber = value
            End Set
        End Property

        ''' <summary>
        ''' Expiration month of the credit card specified by CreditCard
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property CardExpMonth() As String
            Get
                Return cExpMonth
            End Get
            Set(ByVal value As String)
                cExpMonth = Val(value).ToString("00")
            End Set
        End Property

        ''' <summary>
        ''' Expiration Year of the credit card specified by CreditCard
        ''' This field contains the expiration date of the customer's credit card. This field must be in the range 0 - 99, or 2000 - 2099. 
        ''' Any date before the year 2000 cannot be specified.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property CardExpYear() As String
            Get
                Return cExpYear
            End Get
            Set(ByVal value As String)
                Select Case Val(value)
                    Case 0 To 99
                        cExpYear = (Val(DateTime.Now.ToString("yyyy")) \ 100).ToString & Val(value).ToString("00")
                    Case 2000 To 2099
                        cExpYear = Val(value).ToString
                    Case Else
                        ' Nothing
                End Select
            End Set
        End Property

        ''' <summary>
        ''' Three digit security code on back of card (optional). 
        ''' This alphanumeric field contains the three digit Visa "Card Verification Value" (CVV), 
        ''' MasterCard "Card Verification Code" (CVC), 
        ''' or four-digit American Express "Card Identification Number" (CID)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property CardCVVData() As String
            Get
                Return cCVV
            End Get
            Set(ByVal value As String)
                cCVV = value
            End Set
        End Property

        ''' <summary>
        ''' The customer's billing Street address.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property CardHolderAddress() As String
            Get
                Return cCardHolderAddress
            End Get
            Set(ByVal value As String)
                cCardHolderAddress = value
            End Set
        End Property

        ''' <summary>
        ''' Customer's Billing City
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property CardHolderCity As String
            Get
                Return cCardHolderCity
            End Get
            Set(value As String)
                cCardHolderCity = value.Trim
            End Set
        End Property

        ''' <summary>
        ''' Customer's zip code (or postal code if outside of the USA).
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property CardHolderZipCode() As String
            Get
                Return cCardHolderZipCode
            End Get
            Set(ByVal value As String)
                cCardHolderZipCode = value
            End Set
        End Property

        ''' <summary>
        ''' First name on the card
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property CardHolderFirstName() As String
            Get
                Return cCardHolderFirstName
            End Get
            Set(ByVal value As String)
                cCardHolderFirstName = value
            End Set
        End Property

        ''' <summary>
        ''' Last name on the card
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property CardHolderLastName() As String
            Get
                Return cCardHolderLastName
            End Get
            Set(ByVal value As String)
                cCardHolderLastName = value
            End Set
        End Property

        ''' <summary>
        ''' State code for the address of the card
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property CardHolderState() As String
            Get
                Return cCardHolderState
            End Get
            Set(ByVal value As String)
                cCardHolderState = value
            End Set
        End Property

        ''' <summary>
        ''' Set the country of the Card's address
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property CardHolderCountry() As String
            Get
                Return cCardHolderCountry
            End Get
            Set(value As String)
                cCardHolderCountry = value.Trim
            End Set
        End Property

        ''' <summary>
        ''' Email address of card holder
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property CardHolderEmail As String
            Get
                Return cCardHolderEmail
            End Get
            Set(value As String)
                cCardHolderEmail = value.Trim
            End Set
        End Property

        ''' <summary>
        ''' Telephone number of card holder
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property CardHolderTelephone() As String
            Get
                Return cCardHolderTelephone
            End Get
            Set(value As String)
                cCardHolderTelephone = value.Trim
            End Set
        End Property

        ''' <summary>
        ''' Credit Card Type
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property CardType As CreditCardTypes
            Get
                Return cCardType
            End Get
            Set(value As CreditCardTypes)
                cCardType = value
            End Set
        End Property

    End Class

    <Serializable()> _
    Public Class Level2

        'American Express: 

        'DestinationZip (required) *
        'PurchaseIdentifier (required) *
        'TaxAmount (required) *

        'Visa 

        'PurchaseIdentifier (required) *
        'CommercialCardType (optional)
        'InvoiceNumber (optional) *
        'DestinationCountry (optional)
        'DestinationZip (optional) *
        'DiscountAmount (optional) *
        'DutyAmount (optional)
        'FreightAmount (optional) *
        'FreightTaxAmount (optional)
        'FreightTaxRate (optional)
        'MerchantReference (optional)
        'OrderDate (optional)
        'ShipFromZip (optional) *
        'TaxAmount (optional) *

        'MasterCard: 

        'MerchantTaxId (required) *
        'MerchantType (required)
        'PurchaseIdentifier (required) *
        'TaxAmount (required) *
        'CommercialCardType (optional)
        'DestinationCountry (optional)
        'DestinationState (optional)
        'DestinationZip (optional) *
        'DutyAmount (optional)
        'FreightAmount (optional) *
        'MerchantReference (optional)
        'ShipFromZip (optional) *

        Private cCardType As TAC.ARCCCARD.CreditCardTypes
        Private cCommercialCardType As String = String.Empty
        Private cDestinationZip As String = String.Empty
        Private cDiscountAmount As Double = 0
        Private cFreightAmount As Double = 0
        Private cInvoiceNumber As String = String.Empty
        Private cMerchantTaxId As String = String.Empty
        Private cMerchantType As String = "0108"
        Private cOrderDate As String = String.Empty
        Private cPurchaseIdentifier As String = String.Empty
        Private cShipFromZip As String = String.Empty
        Private cTaxAmount As Double = 0
        Private cDestinationState As String = String.Empty

        ' Merchant Types
        ' Position 1 - Business Classification 
        '   0 Unknown 
        '   1 Corporation 
        '   2 Other Business Structure (e.g., Partnership, Sole Proprietorship, Joint Venture, etc.) 

        'Position 2 - Type of Business Owner Classification 
        '   0 Unknown 
        '   1 No Applicable Classification (not female or handicapped) 
        '   2 Female Business Owner 
        '   3 Physically Handicapped Female Business Owner 
        '   4 Physically Handicapped Business Owner 

        'Position 3 - Minority, Small Business, Disadvantaged Certification 
        '   0 Unknown 
        '   1 Certification Not Applicable 
        '   2 SBA Certification as Small Business 
        '   3 SBA Certification as Small Disadvantaged Business 
        '   4 Other Government or Agency Recognized Certification (e.g., Minority Supplier Development Council) 
        '   5 Self-certified Small Business 
        '   6 Definition 2 and Definition 4 
        '   7 Definition 3 and Definition 4 
        '   8 Definition 4 and Definition 5 

        'Position 4 - Racial or Ethnic Type (Unconditional Majority Owner) 
        '   0 Unknown 
        '   1 African American 
        '   2 Asian Pacific American 
        '   3 Subcontinent Asian American 
        '   4 Hispanic American 
        '   5 Native American Indian 
        '   6 Native Hawaiian 
        '   7 Native Alaskan 
        '   8 Caucasian 
        '   9 Other 

        Public Sub New()
            Clear()
        End Sub

        ' Initializes control, clears previous values
        Public Sub Clear()
            cTaxAmount = 0
            cPurchaseIdentifier = String.Empty
            cDestinationZip = String.Empty
            cInvoiceNumber = String.Empty
            cShipFromZip = String.Empty
            cDiscountAmount = 0
            cFreightAmount = 0
            cCommercialCardType = String.Empty
            cMerchantTaxId = String.Empty
            cMerchantType = "0108"
            cCardType = CreditCardTypes.vctUnknown
            cOrderDate = Nothing
            cDestinationState = String.Empty
        End Sub

        ''' <summary>
        ''' Credit Card Type
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property CardType() As TAC.ARCCCARD.CreditCardTypes
            Get
                Return cCardType
            End Get
            Set(ByVal value As TAC.ARCCCARD.CreditCardTypes)
                cCardType = value
            End Set
        End Property

        ''' <summary>
        ''' Discount Amount
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DiscountAmount() As Double
            Get
                Return cDiscountAmount
            End Get
            Set(ByVal value As Double)
                cDiscountAmount = value
            End Set
        End Property

        ''' <summary>
        ''' State where goods are getting shipped
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DestinationState() As String
            Get
                Return cDestinationState
            End Get
            Set(ByVal value As String)
                cDestinationState = value
            End Set
        End Property


        ''' <summary>
        ''' Freight Amount
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property FreightAmount() As Double
            Get
                Return cFreightAmount
            End Get
            Set(ByVal value As Double)
                cFreightAmount = value
            End Set
        End Property

        ''' <summary>
        ''' Invoice Number
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property InvoiceNumber() As String
            Get
                Return cInvoiceNumber
            End Get
            Set(ByVal value As String)
                cInvoiceNumber = value
            End Set
        End Property

        ''' <summary>
        ''' This property contains the government-assigned tax identification number of the merchant from whom the goods or services were purchased. 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property MerchantTaxId() As String
            Get
                Return cMerchantTaxId
            End Get
            Set(ByVal value As String)
                cMerchantTaxId = value
            End Set
        End Property

        ''' <summary>
        ''' Merchant Type is a four-position field defined by MasterCard that describes various merchant classifications. 
        ''' Each position in the field has a separate, special meaning for categorizing the merchant. 
        ''' If the data for a particular position is not available, that position must contain a blank. 
        ''' If it is available, it must contain one of the following business codes:
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property MerchantType() As String
            Get
                Return cMerchantType
            End Get
            Set(ByVal value As String)
                cMerchantType = value
            End Set
        End Property

        ''' <summary>
        ''' This property contains the date the order was taken
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property OrderDate() As Date
            Get
                If cOrderDate Is Nothing Then
                    Return Nothing
                Else
                    Return cOrderDate
                End If
            End Get
            Set(ByVal value As Date)
                cOrderDate = value
            End Set
        End Property

        ''' <summary>
        ''' This property contains a Purchase Identifier (Purchase Order Number or Order Number). 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property PurchaseIdentifier() As String
            Get
                Return cPurchaseIdentifier
            End Get
            Set(ByVal value As String)
                cPurchaseIdentifier = value
            End Set
        End Property

        ''' <summary>
        ''' This property contains the portion of the TransactionAmount which is sales tax
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property TaxAmount() As Double
            Get
                Return cTaxAmount
            End Get
            Set(ByVal value As Double)
                cTaxAmount = value
            End Set
        End Property

        ''' <summary>
        ''' This property contains the zip code where the purchased goods are to be shipped from.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ShipFromZip() As String
            Get
                Return cShipFromZip
            End Get
            Set(ByVal value As String)
                cShipFromZip = value
            End Set
        End Property

        ''' <summary>
        ''' This property contains the zip code where the purchased goods are to be shipped.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DestinationZip() As String
            Get
                Return cDestinationZip
            End Get
            Set(ByVal value As String)
                cDestinationZip = value
            End Set
        End Property

    End Class

    <Serializable()> _
    Public Class Level3

        'List of valid Line Item properties for Visa: 

        'Quantity (required) *
        'UnitCost (required) *
        'CommodityCode (optional)
        'Description (optional) *
        'DiscountAmount (optional) *
        'ProductCode (optional) *
        'TaxAmount (optional) *
        'TaxRate (optional) *
        'Total (optional) *
        'Units (optional) *

        'List of valid Line Item properties for Mastercard: 

        'Description (required) *
        'ProductCode (required) *
        'Quantity (required) *
        'Units (required) *
        'UnitCost (optional) *
        'Total (required) *
        'DiscountAmount (optional) *
        'TaxIncluded (required)
        'TaxRate (required) *
        'TaxAmount (required) *
        'TaxType (required)

        Private cQuantity As Integer = 0
        Private cUnitCost As Double = 0
        Private cDescription As String = String.Empty
        Private cDiscountAmount As Double = 0
        Private cProductCode As String = String.Empty
        Private cTaxAmount As Double = 0
        Private cTaxRate As Double = 0
        Private CTotal As Double = 0
        Private cUnits As String = "each"
        Private cTaxType As TaxTypes = TaxTypes.StateSalesTax

        Public Sub New()
            cQuantity = 0
            cUnitCost = 0
            cDescription = String.Empty
            cDiscountAmount = 0
            cProductCode = String.Empty
            cTaxAmount = 0
            cTaxRate = 0
            CTotal = 0
            cUnits = "each"
            cTaxType = TaxTypes.StateSalesTax
        End Sub

        ''' <summary>
        ''' This field contains an alphanumeric description of the item(s) being supplied. 
        ''' The maximum length of this field is 26 characters.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Description() As String
            Get
                Return cDescription
            End Get
            Set(ByVal value As String)
                cDescription = value
            End Set
        End Property

        ''' <summary>
        ''' Amount of the discount for each line item (if any). 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DiscountAmount() As Double
            Get
                Return cDiscountAmount
            End Get
            Set(ByVal value As Double)
                cDiscountAmount = value
            End Set
        End Property

        ''' <summary>
        ''' This field contains a code assigned to the product by the merchant. 
        ''' This may be a UPC or any other code with which the merchant wishes to identify an individual product. 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ProductCode() As String
            Get
                Return cProductCode
            End Get
            Set(ByVal value As String)
                cProductCode = value
            End Set
        End Property

        ''' <summary>
        ''' This field contains the quantity of items being purchased, in whole numbers. 
        ''' The maximum quantity is 99,999 items
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Quantity() As Integer
            Get
                Return cQuantity
            End Get
            Set(ByVal value As Integer)
                cQuantity = value
                If cQuantity > 99999 Then
                    cQuantity = 99999
                End If
            End Set
        End Property

        ''' <summary>
        ''' This field contains the amount of any Value Added Taxes (VAT) which can be associated with the purchased item.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property TaxAmount() As Double
            Get
                Return cTaxAmount
            End Get
            Set(ByVal value As Double)
                cTaxAmount = value
            End Set
        End Property

        ''' <summary>
        ''' Tax rate used to calculate the TaxAmount. 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property TaxRate() As Double
            Get
                Return cTaxRate
            End Get
            Set(ByVal value As Double)
                cTaxRate = value
            End Set
        End Property

        ''' <summary>
        ''' This field designates the type of value-added taxes (VAT) that are being charged in TaxAmount 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property TaxType() As TaxTypes
            Get
                Return cTaxType
            End Get
            Set(ByVal value As TaxTypes)
                cTaxType = value
            End Set
        End Property

        ''' <summary>
        ''' Total cost of this line item.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Total() As Double
            Get
                Return CTotal
            End Get
            Set(ByVal value As Double)
                CTotal = value
            End Set
        End Property

        ''' <summary>
        ''' Cost of each individual item. 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property UnitCost() As Double
            Get
                Return cUnitCost
            End Get
            Set(ByVal value As Double)
                cUnitCost = value
            End Set
        End Property

        ''' <summary>
        ''' Unit of measure for this Line Item.
        ''' Examples: each, feet, ounce, Defaults to each
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Units() As String
            Get
                Return cUnits
            End Get
            Set(ByVal value As String)
                cUnits = value
            End Set
        End Property

    End Class

    <Serializable()> _
    Public Class Response
        Private cApprovalCode As String = String.Empty 'P/F Contains an authorization code for an approved transaction.  CaptureFlag true/false from FDMS
        Private cAuthSource As String = String.Empty 'P Indicates the source of the authorization code stored in ApprovalCode. 
        Private cAVSResult As String = String.Empty 'P/F Contains the Address Verification System result code. Used for fraud detection. 
        Private cBatchNumber As String = String.Empty 'P Number that identifies the batch that this transaction will be settled in (only applicable to Host Capture. Will be all zeros for Terminal Capture settlements). 
        Private cCode As String = String.Empty 'P/F Indicates the success or failure of the transaction. 
        Private cCVVResult As String = String.Empty 'P/F Contains the returned CVV result code if it was requested. Used for fraud detection. 
        Private cData As String = String.Empty 'P Contains the raw response from the Paymentech host. 
        Private cDebitSurcharge As Double = 0  'P  Additional fee (if any) charged for debit transactions. 
        Private cDebitTrace As String = String.Empty 'P Debit-specific host tracking number for this transaction.  
        Private cRetrievalNumber As String = String.Empty 'P Reference number returned from the Paymentech host. 
        Private cSequenceNumber As String = String.Empty 'P SequenceNumber echoed from the authorization. 
        Private cText As String = String.Empty 'P/F Contains a human-readable message explaining the code.  ValidationCode from FDMS

        Private cCaptureFlag As Boolean = False 'F Indicates whether the authorization was successful, and whether it can be settled.
        Private cCardLevelResult As String = String.Empty 'F  Two character card level results field returned in the response to Visa authorizations on prepaid cards. 
        Private cDatawireReturnCode As String = String.Empty 'F Contains an error code providing more details about the DatawireStatus received. 
        Private cDatawireStatus As String = String.Empty 'F  Status of the communication with Datawire. 
        Private cReturnedACI As String = String.Empty 'F Returned Authorization Characteristics Indicator contains CPS qualification status. 
        Private cTransactionDate As String = String.Empty 'F Local transaction date returned from the server in MMDDYY format. 
        Private cTransactionId As String = String.Empty 'F Contains the Visa Transaction Identifier or MasterCard Reference Number. 
        Private cValidationCode As String = String.Empty 'F Contains the Visa Transaction Identifier or MasterCard Reference Number. 

        Private cDetailAggregate As String = String.Empty 'F aggregate containing details of this transaction, which is then used for settlement
        Private cCommercialCardType As nsoftware.InFDMS.FDMSCommercialCards = nsoftware.InFDMS.FDMSCommercialCards.cctUnknown

        Public Sub New()
            InitializeClass()
        End Sub

        Public Sub New(response As nsoftware.InPay.EPResponse)
            InitializeClass()

            With response
                cApprovalCode = .ApprovalCode
                cCode = .Code
                cData = .Data
                cCaptureFlag = .Approved
                cTransactionId = .TransactionId
                cTransactionDate = DateTime.Now

                cAuthSource = String.Empty
                cAVSResult = .AVSResult
                cBatchNumber = String.Empty
                cCVVResult = String.Empty
                'cDebitSurcharge = String.Empty
                cDebitTrace = String.Empty
                cRetrievalNumber = String.Empty
                cSequenceNumber = String.Empty
                cText = .Text

                cCardLevelResult = String.Empty
                cDatawireReturnCode = String.Empty
                cDatawireStatus = String.Empty
                cReturnedACI = String.Empty
                cValidationCode = String.Empty

            End With
        End Sub

        Public Sub New(ByRef objPtcharge As nsoftware.IBizPtech.Ptcharge)

            InitializeClass()

            With objPtcharge
                cApprovalCode = .ResponseApprovalCode
                cAuthSource = .ResponseAuthSource
                cAVSResult = .ResponseAVS
                cBatchNumber = .ResponseBatchNumber
                cCode = .ResponseCode
                cCVVResult = .ResponseCVVResult
                cData = .ResponseData
                cDebitSurcharge = .ResponseDebitSurcharge
                cDebitTrace = .ResponseDebitTrace
                cRetrievalNumber = .ResponseRetrievalNumber
                cSequenceNumber = .ResponseSequenceNumber
                cText = .ResponseText

                cCaptureFlag = .ResponseApprovalCode = "A"
                cCardLevelResult = String.Empty
                cDatawireReturnCode = String.Empty
                cDatawireStatus = String.Empty
                cReturnedACI = String.Empty
                cTransactionDate = String.Empty
                cTransactionId = String.Empty
                cValidationCode = String.Empty

                cCommercialCardType = nsoftware.InFDMS.FDMSCommercialCards.cctUnknown

            End With
        End Sub

        Public Sub New(ByVal ApprovalCode As String, ByVal ResponseCode As String)

            InitializeClass()

            cApprovalCode = ApprovalCode

            cCode = ResponseCode
            cTransactionDate = DateTime.Now.ToString("MMddyy")

            cDatawireStatus = "OK"
            cCVVResult = "M"
            cAVSResult = "N"
            cText = "Capture"

            If ResponseCode = Approved Then
                cCaptureFlag = True
            Else
                cCaptureFlag = False
            End If

        End Sub

        Public Sub New(ByVal ApprovalCode As String, ByVal ResponseCode As String, ByVal CreditCard As CreditCard, ByVal CreditAmount As Double)

            InitializeClass()

            cApprovalCode = ApprovalCode

            cCode = ResponseCode
            cTransactionDate = DateTime.Now.ToString("MMddyy")

            cDatawireStatus = "OK"
            cCVVResult = "M"
            cAVSResult = "N"
            cReturnedACI = "Z" ' Needed, value in this field identifes transaction as FDMS
            cText = "Credit"

            ' Create the settlement detail aggregate
            Dim FDMSDetailRecord As New nsoftware.InFDMS.Fdmsdetailrecord()
            FDMSDetailRecord.RuntimeLicense = fdmsRuntimeLicense
            With FDMSDetailRecord
                .TransactionType = nsoftware.InFDMS.FdmsdetailrecordTransactionTypes.ttCredit
                .IndustryType = nsoftware.InFDMS.FdmsdetailrecordIndustryTypes.itDirectMarketing
                .AccountDataSource = nsoftware.InFDMS.FdmsdetailrecordAccountDataSources.dsManuallyKeyed
                .DirectMarketingType = nsoftware.InFDMS.FdmsdetailrecordDirectMarketingTypes.dmECommerce

                .CardNumber = CreditCard.CardNumber
                .CardExpMonth = CreditCard.CardExpMonth
                .CardExpYear = CreditCard.CardExpYear
                .AuthorizedAmount = 0
                .SettlementAmount = Val(Math.Abs(CreditAmount) * 100).ToString
                .TransactionTime = DateTime.Now.ToString("hhmmss")
                .TransactionDate = DateTime.Now.ToString("MMddyy")
                .PurchaseIdentifier = ApprovalCode

                cDetailAggregate = .GetDetailAggregate
            End With

            If ResponseCode = Approved Then
                cCaptureFlag = True
            Else
                cCaptureFlag = False
            End If

        End Sub

        Public Sub New(ByRef objFdmsReversal As nsoftware.InFDMS.Fdmsreversal)
            InitializeClass()

            With objFdmsReversal.Response
                cApprovalCode = .ApprovalCode

                cAVSResult = objFdmsReversal.Response.AVSResult
                cCaptureFlag = .CaptureFlag
                If cCaptureFlag = True Then
                    cCode = Approved
                Else
                    cCode = Declined
                End If

                cCardLevelResult = .CardLevelResult
                cCVVResult = .CVVResult
                cDatawireReturnCode = .DatawireReturnCode
                cDatawireStatus = .DatawireStatus
                cReturnedACI = .ReturnedACI
                cTransactionDate = .TransactionDate
                cTransactionId = .TransactionId
                cValidationCode = .ValidationCode
                cText = "Capture/Reversal"
            End With
        End Sub

        Public Sub New(ByRef objFdmsecommerce As nsoftware.InFDMS.Fdmsecommerce, ByVal text As String)

            InitializeClass()

            Dim ResponseCode As String = String.Empty

            If objFdmsecommerce.Response.CaptureFlag = True Then
                ResponseCode = Approved
                cText = "Approved"
            ElseIf objFdmsecommerce.Response.DatawireReturnCode <> "000" Then
                ResponseCode = DatawireError
                cText = "Error Declined"
            Else
                ResponseCode = Declined
                cText = "Declined"
            End If

            With objFdmsecommerce
                cApprovalCode = .Response.ApprovalCode.Trim
                cAuthSource = String.Empty
                cAVSResult = .Response.AVSResult
                cBatchNumber = String.Empty
                cCode = ResponseCode
                cCVVResult = .Response.CVVResult
                cData = String.Empty
                cDebitSurcharge = 0
                cDebitTrace = String.Empty
                cRetrievalNumber = String.Empty
                cSequenceNumber = String.Empty
                'cText = text

                If ResponseCode <> Approved Then
                    cApprovalCode = String.Empty
                    cText = .Response.ApprovalCode.Trim
                End If

                cCaptureFlag = .Response.CaptureFlag
                cCardLevelResult = .Response.CardLevelResult
                cDatawireReturnCode = .Response.DatawireReturnCode
                cDatawireStatus = .Response.DatawireStatus
                cReturnedACI = .Response.ReturnedACI
                cTransactionDate = .Response.TransactionDate
                cTransactionId = .Response.TransactionId
                cValidationCode = .Response.ValidationCode
                cCommercialCardType = .Response.CommercialCard

                ' If the capture is false the a call to .GetDetailAggregate throws an error
                ' Capture was false but I was able to get the DetailAggregate
                Try
                    cDetailAggregate = .GetDetailAggregate()
                Catch ex As Exception
                    cDetailAggregate = String.Empty
                End Try

            End With
        End Sub

        Private Sub InitializeClass()

            cApprovalCode = String.Empty 'P/F Contains an authorization code for an approved transaction.  CaptureFlag true/false from FDMS
            cAuthSource = String.Empty 'P Indicates the source of the authorization code stored in ApprovalCode. 
            cAVSResult = String.Empty 'P/F Contains the Address Verification System result code. Used for fraud detection. 
            cBatchNumber = String.Empty 'P Number that identifies the batch that this transaction will be settled in (only applicable to Host Capture. Will be all zeros for Terminal Capture settlements). 
            cCode = String.Empty 'P/F Indicates the success or failure of the transaction. 
            cCVVResult = String.Empty 'P/F Contains the returned CVV result code if it was requested. Used for fraud detection. 
            cData = String.Empty 'P Contains the raw response from the Paymentech host. 
            cDebitSurcharge = 0  'P  Additional fee (if any) charged for debit transactions. 
            cDebitTrace = String.Empty 'P Debit-specific host tracking number for this transaction.  
            cRetrievalNumber = String.Empty 'P Reference number returned from the Paymentech host. 
            cSequenceNumber = String.Empty 'P SequenceNumber echoed from the authorization. 
            cText = String.Empty 'P/F Contains a human-readable message explaining the code.  ValidationCode from FDMS

            cCaptureFlag = False 'F Indicates whether the authorization was successful, and whether it can be settled.
            cCardLevelResult = String.Empty 'F  Two character card level results field returned in the response to Visa authorizations on prepaid cards. 
            cDatawireReturnCode = String.Empty 'F Contains an error code providing more details about the DatawireStatus received. 
            cDatawireStatus = String.Empty 'F  Status of the communication with Datawire. 
            cReturnedACI = String.Empty 'F Returned Authorization Characteristics Indicator contains CPS qualification status. 
            cTransactionDate = String.Empty 'F Local transaction date returned from the server in MMDDYY format. 
            cTransactionId = String.Empty 'F Contains the Visa Transaction Identifier or MasterCard Reference Number. 
            cValidationCode = String.Empty 'F Contains the Visa Transaction Identifier or MasterCard Reference Number. 
            cDetailAggregate = String.Empty 'F aggregate containing details of this transaction, which is then used for settlement
            cCommercialCardType = nsoftware.InFDMS.FDMSCommercialCards.cctUnknown 'F Indicatse Type of Commercial card Being Used
        End Sub

        ''' <summary>
        ''' Contains an authorization code for an approved transaction. 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ApprovalCode() As String
            Get
                Return (cApprovalCode)
            End Get
            Set(ByVal value As String)
                cApprovalCode = value
            End Set
        End Property

        ''' <summary>
        ''' Indicates the source of the authorization code stored in ApprovalCode.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property AuthSource() As String
            Get
                Return (cAuthSource)
            End Get
            Set(ByVal value As String)
                cAuthSource = value
            End Set

        End Property

        ''' <summary>
        ''' Contains the Address Verification System result code. Used for fraud detection.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property AVSResult() As String
            Get
                Return (cAVSResult)
            End Get
            Set(ByVal value As String)
                cAVSResult = value
            End Set

        End Property

        ''' <summary>
        ''' Number that identifies the batch that this transaction will be settled in (only applicable to Host Capture. Will be all zeros for Terminal Capture settlements). 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property BatchNumber() As String
            Get
                Return (cBatchNumber)
            End Get

            Set(ByVal value As String)
                cBatchNumber = value
            End Set

        End Property

        ''' <summary>
        ''' Two character card level results field returned in the response to Visa authorizations on prepaid cards.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property CardLevelResult() As String
            Get
                Return cCardLevelResult
            End Get

            Set(ByVal value As String)
                cCardLevelResult = value
            End Set

        End Property

        ''' <summary>
        ''' Indictaes the type of commercial card being settled
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property CommercialCardType() As nsoftware.InFDMS.FDMSCommercialCards
            Get
                Return cCommercialCardType
            End Get

            Set(ByVal value As nsoftware.InFDMS.FDMSCommercialCards)
                cCommercialCardType = value
            End Set

        End Property

        ''' <summary>
        ''' Contains the returned CVV result code if it was requested. Used for fraud detection. 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property CVVResult() As String
            Get
                Return (cCVVResult)
            End Get

            Set(ByVal value As String)
                cCVVResult = value
            End Set

        End Property

        ''' <summary>
        ''' Contains the raw response from the Paymentech host. 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Data() As String
            Get
                Return (cData)
            End Get
            Set(ByVal value As String)
                cData = value
            End Set

        End Property

        ''' <summary>
        ''' Contains an error code providing more details about the DatawireStatus received.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DataWireReturnedCode() As String
            Get
                Return cDatawireReturnCode
            End Get
            Set(ByVal value As String)
                cDatawireReturnCode = value
            End Set

        End Property

        ''' <summary>
        ''' Status of the communication with Datawire.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DataWireStatus() As String
            Get
                Return cDatawireStatus
            End Get
            Set(ByVal value As String)
                cDatawireStatus = value
            End Set

        End Property

        ''' <summary>
        '''  Additional fee (if any) charged for debit transactions. 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DebitSurcharge() As String
            Get
                Return (cDebitSurcharge)
            End Get

            Set(ByVal value As String)
                cDebitSurcharge = value
            End Set

        End Property

        ''' <summary>
        '''  Debit-specific host tracking number for this transaction.  
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DebitTrace() As String
            Get
                Return (cDebitTrace)
            End Get

            Set(ByVal value As String)
                cDebitTrace = value
            End Set

        End Property

        ''' <summary>
        ''' Returns an aggregate containing details of this transaction, which is then used for settlement.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DetailAggregate() As String
            Get
                Return cDetailAggregate
            End Get

            Set(ByVal value As String)
                cDetailAggregate = value
            End Set

        End Property

        ''' <summary>
        ''' Indicates the success or failure of the transaction.  
        ''' A = Approved
        ''' E = Error
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ResponseCode() As String
            Get
                Return (cCode)
            End Get
            Set(ByVal value As String)
                cCode = value
            End Set

        End Property

        ''' <summary>
        '''  Reference number returned from the Paymentech host. 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property RetrievalNumber() As String
            Get
                Return (cRetrievalNumber)
            End Get

            Set(ByVal value As String)
                cRetrievalNumber = value
            End Set

        End Property

        ''' <summary>
        ''' Returned ACI from the original response.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ReturnedACI() As String
            Get
                Return cReturnedACI
            End Get

            Set(ByVal value As String)
                cRetrievalNumber = value
            End Set

        End Property

        ''' <summary>
        '''  SequenceNumber echoed from the authorization. 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property SequenceNumber() As String
            Get
                Return (cSequenceNumber)
            End Get

            Set(ByVal value As String)
                cSequenceNumber = value
            End Set

        End Property

        ''' <summary>
        '''  Contains a human-readable message explaining the code.  
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Text() As String
            Get
                Return (cText)
            End Get

            Set(ByVal value As String)
                cText = value
            End Set

        End Property

        ''' <summary>
        ''' Local transaction date returned from the server in MMDDYY format.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property TransactionDate() As String
            Get
                Return cTransactionDate
            End Get
            Set(ByVal value As String)
                cTransactionDate = value
            End Set

        End Property

        ''' <summary>
        ''' Contains the Visa Transaction Identifier or MasterCard Reference Number. 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property TransactionId() As String
            Get
                Return cTransactionId
            End Get
            Set(ByVal value As String)
                cTransactionId = value
            End Set

        End Property

        ''' <summary>
        ''' Additional information generated by the card issuer. Needed for Credits
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ValidationCode() As String
            Get
                Return cValidationCode
            End Get
            Set(ByVal value As String)
                cValidationCode = value
            End Set

        End Property

    End Class

    <Serializable()> _
    Public Class Reversal
        Private cApprovalCode As String = String.Empty
        Private cTransactionId As String = String.Empty
        Private cAuthorizedAmount As Double = 0
        Private cReturnedACI As String = String.Empty
        Private cValidationCode As String = String.Empty
        Private cTransactionNumber As String = String.Empty
        Private cSettlementAmount As Double = 0
        Private cTransactionAmount As Double = 0

        Public CreditCard As New CreditCard

        Public Sub New()
            cApprovalCode = String.Empty
            cTransactionId = String.Empty
            cAuthorizedAmount = 0
            cReturnedACI = String.Empty
            cValidationCode = String.Empty
            cTransactionNumber = String.Empty
            cSettlementAmount = 0
            cTransactionAmount = 0
            CreditCard = New CreditCard
        End Sub

        ''' <summary>
        ''' Approval code of the transaction to be reversed.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ApprovalCode() As String
            Get
                Return cApprovalCode
            End Get
            Set(ByVal value As String)
                cApprovalCode = value
            End Set
        End Property

        ''' <summary>
        ''' Transaction Id from the original response.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property TransactionId() As String
            Get
                Return cTransactionId
            End Get
            Set(ByVal value As String)
                cTransactionId = value
            End Set
        End Property

        ''' <summary>
        ''' Authorized Amount from the original response.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property AuthorizedAmount() As Double
            Get
                Return cAuthorizedAmount
            End Get
            Set(ByVal value As Double)
                cAuthorizedAmount = value
            End Set
        End Property

        ''' <summary>
        ''' Returned ACI from the original response.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ReturnedACI() As String
            Get
                Return cReturnedACI
            End Get
            Set(ByVal value As String)
                cReturnedACI = value
            End Set
        End Property

        ''' <summary>
        ''' Validation Code from the original response.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ValidationCode() As String
            Get
                Return cValidationCode
            End Get
            Set(ByVal value As String)
                cValidationCode = value
            End Set
        End Property

        ''' <summary>
        ''' Uniquely identifies the transaction.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property TransactionNumber() As String
            Get
                Return cTransactionNumber
            End Get
            Set(ByVal value As String)
                cTransactionNumber = value
            End Set
        End Property

        ''' <summary>
        ''' New settlement amount after the reversal.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property SettlementAmount() As Double
            Get
                Return cSettlementAmount
            End Get
            Set(ByVal value As Double)
                cSettlementAmount = value
            End Set
        End Property

        ''' <summary>
        ''' Transaction Amount
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property TransactionAmount() As Double
            Get
                Return cTransactionAmount
            End Get
            Set(ByVal value As Double)
                cTransactionAmount = value
            End Set
        End Property

    End Class

    <Serializable()> _
    Public Class Settle
        Private cPaymentTypeCount As Long = 0
        Private cPaymentTypeName As String()
        Private cPaymentTypeNetAmount As Decimal()
        Private cPaymentTypeTransCount As Integer()
        Private cResponseBatchClose As Date
        Private cResponseBatchNetAmount As Double = 0
        Private cResponseBatchNumber As String = String.Empty
        Private cResponseBatchOpen As String = String.Empty
        Private cResponseBatchTransCount As Integer = 0
        Private cResponseCode As String = String.Empty
        Private cResponseData As String = String.Empty
        Private cResponseInquiryCount As Integer = 0
        Private cResponseSequenceNumber As String = String.Empty
        Private cResponseText As String = String.Empty
        Private cSequenceNumber As Integer = 0

        Public Structure FDMSSettlementDelegates
            Dim Detail As String
            Dim Level2 As String
            Dim level3 As String
            Dim SettlementAmount As Double
        End Structure

        Public FdmsSettlementDetailRecords As New List(Of FDMSSettlementDelegates)

        Public Structure PaymentTypes
            Dim TypeName As String
            Dim TypeNetAmount As Double
            Dim TypeTransCount As String
        End Structure

        Public PaymentTypesList As New List(Of PaymentTypes)

        Public Sub New()
            InitializeClass()
        End Sub

        Public Sub New(ByVal objPtsettle As nsoftware.IBizPtech.Ptsettle)

            InitializeClass()

            With objPtsettle
                cPaymentTypeCount = .PaymentTypeCount

                If cPaymentTypeCount = 0 Then
                    cPaymentTypeName = Nothing
                Else
                    ReDim cPaymentTypeName(cPaymentTypeCount)
                    ReDim cPaymentTypeNetAmount(cPaymentTypeCount)
                    ReDim cPaymentTypeTransCount(cPaymentTypeCount)

                    For i As Integer = 0 To .PaymentTypeCount - 1
                        cPaymentTypeName(i) = .PaymentTypeName(i).Trim
                        cPaymentTypeNetAmount(i) = .PaymentTypeNetAmount(i)
                        cPaymentTypeTransCount(i) = .PaymentTypeTransCount(i)
                    Next
                End If

                cResponseBatchClose = ConvertTime(.ResponseBatchClose)
                cResponseBatchNetAmount = .ResponseBatchNetAmount
                cResponseBatchNumber = .ResponseBatchNumber
                cResponseBatchOpen = ConvertTime(.ResponseBatchOpen)
                cResponseBatchTransCount = .ResponseBatchTransCount
                cResponseCode = .ResponseCode
                cResponseData = .ResponseData
                cResponseInquiryCount = .ResponseInquiryCount
                cResponseSequenceNumber = .ResponseSequenceNumber
                cResponseText = .ResponseText
                cSequenceNumber = .SequenceNumber
            End With
        End Sub

        Private Function ConvertTime(ByVal timeAsString As String) As DateTime

            Try
                Dim gstart As New DateTime
                gstart = System.DateTime.ParseExact(Trim(timeAsString), "MMddyyHHmm", System.Globalization.DateTimeFormatInfo.InvariantInfo).ToString("MM/dd/yyyy HH:mm")
                Return gstart
            Catch ex As Exception
                Return Nothing
            End Try

        End Function

        Public Sub New(ByVal objFdmssettle As nsoftware.InFDMS.Fdmssettle)

            InitializeClass()

            With objFdmssettle
                cPaymentTypeCount = 1

                If cPaymentTypeCount = 0 Then
                    cPaymentTypeName = Nothing
                Else
                    ReDim cPaymentTypeName(cPaymentTypeCount)
                    ReDim cPaymentTypeNetAmount(cPaymentTypeCount)
                    ReDim cPaymentTypeTransCount(cPaymentTypeCount)

                    For i As Integer = 1 To cPaymentTypeCount
                        cPaymentTypeName(i) = "" '.PaymentTypeName(i).Trim
                        cPaymentTypeNetAmount(i) = 0 '.PaymentTypeNetAmount(i)
                        cPaymentTypeTransCount(i) = 1 '.PaymentTypeTransCount(i)
                    Next
                End If

                cResponseBatchClose = DateTime.Now.ToShortDateString
                cResponseBatchNetAmount = 0
                If .Response.BatchNumber.Length > 0 Then
                    cResponseBatchNumber = .Response.BatchNumber
                Else
                    cResponseBatchNumber = DateTime.Now.ToString("yyMMdd")
                End If
                cResponseBatchOpen = DateTime.Now.ToString("MMddyyHHmm")
                cResponseBatchTransCount = 1
                cResponseCode = .Response.DatawireReturnCode
                cResponseData = .Response.DatawireStatus
                cResponseInquiryCount = 1
                cResponseSequenceNumber = .BatchSequenceNumber
                cResponseText = .Response.DatawireStatus
                cSequenceNumber = 1 '.SequenceNumber
                cResponseText = .Response.BatchStatus

            End With
        End Sub

        Public Sub New(ByVal authNetTransactions As List(Of AuthorizeNet.Transaction))
            InitializeClass()

            Dim enumMin As Int16 = 0
            Dim enumMax As Int16 = 0

            For Each i In [Enum].GetValues(GetType(CreditCardNames))
                If CInt(i) > enumMax Then
                    enumMax = CInt(i)
                End If
            Next

            ReDim cPaymentTypeName(enumMax)
            ReDim cPaymentTypeNetAmount(enumMax)
            ReDim cPaymentTypeTransCount(enumMax)

            cPaymentTypeName(CreditCardNames.AmericanExpress) = CreditCardNames.AmericanExpress.ToString
            cPaymentTypeName(CreditCardNames.DinersClubCarteBlanche) = CreditCardNames.DinersClubCarteBlanche.ToString
            cPaymentTypeName(CreditCardNames.Discover) = CreditCardNames.Discover.ToString
            cPaymentTypeName(CreditCardNames.JCB) = CreditCardNames.JCB.ToString
            cPaymentTypeName(CreditCardNames.MasterCard) = CreditCardNames.MasterCard.ToString
            cPaymentTypeName(CreditCardNames.Visa) = CreditCardNames.Visa.ToString
            cResponseBatchNetAmount = 0

            For Each trans As AuthorizeNet.Transaction In authNetTransactions

                If trans.Status = "capturedPendingSettlement" Then
                    Select Case trans.CardType & String.Empty
                        Case "A" ' American Express
                            cPaymentTypeNetAmount(CreditCardNames.AmericanExpress) += trans.AuthorizationAmount
                            cPaymentTypeTransCount(CreditCardNames.AmericanExpress) += 1

                        Case "C" ' Diner's Club
                            cPaymentTypeNetAmount(CreditCardNames.DinersClubCarteBlanche) += trans.AuthorizationAmount
                            cPaymentTypeTransCount(CreditCardNames.DinersClubCarteBlanche) += 1

                        Case "D" ' Discover
                            cPaymentTypeNetAmount(CreditCardNames.Discover) += trans.AuthorizationAmount
                            cPaymentTypeTransCount(CreditCardNames.Discover) += 1

                        Case "E" ' Enroute

                        Case "H" ' eckeck.net

                        Case "J" ' jcb.net
                            cPaymentTypeNetAmount(CreditCardNames.JCB) += trans.AuthorizationAmount
                            cPaymentTypeTransCount(CreditCardNames.JCB) += 1

                        Case "M" ' Mastercard
                            cPaymentTypeNetAmount(CreditCardNames.MasterCard) += trans.AuthorizationAmount
                            cPaymentTypeTransCount(CreditCardNames.MasterCard) += 1

                        Case "V" ' Visa
                            cPaymentTypeNetAmount(CreditCardNames.Visa) += trans.AuthorizationAmount
                            cPaymentTypeTransCount(CreditCardNames.Visa) += 1
                        Case Else
                            cResponseBatchNetAmount += trans.AuthorizationAmount

                    End Select
                End If
            Next

            For i As Integer = enumMin To enumMax
                cResponseBatchNetAmount += cPaymentTypeNetAmount(i)
            Next

        End Sub

        Private Sub InitializeClass()
            FdmsSettlementDetailRecords = New List(Of FDMSSettlementDelegates)
            'Fdmsdetailrecord.Reset()
            'Fdmsdetailrecord.RuntimeLicense = fdmsRuntineLicense
        End Sub

        Public Sub CreatePaymentTypes()
            ReDim cPaymentTypeName(PaymentTypesList.Count)
            ReDim cPaymentTypeNetAmount(PaymentTypesList.Count)
            ReDim cPaymentTypeTransCount(PaymentTypesList.Count)

            cResponseBatchNetAmount = 0
            cResponseBatchTransCount = 0
            cResponseInquiryCount = 0

            Dim i As Integer = 0
            For Each ptype As PaymentTypes In PaymentTypesList
                cPaymentTypeName(i) = ptype.TypeName & String.Empty
                cPaymentTypeNetAmount(i) = Val(ptype.TypeNetAmount & String.Empty).ToString("###0.00")
                cPaymentTypeTransCount(i) = Val(ptype.TypeTransCount & String.Empty)
                If cPaymentTypeName(i).Length > 2 Then cPaymentTypeName(i) = cPaymentTypeName(i).Substring(0, 2)

                i += 1

                cResponseBatchNetAmount += Val(ptype.TypeNetAmount & String.Empty)
                cResponseBatchTransCount += Val(ptype.TypeTransCount & String.Empty)
                cResponseInquiryCount = PaymentTypesList.Count
                cPaymentTypeCount = PaymentTypesList.Count
            Next

        End Sub

        ''' <summary>
        ''' Total number of payment types in the current batch.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property PaymentTypeCount() As Long
            Get
                Return (cPaymentTypeCount)
            End Get
            Set(ByVal value As Long)
                cPaymentTypeCount = value
            End Set
        End Property

        ''' <summary>
        ''' Name for the payment type at the current index.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property PaymentTypeName(ByVal index As Integer) As String
            Get
                Return cPaymentTypeName(index)
            End Get

            Set(ByVal value As String)
                cPaymentTypeName(index) = value
            End Set
        End Property

        ''' <summary>
        ''' Net transaction amount for the payment type at the current index.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property PaymentTypeNetAmount(ByVal index As Integer) As Decimal
            Get
                Return cPaymentTypeNetAmount(index)
            End Get

            Set(ByVal value As Decimal)
                cPaymentTypeNetAmount(index) = value
            End Set
        End Property

        ''' <summary>
        ''' Total number of transactions for the payment type at the current index. 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property PaymentTypeTransCount(ByVal index As Integer) As Long
            Get
                Return cPaymentTypeTransCount(index)
            End Get
            Set(ByVal value As Long)
                cPaymentTypeTransCount(index) = value
            End Set
        End Property

        ''' <summary>
        ''' Date at which the current batch was closed.  
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ResponseBatchClose() As Date
            Get
                Return cResponseBatchClose
            End Get
            Set(ByVal value As Date)
                cResponseBatchClose = value
            End Set
        End Property

        ''' <summary>
        ''' Net amount of the current batch. 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ResponseBatchNetAmount() As Double
            Get
                Return cResponseBatchNetAmount
            End Get
            Set(ByVal value As Double)
                cResponseBatchNetAmount = value
            End Set
        End Property

        ''' <summary>
        ''' Current open batch number 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ResponseBatchNumber() As String
            Get
                Return cResponseBatchNumber
            End Get
            Set(ByVal value As String)

            End Set
        End Property

        ''' <summary>
        ''' Date at which the current batch was opened.  
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ResponseBatchOpen() As String
            Get
                Return cResponseBatchOpen
            End Get

            Set(ByVal value As String)
                cResponseBatchOpen = value
            End Set
        End Property

        ''' <summary>
        ''' Total number of transactions in the batch
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ResponseBatchTransCount() As Long
            Get
                Return cResponseBatchTransCount
            End Get
            Set(ByVal value As Long)

            End Set
        End Property

        ''' <summary>
        ''' Indicates the status of the authorization request. 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ResponseCode() As String
            Get
                Return cResponseCode
            End Get
            Set(ByVal value As String)
                cResponseCode = value
            End Set
        End Property

        ''' <summary>
        ''' Contains the entire contents of the Server response.  
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ResponseData() As String
            Get
                Return cResponseData
            End Get
            Set(ByVal value As String)
                cResponseData = value
            End Set
        End Property

        ''' <summary>
        ''' Number of BatchInquirys performed on the current open batch.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ResponseInquiryCount() As Long
            Get
                Return cResponseInquiryCount
            End Get
            Set(ByVal value As Long)
                cResponseInquiryCount = value
            End Set
        End Property

        ''' <summary>
        ''' SequenceNumber echoed from the authorization. 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ResponseSequenceNumber() As String
            Get
                Return cResponseSequenceNumber
            End Get
            Set(ByVal value As String)
                cResponseSequenceNumber = value
            End Set
        End Property

        ''' <summary>
        ''' Approval/Decline/Error text message information
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ResponseText() As String '
            Get
                Return cResponseText
            End Get
            Set(ByVal value As String)
                cResponseText = value
            End Set
        End Property

        ''' <summary>
        ''' Sequence number of the transaction.  
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property SequenceNumber() As Long
            Get
                Return cSequenceNumber
            End Get
            Set(ByVal value As Long)
                cSequenceNumber = value
            End Set
        End Property

    End Class

#End Region

End Class
