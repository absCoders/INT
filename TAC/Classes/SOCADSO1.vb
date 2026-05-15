Option Explicit On
Imports System.Xml.Schema

Public Class SOCADSO1

#Region "Class Variables"

    Private clsDst As DataSet = New DataSet

    Private lastErrorMessage As String = String.Empty

    Private tblTATTERM1 As DataTable = Nothing
    Private tblSOTSVIA1 As DataTable = Nothing
    Private tblICTWHSE1 As DataTable = Nothing
    Private tblTATCNTRY As DataTable = Nothing
    Private tblTATSSHK1 As DataTable = Nothing
    Private tblWHTTPLP1 As DataTable = Nothing

    Private workFileLocation As String = String.Empty
    Private shipReleaseNo As String = String.Empty

    Private WithEvents Sftp1 As New nsoftware.IPWorksSSH.SFTPClient
    Private ftpFileList As List(Of String)
    Private ftpFile As String = String.Empty

    Private nextDayDeliveries As New List(Of String)
    Private secondDayDeliveries As New List(Of String)
    Private posDeliveries As New List(Of String)
    Private priorityDeliveries As New List(Of String)

    Private Const ADSPOXmlPrefix As String = "WWIMPZPOH"
    Private Const ADSSOXmlPrefix As String = "WWIMPZSOH"
    Private clsXmlXsdError As String = String.Empty
    Private xsdsDirectory As String = String.Empty

    Public tblTasks As DataTable = Nothing

    Private Class sshAppCredentials
        Public SSH_APP_CODE As String = String.Empty
        Public SSH_APP_DESC As String = String.Empty
        Public SSH_APP_USERNAME As String = String.Empty
        Public SSH_APP_PASSWORD As String = String.Empty
        Public SSH_APP_PARTNER_URI_TEST As String = String.Empty
        Public SSH_APP_PARTNER_URI_PROD As String = String.Empty
        Public SSH_APP_PARTNER_PUBKEY_TEST As String = String.Empty
        Public SSH_APP_PARTNER_PUBKEY_PROD As String = String.Empty
        Public SSH_APP_SSH_PUBKEY As String = String.Empty
        Public SSH_APP_SSH_PVTKEY As String = String.Empty
        Public SSH_APP_FOLDER_GET As String = String.Empty
        Public SSH_APP_FOLDER_PUT As String = String.Empty
        Public SSH_APP_PGP_PUBKEY As String = String.Empty
        Public SSH_APP_PGP_PVTKEY As String = String.Empty
        Public SSH_APP_PGP_PVTKEY_PWD As String = String.Empty
        Public SSH_APP_NOTES As String = String.Empty
        Public SSH_APP_SSH_PVTKEY_PWD As String = String.Empty
        Public SSH_APP_PORT As Int32 = 22
    End Class

    Private clssshAppCredentials As New sshAppCredentials

    Private Enum CredentialsType
        SSH_App
        LP_Code
    End Enum

    Private Enum PadDirection
        Left
        Right
    End Enum

    Private Enum FileType
        Header
        Detail
        Batch
        Item
    End Enum

#End Region

#Region "Instantiate Class"

    Public Sub New()
        InitializeVariables()
    End Sub

    Private Sub InitializeVariables()
        lastErrorMessage = String.Empty

        tblTATTERM1 = ASCDATA1.GetDataTable("Select * from TATTERM1", "TATTERM1")
        tblSOTSVIA1 = ASCDATA1.GetDataTable("Select * from SOTSVIA1", "SOTSVIA1")
        tblICTWHSE1 = ASCDATA1.GetDataTable("Select * from ICTWHSE1", "ICTWHSE1")
        tblTATCNTRY = ASCDATA1.GetDataTable("Select * from TATCNTRY", "TATCNTRY")
        tblTATSSHK1 = ASCDATA1.GetDataTable("Select * from TATSSHK1", "TATSSHK1")
        tblWHTTPLP1 = ASCDATA1.GetDataTable("Select * from WHTTPLP1", "WHTTPLP1")

        xsdsDirectory = IO.Path.Combine(ASCMAIN1.Folders("SharedRoot"), "XSDs\")

        workFileLocation = ASCMAIN1.Folders("Temp")
        If workFileLocation.Length > 0 AndAlso Not workFileLocation.EndsWith("\") Then
            workFileLocation &= "\"
        End If

        Sftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwaresftpkey")

        clssshAppCredentials = New sshAppCredentials
    End Sub

#End Region

#Region "Shared Procedures and Properties"

    Private Function GetSSHAppCedentials(SSH_APP_CODE As String) As Boolean

        Try

            clssshAppCredentials = New sshAppCredentials

            Dim rowTATSSHK1 As DataRow = tblTATSSHK1.Rows.Find(SSH_APP_CODE)
            If rowTATSSHK1 Is Nothing Then
                lastErrorMessage = $"GetSSHAppCedentials for {SSH_APP_CODE} does not exist. (TATSSHK1)"
                Return False
            End If

            With clssshAppCredentials
                .SSH_APP_CODE = rowTATSSHK1.Item("SSH_APP_CODE") & String.Empty
                .SSH_APP_DESC = rowTATSSHK1.Item("SSH_APP_DESC") & String.Empty
                .SSH_APP_USERNAME = rowTATSSHK1.Item("SSH_APP_USERNAME") & String.Empty
                .SSH_APP_PASSWORD = rowTATSSHK1.Item("SSH_APP_PASSWORD") & String.Empty
                .SSH_APP_PARTNER_URI_TEST = rowTATSSHK1.Item("SSH_APP_PARTNER_URI_TEST") & String.Empty
                .SSH_APP_PARTNER_URI_PROD = rowTATSSHK1.Item("SSH_APP_PARTNER_URI_PROD") & String.Empty
                .SSH_APP_PARTNER_PUBKEY_TEST = rowTATSSHK1.Item("SSH_APP_PARTNER_PUBKEY_TEST") & String.Empty
                .SSH_APP_PARTNER_PUBKEY_PROD = rowTATSSHK1.Item("SSH_APP_PARTNER_PUBKEY_PROD") & String.Empty
                .SSH_APP_SSH_PUBKEY = rowTATSSHK1.Item("SSH_APP_SSH_PUBKEY") & String.Empty
                .SSH_APP_SSH_PVTKEY = rowTATSSHK1.Item("SSH_APP_SSH_PVTKEY") & String.Empty
                .SSH_APP_FOLDER_GET = rowTATSSHK1.Item("SSH_APP_FOLDER_GET") & String.Empty
                .SSH_APP_FOLDER_PUT = rowTATSSHK1.Item("SSH_APP_FOLDER_PUT") & String.Empty
                .SSH_APP_PGP_PUBKEY = rowTATSSHK1.Item("SSH_APP_PGP_PUBKEY") & String.Empty
                .SSH_APP_PGP_PVTKEY = rowTATSSHK1.Item("SSH_APP_PGP_PVTKEY") & String.Empty
                .SSH_APP_PGP_PVTKEY_PWD = rowTATSSHK1.Item("SSH_APP_PGP_PVTKEY_PWD") & String.Empty
                .SSH_APP_NOTES = rowTATSSHK1.Item("SSH_APP_NOTES") & String.Empty
                .SSH_APP_SSH_PVTKEY_PWD = rowTATSSHK1.Item("SSH_APP_SSH_PVTKEY_PWD") & String.Empty
                .SSH_APP_PORT = Val(rowTATSSHK1.Item("SSH_APP_PORT") & String.Empty)
            End With

            Return True

        Catch ex As Exception
            lastErrorMessage = $"GetSSHAppCedentials Error {ex.Message}"
            Return False
        End Try

    End Function

    ''' <summary>
    ''' returns the last recorded Error
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property LastError As String
        Get
            Return lastErrorMessage
        End Get
    End Property

    ''' <summary>
    ''' Returns the LP_XNO number for this Upload / Release
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ReleaseShipmentNo As String
        Get
            Return shipReleaseNo
        End Get
    End Property

    ''' <summary>
    ''' Returns a list of Priorty Shipments
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetPriorityDeliveries() As List(Of String)
        Get
            Return priorityDeliveries
        End Get
    End Property

    ''' <summary>
    ''' Returns a list of orders where the ship via is flagged as a next day delivery
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetNextDayDeliveries() As List(Of String)
        Get
            Return nextDayDeliveries
        End Get
    End Property

    ''' <summary>
    ''' Returns a list of orders where the ship via is flagged as a 2-day delivery
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetSecondDayDeliveries() As List(Of String)
        Get
            Return secondDayDeliveries
        End Get
    End Property

    ''' <summary>
    ''' Returns a list of orders where the ship via is flagged as a POS
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetPosDeliveries() As List(Of String)
        Get
            Return posDeliveries
        End Get
    End Property

#End Region

#Region "ADS XMl file creation"

    ''' <summary>
    ''' Creates the Header, Details and Batch files for order to be shipped
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PrepareADSSalesOrdersFile(ByVal shipmentList As List(Of String), ByVal WHSE_CODE As String) As Boolean

        Dim successfulTransferOfFiles As Boolean = False
        Dim sql As String = String.Empty
        lastErrorMessage = String.Empty

        Try
            Dim LP_CODE As String = String.Empty
            AddTask("SOCADSO1 Enter PrepareADSSalesOrdersFile")

            AddTask("SOCADSO1.GetCredentials")
            If Not GetCredentials(WHSE_CODE, LP_CODE) Then
                Return False
            End If

            Dim listOfShipments As String = String.Join("', '", shipmentList.ToArray)
            listOfShipments = "'" & listOfShipments & "'"

            ' Done this way in case the screen is loaded and the data on the screen is stale.
            Dim wkSOTPICK1 As String = ASCMAIN1.Temp_Table($"SELECT SOTPICK1.ORDR_NO, SOTPICK1.PICK_NO, SOTPICK1.SHIP_BOL_NO, SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, SOTORDR1.CUST_BILL_TO_CUST
                                                                FROM SOTPICK1, SOTORDR1
                                                                WHERE SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO
                                                                AND SOTORDR1.ORDR_STATUS = 'P'
                                                                AND SOTPICK1.PICK_STATUS = 'P' 
                                                                AND SOTPICK1.SHIP_BOL_NO IN ({listOfShipments})")

            AddTask("SOCADSO1 Fill SOTORDR1")
            sql = $"Select SOTORDR1.*, SOTPICK1.PICK_NO, SOTSHIP1.SHIP_ADDR_TYPE, SOTSHIP1.SHIP_ADDR_CODE, SOTSHIP1.SHIP_BOL_NO
                     , EDT850T1.EDI_SUPPLIER_NO, EDT850T1.EDI_STORE, EDT850T1.EDI_CUSTOMER, SOTSHIP1.SHIP_856_IND
                     From SOTORDR1, EDT850T1, SOTSHIP1, {wkSOTPICK1} SOTPICK1
                     Where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO
                     And SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO
                     And SOTORDR1.EDI_DOC_SEQ_NO = EDT850T1.EDI_DOC_SEQ_NO (+)"
            Dim tblSOTORDR1 As DataTable = ASCDATA1.GetDataTable(sql, "SOTORDR1")

            AddTask("SOCADSO1 Fill SOTORDR2")
            sql = $"Select SOTORDR2.*, SOTPICK2.PICK_QTY, SOTPICK1.PICK_NO, SOTPICK2.PICK_LNO
                        , ICTITEM1.ITEM_UPC_CODE, ICTITEM1.ITEM_EAN_CODE, ICTITEM1.ITEM_UOM
                        , NVL(EDT850T2.EDI_SKU, EDT850T2.EDI_STYLE) EDI_SKU, SOTPICK1.SHIP_BOL_NO
                        , ICTITEM1.NRF_SIZE_CODE, ICTITEM1.SIZE_CODE, ICTITEM1.NRF_COLOR_CODE, ICTITEM1.COLOR_CODE
                        From SOTORDR1, SOTORDR2, ICTITEM1, EDT850T2, SOTPICK2, {wkSOTPICK1} SOTPICK1
                         Where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO
                         And SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO
                         And SOTPICK1.PICK_NO = SOTPICK2.PICK_NO
                         And SOTORDR2.ITEM_CODE = ICTITEM1.ITEM_CODE (+)
                         And SOTORDR2.EDI_DOC_SEQ_NO = EDT850T2.EDI_DOC_SEQ_NO (+)
                         And SOTORDR2.EDI_DTL_SEQ = EDT850T2.EDI_DTL_SEQ (+)
                         And SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO
                         And SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO"
            Dim tblSOTORDR2 As DataTable = ASCDATA1.GetDataTable(sql, "SOTORDR2")

            AddTask("SOCADSO1 Fill SOTORDR5")
            sql = $"Select * from SOTORDR5 where ORDR_NO In (Select ORDR_NO FROM {wkSOTPICK1})"
            Dim tblSOTORDR5 As DataTable = ASCDATA1.GetDataTable(sql, "SOTORDR5")

            AddTask( "SOCADSO1 Fill ARTCUST1")
            sql = $"Select * FROM ARTCUST1 WHERE CUST_CODE In (Select NVL(CUST_BILL_TO_CUST, CUST_CODE) FROM {wkSOTPICK1})"
            Dim tblARTCUST1 As DataTable = ASCDATA1.GetDataTable(sql, "ARTCUST1", 1)

            AddTask( "SOCADSO1 Fill ARTCUST2")
            sql = $"Select * FROM ARTCUST2 WHERE (CUST_CODE, CUST_STORE_NO) In (Select CUST_CODE, CUST_STORE_NO FROM {wkSOTPICK1})"
            Dim tblARTCUST2 As DataTable = ASCDATA1.GetDataTable(sql, "ARTCUST2", 2)

            AddTask( "SOCADSO1 Fill ARTCULP2")
            sql = $"SELECT * FROM ARTCULP2 WHERE LP_CODE = '{LP_CODE}'"
            Dim tblARTCULP2 As DataTable = ASCDATA1.GetDataTable(sql, "ARTCULP2")

            AddTask( "SOCADSO1 Fill ICTITLP2")
            sql = $"SELECT * FROM ICTITLP2 WHERE LP_CODE = '{LP_CODE}'"
            Dim tblICTITLP2 As DataTable = ASCDATA1.GetDataTable(sql, "ICTITLP2")

            AddTask( "SOCADSO1 Fill WHTTPLP1")
            sql = $"SELECT * FROM WHTTPLP1 WHERE LP_CODE = '{LP_CODE}'"
            Dim rowWHTTPLP1 As DataRow = ASCDATA1.GetDataRow(sql)

            Dim dstSalesOrders As New DataSet
            dstSalesOrders.ReadXml(System.IO.Path.Combine(xsdsDirectory, $"{ADSSOXmlPrefix}.xsd"))
            dstSalesOrders.EnforceConstraints = False
            For Each tbl As DataTable In dstSalesOrders.Tables
                tbl.Rows.Clear()
            Next
            dstSalesOrders.EnforceConstraints = True

            For Each fieldName As String In New String() {"SalesOrderDate", "RequestedDeliveryDat", "StartShipDate", "ShipDate"}
                dstSalesOrders.Tables("Order").Columns(fieldName).DataType = GetType(System.String)
            Next

            Dim xsdFile As String = System.IO.Path.Combine(xsdsDirectory, $"{ADSSOXmlPrefix}.xsd")

            Dim tblSOTSHIP1 As DataTable = ASCDATA1.SelectDistinct(tblSOTORDR1, New String() {"SHIP_BOL_NO"})

            Dim ORDER_ID As Int16 = 0

            AddTask( "SOCADSO1 Process Shipments")
            For Each rowSOTSHIP1 As DataRow In tblSOTSHIP1.Select("", "SHIP_BOL_NO")
                Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO") & String.Empty
                AddTask( $"SOCADSO1 Process Shipment {SHIP_BOL_NO}")

                For Each rowSOTORDR1 As DataRow In tblSOTORDR1.Select($"SHIP_BOL_NO = '{SHIP_BOL_NO}'", "PICK_NO")
                    Dim PICK_NO As String = rowSOTORDR1.Item("PICK_NO") & String.Empty

                    ' If for some reason there are no Items To Pick then no need to process
                    If tblSOTORDR2.Select($"PICK_NO = '{PICK_NO}' and PICK_QTY > 0").Length = 0 Then
                        Continue For
                    End If

                    Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO") & String.Empty
                    Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE") & String.Empty
                    Dim CUST_STORE_NO As String = rowSOTORDR1.Item("CUST_STORE_NO") & String.Empty
                    Dim CUST_DC_NO As String = rowSOTORDR1.Item("CUST_DC_NO") & String.Empty
                    Dim CUST_BILL_TO_CUST As String = rowSOTORDR1.Item("CUST_BILL_TO_CUST") & String.Empty

                    CUST_BILL_TO_CUST = CUST_BILL_TO_CUST.Trim
                    If CUST_BILL_TO_CUST.Length = 0 Then
                        CUST_BILL_TO_CUST = CUST_CODE
                    End If

                    Dim CUST_NO_3PL As String = String.Empty ' rowARTCULP2.Item("CUST_NO_3PL") & String.Empty
                    Dim CUST_STORE_NO_3PL As String = String.Empty ' rowARTCULP2.Item("CUST_STORE_NO_3PL") & String.Empty

                    Dim rowARTCULP2 As DataRow = tblARTCULP2.Rows.Find(New Object() {LP_CODE, CUST_CODE, CUST_STORE_NO})
                    If rowWHTTPLP1 Is Nothing OrElse Val(rowWHTTPLP1.Item("TRANSMIT_ALL_STORES") & String.Empty) = 0 Then
                        If rowARTCULP2 Is Nothing Then
                            lastErrorMessage = $"LP Code {LP_CODE} does not have an entry for Ship To {CUST_CODE} / {CUST_STORE_NO}"
                            Return False
                        End If
                    End If

                    If rowARTCULP2 IsNot Nothing Then
                        CUST_NO_3PL = rowARTCULP2.Item("CUST_NO_3PL") & String.Empty
                        CUST_STORE_NO_3PL = rowARTCULP2.Item("CUST_STORE_NO_3PL") & String.Empty
                    Else
                        CUST_NO_3PL = CUST_CODE
                        CUST_STORE_NO_3PL = CUST_STORE_NO
                    End If

                    If CUST_NO_3PL.Length = 0 OrElse CUST_STORE_NO_3PL.Length = 0 Then
                        lastErrorMessage = $"LP Code {LP_CODE}, Ship To {CUST_CODE} / {CUST_STORE_NO} is not assigned an 3PL Customer and Ship To value"
                        Return False
                    End If

                    Dim rowSOTORDR5_ST As DataRow = tblSOTORDR5.Rows.Find(New Object() {ORDR_NO, "ST"})
                    If rowSOTORDR5_ST Is Nothing Then
                        lastErrorMessage = $"Missing Shipping Address Entry for Order No: {ORDR_NO}"
                        Return False
                    End If

                    If rowSOTORDR5_ST.Item("CUST_NAME") & String.Empty = String.Empty Then
                        rowSOTORDR5_ST.Item("CUST_NAME") = rowSOTORDR1.Item("CUST_NAME")
                    End If

                    Dim rowARTCUST1 As DataRow = tblARTCUST1.Rows.Find(CUST_CODE)
                    If rowARTCUST1 Is Nothing Then
                        AddTask( $"SOCADSO1 Load Customer {CUST_CODE}")
                        rowARTCUST1 = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = :PARM1", "V", New Object() {CUST_CODE})
                        If rowARTCUST1 Is Nothing Then
                            lastErrorMessage = $"Invalid Customer Code: {CUST_CODE} for Order No: {ORDR_NO}"
                            Return False
                        End If
                        tblARTCUST1.ImportRow(rowARTCUST1)
                    End If

                    Dim rowSOTORDR5_BT As DataRow = tblARTCUST1.Rows.Find(CUST_BILL_TO_CUST)
                    If rowSOTORDR5_BT Is Nothing Then
                        AddTask( $"SOCADSO1 Load CUST_BILL_TO_CUST {CUST_BILL_TO_CUST}")
                        rowSOTORDR5_BT = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = :PARM1", "V", New Object() {CUST_BILL_TO_CUST})
                        If rowSOTORDR5_BT Is Nothing Then
                            lastErrorMessage = $"Invalid Bill To Customer Code: {CUST_BILL_TO_CUST} for Order No: {ORDR_NO}"
                            Return False
                        End If
                        tblARTCUST1.ImportRow(rowSOTORDR5_BT)
                    End If

                    Dim rowARTCUST2 As DataRow = tblARTCUST2.Rows.Find(New Object() {CUST_CODE, CUST_STORE_NO})
                    If rowARTCUST2 Is Nothing Then
                        lastErrorMessage = $"Invalid Customer Code: {CUST_CODE} Store No: {CUST_STORE_NO} for Order No: {ORDR_NO}"
                        Return False
                    End If

                    Dim rowTATTERM1 As DataRow = tblTATTERM1.Rows.Find(rowSOTORDR1.Item("TERM_CODE") & String.Empty)
                    If rowTATTERM1 Is Nothing Then
                        lastErrorMessage = $"Invalid Terms Code: {rowSOTORDR1.Item("TERM_CODE")} for Customer Code: {CUST_CODE} for Order No: {ORDR_NO}"
                        Return False
                    End If

                    Dim rowSOTSVIA1 As DataRow = tblSOTSVIA1.Rows.Find(rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty)
                    If rowSOTSVIA1 Is Nothing Then
                        lastErrorMessage = $"Invalid Ship Via Code: {rowSOTORDR1.Item("SHIP_VIA_CODE")} for Order No: {ORDR_NO})"
                        Return False
                    End If

                    AddTask( $"SOCADSO1 Gather Sales Order Header Data")
                    Dim rowSOH As DataRow = dstSalesOrders.Tables("Order").NewRow
                    With rowSOH
                        '.Item("SalesSite") = ORDR_NO
                        .Item("AutoCreateIfMissing") = "Y"
                        '.Item("EDISenderID") = ORDR_NO
                        '.Item("EDIReceiverID") = ORDR_NO

                        ' Changed on 07/17/2025, ADS accepts only numbers in the Sales Order Number field.
                        ' The A may have been used if ADS Sales Order Number only permits 7 characters and our Pick Nos roll over.
                        ' 00100000012 and 00200000012 would send over Sales Order Number 12.
                        'Dim SalesOrderNumber As String = Val(PICK_NO).ToString.PadLeft(7, "0")
                        'If SalesOrderNumber.Length > 7 Then
                        '    SalesOrderNumber = "A" & StrReverse(StrReverse(SalesOrderNumber).Substring(0, 6))
                        'End If

                        ' ADS permits a maximum of 7 numeric characters for the Order Number, no leading zeroes
                        .Item("SalesOrderNumber") = Val(StrReverse(StrReverse(PICK_NO).Substring(0, 7)))
                        .Item("ExternalOrderNumber") = SHIP_BOL_NO
                        .Item("SoldToCustomer") = CUST_NO_3PL

                        If IsDate(rowSOTORDR1.Item("ORDR_DATE") & String.Empty) Then
                            .Item("SalesOrderDate") = CDate(rowSOTORDR1.Item("ORDR_DATE") & String.Empty).ToString("yyyy-MM-dd")
                        End If

                        .Item("CustomerPONumber") = rowSOTORDR1.Item("ORDR_CUST_PO")
                        .Item("Department") = rowSOTORDR1.Item("ORDR_DEPT")
                        ' 10/31/2025 - send EDI_MERCH_TYPE instead of SALES_DIVISION_CODE 
                        .Item("Division") = rowSOTORDR1.Item("EDI_MERCH_TYPE")
                        .Item("Currency") = "USD"

                        ' Issue-7255 20251223
                        .Item("Carrier") = rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty ' "TBD"

                        ' 20240904 - Requested by Nathan, Ticket ID INC-6034
                        ' We need to keep “RGOF” for Clarins but, for Military orders which Whse = ADS, can we hard code to send “FXG” as the Ship Via / Carrier in ADS’ file, so Sage can receive as “FXG” directly and not “TBD”?
                        '20251223 - Removed by Nathan
                        'Select Case CUST_CODE
                        '    Case "AAFES", "MCX", "NEXCOM", "USCG", "VETERANS"
                        '        .Item("Carrier") = "FXG"
                        'End Select

                        '.Item("ThirdPartyFrtAcct") = String.Empty
                        '.Item("FreightInvoicing") = String.Empty
                        '.Item("CustomerFreight") = String.Empty
                        '.Item("PackingList") = String.Empty

                        If IsDate(rowSOTORDR1.Item("ORDR_CANCEL_DATE") & String.Empty) Then
                            .Item("RequestedDeliveryDat") = CDate(rowSOTORDR1.Item("ORDR_CANCEL_DATE") & String.Empty).ToString("yyyy-MM-dd")
                        ElseIf IsDate(rowSOTORDR1.Item("ORDR_ARRIVAL_DATE") & String.Empty) Then
                            .Item("RequestedDeliveryDat") = CDate(rowSOTORDR1.Item("ORDR_ARRIVAL_DATE") & String.Empty).ToString("yyyy-MM-dd")
                        End If

                        If IsDate(rowSOTORDR1.Item("ORDR_SHIP_DATE") & String.Empty) Then
                            .Item("StartShipDate") = CDate(rowSOTORDR1.Item("ORDR_SHIP_DATE") & String.Empty).ToString("yyyy-MM-dd")
                            .Item("ShipDate") = CDate(rowSOTORDR1.Item("ORDR_SHIP_DATE") & String.Empty).ToString("yyyy-MM-dd")
                        End If

                        .Item("BillToCode") = "MAIN" 'CUST_BILL_TO_CUST
                        .Item("BillToName") = rowSOTORDR5_BT.Item("CUST_NAME")
                        .Item("BillToAddress1") = rowSOTORDR5_BT.Item("CUST_ADDR1")
                        .Item("BillToAddress2") = rowSOTORDR5_BT.Item("CUST_ADDR2")
                        .Item("BillToAddress3") = rowSOTORDR5_BT.Item("CUST_ADDR3")
                        .Item("BillToCity") = rowSOTORDR5_BT.Item("CUST_CITY")
                        .Item("BillToState") = rowSOTORDR5_BT.Item("CUST_STATE")
                        .Item("BillToZipcode") = rowSOTORDR5_BT.Item("CUST_ZIP_CODE")

                        Dim rowTATCNTRY As DataRow = Nothing
                        Dim CUST_COUNTRY As String = rowSOTORDR5_BT.Item("CUST_COUNTRY") & String.Empty
                        If CUST_COUNTRY.Length = 0 Then
                            CUST_COUNTRY = "US"
                        End If
                        If CUST_COUNTRY.Length = 3 Then
                            rowTATCNTRY = tblTATCNTRY.Rows.Find(CUST_COUNTRY)
                            If rowTATCNTRY IsNot Nothing Then
                                CUST_COUNTRY = rowTATCNTRY.Item("COUNTRY_CODE2") & String.Empty
                            End If
                        End If
                        .Item("BillToCountry") = CUST_COUNTRY

                        .Item("ShipToCode") = CUST_STORE_NO_3PL
                        .Item("ShipToName") = rowSOTORDR5_ST.Item("CUST_NAME")
                        .Item("ShipToAddress1") = rowSOTORDR5_ST.Item("CUST_ADDR1")

                        Dim SHIPMENT_AUTH_NO As String = rowSOTORDR1.Item("SHIPMENT_AUTH_NO") & String.Empty

                        ' ISSUE-7428  Shipment Authorization No. Development
                        ' .Item("ShipToAddress2") = rowSOTORDR5_ST.Item("CUST_ADDR2")
                        If rowSOTORDR5_ST.Item("CUST_ADDR2") & String.Empty = String.Empty Then
                            .Item("ShipToAddress2") = SHIPMENT_AUTH_NO
                            SHIPMENT_AUTH_NO = String.Empty
                        Else
                            .Item("ShipToAddress2") = rowSOTORDR5_ST.Item("CUST_ADDR2")
                        End If

                        ' ISSUE-7428  Shipment Authorization No. Development
                        '.Item("ShipToAddress3") = rowSOTORDR5_ST.Item("CUST_ADDR3")
                        If rowSOTORDR5_ST.Item("CUST_ADDR3") & String.Empty = String.Empty Then
                            .Item("ShipToAddress3") = SHIPMENT_AUTH_NO
                            SHIPMENT_AUTH_NO = String.Empty
                        Else
                            Dim CUST_ADDR3 As String = rowSOTORDR5_ST.Item("CUST_ADDR3") & ""
                            .Item("ShipToAddress3") = CUST_ADDR3 & " " & SHIPMENT_AUTH_NO
                        End If

                        .Item("ShipToCity") = rowSOTORDR5_ST.Item("CUST_CITY")
                        .Item("ShipToState") = rowSOTORDR5_ST.Item("CUST_STATE")
                        .Item("ShipToZipcode") = rowSOTORDR5_ST.Item("CUST_ZIP_CODE")

                        CUST_COUNTRY = rowSOTORDR5_ST.Item("CUST_COUNTRY") & String.Empty
                        If CUST_COUNTRY.Length = 0 Then
                            CUST_COUNTRY = "US"
                        End If
                        If CUST_COUNTRY.Length = 3 Then
                            rowTATCNTRY = tblTATCNTRY.Rows.Find(CUST_COUNTRY)
                            If rowTATCNTRY IsNot Nothing Then
                                CUST_COUNTRY = rowTATCNTRY.Item("COUNTRY_CODE2") & String.Empty
                            End If
                        End If
                        .Item("ShipToCountry") = CUST_COUNTRY

                        .Item("DcCode") = CUST_DC_NO
                        '.Item("Order_ID") = String.Empty
                        '.Item("TaxCode") = String.Empty
                        '.Item("UctTaxAmount") = String.Empty
                        '.Item("UctTaxRate") = String.Empty
                        .Item("DiscountPercent") = 0
                        .Item("DiscountAmount") = 0
                    End With

                    dstSalesOrders.Tables("Order").Rows.Add(rowSOH)
                    ORDER_ID = rowSOH.Item("Order_ID")

                    AddTask( $"SOCADSO1 Gather Sales Order Detail Data")
                    For Each rowSOTORDR2 As DataRow In tblSOTORDR2.Select($"PICK_NO = '{PICK_NO}' AND ORDR_QTY_PICK > 0", "ORDR_LNO")
                        Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE")
                        Dim ORDR_QTY_PICK As Int32 = Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & String.Empty)
                        If ORDR_QTY_PICK <= 0 Then Continue For

                        Dim rowLine As DataRow = dstSalesOrders.Tables("Line").NewRow
                        With rowLine
                            .Item("Order_ID") = ORDER_ID
                            .Item("LineNumber") = Val(rowSOTORDR2.Item("PICK_LNO") & "")

                            Dim ItemReferenceNumber As String = String.Empty
                            If tblICTITLP2.Select($"ITEM_CODE = '{ITEM_CODE}'").Length = 1 Then
                                ItemReferenceNumber = tblICTITLP2.Select($"ITEM_CODE = '{ITEM_CODE}'")(0).Item("ITEM_CODE_3PL") & String.Empty
                            End If

                            If rowWHTTPLP1 Is Nothing OrElse Val(rowWHTTPLP1.Item("TRANSMIT_ALL_ITEMS") & String.Empty) = 0 Then
                                If ItemReferenceNumber.Length = 0 Then
                                    lastErrorMessage = $"Invalid Item Code: {ITEM_CODE} for Order No: {ORDR_NO}. Not mapped to 3PL Item Code."
                                    Return False
                                End If
                            Else
                                If ItemReferenceNumber.Length = 0 Then
                                    ItemReferenceNumber = ITEM_CODE
                                End If
                            End If

                            .Item("ItemReferenceNumber") = ItemReferenceNumber
                            .Item("CustomerItemRef") = rowSOTORDR2.Item("EDI_SKU")
                            Dim ITEM_EAN_CODE As String = rowSOTORDR2.Item("ITEM_EAN_CODE") & ""
                            Dim ITEM_UPC_CODE As String = rowSOTORDR2.Item("ITEM_UPC_CODE") & ""
                            Dim EAN_UPC As String = ITEM_EAN_CODE
                            If EAN_UPC = "" Then
                                EAN_UPC = ITEM_UPC_CODE
                            End If
                            .Item("UpcEanCode") = EAN_UPC
                            .Item("EDISize") = rowSOTORDR2.Item("NRF_SIZE_CODE")
                            .Item("EDIColor") = rowSOTORDR2.Item("NRF_COLOR_CODE")

                            Dim ITEM_DESC As String = rowSOTORDR2.Item("ITEM_DESC") & String.Empty
                            If ITEM_DESC.Length > dstSalesOrders.Tables("Line").Columns("ItemDescription").MaxLength Then
                                ITEM_DESC = ITEM_DESC.Substring(0, dstSalesOrders.Tables("Line").Columns("ItemDescription").MaxLength).Trim
                            End If
                            .Item("ItemDescription") = ITEM_DESC
                            .Item("ItemQuantity") = Val(rowSOTORDR2.Item("PICK_QTY") & "")
                            .Item("UnitOfMeasure") = rowSOTORDR2.Item("ITEM_UOM") ' "EA"
                            '.Item("UomConversionQty") = ""
                            .Item("ItemGrossPrice") = Val(rowSOTORDR2.Item("ITEM_RETAIL_PRICE") & "")
                            .Item("ItemCustGrossPrice") = Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & "")
                            .Item("ItemDiscount1") = 0
                        End With
                        dstSalesOrders.Tables("Line").Rows.Add(rowLine)
                    Next
                Next
            Next

            shipReleaseNo = ASCMAIN1.Next_Control_No("SOTSHIP1.LP_XNO")
            Dim xmlFile As String = System.IO.Path.Combine(ASCMAIN1.Folders("Temp"), $"{ADSSOXmlPrefix}_{shipReleaseNo}.XML")
            If Not My.Computer.FileSystem.DirectoryExists(System.IO.Path.GetDirectoryName(xmlFile)) Then
                AddTask( $"SOCADSO1 Create Directory: {System.IO.Path.GetDirectoryName(xmlFile)}")
                My.Computer.FileSystem.CreateDirectory(System.IO.Path.GetDirectoryName(xmlFile))
            End If

            AddTask( $"SOCADSO1 Start WriteXml Document {xmlFile}")
            dstSalesOrders.WriteXml(xmlFile)
            AddTask( $"SOCADSO1 Finish WriteXml Document {xmlFile}")

            clsXmlXsdError = String.Empty
            If My.Computer.FileSystem.FileExists(xsdFile) Then
                AddTask( $"SOCADSO1 Start LoadValidatedXmlDocument")
                LoadValidatedXmlDocument(xmlFile, xsdFile)
                AddTask( $"SOCADSO1 Finish LoadValidatedXmlDocument")
            End If

            If clsXmlXsdError.Length > 0 Then
                lastErrorMessage = clsXmlXsdError
                Return False
            End If

            ' Send any new items
            Try
                AddTask( $"SOCADSO1 Start Transmit_Document")
                Dim XMIT_NO As String = TAC.ICCMAIN1.Transmit_Document("WHC", "WHC888O1", "AC", "", "ADS")
                AddTask( $"SOCADSO1 Finish Transmit_Document")
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Send Items to ADS", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            ' Send XML to ADS
            If UploadFilesFtp(xmlFile) Then
                Try
                    Dim archiveDirectory As String = System.IO.Path.Combine(ASCMAIN1.Folders("Archive"), "ADS", DateTime.Now.ToString("yyyyMM"))
                    If Not My.Computer.FileSystem.DirectoryExists(archiveDirectory) Then
                        AddTask( $"SOCADSO1 Start Create Archive Directory")
                        My.Computer.FileSystem.CreateDirectory(archiveDirectory)
                        AddTask( $"SOCADSO1 Finish Create Archive Directory")
                    End If

                    Dim destFileName As String = System.IO.Path.Combine(archiveDirectory, System.IO.Path.GetFileName(xmlFile))
                    AddTask( $"SOCADSO1 Start MoveFile {xmlFile} {destFileName}")
                    My.Computer.FileSystem.MoveFile(xmlFile, destFileName, True)
                    AddTask( $"SOCADSO1 Finish MoveFile {xmlFile} {destFileName}")
                Catch ex As Exception

                End Try

                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            lastErrorMessage = ex.Message
            Return False
        Finally

        End Try
    End Function

    ''' <summary>
    ''' Creates PO Xml File for ADS
    ''' </summary>
    ''' <param name="WHSE_CODE"></param>
    ''' <param name="PO_ORDER_NO"></param>
    ''' <param name="PINV_NO"></param>
    ''' <param name="XMIT_NO"></param>
    ''' <returns></returns>
    Public Function PrepreADSPOFile(WHSE_CODE As String,
                                       PO_ORDER_NO As String,
                                       PINV_NO As String,
                                       ByRef XMIT_NO As String) As Boolean

        Try
            Dim LP_CODE As String = String.Empty

            If Not GetCredentials(WHSE_CODE, LP_CODE) Then
                Return False
            End If

            Dim dstPurchaseOrders As New DataSet
            dstPurchaseOrders.ReadXml(System.IO.Path.Combine(xsdsDirectory, $"{ADSPOXmlPrefix}.xml"))
            dstPurchaseOrders.EnforceConstraints = False
            For Each tbl As DataTable In dstPurchaseOrders.Tables
                tbl.Rows.Clear()
            Next
            dstPurchaseOrders.EnforceConstraints = True

            Dim xsdFile As String = System.IO.Path.Combine(xsdsDirectory, $"{ADSPOXmlPrefix}.xsd")

            ASCMAIN1.sql = "Select ICTPINV1.*, POTORDR1.WHSE_CODE
                    from POTORDR1,ICTPINV1
                    where POTORDR1.PO_ORDER_NO = ICTPINV1.PO_ORDER_NO 
                    and ICTPINV1.PINV_STATUS = 'O' 
                    and ICTPINV1.PO_ORDER_NO = :PARM1
                    and ICTPINV1.PINV_NO = :PARM2"

            Dim tblICTPINV1 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
            If tblICTPINV1.Rows.Count = 0 Then
                lastErrorMessage = $"No records found for PO {PO_ORDER_NO}, Pinv No: {PINV_NO}."
                Return False
            End If

            ASCMAIN1.sql = $"Select ICTPINV2.*, POTORDR2.PO_DATE_REQUIRED
                        , ICTITEM1.NRF_SIZE_CODE, ICTITEM1.SIZE_CODE, ICTITEM1.NRF_COLOR_CODE, ICTITEM1.COLOR_CODE
                        from POTORDR2,ICTITEM1,ICTPINV2
                        where ICTITEM1.ITEM_CODE = POTORDR2.ITEM_CODE
                        and POTORDR2.PO_ORDER_NO = ICTPINV2.PO_ORDER_NO
                        and POTORDR2.PO_ORDER_LNO = ICTPINV2.PO_ORDER_LNO
                        AND ICTPINV2.PINV_NO = '{PINV_NO}'"
            Dim tblICTPINV2 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
            If tblICTPINV2.Rows.Count = 0 Then
                lastErrorMessage = $"No detail records found for PO {PO_ORDER_NO}, Pinv No: {PINV_NO}."
                Return False
            End If

            For Each rowICTPINV1 As DataRow In tblICTPINV1.Select("")
                Dim INV_DATE As Date = rowICTPINV1.Item("INV_DATE")
                Dim INV_NUM As String = rowICTPINV1.Item("INV_NUM")

                Dim rowPOH As DataRow = dstPurchaseOrders.Tables("PO").NewRow
                With rowPOH
                    .Item("Supplier") = "MAIN" ' rowICTPINV1.Item("VEND_CODE")
                    .Item("ExpectedReceiptDate") = INV_DATE.AddDays(30)
                    .Item("PurchaseOrderNumber") = INV_NUM
                    .Item("InternalReference") = rowICTPINV1.Item("PO_ORDER_NO")
                End With
                dstPurchaseOrders.Tables("PO").Rows.Add(rowPOH)
                Dim PO_ID As Int64 = Val(rowPOH.Item("PO_ID") & "")

                For Each rowICTPINV2 As DataRow In tblICTPINV2.Select("", "PINV_LNO")
                    Dim ITEM_CODE As String = rowICTPINV2.Item("ITEM_CODE")
                    Dim PINV_QTY As Int32 = Val(rowICTPINV2.Item("PINV_QTY") & "")
                    Dim PINV_LNO As Int32 = Val(rowICTPINV2.Item("PINV_LNO") & "")
                    Dim PO_ORDER_LNO As Int32 = Val(rowICTPINV2.Item("PO_ORDER_LNO") & "")

                    If rowICTPINV2.Item("PO_DATE_REQUIRED") & "" <> "" Then
                        If rowPOH.Item("ExpectedReceiptDate") & "" = "" Then
                            rowPOH.Item("ExpectedReceiptDate") = rowICTPINV2.Item("PO_DATE_REQUIRED")
                        End If
                    End If

                    Dim rowLine As DataRow = dstPurchaseOrders.Tables("Line").NewRow
                    With rowLine
                        .Item("StockNumber") = ITEM_CODE
                        .Item("Quantity") = PINV_QTY
                        .Item("CustomerOrderRef") = ""
                        .Item("EDISize") = rowICTPINV2.Item("NRF_SIZE_CODE")
                        .Item("EDIColor") = rowICTPINV2.Item("NRF_COLOR_CODE")
                        .Item("PO_ID") = PO_ID
                    End With
                    dstPurchaseOrders.Tables("Line").Rows.Add(rowLine)
                Next
            Next

            XMIT_NO = ASCMAIN1.Next_Control_No("WHT3PLX1.XMIT_NO")
            Dim xmlFile As String = System.IO.Path.Combine(ASCMAIN1.Folders("Temp"), $"PO_{XMIT_NO}.XML")
            If Not My.Computer.FileSystem.DirectoryExists(System.IO.Path.GetDirectoryName(xmlFile)) Then
                My.Computer.FileSystem.CreateDirectory(System.IO.Path.GetDirectoryName(xmlFile))
            End If

            dstPurchaseOrders.WriteXml(xmlFile)

            clsXmlXsdError = String.Empty
            If My.Computer.FileSystem.FileExists(xsdFile) Then
                LoadValidatedXmlDocument(xmlFile, xsdFile)
            End If

            If clsXmlXsdError.Length > 0 Then
                lastErrorMessage = clsXmlXsdError
                Return False
            End If

            ' Send XML to ADS
            Return UploadFilesFtp(xmlFile)

        Catch ex As Exception
            lastErrorMessage = ex.Message
            Return False
        End Try
    End Function

    Private Function GetCredentials(WHSE_CODE As String, ByRef LP_CODE As String) As Boolean

        Dim rowICTWHSE1 As DataRow = tblICTWHSE1.Rows.Find(WHSE_CODE)
        If rowICTWHSE1 Is Nothing Then
            lastErrorMessage = $"Cannot locate warehouse {WHSE_CODE} in warehouse master"
            Return False
        End If

        LP_CODE = rowICTWHSE1.Item("LP_CODE") & String.Empty
        LP_CODE = LP_CODE.Trim
        If LP_CODE.Length = 0 Then
            lastErrorMessage = $"Warehouse {WHSE_CODE} is not assigned an LP Code."
            Return False
        End If

        Dim rowWHTTPLP1 As DataRow = tblWHTTPLP1.Rows.Find(LP_CODE)
        If rowWHTTPLP1 Is Nothing Then
            lastErrorMessage = $"Warehouse {WHSE_CODE}, LP Code {LP_CODE} does not exist in WHTTPLP1."
            Return False
        End If

        Dim SSH_APP_CODE As String = rowWHTTPLP1.Item("SSH_APP_CODE") & String.Empty
        SSH_APP_CODE = SSH_APP_CODE.Trim
        If SSH_APP_CODE.Length = 0 Then
            lastErrorMessage = $"Warehouse {WHSE_CODE}, LP Code {LP_CODE} in WHTTPLP1 does not have a value in SSH_APP_CODE."
            Return False
        End If

        Dim rowTATSSHK1 As DataRow = tblTATSSHK1.Rows.Find(SSH_APP_CODE)
        If rowTATSSHK1 Is Nothing Then
            lastErrorMessage = $"Warehouse {WHSE_CODE}, LP Code {LP_CODE}, SSH_APP_CODE Code {SSH_APP_CODE} does not exist in TATSSHK1."
            Return False
        End If

        If Not GetSSHAppCedentials(SSH_APP_CODE) Then
            Return False
        End If

        Return True

    End Function

#End Region

#Region "Clarins Warehouse Processing"

    ''' <summary>
    ''' Creates the Header, Details and Batch files for order to be shipped
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Prepare_Clarins_File(ByRef shipmentList As List(Of String)) As Boolean

        Prepare_Clarins_File = False
        lastErrorMessage = String.Empty
        Dim success As Boolean = False

        nextDayDeliveries.Clear()
        secondDayDeliveries.Clear()
        posDeliveries.Clear()
        priorityDeliveries.Clear()

        Try
            Dim ErrorMessages As New List(Of String)
            Dim listOfShipments As String = String.Empty

            If shipmentList IsNot Nothing AndAlso shipmentList.Count > 0 Then
                listOfShipments = String.Join("', '", shipmentList.ToArray)
                listOfShipments = "'" & listOfShipments & "'"
            End If

            AddTask("SOCADSO1 Enter Prepare_Clarins_File")

            AddTask("SOCADSO1 Load Clarin Shipments")
            Dim tblSHIPMENTS As DataTable = ASCDATA1.GetDataTable("SELECT SOTSHIP1.SHIP_BOL_NO, SOTSHIP1.SHIP_VIA_CODE, SOTSVIA1.SHIP_VIA_TRANSIT_3PL, SOTPICK1.PICK_NO" _
                                                                  & " FROM SOTSHIP1, SOTSVIA1, SOTPICK1" _
                                                                  & " WHERE SOTSHIP1.SHIP_BOL_NO IN (" & listOfShipments & ")" _
                                                                  & " AND SOTSHIP1.SHIP_VIA_CODE = SOTSVIA1.SHIP_VIA_CODE" _
                                                                  & " AND SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO")

            Dim appCmd As String = "WHSE"
            Dim appKey As String = "CLA"

            If listOfShipments.Length > 0 Then
                appCmd = "SHIP"
                appKey = listOfShipments
            End If

            ' This process prepares the files and places them in the sftp directory
            AddTask("SOCADSO1 Call TAC.ICCMAIN1.Transmit_Document")
            ' Change for ADS 07/16/2025, force developer to supply the LP CODE
            TAC.ICCMAIN1.Transmit_Document("WHC", "WHC940O1", appCmd, appKey, ErrorMessages, success, "CLA")
            AddTask("SOCADSO1 Exit TAC.ICCMAIN1.Transmit_Document")

            If TAC.ICCMAIN1.tblTasks IsNot Nothing Then
                For Each row As DataRow In TAC.ICCMAIN1.tblTasks.Select("", "SEQ_NO")
                    tblTasks.Rows.Add({tblTasks.Rows.Count + 1, row.Item("TASK_TIME"), row.Item("TASK_DESC")})
                Next
            End If


            ' Clear the list and extract the shipments that were processed.
            shipmentList.Clear()

            For Each emsg As String In ErrorMessages
                emsg = (emsg & String.Empty).Trim
                If emsg.StartsWith("ShipmentsSentTo3pl:") Then
                    emsg = emsg.Replace("ShipmentsSentTo3pl:", "")
                    emsg = emsg.Replace("'", "")
                    emsg = emsg.Replace(" ", "")

                    Dim bog As String() = emsg.Split(",")
                    For Each shipment As String In bog
                        shipment = shipment.Trim
                        If shipment.Length > 0 Then
                            shipmentList.Add(shipment)

                            For Each rowSHIPMENT As DataRow In tblSHIPMENTS.Select("SHIP_BOL_NO = '" & shipment & "'")
                                Dim PICK_NO As String = rowSHIPMENT.Item("PICK_NO")
                                If Val(rowSHIPMENT.Item("SHIP_VIA_TRANSIT_3PL") & String.Empty) = 1 Then
                                    nextDayDeliveries.Add(PICK_NO)
                                ElseIf Val(rowSHIPMENT.Item("SHIP_VIA_TRANSIT_3PL") & String.Empty) = 2 Then
                                    secondDayDeliveries.Add(PICK_NO)
                                ElseIf Val(rowSHIPMENT.Item("SHIP_VIA_TRANSIT_3PL") & String.Empty) = 3 Then
                                    posDeliveries.Add(PICK_NO)
                                End If
                            Next
                        End If
                    Next
                ElseIf emsg.Length > 0 Then
                    lastErrorMessage &= vbCr & emsg
                End If
            Next

            Prepare_Clarins_File = True

        Catch ex As Exception
            lastErrorMessage = ex.Message
        Finally
            AddTask("SOCADSO1 Exit Prepare_Clarins_File")
        End Try

    End Function

    ''' <summary>
    ''' Convert Clarins Shipment Files to EDT945T1/2 data
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ProcessClarinsShipConfirmationData(ByRef tblEDT945T1 As DataTable, ByRef tblEDT945T2 As DataTable) As Boolean

        Try
            ' It appears sometimes the clarins cancellations come after the shipments. This puts the canceled Pick tickets into EDT945T1
            Dim sql As String = String.Empty
            sql = " BEGIN DECLARE CURSOR C1 IS"
            sql &= " SELECT SOTPICK1.PICK_NO, SOTPICK1.SHIP_BOL_NO"
            sql &= " FROM SOTPICK1,CONV.CFG_CNLORDS CNLORDS,"
            sql &= " (SELECT PICK_NO, MAX(ORDR_NO_3PL) ORDR_NO_3PL FROM SOTPICKC WHERE ORDR_NO_3PL IS NOT NULL GROUP BY PICK_NO) SOTPICKC"
            sql &= " WHERE SOTPICK1.PICK_STATUS = 'P'"
            sql &= " AND SOTPICK1.PICK_NO = SOTPICKC.PICK_NO"
            sql &= " AND SOTPICKC.ORDR_NO_3PL = CNLORDS.OHORDN"
            sql &= " AND SOTPICK1.PICK_NO NOT IN (SELECT EDI_PICK_NO FROM EDT945T1);"
            sql &= " BEGIN FOR R1 IN C1 LOOP"
            sql &= " INSERT INTO EDT945T1 "
            sql &= " Select 'C' || SUBSTR(SOTPICK1.PICK_NO, 2) EDI_DOC_SEQ_NO, NULL GEN_DOC_NO, NULL EDI_ISA_NO, "
            sql &= " NULL EDI_TP_QUAL, NULL EDI_TP_ID, NULL EDI_OUR_QUAL, NULL EDI_OUR_ID, NULL EDI_REPORTING_CODE, "
            sql &= " SOTPICK1.PICK_NO EDI_PICK_NO, EDT945T1.EDI_SHIPMENT_DATE, EDT945T1.EDI_SHIPMENT_ID, "
            sql &= " EDT945T1.EDI_ORDR_CUST_PO, EDT945T1.EDI_DIVISION_CODE, "
            sql &= " EDT945T1.EDI_BOL_NO, EDT945T1.EDI_MASTER_BOL_NO, NULL, 0, "
            sql &= " EDT945T1.EDI_ORDR_SHIP_DATE, EDT945T1.EDI_TRANS_METH_CODE, "
            sql &= " EDT945T1.EDI_CARRIER_NAME, EDT945T1.EDI_CARRIER_CODE, "
            sql &= " EDT945T1.EDI_CARRIER_SCAC_CODE, EDT945T1.EDI_FRT_TERMS, 0, 0, "
            sql &= " EDT945T1.EDI_RECEIVED_DATE, EDT945T1.COMPANY_CODE, EDT945T1.CUST_CODE, '0', "
            sql &= " EDT945T1.EDI_TRAILER_NO, EDT945T1.EDI_LOAD_ID, EDT945T1.SHIP_PICKUP_NO, EDT945T1.SHIP_AUTH_NO "
            sql &= " FROM SOTPICK1, EDT945T1, "
            sql &= " (SELECT PICK_NO, MAX(ORDR_NO_3PL) ORDR_NO_3PL FROM SOTPICKC WHERE PICK_NO = R1.PICK_NO AND ORDR_NO_3PL IS NOT NULL GROUP BY PICK_NO) SOTPICKC"
            sql &= " WHERE SOTPICK1.SHIP_BOL_NO = EDT945T1.EDI_SHIPMENT_ID"
            sql &= " AND EDT945T1.EDI_DOC_SEQ_NO = (SELECT MAX(EDI_DOC_SEQ_NO) FROM EDT945T1 WHERE EDI_SHIPMENT_ID = R1.SHIP_BOL_NO)"
            sql &= " AND SOTPICK1.PICK_NO = R1.PICK_NO"
            sql &= " AND SOTPICKC.PICK_NO = SOTPICKC.PICK_NO"
            sql &= " AND SOTPICKC.ORDR_NO_3PL IN (SELECT OHORDN FROM CONV.CFG_CNLORDS);"
            sql &= " END LOOP; END; END;"
            ASCDATA1.ExecuteSQL(sql)

            ' A Windows Service imports the data.
            Dim tempTable As String = ASCMAIN1.Temp_Table("SELECT EDI_DOC_SEQ_NO FROM EDT945T1 WHERE NVL(EDI_PROCESS_IND, '0') = '0'")

            For Each row As DataRow In ASCDATA1.GetDataTable("SELECT * FROM EDT945T1 WHERE EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM " & tempTable & ")").Rows
                tblEDT945T1.ImportRow(row)
            Next

            For Each row As DataRow In ASCDATA1.GetDataTable("SELECT * FROM EDT945T2 WHERE EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM " & tempTable & ")").Rows
                tblEDT945T2.ImportRow(row)
            Next

            'tblEDT945T1 = ASCDATA1.GetDataTable("SELECT * FROM EDT945T1 WHERE EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM " & tempTable & ")")
            'tblEDT945T2 = ASCDATA1.GetDataTable("SELECT * FROM EDT945T2 WHERE EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM " & tempTable & ")")

            Return True

        Catch ex As Exception
            lastErrorMessage = "ProcessClarinsShipConfirmationData: " & ex.Message
            Return False
        End Try

    End Function

    Public Function ProcessClarinsReturnsTransactions(ByRef tblEDTRTRN1 As DataTable, ByRef tblEDTRTRN2 As DataTable) As Boolean

        Try
            ' A Windows Service imports the data.
            Dim wkTable As String = ASCMAIN1.Temp_Table("SELECT EDI_DOC_SEQ_NO FROM EDTRTRN1 WHERE INIT_DATE >= TRUNC(SYSDATE - 1)")

            For Each row As DataRow In ASCDATA1.GetDataTable("SELECT * FROM EDTRTRN1 WHERE EDI_DOC_SEQ_NO in (SELECT EDI_DOC_SEQ_NO FROM " & wkTable & ")").Rows
                tblEDTRTRN1.ImportRow(row)
            Next

            For Each row As DataRow In ASCDATA1.GetDataTable("SELECT * FROM EDTRTRN2 WHERE EDI_DOC_SEQ_NO in (SELECT EDI_DOC_SEQ_NO FROM " & wkTable & ")").Rows
                tblEDTRTRN2.ImportRow(row)
            Next

            'tblEDTRTRN1 = ASCDATA1.GetDataTable("SELECT * FROM EDTRTRN1 WHERE EDI_DOC_SEQ_NO in (SELECT EDI_DOC_SEQ_NO FROM " & wkTable & ")")
            'tblEDTRTRN2 = ASCDATA1.GetDataTable("SELECT * FROM EDTRTRN2 WHERE EDI_DOC_SEQ_NO in (SELECT EDI_DOC_SEQ_NO FROM " & wkTable & ")")

            Return True

        Catch ex As Exception
            lastErrorMessage = "Process Clarins Returns Transactions: " & ex.Message
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Import Inventory Transactions - Receipts, Adjustments
    ''' </summary>
    ''' <param name="tblEDTTRXN1"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ProcessClarinsInventoryTransactions(ByRef tblEDTTRXN1 As DataTable) As Boolean

        Try
            ' A Windows Service imports the data.
            'tblEDTTRXN1 = ASCDATA1.GetDataTable("SELECT * FROM EDTTRXN1 WHERE IMPORT_DATE >= TRUNC(SYSDATE)")

            For Each row As DataRow In ASCDATA1.GetDataTable("SELECT * FROM EDTTRXN1 WHERE IMPORT_DATE >= TRUNC(SYSDATE - 1)").Rows
                tblEDTTRXN1.ImportRow(row)
            Next

            Return True

        Catch ex As Exception
            lastErrorMessage = "Process Inventory Transactions: " & ex.Message
            Return False
        End Try
    End Function


#End Region

#Region "nSoftware Secure ftp"

    ''' <summary>
    ''' Upload Files via sFtp
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function UploadFilesFtp(fileName As String) As Boolean

        Try
            AddTask($"SOCADSO1 Start UploadFilesFtp")
            Dim FileNameRemote As String = System.IO.Path.GetFileName(fileName)
            Return TAC.TACSCOM1.sftp_put(Nothing, clssshAppCredentials.SSH_APP_CODE, True, fileName, FileNameRemote)
            AddTask($"SOCADSO1 Finish UploadFilesFtp")

            'Sftp1 = New nsoftware.IPWorksSSH.Sftp
            'Sftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwaresftpkey")


            'Dim SSH_APP_PARTNER_URI As String = String.Empty
            'Dim SSH_APP_PARTNER_PUBKEY As String = String.Empty

            'If ASCMAIN1.DBS_COMPANY = ASCMAIN1.DBS_PASSWORD Then
            '    SSH_APP_PARTNER_URI = clssshAppCredentials.SSH_APP_PARTNER_URI_PROD
            '    SSH_APP_PARTNER_PUBKEY = clssshAppCredentials.SSH_APP_PARTNER_PUBKEY_PROD
            'Else
            '    SSH_APP_PARTNER_URI = clssshAppCredentials.SSH_APP_PARTNER_URI_TEST
            '    SSH_APP_PARTNER_PUBKEY = clssshAppCredentials.SSH_APP_PARTNER_PUBKEY_TEST
            'End If

            'Sftp1.SSHUser = clssshAppCredentials.SSH_APP_USERNAME

            'If clssshAppCredentials.SSH_APP_PASSWORD.Length > 0 Then
            '    Sftp1.SSHAuthMode = nsoftware.IPWorksSSH.SCPSSHAuthModes.amPassword
            '    Sftp1.SSHPassword = clssshAppCredentials.SSH_APP_PASSWORD
            'Else
            '    Sftp1.SSHAuthMode = nsoftware.IPWorksSSH.SCPSSHAuthModes.amPublicKey
            '    'sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "C:\Users\wjz\Desktop\Interparfums\JPMC\JPMC_SSH_pvt.ppk", "0ff1c3INT", "*")
            '    'sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "S:\INT\Archive\INT\JPMC\JPMC_SSH_pvt.ppk", "0ff1c3INT", "*")

            '    If ASCMAIN1.Running_in_VS Then
            '        Stop
            '        Sftp1.SSHCert = New nsoftware.IPWorksSSH.Certificate(nsoftware.IPWorksSSH.CertStoreTypes.cstPPKFile, "C:\VS\AHA\Archive\INT\JPMC\JPMC_IPLB_pvt.asc", "0ff1c3INT", "*")
            '        'sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "C:\Users\wjz\Desktop\Interparfums\JPMC\JPMC_SSH_pvt.ppk", "0ff1c3INT", "*")
            '    Else
            '        ' sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "S:\INT\Archive\INT\JPMC\JPMC_IPLB_pvt.asc", "0ff1c3INT", "*")
            '        'sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "S:\INT\Archive\INT\JPMC\JPMC_SSH_pvt.ppk", "0ff1c3INT", "*")
            '        Dim ssh_file As String = ASCMAIN1.Folders("SharedRoot") & "Archive\INT\JPMC\JPMC_SSH_pvt.ppk"
            '        Sftp1.SSHCert = New nsoftware.IPWorksSSH.Certificate(nsoftware.IPWorksSSH.CertStoreTypes.cstPPKFile, ssh_file, "0ff1c3INT", "*")
            '    End If

            'End If

            'If Sftp1.Connected = True Then
            '    Sftp1.SSHLogoff()
            'End If

            'Sftp1.SSHHost = SSH_APP_PARTNER_URI
            'Sftp1.SSHLogon(SSH_APP_PARTNER_URI, clssshAppCredentials.SSH_APP_PORT)
            'Sftp1.RemotePath = clssshAppCredentials.SSH_APP_FOLDER_PUT

            'If Not My.Computer.FileSystem.FileExists(fileName) Then
            '    lastErrorMessage = $"UploadFilesFtp: Cannot locate file {fileName}."
            '    Return False
            'End If

            'Dim ftpFileName As String = System.IO.Path.GetFileName(fileName)

            'Sftp1.LocalFile = fileName
            'Sftp1.RemoteFile = ftpFileName
            'Sftp1.Overwrite = True
            'Sftp1.Upload()

            'Try
            '    My.Computer.FileSystem.MoveFile(fileName, System.IO.Path.Combine(ASCMAIN1.Folders("Archive"), "ADS", ftpFileName))
            'Catch ex As Exception

            'End Try

            'Return True

        Catch ex As Exception
            lastErrorMessage = "UploadFilesFtp Error: " & ex.Message
            Return False

        Finally
            'Sftp1.SSHLogoff()
            'Sftp1.Dispose()
        End Try

    End Function

    '''' <summary>
    '''' Accept Authentication to avoid error
    '''' </summary>
    '''' <param name="sender"></param>
    '''' <param name="e"></param>
    '''' <remarks></remarks>
    'Private Sub Sftp1_OnSSHServerAuthentication(sender As Object, e As nsoftware.IPWorksSSH.SCPSSHServerAuthenticationEventArgs) Handles Sftp1.OnSSHServerAuthentication
    '    e.Accept = True
    'End Sub

    '''' <summary>
    '''' Retrieves a list of files  from the ftp site
    '''' </summary>
    '''' <param name="sender"></param>
    '''' <param name="e"></param>
    '''' <remarks></remarks>
    'Private Sub sFtp1_OnDirList(ByVal sender As System.Object, ByVal e As nsoftware.IPWorksSSH.SftpDirListEventArgs) Handles Sftp1.OnDirList
    '    Dim filename As String = e.FileName
    '    ftpFileList.Add(filename)
    'End Sub

    'Private Sub SSHServerAuthentication(ByVal sender As Object, ByVal e As nsoftware.IPWorksSSH.SCPSSHServerAuthenticationEventArgs) Handles Sftp1.OnSSHServerAuthentication
    '    e.Accept = True
    'End Sub

    'Sub SSHStatus(sender As Object, e As nsoftware.IPWorksSSH.SCPSSHStatusEventArgs) Handles Sftp1.OnSSHStatus
    '    ' MsgBox(e.Message, MsgBoxStyle.OkOnly, "SSHStatus Messages")
    '    'theLog &= e.Message & vbCrLf
    'End Sub

#End Region

#Region "Xml Processing"

    Public Class XmlValidationErrorBuilder
        Private _errors As New List(Of ValidationEventArgs)()

        Public Sub ValidationEventHandler(ByVal sender As Object, ByVal args As ValidationEventArgs)
            If args.Severity = XmlSeverityType.Error Then
                _errors.Add(args)
            End If
        End Sub

        Public Function GetErrors() As String
            If _errors.Count <> 0 Then
                Dim builder As New System.Text.StringBuilder()
                builder.Append("The following ")
                builder.Append(_errors.Count.ToString())
                builder.AppendLine(" error(s) were found while validating the XML document against the XSD:")
                For Each i As ValidationEventArgs In _errors
                    builder.Append("* ")
                    builder.AppendLine(i.Message)
                Next
                Return builder.ToString()
            Else
                Return Nothing
            End If
        End Function
    End Class

    Public Function LoadValidatedXmlDocument(xmlFilePath As String, xsdFilePath As String) As XmlDocument
        Try
            Dim doc As New XmlDocument()
            doc.Load(xmlFilePath)
            doc.Schemas.Add(Nothing, xsdFilePath)
            Dim errorBuilder As New XmlValidationErrorBuilder()
            doc.Validate(New ValidationEventHandler(AddressOf errorBuilder.ValidationEventHandler))
            Dim errorsText As String = errorBuilder.GetErrors()
            If errorsText IsNot Nothing Then
                'Throw New Exception(errorsText)
                clsXmlXsdError = errorsText
            End If
            Return doc

        Catch ex As Exception
            clsXmlXsdError = ex.Message
            Return Nothing
        End Try

    End Function

#End Region

    Private Sub AddTask(ByVal TaskDescription As String)
        tblTasks.Rows.Add({tblTasks.Rows.Count + 1, DateTime.Now, TaskDescription})
    End Sub

End Class
