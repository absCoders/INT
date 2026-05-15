Imports ABSolution
Imports System.IO

' Validates Credit Card Numbers

Public Class ASCCCARD

#Region "Class Variables"

    Public Enum NumericFunctions_CardType
        VISA = 1
        MasterCard = 2
        DinersClub = 3
        enRoute = 4
        Discover = 5
        AmericanExpress = 6
        carteBlanche = 7
        unknown = 8
    End Enum

    Public Enum CreditCardTransactionIdentifer
        CreditCardCharge = 0
        CreditCardVoid = 1
        CreditCardCreditRefundReturn = 2
        CreditCardVoidCredit = 3
        CreditCardPostAuthorization = 4
        CreditCardAuthorizationOnly = 5
        CreditCardBalanceInquiry = 6
    End Enum

    ' Converts Enum setting to Output processiing Identifer Code
    Private CreditCardTransactionIdentiferCode() As String = {"C1", "C2", "C3", "CR", "C5", "C6", "Ci"}

    Private _requestDirectory As String = String.Empty
    Private _responseDirectory As String = String.Empty

    Private _ResponseCodes As Hashtable = Nothing

    Private _approvalCodes As String() = {"Y", "P", "Q", "B", "C"}
    Private _approvalSumbittedCodes As String() = {"Y", "P", "Q", "B", "C", "S", "U"}
    Private _sumbittedCodes As String() = {"S", "U"}

    Public queuedStatusCode As String = "U"
    Public openStatusCode As String = "O"
    Public submittedStatusCode As String = "S"
    Public invalidCreditCardStatusCode As String = "I"
    Public invalidCreditCardExpDateStatusCode As String = "E"
    Public finalizedStatusCode As String = "F"
    Public cancelledStatusCode As String = "L"

#End Region

#Region "Class Constructors"

    ''' <summary>
    ''' Stardard Class Constructor
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        ' Nothing at this moment
        Me.Initialize()
    End Sub

    ''' <summary>
    ''' Parameterized Consrtuctor
    ''' </summary>
    ''' <param name="requestDirectory">Directory to place credit card payment requests</param>
    ''' <param name="responseDirectory">Directory to retrieve processed credit card transaction repsonses</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal requestDirectory As String, ByVal responseDirectory As String)
        Me.Initialize()
        _requestDirectory = requestDirectory.Trim
        _responseDirectory = responseDirectory.Trim
    End Sub

    ''' <summary>
    ''' Initialize Class variables
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Initialize()
        Dim sql As String = String.Empty

        sql = "Select * From ARTCUSTC"
        ASCDATA1.GetDataTable(sql, "ARTCISTC", -1, False, -1)

        _requestDirectory = String.Empty
        _responseDirectory = String.Empty

        _ResponseCodes = New Hashtable

        _ResponseCodes.Add("Y", "Approved")
        _ResponseCodes.Add("N", "Declined")
        _ResponseCodes.Add("P", "Approved not captured")
        _ResponseCodes.Add("Q", "Approved for VISA/MC Purchase Card")
        _ResponseCodes.Add("B", "Approved for VISA/MC Business Card")
        _ResponseCodes.Add("C", "Approved for VISA/MC Corporate Card")

        _ResponseCodes.Add("O", "Open")
        _ResponseCodes.Add("S", "Submitted")
        _ResponseCodes.Add("L", "Cancelled")
        _ResponseCodes.Add("U", "Queued")
        _ResponseCodes.Add("E", "Invalid Exp Date")
        _ResponseCodes.Add("I", "Invalid Credit Card")
        _ResponseCodes.Add("F", "Finalized")

    End Sub

#End Region

#Region "Class Properties"

    ''' <summary>
    ''' Returns String Array of Approval Response Codes
    ''' Used to calcualte charges processed and in the prcoess of getting processed
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetApprovalCodes() As String()
        Get
            Return _approvalCodes
        End Get
    End Property

    ''' <summary>
    ''' Returns String Array of Approval / Submitted Respose Codes
    ''' Used to calcualte charges processed and in the prcoess of getting processed
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetApprovalSubmittedResponseCodes() As String()
        Get
            Return _approvalSumbittedCodes
        End Get
    End Property

    ''' <summary>
    ''' Returns the Open Status Code
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetOpenStatusCode() As String
        Get
            Return openStatusCode
        End Get
    End Property

    ''' <summary>
    ''' Returns the Queued Status Code
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetQueuedStatusCode() As String
        Get
            Return queuedStatusCode
        End Get
    End Property

    ''' <summary>
    ''' Returns the Response Code Text
    ''' </summary>
    ''' <param name="ResponseCode"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>The empty string is returned for an invalid Response Code</remarks>
    Public ReadOnly Property GetResponseCodeValue(ByVal ResponseCode As Char) As String
        Get
            Try
                Return _ResponseCodes(ResponseCode).ToString
            Catch ex As Exception
                Return (String.Empty)
            End Try
        End Get
    End Property

    ''' <summary>
    ''' Returns String Array of Submitted Response Codes
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetSumbittedCodes() As String()
        Get
            Return _sumbittedCodes
        End Get
    End Property

    ''' <summary>
    ''' Get / Set directory to place credit card payment requests
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property RequestDirectory() As String
        Get
            Return _requestDirectory
        End Get
        Set(ByVal value As String)
            _requestDirectory = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Returns Hashtable of Credit Card Processing Status / Response
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ResponseCodes() As Hashtable
        Get
            Return _ResponseCodes
        End Get
    End Property

    ''' <summary>
    ''' Get / Set directory to retrieve processed credit card transaction repsonses
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ResponseDirectory() As String
        Get
            Return _responseDirectory
        End Get
        Set(ByVal value As String)
            _responseDirectory = value.Trim
        End Set
    End Property

#End Region

    ''' <summary>
    ''' Returns the Type of Credit Card based on the Credit Card Number
    ''' </summary>
    ''' <param name="CardNumber">Credit Card Number</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCreditCardType(ByVal CardNumber As String) As NumericFunctions_CardType

        GetCreditCardType = NumericFunctions_CardType.unknown

        Select Case Left$(CardNumber, 4)
            Case 2014
                GetCreditCardType = NumericFunctions_CardType.enRoute
            Case 2149
                GetCreditCardType = NumericFunctions_CardType.enRoute
            Case 3000 To 3059
                GetCreditCardType = NumericFunctions_CardType.DinersClub
            Case 3400 To 3499
                GetCreditCardType = NumericFunctions_CardType.AmericanExpress
            Case 3600 To 3699
                GetCreditCardType = NumericFunctions_CardType.DinersClub
            Case 3700 To 3799
                GetCreditCardType = NumericFunctions_CardType.AmericanExpress
            Case 3800 To 3889
                GetCreditCardType = NumericFunctions_CardType.DinersClub
            Case 3890 To 3899
                GetCreditCardType = NumericFunctions_CardType.carteBlanche
            Case 4000 To 4999
                GetCreditCardType = NumericFunctions_CardType.VISA
            Case 5100 To 5599
                GetCreditCardType = NumericFunctions_CardType.MasterCard
            Case 6011
                GetCreditCardType = NumericFunctions_CardType.Discover
        End Select

    End Function

    ''' <summary>
    ''' Returns the masked representation of a credit card number (last four)
    ''' </summary>
    ''' <param name="CreditCardNumber"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCreditCardMask(ByVal CreditCardNumber As String) As String

        GetCreditCardMask = String.Empty

        CreditCardNumber = CreditCardNumber.Trim

        Select Case Me.GetCreditCardType(CreditCardNumber)
            Case NumericFunctions_CardType.unknown
                Return String.Empty
            Case NumericFunctions_CardType.AmericanExpress
                Return "xxxx-xxxxxx-" & CreditCardNumber.Substring(10, 5)
            Case Else
                Return "xxxx-xxxx-xxxx-" & CreditCardNumber.Substring(CreditCardNumber.Length - 4, 4)
        End Select

    End Function

    ''' <summary>
    ''' Validates a Credit Card Number
    ''' </summary>
    ''' <param name="CreditCardNumber">Credit Card Number</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsCreditCardValid(ByVal CreditCardNumber As String) As Boolean

        Dim CCType As NumericFunctions_CardType = Me.GetCreditCardType(CreditCardNumber)

        If CCType = NumericFunctions_CardType.unknown Then
            Return False
        End If

        Return IsCreditCardValid(CreditCardNumber, CCType)

    End Function

    ''' <summary>
    ''' Validates a Credit Card Number
    ''' </summary>
    ''' <param name="CreditCardNumber">Credit Card Number</param>
    ''' <param name="CCType">Credit Card Type</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsCreditCardValid(ByVal CreditCardNumber As String, ByVal CCType As NumericFunctions_CardType) As Boolean

        Dim X As Long = 0
        Dim Total As Long = 0
        Dim CardNumberLength As Long = 0
        Dim Temp As String = String.Empty
        Dim RevCardNumber As String = String.Empty
        Dim ParsedCardNumber As String = String.Empty
        Dim Character As String = String.Empty
        Dim Digit As Long = 0
        Dim FirstDigit As Long = 0
        Dim FirstFourDigits As Long = 0
        Dim DigitsOK As Boolean = False
        Dim LengthOK As Boolean = False

        'LUHN Formula (in quasi-LaymanTerms)
        '1. Reverse the card number
        '2. Starting with the FIRST DIGIT of the reversed
        '    card number, write the value of that digit
        '    on a piece of paper.
        '    Do the same for the 3rd digit, 5th, 7th, etc.....
        '    writing each value after the previous one
        '    (in essence creating a string of numbers)
        '3. Then, starting at the SECOND DIGIT of the
        '    reversed card number, DOUBLE the value of
        '    that digit; append the doubled value to the
        '    string of numbers from Step 2.
        '    Do the same for the 4th digit, 6th, 8th, etc.....
        '4. Add all the numbers from your string of numbers together.
        '    If this total ENDS IN ZERO, the card number is valid.

        IsCreditCardValid = False

        Temp = String.Empty
        'Parses all non-numeric characters from
        'CardNumber into a temporary string
        For X = 0 To CreditCardNumber.Length - 1
            Character = CreditCardNumber.Substring(X, 1)
            Temp = Temp & IIf(IsNumeric(Character), Character, "")
        Next
        ParsedCardNumber = Temp

        'Save the length of the card number length for later use
        'in editting the card carrier
        CardNumberLength = Temp.Length

        'Exit if number of digits is not in range of testable Card numbers
        If CardNumberLength < 9 Or CardNumberLength > 19 Then
            Exit Function
        End If

        'Save specific leading digits for later use in editting the card carrier
        FirstDigit = Long.Parse(Temp.Substring(0, 1))
        FirstFourDigits = Long.Parse(Temp.Substring(0, 4))

        Select Case CCType
            Case NumericFunctions_CardType.VISA
                DigitsOK = (FirstFourDigits >= 4000) And _
                                (FirstFourDigits <= 4999)
                LengthOK = (CardNumberLength = 13) Or _
                                (CardNumberLength = 16)

            Case NumericFunctions_CardType.MasterCard
                DigitsOK = (FirstFourDigits >= 5100) And _
                                (FirstFourDigits <= 5599)
                LengthOK = (CardNumberLength = 16)

            Case NumericFunctions_CardType.enRoute
                DigitsOK = (FirstFourDigits = 2014) Or _
                                (FirstFourDigits = 2149)
                LengthOK = (CardNumberLength = 15)

            Case NumericFunctions_CardType.Discover
                DigitsOK = (FirstFourDigits = 6011)
                LengthOK = (CardNumberLength = 16)

            Case NumericFunctions_CardType.DinersClub
                DigitsOK = ((FirstFourDigits >= 3000) And _
                                (FirstFourDigits <= 3059)) Or _
                                ((FirstFourDigits >= 3600) And _
                                (FirstFourDigits <= 3699)) Or _
                                ((FirstFourDigits >= 3800) And _
                                (FirstFourDigits <= 3889))
                LengthOK = (CardNumberLength = 14)

            Case NumericFunctions_CardType.carteBlanche
                DigitsOK = (FirstFourDigits >= 3890) And _
                                (FirstFourDigits <= 3899)
                LengthOK = (CardNumberLength = 14)

            Case NumericFunctions_CardType.AmericanExpress
                DigitsOK = ((FirstFourDigits >= 3400) And _
                                (FirstFourDigits <= 3499)) Or _
                                ((FirstFourDigits >= 3700) And _
                                (FirstFourDigits <= 3799))
                LengthOK = (CardNumberLength = 15)

            Case Else
                DigitsOK = False
                LengthOK = False
        End Select

        'If the number sequence and/or card number length
        'do not match the requirements the
        'requirements of the carrier; Exit
        If Not DigitsOK Or Not LengthOK Then
            Exit Function
        End If

        'Reverse the CardNumber
        RevCardNumber = String.Empty
        For X = 0 To Temp.Length - 1
            RevCardNumber = Temp.Substring(X, 1) & RevCardNumber
        Next

        'Iterate through the reversed number; Add the
        'calculated totals to string for subsequent summation.
        'The equation  = (1 + (X - 1) Mod 2) below will result '
        'in a multiplier of 1 if the element is odd (i.e. 1st digit, 3rd, 5th, etc)
        'or a multiplier of 2 if the element is even (i.e. 2nd digit, 4th, 6th, etc)
        Temp = String.Empty
        For X = 1 To RevCardNumber.Length
            Digit = Long.Parse(RevCardNumber.Substring(X - 1, 1))
            Temp = Temp & (Digit * (1 + (X - 1) Mod 2))
        Next

        'Iterate through the string just created and total
        'the individual numbers in each character of the string
        For X = 1 To Temp.Length
            Total = Total + Long.Parse(Temp.Substring(X - 1, 1))
        Next

        'If the total of the added elements ends in zero, '
        'the card number is valid
        IsCreditCardValid = Boolean.Parse(Total Mod 10 = 0)

    End Function

    ''' <summary>
    ''' Verifes if the Expriration of a credit card is Valid
    ''' </summary>
    ''' <param name="expDate">Expiration date in the format MMYY</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsCardExpirationDateValid(ByVal expDate As String) As Boolean

        ' Verify the Exp Date is Valid
        If expDate.Length <> 4 Then
            Return False
        Else
            expDate = expDate.Substring(0, 2) & "/01/" & expDate.Substring(2, 2)
            Try
                If Integer.Parse(DateTime.Parse(expDate).ToString("yyyyMM")) < Integer.Parse(DateTime.Now.ToString("yyyyMM")) Then
                    Return False
                End If
            Catch ex As Exception
                Return False
            End Try
        End If

        Return True

    End Function

    ''' <summary>
    ''' Retrieves Credit Card Processing Information and updates Credit Card Sunmission Status
    ''' Results should be Approved or Declined.
    ''' </summary>
    ''' <param name="tblARTCCPA1">Header Data table of Credit Card Payments</param>
    ''' <returns>Return True if all Icverify where read in; otherwise, false</returns>
    ''' <remarks>Fuction looks for records in the two tables supplied. If not found, it adds them.</remarks>
    Public Function ImportCreditCardResponse(ByRef tblARTCCPA1 As DataTable) As Boolean

        Dim processedImportWithErrors As Boolean = False

        Dim sql As String = String.Empty
        Dim processingWithNoErrors As Boolean = True

        Dim responseData As String = String.Empty
        Dim responseReader As System.IO.StreamReader = Nothing
        Dim responseFileName As String = String.Empty
        Dim responseDirectoryArchive As String = String.Empty

        Dim CCPA_NO As String = String.Empty
        Dim rowARTCCPA1 As DataRow = Nothing

        Dim rowWk As DataRow = Nothing
        Dim tblWk As DataTable = Nothing

        Dim wkString As String = String.Empty
        Dim fieldName As String = String.Empty

        Dim responseCode As String = String.Empty
        Dim dateAuthorized As Date = Nothing
        Dim approvalDeclineMessage As String = String.Empty

        If Me._responseDirectory.Length = 0 Then
            Return False
        End If

        If Not My.Computer.FileSystem.DirectoryExists(_responseDirectory) Then
            Return False
        End If

        responseDirectoryArchive = _responseDirectory & "\Archive"
        If Not My.Computer.FileSystem.DirectoryExists(responseDirectoryArchive) Then
            Return False
        End If

        ' The Filename tells us what Credit Card Request the response is for
        ' Example 0000000002 - CCPA_NO = "0000000002" 
        For Each responseFile As String In My.Computer.FileSystem.GetFiles(_responseDirectory, FileIO.SearchOption.SearchTopLevelOnly, "*.ans")
            responseReader = New StreamReader(responseFile)

            responseFileName = My.Computer.FileSystem.GetName(responseFile)

            dateAuthorized = My.Computer.FileSystem.GetFileInfo(responseFile).CreationTime

            wkString = responseFileName
            wkString = wkString.ToUpper
            wkString = wkString.Replace(".ANS", String.Empty)
            CCPA_NO = wkString.Split("_")(0)

            ' If the Credit Card Payment Header information is not in the table add it; otherwise get it
            sql = "CCPA_NO = '" & CCPA_NO & "'"
            If tblARTCCPA1.Select(sql).Length = 0 Then
                sql = "SELECT * FROM ARTCCPA1 WHERE " & sql
                tblWk = ASCDATA1.GetDataTable(sql)
                If tblWk Is Nothing Then
                    processedImportWithErrors = True
                    Continue For
                End If

                If tblWk.Rows.Count = 0 Then
                    processedImportWithErrors = True
                    Continue For
                End If

                For Each rowWk In tblWk.Rows
                    rowARTCCPA1 = tblARTCCPA1.NewRow

                    For Each dataField As DataColumn In tblWk.Columns
                        fieldName = dataField.ColumnName
                        If rowARTCCPA1.Table.Columns.Contains(fieldName) Then
                            rowARTCCPA1.Item(fieldName) = rowWk.Item(fieldName)
                        End If
                    Next
                    tblARTCCPA1.Rows.Add(rowARTCCPA1)
                    rowARTCCPA1.AcceptChanges()
                    rowARTCCPA1.SetModified()
                Next
            Else
                rowARTCCPA1 = tblARTCCPA1.Select(sql)(0)
            End If

            If tblARTCCPA1 Is Nothing Then
                processedImportWithErrors = True
                Continue For
            End If

            If tblARTCCPA1.Rows.Count = 0 Then
                processedImportWithErrors = True
                Continue For
            End If

            Try
                rowARTCCPA1 = tblARTCCPA1.Select("CCPA_NO = '" & CCPA_NO & "'")(0)
            Catch ex As Exception
                processedImportWithErrors = True
                Continue For
            End Try

            Try
                While responseReader.Peek <> -1
                    responseData = responseReader.ReadLine()

                    approvalDeclineMessage = String.Empty
                    responseCode = responseData.Substring(0, 1)
                    Select Case responseCode
                        Case "N"
                            approvalDeclineMessage = responseData.Substring(1)
                        Case "Y"
                            approvalDeclineMessage = responseData.Substring(1, 6)
                        Case Else
                            approvalDeclineMessage = responseData.Substring(1, 6)
                    End Select

                    approvalDeclineMessage = approvalDeclineMessage.Trim
                    If approvalDeclineMessage.Length > 25 Then
                        approvalDeclineMessage = approvalDeclineMessage.Substring(0, 25)
                    End If

                    rowARTCCPA1.Item("CCPA_STATUS") = responseCode
                    rowARTCCPA1.Item("CCPA_AUTH") = approvalDeclineMessage
                    rowARTCCPA1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                    rowARTCCPA1.Item("LAST_DATE") = DateTime.Now
                    rowARTCCPA1.Item("CCPA_DATE_SALE") = dateAuthorized

                End While

                responseReader.Close()
                My.Computer.FileSystem.MoveFile(responseFile, responseDirectoryArchive & "\" & responseFileName)

            Catch ex As Exception
                processedImportWithErrors = True
            End Try
        Next

        Return (Not processedImportWithErrors)

    End Function

    ''' <summary>
    ''' Sets the Credit Card Paymewnt Request to status queued.
    ''' If the credit card info is missing it fills in the Default Credit Card Information for a Customer Credit Card Payment
    ''' Set the Records Status to one of the following: 
    '''      Queued, Invalid Credit Card Number, Invalid Credit Card Exp Date
    ''' </summary>
    ''' <param name="rowARTCCPA1">Data row to full with credit card information</param>
    ''' <returns>True is credit card is valid; otherwise, false </returns>
    ''' <remarks></remarks>
    Public Function QueueCreditCardForProcessing(ByRef rowARTCCPA1 As DataRow) As Boolean

        Dim tblARTCUSTC As DataTable = Nothing
        Dim rowARTCUSTC As DataRow = Nothing

        Dim CUST_CODE As String = String.Empty
        Dim sql As String = String.Empty

        Dim creditCardNo As String = String.Empty
        Dim creditCardExpDate As String = String.Empty
        Dim creditCardError As String = String.Empty

        Dim cardType As NumericFunctions_CardType = NumericFunctions_CardType.unknown

        Dim validCreditCardFound As String = False

        creditCardError = Me.queuedStatusCode
        validCreditCardFound = False

        Try
            If (rowARTCCPA1.Item("CUST_CREDIT_CARD_NO") & String.Empty).ToString.Trim <> String.Empty Then
                creditCardNo = (rowARTCCPA1.Item("CUST_CREDIT_CARD_NO") & String.Empty).ToString.Trim
                creditCardExpDate = (rowARTCCPA1.Item("CUST_CREDIT_CARD_EXP_DATE") & String.Empty).ToString.Trim
            End If
        Catch ex As Exception
            creditCardNo = String.Empty
        End Try

        If creditCardNo.Length = 0 Then
            ' Get Card on file for the customer
            CUST_CODE = rowARTCCPA1.Item("CUST_CODE")
            sql = "SELECT * FROM ARTCUSTC WHERE CUST_CODE = '" & CUST_CODE & "' AND CUST_CREDIT_CARD_STATUS = 'A' ORDER BY CUST_CREDIT_CARD_EXP_DATE DESC"

            tblARTCUSTC = ASCDATA1.GetDataTable(sql, "ARTCUSTC")

            If tblARTCUSTC Is Nothing Then
                rowARTCCPA1.Item("CCPA_STATUS") = Me.invalidCreditCardStatusCode
                Return False
            End If

            If tblARTCUSTC.Rows.Count = 0 Then
                rowARTCCPA1.Item("CCPA_STATUS") = Me.invalidCreditCardStatusCode
                Return False
            End If

            For Each rowARTCUSTC In tblARTCUSTC.Rows
                creditCardNo = rowARTCUSTC.Item("CUST_CREDIT_CARD_NO") & String.Empty
                creditCardExpDate = rowARTCUSTC.Item("CUST_CREDIT_CARD_EXP_DATE") & String.Empty

                creditCardError = Me.queuedStatusCode

                cardType = Me.GetCreditCardType(creditCardNo)
                If cardType = NumericFunctions_CardType.unknown Then
                    creditCardError = Me.invalidCreditCardStatusCode
                    Continue For
                End If

                If Not Me.IsCreditCardValid(creditCardNo, cardType) Then
                    creditCardError = Me.invalidCreditCardStatusCode
                    Continue For
                End If

                If Not Me.IsCardExpirationDateValid(creditCardExpDate) Then
                    creditCardError = Me.invalidCreditCardExpDateStatusCode
                    Continue For
                End If

                validCreditCardFound = True
                Exit For
            Next

        Else
            ' Use the card data provided in the record
            cardType = Me.GetCreditCardType(creditCardNo)
            If cardType = NumericFunctions_CardType.unknown Then
                creditCardError = Me.invalidCreditCardStatusCode
            ElseIf Not Me.IsCreditCardValid(creditCardNo, cardType) Then
                creditCardError = Me.invalidCreditCardStatusCode
            ElseIf Not Me.IsCardExpirationDateValid(creditCardExpDate) Then
                creditCardError = Me.invalidCreditCardExpDateStatusCode
            End If
            ' Always set to false to preserve original credit card data
            validCreditCardFound = False
        End If

        rowARTCCPA1.Item("CCPA_STATUS") = creditCardError
        If validCreditCardFound = True Then
            rowARTCCPA1.Item("CUST_CREDIT_CARD_NO") = rowARTCUSTC.Item("CUST_CREDIT_CARD_NO") & String.Empty
            rowARTCCPA1.Item("CUST_CREDIT_CARD_EXP_DATE") = rowARTCUSTC.Item("CUST_CREDIT_CARD_EXP_DATE") & String.Empty
            rowARTCCPA1.Item("CUST_CREDIT_CARD_VER_CODE") = rowARTCUSTC.Item("CUST_CREDIT_CARD_VER_CODE") & String.Empty
            rowARTCCPA1.Item("CUST_CREDIT_CARD_NAME") = rowARTCUSTC.Item("CUST_CREDIT_CARD_NAME") & String.Empty
            rowARTCCPA1.Item("CUST_CREDIT_CARD_ADDR1") = rowARTCUSTC.Item("CUST_CREDIT_CARD_ADDR1") & String.Empty
            rowARTCCPA1.Item("CUST_CREDIT_CARD_CITY") = rowARTCUSTC.Item("CUST_CREDIT_CARD_CITY") & String.Empty
            rowARTCCPA1.Item("CUST_CREDIT_CARD_STATE") = rowARTCUSTC.Item("CUST_CREDIT_CARD_STATE") & String.Empty
            rowARTCCPA1.Item("CUST_CREDIT_CARD_ZIP_CODE") = rowARTCUSTC.Item("CUST_CREDIT_CARD_ZIP_CODE") & String.Empty
        End If

        Return creditCardError = queuedStatusCode
    End Function

    ''' <summary>
    ''' Sends Queued Credit Card Payments to the Credit Card Processor for processing
    ''' </summary>
    ''' <param name="PaymentNumber">Unique Payment Number for Payment Header Record</param>
    ''' <returns>True if successful, otherwise false</returns>
    ''' <remarks>Only processes payment with a Status of 'Q' queued for processing</remarks>
    Public Function SubmitCreditCardForProcessing(ByVal PaymentNumber As String, ByVal ProcessingType As CreditCardTransactionIdentifer) As Boolean
        Dim tblARTCCPA1 As DataTable = Nothing

        PaymentNumber = PaymentNumber.Trim

        SubmitCreditCardForProcessing = False

        Try
            tblARTCCPA1 = ASCDATA1.GetDataTable("SELECT * FROM ARTCCPA1 WHERE CCPA_NO = '" & PaymentNumber & "'")

            If tblARTCCPA1.Rows.Count = 0 Then Exit Function

        Catch ex As Exception
            Exit Function
        End Try

        Return SubmitCreditCardForProcessing(tblARTCCPA1, ProcessingType)

    End Function

    ''' <summary>
    ''' Sends Queued Credit Card Payments to the Credit Card Processor for processing
    ''' </summary>
    ''' <param name="tblARTCCPA1">Data table containing credit card header payment records</param>
    ''' <returns>True if successful, otherwise false</returns>
    ''' <remarks>Only processes payment with a Status of 'Q' queued for processing</remarks>
    Public Function SubmitCreditCardForProcessing(ByRef tblARTCCPA1 As DataTable, ByVal ProcessingType As CreditCardTransactionIdentifer) As Boolean

        Dim CCPA_NO As String = String.Empty
        Dim sql As String = String.Empty
        Dim creditCardNo As String = String.Empty
        Dim creditCardExpDate As String = String.Empty
        Dim cardType As NumericFunctions_CardType = NumericFunctions_CardType.unknown
        Dim transactionFile As String = String.Empty
        Dim wkString As String = String.Empty
        Dim processingWithNoErrors As Boolean = True

        Dim streamWriter As System.IO.StreamWriter = Nothing

        If _requestDirectory.Length = 0 Then
            Return False
        End If

        If Not My.Computer.FileSystem.DirectoryExists(_requestDirectory) Then
            Return False
        End If

        Dim processingIdentifer As String = CreditCardTransactionIdentiferCode(ProcessingType)

        For Each rowARTCCPA1 As DataRow In tblARTCCPA1.Select("CCPA_STATUS = '" & queuedStatusCode & "'")
            CCPA_NO = rowARTCCPA1.Item("CCPA_NO") & String.Empty

            creditCardNo = rowARTCCPA1.Item("CUST_CREDIT_CARD_NO") & String.Empty
            creditCardExpDate = rowARTCCPA1.Item("CUST_CREDIT_CARD_EXP_DATE") & String.Empty

            cardType = Me.GetCreditCardType(creditCardNo)
            If cardType = NumericFunctions_CardType.unknown Then
                rowARTCCPA1.Item("CCPA_STATUS") = Me.invalidCreditCardStatusCode
                processingWithNoErrors = False
                Continue For
            End If

            If Me.IsCreditCardValid(creditCardNo, cardType) = False Then
                rowARTCCPA1.Item("CCPA_STATUS") = Me.invalidCreditCardStatusCode
                processingWithNoErrors = False
                Continue For
            End If

            If Me.IsCardExpirationDateValid(creditCardExpDate) = False Then
                rowARTCCPA1.Item("CCPA_STATUS") = Me.invalidCreditCardExpDateStatusCode
                processingWithNoErrors = False
                Continue For
            End If

            transactionFile = rowARTCCPA1.Item("CCPA_NO") & ".req"
            transactionFile = _requestDirectory & "\" & transactionFile

            streamWriter = New System.IO.StreamWriter(transactionFile, False)
            ' Transaction Identifer
            streamWriter.Write(ControlChars.Quote & processingIdentifer & ControlChars.Quote & ",")
            ' Cletrk Information
            streamWriter.Write(ControlChars.Quote & rowARTCCPA1.Item("INIT_OPER") & ControlChars.Quote & ",")
            ' Comment Information (Payment Number and Line Number)
            streamWriter.Write(ControlChars.Quote & rowARTCCPA1.Item("CCPA_NO") & ControlChars.Quote & ",")
            ' Creditr Card Number
            streamWriter.Write(ControlChars.Quote & creditCardNo & ControlChars.Quote & ",")
            ' Credit Card Expiratio Date (YYMM)
            streamWriter.Write(ControlChars.Quote & creditCardExpDate.Substring(2, 2) & creditCardExpDate.Substring(0, 2) & ControlChars.Quote & ",")
            ' Amount to process
            streamWriter.Write(ControlChars.Quote & Double.Parse(rowARTCCPA1.Item("CCPA_AMT")).ToString("0.00") & ControlChars.Quote & ",")

            ' Credit Card Billing Zip Code 5 or 9 digit zip code without spaces or dashes
            wkString = rowARTCCPA1.Item("CUST_CREDIT_CARD_ZIP_CODE")
            For i As Integer = wkString.Length - 1 To 0 Step -1
                If Not Char.IsDigit(wkString.Substring(i, 1)) Then
                    wkString = wkString.Remove(i, 1)
                End If
            Next
            wkString = wkString.Trim
            If Not (wkString.Length = 5 Or wkString.Length = 9) Then
                wkString = String.Empty
            End If
            streamWriter.Write(ControlChars.Quote & wkString & ControlChars.Quote & ",")

            ' Credit Card Billiing Address - Up to 32 alphanumeric characters and allows for spaces and dashes
            wkString = rowARTCCPA1.Item("CUST_CREDIT_CARD_ADDR1")
            For i As Integer = wkString.Length - 1 To 0 Step -1
                If Char.IsLetterOrDigit(wkString.Substring(i, 1)) Then
                    Continue For
                End If

                If wkString.Substring(i, 1) = "-" Or wkString.Substring(i, 1) = " " Then
                    Continue For
                End If

                wkString = wkString.Remove(i, 1)
            Next
            wkString = wkString.Trim
            If wkString.Length > 32 Then
                wkString = wkString.Substring(0, 32)
            End If
            streamWriter.Write(ControlChars.Quote & wkString & ControlChars.Quote)

            ' Line feed
            streamWriter.WriteLine()
            streamWriter.Close()

            rowARTCCPA1.Item("CCPA_STATUS") = Me.submittedStatusCode
            rowARTCCPA1.Item("CCPA_DATE_AUTH") = DateTime.Now
        Next

        Return processingWithNoErrors
    End Function

End Class
