Public Class eCheck

    Public rowGLTBANK1 As DataRow
    Public tblAPTCHCK1 As New DataTable
    Public tblAPTCHCK4 As New DataTable
    Public tblAPTVEND1 As New DataTable
    Public tblAPTVEND2 As New DataTable

    Private eCheckResponse As New eCheckResponseClass
    Private processingTable As Boolean = False
    Public errorList As List(Of String)
    Private BANK_CODE As String = ""
    Public TRAN_NO As String = ""
    Public FOLDERNAME As String = ""
    Public INIT_OPER As String
    Public INIT_DATE As Date

    'CREATE TABLE APTCHCK4 (
    'BANK_CODE                   VARCHAR2(6),
    'CHECK_NUM                   VARCHAR2(10),
    'TRAN_NO                   VARCHAR2(10),
    'INIT_DATE                      DATE,
    'LAST_DATE                      DATE,
    'INIT_OPER                      VARCHAR2(20),
    'LAST_OPER                      VARCHAR2(20),
    'RESPONSE_APPROVAL_CODE         VARCHAR2(10),
    'RESPONSE_CODE                  VARCHAR2(1),
    'RESPONSE_INV_NO                VARCHAR2(20),
    'RESPONSE_TEXT                  VARCHAR2(20),
    'RESPONSE_TRANS_ID              VARCHAR2(60),
    'PRIMARY KEY (BANK_CODE, CHECK_NUM, TRAN_NO));

    'ALTER TABLE APTVEND1 ADD VEND_BANK_ACCT_ID              VARCHAR2(30);
    'ALTER TABLE APTVEND1 ADD VEND_BANK_ROUTING_NO           VARCHAR2(10);
    'ALTER TABLE APTVEND1 ADD VEND_BANK_ACCT_CLASS           VARCHAR2(1);
    'ALTER TABLE APTVEND1 ADD VEND_BANK_ACCT_TYPE            VARCHAR2(1);

    'ALTER TABLE APTVEND2 ADD VEND_ALT_BANK_ACCT_ID              VARCHAR2(30);
    'ALTER TABLE APTVEND2 ADD VEND_ALT_BANK_ROUTING_NO           VARCHAR2(10);
    'ALTER TABLE APTVEND2 ADD VEND_ALT_BANK_ACCT_CLASS           VARCHAR2(1);
    'ALTER TABLE APTVEND2 ADD VEND_ALT_BANK_ACCT_TYPE            VARCHAR2(1);

    'ALTER TABLE GLTBANK1 ADD BANK_MERCHANT_ID               VARCHAR2(10);
    'ALTER TABLE GLTBANK1 ADD BANK_MERCHANT_PASSWORD         VARCHAR2(50);
    'ALTER TABLE GLTBANK1 ADD BANK_PP_IND                    VARCHAR2(1);

    'SELECT * FROM ASTAESC1

    'TABLE_NA COLUMN_NAME                   
    '-------- ------------------------------
    'ASTPARM1 AS_PARM_EFAX_PASSWORD         
    'ASTPARM1 AS_PARM_EMAIL_PASSWORD        
    'ASTUSER1 USER_PASSWORD  

    'INSERT INTO ASTAESC1 VALUES ('APTVEND1','VEND_BANK_ACCT_ID');
    'INSERT INTO ASTAESC1 VALUES ('APTVEND2','VEND_ALT_BANK_ACCT_ID');
    'INSERT INTO ASTAESC1 VALUES ('GLTBANK1','BANK_MERCHANT_PASSWORD');

    'INSERT INTO ASTCODE1 VALUES ('APTVEND1','VEND_PYMT_METHOD','ECHECK','eCheck Payment');
    'INSERT INTO ASTCODE1 VALUES ('APTINVH1','INV_PYMT_METHOD','ECHECK','eCheck Payment');
    'INSERT INTO ASTCODE1 VALUES ('APTCHCK1','PYMT_METHOD','ECHECK','eCheck Payment');
    'INSERT INTO ASTCODE1 VALUES ('GLTBANK1','BANK_PYMT_METHOD','ECHECK','eCheck Payment');

    ' UPDATE GLTBANK1 SET BANK_PP_IND = '1' WHERE SSH_APP_CODE IS NOT NULL;
    'ALTER TABLE APTCHCK1 ADD ECK_PAY_STATUS_IND             VARCHAR2(1);

    Public Sub New()


    End Sub

    Public Enum eCheckTypes
        Authorize
        Credit
    End Enum

    Public Class eCheckResponseClass
        Public ApprovalCode As String = String.Empty
        Public ResponseCode As String = String.Empty
        Public ResponseInvoice As String = String.Empty
        Public ResponseText As String = String.Empty
        Public ResponseTransid As String = String.Empty
        Public ErrorMessage As String = String.Empty
    End Class

    Public ReadOnly Property CheckResponse
        Get
            Return eCheckResponse
        End Get
    End Property

    Public Function Send_eChecks(ByRef rowAPTCHCK1 As DataRow, ByVal method As eCheckTypes) As Boolean
        Send_eChecks = False
        Try
            eCheckResponse = New eCheckResponseClass

            If Val(rowAPTCHCK1.Item("CHECK_AMT") & String.Empty) = 0 Then
                Return True
            End If

            Dim VEND_CODE As String = rowAPTCHCK1.Item("VEND_CODE") & String.Empty
            Dim rowAPTVEND1 As DataRow = tblAPTVEND1.Rows.Find(VEND_CODE)
            Dim VEND_BANK_ACCT_CLASS As String = rowAPTVEND1.Item("VEND_BANK_ACCT_CLASS") & ""
            Dim VEND_BANK_ACCT_TYPE As String = rowAPTVEND1.Item("VEND_BANK_ACCT_TYPE") & ""

            If eCheckResponse.ErrorMessage.Length > 0 Then
                Return False
            End If

            'Dim BANK_CODE As String = rowAPTCHCK1.Item("BANK_CODE") & String.Empty

            Dim epay As New DPayments.InPay.Echeck
            Stop
            ' epay.RuntimeLicense = "XDPNF-AANXR-FAEPA-00565-FBFTR"
            'epay.RuntimeLicense = "XDXNF-AANXR-F4MMB-5741G-0V357"
            epay.RuntimeLicense = "44504E4641414E5852464145504130303536000000000000000000000000000000000000000000004E484D5934553142000058314B365943314434424B530000"
            'epay.LicenseNumber = "XDPNF-AANXR-FAEPA-00565-FBFTR"
            Dim bankInfo As New DPayments.InPay.EPBank(
                routingNumber:=rowAPTVEND1.Item("VEND_BANK_ROUTING_NO"),
                accountNumber:=rowAPTVEND1.Item("VEND_BANK_ACCT_ID"),
                accountClass:=If(VEND_BANK_ACCT_CLASS = "B", DPayments.InPay.AccountClass.acBusiness, DPayments.InPay.AccountClass.acPersonal),
                accountType:=If(VEND_BANK_ACCT_TYPE = "C", DPayments.InPay.AccountTypes.atChecking, DPayments.InPay.AccountTypes.atSavings))
            epay.Bank = bankInfo

            Dim Customer As New DPayments.InPay.EPCustomer
            With Customer
                .FirstName = rowAPTVEND1.Item("VEND_CODE") & String.Empty
                .LastName = rowAPTVEND1.Item("VEND_NAME") & String.Empty
            End With
            epay.Customer = Customer
            epay.MerchantLogin = rowGLTBANK1.Item("BANK_MERCHANT_ID") & String.Empty
            epay.MerchantPassword = rowGLTBANK1.Item("BANK_MERCHANT_PASSWORD") & String.Empty
            Stop ' NEED TO DECRYPT
            epay.MerchantPassword = "J6oGwxW51Ca04"
            epay.CheckNumber = rowAPTCHCK1.Item("CHECK_NUM") & String.Empty
            epay.CompanyName = rowGLTBANK1.Item("ACCT_NAME") & ""
            epay.TransactionAmount = Format(Val(rowAPTCHCK1.Item("CHECK_AMT") & String.Empty), "#.00")
            epay.TransactionDesc = "Payment"

            epay.TransactionId = TRAN_NO
            epay.PaymentType = DPayments.InPay.EcheckPaymentTypes.ptPPD
            epay.Gateway = DPayments.InPay.EcheckGateways.ecgwForte

            epay.InvoiceNumber = "T" & TRAN_NO

            Select Case method
                Case eCheckTypes.Authorize
                    epay.Authorize()
                Case eCheckTypes.Credit
                    epay.Credit("", epay.TransactionAmount)
            End Select

            With eCheckResponse
                .ApprovalCode = epay.Response.ApprovalCode
                .ResponseCode = epay.Response.Code
                .ResponseInvoice = epay.Response.InvoiceNumber
                .ResponseText = epay.Response.Text
                .ResponseTransid = epay.Response.TransactionId

                Stop ' NEED TO SAVE THIS FOR EACH CHECK

                Dim audit As String = .ApprovalCode & vbCrLf & .ResponseCode & vbCrLf & .ResponseInvoice & vbCrLf & .ResponseText & vbCrLf & .ResponseTransid

                Dim FILENAME As String = FOLDERNAME & rowAPTCHCK1.Item("BANK_CODE") & "_" & rowAPTCHCK1.Item("CHECK_NUM") & "_" & TRAN_NO & ".txt"

                Using sw As New System.IO.StreamWriter(FILENAME)
                    sw.Write(audit)
                End Using

                Dim rowAPTCHCK4 As DataRow = tblAPTCHCK4.NewRow
                With rowAPTCHCK4
                    .Item("BANK_CODE") = rowAPTCHCK1.Item("BANK_CODE") & String.Empty
                    .Item("CHECK_NUM") = rowAPTCHCK1.Item("CHECK_NUM") & String.Empty
                    .Item("TRAN_NO") = TRAN_NO
                    .Item("INIT_DATE") = INIT_DATE
                    .Item("INIT_OPER") = INIT_OPER
                    .Item("RESPONSE_APPROVAL_CODE") = eCheckResponse.ApprovalCode
                    .Item("RESPONSE_CODE") = eCheckResponse.ResponseCode
                    .Item("RESPONSE_INV_NO") = eCheckResponse.ResponseInvoice
                    .Item("RESPONSE_TEXT") = eCheckResponse.ResponseText
                    .Item("RESPONSE_TRANS_ID") = eCheckResponse.ResponseTransid
                End With
                tblAPTCHCK4.Rows.Add(rowAPTCHCK4)

            End With

            Send_eChecks = True

        Catch ex As Exception
            eCheckResponse.ErrorMessage = ex.Message
        End Try

    End Function

    Public Function ValidatePayees() As List(Of String)

        errorList = New List(Of String)

        Dim BANK_CODEs As New List(Of String)
        For Each rowAPTCHCK1 As DataRow In tblAPTCHCK1.Select("", "BANK_CODE")
            Dim BANK_CODE As String = rowAPTCHCK1.Item("BANK_CODE")
            If Not BANK_CODEs.Contains(BANK_CODE) Then
                BANK_CODEs.Add(BANK_CODE)
                ValidateBank(BANK_CODE)
            End If
        Next

        If errorList.Count > 0 Then Return errorList

        If BANK_CODEs.Count <> 1 Then
            errorList.Add($"Single Bank Expected - {BANK_CODEs.Count} Banks encountered")
            Return errorList
        End If

        Dim VEND_CODEs As New List(Of String)
        For Each rowAPTCHCK1 As DataRow In tblAPTCHCK1.Select("", "VEND_CODE")
            Dim VEND_CODE As String = rowAPTCHCK1.Item("VEND_CODE")
            Dim VEND_ALT_CODE As String = rowAPTCHCK1.Item("VEND_ALT_CODE") & ""
            If VEND_CODEs.Contains(VEND_CODE & ":" & VEND_ALT_CODE) Then
                ValidatePayee(VEND_CODE, VEND_ALT_CODE)
            End If
        Next

        Return errorList
    End Function

    Public Function ValidatePayee(ByRef VEND_CODE As String, VEND_ALT_CODE As String) As List(Of String)

        Dim VEND_BANK_ROUTING_NO As String = ""
        Dim VEND_BANK_ACCT_ID As String = ""
        Dim VEND_BANK_ACCT_CLASS As String = ""
        Dim VEND_BANK_ACCT_TYPE As String = ""

        Dim rowAPTVEND1 As DataRow = tblAPTVEND1.Rows.Find(VEND_CODE)
        If rowAPTVEND1 Is Nothing Then
            errorList.Add($"Vendor {VEND_CODE} is invalid or missing")
            Return errorList
        End If

        If VEND_ALT_CODE = "" Then
            VEND_BANK_ROUTING_NO = rowAPTVEND1.Item("VEND_BANK_ROUTING_NO") & String.Empty
            VEND_BANK_ACCT_ID = rowAPTVEND1.Item("VEND_BANK_ACCT_ID") & String.Empty
            VEND_BANK_ACCT_CLASS = rowAPTVEND1.Item("VEND_BANK_ACCT_CLASS") & String.Empty
            VEND_BANK_ACCT_TYPE = rowAPTVEND1.Item("VEND_BANK_ACCT_TYPE") & String.Empty
        Else
            Dim rowAPTVEND2 As DataRow = tblAPTVEND2.Rows.Find(New String() {VEND_CODE, VEND_ALT_CODE})
            If rowAPTVEND2 Is Nothing Then
                errorList.Add($"Vendor {VEND_CODE} with Payee Code {VEND_ALT_CODE} is invalid or missing")
                Return errorList
            End If

            VEND_BANK_ROUTING_NO = rowAPTVEND2.Item("VEND_ALT_BANK_ROUTING_NO") & String.Empty
            VEND_BANK_ACCT_ID = rowAPTVEND2.Item("VEND_ALT_BANK_ACCT_ID") & String.Empty
            VEND_BANK_ACCT_CLASS = rowAPTVEND2.Item("VEND_ALT_BANK_ACCT_CLASS") & String.Empty
            VEND_BANK_ACCT_TYPE = rowAPTVEND2.Item("VEND_ALT_BANK_ACCT_TYPE") & String.Empty
        End If


        If VEND_BANK_ROUTING_NO = String.Empty Then
            errorList.Add($"Vendor {VEND_CODE} is missing the Bank Account Routing No")
        End If
        If VEND_BANK_ACCT_ID = String.Empty Then
            errorList.Add($"Vendor {VEND_CODE} is missing the Bank Account No")
        End If
        If VEND_BANK_ACCT_CLASS = String.Empty Or Not New String() {"B", "P"}.Contains(VEND_BANK_ACCT_CLASS) Then
            errorList.Add($"Vendor {VEND_CODE} {VEND_ALT_CODE} Bank Account Class is missing or invalid")
        End If
        If VEND_BANK_ACCT_TYPE = String.Empty Or Not New String() {"C", "S"}.Contains(VEND_BANK_ACCT_TYPE) Then
            errorList.Add($"Vendor {VEND_CODE} {VEND_ALT_CODE} Bank Account Type is missing or invalid")
        End If

        Return errorList
    End Function


    Public Function ValidateBank(BANK_CODE As String) As List(Of String)

        If TRAN_NO = "" Then
            errorList.Add($"eCheck Transaction is missing")
            Return errorList
        End If

        If rowGLTBANK1 Is Nothing OrElse rowGLTBANK1.Item("BANK_CODE") & String.Empty <> BANK_CODE Then
            errorList.Add($"Bank {BANK_CODE} is invalid or missing banking information")
            Return errorList
        End If

        If rowGLTBANK1.Item("BANK_MERCHANT_ID") & String.Empty = String.Empty Then
            errorList.Add($"Bank {BANK_CODE} is missing the Bank Merchant ID")
        End If
        If rowGLTBANK1.Item("BANK_MERCHANT_PASSWORD") & String.Empty = String.Empty Then
            errorList.Add($"Bank {BANK_CODE} is missing the Bank Merchant Password")
        End If

        Return errorList
    End Function

End Class
