Imports Newtonsoft.Json

Public Class GLCJPMC1

    Public Class Account
        Public Property accountId As String
        Public Property accountName As String
        Public Property bankId As String
        Public Property branchId As String
        Public Property bankName As String
        Public Property aba As String
        Public Property swift As Object
        Public Property currency As Currency
    End Class

    Public Class BaiType
        Public Property typeCode As String
        Public Property description As String
        Public Property btrsTypeCode As String
    End Class

    Public Class BankReferenceSearchable
        Public Property standardValue As String
    End Class

    Public Class Currency
        Public Property code As String
        Public Property description As String
    End Class

    Public Class CustomerReferenceSearchable
        Public Property standardValue As String
    End Class

    Public Class Datum
        Public Property account As Account
        Public Property asOfDateTime As DateTime
        Public Property valueDateTime As DateTime
        Public Property asOfDate As String
        Public Property valueDate As String
        Public Property receivedTimestamp As DateTime
        Public Property debitCreditCode As String
        Public Property baiType As BaiType
        Public Property fundsTypeCode As String
        Public Property currency As Currency
        Public Property amount As Double
        Public Property immediateAvailable As Double
        Public Property day1Available As Double
        Public Property day2Available As Double
        Public Property day2PlusAvailable As Object
        Public Property day3PlusAvailable As Double
        Public Property bankReferenceSearchable As BankReferenceSearchable
        Public Property customerReferenceSearchable As CustomerReferenceSearchable
        Public Property repairCode As String
        Public Property reversal As Boolean
        Public Property checkNumber As Integer
        Public Property wireType As String
        Public Property shortDescription As String
        Public Property postCode As String
        Public Property lockbox As Lockbox
        Public Property narrativeText As NarrativeText
        Public Property addenda As List(Of Object)
        Public Property sepaDetailsXml As Object
        Public Property supplementalTextSet As SupplementalTextSet
        Public Property supplementalTextRecordList As Object
        Public Property supplementalText As Object
        Public Property achBatchItems As Object
        Public Property transactionId As String
    End Class

    Public Class Lockbox
        Public Property lockboxSequenceCode As String
        Public Property lockboxItems As Double
        Public Property lockboxNumber As String
        Public Property lockboxDepositDate As Object
        Public Property lockboxDepositTime As Object
    End Class

    Public Class NarrativeText
        <JsonProperty("YOUR REF    ")>
        Public Property YOURREF As String
        <JsonProperty("REC FROM    ")>
        Public Property RECFROM As String
        <JsonProperty("REMARK      ")>
        Public Property REMARK As String
        <JsonProperty("REC GFP     ")>
        Public Property RECGFP As String
        <JsonProperty("B/O CUSTOMER")>
        Public Property BOCUSTOMER As String
        <JsonProperty("B/O BANK    ")>
        Public Property BOBANK As String
        <JsonProperty("CHIP SEQ    ")>
        Public Property CHIPSEQ As String
        <JsonProperty("CHIP REF    ")>
        Public Property CHIPREF As String
        <JsonProperty("ACCT PARTY  ")>
        Public Property ACCTPARTY As String
        <JsonProperty("ULTI BENE   ")>
        Public Property ULTIBENE As String
        <JsonProperty("PAID TO     ")>
        Public Property PAIDTO As String
    End Class

    Public Class Pagination
        Public Property pageSize As Integer
        Public Property totalPages As Integer
        Public Property pageNumber As Integer
        Public Property totalRecords As Integer
    End Class

    Public Class JPMC_Transactions_Response
        Public Property pagination As Pagination
        Public Property data As List(Of Datum)
    End Class

    Public Class SupplementalTextSet
    End Class

End Class
