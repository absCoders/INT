Imports Microsoft.Exchange.WebServices.Data


Public Class ASCNOTEE

    Private NOTE_CODE As String = String.Empty
    Private clsDST As DataSet = Nothing
    Private CUST_CODE As String = String.Empty

    Private tblASTNOTE1 As DataTable
    Private tblASTNOTE2 As DataTable
    Private tblASTNOTE3 As DataTable
    Private tblASTNOTE4 As DataTable

    Private rowARTCUST1 As DataRow
    Private rowSOTSREP1 As DataRow
    Private rowSOTSREP1_ALT As DataRow
    Private rowASTUSER1 As DataRow
    Private rowTATMAIL1 As DataRow
    Private rowASTPARM1 As DataRow

    Private emailSubjectText As String = String.Empty
    Private documentText As String = String.Empty

    Private emailFrom As String = String.Empty
    Private emailTo As String = String.Empty
    Private emailCC As String = String.Empty
    Private emailBCC As String = String.Empty

    Private additionalNote As String = String.Empty
    Private emailSubjectPrefix As String = String.Empty

    Public Folders As New Dictionary(Of String, String)

    Private emailToList As New List(Of String)
    Private emailCCList As New List(Of String)
    Private emailBCCList As New List(Of String)

    Public lastError As String = String.Empty

    Public Enum emailAddresses
        emailTo
        emailCC
        emailBCC
    End Enum

#Region "Instantiate Class"

    Public Sub New(ByVal appFolders As Dictionary(Of String, String))
        InitializeVariables(appFolders)
    End Sub

    Public Sub New(ByVal appFolders As Dictionary(Of String, String), ByVal NoteCode As String, ByRef dst As DataSet)
        InitializeVariables(appFolders)
        NOTE_CODE = NoteCode
        clsDST = dst
    End Sub

    Public Sub New(ByVal appFolders As Dictionary(Of String, String), ByVal NoteCode As String, ByRef dst As DataSet, ByVal CustomerCode As String)
        InitializeVariables(appFolders)
        NOTE_CODE = NoteCode
        clsDST = dst
        CUST_CODE = CustomerCode
    End Sub

    Private Sub InitializeVariables(ByVal appFolders As Dictionary(Of String, String))

        emailSubjectText = String.Empty
        documentText = String.Empty

        emailFrom = String.Empty
        emailTo = String.Empty
        emailCC = String.Empty
        emailBCC = String.Empty

        emailToList = New List(Of String)
        emailCCList = New List(Of String)
        emailBCCList = New List(Of String)

        NOTE_CODE = String.Empty
        clsDST = Nothing
        CUST_CODE = String.Empty
        additionalNote = String.Empty

        Folders = appFolders

        rowTATMAIL1 = ASCDATA1.GetDataRow("SELECT * FROM TATMAIL1 WHERE EMAIL_KEY = 'SO'")
        rowASTPARM1 = ASCDATA1.GetDataRow("SELECT * FROM ASTPARM1 WHERE AS_PARM_KEY = 'Z'")
    End Sub

#End Region

#Region "Properties"

    ''' <summary>
    ''' Get / Set the customer code used to get the Sales Rep and Alternate Sales Rep Code
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CustomerCode() As String
        Get
            Return CUST_CODE
        End Get
        Set(ByVal value As String)
            CUST_CODE = value.ToUpper
        End Set
    End Property

    ''' <summary>
    ''' Set / Get the Note Code used to create document
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property NoteCode() As String
        Get
            Return NOTE_CODE
        End Get
        Set(ByVal value As String)
            NOTE_CODE = value
        End Set
    End Property

    ''' <summary>
    ''' Write Only property to set the dataset used to extract data for the documents merge fields
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property dst() As DataSet
        Set(ByVal value As DataSet)
            clsDST = value
        End Set
    End Property

    ''' <summary>
    ''' Read only property returns Email Subject. Text created since the last 'CreateDocument' call
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property EmailSubject() As String
        Get
            Return emailSubjectText
        End Get
    End Property

    Public Sub SetEmailSubject(ByVal newSubject As String)
        emailSubjectText = newSubject
    End Sub

    ''' <summary>
    ''' Read only property returns Document Text created since the last 'CreateDocument' call
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property DocumentData() As String
        Get
            Return documentText
        End Get
    End Property

    Public Sub SetDocumentBody(ByVal docBody As String)
        documentText = docBody
    End Sub

    ''' <summary>
    ''' Read only property returns Email To created since the last 'CreateDocument' call
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetEmailTo() As String
        Get
            Return emailTo
        End Get
    End Property

    Public Sub SetEmailTo(ByVal sendTo As String)
        emailTo = sendTo
    End Sub

    ''' <summary>
    ''' Read only property returns Email From created since the last 'CreateDocument' call
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetEmailFrom() As String
        Get
            Return emailFrom
        End Get
    End Property

    ''' <summary>
    ''' Read only property returns Email Carbon Copy created since the last 'CreateDocument' call
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetEmailCC() As String
        Get
            Return emailCC
        End Get
    End Property

    ''' <summary>
    ''' Read only property returns Email Blind Carbon Copy created since the last 'CreateDocument' call
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetEmailBCC() As String
        Get
            Return emailBCC
        End Get
    End Property

    ''' <summary>
    ''' Allows a note to be place at the end of the document text.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Note() As String
        Get
            Return additionalNote
        End Get
        Set(ByVal value As String)
            additionalNote = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Text to prefix the Email Subject
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SubjectPrefix() As String
        Get
            Return emailSubjectPrefix
        End Get
        Set(ByVal value As String)
            emailSubjectPrefix = value
        End Set
    End Property

    Public Sub AddEmailAddress(ByVal emailAddress As String, ByVal emailType As emailAddresses)
        Select Case emailType
            Case emailAddresses.emailBCC
                emailBCCList.Add(emailAddress)
            Case emailAddresses.emailCC
                emailCCList.Add(emailAddress)
            Case emailAddresses.emailTo
                emailToList.Add(emailAddress)
        End Select

    End Sub

#End Region

#Region "Procedures"

    Private Sub LoadDocumentData()
        tblASTNOTE1 = ASCDATA1.GetDataTable("SELECT * FROM ASTNOTE1 WHERE NOTE_CODE = :PARM1", String.Empty, "V", NOTE_CODE)
        tblASTNOTE2 = ASCDATA1.GetDataTable("SELECT * FROM ASTNOTE2 WHERE NOTE_CODE = :PARM1", String.Empty, "V", NOTE_CODE)
        tblASTNOTE3 = ASCDATA1.GetDataTable("SELECT * FROM ASTNOTE3 WHERE NOTE_CODE = :PARM1", String.Empty, "V", NOTE_CODE)
        tblASTNOTE4 = ASCDATA1.GetDataTable("SELECT * FROM ASTNOTE4 WHERE NOTE_CODE = :PARM1", String.Empty, "V", NOTE_CODE)
    End Sub

    ''' <summary>
    ''' Creates the Document, Titles by merging Merge Fields with the data in the Dataset.
    ''' and any email address strings
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CreateComponents()

        emailSubjectText = String.Empty
        documentText = String.Empty
        emailTo = String.Empty
        emailFrom = String.Empty
        emailCC = String.Empty
        emailBCC = String.Empty

        LoadDocumentData()

        CreateDocuments()
        CreateEmailComponents()

    End Sub

    Private Sub CreateDocuments()

        Dim tableName As String = String.Empty
        Dim columnName As String = String.Empty
        Dim mergeField As String = String.Empty
        Dim fieldFormat As String = String.Empty
        Dim fieldValue As String = String.Empty

        emailSubjectText = String.Empty
        documentText = String.Empty

        If tblASTNOTE1 Is Nothing OrElse tblASTNOTE1.Rows.Count = 0 Then
            Exit Sub
        End If

        emailSubjectText = tblASTNOTE1.Rows(0).Item("EMAIL_SUBJECT") & String.Empty
        If tblASTNOTE2 IsNot Nothing AndAlso tblASTNOTE1.Rows.Count > 0 Then
            documentText = tblASTNOTE2.Rows(0).Item("NOTE_TEXT") & String.Empty
        End If

        If additionalNote.Length > 0 Then
            documentText &= Environment.NewLine & Environment.NewLine & additionalNote
        End If

        If tblASTNOTE3 IsNot Nothing Then
            For Each rowASTNOTE3 As DataRow In tblASTNOTE3.Rows
                tableName = (rowASTNOTE3.Item("TABLE_NAME") & String.Empty).ToString.Trim
                columnName = (rowASTNOTE3.Item("COLUMN_NAME") & String.Empty).ToString.Trim
                fieldFormat = (rowASTNOTE3.Item("FIELD_FORMAT") & String.Empty).ToString.Trim
                mergeField = "{" & tableName & "." & columnName & "}"

                If tableName.Length = 0 OrElse columnName.Length = 0 Then
                    Continue For
                End If

                If clsDST.Tables.Contains(tableName) _
                    AndAlso clsDST.Tables(tableName).Columns.Contains(columnName) _
                    AndAlso clsDST.Tables(tableName).Rows.Count > 0 Then
                    fieldValue = clsDST.Tables(tableName).Rows(0).Item(columnName) & String.Empty
                    If fieldFormat.Length > 0 Then
                        ' If format fails then trap error and use raw datatable data
                        Try
                            If fieldFormat = "(###) ###-####" Then
                                fieldValue = String.Format("{0:(###) ###-####}", Val(fieldValue))
                            Else
                                fieldValue = Format(Val(fieldValue & String.Empty), fieldFormat)
                            End If
                        Catch ex As Exception
                            ' nothing 
                        End Try
                    End If
                End If

                emailSubjectText = emailSubjectText.Replace(mergeField, fieldValue)
                documentText = documentText.Replace(mergeField, fieldValue)
            Next
        End If

    End Sub

    Private Sub CreateEmailComponents()

        emailTo = String.Empty
        emailFrom = String.Empty
        emailCC = String.Empty
        emailBCC = String.Empty

        Dim srepEmail As String = String.Empty
        Dim altSrepEmail As String = String.Empty
        Dim currentUser As String = String.Empty
        Dim emailaddress As String = String.Empty

        If tblASTNOTE4 Is Nothing OrElse tblASTNOTE4.Rows.Count = 0 Then
            Exit Sub
        End If

        rowARTCUST1 = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = :PARM1", "V", New Object() {CUST_CODE})
        rowASTUSER1 = ASCDATA1.GetDataRow("SELECT * FROM ASTUSER1 WHERE USER_ID = :PARM1", "V", New Object() {ASCMAIN1.USER_ID})
        rowSOTSREP1 = Nothing
        rowSOTSREP1_ALT = Nothing
        If rowARTCUST1 IsNot Nothing Then
            rowSOTSREP1 = ASCDATA1.GetDataRow("SELECT * FROM SOTSREP1 WHERE SREP_CODE = :PARM1", "V", New Object() {rowARTCUST1.Item("SREP_CODE") & String.Empty})
            rowSOTSREP1_ALT = ASCDATA1.GetDataRow("SELECT * FROM SOTSREP1 WHERE SREP_CODE = :PARM1", "V", New Object() {rowARTCUST1.Item("SREP_CODE_ALT") & String.Empty})
        End If

        ' Set up the emails addresss for From, To, CC, BCC
        For Each rowASTCODE4 As DataRow In tblASTNOTE4.Select("", "SEND_LNO")
            emailaddress = String.Empty

            emailaddress = rowASTCODE4.Item("EMAIL_ADDRESS") & String.Empty
            If emailaddress.Length > 0 Then
                emailaddress &= ";"
            End If

            If rowASTCODE4.Item("USE_SREP_CODE") & String.Empty = "1" _
                AndAlso rowSOTSREP1 IsNot Nothing _
                AndAlso rowSOTSREP1.Item("SREP_EMAIL") & String.Empty <> String.Empty Then
                emailaddress &= rowSOTSREP1.Item("SREP_EMAIL") & ";"
            End If

            If rowASTCODE4.Item("ALT_SREP_CODE") & String.Empty = "1" _
                AndAlso rowSOTSREP1_ALT IsNot Nothing _
                AndAlso rowSOTSREP1_ALT.Item("SREP_EMAIL") & String.Empty <> String.Empty Then
                emailaddress &= rowSOTSREP1_ALT.Item("SREP_EMAIL") & ";"
            End If

            If rowASTCODE4.Item("CURRENT_USER") & String.Empty = "1" _
                AndAlso rowASTUSER1 IsNot Nothing _
                AndAlso rowASTUSER1.Item("USER_EMAIL") & String.Empty <> String.Empty Then
                emailaddress &= rowASTUSER1.Item("USER_EMAIL") & ";"
            End If

            If emailaddress.Length = 0 Then Continue For

            Select Case rowASTCODE4.Item("SEND_TYPE") & String.Empty
                Case "F"
                    emailFrom &= emailaddress
                Case "T"
                    emailTo &= emailaddress
                Case "C"
                    emailCC &= emailaddress
                Case "B"
                    emailBCC &= emailaddress
            End Select

        Next

        emailTo &= ";" & String.Join(";", emailToList)
        emailCC &= ";" & String.Join(";", emailCCList)
        emailBCC &= ";" & String.Join(";", emailBCCList)

        While emailFrom.StartsWith(";")
            emailFrom = emailFrom.Substring(1)
        End While

        While emailFrom.Contains(" ")
            emailFrom = emailFrom.Replace(" ", String.Empty)
        End While

        While emailFrom.Contains(";;")
            emailFrom = emailFrom.Replace(";;", ";")
        End While

        If emailFrom.EndsWith(";") Then emailFrom = emailFrom.Substring(0, emailFrom.Length - 1)

        While emailTo.StartsWith(";")
            emailTo = emailTo.Substring(1)
        End While

        While emailTo.Contains(" ")
            emailTo = emailTo.Replace(" ", String.Empty)
        End While

        While emailTo.Contains(";;")
            emailTo = emailTo.Replace(";;", ";")
        End While

        If emailTo.EndsWith(";") Then emailTo = emailTo.Substring(0, emailTo.Length - 1)

        While emailCC.StartsWith(";")
            emailCC = emailCC.Substring(1)
        End While

        While emailCC.Contains(" ")
            emailCC = emailCC.Replace(" ", String.Empty)
        End While

        While emailCC.Contains(";;")
            emailCC = emailCC.Replace(";;", ";")
        End While

        If emailCC.EndsWith(";") Then emailCC = emailCC.Substring(0, emailCC.Length - 1)

        While emailBCC.StartsWith(";")
            emailBCC = emailBCC.Substring(1)
        End While

        While emailBCC.Contains(" ")
            emailBCC = emailBCC.Replace(" ", String.Empty)
        End While

        While emailBCC.Contains(";;")
            emailBCC = emailBCC.Replace(";;", ";")
        End While

        If emailBCC.EndsWith(";") Then emailBCC = emailBCC.Substring(0, emailBCC.Length - 1)

    End Sub

    ''' <summary>
    ''' Sends an email using the Components created from the last call to CreateComponents
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub EmailDocument(Optional emailFileName As String = "")

        If emailTo.Length = 0 OrElse emailFrom.Length = 0 OrElse documentText.Length = 0 Then
            Exit Sub
        End If

        Dim SEND_FROM_SIGNATURE As String = String.Empty
        Dim EMAIL_LOGO As String = String.Empty

        Dim Message As EmailMessage = Nothing

        Dim AS_PARM_EMAIL_USER_ID As String = rowASTPARM1.Item("AS_PARM_EMAIL_USER_ID") & ""
        Dim AS_PARM_EMAIL_PASSWORD As String = ASCMAIN1.DecryptAES(rowASTPARM1.Item("AS_PARM_EMAIL_PASSWORD") & "")

        Dim service As ExchangeService = TACMAIN1.Get_EWS_Service(AS_PARM_EMAIL_USER_ID)
        Message = New EmailMessage(service)

        Try
            Message.From = New EmailAddress(emailFrom, emailFrom)

            For Each sendTo As String In emailTo.Split(";")
                If sendTo.Length > 0 Then
                    Message.ToRecipients.Add(sendTo, sendTo)
                End If
            Next

            For Each cc As String In emailCC.Split(";")
                If cc.Length > 0 Then
                    Message.CcRecipients.Add(cc, cc)
                End If
            Next

            For Each bcc As String In emailBCC.Split(";")
                If bcc.Length > 0 Then
                    Message.BccRecipients.Add(bcc, bcc)
                End If
            Next

            Message.Subject = (emailSubjectPrefix & " " & emailSubjectText).Trim

            If rowTATMAIL1 IsNot Nothing Then
                EMAIL_LOGO = (rowTATMAIL1.Item("EMAIL_LOGO") & String.Empty).ToString.Trim
            End If

            Dim plainView As Net.Mail.AlternateView = Net.Mail.AlternateView.CreateAlternateViewFromString(documentText)
            Dim htmlView As Net.Mail.AlternateView = Net.Mail.AlternateView.CreateAlternateViewFromString("")

            If EMAIL_LOGO <> "" AndAlso Folders.ContainsKey("Images") Then
                Dim logo As FileAttachment = Message.Attachments.AddFileAttachment(Folders("Images") & "ABS\" & EMAIL_LOGO)
                logo.ContentId = "logo"
                Message.Body = "<img src=cid:logo>" & "<p>" & Replace(documentText & vbCrLf & vbCrLf & SEND_FROM_SIGNATURE, vbCrLf, "<br/>") & "</p>"
            Else
                Message.Body = "<p>" & Replace(documentText & vbCrLf & vbCrLf & SEND_FROM_SIGNATURE, vbCrLf, "<br/>") & "</p>"
            End If

            If emailFileName.Length > 0 Then
                Try
                    Message.Save(emailFileName)
                Catch ex As Exception

                End Try
            End If

            Message.SendAndSaveCopy()

        Catch ex As Exception
            lastError = $"Exception: {ex.Message}{Environment.NewLine}Inner Exception: {ex.InnerException.Message}"
        End Try
    End Sub

#End Region

End Class
