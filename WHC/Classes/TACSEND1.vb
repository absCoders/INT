Imports Microsoft.Exchange.WebServices.Data
Public Class TACSEND1
    ' Send email

    Public SEND_FROM As String
    Public SEND_FROM_NAME As String
    Public SEND_FROM_SIGNATURE As String
    Public SEND_TO As String
    Public SEND_TO_NAME As String
    Public SEND_TOs As New Dictionary(Of String, String)
    Public SEND_CC As String
    Public SEND_CC_NAME As String
    Public SEND_BCC As String
    Public SEND_BCC_NAME As String
    Public SEND_SUBJECT As String
    Public SEND_BODY As String
    Public SEND_ENTITY_CAPTION As String
    Public SEND_ENTITY_TABLE As String
    Public SEND_ENTITY_KEY As String
    Public SEND_ENTITY_NAME As String
    Public SEND_METHOD As String
    Public SEND_ATTACHMENT As String
    Public SEND_ATTACHMENTs As Dictionary(Of String, String) = Nothing
    Public SEND_STATUS As String
    Public SEND_ERROR As String

    Public SEND_NO As String
    Public SEND_LOG As String
    Public SEND_ID As String
    Public rowTATSEND1 As DataRow
    Public rowTATMAIL1 As DataRow
    Public EMAIL_KEY As String
    Dim fASCBASE0 As ASCBASE0
    Dim dst As DataSet
    Dim ASCMAIN1 As ASCMAIN1
    Sub New()
        MyBase.New()
    End Sub
    Sub New(f As ASCBASE0)
        fASCBASE0 = f
        dst = fASCBASE0.dst
        ASCMAIN1 = fASCBASE0.ASCMAIN1
        '   Main_Process()
    End Sub
    Public Sub Send_email_automatically(Optional ByVal bcc_User As Boolean = True)

        ' Clear_dst()
        Prepare_Send_Log()

        If Send_email(True) Then
            Update_Send_Log()
        End If
    End Sub

    Function Send_email(Optional ByVal auto_send As Boolean = False) As Boolean

        Send_email = False
        SEND_ERROR = ""

        Try

            ' Evaluate the Email Addresses
            If Not ValidateEmail(SEND_FROM) Then
                Throw New Exception("Invalid Send From email address.")
                Return False
            End If

            If SEND_TOs Is Nothing OrElse SEND_TOs.Count = 0 Then
                For Each SEND_TO_email_address As String In Split(SEND_TO, ";")
                    SEND_TO_email_address = Trim(SEND_TO_email_address)
                    If Not ValidateEmail(SEND_TO_email_address) Then
                        Throw New Exception("Invalid Send To email address.")
                        Return False
                    End If
                Next
            Else
                For Each SEND_TO_EMAIL As String In SEND_TOs.Keys
                    If Not ValidateEmail(SEND_TO_EMAIL) Then
                        Throw New Exception("Invalid Send To email address.")
                        Return False
                    End If
                Next
            End If

            If SEND_CC IsNot Nothing Then
                If SEND_CC <> "" Then
                    For Each SEND_CC_email_address As String In Split(SEND_CC, ";")
                        SEND_CC_email_address = Trim(SEND_CC_email_address)
                        If Not ValidateEmail(SEND_CC_email_address) Then
                            Throw New Exception("Invalid Carbon Copy (cc) email address.")
                            Return False
                        End If
                    Next
                End If
            End If

            If SEND_BCC <> "" AndAlso Not ValidateEmail(SEND_BCC) Then
                Throw New Exception("Invalid Blind Carbon Copy (cc) email address.")
                Return False
            End If

            SEND_NO = ASCMAIN1.Next_Control_No("TATSEND1.SEND_NO")

            Dim folder As String = ASCMAIN1.Folders("Archive") & "email\Sent\"
            If Not My.Computer.FileSystem.DirectoryExists(folder) Then
                My.Computer.FileSystem.CreateDirectory(folder)
            End If

            fASCBASE0.Get_PARM("ASTPARM1")

            sendViaEws()

            SEND_STATUS = "S"
            Update_Send_Log()
            Return True

        Catch ex As Exception
            If ASCMAIN1.Running_in_VS Then
                Stop
            End If
            SEND_STATUS = "E"
            SEND_NO = ""
            SEND_ERROR = ex.Message.ToString

            If Not auto_send Then
                MsgBox("Error Occured: " & ex.Message, MsgBoxStyle.OkOnly, "Could not Send email")
            End If
            Return False
        End Try

    End Function

    Async Sub sendViaEws()

        Dim AS_PARM_EMAIL_USER_ID As String = fASCBASE0.ROWs("ASTPARM1").Item("AS_PARM_EMAIL_USER_ID") & ""
        Dim service As ExchangeService = Await TACMAIN1.Get_EWS_Service(AS_PARM_EMAIL_USER_ID)

        Dim Message As EmailMessage = New EmailMessage(service)


        If SEND_TOs Is Nothing OrElse SEND_TOs.Count = 0 Then
            If InStr(SEND_TO, ";") = 0 Then

                Message.ToRecipients.Add(New EmailAddress(SEND_TO_NAME, SEND_TO))
            Else
                For Each SEND_TO_email_address As String In Split(SEND_TO, ";")
                    SEND_TO_email_address = Trim(SEND_TO_email_address)
                    Dim SEND_TO_email_address_NAME As String = ""
                    ASCMAIN1.sql = "Select * from TATCONT1 " _
                    & " where CONTACT_ENTITY_TABLE = '" & SEND_ENTITY_TABLE & "'" _
                    & "   and CONTACT_ENTITY_KEY = '" & SEND_ENTITY_KEY & "'" _
                    & "   and LOWER(CONTACT_EMAIL) = :PARM1"
                    Dim rowTATCONT1 As DataRow = fASCBASE0.ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", SEND_TO_email_address.ToLower)
                    If rowTATCONT1 IsNot Nothing Then
                        SEND_TO_email_address_NAME = rowTATCONT1.Item("CONTACT_NAME") & ""
                    End If
                    Message.ToRecipients.Add(New EmailAddress(SEND_TO_email_address_NAME, SEND_TO_email_address))
                Next
            End If
        Else
            For Each SEND_TO As String In SEND_TOs.Keys
                Dim SEND_TO_NAME As String = SEND_TOs(SEND_TO)
                Message.ToRecipients.Add(New EmailAddress(SEND_TO_NAME, SEND_TO))
            Next
        End If

        If SEND_CC IsNot Nothing Then
            If SEND_CC <> "" Then
                For Each SEND_CC_email_address As String In Split(SEND_CC, ";")
                    SEND_CC_email_address = Trim(SEND_CC_email_address)
                    Dim SEND_CC_email_address_NAME As String = ""
                    ASCMAIN1.sql = "Select * from TATCONT1 " _
                    & " where CONTACT_ENTITY_TABLE = '" & SEND_ENTITY_TABLE & "'" _
                    & "   and CONTACT_ENTITY_KEY = '" & SEND_ENTITY_KEY & "'" _
                    & "   and LOWER(CONTACT_EMAIL) = :PARM1"
                    Dim rowTATCONT1 As DataRow = fASCBASE0.ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", SEND_CC_email_address.ToLower)
                    If rowTATCONT1 IsNot Nothing Then
                        SEND_CC_email_address_NAME = rowTATCONT1.Item("CONTACT_NAME") & ""
                    End If
                    Message.CcRecipients.Add(New EmailAddress(SEND_CC_email_address_NAME, SEND_CC_email_address))
                Next
            End If
        End If


        If SEND_BCC <> "" Then
            Message.BccRecipients.Add(New EmailAddress(SEND_BCC_NAME, SEND_BCC))
        End If

        Message.From = New EmailAddress(SEND_FROM_NAME, SEND_FROM)
        Message.Subject = IIf(ASCMAIN1.DBS_COMPANY = "TST", "Test Company - ", "") & SEND_SUBJECT

        Dim LINKS As String = ""

        If SEND_ATTACHMENTs Is Nothing Then
            If SEND_ATTACHMENT <> "" Then
                For Each ss As String In SEND_ATTACHMENT.Split(";")
                    If Trim(ss) <> "" Then
                        Message.Attachments.AddFileAttachment(ss.Trim)
                    End If
                Next
            End If
        Else
            For Each ATTACHMENT_FILE As String In SEND_ATTACHMENTs.Keys
                If SEND_ATTACHMENTs(ATTACHMENT_FILE).StartsWith("http://") Then
                    LINKS &= "<br><a href='" & SEND_ATTACHMENTs(ATTACHMENT_FILE) & "'>" & ATTACHMENT_FILE & "</a>"
                Else
                    Message.Attachments.AddFileAttachment(SEND_ATTACHMENTs(ATTACHMENT_FILE))
                End If
            Next
        End If

        If LINKS <> "" Then
            LINKS = vbCrLf & LINKS
        End If

        If rowTATMAIL1 Is Nothing Then
            rowTATMAIL1 = fASCBASE0.LookUp("TATMAIL1", EMAIL_KEY)
        End If

        Dim EMAIL_LOGO As String = ""
        If rowTATMAIL1 IsNot Nothing Then
            EMAIL_LOGO = rowTATMAIL1.Item("EMAIL_LOGO") & ""
        End If

        ' Re: how to embed image in mail body while sending mail in c#.net uisng Exchange2007_SP1.
        ' Aug 25, 2010 10:11 PM|LINK
        ' I don't know if you ever figured this out, but I thought I'd do you and/or anybody else trying to figure this out a solid.
        ' The answer is that the Exchange web service API implements the ContentID property in the AttachmentType class and its derived classes which is similar to the LinkedResource class in System.Net.Mail.  For files, the FileAttachmentType class should be used.
        ' So if you want to embed an image in the HTML you would add the attachment (FileAttachmentType) to the message normally, but also assign its ContentID property. (GUIDs work good for this)
        ' Then simply set the BodyType to HTML and set the src attribute for any image tags in the HTML that reference the image file to "cid:yourcontentid". (Regular expressions work good for this)
        ' Nothing to it.  I hope that helps someone out there and dispells the rumor that this kind of thing is not possible in the EWS API.

        'Dim plainView As AlternateView = AlternateView.CreateAlternateViewFromString(SEND_BODY)
        'Dim htmlView As AlternateView
        If EMAIL_LOGO <> "" Then
            'htmlView = AlternateView.CreateAlternateViewFromString("<img src=cid:logo>" & "<p>" & Replace(SEND_BODY & vbCrLf & vbCrLf & SEND_FROM_SIGNATURE, vbCrLf, "<br>") & "</p>", Nothing, "text/html")
            'Dim logo As New LinkedResource(ASCMAIN1.Folders("Images") & "ABS\" & EMAIL_LOGO)
            'logo.ContentId = "logo"
            'htmlView.LinkedResources.Add(logo)

            Dim logo As FileAttachment = Message.Attachments.AddFileAttachment(ASCMAIN1.Folders("Images") & "ABS\" & EMAIL_LOGO)
            logo.ContentId = "logo"
            Message.Body = "<img src=cid:logo>" & "<p>" & Replace(SEND_BODY & LINKS & vbCrLf & vbCrLf & SEND_FROM_SIGNATURE, vbCrLf, "<br>") & "</p>"

        Else
            'htmlView = AlternateView.CreateAlternateViewFromString("<p>" & SEND_BODY & vbCrLf & vbCrLf & SEND_FROM_SIGNATURE & "</p>", Nothing, "text/html")
            Message.Body = "<p>" & Replace(SEND_BODY & LINKS & vbCrLf & vbCrLf & SEND_FROM_SIGNATURE, vbCrLf, "<br>") & "</p>"

        End If

        'Message.Body.BodyType = BodyType.HTML


        'Message.AlternateViews.Add(plainView)
        'Message.AlternateViews.Add(htmlView)


        Dim folder As String = ASCMAIN1.Folders("Archive") & "email\Sent\"
        If Not My.Computer.FileSystem.DirectoryExists(folder) Then
            My.Computer.FileSystem.CreateDirectory(folder)
        End If

        Message.Save()
        Message.SendAndSaveCopy()
        'Message.SaveToFile(folder & SEND_NO & ".eml")
    End Sub


    Sub Prepare_Send_Log()
        With dst
            If Not dst.Tables.Contains("TATSEND1") Then
                fASCBASE0.Create_TDA(.Tables.Add, "TATSEND1", "*")
            End If
        End With

        rowTATSEND1 = dst.Tables("TATSEND1").NewRow
        rowTATSEND1.Item("SEND_NO") = "0000000000"
        rowTATSEND1.Item("INIT_DATE") = fASCBASE0.DATETIME_STAMP
        dst.Tables("TATSEND1").Rows.Add(rowTATSEND1)

    End Sub

    Sub Update_Send_Log()
        rowTATSEND1.Item("SEND_NO") = SEND_NO
        rowTATSEND1.Item("INIT_OPER") = "service"
        rowTATSEND1.Item("INIT_DATE") = Now
        fASCBASE0.Update_Record_TDA("TATSEND1")
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

End Class
