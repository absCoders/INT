Imports System.Net.Mail
Imports Microsoft.Exchange.WebServices.Data

Public Class TAFSEND1
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
    Public EMAIL As MailMessage

    Public SEND_BODY_OVERRIDE As String = String.Empty

    Private setupScreen As Boolean = True

    Public Sub New(ByVal FF As ASFBASE1)
        frmASFBASE1 = FF
        InitializeComponent()
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Prepare_Send_Log()

        If SEND_ENTITY_CAPTION <> "" Then
            grpEntity.Text = SEND_ENTITY_CAPTION
        End If
        Set_Read_Only(grpEntity, True)
        Set_Read_Only(grpFrom, True)

        If ASCMAIN1.USER_EMAIL <> "" Then
            chkBCC.Text = "BCC " & ASCMAIN1.USER_EMAIL & ", " & ASCMAIN1.USER_NAME
            chkBCC.Visible = True
            If SEND_CC IsNot Nothing AndAlso Not SEND_CC.Contains(ASCMAIN1.USER_EMAIL) Then
                chkBCC.Checked = True
            Else
                chkBCC.Checked = False
            End If
        Else
            chkBCC.Visible = False
            chkBCC.Checked = False
        End If
    End Sub

    Private Sub Form_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Setup_Form_Defaults()
    End Sub

    Private Sub Setup_Form_Defaults()

        Setup_BCC()

        Setup_SEND_TYPE()

        If SEND_METHOD = "E" Then
            Me.Text = "email"
        Else
            Me.Text = "Fax"
        End If

    End Sub

    Sub Setup_BCC()
        If ASCMAIN1.USER_EMAIL = "" Then
            chkBCC.Checked = False
            chkBCC.Visible = False
        Else
            chkBCC.Text = "BCC " & ASCMAIN1.USER_EMAIL
            If SEND_CC IsNot Nothing AndAlso Not SEND_CC.Contains(ASCMAIN1.USER_EMAIL) Then
                chkBCC.Checked = False
            Else
                chkBCC.Checked = True
            End If
        End If
    End Sub

    ''' <summary>
    ''' Send an email without displaying the form
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Send_email_automatically(Optional ByVal bcc_User As Boolean = True)

        setupScreen = False
        optType.Value = "E"

        DATETIME_STAMP = Now + ASCMAIN1.NowTSD
        Clear_dst()
        Prepare_Send_Log()

        If bcc_User Then
            Setup_BCC()
        End If

        If Send_email(True) Then
            Update_Send_Log()
        End If

        Me.Close()
    End Sub

    ''' <summary>
    ''' Send a fax without displaying the form
    ''' </summary>
    ''' <remarks></remarks>
    ''' 
    Public Sub Send_fax_automatically()

        setupScreen = False
        optType.Value = "F"

        DATETIME_STAMP = Now + ASCMAIN1.NowTSD
        Clear_dst()
        Prepare_Send_Log()
        Setup_Form_Defaults()

        If Send_fax(True) Then
            Update_Send_Log()
        End If

        Me.Close()
    End Sub

    Sub Screen_Fields(ByVal LU As String)
        If LU = "Load" Then

            If setupScreen Then
                For Each COLUMN_NAME As String In New String() _
                {"SEND_TO", "SEND_TO_NAME", "SEND_CC", "SEND_SUBJECT", "SEND_ATTACHMENT",
                 "SEND_ENTITY_KEY", "SEND_BODY", "SEND_FROM", "SEND_FROM_NAME", "SEND_FROM_SIGNATURE"}
                    Absx1.txtFor(COLUMN_NAME).MaxLength = rowTATSEND1.Table.Columns(COLUMN_NAME).MaxLength
                Next
            End If

            If SEND_TOs Is Nothing OrElse SEND_TOs.Count = 0 Then
                rowTATSEND1("SEND_TO") = SEND_TO
                rowTATSEND1("SEND_TO_NAME") = SEND_TO_NAME
                If setupScreen Then
                    Absx1.txtFor("SEND_TOS").Visible = True
                    Absx1.txtFor("SEND_TO").Visible = False
                    Absx1.txtFor("SEND_TO_NAME").Visible = False
                    Absx1.txtFor("SEND_TOS").Focus()
                End If
            Else
                Dim SENT_TO_EMAILS As String = ""
                For Each SEND_TO As String In SEND_TOs.Keys
                    Dim SEND_TO_NAME As String = SEND_TOs(SEND_TO)
                    SENT_TO_EMAILS &= ";" & SEND_TO
                Next
                SEND_TOs = Nothing
                If setupScreen Then
                    Absx1.txtFor("SEND_TOS").Text = Mid(SENT_TO_EMAILS, 2)
                    'Absx1.txtFor("SEND_TOS").ReadOnly = True
                    Absx1.txtFor("SEND_TOS").Visible = True
                    Absx1.txtFor("SEND_TO").Visible = False
                    Absx1.txtFor("SEND_TO_NAME").Visible = False
                    Absx1.txtFor("SEND_TOS").Focus()
                End If
            End If
            rowTATSEND1("SEND_CC") = SEND_CC
            rowTATSEND1("SEND_CC_NAME") = SEND_CC_NAME
            rowTATSEND1("SEND_SUBJECT") = SEND_SUBJECT

            If SEND_ATTACHMENTs Is Nothing Then
                rowTATSEND1("SEND_ATTACHMENT") = SEND_ATTACHMENT
            Else
                Dim ATTACHMENT_FILEs As String = ""
                For Each ATTACHMENT_FILE As String In SEND_ATTACHMENTs.Keys
                    ATTACHMENT_FILEs &= ";" & ATTACHMENT_FILE
                Next
                rowTATSEND1("SEND_ATTACHMENT") = Mid(ATTACHMENT_FILEs, 2)
            End If

            rowTATSEND1("SEND_ENTITY_TABLE") = SEND_ENTITY_TABLE
            rowTATSEND1("SEND_ENTITY_KEY") = SEND_ENTITY_KEY
            rowTATSEND1("SEND_ENTITY_NAME") = SEND_ENTITY_NAME
            rowTATSEND1("SEND_METHOD") = SEND_METHOD
            rowTATSEND1("SEND_BODY") = Mid(SEND_BODY, 1, rowTATSEND1.Table.Columns("SEND_BODY").MaxLength)
            rowTATSEND1("SEND_FROM") = SEND_FROM
            rowTATSEND1("SEND_FROM_NAME") = SEND_FROM_NAME
            rowTATSEND1("SEND_FROM_SIGNATURE") = SEND_FROM_SIGNATURE
        Else
            SEND_TO = Absx1.txtFor("SEND_TO").Text.Trim
            SEND_TO_NAME = Absx1.txtFor("SEND_TO_NAME").Text.Trim
            SEND_CC = Absx1.txtFor("SEND_CC").Text.Trim
            ' SEND_CC_NAME = Absx1.txtFor("SEND_CC_NAME").Text.Trim
            SEND_SUBJECT = Absx1.txtFor("SEND_SUBJECT").Text.Trim
            SEND_ATTACHMENT = Absx1.txtFor("SEND_ATTACHMENT").Text.Trim
            SEND_ENTITY_KEY = Absx1.txtFor("SEND_ENTITY_KEY").Text.Trim
            SEND_ENTITY_NAME = Absx1.txtFor("SEND_ENTITY_NAME").Text.Trim
            SEND_METHOD = Absx1.optFor("SEND_METHOD").Value.Trim
            SEND_BODY = Absx1.txtFor("SEND_BODY").Text.Trim
            SEND_FROM = Absx1.txtFor("SEND_FROM").Text.Trim
            SEND_FROM_NAME = Absx1.txtFor("SEND_FROM_NAME").Text.Trim
            SEND_FROM_SIGNATURE = Absx1.txtFor("SEND_FROM_SIGNATURE").Text.Trim
        End If
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        SEND_STATUS = "C"
        Me.Close()
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
        End Select
    End Sub

    Overrides Sub txt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Dim txtctl As UltraWinEditors.UltraTextEditor = DirectCast(sender, UltraWinEditors.UltraTextEditor)

        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
        End Select

    End Sub

    Function Send_email(Optional ByVal auto_send As Boolean = False) As Boolean

        Send_email = False
        SEND_ERROR = ""
        SEND_TO = txtSEND_TOS.Text
        If Not txtSEND_TOS.Visible Then
            SEND_TO = txtSEND_TO.Text
        End If

        Try

            ' Evaluate the Email Addresses
            If Not ValidateEmail(SEND_FROM) Then
                If Not auto_send Then
                    MessageBox.Show("Invalid Send From email address: " & SEND_FROM, "eMail", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    ASCMAIN1.Message = "Invalid Send From email address: " & SEND_FROM
                End If
                Return False
            End If

            If SEND_TOs Is Nothing OrElse SEND_TOs.Count = 0 Then
                For Each SEND_TO_email_address As String In Split(SEND_TO, ";")
                    SEND_TO_email_address = Trim(SEND_TO_email_address)
                    If Not ValidateEmail(SEND_TO_email_address) Then
                        If Not auto_send Then
                            MessageBox.Show("Invalid Send To email address (" & SEND_TO_email_address & ").", "eMail", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Else
                            ASCMAIN1.Message = "Invalid Send To email address (" & SEND_TO_email_address & ")."
                        End If
                        Return False
                    End If
                Next
            Else
                For Each SEND_TO_EMAIL As String In SEND_TOs.Keys
                    If Not ValidateEmail(SEND_TO_EMAIL) Then
                        If Not auto_send Then
                            MessageBox.Show("Invalid Send To email address (" & SEND_TO_EMAIL & ").", "eMail", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Else
                            ASCMAIN1.Message = "Invalid Send To email address (" & SEND_TO_EMAIL & ")."
                        End If
                        Return False
                    End If
                Next
            End If

            If SEND_CC IsNot Nothing Then
                If SEND_CC <> "" Then
                    For Each SEND_CC_email_address As String In Split(SEND_CC, ";")
                        SEND_CC_email_address = Trim(SEND_CC_email_address)
                        If Not ValidateEmail(SEND_CC_email_address) Then
                            If Not auto_send Then
                                MessageBox.Show("Invalid Carbon Copy (cc) email address: " & SEND_CC_email_address, "eMail", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Else
                                ASCMAIN1.Message = "Invalid Carbon Copy (cc) email address: " & SEND_CC_email_address
                            End If
                            Return False
                        End If
                    Next
                End If
            End If

            If SEND_BCC <> "" AndAlso Not ValidateEmail(SEND_BCC) Then
                If Not auto_send Then
                    MessageBox.Show("Invalid Blind Carbon Copy (bcc) email address.", "eMail", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    ASCMAIN1.Message = "Invalid Blind Carbon Copy (bcc) email address: " & SEND_BCC
                End If
                Return False
            End If

            If Not auto_send Then
                Me.Cursor = Cursors.WaitCursor
            End If


            SEND_NO = ASCMAIN1.Next_Control_No("TATSEND1.SEND_NO")

            Dim folder As String = ASCMAIN1.Folders("Archive") & "email\Sent\"
            If Not My.Computer.FileSystem.DirectoryExists(folder) Then
                My.Computer.FileSystem.CreateDirectory(folder)
            End If
            If ASCMAIN1.JOB_STREAM_CODE <> "" Then
                'JobStreamReports
                If Not My.Computer.FileSystem.DirectoryExists(folder & "JobStreamReports\") Then
                    My.Computer.FileSystem.CreateDirectory(folder & "JobStreamReports\")
                End If
            End If


            Get_PARM("ASTPARM1")

            Dim emailType As String = String.Empty
            If EMAIL_KEY.Length > 0 Then
                rowTATMAIL1 = LookUp("TATMAIL1", EMAIL_KEY)
                If rowTATMAIL1 IsNot Nothing Then
                    If rowTATMAIL1.Item("EMAIL_TYPE") & String.Empty <> String.Empty Then
                        emailType = rowTATMAIL1.Item("EMAIL_TYPE") & String.Empty
                    End If
                End If
            End If

            If emailType = "SMTPHVE" Then
                Dim mail As MailMessage = createSmtpEmailMessage()
                If Not sendViaSmtpHVE(mail) Then
                    Return False
                End If
            ElseIf emailType = "SMTP" Then
                Dim mail As MailMessage = createSmtpEmailMessage()
                If Not sendViaSmtp(mail) Then
                    Return False
                End If
            ElseIf ROWs("ASTPARM1").Item("AS_PARM_EMAIL_TYPE") & "" = "EWS" OrElse emailType = "EWS" Then
                sendViaEws()
            Else
                Dim mail As MailMessage = createSmtpEmailMessage()
                Try
                    mail.Save(folder & SEND_NO & ".eml")
                Catch ex As Exception

                End Try

                If Not sendViaSmtp(mail) Then
                    Return False
                End If
            End If

            SEND_STATUS = "S"
            Screen_Fields("Load")
            Update_Send_Log()

            If Not auto_send Then
                MsgBox("email has been sent", MsgBoxStyle.OkOnly, "Verification")
            End If
            Return True

        Catch ex As Exception
            If ASCMAIN1.Running_in_VS Then
                Stop
            End If
            SEND_STATUS = "E"
            SEND_NO = ""
            SEND_ERROR = ex.Message.ToString
            If ASCMAIN1.USER_ID = "wjz" Then MsgBox("catch:  " & ex.InnerException.Message)
            If Not auto_send Then
                MsgBox("Error Occured: " & ex.Message, MsgBoxStyle.OkOnly, "Could not Send email")
            Else
                ASCMAIN1.Message = ex.Message
            End If
            Return False
        End Try

    End Function

    Function createSmtpEmailMessage() As MailMessage
        Dim mail As New MailMessage()
        mail.From = New MailAddress(SEND_FROM, SEND_FROM_NAME)

        If SEND_TOs Is Nothing OrElse SEND_TOs.Count = 0 Then
            If InStr(SEND_TO, ";") = 0 Then

                mail.To.Add(New MailAddress(SEND_TO, SEND_TO_NAME))
            Else
                For Each SEND_TO_email_address As String In Split(SEND_TO, ";")
                    SEND_TO_email_address = Trim(SEND_TO_email_address)
                    Dim SEND_TO_email_address_NAME As String = ""
                    ASCMAIN1.sql = "Select * from TATCONT1 " _
                    & " where CONTACT_ENTITY_TABLE = '" & SEND_ENTITY_TABLE & "'" _
                    & "   and CONTACT_ENTITY_KEY = '" & SEND_ENTITY_KEY & "'" _
                    & "   and LOWER(CONTACT_EMAIL) = :PARM1"
                    Dim rowTATCONT1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", SEND_TO_email_address.ToLower)
                    If rowTATCONT1 IsNot Nothing Then
                        SEND_TO_email_address_NAME = rowTATCONT1.Item("CONTACT_NAME") & ""
                    End If
                    mail.To.Add(New MailAddress(SEND_TO_email_address, SEND_TO_email_address_NAME))
                Next
            End If
        Else
            For Each SEND_TO As String In SEND_TOs.Keys
                Dim SEND_TO_NAME As String = SEND_TOs(SEND_TO)
                mail.To.Add(New MailAddress(SEND_TO, SEND_TO_NAME))
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
                    Dim rowTATCONT1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", SEND_CC_email_address.ToLower)
                    If rowTATCONT1 IsNot Nothing Then
                        SEND_CC_email_address_NAME = rowTATCONT1.Item("CONTACT_NAME") & ""
                    End If
                    mail.CC.Add(New MailAddress(SEND_CC_email_address, SEND_CC_email_address_NAME))
                Next
            End If
        End If


        If SEND_BCC <> "" Then
            mail.Bcc.Add(New MailAddress(SEND_BCC, SEND_BCC_NAME))
        End If

        If chkBCC.Checked Then
            mail.Bcc.Add(New MailAddress(ASCMAIN1.USER_EMAIL, ASCMAIN1.USER_NAME))
        End If

        mail.Subject = IIf(ASCMAIN1.DBS_COMPANY = "TST", "Test Company - ", "") & SEND_SUBJECT

        If SEND_ATTACHMENTs Is Nothing Then
            If SEND_ATTACHMENT <> "" Then
                For Each ss As String In SEND_ATTACHMENT.Split(";")
                    If Trim(ss) <> "" Then
                        mail.Attachments.Add(New System.Net.Mail.Attachment(ss.Trim))
                    End If
                Next
            End If
        Else
            Dim ATTACHMENT_FILEs As String = ""
            For Each ATTACHMENT_FILE As String In SEND_ATTACHMENTs.Keys
                mail.Attachments.Add(New System.Net.Mail.Attachment(SEND_ATTACHMENTs(ATTACHMENT_FILE)))
            Next
        End If

        If EMAIL IsNot Nothing Then
            If EMAIL.Attachments.Count > 0 Then
                For Each MA As System.Net.Mail.Attachment In EMAIL.Attachments
                    mail.Attachments.Add(MA)
                Next
            End If
        End If

        If rowTATMAIL1 Is Nothing Then
            rowTATMAIL1 = LookUp("TATMAIL1", EMAIL_KEY)
        End If

        Dim EMAIL_LOGO As String = ""
        If rowTATMAIL1 IsNot Nothing Then
            EMAIL_LOGO = rowTATMAIL1.Item("EMAIL_LOGO") & ""
        End If

        ' Done since the data is stored n a table anfthe data is truncated to fit in the field
        If SEND_BODY_OVERRIDE.Length > 0 Then
            SEND_BODY = SEND_BODY_OVERRIDE
        End If

        Dim plainView As AlternateView = AlternateView.CreateAlternateViewFromString(SEND_BODY)
        Dim htmlView As AlternateView
        If EMAIL_LOGO <> "" Then
            htmlView = AlternateView.CreateAlternateViewFromString("<img src=cid:logo>" & "<p>" & Replace(SEND_BODY & vbCrLf & vbCrLf & SEND_FROM_SIGNATURE, vbCrLf, "<br>") & "</p>", Nothing, "text/html")
            Dim logo As New LinkedResource(ASCMAIN1.Folders("Images") & "ABS\" & EMAIL_LOGO)
            logo.ContentId = "logo"
            htmlView.LinkedResources.Add(logo)
        Else
            htmlView = AlternateView.CreateAlternateViewFromString("<p>" & SEND_BODY & vbCrLf & vbCrLf & SEND_FROM_SIGNATURE & "</p>", Nothing, "text/html")
        End If

        mail.AlternateViews.Add(plainView)
        mail.AlternateViews.Add(htmlView)

        Return mail
    End Function

    Function sendViaSmtp(ByVal mail As MailMessage) As Boolean
        Dim smtp As New SmtpClient(ASCMAIN1.rowASTPARM1.Item("AS_PARM_EMAIL_SMTP_IP"), Val(ASCMAIN1.rowASTPARM1.Item("AS_PARM_EMAIL_SMTP_PORT") & String.Empty))
        If smtp IsNot Nothing Then
            Dim EMAIL_ACCT_ID As String = rowTATMAIL1.Item("EMAIL_ACCT_ID")
            Dim EMAIL_ACCT_PWD As String = rowTATMAIL1.Item("EMAIL_ACCT_PWD")
            smtp.Credentials = New System.Net.NetworkCredential(EMAIL_ACCT_ID, EMAIL_ACCT_PWD)
        Else
            Dim eMsg As String = "SMTP Client could not be created."
            MsgBox(eMsg, MsgBoxStyle.OkOnly, "Error")
            Return False
        End If

        If ASCMAIN1.DBS_COMPANY = ASCMAIN1.DBS_SERVER Then
            System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Ssl3 Or Net.SecurityProtocolType.Tls Or Net.SecurityProtocolType.Tls11 Or Net.SecurityProtocolType.Tls12
            smtp.Send(mail)
        End If

        Return True
    End Function

    Function sendViaSmtpHVE(ByVal mail As MailMessage) As Boolean

        Dim smtp As New SmtpClient(ASCMAIN1.rowASTPARM1.Item("AS_PARM_EMAIL_HVE_SMTP_IP"), Val(ASCMAIN1.rowASTPARM1.Item("AS_PARM_EMAIL_HVE_SMTP_PORT")))

        If smtp IsNot Nothing Then

            If ASCMAIN1.USER_ID = "wjz" And rowTATMAIL1 Is Nothing Then
                MsgBox("rowTATMAIL1 is Nothing")
            End If
            'If ASCMAIN1.USER_ID = "wjz" Then MsgBox("rowTATMAIL1:" & rowTATMAIL1.ToString)

            Dim EMAIL_ACCT_ID As String = rowTATMAIL1.Item("EMAIL_ACCT_ID")
            Dim EMAIL_ACCT_PWD As String = rowTATMAIL1.Item("EMAIL_ACCT_PWD")

            'If ASCMAIN1.USER_ID = "wjz" Then MsgBox("creds:" & rowTATMAIL1.Item("EMAIL_ACCT_ID") & ":" & rowTATMAIL1.Item("EMAIL_ACCT_PWD"))

            smtp.Credentials = New System.Net.NetworkCredential(EMAIL_ACCT_ID, EMAIL_ACCT_PWD)
        Else
            Dim eMsg As String = "SMTP HVE Client could Not be created."
            MsgBox(eMsg, MsgBoxStyle.OkOnly, "Error")
            Return False
        End If

        If ASCMAIN1.DBS_COMPANY = ASCMAIN1.DBS_SERVER Then
            'If ASCMAIN1.USER_ID = "wjz" Then MsgBox("AS_PARM_EMAIL_HVE_SMTP_TARGET" & ASCMAIN1.rowASTPARM1.Item("AS_PARM_EMAIL_HVE_SMTP_TARGET") & String.Empty)
            System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Ssl3 Or Net.SecurityProtocolType.Tls Or Net.SecurityProtocolType.Tls11 Or Net.SecurityProtocolType.Tls12
            smtp.EnableSsl = True
            smtp.TargetName = ASCMAIN1.rowASTPARM1.Item("AS_PARM_EMAIL_HVE_SMTP_TARGET") & String.Empty
            smtp.Send(mail)
        End If

        Return True
    End Function

    Sub sendViaEws()

        Dim AS_PARM_EMAIL_USER_ID As String = ROWs("ASTPARM1").Item("AS_PARM_EMAIL_USER_ID") & ""
        Dim service As ExchangeService = TACMAIN1.Get_EWS_Service(AS_PARM_EMAIL_USER_ID)

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
                    Dim rowTATCONT1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", SEND_TO_email_address.ToLower)
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
                    Dim rowTATCONT1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", SEND_CC_email_address.ToLower)
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

        If chkBCC.Checked Then
            Message.BccRecipients.Add(New EmailAddress(ASCMAIN1.USER_NAME, ASCMAIN1.USER_EMAIL))
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



        If EMAIL IsNot Nothing Then
            If EMAIL.Attachments.Count > 0 Then
                For Each MA As System.Net.Mail.Attachment In EMAIL.Attachments
                    'Message.Attachments.AddItemAttachment(MA)
                Next
            End If
        End If

        If rowTATMAIL1 Is Nothing Then
            rowTATMAIL1 = LookUp("TATMAIL1", EMAIL_KEY)
            ' IMPORTANT - emails will NOT work without a record in TATMAIL1 showing an email account with Send-As privileges.
            ' sometimes, when wjz sends the email, it works - but that is because usually wjz's ASTUSER1 record reflects an email user id that has send-as privileges
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

        ' Done since the data is stored n a table anfthe data is truncated to fit in the field
        If SEND_BODY_OVERRIDE.Length > 0 Then
            SEND_BODY = SEND_BODY_OVERRIDE
        End If

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



        Dim retry As Boolean = False
        Dim retries As Integer = 0
        Do
            Try
                If retry Then retries += 1
                Message.Save()
                retry = False
            Catch ex As Exception
                If ex.Message = "The request failed. The remote server returned an error: (401) Unauthorized." And Not retry Then
                    ASCMAIN1.absTimer.Enabled = False
                    ASCMAIN1.absTimer.Interval = 1000
                    ASCMAIN1.absTimer.Enabled = True
                    System.Threading.Thread.Sleep(10000)
                    service.Credentials = New OAuthCredentials(ASCMAIN1.authResult.AccessToken)
                    retry = True
                Else
                    ' what now batman?
                    'Throw New Exception("Problem with email Token - email not sent")
                    ' wierd behavior, we rebounded from the End If into the Else once we changed the value of retry
                End If
            End Try

            If retry And retries >= 1 Then
                Throw New Exception("Problem with email Token - email not sent")
            End If
        Loop While retry




        'Message.SendAndSaveCopy()

        Try
            Message.SaveToFile(folder & SEND_NO & ".eml")
        Catch ex As Exception
            If ASCMAIN1.CLIENT <> "INT" Then
                MsgBox("email file not saved")
            End If
        End Try

        ' moved this line down because in dev the emails were not getting saved if we tried to SaveToFile after SendAndSaveCopy
        Message.SendAndSaveCopy()

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

    Private Function Send_fax(Optional ByVal auto_send As Boolean = False) As Boolean

        'Try

        '    SEND_TO = Trim(SEND_TO)
        '    Dim faxnumber As String = String.Empty
        '    Dim zMsg As String = String.Empty

        '    If SEND_TO = "" Then
        '        MessageBox.Show("Hmmm, what fax number should this fax be sent. I do not know, do you?", "Fax", MessageBoxButtons.OK, MessageBoxIcon.Question)
        '        Return False
        '    Else
        '        For Each ch As Char In SEND_TO
        '            If Char.IsDigit(ch) Then
        '                faxnumber &= ch
        '            End If
        '        Next

        '        Select Case Len(faxnumber)
        '            Case 7, 10
        '                ' Should be a good number
        '            Case 11
        '                If Not faxnumber.StartsWith("1") Then
        '                    zMsg = "The provided fax number (" & faxnumber & ") is 11 characters and does not begin with a '1'."
        '                    zMsg &= Environment.NewLine & "Do you want to proceed?"
        '                End If
        '            Case Else
        '                zMsg = "The provided fax number (" & faxnumber & ") does not appear to be a valid telephone number ."
        '                zMsg &= Environment.NewLine & "Do you want to proceed?"
        '        End Select
        '    End If

        '    SEND_TO = faxnumber

        '    If zMsg <> "" Then
        '        If MessageBox.Show(zMsg, "Fax", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
        '            Return False
        '        End If
        '    End If


        '    Dim fax As New TAC.TACFAXS1

        '    fax.fax_Username = ASCMAIN1.rowASTPARM1.Item("AS_PARM_EFAX_USERNAME")
        '    fax.fax_Password = ASCMAIN1.rowASTPARM1.Item("AS_PARM_EFAX_PASSWORD")
        '    fax.fax_CoverFile = ASCMAIN1.Folders("Archive") & "eFax\Cover.rtf"
        '    fax.fax_FaxAttachment = SEND_ATTACHMENT

        '    fax.SEND_BODY = SEND_BODY
        '    fax.SEND_CODE = SEND_ENTITY_KEY
        '    fax.SEND_FROM = SEND_FROM
        '    fax.SEND_FROM_NAME = SEND_FROM_NAME
        '    fax.SEND_NAME = SEND_ENTITY_NAME
        '    fax.SEND_SUBJECT = SEND_SUBJECT
        '    fax.SEND_TO = SEND_TO
        '    fax.SEND_TO_NAME = SEND_TO_NAME
        '    fax.SendFax()

        '    SEND_STATUS = "S"
        '    SEND_LOG = fax.fax_log.ToString
        '    SEND_ID = fax.fax_transportID
        '    SEND_NO = ASCMAIN1.Next_Control_No("TATSEND1.SEND_NO")

        '    If optType.Value = "F" Then
        '        Dim FILENAME As String = ASCMAIN1.Folders("Archive") & "eFax\Logs\" & SEND_NO & ".txt"
        '        If My.Computer.FileSystem.FileExists(FILENAME) Then
        '            My.Computer.FileSystem.DeleteFile(FILENAME)
        '        End If
        '        Using SW As New System.IO.StreamWriter(FILENAME)
        '            SW.Write(SEND_LOG)
        '        End Using
        '    End If

        '    If Not auto_send Then
        '        MsgBox("fax has been sent", MsgBoxStyle.OkOnly, "Verification")
        '    End If
        '    Return True

        'Catch ex As Exception
        '    SEND_STATUS = "E"
        '    If Not auto_send Then
        '        MsgBox("Error Occured: " & ex.Message, MsgBoxStyle.OkOnly, "Could not Send fax")
        '    End If
        '    Return False
        'Finally

        'End Try

    End Function

    Private Sub cmdSend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSend.Click
        Me.Cursor = Cursors.WaitCursor

        Screen_Fields("Unload")

        If optType.Value = "E" Then
            If Not Send_email() Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
        Else
            If Not Send_fax() Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
        End If

        Update_Send_Log()

        Me.Cursor = Cursors.Default
        Me.Close()
    End Sub

    Sub Prepare_Send_Log()
        With dst
            Create_TDA(.Tables.Add, "TATSEND1", "*")
        End With

        rowTATSEND1 = dst.Tables("TATSEND1").NewRow
        rowTATSEND1.Item("SEND_NO") = "0000000000"
        rowTATSEND1.Item("INIT_DATE") = DATETIME_STAMP
        dst.Tables("TATSEND1").Rows.Add(rowTATSEND1)

        If Me.Visible Then
            Screen_Fields("Load")
        End If
    End Sub

    Sub Update_Send_Log()

        rowTATSEND1.Item("SEND_NO") = SEND_NO
        INIT_LAST("TATSEND1", True)
        If optType.Value = "F" Then
            rowTATSEND1.Item("SEND_ID") = SEND_ID
        End If
        Update_Record_TDA("TATSEND1")

        If optType.Value = "F" Then
            Dim FILENAME As String = ASCMAIN1.Folders("Archive") & "eFax\Logs\" & SEND_NO & ".txt"
            If My.Computer.FileSystem.FileExists(FILENAME) Then
                My.Computer.FileSystem.DeleteFile(FILENAME)
            End If
            Using SW As New System.IO.StreamWriter(FILENAME)
                SW.Write(SEND_LOG)
            End Using
        End If


    End Sub

    Private Sub optType_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optType.ValueChanged
        Setup_SEND_TYPE()
    End Sub

    Sub Setup_SEND_TYPE()

        If Not setupScreen Then Exit Sub

        chkBCC.Visible = (optType.Value = "E")

        '     Absx1.txtFor("SEND_CC_NAME").Visible = (optType.Value = "E")
        Absx1.txtFor("SEND_CC").Visible = (optType.Value = "E")
        cmdCC.Visible = (optType.Value = "E")
        lblSignature.Visible = (optType.Value = "E")
        Absx1.txtFor("SEND_FROM_SIGNATURE").Visible = (optType.Value = "E")

        If optType.Value = "E" Then
            lblMethod.Text = "email address"
        Else
            lblMethod.Text = "Fax Number"
        End If

    End Sub

    Sub Add_Contacts(txt As UltraWinEditors.UltraTextEditor)

        Using frmTAFCONT1 As New TAFCONT1(frmASFBASE1)
            With frmTAFCONT1
                .CONTACT_ENTITY_TABLE = SEND_ENTITY_TABLE
                .CONTACT_ENTITY_KEY = SEND_ENTITY_KEY
                .CONTACT_ENTITY_NAME = SEND_ENTITY_NAME
                .ShowDialog()
                If .result Then
                    Dim SEND_TO As String = txt.Text
                    For Each row As DataRow In .dst.Tables("TATCONT1").Select("SEL = '1'")
                        Dim CONTACT_EMAIL As String = row.Item("CONTACT_EMAIL") & ""
                        If InStr(";" & Replace(Trim(SEND_TO).ToLower, " ", "") & ";", ";" & CONTACT_EMAIL.ToLower & ";") = 0 Then
                            If Trim(SEND_TO) <> "" Then
                                SEND_TO &= ";"
                            End If
                            SEND_TO &= CONTACT_EMAIL
                        End If
                        'Add_SEND_TO(row.Item("CONTACT_EMAIL") & "", row.Item("CONTACT_NAME") & "")
                    Next
                    txt.Text = SEND_TO
                End If
            End With
        End Using
    End Sub
    Private Sub cmdTo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTo.Click

        Add_Contacts(txtSEND_TOS)

    End Sub

    Private Sub cmdCC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCC.Click

        Add_Contacts(txtSEND_CC)

        'ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("CONTACT_NO")
        'If ASCMAIN1.CodeSelector.SQL <> "" Then
        '    ASCMAIN1.CodeSelector.MultipleSelections = False
        '    ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = CODE_VALUE
        '    Dim F As New ASFCODE1
        '    F.ShowDialog()
        '    F.Dispose()
        '    If ASCMAIN1.CodeSelector.Selections <> 0 Then
        '        txtSEND_CC.Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("CONTACT_EMAIL")
        '    End If
        'End If
    End Sub

    Public Sub Add_SEND_TO(ByVal SEND_TO_to_add As String, ByVal SEND_TO_NAME_to_add As String)

        If SEND_TOs.ContainsKey(SEND_TO_to_add) Then
            SEND_TOs(SEND_TO_to_add) = SEND_TO_NAME_to_add
        Else
            SEND_TOs.Add(SEND_TO_to_add, SEND_TO_NAME_to_add)
        End If

        SEND_TO = ""
        'SEND_TO_NAME = ""

        For Each SEND_TO_KEY In SEND_TOs.Keys
            SEND_TO &= ";" & SEND_TO_KEY
            'SEND_TO_NAME &= ";" & SEND_TOs(SEND_TO_KEY)
        Next

        If SEND_TO.StartsWith(";") Then
            SEND_TO = Mid(SEND_TO, 2)
        End If
        'If SEND_TO_NAME.StartsWith(";") Then
        '    SEND_TO_NAME = Mid(SEND_TO_NAME, 2)
        'End If

        txtSEND_TO.Text = SEND_TO
        'txtSEND_TO_NAME.Text = SEND_TO_NAME
    End Sub

    Private Sub txtSEND_TO_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSEND_TO.ValueChanged
        If InStr(txtSEND_TO.Text, ";") = 0 Then
            txtSEND_TO_NAME.Visible = True
            txtSEND_TO.Width = txtSEND_TO_NAME.Left
            txtSEND_TO_NAME.Text = Get_Names("SEND_TO", "SEND_TO_NAME")
        Else
            txtSEND_TO.Width = txtSEND_TO_NAME.Left + txtSEND_TO_NAME.Width - txtSEND_TO.Left
            txtSEND_TO_NAME.Visible = False
            txtSEND_TO_NAME.Text = ""
        End If
    End Sub

    Private Sub cmdView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdView.Click
        Dim SEND_ATTACHMENT As String = Absx1.txtFor("SEND_ATTACHMENT").Text

        If SEND_ATTACHMENT = "" Then
            MsgBox("Nothing to View", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
        Else
            For Each FILENAME In Split(SEND_ATTACHMENT, ";")
                If SEND_ATTACHMENTs Is Nothing Then
                    Show_Document(FILENAME)
                Else
                    Show_Document(SEND_ATTACHMENTs(FILENAME))
                End If

            Next
        End If
    End Sub

    'Private Sub txtSEND_CC_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSEND_CC.ValueChanged
    '    txtSEND_CC_NAME.Text = Get_Names("SEND_CC", "SEND_CC_NAME")
    'End Sub

    Function Get_Names(ByVal email_Address_Column As String, ByVal Email_Name_Column As String) As String
        Get_Names = ""
        Dim Email_Address As String = Absx1.txtFor(email_Address_Column).Text

        If Email_Address <> "" Then
            For Each SEND_email_address As String In Split(Absx1.txtFor(email_Address_Column).Text, ";")
                SEND_email_address = Trim(SEND_email_address)
                ASCMAIN1.sql = "Select * from TATCONT1 " _
                & " where LOWER(CONTACT_EMAIL) = :PARM1"
                Dim rowTATCONT1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", SEND_email_address.ToLower)
                If rowTATCONT1 IsNot Nothing Then
                    Get_Names &= rowTATCONT1.Item("CONTACT_NAME")
                End If

                ASCMAIN1.sql = "Select * from ASTUSER1 " _
                & "  where USER_EMAIL = :PARM1"
                Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {SEND_email_address.ToLower})
                If row IsNot Nothing Then
                    Get_Names &= row.Item("USER_NAME") & "" & "; "
                End If
            Next
        End If
        Return Get_Names
    End Function

End Class