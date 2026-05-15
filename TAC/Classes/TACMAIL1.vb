Imports System.IO
Imports System.Reflection
Imports System.Net.Mail
Imports Microsoft.Exchange.WebServices.Data

Public Module TACMAIL1
    Sub New()
    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Save(ByVal Message As MailMessage, ByVal FileName As String)
        Dim assembly As Assembly = GetType(SmtpClient).Assembly
        Dim _mailWriterType As Type = assembly.[GetType]("System.Net.Mail.MailWriter")

        Using _fileStream As New FileStream(FileName, FileMode.Create)
            ' Get reflection info for MailWriter contructor
            Dim _mailWriterContructor As ConstructorInfo = _mailWriterType.GetConstructor(BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, New Type() {GetType(Stream)}, Nothing)

            ' Construct MailWriter object with our FileStream
            Dim _mailWriter As Object = _mailWriterContructor.Invoke(New Object() {_fileStream})

            ' Get reflection info for Send() method on MailMessage
            Dim _sendMethod As MethodInfo = GetType(MailMessage).GetMethod("Send", BindingFlags.Instance Or BindingFlags.NonPublic)

            ' Call method passing in MailWriter
            If ASCMAIN1.Running_in_VS Then
                _sendMethod.Invoke(Message, BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, New Object() {_mailWriter, True, True}, Nothing)
            Else
                _sendMethod.Invoke(Message, BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, New Object() {_mailWriter, True}, Nothing)
            End If
            
            ' Finally get reflection info for Close() method on our MailWriter
            Dim _closeMethod As MethodInfo = _mailWriter.[GetType]().GetMethod("Close", BindingFlags.Instance Or BindingFlags.NonPublic)

            ' Call close method
            _closeMethod.Invoke(_mailWriter, BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, New Object() {}, Nothing)
        End Using
    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub SaveToFile(ByVal Message As EmailMessage, ByVal FileName As String)
        Message.Load(New PropertySet(ItemSchema.MimeContent))
        Dim mimcon As MimeContent = Message.MimeContent
        Using fStream As New FileStream(FileName, FileMode.Create)
            fStream.Write(mimcon.Content, 0, mimcon.Content.Length)
            fStream.Close()
        End Using
    End Sub
End Module
