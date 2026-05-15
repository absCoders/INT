Imports System
Imports System.IO
Imports System.Net.Sockets
Imports System.Text
Imports System.Xml

''' <summary>
''' This class is used to Post an XML Request and Accept an XML Response
''' </summary>
''' <remarks></remarks>
Public Class ASCSOCK1

#Region "Class Variables"

    Private _SendIPAddress As String = String.Empty
    Private _SendPort As String = String.Empty
    Private _TimeOut As Integer = 0
    Private _RequestResponse As String = String.Empty

    Private Const defaultTimeOut As Integer = 15
    Private Const iBufferlen As Integer = 4096

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        _SendIPAddress = String.Empty
        _SendPort = String.Empty
        _TimeOut = defaultTimeOut
        _RequestResponse = String.Empty
    End Sub

    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="RemoteHostIP">Host IP Address</param>
    ''' <param name="RemoteHostPort">Host Port</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal RemoteHostIP As String, ByVal RemoteHostPort As String)

        _SendIPAddress = ValidateIPAddress(RemoteHostIP)
        _SendPort = CStr(Val(RemoteHostPort))
        _RequestResponse = String.Empty

        ' Timeout in seconds converted to Milliseconds
        Try
            _TimeOut = defaultTimeOut
        Catch ex As Exception
            _TimeOut = defaultTimeOut
        End Try

        If _TimeOut <= 0 Then _TimeOut = defaultTimeOut

    End Sub

#End Region

#Region "Class Properties"

    ''' <summary>
    ''' Get / Set IP Address
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property RemoteHostIP() As String
        Get
            Return _SendIPAddress
        End Get

        Set(ByVal value As String)
            _SendIPAddress = ValidateIPAddress(value)
        End Set
    End Property

    ''' <summary>
    ''' Set / Get Remote Host Port 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property RemoteHostPort() As String
        Get
            Return _SendPort
        End Get

        Set(ByVal value As String)
            _SendPort = value
        End Set
    End Property

    ''' <summary>
    ''' Get / Set Timeout in Seconds
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Timeout() As Integer
        Get
            Return _TimeOut
        End Get

        Set(ByVal value As Integer)
            _TimeOut = value
        End Set
    End Property

    ''' <summary>
    ''' Returns the Request's Response.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property RequestResponse() As String
        Get
            Return _RequestResponse
        End Get
    End Property

    ''' <summary>
    ''' Return in the IP Address is valid
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private ReadOnly Property IsValidIPAddress() As Boolean
        Get
            Return _SendIPAddress <> String.Empty
        End Get
    End Property

    ''' <summary>
    ''' Validates a Port Number. Port needs to be bertween 1 and 15000
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private ReadOnly Property IsValidPort() As Boolean
        Get
            Return Val(_SendPort) >= 1 And Val(_SendPort) <= 15000
        End Get
    End Property

#End Region

#Region "Class Private Functions"

    ''' <summary>
    ''' Validates an IP Address
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ValidateIPAddress(ByVal value As String) As String
        Dim IP_Comps() As String = Split(value, ".")
        Dim i As Integer
        Dim IP_Address As String = String.Empty

        ValidateIPAddress = String.Empty

        If IP_Comps.Length <> 4 Then Exit Function

        For i = 0 To 3
            If i = 0 And Val(IP_Comps(i).Trim) = 0 Then Exit Function
            If IP_Comps(i).Trim = String.Empty Then Exit Function
            If Not IsNumeric(IP_Comps(i).Trim) Then Exit Function
            If Val(IP_Comps(i).Trim) > 255 Or Val(IP_Comps(i).Trim) < 0 Then Exit Function

            IP_Address += "." & IP_Comps(i).Trim
        Next

        Return IP_Address.Substring(1)
    End Function

#End Region

#Region "Class Public Functions"
    ''' <summary>
    ''' Posts a Request to an IP Address and Port
    ''' </summary>
    ''' <param name="Post_String">Sting to Post</param>
    ''' <param name="ShutDownOutbound">Should the connection close the outbound after the send</param>
    ''' <param name="WaitForResponse">Wait for a response from the host</param>
    ''' <returns>True if successful; otherwise, false</returns>
    ''' <remarks></remarks>
    Public Function TCPPost(ByVal Post_String As String, ByVal ShutDownOutbound As Boolean, ByVal WaitForResponse As Boolean) As Boolean

        TCPPost = False
        _RequestResponse = String.Empty

        If IsValidIPAddress And IsValidPort Then

            Try
                ' Obtain a Socket Connection to the Server
                Dim ssSocket As New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)

                Dim hostadd As System.Net.IPAddress = System.Net.Dns.GetHostAddresses(_SendIPAddress)(0)
                Dim EPhost As System.Net.IPEndPoint = New System.Net.IPEndPoint(hostadd, CInt(_SendPort))
                ssSocket.Connect(EPhost)

                ' Send Request to Server
                Dim requestAsByte As [Byte]() = Encoding.ASCII.GetBytes(Post_String)
                Try
                    ssSocket.Send(requestAsByte)
                Catch ex As Exception
                    _RequestResponse = ex.Message
                    'Exit Function
                End Try

                If ShutDownOutbound = True Then
                    ' Shutdown client outbound socket request
                    Try
                        ssSocket.Shutdown(SocketShutdown.Send)
                    Catch ex As Exception
                        _RequestResponse = ex.Message
                        Exit Function
                    End Try
                End If

                If WaitForResponse = False Then
                    Return True
                    Exit Function
                End If

                ' Receive response from server
                Dim vTimeOut As Integer = 0
                Try
                    vTimeOut = Timeout * 1000000
                Catch ex As Exception
                    vTimeOut = 15000000
                End Try
                If Not ssSocket.Poll(vTimeOut, SelectMode.SelectRead) Then
                    _RequestResponse = "Could not read from remote host. Timeout Expired"
                    Exit Function
                End If

                ' Receive Server Response
                Dim buffer As Byte()
                ReDim buffer(iBufferlen)

                Dim iRead As Integer = 1
                Dim stResponse As MemoryStream = New MemoryStream

                Try
                    While iRead > 0
                        iRead = ssSocket.Receive(buffer)
                        If iRead > 0 Then stResponse.Write(buffer, 0, iRead)
                    End While
                Catch ex As Exception
                    _RequestResponse = "Could not read from remote host. Timeout Expired"
                    Exit Function
                End Try

                ' Convert response to plain text
                Dim response As String = String.Empty
                Try
                    response = Encoding.ASCII.GetString(stResponse.GetBuffer(), 0, CInt(stResponse.Length))
                Catch ex As Exception
                    _RequestResponse = "Could not read from remote host. Timeout Expired"
                    Exit Function
                End Try

                _RequestResponse = response
                TCPPost = True

            Catch ex As Exception
                _RequestResponse = ex.Message
                Exit Function
            End Try

            TCPPost = True
        Else
            _RequestResponse = "Missing or Invalid Remote IP Address or Port"
            Exit Function
        End If

    End Function

#End Region

End Class
