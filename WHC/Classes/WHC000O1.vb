Imports nsoftware.IPWorksSSH

Public Class WHC000O1
    ' Base Class for Outbound Files to 3PL

    Inherits ASCBASE0

    Public sftp_folder As String = String.Empty
    Public Const sep As String = "," ' vbTab
    Public Const quo As String = Chr(34)
    Public R As Int64 = 0
    Public XMIT_NO As String = ASCMAIN1.Next_Control_No("WHT3PLX1.XMIT_NO")
    Public DTS As String = Format(DATETIME_STAMP, "yyyyMMddHHmmss")
    Protected clsErrorMessage As List(Of String)
    Protected clsSuccessfulExecution As Boolean
    Public LP_CODE As String = ""
    Public FILENAME_TO_SEND As String = ""
    Public theLog As String = ""
    Public Shared tblTasks As DataTable

    Sub New()
        MyBase.New()
    End Sub

    Sub New(g As ABSEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHC000O1"

        tblTasks = New DataTable
        With tblTasks
            .Columns.Add("SEQ_NO", GetType(Int32))
            .Columns.Add("TASK_TIME", GetType(DateTime))
            .Columns.Add("TASK_DESC", GetType(String))
        End With

        sftp_folder = "" _
            & IIf(ASCMAIN1.Running_in_VS, "C:\Temp\", ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT") & "\cfg\") _
              & IIf(g.DBS_SERVER = "INTTST", "TEST", "PROD") _
            & "\FROM_IPLB\"

        If ASCMAIN1.Running_in_VS Then
            Stop
            sftp_folder = ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT") & "\cfg\" & IIf(g.DBS_SERVER = "INTTST", "TEST", "PROD") & "\FROM_IPLB\"
        End If

        If ASCMAIN1.Running_in_VS Then
            If Not My.Computer.FileSystem.DirectoryExists(sftp_folder) Then
                My.Computer.FileSystem.CreateDirectory(sftp_folder)
            End If
        End If

        LP_CODE = g.LP_CODE
        ASCMAIN1.CLIENT = g.CLIENT

        With dst
            With .Tables.Add("WHTEVNT1").Columns
                .Add("EVENT_LOG")
            End With
        End With

        tbl = dst.Tables("WHTEVNT1") ' Log Events
        clsErrorMessage = New List(Of String)
        clsSuccessfulExecution = False
    End Sub

    Sub Update_Record()

        ' Some classes reuse this class and the XMIT_NO does not get updated. causing a unique key error on the insert into WHT3PLX1
        Static eventRecorded As Boolean = False

        BeginTrans()
        Update_Create_File()

        If Not eventRecorded Then
            Try
                ASCDATA1.ExecuteSQL("Insert into WHT3PLX1 Values (:PARM1,:PARM2,:PARM3,:PARM4,:PARM5,:PARM6)",
                                    "VDVDVN",
                                    New Object() {XMIT_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, DATETIME_STAMP, MENU_ITEM_OBJECT, R})
            Catch ex As Exception
            End Try
            eventRecorded = True
        End If

        If R = 0 Then
            ' DO NOTHING
        Else
            Update_Archive()
        End If

        CommitTrans()

        If R > 0 Then
            Post_Update_Archive()
        End If
    End Sub

    Overridable Sub Update_Create_File()

    End Sub

    Overridable Sub Update_Archive()

    End Sub

    Overridable Sub Post_Update_Archive()

    End Sub

    Function SpaceIfNull(t As String) As String
        If t = "" Then t = " "
        Return t
    End Function

    Public ReadOnly Property ErrorMessages As List(Of String)
        Get
            Return clsErrorMessage
        End Get

    End Property

    Public ReadOnly Property SuccessfulExecution As Boolean
        Get
            Return clsSuccessfulExecution
        End Get
    End Property

    Function sftp_put(
        SSH_APP_CODE As String,
        production As Boolean,
        FILENAME_LOCAL As String,
        FILENAME_REMOTE As String) As Boolean

        ' Added 09/26/2025 to prevent sending test environment data
        If ASCMAIN1.DBS_COMPANY <> ASCMAIN1.CLIENT OrElse ASCMAIN1.DBS_SERVER <> ASCMAIN1.CLIENT Then
            If ASCMAIN1.Running_in_VS Then
                Stop
            Else
                MessageBox.Show("You are not in production. sFTP File transfer avoided.", "Update Archive", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return True
            End If
        End If

        Dim rowTATSSHK1 As DataRow = LookUp("TATSSHK1", SSH_APP_CODE)

        '' SHOULD BE USING EXP COMPANY FOR A&E
        'rowTATSSHK1 = ASCDATA1.GetDataRow("Select * from TATSSHK1 where SSH_APP_CODE = '" & SSH_APP_CODE & "'")

        Dim SSH_APP_USERNAME As String = rowTATSSHK1.Item("SSH_APP_USERNAME") & ""
        Dim SSH_APP_PASSWORD As String = rowTATSSHK1.Item("SSH_APP_PASSWORD") & ""
        Dim SSH_APP_FOLDER_PUT As String = rowTATSSHK1.Item("SSH_APP_FOLDER_PUT") & ""
        'If SSH_APP_FOLDER_PUT.EndsWith("\") Then

        'End If

        Dim SSH_APP_PARTNER_URI As String = ""
        Dim SSH_APP_PARTNER_PUBKEY As String = ""
        If production Then
            SSH_APP_PARTNER_URI = rowTATSSHK1.Item("SSH_APP_PARTNER_URI_PROD") & ""
            SSH_APP_PARTNER_PUBKEY = rowTATSSHK1.Item("SSH_APP_PARTNER_PUBKEY_PROD") & ""
        Else
            SSH_APP_PARTNER_URI = rowTATSSHK1.Item("SSH_APP_PARTNER_URI_TEST") & ""
            SSH_APP_PARTNER_PUBKEY = rowTATSSHK1.Item("SSH_APP_PARTNER_PUBKEY_TEST") & ""
        End If

        Dim SSH_APP_PORT As Integer = Val(rowTATSSHK1.Item("SSH_APP_PORT") & "")
        If SSH_APP_PORT = 0 Then
            SSH_APP_PORT = 22
        End If

        Dim success = False
        Dim sftp As New nsoftware.IPWorksSSH.SFTPClient
        theLog = ""

        AddHandler sftp.OnSSHServerAuthentication, AddressOf SSHServerAuthentication
        AddHandler sftp.OnSSHStatus, AddressOf SSHStatus

        sftp.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwaresftpkey")


        sftp.SSHUser = SSH_APP_USERNAME

        If SSH_APP_PASSWORD <> "" Then
            sftp.SSHAuthMode = SCPSSHAuthModes.amPassword
            sftp.SSHPassword = SSH_APP_PASSWORD
        Else
            sftp.SSHAuthMode = SCPSSHAuthModes.amPublicKey
            'sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "C:\Users\wjz\Desktop\Interparfums\JPMC\JPMC_SSH_pvt.ppk", "0ff1c3INT", "*")
            'sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "S:\INT\Archive\INT\JPMC\JPMC_SSH_pvt.ppk", "0ff1c3INT", "*")

            If ASCMAIN1.Running_in_VS Then
                Stop
                sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "C:\VS\AHA\Archive\INT\JPMC\JPMC_IPLB_pvt.asc", "0ff1c3INT", "*")
                'sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "C:\Users\wjz\Desktop\Interparfums\JPMC\JPMC_SSH_pvt.ppk", "0ff1c3INT", "*")
            Else
                ' sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "S:\INT\Archive\INT\JPMC\JPMC_IPLB_pvt.asc", "0ff1c3INT", "*")
                'sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "S:\INT\Archive\INT\JPMC\JPMC_SSH_pvt.ppk", "0ff1c3INT", "*")
                Dim ssh_file As String = ASCMAIN1.Folders("SharedRoot") & "Archive\INT\JPMC\JPMC_SSH_pvt.ppk"
                sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, ssh_file, "0ff1c3INT", "*")
            End If

        End If

        Try

            If sftp.Connected = True Then
                sftp.SSHLogoff()
            End If

            If ASCMAIN1.CLIENT = "INT" Then
                If SSH_APP_CODE = "JPMC" Then
                    sftp.SSHEncryptionAlgorithms = "aes128-ctr,aes192-ctr,aes256-ctr"
                    sftp.Config("LogSSHPackets=True")
                    If ASCMAIN1.USER_ID = "wjz" Then MsgBox(sftp.Config("LogSSHPackets"))
                Else
                    ' COWORX DOES NOT SUPPORT NEW ENCRYPTION
                End If
            End If

            sftp.SSHHost = SSH_APP_PARTNER_URI
            sftp.SSHLogon(SSH_APP_PARTNER_URI, SSH_APP_PORT)
            success = True

            sftp.LocalFile = FILENAME_LOCAL
            sftp.ChangeRemotePath(SSH_APP_FOLDER_PUT)
            'sftp.RemotePath = SSH_APP_FOLDER_PUT

            sftp.RemoteFile = FILENAME_REMOTE
            sftp.Upload()

        Catch ex As Exception
            theLog &= ex.Message
            Dim filename As String = Format(Now, "yyyyMMddhhhhss")
            System.IO.File.WriteAllText(ASCMAIN1.Folders("Work") & filename & ".log", theLog)
            MessageBox.Show(ex.Message, "Secure ftp Setup", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If sftp.Connected = True Then
                sftp.SSHLogoff()
            End If
        End Try

        If sftp.Connected = True Then
            sftp.SSHLogoff()
        End If

        Return success

    End Function

    Public Function sftp_get(
        SSH_APP_CODE As String,
        production As Boolean,
        FILENAME_LOCAL As String,
        FILENAME_REMOTE As String) As List(Of String)

        Dim FILENAMEs As New List(Of String)

        Dim rowTATSSHK1 As DataRow = LookUp("TATSSHK1", SSH_APP_CODE)

        Dim SSH_APP_USERNAME As String = rowTATSSHK1.Item("SSH_APP_USERNAME") & ""
        Dim SSH_APP_PASSWORD As String = rowTATSSHK1.Item("SSH_APP_PASSWORD") & ""
        Dim SSH_APP_FOLDER_GET As String = rowTATSSHK1.Item("SSH_APP_FOLDER_GET") & ""
        'If SSH_APP_FOLDER_PUT.EndsWith("\") Then

        'End If

        Dim SSH_APP_PARTNER_URI As String = ""
        Dim SSH_APP_PARTNER_PUBKEY As String = ""
        If production Then
            SSH_APP_PARTNER_URI = rowTATSSHK1.Item("SSH_APP_PARTNER_URI_PROD") & ""
            SSH_APP_PARTNER_PUBKEY = rowTATSSHK1.Item("SSH_APP_PARTNER_PUBKEY_PROD") & ""
        Else
            SSH_APP_PARTNER_URI = rowTATSSHK1.Item("SSH_APP_PARTNER_URI_TEST") & ""
            SSH_APP_PARTNER_PUBKEY = rowTATSSHK1.Item("SSH_APP_PARTNER_PUBKEY_TEST") & ""
        End If

        Dim success = False
        Dim sftp As New nsoftware.IPWorksSSH.SFTPClient
        AddHandler sftp.OnSSHServerAuthentication, AddressOf SSHServerAuthentication

        sftp.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwaresftpkey")

        sftp.SSHUser = SSH_APP_USERNAME

        If SSH_APP_PASSWORD <> "" Then
            sftp.SSHAuthMode = SCPSSHAuthModes.amPassword
            sftp.SSHPassword = SSH_APP_PASSWORD
        Else
            sftp.SSHAuthMode = SCPSSHAuthModes.amPublicKey
            'sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "C:\Users\wjz\Desktop\Interparfums\JPMC\JPMC_SSH_pvt.ppk", "0ff1c3INT", "*")
            'sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "S:\INT\Archive\INT\JPMC\JPMC_SSH_pvt.ppk", "0ff1c3INT", "*")
            'sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "S:\INT\Archive\INT\JPMC\JPMC_IPLB_pvt.asc", "0ff1c3INT", "*")
            sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, ASCMAIN1.Folders("SharedRoot") & "\Archive\INT\JPMC\JPMC_IPLB_pvt.asc", "0ff1c3INT", "*")
        End If

        Try

            If sftp.Connected = True Then
                sftp.SSHLogoff()
            End If

            sftp.SSHHost = SSH_APP_PARTNER_URI
            sftp.SSHLogon(SSH_APP_PARTNER_URI, 22)
            success = True
            sftp.ChangeRemotePath("/" & SSH_APP_FOLDER_GET)
            'sftp.RemotePath = "/" & SSH_APP_FOLDER_GET

            sftp.ListDirectory()
            For Each s As nsoftware.IPWorksSSH.DirEntry In sftp.DirList
                sftp.RemoteFile = s.FileName
                If Not s.IsDir Then
                    'ASCMAIN1.Progress("-", s.FileName)
                    sftp.LocalFile = FILENAME_LOCAL & s.FileName
                    sftp.Download()
                    '  sftp.RenameFile(FILENAME_LOCAL & "\Archive\" & s.FileName)

                    sftp.DeleteFile(s.FileName)
                    FILENAMEs.Add(FILENAME_LOCAL & s.FileName)
                End If
            Next

            sftp.SSHLogoff()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Secure ftp Setup", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If sftp.Connected = True Then
                sftp.SSHLogoff()
            End If
        End Try

        If sftp.Connected = True Then
            sftp.SSHLogoff()
        End If

        Return FILENAMEs ' success
    End Function

    Public Shared Sub SSHServerAuthentication(sender As Object, e As SFTPClientSSHServerAuthenticationEventArgs) ' SCPSSHServerAuthenticationEventArgs)

        e.Accept = True
    End Sub

    Sub SSHStatus(sender As Object, e As SFTPClientSSHStatusEventArgs) ' SCPSSHStatusEventArgs)

        ' MsgBox(e.Message, MsgBoxStyle.OkOnly, "SSHStatus Messages")
        theLog &= e.Message & vbCrLf

    End Sub

    Function StrToByteArray(ByVal str As String) As Byte()
        Dim encoding As New System.Text.UTF8Encoding()
        Return encoding.GetBytes(str)
    End Function

    Public Shared Sub Addtask(ByVal TaskDescription As String)
        tblTasks.Rows.Add({tblTasks.Rows.Count + 1, DateTime.Now, TaskDescription})
    End Sub


End Class