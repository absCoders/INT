Imports nsoftware.IPWorksSSH
Imports System.Security.Cryptography
Public Class TACSCOM1

    Public Shared theLog As String
    'Public Shared SSH_APP_SSH_PVTKEY_B() As Byte

    Public Shared Function ftp_Files( _
        FOLDERNAME_local As String, _
        FILENAME_local() As String, _
        FOLDERNAME_remote As String, _
        FILENAME_remote() As String, _
        USER As String, _
        PWD As String, _
        HOST As String) As Boolean

        Try
            Dim Ftp1 As New nsoftware.IPWorks.Ftp
            Ftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")

            Ftp1.User = USER
            Ftp1.Password = PWD
            Ftp1.RemoteHost = HOST
            Ftp1.Logon()
            Ftp1.ChangeTransferMode(nsoftware.IPWorks.FTPTransferModes.tmBinary)
            Ftp1.ChangeRemotePath(FOLDERNAME_remote)

            For i As Integer = 0 To FILENAME_local.Length - 1
                'Ftp1.LocalFile = ASCMAIN1.Folders("Temp") & FILENAME_local
                Ftp1.LocalFile = FOLDERNAME_local & FILENAME_local(i)
                Ftp1.RemoteFile = FILENAME_remote(i)
                Ftp1.Upload()
            Next

            Ftp1.Logoff()

            Return True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error ftp'ing File")
            Return False

        End Try
    End Function

    Public Shared Sub SSHServerAuthentication(sender As Object, e As SFTPClientSSHServerAuthenticationEventArgs) '  As SCPSSHServerAuthenticationEventArgs)
        ' e.Fingerprint = e0:62:98:74:ca:18:e2:ed:93:d5:1e:30:90:75:07:06
        '                 96:25:1f:a6:b1:45:27:b2:66:73:3d:89:9b:b5:fa:15:d9:b4:95:6a:0f:12:83:8a:b7:69:36:a3:6d:c2:98:67
        '                 aa:8a:f8:5a:dd:bf:67:b0:22:1c:b9:bb:a7:3f:65:6d:1b:6e:4b:6f:87:83:ab:60:32:ca:ee:b8:9d:ba:f5:fd
        ' e.KeyAlgorithm = ssh-rsa
        ' e.HostKeyB  = 
        e.Accept = True ' NOT NECESSARY SINCE WE USED SSHAcceptServerHostKey 
    End Sub

    Public Shared Sub SSHStatus(sender As Object, e As SFTPClientSSHStatusEventArgs) ' SCPSSHStatusEventArgs)
        ' Sending local version: "SSH-2.0-IP*Works! SSH Client v9.0".
        ' MsgBox(e.Message, MsgBoxStyle.OkOnly, "SSHStatus Messages")
        theLog &= e.Message & vbCrLf

    End Sub

    Public Shared Function sftp_put( _
        frmASFBASE0 As ASFBASE0, _
        SSH_APP_CODE As String, _
        production As Boolean, _
        FILENAME_LOCAL As String, _
        FILENAME_REMOTE As String) As Boolean

        Dim rowTATSSHK1 As DataRow = Nothing ' frmASFBASE0.LookUp("TATSSHK1", SSH_APP_CODE)
        rowTATSSHK1 = ASCDATA1.GetDataRow("Select * from TATSSHK1 where SSH_APP_CODE = '" & SSH_APP_CODE & "'")

        Dim SSH_APP_USERNAME As String = rowTATSSHK1.Item("SSH_APP_USERNAME") & ""
        Dim SSH_APP_PASSWORD As String = rowTATSSHK1.Item("SSH_APP_PASSWORD") & ""
        Dim SSH_APP_FOLDER_PUT As String = rowTATSSHK1.Item("SSH_APP_FOLDER_PUT") & ""
        If frmASFBASE0 IsNot Nothing AndAlso frmASFBASE0.Name = "APFCHCKA" Then ' ACH & WIRES HAVE A DIFFERENT FOLDER THAN POSITIVE PAY
            SSH_APP_FOLDER_PUT = "H2H01"
        End If
        Dim SSH_APP_PORT As Int32 = Val(rowTATSSHK1.Item("SSH_APP_PORT") & "")
        If SSH_APP_PORT = 0 Then SSH_APP_PORT = 22

        Dim SSH_APP_PARTNER_URI As String = ""
        Dim SSH_APP_PARTNER_PUBKEY As String = ""
        If production Then
            SSH_APP_PARTNER_URI = rowTATSSHK1.Item("SSH_APP_PARTNER_URI_PROD") & ""
            SSH_APP_PARTNER_PUBKEY = rowTATSSHK1.Item("SSH_APP_PARTNER_PUBKEY_PROD") & ""
        Else
            SSH_APP_PARTNER_URI = rowTATSSHK1.Item("SSH_APP_PARTNER_URI_TEST") & ""
            SSH_APP_PARTNER_PUBKEY = rowTATSSHK1.Item("SSH_APP_PARTNER_PUBKEY_TEST") & ""
        End If

        Dim SSH_APP_SSH_PVTKEY_FILE As String = rowTATSSHK1.Item("SSH_APP_SSH_PVTKEY_FILE") & ""

        Dim SSH_APP_SSH_PVTKEY As String = rowTATSSHK1.Item("SSH_APP_SSH_PVTKEY") & ""
        Dim SSH_APP_SSH_PUBKEY As String = rowTATSSHK1.Item("SSH_APP_SSH_PUBKEY") & ""

        Dim SSH_APP_PGP_PVTKEY_PWD As String = rowTATSSHK1.Item("SSH_APP_PGP_PVTKEY_PWD") & ""
        Dim SSH_APP_SSH_PVTKEY_PWD As String = rowTATSSHK1.Item("SSH_APP_SSH_PVTKEY_PWD") & ""

        Dim success = False
        Dim sftp As New nsoftware.IPWorksSSH.SFTPClient
        theLog = ""

        AddHandler sftp.OnSSHServerAuthentication, AddressOf SSHServerAuthentication
        AddHandler sftp.OnSSHStatus, AddressOf SSHStatus

        sftp.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwaresftpkey")

        'Dim crypt As New ASCSCRTY

        ' SSH_APP_OUR_PVTKEY = crypt.Decrypt_AES(SSH_APP_OUR_PVTKEY)
        Dim SSH_APP_SSH_PVTKEY_B() As Byte = StrToByteArray(SSH_APP_SSH_PVTKEY)

        ' SSH_APP_PARTNER_PUBKEY = crypt.Decrypt_AES(SSH_APP_PARTNER_PUBKEY)
        Dim SSH_APP_PARTNER_PUBKEY_B() As Byte = StrToByteArray(SSH_APP_PARTNER_PUBKEY)

        If SSH_APP_PARTNER_PUBKEY <> "" Then
            sftp.SSHAcceptServerHostKey = New Certificate(CertStoreTypes.cstSSHPublicKeyBlob, SSH_APP_PARTNER_PUBKEY_B, "", "")
        End If


        If SSH_APP_CODE = "EXP" Then


            sftp.SSHAuthMode = nsoftware.IPWorksSSH.SCPSSHAuthModes.amPublicKey
            sftp.SSHUser = SSH_APP_USERNAME
            sftp.SSHCert = New Certificate(CertStoreTypes.cstPEMKeyBlob, SSH_APP_SSH_PVTKEY_B, "", "*")
            'cstSSHPublicKey
            sftp.SSHAcceptServerHostKey = New Certificate(CertStoreTypes.cstSSHPublicKeyBlob, SSH_APP_PARTNER_PUBKEY_B, "", "*")

            'Verify the signature is authentic using the sender's public key(decrypt Signature block)
            If myReceiver.VerifyHash(mySender.PublicParameters, encrypted, signature) Then
                MsgBox("Signature Valid", MsgBoxStyle.Information)
            Else
                MsgBox("Invalid Signature", MsgBoxStyle.Exclamation)
            End If

            Dim signatureText As String = "A&E"
            'Convert the data string to a byte array.
            toEncrypt = enc.GetBytes(signatureText)

            'Encrypt data using receiver's public key.
            encrypted = mySender.EncryptData(myReceiver.PublicParameters, toEncrypt)

            'Hash the encrypted data and generate a signature block on the hash
            ' using the sender's private key. (Signature Block)
            signature = mySender.HashAndSign(encrypted)

            Try
                If sftp.Connected = True Then
                    sftp.SSHLogoff()
                End If

                sftp.SSHLogon(SSH_APP_PARTNER_URI, SSH_APP_PORT)
                success = True

            Catch ex As Exception
                MessageBox.Show(ex.Message, "Secure ftp Setup", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If sftp.Connected = True Then
                    sftp.SSHLogoff()
                End If
            End Try

        Else

            sftp.SSHUser = SSH_APP_USERNAME

            If SSH_APP_PASSWORD <> "" Then
                sftp.SSHAuthMode = SCPSSHAuthModes.amPassword
                sftp.SSHPassword = SSH_APP_PASSWORD

                ' Added 06/1-0/2025
                If SSH_APP_CODE = "ADS" Then
                    sftp.SSHEncryptionAlgorithms = "aes256-gmc,aes256-ctr,aes128-gcm"
                End If
            Else
                sftp.SSHAuthMode = SCPSSHAuthModes.amPublicKey
                If SSH_APP_SSH_PVTKEY <> "" Then
                    sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKBlob, SSH_APP_SSH_PVTKEY_B, SSH_APP_SSH_PVTKEY_PWD, "*")
                Else
                    Dim ssh_folder As String = ""
                    If ASCMAIN1.Running_in_VS Then
                        Stop
                        ssh_folder = "C:\Users\wjz\Desktop\Interparfums\"
                    Else
                        'ssh_folder = ASCMAIN1.Folders("SharedRoot") & $"Archive\INT\"
                        ssh_folder = ASCMAIN1.Folders("Archive")
                    End If
                    Dim ssh_file As String = $"{ssh_folder}{SSH_APP_CODE}\{SSH_APP_SSH_PVTKEY_FILE}"
                    sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, ssh_file, SSH_APP_SSH_PVTKEY_PWD, "*")
                End If
            End If

            Try

                If sftp.Connected = True Then
                    sftp.SSHLogoff()
                End If

                If ASCMAIN1.CLIENT = "INT" Then
                    If SSH_APP_CODE = "JPMC" Then
                        sftp.SSHEncryptionAlgorithms = "aes128-ctr,aes192-ctr,aes256-ctr"
                        ' DEFAULT WAS aes256-cbc,3des-cbc
                        sftp.Config("LogSSHPackets=True")
                        'If ASCMAIN1.USER_ID = "wjz" Then MsgBox(sftp.Config("LogSSHPackets"))
                    Else
                        ' COWORX DOES NOT SUPPORT NEW ENCRYPTION
                    End If
                End If

                sftp.SSHHost = SSH_APP_PARTNER_URI

                sftp.SSHLogon(SSH_APP_PARTNER_URI, SSH_APP_PORT)

                sftp.LocalFile = FILENAME_LOCAL
                sftp.ChangeRemotePath(SSH_APP_FOLDER_PUT)

                sftp.RemoteFile = FILENAME_REMOTE
                sftp.Upload()

                success = True

            Catch ex As Exception
                theLog &= ex.Message
                Dim filename As String = Format(Now, "yyyyMMddhhhhss")
                System.IO.File.WriteAllText(ASCMAIN1.Folders("Work") & filename & ".log", theLog)
                MessageBox.Show(ex.Message, "Secure ftp Setup", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If sftp.Connected = True Then
                    sftp.SSHLogoff()
                End If
            End Try

        End If

        If sftp.Connected = True Then
            sftp.SSHLogoff()
        End If

        Return success

    End Function

    Public Shared Function sftp_get( _
        frmASFBASE0 As ASFBASE0, _
        SSH_APP_CODE As String, _
        production As Boolean, _
        FILENAME_LOCAL As String, _
        FILENAME_REMOTE As String) As List(Of String)

        Dim FILENAMEs As New List(Of String)

        Dim rowTATSSHK1 As DataRow = Nothing ' frmASFBASE0.LookUp("TATSSHK1", SSH_APP_CODE)
        rowTATSSHK1 = ASCDATA1.GetDataRow("Select * from TATSSHK1 where SSH_APP_CODE = '" & SSH_APP_CODE & "'")

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

            sftp.ListDirectory()
            For Each s As nsoftware.IPWorksSSH.DirEntry In sftp.DirList
                sftp.RemoteFile = s.FileName
                If Not s.IsDir Then
                    ASCMAIN1.Progress("-", s.FileName)
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

    Sub TestDigitalSignature()
        'The hash value to sign.
        Dim HashValue As Byte() = {59, 4, 248, 102, 77, 97, 142, 201, 210, 12, 224, 93, 25, 41, 100, 197, 213, 134, 130, 135}

        'The value to hold the signed value.
        Dim SignedHashValue() As Byte

        Dim cp As New CspParameters

        'Generate a public/private key pair.
        Dim RSA As New RSACryptoServiceProvider()

        'Create an RSAPKCS1SignatureFormatter object and pass it
        'the RSACryptoServiceProvider to transfer the private key.
        Dim RSAFormatter As New RSAPKCS1SignatureFormatter(RSA)

        'Set the hash algorithm to SHA1.
        RSAFormatter.SetHashAlgorithm("SHA1")

        'Create a signature for HashValue and assign it to
        'SignedHashValue.
        SignedHashValue = RSAFormatter.CreateSignature(HashValue)
    End Sub

    Public Shared Function StrToByteArray(ByVal str As String) As Byte()
        Dim encoding As New System.Text.UTF8Encoding()
        Return encoding.GetBytes(str)
    End Function

    Sub Sign_FileSignFile(ByVal FilePath As String, ByVal KeyPath As String)
        ' Signing Step 1: Create the digital signature algorithm object
        Dim signer As DSACryptoServiceProvider = New DSACryptoServiceProvider

        ' Signing Step 2: Store the data to be signed in a byte array.
        Dim file As System.IO.FileStream = New System.IO.FileStream(FilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read)
        Dim reader As System.IO.BinaryReader = New System.IO.BinaryReader(file)
        Dim data As Byte() = reader.ReadBytes(CType(file.Length, Integer))

        ' Signing Step 3: Call the SignData method and create the signature
        Dim signature As Byte() = signer.SignData(data)

        ' Signing Step 4: Export the public key
        ' Save the public key and the Signature in a file. 

        Using sr As New System.IO.StreamWriter(KeyPath)
            sr.Write(signer.ToXmlString(False))
            sr.WriteLine()
            sr.Write("Signature:" & System.Convert.ToBase64String(signature))
        End Using
        reader.Close()
        file.Close()

    End Sub

End Class


Module Globalx
    'Hold message in bytes
    Public toEncrypt() As Byte
    'holds encrypted data
    Public encrypted() As Byte
    'holds signatures
    Public signature() As Byte
    'instance of class Bob
    Public mySender As New Bob
    'instance of class Alice
    Public myReceiver As New Alice
    'new instance of Unicode8 instance
    Public enc As New System.Text.UTF8Encoding
End Module

'Imports System.Security.Cryptography
'Imports System.Text
Public Class Bob
    '=======================================================
    'Bob Sender is who want to exchange and prepares a encrypted message
    '=======================================================
    Private rsaPubParams As RSAParameters          'stores public key
    Private rsaPrivateParams As RSAParameters     'stores private key

    Public Sub New()
        'create new instance of RSACryptoServiceProvider
        Dim rsaCSP As New RSACryptoServiceProvider
        'Generate public and private key data and allowing their exporting.
        'True to include private parameters; otherwise, false
        rsaPrivateParams = rsaCSP.ExportParameters(True)
        rsaPubParams = rsaCSP.ExportParameters(False)
    End Sub 'New

    Public ReadOnly Property PublicParameters() As RSAParameters
        Get
            'just return public key
            Return rsaPubParams
        End Get
    End Property

    'Manually performs hash and then signs hashed value.
    Public Function HashAndSign(ByVal encrypted() As Byte) As Byte()
        'create new instance of RSACryptoServiceProvider
        Dim rsaCSP As New RSACryptoServiceProvider
        'create new instance of SHA1 hash algorithm to compute hash
        Dim hash As New SHA1Managed
        'a byte array to store hash value        
        Dim hashedData() As Byte
        'import private key params into instance of RSACryptoServiceProvider
        rsaCSP.ImportParameters(rsaPrivateParams)
        'compute hash with algorithm specified as here we have SHA!
        hashedData = hash.ComputeHash(encrypted)
        ' Sign Data using private key and  OID is simple name
        ' of the algorithm for which to get the object identifier (OID)
        Return rsaCSP.SignHash(hashedData, CryptoConfig.MapNameToOID("SHA1"))
    End Function 'HashAndSign

    'Encrypts using only the public key data.
    Public Function EncryptData(ByVal rsaParams As RSAParameters, _
                    ByVal toEncrypt() As Byte) As Byte()
        'create new instance of RSACryptoServiceProvider
        Dim rsaCSP As New RSACryptoServiceProvider
        ''import private key params into instance of RSACryptoServiceProvider
        rsaCSP.ImportParameters(rsaParams)
        'true to use OAEP padding PKCS#1 v2  (only available on Windows XP or later)
        ' otherwise, false to use Direct Encryption using PKCS#1 v1.5 padding
        Return rsaCSP.Encrypt(toEncrypt, False)
    End Function 'EncryptData

End Class   'Bob

'Imports System.Security.Cryptography
'Imports System.Text
Public Class Alice
    '=========================================
    ' Alice is Receiver who will decrypt data from sender
    '========================================
    Private rsaPubParams As RSAParameters        ' stores public key
    Private rsaPrivateParams As RSAParameters    'stores private key

    Public Sub New()
        'create new instance of RSACryptoServiceProvider
        Dim rsaCSP As New RSACryptoServiceProvider
        'Generate public and private key data and allowing their exporting.
        'True to include private parameters; otherwise, false
        rsaPrivateParams = rsaCSP.ExportParameters(True)
        rsaPubParams = rsaCSP.ExportParameters(False)
    End Sub 'New

    Public ReadOnly Property PublicParameters() As RSAParameters
        Get
            'Just return public key
            Return rsaPubParams
        End Get
    End Property

    'Manually performs hash and then verifies hashed value.
    Public Function VerifyHash(ByVal rsaParams As RSAParameters, _
                    ByVal signedData() As Byte, _
                    ByVal signature() As Byte) As Boolean
        'create new instance of RSACryptoServiceProvider
        Dim rsaCSP As New RSACryptoServiceProvider
        'create new instance of SHA1 hash algorithm to compute hash
        Dim hash As New SHA1Managed
        'a byte array to store hash value
        Dim hashedData() As Byte
        'import  public key params into instance of RSACryptoServiceProvider
        rsaCSP.ImportParameters(rsaParams)
        'compute hash with algorithm specified as here we have SHA1
        hashedData = hash.ComputeHash(signedData)
        ' Sign Data using public key and OID is simple name
        ' of the algorithm for which to get the object identifier (OID)
        Return rsaCSP.VerifyHash(hashedData, _
               CryptoConfig.MapNameToOID("SHA1"), signature)
    End Function 'VerifyHash

    'Decrypt using the private key data.
    Public Function DecryptData(ByVal encrypted() As Byte) As String
        'a byte array to store decrypted bytes
        Dim fromEncrypt() As Byte
        ' holds orignal message
        Dim roundTrip As String
        'create new instance of RSACryptoServiceProvider
        Dim rsaCSP As New RSACryptoServiceProvider
        'import  private key params into instance of RSACryptoServiceProvider
        rsaCSP.ImportParameters(rsaPrivateParams)
        'store decrypted data into byte array
        fromEncrypt = rsaCSP.Decrypt(encrypted, False)
        'convert bytes to string
        roundTrip = enc.GetString(fromEncrypt)
        Return roundTrip
    End Function 'DecryptData

End Class ' Alice