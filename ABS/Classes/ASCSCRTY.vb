Imports System.Security.Cryptography
Imports System.Text
Imports System.IO

''' <summary>
''' Security Encryption / Decryption Class
''' </summary>
''' <remarks></remarks>
Public Class ASCSCRTY

    Private Const keyWord As String = "P@$$w0rd"
    Private aesException As Exception = New Exception

    Public Sub New()

    End Sub

    ''' <summary>
    ''' Decrypts a string
    ''' </summary>
    ''' <param name="Str2decrypt"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Decrypt(ByVal Str2decrypt As String) As String
        Return Decrypt_AES(Str2decrypt)
    End Function

    ''' <summary>
    ''' Encrypts a string
    ''' </summary>
    ''' <param name="str2encrypt"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Encrypt(ByVal str2encrypt As String) As String
        Return Encrypt_AES(str2encrypt)
    End Function

    ''' <summary>
    ''' Returns the last execption which occurred when Encrypt and Decrypt are called
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetLastException() As Exception
        Get
            Return aesException
        End Get
    End Property

    ''' <summary>
    ''' Encrypts a string
    ''' </summary>
    ''' <param name="str2encrypt">String to Encrypt</param>
    ''' <returns>Encrypted value</returns>
    ''' <remarks></remarks>
    Public Function Encrypt_AES(ByVal str2encrypt As String) As String

        Try

            aesException = New Exception

            Dim hexEncryption As String = Nothing
            Dim encrypted() As Byte
            Dim IV() As Byte
            Dim key(keyWord.Length) As Byte
            Dim myRijndael As New RijndaelManaged()
            Dim iter = 0
            Dim toEncrypt() As Byte
            Dim textConverter As New ASCIIEncoding()

            If str2encrypt.Length = 0 Then
                Return String.Empty
            End If

            For Each var As Char In keyWord
                key(iter) = CByte(Asc(var))
                iter += 1
            Next

            IV = myRijndael.IV

            Dim encryptor As ICryptoTransform = myRijndael.CreateEncryptor(key, IV)

            Dim msEncrypt As New MemoryStream()
            Dim csEncrypt As New CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write)

            toEncrypt = textConverter.GetBytes(str2encrypt)

            csEncrypt.Write(toEncrypt, 0, toEncrypt.Length)
            csEncrypt.FlushFinalBlock()
            encrypted = msEncrypt.ToArray()

            hexEncryption = ByteArrayToString(IV) & ByteArrayToString(encrypted)

            Return hexEncryption.ToUpper
            iter = 0

        Catch ex As Exception
            aesException = ex
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Decrypts a string
    ''' </summary>
    ''' <param name="Str2decrypt">String to decrypt</param>
    ''' <returns>Decrypted String</returns>
    ''' <remarks></remarks>
    Public Function Decrypt_AES(ByVal Str2decrypt As String) As String

        Try
            aesException = New Exception

            Dim myRijndael As New RijndaelManaged()
            Dim fromEncrypt() As Byte
            Dim encrypted(15) As Byte
            Dim convertedStr As String
            Dim textConverter As New ASCIIEncoding()
            Dim key(keyWord.Length) As Byte
            Dim IV(15) As Byte
            Dim iter = 0
            Dim counter As Decimal = -1
            Dim hexbit As String = Nothing
            Dim pswByte As Integer

            If Str2decrypt.Length = 0 Then
                Return String.Empty
            End If

            'Make key
            For Each var As Char In keyWord
                key(iter) = CByte(Asc(var))
                iter += 1
            Next

            iter = 0
            'Seperate IV from Encryption
            For Each letter As Char In Str2decrypt
                If counter < 15 Then
                    If iter = 1 Then
                        hexbit = hexbit & letter
                        pswByte = Convert.ToInt16(hexbit, 16)
                        counter += 0.5
                        IV(counter) = pswByte
                        hexbit = Nothing
                        iter = 0
                    Else
                        hexbit = hexbit & letter
                        iter += 1
                        counter += 0.5
                    End If
                Else
                    If iter = 1 Then
                        hexbit = hexbit & letter
                        pswByte = Convert.ToInt16(hexbit, 16)
                        counter += 0.5
                        ReDim Preserve encrypted(counter - 16)
                        encrypted(counter - 16) = pswByte
                        hexbit = Nothing
                        iter = 0
                    Else
                        hexbit = hexbit & letter
                        iter += 1
                        counter += 0.5
                    End If
                End If
            Next

            Dim decryptor As ICryptoTransform = myRijndael.CreateDecryptor(key, IV)

            Dim msdecrypt As New MemoryStream(encrypted)
            Dim csDecrypt As New CryptoStream(msdecrypt, decryptor, CryptoStreamMode.Read)

            fromEncrypt = New Byte(encrypted.Length) {}
            csDecrypt.Read(fromEncrypt, 0, fromEncrypt.Length)
            convertedStr = textConverter.GetString(fromEncrypt)

            ' Need this here sice result has lots of chr(0) at end of string.
            convertedStr = convertedStr.Replace(Chr(0), String.Empty)

            Return convertedStr
        Catch ex As Exception
            aesException = ex
            Return String.Empty
        End Try

    End Function

    ''' <summary>
    ''' Converts a Byte Array to a String
    ''' </summary>
    ''' <param name="ba"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ByteArrayToString(ByVal ba As Byte()) As String
        Try
            Dim hex As New StringBuilder(ba.Length * 2)
            For Each b As Byte In ba
                hex.AppendFormat("{0:x2}", b)
            Next
            Return hex.ToString()

        Catch ex As Exception
            aesException = ex
            Return String.Empty
        End Try

    End Function


End Class
