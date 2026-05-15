Imports ABSolution
Imports nsoftware.IPWorks.Ftp

Public Class RSCFTPC1

    Friend WithEvents ftp1 As nsoftware.IPWorks.Ftp

#Region "Class Variables"

    Private SO_PARM_MDB_PROVIDER As String = String.Empty
    Private SO_PARM_FTP_IP_ADDRESS As String = String.Empty
    Private SO_PARM_FTP_USER As String = String.Empty
    Private SO_PARM_FTP_PASSWORD As String = String.Empty
    Private SO_PARM_DIR_INBOUND As String = String.Empty
    Private SO_PARM_DIR_OUTBOUND As String = String.Empty
    Private SO_PARM_DIR_ARCHIVE As String = String.Empty
    Private SO_PARM_DOWNLOAD_FILE_NAME As String = String.Empty
    Private SO_PARM_DOWNLOAD_FILE_EXT As String = String.Empty
    Private SO_PARM_LOCAL_DIR_INBOUND As String = String.Empty
    Private SO_PARM_LOCAL_DIR_OUTBOUND As String = String.Empty
    Private SO_PARM_LOCAL_DIR_ARCHIVE As String = String.Empty
    Private SO_PARM_FTP_HOST_NAME As String = String.Empty

    Private Errors As List(Of String)
    Private remoteDirectoryFileList As List(Of String) = New List(Of String)
    Private displayControl As Windows.Forms.Control = Nothing
    Private fileExtensions As List(Of String) = New List(Of String)

#End Region

#Region "Class Constructors"

    ''' <summary>
    ''' Class Constructor
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        Me.Initialize(String.Empty)
    End Sub

    ''' <summary>
    ''' Class Constructor
    ''' </summary>
    ''' <param name="ftpParameterKey">Parameter Key in table SOTPARMF</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal ftpParameterKey As String)
        ftpParameterKey = ftpParameterKey.Trim
        Me.Initialize(ftpParameterKey)
    End Sub

    ''' <summary>
    ''' Initializes the class Variables
    ''' </summary>
    ''' <param name="ftpParameterKey"></param>
    ''' <remarks></remarks>
    Private Sub Initialize(ByVal ftpParameterKey As String)

        SO_PARM_FTP_IP_ADDRESS = String.Empty
        SO_PARM_FTP_USER = String.Empty
        SO_PARM_FTP_PASSWORD = String.Empty
        SO_PARM_DIR_INBOUND = String.Empty
        SO_PARM_DIR_OUTBOUND = String.Empty
        SO_PARM_DIR_ARCHIVE = String.Empty
        SO_PARM_DOWNLOAD_FILE_NAME = String.Empty
        SO_PARM_DOWNLOAD_FILE_EXT = String.Empty
        SO_PARM_LOCAL_DIR_INBOUND = String.Empty
        SO_PARM_LOCAL_DIR_OUTBOUND = String.Empty
        SO_PARM_LOCAL_DIR_ARCHIVE = String.Empty

        If ftpParameterKey.Length > 0 Then

            Dim sql As String = "Select * From SOTPARMF Where SO_PARM_MDB_PROVIDER = '" & ftpParameterKey & "'"
            Dim rowSOTPARMF As DataRow = ABSolution.ASCDATA1.GetDataRow(sql)

            If rowSOTPARMF IsNot Nothing Then
                SO_PARM_MDB_PROVIDER = (rowSOTPARMF.Item("SO_PARM_MDB_PROVIDER") & String.Empty).ToString.Trim
                SO_PARM_FTP_IP_ADDRESS = (rowSOTPARMF.Item("SO_PARM_FTP_IP_ADDRESS") & String.Empty).ToString.Trim
                SO_PARM_FTP_USER = (rowSOTPARMF.Item("SO_PARM_FTP_USER") & String.Empty).ToString.Trim
                SO_PARM_FTP_PASSWORD = (rowSOTPARMF.Item("SO_PARM_FTP_PASSWORD") & String.Empty).ToString.Trim
                SO_PARM_DIR_INBOUND = (rowSOTPARMF.Item("SO_PARM_DIR_INBOUND") & String.Empty).ToString.Trim
                SO_PARM_DIR_OUTBOUND = (rowSOTPARMF.Item("SO_PARM_DIR_OUTBOUND") & String.Empty).ToString.Trim
                SO_PARM_DIR_ARCHIVE = (rowSOTPARMF.Item("SO_PARM_DIR_ARCHIVE") & String.Empty).ToString.Trim
                SO_PARM_DOWNLOAD_FILE_NAME = (rowSOTPARMF.Item("SO_PARM_DOWNLOAD_FILE_NAME") & String.Empty).ToString.Trim
                SO_PARM_DOWNLOAD_FILE_EXT = (rowSOTPARMF.Item("SO_PARM_DOWNLOAD_FILE_EXT") & String.Empty).ToString.Trim
                SO_PARM_LOCAL_DIR_INBOUND = (rowSOTPARMF.Item("SO_PARM_LOCAL_DIR_INBOUND") & String.Empty).ToString.Trim
                SO_PARM_LOCAL_DIR_OUTBOUND = (rowSOTPARMF.Item("SO_PARM_LOCAL_DIR_OUTBOUND") & String.Empty).ToString.Trim
                SO_PARM_LOCAL_DIR_ARCHIVE = (rowSOTPARMF.Item("SO_PARM_LOCAL_DIR_ARCHIVE") & String.Empty).ToString.Trim
                SO_PARM_FTP_HOST_NAME = (rowSOTPARMF.Item("SO_PARM_FTP_HOST_NAME") & String.Empty).ToString.Trim
            End If
        End If

        ' Host Directories
        If SO_PARM_DIR_INBOUND.Length > 0 Then
            If Not SO_PARM_DIR_INBOUND.EndsWith("\") Then
                SO_PARM_DIR_INBOUND &= "\"
            End If
        End If

        If SO_PARM_DIR_OUTBOUND.Length > 0 Then
            If Not SO_PARM_DIR_OUTBOUND.EndsWith("\") Then
                SO_PARM_DIR_OUTBOUND &= "\"
            End If
        End If

        If SO_PARM_DIR_ARCHIVE.Length > 0 Then
            If Not SO_PARM_DIR_ARCHIVE.EndsWith("\") Then
                SO_PARM_DIR_ARCHIVE &= "\"
            End If
        End If

        ' Local Directories
        If SO_PARM_LOCAL_DIR_INBOUND.Length > 0 Then
            If Not SO_PARM_LOCAL_DIR_INBOUND.EndsWith("\") Then
                SO_PARM_LOCAL_DIR_INBOUND &= "\"
            End If
        End If

        If SO_PARM_LOCAL_DIR_INBOUND.Length > 0 Then
            If Not SO_PARM_LOCAL_DIR_INBOUND.EndsWith("\") Then
                SO_PARM_LOCAL_DIR_INBOUND &= "\"
            End If
        End If

        If SO_PARM_LOCAL_DIR_ARCHIVE.Length > 0 Then
            If Not SO_PARM_LOCAL_DIR_ARCHIVE.EndsWith("\") Then
                SO_PARM_LOCAL_DIR_ARCHIVE &= "\"
            End If
        End If


        Errors = New List(Of String)
        remoteDirectoryFileList = New List(Of String)
        fileExtensions = New List(Of String)
        displayControl = Nothing

    End Sub

#End Region

#Region "Class Properties"

    ''' <summary>
    ''' Gets / Sets Control to send messages to. Allows the user to receive status.
    ''' Control has to have a text or caption property
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DisplayMessageControl() As Windows.Forms.Control
        Get
            Return displayControl
        End Get
        Set(ByVal value As Windows.Forms.Control)
            displayControl = value
        End Set
    End Property

    ''' <summary>
    ''' Gets / Sets Ftp Paramters
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ftpParameterKey() As String
        Get
            Return SO_PARM_MDB_PROVIDER
        End Get
        Set(ByVal value As String)
            Me.Initialize(value)
        End Set
    End Property

    ''' <summary>
    ''' Gets / Set any filename wild card to down load
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DownLoadFilesFileName() As String
        Get
            Return Me.SO_PARM_DOWNLOAD_FILE_NAME
        End Get
        Set(ByVal value As String)
            SO_PARM_DOWNLOAD_FILE_NAME = value.Trim
            If SO_PARM_DOWNLOAD_FILE_NAME.Length > 0 Then
                If Not SO_PARM_DOWNLOAD_FILE_NAME.EndsWith("\") Then
                    SO_PARM_DOWNLOAD_FILE_NAME &= "\"
                End If
            End If
        End Set
    End Property

    ''' <summary>
    ''' Gets / Set file extensions to use for downloading purposes
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DownLoadFilesFileExtension() As List(Of String)
        Get
            Return fileExtensions
        End Get

        Set(ByVal value As List(Of String))
            fileExtensions.Clear()
            For Each v As String In value
                fileExtensions.Add(v.Trim.ToUpper)
            Next
        End Set
    End Property

    ''' <summary>
    ''' Gets a list of the files downloaded
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property DownLoadedFiles() As List(Of String)
        Get
            Return remoteDirectoryFileList
        End Get
    End Property

    ''' <summary>
    ''' Gets a list of errors that occured during the FTP
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ErrorList() As List(Of String)
        Get
            Return Errors
        End Get
    End Property


#End Region

#Region "Class Public Procedures"

    ''' <summary>
    ''' Download Files from FTP Site
    ''' </summary>
    ''' <param name="ArchiveFileOnHostMachine"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DownloadFiles(ByVal ArchiveFileOnHostMachine As Boolean, ByVal DeleteFileOnHostMachine As Boolean) As Long

        DownloadFiles = 0
        Errors.Clear()
        remoteDirectoryFileList.Clear()

        Dim iFilesDownloaded As Long = 0
        Dim validDownLoadFile As Boolean = False
        Dim fileName As String = String.Empty

        ftp1 = New nsoftware.IPWorks.Ftp
        ftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")

        Try
            If Not ftp1.Connected Then
                ftp1.RemoteHost = Me.SO_PARM_FTP_IP_ADDRESS
                ftp1.RemotePath = Me.SO_PARM_DIR_OUTBOUND
                ftp1.User = Me.SO_PARM_FTP_USER
                ftp1.Password = Me.SO_PARM_FTP_PASSWORD
                ftp1.RemoteFile = String.Empty

                ftp1.Logon()
            End If

            If Not ftp1.Connected Then
                Errors.Add("Could not connect to remote host to download shipment confirmation files.")
                Return 0
            End If

            Dim downLoadedFiles As List(Of String) = New List(Of String)

            Try
                ftp1.ListDirectory()
            Catch ex As Exception
                ' just in case, no need to bomb out
            End Try
            ftp1.Overwrite = True

            iFilesDownloaded = 0

            For Each remoteDirItem As String In remoteDirectoryFileList
                fileName = remoteDirItem.Trim.ToUpper
                validDownLoadFile = True

                If validDownLoadFile Then

                    If displayControl IsNot Nothing Then
                        Try
                            displayControl.Text = "Downloading " & fileName
                        Catch ex As Exception
                            ' Nothing
                        End Try
                    End If

                    ftp1.RemoteFile = fileName
                    ftp1.LocalFile = Me.SO_PARM_LOCAL_DIR_INBOUND & My.Computer.FileSystem.GetName(fileName)
                    If ASCMAIN1.Running_in_VS Then
                        Stop
                        ftp1.LocalFile = "C:\Documents and Settings\wjz\Desktop\LGI\INBOUND\" & My.Computer.FileSystem.GetName(fileName)
                    End If
                    ftp1.Download()
                    ftp1.DoEvents()

                    ' Move file to the unix archive
                    If ArchiveFileOnHostMachine And Me.SO_PARM_DIR_ARCHIVE.Length > 0 Then
                        Try
                            ftp1.RenameFile("..\" & SO_PARM_DIR_ARCHIVE & fileName)
                            ftp1.DoEvents()
                        Catch ex As Exception
                            Errors.Add(ex.Message)
                        End Try
                    End If

                    If DeleteFileOnHostMachine Then
                        Try
                            ftp1.DeleteFile(fileName)
                            ftp1.DoEvents()
                        Catch ex As Exception
                            Errors.Add(ex.Message)
                        End Try
                    End If

                    ' keep a copy of the downloaded file's name
                    downLoadedFiles.Add(fileName)
                    iFilesDownloaded += 1
                End If
            Next

            ftp1.Logoff()

        Catch ex As Exception
            Errors.Add(ex.Message)
            Try
                If ftp1.Connected Then ftp1.Logoff()
            Catch ex1 As Exception
                ' nothing
            End Try

            Return iFilesDownloaded
        End Try

        Return iFilesDownloaded

    End Function

    ''' <summary>
    ''' Get file list from Remote Directory 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Ftp1_OnDirList(ByVal sender As System.Object, ByVal e As nsoftware.IPWorks.FtpDirListEventArgs) Handles ftp1.OnDirList
        If Not e.IsDir Then
            If fileExtensions.Count = 0 Then
                remoteDirectoryFileList.Add(e.FileName)
            Else
                Dim fileParts() As String = e.FileName.Split(".")
                If fileExtensions.Contains(fileParts(fileParts.GetUpperBound(0)).ToUpper) Then
                    remoteDirectoryFileList.Add(e.FileName)
                End If
            End If
        End If
    End Sub
#End Region
End Class
