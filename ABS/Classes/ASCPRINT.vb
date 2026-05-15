Imports System.IO
Imports System.Drawing.Printing
Imports System.Drawing.Printing.PrinterSettings
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Text
Imports System.Diagnostics

Public Class ASCPRINT

    ''' <summary>
    ''' Various states of a printer
    ''' </summary>
    ''' <remarks></remarks>
    Enum PrintingStates
        Open
        StartDoc
        StartPage
        EndPage
        EndDoc
        Closed
    End Enum

    Enum PrintDataTypes
        RAW
        EMF
        XPS_PASS
    End Enum

    Private clsPrinterName As String = String.Empty
    Private clsLastPrintError As String
    Private clsPrinterState As PrintingStates
    Private clsPrinterHandle As IntPtr
    Private clsPrintDataType As PrintDataTypes
    Private clsSerialPrintPort As System.IO.Ports.SerialPort

#Region "Constructors"

    ''' <summary>
    ''' Class Constructor
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        clsPrinterName = String.Empty
        Call Initialize()
    End Sub

    ''' <summary>
    ''' Class Constructor w/setting the printer used to print
    ''' </summary>
    ''' <param name="PrinterName"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal PrinterName As String)
        clsPrinterName = PrinterName.Trim
        Call Initialize()
    End Sub

    ''' <summary>
    ''' Class Constructor w/setting the Serial Port used to print
    ''' </summary>
    ''' <param name="SerialPrintPort"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal SerialPrintPort As System.IO.Ports.SerialPort)
        Call Initialize()
        clsPrinterName = PrinterName.Trim
        clsSerialPrintPort = SerialPrintPort
    End Sub

    ''' <summary>
    ''' Class Constructor w/setting the Serial Port and printer used to print
    ''' </summary>
    ''' <param name="PrinterName"></param>
    ''' <param name="PrintSerialPort"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal PrinterName As String, ByVal PrintSerialPort As System.IO.Ports.SerialPort)
        clsPrinterName = PrinterName.Trim
        Call Initialize()
        clsSerialPrintPort = PrintSerialPort
    End Sub

    ''' <summary>
    ''' Initialize Class Variables
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Initialize()
        ' Default Printer to Computer Default Printer
        clsLastPrintError = String.Empty
        '_printerName = String.Empty
        Call SetPrinter(clsPrinterName)

        clsPrinterState = PrintingStates.Closed
        clsPrinterHandle = 0
        clsPrintDataType = PrintDataTypes.RAW

        clsSerialPrintPort = Nothing

    End Sub

#End Region

#Region "Class Properties"

    ''' <summary>
    ''' Returns a list of the Installed Printers
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetInstalledPrinters() As List(Of String)
        Get
            Dim instPrinters As New List(Of String)

            Dim printerCount As Integer = InstalledPrinters.Count
            For i As Integer = 0 To printerCount - 1
                instPrinters.Add(InstalledPrinters(i))
            Next

            Return instPrinters
        End Get
    End Property

    ''' <summary>
    ''' Returns whether the supplied printer name is knwon by the computer
    ''' </summary>
    ''' <param name="szPrinterName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property IsValidPrinterName(ByVal szPrinterName As String) As Boolean
        Get
            Dim found As Boolean = False
            Dim instPrinters As New List(Of String)

            Dim printerCount As Integer = InstalledPrinters.Count
            For i As Integer = 0 To printerCount - 1
                If szPrinterName.Trim.ToUpper = InstalledPrinters(i).Trim.ToUpper Then
                    found = True
                    Exit For
                End If
            Next

            Return found
        End Get
    End Property

    ''' <summary>
    ''' Returns The Last Print Error
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property LastPrintError() As String
        Get
            Return (clsLastPrintError)
        End Get
    End Property

    ''' <summary>
    ''' Get / Set Print Data Type used for printing
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PrintDataType() As PrintDataTypes
        Get
            Return clsPrintDataType
        End Get
        Set(ByVal value As PrintDataTypes)
            clsPrintDataType = value
        End Set
    End Property

    ''' <summary>
    ''' Get / Set Printer Name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PrinterName() As String
        Get
            Return clsPrinterName
        End Get
        Set(ByVal value As String)
            SetPrinter(value)
        End Set
    End Property

    ''' <summary>
    ''' Get / Set the serial port to send print jobs to.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SerialPrintPort() As System.IO.Ports.SerialPort
        Get
            Return clsSerialPrintPort
        End Get
        Set(ByVal value As System.IO.Ports.SerialPort)
            clsSerialPrintPort = value
        End Set
    End Property

    ''' <summary>
    ''' Returns the Printer's current state
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property PrinterState() As PrintingStates
        Get
            Return clsPrinterState
        End Get
    End Property

    ''' <summary>
    ''' Set Printer to Default Printer
    ''' </summary>
    ''' <remarks></remarks>
    ''' 
    Private Sub SetPrinter()
        Dim printerName As String = String.Empty
        Call SetPrinter(printerName)
    End Sub

    ''' <summary>
    ''' Set Printer to Printer identified by it's name
    ''' </summary>
    ''' <param name="printerName"></param>
    ''' <remarks></remarks>
    Private Sub SetPrinter(ByVal printerName As String)
        Dim foundPrinter As Boolean = False

        Try
            Dim printerCount As Integer = InstalledPrinters.Count

            If printerCount = 0 Then
                clsPrinterName = String.Empty
                Me.clsLastPrintError = "No installed printers found."
                Exit Sub
            End If

            If printerName.Trim.Length > 0 Then
                If Not IsValidPrinterName(printerName) Then
                    clsPrinterName = String.Empty
                    clsLastPrintError = "Invalid Printer Name"
                    Exit Sub
                End If
            End If

            For i As Integer = 0 To printerCount - 1
                If printerName.Trim <> String.Empty Then
                    If printerName.Trim.ToUpper = InstalledPrinters(i).Trim.ToUpper Then
                        Me.clsPrinterName = InstalledPrinters(i)
                        foundPrinter = True
                        Exit For
                    End If
                Else
                    Dim y As New PrintDocument

                    y.PrinterSettings.PrinterName = InstalledPrinters(i)
                    If y.PrinterSettings.IsDefaultPrinter Then
                        Me.clsPrinterName = InstalledPrinters(i)
                        foundPrinter = True
                        Exit For
                    End If
                End If
            Next

        Catch ex As Exception
            clsPrinterName = String.Empty
            clsLastPrintError = ex.Message
        End Try

        If Not foundPrinter Then
            clsPrinterName = String.Empty
            clsLastPrintError = "Invalid Printer Name"
            Exit Sub
        End If

    End Sub

#End Region

#Region "Unmanaged API References"

    ' Structure and API declarions:
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Private Structure DocInfo
        <MarshalAs(UnmanagedType.LPWStr)> Public pDocName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pOutputFile As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pDataType As String
    End Structure

    <DllImport("winspool.Drv", EntryPoint:="OpenPrinterW", _
       SetLastError:=True, CharSet:=CharSet.Unicode, _
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function OpenPrinter(ByVal src As String, ByRef hPrinter As IntPtr, _
    ByVal pd As IntPtr) As Boolean
    End Function

    <DllImport("winspool.Drv", EntryPoint:="ClosePrinter", _
       SetLastError:=True, CharSet:=CharSet.Unicode, _
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function ClosePrinter(ByVal hPrinter As IntPtr) As Boolean
    End Function

    <DllImport("winspool.Drv", EntryPoint:="StartDocPrinterW", _
       SetLastError:=True, CharSet:=CharSet.Unicode, _
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function StartDocPrinter(ByVal hPrinter As IntPtr, ByVal level As Int32, ByRef pDI As DocInfo) As Boolean
    End Function

    <DllImport("winspool.Drv", EntryPoint:="EndDocPrinter", _
       SetLastError:=True, CharSet:=CharSet.Unicode, _
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function EndDocPrinter(ByVal hPrinter As IntPtr) As Boolean
    End Function

    <DllImport("winspool.Drv", EntryPoint:="StartPagePrinter", _
       SetLastError:=True, CharSet:=CharSet.Unicode, _
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function StartPagePrinter(ByVal hPrinter As IntPtr) As Boolean
    End Function

    <DllImport("winspool.Drv", EntryPoint:="EndPagePrinter", _
       SetLastError:=True, CharSet:=CharSet.Unicode, _
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function EndPagePrinter(ByVal hPrinter As IntPtr) As Boolean
    End Function

    <DllImport("winspool.Drv", EntryPoint:="WritePrinter", _
       SetLastError:=True, CharSet:=CharSet.Unicode, _
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function WritePrinter(ByVal hPrinter As IntPtr, ByVal pBytes As IntPtr, ByVal dwCount As Int32, ByRef dwWritten As Int32) As Boolean
    End Function

#End Region

#Region "Class Public Functions"

    ''' <summary>
    ''' Returns the Case Sensitive Printer name for a printer name given
    ''' </summary>
    ''' <param name="szPrinterName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetProperPrinterName(ByVal szPrinterName As String) As String
        Dim printerName As String = String.Empty

        If IsValidPrinterName(szPrinterName) Then
            Dim instPrinters As New List(Of String)

            Dim printerCount As Integer = InstalledPrinters.Count
            For i As Integer = 0 To printerCount - 1
                If szPrinterName.Trim.ToUpper = InstalledPrinters(i).Trim.ToUpper Then
                    printerName = InstalledPrinters(i)
                    Exit For
                End If
            Next
        End If

        Return printerName
    End Function

    ''' <summary>
    ''' When the function is given a printer name and an unmanaged array of  
    ''' bytes, the function sends those bytes to the print queue.
    ''' Returns True on success or False on failure.
    ''' </summary>
    ''' <param name="szPrinterName"></param>
    ''' <param name="pBytes"></param>
    ''' <param name="dwCount"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SendBytesToPrinter(ByVal szPrinterName As String, ByVal pBytes As IntPtr, ByVal dwCount As Int32) As Boolean
        Return SendBytesToPrinter(szPrinterName, pBytes, dwCount, Application.ProductName, "RAW")
    End Function

    ''' <summary>
    ''' When the function is given a printer name and an unmanaged array of  
    ''' bytes, the function sends those bytes to the print queue.
    ''' Returns True on success or False on failure.
    ''' </summary>
    ''' <param name="szPrinterName"></param>
    ''' <param name="pBytes"></param>
    ''' <param name="dwCount"></param>
    ''' <param name="docName"></param>
    ''' <param name="docDataType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SendBytesToPrinter(ByVal szPrinterName As String, ByVal pBytes As IntPtr, _
        ByVal dwCount As Int32, ByVal docName As String, ByVal docDataType As String) As Boolean
        Dim hPrinter As IntPtr      ' The printer handle.
        Dim dwError As Int32        ' Last error - in case there was trouble.
        Dim di As New DocInfo          ' Describes your document (name, port, data type).
        Dim dwWritten As Int32      ' The number of bytes written by WritePrinter().
        Dim bSuccess As Boolean     ' Your success code.

        ' Set up the DOCINFO structure.
        With di
            .pDocName = docName
            .pDataType = docDataType
        End With

        ' Assume failure unless you specifically succeed.
        bSuccess = False
        If OpenPrinter(szPrinterName, hPrinter, 0&) Then
            If StartDocPrinter(hPrinter, 1, di) Then
                If StartPagePrinter(hPrinter) Then
                    ' Write your printer-specific bytes to the printer.
                    bSuccess = WritePrinter(hPrinter, pBytes, dwCount, dwWritten)
                    EndPagePrinter(hPrinter)
                End If
                EndDocPrinter(hPrinter)
            End If
            ClosePrinter(hPrinter)
        End If
        ' If you did not succeed, GetLastError may give more information
        ' about why not.
        If bSuccess = False Then
            dwError = Marshal.GetLastWin32Error()
        End If
        Return bSuccess
    End Function

    ''' <summary>
    ''' When the function is given a file name the function reads the contents of the file and sends the
    ''' contents to the printer defined by the user. If no printer was defined the defualt pritner is used
    ''' if one is found.
    ''' Presumes that the file contains printer-ready data
    ''' Returns True on success or False on failure.
    ''' </summary>
    ''' <param name="szFileName">Location of file to print</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SendFileToPrinter(ByVal szFileName As String) As Boolean
        If clsPrinterName.Trim = String.Empty Then
            SetPrinter()
        End If

        Return SendFileToPrinter(clsPrinterName, szFileName)
    End Function

    ''' <summary>
    ''' When the function is given a file name and a printer name, the function reads the contents of the file and sends the
    ''' contents to the printer.
    ''' Presumes that the file contains printer-ready data.
    ''' Returns True on success or False on failure.
    ''' </summary>
    ''' <param name="szPrinterName"></param>
    ''' <param name="szFileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SendFileToPrinter(ByVal szPrinterName As String, ByVal szFileName As String) As Boolean
        Return SendFileToPrinter(szPrinterName, szFileName, Application.ProductName, "RAW")
    End Function

    ''' <summary>
    ''' When the function is given a file name and a printer name, the function reads the contents of the file and sends the
    ''' contents to the printer.
    ''' Presumes that the file contains printer-ready data.
    ''' Returns True on success or False on failure.
    ''' </summary>
    ''' <param name="szPrinterName"></param>
    ''' <param name="szFileName"></param>
    ''' <param name="docName"></param>
    ''' <param name="docDataType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SendFileToPrinter(ByVal szPrinterName As String, ByVal szFileName As String, ByVal docName As String, ByVal docDataType As String) As Boolean

        Dim bSuccess As Boolean = False

        Try
            ' Open the file.
            Dim fs As New FileStream(szFileName, FileMode.Open)
            ' Create a BinaryReader on the file.

            Dim br As New BinaryReader(fs)
            ' Dim an array of bytes large enough to hold the file's contents.
            Dim bytes(fs.Length) As Byte
            ' Your unmanaged pointer.
            Dim pUnmanagedBytes As IntPtr

            ' Read the contents of the file into the array.
            bytes = br.ReadBytes(fs.Length)
            ' Allocate some unmanaged memory for those bytes.
            pUnmanagedBytes = Marshal.AllocCoTaskMem(fs.Length)
            ' Copy the managed byte array into the unmanaged array.
            Marshal.Copy(bytes, 0, pUnmanagedBytes, fs.Length)
            ' Send the unmanaged bytes to the printer.
            bSuccess = SendBytesToPrinter(szPrinterName, pUnmanagedBytes, fs.Length, docName, docDataType)
            ' Free the unmanaged memory that you allocated earlier.
            Marshal.FreeCoTaskMem(pUnmanagedBytes)

            fs.Dispose()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Send File To Printer", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return bSuccess
    End Function

    ''' <summary>
    ''' Sends the contents of a byte array to the printer identifed by the serial connection
    ''' </summary>
    ''' <param name="BytesToSend"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SendBytesToSerailPrintPort(ByVal BytesToSend() As Byte) As Boolean

        If clsSerialPrintPort IsNot Nothing Then
            Try
                If Not clsSerialPrintPort.IsOpen Then
                    clsSerialPrintPort.Open()
                End If

                Dim refreshEncoding As System.Text.Encoding = System.Text.Encoding.ASCII

                refreshEncoding = clsSerialPrintPort.Encoding

                clsSerialPrintPort.Encoding = System.Text.Encoding.UTF8
                clsSerialPrintPort.Write(BytesToSend, 0, BytesToSend.Length)
                SendBytesToSerailPrintPort = True

                clsSerialPrintPort.Encoding = refreshEncoding

            Catch ex As Exception
                SendBytesToSerailPrintPort = False
            End Try
        Else
            Return False
        End If

        Return SendBytesToSerailPrintPort

    End Function

    ''' <summary>
    ''' Sends a string of data to the printer as raw bytes
    ''' </summary>
    ''' <param name="szPrinterName"></param>
    ''' <param name="szString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SendStringToPrinter(ByVal szPrinterName As String, ByVal szString As String) As Boolean

        SendStringToPrinter = True

        Dim Sql As String = String.Empty

        If clsSerialPrintPort IsNot Nothing Then
            Try
                If Not clsSerialPrintPort.IsOpen Then
                    clsSerialPrintPort.Open()
                End If

                clsSerialPrintPort.WriteLine(szString)

            Catch ex As Exception
                Return False
            End Try
        Else
            Try
                Dim pBytes As IntPtr
                Dim dwCount As Int32

                ' How many characters are in the string?
                dwCount = szString.Length()

                ' Assume that the printer is expecting ANSI text, and then convert the string to ANSI text.
                pBytes = Marshal.StringToCoTaskMemAnsi(szString)

                ' Send the converted ANSI string to the printer.
                SendBytesToPrinter(szPrinterName, pBytes, dwCount)
                Marshal.FreeCoTaskMem(pBytes)

            Catch ex As Exception
                Return False
            End Try
        End If

        Return True
    End Function

#End Region

#Region "Line Printing Routines"

    ''' <summary>
    ''' Opens the default Printer for Printing.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function OpenPrinter() As Boolean
        ' If no printer is defined then set to default printer
        If Me.clsPrinterName.Trim = String.Empty Then
            SetPrinter()

            If clsPrinterName.Trim = String.Empty Then
                Me.clsLastPrintError = "No installed or default printer found."
                Return False
                Exit Function
            End If

        End If

        Return OpenPrinter(clsPrinterName)

    End Function

    ''' <summary>
    ''' Opens the specified Printer for Printing.
    ''' </summary>
    ''' <param name="szPrinterName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function OpenPrinter(ByVal szPrinterName As String) As Boolean

        Dim di As New DocInfo

        ' Assume failure unless you specifically succeed.
        Dim bSuccess As Boolean = False

        If clsPrinterHandle <> 0 Then
            Try
                Call ClosePrinter(clsPrinterHandle)
            Catch ex As Exception
                clsLastPrintError = ex.Message
                Return False
                Exit Function
            End Try
            clsPrinterHandle = 0
        End If

        Dim ProperPrinterName As String = GetProperPrinterName(szPrinterName)
        If ProperPrinterName = String.Empty Then
            clsLastPrintError = "'" & szPrinterName & "' is not a valid printer name"
            Return bSuccess
            Exit Function
        End If

        Me.clsPrinterName = ProperPrinterName

        Try
            ' Set up the DOCINFO structure.
            With di
                .pDocName = Application.ProductName
                .pDataType = Me.clsPrintDataType.ToString
            End With

            ' Try to Open the Printer
            If OpenPrinter(clsPrinterName, clsPrinterHandle, 0&) Then
                clsPrinterState = Printers.PrintingStates.Open
                ' Try to start a document
                If StartDocPrinter(clsPrinterHandle, 1, di) Then
                    clsPrinterState = Printers.PrintingStates.StartDoc
                    ' Try to start the new page
                    If StartPagePrinter(clsPrinterHandle) Then
                        clsPrinterState = Printers.PrintingStates.StartPage
                        bSuccess = True
                    End If
                End If
            End If

        Catch ex As Exception
            clsLastPrintError = ex.Message
        End Try

        Return bSuccess

    End Function

    ''' <summary>
    ''' Closes a Printer
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReleasePrinter() As Boolean
        Dim bSuccess As Boolean = False
        Try
            If Me.clsPrinterHandle = 0 Then
                Me.clsLastPrintError = "Invalid Printer Handle "
                Return bSuccess
            End If

            If ClosePrinter(Me.clsPrinterHandle) Then
                Me.clsPrinterState = Printers.PrintingStates.Closed
                bSuccess = True
                Me.clsPrinterHandle = 0
            End If

        Catch ex As Exception
            Me.clsLastPrintError = ex.Message
        End Try
        Return bSuccess

    End Function

    ''' <summary>
    ''' Sends a String of text to the printer assigned the provided printer handle
    ''' </summary>
    ''' <param name="pString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PrintString(ByVal pString As String) As Boolean

        Dim bSuccess As Boolean = False
        Dim dwWritten As Int32

        If Me.clsPrinterHandle = 0 Then
            Me.clsLastPrintError = "Invalid Printer Handle"
            Exit Function
        End If

        Try
            ' Unmanaged pointer.
            Dim bytes(pString.Length) As Byte
            Dim pUnmanagedBytes As IntPtr
            Dim encoding As New System.Text.ASCIIEncoding()
            bytes = encoding.GetBytes(pString)

            ' Allocate some unmanaged memory for those bytes.
            pUnmanagedBytes = Marshal.AllocCoTaskMem(pString.Length)

            ' Copy the managed byte array into the unmanaged array.
            Marshal.Copy(bytes, 0, pUnmanagedBytes, pString.Length)

            ' Send the unmanaged bytes to the printer.
            bSuccess = WritePrinter(Me.clsPrinterHandle, pUnmanagedBytes, pString.Length, dwWritten)

            ' Free the unmanaged memory that you allocated earlier.
            Marshal.FreeCoTaskMem(pUnmanagedBytes)

        Catch ex As Exception
            Me.clsLastPrintError = ex.Message
        End Try

        Return bSuccess

    End Function

    ''' <summary>
    ''' Starts a page for the printer
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function StartPage() As Boolean
        Dim bSuccess As Boolean = False
        Try

            If Me.clsPrinterHandle = 0 Then
                Me.clsLastPrintError = "Invalid Printer Handle"
                Return bSuccess
            End If

            If StartPagePrinter(Me.clsPrinterHandle) Then
                Me.clsPrinterState = Printers.PrintingStates.StartPage
                bSuccess = True
            End If

        Catch ex As Exception
            Me.clsLastPrintError = ex.Message
        End Try

        Return bSuccess
    End Function

    ''' <summary>
    ''' Ends the page for a printer
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function EndPage() As Boolean
        Dim bSuccess As Boolean = False

        Try
            If Me.clsPrinterHandle = 0 Then
                Me.clsLastPrintError = "Invalid Printer Handle"
                Return bSuccess
            End If

            If EndPagePrinter(Me.clsPrinterHandle) Then
                Me.clsPrinterState = Printers.PrintingStates.EndPage
                bSuccess = True
            End If

        Catch ex As Exception
            Me.clsLastPrintError = ex.Message
            bSuccess = False
        End Try
        Return bSuccess

    End Function

    ''' <summary>
    ''' Starts a document for a printer
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function StartDoc() As Boolean
        Dim bSuccess As Boolean = False

        Try

            If Me.clsPrinterHandle = 0 Then
                Me.clsLastPrintError = "Invalid Printer Handle"
                Return bSuccess
            End If

            Dim di As New DocInfo
            With di
                .pDocName = Application.ProductName
                .pDataType = Me.clsPrintDataType.ToString
            End With

            If StartDocPrinter(Me.clsPrinterHandle, 1, di) Then
                Me.clsPrinterState = Printers.PrintingStates.StartDoc
                bSuccess = True
            End If

        Catch ex As Exception
            Me.clsLastPrintError = ex.Message
        End Try

        Return bSuccess

    End Function

    ''' <summary>
    ''' Ends a document for a printer
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function EndDoc() As Boolean

        Dim bSuccess As Boolean = False

        Try

            If Me.clsPrinterHandle = 0 Then
                Me.clsLastPrintError = "Invalid Printer Handle."
                Return bSuccess
            End If

            If Me.clsPrinterState = PrintingStates.Closed Then
                Me.clsLastPrintError = "Printer is closed."
                Return bSuccess
            End If

            If Me.clsPrinterState = PrintingStates.Open Then
                Me.clsLastPrintError = "Print document not started."
                Return bSuccess
            End If

            If Me.clsPrinterState = Printers.PrintingStates.StartPage Then
                Me.PrintString(Microsoft.VisualBasic.vbFormFeed)
                Me.EndPage()
            End If

            If EndDocPrinter(Me.clsPrinterHandle) Then
                Me.clsPrinterState = Printers.PrintingStates.EndDoc
                bSuccess = True
            End If

        Catch ex As Exception
            Me.clsLastPrintError = ex.Message
            Return False
        End Try

    End Function

#End Region

#Region "Classes Internal Private Classes"

    Private Class Printers

        ''' <summary>
        ''' Various states of a printer
        ''' </summary>
        ''' <remarks></remarks>
        Enum PrintingStates
            Open
            StartDoc
            StartPage
            EndPage
            EndDoc
            Closed
        End Enum

        Private _PrinterName As String = String.Empty
        Private _hPrinter As IntPtr = 0
        Private _lastPrintError As String = String.Empty
        Private _PrintingState As PrintingStates = PrintingStates.Closed

        ''' <summary>
        ''' Class Constructor
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            _PrinterName = String.Empty
            _hPrinter = 0
            _lastPrintError = String.Empty
            _PrintingState = PrintingStates.Closed
        End Sub

        ''' <summary>
        ''' Get / Set Printer Name
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property PrinterName()
            Get
                Return _PrinterName
            End Get
            Set(ByVal value)
                _PrinterName = value
            End Set
        End Property

        ''' <summary>
        ''' Get / Set Printer Handle
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property PrnterHandle() As IntPtr
            Get
                Return _hPrinter
            End Get
            Set(ByVal value As IntPtr)
                _hPrinter = value
            End Set
        End Property

        ''' <summary>
        ''' Get / Set Last Printer Error
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property LastPrintError() As String
            Get
                Return _lastPrintError
            End Get
            Set(ByVal value As String)
                _lastPrintError = value
            End Set
        End Property

        ''' <summary>
        ''' Gets / Sets the current state of the printer in use
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property PrinterState() As PrintingStates
            Get
                Return _PrintingState
            End Get
            Set(ByVal value As PrintingStates)
                _PrintingState = value
            End Set
        End Property

    End Class

    ''' <summary>
    ''' Class to print to Line printers like the OkiData Microline 186
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ASCPRTLP

        ' Inherits all the functionality of a PrintDocument
        Inherits System.Drawing.Printing.PrintDocument

        ' Private variables to hold default font and text
        Private fntPrintFont As Font
        Private _textToPrint As String
        Private _printerName As String

#Region "Class Constuctors"

        ''' <summary>
        ''' Class Constructor
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            MyBase.New()
            Me.InitializeClass()
        End Sub

        ''' <summary>
        ''' Class Constructor
        ''' </summary>
        ''' <param name="vPrinterName">Printer Name Inches</param>
        ''' <param name="PaperHeight">Paper Height Inches</param>
        ''' <param name="PaperWidth">Paper Width Inches</param>
        ''' <param name="MarginLeft">Left Margin Inches</param>
        ''' <param name="MarginRight">Right Margin Inches</param>
        ''' <param name="MarginTop">Top Margin Inches</param>
        ''' <param name="MarginBottom">Bottom Margin Inches</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal vPrinterName As String, ByVal PaperHeight As Double, ByVal PaperWidth As Double, _
            ByVal MarginLeft As Double, ByVal MarginRight As Double, ByVal MarginTop As Double, ByVal MarginBottom As Double)

            ' Sets the file stream
            MyBase.New()
            Me.InitializeClass()

            Dim ps As New Drawing.Printing.PaperSize("LBL", CInt(PaperWidth * 100), CInt(PaperHeight * 100))
            MyBase.DefaultPageSettings.PaperSize = ps
            MyBase.DefaultPageSettings.Margins.Left = CInt(MarginLeft * 100)
            MyBase.DefaultPageSettings.Margins.Right = CInt(MarginRight * 100)
            MyBase.DefaultPageSettings.Margins.Top = CInt(MarginTop * 100)
            MyBase.DefaultPageSettings.Margins.Bottom = CInt(MarginBottom * 100)

            Me.PrinterName = vPrinterName
        End Sub

        ''' <summary>
        ''' Neoware returns Printer name with garbage appended to it. This passes back the Printer name
        ''' best associated witht he printer name Provided
        ''' </summary>
        ''' <param name="PrinterName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function VerifyPrinterName(ByVal PrinterName As String) As String

            Dim printerNameToUse As String = String.Empty
            Dim printerNameLike As String = String.Empty
            Dim printerNameSessionID As String = String.Empty
            Dim printerClientNameSessionID As String = String.Empty

            Dim UserName As String = My.User.Name
            Dim SessionID As String = "Session " & ASCMAIN1.WTS_SESSION_ID.ToString.Trim
            Dim ClientName As String = System.Environment.GetEnvironmentVariable("CLIENTNAME")


            If UserName Is Nothing Then UserName = String.Empty
            If SessionID Is Nothing Then SessionID = String.Empty
            If ClientName Is Nothing Then ClientName = String.Empty


            VerifyPrinterName = PrinterName.ToUpper

            UserName = UserName.ToUpper
            SessionID = SessionID.ToUpper

            Dim printers As New System.Drawing.Printing.PrintDocument()
            Dim printerSessionId As String = String.Empty

            For Each installedPrinterName As String In PrinterSettings.InstalledPrinters

                If InStr(installedPrinterName.ToUpper, "SESSION") > 0 Then
                    printerSessionId = installedPrinterName.Substring(InStr(installedPrinterName.ToUpper, "SESSION") - 1).Trim.ToUpper
                Else
                    printerSessionId = String.Empty
                End If

                If printerNameToUse.Length > 0 And printerNameLike.Length > 0 And printerNameSessionID.Length = 0 Then
                    Exit For
                End If

                If installedPrinterName.ToUpper = VerifyPrinterName Then
                    printerNameToUse = installedPrinterName
                End If

                If InStr(installedPrinterName.ToUpper, VerifyPrinterName) = 1 Then
                    printerNameLike = installedPrinterName
                End If

                If InStr(installedPrinterName.ToUpper, VerifyPrinterName) = 1 And installedPrinterName.ToUpper.Contains(UserName) And installedPrinterName.ToUpper.Contains(SessionID) _
                        And UserName.Length > 0 And SessionID.Length > 0 And SessionID = printerSessionId Then
                    printerNameSessionID = installedPrinterName
                End If

                If InStr(installedPrinterName.ToUpper, VerifyPrinterName) = 1 And installedPrinterName.ToUpper.Contains(ClientName) And installedPrinterName.ToUpper.Contains(SessionID) _
                            And ClientName.Length > 0 And SessionID.Length > 0 And SessionID = printerSessionId Then
                    printerClientNameSessionID = installedPrinterName
                End If
            Next

            If printerNameToUse.Length > 0 Then
                Return printerNameToUse
            ElseIf printerNameSessionID.Length > 0 Then
                Return printerNameSessionID
            ElseIf printerClientNameSessionID.Length > 0 Then
                Return printerClientNameSessionID
            ElseIf printerNameLike.Length > 0 Then
                Return printerNameLike
            Else
                Return PrinterName
            End If

        End Function

        ''' <summary>
        ''' Initializes Class Variables
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub InitializeClass()

            fntPrintFont = New Font("COURIER NEW", 12)

            Dim ps As New Drawing.Printing.PaperSize("LBL", 850, 1100)

            MyBase.DefaultPageSettings.PaperSize = ps
            MyBase.DefaultPageSettings.Margins.Left = 100
            MyBase.DefaultPageSettings.Margins.Right = 100
            MyBase.DefaultPageSettings.Margins.Top = 100
            MyBase.DefaultPageSettings.Margins.Bottom = 100

            MyBase.DefaultPageSettings.Landscape = False

        End Sub

#End Region

        ''' <summary>
        ''' Returns a list of the Installed Printers
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property GetInstalledPrinters() As List(Of String)
            Get
                Dim instPrinters As New List(Of String)

                For printerCount As Integer = 0 To System.Drawing.Printing.PrinterSettings.InstalledPrinters.Count - 1
                    If System.Drawing.Printing.PrinterSettings.InstalledPrinters(printerCount).ToUpper = PrinterName Then
                        instPrinters.Add(System.Drawing.Printing.PrinterSettings.InstalledPrinters(printerCount))
                    End If
                Next

                Return instPrinters
            End Get
        End Property

        ''' <summary>
        ''' Emuneration to set the pages orientation
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum Orientation
            Landscape
            Portrait
        End Enum

        ''' <summary>
        ''' Gets / Sets the text to be sent to the printer
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property TextToPrint() As String
            Get
                Return _textToPrint
            End Get
            Set(ByVal Value As String)
                _textToPrint = Value
            End Set
        End Property

        ''' <summary>
        ''' Gets / Sets the Printer Font
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Font() As Font
            ' Allows the user to override the default font
            Get
                Return fntPrintFont
            End Get
            Set(ByVal Value As Font)
                Try
                    fntPrintFont = Value
                Catch ex As Exception
                    ' Nothing
                End Try
            End Set
        End Property

        ''' <summary>
        ''' Gets / Sets the page Left margin
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property LeftMargin() As Double
            Get
                Return MyBase.DefaultPageSettings.Margins.Left / 100
            End Get
            Set(ByVal value As Double)
                MyBase.DefaultPageSettings.Margins.Left = CInt(value * 100)
            End Set
        End Property

        ''' <summary>
        ''' Gets / Sets the page Top margin
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property TopMargin() As Double
            Get
                Return MyBase.DefaultPageSettings.Margins.Top / 100
            End Get
            Set(ByVal value As Double)
                MyBase.DefaultPageSettings.Margins.Top = CInt(value * 100)
            End Set
        End Property

        ''' <summary>
        ''' Gets / Sets the page Bottom margin
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property BottomMargin() As Double
            Get
                Return MyBase.DefaultPageSettings.Margins.Bottom / 100
            End Get
            Set(ByVal value As Double)
                MyBase.DefaultPageSettings.Margins.Bottom = CInt(value * 100)
            End Set
        End Property

        ''' <summary>
        ''' Gets / Sets the page Right margin
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property RightMargin() As Double
            Get
                Return MyBase.DefaultPageSettings.Margins.Right / 100
            End Get
            Set(ByVal value As Double)
                MyBase.DefaultPageSettings.Margins.Right = CInt(value * 100)
            End Set
        End Property

        ''' <summary>
        ''' Gets / Sets the Page Paper Width
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property PaperWidth() As Double
            Get
                Return MyBase.DefaultPageSettings.PaperSize.Width / 100
            End Get
            Set(ByVal value As Double)
                Dim ps As New Drawing.Printing.PaperSize("LBL", CInt(value * 100), CInt(PaperHeight * 100))
                MyBase.DefaultPageSettings.PaperSize = ps
            End Set
        End Property

        ''' <summary>
        ''' Gets / Sets the Page Paper Height
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property PaperHeight() As Double
            Get
                Return MyBase.DefaultPageSettings.PaperSize.Height / 100
            End Get
            Set(ByVal value As Double)
                Dim ps As New Drawing.Printing.PaperSize("LBL", CInt(PaperWidth * 100), CInt(value * 100))
                MyBase.DefaultPageSettings.PaperSize = ps
            End Set
        End Property

        ''' <summary>
        ''' Gets / Sets the Page Paper Orientation
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property PageOrientation() As Orientation
            Get
                If MyBase.DefaultPageSettings.Landscape = False Then
                    Return Orientation.Portrait
                Else
                    Return Orientation.Landscape
                End If
            End Get
            Set(ByVal value As Orientation)
                If value = Orientation.Portrait Then
                    MyBase.DefaultPageSettings.Landscape = False
                Else
                    MyBase.DefaultPageSettings.Landscape = True
                End If
            End Set
        End Property

        ''' <summary>
        ''' Gets / Sets the Printer Name
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property PrinterName() As String
            Get
                Return MyBase.PrinterSettings.PrinterName
            End Get

            Set(ByVal value As String)
                value = value.Trim

                Dim printerNameToUse As String = Me.VerifyPrinterName(value)

                If printerNameToUse.Length > 0 Then
                    MyBase.PrinterSettings.PrinterName = printerNameToUse
                    Me._printerName = printerNameToUse
                Else
                    MyBase.PrinterSettings.PrinterName = PrinterName
                    Me._printerName = PrinterName
                End If
            End Set
        End Property

        ''' <summary>
        ''' Sends text to the Selecter Printer
        ''' </summary>
        ''' <param name="textToPrint">Text to send to the printer</param>
        ''' <remarks></remarks>
        Public Sub PrintText(ByVal textToPrint As String)
            _textToPrint = textToPrint

            MyBase.Print()
        End Sub

        Protected Overrides Sub OnBeginPrint(ByVal ev As System.Drawing.Printing.PrintEventArgs)
            ' Run base code
            MyBase.OnBeginPrint(ev)

            ' Sets the default font
            If fntPrintFont Is Nothing Then
                'fntPrintFont = New Font("Times New Roman", 12)
                fntPrintFont = New Font("COURIER NEW", 12)
            End If
        End Sub

        Protected Overrides Sub OnPrintPage(ByVal ev As System.Drawing.Printing.PrintPageEventArgs)
            ' Provides the print logic for our document

            ' Run base code
            MyBase.OnPrintPage(ev)

            ' Variables
            Static intCurrentChar As Integer

            Dim intPrintAreaHeight, intPrintAreaWidth, intMarginLeft, intMarginTop As Integer

            ' Set printing area boundaries and margin coordinates
            With MyBase.DefaultPageSettings
                intPrintAreaHeight = .PaperSize.Height - _
                                   .Margins.Top - .Margins.Bottom
                intPrintAreaWidth = .PaperSize.Width - _
                                  .Margins.Left - .Margins.Right
                intMarginLeft = .Margins.Left 'X
                intMarginTop = .Margins.Top   'Y
            End With

            ' If Landscape set, swap printing height/width
            If MyBase.DefaultPageSettings.Landscape Then
                Dim intTemp As Integer
                intTemp = intPrintAreaHeight
                intPrintAreaHeight = intPrintAreaWidth
                intPrintAreaWidth = intTemp
            End If

            ' Calculate total number of lines
            Dim intLineCount As Int32 = CInt(intPrintAreaHeight / fntPrintFont.Height)

            ' Initialise rectangle printing area
            Dim rectPrintingArea As New RectangleF(intMarginLeft, intMarginTop, intPrintAreaWidth, intPrintAreaHeight)

            ' Initialize StringFormat class, for text layout
            Dim objSF As New StringFormat(StringFormatFlags.LineLimit)

            ' Figure out how many lines will fit into rectangle
            Dim intLinesFilled, intCharsFitted As Int32
            ev.Graphics.MeasureString(Mid(_textToPrint, _
                        UpgradeZeros(intCurrentChar)), fntPrintFont, _
                        New SizeF(intPrintAreaWidth, _
                        intPrintAreaHeight), objSF, _
                        intCharsFitted, intLinesFilled)

            ' Print the text to the page
            ev.Graphics.DrawString(Mid(_textToPrint, _
                UpgradeZeros(intCurrentChar)), fntPrintFont, _
                Brushes.Black, rectPrintingArea, objSF)

            ' Increase current char count
            intCurrentChar += intCharsFitted

            ' Check whether we need to print more
            If intCurrentChar < _textToPrint.Length Then
                ev.HasMorePages = True
            Else
                ev.HasMorePages = False
                intCurrentChar = 0
            End If
        End Sub

        Private Function UpgradeZeros(ByVal Input As Integer) As Integer
            ' Upgrades all zeros to ones
            ' - used as opposed to defunct IIF or messy If statements
            If Input = 0 Then
                Return 1
            Else
                Return Input
            End If
        End Function

        ''' <summary>
        ''' Returns whether a selected printer is installed on
        ''' </summary>
        ''' <param name="PrinterName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsValidPrinterName(ByVal PrinterName As String) As Boolean

            IsValidPrinterName = False
            PrinterName = PrinterName.Trim.ToUpper

            For printerCount As Integer = 0 To System.Drawing.Printing.PrinterSettings.InstalledPrinters.Count - 1
                If System.Drawing.Printing.PrinterSettings.InstalledPrinters(printerCount).ToUpper = PrinterName Then
                    Return True
                End If
            Next

        End Function

    End Class

#End Region
End Class