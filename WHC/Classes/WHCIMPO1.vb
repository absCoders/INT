Imports GemBox.Spreadsheet
Imports Oracle.ManagedDataAccess.Client

Public Class WHCIMPO1
    ' Create Anticipated Receipts File (Outbound) 943

    Inherits WHC000O1

    Private selectedLayoutList As New List(Of String)
    Private layoutsDirectory As String = String.Empty
    Private dataDirectory As String = String.Empty
    Private dataDirectoryArchive As String = String.Empty

    Sub New()
        MyBase.New()
    End Sub

    Sub New(g As ABSEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_OBJECT = "WHCIMPO1"

        Create_Work_Table()

        With dst
            ASCMAIN1.sql = "Select * from CONV.CFG_FILE"
            Create_TDA(.Tables.Add("TATCONVX"), "CONV.CFG_FILE", "**", 0, True)

            ASCMAIN1.sql = "Select * from CONV.CFG_LAYOUT"
            Create_TDA(.Tables.Add("TATCONVY"), "CONV.CFG_LAYOUT", "**", 0, True)

        End With

        Main_Process()
        DisposeOPD()
    End Sub

    Public Sub Main_Process()

        EnforceConstraints(False)

        layoutsDirectory = String.Empty
        dataDirectory = String.Empty
        Dim subDir As String = DateTime.Now.ToString("yyyyMM")
        Dim dupsDir As String = String.Empty

        Dim serverName As String = G.DBS_SERVER

        If G.DBS_SERVER <> G.DBS_COMPANY Then
            serverName = "TST"
        End If

        Select Case (G.APP_CMD & String.Empty).ToUpper
            Case "SHIP", "RETURNS"
                If (G.APP_KEY & String.Empty).Trim.Length > 0 Then
                    selectedLayoutList = (G.APP_KEY & String.Empty).Split(",").ToList
                End If

                layoutsDirectory = ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT") & "\cfg\" _
                    & IIf(serverName = "TST", "TEST", "PROD") & "\Layouts\"

                dataDirectory = ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT") & "\cfg\" _
                    & IIf(serverName = "TST", "TEST", "PROD") & "\FROM_CUSA\"

                dataDirectoryArchive = dataDirectory & "Archive\" & subDir & "\"

            Case Else
                If (G.APP_KEY & String.Empty).Trim.Length > 0 Then
                    selectedLayoutList = (G.APP_KEY & String.Empty).Split(",").ToList
                End If

                layoutsDirectory = ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT") & "\cfg\" _
                    & IIf(serverName = "TST", "TEST", "PROD") & "\Layouts\"

                dataDirectory = ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT") & "\cfg\" _
                    & IIf(serverName = "TST", "TEST", "PROD") & "\FROM_CUSA\"

                dataDirectoryArchive = dataDirectory & "Archive\" & subDir & "\"
        End Select

        dupsDir = dataDirectoryArchive & "DUPS\"
        Try
            If Not My.Computer.FileSystem.DirectoryExists(dupsDir) Then
                My.Computer.FileSystem.CreateDirectory(dupsDir)
            End If
        Catch ex As Exception

        End Try


        For ictr As Int16 = 0 To selectedLayoutList.Count - 1
            selectedLayoutList(ictr) = selectedLayoutList(ictr).ToUpper
        Next

        If layoutsDirectory.Length = 0 OrElse Not My.Computer.FileSystem.DirectoryExists(layoutsDirectory) Then
            Throw New Exception("Invalid Layout Directory: " & layoutsDirectory)
            Exit Sub
        End If

        If dataDirectory.Length = 0 OrElse Not My.Computer.FileSystem.DirectoryExists(dataDirectory) Then
            Throw New Exception("Invalid Data Directory: " & dataDirectory)
            Exit Sub
        End If

        ' If the Archive direcory does not exist then create it as a courtesy
        If Not My.Computer.FileSystem.DirectoryExists(dataDirectoryArchive) Then
            My.Computer.FileSystem.CreateDirectory(dataDirectoryArchive)
        End If

        ' Global Variable that determines the number of records processed.
        R = 0
        Dim dataFiles As New List(Of String)

        ' Work Around for files with same names in Conversion data
        If (G.APP_CMD & String.Empty).ToUpper = "SHIP" Then

            For Each dataFileName As String In My.Computer.FileSystem.GetFiles(dataDirectory)

                Dim fileName As String = My.Computer.FileSystem.GetName(dataFileName)
                Dim renamedFile As String = String.Empty

                If fileName.StartsWith("INVTRANS_", StringComparison.OrdinalIgnoreCase) Then
                    'My.Computer.FileSystem.RenameFile(dataFileName, fileName.ToUpper.Replace("INVTRANS_", "INVTRANSADJ_").ToLower)
                    renamedFile = dataFileName.ToUpper.Replace("INVTRANS_", "INVTRANSADJ_").ToLower
                ElseIf fileName.StartsWith("RECEIPTS_", StringComparison.OrdinalIgnoreCase) Then
                    'My.Computer.FileSystem.RenameFile(dataFileName, fileName.ToUpper.Replace("RECEIPTS_", "RECEIPTSADJ_").ToLower)
                    renamedFile = dataFileName.ToUpper.Replace("RECEIPTS_", "RECEIPTSADJ_").ToLower
                ElseIf fileName.StartsWith("BRAND_", StringComparison.OrdinalIgnoreCase) Then
                    'My.Computer.FileSystem.RenameFile(dataFileName, fileName.ToUpper.Replace("BRAND_", "BRANDIPLB_").ToLower)
                    renamedFile = dataFileName.ToUpper.Replace("BRAND_", "BRANDIPLB_").ToLower
                ElseIf fileName.StartsWith("SUBBRAND_", StringComparison.OrdinalIgnoreCase) Then
                    'My.Computer.FileSystem.RenameFile(dataFileName, fileName.ToUpper.Replace("SUBBRAND_", "SUBBRANDIPLB_").ToLower)
                    renamedFile = dataFileName.ToUpper.Replace("SUBBRAND_", "SUBBRANDIPLB_").ToLower
                ElseIf fileName.StartsWith("INVENTORY_", StringComparison.OrdinalIgnoreCase) Then
                    'My.Computer.FileSystem.RenameFile(dataFileName, fileName.ToUpper.Replace("INVENTORY_", "INVENTORYIPLB_").ToLower)
                    renamedFile = dataFileName.ToUpper.Replace("INVENTORY_", "INVENTORYIPLB_").ToLower
                ElseIf fileName.StartsWith("INVOICE_", StringComparison.OrdinalIgnoreCase) Then
                    'My.Computer.FileSystem.RenameFile(dataFileName, fileName.ToUpper.Replace("INVOICE_", "INVOICEIPLB_").ToLower)
                    renamedFile = dataFileName.ToUpper.Replace("INVOICE_", "INVOICEIPLB_").ToLower
                ElseIf fileName.StartsWith("SHIPRULE_", StringComparison.OrdinalIgnoreCase) Then
                    'My.Computer.FileSystem.RenameFile(dataFileName, fileName.ToUpper.Replace("SHIPRULE_", "SHIPRULEIPLB_").ToLower)
                    renamedFile = dataFileName.ToUpper.Replace("SHIPRULE_", "SHIPRULEIPLB_").ToLower
                ElseIf fileName.StartsWith("CUSTMAST_", StringComparison.OrdinalIgnoreCase) Then
                    'My.Computer.FileSystem.RenameFile(dataFileName, fileName.ToUpper.Replace("CUSTMAST_", "CUSTMASTIPLB_").ToLower)
                    renamedFile = dataFileName.ToUpper.Replace("CUSTMAST_", "CUSTMASTIPLB_").ToLower
                End If

                ' See if we need to rname the file, sFTP and service run at the same time and on occasion the files are duplicated
                If renamedFile.Length > 0 Then
                    If Not My.Computer.FileSystem.FileExists(renamedFile) Then
                        My.Computer.FileSystem.RenameFile(dataFileName, My.Computer.FileSystem.GetName(renamedFile))
                    Else
                        ' Compare the two files if they are the same then move the file
                        Dim file1byte As Int32 = 0
                        Dim file2byte As Int32 = 0
                        Dim fs1 As System.IO.FileStream
                        Dim fs2 As System.IO.FileStream
                        Dim filesAreIdentical As Boolean = False

                        ' Open the two files.
                        fs1 = New System.IO.FileStream(dataFileName, System.IO.FileMode.Open)
                        fs2 = New System.IO.FileStream(renamedFile, System.IO.FileMode.Open)

                        ' Check the file sizes. If they are not the same, the files are not equal.
                        If (fs1.Length = fs2.Length) Then

                            ' Read and compare a byte from each file until either a non-matching set of bytes is found or until the end of file1 is reached.
                            Do
                                ' Read one byte from each file.
                                file1byte = fs1.ReadByte()
                                file2byte = fs2.ReadByte()
                            Loop While ((file1byte = file2byte) AndAlso (file1byte <> -1))

                            ' Return the success of the comparison. "file1byte" is equal to "file2byte" at this point only if the files are the same.
                            filesAreIdentical = ((file1byte - file2byte) = 0)

                        End If

                        ' Close the files.
                        fs1.Close()
                        fs2.Close()

                        If filesAreIdentical Then
                            My.Computer.FileSystem.DeleteFile(dataFileName)
                        Else
                            Dim errorMessage As String = String.Empty
                            errorMessage = "WARNING: Data file " & dataFileName
                            errorMessage &= " was renamed to (" & "@@@" & ") and placed in the DUPS directory."
                            errorMessage &= " File " & renamedFile & " exists and is not identical to " & dataFileName
                            errorMessage = errorMessage.Replace("@@@", dataFileName & "_" & DateTime.Now.ToString("yyyyMMddhhmmss"))
                            Me.ErrorMessages.Add(errorMessage)

                            My.Computer.FileSystem.MoveFile(dataFileName, dupsDir & My.Computer.FileSystem.GetName(dataFileName) & "_" & DateTime.Now.ToString("yyyyMMddhhssmm"), True)
                        End If

                    End If
                End If
            Next
        End If

        Dim cycleAgainToSeeIfThereAreMoreFiles As Boolean = True

        While cycleAgainToSeeIfThereAreMoreFiles
            cycleAgainToSeeIfThereAreMoreFiles = False

            For Each dataFileName As String In My.Computer.FileSystem.GetFiles(dataDirectory)
                Dim layoutFilePrefix As String = My.Computer.FileSystem.GetName(dataFileName)
                layoutFilePrefix = layoutFilePrefix.Split("_")(0)
                If selectedLayoutList.Count > 0 AndAlso selectedLayoutList.Contains(layoutFilePrefix.ToUpper) Then
                    ' dataFiles.Add(dataFileName)
                    If Not dataFiles.Contains(dataFileName) Then
                        dataFiles.Add(dataFileName)
                        cycleAgainToSeeIfThereAreMoreFiles = True
                    End If
                ElseIf selectedLayoutList.Count = 0 Then
                    'dataFiles.Add(dataFileName)
                    If Not dataFiles.Contains(dataFileName) Then
                        dataFiles.Add(dataFileName)
                        cycleAgainToSeeIfThereAreMoreFiles = True
                    End If
                End If
            Next

            ' Wait two minutes to see if other files appear in the directory
            ' This was done in case the service is downloading files while Clarins is generating the files.
            ' At times we miss the rest of the files and they are not picked up until the next pass - 20 minutes
            If cycleAgainToSeeIfThereAreMoreFiles Then
                System.Threading.Thread.Sleep(60000)
                System.Threading.Thread.Sleep(60000)
            End If

        End While

        If dataFiles.Count = 0 Then
            Exit Sub
        End If

        Fill_Records("TATCONVX")
        Fill_Records("TATCONVY")

        Dim sheetNo As Integer = 0
        For Each fileName As String In dataFiles
            If My.Computer.FileSystem.FileExists(dataDirectoryArchive & My.Computer.FileSystem.GetName(fileName)) Then
                Try
                    My.Computer.FileSystem.MoveFile(fileName, dupsDir & My.Computer.FileSystem.GetName(fileName))
                Catch ex As Exception
                End Try

                Me.ErrorMessages.Add("File Exists, placed in Dups Directory: " & dupsDir & My.Computer.FileSystem.GetName(fileName))
                Continue For
            End If
            Load_Workbook(fileName, sheetNo)
        Next

        EnforceConstraints(True)

        Update_Record()

    End Sub

    Public Overrides Sub Update_Create_File()
        MyBase.Update_Create_File()

        ' Global Variable that determine the number of records processed.
        ' Needed so Update_Archive is called from base class.
        R = 0

    End Sub

    Overrides Sub Update_Archive()
        MyBase.Update_Archive()

    End Sub

    Sub Create_Work_Table()

    End Sub

    Overrides Sub Post_Update_Archive()
        MyBase.Post_Update_Archive()

    End Sub

    Sub Load_Workbook(ByVal FILENAME As String, ByRef SHEET_NO As Integer)

        SpreadsheetInfo.SetLicense(ASCMAIN1.GemboxKey) 'v3.1 ("EFYZ-QQSH-LE5Q-NJ7Y")  v3.3 ("EMPX-L9BW-EL8E-4GKJ")  v3.7 (EW1Q-G14I-JKOW-4XS8)

        Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
        Dim fn As String = fi.Name
        Dim ROWCOUNT As Int64 = 0

        Dim TABLE_NAME As String = ""

        If fn.ToUpper.EndsWith(".CSV") Then
            If fn.Contains("_") Then
                TABLE_NAME = Mid(fn, 1, InStr(fn, "_") - 1)
            Else
                Throw New Exception("Invalid File Name - no underscore separating File Name from Date: " & fn)
                Exit Sub
            End If
        Else
            Throw New Exception("Invalid File Type - not CSV: " & fn)
            Exit Sub
        End If

        Dim rowCFG_FILE As DataRow = dst.Tables("TATCONVX").Rows.Find(TABLE_NAME)

        If rowCFG_FILE Is Nothing Then

            If Not My.Computer.FileSystem.FileExists(layoutsDirectory & TABLE_NAME & ".xls") Then
                Exit Sub
            End If

            Dim DDL As String = ""
            Dim DDLKEY As String = ""
            Dim ef2 As ExcelFile = New ExcelFile
            ef2 = ExcelFile.Load(layoutsDirectory & TABLE_NAME & ".xls", LoadOptions.XlsDefault)

            Dim col As Integer = 0
            For Each row As ExcelRow In ef2.Worksheets(0).Rows
                If col > 0 Then
                    If rowCFG_FILE Is Nothing Then
                        rowCFG_FILE = dst.Tables("TATCONVX").NewRow
                        rowCFG_FILE.Item("FILENAME") = TABLE_NAME
                        rowCFG_FILE.Item("FILEEXT") = Mid(fi.Extension, 2)
                        rowCFG_FILE.Item("FILE_AS400") = row.AllocatedCells(0).Value
                        rowCFG_FILE.Item("INIT_DATE") = DATETIME_STAMP
                        dst.Tables("TATCONVX").Rows.Add(rowCFG_FILE)
                    End If

                    Dim rowCFG_LAYOUT As DataRow = dst.Tables("TATCONVY").NewRow
                    For i As Integer = 0 To dst.Tables("TATCONVY").Columns.Count - 1
                        rowCFG_LAYOUT.Item(i) = row.AllocatedCells(i).Value
                    Next

                    Dim COLUMN_NAME As String = rowCFG_LAYOUT.Item("INTERNAL_FIELD_NAME")
                    DDL &= ", " & COLUMN_NAME
                    If CStr(rowCFG_LAYOUT.Item("MORE_INFO") & "").ToUpper = "KEY" Then
                        DDLKEY &= "," & COLUMN_NAME
                    End If
                    Dim COL_DEC As Integer = Val(rowCFG_LAYOUT.Item("DECIMAL_POSITIONS") & "")
                    Dim COL_LEN As Integer = Val(rowCFG_LAYOUT.Item("FIELD_LENGTH_IN_BYTES") & "")
                    Dim COL_DIG As Integer = Val(rowCFG_LAYOUT.Item("NUMBER_OF_DIGITS") & "")
                    If COL_DIG = 0 Then
                        DDL &= " VARCHAR2(" & CStr(COL_LEN) & ")"
                    Else
                        DDL &= " NUMBER(" & CStr(COL_DIG) & "," & CStr(COL_DEC) & ")"
                    End If
                    dst.Tables("TATCONVY").Rows.Add(rowCFG_LAYOUT)
                End If
                col += 1
            Next

            Update_Record_TDA("TATCONVX")
            Update_Record_TDA("TATCONVY")

            ASCMAIN1.sql = "DROP TABLE CONV.CFG_" & TABLE_NAME
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, True)

            DDLKEY = ""

            DDL = "CREATE TABLE CONV.CFG_" & TABLE_NAME & " (" & Mid(DDL, 2) & IIf(DDLKEY <> "", ", PRIMARY KEY (" & Mid(DDLKEY, 2) & ")", "") & ")"
            ASCDATA1.ExecuteSQL(DDL)
        End If

        SHEET_NO += 1

        If dst.Tables.Contains(TABLE_NAME) Then
            'dst.Tables(TABLE_NAME).Rows.Clear()
        Else
            ASCMAIN1.sql = "Select * from CONV.CFG_" & TABLE_NAME
            Create_TDA(dst.Tables.Add(TABLE_NAME), "CONV.CFG_" & TABLE_NAME, "**", 0, True)
        End If

        Dim dt As DataTable = dst.Tables(TABLE_NAME)

        rowCFG_FILE.Item("INIT_DATE") = Now

        Dim use_binding As Boolean = True
        Create_BAs(TABLE_NAME)

        Using sr As New System.IO.StreamReader(FILENAME)
            Do
                Dim lin As String = sr.ReadLine
                If lin Is Nothing OrElse lin.Length = 0 Then
                    Continue Do
                End If

                Dim d() As String = Split(lin, ",")
                For i As Integer = 0 To d.Length - 1
                    If d(i).StartsWith(Chr(34)) Then
                        If d(i).EndsWith(Chr(34)) And d(i).Length > 1 Then
                            d(i) = Trim(Mid(d(i), 2, d(i).Length - 2))
                            If d(i).StartsWith(Chr(34)) Then
                                If d(i).EndsWith(vbTab) Then
                                    d(i) = Mid(d(i), 1, d(i).Length - 1)
                                End If
                                If d(i).EndsWith(Chr(34)) And d(i).Length > 1 Then
                                    i -= 1
                                End If
                            End If
                        Else
                            d(i) &= "," & d(i + 1)
                            d(i + 1) = ""
                            If i + 1 < d.Length - 1 Then
                                For j As Integer = i + 2 To d.Length - 1
                                    d(j - 1) = d(j)
                                    d(j) = ""
                                Next
                            End If
                            i -= 1
                        End If
                    End If
                Next

                Dim r As DataRow = dt.NewRow
                ROWCOUNT += 1

                For i As Integer = 0 To dt.Columns.Count - 1
                    If d.Length - 1 >= i Then
                        If Trim(d(i)) <> "" Then
                            r.Item(i) = Trim(d(i))
                        End If
                    End If
                Next

                If dt.Columns.Contains("IMPORT_FILENAME") AndAlso r.Item("IMPORT_FILENAME") & String.Empty = String.Empty Then
                    r.Item("IMPORT_FILENAME") = fn
                End If

                dt.Rows.Add(r)

                If ROWCOUNT Mod 100000 = 0 Then
                    If use_binding Then
                        Update_BAs(TABLE_NAME)
                    Else
                        Update_Record_TDA(TABLE_NAME)
                    End If
                    dst.Tables(TABLE_NAME).Rows.Clear()
                End If

            Loop While Not sr.EndOfStream
        End Using

        Dim sqldelete As String = String.Empty
        If TABLE_NAME.ToUpper.Contains("OPNORDDTL") OrElse TABLE_NAME.ToUpper.Contains("OPNORDHED") OrElse TABLE_NAME.ToUpper.Contains("ITEMMAST") Then
            sqldelete = "DELETE FROM CONV.CFG_" & TABLE_NAME.ToUpper
        End If

        If dst.Tables(TABLE_NAME).Rows.Count > 0 Then
            If dst.Tables(TABLE_NAME).Rows.Count > 50000 Then
                Update_BAs(TABLE_NAME)
            Else
                Update_Record_TDA(TABLE_NAME, sqldelete)
            End If
        End If

        dst.Tables(TABLE_NAME).Rows.Clear()

        rowCFG_FILE.Item("ROWCOUNT") = ROWCOUNT
        rowCFG_FILE.Item("LAST_DATE") = Now

        Dim subDir As String = DateTime.Now.ToString("yyyyMM")
        ' If the Archive direcory does not exist then create it as a courtesy
        If Not My.Computer.FileSystem.DirectoryExists(fi.Directory.ToString & "\archive\" & subDir) Then
            My.Computer.FileSystem.CreateDirectory(fi.Directory.ToString & "\archive\" & subDir)
        End If

        ' As per Walter on 2/2/2016 have code crash if the file already exists
        My.Computer.FileSystem.MoveFile(FILENAME, fi.Directory.ToString & "\archive\" & subDir & "\" & fi.Name)

    End Sub

    Private Sub DisposeOPD()
        Try
            If clsASCBASE1 Is Nothing Then
                Exit Sub
            End If

            With clsASCBASE1

                If .CMDs IsNot Nothing AndAlso .CMDs.Count <> 0 Then
                    For Each CMD_key As String In .CMDs.Keys
                        Dim cmd As OracleCommand = .CMDs(CMD_key)
                        For Each param As OracleParameter In cmd.Parameters
                            param.Dispose()
                        Next
                        cmd.Dispose()
                    Next
                End If
                .CMDs = Nothing

                If .BA_CMDs IsNot Nothing AndAlso .BA_CMDs.Count <> 0 Then
                    For Each CMD_key As String In .BA_CMDs.Keys
                        Dim cmds() As OracleCommand = .BA_CMDs(CMD_key)
                        For Each cmd As OracleCommand In cmds
                            For Each param As OracleParameter In cmd.Parameters
                                param.Dispose()
                            Next
                            cmd.Dispose()
                        Next
                        cmds = Nothing
                    Next
                End If
                .BA_CMDs = Nothing

                If .TDAs IsNot Nothing Then
                    For Each tda As OracleDataAdapter In .TDAs.Values
                        tda.Dispose()
                    Next
                End If
                .TDAs = Nothing

                .Dispose()
            End With

        Catch ex As Exception

        End Try
    End Sub


End Class