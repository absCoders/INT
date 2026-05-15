Public MustInherit Class WHCRPTS1 'Base Class for ABSolution Reports
    Inherits ASCBASE0

    Public R As Int64 = 0
    Protected clsErrorMessage As List(Of String)

    Public tblASTSPRF1 As DataTable

    MustOverride Property reportName As String
    Public Property reportFilename As String
    Protected Property CR_params As New Dictionary(Of String, String)
    Protected Property SUBT As String

    Protected inTestMode As Boolean = True

    Public logsFolder As String = String.Empty
    Private logFilename As String = String.Empty
    Private logStreamWriter As System.IO.StreamWriter

    Sub New()
        MyBase.New()
    End Sub

    Sub New(g As ABSEnvironment)
        MyBase.New(g)

        clsErrorMessage = New List(Of String)
        If inTestMode = True Then RecordLogEntry("Me.MENU_ITEM_TYPE = 'C'")

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRPTS1"

        With dst
            With .Tables.Add("TATEVNT1").Columns
                .Add("EVENT_LOG")
            End With
        End With

        tbl = dst.Tables("TATEVNT1") ' Log Events
    End Sub

 
    Sub Update_Record()

        BeginTrans()

        Update_Create_File() ' THIS IS THE PROC IN YOUR CLASS THAT WOULD CONTAIN YOUR REAL UPDATE

        ' LOG THAT THE UPDATE HAS OCCURED

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

    Public ReadOnly Property ErrorMessages As List(Of String)
        Get
            Return clsErrorMessage
        End Get

        'Set(value As List(Of String))
        '    clsErrorMessage = value
        'End Set
    End Property

    Public Function GenerateReport(ByVal clientData As Dictionary(Of String, Object)) As String

        Dim ExportFormat As String = "RPT"

        If clientData.ContainsKey("ExportFormat") Then
            ExportFormat = clientData("ExportFormat")
        End If

        Dim RPT_FILENAME
        If ASCMAIN1.Running_in_VS Then
            Dim XSD_FILENAME As String = ASCMAIN1.Folders("Temp") & Me.G.FormName & ".XSD"
            If Not My.Computer.FileSystem.FileExists(XSD_FILENAME) Then
                dst.WriteXml(XSD_FILENAME, XmlWriteMode.WriteSchema)
            End If
            RPT_FILENAME = ASCMAIN1.Folders("Reports") & reportName & ".RPT"
        Else

            If ASCMAIN1.DBS_SERVER = "VAN" Then ' necessary because V1 uses G: drive in ASTPARM1
                RPT_FILENAME = "R:\VDI\REPORTS\" & reportName & ".RPT"
            ElseIf ASCMAIN1.DBS_SERVER = "ANE" Then ' necessary because V1 uses G: drive in ASTPARM1
                RPT_FILENAME = "G:\EXP\REPORTS\" & reportName & ".RPT"
            Else
                RPT_FILENAME = ASCMAIN1.rowASTPARM1.Item("AS_PARM_REPORTS_DIR") & "\" & reportName & ".RPT"
            End If
        End If

        ' Override Value 
        If Not String.IsNullOrEmpty(G.ReportsDirectory) Then
            RPT_FILENAME = IO.Path.Combine(G.ReportsDirectory, reportName & ".RPT")
        End If

        'Job Processing Service
        'Report Response Error: Generate Report Error (CR_RPT.Load(RPT_FILENAME)): Load report failed.
        'Generate Report Error (CR_RPT.Load(RPT_FILENAME)): Load report failed.
        'Class did not return a Report

        If inTestMode Then RecordLogEntry("RPT_FILENAME = " & RPT_FILENAME)

        Dim CR_RPT As CrystalDecisions.CrystalReports.Engine.ReportDocument = New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim CR_SubRpt As CrystalDecisions.CrystalReports.Engine.ReportDocument = New CrystalDecisions.CrystalReports.Engine.ReportDocument

        Try
            If inTestMode Then RecordLogEntry("CR_RPT.Load(RPT_FILENAME) = " & RPT_FILENAME)
            CR_RPT.Load(RPT_FILENAME)
        Catch ex As Exception
            clsErrorMessage.Add("Generate Report Error (CR_RPT.Load(RPT_FILENAME)): " & ex.Message & " Inner Exception: " & ex.InnerException.Message)
            Return ""
        End Try

        Try
            If inTestMode Then RecordLogEntry("CR_RPT.SetDataSource(dst)")
            If inTestMode Then RecordLogEntry("dst.Tables.Count = " & dst.Tables.Count)
            CR_RPT.SetDataSource(dst)
        Catch ex As Exception
            clsErrorMessage.Add("Generate Report Error (CR_RPT.SetDataSource(dst)): " & ex.Message & " Inner Exception: " & ex.InnerException.Message)
            Return ""
        End Try

        If inTestMode Then RecordLogEntry("CR_RPT.SetDataSource(dst) Completed")

        For Each sr As CrystalDecisions.CrystalReports.Engine.ReportDocument In CR_RPT.Subreports
            Try
                If inTestMode Then RecordLogEntry("sr.SetDataSource(dst)")
                sr.SetDataSource(dst)
            Catch ex As Exception
                clsErrorMessage.Add("Generate Report Error (sr.SetDataSource(dst)): " & ex.Message & " Inner Exception: " & ex.InnerException.Message)
                Return ""
            End Try
        Next

        If inTestMode Then RecordLogEntry("sr.SetDataSource(dst) Completed")

        Dim REPORT_NO As String = ASCMAIN1.Next_Control_No("ASTSPRF1.REPORT_NO")
        Dim filename As String = ASCMAIN1.DBS_COMPANY & "_" & REPORT_NO & "." & ExportFormat
        Dim DestOpt As New CrystalDecisions.Shared.DiskFileDestinationOptions

        If inTestMode Then RecordLogEntry("REPORT_NO = " & REPORT_NO)


        Dim rptDirectory As String
        If clientData.ContainsKey("directory") Then
            rptDirectory = clientData("directory")
        Else
            rptDirectory = If(G.TempDirectory, ASCMAIN1.Folders("Temp"))
        End If
        DestOpt.DiskFileName = IO.Path.Combine(rptDirectory, filename)

        If inTestMode Then RecordLogEntry("With CR_RPT.ExportOptions")
        With CR_RPT.ExportOptions
            .DestinationOptions = DestOpt
            .ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
            Select Case ExportFormat
                Case "RPT"
                    .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.CrystalReport
                Case "PDF"
                    .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                Case "HTM"
                    .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.HTML40
                Case "RTF"
                    .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.EditableRTF
                Case Else
                    .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.CrystalReport
            End Select
        End With

        Try
            SetParameterValue("USERID", ASCMAIN1.USER_ID, CR_RPT, CR_SubRpt)
            SetParameterValue("UID", ASCMAIN1.DBS_COMPANY, CR_RPT, CR_SubRpt)
            SetParameterValue("YPD", ASCMAIN1.Get_Legend(ASCMAIN1.CYP), CR_RPT, CR_SubRpt)
            SetParameterValue("INSTNAME", ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_NAME"), CR_RPT, CR_SubRpt)
            SetParameterValue("SESSIONID", REPORT_NO, CR_RPT, CR_SubRpt)
            SetParameterValue("RPT", reportName, CR_RPT, CR_SubRpt)
            SetParameterValue("XNO", G.XNO, CR_RPT, CR_SubRpt)
            SetParameterValue("VERSIONNO", ASCMAIN1.VERSION_NO & "", CR_RPT, CR_SubRpt)
            SetParameterValue("RPT_TITLE", CR_RPT.SummaryInfo.ReportTitle & "", CR_RPT, CR_SubRpt)
            SetParameterValue("SUBT", SUBT, CR_RPT, CR_SubRpt)

            For Each k As String In CR_params.Keys
                If inTestMode Then RecordLogEntry("SetParameterValue k = " & k)
                SetParameterValue(k, CR_params.Item(k), CR_RPT, CR_SubRpt)
            Next

        Catch ex As Exception
            ErrorMessages.Add("Error Setting Report Parameters" & vbCr & ex.Message)
            Return ""
        End Try

        Try
            If inTestMode Then RecordLogEntry("CR_RPT.Export()")
            CR_RPT.Export()
        Catch ex As Exception
            clsErrorMessage.Add("Generate Report Error (CR_RPT.Export()): " & ex.Message)
        End Try

        If ExportFormat = "RPT" Then
            If tblASTSPRF1 Is Nothing Then
                ASCMAIN1.sql = "Select * from ASTSPRF1 where ROWNUM <1"
                tblASTSPRF1 = ASCDATA1.GetDataTable
            End If

            Dim rowASTSPRF1 As DataRow = tblASTSPRF1.NewRow
            With rowASTSPRF1
                .Item("REPORT_NO") = REPORT_NO
                .Item("FORM_NAME") = Me.G.FormName ' FORM_NAME
                .Item("XNO") = Me.G.XNO
                .Item("USER_ID") = ASCMAIN1.USER_ID
                .Item("YYYYPP") = ASCMAIN1.CYP
                .Item("YP_LEGEND") = ASCMAIN1.Get_Legend(ASCMAIN1.CYP)
                .Item("RPT_TITLE") = "" 'ASCMAIN1.CR_RPT.SummaryInfo.ReportTitle
                .Item("RPT") = reportName
                .Item("COMPUTER_NAME") = ASCMAIN1.COMPUTER_NAME
                .Item("REPORT_DATE") = Now + ASCMAIN1.NowTSD
                .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                .Item("MENU_ITEM_OBJECT") = "" 'frmASFBASE0.MENU_ITEM_OBJECT
                .Item("MENU_ITEM_TYPE") = "" 'frmASFBASE0.MENU_ITEM_TYPE
                .Item("MENU_ID") = "" 'frmASFBASE0.MENU_ID
                .Item("MENU_ITEM_SECURITY") = "" 'frmASFBASE0.MENU_ITEM_SECURITY
                .Item("VERSION_NO") = ASCMAIN1.VERSION_NO
            End With

            tblASTSPRF1.Rows.Add(rowASTSPRF1)

            Dim r As DataRow = tblASTSPRF1.NewRow
            r.ItemArray = rowASTSPRF1.ItemArray

        End If

        If inTestMode Then RecordLogEntry("Return filename: " & filename)

        ' Dispose Objects
        Try
            If CR_SubRpt IsNot Nothing Then
                CR_SubRpt.Close()
                CR_SubRpt.Dispose()
                CR_SubRpt = Nothing
            End If

            If CR_RPT IsNot Nothing Then
                CR_RPT.Close()
                CR_RPT.Dispose()
                CR_RPT = Nothing
            End If
        Catch ex As Exception

        End Try

        Return filename

    End Function

    Sub SetParameterValue( _
    ByVal pfName As String, _
    ByVal pfValue As String, _
    ByVal CR_RPT As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
    ByVal CR_SubRpt As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
    Optional ByVal Sub_Report As Boolean = False)

        Dim Par As CrystalDecisions.Shared.ParameterValues = Nothing
        Dim ParD As New CrystalDecisions.Shared.ParameterDiscreteValue()

        Try
            If Sub_Report Then
                Par = CR_SubRpt.DataDefinition.ParameterFields.Item(pfName).CurrentValues
            Else
                Par = CR_RPT.DataDefinition.ParameterFields.Item(pfName).CurrentValues
            End If
        Catch ex As Exception
            ErrorMessages.Add("Error Adding Parameter " & pfName & " to Report " & CR_RPT.Name)
        End Try

        ParD.Value = pfValue
        Par.Add(ParD)

        Try
            If Sub_Report Then
                CR_SubRpt.DataDefinition.ParameterFields.Item(pfName).ApplyCurrentValues(Par)
            Else
                CR_RPT.DataDefinition.ParameterFields.Item(pfName).ApplyCurrentValues(Par)
            End If
        Catch ex As Exception
            ErrorMessages.Add("Error Adding Parameter " & pfName & " to Report " & CR_RPT.Name)
        End Try
    End Sub


#Region "Log Procedures"

    Private Function OpenLogFile() As Boolean

        Try

            If logsFolder.Length = 0 Then
                Exit Function
            End If

            If Not My.Computer.FileSystem.DirectoryExists(logsFolder) Then
                My.Computer.FileSystem.CreateDirectory(logsFolder)
            End If

            logFilename = Format(Now, "yyyyMMdd") & ".log"
            If logStreamWriter IsNot Nothing Then
                logStreamWriter.Close()
                logStreamWriter.Dispose()
            End If

            Dim logdirectory As String = logsFolder
            If Not logdirectory.EndsWith("\") Then logdirectory &= "\"
            logdirectory &= "Logs\"

            If Not My.Computer.FileSystem.DirectoryExists(logdirectory) Then
                My.Computer.FileSystem.CreateDirectory(logdirectory)
            End If

            logStreamWriter = New System.IO.StreamWriter(logdirectory & logFilename, True)

            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Sub RecordLogEntry(ByVal message As String)
        Try
            If logsFolder.Length = 0 Then
                Exit Sub
            End If

            OpenLogFile()
            logStreamWriter.WriteLine(DateTime.Now & ": " & message)
            CloseLog()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CloseLog()
        Try
            If logStreamWriter IsNot Nothing Then
                logStreamWriter.Close()
                logStreamWriter.Dispose()
                logStreamWriter = Nothing
            End If
        Catch ex As Exception

        Finally

        End Try
    End Sub

#End Region

#Region "Report Functions"

    Public Function SQLA( _
        ByVal PB_COLUMN_NAME As String, tblASTDSQLA As DataTable, _
        Optional ByVal COLUMN_NAME As String = "CODE_VALUES", _
        Optional ByVal SQL_List As Boolean = False) As String

        If tblASTDSQLA Is Nothing Then
            Return String.Empty
        End If

        Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find(PB_COLUMN_NAME)
        If rowASTDSQLA Is Nothing Then
            Return String.Empty
        End If

        SQLA = rowASTDSQLA.Item(COLUMN_NAME) & ""
        If SQL_List And SQLA <> "" Then
            SQLA = "'" & Replace(SQLA, ",", "','") & "'"
        End If

        Return SQLA

    End Function
#End Region

End Class