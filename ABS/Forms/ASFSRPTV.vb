Imports CrystalDecisions.CrystalReports

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports OutlookApp = Microsoft.Office.Interop.Outlook

Public Class ASFSRPTV

    Public tblASTSPRF1 As DataTable
    Public CRs As New Dictionary(Of String, CrystalDecisions.CrystalReports.Engine.ReportDocument)
    Public PRINTER_NAME As String
    Public REPORT_FILENAMES As New Dictionary(Of String, String)
    Public REPORT_ARCHIVES As New Dictionary(Of String, DataRow)
    Public source_dst As DataSet = Nothing
    Public publish_documents As Boolean = False
    Public FILENAME_to_export As String
    Public ExportFormat As String = "XLS"
    Dim FORM_NAME As String = ""
    Dim SET_ID As String = ""

    Public frmASFBASE0 As ASFBASE0

    Dim ok_to_archive_report As String = ""
    Dim USER_GROUP_IDs As New List(Of String)

    Public Sub New( _
    Optional ByVal incoming_dst As DataSet = Nothing, _
    Optional ByVal incoming_frm As ASFBASE0 = Nothing)

        If incoming_frm Is Nothing Then
            frmASFBASE0 = ASCMAIN1.ActiveForm
        Else
            frmASFBASE0 = incoming_frm
        End If

        If incoming_dst Is Nothing Then
            If frmASFBASE0 IsNot Nothing Then
                source_dst = frmASFBASE0.dst
            End If
        Else
            source_dst = incoming_dst
        End If

        If frmASFBASE0 IsNot Nothing Then
            FILENAMEs_to_Publish = frmASFBASE0.FILENAMEs_to_Publish
        End If

        InitializeComponent()
    End Sub

    Sub Setup_dst()
        With dst

            ASCMAIN1.sql = "Select ASTROPT4.*, ASTUSER1.USER_NAME, ASTUSER1.USER_EMAIL, ASTUSER1.USER_STATUS" _
            & " from ASTROPT4,ASTUSER1 where ASTUSER1.USER_ID = ASTROPT4.USER_ID" _
            & " and ASTROPT4.FORM_NAME = :PARM1 and ASTROPT4.SET_ID = :PARM2"
            Create_TDA(.Tables.Add, "ASTROPT4", "**", , , "VV")

            With .Tables.Add("ASTROPTA")
                .Columns.Add("USER_GROUP_ID")
                .Columns.Add("USER_ID")
                .Columns.Add("USER_NAME")
                .Columns.Add("USER_EMAIL")
                .PrimaryKey = New DataColumn() {.Columns("USER_GROUP_ID"), .Columns("USER_ID")}
            End With

            .Relations.Add("ASTROPT4_ASTROPTA", _
                           New DataColumn() {.Tables("ASTROPT4").Columns("USER_ID")}, _
                           New DataColumn() {.Tables("ASTROPTA").Columns("USER_GROUP_ID")})

            Create_TDA(.Tables.Add, "ASTWRPT0", "*")
            
            If publish_documents Then
                dst.Tables("ASTWRPT0").Columns.Add("FILENAME")
                Create_TDA(.Tables.Add, "ASTLIST1", "*")
                Create_TDA(.Tables.Add, "ASTLIST2", "*")
            End If

            Create_TDA(.Tables.Add, "ASTSPRF1", "*")

            Create_TDA(.Tables.Add, "ASTATTA1", "*", 0, False)
            Fill_Records("ASTATTA1")

            If tblASTSPRF1 IsNot Nothing Then
                For Each rowASTSPRF1 As DataRow In tblASTSPRF1.Rows
                    Dim row As DataRow = dst.Tables("ASTSPRF1").NewRow
                    row.ItemArray = rowASTSPRF1.ItemArray
                    dst.Tables("ASTSPRF1").Rows.Add(row)
                Next
            End If
            dst.Tables("ASTSPRF1").Columns.Add("SELECT", GetType(System.String))
            dst.Tables("ASTSPRF1").Columns("SELECT").ReadOnly = False

            tblASTSPRF1 = dst.Tables("ASTSPRF1")
        End With

        If publish_documents AndAlso FILENAMEs_to_Publish.Count <> 0 Then
            For Each FILENAME As String In FILENAMEs_to_Publish
                Add_File(FILENAME)
            Next
        End If
    End Sub

    Private Sub ASFSRPTV_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Width = Screen.PrimaryScreen.WorkingArea.Width * 0.85
        Me.Height = Screen.PrimaryScreen.WorkingArea.Height * 0.85
        ASCMAIN1.Center(Me)

        Setup_dst()

        grdASTSPRF1.DataSource = dst.Tables("ASTSPRF1")

        For Each row As DataRow In dst.Tables("ASTSPRF1").Rows
            Make_Tab(row.Item( _
                "RPT_TITLE") & "", _
                ASCMAIN1.DBS_COMPANY & "_" & row.Item("REPORT_NO") & ".RPT", _
                row.Item("REPORT_NO"))
            FORM_NAME = row.Item("FORM_NAME") & String.Empty
            SET_ID = row.Item("XNO") & ""
            row.Item("SELECT") = "1"
        Next

        'grdASTROPTA.DataMember = "ASTROPTA"
        'grdASTROPTA.DataSource = dst
        grdASTROPT4.DataSource = dst.Tables("ASTROPT4")
        grdASTROPT4.DisplayLayout.Override.ExpansionIndicator = UltraWinGrid.ShowExpansionIndicator.CheckOnDisplay

        If Me.publish_documents Then
            grdASTWRPT0.DataSource = dst.Tables("ASTWRPT0")
            grdASTWRPT0.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            grdASTWRPT0.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            For Each gcol As UltraWinGrid.UltraGridColumn In grdASTWRPT0.DisplayLayout.Bands(0).Columns
                If gcol.Key = "RPT_TITLE" Or gcol.Key = "SET_DESC" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Color.WhiteSmoke
                End If
            Next

            txtSubject.Text = "Web Links to Documents"
            txtSubject.Visible = True
            lblSubject.Visible = True

        End If
        
        'TEMP
        If ASCMAIN1.Running_in_VS Then
        Else
            'cmdPublish.Enabled = False
        End If

        If publish_documents Then
            UltraExplorerBar1.Groups("Printing Options").Visible = False
            UltraExplorerBar1.Groups("Publish").Visible = False
            UltraExplorerBar1.Groups("Publish to Portal").Visible = True
            splRD.Panel1Collapsed = True
            'grdASTSPRF1.Visible = False
            'grdASTWRPT0.Visible = True
            Me.Text = "Publish Documents to ABSolution Web Portal"

            chkEmailLinks.Visible = True
        Else
            UltraExplorerBar1.Groups("Publish to Portal").Visible = False
            splRD.Panel2Collapsed = True
            'grdASTSPRF1.Visible = True
            'grdASTWRPT0.Visible = False
        End If

        lblSubject.Appearance.ForeColor = Color.White
        lblSubject.Appearance.BackColor = Color.FromArgb(98, 160, 232)
        lblSubject.Appearance.BackColor2 = Color.FromArgb(83, 115, 191)
        lblSubject.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical

        chkEmailLinks.Appearance.ForeColor = Color.White
        chkEmailLinks.Appearance.BackColor = Color.FromArgb(98, 160, 232)
        chkEmailLinks.Appearance.BackColor2 = Color.FromArgb(83, 115, 191)
        chkEmailLinks.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical



    End Sub

    Private Sub ASFSRPTV_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        If publish_documents Then Exit Sub

        For Each utp As Infragistics.Win.UltraWinTabControl.UltraTabPageControl In Me.UltraTabControl1.Controls

            For Each ctl As Control In utp.Controls
                If ctl.GetType.ToString = "CrystalDecisions.Windows.Forms.CrystalReportViewer" Then
                    Dim CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer = DirectCast(ctl, CrystalDecisions.Windows.Forms.CrystalReportViewer)
                    Try
                        CRV.Dispose()
                    Catch ex As Exception

                    End Try

                End If
            Next
        Next

        For Each cr As CrystalDecisions.CrystalReports.Engine.ReportDocument In CRs.Values
            cr.Dispose()
        Next

        If ASCMAIN1.CR_SubRpt IsNot Nothing Then
            ASCMAIN1.CR_SubRpt.Close()
            ASCMAIN1.CR_SubRpt.Dispose()
            ASCMAIN1.CR_SubRpt = Nothing
        End If

        If ASCMAIN1.CR_RPT IsNot Nothing Then
            ASCMAIN1.CR_RPT.Close()
            ASCMAIN1.CR_RPT.Dispose()
            ASCMAIN1.CR_RPT = Nothing
        End If

    End Sub

    Private Sub ASFSRPTV_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

    End Sub

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdASTSPRF1, "SBB", "CardView", "Select All", "De-Select All")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
        End If
        If tlb_pop.Tools.Exists("CardView") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("CardView"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = grd.DisplayLayout.Bands(0).CardView
        End If

        If grd.Name = "grdASTSPRF1" _
        Or grd.Name = "grdASTROPT4" Then
            Exit Sub
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                Case ""
                    If grd.DisplayLayout.Override.AllowUpdate <> DefaultableBoolean.True Then
                        e.Cancel = True
                    End If
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "CardView"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).CardView = tlb_sbt.Checked

            Case "Select All"
                For Each rowASTROPT1 As DataRow In dst.Tables("ASTSPRF1").Select
                    rowASTROPT1.Item("SELECT") = "1"
                Next

            Case "De-Select All"
                For Each rowASTROPT1 As DataRow In dst.Tables("ASTSPRF1").Select
                    rowASTROPT1.Item("SELECT") = "0"
                Next

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Customer Inquiry"
            '    Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
            '    Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")
        End Select
    End Sub

#End Region

    Public Sub Make_Tab( _
    ByVal tabtext As String, _
    ByVal FILENAME As String, _
    ByVal REPORT_NO As String)

        Dim utp As New Infragistics.Win.UltraWinTabControl.UltraTabPageControl
        Dim crv As New CrystalDecisions.Windows.Forms.CrystalReportViewer

        ' Add a Tab Page Control to the Tab Control
        Dim i As Integer = Me.UltraTabControl1.Tabs.Count
        Me.UltraTabControl1.Controls.Add(utp)
        Me.UltraTabControl1.Tabs(i).Text = tabtext
        Me.UltraTabControl1.Tabs(i).Key = REPORT_NO

        ' Add a Crystal Report Viewer to the Tab Page Control & Configure it
        utp.Controls.Add(crv)
        crv.ActiveViewIndex = -1
        crv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        crv.Dock = System.Windows.Forms.DockStyle.Fill
        crv.BackColor = System.Drawing.Color.FromArgb(222, 223, 206)

        Dim RPT As CrystalDecisions.CrystalReports.Engine.ReportDocument
        If Not CRs.ContainsKey(REPORT_NO) Then
            Dim REPORT_FILENAME As String = ""
            'Load the Report into a ReportDocument, looking at Temp first, then at Archive
            If Not CRs.ContainsKey(REPORT_NO) Then
                If My.Computer.FileSystem.FileExists(ASCMAIN1.Folders("Temp") & FILENAME) Then
                    REPORT_FILENAME = ASCMAIN1.Folders("Temp") & FILENAME
                Else
                    REPORT_FILENAME = ASCMAIN1.Folders("Archive") & "Reports\" & Mid(FILENAME, 1, 3) & "\" & Mid(FILENAME, 5, 5) & "\" & FILENAME
                End If
            Else
                REPORT_FILENAME = ASCMAIN1.Folders("Temp") & FILENAME
            End If

            If My.Computer.FileSystem.FileExists(REPORT_FILENAME) Then
                RPT = New CrystalDecisions.CrystalReports.Engine.ReportDocument
                Try
                    RPT.Load(REPORT_FILENAME)
                    crv.ReportSource = RPT
                Catch ex As Exception
                    MsgBox("Problem Report: " & REPORT_FILENAME & vbCr & vbCr & ex.Message, MsgBoxStyle.OkOnly, "Cannot Load Report " & REPORT_NO)
                    RPT.Dispose()
                    Exit Sub
                End Try
            Else
                MsgBox("Problem Report: " & REPORT_FILENAME & vbCr & vbCr & "Report file does not exist in Archive", MsgBoxStyle.OkOnly, "Cannot Load Report " & REPORT_NO)
                Exit Sub
            End If

            Add_Report(REPORT_NO, RPT)
        Else
            RPT = CRs(REPORT_NO)
            crv.ReportSource = RPT
        End If

        Me.UltraTabControl1.SelectedTab = UltraTabControl1.Tabs(0)
        Me.UltraTabControl1.TabIndex = 0
    End Sub

    Public Shadows Function Generate_Report( _
    ByVal RPT As String, _
    Optional ByVal RPT_TITLE As String = "", _
    Optional ByVal SUBT As String = "", _
    Optional ByVal Show_Report As Boolean = False, _
    Optional ByVal PB_Report As Boolean = False, _
    Optional ByVal RecordSelectionFormula As String = "", _
    Optional ByVal ExportFormat As String = "", _
    Optional ByVal TempExportFilenameBody As String = "", _
    Optional ByVal archive_this_report As Boolean = True) As String

        If ExportFormat = "" Then
            ExportFormat = "RPT"
        End If

        Dim RPT_FILENAME As String = ""
        If ASCMAIN1.Running_in_VS Then
            Dim XSD_FILENAME As String = ASCMAIN1.Folders("Temp") & frmASFBASE0.Name & ".XSD"
            If Not My.Computer.FileSystem.FileExists(XSD_FILENAME) Then
                source_dst.WriteXml(XSD_FILENAME, XmlWriteMode.WriteSchema)
            End If
            RPT_FILENAME = ASCMAIN1.Folders("Reports") & RPT & ".RPT"
        Else

            If ASCMAIN1.DBS_SERVER = "VAN" Then ' necessary because V1 uses G: drive in ASTPARM1
                RPT_FILENAME = "R:\VDI\REPORTS\" & RPT & ".RPT"
                'If ASCMAIN1.USER_ID = "wjz" Then
                '    MsgBox("Current: " & RPT_FILENAME)
                '    MsgBox(ASCMAIN1.Folders("SharedRoot") & RPT & ".RPT")
                'End If
            ElseIf ASCMAIN1.DBS_SERVER = "ANE" Then ' necessary because V1 uses G: drive in ASTPARM1
                'RPT_FILENAME = "G:\AHA\REPORTS\" & RPT & ".RPT"
                RPT_FILENAME = "G:\EXP\REPORTS\" & RPT & ".RPT"
            Else
                RPT_FILENAME = ASCMAIN1.rowASTPARM1.Item("AS_PARM_REPORTS_DIR") & "\" & RPT & ".RPT"
            End If
        End If

        ASCMAIN1.CR_RPT = New CrystalDecisions.CrystalReports.Engine.ReportDocument
        ASCMAIN1.CR_SubRpt = New CrystalDecisions.CrystalReports.Engine.ReportDocument

        Try
            ASCMAIN1.CR_RPT.Load(RPT_FILENAME)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error Trying to Load Report: " & RPT_FILENAME)
            Return ""
            Exit Function
        End Try

        If RPT_TITLE <> "" Then
            ASCMAIN1.CR_RPT.SummaryInfo.ReportTitle = RPT_TITLE
        Else
            Try
                ASCMAIN1.CR_RPT.SummaryInfo.ReportTitle = frmASFBASE0.MENU_ITEM_DESC
            Catch ex As Exception

            End Try
        End If

        Try
            ASCMAIN1.CR_RPT.SetDataSource(source_dst)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        For Each sr As Engine.ReportDocument In ASCMAIN1.CR_RPT.Subreports
            Try
                sr.SetDataSource(source_dst)

            Catch ex As Exception
                If ASCMAIN1.Running_in_VS Then
                    Stop
                End If
            End Try
        Next

        'ASCMAIN1.CR_RPT.DataDefinition.SortFields.Item
        'Dim myDatabaseFieldDefinition As CrystalDecisions.CrystalReports.Engine.DatabaseFieldDefinition
        'myDatabaseFieldDefinition = ASCMAIN1.CR_RPT.Database.Tables.Item(tableName).Fields.Item(fieldName)
        'Dim mySortField As CrystalDecisions.CrystalReports.Engine.SortField = ASCMAIN1.CR_RPT.DataDefinition.SortFields.Item(1)
        'If (mySortField.SortType = CrystalDecisions.Shared.SortFieldType.RecordSortField) Then
        '    mySortField.Field = myDatabaseFieldDefinition
        '    mySortField.SortDirection = CrystalDecisions.Shared.SortDirection.AscendingOrder
        'End If


        Dim REPORT_NO As String = ASCMAIN1.Next_Control_No("ASTSPRF1.REPORT_NO")
        Dim filename As String = ASCMAIN1.DBS_COMPANY & "_" & REPORT_NO & "." & ExportFormat
        Dim DestOpt As New CrystalDecisions.Shared.DiskFileDestinationOptions
        DestOpt.DiskFileName = ASCMAIN1.Folders("Temp") & filename

        REPORT_FILENAMES.Add(REPORT_NO, DestOpt.DiskFileName)

        With ASCMAIN1.CR_RPT.ExportOptions
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

        '        ASCMAIN1.CR_RPT.Database.Verify()
        Try
            SetParameterValue("USERID", ASCMAIN1.USER_ID)
            SetParameterValue("UID", ASCMAIN1.DBS_COMPANY)
            SetParameterValue("YPD", ASCMAIN1.Get_Legend(ASCMAIN1.CYP))
            SetParameterValue("INSTNAME", ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_NAME"))
            'SetParameterValue("SESSIONID", ASCMAIN1.SESSION_NO)
            SetParameterValue("SESSIONID", REPORT_NO)
            SetParameterValue("RPT", RPT)
            SetParameterValue("XNO", frmASFBASE0.XNO)
            SetParameterValue("VERSIONNO", ASCMAIN1.VERSION_NO & "")
            SetParameterValue("RPT_TITLE", ASCMAIN1.CR_RPT.SummaryInfo.ReportTitle & "")
            SetParameterValue("SUBT", SUBT & "")

            For Each k As String In frmASFBASE0.CR_params.Keys
                SetParameterValue(k, frmASFBASE0.CR_params.Item(k))
            Next
            frmASFBASE0.CR_params.Clear()

            If RecordSelectionFormula <> "" Then
                ASCMAIN1.CR_RPT.RecordSelectionFormula = RecordSelectionFormula
            End If

        Catch ex As Exception
            MsgBox("Error Setting Report Parameters" & vbCr & ex.Message, MsgBoxStyle.OkOnly, "Report Design Issue")
            Return ""
        End Try

        Call ASCMAIN1.Progress("Exporting " & RPT & " (" & filename & ")")

        Try
            ASCMAIN1.CR_RPT.Export()
            If ASCMAIN1.Running_in_VS Then
                If ASCMAIN1.DBS_SERVER = "" Then
                    ok_to_archive_report = "N"
                End If
                If ok_to_archive_report = "" And archive_this_report AndAlso ASCMAIN1.JOB_STREAM_CODE.Length = 0 Then
                    If MsgBoxResult.No = MsgBox("Copy File to Archive?", MsgBoxStyle.YesNo, "You are running a Report on a Development Machine") Then
                        ok_to_archive_report = "N"
                    Else
                        ok_to_archive_report = "Y"
                    End If
                End If
            Else
                ok_to_archive_report = "Y"
            End If
            If ok_to_archive_report = "Y" And archive_this_report Then
                My.Computer.FileSystem.CopyFile( _
                    ASCMAIN1.Folders("Temp") & filename, _
                    ASCMAIN1.Folders("Archive") & "Reports\" & Mid(filename, 1, 3) & "\" & Mid(filename, 5, 5) & "\" & filename)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error Exporting Report File")

        End Try

        Call Add_Report(REPORT_NO, ASCMAIN1.CR_RPT)

        If archive_this_report Then
            Call ASCMAIN1.Progress("Writing to Archive")
        End If

        If ExportFormat = "RPT" Then
            If tblASTSPRF1 Is Nothing Then
                ASCMAIN1.sql = "Select * from ASTSPRF1 where ROWNUM <1"
                tblASTSPRF1 = ASCDATA1.GetDataTable
            End If

            Dim rowASTSPRF1 As DataRow = tblASTSPRF1.NewRow
            With rowASTSPRF1
                .Item("REPORT_NO") = REPORT_NO
                .Item("FORM_NAME") = frmASFBASE0.MENU_ITEM_OBJECT ' FORM_NAME
                .Item("XNO") = frmASFBASE0.XNO
                .Item("USER_ID") = ASCMAIN1.USER_ID
                .Item("YYYYPP") = ASCMAIN1.CYP
                .Item("YP_LEGEND") = ASCMAIN1.Get_Legend(ASCMAIN1.CYP)
                .Item("RPT_TITLE") = ASCMAIN1.CR_RPT.SummaryInfo.ReportTitle
                .Item("RPT") = RPT
                .Item("COMPUTER_NAME") = ASCMAIN1.COMPUTER_NAME
                .Item("REPORT_DATE") = Now + ASCMAIN1.NowTSD
                .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                .Item("MENU_ITEM_OBJECT") = frmASFBASE0.MENU_ITEM_OBJECT
                .Item("MENU_ITEM_TYPE") = frmASFBASE0.MENU_ITEM_TYPE
                .Item("MENU_ID") = frmASFBASE0.MENU_ID
                .Item("MENU_ITEM_SECURITY") = frmASFBASE0.MENU_ITEM_SECURITY
                .Item("VERSION_NO") = ASCMAIN1.VERSION_NO
            End With

            tblASTSPRF1.Rows.Add(rowASTSPRF1)

            If Not archive_this_report Then
                rowASTSPRF1.AcceptChanges()
            End If

            Dim r As DataRow = tblASTSPRF1.NewRow
            r.ItemArray = rowASTSPRF1.ItemArray

            REPORT_ARCHIVES.Add(REPORT_NO, r)
        End If

        If TempExportFilenameBody <> "" Then
            Dim TempExportFilename As String = _
                TempExportFilenameBody & "." & ExportFormat
            If My.Computer.FileSystem.FileExists(ASCMAIN1.Folders("Temp") & TempExportFilename) Then
                Try
                    My.Computer.FileSystem.DeleteFile(ASCMAIN1.Folders("Temp") & TempExportFilename)

                Catch ex As Exception
                    MsgBox("Perhaps Report is already Open (look at taskbar)" & vbCrLf & vbCrLf & "System Error: " & ex.Message, MsgBoxStyle.OkOnly, "Cannot Export Report")
                    Return REPORT_NO
                End Try
            End If

            Try
                My.Computer.FileSystem.RenameFile(My.Computer.FileSystem.GetFileInfo(DestOpt.DiskFileName).FullName, TempExportFilename)
                REPORT_FILENAMES(REPORT_NO) = ASCMAIN1.Folders("Temp") & TempExportFilename
            Catch ex As Exception

            End Try
        End If

        DestOpt = Nothing

        If Show_Report Then
            Call Show_Reports()
        End If

        Return REPORT_NO
    End Function

    Sub SetParameterValue( _
    ByVal pfName As String, _
    ByVal pfValue As String, _
    Optional ByVal Sub_Report As Boolean = False)

        Dim Par As CrystalDecisions.Shared.ParameterValues = Nothing
        Dim ParD As New CrystalDecisions.Shared.ParameterDiscreteValue()

        Try
            If Sub_Report Then
                Par = ASCMAIN1.CR_SubRpt.DataDefinition.ParameterFields.Item(pfName).CurrentValues
            Else
                Par = ASCMAIN1.CR_RPT.DataDefinition.ParameterFields.Item(pfName).CurrentValues
            End If
        Catch ex As Exception
            MsgBox("Error Adding Parameter " & pfName & " to Report " & ASCMAIN1.CR_RPT.Name)
        End Try

        ParD.Value = pfValue
        Par.Add(ParD)

        Try
            If Sub_Report Then
                ASCMAIN1.CR_SubRpt.DataDefinition.ParameterFields.Item(pfName).ApplyCurrentValues(Par)
            Else
                ASCMAIN1.CR_RPT.DataDefinition.ParameterFields.Item(pfName).ApplyCurrentValues(Par)
            End If
        Catch ex As Exception
            MsgBox("Error Adding Parameter " & pfName & " to Report " & ASCMAIN1.CR_RPT.Name)
        End Try
    End Sub

    Public Sub Show_Reports(Optional ByVal progress As Boolean = False)

        If progress Then
            Call ASCMAIN1.Progress("Showing Reports")
        End If

        Call Set_Table(tblASTSPRF1)


        If progress Then
            Call ASCMAIN1.Progress("Updating Server")
        End If

        Update_Record_TDA("ASTSPRF1")

        If UltraTabControl1 IsNot Nothing AndAlso UltraTabControl1.Tabs.Count > 1 Then UltraTabControl1.SelectedTab = UltraTabControl1.Tabs(1)
    End Sub

    Public Sub Dont_Show_Reports()
        Set_Table(tblASTSPRF1)
        Update_Record_TDA("ASTSPRF1")
    End Sub

    Public Sub Print_Reports( _
    Optional ByVal progress As Boolean = False, _
    Optional ByVal selected_reports_only As Boolean = False, _
    Optional ByVal Record_ASTSPRF1 As Boolean = False, _
    Optional ByVal number_of_copies As Int32 = 1, _
    Optional ByVal streamIPandPort As String = "")

        If Record_ASTSPRF1 Then
            Update_Record_TDA("ASTSPRF1")
        End If

        If progress Then
            Call ASCMAIN1.Progress("Now Printing Report(s)")
        End If
        If tblASTSPRF1 IsNot Nothing Then ' Necessary to check if we ever cann Print_Report_Begin and then do nothing and then call Print_Report_Begin
            For Each row As DataRow In tblASTSPRF1.Select(IIf(selected_reports_only, "SELECT = '1'", ""))
                If progress Then
                    Call ASCMAIN1.Progress("-", row.Item("REPORT_NO"))
                End If
                Call Print_Report(row.Item("REPORT_NO"), number_of_copies, streamIPandPort)
            Next
        End If
        'For Each row As DataRow In dst.Tables("ASTSPRF1").Select(IIf(selected_reports_only, "SELECT = '1'", ""))


        If progress Then
            Call ASCMAIN1.Progress("")
        End If
    End Sub

    Public Sub Set_Table(ByVal t As DataTable)
        tblASTSPRF1 = t
        Me.WindowState = FormWindowState.Normal

        Me.Show()
    End Sub

    Private Sub UltraExplorerBar1_ItemClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinExplorerBar.ItemEventArgs) Handles UltraExplorerBar1.ItemClick

        Select Case e.Item.Key

            Case "Select Printer"
                Dim PrintDialog1 As New PrintDialog()
                Dim result As DialogResult = PrintDialog1.ShowDialog()
                If result.ToString = "OK" Then
                    PRINTER_NAME = PrintDialog1.PrinterSettings.PrinterName
                    UltraExplorerBar1.Groups("Printing Options").Items("Print All").ToolTipText = "Prints All of the Reports" & IIf(PRINTER_NAME = "", "", " to " & PRINTER_NAME)
                End If

            Case "Print All"
                Call Print_Reports(True, False)

            Case "Print Selected"
                If tblASTSPRF1.Select("SELECT = '1'").Length = 0 Then
                    MsgBox("Nothing Selected", MsgBoxStyle.OkOnly, "Cannot Print Selected Reports")
                    Exit Sub
                End If

                Call Print_Reports(True, True)

                'Case "Web Report"

                '    For Each REPORT_NO As String In CRs.Keys
                '        Dim CR As CrystalDecisions.CrystalReports.Engine.ReportDocument = CRs(REPORT_NO)
                '        CR.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, "C:\Projects\ABSolute1\ABSolute1.Web\Reports\" & REPORT_NO & ".PDF")
                '    Next

            Case "email"
                Dim ATTACHMENTs As New Dictionary(Of String, String)
                Dim Subject As String = ""

                Dim report_count As Integer = 0

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Creating PDFs to attach to email")

                For i As Int16 = 1 To UltraTabControl1.Tabs.Count - 1
                    Dim REPORT_NO As String = REPORT_FILENAMES.Keys(i - 1)
                    Dim rowASTSRPT1 As DataRow = dst.Tables("ASTSPRF1").Rows.Find(REPORT_NO)
                    If rowASTSRPT1.Item("SELECT") & "" = "1" Then
                        report_count += 1

                        ASCMAIN1.Progress("-", CStr(report_count))

                        Dim crv As CrystalDecisions.Windows.Forms.CrystalReportViewer = DirectCast(UltraTabControl1.Tabs(i).TabPage.Controls(0), CrystalDecisions.Windows.Forms.CrystalReportViewer)
                        Dim DestOpt As New CrystalDecisions.Shared.DiskFileDestinationOptions ' MicrosoftMailDestinationOptions
                        Dim filename As String = ASCMAIN1.Next_Control_No("ASFSRPTV.EXPORT") & ".PDF"
                        DestOpt.DiskFileName = ASCMAIN1.Folders("Temp") & filename
                        ASCMAIN1.CR_RPT = crv.ReportSource
                        With ASCMAIN1.CR_RPT.ExportOptions
                            .DestinationOptions = DestOpt
                            .ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile ' CrystalDecisions.Shared.ExportDestinationType.MicrosoftMail
                            .FormatOptions = New CrystalDecisions.Shared.PdfRtfWordFormatOptions
                            .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        End With

                        Try
                            ASCMAIN1.CR_RPT.Export()
                            ATTACHMENTs.Add(filename, ASCMAIN1.Folders("Temp") & filename)
                            Subject &= ";" & UltraTabControl1.Tabs(i).Text

                        Catch ex As Exception
                            MsgBox("email attempt failed - Client Station may not be set up for email")
                            Exit Sub
                        End Try

                    End If
                Next


                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")


                If report_count = 0 Then
                    MsgBox("No Reports Selected", MsgBoxStyle.OkOnly, "Cannot email")
                    Exit Sub
                End If

                Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
                'EMAIL_ADDRESSs.Add("wjz@absolution.com", "Walter J. Zielenski")

                Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                       (ASCMAIN1.ActiveForm, Nothing, ATTACHMENTs, _
                         Mid(Subject, 2), "RPT")

            Case "email - Outlook"

                Dim OutlookEmail As New Email()
                Dim Subject As String = ""

                For i As Int16 = 1 To UltraTabControl1.Tabs.Count - 1
                    Dim crv As CrystalDecisions.Windows.Forms.CrystalReportViewer = DirectCast(UltraTabControl1.Tabs(i).TabPage.Controls(0), CrystalDecisions.Windows.Forms.CrystalReportViewer)

                    Dim DestOpt As New CrystalDecisions.Shared.DiskFileDestinationOptions ' MicrosoftMailDestinationOptions

                    Dim filename As String = ASCMAIN1.Next_Control_No("ASFSRPTV.EXPORT") & ".PDF"

                    DestOpt.DiskFileName = ASCMAIN1.Folders("Temp") & filename

                    ASCMAIN1.CR_RPT = crv.ReportSource

                    With ASCMAIN1.CR_RPT.ExportOptions
                        .DestinationOptions = DestOpt
                        .ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile ' CrystalDecisions.Shared.ExportDestinationType.MicrosoftMail
                        .FormatOptions = New CrystalDecisions.Shared.PdfRtfWordFormatOptions
                        .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                    End With

                    Try
                        ASCMAIN1.CR_RPT.Export()

                        OutlookEmail.Message.Attachments.Add(ASCMAIN1.Folders("Temp") & filename, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, , UltraTabControl1.Tabs(i).Text)
                        Subject &= ";" & UltraTabControl1.Tabs(i).Text

                    Catch ex As Exception
                        MsgBox("email attempt failed - Client Station may not be set up for email")
                        Exit Sub
                    End Try

                Next

                OutlookEmail.Message.Subject = Mid(Subject, 2)
                OutlookEmail.Message.Body = vbCrLf & vbCrLf & "Attached is the report you requested." & vbCrLf & vbCrLf & ASCMAIN1.USER_NAME

                OutlookEmail.Show()

            Case "Browse"
                Dim FILENAME As String = ""
                Using openFileDialog1 As New OpenFileDialog
                    openFileDialog1.Title = "Select a File to Publish"
                    openFileDialog1.Filter = "xls files (*.xls)|*.xls|xlsx files (*.xlsx)|*.xlsx|pdf files (*.pdf)|*.pdf|doc files (*.doc)|*.doc"
                    openFileDialog1.RestoreDirectory = True

                    If openFileDialog1.ShowDialog() = DialogResult.OK Then
                        FILENAME = openFileDialog1.FileName
                    End If
                End Using

                If FILENAME <> "" Then
                    Add_File(FILENAME)
                End If

            Case "Publish"
                Publish_to_Portal()


            Case "Exit"
                Me.Close()
        End Select
    End Sub

    Sub Publish_to_Portal()
        If dst.Tables("ASTROPT4").Rows.Count = 0 Then
            MsgBox("Nobody specified in the Distribution List", MsgBoxStyle.OkOnly, "Cannot Publish")
            Exit Sub
        End If

        If grdASTWRPT0.Rows.Count = 0 Then
            MsgBox("No Documents Specified", MsgBoxStyle.OkOnly, "Cannot Publish")
            Exit Sub
        End If


        Dim ok_to_publish_to_portal As String = ""

        If ASCMAIN1.Running_in_VS Then
            If ASCMAIN1.DBS_SERVER = "" Then
                ok_to_publish_to_portal = "N"
                ok_to_publish_to_portal = ""
            End If
            If ok_to_publish_to_portal = "" Then
                If MsgBoxResult.No = MsgBox("Copy File to Portal?", MsgBoxStyle.YesNo, "You are running on a Development Machine") Then
                    ok_to_publish_to_portal = "N"
                Else
                    ok_to_publish_to_portal = "Y"
                End If
            End If
        Else
            ok_to_publish_to_portal = "Y"
        End If

        Dim ATTACHMENTs As New Dictionary(Of String, String)

        For Each row As DataRow In dst.Tables("ASTWRPT0").Select("")
            Dim FILENAME_ORIG As String = row.Item("FILENAME")
            Dim FILETYPE As String = row.Item("REPORT_TYPE")
            Dim REPORT_NO As String = row.Item("REPORT_NO")

            Dim FOLDERNAME As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_WEB_REPORTS_DIR") _
                & "\" & ASCMAIN1.DBS_COMPANY & "\" & Mid(REPORT_NO, 1, 5) & "\"
            Dim FILENAME As String = FOLDERNAME & REPORT_NO & "." & FILETYPE

            If ok_to_publish_to_portal = "Y" Then
                If Not My.Computer.FileSystem.DirectoryExists(FOLDERNAME) Then
                    My.Computer.FileSystem.CreateDirectory(FOLDERNAME)
                End If
                'If ASCMAIN1.Running_in_VS Then
                '    FILENAME_ORIG = ASCMAIN1.Folders("Work") & FILENAME_ORIG
                'End If
                My.Computer.FileSystem.CopyFile(FILENAME_ORIG, FILENAME, False)
            End If

            row.Item("FILENAME") = FILENAME

            '     ATTACHMENTs.Add(RPT_TITLE & IIf(SET_DESC = "", "", " - " & SET_DESC), FILENAME_link)


            'Dim REPORT_NO As String = row.Item("REPORT_NO")
            'Dim FILENAME As String = Export_Report(REPORT_NO, optPublishFormat.Value)
            Dim RPT_TITLE As String = row.Item("RPT_TITLE") & ""
            Dim SET_DESC As String = row.Item("SET_DESC") & ""
            If SET_DESC = "{SubTitle}" Then SET_DESC = ""

            Dim AS_PARM_WEB_REPORTS_URL As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_WEB_REPORTS_URL") & ""
            Dim AS_PARM_WEB_REPORTS_DIR As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_WEB_REPORTS_DIR") & ""
            ''http://portal.interparfums.com/webreports/INT/00076/00076439.PDF
            Dim FILENAME_link As String = Replace(FILENAME, AS_PARM_WEB_REPORTS_DIR, AS_PARM_WEB_REPORTS_URL)
            FILENAME_link = Replace(Replace(Replace(FILENAME_link, "\", "/"), "//", "/"), "http:/", "http://")

            ' ATTACHMENTs.Add(RPT_TITLE & IIf(SET_DESC = "", "", " - " & SET_DESC), FILENAME_link)
            ATTACHMENTs.Add(RPT_TITLE & IIf(SET_DESC = "", "", " - " & SET_DESC), FILENAME_link)
        Next


        Dim USER_IDs As New List(Of String)
        For Each rowASTROPT4 As DataRow In dst.Tables("ASTROPT4").Select
            Dim USER_ID As String = rowASTROPT4.Item("USER_ID")
            USER_IDs.Add(USER_ID)
        Next

        For Each USER_ID As String In USER_IDs
            For Each rowASTWRPT0 As DataRow In dst.Tables("ASTWRPT0").Select("")
                rowASTWRPT0.Item("USER_ID") = USER_ID
                rowASTWRPT0.AcceptChanges()
                rowASTWRPT0.SetAdded()
            Next
            Update_Record_TDA("ASTWRPT0")
        Next


        If chkEmailLinks.Checked Then
            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
            For Each USER_ID As String In USER_IDs
                Dim rowASTUSER1 As DataRow = LookUp("ASTUSER1", USER_ID)
                Dim USER_EMAIL As String = rowASTUSER1.Item("USER_EMAIL")
                Dim USER_NAME As String = rowASTUSER1.Item("USER_NAME")
                If Not EMAIL_ADDRESSs.ContainsKey(USER_EMAIL) Then
                    EMAIL_ADDRESSs.Add(USER_EMAIL, USER_NAME)
                End If
            Next
 
            ASCMAIN1.TACMAIN1.Send_email _
                (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs, _
                 txtSubject.Text, MENU_ITEM_OBJECT & "L", True)
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")

        MsgBox("Documents Published Successfully", MsgBoxStyle.OkOnly, "Verification")
        Me.Close()
    End Sub

    Function Export_Report( _
    ByVal REPORT_NO As String, _
    Optional ByVal ExportFormat As String = "PDF")
        Dim FILENAME As String = REPORT_NO & "." & ExportFormat
        Dim FULL_FILENAME As String 

        If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
            ' until we straighten out the real folders at AHA
            FULL_FILENAME = ASCMAIN1.rowASTPARM1.Item("AS_PARM_WEB_REPORTS_DIR") _
            & "\" & Mid(FILENAME, 1, 7) & "\" & FILENAME
        Else
            FULL_FILENAME = ASCMAIN1.rowASTPARM1.Item("AS_PARM_WEB_REPORTS_DIR") _
             & "\" & ASCMAIN1.DBS_COMPANY & "\" & Mid(REPORT_NO, 1, 5) & "\" & FILENAME
        End If

        If Not My.Computer.FileSystem.FileExists(FULL_FILENAME) Then
            Try
                Dim DestOpt As New CrystalDecisions.Shared.DiskFileDestinationOptions
                DestOpt.DiskFileName = ASCMAIN1.Folders("Temp") & FILENAME
                ASCMAIN1.CR_RPT = CRs(REPORT_NO)

                Dim T As Integer = 0
                For Each CRsREPORT_NO As String In CRs.Keys
                    T += 1
                    If CRsREPORT_NO = REPORT_NO Then
                        Exit For
                    End If
                Next

                Dim crv As CrystalDecisions.Windows.Forms.CrystalReportViewer = DirectCast(UltraTabControl1.Tabs(T).TabPage.Controls(0), CrystalDecisions.Windows.Forms.CrystalReportViewer)
                ASCMAIN1.CR_RPT = crv.ReportSource

                With ASCMAIN1.CR_RPT.ExportOptions
                    .DestinationOptions = DestOpt
                    .ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                    Select Case ExportFormat
                        Case "RPT"
                            .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.CrystalReport
                        Case "XLS"
                            .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.Excel
                        Case "PDF"
                            .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        Case Else
                            .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.CrystalReport
                    End Select
                End With

                ASCMAIN1.Progress("Exporting Report " & REPORT_NO)
                ASCMAIN1.CR_RPT.Export()


                If ASCMAIN1.Running_in_VS Then
                    If ASCMAIN1.DBS_SERVER = "" Then
                        '  ok_to_archive_report = "N"
                        ok_to_archive_report = ""
                    End If
                    If ok_to_archive_report = "" Then
                        If MsgBoxResult.No = MsgBox("Copy File to Portal?", MsgBoxStyle.YesNo, "You are running on a Development Machine") Then
                            ok_to_archive_report = "N"
                        Else
                            ok_to_archive_report = "Y"
                        End If
                    End If
                Else
                    ok_to_archive_report = "Y"
                End If
                If ok_to_archive_report = "Y" Then
                    Try
                        My.Computer.FileSystem.CopyFile(ASCMAIN1.Folders("Temp") & FILENAME, FULL_FILENAME, True)
                    Catch ex As Exception
                        MsgBox("Error " & ex.Message & " Trying to Copy File to Portal Folder", MsgBoxStyle.OkOnly, "Cannot Publish File")
                    End Try

                End If


                ASCMAIN1.Progress("")

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error Exporting Report")
            End Try
        End If

        Return FULL_FILENAME
    End Function

    Sub Print_Report( _
    ByVal REPORT_NO As String, _
    Optional ByVal number_of_copies As Int32 = 1, _
    Optional ByVal streamIPandPort As String = "")

        Dim w2 As Int16 = 0
        Try
            Dim FILENAME As String = ASCMAIN1.DBS_COMPANY & "_" & REPORT_NO & ".RPT"
            If ASCMAIN1.CR_RPT Is Nothing Then
                ASCMAIN1.CR_RPT = New CrystalDecisions.CrystalReports.Engine.ReportDocument
            End If
            ASCMAIN1.CR_RPT.Load(ASCMAIN1.Folders("Temp") & FILENAME)
            If PRINTER_NAME <> "" Then
                ASCMAIN1.CR_RPT.PrintOptions.PrinterName = PRINTER_NAME
            End If
            If streamIPandPort <> "" Then
                Dim tempfilename As String = ASCMAIN1.Folders("Temp") & ASCMAIN1.DBS_COMPANY & "_" & REPORT_NO & ".TMP"
                Dim prtdoc As New System.Drawing.Printing.PrintDocument
                Dim prset As System.Drawing.Printing.PrinterSettings = prtdoc.PrinterSettings
                If PRINTER_NAME <> "" Then
                    prset.PrinterName = PRINTER_NAME
                End If
                prset.PrintFileName = tempfilename
                prset.PrintToFile = True

                ' the following line is sometimes printing to printer instead of printing to file
                Dim pgset As New System.Drawing.Printing.PageSettings
                pgset.PrinterSettings.PrinterName = prset.PrinterName
                pgset.PrinterSettings.PrintToFile = True
                ASCMAIN1.CR_RPT.PrintToPrinter(prset, pgset, False)
                Using ipp As New nsoftware.IPWorks.TCPClient
                    ipp.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareipportkey")

                    ipp.RemoteHost = Split(streamIPandPort, ":")(0)
                    ipp.RemotePort = Val(Split(streamIPandPort, ":")(1))

                    'ipp.Connect(Split(streamIPandPort, ":")(0), Val(Split(streamIPandPort, ":")(1)))
                    While w2 < 12 And Not My.Computer.FileSystem.FileExists(tempfilename)
                        System.Threading.Thread.Sleep(3000)
                        w2 += 1
                        If w2 = 3 Or w2 = 6 Or w2 = 9 Then
                            prset.PrintFileName = tempfilename
                            prset.PrintToFile = True
                            ASCMAIN1.CR_RPT.PrintToPrinter(prset, pgset, False)
                            System.Threading.Thread.Sleep(1000)
                        End If
                    End While
                    Using BSR As New System.IO.BinaryReader(System.IO.File.Open(tempfilename, System.IO.FileMode.Open))
                        Dim W As Int32 = 0
                        Do
                            If Not ipp.Connected Then
                                System.Threading.Thread.Sleep(1000)
                                W += 1
                            End If
                        Loop While Not ipp.Connected And W < 5
                        ipp.Send(BSR.ReadBytes(BSR.BaseStream.Length))
                        System.Threading.Thread.Sleep(1000)
                    End Using
                    'Using BSR As New System.IO.BinaryReader(ASCMAIN1.CR_RPT.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat))
                    '    ipp.Send(BSR.ReadBytes(BSR.BaseStream.Length))
                    'End Using
                    '' ipp.SendStream(ASCMAIN1.CR_RPT.ExportToStream(CrystalDecisions.Shared.ExportFormatType.CrystalReport))
                    ipp.Disconnect()
                End Using
            Else
                ASCMAIN1.CR_RPT.PrintToPrinter(number_of_copies, False, 0, 0)
            End If

        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf _
                & "StreamIPandPort = " & streamIPandPort & vbCrLf _
                & "Printer Name = " & PRINTER_NAME & vbCrLf _
                , MsgBoxStyle.OkOnly, "Error (IP)Printing Report")
            If ASCMAIN1.Running_in_VS Then
                Stop
            End If

        End Try
    End Sub

    Sub Clear_All()
        For Each row As DataRow In dst.Tables("ASTSPRF1").Rows
            row.Item("SELECT") = "0"
        Next
    End Sub

    Private Sub grdASTSPRF1_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdASTSPRF1.InitializeLayout
        Call ASCMAIN1.grdInitializeLayout(grdASTSPRF1)
        grdASTSPRF1.DisplayLayout.AutoFitStyle = UltraWinGrid.AutoFitStyle.ResizeAllColumns
    End Sub

    Sub Mail_PDF_Report( _
    ByVal MailSubject As String, _
    ByVal MailToList As String, _
    ByVal MailMessage As String, _
    ByVal MailCCList As String, _
    ByVal AttachmentFileName As String, _
    Optional ByVal RecordSelectionFormula As String = "")

        ' THIS SECTION ASSUMES THAT THE REPORT TO MAIL IS ASCMAIN1.CR_RPT
        ' THE CODE SHOULD BE USING CRs

        If RecordSelectionFormula <> "" Then
            ASCMAIN1.CR_RPT.RecordSelectionFormula = RecordSelectionFormula
        End If

        Dim DestOpt As New CrystalDecisions.Shared.MicrosoftMailDestinationOptions
        DestOpt.MailSubject = MailSubject
        DestOpt.MailToList = MailToList
        DestOpt.MailMessage = MailMessage
        DestOpt.MailCCList = MailCCList

        With ASCMAIN1.CR_RPT.ExportOptions
            .DestinationOptions = DestOpt
            .ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.MicrosoftMail
            .FormatOptions = New CrystalDecisions.Shared.PdfRtfWordFormatOptions
            .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
        End With

        Try
            ASCMAIN1.CR_RPT.Export()
        Catch ex As Exception
            'Stop
            MsgBox("email attempt failed - Client Station may not be set up for email")
        End Try

    End Sub

    Sub Add_Report(ByVal REPORT_NO As String, _
    ByRef CRP As CrystalDecisions.CrystalReports.Engine.ReportDocument)
        CRs.Add(REPORT_NO, CRP)

        If Not REPORT_FILENAMES.ContainsKey(REPORT_NO) Then
            REPORT_FILENAMES.Add(REPORT_NO, CRP.FileName)
        End If
    End Sub

    Private Sub optPublishVia_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPublishVia.ValueChanged
        optPublishAs.Visible = (optPublishVia.Value = "E")
        lblSubject.Visible = (optPublishVia.Value = "E")
        txtSubject.Visible = (optPublishVia.Value = "E")
    End Sub

    Private Sub cmdPublish_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPublish.Click
        If dst.Tables("ASTROPT4").Rows.Count = 0 Then
            MsgBox("Nobody specified in the Distribution List", MsgBoxStyle.OkOnly, "Cannot Publish")
            Exit Sub
        End If

            If dst.Tables("ASTSPRF1").Select("SELECT = '1'").Length = 0 Then
            MsgBox("No Reports Selected", MsgBoxStyle.OkOnly, "Cannot Publish")
            Exit Sub
        End If

        'If optPublishVia.Value = "E" Then
        '    If dst.Tables("ASTROPT4").Select("USER_EMAIL = NULL").Length <> 0 Then
        '        MsgBox("Nobody specified in the Distribution List", MsgBoxStyle.OkOnly, "Cannot Publish")
        '    End If
        'End If

        Dim USER_IDs As New List(Of String)
        For Each rowASTROPT4 As DataRow In dst.Tables("ASTROPT4").Select
            Dim USER_ID As String = rowASTROPT4.Item("USER_ID")
            USER_IDs.Add(USER_ID)
        Next
        If optPublishVia.Value = "E" Then
            ' FOR WEB REPORTS, USERS WILL PICK UP THESE RECORDS IMPLICITLY FROM GROUP MEMBERSHIP
            For Each rowASTROPTA As DataRow In dst.Tables("ASTROPTA").Select
                Dim USER_ID As String = rowASTROPTA.Item("USER_ID")
                If Not USER_IDs.Contains(USER_ID) Then
                    USER_IDs.Add(USER_ID)
                End If
            Next
        End If

        Dim ATTACHMENTs As New Dictionary(Of String, String)

        dst.Tables("ASTWRPT0").Rows.Clear()

        Dim LINKs As New List(Of String)

        For Each row As DataRow In tblASTSPRF1.Select("SELECT = '1'", "")

            Dim REPORT_NO As String = row.Item("REPORT_NO")
            Dim FILENAME As String = Export_Report(REPORT_NO, optPublishFormat.Value)
            Dim RPT_TITLE As String = row.Item("RPT_TITLE") & ""
            Dim SET_DESC As String = row.Item("SET_DESC") & ""

            If optPublishAs.Value = "L" Then
                Dim AS_PARM_WEB_REPORTS_URL As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_WEB_REPORTS_URL") & ""
                Dim AS_PARM_WEB_REPORTS_DIR As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_WEB_REPORTS_DIR") & ""
                'http://portal.interparfums.com/webreports/INT/00076/00076439.PDF
                Dim FILENAME_link As String = Replace(FILENAME, AS_PARM_WEB_REPORTS_DIR, AS_PARM_WEB_REPORTS_URL)
                FILENAME_link = Replace(Replace(Replace(FILENAME_link, "\", "/"), "//", "/"), "http:/", "http://")
                LINKs.Add(FILENAME_link)
                'ATTACHMENTs.Add(REPORT_NO & "." & optPublishFormat.Value, FILENAME_link)
                ATTACHMENTs.Add(RPT_TITLE & IIf(SET_DESC = "", "", " - " & SET_DESC), FILENAME_link)
            ElseIf optPublishAs.Value = "F" Then
                Dim FILENAME_local As String = ASCMAIN1.Folders("Temp") & REPORT_NO & "." & optPublishFormat.Value
                ATTACHMENTs.Add(REPORT_NO & "." & optPublishFormat.Value, FILENAME_local)
            End If

            'If optPublishVia.Value = "W" Then
            For Each USER_ID As String In USER_IDs
                Dim rowASTWRPT0 As DataRow = dst.Tables("ASTWRPT0").NewRow
                With rowASTWRPT0
                    .Item("USER_ID") = USER_ID
                    .Item("REPORT_NO") = row.Item("REPORT_NO")
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("REPORT_TYPE") = optPublishFormat.Value
                    .Item("RPT_TITLE") = row.Item("RPT_TITLE")
                    .Item("SET_DESC") = row.Item("SET_DESC")
                    .Item("SET_ID") = row.Item("SET_ID")
                    .Item("LINK_URL_TOKEN") = Get_LINK_URL_TOKEN(row.Item("REPORT_NO"))
                End With
                dst.Tables("ASTWRPT0").Rows.Add(rowASTWRPT0)
            Next
            'End If
        Next

        If optPublishVia.Value = "E" Then
            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
            For Each USER_ID As String In USER_IDs
                Dim rowASTUSER1 As DataRow = LookUp("ASTUSER1", USER_ID)
                Dim USER_EMAIL As String = rowASTUSER1.Item("USER_EMAIL")
                Dim USER_NAME As String = rowASTUSER1.Item("USER_NAME")
                If Not EMAIL_ADDRESSs.ContainsKey(USER_EMAIL) Then
                    EMAIL_ADDRESSs.Add(USER_EMAIL, USER_NAME)
                End If
            Next
            ASCMAIN1.TACMAIN1.Send_email _
                (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs, _
                 txtSubject.Text, MENU_ITEM_OBJECT & optPublishAs.Value, True)
        End If

        Update_Record_TDA("ASTWRPT0")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")

        If optPublishVia.Value = "E" Then
            MsgBox("Reports Published Successfully via email", MsgBoxStyle.OkOnly, "Verification")
        Else
            MsgBox("Reports Published Successfully to Web", MsgBoxStyle.OkOnly, "Verification")
        End If

    End Sub

    Function Get_LINK_URL_TOKEN(REPORT_NO As String)
        Dim LINK_URL_TOKEN As String = Guid.NewGuid.ToString
        Return LINK_URL_TOKEN
    End Function
    Private Sub optPublishFormat_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPublishFormat.ValueChanged

    End Sub

    Private Sub optPublishAs_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPublishAs.ValueChanged

    End Sub


#Region "grdASTROPT4"

    Private Sub grdASTROPT4_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTROPT4.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "USER_ID"
                grdCodeDesc(grdASTROPT4, "ASTUSER1", "USER_ID", "USER_NAME")
                grdCodeDesc(grdASTROPT4, "ASTUSER1", "USER_ID", "USER_EMAIL")
                grdCodeDesc(grdASTROPT4, "ASTUSER1", "USER_ID", "USER_STATUS")
        End Select

        If grdASTROPT4.ActiveCell.Column.Key = "USER_ID" And grdASTROPT4.ActiveCell.Value & "" <> "" AndAlso grdASTROPT4.ActiveCell.Value <> grdASTROPT4.ActiveCell.Value.ToString.ToLower Then
            grdASTROPT4.ActiveCell.Value = grdASTROPT4.ActiveCell.Value.ToString.ToLower
        End If

    End Sub

    Private Sub grdASTROPT4_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTROPT4.AfterRowActivate
        With grdASTROPT4.DisplayLayout.Bands(0)
            If grdASTROPT4.ActiveRow.IsAddRow Then
                .Columns("USER_ID").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdASTROPT4.ActiveCell = grdASTROPT4.ActiveRow.Cells("USER_ID")
                grdASTROPT4.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("USER_ID").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdASTROPT4_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTROPT4.AfterRowsDeleted
        For Each USER_GROUP_ID As String In USER_GROUP_IDs
            Delete_Rows("ASTROPTA", "USER_GROUP_ID = '" & USER_GROUP_ID & "'")
        Next
        USER_GROUP_IDs.Clear()
    End Sub

    Private Sub grdASTROPT4_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdASTROPT4.AfterRowUpdate
        If e.Row.Band.Key = "ASTROPT4" Then

            Delete_Rows("ASTROPTA", "USER_GROUP_ID = '" & e.Row.Cells("USER_ID").Text & "'")

            If e.Row.Cells("USER_STATUS").Text = "G" Then
                ASCMAIN1.sql = "Select ASTUSER3.*,ASTUSER1.USER_NAME,ASTUSER1.USER_EMAIL from ASTUSER3,ASTUSER1" _
                & " where ASTUSER1.USER_ID = ASTUSER3.USER_ID and ASTUSER3.USER_GROUP_ID = :PARM1"
                For Each rowASTUSER3 As DataRow In ASCDATA1.GetDataTable _
                    (ASCMAIN1.sql, , , , , "V", New Object() {e.Row.Cells("USER_ID").Text}).Rows
                    Dim rowASTROPTA As DataRow = dst.Tables("ASTROPTA").NewRow
                    rowASTROPTA.Item("USER_GROUP_ID") = rowASTUSER3.Item("USER_GROUP_ID")
                    rowASTROPTA.Item("USER_ID") = rowASTUSER3.Item("USER_ID")
                    rowASTROPTA.Item("USER_NAME") = rowASTUSER3.Item("USER_NAME")
                    rowASTROPTA.Item("USER_EMAIL") = rowASTUSER3.Item("USER_EMAIL")
                    dst.Tables("ASTROPTA").Rows.Add(rowASTROPTA)
                Next
            End If
        End If
    End Sub

    Private Sub grdASTROPT4_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdASTROPT4.BeforeCellUpdate
        'grdFieldFormat(grdASTROPT4)
        'If e.Cell.Column.Key = "USER_ID" And e.Cell.Value & "" <> "" AndAlso e.Cell.Value <> e.Cell.Value.ToString.ToLower Then
        '    e.Cell.Value = e.Cell.Value.ToString.ToLower
        'End If
    End Sub


    Private Sub grdASTROPT4_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdASTROPT4.BeforeExitEditMode
    End Sub

    Private Sub grdASTROPT4_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdASTROPT4.BeforeRowsDeleted

        ' the msgbox to hides behind the modal form
        e.DisplayPromptMsg = False

        'If grdASTROPT4.Selected.Rows.Count <> 1 Then
        '    e.Cancel = True
        'Else
        'End If

        USER_GROUP_IDs.Clear()
        For Each gr As UltraWinGrid.UltraGridRow In grdASTROPT4.Selected.Rows
            USER_GROUP_IDs.Add(gr.Cells("USER_ID").Text)
        Next

    End Sub

    Private Sub grdASTROPT4_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdASTROPT4.BeforeRowUpdate
        With grdASTROPT4

            Dim rowASTUSER1 As DataRow = LookUp("ASTUSER1", e.Row.Cells("USER_ID").Text)
            If rowASTUSER1 Is Nothing Then
                e.Cancel = True
            End If
            If e.Row.IsAddRow Then
                .ActiveRow.Cells("FORM_NAME").Value = FORM_NAME
                .ActiveRow.Cells("SET_ID").Value = SET_ID
            End If
        End With
    End Sub

    Private Sub grdASTROPT4_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTROPT4.ClickCellButton
        Dim sql_where As String = ""
        If grdASTROPT4.ActiveCell.Column.Key = "USER_ID" Then
            'sql_where = " and VEND_CODE = '" & Absx1.txtFor("VEND_CODE").Text & "'"
        End If
        Call grdClickCellButton(grdASTROPT4, sql_where, False)
    End Sub

    Private Sub grdASTROPT4_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdASTROPT4.Error
        grdASTROPT4.ActiveRow.CancelUpdate()
    End Sub
#End Region

    Function Add_File(FILENAME As String) As DataRow

        Dim fi As System.IO.FileInfo

        Try
            fi = My.Computer.FileSystem.GetFileInfo(FILENAME)
            If dst.Tables("ASTATTA1").Rows.Find(Mid(fi.Extension, 2).ToUpper) Is Nothing Then
                MsgBox("Cannot Include " & FILENAME, MsgBoxStyle.OkOnly, "Unsupported File Type (" & fi.Extension & ")")
                Return Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.InnerException.Message, MsgBoxStyle.OkOnly, "Error gettting File Information")
            Return Nothing
        End Try
        
        Dim REPORT_NO As String = ASCMAIN1.Next_Control_No("ASTSPRF1.REPORT_NO")

        Dim rowASTWRPT0 As DataRow = dst.Tables("ASTWRPT0").NewRow
        With rowASTWRPT0
            .Item("USER_ID") = "*"
            .Item("REPORT_NO") = REPORT_NO
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("REPORT_TYPE") = Mid(fi.Extension, 2).ToUpper
            .Item("RPT_TITLE") = fi.Name
            .Item("SET_DESC") = "{SubTitle}"
            .Item("SET_ID") = ""
            .Item("FILENAME") = FILENAME
            .Item("LINK_URL_TOKEN") = Get_LINK_URL_TOKEN(REPORT_NO)
        End With
        dst.Tables("ASTWRPT0").Rows.Add(rowASTWRPT0)
        Return rowASTWRPT0

    End Function

    Private Sub grpPubParms_DragEnter(sender As Object, e As DragEventArgs)
    End Sub

    Private Sub grdASTWRPT0_BeforeRowsDeleted(sender As Object, e As UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdASTWRPT0.BeforeRowsDeleted
        e.DisplayPromptMsg = False
    End Sub

    Private Sub grdASTWRPT0_DragDrop(sender As Object, e As DragEventArgs) Handles grdASTWRPT0.DragDrop
        Application.DoEvents()

        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        If files.Length > 0 Then
            Add_File(files(0))
        End If
    End Sub

    Private Sub grdASTWRPT0_DragEnter(sender As Object, e As DragEventArgs) Handles grdASTWRPT0.DragEnter
        If grdASTWRPT0.AllowDrop Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    Private Sub cmdSaveList_Click(sender As Object, e As EventArgs) Handles cmdSaveList.Click

        If txtLIST_DESC.Text = "" Then
            MsgBox("Please Enter a Description for this List", MsgBoxStyle.OkOnly, "Cannot Save List")
            Exit Sub
        End If

        If grdASTROPT4.Rows.Count = 0 Then
            MsgBox("Please Enter at leat 1 User for this List", MsgBoxStyle.OkOnly, "Cannot Save List")
            Exit Sub
        End If

        Dim LIST_CODE As String = ""

        Dim sqld As String = ""

        If cmdDeleteList.Tag <> "" Then
            Dim iresponse As Microsoft.VisualBasic.MsgBoxResult

            If Not chkLIST_MODIFIABLE.Checked Then
                iresponse = MsgBoxResult.No
            Else
                iresponse = MsgBox("Overwrite List '" & Split(cmdDeleteList.Tag, vbTab)(1) & "'?", MsgBoxStyle.YesNoCancel, "Option to Replace List")
                If iresponse = MsgBoxResult.Yes Then
                    LIST_CODE = Split(cmdDeleteList.Tag, vbTab)(0)
                    sqld = "LIST_CODE = '" & LIST_CODE & "'"
                ElseIf iresponse = MsgBoxResult.No Then
                    cmdDeleteList.Tag = ""
                    cmdDeleteList.Visible = False
                Else
                    Exit Sub
                End If
            End If
        End If

        If LIST_CODE = "" Then LIST_CODE = ASCMAIN1.Next_Control_No("ASTLIST1.LIST_CODE")

        dst.Tables("ASTLIST1").Rows.Clear()
        dst.Tables("ASTLIST2").Rows.Clear()

        Dim rowASTLIST1 As DataRow = dst.Tables("ASTLIST1").NewRow
        LIST_CODE = ASCMAIN1.Next_Control_No("ASTLIST1.LIST_CODE")

        rowASTLIST1.Item("LIST_CODE") = LIST_CODE
        rowASTLIST1.Item("LIST_DESC") = txtLIST_DESC.Text
        rowASTLIST1.Item("COLUMN_NAME") = "USER_ID"
        rowASTLIST1.Item("INIT_DATE") = DATETIME_STAMP
        rowASTLIST1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowASTLIST1.Item("LAST_DATE") = DATETIME_STAMP
        rowASTLIST1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowASTLIST1.Item("LIST_SHAREABLE") = IIf(chkLIST_SHAREABLE.Checked, "1", "0")
        rowASTLIST1.Item("LIST_MODIFIABLE") = IIf(chkLIST_MODIFIABLE.Checked, "1", "0")

        dst.Tables("ASTLIST1").Rows.Add(rowASTLIST1)

        For Each row As DataRow In dst.Tables("ASTROPT4").Select("")
            Dim rowASTLIST2 As DataRow = dst.Tables("ASTLIST2").NewRow
            rowASTLIST2.Item("LIST_CODE") = LIST_CODE
            rowASTLIST2.Item("CODE_VALUE") = row.Item("USER_ID")
            dst.Tables("ASTLIST2").Rows.Add(rowASTLIST2)
        Next

        Update_Record_TDA("ASTLIST1", sqld)
        Update_Record_TDA("ASTLIST2", sqld)

        MsgBox("Distribution List has been Saved", MsgBoxStyle.OkOnly, "Verification")

        Clear_Delete()

    End Sub

    Sub Clear_Delete()

        cmdDeleteList.Visible = False
        cmdClear.Visible = False
        lblLIST_CODE.Visible = False
        txtLIST_DESC.Text = ""
        chkLIST_SHAREABLE.Checked = False
        chkLIST_MODIFIABLE.Checked = False
        chkLIST_SHAREABLE.Enabled = True
        chkLIST_MODIFIABLE.Enabled = True
    End Sub

    Private Sub cmdDeleteList_Click(sender As Object, e As EventArgs) Handles cmdDeleteList.Click
        Dim LIST_CODE As String = lblLIST_CODE.Text
        If MsgBox("OK to Delete Distribution List " & LIST_CODE & "?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
            ASCDATA1.ExecuteSQL("Delete from ASTLIST1 where LIST_CODE = '" & LIST_CODE & "'")
            ASCDATA1.ExecuteSQL("Delete from ASTLIST2 where LIST_CODE = '" & LIST_CODE & "'")
            MsgBox("List " & "" & " has been Deleted", MsgBoxStyle.OkOnly, "Verification")
            Clear_Delete()
        End If
    End Sub

    Private Sub cmdDistributionLists_Click(sender As Object, e As EventArgs) Handles cmdDistributionLists.Click
        Dim sql_where As String = "COLUMN_NAME = 'USER_ID' and (INIT_OPER = '" & ASCMAIN1.USER_ID & "' or LIST_SHAREABLE = '1')"
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("LIST_CODE", "ASTLIST1", sql_where)
        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = False
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Dim LIST_CODE As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("LIST_CODE")
                Dim LIST_DESC As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("LIST_DESC")

                lblLIST_CODE.Text = LIST_CODE
                lblLIST_CODE.Visible = True

                txtLIST_DESC.Text = LIST_DESC

                Dim rowASTLIST1 As DataRow = LookUp("ASTLIST1", LIST_CODE)
                chkLIST_SHAREABLE.Checked = (rowASTLIST1.Item("LIST_SHAREABLE") & "" = "1")
                chkLIST_MODIFIABLE.Checked = (rowASTLIST1.Item("LIST_MODIFIABLE") & "" = "1")

                Dim tf As Boolean = (chkLIST_MODIFIABLE.Checked Or rowASTLIST1.Item("INIT_OPER") & "" = ASCMAIN1.USER_ID)
                chkLIST_SHAREABLE.Enabled = tf
                chkLIST_MODIFIABLE.Enabled = tf
 
                cmdDeleteList.Visible = tf
                cmdDeleteList.Tag = LIST_CODE & vbTab & LIST_DESC
                cmdClear.Visible = True

                ASCMAIN1.sql = "Select * from ASTLIST2 where LIST_CODE = '" & LIST_CODE & "'"
                For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "CODE_VALUE")
                    Dim USER_ID As String = row.Item("CODE_VALUE")
                    If dst.Tables("ASTROPT4").Rows.Find(New String() {FORM_NAME, SET_ID, USER_ID}) Is Nothing Then
                        If grdASTROPT4.ActiveRow IsNot Nothing AndAlso grdASTROPT4.ActiveRow.IsAddRow Then
                            grdASTROPT4.ActiveRow = Nothing
                        End If
                        grdASTROPT4.DisplayLayout.Bands(0).AddNew()
                        With grdASTROPT4.ActiveRow
                            .Cells("USER_ID").Value = USER_ID
                            .Update()
                        End With
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub cmdClear_Click(sender As Object, e As EventArgs) Handles cmdClear.Click
        Clear_Delete()
    End Sub

    Private Sub chkEmailLinks_CheckedChanged(sender As Object, e As EventArgs) Handles chkEmailLinks.CheckedChanged

    End Sub
End Class

    'Namespace Common.Outlook

    ' <summary>
    ' Provides functionality so Outlook can be used to generate Emails.  Use the
    ' write event to log the email.
    ' <example>
    ' Email outlookMsg = new Email();
    ' outlookMsg.Message.Subject = "Foo bar";
    ' outlookMsg.Show();
    ' </example>
    ' </summary>

Public Class Email

    ' <summary>
    ' Gets the current instance of Outlook
    ' </summary>

    Private _outlookInstance As OutlookApp.Application = New Microsoft.Office.Interop.Outlook.Application
    Private _message As OutlookApp.MailItem

    Public Property Message() As OutlookApp.MailItem
        Get
            Return _message
        End Get
        Set(ByVal value As OutlookApp.MailItem)
            _message = value
        End Set
    End Property


    ' <summary>
    ' Constructor, gets current outlook instance
    ' and creates a blank email message
    ' </summary>

    Public Sub New()
        Initialize()
    End Sub

    Private Sub Initialize()
        ' create a blank email
        _message = _outlookInstance.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
        ' _message = (OutlookApp.MailItem)_outlookInstance.CreateItem(OutlookApp.OlItemType.olMailItem);

        ' wire up the write event for logging
        AddHandler _message.Write, AddressOf Message_Write
        ' _message.Write += new Microsoft.Office.Interop.Outlook.ItemEvents_10_WriteEventHandler(Message_Write);

    End Sub

    ' <summary>
    '  Used for logging after the end user presses the send
    '  button in Outlook.  If you need to log the email that was
    '  sent to a web service or something else, fill this in.  This is
    '  called after the email is sent via Outlook.
    ' </summary>

    ' <param name="Cancel"></param>

    Sub Message_Write(ByRef Cancel As Boolean)
        'ADD LOGGING HERE IF YOU NEED IT
        ' <summary>
        ' Displays the outlook screen and shows the email message.
        ' </summary>

    End Sub

    Public Sub Show()
        _message.Display(False)
    End Sub

End Class