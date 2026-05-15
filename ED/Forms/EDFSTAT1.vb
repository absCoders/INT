Public Class EDFSTAT1

    Private sqlFlattened As String = String.Empty
    Private DTE0 As String = String.Empty
    Private DTE1 As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = " Select " & vbCrLf _
            & " INTERCHANGE.""InterchangeKEY""  INTERCHANGE_KEY, " & vbCrLf _
            & " DOC.""PartnerKEY"" PARTNER_KEY, ""PartnerName"" PARTNER_NAME, " & vbCrLf _
            & " DECODE(INTERCHANGE.""Direction"",'0','Inbound','Outbound') DIRECTION," & vbCrLf _
            & " TRIM(INTERCHANGE.""ControlNumber"") ISA_CTL_NO ,INTERCHANGE.""Filename"" INT_FILENAME," & vbCrLf _
            & " Trunc(NEW_TIME((TO_DATE('01/01/1970', 'MM-DD-YYYY') + INTERCHANGE.""TimeCreated""/86400), 'GMT', 'EST')) ISA_INIT_DATE," & vbCrLf _
            & " NEW_TIME((TO_DATE('01/01/1970', 'MM-DD-YYYY') + INTERCHANGE.""TimeCreated""/86400), 'GMT', 'EST') ISA_INIT_DATETIME," & vbCrLf _
            & " DOC.""InterchangeVersion"" ISA_VERSION, DOC.""GroupVersion"" GROUP_VERSION, DOC.""DocumentVersion"" DOC_VERSION," & vbCrLf _
            & " DOC.""DocumentName"" DOC_NAME," & vbCrLf _
            & " DOC.""ReferenceData"" DOC_REF," & vbCrLf _
            & " DOC.""AppField1"" APP_FIELD1," & vbCrLf _
            & " DOC.""AppField2"" APP_FIELD2," & vbCrLf _
            & " DOC.""AppField3"" APP_FIELD3," & vbCrLf _
            & " DOC.""DocumentKEY"" DOC_KEY," & vbCrLf _
            & " DOC.""TransactionSetID"" TRANS_NO," & vbCrLf _
            & " DOC.""FunctionalGroupID"" FG_ID," & vbCrLf _
            & " Trunc(NEW_TIME((TO_DATE('01/01/1970', 'MM-DD-YYYY') + DOC.""TimeCreated""/86400), 'GMT', 'EST')) DOC_DATE, " & vbCrLf _
            & "   DOC.""ControlNumber"" DOC_CTL_NO," & vbCrLf _
            & "   DECODE(""LocationStatus"",'0','In Drawer','1','Out Drawer','2','In Documents','3','?In Documents'," & vbCrLf _
            & "   '4','Out Documents','5','?Out Documents','6','Workspace','7','Queued',""LocationStatus"")  DOC_LOCATION," & vbCrLf _
            & " DECODE(""ComplianceStatus"",'0','Incomplete','1','NonCompliant','2','OK','3','DocQueued'," & vbCrLf _
            & " '4','Sent','5','Net Received','6','Net Delivered','7','Acknowledged','8','Waiting for Ack','9','Ack Overdue'," & vbCrLf _
            & " '10','NetWarning','11','NetError','12','AckErr','13','FAPartial','14','FAReject','15','NetPickedUp','16','Duplicate'," & vbCrLf _
            & " '17','ReadyToSend','18','SendFailed',""ComplianceStatus"") DOC_STATUS," & vbCrLf _
            & " DOC.""DocumentBlobKEY"" DOC_BLOB_KEY" & vbCrLf _
            & "   from GEN.""Document_tb"" DOC, GEN.""Partner_tb"" PARTNER, " & vbCrLf _
            & "   GEN.""Track_tb"" TRACK, GEN.""Interchange_tb"" INTERCHANGE" & vbCrLf _
            & "   Where PARTNER.""PartnerKEY"" = DOC.""PartnerKEY""" & vbCrLf _
            & "   and TRACK.""InterchangeKEY"" = INTERCHANGE.""InterchangeKEY""" & vbCrLf _
            & "   and DOC.""DocumentKEY"" = TRACK.""DocumentKEY""" & vbCrLf _
            & "   and TRUNC(NEW_TIME((TO_DATE('01/01/1970', 'MM-DD-YYYY') + DOC.""TimeCreated""/86400), 'GMT', 'EST'))" & vbCrLf _
            & "   Between :PARM1 and :PARM2"
            sqlFlattened = ASCMAIN1.sql

            Create_TDA(.Tables.Add, "FLAT", ASCMAIN1.sql, 0, False, "VV", 0)

            .Tables.Add("INTERCHANGE")
            .Tables("INTERCHANGE").Columns.Add("INTERCHANGE_KEY", GetType(System.String))
            .Tables("INTERCHANGE").Columns.Add("PARTNER_KEY", GetType(System.String))
            .Tables("INTERCHANGE").Columns.Add("PARTNER_NAME", GetType(System.String))
            .Tables("INTERCHANGE").Columns.Add("DIRECTION", GetType(System.String))
            .Tables("INTERCHANGE").Columns.Add("ISA_CTL_NO", GetType(System.String))
            .Tables("INTERCHANGE").Columns.Add("ISA_INIT_DATE", GetType(System.DateTime))
            .Tables("INTERCHANGE").Columns.Add("ISA_INIT_DATETIME", GetType(System.DateTime))
            .Tables("INTERCHANGE").Columns.Add("ISA_VERSION", GetType(System.String))
            .Tables("INTERCHANGE").Columns.Add("INT_FILENAME", GetType(System.String))
            .Tables("INTERCHANGE").Columns.Add("NUM_RECORDS", GetType(System.Int64))

            '.Tables.Add("DOCUMENTS")
            '.Tables("DOCUMENTS").Columns.Add("INTERCHANGE_KEY", GetType(System.String))
            '.Tables("DOCUMENTS").Columns.Add("DOC_DATE", GetType(System.DateTime))
            '.Tables("DOCUMENTS").Columns.Add("DOC_CTL_NO", GetType(System.String))
            '.Tables("DOCUMENTS").Columns.Add("DOC_NAME", GetType(System.String))
            '.Tables("DOCUMENTS").Columns.Add("DOC_REF", GetType(System.String))
            '.Tables("DOCUMENTS").Columns.Add("DOC_STATUS", GetType(System.String))
            '.Tables("DOCUMENTS").Columns.Add("TRANS_NO", GetType(System.String))
            '.Tables("DOCUMENTS").Columns.Add("FG_ID", GetType(System.String))
            '.Tables("DOCUMENTS").Columns.Add("DOC_VERSION", GetType(System.String))
            '.Tables("DOCUMENTS").Columns.Add("DOC_BLOB_KEY", GetType(System.String))
            '.Tables("DOCUMENTS").Columns.Add("APP_FIELD1", GetType(System.String))
            '.Tables("DOCUMENTS").Columns.Add("APP_FIELD2", GetType(System.String))
            '.Tables("DOCUMENTS").Columns.Add("APP_FIELD3", GetType(System.String))

        End With

        Get_PARM("EDTPARM1")

        grdFlat.DataSource = dst.Tables("FLAT")
        grdInterchanges.DataSource = dst.Tables("INTERCHANGE")
        grdDocuments.DataSource = dst.Tables("FLAT")

        MyBase.Absx1.dteFor("DTE0").MaxDate = DateAdd(DateInterval.Day, 1, DateTime.Now)
        MyBase.Absx1.dteFor("DTE0").MinDate = DateAdd(DateInterval.Month, -36, DateTime.Now)

        MyBase.Absx1.dteFor("DTE1").MaxDate = MyBase.Absx1.dteFor("DTE0").MaxDate
        MyBase.Absx1.dteFor("DTE1").MinDate = MyBase.Absx1.dteFor("DTE0").MinDate

        Create_Summary(grdInterchanges, "PARTNER_KEY", "Count")
        Create_Summary(grdDocuments, "DOC_DATE", "Count")
        Create_Summary(grdFlat, "INTERCHANGE_KEY", "Count")

        grdFlat.DisplayLayout.Bands(0).Columns("INTERCHANGE_KEY").Format = "####"
        grdFlat.DisplayLayout.Bands(0).Columns("DOC_KEY").Format = "####"


        MyBase.Absx1.dteFor("DTE1").DateTime = DateTime.Now
        MyBase.Absx1.dteFor("DTE0").DateTime = DateAdd(DateInterval.Day, -1, DateTime.Now)

        ASCMAIN1.sql = "SELECT ""PartnerKEY"" PARTNER_CODE, ""PartnerName"" PARTNER_NAME FROM GEN.""Partner_tb"""
        Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
        tbl.Columns("PARTNER_CODE").MaxLength = 50
        tbl.Rows.Add(New Object() {"{All Partners}", "All Partners"})
        cmbPartnerCode.DataSource = tbl
        Sort_cmbColumns(cmbPartnerCode, "PARTNER_CODE")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"
                ' Validate Dates
                If CDate(MyBase.Absx1.dteFor("DTE0").DateTime.ToShortDateString) > CDate(MyBase.Absx1.dteFor("DTE1").DateTime.ToShortDateString) Then
                    EMsg &= vbCr & "Start Date must be Less Equal End Date."
                Else
                    DTE0 = MyBase.Absx1.dteFor("DTE0").DateTime.ToString("dd-MMM-yyyy")
                    DTE1 = MyBase.Absx1.dteFor("DTE1").DateTime.ToString("dd-MMM-yyyy")
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "View"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("View").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabMain.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"FLAT", "INTERCHANGE"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        cmbPartnerCode.SelectedRow = cmbPartnerCode.Rows(0)
        DTE0 = String.Empty
        DTE1 = String.Empty

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Gathering Data", "")

        ASCMAIN1.sql = sqlFlattened

        ASCMAIN1.sql = ASCMAIN1.sql.Replace(":PARM1", "'" & DTE0 & "'")
        ASCMAIN1.sql = ASCMAIN1.sql.Replace(":PARM2", "'" & DTE1 & "'")

        ' See if it is for a specified partner
        If Not cmbPartnerCode.Text.StartsWith("{") Then
            ASCMAIN1.sql &= " AND DOC.""PartnerKEY"" = '" & cmbPartnerCode.Text & "'"
        End If

        Select Case optDirection.Value
            Case "0"
                ASCMAIN1.sql &= " AND DOC.""Direction"" = '0'"
            Case "1"
                ASCMAIN1.sql &= " AND DOC.""Direction"" = '1'"
        End Select

        If chkNotAcknowledged.Checked Then
            ASCMAIN1.sql &= " And DOC.""ComplianceStatus""  in  ('8','9','12','13','14')"
        End If


        EnforceConstraints(False)
        Fill_Records("FLAT", String.Empty, True, ASCMAIN1.sql)

        Dim rowFLAT As DataRow = Nothing
        Dim rowINTERCHANGE As DataRow = Nothing
        Dim rowDOCUMENT As DataRow = Nothing

        For Each row As DataRow In ASCDATA1.SelectDistinct("FLAT", New String() {"INTERCHANGE_KEY"}).Select("")
            rowFLAT = dst.Tables("FLAT").Select("INTERCHANGE_KEY = '" & row.Item("INTERCHANGE_KEY") & "'")(0)
            ASCMAIN1.Progress("-", rowFLAT.Item("INTERCHANGE_KEY") & "")

            rowINTERCHANGE = dst.Tables("INTERCHANGE").NewRow
            rowINTERCHANGE.Item("INTERCHANGE_KEY") = rowFLAT.Item("INTERCHANGE_KEY")
            rowINTERCHANGE.Item("PARTNER_KEY") = rowFLAT.Item("PARTNER_KEY")
            rowINTERCHANGE.Item("PARTNER_NAME") = rowFLAT.Item("PARTNER_NAME")
            rowINTERCHANGE.Item("DIRECTION") = rowFLAT.Item("DIRECTION")
            rowINTERCHANGE.Item("ISA_CTL_NO") = rowFLAT.Item("ISA_CTL_NO")
            rowINTERCHANGE.Item("ISA_INIT_DATE") = rowFLAT.Item("ISA_INIT_DATE")
            rowINTERCHANGE.Item("ISA_INIT_DATETIME") = rowFLAT.Item("ISA_INIT_DATETIME")
            rowINTERCHANGE.Item("ISA_VERSION") = rowFLAT.Item("ISA_VERSION")
            rowINTERCHANGE.Item("NUM_RECORDS") = dst.Tables("FLAT").Select("INTERCHANGE_KEY = '" & row.Item("INTERCHANGE_KEY") & "'").Length
            rowINTERCHANGE.Item("INT_FILENAME") = rowFLAT.Item("INT_FILENAME")
            dst.Tables("INTERCHANGE").Rows.Add(rowINTERCHANGE)
        Next

        EnforceConstraints(True)

        Sort_grdColumns(grdInterchanges, "ISA_INIT_DATETIME".ToLower)
        Sort_grdColumns(grdDocuments, "DOC_DATE")
        Sort_grdColumns(grdFlat, "INTERCHANGE_KEY")

        ASCMAIN1.Progress("", "")
        Me.Cursor = Cursors.Default

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdInterchanges, "SSS", "Show Filter", "Show GroupBox", "Show Pins", "Show Raw EDI")
        Load_Popup_Menu(grdDocuments, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Show Raw EDI")
        Load_Popup_Menu(grdFlat, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Show Raw EDI")

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

        Select Case e.SourceControl.Name
            Case ""

        End Select

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case ""

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Show Raw EDI"

                Try
                    Dim ED_PARM_RAW_ARCHIVE As String = String.Empty

                    If ROWs("EDTPARM1").Table.Columns.Contains("ED_PARM_RAW_ARCHIVE") Then
                        ED_PARM_RAW_ARCHIVE = ROWs("EDTPARM1").Item("ED_PARM_RAW_ARCHIVE") & String.Empty
                        ED_PARM_RAW_ARCHIVE = ED_PARM_RAW_ARCHIVE.Trim
                        If ED_PARM_RAW_ARCHIVE.Length > 0 AndAlso Not ED_PARM_RAW_ARCHIVE.EndsWith("\") Then
                            ED_PARM_RAW_ARCHIVE &= "\"
                        End If
                    Else
                        ED_PARM_RAW_ARCHIVE = "\\192.168.170.103\gensrvnt\Documents\"
                    End If
                    ED_PARM_RAW_ARCHIVE = ED_PARM_RAW_ARCHIVE.ToUpper

                    Dim RAW_DATA As String = String.Empty
                    Dim RAW_DATA_FILE As String = String.Empty
                    Dim RAW_DATA_FILE_EXT As String = String.Empty

                    If grd.Name = "grdInterchanges" Then
                        ED_PARM_RAW_ARCHIVE = ED_PARM_RAW_ARCHIVE.Replace("DOCUMENTS\", String.Empty)
                        Dim DIRECTION As String = grd.ActiveRow.Cells("DIRECTION").Value
                        ED_PARM_RAW_ARCHIVE &= IIf(DIRECTION = "Inbound", "IntIn", "IntOut") & "\"
                        RAW_DATA_FILE = grd.ActiveRow.Cells("INT_FILENAME").Value
                    Else
                        RAW_DATA_FILE = grd.ActiveRow.Cells("DOC_BLOB_KEY").Value
                        RAW_DATA_FILE_EXT = ".DOC"
                    End If

                    If RAW_DATA_FILE <> "" Then
                        Dim FILENAME As String = ED_PARM_RAW_ARCHIVE & RAW_DATA_FILE & RAW_DATA_FILE_EXT
                        If My.Computer.FileSystem.FileExists(FILENAME) Then
                            RAW_DATA = My.Computer.FileSystem.ReadAllText(FILENAME)
                        Else
                            MessageBox.Show("File: " & FILENAME & " cannot be found.", "Show Raw EDI", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Exit Sub
                        End If
                    End If

                    Using frm As New ASFTEXT1
                        frm.t = RAW_DATA
                        frm.Text = "Raw EDI for " & RAW_DATA_FILE
                        frm.ShowDialog()
                    End Using
                Catch ex As Exception
                    MessageBox.Show("The following error occurred: " & ex.Message, "Show Raw EDI", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try


        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

#End Region

    Private Sub grdInterchanges_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdInterchanges.AfterRowActivate

        Dim INTERCHANGE_KEY As String = "0"
        Dim ISA_CTL_NO As String = String.Empty

        If grdInterchanges.ActiveRow IsNot Nothing AndAlso grdInterchanges.ActiveRow.IsDataRow Then
            INTERCHANGE_KEY = grdInterchanges.ActiveRow.Cells("INTERCHANGE_KEY").Value
            ISA_CTL_NO = grdInterchanges.ActiveRow.Cells("ISA_CTL_NO").Value
        End If

        Dim viewDocuments As New DataView(dst.Tables("FLAT"))
        viewDocuments.RowFilter = "INTERCHANGE_KEY = '" & INTERCHANGE_KEY & "'"
        grdDocuments.DataSource = viewDocuments
        grdDocuments.Text = "Documents for ISA Ctl No: " & ISA_CTL_NO

    End Sub

End Class