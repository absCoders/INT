Imports Infragistics.Win.UltraWinGrid

Public Class SPFCWRXG

    Dim FILENAME As String

    Dim FILENAMEs_to_Archive As New List(Of String)
    Dim COLUMN_NAMEs() As String = {"RSSP_TAX_ID", "RSSP_ID", "RSSP_NAME_LAST", "RSSP_NAME_FIRST", "RSSP_TITLE",
                            "RSSP_ADDR1", "RSSP_ADDR2", "RSSP_CITY", "RSSP_STATE", "RSSP_ZIP_CODE",
                            "RSSP_SHIP_TO_ADDR1", "RSSP_SHIP_TO_ADDR2", "RSSP_SHIP_TO_CITY", "RSSP_SHIP_TO_STATE", "RSSP_SHIP_TO_ZIP_CODE",
                            "RSSP_EMAIL", "RSSP_PHONE", "RSSP_CELL", "RSSP_PAY_RATE", "SELL_NAME", "SELL_CODE",
                            "RSSP_DATE_HIRED", "RSSP_DATE_CHANGED", "RSSP_DATE_TERM", "RECORD_TYPE"}
    Dim IMPORT_XNO As String
    Dim update_blocked As Boolean
    Dim YW_INVOICE As String
    Dim replacement_method As Boolean = False
    Dim import_errors As New List(Of String)
    Dim tblEXP As DataTable

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SPTPARM1")
        Get_PARM("GLTPARM1")
        Get_PARM("APTPARM1")

        grdSPTARCHD.Visible = False
        SplitContainer3.Panel2Collapsed = True
        With dst

            Create_TDA(.Tables.Add, "SPTRSSP1", "*", 0)
            Create_TDA(.Tables.Add, "SPTRSSP2", "*", 0)

            Create_TDA(.Tables.Add, "SPTCWRX1", "*", 0)
            Create_TDA(.Tables.Add, "SPTCWRX2", "*", 0)
            Create_TDA(.Tables.Add, "SPTCWRX3", "*", 0)

            Create_TDA(.Tables.Add, "SPTCWRXP", "*", 0)
            Create_TDA(.Tables.Add, "SPTCWRX4", "*", 0)


            Create_TDA(.Tables.Add, "APTINVH1", "*", 0)
            Create_TDA(.Tables.Add, "APTINVH2", "*", 0)
            Create_TDA(.Tables.Add, "ASTATTA2", "*", 0)

            Create_TDA(.Tables.Add, "SPTPYXI1", "*", 0)
            Create_TDA(.Tables.Add, "SPTPYXI2", "*", 0)
            Create_Relation("SPTPYXI1", "SPTPYXI2", "CTRL_NO,CTRL_LNO")

            Create_TDA(.Tables.Add, "SPTCWRXC", "*", 0, False)
            Create_TDA(.Tables.Add, "SPTCWRXA", "*", 0, False)
            Create_TDA(.Tables.Add, "GLTPARM3", "*", 0, False)

            Create_TDA(.Tables.Add, "ARTCUST2", "*", 0, False)
            Create_TDA(.Tables.Add, "ARTCUST1", "*", 0, False)

            With .Tables.Add("SPTCWRXF")
                .Columns.Add("FILENAME")
                .Columns.Add("FILEDATE", GetType(System.DateTime))
                .Columns.Add("FILESIZE", GetType(System.Int64))
                .Columns.Add("FILEABBR")
            End With

            With .Tables.Add("SPTARCHD")
                .Columns.Add("FILENAME", GetType(String))
                .Columns.Add("FILEDATE", GetType(System.DateTime))
                .Columns.Add("FILESIZE", GetType(System.Int64))
                .Columns.Add("FILEABBR")
            End With

            'With .Tables.Add("SPTCWRXI")
            '    .Columns.Add("FILENAME")
            'End With

            Create_TDA(.Tables.Add, "ARTCUSTT", "*")

        End With

        Fill_Records("ARTCUST1")
        Fill_Records("ARTCUST2")
        Fill_Records("SPTCWRXA")
        Fill_Records("GLTPARM3")
        Fill_Records("ARTCUSTT", String.Empty, True, "SELECT * FROM ARTCUSTT")
        Fill_Records("SPTCWRXP")


        grdSPTCWRXF.DataSource = dst.Tables("SPTCWRXF")
        grdSPTRSSP2.DataSource = dst.Tables("SPTRSSP2")
        grdSPTCWRX2.DataSource = dst.Tables("SPTCWRX2")
        grdSPTPYXI2.DataSource = dst.Tables("SPTPYXI1")
        grdSPTCWRX4.DataSource = dst.Tables("SPTCWRX4")
        grdSPTARCHD.DataSource = dst.Tables("SPTARCHD")
        Sort_grdColumns(grdSPTARCHD, "filedate")

        Create_Summary(grdSPTCWRX2, "SALES_CHECKBOOK")
        Create_Summary(grdSPTCWRX2, "SALES_OTHER")
        Create_Summary(grdSPTCWRX2, "WAGES_SHIFT")
        Create_Summary(grdSPTCWRX2, "WAGES_RETRO")
        Create_Summary(grdSPTCWRX2, "WAGES_TRAIN")
        Create_Summary(grdSPTCWRX2, "WAGES_SICK")
        Create_Summary(grdSPTCWRX2, "WAGES_BONUS")
        Create_Summary(grdSPTCWRX2, "WAGES_EXP")
        Create_Summary(grdSPTCWRX2, "TRAVEL_EXP")
        Create_Summary(grdSPTCWRX2, "OTHER_TAXES")
        Create_Summary(grdSPTCWRX2, "COWORX_FEE")
        Create_Summary(grdSPTCWRX2, "BILL_AMT")
        Create_Summary(grdSPTCWRX2, "NET_AMT")

        Create_Summary(grdSPTPYXI2, "SPEND_AMT")
        Create_Summary(grdSPTPYXI2, "TOT_SALES_AMT")

        Create_Summary(grdSPTCWRX4, "CWRX_EXP_PCT")
        Create_Summary(grdSPTCWRX4, "CWRX_EXP_AMT")


        Fill_Records("SPTCWRXC")

        Create_Summary(grdSPTCWRXF, "FILENAME", "Count")
        Create_Summary(grdSPTARCHD, "FILENAME", "Count")

        If Me.replacement_method Then
            MsgBox("You are now in File Replacement Mode", MsgBoxStyle.OkOnly, "Exit if you don't think you should be here")

        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "New"

            Case "Load Files"
                Dim INV_files_count As Integer = dst.Tables("SPTCWRXF").Select("FILEABBR LIKE 'INV%'", "").Length
                Dim PYX_files_count As Integer = dst.Tables("SPTCWRXF").Select("FILEABBR LIKE 'PYX%'", "").Length
                If INV_files_count = 0 And PYX_files_count = 0 Then
                    ' this is ok - no invoice nor payroll export files
                Else
                    If INV_files_count <> PYX_files_count Or INV_files_count <> 1 Then
                        EMsg &= vbCr & "This function is designed to import" & vbCr & " a matched pair of INV/PYX files each week" & vbCr & " and there appears to be a mis-match." & vbCr & "Please call ABS for support"
                    End If
                End If

            Case "Update"

                If tblEXP IsNot Nothing Then
                    If tblEXP.Rows.Count > 0 Then
                        Dim YWMAX As String = tblEXP.Compute("MAX(YW)", "")
                        Dim ROW = tblEXP.Rows.Find(YWMAX)
                        If Val(ROW.ITEM("NET_AMT_PREV") & "") <> 0 Then
                            If MsgBox("There are postings in the latest week referenced by This File" _
                                      & vbCrLf & "This might be a duplicate posting." _
                                      & vbCrLf & vbCrLf & "Continue to Update anyway?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If

                Dim INV_AMT_MENS As Decimal = Val(dst.Tables("SPTCWRX4").Compute("SUM (CWRX_EXP_PCT)", "CHECKBOOK = 'MENS'") & "")
                Dim INV_AMT_WMEN As Decimal = Val(dst.Tables("SPTCWRX4").Compute("SUM (CWRX_EXP_PCT)", "CHECKBOOK = 'WMEN'") & "")
                If dst.Tables("SPTCWRX4").Select("").Length <> 0 Then
                    If INV_AMT_MENS <> 100 Then
                        EMsg &= vbCr & "CheckBook 'MENS' Exp Distribution " & INV_AMT_MENS & "% does not = 100%" & vbCr & "Please Correct in File Maintenance"
                    End If
                    If INV_AMT_WMEN <> 100 Then
                        EMsg &= vbCr & "CheckBook 'WMEN' Exp Distribution " & INV_AMT_WMEN & "% does not = 100%" & vbCr & "Please Correct in File Maintenance"
                    End If
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)
        Select Case eItemKey

            Case "Get Files"
                Get_Files()
            Case "Load Files"
                Load_Files()
                EntryMode = "L"
                Mode_Settings(True)
            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Get Files").Settings.Enabled = not_iScreenMode
                    .Items("Load Files").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode

                    .Items("Update").Visible = Not update_blocked
                End With

                With .Groups("Load Archived Docs")
                    .Expanded = False
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        'grdSPTCWRXF.Visible = Not ScreenMode
        'grdSPTARCHD.Visible = Not ScreenMode
        SplitContainer3.Visible = Not ScreenMode
        tabMain.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SPTCWRXF", "SPTARCHD", "SPTRSSP1",
             "SPTCWRX1", "SPTCWRX2", "SPTCWRX3",
             "SPTCWRX4", "SPTPYXI1", "SPTPYXI2",
             "APTINVH1", "APTINVH2", "ASTATTA2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
        If tblEXP IsNot Nothing Then
            tblEXP.Rows.Clear()
        End If

        List_Files()
    End Sub

    Sub Load_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Save_Header_Fields(UltraGroupBox1)
        EnforceConstraints(False)

        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating Database")

        BeginTrans()

        Update_Record_TDA("SPTRSSP1")
        Update_Record_TDA("SPTRSSP2")

        Update_Record_TDA("SPTCWRX1")
        Update_Record_TDA("SPTCWRX2")
        Update_Record_TDA("SPTCWRX3")
        Update_Record_TDA("SPTCWRX4", $"CTRL_NO = '{IMPORT_XNO}'")

        ASCMAIN1.sql = $"Insert into SPTCWRX5
            Select SPTCWRX2.CTRL_NO, SPTCWRX2.CUST_CODE, SPTCWRX2.CUST_STORE_NO, SPTCWRX2.CHECKBOOK
            , SPTCWRX2.OPS_YYYYPP, SPTCWRX2.OPS_YYYYWW, SPTCWRX2.BILL_AMT
            , SPTCWRXP.COLLECTION_CODE, SPTCWRX2.BILL_AMT * SPTCWRXP.CWRX_EXP_PCT/100 ACT
             from SPTCWRX2,SPTCWRXP
             where SPTCWRXP.CHECKBOOK = SPTCWRX2.CHECKBOOK AND SPTCWRX2.CTRL_NO = '{IMPORT_XNO}'"
        ASCDATA1.ExecuteSQL()

        Update_Record_TDA("SPTPYXI1")
        Update_Record_TDA("SPTPYXI2")

        If dst.Tables("APTINVH1").Rows.Count > 0 Then
            Update_Record_TDA("APTINVH1")
            Update_Record_TDA("APTINVH2")
            Update_Record_TDA("ASTATTA2")
        End If


        CommitTrans("Update Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Archiving Files")

        For Each FILENAME As String In FILENAMEs_to_Archive
            Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)

            If fi.Name.ToUpper.StartsWith("INV") Or fi.Name.ToUpper.StartsWith("PYX") Then
                Dim row() As DataRow = dst.Tables("ASTATTA2").Select("ATTACHMENT_FILENAME = '" & FILENAME & "'")
                Dim ATTACHMENT_NO As String = row(0).Item("ATTACHMENT_NO")
                Dim ATTACH_PATH As String = ASCMAIN1.Folders("Attach") & ATTACHMENT_NO
                My.Computer.FileSystem.CopyFile(FILENAME, ATTACH_PATH)
            End If

            Dim ARCHIVE_PATH As String = fi.DirectoryName & "\archive\" & fi.Name
            My.Computer.FileSystem.MoveFile(FILENAME, ARCHIVE_PATH)
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        List_Files()
    End Sub

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSPTCWRXF, "SSB", "Show Filter", "Show GroupBox", "Delete Files")
        Load_Popup_Menu(grdSPTRSSP2, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSPTCWRX2, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSPTPYXI2, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSPTCWRX4, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSPTARCHD, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If
        If Not GRDs.ContainsKey(Mid(e.SourceControl.Name, 4)) Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)

        Select Case e.SourceControl.Name
            Case "grdSATSLSC1"

            Case Else
        End Select
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Delete Files"

                If grdSPTCWRXF.Selected.Rows.Count = 0 AndAlso grdSPTCWRXF.ActiveRow IsNot Nothing Then
                    grdSPTCWRXF.ActiveRow.Selected = True
                End If

                Dim FILENAMEs_to_Delete As New List(Of String)
                For Each grow As UltraWinGrid.UltraGridRow In grdSPTCWRXF.Selected.Rows
                    FILENAMEs_to_Delete.Add(grow.Cells("FILENAME").Value)
                Next

                If FILENAMEs_to_Delete.Count = 0 Then
                    MsgBox("No Files Selected to Delete", MsgBoxStyle.OkOnly, "Cannot Delete")
                    Exit Sub
                End If

                If MsgBox("OK to Delete the " & CStr(FILENAMEs_to_Delete.Count) & " File(s) selected?" _
                          & vbCrLf & vbCrLf & "Remember - the Invoice file and the Payroll Export files need" _
                          & vbCrLf & " to be processed as a correlated pair",
                          MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Deleting Files")

                For Each FILENAME As String In FILENAMEs_to_Delete
                    Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
                    Dim sfx As String = $"_DELETED_{ASCMAIN1.USER_ID}_{Format(Now, "yyyyMMddHHmmss")}"
                    Dim DELETED_PATH As String = fi.DirectoryName & "\deleted\" & fi.Name & sfx
                    My.Computer.FileSystem.MoveFile(FILENAME, DELETED_PATH)
                Next

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

                MsgBox("Files have been Deleted", MsgBoxStyle.OkOnly, "Verification")
                List_Files()

        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"
    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            'Case "BRAND_CODE"
            '    If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
            '        If Absx1.txtFor("OPS_YYYYPP").Text <> "" Then
            '            Click_Command("Load", e)
            '        End If
            '    End If
        End Select
    End Sub
#End Region
    Sub Get_Files()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now getting files via sftp")

        Dim FOLDERNAME As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT") & "\" & "COWORX" & "\" & "FromCoworx" & "\"
        Dim FILENAMEs As List(Of String)
        FILENAMEs = TAC.TACSCOM1.sftp_get(Me, "COWORX", True, FOLDERNAME, "")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        List_Files()

        MsgBox(CStr(FILENAMEs.Count) & " Files Retrieved", MsgBoxStyle.OkOnly, "Verification")

    End Sub

    Sub Load_Files()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data from files")

        dst.Tables("SPTRSSP2").Rows.Clear()
        Dim files As Integer = 0
        FILENAMEs_to_Archive.Clear()
        update_blocked = False
        import_errors.Clear()

        ' need to do RSCs first before doing INVs

        For Each PFX As String In New String() {"RSC", "INV", "PYX"}

            Dim FILENAME_old As String = ""

            For Each row As DataRow In dst.Tables("SPTCWRXF").Select("", "FILEDATE")
                Dim FILENAME As String = row.Item("FILENAME")
                Dim FILEABBR As String = row.Item("FILEABBR")

                Dim FILENAME_fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)

                If FILEABBR.ToUpper.StartsWith(PFX) Then
                    FILENAMEs_to_Archive.Add(FILENAME)

                    If FILENAME < FILENAME_old Then
                        MsgBox("Filenames appear to be out of sequence - please call ABS - DO NOT UPDATE")
                        ' THE FILEDATE SORT DID NOT PRESENT THE FILES IN CHRONOLOGICAL ORDER
                    End If
                    FILENAME_old = FILENAME

                    files += 1
                    IMPORT_XNO = ASCMAIN1.Next_Control_No("SPTCWRXG.IMPORT_XNO")



                    ' USING THIS METHOD WOULD REQUIRE CHANGES TO EACH PROCESS 
                    'Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(FILENAME)
                    '    MyReader.TextFieldType = FileIO.FieldType.Delimited
                    '    MyReader.SetDelimiters(",")
                    '    Dim T As String()
                    '    While Not MyReader.EndOfData
                    '        Try
                    '            T = MyReader.ReadFields()

                    '        Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    '            Throw New Exception("Line " & ex.Message & "is not valid and will be skipped.")
                    '        End Try
                    '    End While
                    'End Using

                    Using sr As New System.IO.StreamReader(FILENAME)
                        Dim T() As String = Split(sr.ReadToEnd, vbCrLf)

                        Select Case Mid(FILEABBR.ToUpper, 1, 3)
                            Case "RSC"
                                Process_RSC(T)
                            Case "INV"
                                Process_INV(T)
                            Case "PYX"
                                Process_PYX(T)
                            Case Else
                                MsgBox("Unknown File Type (" & FILENAME & ")", MsgBoxStyle.OkOnly, "Please take Screenshot and send to ABS")
                        End Select

                        If PFX = "INV" Or PFX = "PYX" Then
                            Dim rowASTATTA2 As DataRow = dst.Tables("ASTATTA2").NewRow
                            With rowASTATTA2
                                .Item("TABLE_NAME") = "APTINVH1"
                                .Item("COLUMN_NAME") = "VOUCHER_NO"
                                .Item("CODE_VALUE") = "X"
                                Dim ATTACHMENT_NO As String = ASCMAIN1.Next_Control_No("ASTATTA2.ATTACHMENT_NO")
                                .Item("ATTACHMENT_NO") = ATTACHMENT_NO
                                .Item("ATTACHMENT_DESC") = IIf(PFX = "INV", "CoWorx Invoice File", "CoWorx Payroll Export File")
                                .Item("ATTACHMENT_FILENAME") = FILENAME
                                .Item("ATTACHMENT_EXT") = "CSV"
                                .Item("COMPUTER_NAME") = My.Computer.Name
                                .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                                .Item("INIT_DATE") = DATETIME_STAMP
                                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                                .Item("ATTACHMENT_TYPE") = "CWX"
                                '.Item("ATTACHMENT_ORIGINATOR") = ""
                                .Item("ATTACHMENT_DATETIME") = FILENAME_fi.LastWriteTime
                                '.Item("ATTACHMENT_STATUS") = ""
                                '.Item("ATTACHMENT_NOTES") = ""
                            End With
                            dst.Tables("ASTATTA2").Rows.Add(rowASTATTA2)
                        End If

                        sr.Close()
                        sr.Dispose()
                    End Using
                End If
            Next

            If PFX = "RSC" Then
                Apply_RSC_Changes()
            End If
        Next


        Create_AP()
        If import_errors.Count <> 0 Then
            update_blocked = True
            Dim import_errors_reported As New List(Of String)
            For i As Integer = 0 To import_errors.Count - 1
                import_errors_reported.Add(import_errors(i))
                If i > 10 Then
                    import_errors_reported.Add("more ...")
                    Exit For
                End If
            Next

            MsgBox("Import Errors (" & import_errors.Count & "):" & vbCrLf & vbCrLf & Join(import_errors_reported.ToArray, vbCrLf), vbOKOnly, "Cannot Process File")
        End If

        MsgBox(CStr(files) & " Files have been Loaded", MsgBoxStyle.OkOnly, "Verification")
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        If import_errors.Count = 0 Then
            Dim OPS_YYYYWW_max As String = dst.Tables("SPTCWRX2").Compute("MAX(OPS_YYYYWW)", "") & ""

            If OPS_YYYYWW_max >= ASCMAIN1.CYW Then
                grdSPTCWRX2.DisplayLayout.GroupByBox.Hidden = False
                grdSPTCWRX2.DisplayLayout.Bands(0).SortedColumns.Clear()
                grdSPTCWRX2.DisplayLayout.Bands(0).SortedColumns.Add("OPS_YYYYWW", False, True)
                tabMain.SelectedTab = tabMain.Tabs("Invoices")

                For Each grow As UltraWinGrid.UltraGridRow In grdSPTCWRX2.Rows
                    If grow.IsGroupByRow Then
                        Dim growgby As UltraWinGrid.UltraGridGroupByRow = DirectCast(grow, UltraWinGrid.UltraGridGroupByRow)
                        If growgby.Value >= ASCMAIN1.CYW Then
                            growgby.Expanded = True
                        End If
                        '     Stop
                    End If
                Next

                MsgBox("There are records in the is file that appear to be later than last week", MsgBoxStyle.OkOnly, "Please Check")
            End If
        End If
    End Sub

    Sub List_Files()

        dst.Tables("SPTCWRXF").Rows.Clear()

        ' WHILE TESTING
        If ASCMAIN1.Running_in_VS Then ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT") = "C:\Users\wjz\Desktop\Interparfums" ' "C:\Users\nicholas\Desktop" ' \coworx\FromCoworx"

        Dim FOLDERNAME As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT") & "\" & "COWORX" & "\" & "FromCoworx"
        For Each FILENAME As String In My.Computer.FileSystem.GetFiles(FOLDERNAME)
            Dim FI As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
            dst.Tables("SPTCWRXF").Rows.Add(New Object() {FI.FullName, FI.CreationTime, FI.Length, FI.Name})
        Next
        Sort_grdColumns(grdSPTCWRXF, "FILENAME")
    End Sub

    Sub Process_RSC(T() As String)

        For r As Integer = 0 To T.Length - 1
            Dim line As String = T(r)

            If r <> 0 Then
                If line = "" Then
                    '  r = r - 1
                Else
                    Dim rowSPTRSSP2 As DataRow = dst.Tables("SPTRSSP2").NewRow
                    Dim data() As String = Split(line, ",")
                    Dim i As Integer = -1
                    For Each COLUMN_NAME As String In COLUMN_NAMEs
                        i += 1
                        If COLUMN_NAME = "SELL_NAME" Then
                        Else
                            If data(i) <> "" Then

                                If dst.Tables("SPTRSSP2").Columns(COLUMN_NAME).DataType.ToString = "System.DateTime" Then
                                    If Trim(data(i)) <> "" Then
                                        rowSPTRSSP2.Item(COLUMN_NAME) = data(i)
                                    End If

                                Else

                                    If COLUMN_NAME = "RSSP_PHONE" Or COLUMN_NAME = "RSSP_CELL" Then
                                        data(i) = Trim(Replace(data(i), "-", ""))
                                    End If
                                    'If COLUMN_NAME = "RSSP_PAY_RATE" Then ' waiting for catherine
                                    '    rowSPTRSSP2.Item(COLUMN_NAME) = Val(data(i))
                                    'Else
                                    If COLUMN_NAME = "RSSP_NAME_LAST" Or COLUMN_NAME = "RSSP_NAME_FIRST" Then
                                        data(i) = Trim(data(i))
                                    End If
                                    rowSPTRSSP2.Item(COLUMN_NAME) = data(i)
                                    'End If
                                End If

                            End If
                        End If
                    Next

                    rowSPTRSSP2.Item("IMPORT_XNO") = IMPORT_XNO
                    rowSPTRSSP2.Item("IMPORT_DATE") = DATETIME_STAMP

                    dst.Tables("SPTRSSP2").Rows.Add(rowSPTRSSP2)
                End If
            End If
        Next
    End Sub

    Sub Apply_RSC_Changes()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now loading RSC file")

        Fill_Records("SPTRSSP1")

        For Each row As DataRow In ASCDATA1.SelectDistinct("SPTRSSP2", New String() {"IMPORT_XNO"}).Select("", "IMPORT_XNO")
            Dim IMPORT_XNO As String = row.Item("IMPORT_XNO")

            Dim RSSP_IDs As New List(Of String)

            For Each rowSPTRSSP2 As DataRow In dst.Tables("SPTRSSP2").Select("IMPORT_XNO = '" & IMPORT_XNO & "'")
                Dim RSSP_ID As String = rowSPTRSSP2.Item("RSSP_ID")
                RSSP_IDs.Add(RSSP_ID)

                Dim sql As String = "RSSP_ID = '" & RSSP_ID & "'"
                Dim rowSPTRSSP1s() As DataRow = dst.Tables("SPTRSSP1").Select(sql)
                Dim rowSPTRSSP1 As DataRow = Nothing
                If rowSPTRSSP1s.Length = 0 Then
                    rowSPTRSSP1 = dst.Tables("SPTRSSP1").NewRow
                    rowSPTRSSP1.Item("RSSP_CODE") = ASCMAIN1.Next_Control_No("SPTRSSP1.RSSP_CODE")
                    dst.Tables("SPTRSSP1").Rows.Add(rowSPTRSSP1)
                Else
                    rowSPTRSSP1 = rowSPTRSSP1s(0)
                End If

                Dim RECORD_TYPE As String = rowSPTRSSP2.Item("RECORD_TYPE") & ""

                For Each COLUMN_NAME As String In COLUMN_NAMEs
                    If COLUMN_NAME = "RECORD_TYPE" Or COLUMN_NAME = "SELL_NAME" Then
                    Else
                        If Me.replacement_method Then
                            ' WHEN USING REPLACEMENT METHOD - JUST TAKE ALL OF THE DATE INFO PROVIDED
                            rowSPTRSSP1.Item(COLUMN_NAME) = rowSPTRSSP2.Item(COLUMN_NAME)

                        Else
                            If (COLUMN_NAME = "RSSP_DATE_HIRED" And RECORD_TYPE <> "N") _
                            Or (COLUMN_NAME = "RSSP_DATE_CHANGED" And RECORD_TYPE <> "C") _
                            Or (COLUMN_NAME = "RSSP_DATE_TERM" And RECORD_TYPE <> "T") Then
                                ' DO NOTHING
                            Else

                                rowSPTRSSP1.Item(COLUMN_NAME) = rowSPTRSSP2.Item(COLUMN_NAME)

                            End If
                        End If
                    End If
                Next

                ' IN REPLACEMENT MODE - THIS MIGHT BE DONE USING A STATUS FIELD RATHER THAN AN ACTION FIELD
                If RECORD_TYPE = "N" Or RECORD_TYPE = "C" Then
                    rowSPTRSSP1.Item("RSSP_DATE_TERM") = DBNull.Value
                End If

                Dim RSSP_NAME As String = rowSPTRSSP2.Item("RSSP_NAME_FIRST") & " " & rowSPTRSSP2.Item("RSSP_NAME_LAST")

                rowSPTRSSP1.Item("RSSP_NAME") = RSSP_NAME
                rowSPTRSSP1.Item("RSSP_SHIP_TO_NAME") = RSSP_NAME
                rowSPTRSSP1.Item("IMPORT_XNO") = IMPORT_XNO
                rowSPTRSSP1.Item("IMPORT_DATE") = DATETIME_STAMP

                Dim RSSP_TITLE As String = Trim(rowSPTRSSP1.Item("RSSP_TITLE") & "")
                Dim RSSP_TYPE As String = ""
                Select Case RSSP_TITLE
                    Case "RSC"
                        RSSP_TYPE = "C"
                    Case "FM"
                        RSSP_TYPE = "F"
                    Case "SDS"
                        RSSP_TYPE = "D"
                    Case Else
                        RSSP_TYPE = "?"
                End Select
                rowSPTRSSP1.Item("RSSP_TYPE") = RSSP_TYPE

                If Me.replacement_method Then
                    ' i don't think that this should be happening because on 06/23/21 LBM wants RSSP_STATUS to be A or I depending on RSSP_DATE_TERM
                    ' rowSPTRSSP1.Item("RSSP_STATUS") = RECORD_TYPE
                End If


                If rowSPTRSSP1.Item("RSSP_DATE_TERM") & "" = "" Then
                    rowSPTRSSP1.Item("RSSP_STATUS") = "A"
                Else
                    rowSPTRSSP1.Item("RSSP_STATUS") = "I"
                End If
            Next

            If Me.replacement_method Then
                Dim RSSP_IDs_to_Terminate As New Dictionary(Of String, String)
                Stop ' NOT SURE WHAT WAS INTENDED HERE - BUT IT LOOKS LIKE replacement_method NEEDS A LITTLE WORK
                For Each rowSPTRSSP1 As DataRow In dst.Tables("SPTRSSP1").Select("ISNULL(RSSP_STATUS,'?') <> 'X'")
                    Dim RSSP_ID As String = rowSPTRSSP1.Item("RSSP_ID")
                    Dim RSSP_CODE As String = rowSPTRSSP1.Item("RSSP_CODE")
                    Dim RSSP_STATUS As String = rowSPTRSSP1.Item("RSSP_STATUS") & ""
                    If Not RSSP_IDs.Contains(RSSP_ID) Then
                        RSSP_IDs_to_Terminate.Add(RSSP_ID, RSSP_CODE)
                        ' rowSPTRSSP1.Item("") = ""
                    End If
                Next
                If RSSP_IDs_to_Terminate.Count <> 0 Then
                    If MsgBox("OK to Terminate " & CStr(RSSP_IDs_to_Terminate.Count) & " RSCs not Included in File?",
                              MsgBoxStyle.OkCancel,
                              "You are now in Replacement Mode") = MsgBoxResult.Ok Then

                        For Each RSSP_ID As String In RSSP_IDs_to_Terminate.Keys
                            Dim RSSP_CODE As String = RSSP_IDs_to_Terminate(RSSP_ID)
                            Dim rowSPTRSSP1 As DataRow = dst.Tables("SPTRSSP1").Rows.Find(RSSP_CODE)
                            rowSPTRSSP1.Item("RSSP_DATE_TERM") = DATETIME_STAMP.Date
                            'rowSPTRSSP1.Item("RSSP_STATUS") = "X"
                            rowSPTRSSP1.Item("RSSP_STATUS") = "I" ' SEE LBM 06/23/21
                        Next
                    End If
                End If
            End If
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Process_PYX(T() As String)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now loading PYX file")

        Dim PYX_HEADER As String = "SPTPYXI1"
        Dim PYX_DETAIL As String = "SPTPYXI2"
        Dim CTRL_NO As String = IMPORT_XNO
        Dim records As Integer = 0
        Dim YWs As New Dictionary(Of String, String)
        Dim Sub_Brands(0) As String

        Dim Sub_Brands_count As Integer = 0

        ' Dim dataFromStream() As String

        For r As Integer = 0 To T.Length - 1

            Dim line As String = T(r)
            Try
                If r <> 0 Then ' If r <> -1 Then ' switch back to 0 because the file now has 2 header lines, the 1st being a Report title with a date range
                    ' also - T.length - X needs to be adjusted
                    If line = "" Or line.StartsWith(",,") Then
                    Else
                        records += 1

                        Dim data() As String = Split(line, ",")


                        Dim lineStream As System.IO.Stream = New System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(line))
                        Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(lineStream)
                            MyReader.TextFieldType = FileIO.FieldType.Delimited
                            MyReader.SetDelimiters(",")
                            While Not MyReader.EndOfData
                                Try
                                    data = MyReader.ReadFields()
                                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                                    Throw New Exception("Line " & ex.Message & "is not valid and will be skipped.")
                                End Try
                            End While
                        End Using


                        If r = 1 Then ' THIS MUST CHANGE IF NUMBER OF HEADER LINES CHANGES
                            ReDim Sub_Brands(data.Length - 15)
                            Sub_Brands_count = 0
                            For i As Integer = 0 To Sub_Brands.Count - 1
                                Sub_Brands(i) = data(i + 13)
                                Sub_Brands_count += 1
                                If data(i + 13) = "Total Retail Sales" Then
                                    Exit For
                                End If
                            Next
                        Else 'If r = 2 Then

                            'ReDim Sub_Brands(data.Length - 15)
                            'For i As Integer = 0 To Sub_Brands.Count - 1
                            '    Sub_Brands(i) = data(i + 13)
                            'Next
                            '   Else


                            Dim WORK_WE_DATE As Date = CDate(data(5))
                            Dim YYYYMMDD As String = Format(WORK_WE_DATE, "yyyyMMdd")

                            Dim YW As String
                            If YWs.ContainsKey(YYYYMMDD) Then
                                YW = YWs(YYYYMMDD)
                            Else
                                ASCMAIN1.sql = "SELECT MIN (YYYYWW) FROM GLTPARM3 WHERE WEEK_END_DATE >= '" & Format(WORK_WE_DATE, "dd-MMM-yyyy") & "'"
                                YW = ASCDATA1.GetDataValue
                                YWs.Add(YYYYMMDD, YW)
                            End If

                            Dim rowGLTPARM3 As DataRow = dst.Tables("GLTPARM3").Rows.Find(YW)


                            Dim DOOR_CODE As String = data(2)
                            Dim CS As String = Split(DOOR_CODE, "-")(4)


                            Dim CUST_STORE_NO = Mid(CS, 4)
                            Dim CUST_CODE As String = ""
                            If Mid(CS, 1, 3) = "SDS" Then
                                CUST_CODE = "SDS"
                            Else
                                Dim rowSPTCWRXA As DataRow = dst.Tables("SPTCWRXA").Rows.Find(Mid(CS, 1, 3))
                                CUST_CODE = rowSPTCWRXA.Item("CUST_CODE")

                                Dim rowARTCUST2 As DataRow = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
                                If rowARTCUST2 Is Nothing Then
                                    Dim rowARTCUSTT As DataRow = dst.Tables("ARTCUSTT").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
                                    If rowARTCUSTT IsNot Nothing Then
                                        Dim NEW_CUST_CODE As String = rowARTCUSTT.Item("NEW_CUST_CODE") & String.Empty
                                        Dim NEW_CUST_STORE_NO As String = rowARTCUSTT.Item("NEW_CUST_STORE_NO") & String.Empty
                                        rowARTCUST2 = dst.Tables("ARTCUST2").Rows.Find(New String() {NEW_CUST_CODE, NEW_CUST_STORE_NO})
                                    End If

                                    If rowARTCUST2 Is Nothing Then
                                        update_blocked = True
                                        If MsgBox("Bad Customer-Store: " & CUST_CODE & "-" & CUST_STORE_NO, MsgBoxStyle.OkCancel, "Update will not be permitted") = MsgBoxResult.Cancel Then
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If

                            Dim rowPYX_HEADER As DataRow = dst.Tables(PYX_HEADER).NewRow
                            rowPYX_HEADER.Item("CTRL_NO") = CTRL_NO
                            rowPYX_HEADER.Item("CTRL_LNO") = r
                            rowPYX_HEADER.Item("CUST_STORE_NAME") = data(0)
                            rowPYX_HEADER.Item("RSC_NAME") = data(1)
                            rowPYX_HEADER.Item("DOOR_CODE") = data(2)
                            rowPYX_HEADER.Item("CHECKBOOK") = data(3)
                            rowPYX_HEADER.Item("COWORX_ID") = data(4)
                            rowPYX_HEADER.Item("WORK_DATE") = data(5)
                            rowPYX_HEADER.Item("START_TIME") = data(6)
                            rowPYX_HEADER.Item("END_TIME") = data(7)
                            rowPYX_HEADER.Item("PAY_TYPE") = data(8)
                            rowPYX_HEADER.Item("PAY_HOURS") = data(9)
                            rowPYX_HEADER.Item("PAY_RATE") = data(10)
                            rowPYX_HEADER.Item("CUST_STORE_STATE") = data(11)
                            rowPYX_HEADER.Item("SPEND_AMT") = data(12)
                            rowPYX_HEADER.Item("TOT_SALES_AMT") = data(Sub_Brands_count + 13 - 1)

                            rowPYX_HEADER.Item("CUST_CODE") = CUST_CODE
                            rowPYX_HEADER.Item("CUST_STORE_NO") = CUST_STORE_NO
                            rowPYX_HEADER.Item("OPS_YYYYPP") = rowGLTPARM3.Item("YYYYPP")
                            rowPYX_HEADER.Item("OPS_YYYYWW") = YW

                            dst.Tables(PYX_HEADER).Rows.Add(rowPYX_HEADER)


                            For i As Integer = 0 To Sub_Brands_count - 1 - 1 'Sub_Brands.Count - 1
                                If Val(data(i + 13)) <> 0 Then
                                    Dim rowPYX_DETAIL As DataRow = dst.Tables(PYX_DETAIL).NewRow
                                    rowPYX_DETAIL.Item("CTRL_NO") = CTRL_NO
                                    rowPYX_DETAIL.Item("CTRL_LNO") = r
                                    rowPYX_DETAIL.Item("SUB_BRAND") = Sub_Brands(i)
                                    rowPYX_DETAIL.Item("SALES_AMT") = data(i + 13)
                                    dst.Tables(PYX_DETAIL).Rows.Add(rowPYX_DETAIL)
                                End If

                            Next
                        End If
                    End If
                End If
            Catch ex As Exception
                import_errors.Add(ex.Message & ":" & "Error processing Invoice file line " & CStr(r + 1) & vbCrLf & line)
                '  MsgBox(ex.Message & vbCrLf & vbCrLf & "Error processing Invoice file line " & line)
            End Try

        Next

        'Dim rowSPTCWRX1 As DataRow = dst.Tables("SPTCWRX1").NewRow
        'rowSPTCWRX1.Item("CTRL_NO") = CTRL_NO
        'rowSPTCWRX1.Item("DATE_PROCESSED") = DATETIME_STAMP
        'rowSPTCWRX1.Item("RECORD_COUNT") = records
        'dst.Tables("SPTCWRX1").Rows.Add(rowSPTCWRX1)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Process_INV(T() As String)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now loading INV file")

        Dim CTRL_NO As String = IMPORT_XNO
        Dim records As Integer = 0
        Dim YWs As New Dictionary(Of String, String)

        Dim INV_NUM As String = ""
        Dim INV_DATE As Date
        Dim INV_AMT As Decimal = 0

        Dim RSCPAY As Decimal = 0
        Dim RSCFRG As Decimal = 0
        Dim RSCEXP As Decimal = 0
        Dim SDSPAY As Decimal = 0
        Dim SDSFRG As Decimal = 0
        Dim SDSEXP As Decimal = 0

        EMsg = ""
        YW_INVOICE = ""

        Fill_Records("SPTCWRX4", String.Empty, True, "SELECT '" & CTRL_NO & "' CTRL_NO,SPTCWRXP.*,0 CWRX_EXP_AMT FROM SPTCWRXP")
        With grdSPTCWRX4.DisplayLayout.Bands(0)
            .Columns(0).HiddenWhenGroupBy = DefaultableBoolean.True
            .SortedColumns.Add("CHECKBOOK", False, True)
        End With

        '   MsgBox("ABS MUST CHECK TO SEE IF THE FILE CONTAINS HEADERS" & vbCrLf & vbCrLf & "Header Row: " & T(0), MsgBoxStyle.OkOnly, "PLEASE CONTACT ABS BEFORE PROCEEDING")

        Dim OPS_YYYYWWs As New Dictionary(Of String, Decimal)

        For r As Integer = 0 To T.Length - 1
            Dim line As String = T(r)
            If r <> 0 Then ' use If r <> 0 if the file contains a header.  USE IF R <> -1 OF THE FILE DOES NOT CONTAIN A HEADER

                If line = "" Then
                Else
                    records += 1

                    'Invoice Date	Company	ID#	        Job Title	Account	Door #	State	AE	Brand/House	
                    '20160210	    IPLB	020101AB3	RSC	        MACYS   000307	MI	    330	JCM	        
                    'Regular Wage Rate	Regular Hours Worked	Overtime Hours Worked	Other Hours/Units	Wage Code Description	
                    '21	                4	                    0	                    0		                                    
                    'Check Book Reported Sales	Other Reported Sales	Work Weekending	
                    '610	                    0	                    20160206	    
                    'Shift Wages	Retro Wages	Training Wages	Sick Pay Wages	Bonus	Expense	Travel Expense	
                    '84	            0	        0	            0	            0	    0	    0	            
                    'Other taxes ($)	Sales Tax ($)	CoWorx fee ($)	Invoice Amount	Fee %
                    '0	                0	            18.28	        102.28	        21.76

                    '' 0    Invoice Date	20160210
                    '' 1    Company		    IPLB
                    '' 2    ID#		        ABCDEFGG8
                    '' 3 	Title	        RSC
                    '' 4    Account		    MACYS
                    '' 5    Door #		    000307
                    '' 6    State	        MI
                    '' 7    AE	            330
                    '' 8    Brand/House	    JCM         Checkbook
                    '' 9	Regular Wage Rate	    21
                    ''10    Regular Hours Worked    4
                    ''11    Overtime Hours Worked   0
                    ''12    Other Hours/Units       0
                    ' 13	Wage Code Description   
                    ''14	Check Book Reported Sales	610
                    ' 15	Other Reported Sales	    0.00
                    ''16	Work Weekending	        20160206
                    ''17    Shift Wages	        0
                    ''18    Retro Wages	        0
                    ''19    Training Wages	    0
                    ''20    Sick Pay Wages	    0
                    ''21    Bonus	            0
                    ''22    Expense	            0
                    ''23    Travel Expense      0
                    ''24    Other taxes ($)     0
                    ''25	Sales Tax ($)	    0
                    ''26    CoWorx fee ($)      18.28
                    ''27	Invoice Amount	    102.28
                    ''28    Fee %               21.76

                    ' how to calculate pay based on overtime
                    ' does total bill include sales tax?
                    ' I have not used Other Sales nor Hours/Unit Type in my data load
                    ' there is a mismatch between test data and the layout - Hours/Unit Type

                    Dim data() As String = Split(line, ",")

                    Dim YYYYMMDD As String = data(0)
                    INV_DATE = CDate(Mid(YYYYMMDD, 5, 2) & "/" & Mid(YYYYMMDD, 7, 2) & "/" & Mid(YYYYMMDD, 1, 4))
                    YYYYMMDD = data(16)
                    Dim WORK_WE_DATE As Date = CDate(Mid(YYYYMMDD, 5, 2) & "/" & Mid(YYYYMMDD, 7, 2) & "/" & Mid(YYYYMMDD, 1, 4))

                    Dim YW As String
                    If YWs.ContainsKey(YYYYMMDD) Then
                        YW = YWs(YYYYMMDD)
                    Else
                        ASCMAIN1.sql = "SELECT MIN (YYYYWW) FROM GLTPARM3 WHERE WEEK_END_DATE >= '" & Format(WORK_WE_DATE, "dd-MMM-yyyy") & "'"
                        YW = ASCDATA1.GetDataValue
                        YWs.Add(YYYYMMDD, YW)
                    End If

                    If YW_INVOICE = "" Then
                        Dim POSTING_DATE As Date = INV_DATE.AddDays(-7)
                        ASCMAIN1.sql = "SELECT MIN (YYYYWW) FROM GLTPARM3 WHERE WEEK_END_DATE >= '" & Format(POSTING_DATE, "dd-MMM-yyyy") & "'"
                        YW_INVOICE = ASCDATA1.GetDataValue
                    End If

                    Dim rowGLTPARM3 As DataRow = dst.Tables("GLTPARM3").Rows.Find(YW)

                    Dim CUST_CODE As String = Trim(data(4))
                    Dim CUST_STORE_NO As String = Trim(data(5))
                    If CUST_STORE_NO.Length < 6 Then
                        CUST_STORE_NO = CUST_STORE_NO.PadLeft(6, "0")
                    End If

                    If CUST_CODE = "NEIMAN MARCUS" Then CUST_CODE = "NEIMAN"

                    If CUST_CODE = "SDS" Then
                        ' SDS WON'T COME IN WITH A REAL CUSTOMER
                        ' BUT SDS JOB TITLE MIGHT COME IN WITH A REAL CUSTOMER- SEE CM EMAIL 01/10/19
                        'In some cases this past year, an SDS worked RSC hours for a specific account and brand.  Below is an example of an invoice with this situation.  Right now it seems that the “Job Title” part of the invoice is being used to determine if shift wages should be applied to the SDS line for Finance.  The true SDS hours are coded under the customer; however our system is capturing the additional hours (like the 11 below) as time worked under the SDS budget rather than under freelance.  
                        'We’d like to update the system to pull the “Customer” line rather than the “Job Title” line so the wages hit the correct budgets going forward.  
                        'The additional RSC hours being worked are correctly hitting our freelance budget (spend in the workbooks), just hitting the wrong line for Finance. 
                    Else
                        Dim rowARTCUST2 As DataRow = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
                        If rowARTCUST2 Is Nothing Then
                            Dim rowARTCUSTT As DataRow = dst.Tables("ARTCUSTT").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
                            If rowARTCUSTT IsNot Nothing Then
                                Dim NEW_CUST_CODE As String = rowARTCUSTT.Item("NEW_CUST_CODE") & String.Empty
                                Dim NEW_CUST_STORE_NO As String = rowARTCUSTT.Item("NEW_CUST_STORE_NO") & String.Empty
                                rowARTCUST2 = dst.Tables("ARTCUST2").Rows.Find(New String() {NEW_CUST_CODE, NEW_CUST_STORE_NO})
                            End If

                            If rowARTCUST2 Is Nothing Then
                                update_blocked = True
                                If MsgBox("Bad Customer-Store: " & CUST_CODE & "-" & CUST_STORE_NO, MsgBoxStyle.OkCancel, "Update will not be permitted") = MsgBoxResult.Cancel Then
                                    Exit For
                                End If
                            End If
                        End If
                    End If


                    Dim CHECKBOOK As String = Trim(data(8))
                    Dim COLLECTION_CODE As String = ""
                    If CHECKBOOK = "SDS" Then
                        COLLECTION_CODE = "000"
                    Else
                        Dim rowSPTCWRXC As DataRow = dst.Tables("SPTCWRXC").Rows.Find(CHECKBOOK)
                        COLLECTION_CODE = rowSPTCWRXC.Item("COLLECTION_CODE") & ""
                    End If

                    Dim rowSPTRSSP1s() As DataRow = dst.Tables("SPTRSSP1").Select("RSSP_ID = '" & data(2) & "'")
                    Dim rowSPTRSSP1 As DataRow = Nothing
                    If rowSPTRSSP1s.Length = 1 Then
                        rowSPTRSSP1 = rowSPTRSSP1s(0)
                    End If

                    Dim rowSPTCWRX2 As DataRow = dst.Tables("SPTCWRX2").NewRow

                    rowSPTCWRX2.Item("CTRL_NO") = CTRL_NO
                    rowSPTCWRX2.Item("CTRL_LNO") = r
                    rowSPTCWRX2.Item("SSN") = data(2)
                    If rowSPTRSSP1 IsNot Nothing Then
                        rowSPTCWRX2.Item("EMP_LASTNAME") = rowSPTRSSP1.Item("RSSP_NAME_LAST")
                        rowSPTCWRX2.Item("EMP_FIRSTNAME") = rowSPTRSSP1.Item("RSSP_NAME_FIRST")
                    End If
                    rowSPTCWRX2.Item("COLLECTION_CODE") = COLLECTION_CODE
                    rowSPTCWRX2.Item("TERRITORY") = Trim(data(7))
                    rowSPTCWRX2.Item("CUST_CODE") = CUST_CODE
                    rowSPTCWRX2.Item("CUST_STORE_NO") = CUST_STORE_NO
                    rowSPTCWRX2.Item("PAY_RATE") = Val(data(9))
                    rowSPTCWRX2.Item("PAY_HOURS") = Val(data(10)) + Val(data(11))
                    rowSPTCWRX2.Item("BILL_RATE") = Val(data(9))
                    rowSPTCWRX2.Item("BILL_HOURS") = Val(data(10)) + Val(data(11))
                    'rowSPTCWRX2.Item("CHECK_NO") = data(i)
                    'rowSPTCWRX2.Item("CHECK_DATE") = data(i)
                    rowSPTCWRX2.Item("WORK_DATE") = WORK_WE_DATE
                    rowSPTCWRX2.Item("INVOICE_DATE") = INV_DATE
                    rowSPTCWRX2.Item("BILL_AMT") = Val(data(17)) + Val(data(18)) + Val(data(19)) + Val(data(20))
                    'rowSPTCWRX2.Item("DISCOUNT_AMT") = data(i)
                    rowSPTCWRX2.Item("SALES_TAX_AMT") = Val(data(25))
                    rowSPTCWRX2.Item("NET_AMT") = Val(data(27))

                    Dim OPS_YYYYWW As String = rowGLTPARM3.Item("YYYYWW")
                    If Not OPS_YYYYWWs.ContainsKey(OPS_YYYYWW) Then
                        OPS_YYYYWWs.Add(OPS_YYYYWW, 0)
                    End If

                    OPS_YYYYWWs(OPS_YYYYWW) += Val(data(27))

                    rowSPTCWRX2.Item("CONTROL_NO") = CTRL_NO
                    rowSPTCWRX2.Item("OPS_YYYYPP") = rowGLTPARM3.Item("YYYYPP") ' Format(INV_DATE, "yyyyMM")
                    'rowSPTCWRX2.Item("CUST_STORE_SUBMIT") = data(i)
                    rowSPTCWRX2.Item("RETAIL_VALUE_SOLD") = Val(data(14)) + Val(data(15))
                    rowSPTCWRX2.Item("OPS_YYYYWW") = YW

                    rowSPTCWRX2.Item("JOB_TITLE") = Trim(data(3))
                    rowSPTCWRX2.Item("STATE") = Trim(data(6))
                    rowSPTCWRX2.Item("CHECKBOOK") = CHECKBOOK
                    rowSPTCWRX2.Item("HOURS_REG") = Val(data(10))
                    rowSPTCWRX2.Item("HOURS_OVT") = Val(data(11))
                    rowSPTCWRX2.Item("HOURS_OTH") = Val(data(12))
                    rowSPTCWRX2.Item("WAGE_CODE_DESC") = Trim(data(13))
                    rowSPTCWRX2.Item("SALES_CHECKBOOK") = Val(data(14))
                    rowSPTCWRX2.Item("SALES_OTHER") = Val(data(15))
                    rowSPTCWRX2.Item("WAGES_SHIFT") = Val(data(17))
                    rowSPTCWRX2.Item("WAGES_RETRO") = Val(data(18))
                    rowSPTCWRX2.Item("WAGES_TRAIN") = Val(data(19))
                    rowSPTCWRX2.Item("WAGES_SICK") = Val(data(20))
                    rowSPTCWRX2.Item("WAGES_BONUS") = Val(data(21))
                    rowSPTCWRX2.Item("WAGES_EXP") = Val(data(22))
                    rowSPTCWRX2.Item("TRAVEL_EXP") = Val(data(23))
                    rowSPTCWRX2.Item("OTHER_TAXES") = Val(data(24))
                    rowSPTCWRX2.Item("COWORX_FEE") = Val(data(26))
                    rowSPTCWRX2.Item("COWORX_FEE_PCT") = Val(data(28))

                    INV_AMT += Val(data(27))

                    With rowSPTCWRX2
                        'LR email 02/10/2016
                        'Spoke with Debbie and here is the breakdown (color coded with the spreadsheet):

                        'SDS:
                        'Wages (green):  641000 (SDS SALARY)
                        'Fringe and tax (purple):  642000 (SDS FRINGE)
                        'Expenses (blue):  643000 (SDS T&E)

                        'RSC:
                        'Wages (green):  632000 (INDEPENDENT PROMOTERS—RSC)
                        'Fringe and tax (purple):  633000 (INDEPENDENT PROMOTER—FRINGE)

                        ' NEW
                        Dim PAY As Decimal = 0
                        Dim FRG As Decimal = 0
                        Dim EXP As Decimal = 0
                        For Each rowSPTCWRX4 As DataRow In dst.Tables("SPTCWRX4").Select($"CHECKBOOK = '{CHECKBOOK}'")
                            If rowSPTCWRX4.Item("CHECKBOOK") = CHECKBOOK Then
                                'Update  PAY, FRG, EXP * Percentage, Write to 3 each SPTCWRX4 record
                                PAY = Val(.Item("WAGES_SHIFT")) + Val(.Item("WAGES_RETRO")) + Val(.Item("WAGES_TRAIN")) + Val(.Item("WAGES_SICK")) + Val(.Item("WAGES_BONUS"))
                                'Purple
                                FRG = Val(.Item("OTHER_TAXES")) + Val(.Item("SALES_TAX_AMT")) + Val(.Item("COWORX_FEE"))
                                'Blue
                                EXP = Val(.Item("WAGES_EXP")) + Val(.Item("TRAVEL_EXP"))
                                COLLECTION_CODE = rowSPTCWRX4.Item("COLLECTION_CODE")

                                If CHECKBOOK <> "SDS" Then
                                    PAY = System.Math.Round(Val((rowSPTCWRX4.Item("CWRX_EXP_PCT")) / 100) * PAY, 2)
                                    FRG = System.Math.Round(Val((rowSPTCWRX4.Item("CWRX_EXP_PCT")) / 100) * FRG, 2)
                                    EXP = System.Math.Round(Val((rowSPTCWRX4.Item("CWRX_EXP_PCT")) / 100) * EXP, 2)
                                End If

                                Dim rowSPTCWRX3 As DataRow = dst.Tables("SPTCWRX3").Rows.Find(New String() {CTRL_NO, COLLECTION_CODE})
                                If rowSPTCWRX3 Is Nothing Then
                                    rowSPTCWRX3 = dst.Tables("SPTCWRX3").NewRow
                                    rowSPTCWRX3.Item("CTRL_NO") = CTRL_NO
                                    rowSPTCWRX3.Item("COLLECTION_CODE") = COLLECTION_CODE
                                    dst.Tables("SPTCWRX3").Rows.Add(rowSPTCWRX3)
                                End If


                                If Trim(data(3)) = "RSC" Or Trim(data(3)) = "FM" Or (Trim(data(3)) = "SDS" And CUST_CODE <> "SDS") Then
                                    RSCPAY += PAY
                                    RSCFRG += FRG
                                    RSCEXP += EXP

                                    Dim RSCEXP_field As String = "RSCEXP"
                                    If EXP <> 0 Then
                                        If data(13) = "BYOD Expense Reimbursement" Then
                                            RSCEXP_field = "RSCFRG"
                                        Else
                                            import_errors.Add("Invalid RSC Expense: " & Format(EXP, "$#,##0.00") & ", see " & data(2))
                                        End If
                                    End If
                                    rowSPTCWRX3.Item("RSCPAY") = Val(rowSPTCWRX3.Item("RSCPAY") & "") + PAY
                                    rowSPTCWRX3.Item("RSCFRG") = Val(rowSPTCWRX3.Item("RSCFRG") & "") + FRG
                                    rowSPTCWRX3.Item(RSCEXP_field) = Val(rowSPTCWRX3.Item(RSCEXP_field) & "") + EXP

                                ElseIf Trim(data(3)) = "SDS" Then
                                    SDSPAY += PAY
                                    SDSFRG += FRG
                                    SDSEXP += EXP

                                    Dim SDSEXP_field As String = "SDSEXP"
                                    'If ASCMAIN1.Running_in_VS AndAlso data(13) <> "" Then Stop
                                    If data(13) = "BYOD Expense Reimbursement" Then
                                        SDSEXP_field = "SDSFRG"
                                    End If

                                    rowSPTCWRX3.Item("SDSPAY") = Val(rowSPTCWRX3.Item("SDSPAY") & "") + PAY
                                    rowSPTCWRX3.Item("SDSFRG") = Val(rowSPTCWRX3.Item("SDSFRG") & "") + FRG
                                    rowSPTCWRX3.Item(SDSEXP_field) = Val(rowSPTCWRX3.Item(SDSEXP_field) & "") + EXP
                                Else
                                    Dim z As String = "Cannot Map Job Title " & Trim(data(3))
                                    If Not EMsg.Contains(z) Then
                                        EMsg &= vbCr & z
                                    End If
                                End If
                            End If
                        Next

                        ' Green
                        'Dim PAY As Decimal = Val(.Item("WAGES_SHIFT")) + Val(.Item("WAGES_RETRO")) + Val(.Item("WAGES_TRAIN")) + Val(.Item("WAGES_SICK")) + Val(.Item("WAGES_BONUS"))
                        ''Purple
                        'Dim FRG As Decimal = Val(.Item("OTHER_TAXES")) + Val(.Item("SALES_TAX_AMT")) + Val(.Item("COWORX_FEE"))
                        ''Blue
                        'Dim EXP As Decimal = Val(.Item("WAGES_EXP")) + Val(.Item("TRAVEL_EXP"))

                        ''Dim rowSPTCWRX3 As DataRow = dst.Tables("SPTCWRX3").Rows.Find(New String() {CTRL_NO, COLLECTION_CODE})
                        ''If rowSPTCWRX3 Is Nothing Then
                        ''    rowSPTCWRX3 = dst.Tables("SPTCWRX3").NewRow
                        ''    rowSPTCWRX3.Item("CTRL_NO") = CTRL_NO
                        ''    rowSPTCWRX3.Item("COLLECTION_CODE") = COLLECTION_CODE
                        ''    dst.Tables("SPTCWRX3").Rows.Add(rowSPTCWRX3)
                        ''End If


                        ''If Trim(data(3)) = "RSC" Or Trim(data(3)) = "FM" Or (Trim(data(3)) = "SDS" And CUST_CODE <> "SDS") Then
                        ''    RSCPAY += PAY
                        ''    RSCFRG += FRG
                        ''    RSCEXP += EXP

                        ''    If EXP <> 0 Then
                        ''        import_errors.Add("Invalid RSC Expense: " & Format(EXP, "$#,##0.00") & ", see " & data(2))
                        ''    End If
                        ''    rowSPTCWRX3.Item("RSCPAY") = Val(rowSPTCWRX3.Item("RSCPAY") & "") + PAY
                        ''    rowSPTCWRX3.Item("RSCFRG") = Val(rowSPTCWRX3.Item("RSCFRG") & "") + FRG
                        ''    rowSPTCWRX3.Item("RSCEXP") = Val(rowSPTCWRX3.Item("RSCEXP") & "") + EXP

                        ''ElseIf Trim(data(3)) = "SDS" Then
                        ''    SDSPAY += PAY
                        ''    SDSFRG += FRG
                        ''    SDSEXP += EXP

                        ''    rowSPTCWRX3.Item("SDSPAY") = Val(rowSPTCWRX3.Item("SDSPAY") & "") + PAY
                        ''    rowSPTCWRX3.Item("SDSFRG") = Val(rowSPTCWRX3.Item("SDSFRG") & "") + FRG
                        ''    rowSPTCWRX3.Item("SDSEXP") = Val(rowSPTCWRX3.Item("SDSEXP") & "") + EXP
                        ''Else
                        ''    Dim z As String = "Cannot Map Job Title " & Trim(data(3))
                        ''    If Not EMsg.Contains(z) Then
                        ''        EMsg &= vbCr & z
                        ''    End If
                        ''End If

                    End With

                    dst.Tables("SPTCWRX2").Rows.Add(rowSPTCWRX2)

                End If
            End If
        Next


        Dim YWx As String = "'" & Join(OPS_YYYYWWs.Keys.ToArray, "','") & "'"
        ASCMAIN1.sql = $"Select OPS_YYYYWW YW, SUM (NET_AMT) NET_AMT_PREV from SPTCWRX2 where OPS_YYYYWW IN ({YWx}) group by OPS_YYYYWW"
        tblEXP = ASCDATA1.GetDataTable
        tblEXP.PrimaryKey = New DataColumn() {tblEXP.Columns(0)}
        tblEXP.Columns.Add("NET_AMT_CURR", GetType(System.Decimal))
        tblEXP.Columns("YW").Caption = "YW"
        tblEXP.Columns("NET_AMT_PREV").Caption = "Prev Files"
        tblEXP.Columns("NET_AMT_CURR").Caption = "This File"

        For Each OPS_YYYYWW As String In OPS_YYYYWWs.Keys
            Dim row As DataRow = tblEXP.Rows.Find(OPS_YYYYWW)
            If row Is Nothing Then
                row = tblEXP.NewRow
                row.Item("YW") = OPS_YYYYWW
                tblEXP.Rows.Add(row)
            End If
            row.Item("NET_AMT_CURR") = OPS_YYYYWWs(OPS_YYYYWW)
        Next

        grdEXP.DisplayLayout.Bands(0).Summaries.Clear()

        grdEXP.DataSource = tblEXP
        Create_Summary(grdEXP, "NET_AMT_PREV")
        Create_Summary(grdEXP, "NET_AMT_CURR")

        Sort_grdColumns(grdEXP, "YW".ToLower)

        Dim INV_AMT_MENS As Decimal = Val(dst.Tables("SPTCWRX2").Compute("SUM (NET_AMT)", "CHECKBOOK = 'MENS'") & "")
        Dim INV_AMT_WMEN As Decimal = Val(dst.Tables("SPTCWRX2").Compute("SUM (NET_AMT)", "CHECKBOOK = 'WMEN'") & "")
        Dim INV_AMT_SDS As Decimal = Val(dst.Tables("SPTCWRX2").Compute("SUM (NET_AMT)", "CHECKBOOK = 'SDS'") & "")

        ' FIGURE OUT AMOUNT
        For Each rowSPTCWRX4 As DataRow In dst.Tables("SPTCWRX4").Select()
            If rowSPTCWRX4.Item("CHECKBOOK") = "WMEN" Then
                rowSPTCWRX4.Item("CWRX_EXP_AMT") = System.Math.Round(Val((rowSPTCWRX4.Item("CWRX_EXP_PCT")) / 100) * INV_AMT_WMEN, 2)
            ElseIf rowSPTCWRX4.Item("CHECKBOOK") = "MENS" Then
                rowSPTCWRX4.Item("CWRX_EXP_AMT") = System.Math.Round(Val((rowSPTCWRX4.Item("CWRX_EXP_PCT")) / 100) * INV_AMT_MENS, 2)
            ElseIf rowSPTCWRX4.Item("CHECKBOOK") = "SDS" Then
                rowSPTCWRX4.Item("CWRX_EXP_AMT") = System.Math.Round(Val(INV_AMT_SDS), 2)
            End If
        Next

        Dim rowSPTCWRX1 As DataRow = dst.Tables("SPTCWRX1").NewRow
        rowSPTCWRX1.Item("CTRL_NO") = CTRL_NO
        rowSPTCWRX1.Item("DATE_PROCESSED") = DATETIME_STAMP
        rowSPTCWRX1.Item("RECORD_COUNT") = records
        INV_NUM = Format(INV_DATE, "yyyyMMdd")
        rowSPTCWRX1.Item("INV_NUM") = INV_NUM
        rowSPTCWRX1.Item("INV_DATE") = INV_DATE
        rowSPTCWRX1.Item("INV_AMT") = INV_AMT

        rowSPTCWRX1.Item("RSCPAY") = RSCPAY
        rowSPTCWRX1.Item("RSCFRG") = RSCFRG
        rowSPTCWRX1.Item("RSCEXP") = RSCEXP

        rowSPTCWRX1.Item("SDSPAY") = SDSPAY
        rowSPTCWRX1.Item("SDSFRG") = SDSFRG
        rowSPTCWRX1.Item("SDSEXP") = SDSEXP

        dst.Tables("SPTCWRX1").Rows.Add(rowSPTCWRX1)

        If EMsg <> "" Then
            MsgBox("Encountered an unmappable Job Title." _
                   & vbCrLf & "AP Invoice might Not balance." _
                   & vbCrLf & "Do Not Update- Call ABS." _
                   & vbCrLf & vbCrLf & Mid(EMsg, 2),
                   MsgBoxStyle.OkOnly, "Please Call ABS For Support")
            update_blocked = True
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Create_AP()

        Dim VEND_CODE As String = ROWs("SPTPARM1").Item("SP_PARM_MODEL_VEND_CODE")

        Dim DIST_PCTs As New Dictionary(Of String, Decimal)

        ASCMAIN1.sql = "Select GLTDIST2.* from GLTDIST2,GLTDIST1" & vbCrLf _
            & " where GLTDIST2.DIST_CODE = GLTDIST1.DIST_CODE" & vbCrLf _
            & "   And GLTDIST1.DIST_STATUS = 'A'" & vbCrLf _
            & "   and GLTDIST2.DIST_PCT <> 0" & vbCrLf _
            & "   and GLTDIST2.DIST_CODE = :PARM1"
        Dim DIST_PCT_total As Decimal = 0
        Dim DIST_AMT_total As Decimal = 0
        Dim ACCT_SEG_CODE_last As String = ""
        For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, "GLTDIST2", "V", "COWORX").Select("", "DIST_PCT")
            Dim ACCT_SEG_CODE As String = row.Item("ACCT_SEG_CODE")
            Dim DIST_PCT As Decimal = Val(row.Item("DIST_PCT") & "")
            DIST_PCT_total += DIST_PCT
            DIST_PCTs.Add(ACCT_SEG_CODE, DIST_PCT)
            ACCT_SEG_CODE_last = ACCT_SEG_CODE
        Next
        If DIST_PCTs.Count <> 0 Then
            If DIST_PCT_total <> 100 Then
                DIST_PCTs(ACCT_SEG_CODE_last) = 100 - DIST_PCT_total
            End If
        End If

        For Each rowSPTCWRX1 As DataRow In dst.Tables("SPTCWRX1").Select("")
            Dim CTRL_NO As String = rowSPTCWRX1.Item("CTRL_NO")

            Dim INV_NUM As String = rowSPTCWRX1.Item("INV_NUM")
            Dim INV_DATE As Date = rowSPTCWRX1.Item("INV_DATE")
            Dim INV_AMT As String = Val(rowSPTCWRX1.Item("INV_AMT") & "")

            Dim VOUCHER_NO As String = ASCMAIN1.Next_Control_No("APTINVH1.VOUCHER_NO")
            rowSPTCWRX1.Item("VOUCHER_NO") = VOUCHER_NO

            For Each rowASTATTA2 As DataRow In dst.Tables("ASTATTA2").Select("")
                rowASTATTA2.Item("CODE_VALUE") = VOUCHER_NO
            Next

            Dim VOUCHER_LNO As Integer = 0

            Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", VEND_CODE)

            Dim SPLITS As New Dictionary(Of Integer, Decimal)

            For Each rowSPTCWRX3 As DataRow In dst.Tables("SPTCWRX3").Select("CTRL_NO = '" & CTRL_NO & "'", "COLLECTION_CODE")
                Dim COLLECTION_CODE As String = rowSPTCWRX3.Item("COLLECTION_CODE")
                Dim SEG4_CODE As String = COLLECTION_CODE
                If COLLECTION_CODE <> "000" Then
                    Dim rowICTCOLL1 As DataRow = LookUp("ICTCOLL1", COLLECTION_CODE)
                    If rowICTCOLL1.Item("SEG4_CODE") & "" <> "" Then
                        SEG4_CODE = rowICTCOLL1.Item("SEG4_CODE")
                    End If
                End If

                For Each COL As String In New String() {"RSCPAY", "RSCFRG", "RSCEXP", "SDSPAY", "SDSFRG", "SDSEXP"}
                    Dim INV_LINE_AMT As Decimal = Val(rowSPTCWRX3.Item(COL) & "")
                    If INV_LINE_AMT <> 0 Then
                        Dim rowAPTINVH2 As DataRow = dst.Tables("APTINVH2").NewRow
                        With rowAPTINVH2
                            .Item("VOUCHER_NO") = VOUCHER_NO
                            VOUCHER_LNO = VOUCHER_LNO + 1
                            .Item("VOUCHER_LNO") = VOUCHER_LNO
                            .Item("ACCT_CODE") = ROWs("SPTPARM1").Item("SP_PARM_MODEL_ACCT_CODE_" & COL)
                            .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                            .Item("SEG3_CODE") = "DPT" ' ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                            .Item("SEG4_CODE") = SEG4_CODE
                            .Item("INV_LINE_AMT") = INV_LINE_AMT

                            If .Item("ACCT_CODE") & "" = "" Then
                                import_errors.Add("Invoice " & INV_NUM & " - No GL Account set up for " & COL & ", " & Format(INV_LINE_AMT, "$#,###.00"))
                            End If

                        End With
                        dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)
                        If SEG4_CODE = "000" Then
                            SPLITS.Add(VOUCHER_LNO, INV_LINE_AMT)
                        End If
                    End If
                Next
            Next

            If SPLITS.Count <> 0 Then
                For Each VOUCHER_LNO_000 As Integer In SPLITS.Keys
                    Dim rowAPTINVH2_000 As DataRow = dst.Tables("APTINVH2").Rows.Find(New Object() {VOUCHER_NO, VOUCHER_LNO_000})
                    Dim rowAPTINVH2 As DataRow = Nothing
                    Dim INV_LINE_AMT As Decimal = SPLITS(VOUCHER_LNO_000)

                    If DIST_PCTs.Count = 0 Then
                        Dim INV_LINE_AMT_HALF = System.Math.Round(INV_LINE_AMT / 2, 2)

                        rowAPTINVH2 = dst.Tables("APTINVH2").NewRow
                        rowAPTINVH2.ItemArray = rowAPTINVH2_000.ItemArray
                        VOUCHER_LNO = VOUCHER_LNO + 1
                        rowAPTINVH2.Item("VOUCHER_LNO") = VOUCHER_LNO
                        rowAPTINVH2.Item("SEG4_CODE") = "JCHSIG"
                        rowAPTINVH2.Item("INV_LINE_AMT") = INV_LINE_AMT_HALF
                        rowAPTINVH2.Item("INV_COMMENT_DTL") = "JCHSIG SPLIT"
                        dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)

                        rowAPTINVH2 = dst.Tables("APTINVH2").NewRow
                        rowAPTINVH2.ItemArray = rowAPTINVH2_000.ItemArray
                        VOUCHER_LNO = VOUCHER_LNO + 1
                        rowAPTINVH2.Item("VOUCHER_LNO") = VOUCHER_LNO
                        rowAPTINVH2.Item("SEG4_CODE") = "MBCLEH"
                        rowAPTINVH2.Item("INV_LINE_AMT") = INV_LINE_AMT - INV_LINE_AMT_HALF
                        rowAPTINVH2.Item("INV_COMMENT_DTL") = "MBCLEH SPLIT"
                        dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)
                    Else
                        Dim INV_LINE_AMT_total As Decimal = 0
                        For Each ACCT_SEG_CODE As String In DIST_PCTs.Keys
                            rowAPTINVH2 = dst.Tables("APTINVH2").NewRow
                            rowAPTINVH2.ItemArray = rowAPTINVH2_000.ItemArray
                            VOUCHER_LNO = VOUCHER_LNO + 1
                            rowAPTINVH2.Item("VOUCHER_LNO") = VOUCHER_LNO
                            rowAPTINVH2.Item("SEG4_CODE") = ACCT_SEG_CODE
                            Dim INV_LINE_AMT_dist As Decimal = System.Math.Round(DIST_PCTs(ACCT_SEG_CODE) * INV_LINE_AMT / 100 + 0.001, 2)
                            INV_LINE_AMT_total += INV_LINE_AMT_dist
                            rowAPTINVH2.Item("INV_LINE_AMT") = INV_LINE_AMT_dist
                            rowAPTINVH2.Item("INV_COMMENT_DTL") = "SDS SPLIT"
                            dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)
                        Next

                        If INV_LINE_AMT_total <> INV_LINE_AMT Then
                            rowAPTINVH2.Item("INV_LINE_AMT") += INV_LINE_AMT - INV_LINE_AMT_total
                        End If
                    End If


                    rowAPTINVH2_000.Delete()
                Next
            End If

            Dim INV_LINE_AMT_all As Decimal = Val(dst.Tables("APTINVH2").Compute("SUM(INV_LINE_AMT)", "") & "")
            If INV_AMT <> INV_LINE_AMT_all Then
                Dim rowAPTINVH2 As DataRow = dst.Tables("APTINVH2").NewRow
                With rowAPTINVH2
                    .Item("VOUCHER_NO") = VOUCHER_NO
                    VOUCHER_LNO = VOUCHER_LNO + 1
                    .Item("VOUCHER_LNO") = VOUCHER_LNO
                    .Item("ACCT_CODE") = ROWs("GLTPARM1").Item("GL_PARM_ACCT_ROUNDING")
                    .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                    .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                    .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                    .Item("INV_LINE_AMT") = INV_AMT - INV_LINE_AMT_all
                End With
                dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)
            End If


            'Dim INV_AMT As Decimal = Val(Absx1.numFor("PYMT_REF_AMT").Value & "")

            Dim rowAPTINVH1 As DataRow = dst.Tables("APTINVH1").NewRow
            With rowAPTINVH1
                .Item("VOUCHER_NO") = VOUCHER_NO
                .Item("VEND_CODE") = VEND_CODE
                .Item("INV_TYPE") = "I"
                .Item("INV_NUM") = INV_NUM
                .Item("INV_DATE") = INV_DATE

                .Item("INV_AMT") = INV_AMT
                .Item("INV_REF") = rowSPTCWRX1.Item("CTRL_NO")

                .Item("VEND_CODE_AP") = rowAPTVEND1.Item("VEND_CODE_AP")
                If rowAPTVEND1.Item("VEND_PYMT_ADDR") & "" = "" Then
                    .Item("VEND_ALT_CODE") = ""
                    .Item("INV_REMIT_TO") = "V"
                Else
                    .Item("VEND_ALT_CODE") = rowAPTVEND1.Item("VEND_PYMT_ADDR")
                    .Item("INV_REMIT_TO") = "A"
                End If

                .Item("INV_SEP_CHECK") = rowAPTVEND1.Item("VEND_SEP_CHECKS")
                .Item("TERM_CODE") = rowAPTVEND1.Item("TERM_CODE")

                If rowAPTVEND1.Item("BANK_CODE") & "" = "" Then
                    .Item("BANK_CODE") = ROWs("APTPARM1").Item("AP_PARM_BANK_CODE")
                Else
                    .Item("BANK_CODE") = rowAPTVEND1.Item("BANK_CODE")
                End If

                If rowAPTVEND1.Item("VEND_PYMT_METHOD") & "" = "" Then
                    If .Item("BANK_CODE") & "" <> "" Then
                        Dim rowGLTBANK1 As DataRow = LookUp("GLTBANK1", .Item("BANK_CODE"))
                        .Item("INV_PYMT_METHOD") = rowGLTBANK1.Item("BANK_PYMT_METHOD")
                    End If
                Else
                    .Item("INV_PYMT_METHOD") = rowAPTVEND1.Item("VEND_PYMT_METHOD")
                End If

                .Item("INV_PYMT_CYCLE") = rowAPTVEND1.Item("VEND_PYMT_CYCLE")

                If rowAPTVEND1.Item("POST_CODE") & "" <> "" Then
                    .Item("POST_CODE") = rowAPTVEND1.Item("POST_CODE")
                Else
                    .Item("POST_CODE") = ROWs("APTPARM1").Item("AP_PARM_POST_CODE")
                End If

                .Item("INV_STATUS") = "O"
                .Item("INV_PYMT_CYCLE") = DBNull.Value
                .Item("INV_DUE_DATE") = TAC.TACMAIN1.Calculate_INV_DUE_DATE(Me, rowAPTVEND1.Item("TERM_CODE") & "", Nothing, .Item("INV_DATE"))
                .Item("INV_BALANCE") = INV_AMT
                .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
                .Item("CURR_EXCH_RATE") = 1

                .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")

                .Item("REGISTER_IND") = "0"
                .Item("INV_BL_DATE") = .Item("INV_DATE")
                .Item("INV_AMT_VEND") = .Item("INV_AMT")

                'Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", YW_INVOICE)
                'Dim OPS_YYYYPP_ACCRUE As String = rowGLTPARM3.Item("YYYYPP")
                Dim OPS_YYYYPP_ACCRUE As String = Format(INV_DATE.AddDays(-7), "yyyyMM")
                If OPS_YYYYPP_ACCRUE <> ASCMAIN1.CYP Then
                    .Item("OPS_YYYYPP_ACCRUE") = OPS_YYYYPP_ACCRUE
                End If

                If rowAPTVEND1.Item("VEND_AUTO_APPROVE") & "" = "1" Then
                    .Item("INV_APPR_STATUS") = "A"
                    Write_Event_Log("APTINVH1", VOUCHER_NO, "Auto Approved")
                Else
                    .Item("INV_APPR_STATUS") = "P"
                End If

                .Item("VEND_BUYER_CODE") = rowAPTVEND1.Item("VEND_BUYER_CODE")
            End With

            dst.Tables("APTINVH1").Rows.Add(rowAPTINVH1)
        Next


    End Sub
    Sub PopulateSPTARCHD()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now loading archived documents")

        Dim archivePath As String = "\\nymain-abs-iis1\sftp\coworx\FromCoworx\Archive"
        Dim tempPath As String = ASCMAIN1.Folders("Temp")
        Dim di As New System.IO.DirectoryInfo(archivePath)
        Dim files As System.IO.FileInfo() = di.GetFiles("*.csv").Concat(di.GetFiles("*.xls")).Concat(di.GetFiles("*.xlsx")).ToArray()

        Dim dt As DataTable = dst.Tables("SPTARCHD")
        dt.Clear()

        For Each fi As System.IO.FileInfo In files
            Dim tempFilePath As String = System.IO.Path.Combine(tempPath, fi.Name)
            System.IO.File.Copy(fi.FullName, tempFilePath, True)

            Dim dr As DataRow = dt.NewRow()
            dr("FILENAME") = tempFilePath  ' Store the path of the copied file
            dr("FILEDATE") = fi.CreationTime
            dr("FILESIZE") = fi.Length
            dr("FILEABBR") = fi.Name
            dt.Rows.Add(dr)
        Next

        grdSPTARCHD.DataSource = dt

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub


    Private Sub UltraButton3_Click(sender As Object, e As EventArgs) Handles UltraButton3.Click
        grdSPTARCHD.Visible = True
        SplitContainer3.Panel2Collapsed = False  '

        PopulateSPTARCHD()
    End Sub

    Private Sub grdSPTARCHD_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdSPTARCHD.DoubleClickRow
        Dim row As UltraGridRow = grdSPTARCHD.ActiveRow
        Dim FILENAME As String = row.Cells("FILEABBR").Value.ToString()
        Dim tempPath As String = ASCMAIN1.Folders("Temp")
        Dim PATH As String = System.IO.Path.Combine(tempPath, FILENAME)

        If System.IO.File.Exists(PATH) Then
            Process.Start(PATH)
        Else
            MessageBox.Show("File not found: " & PATH, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class