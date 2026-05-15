Imports ABSolution
Imports Infragistics.Win
Imports System.Windows.forms

Public Class EDF852I1

    Dim EDI_DOC_SEQ_NOs As New List(Of String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("EDTPARM1")

        With dst

            '& ", (Select EDI_DOC_SEQ_NO, Count(*) ERRS_ITEM from (Select Distinct EDT852T0.EDI_DOC_SEQ_NO, EDT852T0.EDI_ITEM_CODE from EDT852T0,EDT852T1 where EDT852T1.EDI_DOC_SEQ_NO = EDT852T0.EDI_DOC_SEQ_NO) XI" _
            '& ", () XS" _

            ASCMAIN1.sql = "Select EDT852T1.*, ARTCUST1.CUST_NAME" _
            & ", '0' SELECTED " _
            & " from EDT852T1,ARTCUST1 " _
            & " where EDT852T1.EDI_STATUS = :PARM1 " _
            & " and ARTCUST1.CUST_CODE (+) = EDT852T1.CUST_CODE"
            Create_TDA(.Tables.Add, "EDT852T1", "**", 0, True, "V")


            ASCMAIN1.sql = "Select EDT852T0.EDI_DOC_SEQ_NO" _
            & ", EDT852T0.EDI_ITEM_CODE, EDT852T0.ITEM_CODE, ICTITEM1.ITEM_DESC" _
            & ", COUNT(*) RECORDS" _
            & ", SUM (DECODE(EDT852T0.EDI_TRAN_TYPE,'QA',NVL(EDT852T0.EDI_QTY,0),0)) QTY_ONH" _
            & ", SUM (DECODE(EDT852T0.EDI_TRAN_TYPE,'QS',NVL(EDT852T0.EDI_QTY,0),0)) QTY_SLS" _
            & ", SUM (DECODE(EDT852T0.EDI_TRAN_TYPE,'QU',NVL(EDT852T0.EDI_QTY,0),0)) QTY_RTN" _
            & " from EDT852T0,ICTITEM1 " _
            & " where EDT852T0.EDI_DOC_SEQ_NO = :PARM1 " _
            & " and ICTITEM1.ITEM_CODE (+) = EDT852T0.ITEM_CODE" _
            & " group by EDT852T0.EDI_DOC_SEQ_NO" _
            & ", EDT852T0.EDI_ITEM_CODE, EDT852T0.ITEM_CODE, ICTITEM1.ITEM_DESC"
            Create_TDA(.Tables.Add, "EDT852TI", "**", 0, False, "V")

            For Each dc As String In New String() {"RECORDS", "QTY_ONH", "QTY_SLS", "QTY_RTN"}
                dst.Tables("EDT852TI").Columns(dc).DataType = GetType(System.Int32)
            Next

            ASCMAIN1.sql = "Select EDT852T0.EDI_DOC_SEQ_NO" _
            & ", EDT852T0.EDI_STORE_NO, EDT852T0.CUST_CODE" _
            & ", ARTCUST2.CUST_STORE_NO" _
            & ", NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_CITY) CUST_STORE_LOCATION" _
            & ", COUNT(*) RECORDS" _
            & ", SUM (DECODE(EDT852T0.EDI_TRAN_TYPE,'QA',NVL(EDT852T0.EDI_QTY,0),0)) QTY_ONH" _
            & ", SUM (DECODE(EDT852T0.EDI_TRAN_TYPE,'QS',NVL(EDT852T0.EDI_QTY,0),0)) QTY_SLS" _
            & ", SUM (DECODE(EDT852T0.EDI_TRAN_TYPE,'QU',NVL(EDT852T0.EDI_QTY,0),0)) QTY_RTN" _
            & " from EDT852T0,ARTCUST2 " _
            & " where EDT852T0.EDI_DOC_SEQ_NO = :PARM1 " _
            & " and ARTCUST2.CUST_CODE (+) = EDT852T0.CUST_CODE" _
            & " and ARTCUST2.CUST_STORE_NO (+) = EDT852T0.CUST_STORE_NO" _
            & " group by EDT852T0.EDI_DOC_SEQ_NO" _
            & ", EDT852T0.EDI_STORE_NO, EDT852T0.CUST_CODE" _
            & ", ARTCUST2.CUST_STORE_NO" _
            & ", NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_CITY)"
            Create_TDA(.Tables.Add, "EDT852TS", "**", 0, False, "V")

            For Each dc As String In New String() {"RECORDS", "QTY_ONH", "QTY_SLS", "QTY_RTN"}
                dst.Tables("EDT852TS").Columns(dc).DataType = GetType(System.Int32)
            Next


            Create_TDA(.Tables.Add, "EDT852T2", "*", 1)
            Create_TDA(.Tables.Add, "EDT852T3", "*", 1)

            .Relations.Add("EDT852T2_EDT852T3" _
                           , New DataColumn() _
                             {.Tables("EDT852T2").Columns("EDI_DOC_SEQ_NO") _
                              , .Tables("EDT852T2").Columns("EDI_LINE_NO")} _
                             , New DataColumn() _
                             {.Tables("EDT852T3").Columns("EDI_DOC_SEQ_NO") _
                              , .Tables("EDT852T3").Columns("EDI_LINE_NO")})

            Create_TDA(.Tables.Add, "EDTFILE1", "*")
            With .Tables("EDTFILE1").Columns
                .Add("EDI_JRNL_NO")
                .Add("EDI_SENDER_QUAL")
                .Add("EDI_SENDER_ID")
                .Add("EDI_ISA_CTL_NO")
                .Add("EDI_ISA_CTL_DATE", GetType(System.DateTime))
                .Add("DOC_EDI", GetType(System.Int32))
                .Add("DOC_852", GetType(System.Int32))
                .Add("NOTES")
            End With

            ASCMAIN1.sql = "Select CUST_CODE from ARTCUST1"
            Create_TDA(.Tables.Add, "EDT852TC", "**", 0, False)
            With .Tables("EDT852TC").Columns
                For I As Integer = 1 To 27
                    .Add("F" & Format(I, "00"), GetType(System.Int32))
                Next
            End With

            ASCMAIN1.sql = "Select EDTJRNL3.*" _
            & " from EDTJRNL3 " _
            & " where EDI_DOC_SEQ_NO = :PARM1"
            Create_TDA(.Tables.Add, "EDTJRNL3", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "EDTJRNL1", "*")
            Create_TDA(.Tables.Add, "EDTJRNL2", "*")

            .Relations.Add("EDTJRNL1_EDTJRNL2" _
                             , .Tables("EDTJRNL1").Columns("EDI_JRNL_NO") _
                             , .Tables("EDTJRNL2").Columns("EDI_JRNL_NO"))

            .Relations.Add("EDTJRNL2_EDTJRNL3" _
                           , New DataColumn() _
                             {.Tables("EDTJRNL2").Columns("EDI_JRNL_NO") _
                              , .Tables("EDTJRNL2").Columns("EDI_GS_NO")} _
                             , New DataColumn() _
                             {.Tables("EDTJRNL3").Columns("EDI_JRNL_NO") _
                              , .Tables("EDTJRNL3").Columns("EDI_GS_NO")})

        End With

        grdEDTFILE1.DataSource = dst.Tables("EDTFILE1")

        grdEDT852TC.DataSource = dst.Tables("EDT852TC")

        grdEDTJRNL1.DataSource = dst.Tables("EDTJRNL1")
        'grdEDTJRNL1.DataMember = "EDTJRNL1"
        'grdEDTJRNL1.DataSource = dst

        grdEDT852T1.DataSource = dst.Tables("EDT852T1")
        grdEDT852TI.DataSource = dst.Tables("EDT852TI")
        grdEDT852TS.DataSource = dst.Tables("EDT852TS")

        'grdEDT852T2.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.MultiBand
        grdEDT852T2.DataSource = dst.Tables("EDT852T2")
        'grdEDT852T2.DataMember = "EDT852T2"
        'grdEDT852T2.DataSource = dst

        Create_Summary(grdEDTFILE1, "EDI_FILENAME", "Count")
        Create_Summary(grdEDTFILE1, "DOC_EDI")
        Create_Summary(grdEDTFILE1, "DOC_852")

        Create_Summary(grdEDT852T1, "EDI_DOC_SEQ_NO", "Count")
        Create_Summary(grdEDT852T1, "SELECTED")

        Create_Summary(grdEDT852TC, "CUST_CODE", "Count")
        For W As Integer = 1 To 27
            Create_Summary(grdEDT852TC, "F" & Format(W, "00"))
        Next

        Create_Summary(grdEDT852TI, "EDI_ITEM_CODE", "Count")
        Create_Summary(grdEDT852TI, "RECORDS")
        Create_Summary(grdEDT852TI, "QTY_ONH")
        Create_Summary(grdEDT852TI, "QTY_SLS")
        Create_Summary(grdEDT852TI, "QTY_RTN")

        Create_Summary(grdEDT852TS, "EDI_STORE_NO", "Count")
        Create_Summary(grdEDT852TS, "RECORDS")
        Create_Summary(grdEDT852TS, "QTY_ONH")
        Create_Summary(grdEDT852TS, "QTY_SLS")
        Create_Summary(grdEDT852TS, "QTY_RTN")

        With grdEDT852T1.DisplayLayout.Bands("EDT852T1")
            .Columns("EDI_DOC_SEQ_NO").Header.Fixed = True
        End With

        With grdEDTFILE1.DisplayLayout.Bands("EDTFILE1")
            .Columns("EDI_FILENAME").Header.Fixed = True
        End With

        With grdEDT852TC.DisplayLayout.Bands("EDT852TC")
            .Columns("CUST_CODE").Header.Fixed = True
        End With

        With grdEDT852T2.DisplayLayout.Bands("EDT852T2_EDT852T3")
            For I As Integer = 1 To 10
                .Columns("EDI_SDQ_STORE_" & Format(I, "00")).Header.Caption = "S" & Format(I, "00")
                .Columns("EDI_SDQ_QTY_AMT_" & Format(I, "00")).Header.Caption = "Q" & Format(I, "00")
                .Columns("EDI_SDQ_STORE_" & Format(I, "00")).Width = 45
                .Columns("EDI_SDQ_QTY_AMT_" & Format(I, "00")).Width = 45
                .Columns("EDI_SDQ_STORE_" & Format(I, "00")).CellAppearance.BackColor = Drawing.Color.Beige
            Next
        End With

        For Each gcol As UltraWinGrid.UltraGridColumn In grdEDT852T1.DisplayLayout.Bands(0).Columns
            If gcol.Key = "SELECTED" Then
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next


        grdEDT852TC.DisplayLayout.Bands(0).Columns("CUST_CODE").CellAppearance.BackColor = Drawing.Color.Beige
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("CUST_CODE")
                Validate_Code("OPS_YYYYPP")

                'If Absx1.dteFor("DTE0").Value & "" = "" Then
                '    EMsg &= vbCr & "You must Specify a Starting Date"
                'End If

            Case "Import Raw EDI Files"
                If chkRawPreviouslyImported.Checked Then
                    EMsg &= vbCr & "You must first uncheck the option to Show Previously Imported files"
                End If

            Case "Load 852 Data", "Retract 852 Data", "Restore Deleted"
                If dst.Tables("EDT852T1").Select("SELECTED = '1'").Length = 0 Then
                    EMsg &= vbCr & "No Documents Selected"
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

            Case "Load"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Import Raw EDI Files"
                Call Mode_Settings(True)
                Call Import_Raw_EDI()
                Call Mode_Settings(False)

            Case "Load 852 Data", "Retract 852 Data", "Restore Deleted"
                Call Mode_Settings(True)
                Call Load_852_Data()
                Call Mode_Settings(False)

            Case "Cancel"
                Call Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
            End With
        End If

        'Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        'UltraTabControl1.Visible = tf

        optMode.Enabled = Not ScreenMode
        SplitContainer5.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
            Setup_tabMain()
        End If

    End Sub

    Sub Clear_Record()
        Select_tabMain()

        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("OPS_YYYYPP").Text = ""

        dst.EnforceConstraints = False
        dst.Tables("EDTJRNL1").Rows.Clear()
        dst.Tables("EDTJRNL2").Rows.Clear()
        dst.Tables("EDTJRNL3").Rows.Clear()

        dst.Tables("EDT852T1").Rows.Clear()
        dst.Tables("EDT852T2").Rows.Clear()
        dst.Tables("EDT852T3").Rows.Clear()
        dst.EnforceConstraints = True

        Prepare_852_Queue()
        Setup_tabRaw()
    End Sub

    Sub Load_Record()

        Call ASCMAIN1.Progress("Now Loading Gentran Data")

        Call Save_Header_Fields(UltraGroupBox1)

        dst.EnforceConstraints = False

        dst.EnforceConstraints = True

        Call ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdEDT852T1, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Dim tlb_sbt As UltraWinToolbars.StateButtonTool
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden

            Select Case e.SourceControl.Name
                Case "grdEDT852T1"

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)
            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub dte_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.dte_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "DTE0", "DTE1"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    Call Click_Command("Load", e)
                End If
        End Select

    End Sub

#End Region

    Sub Print_Report()
        Call Print_Report_Begin()
        Dim SUBT As String = ""

        Dim RecordSelectionFormula As String = ""

        Generate_Report("EDRSTATI", "Inbound EDI Transactions", SUBT, RecordSelectionFormula)

        Call Print_Report_End()
    End Sub

    Private Sub View_doc(ByVal DOCUMENTBLOBKEY As String)
        Dim FILENAME As String = "\\192.168.130.206\E$\GENSRVNT\DOCUMENTS\" & DOCUMENTBLOBKEY & ".DOC"

        If My.Computer.FileSystem.FileExists(FILENAME) Then
            Dim TEMP_FILENAME As String = ASCMAIN1.Folders("Temp") & DOCUMENTBLOBKEY & ".DOC"
            If My.Computer.FileSystem.FileExists(TEMP_FILENAME) Then
                My.Computer.FileSystem.DeleteFile(TEMP_FILENAME)
            End If
            My.Computer.FileSystem.CopyFile(FILENAME, TEMP_FILENAME)
            Dim p As Process = Process.Start("NOTEPAD.EXE", TEMP_FILENAME)
        End If
    End Sub

    Private Sub UltraTabControl1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()
        If SELECTION_NO = 0 Then Exit Sub

        If tabMain.SelectedTab Is Nothing Then Exit Sub

        With UltraExplorerBar1.Groups("Batch Control")
            .Items("Import Raw EDI Files").Visible = (tabMain.SelectedTab.Key = "Raw EDI Files")
            .Items("Load 852 Data").Visible = (tabMain.SelectedTab.Key = "852 Data" And opt852Data.Value = "0")
            .Items("Retract 852 Data").Visible = (tabMain.SelectedTab.Key = "852 Data" And opt852Data.Value = "1")
            .Items("Restore Deleted").Visible = (tabMain.SelectedTab.Key = "852 Data" And opt852Data.Value = "D")
        End With

        With UltraExplorerBar1

            .Groups("Screen Control").Visible = (tabMain.SelectedTab.Key = "Manual Entry")

            Select Case tabMain.SelectedTab.Key
                Case "Raw EDI Files"
                    grdEDTJRNL1.Dock = DockStyle.None
                    grdEDTJRNL1.Parent = tabRaw.Tabs("Control Records").TabPage
                    grdEDTJRNL1.Dock = DockStyle.Fill

                    grpRawEDI.Dock = DockStyle.None
                    grpRawEDI.Parent = tabRaw.Tabs("Raw EDI").TabPage
                    grpRawEDI.Dock = DockStyle.Fill

                    grdEDT852T1.Dock = DockStyle.None
                    grdEDT852T1.Parent = tabRaw.Tabs("Documents").TabPage
                    grdEDT852T1.Dock = DockStyle.Fill
                Case "852 Data"
                    grdEDTJRNL1.Dock = DockStyle.None
                    grdEDTJRNL1.Parent = tab852.Tabs("Control Records").TabPage
                    grdEDTJRNL1.Dock = DockStyle.Fill

                    grpRawEDI.Dock = DockStyle.None
                    grpRawEDI.Parent = tab852.Tabs("Raw EDI").TabPage
                    grpRawEDI.Dock = DockStyle.Fill

                    grdEDT852T1.Dock = DockStyle.None
                    grdEDT852T1.Parent = SplitContainer3.Panel1
                    grdEDT852T1.Dock = DockStyle.Fill

                    Setup_EDT852T1()

                Case "Manual Entry"
                    grdEDT852T1.Dock = DockStyle.None
                    grdEDT852T1.Parent = SplitContainer5.Panel1
                    grdEDT852T1.Dock = DockStyle.Fill

                    'Setup_EDT852T1()
            End Select

            .Groups("Raw EDI Files").Visible = (tabMain.SelectedTab.Key = "Raw EDI Files")
            .Groups("852 Data").Visible = (tabMain.SelectedTab.Key = "852 Data")
            .Groups("Manual Entry").Visible = (tabMain.SelectedTab.Key = "Manual Entry")
            .Groups("Batch Control").Visible = (tabMain.SelectedTab.Key <> "Manual Entry")

        End With
    End Sub

    Sub Import_Raw_EDI()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Processing EDI Raw Files")

        Dim EDCIBND1 As New EDCIBND1()

        dst.Tables("EDTFILE1").Rows.Clear()

        Dim file_counter As Integer = 0

        Dim wildcard As String = "*.edi" ' "MAIL_IN.TXT"

        Dim ED_PARM_RAW_INBOUND As String = ROWs("EDTPARM1").Item("ED_PARM_RAW_INBOUND") & ""
        For Each FILE As String In My.Computer.FileSystem.GetFiles _
        (ED_PARM_RAW_INBOUND, FileIO.SearchOption.SearchAllSubDirectories, wildcard)
            Dim FILEINFO As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILE)
            Dim rowEDTFILE1 As DataRow = dst.Tables("EDTFILE1").NewRow
            Dim FILENAME As String = Mid(FILEINFO.FullName, ED_PARM_RAW_INBOUND.Length + 2)

            Dim row As DataRow = LookUp("EDTFILE1", FILENAME)
            If row Is Nothing Then
                rowEDTFILE1.Item("EDI_FILENAME") = FILENAME
                rowEDTFILE1.Item("EDI_FILESIZE") = FILEINFO.Length
                rowEDTFILE1.Item("EDI_DATETIME") = FILEINFO.LastWriteTime
                dst.Tables("EDTFILE1").Rows.Add(rowEDTFILE1)

                file_counter += 1

                ASCMAIN1.Progress("Processing " & FILENAME)
                Dim EDI_JRNL_NOs As List(Of String)
                EDI_JRNL_NOs = EDCIBND1.Process_File( _
                ED_PARM_RAW_INBOUND, FILENAME, _
                FILEINFO, "852", _
                ROWs("EDTPARM1").Item("ED_PARM_RAW_ARCHIVE") & "", _
                "", False)

                If EDI_JRNL_NOs.Count > 0 Then
                    Dim rowEDTJRNL1 = LookUp("EDTJRNL1", EDI_JRNL_NOs(0))
                    rowEDTFILE1.Item("EDI_JRNL_NO") = rowEDTJRNL1.item("EDI_JRNL_NO")
                    rowEDTFILE1.Item("EDI_SENDER_QUAL") = rowEDTJRNL1.item("EDI_SENDER_QUAL")
                    rowEDTFILE1.Item("EDI_SENDER_ID") = rowEDTJRNL1.item("EDI_SENDER_ID")
                    rowEDTFILE1.Item("EDI_ISA_CTL_NO") = rowEDTJRNL1.item("EDI_ISA_CTL_NO")
                    rowEDTFILE1.Item("EDI_ISA_CTL_DATE") = rowEDTJRNL1.item("EDI_ISA_CTL_DATE")
                    ASCMAIN1.sql = "Select Count (*) from EDTJRNL3 where EDI_JRNL_NO = '" & EDI_JRNL_NOs(0) & "'"
                    Dim DOC_EDI As Integer = ASCDATA1.GetDataValue
                    ASCMAIN1.sql = "Select Count (*) from EDTJRNL3 where EDI_JRNL_NO = '" & EDI_JRNL_NOs(0) & "' and EDI_DOC_NO = '852'"
                    Dim DOC_852 As Integer = ASCDATA1.GetDataValue
                    rowEDTFILE1.Item("DOC_EDI") = DOC_EDI
                    rowEDTFILE1.Item("DOC_852") = DOC_852
                    rowEDTFILE1.Item("NOTES") = ""
                End If

                ASCMAIN1.Progress("Now Processing EDI Raw Files")
            End If

        Next

        Update_Record_TDA("EDTFILE1")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        EDCIBND1 = Nothing

        Prepare_852_Queue()

        MsgBox(CStr(file_counter) & " Files have been Imported", MsgBoxStyle.OkOnly, "Verification")

    End Sub

    Sub Prepare_852_Queue()
        ASCMAIN1.sql = "Update EDT852T1 set EDI_STATUS = '0' " _
        & " where EDI_STATUS is Null"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS " _
        & " SELECT EDT852T1.EDI_DOC_SEQ_NO, EDTTRPM1.CUST_CODE " _
        & " FROM EDT852T1,EDTJRNL3,EDTJRNL1,EDTTRPM1 " _
        & " WHERE EDTJRNL3.EDI_DOC_SEQ_NO = EDT852T1.EDI_DOC_SEQ_NO " _
        & " AND EDTJRNL1.EDI_JRNL_NO = EDTJRNL3.EDI_JRNL_NO " _
        & " AND EDTTRPM1.EDI_TP_QUAL = TRIM(EDTJRNL1.EDI_SENDER_QUAL)" _
        & " AND EDTTRPM1.EDI_TP_ID = TRIM(EDTJRNL1.EDI_SENDER_ID)" _
        & " AND EDT852T1.EDI_STATUS = '0'; " _
        & " BEGIN FOR R1 IN C1 LOOP " _
        & " UPDATE EDT852T1 SET CUST_CODE = NVL(CUST_CODE,R1.CUST_CODE) " _
        & " WHERE EDI_DOC_SEQ_NO = R1.EDI_DOC_SEQ_NO and CUST_CODE is Null; " _
        & " END LOOP; END; END; "
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update EDT852T1 Set OPS_YYYYWW = " _
        & " (Select Min (YYYYWW) from GLTPARM3 " _
        & " where WEEK_END_DATE >= EDT852T1.EDI_FROM_DATE)" _
        & " where EDI_STATUS = '0' and OPS_YYYYWW is Null"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update EDT852T1 Set OPS_YYYYPP = " _
        & " (Select YYYYPP from GLTPARM3 " _
        & " where YYYYWW = EDT852T1.OPS_YYYYWW)" _
        & " where EDI_STATUS = '0' and OPS_YYYYPP is Null"
        ASCDATA1.ExecuteSQL()

        If opt852Data.CheckedIndex <> 0 Then
            opt852Data.CheckedIndex = 0
        Else
            Setup_EDT852T1()
        End If

    End Sub

    Private Sub grdEDT852TC_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdEDT852T2.InitializeLayout

    End Sub

    Private Sub UltraTabControl2_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab852.SelectedTabChanged

    End Sub

    Private Sub grdEDT852T1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdEDT852T1.AfterRowActivate
        Setup_EDT852T1_Details()
    End Sub

    Private Sub grdEDT852T1_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdEDT852T1.AfterRowsDeleted
        Dim sql As String = ""
        For Each EDI_DOC_SEQ_NO As String In EDI_DOC_SEQ_NOs
            sql &= ",'" & EDI_DOC_SEQ_NO & "'"
        Next
        ASCMAIN1.sql = "Update EDT852T1 Set EDI_STATUS = 'D' where EDI_DOC_SEQ_NO in (" & Mid(sql, 2) & ")"
        ASCDATA1.ExecuteSQL()
        dst.Tables("EDT852T1").AcceptChanges()
    End Sub

    Private Sub grdEDT852T1_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdEDT852T1.BeforeRowsDeleted
        EDI_DOC_SEQ_NOs.Clear()
        For Each grow As UltraWinGrid.UltraGridRow In grdEDT852T1.Selected.Rows
            EDI_DOC_SEQ_NOs.Add(grow.Cells("EDI_DOC_SEQ_NO").Text)
        Next
    End Sub

    Private Sub grdEDT852T1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdEDT852T1.InitializeLayout

    End Sub

    Sub Setup_EDT852T1()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Retrieving Previously Loaded 852 Data")

        grdEDT852T1.Text = "852 Documents (" & opt852Data.Text & ")"


        With UltraExplorerBar1.Groups("Batch Control")
            .Items("Load 852 Data").Visible = (opt852Data.Value = "0")
            .Items("Retract 852 Data").Visible = (opt852Data.Value = "1")
            .Items("Restore Deleted").Visible = (opt852Data.Value = "D")
        End With

        Dim EDI_STATUS As String = opt852Data.Value
        If opt852Data.Value = "1" Then
            grdEDT852T1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            Load_Customer_Summary()
        ElseIf opt852Data.Value = "0" Then
            grdEDT852T1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            Load_Customer_Summary()
        ElseIf opt852Data.Value = "D" Then
            grdEDT852T1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        End If

        tab852.Tabs("Items").Visible = (EDI_STATUS = "1")
        tab852.Tabs("Stores").Visible = (EDI_STATUS = "1")

        Fill_Records("EDT852T1", EDI_STATUS)
        Sort_grdColumns(grdEDT852T1, "EDI_DOC_SEQ_NO".ToLower)
        If grdEDT852T1.Rows.Count = 0 Then
            Setup_EDT852T1_Details()
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Setup_EDT852T1_Details()
        If grdEDT852T1.ActiveRow Is Nothing OrElse Not grdEDT852T1.ActiveRow.IsDataRow Then
            SplitContainer3.Panel2Collapsed = True
            Exit Sub
        End If

        SplitContainer3.Panel2Collapsed = False
        Dim EDI_DOC_SEQ_NO As String = grdEDT852T1.ActiveRow.Cells("EDI_DOC_SEQ_NO").Text
        dst.EnforceConstraints = False
        Fill_Records("EDT852T2", EDI_DOC_SEQ_NO)
        Fill_Records("EDT852T3", EDI_DOC_SEQ_NO)

        Fill_Records("EDTJRNL3", EDI_DOC_SEQ_NO)

        Dim rowEDTJRNL3 As DataRow = dst.Tables("EDTJRNL3").Rows(0)
        Dim EDI_JRNL_NO As String = rowEDTJRNL3.Item("EDI_JRNL_NO")
        Dim EDI_GS_NO As Integer = rowEDTJRNL3.Item("EDI_GS_NO")
        Fill_Records("EDTJRNL1", EDI_JRNL_NO)
        Fill_Records("EDTJRNL2", New String() {EDI_JRNL_NO, CStr(EDI_GS_NO)})
        dst.EnforceConstraints = True

        Dim rowEDTJRNL1 As DataRow = dst.Tables("EDTJRNL1").Rows(0)
        Dim EDI_FOLDERNAME As String = rowEDTJRNL1.Item("EDI_FOLDERNAME")
        Dim EDI_FILENAME As String = rowEDTJRNL1.Item("EDI_FILENAME")
        Dim FILENAME As String = EDI_FOLDERNAME & "\" & EDI_FILENAME

        Dim FI As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
        grpRawEDI.Text = FI.Name & " " & Format$(FI.LastWriteTime, "MM/dd/yy HH:mm")
        txtRawEDI.Text = ""
        Using SR As New System.IO.StreamReader(FILENAME)
            Dim RAW As String = SR.ReadToEnd
            txtRawEDI.Text = Replace(RAW, Mid(RAW, 106, 1), vbCrLf)
        End Using

        If opt852Data.Value = "1" Then
            Fill_Records("EDT852TI", EDI_DOC_SEQ_NO)
            Sort_grdColumns(grdEDT852TI, "EDI_ITEM_CODE")
            Fill_Records("EDT852TS", EDI_DOC_SEQ_NO)
            Sort_grdColumns(grdEDT852TS, "EDI_STORE_NO")
        End If

    End Sub

    Sub FIX_EDTJRNL1()

        Dim ED_PARM_RAW_INBOUND As String = ROWs("EDTPARM1").Item("ED_PARM_RAW_INBOUND") & ""

        dst.EnforceConstraints = False
        Dim Sql As String = "Select * from EDTJRNL1"
        Fill_Records("EDTJRNL1", "", True, Sql)
        For Each row As DataRow In dst.Tables("EDTJRNL1").Rows
            Dim FILENAME As String = ED_PARM_RAW_INBOUND & "\" & row.Item("EDI_FILENAME")
            row.Item("EDI_FILESIZE") = My.Computer.FileSystem.GetFileInfo(FILENAME).Length
            row.Item("EDI_DATETIME") = My.Computer.FileSystem.GetFileInfo(FILENAME).LastWriteTime
        Next
        Call Update_Record_TDA("EDTJRNL1")
        Stop

    End Sub

    Private Sub opt852Data_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles opt852Data.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_EDT852T1()
    End Sub

    Private Sub cmd852SelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd852SelectAll.Click
        For Each rowEDT852T1 As DataRow In dst.Tables("EDT852T1").Rows
            rowEDT852T1.Item("SELECTED") = "1"
        Next
    End Sub

    Private Sub cmd852DeSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd852DeSelectAll.Click
        For Each rowEDT852T1 As DataRow In dst.Tables("EDT852T1").Rows
            rowEDT852T1.Item("SELECTED") = "0"
        Next
    End Sub

    Sub Load_852_Data()

        dst.Tables("EDT852T1").AcceptChanges()

        BeginTrans()

        For Each rowEDT852T1 As DataRow In dst.Tables("EDT852T1").Select("SELECTED = '1'", "EDI_DOC_SEQ_NO")
            Dim EDI_DOC_SEQ_NO As String = rowEDT852T1.Item("EDI_DOC_SEQ_NO")
            ASCMAIN1.Progress("Now Processing Document " & EDI_DOC_SEQ_NO, "")

            If opt852Data.Value <> "D" Then
                If opt852Data.Value = "0" Then

                    ASCDATA1.ExecuteSQL("Delete from EDT852T0 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
                    Dim SQL As String = ""
                    For i As Integer = 1 To 10
                        Dim z As String = Format$(i, "00")
                        SQL = "Insert INTO EDT852T0" & vbCrLf _
                        & " (EDI_DOC_SEQ_NO, EDI_LINE_NO, EDI_ITEM_CODE, EDI_TRAN_TYPE " & vbCrLf _
                        & ", EDI_STORE_NO, EDI_QTY, CUST_CODE, CUST_STORE_NO)" & vbCrLf _
                        & " Select EDT852T2.EDI_DOC_SEQ_NO, EDT852T2.EDI_LINE_NO" & vbCrLf _
                        & ", NVL(EDT852T2.EDI_ITEM_UP, EDT852T2.EDI_ITEM_EN) EDI_ITEM_CODE" & vbCrLf _
                        & ", EDT852T3.EDI_ZA_TRAN_TYPE EDI_TRAN_TYPE" & vbCrLf _
                        & ", EDT852T3.EDI_SDQ_STORE_" & z & " EDI_STORE_NO" & vbCrLf _
                        & ", EDT852T3.EDI_SDQ_QTY_AMT_" & z & " EDI_QTY" & vbCrLf _
                        & ", EDT852T1.CUST_CODE" & vbCrLf _
                        & ", LPAD(EDT852T3.EDI_SDQ_STORE_" & z & ",6,'0') CUST_STORE_NO" & vbCrLf _
                        & " from EDT852T3, EDT852T2, EDT852T1" & vbCrLf _
                        & " where EDT852T2.EDI_DOC_SEQ_NO = EDT852T1.EDI_DOC_SEQ_NO" & vbCrLf _
                        & "   and EDT852T3.EDI_DOC_SEQ_NO = EDT852T2.EDI_DOC_SEQ_NO" & vbCrLf _
                        & "   and EDT852T3.EDI_LINE_NO = EDT852T2.EDI_LINE_NO" & vbCrLf _
                        & "   and EDT852T3.EDI_ZA_TRAN_TYPE IN ('QS','QU')" & vbCrLf _
                        & "   and EDT852T1.EDI_STATUS = '0'" & vbCrLf _
                        & "   and EDT852T1.EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" & vbCrLf _
                        & "   and EDT852T3.EDI_SDQ_STORE_" & z & " IS NOT NULL" & vbCrLf _
                        & "   and NVL(EDT852T3.EDI_SDQ_QTY_AMT_" & z & ",0) <> 0"
                        ASCDATA1.ExecuteSQL(SQL)
                        'SQL = SQL & " AND EDT852T4.EDI_SDQ_TYPE ='EA'"
                    Next i
                    Set_ITEM_CODE(EDI_DOC_SEQ_NO)
                End If

                Update_RSTRETL1(EDI_DOC_SEQ_NO)
            End If

            Dim EDI_STATUS As String = "1"
            If opt852Data.Value = "1" Or opt852Data.Value = "D" Then EDI_STATUS = "0"

            ASCMAIN1.sql = "Update EDT852T1 set EDI_STATUS = '" & EDI_STATUS & "'" _
            & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
            ASCDATA1.ExecuteSQL()

            'rowEDT852T1.Item("EDI_STATUS") = "1"
        Next

        'Update_Record_TDA("EDT852T1")
        ASCMAIN1.Progress("", "")
        CommitTrans("Update Complete")

        Prepare_852_Queue()

    End Sub

    Sub Set_ITEM_CODE(ByVal EDI_DOC_SEQ_NO As String)
        ASCMAIN1.sql = "" _
        & " BEGIN DECLARE CURSOR C1 IS" _
        & " SELECT ITEM_UPC_CODE EDI_ITEM_CODE, ITEM_CODE " _
        & " FROM ICTITEM1 WHERE ITEM_UPC_CODE IN (" _
        & " SELECT DISTINCT EDI_ITEM_CODE FROM EDT852T0 " _
        & " WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" _
        & " AND EDI_ITEM_CODE IS NOT NULL)" _
        & " UNION" _
        & " SELECT ITEM_EAN_CODE EDI_ITEM_CODE, ITEM_CODE " _
        & " FROM ICTITEM1 WHERE ITEM_EAN_CODE IN (" _
        & " SELECT DISTINCT EDI_ITEM_CODE FROM EDT852T0 " _
        & " WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" _
        & " AND EDI_ITEM_CODE IS NOT NULL);" _
        & " BEGIN FOR R1 IN C1 LOOP" _
        & " UPDATE EDT852T0 SET ITEM_CODE = R1.ITEM_CODE" _
        & " WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" _
        & " AND EDI_ITEM_CODE = R1.EDI_ITEM_CODE;" _
        & " END LOOP; END; END;"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_RSTRETL1(ByVal EDI_DOC_SEQ_NO As String)

        If opt852Data.Value = "1" Then
            Update_RSTRETLx(EDI_DOC_SEQ_NO)

            ASCMAIN1.sql = "Delete from RSTRETL1 " _
            & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
            ASCDATA1.ExecuteSQL()
        Else
            ASCMAIN1.sql = "" _
            & " Insert INTO RSTRETL1" _
            & " Select EDT852T0.EDI_DOC_SEQ_NO" _
            & " , EDT852T0.CUST_CODE, EDT852T0.CUST_STORE_NO" _
            & " , EDT852T0.ITEM_CODE" _
            & " , SUM (DECODE(EDI_TRAN_TYPE,'QS',NVL(EDI_QTY,0),'QU',-1 * NVL(EDI_QTY,0),0)) QTY_SOLD" _
            & " , SUM (DECODE(EDI_TRAN_TYPE,'QS',NVL(EDI_QTY,0),'QU',-1 * NVL(EDI_QTY,0),0) * NVL(ICTITEM1.ITEM_RETAIL_PRICE,0)) AMT_SOLD" _
            & " , EDT852T1.OPS_YYYYPP, EDT852T1.OPS_YYYYWW" _
            & " , SUM (DECODE(EDI_TRAN_TYPE,'QA',EDI_QTY)) QTY_SOLD" _
            & " from EDT852T0,EDT852T1,ICTITEM1" _
            & " where EDT852T0.EDI_DOC_SEQ_NO = EDT852T1.EDI_DOC_SEQ_NO" _
            & " and EDT852T0.ITEM_CODE is Not Null" _
            & " and ICTITEM1.ITEM_CODE = EDT852T0.ITEM_CODE" _
            & " and EDT852T0.EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" _
            & " group by EDT852T0.EDI_DOC_SEQ_NO" _
            & " , EDT852T0.CUST_CODE, EDT852T0.CUST_STORE_NO" _
            & " , EDT852T0.ITEM_CODE" _
            & " , EDT852T1.OPS_YYYYPP, EDT852T1.OPS_YYYYWW"
            ASCDATA1.ExecuteSQL()

            Update_RSTRETLx(EDI_DOC_SEQ_NO)
        End If
    End Sub

    Private Sub chkRawProcessed_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRawPreviouslyImported.CheckedChanged
        Setup_tabRaw()
    End Sub

    Sub Setup_tabRaw()
        SplitContainer2.Panel2Collapsed = Not (chkRawPreviouslyImported.Checked)
        If Not chkRawPreviouslyImported.Checked Then
            dst.Tables("EDTFILE1").Rows.Clear()

            grdEDTFILE1.DisplayLayout.GroupByBox.Hidden = True
            Show_Filter(grdEDTFILE1, False)

        Else

            grdEDTFILE1.DisplayLayout.GroupByBox.Hidden = False
            Show_Filter(grdEDTFILE1, True)

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Retrieving Files Processed")

            Dim SQL As String = "Select EDTJRNL1.EDI_FILENAME" _
            & ", MIN (EDTJRNL1.EDI_FILESIZE) EDI_FILESIZE" _
            & ", MIN (EDTJRNL1.EDI_DATETIME) EDI_DATETIME" _
            & ", MIN (EDTJRNL1.EDI_JRNL_NO) EDI_JRNL_NO" _
            & ", MIN (EDTJRNL1.EDI_SENDER_QUAL) EDI_SENDER_QUAL" _
            & ", MIN (EDTJRNL1.EDI_SENDER_ID) EDI_SENDER_ID" _
            & ", MIN (EDTJRNL1.EDI_ISA_CTL_NO) EDI_ISA_CTL_NO" _
            & ", MIN (EDTJRNL1.EDI_ISA_CTL_DATE) EDI_ISA_CTL_DATE" _
            & ", COUNT (EDTJRNL3.EDI_DOC_NO) DOC_EDI" _
            & ", COUNT (EDTJRNL3.EDI_DOC_SEQ_NO) DOC_852, NULL NOTES from EDTJRNL1,EDTJRNL3" _
            & " where EDTJRNL1.EDI_JRNL_NO = EDTJRNL3.EDI_JRNL_NO (+)" _
            & " group by EDTJRNL1.EDI_FILENAME"
            Fill_Records("EDTFILE1", "", True, SQL)
            Sort_grdColumns(grdEDTFILE1, "EDI_FILENAME")

            If grdEDTFILE1.ActiveRow Is Nothing Then
                tabRaw.Visible = False
            Else
                tabRaw.Visible = True
                Setup_EDTFILE1_Details()
            End If

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")

        End If
    End Sub

    Private Sub grdEDTFILE1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdEDTFILE1.AfterRowActivate
        If grdEDTFILE1.ActiveRow.IsDataRow Then
            Setup_EDTFILE1_Details()
        End If
    End Sub

    Private Sub grdEDTFILE1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdEDTFILE1.InitializeLayout

    End Sub

    Sub Setup_EDTFILE1_Details()

        Dim ED_PARM_RAW_INBOUND As String = ROWs("EDTPARM1").Item("ED_PARM_RAW_INBOUND") & ""

        Dim FILENAME As String = ED_PARM_RAW_INBOUND & "\" & grdEDTFILE1.ActiveRow.Cells("EDI_FILENAME").Text

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data for file")

        Dim FI As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
        grpRawEDI.Text = FI.Name & " " & Format$(FI.LastWriteTime, "MM/dd/yy HH:mm")
        txtRawEDI.Text = ""
        Using SR As New System.IO.StreamReader(FILENAME)
            Dim RAW As String = SR.ReadToEnd
            txtRawEDI.Text = Replace(RAW, Mid(RAW, 106, 1), vbCrLf)
        End Using

        Dim EDI_JRNL_NO As String = grdEDTFILE1.ActiveRow.Cells("EDI_JRNL_NO").Text
        Dim sql As String = ""


        dst.EnforceConstraints = False

        'Fill_Records("EDT852T2", EDI_DOC_SEQ_NO)
        'Fill_Records("EDT852T3", EDI_DOC_SEQ_NO)

        sql = "Select EDT852T1.*, ARTCUST1.CUST_NAME, '0' SELECTED " _
        & " from EDT852T1,ARTCUST1 " _
        & " where EDT852T1.EDI_DOC_SEQ_NO in (Select EDI_DOC_SEQ_NO FROM EDTJRNL1,EDTJRNL3 where EDTJRNL1.EDI_JRNL_NO = EDTJRNL3.EDI_JRNL_NO and EDTJRNL1.EDI_JRNL_NO = '" & EDI_JRNL_NO & "')" _
        & " and ARTCUST1.CUST_CODE (+) = EDT852T1.CUST_CODE"
        Fill_Records("EDT852T1", "", True, sql)

        Fill_Records("EDTJRNL1", EDI_JRNL_NO)

        sql = "Select EDTJRNL2.* from EDTJRNL2 " _
        & " where EDTJRNL2.EDI_JRNL_NO = '" & EDI_JRNL_NO & "'"
        Fill_Records("EDTJRNL2", "", True, sql)

        sql = "Select EDTJRNL3.* from EDTJRNL3 " _
        & " where EDTJRNL3.EDI_JRNL_NO = '" & EDI_JRNL_NO & "'"
        Fill_Records("EDTJRNL3", "", True, sql)

        dst.EnforceConstraints = True

        Sort_grdColumns(grdEDT852T1, "EDI_DOC_SEQ_NO")

        If grdEDTJRNL1.Rows.Count > 0 Then
            grdEDTJRNL1.Rows(0).ExpandAll()
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Load_Customer_Summary()

        Dim sql As String = ""
        Dim W As Integer = 0
        Dim RYP As String = ASCMAIN1.CYP
        With grdEDT852TC.DisplayLayout.Bands(0)
            If .Groups.Count <> 0 Then
                For g As Integer = .Groups.Count - 1 To 0 Step -1
                    .Groups.Remove(g)
                Next
            End If
            .Groups.Add("CUST_CODE")
            .Groups("CUST_CODE").Header.Caption = ""
            .Columns("CUST_CODE").Group = .Groups("CUST_CODE")
            For P As Integer = 2 To 0 Step -1
                Dim YP As String = ASCMAIN1.Period_Calc(RYP, -1 * P)
                .Groups.Add(YP)
                Dim LEGEND As String = ASCMAIN1.Get_Legend(YP)
                .Groups(YP).Header.Caption = Mid(LEGEND, 10, 6)
                .Groups(YP).Header.Appearance.BackColor = Drawing.Color.Yellow
                .Groups(YP).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                ASCMAIN1.sql = "Select YYYYWW from GLTPARM3 where YYYYPP = '" & YP & "'"
                For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "YYYYWW")
                    W += 1
                    Dim COLUMN_NAME As String = "F" & Format(W, "00")
                    .Columns(COLUMN_NAME).Group = .Groups(YP)
                    Dim YW As String = row.Item("YYYYWW")
                    .Columns(COLUMN_NAME).Tag = YW
                    sql &= ", Sum (Decode(OPS_YYYYWW,'" & YW & "',1,0)) " & COLUMN_NAME

                    If P Mod 2 = 0 Then
                        .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightBlue
                    Else
                        .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightPink
                    End If
                    .Columns(COLUMN_NAME).Hidden = False
                    .Columns(COLUMN_NAME).Header.Caption = Mid(YW, 5, 2)
                    .Columns(COLUMN_NAME).Width = 30
                Next
            Next
            If W < 27 Then
                For I As Integer = W + 1 To 27
                    Dim COLUMN_NAME As String = "F" & Format(I, "00")
                    .Columns(COLUMN_NAME).Hidden = True
                Next
            End If
        End With

        sql = "Select CUST_CODE" & sql _
        & " from EDT852T1 " _
        & " where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(RYP, -5) & "'" _
        & " and OPS_YYYYPP <= '" & RYP & "'" _
        & " and EDI_STATUS = '1'" _
        & " group by CUST_CODE"
        Fill_Records("EDT852TC", "", True, sql)
        Sort_grdColumns(grdEDT852TC, "CUST_CODE")
    End Sub

    Private Sub grdEDT852TI_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdEDT852TI.InitializeRow
        If e.Row.Cells("ITEM_CODE").Text = "" Then
            e.Row.Cells("EDI_ITEM_CODE").Appearance.ForeColor = Drawing.Color.Red
        End If
    End Sub

    Private Sub grdEDT852TS_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdEDT852TS.InitializeRow
        If e.Row.Cells("CUST_STORE_NO").Text = "" Then
            e.Row.Cells("EDI_STORE_NO").Appearance.ForeColor = Drawing.Color.Red
        End If
    End Sub

    Private Sub optMode_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optMode.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Select_tabMain()
    End Sub

    Sub Select_tabMain()
        'tabMain.VisibleTab
        tabMain.Tabs(optMode.Value).visible = True
        tabMain.SelectedTab = tabMain.Tabs(optMode.Value)
        'tabMain.SelectedTab.Visible = True
        For Each TAB As UltraWinTabControl.UltraTab In tabMain.Tabs
            If Not TAB.Selected Then
                TAB.Visible = False
            End If
        Next
    End Sub

    Sub Update_RSTRETLx(ByVal EDI_DOC_SEQ_NO As String)

        Dim plus_or_minus As String = "+"
        If opt852Data.Value = "1" Then
            plus_or_minus = "-"
        End If

        Dim sql As String = ""

        sql = "BEGIN DECLARE CURSOR C1 IS"
        sql = sql & " SELECT RSTRETL1.OPS_YYYYPP, RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO,"
        sql = sql & "  ICTITEM1.COLLECTION_CODE, SUM (RSTRETL1.AMT_SOLD) AS RETAIL_SALES"
        sql = sql & "  From RSTRETL1, ICTITEM1"
        sql = sql & "  Where RSTRETL1.ITEM_CODE = ICTITEM1.ITEM_CODE (+) "
        sql = sql & "    and EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
        sql = sql & "  GROUP BY RSTRETL1.OPS_YYYYPP, RSTRETL1.CUST_CODE, "
        sql = sql & "           RSTRETL1.CUST_STORE_NO, ICTITEM1.COLLECTION_CODE"
        sql = sql & "  HAVING SUM (RSTRETL1.AMT_SOLD) <> 0;"
        sql = sql & " BEGIN FOR R1 IN C1 LOOP"
        sql = sql & " Update RSTRETL2"
        sql = sql & " Set RETAIL_SALES = NVL(RETAIL_SALES, 0) " & plus_or_minus & " NVL(R1.RETAIL_SALES,0)"
        sql = sql & "  Where OPS_YYYYPP = R1.OPS_YYYYPP And CUST_CODE = R1.CUST_CODE      "
        sql = sql & "    AND CUST_STORE_NO = R1.CUST_STORE_NO AND COLLECTION_CODE = R1.COLLECTION_CODE;"
        sql = sql & " IF SQL%NOTFOUND THEN"
        sql = sql & "  INSERT INTO RSTRETL2 VALUES"
        sql = sql & "   (R1.OPS_YYYYPP, R1.CUST_CODE, R1.CUST_STORE_NO, R1.COLLECTION_CODE, " & plus_or_minus & "1 * " & "NVL(R1.RETAIL_SALES,0));"
        sql = sql & " END IF;"
        sql = sql & " END LOOP; END; END;"
        ASCDATA1.ExecuteSQL(sql)

        sql = "BEGIN DECLARE CURSOR C1 IS" _
        & "  SELECT RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, ICTITEM1.COLLECTION_CODE, " _
        & "   ICTCOLL1.BRAND_CODE," _
        & "   NULL SREP_CODE, NULL SELL_CODE," _
        & "   RSTRETL1.OPS_YYYYWW, RSTRETL1.OPS_YYYYPP" _
        & ",  SUM(RSTRETL1.AMT_SOLD) AS AMT_SOLD" _
        & ",  SUM(RSTRETL1.QTY_SOLD) AS QTY_SOLD" _
        & ",  SUM(RSTRETL1.QTY_EOW * ICTITEM1.ITEM_RETAIL_PRICE) QTY_EOW" _
        & "   From RSTRETL1, ICTITEM1, ICTCOLL1" _
        & "   Where RSTRETL1.ITEM_CODE = ICTITEM1.ITEM_CODE (+) " _
        & "   AND ICTITEM1.COLLECTION_CODE = ICTCOLL1.COLLECTION_CODE (+)" _
        & "    and EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" _
        & "   GROUP BY RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, ICTITEM1.COLLECTION_CODE, " _
        & "   ICTCOLL1.BRAND_CODE, " _
        & "   NULL, NULL, RSTRETL1.OPS_YYYYWW, RSTRETL1.OPS_YYYYPP" _
        & "   HAVING SUM (RSTRETL1.AMT_SOLD) <> 0 OR SUM (RSTRETL1.QTY_SOLD) <> 0 OR SUM(QTY_EOW) <> 0;" _
        & "  BEGIN FOR R1 IN C1 LOOP" _
        & "  Update RSTRETL4" _
        & " Set AMT_SOLD = NVL(AMT_SOLD, 0) " & plus_or_minus & " R1.AMT_SOLD," _
        & "    QTY_SOLD = NVL(QTY_SOLD, 0) " & plus_or_minus & " R1.QTY_SOLD," _
        & "    QTY_EOW = NVL(QTY_EOW, 0) " & plus_or_minus & " R1.QTY_EOW" _
        & "   Where  CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO AND COLLECTION_CODE = R1.COLLECTION_CODE" _
        & "     And OPS_YYYYWW = R1.OPS_YYYYWW And OPS_YYYYPP = R1.OPS_YYYYPP;" _
        & "  IF SQL%NOTFOUND THEN" _
        & "   INSERT INTO RSTRETL4 VALUES" _
        & "    (R1.CUST_CODE, R1.CUST_STORE_NO, R1.COLLECTION_CODE, R1.OPS_YYYYWW, R1.OPS_YYYYPP, " _
        & plus_or_minus & "1 * R1.QTY_SOLD, " & plus_or_minus & "1 * R1.AMT_SOLD, R1.BRAND_CODE, R1.SELL_CODE, R1.SREP_CODE, R1.QTY_EOW);" _
        & "  END IF;" _
        & "  END LOOP; END; END;"
        ASCDATA1.ExecuteSQL(sql)

    End Sub

    Private Sub grdEDT852TC_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdEDT852TC.DoubleClickCell
        Try
            opt852Data.Value = "1"
            Dim OPS_YYYYWW As String = e.Cell.Column.Tag
            Dim CUST_CODE As String = e.Cell.Row.Cells("CUST_CODE").Text
            With grdEDT852T1.DisplayLayout.Bands(0)
                .ColumnFilters.ClearAllFilters()
                .ColumnFilters("CUST_CODE").FilterConditions.Add(Infragistics.Win.UltraWinGrid.FilterComparisionOperator.Match, CUST_CODE)
                .ColumnFilters("OPS_YYYYWW").FilterConditions.Add(Infragistics.Win.UltraWinGrid.FilterComparisionOperator.Match, OPS_YYYYWW)

                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
                tlb_sbt.Checked = True
                Show_Filter(grdEDT852T1, True)

            End With
        Catch ex As Exception

        End Try
    End Sub

    Private Sub grdEDT852TC_InitializeLayout_1(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdEDT852TC.InitializeLayout

    End Sub
End Class