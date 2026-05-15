Imports GemBox.Spreadsheet

Public Class TAFCONVC
    Dim grdX2Os() As UltraWinGrid.UltraGrid

    Dim excelFileNum As Integer
     Dim loaded_once As Boolean = False
    Dim cfg_folder As String

    Dim dt_errors As New DataTable
    Dim INCR_FROMs As New Dictionary(Of String, String)
    Dim skip_for_now = True

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'IF WE GET A NEW LAYOUT, THIS IS HOW YOU GET THIS FORM TO RECREATE THAT TABLE
        ' NOTE: THIS ONLY WORKS FOR TABLES WHERE WE GET ALL OF THE DATA, ALL OF THE TIME
        '       IT DOES NOT WORK FOR INCREMENTALS - WHERE I USUALLY DO THESE COMMANDS ON MY LAPTOP AND THEN MANUALLY PORT OVER RECORDS AND DDL TO THE CONV SCHEMA

        'DROP TABLE CONV.CFG_CUSTMAST;
        'DELETE FROM CONV.CFG_LAYOUT where FILE_AS400 = (SELECT FILE_AS400 FROM CONV.CFG_FILE WHERE FILENAME = 'custmast');
        'DELETE FROM CONV.CFG_FILE WHERE FILENAME = 'custmast';

        With dst
            ASCMAIN1.sql = "Select * from CONV.CFG_FILE"
            Create_TDA(.Tables.Add("TATCONVX"), "CONV.CFG_FILE", "**", 0, True)

            ASCMAIN1.sql = "Select * from CONV.CFG_LAYOUT"
            Create_TDA(.Tables.Add("TATCONVY"), "CONV.CFG_LAYOUT", "**", 0, True)

            Create_Relation("TATCONVX", "TATCONVY", "FILE_AS400")
        End With

        grdTATCONVX.DataSource = dst.Tables("TATCONVX")

        Create_Summary(grdTATCONVX, "FILENAME", "Count")
        Create_Summary(grdTATCONVX, "SECS")

        Fill_Records("TATCONVX")
        Fill_Records("TATCONVY")

        If ASCMAIN1.Running_in_VS Then
            cfg_folder = "C:\Users\wjz\Desktop\CFG"
        Else
            cfg_folder = ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT") & "\CFG"
        End If


    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load CSVs"
                If loaded_once Then
                    EMsg &= vbCr & "You may load only once.  Exit to the Menu and then re-execute."
                End If

                If EMsg = "" Then

                    INCR_FROMs.Clear()
                    For Each ROW As DataRow In dst.Tables("TATCONVX").Select("INCR_FROM IS NOT NULL")
                        Dim FILENAME As String = ROW.Item("FILENAME").ToString.ToUpper
                        Dim INCR_FROM As String = ROW.Item("INCR_FROM")
                        INCR_FROMs.Add(FILENAME, INCR_FROM)
                    Next

                    Dim FILES As New List(Of String)
                    For Each FILE As String In My.Computer.FileSystem.GetFiles(cfg_folder)
                        FILES.Add(FILE)
                    Next

                    If FILES.Count = 0 Then
                        EMsg &= vbCr & "No Files to Import"
                    Else

                        If tabX2O.Tabs.Count > 1 Then
                            For T As Int16 = tabX2O.Tabs.Count - 1 To 1
                                tabX2O.Tabs.Remove(tabX2O.Tabs(T))
                            Next
                        End If

                        ReDim grdX2Os(FILES.Count - 1)
                        Dim SHEET_NO As Integer = 0
                        ASCMAIN1.Progress("Now Loading", "")

                        For Each FILENAME As String In FILES

                            Load_Workbook(FILENAME, SHEET_NO)
                        Next
                        ASCMAIN1.Progress("", "")
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

            Case "Load CSVs"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)


            Case "Run Scripts"
                'SQLPLUS INT/INT @INT_CFG_CONVERT.SQL > INT_CFG_CONVERT.LOG
                'SQLPLUS INT/INT @INT_GP_CONVERT.SQL > INT_GP_CONVERT.LOG

                Dim CMD As String = ""

                Dim FILENAME As String = ASCMAIN1.Folders("Work") & "\INT_CONVERT.BAT"
                Using sw As New System.IO.StreamWriter(FILENAME)

                    'CMD = "SQLPLUS " & ASCMAIN1.DBS_COMPANY & "/" & ASCMAIN1.DBS_PASSWORD & "@INT @S:\INT\INT_GP_CONVERT.SQL > S:\INT\ARCHIVE\" & ASCMAIN1.DBS_COMPANY & "\INT_GP_CONVERT.LOG"
                    CMD = "SQLPLUS " & ASCMAIN1.DBS_COMPANY & "/" & ASCMAIN1.DBS_PASSWORD & "@INT @" & ASCMAIN1.Folders("SharedRoot") & "\INT_GP_CONVERT.SQL > " & ASCMAIN1.Folders("SharedRoot") & "\ARCHIVE\" & ASCMAIN1.DBS_COMPANY & "\INT_GP_CONVERT.LOG"
                    sw.WriteLine(CMD)
                    'CMD = "SQLPLUS " & ASCMAIN1.DBS_COMPANY & "/" & ASCMAIN1.DBS_PASSWORD & "@INT @S:\INT\INT_CFG_CONVERT.SQL > S:\INT\ARCHIVE\" & ASCMAIN1.DBS_COMPANY & "\INT_CFG_CONVERT.LOG"
                    CMD = "SQLPLUS " & ASCMAIN1.DBS_COMPANY & "/" & ASCMAIN1.DBS_PASSWORD & "@INT @" & ASCMAIN1.Folders("SharedRoot") & "\INT_CFG_CONVERT.SQL > " & ASCMAIN1.Folders("SharedRoot") & "\ARCHIVE\" & ASCMAIN1.DBS_COMPANY & "\INT_CFG_CONVERT.LOG"
                    sw.WriteLine(CMD)
                End Using

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Running Data Conversion Scripts")
                Show_Document(FILENAME)

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Load CSVs").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            .Groups("Data Controls").Visible = False 'ScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdTATCONVX.Visible = Not ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading")
        Me.Cursor = Cursors.WaitCursor

        Update_Record_TDA("TATCONVX")
        Update_Record_TDA("TATCONVY")

        'ASCMAIN1.sql = "Delete from CONV.CFG_SELLTHRU where TRIM(SLPRD) LIKE '1201%' or TRIM(SLPRD) LIKE '1202%'"
        'ASCDATA1.ExecuteSQL()
        'ASCMAIN1.sql = "Insert into CONV.CFG_SELLTHRU Select * from CONV.CFG_IPLB121"
        'ASCDATA1.ExecuteSQL()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Clear_Record()

        'For Each TABLE_NAME As String In New String() {"TATCONVY"}
        '    dst.Tables(TABLE_NAME).Rows.Clear()
        'Next

        'If dst.Tables("TATCONVX").Rows.Count = 0 Then
        '    ASCMAIN1.sql = "Select * from TATCONVX"
        '    Fill_Records("TATCONVX", "", , ASCMAIN1.sql)
        '    Sort_grdColumns(grdTATCONVX, "ODOBNM")
        'End If

        Application.DoEvents()
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()

        Load_Popup_Menu(grdTATCONVX, "SSS", "Show GroupBox", "Show Filter", "Show Pins")

    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        Select Case grd.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case Else
                'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select
    End Sub

#End Region


    Sub Load_Workbook(ByVal FILENAME As String, ByRef SHEET_NO As Integer)

        SpreadsheetInfo.SetLicense(ASCMAIN1.GemboxKey) ' ("EMPX-L9BW-EL8E-4GKJ") ' ("EFYZ-QQSH-LE5Q-NJ7Y")

        loaded_once = True

        Dim INIT_DATE As DateTime
        Dim LAST_DATE As DateTime

        Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
        Dim fn As String = fi.Name
        Dim ROWCOUNT As Int64 = 0

        Dim TABLE_NAME As String = ""

        If fn.EndsWith(".csv") Then
            If fn.Contains("_") Then
                TABLE_NAME = Mid(fn, 1, InStr(fn, "_") - 1)
            Else
                Throw New Exception("Invalid File Name - no underscore separating File Name from Date")
            End If
        Else
            Throw New Exception("Invalid File Type - not CSV")
        End If

        ASCMAIN1.Progress("-", TABLE_NAME)

        Dim rowCFG_FILE As DataRow = dst.Tables("TATCONVX").Rows.Find(TABLE_NAME)

        If rowCFG_FILE Is Nothing Then

            Dim DDL As String = ""
            Dim DDLKEY As String = ""
            Dim ef2 As ExcelFile = New ExcelFile
            ef2 = ExcelFile.Load(fi.DirectoryName & "\layouts\" & TABLE_NAME & ".xls", LoadOptions.XlsDefault)

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

            'If INCR_FROMs.ContainsKey(TABLE_NAME.ToUpper) Then
            '    Stop
            'End If
            'If TABLE_NAME.ToUpper = "ORDRHEAD" Or TABLE_NAME.ToUpper = "ORDRDETL" Or TABLE_NAME.ToUpper = "SCRORDDTL" Or TABLE_NAME.ToUpper = "SCRORDHED" Then
            '    Stop
            'End If

            ASCMAIN1.sql = "DROP TABLE CONV.CFG_" & TABLE_NAME
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, True)

            DDLKEY = ""

            DDL = "CREATE TABLE CONV.CFG_" & TABLE_NAME & " (" & Mid(DDL, 2) & IIf(DDLKEY <> "", ", PRIMARY KEY (" & Mid(DDLKEY, 2) & ")", "") & ")"
            ASCDATA1.ExecuteSQL(DDL)
        End If

        SHEET_NO += 1
        Dim grd As UltraWinGrid.UltraGrid

        If SHEET_NO = 1 Then
            grd = grdX2O
            grd.DataSource = Nothing
        Else
            tabX2O.Tabs.Add(New UltraWinTabControl.UltraTab)
            grd = New UltraWinGrid.UltraGrid
            grd.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.SortMulti
            grd.Parent = tabX2O.Tabs(SHEET_NO - 1).TabPage
            grd.Visible = True
            grd.Dock = DockStyle.Fill
        End If

        tabX2O.Tabs(SHEET_NO - 1).Text = TABLE_NAME
        tabX2O.Tabs(SHEET_NO - 1).Tag = TABLE_NAME

        grdX2Os(SHEET_NO - 1) = grd

        ASCMAIN1.sql = "Select * from CONV.CFG_" & TABLE_NAME

        If dst.Tables.Contains(TABLE_NAME) Then
            dst.Tables(TABLE_NAME).Rows.Clear()
        Else
            Create_TDA(dst.Tables.Add(TABLE_NAME), "CONV.CFG_" & TABLE_NAME, "**", 0, True)
        End If
        Dim dt As DataTable = dst.Tables(TABLE_NAME)
 
        If INCR_FROMs.ContainsKey(TABLE_NAME.ToUpper) Then
            '    If TABLE_NAME.ToUpper = "ORDRHEAD" Or TABLE_NAME.ToUpper = "ORDRDETL" Or TABLE_NAME.ToUpper = "SCRORDDTL" Or TABLE_NAME.ToUpper = "SCRORDHED" Then
            ' DO NOTHING
            Dim MMDDYY As String = Split(Split(fn, "_")(1), ".")(0)
            If MMDDYY.Length <> 6 OrElse Mid(MMDDYY, 5, 2) <> "15" Then Stop
            Dim INCR_FROM As String = INCR_FROMs(TABLE_NAME.ToUpper)
            'If MMDDYY < INCR_FROM Then
            '    Stop
            '    ASCDATA1.ExecuteSQL("Truncate Table CONV.CFG_" & TABLE_NAME)
            'End If

            ASCDATA1.ExecuteSQL("Truncate Table CONV.CFG_" & TABLE_NAME)
            If skip_for_now Then
            Else
                ASCDATA1.ExecuteSQL("Insert into CONV.CFG_" & TABLE_NAME & " Select * from CONV.CFG_" & TABLE_NAME & "_1020")
            End If
        Else
            ASCDATA1.ExecuteSQL("Truncate Table CONV.CFG_" & TABLE_NAME)
        End If

        INIT_DATE = Now
        rowCFG_FILE.Item("INIT_DATE") = INIT_DATE

        Dim use_binding As Boolean = True
        Create_BAs(TABLE_NAME)

        Using sr As New System.IO.StreamReader(FILENAME)
            Do
                Dim lin As String = sr.ReadLine
                Dim d() As String = Split(lin, ",")

                If TABLE_NAME = "ordrhead" Then
                    If d(1) = Chr(34) & "6706150" & Chr(34) AndAlso d.Length = 15 Then
                        For a As Integer = 1 To 2
                            Dim lin2 As String = sr.ReadLine
                            lin &= lin2
                        Next
                        d = Split(lin, ",")
                    End If
                End If

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
                            If InStr(d(i), Chr(34) & Chr(34)) <> 0 Then
                                d(i) = Replace(d(i), Chr(34) & Chr(34), Chr(34))
                            End If
                            ' If i = 1 And Len(d(i)) > 30 Then Stop
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

                If TABLE_NAME = "shpviascac" AndAlso Len(r.Item(2)) > 30 Then
                    If ASCMAIN1.Running_in_VS Then Stop
                    r.Item(2) = Mid(r.Item(2), 1, 30)
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

        If use_binding Then
            Update_BAs(TABLE_NAME)
        Else
            Update_Record_TDA(TABLE_NAME)
        End If

        rowCFG_FILE.Item("ROWCOUNT") = ROWCOUNT
        LAST_DATE = Now
        rowCFG_FILE.Item("LAST_DATE") = LAST_DATE
        rowCFG_FILE.Item("SECS") = LAST_DATE.Subtract(INIT_DATE).Seconds

        If TABLE_NAME.ToUpper = "ITEMMAST" Then
            ASCMAIN1.sql = "Insert into CONV.CFG_ITEMMAST Select * from CONV.CFG_ITEMMAST_SAVE where (ITCONO,ITITEM) in (SELECT ITCONO, ITITEM FROM CONV.CFG_ITEMMAST_SAVE MINUS SELECT ITCONO, ITITEM FROM CONV.CFG_ITEMMAST)"
            ASCDATA1.ExecuteSQL()
        End If

        If INCR_FROMs.ContainsKey(TABLE_NAME.ToUpper) Then
 
            If INCR_FROMs(TABLE_NAME.ToUpper) = "000000" And Not skip_for_now Then
                ASCDATA1.ExecuteSQL("Drop Table CONV.CFG_" & TABLE_NAME & "_MERGE", True)
                ASCDATA1.ExecuteSQL("Create Table CONV.CFG_" & TABLE_NAME & "_MERGE as Select Distinct X.* from CONV.CFG_" & TABLE_NAME & " X")
                ASCDATA1.ExecuteSQL("Truncate Table CONV.CFG_" & TABLE_NAME)
                ASCDATA1.ExecuteSQL("Insert into CONV.CFG_" & TABLE_NAME & " Select * from CONV.CFG_" & TABLE_NAME & "_MERGE")
            End If
        End If


        grd.DataSource = dt
        ASCMAIN1.grdInitializeLayout(grd)
        grd.DisplayLayout.GroupByBox.Hidden = False
        Show_Filter(grd, True)

        For I As Int16 = 0 To dt.Columns.Count - 1
            grd.DisplayLayout.Bands(0).Columns(I).Header.Appearance.TextHAlign = HAlign.Center
        Next
        With grd.DisplayLayout.Override
            .RowSelectors = DefaultableBoolean.True
            .RowSelectorNumberStyle = UltraWinGrid.RowSelectorNumberStyle.VisibleIndex
            .RowSelectorHeaderStyle = UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
            .RowSelectorAppearance.TextHAlign = HAlign.Center
        End With

        grd.DisplayLayout.CaptionVisible = DefaultableBoolean.True
        grd.DisplayLayout.CaptionAppearance.TextHAlign = HAlign.Left
        grd.Text = FILENAME

        My.Computer.FileSystem.MoveFile(FILENAME, fi.Directory.ToString & "\archive\" & fi.Name)

    End Sub
    Function Excel_Column(ByVal i As Int16) As String
        Dim I1 As Int16 = i Mod 26
        Dim I2 As Int16 = (i - I1) / 26
        Dim C As String = Chr(Asc("A") + I1)
        If I2 > 0 Then
            C = Chr(Asc("A") + (I2 - 1)) & C
        End If

        Return C
    End Function

    Private Sub grdTATCONVX_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdTATCONVX.InitializeLayout

    End Sub
End Class