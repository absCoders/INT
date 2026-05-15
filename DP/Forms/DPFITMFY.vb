Public Class DPFITMFY

    Dim SPTCWXBD As String
    Dim SPTCWXBI As String
    Dim rowSPTCWRXX As DataRow
    Dim CWRX_NO As String
    Dim FILENAME As String
    Dim Budget_Seasons As String
    Dim SEASON_YEAR As String
    Dim SEASON_TYPE As String

    Dim workbook As SpreadsheetGear.IWorkbook = Nothing
    Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
    Dim range As SpreadsheetGear.IRange = Nothing
    Dim WithEvents ws As SpreadsheetGear.IWorksheet

    Dim rangeCopyFrom As SpreadsheetGear.IRange
    Dim rangePaste_To As SpreadsheetGear.IRange

    Dim XLS_NO As String
    Dim XLS_PWD As String = "ABS"
    Dim c_FC_Start = -1
    Dim c_FC_End = -1
    Dim r_Start As Integer = 6
    Dim c_UPC As Integer = 2
    Dim c_Type As Integer = 0


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Create_Work_Tables(True)

        With dst
            Create_TDA(.Tables.Add, "GLTPARM3", "*", 0)

            With .Tables.Add("DPTSLPFI").Columns
                .Add("ITEM_CODE")
                .Add("ITEM_DESC")

                For FC As Integer = 0 To 12
                    Dim FCX As String = "FC" & Format(FC, "00")
                    .Add(FCX, GetType(System.Int32))
                Next

            End With
            .Tables("DPTSLPFI").PrimaryKey = New DataColumn() { .Tables("DPTSLPFI").Columns("ITEM_CODE")}



            'ASCMAIN1.sql = "Select * from SPTCWRXX where CWRX_EXPORT_TYPE = 'B'"
            'Create_TDA(.Tables.Add, "SPTCWRXX", "**", 0)

            'ASCMAIN1.sql = "Select * from " & SPTCWXBD
            'Create_TDA(.Tables.Add, "SPTCWXBD", "**", 0, False, "", 0)

            'ASCMAIN1.sql = "Select * from " & SPTCWXBI
            'Create_TDA(.Tables.Add, "SPTCWXBI", "**", 0, False, "", 0)
        End With

        grdDPTSLPFI.DataSource = dst.Tables("DPTSLPFI")

        Create_Summary(grdDPTSLPFI, "ITEM_CODE", "Count")
        With grdDPTSLPFI.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = System.Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If gcol.Key.StartsWith("FC") Then
                    gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                    Create_Summary(grdDPTSLPFI, gcol.Key)
                    Dim MM As Integer = Val(Mid(gcol.Key, 3))
                    If MM = 0 Then
                        gcol.Header.Caption = "PD FC"
                    Else
                        Dim DT As String = $"{Format(MM, "00")}/01/2024"
                        gcol.Header.Caption = Replace(Format(DateValue(DT), "MMM/yy"), "/", "'")
                    End If

                Else
                    If gcol.Key = "ITEM_CODE" Then
                        gcol.Header.Caption = "Item Code"
                    End If
                    If gcol.Key = "ITEM_DESC" Then
                        gcol.Header.Caption = "Item Description"
                    End If
                    gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
        End With

        'Create_Summary(grdSPTCWXBD, "CUST_CODE", "Count")
        'Create_Summary(grdSPTCWXBD, "BUDGET")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Load"
                FILENAME = Import_from_Excel()

                If FILENAME = "" Then
                    EMsg &= vbCr & "You must select a file to Import forecast data"
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

            Case "Load"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

                'If chkInitialize.Checked Then
                '    MsgBox("Please be aware that you have created a Budget file with ALL budgets, not just those that were added/changed since the last export." & vbCrLf & vbCrLf & "Please compare totals to another source.", MsgBoxStyle.OkOnly, "Verification")
                'End If

            Case "Load Sales"
                Load_Sales()

            Case "Refresh FC by Item"
                Refresh_FC_by_Item()

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
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Refresh FC by Item").Visible = ScreenMode
                    .Items("Load Sales").Visible = ScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                End With

                ' .Groups("Options").Enabled = Not tf
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        'Set_Read_Only_for_ctl(chkInitialize, ScreenMode)

        grdDPTSLPFC.Visible = Not ScreenMode
        tabMain.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        'For Each TABLE_NAME As String In New String() {"SPTCWRXX", "SPTCWXBD", "SPTCWXBI"}
        '    dst.Tables(TABLE_NAME).Rows.Clear()
        'Next
        EnforceConstraints(True)

        'Fill_Records("SPTCWRXX")
        'Sort_grdColumns(grdSPTCWRXX, "CWRX_NO".ToLower)
    End Sub

    Sub Load_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Save_Header_Fields(UltraGroupBox1)

        Open_Workbook()



        'CWRX_NO = ASCMAIN1.Next_Control_No("SPTCWRXX.CWRX_NO")
        'dst.Tables("SPTCWRXX").AcceptChanges()
        'rowSPTCWRXX = dst.Tables("SPTCWRXX").NewRow
        'rowSPTCWRXX.Item("CWRX_NO") = CWRX_NO
        'rowSPTCWRXX.Item("CWRX_EXPORT_TYPE") = "B"
        'rowSPTCWRXX.Item("CWRX_DATE") = Now.Date
        'dst.Tables("SPTCWRXX").Rows.Add(rowSPTCWRXX)

        'EnforceConstraints(False)

        'Fill_Records("GLTPARM3")

        '' Create Budget Seasons as Current Season and Next Season
        'Dim YYYYWW As String = ASCMAIN1.Week_Calc(ASCMAIN1.CYW, -4)
        'Dim row As DataRow = dst.Tables("GLTPARM3").Rows.Find(YYYYWW)
        'Dim YYYYMM As String = row.Item("YYYYMM")
        'SEASON_YEAR = Mid(YYYYMM, 1, 4)
        'SEASON_TYPE = IIf(Mid(YYYYMM, 5, 2) > 6, "F", "S")

        'Budget_Seasons = "'" & SEASON_YEAR & SEASON_TYPE & "'"
        'If Mid(Budget_Seasons, 6, 1) = "F" Then
        '    Budget_Seasons &= ",'" & Val(SEASON_YEAR) + 1 & "S'"
        'Else
        '    Budget_Seasons &= ",'" & SEASON_YEAR & "F'"
        'End If

        ''Create_Work_Tables(False)

        'Fill_Records("SPTCWXBD")
        'Sort_grdColumns(grdSPTCWXBD, "CUST_CODE,CUST_STORE_NO")

        '' Create_Budget_File()
        'rowSPTCWRXX.Item("FILENAME") = FILENAME

        'rowSPTCWRXX.Item("INIT_OPER") = ASCMAIN1.USER_ID
        'rowSPTCWRXX.Item("INIT_DATE") = DATETIME_STAMP

        'EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now sftp'ing the file")

        BeginTrans()

        ' Delta - this is what was sent to CoWorx
        ASCMAIN1.sql = "Insert into SPTCWXBD Select * from " & SPTCWXBD
        ASCDATA1.ExecuteSQL()

        ' Image - this is what we will compare to next time
        ASCMAIN1.sql = "Delete from SPTCWXBI"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into SPTCWXBI Select * from " & SPTCWXBI
        ASCDATA1.ExecuteSQL()

        My.Computer.FileSystem.CopyFile(ASCMAIN1.Folders("Temp") & "\" & FILENAME, ASCMAIN1.Folders("Archive") & "\COWORX\" & FILENAME)
        If dst.Tables("SPTCWXBD").Rows.Count > 0 Then
            TAC.TACSCOM1.sftp_put(Me, "COWORX", True, ASCMAIN1.Folders("Temp") & "\" & FILENAME, FILENAME)
        End If

        Update_Record_TDA("SPTCWRXX")

        CommitTrans("Update Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdDPTSLPFC, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdDPTSLPFI, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
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

        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"
    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "BRAND_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    If Absx1.txtFor("OPS_YYYYPP").Text <> "" Then
                        Click_Command("Load", e)
                    End If
                End If
        End Select
    End Sub
#End Region


    Sub Open_Workbook()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Opening Forecast XLS")

        'Dim FILENAME As String = "C:\Users\wjz\Desktop\SLP\Demand Planning\Forecast_2023_December Production_V02.XLSX"
        WorkbookView1.ActiveWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME, System.Globalization.CultureInfo.CurrentCulture)

        Refresh_FC_by_Item()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")


    End Sub 'Main

    Sub Refresh_FC_by_Item()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Refreshing Forecast by Item")

        Dim MMM2MM As New Dictionary(Of String, Integer)
        MMM2MM.Add("Jan", 1)
        MMM2MM.Add("Feb", 2)
        MMM2MM.Add("Mar", 3)
        MMM2MM.Add("Apr", 4)
        MMM2MM.Add("May", 5)
        MMM2MM.Add("Jun", 6)
        MMM2MM.Add("Jul", 7)
        MMM2MM.Add("Aug", 8)
        MMM2MM.Add("Sep", 9)
        MMM2MM.Add("Oct", 10)
        MMM2MM.Add("Nov", 11)
        MMM2MM.Add("Dec", 12)

        dst.Tables("DPTSLPFI").Rows.Clear()


        c_FC_Start = -1
        c_FC_End = -1

        WorkbookView1.GetLock()
        workbook = WorkbookView1.ActiveWorkbook
        ws = workbook.Worksheets("Core Forecast")

        For c As Integer = 0 To 500
            If ws.Cells(0, c).Value & "" = "2024" And ws.Cells(1, c).Value & "" = "Jan" Then

                If c_fc_start = -1 Then c_fc_start = c

                Dim c0 As Integer = c
                Do Until ws.Cells(0, c).Value & "" <> "2024"
                    Dim mmm As String = ws.Cells(1, c).Value & ""
                    Dim mm As Integer = MMM2MM(mmm)

                    Dim r As Integer = r_Start
                    Dim c_UPC As Integer = 2
                    Dim c_Type As Integer = 0
                    Do Until ws.Cells(r, c_UPC).Value & "" = ""
                        If ws.Cells(r, c_Type).Value & "" = "Sales" Then
                            Dim FC As Integer = Val(ws.Cells(r, c).Value & "")
                            If FC <> 0 Then
                                Dim ITEM_CODE As String = ws.Cells(r, c_UPC).Value & ""
                                ' dst.Tables("DPTSLPFI").PrimaryKey = New DataColumn() {dst.Tables("DPTSLPFI").Columns("ITEM_CODE")}
                                Dim row As DataRow = dst.Tables("DPTSLPFI").Rows.Find(ITEM_CODE)
                                If row Is Nothing Then
                                    row = dst.Tables("DPTSLPFI").NewRow
                                    row.Item("ITEM_CODE") = ITEM_CODE
                                    row.Item("ITEM_DESC") = ws.Cells(r, c_UPC + 1).Value & ""
                                    dst.Tables("DPTSLPFI").Rows.Add(row)
                                End If
                                Dim FCX As String = "FC" & Format(mm, "00")

                                row.Item(FCX) = Val(row.Item(FCX) & "") + FC
                            End If
                        End If
                        r += 1

                        If ws.Cells(r, c_UPC).Value & "" = "" Then
                            r += 1
                        End If
                    Loop
                    c += 1
                    c_FC_End = c
                Loop
            End If
        Next

        Sort_grdColumns(grdDPTSLPFI, "ITEM_CODE")

        'ws.Protect(XLS_PWD)
        WorkbookView1.ReleaseLock()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Load_Sales()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Sales")

        WorkbookView1.GetLock()

        workbook = WorkbookView1.ActiveWorkbook
        ws = workbook.Worksheets("Core Forecast")
        ws = workbook.Worksheets("Stock Movement Report")

        Dim c As Integer = c_FC_Start
        Dim r As Integer = 1

        Do Until ws.Cells(0, r).Value & "" = ""

            'If ws.Cells(r, 8).Value & "" = "FY24" And ws.Cells(r, 9).Value & "" = "Jan" Then
            Dim sls As Integer = Val(ws.Cells(r, 6).Value & "")
            ws.Cells(r, 6).Value = sls + 1
            'End If
            r += 1
        Loop

        workbook.WorkbookSet.Calculate() ' = SpreadsheetGear.Calculation.Automatic

        WorkbookView1.ReleaseLock()


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Function Import_from_Excel() As String
        Dim FILENAME As String = ""

        Dim openFileDialog1 As New OpenFileDialog
        openFileDialog1.InitialDirectory = "C:\Users\wjz\Desktop\SLP\Demand Planning\"
        openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
        openFileDialog1.Filter = "xls files (*.xls)|*.xls|xlsX files (*.xlsX)|*.xlsX"
        'openFileDialog1.FilterIndex = 1
        openFileDialog1.RestoreDirectory = True

        If openFileDialog1.ShowDialog() = DialogResult.OK Then

            FILENAME = openFileDialog1.FileName

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End If

        Return FILENAME
    End Function

End Class