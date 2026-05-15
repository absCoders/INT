Public Class ASFCOMP1
    Dim SOLUTION_1 As String
    Dim SOLUTION_1_S As Integer
    Dim SOLUTIONS As New Dictionary(Of String, Integer)
    Dim FOLDERS As New Dictionary(Of String, String)
    Dim PROJECTS As New List(Of String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Form_Load_Custom()
        Catch ex As Exception
            error_has_occured = ex
        End Try
    End Sub

    Sub Form_Load_Custom()
        With dst
            With .Tables.Add("ASTCOMPS")
                .Columns.Add("SOLUTION")
                .Columns.Add("FOLDER")
                .Columns.Add("PROJECTS")
                .PrimaryKey = New DataColumn() {.Columns("SOLUTION")}
            End With

            With .Tables.Add("ASTCOMPP")
                .Columns.Add("PROJECT")
                .Columns.Add("SELECTED")
                .Columns("SELECTED").DefaultValue = "0"
                .PrimaryKey = New DataColumn() {.Columns("PROJECT")}
            End With

            With .Tables.Add("ASTCOMP1")
                .Columns.Add("PROJECT")
                .Columns.Add("OBJECT_TYPE")
                .Columns.Add("OBJECT_NAME")
                .Columns.Add("DATETIME", GetType(System.DateTime))
                .Columns.Add("FILESIZE", GetType(System.Int64))
                .Columns.Add("SOLUTION1")
                .Columns.Add("SOLUTION2")
                .Columns.Add("SOLUTION3")
                .Columns.Add("SOLUTION4")
                .Columns.Add("SOLUTION5")
                .PrimaryKey = New DataColumn() {.Columns("PROJECT"), .Columns("OBJECT_TYPE"), .Columns("OBJECT_NAME")}
            End With

            With .Tables.Add("ASTCOMP2")
                .Columns.Add("FILENAME")
                .Columns.Add("SOLUTION")
                .Columns.Add("PROJECT")
                .Columns.Add("OBJECT_TYPE")
                .Columns.Add("OBJECT_NAME")
                .Columns.Add("DATETIME", GetType(System.DateTime))
                .Columns.Add("FILESIZE", GetType(System.Int64))
                .PrimaryKey = New DataColumn() {.Columns("FILENAME")}
            End With

            .Relations.Add(ASCDATA1.GetRelation(dst, "ASTCOMP1", "ASTCOMP2", "PROJECT,OBJECT_TYPE,OBJECT_NAME"))


            With .Tables.Add("ASTDBAS1")
                .Columns.Add("SCHEMA", GetType(System.String))
                .Columns.Add("TABLE_NAME", GetType(System.String))
                .Columns.Add("DDL", GetType(System.String))
                .PrimaryKey = New DataColumn() {.Columns("SCHEMA"), .Columns("TABLE_NAME")}
            End With

            With .Tables.Add("ASTDBAS2")
                .Columns.Add("SCHEMA", GetType(System.String))
                .Columns.Add("TABLE_NAME", GetType(System.String))
                .Columns.Add("COLUMN_ID", GetType(System.Int16))
                .Columns.Add("COLUMN_NAME", GetType(System.String))
                .Columns.Add("KEY", GetType(System.String))
                .Columns.Add("DATA_SPEC", GetType(System.String))
                .PrimaryKey = New DataColumn() { _
                .Columns("SCHEMA"), _
                .Columns("TABLE_NAME"), _
                .Columns("COLUMN_ID")}
            End With
            Create_Relation("ASTDBAS1", "ASTDBAS2", "SCHEMA,TABLE_NAME")

            With .Tables.Add("ASTDBAS3")
                .Columns.Add("SCHEMA", GetType(System.String))
                .Columns.Add("TABLE_NAME", GetType(System.String))
                .Columns.Add("INDEX_NAME", GetType(System.String))
                .Columns.Add("COLUMN_NAMES", GetType(System.String))
                .Columns.Add("PK", GetType(System.String))
                .PrimaryKey = New DataColumn() { _
                .Columns("SCHEMA"), _
                .Columns("TABLE_NAME"), _
                .Columns("INDEX_NAME")}
            End With
            Create_Relation("ASTDBAS1", "ASTDBAS3", "SCHEMA,TABLE_NAME")
        End With

        'ASCMAIN1.sql = "sELECT * FROM NOTABLE"
        'Dim DT As DataTable = ASCDATA1.GetDataTable


        grdASTDBAS1.DataSource = dst.Tables("ASTDBAS1")

        grdASTCOMP1.DataSource = dst.Tables("ASTCOMP1")
        grdASTCOMP2.DataSource = dst.Tables("ASTCOMP2")
        grdASTCOMPS.DataSource = dst.Tables("ASTCOMPS")
        grdASTCOMPP.DataSource = dst.Tables("ASTCOMPP")

        Create_Summary(grdASTCOMPS, "SOLUTION", "Count")
        Create_Summary(grdASTCOMPP, "PROJECT", "Count")
        Create_Summary(grdASTCOMPP, "SELECTED")

        spl.Panel1Collapsed = True

        grdASTCOMP1.DisplayLayout.Bands(0).Columns("PROJECT").HiddenWhenGroupBy = DefaultableBoolean.True
        grdASTCOMP1.DisplayLayout.Bands(0).Columns("OBJECT_TYPE").HiddenWhenGroupBy = DefaultableBoolean.True

    End Sub
    Private Sub Form_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                'Try
                '    ASCMAIN1.sql = "sELECT * FROM NOTABLE"
                '    Dim DT As DataTable = ASCDATA1.GetDataTable
                'Catch ex As Exception
                '    If TypeOf (ASCMAIN1.ActiveForm) Is ASFBASE1 Then
                '        Dim ABSF As ASFBASE1 = DirectCast(ASCMAIN1.ActiveForm, ASFBASE1)
                '        ASCMAIN1.Disable_Form(ABSF)
                '    End If
                '    Exit Sub
                'End Try

                If UltraTabControl1.SelectedTab.Key = "Database Definitions" Then
                    If txtDB1.Text = "" Or txtDB2.Text = "" Then
                        EMsg &= vbCr & "You Must Specify 2 Schema's to Compare"
                    End If
                Else
                    If dst.Tables("ASTCOMPS").Rows.Count = 0 Then
                        EMsg &= vbCr & "No Solutions Specified"
                    Else
                        If dst.Tables("ASTCOMPS").Rows.Count = 1 Then
                            EMsg &= vbCr & "At least 2 Projects are Required"
                        ElseIf dst.Tables("ASTCOMPS").Rows.Count = 1 Then
                            EMsg &= vbCr & "No More than 5 Projects are Allowed"
                        End If

                        If dst.Tables("ASTCOMPP").Select("SELECTED = '1'").Length = 0 Then
                            EMsg &= vbCr & "No Projects Selected"
                        End If
                    End If
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

            Case "Refresh"
                Load_File_Attributes()
            Case "Done"
                Call Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
            If Not ScreenMode Then
                .Groups("Screen Control").Items("Refresh").Settings.Enabled = iScreenMode
                '.Groups("Screen Control").Items("Print").Settings.Enabled = iScreenMode
            Else
                If UltraTabControl1.SelectedTab.Key = "Database Definitions" Then
                    '.Groups("Screen Control").Items("Print").Settings.Enabled = iScreenMode
                Else
                    .Groups("Screen Control").Items("Refresh").Settings.Enabled = iScreenMode
                End If
            End If
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            .Groups("Database Definitions").Visible = Not ScreenMode And UltraTabControl1.SelectedTab.Key = "Database Definitions"
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        splASTCOMPS.Visible = Not ScreenMode
        grdASTDBAS1.Visible = False
        splDB.Visible = ScreenMode
        If ScreenMode Then
            If UltraTabControl1.SelectedTab.Key = "Database Definitions" Then
                UltraTabControl1.Tabs("Source Code").Enabled = False
                grdASTDBAS1.Visible = True
            Else
                UltraTabControl1.Tabs("Database Definitions").Enabled = False
                SplitContainer1.Visible = ScreenMode
            End If

        Else
            Clear_Record()
            SplitContainer1.Visible = ScreenMode

            grdASTCOMPS.Text = "Solutions"

            UltraTabControl1.Tabs("Source Code").Enabled = True
            UltraTabControl1.Tabs("Database Definitions").Enabled = True

        End If

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Files")
        Me.Cursor = Cursors.WaitCursor

        If UltraTabControl1.SelectedTab.Key = "Database Definitions" Then
            Call Load_Record_Database()
        Else
            Call Load_Record_Source()
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        dst.Tables("ASTCOMPS").Rows.Clear()
        dst.Tables("ASTCOMPP").Rows.Clear()
        dst.Tables("ASTCOMP1").Rows.Clear()
        dst.Tables("ASTCOMP2").Rows.Clear()

        dst.Tables("ASTDBAS1").Rows.Clear()
        dst.Tables("ASTDBAS2").Rows.Clear()
        dst.Tables("ASTDBAS3").Rows.Clear()

        EnforceConstraints(True)
        SOLUTION_1 = ""

        For S As Integer = 1 To 5
            With grdASTCOMP1.DisplayLayout.Bands(0).Columns("SOLUTION" & CStr(S))
                .Hidden = True
            End With
        Next

    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As System.Windows.Forms.Control)
        'If COLUMN_NAME = "USER_ID" Then
        '    If ctl.Text <> "" Then
        '        'Call Click_Command("Load Reports")
        '    End If
        'End If
    End Sub

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        'If Absx1.GetABSColumnName(sender) = "USER_ID" Then
        '    If e.KeyCode = Windows.Forms.Keys.Enter Then
        '        Call Click_Command("Load Reports", e)
        '    End If
        'End If
    End Sub
#End Region

#Region "grdASTSPRF1"

#End Region

    Private Sub grdASTCOMPS_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTCOMPS.AfterRowActivate

    End Sub

    Private Sub grdASTCOMPS_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTCOMPS.AfterRowsDeleted
        List_Projects()
    End Sub

    Private Sub grdASTCOMPS_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdASTCOMPS.AfterRowUpdate
        List_Projects()
    End Sub

    Private Sub grdASTCOMPS_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdASTCOMPS.BeforeCellUpdate
    End Sub

    Private Sub grdASTCOMPS_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdASTCOMPS.BeforeExitEditMode
        If grdASTCOMPS.ActiveCell.Column.Key = "SOLUTION" Then
            If grdASTCOMPS.ActiveCell.Text <> "" Then
                grdASTCOMPS.ActiveCell.Value = grdASTCOMPS.ActiveCell.Text.ToString.ToUpper
            End If
        End If
    End Sub

    Private Sub grdASTCOMPS_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdASTCOMPS.BeforeRowUpdate
        If e.Row.IsAddRow Then
            If e.Row.Cells("SOLUTION").Text <> "" Then
                Dim FOLDER As String = "C:\VS\" & e.Row.Cells("SOLUTION").Text & "\"
                If Not My.Computer.FileSystem.DirectoryExists(FOLDER) Then
                    MsgBox("Solution Folder Not Found for " & e.Row.Cells("SOLUTION").Text, MsgBoxStyle.OkOnly, "Bad Solution")
                    e.Cancel = True
                    Exit Sub
                End If

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Listing all Projects")
                e.Row.Cells("FOLDER").Value = FOLDER
                Dim PROJECTS As String = ""

                For Each FILENAME As String In My.Computer.FileSystem.GetFiles(FOLDER, FileIO.SearchOption.SearchAllSubDirectories, "*.VBPROJ")
                    Dim FILEINFO As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
                    Dim PROJECT As String = Split(FILEINFO.Name, ".")(0)
                    If FILEINFO.DirectoryName = FOLDER & PROJECT Then
                        PROJECTS &= "," & PROJECT
                    End If
                Next
                e.Row.Cells("PROJECTS").Value = Mid(PROJECTS, 2)
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
            Else
                If e.Row.DataChanged Then
                    e.Cancel = True
                End If
            End If

        End If
    End Sub

    Sub List_Projects()
        Dim PROJECTS_in_SOLUTIONS As New List(Of String)
        For Each row As DataRow In dst.Tables("ASTCOMPS").Select
            Dim PROJECTS As String = row.Item("PROJECTS") & ""

            For Each PROJECT As String In Split(PROJECTS, ",")
                If Not PROJECTS_in_SOLUTIONS.Contains(PROJECT) Then
                    PROJECTS_in_SOLUTIONS.Add(PROJECT)
                End If
            Next
        Next

        For Each row As DataRow In dst.Tables("ASTCOMPP").Select
            Dim PROJECT As String = row.Item("PROJECT")
            If Not PROJECTS_in_SOLUTIONS.Contains(PROJECT) Then
                row.Item("SELECTED") = "2"
            End If
        Next

        ASCDATA1.DeleteRows("ASTCOMPP", "SELECTED = '2'")

        For Each PROJECT As String In PROJECTS_in_SOLUTIONS
            Dim row As DataRow = dst.Tables("ASTCOMPP").Rows.Find(PROJECT)
            If row Is Nothing Then
                row = dst.Tables("ASTCOMPP").NewRow
                row.Item("PROJECT") = PROJECT
                dst.Tables("ASTCOMPP").Rows.Add(row)
            End If
        Next

        If grdASTCOMPS.Rows.Count = 1 Then
            SOLUTION_1 = grdASTCOMPS.Rows(0).Cells("SOLUTION").Text
        ElseIf grdASTCOMPS.Rows.Count = 0 Then
            SOLUTION_1 = ""
        End If

        If SOLUTION_1 = "" Then
            grdASTCOMPS.Text = "Solutions"
            grdASTCOMP1.Text = "Comparison Chart"
        Else
            grdASTCOMPS.Text = "Solutions (Primary Solution = " & SOLUTION_1 & ")"
            grdASTCOMP1.Text = "Comparison Chart (Primary Solution = " & SOLUTION_1 & ")"
        End If

        Sort_grdColumns(grdASTCOMPP, "PROJECT")
    End Sub

    Private Sub grdASTCOMP1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTCOMP1.AfterRowActivate

        If grdASTCOMP1.ActiveRow.IsDataRow Then
            grdASTCOMP2.Visible = True
            Dim PROJECT As String = grdASTCOMP1.ActiveRow.Cells("PROJECT").Value & ""
            Dim OBJECT_TYPE As String = grdASTCOMP1.ActiveRow.Cells("OBJECT_TYPE").Value & ""
            Dim OBJECT_NAME As String = grdASTCOMP1.ActiveRow.Cells("OBJECT_NAME").Value & ""

            Dim dvw As DataView = DirectCast(grdASTCOMP2.DataSource, DataTable).DefaultView
            dvw.RowFilter = "PROJECT = '" & PROJECT & "'" _
            & " and OBJECT_TYPE = '" & OBJECT_TYPE & "'" _
            & " and OBJECT_NAME = '" & OBJECT_NAME & "'"
            grdASTCOMP2.Text = "File Details (" & PROJECT & ":" & OBJECT_TYPE & ":" & OBJECT_NAME & ")"
        Else
            grdASTCOMP2.Visible = False
        End If

    End Sub

    Private Sub grdASTCOMP1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdASTCOMP1.InitializeRow

        If grdASTCOMP1.Tag <> "" Then
            Exit Sub
        End If

        If e.Row.Band.Key <> "ASTCOMP1" Then
            Exit Sub
        End If

        For Each SOLUTION As String In SOLUTIONS.Keys
            If SOLUTION <> SOLUTION_1 Then
                Dim FOLDER As String = FOLDERS(SOLUTION)
                Dim PROJECT As String = e.Row.Cells("PROJECT").Value
                Dim OBJECT_TYPE As String = e.Row.Cells("OBJECT_TYPE").Value
                Dim OBJECT_NAME As String = e.Row.Cells("OBJECT_NAME").Value
                'If OBJECT_NAME Like "ASCPRINT.*" Then Stop
                Dim FILENAME As String = FOLDER & PROJECT & "\" & OBJECT_TYPE & "\" & OBJECT_NAME
                Dim rowASTCOMP2 As DataRow = dst.Tables("ASTCOMP2").Rows.Find(FILENAME)
                Dim S As Integer = SOLUTIONS(SOLUTION)
                e.Row.Cells("SOLUTION" & CStr(S)).Appearance.BackColor = Drawing.Color.Empty
                If rowASTCOMP2 Is Nothing Then
                    e.Row.Cells("SOLUTION" & CStr(S)).Appearance.BackColor = Drawing.Color.Yellow
                Else
                    If Val(e.Row.Cells("FILESIZE").Value & "") <> Val(rowASTCOMP2.Item("FILESIZE") & "") Then
                        e.Row.Cells("SOLUTION" & CStr(S)).Appearance.BackColor = Drawing.Color.Red
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub grdASTCOMP1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdASTCOMP1.InitializeLayout

    End Sub

    Private Sub grdASTCOMP1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdASTCOMP1.DoubleClickRow

    End Sub

    Private Sub grdASTCOMP1_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdASTCOMP1.DoubleClickCell

        If e.Cell.Column.Key Like "SOLUTION*" Then
            If e.Cell.Value & "" <> "1" Or e.Cell.Row.Cells("SOLUTION" & CStr(SOLUTION_1_S)).Value & "" <> "1" Then
                Exit Sub
            End If
            Dim SOLUTION As String = e.Cell.Column.Header.Caption
            If SOLUTION <> SOLUTION_1 Then
                Dim PROJECT As String = e.Cell.Row.Cells("PROJECT").Value
                Dim OBJECT_TYPE As String = e.Cell.Row.Cells("OBJECT_TYPE").Value
                Dim OBJECT_NAME As String = e.Cell.Row.Cells("OBJECT_NAME").Value
                Dim CL As String = "C:\Program Files (x86)\Compare It!\wincmp3.exe " _
                & "C:\VS\" & SOLUTION_1 & "\" & PROJECT & "\" & OBJECT_TYPE & "\" & OBJECT_NAME & " " _
                & "C:\VS\" & SOLUTION & "\" & PROJECT & "\" & OBJECT_TYPE & "\" & OBJECT_NAME
                Shell(CL)
            End If
        End If

    End Sub

    'Private Sub FileSystemWatcher1_Changed(ByVal sender As System.Object, ByVal e As System.IO.FileSystemEventArgs)
    '    Stop

    '    If ScreenMode Then
    '        Dim rowASTCOMP2 As DataRow = dst.Tables("ASTCOMP2").Rows.Find(e.FullPath)
    '        Dim FILEINFO As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(e.FullPath)
    '        rowASTCOMP2.Item("DATETIME") = FILEINFO.LastWriteTime
    '        rowASTCOMP2.Item("FILESIZE") = FILEINFO.Length
    '        If rowASTCOMP2.Item("SOLUTION") = SOLUTION_1 Then
    '            Dim rowASTCOMP1 As DataRow = rowASTCOMP2.GetParentRow("ASTCOMP1_ASTCOMP2")
    '            rowASTCOMP1.Item("DATETIME") = FILEINFO.LastWriteTime
    '            rowASTCOMP1.Item("FILESIZE") = FILEINFO.Length
    '        End If
    '    End If
    'End Sub

    Sub Load_File_Attributes()

        ASCMAIN1.Progress("Now Loading Files")

        Dim S As Integer = 0

        dst.Tables("ASTCOMP2").Rows.Clear()
        dst.Tables("ASTCOMP1").Rows.Clear()

        'grdASTCOMP1.DataSource = Nothing

        'grdASTCOMP1.Tag = "x"

        For Each rowASTCOMPS As DataRow In dst.Tables("ASTCOMPS").Select("", "SOLUTION")
            Dim SOLUTION As String = rowASTCOMPS.Item("SOLUTION")
            Dim FOLDER As String = rowASTCOMPS.Item("FOLDER")
            For Each PROJECT As String In PROJECTS
                If My.Computer.FileSystem.DirectoryExists(FOLDER & PROJECT) Then
                    For Each FILENAME As String In My.Computer.FileSystem.GetFiles(FOLDER & PROJECT, FileIO.SearchOption.SearchAllSubDirectories, "*.VB")
                        Dim FILEINFO As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
                        Dim OBJECT_TYPE As String = ""
                        If FILEINFO.DirectoryName = FOLDER & PROJECT & "\Classes" Then
                            OBJECT_TYPE = "Classes"
                        ElseIf FILEINFO.DirectoryName = FOLDER & PROJECT & "\Forms" Then
                            OBJECT_TYPE = "Forms"
                        ElseIf FILEINFO.DirectoryName = FOLDER & PROJECT & "\Tables" Then
                            OBJECT_TYPE = "Tables"
                        ElseIf FILEINFO.DirectoryName = FOLDER & PROJECT & "\Reports" Then
                            OBJECT_TYPE = "Reports"
                        End If
                        If OBJECT_TYPE <> "" Then
                            Dim OBJECT_NAME As String = FILEINFO.Name

                            Dim rowASTCOMP1 As DataRow = dst.Tables("ASTCOMP1").Rows.Find _
                            (New String() {PROJECT, OBJECT_TYPE, OBJECT_NAME})
                            If rowASTCOMP1 Is Nothing Then
                                rowASTCOMP1 = dst.Tables("ASTCOMP1").NewRow
                                rowASTCOMP1.Item("PROJECT") = PROJECT
                                rowASTCOMP1.Item("OBJECT_TYPE") = OBJECT_TYPE
                                rowASTCOMP1.Item("OBJECT_NAME") = OBJECT_NAME
                                For S = 1 To 5
                                    rowASTCOMP1.Item("SOLUTION" & CStr(S)) = "0"
                                Next
                                dst.Tables("ASTCOMP1").Rows.Add(rowASTCOMP1)
                            End If
                            S = SOLUTIONS(SOLUTION)
                            rowASTCOMP1.Item("SOLUTION" & CStr(S)) = "1"

                            If SOLUTION = SOLUTION_1 Then
                                rowASTCOMP1.Item("DATETIME") = FILEINFO.LastWriteTime
                                rowASTCOMP1.Item("FILESIZE") = FILEINFO.Length
                            Else

                            End If


                            Dim rowASTCOMP2 As DataRow = dst.Tables("ASTCOMP2").NewRow
                            rowASTCOMP2.Item("FILENAME") = FILENAME
                            rowASTCOMP2.Item("SOLUTION") = SOLUTION
                            rowASTCOMP2.Item("PROJECT") = PROJECT
                            rowASTCOMP2.Item("OBJECT_TYPE") = OBJECT_TYPE
                            rowASTCOMP2.Item("OBJECT_NAME") = OBJECT_NAME
                            rowASTCOMP2.Item("DATETIME") = FILEINFO.LastWriteTime
                            rowASTCOMP2.Item("FILESIZE") = FILEINFO.Length
                            dst.Tables("ASTCOMP2").Rows.Add(rowASTCOMP2)
                        End If
                    Next
                End If
            Next
        Next

        'grdASTCOMP1.DataSource = dst.Tables("ASTCOMP1")

        With grdASTCOMP1.DisplayLayout.Bands(0).SortedColumns
            .Clear()
            .Add("PROJECT", False, True)

            .Add("OBJECT_TYPE", False, True)
        End With
        grdASTCOMP1.Rows.ExpandAll(True)

        'grdASTCOMP1.Tag = ""

        grdASTCOMP1.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)


        ASCMAIN1.Progress("")
    End Sub

    Private Sub UltraTabControl1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        UltraExplorerBar1.Groups("Database Definitions").Visible = (UltraTabControl1.SelectedTab.Key = "Database Definitions")
    End Sub

    Sub Load_Record_Source()

        SOLUTION_1_S = 0

        Dim S As Integer = 0
        FOLDERS.Clear()
        SOLUTIONS.Clear()
        For Each rowASTCOMPS As DataRow In dst.Tables("ASTCOMPS").Select()
            Dim SOLUTION As String = rowASTCOMPS.Item("SOLUTION")
            S += 1
            SOLUTIONS.Add(SOLUTION, S)
            With grdASTCOMP1.DisplayLayout.Bands(0).Columns("SOLUTION" & CStr(S))
                .Hidden = False
                .Header.Caption = SOLUTION
                .Width = 60
                .Header.Appearance.BackColor = Drawing.Color.Empty
                .Header.Appearance.ForeColor = Drawing.Color.Empty
                .Header.Appearance.BackGradientStyle = GradientStyle.None
                If SOLUTION = SOLUTION_1 Then
                    .Header.Appearance.BackColor = Drawing.Color.Green
                    .Header.Appearance.ForeColor = Drawing.Color.White
                    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                End If
            End With
            If SOLUTION = SOLUTION_1 Then SOLUTION_1_S = S
            FOLDERS.Add(SOLUTION, rowASTCOMPS.Item("FOLDER"))
        Next

        PROJECTS.Clear()
        For Each rowASTCOMPP As DataRow In dst.Tables("ASTCOMPP").Select("SELECTED = '1'")
            Dim PROJECT As String = rowASTCOMPP.Item("PROJECT")
            PROJECTS.Add(PROJECT)
        Next

        Load_File_Attributes()
    End Sub

    Sub Load_Record_Database()

        For Each SCHEMA As String In New String() {txtDB1.Text, txtDB2.Text}
            Dim TNS As String
            If SCHEMA = txtDB1.Text Then
                TNS = txtTNS1.Text
            Else
                TNS = txtTNS2.Text
            End If
            If TNS <> "" Then
                TNS = "@" & TNS
            End If
            ASCMAIN1.sql = "Select * from ALL_TAB_COLUMNS" & TNS _
            & " where OWNER = '" & SCHEMA & "'" _
            & "   and TABLE_NAME like '" & txtTableNamePFX.Text & "%'"
            Dim TABLE_NAME As String = ""
            Dim KEY_COLUMNS As New List(Of String)
            Dim COLUMN_NAME As String
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "TABLE_NAME,COLUMN_ID")
                If row.Item("TABLE_NAME") & "" <> TABLE_NAME Then
                    TABLE_NAME = row.Item("TABLE_NAME")
                    ASCMAIN1.Progress(SCHEMA & ":" & TABLE_NAME)
                    Application.DoEvents()
                    Dim rowASTDBAS1 As DataRow = dst.Tables("ASTDBAS1").NewRow
                    rowASTDBAS1.Item("SCHEMA") = SCHEMA
                    rowASTDBAS1.Item("TABLE_NAME") = TABLE_NAME
                    dst.Tables("ASTDBAS1").Rows.Add(rowASTDBAS1)

                    ASCMAIN1.sql = "Select ALL_INDEXES.*, ALL_CONSTRAINTS.CONSTRAINT_TYPE " _
                    & " from ALL_INDEXES" & TNS & ", ALL_CONSTRAINTS" & TNS _
                    & " where ALL_INDEXES.OWNER = '" & SCHEMA & "'" _
                    & "   and ALL_INDEXES.TABLE_NAME = '" & TABLE_NAME & "'" _
                    & "   and ALL_CONSTRAINTS.OWNER (+) = ALL_INDEXES.OWNER" _
                    & "   and ALL_CONSTRAINTS.CONSTRAINT_NAME (+) = ALL_INDEXES.INDEX_NAME"
                    For Each row2 As DataRow In ASCDATA1.GetDataTable.Select("", "INDEX_NAME")
                        Dim INDEX_NAME As String = row2.Item("INDEX_NAME")
                        Dim PK As String = IIf(row2.Item("CONSTRAINT_TYPE") & "" = "P", "1", "0")

                        Dim COLUMN_NAMES As String = ""
                        ASCMAIN1.sql = "Select * from ALL_IND_COLUMNS" & TNS _
                        & " where ALL_IND_COLUMNS.INDEX_OWNER = '" & SCHEMA & "'" _
                        & "   and ALL_IND_COLUMNS.INDEX_NAME = '" & INDEX_NAME & "'"
                        For Each row3 As DataRow In ASCDATA1.GetDataTable.Select("", "COLUMN_POSITION")
                            COLUMN_NAME = row3.Item("COLUMN_NAME")
                            COLUMN_NAMES &= "," & COLUMN_NAME
                            If row3.Item("DESCEND") & "" = "DESC" Then
                                COLUMN_NAMES &= " DESC"
                            End If
                            If PK = "1" Then
                                KEY_COLUMNS.Add(COLUMN_NAME)
                            End If
                        Next

                        Dim rowASTDBAS3 As DataRow = dst.Tables("ASTDBAS3").NewRow
                        rowASTDBAS3.Item("SCHEMA") = SCHEMA
                        rowASTDBAS3.Item("TABLE_NAME") = TABLE_NAME
                        rowASTDBAS3.Item("INDEX_NAME") = INDEX_NAME
                        rowASTDBAS3.Item("COLUMN_NAMES") = Mid(COLUMN_NAMES, 2)
                        rowASTDBAS3.Item("PK") = PK
                        dst.Tables("ASTDBAS3").Rows.Add(rowASTDBAS3)
                    Next
                End If

                Dim rowASTDBAS2 As DataRow = dst.Tables("ASTDBAS2").NewRow
                rowASTDBAS2.Item("SCHEMA") = SCHEMA
                rowASTDBAS2.Item("TABLE_NAME") = TABLE_NAME
                rowASTDBAS2.Item("COLUMN_ID") = row.Item("COLUMN_ID")
                COLUMN_NAME = row.Item("COLUMN_NAME")
                rowASTDBAS2.Item("COLUMN_NAME") = COLUMN_NAME
                rowASTDBAS2.Item("KEY") = IIf(KEY_COLUMNS.Contains(COLUMN_NAME), "1", "0")

                Dim DATA_SPEC As String = ""
                Select Case row.Item("DATA_TYPE").ToString
                    Case "VARCHAR2"
                        DATA_SPEC = "VARCHAR2(" & row.Item("DATA_LENGTH") & ")"
                    Case "NUMBER"
                        If row.Item("DATA_PRECISION").Equals(System.DBNull.Value) Or row.Item("DATA_SCALE").Equals(System.DBNull.Value) Then
                            DATA_SPEC = "NUMBER"
                        Else
                            DATA_SPEC = "NUMBER(" & row.Item("DATA_PRECISION") & IIf(row.Item("DATA_SCALE") = 0, "", "," & row.Item("DATA_SCALE")) & ")"
                        End If
                    Case "DATE"
                        DATA_SPEC = "DATE"
                    Case Else
                        DATA_SPEC = row.Item("DATA_TYPE").ToString
                End Select

                rowASTDBAS2.Item("DATA_SPEC") = DATA_SPEC
                dst.Tables("ASTDBAS2").Rows.Add(rowASTDBAS2)
            Next

            Dim FILENAME As String = ASCMAIN1.Folders("Temp") & SCHEMA & ".txt"
            If My.Computer.FileSystem.FileExists(FILENAME) Then
                My.Computer.FileSystem.DeleteFile(FILENAME)
            End If

            Using sw As New System.IO.StreamWriter(FILENAME)
                For Each rowASTDBAS1 As DataRow In dst.Tables("ASTDBAS1") _
                    .Select("SCHEMA = '" & SCHEMA & "'", "TABLE_NAME")
                    Dim DDL As String = ""
                    Dim KEY As String = ""
                    For Each rowASTDBAS2 As DataRow In rowASTDBAS1.GetChildRows("ASTDBAS1_ASTDBAS2")
                        DDL = DDL & "," & vbCrLf & rowASTDBAS2.Item("COLUMN_NAME") & " " & rowASTDBAS2.Item("DATA_SPEC")
                        If rowASTDBAS2.Item("KEY") & "" = "1" Then
                            KEY = KEY & ", " & rowASTDBAS2.Item("COLUMN_NAME")
                        End If
                    Next
                    DDL = "Create Table " & rowASTDBAS1.Item("TABLE_NAME") & " (" & Mid$(DDL, 2)
                    If KEY <> "" Then
                        DDL = DDL & "," & vbCrLf & "Primary Key (" & Mid$(KEY, 3) & ")"
                    End If
                    DDL = DDL & ");"

                    For Each rowASTDBAS3 As DataRow In rowASTDBAS1.GetChildRows("ASTDBAS1_ASTDBAS3")
                        If rowASTDBAS3.Item("PK") & "" <> "1" Then
                            DDL = DDL & vbCrLf & "Create Index " & rowASTDBAS3.Item("INDEX_NAME") & " on " & rowASTDBAS1.Item("TABLE_NAME") & " (" & rowASTDBAS3.Item("COLUMN_NAMES") & ");"
                        End If
                    Next

                    rowASTDBAS1.Item("DDL") = DDL
                    sw.Write(vbCrLf & DDL & vbCrLf)
                Next

            End Using
        Next

        'With grdASTDBAS1.DisplayLayout.Bands(0)
        '    .ColumnFilters.ClearAllFilters()
        '    .ColumnFilters("SCHEMA").FilterConditions.Add _
        '    (UltraWinGrid.FilterComparisionOperator.Equals, txtDB1.Text)
        'End With


        Dim CL As String = "C:\Program Files (x86)\Compare It!\wincmp3.exe " _
        & ASCMAIN1.Folders("Temp") & txtDB1.Text & ".txt" & " " _
        & ASCMAIN1.Folders("Temp") & txtDB2.Text & ".txt"
        Shell(CL)

        grpDDL1.Text = txtDB1.Text
        grpDDL2.Text = txtDB2.Text
    End Sub

    Private Sub grdASTDBAS1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTDBAS1.ClickCellButton
        Dim TABLE_NAME As String = e.Cell.Row.Cells("TABLE_NAME").Value
        Dim rowASTDBAS1 As DataRow

        rowASTDBAS1 = dst.Tables("ASTDBAS1").Rows.Find _
            (New String() {txtDB1.Text, TABLE_NAME})
        If rowASTDBAS1 Is Nothing Then
            txtDDL1.Text = ""
            grpDDL1.Visible = False
        Else
            txtDDL1.Text = rowASTDBAS1.Item("DDL")
            grpDDL1.Visible = True
        End If

        rowASTDBAS1 = dst.Tables("ASTDBAS1").Rows.Find _
            (New String() {txtDB2.Text, TABLE_NAME})
        If rowASTDBAS1 Is Nothing Then
            txtDDL2.Text = ""
            grpDDL2.Visible = False
        Else
            txtDDL2.Text = rowASTDBAS1.Item("DDL")
            grpDDL2.Visible = True
        End If

    End Sub


    Private Sub grdASTDBAS1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdASTDBAS1.InitializeLayout

    End Sub

    Private Sub grdASTDBAS1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdASTDBAS1.InitializeRow

        If e.Row.Band.Key = "ASTDBAS1" Then
            Dim SCHEMA As String = e.Row.Cells("SCHEMA").Value
            Dim TABLE_NAME As String = e.Row.Cells("TABLE_NAME").Value
            Dim rowASTDBAS1 As DataRow

            Dim SCHEMA_OTHER As String
            If SCHEMA = txtDB1.Text Then
                SCHEMA_OTHER = txtDB2.Text
            Else
                SCHEMA_OTHER = txtDB1.Text
            End If
            rowASTDBAS1 = dst.Tables("ASTDBAS1").Rows.Find _
                (New String() {SCHEMA_OTHER, TABLE_NAME})
            If rowASTDBAS1 Is Nothing Then
                e.Row.Cells("TABLE_NAME").Appearance.ForeColor = Drawing.Color.Red
            Else
                If rowASTDBAS1.Item("DDL") & "" <> e.Row.Cells("DDL").Value & "" Then
                    e.Row.Cells("TABLE_NAME").Appearance.BackColor = Drawing.Color.Yellow
                End If
            End If
        End If
      
    End Sub
End Class