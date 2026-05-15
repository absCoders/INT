Public Class ASFSCAN1
    Dim AS_PARM_SCAN_FOLDER As String = ""

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst

            Create_TDA(.Tables.Add, "ASTATTA1", "*")
            Create_TDA(.Tables.Add, "ASTATTA2", "*")

            ASCMAIN1.sql = "SELECT ASTATTA3.* from ASTATTA3 where SCAN_FOLDER is Not Null"
            Create_TDA(.Tables.Add, "ASTATTA3", "**", 0, False, "V", 3)
            '.Tables("ASTATTA3").Columns.Add("FILES", GetType(System.Int32))

            With .Tables.Add("ASTATTAF")
                .Columns.Add("TABLE_NAME", GetType(System.String))
                .Columns.Add("COLUMN_NAME", GetType(System.String))
                .Columns.Add("DOCUMENT_TYPE", GetType(System.String))
                .Columns.Add("FILENAME", GetType(System.String))
                .Columns.Add("FILESIZE", GetType(System.Int32))
                .Columns.Add("FILEDATE", GetType(System.DateTime))
                .Columns.Add("CODE_VALUE", GetType(System.String))
                .Columns("CODE_VALUE").ReadOnly = False
                .Columns.Add("DESC_VALUE", GetType(System.String))
                .Columns("DESC_VALUE").ReadOnly = False
                .Columns.Add("ATTACHMENT_DESC", GetType(System.String))
            End With

            .Relations.Add("ASTATTA3_ASTATTAF" _
            , New DataColumn() { _
            .Tables("ASTATTA3").Columns("TABLE_NAME"), _
            .Tables("ASTATTA3").Columns("COLUMN_NAME"), _
            .Tables("ASTATTA3").Columns("DOCUMENT_TYPE")} _
            , New DataColumn() { _
            .Tables("ASTATTAF").Columns("TABLE_NAME"), _
            .Tables("ASTATTAF").Columns("COLUMN_NAME"), _
            .Tables("ASTATTAF").Columns("DOCUMENT_TYPE")})

            .Tables("ASTATTA3").Columns.Add("FILES", GetType(System.Int32), "COUNT(CHILD.FILENAME)")
        End With

        'Create_Lookup("GLTBANK1")

        'grdASTATTA3.DataSource = dst.Tables("ASTATTA3")
        'grdASTATTAF.DataSource = dst.Tables("ASTATTAF")

        grdASTATTA3.DataMember = "ASTATTA3"
        grdASTATTA3.DataSource = dst

        grdASTATTAF.DataMember = "ASTATTA3.ASTATTA3_ASTATTAF"
        grdASTATTAF.DataSource = dst

        Call Create_Summary(grdASTATTA3, "DOCUMENT_TYPE", "Count")
        Call Create_Summary(grdASTATTA3, "FILES")
        Call Create_Summary(grdASTATTAF, "FILENAME", "Count")

        Get_PARM("ASTPARM1")
        AS_PARM_SCAN_FOLDER = ROWs("ASTPARM1").Item("AS_PARM_SCAN_FOLDER") & ""
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

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
                EntryMode = "N"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Update"
                Call Update_Record()
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

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        SplitContainer1.Visible = tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        dst.EnforceConstraints = False
        dst.Tables("ASTATTA1").Rows.Clear()
        dst.Tables("ASTATTA2").Rows.Clear()
        dst.Tables("ASTATTA3").Rows.Clear()
        dst.Tables("ASTATTAF").Rows.Clear()
        Absx1.txtFor("AS_PARM_SCAN_FOLDER").Text = AS_PARM_SCAN_FOLDER
        pic.Image = Nothing
        dst.EnforceConstraints = True
    End Sub

    Sub Load_Record()

        Call Save_Header_Fields(UltraGroupBox1)

        dst.EnforceConstraints = False

        Call Fill_Records("ASTATTA3")
        Sort_grdColumns(grdASTATTA3, "TABLE_NAME,COLUMN_NAME,SCAN_FOLDER")

        dst.Tables("ASTATTAF").Rows.Clear()
        For Each rowASTATTA3 As DataRow In dst.Tables("ASTATTA3").Rows
            Dim SCAN_FOLDER As String = rowASTATTA3.Item("SCAN_FOLDER")

            For Each FILENAME As String In My.Computer.FileSystem.GetFiles(AS_PARM_SCAN_FOLDER & "\" & SCAN_FOLDER)
                If FILENAME.ToUpper Like "*THUMBS.DB" Then
                Else
                    Dim FI As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
                    Dim rowASTATTAF As DataRow = dst.Tables("ASTATTAF").NewRow
                    rowASTATTAF.Item("TABLE_NAME") = rowASTATTA3.Item("TABLE_NAME")
                    rowASTATTAF.Item("COLUMN_NAME") = rowASTATTA3.Item("COLUMN_NAME")
                    rowASTATTAF.Item("DOCUMENT_TYPE") = rowASTATTA3.Item("DOCUMENT_TYPE")
                    rowASTATTAF.Item("FILENAME") = FI.Name
                    rowASTATTAF.Item("FILESIZE") = FI.Length
                    rowASTATTAF.Item("FILEDATE") = FI.LastWriteTime
                    dst.Tables("ASTATTAF").Rows.Add(rowASTATTAF)
                End If
            Next
        Next
        dst.Tables("ASTATTAF").AcceptChanges()

        dst.EnforceConstraints = True
    End Sub

    Sub Delete_Record()
        Call BeginTrans()

        Stop
        'Call Delete_Rows("ARTPYMT1")
        'Call Update_Record_TDA("ARTPYMT1")

        Call CommitTrans("Delete")
    End Sub

    Sub Update_Record()
        Call BeginTrans()

        For Each rowASTATTAF As DataRow In dst.Tables("ASTATTAF").Select("CODE_VALUE IS NOT NULL")

            Dim TABLE_NAME As String = rowASTATTAF.Item("TABLE_NAME")
            Dim COLUMN_NAME As String = rowASTATTAF.Item("COLUMN_NAME")
            Dim CODE_VALUE As String = rowASTATTAF.Item("CODE_VALUE")
            Dim DESC_VALUE As String = rowASTATTAF.Item("DESC_VALUE")

            Dim rowASTATTA1 As DataRow = Fill_Record("ASTATTA1", New String() {TABLE_NAME, COLUMN_NAME, CODE_VALUE}, True, False)
            rowASTATTA1.Item("DESC_VALUE") = DESC_VALUE
            If rowASTATTA1.Item("INIT_OPER") & "" = "" Then
                rowASTATTA1.Item("INIT_DATE") = DATETIME_STAMP
                rowASTATTA1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            End If
            rowASTATTA1.Item("LAST_DATE") = DATETIME_STAMP
            rowASTATTA1.Item("LAST_OPER") = ASCMAIN1.USER_ID

            Dim rowASTATTA2 As DataRow = dst.Tables("ASTATTA2").NewRow
            rowASTATTA2.Item("TABLE_NAME") = TABLE_NAME
            rowASTATTA2.Item("COLUMN_NAME") = COLUMN_NAME
            rowASTATTA2.Item("CODE_VALUE") = CODE_VALUE
            Dim ATTACHMENT_NO As String = ASCMAIN1.Next_Control_No("ASTATTA2.ATTACHMENT_NO")
            rowASTATTA2.Item("ATTACHMENT_NO") = ATTACHMENT_NO
            Dim FILENAME As String = AS_PARM_SCAN_FOLDER & "\" & rowASTATTAF.GetParentRow("ASTATTA3_ASTATTAF").Item("SCAN_FOLDER") & "\" & rowASTATTAF.Item("FILENAME")
            rowASTATTA2.Item("ATTACHMENT_FILENAME") = FILENAME
            Dim FILENAME_SEGMENTS() As String = Split(FILENAME, ".")
            Dim FILENAME_EXT As String = FILENAME_SEGMENTS(UBound(FILENAME_SEGMENTS))
            rowASTATTA2.Item("ATTACHMENT_TYPE") = FILENAME_EXT.ToUpper
            rowASTATTA2.Item("COMPUTER_NAME") = ASCMAIN1.COMPUTER_NAME
            rowASTATTA2.Item("SESSION_NO") = ASCMAIN1.SESSION_NO
            rowASTATTA2.Item("INIT_DATE") = DATETIME_STAMP
            rowASTATTA2.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowASTATTA2.Item("LAST_DATE") = DATETIME_STAMP
            rowASTATTA2.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowASTATTA2.Item("DOCUMENT_TYPE") = rowASTATTAF.Item("DOCUMENT_TYPE")
            rowASTATTA2.Item("ATTACHMENT_DESC") = rowASTATTAF.Item("ATTACHMENT_DESC")
            My.Computer.FileSystem.CopyFile(FILENAME, ASCMAIN1.Folders("Attach") & ATTACHMENT_NO)
            My.Computer.FileSystem.DeleteFile(FILENAME)

            dst.Tables("ASTATTA2").Rows.Add(rowASTATTA2)
        Next

        For Each rowASTATTAF As DataRow In dst.Tables("ASTATTAF").Select("", "", DataViewRowState.Deleted)
            Dim FILENAME As String = AS_PARM_SCAN_FOLDER & "\" & rowASTATTAF.GetParentRow("ASTATTA3_ASTATTAF").Item("SCAN_FOLDER") & "\" & rowASTATTAF.Item("FILENAME")
            Dim FILENAME_SEGMENTS() As String = Split(FILENAME, ".")
            Dim FILENAME_EXT As String = FILENAME_SEGMENTS(UBound(FILENAME_SEGMENTS))
            My.Computer.FileSystem.DeleteFile(FILENAME)
        Next

        Call Update_Record_TDA("ASTATTA1")
        Call Update_Record_TDA("ASTATTA2")

        Call CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "PYMT_BATCH_NO"
                'sql_where = "STATUS = '0'"
        End Select

    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As System.Windows.Forms.Control)
        If COLUMN_NAME = "USER_ID" Then
            'If ctl.Text <> "" Then
            '    Call Click_Command("Load Reports")
            'End If
        End If
    End Sub

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If Absx1.GetABSColumnName(sender) = "USER_ID" Then
            'If e.KeyCode = Windows.Forms.Keys.Enter Then
            '    Call Click_Command("Load Reports", e)
            'End If
        End If
    End Sub
#End Region

    Private Sub grdASTATTA3_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTATTA3.AfterRowActivate

        'Dim dvw As DataView = dst.Tables("ASTATTAF").DefaultView
        'dvw.RowFilter = "TABLE_NAME"
        grdASTATTAF.Text = "Scanned Files in " & grdASTATTA3.ActiveRow.Cells("SCAN_FOLDER").Text
    End Sub

    Private Sub grdASTATTAF_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTATTAF.AfterCellUpdate
        With grdASTATTAF.ActiveRow
            Select Case e.Cell.Column.Key
                Case "CODE_VALUE"
                    If e.Cell.Text = "" Then
                        .Cells("DESC_VALUE").Value = ""
                    Else
                        Dim TABLE_NAME As String = grdASTATTA3.ActiveRow.Cells("TABLE_NAME").Text
                        Dim COLUMN_NAME_DESC As String = ""
                        If TABLE_NAME = "ARTCUST1" Then
                            COLUMN_NAME_DESC = "CUST_NAME"
                        End If
                        If COLUMN_NAME_DESC <> "" Then
                            .Cells("DESC_VALUE").Value = LookUp(TABLE_NAME, e.Cell.Text, True).Item(COLUMN_NAME_DESC)
                        End If
                    End If
            End Select
        End With
    End Sub

    Private Sub grdASTATTAF_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTATTAF.AfterRowActivate
        If chkAutoDisplay.Checked Then
            DisplayScannedImage()
        End If
    End Sub

    Private Sub grdASTATTAF_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdASTATTAF.BeforeCellUpdate
    End Sub

    Private Sub grdASTATTAF_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdASTATTAF.BeforeExitEditMode
        If grdASTATTAF.ActiveCell IsNot Nothing Then
            With grdASTATTAF.ActiveCell
                If .Column.Key = "CODE_VALUE" Then
                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Row.Cells("COLUMN_NAME").Text, , True)
                    End If
                End If
            End With
        End If
    End Sub

    Private Sub grdASTATTAF_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTATTAF.ClickCellButton
        Dim sql_where As String = ""
        'Call grdClickCellButton(grdARTPYMT2, sql_where, sql_where <> "")
        Dim VIEW_NAME As String = grdASTATTAF.ActiveRow.Cells("COLUMN_NAME").Text
        Dim TABLE_NAME As String = grdASTATTAF.ActiveRow.Cells("TABLE_NAME").Text
        Call View_Lookup(grdASTATTAF.ActiveCell, "CODE_VALUE", VIEW_NAME, TABLE_NAME, sql_where)
    End Sub

    Private Sub grdASTATTAF_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdASTATTAF.DoubleClickRow

        If grdASTATTAF.ActiveRow.Cells("FILENAME").Text.ToUpper Like "*.PDF" _
        Or grdASTATTAF.ActiveRow.Cells("FILENAME").Text.ToUpper Like "*.JPG" Then
            Try
                Dim FILENAME As String = AS_PARM_SCAN_FOLDER & "\" & grdASTATTA3.ActiveRow.Cells("SCAN_FOLDER").Text & "\" & grdASTATTAF.ActiveRow.Cells("FILENAME").Text
                If My.Computer.FileSystem.FileExists(FILENAME) Then
                    Dim p As Process = Process.Start(FILENAME)
                End If

            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub grdASTATTAF_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdASTATTAF.InitializeLayout

    End Sub

    Private Sub chkAutoDisplay_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAutoDisplay.CheckedChanged
        If chkAutoDisplay.Checked And grdASTATTAF.ActiveRow IsNot Nothing Then
            DisplayScannedImage()
        Else
            pic.Image = Nothing
        End If
    End Sub

    Sub DisplayScannedImage()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Scanned Image")
        pic.Image = ASCMAIN1.Get_Image(AS_PARM_SCAN_FOLDER & "\" & grdASTATTA3.ActiveRow.Cells("SCAN_FOLDER").Text & "\", grdASTATTAF.ActiveRow.Cells("FILENAME").Text)
        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub
End Class