Public Class ASFDATA1

    Dim tblASTDATA1 As New DataTable
    Dim TABLE_NAME_shown As String

    Public Sub New(ByVal FF As ASFBASE1)
        frmASFBASE1 = FF
        InitializeComponent()
    End Sub

    Private Sub ASFDSET1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If ASCMAIN1.ActiveForm Is Nothing OrElse ASCMAIN1.ActiveForm.dst Is Nothing Then
            MsgBox("Nothing to Show")
            Exit Sub
        End If

        Try
            'dst = ASCMAIN1.ActiveForm.dst
            dst.Tables.Clear()
            For Each tbl As DataTable In ASCMAIN1.ActiveForm.dst.Tables
                Dim tbl2 As DataTable = dst.Tables.Add(tbl.TableName)
                For Each dc As DataColumn In tbl.Columns
                    tbl2.Columns.Add(dc.ColumnName, dc.DataType)
                Next
                For Each row As DataRow In tbl.Select
                    tbl2.Rows.Add(row.ItemArray)
                Next
            Next
        Catch ex As Exception

        End Try

        With tblASTDATA1
            .Columns.Add("TABLE_NAME")
            .Columns.Add("TABLE_DESC")
            .Columns.Add("ROW_COUNT", GetType(System.Int64))
            .PrimaryKey = New DataColumn() {.Columns("TABLE_NAME")}
            .TableName = "ASTDATA1"
        End With

        For Each TABLE_NAME As String In frmASFBASE1.ASTDATA1s.Keys
            tblASTDATA1.Rows.Add(New Object() {TABLE_NAME, frmASFBASE1.ASTDATA1s(TABLE_NAME), 0})
        Next

        'Create_Summary(grd, grd.DisplayLayout.Bands(0).Columns(0).Key, "Count")

        If dst IsNot Nothing Then

            grdASTDATA1.DataSource = tblASTDATA1

            For Each row As DataRow In tblASTDATA1.Rows
                Dim TABLE_NAME As String = row.Item("TABLE_NAME")

                'row.Item("ROW_COUNT") = dst.Tables(TABLE_NAME).Select.Length
                row.Item("ROW_COUNT") = ASCMAIN1.ActiveForm.dst.Tables(TABLE_NAME).Select.Length

                'Dim KEY_COLUMNS As String = ""
                'If tbl.PrimaryKey.Length <> 0 Then
                '    Dim DC() As DataColumn = tbl.PrimaryKey
                '    For Each DCX As DataColumn In DC
                '        KEY_COLUMNS &= "," & DCX.ColumnName
                '    Next
                'End If
                'row0.Item("KEY_COLUMNS") = Mid(KEY_COLUMNS, 2)
            Next
        End If

        Me.Text &= " - " & frmASFBASE1.Text

        MENU_ITEM_DESC = Me.Text

        ASCMAIN1.Center(Me)
    End Sub

    Private Sub grdASTDATA1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTDATA1.AfterRowActivate
        Dim TABLE_NAME As String = grdASTDATA1.ActiveRow.Cells("TABLE_NAME").Text
        Dim TABLE_DESC As String = grdASTDATA1.ActiveRow.Cells("TABLE_DESC").Text
        grdData.Text = TABLE_DESC & " (" & TABLE_NAME & ")"

        grdData.DisplayLayout.Bands(0).Summaries.Clear()
        grdData.DataSource = Nothing
        grdData.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
        grdData.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
        'ASCMAIN1.grdInitializeLayout(grdData)
        'grdData.DataSource = dst.Tables(TABLE_NAME)
        grdData.DisplayLayout.MaxBandDepth = 1
        grdData.DataSource = ASCMAIN1.ActiveForm.dst.Tables(TABLE_NAME)

        If grdData.DataSource IsNot Nothing Then
            'C1OlapPage1.DataSource = ASCMAIN1.ActiveForm.dst.Tables(TABLE_NAME)
            'C1OlapPage1.OlapEngine.ValueFields.MaxItems = 5

            ASCMAIN1.grdInitializeLayout(grdData)
            grdData.DisplayLayout.Bands(0).SortedColumns.Clear()
            Create_Summary(grdData, grdData.DisplayLayout.Bands(0).Columns(0).Key, "Count")
            grdData.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        End If
        'Sort_grdColumns(grdData, "")

    End Sub

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        'Load_Popup_Menu(grdData, "SSBB", "Show Filter", "Show GroupBox", "Toggle Mask", "Add Total")
        Load_Popup_Menu(grdASTDATA1, "BBBS", "Export All - XLS", "Export All - CSV", "Export All - TAB", "Multi-Band")
        Load_Popup_Menu(grdData, "SS", "Show Filter", "Show GroupBox")
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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
        Else
            Select Case e.SourceControl.Name

            End Select
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key

            Case "Export All - XLS"
                Export_to_Excel_All()
            Case "Export All - CSV"
                Export_Streamwriter("CSV")
            Case "Export All - TAB"
                Export_Streamwriter("TAB")
            Case "Multi-Band"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    grdData.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.MultiBand
                Else
                    grdData.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand
                End If

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Toggle Mask"
                If grd.ActiveCell Is Nothing Then
                Else
                    Dim COLUMN_NAME As String = grd.ActiveCell.Column.Key
                    Dim TABLE_NAME As String = Mid(grd.Name, 4)
                    Dim tbl As DataTable = dst.Tables(TABLE_NAME)

                    Dim dt As System.Type = tbl.Columns(COLUMN_NAME).DataType
                    If dt.ToString = "System.Int64" Or dt.ToString = "System.Int32" Or dt.ToString = "System.Integer" Then
                        grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Format = "###,##0"
                    ElseIf dt.ToString = "System.Decimal" Then
                        If grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Format = "###,##0.00" Then
                            grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Format = "###,##0"
                        Else
                            grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Format = "###,##0.00"
                        End If
                    ElseIf dt.ToString = "System.DateTime" Then
                        If grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Format = "MM/dd/yy HH:mm:ss tt" Then
                            grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Format = "MM/dd/yy"
                        Else
                            grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Format = "MM/dd/yy HH:mm:ss tt"
                        End If

                    End If

                End If

            Case "Add Total"
                If grd.ActiveCell Is Nothing Then
                Else
                    Dim COLUMN_NAME As String = grd.ActiveCell.Column.Key
                    Create_Summary(grd, COLUMN_NAME)
                End If

        End Select
    End Sub

#End Region

    Sub Export_to_Excel_All()

        Me.Cursor = Cursors.WaitCursor
        'Dim xf As Object = Nothing
        'Dim xf As GemBox.Spreadsheet.ExcelFile = Nothing
        Dim xf As Infragistics.Documents.Excel.Workbook = Nothing

        For Each grow As UltraWinGrid.UltraGridRow In grdASTDATA1.Rows
            grow.Activate()
            If xf Is Nothing Then
                'xf = Gembox_Export_to_Excel(grdData, False)
                xf = Export_to_Excel(grdData, False)
            Else
                'Gembox_Export_to_Excel_Add_grd(xf, grdData)
                Export_to_Excel_Add_grd(xf, grdData)
            End If
        Next

        'Gembox_Export_to_Excel_Show(xf, Me.Text)
        Export_to_Excel_Show(xf, Me.Text)

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub cmdCSV_Click(sender As System.Object, e As System.EventArgs) ' Handles cmdCSV.Click


        If TABLE_NAME_shown <> "" Then

            Dim FILENAME As String = ASCMAIN1.Folders("Work") & TABLE_NAME_shown & ".CSV"
            If My.Computer.FileSystem.FileExists(FILENAME) Then
                My.Computer.FileSystem.DeleteFile(FILENAME)
            End If

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Creating CSV")

            Using sw As New System.IO.StreamWriter(FILENAME)
                Dim headings As String = ""
                Dim DEL As String = Chr(34)
                For i As Integer = 0 To dst.Tables(TABLE_NAME_shown).Columns.Count - 1

                    headings &= "," & DEL & dst.Tables(TABLE_NAME_shown).Columns(i).ColumnName & DEL
                Next
                headings = Mid(headings, 2)
                sw.WriteLine(headings)

                For Each row As DataRow In dst.Tables(TABLE_NAME_shown).Select("")
                    Dim record As String = ""


                    For i As Integer = 0 To dst.Tables(TABLE_NAME_shown).Columns.Count - 1
                        If dst.Tables(TABLE_NAME_shown).Columns(i).DataType.ToString = "System.String" Then
                            DEL = Chr(34)
                        Else
                            DEL = ""
                        End If
                        record &= "," & DEL & row.Item(i) & DEL
                    Next
                    record = Mid(record, 2)
                    sw.WriteLine(record)
                Next
            End Using

            Show_Document(FILENAME)

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")

        End If
    End Sub

    Sub Export_Streamwriter(ByVal filetype As String)
        Dim tab_delimited As Boolean
        Dim excel As New Process

        If filetype = "CSV" Then
            tab_delimited = False
        Else
            tab_delimited = True
        End If

        Dim vbquo As String = Chr(34)

        Try
            Dim FILENAME As String = ASCMAIN1.Folders("Work") & "ASFDSET1." & IIf(tab_delimited, "XLS", "CSV")
            If My.Computer.FileSystem.FileExists(FILENAME) Then
                My.Computer.FileSystem.DeleteFile(FILENAME)
            End If

            For Each grow As UltraWinGrid.UltraGridRow In grdASTDATA1.Rows
                Dim tname As String = grow.Cells(0).Value
                Dim tbl As DataTable = dst.Tables(grow.Cells(0).Value)


                Using SW As New System.IO.StreamWriter(FILENAME)
                    Dim Z As String = ""
                    For i As Integer = 0 To tbl.Columns.Count - 1
                        If tab_delimited Then
                            Z &= vbTab & vbquo & tbl.Columns(i).ColumnName & vbquo
                        Else
                            Z &= "," & vbquo & tbl.Columns(i).ColumnName & vbquo
                        End If
                    Next
                    SW.WriteLine(Mid(Z, 2))

                    For Each row As DataRow In tbl.Rows
                        Z = ""
                        For i As Integer = 0 To tbl.Columns.Count - 1
                            If tab_delimited Then
                                Z &= vbTab & vbquo & row.Item(i) & vbquo
                            Else
                                Z &= "," & vbquo & row.Item(i) & vbquo
                            End If
                        Next
                        SW.WriteLine(Mid(Z, 2))
                    Next
                End Using
            Next

            MsgBox("Success")

            excel.StartInfo.Arguments = """" + FILENAME + """ /e"
            excel.StartInfo.FileName = ASCMAIN1.Folders("Work") & FILENAME
            excel.Start()


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Failure")

        End Try

    End Sub

    Private Sub grdASTDATA1_Click(sender As Object, e As System.EventArgs) Handles grdASTDATA1.Click
    End Sub
End Class