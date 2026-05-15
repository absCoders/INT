Public Class ASFDSET1
    Dim CMs As New Dictionary(Of String, CurrencyManager)
    Dim tblColumns As New DataTable
    Dim tblRelations As New DataTable
    Dim tblRelationDataColumns As New DataTable
    Dim tblConstraints As New DataTable
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
            dst = ASCMAIN1.ActiveForm.dst
        Catch ex As Exception

        End Try


        Dim dstViewer As New DataSet

        dstViewer.Tables.Add(tblColumns)
        With tblColumns.Columns
            .Add("TableName")
            .Add("Ordinal", GetType(System.Int32))
            .Add("ColumnName")
            .Add("DataType")
            .Add("ReadOnly", GetType(System.Boolean))
            .Add("MaxLength", GetType(System.Int32))
            .Add("AllowDBNull", GetType(System.Boolean))
            .Add("Expression", GetType(System.String))
        End With

        dstViewer.Tables.Add(tblRelations)
        With tblRelations.Columns
            .Add("RelationName")
            .Add("TableNameParent")
            .Add("TableNameChild")
        End With
        With tblRelations
            .PrimaryKey = New DataColumn() {.Columns("RelationName")}
        End With

        dstViewer.Tables.Add(tblRelationDataColumns)
        With tblRelationDataColumns.Columns
            .Add("RelationName")
            .Add("Ordinal")
            .Add("ColumnNameParent")
            .Add("ColumnNameChild")
            '.Add("Ordinal", GetType(System.Int32))
            '.Add("ColumnName")
            '.Add("DataType")
            '.Add("ReadOnly", GetType(System.Boolean))
            '.Add("MaxLength", GetType(System.Int32))
            '.Add("AllowDBNull", GetType(System.Boolean))
        End With
        With tblRelationDataColumns
            .PrimaryKey = New DataColumn() {.Columns("RelationName"), .Columns("Ordinal")}
        End With

        dstViewer.Relations.Add("Relations_RelationDataColumns", _
                                  tblRelations.Columns("RelationName"), _
                                  tblRelationDataColumns.Columns("RelationName"))


        Dim tbl0 As New DataTable
        tbl0.Columns.Add("TABLE_NAME")
        tbl0.Columns.Add("ROW_COUNT", GetType(System.Int32))
        tbl0.Columns.Add("KEY_COUNT", GetType(System.Int32))
        tbl0.Columns.Add("KEY_COLUMNS")
        tbl0.Columns.Add("BINDINGS")



        If dst IsNot Nothing Then

            For Each tbl As DataTable In dst.Tables
                Dim utp As New Infragistics.Win.UltraWinTabControl.UltraTabPageControl
                Dim grd As New Infragistics.Win.UltraWinGrid.UltraGrid
                grd.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.SortMulti
                grd.DisplayLayout.Override.RowSelectors = DefaultableBoolean.Default
                grd.DisplayLayout.Override.RowSelectorHeaderStyle = UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
                grd.Name = "grd" & tbl.TableName
                grd.DisplayLayout.GroupByBox.Hidden = False
                grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True
                Call ASCMAIN1.grdInitializeLayout(grd)
                grd.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
                grd.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
                grd.DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.OutlookGroupBy
                AddHandler grd.InitializeRow, AddressOf grd_InitializeRow

                'Call ASCMAIN1.grdInitializeLayout(grd)

                Dim i As Integer = tab.Tabs.Count
                tab.Controls.Add(utp)
                tab.Tabs(i).Text = tbl.TableName
                tab.Tabs(i).Key = tbl.TableName
                utp.Controls.Add(grd)
                grd.Dock = System.Windows.Forms.DockStyle.Fill
                Initialize_Controls_for_a_Container(utp)

                Dim dvw As New DataView(tbl)
                'dvw.RowFilter = ""
                'grd.DataSource = dvw

                grd.DataMember = tbl.TableName
                grd.DataSource = dst

                grd.DisplayLayout.Override.FilterClearButtonLocation = UltraWinGrid.FilterClearButtonLocation.Row
                'grd.DisplayLayout.Override.FilterOperatorLocation = UltraWinGrid.FilterOperatorLocation.WithOperand 
                grd.DisplayLayout.Override.FilterUIType = UltraWinGrid.FilterUIType.FilterRow
                grd.DisplayLayout.Override.FilterRowAppearance.BackColor = Color.AliceBlue
                grd.DisplayLayout.Override.FilterOperatorLocation = UltraWinGrid.FilterOperatorLocation.AboveOperand

                grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                grd.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText


                Load_Popup_Menu(grd, "SSBB", "Show Filter", "Show GroupBox", "Toggle Mask", "Add Total")
                If grd.DisplayLayout.Bands.Count > 0 AndAlso grd.DisplayLayout.Bands(0).Columns.Count > 0 Then
                Create_Summary(grd, grd.DisplayLayout.Bands(0).Columns(0).Key, "Count")
                End If

                Dim X As CurrencyManager = ASCMAIN1.ActiveForm.BindingContext(tbl)
                CMs.Add(tbl.TableName, X)

                Dim row0 As DataRow = tbl0.NewRow
                row0.Item("TABLE_NAME") = tbl.TableName
                row0.Item("ROW_COUNT") = tbl.Rows.Count
                row0.Item("KEY_COUNT") = tbl.PrimaryKey.Length
                Dim KEY_COLUMNS As String = ""
                If tbl.PrimaryKey.Length <> 0 Then
                    Dim DC() As DataColumn = tbl.PrimaryKey
                    For Each DCX As DataColumn In DC
                        KEY_COLUMNS &= "," & DCX.ColumnName
                    Next
                End If
                row0.Item("KEY_COLUMNS") = Mid(KEY_COLUMNS, 2)
                tbl0.Rows.Add(row0)

                For Each col As DataColumn In tbl.Columns
                    tblColumns.Rows.Add(New Object() {tbl.TableName, col.Ordinal, _
                                                      col.ColumnName, Split(col.DataType.ToString, ".")(1), _
                                                      col.ReadOnly, col.MaxLength, col.AllowDBNull, col.Expression})
                Next

            Next


            For Each rel As DataRelation In dst.Relations
                Dim rowRelations As DataRow = tblRelations.NewRow

                rowRelations.Item("RelationName") = rel.RelationName
                rowRelations.Item("TableNameParent") = rel.ParentTable.TableName
                rowRelations.Item("TableNameChild") = rel.ChildTable.TableName
                tblRelations.Rows.Add(rowRelations)
                Dim rel_ordinal As Integer
                rel_ordinal = 0
                For Each relcolp As DataColumn In rel.ParentColumns
                    tblRelationDataColumns.Rows.Add(New Object() {rel.RelationName, rel_ordinal, relcolp.ColumnName})
                    rel_ordinal += 1
                Next
                rel_ordinal = 0
                For Each relcolc As DataColumn In rel.ChildColumns
                    Dim rowRelationDataCOlumns As DataRow = _
                    tblRelationDataColumns.Rows.Find(New Object() {rel.RelationName, rel_ordinal})
                    rowRelationDataCOlumns.Item("ColumnNameChild") = relcolc.ColumnName
                    rel_ordinal += 1
                Next
            Next

            tab.SelectedTab = tab.Tabs(0)
            tab.TabIndex = 0

            Call ASCMAIN1.grdInitializeLayout(grdTables)
            grdTables.DataSource = tbl0
            ASCMAIN1.ActiveForm.Sort_grdColumns(grdTables, "TABLE_NAME")
        End If


        grdRelations.DataSource = tblRelations
        grdRelations.DisplayLayout.Bands(1).Columns("RelationName").Hidden = True

        'Dim dvwConstraints As DataView = tblConstraints.DefaultView
        'dvwConstraints = tblColumns.DefaultView
        'dvwConstraints.RowFilter = "TableName = '" & TABLE_NAME & "'"
        'dvwConstraints.Sort = "Ordinal"
        'grdColumns.DataSource = dvwConstraints


        'Dim CBE As New Infragistics.Win.UltraWinEditors.UltraComboEditor
        'CBE.DataSource = CMs("APTINVH1")
        'UltraGrid1.DisplayLayout.Bands(0).Columns("BINDINGS").EditorControl = CBE

        'F.Sort_grdColumns(grdTables, "TABLE_NAME")

        Call ASCMAIN1.Center(Me)
    End Sub

    Private Sub grdTables_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdTables.AfterRowActivate
        Dim TABLE_NAME As String = grdTables.ActiveRow.Cells("TABLE_NAME").Text
        UltraTabControl1.Tabs(0).Text = "Columns for Table " & TABLE_NAME

        Dim dvwColumns As DataView = tblColumns.DefaultView
        dvwColumns.RowFilter = "TableName = '" & TABLE_NAME & "'"
        dvwColumns.Sort = "Ordinal"
        grdColumns.DataSource = dvwColumns

    End Sub

    Private Sub grdTables_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdTables.DoubleClickRow
        TABLE_NAME_shown = e.Row.Cells("TABLE_NAME").Text
        tab.SelectedTab = tab.Tabs(TABLE_NAME_shown)
    End Sub

    Private Sub btnSummary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSummary.Click
        tab.SelectedTab = tab.Tabs(0)
    End Sub


#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        'Call Load_Popup_Menu(grdICTTRANX, "SS", "Show Filter", "Show GroupBox")
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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
        End If



        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked

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

                'Case "Job Order Inquiry"
                '    Dim JOB_NO As String = grd.ActiveRow.Cells("JOB_NO").Text
                '    Context_Launch("Load", JOB_NO, e.Tool.Key, "DEFJOBMI")


        End Select
    End Sub

#End Region

    Private Sub grd_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs)
        Dim i As Int32 = e.Row.ListIndex
        Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        Dim TABLE_NAME As String = Mid(grd.Name, 4)
        Select Case dst.Tables(TABLE_NAME).Rows(i).RowState
            Case DataRowState.Added
                'e.Row.Appearance.BackColor = Color.LightGreen
                e.Row.RowSelectorAppearance.BackColor = Color.LightGreen
                e.Row.RowSelectorAppearance.BackColor2 = Color.Green
            Case DataRowState.Modified
                'e.Row.Appearance.BackColor = Color.LightSkyBlue
                e.Row.RowSelectorAppearance.BackColor = Color.LightSkyBlue
                e.Row.RowSelectorAppearance.BackColor2 = Color.Blue
        End Select

    End Sub

    Private Sub UltraButton1_Click(sender As System.Object, e As System.EventArgs) ' Handles UltraButton1.Click

        Dim filename As String = "C:\Users\wjz\Desktop\RGI\datasets\PODBATC1.XML"

        'dataSet.ReadXmlSchema("schema.XML");
        dst.ReadXml(filename)

        Dim tbl0 As New DataTable
        tbl0.Columns.Add("TABLE_NAME")
        tbl0.Columns.Add("ROW_COUNT", GetType(System.Int32))
        tbl0.Columns.Add("KEY_COUNT", GetType(System.Int32))
        tbl0.Columns.Add("KEY_COLUMNS")
        tbl0.Columns.Add("BINDINGS")

        For Each tbl As DataTable In dst.Tables
            Dim utp As New Infragistics.Win.UltraWinTabControl.UltraTabPageControl
            Dim grd As New Infragistics.Win.UltraWinGrid.UltraGrid
            grd.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.SortMulti
            grd.DisplayLayout.Override.RowSelectors = DefaultableBoolean.Default
            grd.DisplayLayout.Override.RowSelectorHeaderStyle = UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
            grd.Name = "grd" & tbl.TableName
            grd.DisplayLayout.GroupByBox.Hidden = False
            grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True
            Call ASCMAIN1.grdInitializeLayout(grd)
            grd.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
            grd.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
            grd.DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.OutlookGroupBy
            AddHandler grd.InitializeRow, AddressOf grd_InitializeRow

            'Call ASCMAIN1.grdInitializeLayout(grd)

            Dim i As Integer = tab.Tabs.Count
            tab.Controls.Add(utp)
            tab.Tabs(i).Text = tbl.TableName
            tab.Tabs(i).Key = tbl.TableName
            utp.Controls.Add(grd)
            grd.Dock = System.Windows.Forms.DockStyle.Fill
            Initialize_Controls_for_a_Container(utp)

            Dim dvw As New DataView(tbl)
            'dvw.RowFilter = ""
            'grd.DataSource = dvw

            grd.DataMember = tbl.TableName
            grd.DataSource = dst

            grd.DisplayLayout.Override.FilterClearButtonLocation = UltraWinGrid.FilterClearButtonLocation.Row
            'grd.DisplayLayout.Override.FilterOperatorLocation = UltraWinGrid.FilterOperatorLocation.WithOperand 
            grd.DisplayLayout.Override.FilterUIType = UltraWinGrid.FilterUIType.FilterRow
            grd.DisplayLayout.Override.FilterRowAppearance.BackColor = Color.AliceBlue
            grd.DisplayLayout.Override.FilterOperatorLocation = UltraWinGrid.FilterOperatorLocation.AboveOperand

            grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            grd.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText


            Load_Popup_Menu(grd, "SSBB", "Show Filter", "Show GroupBox", "Toggle Mask", "Add Total")
            If grd.DisplayLayout.Bands.Count > 0 AndAlso grd.DisplayLayout.Bands(0).Columns.Count > 0 Then
                Create_Summary(grd, grd.DisplayLayout.Bands(0).Columns(0).Key, "Count")
            End If

            'Dim X As CurrencyManager = ASCMAIN1.ActiveForm.BindingContext(tbl)
            'CMs.Add(tbl.TableName, X)

            Dim row0 As DataRow = tbl0.NewRow
            row0.Item("TABLE_NAME") = tbl.TableName
            row0.Item("ROW_COUNT") = tbl.Rows.Count
            row0.Item("KEY_COUNT") = tbl.PrimaryKey.Length
            Dim KEY_COLUMNS As String = ""
            If tbl.PrimaryKey.Length <> 0 Then
                Dim DC() As DataColumn = tbl.PrimaryKey
                For Each DCX As DataColumn In DC
                    KEY_COLUMNS &= "," & DCX.ColumnName
                Next
            End If
            row0.Item("KEY_COLUMNS") = Mid(KEY_COLUMNS, 2)
            tbl0.Rows.Add(row0)

            For Each col As DataColumn In tbl.Columns
                tblColumns.Rows.Add(New Object() {tbl.TableName, col.Ordinal, _
                                                  col.ColumnName, Split(col.DataType.ToString, ".")(1), _
                                                  col.ReadOnly, col.MaxLength, col.AllowDBNull, col.Expression})
            Next

        Next


        For Each rel As DataRelation In dst.Relations
            Dim rowRelations As DataRow = tblRelations.NewRow

            rowRelations.Item("RelationName") = rel.RelationName
            rowRelations.Item("TableNameParent") = rel.ParentTable.TableName
            rowRelations.Item("TableNameChild") = rel.ChildTable.TableName
            tblRelations.Rows.Add(rowRelations)
            Dim rel_ordinal As Integer
            rel_ordinal = 0
            For Each relcolp As DataColumn In rel.ParentColumns
                tblRelationDataColumns.Rows.Add(New Object() {rel.RelationName, rel_ordinal, relcolp.ColumnName})
                rel_ordinal += 1
            Next
            rel_ordinal = 0
            For Each relcolc As DataColumn In rel.ChildColumns
                Dim rowRelationDataCOlumns As DataRow = _
                tblRelationDataColumns.Rows.Find(New Object() {rel.RelationName, rel_ordinal})
                rowRelationDataCOlumns.Item("ColumnNameChild") = relcolc.ColumnName
                rel_ordinal += 1
            Next
        Next

        tab.SelectedTab = tab.Tabs(0)
        tab.TabIndex = 0

        Call ASCMAIN1.grdInitializeLayout(grdTables)
        grdTables.DataSource = tbl0
        ASCMAIN1.ActiveForm.Sort_grdColumns(grdTables, "TABLE_NAME")

    End Sub

    Private Sub cmdCSV_Click(sender As System.Object, e As System.EventArgs) 'Handles cmdCSV.Click
        If tab.SelectedTab.Index <> 0 Then

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

    Private Sub tab_SelectedTabChanged(sender As Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab.SelectedTabChanged
        TABLE_NAME_shown = e.Tab.Text
    End Sub
End Class