Imports ABSolution
Imports Infragistics.Win
Imports System.Windows.forms

Public Class EDFSQLI1

    Dim viewCatalogTables As New DataView
    Private Const noUPDATE_TYPE = "0"
    Private Const noUPDATE_FREQUENCY = "N"
    Dim sqlConnectString As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("EDTPARM1")
        Dim sql As String

        With dst

            Create_TDA(.Tables.Add, "EDTSQLI1", "*", 0, True, String.Empty, 2, String.Empty)
            .Tables("EDTSQLI1").Columns.Add("INVALID", GetType(System.String))
            .Tables("EDTSQLI1").Columns("INVALID").DefaultValue = False

            .Tables("EDTSQLI1").Columns("UPDATE_TYPE").DefaultValue = noUPDATE_TYPE
            .Tables("EDTSQLI1").Columns("UPDATE_FREQUENCY").DefaultValue = noUPDATE_FREQUENCY

            sql = "Select LPAD(' ', 100) CATALOG from dual"
            Create_TDA(.Tables.Add, "dbs", sql, 0, False, String.Empty, 0, String.Empty)

            sql = "Select LPAD(' ', 100) CATALOG, LPAD(' ', 100) TABLE_NAME from dual"
            Create_TDA(.Tables.Add, "tbls", sql, 0, False, String.Empty, 0, String.Empty)

            Create_TDA(.Tables.Add, "EDTSQLI2", "*")

        End With

        grdEDTSQLI1.DataSource = dst.Tables("EDTSQLI1")

        Create_Summary(grdEDTSQLI1, "TABLE_NAME", "Count")
        Create_Summary(grdEDTSQLI1, "SELECTED")
        Create_Summary(grdEDTSQLI1, "MINUTES")
        Create_Summary(grdEDTSQLI1, "ROW_COUNT")
        Create_Summary(grdEDTSQLI1, "COL_COUNT")

        With grdEDTSQLI1.DisplayLayout.Bands("EDTSQLI1")
            '.Columns("UPDATE_TYPE").Header.Fixed = True
            '.Columns("UPDATE_FREQUENCY").Header.Fixed = True
            .Columns("CATALOG").Header.Fixed = True
            .Columns("TABLE_NAME").Header.Fixed = True
        End With

        grdEDTSQLI1.DisplayLayout.MaxColScrollRegions = 1
        grdEDTSQLI1.DisplayLayout.MaxRowScrollRegions = 1

        If Not ASCMAIN1.Running_in_VS Then
            optSS.CheckedIndex = 1
        End If

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

            Case "Import Data"

            Case "Update Table Info"

            Case "Update Sql Server"
                EMsg = "Proceeding deletes all records in the EDTSQLI1 table for all Catalogues."
                EMsg &= " Only valid entries (Update Type <> None, Update Frequency <> None, Catalogue, Table Name) will be uploaded."
                EMsg &= Environment.NewLine & "Do you want to proceed?"
                If MessageBox.Show(EMsg, "Update SQL Server", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If
                EMsg = String.Empty
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

            Case "Import Data"
                Call Import_Data()
                Call Mode_Settings(True)

            Case "Update Table Info"
                Try
                    Me.grdEDTSQLI1.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
                Catch ex As Exception

                End Try
                Update_Record_TDA("EDTSQLI1")
                MsgBox("Table Info Updated", MsgBoxStyle.OkOnly, "Verification")
                Call Mode_Settings(True)

            Case "Update Sql Server"
                Call Update_SQL_Server()
                MsgBox("SQL Server is Updated", MsgBoxStyle.OkOnly, "Verification")
                Call Mode_Settings(False)

            Case "SQL Scripts"
                SQL_Scripts()

            Case "Done"
                Call Mode_Settings(False)

            Case "Select All"
                Select_All()

            Case "De-Select All"
                DeSelect_All()

            Case "Select New"
                Select_New()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Import Data").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Update Table Info").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("SQL Scripts").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update SQL Server").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode

                .Groups("Table Selection").Visible = ScreenMode

                Setup_Tab()
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        UltraTabControl1.Visible = tf
        UltraGroupBox1.Visible = Not tf

        If ScreenMode Then
        Else
            Clear_Record()
            UltraGrid1.Visible = False
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("EDTSQLI1").Rows.Clear()
        dst.EnforceConstraints = True

        chkUpdateCosts.Checked = False
        chkUpdateCosts.Text = "Update Costs for " & ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1), False, False)

        'ASCMAIN1.sql = "Select * from TATCTLN1 where CTL_NO_TYPE = 'EDFSQLI1.FORM_INSTANCE_NO'"
        'Dim row As DataRow = ASCDATA1.GetDataRow
        'If row.Item("CTL_NO_TEXT") & "" = "" OrElse Val(row.Item("CTL_NO_TEXT") & "") < ASCMAIN1.CYP Then
        '    If Format(Now, "dd") > 7 Then
        '        chkUpdateCosts.Visible = True
        '        chkUpdateCosts.Checked = True
        '        chkUpdateCosts.Text = "Update Costs for " & ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1), False, False)
        '    End If
        'End If
    End Sub

    Sub Load_Record()

        Call ASCMAIN1.Progress("Now Loading Data")
        Call Save_Header_Fields(UltraGroupBox1)

        Load_Meta_Data()

        MyBase.dst.EnforceConstraints = False
        MyBase.Fill_Records("EDTSQLI1")
        Sort_grdColumns(grdEDTSQLI1, "CATALOG,TABLE_NAME")

        'If optSS.CheckedIndex = 0 Then
        '    sqlConnectString = "Data Source=.;Integrated Security=SSPI;"
        'Else
        '    sqlConnectString = "User ID=sa;Password=jhinyc;Server=172.19.0.75;Data Source=.;Integrated Security=SSPI;"
        'End If

        Using sqlConnection As New System.Data.SqlClient.SqlConnection(sqlConnectString)
            sqlConnection.Open()

            For Each rowEDTSQLI1 As DataRow In dst.Tables("EDTSQLI1").Rows
                Dim CATALOG = rowEDTSQLI1.Item("CATALOG") & String.Empty
                Dim TABLE_NAME = rowEDTSQLI1.Item("TABLE_NAME") & String.Empty
                ' If CATALOG = "NAV" Then Stop

                If CATALOG = "AHNA" Or CATALOG = "NYAG" Then
                Else
                    'Dim sql As String = "Select * from [" & CATALOG & "].[dbo].[EDTSQLI1] where TABLE_NAME = '" & TABLE_NAME & "'"
                    'Dim sqlCommand As New System.Data.SqlClient.SqlCommand(sql, sqlConnection)
                    'Using rdr As System.Data.SqlClient.SqlDataReader = sqlCommand.ExecuteReader
                    '    If rdr.HasRows Then
                    '        rdr.Read()
                    '        rowEDTSQLI1.Item("INIT_DATE") = rdr.Item("LAST_DATE")
                    '        rowEDTSQLI1.Item("INIT_OPER") = rdr.Item("LAST_OPER")
                    '    End If
                    'End Using

                End If


                'Dim ADA As New SqlClient.SqlDataAdapter(sqlCommand)
                'Dim TBL As New DataTable
                'ADA.Fill(TBL)


                'rowEDTSQLI1.Item("INVALID") = dst.Tables("tbls").Select("CATALOG = '" & CATALOG & "' AND TABLE_NAME = '" & TABLE_NAME & "'").Length = 0
            Next

            sqlConnection.Close()
        End Using

        MyBase.dst.EnforceConstraints = True
        Call ASCMAIN1.Progress("")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdEDTSQLI1, "SS", "Show Filter", "Show GroupBox")
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

            'Case "Track Shipment"
            '    If grd.ActiveRow.Cells("SHIP_REF").Text <> "" Then
            '        Me.Cursor = Cursors.WaitCursor
            '        Call ASCMAIN1.Progress("Now Locating DHL POD")
            '        System.Diagnostics.Process.Start("http://track.dhl-usa.com/TrackByNbr.asp?ShipmentNumber=" & grd.ActiveRow.Cells("SHIP_REF").Text)
            '        Me.Cursor = Cursors.Default
            '        Call ASCMAIN1.Progress("")
            '    End If

            'Case "Job Order Inquiry"
            '    Dim JOB_NO As String = grd.ActiveRow.Cells("JOB_NO").Text
            '    Context_Launch("Load", JOB_NO, e.Tool.Key, "DEFJOBMI")

            Case Else
                'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
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
        Generate_Report("EDRSQLI", "SQL Server Data Import", SUBT, RecordSelectionFormula)
        Call Print_Report_End()
    End Sub

    Private Sub UltraTabControl1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged
        Setup_Tab()
    End Sub

    Sub Setup_Tab()
        If SELECTION_NO = 0 Then Exit Sub

        If UltraTabControl1.SelectedTab Is Nothing Then Exit Sub
        'With UltraExplorerBar1
        '    If UltraTabControl1.SelectedTab.Key = UltraTabControl1.Tabs("Import Raw EDI Files").Key Then
        '        .Groups("Screen Control").Items("Import Raw EDI Files").Settings.Enabled = DefaultableBoolean.True
        '    Else
        '        .Groups("Screen Control").Items("Import Raw EDI Files").Settings.Enabled = DefaultableBoolean.False
        '    End If
        '    .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
        'End With
    End Sub

    Sub Import_Data()

        ASCMAIN1.Progress("Now Importing Data")
        Me.Cursor = Cursors.WaitCursor

        Dim oracleTABLE_NAME As String = String.Empty
        Dim oracleSCHEMA_NAME As String = "CONV"

        UltraGrid1.Visible = True

        For Each rowEDTSQLI1 As DataRow In dst.Tables("EDTSQLI1").Select("SELECTED = '1'", "TABLE_NAME")
            Dim TABLE_NAME As String = rowEDTSQLI1.Item("TABLE_NAME")
            Dim CATALOG As String = rowEDTSQLI1.Item("CATALOG")

            oracleTABLE_NAME = Replace(Replace(CATALOG & "_" & TABLE_NAME, " ", "_"), "-", "_")
            If oracleTABLE_NAME.ToString.Length > 30 Then
                oracleTABLE_NAME = oracleTABLE_NAME.ToString.Substring(0, 30)
            End If

            Dim cs As String = "Data Source=.;Initial Catalog=" & CATALOG & ";Integrated Security=SSPI;"
            cs = sqlConnectString
            Using conn As New System.Data.SqlClient.SqlConnection(cs)

                Dim INIT_DATE As Date = Now
                UltraGrid1.DataSource = Nothing

                ASCMAIN1.Progress("Now Importing Data from " & CATALOG & ":" & TABLE_NAME)

                ASCDATA1.ExecuteSQL("Drop Table " & oracleSCHEMA_NAME & "." & oracleTABLE_NAME, True)

                Dim sql As String = "Select * from " & CATALOG & ".dbo.[" & TABLE_NAME & "]"
                Dim sc As New System.Data.SqlClient.SqlCommand(sql, conn)
                Dim da As New System.Data.SqlClient.SqlDataAdapter(sql, conn)
                Dim tbl As New DataTable
                da.FillSchema(tbl, SchemaType.Source)

                If tbl.Columns.Contains("DESC") Then
                    tbl.Columns("DESC").ColumnName = "DESCRIPTION"
                End If
                If tbl.Columns.Contains("EXCLUSIVE") Then
                    tbl.Columns("EXCLUSIVE").ColumnName = "EXCLUSIVE_X"
                End If

                Dim IGNORE As New List(Of String)

                Dim ddl As String = ""
                For Each dc As DataColumn In tbl.Columns
                    Dim COLUMN_NAME As String = Replace(Replace(Replace(Replace(Replace(dc.ColumnName.ToUpper, " ", "_"), "%", "PCT"), "-", "_"), ")", ""), "(", "")
                    Do While InStr(COLUMN_NAME, "__") <> 0
                        COLUMN_NAME = Replace(COLUMN_NAME, "__", "_")
                    Loop
                    'If COLUMN_NAME = "DESC" Then COLUMN_NAME = "DESCRIPTION"
                    'If COLUMN_NAME = "EXCLUSIVE" Then COLUMN_NAME = "EXCLUSIVE_X"
                    'ddl &= ", " & Chr(34) & COLUMN_NAME & Chr(34) & " "
                    ddl &= ", " & Chr(34) & COLUMN_NAME & Chr(34) & " "

                    Select Case dc.DataType.ToString
                        Case "System.String"
                            Dim MAXLEN As Integer = dc.MaxLength
                            If MAXLEN > 2000 Then
                                MAXLEN = 2000
                            End If
                            ddl &= "VARCHAR2(" & CStr(MAXLEN) & ")"
                        Case "System.Integer"
                            Stop
                        Case "System.Decimal"
                            ddl &= "NUMBER(20,2)"
                        Case "System.Decimal[]"
                            'ddl &= "NUMBER(13,2)"
                            IGNORE.Add(COLUMN_NAME)
                            ddl &= "VARCHAR2(1)"
                        Case "System.DateTime"
                            ddl &= "DATE"
                        Case "System.Int16"
                            ddl &= "NUMBER(6,0)"
                        Case "System.Int32"
                            ddl &= "NUMBER(10,0)"
                        Case "System.Int64"
                            ddl &= "NUMBER(14,0)"
                        Case "System.Byte"
                            ddl &= "NUMBER"
                        Case "System.Byte[]"
                            'ddl &= "LONG"
                            IGNORE.Add(COLUMN_NAME)
                            ddl &= "VARCHAR2(1)"
                        Case "System.Boolean"
                            ddl &= "NUMBER(3,0)"
                        Case "System.Double"
                            ddl &= "NUMBER(20,6)"
                        Case "System.Single"
                            ddl &= "NUMBER(10,4)"
                        Case Else
                            Stop

                    End Select
                    ddl &= vbCr
                Next

                'UltraGrid1.BindingContext = Nothing
                UltraGrid1.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
                UltraGrid1.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
                UltraGrid1.DataSource = tbl
                UltraGrid1.Text = TABLE_NAME & " - 1st 100 Rows"
                Application.DoEvents()

                ddl = "Create Table " & oracleSCHEMA_NAME & "." & oracleTABLE_NAME & "(" & Mid(ddl, 3) & ")"
                ASCDATA1.ExecuteSQL(ddl)

                'dst.Tables.Remove(oracleTABLE_NAME)
                If dst.Tables.Contains(oracleTABLE_NAME) Then
                    dst.Tables(oracleTABLE_NAME).Rows.Clear()
                Else
                    Create_TDA(dst.Tables.Add(oracleTABLE_NAME), oracleSCHEMA_NAME & "." & oracleTABLE_NAME, "*")
                End If

                Dim r As Integer = 0
                conn.Open()

                Dim use_binding As Boolean = True
                Create_BAs(oracleTABLE_NAME)

                Using sr As System.Data.SqlClient.SqlDataReader = sc.ExecuteReader
                    While sr.Read
                        r += 1

                        Dim row As DataRow = dst.Tables(oracleTABLE_NAME).NewRow
                        For c As Integer = 0 To tbl.Columns.Count - 1
                            row.Item(c) = sr.Item(c)
                        Next
                        If IGNORE.Count <> 0 Then
                            For Each COLUMN_NAME In IGNORE
                                row.Item(COLUMN_NAME) = DBNull.Value
                            Next
                        End If
                        dst.Tables(oracleTABLE_NAME).Rows.Add(row)
                        If r <= 100 Then
                            Dim rowx As DataRow = tbl.NewRow
                            For c As Integer = 0 To tbl.Columns.Count - 1
                                rowx.Item(c) = sr.Item(c)
                            Next
                            tbl.Rows.Add(rowx)
                        End If

                        If r Mod 1000 = 0 Then
                            Application.DoEvents()
                            If use_binding Then
                                Update_BAs(oracleTABLE_NAME)
                            Else
                                Update_Record_TDA(oracleTABLE_NAME)
                            End If
                            dst.Tables(oracleTABLE_NAME).Rows.Clear()
                            ASCMAIN1.Progress("Now Importing Data from " & CATALOG & ":" & TABLE_NAME, CStr(r))
                        End If

                    End While
                End Using
                conn.Close()

                If use_binding Then
                    If r Mod 1000 <> 0 Then ' THERE WERE ROWS LEFT
                        Update_BAs(oracleTABLE_NAME)
                    End If
                Else
                    Update_Record_TDA(oracleTABLE_NAME)
                End If

                dst.Tables(oracleTABLE_NAME).Rows.Clear()

                Dim LAST_DATE As Date = Now

                rowEDTSQLI1.Item("ROW_COUNT") = r
                rowEDTSQLI1.Item("COL_COUNT") = tbl.Columns.Count
                rowEDTSQLI1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowEDTSQLI1.Item("INIT_DATE") = INIT_DATE
                rowEDTSQLI1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowEDTSQLI1.Item("LAST_DATE") = LAST_DATE
                rowEDTSQLI1.Item("MINUTES") = LAST_DATE.Subtract(INIT_DATE).TotalMinutes
                rowEDTSQLI1.Item("DDL") = Mid(ddl, 1, 3000)
            End Using
        Next
        'ASCMAIN1.AnalyzeTable(oracleTABLE_NAME, oracleSCHEMA_NAME)
        Update_Record_TDA("EDTSQLI1")
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("")

        UltraGrid1.Visible = False
        MsgBox("Import Complete", MsgBoxStyle.OkOnly, "Verification")
    End Sub

    Sub Update_SQL_Server()

        Dim sql As String = String.Empty
        Dim deletedRecords As Integer = 0

        ' Update Orcle before hand.
        Update_Record_TDA("EDTSQLI1")

        Dim tblSqlCommand As System.Data.SqlClient.SqlCommand
        Dim tblConnection As New System.Data.SqlClient.SqlConnection(sqlConnectString)
        tblConnection.Open()

        For Each rowdbs As DataRow In dst.Tables("dbs").Rows
            Dim CATALOG As String = rowdbs.Item("CATALOG")
            Dim TABLE_COUNT As Int32 = dst.Tables("EDTSQLI1").Select("CATALOG = '" & CATALOG & "'").Length
            If CATALOG <> "" And TABLE_COUNT <> 0 Then

                Dim sqlada As New System.Data.SqlClient.SqlDataAdapter("Select * from [" & CATALOG & "].[dbo].[EDTSQLI1]", sqlConnectString)
                Dim sqlbld As New System.Data.SqlClient.SqlCommandBuilder(sqlada)
                sqlada.InsertCommand = sqlbld.GetInsertCommand(True)
                Dim DT As New DataTable
                sqlada.Fill(DT)
                DT.Rows.Clear()
                ASCMAIN1.Progress("Process Catalog: " & CATALOG, String.Empty)

                Try
                    sql = "DELETE FROM [" & CATALOG & "].[dbo].[EDTSQLI1]"
                    tblSqlCommand = New System.Data.SqlClient.SqlCommand(sql, tblConnection)
                    deletedRecords = tblSqlCommand.ExecuteNonQuery()

                    For Each row As DataRow In dst.Tables("EDTSQLI1").Select _
                    ("CATALOG = '" & CATALOG & "'") '  and UPDATE_TYPE <> '0' and UPDATE_FREQUENCY <> 'N'")
                        Dim rowDT As DataRow = DT.NewRow
                        'rowDT.ItemArray = row.ItemArray
                        For I As Int32 = 0 To DT.Columns.Count - 1
                            rowDT.Item(I) = row.Item(I)
                        Next
                        rowDT.Item("LAST_OPER") = row.Item("INIT_OPER")
                        rowDT.Item("LAST_DATE") = row.Item("INIT_DATE")
                        DT.Rows.Add(rowDT)
                    Next

                    Application.DoEvents()
                    sqlada.Update(DT)
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If

        Next
        ASCMAIN1.Progress("", String.Empty)
        tblConnection.Close()
    End Sub

    Private Sub grdEDTSQLI1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdEDTSQLI1.AfterRowActivate
        If grdEDTSQLI1.ActiveRow Is Nothing Then
            Exit Sub
        End If

        If grdEDTSQLI1.ActiveRow.IsAddRow Then
            grdEDTSQLI1.DisplayLayout.Bands(0).Columns("CATALOG").CellActivation = UltraWinGrid.Activation.AllowEdit
            grdEDTSQLI1.DisplayLayout.Bands(0).Columns("TABLE_NAME").CellActivation = UltraWinGrid.Activation.AllowEdit
            grdEDTSQLI1.DisplayLayout.Bands(0).Columns("TABLE_DESC").CellActivation = UltraWinGrid.Activation.AllowEdit
            viewCatalogTables.RowFilter = "CATALOG = '**'"
        Else
            grdEDTSQLI1.DisplayLayout.Bands(0).Columns("CATALOG").CellActivation = UltraWinGrid.Activation.NoEdit
            grdEDTSQLI1.DisplayLayout.Bands(0).Columns("TABLE_NAME").CellActivation = UltraWinGrid.Activation.NoEdit
            grdEDTSQLI1.DisplayLayout.Bands(0).Columns("TABLE_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
            If grdEDTSQLI1.ActiveRow Is Nothing Then
                viewCatalogTables.RowFilter = ""
            Else
                Try
                    '  viewCatalogTables.RowFilter = "CATALOG = '" & grdEDTSQLI1.ActiveRow.Cells("CATALOG").Text & "'"
                Catch ex As Exception

                End Try
            End If
        End If


    End Sub

    Private Sub grdEDTSQLI1_AfterCellActivate(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles grdEDTSQLI1.AfterCellActivate
        If grdEDTSQLI1.ActiveRow Is Nothing Then
            Exit Sub
        End If

        If grdEDTSQLI1.ActiveCell.Column.Key = "TABLE_NAME" Then
            viewCatalogTables.RowFilter = "CATALOG = '" & grdEDTSQLI1.ActiveRow.Cells("CATALOG").Text & "'"
        End If

    End Sub

    Private Sub Form_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        'Load_Record()
        'Call Mode_Settings(True)
    End Sub

    Sub Select_All()
        For Each rowEDTSQLI1 As DataRow In dst.Tables("EDTSQLI1").Rows
            rowEDTSQLI1.Item("SELECTED") = "1"
        Next
    End Sub

    Sub DeSelect_All()
        For Each rowEDTSQLI1 As DataRow In dst.Tables("EDTSQLI1").Select("", "", DataViewRowState.CurrentRows)
            rowEDTSQLI1.Item("SELECTED") = "0"
        Next
    End Sub

    Sub Select_New()
        For Each rowEDTSQLI1 As DataRow In dst.Tables("EDTSQLI1").Rows
            If rowEDTSQLI1.Item("INIT_DATE") & "" = "" Or rowEDTSQLI1.Item("LAST_DATE") & "" = "" Then
                If rowEDTSQLI1.Item("LAST_DATE") & "" = "" Then
                    rowEDTSQLI1.Item("SELECTED") = "1"
                End If
            Else
                Dim INIT_DATE As String = Format(rowEDTSQLI1.Item("INIT_DATE"), "yyyyMMddHHmm")
                Dim LAST_DATE As String = Format(rowEDTSQLI1.Item("LAST_DATE"), "yyyyMMddHHmm")
                If INIT_DATE > LAST_DATE Then ' LAST_DATE from sql server is loaded into INIT_DATE in this work table
                    rowEDTSQLI1.Item("SELECTED") = "1"
                End If
            End If
        Next
    End Sub

    Private Sub grdEDTSQLI1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdEDTSQLI1.InitializeLayout

        Dim vlUPDATE_TYPE As ValueList

        If (Not e.Layout.ValueLists.Exists("vlUPDATE_TYPE")) Then
            vlUPDATE_TYPE = e.Layout.ValueLists.Add("vlUPDATE_TYPE")
            vlUPDATE_TYPE.ValueListItems.Add("0", "None")
            vlUPDATE_TYPE.ValueListItems.Add("1", "Type 1")
            vlUPDATE_TYPE.ValueListItems.Add("2", "Type 2")
            vlUPDATE_TYPE.ValueListItems.Add("3", "Type 3")
        End If
        e.Layout.Bands(0).Columns("UPDATE_TYPE").ValueList = e.Layout.ValueLists("vlUPDATE_TYPE")

        Dim vlUPDATE_FREQUENCY As ValueList

        If (Not e.Layout.ValueLists.Exists("vlUPDATE_FREQUENCY")) Then
            vlUPDATE_FREQUENCY = e.Layout.ValueLists.Add("vlUPDATE_FREQUENCY")
            vlUPDATE_FREQUENCY.ValueListItems.Add("N", "None")
            vlUPDATE_FREQUENCY.ValueListItems.Add("D", "Daily")
            vlUPDATE_FREQUENCY.ValueListItems.Add("W", "Weekly")
            vlUPDATE_FREQUENCY.ValueListItems.Add("M", "Monthly")
        End If
        e.Layout.Bands(0).Columns("UPDATE_FREQUENCY").ValueList = e.Layout.ValueLists("vlUPDATE_FREQUENCY")

        e.Layout.Bands(0).Columns("CATALOG").ValueList = Me.uddCatalog
        e.Layout.Bands(0).Columns("TABLE_NAME").ValueList = Me.uddTables
    End Sub

    Private Sub grdEDTSQLI1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdEDTSQLI1.InitializeRow
        If e.Row.Cells("INVALID").Value = True Then
            'e.Row.Appearance.BackColor = Drawing.Color.Red
            'e.Row.Appearance.ForeColor = Drawing.Color.White

            For Each gCell As UltraWinGrid.UltraGridCell In e.Row.Cells
                gCell.Appearance.BackColor = Drawing.Color.Red
                gCell.Appearance.ForeColor = Drawing.Color.White
            Next

        Else
            For Each gCell As UltraWinGrid.UltraGridCell In e.Row.Cells
                e.Row.Cells("INVALID").Appearance.Reset()
            Next
        End If

        If e.Row.Cells("LAST_DATE").Value & "" <> "" And _
            e.Row.Cells("INIT_DATE").Value & "" <> "" Then
            Dim LAST_DATE As String = Format(e.Row.Cells("LAST_DATE").Value, "yyyyMMddHHmm")
            Dim INIT_DATE As String = Format(e.Row.Cells("INIT_DATE").Value, "yyyyMMddHHmm")
            If INIT_DATE > LAST_DATE Then
                e.Row.Cells("INIT_DATE").Appearance.ForeColor = Drawing.Color.Green
            End If
        End If

    End Sub

    Sub Load_Meta_Data()

        Dim sql As String = String.Empty
        Dim Catalog As String = String.Empty
        Dim TABLE_NAME As String = String.Empty


        sql = "Select * from master.dbo.sysdatabases" '  where cmptlevel = 80"
        If optSS.CheckedIndex = 0 Then
            sqlConnectString = "Data Source=.;Integrated Security=SSPI;"
        Else
            'sqlConnectString = "User ID=sa;Password=jhinyc;Server=172.19.0.75;Data Source=.;" ' Integrated Security=SSPI;"
            'sqlConnectString = "User ID=abs;Password=office;Server=172.19.0.75;Data Source=.;" ' Integrated Security=SSPI;"
            sqlConnectString = "Data Source=172.19.0.75;Initial Catalog=JHC;Persist Security Info=True;User ID=abs;Password=office"
            sqlConnectString = "Data Source=172.17.0.23;Initial Catalog=JHC;Persist Security Info=True;User ID=sa;Password=jhinyc"
        End If
        sqlConnectString = "Data Source=10.0.1.2;Persist Security Info=True;User ID=sa;Password=NYAGahava"
        sqlConnectString = "Data Source=.;Integrated Security=SSPI;"

        Dim dbsConnection As New System.Data.SqlClient.SqlConnection(sqlConnectString)
        Dim dbsSqlCommand As New System.Data.SqlClient.SqlCommand(Sql, dbsConnection)

        dbsConnection.Open()

        dst.Tables("dbs").Rows.Clear()
        dst.Tables("tbls").Rows.Clear()

        Using dbsSR As System.Data.SqlClient.SqlDataReader = dbsSqlCommand.ExecuteReader
            While dbsSR.Read
                Catalog = dbsSR.Item("Name")
                Dim rowDBS As DataRow = dst.Tables("dbs").NewRow
                rowDBS.Item("CATALOG") = Catalog
                dst.Tables("dbs").Rows.Add(rowDBS)

                ' Get the tables for this Catalog
                Sql = "SELECT TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE"
                sql &= " FROM [" & Catalog & "].INFORMATION_SCHEMA.TABLES"
                Sql &= " WHERE TABLE_CATALOG = '" & Catalog & "'"
                Sql &= " AND (TABLE_TYPE = 'BASE TABLE') AND (TABLE_NAME <> 'dtproperties') AND (TABLE_NAME <> 'EDTSQLI1')"

                Dim tblConnection As New System.Data.SqlClient.SqlConnection(sqlConnectString)
                Dim tblSqlCommand As New System.Data.SqlClient.SqlCommand(Sql, tblConnection)

                tblConnection.Open()

                Using tblSR As System.Data.SqlClient.SqlDataReader = tblSqlCommand.ExecuteReader
                    While tblSR.Read
                        TABLE_NAME = tblSR.Item("TABLE_NAME")
                        Dim rowTBLS As DataRow = dst.Tables("tbls").NewRow
                        rowTBLS.Item("CATALOG") = Catalog
                        rowTBLS.Item("TABLE_NAME") = TABLE_NAME
                        dst.Tables("tbls").Rows.Add(rowTBLS)
                    End While
                End Using

                tblConnection.Close()

            End While
        End Using

        dbsConnection.Close()

        Me.uddCatalog.SetDataBinding(dst.Tables("dbs"), Nothing)
        Me.uddCatalog.ValueMember = "CATALOG"
        Me.uddCatalog.DisplayMember = "CATALOG"

        viewCatalogTables = New DataView(dst.Tables("tbls"))
        viewCatalogTables.Sort = "TABLE_NAME"
        viewCatalogTables.RowFilter = "CATALOG = '**'"

        ' uddTables
        Me.uddTables.SetDataBinding(viewCatalogTables, Nothing)
        Me.uddTables.ValueMember = "TABLE_NAME"
        Me.uddTables.DisplayMember = "TABLE_NAME"
    End Sub

    Sub SQL_Scripts()

        dst.Tables("EDTSQLI2").Rows.Clear()

        For Each XXX As String In New String() {"JHU", "JHI", "JHC"}
            ASCMAIN1.sql = "" _
            & "DELETE FROM JHX." & XXX & "_General_Posting_Setup" _
            & " WHERE Gen_Bus_Posting_Group IS NULL OR Gen_Prod_Posting_Group IS NULL"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "" _
            & "ALTER TABLE JHX." & XXX & "_General_Posting_Setup" _
            & " ADD PRIMARY KEY (GEN_BUS_POSTING_GROUP," _
            & " GEN_PROD_POSTING_GROUP);"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, True)
        Next

        If chkUpdateCosts.Checked Then
            ASCMAIN1.sql = "UPDATE TATCTLN1 " _
            & " SET CTL_NO_TEXT = '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1) & "'" _
            & " WHERE TATCTLN1.CTL_NO_TYPE = 'EDFSQLI1.FORM_INSTANCE_NO'"
            ASCDATA1.ExecuteSQL()
        End If

        Dim FILENAME As String = ASCMAIN1.Folders("Images") & "\CONVERT_SQL.TXT"
        If ASCMAIN1.Running_in_VS Then
            FILENAME = "C:\Users\wjz\Desktop\Clients\JHI\wjz\CONVERT_SQL.TXT"
        End If

        Dim CONVERT_SQL As String = ""
        Using sr As New System.IO.StreamReader(FILENAME)
            CONVERT_SQL = sr.ReadToEnd
        End Using

        Dim SQLs() As String = Split(CONVERT_SQL, vbCrLf & vbCrLf)

        Dim SQL_XNO As String = ASCMAIN1.Next_Control_No("EDTSQLI2.SQL_XNO")
        Dim SQL_LNO As Int32 = 0
        Dim SQL_RESULT As Int32 = 0
        Dim SQL_COMMENT As String = ""

        Me.Cursor = Cursors.WaitCursor

        Dim SQL_2_EXECUTE As String = ""

        Try

            For Each SQL As String In SQLs
                SQL = SQL.Trim

                If SQL Like ("-- End*") Then Exit For

                If SQL Like ("-- *") Then
                    SQL_COMMENT = Mid(SQL, 1, InStr(SQL & vbCrLf, vbCrLf) - 1)
                    SQL_COMMENT = Mid(SQL_COMMENT, 4)
                    ASCMAIN1.Progress(SQL_COMMENT, "")
                    SQL = Mid(SQL, InStr(SQL, vbCrLf) + 2)
                Else
                    SQL_COMMENT = ""
                End If

                SQL_2_EXECUTE = "..."

                If SQL.EndsWith(";") Then
                    Console.Write(SQL)

                    Dim SQLx() As String
                    ReDim SQLx(0)
                    SQLx(0) = SQL

                    If Mid(SQL.ToUpper, 1, 5) <> "BEGIN" Then
                        SQLx = Split(SQL & vbCrLf, ";" & vbCrLf)
                    Else
                        SQLx(0) = Replace(SQLx(0), vbCrLf, " ")
                    End If

                    If SQL_COMMENT.StartsWith("Customer Master") Then
                        Add_XXX("JHC,JHI", "ARTCUST1", "CUST_CODE", SQLx)
                        Add_SQL("Insert into ARTCUST1 Select * from ARTCUST1_COB where CUST_CODE in (Select CUST_CODE from ARTCUST1_COB minus Select CUST_CODE from ARTCUST1)", SQLx)
                    End If
                    If SQL_COMMENT = "Customer Ship-To Addresses" Then
                        Add_XXX("JHC,JHI", "ARTCUST2", "CUST_CODE,CUST_STORE_NO", SQLx)
                        Add_SQL("Insert into ARTCUST2 Select * from ARTCUST2_COB where (CUST_CODE,CUST_STORE_NO) in (Select CUST_CODE,CUST_STORE_NO from ARTCUST2_COB minus Select CUST_CODE,CUST_STORE_NO from ARTCUST2)", SQLx)
                    End If
                    If SQL_COMMENT = "Sales Reps" Then Add_XXX("JHC,JHI", "SOTSREP1", "SREP_CODE", SQLx)
                    If SQL_COMMENT = "Terms Codes" Then Add_XXX("JHC,JHI", "TATTERM1", "TERM_CODE", SQLx)

                    If SQL_COMMENT = "Sales Invoice Header" Then Add_XXX("JHC,JHI", "SOTINVH1", "", SQLx)
                    If SQL_COMMENT = "Sales Invoice Details" Then Add_XXX("JHC,JHI", "SOTINVH2", "", SQLx)
                    If SQL_COMMENT = "Sales CR Memo Header" Then Add_XXX("JHC,JHI", "SOTINVH1", "", SQLx, 0)
                    If SQL_COMMENT = "Sales CR Memo Details" Then Add_XXX("JHC,JHI", "SOTINVH2", "", SQLx, 0)

                    If SQL_COMMENT = "Open Sales Orders (Header)" Then Add_XXX("JHC,JHI", "SOTORDR1", "", SQLx)
                    If SQL_COMMENT = "Open Sales Orders (Detail)" Then Add_XXX("JHC,JHI", "SOTORDR2", "", SQLx)
                    If SQL_COMMENT = "Warehouses" Then Add_XXX("JHC,JHI", "ICTWHSE1", "WHSE_CODE", SQLx)

                    For Each SQL2 As String In SQLx
                        If SQL2 <> "" Then
                            SQL_2_EXECUTE = SQL2

                            Dim rowEDTSQLI2 As DataRow = dst.Tables("EDTSQLI2").NewRow

                            rowEDTSQLI2.Item("SQL_XNO") = SQL_XNO
                            SQL_LNO += 1
                            rowEDTSQLI2.Item("SQL_LNO") = SQL_LNO
                            Dim SQL_BEGIN As Date = Now + ASCMAIN1.NowTSD
                            rowEDTSQLI2.Item("SQL_BEGIN") = SQL_BEGIN
                            rowEDTSQLI2.Item("SQL_TEXT") = Mid(SQL2, 1, 2000)
                            rowEDTSQLI2.Item("SQL_COMMENT") = SQL_COMMENT
                            'If MsgBox(SQL2, MsgBoxStyle.YesNo, "Click Y to Execute, N to Pause") = MsgBoxResult.No Then
                            '    Stop
                            'End If
                            SQL_RESULT = ASCDATA1.ExecuteSQL(SQL2)
                            rowEDTSQLI2.Item("SQL_RESULT") = SQL_RESULT

                            Dim SQL_END As Date = Now + ASCMAIN1.NowTSD
                            rowEDTSQLI2.Item("SQL_END") = SQL_END
                            rowEDTSQLI2.Item("SQL_TIME_SECS") = DateDiff("S", SQL_BEGIN, SQL_END)

                            dst.Tables("EDTSQLI2").Rows.Add(rowEDTSQLI2)
                        End If
                    Next
                End If
            Next

            Update_Record_TDA("EDTSQLI2")

            MsgBox("SQL Script Complete", MsgBoxStyle.OkOnly, "Verification")

        Catch ex As Exception
            MsgBox("SQL_COMMENT = " & SQL_COMMENT & vbCrLf & "SQL = " & SQL_2_EXECUTE & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Please report to ABS - do Not click OK without a Screenshot")

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
            Exit Sub
        End Try

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Add_SQL(ByVal sql As String, ByRef SQLx() As String)
        Dim p As Integer = UBound(SQLx)
        ReDim Preserve SQLx(p + 1)
        SQLx(p + 1) = sql
    End Sub

    Sub Add_XXX( _
    ByVal XXXs As String, _
    ByVal TABLE_NAME As String, _
    ByVal keyCOLUMNs As String, _
    ByRef SQLx() As String, _
    Optional ByVal iSQL As Integer = 1)

        For Each XXX As String In Split(XXXs, ",")
            Dim sqly As String = ""
            Dim p As Integer = UBound(SQLx)
            ReDim Preserve SQLx(p + 3)
            Dim TABLE_NAME_XXX As String = TABLE_NAME & "_" & XXX

            ASCDATA1.ExecuteSQL("DROP TABLE " & TABLE_NAME_XXX, True)
            ASCDATA1.ExecuteSQL("CREATE TABLE " & TABLE_NAME_XXX _
                                & " AS SELECT * FROM " & TABLE_NAME & " WHERE ROWNUM < 1")

            SQLx(p + 1) = "Truncate Table " & TABLE_NAME_XXX
            sqly = SQLx(iSQL)
            sqly = Replace(sqly, TABLE_NAME & " ", TABLE_NAME & "_" & XXX & " ")
            sqly = Replace(sqly, "JHX.JHU_", "JHX." & XXX & "_")
            sqly = Replace(sqly, "'JHU'", "'" & XXX & "'")

            If TABLE_NAME = "ARTCUST2" Then
                If XXX = "JHI" Then
                    sqly = Replace(sqly, "EDI_STORE_ABBR ", "NULL ")
                    sqly = Replace(sqly, "EDI_SHIP_TO ", "NULL ")
                End If
            End If
            If TABLE_NAME = "SOTINVH1" Then
                If XXX = "JHI" And iSQL = 0 Then ' SPECIAL FOR CR MEMO FOR JHI
                    sqly = Replace(sqly, "DEPARTMENT ", "NULL ")
                    sqly = Replace(sqly, "SALES_TYPE ", "'JHI' ")
                    sqly = Replace(sqly, "JHU_CUST_PO_NO_ ", "NULL ")
                    sqly = Replace(sqly, "STORE_CODE", "NULL")
                End If
            End If
            If TABLE_NAME = "SOTINVH2" Then
                If XXX = "JHI" And iSQL = 0 Then ' SPECIAL FOR CR MEMO FOR JHI
                    sqly = Replace(sqly, "L.SO_NO_ ", "NULL ")
                End If
            End If

            SQLx(p + 2) = sqly

            SQLx(p + 3) = "Insert into " & TABLE_NAME & " Select * from " & TABLE_NAME_XXX
            If keyCOLUMNs <> "" Then
                SQLx(p + 3) &= "" _
                & " where (" & keyCOLUMNs & ") in (Select " & keyCOLUMNs & " from " & TABLE_NAME_XXX _
                & " minus Select " & keyCOLUMNs & " from " & TABLE_NAME & ")"
            End If
        Next
    End Sub
End Class