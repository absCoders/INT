Imports ABSolution
Imports Infragistics.Win
Imports System.Windows.Forms
Imports System.IO

Public Class EDFMDBIM

    Dim SO_PARM_LOCAL_DIR_INBOUND As String = String.Empty
    Dim SO_PARM_MDB_PROVIDER As String = String.Empty

    Private Enum FormStates
        LoadDatabases
        LoadMDBTables
        LoadTableData
        UpdateToOracle
    End Enum

    Dim fillTables As List(Of String) = New List(Of String)

    Private currentState As FormStates = FormStates.LoadDatabases

    Private viewMDBTABLES As DataView = New DataView
    Private viewTableData As DataView = New DataView

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim sql As String = String.Empty
        With dst
            Create_TDA(.Tables.Add, "SOTPARMF", "*")

            sql = "Select 0 IMPORT, LPAD(' ', 100) FILENAME, SYSDATE DATE_LAST_MODIFIED FROM DUAL"
            Create_TDA(.Tables.Add, "DATABASES", sql, 0, False, String.Empty, 0, String.Empty)

            sql = "Select 0 IMPORT, LPAD(' ', 100) DBNAME, LPAD(' ', 100) TABLENAME, 0 IN_ORACLE FROM DUAL"
            Create_TDA(.Tables.Add, "MDBTABLES", sql, 0, False, String.Empty, 0, String.Empty)

        End With

        grdDatabase.DataSource = dst.Tables("DATABASES")

        'grdMDBTables.DataSource = dst.Tables("MDBTABLES")
        viewMDBTABLES = New DataView(dst.Tables("MDBTABLES"))
        viewMDBTABLES.RowFilter = "TABLENAME = '**'"
        viewMDBTABLES.Sort = "TABLENAME"
        grdMDBTables.DataSource = viewMDBTABLES

        viewTableData = New DataView
        'viewTableData.RowFilter = "TABLENAME = '**'"
        'viewTableData.Sort = "TABLENAME"
        'grdTableData.DataSource = viewTableData

        currentState = FormStates.LoadDatabases
        Me.Clear_Record()

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load Databases"
                SO_PARM_LOCAL_DIR_INBOUND = String.Empty
                SO_PARM_MDB_PROVIDER = MyBase.Absx1.txtFor("SO_PARM_MDB_PROVIDER").Text
                SO_PARM_MDB_PROVIDER = SO_PARM_MDB_PROVIDER.Trim
                MyBase.Fill_Records("SOTPARMF", SO_PARM_MDB_PROVIDER)
                If dst.Tables("SOTPARMF").Rows.Count = 0 Then
                    SO_PARM_MDB_PROVIDER = String.Empty
                    EMsg = "Missing or Invalid MDB Provider"
                Else
                    SO_PARM_LOCAL_DIR_INBOUND = (dst.Tables("SOTPARMF").Rows(0).Item("SO_PARM_LOCAL_DIR_INBOUND") & String.Empty).ToString.Trim
                    MyBase.Absx1.txtFor("SO_PARM_FTP_HOST_NAME").Text = dst.Tables("SOTPARMF").Rows(0).Item("SO_PARM_FTP_HOST_NAME") & String.Empty
                    If SO_PARM_LOCAL_DIR_INBOUND.Length = 0 Then
                        EMsg = "MDB Provider missing inbound directory."
                    Else
                        If Not My.Computer.FileSystem.DirectoryExists(SO_PARM_LOCAL_DIR_INBOUND) Then
                            EMsg = "MDB Provider inbound directory (" & SO_PARM_LOCAL_DIR_INBOUND & ") cannot be found. "
                        End If
                    End If
                End If

            Case "Load MDB Tables"
                Me.grdDatabase.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
                If dst.Tables("DATABASES").Select("IMPORT <> 0").Length = 0 Then
                    EMsg = "There are no selected Databases to import."
                End If

            Case "Load Table Data"
                Me.grdMDBTables.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
                If dst.Tables("MDBTables").Select("IMPORT <> 0").Length = 0 Then
                    EMsg = "There are no selected Database Tables to import."
                End If

            Case "Update"
                If Me.fillTables.Count = 0 Then
                    EMsg = "There are no tables to Update."
                End If

            Case "Cancel"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load Databases"
                Me.Load_Record()
                Me.Mode_Settings(True)

            Case "Load MDB Tables"
                Me.Import_Tables()
                If Me.grdMDBTables.Rows.Count > 0 Then
                    grdMDBTables.ActiveRow = grdMDBTables.Rows(0)
                End If
                Me.Mode_Settings(True)

            Case "Load Table Data"
                Me.CreateAndFillOrcaleTables()
                If Me.grdMDBTables.Rows.Count > 0 Then
                    grdMDBTables.ActiveRow = grdMDBTables.Rows(0)
                End If
                Me.Mode_Settings(True)

            Case "Update"
                Me.Update_Record()
                currentState = FormStates.LoadDatabases
                Me.Mode_Settings(False)

            Case "Cancel"
                currentState = FormStates.LoadDatabases
                Me.Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                '.Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                '.Groups("Screen Control").Items("Import Data").Settings.Enabled = iScreenMode
                '.Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode

                Select Case currentState
                    Case FormStates.LoadDatabases
                        .Groups("Screen Control").Items("Load Databases").Settings.Enabled = DefaultableBoolean.True
                        .Groups("Screen Control").Items("Load MDB Tables").Settings.Enabled = DefaultableBoolean.False
                        .Groups("Screen Control").Items("Load Table Data").Settings.Enabled = DefaultableBoolean.False
                        .Groups("Screen Control").Items("Update").Settings.Enabled = DefaultableBoolean.False
                        .Groups("Screen Control").Items("Cancel").Settings.Enabled = DefaultableBoolean.False
                        Me.grdDatabase.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                        Me.grdMDBTables.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                        Me.grdTableData.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

                    Case FormStates.LoadMDBTables
                        .Groups("Screen Control").Items("Load Databases").Settings.Enabled = DefaultableBoolean.False
                        .Groups("Screen Control").Items("Load MDB Tables").Settings.Enabled = DefaultableBoolean.True
                        .Groups("Screen Control").Items("Load Table Data").Settings.Enabled = DefaultableBoolean.False
                        .Groups("Screen Control").Items("Update").Settings.Enabled = DefaultableBoolean.False
                        .Groups("Screen Control").Items("Cancel").Settings.Enabled = DefaultableBoolean.True
                        Me.grdDatabase.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                        Me.grdMDBTables.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                        Me.grdTableData.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

                    Case FormStates.LoadTableData
                        .Groups("Screen Control").Items("Load Databases").Settings.Enabled = DefaultableBoolean.False
                        .Groups("Screen Control").Items("Load MDB Tables").Settings.Enabled = DefaultableBoolean.False
                        .Groups("Screen Control").Items("Load Table Data").Settings.Enabled = DefaultableBoolean.True
                        .Groups("Screen Control").Items("Update").Settings.Enabled = DefaultableBoolean.False
                        .Groups("Screen Control").Items("Cancel").Settings.Enabled = DefaultableBoolean.True
                        Me.grdDatabase.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                        Me.grdMDBTables.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                        Me.grdTableData.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

                    Case FormStates.UpdateToOracle
                        .Groups("Screen Control").Items("Load Databases").Settings.Enabled = DefaultableBoolean.False
                        .Groups("Screen Control").Items("Load MDB Tables").Settings.Enabled = DefaultableBoolean.False
                        .Groups("Screen Control").Items("Load Table Data").Settings.Enabled = DefaultableBoolean.False
                        .Groups("Screen Control").Items("Update").Settings.Enabled = DefaultableBoolean.True
                        .Groups("Screen Control").Items("Cancel").Settings.Enabled = DefaultableBoolean.True
                        Me.grdDatabase.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                        Me.grdMDBTables.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                        Me.grdTableData.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                End Select


            End With
        End If


        Me.SplitContainer2.Visible = tf
        MyBase.Absx1.txtFor("SO_PARM_MDB_PROVIDER").ReadOnly = tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("DATABASES").Clear()
        dst.Tables("MDBTABLES").Clear()
        dst.EnforceConstraints = True

        MyBase.Absx1.txtFor("SO_PARM_FTP_HOST_NAME").Clear()
        MyBase.Absx1.txtFor("SO_PARM_MDB_PROVIDER").Clear()
        MyBase.Absx1.txtFor("SO_PARM_MDB_PROVIDER").Focus()
        fillTables.Clear()

    End Sub

    Sub Load_Record()

        Call ASCMAIN1.Progress("Loading Database Files")

        Dim rowDATABASES As DataRow = Nothing
        Dim fld As New System.IO.DirectoryInfo(SO_PARM_LOCAL_DIR_INBOUND)

        For Each mdbFile As System.IO.FileInfo In fld.GetFiles("*.mdb")
            rowDATABASES = dst.Tables("DATABASES").NewRow
            rowDATABASES.Item("IMPORT") = False
            rowDATABASES.Item("FILENAME") = mdbFile.Name
            rowDATABASES.Item("DATE_LAST_MODIFIED") = mdbFile.LastWriteTime
            dst.Tables("DATABASES").Rows.Add(rowDATABASES)
        Next

        If dst.Tables("DATABASES").Rows.Count > 0 Then
            Me.grdDatabase.Text = "MDBs located in " & SO_PARM_LOCAL_DIR_INBOUND
        Else
            Me.grdDatabase.Text = SO_PARM_LOCAL_DIR_INBOUND & " contains no MDBs."
        End If

        Call ASCMAIN1.Progress("")
        currentState = FormStates.LoadMDBTables
    End Sub

    Sub Update_Record()

        MyBase.BeginTrans()

        Try
            For Each tablename As String In Me.fillTables
                MyBase.Update_Record_TDA(tablename)
            Next

        Catch ex As Exception
            MyBase.Rollback(ex.Message)
            Exit Sub
        End Try

        MyBase.CommitTrans("Tables Added Successfully")

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

    Private Sub Import_Tables()

        Dim mdbAccess As New EDCPRMDB
        Dim dbLocation As String = String.Empty


        If Not SO_PARM_LOCAL_DIR_INBOUND.EndsWith("\") Then
            SO_PARM_LOCAL_DIR_INBOUND &= "\"
        End If

        Dim rowMDBTABLES As DataRow = Nothing

        For Each rowDATABASE As DataRow In dst.Tables("DATABASES").Select("IMPORT <> 0")
            mdbAccess = New EDCPRMDB
            dbLocation = SO_PARM_LOCAL_DIR_INBOUND & rowDATABASE.Item("FILENAME")
            mdbAccess.CreateConnectionSting(4.0, dbLocation, String.Empty, String.Empty)

            Dim tableList As List(Of String) = mdbAccess.GetListOfMdbTables
            For Each tableName As String In tableList
                rowMDBTABLES = dst.Tables("MDBTABLES").NewRow
                rowMDBTABLES.Item("IMPORT") = False
                rowMDBTABLES.Item("DBNAME") = rowDATABASE.Item("FILENAME")
                rowMDBTABLES.Item("TABLENAME") = tableName

                If ASCDATA1.GetDataRow("SELECT * FROM USER_TABLES WHERE TABLE_NAME = '" & tableName & "'") Is Nothing Then
                    rowMDBTABLES.Item("IN_ORACLE") = False
                Else
                    rowMDBTABLES.Item("IN_ORACLE") = True
                End If

                dst.Tables("MDBTABLES").Rows.Add(rowMDBTABLES)
            Next

        Next
        currentState = FormStates.LoadTableData

    End Sub

    Private Sub CreateAndFillOrcaleTables()

        fillTables.Clear()
        Dim databaseName As String = String.Empty
        Dim dbLocation As String = String.Empty

        For Each rowDATABASES As DataRow In dst.Tables("DATABASES").Select("IMPORT <> 0")

            databaseName = rowDATABASES.Item("FILENAME")

            For Each rowMDBTABLES As DataRow In dst.Tables("MDBTABLES").Select("IMPORT <> 0 AND DBNAME = '" & databaseName & "'")
                Dim tableName As String = rowMDBTABLES.Item("TABLENAME")

                Dim clsEDCPRMDB As EDCPRMDB = New EDCPRMDB
                If Not SO_PARM_LOCAL_DIR_INBOUND.EndsWith("\") Then
                    SO_PARM_LOCAL_DIR_INBOUND = SO_PARM_LOCAL_DIR_INBOUND & "\"
                End If

                dbLocation = SO_PARM_LOCAL_DIR_INBOUND & rowDATABASES.Item("FILENAME")
                clsEDCPRMDB.CreateConnectionSting(4.0, dbLocation, String.Empty, String.Empty)

                Dim tbl As DataTable = New DataTable
                If rowMDBTABLES.Item("IN_ORACLE") = 1 Then
                    tbl = ASCDATA1.GetDataTable("SELECT * FROM " & tableName & " WHERE ROWNUM < 1")
                Else
                    tbl = clsEDCPRMDB.CreateOracleTableFromMdbSchema(tableName, tableName)
                End If

                If tbl Is Nothing Then
                    If clsEDCPRMDB.Errors.Count > 0 Then
                        Dim eMsg As String = String.Empty
                        For Each msg As String In clsEDCPRMDB.Errors
                            eMsg &= msg & Environment.NewLine
                        Next
                        MessageBox.Show(eMsg, "Create Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                    'Stop
                    Exit Sub
                Else
                    Try
                        dst.Tables.Remove(tableName)
                    Catch ex As Exception
                        ' Nothing
                    End Try

                    MyBase.Create_TDA(dst.Tables.Add, tableName, "*")
                    tbl = clsEDCPRMDB.CopyMdbIntoOracle(tableName, tableName, False)
                    If tbl Is Nothing Then
                        Stop
                    Else
                        dst.Tables(tbl.TableName).Merge(tbl)
                        fillTables.Add(tableName)
                    End If
                End If
            Next
        Next

        Me.grdMDBTables.ActiveRow = Me.grdMDBTables.Rows(0)
        Dim vSender As Object = Nothing
        Dim e As System.EventArgs = Nothing
        Me.grdMDBTables_AfterRowActivate(vSender, e)


        currentState = FormStates.UpdateToOracle
    End Sub

    Private Sub grdDatabase_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdDatabase.AfterRowActivate

        Try
            Dim DBNAME As String = grdDatabase.ActiveRow.Cells("FILENAME").Value & String.Empty
            viewMDBTABLES.RowFilter = "DBNAME = '" & DBNAME & "'"
            If grdDatabase.Rows.Count > 0 Then
                grdMDBTables.Text = "Tables in Database : " & DBNAME
            Else
                grdMDBTables.Text = String.Empty
            End If
        Catch ex As Exception
            grdMDBTables.Text = String.Empty
        End Try

    End Sub

    Private Sub grdMDBTables_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdMDBTables.AfterRowActivate

        Dim tableName As String = grdMDBTables.ActiveRow.Cells("TABLENAME").Value & String.Empty
        viewTableData = New DataView

        Try
            viewTableData = New DataView(dst.Tables(tableName))
            grdTableData.DataSource = Nothing
            ''(dst.Tables(tableName))
            While grdTableData.DisplayLayout.Bands(0).Columns.Count > 0
                grdTableData.DisplayLayout.Bands(0).Columns.Remove(0)
            End While

            For Each col As DataColumn In viewTableData.Table.Columns
                grdTableData.DisplayLayout.Bands(0).Columns.Add(col.ColumnName, col.ColumnName)
                grdTableData.DisplayLayout.Bands(0).Columns(col.ColumnName).Hidden = False
            Next

            grdTableData.DataSource = Me.viewTableData
            grdTableData.Text = "Data in table : " & tableName
        Catch ex As Exception
            grdTableData.Text = String.Empty
        End Try

    End Sub

End Class