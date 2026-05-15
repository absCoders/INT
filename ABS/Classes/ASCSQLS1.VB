Imports System.Data.SqlClient
Imports System.Data.Common

Public Class ASCSQLS1

    Public sqlServerConnection As New SqlConnection
    Public sqlServerCommand As New SqlCommand
    Public sqlServerSql As String = String.Empty
    Public sqlServerDst As New DataSet()
    Private sqlServerTransaction As SqlTransaction

    Public TDAs As New Dictionary(Of String, SqlDataAdapter)
    Public TBLs As New Dictionary(Of String, DataTable)
    Public pROWs As New Dictionary(Of String, DataRow)
    Public ROWs As New Dictionary(Of String, DataRow)
    Public CMDs As New Dictionary(Of String, SqlCommand)
    Private temps As New List(Of Integer)

    Public DataTableSchema As DataTable
    Public DTS As New SortedList(Of String, DataTable)

    Public Class DataReaderAdapter
        Inherits DbDataAdapter

        Public Function FillFromReader(ByVal dataTable As DataTable, ByVal dataReader As IDataReader) As Integer
            Return Me.Fill(dataTable, dataReader)
        End Function

        Protected Overloads Overrides Function CreateRowUpdatedEvent(ByVal dataRow As DataRow, ByVal command As IDbCommand, ByVal statementType As StatementType, ByVal tableMapping As DataTableMapping) As RowUpdatedEventArgs
            Return Nothing
        End Function

        Protected Overloads Overrides Function CreateRowUpdatingEvent(ByVal dataRow As DataRow, ByVal command As IDbCommand, ByVal statementType As StatementType, ByVal tableMapping As DataTableMapping) As RowUpdatingEventArgs
            Return Nothing
        End Function

        Protected Overloads Overrides Sub OnRowUpdated(ByVal value As RowUpdatedEventArgs)
        End Sub

        Protected Overloads Overrides Sub OnRowUpdating(ByVal value As RowUpdatingEventArgs)
        End Sub
    End Class

    Public Class SQLCreateTable

        Private _connection As SqlConnection
        Private _transaction As SqlTransaction
        Private _tableName As String
        Private _dropTableIfExist As Boolean = False

        Public Property Connection() As SqlConnection
            Get
                Return _connection
            End Get
            Set(ByVal value As SqlConnection)
                _connection = value
            End Set
        End Property

        Public Property Transaction() As SqlTransaction
            Get
                Return _transaction
            End Get
            Set(ByVal value As SqlTransaction)
                _transaction = value
            End Set
        End Property

        Public Property DestinationTableName() As String
            Get
                Return _tableName
            End Get
            Set(ByVal value As String)
                _tableName = value
            End Set
        End Property

        Public Property DropTableIfExists() As Boolean
            Get
                Return _dropTableIfExist
            End Get
            Set(ByVal value As Boolean)
                _dropTableIfExist = value
            End Set
        End Property

        Public Sub New()
        End Sub

        Public Sub New(ByVal connection As SqlConnection)
            MyClass.New(connection, Nothing)
        End Sub

        Public Sub New(ByVal connection As SqlConnection, ByVal transaction As SqlTransaction)
            _connection = connection
            _transaction = transaction
        End Sub

        Public Function DropTable(ByVal tablename As String) As Object
            Dim sql As String = "IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & tablename & "]') AND type in (N'U')) DROP Table [dbo].[" & tablename & "]"
            Dim cmd As SqlCommand
            If Not (_transaction Is Nothing) AndAlso Not (_transaction.Connection Is Nothing) Then
                cmd = New SqlCommand(sql, _connection, _transaction)
            Else
                cmd = New SqlCommand(sql, _connection)
            End If
            Console.WriteLine("DropTable: " & sql)
            Return cmd.ExecuteNonQuery
        End Function

        Public Function Create(ByVal schema As DataTable) As Object
            Return Create(schema, 0)
        End Function

        Public Function Create(ByVal schema As DataTable, ByVal numKeys As Integer) As Object
            Dim primaryKeys(numKeys) As Integer
            Dim i As Integer = 0
            While i < numKeys
                primaryKeys(i) = i
                System.Math.Min(System.Threading.Interlocked.Increment(i), i - 1)
            End While
            Return Create(schema, primaryKeys)
        End Function

        Public Function Create(ByVal schema As DataTable, ByVal primaryKeys As Integer()) As Object

            If DropTableIfExists Then
                DropTable(DestinationTableName)
            End If
            Dim sql As String = GetCreateSQL(_tableName, schema, primaryKeys)
            Dim cmd As SqlCommand
            If Not (_transaction Is Nothing) AndAlso Not (_transaction.Connection Is Nothing) Then
                cmd = New SqlCommand(sql, _connection, _transaction)
            Else
                cmd = New SqlCommand(sql, _connection)
            End If
            Return cmd.ExecuteNonQuery
        End Function

        Public Function CreateFromDataTable(ByVal table As DataTable, ByVal dtSchema As DataTable) As Object

            If DropTableIfExists Then
                DropTable(DestinationTableName)
            End If

            Dim sql As String = GetCreateFromDataTableSQL(_tableName, table, dtSchema)
            Dim cmd As SqlCommand
            If Not (_transaction Is Nothing) AndAlso Not (_transaction.Connection Is Nothing) Then
                cmd = New SqlCommand(sql, _connection, _transaction)
            Else
                cmd = New SqlCommand(sql, _connection)
            End If
            Return cmd.ExecuteNonQuery
        End Function


        Public Shared Function GetCreateSQL(ByVal tableName As String, ByVal schema As DataTable, ByVal primaryKeys As Integer()) As String
            Dim sql As String = "IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & tableName & "]') AND type in (N'U')) CREATE TABLE " & tableName & " (" & Microsoft.VisualBasic.Chr(10) & ""
            For Each column As DataRow In schema.Rows
                If Not (schema.Columns.Contains("IsHidden") AndAlso CType(column("IsHidden"), Boolean)) Then
                    sql += column("ColumnName").ToString & " " & SQLGetType(column) & "," & Microsoft.VisualBasic.Chr(10) & ""
                End If
            Next
            sql = sql.TrimEnd(New Char() {","c, Microsoft.VisualBasic.Chr(10)}) & "" & Microsoft.VisualBasic.Chr(10) & ""
            Dim pk As String = "CONSTRAINT PK_" & tableName & " PRIMARY KEY CLUSTERED ("
            Dim hasKeys As Boolean = (Not (primaryKeys Is Nothing) AndAlso primaryKeys.Length > 0)
            If hasKeys Then
                For Each key As Integer In primaryKeys
                    pk += schema.Rows(key)("ColumnName").ToString & ", "
                Next
            Else
                Dim keys As String = String.Join(", ", GetPrimaryKeys(schema))
                pk += keys
                hasKeys = keys.Length > 0
            End If
            pk = pk.TrimEnd(New Char() {","c, " "c, Microsoft.VisualBasic.Chr(10)}) & ")" & Microsoft.VisualBasic.Chr(10) & ""
            If hasKeys Then
                sql += pk
            End If
            sql += ")"
            Console.WriteLine("Primary Keys: " & primaryKeys.Length & " " & "GetCreateSQL: " & sql)
            Return sql
        End Function

        Public Shared Function GetCreateFromDataTableSQL(ByVal tableName As String, ByVal table As DataTable) As String

            Dim sql As String = "IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & tableName & "]') AND type in (N'U')) CREATE TABLE [" & tableName & "] (" & Microsoft.VisualBasic.Chr(10) & ""

            For Each column As DataColumn In table.Columns
                sql += "[" & column.ColumnName & "] " & SQLGetType(column) & "," & Microsoft.VisualBasic.Chr(10) & ""
            Next


            sql = sql.TrimEnd(New Char() {","c, Microsoft.VisualBasic.Chr(10)}) & "" & Microsoft.VisualBasic.Chr(10) & ""

            If table.PrimaryKey.Length > 0 Then
                sql += "CONSTRAINT [PK_" & tableName & "] PRIMARY KEY CLUSTERED ("
                For Each column As DataColumn In table.PrimaryKey
                    sql += "[" & column.ColumnName & "],"
                Next
                sql = sql.TrimEnd(New Char() {","c}) & "))" & Microsoft.VisualBasic.Chr(10)

            End If

            sql += ")"
            Console.WriteLine("Primary Keys: " & table.PrimaryKey.Length & " " & "GetCreateFromDataTableSQL: " & sql)
            Return sql
        End Function

        Public Shared Function GetCreateFromDataTableSQL(ByVal tableName As String, ByVal table As DataTable, ByVal schema As DataTable) As String

            Dim sql As String = "IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & tableName & "]') AND type in (N'U')) CREATE TABLE [" & tableName & "] (" & Microsoft.VisualBasic.Chr(10) & ""

            For Each column As DataColumn In table.Columns
                Dim rowSchema As DataRow = Nothing
                Dim DATA_LENGTH As Integer = -1
                Dim DATA_PRECISION As Integer = 0
                Dim DATA_SCALE As Integer = 0

                Try
                    rowSchema = schema.Select("COLUMN_NAME = '" & column.ColumnName & "'")(0)
                    DATA_LENGTH = Val(rowSchema.Item("DATA_LENGTH") & String.Empty)
                    DATA_PRECISION = Val(rowSchema.Item("DATA_PRECISION") & String.Empty)
                    DATA_SCALE = Val(rowSchema.Item("DATA_SCALE") & String.Empty)
                Catch ex As Exception

                End Try

                sql += "[" & column.ColumnName & "] " & SQLGetType(column.DataType, DATA_LENGTH, DATA_PRECISION, DATA_SCALE) & "," & Microsoft.VisualBasic.Chr(10) & ""
            Next

            sql = sql.TrimEnd(New Char() {","c, Microsoft.VisualBasic.Chr(10)}) & "" & Microsoft.VisualBasic.Chr(10) & ""

            Dim keyfield As Boolean = False
            For Each row As DataRow In schema.Select(String.Empty, "COLUMN_ID")
                If row.Item("NULLABLE") & String.Empty <> "N" Then Exit For
                If keyfield = False Then
                    keyfield = True
                    sql += "CONSTRAINT [PK_" & tableName & "] PRIMARY KEY CLUSTERED ("
                End If
                sql += "[" & row.Item("Column_Name") & "],"
            Next
            If keyfield = True Then
                sql = sql.TrimEnd(New Char() {","c}) & ")" & Microsoft.VisualBasic.Chr(10)
            End If

            sql += ")"
            Return sql
        End Function


        Public Shared Function GetPrimaryKeys(ByVal schema As DataTable) As String()
            Dim Keys As New List(Of String)

            For Each column As DataRow In schema.Rows
                If schema.Columns.Contains("IsKey") AndAlso CType(column("IsKey"), Boolean) Then
                    Keys.Add(column("ColumnName").ToString)
                End If
            Next
            Return Keys.ToArray
        End Function

        Public Shared Function SQLGetType(ByVal type As Object, ByVal columnSize As Integer, ByVal numericPrecision As Integer, ByVal numericScale As Integer) As String

            Select Case type.ToString
                Case "System.String"
                    Return "VARCHAR(" & (Microsoft.VisualBasic.IIf((columnSize = -1), 255, columnSize)) & ")"

                Case "System.Decimal"
                    If numericScale > 0 Then
                        Return "REAL"
                    Else
                        If numericPrecision > 10 Then
                            Return "BIGINT"
                        Else
                            Return "INT"
                        End If
                    End If

                Case "System.Double", "System.Single"
                    Return "REAL"

                Case "System.Int64"
                    Return "BIGINT"

                Case "System.Int16", "System.Int32"
                    Return "INT"

                Case "System.DateTime"
                    Return "DATETIME"

                Case Else
                    Throw New Exception(type.ToString & " not implemented.")
            End Select
        End Function

        Public Shared Function SQLGetType(ByVal schemaRow As DataRow) As String
            Return SQLGetType(schemaRow("DataType"), Integer.Parse(schemaRow("ColumnSize").ToString), Integer.Parse(schemaRow("NumericPrecision").ToString), Integer.Parse(schemaRow("NumericScale").ToString))
        End Function

        Public Shared Function SQLGetType(ByVal column As DataColumn) As String
            Return SQLGetType(column.DataType, column.MaxLength, 10, 2)
        End Function

    End Class

    ''' <summary>
    ''' Begin a SQL Server Transaction
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub sqlServerBeginTrans()
        sqlServerTransaction = sqlServerConnection.BeginTransaction
        sqlServerCommand = sqlServerConnection.CreateCommand()
        sqlServerCommand.Transaction = sqlServerTransaction
    End Sub

    ''' <summary>
    ''' Commit a SQL Server Trans Action
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub sqlServerCommitTrans()
        sqlServerTransaction.Commit()
        sqlServerTransaction = Nothing
    End Sub

    ''' <summary>
    ''' Rollback a SQL Server Transaction
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub sqlServerRollback()
        sqlServerTransaction.Rollback()
    End Sub

    ''' <summary>
    ''' Creates a DataAdapter which will be used to Fill and optionally to Update data.
    ''' </summary>
    ''' <param name="tbl">The DataTable used to contain the Data on the Client.</param>
    ''' <param name="TABLE_NAME">A logical name for the DataTable.  This should be the TABLE_NAME in the Database for Adapters which are used for Updates.</param>
    ''' <param name="sql_custom">The SQL statement used to generate the Result Set for the Select Statement.  If "*", then all columns are used.  If "**" then ASCMAIN1.sql is used for the result set, but the Insert and Update statements will be prepared using * from the TABLE_NAME</param>
    ''' <param name="NumberOfKeysUsedToSelect"></param>
    ''' <param name="for_update">Set this to True to cause the Insert/Update/Delete statements to be generated for the Adapter.  A Valid (Database) TABLE_NAME must be specified when this field is set (or defaulted) to True</param>
    ''' <param name="custom_parameters">A string indicating the number and types of parameters used in the where clause, such as "VVIND" to indicate 5 parameters, 2 Varchar2, 1 Integer, 1 Number, and 1 Date.  Always use :PARMx as your parameter names in the where clause.</param>
    ''' <param name="Key_Field_Count">The number of leading columns that are to be used to create a key for the DataTable generated by the Fill.  This field is only necessary when specifying a Read-Only Result Set.</param>
    ''' <remarks></remarks>
    Public Sub sqlSvrCreate_TDA( _
    ByRef tbl As DataTable, _
    ByVal TABLE_NAME As String, _
    ByVal sql_custom As String, _
    Optional ByVal NumberOfKeysUsedToSelect As Integer = -1, _
    Optional ByVal for_update As Boolean = True, _
    Optional ByVal custom_parameters As String = "", _
    Optional ByVal Key_Field_Count As Integer = -1, _
    Optional ByVal Update_COLUMN_NAMEs As String = "")

        Dim tblTABLE_NAME As String = TABLE_NAME
        If Not tbl.TableName Like "Table*" Then
            tblTABLE_NAME = tbl.TableName
        End If

        TDAs.Add(tblTABLE_NAME, sqlSvrGetDataAdapter(tbl, TABLE_NAME, sql_custom, for_update, Key_Field_Count, False, NumberOfKeysUsedToSelect, Update_COLUMN_NAMEs, custom_parameters, tblTABLE_NAME))
        'If custom_parameters.Length <> 0 Then '  <> "" Then
        '    Dim ptbl As New DataTable
        '    Call ascsqls1.Create_Parameters(TDAs(tblTABLE_NAME).SelectCommand, custom_parameters, ptbl)
        '    pROWs.Add(tblTABLE_NAME, ptbl.NewRow)
        'End If
        'Dim tbl_schema As DataTable
        'Using rdr = TDAs(tblTABLE_NAME).SelectCommand.ExecuteReader

        '    tbl_schema = rdr.GetSchemaTable
        '    tbl_schema.PrimaryKey = New DataColumn() {tbl_schema.Columns("ColumnName")}

        '    TBL_SCHEMAs.Add(tblTABLE_NAME, tbl_schema)
        'End Using
        For Each dc As DataColumn In tbl.Columns
            'Dim row_Schema As DataRow = tbl_Schema.Rows.Find(dc.ColumnName)
            'If Not row_Schema("NumericPrecision").Equals(DBNull.Value) Then
            '    If Val(row_Schema("NumericScale")) = 0 Then
            '        'dc.DataType = GetType(System.Int64)
            '    Else
            '        'dc.DataType = GetType(System.Decimal)
            '    End If
            'End If
            dc.ReadOnly = False
        Next

        TBLs.Add(tblTABLE_NAME, tbl)
        Dim dvw As New DataView(tbl)
        'DVWs.Add(tblTABLE_NAME, dvw)
    End Sub

    ''' <summary>
    ''' Performs the Update Command for the Table Data Adapter.  Use sql_Delete only if you want to delete rows (using the supplied sql_Delete statement) and then have the current rows Inserted.
    ''' </summary>
    ''' <param name="TABLE_NAME"></param>
    ''' <param name="sql_Delete">Specify a complete Delete statement, or else just the where clause.  If sql_delete does not begin with the word 'Delete', then the clause 'Delete from {TABLE_NAME} where ' will be pre-pended to the clause supplied.</param>
    ''' <remarks></remarks>
    Public Sub sqlSvrUpdate_Record_TDA(ByVal TABLE_NAME As String, Optional ByVal sql_Delete As String = "")

        ' Without the next 2 lines, data in text boxes in single row datatables (like header tables) will not get written to Oracle
        'Dim X As CurrencyManager = Me.BindingContext(dst.Tables(TABLE_NAME))
        'X.EndCurrentEdit()

        'If AUDIT.ContainsKey(TABLE_NAME) Then
        '    WriteAuditTrail(TABLE_NAME)
        'End If

        If sql_Delete <> "" Then
            If Mid(Trim(sql_Delete), 1, 6).ToUpper <> "DELETE" Then
                sql_Delete = "Delete from " & TABLE_NAME & " where " & sql_Delete
            End If
            sqlSvrExecuteSQL(sql_Delete)
            TBLs(TABLE_NAME).AcceptChanges()
            For Each row As DataRow In TBLs(TABLE_NAME).Rows
                row.SetAdded()
            Next
        Else
        End If

        If sqlServerTransaction IsNot Nothing Then
            If TDAs(TABLE_NAME).UpdateCommand IsNot Nothing Then
                TDAs(TABLE_NAME).UpdateCommand.Transaction = sqlServerTransaction
            End If
            If TDAs(TABLE_NAME).InsertCommand IsNot Nothing Then
                TDAs(TABLE_NAME).InsertCommand.Transaction = sqlServerTransaction
            End If
            If TDAs(TABLE_NAME).DeleteCommand IsNot Nothing Then
                TDAs(TABLE_NAME).DeleteCommand.Transaction = sqlServerTransaction
            End If
        Else
            If TDAs(TABLE_NAME).UpdateCommand IsNot Nothing Then
                TDAs(TABLE_NAME).UpdateCommand.Transaction = Nothing
            End If
            If TDAs(TABLE_NAME).InsertCommand IsNot Nothing Then
                TDAs(TABLE_NAME).InsertCommand.Transaction = Nothing
            End If
            If TDAs(TABLE_NAME).DeleteCommand IsNot Nothing Then
                TDAs(TABLE_NAME).DeleteCommand.Transaction = Nothing
            End If
        End If

        TDAs(TABLE_NAME).Update(sqlServerDst.Tables(TABLE_NAME))
    End Sub

    Public Function sqlSvrFill_Records( _
          ByVal TABLE_NAME As String, _
          ByVal KEY_VALUE As String, _
          Optional ByVal ClearBeforeFilling As Boolean = True, _
          Optional ByVal Temp_Select As String = "", _
          Optional ByVal tblSubstitute As DataTable = Nothing) As Integer

        If KEY_VALUE = "" Then
            Return sqlSvrFill_Records(TABLE_NAME, , ClearBeforeFilling, Temp_Select, tblSubstitute)
        Else
            Return sqlSvrFill_Records(TABLE_NAME, New String() {KEY_VALUE}, ClearBeforeFilling, Temp_Select, tblSubstitute)
        End If

    End Function

    Public Function sqlSvrFill_Records( _
    ByVal TABLE_NAME As String, _
    Optional ByVal Parameters() As Object = Nothing, _
    Optional ByVal ClearBeforeFilling As Boolean = True, _
    Optional ByVal Temp_Select As String = "", _
    Optional ByVal tblSubstitute As DataTable = Nothing) As Integer

        Dim tbl As DataTable
        If tblSubstitute Is Nothing Then
            tbl = TBLs(TABLE_NAME)
        Else
            tbl = tblSubstitute
        End If

        If Not TDAs.ContainsKey(TABLE_NAME) Then
            Stop ' SHOULD RAISE AN ERROR
        Else

            If ClearBeforeFilling Then
                'dst.Tables(TABLE_NAME).Rows.Clear()
                'TBLs(TABLE_NAME).Rows.Clear()
                tbl.Rows.Clear()
            End If

            Dim UseRow As Boolean = False ' we need to know how to autodetect when to use this

            If UseRow Then
                For i As Integer = 0 To pROWs(TABLE_NAME).ItemArray.Length - 1
                    TDAs(TABLE_NAME).SelectCommand.Parameters(i).Value = pROWs(TABLE_NAME).Item(i)
                Next
            Else
                With TDAs(TABLE_NAME).SelectCommand
                    If Parameters IsNot Nothing Then
                        For i As Integer = 0 To UBound(Parameters)
                            .Parameters(i).Value = Parameters(i)
                        Next
                    Else
                        For i As Integer = 0 To .Parameters.Count - 1
                            .Parameters(i).Value = System.DBNull.Value
                        Next
                    End If
                End With
            End If

            Dim SelectCommandSQL As String = ""
            If Temp_Select <> "" Then
                SelectCommandSQL = TDAs(TABLE_NAME).SelectCommand.CommandText
                TDAs(TABLE_NAME).SelectCommand.CommandText = Temp_Select
            End If

            'Fill_Records = TDAs(TABLE_NAME).Fill(TBLs(TABLE_NAME))

            Dim XDT As Date = Now

            If sqlServerTransaction IsNot Nothing Then
                TDAs(TABLE_NAME).SelectCommand.Transaction = sqlServerTransaction
            End If
            sqlSvrFill_Records = TDAs(TABLE_NAME).Fill(tbl)

            If Temp_Select <> "" Then
                TDAs(TABLE_NAME).SelectCommand.CommandText = SelectCommandSQL
            End If
        End If
    End Function

    Public Function sqlSvrLookUp( _
    ByVal TABLE_NAME As String, _
    ByVal KEY() As String, _
    Optional ByVal Return_Empty_Row_if_Missing As Boolean = False) As DataRow

        Dim row As DataRow = Nothing

        If Not ROWs.ContainsKey(TABLE_NAME) Then
            Call sqlSvrCreate_Lookup(TABLE_NAME)
        End If

        ROWs(TABLE_NAME).Table.Rows.Clear()

        With CMDs(TABLE_NAME)
            For i As Integer = 0 To UBound(KEY)
                .Parameters(i).Value = KEY(i)
            Next

            Using rdr As SqlDataReader = .ExecuteReader
                If rdr.HasRows Then
                    rdr.Read()
                    row = ROWs(TABLE_NAME).Table.NewRow
                    'row = TBLs(TABLE_NAME).NewRow
                    For i As Integer = 0 To rdr.FieldCount - 1
                        row.Item(i) = rdr.Item(i)
                    Next
                Else
                    If Return_Empty_Row_if_Missing Then
                        'Stop
                        row = ROWs(TABLE_NAME).Table.NewRow
                        If ROWs(TABLE_NAME).Table.PrimaryKey.Length = KEY.Length Then
                            For i As Integer = 0 To UBound(KEY)
                                row.Item(i) = KEY(i)
                            Next
                        End If
                    End If
                End If
            End Using
        End With

        If row IsNot Nothing Then
            ROWs(TABLE_NAME) = row
        End If
        Return row
    End Function

    Public Function sqlSvrLookUp( _
    ByVal TABLE_NAME As String, _
    ByVal KEY As String, _
    Optional ByVal Return_Empty_Row_if_Missing As Boolean = False) As DataRow
        Return sqlSvrLookUp(TABLE_NAME, New String() {KEY}, Return_Empty_Row_if_Missing)
    End Function

    Public Sub sqlSvrCreate_Lookup( _
    ByVal LOOKUP_NAME As String, _
    Optional ByVal column_list As String = "*", _
    Optional ByVal where_clause As String = "", _
    Optional ByVal custom_parameters As String = "", _
    Optional ByVal create_parameters_for_key As Boolean = True)

        ' Examples of usage:
        'Create_Lookup("GLTACCT1")
        'Create_Lookup("POTORDR1.OPEN", , "PO_STATUS_CODE = 'O' and PO_ORDER_TYPE = 'S'")
        'Create_Lookup("ICTPRCAT", , "PROD_CODE in (Select Distinct PROD_CODE from ICTPROD1 where VEND_CODE = :PARM1)", "V")
        'Create_Lookup("ICTITEM1.ITEM_UPC_CODE", , "ITEM_UPC_CODE = :PARM1", "V", False)

        Dim TABLE_NAME As String = LOOKUP_NAME
        If InStr(TABLE_NAME, ".") <> 0 Then
            TABLE_NAME = Split(TABLE_NAME, ".")(0)
        End If

        Dim tbl As DataTable

        If column_list = "" Or column_list = "*" Then
            tbl = sqlSvrGetDataTable("*", TABLE_NAME, -1, False)
        Else
            sqlServerSql = "Select " & column_list & " from " & TABLE_NAME
            tbl = sqlSvrGetDataTable("", TABLE_NAME, -1, False)
        End If
        'tbl.TableName = LOOKUP_NAME ?

        Dim cmd As SqlCommand = sqlServerConnection.CreateCommand
        Dim sql As String = ""
        If create_parameters_for_key Then
            For Each dc As DataColumn In tbl.PrimaryKey
                Dim COLUMN_NAME As String = dc.ColumnName
                sql = sql & " and (" & COLUMN_NAME & " = @S_" & COLUMN_NAME & ")"
                Select Case dc.DataType.Name
                    Case "String", "Int16", "Int32"
                        cmd.Parameters.Add("S_" & COLUMN_NAME, SqlDbType.VarChar)
                    Case Else
                        Stop
                End Select
            Next
        End If

        If custom_parameters <> "" Then Call sqlSvrCreate_Parameters(cmd, custom_parameters)

        If where_clause <> "" Then
            where_clause = " and " & where_clause
        End If
        cmd.CommandText = "Select * from " & TABLE_NAME & sqlSvrSQL_Add_WHERE(sql & where_clause)

        If CMDs.ContainsKey(LOOKUP_NAME) Then
            CMDs.Remove(LOOKUP_NAME)
        End If
        CMDs.Add(LOOKUP_NAME, cmd)
        If ROWs.ContainsKey(LOOKUP_NAME) Then
            ROWs.Remove(LOOKUP_NAME)
        End If
        ROWs.Add(LOOKUP_NAME, tbl.NewRow)
    End Sub

    Public Function sqlSvrSQL_Add_WHERE(ByVal WHEREless_Clause As String) As String
        Dim z As String
        z = LTrim$(WHEREless_Clause)
        If InStr(UCase$(z), "WHERE") = 1 Then
            sqlSvrSQL_Add_WHERE = z
        ElseIf InStr(UCase$(z), "AND") = 1 Then
            sqlSvrSQL_Add_WHERE = " where " & Mid$(z, 4)
        ElseIf z <> "" Then
            sqlSvrSQL_Add_WHERE = " where " & z
        Else
            sqlSvrSQL_Add_WHERE = WHEREless_Clause
        End If
    End Function

    Public Sub sqlSvrResetState()
        sqlServerDst.Reset()
        TDAs.Clear()
        TBLs.Clear()
        'TBL_SCHEMAs.Clear()
        ROWs.Clear()
        'DVWs.Clear()
        pROWs.Clear()
        temps.Clear()
    End Sub

    Public Function sqlSvrTemp_Table( _
        Optional ByVal sql As String = "", _
        Optional ByVal FORM_NAME As String = "", _
        Optional ByVal XNO As String = "", _
        Optional ByVal custom_parameters As String = "", _
        Optional ByVal PARMs() As String = Nothing) As String

        If sql = "" Then
            sql = sqlServerSql
        End If

        'Try
        '    If FORM_NAME = "" Then FORM_NAME = ASCMAIN1.ActiveForm.MENU_ITEM_OBJECT
        '    If XNO = "" Then XNO = ASCMAIN1.ActiveForm.XNO
        'Catch ex As Exception
        'End Try


        Dim TABLE_NAME As String = "ASW" & sqlSvrNext_Control_No("TEMP_NO")
        'Dim sqlASTTEMP1 As String = "Insert into ASTTEMP1 Values ('" & TABLE_NAME & "','" & SESSION_NO & "',sysdate,'" & USER_ID & "','" & FORM_NAME & "','" & XNO & "'," & CStr(DBS_SESSION_ID) & ")"
        'ascsqls1.ExecuteSQL(sqlASTTEMP1)
        Dim insertAt = sql.ToUpper.IndexOf("FROM")

        If insertAt > 0 Then sql = sql.Insert(insertAt, "INTO " & TABLE_NAME & " ")
        sql = sql '& " INTO " & TABLE_NAME '"Create Table " & TABLE_NAME & " as " & sql
        If custom_parameters = "" Then
            sqlSvrExecuteSQL(sql)
        Else
            sqlSvrExecuteSQL(sql, custom_parameters, PARMs)
        End If
        Return TABLE_NAME
    End Function

    Public Function sqlSvrNext_Control_No( _
    ByVal CTL_NO_TYPE As String, _
    Optional ByVal How_Many As Long = 1) As String

        Dim CTL_NO_LAST As Long
        Dim CTL_NO_LENGTH As Integer

        Dim tblTATCTLN1 As New DataTable
        'Dim sql As String = "Select * from TATCTLN1 where CTL_NO_TYPE = '" & CTL_NO_TYPE & "'"
        Dim sql As String = "Select * from abs_TATCTLN1 (UPDLOCK) where CTL_NO_TYPE = '" & CTL_NO_TYPE & "'"
        Dim sqlAda As SqlDataAdapter = sqlSvrGetDataAdapter(tblTATCTLN1, "abs_TATCTLN1", sql, True)

        Dim rowTATCTLN1 As DataRow
        If tblTATCTLN1.Rows.Count = 0 Then
            rowTATCTLN1 = tblTATCTLN1.NewRow()
            rowTATCTLN1.Item("CTL_NO_TYPE") = CTL_NO_TYPE
            rowTATCTLN1.Item("CTL_NO_LAST") = 0
            tblTATCTLN1.Rows.Add(rowTATCTLN1)
        Else
            rowTATCTLN1 = tblTATCTLN1.Rows(0)
        End If

        CTL_NO_LENGTH = Val(rowTATCTLN1.Item("CTL_NO_LENGTH") & "")
        If CTL_NO_LENGTH = 0 Then
            'If InStr(CTL_NO_TYPE, ".") <> 0 Then
            '    Dim sqlCTL_NO_LENGTH As String = "Select DATA_LENGTH from USER_TAB_COLUMNS where TABLE_NAME = '" & Split(CTL_NO_TYPE, ".")(0) & "' and COLUMN_NAME = '" & Split(CTL_NO_TYPE, ".")(1) & "'"
            '    CTL_NO_LENGTH = Val(ascsqls1.GetDataValue(sqlCTL_NO_LENGTH) & "")
            'End If
            If CTL_NO_LENGTH = 0 Then
                CTL_NO_LENGTH = 10
            End If
            rowTATCTLN1.Item("CTL_NO_LENGTH") = CTL_NO_LENGTH
        End If

        If How_Many <= 0 Then
            How_Many = 1
        End If

        CTL_NO_LAST = Val(rowTATCTLN1.Item("CTL_NO_LAST") & "") + 1

        If CTL_NO_LAST >= 10 ^ CTL_NO_LENGTH Then
            CTL_NO_LAST = 1
        End If

        sqlSvrNext_Control_No = Format$(CTL_NO_LAST, "".PadLeft(CTL_NO_LENGTH, "0"))
        CTL_NO_LAST = CTL_NO_LAST + How_Many - 1
        If CTL_NO_LAST >= 10 ^ CTL_NO_LENGTH Then
            CTL_NO_LAST = CTL_NO_LAST - 10 ^ CTL_NO_LENGTH + 1
        End If
        rowTATCTLN1.Item("CTL_NO_LAST") = CTL_NO_LAST
        Dim SUCCESS As Integer = sqlAda.Update(tblTATCTLN1)

        'sqlAda.Dispose()
        'sqlAda = Nothing

    End Function

    Public Sub sqlSvrEnforceConstraints(Optional ByVal enforce As Boolean = True) 'As String
        Dim hasConstraintErrors As Boolean = False

        Try
            sqlServerDst.EnforceConstraints = enforce
        Catch ex As ConstraintException
            hasConstraintErrors = True
        End Try

        If Not hasConstraintErrors Then
            'Return ""
        End If

    End Sub

    ''' <summary>
    ''' Returns a DataTable representing the result set for the sql provided
    '''  using the (assumed to be open) db Connection 
    ''' </summary>
    ''' <param name="sql"></param>
    ''' <param name="TABLE_NAME"></param>
    ''' <param name="Key_Field_Count"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function sqlSvrGetDataTable( _
    Optional ByVal sql As String = "", _
    Optional ByVal TABLE_NAME As String = "", _
    Optional ByVal Key_Field_Count As Integer = -1, _
    Optional ByVal Include_Data As Boolean = True, _
    Optional ByVal NumberOfKeysUsedToSelect As Integer = -1) _
    As DataTable

        Dim tbl As New DataTable
        Dim ada As SqlDataAdapter = sqlSvrGetDataAdapter(tbl, TABLE_NAME, sql, False, Key_Field_Count, Include_Data, NumberOfKeysUsedToSelect)
        ada.Dispose()
        Return tbl

    End Function

    Public Function sqlSvrGetDataAdapter( _
    ByRef tbl As DataTable, _
    ByVal TABLE_NAME As String, _
    ByVal sql_custom As String, _
    Optional ByVal for_update As Boolean = False, _
    Optional ByVal Key_Field_Count As Integer = -1, _
    Optional ByVal Include_Data As Boolean = True, _
    Optional ByVal NumberOfKeysUsedToSelect As Integer = -1, _
    Optional ByVal Update_COLUMN_NAMEs As String = "", _
    Optional ByVal custom_parameters As String = "", _
    Optional ByVal tblTABLE_NAME As String = "") As SqlDataAdapter

        'If Oracle.oraCon.ConnectionString = "" Then
        '    Oracle.oraCon.ConnectionString = "Data Source=" & "" & ";User ID=" & "ODG" & ";Password=" & "ODG"
        '    Oracle.oraCon.Open()
        'End If

        Dim sql As String = String.Empty  ' ""

        ' need a GlobalConst Class
        'If sql_custom = GlobalConst.SelectAll Or sql_custom = GlobalConst.ExternalSelect Then
        '    If (sql_custom = GlobalConst.ExternalSelect And Not for_update) Then
        '        sql = Oracle.oraSql
        '    Else
        '        sql = "Select * from " & TABLE_NAME
        '    End If

        'End If

        If sql_custom = "*" Or sql_custom = "**" Then
            If (sql_custom = "**" And Not for_update) Then
                'sql =  sqlServerSql
                sql = sqlServerSql
            Else
                sql = "Select * from " & TABLE_NAME
            End If
        Else
            If sql_custom = "" Then
                sql = sqlServerSql
            Else
                sql = sql_custom
            End If
        End If

        Dim includes_where_clause As Boolean = (sql_custom = "**" And for_update) Or (Not (sql_custom = "*") And (InStr(sql.ToUpper, " WHERE ") <> 0))

        Dim Ada As New SqlDataAdapter
        Ada.SelectCommand = sqlServerConnection.CreateCommand
        Ada.SelectCommand.CommandText = sql
        If custom_parameters.Length > 0 Then
            Dim ptbl As New DataTable
            sqlSvrCreate_Parameters(Ada.SelectCommand, custom_parameters, ptbl)
            pROWs.Add(tblTABLE_NAME, ptbl.NewRow)
        End If

        If sqlServerTransaction IsNot Nothing Then
            Ada.SelectCommand.Transaction = sqlServerTransaction
        End If

        'Ada.SelectCommand.Parameters
        '    Dim ptbl As New DataTable
        '    Call ascsqls1.Create_Parameters(TDAs(tblTABLE_NAME).SelectCommand, custom_parameters, ptbl)
        '    pROWs.Add(tblTABLE_NAME, ptbl.NewRow)
        If tbl Is Nothing Then
            tbl = New DataTable
        End If

        If TABLE_NAME <> "" Then
            If tbl.TableName Like "Table*" Or tbl.TableName = "" Then
                tbl.TableName = TABLE_NAME
            End If
        End If

        Ada.FillSchema(tbl, SchemaType.Source)
        'Dim TBL2 As New DataTable
        'Ada.Fill(TBL2)
        'Stop
        '        Dim KeyFields_from_Table As Integer = tbl.PrimaryKey.Length

        Dim OBJPFX As String = ""
        Dim OBJSFX As String = ""

        If Key_Field_Count = -1 Then
            Key_Field_Count = tbl.PrimaryKey.Length ' KeyFields_from_Table
        Else
            If Key_Field_Count = 0 Then
                If tbl.PrimaryKey.Length <> 0 Then
                    For i As Integer = 0 To tbl.PrimaryKey.Length - 1
                        tbl.Columns(i).AllowDBNull = True
                    Next
                    tbl.PrimaryKey = Nothing
                End If
            Else

                Dim Key_Fields(Key_Field_Count - 1) As DataColumn
                For i As Integer = 1 To Key_Field_Count
                    Key_Fields(i - 1) = tbl.Columns(i - 1)
                Next
                If tbl.PrimaryKey Is Nothing Then
                Else
                    Try
                        tbl.PrimaryKey = Key_Fields
                    Catch
                    End Try
                End If
            End If
        End If

        Dim sql_key As String = ""
        Dim UCN As New List(Of String)

        If for_update Then
            Dim sql_cols As String = ""
            Dim sql_parms As String = ""

            If tbl.PrimaryKey.Length <> 0 Then
                For i As Integer = 0 To tbl.PrimaryKey.Length - 1
                    sql_key = sql_key & " and (" & OBJPFX & TABLE_NAME & "." & tbl.Columns(i).ColumnName & OBJSFX & "=" & "@" & tbl.Columns(i).ColumnName & ")"
                    sql_cols = sql_cols & "," & OBJPFX & tbl.Columns(i).ColumnName & OBJSFX
                    sql_parms = sql_parms & "," & "@" & tbl.Columns(i).ColumnName
                Next i
                sql_key = " where " & Mid(sql_key, 6)
            End If

            Dim sql_data As String = ""
            For i As Integer = tbl.PrimaryKey.Length To tbl.Columns.Count - 1
                sql_data = sql_data & "," & OBJPFX & TABLE_NAME & "." & tbl.Columns(i).ColumnName & OBJSFX & "=" & "@" & tbl.Columns(i).ColumnName
                sql_cols = sql_cols & "," & OBJPFX & tbl.Columns(i).ColumnName & OBJSFX
                sql_parms = sql_parms & "," & "@" & tbl.Columns(i).ColumnName
            Next i
            sql_data = " Set " & Mid(sql_data, 2)

            sql_cols = Mid(sql_cols, 2)
            sql_parms = Mid(sql_parms, 2)

            sql = "Insert into " & OBJPFX & TABLE_NAME & OBJSFX & " (" & sql_cols & ") values (" & sql_parms & ")"
            Ada.InsertCommand = New SqlCommand(sql, sqlServerConnection)
            'Ada.InsertCommand.BindByName = True

            Dim sql_data2 As String = ""
            If Update_COLUMN_NAMEs <> "" Then
                For Each Update_COLUMN_NAME As String In Split(Update_COLUMN_NAMEs, ",")
                    UCN.Add(UCase(Trim(Update_COLUMN_NAME)))
                    sql_data2 = sql_data2 & "," & OBJPFX & TABLE_NAME & "." & Trim(Update_COLUMN_NAME) & OBJSFX & "=" & "@" & Trim(Update_COLUMN_NAME)
                Next
                sql_data2 = " Set " & Mid(sql_data2, 2)
            End If
            If UCN.Count = 0 Then
                sql_data2 = sql_data
            End If
            'If TABLE_NAME = "APTINVH1" Then Stop

            sql = "Update " & OBJPFX & TABLE_NAME & OBJSFX & sql_data2 & Replace(sql_key, "=@", "=@U_")
            Ada.UpdateCommand = New SqlCommand(sql, sqlServerConnection)
            'Ada.UpdateCommand.BindByName = True

            sql = "Delete from " & OBJPFX & TABLE_NAME & OBJSFX & Replace(sql_key, "=@", "=@D_")
            Ada.DeleteCommand = New SqlCommand(sql, sqlServerConnection)
            'Ada.DeleteCommand.BindByName = True
        End If

        Dim COLUMN_NAME As String
        Dim DATA_DTYPE As String
        Dim DATA_LENGTH As Integer
        'Dim DATA_PRECISION As Integer
        'Dim DATA_SCALE As Integer

        If for_update Or NumberOfKeysUsedToSelect <> 0 Then
            For i As Integer = 0 To tbl.Columns.Count - 1

                'tbl.Columns(i).ReadOnly = False

                COLUMN_NAME = tbl.Columns(i).ColumnName
                DATA_DTYPE = tbl.Columns(i).DataType.ToString
                DATA_LENGTH = tbl.Columns(i).MaxLength
                'DATA_PRECISION = tbl.Columns(i)
                'DATA_SCALE = tbl.Columns(i)

                Dim p_I_CURR As New SqlParameter
                p_I_CURR.ParameterName = COLUMN_NAME
                p_I_CURR.SourceColumn = COLUMN_NAME
                p_I_CURR.SourceVersion = DataRowVersion.Current

                Dim p_D_ORIG As New SqlParameter
                p_D_ORIG.ParameterName = "D_" & COLUMN_NAME
                p_D_ORIG.SourceColumn = COLUMN_NAME
                p_D_ORIG.SourceVersion = DataRowVersion.Original

                Dim p_U_CURR As New SqlParameter
                p_U_CURR.ParameterName = COLUMN_NAME
                p_U_CURR.SourceColumn = COLUMN_NAME
                p_U_CURR.SourceVersion = DataRowVersion.Current

                Dim p_U_ORIG As New SqlParameter
                p_U_ORIG.ParameterName = "U_" & COLUMN_NAME
                p_U_ORIG.SourceColumn = COLUMN_NAME
                p_U_ORIG.SourceVersion = DataRowVersion.Original

                Dim p_S_ORIG As New SqlParameter
                p_S_ORIG.ParameterName = "S_" & COLUMN_NAME
                p_S_ORIG.SourceColumn = COLUMN_NAME
                p_S_ORIG.SourceVersion = DataRowVersion.Original

                Select Case DATA_DTYPE

                    Case "System.String", "System.Byte"
                        p_I_CURR.SqlDbType = SqlDbType.VarChar
                        p_U_CURR.SqlDbType = SqlDbType.VarChar
                        p_U_ORIG.SqlDbType = SqlDbType.VarChar
                        p_D_ORIG.SqlDbType = SqlDbType.VarChar
                        p_S_ORIG.SqlDbType = SqlDbType.VarChar

                        p_I_CURR.Size = DATA_LENGTH
                        p_U_CURR.Size = DATA_LENGTH
                        p_U_ORIG.Size = DATA_LENGTH
                        p_D_ORIG.Size = DATA_LENGTH
                        p_S_ORIG.Size = DATA_LENGTH

                    Case "System.DateTime"
                        p_I_CURR.SqlDbType = SqlDbType.DateTime
                        p_U_CURR.SqlDbType = SqlDbType.DateTime
                        p_U_ORIG.SqlDbType = SqlDbType.DateTime
                        p_D_ORIG.SqlDbType = SqlDbType.DateTime
                        p_S_ORIG.SqlDbType = SqlDbType.DateTime

                    Case "System.Double"
                        p_I_CURR.SqlDbType = SqlDbType.Float
                        p_U_CURR.SqlDbType = SqlDbType.Float
                        p_U_ORIG.SqlDbType = SqlDbType.Float
                        p_D_ORIG.SqlDbType = SqlDbType.Float
                        p_S_ORIG.SqlDbType = SqlDbType.Float

                    Case "System.Single"
                        p_I_CURR.SqlDbType = SqlDbType.Real
                        p_U_CURR.SqlDbType = SqlDbType.Real
                        p_U_ORIG.SqlDbType = SqlDbType.Real
                        p_D_ORIG.SqlDbType = SqlDbType.Real
                        p_S_ORIG.SqlDbType = SqlDbType.Real

                        'p.DbType = DbType.Decimal
                        'p_U.DbType = DbType.Decimal
                        'Stop
                        'p.Precision = DATA_PRECISION
                        'p.Scale = DATA_SCALE
                        'p.Value = 0

                    Case "System.Decimal"
                        p_I_CURR.SqlDbType = SqlDbType.Decimal
                        p_U_CURR.SqlDbType = SqlDbType.Decimal
                        p_U_ORIG.SqlDbType = SqlDbType.Decimal
                        p_D_ORIG.SqlDbType = SqlDbType.Decimal
                        p_S_ORIG.SqlDbType = SqlDbType.Decimal

                    Case "System.Int16"
                        p_I_CURR.SqlDbType = SqlDbType.SmallInt
                        p_U_CURR.SqlDbType = SqlDbType.SmallInt
                        p_U_ORIG.SqlDbType = SqlDbType.SmallInt
                        p_D_ORIG.SqlDbType = SqlDbType.SmallInt
                        p_S_ORIG.SqlDbType = SqlDbType.SmallInt

                    Case "System.Int32"
                        p_I_CURR.SqlDbType = SqlDbType.Int
                        p_U_CURR.SqlDbType = SqlDbType.Int
                        p_U_ORIG.SqlDbType = SqlDbType.Int
                        p_D_ORIG.SqlDbType = SqlDbType.Int
                        p_S_ORIG.SqlDbType = SqlDbType.Int

                    Case "System.Int64"
                        p_I_CURR.SqlDbType = SqlDbType.BigInt
                        p_U_CURR.SqlDbType = SqlDbType.BigInt
                        p_U_ORIG.SqlDbType = SqlDbType.BigInt
                        p_D_ORIG.SqlDbType = SqlDbType.BigInt
                        p_S_ORIG.SqlDbType = SqlDbType.BigInt

                    Case Else
                        'Stop
                        Throw New Exception("Unsupported Data Type")

                End Select

                If for_update Then
                    Ada.InsertCommand.Parameters.Add(p_I_CURR)
                End If


                If i <= Key_Field_Count - 1 Then
                    If for_update Then
                        Ada.DeleteCommand.Parameters.Add(p_D_ORIG)
                        'If UCN.Count = 0 OrElse UCN.Contains(COLUMN_NAME) Then
                        Ada.UpdateCommand.Parameters.Add(p_U_ORIG)
                        'End If
                    End If
                    If NumberOfKeysUsedToSelect <> 0 And _
                    NumberOfKeysUsedToSelect >= i + 1 Then
                        '(NumberOfKeysUsedToSelect = -1 _
                        'Or NumberOfKeysUsedToSelect >= i + 1) Then
                        'Ada.DeleteCommand.Parameters.Add(p_D_ORIG)
                        'Ada.UpdateCommand.Parameters.Add(p_U_ORIG)
                        Ada.SelectCommand.Parameters.Add(p_S_ORIG)
                        'Ada.SelectCommand.Parameters(0).Value = 
                    End If
                Else
                    If for_update Then
                        If UCN.Count = 0 OrElse UCN.Contains(COLUMN_NAME) Then
                            Ada.UpdateCommand.Parameters.Add(p_U_CURR)
                        End If
                    End If
                End If
            Next i
        End If

        sql_key = ""

        Dim XDT As Date = Now

        If sql_custom = "**" And for_update Then
            Ada.SelectCommand.CommandText = sqlServerSql
            ' next line will put a key on tbl (ICFINVA1.ICFINVAT) where you might not want one 
            ' (like if Key_Field_Count = 0) - but this is nec for ASFMENU1.ASTSECM1,
            '  so the trick is to remember to set the for_update flag to false 
            ' for tables like ICTINVAT
            Ada.FillSchema(tbl, SchemaType.Source)
        End If

        If Include_Data Then
            Try
                Ada.Fill(tbl)
            Catch ex As Exception
                Stop
            End Try
        End If


        If NumberOfKeysUsedToSelect <> 0 Then
            Dim KEYS As Integer = 0
            If NumberOfKeysUsedToSelect <> -1 Then
                KEYS = NumberOfKeysUsedToSelect
            Else
                If tbl.PrimaryKey.Length <> 0 Then
                    KEYS = tbl.PrimaryKey.Length
                Else
                    KEYS = 0
                End If
            End If

            If KEYS <> 0 Then
                For i As Integer = 0 To KEYS - 1
                    'IIF INSTALLED WHEN THE DATATABLE NAME IS A BOGUS ONE, AND THE TDA IS NOT GOING TO BE USED FOR UPDATE
                    ' STILL NOT THE BEST SOLUTION, SINCE WE MAY BE USING KEYS TO GRAB A RECORD FROM A TABLE WHERE THERE MAY BE AMBIGUITY

                    sql_key = sql_key & " and (" & OBJPFX & IIf(for_update, TABLE_NAME & ".", "") & tbl.Columns(i).ColumnName & OBJSFX & "=" & "@" & tbl.Columns(i).ColumnName & ")"
                Next i
                If Not includes_where_clause Then
                    sql_key = " where " & Mid(sql_key, 6)
                End If
                'Ada.SelectCommand.BindByName = True
                Ada.SelectCommand.CommandText &= Replace(sql_key, "=@", "=@S_")
            End If
        End If


        Return Ada
    End Function

    Public Sub sqlSvrCreate_Parameters( _
    ByRef cmd As SqlCommand, _
    ByVal custom_parameters As String, _
    Optional ByRef ptbl As DataTable = Nothing)

        With cmd.Parameters
            For i As Integer = 1 To Len(custom_parameters)
                Dim COLUMN_NAME As String = "PARM" & CStr(i)
                'If 1 = 1 Then
                '    COLUMN_NAME = ""
                'End If

                Select Case Mid$(custom_parameters, i, 1)
                    Case "V"    ' Varchar2
                        .Add(COLUMN_NAME, SqlDbType.VarChar)
                        If ptbl IsNot Nothing Then
                            ptbl.Columns.Add(COLUMN_NAME, GetType(System.String))
                        End If
                    Case "I"    ' Integer
                        .Add(COLUMN_NAME, SqlDbType.Int)
                        If ptbl IsNot Nothing Then
                            ptbl.Columns.Add(COLUMN_NAME, GetType(System.Int32))
                        End If
                    Case "N"    ' Number
                        .Add(COLUMN_NAME, SqlDbType.Decimal)
                        If ptbl IsNot Nothing Then
                            ptbl.Columns.Add(COLUMN_NAME, GetType(System.Double))
                        End If
                    Case "D"    ' Date
                        .Add(COLUMN_NAME, SqlDbType.DateTime)
                        If ptbl IsNot Nothing Then
                            ptbl.Columns.Add(COLUMN_NAME, GetType(System.DateTime))
                        End If
                End Select
            Next
        End With
    End Sub

    ''' <summary>
    ''' Executes the sql statement passed in using the Oracle.oraCmd Object (which originates from the Oracle.oraCon Connection object).
    ''' If you have begun a Transacton with Oracle.oraCon, then this statement will execute within the context of that transaction.
    ''' </summary>
    ''' <param name="sql_to_Execute"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function sqlSvrExecuteSQL( _
    ByVal sql_to_Execute As String, _
    ByVal custom_parameters As String, _
    ByVal ParamArray PARMs() As Object) As Integer
        Using cmd As New SqlCommand(sql_to_Execute, sqlServerConnection)
            If custom_parameters <> "" Then
                sqlSvrCreate_Parameters(cmd, custom_parameters)
                For i As Integer = 1 To Len(custom_parameters)
                    With cmd.Parameters("PARM" & CStr(i))
                        'Debug.Print(PARMs(i - 1) & "")
                        If PARMs Is Nothing Then
                            '.Value = ""
                            .Value = System.DBNull.Value
                        Else
                            If PARMs.Length >= i Then
                                .Value = PARMs(i - 1)
                            Else
                                '.Value = ""
                                .Value = System.DBNull.Value
                            End If
                        End If
                    End With
                Next
            End If
            'cmd.Parameters(0).
            'cmd.Parameters.Add("p1", OracleType.VarChar, 255);
            'cmd.Parameters[0].Direction = ParameterDirection.Output;

            cmd.CommandText = sql_to_Execute
            Try
                'cmd.BindByName = True

                Dim XDT As Date = Now

                sqlSvrExecuteSQL = cmd.ExecuteNonQuery

                Dim commandText As String = cmd.CommandText
                If commandText.Length > 4000 Then
                    commandText = commandText.Substring(0, 4000).Trim
                End If

            Catch ex As Exception
                Stop
            End Try
        End Using
    End Function

    ''' <summary>
    ''' Executes the sql statement passed in using the Oracle.oraCmd Object (which originates from the Oracle.oraCon Connection object).
    ''' If you have begun a Transacton with Oracle.oraCon, then this statement will execute within the context of that transaction.
    ''' </summary>
    ''' <param name="sql_to_Execute"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function sqlSvrExecuteSQL( _
    Optional ByVal sql_to_Execute As String = "", _
    Optional ByVal ok_to_fail As Boolean = False) As Integer
        If sql_to_Execute = "" Then
            sql_to_Execute = sqlServerSql
        End If

        Try

            Dim XDT As Date = Now

            sqlServerCommand.CommandText = sql_to_Execute
            sqlServerCommand.Connection = sqlServerConnection

            Dim icmd As Int32 = sqlServerCommand.ExecuteNonQuery()

            Dim commandText As String = sql_to_Execute
            If commandText.Length > 4000 Then
                commandText = commandText.Substring(0, 4000).Trim
            End If

            Return icmd
        Catch ex As Exception
            Dim oex As SqlException
            oex = DirectCast(ex, SqlException)
        End Try
    End Function

    ''' <summary>
    ''' Get a datarow using the SQL passed in the first parameter, and return an emtpy row if missing depending on the second parameter
    ''' </summary>
    ''' <param name="sql_to_Execute">Defaults to Oracle.oraSql if omitted</param>
    ''' <param name="Return_Empty_Row_if_Missing">Defaults to False if omitted</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function sqlSvrGetDataRow( _
    Optional ByVal sql_to_Execute As String = "", _
    Optional ByVal Return_Empty_Row_if_Missing As Boolean = False) As DataRow
        Try
            Dim cmd As SqlCommand = sqlServerConnection.CreateCommand
            If sql_to_Execute = "" Then
                sql_to_Execute = sqlServerSql
            End If
            Dim tbl As DataTable = sqlSvrGetDataTable(sql_to_Execute)
            Dim row As DataRow
            If tbl.Rows.Count = 0 Then
                If Return_Empty_Row_if_Missing Then
                    row = tbl.NewRow
                Else
                    row = Nothing
                End If
            Else
                row = tbl.Rows(0)
            End If
            Return row
        Catch ex As Exception
            Stop
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Get a datarow using the SQL set in the global variable and doesn't return empty row if missing
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function sqlSvrGetDataRowL()
        Return sqlSvrGetDataRowL(String.Empty, False)
    End Function

    ''' <summary>
    ''' Get a datarow using the SQL passed in the first parameter, and doesn't return empty row if missing
    ''' </summary>
    ''' <param name="sqlToExecute"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function sqlSvrGetDataRowL(ByVal sqlToExecute As String)
        Return sqlSvrGetDataRowL(sqlToExecute, False)
    End Function

    ''' <summary>
    ''' Get a datarow using the SQL set in the global variable, and return a an empty row if missing depending on the parameter passed
    ''' </summary>
    ''' <param name="returnEmptyRowIfMissing"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function sqlSvrGetDataRowL(ByVal returnEmptyRowIfMissing As Boolean)
        Return sqlSvrGetDataRowL(String.Empty, returnEmptyRowIfMissing)
    End Function

    ''' <summary>
    ''' Get a datarow using the SQL passed in the first parameter, and return an emtpy row if missing depending on the second parameter
    ''' </summary>
    ''' <param name="sql_to_Execute"></param>
    ''' <param name="Return_Empty_Row_if_Missing"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function sqlSvrGetDataRowL( _
    ByVal sql_to_Execute As String, _
    ByVal Return_Empty_Row_if_Missing As Boolean) As DataRow
        Try
            Dim cmd As SqlCommand = sqlServerConnection.CreateCommand
            If sql_to_Execute = "" Then
                sql_to_Execute = sqlServerSql
            End If
            Dim tbl As DataTable = sqlSvrGetDataTable(sql_to_Execute)
            Dim row As DataRow
            If tbl.Rows.Count = 0 Then
                If Return_Empty_Row_if_Missing Then
                    row = tbl.NewRow
                Else
                    row = Nothing
                End If
            Else
                row = tbl.Rows(0)
            End If
            Return row
        Catch ex As Exception
            Stop
            Return Nothing
        End Try
    End Function

    Public Function sqlSvrGetDataValue(Optional ByVal sql_to_Execute As String = "") As String
        Try
            Dim cmd As SqlCommand = sqlServerConnection.CreateCommand()
            If sql_to_Execute = "" Then
                sql_to_Execute = sqlServerSql
            End If
            cmd.CommandText = sql_to_Execute
            Dim result As String = cmd.ExecuteScalar & ""
            Return result
        Catch
            Return ""
        End Try
    End Function

    Public Function sqlSvrTable2Hash(ByVal tbl As DataTable, ByVal COLUMN_NAME_key As String, ByVal COLUMN_NAME_item As String)
        Dim htb As New Hashtable
        For Each r As DataRow In tbl.Rows
            Try
                htb.Add(r.Item(COLUMN_NAME_key), r.Item(COLUMN_NAME_item))
            Catch ex As Exception
            End Try
        Next
        Return htb
    End Function

    Public Function sqlSvrGetRelation(ByVal dst As DataSet, _
    ByVal TABLE_NAME_parent As String, _
    ByVal TABLE_NAME_child As String, _
    Optional ByVal COLUMN_NAMEs As String = "") As DataRelation

        Dim DC_parent() As DataColumn
        Dim DC_child() As DataColumn
        ReDim DC_parent(0)
        ReDim DC_child(0)
        If COLUMN_NAMEs = "" Then
            Dim dt As DataTable = dst.Tables(TABLE_NAME_parent)
            For i As Integer = 0 To dt.PrimaryKey.Length - 1
                Dim COLUMN_NAME As String = dt.Columns(i).ColumnName
                COLUMN_NAMEs &= COLUMN_NAME
                ReDim Preserve DC_parent(i)
                DC_parent(i) = dst.Tables(TABLE_NAME_parent).Columns(COLUMN_NAME)
                ReDim Preserve DC_child(i)
                DC_child(i) = dst.Tables(TABLE_NAME_child).Columns(COLUMN_NAME)
            Next
            COLUMN_NAMEs = Mid$(COLUMN_NAMEs, 2)
        Else
            Dim i As Integer = 0
            For Each COLUMN_NAME As String In Split(COLUMN_NAMEs, ",")
                ReDim Preserve DC_parent(i)
                DC_parent(i) = dst.Tables(TABLE_NAME_parent).Columns(COLUMN_NAME)
                ReDim Preserve DC_child(i)
                DC_child(i) = dst.Tables(TABLE_NAME_child).Columns(COLUMN_NAME)
                i = i + 1
            Next
        End If

        'Dim R As New DataRelation( _
        'TABLE_NAME_parent & "_" & TABLE_NAME_child, _
        'TABLE_NAME_parent, TABLE_NAME_child, _
        'Split(COLUMN_NAMEs, ","), _
        'Split(COLUMN_NAMEs, ","), False)

        Dim R As New DataRelation( _
        TABLE_NAME_parent & "_" & TABLE_NAME_child, _
        DC_parent, _
        DC_child)

        Return R
    End Function

    Public Function sqlSvrDeleteRows( _
    ByVal TABLE_NAME As String, _
    ByVal sqlWhere As String)

        Dim rows() As DataRow = sqlServerDst.Tables(TABLE_NAME).Select(sqlWhere)
        If rows.Length > 0 Then
            For Each row As DataRow In rows
                'dst.Tables("EDTMAPT2").Rows.Remove(row)
                row.Delete()
            Next
        End If
        Return rows.Length
    End Function

#Region "SelectDistinct"

    Public Function sqlSvrSelectDistinct( _
    ByVal Rows As DataRow(), _
    ByVal ParamArray FieldNames() As String) As DataTable
        If Rows.Length = 0 Then
            Return New DataTable
        Else
            Dim SourceTable As DataTable = Rows(0).Table.Clone

            Try
                SourceTable.LoadDataRow(Rows, True)

            Catch ex As Exception
                For I As Long = 0 To Rows.Length - 1
                    Dim RR As DataRow = SourceTable.NewRow
                    RR.ItemArray = Rows(I).ItemArray
                    'For J As Integer = 0 To RR.ItemArray.Length - 1
                    '    RR.Item(J) = Rows(I).Item(J)
                    'Next
                    SourceTable.Rows.Add(RR)
                Next
                SourceTable.AcceptChanges()
            End Try
            'Stop
            Return sqlSvrSelectDistinct(SourceTable, FieldNames)
        End If

    End Function

    Public Function sqlSvrSelectDistinct( _
    ByVal SourceTableName As String, _
    ByVal ParamArray FieldNames() As String) As DataTable
        Return sqlSvrSelectDistinct(sqlServerDst.Tables(SourceTableName), FieldNames)
    End Function

    Public Function sqlSvrSelectDistinct( _
    ByVal SourceTable As DataTable, _
    ByVal ParamArray FieldNames() As String) As DataTable

        Dim lastValues() As Object
        Dim newTable As DataTable

        If FieldNames Is Nothing OrElse FieldNames.Length = 0 Then
            Throw New ArgumentNullException("FieldNames")
        End If

        lastValues = New Object(FieldNames.Length - 1) {}
        newTable = New DataTable

        For Each field As String In FieldNames
            newTable.Columns.Add(field, SourceTable.Columns(field).DataType)
        Next

        For Each Row As DataRow In SourceTable.Select("", String.Join(", ", FieldNames))
            If Not sqlSvrfieldValuesAreEqual(lastValues, Row, FieldNames) Then
                newTable.Rows.Add(sqlSvrcreateRowClone(Row, newTable.NewRow(), FieldNames))

                sqlSvrsetLastValues(lastValues, Row, FieldNames)
            End If
        Next

        Return newTable
    End Function

    Private Function sqlSvrfieldValuesAreEqual(ByVal lastValues() As Object, ByVal currentRow As DataRow, ByVal fieldNames() As String) As Boolean
        Dim areEqual As Boolean = True

        For i As Integer = 0 To fieldNames.Length - 1
            If lastValues(i) Is Nothing OrElse Not lastValues(i).Equals(currentRow(fieldNames(i))) Then
                areEqual = False
                Exit For
            End If
        Next

        Return areEqual
    End Function

    Private Function sqlSvrcreateRowClone(ByVal sourceRow As DataRow, ByVal newRow As DataRow, ByVal fieldNames() As String) As DataRow
        For Each field As String In fieldNames
            newRow(field) = sourceRow(field)
        Next

        Return newRow
    End Function

    Private Sub sqlSvrsetLastValues(ByVal lastValues() As Object, ByVal sourceRow As DataRow, ByVal fieldNames() As String)
        For i As Integer = 0 To fieldNames.Length - 1
            lastValues(i) = sourceRow(fieldNames(i))
        Next
    End Sub

#End Region

End Class
