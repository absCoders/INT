Imports ABSolution
Imports System.Data.OleDb

Public Class EDCPRMDB

#Region "Class Variables"

    Private clsMdbDatabaseQuery As String
    Private clsMdbDatabaseConnectionString As String
    Private clsMdbFields() As String
    Private clsMdbTablename As String
    Private clsMdbLocation As String
    Private clsMdbTable As DataTable = Nothing

    Private clsOracleTableName As String
    Private clsOracleFields() As String
    Private clsOracleTable As DataTable = Nothing

    Private ErrorList As List(Of String) = New List(Of String)

    Private Const mdbTableNameDefault = "WorkTable"

    Private mdbConnection As OleDbConnection = Nothing
    Private mdbCommand As OleDbCommand = Nothing
    Private mdbDataAdapter As OleDbDataAdapter = Nothing
    Private mdbResultSet As DataSet = Nothing

    Const integerDataType As Integer = 2
    Const longIntegerDataType As Integer = 3
    Const singleDataType As Integer = 4
    Const doubleDataType As Integer = 5
    Const currencyDataType As Integer = 6
    Const dateDataType As Integer = 7
    Const booleanDataType As Integer = 11
    Const byteDataType As Integer = 17
    Const textMemoDataType As Integer = 130
    Const decimalDataType As Integer = 131

#End Region

#Region "Class Constructors"

    ''' <summary>
    ''' Class Constructor
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

        clsMdbDatabaseQuery = String.Empty
        clsMdbDatabaseConnectionString = String.Empty
        ReDim clsMdbFields(1)
        clsMdbTablename = String.Empty
        clsMdbLocation = String.Empty
        clsMdbTable = Nothing

        clsOracleTableName = String.Empty
        ReDim clsOracleFields(1)
        clsOracleTable = Nothing

        ErrorList = New List(Of String)

        mdbConnection = Nothing
        mdbCommand = Nothing
        mdbDataAdapter = Nothing
        mdbResultSet = Nothing

    End Sub

#End Region

#Region "Class Properties / Property type Procedures"

    ''' <summary>
    ''' Creates the Connection String for MS Access MDB connection
    ''' </summary>
    ''' <param name="JetVersion">Version of the Jet Engine. Ex: 4.0, 10.0</param>
    ''' <param name="DataSource">The full path of the Database including name</param>
    ''' <param name="UserID">Userid to gain access to the MDB, may be blank</param>
    ''' <param name="Password">Password to gain access to the mdb, may be blank</param>
    ''' <remarks>Sample: Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\temp\mydb.mdb;User ID=Admin;Password=Admin</remarks>
    Public Sub CreateConnectionSting(ByVal JetVersion As Double, ByVal DataSource As String, _
                                    ByVal UserID As String, ByVal Password As String)

        ' Sample Connection String
        ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\temp\mydb.mdb;User ID=Admin;Password="

        MdbDatabaseConnectionString = "Provider=Microsoft.Jet.OLEDB." & JetVersion.ToString("0.0") & ";" _
            & "Data Source=" & DataSource.Trim _
            & ";User ID=" & UserID.Trim & ";" _
            & "Password=" & Password.Trim

        Me.clsMdbLocation = DataSource
    End Sub

    ''' <summary>
    ''' Clears the Error List
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ClearErrors()
        ErrorList.Clear()
    End Sub

    ''' <summary>
    ''' Gets the list of Errors recorded in the last fuction call
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Errors() As List(Of String)
        Get
            Return ErrorList
        End Get
    End Property

    ''' <summary>
    ''' Gets / Sets a Connection to a MDB 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property GetMdbConnection() As OleDbConnection
        Get
            Return mdbConnection
        End Get

        Set(ByVal value As OleDbConnection)
            mdbConnection = value
        End Set
    End Property

    ''' <summary>
    ''' Get / Set MDB ConnectionString
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property MdbDatabaseConnectionString() As String
        Get
            Return clsMdbDatabaseConnectionString
        End Get
        Set(ByVal value As String)
            clsMdbDatabaseConnectionString = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Get / Set MDB Query String
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property MdbDatabaseQuery() As String
        Get
            Return clsMdbDatabaseQuery
        End Get
        Set(ByVal value As String)
            clsMdbDatabaseQuery = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Get / Set MDB List of fields, they should be in the same order and the Oralce Fields
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property MdbFields() As String()
        Get
            Return clsMdbFields
        End Get
        Set(ByVal value As String())
            clsMdbFields = value
        End Set
    End Property

    ''' <summary>
    ''' Returns the data extracted from the MS Access MDB
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property MDBTable() As DataTable
        Get
            Return clsMdbTable
        End Get
    End Property

    ''' <summary>
    ''' Gets / Sets the MDB Table to clone into Oracle
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property MdbTableName() As String
        Get
            Return clsMdbTablename
        End Get
        Set(ByVal value As String)
            clsMdbTablename = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Get / Set Oracle List of fields, they should be in the same order and the MDB Fields
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property OracleFields() As String()
        Get
            Return clsOracleFields
        End Get
        Set(ByVal value As String())
            clsOracleFields = value
        End Set
    End Property

    ''' <summary>
    ''' Get / Set Oracle Tablename
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property OracleTableName() As String
        Get
            Return clsOracleTableName
        End Get
        Set(ByVal value As String)
            clsOracleTableName = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Gets the class tables filled by the CopyMdbIntoOracle Function
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property OracleTable() As DataTable
        Get
            Return clsOracleTable
        End Get
    End Property

#End Region

#Region "Public Procedures"

    ''' <summary>
    ''' Copies the MS Access MDB data into an Oracle Table
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CopyMdbIntoOracle() As DataTable
        Return CopyMdbIntoOracle(clsMdbTablename, clsOracleTableName, True)
    End Function

    ''' <summary>
    ''' Copies the MS Access MDB data into an Oracle Table
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CopyMdbIntoOracle(ByVal vMdbTableName As String, _
            ByVal vOracleTableName As String, ByVal UseFieldLists As Boolean) As DataTable

        ' Clear any existing Errors
        ErrorList.Clear()

        ' Validate the data.
        If mdbConnection Is Nothing Then
            If Me.CreateMdbConnection = False Then
                Return Nothing
            End If
        End If

        vOracleTableName = vOracleTableName.Trim
        If vOracleTableName.Length = 0 Then
            ErrorList.Add("Missing Oracle tablename")
            Return Nothing
        End If

        vMdbTableName = vMdbTableName.Trim
        If vMdbTableName.Length = 0 Then
            ErrorList.Add("Missing MDB tablename")
            Return Nothing
        End If

        If UseFieldLists = True Then

            If MdbFields.Length = 0 Then
                ErrorList.Add("Missing MDB fields")
                Return Nothing
            End If

            If OracleFields.Length = 0 Then
                ErrorList.Add("Missing Oracle fields")
                Return Nothing
            End If

            If MdbFields.Length <> OracleFields.Length Then
                ErrorList.Add("Oracle field count is unequal MDB field count.")
                Return Nothing
            End If

            For i As Integer = MdbFields.GetLowerBound(0) To MdbFields.GetUpperBound(0)
                If MdbFields(i).Trim.Length = 0 Then
                    ErrorList.Add("MDB field array contains Null field name.")
                    Return Nothing
                End If
            Next

            For i As Integer = OracleFields.GetLowerBound(0) To OracleFields.GetUpperBound(0)
                If OracleFields(i).Trim.Length = 0 Then
                    ErrorList.Add("Oracle field array contains Null field name.")
                    Return Nothing
                End If
            Next

        End If

        Try
            ' Connect to the MDB
            clsMdbTable = Me.FetchAccessDataTable("Select * From " & vMdbTableName)
            If clsMdbTable Is Nothing Then
                ErrorList.Add("MDB table " & MdbDatabaseQuery & " is nothing.")
                Return Nothing
            End If

            Dim clsDataRow As DataRow = Nothing
            Dim fieldname As String = String.Empty

            Dim sql As String = "Select * From " & vOracleTableName & " Where Rownum < 1"
            clsOracleTable = ASCDATA1.GetDataTable(sql, String.Empty, 0, False, 0)

            If clsOracleTable Is Nothing Then
                ErrorList.Add("Could not access Oracle Table: " & vOracleTableName)
                Return Nothing
            End If

            For Each mdbRow As DataRow In clsMdbTable.Rows
                clsDataRow = clsOracleTable.NewRow

                If UseFieldLists = True Then
                    For i As Integer = MdbFields.GetLowerBound(0) To MdbFields.GetUpperBound(0)
                        fieldname = MdbFields(i).Trim
                        Try
                            clsDataRow.Item(fieldname) = mdbRow.Item(fieldname)
                        Catch ex As Exception
                            Stop
                            ' Need to catch Null errors were we try to put a null into the table
                            ' We shold allow this, but at times it causes an error
                        End Try
                    Next

                Else
                    For i As Integer = 0 To clsOracleTable.Columns.Count - 1
                        fieldname = clsOracleTable.Columns(i).ColumnName
                        Try
                            clsDataRow.Item(fieldname) = mdbRow.Item(fieldname)
                        Catch ex As Exception
                            ' Stop
                            ' Need to catch Null errors were we try to put a null into the table
                            ' We shold allow this, but at times it causes an error
                        End Try
                    Next
                End If

                clsOracleTable.Rows.Add(clsDataRow)
            Next

        Catch ex As Exception
            ErrorList.Add(ex.Message)
            Return Nothing
        End Try

        Return clsOracleTable

    End Function


    ''' <summary>
    ''' Establishes a connection to a MDB
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateMdbConnection() As Boolean
        Return CreateMdbConnection(MdbDatabaseConnectionString)
    End Function

    ''' <summary>
    ''' Establishes a connection to a MDB
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateMdbConnection(ByVal ConnectionString As String) As Boolean

        Try
            'Dispose the connector objects
            If Not (mdbConnection Is Nothing) Then mdbConnection.Dispose()
            mdbConnection = Nothing

        Catch ex As Exception
            ErrorList.Add(ex.Message)

        End Try

        ConnectionString = ConnectionString.Trim
        If ConnectionString.Length = 0 Then
            ErrorList.Add("Missing Mdb Database Connection String")
            Return False
        End If

        'Instantiate the DB Conection
        Try
            mdbConnection = New OleDbConnection(ConnectionString)
            mdbConnection.Open()
            Return True
        Catch ex As Exception
            Me.ErrorList.Add(ex.Message)
            Return False
        End Try

    End Function


    ''' <summary>
    ''' Creates an Orcale table from 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateOracleTableFromMdbSchema() As DataTable
        Return CreateOracleTableFromMdbSchema(clsMdbTablename, clsOracleTableName)
    End Function

    ''' <summary>
    ''' Creates an Oracle Table from an MS Access table. Only Creates the table Schema
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateOracleTableFromMdbSchema(ByVal vMdbTableName As String, ByVal vOracleTableName As String) As DataTable

        Dim oracleTable As DataTable = Nothing

        CreateOracleTableFromMdbSchema = oracleTable

        ' Validate data
        If vMdbTableName.Length = 0 Then
            ErrorList.Add("Missing MS Access Tablename")
            Return oracleTable
        End If

        If vOracleTableName.Length = 0 Then
            vOracleTableName = vMdbTableName
            ErrorList.Add("Missing Oracle Tablename, using MDB table name")
        End If

        If Me.mdbConnection Is Nothing Then
            If Not CreateMdbConnection() Then
                Return oracleTable
            End If
        End If

        Dim sql As String = "Select * From User_Tables Where Table_Name = '" & vOracleTableName & "'"
        If ASCDATA1.GetDataRow(sql) IsNot Nothing Then
            Errors.Add("The supplied oracle table " & vOracleTableName & " already exists.")
            Return oracleTable
        End If

        Dim mdbQueryString As String = "Select * From " & vMdbTableName & " Where 1 = 2"
        clsMdbTable = Me.FetchAccessDataTable(mdbQueryString)

        If clsMdbTable Is Nothing Then
            ErrorList.Add("MDB table is nothing.")
            Return oracleTable
        End If

        Try
            Dim ddl As String = GetOracleDDL(vMdbTableName, vOracleTableName)
            ddl = ddl.Trim
            If ddl.Length > 0 Then
                ASCDATA1.ExecuteSQL(ddl)
                oracleTable = ASCDATA1.GetDataTable("Select * From " & vOracleTableName & " Where Rownum < 1")
            End If
        Catch ex As Exception
            Errors.Add(ex.Message)
        End Try

        Return oracleTable

    End Function


    ''' <summary>
    ''' Fetches the MS Access MDB Query data and places it into a dataset
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FetchAccessDataTable() As DataTable
        Return FetchAccessDataTable(clsMdbDatabaseQuery)
    End Function

    ''' <summary>
    ''' Fetches the MS Access MDB Query data and places it into a dataset
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FetchAccessDataTable(ByVal mdbQueryString As String) As DataTable

        Try
            If mdbConnection Is Nothing Then
                If Not CreateMdbConnection() Then
                    Return Nothing
                End If
            End If

            mdbResultSet = New DataSet
            mdbCommand = New OleDbCommand(mdbQueryString, mdbConnection)
            mdbDataAdapter = New OleDbDataAdapter(mdbCommand)
            mdbDataAdapter.Fill(mdbResultSet, mdbTableNameDefault)

        Catch ex As OleDb.OleDbException
            Me.ErrorList.Add("OleDbException: " & ex.Message)
            FetchAccessDataTable = Nothing

        Catch ex As Exception
            Me.ErrorList.Add(ex.Message)
            FetchAccessDataTable = Nothing

        Finally
            'Dispose the connector objects
            If Not (mdbCommand Is Nothing) Then mdbCommand.Dispose()
            mdbCommand = Nothing
            If Not (mdbDataAdapter Is Nothing) Then mdbDataAdapter.Dispose()
            mdbDataAdapter = Nothing
        End Try

        'Return results
        Try
            Return mdbResultSet.Tables(mdbTableNameDefault)
        Catch ex As Exception
            Return Nothing
        End Try

    End Function


    ''' <summary>
    ''' Returns a list of Tables from an MS Access MDB
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetListOfMdbTables() As List(Of String)
        Return GetListOfMdbTables(mdbConnection)
    End Function

    ''' <summary>
    ''' Returns a list of Tables from an MS Access MDB
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetListOfMdbTables(ByVal MdbDBConnection As OleDbConnection) As List(Of String)

        Dim tblList As List(Of String) = New List(Of String)
        GetListOfMdbTables = tblList

        'Define the connectors
        If (MdbDBConnection Is Nothing) Then
            If Not CreateMdbConnection() Then
                Return tblList
            End If
        End If

        Try
            Dim restrictions() As String = New String() {Nothing, Nothing, Nothing, "TABLE"}
            Dim tbls As DataTable = mdbConnection.GetSchema("Tables", restrictions)
            For Each rowTablename As DataRow In tbls.Rows
                tblList.Add(rowTablename.Item("TABLE_NAME"))
            Next

        Catch ex As OleDb.OleDbException
            Me.ErrorList.Add("OleDbException: " & ex.Message)

        Catch ex As Exception
            Me.ErrorList.Add(ex.Message)

        Finally
            ' Nothing
        End Try

        GetListOfMdbTables = tblList

    End Function


    ''' <summary>
    ''' Returns the Oracle DDL to Create a table in Oracle using a table Schema from MS Access
    ''' </summary>
    ''' <param name="vMdbTableName">MS Access Table Name</param>
    ''' <param name="vOracleTableName">Oracle table Name</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetOracleDDL(ByVal vMdbTableName As String, ByVal vOracleTableName As String) As String

        GetOracleDDL = String.Empty
        Dim oracleDDL As String = String.Empty

        ' Validate data
        If vMdbTableName.Length = 0 Then
            ErrorList.Add("Missing MS Access Tablename")
            Return False
        End If

        If Me.mdbConnection Is Nothing Then
            If Not CreateMdbConnection() Then
                Return False
            End If
        End If

        'Dim mdbQueryString As String = "Select * From " & vMdbTableName & " Where 1 = 2"
        'Dim mdbTable As DataTable = Me.FetchAccessDataTable(mdbQueryString)

        ' {TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE}.
        Dim restrictions() = New String() {Nothing, Nothing, vMdbTableName, Nothing}
        Dim mdbTable As DataTable = mdbConnection.GetSchema("Columns", restrictions)

        If mdbTable Is Nothing Then
            ErrorList.Add("MDB table is nothing.")
            Return False
        End If

        Dim fieldLength As Integer = 0
        Try
            For Each rowColumns As DataRow In mdbTable.Select(String.Empty, "ORDINAL_POSITION")
                Dim COLUMN_NAME As String = rowColumns.Item("COLUMN_NAME").ToString.ToUpper

                oracleDDL &= ", " & COLUMN_NAME & " "

                Select Case Val(rowColumns.Item("DATA_TYPE") & String.Empty)

                    Case textMemoDataType
                        fieldLength = Val(rowColumns.Item("CHARACTER_MAXIMUM_LENGTH") & String.Empty)
                        If fieldLength > 2000 Then
                            fieldLength = 2000
                        ElseIf fieldLength = 0 Then
                            fieldLength = 2000 ' MEMO FIELD
                        End If
                        oracleDDL &= "VARCHAR2(" & CStr(fieldLength) & ")"

                    Case integerDataType
                        oracleDDL &= "NUMBER(5)"

                    Case doubleDataType
                        oracleDDL &= "NUMBER(15,6)"

                    Case decimalDataType
                        oracleDDL &= "NUMBER(18)"

                    Case dateDataType
                        oracleDDL &= "DATE"

                    Case longIntegerDataType
                        oracleDDL &= "NUMBER(10)"

                    Case singleDataType
                        oracleDDL &= "NUMBER(10,4)"

                    Case byteDataType
                        oracleDDL &= "NUMBER(3)"

                    Case currencyDataType
                        oracleDDL &= "NUMBER(19,2)"

                    Case booleanDataType
                        oracleDDL &= "NUMBER(3)"

                    Case Else
                        Stop
                End Select

                If rowColumns.Item("IS_NULLABLE") = False Then
                    oracleDDL &= " NOT NULL"
                End If
                oracleDDL &= vbCr
            Next

            oracleDDL = "Create Table " & vOracleTableName & " (" & Mid(oracleDDL, 3) & " )"

        Catch ex As Exception
            Errors.Add(ex.Message)
            Return String.Empty
        End Try

        Return oracleDDL

    End Function

#End Region

End Class
