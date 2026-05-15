Imports Oracle.ManagedDataAccess.Client

Public Class ASCTABLE
    Public TABLE_NAME As String
    Public TABLE_DDL As String

    ' This is how you needed to do this prior to Generics
    '    Public COLUMNS As ArrayList ' (could not say "of COLUMN")
    '    Public INDEXES As ArrayList ' (could not say "of INDEX")

    Public Columns As New List(Of COLUMN)
    Public Indexes As New List(Of INDEX)
    Public PrimaryKey As New List(Of String)

    ' the 2 variables XColumns and XIndexes, 
    ' and the classes COLUMN, INDEX, colCOLUMNS, and colINDEXES (below)
    ' are preserved in this class to show how to do it before Generics
    Dim XColumns As New colCOLUMNS
    Dim XIndexes As New colINDEXES

    ''' <summary>
    ''' Will retrieve the Table Definition of TABLE_NAME from Oracle Metadata
    ''' 
    ''' Assumes that oraCon is a live connection to Oracle
    ''' </summary>
    ''' <remarks></remarks>

    Sub Get_Table_Definition(ByVal oraCon As OracleConnection)

        Dim COLUMN_NAME As String
        Dim DATA_TYPE As String
        Dim DATA_LENGTH As Integer
        Dim DATA_PRECISION As Integer
        Dim DATA_SCALE As Integer

        ' Determine Primary Key Columns

        Dim PrimaryKeyColumns As String = ""
        Dim pkyINDEX_NAME As String = ""

        Dim sql As String

        sql = "SELECT USER_IND_COLUMNS.INDEX_NAME, USER_IND_COLUMNS.COLUMN_NAME "
        sql = sql & " FROM USER_IND_COLUMNS, USER_CONSTRAINTS "
        sql = sql & " WHERE(USER_IND_COLUMNS.INDEX_NAME = USER_CONSTRAINTS.CONSTRAINT_NAME)"
        sql = sql & " AND USER_CONSTRAINTS.TABLE_NAME = '" & TABLE_NAME & "' "
        sql = sql & " AND USER_CONSTRAINTS.CONSTRAINT_TYPE = 'P'"
        sql = sql & " ORDER BY COLUMN_POSITION"

        Dim oraCmd As New OracleCommand("", oraCon)
        oraCmd.CommandText = sql

        If oraCon.State = ConnectionState.Closed Then
            oraCon.Open()
        End If
        With oraCmd.ExecuteReader()
            While .Read()
                COLUMN_NAME = .Item("COLUMN_NAME").ToString
                pkyINDEX_NAME = .Item("INDEX_NAME").ToString
                PrimaryKeyColumns = PrimaryKeyColumns & "," & COLUMN_NAME
                PrimaryKey.Add(COLUMN_NAME)
            End While
            .Close()
            .Dispose()
        End With
        PrimaryKeyColumns = Mid$(PrimaryKeyColumns, 2)

        ' Construct the DDL required to create the table

        TABLE_DDL = ""
        'Dim columns As New colCOLUMNS

        sql = "SELECT COLUMN_NAME, DATA_TYPE, DATA_LENGTH"
        sql = sql & ", DATA_PRECISION, DATA_SCALE "
        sql = sql & " from USER_TAB_COLUMNS "
        sql = sql & " where TABLE_NAME = '" & TABLE_NAME & "'"
        sql = sql & " ORDER BY COLUMN_ID"
        oraCmd.CommandText = sql
        With oraCmd.ExecuteReader()
            While .Read()
                COLUMN_NAME = .Item("COLUMN_NAME")
                TABLE_DDL = TABLE_DDL & ", " & COLUMN_NAME
                DATA_TYPE = .Item("DATA_TYPE")
                DATA_LENGTH = .Item("DATA_LENGTH")
                DATA_PRECISION = Val(.Item("DATA_PRECISION") & "")
                DATA_SCALE = Val(.Item("DATA_SCALE") & "")

                Dim oraCOLUMN As New COLUMN
                oraCOLUMN.COLUMN_NAME = COLUMN_NAME
                oraCOLUMN.DATA_TYPE = DATA_TYPE
                oraCOLUMN.DATA_LENGTH = DATA_LENGTH
                oraCOLUMN.DATA_PRECISION = DATA_PRECISION
                oraCOLUMN.DATA_SCALE = DATA_SCALE
                Call Columns.Add(oraCOLUMN)

                Select Case DATA_TYPE
                    Case "VARCHAR2", "CHAR"
                        TABLE_DDL = TABLE_DDL & " VARCHAR (" & DATA_LENGTH & ")"
                    Case "DATE"
                        TABLE_DDL = TABLE_DDL & " DATETIME "
                    Case "NUMBER"
                        If DATA_SCALE = 0 Then
                            DATA_SCALE = 0
                        End If
                        If DATA_PRECISION = 0 Then
                            DATA_PRECISION = 3
                        End If
                        TABLE_DDL = TABLE_DDL & " DECIMAL (" & DATA_PRECISION
                        If DATA_SCALE <> 0 Then
                            TABLE_DDL = TABLE_DDL & "," & DATA_SCALE
                        End If
                        TABLE_DDL = TABLE_DDL & ")"
                    Case Else
                        Stop
                End Select
            End While
            .Close()
            .Dispose()
        End With

        TABLE_DDL = "Create Table " & TABLE_NAME & " (" & Mid$(TABLE_DDL, 2)
        If PrimaryKeyColumns <> "" Then
            TABLE_DDL = TABLE_DDL & ", PRIMARY KEY (" & PrimaryKeyColumns & ")"
        End If
        TABLE_DDL = TABLE_DDL & ")"

        ' Construct DDL for each Index and Create

        Dim INDEX_NAME As String = ""
        Dim INDEX_DDL As String = ""

        sql = "Select INDEX_NAME from USER_INDEXES"
        sql = sql & " where TABLE_NAME = '" & TABLE_NAME & "'"
        sql = sql & " and INDEX_NAME <>  '" & pkyINDEX_NAME & "'"
        oraCmd.CommandText = sql
        With oraCmd.ExecuteReader()
            While .Read()
                INDEX_NAME = .Item("INDEX_NAME")
                INDEX_DDL = ""

                Dim oraINDEX As New INDEX
                oraINDEX.INDEX_NAME = INDEX_NAME
                oraINDEX.INDEX_DDL = INDEX_DDL
                Call Indexes.Add(oraINDEX)

                sql = "SELECT USER_TAB_COLUMNS.COLUMN_NAME, USER_TAB_COLUMNS.DATA_TYPE, USER_TAB_COLUMNS.DATA_LENGTH"
                sql = sql & ", USER_TAB_COLUMNS.DATA_PRECISION, USER_TAB_COLUMNS.DATA_SCALE "
                sql = sql & " from USER_TAB_COLUMNS, USER_IND_COLUMNS "
                sql = sql & " where USER_IND_COLUMNS.INDEX_NAME = '" & INDEX_NAME & "'"
                sql = sql & "   and USER_IND_COLUMNS.TABLE_NAME = USER_TAB_COLUMNS.TABLE_NAME"
                sql = sql & "   and USER_IND_COLUMNS.COLUMN_NAME = USER_TAB_COLUMNS.COLUMN_NAME"
                sql = sql & " ORDER BY USER_IND_COLUMNS.COLUMN_POSITION"
                oraCmd.CommandText = sql
                With oraCmd.ExecuteReader()
                    While .Read()
                        COLUMN_NAME = .Item("COLUMN_NAME")
                        INDEX_DDL = INDEX_DDL & "," & COLUMN_NAME

                        DATA_TYPE = .Item("DATA_TYPE")
                        DATA_LENGTH = .Item("DATA_LENGTH")
                        DATA_PRECISION = Val(.Item("DATA_PRECISION") & "")
                        DATA_SCALE = Val(.Item("DATA_SCALE") & "")

                        Dim oraCOLUMN As New COLUMN
                        oraCOLUMN.COLUMN_NAME = COLUMN_NAME
                        oraCOLUMN.DATA_TYPE = DATA_TYPE
                        oraCOLUMN.DATA_LENGTH = DATA_LENGTH
                        oraCOLUMN.DATA_PRECISION = DATA_PRECISION
                        oraCOLUMN.DATA_SCALE = DATA_SCALE
                        Call oraINDEX.Columns.Add(oraCOLUMN)

                    End While
                    .Close()
                    .Dispose()
                End With
               

                INDEX_DDL = "Create Index " & INDEX_NAME _
                    & " on " & TABLE_NAME _
                    & " (" & Mid$(INDEX_DDL, 2) & ")"

                oraINDEX.INDEX_DDL = INDEX_DDL
            End While
            .Close()
            .Dispose()
        End With
        
    End Sub
    ''' <summary>
    ''' Returns a comma-separated Primary Key Column List
    '''  prepended by " ORDER BY "
    ''' </summary>
    ''' <param name="prepend_ORDER_BY"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPrimaryKey(ByVal prepend_ORDER_BY As Boolean) As String
        Dim zz As String
        zz = GetPrimaryKey()

        If zz <> "" And prepend_ORDER_BY Then
            zz = " ORDER BY " & zz
        End If
        GetPrimaryKey = zz
    End Function
    ''' <summary>
    ''' Returns a comma-separated Primary Key Column List
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPrimaryKey() As String
        Dim zz As String = ""

        GetPrimaryKey = ""
        If PrimaryKey.Count > 0 Then
            For Each z As String In PrimaryKey
                zz = zz & "," & z
            Next
            zz = Mid(zz, 2)
            GetPrimaryKey = zz
        End If

    End Function
    ''' <summary>
    ''' Will create a class which includes various metadata for the TABLE_NAME specified
    ''' using the (live) Oracle Connection specified in oraCon
    ''' </summary>
    ''' <param name="TABLE_NAME"></param>
    ''' <param name="oraCon"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal TABLE_NAME As String, ByRef oraCon As OracleConnection)
        Me.TABLE_NAME = TABLE_NAME
        Call Get_Table_Definition(oraCon)
    End Sub
    ''' <summary>
    ''' Will create a class which includes various metadata for the TABLE_NAME specified
    ''' using an assumed (and assumed to be live) Oracle Connection oraCon
    ''' </summary>
    ''' <param name="TABLE_NAME"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal TABLE_NAME As String)
        Me.TABLE_NAME = TABLE_NAME
        Call Get_Table_Definition(ASCMAIN1.oraCon)
    End Sub
    ''' <summary>
    ''' Will create a SQL Server Table with the same definition 
    '''  as the Oracle Table in the ASCTABLE object, 
    '''  including Primary Key Definition and Indexes,
    '''  but using the table name provided (TABLE_NAME)
    ''' </summary>
    ''' <param name="TABLE_NAME"></param>
    ''' <remarks></remarks>
    Public Sub Create_SQL_Table(ByVal TABLE_NAME As String)
        ' Declare working variables

        Dim sql As String
        Dim i As Integer

        Dim sqlCmd As OracleCommand = ASCMAIN1.oraCon.CreateCommand

        ' Drop the table if it already exists

        sql = "Drop table " & TABLE_NAME
        sqlCmd.CommandText = sql
        Try
            i = sqlCmd.ExecuteNonQuery()
        Catch
        End Try

        ' Create the TABLE

        sqlCmd.CommandText = TABLE_DDL
        Try
            i = sqlCmd.ExecuteNonQuery()
        Catch
        End Try

        ' Create the INDEXES

        For Each X As INDEX In Indexes
            sqlCmd.CommandText = X.INDEX_DDL
            Try
                i = sqlCmd.ExecuteNonQuery()
            Catch
            End Try
        Next

    End Sub
    ''' <summary>
    ''' Will create a SQL Server Table with the same definition 
    '''  as the Oracle Table in the ASCTABLE object, 
    '''  including Primary Key Definition and Indexes,
    '''  and using the same name (TABLE_NAME)
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Create_SQL_Table()
        Call Create_SQL_Table(Me.TABLE_NAME)
    End Sub

End Class

Public Class colCOLUMNS
    Inherits ArrayList
    Public Property Itemx(ByVal I As Integer) As COLUMN
        Get
            Return MyBase.Item(I)
        End Get

        Set(ByVal value As COLUMN)
            MyBase.Item(I) = value
        End Set
    End Property
End Class

Public Class colINDEXES
    Inherits ArrayList
    Public Property Itemx(ByVal I As Integer) As INDEX
        Get
            Return MyBase.Item(I)
        End Get

        Set(ByVal value As INDEX)
            MyBase.Item(I) = value
        End Set
    End Property
End Class

Public Class COLUMN
    'Inherits System.Collections.CollectionBase
    Public COLUMN_NAME As String
    Public DATA_TYPE As String
    Public DATA_LENGTH As Integer
    Public DATA_PRECISION As Integer
    Public DATA_SCALE As Integer
End Class

Public Class INDEX
    Public INDEX_NAME As String
    Public INDEX_DDL As String
    Public Columns As New colCOLUMNS
End Class
