Imports System.Data.Common
Imports Oracle.ManagedDataAccess.Client

Public Class ASCDATA1
    Implements IDisposable

    Public DataTableSchema As DataTable
    Public DTS As New SortedList(Of String, DataTable)
    Public ASCMAIN1 As ASCMAIN1

    ''' <summary>
    ''' Returns a DataTable representing the result set for the sql provided
    '''  using the (assumed to be open) db Connection 
    ''' </summary>
    ''' <param name="sql"></param>
    ''' <param name="TABLE_NAME"></param>
    ''' <param name="Key_Field_Count"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDataTable( _
    Optional ByVal sql As String = "", _
    Optional ByVal TABLE_NAME As String = "", _
    Optional ByVal Key_Field_Count As Integer = -1, _
    Optional ByVal Include_Data As Boolean = True, _
    Optional ByVal NumberOfKeysUsedToSelect As Integer = -1, _
    Optional ByVal custom_parameters As String = "", _
    Optional ByVal custom_parameter_values() As Object = Nothing) _
    As DataTable

        Dim tbl As New DataTable
        With GetDataAdapter(tbl, TABLE_NAME, sql, False, Key_Field_Count, _
                            Include_Data, NumberOfKeysUsedToSelect, , , _
                            custom_parameters, custom_parameter_values)
            .Dispose()
        End With
        Return tbl
    End Function

    Public Function GetDataTable( _
    ByVal sql As String, _
    ByVal TABLE_NAME As String, _
    ByVal params As String, _
    ByVal ParamArray PARMs() As Object) As DataTable
        Dim dt As New DataTable(TABLE_NAME)
        With GetDataAdapter(dt, TABLE_NAME, SQL_Server_Conversions(sql), params, PARMs)
            .Fill(dt)
            .Dispose()
        End With
        Return dt
    End Function

    Public Function GetDataRow( _
    ByVal sql As String, _
    ByVal params As String, _
    ByVal PARMs() As Object) As DataRow
        Dim tbl As DataTable = GetDataTable(sql, "DataRowTemp", params, PARMs)
        If tbl.Rows.Count > 0 Then
            Return tbl.Rows(0)
        Else
            Return Nothing
        End If
    End Function

    Public Function GetDataAdapter( _
    ByRef tbl As DataTable, _
    ByVal TABLE_NAME As String, _
    ByVal sql As String, _
    ByVal params As String, _
    ByVal ParamArray PARMs() As Object) As OracleDataAdapter
        Return GetDataAdapter(tbl, TABLE_NAME, sql, , , False, 0, , , params, PARMs)
    End Function

    Public Function GetDataAdapter( _
    ByRef tbl As DataTable, _
    ByVal TABLE_NAME As String, _
    ByVal sql_custom As String, _
    Optional ByVal for_update As Boolean = False, _
    Optional ByVal Key_Field_Count As Integer = -1, _
    Optional ByVal Include_Data As Boolean = True, _
    Optional ByVal NumberOfKeysUsedToSelect As Integer = -1, _
    Optional ByVal Update_COLUMN_NAMEs As String = "", _
    Optional ByVal SCHEMA As String = "", _
    Optional ByVal custom_parameters As String = "", _
    Optional ByVal custom_parameter_values() As Object = Nothing _
    ) As OracleDataAdapter

        Dim XDT As Date = Now

        Dim SCHEMA_dot As String = ""
        If SCHEMA <> "" Then
            SCHEMA_dot = SCHEMA & "."
        End If

        'If ASCMAIN1.oraCon.ConnectionString = "" Then
        '    ASCMAIN1.oraCon.ConnectionString = "Data Source=" & "" & ";User ID=" & "ODG" & ";Password=" & "ODG"
        '    ASCMAIN1.oraCon.Open()
        'End If

        Dim sql As String = ""

        If sql_custom = "*" Or sql_custom = "**" Then
            If (sql_custom = "**" And Not for_update) Then
                'Q: WHY WAS THE And Not clause necessary?  SOTSREP1 use of ARTCUST1 has an issue with that
                'A: because if you append columns, the code below has no way of knowing which columns are appeneded vs those that belong to the table when constructing the insert and update commands
                '   so the deal is (for now) that you must always use TABLE_NAME.*,APPENDED_COL1,APPENDED_COL2 when setting up a Data Adapter for update
                sql = ASCMAIN1.sql
            Else
                sql = "Select * from " & SCHEMA_dot & TABLE_NAME
            End If
        Else
            If sql_custom = "" Then
                sql = ASCMAIN1.sql
            Else
                sql = sql_custom
            End If
        End If

        Dim includes_where_clause As Boolean = (sql_custom = "**" And for_update) Or (Not (sql_custom = "*") And (InStr(sql.ToUpper, " WHERE ") <> 0))

        Dim Ada As New OracleDataAdapter
        Ada.SelectCommand = ASCMAIN1.oraCon.CreateCommand
        'Ada.SelectCommand.Transaction = ASCMAIN1.T
        sql = SQL_Server_Conversions(sql)
        Ada.SelectCommand.CommandText = sql

        ASCMAIN1.sql = SQL_Server_Conversions(ASCMAIN1.sql)

        If tbl Is Nothing Then
            tbl = New DataTable
        End If

        If TABLE_NAME <> "" Then
            If tbl.TableName Like "Table*" Or tbl.TableName = "" Then
                If tbl.DataSet IsNot Nothing AndAlso tbl.DataSet.Tables.Contains(TABLE_NAME) Then
                    tbl.DataSet.Tables.Remove(TABLE_NAME)
                End If
                tbl.TableName = TABLE_NAME
            End If
        End If

        'For Each KK As String In ASCMAIN1.FilledSchemas.Keys
        '    Debug.Print(KK)
        'Next

        Dim sqlC As String = Mid(sql.ToUpper.Replace(" ", "").Replace(vbCrLf, ""), 1, 24)  ' "Select * from ASTMTSK2  where"
        If sqlC.Length >= 19 And sqlC.StartsWith("SELECT*FROM") And Not sqlC.StartsWith("SELECT*FROMASW") Then
            If sqlC.Length = 19 Or (sqlC.Length = 24 And sqlC.EndsWith("WHERE")) Then
                sqlC = sqlC.Substring(0, 19)
            Else
                sqlC = ""
            End If
        Else
            sqlC = ""
        End If

        If ASCMAIN1.FilledSchemas.ContainsKey(sqlC) Then
            Dim TableName As String = tbl.TableName
            Dim t As New DataTable
            t = ASCMAIN1.FilledSchemas(sqlC).Clone
            For Each dc As DataColumn In t.Columns
                Dim DC2 As New DataColumn(dc.ColumnName, dc.DataType)
                DC2.MaxLength = dc.MaxLength
                tbl.Columns.Add(DC2)
                'tbl.Columns.Add(dc.ColumnName, dc.DataType)
            Next
            If t.PrimaryKey.Length <> 0 Then
                Dim KCs() As DataColumn
                ReDim KCs(t.PrimaryKey.Length - 1)
                Dim KCL As Int16 = 0
                For Each KC As DataColumn In t.PrimaryKey
                    KCs(KCL) = tbl.Columns(KC.ColumnName)
                    KCL += 1
                Next
                tbl.PrimaryKey = KCs
            End If
            'tbl = t
            'tbl.TableName = TableName
        Else
            Fill_Schema(Ada, tbl)
            If 1 <> 1 Then
                Return Nothing
            End If
            If sqlC <> "" Then
                ASCMAIN1.FilledSchemas.Add(sqlC, tbl.Clone)
            End If
        End If

        '        Dim KeyFields_from_Table As Integer = tbl.PrimaryKey.Length

        Dim OBJPFX As String = ASCMAIN1.DBS_OBJPFX
        Dim OBJSFX As String = ASCMAIN1.DBS_OBJSFX
        OBJPFX = ""
        OBJSFX = ""

        If Key_Field_Count = -1 Then
            'If Key_Field_Count <> -1 Then
            '    If Key_Field_Count <> 0 Then
            '        tbl.PrimaryKey = Nothing
            '        Dim DC As DataColumn()
            '        ReDim DC(Key_Field_Count - 1)
            '        For i As Int16 = 1 To Key_Field_Count
            '            DC(i - 1) = tbl.Columns(i - 1)
            '        Next
            '        tbl.PrimaryKey = DC
            '    End If
            'Else
            '    Key_Field_Count = tbl.PrimaryKey.Length ' KeyFields_from_Table
            'End If
            Key_Field_Count = tbl.PrimaryKey.Length

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
                    sql_key = sql_key & " and (" & OBJPFX & TABLE_NAME & "." & tbl.Columns(i).ColumnName & OBJSFX & "=" & ASCMAIN1.DBS_PARAMETER & tbl.Columns(i).ColumnName & ")"
                    sql_cols = sql_cols & "," & OBJPFX & tbl.Columns(i).ColumnName & OBJSFX
                    sql_parms = sql_parms & "," & ASCMAIN1.DBS_PARAMETER & tbl.Columns(i).ColumnName
                Next i
                sql_key = " where " & Mid(sql_key, 6)
            End If

            Dim sql_data As String = ""
            For i As Integer = tbl.PrimaryKey.Length To tbl.Columns.Count - 1
                sql_data = sql_data & "," & OBJPFX & TABLE_NAME & "." & tbl.Columns(i).ColumnName & OBJSFX & "=" & ASCMAIN1.DBS_PARAMETER & tbl.Columns(i).ColumnName
                sql_cols = sql_cols & "," & OBJPFX & Chr(34) & tbl.Columns(i).ColumnName & Chr(34) & OBJSFX
                sql_parms = sql_parms & "," & ASCMAIN1.DBS_PARAMETER & tbl.Columns(i).ColumnName
            Next i
            sql_data = " Set " & Mid(sql_data, 2)

            sql_cols = Mid(sql_cols, 2)
            sql_parms = Mid(sql_parms, 2)

            sql = "Insert into " & SCHEMA_dot & OBJPFX & TABLE_NAME & OBJSFX & " (" & sql_cols & ") values (" & sql_parms & ")"
            Ada.InsertCommand = New OracleCommand(sql, ASCMAIN1.oraCon)
            Ada.InsertCommand.BindByName = True

            Dim sql_data2 As String = ""
            If Update_COLUMN_NAMEs <> "" Then
                For Each Update_COLUMN_NAME As String In Split(Update_COLUMN_NAMEs, ",")
                    UCN.Add(UCase(Trim(Update_COLUMN_NAME)))
                    sql_data2 = sql_data2 & "," & OBJPFX & TABLE_NAME & "." & Trim(Update_COLUMN_NAME) & OBJSFX & "=" & ASCMAIN1.DBS_PARAMETER & Trim(Update_COLUMN_NAME)
                Next
                sql_data2 = " Set " & Mid(sql_data2, 2)
            End If
            If UCN.Count = 0 Then
                sql_data2 = sql_data
            End If
            'If TABLE_NAME = "APTINVH1" Then Stop

            If sql_key <> "" Then
                sql = "Update " & SCHEMA_dot & OBJPFX & TABLE_NAME & OBJSFX & sql_data2 & Replace(sql_key, "=" & ASCMAIN1.DBS_PARAMETER, "=" & ASCMAIN1.DBS_PARAMETER & "U_")
                Ada.UpdateCommand = New OracleCommand(sql, ASCMAIN1.oraCon)
                Ada.UpdateCommand.BindByName = True

                sql = "Delete from " & SCHEMA_dot & OBJPFX & TABLE_NAME & OBJSFX & Replace(sql_key, "=" & ASCMAIN1.DBS_PARAMETER, "=" & ASCMAIN1.DBS_PARAMETER & "D_")
                Ada.DeleteCommand = New OracleCommand(sql, ASCMAIN1.oraCon)
                Ada.DeleteCommand.BindByName = True
            End If
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

                Dim p_I_CURR As New OracleParameter
                p_I_CURR.ParameterName = COLUMN_NAME
                p_I_CURR.SourceColumn = COLUMN_NAME
                p_I_CURR.SourceVersion = DataRowVersion.Current

                Dim p_D_ORIG As New OracleParameter
                p_D_ORIG.ParameterName = "D_" & COLUMN_NAME
                p_D_ORIG.SourceColumn = COLUMN_NAME
                p_D_ORIG.SourceVersion = DataRowVersion.Original

                Dim p_U_CURR As New OracleParameter
                p_U_CURR.ParameterName = COLUMN_NAME
                p_U_CURR.SourceColumn = COLUMN_NAME
                p_U_CURR.SourceVersion = DataRowVersion.Current

                Dim p_U_ORIG As New OracleParameter
                p_U_ORIG.ParameterName = "U_" & COLUMN_NAME
                p_U_ORIG.SourceColumn = COLUMN_NAME
                p_U_ORIG.SourceVersion = DataRowVersion.Original

                Dim p_S_ORIG As New OracleParameter
                p_S_ORIG.ParameterName = "S_" & COLUMN_NAME
                p_S_ORIG.SourceColumn = COLUMN_NAME
                p_S_ORIG.SourceVersion = DataRowVersion.Original
                p_S_ORIG.Value = DBNull.Value

                Dim DBT As DbType = Nothing

                Select Case DATA_DTYPE

                    Case "System.String"
                        DBT = DbType.String

                        p_I_CURR.Size = DATA_LENGTH
                        p_U_CURR.Size = DATA_LENGTH
                        p_U_ORIG.Size = DATA_LENGTH
                        p_D_ORIG.Size = DATA_LENGTH
                        p_S_ORIG.Size = DATA_LENGTH

                    Case "System.DateTime"
                        DBT = DbType.Date
                    Case "System.Double"
                        DBT = DbType.Double
                    Case "System.Single"
                        DBT = DbType.Single
                    Case "System.Decimal"
                        DBT = DbType.Decimal
                    Case "System.Int16"
                        DBT = DbType.Int16
                    Case "System.Int32"
                        DBT = DbType.Int32
                    Case "System.Int64"
                        DBT = DbType.Int64
                    Case Else
                        Throw New Exception("Unsupported Data Type")
                End Select

                p_I_CURR.DbType = DBT
                p_U_CURR.DbType = DBT
                p_U_ORIG.DbType = DBT
                p_D_ORIG.DbType = DBT
                p_S_ORIG.DbType = DBT

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
                    (NumberOfKeysUsedToSelect = -1 _
                    Or NumberOfKeysUsedToSelect >= i + 1) Then
                        'Ada.DeleteCommand.Parameters.Add(p_D_ORIG)
                        'Ada.UpdateCommand.Parameters.Add(p_U_ORIG)
                        Ada.SelectCommand.Parameters.Add(p_S_ORIG)
                    End If
                Else
                    If for_update Then
                        If sql_key <> "" Then
                            If UCN.Count = 0 OrElse UCN.Contains(COLUMN_NAME) Then
                                Ada.UpdateCommand.Parameters.Add(p_U_CURR)
                            End If
                        End If
                    End If
                End If

                If Not for_update And i >= Key_Field_Count - 1 Then
                    Exit For ' optimize
                End If
            Next i
        End If

        sql_key = ""

        'Dim XDT As Date = Now

        If sql_custom = "**" And for_update Then
            Ada.SelectCommand.CommandText = ASCMAIN1.sql

            Dim sqlC2 As String = Mid(ASCMAIN1.sql.ToUpper.Replace(" ", "").Replace(vbCrLf, ""), 1, 24)
            If sqlC2 = sqlC Or sqlC2 = sqlC & "WHERE" Then
            Else
                Fill_Schema(Ada, tbl)
            End If

        End If

        If custom_parameters.Length <> 0 Then
            Create_Parameters(Ada.SelectCommand, custom_parameters, custom_parameter_values)
        End If

        If Include_Data Then
            Try
                Ada.Fill(tbl)
            Catch ex As Exception
                Throw New Exception(String.Format("Exception in GetDataAdapter: {0}", ex.Message))
            End Try
        End If

        If ASCMAIN1.tblASTSQLX1 IsNot Nothing Then
            Dim commandText As String = Ada.SelectCommand.CommandText
            If commandText.Length > 4000 Then
                commandText = commandText.Substring(0, 4000).Trim
            End If

            Dim SELECTION_NO As Int32 = -1
            Dim RE_XNO As Int32 = -1

            If ASCMAIN1.ActiveForm IsNot Nothing Then
                SELECTION_NO = ASCMAIN1.ActiveForm.SELECTION_NO
                RE_XNO = ASCMAIN1.ActiveForm.RE_XNO
            End If

            If ASCMAIN1.tblASTSQLX1.Columns.Count <> 0 Then
                ASCMAIN1.tblASTSQLX1.Rows.Add(New Object() _
                {ASCMAIN1.SESSION_NO, SELECTION_NO, RE_XNO, _
                 XDT, Now.Subtract(XDT).Ticks / 10000000, commandText})
            End If
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

                    sql_key = sql_key & " and (" & OBJPFX & IIf(for_update, TABLE_NAME & ".", "") & tbl.Columns(i).ColumnName & OBJSFX & "=" & ASCMAIN1.DBS_PARAMETER & tbl.Columns(i).ColumnName & ")"
                Next i
                If Not includes_where_clause Then
                    sql_key = " where " & Mid(sql_key, 6)
                End If
                Ada.SelectCommand.BindByName = True
                Ada.SelectCommand.CommandText &= Replace(sql_key, "=" & ASCMAIN1.DBS_PARAMETER, "=" & ASCMAIN1.DBS_PARAMETER & "S_")
            End If
        End If
        Return Ada
    End Function

    ''' <summary>
    ''' Executes the Store Procedure in using the ASCMAIN1.oraSP Object (which originates from the oraCon Connection object).
    ''' If you have begun a Transacton with oraCon, then this statement will execute within the context of that transaction.
    ''' </summary>
    ''' <param name="Stored_Procedure_Name"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteSP(ByVal Stored_Procedure_Name As String) As Integer
        Try
            ASCMAIN1.oraSP.CommandText = Stored_Procedure_Name
            Return ASCMAIN1.oraSP.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception(String.Format("Exception in ExecuteSP: {0}", ex.Message))
        End Try
    End Function

    Public Function ExecuteSP( _
    ByVal Stored_Procedure_Name As String, _
    ByVal custom_parameters As String, _
    ByVal PARMs() As Object, _
    ByVal PARM_NAMEs() As String) As Integer

        Try
            Dim sp As New OracleCommand(Stored_Procedure_Name, ASCMAIN1.oraCon)
            sp.CommandType = CommandType.StoredProcedure
            'sp.Transaction = ASCMAIN1.T
            If custom_parameters <> "" Then
                Call Create_Parameters(sp, custom_parameters)
                For i As Integer = 1 To Len(custom_parameters)
                    With sp.Parameters("PARM" & CStr(i))
                        .ParameterName = PARM_NAMEs(i - 1)
                        If PARMs Is Nothing Then
                            .Value = System.DBNull.Value
                        Else
                            If PARMs.Length >= i Then
                                .Value = PARMs(i - 1)
                            Else
                                .Value = System.DBNull.Value
                            End If
                        End If
                    End With
                Next
            End If

            Try
                ExecuteSP = sp.ExecuteNonQuery()
            Catch ex As Exception
                ExecuteSP = 0
                Throw New Exception(String.Format("Exception in ExecuteSQL: {0}", ex.Message))
            End Try
        Catch ex As Exception
            ExecuteSP = 0
            Throw ex
        End Try

    End Function

    Public Function ExecuteSF(ByVal sfName As String, ByVal ParamArray PARMs() As Object)
        Using cmd As New OracleCommand(sfName, ASCMAIN1.oraCon)
            cmd.CommandType = CommandType.StoredProcedure
            Dim dummy As Object = Nothing
            cmd.Parameters.Add("returnValue", OracleDbType.Varchar2, 2000, dummy, ParameterDirection.ReturnValue)
            If PARMs IsNot Nothing Then
                CreateParametersByPosition(cmd, PARMs)
            End If
            cmd.ExecuteNonQuery()
            ExecuteSF = cmd.Parameters("returnValue").Value.ToString
        End Using
    End Function

    Private Shared Sub CreateParametersByPosition( _
                    ByRef cmd As OracleCommand, _
                    ByVal ParamArray PARMS() As Object)

        With cmd.Parameters
            For i As Integer = 1 To PARMS.Length
                Dim PP As New OracleParameter
                PP.Direction = ParameterDirection.Input
                If PARMS Is Nothing OrElse PARMS.Length < i Then
                    PP.Value = System.DBNull.Value
                Else
                    PP.Value = PARMS(i - 1)
                End If
                .Add(PP)
            Next
        End With
    End Sub

    Public Sub Create_Parameters( _
    ByRef cmd As OracleCommand, _
    ByVal custom_parameters As String, _
    ByVal PARMS() As Object, _
    ByRef ptbl As DataTable)

        With cmd.Parameters
            For i As Integer = 1 To Len(custom_parameters)

                Dim COLUMN_NAME As String = "PARM" & CStr(i)
                Dim P As New OracleParameter
                P.ParameterName = COLUMN_NAME
                P.Direction = ParameterDirection.Input

                Select Case Mid$(custom_parameters, i, 1)
                    Case "V"    ' Varchar2
                        P.OracleDbType = OracleDbType.Varchar2
                        If ptbl IsNot Nothing Then ptbl.Columns.Add(COLUMN_NAME, GetType(System.String))
                    Case "I"    ' Integer
                        P.OracleDbType = OracleDbType.Long
                        If ptbl IsNot Nothing Then ptbl.Columns.Add(COLUMN_NAME, GetType(System.Int32))
                    Case "N"    ' Number
                        P.OracleDbType = OracleDbType.Double
                        If ptbl IsNot Nothing Then ptbl.Columns.Add(COLUMN_NAME, GetType(System.Double))
                    Case "D"    ' Date
                        P.OracleDbType = OracleDbType.Date
                        If ptbl IsNot Nothing Then ptbl.Columns.Add(COLUMN_NAME, GetType(System.DateTime))
                End Select

                cmd.Parameters.Add(P)
                If PARMS Is Nothing OrElse PARMS.Length < i Then
                    P.Value = System.DBNull.Value
                Else
                    If P.OracleDbType = OracleDbType.Date And TypeOf (PARMS(i - 1)) Is Date Then
                        If PARMS(i - 1) & "" <> "" AndAlso PARMS(i - 1).ToString = "1/1/0001 12:00:00 AM" Then
                            P.Value = DBNull.Value
                        Else
                            P.Value = PARMS(i - 1)
                        End If
                        '  P.Value = PARMS(i - 1)
                        'If P.OracleDbType = OracleDbType.Date And TypeOf (PARMS(i - 1)) Is Date Then
                        '    Dim D As Date = PARMS(i - 1) ' NOTE THAT IF D HAS TIME, THAT THE NEXT LINE IS STRIPPING IT OUT
                        '    P.Value = Format(PARMS(i - 1), "dd-MMM-yyyy")
                        'Else
                        '    P.Value = PARMS(i - 1)
                        'End If
                    Else
                        P.Value = PARMS(i - 1)
                    End If
                End If
            Next
        End With
    End Sub

    Public Sub Create_Parameters( _
    ByRef cmd As OracleCommand, _
    ByVal custom_parameters As String, _
    Optional ByRef ptbl As DataTable = Nothing)

        Create_Parameters(cmd, custom_parameters, Nothing, ptbl)
    End Sub

    Public Sub Create_Parameters( _
    ByRef cmd As OracleCommand, _
    ByVal custom_parameters As String, _
    ByVal PARMS() As Object)

        Create_Parameters(cmd, custom_parameters, PARMS, Nothing)
    End Sub

    Public Function ExecuteSQL( _
    ByVal sql_to_Execute As String, _
    ByVal custom_parameters As String, _
    ByVal PARM As Object) As Integer
        ExecuteSQL(sql_to_Execute, custom_parameters, New Object() {PARM})
    End Function

    ''' <summary>
    ''' Executes the sql statement passed in using the ASCMAIN1.oraCmd Object (which originates from the oraCon Connection object).
    ''' If you have begun a Transacton with oraCon, then this statement will execute within the context of that transaction.
    ''' </summary>
    ''' <param name="sql_to_Execute"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteSQL( _
    ByVal sql_to_Execute As String, _
    ByVal custom_parameters As String, _
    ByVal PARMs() As Object, _
    Optional ByVal build_or_execute As ASCMAIN1.ExecuteSQL_types = ASCMAIN1.ExecuteSQL_types.Build_and_Execute, _
    Optional ByVal command_name As String = "") As Integer

        If build_or_execute = ASCMAIN1.ExecuteSQL_types.Execute_Only Then
            Using cmd As OracleCommand = ASCMAIN1.ActiveForm.CMDs(command_name)
                Try
                    cmd.BindByName = True 'already set above?

                    Dim XDT As Date = Now

                    ExecuteSQL = cmd.ExecuteNonQuery

                    If ASCMAIN1.tblASTSQLX1 IsNot Nothing Then
                        Dim commandText As String = cmd.CommandText
                        If commandText.Length > 4000 Then
                            commandText = commandText.Substring(0, 4000).Trim
                        End If

                        Dim SELECTION_NO As Int32 = -1
                        Dim RE_XNO As Int32 = -1

                        If ASCMAIN1.ActiveForm IsNot Nothing Then
                            SELECTION_NO = ASCMAIN1.ActiveForm.SELECTION_NO
                            RE_XNO = ASCMAIN1.ActiveForm.RE_XNO
                        End If

                        ASCMAIN1.tblASTSQLX1.Rows.Add(New Object() _
                        {ASCMAIN1.SESSION_NO, SELECTION_NO, RE_XNO, _
                         XDT, Now.Subtract(XDT).Ticks / 10000000, commandText})
                    End If

                Catch ex As Exception
                    Throw New Exception(String.Format("Exception in ExecuteSQL: {0}", ex.Message))
                End Try
            End Using

            Exit Function
        End If

        Using cmd As New OracleCommand(sql_to_Execute, ASCMAIN1.oraCon)
            cmd.BindByName = True
            'cmd.Transaction = ASCMAIN1.T
            If custom_parameters <> "" Then
                Create_Parameters(cmd, custom_parameters, PARMs)
            End If

            cmd.CommandText = SQL_Server_Conversions(sql_to_Execute)

            If build_or_execute = ASCMAIN1.ExecuteSQL_types.Build_Only Then
                ASCMAIN1.ActiveForm.CMDs.Add(command_name, cmd)
                Exit Function
            End If

            Try
                cmd.BindByName = True 'already set above?

                Dim XDT As Date = Now

                ExecuteSQL = cmd.ExecuteNonQuery

                If ASCMAIN1.tblASTSQLX1 IsNot Nothing Then
                    Dim commandText As String = cmd.CommandText
                    If commandText.Length > 4000 Then
                        commandText = commandText.Substring(0, 4000).Trim
                    End If

                    Dim SELECTION_NO As Int32 = -1
                    Dim RE_XNO As Int32 = -1

                    If ASCMAIN1.ActiveForm IsNot Nothing Then
                        SELECTION_NO = ASCMAIN1.ActiveForm.SELECTION_NO
                        RE_XNO = ASCMAIN1.ActiveForm.RE_XNO
                    End If

                    ASCMAIN1.tblASTSQLX1.Rows.Add(New Object() _
                    {ASCMAIN1.SESSION_NO, SELECTION_NO, RE_XNO, _
                     XDT, Now.Subtract(XDT).Ticks / 10000000, commandText})
                End If

            Catch ex As Exception
                Throw New Exception(String.Format("Exception in ExecuteSQL: {0}", ex.Message))
            End Try
        End Using
    End Function

    ''' <summary>
    ''' Executes the sql statement passed in using the ASCMAIN1.oraCmd Object (which originates from the oraCon Connection object).
    ''' If you have begun a Transacton with oraCon, then this statement will execute within the context of that transaction.
    ''' </summary>
    ''' <param name="sql_to_Execute"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteSQL( _
    Optional ByVal sql_to_Execute As String = "", _
    Optional ByVal ok_to_fail As Boolean = False) As Integer
        If sql_to_Execute = "" Then
            sql_to_Execute = ASCMAIN1.sql
        End If

        Try

            Dim XDT As Date = Now

            sql_to_Execute = SQL_Server_Conversions(sql_to_Execute)

            ASCMAIN1.oraCmd.CommandText = sql_to_Execute
            'ASCMAIN1.oraCmd.Transaction = ASCMAIN1.T
            Dim icmd As Int32 = ASCMAIN1.oraCmd.ExecuteNonQuery()

            If ASCMAIN1.tblASTSQLX1 IsNot Nothing Then
                Dim commandText As String = sql_to_Execute
                If commandText.Length > 4000 Then
                    commandText = commandText.Substring(0, 4000).Trim
                End If

                Dim SELECTION_NO As Int32 = -1
                Dim RE_XNO As Int32 = -1

                If ASCMAIN1.ActiveForm IsNot Nothing Then
                    SELECTION_NO = ASCMAIN1.ActiveForm.SELECTION_NO
                    RE_XNO = ASCMAIN1.ActiveForm.RE_XNO
                End If

                ASCMAIN1.tblASTSQLX1.Rows.Add(New Object() _
                 {ASCMAIN1.SESSION_NO, SELECTION_NO, RE_XNO, _
                  XDT, Now.Subtract(XDT).Ticks / 10000000, commandText})
            End If

            Return icmd
        Catch ex As Exception
            Dim oex As OracleException
            oex = DirectCast(ex, OracleException)
            If Not ok_to_fail Then
                'MsgBox(sql_to_Execute & Environment.NewLine & " caused the following error " & Environment.NewLine & ex.Message, MsgBoxStyle.OkOnly, "SQL Error")

                Throw New Exception(String.Format("Exception in ExecuteSQL: {0}", ex.Message))

            End If
        End Try
    End Function

    Public Function GetDataRow( _
    ByVal sql_to_Execute As String, _
    ByVal custom_parameters As String, _
    ByVal custom_parameter_value As Object) As DataRow
        Return GetDataRow(sql_to_Execute, custom_parameters, New Object() {custom_parameter_value})
    End Function

    ''' <summary>
    ''' Get a datarow using the SQL passed in the first parameter, and return an emtpy row if missing depending on the second parameter
    ''' </summary>
    ''' <param name="sql_to_Execute">Defaults to ASCMAIN1.SQL if omitted</param>
    ''' <param name="Return_Empty_Row_if_Missing">Defaults to False if omitted</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDataRow( _
    Optional ByVal sql_to_Execute As String = "", _
    Optional ByVal Return_Empty_Row_if_Missing As Boolean = False, _
    Optional ByVal custom_parameters As String = "", _
    Optional ByVal custom_parameter_values() As Object = Nothing) As DataRow
        Try
            If sql_to_Execute = "" Then
                sql_to_Execute = ASCMAIN1.sql
            End If
            Dim tbl As DataTable = GetDataTable(sql_to_Execute, , , , 0, custom_parameters, custom_parameter_values)
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
            Throw New Exception(String.Format("Exception in GetDataRow: {0}", ex.Message))
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Get a datarow using the SQL set in the global variable and doesn't return empty row if missing
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDataRowL()
        Return GetDataRowL(String.Empty, False)
    End Function

    ''' <summary>
    ''' Get a datarow using the SQL passed in the first parameter, and doesn't return empty row if missing
    ''' </summary>
    ''' <param name="sqlToExecute"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDataRowL(ByVal sqlToExecute As String)
        Return GetDataRowL(sqlToExecute, False)
    End Function

    ''' <summary>
    ''' Get a datarow using the SQL set in the global variable, and return a an empty row if missing depending on the parameter passed
    ''' </summary>
    ''' <param name="returnEmptyRowIfMissing"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDataRowL(ByVal returnEmptyRowIfMissing As Boolean)
        Return GetDataRowL(String.Empty, returnEmptyRowIfMissing)
    End Function

    ''' <summary>
    ''' Get a datarow using the SQL passed in the first parameter, and return an emtpy row if missing depending on the second parameter
    ''' </summary>
    ''' <param name="sql_to_Execute"></param>
    ''' <param name="Return_Empty_Row_if_Missing"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDataRowL( _
    ByVal sql_to_Execute As String, _
    ByVal Return_Empty_Row_if_Missing As Boolean) As DataRow
        Try
            If sql_to_Execute = "" Then
                sql_to_Execute = ASCMAIN1.sql
            End If
            Dim tbl As DataTable = GetDataTable(sql_to_Execute)
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
            Throw New Exception(String.Format("Exception in GetDataRowL: {0}", ex.Message))
            Return Nothing
        End Try
    End Function

    Public Function GetDataValue(Optional ByVal sql_to_Execute As String = "") As String
        Dim result As String = ""

        Try
            With ASCMAIN1.oraCon.CreateCommand
                If sql_to_Execute = "" Then
                    sql_to_Execute = ASCMAIN1.sql
                End If
                .CommandText = sql_to_Execute
                If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                    '.Transaction = ASCMAIN1.T
                End If
                result = .ExecuteScalar & ""
                .Dispose()
                Return result
            End With

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

        Return result
    End Function

    Public Function GetDataValue( _
    ByVal sql_to_execute As String, _
    ByVal params As String, _
    ByVal ParamArray PARMs() As Object) As String

        Dim result As String = ""
        Try
            Using cmd As OracleCommand = ASCMAIN1.oraCon.CreateCommand
                If sql_to_execute = "" Then
                    sql_to_execute = ASCMAIN1.sql
                End If

                cmd.CommandText = SQL_Server_Conversions(sql_to_execute)
                '.Transaction = ASCMAIN1.T

                If params <> "" Then
                    Create_Parameters(cmd, params, PARMs)
                End If
                result = cmd.ExecuteScalar & ""
            End Using

        Catch ex As Exception

        End Try

        Return result
    End Function

    Public Function Table2Hash(ByVal tbl As DataTable, ByVal COLUMN_NAME_key As String, ByVal COLUMN_NAME_item As String)
        Dim htb As New Hashtable
        For Each r As DataRow In tbl.Rows
            Try
                htb.Add(r.Item(COLUMN_NAME_key), r.Item(COLUMN_NAME_item))
            Catch ex As Exception
            End Try
        Next
        Return htb
    End Function

    Public Function GetRelation(ByVal dst As DataSet, _
    ByVal TABLE_NAME_parent As String, _
    ByVal TABLE_NAME_child As String, _
    Optional ByVal COLUMN_NAMEsParent As String = "", _
    Optional ByVal COLUMN_NAMEsChild As String = "") As DataRelation

        Dim DC_parent() As DataColumn
        Dim DC_child() As DataColumn
        ReDim DC_parent(0)
        ReDim DC_child(0)
        If COLUMN_NAMEsParent = "" Then
            Dim dt As DataTable = dst.Tables(TABLE_NAME_parent)
            For i As Integer = 0 To dt.PrimaryKey.Length - 1
                Dim COLUMN_NAME As String = dt.Columns(i).ColumnName
                COLUMN_NAMEsParent &= COLUMN_NAME
                ReDim Preserve DC_parent(i)
                DC_parent(i) = dst.Tables(TABLE_NAME_parent).Columns(COLUMN_NAME)
                ReDim Preserve DC_child(i)
                DC_child(i) = dst.Tables(TABLE_NAME_child).Columns(COLUMN_NAME)
            Next
            COLUMN_NAMEsParent = Mid$(COLUMN_NAMEsParent, 2)
        Else
            If COLUMN_NAMEsChild = "" Then
                COLUMN_NAMEsChild = COLUMN_NAMEsParent
            End If
            Dim i As Integer = 0
            For Each COLUMN_NAME As String In Split(COLUMN_NAMEsParent, ",")
                COLUMN_NAME = Trim(COLUMN_NAME)
                ReDim Preserve DC_parent(i)
                DC_parent(i) = dst.Tables(TABLE_NAME_parent).Columns(COLUMN_NAME)
                ReDim Preserve DC_child(i)
                Dim COLUMN_NAME_child = Trim(Split(COLUMN_NAMEsChild, ",")(i))
                DC_child(i) = dst.Tables(TABLE_NAME_child).Columns(COLUMN_NAME_child)
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

    Public Function DeleteRows( _
    ByVal TABLE_NAME As String, _
    ByVal sqlWhere As String)
        Return DeleteRows(ASCMAIN1.ActiveForm.dst.Tables(TABLE_NAME), sqlWhere)
    End Function

    Public Function DeleteRows( _
    ByVal tbl As DataTable, _
    ByVal sqlWhere As String)

        Dim rows() As DataRow = tbl.Select(sqlWhere)
        If rows.Length > 0 Then
            For Each row As DataRow In rows
                row.Delete()
            Next
        End If
        Return rows.Length
    End Function

    Public Sub Keep_Rows(ByVal tbl As DataTable, _
                                ByVal sql_where_to_keep As String)
        Dim tbl_temp As DataTable = tbl.Clone
        tbl_temp.Merge(New DataView(tbl, sql_where_to_keep, "", DataViewRowState.CurrentRows).ToTable)
        tbl.Rows.Clear()
        tbl.Merge(tbl_temp)
    End Sub

#Region "SelectDistinct"

    Public Function SelectDistinct( _
    ByVal Rows As DataRow(), _
    ByVal ParamArray FieldNames() As String) As DataTable
        If Rows.Length = 0 Then
            Return New DataTable
        Else
            Dim sourceTable As New DataTable
            For Each DC As DataColumn In Rows(0).Table.Columns
                sourceTable.Columns.Add(DC.ColumnName, DC.DataType)
            Next
            'Dim SourceTable As DataTable = Rows(0).Table.Clone

            For I As Long = 0 To Rows.Length - 1
                Dim RR As DataRow = sourceTable.NewRow
                RR.ItemArray = Rows(I).ItemArray
                'For J As Integer = 0 To RR.ItemArray.Length - 1
                '    RR.Item(J) = Rows(I).Item(J)
                'Next
                sourceTable.Rows.Add(RR)
            Next
            sourceTable.AcceptChanges()

            'Try
            '    SourceTable.LoadDataRow(Rows, True)

            'Catch ex As Exception
            'End Try
            ''Stop
            Return SelectDistinct(sourceTable, FieldNames)
        End If

    End Function

    Public Function SelectDistinct( _
    ByVal SourceTableName As String, _
    ByVal ParamArray FieldNames() As String) As DataTable
        Return SelectDistinct(ASCMAIN1.ActiveForm.dst.Tables(SourceTableName), FieldNames)
    End Function

    Public Function SelectDistinct( _
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
            If Not fieldValuesAreEqual(lastValues, Row, FieldNames) Then
                newTable.Rows.Add(createRowClone(Row, newTable.NewRow(), FieldNames))

                setLastValues(lastValues, Row, FieldNames)
            End If
        Next

        Return newTable
    End Function

    Private Shared Function fieldValuesAreEqual(ByVal lastValues() As Object, ByVal currentRow As DataRow, ByVal fieldNames() As String) As Boolean
        Dim areEqual As Boolean = True

        For i As Integer = 0 To fieldNames.Length - 1
            If lastValues(i) Is Nothing OrElse Not lastValues(i).Equals(currentRow(fieldNames(i))) Then
                areEqual = False
                Exit For
            End If
        Next

        Return areEqual
    End Function

    Private Shared Function createRowClone(ByVal sourceRow As DataRow, ByVal newRow As DataRow, ByVal fieldNames() As String) As DataRow
        For Each field As String In fieldNames
            newRow(field) = sourceRow(field)
        Next

        Return newRow
    End Function

    Private Shared Sub setLastValues(ByVal lastValues() As Object, ByVal sourceRow As DataRow, ByVal fieldNames() As String)
        For i As Integer = 0 To fieldNames.Length - 1
            lastValues(i) = sourceRow(fieldNames(i))
        Next
    End Sub
#End Region

    Public Function SQL_Server_Conversions(ByVal sql As String) As String
        Dim sql2 As String = sql
        If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
            sql2 = Replace(sql2, "(+)", "") ' OUTER JOIN PROBLEM
            sql2 = Replace(sql2, "NVL(", "ISNULL(")
            sql2 = Replace(sql2, "TO_DATE(", "CONVERT(DATETIME,")
            sql2 = Replace(sql2, "SUM(DECODE", "SUM(dbo.DECODEN")
            sql2 = Replace(sql2, "SUM (DECODE", "SUM (dbo.DECODEN")
            sql2 = Replace(sql2, "DECODE", "dbo.DECODE")
            sql2 = Replace(sql2, "dbo.dbo.DECODE", "dbo.DECODE")
            sql2 = Replace(sql2, ":PARM", "@PARM")
            sql2 = Replace(sql2, "SUBSTR(", "SUBSTRING(")
            sql2 = Replace(sql2, "TRIM(", "dbo.TRIM(")
            sql2 = Replace(sql2, "TRIM (", "dbo.TRIM (")
            sql2 = Replace(sql2, " || ", " + ")
            sql2 = Replace(sql2, " ROWNUM <1", " 1<>1")
            'sql2 = Replace(sql2, "ROWNUM", "ROW_NUMBER() OVER (ORDER BY GETDATE())")
        End If
        Return sql2
    End Function

    Public Sub Fill_Schema(ByVal ada As OracleDataAdapter, ByVal tbl As DataTable)
        Dim sql As String = ada.SelectCommand.CommandText
        'If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
        Dim parmx As Int16 = 0
        Do
            sql = Replace(sql, ASCMAIN1.DBS_PARAMETER & "PARM" & CStr(parmx), "NULL")
            parmx += 1
        Loop While InStr(sql & CStr(parmx), ASCMAIN1.DBS_PARAMETER & "PARM") <> 0
        'End If

        With New OracleDataAdapter(sql, ASCMAIN1.oraCon)
            '.SelectCommand.Transaction = ASCMAIN1.T
            .FillSchema(tbl, SchemaType.Source)
            .Dispose()
        End With
    End Sub

    Public Sub New()
        ASCMAIN1 = New ASCMAIN1
    End Sub


    Public Sub Dispose() Implements IDisposable.Dispose
        If ASCMAIN1 IsNot Nothing Then
            ASCMAIN1.Dispose()
            ASCMAIN1 = Nothing
        End If
    End Sub
End Class