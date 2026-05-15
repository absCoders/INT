Imports Oracle.ManagedDataAccess.Client

Public Class ASCBASE1
    Implements IDisposable

    Public dst As New DataSet
    Public TDAs As New Dictionary(Of String, OracleDataAdapter)
    Public TBLs As New Dictionary(Of String, DataTable)
    Public pROWs As New Dictionary(Of String, DataRow)
    Public TBL_SCHEMAs As New Dictionary(Of String, DataTable)
    Public DVWs As New Dictionary(Of String, DataView)
    
    Public ROWs As New Dictionary(Of String, DataRow)
    Public CMDs As New Dictionary(Of String, OracleCommand)
    Public BA_CMDs As New Dictionary(Of String, OracleCommand())
    '  Public frmASFBASE0 As ASFBASE0
    Dim TABLE_NAMEs As List(Of String)
    Public ASCMAIN1 As ASCMAIN1
    Public ASCDATA1 As ASCDATA1

    Sub Dispose() Implements IDisposable.Dispose
        If CMDs IsNot Nothing AndAlso CMDs.Count <> 0 Then
            For Each CMD_key As String In CMDs.Keys
                Dim cmd As OracleCommand = CMDs(CMD_key)
                For Each param As OracleParameter In cmd.Parameters
                    param.Dispose()
                Next
                cmd.Dispose()
            Next
        End If
        CMDs = Nothing

        If BA_CMDs IsNot Nothing AndAlso BA_CMDs.Count <> 0 Then
            For Each CMD_key As String In BA_CMDs.Keys
                Dim cmds() As OracleCommand = BA_CMDs(CMD_key)
                For Each cmd As OracleCommand In cmds
                    If cmd IsNot Nothing Then
                        For Each param As OracleParameter In cmd.Parameters
                            param.Dispose()
                        Next
                        cmd.Dispose()
                    End If
                Next
                cmds = Nothing
            Next
        End If
        BA_CMDs = Nothing

        If TDAs IsNot Nothing Then
            For Each tda As OracleDataAdapter In TDAs.Values
                tda.Dispose()
            Next
        End If
        TDAs = Nothing
    End Sub

    Public Sub New( _
    ByVal frm As ASCBASE0, _
    Optional ByVal clone_dst As Boolean = False, _
    Optional ByVal vTABLE_NAMEs As List(Of String) = Nothing)

        '  frmASCBASE0 = frm

        ASCDATA1 = New ASCDATA1
        ASCMAIN1 = ASCDATA1.ASCMAIN1
        ASCMAIN1.ASCDATA1 = ASCDATA1

        If clone_dst Then
            TABLE_NAMEs = vTABLE_NAMEs

            dst = frm.dst.Copy

            Copy_Dictionary(TDAs, frm.TDAs)
            Copy_Dictionary(TBLs, frm.TBLs)
            Copy_Dictionary(pROWs, frm.pROWs)
            Copy_Dictionary(TBL_SCHEMAs, frm.TBL_SCHEMAs)
            Copy_Dictionary(DVWs, frm.DVWs)

            Copy_Dictionary(ROWs, frm.ROWs)
            Copy_Dictionary(CMDs, frm.CMDs)
        End If
    End Sub

    Sub Copy_Dictionary( _
    ByVal D1 As Dictionary(Of String, OracleDataAdapter), _
    ByRef D2 As Dictionary(Of String, OracleDataAdapter))

        For Each k As String In D2.Keys
            If TABLE_NAMEs Is Nothing OrElse TABLE_NAMEs.Contains(k) Then
                D1.Add(k, D2(k))
            End If
        Next
    End Sub

    Sub Copy_Dictionary( _
    ByVal D1 As Dictionary(Of String, DataTable), _
    ByRef D2 As Dictionary(Of String, DataTable))

        For Each k As String In D2.Keys
            If TABLE_NAMEs Is Nothing OrElse TABLE_NAMEs.Contains(k) Then
                D1.Add(k, dst.Tables(k))
            End If
        Next
    End Sub

    Sub Copy_Dictionary( _
    ByVal D1 As Dictionary(Of String, DataRow), _
    ByRef D2 As Dictionary(Of String, DataRow))

        For Each k As String In D2.Keys
            D1.Add(k, D2(k))
        Next
    End Sub

    Sub Copy_Dictionary( _
    ByVal D1 As Dictionary(Of String, DataView), _
    ByRef D2 As Dictionary(Of String, DataView))

        For Each k As String In D2.Keys
            D1.Add(k, New DataView(dst.Tables(k)))
        Next
    End Sub

    Sub Copy_Dictionary( _
    ByVal D1 As Dictionary(Of String, String), _
    ByRef D2 As Dictionary(Of String, String))

        For Each k As String In D2.Keys
            D1.Add(k, D2(k))
        Next
    End Sub

    Sub Copy_Dictionary( _
    ByVal D1 As Dictionary(Of String, OracleCommand), _
    ByRef D2 As Dictionary(Of String, OracleCommand))

        For Each k As String In D2.Keys
            D1.Add(k, D2(k))
        Next
    End Sub

    Function Fill_Record( _
    ByVal TABLE_NAME As String, _
    ByVal KEY_VALUE As Object, _
    Optional ByVal create_row_if_non_existent As Boolean = False, _
    Optional ByVal ClearBeforeFilling As Boolean = True) As DataRow

        If KEY_VALUE Is Nothing Then
            Return Fill_Record(TABLE_NAME, , create_row_if_non_existent, ClearBeforeFilling)
        Else
            Return Fill_Record(TABLE_NAME, New Object() {KEY_VALUE}, create_row_if_non_existent, ClearBeforeFilling)
        End If

    End Function

    Function Fill_Record( _
    ByVal TABLE_NAME As String, _
    Optional ByVal Parameters() As Object = Nothing, _
    Optional ByVal create_row_if_non_existent As Boolean = False, _
    Optional ByVal ClearBeforeFilling As Boolean = True) As DataRow

        Dim row As DataRow = Nothing

        If Not ClearBeforeFilling Then

            'If Not create_row_if_non_existent And ASCMAIN1.Running_in_VS Then Stop - ARFPYMT2 WAS STOPPING HERE AFTER LOOKING AT THE DATATABLE, SEEING THAT THERE WAS NO ROW FOR THE KEY SPECIFIED, AND THEN TRYING TO FILL THE DATATABLE FROM ORACLE - IT APPEARS THAT FILL_RECORD WOULD HAVE CHECKED THE DATATABLE FOR ME USING THE CODE BELOW,  MAKING A KEYED READ TO THE DATATABLE UNNECESSARY
            'If create_row_if_non_existent And Not ClearBeforeFilling Then
            ' APFCHKP1 RELIES ON US FINDING THE SAME RECORD AND NOT RE-FILLING FROM THE DB IF THESE CONDITIONS ARE TRUE
            ' think of (Not ClearBeforeFilling) as the parameter you would use to
            '  force this routine to check the local datatable before going out to Oracle
            ' sometimes you need to get a refreshed version of a row from a database 
            '  even though the local cache (datatable) already has a row with that key 
            '  - so use ClearBeforeFilling for that one

            row = TBLs(TABLE_NAME).Rows.Find(Parameters)
            If row IsNot Nothing Then
                Return row
                Exit Function
            End If
        End If

        Dim records_filled As Integer = _
        Fill_Records(TABLE_NAME, Parameters, ClearBeforeFilling)

        If records_filled = 0 Then
            If create_row_if_non_existent Then
                row = TBLs(TABLE_NAME).NewRow
                If Parameters IsNot Nothing Then
                    For i As Integer = 0 To UBound(Parameters)
                        row.Item(i) = Parameters(i)
                    Next
                End If
                TBLs(TABLE_NAME).Rows.Add(row)
            End If
        Else
            row = TBLs(TABLE_NAME).Rows(TBLs(TABLE_NAME).Rows.Count - 1)
        End If
        Return row
    End Function

    Function Fill_Records( _
    ByVal TABLE_NAME As String, _
    ByVal KEY_VALUE As String, _
    Optional ByVal ClearBeforeFilling As Boolean = True, _
    Optional ByVal Temp_Select As String = "", _
    Optional ByVal tblSubstitute As DataTable = Nothing) As Integer

        If KEY_VALUE = "" Then
            Return Fill_Records(TABLE_NAME, , ClearBeforeFilling, Temp_Select, tblSubstitute)
        Else
            Return Fill_Records(TABLE_NAME, New String() {KEY_VALUE}, ClearBeforeFilling, Temp_Select, tblSubstitute)
        End If

    End Function

    Public Function Fill_Records( _
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
            Throw New Exception("Unknown Table " & TABLE_NAME & " in Fill Records")
        Else

            If ClearBeforeFilling Then
                tbl.Rows.Clear()

                Dim dvw As DataView = tbl.DefaultView
                dvw.RowFilter = ""

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
                            If .Parameters(i).DbType = DbType.DateTime Then
                                .Parameters(i).Value = CDate(Parameters(i))
                            Else
                                .Parameters(i).Value = Parameters(i)
                            End If
                            ' .Parameters(i).Value = Parameters(i)
                        Next
                    Else
                        For i As Integer = 0 To .Parameters.Count - 1
                            .Parameters(i).Value = System.DBNull.Value
                        Next
                    End If
                End With
            End If

            Dim ada As OracleDataAdapter
            Dim original_select_commandtext As String = TDAs(TABLE_NAME).SelectCommand.CommandText
            If Temp_Select = "" Then
                ada = TDAs(TABLE_NAME)
            Else
                If Parameters IsNot Nothing Then
                    ada = TDAs(TABLE_NAME)
                    ada.SelectCommand.CommandText = Temp_Select
                Else
                    ada = New OracleDataAdapter(Temp_Select, ASCMAIN1.oraCon)
                End If
            End If

            Dim XDT As Date = Now
            'ada.SelectCommand.Transaction = ASCMAIN1.T
            'ada.SelectCommand.CommandTimeout = 60
            Fill_Records = ada.Fill(tbl)

            If ASCMAIN1.ActiveForm IsNot Nothing AndAlso dst.Tables.Contains("ASTSQLX1") Then

                Dim commandText As String = ada.SelectCommand.CommandText
                If commandText.Length > 4000 Then
                    commandText = commandText.Substring(0, 4000).Trim
                End If

                dst.Tables("ASTSQLX1").Rows.Add(New Object() _
                {ASCMAIN1.SESSION_NO, _
                 ASCMAIN1.ActiveForm.SELECTION_NO, ASCMAIN1.ActiveForm.RE_XNO, _
                 XDT, Now.Subtract(XDT).Seconds, commandText})
            End If

            TDAs(TABLE_NAME).SelectCommand.CommandText = original_select_commandtext
        End If
    End Function


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
    Public Sub Create_TDA( _
    ByRef tbl As DataTable, _
    ByVal TABLE_NAME As String, _
    ByVal sql_custom As String, _
    Optional ByVal NumberOfKeysUsedToSelect As Integer = -1, _
    Optional ByVal for_update As Boolean = True, _
    Optional ByVal custom_parameters As String = "", _
    Optional ByVal Key_Field_Count As Integer = -1, _
    Optional ByVal Update_COLUMN_NAMEs As String = "", _
    Optional ByVal SCHEMA As String = "")

        Dim tblTABLE_NAME As String = TABLE_NAME
        If Not tbl.TableName Like "Table*" Then
            tblTABLE_NAME = tbl.TableName
        End If

        If TDAs.ContainsKey(tblTABLE_NAME) Then
            TDAs.Remove(tblTABLE_NAME)
        End If
        TDAs.Add(tblTABLE_NAME, ASCDATA1.GetDataAdapter(tbl, TABLE_NAME, sql_custom, for_update, Key_Field_Count, False, NumberOfKeysUsedToSelect, Update_COLUMN_NAMEs, SCHEMA))
        If custom_parameters.Length <> 0 Then
            Dim ptbl As New DataTable
            Call ASCDATA1.Create_Parameters(TDAs(tblTABLE_NAME).SelectCommand, custom_parameters, ptbl)
            pROWs.Add(tblTABLE_NAME, ptbl.NewRow)
        End If

        Dim tbl_Schema As DataTable = Nothing
        If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
            'why do we need tbl_schemas? - SO THAT WE HAVE CONTROL OVER THE DATATYPE THAT THE DATACOLUMNS BECOME
            Dim sqltemp As String = TDAs(tblTABLE_NAME).SelectCommand.CommandText
            Dim parmx2 As Int16 = 0
            Do
                sqltemp = Replace(sqltemp, ASCMAIN1.DBS_PARAMETER & "PARM" & CStr(parmx2), "NULL")
                parmx2 += 1
            Loop While InStr(sqltemp & CStr(parmx2), ASCMAIN1.DBS_PARAMETER & "PARM") <> 0
            For Each p As OracleParameter In TDAs(tblTABLE_NAME).SelectCommand.Parameters
                sqltemp = Replace(sqltemp, ASCMAIN1.DBS_PARAMETER & p.ParameterName, "NULL")
            Next
            sqltemp = "Select * from (" & sqltemp & ") x where 1 <> 1"
            ASCMAIN1.oraCmd.CommandText = sqltemp
            'ASCMAIN1.oraCmd.Transaction = ASCMAIN1.T
            With ASCMAIN1.oraCmd.ExecuteReader
                tbl_Schema = .GetSchemaTable
                .Close()
                .Dispose()
            End With
        Else
            With TDAs(tblTABLE_NAME).SelectCommand.ExecuteReader
                tbl_Schema = .GetSchemaTable
                .Close()
                .Dispose()
            End With
        End If

        tbl_Schema.PrimaryKey = New DataColumn() {tbl_Schema.Columns("ColumnName")}

        TBL_SCHEMAs.Add(tblTABLE_NAME, tbl_Schema)

        For Each dc As DataColumn In tbl.Columns
            Dim row_Schema As DataRow = tbl_Schema.Rows.Find(dc.ColumnName)
            If row_Schema IsNot Nothing Then
            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer And row_Schema("DataType") Is GetType(System.String) Or row_Schema("DataType") Is GetType(System.DateTime) Then
                ' do nothing 
            Else
                If Not row_Schema("NumericPrecision").Equals(DBNull.Value) Then
                    If Val(row_Schema("NumericScale")) = 0 Then
                        dc.DataType = GetType(System.Int64)
                    Else
                        dc.DataType = GetType(System.Decimal)
                    End If
                End If
            End If

            dc.ReadOnly = False
            End If
        Next

        TBLs.Add(tblTABLE_NAME, tbl)
        Dim dvw As New DataView(tbl)
        DVWs.Add(tblTABLE_NAME, dvw)
    End Sub
     

    ''' <summary>
    ''' This function returns a single row using the TABLE_NAME (ie: the LOOKUP_NAME or RESULTSET_NAME) specified.
    ''' The KEY may be a single string value, or else a string array of values.
    ''' If multiple rows are returned from the database, only the 1st row detected will be returned.
    ''' Nothing is returned if no rows are found in the database.
    ''' An error will result if the TABLE_NAME is not found in the CMDs collection.
    ''' Specify values for KEYs and/or custom parameters in the KEYS() array in the order in which they were defined, keys first if both were used.
    ''' The row is returned by the function, but is also available in the variable cdr (current data row), and as ROWs(TABLE_NAME).
    ''' If an empty row is returned, it does NOT have the key values packed into the key columns of the empty row even if keys were specified - the row is a .NewRow from the table.
    ''' </summary>
    ''' <param name="TABLE_NAME">The LOOKUP_NAME or the RESULTSET_NAME used to create the read-only command in the CMSs collection.</param>
    ''' <param name="KEY"></param>
    ''' <param name="Return_Empty_Row_if_Missing"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LookUp( _
    ByVal TABLE_NAME As String, _
    ByVal KEY() As String, _
    Optional ByVal Return_Empty_Row_if_Missing As Boolean = False) As DataRow

        Dim row As DataRow = Nothing

        If Not ROWs.ContainsKey(TABLE_NAME) Then
            Call Create_Lookup(TABLE_NAME)
        End If

        ROWs(TABLE_NAME).Table.Rows.Clear()

        With CMDs(TABLE_NAME)

            ' THIS LINE WAS ADDED WHEN WE TRIED READ GLTACCT1 IN UPDATE OF PMFBILL1
            '.Transaction = ASCMAIN1.T

            For i As Integer = 0 To UBound(KEY)
                .Parameters(i).Value = KEY(i)
            Next

            With .ExecuteReader
                If .HasRows Then
                    .Read()
                    row = ROWs(TABLE_NAME).Table.NewRow
                    'row = TBLs(TABLE_NAME).NewRow
                    For i As Integer = 0 To .FieldCount - 1
                        row.Item(i) = .Item(i)
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
                .Close()
                .Dispose()
            End With
        End With

        If row IsNot Nothing Then
            ROWs(TABLE_NAME) = row
        End If
        Return row
    End Function

    Public Function LookUp( _
    ByVal TABLE_NAME As String, _
    ByVal KEY As String, _
    Optional ByVal Return_Empty_Row_if_Missing As Boolean = False) As DataRow
        Return LookUp(TABLE_NAME, New String() {KEY}, Return_Empty_Row_if_Missing)
    End Function

    ''' <summary>
    ''' This routine creates a command that (when invoked) returns a single row.  
    ''' This row is READ-ONLY.
    ''' The most common use for creating a Lookup is to return a row from a master table following a keyed lookup.
    ''' However, you may also provide a filter condition on that lookup (such as STATUS = 'A'), or a parameterized condition (STATUS = :PARM1), or an alternate key (USER_NAME = :PARM1).
    ''' In code, you would normally use a Lookup as follows: LookUp("ASTUSER1","joe").
    ''' This would return a single row in the variable cdr (current data row), as well as in ROWs(LOOKUP_NAME).
    ''' The variable cdr would be set to nothing if the row does not exist in the DBS table.
    ''' This function will cache the command in the CMDs collection, a table definition in the TBLs collection, and a row definition in the ROWs collection, all keyed by LOOKUP_NAME.
    ''' Since the CMDs collection is defined to ASFBASE1, the Lookups defined in a single form belong to that form, and are not usable across forms.
    ''' If a Lookup is defined twice in the same form, an error will result.
    ''' Lookups may be automatically created (using defaults) and used (with parameters for the key) by specifying a value for the ABSLookupTableName property of a control.
    ''' </summary>
    ''' <param name="LOOKUP_NAME">A logical name for the Lookup.  This is normally the TABLE_NAME used in the Lookup, but may also be a Column Name or Logical Name prefixed by a Table Name (like ASTUSER1.ACTIVE_USERS).  IMPORTANT: Always use a valid DBS Table Name, either as the LOOKUP_NAME, or as the dotted prefix in the LOOKUP_NAME.</param>
    ''' <param name="column_list">A list of columns to be returned from the Table.  Normally this omitted or "", to signify "*", but can return a subset of the columns available.</param>
    ''' <param name="where_clause">Used for 2 purposes: 1) to qualify the row returned, such as STATUS = 'A' and 2) to set up a custom parameter, such as STATUS = :PARM1.  Note that custom parameters may be specified in addition to keys (this provides a parameterized filter).  To use custom parameters instead of keys, make sure you specify create_parameters_for_key = false.</param>
    ''' <param name="custom_parameters">A string indicating the number and types of parameters used in the where clause, such as "VVIND" to indicate 5 parameters, 2 Varchar2, 1 Integer, 1 Number, and 1 Date.  Always use :PARMx as your parameter names in the where clause.</param>
    ''' <param name="create_parameters_for_key">True/False indicating whether parameters should be set up for the key field(s).  Note that parameters will always be set up for the key unless you explicitly indicate they should not; this is because the most common use for lookups is the simple retreival of a row to a table for a given key.</param>
    ''' <remarks>these are the remarks - where do they appear?</remarks>
    Public Sub Create_Lookup( _
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
            tbl = ASCDATA1.GetDataTable("*", TABLE_NAME, -1, False)
        Else
            ASCMAIN1.sql = "Select " & column_list & " from " & TABLE_NAME
            tbl = ASCDATA1.GetDataTable("", TABLE_NAME, -1, False)
        End If
        If tbl.PrimaryKey.Length = 0 Then
            tbl.PrimaryKey = New DataColumn() {tbl.Columns(0)}
        End If
        'tbl.TableName = LOOKUP_NAME ?

        Dim cmd As OracleCommand = ASCMAIN1.oraCon.CreateCommand
        Dim sql As String = ""
        If create_parameters_for_key Then
            For Each dc As DataColumn In tbl.PrimaryKey
                Dim COLUMN_NAME As String = dc.ColumnName
                sql = sql & " and (" & COLUMN_NAME & " = " & ASCMAIN1.DBS_PARAMETER & "S_" & COLUMN_NAME & ")"
                Dim pp As New OracleParameter
                pp.ParameterName = "S_" & COLUMN_NAME
                pp.Direction = ParameterDirection.Input
                cmd.Parameters.Add(pp)
                Select Case dc.DataType.Name
                    Case "String"
                        'pp.DbType = sqlDbType.NVarChar 
                        pp.DbType = DbType.String
                        'SETTING THE SIZE MADE IT SO A LOOKUP ON AN 11 CHAR VALUE WHERE THE 1ST 10 CARS WERE VALID, RETURNED A RECORD
                        'pp.Size = tbl.Columns(COLUMN_NAME).MaxLength
                    Case "Int16"
                        'pp.DbType = SqlDbType.SmallInt
                        pp.DbType = DbType.Int16
                    Case "Int32"
                        'pp.DbType = SqlDbType.Int
                        pp.DbType = DbType.Int32
                    Case "Int64"
                        'pp.DbType = SqlDbType.BigInt
                        pp.DbType = DbType.Int64
                    Case "Decimal"
                        pp.DbType = DbType.Decimal
                    Case Else
                        'Stop
                        Throw New Exception("Unknown Type in Create_Lookup")
                End Select
            Next
        End If

        If custom_parameters <> "" Then Call ASCDATA1.Create_Parameters(cmd, custom_parameters)

        If where_clause <> "" Then
            where_clause = " and " & where_clause
        End If
        cmd.CommandText = "Select * from " & TABLE_NAME & ASCMAIN1.SQL_Add_WHERE(sql & where_clause)

        If CMDs.ContainsKey(LOOKUP_NAME) Then
            CMDs.Remove(LOOKUP_NAME)
        End If
        CMDs.Add(LOOKUP_NAME, cmd)
        If ROWs.ContainsKey(LOOKUP_NAME) Then
            ROWs.Remove(LOOKUP_NAME)
        End If
        ROWs.Add(LOOKUP_NAME, tbl.NewRow)
    End Sub

    Sub Write_DataSet( _
    Optional ByVal xml As Boolean = False, _
    Optional ByVal FOLDER_NAME As String = "", _
    Optional ByVal FILE_NAME As String = "")

        If dst Is Nothing Then
            Exit Sub

        End If
        Try

        Catch ex As Exception

        End Try
        If dst.Tables.Count > 0 Then
            'Dim DATASET_NAME As String = Me.Name & "dst"
            ''Mid$(DATASET_NAME, 3, 1) = "D"
            'dst.DataSetName = DATASET_NAME
            Dim DATASET_NAME As String = dst.DataSetName

            If FOLDER_NAME = "" Then
                FOLDER_NAME = ASCMAIN1.Folders("DataSets")
                If ASCMAIN1.Running_in_VS Then
                    FOLDER_NAME = ASCMAIN1.Folders("root") & Mid$(DATASET_NAME, 1, 2) & "\DataSets\"
                End If
            End If
            If Not My.Computer.FileSystem.DirectoryExists(FOLDER_NAME) Then
                My.Computer.FileSystem.CreateDirectory(FOLDER_NAME)
            End If

            If FILE_NAME = "" Then
                FILE_NAME = DATASET_NAME
            End If

            If xml Then
                dst.WriteXml(FOLDER_NAME & FILE_NAME & ".xml", XmlWriteMode.WriteSchema)
            Else
                dst.WriteXmlSchema(FOLDER_NAME & FILE_NAME & ".xsd")
            End If
        End If
    End Sub

    ''' <summary>
    ''' Performs the Update Command for the Table Data Adapter.  Use sql_Delete only if you want to delete rows (using the supplied sql_Delete statement) and then have the current rows Inserted.
    ''' </summary>
    ''' <param name="TABLE_NAME"></param>
    ''' <param name="sql_Delete">Specify a complete Delete statement, or else just the where clause.  If sql_delete does not begin with the word 'Delete', then the clause 'Delete from {TABLE_NAME} where ' will be pre-pended to the clause supplied.</param>
    ''' <remarks></remarks>
    Sub Update_Record_TDA(ByVal TABLE_NAME As String, Optional ByVal sql_Delete As String = "")

        ' Without the next 2 lines, data in text boxes in single row datatables (like header tables) will not get written to Oracle
        'Dim X As CurrencyManager = Me.BindingContext(dst.Tables(TABLE_NAME))
        'X.EndCurrentEdit()

        'If AUDIT.ContainsKey(TABLE_NAME) Then
        '    WriteAuditTrail(TABLE_NAME)
        'End If

        If sql_Delete <> "" Then

            Dim TABLE_NAME_ora As String = TABLE_NAME

            ' need to find out if the Oracle Table is not the same as the ADO.Net datatable - this happens when the Oracle table is a Temp ASW Table
            ' this code probably needs to be handled by a dictionary that points directly to the Oracle table connected to the TDA - 
            Dim adaT As String = ""
            If TDAs(TABLE_NAME).UpdateCommand IsNot Nothing Then
                adaT = Mid(TDAs(TABLE_NAME).UpdateCommand.CommandText, 8, 8)
            ElseIf TDAs(TABLE_NAME).InsertCommand IsNot Nothing Then
                adaT = Mid(TDAs(TABLE_NAME).InsertCommand.CommandText, 8, 8)
            ElseIf TDAs(TABLE_NAME).DeleteCommand IsNot Nothing Then
                adaT = Mid(TDAs(TABLE_NAME).DeleteCommand.CommandText, 8, 8)
            End If
            If adaT <> TABLE_NAME And adaT.StartsWith("ASW") Then
                TABLE_NAME_ora = adaT
            End If

            If Mid(Trim(sql_Delete), 1, 6).ToUpper <> "DELETE" Then
                sql_Delete = "Delete from " & TABLE_NAME_ora & " where " & sql_Delete
            End If
            ASCDATA1.ExecuteSQL(sql_Delete)
            TBLs(TABLE_NAME).AcceptChanges()
            For Each row As DataRow In TBLs(TABLE_NAME).Rows
                row.SetAdded()
            Next
        Else
        End If

        'If TDAs(TABLE_NAME).InsertCommand IsNot Nothing Then TDAs(TABLE_NAME).InsertCommand.Transaction = ASCMAIN1.T
        'If TDAs(TABLE_NAME).DeleteCommand IsNot Nothing Then TDAs(TABLE_NAME).DeleteCommand.Transaction = ASCMAIN1.T
        'If TDAs(TABLE_NAME).UpdateCommand IsNot Nothing Then TDAs(TABLE_NAME).UpdateCommand.Transaction = ASCMAIN1.T               

        Dim I As Int64 = TDAs(TABLE_NAME).Update(dst.Tables(TABLE_NAME))
    End Sub

    Sub Create_BAs(ByVal TABLE_NAME As String)

        ' SUPPORTING ONLY INSERT COMMANDS (FOR NOW) - OTHERS MAY NEED SPECIAL ARRAYS FOR KEY COLUMNS

        Dim CMDs(3) As OracleCommand

        For i As Integer = 0 To 3
            If i = 1 Then
                CMDs(i) = New OracleCommand
                CMDs(i).Connection = ASCMAIN1.oraCon
            End If
        Next i

        'CMDs(0).CommandText = TDAs(TABLE_NAME).SelectCommand.CommandText
        CMDs(1).CommandText = TDAs(TABLE_NAME).InsertCommand.CommandText
        'CMDs(2).CommandText = TDAs(TABLE_NAME).UpdateCommand.CommandText
        'CMDs(3).CommandText = TDAs(TABLE_NAME).DeleteCommand.CommandText

        If BA_CMDs.Count <> 0 Then
            BA_CMDs.Clear()
        End If
        BA_CMDs.Add(TABLE_NAME, CMDs)

        ' Dim BAs As New Dictionary(Of String, Object())

        For i As Integer = 0 To TBLs(TABLE_NAME).Columns.Count - 1

            Dim COLUMN_NAME As String = TBLs(TABLE_NAME).Columns(i).ColumnName
            Dim DATA_DTYPE As String = TBLs(TABLE_NAME).Columns(i).DataType.ToString
            Dim DATA_LENGTH As Integer = TBLs(TABLE_NAME).Columns(i).MaxLength

            Dim p As New OracleParameter
            p.ParameterName = COLUMN_NAME
            Select Case DATA_DTYPE

                Case "System.String"
                    p.DbType = DbType.String
                    p.Size = DATA_LENGTH
                    Dim A() As String = Nothing
                    p.Value = A

                Case "System.DateTime"
                    p.DbType = DbType.Date
                    Dim A() As Date = Nothing
                    p.Value = A

                Case "System.Double"
                    p.DbType = DbType.Double
                    Dim A() As Double = Nothing
                    p.Value = A

                Case "System.Single"
                    p.DbType = DbType.Single
                    Dim A() As Single = Nothing
                    p.Value = A

                Case "System.Decimal"
                    p.DbType = DbType.Decimal
                    Dim A() As Decimal = Nothing
                    p.Value = A

                Case "System.Int16"
                    p.DbType = DbType.Int16
                    Dim A() As Int16 = Nothing
                    p.Value = A

                Case "System.Int32"
                    p.DbType = DbType.Int32
                    Dim A() As Int32 = Nothing
                    p.Value = A

                Case "System.Int64"
                    p.DbType = DbType.Int64
                    Dim A() As Int64 = Nothing
                    p.Value = A

                Case Else
                    Throw New Exception("Unsupported Data Type")
            End Select

            CMDs(1).Parameters.Add(p)

        Next i
    End Sub

    Sub Update_BAs(ByVal TABLE_NAME As String)

        Dim tbl As DataTable = TBLs(TABLE_NAME)
        Dim row_count As Integer = tbl.Rows.Count
        BA_CMDs(TABLE_NAME)(1).ArrayBindCount = row_count

        For c As Integer = 0 To tbl.Columns.Count - 1
            Dim COLUMN_NAME As String = tbl.Columns(c).ColumnName
            Select Case tbl.Columns(c).DataType.ToString
                Case "System.String"
                    Dim A() As String
                    ReDim A(row_count - 1)
                    For r As Integer = 0 To row_count - 1
                        A(r) = tbl.Rows(r).Item(c) & ""
                    Next
                    BA_CMDs(TABLE_NAME)(1).Parameters(COLUMN_NAME).Value = A

                Case "System.DateTime"
                    Dim A() As DateTime
                    ReDim A(row_count - 1)
                    For r As Integer = 0 To row_count - 1
                        If Not tbl.Rows(r).Item(c).Equals(System.DBNull.Value) Then
                            A(r) = tbl.Rows(r).Item(c)
                        End If
                    Next
                    BA_CMDs(TABLE_NAME)(1).Parameters(COLUMN_NAME).Value = A

                Case "System.Double"
                    Dim A() As Double
                    ReDim A(row_count - 1)
                    For r As Integer = 0 To row_count - 1
                        A(r) = tbl.Rows(r).Item(c)
                    Next
                    BA_CMDs(TABLE_NAME)(1).Parameters(COLUMN_NAME).Value = A

                Case "System.Single"
                    Dim A() As Single
                    ReDim A(row_count - 1)
                    For r As Integer = 0 To row_count - 1
                        A(r) = tbl.Rows(r).Item(c)
                    Next
                    BA_CMDs(TABLE_NAME)(1).Parameters(COLUMN_NAME).Value = A

                Case "System.Decimal"
                    Dim A() As Decimal
                    ReDim A(row_count - 1)
                    For r As Integer = 0 To row_count - 1
                        If Not tbl.Rows(r).Item(c).Equals(System.DBNull.Value) Then
                            A(r) = tbl.Rows(r).Item(c)
                        End If
                    Next
                    BA_CMDs(TABLE_NAME)(1).Parameters(COLUMN_NAME).Value = A

                Case "System.Int16"
                    Dim A() As Int16
                    ReDim A(row_count - 1)
                    For r As Integer = 0 To row_count - 1
                        A(r) = tbl.Rows(r).Item(c)
                    Next
                    BA_CMDs(TABLE_NAME)(1).Parameters(COLUMN_NAME).Value = A

                Case "System.Int32"
                    Dim A() As Int32
                    ReDim A(row_count - 1)
                    For r As Integer = 0 To row_count - 1
                        A(r) = tbl.Rows(r).Item(c)
                    Next
                    BA_CMDs(TABLE_NAME)(1).Parameters(COLUMN_NAME).Value = A

                Case "System.Int64"
                    Dim A() As Int64
                    ReDim A(row_count - 1)
                    For r As Integer = 0 To row_count - 1
                        If Not tbl.Rows(r).Item(c).Equals(System.DBNull.Value) Then
                            A(r) = tbl.Rows(r).Item(c)
                        End If
                    Next
                    BA_CMDs(TABLE_NAME)(1).Parameters(COLUMN_NAME).Value = A
                Case Else
                    Throw New Exception("Unsupported Data Type")
            End Select
        Next
        Try
        BA_CMDs(TABLE_NAME)(1).ExecuteNonQuery()

        Catch ex As Exception
            Throw New Exception(String.Format("Exception in Update_BAs: {0}", ex.Message))
        End Try
    End Sub

    Function EnforceConstraints(Optional ByVal enforce As Boolean = True) As String
        Dim hasConstraintErrors As Boolean = False

        Try
            dst.EnforceConstraints = enforce
        Catch ex As ConstraintException
            hasConstraintErrors = True
        End Try

        If Not hasConstraintErrors Then
            Return ""
        End If

        Dim ErrorList As List(Of String) = New List(Of String)
        Dim i As Integer = 0

        For Each tbl As DataTable In dst.Tables
            For Each row As DataRow In tbl.GetErrors()
                ErrorList.Add("Table: " & tbl.TableName & vbCrLf & " Row Error: " & row.RowError)
                i += 1
                If i = 20 Then Exit For
            Next
            If ErrorList.Count > 20 Then Exit For
        Next

        Dim eMessage As String = String.Empty
        For Each sString As String In ErrorList
            eMessage &= vbCrLf & Environment.NewLine & sString
        Next

        Return eMessage
    End Function
End Class
