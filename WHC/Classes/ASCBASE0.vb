Imports System.Math
Imports Oracle.ManagedDataAccess.Client

Public Class ASCBASE0
    Public G As ABSEnvironment
    Public GunEnv As GunEnvironment
    Public tbl As DataTable

    Public Name As String = "ASCBASE0"
    Public MENU_ID As String
    Public MENU_ITEM_TYPE As String
    Public MENU_ITEM_OBJECT As String
    Public MENU_ITEM_DESC As String
    Public MENU_ITEM_SECURITY As String
    Public MENU_ITEM_PP As String
    Public MENU_ITEM_FORM As String
    Public MODULE_ID As String

    Public tblASTOPST1 As New DataTable
    Public tdaASTOPST1 As OracleDataAdapter
    Public rowASTOPST1 As DataRow

    Public XNO As String = "" ' Generated each time you click Proceed, or Change Modes to True
    Public FORM_INSTANCE_NO As String  ' Generated each time you Launch from the Menu

    Public AUDIT As New Dictionary(Of String, String)

    Public TABLE_NAME As String = ""
    Public TABLE_NAME_view As String
    Public COLUMN_NAME As String = ""
    Public VIEW_NAME As String = ""
    Public CODE_VALUE As String = ""
    Public EMsg As String = ""
    Public SELECTION_NO As Integer = 0
    Public tblASFBASE1 As New DataTable
    Public tblASFBASE1_Schema As New DataTable
    Public rowASFBASE1 As DataRow
    Public ScreenMode As Boolean
    Public InquiryMode As Boolean = False
    Public This_Record_Inquiry_Only As Boolean
    Public EntryMode As String
    Public iScreenMode As Integer
    Public not_iScreenMode As Integer
    Public htbCOLUMN_NAME As New Hashtable
    Public Update_CMDs As New Dictionary(Of String, OracleCommand)
    Public error_has_occured As Exception

    Public DATETIME_STAMP As Date
    Public IsLoading As Boolean = False
    Public IsClosing As Boolean = False
    Public IsDone As Boolean = False
    Public remotely_controlled As Boolean = False
    Public bind_to_TABLE_NAME As Boolean = True
    Public ABSReadOnly As New List(Of String)
    Public HFs As New Dictionary(Of String, String)
    Public T As OracleTransaction
    Public Null As System.DBNull = DBNull.Value

    Public RE_XNO As Integer = -1
    Public RE_XNO_STAT As Integer = 0
    Public cdr As DataRow   ' Current DataRow returned from most recent LookUp or Validate_Code

    Dim disable_arrows As Boolean = False

    Public frmASFBASE1s As New Dictionary(Of String, ASCBASE0)

    ' References to ASCBASE1 Objects
    Public dst As New DataSet
    Public TDAs As Dictionary(Of String, OracleDataAdapter)
    Public TBLs As Dictionary(Of String, DataTable)
    Public pROWs As Dictionary(Of String, DataRow)
    Public TBL_SCHEMAs As Dictionary(Of String, DataTable)
    Public DVWs As Dictionary(Of String, DataView)

    Public ROWs As Dictionary(Of String, DataRow)
    Public CMDs As Dictionary(Of String, OracleCommand)
    Public BA_CMDs As Dictionary(Of String, OracleCommand())

    Public ASTDATA1s As New Dictionary(Of String, String)

    Public Bound_DataSources As New List(Of Object)

    Dim EXT_allowed As List(Of String)

    Public ASCDATA1 As ASCDATA1
    Public ASCMAIN1 As ASCMAIN1
    Public clsASCBASE1 As ASCBASE1

    Public Structure Log_Entity
        Public TABLE_NAME As String
        Public TABLE_KEY As String
        Public TABLE_KEY_CAPTION As String
        Public TABLE_KEY_DESC As String
        Public read_only As Boolean
        Public enabled As Boolean
        Public TABLE_KEY_locked As Boolean
    End Structure

    Public Structure Data_Export_Entity
        Public enabled As Boolean
    End Structure

    Public Structure Audit_Entity
        Public TABLE_NAME As String
        Public TABLE_DESC As String
        Public KEY_VALUE As String
        Public KEY_DESC As String
    End Structure
    Sub New()

    End Sub
    Sub New(g As GunEnvironment)

    End Sub
    Sub New(g As ABSEnvironment)
        Me.G = g

        clsASCBASE1 = New ASCBASE1(Me)
        ASCMAIN1 = clsASCBASE1.ASCMAIN1
        ASCDATA1 = clsASCBASE1.ASCDATA1

        ASCMAIN1.DBS_COMPANY = g.DBS_COMPANY
        ASCMAIN1.DBS_PASSWORD = g.DBS_PASSWORD
        ASCMAIN1.DBS_SERVER = g.DBS_SERVER
        If Not Logon_Attempt_Succeeded() Then
            Throw New Exception("Cannot Create Oracle Connection")
        End If

        ASCMAIN1.ActiveForm = Me

        IsLoading = True
        Clear_dst()

        ASCMAIN1.tblASTSQLX1 = Nothing
        Create_TDA(dst.Tables.Add, "ASTSQLX1", "*")
        ASCMAIN1.tblASTSQLX1 = dst.Tables("ASTSQLX1")

        'SELECTION_NO = ASCMAIN1.Register_Form(Me)

        Dim folder_prefix As String

        If Debugger.IsAttached Then ' If UCase(My.Application.Info.DirectoryPath) Like "C:\VS\*" Then
            ASCMAIN1.Running_in_VS = True
            folder_prefix = "\..\..\..\..\"
            ASCMAIN1.CLIENT = "AHA" 'UCase(Mid(My.Application.Info.DirectoryPath, 7, 3))
        Else
            ASCMAIN1.Running_in_VS = False
            folder_prefix = "\..\"
            ASCMAIN1.CLIENT = "AHA" 'UCase(Split(My.Application.Info.DirectoryPath, "\")(3))
        End If



        ASCMAIN1.Folders.Add("Images", ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Images\"))
        ASCMAIN1.Folders.Add("Reports", ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Reports\"))
        ASCMAIN1.Folders.Add("DataSets", ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "DataSets\"))
        ASCMAIN1.Folders.Add("Temp", ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Temp\"))
        ASCMAIN1.Folders.Add("Work", ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Work\"))
        ASCMAIN1.Folders.Add("bin", ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "bin\"))
        ASCMAIN1.Folders.Add("Help", ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Help\"))
        ASCMAIN1.Folders.Add("Archive", ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Archive\"))
        ASCMAIN1.Folders.Add("Attach", ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix & "Attach\"))
        ASCMAIN1.Folders.Add("root", ASCMAIN1.GetPath(My.Application.Info.DirectoryPath & folder_prefix))
        ASCMAIN1.Folders.Add("SharedRoot", "R:\" & ASCMAIN1.CLIENT & "\")

        If My.Computer.Name = "WJZ64B" Then
            ASCMAIN1.Folders.Add("Oracle", "C:\oracle\product\11.2.0\dbhome_1\")
        Else
            ASCMAIN1.Folders.Add("Oracle", "C:\oracle\product\11.2.0\Client_1\")
        End If



        ASCMAIN1.Set_DBS_Dependent_Strings()

        Dim Title As String = My.Application.Info.Title

        SELECTION_NO = g.THREAD_NO + 1 ' REGISTER_FORM
        FORM_INSTANCE_NO = ASCMAIN1.Next_Control_No(MENU_ITEM_OBJECT & ".FORM_INSTANCE_NO")

        If TABLE_NAME = "" Then
            If MENU_ITEM_FORM <> "" Then
                TABLE_NAME = MENU_ITEM_FORM
            Else
                TABLE_NAME = MENU_ITEM_OBJECT
            End If
            'TABLE_NAME = Me.Name
            If TABLE_NAME = "" Then
                TABLE_NAME = Me.Name
            End If
            Mid$(TABLE_NAME, 3, 1) = "T"
        End If

        Create_TDA(dst.Tables.Add, "ASTAUDT1", "*")
        Create_TDA(dst.Tables.Add, "ASTOPST2", "*")

        ' ************************* LOG ON SUCCESSFUL ************************

        If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
            ASCMAIN1.sql = "Select GETDATE()"
        Else
            ASCMAIN1.sql = "Select SYSDATE from DUAL"
        End If
        ASCMAIN1.oraCmd.CommandText = ASCMAIN1.sql
        Dim db_datetime As Date = ASCMAIN1.oraCmd.ExecuteScalar
        ASCMAIN1.NowTSD = db_datetime.Subtract(Now)

        DATETIME_STAMP = Now + ASCMAIN1.NowTSD

        ASCMAIN1.USER_ID = g.USER_ID
        '    ASCMAIN1.USER_PASSWORD = txtUSER_PASSWORD.Text

        ASCMAIN1.sql = "Select * from ASTSECK1"
        ASCMAIN1.tblASTSECK1 = ASCDATA1.GetDataTable

        If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
            ASCMAIN1.USER_SECURITY_CODEs = ""
            ASCMAIN1.sql = "Select SECURITY_CODE from ASTUSER2 where USER_ID = '" & ASCMAIN1.USER_ID & "'"
            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                ASCMAIN1.USER_SECURITY_CODEs &= "," & row.Item(0)
            Next
            If ASCMAIN1.USER_SECURITY_CODEs <> "" Then
                ASCMAIN1.USER_SECURITY_CODEs = ASCMAIN1.USER_SECURITY_CODEs.Substring(1)
            End If
        Else
            ASCMAIN1.sql = ASCMAIN1.Flattened_List("USER_ID", "SECURITY_CODE", "ASTUSER2", ",", "USER_ID = '" & ASCMAIN1.USER_ID & "'")
            Dim tblASTUSER2s As DataTable
            tblASTUSER2s = ASCDATA1.GetDataTable(ASCMAIN1.sql)
            If tblASTUSER2s.Rows.Count = 0 Then
                ASCMAIN1.USER_SECURITY_CODEs = ""
            Else
                ASCMAIN1.USER_SECURITY_CODEs = tblASTUSER2s.Rows(0).Item(1)
            End If
        End If

        ASCMAIN1.SESSION_NO = ASCMAIN1.Next_Control_No("ASTLOGS1.SESSION_NO")
        If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
            ASCMAIN1.DBS_SESSION_ID = 1
        Else
            Dim rowSession As DataRow = ASCDATA1.GetDataRow("Select UserEnv('SESSIONID'), UserEnv('TERMINAL') from DUAL")
            ASCMAIN1.DBS_SESSION_ID = rowSession.Item(0)
        End If
        ASCMAIN1.COMPUTER_NAME = My.Computer.Name

        ' Kill_Ghost_Sessions()  ' DO WE NEED THIS?
        ASCMAIN1.Get_Current_YP()

        ASCMAIN1.sql = "Select * from ASTPARM1 where AS_PARM_KEY = 'Z'"
        Dim tblASTPARM1 As DataTable = ASCDATA1.GetDataTable
        ASCMAIN1.rowASTPARM1 = tblASTPARM1.Rows(0)


        If Not ASCMAIN1.Running_in_VS Then
            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then ' necessary because V1 uses G: drive in ASTPARM1
                ASCMAIN1.Folders("Archive") = "R:\VDI\ARCHIVE\" & ASCMAIN1.DBS_COMPANY & "\"
                ASCMAIN1.Folders("Attach") = "R:\VDI\ATTACH\" & ASCMAIN1.DBS_COMPANY & "\"
                If Not ASCMAIN1.Running_in_VS Then
                    ASCMAIN1.Folders("Reports") = "R:\VDI\REPORTS\"
                End If
                ASCMAIN1.Folders("SharedRoot") = "R:\VDI\"
            Else
                ASCMAIN1.Folders("Archive") = ASCMAIN1.rowASTPARM1("AS_PARM_ARCHIVE_FOLDER") & "\" & ASCMAIN1.DBS_COMPANY & "\"
                ASCMAIN1.Folders("Attach") = ASCMAIN1.rowASTPARM1("AS_PARM_ATTACHMENT_FOLDER") & "\" & ASCMAIN1.DBS_COMPANY & "\"
                ASCMAIN1.Folders("SharedRoot") = ASCMAIN1.rowASTPARM1("AS_PARM_SHARED_ROOT_FOLDER") & "\"
            End If

            If ASCMAIN1.DBS_SERVER = "ANE" Or ASCMAIN1.DBS_COMPANY = "ANE" Then ' necessary because V1 uses G: drive in ASTPARM1
                ASCMAIN1.Folders("Archive") = "G:\EXP\ARCHIVE\" & ASCMAIN1.DBS_COMPANY & "\"
                ASCMAIN1.Folders("Attach") = "G:\EXP\ATTACH\" & ASCMAIN1.DBS_COMPANY & "\"
                If Not ASCMAIN1.Running_in_VS Then
                    ASCMAIN1.Folders("Reports") = "G:\EXP\REPORTS\"
                End If
                ASCMAIN1.Folders("SharedRoot") = "G:\EXP\"

            End If
        Else
            ASCMAIN1.Folders("Archive") = ASCMAIN1.Folders("Archive") & ASCMAIN1.DBS_COMPANY & "\"
            ASCMAIN1.Folders("Attach") = ASCMAIN1.Folders("Attach") & ASCMAIN1.DBS_COMPANY & "\"
            ASCMAIN1.Folders("SharedRoot") = ASCMAIN1.rowASTPARM1("AS_PARM_SHARED_ROOT_FOLDER") & "\"
        End If

        ASCMAIN1.Temp_Table_Cleanup()

        ASCMAIN1.tblASTFFMT1 = ASCDATA1.GetDataTable("*", "ASTFFMT1")

        Dim INIT_DATE As Date = Now + ASCMAIN1.NowTSD

        Dim tblASTOPST1 As New DataTable
        With ASCDATA1.GetDataAdapter(tblASTOPST1, "ASTOPST1", "*", True, -1, False)
            Dim rowASTOPST1 As DataRow = tblASTOPST1.NewRow
            rowASTOPST1.Item("USER_ID") = ASCMAIN1.USER_ID
            rowASTOPST1.Item("SESSION_NO") = ASCMAIN1.SESSION_NO
            rowASTOPST1.Item("INIT_DATE") = INIT_DATE
            rowASTOPST1.Item("YYYYPP") = ASCMAIN1.CYP
            rowASTOPST1.Item("SELECTION_NO") = 0
            rowASTOPST1.Item("RE_XNO") = 0
            rowASTOPST1.Item("PRD_CLOSE_IND") = ASCMAIN1.EOM
            rowASTOPST1.Item("FORM_INSTANCE_NO") = ASCMAIN1.Next_Control_No("ASFLOGON.FORM_INSTANCE_NO")
            tblASTOPST1.Rows.Add(rowASTOPST1)
            .Update(tblASTOPST1)
            .Dispose()
        End With

        Dim tblASTLOGS1 As New DataTable
        With ASCDATA1.GetDataAdapter(tblASTLOGS1, "ASTLOGS1", "*", True, -1, False)
            Dim rowASTLOGS1 As DataRow = tblASTLOGS1.NewRow
            rowASTLOGS1.Item("SESSION_NO") = ASCMAIN1.SESSION_NO
            rowASTLOGS1.Item("USER_ID") = ASCMAIN1.USER_ID
            rowASTLOGS1.Item("SESSION_ID") = ASCMAIN1.DBS_SESSION_ID
            rowASTLOGS1.Item("COMPUTER_NAME") = ASCMAIN1.COMPUTER_NAME
            rowASTLOGS1.Item("DATE_LOGGED_ON") = INIT_DATE
            rowASTLOGS1.Item("SESSION_STATUS") = "A"
            tblASTLOGS1.Rows.Add(rowASTLOGS1)
            .Update(tblASTLOGS1)
            .Dispose()
        End With


        ' WTS Session ID

        ASCMAIN1.WTS_SESSION_ID = GetSessionId()
        ASCMAIN1.EncryptionKey = "0ff1c3" & ASCMAIN1.DBS_COMPANY

    End Sub

    Declare Function ProcessIdToSessionId Lib "kernel32.dll" (ByVal dwProcessId As Int32, ByRef pSessionId As Int32) As Int32

    Public Function GetSessionId() As Int32
        Dim _currentProcess As Process = Process.GetCurrentProcess()
        Dim _processID As Int32 = _currentProcess.Id
        Dim _sessionID As Int32
        Dim _result As Boolean = ProcessIdToSessionId(_processID, _sessionID)
        Return _sessionID
    End Function

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

        If clsASCBASE1 Is Nothing Then Return Nothing

        cdr = clsASCBASE1.LookUp(TABLE_NAME, KEY, Return_Empty_Row_if_Missing)
        Return cdr

    End Function

    Public Function LookUp( _
    ByVal TABLE_NAME As String, _
    ByVal KEY As String, _
    Optional ByVal Return_Empty_Row_if_Missing As Boolean = False) As DataRow
        Return LookUp(TABLE_NAME, New String() {KEY}, Return_Empty_Row_if_Missing)

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

        clsASCBASE1.Create_TDA _
        (tbl, TABLE_NAME, sql_custom, NumberOfKeysUsedToSelect, for_update, _
         custom_parameters, Key_Field_Count, Update_COLUMN_NAMEs, SCHEMA)
    End Sub

    Public Sub Get_PARM(ByVal PARM_TABLE_NAME As String)
        'If Not ROWs.ContainsKey(PARM_TABLE_NAME) Then
        '    Create_Lookup(PARM_TABLE_NAME)
        'End If
        ROWs(PARM_TABLE_NAME) = LookUp(PARM_TABLE_NAME, "Z")
    End Sub

    Function Fill_Record( _
    ByVal TABLE_NAME As String, _
    ByVal KEY_VALUE As Object, _
    Optional ByVal create_row_if_non_existent As Boolean = False, _
    Optional ByVal ClearBeforeFilling As Boolean = True) As DataRow
        If KEY_VALUE Is Nothing Then
            cdr = clsASCBASE1.Fill_Record(TABLE_NAME, , create_row_if_non_existent, ClearBeforeFilling)
        Else
            cdr = clsASCBASE1.Fill_Record(TABLE_NAME, New Object() {KEY_VALUE}, create_row_if_non_existent, ClearBeforeFilling)
        End If
        Return cdr
    End Function

    Function Fill_Record( _
    ByVal TABLE_NAME As String, _
    Optional ByVal Parameters() As Object = Nothing, _
    Optional ByVal create_row_if_non_existent As Boolean = False, _
    Optional ByVal ClearBeforeFilling As Boolean = True) As DataRow

        Return clsASCBASE1.Fill_Record _
        (TABLE_NAME, Parameters, create_row_if_non_existent, ClearBeforeFilling)

    End Function

    Function Fill_Records( _
    ByVal TABLE_NAME As String, _
    ByVal KEY_VALUE As String, _
    Optional ByVal ClearBeforeFilling As Boolean = True, _
    Optional ByVal Temp_Select As String = "", _
    Optional ByVal tblSubstitute As DataTable = Nothing) As Integer

        If KEY_VALUE = "" Then
            Return clsASCBASE1.Fill_Records(TABLE_NAME, , ClearBeforeFilling, Temp_Select, tblSubstitute)
        Else
            Return clsASCBASE1.Fill_Records(TABLE_NAME, New String() {KEY_VALUE}, ClearBeforeFilling, Temp_Select, tblSubstitute)
        End If

    End Function

    Function Fill_Records( _
    ByVal TABLE_NAME As String, _
    Optional ByVal Parameters() As Object = Nothing, _
    Optional ByVal ClearBeforeFilling As Boolean = True, _
    Optional ByVal Temp_Select As String = "", _
    Optional ByVal tblSubstitute As DataTable = Nothing) As Integer

        Return clsASCBASE1.Fill_Records _
        (TABLE_NAME, Parameters, ClearBeforeFilling, Temp_Select, tblSubstitute)

    End Function

    ''' <summary>
    ''' Performs the Update Command for the Table Data Adapter.  Use sql_Delete only if you want to delete rows (using the supplied sql_Delete statement) and then have the current rows Inserted.
    ''' </summary>
    ''' <param name="TABLE_NAME"></param>
    ''' <param name="sql_Delete">Specify a complete Delete statement, or else just the where clause.  If sql_delete does not begin with the word 'Delete', then the clause 'Delete from {TABLE_NAME} where ' will be pre-pended to the clause supplied.</param>
    ''' <remarks></remarks>
    Sub Update_Record_TDA(ByVal TABLE_NAME As String, Optional ByVal sql_Delete As String = "")

        If dst.Tables.Contains(TABLE_NAME) Then
            'If AUDIT.ContainsKey(TABLE_NAME) Then
            '    WriteAuditTrail(TABLE_NAME)
            'End If
            clsASCBASE1.Update_Record_TDA(TABLE_NAME, sql_Delete)
        End If
    End Sub

    Sub Create_BAs(ByVal TABLE_NAME As String)
        clsASCBASE1.Create_BAs(TABLE_NAME)
    End Sub

    Sub Clear_dst()
        clsASCBASE1.dst = New DataSet

        Setup_DataLayer()

        Dim DATASET_NAME As String = Me.Name ' & "dst"
        Mid$(DATASET_NAME, 3, 1) = "D"
        dst.DataSetName = DATASET_NAME
    End Sub

    Sub Setup_DataLayer()
        dst = clsASCBASE1.dst
        TBLs = clsASCBASE1.TBLs
        TDAs = clsASCBASE1.TDAs
        pROWs = clsASCBASE1.pROWs
        DVWs = clsASCBASE1.DVWs
        TBL_SCHEMAs = clsASCBASE1.TBL_SCHEMAs
        ROWs = clsASCBASE1.ROWs
        CMDs = clsASCBASE1.CMDs
        BA_CMDs = clsASCBASE1.BA_CMDs
    End Sub

    Sub Update_BAs(ByVal TABLE_NAME As String)
        clsASCBASE1.Update_BAs(TABLE_NAME)
    End Sub

    Sub EnforceConstraints(Optional ByVal enforce As Boolean = True)

        Dim eMessage As String = clsASCBASE1.EnforceConstraints(enforce)
        If eMessage <> "" Then
            'MsgBox(eMessage, MsgBoxStyle.OkOnly, "ABSolution will Terminate")
            Throw New Exception(eMessage)
        End If
    End Sub

    Sub BeginTrans(Optional ByVal Begin_Message As String = "")
        ASCMAIN1.T = ASCMAIN1.oraCon.BeginTransaction
    End Sub

    Sub CommitTrans(Optional ByVal Commit_Message As String = "")
        ASCMAIN1.T.Commit()
    End Sub

    Function Logon_Attempt_Succeeded() As Boolean
        Logon_Attempt_Succeeded = False

        If ASCMAIN1.DBS_PASSWORD <> "" Then
            Try
                If ASCMAIN1.oraCon.State = ConnectionState.Open Then
                    ASCMAIN1.oraCon.Close()
                End If

                Dim DEVELOPMENT_MACHINE_TNS As String = "(DESCRIPTION =(ADDRESS_LIST =(ADDRESS = (PROTOCOL = TCP)(HOST = localhost)(PORT = 1521)))(CONNECT_DATA = (SERVICE_NAME = orcl)))"
                DEVELOPMENT_MACHINE_TNS = ""

                If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                    ASCMAIN1.oraCon.ConnectionString = "Data Source=" & IIf(ASCMAIN1.DBS_SERVER = "", ".", ASCMAIN1.DBS_SERVER) & ";Initial Catalog=" & ASCMAIN1.DBS_COMPANY & "; " & IIf(ASCMAIN1.DBS_SERVER = "", "User ID='ODG'", "User ID='sa';Password='0ff1c3';") & ";Integrated Security=" & IIf(ASCMAIN1.DBS_SERVER = "", "True", "False") & ";MultipleActiveResultSets=True"
                Else
                    ASCMAIN1.oraCon.ConnectionString = "Data Source=" & IIf(ASCMAIN1.DBS_SERVER = "", DEVELOPMENT_MACHINE_TNS, ASCMAIN1.DBS_SERVER) & ";User ID=" & ASCMAIN1.DBS_COMPANY & ";Password=" & ASCMAIN1.DBS_PASSWORD & ";pooling=false"
                End If

                ASCMAIN1.oraCon.Open()
                ASCMAIN1.oraCmd = ASCMAIN1.oraCon.CreateCommand

                ASCMAIN1.oraSP.CommandType = CommandType.StoredProcedure
                ASCMAIN1.oraSP.Connection = ASCMAIN1.oraCon

                Logon_Attempt_Succeeded = True

                Dim myWorkstation As String = System.Net.Dns.GetHostName()
                Dim IPAddress As String = _
                System.Net.Dns.GetHostEntry(myWorkstation).AddressList(0).ToString()
                ASCMAIN1.DBS_IP_ADDRESS = IPAddress
                ASCMAIN1.DBS_SERVER_NAME = myWorkstation

            Catch ex As Exception
            End Try
        End If

    End Function

    Function Create_Relation( _
    ByVal TABLE_NAME_parent As String, _
    ByVal TABLE_NAME_child As String, _
    ByVal COLUMN_NAMEsParent As String, _
    Optional ByVal COLUMN_NAMEsChild As String = "")

        Dim R As DataRelation = ASCDATA1.GetRelation _
        (dst, TABLE_NAME_parent, TABLE_NAME_child, COLUMN_NAMEsParent, COLUMN_NAMEsChild)

        dst.Relations.Add(R)

        Return R
    End Function

    Public Function Create_Temporary_Table(TABLE_NAME As String, keyCOLUMN_NAMEs As String) As String

        ASCMAIN1.sql = "Select * from " & TABLE_NAME & " where ROWNUM < 1"
        Dim TABLE_NAME_temp As String = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCDATA1.ExecuteSQL("Alter Table " & TABLE_NAME_temp & " Add Primary Key (" & keyCOLUMN_NAMEs & ")")
        ASCMAIN1.sql = "Select * from " & TABLE_NAME_temp & " where ROWNUM < 1"
        Create_TDA(dst.Tables.Add(TABLE_NAME), TABLE_NAME_temp, ASCMAIN1.sql, 0)

        Return TABLE_NAME_temp
    End Function
End Class

Public Class GunEnvironment
    Inherits ABSEnvironment
    'Public DBS_SERVER As String
    'Public DBS_COMPANY As String
    'Public DBS_PASSWORD As String

    'Public THREAD_NO As Integer
    'Public APP_ID As String
    'Public APP_DESC As String
    'Public USER_ID As String
    Public GUN_LOC As String
    Public PICK_TYPE As String
    Public WHSE_CODE As String

    ' Instructions for the App

    'Public APP_CMD As String
    'Public APP_KEY As String
End Class

Public Class ABSEnvironment
    Public Property DBS_SERVER As String
    Public Property DBS_COMPANY As String
    Public Property DBS_PASSWORD As String

    Public Property THREAD_NO As Integer
    Public Property APP_ID As String
    Public Property APP_DESC As String
    Public Property USER_ID As String


    Public Property TempDirectory As String
    Public Property ReportsDirectory As String
    Public Property ArchiveDirectory As String
    Public Property FormName As String

    Public Property XNO As String

    Public Property APP_CMD As String
    Public Property APP_KEY As String
    Public Property LP_CODE As String
    Public Property CLIENT As String
End Class
