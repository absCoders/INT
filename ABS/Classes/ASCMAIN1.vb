'Imports System.Windows.Forms
Imports System.Math
Imports Infragistics.Win
Imports Oracle.ManagedDataAccess.Client
'Imports CrystalDecisions.CrystalReports
Imports System.IO
Imports System.Net.Http.Headers
Imports System.Net.Http.Formatting
Imports System.Net.Http
'Imports System.Text
'Imports Microsoft.Exchange.WebServices.Data
Imports Microsoft.Identity.Client
Imports System.Configuration
Imports System.Timers

Public Class ASCMAIN1

    ' Application
    Public Shared FORM_NAME As String
    Public Shared SET_ID As String
    Public Shared ActiveForm As ASFBASE1 '  Windows.Forms.Form
    Public Shared MainForm As Windows.Forms.Form
    Public Shared MainForm_pgd As Windows.Forms.PropertyGrid
    Public Shared ABS_FORMS() As System.Windows.Forms.Form
    Public Shared Folders As New Dictionary(Of String, String)

    ' General Purpose
    Public Shared sql As String
    Public Shared CodeSelector As New ASCCODE1
    Public Shared rowASTPARM1 As DataRow
    Public Shared NowTSD As TimeSpan
    Public Shared grdRows As New Collection
    Public Shared response As Integer
    Public Shared Message As String

    ' User Logged In
    Public Shared USER_ID As String
    Public Shared USER_NAME As String
    Public Shared USER_PASSWORD As String
    Public Shared USER_ADMIN As String
    Public Shared USER_CODES As String
    Public Shared USER_SECURITY_CODEs As String
    Public Shared USER_MENU_ITEM_OBJECT As String
    Public Shared SESSION_NO As String
    Public Shared USER_EMAIL As String
    Public Shared WTS_SESSION_ID As Int32
    Public Shared MENU_ITEM_OBJECTs As List(Of String)
    Public Shared client_schema As String = ""

    Public Shared JOB_STREAM_CODE As String = String.Empty
    Public Shared JOB_STREAM_TYPE As String = String.Empty
    Public Shared JOB_STREAM_FORM_NAME As String = String.Empty
    Public Shared JOB_STREAM_XNO As String = String.Empty
    Public Shared JOB_STREAM_LNO As Int16 = 0
    Public Shared JOB_STREAM_EXPORT_FOLDER As String = String.Empty
    Public Shared JOB_STREAM_EXPORT_OPTIONS As String = String.Empty

    Public Shared JOB_STREAM_REPORT_FILES As New Dictionary(Of String, String)
    Public Shared JOB_STREAM_REPORT_XNO As String = ""

    ' DB & Schema
    'Public Shared DBS As String
    'Public ABS_Data As New ABS_Data.ABS_DataAdapter

    Public Shared DBS_SERVER As String
    Public Shared DBS_COMPANY As String
    Public Shared DBS_PASSWORD As String
    Public Shared DBS_CONCAT As String
    Public Shared DBS_PARAMETER As String
    Public Shared DBS_NOROWS As String
    Public Shared DBS_OBJPFX As String
    Public Shared DBS_OBJSFX As String
    Public Shared DBS_SESSION_ID As Long
    Public Shared DBS_IP_ADDRESS As String
    Public Shared DBS_SERVER_NAME As String
    Public Shared DBS_ENCRYPTION_KEY As String

    Public Shared authResult As AuthenticationResult
    Public Shared authResult_counter As Integer
    Public Shared authResult_timestamp As Date
    Public Shared WithEvents absTimer As Timer = New Timer(2000)

    ' Web API
    Public Shared USER_JWT As String
    Public Shared USER_JWT_EXPIRES As DateTime
    Public Shared API_ENDPOINT As String

    Public Shared SEXE_NO As String
    Public Shared rowASTSEXE1 As DataRow

    Public Enum DBS_TYPE_types
        Oracle
        SQLServer
    End Enum

    Public Enum ExecuteSQL_types
        Build_and_Execute
        Build_Only
        Execute_Only
    End Enum

    Public Shared DBS_TYPE As DBS_TYPE_types
    Public Shared FilledSchemas As New Dictionary(Of String, DataTable)

    ' DB Connection
    Public Shared oraCon As New OracleConnection
    Public Shared oraCmd As New OracleCommand
    Public Shared oraSP As New OracleCommand    ' Used to execute Stored Procedures
    Public Shared T As OracleTransaction
    Public Shared oraAda As New OracleDataAdapter

    ' Universe
    Public Shared COMPUTER_NAME As String
    Public Shared ABSWEB As Boolean
    Public Shared APP_PATH As String
    Public Shared MAIN_FOLDER As String
    Public Shared ASSEMBLY_NAME As String
    Public Shared VERSION_NO As String
    Public Shared Running_in_VS As Boolean
    Public Shared developerMode As Boolean = False
    Public Shared testMode As Boolean = False
    ' there is a menu option which will toggle whether ABS is running in design mode or run mode (like to set the explorerbar item properties)
    ' need a 3rd setting to conveniently know if the user logged in is a designer (ie, has SY)

    Public Shared developerModeOptions As New ASCDEVMO
    Public Shared testModeOptions As New ASCTSTF1

    Public Shared GridDoubleClickAllowed As Boolean = True

    Public Shared SOLUTION As String
    Public Shared CLIENT As String
    Public Shared EncryptionKey As String = String.Empty

    ' Crystal Reports
    '    Public Shared CR_RPT As New CrystalDecisions.CrystalReports.Engine.ReportClass
    Public Shared CR_RPT As CrystalDecisions.CrystalReports.Engine.ReportDocument
    Public Shared CR_SubRpt As CrystalDecisions.CrystalReports.Engine.ReportDocument

    ' Excel
    'Public Shared xls_File As New GemBox.Spreadsheet.ExcelFile
    'Public Shared xls_Sheet As GemBox.Spreadsheet.ExcelWorksheet
    'Public Shared xls_Row As GemBox.Spreadsheet.ExcelRow
    'Public Shared xls_Cell As GemBox.Spreadsheet.ExcelCell

    ' Environment
    Public Shared tblASTFFMT1 As New DataTable
    Public Shared tblASTSECK1 As New DataTable
    Public Shared tblASTSQLX1 As DataTable
    Public Shared dstASTVIEWS As New DataSet
    Public Shared MRUs As New Dictionary(Of String, List(Of String))
    Public Shared MRU_COLUMN_NAME As String
    Public Shared MRU_txtctl As UltraWinEditors.UltraTextEditor
    Public Shared MRU_cmbctl As UltraWinGrid.UltraCombo
    Public Shared MRU_used As Boolean
    Public Shared TACMAIN1 As TACMAIN1
    Public Shared ABS_Assemblies As New Dictionary(Of String, System.Reflection.Assembly)
    Public Shared tblASTMENU1 As New DataTable
    Public Shared tblASTAESC1 As New DataTable

    ' Period Context
    Public Shared CYW As String     ' Current Year and Week, YYYYWW
    Public Shared CYP As String     ' Current Year and Period, YYYYPP
    Public Shared CYM As String     ' Calendar Year and Month of CYP, YYYYMM
    Public Shared PCO As Integer    ' Period Calendar Offset to Regular Calendar, 1 = Year ends in Jan, 2 = Year ends in Feb, 0 = Year ends in Dec
    Public Shared WCO As Integer    ' Week Calendar Offset to Regular Calendar
    Public Shared EOM As String     ' 1 = Period End is in Process
    Public Shared Timer As Date

    Public Shared MENU_ID As String
    Public Shared MENU_ITEM_TYPE As String
    Public Shared MENU_ITEM_OBJECT As String

    Public Shared Sub Set_DBS_Dependent_Strings()
        If ASCMAIN1.oraCon.GetType.ToString() = "System.Data.SqlClient.SqlConnection" Then
            ASCMAIN1.DBS_TYPE = DBS_TYPE_types.SQLServer
            ASCMAIN1.DBS_CONCAT = "+"
            ASCMAIN1.DBS_PARAMETER = "@"
            ASCMAIN1.DBS_NOROWS = "1 <> 1"
            ASCMAIN1.DBS_OBJPFX = "["
            ASCMAIN1.DBS_OBJSFX = "]"
        Else
            ASCMAIN1.DBS_TYPE = DBS_TYPE_types.Oracle
            ASCMAIN1.DBS_CONCAT = "||"
            ASCMAIN1.DBS_PARAMETER = ":"
            ASCMAIN1.DBS_NOROWS = "ROWNUM < 1"
            ASCMAIN1.DBS_OBJPFX = Chr(34)
            ASCMAIN1.DBS_OBJSFX = Chr(34)
        End If
    End Sub

    Public Shared Sub AutoFitGridColumns(ByVal grd As UltraWinGrid.UltraGrid, Optional ByVal band As Integer = -1, Optional ByVal nRows As Integer = 10)
        Try
            Dim startBand As Integer = band
            Dim endBand As Integer = band

            If band = -1 Then
                startBand = 0
                endBand = grd.DisplayLayout.Bands.Count - 1
            End If

            For iBand As Integer = startBand To endBand
                For Each col As Infragistics.Win.UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(iBand).Columns
                    col.PerformAutoResize(nRows)
                Next
            Next

        Catch ex As Exception
            ' do not want to throw error is user types in an invalid band number
        End Try
    End Sub


    Public Shared Sub Center(ByVal F As Form)
        F.SetBounds((System.Windows.Forms.Screen.GetBounds(F).Width / 2) - (F.Width / 2),
    (System.Windows.Forms.Screen.GetBounds(F).Height / 2) - (F.Height / 2),
    F.Width, F.Height, System.Windows.Forms.BoundsSpecified.Location)
    End Sub

    Public Shared Function Make_Caption(ByVal COLUMN_NAME As String) As String
        Dim f As String
        Dim j As Integer
        f = LCase(COLUMN_NAME)
        If f <> "" Then
            Mid$(f, 1, 1) = UCase(Mid$(f, 1, 1))
            Do While InStr(f, "_") <> 0
                j = InStr(f, "_")
                Mid$(f, j, 1) = " "
                If Len(f) > j Then
                    Mid$(f, j + 1, 1) = UCase(Mid$(f, j + 1, 1))
                End If
            Loop
        End If
        Make_Caption = f
    End Function

    ''' <summary>
    ''' How_Many is (optionally) how many control numbers you want,
    ''' default = 1, do not use this unless you want >1
    ''' you will need to format the control no to handle wrap around
    ''' </summary>
    ''' <param name="CTL_NO_TYPE"></param>
    ''' <param name="How_Many"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Next_Control_No(
    ByVal CTL_NO_TYPE As String,
    Optional ByVal How_Many As Long = 1) As String

        Dim CTL_NO_LAST As Long
        Dim CTL_NO_LENGTH As Integer

        'If CTL_NO_TYPE = "ARFCINQ1.FORM_INSTANCE_NO" Then
        '    Next_Control_No = ASCDATA1.ExecuteSF("TAPSEQN1", New String() {"CTL_NO_TYPE_IN"}, New Object() {CTL_NO_TYPE})
        'Else


        Dim tblTATCTLN1 As New DataTable
        Dim sql As String = "Select * from TATCTLN1 where CTL_NO_TYPE = '" & CTL_NO_TYPE & "'"
        If DBS_TYPE = DBS_TYPE_types.SQLServer Then
        Else
            sql &= " FOR UPDATE"
        End If
        With ASCDATA1.GetDataAdapter(tblTATCTLN1, "TATCTLN1", sql, True)

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
                If InStr(CTL_NO_TYPE, ".") <> 0 Then

                    Dim sqlCTL_NO_LENGTH As String =
                        "Select DATA_LENGTH from USER_TAB_COLUMNS " _
                        & " where TABLE_NAME = '" & Split(CTL_NO_TYPE, ".")(0) & "' and COLUMN_NAME = '" _
                        & Split(CTL_NO_TYPE, ".")(1) & "'"

                    If DBS_TYPE = DBS_TYPE_types.SQLServer Then
                        sqlCTL_NO_LENGTH = "SELECT CHARACTER_MAXIMUM_LENGTH " _
                        & " FROM INFORMATION_SCHEMA.COLUMNS " _
                        & " where TABLE_CATALOG = '" & ASCMAIN1.DBS_COMPANY & "' " _
                        & " and TABLE_NAME = '" & Split(CTL_NO_TYPE, ".")(0) & "'" _
                        & " and COLUMN_NAME = '" & Split(CTL_NO_TYPE, ".")(1) & "'"
                    End If

                    CTL_NO_LENGTH = Val(ASCDATA1.GetDataValue(sqlCTL_NO_LENGTH) & "")
                End If
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

            Next_Control_No = Format$(CTL_NO_LAST, "".PadLeft(CTL_NO_LENGTH, "0"))
            CTL_NO_LAST = CTL_NO_LAST + How_Many - 1
            If CTL_NO_LAST >= 10 ^ CTL_NO_LENGTH Then
                CTL_NO_LAST = CTL_NO_LAST - 10 ^ CTL_NO_LENGTH + 1
            End If
            rowTATCTLN1.Item("CTL_NO_LAST") = CTL_NO_LAST
            Dim SUCCESS As Integer = .Update(tblTATCTLN1)
            .Dispose()
        End With

        Dim S As String = "SYSDATE"
        If DBS_TYPE = DBS_TYPE_types.SQLServer Then
            S = "GETDATE()"
        End If


        ASCDATA1.ExecuteSQL("Insert into TATCTLN2 " _
                        & " (CTL_NO_TYPE,CTL_NO_LAST,CTL_NO_KEY,HOW_MANY,INIT_DATE,INIT_OPER) " _
                        & " VALUES (:PARM1,:PARM2,:PARM3,:PARM4," & S & ",:PARM5)" _
                        , "VNVNV" _
                        , New Object() {CTL_NO_TYPE, CTL_NO_LAST, Next_Control_No, How_Many, ASCMAIN1.USER_ID})
        'End If

    End Function

    ''' <summary>
    ''' Pass in a date (as a string, date, or variant) and number of week days
    ''' to calculate the date which is number_of_week_days from the base_date
    ''' if you pass in base_date of Sat or Sun, then then base_date will be 
    ''' considered to be the subsequent Mon
    ''' </summary>
    ''' <param name="base_date"></param>
    ''' <param name="number_of_week_days">Number of Week Days may be positive or negative</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DateDiff_Weekday(
    ByVal base_date As Object,
    ByVal number_of_week_days As Integer) As Object

        Dim i As Integer
        Dim working_date As Date
        Dim full_weeks As Integer   ' number of full weeks represented in number_of_weekdays
        Dim days As Integer         ' number of days beyond full_weeks
        Dim direction As Integer    ' +1 for calcs into the future, -1 for into the past

        direction = Sign(number_of_week_days)
        full_weeks = Int(Abs(number_of_week_days) / 5)
        days = Abs(number_of_week_days) - full_weeks * 5

        working_date = base_date
        Call DateDiff_Weekday_Adjust_for_Weekend(working_date, direction)

        working_date = DateAdd("d", direction * full_weeks * 7, working_date)
        If days <> 0 Then
            For i = 1 To days
                working_date = DateAdd("d", direction, working_date)
                Call DateDiff_Weekday_Adjust_for_Weekend(working_date, direction)
            Next i
        End If

        Return working_date
    End Function
    Public Shared Sub DateDiff_Weekday_Adjust_for_Weekend(
    ByRef working_date As Date,
    ByVal direction As Integer)

        If Weekday(working_date) = 1 Then ' Sunday
            If direction = 1 Then
                working_date = DateAdd("d", 1, working_date)
            Else
                working_date = DateAdd("d", -2, working_date)
            End If
        End If
        If Weekday(working_date) = 7 Then ' Saturday
            If direction = 1 Then
                working_date = DateAdd("d", 2, working_date)
            Else
                working_date = DateAdd("d", -1, working_date)
            End If
        End If
        Return

    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="TABLE_NAME">Name of the DataTable to Return</param>
    ''' <param name="dt_source">Source of Data</param>
    ''' <param name="COLUMN_NAMEs">List of Columns for which Distinct Values are Requested</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Distinct_Values(
    ByVal TABLE_NAME As String,
    ByRef dt_source As DataTable,
    ByVal ParamArray COLUMN_NAMEs() As String) As DataTable

        'Dim dt As DataTable = New DataView(dt_source, RowFilter).ToTable(True, COLUMN_NAMEs)
        'If TABLE_NAME <> "" Then
        '    dt.TableName = TABLE_NAME
        'End If

        Return Distinct_Values(TABLE_NAME, "", dt_source, COLUMN_NAMEs)

    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="TABLE_NAME">Name of the DataTable to Return</param>
    ''' <param name="dt_source">Source of Data</param>
    ''' <param name="RowFilter">where clause, without the where</param>
    ''' <param name="COLUMN_NAMEs">List of Columns for which Distinct Values are Requested</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Distinct_Values(
    ByVal TABLE_NAME As String,
    ByVal RowFilter As String,
    ByRef dt_source As DataTable,
    ByVal ParamArray COLUMN_NAMEs() As String) As DataTable

        Dim dt As DataTable = New DataView(dt_source, RowFilter, "", DataViewRowState.CurrentRows).ToTable(True, COLUMN_NAMEs)
        If TABLE_NAME <> "" Then
            dt.TableName = TABLE_NAME
        End If

        Return dt

    End Function

    Public Shared Function Temp_Table(
    Optional ByVal sql As String = "",
    Optional ByVal FORM_NAME As String = "",
    Optional ByVal XNO As String = "",
    Optional ByVal custom_parameters As String = "",
    Optional ByVal PARMs() As String = Nothing) As String

        If sql = "" Then
            sql = ASCMAIN1.sql
        End If

        If ASCMAIN1.ActiveForm IsNot Nothing Then
            Try
                If FORM_NAME = "" Then FORM_NAME = ASCMAIN1.ActiveForm.MENU_ITEM_OBJECT
                If XNO = "" Then XNO = ASCMAIN1.ActiveForm.XNO
            Catch ex As Exception

            End Try
        End If

        Dim TABLE_NAME As String = "ASW" & Next_Control_No("ASTTEMP1")
        Dim S As String = "SYSDATE"
        If DBS_TYPE = DBS_TYPE_types.SQLServer Then
            S = "GETDATE()"
        End If
        Dim sqlASTTEMP1 As String = "Insert into ASTTEMP1 Values ('" & TABLE_NAME & "','" & ASCMAIN1.SESSION_NO & "', " & S & ",'" & ASCMAIN1.USER_ID & "','" & FORM_NAME & "','" & XNO & "'," & CStr(ASCMAIN1.DBS_SESSION_ID) & ")"
        ASCDATA1.ExecuteSQL(sqlASTTEMP1)

        If ASCMAIN1.DBS_TYPE = DBS_TYPE_types.SQLServer Then
            sql = Replace(sql, " from ", " into " & TABLE_NAME & " from ", , 1)
            sql = Replace(sql, "ROWNUM < 1", "1<>1")
            sql = Replace(sql, "NVL(", "ISNULL(")
        Else
            sql = "Create Table " & TABLE_NAME & " as " & sql
        End If

        If custom_parameters = "" Then
            ASCDATA1.ExecuteSQL(sql)
        Else
            ASCDATA1.ExecuteSQL(sql, custom_parameters, PARMs)
        End If

        Return TABLE_NAME
    End Function

    Public Shared Sub Temp_Table_Cleanup(
    Optional ByVal General_Cleanup As Boolean = True)

        If General_Cleanup Then
            ' this causes many slowdowns exiting ABSolution - handle it in backup procedure
            Exit Sub
        End If

        If ASCMAIN1.rowASTPARM1 Is Nothing OrElse ASCMAIN1.rowASTPARM1.Item("AS_PARM_NO_DROP_TEMP_TABLES") & "" = "1" Then
            Exit Sub
        End If

        Dim sql As String = ""
        sql = "Select * from ASTTEMP1 where (USER_ID = '" & ASCMAIN1.USER_ID & "'"
        If General_Cleanup Then
            sql = sql & " and SESSION_ID not in (Select AUDSID from V$SESSION)"
        Else
            sql = sql & " and SESSION_NO = '" & ASCMAIN1.SESSION_NO & "'"
        End If
        If DBS_PARAMETER = "@" Then
            sql = sql & ") or DATE_CREATED < GETDATE() -2"
        Else
            sql = sql & ") or DATE_CREATED < SYSDATE -2"
        End If

        Dim tblASTTEMP1 As New DataTable
        With ASCDATA1.GetDataAdapter(tblASTTEMP1, "ASTTEMP1", sql, True)

            For Each r As DataRow In tblASTTEMP1.Rows
                Try
                    Call ASCMAIN1.Progress("Now Dropping Work Tables", r.Item("TABLE_NAME"))
                    Application.DoEvents()

                    If DBS_TYPE = DBS_TYPE_types.SQLServer Then
                        sql = "Drop Table " & r.Item("TABLE_NAME")
                    Else
                        sql = "Drop Table " & r.Item("TABLE_NAME") & " Purge"
                        'purge option is taking too long
                        'sql = "Drop Table " & r.Item("TABLE_NAME")
                    End If
                    ASCDATA1.ExecuteSQL(sql, True)
                Catch ex As Exception
                    If ASCMAIN1.Running_in_VS Then
                        Stop
                    End If
                End Try
                r.Delete()
            Next
            Try
                .Update(tblASTTEMP1)
                .Dispose()
            Catch ex As Exception

            End Try
        End With

        ' Queue up tables to be dropped in 2 days if they are still around
        ' We are assuming that Oracle will analyze everything by filtering on the date last analyzed,
        ' and this gives us a little breathing from to stay away from temporary tables created recently
        ' which (for whatever reason) are not represented (yet) in ASTTEMP1

        ASCMAIN1.sql = "Select TABLE_NAME from USER_TABLES where TABLE_NAME LIKE 'ASW%' and LAST_ANALYZED < SYSDATE -2" _
            & " minus Select TABLE_NAME from ASTTEMP1"
        ASCDATA1.ExecuteSQL("Insert into ASTTEMP1 (TABLE_NAME,SESSION_NO,DATE_CREATED,USER_ID,FORM_NAME,XNO,SESSION_ID) Select TABLE_NAME, Null SESSION_NO, SYSDATE DATE_CREATED, '" & ASCMAIN1.USER_ID & "' USER_ID,'" & ASCMAIN1.USER_ID & "' FORM_NAME, Null XNO, Null SESSION_ID from (" & ASCMAIN1.sql & ")")

    End Sub

    Public Shared Sub Directory_Cleanup()

        'If we ever wanted to make this a parm option
        'If ASCMAIN1.rowASTPARM1.Item("AS_PARM_NO_DIR_CLEANUP") & "" = "1" Then
        '    Exit Sub
        'End If

        If ASCMAIN1.Running_in_VS Then
            Exit Sub
        End If

        Call ASCMAIN1.Progress("Now Cleaning Application Directories", "")

        Dim aPath As String = Application.StartupPath()
        'Temp
        If Directory.Exists(aPath & "\Temp") Then
            Directory.Delete(aPath & "\Temp")
            Directory.CreateDirectory(aPath & "\Temp")
        End If
        'Datasets
        If Directory.Exists(aPath & "\Datasets") Then
            Directory.Delete(aPath & "\Datasets")
            Directory.CreateDirectory(aPath & "\Datasets")
        End If
        'Reports
        If Directory.Exists(aPath & "\Reports") Then
            Directory.Delete(aPath & "\Reports")
            Directory.CreateDirectory(aPath & "\Reports")
        End If
        'Work
        If Directory.Exists(aPath & "\Work") Then
            Directory.Delete(aPath & "\Work")
            Directory.CreateDirectory(aPath & "\Work")
        End If

        Call ASCMAIN1.Progress("", "")

    End Sub

    Public Shared Sub Track(ByVal MSG1 As String, ByVal MSG2 As String)
        '        Dim z As String

    End Sub

    Public Shared Function Pos(ByVal S As String, ByVal h As String, ByVal T As String, ByVal l As Integer) As Integer
        Dim p As Integer
        Dim i As Integer
        Dim X As Integer

        Dim v1 As Object
        Dim v2 As Object

        p = 0
        If Len(T) Mod l = 0 Then
            X = Len(T) / l
        Else
            X = Len(T) / l + 1
        End If
        v1 = S
        If h = "=>" Then
            h = ">="
        End If
        If h = "=<" Then
            h = "<="
        End If
        For i = 1 To X
            v2 = Mid$(T, (i - 1) * l + 1)
            If Len(v2) > Len(S) Then
                v2 = Left$(v2, Len(S))
            End If
            v1 = S
            If Len(v1) > Len(v2) Then
                v1 = Left$(v1, Len(v2))
            End If
            If (InStr(h, "=") <> 0 And v1 = v2) Or (h = "<>" And v1 <> v2) Or ((h = "<" Or h = "<=") And StrComp(v1, v2, 0) = -1) Or ((h = ">" Or h = ">=") And StrComp(v1, v2, 0) = 1) Then
                p = (i - 1) * l + 1
                Exit For
            End If
        Next i
        Pos = p

    End Function
    ''' <summary>
    ''' Returns a string which is ready to be used in a SQL statement IN expression.
    ''' Passing in a CODEs value of AABBCC with a CODE_length of 2 would return 'AA','BB','CC'.
    ''' </summary>
    ''' <param name="CODEs"></param>
    ''' <param name="CODE_length"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Select_List(ByVal CODEs As String, ByVal CODE_length As Integer) As String
        Dim i As Integer
        Dim sqlCODEs As String

        sqlCODEs = ""
        If CODEs <> "" Then
            For i = 1 To Len(CODEs) / CODE_length
                sqlCODEs = sqlCODEs & "'" & Mid$(CODEs, (i - 1) * CODE_length + 1, CODE_length) & "',"
            Next i
            sqlCODEs = Mid$(sqlCODEs, 1, Len(sqlCODEs) - 1)
        End If
        Select_List = sqlCODEs
    End Function

    ''' <summary>
    ''' Returns a string which is a sorted list of Code Values, where the length of each Code Value is a fixed number of characters.
    ''' Passing in a CODEs value of BBCCAA with a CODE_length of 2 would return AABBCC.
    ''' </summary>
    ''' <param name="CODEs"></param>
    ''' <param name="CODE_length"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Sort_List(ByVal CODEs As String, ByVal CODE_length As Integer) As String
        Dim i As Integer
        Dim j As Integer
        Dim sortedCODEs As String
        Dim zz As String

        sortedCODEs = ""
        For i = 1 To Len(CODEs) / CODE_length
            zz = Mid$(CODEs, (i - 1) * CODE_length + 1, CODE_length)
            j = Pos(zz, "<", sortedCODEs & Chr(255), CODE_length)
            sortedCODEs = Mid$(sortedCODEs, 1, j - 1) & zz & Mid$(sortedCODEs, j)
        Next i
        Sort_List = sortedCODEs
    End Function
    ''' <summary>
    ''' Returns a where clause with the keyword WHERE for the first condition.
    ''' This function is useful to place the WHERE keyword in front of a dynamically constructed where clause built using ANDs only.
    ''' This function will return an empty string if an empty string was passed in.
    ''' This function will not harm a WHEREless_Clause which already begins with WHERE.
    ''' Case is not an issue.
    ''' If the clause does not begin with either an AND or a WHERE, then you get back what you passed in with a WHERE pre-pended to it.
    ''' </summary>
    ''' <param name="WHEREless_Clause"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SQL_Add_WHERE(ByVal WHEREless_Clause As String) As String
        If Trim(Replace(WHEREless_Clause, vbCrLf, "")) = "" Then
            Return ""
        End If
        Dim z As String
        z = LTrim$(WHEREless_Clause)
        If InStr(UCase$(z), "WHERE") = 1 Then
            SQL_Add_WHERE = z
        ElseIf InStr(UCase$(z), "AND") = 1 Then
            SQL_Add_WHERE = " where " & Mid$(z, 4)
        ElseIf z <> "" Then
            SQL_Add_WHERE = " where " & z
        Else
            SQL_Add_WHERE = WHEREless_Clause
        End If
    End Function

    Public Shared Function Make_Plural(ByVal singular_noun As String) As String
        Dim z As String = singular_noun
        If Right$(z, 1) = "y" Then
            z = Mid$(z, 1, Len(z) - 1) & "ies"
        ElseIf Right$(z, 2) = "ss" Then
            z = z & "es"
        ElseIf Right$(z, 1) = "s" Then
            z = z
        Else
            z = z & "s"
        End If
        Make_Plural = z
    End Function

    Public Shared Sub Get_Current_YP()
        Dim CYYYY As String
        Dim CPP As String
        Dim CWW As String
        sql = "Select * from ASTPCTL1"
        Dim tblASTPCTL1 As DataTable = ASCDATA1.GetDataTable(sql)
        Dim rowASTPCTL1 As DataRow = tblASTPCTL1.Rows(0)
        CYYYY = Format(Val(rowASTPCTL1.Item("CURR_YEAR") & ""), "0000")
        CPP = Format(Val(rowASTPCTL1.Item("CURR_PERIOD") & ""), "00")

        CWW = Format(Val(rowASTPCTL1.Item("CURR_WEEK") & ""), "00")
        If CWW = "00" Then
            sql = "Select YYYYWW from GLTPARM3 where WEEK_END_DATE >= '" & Format$(Now + NowTSD, ASCMAIN1.Get_Date_Mask) & "'"
            sql = sql & " and WEEK_END_DATE < '" & Format$(DateAdd(DateInterval.Day, 7, (Now + NowTSD)), ASCMAIN1.Get_Date_Mask) & "'"
            Dim tblGLTPARM3X As DataTable = ASCDATA1.GetDataTable(sql)
            If tblGLTPARM3X.Rows.Count = 0 Then
                CWW = Format$(Val(CPP) * 4.5, "00")
                If CWW > "52" Then
                    CWW = "52"
                End If
                CYW = CYYYY & CWW
            Else
                Dim rowGLTPARM3X As DataRow = tblGLTPARM3X.Rows(0)
                CYW = rowGLTPARM3X.Item("YYYYWW")
                CWW = Mid$(CYW, 5, 2)
            End If
        Else
            CYW = CYYYY & Format$(Val(CWW), "00")
        End If
        PCO = Val(rowASTPCTL1.Item("P01_CAL_OFFSET") & "")
        WCO = Val(rowASTPCTL1.Item("P01_CAL_OFFSET_YW") & "")
        EOM = rowASTPCTL1.Item("PRD_CLOSE_IND") & ""
        CYP = CYYYY & CPP
        CYM = Get_YYYYMM(CYP, 0)
    End Sub

    Public Shared Function Get_YYYYMM(
    ByVal YYYYPP As String,
    Optional ByVal PERIOD_OFFSET As Integer = 0)
        If YYYYPP = "" Then
            Return ""
        Else
            Return Format$(DateAdd("M", PCO + PERIOD_OFFSET, Mid$(YYYYPP, 5, 2) & "/01/" & Mid$(YYYYPP, 3, 2)), "yyyyMM")
        End If
    End Function

    Public Shared Function Get_Dates(ByVal YP1 As String) As Date()

        Dim YP0 As String = Period_Calc(YP1, -1)
        ASCMAIN1.sql = "Select OPS_YYYYPP, PRD_END_DATE from GLTPARM2 " _
        & " where OPS_YYYYPP between '" & YP0 & "'" _
        & "   and '" & YP1 & "' order by OPS_YYYYPP"
        Dim tbl As DataTable = ASCDATA1.GetDataTable
        Dim DT0 As Date = tbl.Rows.Find(YP0).Item("PRD_END_DATE")
        Dim DT1 As Date = tbl.Rows.Find(YP1).Item("PRD_END_DATE")

        Dim dates() As Date
        ReDim dates(0)
        dates(0) = DT0

        Dim ts As TimeSpan = DT1 - DT0
        For I As Integer = 1 To ts.Days
            ReDim Preserve dates(I)
            dates(I) = DT0.AddDays(I)
        Next

        Return dates
    End Function

    Public Shared Function Get_Date_Mask() As String
        Get_Date_Mask = ""
        Get_Date_Mask = "dd-MMM-yyyy"
        '        Get_Date_Mask = "MM/dd/yyyy"
    End Function


    Public Shared Sub Add_Menu_to_Tree(
        ByVal MENU_ID As String,
        ByVal KEY_PREFIX As String,
        ByVal tvw As Infragistics.Win.UltraWinTree.UltraTree,
        ByVal level As Integer,
        ByRef tblASTMENU1 As DataTable)

        Dim KEY As String
        Dim MENU_ITEM_DESC As String
        Dim IMAGE_FOLDER As String = ASCMAIN1.Folders("Images") & "ABS\Menu\Tree\"

        Dim aNode As New Infragistics.Win.UltraWinTree.UltraTreeNode

        Dim sql As String
        Dim orderby As String

        If MENU_ID = "*" Then
            sql = "MENU_ITEM_DESC_FAVORITE IS NOT NULL"
            orderby = "MENU_ITEM_DESC_FAVORITE"
        Else
            sql = "MENU_ID = '" & MENU_ID & "'"
            If MENU_ID = "MAIN" Then
                sql = sql & " and MENU_ITEM_TYPE = 'M'"
            End If
            orderby = "MENU_ITEM_SEQ"
        End If

        For Each row As DataRow In tblASTMENU1.Select(sql, orderby)
            Dim skip_favorite_no_security As Boolean = False
            If ASCMAIN1.tblASTMENU1.Rows.Find(New String() {row.Item("MENU_ID"), row.Item("MENU_ITEM_TYPE"), row.Item("MENU_ITEM_OBJECT")}) Is Nothing Then
                If MENU_ID = "*" Then
                    skip_favorite_no_security = True
                Else
                    ASCMAIN1.tblASTMENU1.Rows.Add(row.ItemArray)
                End If
            End If

            If Not skip_favorite_no_security Then

                KEY = row.Item("MENU_ITEM_TYPE") & Chr(1) & row.Item("MENU_ITEM_OBJECT")
                If MENU_ID = "*" Then
                    MENU_ITEM_DESC = row.Item("MENU_ITEM_DESC_FAVORITE")
                Else
                    MENU_ITEM_DESC = row.Item("MENU_ITEM_DESC")
                End If

                If level = 0 Then
                    aNode = tvw.Nodes.Add()
                Else
                    aNode = tvw.GetNodeByKey(Mid(KEY_PREFIX, 1, Len(KEY_PREFIX) - 1)).Nodes.Add
                End If

                Try
                    aNode.Key = KEY_PREFIX & KEY
                    aNode.Text = MENU_ITEM_DESC
                    aNode.Tag = row.Item("MENU_ID") & Chr(1) & KEY
                    aNode.Expanded = False
                    If row.Item("MENU_ITEM_TYPE") = "M" Then
                        aNode.Override.NodeAppearance.Image = Get_Image(IMAGE_FOLDER, "M")
                        aNode.Override.ExpandedNodeAppearance.Image = Get_Image(IMAGE_FOLDER, "M_OPEN")
                    Else
                        Dim MENU_ITEM_TYPE_image As String = row.Item("MENU_ITEM_TYPE")
                        aNode.LeftImages.Add(Get_Image(IMAGE_FOLDER, MENU_ITEM_TYPE_image))
                        'If row.Item("MENU_ITEM_TYPE") = "R" Then
                        If Not ASCMAIN1.MENU_ITEM_OBJECTs.Contains(row.Item("MENU_ITEM_OBJECT")) Then
                            ASCMAIN1.MENU_ITEM_OBJECTs.Add(row.Item("MENU_ITEM_OBJECT"))
                        End If
                        'End If
                    End If

                    If aNode.Cells.Count <> 0 Then
                        aNode.Cells("MENU_ITEM_DESC").Value = row.Item("MENU_ITEM_DESC")

                        aNode.Cells("MENU_ITEM_OBJECT").Value = row.Item("MENU_ITEM_OBJECT")
                        aNode.Cells("MENU_ITEM_SECURITY").Value = row.Item("MENU_ITEM_SECURITY")
                        aNode.Cells("MENU_ITEM_PP").Value = row.Item("MENU_ITEM_PP")
                        aNode.Cells("MENU_ITEM_PASSWORD").Value = row.Item("MENU_ITEM_PASSWORD")
                        If Val(row.Item("MENU_ITEM_HIDDEN") & "") = 0 Then
                            aNode.Cells("MENU_ITEM_HIDDEN").Value = False
                        Else
                            aNode.Cells("MENU_ITEM_HIDDEN").Value = True
                        End If
                        aNode.Cells("MENU_ITEM_FORM").Value = row.Item("MENU_ITEM_FORM")
                        'aNode.Cells("MENU_ITEM_EOM_CHECK").Value = row.Item("MENU_ITEM_EOM_CHECK")
                        If Val(row.Item("MENU_ITEM_EOM_CHECK") & "") = 0 Then
                            aNode.Cells("MENU_ITEM_EOM_CHECK").Value = False
                        Else
                            aNode.Cells("MENU_ITEM_EOM_CHECK").Value = True
                        End If
                        If Val(row.Item("MENU_ITEM_STANDALONE") & "") = 0 Then
                            aNode.Cells("MENU_ITEM_STANDALONE").Value = False
                        Else
                            aNode.Cells("MENU_ITEM_STANDALONE").Value = True
                        End If
                        aNode.Cells("MENU_ITEM_TYPE").Value = row.Item("MENU_ITEM_TYPE")
                        aNode.Cells("MENU_ITEM_DESC").Appearance.Image = Get_Image(IMAGE_FOLDER, row.Item("MENU_ITEM_TYPE"))
                        aNode.Cells("MENU_ITEM_STATUS").Value = row.Item("MENU_ITEM_STATUS")
                    End If

                    If row.Item("MENU_ITEM_TYPE") = "M" Then
                        Call Add_Menu_to_Tree(row.Item("MENU_ITEM_OBJECT"), KEY_PREFIX & KEY & Chr(0), tvw, level + 1, tblASTMENU1)
                    End If
                Catch ex As Exception
                    tvw.Nodes.Remove(aNode)
                End Try

            End If
        Next
    End Sub

    Public Shared Function Get_Image(
        ByVal IMAGE_FOLDER As String,
        ByVal IMAGE_FILE As String,
        Optional return_byte_array As Boolean = False,
        Optional ByRef IMAGE_FOLDER_USED As String = "",
        Optional ByRef IMAGE_FILE_USED As String = "",
        Optional ByRef byte_array() As Byte = Nothing) As System.Drawing.Bitmap

        Dim img As System.Drawing.Bitmap = Nothing

        IMAGE_FOLDER_USED = IMAGE_FOLDER
        IMAGE_FILE_USED = IMAGE_FILE

        Dim image_file_found As Boolean = True

        If IMAGE_FILE = "\.jpg" Then
            image_file_found = False
            Return My.Resources.ABS
        End If


        If Not IMAGE_FOLDER_USED.EndsWith("\") Then IMAGE_FOLDER_USED &= "\"
        Dim IMAGE_FILENAME As String = IMAGE_FOLDER_USED & IMAGE_FILE
        Try
            If My.Computer.FileSystem.FileExists(IMAGE_FILENAME) Then
                img = System.Drawing.Image.FromFile(IMAGE_FILENAME)
            ElseIf My.Computer.FileSystem.FileExists(IMAGE_FILENAME & ".PNG") Then
                img = System.Drawing.Image.FromFile(IMAGE_FILENAME & ".PNG")
                IMAGE_FILE_USED &= ".PNG"
            ElseIf My.Computer.FileSystem.FileExists(IMAGE_FILENAME & ".JPG") Then
                img = System.Drawing.Image.FromFile(IMAGE_FILENAME & ".JPG")
                IMAGE_FILE_USED &= ".JPG"
            Else
                image_file_found = False
                img = My.Resources.ABS
            End If

        Catch ex As Exception
            image_file_found = False
            img = My.Resources.ABS
        End Try

        Try
            img.MakeTransparent(System.Drawing.Color.White)
        Catch ex As Exception

        End Try

        If img IsNot Nothing And image_file_found And return_byte_array Then
            byte_array = GetImageData(IMAGE_FOLDER_USED & IMAGE_FILE_USED)
        End If

        Get_Image = img
    End Function

    Public Shared Function Get_Filename(
    ByVal EXT As String) As String
        Select Case EXT
            Case "TXT"
                Get_Filename = ""
            Case "XLS", "XLSX"
                Get_Filename = "EXCEL"
            Case "PDF"
                Get_Filename = "PDF"
            Case "MSG", "EML"
                Get_Filename = "MAIL"
            Case "DOC", "DOCX"
                Get_Filename = "WORD"
            Case Else
                Get_Filename = ""
        End Select

    End Function
    Public Shared Function Flattened_List(
    ByVal COLUMN_KEY As String,
    ByVal COLUMN_DATA As String,
    ByVal TABLE_NAME As String,
    Optional ByVal DELIMITER As String = ",",
    Optional ByVal WHERE_CLAUSE As String = "",
    Optional ByVal ORDER_BY As String = "") As String

        If ORDER_BY = "" Then ORDER_BY = COLUMN_DATA

        Dim SQL As String
        SQL = "Select " & COLUMN_KEY & ","
        SQL = SQL & "  ltrim(sys_connect_by_path(" & COLUMN_DATA & ",'" & DELIMITER & "'),'" & DELIMITER & "') " & COLUMN_DATA & "S"
        SQL = SQL & "      from"
        SQL = SQL & "    (select " & COLUMN_KEY & ", " & COLUMN_DATA & ","
        SQL = SQL & "           row_number() over(partition by " & COLUMN_KEY & " order by " & ORDER_BY & ") rn,"
        SQL = SQL & "          row_number() over(partition by " & COLUMN_KEY & " order by " & ORDER_BY & " desc)"
        SQL = SQL & "  rn_desc"
        SQL = SQL & " FROM " & TABLE_NAME
        If WHERE_CLAUSE <> "" Then
            SQL = SQL & ASCMAIN1.SQL_Add_WHERE(WHERE_CLAUSE)
        End If
        SQL = SQL & ")"
        SQL = SQL & "     Where rn_desc = 1"
        SQL = SQL & "     start with rn = 1"
        SQL = SQL & "    connect by prior " & COLUMN_KEY & " = " & COLUMN_KEY & ""
        SQL = SQL & "  and prior rn = rn-1"

        Flattened_List = SQL
    End Function

    Public Shared Sub Setup_Commands(ByRef UltraToolbarsManager1 As Infragistics.Win.UltraWinToolbars.UltraToolbarsManager)
        Try
            UltraToolbarsManager1.Toolbars("Commands").Settings.Appearance.BackColor = System.Drawing.Color.FromArgb(255, 245, 245, 230)
            UltraToolbarsManager1.Toolbars("Commands").Settings.Appearance.BackColor2 = System.Drawing.Color.FromArgb(255, 165, 165, 150)
            UltraToolbarsManager1.Toolbars("Commands").Settings.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
            Dim toolClose As New Infragistics.Win.UltraWinToolbars.ButtonTool("Close")
            toolClose.SharedProps.Caption = "Close"
            UltraToolbarsManager1.Tools.Add(toolClose)
            UltraToolbarsManager1.Toolbars("Commands").Tools.Add(toolClose)
            UltraToolbarsManager1.Toolbars("Commands").Settings.AllowCustomize = Infragistics.Win.DefaultableBoolean.False
        Catch ex As Exception
        End Try

        UltraToolbarsManager1.Toolbars.AddToolbar("Status")
        UltraToolbarsManager1.Toolbars("Status").DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Bottom
        UltraToolbarsManager1.Toolbars("Status").Settings.AllowCustomize = Infragistics.Win.DefaultableBoolean.False

        'Dim toolPrompt1 As New Infragistics.Win.UltraWinToolbars.LabelTool("Prompt1")
        'toolPrompt1.SharedProps.Caption = ""
        'toolPrompt1.SharedProps.Width = 300
        'UltraToolbarsManager1.Tools.Add(toolPrompt1)
        'UltraToolbarsManager1.Toolbars("Status").Tools.Add(toolPrompt1)

        'Dim toolPrompt2 As New Infragistics.Win.UltraWinToolbars.LabelTool("Prompt2")
        'toolPrompt2.SharedProps.Caption = ""
        'toolPrompt2.SharedProps.Width = 200
        'UltraToolbarsManager1.Tools.Add(toolPrompt2)
        'UltraToolbarsManager1.Toolbars("Status").Tools.Add(toolPrompt2)

        Dim folder As String
        folder = ASCMAIN1.Folders("Images") & "ABS\Toolbar\"
        For i As Integer = 0 To UltraToolbarsManager1.Tools.Count - 1
            If UltraToolbarsManager1.Tools(i).SharedProps.AppearancesSmall.Appearance.Image Is Nothing Then
                UltraToolbarsManager1.Tools(i).SharedProps.AppearancesSmall.Appearance.Image = ASCMAIN1.Get_Image(folder, UltraToolbarsManager1.Tools(i).Key)
            End If
        Next

    End Sub

    Public Shared Sub Enable_Commands(ByVal toolbar As String, ByVal keys() As String, ByRef UltraToolbarsManager1 As Infragistics.Win.UltraWinToolbars.UltraToolbarsManager)
        For i As Integer = 0 To UltraToolbarsManager1.Toolbars(toolbar).Tools.Count - 1
            UltraToolbarsManager1.Toolbars(toolbar).Tools(i).SharedProps.Enabled = False
        Next i

        For Each key As String In keys
            UltraToolbarsManager1.Toolbars(toolbar).Tools(key).SharedProps.Enabled = True
        Next
    End Sub

    Public Shared Sub Load_ExplorerBar(ByRef f As ASFBASE1, ByRef exb As UltraWinExplorerBar.UltraExplorerBar)

        'exb.Groups.Clear()

        If exb.Groups.Count > 0 Then
            ASFMAIN1.tblImages.Rows.Clear()
            Dim GROUP_INDEX As Integer = -1
            For Each grp As UltraWinExplorerBar.UltraExplorerBarGroup In exb.Groups
                GROUP_INDEX += 1
                If grp.Key = "" Then grp.Key = grp.Text
                'grp.ToolTipText = rowASTEXPB1("GROUP_TOOLTIPTEXT") & ""
                grp.Settings.AllowEdit = DefaultableBoolean.False
                grp.Settings.AllowDrag = DefaultableBoolean.False
                grp.Settings.AllowItemDrop = DefaultableBoolean.False
                grp.Settings.AllowItemUncheck = DefaultableBoolean.False
                grp.Expanded = True

                Dim ITEM_INDEX As Integer = -1
                For Each itm As UltraWinExplorerBar.UltraExplorerBarItem In grp.Items
                    ITEM_INDEX += 1

                    If itm.Key = "" Then itm.Key = itm.Text
                    'itm.ToolTipText = rowASTEXPB2("ITEM_TOOLTIPTEXT") & ""
                    itm.Settings.AppearancesLarge.Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "32\", itm.Key)
                    itm.Settings.AppearancesSmall.Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", itm.Key)
                    itm.Settings.AllowEdit = DefaultableBoolean.False
                    itm.Settings.AllowDragMove = DefaultableBoolean.False
                    itm.Settings.AllowDragCopy = DefaultableBoolean.False
                    'itm.Settings.Style = Val(rowASTEXPB2("ITEM_STYLE") & "")
                    itm.Settings.UseMnemonics = DefaultableBoolean.True
                    LoadKeysFromPopup(itm.Key, itm.Text, itm.ToolTipText)
                Next
            Next
            ASFMAIN1.grdImages.DataSource = ASFMAIN1.tblImages

            'If ASCMAIN1.DBS_SERVER = "" And ASCMAIN1.USER_ID = "wjz" Then
            '    ASCMAIN1.sql = "Delete from ASTEXPB1 where FORM_NAME = '" & f.Name & "'"
            '    ASCDATA1.ExecuteSQL()
            '    ASCMAIN1.sql = "Delete from ASTEXPB2 where FORM_NAME = '" & f.Name & "'"
            '    ASCDATA1.ExecuteSQL()
            'End If
        Else
            Dim FORM_NAME As String = f.Name
            If f.MENU_ITEM_TYPE = "F" Then
                FORM_NAME = f.MENU_ITEM_OBJECT
            Else
                'If f.MENU_ITEM_FORM <> "" Then
                '    FORM_NAME = f.MENU_ITEM_FORM
                'End If
                If f.MENU_ITEM_TYPE = "T" Then
                    FORM_NAME = "ASFCODEM"
                End If
                If f.MENU_ITEM_TYPE = "R" Then
                    FORM_NAME = "ASFSRPTM"
                End If
            End If

            ASCMAIN1.sql = "Select * from ASTEXPB1 where FORM_NAME = '" & FORM_NAME & "'"
            Dim tblASTEXPB1 As DataTable = ASCDATA1.GetDataTable("", "ASTEXPB1")

            ASCMAIN1.sql = "Select * from ASTEXPB2 where FORM_NAME = '" & FORM_NAME & "'"
            Dim tblASTEXPB2 As DataTable = ASCDATA1.GetDataTable("", "ASTEXPB2")
            'ASFMAIN1.grdImages.DataSource = tblASTEXPB2

            For Each rowASTEXPB1 As DataRow In tblASTEXPB1.Rows
                Dim grp As UltraWinExplorerBar.UltraExplorerBarGroup = exb.Groups.Add(rowASTEXPB1("GROUP_KEY"))
                grp.Text = rowASTEXPB1("GROUP_TEXT")
                grp.ToolTipText = rowASTEXPB1("GROUP_TOOLTIPTEXT") & ""
                grp.Settings.AllowEdit = IIf(ASCMAIN1.Running_in_VS, DefaultableBoolean.True, DefaultableBoolean.False)
                grp.Settings.AllowDrag = IIf(ASCMAIN1.Running_in_VS, DefaultableBoolean.True, DefaultableBoolean.False)
                grp.Settings.AllowItemDrop = IIf(ASCMAIN1.Running_in_VS, DefaultableBoolean.True, DefaultableBoolean.False)
                grp.Settings.AllowItemUncheck = IIf(ASCMAIN1.Running_in_VS, DefaultableBoolean.True, DefaultableBoolean.False)

                For Each rowASTEXPB2 As DataRow In tblASTEXPB2.Select("GROUP_INDEX = " & rowASTEXPB1("GROUP_INDEX"))
                    Dim itm As UltraWinExplorerBar.UltraExplorerBarItem = grp.Items.Add(rowASTEXPB2("ITEM_KEY"))
                    itm.Text = rowASTEXPB2("ITEM_TEXT")
                    itm.ToolTipText = rowASTEXPB2("ITEM_TOOLTIPTEXT") & ""
                    itm.Settings.AppearancesLarge.Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "32\", itm.Key)
                    itm.Settings.AppearancesSmall.Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", itm.Key)
                    itm.Settings.AllowEdit = IIf(ASCMAIN1.Running_in_VS, DefaultableBoolean.True, DefaultableBoolean.False)
                    itm.Settings.AllowDragMove = IIf(ASCMAIN1.Running_in_VS, DefaultableBoolean.True, DefaultableBoolean.False)
                    itm.Settings.AllowDragCopy = IIf(ASCMAIN1.Running_in_VS, DefaultableBoolean.True, DefaultableBoolean.False)
                    itm.Settings.Style = Val(rowASTEXPB2("ITEM_STYLE") & "")
                    itm.Settings.UseMnemonics = DefaultableBoolean.True
                Next
            Next
        End If

        exb.Tag = ""

    End Sub

    Public Shared Sub Load_Views()

        dstASTVIEWS.Clear()
        dstASTVIEWS.Tables.Clear()

        With dstASTVIEWS.Tables
            .Add(ASCDATA1.GetDataTable("*", "ASTVIEW1"))
            .Add(ASCDATA1.GetDataTable("*", "ASTVIEW2"))
            .Add(ASCDATA1.GetDataTable("*", "ASTVIEW3"))
            .Add(ASCDATA1.GetDataTable("*", "ASTVIEW4"))
            .Add(ASCDATA1.GetDataTable("*", "ASTVIEW5"))
        End With
    End Sub

#Region "Multi-Tasking Routines"

    ''' <summary>
    ''' This procedure clears all Multi-Task locks which belong to Oracle Sessions (identified by the AUDSID) which are not found in the V$SESSION view (meaning, the connection is no longer active as far as Oracle is concerned).
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub Multi_Task_Cleanup()

        Dim ENTITY_TYPE As String
        Dim ENTITY As String
        Dim SESSION_NO As String
        Dim SELECTION_NO As Integer
        Dim MT_MENU As String

        Dim sql As String = "Select ENTITY_TYPE, ENTITY, SESSION_NO, SELECTION_NO, MT_MENU"
        sql = sql & " from ASTMTSK2"
        sql = sql & " where SESSION_ID not in (Select AUDSID from V$SESSION)"
        Dim tblASTMTSK2 As DataTable = ASCDATA1.GetDataTable(sql, "ASTMTSK2")
        For Each rowASTMTSK2 As DataRow In tblASTMTSK2.Rows
            ENTITY_TYPE = rowASTMTSK2.Item("ENTITY_TYPE")
            ENTITY = rowASTMTSK2.Item("ENTITY")
            SESSION_NO = rowASTMTSK2.Item("SESSION_NO")
            SELECTION_NO = Val(rowASTMTSK2.Item("SELECTION_NO") & "")
            MT_MENU = Val(rowASTMTSK2.Item("MT_MENU") & "").ToString
            MultiTask(ENTITY_TYPE, ENTITY, "?", -1, "", SESSION_NO, SELECTION_NO, False, MT_MENU)
        Next

        sql = "Delete from ASTMTSK2 where SESSION_ID not in (Select AUDSID from V$SESSION)"
        ASCDATA1.ExecuteSQL(sql)
        sql = "Delete from ASTMTSK1 where (ENTITY_TYPE, ENTITY) not in "
        sql = sql & " (Select ENTITY_TYPE, ENTITY from ASTMTSK2)"
        ASCDATA1.ExecuteSQL(sql)
    End Sub
    ''' <summary>
    ''' This procedure attempts to Logically Lock the ENTITY_TYPE.ENTITY specified.  If the lock attempt fails, then an error message is issued to the screen, and all other logical locks and opens attempted previous to this one (within this session) are reversed.
    ''' </summary>
    ''' <param name="ENTITY_TYPE"></param>
    ''' <param name="ENTITY"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Logical_Lock(
    ByVal ENTITY_TYPE As String,
    ByVal ENTITY As String,
    Optional ByVal menu_level As Boolean = False,
    Optional ByVal show_message As Boolean = True,
    Optional ByVal reverse_all_previous_if_unsuccessful As Boolean = True,
    Optional ByVal MT_LEVEL As Integer = 0) As Boolean
        Dim EMsg As String = ""
        If Not MultiTask(ENTITY_TYPE, ENTITY, "L", 0, EMsg, , ,
                         menu_level, , reverse_all_previous_if_unsuccessful, MT_LEVEL) Then
            If show_message Then MsgBox(EMsg, vbOKOnly, "Cannot Proceed")
            Logical_Lock = False
        Else
            Logical_Lock = True
        End If
    End Function

    Public Shared Function Logical_Open(
    ByVal ENTITY_TYPE As String,
    ByVal ENTITY As String,
    Optional ByVal menu_level As Boolean = False,
    Optional ByVal show_message As Boolean = True,
    Optional ByVal reverse_all_previous_if_unsuccessful As Boolean = True,
    Optional ByVal MT_LEVEL As Integer = 0) As Boolean
        Dim EMsg As String = ""
        If Not MultiTask(ENTITY_TYPE, ENTITY, "O", 1, EMsg, , ,
                         menu_level, , reverse_all_previous_if_unsuccessful, MT_LEVEL) Then
            MsgBox(EMsg, vbOKOnly, "Cannot Proceed")
            Logical_Open = False
        Else
            Logical_Open = True
        End If
    End Function

    Public Shared Function MultiTask(
    ByVal ENTITY_TYPE As String,
    ByVal ENTITY As String,
    ByVal MT_ACTION As String,
    ByVal OPEN_COUNT As Integer,
    Optional ByRef EMsg As String = "",
    Optional ByVal SESSION_NO As String = "",
    Optional ByVal SELECTION_NO As Integer = 0,
    Optional ByVal menu_level As Boolean = False,
    Optional ByVal MT_MENU As String = "0",
    Optional ByVal reverse_all_previous_if_unsuccessful As Boolean = True,
    Optional ByVal MT_LEVEL As Integer = 0) As Boolean

        If ASCMAIN1.developerModeOptions.BypassMultiTask Then
            MultiTask = True
            Return MultiTask
        End If

        If SESSION_NO = "" Then
            SESSION_NO = ASCMAIN1.SESSION_NO
        End If
        If menu_level Then
            SELECTION_NO = UBound(ABS_FORMS) + 1
        Else
            If SELECTION_NO = 0 Then
                If ASCMAIN1.ActiveForm Is Nothing Then
                    Exit Function
                End If
                SELECTION_NO = ASCMAIN1.ActiveForm.SELECTION_NO
            End If
        End If

        MultiTask = False
        Dim sql As String = ""

        Dim sqlx As String = ""
        sqlx &= " where ENTITY_TYPE = :PARM1"
        sqlx &= "   and ENTITY = :PARM2"

        If OPEN_COUNT = -1 Then
            sql = "Select * from ASTMTSK2 " & sqlx _
             & " and SESSION_NO = '" & SESSION_NO & "'" _
             & " and SELECTION_NO = " & CStr(SELECTION_NO) _
             & "   and MT_MENU = '" & MT_MENU & "'"
            If MT_LEVEL <> 0 Then
                sql &= "   and MT_LEVEL = " & CStr(MT_LEVEL)
            End If

            Using tblASTMTSK2 As DataTable = ASCDATA1.GetDataTable(sql, "LOCKTABLE", "VV", New String() {ENTITY_TYPE, ENTITY})
                If tblASTMTSK2.Rows.Count = 0 Then
                    Exit Function
                End If
            End Using
        End If

        Dim tblASTMTSK1 As New DataTable
        sql = "Select * from ASTMTSK1" & sqlx

        If OPEN_COUNT <> -1 And MT_MENU <> "1" Then
            T = oraCon.BeginTransaction
        End If

        With ASCDATA1.GetDataAdapter(tblASTMTSK1, "ASTMTSK1", sql, True, , , 0, , , "VV",
                                     New String() {ENTITY_TYPE, ENTITY})
            Try
                Dim rowASTMTSK1 As DataRow

                If MT_ACTION = "?" Then
                    MT_ACTION = "O"
                    If tblASTMTSK1.Rows.Count <> 0 Then
                        If tblASTMTSK1.Rows(0).Item("MT_ACTION") = "L" Then
                            MT_ACTION = "U"
                        End If
                    End If
                End If

                Select Case MT_ACTION
                    Case "L", "X"
                        If tblASTMTSK1.Rows.Count = 0 Then
                            rowASTMTSK1 = tblASTMTSK1.NewRow
                            rowASTMTSK1.Item("ENTITY_TYPE") = ENTITY_TYPE
                            rowASTMTSK1.Item("ENTITY") = ENTITY
                            rowASTMTSK1.Item("MT_ACTION") = MT_ACTION
                            rowASTMTSK1.Item("LOCK_BY") = ASCMAIN1.USER_ID
                            rowASTMTSK1.Item("OPEN_COUNT") = OPEN_COUNT
                            tblASTMTSK1.Rows.Add(rowASTMTSK1)
                            .Update(tblASTMTSK1)
                            MultiTask = MultiTask_Detail(ENTITY_TYPE, ENTITY, OPEN_COUNT, SESSION_NO, SELECTION_NO, menu_level, MT_MENU, MT_LEVEL)
                        Else
                            If tblASTMTSK1.Rows.Count = 1 AndAlso tblASTMTSK1.Rows(0).Item("LOCK_BY") & "" = ASCMAIN1.USER_ID AndAlso tblASTMTSK1.Rows(0).Item("MT_ACTION") & "" = "L" Then
                                Dim tbl As DataTable = ASCDATA1.GetDataTable("Select * from ASTMTSK2 where ENTITY_TYPE = '" & ENTITY_TYPE & "' and ENTITY = '" & ENTITY & "'")
                                If tbl.Rows.Count = 1 AndAlso tbl.Rows(0).Item("SESSION_NO") = SESSION_NO AndAlso Val(tbl.Rows(0).Item("SELECTION_NO") & "") = SELECTION_NO Then
                                    MultiTask = True
                                End If
                            End If
                            If Not MultiTask Then
                                EMsg = vbCr & "Multi-Tasking Conflict w/Other Users: " _
                                 & MultiTask_Get_Users(ENTITY_TYPE, ENTITY, MT_ACTION)
                            End If
                        End If

                    Case "U"
                        If tblASTMTSK1.Rows.Count <> 0 Then
                            tblASTMTSK1.Rows(0).Delete()
                            .Update(tblASTMTSK1)
                            MultiTask = MultiTask_Detail(ENTITY_TYPE, ENTITY, OPEN_COUNT, SESSION_NO, SELECTION_NO, menu_level, MT_MENU, MT_LEVEL)
                        End If

                    Case "O"
                        If tblASTMTSK1.Rows.Count <> 0 Then
                            If OPEN_COUNT = -1 Then
                                tblASTMTSK1.Rows(0).Item("OPEN_COUNT") = Val(tblASTMTSK1.Rows(0).Item("OPEN_COUNT") & "") + OPEN_COUNT
                                If Val(tblASTMTSK1.Rows(0).Item("OPEN_COUNT") & "") = 0 Then
                                    tblASTMTSK1.Rows(0).Delete()
                                End If
                                .Update(tblASTMTSK1)
                                MultiTask = MultiTask_Detail(ENTITY_TYPE, ENTITY, OPEN_COUNT, SESSION_NO, SELECTION_NO, menu_level, MT_MENU, MT_LEVEL)
                            Else
                                If tblASTMTSK1.Rows(0).Item("MT_ACTION") = "L" Then
                                    Dim Locking_Application As String
                                    If DBS_TYPE = DBS_TYPE_types.SQLServer Then
                                        Locking_Application = ASCDATA1.GetDataValue("Select MENU_ITEM_DESC FROM ASTMENU1 WHERE (MENU_ITEM_TYPE + '-' + MENU_ITEM_OBJECT) in (Select MENU_ITEM_TYPE + '-' + MENU_ITEM_OBJECT from ASTMTSK2 WHERE ENTITY_TYPE = :PARM1 AND ENTITY = :PARM2)", "VV", New String() {ENTITY_TYPE, ENTITY})
                                    Else
                                        Locking_Application = ASCDATA1.GetDataValue("Select MENU_ITEM_DESC FROM ASTMENU1 WHERE (MENU_ITEM_TYPE, MENU_ITEM_OBJECT) in (Select MENU_ITEM_TYPE, MENU_ITEM_OBJECT from ASTMTSK2 WHERE ENTITY_TYPE = :PARM1 AND ENTITY = :PARM2)", "VV", New String() {ENTITY_TYPE, ENTITY})
                                    End If

                                    EMsg = vbCr & "Multi-Tasking Conflict w/User " & tblASTMTSK1.Rows(0).Item("LOCK_BY") & " (" & Locking_Application & ")"
                                Else
                                    'EDITING A ROW?
                                    tblASTMTSK1.Rows(0).Item("OPEN_COUNT") = Val(tblASTMTSK1.Rows(0).Item("OPEN_COUNT") & "") + OPEN_COUNT
                                    .Update(tblASTMTSK1)
                                    MultiTask = MultiTask_Detail(ENTITY_TYPE, ENTITY, OPEN_COUNT, SESSION_NO, SELECTION_NO, menu_level, MT_MENU, MT_LEVEL)
                                End If
                            End If
                        Else
                            rowASTMTSK1 = tblASTMTSK1.NewRow
                            rowASTMTSK1.Item("ENTITY_TYPE") = ENTITY_TYPE
                            rowASTMTSK1.Item("ENTITY") = ENTITY
                            rowASTMTSK1.Item("MT_ACTION") = MT_ACTION
                            rowASTMTSK1.Item("OPEN_COUNT") = OPEN_COUNT
                            If MT_MENU = "1" Then
                                rowASTMTSK1.Item("LOCK_BY") = ASCMAIN1.USER_ID
                            End If
                            tblASTMTSK1.Rows.Add(rowASTMTSK1)
                            .Update(tblASTMTSK1)

                            MultiTask = MultiTask_Detail(ENTITY_TYPE, ENTITY, OPEN_COUNT, SESSION_NO, SELECTION_NO, menu_level, MT_MENU, MT_LEVEL)
                        End If
                End Select

                If Not MultiTask And reverse_all_previous_if_unsuccessful Then
                    If MT_ACTION = "L" Or OPEN_COUNT > 0 Then
                        Call MultiTask_Release(SESSION_NO, SELECTION_NO, MT_LEVEL)
                    End If
                End If

            Catch ex As Exception
                'MsgBox("Error " & ex.Message)
                'Stop

            End Try

        End With
        If OPEN_COUNT <> -1 And MT_MENU <> "1" Then
            T.Commit()
        End If

    End Function

    Public Shared Function MultiTask_Detail(
    ByVal ENTITY_TYPE As String,
    ByVal ENTITY As String,
    ByVal OPEN_COUNT As Integer,
    Optional ByVal SESSION_NO As String = "",
    Optional ByVal SELECTION_NO As Integer = 0,
    Optional ByVal menu_check As Boolean = False,
    Optional ByVal MT_MENU As String = "0",
    Optional ByVal MT_LEVEL As Integer = 0) As Boolean

        If ASCMAIN1.developerModeOptions.BypassMultiTask Then
            MultiTask_Detail = True
            Return MultiTask_Detail
        End If

        If SESSION_NO = "" Then
            SESSION_NO = ASCMAIN1.SESSION_NO
        End If
        If SELECTION_NO = 0 Then
            SELECTION_NO = ASCMAIN1.ActiveForm.SELECTION_NO
        End If

        MultiTask_Detail = False
        Dim sql As String = ""

        Try
            If OPEN_COUNT = -1 Then

                Dim sqlx As String = ""
                sqlx &= " where ENTITY_TYPE = :PARM1"
                sqlx &= "   and ENTITY = :PARM2"

                sql = "Delete from ASTMTSK2 " & sqlx
                sql = sql & "   and SESSION_NO = '" & SESSION_NO & "'"
                sql = sql & "   and SELECTION_NO = " & SELECTION_NO
                sql = sql & "   and MT_MENU = " & MT_MENU
                If MT_LEVEL <> 0 Then
                    sql = sql & "   and MT_LEVEL = " & CStr(MT_LEVEL)
                End If
                ASCDATA1.ExecuteSQL(sql, "VV", New String() {ENTITY_TYPE, ENTITY})

            Else
                Dim tblASTMTSK2 As New DataTable

                ASCMAIN1.sql = "Select * from ASTMTSK2 where ENTITY_TYPE = :PARM1 and ENTITY = :PARM2 and SESSION_NO = :PARM3 and SELECTION_NO = :PARM4"
                With ASCDATA1.GetDataAdapter(tblASTMTSK2, "ASTMTSK2", "**", True, -1, True, 0, , , "VVVN", New Object() {ENTITY_TYPE, ENTITY, SESSION_NO, SELECTION_NO})

                    If tblASTMTSK2.Rows.Count = 0 Then
                        Dim rowASTMTSK2 As DataRow = tblASTMTSK2.NewRow
                        rowASTMTSK2.Item("ENTITY_TYPE") = ENTITY_TYPE
                        rowASTMTSK2.Item("ENTITY") = ENTITY
                        rowASTMTSK2.Item("SESSION_NO") = SESSION_NO
                        rowASTMTSK2.Item("SELECTION_NO") = SELECTION_NO
                        rowASTMTSK2.Item("INIT_OPER") = ASCMAIN1.USER_ID
                        rowASTMTSK2.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                        If menu_check Then
                            rowASTMTSK2.Item("MENU_ITEM_TYPE") = ASFMAIN1.MENU_ITEM_TYPE
                            rowASTMTSK2.Item("MENU_ITEM_OBJECT") = ASFMAIN1.MENU_ITEM_OBJECT
                        Else
                            rowASTMTSK2.Item("MENU_ITEM_TYPE") = ASCMAIN1.ActiveForm.MENU_ITEM_TYPE
                            rowASTMTSK2.Item("MENU_ITEM_OBJECT") = ASCMAIN1.ActiveForm.MENU_ITEM_OBJECT
                        End If
                        rowASTMTSK2.Item("MT_MENU") = MT_MENU
                        rowASTMTSK2.Item("MT_LEVEL") = MT_LEVEL
                        rowASTMTSK2.Item("SESSION_ID") = ASCMAIN1.DBS_SESSION_ID

                        tblASTMTSK2.Rows.Add(rowASTMTSK2)
                        .Update(tblASTMTSK2)
                    End If
                End With
            End If
            MultiTask_Detail = True

        Catch ex As Exception
            Stop
        End Try

    End Function

    Public Shared Function MultiTask_Get_Users(
    ByVal ENTITY_TYPE As String,
    ByVal ENTITY As String,
    ByVal MT_ACTION As String) As String

        Dim Users As String = ""

        Dim sql As String = "Select " _
         & " ASTMTSK2.INIT_OPER, ASTMTSK2.INIT_DATE, ASTUSER1.USER_NAME, " _
         & " ASTMTSK2.MENU_ITEM_TYPE, ASTMTSK2.MENU_ITEM_OBJECT " _
         & " from ASTMTSK2, ASTUSER1" _
         & " where ASTMTSK2.ENTITY_TYPE = :PARM1" _
         & "   and ASTMTSK2.ENTITY = :PARM2" _
         & "   and ASTUSER1.USER_ID (+) = ASTMTSK2.INIT_OPER"

        Using tbl As DataTable = ASCDATA1.GetDataTable(sql, "MTUSER", "VV", New String() {ENTITY_TYPE, ENTITY})
            For Each row As DataRow In tbl.Rows
                Users &= vbCrLf & row.Item("INIT_OPER") & " (" & row.Item("USER_NAME") & ")"
                Users = Users & vbCrLf & vbCrLf & "System Codes: " & vbCr & ENTITY_TYPE & " : " & ENTITY

                'If MT_ACTION <> "L" Then
                sql = "Select * from ASTMENU1 " _
                 & " where (MENU_ITEM_TYPE = :PARM1" _
                 & "   and  MENU_ITEM_OBJECT = :PARM2)"
                Dim rowASTMENU1 As DataRow = ASCDATA1.GetDataRow(sql, "VV", New String() {row.Item("MENU_ITEM_TYPE") & "", row.Item("MENU_ITEM_OBJECT") & ""})
                If rowASTMENU1 IsNot Nothing Then

                    Users = Users & " (" & rowASTMENU1.Item("MENU_ITEM_DESC") & ")"
                Else
                    If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                        sql = "Select * from WHTGUNA1 where PROCEDURE_NAME = '" & row.Item("MENU_ITEM_OBJECT") & "'"
                        Dim rowWHTGUNA1 As DataRow = ASCDATA1.GetDataRow(sql)
                        If rowWHTGUNA1 IsNot Nothing Then
                            Users = Users & " (" & rowWHTGUNA1.Item("APP_DESC") & ")"
                        End If
                    End If
                End If

                'End If
            Next
        End Using
        MultiTask_Get_Users = Users
    End Function

    ''' <summary>
    ''' This procedure releases all Multi-Task Locks and Opens that are associated with the specified SESSION_NO and SELECTION_NO.  If SESSION_NO is ommitted, then the SESSION_NO of the Active Form is used.  If SELECTION_NO is omitted, then the SELECTION_NO of the Active form is used.
    ''' </summary>
    ''' <param name="SESSION_NO"></param>
    ''' <param name="SELECTION_NO"></param>
    ''' <remarks></remarks>
    Public Shared Sub MultiTask_Release(
    Optional ByVal SESSION_NO As String = "",
    Optional ByVal SELECTION_NO As Integer = 0,
    Optional ByVal MT_LEVEL As Integer = 0)

        If SESSION_NO = "" Then
            SESSION_NO = ASCMAIN1.SESSION_NO
        End If
        If SELECTION_NO = 0 Then
            If ASCMAIN1.ActiveForm Is Nothing Then
                Exit Sub
            End If
            SELECTION_NO = ASCMAIN1.ActiveForm.SELECTION_NO
        End If

        Dim sql As String = "Select * from ASTMTSK2 " _
         & " where SESSION_NO = :PARM1" _
         & " and SELECTION_NO = :PARM2" _
         & " and NVL(MT_MENU,'0') <> '1'"
        If MT_LEVEL <> 0 Then
            sql &= " and MT_LEVEL = " & CStr(MT_LEVEL)
        End If
        Dim tblASTMTSK2 As DataTable = ASCDATA1.GetDataTable(sql, "MTREL", "VV", New String() {SESSION_NO, CStr(SELECTION_NO)})

        For Each rowASTMTSK2 As DataRow In tblASTMTSK2.Rows
            MultiTask(rowASTMTSK2.Item("ENTITY_TYPE"), rowASTMTSK2.Item("ENTITY"), "?", -1, "", SESSION_NO, SELECTION_NO, , , , MT_LEVEL)
        Next
    End Sub

    Public Shared Function Multi_Task_Menu_Item(
    ByVal MENU_ITEM_TYPE As String,
    ByVal MENU_ITEM_OBJECT As String,
    ByVal OPEN_COUNT As Integer,
    Optional ByVal menu_check As Boolean = False,
    Optional ByVal MENU_ITEM_STANDALONE As String = "") As Boolean

        Dim ENTITY_TYPE As String
        Dim ENTITY As String
        Dim MT_ACTION As String
        Dim EMsg As String = ""

        Dim sql As String = ""
        Multi_Task_Menu_Item = False

        T = oraCon.BeginTransaction()

        If Not MultiTask(MENU_ITEM_TYPE, MENU_ITEM_OBJECT, "O", OPEN_COUNT, EMsg, , , menu_check, "1") Then
            T.Rollback()
            If OPEN_COUNT <> -1 Then
                MsgBox(EMsg, 0, "Cannot Make Requested Menu Selection At This Time")
            End If
            Exit Function
        End If

        sql = "Select * from ASTMTKC1 " _
         & " where MENU_ITEM_TYPE = :PARM1" _
         & "   and MENU_ITEM_OBJECT = :PARM2"

        Dim tblASTMTKC1 As DataTable = ASCDATA1.GetDataTable(sql, "ASTMTKC1", "VV", New String() {MENU_ITEM_TYPE, MENU_ITEM_OBJECT})
        Dim rowASTMTKC1 As DataRow = tblASTMTKC1.NewRow
        rowASTMTKC1.Item("MENU_ITEM_TYPE") = MENU_ITEM_TYPE
        rowASTMTKC1.Item("MENU_ITEM_OBJECT") = MENU_ITEM_OBJECT
        rowASTMTKC1.Item("ENTITY_TYPE") = "S"
        rowASTMTKC1.Item("ENTITY") = ASCMAIN1.SESSION_NO
        If MENU_ITEM_STANDALONE = "1" Then
            rowASTMTKC1.Item("MT_ACTION") = "L"
        Else
            rowASTMTKC1.Item("MT_ACTION") = "O"
        End If
        tblASTMTKC1.Rows.Add(rowASTMTKC1)

        For Each rowASTMTKC1 In tblASTMTKC1.Rows

            ENTITY_TYPE = rowASTMTKC1.Item("ENTITY_TYPE") & ""
            ENTITY = rowASTMTKC1.Item("ENTITY") & ""
            MT_ACTION = rowASTMTKC1.Item("MT_ACTION") & ""
            If OPEN_COUNT = 1 Then
                If ENTITY_TYPE = MENU_ITEM_TYPE And ENTITY = MENU_ITEM_OBJECT Then
                    MultiTask(MENU_ITEM_TYPE, MENU_ITEM_OBJECT, "O", -1, EMsg, , , menu_check, "1")
                End If
                If Not MultiTask(ENTITY_TYPE, ENTITY, MT_ACTION, OPEN_COUNT, EMsg, , , menu_check, "1") Then
                    MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Make Requested Menu Selection At This Time")
                    T.Rollback()
                    Exit Function
                End If
            Else
                If MT_ACTION = "L" Then
                    MT_ACTION = "U"
                End If
                Multi_Task_Menu_Item = MultiTask(ENTITY_TYPE, ENTITY, MT_ACTION, OPEN_COUNT, EMsg, , , menu_check, "1")
            End If
        Next
        T.Commit()
        T = Nothing

        If OPEN_COUNT = 1 Then
            Multi_Task_Menu_Item = True
        End If
    End Function

    Public Shared Function Reset(ByVal F As ASFBASE1) As Boolean
        If F.ScreenMode Then
            'Dim frmASFMSGBF As New ASFMSGBF
            'Dim A As String = frmASFMSGBF.Get_txt_from_User("You may not Close this Window at this time", "Form has an Active Document or Transaction Pending", True)
            'frmASFMSGBF.Dispose()

            Dim A As String = ASCMAIN1.Get_txt_from_User("You may not Close this Window at this time", "Form has an Active Document or Transaction Pending", True)
            If A = "0ff1c3" Then
                F.ScreenMode = False
            End If
        End If

        Reset = F.ScreenMode

        If F.ScreenMode Then
        Else

            ' inside this next call, it is using abs_forms + 1 even when use count is -1
            ASCMAIN1.Multi_Task_Menu_Item(F.MENU_ITEM_TYPE, F.MENU_ITEM_OBJECT, -1)

            ASFMAIN1.UltraStatusBar1.Panels("MSG1").Text = ""
            ASFMAIN1.UltraStatusBar1.Panels("MSG2").Text = ""
            ASFMAIN1.UltraStatusBar1.Panels("MODE").Text = ""
            ASFMAIN1.UltraStatusBar1.Panels("MENU_ITEM_OBJECT").Text = ""
            ASFMAIN1.UltraStatusBar1.Panels("SELECTION_NO").Text = ""

        End If
    End Function
#End Region

    Public Shared Sub Activate_Form(ByRef F As ASFBASE1)

        If ASCMAIN1.MainForm IsNot Nothing Then
            DirectCast(ASCMAIN1.MainForm.Controls("UltraStatusBar1"), UltraWinStatusBar.UltraStatusBar).Panels("MENU_ITEM_OBJECT").Text = F.Name
            DirectCast(ASCMAIN1.MainForm.Controls("UltraStatusBar1"), UltraWinStatusBar.UltraStatusBar).Panels("SELECTION_NO").Text = F.SELECTION_NO
        End If
        ASCMAIN1.ActiveForm = F
    End Sub

    Public Shared Function Register_Form(ByVal F As Form) As Integer
        'MsgBox("2xa")
        If F.Name = "ASFMAIN1" Then
            Register_Form = 0
        Else
            Register_Form = UBound(ABS_FORMS) + 1
        End If
        'MsgBox("2xb")
        ReDim Preserve ABS_FORMS(Register_Form)
        'MsgBox("2xc")
        ABS_FORMS(Register_Form) = F
        'MsgBox("2xd")
        ' ASSOCIATE HELP CONTEXT WITH THIS FORM
    End Function

    Public Shared Sub Navigate_like_Excel(ByRef grd As UltraWinGrid.UltraGrid, ByRef e As System.Windows.Forms.KeyEventArgs)

        With grd
            Select Case e.KeyValue

                Case Keys.Up

                    .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
                    .PerformAction(UltraWinGrid.UltraGridAction.AboveCell, False, False)
                    e.Handled = True
                    .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode, False, False)

                Case Keys.Down

                    .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
                    .PerformAction(UltraWinGrid.UltraGridAction.BelowCell, False, False)
                    e.Handled = True
                    .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode, False, False)

                Case Keys.Right

                    .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
                    .PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab, False, False)
                    e.Handled = True
                    .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode, False, False)

                Case Keys.Left

                    .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
                    .PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab, False, False)
                    e.Handled = True
                    .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode, False, False)

            End Select
        End With

    End Sub

    Public Shared Function Format_Field(ByVal txt As String, ByVal COLUMN_NAME As String, Optional ByVal tbl As DataTable = Nothing, Optional ByVal treat_as_code As Boolean = False)

        Dim row As DataRow = ASCMAIN1.tblASTFFMT1.Rows.Find(COLUMN_NAME)

        If row Is Nothing Then
            If treat_as_code Then
                row = ASCMAIN1.tblASTFFMT1.NewRow
            Else
                If Not (tbl Is Nothing) Then
                    For Each DC As DataColumn In tbl.PrimaryKey
                        If DC.ColumnName = COLUMN_NAME Then
                            row = ASCMAIN1.tblASTFFMT1.NewRow
                        End If
                    Next
                End If
            End If
        End If


        If Not (row Is Nothing) Then
            Dim FIELD_LENGTH As Integer = Val(row.Item("FIELD_LENGTH") & "")

            If row.Item("FIXED_LENGTH") & "" = "1" Then
                Dim JUSTIFY As String = row.Item("JUSTIFY") & ""
                Dim FILL_CHAR As String = row.Item("FILL_CHAR") & ""
                If FILL_CHAR <> "" Then
                    If JUSTIFY = "L" Then
                        txt = txt.PadRight(FIELD_LENGTH, FILL_CHAR)
                    Else
                        txt = txt.PadLeft(FIELD_LENGTH, FILL_CHAR)
                    End If
                End If
            End If

            If row.Item("ALLOW_LOWER_CASE") & "" <> "1" Then
                txt = UCase$(txt)
            End If

            If row.Item("ALPHA_NUMERIC") & "" <> "B" Then
                If row.Item("ALPHA_NUMERIC") & "" = "A" Then
                    For i As Integer = txt.Length To 1 Step -1
                        If Not "ABCDEFGHIJKLMNOPQRSTUVWXYZ".Contains(Mid$(txt, i, 1)) Then
                            txt = Mid$(txt, 1, i - 1) & Mid$(txt, i + 1)
                        End If
                    Next
                End If
                If row.Item("ALPHA_NUMERIC") & "" = "N" Then
                    For i As Integer = txt.Length To 1 Step -1
                        If Not "0123456789".Contains(Mid$(txt, i, 1)) Then
                            txt = Mid$(txt, 1, i - 1) & Mid$(txt, i + 1)
                        End If
                    Next
                End If
                If row.Item("ALPHA_NUMERIC") & "" = "C" Then
                    Dim IsNumeric As Boolean = True
                    For i As Integer = txt.Length To 1 Step -1
                        If Not "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ".Contains(Mid$(txt, i, 1)) Then
                            txt = Mid$(txt, 1, i - 1) & Mid$(txt, i + 1)
                        Else
                            If Not "0123456789".Contains(Mid$(txt, i, 1)) Then
                                IsNumeric = False
                            End If
                        End If
                    Next
                    If IsNumeric Then
                        If row.Item("FIXED_LENGTH") & "" = "2" Then
                            Dim JUSTIFY As String = row.Item("JUSTIFY") & ""
                            Dim FILL_CHAR As String = row.Item("FILL_CHAR") & ""
                            If JUSTIFY = "L" Then
                                txt = txt.PadRight(FIELD_LENGTH, FILL_CHAR)
                            Else
                                txt = txt.PadLeft(FIELD_LENGTH, FILL_CHAR)
                            End If
                        End If
                    End If
                End If

                If row.Item("FIXED_LENGTH") & "" = "1" Then
                    If txt.Length <> FIELD_LENGTH Then
                        txt = ""
                    End If
                End If

            End If

        End If

        Return txt

    End Function

    Public Shared Function Get_Legend(
    ByVal YYYYPP As String,
    Optional ByVal AppendPeriodStatus As Boolean = True,
    Optional ByVal abbreviated As Boolean = False)

        Dim YM As String
        YM = Get_YYYYMM(YYYYPP, 0)

        Dim EOM_indicator As String = ""
        If YYYYPP = CYP Then
            If EOM = "1" Then
                EOM_indicator = "I"
            Else
                EOM_indicator = " "
            End If
        Else
            If YYYYPP < CYP Then
                EOM_indicator = "F"
            End If
        End If

        Dim LEGEND As String = Mid$(YYYYPP, 1, 4) & "-" & Mid$(YYYYPP, 5, 2) & " (" & Format$(DateValue(Mid$(YM, 5, 2) & "/01/" & Mid$(YM, 3, 2)), "MMM") & "'" & Mid$(YM, 3, 2) & ")" & IIf(AppendPeriodStatus, " " & EOM_indicator, "")
        If abbreviated Then
            LEGEND = Mid(LEGEND, 10, 6)
        End If
        Return LEGEND
    End Function

    Public Shared Function Get_Legend_Wk(
    ByVal YYYYWW As String,
    Optional ByVal abbreviated As Boolean = False)
        ASCMAIN1.sql = "Select LEGEND from GLTPARM3 where YYYYWW = '" & YYYYWW & "'"
        Dim LEGEND As String = ASCDATA1.GetDataValue()
        If abbreviated Then
            LEGEND = Mid(LEGEND, 10, 7)
        End If
        Return LEGEND
    End Function

    Public Shared Sub Load_MRUs()
        MRUs.Clear()
        For Each row As DataRow In ASCDATA1.GetDataTable("*", "ASTMRUL1").Rows
            Dim MRU_List As New List(Of String)
            MRUs.Add(row.Item("COLUMN_NAME"), MRU_List)
        Next
    End Sub

    Public Shared Function Period_Calc(
    ByVal base_YP As String,
    ByVal number_of_periods As Integer) As String

        Dim p As Integer = Val(Mid$(base_YP, 1, 4)) * 12 + Val(Mid$(base_YP, 5, 2))
        p = p + number_of_periods

        Dim m As Integer
        Dim Y As Integer
        m = 1 + ((p - 1) Mod 12)
        Y = (p - m) / 12

        Return Format$(Y, "0000") & Format$(m, "00")

    End Function

    Public Shared Function Period_Diff(
    ByVal base_YP As String,
    ByVal other_YP As String) As Integer
        Return 12 * (Val(Mid$(other_YP, 1, 4)) - Val(Mid$(base_YP, 1, 4))) + (Val(Mid$(other_YP, 5, 2)) - Val(Mid$(base_YP, 5, 2)))
    End Function

    Public Shared Function Week_Calc(
    ByVal base_YW As String,
    ByVal number_of_periods As Integer) As String

        Dim Sql As String

        If number_of_periods >= 0 Then
            Sql = "Select ROWNUM -1 RELYW, YYYYWW from (Select YYYYWW from GLTPARM3 where YYYYWW >= '" & base_YW & "' order by YYYYWW)"
        Else
            Sql = "Select ROWNUM -1 RELYW, YYYYWW from (Select YYYYWW from GLTPARM3 where YYYYWW <= '" & base_YW & "' order by YYYYWW Desc)"
        End If

        Sql = "Select * from (" & Sql & ") where RELYW = " & CStr(Abs(number_of_periods))
        Dim row As DataRow = ASCDATA1.GetDataRow(Sql)
        If row Is Nothing Then
            Return "000000"
        Else
            Return row.Item("YYYYWW")
        End If
    End Function

    Public Shared Function Week_Diff(
    ByVal base_YW As String,
    ByVal other_YW As String) As Integer

        ASCMAIN1.sql = "Select Count (*) from GLTPARM3" _
        & " where YYYYWW between '" & IIf(base_YW <= other_YW, base_YW, other_YW) _
        & "' and '" & IIf(base_YW <= other_YW, other_YW, base_YW) & "'"
        Dim number_of_weeks As Integer = 0
        If base_YW <> other_YW Then
            number_of_weeks = Val(ASCDATA1.GetDataValue & "") - 1
            If base_YW > other_YW Then
                number_of_weeks = -1 * number_of_weeks
            End If
        End If
        Return number_of_weeks

    End Function

    ''' <summary>
    ''' Returns a SQL statement which will provide a result set 
    ''' showing a primary key column (COLUMN_NAME_key) and then a list of values
    ''' from a related column (COLUMN_NAME_list) using a TABLE_NAME and WHERE_CLAUSE.
    ''' Optionally, the COLUMN_NAME_list could be an expression (COLUMN_EXPRESSION_list) where the COLUMN_NAME_list would serve as an alias to expression.
    ''' The list of values is separated by commas. 
    ''' The column containing the list is aliased using the COLUMN_NAME_list appended with an S.
    ''' </summary>
    ''' <param name="TABLE_NAME"></param>
    ''' <param name="COLUMN_NAME_key"></param>
    ''' <param name="COLUMN_NAME_list"></param>
    ''' <param name="COLUMN_EXPRESSION_list"></param>
    ''' <param name="WHERE_CLAUSE"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SQL_CodeList(
    ByVal TABLE_NAME As String,
    ByVal COLUMN_NAME_key As String,
    ByVal COLUMN_NAME_list As String,
    Optional ByVal COLUMN_EXPRESSION_list As String = "",
    Optional ByVal WHERE_CLAUSE As String = "") As String

        If COLUMN_EXPRESSION_list = "" Then
            COLUMN_EXPRESSION_list = COLUMN_NAME_list
        End If

        Dim sql As String
        sql = "Select " & COLUMN_NAME_key & "," & vbCr
        sql = sql & "   ltrim(sys_connect_by_path(" & COLUMN_NAME_list & ",', '),', ') " & COLUMN_NAME_list & "S" & vbCr
        sql = sql & "      from" & vbCr
        sql = sql & "    (select " & COLUMN_NAME_key & "," & COLUMN_EXPRESSION_list & " " & COLUMN_NAME_list & "," & vbCr
        sql = sql & "           row_number() over(partition by " & COLUMN_NAME_key & " order by " & COLUMN_EXPRESSION_list & ") rn," & vbCr
        sql = sql & "          row_number() over(partition by " & COLUMN_NAME_key & " order by " & COLUMN_EXPRESSION_list & " desc)" & vbCr
        sql = sql & "  rn_desc" & vbCr
        sql = sql & " FROM " & TABLE_NAME & " " & WHERE_CLAUSE & ")" & vbCr
        sql = sql & "     Where rn_desc = 1" & vbCr
        sql = sql & "     start with rn = 1" & vbCr
        sql = sql & "    connect by prior " & COLUMN_NAME_key & " = " & COLUMN_NAME_key & vbCr
        sql = sql & "  and prior rn = rn-1" & vbCr

        Return sql
    End Function

    Public Shared Sub grdInitializeLayout(
    ByRef cmb As UltraWinGrid.UltraCombo,
    Optional ByVal F As ASFBASE0 = Nothing)
        grdInitializeLayout(cmb.DisplayLayout, F)
    End Sub

    Public Shared Sub grdInitializeLayout(
    ByRef grd As UltraWinGrid.UltraGrid,
    Optional ByVal F As ASFBASE0 = Nothing)
        grdInitializeLayout(grd.DisplayLayout, F)
        ' as per Infragistics Tech Support
        grd.RowUpdateCancelAction = UltraWinGrid.RowUpdateCancelAction.RetainDataAndActivation
    End Sub

    Public Shared Sub grdInitializeLayout(
    ByRef DL As UltraWinGrid.UltraGridLayout,
    Optional ByVal F As ASFBASE0 = Nothing)
        With DL
            .Override.AllowMultiCellOperations = UltraWinGrid.AllowMultiCellOperation.Copy ' Infragistics.Win.UltraWinGrid.AllowMultiCellOperation.All - this caused the grid column headers to get copied with ctrl-c
            '.DisplayLayout.Override.ExpansionIndicator = UltraWinGrid.ShowExpansionIndicator.CheckOnDisplay
            If .Override.CellClickAction <> UltraWinGrid.CellClickAction.RowSelect Then
                .Override.CellClickAction = UltraWinGrid.CellClickAction.CellSelect
            End If


            If .Override.RowSelectorHeaderStyle = UltraWinGrid.RowSelectorHeaderStyle.SeparateElement Then
                .Override.RowSelectorHeaderAppearance.ImageBackground = My.Resources.EXCEL12 ' My.Resources.EXCEL ' CType(Resources.GetObject("EXCEL"), System.Drawing.Image)
                .Override.RowSelectorHeaderAppearance.ImageBackgroundStyle = ImageBackgroundStyle.Centered
            End If

            .Override.GroupBySummaryDisplayStyle = UltraWinGrid.GroupBySummaryDisplayStyle.SummaryCells

            If .Bands.Count < 2 Then
                .ViewStyle = UltraWinGrid.ViewStyle.SingleBand
            End If

            .UseFixedHeaders = True
            .Override.FixedHeaderIndicator = UltraWinGrid.FixedHeaderIndicator.None

            If .Override.RowSizing = UltraWinGrid.RowSizing.Default Then
                .Override.RowSizing = UltraWinGrid.RowSizing.Fixed
            End If

            .NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Hide
            .NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Hide

            Dim gMRUs As New List(Of String)

            For Each b As UltraWinGrid.UltraGridBand In .Bands
                b.SortedColumns.Clear()
                ' b.UseRowLayout = False

                For Each c As UltraWinGrid.UltraGridColumn In b.Columns
                    c.HiddenWhenGroupBy = DefaultableBoolean.False
                    Try
                        If Not c.Hidden AndAlso c.IsFirstVisibleColumnOnLevel Then
                            'If grd.Name = "grdASTDSQLV" Then Stop
                            b.SortedColumns.Add(c, False)
                        End If
                    Catch ex As Exception

                    End Try
                    '    If C.MaxLength = -1 Then
                    '        C.MaxLength = DirectCast(.DataSource, DataTable).Columns(C.Key).MaxLength
                    '    End If
                    If c.Style = UltraWinGrid.ColumnStyle.CheckBox Then
                        c.Editor.DataFilter = New CheckEditorDataFilter
                        ' C.DefaultCellValue = "0"
                    End If

                    If c.Style = UltraWinGrid.ColumnStyle.EditButton Then
                        If c.CellButtonAppearance.Image Is Nothing Then ' grd.Name <> "grdSetup" Then
                            c.CellButtonAppearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "ARROW_UP_BLUE")
                        End If
                    End If

                    If b.CardView Then
                        c.Header.Appearance.TextHAlign = HAlign.Left
                    End If


                    If c.CellActivation = UltraWinGrid.Activation.Disabled Or c.CellActivation = UltraWinGrid.Activation.NoEdit Then
                        c.TabStop = False
                    End If

                    'If .DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.RowSelect Then
                    '    If c.CellClickAction = UltraWinGrid.CellClickAction.Default Then
                    '        c.CellClickAction = UltraWinGrid.CellClickAction.CellSelect
                    '    End If
                    'End If

                    Dim COLUMN_NAME As String = c.Key
                    If ASCMAIN1.MRUs.ContainsKey(COLUMN_NAME) Then
                        gMRUs.Add(COLUMN_NAME)
                        '                        ASFMAIN1.UltraToolbarsManager1.SetContextMenuUltra(c., "txtMenu")
                    End If
                Next
            Next

            ' NOT SURE ABOUT THESE NEXT 4
            'If F IsNot Nothing Then
            '    If Not F.grdMRUs.ContainsKey(.Name) Then
            '        F.grdMRUs.Add(.Name, gMRUs)
            '    End If
            'End If


            For Each b As UltraWinGrid.UltraGridBand In .Bands
                For Each c As UltraWinGrid.UltraGridColumn In b.Columns
                    If c.Style = UltraWinGrid.ColumnStyle.Button Then
                        If c.Header.Appearance.TextHAlign = HAlign.Default Then
                            c.Header.Appearance.TextHAlign = HAlign.Center
                        End If
                        If c.CellAppearance.TextHAlign = HAlign.Default Then
                            c.CellAppearance.TextHAlign = HAlign.Center
                        End If
                    End If
                    If c.DataType.ToString = "System.DateTime" _
                    Or c.Style = UltraWinGrid.ColumnStyle.CheckBox Then
                        If c.Header.Appearance.TextHAlign = HAlign.Default Then
                            If b.CardView Then
                                c.Header.Appearance.TextHAlign = HAlign.Left
                            Else
                                c.Header.Appearance.TextHAlign = HAlign.Center
                            End If

                            ' the following change made a checkbox work properly - the cb is in a row with other fields that are changeable, and it was taking several clicks in rapid succession to get the cb to check
                            ' Note that we need to toggle the setting before leaving it with the proper setting
                            If c.Style = UltraWinGrid.ColumnStyle.CheckBox Then
                                ' WITHOUT THE PRECEDING IF, DATETIME CELLS WOULD NOT RESPOND PROPERLY TO A DBL CLICK IN ASFCODE1 TO SELECT THE ROW
                                c.CellClickAction = UltraWinGrid.CellClickAction.CellSelect
                                c.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
                            End If


                        End If
                        If c.CellAppearance.TextHAlign = HAlign.Default Then
                            If b.CardView Then
                                c.CellAppearance.TextHAlign = HAlign.Left
                            Else
                                c.CellAppearance.TextHAlign = HAlign.Center
                            End If

                            c.CellAppearance.TextHAlign = HAlign.Center
                        End If
                        'Stop
                    End If
                    If c.DataType.ToString = "System.Int16" _
                    Or c.DataType.ToString = "System.Int32" _
                    Or c.DataType.ToString = "System.Int64" Then
                        If c.Format = "" Then
                            c.Format = "###,##0"
                        End If
                        If c.Style = UltraWinGrid.ColumnStyle.Default Then
                            c.Style = UltraWinGrid.ColumnStyle.Integer
                        End If
                        If c.Header.Appearance.TextHAlign = HAlign.Default Then
                            If b.CardView Then
                                c.Header.Appearance.TextHAlign = HAlign.Left
                            Else
                                c.Header.Appearance.TextHAlign = HAlign.Right
                            End If
                        End If
                        If c.CellAppearance.TextHAlign = HAlign.Default Then
                            c.CellAppearance.TextHAlign = HAlign.Right
                        End If
                    End If
                    If c.DataType.ToString = "System.Double" _
                    Or c.DataType.ToString = "System.Single" _
                    Or c.DataType.ToString = "System.Decimal" Then
                        If c.Format = "" Then
                            c.Format = "###,##0.00"
                        End If
                        '                    c.MaskInput = ""
                        If c.Style = UltraWinGrid.ColumnStyle.Default Then
                            c.Style = UltraWinGrid.ColumnStyle.Double
                        End If
                        If c.Header.Appearance.TextHAlign = HAlign.Default Then
                            If b.CardView Then
                                c.Header.Appearance.TextHAlign = HAlign.Left
                            Else
                                c.Header.Appearance.TextHAlign = HAlign.Right
                            End If
                        End If
                        If c.CellAppearance.TextHAlign = HAlign.Default Then
                            c.CellAppearance.TextHAlign = HAlign.Right
                        End If
                    End If

                    If c.DataType.ToString = "System.String" Then
                        If c.Format = "" Then
                            If c.Key Like "*_PHONE_NO" _
                            Or c.Key Like "*_TEL_NO" _
                            Or c.Key Like "*_FAX_NO" _
                            Or c.Key Like "*_PHONE" _
                            Or c.Key Like "*_TEL" _
                            Or c.Key Like "*_FAX" Then
                                'c.Format = "(###) ###-####"
                                c.MaskInput = "(###) ###-####"
                                c.CellDisplayStyle = UltraWinGrid.CellDisplayStyle.FormattedText
                                'c.CellDisplayStyle = UltraWinGrid.CellDisplayStyle.FullEditorDisplay

                                c.MaskDisplayMode = UltraWinMaskedEdit.MaskMode.IncludeBoth
                                c.MaskDataMode = UltraWinMaskedEdit.MaskMode.Raw
                            End If

                        End If
                    End If


                    c.PromptChar = String.Empty
                    'Stop
                Next

                'b.UseRowLayout = False

            Next

            ' as per Infragistics Tech Support
            .Override.RowLayoutCellNavigationVertical = UltraWinGrid.RowLayoutCellNavigation.Adjacent

            '.DisplayLayout.Override.SummaryDisplayArea = UltraWinGrid.SummaryDisplayAreas.BottomFixed
            .Override.SummaryDisplayArea = UltraWinGrid.SummaryDisplayAreas.BottomFixed _
                                                        Or UltraWinGrid.SummaryDisplayAreas.InGroupByRows _
                                                        Or UltraWinGrid.SummaryDisplayAreas.GroupByRowsFooter

            .Override.GroupBySummaryValueAppearance.TextHAlign = HAlign.Right
            '.DisplayLayout.Override.RowSelectorNumberStyle = UltraWinGrid.RowSelectorNumberStyle.ListIndex

        End With
    End Sub

    Public Shared Sub Notify(ByVal Msg As String, Optional ByVal seconds As Int32 = 1)
        'ASFMAIN1.NotifyIcon1.BalloonTipText = Msg
        'ASFMAIN1.NotifyIcon1.ShowBalloonTip(seconds)
    End Sub

    Public Shared Sub Progress(ByVal Msg1 As String, Optional ByVal msg2 As String = "")

        If Msg1 = "-" Then
            ASFMAIN1.UltraStatusBar1.Panels("MSG2").Text = msg2
        Else
            ASFMAIN1.UltraStatusBar1.Panels("MSG1").Text = Msg1
            ASFMAIN1.UltraStatusBar1.Panels("MSG2").Text = msg2
        End If

        With ASFMAIN1.UltraStatusBar1.Panels("MSG1").Appearance
            If Msg1 <> "" Then
                '.BackColor = Color.Blue
                .BackColor = Color.Firebrick
                .ForeColor = Color.White

                .BackColor2 = Color.DarkOrange
                '.BackColor2 = Color.LightBlue
                '.ForeColor = Color.Black
                .BackGradientStyle = GradientStyle.ForwardDiagonal
            Else
                .BackColor = Color.Empty
                .BackColor2 = Color.Empty
                .BackGradientStyle = GradientStyle.None
            End If
        End With

        'If Msg1 <> "" Then
        '    ASFMAIN1.animationControl1.Left = 0 ' ASFMAIN1.Width - ASFMAIN1.animationControl1.Width
        '    ASFMAIN1.animationControl1.Top = 0 ' ASFMAIN1.Height - ASFMAIN1.animationControl1.Height
        '    ASFMAIN1.animationControl1.Play()
        '    ASFMAIN1.animationControl1.Visible = True
        'Else
        '    ASFMAIN1.animationControl1.Stop()
        '    ASFMAIN1.animationControl1.Visible = False
        '    ASFMAIN1.animationControl1.Left = -1000
        'End If

        ' ENABLING THIS MIGHT BE VERY DANGEROUS
        'WHEN i AM IN A TIGHLYBOUND LOOP (LIKE BRINGING IN TRACINGPOINTS DATA)
        '- WHERE i CLICK A BUTTON ON THE SCREEN - WHY DOES PROCESSING JUMP
        'DO TO PERFORM THE CODE BEHIND THE BUTTON - DOES IT JUST LEAVE THE
        'LOOP FOR A MOMENT?

        'Application.DoEvents()

        ' this appears to do the trick nicely
        ASFMAIN1.UltraStatusBar1.Refresh()
    End Sub

    Public Shared Function Print_REPORT_NO(
    ByVal REPORT_NO As String,
    Optional ByVal PRINTER_NAME As String = "")

        Try
            Dim clsASCBASE1 As New ABSolution.ASCBASE1
            clsASCBASE1.F = New ABSolution.ASFSRPTV
            clsASCBASE1.F.PRINTER_NAME = PRINTER_NAME

            ASCMAIN1.CR_RPT = New CrystalDecisions.CrystalReports.Engine.ReportDocument

            Dim reportFileName As String = ASCMAIN1.DBS_COMPANY & "_" & REPORT_NO & ".RPT"
            Dim tempFolder As String = ASCMAIN1.Folders("Temp")
            Dim reportFolder As String = ASCMAIN1.Folders("Archive") & "Reports\"

            If Not System.IO.File.Exists(tempFolder & reportFileName) Then
                System.IO.File.Copy(reportFolder & Mid(reportFileName, 1, 9) & "\" & reportFileName, tempFolder & reportFileName)
            End If

            clsASCBASE1.F.Print_Report(REPORT_NO)
            System.IO.File.Delete(tempFolder & reportFileName)

            clsASCBASE1 = Nothing

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

    Public Shared Function Excel_Cell(ByVal R As Integer, ByVal C As Integer) As String

        Dim AZ As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim RC As String = Mid(AZ, ((C - 1) Mod 26) + 1, 1)
        If C > 26 Then
            RC = Mid(AZ, (C - 1) \ 26, 1) & RC
        End If
        If R <> 0 Then
            RC &= CStr(R)
        End If

        Return RC

    End Function

    Public Shared Function Excel_Sheet_Name(
    ByVal Sheet_Name As String) As String

        Dim i As Integer
        Dim z As String

        Dim ILLEGAL_CHARACTERS As String
        ILLEGAL_CHARACTERS = ":\/?*[]"

        For i = 1 To Len(ILLEGAL_CHARACTERS)
            z = Mid$(ILLEGAL_CHARACTERS, i, 1)
            Do While InStr(Sheet_Name, z) <> 0
                Mid$(Sheet_Name, InStr(Sheet_Name, z), 1) = "_"
            Loop
        Next i
        Excel_Sheet_Name = Mid$(Sheet_Name, 1, 31)
    End Function
    Public Shared Function Export_To_CSV(sql As String, expRows As Integer, Optional fn As String = "DataExtract") As String
        Dim columnSql As String = "Select * from (" & sql & ") where rownum < 1"
        ASCMAIN1.sql = columnSql
        Dim dtc As DataTable = ASCDATA1.GetDataTable

        ASCMAIN1.oraCmd.CommandText = sql

        Dim reader As OracleDataReader = ASCMAIN1.oraCmd.ExecuteReader

        Dim FILENAME As String = ASCMAIN1.Folders("Work") & fn & ".CSV"
        If My.Computer.FileSystem.FileExists(FILENAME) Then
            My.Computer.FileSystem.DeleteFile(FILENAME)
        End If

        ASCMAIN1.Progress("Now Creating CSV")
        Dim rowsProcessed As Integer = 0
        Dim pctProcessed As String = ""
        Dim sep As String = ","
        Using sw As New System.IO.StreamWriter(FILENAME)
            Dim headings As String = ""
            Dim DEL As String = Chr(34)
            For i As Integer = 0 To dtc.Columns.Count - 1
                headings &= "," & DEL & dtc.Columns(i).ColumnName & DEL
            Next
            headings = Mid(headings, 2)
            sw.WriteLine(headings)
            Dim fields As Integer = reader.FieldCount - 1
            While reader.Read()
                Dim record As String = ""
                Dim i As Integer = 0
                While i <= fields
                    If dtc.Columns(i).DataType.ToString = "System.String" Then
                        DEL = Chr(34)
                    Else
                        DEL = ""
                    End If
                    record &= "," & DEL & reader(i) & DEL
                    i += 1
                End While
                record = Mid(record, 2)
                sw.WriteLine(record)
                rowsProcessed += 1
                pctProcessed = Math.Round(100 * (rowsProcessed / expRows), 2) & "%"
                ASCMAIN1.Progress("Now Creating CSV", pctProcessed)
            End While

        End Using
        ASCMAIN1.Progress("", "")

        Return FILENAME

        ASCMAIN1.Progress("")
    End Function
    Public Shared Function FormatTel(ByVal TEL As String, Optional ByVal EXT As String = "") As String
        FormatTel = "(" & Mid(TEL, 1, 3) & ")" & Mid(TEL, 4, 3) & "-" & Mid(TEL, 7, 4)
        If EXT <> "" Then
            FormatTel &= " x" & EXT
        End If
    End Function

    Public Shared Sub AnalyzeTable(
    ByVal TABLE_NAME As String,
    Optional ByVal SCHEMA As String = "")
        If SCHEMA = "" Then
            SCHEMA = ASCMAIN1.DBS_COMPANY
        End If
        If ASCMAIN1.DBS_TYPE = DBS_TYPE_types.SQLServer Then
        Else
            'ASCDATA1.ExecuteSQL("Analyze Table " & TABLE_NAME & " Compute Statistics")
            ASCDATA1.ExecuteSQL("Begin dbms_stats.gather_table_stats('" & Chr(34) & SCHEMA & Chr(34) & "','" & Chr(34) & TABLE_NAME & Chr(34) & "', CASCADE=>TRUE, METHOD_OPT=>'FOR ALL COLUMNS SIZE 1'); End;")

        End If
    End Sub

    Public Shared Function DEC(ByVal a As String) As String
        DEC = ""
        Dim i As Integer

        If a <> "" Then
            For i = 1 To Len(a)
                DEC = DEC & " " & Format(Asc(Mid(a, i, 1)), "000")
            Next
            DEC = Mid(DEC, 2)
        End If
    End Function

    Public Shared Function HTA(ByVal a As String) As String
        HTA = ""
        Dim i As Integer
        Dim HTA1 As String
        Dim HTA2 As String
        Dim z As String
        Dim d As Integer
        Dim ASCII As String = "0123456789ABCDEF"

        If a <> "" Then
            For i = 1 To Len(a)
                z = Mid(a, i, 1)
                d = Asc(z)
                HTA2 = Mid(ASCII, (d Mod 16) + 1, 1)
                HTA1 = Mid(ASCII, 1 + (d \ 16), 1)
                HTA = HTA & " " & HTA1 & HTA2
            Next
            HTA = Mid(HTA, 2)
        End If
    End Function

    Public Shared Function ATH(ByVal a As String) As Long
        ATH = 0
        Dim HEX As String = "0123456789ABCDEF"
        If Len(a) Mod 2 = 0 Then
            Dim wATH As Long = 0
            Dim pairs As Integer = Len(a) / 2
            For i As Integer = pairs To 1 Step -1
                Dim pair As String = UCase(Mid(a, (i - 1) * 2 + 1, 2))
                Dim H1 As Integer = InStr(HEX, Mid(pair, 1, 1))
                Dim H2 As Integer = InStr(HEX, Mid(pair, 2, 1))
                If H1 = 0 Or H2 = 0 Then
                    Exit Function
                End If
                wATH += ((H1 - 1) * 16 + (H2 - 1) * 1) * 256 ^ (pairs - i)
            Next
            ATH = wATH
        End If
    End Function

    Public Shared Function Get_where_from_Filter(ByVal grd As UltraWinGrid.UltraGrid) As String

        Dim sqlwhere As String = ""
        Dim sqlwhere_segment As String = ""

        For Each U As UltraWinGrid.ColumnFilter In grd.DisplayLayout.Bands(0).ColumnFilters
            For Each F As UltraWinGrid.FilterCondition In U.FilterConditions
                Dim FCOMP As String = F.CompareValue.ToString.ToUpper
                FCOMP = Replace(FCOMP, "'", "''")
                If F.CompareValue.GetType Is GetType(System.DateTime) Then
                    FCOMP = Format(CDate(FCOMP), "dd-MMM-yyyy")
                End If

                Select Case F.ComparisionOperator
                    Case UltraWinGrid.FilterComparisionOperator.StartsWith ' 10 ' StartsWith
                        sqlwhere_segment = " and UPPER(" & U.Column.Key & ") like '" & FCOMP & "%'"
                    Case UltraWinGrid.FilterComparisionOperator.Contains
                        sqlwhere_segment = " and UPPER(" & U.Column.Key & ") like '%" & FCOMP & "%'"

                    Case UltraWinGrid.FilterComparisionOperator.Equals
                        If FCOMP = "(BLANKS)" AndAlso F.CompareValue.ToString = "(Blanks)" Then
                            sqlwhere_segment = " and UPPER(" & U.Column.Key & ") IS NULL"

                        Else
                            sqlwhere_segment = " and UPPER(" & U.Column.Key & ") = '" & FCOMP & "'"
                        End If
                    Case UltraWinGrid.FilterComparisionOperator.GreaterThan
                        sqlwhere_segment = " and UPPER(" & U.Column.Key & ") > '" & FCOMP & "'"
                    Case UltraWinGrid.FilterComparisionOperator.GreaterThanOrEqualTo
                        sqlwhere_segment = " and UPPER(" & U.Column.Key & ") >= '" & FCOMP & "'"
                    Case UltraWinGrid.FilterComparisionOperator.LessThan
                        sqlwhere_segment = " and UPPER(" & U.Column.Key & ") < '" & FCOMP & "'"
                    Case UltraWinGrid.FilterComparisionOperator.LessThanOrEqualTo
                        sqlwhere_segment = " and UPPER(" & U.Column.Key & ") <= '" & FCOMP & "'"

                    Case UltraWinGrid.FilterComparisionOperator.EndsWith
                        sqlwhere_segment = " and UPPER(" & U.Column.Key & ") like '%" & FCOMP & "'"

                    Case Else
                        'Stop
                        MsgBox("ABS needs to code for this filter - NOT SUPPORTED")
                        '  Stop
                        sqlwhere_segment = " AND ROWNUM < 1"
                        ' sqlwhere_segment = " and UPPER(" & U.Column.Key & ") like '%" & FCOMP & "%'"

                End Select

                sqlwhere_segment = ASCMAIN1.TACMAIN1.Custom_sqlwhere(sqlwhere_segment, grd, U.Column.Key)

                sqlwhere &= sqlwhere_segment
            Next
        Next


        Return sqlwhere
    End Function

    Public Shared Function CheckDigitUPC(ByVal ValueToCheckDigit As String) As String

        ' Note: Check Digit Calculation applies to the 11-digits prior to the check digit
        '       These 11 digits are usually made up from the 6 digit Vendor ID prepended to the 5 digit UPC Serial Number

        Dim odd_digits As Integer = 0
        Dim even_digits As Integer = 0

        For i As Integer = 1 To Len(ValueToCheckDigit) Step 2
            odd_digits = odd_digits + Val(Mid$(ValueToCheckDigit, i, 1))
            If Len(ValueToCheckDigit) > i Then
                even_digits = even_digits + Val(Mid$(ValueToCheckDigit, i + 1, 1))
            End If
        Next i

        Dim check_digit As Integer
        check_digit = (odd_digits * 3 + even_digits) Mod 10
        If check_digit <> 0 Then
            check_digit = 10 - check_digit
        End If

        Return Format$(check_digit, "0")
    End Function

    Public Shared Function CheckDigit(
    ByVal ValueToCheckDigit As String) As String

        ' Mod 10 from PNC

        Dim product As Integer = 0
        Dim factor As Integer = 2

        For i As Integer = Len(ValueToCheckDigit) To 1 Step -1
            Dim cell As String = CStr(Val(Mid(ValueToCheckDigit, i, 1)) * factor)
            For j As Integer = 1 To Len(cell)
                product += Val(Mid(cell, j, 1))
            Next
            If factor = 2 Then factor = 1 Else factor = 2
        Next i

        Dim check_digit As Integer = (10 - (product Mod 10)) Mod 10
        Return Format$(check_digit, "0")

    End Function


    ' we are missing a few functions (like AscToDec which would return a Long, with options for signed and unsigned integer modes) 
    '  to make this library complete
    ' we will write them as we need them


    ''' <summary>
    ''' This function will accept an Ascii string (ex:"12") and return the Decimal equivalent of each character ("049050" or "049 050", depending on the optional withSpaces argument)
    ''' </summary>
    ''' <param name="AsciiString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AscToDecBytes(
    ByVal AsciiString As String,
    Optional ByVal withSpaces As Boolean = False) As String

        AscToDecBytes = ""
        Dim i As Integer

        Dim delimiter As String = ""
        If withSpaces Then
            delimiter = " "
        End If

        If AsciiString <> "" Then
            For i = 1 To AsciiString.Length
                AscToDecBytes = AscToDecBytes & delimiter & Format(Asc(Mid(AsciiString, i, 1)), "000")
            Next
            If delimiter.Length <> 0 Then
                AscToDecBytes = Mid(AscToDecBytes, delimiter.Length + 1)
            End If
        End If
    End Function

    ''' <summary>
    ''' This function will accept a Hex Representation of a String (ex:"3132") and return the String ("12")
    ''' </summary>
    ''' <param name="HexString"></param>
    ''' <returns></returns>
    ''' <remarks>Input string should be an even number of characters in length, or else an empty string "" will be returned.  If invalid characters are found in the input string, an empty string "" will be returned.</remarks>
    Public Shared Function HexToAsc(ByVal HexString As String) As String
        HexToAsc = ""
        Dim HEX As String = "0123456789ABCDEF"
        If HexString.Length Mod 2 = 0 Then
            Dim pairs As Integer = Len(HexString) / 2
            For i As Integer = 1 To pairs
                Dim pair As String = UCase(Mid(HexString, (i - 1) * 2 + 1, 2))
                Dim H1 As Integer = InStr(HEX, Mid(pair, 1, 1))
                Dim H2 As Integer = InStr(HEX, Mid(pair, 2, 1))
                If H1 = 0 Or H2 = 0 Then
                    HexToAsc = ""
                    Exit Function
                End If
                HexToAsc &= Chr(((H1 - 1) * 16 + (H2 - 1) * 1))
            Next
        End If
    End Function

    ''' <summary>
    ''' This function will accept an ASCII String (ex:"12") and return the Hexadecimal representation of that string ("3131" or "31 32", depending on the optional withSpaces argument)
    ''' </summary>
    ''' <param name="AsciiString"></param>
    ''' <param name="withSpaces"></param>
    ''' <returns>Hexadecimal representation of that string</returns>
    ''' <remarks>Input string may be any length</remarks>
    Public Shared Function AscToHex(
    ByVal AsciiString As String,
    Optional ByVal withSpaces As Boolean = False) As String

        AscToHex = ""
        Dim i As Integer
        Dim HTA1 As String
        Dim HTA2 As String
        Dim z As String
        Dim d As Integer
        Dim ASCII As String = "0123456789ABCDEF"

        Dim delimiter As String = ""
        If withSpaces Then
            delimiter = " "
        End If

        If AsciiString <> "" Then
            For i = 1 To Len(AsciiString)
                z = Mid(AsciiString, i, 1)
                d = Asc(z)
                HTA2 = Mid(ASCII, (d Mod 16) + 1, 1)
                HTA1 = Mid(ASCII, 1 + (d \ 16), 1)
                AscToHex = AscToHex & delimiter & HTA1 & HTA2
            Next
            If delimiter.Length <> 0 Then
                AscToHex = Mid(AscToHex, delimiter.Length + 1)
            End If
        End If
    End Function

    ''' <summary>
    ''' This function will accept a Hexadecimal string (ex:"31") and return its Decimal equivalent (49)
    ''' </summary>
    ''' <param name="HexString"></param>
    ''' <returns></returns>
    ''' <remarks>Input string should be an even number of characters in length, or else a 0 value will be returned.  If invalid characters are found in the input string, a 0 value will be returned.</remarks>
    Public Shared Function HexToDec(ByVal HexString As String) As Long
        HexToDec = 0
        Dim HEX As String = "0123456789ABCDEF"
        If HexString.Length Mod 2 = 0 Then
            Dim pairs As Integer = Len(HexString) / 2
            For i As Integer = pairs To 1 Step -1
                Dim pair As String = UCase(Mid(HexString, (i - 1) * 2 + 1, 2))
                Dim H1 As Integer = InStr(HEX, Mid(pair, 1, 1))
                Dim H2 As Integer = InStr(HEX, Mid(pair, 2, 1))
                If H1 = 0 Or H2 = 0 Then
                    HexToDec = 0
                    Exit Function
                End If
                HexToDec += ((H1 - 1) * 16 + (H2 - 1) * 1) * 256 ^ (pairs - i)
            Next
        End If
    End Function

    Public Shared Function Launch_Attachment(ByVal ATTACHMENT_NO As String, ByVal ATTACHMENT_TYPE As String)

        Dim ataFileName As String = ASCMAIN1.Folders("Attach") & ATTACHMENT_NO
        If Not My.Computer.FileSystem.FileExists(ataFileName) Then
            'If ASCMAIN1.USER_ID = "wjz" Then MsgBox("Cannot Find File " & ataFileName)
            Return "Cannot Find File " & ataFileName
            Exit Function
        Else
            Dim finfo As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(ataFileName)
            ataFileName = finfo.FullName
        End If

        Select Case ATTACHMENT_TYPE

            Case "XLS", "XLSX"
                Try
                    'Infragistics.Documents.Excel.BIFF8Writer.WriteWorkbookToFile(myWorkbook, ASCMAIN1.Folders("Work") & xlsFileName & ".xls")
                    'xlsFileName_sfx = ""
                    Dim xlsFileName As String = ASCMAIN1.Folders("Work") & ATTACHMENT_NO & "." & ATTACHMENT_TYPE ' ".xls"
                    My.Computer.FileSystem.CopyFile(ataFileName, xlsFileName, True)

                    Dim excel As New Process
                    excel.StartInfo.Arguments = """" + ATTACHMENT_NO + """ /e"
                    excel.StartInfo.FileName = xlsFileName
                    excel.Start()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error trying to Launch Attachment")
                End Try


            Case "CSV"
                Try
                    Dim xlsFileName As String = ASCMAIN1.Folders("Work") & ATTACHMENT_NO & ".csv"
                    My.Computer.FileSystem.CopyFile(ataFileName, xlsFileName, True)

                    Dim excel As New Process
                    excel.StartInfo.Arguments = """" + ATTACHMENT_NO + """ /e"
                    excel.StartInfo.FileName = xlsFileName
                    excel.Start()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error trying to Launch Attachment")
                End Try

            Case "EML"
                Dim appFileName As String = ASCMAIN1.Folders("Work") & ATTACHMENT_NO & "." & ATTACHMENT_TYPE
                Try
                    'If ASCMAIN1.USER_ID = "wjz" Then MsgBox($"Copy {ataFileName} to {appFileName}")
                    My.Computer.FileSystem.CopyFile(ataFileName, appFileName, True)
                    appFileName = My.Computer.FileSystem.GetFileInfo(appFileName).FullName
                    ASCMAIN1.ActiveForm.Show_Document(appFileName)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error trying to Launch EML Attachment")
                End Try

            Case "PDF", "BMP", "TXT", "JPG", "MSG", "DOC", "DOCX", "PNG"
                Dim appFileName As String = ASCMAIN1.Folders("Work") & ATTACHMENT_NO & "." & ATTACHMENT_TYPE

                Try

                    My.Computer.FileSystem.CopyFile(ataFileName, appFileName, True)
                    appFileName = My.Computer.FileSystem.GetFileInfo(appFileName).FullName

                    Dim p As Process = Process.Start(appFileName)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error trying to Launch Attachment")
                End Try

            Case Else
                Return "Unknown Application (" & ATTACHMENT_TYPE & ")"
                Exit Function
        End Select

        Return ""

    End Function

    Public Shared Function GetPath(ByVal path As String) As String
        GetPath = path
        Try
            Dim fullname As String = My.Computer.FileSystem.GetDirectoryInfo(path).FullName
            If fullname <> "" Then
                GetPath = fullname
            End If
        Catch ex As Exception

        End Try
    End Function

    Public Shared Sub LoadKeysFromPopup(ByVal ITEM_KEY As String, ByVal ITEM_TEXT As String, ByVal ITEM_TOOLTIP As String)
        If ASFMAIN1.tblImages.PrimaryKey Is Nothing Then Exit Sub
        Dim tbl As DataTable = ASFMAIN1.tblImages

        If ASFMAIN1.tblImages IsNot Nothing AndAlso ASFMAIN1.tblImages.Rows IsNot Nothing _
             AndAlso ASFMAIN1.tblImages.PrimaryKey.Length > 0 _
            AndAlso ASFMAIN1.tblImages.Rows.Find(ITEM_KEY) Is Nothing Then
            ASFMAIN1.tblImages.Rows.Add(New String() {ITEM_KEY, ITEM_TEXT, ITEM_TOOLTIP})
        End If
    End Sub

    Public Shared Function GetImageData(ByVal fileName As String) As Byte()
        'Method to load an image from disk and return it as a bytestream
        Dim fs As System.IO.FileStream =
        New System.IO.FileStream(fileName,
        System.IO.FileMode.Open, System.IO.FileAccess.Read)
        Dim br As System.IO.BinaryReader = New System.IO.BinaryReader(fs)
        ' Return (br.ReadBytes(Convert.ToInt32(br.BaseStream.Length)))
        Return (br.ReadBytes(br.BaseStream.Length))
    End Function

    Public Shared Function Add_Value_List(
    ByVal grd As UltraWinGrid.UltraGrid,
    ByVal COLUMN_NAME As String,
    ByVal sql As String) As ValueList

        Return Add_Value_List(grd, COLUMN_NAME, , , , sql)

    End Function

    Public Shared Function Add_Value_List(
    ByVal cbe As UltraWinEditors.UltraComboEditor,
    ByVal COLUMN_NAME As String,
    ByVal sql As String) As ValueList

        Return Add_Value_List(cbe, COLUMN_NAME, , , sql)

    End Function

    Public Shared Function Add_Value_List(
    ByVal grd As UltraWinGrid.UltraGrid,
    ByVal TABLE_COLUMN As String,
    Optional ByVal RemoveCodes() As String = Nothing,
    Optional ByVal AddCodes() As String = Nothing,
    Optional ByVal BandIndex As Integer = 0,
    Optional ByVal sql As String = "") As ValueList

        Dim VL As ValueList = ValueListFor(TABLE_COLUMN, RemoveCodes, AddCodes, sql)

        Dim TABLE_NAME As String = ""
        Dim COLUMN_NAME As String = TABLE_COLUMN
        If InStr(COLUMN_NAME, ".") <> 0 Then
            TABLE_NAME = Split(TABLE_COLUMN, ".")(0)
            COLUMN_NAME = Split(TABLE_COLUMN, ".")(1)
        End If

        With grd.DisplayLayout
            If .ValueLists.Exists(COLUMN_NAME) Then
                .ValueLists.Remove(COLUMN_NAME)
            End If

            With .ValueLists.Add(COLUMN_NAME)
                For Each VLI As ValueListItem In VL.ValueListItems
                    .ValueListItems.Add(VLI.DataValue, VLI.DisplayText)
                Next
            End With

            .Bands(BandIndex).Columns(COLUMN_NAME).ValueList = .ValueLists(COLUMN_NAME)
            .Bands(BandIndex).Columns(COLUMN_NAME).Style = UltraWinGrid.ColumnStyle.DropDownList
        End With

        Return VL

    End Function

    Public Shared Function Add_Value_List(
    ByVal cbe As UltraWinEditors.UltraComboEditor,
    ByVal TABLE_COLUMN As String,
    Optional ByVal RemoveCodes() As String = Nothing,
    Optional ByVal AddCodes() As String = Nothing,
    Optional ByVal sql As String = "") As ValueList

        Dim VL As ValueList = ValueListFor(TABLE_COLUMN, RemoveCodes, AddCodes, sql)
        With cbe.Items.ValueList
            For Each VLI As ValueListItem In VL.ValueListItems
                .ValueListItems.Add(VLI.DataValue, VLI.DisplayText)
            Next
        End With

        Return VL
    End Function

    Public Shared Function Add_Value_List(
    ByVal opt As UltraWinEditors.UltraOptionSet,
    ByVal TABLE_COLUMN As String,
    Optional ByVal RemoveCodes() As String = Nothing,
    Optional ByVal AddCodes() As String = Nothing,
    Optional ByVal sql As String = "") As ValueList

        Dim VL As ValueList = ValueListFor(TABLE_COLUMN, RemoveCodes, AddCodes, sql)
        With opt.Items.ValueList
            For Each VLI As ValueListItem In VL.ValueListItems
                .ValueListItems.Add(VLI.DataValue, VLI.DisplayText)
            Next
        End With

        Return VL
    End Function

    Public Shared Function Validate_User_Password(
        ByVal DecryptPassword As Boolean,
        ByVal USER_ID As String,
        ByVal USER_PASSWORD As String,
        ByVal rowASTPARMP As DataRow) As String

        If USER_PASSWORD.Length = 32 Then
            Return "" ' password already encrypted
        End If

        ' Parameters
        ' DecryptPassword - If True then Decrypt the Password
        ' UserID - If provided then only process this user
        ' User_Password - If Provided then only validate this Password

        Dim EMsg As String = String.Empty

        Dim has_alpha As Boolean = False
        Dim has_numeric As Boolean = False
        Dim has_upper As Boolean = False
        Dim has_lower As Boolean = False
        Dim has_non_an As Boolean = False

        Dim AS_PARM_PWD_MIN_LEN As Integer = Val(rowASTPARMP.Item("AS_PARM_PWD_MIN_LEN") & "")
        Dim AS_PARM_PWD_REQ_MIX_AN As String = rowASTPARMP.Item("AS_PARM_PWD_REQ_MIX_AN") & ""
        Dim AS_PARM_PWD_REQ_MIX_CASE As String = rowASTPARMP.Item("AS_PARM_PWD_REQ_MIX_CASE") & ""
        Dim AS_PARM_PWD_REQ_MIX_NON_AN As String = rowASTPARMP.Item("AS_PARM_PWD_REQ_MIX_NON_AN") & ""
        Dim AS_PARM_PWD_NO_USER_ID As String = rowASTPARMP.Item("AS_PARM_PWD_NO_USER_ID") & ""
        Dim AS_PARM_PWD_NO_USER_ID_PERM As String = rowASTPARMP.Item("AS_PARM_PWD_NO_USER_ID_PERM") & ""
        Dim AS_PARM_PWD_REUSE As Integer = Val(rowASTPARMP.Item("AS_PARM_PWD_REUSE") & "")
        Dim AS_PARM_PWD_ENCRYPTED As String = rowASTPARMP.Item("AS_PARM_PWD_ENCRYPTED") & ""

        If USER_ID <> USER_ID.ToLower Then
            EMsg &= vbCr & "User ID must be all Lower Case"
        End If

        If USER_PASSWORD.Length < AS_PARM_PWD_MIN_LEN Then
            EMsg &= vbCr & "Minimum Length (must be at least " & CStr(AS_PARM_PWD_MIN_LEN) & " characters)"
        End If

        Dim Usorted As String = X_Sort(USER_ID, has_alpha, has_numeric, has_upper, has_lower, has_non_an)
        Dim Psorted As String = X_Sort(USER_PASSWORD, has_alpha, has_numeric, has_upper, has_lower, has_non_an)

        If AS_PARM_PWD_REQ_MIX_AN = "1" AndAlso Abs(has_alpha + has_numeric + has_non_an) < 2 Then
            EMsg &= vbCr & "Password may not be pure alpha or pure numeric."
        End If

        If AS_PARM_PWD_REQ_MIX_CASE = "1" AndAlso Abs(has_upper + has_lower) < 2 Then
            EMsg &= vbCr & "Password requires mix case."
        End If

        If AS_PARM_PWD_REQ_MIX_NON_AN = "1" AndAlso Not has_non_an Then
            EMsg &= vbCr & "Password requires special characters."
        End If

        If USER_PASSWORD.Contains("'") Then
            EMsg &= vbCr & "Password may not use the apostrophe special character."
        End If

        If USER_ID = USER_PASSWORD Then
            If AS_PARM_PWD_NO_USER_ID = "1" Then
                EMsg &= vbCr & "Password may not be the same as the User ID."
            End If
        Else
            If AS_PARM_PWD_NO_USER_ID_PERM = "1" AndAlso Usorted.ToUpper = Psorted.ToUpper Then
                EMsg &= vbCr & "Password may not be a permutation of User ID."
            End If
        End If

        If AS_PARM_PWD_REUSE > 0 Then
            Dim sqlx As String = $"Select * From ASTUSERP  
                Where USER_ID = '{USER_ID}' 
                And UPPER(USER_PASSWORD) = '{USER_PASSWORD.ToUpper}'
                And INIT_DATE > '{Format(DateAdd(DateInterval.Day, -1 * AS_PARM_PWD_REUSE, Now), "dd-MMM-yyyy")}'"
            Dim dr As DataRow = ASCDATA1.GetDataRow(sqlx)
            If Not dr Is Nothing Then
                EMsg &= vbCr & "Password has already been used in the past."
            End If
        End If

        If EMsg.Length = 0 Then
            If AS_PARM_PWD_ENCRYPTED = "1" Then
                USER_PASSWORD = ASCMAIN1.EncryptAES(USER_PASSWORD)
            End If
        End If

        Return EMsg

    End Function

    Public Shared Function X_Sort(
    ByRef a As String,
    ByRef has_alpha As Boolean,
    ByRef has_numeric As Boolean,
    ByRef has_upper As Boolean,
    ByRef has_lower As Boolean,
    ByRef has_non_an As Boolean) As String

        Dim z As String
        Dim i As Integer
        Dim j As Integer

        Dim b As String
        b = ""

        has_alpha = False
        has_numeric = False
        has_upper = False
        has_lower = False
        has_non_an = False

        For i = 1 To Len(a)
            z = Mid$(a, i, 1)

            If UCase$(z) >= "A" And UCase$(z) <= "Z" Then
                has_alpha = True
            End If

            If z >= "0" And z <= "9" Then
                has_numeric = True
            End If

            If z >= "A" And z <= "Z" Then
                has_upper = True
            End If

            If z >= "a" And z <= "z" Then
                has_lower = True
            End If

            If (UCase$(z) < "A" Or UCase$(z) > "Z") And (z < "0" Or z > "9") Then
                has_non_an = True
            End If

            If b = "" Then
                b = z
            Else
                If z >= Right$(b, 1) Then
                    b = b & z
                ElseIf z <= Left$(b, 1) Then
                    b = z & b
                Else
                    For j = 1 To Len(b)
                        If z <= Mid$(b, j, 1) Then
                            b = Left$(b, j - 1) & z & Mid$(b, j)
                            Exit For
                        End If
                    Next j
                End If
            End If
        Next i

        X_Sort = b
    End Function

    Public Shared Function ValueListFor(
        ByVal TABLE_COLUMN As String,
        Optional ByVal RemoveCodes() As String = Nothing,
        Optional ByVal AddCodes() As String = Nothing,
        Optional ByVal sql As String = "") As ValueList

        Dim TABLE_NAME As String = ""
        Dim COLUMN_NAME As String = TABLE_COLUMN
        If InStr(COLUMN_NAME, ".") <> 0 Then
            TABLE_NAME = Split(TABLE_COLUMN, ".")(0)
            COLUMN_NAME = Split(TABLE_COLUMN, ".")(1)
        End If

        Dim CVL As Dictionary(Of String, String) = TACMAIN1.CodeValues(TABLE_COLUMN)
        If CVL Is Nothing OrElse CVL.Count = 0 Then
            Dim FORM_NAME As String = TABLE_NAME
            If ASCMAIN1.ActiveForm IsNot Nothing And TABLE_NAME = "" Then FORM_NAME = ASCMAIN1.ActiveForm.Name
            Dim SQL_Code As String = sql
            Dim tbl As New DataTable
            If SQL_Code = "" Then SQL_Code = ASCMAIN1.TACMAIN1.Get_Code_SQL_X(FORM_NAME, COLUMN_NAME, "")
            If SQL_Code = "" Then
                Dim sqlc As String = "Select T_CODE, T_DESC, TABLE_NAME from ASTCODE1 where COLUMN_NAME = '" & COLUMN_NAME & "'"
                If TABLE_NAME <> "" Then
                    tbl = ASCDATA1.GetDataTable(sqlc & " and TABLE_NAME = '" & TABLE_NAME & "'")
                End If
                If tbl Is Nothing OrElse tbl.Rows.Count = 0 Then
                    Dim PFX As String = ""
                    If TABLE_COLUMN = "" Then
                        PFX = Mid(TABLE_NAME, 1, 2)
                    Else
                        PFX = Mid(FORM_NAME, 1, 2)
                    End If
                    tbl = ASCDATA1.GetDataTable(sqlc & " and TABLE_NAME like '" & PFX & "%'")
                    If tbl Is Nothing OrElse tbl.Rows.Count = 0 Then
                        tbl = ASCDATA1.GetDataTable(sqlc)
                    End If
                    If tbl IsNot Nothing AndAlso tbl.Rows.Count <> 0 Then
                        Dim TABLE_NAME_SELECTED As String = tbl.Rows(0).Item("TABLE_NAME")
                        For i As Integer = tbl.Rows.Count - 1 To 0 Step -1
                            If tbl.Rows(i).Item("TABLE_NAME") <> TABLE_NAME_SELECTED Then
                                tbl.Rows(i).Delete()
                            End If
                        Next
                    End If
                End If
            Else
                tbl = ASCDATA1.GetDataTable(SQL_Code)
            End If

            If tbl IsNot Nothing AndAlso tbl.Rows.Count <> 0 Then

                For Each row As DataRow In tbl.Rows
                    If Not CVL.ContainsKey(row.Item(0) & "") Then
                        CVL.Add(row.Item(0) & "", row.Item(1) & "")
                    End If
                Next
            End If
        End If

        Dim VL As New ValueList

        Dim xVL As New List(Of String)
        If RemoveCodes IsNot Nothing AndAlso RemoveCodes.Length > 0 Then
            For Each dataValue As String In RemoveCodes
                xVL.Add(dataValue)
            Next
        End If

        With VL.ValueListItems
            For Each dataValue As String In CVL.Keys
                If Not xVL.Contains(dataValue) Then
                    Dim displayText As String = CVL(dataValue)
                    .Add(dataValue, displayText)
                End If
            Next

            If AddCodes IsNot Nothing AndAlso AddCodes.Length > 1 Then
                Dim splitChar As String = AddCodes(0)
                For i As Integer = 1 To AddCodes.Length - 1
                    Dim datavalue As String = Split(AddCodes(i), splitChar)(0)
                    Dim displaytext As String = Split(AddCodes(i) & splitChar, splitChar)(1)
                    If displaytext = "" Then displaytext = datavalue
                    .Add(datavalue, displaytext)
                Next
            End If
        End With

        Return VL
    End Function

    Public Shared Function Launch_Form(
    ByVal MENU_ITEM_OBJECT As String,
    Optional ByVal MENU_ITEM_TYPE As String = "",
    Optional ByVal MENU_ID As String = "") As ASFBASE1

        ASCMAIN1.sql = "Select * from ASTMENU1" _
        & " where MENU_ITEM_OBJECT = '" & MENU_ITEM_OBJECT & "'"
        If MENU_ITEM_TYPE <> "" Then
            ASCMAIN1.sql &= " and MENU_ITEM_TYPE = '" & MENU_ITEM_TYPE & "'"
        End If
        If MENU_ID <> "" Then
            ASCMAIN1.sql &= " and MENU_ID = '" & MENU_ID & "'"
        Else
            ASCMAIN1.sql &= " and MENU_ID IN (" & "'" & Replace(USER_SECURITY_CODEs, ",", "','") & "'" & ")"
        End If

        Dim rowASTMENU1 As DataRow = ASCDATA1.GetDataRow

        If rowASTMENU1 Is Nothing Then
            Dim sql As String = "MENU_ITEM_OBJECT = '" & MENU_ITEM_OBJECT & "'"
            If MENU_ITEM_TYPE <> "" Then
                ASCMAIN1.sql &= " and MENU_ITEM_TYPE = '" & MENU_ITEM_TYPE & "'"
            End If
            If MENU_ID <> "" Then
                ASCMAIN1.sql &= " and MENU_ID = '" & MENU_ID & "'"
            End If
            Dim rowASTMENU1s() As DataRow = tblASTMENU1.Select(sql)
            If rowASTMENU1s.Length <> 0 Then
                rowASTMENU1 = rowASTMENU1s(0)
            End If
        End If

        If rowASTMENU1 Is Nothing Then
            Return Nothing
        Else
            Return ASFMAIN1.Launch_Form(rowASTMENU1)
        End If
    End Function

    Public Shared Function GemboxKey(Optional ByVal key As String = "") As String

        Select Case key
            Case ""
                Return "EW1Q-G14I-JKOW-4XS8" ' this key is ver version 3.7 
                ' Return "EMPX-L9BW-EL8E-4GKJ" ' this key is ver version 3.3
                ' Return "EFYZ-QQSH-LE5Q-NJ7Y" this is the old 3.1 key
            Case Else
                Return ""
        End Select
    End Function

    Public Shared Function nSoftwareKeys(ByVal key As String) As String

        Select Case key
            Case "nSoftwareZipkey"
                Return TACMAIN1.nSoftwareZipkey
            Case "nSoftwareftpkey"
                Return TACMAIN1.nSoftwareftpkey
            Case "nSoftwareipportkey"
                Return TACMAIN1.nSoftwareipportkey
            Case "nSoftwarehttpkey"
                Return TACMAIN1.nSoftwarehttpkey
            Case "nSoftwarepopkey"
                Return TACMAIN1.nSoftwarepopkey
            'Case "nSoftwareInship"
            '    Return TACMAIN1.nSoftwareInship
            Case "nSoftwaresftpkey"
                Return TACMAIN1.nSoftwaresftpkey
            Case "nSoftwareEncryptionkey"
                Return TACMAIN1.nSoftwareEncryptionkey
            Case Else
                Return ""
        End Select
    End Function

    Public Shared Sub Record_Event(
    ByVal TABLE_NAME As String,
    ByVal TABLE_KEY As String,
    ByVal TABLE_KEY2 As String,
    ByVal INIT_DATE As Date,
    ByVal INIT_OPER As String,
    ByVal EVENT_TYPE As String,
    ByVal EVENT_DESC As String,
    ByVal EVENT_KEY As String)
        If TABLE_KEY2 <> "" Then
            TABLE_KEY &= ":" & TABLE_KEY2
        End If
        ASCMAIN1.sql = "Insert into TATEVNT1 Values (:PARM1,:PARM2,:PARM3,:PARM4,:PARM5,:PARM6,:PARM7,:PARM8,:PARM9,:PARM10,:PARM11)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVDVVVVVVNV",
                            New Object() {TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY,
                                          ASCMAIN1.ActiveForm.Name, ASCMAIN1.SESSION_NO, ASCMAIN1.ActiveForm.SELECTION_NO, ASCMAIN1.ActiveForm.XNO})
    End Sub

    Public Shared Sub Get_Period_Range(
     ByVal Number_of_Periods As Integer,
     ByRef YP_Dates() As Date,
     ByRef YP_Periods(,) As String,
     Optional ByVal YP_Base As String = "")

        If YP_Base = "" Then
            YP_Base = ASCMAIN1.CYP
        End If

        Dim YP0 As String = YP_Base
        Dim YP1 As String = ASCMAIN1.Period_Calc(YP_Base, Number_of_Periods)

        Dim AD As String = ""
        If YP0 > YP1 Then
            YP0 = YP1
            YP1 = YP_Base
            AD = " DESC"
        End If

        Dim P As Integer = 0

        ReDim YP_Periods(Abs(Number_of_Periods), 1)
        ReDim YP_Dates(Abs(Number_of_Periods))

        ASCMAIN1.sql = "Select * from GLTPARM2 " _
        & " where OPS_YYYYPP between '" & YP0 & "' and '" & YP1 & "'"
        P = 0
        For Each rowGLTPARM2 As DataRow In ASCDATA1.GetDataTable.Select("", "OPS_YYYYPP" & AD)
            YP_Periods(P, 0) = rowGLTPARM2.Item("OPS_YYYYPP")
            YP_Periods(P, 1) = Mid(rowGLTPARM2.Item("LEGEND"), 10, 6)
            YP_Dates(P) = rowGLTPARM2.Item("PRD_END_DATE")
            P += 1
        Next
    End Sub

    Public Shared Sub Get_Week_Range(
     ByVal Number_of_Weeks As Integer,
     ByRef YW_Dates() As Date,
     ByRef YW_Weeks(,) As String,
     Optional ByVal YW_Base As String = "")

        If YW_Base = "" Then
            YW_Base = ASCMAIN1.CYW
        End If

        Dim YW0 As String = YW_Base
        Dim YW1 As String = ASCMAIN1.Week_Calc(YW_Base, Number_of_Weeks)

        Dim AD As String = ""
        If YW0 > YW1 Then
            YW0 = YW1
            YW1 = YW_Base
            AD = " DESC"
        End If

        Dim W As Integer = 0

        ReDim YW_Weeks(Abs(Number_of_Weeks), 1)
        ReDim YW_Dates(Abs(Number_of_Weeks))

        ASCMAIN1.sql = "Select * from GLTPARM3 " _
        & " where YYYYWW between '" & YW0 & "' and '" & YW1 & "'"
        W = 0
        For Each rowGLTPARM3 As DataRow In ASCDATA1.GetDataTable.Select("", "YYYYWW" & AD)
            YW_Weeks(W, 0) = rowGLTPARM3.Item("YYYYWW")
            YW_Weeks(W, 1) = Mid(rowGLTPARM3.Item("LEGEND"), 10, 7)
            YW_Dates(W) = rowGLTPARM3.Item("WEEK_END_DATE")
            W += 1
        Next
    End Sub

    Public Shared Sub Load_Popup_Menu(
    ByVal tlb As UltraWinToolbars.UltraToolbarsManager,
    ByVal ctl As Control,
    ByVal ToolTypes As String,
    ByVal ParamArray Tools() As String)

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool
        tlb_pop = New UltraWinToolbars.PopupMenuTool(ctl.Name)
        tlb.Tools.Add(tlb_pop)

        Dim TOOL_KEY As String
        Dim TOOL_CAPTION As String
        Dim IsFirstInGroup As Boolean = False

        Dim T As Integer = 0
        Dim subMenuCount As Integer = 0

        Dim tkeys As New List(Of String)
        If tlb.Tools.Count > 0 Then
            For i As Integer = 0 To tlb.Tools.Count - 1
                tkeys.Add(tlb.Tools(i).Key)
            Next
        End If

        For Each Tool As String In Tools

            IsFirstInGroup = False

            T += 1
            Dim ToolType As String = Mid$(ToolTypes, T, 1)
            Dim OptionSetKey As String = ""


            If subMenuCount > 0 Then
                If ToolType = "M" Then
                    subMenuCount += 1
                    Continue For
                ElseIf ToolType = "X" Then
                    subMenuCount -= 1
                    If subMenuCount = 0 And T < ToolTypes.Count Then
                        T += 1
                        ToolType = Mid$(ToolTypes, T, 1)
                    Else
                        Continue For
                    End If
                Else
                    Continue For
                End If
            End If

            If ToolType = "P" Then
                IsFirstInGroup = True
                T += 1
                ToolType = Mid$(ToolTypes, T, 1)
            End If

            If ToolType = "O" Then
                OptionSetKey = Split(Tool & ":", ":")(0)
                Tool = Split(Tool & ":", ":")(1)

                If Not tlb.OptionSets.Exists(OptionSetKey) Then
                    'Dim OptionSet1 As Infragistics.Win.UltraWinToolbars.OptionSet = New Infragistics.Win.UltraWinToolbars.OptionSet(OptionSetKey)
                    'OptionSet1.AllowAllUp = False
                    'Me.tlb.OptionSets.Add(OptionSet1)
                    Try
                        tlb.OptionSets.Add(False, OptionSetKey)
                    Catch ex As Exception

                    End Try
                End If


            End If

            TOOL_KEY = Split(Tool & "|", "|")(0)
            TOOL_CAPTION = Split(Tool & "|", "|")(1)
            If TOOL_CAPTION = "" Then
                TOOL_CAPTION = TOOL_KEY
            End If

            Dim OptionSetStarted As Boolean = False

            If ToolType <> "O" Then
                If OptionSetStarted Then
                    OptionSetStarted = False
                End If
            End If

            Select Case ToolType
                Case "B", ""
                    Dim tlb_btn As UltraWinToolbars.ButtonTool
                    If tkeys.Contains(TOOL_KEY) Then
                        tlb_btn = tlb.Tools(TOOL_KEY)
                    Else
                        tlb_btn = New UltraWinToolbars.ButtonTool(TOOL_KEY)
                        tlb.Tools.Add(tlb_btn)
                        tkeys.Add(TOOL_KEY)
                    End If

                Case "S"
                    Dim tlb_sbt As UltraWinToolbars.StateButtonTool
                    If tkeys.Contains(TOOL_KEY) Then
                        tlb_sbt = tlb.Tools(TOOL_KEY)
                    Else
                        tlb_sbt = New UltraWinToolbars.StateButtonTool(TOOL_KEY)
                        tlb_sbt.MenuDisplayStyle = UltraWinToolbars.StateButtonMenuDisplayStyle.DisplayCheckmark
                        tlb.Tools.Add(tlb_sbt)
                        tkeys.Add(TOOL_KEY)
                    End If

                Case "O"

                    Dim tlb_sbt As UltraWinToolbars.StateButtonTool
                    If tkeys.Contains(TOOL_KEY) Then
                        tlb_sbt = tlb.Tools(TOOL_KEY)
                    Else
                        tlb_sbt = New UltraWinToolbars.StateButtonTool(TOOL_KEY)
                        tlb_sbt.MenuDisplayStyle = UltraWinToolbars.StateButtonMenuDisplayStyle.DisplayCheckmark
                        tlb_sbt.OptionSetKey = OptionSetKey
                        tlb.Tools.Add(tlb_sbt)
                        tkeys.Add(TOOL_KEY)
                    End If

                    If Not OptionSetStarted Then
                        OptionSetStarted = True
                    End If

                Case "C"
                    Dim tlb_cpt As UltraWinToolbars.PopupColorPickerTool
                    If tkeys.Contains(TOOL_KEY) Then
                        tlb_cpt = tlb.Tools(TOOL_KEY)
                    Else
                        tlb_cpt = New UltraWinToolbars.PopupColorPickerTool(TOOL_KEY)
                        tlb.Tools.Add(tlb_cpt)
                        tkeys.Add(TOOL_KEY)
                    End If
                Case "M"
                    subMenuCount += 1
                    Dim passCount = 1
                    For x As Integer = T + 1 To ToolTypes.Count
                        If Mid(ToolTypes, x, 1) = "M" Then
                            passCount += 1
                        ElseIf Mid(ToolTypes, x, 1) = "X" Then
                            passCount -= 1
                        End If
                        If passCount = 0 Then
                            passCount = x
                            Exit For
                        End If
                    Next
                    Dim tempTools As List(Of String) = New List(Of String)
                    Dim toolCount As Integer = 0
                    For a As Integer = 1 To passCount
                        Select Case Mid(ToolTypes, a, 1)
                            Case "P", "X"

                            Case "M"
                                toolCount += 1
                                If a > T Then
                                    tempTools.Add(Tools(toolCount - 1))
                                End If
                            Case Else
                                toolCount += 1
                                If a >= T Then
                                    tempTools.Add(Tools(toolCount - 1))
                                End If
                        End Select
                    Next


                    Dim index = 1 'tlb_pop.Key.Substring(2, 1)
                    While tlb.Tools.Exists("sm" & index & tlb_pop.Key.Substring(3))
                        index += 1
                    End While
                    TOOL_KEY = "sm" & index & tlb_pop.Key.Substring(3)

                    Dim xctl As New Control(TOOL_KEY)
                    xctl.Name = TOOL_KEY
                    Load_Popup_Menu(tlb, xctl, Mid(ToolTypes, T + 1, passCount - T - 1), tempTools.ToArray())
                Case Else
                    Continue For
            End Select

            tlb_pop.Tools.AddTool(TOOL_KEY)
            tlb_pop.Tools(TOOL_KEY).InstanceProps.IsFirstInGroup = IsFirstInGroup
            IsFirstInGroup = False
            tlb_pop.Tools(TOOL_KEY).SharedProps.Caption = TOOL_CAPTION
            Dim IMAGE_FOLDER As String = ASCMAIN1.Folders("Images") & "16\"
            tlb_pop.Tools(TOOL_KEY).SharedProps.AppearancesSmall.Appearance.Image = ASCMAIN1.Get_Image(IMAGE_FOLDER, TOOL_KEY & ".PNG")
            ASCMAIN1.LoadKeysFromPopup(TOOL_KEY, TOOL_CAPTION, "")
        Next

        If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" Then

            TOOL_KEY = "Export to JSON"
            TOOL_CAPTION = TOOL_KEY
            Dim tlb_btn As UltraWinToolbars.ButtonTool
            If tkeys.Contains(TOOL_KEY) Then
                tlb_btn = tlb.Tools(TOOL_KEY)
            Else
                tlb_btn = New UltraWinToolbars.ButtonTool(TOOL_KEY)

                tlb.Tools.Add(tlb_btn)
                tkeys.Add(TOOL_KEY)
                tlb_pop.Tools.AddTool(TOOL_KEY)
                tlb_pop.Tools(TOOL_KEY).SharedProps.Caption = TOOL_CAPTION
            End If


        End If

        tlb.SetContextMenuUltra(ctl, ctl.Name)
    End Sub

    Sub DownloadFromAWebServer()

        ' basic connection to a web resource
        Dim myWebClient As New Net.WebClient
        Dim webstream As IO.Stream = myWebClient.OpenRead("http://www.absolution.com")
        Dim reader As New IO.StreamReader(webstream)
        Dim text As String = reader.ReadToEnd
        reader.Close()

        Console.WriteLine(text)
        Console.ReadLine()

        ' to upload values to a web site
        Dim loginForm As New Net.WebClient
        Dim loginFormElements As New System.Collections.Specialized.NameValueCollection
        loginFormElements.Add("UserID", "rboman")
        loginForm.UploadValues("http://www.sample/cmo/login.asp", loginFormElements)

        ' to login using credentials
        Dim restrictedPage As New Net.WebClient
        Dim myCache As New Net.CredentialCache
        Dim credential As New Net.NetworkCredential("myUser", "myPass")
        myCache.Add(New Uri("http://www.sample.com/restricted/"), "Basic", credential)
        restrictedPage.Credentials = myCache


    End Sub

    Sub ConnectToAServer()
        'Dim client As New Net.Sockets.TcpClient
        'Dim myStream As IO.Stream = client.GetStream

        ''client.LingerState.LingerTime = 2
        ''client.LingerState.Enabled = True

        'Dim myReader As New IO.StreamReader(myStream)
        'Dim data As String = myReader.ReadToEnd
        'myStream.Close()

        'Console.WriteLine(data)
        'Console.ReadLine()
    End Sub

    Sub Accept_Connection()
        'Dim server As New Net.Sockets.TcpListener(90)
        'server.Start()
        'Console.Write("Server is waiting ...")

        'Dim mySocket As Net.Sockets.Socket = server.AcceptSocket
        'Dim sendData As String = "Hi. Welcome to this server." & _
        'ControlChars.CrLf & Now.ToShortDateString & " " & _
        'Now.ToLongTimeString
        'Console.WriteLine("Sending: " & sendData)
        'Dim data As Byte() = System.Text.Encoding.ASCII.GetBytes(sendData)
        'mySocket.Send(data)

        'Console.WriteLine("Closing Connection ...")
        'mySocket.Close()
        'server.Stop()
        'Console.ReadLine()
    End Sub

    Sub Accept_Connection2()
        'Dim server As New Net.Sockets.TcpListener(90)
        'server.Start()
        'Dim mySocket As Net.Sockets.Socket = server.AcceptSocket
        'Dim buffer(100) As Byte, data As String = "", size As Integer = 0
        'Do Until InStr(Headers, ControlChars.CrLf) > 0
        '    size = mySocket.Receive(buffer)
        '    data += System.Text.Encoding.ASCII.GetString(buffer)
        'Loop
        'mySocket.Close()
        'server.Stop()
        'Console.WriteLine("Received a carriage return.")
        'Console.ReadLine()
    End Sub

    Public Shared Function Get_num_from_User(
     ByVal Label_Text As String,
     ByVal Form_Caption As String,
     Optional ByVal decimal_places As Long = 0,
     Optional ByVal maxValue As Decimal = 2147483647,
     Optional ByVal minValue As Decimal = -2147483648,
     Optional ByVal defaultValue As String = "") As Decimal

        Dim f As New ASFMSGBF
        Dim numEntry As Decimal
        If decimal_places = 0 Then
            numEntry = f.Get_numint_from_User(Label_Text, Form_Caption, maxValue, minValue, defaultValue)
        Else
            numEntry = f.Get_numdec_from_User(Label_Text, Form_Caption, maxValue, minValue, defaultValue)
        End If

        ASCMAIN1.response = f.user_option
        f.Dispose()

        Return numEntry
    End Function

    Public Shared Function Get_txt_from_User(
     ByVal Label_Text As String,
     ByVal Form_Caption As String,
     Optional ByVal password As Boolean = False,
     Optional ByVal maxLength As Long = 0,
     Optional ByVal defaultValue As String = "") As String

        Dim f As New ASFMSGBF
        Dim txtEntry As String = f.Get_txt_from_User(Label_Text, Form_Caption, password, maxLength, defaultValue)
        ASCMAIN1.response = f.user_option
        f.Dispose()

        Return txtEntry
    End Function

    Public Shared Function EnumToList(Of T)() As List(Of T)
        Dim enumType As Type = GetType(T)

        ' Can't use type constraints on value types, so have to do check like this
        If Not enumType.BaseType.Equals(GetType([Enum])) Then
            Throw New ArgumentException("T must be of type System.Enum")
        End If

        Dim enumValArray As Array = [Enum].GetValues(enumType)

        Dim enumValList As New List(Of T)(enumValArray.Length)

        For Each val As Integer In enumValArray
            enumValList.Add(DirectCast([Enum].Parse(enumType, val.ToString()), T))
            'enumValList.Add([Enum].Parse(enumType, val.ToString()))

        Next

        Return enumValList
    End Function

    ''' <summary>
    ''' Concatenates Files into One File
    ''' </summary>
    ''' <param name="strTarget"></param>
    ''' <param name="inputFiles"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ConCatFiles(ByVal strTarget As String, ByVal inputFiles As List(Of String)) As Boolean

        If inputFiles Is Nothing OrElse inputFiles.Count = 0 Then
            Return False
        End If

        Dim FSwrite As System.IO.FileStream
        Dim FSread As System.IO.FileStream
        Dim intBytesRead As Integer = 0
        Dim MyBuffer() As Byte
        Dim blLastSource As Boolean = False
        Dim numProccessed As Int16 = 0

        Try
            FSwrite = New System.IO.FileStream(strTarget, IO.FileMode.Create)

            For Each strFileName As String In inputFiles

                numProccessed += 1

                If strFileName.Length = 0 Then Continue For
                If Not My.Computer.FileSystem.FileExists(strFileName) Then Continue For

                blLastSource = (numProccessed = inputFiles.Count)

                ReDim MyBuffer(4095)  'Memory buffer for file transfer

                FSread = New System.IO.FileStream(strFileName, IO.FileMode.Open)

                'Write the contents of the source file to the target, one buffer-full at a time.
                intBytesRead = MyBuffer.Length
                While (intBytesRead = MyBuffer.Length)

                    'Read 1 buffer-full of bytes from source.
                    intBytesRead = FSread.Read(MyBuffer, 0, MyBuffer.Length)

                    ' Check last byte read for an EOF = $1A = ASCII 26.  Ignore if found,
                    ' unless this is the last source file.
                    If intBytesRead > 0 And MyBuffer(intBytesRead - 1) = 26 _
                       And blLastSource = False Then
                        intBytesRead -= 1
                    End If

                    'If we read any bytes, write them to the target.
                    If intBytesRead > 0 Then
                        FSwrite.Write(MyBuffer, 0, intBytesRead)
                    End If

                End While

                'Done with this source file.  Close it.
                FSread.Close()

            Next

            FSwrite.Close()
            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Shared Function ordinal(ByVal intNumber As Integer) As String
        If CType(intNumber, String).Length > 2 Then
            Dim intEndNum As Integer = CType(CType(intNumber, String).Substring(CType(intNumber, String).Length - 2, 2), Integer)
            If intEndNum >= 11 And intEndNum <= 13 Then
                Select Case intEndNum
                    Case 11, 12, 13
                        Return CStr(intNumber) & "th"
                End Select
            End If
        End If
        If intNumber >= 21 Then
            ' Handles 21st, 22nd, 23rd, et al
            Select Case CType(intNumber.ToString.Substring(intNumber.ToString.Length - 1, 1), Integer)
                Case 1
                    Return CStr(intNumber) & "st"
                Case 2
                    Return CStr(intNumber) & "nd"
                Case 3
                    Return CStr(intNumber) & "rd"
                Case 0, 4 To 9
                    Return CStr(intNumber) & "th"
            End Select
        Else
            ' Handles 1st to 20th
            Select Case intNumber
                Case 1
                    Return CStr(intNumber) & "st"
                Case 2
                    Return CStr(intNumber) & "nd"
                Case 3
                    Return CStr(intNumber) & "rd"
                Case 4 To 20
                    Return CStr(intNumber) & "th"
            End Select
        End If
        ' If here, no match was found. Should be impossible. In fact, if we're here the day of the month is a non-number. Best not return an ordinal at all.
        Return CStr(intNumber) & ""
    End Function

    Public Shared Sub Disable_Form(ByVal ABSF As ASFBASE1)
        ABSF.ScreenMode = False
        'ABSF.Enabled = False
        ABSF.UltraExplorerBar1.Enabled = False
        ABSF.ASFBASE1_Fill_Panel.Enabled = False

        'ABSF.Rollback() ' NEED TO FIND OUT WHETHER TRANSACTION IS PENDING
        ASCMAIN1.MultiTask_Release(ASCMAIN1.SESSION_NO, ABSF.SELECTION_NO)
        MsgBox("Form has been Disabled - please email a screenshot",
               MsgBoxStyle.OkOnly,
               "Please Call ABS prior to Closing this Form")
    End Sub

    Public Shared Sub Set_Form_Caption_on_Tab(ByVal frmASFBASE0 As ASFBASE0)
        ASFMAIN1.UltraTabbedMdiManager1.ActiveTab.Text = frmASFBASE0.Text
    End Sub

    Public Shared Sub Initalize_FCB(ByVal COLUMN_NAME As String, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs _
                                    , Optional ByVal ON_VALUE As String = "1")
        Dim FILENAME As String = ""

        If e.Row.Cells(COLUMN_NAME).Value & "" = ON_VALUE Then
            FILENAME = "Selected"
        Else
            FILENAME = "Blank Selection"
        End If
        If FILENAME <> "" Then
            e.Row.Cells(COLUMN_NAME).ButtonAppearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", FILENAME)
        End If

    End Sub

    Public Shared Sub Set_FCB(ByVal COLUMN_NAME As String, ByVal grd As UltraWinGrid.UltraGrid _
                              , Optional ByVal ON_VALUE As String = "1", Optional ByVal OFF_VALUE As String = "0")
        With grd.ActiveRow
            If .Cells(COLUMN_NAME).Value & "" <> ON_VALUE Then
                .Cells(COLUMN_NAME).Value = ON_VALUE
            Else
                .Cells(COLUMN_NAME).Value = OFF_VALUE
            End If
            .Update()
        End With
    End Sub

    Public Shared Sub Design_FCB(ByVal COLUMN_NAME As String, ByVal grd As UltraWinGrid.UltraGrid)
        With grd.DisplayLayout.Bands(0)
            .Columns(COLUMN_NAME).ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
            .Columns(COLUMN_NAME).CellAppearance.ForeColor = Drawing.Color.Transparent
            .Columns(COLUMN_NAME).CellAppearance.ForegroundAlpha = Alpha.Transparent
            .Columns(COLUMN_NAME).CellAppearance.TextHAlign = HAlign.Center 'Not necessary but recommended
            .Columns(COLUMN_NAME).CellButtonAppearance.ImageHAlign = HAlign.Center 'Not necessary but recommended
            .Columns(COLUMN_NAME).CellButtonAppearance.ImageVAlign = VAlign.Middle 'Not necessary but recommended
            .Columns(COLUMN_NAME).Style = UltraWinGrid.ColumnStyle.Button
        End With
    End Sub

    Public Shared Function Execute_Debug_Code() As Boolean
        Dim ok_to_execute As Boolean = False
        If developerModeOptions.RunDebugCode Then
            If developerModeOptions.RunDebugCodePrompt Then
                Dim dcp As String = "You are about to execute code a developer has designated as debug only." _
                & vbCrLf & "Please be aware of what you are about to do." & vbCr & vbCr & "Continue Anyway?"
                If vbNo = MsgBox(dcp, vbQuestion + vbYesNo, "Run Debug Code Prompt") Then
                    ok_to_execute = False
                Else
                    ok_to_execute = True
                End If
            Else
                ok_to_execute = True
            End If
        End If
        Return ok_to_execute
    End Function

    ' use Excel_Cell or Excel_Cell0
    'Public Shared Function XC( _
    '    ByVal C As Int16, _
    '    Optional ByVal R As Int16 = 0, _
    '    Optional ByVal absolute As Boolean = False) As String

    '    Dim COL As String = ""
    '    If C >= 1 Then
    '        Dim B As Int16 = (C - 1) Mod 26 + 1
    '        Dim A As Int16 = (C - B) / 26
    '        COL = Chr(Asc("A") + B - 1)
    '        If A > 0 Then
    '            COL = Chr(Asc("A") + A - 1) & COL
    '        End If
    '        If absolute Then
    '            COL = "$" & COL
    '        End If

    '        If R = 0 Then
    '            COL = COL & ":" & COL
    '        ElseIf R > 0 Then
    '            COL = COL & IIf(absolute, "$", "") & CStr(R)
    '        End If
    '    End If

    '    Return COL
    'End Function

    Public Shared Function Get_Hash(strToHash As String) As String

        Dim sha1Obj As New Security.Cryptography.SHA1CryptoServiceProvider
        Dim bytesToHash() As Byte = System.Text.Encoding.ASCII.GetBytes(strToHash)

        bytesToHash = sha1Obj.ComputeHash(bytesToHash)

        Dim strResult As String = ""

        For Each b As Byte In bytesToHash
            strResult += b.ToString("x2")
        Next

        Return strResult
    End Function

    Public Shared Function EncryptAES(strToHash As String) As String
        Dim Ezcrypt1 As New nsoftware.IPWorksEncrypt.Ezcrypt()
        Ezcrypt1.RuntimeLicense = TACMAIN1.nSoftwareEncryptionkey
        Ezcrypt1.Reset()
        Ezcrypt1.Algorithm = nsoftware.IPWorksEncrypt.EzcryptAlgorithms.ezAES
        Ezcrypt1.UseHex = True
        Ezcrypt1.InputMessage = strToHash
        Ezcrypt1.KeyPassword = "0ff1c3"
        Ezcrypt1.Encrypt()
        Dim encrypted As String = Ezcrypt1.OutputMessage

        Return encrypted
    End Function

    Public Shared Function DecryptAES(strToHash As String) As String
        Dim Ezcrypt1 As New nsoftware.IPWorksEncrypt.Ezcrypt()
        Ezcrypt1.RuntimeLicense = TACMAIN1.nSoftwareEncryptionkey
        Ezcrypt1.Reset()
        Ezcrypt1.Algorithm = nsoftware.IPWorksEncrypt.EzcryptAlgorithms.ezAES
        Ezcrypt1.UseHex = True
        Ezcrypt1.InputMessage = strToHash
        Ezcrypt1.KeyPassword = "0ff1c3"
        Ezcrypt1.Decrypt()
        Dim decrypted As String = Ezcrypt1.OutputMessage

        Return decrypted
    End Function


    Public Shared Function Get_User_JWT() As String

        If ASCMAIN1.USER_JWT <> "" AndAlso DateTime.Now < ASCMAIN1.USER_JWT_EXPIRES Then
        Else
            Dim client As New HttpClient()
            client.BaseAddress = New Uri(Get_API_Endpoint(ASCMAIN1.Running_in_VS))
            Dim apiMethod As String = "login"
            Dim API_CONTROLLER As String = "AS/" & apiMethod

            client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
            client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/x-www-form-urlencoded"))

            Dim frmtr As MediaTypeFormatter = New JsonMediaTypeFormatter()
            Dim UC As New userCreds
            UC.USER_ID = ASCMAIN1.USER_ID
            UC.USER_PASSWORD = ASCMAIN1.USER_PASSWORD
            Dim content As HttpContent = New ObjectContent(Of userCreds)(UC, frmtr)

            Dim resp As HttpResponseMessage = Nothing
            Dim resp_err As String = ""
            Try
                resp = client.PostAsync(API_CONTROLLER, content).Result
                If resp.StatusCode = Net.HttpStatusCode.OK Then
                    Dim apiResponseString As String = ""
                    Dim responseObject As Object = Nothing
                    responseObject = resp.Content.ReadAsAsync(Of Object)().Result
                    ASCMAIN1.USER_JWT = responseObject("BearerToken").ToString
                    ASCMAIN1.USER_JWT_EXPIRES = DateTime.Now.AddMinutes(30)
                End If
            Catch ex As Exception
                resp_err = ex.InnerException.InnerException.Message
            End Try
        End If


        Return ASCMAIN1.USER_JWT

    End Function

    Public Shared Function Get_API_Endpoint(Optional useLocalHost As Boolean = False) As String
        If useLocalHost Then
            Dim localApiEndpoint As String = "http://localhost:1642/api/"
            If ASCMAIN1.API_ENDPOINT = "" Or ASCMAIN1.API_ENDPOINT <> localApiEndpoint Then
                MsgBox("Using local API endpoint: " & localApiEndpoint, vbOKOnly, "Developer Alert")
            End If
            ASCMAIN1.API_ENDPOINT = localApiEndpoint
        Else
            If ASCMAIN1.API_ENDPOINT = "" Then
                Dim rowWBTPARM1 As DataRow = ASFBASE1.LookUp("WBTPARM1", "Z")
                ASCMAIN1.API_ENDPOINT = rowWBTPARM1.Item("WB_PARM_API_URL")
            End If
        End If
        Return ASCMAIN1.API_ENDPOINT
    End Function


#Region "Serial and Comm Port Connections"

    Public Shared Property LaserPrinterName() As String
        Get
            Return (ASFMAIN1.laserPrinterName)
        End Get
        Set(value As String)
            ASFMAIN1.laserPrinterName = value
        End Set
    End Property

    Public Shared Property LaserPrinterIpAddress() As String
        Get
            Return (ASFMAIN1.laserPrinterIP)
        End Get
        Set(ByVal value As String)
            ASFMAIN1.laserPrinterIP = value
        End Set
    End Property

    Public Shared Property LabelPrinterSerialPort() As System.IO.Ports.SerialPort
        Get
            Return ASFMAIN1.labelPrinterSerialPort
        End Get
        Set(ByVal value As System.IO.Ports.SerialPort)
            ASFMAIN1.labelPrinterSerialPort = value
        End Set
    End Property

    'Public Shared Property ScaleSerialPort() As System.IO.Ports.SerialPort
    '    Get
    '        Return ASFMAIN1.scaleSerialPort
    '    End Get
    '    Set(ByVal value As System.IO.Ports.SerialPort)
    '        ASFMAIN1.scaleSerialPort = value
    '    End Set
    'End Property

    'Public Shared Property ScaleWeightDelegate() As ASFMAIN1.ScaleDelegate
    '    Get
    '        Return ASFMAIN1.scaleWeightDelegate
    '    End Get
    '    Set(ByVal value As ASFMAIN1.ScaleDelegate)
    '        ASFMAIN1.scaleWeightDelegate = value
    '    End Set
    'End Property

#End Region

    Private Shared Async Sub absTimer_TimerElapsed(ByVal sender As Object, ByVal e As ElapsedEventArgs) Handles absTimer.Elapsed
        Console.WriteLine(e.SignalTime)

        Dim interval_mins As Integer = 3000 ' 3000000 ms
        If absTimer.Interval <> interval_mins * 1000 Then
            absTimer.Interval = interval_mins * 1000
        End If

        Try

            Dim appID As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_APP_ID") & ""
            Dim tenantID As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_TENANT_ID") & ""
            Dim clientSecret As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_CLIENT_SECRET") & ""
            If appID <> "" And tenantID <> "" And clientSecret <> "" Then
                clientSecret = ASCMAIN1.DecryptAES(clientSecret)
            Else
                appID = ConfigurationManager.AppSettings("appId")
                clientSecret = ConfigurationManager.AppSettings("clientSecret")
                tenantID = ConfigurationManager.AppSettings("tenantId")
            End If
            '.Create(ConfigurationManager.AppSettings("appId")) _
            '.WithClientSecret(ConfigurationManager.AppSettings("clientSecret")) _
            '.WithTenantId(ConfigurationManager.AppSettings("tenantId")) _

            Dim cca As IConfidentialClientApplication = ConfidentialClientApplicationBuilder _
            .Create(appID) _
                .WithClientSecret(clientSecret) _
                .WithTenantId(tenantID) _
                .Build()
            Dim ewsScopes As String() = New String() {"https://outlook.office365.com/.default"}
            Dim authResult As AuthenticationResult = Await cca.AcquireTokenForClient(ewsScopes).ExecuteAsync()
            ASCMAIN1.authResult = authResult
            ASCMAIN1.authResult_counter += 1
            ASCMAIN1.authResult_timestamp = Now
            Console.WriteLine(authResult.ToString)
            Console.WriteLine(ASCMAIN1.authResult_counter)
            Console.WriteLine(ASCMAIN1.authResult_timestamp)

        Catch ex As Exception
            If ASCMAIN1.Running_in_VS Then Stop
            If ASCMAIN1.USER_ID = "wjz" Then MsgBox("Issue with Auith Token:" & ex.Message)
        End Try

    End Sub

    Class userCreds
        Public USER_ID As String
        Public USER_PASSWORD As String
    End Class

End Class