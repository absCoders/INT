Imports System.Windows.Forms
Imports System.Math
Imports Oracle.ManagedDataAccess.Client
Imports System.IO


Public Class ASCMAIN1
    Implements IDisposable

    ' Application
    Public FORM_NAME As String
    Public SET_ID As String
    Public ActiveForm As ASCBASE0 '  Windows.Forms.Form
    Public MainForm As Windows.Forms.Form
    Public MainForm_pgd As Windows.Forms.PropertyGrid
    Public ABS_FORMS() As System.Windows.Forms.Form
    Public Folders As New Dictionary(Of String, String)

    ' General Purpose
    Public sql As String

    Public Shared rowASTPARM1 As DataRow
    Public NowTSD As TimeSpan
    Public grdRows As New Collection
    Public response As Integer
    Public Message As String

    ' User Logged In
    Public USER_ID As String
    Public USER_NAME As String
    Public USER_PASSWORD As String
    Public USER_CODES As String
    Public USER_SECURITY_CODEs As String
    Public USER_MENU_ITEM_OBJECT As String
    Public SESSION_NO As String
    Public USER_EMAIL As String
    Public WTS_SESSION_ID As Int32
    Public MENU_ITEM_OBJECTs As List(Of String)
    Public client_schema As String = ""
    ' DB & Schema
    'Public DBS As String
    'Public ABS_Data As New ABS_Data.ABS_DataAdapter

    Public DBS_SERVER As String
    Public DBS_COMPANY As String
    Public DBS_PASSWORD As String
    Public DBS_CONCAT As String
    Public DBS_PARAMETER As String
    Public DBS_NOROWS As String
    Public DBS_OBJPFX As String
    Public DBS_OBJSFX As String
    Public DBS_SESSION_ID As Long
    Public DBS_IP_ADDRESS As String
    Public DBS_SERVER_NAME As String

    Public ASCDATA1 As ASCDATA1

    Public Enum DBS_TYPE_types
        Oracle
        SQLServer
    End Enum

    Public Enum ExecuteSQL_types
        Build_and_Execute
        Build_Only
        Execute_Only
    End Enum

    Public DBS_TYPE As DBS_TYPE_types
    Public FilledSchemas As New Dictionary(Of String, DataTable)

    ' DB Connection
    Public oraCon As New OracleConnection
    Public oraCmd As New OracleCommand
    Public oraSP As New OracleCommand    ' Used to execute Stored Procedures
    Public T As OracleTransaction
    Public oraAda As New OracleDataAdapter

    ' Universe
    Public COMPUTER_NAME As String
    Public ABSWEB As Boolean
    Public APP_PATH As String
    Public MAIN_FOLDER As String
    Public ASSEMBLY_NAME As String
    Public VERSION_NO As String
    Public Running_in_VS As Boolean

    Public GridDoubleClickAllowed As Boolean = True


    Public Shared SOLUTION As String
    Public Shared CLIENT As String
    Public Shared EncryptionKey As String = String.Empty

    ' Environment
    Public tblASTFFMT1 As New DataTable
    Public tblASTSECK1 As New DataTable
    Public tblASTSQLX1 As DataTable
    Public dstASTVIEWS As New DataSet

    Public TACMAIN1 As TACMAIN1
    Public ABS_Assemblies As New Dictionary(Of String, System.Reflection.Assembly)
    Public tblASTMENU1 As New DataTable

    ' Period Context
    Public CYW As String     ' Current Year and Week, YYYYWW
    Public CYP As String     ' Current Year and Period, YYYYPP
    Public CYM As String     ' Calendar Year and Month of CYP, YYYYMM
    Public PCO As Integer    ' Period Calendar Offset to Regular Calendar, 1 = Year ends in Jan, 2 = Year ends in Feb, 0 = Year ends in Dec
    Public WCO As Integer    ' Week Calendar Offset to Regular Calendar
    Public EOM As String     ' 1 = Period End is in Process
    Public Timer As Date

    Public JOB_STREAM_CODE As String = String.Empty
    Public JOB_STREAM_FORM_NAME As String = String.Empty
    Public JOB_STREAM_XNO As String = String.Empty
    Public JOB_STREAM_LNO As Int16 = 0

    Public Sub Set_DBS_Dependent_Strings()
        If oraCon.GetType.ToString() = "System.Data.SqlClient.SqlConnection" Then
            DBS_TYPE = DBS_TYPE_types.SQLServer
            DBS_CONCAT = "+"
            DBS_PARAMETER = "@"
            DBS_NOROWS = "1 <> 1"
            DBS_OBJPFX = "["
            DBS_OBJSFX = "]"
        Else
            DBS_TYPE = DBS_TYPE_types.Oracle
            DBS_CONCAT = "||"
            DBS_PARAMETER = ":"
            DBS_NOROWS = "ROWNUM < 1"
            DBS_OBJPFX = Chr(34)
            DBS_OBJSFX = Chr(34)
        End If
    End Sub

    Public Sub Center(ByVal F As Form)
        F.SetBounds((System.Windows.Forms.Screen.GetBounds(F).Width / 2) - (F.Width / 2), _
    (System.Windows.Forms.Screen.GetBounds(F).Height / 2) - (F.Height / 2), _
    F.Width, F.Height, System.Windows.Forms.BoundsSpecified.Location)
    End Sub

    Public Function Make_Caption(ByVal COLUMN_NAME As String) As String
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
    Public Function Next_Control_No( _
    ByVal CTL_NO_TYPE As String, _
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

                    Dim sqlCTL_NO_LENGTH As String = _
                        "Select DATA_LENGTH from USER_TAB_COLUMNS " _
                        & " where TABLE_NAME = '" & Split(CTL_NO_TYPE, ".")(0) & "' and COLUMN_NAME = '" _
                        & Split(CTL_NO_TYPE, ".")(1) & "'"

                    If DBS_TYPE = DBS_TYPE_types.SQLServer Then
                        sqlCTL_NO_LENGTH = "SELECT CHARACTER_MAXIMUM_LENGTH " _
                        & " FROM INFORMATION_SCHEMA.COLUMNS " _
                        & " where TABLE_CATALOG = '" & DBS_COMPANY & "' " _
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
            ' Dim ff = .UpdateCommand.CommandText
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
                        , New Object() {CTL_NO_TYPE, CTL_NO_LAST, Next_Control_No, How_Many, USER_ID})
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
    Public Function DateDiff_Weekday( _
    ByVal base_date As Object, _
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
    Public Sub DateDiff_Weekday_Adjust_for_Weekend( _
    ByRef working_date As Date, _
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
    Public Function Distinct_Values( _
    ByVal TABLE_NAME As String, _
    ByRef dt_source As DataTable, _
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
    Public Function Distinct_Values( _
    ByVal TABLE_NAME As String, _
    ByVal RowFilter As String, _
    ByRef dt_source As DataTable, _
    ByVal ParamArray COLUMN_NAMEs() As String) As DataTable

        Dim dt As DataTable = New DataView(dt_source, RowFilter, "", DataViewRowState.CurrentRows).ToTable(True, COLUMN_NAMEs)
        If TABLE_NAME <> "" Then
            dt.TableName = TABLE_NAME
        End If

        Return dt

    End Function

    Public Function Temp_Table( _
    Optional ByVal sql As String = "", _
    Optional ByVal FORM_NAME As String = "", _
    Optional ByVal XNO As String = "", _
    Optional ByVal custom_parameters As String = "", _
    Optional ByVal PARMs() As String = Nothing) As String

        If sql = "" Then
            sql = sql
        End If

        If ActiveForm IsNot Nothing Then
            Try
                If FORM_NAME = "" Then FORM_NAME = ActiveForm.MENU_ITEM_OBJECT
                If XNO = "" Then XNO = ActiveForm.XNO
            Catch ex As Exception

            End Try
        End If

        Dim TABLE_NAME As String = "ASW" & Next_Control_No("ASTTEMP1")
        Dim S As String = "SYSDATE"
        If DBS_TYPE = DBS_TYPE_types.SQLServer Then
            S = "GETDATE()"
        End If
        Dim sqlASTTEMP1 As String = "Insert into ASTTEMP1 Values ('" & TABLE_NAME & "','" & SESSION_NO & "', " & S & ",'" & USER_ID & "','" & FORM_NAME & "','" & XNO & "'," & CStr(DBS_SESSION_ID) & ")"
        ASCDATA1.ExecuteSQL(sqlASTTEMP1)

        If DBS_TYPE = DBS_TYPE_types.SQLServer Then
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

    Public Sub Temp_Table_Cleanup( _
    Optional ByVal General_Cleanup As Boolean = True)

        If General_Cleanup Then
            ' this causes many slowdowns exiting ABSolution - handle it in backup procedure
            Exit Sub
        End If

        If rowASTPARM1 Is Nothing OrElse rowASTPARM1.Item("AS_PARM_NO_DROP_TEMP_TABLES") & "" = "1" Then
            Exit Sub
        End If

        Dim sql As String = ""
        sql = "Select * from ASTTEMP1 where (USER_ID = '" & USER_ID & "'"
        If General_Cleanup Then
            sql = sql & " and SESSION_ID not in (Select AUDSID from V$SESSION)"
        Else
            sql = sql & " and SESSION_NO = '" & SESSION_NO & "'"
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
                    If Running_in_VS Then
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

        sql = "Select TABLE_NAME from USER_TABLES where TABLE_NAME LIKE 'ASW%' and LAST_ANALYZED < SYSDATE -2" _
            & " minus Select TABLE_NAME from ASTTEMP1"
        ASCDATA1.ExecuteSQL("Insert into ASTTEMP1 (TABLE_NAME,SESSION_NO,DATE_CREATED,USER_ID,FORM_NAME,XNO,SESSION_ID) Select TABLE_NAME, Null SESSION_NO, SYSDATE DATE_CREATED, '" & USER_ID & "' USER_ID,'" & USER_ID & "' FORM_NAME, Null XNO, Null SESSION_ID from (" & sql & ")")

    End Sub

    ''' <summary>
    ''' Returns a string which is ready to be used in a SQL statement IN expression.
    ''' Passing in a CODEs value of AABBCC with a CODE_length of 2 would return 'AA','BB','CC'.
    ''' </summary>
    ''' <param name="CODEs"></param>
    ''' <param name="CODE_length"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Select_List(ByVal CODEs As String, ByVal CODE_length As Integer) As String
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
    Public Function SQL_Add_WHERE(ByVal WHEREless_Clause As String) As String
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

    Public Function Make_Plural(ByVal singular_noun As String) As String
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

    Public Sub Get_Current_YP()
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
            sql = "Select YYYYWW from GLTPARM3 where WEEK_END_DATE >= '" & Format$(Now + NowTSD, Get_Date_Mask) & "'"
            sql = sql & " and WEEK_END_DATE < '" & Format$(DateAdd(DateInterval.Day, 7, (Now + NowTSD)), Get_Date_Mask) & "'"
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

    Public Function Get_YYYYMM( _
    ByVal YYYYPP As String, _
    Optional ByVal PERIOD_OFFSET As Integer = 0)
        If YYYYPP = "" Then
            Return ""
        Else
            Return Format$(DateAdd("M", PCO + PERIOD_OFFSET, Mid$(YYYYPP, 5, 2) & "/01/" & Mid$(YYYYPP, 3, 2)), "yyyyMM")
        End If
    End Function

    Public Function Get_Dates(ByVal YP1 As String) As Date()

        Dim YP0 As String = Period_Calc(YP1, -1)
        sql = "Select OPS_YYYYPP, PRD_END_DATE from GLTPARM2 " _
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

    Public Function Get_Date_Mask() As String
        Get_Date_Mask = ""
        Get_Date_Mask = "dd-MMM-yyyy"
        '        Get_Date_Mask = "MM/dd/yyyy"
    End Function

    Public Function Get_Image( _
        ByVal IMAGE_FOLDER As String, _
        ByVal IMAGE_FILE As String, _
        Optional return_byte_array As Boolean = False, _
        Optional ByRef IMAGE_FOLDER_USED As String = "", _
        Optional ByRef IMAGE_FILE_USED As String = "", _
        Optional ByRef byte_array() As Byte = Nothing) As System.Drawing.Bitmap

        Dim img As System.Drawing.Bitmap = Nothing

        IMAGE_FOLDER_USED = IMAGE_FOLDER
        IMAGE_FILE_USED = IMAGE_FILE

        Dim image_file_found As Boolean = True

        If IMAGE_FILE = "\.jpg" Then
            image_file_found = False
            Return Nothing ' My.Resources.ABS
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
                img = Nothing ' My.Resources.ABS
            End If

        Catch ex As Exception
            image_file_found = False
            img = Nothing ' My.Resources.ABS
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

    Public Function Get_Filename( _
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
    Public Function Flattened_List( _
    ByVal COLUMN_KEY As String, _
    ByVal COLUMN_DATA As String, _
    ByVal TABLE_NAME As String, _
    Optional ByVal DELIMITER As String = ",", _
    Optional ByVal WHERE_CLAUSE As String = "") As String

        Dim SQL As String
        SQL = "Select " & COLUMN_KEY & ","
        SQL = SQL & "  ltrim(sys_connect_by_path(" & COLUMN_DATA & ",'" & DELIMITER & "'),'" & DELIMITER & "') " & COLUMN_DATA & "S"
        SQL = SQL & "      from"
        SQL = SQL & "    (select " & COLUMN_KEY & ", " & COLUMN_DATA & ","
        SQL = SQL & "           row_number() over(partition by " & COLUMN_KEY & " order by " & COLUMN_DATA & ") rn,"
        SQL = SQL & "          row_number() over(partition by " & COLUMN_KEY & " order by " & COLUMN_DATA & " desc)"
        SQL = SQL & "  rn_desc"
        SQL = SQL & " FROM " & TABLE_NAME
        If WHERE_CLAUSE <> "" Then
            SQL = SQL & SQL_Add_WHERE(WHERE_CLAUSE)
        End If
        SQL = SQL & ")"
        SQL = SQL & "     Where rn_desc = 1"
        SQL = SQL & "     start with rn = 1"
        SQL = SQL & "    connect by prior " & COLUMN_KEY & " = " & COLUMN_KEY & ""
        SQL = SQL & "  and prior rn = rn-1"

        Flattened_List = SQL
    End Function

    Public Sub Load_Views()

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
    Public Sub Multi_Task_Cleanup()

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
    Public Function Logical_Lock( _
    ByVal ENTITY_TYPE As String, _
    ByVal ENTITY As String, _
    Optional ByVal menu_level As Boolean = False, _
    Optional ByVal show_message As Boolean = True, _
    Optional ByVal reverse_all_previous_if_unsuccessful As Boolean = True, _
    Optional ByVal MT_LEVEL As Integer = 0) As Boolean
        Dim EMsg As String = ""
        If Not MultiTask(ENTITY_TYPE, ENTITY, "L", 0, EMsg, , , _
                         menu_level, , reverse_all_previous_if_unsuccessful, MT_LEVEL) Then
            'If show_message Then MsgBox(EMsg, vbOKOnly, "Cannot Proceed")
            Logical_Lock = False
        Else
            Logical_Lock = True
        End If
    End Function

    Public Function Logical_Open( _
    ByVal ENTITY_TYPE As String, _
    ByVal ENTITY As String, _
    Optional ByVal menu_level As Boolean = False, _
    Optional ByVal show_message As Boolean = True, _
    Optional ByVal reverse_all_previous_if_unsuccessful As Boolean = True, _
    Optional ByVal MT_LEVEL As Integer = 0) As Boolean
        Dim EMsg As String = ""
        If Not MultiTask(ENTITY_TYPE, ENTITY, "O", 1, EMsg, , , _
                         menu_level, , reverse_all_previous_if_unsuccessful, MT_LEVEL) Then
            'MsgBox(EMsg, vbOKOnly, "Cannot Proceed")
            Logical_Open = False
        Else
            Logical_Open = True
        End If
    End Function

    Public Function MultiTask( _
    ByVal ENTITY_TYPE As String, _
    ByVal ENTITY As String, _
    ByVal MT_ACTION As String, _
    ByVal OPEN_COUNT As Integer, _
    Optional ByRef EMsg As String = "", _
    Optional ByVal SESSION_NO As String = "", _
    Optional ByVal SELECTION_NO As Integer = 0, _
    Optional ByVal menu_level As Boolean = False, _
    Optional ByVal MT_MENU As String = "0", _
    Optional ByVal reverse_all_previous_if_unsuccessful As Boolean = True, _
    Optional ByVal MT_LEVEL As Integer = 0) As Boolean

        If SESSION_NO = "" Then
            SESSION_NO = Me.SESSION_NO
        End If
        If menu_level Then
            SELECTION_NO = UBound(ABS_FORMS) + 1
        Else
            If SELECTION_NO = 0 Then
                SELECTION_NO = ActiveForm.SELECTION_NO
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

        With ASCDATA1.GetDataAdapter(tblASTMTSK1, "ASTMTSK1", sql, True, , , 0, , , "VV", _
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
                            rowASTMTSK1.Item("LOCK_BY") = USER_ID
                            rowASTMTSK1.Item("OPEN_COUNT") = OPEN_COUNT
                            tblASTMTSK1.Rows.Add(rowASTMTSK1)
                            .Update(tblASTMTSK1)
                            MultiTask = MultiTask_Detail(ENTITY_TYPE, ENTITY, OPEN_COUNT, SESSION_NO, SELECTION_NO, menu_level, MT_MENU, MT_LEVEL)
                        Else
                            If tblASTMTSK1.Rows.Count = 1 AndAlso tblASTMTSK1.Rows(0).Item("LOCK_BY") & "" = USER_ID AndAlso tblASTMTSK1.Rows(0).Item("MT_ACTION") & "" = "L" Then
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
                                rowASTMTSK1.Item("LOCK_BY") = USER_ID
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

    Public Function MultiTask_Detail( _
    ByVal ENTITY_TYPE As String, _
    ByVal ENTITY As String, _
    ByVal OPEN_COUNT As Integer, _
    Optional ByVal SESSION_NO As String = "", _
    Optional ByVal SELECTION_NO As Integer = 0, _
    Optional ByVal menu_check As Boolean = False, _
    Optional ByVal MT_MENU As String = "0", _
    Optional ByVal MT_LEVEL As Integer = 0) As Boolean

        If SESSION_NO = "" Then
            SESSION_NO = Me.SESSION_NO
        End If
        If SELECTION_NO = 0 Then
            SELECTION_NO = ActiveForm.SELECTION_NO
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

                sql = "Select * from ASTMTSK2 where ENTITY_TYPE = :PARM1 and ENTITY = :PARM2 and SESSION_NO = :PARM3 and SELECTION_NO = :PARM4"
                With ASCDATA1.GetDataAdapter(tblASTMTSK2, "ASTMTSK2", sql, True, -1, True, 0, , , "VVVN", New Object() {ENTITY_TYPE, ENTITY, SESSION_NO, SELECTION_NO})
                    ' With ASCDATA1.GetDataAdapter(tblASTMTSK2, "ASTMTSK2", "**", True, -1, True, 0, , , "VVVN", New Object() {ENTITY_TYPE, ENTITY, SESSION_NO, SELECTION_NO})


                    If tblASTMTSK2.Rows.Count = 0 Then
                        Dim rowASTMTSK2 As DataRow = tblASTMTSK2.NewRow
                        rowASTMTSK2.Item("ENTITY_TYPE") = ENTITY_TYPE
                        rowASTMTSK2.Item("ENTITY") = ENTITY
                        rowASTMTSK2.Item("SESSION_NO") = SESSION_NO
                        rowASTMTSK2.Item("SELECTION_NO") = SELECTION_NO
                        rowASTMTSK2.Item("INIT_OPER") = USER_ID
                        rowASTMTSK2.Item("INIT_DATE") = Now + NowTSD
                        'If menu_check Then
                        '    rowASTMTSK2.Item("MENU_ITEM_TYPE") = ASFMAIN1.MENU_ITEM_TYPE
                        '    rowASTMTSK2.Item("MENU_ITEM_OBJECT") = ASFMAIN1.MENU_ITEM_OBJECT
                        'Else
                        '    rowASTMTSK2.Item("MENU_ITEM_TYPE") = ActiveForm.MENU_ITEM_TYPE
                        '    rowASTMTSK2.Item("MENU_ITEM_OBJECT") = ActiveForm.MENU_ITEM_OBJECT
                        'End If
                        rowASTMTSK2.Item("MENU_ITEM_TYPE") = ActiveForm.MENU_ITEM_TYPE ' "C"
                        rowASTMTSK2.Item("MENU_ITEM_OBJECT") = ActiveForm.MENU_ITEM_OBJECT

                        rowASTMTSK2.Item("MT_MENU") = MT_MENU
                        rowASTMTSK2.Item("MT_LEVEL") = MT_LEVEL
                        rowASTMTSK2.Item("SESSION_ID") = DBS_SESSION_ID

                        tblASTMTSK2.Rows.Add(rowASTMTSK2)
                        .Update(tblASTMTSK2)
                    End If
                End With
            End If
            MultiTask_Detail = True

        Catch ex As Exception
            Throw New Exception(String.Format("Exception in MultiTask_Detail: {0}", ex.Message))
        End Try

    End Function

    Public Function MultiTask_Get_Users( _
    ByVal ENTITY_TYPE As String, _
    ByVal ENTITY As String, _
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
                'If MT_ACTION <> "L" Then
                sql = "Select * from ASTMENU1 " _
                 & " where (MENU_ITEM_TYPE = :PARM1" _
                 & "   and  MENU_ITEM_OBJECT = :PARM2)"
                Dim rowASTMENU1 As DataRow = ASCDATA1.GetDataRow(sql, "VV", New String() {row.Item("MENU_ITEM_TYPE"), row.Item("MENU_ITEM_OBJECT")})
                If rowASTMENU1 IsNot Nothing Then
                    Users = Users & vbCrLf & vbCrLf & "System Codes: " & vbCr & ENTITY_TYPE & " : " & ENTITY
                    Users = Users & " (" & rowASTMENU1.Item("MENU_ITEM_DESC") & ")"
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
    Public Sub MultiTask_Release( _
    Optional ByVal SESSION_NO As String = "", _
    Optional ByVal SELECTION_NO As Integer = 0, _
    Optional ByVal MT_LEVEL As Integer = 0)

        If SESSION_NO = "" Then
            SESSION_NO = Me.SESSION_NO
        End If
        If SELECTION_NO = 0 Then
            If ActiveForm Is Nothing Then
                Exit Sub
            End If
            SELECTION_NO = ActiveForm.SELECTION_NO
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

    Public Function Multi_Task_Menu_Item( _
    ByVal MENU_ITEM_TYPE As String, _
    ByVal MENU_ITEM_OBJECT As String, _
    ByVal OPEN_COUNT As Integer, _
    Optional ByVal menu_check As Boolean = False, _
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
                'MsgBox(EMsg, 0, "Cannot Make Requested Menu Selection At This Time")
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
        rowASTMTKC1.Item("ENTITY") = SESSION_NO
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
                    'MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Make Requested Menu Selection At This Time")
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
#End Region

    Public Function Format_Field(ByVal txt As String, ByVal COLUMN_NAME As String, Optional ByVal tbl As DataTable = Nothing, Optional ByVal treat_as_code As Boolean = False)

        Dim row As DataRow = tblASTFFMT1.Rows.Find(COLUMN_NAME)

        If row Is Nothing Then
            If treat_as_code Then
                row = tblASTFFMT1.NewRow
            Else
                If Not (tbl Is Nothing) Then
                    For Each DC As DataColumn In tbl.PrimaryKey
                        If DC.ColumnName = COLUMN_NAME Then
                            row = tblASTFFMT1.NewRow
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

    Public Function Get_Legend( _
    ByVal YYYYPP As String, _
    Optional ByVal AppendPeriodStatus As Boolean = True, _
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

    Public Function Get_Legend_Wk( _
    ByVal YYYYWW As String, _
    Optional ByVal abbreviated As Boolean = False)
        sql = "Select LEGEND from GLTPARM3 where YYYYWW = '" & YYYYWW & "'"
        Dim LEGEND As String = ASCDATA1.GetDataValue()
        If abbreviated Then
            LEGEND = Mid(LEGEND, 10, 7)
        End If
        Return LEGEND
    End Function

    Public Function Period_Calc( _
    ByVal base_YP As String, _
    ByVal number_of_periods As Integer) As String

        Dim p As Integer = Val(Mid$(base_YP, 1, 4)) * 12 + Val(Mid$(base_YP, 5, 2))
        p = p + number_of_periods

        Dim m As Integer
        Dim Y As Integer
        m = 1 + ((p - 1) Mod 12)
        Y = (p - m) / 12

        Return Format$(Y, "0000") & Format$(m, "00")

    End Function

    Public Function Period_Diff( _
    ByVal base_YP As String, _
    ByVal other_YP As String) As Integer
        Return 12 * (Val(Mid$(other_YP, 1, 4)) - Val(Mid$(base_YP, 1, 4))) + (Val(Mid$(other_YP, 5, 2)) - Val(Mid$(base_YP, 5, 2)))
    End Function

    Public Function Week_Calc( _
    ByVal base_YW As String, _
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

    Public Function Week_Diff( _
    ByVal base_YW As String, _
    ByVal other_YW As String) As Integer

        sql = "Select Count (*) from GLTPARM3" _
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
    Public Function SQL_CodeList( _
    ByVal TABLE_NAME As String, _
    ByVal COLUMN_NAME_key As String, _
    ByVal COLUMN_NAME_list As String, _
    Optional ByVal COLUMN_EXPRESSION_list As String = "", _
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

    Public Function Excel_Cell(ByVal R As Integer, ByVal C As Integer) As String

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

    Public Function Excel_Sheet_Name( _
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

    Public Function FormatTel(ByVal TEL As String, Optional ByVal EXT As String = "") As String
        FormatTel = "(" & Mid(TEL, 1, 3) & ")" & Mid(TEL, 4, 3) & "-" & Mid(TEL, 7, 4)
        If EXT <> "" Then
            FormatTel &= " x" & EXT
        End If
    End Function

    Public Sub AnalyzeTable( _
    ByVal TABLE_NAME As String, _
    Optional ByVal SCHEMA As String = "")
        If SCHEMA = "" Then
            SCHEMA = DBS_COMPANY
        End If
        If DBS_TYPE = DBS_TYPE_types.SQLServer Then
        Else
            'ASCDATA1.ExecuteSQL("Analyze Table " & TABLE_NAME & " Compute Statistics")
            ASCDATA1.ExecuteSQL("Begin dbms_stats.gather_table_stats('" & Chr(34) & SCHEMA & Chr(34) & "','" & Chr(34) & TABLE_NAME & Chr(34) & "', CASCADE=>TRUE, METHOD_OPT=>'FOR ALL COLUMNS SIZE 1'); End;")

        End If
    End Sub

    Public Function DEC(ByVal a As String) As String
        DEC = ""
        Dim i As Integer

        If a <> "" Then
            For i = 1 To Len(a)
                DEC = DEC & " " & Format(Asc(Mid(a, i, 1)), "000")
            Next
            DEC = Mid(DEC, 2)
        End If
    End Function

    Public Function HTA(ByVal a As String) As String
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

    Public Function ATH(ByVal a As String) As Long
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

    Public Function CheckDigitUPC(ByVal ValueToCheckDigit As String) As String

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

    Public Function CheckDigit( _
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
    Public Function AscToDecBytes( _
    ByVal AsciiString As String, _
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
    Public Function HexToAsc(ByVal HexString As String) As String
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
    Public Function AscToHex( _
    ByVal AsciiString As String, _
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
    Public Function HexToDec(ByVal HexString As String) As Long
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

    Public Function Launch_Attachment(ByVal ATTACHMENT_NO As String, ByVal ATTACHMENT_TYPE As String)

        Dim ataFileName As String = Folders("Attach") & ATTACHMENT_NO
        If Not My.Computer.FileSystem.FileExists(ataFileName) Then
            Return "Cannot Find File " & ataFileName
            Exit Function
        Else
            Dim finfo As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(ataFileName)
            ataFileName = finfo.FullName
        End If

        Select Case ATTACHMENT_TYPE

            Case "XLS", "XLSX"
                Try
                    'Infragistics.Documents.Excel.BIFF8Writer.WriteWorkbookToFile(myWorkbook, Folders("Work") & xlsFileName & ".xls")
                    'xlsFileName_sfx = ""
                    Dim xlsFileName = Folders("Work") & ATTACHMENT_NO & ".xls"
                    My.Computer.FileSystem.CopyFile(ataFileName, xlsFileName, True)

                    Dim excel As New Process
                    excel.StartInfo.Arguments = """" + ATTACHMENT_NO + """ /e"
                    excel.StartInfo.FileName = xlsFileName
                    excel.Start()

                Catch ex As Exception
                    'MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error trying to Launch Attachment")
                End Try

            Case "PDF", "BMP", "TXT", "JPG", "MSG", "DOC", "DOCX"
                Dim appFileName = Folders("Work") & ATTACHMENT_NO & "." & ATTACHMENT_TYPE
                Try
                    My.Computer.FileSystem.CopyFile(ataFileName, appFileName, True)
                    appFileName = My.Computer.FileSystem.GetFileInfo(appFileName).FullName

                    Dim p As Process = Process.Start(appFileName)
                Catch ex As Exception
                    'MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error trying to Launch Attachment")
                End Try

            Case Else
                Return "Unknown Application (" & ATTACHMENT_TYPE & ")"
                Exit Function
        End Select

        Return ""

    End Function

    Public Function GetPath(ByVal path As String) As String
        GetPath = path
        Try
            Dim fullname As String = My.Computer.FileSystem.GetDirectoryInfo(path).FullName
            If fullname <> "" Then
                GetPath = fullname
            End If
        Catch ex As Exception

        End Try
    End Function

    Public Function GetImageData(ByVal fileName As String) As Byte()
        'Method to load an image from disk and return it as a bytestream
        Dim fs As System.IO.FileStream = _
        New System.IO.FileStream(fileName, _
        System.IO.FileMode.Open, System.IO.FileAccess.Read)
        Dim br As System.IO.BinaryReader = New System.IO.BinaryReader(fs)
        Return (br.ReadBytes(Convert.ToInt32(br.BaseStream.Length)))

    End Function

    Public Function X_Sort( _
    ByRef a As String, _
    ByRef has_alpha As Boolean, _
    ByRef has_numeric As Boolean, _
    ByRef has_upper As Boolean, _
    ByRef has_lower As Boolean, _
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

    Public Function GemboxKey(Optional ByVal key As String = "") As String

        Select Case key
            Case ""
                Return "EW1Q-G14I-JKOW-4XS8" ' this key is ver version 3.7
                ' Return "EMPX-L9BW-EL8E-4GKJ" ' this key is ver version 3.3
                ' Return "EFYZ-QQSH-LE5Q-NJ7Y" this is the old 3.1 key
            Case Else
                Return ""
        End Select
    End Function

    Public Function nSoftwareKeys(ByVal key As String) As String

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
            Case "nSoftwareInship"
                Return TACMAIN1.nSoftwareInship
            Case "nSoftwaresftpkey"
                Return TACMAIN1.nSoftwaresftpkey
            Case Else
                Return ""
        End Select
    End Function

    Public Sub Record_Event( _
    ByVal TABLE_NAME As String, _
    ByVal TABLE_KEY As String, _
    ByVal TABLE_KEY2 As String, _
    ByVal INIT_DATE As Date, _
    ByVal INIT_OPER As String, _
    ByVal EVENT_TYPE As String, _
    ByVal EVENT_DESC As String, _
    ByVal EVENT_KEY As String)
        If TABLE_KEY2 <> "" Then
            TABLE_KEY &= ":" & TABLE_KEY2
        End If
        sql = "Insert into TATEVNT1 Values (:PARM1,:PARM2,:PARM3,:PARM4,:PARM5,:PARM6,:PARM7,:PARM8)"
        ASCDATA1.ExecuteSQL(sql, "VVDVVVVV", New Object() {TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY, ActiveForm.Name})
    End Sub

    Public Sub Get_Period_Range( _
     ByVal Number_of_Periods As Integer, _
     ByRef YP_Dates() As Date, _
     ByRef YP_Periods(,) As String, _
     Optional ByVal YP_Base As String = "")

        If YP_Base = "" Then
            YP_Base = CYP
        End If

        Dim YP0 As String = YP_Base
        Dim YP1 As String = Period_Calc(YP_Base, Number_of_Periods)

        Dim AD As String = ""
        If YP0 > YP1 Then
            YP0 = YP1
            YP1 = YP_Base
            AD = " DESC"
        End If

        Dim P As Integer = 0

        ReDim YP_Periods(Abs(Number_of_Periods), 1)
        ReDim YP_Dates(Abs(Number_of_Periods))

        sql = "Select * from GLTPARM2 " _
        & " where OPS_YYYYPP between '" & YP0 & "' and '" & YP1 & "'"
        P = 0
        For Each rowGLTPARM2 As DataRow In ASCDATA1.GetDataTable.Select("", "OPS_YYYYPP" & AD)
            YP_Periods(P, 0) = rowGLTPARM2.Item("OPS_YYYYPP")
            YP_Periods(P, 1) = Mid(rowGLTPARM2.Item("LEGEND"), 10, 6)
            YP_Dates(P) = rowGLTPARM2.Item("PRD_END_DATE")
            P += 1
        Next
    End Sub

    Public Sub Get_Week_Range( _
     ByVal Number_of_Weeks As Integer, _
     ByRef YW_Dates() As Date, _
     ByRef YW_Weeks(,) As String, _
     Optional ByVal YW_Base As String = "")

        If YW_Base = "" Then
            YW_Base = CYW
        End If

        Dim YW0 As String = YW_Base
        Dim YW1 As String = Week_Calc(YW_Base, Number_of_Weeks)

        Dim AD As String = ""
        If YW0 > YW1 Then
            YW0 = YW1
            YW1 = YW_Base
            AD = " DESC"
        End If

        Dim W As Integer = 0

        ReDim YW_Weeks(Abs(Number_of_Weeks), 1)
        ReDim YW_Dates(Abs(Number_of_Weeks))

        sql = "Select * from GLTPARM3 " _
        & " where YYYYWW between '" & YW0 & "' and '" & YW1 & "'"
        W = 0
        For Each rowGLTPARM3 As DataRow In ASCDATA1.GetDataTable.Select("", "YYYYWW" & AD)
            YW_Weeks(W, 0) = rowGLTPARM3.Item("YYYYWW")
            YW_Weeks(W, 1) = Mid(rowGLTPARM3.Item("LEGEND"), 10, 7)
            YW_Dates(W) = rowGLTPARM3.Item("WEEK_END_DATE")
            W += 1
        Next
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

    Public Function EnumToList(Of T)() As List(Of T)
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
    Public Function ConCatFiles(ByVal strTarget As String, ByVal inputFiles As List(Of String)) As Boolean

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

    Public Function ordinal(ByVal intNumber As Integer) As String
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

    Public Function XC( _
        ByVal C As Int16, _
        Optional ByVal R As Int16 = 0, _
        Optional ByVal absolute As Boolean = False) As String

        Dim COL As String = ""
        If C >= 1 Then
            Dim B As Int16 = (C - 1) Mod 26 + 1
            Dim A As Int16 = (C - B) / 26
            COL = Chr(Asc("A") + B - 1)
            If A > 0 Then
                COL = Chr(Asc("A") + A - 1) & COL
            End If
            If absolute Then
                COL = "$" & COL
            End If

            If R = 0 Then
                COL = COL & ":" & COL
            ElseIf R > 0 Then
                COL = COL & IIf(absolute, "$", "") & CStr(R)
            End If
        End If

        Return COL
    End Function


    Public Shared Function EncryptAES(strToHash As String) As String
        Dim Ezcrypt1 As New nsoftware.IPWorksEncrypt.Ezcrypt()
        ' SOMETHING IS NOT INSTANTIATED YET
        'Ezcrypt1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareEncryptionkey") ' ASCMAIN1.TACMAIN1.nSoftwareEncryptionkey
        Ezcrypt1.RuntimeLicense = "31454E3941413153554252415331544531414D483134323600000000000000000000000000000000303448445954314500005343315A53463257475453520000"
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
        ' SOMETHING IS NOT INSTANTIATED YET
        'Ezcrypt1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareEncryptionkey") ' ASCMAIN1.TACMAIN1.nSoftwareEncryptionkey
        Ezcrypt1.RuntimeLicense = "31454E3941413153554252415331544531414D483134323600000000000000000000000000000000303448445954314500005343315A53463257475453520000"
        Ezcrypt1.Reset()
        Ezcrypt1.Algorithm = nsoftware.IPWorksEncrypt.EzcryptAlgorithms.ezAES
        Ezcrypt1.UseHex = True
        Ezcrypt1.InputMessage = strToHash
        Ezcrypt1.KeyPassword = "0ff1c3"
        Ezcrypt1.Decrypt()
        Dim decrypted As String = Ezcrypt1.OutputMessage

        Return decrypted
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                oraCmd.Dispose()
                oraSP.Dispose()
                oraAda.Dispose()
                oraCon.Close()
                oraCon.Dispose()
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.       
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class