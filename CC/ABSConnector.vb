Imports ABSolution
Imports Infragistics.Win

Public Class ABSConnector

    Private clsASFBASE0 As New ABSolution.ASFBASE0
    Private clsASFBASE1 As New ABSolution.ASFBASE1
    Private clsASCBASE1 As New ABSolution.ASCBASE1
    Private objCCProcessor As Object
    Public dst As DataSet = New DataSet

    Public Sub New()

        clsASFBASE0 = New ABSolution.ASFBASE0
        clsASFBASE1 = New ABSolution.ASFBASE1
        clsASCBASE1 = New ABSolution.ASCBASE1

        dst = clsASCBASE1.dst

        clsASFBASE1.dst = clsASCBASE1.dst
    End Sub


#Region "Properties"

    Public ReadOnly Property Audit() As Dictionary(Of String, String)
        Get
            If clsASFBASE0.AUDIT Is Nothing Then
                clsASFBASE0.AUDIT = New Dictionary(Of String, String)
            End If
            Return (clsASFBASE0.AUDIT)
        End Get
    End Property


    Public Property DATETIME_STAMP() As Date
        Get
            Return clsASFBASE1.DATETIME_STAMP
        End Get
        Set(ByVal value As Date)
            clsASFBASE1.DATETIME_STAMP = value
        End Set
    End Property

    Public ReadOnly Property Folders(ByVal FolderName As String)
        Get
            Return ABSolution.ASCMAIN1.Folders(FolderName)
        End Get
    End Property

    Public ReadOnly Property UserID() As String
        Get
            Return ABSolution.ASCMAIN1.USER_ID
        End Get
    End Property

    Public ReadOnly Property CYP() As String
        Get
            Return ABSolution.ASCMAIN1.CYP
        End Get
    End Property

    Public ReadOnly Property CYM() As String
        Get
            Return ABSolution.ASCMAIN1.CYM
        End Get
    End Property

    Public ReadOnly Property NowTSD() As System.TimeSpan
        Get
            Return ABSolution.ASCMAIN1.NowTSD
        End Get
    End Property

    Public ReadOnly Property ActiveFormName() As String
        Get
            Return ABSolution.ASCMAIN1.ActiveForm.Name
        End Get
    End Property

#End Region

#Region "ABSolution Sub / Functions"

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

        clsASCBASE1.Create_TDA(tbl, TABLE_NAME, sql_custom, NumberOfKeysUsedToSelect, for_update, custom_parameters, Key_Field_Count, Update_COLUMN_NAMEs, SCHEMA)

    End Sub

    Public Function Fill_Records( _
    ByVal TABLE_NAME As String, _
    Optional ByVal Parameters() As Object = Nothing, _
    Optional ByVal ClearBeforeFilling As Boolean = True, _
    Optional ByVal Temp_Select As String = "", _
    Optional ByVal tblSubstitute As DataTable = Nothing) As Integer

        clsASCBASE1.Fill_Records(TABLE_NAME, Parameters, ClearBeforeFilling, Temp_Select, tblSubstitute)
    End Function

    Public Function Fill_Records( _
    ByVal TABLE_NAME As String, _
    ByVal KEY_VALUE As String, _
    Optional ByVal ClearBeforeFilling As Boolean = True, _
    Optional ByVal Temp_Select As String = "", _
    Optional ByVal tblSubstitute As DataTable = Nothing) As Integer
        clsASCBASE1.Fill_Records(TABLE_NAME, KEY_VALUE, ClearBeforeFilling, Temp_Select, tblSubstitute)
    End Function

    Public Function GetDataRow( _
           ByVal sql As String, _
           ByVal params As String, _
           ByVal ParamArray PARMs() As Object) As DataRow
        Return ABSolution.ASCDATA1.GetDataRow(sql, params, PARMs)
    End Function

    Public Function GetDataRow( _
         Optional ByVal sql_to_Execute As String = "", _
         Optional ByVal Return_Empty_Row_if_Missing As Boolean = False, _
         Optional ByVal custom_parameters As String = "", _
         Optional ByVal custom_parameter_values() As Object = Nothing) As DataRow
        Return ABSolution.ASCDATA1.GetDataRow(sql_to_Execute, Return_Empty_Row_if_Missing, custom_parameters, custom_parameter_values)
    End Function

    Public Function GetDataTable( _
           ByVal sql As String, _
           ByVal TABLE_NAME As String, _
           ByVal params As String, _
           ByVal ParamArray PARMs() As Object) As DataTable
        Return ABSolution.ASCDATA1.GetDataTable(sql, TABLE_NAME, params, PARMs)

    End Function

    Public Function GetDataTable( _
           Optional ByVal sql As String = "", _
           Optional ByVal TABLE_NAME As String = "", _
           Optional ByVal Key_Field_Count As Integer = -1, _
           Optional ByVal Include_Data As Boolean = True, _
           Optional ByVal NumberOfKeysUsedToSelect As Integer = -1, _
           Optional ByVal custom_parameters As String = "", _
           Optional ByVal custom_parameter_values() As Object = Nothing) _
           As DataTable
        Return ABSolution.ASCDATA1.GetDataTable(sql, TABLE_NAME, Key_Field_Count, Include_Data, NumberOfKeysUsedToSelect, custom_parameters, custom_parameter_values)
    End Function

    Public Function Add_Value_List( _
           ByVal cbe As UltraWinEditors.UltraComboEditor, _
           ByVal TABLE_COLUMN As String, _
           Optional ByVal RemoveCodes() As String = Nothing, _
           Optional ByVal AddCodes() As String = Nothing) As ValueList
        Return ABSolution.ASCMAIN1.Add_Value_List(cbe, TABLE_COLUMN, RemoveCodes, AddCodes)
    End Function

    Public Function Add_Value_List( _
            ByVal opt As UltraWinEditors.UltraOptionSet, _
            ByVal TABLE_COLUMN As String, _
            Optional ByVal RemoveCodes() As String = Nothing, _
            Optional ByVal AddCodes() As String = Nothing) As ValueList
        Return ABSolution.ASCMAIN1.Add_Value_List(opt, TABLE_COLUMN, RemoveCodes, AddCodes)
    End Function

    Public Function Add_Value_List( _
            ByVal grd As UltraWinGrid.UltraGrid, _
            ByVal COLUMN_NAME As String, _
            ByVal sql As String) As ValueList
        Return ABSolution.ASCMAIN1.Add_Value_List(grd, COLUMN_NAME, sql)
    End Function

    Public Function Add_Value_List( _
            ByVal grd As UltraWinGrid.UltraGrid, _
            ByVal TABLE_COLUMN As String, _
            Optional ByVal RemoveCodes() As String = Nothing, _
            Optional ByVal AddCodes() As String = Nothing, _
            Optional ByVal BandIndex As Integer = 0, _
            Optional ByVal sql As String = "") As ValueList
        Return ABSolution.ASCMAIN1.Add_Value_List(grd, TABLE_COLUMN, RemoveCodes, AddCodes, BandIndex, sql)
    End Function

    Public Sub grdClickCellButton( _
        ByRef grd As UltraWinGrid.UltraGrid, _
        Optional ByVal sql_where As String = "", _
        Optional ByVal commit_row As Boolean = False, _
        Optional ByVal COLUMN_NAME As String = "", _
        Optional ByVal VIEW_NAME As String = "")
        clsASFBASE1.grdClickCellButton(grd, sql_where, commit_row, COLUMN_NAME, VIEW_NAME)
    End Sub

    Public Sub Set_Read_Only( _
          ByRef ctl As Control, _
          ByVal tf As Boolean)
        clsASFBASE1.Set_Read_Only(ctl, tf)
    End Sub

    Public Function Logical_Lock( _
        ByVal ENTITY_TYPE As String, _
        ByVal ENTITY As String, _
        Optional ByVal menu_level As Boolean = False, _
        Optional ByVal show_message As Boolean = True, _
        Optional ByVal reverse_all_previous_if_unsuccessful As Boolean = True, _
        Optional ByVal MT_LEVEL As Integer = 0) As Boolean
        Return ABSolution.ASCMAIN1.Logical_Lock(ENTITY_TYPE, ENTITY, menu_level, show_message, reverse_all_previous_if_unsuccessful, MT_LEVEL)
    End Function

    Public Sub MultiTask_Release( _
           Optional ByVal SESSION_NO As String = "", _
           Optional ByVal SELECTION_NO As Integer = 0, _
           Optional ByVal MT_LEVEL As Integer = 0)
        ABSolution.ASCMAIN1.MultiTask_Release(SELECTION_NO, SELECTION_NO, MT_LEVEL)
    End Sub

    Public Sub BeginTrans()
        clsASFBASE1.BeginTrans()
    End Sub

    Public Sub CommitTrans(Optional ByVal Commit_Message As String = "")
        clsASFBASE1.CommitTrans(Commit_Message)
    End Sub

    Public Sub Rollback(Optional ByVal Error_Message As String = "", Optional ByRef ex2Record As Exception = Nothing)
        clsASFBASE1.Rollback(Error_Message, ex2Record)
    End Sub

    Public Function ExecuteSQL( _
            ByVal sql_to_Execute As String, _
            ByVal custom_parameters As String, _
            ByVal ParamArray PARMs() As Object) As Integer
        ABSolution.ASCDATA1.ExecuteSQL(sql_to_Execute, custom_parameters, PARMs)
    End Function

    Public Function ExecuteSQL( _
            Optional ByVal sql_to_Execute As String = "", _
            Optional ByVal ok_to_fail As Boolean = False) As Integer
        ABSolution.ASCDATA1.ExecuteSQL(sql_to_Execute, ok_to_fail)
    End Function

    Public Sub INIT_LAST(ByVal TABLE_NAME As String, _
            Optional ByVal UseRowState As Boolean = False, _
            Optional ByVal sqlWhere As String = "", _
            Optional ByVal LASTonINIT As Boolean = False)
        clsASFBASE1.INIT_LAST(TABLE_NAME, UseRowState, sqlWhere, LASTonINIT)
    End Sub

    Public Sub Write_Audit_Trail(ByVal row As DataRow, Optional ByVal FM_MODE As String = "")
        clsASFBASE1.Write_Audit_Trail(row, FM_MODE)
    End Sub

    Public Sub Write_Audit_Trail(ByRef row_current As DataRow, ByRef row_original As DataRow, Optional ByVal FM_MODE As String = "")
        clsASFBASE1.Write_Audit_Trail(row_current, row_original, FM_MODE)
    End Sub

    Public Sub Update_Record_TDA(ByVal TABLE_NAME As String, Optional ByVal sql_Delete As String = "")
        clsASCBASE1.Update_Record_TDA(TABLE_NAME, sql_Delete)
    End Sub

    'Public Sub Update_Record_TDA(ByVal TABLE_NAME As String, ByVal sql_Delete As String, ByVal custom_parameters As String, ByVal ParamArray PARMs() As Object)
    '    clsASCBASE1.Update_Record_TDA(TABLE_NAME, sql_Delete, custom_parameters, PARMs)
    'End Sub

    Public Function Get_Image( _
         ByVal IMAGE_FOLDER As String, _
         ByVal IMAGE_FILE As String) As System.Drawing.Bitmap
        Return ABSolution.ASCMAIN1.Get_Image(IMAGE_FOLDER, IMAGE_FILE)
    End Function

    Public Function Period_Calc(ByVal base_YP As String, ByVal number_of_periods As Integer) As String
        Return ABSolution.ASCMAIN1.Period_Calc(base_YP, number_of_periods)
    End Function

    Public Function LookUp( _
           ByVal TABLE_NAME As String, _
           ByVal KEY As String, _
           Optional ByVal Return_Empty_Row_if_Missing As Boolean = False) As DataRow
        Return clsASCBASE1.LookUp(TABLE_NAME, KEY, Return_Empty_Row_if_Missing)
    End Function

    Public Function LookUp( _
        ByVal TABLE_NAME As String, _
        ByVal KEY() As String, _
        Optional ByVal Return_Empty_Row_if_Missing As Boolean = False) As DataRow
        Return clsASCBASE1.LookUp(TABLE_NAME, KEY, Return_Empty_Row_if_Missing)
    End Function

    Public Function Next_Control_No( _
           ByVal CTL_NO_TYPE As String, _
           Optional ByVal How_Many As Long = 1) As String
        Return ABSolution.ASCMAIN1.Next_Control_No(CTL_NO_TYPE, How_Many)
    End Function

    Public Sub Create_Summary( _
            ByRef grd As UltraWinGrid.UltraGrid, _
            ByVal COLUMN_NAME As String, _
            Optional ByVal summary_type As String = "Sum", _
            Optional ByVal BandKey As String = "", _
            Optional ByVal mask As String = "", _
            Optional ByVal Calculator As Infragistics.Win.UltraWinGrid.ICustomSummaryCalculator = Nothing)
        clsASFBASE1.Create_Summary(grd, COLUMN_NAME, summary_type, BandKey, mask, Calculator)
    End Sub

    Public Sub Create_Summary( _
          ByRef grd As UltraWinGrid.UltraGrid, _
          ByVal COLUMN_NAMEs() As String, _
          Optional ByVal summary_type As String = "Sum", _
          Optional ByVal BandKey As String = "", _
          Optional ByVal mask As String = "", _
          Optional ByVal Calculator As Infragistics.Win.UltraWinGrid.ICustomSummaryCalculator = Nothing)
        clsASFBASE1.Create_Summary(grd, COLUMN_NAMEs, summary_type, BandKey, mask, Calculator)
    End Sub

    Public Sub Sort_grdColumns( _
            ByRef grd As UltraWinGrid.UltraGrid, _
            Optional ByVal COLUMN_NAMEs As String = "", _
            Optional ByVal LockSort As Boolean = False, _
            Optional ByVal BAND As Int32 = 0)
        clsASFBASE1.Sort_grdColumns(grd, COLUMN_NAMEs, LockSort, BAND)
    End Sub

#End Region

End Class
