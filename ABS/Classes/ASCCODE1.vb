Public Class ASCCODE1
    Public Selections As Integer    ' Number of Selected Codes
    Public SelectedCode As String   ' The Selected Code (i.e., CUST_CODE)
    Public SelectedCodes As New Collection  ' A collection of Selected Codes
    Public SelectedCodes0 As String ' Chr$(0) delimited string of Selected Codes
    Public PreviouslySelectedCodes0 As String   ' Used to start ASFCODE1 with Previous Selection(s)
    Public SelectedRows As New List(Of DataRow) ' Rows from the View for the Selected Codes
    Public MultipleSelections As Boolean    ' Indicating whether ASFCODE1 should allow Multiple vs Single Selections
    Public ForceFilterFirst As Boolean      ' Indicating whether the FilterFirst flag should be used regardless of how the view was set up
    Public DoNotFilterFirst As Boolean = False   ' to avoid filterfirst, perhaps because a constraining sqlwhere clause has been supplied
    Public SQL As String    ' Work Variable
    Public ParamTypes As String
    Public Params As Object()
    Public MODULE_ID As String  ' Providing Context for commonly used Column Names, ex: AR, AP, GL, SO
    Public VIEW_NAME As String  ' usually a Column Name
    Public VIEW_DESC As String  ' Description of the View - appears in the Caption of ASFCODE1
    Public TABLE_NAME As String ' providing context for the Column Name
    Public TABLE_NAME_temp As String ' optional temporary table containing the data to be shown in the view
    Public COLUMN_NAME As String    ' ?
    Public grdColumns As New List(Of DataRow)
    Public Hierarchal_Views As Collection       ' H Views supported in ASFCODE1 when this VIEW_NAME is used
    Public Hierarchal_Views_dr As Collection    ' H View Details
    Public tblASTVIEW1 As DataTable     ' This VIEW_NAME & TABLE_NAME
    Public Precedent_Keys As New Dictionary(Of String, Object)
    Public COLUMN_PREKEYs As New Dictionary(Of String, String)
    Public Caption As String  ' Overriding Caption for Code Selection Form
    Public PreFilter As New Dictionary(Of String, Object)
    Public UseDataFromTable As DataTable
    Public Custom_sql_where As String
    Public Custom_sqlkey As String

    ''' <summary>
    ''' Returns the SQL required to provide the VIEW_NAME specified.  If no TABLE_NAME is specified, then the first VIEW_NAME encountered will be used.
    ''' </summary>
    ''' <param name="VIEW_NAME"></param>
    ''' <param name="TABLE_NAME"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Get_SQL( _
    ByVal VIEW_NAME As String, _
    Optional ByVal TABLE_NAME As String = "", _
    Optional ByVal sql_where As String = "", _
    Optional ByVal key_value As String = "", _
    Optional ByVal TABLE_NAME_temp As String = "") As String

        Dim sql As String
        Dim i As Integer
        Dim j As Integer

        ASCMAIN1.CodeSelector.MODULE_ID = ""
        ASCMAIN1.CodeSelector.TABLE_NAME = ""
        ASCMAIN1.CodeSelector.COLUMN_NAME = ""
        ASCMAIN1.CodeSelector.Caption = ""
        ASCMAIN1.CodeSelector.PreFilter.Clear()
        ASCMAIN1.CodeSelector.COLUMN_PREKEYs.Clear()
        ASCMAIN1.CodeSelector.DoNotFilterFirst = False
        ASCMAIN1.CodeSelector.Custom_sqlkey = ""
        ASCMAIN1.CodeSelector.Custom_sql_where = ""

        'sql = "Select ASTVIEW1.*, ASTVIEW3.MODULE_ID AS MODULE_ID_RELATED"
        'sql = sql & " from ASTVIEW1 LEFT OUTER
        'ASTVIEW3 ON SUBSTR(ASTVIEW1.TABLE_NAME, 1, 2) = ASTVIEW3.MODULE_ID_RELATED"
        'sql = sql & " where VIEW_NAME = '" & VIEW_NAME & "'"
        'If TABLE_NAME <> "" Then
        '    sql = sql & " order by "
        '    sql = sql & " CASE "
        '    sql = sql & " WHEN ASTVIEW1.TABLE_NAME = '" & TABLE_NAME & "' THEN 0 "
        '    sql = sql & " WHEN ASTVIEW1.TABLE_NAME LIKE '" & Mid$(TABLE_NAME, 1, 2) & "%' THEN 1 "
        '    sql = sql & " WHEN ASTVIEW3.MODULE_ID = '" & Mid$(TABLE_NAME, 1, 2) & "' THEN 2 "
        '    sql = sql & " ELSE 9"
        '    sql = sql & " END"
        'End If
        'Dim tblASTVIEW1 As DataTable = ASCDATA1.GetDataTable(sql)

        ' need to do something about view3 

        Dim sqlwhere As String
        If TABLE_NAME <> "" Then
            sqlwhere = "VIEW_NAME = '" & VIEW_NAME & "' AND TABLE_NAME = '" & TABLE_NAME & "'"
        Else
            If ASCMAIN1.ActiveForm Is Nothing Then
                sqlwhere = "VIEW_NAME = '" & VIEW_NAME & "'"
            Else
                If ASCMAIN1.ActiveForm.MENU_ITEM_TYPE = "T" Then
                    ' I DON'T KNOW WHY TABLE_NAME PASSED IN HERE WASN'T THE ACTUAL FORM TABLE
                    If TABLE_NAME = "" Then
                        TABLE_NAME = ASCMAIN1.ActiveForm.TABLE_NAME
                    End If
                    sqlwhere = "VIEW_NAME = '" & VIEW_NAME & "' AND TABLE_NAME = '" & TABLE_NAME & "'"
                Else
                    sqlwhere = "VIEW_NAME = '" & VIEW_NAME & "' AND TABLE_NAME like '" & ASCMAIN1.ActiveForm.MODULE_ID & "*'"
                End If
            End If
        End If

        'Dim tblASTVIEW1 As DataTable = New DataView(ASCMAIN1.dstASTVIEWS.Tables("ASTVIEW1"), sqlwhere, "", DataViewRowState.CurrentRows).ToTable
        tblASTVIEW1 = New DataView(ASCMAIN1.dstASTVIEWS.Tables("ASTVIEW1"), sqlwhere, "", DataViewRowState.CurrentRows).ToTable
        If tblASTVIEW1.Rows.Count = 0 Then
            sqlwhere = "VIEW_NAME = '" & VIEW_NAME & "' and TABLE_NAME LIKE '" & Mid(TABLE_NAME, 1, 2) & "%'"
            tblASTVIEW1 = New DataView(ASCMAIN1.dstASTVIEWS.Tables("ASTVIEW1"), sqlwhere, "", DataViewRowState.CurrentRows).ToTable
            If tblASTVIEW1.Rows.Count = 0 Then
                sqlwhere = "VIEW_NAME = '" & VIEW_NAME & "'"
                tblASTVIEW1 = New DataView(ASCMAIN1.dstASTVIEWS.Tables("ASTVIEW1"), sqlwhere, "", DataViewRowState.CurrentRows).ToTable
            End If
        Else
            ' THIS SECTION IS NOT COMPLETE - IT NEEDS TO USE TABLE_NAME_ALT IF SPECIFIED WITHOUT A VIEW_NAME_ALT
            If tblASTVIEW1.Rows(0).Item("VIEW_NAME_ALT") & "" <> "" Then
                VIEW_NAME = tblASTVIEW1.Rows(0).Item("VIEW_NAME_ALT")
                sqlwhere = "VIEW_NAME = '" & VIEW_NAME & "' AND TABLE_NAME = '" & tblASTVIEW1.Rows(0).Item("TABLE_NAME_ALT") & "'"
                tblASTVIEW1 = New DataView(ASCMAIN1.dstASTVIEWS.Tables("ASTVIEW1"), sqlwhere, "", DataViewRowState.CurrentRows).ToTable
            End If
        End If

        If tblASTVIEW1.Rows.Count > 1 Then
            Dim MODULE_ID As String = ""
            If ASCMAIN1.ActiveForm IsNot Nothing Then
                MODULE_ID = Mid(ASCMAIN1.ActiveForm.Name, 1, 2)
                Dim row() As DataRow = tblASTVIEW1.Select("TABLE_NAME LIKE '" & MODULE_ID & "*'")
                If row.Length <> 0 Then
                    sqlwhere &= " and TABLE_NAME = '" & row(0).Item("TABLE_NAME") & "'"
                    tblASTVIEW1 = New DataView(ASCMAIN1.dstASTVIEWS.Tables("ASTVIEW1"), sqlwhere, "", DataViewRowState.CurrentRows).ToTable
                End If
            End If
            If tblASTVIEW1.Rows.Count > 1 Then
                Dim row() As DataRow = tblASTVIEW1.Select("TABLE_NAME LIKE 'TA*'")
                If row.Length <> 0 Then
                    sqlwhere &= " and TABLE_NAME = '" & row(0).Item("TABLE_NAME") & "'"
                    tblASTVIEW1 = New DataView(ASCMAIN1.dstASTVIEWS.Tables("ASTVIEW1"), sqlwhere, "", DataViewRowState.CurrentRows).ToTable
                End If
            End If
            If tblASTVIEW1.Rows.Count > 1 Then
                Dim row() As DataRow = tblASTVIEW1.Select("TABLE_NAME LIKE 'AS*'")
                If row.Length <> 0 Then
                    sqlwhere &= " and TABLE_NAME = '" & row(0).Item("TABLE_NAME") & "'"
                    tblASTVIEW1 = New DataView(ASCMAIN1.dstASTVIEWS.Tables("ASTVIEW1"), sqlwhere, "", DataViewRowState.CurrentRows).ToTable
                End If
            End If
            If tblASTVIEW1.Rows.Count > 1 Then
                Dim row() As DataRow = tblASTVIEW1.Select("")
                sqlwhere &= " and TABLE_NAME = '" & row(0).Item("TABLE_NAME") & "'"
                tblASTVIEW1 = New DataView(ASCMAIN1.dstASTVIEWS.Tables("ASTVIEW1"), sqlwhere, "", DataViewRowState.CurrentRows).ToTable
            End If
        End If

        If VIEW_NAME = "SEG2_CODE" Or VIEW_NAME = "SEG3_CODE" Or VIEW_NAME = "SEG4_CODE" Then
            tblASTVIEW1.Rows(0).Item("VIEW_DESC") = ASCDATA1.GetDataValue("Select GL_PARM_" & Mid$(VIEW_NAME, 1, 4) & "_DESC from GLTPARM1 where GL_PARM_KEY = 'Z'") & " Codes"
        End If


        Hierarchal_Views = New Collection
        Hierarchal_Views_dr = New Collection


        If tblASTVIEW1.Rows.Count = 0 Then
            Get_SQL = ""
        Else
            TABLE_NAME = tblASTVIEW1.Rows(0).Item("TABLE_NAME")
            VIEW_DESC = tblASTVIEW1.Rows(0).Item("VIEW_DESC") & ""

            If TABLE_NAME.Length > 0 AndAlso sqlwhere.Contains("'%'") Then
                sqlwhere = "VIEW_NAME = '" & VIEW_NAME & "' AND TABLE_NAME = '" & TABLE_NAME & "'"
            End If

            Dim tblASTVIEW2 As DataTable = New DataView(ASCMAIN1.dstASTVIEWS.Tables("ASTVIEW2"), sqlwhere, "COLUMN_POSITION", DataViewRowState.CurrentRows).ToTable
            Dim tblASTVIEW5 As DataTable = New DataView(ASCMAIN1.dstASTVIEWS.Tables("ASTVIEW5"), sqlwhere, "TABLE_NAME_JOIN", DataViewRowState.CurrentRows).ToTable

            grdColumns.Clear()

            If tblASTVIEW1.Rows(0).Item("CODE_TABLE") & "" <> "" Then
                TABLE_NAME = "ASTCODE1"
                tblASTVIEW1.Rows(0).Item("WHERE_CLAUSE") = "TABLE_NAME = '" & tblASTVIEW1.Rows(0).Item("CODE_TABLE") & "' AND COLUMN_NAME = '" & tblASTVIEW1.Rows(0).Item("CODE_COLUMN") & "'"
            End If

            Dim i_return As Integer = 0
            i = 0
            sql = ""
            For Each dr As DataRow In tblASTVIEW2.Rows
                i = i + 1
                If dr.Item("COLUMN_EXPRESSION") & "" <> "" Then
                    sql = sql & ", " & dr.Item("COLUMN_EXPRESSION")
                Else
                    If tblASTVIEW1.Rows(0).Item("CODE_TABLE") & "" <> "" Then
                        If i = 1 Then
                            dr.Item("COLUMN_NAME") = "T_CODE"
                        ElseIf i = 2 Then
                            dr.Item("COLUMN_NAME") = "T_DESC"
                        End If
                    End If

                    '  Modified  by Ed on 20100204
                    ' If the Column Name contains a dot then it is already pefixed by its table
                    If (dr.Item("COLUMN_NAME") & String.Empty).ToString.Contains(".") Then
                        sql = sql & ", " & dr.Item("COLUMN_NAME")
                    Else
                        sql = sql & ", " & TABLE_NAME & "." & dr.Item("COLUMN_NAME")
                    End If
                End If

                If dr.Item("COLUMN_ALIAS") & "" <> "" Then
                    sql = sql & " " & dr.Item("COLUMN_ALIAS")
                End If

                If dr.Item("COLUMN_PREKEY") & "" = "1" Then
                    COLUMN_PREKEYs.Add(dr.Item("COLUMN_NAME"), "")
                    ' RIGHT NOW, 2ND DIMENSION NOT BEING USED FOR ANYTHING- PERHAPS VALUES
                End If

                If dr.Item("COLUMN_NAME") = tblASTVIEW1.Rows(0).Item("COLUMN_NAME") & "" Then
                    i_return = i - 1
                End If

                grdColumns.Add(dr)
            Next

            Get_SQL = "Select " & Mid$(sql, 3) & " from " _
                & IIf(TABLE_NAME_temp = "", "", TABLE_NAME_temp & " ") & TABLE_NAME

            If tblASTVIEW1.Rows(0).Item("ADDL_TABLE_NAMES") & "" <> "" Then
                Get_SQL &= "," & tblASTVIEW1.Rows(0).Item("ADDL_TABLE_NAMES")
            End If

            If tblASTVIEW1.Rows(0).Item("WHERE_CLAUSE") & "" <> "" Then
                Get_SQL = Get_SQL & " where " & tblASTVIEW1.Rows(0).Item("WHERE_CLAUSE")
            End If

            For Each rowASTVIEW5 As DataRow In tblASTVIEW5.Rows
                Dim OUTER As String = ""
                If rowASTVIEW5.Item("OUTER_JOIN") & "" = "1" Then
                    OUTER = "OUTER "
                End If
                Get_SQL &= " LEFT " & OUTER & "JOIN " & rowASTVIEW5.Item("TABLE_NAME_JOIN")
                If rowASTVIEW5.Item("TABLE_NAME_ALIAS") & "" <> "" Then
                    Get_SQL &= " AS " & rowASTVIEW5.Item("TABLE_NAME_ALIAS")
                End If
                Get_SQL &= " ON " & rowASTVIEW5.Item("JOIN_CRITERIA")
            Next

            ' cache statement here, before adding optional where clause
            If sql_where <> "" Or key_value <> "" Then

                ASCMAIN1.CodeSelector.Custom_sql_where = sql_where
                ASCMAIN1.CodeSelector.Custom_sqlkey = key_value

                Dim sql_key_value As String = ""
                If key_value <> "" Then
                    Dim COLUMN_EXPRESSION As String = grdColumns(i_return).Item(9) & ""
                    Dim COLUMN_NAME As String = grdColumns(i_return).Item(3) & ""
                    Dim VIEW_NAME_key As String = grdColumns(i_return).Item(0) & ""
                    sql_key_value = " and " & IIf(COLUMN_EXPRESSION <> "", COLUMN_EXPRESSION, IIf(COLUMN_NAME = "", VIEW_NAME_key, COLUMN_NAME)) & " = '" & key_value & "'"
                End If

                If tblASTVIEW1.Rows(0).Item("WHERE_CLAUSE") & "" <> "" Then
                    If Trim(sql_where) <> "" And Mid(Trim(UCase(sql_where)), 1, 3) <> "AND" Then
                        sql_where = " and " & sql_where
                    End If
                    Get_SQL = Get_SQL & " " & sql_where & " " & sql_key_value
                Else
                    Get_SQL = Get_SQL & ASCMAIN1.SQL_Add_WHERE(sql_where & " " & sql_key_value)
                End If
            End If

            Dim HV(,) As String
            Dim tblASTVIEW4 As DataTable = New DataView(ASCMAIN1.dstASTVIEWS.Tables("ASTVIEW4"), sqlwhere, "HIERARCHAL_VIEW, HIERARCHAL_VIEW_LEVEL", DataViewRowState.CurrentRows).ToTable

            i = 0
            'Dim HVDR() As List(Of DataRow)
            Dim HVDR As New List(Of DataRow)

            ReDim HV(4, 0)
            For Each dr As DataRow In tblASTVIEW4.Rows
                If HV(0, i) <> dr.Item("HIERARCHAL_VIEW") Then
                    If i <> 0 Then
                        Dim DRX As DataRow = tblASTVIEW4.NewRow
                        DRX.Item("COLUMN_NAME_CODE") = grdColumns(0).Item("COLUMN_NAME")
                        DRX.Item("COLUMN_NAME_DESC") = grdColumns(1).Item("COLUMN_NAME")
                        DRX.Item("COLUMN_NAME_LINK") = ""
                        DRX.Item("HIERARCHAL_TABLE_NAME") = TABLE_NAME
                        HVDR.Add(DRX)

                        Hierarchal_Views_dr.Add(HVDR)
                        HVDR.Clear()
                    End If
                    i = i + 1
                    ReDim Preserve HV(4, i)
                    HV(0, i) = dr.Item("HIERARCHAL_VIEW")
                    'ReDim Preserve HVDR(i)
                    'HVDR(i) = New List(Of DataRow)
                End If
                HV(1, i) = HV(1, i) & " / " & dr.Item("HIERARCHAL_VIEW_LEVEL_DESC")
                j = j + 1
                'HVDR(i).Add(dr)
                HVDR.Add(dr)
                'HIERARCHAL_VIEW_NAME VARCHAR2(20),HIERARCHAL_TABLE_NAME VARCHAR2(8),
                'HIERARCHAL_VIEW_DESC VARCHAR2(20), COLUMN_NAME VARCHAR2(20),
            Next
            If i <> 0 Then
                'Dim LL As List(Of DataRow) = HVDR.
                Dim DRX As DataRow = tblASTVIEW4.NewRow
                DRX.Item("COLUMN_NAME_CODE") = grdColumns(0).Item("COLUMN_NAME")
                DRX.Item("COLUMN_NAME_DESC") = grdColumns(1).Item("COLUMN_NAME")
                DRX.Item("COLUMN_NAME_LINK") = ""
                DRX.Item("HIERARCHAL_TABLE_NAME") = TABLE_NAME
                HVDR.Add(DRX)

                Hierarchal_Views_dr.Add(HVDR)
                'HVDR.Clear()
            End If
            For i = 1 To UBound(HV, 2)
                Hierarchal_Views.Add(Mid$(HV(1, i), 4))
                '                Hierarchal_Views_dr.Add(HVDR(i))
            Next

            Me.VIEW_NAME = VIEW_NAME
            Me.TABLE_NAME = TABLE_NAME
        End If
    End Function
End Class