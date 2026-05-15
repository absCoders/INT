Imports System.Math
Imports Infragistics.Win

Public Class ABSC
    Public Parent_Form As ASFBASE1
    Public tlb As New Infragistics.Win.UltraWinToolbars.UltraToolbarsManager


#Region "Declarations"

    Public SEQs As Integer       ' Number of Columns Selected to Sort
    Public FORM_NAME As String   ' Form Object Name for Report (Me.Name)

    Public sql_SELECT_cols As String     ' PB Columns used in the Select List
    Public sql_GROUP_BY_cols As String   ' PB Columns used in the Group By
    Public sql_WHERE As String           ' PB where clause
    Public sql_TABLE_NAMEs As String     ' PB Tables
    Public sql_JOIN As String            ' PB Join
    Public sql_TABLE_NAME As String      ' PB Primary Table for Data Source

    Public sql_SELECT_COLUMNs As New List(Of String)     ' PB Columns used in the Select List

    Public COLUMN_NAMEs As New ArrayList         ' PB Column Names
    Public COLUMN_CAPTIONs As New ArrayList      ' PB Column Captions
    Public GROUP_ALL_OTHERSs As New ArrayList    ' Whether others should be grouped
    Public PAGE_BREAKs As String                       ' PB Level Page Breaks
    Public xErrMsg As String     ' if <> "" (error); exit Main Process in Sub Proceed
    Public COLUMN_NAME_first As String = ""  ' Forced First Column of a PB Report
    Public COLUMN_NAME_last As String = ""   ' Forced Last Column of a PB Report
    Public COLUMN_NAME_sum As New Dictionary(Of String, String)
    Public COLUMN_NAMEs_appended As String = ""
    Public tblASTDSQLH As DataTable
    Public tblASTDSQLS As DataTable
    Public PB_Report As Boolean = False
    Public ASTSRPT1 As String

    'Public COLUMN_appearance As New Dictionary(Of String, Infragistics.Win.Appearance)

    Public sql As String

    Public tblASTDSQLA As DataTable
    Public tblASTDSQLB As New DataTable
    Public tblASTDSQLC As New DataTable
    Public tblASTDSQLD As New DataTable
    Public tblASTDSQLE As New DataTable
    Public tblASTDSQLF As New DataTable
    Public tblASTDSQLJ As New DataTable

    Public Event Get_SQL_for_COLUMN_NAME(ByVal sender As Object, ByVal e As EventArgs)

#End Region

#Region "Eventargs"
    Public Class grdSetupClickCellEventArgs
        Inherits EventArgs

        Public COLUMN_NAME As String
        Public SQL As String
    End Class
#End Region

#Region "Properties"
    Public ReadOnly Property grdSetupDataSource() As DataTable
        Get
            Return DirectCast(grdSetup.DataSource, DataTable)
        End Get
    End Property



#End Region
#Region "Initialization"

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Me.ParentForm IsNot Nothing Then
            Parent_Form = DirectCast(Me.ParentForm, ASFBASE0)
        End If
        If Parent_Form IsNot Nothing Then
            FORM_NAME = Parent_Form.Name
            Build_Workfile_DB_Init()
            Setup_grdSetup()

            grdSetup.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False

            If ASCMAIN1.USER_ID <> "" Then
                With grdSetup.DisplayLayout.Bands(0)
                    .Columns("COLUMN_CAPTION").CellAppearance.BackColor = Drawing.Color.SkyBlue
                    '.Columns("EXCLUDE").CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
                    .Columns("EXCLUDE").CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.CellSelect
                    .Columns("EXCLUDE").Width = 30
                    .Columns("PAGE_BREAK").CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
                    .Columns("PAGE_BREAK").Hidden = True
                    .Columns("GROUP_ALL_OTHERS").CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
                    .Columns("GROUP_ALL_OTHERS").Hidden = True
                    .Columns("SEQUENCE").Width = 30
                End With
                ASCMAIN1.grdInitializeLayout(grdSetup)
                'Parent_Form.Load_Popup_Menu(grdSetup, "BB", "Save As", "Load Settings")
            End If
        End If

        AddHandler tlb.ToolClick, AddressOf tlb_ToolClick

        If ASCMAIN1.USER_ID <> "" Then
            ASCMAIN1.Load_Popup_Menu(tlb, grdSetup, "B", "Clear Values")
        End If
    End Sub
#End Region

#Region "grdSetup"

    Sub Clear_grdSetup(ByVal Clear_All As Boolean)
        grdSetup.UpdateData()
        grdSetup.ActiveRow = Nothing
        For Each dr As DataRow In DirectCast(grdSetup.DataSource, DataTable).Rows
            dr.Item("SEQUENCE") = DBNull.Value
            If Clear_All Then
                dr.Item("PAGE_BREAK") = "0"
                dr.Item("EXCLUDE") = "0"
                dr.Item("GROUP_ALL_OTHERS") = "0"
                dr.Item("CODE_VALUES") = ""
            End If
        Next
        SEQs = 0
        Call Re_SEQ()

    End Sub

    Sub Setup_grdSetup()
        Create_tblASTDSQLA()

        'Parent_Form.Get_PARM("GLTPARM1")

        Dim COLUMN_CAPTION As String = ""
        For Each dr As DataRow In ASCDATA1.GetDataTable("Select ASTDSQLA.COLUMN_NAME, NVL(ASTDSQLA.COLUMN_CAPTION,ASTDSQLK.COLUMN_CAPTION) COLUMN_CAPTION, ASTDSQLA.SORTABLE, ASTDSQLA.COLUMN_LAST from ASTDSQLA,ASTDSQLK WHERE ASTDSQLK.COLUMN_NAME (+) = ASTDSQLA.COLUMN_NAME and ASTDSQLA.FORM_NAME = '" & FORM_NAME & "' ORDER BY NVL(ASTDSQLA.COLUMN_CAPTION,ASTDSQLK.COLUMN_CAPTION)").Rows
            'If dr.Item("COLUMN_NAME") = "SEG2_CODE" And Parent_Form.ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & "" = "" _
            'Or dr.Item("COLUMN_NAME") = "SEG3_CODE" And Parent_Form.ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & "" = "" _
            'Or dr.Item("COLUMN_NAME") = "SEG4_CODE" And Parent_Form.ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & "" = "" _
            'Then
            '    ' SKIP IT
            'Else
            COLUMN_CAPTION = dr.Item("COLUMN_CAPTION") & ""
            'If dr.Item("COLUMN_NAME") = "SEG2_CODE" Then
            '    COLUMN_CAPTION = Parent_Form.ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC")
            'ElseIf dr.Item("COLUMN_NAME") = "SEG3_CODE" Then
            '    COLUMN_CAPTION = Parent_Form.ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC")
            'ElseIf dr.Item("COLUMN_NAME") = "SEG4_CODE" Then
            '    COLUMN_CAPTION = Parent_Form.ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC")
            'End If
            If dr.Item("SORTABLE") & "" = "1" Or dr.Item("COLUMN_LAST") & "" = "1" Then
                PB_Report = True
            End If
            If dr.Item("COLUMN_LAST") & "" = "1" Then
                COLUMN_NAME_last = dr.Item("COLUMN_NAME")
                dr.Item("SORTABLE") = "0"
                'PB_Report = True ?
            End If
            Call Add_Row(COLUMN_CAPTION, dr.Item("COLUMN_NAME") & "", dr.Item("SORTABLE") & "")
            'End If

        Next dr



        'If PB_Report Then
        '    UltraTabControl1.Tabs(0).Text = "Sort && Filter"
        'Else
        '    UltraTabControl1.Tabs(0).Text = "Filter"
        'End If

        If PB_Report Then
            tblASTDSQLH = ASCDATA1.GetDataTable("Select * from ASTDSQLH WHERE FORM_NAME = '" & FORM_NAME & "'")
            tblASTDSQLS = ASCDATA1.GetDataTable("Select * from ASTDSQLS WHERE FORM_NAME = '" & FORM_NAME & "'")
            'For Each row As DataRow In ASCDATA1.GetDataTable("Select * from ASTDSQLS WHERE FORM_NAME = '" & FORM_NAME & "' ORDER BY COLUMN_SEQ").Rows
            For Each row As DataRow In tblASTDSQLS.Select("", "COLUMN_SEQ")
                Try
                    COLUMN_NAME_sum.Add(row.Item("COLUMN_NAME"), row.Item("COLUMN_TYPE"))
                    If Val(row.Item("COLUMN_WIDTH") & "") = 0 Then
                        row.Item("COLUMN_WIDTH") = 80
                    End If
                    If row.Item("COLUMN_CAPTION") & "" = "" Then
                        row.Item("COLUMN_CAPTION") = row.Item("COLUMN_NAME")
                    End If

                Catch ex As Exception
                    MsgBox("Cannot Add Report Summary Column for " & row.Item("COLUMN_NAME"))
                End Try
            Next
        End If



        grdSetup.DataSource = tblASTDSQLA
        grdSetup.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChangeOrLostFocus

        SEQs = 0
        Call Re_SEQ()
        'grdSetup.DisplayLayout.Bands(0).SortedColumns.Add(grdSetup.DisplayLayout.Bands(0).Columns("COLUMN_CAPTION"), False)

        ' GET TO THE TOP
        If grdSetup.Rows.Count > 0 Then
            grdSetup.ActiveRow = grdSetup.Rows(0)
        End If

        grdSetup.UpdateData()

        grdSetup.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdSetup.DisplayLayout.Bands(0).SortedColumns.Add("COLUMN_CAPTION", False)
    End Sub

    Sub Create_tblASTDSQLA()
        tblASTDSQLA = New DataTable
        tblASTDSQLA.Columns.Add("COLUMN_NAME")
        tblASTDSQLA.Columns.Add("COLUMN_CAPTION")
        tblASTDSQLA.Columns.Add("CODE_VALUES")
        tblASTDSQLA.Columns.Add("EXCLUDE")
        tblASTDSQLA.Columns.Add("SEQUENCE", GetType(System.Int16))
        tblASTDSQLA.Columns.Add("PAGE_BREAK")
        tblASTDSQLA.Columns.Add("SORTABLE")
        tblASTDSQLA.Columns.Add("GROUP_ALL_OTHERS")
        tblASTDSQLA.Columns.Add("COLUMN_LAST")
        tblASTDSQLA.PrimaryKey = New DataColumn() {tblASTDSQLA.Columns("COLUMN_NAME")}
    End Sub

    Sub Add_Row(ByVal COLUMN_CAPTION As String, ByVal COLUMN_NAME As String, ByVal SORTABLE As String)
        Dim dr As DataRow
        dr = tblASTDSQLA.NewRow
        dr.Item("COLUMN_NAME") = COLUMN_NAME
        dr.Item("COLUMN_CAPTION") = COLUMN_CAPTION
        dr.Item("EXCLUDE") = "0"
        dr.Item("PAGE_BREAK") = "0"
        dr.Item("SORTABLE") = SORTABLE
        dr.Item("GROUP_ALL_OTHERS") = "0"
        tblASTDSQLA.Rows.Add(dr)
    End Sub

    Sub Re_SEQ( _
    Optional ByVal COLUMN_NAME As String = "", _
    Optional ByVal add_to_sort As Boolean = False)

        'grdSetup.Update 
        grdSetup.UpdateData()

        Dim tbl As DataTable = DirectCast(grdSetup.DataSource, DataTable)
        Dim row As DataRow

        If COLUMN_NAME <> "" Then
            row = tbl.Rows.Find(COLUMN_NAME)
            If add_to_sort Then
                row.Item("SEQUENCE") = 9
            Else
                row.Item("SEQUENCE") = System.DBNull.Value
                row.Item("PAGE_BREAK") = "0"
            End If
        End If

        If COLUMN_NAME_last <> "" Then
            row = tbl.Rows.Find(COLUMN_NAME_last)
            row.Item("SEQUENCE") = System.DBNull.Value
            row.Item("PAGE_BREAK") = "0"
        End If

        SEQs = 0
        For Each dr As DataRow In tbl.Select _
            ("SEQUENCE IS NOT NULL OR SEQUENCE <> ''", "SEQUENCE")
            SEQs = SEQs + 1
            dr.Item("SEQUENCE") = SEQs
        Next

        If COLUMN_NAME_last <> "" Then
            row = tbl.Rows.Find(COLUMN_NAME_last)
            SEQs = SEQs + 1
            row.Item("SEQUENCE") = SEQs
        End If

    End Sub

    Private Sub grdSetup_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSetup.AfterCellUpdate
        If e.Cell.Column.Key = "SEQUENCE" Then
            'grdSetup.UpdateData()
            grdSetup.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.CommitRow)
            'grdSetup.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode)
        End If
    End Sub

    Private Sub grdSetup_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSetup.BeforeRowUpdate

        If Val(e.Row.Cells("SEQUENCE").Value & "") = 0 Then
            grdSetup.DisplayLayout.Rows(e.Row.Index).Cells("PAGE_BREAK").Value = "0"
        End If

        Dim COLUMN_NAME As String = e.Row.Cells("COLUMN_NAME").Text ' grdSetup.ActiveRow.Cells("COLUMN_NAME").Text
        Dim sql As String = ASCMAIN1.CodeSelector.Get_SQL(COLUMN_NAME)
        If sql <> "" Then
            Dim CODE_VALUES_new As String = ""
            Dim CODE_VALUES As String = e.Row.Cells("CODE_VALUES").Text
            Dim KEY_EXPRESSION As String = ASCMAIN1.CodeSelector.grdColumns(0).Item("COLUMN_NAME")

            If CODE_VALUES <> "" Then
                Dim CODE_VALUES_old As String = ""
                For Each txt As String In Split(Replace(CODE_VALUES, "'", ""), ",")
                    CODE_VALUES_old = CODE_VALUES_old & ",'" & ASCMAIN1.Format_Field(txt, COLUMN_NAME, , True) & "'"
                Next
                CODE_VALUES_old = Mid$(CODE_VALUES_old, 2)
                Dim where_or_and As String = " where "
                If ASCMAIN1.CodeSelector.tblASTVIEW1.Rows(0).Item("WHERE_CLAUSE") & "" <> "" Then
                    where_or_and = " and "
                End If

                For Each dr As DataRow In ASCDATA1.GetDataTable(sql & where_or_and & KEY_EXPRESSION & " IN (" & CODE_VALUES_old & ")").Rows
                    CODE_VALUES_new = CODE_VALUES_new & "," & dr.Item(0)
                Next
            End If

            CODE_VALUES_new = Mid(CODE_VALUES_new, 2)
            If CODE_VALUES_new <> CODE_VALUES Then
                grdSetup.DisplayLayout.Rows(e.Row.Index).Cells("CODE_VALUES").Value = CODE_VALUES_new '  .ActiveRow.Cells("CODE_VALUES").Value = CODE_VALUES_new
            End If
        End If

    End Sub

    Private Sub grdSetup_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSetup.ClickCellButton
        If e.Cell.Column.Key = "SET_SEQ" Then
            If e.Cell.Row.Cells("SORTABLE").Text = "1" Then
                If e.Cell.Row.Cells("SEQUENCE").Text <> "" Then
                    Call Re_SEQ(e.Cell.Row.Cells("COLUMN_NAME").Text, False)
                Else
                    Call Re_SEQ(e.Cell.Row.Cells("COLUMN_NAME").Text, True)
                End If
            End If
        ElseIf e.Cell.Column.Key = "COLUMN_CAPTION" Then
            ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL(grdSetup.ActiveRow.Cells("COLUMN_NAME").Text)

            Dim ev As New grdSetupClickCellEventArgs
            ev.SQL = ASCMAIN1.CodeSelector.SQL
            ev.COLUMN_NAME = grdSetup.ActiveRow.Cells("COLUMN_NAME").Text
            RaiseEvent Get_SQL_for_COLUMN_NAME(Me, ev)
            ASCMAIN1.CodeSelector.SQL = ev.SQL

            If ASCMAIN1.CodeSelector.SQL <> "" Then
                ASCMAIN1.CodeSelector.MultipleSelections = True
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = Replace(grdSetup.ActiveRow.Cells("CODE_VALUES").Text & "", ",", Chr(0))
                Dim F As New ASFCODE1
                F.ShowDialog()
                F.Dispose()
                If ASCMAIN1.CodeSelector.Selections <> 0 Then
                    grdSetup.ActiveRow.Cells("CODE_VALUES").Value = Mid$(Replace(ASCMAIN1.CodeSelector.SelectedCodes0, Chr(0), ","), 2)
                    ' grdSetup.UpdateData()
                    grdSetup.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.CommitRow)

                End If
            End If
        End If
    End Sub

    Private Sub grdSetup_DoubleClickHeader(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickHeaderEventArgs) Handles grdSetup.DoubleClickHeader
        If e.Header.Column.Key = "SET_SEQ" Then
            Call Clear_grdSetup(False)
        End If
    End Sub

    Private Sub grdSetup_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSetup.InitializeRow
        If e.Row.Cells("SORTABLE").Text <> "1" Then
            e.Row.Cells("COLUMN_CAPTION").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Edit
        End If
    End Sub

    Private Sub grdSetup_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grdSetup.KeyDown
        If grdSetup.ActiveCell Is Nothing Then Exit Sub
        If e.KeyValue = Windows.Forms.Keys.Delete Then
            If grdSetup.ActiveCell.Column.Key = "SEQUENCE" Then
                If grdSetup.ActiveCell.Text <> "" Then
                    'grdSetup.ActiveCell.Value = DBNull.Value
                    'grdSetup.UpdateData()
                    Call Re_SEQ(grdSetup.ActiveRow.Cells("COLUMN_NAME").Text, False)
                End If
            End If
        End If

        If e.KeyValue = Windows.Forms.Keys.Enter Then
            If grdSetup.ActiveCell.Column.Key = "CODE_VALUES" Then
                grdSetup.Update()
            End If
        End If
    End Sub

    Private Sub grdSetup_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles grdSetup.KeyPress
        If grdSetup.ActiveCell IsNot Nothing Then
            If grdSetup.ActiveCell.Column.Key = "SEQUENCE" And grdSetup.ActiveRow.Cells("SORTABLE").Text = "1" Then
                Dim COLUMN_NAME As String = grdSetup.ActiveRow.Cells("COLUMN_NAME").Text
                Dim SEQcur As Integer = Val(grdSetup.ActiveCell.Text)
                Dim SEQnew As Integer = Val(e.KeyChar)
                If SEQnew < 1 Or SEQnew = SEQcur Or (SEQcur = 0 And SEQnew > SEQs + 1) Or (SEQcur <> 0 And SEQnew > SEQs) Then
                    Exit Sub
                End If

                grdSetup.ActiveCell.Value = SEQnew
                grdSetup.UpdateData()

                Dim i As Integer
                Dim z As String
                If SEQnew < SEQcur Or SEQcur = 0 Then
                    z = ">"
                    i = SEQnew
                Else
                    z = "<"
                    i = 0
                End If
                For Each dr As DataRow In DirectCast(grdSetup.DataSource, DataTable).Select("SEQUENCE " & z & "= " & CStr(SEQnew), "SEQUENCE")
                    If dr.Item("COLUMN_NAME") <> COLUMN_NAME Then
                        i = i + 1
                        dr.Item("SEQUENCE") = i
                    End If
                Next

                If SEQcur = 0 Then
                    SEQs = SEQs + 1
                End If
            End If
        End If
    End Sub

    Private Sub grdSetup_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSetup.Leave
        grdSetup.UpdateData()
    End Sub

    Private Sub grdSetup_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grdSetup.MouseUp

        If grdSetup.ActiveCell IsNot Nothing Then
            If grdSetup.ActiveCell.Column.Key = "EXCLUDE" Then
                grdSetup.UpdateData()
            End If
        End If
    End Sub
#End Region

#Region "Dynamic SQL - Generation of SQL"

    Public Function grdSetupCOLUMNs() As String()
        Dim C() As String = Nothing
        For Each row As DataRow In tblASTDSQLA.Select("ISNULL(SEQUENCE,0) <> 0", "SEQUENCE")
            If C Is Nothing Then
                ReDim C(0)
            Else
                ReDim Preserve C(C.Length)
            End If
            C(C.Length - 1) = row.Item("COLUMN_NAME")
        Next
        Return C
    End Function

    Sub Get_SQL( _
    ByVal DATA_SOURCE As String, _
    Optional ByVal TABLE_NAME_temp As String = "")

        Call ASCMAIN1.Track("Extracts from Data Sources", DATA_SOURCE)

        Dim TABLE_NAME As String
        Dim COLUMN_NAME As String
        Dim rowASTDSQLC As DataRow

        Dim sql_SELECT_col As String
        Dim sql_Select_col_count As Integer = 0

        DATA_SOURCE = IIf(DATA_SOURCE = "", "*", DATA_SOURCE)

        COLUMN_NAMEs.Clear()
        COLUMN_CAPTIONs.Clear()

        If PB_Report Then
            Dim rowASTDSQLB As DataRow = tblASTDSQLB.Rows.Find(New Object() {FORM_NAME, DATA_SOURCE})
            sql_TABLE_NAME = rowASTDSQLB.Item("TABLE_NAME")

            sql_SELECT_cols = ""
            sql_GROUP_BY_cols = ""
            sql_WHERE = ""
            sql_TABLE_NAMEs = ""
            sql_JOIN = ""

            ' Forced Joins - SHOULDN'T WE BE LOOKING AT J FOR THIS?

            For Each rowASTDSQLJ As DataRow In tblASTDSQLJ.Select("FORM_NAME = '" & FORM_NAME & "' and DATA_SOURCE = '" & DATA_SOURCE & "' and ALWAYS_JOIN = '1'")
                TABLE_NAME = rowASTDSQLJ.Item("TABLE_NAME")
                Call Get_SQL_Join_Criteria(TABLE_NAME, DATA_SOURCE)
            Next

            ' Sort

            sql_SELECT_COLUMNs.Clear()

            For Each rowASTDSQLA As DataRow In tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE")
                sql_SELECT_col = ""
                'sql_GROUP_BY_col = ""

                COLUMN_NAME = rowASTDSQLA.Item("COLUMN_NAME")
                COLUMN_NAMEs.Add(COLUMN_NAME)
                COLUMN_CAPTIONs.Add(rowASTDSQLA.Item("COLUMN_CAPTION"))

                rowASTDSQLC = tblASTDSQLC.Rows.Find(New Object() {FORM_NAME, DATA_SOURCE, COLUMN_NAME})


                If rowASTDSQLC Is Nothing Then
                    rowASTDSQLC = tblASTDSQLC.Rows.Find(New Object() {FORM_NAME, "*", COLUMN_NAME})
                End If

                TABLE_NAME = rowASTDSQLC.Item("TABLE_NAME") & ""
                If TABLE_NAME <> "" Then
                    If TABLE_NAME <> sql_TABLE_NAME Then
                        Call Get_SQL_Join_Criteria(TABLE_NAME, DATA_SOURCE)
                    End If
                Else
                    TABLE_NAME = sql_TABLE_NAME
                End If

                'If rowASTDSQLC.Item("EXPRESSION_IND") & "" = "1" Then
                If rowASTDSQLC.Item("COLUMN_EXPRESSION") & "" <> "" Then
                    sql_SELECT_col = rowASTDSQLC.Item("COLUMN_EXPRESSION") & ""
                    'sql_GROUP_BY_col = rowASTDSQLC.Item("COLUMN_EXPRESSION_Y") & ""
                Else
                    sql_SELECT_col = TABLE_NAME & "." & COLUMN_NAME
                End If

                If rowASTDSQLA.Item("GROUP_ALL_OTHERS") & "" = "1" And rowASTDSQLA.Item("CODE_VALUES") & "" <> "" Then
                    sql_SELECT_col = "Case When " & sql_SELECT_col & " in (" & Replace(rowASTDSQLA.Item("CODE_VALUES"), ",", "','") & "') Then " & sql_SELECT_col & " else '*' End"
                End If

                ' NEED TO DO THIS WITH A ROUTINE EXPOSED BY ASCMAIN1
                'ASCMAIN1.TACMAIN1.Get_Column_Expression_Exceptions(FORM_NAME, DATA_SOURCE, COLUMN_NAME, sql_SELECT_col) ' , sql_GROUP_BY_col)
                sql_SELECT_COLUMNs.Add(sql_SELECT_col)
                sql_SELECT_cols = sql_SELECT_cols & ", " & sql_SELECT_col & " AS " & COLUMN_NAME
                'If sql_GROUP_BY_col = "" Then
                '    sql_GROUP_BY_col = sql_SELECT_col
                'End If
                'sql_GROUP_BY_cols = sql_GROUP_BY_cols & ", " & sql_GROUP_BY_col
                sql_GROUP_BY_cols = sql_GROUP_BY_cols & ", " & sql_SELECT_col
                sql_Select_col_count = sql_Select_col_count + 1
            Next

            If COLUMN_NAMEs.Count > sql_Select_col_count Then
                sql_SELECT_col = COLUMN_NAMEs(COLUMN_NAMEs.Count - 1)
                sql_SELECT_cols = sql_SELECT_cols & ", " & sql_SELECT_col
                sql_GROUP_BY_cols = sql_GROUP_BY_cols & ", " & sql_SELECT_col
            End If

            sql_SELECT_cols = Mid$(sql_SELECT_cols, 3)
            sql_GROUP_BY_cols = Mid$(sql_GROUP_BY_cols, 3)
        End If

        ' Filter

        For Each rowASTDSQLA As DataRow In tblASTDSQLA.Select("CODE_VALUES is Not Null AND CODE_VALUES <> ''")
            COLUMN_NAME = rowASTDSQLA.Item("COLUMN_NAME")

            rowASTDSQLC = tblASTDSQLC.Rows.Find(New Object() {FORM_NAME, DATA_SOURCE, COLUMN_NAME})
            If rowASTDSQLC Is Nothing Then
                rowASTDSQLC = tblASTDSQLC.Rows.Find(New Object() {FORM_NAME, "*", COLUMN_NAME})
            End If

            TABLE_NAME = rowASTDSQLC.Item("TABLE_NAME") & ""
            If TABLE_NAME <> "" Then
                If TABLE_NAME <> sql_TABLE_NAME Then
                    Call Get_SQL_Join_Criteria(TABLE_NAME, DATA_SOURCE)
                End If
            Else
                TABLE_NAME = sql_TABLE_NAME
            End If

            If rowASTDSQLC.Item("JOIN_SPECIAL") & "" = "1" Then
                sql_SELECT_col = GetSpecialSelectedJoin(COLUMN_NAME, DATA_SOURCE)
            ElseIf rowASTDSQLC.Item("COLUMN_EXPRESSION") & "" <> "" Then
                sql_SELECT_col = rowASTDSQLC.Item("COLUMN_EXPRESSION") & ""
            Else
                sql_SELECT_col = TABLE_NAME & "." & COLUMN_NAME
            End If

            Dim in_or_equal As String
            Dim not_in_or_not_equal As String

            Dim CODE_VALUES_sql As String = "'" & Replace(rowASTDSQLA.Item("CODE_VALUES"), ",", "','") & "'"
            If InStr(CODE_VALUES_sql, ",") = 0 Then
                in_or_equal = "="
                not_in_or_not_equal = "<>"
            Else
                in_or_equal = "IN"
                not_in_or_not_equal = "NOT IN"
            End If

            If rowASTDSQLA.Item("EXCLUDE") & "" = "1" Then
                sql_WHERE = sql_WHERE & " AND (" & sql_SELECT_col & " IS NULL OR " & sql_SELECT_col & " " & not_in_or_not_equal & " (" & CODE_VALUES_sql & "))"
            Else
                sql_WHERE = sql_WHERE & " AND " & sql_SELECT_col & " " & in_or_equal & " (" & CODE_VALUES_sql & ")"
            End If
        Next

        'tblASTDSQLC = Nothing

        If PB_Report Then
            ' Pad sql_SELECT_cols for unused Group By's

            If COLUMN_NAMEs.Count < 9 Then
                For i As Integer = COLUMN_NAMEs.Count + 1 To 9
                    sql_SELECT_cols = sql_SELECT_cols & ", 'x' as G" & CStr(i)
                Next
                If COLUMN_NAMEs.Count = 0 Then
                    sql_SELECT_cols = Mid(sql_SELECT_cols, 3)
                End If
            End If

            ' Create Report Work File

            If ASTSRPT1 = "" Then
                COLUMN_NAMEs_appended = ""
                Dim sql_sum As String = ""
                Dim sql_sum_group_by As String = ""
                Dim sql As String = ""
                For i As Integer = 1 To 9
                    sql &= ",ASTSRPT1.G" & CStr(i)
                    sql_sum &= ",G" & CStr(i)
                Next
                sql = "Select " & Mid(sql, 2)
                sql_sum_group_by = Mid(sql_sum, 2)
                sql_sum = "Select " & sql_sum_group_by
                Dim ZTBL As String = ""
                If tblASTDSQLH.Rows.Count <> 0 Then
                    For Each ROW As DataRow In tblASTDSQLH.Select("", "COLUMN_SEQ")
                        Dim TABLE_NAME_appended_column As String = ROW.Item("TABLE_NAME") & ""
                        If TABLE_NAME_appended_column = "" Then
                            If TABLE_NAME_temp <> "" Then
                                TABLE_NAME_appended_column = TABLE_NAME_temp
                            Else
                                Dim rowASTDSQLB As DataRow = tblASTDSQLB.Rows.Find(New Object() {FORM_NAME, DATA_SOURCE})
                                TABLE_NAME_appended_column = rowASTDSQLB("TABLE_NAME") & ""
                            End If
                        End If
                        sql = sql & "," & TABLE_NAME_appended_column & "." & ROW.Item("COLUMN_NAME") & " " & ROW.Item("COLUMN_ALIAS")
                        If InStr(ZTBL, "," & TABLE_NAME_appended_column) = 0 Then
                            ZTBL &= "," & TABLE_NAME_appended_column
                        End If
                        Dim COLUMN_ALIAS As String = ROW.Item("COLUMN_ALIAS") & ""
                        If COLUMN_ALIAS = "" Then
                            COLUMN_ALIAS = ROW.Item("COLUMN_NAME")
                        End If
                        sql_sum &= "," & COLUMN_ALIAS
                        COLUMN_NAMEs_appended &= "," & COLUMN_ALIAS
                        sql_sum_group_by &= "," & COLUMN_ALIAS
                    Next
                End If

            End If
        End If

    End Sub

    Private Sub Get_SQL_Join_Criteria( _
    ByVal TABLE_NAME As String, _
    ByVal DATA_SOURCE As String)

        Dim rowASTDSQLJ As DataRow = tblASTDSQLJ.Rows.Find(New Object() {FORM_NAME, DATA_SOURCE, TABLE_NAME})
        If rowASTDSQLJ Is Nothing Then
            rowASTDSQLJ = tblASTDSQLJ.Rows.Find(New Object() {FORM_NAME, "*", TABLE_NAME})
        End If
        Dim TABLE_NAME_IS_ALIAS_FOR As String = rowASTDSQLJ.Item("TABLE_NAME_IS_ALIAS_FOR") & ""

        Dim TABLE_NAME_to_look_for As String = TABLE_NAME
        If TABLE_NAME_IS_ALIAS_FOR <> "" Then
            TABLE_NAME_to_look_for = TABLE_NAME_IS_ALIAS_FOR & " " & TABLE_NAME
        End If
        If InStr(sql_TABLE_NAMEs, "," & TABLE_NAME_to_look_for) <> 0 Then
            Exit Sub
        Else
            sql_TABLE_NAMEs = sql_TABLE_NAMEs & "," & TABLE_NAME_to_look_for
        End If

        Dim COLUMN_NAME As String
        Dim sql As String


        Dim JOIN_TYPE As String = ""
        If rowASTDSQLJ IsNot Nothing AndAlso rowASTDSQLJ.Item("OUTER_JOIN") & "" = "1" Then
            JOIN_TYPE = "(+)"
        End If

        Dim drsASTDSQLD() As DataRow
        sql = "FORM_NAME = '" & FORM_NAME & "' and TABLE_NAME = '" & TABLE_NAME _
            & "' and DATA_SOURCE = '" & DATA_SOURCE & "'"
        drsASTDSQLD = tblASTDSQLD.Select(sql)
        If drsASTDSQLD.Length = 0 Then
            sql = "FORM_NAME = '" & FORM_NAME & "' and TABLE_NAME = '" & TABLE_NAME _
                & "' and DATA_SOURCE = '" & "*" & "'"
            drsASTDSQLD = tblASTDSQLD.Select(sql)
        End If

        If drsASTDSQLD.Length = 0 Then
            Dim tbl As DataTable = ASCDATA1.GetDataTable("*", TABLE_NAME, -1, False)
            ReDim drsASTDSQLD(tbl.PrimaryKey.Length - 1)
            For i As Integer = 0 To tbl.PrimaryKey.Length - 1
                Dim row As DataRow = tblASTDSQLD.NewRow
                Dim dc As DataColumn = tbl.PrimaryKey(i)
                row.Item("FORM_NAME") = FORM_NAME
                row.Item("DATA_SOURCE") = DATA_SOURCE
                row.Item("TABLE_NAME") = TABLE_NAME
                row.Item("COLUMN_NAME") = dc.ColumnName
                drsASTDSQLD(i) = row
            Next
        End If

        Dim rowASTDSQLD As DataRow
        For i As Integer = 0 To UBound(drsASTDSQLD)
            rowASTDSQLD = drsASTDSQLD(i)

            If rowASTDSQLD.Item("TABLE_NAME_JOIN") & "" <> "" Then
                Call Get_SQL_Join_Criteria(rowASTDSQLD.Item("TABLE_NAME_JOIN"), DATA_SOURCE)
            End If

            COLUMN_NAME = rowASTDSQLD.Item("COLUMN_NAME")
            Dim TABLE_NAME_JOIN As String = rowASTDSQLD.Item("TABLE_NAME_JOIN") & ""
            If TABLE_NAME_JOIN = "" Then
                TABLE_NAME_JOIN = sql_TABLE_NAME
            End If

            sql = TABLE_NAME & "." & COLUMN_NAME & JOIN_TYPE & " = "

            'sql = JOIN_TYPE & TABLE_NAME & " ON " & COLUMN_NAME & " = "
            ''If rowASTDSQLD.Item("EXPRESSION_IND") & "" = "1" Then
            '' sql = sql & rowASTDSQLD.Item("COLUMN_NAME_JOIN")
            '' Else
            Dim COLUMN_NAME_JOIN As String = rowASTDSQLD.Item("COLUMN_NAME_JOIN") & ""
            If COLUMN_NAME_JOIN = "" Then
                COLUMN_NAME_JOIN = COLUMN_NAME
            End If
            If InStr(COLUMN_NAME_JOIN, ".") = 0 Then
                sql = sql & TABLE_NAME_JOIN & "." & COLUMN_NAME_JOIN
            Else
                sql = sql & COLUMN_NAME_JOIN
            End If
            'sql = sql & TABLE_NAME_JOIN & "." & COLUMN_NAME_JOIN
            'End If

            sql_JOIN = sql_JOIN & " AND " & sql
        Next

    End Sub

    Private Sub Get_SQL_Join_Criteria_Special(ByVal FORM_NAME As String, ByVal TABLE_NAME As String, ByVal DATA_SOURCE As String)

        'If TABLE_NAME = "" Then Exit Sub

        'jz = ""
        '' Special Conditions
        'Select Case FORM_NAME
        '    Case "SOFWHOD1"
        '        Select Case TABLE_NAME
        '            Case "ARTCUST1"
        '                If DATA_SOURCE = "A" Then
        '                    jz = "ARTCUST1.CUST_CODE (+) = Y2.CUST_CODE"
        '                Else
        '                    jz = "ARTCUST1.CUST_CODE (+) = X.CUST_CODE"
        '                End If

        '            Case "ICTITEM1"
        '                If DATA_SOURCE = "A" Then
        '                    jz = "ICTITEM1.ITEM_CODE (+) = Y1.ITEM_CODE"
        '                Else
        '                    jz = "ICTITEM1.ITEM_CODE (+) = X.ITEM_CODE"
        '                End If

        '        End Select

        'End Select

        'If jz <> "" And InStr(1, sqljoin, jz) = 0 Then
        '    sqljoin = sqljoin & " AND " & jz
        '    If InStr(sql_TABLE_NAMEs, "," & TABLE_NAME) = 0 Then 'Exit Sub
        '        sql_TABLE_NAMEs = sql_TABLE_NAMEs & "," & TABLE_NAME
        '    End If
        'End If

    End Sub

    Private Function GetSpecialSelectedJoin(ByVal COLUMN_NAME As String, ByVal DATA_SOURCE As String) As String
        Dim z As String = ""

        'Select Case FORM_NAME
        '    Case "SOFSLSF1"
        '        Select Case COLUMN_NAME
        '            Case "ITEM_BRAND_CODE"
        '        End Select
        '    Case "SOFWHOD1"
        '        Select Case COLUMN_NAME
        '            Case "CUST_CODE"
        '                If DATA_SOURCE = "A" Then
        '                    z = "Y2.CUST_CODE"
        '                Else
        '                    z = "X.CUST_CODE"
        '                End If

        '            Case "ITEM_CODE"
        '                If DATA_SOURCE = "C" Or DATA_SOURCE = "B" Then
        '                    z = ""
        '                ElseIf DATA_SOURCE = "A" Then
        '                    z = "Y1.ITEM_CODE"
        '                Else
        '                    z = "X.ITEM_CODE"
        '                End If
        '        End Select
        'End Select

        GetSpecialSelectedJoin = z

    End Function
#End Region

#Region "Supporting Routines"

    Function SQL_in(ByVal COLUMN_NAME As String, _
    Optional ByVal DB_COLUMN_NAME As String = "") As String

        Dim CODE_VALUES = SQLA(COLUMN_NAME, "CODE_VALUES", True)
        Dim sql As String = ""

        If CODE_VALUES <> "" Then
            Dim single_code_value As Boolean = (InStr(CODE_VALUES, "','") = 0)

            If single_code_value Then
                sql = sql & IIf(SQLA(COLUMN_NAME, "EXCLUDE") = "1", " <> ", " = ") & CODE_VALUES
            Else
                sql = sql & IIf(SQLA(COLUMN_NAME, "EXCLUDE") = "1", " NOT", "") & " in (" & CODE_VALUES & ")"
            End If

            If DB_COLUMN_NAME <> "" Then
                sql = " and " & DB_COLUMN_NAME & sql
            Else
                sql = " and " & COLUMN_NAME & sql
            End If
        End If

        Return sql
    End Function

    Public Function SQLA( _
    ByVal PB_COLUMN_NAME As String, _
    Optional ByVal COLUMN_NAME As String = "CODE_VALUES", _
    Optional ByVal SQL_List As Boolean = False) As String
        Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find(PB_COLUMN_NAME)
        If rowASTDSQLA Is Nothing Then
            SQLA = ""
        Else
            SQLA = rowASTDSQLA.Item(COLUMN_NAME) & ""
            If SQL_List And SQLA <> "" Then
                SQLA = "'" & Replace(SQLA, ",", "','") & "'"
            End If
        End If
        Return SQLA
    End Function

    Function SQLA_filter( _
    ByVal PB_COLUMN_NAME As String, _
    Optional ByVal DB_TABLE_NAME As String = "", _
    Optional ByVal DB_COLUMN_NAME As String = "") As String

        If DB_TABLE_NAME <> "" Then
            If DB_COLUMN_NAME = "" Then
                DB_COLUMN_NAME = DB_TABLE_NAME & "." & PB_COLUMN_NAME
            Else
                DB_COLUMN_NAME = DB_TABLE_NAME & "." & DB_COLUMN_NAME
            End If
        End If

        If DB_COLUMN_NAME = "" Then
            DB_COLUMN_NAME = PB_COLUMN_NAME
        End If

        Dim z As String
        z = SQLA(PB_COLUMN_NAME, "CODE_VALUES", True)
        If z <> "" Then
            SQLA_filter = " AND " & DB_COLUMN_NAME & IIf(SQLA(PB_COLUMN_NAME, "EXCLUDE") = "1", " NOT", "") & " IN (" & z & ")" & vbCr
        Else
            SQLA_filter = ""
        End If
        Return SQLA_filter
    End Function

    Function Get_Filter(ByVal COLUMN_NAME As String, ByVal SQL_ELEMENT_TO_COMPARE_TO As String) As String
        Dim sqlw As String = ""
        If SQLA(COLUMN_NAME, "CODE_VALUES") <> "" Then
            sqlw = " and " & SQL_ELEMENT_TO_COMPARE_TO & " " & IIf(SQLA(COLUMN_NAME, "EXCLUDE") = "1", "Not ", "") & "in (" & SQLA(COLUMN_NAME, "CODE_VALUES", True) & ")"
        End If
        Return sqlw

    End Function

#End Region

    Sub Build_Workfile_DB_Init()
        sql = "Select * from ASTDSQLF where FORM_NAME = '" & FORM_NAME & "'"
        tblASTDSQLF = ASCDATA1.GetDataTable(sql, "ASTDSQLF")
        If tblASTDSQLF.Rows.Count = 1 Then
            Dim OPTIMIZER_MODE As String
            OPTIMIZER_MODE = tblASTDSQLF.Rows(0).Item("OPTIMIZER_MODE") & ""
            If OPTIMIZER_MODE = "C" Then
                OPTIMIZER_MODE = "CHOOSE"
            ElseIf OPTIMIZER_MODE = "R" Then
                OPTIMIZER_MODE = "RULE"
            Else
                OPTIMIZER_MODE = ""
            End If
            If OPTIMIZER_MODE <> "" Then
                ASCDATA1.ExecuteSQL("ALTER SESSION SET OPTIMIZER_MODE = " & OPTIMIZER_MODE)
            End If
        End If

        ' Prepare for Dynamic SQL Generation

        tblASTDSQLB = ASCDATA1.GetDataTable("Select * from ASTDSQLB where FORM_NAME = '" & FORM_NAME & "'", "ASTDSQLB")

        ASCMAIN1.sql = "SELECT X.FORM_NAME, X.DATA_SOURCE, X.COLUMN_NAME, " _
            & "ASTDSQLC.TABLE_NAME, ASTDSQLC.COLUMN_EXPRESSION, ASTDSQLC.JOIN_SPECIAL " _
            & "FROM ASTDSQLC, ( " _
            & "SELECT ASTDSQLF.FORM_NAME, ASTDSQLB.DATA_SOURCE, ASTDSQLA.COLUMN_NAME " _
            & "FROM ASTDSQLF, ASTDSQLA, ASTDSQLB " _
            & "WHERE ASTDSQLF.FORM_NAME = '" & FORM_NAME & "' " _
            & "AND ASTDSQLF.FORM_NAME = ASTDSQLA.FORM_NAME " _
            & "AND ASTDSQLF.FORM_NAME = ASTDSQLB.FORM_NAME " _
            & ") X " _
            & "WHERE ASTDSQLC.FORM_NAME (+) = X.FORM_NAME " _
            & "AND ASTDSQLC.DATA_SOURCE (+) = X.DATA_SOURCE " _
            & "AND ASTDSQLC.COLUMN_NAME (+) = X.COLUMN_NAME "

        tblASTDSQLC = ASCDATA1.GetDataTable("", "ASTDSQLC", 3)
        tblASTDSQLC.Columns("COLUMN_EXPRESSION").ReadOnly = False
        tblASTDSQLD = ASCDATA1.GetDataTable("Select * from ASTDSQLD where FORM_NAME = '" & FORM_NAME & "'", "ASTDSQLD")
        tblASTDSQLJ = ASCDATA1.GetDataTable("Select * from ASTDSQLJ where FORM_NAME = '" & FORM_NAME & "'", "ASTDSQLJ")
    End Sub

    Private Sub tlb_ToolClick(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)

        Select Case e.Tool.Key
            Case "Clear Values"
                If grdSetup.ActiveRow IsNot Nothing Then
                    grdSetup.ActiveRow.Cells("CODE_VALUES").Value = ""
                    grdSetup.UpdateData()
                End If
        End Select

    End Sub
End Class