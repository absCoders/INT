Imports System.Math
Imports Oracle.ManagedDataAccess.Client
Imports GemBox.Spreadsheet

Public Class ASFBASE0
    Friend exbOLDKEY As String
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

    Public tblASTLOGX1 As New DataTable
    Public tdaASTLOGX1 As OracleDataAdapter

    ' Maintained by MENU_ITEM_OBJECT
    Public XNO As String = "" ' Generated each time you click Proceed, or Change Modes to True
    Public FORM_INSTANCE_NO As String  ' Generated each time you Launch from the Menu

    Public AUDIT As New Dictionary(Of String, String)

    Public TABLE_NAME As String = ""
    Public TABLE_NAME_view As String
    Public COLUMN_NAME As String = ""
    Public VIEW_NAME As String = ""
    Public CODE_VALUE As String = ""
    Public EMsg As String = ""
    Public EMsg2 As String = "" ' used to communicate error messages when loading grid from excel
    Public loading_grd_from_Excel As Boolean = False
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
    'Public BeforeRowsDeletedRows() As UltraWinGrid.UltraGridRow
    Public BeforeRowsDeletedRows As List(Of VariantType())
    Public HFs As New Dictionary(Of String, String)
    Public grdMRUs As New Dictionary(Of String, List(Of String))
    Public T As OracleTransaction
    Public Null As System.DBNull = DBNull.Value
    Public CMBs As New Dictionary(Of String, UltraWinGrid.UltraCombo)
    Public GRDs As New Dictionary(Of String, UltraWinGrid.UltraGrid)
    Public cmbYP As New Dictionary(Of String, cmbYPparms)
    Public cmbYW As New Dictionary(Of String, cmbYWparms)
    Public ctl1 As Control
    Public RE_XNO As Integer = -1
    Public RE_XNO_STAT As Integer = 0
    Public cdr As DataRow   ' Current DataRow returned from most recent LookUp or Validate_Code

    Dim disable_arrows As Boolean = False
    Dim doubleclicked As Boolean = False
    Dim tipDisplayed As Boolean = False
    Dim grdKeyValue As Integer
    Dim grdKeyData As System.Windows.Forms.Keys
    Dim grdError As UltraWinGrid.ErrorEventArgs

    Private gemboxUseXlsx As Boolean = False
    Private xlsExportAborted As Boolean = False

    Public clsASCBASE1 As New ASCBASE1(Me)
    Public frmASFBASE1s As New Dictionary(Of String, ASFBASE1)

    Public CurrentControl As Control
    Public CurrentGridColumn As String
    Public CurrentGridBand As String

    ' References to ASCBASE1 Objects
    Public dst As New DataSet
    Public TDAs As Dictionary(Of String, OracleDataAdapter)
    Public TBLs As Dictionary(Of String, DataTable)
    Public pROWs As Dictionary(Of String, DataRow)
    Public TBL_SCHEMAs As Dictionary(Of String, DataTable)
    Public DVWs As Dictionary(Of String, DataView)
    Protected F As ASFSRPTV
    Public CR_params As Dictionary(Of String, String)
    Public ROWs As Dictionary(Of String, DataRow)
    Public CMDs As Dictionary(Of String, OracleCommand)
    Public BA_CMDs As Dictionary(Of String, OracleCommand())
    Dim pressedKeys As New Dictionary(Of Keys, Boolean)
    Public oraDeps As New Dictionary(Of String, OracleDependency)

    Public REPORTS As New Dictionary(Of String, ASFSRPTM)
    Public ASTDATA1s As New Dictionary(Of String, String)

    Public Bound_DataSources As New List(Of Object)

    ' the tooltip that we will use when the cursor is over a cell of the grid
    Dim tooltip As New System.Windows.Forms.ToolTip()
    Public grds_with_Attachments As New Dictionary(Of String, Attachment_Button)
    Public FILENAMEs_to_Publish As New List(Of String)

    Private Structure strColumnFilters
        Dim band As Integer
        Dim ColumnName As String
        Dim filterCondition As Infragistics.Win.UltraWinGrid.FilterCondition
    End Structure

    Private grdColumnFilters As List(Of strColumnFilters)

    ' this allows our tooltips to have a delay before appearing
    Dim timer As New Timer()
    Public myWorkSheet As Infragistics.Documents.Excel.Worksheet
    Public eDND As System.Windows.Forms.DragEventArgs = Nothing
    Public ENTITY As Dropped_On_Entity
    Dim EXT_allowed As List(Of String)

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

    Public Structure Events_Entity
        Public TABLE_NAME As String
        Public TABLE_KEY As String
        Public TABLE_KEY_CAPTION As String
        Public TABLE_KEY_DESC As String
        Public read_only As Boolean
        Public enabled As Boolean
        Public TABLE_KEY_locked As Boolean
    End Structure

    Public Structure Dropped_On_Entity
        Public TABLE_NAME As String
        Public COLUMN_NAME As String
        Public CODE_VALUE As String
        Public DESC_VALUE As String
        Public ATTACHMENT_NOTES As String
        Public READ_ONLY As Boolean
        Public ATTACHMENT_NO As String
        Public RESTRICTIONS As String
        Public CUSTOM_SQL As String
        Public OTHER_ENTITIES As List(Of Dropped_On_Entity_Other)
    End Structure

    Public Overridable Sub Attachments_Changed()

    End Sub

    Public Structure Dropped_On_Entity_Other
        Public TABLE_NAME As String
        Public COLUMN_NAME As String
        Public COLUMN_NAME_linked As String
        Public sql_for_link As String
    End Structure

    Public Structure Audit_Entity
        Public TABLE_NAME As String
        Public TABLE_DESC As String
        Public KEY_VALUE As String
        Public KEY_DESC As String
    End Structure

    Public Structure cmbYPparms
        Public Base_YYYYPP As String
        Public RelativeStartingPeriod As Integer
        Public RelativeEndingPeriod As Integer
        Public RelativeDefaultPeriod As Integer
        Public TotalRelativePeriods As Integer
        Public Parent_cmbYP As String
        Public Child_cmbYP As String
    End Structure

    Public Structure cmbYWparms
        Public Base_YYYYWW As String
        Public RelativeStartingWeek As Integer
        Public RelativeEndingWeek As Integer
        Public RelativeDefaultWeek As Integer
        Public TotalRelativeWeeks As Integer
        Public Parent_cmbYW As String
        Public Child_cmbYW As String
    End Structure

    Public Structure grdSetupParms
        Public SEQs As Int16
        Public COLUMN_NAME_last As String
        Public FORM_NAME As String
        Public PB_Report As Boolean
    End Structure

    Public cmbYPparm As cmbYPparms
    Public cmbYWparm As cmbYWparms

    Public Structure Attachment_Button
        Public TABLE_NAME As String
        Public COLUMN_NAME As String
        Public allow_update As String
        Public grd As UltraWinGrid.UltraGrid
    End Structure

    ' Lets us know a scan is comming our way
    ' KeyDown sets it to true when the beginning of a scan
    ' All other keys set to false
    ' The form that accepts scans should do processing in the Forms KeyDown event
    ' Forexample, set focus to a control depending on the state of the form.
    Public ScannerInUse As Boolean = False
    Public kill_switch As Boolean = False

#Region "Checkbox On Grid Header"

    ' At the end of this Class is CheckBoxOnHeader_CreationFilter, the class that does the work

    ''' <summary>
    '''  Create an instance of the CreationFilter. This needs to be Form-Level so 
    '''  we can catch the event that fires when the CheckBox is clicked. 
    ''' </summary>
    ''' <remarks></remarks>
    Dim WithEvents aCheckBoxOnHeader_CreationFilter As New CheckBoxOnHeader_CreationFilter()

    ''' <summary>
    ''' This event on the CreationFilter fires when the CheckBox in a Header is clicked. 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub aCheckBoxOnHeader_CreationFilter_HeaderCheckBoxClicked(ByVal sender As Object, ByVal e As CheckBoxOnHeader_CreationFilter.HeaderCheckBoxEventArgs) Handles aCheckBoxOnHeader_CreationFilter.HeaderCheckBoxClicked
        ' Check to see if the column is of style checkbox.  If it is, set all the cells in that column to
        ' whatever value the header checkbox is.
        Dim aRow As UltraWinGrid.UltraGridRow
        Dim level As Integer = e.Header.Column.Level

        If e.Header.Column.Style = UltraWinGrid.ColumnStyle.CheckBox Then
            For Each aRow In e.Rows
                aRow.Cells(e.Header.Column.Index).Value = IIf((e.CheckState.Equals(CheckState.Checked)), "1", "0")
                aRow.Update()
            Next
        End If
    End Sub

    ''' <summary>
    ''' Sub to tell the Creation filter what columns in a grid are to have the Checkbox
    ''' If no list of columns is provided then all Check Box columns have the Header Checkbox
    ''' </summary>
    ''' <param name="grd"></param>
    ''' <remarks></remarks>
    Public Sub DisplayHeaderCheckBox(ByRef grd As Infragistics.Win.UltraWinGrid.UltraGrid)
        Dim columnList As List(Of String) = New List(Of String)
        DisplayHeaderCheckBox(grd, Nothing)
    End Sub

    ''' <summary>
    ''' Sub to tell the Creation filter what columns in a grid are to have the Checkbox
    ''' If no list of columns is provided then all Check Box columns have the Header Checkbox
    ''' </summary>
    ''' <param name="grd"></param>
    ''' <param name="fieldColumnList"></param>
    ''' <remarks></remarks>
    Public Sub DisplayHeaderCheckBox(ByRef grd As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal fieldColumnList() As String)

        Dim columnList As List(Of String) = New List(Of String)

        If fieldColumnList IsNot Nothing Then
            For Each key As String In fieldColumnList
                If key.Length > 0 Then
                    columnList.Add(key)
                End If
            Next
        End If

        ' Set the columns to receive the checkbox
        aCheckBoxOnHeader_CreationFilter.ColumnNames = columnList
        grd.CreationFilter = aCheckBoxOnHeader_CreationFilter
    End Sub

#End Region

    Private Sub ASFBASE0_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        'If ASCMAIN1.USER_ID = "" Then
        '    Exit Sub
        'End If

        ASCMAIN1.tblASTSQLX1 = dst.Tables("ASTSQLX1")
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(ASFMAIN1.UltraToolbarsManager1.Tools("Show Tool-Tips"), UltraWinToolbars.StateButtonTool)
        tip.Enabled = tlb_sbt.Checked

        If Me.OwnedForms.Count <> 0 Then
            For Each frm As Form In Me.OwnedForms
                If frm.Name <> "FloatingWindowContainer" Then
                    frm.Visible = True
                End If
            Next
        End If

    End Sub

    Private Sub ASFBASE0_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        If ASCMAIN1.tblASTSQLX1 IsNot Nothing Then
            If Not dst.Tables.Contains("ASTSQLX1") OrElse dst.Tables("ASTSQLX1") Is Nothing Then
                ASCMAIN1.tblASTSQLX1 = Nothing
            Else
                If ASCMAIN1.tblASTSQLX1.Equals(dst.Tables("ASTSQLX1")) Then
                    ASCMAIN1.tblASTSQLX1 = Nothing
                End If
            End If
        End If

        If Me.OwnedForms.Count <> 0 Then
            For Each frm As Form In Me.OwnedForms
                frm.Visible = False
            Next
        End If

    End Sub

    Private Sub ASFBASE0_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed

        If dst IsNot Nothing Then
            dst.Dispose()
        End If
        If tblASFBASE1 IsNot Nothing Then
            tblASFBASE1.Dispose()
        End If
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

        If oraDeps.Count > 0 Then
            For Each T As String In oraDeps.Keys
                Dim dep As OracleDependency = oraDeps(T)
                dep.RemoveRegistration(ASCMAIN1.oraCon)
                dep = Nothing
            Next
        End If

        If TBLs IsNot Nothing AndAlso TBLs.Count <> 0 Then
            For Each TBL_key As String In TBLs.Keys
                Dim tbl As DataTable = TBLs(TBL_key)
                tbl.Dispose()
            Next
        End If
        TBLs = Nothing

        If TBL_SCHEMAs IsNot Nothing AndAlso TBL_SCHEMAs.Count <> 0 Then
            For Each TBL_SCHEMAs_key As String In TBL_SCHEMAs.Keys
                Dim TBL_SCHEMA As DataTable = TBL_SCHEMAs(TBL_SCHEMAs_key)
                TBL_SCHEMA.Dispose()
            Next
        End If
        TBL_SCHEMAs = Nothing

        If TDAs IsNot Nothing AndAlso TDAs.Count <> 0 Then
            For Each TDA_key As String In TDAs.Keys
                Dim tda As OracleDataAdapter = TDAs(TDA_key)
                If tda IsNot Nothing Then
                    tda.Dispose()
                End If
            Next
        End If
        TDAs = Nothing

        If tblASTOPST1 IsNot Nothing Then
            tblASTOPST1.Dispose()
        End If
        If tdaASTOPST1 IsNot Nothing Then
            tdaASTOPST1.Dispose()
        End If

        If clsASCBASE1 IsNot Nothing Then
            clsASCBASE1.Dispose()
        End If
        clsASCBASE1 = Nothing

        dst = Nothing
        tblASFBASE1 = Nothing
        rowASFBASE1 = Nothing
        htbCOLUMN_NAME = Nothing
        CMDs = Nothing
        TBLs = Nothing
        TBL_SCHEMAs = Nothing
        cdr = Nothing
        TDAs = Nothing
        HFs = Nothing
        rowASTOPST1 = Nothing
        tblASTOPST1 = Nothing
        tdaASTOPST1 = Nothing
    End Sub

    Private Sub ASFBASE0_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        If My.Application.OpenForms.Count = 1 Then
            If ASFMAIN1.UltraPictureBox1.Image IsNot Nothing Then
                ASFMAIN1.UltraPictureBox1.Visible = True
            End If
        End If

        If ASCMAIN1.USER_MENU_ITEM_OBJECT <> "" Then
            If My.Application.OpenForms.Count = 1 Then
                End
            End If
        End If
    End Sub

    Private Sub ASFBASE0_HelpRequested(ByVal sender As Object, ByVal hlpevent As System.Windows.Forms.HelpEventArgs) Handles Me.HelpRequested
        Dim HELP_FILENAME As String = ASCMAIN1.SOLUTION & ".CHM"
        Help.ShowHelp(Me, ASCMAIN1.Folders("Help") & HELP_FILENAME, HelpNavigator.Topic, "HTML\" & Me.Name & ".HTM")
    End Sub

#Region "Key Press"
    Private Sub ASFBASE0_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        ' Used to ignore scanner codes
        If e.KeyValue = 17 Or e.KeyValue = 220 Then
            ScannerInUse = True
            e.SuppressKeyPress = True
            Exit Sub
        Else
            ScannerInUse = False
        End If

        Select Case e.KeyData
            Case Keys.F2, Keys.Shift Or Keys.F2
                If ASCMAIN1.Running_in_VS Or ASCMAIN1.USER_SECURITY_CODEs.Contains("SY") Then

                    Dim ctl As Control = Me.ActiveControl
                    Dim cmp As System.ComponentModel.Component = DirectCast(Me.ActiveControl, System.ComponentModel.Component)
                    If ctl IsNot Nothing Then
                        Dim FORM_NAME As String = Me.Name
                        Dim TABLE_NAME_TIP As String = Absx1.GetABSTableName(ctl)
                        If TABLE_NAME_TIP = "" Then
                            TABLE_NAME_TIP = TABLE_NAME
                        End If
                        'Dim txtctl As UltraWinEditors.UltraTextEditor = DirectCast(ctl, UltraWinEditors.UltraTextEditor)
                        'Dim COLUMN_NAME As String = ASCMAIN1.ActiveForm.Absx1.GetABSColumnName(ctl)

                        Dim COLUMN_NAME As String = ""

                        If e.Shift Then
                            TABLE_NAME_TIP = "*"
                            COLUMN_NAME = "HEADER"
                        Else
                            If TypeOf CurrentControl Is UltraWinGrid.UltraGrid Then
                                TABLE_NAME_TIP = CurrentGridBand
                                COLUMN_NAME = CurrentGridColumn
                            Else
                                If CurrentControl Is Nothing Then
                                    COLUMN_NAME = ""
                                Else
                                    COLUMN_NAME = Absx1.GetABSColumnName(CurrentControl)
                                End If
                            End If
                        End If

                        If TABLE_NAME_TIP.Length > 8 Then
                            Exit Sub ' lauren hits F2 to edit and this is no good when tablename is IPSA_ICTITEM1
                        End If

                        If Not ASCMAIN1.ActiveForm.dst.Tables.Contains("ASTTTIP1") Then
                            Create_TDA(ASCMAIN1.ActiveForm.dst.Tables.Add, "ASTTTIP1", "*")
                        End If

                        Dim rowASTTTIP1 As DataRow = ASCMAIN1.ActiveForm.dst.Tables("ASTTTIP1").Rows.Find _
                        (New String() {FORM_NAME, TABLE_NAME_TIP, COLUMN_NAME})
                        If rowASTTTIP1 Is Nothing Then
                            rowASTTTIP1 = dst.Tables("ASTTTIP1").NewRow
                            rowASTTTIP1.Item("FORM_NAME") = FORM_NAME
                            rowASTTTIP1.Item("TABLE_NAME") = TABLE_NAME_TIP
                            rowASTTTIP1.Item("COLUMN_NAME") = COLUMN_NAME
                            If TABLE_NAME_TIP = "*" Then
                                rowASTTTIP1.Item("TOOLTIP_TITLE") = MENU_ITEM_DESC  ' Me.Text
                            Else
                                rowASTTTIP1.Item("TOOLTIP_TITLE") = ASCMAIN1.Make_Caption(COLUMN_NAME)
                            End If

                            Dim TOOLTIP_TEXT As String = "" ' "{Text goes here}"
                            TOOLTIP_TEXT = "This screen is used to maintain records in the " & MENU_ITEM_DESC & " table."
                            rowASTTTIP1.Item("TOOLTIP_TEXT") = TOOLTIP_TEXT
                            rowASTTTIP1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                            rowASTTTIP1.Item("INIT_DATE") = DATETIME_STAMP

                            dst.Tables("ASTTTIP1").Rows.Add(rowASTTTIP1)

                            If TABLE_NAME_TIP = "*" And COLUMN_NAME = "HEADER" Then
                                COLUMN_NAME = "BODY"
                                Dim row2 As DataRow = dst.Tables("ASTTTIP1").NewRow
                                row2.ItemArray = rowASTTTIP1.ItemArray
                                row2.Item("COLUMN_NAME") = COLUMN_NAME
                                row2.Item("TOOLTIP_TEXT") = "{Text goes here}"
                                dst.Tables("ASTTTIP1").Rows.Add(row2)
                                rowASTTTIP1 = row2
                            End If
                        End If

                        Dim f As New ASFTTIP1
                        f.rowASTTTIP1 = rowASTTTIP1
                        f.FP = Me

                        f.ShowDialog()
                        f.Dispose()
                        f = Nothing

                        Me.Cursor = Cursors.WaitCursor
                        Me.Cursor = Cursors.Default

                        If CurrentControl IsNot Nothing Then
                            CurrentControl.Select()
                        End If

                    End If
                End If
        End Select

        If e.Control AndAlso e.KeyCode = Keys.A Then
            SelectAll()
        End If

        If Not pressedKeys.ContainsKey(e.KeyCode) Then
            pressedKeys.Add(e.KeyCode, True)
        End If
        'EvaluatePressedKeys()


    End Sub
    Private Sub ASFBASE0_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        If pressedKeys.ContainsKey(e.KeyCode) Then
            pressedKeys.Remove(e.KeyCode)
        End If
        'EvaluatePressedKeys()
    End Sub

    Private Sub EvaluatePressedKeys()
        If pressedKeys.ContainsKey(Keys.ControlKey) AndAlso pressedKeys.ContainsKey(Keys.A) Then
            If TypeName(ASCMAIN1.ActiveForm.CurrentControl) = "UltraGrid" Then
                Dim activeGrid As UltraWinGrid.UltraGrid = ASCMAIN1.ActiveForm.CurrentControl
                activeGrid.Selected.Rows.AddRange(CType(activeGrid.Rows.All, UltraWinGrid.UltraGridRow()))
                If Not IsNothing(activeGrid.ActiveCell) Then
                    activeGrid.ActiveCell.Row.Activate()
                End If
            End If
        End If
    End Sub
    Private Sub SelectAll()


        Select Case TypeName(ASCMAIN1.ActiveForm.CurrentControl)
            Case "UltraGrid"
                Dim activeGrid As UltraWinGrid.UltraGrid = ASCMAIN1.ActiveForm.CurrentControl
                activeGrid.Selected.Rows.AddRange(CType(activeGrid.Rows.All, UltraWinGrid.UltraGridRow()))
                If Not IsNothing(activeGrid.ActiveCell) Then
                    activeGrid.ActiveCell.Row.Activate()
                End If
            Case "UltraTextEditor"
                Dim activeTextEditor As Infragistics.Win.UltraWinEditors.UltraTextEditor = ASCMAIN1.ActiveForm.CurrentControl
                activeTextEditor.SelectAll()
        End Select

    End Sub

#End Region

    Private Sub ASFBASE0_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        '    Stop
        'If e.KeyChar = "=" Then
        '    If ASCMAIN1.Running_in_VS Then ASCMAIN1.MainForm_pgd.SelectedObject = Me.ActiveControl
        'End If

    End Sub

    Private Sub Form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If ASCMAIN1.USER_ID = "" Then
            Exit Sub
        End If

        IsLoading = True

        AddHandler Me.Disposed, AddressOf ASFMAIN1.ChildFormDisposed

        'Ftp1.RuntimeLicense = ASCMAIN1.TACMAIN1.nSoftwareftpkey

        Clear_dst()

        If ASCMAIN1.USER_ID = "" Then
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor

        ASCMAIN1.tblASTSQLX1 = Nothing
        Create_TDA(dst.Tables.Add, "ASTSQLX1", "*")
        ASCMAIN1.tblASTSQLX1 = dst.Tables("ASTSQLX1")

        SELECTION_NO = ASCMAIN1.Register_Form(Me)
        FORM_INSTANCE_NO = ASCMAIN1.Next_Control_No(MENU_ITEM_OBJECT & ".FORM_INSTANCE_NO")

        bind_to_TABLE_NAME = Absx1.GetABSBindToTable(Me)
        Load_Tool_Tips()
        Initialize_Controls_for_a_Container(Me)

        AddHandler tlb.ToolClick, AddressOf tlb_ToolClick
        AddHandler tlb.BeforeToolDropdown, AddressOf tlb_BeforeToolDropdown
        AddHandler tlb.ToolValueChanged, AddressOf tlb_ToolValueChanged

        TABLE_NAME = Absx1.GetABSTableName(Me)
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
        Absx1.TABLE_NAME_base = TABLE_NAME
        Absx1.Load_COLUMN_NAMEs()

        Load_Popup_Menus()

        Create_TDA(dst.Tables.Add, "ASTAUDT1", "*")
        Create_TDA(dst.Tables.Add, "ASTOPST2", "*")

        Set_Security_Context()

    End Sub

    Private Sub Form_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        IsLoading = False

        If ASCMAIN1.USER_ID = "" Then
            Exit Sub
        End If

        If MENU_ITEM_TYPE = "T" Then
            Try
                'why do we need tbl_schemas?
                ASCMAIN1.sql = "Select * from " & TABLE_NAME & " where ROWNUM <1"
                ASCMAIN1.sql = ASCDATA1.SQL_Server_Conversions(ASCMAIN1.sql)
                ASCMAIN1.oraCmd.CommandText = ASCMAIN1.sql
                With ASCMAIN1.oraCmd.ExecuteReader
                    tblASFBASE1_Schema = .GetSchemaTable
                    .Close()
                    .Dispose()
                End With
                tblASFBASE1_Schema.PrimaryKey = New DataColumn() {tblASFBASE1_Schema.Columns("ColumnName")}

            Catch ex As Exception

            End Try
        End If

        If MENU_ITEM_TYPE <> "T" Then
            If TABLE_NAME <> "" Then
                ' LET'S SEE IF THIS WORKS ALL OF THE TIME
                tblASFBASE1 = dst.Tables(TABLE_NAME)
            End If
        End If

        Call Populate_CMBs()

        If MENU_ITEM_TYPE <> "R" Then ' SINCE WE PUT ASTAUDT1 IN dst THERE IS CODE IN Bind_Controls WHICH THINKS THAT IF THERE IS ANYTHING IN dst, then ABSCOLUMN_NAMES SHOULD BE BINDABLE, AND THIS IS NOT TRUE FOR REPORTS
            Dim tf As Boolean = Absx1.GetABSBindToTable(Me)
            If tf Then ' WHY DO WE NEED TO BIND CONTROLS IF THIS CONTROL (OR FORM) STARTS OUT WITH BIND TO TABLE = FALSE
                'WITHOUT THE ABOVE LINE PMFNEXT2 BLOWS UP
                Call Bind_Controls(Me, TABLE_NAME, , tf, , True)
            End If
        End If

        If ASCMAIN1.Running_in_VS Then 'adding event handler for grid datasource tool tips in developer mode
            Add_Control_Tooltip_Events(Me)
        End If

        'Initialize_Controls_for_a_Container(Me) ' testing adding handlers after binding is complete

        ' IF THERE IS A PROBLEM WITH FORM LOAD, PLACE STOP HERE AND THEN ENABLE EXCEPTION CATCHING
        If MENU_ITEM_TYPE <> "T" And MENU_ITEM_TYPE <> "R" Then
            Call Mode_Settings(False)
        End If

        If MENU_ID = "XX" Then ' this form a derived from ASFBASE2 - check out its Constructor - probably need a boolean for this
        Else
            tdaASTOPST1 = ASCDATA1.GetDataAdapter(tblASTOPST1, "ASTOPST1", "*", True, -1, False)
            Call Create_ASTOPST1()
        End If

        Me.Cursor = Cursors.Default
    End Sub

    Overridable Sub Load_Popup_Menus()

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ctl">Control that displays the popup menu</param>
    ''' <param name="ToolTypes">Codes representing the menu types B=Button, S=State Button, O=Option Set, C=Color Picker, P=Separator, M=Begin Submenu, X=End Submenu</param>
    ''' <param name="Tools">The captions (which is also the key) to appear on the tool. Do NOT provide a caption for 'P' separators or for 'X' end of submenu. 
    '''  The Tool Captions should match the order of the ToolTypes</param>
    ''' <remarks></remarks>
    Sub Load_Popup_Menu(
    ByVal ctl As Control,
    ByVal ToolTypes As String,
    ByVal ParamArray Tools() As String)

        ASCMAIN1.Load_Popup_Menu(tlb, ctl, ToolTypes, Tools)
    End Sub

    Sub Create_ASTOPST1()
        rowASTOPST1 = tblASTOPST1.NewRow
        rowASTOPST1.Item("USER_ID") = ASCMAIN1.USER_ID
        rowASTOPST1.Item("MENU_ID") = MENU_ID
        rowASTOPST1.Item("MENU_ITEM_TYPE") = MENU_ITEM_TYPE
        rowASTOPST1.Item("MENU_ITEM_OBJECT") = MENU_ITEM_OBJECT
        rowASTOPST1.Item("SESSION_NO") = ASCMAIN1.SESSION_NO
        RE_XNO = RE_XNO + 1
        rowASTOPST1.Item("RE_XNO") = RE_XNO
        rowASTOPST1.Item("SELECTION_NO") = SELECTION_NO
        rowASTOPST1.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
        rowASTOPST1.Item("YYYYPP") = ASCMAIN1.CYP
        rowASTOPST1.Item("XNO") = ""
        rowASTOPST1.Item("PRD_CLOSE_IND") = ASCMAIN1.EOM
        rowASTOPST1.Item("FORM_INSTANCE_NO") = FORM_INSTANCE_NO
        rowASTOPST1.Item("VERSION_NO") = ASCMAIN1.VERSION_NO
        tblASTOPST1.Rows.Add(rowASTOPST1)
        tdaASTOPST1.Update(tblASTOPST1)

        If RE_XNO = 0 Then
            If dst.Tables.Contains("ASTSQLX1") Then
                For Each rowASTSQLX1 As DataRow In dst.Tables("ASTSQLX1").Select
                    rowASTSQLX1.Item("SELECTION_NO") = SELECTION_NO
                    rowASTSQLX1.Item("RE_XNO") = RE_XNO
                Next
            End If
        End If
    End Sub

    Sub Bind_Controls(ByVal c As Control)


        For Each cc As Control In c.Controls
            If cc.Controls.Count > 0 Then
                Call Bind_Controls(cc)
            End If
            Dim ABSColumnName As String = Absx1.GetABSColumnName(cc)
            Dim TABLE_NAME_control As String = Absx1.GetABSTableName(cc)

            If ABSColumnName <> "" Then
                Dim ABSTableName As String = Absx1.GetABSTableName(cc)
                Dim ABSBindToTable As Boolean = Absx1.GetABSBindToTable(cc)
                Dim rowASFBASE1 As DataRow = Nothing
                If ABSTableName <> "" And ABSBindToTable Then
                    rowASFBASE1 = TBL_SCHEMAs(ABSTableName).Rows.Find(ABSColumnName)
                Else
                    If tblASFBASE1_Schema.Rows.Count = 0 Then
                        rowASFBASE1 = Nothing
                    Else
                        rowASFBASE1 = tblASFBASE1_Schema.Rows.Find(ABSColumnName)
                    End If
                End If

                Dim Text_or_Value As String = "Text"

                If TypeOf cc Is UltraWinEditors.UltraTextEditor Then
                    Dim txtctl As UltraWinEditors.UltraTextEditor = DirectCast(cc, UltraWinEditors.UltraTextEditor)
                    Try
                        txtctl.MaxLength = dst.Tables(TABLE_NAME).Columns(ABSColumnName).MaxLength
                    Catch ex As Exception

                    End Try
                    Text_or_Value = "Value"
                End If

                If TypeOf cc Is UltraWinGrid.UltraCombo Then
                    Dim cmbctl As UltraWinGrid.UltraCombo = DirectCast(cc, UltraWinGrid.UltraCombo)
                    Try
                        cmbctl.MaxLength = dst.Tables(TABLE_NAME).Columns(ABSColumnName).MaxLength
                    Catch ex As Exception

                    End Try
                    Text_or_Value = "Value"
                End If

                If TypeOf cc Is UltraWinEditors.UltraComboEditor Then
                    Dim cbectl As UltraWinEditors.UltraComboEditor = DirectCast(cc, UltraWinEditors.UltraComboEditor)
                    Try
                        cbectl.MaxLength = tblASFBASE1.Columns(ABSColumnName).MaxLength
                    Catch ex As Exception

                    End Try
                    Text_or_Value = "Value"
                End If

                If TypeOf cc Is UltraWinEditors.UltraOptionSet Then
                    Text_or_Value = "Value"
                End If
                If TypeOf cc Is UltraWinEditors.UltraCheckEditor Then
                    Dim chkctl As UltraWinEditors.UltraCheckEditor = DirectCast(cc, UltraWinEditors.UltraCheckEditor)
                End If

                If TypeOf cc Is ABSCS.ABSCheckBox Then
                    Dim chkctl As ABSCS.ABSCheckBox = DirectCast(cc, ABSCS.ABSCheckBox)
                    Text_or_Value = "ABSChecked"
                End If

                If TypeOf cc Is UltraWinEditors.UltraDateTimeEditor Then
                    Text_or_Value = "Value"
                End If
                If TypeOf cc Is UltraWinSchedule.UltraCalendarCombo Then
                    Text_or_Value = "Value"
                End If

                If TypeOf cc Is UltraWinEditors.UltraNumericEditor Then
                    Text_or_Value = "Value"
                    Dim numctl As UltraWinEditors.UltraNumericEditor = DirectCast(cc, UltraWinEditors.UltraNumericEditor)
                    'numctl.MaskInput = ""
                    numctl.TabNavigation = UltraWinMaskedEdit.MaskedEditTabNavigation.NextControl
                    numctl.Nullable = True
                    If numctl.FormatString = "" Then
                        If Not rowASFBASE1 Is Nothing Then
                            Dim NumericPrecision As Integer = Val(rowASFBASE1.Item("NumericPrecision") & "")
                            Dim NumericScale As Integer = Val(rowASFBASE1.Item("NumericScale") & "")
                            Dim mask As String = "".PadLeft(NumericPrecision - NumericScale, "n")
                            mask = Replace(mask, "nnn", "nnn,")
                            Dim new_mask As String = ""
                            For i As Integer = 1 To Len(mask)
                                new_mask &= Mid(mask, Len(mask) - i + 1, 1)
                            Next
                            mask = new_mask

                            If NumericScale > 0 Then
                                mask = mask & "." & "".PadLeft(NumericScale, "n")
                                numctl.NumericType = UltraWinEditors.NumericType.Double
                            Else
                                numctl.NumericType = UltraWinEditors.NumericType.Integer
                            End If

                            If mask.Length > 0 AndAlso mask.Substring(0, 1) = "," Then
                                mask = mask.Substring(1)
                            End If

                            If numctl.MaxValue = 2147483647 Then
                                numctl.MaxValue = Val(Replace(Replace(mask, ",", ""), "n", "9"))
                            End If
                            If numctl.MinValue = -2147483648 Then
                                numctl.MinValue = -1 * Val(Replace(Replace(mask, ",", ""), "n", "9"))
                            End If

                            numctl.MaskInput = mask

                        End If
                    End If
                End If

                If dst.Tables.Count <> 0 And bind_to_TABLE_NAME Then
                    If (ABSTableName <> "" And ABSTableName <> TABLE_NAME) _
                    Or Not ABSBindToTable Then
                        ' do nothing
                    Else
                        If Not Bound_DataSources.Contains(dst.Tables(TABLE_NAME)) Then
                            Bound_DataSources.Add(dst.Tables(TABLE_NAME))
                        End If
                        Try
                            cc.DataBindings.Add(Text_or_Value, dst.Tables(TABLE_NAME), ABSColumnName)
                        Catch ex As Exception
                            MsgBox("Problem trying to bind " & ABSColumnName & " to DataSource (" & TABLE_NAME & ")", MsgBoxStyle.OkOnly, "Check properties of Controls")
                        End Try
                    End If
                End If
            End If
        Next
    End Sub

    Sub Bind_Controls(
    ByVal c As Control,
    ByVal TABLE_NAME As String,
    Optional ByVal ds As Object = Nothing,
    Optional ByVal bind_to_TABLE_NAME As Boolean = True,
    Optional ByVal DataMember As String = "",
    Optional ByVal default_table As Boolean = False)

        If ds Is Nothing Then
            ds = dst.Tables(TABLE_NAME)
        End If

        'If ds Is Nothing Then Exit Sub

        If Not Bound_DataSources.Contains(ds) Then
            Bound_DataSources.Add(ds)
        End If

        Dim ABSTableName_container As String = Absx1.GetABSTableName(c)
        If ABSTableName_container <> "" Then

        End If
        For Each cc As Control In c.Controls
            If cc.Controls.Count > 0 Then
                Call Bind_Controls(cc, TABLE_NAME, ds, bind_to_TABLE_NAME, DataMember, default_table)
            End If
            Dim ABSColumnName As String = Absx1.GetABSColumnName(cc)
            Dim ABSTableName As String = Absx1.GetABSTableName(cc)
            Dim ABSBindToTable As Boolean = Absx1.GetABSBindToTable(cc)
            Dim ABSParentColumnName As String = Absx1.GetABSParentColumnName(cc)
            If ABSTableName = "" And ABSTableName_container <> "" Then
                ABSTableName = ABSTableName_container
            End If
            'If ABSColumnName <> "" Then Stop
            ' TBL_SCHEMAs only has oracle definitions, so columns added to the DataTable are not available to routines here
            If ABSColumnName <> "" And (ABSTableName = TABLE_NAME Or ABSTableName = "" And default_table) Then
                Dim rowASFBASE1 As DataRow = Nothing

                'If (ABSTableName <> "" Or (ABSTableName = "" And default_table)) And ABSBindToTable And TBL_SCHEMAs.ContainsKey(IIf(ABSTableName = "", TABLE_NAME, ABSTableName)) Then
                '    rowASFBASE1 = TBL_SCHEMAs(IIf(ABSTableName = "", TABLE_NAME, ABSTableName)).Rows.Find(ABSColumnName)

                If ABSTableName <> "" And ABSBindToTable And TBL_SCHEMAs.ContainsKey(ABSTableName) Then
                    rowASFBASE1 = TBL_SCHEMAs(ABSTableName).Rows.Find(ABSColumnName)

                Else
                    If tblASFBASE1_Schema.Rows.Count = 0 Then
                        rowASFBASE1 = Nothing
                    Else
                        rowASFBASE1 = tblASFBASE1_Schema.Rows.Find(ABSColumnName)
                    End If
                End If

                Dim Text_or_Value As String = "Text"

                If TypeOf cc Is UltraWinEditors.UltraTextEditor Then
                    Dim txtctl As UltraWinEditors.UltraTextEditor = DirectCast(cc, UltraWinEditors.UltraTextEditor)

                    If dst.Tables.Contains(TABLE_NAME) Then
                        If dst.Tables(TABLE_NAME).Columns.Contains(ABSColumnName) Then
                            Try
                                Dim MAXLENGTH As Int32 = dst.Tables(TABLE_NAME).Columns(ABSColumnName).MaxLength
                                If MAXLENGTH > 0 Then
                                    txtctl.MaxLength = MAXLENGTH
                                End If
                            Catch ex As Exception

                            End Try
                        End If
                    End If

                    Text_or_Value = "Value"
                End If

                If TypeOf cc Is UltraWinGrid.UltraCombo Then
                    Dim cmbctl As UltraWinGrid.UltraCombo = DirectCast(cc, UltraWinGrid.UltraCombo)
                    If dst.Tables.Contains(TABLE_NAME) Then
                        If dst.Tables(TABLE_NAME).Columns.Contains(ABSColumnName) Then
                            Try
                                cmbctl.MaxLength = dst.Tables(TABLE_NAME).Columns(ABSColumnName).MaxLength
                            Catch ex As Exception

                            End Try
                        End If
                        Text_or_Value = "Value"
                    End If
                End If

                If TypeOf cc Is UltraWinEditors.UltraComboEditor Then
                    Dim cbectl As UltraWinEditors.UltraComboEditor = DirectCast(cc, UltraWinEditors.UltraComboEditor)
                    If dst.Tables.Contains(TABLE_NAME) AndAlso dst.Tables(TABLE_NAME).Columns.Contains(ABSColumnName) Then
                        If dst.Tables(TABLE_NAME).Columns(ABSColumnName).MaxLength > 0 Then
                            Try
                                cbectl.MaxLength = dst.Tables(TABLE_NAME).Columns(ABSColumnName).MaxLength
                            Catch ex As Exception

                            End Try
                        End If
                    End If
                    'If tblASFBASE1.Columns.Contains(ABSColumnName) Then
                    '    Try
                    '        cbectl.MaxLength = tblASFBASE1.Columns(ABSColumnName).MaxLength
                    '    Catch ex As Exception

                    '    End Try
                    'End If
                    Text_or_Value = "Value"
                End If

                If TypeOf cc Is UltraWinEditors.UltraOptionSet Then
                    Text_or_Value = "Value"
                End If
                If TypeOf cc Is UltraWinEditors.UltraCheckEditor Then
                    Dim chkctl As UltraWinEditors.UltraCheckEditor = DirectCast(cc, UltraWinEditors.UltraCheckEditor)
                End If

                If TypeOf cc Is ABSCS.ABSCheckBox Then
                    Dim chkctl As ABSCS.ABSCheckBox = DirectCast(cc, ABSCS.ABSCheckBox)
                    Text_or_Value = "ABSChecked"
                End If

                If TypeOf cc Is UltraWinEditors.UltraDateTimeEditor Then
                    Text_or_Value = "Value"
                End If
                If TypeOf cc Is UltraWinSchedule.UltraCalendarCombo Then
                    Text_or_Value = "Value"
                End If
                If TypeOf cc Is UltraWinEditors.UltraNumericEditor Then
                    Text_or_Value = "Value"
                    Dim numctl As UltraWinEditors.UltraNumericEditor = DirectCast(cc, UltraWinEditors.UltraNumericEditor)
                    'numctl.MaskInput = ""
                    numctl.TabNavigation = UltraWinMaskedEdit.MaskedEditTabNavigation.NextControl
                    numctl.Nullable = True

                    numctl.AlwaysInEditMode = False
                    numctl.MaskDisplayMode = UltraWinMaskedEdit.MaskMode.IncludeLiterals

                    If numctl.FormatString = "" And numctl.MaskInput = "" Then
                        If Not rowASFBASE1 Is Nothing Then
                            Dim NumericPrecision As Integer = Val(rowASFBASE1.Item("NumericPrecision") & "")
                            Dim NumericScale As Integer = Val(rowASFBASE1.Item("NumericScale") & "")
                            Dim DECIMAL_PLACES As Integer = NumericPrecision - NumericScale
                            If DECIMAL_PLACES < 0 Then DECIMAL_PLACES = 0
                            Dim mask As String = "".PadLeft(DECIMAL_PLACES, "n")
                            mask = Replace(mask, "nnn", "nnn,")
                            Dim new_mask As String = ""
                            For i As Integer = 1 To Len(mask)
                                new_mask &= Mid(mask, Len(mask) - i + 1, 1)
                            Next
                            mask = new_mask

                            If NumericScale > 0 Then
                                mask = mask & "." & "".PadLeft(NumericScale, "n")
                                numctl.NumericType = UltraWinEditors.NumericType.Double
                            Else
                                numctl.NumericType = UltraWinEditors.NumericType.Integer
                            End If

                            If mask.Length > 0 AndAlso mask.Substring(0, 1) = "," Then
                                mask = mask.Substring(1)
                            End If

                            If numctl.MaxValue = 2147483647 Then
                                numctl.MaxValue = Val(Replace(Replace(mask, ",", ""), "n", "9"))
                            End If
                            If numctl.MinValue = -2147483648 Then
                                numctl.MinValue = -1 * Val(Replace(Replace(mask, ",", ""), "n", "9"))
                            End If

                            'numctl.MaskInput = mask ' why do this only for those columns that exist in the oracle table
                            mask = Replace(mask, "n." & "".PadLeft(NumericScale, "n"), "0." & "".PadLeft(NumericScale, "0"))
                            mask = Replace(mask, "n", "#")
                            If mask <> "" Then If Mid(mask, Len(mask), 1) = "#" Then Mid(mask, Len(mask), 1) = "0"
                            numctl.FormatString = mask
                        Else
                            If numctl.NumericType = UltraWinEditors.NumericType.Integer Then
                                'numctl.MaskInput = "nnn,nnn,nnn,nnn"
                                'numctl.FormatString = "###,###,###,##0"
                                numctl.FormatString = "#,##0"
                            Else
                                'numctl.MaskInput = "nnn,nnn,nnn,nnn.nn"
                                'numctl.FormatString = "###,###,###,##0.00"
                                numctl.FormatString = "#,##0.00"
                            End If

                        End If
                    End If
                End If

                If bind_to_TABLE_NAME And ABSBindToTable Then
                    If ds IsNot Nothing Then
                        Try
                            cc.DataBindings.Clear()
                            If DataMember = "" Then
                                If ds.GetType.ToString = "System.Data.DataTable" Then
                                    Dim tbl As DataTable = DirectCast(ds, DataTable)
                                    If tbl.Columns.Contains(ABSColumnName) Then
                                        cc.DataBindings.Add(Text_or_Value, ds, ABSColumnName)
                                    Else
                                        If ABSParentColumnName = "" Then
                                            If ASCMAIN1.Running_in_VS Then
                                                MsgBox("Problem trying to bind " & ABSColumnName & " to DataSource (" & TABLE_NAME & ")", MsgBoxStyle.OkOnly, "Check properties of Controls")
                                            End If
                                        End If
                                    End If
                                ElseIf ds.GetType.ToString = "System.Data.DataView" Then
                                    Dim tbl As DataTable = DirectCast(ds, DataView).ToTable
                                    If tbl.Columns.Contains(ABSColumnName) Then
                                        cc.DataBindings.Add(Text_or_Value, ds, ABSColumnName)
                                    Else
                                        If ABSParentColumnName = "" Then
                                            If ASCMAIN1.Running_in_VS Then
                                                MsgBox("Problem trying to bind " & ABSColumnName & " to DataSource (" & TABLE_NAME & ")", MsgBoxStyle.OkOnly, "Check properties of Controls")
                                            End If
                                        End If
                                    End If
                                Else
                                    MsgBox("Error trying to bind " & ABSColumnName, MsgBoxStyle.OkOnly, "Please Call ABS")

                                    Stop ' don't know what to do here
                                End If
                            Else
                                cc.DataBindings.Add(Text_or_Value, ds, DataMember & "." & ABSColumnName)
                            End If
                        Catch ex As Exception
                            If ABSParentColumnName = "" Then
                                If ASCMAIN1.Running_in_VS Then
                                    MsgBox("Problem trying to bind " & ABSColumnName & " to DataSource (" & TABLE_NAME & ")", MsgBoxStyle.OkOnly, "Check properties of Controls")
                                End If
                            End If
                        End Try
                    End If
                End If
            End If
        Next
    End Sub

#Region "Developer Mode Control Datasource Tooltip Logic"

    Sub Add_Control_Tooltip_Events(ByVal c As Control)

        For Each cc As Control In c.Controls

            If cc.Controls.Count > 0 Then
                Add_Control_Tooltip_Events(cc)
            End If

            If TypeOf cc Is UltraWinGrid.UltraGrid Then
                Dim grdctl As UltraWinGrid.UltraGrid = DirectCast(cc, UltraWinGrid.UltraGrid)
                AddHandler grdctl.MouseMove, AddressOf Grid_DataSource_Tooltip
                AddHandler grdctl.MouseEnterElement, AddressOf Grid_Column_Header_Tooltip
            End If

            If TypeOf cc Is UltraWinEditors.UltraTextEditor Then
                Dim txtctl As UltraWinEditors.UltraTextEditor = DirectCast(cc, UltraWinEditors.UltraTextEditor)
                AddHandler txtctl.MouseEnter, AddressOf Ctl_Show_Developer_Tooltip
                AddHandler txtctl.MouseLeave, AddressOf Ctl_Hide_Developer_Tooltip
            End If

            If TypeOf cc Is UltraWinEditors.UltraOptionSet Then
                Dim optctl As UltraWinEditors.UltraOptionSet = DirectCast(cc, UltraWinEditors.UltraOptionSet)
                AddHandler optctl.MouseEnter, AddressOf Ctl_Show_Developer_Tooltip
                AddHandler optctl.MouseLeave, AddressOf Ctl_Hide_Developer_Tooltip
            End If

            If TypeOf cc Is UltraWinEditors.UltraDateTimeEditor Then
                Dim dtectl As UltraWinEditors.UltraDateTimeEditor = DirectCast(cc, UltraWinEditors.UltraDateTimeEditor)
                AddHandler dtectl.MouseEnter, AddressOf Ctl_Show_Developer_Tooltip
                AddHandler dtectl.MouseLeave, AddressOf Ctl_Hide_Developer_Tooltip
            End If

            If TypeOf cc Is UltraWinEditors.UltraNumericEditor Then
                Dim numctl As UltraWinEditors.UltraNumericEditor = DirectCast(cc, UltraWinEditors.UltraNumericEditor)
                AddHandler numctl.MouseEnter, AddressOf Ctl_Show_Developer_Tooltip
                AddHandler numctl.MouseLeave, AddressOf Ctl_Hide_Developer_Tooltip
            End If

            If TypeOf cc Is UltraWinMaskedEdit.UltraMaskedEdit Then
                Dim mskctl As UltraWinMaskedEdit.UltraMaskedEdit = DirectCast(cc, UltraWinMaskedEdit.UltraMaskedEdit)
                AddHandler mskctl.MouseEnter, AddressOf Ctl_Show_Developer_Tooltip
                AddHandler mskctl.MouseLeave, AddressOf Ctl_Hide_Developer_Tooltip
            End If

            If TypeOf cc Is UltraWinEditors.UltraCheckEditor Then
                Dim chkctl As UltraWinEditors.UltraCheckEditor = DirectCast(cc, UltraWinEditors.UltraCheckEditor)
                AddHandler chkctl.MouseEnter, AddressOf Ctl_Show_Developer_Tooltip
                AddHandler chkctl.MouseLeave, AddressOf Ctl_Hide_Developer_Tooltip
            End If
            If TypeOf cc Is UltraWinEditors.UltraComboEditor Then
                Dim cbectl As UltraWinEditors.UltraComboEditor = DirectCast(cc, UltraWinEditors.UltraComboEditor)
                AddHandler cbectl.MouseEnter, AddressOf Ctl_Show_Developer_Tooltip
                AddHandler cbectl.MouseLeave, AddressOf Ctl_Hide_Developer_Tooltip
            End If

            If TypeOf cc Is ABSCS.ABSCheckBox Then
                Dim chkctl As ABSCS.ABSCheckBox = DirectCast(cc, ABSCS.ABSCheckBox)
                AddHandler chkctl.MouseEnter, AddressOf Ctl_Show_Developer_Tooltip
                AddHandler chkctl.MouseLeave, AddressOf Ctl_Hide_Developer_Tooltip
            End If
        Next

    End Sub

    Sub Ctl_Hide_Developer_Tooltip(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim needToHide As Boolean = tip.IsToolTipVisible(sender)
        If needToHide Then
            tip.HideToolTip()
        End If
    End Sub

    Sub Ctl_Show_Developer_Tooltip(ByVal sender As Object, ByVal e As System.EventArgs)

        tip.Enabled = False

        If Not ASCMAIN1.developerModeOptions.DataSourceToolTip Then Exit Sub

        If My.Computer.Keyboard.CtrlKeyDown Then
            tip.Enabled = True
            Dim ctl As Control = DirectCast(sender, Control)
            ctl.Focus()

            Dim tipText As String = ""
            Dim parentTable As String = ""
            Dim ctlName As String = ctl.Name
            Dim ABSColumnName As String = Absx1.GetABSColumnName(sender) & ""
            Dim ABSTableName As String = Absx1.GetABSTableName(sender) & ""
            Dim ABSBindToTable As Boolean = Absx1.GetABSBindToTable(sender) & ""
            Dim ABSParentColumnName As String = Absx1.GetABSParentColumnName(sender) & ""
            Dim ABSViewName As String = Absx1.GetABSViewName(sender) & ""

            Dim dtbl As String = ""
            If ABSTableName = "" Then
                dtbl = TABLE_NAME
            Else
                dtbl = ABSTableName
            End If

            tipText = "<b> Field </b><br/>" & IIf(ABSColumnName = "", "None", ABSColumnName)
            If ABSBindToTable Then
                tipText = tipText & "<hr NoShade='true' size='2px' style='color:#000000;' color='#000000'/><b> Table </b>"
                tipText = tipText & "<br/><span align='center'>" & IIf(dtbl = "", "Unknown", dtbl) & "</span>"
            End If
            tipText = tipText & "<hr NoShade='true' size='2px' style='color:#000000;' color='#000000'/><b> ABS Bind </b>"
            tipText = tipText & "<br/><span align='center'>" & ABSBindToTable.ToString & "</span>"
            If ABSParentColumnName <> "" Then
                tipText = tipText & "<hr NoShade='true' size='2px' style='color:#000000;' color='#000000'/><b> ABS Parent </b>"
                tipText = tipText & "<br/><span align='center'>" & ABSParentColumnName & "</span>"
            End If
            If ABSViewName <> "" Then
                tipText = tipText & "<hr NoShade='true' size='2px' style='color:#000000;' color='#000000'/><b> ABS View </b>"
                tipText = tipText & "<br/><span align='center'>" & ABSViewName & "</span>"
            End If

            If TypeOf sender Is UltraWinEditors.UltraNumericEditor Then
                Dim numctl As UltraWinEditors.UltraNumericEditor = DirectCast(sender, UltraWinEditors.UltraNumericEditor)
                If numctl.MaskInput <> "" Then
                    tipText = tipText & "<hr NoShade='true' size='2px' style='color:#000000;' color='#000000'/><b> Mask </b>"
                    tipText = tipText & "<br/><span align='center'>" & numctl.MaskInput & "</span>"
                End If
                If numctl.FormatString <> "" Then
                    tipText = tipText & "<hr NoShade='true' size='2px' style='color:#000000;' color='#000000'/><b> Format </b>"
                    tipText = tipText & "<br/><span align='center'>" & numctl.FormatString & "</span>"
                End If
            End If

            If TypeOf sender Is UltraWinEditors.UltraTextEditor Then
                Dim txtctl As UltraWinEditors.UltraTextEditor = DirectCast(sender, UltraWinEditors.UltraTextEditor)
                If dst.Tables.Contains(dtbl) Then
                    If dst.Tables(dtbl).Columns.Contains(ABSColumnName) Then
                        Try
                            Dim MAXLENGTH As Int32 = dst.Tables(dtbl).Columns(ABSColumnName).MaxLength
                            tipText = tipText & "<hr NoShade='true' size='2px' style='color:#000000;' color='#000000'/><b> Max Length </b>"
                            tipText = tipText & "<br/><span align='center'>" & MAXLENGTH.ToString & "</span>"
                        Catch ex As Exception

                        End Try
                    End If
                End If
            End If

            If TypeOf sender Is UltraWinEditors.UltraComboEditor Then
                Dim cbectl As UltraWinEditors.UltraComboEditor = DirectCast(sender, UltraWinEditors.UltraComboEditor)
                If dst.Tables.Contains(dtbl) Then
                    If dst.Tables(dtbl).Columns.Contains(ABSColumnName) Then
                        Try
                            Dim MAXLENGTH As Int32 = dst.Tables(dtbl).Columns(ABSColumnName).MaxLength
                            tipText = tipText & "<hr NoShade='true' size='2px' style='color:#000000;' color='#000000'/><b> Max Length </b>"
                            tipText = tipText & "<br/><span align='center'>" & MAXLENGTH.ToString & "</span>"
                        Catch ex As Exception

                        End Try
                    End If
                End If
            End If

            If TypeOf sender Is UltraWinMaskedEdit.UltraMaskedEdit Then
                Dim mskctl As UltraWinMaskedEdit.UltraMaskedEdit = DirectCast(sender, UltraWinMaskedEdit.UltraMaskedEdit)
                If mskctl.InputMask <> "" Then
                    tipText = tipText & "<hr NoShade='true' size='2px' style='color:#000000;' color='#000000'/><b> Input Mask </b>"
                    tipText = tipText & "<br/><span align='center'>" & System.Web.HttpUtility.HtmlEncode(mskctl.InputMask) & "</span>"
                End If
            End If

            tipText = tipText & "<hr NoShade='true' size='2px' style='color:#000000;' color='#000000'/><b> Control Name </b>"
            tipText = tipText & "<br/><span align='center'>" & ctlName & "</span>"

            Dim tipInfo As New UltraWinToolTip.UltraToolTipInfo()
            tipInfo.ToolTipTextStyle = ToolTipTextStyle.Formatted
            tipInfo.ToolTipTextFormatted = tipText

            tip.DisplayStyle = ToolTipDisplayStyle.Office2007
            tip.InitialDelay = 0
            tip.AutoPopDelay = 5000
            tip.SetUltraToolTip(sender, tipInfo)
            If Not tip.IsToolTipVisible(sender) Then
                tip.ShowToolTip(sender)
            End If
        Else
            tip.HideToolTip()
        End If

        Return

    End Sub

    Sub Grid_Column_Header_Tooltip(ByVal sender As Object, ByVal e As Infragistics.Win.UIElementEventArgs)

        If Not ASCMAIN1.developerModeOptions.DataSourceToolTip Then
            tip.SetUltraToolTip(sender, Nothing)
            tip.HideToolTip()
            Exit Sub
        End If

        If Not My.Computer.Keyboard.CtrlKeyDown Then
            tip.SetUltraToolTip(sender, Nothing)
            tip.HideToolTip()
            Exit Sub
        End If

        Dim grid As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        Dim columnHeaderUIElement As UltraWinGrid.ColumnHeader = e.Element.GetContext(GetType(Infragistics.Win.UltraWinGrid.ColumnHeader))

        If Not columnHeaderUIElement Is Nothing Then
            Dim tipText As String = ""

            Dim colKey As String = columnHeaderUIElement.Column.Key.ToString
            tipText = "<b> Column Key </b><br/>" & colKey

            Dim colDataType As String = columnHeaderUIElement.Column.DataType.ToString
            tipText = tipText & "<hr NoShade='true' size='2px' style='color:#000000;' color='#000000'/><b> Data Type </b>"
            tipText = tipText & "<br/><span align='center'>" & colDataType & "</span>"

            Dim colIsBound As String = columnHeaderUIElement.Column.IsBound.ToString
            tipText = tipText & "<hr NoShade='true' size='2px' style='color:#000000;' color='#000000'/><b> Bound </b>"
            tipText = tipText & "<br/><span align='center'>" & colIsBound & "</span>"

            Dim colwidth As String = columnHeaderUIElement.Column.Width.ToString
            tipText = tipText & "<hr NoShade='true' size='2px' style='color:#000000;' color='#000000'/><b> Width </b>"
            tipText = tipText & "<br/><span align='center'>" & colwidth & "</span>"

            Dim tipInfo As New UltraWinToolTip.UltraToolTipInfo()
            tipInfo.ToolTipTextStyle = ToolTipTextStyle.Formatted
            tipInfo.ToolTipTextFormatted = tipText

            tip.DisplayStyle = ToolTipDisplayStyle.Office2007
            tip.InitialDelay = 0
            tip.AutoPopDelay = 5000
            tip.SetUltraToolTip(grid, tipInfo)
            tip.Enabled = True
            tip.ShowToolTip(grid)
            Return
        Else
            Dim element As UIElement = grid.DisplayLayout.UIElement.LastElementEntered
            Dim captionAreaUIElement As UltraWinGrid.CaptionAreaUIElement = TryCast(element, UltraWinGrid.CaptionAreaUIElement)
            If captionAreaUIElement IsNot Nothing Then
                Exit Sub
            Else
                tip.Enabled = False
                tip.HideToolTip()
            End If

        End If

    End Sub

    Sub Grid_DataSource_Tooltip(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)

        If Not ASCMAIN1.developerModeOptions.DataSourceToolTip Then
            tip.SetUltraToolTip(sender, Nothing)
            tip.HideToolTip()
            Exit Sub
        End If

        If Not My.Computer.Keyboard.CtrlKeyDown Then
            tip.SetUltraToolTip(sender, Nothing)
            tip.HideToolTip()
            Exit Sub
        End If

        Dim grid As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        Dim element As UIElement = grid.DisplayLayout.UIElement.LastElementEntered
        Dim elementPoint As UIElement = grid.DisplayLayout.UIElement.ElementFromPoint(New Point(e.X, e.Y))

        Dim ttv As Boolean = tip.IsToolTipVisible(grid)

        If element IsNot Nothing And Not ttv Then
            Dim captionAreaUIElement As UltraWinGrid.CaptionAreaUIElement = TryCast(element, UltraWinGrid.CaptionAreaUIElement)
            Dim captionAreaUIElementPoint As UltraWinGrid.CaptionAreaUIElement = TryCast(elementPoint, UltraWinGrid.CaptionAreaUIElement)
            Dim columnHeaderUIElement As UltraWinGrid.ColumnHeader = TryCast(elementPoint.GetContext(GetType(Infragistics.Win.UltraWinGrid.ColumnHeader)), UltraWinGrid.ColumnHeader)

            If columnHeaderUIElement IsNot Nothing Then
                Exit Sub
            End If
            Dim tipText As String = ""
            Dim showTT As Boolean = True
            If grid.Text <> "" Then
                showTT = (e.Y <= 24)
            End If
            If captionAreaUIElement IsNot Nothing Or showTT Then

                Dim tblName As String = ""
                Dim tbl As DataTable = Nothing
                Dim dvwRowFilter As String = ""
                Dim dsType As String = "Source"

                If grid.DataSource IsNot Nothing Then
                    tblName = grid.DataSource.ToString
                    If tblName = "System.Data.DataView" Then
                        dsType = "View"
                        Dim dvw As DataView = grid.DataSource
                        tblName = dvw.Table.TableName
                        dvwRowFilter = dvw.RowFilter
                    End If
                    tbl = dst.Tables(tblName)
                Else
                    tblName = "None"
                End If

                'Datasource
                tipText = "<b> Data " & dsType & " </b><br/>" & tblName

                If tbl IsNot Nothing Then
                    'Primary Key
                    Dim columns() As DataColumn
                    columns = tbl.PrimaryKey
                    If columns.Count > 0 Then
                        tipText = tipText & "<hr NoShade='true' size='2px' style='color:#000000;' color='#000000'/><b> Primary Key </b>"
                        Dim i As Integer
                        For i = 0 To columns.GetUpperBound(0)
                            tipText = tipText & "<br/><span align='center'>" & columns(i).ColumnName & "&nbsp;&nbsp;&nbsp;" & columns(i).DataType.ToString() & "</span>"
                        Next i

                    End If

                    'Relationships
                    Dim r2 As Boolean = False
                    Dim relation As DataRelation
                    For Each relation In tbl.ChildRelations
                        If Not r2 Then
                            tipText = tipText & "<hr NoShade='true' size='2px' style='color:#000000;' color='#000000'/><b> Relationships </b>"
                            r2 = True
                        End If
                        tipText = tipText & "<br/>" & relation.RelationName
                    Next

                    'Dataview Rowfilter
                    If dvwRowFilter <> "" Then
                        tipText = tipText & "<hr NoShade='true' size='2px' style='color:#000000;' color='#000000'/><b> Rowfilter </b>"
                        tipText = tipText & "<br/>" & System.Web.HttpUtility.HtmlEncode(dvwRowFilter)
                    End If

                End If

                Dim tipInfo As New UltraWinToolTip.UltraToolTipInfo()
                tipInfo.ToolTipTextStyle = ToolTipTextStyle.Formatted
                tipInfo.ToolTipTextFormatted = tipText

                tip.DisplayStyle = ToolTipDisplayStyle.Office2007
                tip.InitialDelay = 0
                tip.AutoPopDelay = 5000
                tip.SetUltraToolTip(grid, tipInfo)
                tip.Enabled = True
                tip.ShowToolTip(grid)
                Return
            Else
                tip.Enabled = False
                tip.HideToolTip()
            End If
        End If


    End Sub

#End Region

    Sub Bind_Controls_OLD(
    ByVal c As Control,
    ByVal TABLE_NAME As String,
    Optional ByVal ds As Object = Nothing,
    Optional ByVal bind_to_TABLE_NAME As Boolean = True,
    Optional ByVal DataMember As String = "",
    Optional ByVal default_table As Boolean = False)

        If ds Is Nothing Then
            ds = dst.Tables(TABLE_NAME)
        End If

        'If ds Is Nothing Then Exit Sub

        If Not Bound_DataSources.Contains(ds) Then
            Bound_DataSources.Add(ds)
        End If

        For Each cc As Control In c.Controls
            If cc.Controls.Count > 0 Then
                Call Bind_Controls(cc, TABLE_NAME, ds, bind_to_TABLE_NAME, DataMember, default_table)
            End If
            Dim ABSColumnName As String = Absx1.GetABSColumnName(cc)
            Dim ABSTableName As String = Absx1.GetABSTableName(cc)
            Dim ABSBindToTable As Boolean = Absx1.GetABSBindToTable(cc)
            Dim ABSParentColumnName As String = Absx1.GetABSParentColumnName(cc)

            'If ABSColumnName = "ITEM_CODE" Then Stop
            Dim TABLE_NAME_schema As String = ABSTableName
            If ABSTableName = "" And default_table Then
                TABLE_NAME_schema = TABLE_NAME
            End If

            If ABSColumnName <> "" And TBL_SCHEMAs.ContainsKey(TABLE_NAME_schema) And (ABSTableName = TABLE_NAME Or ABSTableName = "" And default_table) Then
                Dim rowASFBASE1 As DataRow = Nothing

                'If ABSTableName <> "" And ABSBindToTable Then
                '    rowASFBASE1 = TBL_SCHEMAs(ABSTableName).Rows.Find(ABSColumnName)

                If TABLE_NAME_schema <> "" And ABSBindToTable Then
                    rowASFBASE1 = TBL_SCHEMAs(TABLE_NAME_schema).Rows.Find(ABSColumnName)
                Else
                    If tblASFBASE1_Schema.Rows.Count = 0 Then
                        rowASFBASE1 = Nothing
                    Else
                        rowASFBASE1 = tblASFBASE1_Schema.Rows.Find(ABSColumnName)
                    End If
                End If

                Dim Text_or_Value As String = "Text"

                If TypeOf cc Is UltraWinEditors.UltraTextEditor Then
                    Dim txtctl As UltraWinEditors.UltraTextEditor = DirectCast(cc, UltraWinEditors.UltraTextEditor)

                    If dst.Tables.Contains(TABLE_NAME) Then
                        If dst.Tables(TABLE_NAME).Columns.Contains(ABSColumnName) Then
                            Try
                                txtctl.MaxLength = dst.Tables(TABLE_NAME).Columns(ABSColumnName).MaxLength
                            Catch ex As Exception

                            End Try
                        End If
                    End If

                    Text_or_Value = "Value"
                End If

                If TypeOf cc Is UltraWinGrid.UltraCombo Then
                    Dim cmbctl As UltraWinGrid.UltraCombo = DirectCast(cc, UltraWinGrid.UltraCombo)
                    If dst.Tables.Contains(TABLE_NAME) Then
                        If dst.Tables(TABLE_NAME).Columns.Contains(ABSColumnName) Then
                            Try
                                cmbctl.MaxLength = dst.Tables(TABLE_NAME).Columns(ABSColumnName).MaxLength
                            Catch ex As Exception

                            End Try
                        End If
                        Text_or_Value = "Value"
                    End If
                End If

                If TypeOf cc Is UltraWinEditors.UltraComboEditor Then
                    Dim cbectl As UltraWinEditors.UltraComboEditor = DirectCast(cc, UltraWinEditors.UltraComboEditor)
                    If dst.Tables.Contains(TABLE_NAME) AndAlso dst.Tables(TABLE_NAME).Columns.Contains(ABSColumnName) Then
                        Try
                            cbectl.MaxLength = dst.Tables(TABLE_NAME).Columns(ABSColumnName).MaxLength
                        Catch ex As Exception

                        End Try
                    End If
                    'If tblASFBASE1.Columns.Contains(ABSColumnName) Then
                    '    Try
                    '        cbectl.MaxLength = tblASFBASE1.Columns(ABSColumnName).MaxLength
                    '    Catch ex As Exception

                    '    End Try
                    'End If
                    Text_or_Value = "Value"
                End If

                If TypeOf cc Is UltraWinEditors.UltraOptionSet Then
                    Text_or_Value = "Value"
                End If
                If TypeOf cc Is UltraWinEditors.UltraCheckEditor Then
                    Dim chkctl As UltraWinEditors.UltraCheckEditor = DirectCast(cc, UltraWinEditors.UltraCheckEditor)
                End If

                If TypeOf cc Is ABSCS.ABSCheckBox Then
                    Dim chkctl As ABSCS.ABSCheckBox = DirectCast(cc, ABSCS.ABSCheckBox)
                    Text_or_Value = "ABSChecked"
                End If

                If TypeOf cc Is UltraWinEditors.UltraDateTimeEditor Then
                    Text_or_Value = "Value"
                End If
                If TypeOf cc Is UltraWinEditors.UltraNumericEditor Then
                    Text_or_Value = "Value"
                    Dim numctl As UltraWinEditors.UltraNumericEditor = DirectCast(cc, UltraWinEditors.UltraNumericEditor)
                    'numctl.MaskInput = ""
                    numctl.TabNavigation = UltraWinMaskedEdit.MaskedEditTabNavigation.NextControl
                    numctl.Nullable = True
                    If numctl.FormatString = "" Then
                        If Not rowASFBASE1 Is Nothing Then
                            Dim NumericPrecision As Integer = Val(rowASFBASE1.Item("NumericPrecision") & "")
                            Dim NumericScale As Integer = Val(rowASFBASE1.Item("NumericScale") & "")
                            Dim mask As String = "".PadLeft(NumericPrecision - NumericScale, "n")
                            mask = Replace(mask, "nnn", "nnn,")
                            Dim new_mask As String = ""
                            For i As Integer = 1 To Len(mask)
                                new_mask &= Mid(mask, Len(mask) - i + 1, 1)
                            Next
                            mask = new_mask

                            If NumericScale > 0 Then
                                mask = mask & "." & "".PadLeft(NumericScale, "n")
                                numctl.NumericType = UltraWinEditors.NumericType.Double
                            Else
                                numctl.NumericType = UltraWinEditors.NumericType.Integer
                            End If

                            If mask.Length > 0 AndAlso mask.Substring(0, 1) = "," Then
                                mask = mask.Substring(1)
                            End If

                            If numctl.MaxValue = 2147483647 Then
                                numctl.MaxValue = Val(Replace(Replace(mask, ",", ""), "n", "9"))
                            End If
                            If numctl.MinValue = -2147483648 Then
                                numctl.MinValue = -1 * Val(Replace(Replace(mask, ",", ""), "n", "9"))
                            End If

                            numctl.MaskInput = mask

                        End If
                    End If
                End If

                If bind_to_TABLE_NAME And ABSBindToTable Then
                    If ds IsNot Nothing Then
                        Try
                            cc.DataBindings.Clear()
                            If DataMember = "" Then
                                If ds.GetType.ToString = "System.Data.DataTable" Then
                                    Dim tbl As DataTable = DirectCast(ds, DataTable)
                                    If tbl.Columns.Contains(ABSColumnName) Then
                                        cc.DataBindings.Add(Text_or_Value, ds, ABSColumnName)
                                    Else
                                        If ABSParentColumnName = "" Then
                                            If ASCMAIN1.Running_in_VS Then
                                                MsgBox("Problem trying to bind " & ABSColumnName & " to DataSource (" & TABLE_NAME & ")", MsgBoxStyle.OkOnly, "Check properties of Controls")
                                            End If
                                        End If
                                    End If
                                Else
                                    MsgBox("Error trying to bind " & ABSColumnName, MsgBoxStyle.OkOnly, "Please Call ABS")

                                    Stop ' don't know what to do here
                                End If
                            Else
                                cc.DataBindings.Add(Text_or_Value, ds, DataMember & "." & ABSColumnName)
                            End If
                        Catch ex As Exception
                            If ABSParentColumnName = "" Then
                                If ASCMAIN1.Running_in_VS Then
                                    MsgBox("Problem trying to bind " & ABSColumnName & " to DataSource (" & TABLE_NAME & ")", MsgBoxStyle.OkOnly, "Check properties of Controls")
                                End If
                            End If
                        End Try
                    End If
                End If
            End If
        Next
    End Sub

    Public Overridable Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

    End Sub

    Sub Record_Exception(ByVal ex2record As Exception)
        Try
            Dim tbl As New DataTable("EXCEPTION")
            dst.Tables.Add(tbl)
            tbl.Columns.Add("EXCEPTION", GetType(System.Exception))
            Dim FILE_NAME As String = "Exception_" & Format(Now, "yyyyMMddHHmmss")
            Dim FOLDER_NAME As String = Application.StartupPath
            Call Write_DataSet(True, FOLDER_NAME, FILE_NAME)
        Catch ex As Exception

        End Try
    End Sub

    Sub Write_DataSet(
    Optional ByVal xml As Boolean = False,
    Optional ByVal FOLDER_NAME As String = "",
    Optional ByVal FILE_NAME As String = "")

        clsASCBASE1.Write_DataSet(xml, FOLDER_NAME, FILE_NAME)

    End Sub

    Sub Populate_CMBs()

        For Each COLUMN_NAME As String In CMBs.Keys

            Dim CTL As Control = CMBs(COLUMN_NAME)
            Dim cmbctl As UltraWinGrid.UltraCombo = DirectCast(CTL, UltraWinGrid.UltraCombo)
            If cmbctl.DisplayLayout.Bands(0).Columns.Count <> 0 Then
                ' application is handling populating of cmb in form_load, like GLFJRNL1
            Else
                VIEW_NAME = Absx1.GetABSViewName(CTL)
                If VIEW_NAME = "" Then
                    VIEW_NAME = COLUMN_NAME
                End If
                TABLE_NAME_view = Absx1.GetABSLookUpTableName(CTL)
                If cmbctl.ReadOnly Then
                    ' PROBABLY NEED TABLE_NAME. IN FRONT OF COLUMN_NAME
                    'ABSReadOnly.Add(COLUMN_NAME)
                    ABSReadOnly.Add(cmbctl.Name)
                End If
                If VIEW_NAME <> "" Then

                    cmbYPparm = New cmbYPparms
                    If cmbYP.ContainsKey(COLUMN_NAME) Then
                        cmbYPparm = cmbYP(COLUMN_NAME)
                    End If

                    cmbYWparm = New cmbYWparms
                    If cmbYW.ContainsKey(COLUMN_NAME) Then
                        cmbYWparm = cmbYW(COLUMN_NAME)
                    End If

                    Dim sql As String
                    If Mid(Me.Name, 3, 1) = "R" Then
                        ' Stop ' NEED TO DO SOMETHING FOR WEEKS HERE
                        Dim sql_where As String = ""
                        If cmbYPparm.Base_YYYYPP <> "" Then
                            sql_where = sql_where & "OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(cmbYPparm.Base_YYYYPP, cmbYPparm.RelativeStartingPeriod) & "' and OPS_YYYYPP <= '" & ASCMAIN1.Period_Calc(cmbYPparm.Base_YYYYPP, cmbYPparm.RelativeEndingPeriod) & "'"
                        End If
                        If cmbYPparm.Parent_cmbYP <> "" Then
                            Dim Base_cmbYPparm As New cmbYPparms
                            Base_cmbYPparm = cmbYP(cmbYPparm.Parent_cmbYP)
                            ' this sql allows for x periods in the future, but does not provide for a cap on the latest period available
                            ' not really important for the app that drew me here (APRMCHK1), so waiting for a compelling app before changing it
                            sql_where = sql_where & "OPS_YYYYPP >= '" & Base_cmbYPparm.Base_YYYYPP & "' and OPS_YYYYPP <= '" & ASCMAIN1.Period_Calc(Base_cmbYPparm.Base_YYYYPP, cmbYPparm.TotalRelativePeriods - 1) & "'"
                        End If


                        If cmbYWparm.Base_YYYYWW <> "" Then
                            sql_where = sql_where & "YYYYWW >= '" & ASCMAIN1.Week_Calc(cmbYWparm.Base_YYYYWW, cmbYWparm.RelativeStartingWeek) & "' and YYYYWW <= '" & ASCMAIN1.Week_Calc(cmbYWparm.Base_YYYYWW, cmbYWparm.RelativeEndingWeek) & "'"
                        End If
                        If cmbYWparm.Parent_cmbYW <> "" Then
                            Dim Base_cmbYWparm As New cmbYWparms
                            Base_cmbYWparm = cmbYW(cmbYWparm.Parent_cmbYW)
                            ' this sql allows for x periods in the future, but does not provide for a cap on the latest period available
                            ' not really important for the app that drew me here (APRMCHK1), so waiting for a compelling app before changing it
                            sql_where = sql_where & "YYYYWW >= '" & Base_cmbYWparm.Base_YYYYWW & "' and YYYYWW <= '" & ASCMAIN1.Week_Calc(Base_cmbYWparm.Base_YYYYWW, cmbYWparm.TotalRelativeWeeks - 1) & "'"
                        End If

                        Prepare_for_View_Lookup_Special(CTL, COLUMN_NAME, sql_where)
                        sql = ASCMAIN1.CodeSelector.Get_SQL(VIEW_NAME, TABLE_NAME_view, sql_where) ' GLRTBAL1
                    Else

                        If ASCMAIN1.ActiveForm.MENU_ITEM_TYPE = "T" Then
                            sql = ASCMAIN1.CodeSelector.Get_SQL(VIEW_NAME, TABLE_NAME_view) ' ASTDSQLF
                        Else
                            Dim sql_where As String = ""
                            If cmbYPparm.Base_YYYYPP <> "" Then
                                sql_where = sql_where & "OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(cmbYPparm.Base_YYYYPP, cmbYPparm.RelativeStartingPeriod) & "' and OPS_YYYYPP <= '" & ASCMAIN1.Period_Calc(cmbYPparm.Base_YYYYPP, cmbYPparm.RelativeEndingPeriod) & "'"
                                Prepare_for_View_Lookup_Special(CTL, COLUMN_NAME, sql_where)
                                sql = ASCMAIN1.CodeSelector.Get_SQL(VIEW_NAME, TABLE_NAME_view, sql_where) ' GLFFINM1
                            ElseIf cmbYWparm.Base_YYYYWW <> "" Then
                                sql_where = sql_where & "YYYYWW >= '" & ASCMAIN1.Week_Calc(cmbYWparm.Base_YYYYWW, cmbYWparm.RelativeStartingWeek) & "' and YYYYWW <= '" & ASCMAIN1.Week_Calc(cmbYWparm.Base_YYYYWW, cmbYWparm.RelativeEndingWeek) & "'"
                                Prepare_for_View_Lookup_Special(CTL, COLUMN_NAME, sql_where)
                                sql = ASCMAIN1.CodeSelector.Get_SQL(VIEW_NAME, TABLE_NAME_view, sql_where) ' GLFFINM1
                            Else
                                'sql = ASCMAIN1.CodeSelector.Get_SQL(VIEW_NAME, TABLE_NAME_view, "ROWNUM < 1")
                                Prepare_for_View_Lookup_Special(CTL, COLUMN_NAME, sql_where)
                                sql = ASCMAIN1.CodeSelector.Get_SQL(VIEW_NAME, TABLE_NAME_view) ' GLFJRNL1
                            End If

                        End If
                    End If

                    If sql = "" Then
                        MsgBox("No SQL Statement for cmb control (" & VIEW_NAME & ")")
                        Stop
                    End If
                    Dim tbl As DataTable = ASCDATA1.GetDataTable(sql)
                    cmbctl.DataSource = tbl
                End If

                If cmbYP.ContainsKey(COLUMN_NAME) Then
                    cmbctl.DisplayLayout.Bands(0).SortedColumns.Add(cmbctl.DisplayLayout.Bands(0).Columns(0).Key, False)
                    If cmbYPparm.Parent_cmbYP <> "" Then
                        'cmbctl.ActiveRow = cmbctl.Rows(cmbYPparm.TotalRelativePeriods - 1)
                        cmbctl.ActiveRow = cmbctl.Rows(cmbYPparm.RelativeDefaultPeriod)
                    Else
                        cmbctl.ActiveRow = cmbctl.Rows(cmbYPparm.RelativeDefaultPeriod - cmbYPparm.RelativeStartingPeriod)
                    End If
                End If

                If cmbYW.ContainsKey(COLUMN_NAME) Then
                    cmbctl.DisplayLayout.Bands(0).SortedColumns.Add(cmbctl.DisplayLayout.Bands(0).Columns(0).Key, False)
                    If cmbYWparm.Parent_cmbYW <> "" Then
                        'cmbctl.ActiveRow = cmbctl.Rows(cmbYPparm.TotalRelativePeriods - 1)
                        cmbctl.ActiveRow = cmbctl.Rows(cmbYWparm.RelativeDefaultWeek)
                    Else
                        cmbctl.ActiveRow = cmbctl.Rows(cmbYWparm.RelativeDefaultWeek - cmbYWparm.RelativeStartingWeek)
                    End If
                End If

                If ASCMAIN1.CodeSelector.VIEW_NAME <> "" Then
                    If ASCMAIN1.CodeSelector.grdColumns.Count <> 0 Then
                        For I As Integer = 1 To ASCMAIN1.CodeSelector.grdColumns.Count
                            cmbctl.DisplayLayout.Bands(0).Columns(I - 1).Header.Caption = ASCMAIN1.CodeSelector.grdColumns(I - 1).Item("COLUMN_CAPTION")
                            If Val(ASCMAIN1.CodeSelector.grdColumns(I - 1).Item("COLUMN_WIDTH") & "") <> 0 Then
                                cmbctl.DisplayLayout.Bands(0).Columns(I - 1).Width = ASCMAIN1.CodeSelector.grdColumns(I - 1).Item("COLUMN_WIDTH")
                            End If
                        Next
                        cmbctl.DisplayLayout.Bands(0).SortedColumns.Add(cmbctl.DisplayLayout.Bands(0).Columns(0), False)
                    End If
                End If
            End If
        Next

    End Sub

    Sub Initialize_Controls_for_a_Container(ByVal c As Control)

        For Each CTL As Control In c.Controls
            If CTL.Controls.Count > 0 Then
                Call Initialize_Controls_for_a_Container(CTL)
            End If
            Dim PARENT_COLUMN_NAME As String = ""

            COLUMN_NAME = Absx1.GetABSColumnName(CTL)
            'If COLUMN_NAME = "OPS_YYYYPP" Then Stop
            Dim TABLE_NAME As String = Absx1.GetABSTableName(CTL)
            Dim TC As String = COLUMN_NAME
            If TABLE_NAME <> "" Then
                TC = TABLE_NAME & "." & COLUMN_NAME
            End If

            If TypeOf CTL Is Infragistics.Win.Misc.UltraGroupBox Then
                Dim grpctl As Infragistics.Win.Misc.UltraGroupBox = DirectCast(CTL, Infragistics.Win.Misc.UltraGroupBox)
                If grpctl.BorderStyle = Misc.GroupBoxBorderStyle.Default Then
                    grpctl.BorderStyle = Misc.GroupBoxBorderStyle.Rectangular3D
                End If
            End If

            If COLUMN_NAME <> "" Then
                Dim FORM_NAME As String = Me.Name
                Dim TABLE_NAME_TIP As String = TABLE_NAME
                If TABLE_NAME_TIP = "" Then
                    TABLE_NAME_TIP = Absx1.GetABSColumnName(Me)
                End If
                If TABLE_NAME_TIP = "" Then
                    TABLE_NAME_TIP = Me.Name
                    Mid(TABLE_NAME_TIP, 3, 1) = "T"
                End If

                Dim rowASTTTIP1 As DataRow = dst.Tables("ASTTTIP1").Rows.Find _
                (New String() {FORM_NAME, TABLE_NAME_TIP, COLUMN_NAME})
                If rowASTTTIP1 IsNot Nothing Then
                    Dim TTI As New UltraWinToolTip.UltraToolTipInfo
                    TTI.ToolTipTitle = rowASTTTIP1.Item("TOOLTIP_TITLE") & ""
                    TTI.ToolTipTextFormatted = vbCrLf & rowASTTTIP1.Item("TOOLTIP_TEXT") & ""

                    If TABLE_NAME_TIP <> "*" Then
                        tip.SetUltraToolTip(CTL, TTI)
                    Else
                        tip.SetUltraToolTip(Me, TTI)
                    End If

                    tip.AutoPopDelay = 15000
                    tip.InitialDelay = 1500
                    'tip.SetUltraToolTip(CurrentControl, TTI)
                    'tip.ShowToolTip(CurrentControl)
                End If

                PARENT_COLUMN_NAME = Absx1.GetABSParentColumnName(CTL)
                Try
                    If PARENT_COLUMN_NAME = "" Then
                        If Not htbCOLUMN_NAME.Contains(COLUMN_NAME) Then
                            htbCOLUMN_NAME.Add(COLUMN_NAME, CTL)
                        End If
                    Else
                        If Not htbCOLUMN_NAME.Contains(PARENT_COLUMN_NAME & "." & COLUMN_NAME) Then
                            htbCOLUMN_NAME.Add(PARENT_COLUMN_NAME & "." & COLUMN_NAME, CTL)
                        End If
                    End If

                    AddHandler CTL.MouseHover, AddressOf ctl_MouseHover
                    AddHandler CTL.MouseEnter, AddressOf ctl_MouseEnter

                    'AddHandler CTL.MouseUp, AddressOf ctl_MouseUp
                    'AddHandler CTL.MouseMove, AddressOf ctl_MouseMove
                Catch ex As Exception
                    'MsgBox("Problem with Control linked to " & COLUMN_NAME, MsgBoxStyle.OkOnly, "Form May Not Load Correctly")
                End Try

                If TypeOf CTL Is UltraWinGrid.UltraCombo Then
                    Dim cmbctl As UltraWinGrid.UltraCombo = DirectCast(CTL, UltraWinGrid.UltraCombo)

                    CMBs.Add(COLUMN_NAME, cmbctl)

                    AddHandler cmbctl.KeyDown, AddressOf cmb_KeyDown
                    AddHandler cmbctl.ValueChanged, AddressOf cmb_ValueChanged
                    AddHandler cmbctl.Enter, AddressOf cmb_Enter
                    AddHandler cmbctl.Leave, AddressOf cmb_Leave
                    AddHandler cmbctl.BeforeDropDown, AddressOf cmb_BeforeDropDown
                    AddHandler cmbctl.InitializeLayout, AddressOf cmb_InitializeLayout
                    AddHandler cmbctl.AfterCloseUp, AddressOf cmb_AfterCloseUp
                    ASFMAIN1.UltraToolbarsManager1.SetContextMenuUltra(cmbctl, "txtMenu")
                End If


                If TypeOf CTL Is UltraWinEditors.UltraComboEditor Then
                    Dim cbectl As UltraWinEditors.UltraComboEditor = DirectCast(CTL, UltraWinEditors.UltraComboEditor)

                    'CMBs.Add(COLUMN_NAME, cmbctl)

                    AddHandler cbectl.KeyDown, AddressOf cbe_KeyDown
                    AddHandler cbectl.ValueChanged, AddressOf cbe_ValueChanged
                    'AddHandler cbectl.Enter, AddressOf cbe_Enter
                    'AddHandler cbectl.Leave, AddressOf cbe_Leave
                    AddHandler cbectl.BeforeDropDown, AddressOf cbe_BeforeDropDown
                    'AddHandler cbectl.InitializeLayout, AddressOf cbe_InitializeLayout
                    'AddHandler cbectl.AfterCloseUp, AddressOf cbe_AfterCloseUp
                    'ASFMAIN1.UltraToolbarsManager1.SetContextMenuUltra(cbectl, "txtMenu")
                End If

                If TypeOf CTL Is UltraWinEditors.UltraTextEditor Then
                    Dim txtctl As UltraWinEditors.UltraTextEditor = DirectCast(CTL, UltraWinEditors.UltraTextEditor)
                    If txtctl.ReadOnly Then
                        'ABSReadOnly.Add(COLUMN_NAME)
                        'ABSReadOnly.Add(TC)
                        ABSReadOnly.Add(CTL.Name)
                    End If
                    VIEW_NAME = Absx1.GetABSViewName(txtctl)
                    PARENT_COLUMN_NAME = Absx1.GetABSParentColumnName(txtctl)
                    If Absx1.GetABSHasButton(txtctl) Then
                        If COLUMN_NAME <> "" Then
                            If VIEW_NAME = "" Then
                                VIEW_NAME = COLUMN_NAME
                            End If
                            Dim tbl As DataTable = ASCDATA1.GetDataTable _
                            ("Select * from ASTVIEW1 where VIEW_NAME = '" & VIEW_NAME & "'", "ASTVIEW1")
                            If tbl.Rows.Count > 0 Then
                                Dim btn As New UltraWinEditors.EditorButton
                                txtctl.ButtonsRight.Add(btn)
                                btn.Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "ARROW_UP_BLUE")

                                AddHandler txtctl.EditorButtonClick, AddressOf txt_EditorButtonClick
                            End If
                        End If
                    End If
                    AddHandler txtctl.KeyDown, AddressOf txt_KeyDown
                    AddHandler txtctl.ValueChanged, AddressOf txt_ValueChanged
                    AddHandler txtctl.Leave, AddressOf txt_Leave
                    AddHandler txtctl.Enter, AddressOf txt_Enter
                    'AddHandler txtctl.GotFocus, AddressOf txt_GotFocus
                    ASFMAIN1.UltraToolbarsManager1.SetContextMenuUltra(txtctl, "txtMenu")
                End If

                If TypeOf CTL Is UltraWinEditors.UltraCheckEditor Then
                    Dim chkctl As UltraWinEditors.UltraCheckEditor = DirectCast(CTL, UltraWinEditors.UltraCheckEditor)
                    AddHandler chkctl.CheckedChanged, AddressOf chk_CheckedChanged
                    AddHandler chkctl.BeforeCheckStateChanged, AddressOf chk_BeforeCheckStateChanged
                    'AddHandler chkctl.BindingContextChanged, AddressOf chk_BindingContextChanged
                    chkctl.Appearance.ForeColorDisabled = System.Drawing.Color.Black
                End If

                If TypeOf CTL Is UltraWinEditors.UltraNumericEditor Then
                    Dim numctl As UltraWinEditors.UltraNumericEditor = DirectCast(CTL, UltraWinEditors.UltraNumericEditor)
                    If numctl.ReadOnly Then
                        'ABSReadOnly.Add(COLUMN_NAME)
                        'ABSReadOnly.Add(TC)
                        ABSReadOnly.Add(CTL.Name)
                    End If
                    AddHandler numctl.KeyDown, AddressOf num_KeyDown
                    AddHandler numctl.ValueChanged, AddressOf num_ValueChanged
                    AddHandler numctl.Leave, AddressOf num_Leave
                    AddHandler numctl.ValidationError, AddressOf num_ValidationError
                End If



                If TypeOf CTL Is UltraWinEditors.UltraDateTimeEditor Then
                    Dim dtectl As UltraWinEditors.UltraDateTimeEditor = DirectCast(CTL, UltraWinEditors.UltraDateTimeEditor)
                    If dtectl.ReadOnly Then
                        'ABSReadOnly.Add(COLUMN_NAME)
                        'ABSReadOnly.Add(TC)
                        ABSReadOnly.Add(CTL.Name)
                    End If
                    AddHandler dtectl.KeyDown, AddressOf dte_KeyDown
                    AddHandler dtectl.ValueChanged, AddressOf dte_ValueChanged
                End If

                If TypeOf CTL Is UltraWinSchedule.UltraCalendarCombo Then
                    Dim dtcctl As UltraWinSchedule.UltraCalendarCombo = DirectCast(CTL, UltraWinSchedule.UltraCalendarCombo)
                    If dtcctl.ReadOnly Then
                        'ABSReadOnly.Add(COLUMN_NAME)
                        'ABSReadOnly.Add(TC)
                        ABSReadOnly.Add(CTL.Name)
                    End If
                    'dtcctl.DayOfWeekCaptionStyle = UltraWinSchedule.DayOfWeekCaptionStyle.ShortDescription
                    'dtcctl.DayOfWeekCaptionStyle = UltraWinSchedule.DayOfWeekCaptionStyle.ShortDescription
                    dtcctl.Format = "MM/dd/yyyy" ' If dtcctl.Format = "" Then dtcctl.Format = "MM/dd/yyyyy"
                    dtcctl.YearScrollButtonsVisible = DefaultableBoolean.True
                    dtcctl.NullDateLabel = ""
                    dtcctl.DropDownAppearance.BackColor2 = Color.DodgerBlue
                    dtcctl.DropDownAppearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

                    dtcctl.MonthScrollButtonAppearance.BackColor2 = Color.DodgerBlue
                    dtcctl.MonthScrollButtonAppearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal


                    dtcctl.MonthPopupAppearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    dtcctl.MonthPopupAppearance.BackColor2 = Color.DodgerBlue
                    'AddHandler dtcctl.KeyDown, AddressOf dtc_KeyDown
                    'AddHandler dtcctl.ValueChanged, AddressOf dtc_ValueChanged
                    AddHandler dtcctl.ValidationError, AddressOf cal_ValidationError
                    AddHandler dtcctl.KeyDown, AddressOf cal_KeyDown

                End If

                If TypeOf CTL Is UltraWinMaskedEdit.UltraMaskedEdit Then
                    Dim medctl As UltraWinMaskedEdit.UltraMaskedEdit = DirectCast(CTL, UltraWinMaskedEdit.UltraMaskedEdit)
                    If medctl.ReadOnly Then
                        'ABSReadOnly.Add(COLUMN_NAME)
                        'ABSReadOnly.Add(TC)
                        ABSReadOnly.Add(CTL.Name)
                    End If
                    medctl.DataMode = UltraWinMaskedEdit.MaskMode.Raw
                    AddHandler medctl.KeyDown, AddressOf med_KeyDown
                    AddHandler medctl.Validated, AddressOf med_Validated
                    AddHandler medctl.Invalidated, AddressOf med_Invalidated
                    AddHandler medctl.Validating, AddressOf med_Validating
                    AddHandler medctl.MaskValidationError, AddressOf med_MaskValidationError
                End If

                If TypeOf CTL Is UltraWinEditors.UltraOptionSet Then
                    Dim optctl As UltraWinEditors.UltraOptionSet = DirectCast(CTL, UltraWinEditors.UltraOptionSet)
                    'If optctl.ReadOnly Then
                    '    ABSReadOnly.Add(COLUMN_NAME)
                    'End If
                    AddHandler optctl.ValueChanged, AddressOf opt_ValueChanged
                    AddHandler optctl.KeyDown, AddressOf opt_KeyDown
                End If

            End If


            If TypeOf CTL Is UltraWinEditors.UltraOptionSet Then
                Dim optctl As UltraWinEditors.UltraOptionSet = DirectCast(CTL, UltraWinEditors.UltraOptionSet)
                optctl.Appearance.BackColorAlpha = Infragistics.Win.Alpha.Transparent
                optctl.Appearance.ForeColorDisabled = System.Drawing.Color.Black
                optctl.ItemAppearance.ForeColorDisabled = System.Drawing.Color.Black
            End If


            If TypeOf CTL Is UltraWinTabControl.UltraTabControl Then
                Dim tabctl As UltraWinTabControl.UltraTabControl = DirectCast(CTL, UltraWinTabControl.UltraTabControl)

                For Each TAB As UltraWinTabControl.UltraTab In tabctl.Tabs
                    If TAB.Key = "" Then
                        TAB.Key = TAB.Text
                    End If
                Next
            End If

            If TypeOf CTL Is UltraWinGrid.UltraGrid Then
                Dim grdctl As UltraWinGrid.UltraGrid = DirectCast(CTL, UltraWinGrid.UltraGrid)

                If grdctl.Name Like "grd*" Then
                    If Not GRDs.ContainsKey(Mid(grdctl.Name, 4)) Then
                        GRDs.Add(Mid(grdctl.Name, 4), grdctl)
                    End If
                End If
                Add_Handlers_grd(grdctl)
                'AddHandler grdctl.BeforeExitEditMode, AddressOf grd_BeforeExitEditMode
                'AddHandler grdctl.KeyDown, AddressOf grd_KeyDown
                'AddHandler grdctl.KeyPress, AddressOf grd_KeyPress
                'AddHandler grdctl.LostFocus, AddressOf grd_LostFocus
                'AddHandler grdctl.Leave, AddressOf grd_Leave
                'AddHandler grdctl.InitializeLayout, AddressOf grd_InitializeLayout
                'AddHandler grdctl.AfterRowUpdate, AddressOf grd_AfterRowUpdate
                'AddHandler grdctl.AfterRowInsert, AddressOf grd_AfterRowInsert
                'AddHandler grdctl.MouseUp, AddressOf grd_MouseUp
                'AddHandler grdctl.MouseDown, AddressOf grd_MouseDown
                'AddHandler grdctl.MouseEnterElement, AddressOf grd_MouseEnterElement
                'AddHandler grdctl.MouseLeaveElement, AddressOf grd_MouseLeaveElement
                'AddHandler grdctl.AfterEnterEditMode, AddressOf grd_AfterEnterEditMode
                'AddHandler grdctl.AfterExitEditMode, AddressOf grd_AfterExitEditMode
                'AddHandler grdctl.BeforeRowsDeleted, AddressOf grd_BeforeRowsDeleted
                'AddHandler grdctl.CellChange, AddressOf grd_CellChange
                'AddHandler grdctl.DoubleClickCell, AddressOf grd_DoubleClickCell
                ''AddHandler grdctl.Click, AddressOf grd_Click
                'AddHandler grdctl.InitializeRow, AddressOf grd_InitializeRow
                'AddHandler grdctl.ClickCellButton, AddressOf grd_ClickCellButton
                ''ASFMAIN1.UltraToolbarsManager1.SetContextMenuUltra(grdctl, "txtMenu")

            End If
        Next
    End Sub

    Sub Add_Handlers_grd(grdctl As UltraWinGrid.UltraGrid)
        AddHandler grdctl.BeforeExitEditMode, AddressOf grd_BeforeExitEditMode
        AddHandler grdctl.KeyDown, AddressOf grd_KeyDown
        AddHandler grdctl.KeyPress, AddressOf grd_KeyPress
        AddHandler grdctl.LostFocus, AddressOf grd_LostFocus
        AddHandler grdctl.Leave, AddressOf grd_Leave
        AddHandler grdctl.InitializeLayout, AddressOf grd_InitializeLayout
        AddHandler grdctl.AfterRowUpdate, AddressOf grd_AfterRowUpdate
        AddHandler grdctl.AfterRowInsert, AddressOf grd_AfterRowInsert
        AddHandler grdctl.MouseUp, AddressOf grd_MouseUp
        AddHandler grdctl.MouseDown, AddressOf grd_MouseDown
        AddHandler grdctl.MouseEnterElement, AddressOf grd_MouseEnterElement
        AddHandler grdctl.MouseLeaveElement, AddressOf grd_MouseLeaveElement
        AddHandler grdctl.AfterEnterEditMode, AddressOf grd_AfterEnterEditMode
        AddHandler grdctl.AfterExitEditMode, AddressOf grd_AfterExitEditMode
        AddHandler grdctl.BeforeRowsDeleted, AddressOf grd_BeforeRowsDeleted
        AddHandler grdctl.CellChange, AddressOf grd_CellChange
        AddHandler grdctl.DoubleClickCell, AddressOf grd_DoubleClickCell
        'AddHandler grdctl.Click, AddressOf grd_Click
        AddHandler grdctl.InitializeRow, AddressOf grd_InitializeRow
        AddHandler grdctl.ClickCellButton, AddressOf grd_ClickCellButton
        AddHandler grdctl.Error, AddressOf grd_Error
        'ASFMAIN1.UltraToolbarsManager1.SetContextMenuUltra(grdctl, "txtMenu")
    End Sub
    Public Sub Set_ScreenMode_Base(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        ' ScreenMode should not be used for ASFBASE2 forms - only for ASFBASE1 forms, 
        '  and for that reason we probably should relocate some of these methods 
        '  and form level variables to ASFBASE1

        'If New String() {"SOFORDR1"}.Contains(ASCMAIN1.ActiveForm.MENU_ITEM_OBJECT) Then
        '    If ASCMAIN1.DBS_COMPANY = "TFP" And ASCMAIN1.DBS_SERVER = "EMP" And Not ASCMAIN1.Running_in_VS And Not InquiryMode Then
        '        ASCMAIN1.ActiveForm.ASFBASE1_Fill_Panel.Visible = False
        '        MsgBox("Not Available in Live Company (yet)", MsgBoxStyle.OkOnly, "Please Exit this Form")
        '    End If
        '    If ASCMAIN1.DBS_COMPANY = "EMP" And ASCMAIN1.DBS_SERVER = "EMP" And Not ASCMAIN1.Running_in_VS And Not InquiryMode Then
        '        ASCMAIN1.ActiveForm.ASFBASE1_Fill_Panel.Visible = False
        '        MsgBox("Not Available in Live Company (yet)", MsgBoxStyle.OkOnly, "Please Exit this Form")
        '    End If
        '    If ASCMAIN1.DBS_COMPANY = "COS" And ASCMAIN1.DBS_SERVER = "EMP" And Not ASCMAIN1.Running_in_VS And Not InquiryMode Then
        '        ASCMAIN1.ActiveForm.ASFBASE1_Fill_Panel.Visible = False
        '        MsgBox("Not Available in Live Company (yet)", MsgBoxStyle.OkOnly, "Please Exit this Form")
        '    End If
        'End If

        ScreenMode = tf

        If Not tf Then
            IsDone = True
            If ctl1 Is Nothing Then
                'ctl1 = Me.ActiveControl
            Else
                Try
                    If ctl1 IsNot Nothing Then
                        ctl1.Focus()
                        ctl1.SelectNextControl(ctl1, True, False, True, True)
                        ctl1.Focus()
                    End If
                Catch ex As Exception

                End Try
            End If
        End If

        iScreenMode = IIf(ScreenMode, Infragistics.Win.DefaultableBoolean.True, Infragistics.Win.DefaultableBoolean.False)
        not_iScreenMode = IIf(ScreenMode, Infragistics.Win.DefaultableBoolean.False, Infragistics.Win.DefaultableBoolean.True)

        If Not ScreenMode Then
            Call ASCMAIN1.MultiTask_Release()
            EntryMode = ""
            XNO = ""
            This_Record_Inquiry_Only = False
            FILENAMEs_to_Publish.Clear()
        Else
            XNO = ASCMAIN1.Next_Control_No(MENU_ITEM_OBJECT & ".XNO")
        End If

        Call Set_MODE_in_StatusBar1(MODE_description)

    End Sub

    Sub Set_MODE_in_StatusBar1(Optional ByVal MODE_description As String = "")
        If Not ScreenMode Then
            ASFMAIN1.UltraStatusBar1.Panels("MODE").Text = ""
            ASFMAIN1.UltraStatusBar1.Panels("MODE").Appearance.BackColor = Color.Empty
            ASFMAIN1.UltraStatusBar1.Panels("MODE").Appearance.ForeColor = Color.Empty

        Else
            Dim MODE As String = ""
            If MODE_description <> "" Then
                MODE = MODE_description
            Else
                If This_Record_Inquiry_Only Then
                    MODE = "Inquiry"
                ElseIf EntryMode = "N" Then
                    MODE = "New"
                ElseIf EntryMode = "E" Then
                    MODE = "Edit"
                End If
            End If

            ASFMAIN1.UltraStatusBar1.Panels("MODE").Text = MODE

            If MODE <> "" Then
                ASFMAIN1.UltraStatusBar1.Panels("MODE").Appearance.BackColor = Color.OrangeRed ' .Orange ' .Gold ' .Lime ' .OrangeRed '  .BlueViolet
                ASFMAIN1.UltraStatusBar1.Panels("MODE").Appearance.TextHAlign = HAlign.Center
                ASFMAIN1.UltraStatusBar1.Panels("MODE").Appearance.ForeColor = Color.White
            End If

        End If
    End Sub

    Sub Populate_Dependent_Controls(ByVal COLUMN_NAME As String, ByRef dr As DataRow)

        ' asfsprf1 uses this, but what is the difference between this and
        ' populatecontrolswithparents
        'If COLUMN_NAME <> "" Then
        '    Call Leaving_txt_Special_Before(COLUMN_NAME, txtctl)
        '    Call Populate_Controls_with_Parents(COLUMN_NAME, txtctl)
        '    Call Leaving_txt_Special_After(COLUMN_NAME, txtctl)
        'End If

        For Each ctl As Control In Absx1.CtlsFor(COLUMN_NAME)
            If TypeOf ctl Is ABSCS.ABSCheckBox Then
                Dim chk As ABSCS.ABSCheckBox = DirectCast(ctl, ABSCS.ABSCheckBox)
                chk.ABSChecked = dr.Item(Absx1.GetABSColumnName(ctl)) & ""
            Else
                ctl.Text = dr.Item(Absx1.GetABSColumnName(ctl)) & ""
            End If
        Next
        Call Populate_Dependent_Controls_Special(COLUMN_NAME, dr)
    End Sub

    Public Overridable Sub Populate_Dependent_Controls_Special(ByVal COLUMN_NAME As String, ByRef dr As DataRow)

    End Sub

    Public Overridable Function OK_to_do_View_Lookup(ByVal txtctl As UltraWinEditors.UltraTextEditor) As Boolean
        Return True
    End Function

    Public Overridable Sub txt_EditorButtonClick(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinEditors.EditorButtonEventArgs)

        Dim txtctl As UltraWinEditors.UltraTextEditor
        txtctl = DirectCast(sender, UltraWinEditors.UltraTextEditor)

        If txtctl.ReadOnly Then
            Exit Sub
        End If

        If Not OK_to_do_View_Lookup(txtctl) Then
            Exit Sub
        End If

        Call Prepare_for_View_Lookup(txtctl)

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = False
            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = DirectCast(sender, UltraWinEditors.UltraTextEditor).Text

            Dim PKs As String = Absx1.GetABSPrecedentKeys(txtctl)
            If PKs <> "" Then
                ' need to be able to get .value of an optionset
                ' need to be able to work with multiple (perhaps comma separated) column names in pks
                Dim COLUMN_NAME_pk As String = PKs
                Dim COLUMN_NAME_alias As String = PKs
                If COLUMN_NAME_pk.Contains(":") Then
                    COLUMN_NAME_pk = Split(PKs, ":")(0)
                    COLUMN_NAME_alias = Split(PKs, ":")(1)
                End If

                ASCMAIN1.CodeSelector.Precedent_Keys.Add(COLUMN_NAME_alias, Absx1.txtFor(COLUMN_NAME_pk).Text)
                'ASCMAIN1.CodeSelector.Precedent_Keys.Add(PKs, Absx1.txtFor(PKs).Text)
            End If

            Dim WHERE_CLAUSEs As String = ""
            For Each rowASTSECK1 As DataRow In ASCMAIN1.tblASTSECK1.Select("TABLE_NAME = '" & ASCMAIN1.CodeSelector.TABLE_NAME & "'")
                If WHERE_CLAUSEs = "" Then
                    WHERE_CLAUSEs = "x"
                End If
                Dim SECURITY_CODE As String = rowASTSECK1.Item("SECURITY_CODE")
                If SECURITY_CODE = "**" Or ASCMAIN1.USER_SECURITY_CODEs.Contains(SECURITY_CODE) Then
                    Dim WHERE_CLAUSE As String = rowASTSECK1.Item("WHERE_CLAUSE") & ""
                    If WHERE_CLAUSE = "" Then
                        WHERE_CLAUSEs = ""
                        Exit For
                    Else
                        WHERE_CLAUSEs &= " or (" & WHERE_CLAUSE & ")"
                    End If
                End If
            Next
            If WHERE_CLAUSEs <> "" Then
                If WHERE_CLAUSEs = "x" Then
                    WHERE_CLAUSEs &= " or ROWNUM < 1"
                End If
                'ASCMAIN1.CodeSelector.SQL = "Select * from (" & ASCMAIN1.CodeSelector.SQL & ") where " & Mid(WHERE_CLAUSEs, 4)
                Dim WHERE_OR_AND As String = " WHERE "
                If InStr(ASCMAIN1.CodeSelector.SQL.ToUpper, " WHERE ") <> 0 Then
                    WHERE_OR_AND = " AND "
                End If
                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.SQL & WHERE_OR_AND & "(" & Mid(WHERE_CLAUSEs, 6) & ")"
            End If

            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                txtctl.Text = ASCMAIN1.CodeSelector.SelectedCode
                Call txt_EditorButtonClick_Special(txtctl)
                Dim COLUMN_NAME As String = Absx1.GetABSColumnName(txtctl)
                Call Leaving_txt_Special_Before(COLUMN_NAME, txtctl)
                Call Populate_Controls_with_Parents(COLUMN_NAME, txtctl)
                Call Leaving_txt_Special_After(COLUMN_NAME, txtctl)
            End If
        End If

    End Sub

    Overridable Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)

    End Sub

    Sub Prepare_for_View_Lookup(
    ByVal ctl As Control,
    Optional ByVal key_value As String = "",
    Optional ByVal sql_where As String = "")

        COLUMN_NAME = Absx1.GetABSColumnName(ctl)
        VIEW_NAME = Absx1.GetABSViewName(ctl)
        If VIEW_NAME = "" Then
            VIEW_NAME = COLUMN_NAME
        End If
        TABLE_NAME_view = Absx1.GetABSLookUpTableName(ctl) '  .GetABSTableName(txtctl)
        Dim TABLE_NAME_ctl As String = Absx1.GetABSTableName(ctl)
        ' changes here involving TABLE_NAME_ctl and TABLE_NAME_view 12/08/07 on LENS_DESIGN_CODE in DEFJOBM1
        'TABLE_NAME_view = Absx1.GetABSTableName(ctl)
        If TABLE_NAME_view = "" Then
            'TABLE_NAME_view = TABLE_NAME

            ' next line remmed and following 5 lines added to get APTVEND1 to show AP Post codes rather than AR post codes
            'TABLE_NAME_view = TABLE_NAME_ctl
            If TABLE_NAME_ctl <> "" Then
                TABLE_NAME_view = TABLE_NAME_ctl
            Else
                TABLE_NAME_view = TABLE_NAME
            End If

        End If

        Dim Cancel As Boolean
        Call Prepare_for_View_Lookup_Special(ctl, COLUMN_NAME, sql_where, Cancel)
        If Cancel Then
            ASCMAIN1.CodeSelector.SQL = ""
        Else
            ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL(VIEW_NAME, TABLE_NAME_view, sql_where, key_value)
        End If

    End Sub

    Public Overridable Sub Prepare_for_View_Lookup_Special(
    ByVal ctl As Control,
    ByVal COLUMN_NAME As String,
    Optional ByRef sql_where As String = "",
    Optional ByRef Cancel As Boolean = False)

    End Sub

    Private Sub txt_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        Set_MRU(sender)
    End Sub

    Sub Set_MRU(ByVal sender As Object)
        ASCMAIN1.MRU_used = False
        ASCMAIN1.MRU_COLUMN_NAME = ""

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(DirectCast(sender, Control))
        If ASCMAIN1.MRUs.ContainsKey(COLUMN_NAME) Then
            ASCMAIN1.MRU_txtctl = DirectCast(sender, UltraWinEditors.UltraTextEditor)
            If Not ASCMAIN1.MRU_txtctl.ReadOnly Then
                ASCMAIN1.MRU_COLUMN_NAME = COLUMN_NAME
            End If
        End If
    End Sub

    Private Sub txt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Public Overridable Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            Dim txt As UltraWinEditors.UltraTextEditor
            txt = DirectCast(sender, UltraWinEditors.UltraTextEditor)

            If Not txt.Multiline Then
                txt_Leave(sender, e)
                'txt.SelectNextControl(txt, True, True, True, True) ' doesn't work
                'txt.GetNextControl(txt, True).Focus() ' doesn't work
                ' why was this sendkeys knocked out?  needed for things like quick entry

                SendKeys.Send(Chr(9))
            End If
            'Dim ctl As Control
            'ctl = txt.GetNextControl(txt, False)
            'ctl.Focus()
        End If
    End Sub

    Public Overridable Sub txt_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim COLUMN_NAME As String

        ASCMAIN1.MRU_COLUMN_NAME = ""

        Dim txtctl As UltraWinEditors.UltraTextEditor
        txtctl = DirectCast(sender, UltraWinEditors.UltraTextEditor)

        If txtctl.Modified Then
            COLUMN_NAME = Absx1.GetABSColumnName(txtctl)
            txtctl.Text = Trim$(txtctl.Text)
            If txtctl.Text <> "" Then
                txtctl.Text = ASCMAIN1.Format_Field(txtctl.Text, COLUMN_NAME, tblASFBASE1)
            End If

            If COLUMN_NAME <> "" Then
                Call Leaving_txt_Special_Before(COLUMN_NAME, txtctl)
                Call Populate_Controls_with_Parents(COLUMN_NAME, txtctl)
                Call Leaving_txt_Special_After(COLUMN_NAME, txtctl)
            End If
        End If

    End Sub

    Public Overridable Sub Leaving_txt_Special_Before(ByVal COLUMN_NAME As String, ByVal ctl As Control)

    End Sub

    Public Overridable Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)

    End Sub

    Public Overridable Sub txt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim txtctl As UltraWinEditors.UltraTextEditor = DirectCast(sender, UltraWinEditors.UltraTextEditor)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(txtctl)
        'If COLUMN_NAME = "ORIG_CODE" Then Debug.Print("ORIG CHANGED")
        If COLUMN_NAME <> "" Then
            'Call Populate_Controls_with_Parents(COLUMN_NAME, txtctl) ' ENABLING THIS HERE CAUSED SOFORDR1 TO GET CONFUSED WHEN THERE WERE 2 TEXT BOXES CALLED SHIP_VIA_DESC
            Dim TABLE_NAME As String = Absx1.GetABSTableName(txtctl)
            If TABLE_NAME <> "" And TABLE_NAME <> Me.TABLE_NAME Then ' NEEDED TO ADD AND TABLE_NAME <> ME.TABLE_NAME TO GET VEND_NAME TO POPULATE WHEN CALLING UP AN AP VOUCHER
                Call Populate_Controls_with_Parents(TABLE_NAME & "." & COLUMN_NAME, txtctl)
            Else
                Call Populate_Controls_with_Parents(COLUMN_NAME, txtctl)
            End If
        End If
        If Not txtctl.Focused Or ASCMAIN1.MRU_used Then
            ASCMAIN1.MRU_used = False
        End If

    End Sub

    Public Overridable Sub opt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Sub Populate_Controls_with_Parents(
    ByVal PARENT_COLUMN_NAME As String,
    ByVal txtctl As Control)

        Dim cdr2 As DataRow = Nothing

        Dim initialized As Boolean = False
        'If PARENT_COLUMN_NAME = "SHIP_VIA_CODE" Or PARENT_COLUMN_NAME = "ARTCUST3.SHIP_VIA_CODE" Then Stop
        For Each CTL As Control In Absx1.CtlsFor(PARENT_COLUMN_NAME)
            If txtctl.Text = "" Then
                If TypeOf CTL Is ABSCS.ABSCheckBox Then
                    Dim chk As ABSCS.ABSCheckBox = DirectCast(CTL, ABSCS.ABSCheckBox)
                    chk.ABSChecked = "0"
                Else
                    If TypeOf CTL Is UltraWinEditors.UltraOptionSet Then
                        Dim optctl As UltraWinEditors.UltraOptionSet = DirectCast(CTL, UltraWinEditors.UltraOptionSet)
                        optctl.CheckedIndex = -1
                    Else
                        CTL.Text = ""
                    End If
                End If
            Else
                If Not initialized Then
                    initialized = True
                    Dim PRECEDENT_KEYS As String = Absx1.GetABSPrecedentKeys(txtctl)
                    'If PARENT_COLUMN_NAME = "INV_PYMT_METHOD" Then Stop
                    If PRECEDENT_KEYS <> "" Then
                        cdr = LookUp_for_txtctl(txtctl, Add_PK(PRECEDENT_KEYS))
                    Else
                        cdr = LookUp_for_txtctl(txtctl)
                    End If
                    If cdr IsNot Nothing Then
                        cdr2 = cdr.Table.NewRow
                        cdr2.ItemArray = cdr.ItemArray
                    End If
                End If

                If cdr2 Is Nothing Then
                    If TypeOf CTL Is ABSCS.ABSCheckBox Then
                        Dim chk As ABSCS.ABSCheckBox = DirectCast(CTL, ABSCS.ABSCheckBox)
                        chk.ABSChecked = "0"
                    Else
                        If TypeOf CTL Is UltraWinEditors.UltraOptionSet Then
                            Dim optctl As UltraWinEditors.UltraOptionSet = DirectCast(CTL, UltraWinEditors.UltraOptionSet)
                            optctl.CheckedIndex = -1
                        Else
                            CTL.Text = ""
                        End If
                    End If
                Else
                    If cdr2.Table.TableName = "ASTCODE1" Then
                        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(CTL)
                        If cdr2.Table.Columns.Contains(COLUMN_NAME) Then
                            CTL.Text = cdr2.Item(COLUMN_NAME) & ""
                        Else
                            CTL.Text = cdr2.Item("T_DESC") & ""
                        End If
                    Else
                        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(CTL)
                        If cdr2.Table.Columns.Contains(COLUMN_NAME) Then
                            Try
                                If TypeOf CTL Is ABSCS.ABSCheckBox Then
                                    Dim chkctl As ABSCS.ABSCheckBox = DirectCast(CTL, ABSCS.ABSCheckBox)
                                    chkctl.ABSChecked = cdr2.Item(COLUMN_NAME) & ""
                                Else
                                    If TypeOf CTL Is UltraWinEditors.UltraOptionSet Then
                                        Dim optctl As UltraWinEditors.UltraOptionSet = DirectCast(CTL, UltraWinEditors.UltraOptionSet)
                                        optctl.Value = cdr2.Item(COLUMN_NAME) & ""
                                    Else
                                        CTL.Text = cdr2.Item(COLUMN_NAME) & ""
                                    End If
                                End If
                            Catch ex As Exception
                                If TypeOf CTL Is ABSCS.ABSCheckBox Then
                                    Dim chk As ABSCS.ABSCheckBox = DirectCast(CTL, ABSCS.ABSCheckBox)
                                    chk.ABSChecked = "0"
                                Else
                                    If TypeOf CTL Is UltraWinEditors.UltraOptionSet Then
                                        Dim optctl As UltraWinEditors.UltraOptionSet = DirectCast(CTL, UltraWinEditors.UltraOptionSet)
                                        optctl.CheckedIndex = -1
                                    Else
                                        CTL.Text = "?"
                                    End If
                                End If
                            End Try
                        Else
                            If TypeOf CTL Is ABSCS.ABSCheckBox Then
                                Dim chk As ABSCS.ABSCheckBox = DirectCast(CTL, ABSCS.ABSCheckBox)
                                chk.ABSChecked = "0"
                            Else
                                If TypeOf CTL Is UltraWinEditors.UltraOptionSet Then
                                    Dim optctl As UltraWinEditors.UltraOptionSet = DirectCast(CTL, UltraWinEditors.UltraOptionSet)
                                    optctl.CheckedIndex = -1
                                Else
                                    If TypeOf CTL Is UltraWinEditors.UltraNumericEditor Then
                                        ' CTL.VALUE = 0
                                    Else
                                        CTL.Text = "?"
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub ctl_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        If ASCMAIN1.Running_in_VS Then
            ASCMAIN1.MainForm_pgd.SelectedObject = sender

            CurrentControl = DirectCast(sender, Control)
            'ASFMAIN1.lbltip.Text = CurrentControl.Name
            ''ASFMAIN1.tip.ContextMenuItems = FormattedLinkLabel.FormattedTextMenuItems.All
            'Dim COLUMN_NAME As String = Absx1.GetABSColumnName(CurrentControl)
            'Dim TABLE_NAME As String = Absx1.GetABSTableName(CurrentControl)
            'If TABLE_NAME = "" Then
            '    TABLE_NAME = Absx1.GetABSTableName(Me)
            'End If
            'If TABLE_NAME = "" Then
            '    TABLE_NAME = Me.Name
            '    Mid(TABLE_NAME, 3, 1) = "T"
            'End If

            'Try
            '    tip.ShowToolTip(CurrentControl)
            'Catch ex As Exception

            'End Try
        End If
    End Sub

    Public Overridable Sub ctl_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim ctlCOLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        'Dim ctlTABLE_NAME As String = Absx1.GetABSTableName(sender)
        'If ctlTABLE_NAME = "" Then
        '    ctlTABLE_NAME = TABLE_NAME
        'End If
        'ASFMAIN1.UltraStatusBar1.Panels(0).Text = ctlCOLUMN_NAME
    End Sub

    Public Overridable Sub ctl_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        If ASCMAIN1.Running_in_VS Then ASCMAIN1.MainForm_pgd.SelectedObject = sender
    End Sub

    Public Overridable Sub ctl_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        If ASCMAIN1.Running_in_VS Then ASCMAIN1.MainForm_pgd.SelectedObject = sender
    End Sub

    Public Overridable Sub chk_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim chkctl As UltraWinEditors.UltraCheckEditor
        chkctl = DirectCast(sender, UltraWinEditors.UltraCheckEditor)

        COLUMN_NAME = Absx1.GetABSColumnName(chkctl)

        If COLUMN_NAME Like "CHK#DATE*" Then
            Call Set_Date_Combo(chkctl, COLUMN_NAME)
        End If

        If COLUMN_NAME <> "" Then
            Call CheckedChanged_Special(COLUMN_NAME, chkctl)
        End If
    End Sub

    Sub Set_Date_Combo(
    ByVal chkctl As UltraWinEditors.UltraCheckEditor,
    ByVal COLUMN_NAME As String)

        Dim dteCOLUMN_NAME As String = "DTE" & Mid(COLUMN_NAME, 4)
        Dim DATE_VALUE As Date = Now

        Try
            If chkctl.CheckState = CheckState.Checked Then
                Absx1.dteFor(dteCOLUMN_NAME).Value = ""
                Absx1.dteFor(dteCOLUMN_NAME).Enabled = False
            Else
                Absx1.dteFor(dteCOLUMN_NAME).Value = DATE_VALUE
                Absx1.dteFor(dteCOLUMN_NAME).Enabled = True
            End If
        Catch ex As Exception
            ' Nothing 
        End Try
    End Sub

    Public Sub chk_BeforeCheckStateChanged(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Dim chkctl As UltraWinEditors.UltraCheckEditor = DirectCast(sender, UltraWinEditors.UltraCheckEditor)
        'Stop

        'If chkctl.Tag = "1" Then
        '    chkctl.Checked = True
        'End If
    End Sub

    Public Overridable Sub CheckedChanged_Special(ByVal COLUMN_NAME As String, ByVal chk As UltraWinEditors.UltraCheckEditor)

    End Sub

    Public Sub Create_TDA_NEW(
    ByRef tbl As DataTable,
    ByVal TABLE_NAME As String,
    ByVal selectOption As SelectOption,
    Optional ByVal numberOfKeysUsedToSelect As Integer = -1,
    Optional ByVal forUpdate As Boolean = True,
    Optional ByVal customParameters As String = "",
    Optional ByVal keyFieldCount As Integer = -1,
    Optional ByVal updateCOLUMN_NAMEs As String = "")

        Dim sqlCustom As String = String.Empty

        Select Case selectOption
            Case ABSolution.SelectOption.AllColumns
                sqlCustom = String.Format _
                ("Select * from {0}", TABLE_NAME)

            Case ABSolution.SelectOption.ASCMAIN1sql
                sqlCustom = ASCMAIN1.sql
        End Select

        Me.Create_TDA(tbl, TABLE_NAME, sqlCustom,
            numberOfKeysUsedToSelect, forUpdate,
            customParameters, keyFieldCount,
            updateCOLUMN_NAMEs)

    End Sub

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
    Public Sub Create_TDA(
    ByRef tbl As DataTable,
    ByVal TABLE_NAME As String,
    ByVal sql_custom As String,
    Optional ByVal NumberOfKeysUsedToSelect As Integer = -1,
    Optional ByVal for_update As Boolean = True,
    Optional ByVal custom_parameters As String = "",
    Optional ByVal Key_Field_Count As Integer = -1,
    Optional ByVal Update_COLUMN_NAMEs As String = "",
    Optional ByVal SCHEMA As String = "")

        clsASCBASE1.Create_TDA _
        (tbl, TABLE_NAME, sql_custom, NumberOfKeysUsedToSelect, for_update,
         custom_parameters, Key_Field_Count, Update_COLUMN_NAMEs, SCHEMA)
    End Sub

    Public Sub Get_PARM(ByVal PARM_TABLE_NAME As String)
        If Not ROWs.ContainsKey(PARM_TABLE_NAME) Then
            Call Create_Lookup(PARM_TABLE_NAME)
        End If
        ROWs(PARM_TABLE_NAME) = LookUp(PARM_TABLE_NAME, "Z")
    End Sub

    ''' <summary>
    ''' Returns an empty DataTable representing the ResultSet generated from the SQL statement provided.
    ''' A DataAdapter is created and linked to a SelectCommand formed with the SQL statement.
    ''' In code, you would normally get the ResultSet using one of the FillRecords methods: FillRecords("ASTMENU1").
    ''' The ResultSet is READ-ONLY (the Adapter is not prepared to perform an Update).
    ''' This function will cache the DataAdapter created in the TDAs collection, and the DataTable created in the TBLs collection, both keyed by RESULTSET_NAME.
    ''' Since the TDAs and TBLs collections are defined to ASFBASE1, the ResultSets defined in a single form belong to that form, and are not usable across forms.
    ''' If a ResultSet is defined twice in the same form, an error will result.
    ''' </summary>
    ''' <param name="RESULTSET_NAME">A logical name for the ResultSet.  
    ''' This is usually the TABLE_NAME used in the query.  
    ''' Note that the RESULTSET_NAME uses the same TDAs and TBLs collections as the Create_TDA function, 
    ''' so you should not use a RESULTSET_NAME which may conflict with a TABLE_NAME used to create an updatable DataAdapter (using Create_TDA).</param>
    ''' <param name="custom_parameters">A string indicating the either * for Select * from RESULTSET_NAME, or the number and types of parameters used in the where clause of ASCMAIN1.sql, such as "VVIND" to indicate 5 parameters, 2 Varchar2, 1 Integer, 1 Number, and 1 Date.  Always name your custom parameters like :PARMx.  Note that if you use "*" then there is no need for ASCMAIN1.sql, and if you do not use "*", then you must use ASCMAIN1.sql.</param>
    ''' <param name="Key_Field_Count">The number of leading columns in the DataTable that should be set up as a PrimaryKey for the DataTable.  If the SQL statement draws data from a single Database Table, then the PrimaryKey count can be set automatically from the Database definition.  A value of -1 will use whatever PrimaryKey that that DataAdapter infers from the FillSchema method, and a value of 0 means that you want no PrimaryKey.</param>
    ''' <remarks></remarks>
    Public Function Create_ResultSet(
    ByVal RESULTSET_NAME As String,
    Optional ByVal custom_parameters As String = "",
    Optional ByVal Key_Field_Count As Integer = -1) As DataTable

        ' probably need to have this procedure call Create_TDA, because there things such as the Decimal and Int64 settings that have been added to that procedure
        ' and those improvements probably ought to have been made to GetDataTable

        Dim sql As String
        If custom_parameters = "*" Then
            sql = "Select * from " & RESULTSET_NAME
        Else
            sql = ASCMAIN1.sql
        End If
        sql = ASCDATA1.SQL_Server_Conversions(sql)
        Dim tbl As New DataTable(RESULTSET_NAME)
        Dim ada As New OracleDataAdapter(sql, ASCMAIN1.oraCon)

        With ada
            If custom_parameters <> "" And custom_parameters <> "*" Then
                Dim ptbl As New DataTable
                Call ASCDATA1.Create_Parameters(ada.SelectCommand, custom_parameters, ptbl)
                pROWs.Add(RESULTSET_NAME, ptbl.NewRow)
            End If

            ASCDATA1.Fill_Schema(ada, tbl)

            If Key_Field_Count <> -1 Then
                If Key_Field_Count = 0 Then
                    tbl.PrimaryKey = Nothing
                Else
                    Dim PK(Key_Field_Count - 1) As DataColumn
                    For i As Integer = 0 To Key_Field_Count - 1
                        PK(i) = tbl.Columns(i)
                    Next
                    tbl.PrimaryKey = PK
                End If
            End If
            TDAs.Add(RESULTSET_NAME, ada)
            TBLs.Add(RESULTSET_NAME, tbl)
            Dim dvw As New DataView(tbl)
            DVWs.Add(RESULTSET_NAME, dvw) ' was TABLE_NAME - causing an error in GLFPEND1
        End With

        For Each dc As DataColumn In tbl.Columns
            If dc.ReadOnly Then
                dc.ReadOnly = False
            End If
        Next

        If CMDs.ContainsKey(RESULTSET_NAME) Then
            CMDs.Remove(RESULTSET_NAME)
        End If
        CMDs.Add(RESULTSET_NAME, ada.SelectCommand)
        If ROWs.ContainsKey(RESULTSET_NAME) Then
            ROWs.Remove(RESULTSET_NAME)
        End If
        ROWs.Add(RESULTSET_NAME, tbl.NewRow)



        Return tbl

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
    Public Sub Create_Lookup(
    ByVal LOOKUP_NAME As String,
    Optional ByVal column_list As String = "*",
    Optional ByVal where_clause As String = "",
    Optional ByVal custom_parameters As String = "",
    Optional ByVal create_parameters_for_key As Boolean = True)

        ' Examples of usage:
        'Create_Lookup("GLTACCT1")
        'Create_Lookup("POTORDR1.OPEN", , "PO_STATUS_CODE = 'O' and PO_ORDER_TYPE = 'S'")
        'Create_Lookup("ICTPRCAT", , "PROD_CODE in (Select Distinct PROD_CODE from ICTPROD1 where VEND_CODE = :PARM1)", "V")
        'Create_Lookup("ICTITEM1.ITEM_UPC_CODE", , "ITEM_UPC_CODE = :PARM1", "V", False)

        clsASCBASE1.Create_Lookup(LOOKUP_NAME, column_list, where_clause,
                                    custom_parameters, create_parameters_for_key)
    End Sub

    Public Function LookUp_for_txtctl(
    ByVal txtctl As Control,
    Optional ByVal Precedent_Key_Values As List(Of String) = Nothing) As DataRow

        cdr = Nothing

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(txtctl)

        Dim CODE_VALUE As String = txtctl.Text
        If TypeOf (txtctl) Is UltraWinEditors.UltraComboEditor Then
            CODE_VALUE = Absx1.cbeFor(COLUMN_NAME).Value
        End If

        Dim VIEW_NAME As String = Absx1.GetABSViewName(txtctl)
        ' WHY DON'T WE USE THE COLUMN NAME AS THE VIEW NAME IF THE VIEW NAME IS BLANK - THIS WOULD AVOID US HAVING TO SPECIFY THE VIEW NAME EXPLICITLY IN THE TEXT BOX = SEE EMP SOFORDR1.FRT_TERMS -> SOFORDR1.FRT_TERMS_DESC - BECAUSE THE DESC IS ALIASED, WE NEED TO USE THE VIEW NAME
        If VIEW_NAME <> "" Then
            ' LATER WE NEED TO LOOK AT THE TYPE OF CONTROL AND USE .VALUE OR .TEXT
            ' NEED TO PASS IN PRECEDENT KEYS - SO THAT WE CAN SHOW COLOR_DESC IN ICTITEM1 SCREEN
            Call Prepare_for_View_Lookup(txtctl, CODE_VALUE)

            If ASCMAIN1.CodeSelector.SQL <> "" Then
                ASCMAIN1.sql = ASCMAIN1.CodeSelector.SQL

                ' there can only be 1 precedent key for a text control - so why are we passing in a list?
                If Precedent_Key_Values IsNot Nothing AndAlso Precedent_Key_Values.Count <> 0 Then
                    Dim COLUMN_NAME_pkey As String = Absx1.GetABSPrecedentKeys(txtctl)


                    Dim COLUMN_NAME_pk As String = COLUMN_NAME_pkey
                    Dim COLUMN_NAME_alias As String = COLUMN_NAME_pkey
                    If COLUMN_NAME_pk.Contains(":") Then
                        COLUMN_NAME_pk = Split(COLUMN_NAME_pkey, ":")(0)
                        COLUMN_NAME_alias = Split(COLUMN_NAME_pkey, ":")(1)
                    End If

                    ' we should always be setting the key, so an and should be what we need here
                    ASCMAIN1.sql &= " and " & COLUMN_NAME_alias & " = '" & Precedent_Key_Values(0) & "'"
                    'ASCMAIN1.sql &= " and " & COLUMN_NAME_pkey & " = '" & Precedent_Key_Values(0) & "'"
                End If

                Dim tbl As DataTable = ASCDATA1.GetDataTable
                If tbl.Rows.Count = 0 Then
                    cdr = Nothing
                Else
                    cdr = tbl.Rows(0)
                End If
            End If
        Else
            Dim LOOKUP_TABLE_NAME As String = Absx1.GetABSLookUpTableName(txtctl)
            If LOOKUP_TABLE_NAME = "" Then
                Dim tbl As DataTable
                tbl = ASCDATA1.GetDataTable("Select * from ASTVIEW1 where VIEW_NAME = '" & COLUMN_NAME & "' and TABLE_NAME = '" & TABLE_NAME & "'")
                If tbl.Rows.Count = 0 Then
                    tbl = ASCDATA1.GetDataTable("Select * from ASTVIEW1 where VIEW_NAME = '" & COLUMN_NAME & "' and TABLE_NAME LIKE '" & Mid(TABLE_NAME, 1, 2) & "%'")
                End If
                'If tbl.Rows.Count = 0 Then
                ' THIS ONE SHOULD TRY TO USE ASTVIEW3
                '    tbl = ASCDATA1.GetDataTable("Select * from ASTVIEW1 where VIEW_NAME = '" & COLUMN_NAME & "' and TABLE_NAME LIKE '" & Mid(TABLE_NAME, 1, 2) & "'")
                'End If
                If tbl.Rows.Count = 0 Then
                    tbl = ASCDATA1.GetDataTable("Select * from ASTVIEW1 where VIEW_NAME = '" & COLUMN_NAME & "'")
                End If
                If tbl.Rows.Count <> 0 Then
                    LOOKUP_TABLE_NAME = tbl.Rows(0).Item("TABLE_NAME")
                    If tbl.Rows(0).Item("CODE_TABLE") & "" <> "" Then
                        Precedent_Key_Values = New List(Of String)
                        Precedent_Key_Values.Add(tbl.Rows(0).Item("CODE_TABLE"))
                        Precedent_Key_Values.Add(tbl.Rows(0).Item("CODE_COLUMN"))
                        LOOKUP_TABLE_NAME = "ASTCODE1"
                    End If
                End If
                If LOOKUP_TABLE_NAME = "" Then
                    MsgBox("No Lookup Table for " & COLUMN_NAME, MsgBoxStyle.OkOnly, "Cannot Find Record for Code Value")
                    'Stop ' PROBABLY NEED TO PUT THE LOOKUPTABLENAME INTO THE CONTROL
                    Return Nothing
                End If
            End If

            If Not CMDs.ContainsKey(LOOKUP_TABLE_NAME) Then
                Create_Lookup(LOOKUP_TABLE_NAME)
            End If
            If Precedent_Key_Values Is Nothing Then
                cdr = LookUp(LOOKUP_TABLE_NAME, CODE_VALUE)
            Else
                Dim KEYs() As String
                ReDim KEYs(Precedent_Key_Values.Count)
                For i As Integer = 0 To Precedent_Key_Values.Count - 1
                    KEYs(i) = Precedent_Key_Values.Item(i)
                Next
                KEYs(Precedent_Key_Values.Count) = CODE_VALUE
                cdr = LookUp(LOOKUP_TABLE_NAME, KEYs)
            End If
        End If
        Return cdr
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
    Public Function LookUp(
    ByVal TABLE_NAME As String,
    ByVal KEY() As String,
    Optional ByVal Return_Empty_Row_if_Missing As Boolean = False) As DataRow

        If clsASCBASE1 Is Nothing Then Return Nothing

        cdr = clsASCBASE1.LookUp(TABLE_NAME, KEY, Return_Empty_Row_if_Missing)
        Return cdr

    End Function

    Public Function LookUp(
    ByVal TABLE_NAME As String,
    ByVal KEY As String,
    Optional ByVal Return_Empty_Row_if_Missing As Boolean = False) As DataRow
        Return LookUp(TABLE_NAME, New String() {KEY}, Return_Empty_Row_if_Missing)

    End Function

    Public Sub grdCodeDesc _
    (ByRef grd As UltraWinGrid.UltraGrid,
     ByVal TABLE_NAME_CODE As String,
     ByVal COLUMN_NAME_CODE As String,
     ByVal COLUMN_NAME_DESC As String,
     Optional ByVal COLUMN_NAME_DESC_FROM_LOOKUP As String = "")

        If COLUMN_NAME_DESC_FROM_LOOKUP = "" Then
            COLUMN_NAME_DESC_FROM_LOOKUP = COLUMN_NAME_DESC
        End If

        With grd
            If .ActiveRow IsNot Nothing Then
                cdr = LookUp(TABLE_NAME_CODE, .ActiveRow.Cells(COLUMN_NAME_CODE).Text)
                If cdr Is Nothing Then
                    Try
                        .ActiveRow.Cells(COLUMN_NAME_DESC).Value = ""
                    Catch ex As Exception

                    End Try
                Else
                    Try
                        'If .ActiveRow.Cells.Contains(COLUMN_NAME_DESC) Then
                        .ActiveRow.Cells(COLUMN_NAME_DESC).Value = cdr.Item(COLUMN_NAME_DESC_FROM_LOOKUP)
                        'End If
                    Catch ex As Exception

                    End Try
                End If
            End If
        End With
    End Sub

    Public Sub grdFieldFormat(ByRef grd As UltraWinGrid.UltraGrid)
        With grd.ActiveCell
            If .Column.Key <> "" Then
                If .Column.Style = UltraWinGrid.ColumnStyle.EditButton _
                And .Column.Key <> "" _
                And .Column.DataType.Name = "String" _
                And .Text <> "" Then
                    Dim newvalue As String = ASCMAIN1.Format_Field(.Text, .Column.Key, , True)
                    If newvalue <> .Text Then
                        .Value = newvalue
                    End If
                End If
            End If
        End With
    End Sub

    Public Sub grdClickCellButton(
    ByRef grd As UltraWinGrid.UltraGrid,
    Optional ByVal sql_where As String = "",
    Optional ByVal commit_row As Boolean = False,
    Optional ByVal COLUMN_NAME As String = "",
    Optional ByVal VIEW_NAME As String = "")

        If grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False And grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No Then
            Exit Sub ' SO THAT VOUCHER INQ GRDAPTINVH2 DOES NOT PERMIT CLICKING THE UP ARROW
        End If
        With grd
            If .ActiveRow Is Nothing Then
            Else
                ' i don't know why we don't do these next few lines
                'If .ActiveRow.Band.Override.AllowUpdate = DefaultableBoolean.False Then
                '    Exit Sub
                'End If

                'If .ActiveRow.IsAddRow Then
                'change made to accomodate j/e
                If .ActiveRow.IsAddRow OrElse (.ActiveRow.Band.Override.AllowUpdate <> DefaultableBoolean.False AndAlso .ActiveCell.CanEnterEditMode) Then
                    Call View_Lookup(.ActiveCell, COLUMN_NAME, VIEW_NAME, , sql_where)
                    If ASCMAIN1.CodeSelector.Selections <> 0 And commit_row Then
                        .PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
                    End If
                End If
            End If
        End With
    End Sub


    Public Shared Function View_Lookup(
    ByRef txtctl As Object,
    Optional ByVal COLUMN_NAME As String = "",
    Optional ByVal VIEW_NAME As String = "",
    Optional ByVal TABLE_NAME_view As String = "",
    Optional ByVal sql_where As String = "") As String

        If COLUMN_NAME = "" Then
            If TypeOf txtctl Is UltraWinGrid.UltraGridCell Then
                COLUMN_NAME = DirectCast(txtctl, UltraWinGrid.UltraGridCell).Column.Key
            Else
                Stop
            End If
        End If

        If VIEW_NAME = "" Then
            VIEW_NAME = COLUMN_NAME
        End If

        If TABLE_NAME_view.Length = 0 AndAlso VIEW_NAME.Contains(".") Then
            ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL(
                VIEW_NAME.Split(".")(0), VIEW_NAME.Split(".")(1), sql_where)
        Else
            ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL(
                VIEW_NAME, TABLE_NAME_view, sql_where)
        End If

        'If TABLE_NAME_view = "" Then
        '    TABLE_NAME_view = TABLE_NAME
        'End If
        'ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL( _
        'VIEW_NAME, TABLE_NAME_view, sql_where)

        Dim CODE_VALUE As String = ""

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = False
            If txtctl IsNot Nothing Then
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = txtctl.Text
            End If
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                CODE_VALUE = ASCMAIN1.CodeSelector.SelectedCode

                If TypeOf txtctl Is UltraWinGrid.UltraGridCell Then
                    If DirectCast(txtctl, UltraWinGrid.UltraGridCell).Value & "" <> CODE_VALUE Then
                        DirectCast(txtctl, UltraWinGrid.UltraGridCell).Value = CODE_VALUE
                    End If
                End If
                If TypeOf txtctl Is UltraWinEditors.UltraTextEditor Then
                    If DirectCast(txtctl, UltraWinEditors.UltraTextEditor).Text & "" <> CODE_VALUE Then
                        DirectCast(txtctl, UltraWinEditors.UltraTextEditor).Text = CODE_VALUE
                    End If
                End If
            End If
        End If

        Return CODE_VALUE
    End Function

    Function Fill_Record(
    ByVal TABLE_NAME As String,
    ByVal KEY_VALUE As Object,
    Optional ByVal create_row_if_non_existent As Boolean = False,
    Optional ByVal ClearBeforeFilling As Boolean = True) As DataRow
        If KEY_VALUE Is Nothing Then
            cdr = clsASCBASE1.Fill_Record(TABLE_NAME, , create_row_if_non_existent, ClearBeforeFilling)
        Else
            cdr = clsASCBASE1.Fill_Record(TABLE_NAME, New Object() {KEY_VALUE}, create_row_if_non_existent, ClearBeforeFilling)
        End If
        Return cdr
    End Function

    Function Fill_Record(
    ByVal TABLE_NAME As String,
    Optional ByVal Parameters() As Object = Nothing,
    Optional ByVal create_row_if_non_existent As Boolean = False,
    Optional ByVal ClearBeforeFilling As Boolean = True) As DataRow

        Return clsASCBASE1.Fill_Record _
        (TABLE_NAME, Parameters, create_row_if_non_existent, ClearBeforeFilling)

    End Function

    Function Fill_Records(
    ByVal TABLE_NAME As String,
    ByVal KEY_VALUE As String,
    Optional ByVal ClearBeforeFilling As Boolean = True,
    Optional ByVal Temp_Select As String = "",
    Optional ByVal tblSubstitute As DataTable = Nothing) As Integer

        If KEY_VALUE = "" Then
            Return clsASCBASE1.Fill_Records(TABLE_NAME, , ClearBeforeFilling, Temp_Select, tblSubstitute)
        Else
            Return clsASCBASE1.Fill_Records(TABLE_NAME, New String() {KEY_VALUE}, ClearBeforeFilling, Temp_Select, tblSubstitute)
        End If

    End Function

    Function Fill_Records(
    ByVal TABLE_NAME As String,
    Optional ByVal Parameters() As Object = Nothing,
    Optional ByVal ClearBeforeFilling As Boolean = True,
    Optional ByVal Temp_Select As String = "",
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
            If Me.BindingContext.Contains(dst.Tables(TABLE_NAME)) Then
                ' Without the next 2 lines, data in text boxes in single row datatables (like header tables) will not get written to Oracle
                Dim X As CurrencyManager = Me.BindingContext(dst.Tables(TABLE_NAME))


                If ASCMAIN1.CLIENT = "INT" And TABLE_NAME = "SOTORDR2" Then
                    Try
                        X.EndCurrentEdit()

                        ' See SOFORDR1.grdSOTORDR2_Error method:
                        'grdSOTORDR2.ActiveRow.CancelUpdate()
                        ' NASTY ERROR SHOWS HERE AND IS A PRECURSOR TO ERROR WHEN CLICKING UPDATE: Column 'ORDR_NO' does not allow nulls.
                        ' CAN MAKE IT HAPPEN BY ENTERING A STYLE, COLOR, QTY, ND THEN CLICK INTO THE GRID
                        ' NOT SURE WHY CLICKING INTO THE GRID CAUSES THIS ERROR
                        ' TRIED OTHER FORMS (ICFIADJ1) AND DO NOT GET SAME BEHAVIOR
                        ' MUST BE A PROPERTY, OR CODE IN ONE OF THESE EVENT PROCEDRES

                    Catch ex As Exception
                        If MsgBox("Do you want to proceed with the Update", MsgBoxStyle.YesNo,
                              "Something went wrong at Position " & CStr(X.Position) & " when Ending Current Edit for " & TABLE_NAME) = MsgBoxResult.Yes Then
                            If ASCMAIN1.Running_in_VS Then Stop
                            X.CancelCurrentEdit()

                        Else
                            Throw New Exception(ex.Message, ex.InnerException)
                        End If
                    End Try
                Else
                    X.EndCurrentEdit()
                End If

                '  X.EndCurrentEdit()
            End If

            If AUDIT.ContainsKey(TABLE_NAME) Then
                WriteAuditTrail(TABLE_NAME)
            End If

            clsASCBASE1.Update_Record_TDA(TABLE_NAME, sql_Delete)

        End If

    End Sub

    Public Overridable Sub num_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            SendKeys.Send(Chr(9))
        End If
    End Sub

    Public Overridable Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Public Overridable Sub num_ValidationError(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinEditors.ValidationErrorEventArgs)
        e.RetainFocus = False
    End Sub

    Public Overridable Sub num_Leave(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Public Overridable Sub med_Invalidated(ByVal sender As Object, ByVal e As System.Windows.Forms.InvalidateEventArgs)

    End Sub

    Private Sub med_MaskValidationError(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinMaskedEdit.MaskValidationErrorEventArgs)
        e.RetainFocus = False
    End Sub

    Public Overridable Sub med_Validated(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Public Overridable Sub med_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)

    End Sub


    Sub Create_Summary(
    ByRef grd As UltraWinGrid.UltraGrid,
    ByVal COLUMN_NAMEs() As String,
    Optional ByVal summary_type As String = "Sum",
    Optional ByVal BandKey As String = "",
    Optional ByVal mask As String = "",
    Optional ByVal Calculator As Infragistics.Win.UltraWinGrid.ICustomSummaryCalculator = Nothing)

        For i As Integer = 0 To UBound(COLUMN_NAMEs)
            Call Create_Summary(grd, COLUMN_NAMEs(i), summary_type, BandKey, mask, Calculator)
        Next
    End Sub

    Sub Create_Summary(
    ByRef grd As UltraWinGrid.UltraGrid,
    ByVal COLUMN_NAME As String,
    Optional ByVal summary_type As String = "Sum",
    Optional ByVal BandKey As String = "",
    Optional ByVal mask As String = "",
    Optional ByVal Calculator As Infragistics.Win.UltraWinGrid.ICustomSummaryCalculator = Nothing)

        If BandKey = "" Then
            BandKey = grd.DisplayLayout.Bands(0).Key
        End If

        If mask = "" Then
            ' why can't we use the .format of the column?
            ' - 05/18/07 - I don't know, and I need it to do this now, so changing it and waiting for a good reason why not
            If summary_type = "Count" Then
                mask = "#,##0"
            Else
                mask = grd.DisplayLayout.Bands(BandKey).Columns(COLUMN_NAME).Format
            End If
            If mask = "" Then
                Select Case grd.DisplayLayout.Bands(BandKey).Columns(COLUMN_NAME).DataType.Name
                    Case "Int16", "Int32"
                        mask = "#,##0"
                    Case "Double", "Decimal"
                        mask = "#,##0.00"
                    Case "String"
                    Case Else
                        mask = "#,##0"
                End Select
            End If
        End If
        Dim DisplayFormat As String = "{0:" & mask & "}"

        Dim B As UltraWinGrid.UltraGridBand
        If BandKey = "" Then
            B = grd.DisplayLayout.Bands(0)
        Else
            B = grd.DisplayLayout.Bands(BandKey)
        End If

        With B
            If .SummaryFooterCaption = "Grand Summaries" Then
                .SummaryFooterCaption = "Totals"
            End If

            Dim summary As UltraWinGrid.SummarySettings
            Select Case summary_type
                Case "Average", "Avg"
                    summary = .Summaries.Add(UltraWinGrid.SummaryType.Average, .Columns(COLUMN_NAME))
                Case "Count"
                    summary = .Summaries.Add(UltraWinGrid.SummaryType.Count, .Columns(COLUMN_NAME))
                Case "Maximum", "Max"
                    summary = .Summaries.Add(UltraWinGrid.SummaryType.Maximum, .Columns(COLUMN_NAME))
                Case "Minimum", "Min"
                    summary = .Summaries.Add(UltraWinGrid.SummaryType.Minimum, .Columns(COLUMN_NAME))
                Case "Custom"
                    If Calculator Is Nothing Then
                        Calculator = New CustomSummary(Me, grd)
                    End If
                    summary = .Summaries.Add(UltraWinGrid.SummaryType.Custom, Calculator, .Columns(COLUMN_NAME), UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, .Columns(COLUMN_NAME))
                Case "CustomString"
                    If Calculator Is Nothing Then
                        Calculator = New CustomSummary(Me, grd, "String")
                    End If
                    summary = .Summaries.Add(UltraWinGrid.SummaryType.Custom, Calculator, .Columns(COLUMN_NAME), UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, .Columns(COLUMN_NAME))
                Case Else ' "Sum"
                    'summary = .Summaries.Add(UltraWinGrid.SummaryType.Sum, .Columns(COLUMN_NAME))
                    If grd.CalcManager Is Nothing Then
                        Dim calcManager As Infragistics.Win.UltraWinCalcManager.UltraCalcManager
                        calcManager = New Infragistics.Win.UltraWinCalcManager.UltraCalcManager(Me.Container)
                        grd.CalcManager = calcManager
                    End If

                    summary = .Summaries.Add(UltraWinGrid.SummaryType.Formula, .Columns(COLUMN_NAME))
                    summary.Formula = "Sum([" & COLUMN_NAME & "])"
            End Select
            summary.DisplayFormat = DisplayFormat
            summary.Key = COLUMN_NAME
            'summary.Appearance.BackColor = Drawing.Color.LightGray
            If .Columns(COLUMN_NAME).Level > 0 Then
                summary.SummaryPositionColumn = .Columns(COLUMN_NAME).Group.Columns(0)
            End If
            If .Columns(COLUMN_NAME).Style = UltraWinGrid.ColumnStyle.CheckBox Then
                summary.Appearance.TextHAlign = HAlign.Center
            Else

                If .Columns(COLUMN_NAME).DataType.Name = "String" And .Columns(COLUMN_NAME).CellAppearance.TextHAlign = HAlign.Default Then
                    If summary_type = "Count" Then
                        summary.Appearance.TextHAlign = HAlign.Right
                    Else
                        summary.Appearance.TextHAlign = HAlign.Left
                    End If
                Else
                    summary.Appearance.TextHAlign = .Columns(COLUMN_NAME).CellAppearance.TextHAlign
                End If
            End If

        End With
    End Sub

#Region "Excel Export"

    Overridable Function Excel_Export(ByVal grd As UltraWinGrid.UltraGrid) As GemBox.Spreadsheet.ExcelFile

        ' Gembox does not export multi banded grids
        ' Gembox has a memory constraint error when exporting large amounts of data
        If (ASCMAIN1.CLIENT = "INT" And grd.Name = "grdDPTPLANX") Or ((ASCMAIN1.CLIENT <> "INT" Or grd.Name <> "grdASTEXPT1") AndAlso grd.DisplayLayout.Bands.Count = 1) Then
            Return Gembox_Excel_Export(grd)
        Else
            'export excel spreadsheet
            Dim colheaders As Boolean = False
            If Not grd.DisplayLayout.Bands(0).ColHeadersVisible Then

                If grd.DisplayLayout.Bands(0).Groups.Count = 0 Then
                    colheaders = True
                    grd.DisplayLayout.Bands(0).ColHeadersVisible = True
                End If
            End If
            Dim x As Infragistics.Documents.Excel.Workbook = Export_to_Excel_General(grd)
            If colheaders Then
                grd.DisplayLayout.Bands(0).ColHeadersVisible = False
            End If

            Return Nothing
        End If
    End Function

    Overridable Function Export_to_Excel_General(ByVal grd As UltraWinGrid.UltraGrid)
        Return Export_to_Excel(grd)
    End Function

    Function Export_to_Excel(
    ByVal grd As UltraWinGrid.UltraGrid,
    Optional ByVal ShowExcelWorkbook As Boolean = True,
    Optional ByVal Show_HFs As Boolean = False,
    Optional ByVal CAPTION As String = "",
    Optional ByVal HeadingType As String = "A"
    ) As Infragistics.Documents.Excel.Workbook
        Return Export_to_Excel(New UltraWinGrid.UltraGrid() {grd}, ShowExcelWorkbook, Show_HFs, CAPTION, HeadingType)
    End Function

    Sub Export_to_Excel_Show(
    ByRef myWorkbook As Infragistics.Documents.Excel.Workbook,
    Optional ByVal FILE_NAME As String = "")

        Dim xlsFileName_sfx As String = ""
        Dim xlsFileName As String = ""

        If FILE_NAME = "" Then
            FILE_NAME = ASCMAIN1.ActiveForm.Name
        End If

        Do
            Try
                xlsFileName = FILE_NAME & IIf(xlsFileName_sfx = "", "", "_") & xlsFileName_sfx
                'Infragistics.Documents.Excel.BIFF8Writer.WriteWorkbookToFile(myWorkbook, ASCMAIN1.Folders("Work") & xlsFileName & ".xls")
                myWorkbook.Save(ASCMAIN1.Folders("Work") & xlsFileName & ".xls")

                xlsFileName_sfx = ""

                Dim excel As New Process
                excel.StartInfo.Arguments = """" + xlsFileName + """ /e"
                excel.StartInfo.FileName = ASCMAIN1.Folders("Work") & xlsFileName & ".xls"
                excel.Start()

            Catch ex As Exception
                xlsFileName_sfx = CStr(Val(xlsFileName_sfx) + 1)
            End Try
        Loop While xlsFileName_sfx <> "" And Val(xlsFileName_sfx) < 10

        ASCMAIN1.Progress("")

    End Sub
    Sub Export_Grids_To_Excel(ByVal expDict As Dictionary(Of UltraWinGrid.UltraGrid, String) _
                              , Optional showWorkbook As Boolean = True _
                              , Optional fileName As String = "")

        Me.Cursor = Cursors.WaitCursor
        Dim myWorkbook As New Infragistics.Documents.Excel.Workbook
        For Each kvp As KeyValuePair(Of UltraWinGrid.UltraGrid, String) In expDict
            ASCMAIN1.Progress("Exporting " & kvp.Value)
            Export_to_Excel_Add_grd(myWorkbook, kvp.Key, False, kvp.Value)
        Next
        If showWorkbook Then
            Export_to_Excel_Show(myWorkbook, fileName)
        End If
        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub
    Sub Export_to_Excel_Add_grd(
    ByRef myWorkbookFile As String,
    ByVal grd As UltraWinGrid.UltraGrid,
    Optional ByVal Show_HFs As Boolean = False,
    Optional ByVal CAPTION As String = "",
    Optional ByVal HeadingType As String = "A",
    Optional ByVal TITLE As String = "")
        Dim myWorkbook As Infragistics.Documents.Excel.Workbook = Infragistics.Documents.Excel.Workbook.Load(myWorkbookFile)
        Call Export_to_Excel_Add_grd(myWorkbook, New UltraWinGrid.UltraGrid() {grd}, Show_HFs, CAPTION, HeadingType, TITLE)
    End Sub
    Sub Export_to_Excel_Add_grd(
    ByRef myWorkbook As Infragistics.Documents.Excel.Workbook,
    ByVal grd As UltraWinGrid.UltraGrid,
    Optional ByVal Show_HFs As Boolean = False,
    Optional ByVal CAPTION As String = "",
    Optional ByVal HeadingType As String = "A",
    Optional ByVal TITLE As String = "")
        Call Export_to_Excel_Add_grd(myWorkbook, New UltraWinGrid.UltraGrid() {grd}, Show_HFs, CAPTION, HeadingType, TITLE)
    End Sub

    Sub Export_to_Excel_Add_grd(
    ByRef myWorkbook As Infragistics.Documents.Excel.Workbook,
    ByVal grd() As UltraWinGrid.UltraGrid,
    Optional ByVal Show_HFs As Boolean = False,
    Optional ByVal CAPTION As String = "",
    Optional ByVal HeadingType As String = "A",
    Optional ByVal TITLE As String = "")

        For Each grdx As UltraWinGrid.UltraGrid In grd
            If CAPTION = "" Then
                CAPTION = grdx.Text
            End If
            Call ASCMAIN1.Progress("Now Exporting " & CAPTION)

            Dim myWorksheet As Infragistics.Documents.Excel.Worksheet
            Dim TABLE_NAME As String = ""

            Try
                If TypeOf grdx.DataSource Is DataTable Then
                    TABLE_NAME = DirectCast(grdx.DataSource, DataTable).TableName
                ElseIf TypeOf grdx.DataSource Is DataSet Then
                    TABLE_NAME = grdx.DataMember
                ElseIf TypeOf grdx.DataSource Is DataView Then
                    TABLE_NAME = DirectCast(grdx.DataSource, DataView).Table.TableName
                Else
                    TABLE_NAME = grdx.Name
                End If
            Catch ex As Exception

            End Try

            Dim SHEET_NAME As String = ASCMAIN1.Excel_Sheet_Name(IIf(CAPTION <> "", CAPTION, TABLE_NAME))
            Dim SHEET_NAME_sfx As Integer = 0

            Dim found As Boolean = False
            Do
                found = False
                If myWorkbook.Worksheets.Count <> 0 Then
                    For i As Integer = 0 To myWorkbook.Worksheets.Count - 1
                        If myWorkbook.Worksheets(i).Name = SHEET_NAME Then
                            found = True
                            Exit For
                        End If
                    Next
                End If

                If found Then
                    SHEET_NAME_sfx += 1
                    SHEET_NAME = ASCMAIN1.Excel_Sheet_Name(CAPTION)
                    If Len(SHEET_NAME) > 31 - Len(CStr(SHEET_NAME_sfx)) Then
                        SHEET_NAME = Mid(SHEET_NAME, 1, 31 - Len(CStr(SHEET_NAME_sfx)))
                    End If
                    SHEET_NAME = SHEET_NAME & CStr(SHEET_NAME_sfx)
                End If
            Loop While found = True

            If SHEET_NAME = "" Then
                SHEET_NAME = "Sheet1"
            End If
            myWorksheet = myWorkbook.Worksheets.Add(SHEET_NAME)

            For i As Integer = 0 To grdx.DisplayLayout.Bands(0).Columns.Count - 1
                If i > 255 Then
                    Exit For
                End If
                myWorksheet.Columns(i).CellFormat.Font.Name = "Verdana"
            Next i

            Dim HeadingLines As Integer = 0

            Dim MENU_ITEM_DESC_NET As String = ""
            If TITLE <> "" Then
                MENU_ITEM_DESC_NET = TITLE
            Else
                If MENU_ITEM_TYPE = "T" Then
                    ' need a better way to figure out of we are exporting the audit trail to excel
                    'MENU_ITEM_DESC_NET = MENU_ITEM_DESC & IIf(Me.UltraExplorerBar1.Groups("Screen Mode").CheckedItem.Key = "Audit Trail", " - Audit Trail", "")
                Else
                    MENU_ITEM_DESC_NET = MENU_ITEM_DESC
                End If
            End If

            Select Case HeadingType
                Case "A"
                    myWorksheet.Rows(1).Cells(0).CellFormat.Font.Color = Color.Blue
                    myWorksheet.Rows(1).Cells(0).CellFormat.Font.Height = 300
                    myWorksheet.Rows(1).Cells(0).Value = MENU_ITEM_DESC_NET
                    myWorksheet.Rows(0).Cells(1).Value = TABLE_NAME
                    myWorksheet.Rows(2).Cells(0).Value = CAPTION
                    HeadingLines = 3
            End Select

            Dim hfs_items As Integer = 0
            If Show_HFs Then
                hfs_items = HFs.Count

                If HFs.Count > 0 Then
                    Dim i As Integer
                    For Each HFsKey As String In HFs.Keys
                        myWorksheet.Rows(HeadingLines + i).Cells(0).Value = HFsKey
                        myWorksheet.Rows(HeadingLines + i).Cells(2).Value = HFs(HFsKey)
                        i = i + 1
                    Next
                End If
            End If

            If HeadingType <> "" Then
                myWorksheet.Rows(0).Cells(0).CellFormat.Alignment = Infragistics.Documents.Excel.HorizontalCellAlignment.Left
                myWorksheet.Rows(0).Cells(0).CellFormat.FormatString = "mm/dd/yy;@"
                myWorksheet.Rows(0).Cells(0).Value = Now
            End If

            Dim Row0 As Integer = HeadingLines + IIf(HeadingType = "", 0, 1) + hfs_items + IIf(hfs_items = 0, 0, 1)

            ' Preserve the values of Checkbox Column Filters
            ' The Grid uses UnChecked / Checked, ABSolution uses 0 / 1
            grdColumnFilters = New List(Of strColumnFilters)
            Dim colFilterObj As Infragistics.Win.UltraWinGrid.ColumnFilter
            Dim columnName As String = String.Empty
            Dim columnFilter As New strColumnFilters

            For bands As Integer = 0 To grdx.DisplayLayout.Bands.Count - 1
                For filterNo As Integer = 0 To grdx.DisplayLayout.Bands(bands).ColumnFilters.Count - 1
                    colFilterObj = grdx.DisplayLayout.Bands(bands).ColumnFilters.Item(filterNo)
                    columnName = colFilterObj.Column.Key
                    If grdx.DisplayLayout.Bands(bands).Columns(columnName).Style = UltraWinGrid.ColumnStyle.CheckBox Then
                        If colFilterObj.FilterConditions.Count > 0 Then
                            For filterCondition As Integer = 0 To colFilterObj.FilterConditions.Count - 1
                                columnFilter = New strColumnFilters
                                columnFilter.band = bands
                                columnFilter.ColumnName = columnName
                                columnFilter.filterCondition = colFilterObj.FilterConditions(filterCondition).Clone
                                grdColumnFilters.Add(columnFilter)

                                If colFilterObj.FilterConditions(filterCondition).CompareValue.ToString = "Checked" Then
                                    colFilterObj.FilterConditions(filterCondition).CompareValue = "1"
                                ElseIf colFilterObj.FilterConditions(filterCondition).CompareValue.ToString = "Unchecked" Then
                                    colFilterObj.FilterConditions(filterCondition).CompareValue = System.DBNull.Value
                                End If

                            Next
                        End If
                    End If
                Next
            Next

            Try
                UltraGridExcelExporter1.Export(grdx, myWorksheet, Row0, 0)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Cannot Complete this Export")
                Exit Sub
            End Try


            'If grdx.DisplayLayout.Bands.Count = 1 Then
            '    For c As Integer = 1 To 100
            '        myWorksheet.Rows(3).Cells(c).Value = myWorksheet.Rows(6).Cells(c).Value
            '        myWorksheet.Rows(3).Cells(c).CellFormat.Font.Color = myWorksheet.Rows(6).Cells(c).CellFormat.Font.Color
            '        ' myWorksheet.Rows(0).Cells(0).CellFormat.Font.Height = 500

            '        'myWorksheet.Rows(3).Cells(c). = myWorksheet.Rows(6).Cells(c).CellFormat
            '    Next
            '    myWorksheet.Rows(3).Height = myWorksheet.Rows(6).Height
            'End If

            ' Reset Checkbox Column Filters
            For bands As Integer = 0 To grdx.DisplayLayout.Bands.Count - 1
                For filterNo As Integer = 0 To grdx.DisplayLayout.Bands(bands).ColumnFilters.Count - 1
                    colFilterObj = grdx.DisplayLayout.Bands(bands).ColumnFilters.Item(filterNo)
                    columnName = colFilterObj.Column.Key

                    For Each columnFilterObj As strColumnFilters In grdColumnFilters
                        If Not (columnFilterObj.band = bands AndAlso columnFilterObj.ColumnName = columnName) Then
                            Continue For
                        Else
                            For filterCondition As Integer = 0 To colFilterObj.FilterConditions.Count - 1
                                If colFilterObj.FilterConditions(filterCondition).CompareValue.ToString = "1" Then
                                    colFilterObj.FilterConditions(filterCondition).CompareValue = "Checked"
                                ElseIf colFilterObj.FilterConditions(filterCondition).CompareValue.ToString = "" Then
                                    colFilterObj.FilterConditions(filterCondition).CompareValue = "Unchecked"
                                End If
                            Next
                        End If
                    Next
                Next
            Next

            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                ASCMAIN1.sql = "Insert into ASTEXCLX values (@PARM1, GETDATE(), @PARM2, @PARM3, @PARM4, @PARM5, @PARM6, @PARM7)"
            Else
                ASCMAIN1.sql = "Insert into ASTEXCLX values (:PARM1, SYSDATE, :PARM2, :PARM3, :PARM4, :PARM5, :PARM6, :PARM7)"
            End If
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVNVVV", New Object() {ASCMAIN1.USER_ID, grdx.Name, grdx.Text, grdx.Rows.FilteredInRowCount, MENU_ITEM_OBJECT, MENU_ITEM_DESC_NET, TABLE_NAME})

            'NOTE THAT THE TROW IS WRONG FOR EXPORTS EXCEPT FOR THOSE COMING OUT OF PB REPORTS BECAUSE OF THE MANUFACTURED TOTALS ROW
            ' AND THAT IS WHY THE TOTALS FORMULA HAS BEEN DISABLED
            Dim Trow As Integer = Row0 + grdx.Rows.FilteredInRowCount

            ' Commented out by Edz on 21/1/09. Code no longer used and generates an error

            'If grdx.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand Then
            '    For i As Integer = 1 To grdx.DisplayLayout.Bands(0).Columns.Count - 1
            '        If myWorksheet.Rows(Trow).Cells(i).Value & "" <> "" Then
            '            Try
            '                Dim XCOL As String = Chr(65 + (i Mod 26))
            '                If i >= 26 Then XCOL = Chr(64 + i \ 26) & XCOL
            '                'Stop
            '                '' myWorksheet.Rows(Trow).Cells(i).ApplyFormula("=SUM(" & XCOL & CStr(Row0 + 1 + 1) & ":" & XCOL & CStr(Trow) & ")") ' ("=C5*6+1")
            '            Catch ex As Exception

            '            End Try
            '        End If

            '    Next
            'End If


            Dim XRow As Integer = Row0
            If grdx.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand Then
                If grdx.DisplayLayout.Bands(0).Groups.Count > 0 Then
                    Dim Gn As Integer = 0
                    Dim Cn As Integer = 0
                    Dim g As UltraWinGrid.UltraGridGroup = grdx.DisplayLayout.Bands(0).Groups(0).GetRelatedVisibleGroup(UltraWinGrid.VisibleRelation.First)
                    Do
                        If Not g.Hidden Then
                            myWorksheet.Rows(Row0 - 1).Cells(Cn).Value = g.Header.Caption
                            myWorksheet.Rows(Row0 - 1).Cells(Cn).CellFormat.SetFormatting(myWorksheet.Rows(Row0).Cells(Cn).CellFormat)
                            Gn += 1
                            'Cn += g.Columns.Count
                            Dim VCs As Integer = 1
                            Dim cols_by_level() As Integer
                            ReDim cols_by_level(g.Band.LevelCount - 1)

                            If g.Columns.Count > 1 Then
                                VCs = 0
                                For i As Integer = 0 To g.Columns.Count - 1
                                    If Not g.Columns(i).Hidden Then
                                        cols_by_level(g.Columns(i).Level) += 1
                                        If cols_by_level(g.Columns(i).Level) > VCs Then
                                            VCs = cols_by_level(g.Columns(i).Level)
                                        End If
                                        'VCs += 1
                                    End If
                                Next
                            End If
                            If VCs > 1 Then
                                Try

                                    Dim mergedRegion1 As Infragistics.Documents.Excel.WorksheetMergedCellsRegion =
                                        myWorksheet.MergedCellsRegions.Add(Row0 - 1, Cn, Row0 - 1, Cn + VCs - 1)
                                    mergedRegion1.CellFormat.Alignment = g.Header.Appearance.TextHAlign

                                Catch ex As Exception

                                End Try

                            End If
                            Cn += VCs
                        End If
                        g = g.GetRelatedVisibleGroup(UltraWinGrid.VisibleRelation.Next)
                    Loop While g IsNot Nothing

                    'For r As Integer = 1 To grdx.Rows.Count

                    'Next
                End If
            End If



            Dim groups As Boolean = False
            If grdx.DisplayLayout.Bands(0).SortedColumns.Count <> 0 Then
                For Each c As UltraWinGrid.UltraGridColumn In grdx.DisplayLayout.Bands(0).SortedColumns
                    If c.IsGroupByColumn Then
                        groups = True
                        Exit For
                    End If
                Next
            End If

            If groups Then
                With myWorksheet
                    .Rows(3).Height = .Rows(6).Height
                    For I As Integer = 0 To .Rows(6).Cells.Count - 1
                        .Rows(3).Cells(I).Value = .Rows(6).Cells(I).Value
                        .Rows(3).Cells(I).CellFormat.SetFormatting(.Rows(6).Cells(I).CellFormat)
                    Next
                End With
            End If
        Next
    End Sub

    Sub Export_to_Excel_Add_grd_to_Sheet(
 ByVal grdx As UltraWinGrid.UltraGrid,
 ByVal row0 As Int32,
 ByVal col0 As Int32)

        UltraGridExcelExporter1.Export(grdx, myWorkSheet, row0, col0)
        'Dim Trow As Integer = row0 + grdx.Rows.Count
        'If grdx.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand Then
        '    For i As Integer = 1 To grdx.DisplayLayout.Bands(0).Columns.Count - 1
        '        If myWorksheet.Rows(Trow).Cells(i).Value & "" <> "" Then
        '            Try
        '                Dim XCOL As String = Chr(65 + (i Mod 26))
        '                If i >= 26 Then XCOL = Chr(64 + i \ 26) & XCOL
        '            Catch ex As Exception

        '            End Try
        '        End If

        '    Next
        'End If

        myWorkSheet.Rows(row0 - 1).Cells(col0).Value = grdx.Text

        Dim XRow As Integer = row0
        If grdx.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand Then
            If grdx.DisplayLayout.Bands(0).Groups.Count > 0 Then
                Dim Gn As Integer = 0
                Dim Cn As Integer = 0
                Dim g As UltraWinGrid.UltraGridGroup = grdx.DisplayLayout.Bands(0).Groups(0).GetRelatedVisibleGroup(UltraWinGrid.VisibleRelation.First)
                Do
                    If Not g.Hidden Then
                        myWorkSheet.Rows(row0 - 1).Cells(Cn).Value = g.Header.Caption
                        myWorkSheet.Rows(row0 - 1).Cells(Cn).CellFormat.SetFormatting(myWorkSheet.Rows(row0).Cells(Cn).CellFormat)
                        Gn += 1
                        'Cn += g.Columns.Count
                        Dim VCs As Integer = 1
                        If g.Columns.Count > 1 Then
                            VCs = 0
                            For i As Integer = 0 To g.Columns.Count - 1
                                If Not g.Columns(i).Hidden Then
                                    VCs += 1
                                End If
                            Next
                        End If
                        If VCs > 1 Then
                            Try

                                Dim mergedRegion1 As Infragistics.Documents.Excel.WorksheetMergedCellsRegion =
                                    myWorkSheet.MergedCellsRegions.Add(row0 - 1, Cn, row0 - 1, Cn + VCs - 1)
                                mergedRegion1.CellFormat.Alignment = g.Header.Appearance.TextHAlign
                            Catch ex As Exception

                            End Try
                        End If
                        Cn += VCs
                    End If
                    g = g.GetRelatedVisibleGroup(UltraWinGrid.VisibleRelation.Next)
                Loop While g IsNot Nothing

            End If
        End If

    End Sub

    Function Export_to_Excel(
    ByVal grd() As UltraWinGrid.UltraGrid,
    Optional ByVal ShowExcelWorkbook As Boolean = True,
    Optional ByVal Show_HFs As Boolean = False,
    Optional ByVal CAPTION As String = "",
    Optional ByVal HeadingType As String = "A"
    ) As Infragistics.Documents.Excel.Workbook

        Me.Cursor = Cursors.WaitCursor
        Dim myWorkbook As New Infragistics.Documents.Excel.Workbook

        Call Export_to_Excel_Add_grd(myWorkbook, grd, Show_HFs, CAPTION, HeadingType)

        If ShowExcelWorkbook Then
            Call Export_to_Excel_Show(myWorkbook)
        End If

        Call ASCMAIN1.Progress("")

        Me.Cursor = Cursors.Default
        Return myWorkbook

    End Function
#End Region

#Region "Excel Export - Using Gembox"

    Overridable Sub Gembox_Excel_Export_grd(ByVal grd As UltraWinGrid.UltraGrid)
        Gembox_Export_to_Excel_General(grd)
    End Sub

    Overridable Function Gembox_Excel_Export(ByVal grd As UltraWinGrid.UltraGrid) As GemBox.Spreadsheet.ExcelFile
        Return Gembox_Export_to_Excel_General(grd)
    End Function

    Overridable Function Gembox_Export_to_Excel_General(ByVal grd As UltraWinGrid.UltraGrid) As GemBox.Spreadsheet.ExcelFile
        Return Gembox_Export_to_Excel(grd)
    End Function

    Function Gembox_Export_to_Excel(
    ByVal grd As UltraWinGrid.UltraGrid,
    Optional ByVal ShowExcelWorkbook As Boolean = True,
    Optional ByVal Show_HFs As Boolean = False,
    Optional ByVal CAPTION As String = "",
    Optional ByVal HeadingType As String = "A"
    ) As GemBox.Spreadsheet.ExcelFile
        Return Gembox_Export_to_Excel(New UltraWinGrid.UltraGrid() {grd}, ShowExcelWorkbook, Show_HFs, CAPTION, HeadingType)
    End Function

    Function Gembox_Export_to_Excel_Show(
    ByRef myWorkbook As GemBox.Spreadsheet.ExcelFile,
    Optional ByVal FILE_NAME As String = "",
    Optional ByVal xls_type As String = ".xlsx") As String

        Dim xlsFileName As String = GetFileName(myWorkbook, FILE_NAME, xls_type)

        Dim excel As New Process

        Try
            excel.StartInfo.Arguments = """" + xlsFileName + """ /e"
            excel.StartInfo.FileName = ASCMAIN1.Folders("Work") & xlsFileName & xls_type ' ".xls"
            excel.Start()
        Catch ex As Exception

        End Try

        ASCMAIN1.Progress("")

        Return ASCMAIN1.Folders("Work") & xlsFileName & xls_type
        'Return excel

    End Function

    Function GetFileName(
    ByRef myWorkbook As GemBox.Spreadsheet.ExcelFile,
    Optional ByVal FILE_NAME As String = "",
    Optional ByVal xls_type As String = ".xls") As String

        Dim xlsFileName_sfx As String = ""
        Dim xlsFileName As String = ""

        If FILE_NAME = "" Then
            FILE_NAME = ASCMAIN1.ActiveForm.Name
        End If

        Do
            Try
                xlsFileName = FILE_NAME & IIf(xlsFileName_sfx = "", "", "_") & xlsFileName_sfx
                myWorkbook.Save(ASCMAIN1.Folders("Work") & xlsFileName & xls_type) ' ".xls")

                'myWorkbook.SaveXls(ASCMAIN1.Folders("Work") & xlsFileName & ".xls")
                'myWorkbook.ClosePreservedXlsx()

                myWorkbook = Nothing
                xlsFileName_sfx = ""

            Catch ex As Exception
                xlsFileName_sfx = CStr(Val(xlsFileName_sfx) + 1)
            End Try
        Loop While xlsFileName_sfx <> "" And Val(xlsFileName_sfx) < 10

        Try
            Gembox_Excel_Formatting(xlsFileName & xls_type) ' ".xls")
        Catch ex As Exception

        End Try

        Return xlsFileName

    End Function

    Sub Gembox_Excel_Formatting(ByVal xlsFilename As String)

        Dim excel As New Microsoft.Office.Interop.Excel.Application
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = excel.Workbooks.Open(ASCMAIN1.Folders("Work") & xlsFilename)

        If dst.Tables.Contains("ASTGRIDC") Then
            If dst.Tables("ASTGRIDC").Rows.Count <> 0 Then

                Dim SHEET As Int32 = 0
                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing

                For Each rowASTGRIDC As DataRow In dst.Tables("ASTGRIDC").Select("", "SHEET")
                    If rowASTGRIDC.Item("SHEET") <> SHEET Then
                        SHEET = rowASTGRIDC.Item("SHEET")
                        ws = wb.Sheets(SHEET)

                        With ws.Range(ws.Cells(2, 1), ws.Cells(2, 1))
                            With .Interior
                                .Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternLinearGradient
                                .Gradient.Degree = 45
                                .Gradient.ColorStops.Clear()
                            End With
                            With .Interior.Gradient.ColorStops.Add(0)
                                .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                                .TintAndShade = 0
                            End With
                            With .Interior.Gradient.ColorStops.Add(1)
                                .ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark2
                                .TintAndShade = -0.498031556138798
                            End With
                        End With

                    End If

                    Dim ROW As Int32 = Val(rowASTGRIDC.Item("ROW"))
                    Dim COL As Int32 = Val(rowASTGRIDC.Item("COL"))

                    With ws.Range(ws.Cells(ROW + 1, COL + 1), ws.Cells(ROW + 1, COL + 1))
                        Dim COLOR1 As Double = Val(rowASTGRIDC.Item("COLOR1")) ' .Interior.Color  ' .Interior.PatternColor
                        Dim COLOR2 As Double = Val(rowASTGRIDC.Item("COLOR2")) '  211 * 256 * 256 + 211 * 256 + 211 ' LIGHTGREY RGB
                        Dim GRADIENT As Int32 = rowASTGRIDC.Item("GRADIENT")
                        .WrapText = True
                        With .Interior
                            Select Case GRADIENT
                                Case 2
                                    .Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternLinearGradient
                                    .Gradient.Degree = 90
                                Case 3
                                    .Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternLinearGradient
                                    .Gradient.Degree = 0
                                Case 4
                                    .Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternLinearGradient
                                    .Gradient.Degree = 135
                                Case 5
                                    .Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternLinearGradient
                                    .Gradient.Degree = 45
                                Case Else
                                    .Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternLinearGradient
                                    .Gradient.Degree = 90
                            End Select

                            '.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternLinearGradient
                            '.Gradient.Degree = 90
                            .Gradient.ColorStops.Clear()
                        End With
                        With .Interior.Gradient.ColorStops.Add(0)
                            '.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                            .Color = COLOR1
                            .TintAndShade = 0
                        End With
                        With .Interior.Gradient.ColorStops.Add(1)
                            .Color = COLOR2
                            .TintAndShade = 0
                        End With

                    End With
                Next
            End If
        End If


        excel.DisplayAlerts = False
        wb.Save()
        excel.Quit()
        excel = Nothing
    End Sub

    Sub Gembox_Export_to_Excel_Add_grd(
    ByRef myWorkbook As GemBox.Spreadsheet.ExcelFile,
    ByVal grd As UltraWinGrid.UltraGrid,
    Optional ByVal Show_HFs As Boolean = False,
    Optional ByVal CAPTION As String = "",
    Optional ByVal HeadingType As String = "A",
    Optional ByVal TITLE As String = "")
        Gembox_Export_to_Excel_Add_grd(myWorkbook, New UltraWinGrid.UltraGrid() {grd}, Show_HFs, CAPTION, HeadingType, TITLE)
    End Sub

    Sub Gembox_Export_to_Excel_Add_grd(
    ByRef myWorkbook As GemBox.Spreadsheet.ExcelFile,
    ByVal grd() As UltraWinGrid.UltraGrid,
    Optional ByVal Show_HFs As Boolean = False,
    Optional ByVal CAPTION As String = "",
    Optional ByVal HeadingType As String = "A",
    Optional ByVal TITLE As String = "")

        For Each grdx As UltraWinGrid.UltraGrid In grd
            If CAPTION = "" Then
                CAPTION = grdx.Text
            End If
            ASCMAIN1.Progress("Now Exporting " & CAPTION)


            Dim TABLE_NAME As String = ""
            Dim rowsInDataSource As Integer = 0
            Dim columnsInDataSource As Integer = 0

            Try
                If TypeOf grdx.DataSource Is DataTable Then
                    TABLE_NAME = DirectCast(grdx.DataSource, DataTable).TableName
                    rowsInDataSource = DirectCast(grdx.DataSource, DataTable).Rows.Count
                    columnsInDataSource = DirectCast(grdx.DataSource, DataTable).Columns.Count
                    If Not gemboxUseXlsx And columnsInDataSource > 255 Then
                        gemboxUseXlsx = True
                    End If
                ElseIf TypeOf grdx.DataSource Is DataSet Then
                    TABLE_NAME = grdx.DataMember
                ElseIf TypeOf grdx.DataSource Is DataView Then
                    TABLE_NAME = DirectCast(grdx.DataSource, DataView).Table.TableName
                Else
                    TABLE_NAME = grdx.Name
                End If
            Catch ex As Exception

            End Try

            Dim SHEET_NAME As String = ASCMAIN1.Excel_Sheet_Name(IIf(CAPTION <> "", CAPTION, TABLE_NAME))
            Dim SHEET_NAME_sfx As Integer = 0

            Dim found As Boolean = False
            Do
                found = False
                If myWorkbook.Worksheets.Count <> 0 Then
                    For i As Integer = 0 To myWorkbook.Worksheets.Count - 1
                        If myWorkbook.Worksheets(i).Name = SHEET_NAME Then
                            found = True
                            Exit For
                        End If
                    Next
                End If

                If found Then
                    SHEET_NAME_sfx += 1
                    SHEET_NAME = ASCMAIN1.Excel_Sheet_Name(CAPTION)
                    If Len(SHEET_NAME) > 31 - Len(CStr(SHEET_NAME_sfx)) Then
                        SHEET_NAME = Mid(SHEET_NAME, 1, 31 - Len(CStr(SHEET_NAME_sfx)))
                    End If
                    SHEET_NAME = SHEET_NAME & CStr(SHEET_NAME_sfx)
                End If
            Loop While found = True

            If SHEET_NAME = "" Then
                SHEET_NAME = "Sheet1"
            End If
            Dim myWorksheet As GemBox.Spreadsheet.ExcelWorksheet = myWorkbook.Worksheets.Add(SHEET_NAME)
            Dim HeadingLines As Integer = 0

            With myWorksheet.PrintOptions
                ' Necessary since upgrade to 3.7 since the default causes A4 to be used
                .PaperType = GemBox.Spreadsheet.PaperType.Letter
                '.FitWorksheetWidthToPages = 1
                '.Portrait = False
                '.PrintGridlines = True

                '.BottomMargin = 0.25
                '.TopMargin = 0.25
                '.LeftMargin = 0.25
                '.RightMargin = 0.25
            End With

            Dim Display_Group_Headers_at_Col0 As Boolean = True
            Dim Display_Group_Headers_on_Left As Boolean = False

            Dim gbyCols As New List(Of UltraWinGrid.UltraGridColumn)
            If grdx.DisplayLayout.Bands(0).SortedColumns.Count > 0 Then
                For i As Integer = 0 To grdx.DisplayLayout.Bands(0).SortedColumns.Count - 1
                    Dim gcol As UltraWinGrid.UltraGridColumn = grdx.DisplayLayout.Bands(0).SortedColumns(i)
                    If gcol.IsGroupByColumn Then
                        gbyCols.Add(gcol)
                    End If
                Next
            End If

            Dim Col0 As Integer = 0
            If Display_Group_Headers_on_Left Then
                Col0 = gbyCols.Count
            End If

            Dim MENU_ITEM_DESC_NET As String = ""
            If TITLE <> "" Then
                MENU_ITEM_DESC_NET = TITLE
            Else
                If MENU_ITEM_TYPE = "T" Then
                    ' need a better way to figure out of we are exporting the audit trail to excel
                    'MENU_ITEM_DESC_NET = MENU_ITEM_DESC & IIf(Me.UltraExplorerBar1.Groups("Screen Mode").CheckedItem.Key = "Audit Trail", " - Audit Trail", "")
                    MENU_ITEM_DESC_NET = MENU_ITEM_DESC
                Else
                    MENU_ITEM_DESC_NET = MENU_ITEM_DESC
                End If
            End If

            Select Case HeadingType
                Case "A"
                    myWorksheet.Cells(1, Col0).Style.Font.Color = Color.Blue
                    myWorksheet.Cells(1, Col0).Style.Font.Size = 300
                    myWorksheet.Cells(1, Col0).Style.Font.Name = "Times New Roman"
                    myWorksheet.Cells(1, Col0).Value = MENU_ITEM_DESC_NET
                    myWorksheet.Cells(0, Col0 + 1).Value = TABLE_NAME
                    myWorksheet.Cells(2, Col0).Value = CAPTION
                    HeadingLines = 3
            End Select

            Dim hfs_items As Integer = 0
            If Show_HFs Then
                hfs_items = HFs.Count

                If HFs.Count > 0 Then
                    Dim i As Integer
                    For Each HFsKey As String In HFs.Keys
                        myWorksheet.Rows(HeadingLines + i).Cells(Col0).Value = HFsKey
                        myWorksheet.Rows(HeadingLines + i).Cells(Col0 + 2).Value = HFs(HFsKey)
                        i = i + 1
                    Next
                End If
            End If

            If HeadingType <> "" Then
                myWorksheet.Rows(0).Cells(Col0 + 0).Style.HorizontalAlignment = HorizontalAlignmentStyle.Left
                myWorksheet.Rows(0).Cells(Col0 + 0).Style.NumberFormat = "mm/dd/yy;@"
                myWorksheet.Rows(0).Cells(Col0 + 0).Value = Now
            End If

            Dim Row0 As Integer = HeadingLines + IIf(HeadingType = "", 0, 1) + hfs_items + IIf(hfs_items = 0, 0, 1)

            If Col0 > 0 Then
                For i As Integer = 0 To Col0 - 1
                    myWorksheet.Columns(i).Width = 20
                Next
            End If

            Gembox_Export(grdx, myWorksheet, Row0, Col0, gbyCols, Display_Group_Headers_at_Col0, Display_Group_Headers_on_Left)

            Gembox_Export_Custom_Post_Processing(grdx, myWorksheet)

            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                ASCMAIN1.sql = "Insert into ASTEXCLX values (@PARM1, GETDATE(), @PARM2, @PARM3, @PARM4, @PARM5, @PARM6, @PARM7)"
            Else
                ASCMAIN1.sql = "Insert into ASTEXCLX values (:PARM1, SYSDATE, :PARM2, :PARM3, :PARM4, :PARM5, :PARM6, :PARM7)"
            End If
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVNVVV", New Object() {ASCMAIN1.USER_ID, grdx.Name, grdx.Text, grdx.Rows.FilteredInRowCount, MENU_ITEM_OBJECT, MENU_ITEM_DESC_NET, TABLE_NAME})

        Next
    End Sub

    Overridable Sub Gembox_Export_Custom_Post_Processing(ByVal grd As UltraWinGrid.UltraGrid, ByVal ws As GemBox.Spreadsheet.ExcelWorksheet)

    End Sub

    Sub Gembox_Export_to_Excel_Add_grd_to_Sheet(
 ByVal grdx As UltraWinGrid.UltraGrid,
 ByVal row0 As Int32,
 ByVal col0 As Int32,
 ByVal myWorkSheet As GemBox.Spreadsheet.ExcelWorksheet)

        myWorkSheet.Rows(row0 - 1).Cells(col0).Value = grdx.Text
        Dim XRow As Integer = row0
        If grdx.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand Then
            If grdx.DisplayLayout.Bands(0).Groups.Count > 0 Then
                Dim Gn As Integer = 0
                Dim Cn As Integer = 0
                Dim g As UltraWinGrid.UltraGridGroup = grdx.DisplayLayout.Bands(0).Groups(0).GetRelatedVisibleGroup(UltraWinGrid.VisibleRelation.First)
                Do
                    If Not g.Hidden Then
                        myWorkSheet.Rows(row0 - 1).Cells(Cn).Value = g.Header.Caption
                        myWorkSheet.Rows(row0 - 1).Cells(Cn).Style = myWorkSheet.Rows(row0).Cells(Cn).Style
                        Gn += 1
                        Dim VCs As Integer = 1
                        If g.Columns.Count > 1 Then
                            VCs = 0
                            For i As Integer = 0 To g.Columns.Count - 1
                                If Not g.Columns(i).Hidden Then
                                    VCs += 1
                                End If
                            Next
                        End If
                        If VCs > 1 Then
                            Try
                                'Dim mergedRegion1 As Infragistics.Documents.Excel.WorksheetMergedCellsRegion = _
                                '    myWorkSheet.MergedCellsRegions.Add(row0 - 1, Cn, row0 - 1, Cn + VCs - 1)
                                'mergedRegion1.CellFormat.Alignment = g.Header.Appearance.TextHAlign
                            Catch ex As Exception

                            End Try
                        End If
                        Cn += VCs
                    End If
                    g = g.GetRelatedVisibleGroup(UltraWinGrid.VisibleRelation.Next)
                Loop While g IsNot Nothing

            End If
        End If

    End Sub

    Function Gembox_Export_to_Excel(
    ByVal grd() As UltraWinGrid.UltraGrid,
    Optional ByVal ShowExcelWorkbook As Boolean = True,
    Optional ByVal Show_HFs As Boolean = False,
    Optional ByVal CAPTION As String = "",
    Optional ByVal HeadingType As String = "A"
    ) As GemBox.Spreadsheet.ExcelFile

        Me.Cursor = Cursors.WaitCursor
        SpreadsheetInfo.SetLicense(ASCMAIN1.GemboxKey)

        Dim myWorkbook As New GemBox.Spreadsheet.ExcelFile

        If dst.Tables.Contains("ASTGRIDC") Then
            dst.Tables("ASTGRIDC").Rows.Clear()
        Else
            With dst.Tables.Add("ASTGRIDC")
                .Columns.Add("SHEET", GetType(System.Int32))
                .Columns.Add("ROW", GetType(System.Int32))
                .Columns.Add("COL", GetType(System.Int32))
                .Columns.Add("COLOR1", GetType(System.Int64))
                .Columns.Add("COLOR2", GetType(System.Int64))
                .Columns.Add("GRADIENT", GetType(System.Int32))
            End With
        End If

        Gembox_Export_to_Excel_Add_grd(myWorkbook, grd, Show_HFs, CAPTION, HeadingType)

        If ShowExcelWorkbook Then
            'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = ASFMAIN1.UltraToolbarsManager1.Tools("Publish Documents to Portal")
            'If tlb_sbt.Checked Then
            '    Dim FF As New ASFSRPTV
            '    FF.publish_documents = True
            '    Dim xlsFileName As String = GetFileName(myWorkbook)
            '    FF.FILENAME_to_export = xlsFileName
            '    FF.ExportFormat = "XLS"
            '    FF.Show()
            'Else
            '    Gembox_Export_to_Excel_Show(myWorkbook)
            'End If
            Dim FILENAME As String = Gembox_Export_to_Excel_Show(myWorkbook)
            FILENAMEs_to_Publish.Add(FILENAME)

        End If

        ASCMAIN1.Progress("")

        Me.Cursor = Cursors.Default
        Return myWorkbook
    End Function

    Sub Gembox_Export_Paint_Column(
    ByVal grd As UltraWinGrid.UltraGrid,
    ByVal ws As GemBox.Spreadsheet.ExcelWorksheet,
    ByVal gcol As UltraWinGrid.UltraGridColumn,
    ByVal r As Int16,
    ByVal col As Int16)

        Dim row As DataRow = dst.Tables("ASTGRIDC").NewRow
        row.Item("SHEET") = 1 ' ws.Name
        row.Item("ROW") = r
        row.Item("COL") = col
        dst.Tables("ASTGRIDC").Rows.Add(row)

        ws.Cells(r, col).Value = gcol.Header.Caption
        ws.Cells(r, col).Style.Font.Name = grd.Font.Name
        If Not gcol.Header.Appearance.ForeColor = Color.Empty Then
            ws.Cells(r, col).Style.Font.Color = gcol.Header.Appearance.ForeColor
        End If
        If Not gcol.Header.Appearance.BackColor = Color.Empty Then
            ws.Cells(r, col).Style.FillPattern.PatternForegroundColor = gcol.Header.Appearance.BackColor
        ElseIf Not gcol.Header.Appearance.BackColor2 = Color.Empty Then
            ws.Cells(r, col).Style.FillPattern.PatternForegroundColor = gcol.Header.Appearance.BackColor2
        Else
            ws.Cells(r, col).Style.FillPattern.PatternForegroundColor = Color.LightGray
        End If
        ws.Cells(r, col).Style.FillPattern.PatternStyle = FillPatternStyle.Solid

        Dim COLOR1 As Color = gcol.Header.Appearance.BackColor
        If COLOR1 = Color.Empty Then COLOR1 = Color.White
        Dim COLOR2 As Color = gcol.Header.Appearance.BackColor2
        If COLOR2 = Color.Empty Then COLOR2 = Color.Gray ' Color.FromArgb(222, 223, 206) ' Color.Gray ' Color.LightGray 'Color.FromArgb(222, 223, 200) ' 

        row.Item("COLOR1") = COLOR1.B * 256 * 256 + COLOR1.G * 256 + COLOR1.R
        row.Item("COLOR2") = COLOR2.B * 256 * 256 + COLOR2.G * 256 + COLOR2.R
        row.Item("GRADIENT") = gcol.Header.Appearance.BackGradientStyle

        ws.Columns(col).Width = gcol.Width * 2340 / 130 * 2
        ws.Cells(r, col).Style.VerticalAlignment = VerticalAlignmentStyle.Center
        If gcol.Header.Appearance.TextHAlign = HAlign.Right Then
            ws.Cells(r, col).Style.HorizontalAlignment = HorizontalAlignmentStyle.Right
        ElseIf gcol.Header.Appearance.TextHAlign = HAlign.Center Then
            ws.Cells(r, col).Style.HorizontalAlignment = HorizontalAlignmentStyle.Center
        End If
    End Sub

    Sub Gembox_Export(
    ByVal grd As UltraWinGrid.UltraGrid,
    ByVal ws As GemBox.Spreadsheet.ExcelWorksheet,
    ByVal Row0 As Int64,
    ByVal Col0 As Int64,
    ByVal gbyCols As List(Of UltraWinGrid.UltraGridColumn),
    ByVal Display_Group_Headers_at_Col0 As Boolean,
    ByVal Display_Group_Headers_on_Left As Boolean)

        If dst.Tables.Contains("ASTGRIDB") Then
            dst.Tables("ASTGRIDB").Rows.Clear()
        Else
            With dst.Tables.Add("ASTGRIDB")
                .Columns.Add("BAND", GetType(System.Int32))
                .Columns.Add("COLUMN", GetType(System.Int32))
                .Columns.Add("ROW", GetType(System.Int32))
                .Columns.Add("COL", GetType(System.Int32))
                .Columns.Add("GROUP", GetType(System.Int32))
                .Columns.Add("GHVP", GetType(System.Int32))
                .Columns.Add("CHVP", GetType(System.Int32))
                .Columns.Add("COLUMN_NAME")
                .Columns.Add("LVL", GetType(System.Int32))
                .Columns.Add("COLSPAN", GetType(System.Int32))
                .PrimaryKey = New DataColumn() { .Columns("BAND"), .Columns("COLUMN")}
            End With
        End If

        Dim r As Integer = Row0
        Dim col_max As Integer = 0
        For Each b As UltraWinGrid.UltraGridBand In grd.DisplayLayout.Bands
            With b
                For i As Integer = 0 To .Columns.Count - 1
                    If Not .Columns(i).Hidden And (.Groups.Count = 0 OrElse (.Columns(i).Group IsNot Nothing AndAlso Not .Columns(i).Group.Hidden)) Then
                        Dim rowDT As DataRow = dst.Tables("ASTGRIDB").NewRow
                        rowDT.Item("BAND") = .Index
                        rowDT.Item("COLUMN") = i
                        Dim ROW As Int32 = 0
                        If .Columns(i).Group IsNot Nothing Then
                            ROW = .Columns(i).Level
                            rowDT.Item("GROUP") = .Columns(i).Group.Index
                            rowDT.Item("GHVP") = .Columns(i).Group.Header.VisiblePosition
                        End If
                        Dim COL As Int32 = .Columns(i).Header.VisiblePosition
                        rowDT.Item("CHVP") = .Columns(i).Header.VisiblePosition
                        If .Columns(i).Group IsNot Nothing Then
                            COL += .Columns(i).Group.Header.VisiblePosition
                        End If
                        If COL > col_max Then
                            col_max = COL
                        End If
                        rowDT.Item("ROW") = ROW
                        rowDT.Item("COL") = COL
                        rowDT.Item("COLUMN_NAME") = .Columns(i).Key
                        rowDT.Item("LVL") = .Columns(i).Level
                        rowDT.Item("COLSPAN") = .Columns(i).ColSpan
                        dst.Tables("ASTGRIDB").Rows.Add(rowDT)
                    End If
                Next

                If .Groups.Count <> 0 Then
                    For g As Integer = 0 To .Groups.Count - 1
                        If Not .Groups(g).Hidden Then
                            For l As Integer = 0 To .LevelCount - 1
                                Dim CHVP As Integer = -1
                                For Each rowDT As DataRow In dst.Tables("ASTGRIDB").Select _
                                    ("BAND = " & CStr(.Index) & " AND GROUP = " & CStr(g) & " AND ROW = " & CStr(l), "CHVP")
                                    Dim COL As Integer = Val(rowDT.Item("COL") & "")
                                    COL -= Val(rowDT.Item("CHVP") & "")
                                    CHVP += 1
                                    COL += CHVP
                                    rowDT.Item("COL") = COL
                                    'rowDT.Item("CHVP") = CHVP
                                Next
                            Next
                        End If
                    Next
                End If

            End With

            Exit For
        Next

        ' Print Headings for a Band (currently hard coded for Band 0 only)

        With grd.DisplayLayout.Bands(0)

            col_max = -1
            Dim col_max_prev_group As Integer = 0
            Dim sqlw As String = ""
            If .Groups.Count <> 0 Then
                sqlw = "GHVP IS NOT NULL"
            End If
            Dim GHVP As Int32 = -1
            Dim col_lvl As Int32 = 0
            Dim row_lvl As Int32 = 0
            For Each rowDT As DataRow In dst.Tables("ASTGRIDB").Select(sqlw, "GHVP,ROW,CHVP,COL")
                If Val(rowDT.Item("GHVP") & "") <> GHVP Then
                    GHVP = Val(rowDT.Item("GHVP") & "")
                    col_max_prev_group = col_max
                    row_lvl = -1
                End If
                If Val(rowDT.Item("ROW") & "") <> row_lvl Then
                    row_lvl = Val(rowDT.Item("ROW") & "")
                    col_lvl = col_max_prev_group + 1
                End If
                If col_lvl > col_max Then
                    col_max = col_lvl
                End If
                rowDT.Item("COL") = col_lvl
                col_lvl += Val(rowDT.Item("COLSPAN") & "")
            Next

            r += 1
            If .Groups.Count = 0 Then
                For i As Integer = 0 To .Columns.Count - 1
                    If Not .Columns(i).Hidden Then
                        Dim rowDT As DataRow = dst.Tables("ASTGRIDB").Rows.Find(New Object() {0, i}) ' 0 is hard-coded band 0
                        Dim col As Integer = Val(rowDT.Item("COL") & "") + Col0
                        Gembox_Export_Paint_Column(grd, ws, .Columns(i), r, col)
                    End If
                Next
                ws.Rows(r).Height = .ColHeaderLines * 255 * 1.5
                ws.Rows(r).Style.VerticalAlignment = VerticalAlignmentStyle.Center
            Else
                Dim col As Integer = Col0 + 0
                For i As Integer = 0 To .Groups.Count - 1
                    If Not .Groups(i).Hidden Then
                        Dim max_cols_in_level As Int16 = 0
                        Dim min_colspan_in_group As Int16 = 0
                        Dim glvls() As String
                        ReDim glvls(.LevelCount - 1)
                        For Each gcol As UltraWinGrid.UltraGridColumn In .Groups(i).Columns
                            If Not gcol.Hidden Then
                                Dim lvl As Int16 = gcol.Level
                                glvls(lvl) &= CStr(gcol.ColSpan)
                                If glvls(lvl).Length > max_cols_in_level Then max_cols_in_level = glvls(lvl).Length
                                If gcol.ColSpan < min_colspan_in_group Or min_colspan_in_group = 0 Then min_colspan_in_group = gcol.ColSpan
                            End If
                        Next

                        'If .Groups(i).Header.Caption = "Totals" Then Stop
                        ws.Cells(r, col).Value = .Groups(i).Header.Caption
                        ws.Cells(r, col).Style.Font.Name = grd.Font.Name

                        Dim row As DataRow = dst.Tables("ASTGRIDC").NewRow
                        row.Item("SHEET") = 1 ' ws.Name
                        row.Item("ROW") = r
                        row.Item("COL") = col
                        dst.Tables("ASTGRIDC").Rows.Add(row)

                        If Not .Groups(i).Header.Appearance.ForeColor = Color.Empty Then
                            ws.Cells(r, col).Style.Font.Color = .Groups(i).Header.Appearance.ForeColor
                        End If
                        If Not .Groups(i).Header.Appearance.BackColor = Color.Empty Then
                            ws.Cells(r, col).Style.FillPattern.PatternForegroundColor = .Groups(i).Header.Appearance.BackColor
                        ElseIf Not .Groups(i).Header.Appearance.BackColor2 = Color.Empty Then
                            ws.Cells(r, col).Style.FillPattern.PatternForegroundColor = .Groups(i).Header.Appearance.BackColor2
                        Else
                            ws.Cells(r, col).Style.FillPattern.PatternForegroundColor = Color.LightGray
                        End If

                        ws.Cells(r, col).Style.FillPattern.PatternStyle = FillPatternStyle.Solid


                        Dim COLOR1 As Color = .Groups(i).Header.Appearance.BackColor
                        If COLOR1 = Color.Empty Then COLOR1 = Color.White
                        Dim COLOR2 As Color = .Groups(i).Header.Appearance.BackColor2
                        If COLOR2 = Color.Empty Then COLOR2 = Color.Gray ' Color.FromArgb(222, 223, 206) ' Color.Gray ' Color.LightGray 'Color.FromArgb(222, 223, 200) ' 

                        'If .Groups(i).Header.Caption = "Totals" Then Stop

                        row.Item("COLOR1") = COLOR1.B * 256 * 256 + COLOR1.G * 256 + COLOR1.R
                        row.Item("COLOR2") = COLOR2.B * 256 * 256 + COLOR2.G * 256 + COLOR2.R
                        row.Item("GRADIENT") = .Groups(i).Header.Appearance.BackGradientStyle

                        ws.Columns(col).Width = .Groups(i).Width * 2340 / 130 * 2
                        If .Groups(i).Header.Appearance.TextHAlign = HAlign.Right Then
                            ws.Cells(r, col).Style.HorizontalAlignment = HorizontalAlignmentStyle.Right
                        ElseIf .Groups(i).Header.Appearance.TextHAlign = HAlign.Center Then
                            ws.Cells(r, col).Style.HorizontalAlignment = HorizontalAlignmentStyle.Center
                        End If

                        If max_cols_in_level > 1 Then
                            ws.Cells.GetSubrangeRelative(r, col, max_cols_in_level, 1).Merged = True
                            With ws.Cells.GetSubrangeRelative(r, col, max_cols_in_level, 1)
                                .Style.VerticalAlignment = VerticalAlignmentStyle.Center
                                '.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center
                            End With
                        Else
                            With ws.Cells(r, col)
                                .Style.VerticalAlignment = VerticalAlignmentStyle.Center
                                '.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center
                            End With
                        End If

                        If .ColHeadersVisible Then
                            For Each gcol As UltraWinGrid.UltraGridColumn In .Groups(i).Columns
                                If Not gcol.Hidden Then
                                    Dim lvl As Int16 = gcol.Level
                                    'Gembox_Export_Paint_Column(grd, ws, .Columns(i), r, col)

                                    Dim rowDT As DataRow = dst.Tables("ASTGRIDB").Rows.Find(New Object() {0, gcol.Index}) ' 0 is hard-coded band 0
                                    Dim colDT As Integer = Val(rowDT.Item("COL") & "") + Col0 ' THIS LINE CAUSED EMPTY COLUMNS HEADERS TO APPEAR IN BETWEEN COLUMN HEADERS; DATA WAS FINE, THO
                                    Gembox_Export_Paint_Column(grd, ws, gcol, r + lvl + 1, colDT) ' col + gcol.Header.VisiblePosition)
                                    ' need to set up a map - visibleposition is not going to work if any columns have a span > 1
                                End If
                            Next
                            For lvl As Int16 = 0 To .LevelCount - 1
                                ws.Rows(r + lvl + 1).Height = .ColHeaderLines * 255 * 1.5
                                ws.Rows(r + lvl + 1).Style.VerticalAlignment = VerticalAlignmentStyle.Center
                            Next
                        End If

                        col += max_cols_in_level
                        col_max = col - 1
                    End If
                Next

                ws.Rows(r).Height = .GroupHeaderLines * 255 * 1.5
                ws.Rows(r).Style.VerticalAlignment = VerticalAlignmentStyle.Center

                If .ColHeadersVisible Then
                    r += .LevelCount
                End If
            End If

            Dim colors() As System.Drawing.Color =
            {Color.PaleGoldenrod, Color.PaleGreen, Color.PaleTurquoise, Color.Beige,
             Color.PaleGoldenrod, Color.PaleGreen, Color.PaleTurquoise, Color.Beige,
             Color.PaleGoldenrod, Color.PaleGreen, Color.PaleTurquoise, Color.Beige}

            If Display_Group_Headers_on_Left Then
                If gbyCols.Count <> 0 Then
                    For i As Integer = 0 To gbyCols.Count - 1
                        ws.Cells(r, i).Value = gbyCols(i).Header.Caption
                        ws.Cells(r, i).Style.FillPattern.PatternForegroundColor = colors(i)
                        ws.Cells(r, i).Style.FillPattern.PatternStyle = FillPatternStyle.Solid
                    Next
                End If
            End If

            Gembox_Export_GridRows(grd, ws, Row0, Col0, 0, r, dst.Tables("ASTGRIDB"), grd.Rows, col_max, gbyCols, colors,
                                   Display_Group_Headers_at_Col0, Display_Group_Headers_on_Left)

            If grd.Rows.SummaryValues.Count > 0 Then

                r += 1
                If grd.DisplayLayout.Bands.Count = 1 Then
                    ws.Cells(r, 0).Value = grd.DisplayLayout.Bands(0).SummaryFooterCaption
                Else
                    ws.Cells(r, 0).Value = "Totals"
                End If

                For col As Integer = 0 + Col0 To col_max + Col0
                    ws.Cells(r, col).Style.Font.Name = grd.Font.Name
                    ws.Cells(r, col).Style.FillPattern.PatternForegroundColor = Color.LightGray
                    ws.Cells(r, col).Style.FillPattern.PatternStyle = FillPatternStyle.Solid
                Next

                For Each sum As UltraWinGrid.SummaryValue In grd.Rows.SummaryValues
                    Dim gcol As UltraWinGrid.UltraGridColumn = sum.SummarySettings.SourceColumn
                    If Not gcol.Hidden And (gcol.Band.Groups.Count = 0 OrElse (gcol.Group IsNot Nothing AndAlso Not gcol.Group.Hidden)) Then
                        Dim row As Integer = r + gcol.Level + 1
                        Dim rowDT As DataRow = dst.Tables("ASTGRIDB").Rows.Find(New Object() {0, gcol.Index}) ' 0 is hard coded band 0
                        Dim col As Integer = Val(rowDT.Item("COL") & "") + Col0

                        With ws.Cells(row, col)
                            .Value = sum.Value
                            .Style.Font.Name = grd.Font.Name
                            If gcol.Format & "" <> "" Then
                                .Style.NumberFormat = gcol.Format
                            End If
                        End With
                    End If
                Next
            End If

            For col As Integer = 0 + Col0 To col_max + Col0
                ws.Columns(col).Style.Font.Name = grd.Font.Name
            Next

            'ws.ViewOptions.OutlineRowButtonsBelow = False

        End With

        ws.Cells.GetSubrangeRelative(1, 0, col_max + 1, 1).Merged = True

    End Sub

    Sub Gembox_Export_GridRows(
    ByVal grd As UltraWinGrid.UltraGrid,
    ByVal ws As GemBox.Spreadsheet.ExcelWorksheet,
    ByVal Row0 As Int64,
    ByVal Col0 As Int64,
    ByVal OutlineLevel As Integer,
    ByRef r As Integer,
    ByVal DT As DataTable,
    ByVal grows As UltraWinGrid.RowsCollection,
    ByVal col_max As Integer,
    ByVal gbyCols As List(Of UltraWinGrid.UltraGridColumn),
    ByVal colors() As System.Drawing.Color,
    ByVal Display_Group_Headers_at_Col0 As Boolean,
    ByVal Display_Group_Headers_on_Left As Boolean)

        With grd.DisplayLayout.Bands(0)

            For Each grow As UltraWinGrid.UltraGridRow In grows
                If Not grow.IsFilteredOut Then
                    If grow.IsGroupByRow Then
                        Dim gbrow As UltraWinGrid.UltraGridGroupByRow = grow

                        Dim gbr As UltraWinGrid.UltraGridGroupByRow = gbrow

                        If Display_Group_Headers_at_Col0 Then
                            r += 1
                            ws.Cells(r, Col0).Value = gbr.Description  ' gbrow.Value
                            If OutlineLevel > 0 Then
                                ws.Rows(r).OutlineLevel = OutlineLevel
                            End If
                        End If

                        If Display_Group_Headers_on_Left Then
                            For i As Integer = OutlineLevel To 0 Step -1
                                ws.Cells(r + 1, i).Value = gbr.Value
                                ws.Cells(r + 1, i).Style.FillPattern.PatternForegroundColor = colors(i)
                                ws.Cells(r + 1, i).Style.FillPattern.PatternStyle = FillPatternStyle.Solid
                                If i <> 0 Then
                                    gbr = gbr.ParentRow
                                End If
                            Next
                            If OutlineLevel < gbyCols.Count - 1 Then
                                For i As Integer = OutlineLevel + 1 To gbyCols.Count - 1
                                    ws.Cells(r + 1, i).Style.FillPattern.PatternForegroundColor = colors(OutlineLevel)
                                    ws.Cells(r + 1, i).Style.FillPattern.PatternStyle = FillPatternStyle.Solid
                                Next
                            End If
                        End If


                        If gbrow.Rows.Count <> 0 AndAlso Not gbrow.GetChild(UltraWinGrid.ChildRow.First).IsGroupByRow Then
                            ' Repaint Header
                            r += 1
                            ws.Rows(5).Cells.CopyTo(r, 0)
                            ws.Rows(r).OutlineLevel = OutlineLevel + 1
                        End If

                        Gembox_Export_GridRows(grd, ws, Row0, Col0, OutlineLevel + 1, r, DT, gbrow.Rows, col_max, gbyCols, colors,
                                                   Display_Group_Headers_at_Col0, Display_Group_Headers_on_Left)

                        ' Banded Group-By Row Totals

                        r += 1

                        If Display_Group_Headers_on_Left Then
                            gbr = gbrow
                            For i As Integer = OutlineLevel To 0 Step -1
                                ws.Cells(r, i).Value = gbr.Value
                                ws.Cells(r, i).Style.FillPattern.PatternForegroundColor = colors(i)
                                ws.Cells(r, i).Style.FillPattern.PatternStyle = FillPatternStyle.Solid
                                If i <> 0 Then
                                    gbr = gbr.ParentRow
                                End If
                            Next
                            If OutlineLevel < gbyCols.Count - 1 Then
                                For i As Integer = OutlineLevel + 1 To gbyCols.Count - 1
                                    ws.Cells(r, i).Style.FillPattern.PatternForegroundColor = colors(OutlineLevel)
                                    ws.Cells(r, i).Style.FillPattern.PatternStyle = FillPatternStyle.Solid
                                Next
                            End If
                        End If

                        If OutlineLevel > 0 Then
                            ws.Rows(r).OutlineLevel = OutlineLevel
                        End If

                        For col As Integer = 0 + Col0 To col_max + Col0
                            ws.Cells(r, col).Style.Font.Name = grd.Font.Name
                            ws.Cells(r, col).Style.FillPattern.PatternForegroundColor = colors(OutlineLevel)
                            ws.Cells(r, col).Style.FillPattern.PatternStyle = FillPatternStyle.Solid
                        Next

                        For Each sum As UltraWinGrid.SummaryValue In gbrow.Rows.SummaryValues
                            Dim gcol As UltraWinGrid.UltraGridColumn = sum.SummarySettings.SourceColumn
                            Dim row As Integer = r + gcol.Level
                            Dim rowDT As DataRow = DT.Rows.Find(New Object() {0, gcol.Index})
                            If rowDT IsNot Nothing Then
                                Dim col As Integer = Val(rowDT.Item("COL") & "") + Col0

                                With ws.Cells(row, col)
                                    .Value = sum.Value
                                    .Style.Font.Name = grd.Font.Name
                                    If gcol.Format & "" <> "" Then
                                        .Style.NumberFormat = gcol.Format
                                    End If
                                End With
                            End If
                        Next
                    Else
                        If Display_Group_Headers_on_Left Then
                            For i As Integer = 0 To OutlineLevel - 1
                                For j As Integer = 1 To grow.Band.LevelCount
                                    ws.Cells(r + j, i).Value = grow.Cells(gbyCols(i)).Value
                                    ws.Cells(r + j, i).Style.FillPattern.PatternForegroundColor = colors(i)
                                    ws.Cells(r + j, i).Style.FillPattern.PatternStyle = FillPatternStyle.Solid
                                Next
                            Next
                        End If

                        For Each gcell As UltraWinGrid.UltraGridCell In grow.Cells
                            If Not gcell.Column.Hidden And (gcell.Column.Band.Groups.Count = 0 OrElse (gcell.Column.Group IsNot Nothing AndAlso Not gcell.Column.Group.Hidden)) Then
                                Dim row As Integer = r + gcell.Column.Level + 1
                                Dim rowDT As DataRow = DT.Rows.Find(New Object() {0, gcell.Column.Index})
                                Dim col As Integer = Val(rowDT.Item("COL") & "") + Col0
                                With ws.Cells(row, col)
                                    If gcell.Column.ColSpan > 1 Then
                                        ws.Cells.GetSubrangeRelative(row, col, gcell.Column.ColSpan, 1).Merged = True
                                        'ws.Cells.GetSubrangeRelative(row, col, gcell.Column.ColSpan, 1).Width = gcell.Column.Width
                                    End If

                                    Dim skip_formatting As Boolean = False ' True

                                    If Not skip_formatting Then
                                        If Not gcell.Appearance.ForeColor = Color.Empty Then
                                            .Style.Font.Color = gcell.Appearance.ForeColor
                                        Else
                                            If Not gcell.Column.CellAppearance.ForeColor = Color.Empty Then
                                                .Style.Font.Color = gcell.Column.CellAppearance.ForeColor
                                            End If
                                        End If

                                        If Not gcell.Appearance.BackColor = Color.Empty Then
                                            .Style.FillPattern.PatternForegroundColor = gcell.Appearance.BackColor
                                            .Style.FillPattern.PatternStyle = FillPatternStyle.Solid
                                        Else
                                            If Not gcell.Column.CellAppearance.BackColor = Color.Empty Then
                                                .Style.FillPattern.PatternForegroundColor = gcell.Column.CellAppearance.BackColor
                                                .Style.FillPattern.PatternStyle = FillPatternStyle.Solid
                                            Else
                                                If gcell.Column.Group IsNot Nothing AndAlso Not gcell.Column.Group.CellAppearance.BackColor = Color.Empty Then
                                                    .Style.FillPattern.PatternForegroundColor = gcell.Column.Group.CellAppearance.BackColor
                                                    .Style.FillPattern.PatternStyle = FillPatternStyle.Solid
                                                End If
                                            End If
                                        End If
                                    End If
                                    If TypeOf (gcell.Value) Is System.Array Then
                                    Else
                                        .Value = gcell.Value
                                    End If

                                    .Style.Font.Name = grd.Font.Name
                                    If gcell.Column.Format & "" <> "" Then
                                        If gcell.Column.DataType.ToString = "System.DateTime" Then
                                            .Style.NumberFormat = gcell.Column.Format.ToLower & ";@"
                                        Else
                                            .Style.NumberFormat = gcell.Column.Format
                                        End If
                                    Else
                                        If gcell.Column.DataType.ToString = "System.DateTime" Then
                                            .Style.NumberFormat = "mm/dd/yy;@"
                                        Else
                                            If gcell.Column.Style = UltraWinGrid.ColumnStyle.CheckBox Then
                                                .Value = gcell.Value ' for checkboxes
                                            Else
                                                .Value = gcell.Text ' for value lists
                                            End If
                                        End If
                                    End If


                                    If gcell.Column.CellAppearance.TextHAlign = HAlign.Right Then
                                        .Style.HorizontalAlignment = HorizontalAlignmentStyle.Right
                                    ElseIf gcell.Column.CellAppearance.TextHAlign = HAlign.Center Then
                                        .Style.HorizontalAlignment = HorizontalAlignmentStyle.Center
                                    End If
                                End With
                            End If
                        Next

                        r += .LevelCount
                        'If OutlineLevel > 0 Then
                        ws.Rows(r).Height = 300
                        ws.Rows(r).Style.WrapText = False
                        ws.Rows(r).OutlineLevel = OutlineLevel
                        'End If
                    End If
                End If
            Next
        End With
    End Sub
#End Region

    Private Sub grd_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs)
        ASCMAIN1.Progress("")
        disable_arrows = False
    End Sub

    Private Sub grd_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs)

        If e.Cell.Column.Key = "ATTACHMENTS" Then
            If grds_with_Attachments.ContainsKey(e.Cell.Row.Band.Key) Then
                Dim AB As Attachment_Button = grds_with_Attachments(e.Cell.Row.Band.Key)

                Dim F As New ASFATTA1

                Dim EE As New Dropped_On_Entity
                EE.TABLE_NAME = AB.TABLE_NAME
                EE.COLUMN_NAME = AB.COLUMN_NAME
                EE.CODE_VALUE = e.Cell.Row.Cells(AB.COLUMN_NAME).Value
                EE.DESC_VALUE = "Attachments for " & e.Cell.Row.Cells(AB.COLUMN_NAME).Text
                EE.ATTACHMENT_NOTES = ""

                If AB.allow_update = "I" Then
                    EE.READ_ONLY = InquiryMode
                ElseIf AB.allow_update = "Y" Or
                    (AB.allow_update = "G" And AB.grd.DisplayLayout.Override.AllowUpdate <> DefaultableBoolean.False) Then
                    EE.READ_ONLY = False
                Else
                    EE.READ_ONLY = True
                End If

                F.ENTITY = EE
                F.ShowDialog()
                F.Dispose()

                AB.grd.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
            End If
        End If
    End Sub

    Private Sub grd_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs)
        disable_arrows = True
        doubleclicked = True

        Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsFilterRow AndAlso grd.ActiveCell IsNot Nothing And 1 <> 1 Then
            ' DISABLING THE DOUBLE CLICK ON THE FILTER ROW- CAUSING TOO MANY PROBLEMS
            ' MAYBE WE ENABLE A DOUBLECLICK ON THE FILTER ROW BY USING OPTIONS, OR SOME OTHER DEVICE SIGNALLING INTENTIONAL DESIRE TO SEE A LOOKUP
            Try
                Dim COLUMN_NAME As String = grd.ActiveCell.Column.Key
                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL(COLUMN_NAME)
                If ASCMAIN1.CodeSelector.SQL <> "" Then
                    ASCMAIN1.CodeSelector.MultipleSelections = False
                    ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                    Using F As New ASFCODE1
                        F.ShowDialog()
                    End Using
                    If ASCMAIN1.CodeSelector.SelectedCode <> "" Then
                        grd.ActiveCell.Value = ASCMAIN1.CodeSelector.SelectedCode
                    End If
                End If

            Catch ex As Exception

            Finally
                Me.Cursor = Cursors.Default
            End Try


        End If
    End Sub

    Private Sub grd_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)

        'If e.KeyValue = 13 Then Stop
        grdKeyValue = e.KeyValue
        grdKeyData = e.KeyData

        'Select Case e.KeyCode
        '    Case Keys.Control Or Keys.V
        '        With DirectCast(sender, UltraWinGrid.UltraGrid)
        '            If .DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False Then
        '                e.Handled = True
        '            Else
        '                .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)
        '                .ActiveCell.Column.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.CellSelect
        '            End If
        '        End With
        '    Case Else
        '        'Stop
        'End Select

        If grdKeyData = Keys.F7 Then
            ' Stop
            With DirectCast(sender, UltraWinGrid.UltraGrid)
                If .ActiveRow IsNot Nothing AndAlso .ActiveRow.IsAddRow Then
                    .UpdateData()
                    .ActiveRow = .DisplayLayout.Bands(0).AddNew
                    If .ActiveRow IsNot Nothing Then
                        For I As Integer = 0 To .ActiveRow.Cells.Count - 1
                            If .ActiveRow.Cells(I).Column.CellActivation = UltraWinGrid.Activation.AllowEdit Then
                                .ActiveCell = .ActiveRow.Cells(I)
                                Exit For
                            End If
                        Next
                    End If
                End If
            End With

            Exit Sub
        End If

        If grdKeyValue = Keys.Enter Then
            With DirectCast(sender, UltraWinGrid.UltraGrid)


                If .ActiveCell Is Nothing Then
                    Exit Sub
                End If

                If .ActiveCell.IsInEditMode Then
                    Dim eee As System.Windows.Forms.KeyPressEventArgs = New System.Windows.Forms.KeyPressEventArgs(ChrW(grdKeyValue))
                    'If .ActiveCell.IsInEditMode Then

                    'End If
                    Me.grd_KeyPress(sender, eee)
                    'If .ActiveCell.IsInEditMode Then
                    '    grdKeyData = Keys.Tab
                    '    Exit Sub
                    'End If
                    e.Handled = eee.Handled ' True
                    If .ActiveCell Is Nothing Then
                        Exit Sub
                    Else
                        If .ActiveCell.Column.CellMultiLine Then Exit Sub
                    End If
                End If
                e.Handled = True

                If .ActiveRow Is Nothing Then
                    Exit Sub
                End If

                Dim Highest_Visible_Position_Encountered As Integer = -1
                Dim Highest_Visible_Position_Encountered_Permitting_Edit As Integer = -1
                Dim Key_Of_Last_Visible_Column As String = ""
                Dim Key_Of_Last_Visible_Column_Permitting_Edit As String = ""
                For Each Grid_Cell As UltraWinGrid.UltraGridCell In .ActiveRow.Cells
                    Dim GH As Boolean = False
                    If Grid_Cell.Column.Group IsNot Nothing AndAlso Grid_Cell.Column.Group.Hidden Then GH = True

                    If Grid_Cell.Column.Header.VisiblePositionWithinBand > Highest_Visible_Position_Encountered And Not Grid_Cell.Column.Hidden And Not GH Then
                        Key_Of_Last_Visible_Column = Grid_Cell.Column.Key
                        Highest_Visible_Position_Encountered = Grid_Cell.Column.Header.VisiblePositionWithinBand
                        If Grid_Cell.Column.CellActivation = UltraWinGrid.Activation.AllowEdit Then
                            Key_Of_Last_Visible_Column_Permitting_Edit = Grid_Cell.Column.Key
                            Highest_Visible_Position_Encountered_Permitting_Edit = Grid_Cell.Column.Header.VisiblePositionWithinBand
                        End If
                    End If
                Next

                e.Handled = True
                If .ActiveCell.IsInEditMode Then
                    .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
                End If
                If .ActiveCell IsNot Nothing AndAlso .ActiveCell.IsInEditMode Then
                    Exit Sub
                End If
                If .ActiveRow Is Nothing Then
                    Exit Sub
                End If

                If .ActiveRow.IsAddRow And .ActiveCell IsNot Nothing Then
                    'If .ActiveCell IsNot Nothing Then

                    If .ActiveCell.Column.Key = Key_Of_Last_Visible_Column_Permitting_Edit _
                    Or .ActiveCell.Column.Header.VisiblePositionWithinBand > Highest_Visible_Position_Encountered_Permitting_Edit Then ' Key_Of_Last_Visible_Column Then
                        ' THE NEXT 3 LINES WERE REMMED/UNREMMED TO ACCOMODATE TS ENTRY WHERE THE USER PICKS A TS FROM ANOTHER JOB, AND THEN CR'S THROUGH THE ADDROW - THE ADDROW WAS NOT GETTING ADDED TO THE GRID UNTIL WE CHANGED THE 3 LINES
                        ' HAD TO PUT THESE LINES BACK IN ORDER TO GET ROWS TO BE ADDED IN LNO SEQ, IN ORDER TO STAY ON THE ADDROW, IN PMTJOBM1.GRDPMTJOBL4
                        e.Handled = False
                        'e.Handled = True ' UNREMMED 10/19 TO GET TS ENTRY ADDROW TO WORK
                        'Dim B As Int16 = .ActiveRow.Band.Index
                        '.ActiveRow.Update() ' UNREMMED 10/19 TO GET TS ENTRY ADDROW TO WORK
                        'If .ActiveRow IsNot Nothing AndAlso .ActiveRow.IsAddRow Then
                        '    Try
                        '        .DisplayLayout.Bands(B).AddNew.Activate()

                        '    Catch ex As Exception

                        '    End Try

                        'End If
                        '.PerformAction(UltraWinGrid.UltraGridAction.ActivateCell)

                        '.UpdateData()
                        'e.Handled = True
                        'Exit Sub
                        ''.PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab, False, False)

                    Else
                        'If Not e.Handled Then
                        If .ActiveCell IsNot Nothing Then
                            If grdKeyData <> Keys.Tab Then ' set to tab to indicate that the before exit edit mode code had a problem with the value
                                Dim LAST_CELL As String = .ActiveCell.Column.Key & ""
                                Do
                                    Dim gr As UltraWinGrid.UltraGridRow = .ActiveCell.Row
                                    .PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab, False, False)

                                    If .ActiveCell IsNot Nothing Then

                                        If Not gr.Equals(.ActiveCell.Row) Then
                                            Exit Do
                                        End If

                                        If .ActiveCell.Column.Key = LAST_CELL Then
                                            Exit Do
                                        Else
                                            LAST_CELL = .ActiveCell.Column.Key
                                        End If
                                    End If
                                Loop While .ActiveCell IsNot Nothing AndAlso .ActiveCell.Column.CellActivation = UltraWinGrid.Activation.NoEdit
                            End If
                        End If
                        'e.Handled = True
                        'End If
                        'If .ActiveCell IsNot Nothing Then
                        '    If .ActiveCell.Column.CellActivation <> UltraWinGrid.Activation.AllowEdit Then
                        '        ' we must be at the end of the row and there are no more cells left which permit editing
                        '    End If
                        'End If
                        'Stop
                    End If
                Else
                    If .ActiveCell IsNot Nothing Then
                        If .ActiveCell.Column.Key = Key_Of_Last_Visible_Column Then
                            .UpdateData()
                        End If

                        ' the following if was necessary so that pressing return navigates to the next row, same cell in SOFORDR1.grdSOTORDR3 LOT_ORDER_QTY field
                        If .ActiveCell.Column.Key = Key_Of_Last_Visible_Column_Permitting_Edit Then
                            .PerformAction(UltraWinGrid.UltraGridAction.BelowCell, False, False)
                        Else

                            Try
                                Do
                                    Dim CURRENT_CELL As String = .ActiveCell.Column.Key

                                    Dim gr As UltraWinGrid.UltraGridRow = .ActiveCell.Row

                                    .PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab, False, False)


                                    If Not gr.Equals(.ActiveCell.Row) Then
                                        Exit Do
                                    End If

                                    If .ActiveRow.IsFilterRow Then
                                        Exit Do
                                    End If

                                    If .ActiveCell.Column.Key = CURRENT_CELL Then
                                        ' WE ARE STUCK, SO GET THE HECK OUT OF THIS LOOP
                                        Exit Do
                                    End If
                                Loop While .ActiveCell.Column.CellActivation = UltraWinGrid.Activation.NoEdit
                            Catch ex As Exception

                            End Try
                        End If


                        '.PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab, False, False)
                    End If
                End If
            End With
        End If

        If grdKeyValue = Keys.Up Or grdKeyValue = Keys.Down Or grdKeyValue = Keys.PageDown Or grdKeyValue = Keys.PageUp Then
            With DirectCast(sender, UltraWinGrid.UltraGrid)
                If .ActiveCell IsNot Nothing AndAlso .ActiveCell.IsInEditMode Then
                    Dim eee As System.Windows.Forms.KeyPressEventArgs = New System.Windows.Forms.KeyPressEventArgs(ChrW(grdKeyValue))
                    Me.grd_KeyPress(sender, eee)
                    e.Handled = True
                End If
            End With
        End If
    End Sub

    Private Sub grd_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

        If e.KeyChar = "" Then

        End If

        With DirectCast(sender, UltraWinGrid.UltraGrid)
            Dim COLUMN_NAME As String = ""
            If .ActiveCell IsNot Nothing Then
                COLUMN_NAME = .ActiveCell.Column.Key
            End If

            If grdKeyData = Keys.Enter Then
                If .ActiveCell IsNot Nothing AndAlso .ActiveCell.IsInEditMode Then

                    If .ActiveCell.Column.CellMultiLine Then
                        'If .ActiveCell.SelStart = .ActiveCell.Text.Length Then
                        '    Exit Sub
                        'End If
                        '.DisplayLayout.Override.RowSizing = UltraWinGrid.RowSizing.AutoFree

                        Exit Sub
                        'e.Handled = True
                    End If

                    .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)
                    If .ActiveCell IsNot Nothing Then
                        If .ActiveCell.IsInEditMode Then
                            grdKeyData = Keys.Tab ' set to tab to indicate that the before exit edit mode code had a problem with the value
                        End If
                    End If
                    Exit Sub
                End If
            End If

            Dim character As Char = e.KeyChar
            grdKeyData = Asc(character)
            '

            If .ActiveCell IsNot Nothing Then
                If Not .ActiveCell.IsInEditMode Then
                    If ((grdKeyData >= 96 And grdKeyData <= 122) _
                    Or (grdKeyData >= 65 And grdKeyData <= 90) _
                    Or (grdKeyData >= 48 And grdKeyData <= 57) _
                    Or (grdKeyData >= 32 And grdKeyData <= 32) _
                    Or (grdKeyData >= 32 + 128 And grdKeyData <= 126 + 128) _
                    Or (grdKeyData = 43 Or grdKeyData = 45 Or grdKeyData = 46)
                    ) And (.DisplayLayout.Override.AllowAddNew <> UltraWinGrid.AllowAddNew.No Or .DisplayLayout.Override.AllowUpdate <> DefaultableBoolean.False) Then
                        .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
                        If .ActiveCell.IsInEditMode Then
                            ' get its editor
                            Dim editor As EmbeddableEditorBase = .ActiveCell.EditorResolved

                            ' if the editor supports selectable text
                            If editor.SupportsSelectableText Then
                                ' select all the text so it can be replaced
                                editor.SelectionStart = 0
                                editor.SelectionLength = editor.TextLength

                                If TypeOf editor Is EditorWithMask Then
                                    ' just clear the selected text and let the grid
                                    ' forward the keypress to the editor
                                    editor.SelectedText = String.Empty
                                Else
                                    ' then replace the selected text with the character
                                    editor.SelectedText = ChrW(grdKeyData) ' New String(CChar(CStr(grdKeyData)), 1)
                                    ' mark the event as handled so the grid doesn't process it
                                    e.Handled = True
                                End If
                            End If
                        End If
                        Exit Sub
                    End If
                Else

                    'Select Case e.KeyValue
                    '    Case Keys.Escape
                    '        .PerformAction(UltraWinGrid.UltraGridAction.UndoCell)
                    '        e.Handled = True
                    '        Exit Sub
                    'End Select
                End If
            End If

            If Not disable_arrows Then
                Select Case grdKeyValue
                    Case Keys.PageUp
                        .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
                        .PerformAction(UltraWinGrid.UltraGridAction.PageUpCell, False, False)
                        e.Handled = True
                        '.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
                    Case Keys.PageDown
                        .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
                        .PerformAction(UltraWinGrid.UltraGridAction.PageDownCell, False, False)
                        e.Handled = True
                        '.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
                    Case Keys.Up
                        .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
                        .PerformAction(UltraWinGrid.UltraGridAction.AboveCell, False, False)
                        e.Handled = True
                        '.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
                    Case Keys.Down
                        .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
                        If .ActiveRow IsNot Nothing AndAlso .ActiveRow.IsAddRow Then
                            .UpdateData()
                        End If
                        'Try
                        .PerformAction(UltraWinGrid.UltraGridAction.BelowCell, False, False)
                        'Catch ex As Exception

                        'End Try
                        e.Handled = True
                        '.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
                    Case Keys.Right
                        .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
                        .PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab, False, False)
                        e.Handled = True
                        '.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
                    Case Keys.Left
                        .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
                        .PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab, False, False)
                        e.Handled = True
                        '.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
                End Select
            Else
                Select Case grdKeyValue
                    Case Keys.PageUp
                        e.Handled = True
                    Case Keys.PageDown
                        e.Handled = True
                    Case Keys.Up
                        e.Handled = True
                    Case Keys.Down
                        e.Handled = True
                    Case Keys.Right
                        'e.Handled = True
                    Case Keys.Left
                        'e.Handled = True
                End Select
            End If

            Select Case grdKeyValue
                Case Keys.Delete
                    If .Selected.Rows.Count = 0 Then
                        If .ActiveCell IsNot Nothing Then
                            If .ActiveCell.Column.CellActivation = UltraWinGrid.Activation.AllowEdit Then
                                .ActiveCell.Value = DBNull.Value
                            End If
                        End If
                    End If

            End Select

            If .ActiveCell IsNot Nothing Then
                Select Case grdKeyData

                    Case Keys.Control Or Keys.V
                        If .DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False Then
                            e.Handled = True
                        Else
                            .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)
                            .ActiveCell.Column.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.CellSelect
                            'Call Paste_Data()
                            '.ActiveCell.Column.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit
                        End If

                    Case Keys.Control Or Keys.C
                        .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)
                        .ActiveCell.Column.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.CellSelect
                        'Call Paste_Data()
                        '.ActiveCell.Column.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit
                End Select
            End If

        End With
    End Sub

    'Private Sub grd_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    With DirectCast(sender, UltraWinGrid.UltraGrid)

    '        Dim COLUMN_NAME As String = ""
    '        If .ActiveCell IsNot Nothing Then
    '            COLUMN_NAME = .ActiveCell.Column.Key
    '        End If

    '        If .ActiveCell IsNot Nothing Then
    '            If Not .ActiveCell.IsInEditMode Then
    '                If ((e.KeyData >= 96 And e.KeyData <= 111) Or (e.KeyData >= 65 And e.KeyData <= 90) Or (e.KeyData >= 48 And e.KeyData <= 57) Or (e.KeyData >= 32 And e.KeyData <= 32) Or (e.KeyData >= 32 + 128 And e.KeyData <= 126 + 128)) Then
    '                    .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)

    '                    If .ActiveCell.IsInEditMode Then
    '                        ' get its editor
    '                        Dim editor As EmbeddableEditorBase = .ActiveCell.EditorResolved

    '                        ' if the editor supports selectable text
    '                        If editor.SupportsSelectableText Then
    '                            ' select all the text so it can be replaced
    '                            editor.SelectionStart = 0
    '                            editor.SelectionLength = editor.TextLength

    '                            If TypeOf editor Is EditorWithMask Then
    '                                ' just clear the selected text and let the grid
    '                                ' forward the keypress to the editor
    '                                editor.SelectedText = String.Empty
    '                            Else
    '                                ' then replace the selected text with the character
    '                                editor.SelectedText = ChrW(e.KeyData) ' New String(CChar(CStr(e.KeyData)), 1)

    '                                ' mark the event as handled so the grid doesn't process it
    '                                e.Handled = True
    '                            End If
    '                        End If
    '                    End If

    '                    Exit Sub
    '                End If
    '            Else

    '                'Select Case e.KeyValue
    '                '    Case Keys.Escape
    '                '        .PerformAction(UltraWinGrid.UltraGridAction.UndoCell)
    '                '        e.Handled = True
    '                '        Exit Sub
    '                'End Select

    '            End If
    '        End If

    '        If Not disable_arrows Then
    '            Select Case e.KeyValue
    '                Case Keys.PageUp
    '                    .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
    '                    .PerformAction(UltraWinGrid.UltraGridAction.PageUpCell, False, False)
    '                    e.Handled = True
    '                    '.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
    '                Case Keys.PageDown
    '                    .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
    '                    .PerformAction(UltraWinGrid.UltraGridAction.PageDownCell, False, False)
    '                    e.Handled = True
    '                    '.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode, False, False)

    '                Case Keys.Up
    '                    .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
    '                    .PerformAction(UltraWinGrid.UltraGridAction.AboveCell, False, False)
    '                    e.Handled = True
    '                    '.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
    '                Case Keys.Down
    '                    .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
    '                    .PerformAction(UltraWinGrid.UltraGridAction.BelowCell, False, False)
    '                    e.Handled = True
    '                    If .ActiveRow IsNot Nothing AndAlso .ActiveRow.IsAddRow Then
    '                        .UpdateData()
    '                    End If
    '                    '.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
    '                Case Keys.Right
    '                    .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
    '                    .PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab, False, False)
    '                    e.Handled = True
    '                    '.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
    '                Case Keys.Left
    '                    .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
    '                    .PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab, False, False)
    '                    e.Handled = True
    '                    '.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
    '            End Select
    '        Else
    '            Select Case e.KeyValue
    '                Case Keys.PageUp
    '                    e.Handled = True
    '                Case Keys.PageDown
    '                    e.Handled = True
    '                Case Keys.Up
    '                    e.Handled = True
    '                Case Keys.Down
    '                    e.Handled = True
    '                Case Keys.Right
    '                    'e.Handled = True
    '                Case Keys.Left
    '                    'e.Handled = True
    '            End Select

    '        End If

    '        Select Case e.KeyValue
    '            Case Keys.Enter
    '                If .ActiveRow Is Nothing Then
    '                    Exit Sub
    '                End If

    '                Dim Highest_Visible_Position_Encountered As Integer = -1
    '                Dim Key_Of_Last_Visible_Column As String = ""
    '                Dim Key_Of_Last_Visible_Column_Permitting_Edit As String = ""
    '                For Each Grid_Cell As UltraWinGrid.UltraGridCell In .ActiveRow.Cells
    '                    If Grid_Cell.Column.Header.VisiblePositionWithinBand > Highest_Visible_Position_Encountered And Not Grid_Cell.Column.Hidden Then
    '                        Key_Of_Last_Visible_Column = Grid_Cell.Column.Key
    '                        If Grid_Cell.Column.CellActivation = UltraWinGrid.Activation.AllowEdit Then
    '                            Key_Of_Last_Visible_Column_Permitting_Edit = Grid_Cell.Column.Key
    '                            Highest_Visible_Position_Encountered = Grid_Cell.Column.Header.VisiblePositionWithinBand
    '                        End If
    '                    End If
    '                Next

    '                e.Handled = True
    '                .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
    '                If .ActiveCell IsNot Nothing AndAlso .ActiveCell.IsInEditMode Then
    '                    Exit Sub
    '                End If
    '                If .ActiveRow Is Nothing Then
    '                    Exit Sub
    '                End If

    '                If .ActiveRow.IsAddRow And .ActiveCell IsNot Nothing Then
    '                    If .ActiveCell.Column.Key = Key_Of_Last_Visible_Column_Permitting_Edit Then ' Key_Of_Last_Visible_Column Then
    '                        e.Handled = False
    '                    Else
    '                        .PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab, False, False)
    '                        'If .ActiveCell IsNot Nothing Then
    '                        '    If .ActiveCell.Column.CellActivation <> UltraWinGrid.Activation.AllowEdit Then
    '                        '        ' we must be at the end of the row and there are no more cells left which permit editing

    '                        '    End If
    '                        'End If
    '                        'Stop
    '                    End If
    '                Else
    '                    If .ActiveCell IsNot Nothing Then
    '                        If .ActiveCell.Column.Key = Key_Of_Last_Visible_Column Then
    '                            .UpdateData()
    '                        End If
    '                        .PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab, False, False)
    '                    End If
    '                End If

    '            Case Keys.Delete
    '                If .Selected.Rows.Count = 0 Then
    '                    If .ActiveCell IsNot Nothing Then
    '                        If .ActiveCell.Column.CellActivation = UltraWinGrid.Activation.AllowEdit Then
    '                            .ActiveCell.Value = DBNull.Value
    '                        End If
    '                    End If
    '                End If


    '        End Select

    '        If .ActiveCell IsNot Nothing Then
    '            Select Case e.KeyData

    '                Case Keys.Control Or Keys.V
    '                    .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)
    '                    .ActiveCell.Column.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.CellSelect
    '                    'Call Paste_Data()
    '                    '.ActiveCell.Column.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit

    '                Case Keys.Control Or Keys.C
    '                    .PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)
    '                    .ActiveCell.Column.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.CellSelect
    '                    'Call Paste_Data()
    '                    '.ActiveCell.Column.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit
    '            End Select
    '        End If

    '    End With
    'End Sub

    Private Sub grd_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs)
        If e.CancellingEditOperation Then

        End If
    End Sub


    Private Sub grd_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        'Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        'If Not grd.ActiveRow Is Nothing AndAlso grd.ActiveRow.DataChanged Then
        '    'grd.ActiveRow.CancelUpdate()
        'ElseIf Not grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.DataChanged Then
        '    grd.ActiveRow.Update()
        'End If
    End Sub

    Private Sub grd_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)

        If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsFilterRow Then
            If grd.ActiveCell IsNot Nothing Then
                'THIS LINE OF CODE WILL CAUSE U TO LOSE THE VALUE U PLACE IN THE FILTER ROW IF U CLICK SOMETHING OUTSIDE THE GRID, BUT VALUE WILL REMAIN IN FILTER IF U CLICK ON A GRID ROW, SO NOT SURE WHY WE NEED IT
                'grd.ActiveCell.CancelUpdate()
            End If
            Exit Sub
        End If

        If Not grd.ActiveRow Is Nothing AndAlso grd.ActiveRow.DataChanged Then
            Try
                ' this next line solved the problem where the grid BeforeRowUpdate event was firing before the cell value was updated
                grd.PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)
                grd.UpdateData()
                If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.DataChanged Then
                    grd.ActiveRow.CancelUpdate()
                    'grd.PerformAction(UltraWinGrid.UltraGridAction.UndoRow)
                End If
            Catch ex As Exception

            End Try

            Try
                grd.PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)
                If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.DataChanged Then
                    grd.ActiveRow.CancelUpdate()
                    'grd.PerformAction(UltraWinGrid.UltraGridAction.UndoRow)
                End If
            Catch ex As Exception
            End Try
            grd.PerformAction(UltraWinGrid.UltraGridAction.UndoRow)
        End If
    End Sub

    Private Sub grd_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs)
        Call ASCMAIN1.grdInitializeLayout(DirectCast(sender, UltraWinGrid.UltraGrid), Me)
    End Sub

    Private Sub grd_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs)
        grd_RowColor(sender, e)
    End Sub

    Sub grd_RowColor(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs)
        If e.Row.IsAddRow Then Exit Sub

        Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        If TypeOf (grd.DataSource) Is DataTable Then
            Dim tbl As DataTable = DirectCast(grd.DataSource, DataTable)

            grd_RowColor(tbl, e.Row)
        End If
    End Sub

    Sub grd_RowColor(ByVal tbl As DataTable, ByVal eRow As UltraWinGrid.UltraGridRow)

        Dim rowstate As String = ""
        If eRow.IsAddRow Then
            rowstate = "Added"
        Else
            If eRow.Band.Index = 0 Then
                Dim i As Int32 = eRow.ListIndex
                If i = -1 Then Exit Sub
                If i < tbl.Rows.Count Then
                    Select Case tbl.Rows(i).RowState
                        Case DataRowState.Added
                            rowstate = "Added"
                        Case DataRowState.Modified
                            rowstate = "Modified"
                    End Select
                End If
            End If
        End If

        Select Case rowstate
            Case "Added"
                'e.Row.Appearance.BackColor = Color.LightGreen
                eRow.RowSelectorAppearance.BackColor = System.Drawing.Color.LightGreen
                eRow.RowSelectorAppearance.BackColor2 = System.Drawing.Color.Green
            Case "Modified"
                'e.Row.Appearance.BackColor = Color.LightSkyBlue
                eRow.RowSelectorAppearance.BackColor = System.Drawing.Color.LightSkyBlue
                eRow.RowSelectorAppearance.BackColor2 = System.Drawing.Color.Blue
        End Select

        If grds_with_Attachments.ContainsKey(eRow.Band.Key) Then

            Dim AB As Attachment_Button = grds_with_Attachments(eRow.Band.Key)
            'If e.Row.Band.Columns.Contains(AB.COLUMN_NAME) Then
            ASCMAIN1.sql = "Select Count (*) from ASTATTA2" _
                & " where TABLE_NAME = :PARM1" _
                & "   and COLUMN_NAME = :PARM2" _
                & "   and CODE_VALUE = :PARM3"

            Dim ATTACHMENTS As Int16 = Val(ASCDATA1.GetDataValue _
                    (ASCMAIN1.sql, "VVV", New Object() _
                     {AB.TABLE_NAME,
                      AB.COLUMN_NAME,
                      eRow.Cells(AB.COLUMN_NAME).Text}))

            If ATTACHMENTS = 0 Then
                eRow.Cells("ATTACHMENTS").Value = DBNull.Value
            Else
                eRow.Cells("ATTACHMENTS").Value = ATTACHMENTS
            End If
            'End If

        End If

    End Sub

    Private Sub grd_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs)

        If BeforeRowsDeletedRows IsNot Nothing Then
            Try
                BeforeRowsDeletedRows.Clear()
                Dim c() As VariantType
                For Each x As UltraWinGrid.UltraGridRow In e.Rows
                    ReDim c(x.Cells.Count - 1)
                    For i As Integer = 0 To x.Cells.Count - 1
                        c(i) = x.Cells(i).Text
                    Next
                    BeforeRowsDeletedRows.Add(c)
                Next

            Catch ex As Exception

            End Try
        End If

    End Sub

    Sub Set_Read_Only(
    ByRef ctl As Control,
    ByVal tf As Boolean,
    Optional ByVal TABLE_NAME_PARENT As String = "")

        If TABLE_NAME_PARENT = "" Then
            TABLE_NAME_PARENT = Absx1.GetABSTableName(ctl)
        End If

        If Not tf And ctl1 Is Nothing Then
            For Each ctlx As Control In ctl.Controls
                If Absx1.GetABSColumnName(ctlx) <> "" Then

                    If ctl1 Is Nothing Then
                        ctl1 = ctlx
                    Else
                        If System.Math.Sqrt(ctlx.Location.X ^ 2 + ctlx.Location.Y ^ 2) _
                         < System.Math.Sqrt(ctl1.Location.X ^ 2 + ctl1.Location.Y ^ 2) Then
                            ctl1 = ctlx
                        End If
                    End If
                End If
            Next
            If ctl1 IsNot Nothing Then
                ctl1.Focus()
            End If
        End If

        For Each child_ctl As Control In ctl.Controls
            If child_ctl.HasChildren Then
                Set_Read_Only(child_ctl, tf, TABLE_NAME_PARENT)
            End If

            Set_Read_Only_for_ctl(child_ctl, tf, TABLE_NAME_PARENT)
        Next

        Set_Read_Only_for_ctl(ctl, tf)

    End Sub

    Sub Set_Read_Only_for_ctl(
    ByRef child_ctl As Control,
    ByVal tf As Boolean,
    Optional ByVal TABLE_NAME_PARENT As String = "")

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(child_ctl)
        Dim TABLE_NAME As String = Absx1.GetABSTableName(child_ctl)
        If TABLE_NAME = "" Then TABLE_NAME = TABLE_NAME_PARENT

        'If COLUMN_NAME = "CUST_NAME" Then Stop
        If tf Then ' If Not tf Then
            ' we may need to MRU columns a little bit differently than the way we are doing it now
            ' now, only those controls in the header group box are having their values saved in the MRU list
            ' need to hook this up for when modes is true for all text boxes
            ' but need to not MRU values that were not keyed in
            ' also need a solution for FM that would work for key fields

            If child_ctl.Text <> "" Then
                If ASCMAIN1.MRUs.ContainsKey(COLUMN_NAME) Then
                    If ASCMAIN1.MRUs(COLUMN_NAME).Contains(child_ctl.Text) Then
                        ASCMAIN1.MRUs(COLUMN_NAME).Remove(child_ctl.Text)
                    End If
                    ASCMAIN1.MRUs(COLUMN_NAME).Add(child_ctl.Text)
                End If
            End If
        End If

        Dim TC As String = COLUMN_NAME
        If TABLE_NAME <> "" Then
            TC = TABLE_NAME & "." & COLUMN_NAME
        End If
        If ABSReadOnly.Contains(child_ctl.Name) Then ' ABSReadOnly.Contains(TC) Then

        Else
            If TypeOf child_ctl Is UltraWinEditors.UltraTextEditor Then
                DirectCast(child_ctl, UltraWinEditors.UltraTextEditor).ReadOnly = tf
                If DirectCast(child_ctl, UltraWinEditors.UltraTextEditor).ButtonsRight.Count = 1 Then
                    DirectCast(child_ctl, UltraWinEditors.UltraTextEditor).ButtonsRight(0).Enabled = Not tf
                End If
            End If
            If TypeOf child_ctl Is UltraWinGrid.UltraCombo Then
                DirectCast(child_ctl, UltraWinGrid.UltraCombo).ReadOnly = tf
            End If
            If TypeOf child_ctl Is UltraWinEditors.UltraDateTimeEditor Then
                DirectCast(child_ctl, UltraWinEditors.UltraDateTimeEditor).ReadOnly = tf
            End If
            If TypeOf child_ctl Is UltraWinSchedule.UltraCalendarCombo Then
                DirectCast(child_ctl, UltraWinSchedule.UltraCalendarCombo).ReadOnly = tf
            End If
            If TypeOf child_ctl Is UltraWinMaskedEdit.UltraMaskedEdit Then
                DirectCast(child_ctl, UltraWinMaskedEdit.UltraMaskedEdit).ReadOnly = tf
            End If
            If TypeOf child_ctl Is UltraWinEditors.UltraNumericEditor Then
                DirectCast(child_ctl, UltraWinEditors.UltraNumericEditor).ReadOnly = tf
            End If
            If TypeOf child_ctl Is UltraWinEditors.UltraComboEditor Then
                DirectCast(child_ctl, UltraWinEditors.UltraComboEditor).ReadOnly = tf
            End If
            If TypeOf child_ctl Is UltraWinEditors.UltraOptionSet Then
                Dim optctl As UltraWinEditors.UltraOptionSet = DirectCast(child_ctl, UltraWinEditors.UltraOptionSet)
                optctl.Enabled = Not tf
                optctl.Appearance.BackColorDisabled = optctl.Appearance.BackColor
                optctl.Appearance.ForeColorDisabled = optctl.Appearance.ForeColor
            End If
            If TypeOf child_ctl Is UltraWinEditors.UltraCheckEditor Then
                Dim chkctl As UltraWinEditors.UltraCheckEditor = DirectCast(child_ctl, UltraWinEditors.UltraCheckEditor)
                chkctl.Enabled = Not tf
            End If
        End If
    End Sub

    Overridable Function Validate_Code_Special(
    ByVal COLUMN_NAME As String,
    ByVal IsValid As Boolean) As Boolean
        Return IsValid
    End Function

    Function Validate_Code(
    ByVal COLUMN_NAME As String,
    Optional ByVal should_not_exist As Boolean = False,
    Optional ByVal valid_if_blank As Boolean = False) As Boolean
        ' THE LINE BELOW WILL EXECUTE Validate_Code_Special twice because Validate_Code with the Nothing in it does a call to Validate_Code_Special
        'Return Validate_Code_Special(COLUMN_NAME, Validate_Code(COLUMN_NAME, Nothing, should_not_exist))
        Return Validate_Code(COLUMN_NAME, Nothing, should_not_exist, valid_if_blank)

    End Function

    Function Validate_Code(
    ByVal COLUMN_NAME As String,
    ByVal Precedent_Key_Values As List(Of String),
    Optional ByVal should_not_exist As Boolean = False,
    Optional ByVal valid_if_blank As Boolean = False) As Boolean

        cdr = Nothing

        Dim ctl As Control = Absx1.CtlFor(COLUMN_NAME)
        Dim CODE_VALUE As String = ctl.Text
        If TypeOf (ctl) Is UltraWinEditors.UltraComboEditor Then
            CODE_VALUE = Absx1.cbeFor(COLUMN_NAME).Value
        End If

        Dim COLUMN_CAPTION As String = Absx1.GetABSColumnCaption(ctl)
        If COLUMN_CAPTION = "" Then
            COLUMN_CAPTION = ASCMAIN1.Make_Caption(COLUMN_NAME)
        End If

        If CODE_VALUE = "" Then
            If Not valid_if_blank Then
                EMsg &= vbCr & "No Value Specified for " & COLUMN_CAPTION
            End If
        Else
            'cdr = LookUp(TABLE_NAME, CODE_VALUE)
            cdr = LookUp_for_txtctl(ctl, Precedent_Key_Values)
            If cdr Is Nothing Then
                If should_not_exist Then
                    Return Validate_Code_Special(COLUMN_NAME, True)
                Else
                    EMsg &= vbCr & "Invalid Value Specified for " & COLUMN_CAPTION & " (" & CODE_VALUE & ")"
                    Return False
                End If
            Else
                If should_not_exist Then
                    EMsg &= vbCr & "Record Exists for Value Specified (" & CODE_VALUE & ")"
                    Return False
                Else
                    Call Populate_Dependent_Controls(COLUMN_NAME, cdr)
                    Return Validate_Code_Special(COLUMN_NAME, True)
                End If
            End If
        End If
    End Function

    Public Overridable Sub Delete_Rows(
    ByVal TABLE_NAME As String)
        With dst.Tables(TABLE_NAME)
            'For Each row As DataRow In .Rows
            '    row.Delete()
            'Next
            For i As Long = .Rows.Count - 1 To 0 Step -1
                .Rows(i).Delete()
            Next
        End With
    End Sub

    Sub Delete_Rows(
    ByVal TABLE_NAME As String,
    ByVal sql_where As String)

        Dim NumberOfKeyFields As Integer = dst.Tables(TABLE_NAME).PrimaryKey.Length

        Dim KeysToDelete As New List(Of Object)

        For Each row As DataRow In dst.Tables(TABLE_NAME).Select(sql_where, "")
            Dim Keys As New List(Of Object)
            For i As Integer = 0 To NumberOfKeyFields - 1
                Keys.Add(row(i).ToString)
            Next
            KeysToDelete.Add(Keys)
        Next

        If KeysToDelete.Count <> 0 Then
            For Each Keys As List(Of Object) In KeysToDelete
                Dim row As DataRow =
                dst.Tables(TABLE_NAME).Rows.Find(Keys.ToArray)
                row.Delete()
            Next
        End If

    End Sub

    Public Sub Save_Header_Fields(ByVal c As Control, Optional ByVal Clear As Boolean = True)
        If Clear Then
            HFs.Clear()
        End If

        For Each ctl As Control In c.Controls
            If ctl.Controls.Count > 0 Then
                Save_Header_Fields(ctl, False)
            End If
            Dim COLUMN_NAME As String = Absx1.GetABSColumnName(ctl)
            If COLUMN_NAME <> "" Then
                If TypeOf ctl Is UltraWinEditors.UltraTextEditor Then
                    Dim txtctl As UltraWinEditors.UltraTextEditor = DirectCast(ctl, UltraWinEditors.UltraTextEditor)
                    HFs.Add(COLUMN_NAME, txtctl.Text)
                End If
                If TypeOf ctl Is UltraWinEditors.UltraOptionSet Then
                    Dim optctl As UltraWinEditors.UltraOptionSet = DirectCast(ctl, UltraWinEditors.UltraOptionSet)
                    HFs.Add(COLUMN_NAME, optctl.Value & "")
                End If
                If TypeOf ctl Is UltraWinEditors.UltraNumericEditor Then
                    Dim numctl As UltraWinEditors.UltraNumericEditor = DirectCast(ctl, UltraWinEditors.UltraNumericEditor)
                    HFs.Add(COLUMN_NAME, numctl.Value & "")
                End If
                If TypeOf ctl Is UltraWinEditors.UltraDateTimeEditor Then
                    Dim dtectl As UltraWinEditors.UltraDateTimeEditor = DirectCast(ctl, UltraWinEditors.UltraDateTimeEditor)
                    HFs.Add(COLUMN_NAME, dtectl.Value & "")
                End If
                If TypeOf ctl Is UltraWinSchedule.UltraCalendarCombo Then
                    Dim dtectl As UltraWinSchedule.UltraCalendarCombo = DirectCast(ctl, UltraWinSchedule.UltraCalendarCombo)
                    HFs.Add(COLUMN_NAME, dtectl.Value & "")
                End If
                If TypeOf ctl Is UltraWinEditors.UltraComboEditor Then
                    Dim cbectl As UltraWinEditors.UltraComboEditor = DirectCast(ctl, UltraWinEditors.UltraComboEditor)
                    HFs.Add(COLUMN_NAME, cbectl.Value & "")
                End If
                If TypeOf ctl Is UltraWinGrid.UltraCombo Then
                    Dim cmbctl As UltraWinGrid.UltraCombo = DirectCast(ctl, UltraWinGrid.UltraCombo)
                    HFs.Add(COLUMN_NAME, cmbctl.Value)
                End If
                If TypeOf ctl Is UltraWinEditors.UltraCheckEditor Then
                    Dim chkctl As UltraWinEditors.UltraCheckEditor = DirectCast(ctl, UltraWinEditors.UltraCheckEditor)
                    HFs.Add(COLUMN_NAME, IIf(chkctl.Checked, "1", "0"))
                End If
            End If
        Next

    End Sub

    Function Add_PK(ByVal COLUMN_NAME) As List(Of String)
        Return Add_PK(New String() {COLUMN_NAME})
    End Function

    Function Add_PK(ByVal COLUMN_NAMEs() As String) As List(Of String)
        Dim PKs As New List(Of String)
        For i As Integer = 0 To UBound(COLUMN_NAMEs)

            Dim COLUMN_NAME_pk As String = COLUMN_NAMEs(i)
            Dim COLUMN_NAME_alias As String = COLUMN_NAMEs(i)
            If COLUMN_NAME_pk.Contains(":") Then
                COLUMN_NAME_pk = Split(COLUMN_NAMEs(i), ":")(0)
                COLUMN_NAME_alias = Split(COLUMN_NAMEs(i), ":")(1)
            End If

            PKs.Add(Absx1.txtFor(COLUMN_NAME_pk).Text)
            'PKs.Add(Absx1.txtFor(COLUMN_NAMEs(i)).Text)
        Next
        'PKs.Add(Absx1.txtFor(COLUMN_NAME).Text)
        Return PKs
    End Function

    Sub BeginTrans(Optional ByVal Begin_Message As String = "")
        Me.Cursor = Cursors.WaitCursor
        T = ASCMAIN1.oraCon.BeginTransaction

        If Begin_Message <> "" Then
            ASCMAIN1.Progress(Begin_Message)
        End If

    End Sub

    Sub CommitTrans(Optional ByVal Commit_Message As String = "")
        T.Commit()
        Me.Cursor = Cursors.Default
        If Commit_Message <> "" And Not remotely_controlled Then
            MsgBox(Commit_Message, MsgBoxStyle.OkOnly, "Verification")
        End If
    End Sub

    Sub Rollback(
    Optional ByVal Error_Message As String = "",
    Optional ByRef ex2Record As Exception = Nothing)

        Try
            If ex2Record IsNot Nothing Then
                Call Record_Exception(ex2Record)
            End If

            T.Rollback()
            If Error_Message <> "" Then
                MsgBox(Error_Message, MsgBoxStyle.OkOnly, "Update Rolled Back")
            End If
        Catch ex As Exception
            MsgBox("Please call ABS", MsgBoxStyle.OkOnly, "Error trying to process Rollback")
        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub grd_AfterRowInsert(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs)
        If e.Row.Band.SortedColumns.Tag = "N" Then
            e.Row.Band.SortedColumns.Tag = ""
            e.Row.Band.SortedColumns.Clear()
        Else
            e.Row.Band.SortedColumns.RefreshSort(False)
            e.Row.Band.Layout.RowScrollRegions(0).Scroll(UltraWinGrid.RowScrollAction.Bottom)
            'Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
            'grd.DisplayLayout.RowScrollRegions(0).Scroll(UltraWinGrid.RowScrollAction.Bottom)
        End If
    End Sub

    Private Sub grd_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs)
        Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        If Not grdMRUs.ContainsKey(grd.Name) Then Exit Sub
        For Each COLUMN_NAME As String In grdMRUs(grd.Name)
            'If grd.ActiveRow IsNot Nothing Then
            '    With grd.ActiveRow

            If e.Row IsNot Nothing Then
                With e.Row
                    If .Band.Columns.Contains(COLUMN_NAME) Then
                        Try
                            If .Cells(COLUMN_NAME).Text <> "" Then
                                If ASCMAIN1.MRUs(COLUMN_NAME).Contains(.Cells(COLUMN_NAME).Text) Then
                                    ASCMAIN1.MRUs(COLUMN_NAME).Remove(.Cells(COLUMN_NAME).Text)
                                End If
                                ASCMAIN1.MRUs(COLUMN_NAME).Add(.Cells(COLUMN_NAME).Text)
                            End If
                        Catch ex As Exception

                        End Try
                    End If
                End With
            End If
        Next
    End Sub

    Sub CalculateTotals(grd As UltraWinGrid.UltraGrid)

        With grd
            If .Selected.Cells.Count = 0 Then
                .Text = "Total Selected Cells"
            Else

                Dim L As Integer = .DisplayLayout.Bands(0).LevelCount

                Dim T() As Decimal
                ReDim T(L - 1)
                For Each cell As UltraWinGrid.UltraGridCell In .Selected.Cells
                    Dim LX As Integer = cell.Column.Level
                    T(LX) += Val(cell.Value & "")
                Next
                Dim LZ As String = ""
                For II As Integer = 0 To L - 1
                    If T(II) = Val(Format(T(II), "##0")) Then
                        LZ &= ", " & Format(T(II), "#,##0") ' CStr(T(II))
                    Else
                        LZ &= ", " & Format(T(II), "#,##0.00####") '  CStr(T(II))
                    End If


                Next
                Dim totals As String = $"Total Selected Cells ({ .Selected.Cells.Count}) = {Mid(LZ, 3)}"
                ASCMAIN1.Progress(totals)
            End If
        End With
    End Sub

    Private Sub grd_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)

        Dim sbt As UltraWinToolbars.StateButtonTool = DirectCast(ASFMAIN1.UltraToolbarsManager1.Tools("Show Total of Selected Grid Cells"), UltraWinToolbars.StateButtonTool)

        If sbt.Checked Then
            CalculateTotals(sender)
        End If

        Static cell As UltraWinGrid.UltraGridCell = Nothing

        Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)



        If grd.ActiveCell IsNot Nothing And e.Button = Windows.Forms.MouseButtons.Left Then

            'problem where check box takes several clicks before checking: 
            'grd.DisplayLayout.Bands(0).Columns("PROM_NON_QUAL").CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText

            'grd.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.CellSelect
            'the line of code above fixes the problem where a check box is selected 
            '  when you click the eggwhites but not the yolk 
            ' - you should fix this in the grid designer for the column (not for the whole grid)

            'IMPORTANT - forget all of the malarky above - this problem may be easily solved by changing the default value of the column to "0"

            Dim pt As System.Drawing.Point = New System.Drawing.Point(e.X, e.Y)
            Dim elem As Infragistics.Win.UIElement = grd.DisplayLayout.UIElement.ElementFromPoint(pt)
            If elem Is Nothing Then
                Exit Sub
            End If
            If elem.GetType.Equals(GetType(Infragistics.Win.UltraWinGrid.CellUIElement)) _
            Or elem.GetType.Equals(GetType(Infragistics.Win.CheckEditorCheckBoxUIElement)) _
            Or elem.GetType.Equals(GetType(Infragistics.Win.UltraWinGrid.RowUIElement)) _
            Or elem.GetType.Equals(GetType(Infragistics.Win.CheckIndicatorUIElement)) Then '  .RowSelectorHeaderUIElement)) Then
                If grd.ActiveCell.Column.Style = UltraWinGrid.ColumnStyle.CheckBox _
                And grd.ActiveCell.CanEnterEditMode And grd.DisplayLayout.Override.AllowUpdate <> DefaultableBoolean.False Then
                    'grd.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
                    Dim UPDATE_NOW As Boolean = False
                    ' CHANGE MADE TO MAKE WHTIADJ1 WORK WITH DEL & SEL CHECKBOXES
                    If grd.ActiveRow.DataChanged Then ' If Not grd.ActiveRow.DataChanged Then
                        UPDATE_NOW = True
                    End If

                    Dim others As Boolean = False
                    For Each Grid_Cell As UltraWinGrid.UltraGridCell In grd.ActiveCell.Row.Cells
                        If Not Grid_Cell.Column.Hidden Then
                            If Grid_Cell.Column.Style <> UltraWinGrid.ColumnStyle.CheckBox Then
                                ' - UPDATE 05/08/08 - changing check for others to include whether they are editable
                                If Grid_Cell.Column.CellActivation = UltraWinGrid.Activation.AllowEdit Then
                                    others = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next

                    'NEXT LINE NEEDS TO BE DISABLED SO THAT WE CAN CLICK DEAD CENTER OF THE CHECK BOX (SEE PYMT APPL BANKS)
                    ' NEED TO ELIMINATE THE IF SO THAT JHI CLICKS IN RSFRETL1 CHKBOXES WORK
                    'If others Then
                    If Math.Abs(Val(grd.ActiveCell.Value & "") - 1).ToString <> grd.ActiveCell.OriginalValue & "" Then
                        grd.ActiveCell.Value = Math.Abs(Val(grd.ActiveCell.Value & "") - 1).ToString
                    End If
                    'End If

                    If UPDATE_NOW Then
                        grd.UpdateData()
                    End If
                End If
            End If

            If grd.ActiveCell Is Nothing Then
                Exit Sub
            End If

            If doubleclicked And Not grd.ActiveCell.IsInEditMode Then
                grd.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            End If

            Try
                If grd.ActiveCell.IsInEditMode And (cell Is Nothing Or Not grd.ActiveCell.Equals(cell)) Then
                    If grd.ActiveCell.Text <> "" And grd.ActiveCell.Editor IsNot Nothing AndAlso grd.ActiveCell.Editor.SupportsSelectableText Then
                        grd.ActiveCell.SelStart = 0
                        grd.ActiveCell.SelLength = grd.ActiveCell.Text.Length
                    End If
                    cell = grd.ActiveCell
                Else
                    cell = Nothing
                End If

            Catch ex As Exception
                cell = Nothing
            End Try
        End If

        doubleclicked = False
    End Sub

    Private Sub grd_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)

        Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)

        Try
            Dim pt As System.Drawing.Point = New System.Drawing.Point(e.X, e.Y)
            Dim elem As Infragistics.Win.UIElement
            elem = grd.DisplayLayout.UIElement.ElementFromPoint(pt)
            If elem IsNot Nothing Then
                If elem.GetType.Equals(GetType(Infragistics.Win.UltraWinGrid.RowSelectorHeaderUIElement)) Then

                    If e.Button = Windows.Forms.MouseButtons.Right Then
                        If grd.Rows.Count = 0 OrElse Not grd.Rows(0).IsGroupByRow Then
                            If grd.DisplayLayout.Override.AllowAddNew <> UltraWinGrid.AllowAddNew.No _
                            Or grd.DisplayLayout.Override.AllowUpdate <> DefaultableBoolean.False Then
                                Windows.Forms.Cursor.Position = New Point(-1, 0)
                                Excel_Import(grd)
                                'Excel_Import_SG(grd)
                            End If
                        End If
                    Else
                        ' And grd.DisplayLayout.Bands.Count = 1
                        If (ASCMAIN1.USER_ID = "pcarraro" Or ASCMAIN1.USER_ID = "wjz") And grd.DisplayLayout.Bands.Count = 1 And grd.DisplayLayout.Bands(0).Groups.Count < 1 Then  ' If grd.Rows.Count > 10000 Then ' this is total rows, not visible rows
                            Dim FILENAME As String = Excel_Export_to_SSG(grd)
                        Else
                            Dim gb As GemBox.Spreadsheet.ExcelFile = Excel_Export(grd)
                            Excel_Export_Post_Process(gb)
                            If FILENAMEs_to_Publish IsNot Nothing AndAlso FILENAMEs_to_Publish.Count > 0 Then
                                Dim FILENAME As String = FILENAMEs_to_Publish(FILENAMEs_to_Publish.Count - 1)
                                Excel_Export_Post_Process(FILENAME)
                            End If
                        End If
                    End If
                    Exit Sub
                End If
            End If


        Catch ex As Exception

        End Try

        Try

            If e.Button = Windows.Forms.MouseButtons.Right Then
                Dim mousePoint As Point = New Point(e.X, e.Y)
                Dim element As UIElement = CType(sender, UltraWinGrid.UltraGrid).DisplayLayout.UIElement.ElementFromPoint(mousePoint)
                If element IsNot Nothing Then
                    Dim cell As UltraWinGrid.UltraGridCell = CType(element.GetContext(GetType(UltraWinGrid.UltraGridCell)), UltraWinGrid.UltraGridCell)
                    If Not cell Is Nothing Then
                        Try
                            'grd.Selected.Rows.Clear()
                            'grd.Selected.Rows.Add(grd.ActiveRow)
                        Catch ex As Exception

                        End Try

                        grd.ActiveRow = CType(element.GetContext(GetType(UltraWinGrid.UltraGridRow)), UltraWinGrid.UltraGridRow)
                    Else

                        ' IF WE DO THE NEXT LINE, THEN RIGHT CLICK CONTEXT MENUS, LIKE IN ICTSTYL1.grdICTSTYC1, will unnecessarily cause activerows to then be de-activated, and all other child grids will now have no context
                        ' grd.ActiveRow = Nothing

                    End If
                End If

            End If

        Catch ex As Exception

        End Try
    End Sub

    Function Excel_Export_to_SSG(grd As UltraWinGrid.UltraGrid, Optional show_XLS As Boolean = True) As String
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Exporting to SSG")

        Dim FILENAME As String = ASCMAIN1.Next_Control_No("XNO")
        FILENAME = ASCMAIN1.Folders("Temp") & FILENAME & ".xlsx"

        Dim cols() As String
        ReDim cols(grd.DisplayLayout.Bands(0).Columns.Count - 1)

        Dim grdtbl As DataTable
        If TypeOf (grd.DataSource) Is System.Data.DataView Then
            grdtbl = DirectCast(grd.DataSource, DataView).ToTable
        Else
            If TypeOf (grd.DataSource) Is System.Data.DataSet Then
                Dim Ds As DataSet = DirectCast(grd.DataSource, DataSet)
                grdtbl = Ds.Tables(grd.DataMember)
            Else
                grdtbl = DirectCast(grd.DataSource, DataTable)
            End If

        End If

        Dim tbl As New DataTable
        With tbl
            For Each grdcol As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
                Dim vc As Integer = grdcol.Header.VisiblePosition
                If Not grdcol.Hidden Then
                    cols(vc) = grdcol.Key
                End If
            Next
            Dim column_names As New List(Of String)
            For vc As Integer = 0 To cols.Length - 1
                If cols(vc) <> "" Then
                    column_names.Add(cols(vc))
                    If Not grdtbl.Columns.Contains(cols(vc)) Then
                        .Columns.Add(cols(vc), GetType(String))
                    Else
                        .Columns.Add(cols(vc), grdtbl.Columns(cols(vc)).DataType)
                    End If

                End If
            Next
            For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                If Not grow.IsFilteredOut Then
                    Dim row As DataRow = tbl.NewRow
                    For Each COLUMN_NAME In column_names
                        row.Item(COLUMN_NAME) = grow.Cells(COLUMN_NAME).Value
                    Next
                    tbl.Rows.Add(row)
                End If
            Next

            Dim oWB As SpreadsheetGear.IWorkbook
            Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
            Dim range As SpreadsheetGear.IRange = Nothing

            oWB = SpreadsheetGear.Factory.GetWorkbook()

            oSheet = oWB.Worksheets(0)

            Dim c As Integer = -1
            For Each COLUMN_NAME In column_names
                With grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                    c += 1
                    oSheet.Cells(1, c).Value = .Header.Caption

                    oSheet.Cells(1, c).EntireColumn.ColumnWidth = .Width * 8 / 27 / 2

                    If tbl.Columns(COLUMN_NAME).DataType.ToString = "System.String" Then
                        oSheet.Cells(1, c).EntireColumn.NumberFormat = "@"
                    End If


                    If tbl.Columns(COLUMN_NAME).DataType.ToString = "System.DateTime" Then
                        oSheet.Cells(1, c).EntireColumn.NumberFormat = "MM/dd/yyyy"
                    End If

                    If .Format <> "" Then
                        oSheet.Cells(1, c).EntireColumn.NumberFormat = .Format
                    End If

                    If .CellAppearance.TextHAlign = HAlign.Center Then
                        oSheet.Cells(1, c).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    ElseIf .CellAppearance.TextHAlign = HAlign.Right Then
                        oSheet.Cells(1, c).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    End If

                    If .Header.Appearance.TextHAlign = HAlign.Center Then
                        oSheet.Cells(1, c).HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    ElseIf .Header.Appearance.TextHAlign = HAlign.Right Then
                        oSheet.Cells(1, c).HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    End If
                End With

            Next

            If grd.DisplayLayout.Bands(0).Summaries.Count > 0 Then
                For Each summary As UltraWinGrid.SummarySettings In grd.DisplayLayout.Bands(0).Summaries
                    Dim COLUMN_NAME As String = summary.SourceColumn.Key
                    Dim C2 As Integer = column_names.IndexOf(COLUMN_NAME)
                    Dim SUMMARY_TYPE As String = summary.SummaryType
                    If column_names.Contains(COLUMN_NAME) Then

                        If SUMMARY_TYPE = UltraWinGrid.SummaryType.Formula Then
                            oSheet.Cells(0, C2).Formula = $"=SUBTOTAL(9,{Excel_Cell0(1 + 1, C2)}:{Excel_Cell0(1 + tbl.Rows.Count, C2)})"
                        ElseIf SUMMARY_TYPE = UltraWinGrid.SummaryType.Sum Then
                            oSheet.Cells(0, C2).Formula = $"=SUBTOTAL(9,{Excel_Cell0(1 + 1, C2)}:{Excel_Cell0(1 + tbl.Rows.Count, C2)})"
                        ElseIf SUMMARY_TYPE = UltraWinGrid.SummaryType.Count Then
                            oSheet.Cells(0, C2).NumberFormat = "#,##0"
                            oSheet.Cells(0, C2).Formula = $"=SUBTOTAL(3,{Excel_Cell0(1 + 1, C2)}:{Excel_Cell0(1 + tbl.Rows.Count, C2)})"
                        End If
                    End If
                Next
            End If

            range = oSheet.Range(Excel_Cell0(2, 0))
            range.CopyFromDataTable(tbl, SpreadsheetGear.Data.SetDataFlags.NoColumnHeaders)

            oSheet.Range(1, 0, 1, column_names.Count - 1).Interior.Color = SpreadsheetGear.Colors.AliceBlue
            oSheet.Range(1, 0, 1, column_names.Count - 1).AutoFilter()

            oSheet.Cells(2, 0).Select()
            oSheet.WindowInfo.FreezePanes() = True

            oWB.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)

            If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" Then
                'oWB.SaveAs(Replace(FILENAME, ".xlsx", ".txt"), SpreadsheetGear.FileFormat.UnicodeText)
                'oWB.SaveAs("C:\VS\AHA\Temp\wjz.txt", SpreadsheetGear.FileFormat.UnicodeText)

                'Dim excel As New Microsoft.Office.Interop.Excel.Application
                'Dim wb As Microsoft.Office.Interop.Excel.Workbook = excel.Workbooks.Open(FILENAME)
                'wb.SaveAs(Replace(FILENAME, ".xlsx", ".txt"), FileFormat:=excel.XlFileFormat.xlText)
            End If


            If ASCMAIN1.JOB_STREAM_CODE = "" And show_XLS Then
                Show_Document(FILENAME)
            End If

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")

        End With

        Return FILENAME
    End Function
    Public Overridable Sub Excel_Export_Post_Process(FILENAME As String)

    End Sub
    Public Overridable Sub Excel_Export_Post_Process(f As GemBox.Spreadsheet.ExcelFile)

    End Sub
    Private Sub grd_MouseEnterElement(ByVal sender As Object, ByVal e As Infragistics.Win.UIElementEventArgs)

        'Dim grd As Infragistics.Win.UltraWinGrid.UltraGrid = DirectCast(sender, Infragistics.Win.UltraWinGrid.UltraGrid)
        'ASFMAIN1.UltraStatusBar1.Panels(8).Text = grd.name

        If e.Element.GetType().ToString() = "Infragistics.Win.UltraWinGrid.CellUIElement" Then
            Dim elem As Infragistics.Win.UIElement = e.Element
            If elem.GetType.Equals(GetType(Infragistics.Win.UltraWinGrid.CellUIElement)) Then
                'Dim C As UltraWinGrid.UltraGridCell = DirectCast(elem., UltraWinGrid.UltraGridCell)
                Dim cell As UltraWinGrid.UltraGridCell = CType(elem.GetContext(GetType(UltraWinGrid.UltraGridCell)), UltraWinGrid.UltraGridCell)
                If Not cell Is Nothing Then
                    CurrentGridColumn = cell.Column.Key
                    CurrentGridBand = cell.Band.Key
                    CurrentControl = DirectCast(sender, UltraWinGrid.UltraGrid)
                    timer.Stop()
                    timer.Start()
                End If
            End If
        End If
    End Sub

    Private Sub grd_MouseLeaveElement(ByVal sender As Object, ByVal e As Infragistics.Win.UIElementEventArgs)

        If TypeOf e.Element Is Infragistics.Win.UltraWinScrollBar.ScrollBarUIElement Then
            ASCMAIN1.GridDoubleClickAllowed = True
        End If


        ' if we are not leaving a cell, then don't anything
        If Not e.Element.GetType().Equals(GetType(UltraWinGrid.CellUIElement)) Then
            Exit Sub
        End If

        ' prevent the timer from ticking again
        timer.Stop()

        ' destroy the tooltip
        If Not tooltip Is Nothing Then
            tooltip.SetToolTip(Me, String.Empty)
            tooltip.Dispose()
            tooltip = Nothing
        End If
    End Sub

    Private Sub OnTimerTick(ByVal sender As Object, ByVal e As EventArgs)
        tooltip = New System.Windows.Forms.ToolTip()


        If dst IsNot Nothing Then
            If dst.Tables.Contains("ASTTTIP1") AndAlso dst.Tables("ASTTTIP1").PrimaryKey.Length <> 0 Then
                Try

                    Dim rowASTTTIP1 As DataRow = dst.Tables("ASTTTIP1").Rows.Find _
                        (New String() {Me.Name, CurrentGridBand, CurrentGridColumn})
                    If rowASTTTIP1 IsNot Nothing Then
                        Dim TTI As New UltraWinToolTip.UltraToolTipInfo
                        TTI.ToolTipTitle = rowASTTTIP1.Item("TOOLTIP_TITLE") & ""
                        TTI.ToolTipTextFormatted = rowASTTTIP1.Item("TOOLTIP_TEXT") & ""
                        tip.SetUltraToolTip(CurrentControl, TTI)
                        tip.ShowToolTip(CurrentControl)
                    End If

                    'tooltip.SetToolTip(Me, tooltip_msg)
                    'tooltip.ToolTipTitle = tooltip_title
                    'tooltip.AutoPopDelay = 12000
                Catch ex As Exception

                End Try
            End If
        End If
        ' once the timer has ticked, stop it
        timer.Stop()
    End Sub


    Sub Build_MRU_ContextMenu(ByVal COLUMN_NAME As String)

        If COLUMN_NAME = "" Then
            Exit Sub

        Else
            ASCMAIN1.MRU_COLUMN_NAME = COLUMN_NAME
            Dim PopupMenuTool As UltraWinToolbars.PopupMenuTool = ASFMAIN1.UltraToolbarsManager1.Tools("txtMenu")
            PopupMenuTool.ShowPopup()
            Exit Sub
        End If

        'Dim cMenu As ContextMenu = New ContextMenu

        'Dim ListTool As UltraWinToolbars.ListTool = New UltraWinToolbars.ListTool("Most Recently Used")
        'Dim ComboBoxTool As UltraWinToolbars.ComboBoxTool = New UltraWinToolbars.ComboBoxTool("History")
        '' Dim PopupMenuTool As UltraWinToolbars.PopupMenuTool = New UltraWinToolbars.PopupMenuTool("txtMenu")
        'Dim LabelTool As UltraWinToolbars.LabelTool = New UltraWinToolbars.LabelTool("COLUMN_NAME")

        'ListTool.SharedProps.Caption = "Most Recently Used"
        'ComboBoxTool.SharedProps.Caption = "History"

        'PopupMenuTool.SharedProps.Caption = "txtMenu"
        'ComboBoxTool.InstanceProps.IsFirstInGroup = True
        'PopupMenuTool.Tools.AddRange(New Infragistics.Win.UltraWinToolbars.ToolBase() {LabelTool, ComboBoxTool, ListTool})
        'LabelTool.SharedProps.Caption = "COLUMN_NAME"


        'PopupMenuTool = ASFMAIN1.UltraToolbarsManager1.Tools("txtMenu")

        ''LabelTool.SharedProps.Caption = ASCMAIN1.Make_Caption(COLUMN_NAME)
        'PopupMenuTool.Tools("COLUMN_NAME").SharedProps.Caption = ASCMAIN1.Make_Caption(COLUMN_NAME)
        'Dim LL As List(Of String) = ASCMAIN1.MRUs(COLUMN_NAME)
        'If LL.Count = 0 Then
        '    Exit Sub
        'End If

        'Dim ValueList As ValueList = New ValueList(0)
        'For i As Integer = 0 To LL.Count - 1
        '    ValueList.ValueListItems.Add(New ValueListItem(LL(i)))
        'Next
        'Dim ComboBoxTool As UltraWinToolbars.ComboBoxTool = PopupMenuTool.Tools("History")
        'ComboBoxTool.ValueList = ValueList
        'ComboBoxTool.SelectedIndex = 0
        ''comboboxtool.DropDownStyle = DropDownStyle.DropDown
        'ComboBoxTool.DropDownStyle = DropDownStyle.DropDownList
        'ComboBoxTool.Value = "<Select Value>"
        'ComboBoxTool.SharedProps.ToolTipText = "Code Values Entered sorted from First to Last"
        'ComboBoxTool.AutoComplete = True

        'Dim ListTool As UltraWinToolbars.ListTool = PopupMenuTool.Tools("Most Recently Used")
        'ListTool.ListToolItems.Clear()
        'If LL.Count > 0 Then
        '    For i As Integer = LL.Count - 1 To 0 Step -1
        '        listtool.ListToolItems.Add(LL(i).ToString, LL(i).ToString)
        '        If LL.Count - i > 3 Then
        '            Exit For
        '        End If
        '    Next
        'End If

        'PopupMenuTool.ShowPopup()
        ''        cMenu.Show(UltraGrid1, mousePoint)


    End Sub

    Private Sub grd_AfterEnterEditMode(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)

        If grd.ActiveCell.IsInEditMode Then
            disable_arrows = False
        Else
            disable_arrows = True
        End If

    End Sub

    Private Sub grd_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs)

        If e.Cell.Column.Style = UltraWinGrid.ColumnStyle.CheckBox Then
            ''Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
            ''grd.UpdateData()

            ' if there are other non-checkbox columns in the grid, then do not do the update
            ' (instead of doing this, and the block below where we check .Tag = "N", perhaps we should update the grid only if setting .Tag = "Y" - just a thought

            Dim others As Boolean = False
            For Each Grid_Cell As UltraWinGrid.UltraGridCell In e.Cell.Row.Cells
                If Not Grid_Cell.Column.Hidden Then
                    If Grid_Cell.Column.Style <> UltraWinGrid.ColumnStyle.CheckBox Then
                        ' - UPDATE 05/08/08 - changing check for others to include whether they are editable
                        If Grid_Cell.Column.CellActivation = UltraWinGrid.Activation.AllowEdit Then
                            others = True
                            Exit For
                        End If
                    End If
                End If
            Next

            If e.Cell.Column.Tag <> "N" And Not others Then
                e.Cell.Row.Update()
            End If
            ''grd.PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)
        End If
    End Sub


    Public Overridable Sub cmb_BeforeDropDown(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)

    End Sub

    Private Sub cmb_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs)

    End Sub

    Private Sub cmb_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        ASCMAIN1.MRU_used = False
        ASCMAIN1.MRU_COLUMN_NAME = ""

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(DirectCast(sender, Control))
        If ASCMAIN1.MRUs.ContainsKey(COLUMN_NAME) Then
            ASCMAIN1.MRU_cmbctl = DirectCast(sender, UltraWinGrid.UltraCombo)
            If Not ASCMAIN1.MRU_cmbctl.ReadOnly Then
                ASCMAIN1.MRU_COLUMN_NAME = COLUMN_NAME
            End If
        End If
    End Sub

    Public Overridable Sub cmb_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            SendKeys.Send(Chr(9))
        End If
    End Sub

    Private Sub cmb_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        ASCMAIN1.MRU_COLUMN_NAME = ""

        Dim cmbctl As UltraWinGrid.UltraCombo
        cmbctl = DirectCast(sender, UltraWinGrid.UltraCombo)

        '        If cmbctl.Modified Then
        COLUMN_NAME = Absx1.GetABSColumnName(cmbctl)
        cmbctl.Text = Trim(cmbctl.Text)
        If cmbctl.Text <> "" Then
            cmbctl.Text = ASCMAIN1.Format_Field(cmbctl.Text, COLUMN_NAME, tblASFBASE1)
        End If

        If COLUMN_NAME <> "" Then
            Call Populate_Controls_with_Parents(COLUMN_NAME, cmbctl)
        End If
        '        End If
    End Sub
    ' Private Sub cmb_ValueChanged
    Public Overridable Sub cmb_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cmbctl As UltraWinGrid.UltraCombo = DirectCast(sender, UltraWinGrid.UltraCombo)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(cmbctl)

        If Mid(Me.Name, 3, 1) = "R" Then
            If cmbYP.ContainsKey(COLUMN_NAME) Then
                Dim cmbYPparm As New cmbYPparms
                cmbYPparm = cmbYP(COLUMN_NAME)

                If cmbYPparm.Child_cmbYP <> "" Then
                    Dim cmbYPparm_Child As New cmbYPparms
                    cmbYPparm_Child = cmbYP(cmbYPparm.Child_cmbYP)

                    Dim cmbctl_child As UltraWinGrid.UltraCombo
                    cmbctl_child = Absx1.cmbFor(cmbYPparm.Child_cmbYP)

                    If cmbctl.Text = "" Then
                        cmbctl_child.Text = ""
                    Else
                        Dim BYP As String = cmbctl.ActiveRow.Cells("OPS_YYYYPP").Text
                        Dim sql_where As String = "OPS_YYYYPP >= '" & BYP & "' and OPS_YYYYPP <= '" & ASCMAIN1.Period_Calc(BYP, cmbYPparm_Child.TotalRelativePeriods - 1) & "'"
                        Dim Sql As String = ASCMAIN1.CodeSelector.Get_SQL(Absx1.GetABSViewName(cmbctl), Absx1.GetABSLookUpTableName(cmbctl), sql_where)

                        Dim tbl As DataTable = ASCDATA1.GetDataTable(Sql)
                        cmbctl_child.DataSource = tbl
                        cmbctl_child.ActiveRow = cmbctl_child.Rows(cmbYPparm_Child.RelativeDefaultPeriod)
                    End If
                End If

                If cmbYWparm.Child_cmbYW <> "" Then
                    Dim cmbYWparm_Child As New cmbYWparms
                    cmbYWparm_Child = cmbYW(cmbYWparm.Child_cmbYW)

                    Dim cmbctl_child As UltraWinGrid.UltraCombo
                    cmbctl_child = Absx1.cmbFor(cmbYWparm.Child_cmbYW)

                    If cmbctl.Text = "" Then
                        cmbctl_child.Text = ""
                    Else
                        Dim BYW As String = cmbctl.ActiveRow.Cells("OPS_YYYYWW").Text
                        Dim sql_where As String = "YYYYWW >= '" & BYW & "' and YYYYWW <= '" & ASCMAIN1.Week_Calc(BYW, cmbYWparm_Child.TotalRelativeWeeks - 1) & "'"
                        Dim Sql As String = ASCMAIN1.CodeSelector.Get_SQL(Absx1.GetABSViewName(cmbctl), Absx1.GetABSLookUpTableName(cmbctl), sql_where)

                        Dim tbl As DataTable = ASCDATA1.GetDataTable(Sql)
                        cmbctl_child.DataSource = tbl
                        cmbctl_child.ActiveRow = cmbctl_child.Rows(cmbYWparm_Child.RelativeDefaultWeek)
                    End If
                End If


            ElseIf cmbYW.ContainsKey(COLUMN_NAME) Then
                Dim cmbYWparm As New cmbYWparms
                cmbYWparm = cmbYW(COLUMN_NAME)

                If cmbYWparm.Child_cmbYW <> "" Then
                    Dim cmbYWparm_Child As New cmbYWparms
                    cmbYWparm_Child = cmbYW(cmbYWparm.Child_cmbYW)

                    Dim cmbctl_child As UltraWinGrid.UltraCombo
                    cmbctl_child = Absx1.cmbFor(cmbYWparm.Child_cmbYW)

                    If cmbctl.Text = "" Then
                        cmbctl_child.Text = ""
                    Else
                        Dim BYW As String = cmbctl.ActiveRow.Cells("YYYYWW").Text
                        Dim sql_where As String = "YYYYWW >= '" & BYW & "' and YYYYWW <= '" & ASCMAIN1.Period_Calc(BYW, cmbYWparm_Child.TotalRelativeWeeks - 1) & "'"
                        Dim Sql As String = ASCMAIN1.CodeSelector.Get_SQL(Absx1.GetABSViewName(cmbctl), Absx1.GetABSLookUpTableName(cmbctl), sql_where)

                        Dim tbl As DataTable = ASCDATA1.GetDataTable(Sql)
                        cmbctl_child.DataSource = tbl
                        cmbctl_child.ActiveRow = cmbctl_child.Rows(cmbYWparm_Child.RelativeDefaultWeek)
                    End If
                End If
            End If
        Else
            If tblASFBASE1 IsNot Nothing Then
                If tblASFBASE1.Columns.Count > 0 Then
                    If Not cmbctl.Focused Or ASCMAIN1.MRU_used Then
                        If COLUMN_NAME <> "" Then
                            Call Populate_Controls_with_Parents(COLUMN_NAME, cmbctl)
                        End If
                        ASCMAIN1.MRU_used = False
                    End If
                End If
            End If
        End If

    End Sub

    Public Sub Load_Drop_Down(
    ByVal COLUMN_NAME As String,
    Optional ByVal sql_where As String = "")

        Dim ctl As Control = Absx1.CtlFor(COLUMN_NAME)

        VIEW_NAME = Absx1.GetABSViewName(ctl)
        If VIEW_NAME = "" Then
            VIEW_NAME = COLUMN_NAME
        End If
        TABLE_NAME_view = Absx1.GetABSLookUpTableName(ctl)
        Dim cmbctl As UltraWinGrid.UltraCombo = DirectCast(ctl, UltraWinGrid.UltraCombo)
        Dim sql As String = ASCMAIN1.CodeSelector.Get_SQL(VIEW_NAME, TABLE_NAME_view, sql_where)


        If cmbctl.DataSource Is Nothing Then
            cmbctl.DataSource = ASCDATA1.GetDataTable(sql)
            For I As Integer = 0 To ASCMAIN1.CodeSelector.grdColumns.Count - 1
                cmbctl.DisplayLayout.Bands(0).Columns(I).Header.Caption = ASCMAIN1.CodeSelector.grdColumns(I).Item("COLUMN_CAPTION")
                cmbctl.DisplayLayout.Bands(0).Columns(I).Width = ASCMAIN1.CodeSelector.grdColumns(I).Item("COLUMN_WIDTH")
            Next
            cmbctl.DropDownStyle = UltraWinGrid.UltraComboStyle.DropDownList
            cmbctl.DisplayLayout.Bands(0).SortedColumns.Clear()
            cmbctl.DisplayLayout.Bands(0).SortedColumns.Add(cmbctl.DisplayLayout.Bands(0).Columns(0), False)
        Else
            cmbctl.DropDownStyle = UltraWinGrid.UltraComboStyle.DropDown
            Dim oldvalue As String = cmbctl.Value
            Dim oldtext As String = cmbctl.Text
            cmbctl.DropDownStyle = UltraWinGrid.UltraComboStyle.DropDownList
            Dim tbl As DataTable = DirectCast(cmbctl.DataSource, DataTable)
            tbl.Rows.Clear()
            tbl.Load(ASCDATA1.GetDataTable(sql).CreateDataReader)
            cmbctl.Value = oldvalue
            cmbctl.Text = oldtext
            If cmbctl.Text <> oldtext And ASCMAIN1.Running_in_VS Then
                MsgBox("Could not set combo for " & COLUMN_NAME & " to " & oldtext)
            End If
        End If
    End Sub

    Private Sub cmb_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs)
        'Call txt_EditorButtonClick_Special(txtctl)
        'Call Leaving_txt_Special_Before(COLUMN_NAME, txtctl)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Call Populate_Controls_with_Parents(COLUMN_NAME, sender)
        'Call Leaving_txt_Special_After(COLUMN_NAME, txtctl)
    End Sub

    Public Function Set_cmbYP(
    ByVal COLUMN_NAME As String,
    ByVal Base_YYYYPP As String,
    ByVal RelativeStartingPeriod As Integer,
    ByVal RelativeEndingPeriod As Integer,
    ByVal RelativeDefaultPeriod As Integer
    ) As cmbYPparms

        cmbYPparm = New cmbYPparms
        cmbYPparm.Base_YYYYPP = Base_YYYYPP
        cmbYPparm.RelativeStartingPeriod = RelativeStartingPeriod
        cmbYPparm.RelativeEndingPeriod = RelativeEndingPeriod
        cmbYPparm.RelativeDefaultPeriod = RelativeDefaultPeriod

        If Not cmbYP.ContainsKey(COLUMN_NAME) Then
            cmbYP.Add(COLUMN_NAME, cmbYPparm)
        End If
        Return cmbYPparm

    End Function

    Public Function Set_cmbYP_Child(
    ByVal COLUMN_NAME As String,
    ByVal TotalRelativePeriods As Integer,
    ByVal Parent_cmbYP As String,
    Optional ByVal RelativeDefaultPeriod As Integer = -1
    ) As cmbYPparms

        cmbYPparm = New cmbYPparms
        cmbYPparm.TotalRelativePeriods = TotalRelativePeriods
        If RelativeDefaultPeriod = -1 Then
            cmbYPparm.RelativeDefaultPeriod = TotalRelativePeriods - 1
        Else
            If RelativeDefaultPeriod < 0 Then
                RelativeDefaultPeriod = 0
            ElseIf RelativeDefaultPeriod > TotalRelativePeriods - 1 Then
                RelativeDefaultPeriod = TotalRelativePeriods - 1
            End If
            cmbYPparm.RelativeDefaultPeriod = RelativeDefaultPeriod
        End If

        cmbYPparm.Parent_cmbYP = Parent_cmbYP

        Dim Parent_cmbYPparm As New cmbYPparms
        Parent_cmbYPparm = cmbYP(Parent_cmbYP)
        Parent_cmbYPparm.Child_cmbYP = COLUMN_NAME
        cmbYP(Parent_cmbYP) = Parent_cmbYPparm

        If Not cmbYP.ContainsKey(COLUMN_NAME) Then
            cmbYP.Add(COLUMN_NAME, cmbYPparm)
        End If
        Return cmbYPparm
    End Function

    Public Function Set_cmbYW(
    ByVal COLUMN_NAME As String,
    ByVal Base_YYYYWW As String,
    ByVal RelativeStartingWeek As Integer,
    ByVal RelativeEndingWeek As Integer,
    ByVal RelativeDefaultWeek As Integer
    ) As cmbYWparms

        cmbYWparm = New cmbYWparms
        cmbYWparm.Base_YYYYWW = Base_YYYYWW
        cmbYWparm.RelativeStartingWeek = RelativeStartingWeek
        cmbYWparm.RelativeEndingWeek = RelativeEndingWeek
        cmbYWparm.RelativeDefaultWeek = RelativeDefaultWeek

        If Not cmbYW.ContainsKey(COLUMN_NAME) Then
            cmbYW.Add(COLUMN_NAME, cmbYWparm)
        End If
        Return cmbYWparm

    End Function

    Public Function Set_cmbYW_Child(
    ByVal COLUMN_NAME As String,
    ByVal TotalRelativeWeeks As Integer,
    ByVal Parent_cmbYW As String,
    Optional ByVal RelativeDefaultWeek As Integer = -1
    ) As cmbYWparms

        cmbYWparm = New cmbYWparms
        cmbYWparm.TotalRelativeWeeks = TotalRelativeWeeks
        If RelativeDefaultWeek = -1 Then
            cmbYWparm.RelativeDefaultWeek = TotalRelativeWeeks - 1
        Else
            If RelativeDefaultWeek < 0 Then
                RelativeDefaultWeek = 0
            ElseIf RelativeDefaultWeek > TotalRelativeWeeks - 1 Then
                RelativeDefaultWeek = TotalRelativeWeeks - 1
            End If
            cmbYWparm.RelativeDefaultWeek = RelativeDefaultWeek
        End If

        cmbYWparm.Parent_cmbYW = Parent_cmbYW

        Dim Parent_cmbYWparm As New cmbYWparms
        Parent_cmbYWparm = cmbYW(Parent_cmbYW)
        Parent_cmbYWparm.Child_cmbYW = COLUMN_NAME
        cmbYW(Parent_cmbYW) = Parent_cmbYWparm

        If Not cmbYW.ContainsKey(COLUMN_NAME) Then
            cmbYW.Add(COLUMN_NAME, cmbYWparm)
        End If
        Return cmbYWparm
    End Function

#Region "GL"
    Sub Set_SEGS(ByVal grd As UltraWinGrid.UltraGrid, Optional ByVal BandKey As String = "")
        If BandKey = "" Then
            BandKey = grd.DisplayLayout.Bands(0).Key
        End If
        If Not ROWs.ContainsKey("GLTPARM1") Then
            Get_PARM("GLTPARM1")
        End If
        With grd.DisplayLayout.Bands(BandKey)
            For I As Integer = 2 To 4
                Dim Z As String = "SEG" & CStr(I)
                If ROWs("GLTPARM1").Item("GL_PARM_" & Z & "_DESC") & "" = "" Then
                    .Columns(Z & "_CODE").Hidden = True
                Else
                    .Columns(Z & "_CODE").Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_" & Z & "_DESC")
                End If
            Next
        End With
    End Sub

    Sub Add_ACCT_TYPEs(
    ByVal TABLE_NAME As String,
    Optional ByVal include_ACCT_TYPE_SEQ As Boolean = False)

        Call Add_ACCT_TYPE(TABLE_NAME, "A", "Asset", include_ACCT_TYPE_SEQ, 1)
        Call Add_ACCT_TYPE(TABLE_NAME, "L", "Liability", include_ACCT_TYPE_SEQ, 2)
        Call Add_ACCT_TYPE(TABLE_NAME, "E", "Equity", include_ACCT_TYPE_SEQ, 3)
        Call Add_ACCT_TYPE(TABLE_NAME, "I", "Income", include_ACCT_TYPE_SEQ, 4)
        Call Add_ACCT_TYPE(TABLE_NAME, "X", "Expense", include_ACCT_TYPE_SEQ, 5)
    End Sub

    Sub Add_ACCT_TYPE(
    ByVal TABLE_NAME As String,
    ByVal ACCT_TYPE As String,
    ByVal ACCT_TYPE_DESC As String,
    ByVal include_ACCT_TYPE_SEQ As Boolean,
    ByVal ACCT_TYPE_SEQ As Integer)
        Dim dr As DataRow = dst.Tables(TABLE_NAME).NewRow
        dr.Item("ACCT_TYPE") = ACCT_TYPE
        dr.Item("ACCT_TYPE_DESC") = ACCT_TYPE_DESC
        If include_ACCT_TYPE_SEQ Then
            dr.Item("ACCT_TYPE_SEQ") = ACCT_TYPE_SEQ
        End If
        dst.Tables(TABLE_NAME).Rows.Add(dr)
    End Sub

    Sub Update_GLTACCT3(ByVal JOURNAL_NO As String, ByVal OPS_YYYYPP As String)
        If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
            ASCDATA1.ExecuteSQL("EXEC [GLPDETL1_POST_SUMMARY] @JOURNAL_NO_IN = N'" & JOURNAL_NO & "', @OPS_YYYYPP_IN = N'" & OPS_YYYYPP & "'")
            ASCDATA1.ExecuteSQL("EXEC [GLPACCT3_J] @JOURNAL_NO_IN = N'" & JOURNAL_NO & "', @OPS_YYYYPP_IN = N'" & OPS_YYYYPP & "', @SGN = N'1'")
        Else
            ASCDATA1.ExecuteSQL("Begin GLPDETL1_POST_SUMMARY ('" & JOURNAL_NO & "','" & OPS_YYYYPP & "'); End;")
            ASCDATA1.ExecuteSQL("Begin GLPACCT3_J ('" & JOURNAL_NO & "','" & OPS_YYYYPP & "',1); End;")
        End If
    End Sub

    Sub InterCompany(ByVal JOURNAL_NO As String, ByVal OPS_YYYYPP As String)
        Dim sqlx As String = "OPS_YYYYPP = '" & OPS_YYYYPP & "' and JOURNAL_NO = '" & JOURNAL_NO & "'"
        Dim JOURNAL_LNO As Integer = Val(dst.Tables("GLTDETL1").Compute("MAX(JOURNAL_LNO)", sqlx) & "")
        For i As Integer = 2 To 4
            Dim ACCT_SEG_ID As String = CStr(i)
            If ROWs("GLTPARM1").Item("GL_PARM_SEG" & CStr(i) & "_REC") & "" = "1" Then
                Dim COLUMN_NAME As String = "SEG" & CStr(i) & "_CODE"
                Dim rowGLTJRNL1 As DataRow() = dst.Tables("GLTJRNL1").Select("JOURNAL_NO = '" & JOURNAL_NO & "'")
                ' in a true >2 company scenario, we may want companies to interact with eachother instead of with the default company
                ' in that case, we would need to set up journal ownership in the code that created GLTJRNL1
                ' if only 2 companies in a J/E, then either could become the owner (suggest the 1st one we see)
                ' if >2 companies, then we need to pick one, and perhaps would pick the default company
                Dim JOURNAL_OWNER As String = rowGLTJRNL1(0).Item(COLUMN_NAME) & ""
                If JOURNAL_OWNER = "" Then
                    JOURNAL_OWNER = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i))
                End If
                Dim ACCT_SEG_REC_ACCT_JOURNAL_OWNER As String =
                LookUp("GLTSEGM1", New String() {ACCT_SEG_ID, JOURNAL_OWNER}, True).Item("ACCT_SEG_REC_ACCT") & ""

                For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("GLTDETL1") _
                    .Select(sqlx & " and " & COLUMN_NAME & " <> '" & JOURNAL_OWNER & "'"), COLUMN_NAME).Rows
                    Dim ACCT_SEG_CODE = row.Item(COLUMN_NAME)
                    If ACCT_SEG_CODE <> JOURNAL_OWNER Then
                        Dim sqly As String = sqlx & " and " & COLUMN_NAME & " = '" & ACCT_SEG_CODE & "'"
                        Dim DETL_POSTING_AMT As Double = Val(dst.Tables("GLTDETL1").Compute("SUM (DETL_POSTING_AMT)", sqly) & "")
                        DETL_POSTING_AMT = Round(DETL_POSTING_AMT, 2)
                        If DETL_POSTING_AMT <> 0 Then
                            Dim ACCT_SEG_REC_ACCT As String =
                            LookUp("GLTSEGM1", New String() {ACCT_SEG_ID, ACCT_SEG_CODE}).Item("ACCT_SEG_REC_ACCT")
                            For j As Integer = 0 To 1
                                Dim ACCT_CODE As String = New String() {ACCT_SEG_REC_ACCT, ACCT_SEG_REC_ACCT_JOURNAL_OWNER}(j)
                                Dim SEGX_CODE As String = New String() {JOURNAL_OWNER, ACCT_SEG_CODE}(j)
                                Dim rowGLTDETL1 As DataRow = dst.Tables("GLTDETL1").NewRow
                                With rowGLTDETL1
                                    .Item("OPS_YYYYPP") = OPS_YYYYPP
                                    .Item("JOURNAL_NO") = JOURNAL_NO
                                    JOURNAL_LNO = JOURNAL_LNO + 1
                                    .Item("JOURNAL_LNO") = JOURNAL_LNO
                                    .Item("DETL_CTL_DATE") = Format(DATETIME_STAMP, "MM/dd/yyyy")
                                    .Item("ACCT_CODE") = ACCT_CODE
                                    .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & ""
                                    .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & ""
                                    .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & ""
                                    .Item(COLUMN_NAME) = SEGX_CODE
                                    .Item("DETL_POSTING_AMT") = DETL_POSTING_AMT
                                    .Item("DETL_EXE_NO") = XNO
                                End With
                                dst.Tables("GLTDETL1").Rows.Add(rowGLTDETL1)
                                DETL_POSTING_AMT = -1 * DETL_POSTING_AMT
                            Next
                        End If
                    End If
                Next
            End If
        Next
    End Sub

    Sub Prepare_GL_Account_Activity_Recaps(ByVal TABLE_NAME As String)
        If dst.Tables.Contains("GLTINTF2") Then
            dst.Tables.Remove("GLTINTF2")
        End If

        ASCMAIN1.sql = "Select GLTDETL1.JOURNAL_NO, GLTDETL1.OPS_YYYYPP, GLTSEGM1.ACCT_SEG_ID, GLTSEGM1.ACCT_SEG_CODE, GLTDETL1.ACCT_CODE from GLTSEGM1,GLTDETL1 where ROWNUM <1"

        dst.Tables.Add(ASCDATA1.GetDataTable("**", "GLTINTF2", 5))
        With dst.Tables("GLTINTF2")
            .Columns.Add("ACCT_SEG_DR", GetType(System.Double))
            .Columns.Add("ACCT_SEG_CR", GetType(System.Double))
            .Columns.Add("ACCT_SEG_NET", GetType(System.Double))
            .Columns.Add("ACCT_SEG_TRANS", GetType(System.Int32))
        End With
        For i As Integer = 2 To 4
            Dim SEGX_CODE As String = "SEG" & CStr(i) & "_CODE"
            For Each row As DataRow In ASCMAIN1.Distinct_Values("",
            dst.Tables(TABLE_NAME), "JOURNAL_NO", "OPS_YYYYPP", SEGX_CODE, "ACCT_CODE").Rows
                Dim sql As String = "JOURNAL_NO = '" & row("JOURNAL_NO") & "' and OPS_YYYYPP = '" & row("OPS_YYYYPP") & "' and SEG" & CStr(i) & "_CODE = '" & row(SEGX_CODE) & "' and ISNULL(ACCT_CODE,'') = '" & row("ACCT_CODE") & "'"
                Dim ACCT_SEG_DR As Double = Val(dst.Tables(TABLE_NAME).Compute("SUM(DETL_POSTING_AMT)", sql & " AND DETL_POSTING_AMT >= 0") & "")
                Dim ACCT_SEG_CR As Double = Val(dst.Tables(TABLE_NAME).Compute("SUM(DETL_POSTING_AMT)", sql & " AND DETL_POSTING_AMT < 0") & "")
                Dim ACCT_SEG_NET As Double = Val(dst.Tables(TABLE_NAME).Compute("SUM(DETL_POSTING_AMT)", sql) & "")
                Dim ACCT_SEG_TRANS As Double = Val(dst.Tables(TABLE_NAME).Compute("COUNT(ACCT_CODE)", sql) & "")
                Dim rowGLTINTF2 As DataRow = dst.Tables("GLTINTF2").NewRow
                rowGLTINTF2("JOURNAL_NO") = row("JOURNAL_NO")
                rowGLTINTF2("OPS_YYYYPP") = row("OPS_YYYYPP")
                rowGLTINTF2("ACCT_SEG_ID") = CStr(i)
                rowGLTINTF2("ACCT_SEG_CODE") = row(SEGX_CODE)
                rowGLTINTF2("ACCT_CODE") = row("ACCT_CODE")
                If rowGLTINTF2("ACCT_CODE") & "" = "" Then
                    rowGLTINTF2("ACCT_CODE") = "?"
                End If
                rowGLTINTF2("ACCT_SEG_DR") = ACCT_SEG_DR
                rowGLTINTF2("ACCT_SEG_CR") = ACCT_SEG_CR
                rowGLTINTF2("ACCT_SEG_NET") = ACCT_SEG_NET
                rowGLTINTF2("ACCT_SEG_TRANS") = ACCT_SEG_TRANS
                dst.Tables("GLTINTF2").Rows.Add(rowGLTINTF2)
            Next
        Next

    End Sub

    Function GL_Prep(
    ByVal YYYY_beg As String,
    ByVal YYYY_end As String,
    Optional ByVal budget As Boolean = False,
    Optional ByVal OFFSET As Integer = 0,
    Optional ByVal OFFSET_Y As Integer = 0,
    Optional ByVal budTblsfx24 As String = "",
    Optional ByVal TABLE_NAME As String = "") As String

        ' determine YYYY_gyp as lesser of YYYY_beg and GYP
        ' get all years into work table from YYYY_gyp thru endyear
        ' change nulls to zeroes
        ' get balance forward set for years > YYYY_gyp thru YYYY_end
        ' close net profit into RTE for all years from GYP thru YYYY_end

        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim z As String
        Dim sqlbs As String
        Dim sqlis As String
        Dim sql As String

        Dim GYP As String
        Dim RTE As String
        Dim RTEsql As String
        Dim RTEsql_group_by As String

        GYP = ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP")
        RTE = ROWs("GLTPARM1").Item("GL_PARM_RET_EARN_ACCT")

        RTEsql = ""
        RTEsql_group_by = ""
        Dim seg(4) As String
        seg(2) = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
        seg(3) = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
        seg(4) = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")

        Dim YYYY_gyp As String
        If YYYY_beg < Mid$(GYP, 1, 4) Then
            YYYY_gyp = YYYY_beg
        Else
            YYYY_gyp = Mid$(GYP, 1, 4)
        End If

        Dim YRS As String
        YRS = ""
        For i = Val(YYYY_gyp) To Val(YYYY_end)
            YRS = YRS & ",'" & Format$(i, "0000") & "'"
        Next i
        YRS = Mid$(YRS, 2)

        '' Make sure that all Segment Codes are accounted for in Segment Master File

        'For i = 2 To 4
        '    z = Format$(i, "0")
        '    sql = "INSERT INTO GLTSEGM1 (ACCT_SEG_ID, ACCT_SEG_CODE, ACCT_SEG_DESC)"
        '    sql = sql & " SELECT '" & z & "', SEG" & z & "_CODE, 'Code ' || SEG" & z & "_CODE "
        '    sql = sql & " FROM"
        '    sql = sql & " (SELECT DISTINCT SEG" & z & "_CODE FROM GLTACCT3 "
        '    sql = sql & " MINUS "
        '    sql = sql & "  SELECT ACCT_SEG_CODE FROM GLTSEGM1 WHERE ACCT_SEG_ID = '" & z & "')"
        '    OraD.ExecuteSQL(sql)
        'Next i

        sql = "Select GLTACCT3.*, GLTACCT1.ACCT_TYPE "
        If budget Then
            If budTblsfx24 = "" Then
                budTblsfx24 = "2"
            End If
            sql = sql & " from GLTACCT" & budTblsfx24 & " GLTACCT3,GLTACCT1"
        Else
            sql = sql & " from GLTACCT3,GLTACCT1"
        End If
        sql = sql & " where GLTACCT1.ACCT_CODE (+) = GLTACCT3.ACCT_CODE"
        sql = sql & "   and GLTACCT3.ACCT_YEAR in (" & YRS & ")"
        Dim TT As String = ""
        If TABLE_NAME <> "" Then
            TT = TABLE_NAME
            ASCDATA1.ExecuteSQL("Delete from " & TT)
            ASCDATA1.ExecuteSQL("Insert into " & TT & " " & sql)
        Else
            TT = ASCMAIN1.Temp_Table(sql)
            ASCDATA1.ExecuteSQL("Alter Table " & TT & " add Primary Key (ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE,ACCT_YEAR)")
        End If

        For i = 2 To 4
            z = "SEG" & Format$(i, "0")
            If ROWs("GLTPARM1").Item("GL_PARM_" & z & "_RTE") & "" = "1" Then
                RTEsql = RTEsql & z & "_CODE,"
                RTEsql_group_by = RTEsql_group_by & z & "_CODE,"
            Else
                RTEsql = RTEsql & "'" & seg(i) & "' " & z & "_CODE,"
                'RTEsql_group_by = RTEsql_group_by & "'" & seg(i) & "',"
            End If
        Next i

        ASCDATA1.ExecuteSQL("Update " & TT & " set ACCT_BEG_BAL = 0 where ACCT_BEG_BAL is Null")

        sqlbs = "Select ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE, ACCT_TYPE, NVL(ACCT_BEG_BAL,0) "
        sqlis = "Select " & RTEsql & " Sum (NVL(ACCT_BEG_BAL,0) "
        For i = 1 To 13
            If budget Then
                z = "ACCT_BUD_P" & Format$(i, "00")
            Else
                z = "ACCT_ACT_P" & Format$(i, "00")
            End If
            ASCDATA1.ExecuteSQL("Update " & TT & " Set " & z & " = 0 where " & z & " is Null")
            sqlbs = sqlbs & " + NVL(" & z & ",0)"
            sqlis = sqlis & " + NVL(" & z & ",0)"
        Next i
        sqlis = sqlis & ")"
        sqlbs = sqlbs & " ACCT_BEG_BAL"

        If TABLE_NAME = "" Then
            Call Create_TDA(dst.Tables.Add, TT, "*")
            For j = 0 To 13
                dst.Tables(TT).Columns(5 + j).DefaultValue = 0
            Next
        Else
            '            Fill_Records(TT)
        End If

        Dim RTE_imax As Integer

        If Val(Mid$(GYP, 1, 4)) <= YYYY_end - 1 Then
            Dim yz As String
            For i = Val(Mid$(GYP, 1, 4)) To YYYY_end - 1
                yz = Format$(i + 1, "0000")
                sql = sqlbs & " from " & TT & " where ACCT_TYPE in ('A','L','E') and ACCT_YEAR = '" & Format$(i, "0000") & "'"
                For Each row As DataRow In ASCDATA1.GetDataTable(sql, "GLTACCTX").Rows
                    Dim ACCT_BEG_BAL As Double = Val(row.Item("ACCT_BEG_BAL") & "")
                    If ACCT_BEG_BAL <> 0 Then
                        Dim rowTT As DataRow = Fill_Record(TT, New String() {row.Item("ACCT_CODE"),
                        row.Item("SEG2_CODE"), row.Item("SEG3_CODE"), row.Item("SEG4_CODE"), yz}, True)
                        rowTT.Item("ACCT_TYPE") = row.Item("ACCT_TYPE")
                        rowTT.Item("ACCT_BEG_BAL") = Val(rowTT.Item("ACCT_BEG_BAL") & "") + ACCT_BEG_BAL
                        TDAs(TT).Update(dst.Tables(TT))
                    End If
                Next
                RTE_imax = i
                Call RTE_Calc(i, YYYY_gyp, TT, RTEsql_group_by, sqlis, RTE_imax, RTE)
            Next i
        End If

        If OFFSET <> 0 Then
            Stop ' WHEN WE HAVE A FRESH MIND
            '    Dim jmax As Integer
            '    j = 0
            '    If budget Then
            '        z = "BUD"
            '    Else
            '        z = "ACT"
            '    End If
            '    sql = "Select ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE, ACCT_TYPE" & vbCr
            '    sql = sql & ", SUM (DECODE(ACCT_YEAR,'" & Format$(YYYY_gyp, "0000") & "',NVL(ACCT_BEG_BAL,0))) P000" & vbCr
            '    For i = Val(YYYY_gyp) To Val(YYYY_end)
            '        For k = 1 To 12
            '            j = j + 1
            '            sql = sql & ", SUM (DECODE(ACCT_YEAR,'" & Format$(i, "0000") & "', NVL(ACCT_" & z & "_P" & Format$(k, "00") & ",0))) P" & Format$(j, "000") & vbCr
            '        Next k
            '    Next i
            '    sql = sql & " from " & TT & " group by "
            '    sql = sql & "ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE, ACCT_TYPE" & vbCr
            '    jmax = j
            '    Dim dyn As OraDynaset
            '    dyn = OraD.CreateDynaset(sql, 8&)
            '    ASCDATA1.ExecuteSQL("Delete from " & TT)
            '    sql = "Select * from " & TT & " where ROWNUM < 1"
            '    Dim dyntt As OraDynaset
            '    dyntt = OraD.CreateDynaset(sql, 0&)
            '    Dim a As Double
            '    Dim AMT() As Double
            '    Do While Not dyn.EOF
            '        ReDim AMT(12)
            '        k = 0
            '        i = OFFSET_Y
            '        For j = 0 To jmax
            '            a = Val(dyn.Fields("P" & Format$(j, "000")).Value & "")
            '            If j <= OFFSET Then
            '                AMT(0) = AMT(0) + a
            '            Else
            '                k = k + 1
            '                AMT(k) = a
            '                If k = 12 Or j = jmax Then
            '                    If InStr("ALE", dyn.Fields("ACCT_TYPE").Value & "") = 0 Then
            '                        If i = OFFSET_Y Then
            '                            dyntt.AddNew()
            '                            dyntt.Fields("ACCT_CODE").Value = dyn.Fields("ACCT_CODE").Value
            '                            dyntt.Fields("SEG2_CODE").Value = dyn.Fields("SEG2_CODE").Value
            '                            dyntt.Fields("SEG3_CODE").Value = dyn.Fields("SEG3_CODE").Value
            '                            dyntt.Fields("SEG4_CODE").Value = dyn.Fields("SEG4_CODE").Value
            '                            dyntt.Fields("ACCT_YEAR").Value = "0000" ' Val(y0) ' + i - 1
            '                            dyntt.Fields("ACCT_TYPE").Value = dyn.Fields("ACCT_TYPE").Value & ""
            '                            dyntt.Fields("ACCT_BEG_BAL").Value = AMT(0)
            '                            dyntt.Update()
            '                        End If
            '                        AMT(0) = 0
            '                    End If
            '                    If Val(YYYY_gyp) + i <= Val(YYYY_end) Then ' And (amt(0) <> 0 Or amt(1) <> 0 Or amt(2) <> 0 Or amt(3) <> 0 Or amt(4) <> 0 Or amt(5) <> 0 Or amt(6) <> 0 Or amt(7) <> 0 Or amt(8) <> 0 Or amt(9) <> 0 Or amt(10) <> 0 Or amt(11) <> 0 Or amt(12) <> 0) Then
            '                        dyntt.AddNew()
            '                        dyntt.Fields("ACCT_CODE").Value = dyn.Fields("ACCT_CODE").Value
            '                        dyntt.Fields("SEG2_CODE").Value = dyn.Fields("SEG2_CODE").Value
            '                        dyntt.Fields("SEG3_CODE").Value = dyn.Fields("SEG3_CODE").Value
            '                        dyntt.Fields("SEG4_CODE").Value = dyn.Fields("SEG4_CODE").Value
            '                        dyntt.Fields("ACCT_YEAR").Value = Val(y0) + i
            '                        dyntt.Fields("ACCT_TYPE").Value = dyn.Fields("ACCT_TYPE").Value & ""
            '                        dyntt.Fields("ACCT_BEG_BAL").Value = AMT(0)
            '                        For k = 1 To 12
            '                            dyntt.Fields("ACCT_" & z & "_P" & Format$(k, "00")).Value = AMT(k)
            '                            If InStr("ALE", dyn.Fields("ACCT_TYPE").Value & "") <> 0 Then
            '                                AMT(0) = AMT(0) + AMT(k)
            '                            End If
            '                            AMT(k) = 0
            '                        Next k
            '                        dyntt.Update()
            '                    End If
            '                    i = i + 1
            '                    k = 0
            '                End If
            '            End If
            '        Next j
            '        dyn.MoveNext()
            '    Loop

            '    i = 0
            '    RTE_imax = Val(YYYY_end)
            'GoSub Calc_RTE

            '    For i = Val(YYYY_gyp) To Val(YYYY_end)
            '        RTE_imax = Val(YYYY_end)
            '    GoSub Calc_RTE
            '    Next i

            '    sql = "Update " & TT & " SET ACCT_BEG_BAL = 0 "
            '    sql = sql & " where ACCT_TYPE in ('I','X') "
            '    sql = sql & " and ACCT_YEAR = '0000'"
            '    ASCDATA1.ExecuteSQL(sql) ' Clear out Accum R/E from periods prior to start of re-calendarized year which was stuffed into Op Accts

        End If

        ASCDATA1.ExecuteSQL("Delete from " & TT & " where ACCT_YEAR < '" & Format$(Val(YYYY_beg) + OFFSET_Y * Sign(Abs(OFFSET)), "0000") & "'")
        ASCDATA1.ExecuteSQL("Delete from " & TT & " where ACCT_YEAR > '" & Format$(Val(YYYY_end), "0000") & "'")

        If budget Then
            z = "BUD"
        Else
            z = "ACT"
        End If
        sql = "Delete from " & TT
        sql = sql & " where NVL(ACCT_BEG_BAL,0) = 0" & vbCr
        For k = 1 To 12
            sql = sql & " and NVL(ACCT_" & z & "_P" & Format$(k, "00") & ",0) = 0" & vbCr
        Next k
        'OraD.ExecuteSQL sql ' this throws off the TBAL where an account may have had activity which nets to 0

        If TABLE_NAME = "" Then
            ASCDATA1.ExecuteSQL("Create Index I_" & TT & "_1 on " & TT & " (ACCT_YEAR,ACCT_TYPE)")
        End If

        Call ASCMAIN1.AnalyzeTable(TT)

        Return TT

    End Function

    Sub RTE_Calc(
    ByVal YYYY As Integer,
    ByVal YYYY_gyp As String,
    ByVal TT As String,
    ByVal RTEsql_group_by As String,
    ByVal sqlis As String,
    ByVal RTE_imax As Integer,
    ByVal RTE As String)

        Dim RTE_imin As Integer
        If YYYY = 0 Then
            RTE_imin = YYYY_gyp
        Else
            RTE_imin = YYYY
        End If
        Dim sql As String
        sql = sqlis & " from " & TT & " where ACCT_TYPE in ('I','X') "
        sql = sql & " and ACCT_YEAR = '" & Format$(YYYY, "0000") & "'"
        If RTEsql_group_by.Length <> 0 Then sql = sql & " group by " & Mid$(RTEsql_group_by, 1, Len(RTEsql_group_by) - 1)
        For Each rowRTE As DataRow In ASCDATA1.GetDataTable(sql, "").Rows
            Dim ACCT_BEG_BAL As Double = Val(rowRTE.Item(3) & "")
            If ACCT_BEG_BAL <> 0 Then
                For RTE_i As Integer = RTE_imin To RTE_imax
                    Dim row As DataRow = Fill_Record(TT, New String() {RTE,
                    rowRTE.Item("SEG2_CODE"), rowRTE.Item("SEG3_CODE"), rowRTE.Item("SEG4_CODE"),
                    Format$(RTE_i + 1, "0000")}, True)
                    row.Item("ACCT_TYPE") = "E"
                    row.Item("ACCT_BEG_BAL") = Val(row.Item("ACCT_BEG_BAL") & "") + ACCT_BEG_BAL
                    TDAs(TT).Update(dst.Tables(TT))
                Next RTE_i
            End If
        Next
    End Sub

    Sub Breakout_By()
        For i As Integer = 2 To 4
            Dim z As String = "SEG" & CStr(i)
            If ROWs("GLTPARM1").Item("GL_PARM_" & z & "_DESC") & "" = "" Then
                Absx1.CtlFor(z & "_CODE").Visible = False
                If Absx1.chkFor(z & "_CODE").Checked Then
                    Absx1.chkFor(z & "_CODE").Checked = False
                End If
            Else
                Absx1.CtlFor(z & "_CODE").Text = ROWs("GLTPARM1").Item("GL_PARM_" & z & "_DESC") & ""
            End If
        Next
    End Sub

    Sub Breakout_By_Class()
        For i As Integer = 2 To 4
            Dim z As String = "SEG" & CStr(i) & "_CLASS"
            If ROWs("GLTPARM1").Item("GL_PARM_" & z & "_DESC") & "" = "" Then
                Absx1.CtlFor(z & "_CODE").Visible = False
                If Absx1.chkFor(z & "_CODE").Checked Then
                    Absx1.chkFor(z & "_CODE").Checked = False
                End If
            Else
                Absx1.CtlFor(z & "_CODE").Text = ROWs("GLTPARM1").Item("GL_PARM_" & z & "_DESC") & ""
            End If
        Next
    End Sub
#End Region

    Public Overridable Sub cbe_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Stop
    End Sub

    Public Overridable Sub cbe_BeforeDropDown(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        If MENU_ITEM_TYPE = "T" And Not ScreenMode Then
            e.Cancel = True
        End If
    End Sub

    Public Overridable Sub cbe_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Delete Then
            Dim cbectl As UltraWinEditors.UltraComboEditor = DirectCast(sender, UltraWinEditors.UltraComboEditor)
            If cbectl.Nullable Then
                cbectl.Value = DBNull.Value
            End If
        End If

        If e.KeyCode = Windows.Forms.Keys.Enter Then
            SendKeys.Send(Chr(9))
        End If
    End Sub

    Sub Write_Audit_Trail(ByVal row As DataRow, Optional ByVal FM_MODE As String = "")
        Dim row_original As DataRow = row.Table.NewRow
        If row.RowState = DataRowState.Added Then
            ' NO ORIGINAL VALUES TO BE HAD HERE
        Else
            For i As Int16 = 0 To row.Table.Columns.Count - 1
                row_original.Item(i) = row.Item(i, DataRowVersion.Original)
            Next
        End If
        Write_Audit_Trail(row, row_original, FM_MODE)
    End Sub

    Sub Write_Audit_Trail(
    ByRef row_current As DataRow,
    ByRef row_original As DataRow,
    Optional ByVal FM_MODE As String = "")

        If FM_MODE = "" Then
            FM_MODE = IIf(EntryMode = "New", "N", "E")
        End If

        Dim tbl As DataTable = row_current.Table
        Dim TABLE_NAME As String = tbl.TableName()

        If row_original Is Nothing Then
            If FM_MODE <> "N" Then
                Dim row_index As Integer = tbl.Rows.IndexOf(row_current)
                tbl.Select("", "", DataViewRowState.ModifiedOriginal)

                'Dim dvw As DataView = New DataView(tbl, "", "", DataViewRowState.OriginalRows)
                'row_original = dvw.Item(row_index).Row
                'Dim drv As DataRowView = dvw.Item(row_index)
            End If
        End If

        Dim tblASTAUDT1 As New DataTable
        With ASCDATA1.GetDataAdapter(tblASTAUDT1, "ASTAUDT1", "*", True, -1, False, 0)

            Dim KEY_VALUE As String = ""
            For i As Integer = 0 To tbl.PrimaryKey.Length - 1
                If FM_MODE = "D" Then
                    KEY_VALUE &= ":" & row_original.Item(i)
                Else
                    ' important for logging deleted rows from grids shown in FM screens
                    If row_current.RowState = DataRowState.Deleted Then
                        KEY_VALUE &= ":" & row_current.Item(i, DataRowVersion.Original)
                    Else
                        KEY_VALUE &= ":" & row_current.Item(i)
                    End If
                    'KEY_VALUE &= ":" & row_current.Item(i)
                End If
            Next
            KEY_VALUE = Mid$(KEY_VALUE, 2)

            For i As Integer = 0 To row_current.Table.Columns.Count - 1 ' tblASFBASE1.Columns.Count - 1
                Dim audit As Boolean = False
                Dim OLD_VALUE As Object = Nothing
                If FM_MODE = "N" Then
                    audit = (row_current.Item(i) & "" <> "")
                    OLD_VALUE = System.DBNull.Value
                ElseIf FM_MODE = "E" Then
                    If row_original Is Nothing Then
                        audit = (row_current.Item(i, DataRowVersion.Original) & "" <> row_current.Item(i) & "")
                        OLD_VALUE = row_current.Item(i, DataRowVersion.Original) & ""
                    Else
                        If row_current.RowState = DataRowState.Deleted Then
                            OLD_VALUE = row_original.Item(i) & ""
                            If OLD_VALUE <> "" Then
                                audit = True
                            End If
                        Else
                            Dim dt As String = row_current.Table.Columns(i).DataType.ToString
                            If dt = "System.Decimal" Or dt = "System.Integer" Or dt = "System.Int32" Or dt = "System.Int64" Then
                                audit = Val(row_original.Item(i) & "") <> Val(row_current.Item(i) & "")
                                OLD_VALUE = row_original.Item(i) & ""
                            Else
                                audit = (row_original.Item(i) & "" <> row_current.Item(i) & "")
                                OLD_VALUE = row_original.Item(i) & ""
                            End If

                        End If
                    End If
                ElseIf FM_MODE = "D" Then
                    OLD_VALUE = row_original.Item(i) & ""
                    If OLD_VALUE <> "" Then
                        audit = True
                    End If
                End If
                If audit Then
                    Dim rowASTAUDT1 As DataRow = tblASTAUDT1.NewRow
                    rowASTAUDT1.Item("TABLE_NAME") = TABLE_NAME
                    rowASTAUDT1.Item("KEY_VALUE") = KEY_VALUE
                    rowASTAUDT1.Item("COLUMN_NAME") = row_current.Table.Columns(i).ColumnName ' tblASFBASE1.Columns(i).ColumnName
                    rowASTAUDT1.Item("USER_ID") = ASCMAIN1.USER_ID
                    rowASTAUDT1.Item("INIT_DATE") = DATETIME_STAMP
                    If Len(OLD_VALUE & "") > 255 Then
                        OLD_VALUE = Mid(OLD_VALUE, 1, 255)
                    End If
                    rowASTAUDT1.Item("OLD_VALUE") = OLD_VALUE
                    If FM_MODE <> "D" And row_current.RowState <> DataRowState.Deleted Then
                        Dim NEW_VALUE As String = row_current.Item(i) & ""
                        If Len(NEW_VALUE & "") > 255 Then
                            NEW_VALUE = Mid(NEW_VALUE, 1, 255)
                        End If
                        rowASTAUDT1.Item("NEW_VALUE") = NEW_VALUE
                    End If
                    rowASTAUDT1.Item("FM_MODE") = FM_MODE
                    rowASTAUDT1.Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                    rowASTAUDT1.Item("SELECTION_NO") = SELECTION_NO
                    rowASTAUDT1.Item("XNO") = XNO
                    tblASTAUDT1.Rows.Add(rowASTAUDT1)
                End If
            Next
            .Update(tblASTAUDT1)
            .Dispose()
        End With

    End Sub

    Overloads Sub Write_Event_Log(
    ByRef TABLE_NAME As String,
    ByRef TABLE_KEY As String,
    ByRef EVENT_DESC As String)

        Me.Write_Event_Log(TABLE_NAME, TABLE_KEY, EVENT_DESC, String.Empty)

    End Sub

    Overloads Sub Write_Event_Log(ByRef TABLE_NAME As String,
    ByRef TABLE_KEY As String,
    ByRef EVENT_DESC As String,
    ByRef EVENT_TYPE As String)

        Dim tblTATEVNT1 As New DataTable
        With ASCDATA1.GetDataAdapter(tblTATEVNT1, "TATEVNT1", "*", True, -1, False, 0)

            Dim rowTATEVNT1 As DataRow = tblTATEVNT1.NewRow
            With rowTATEVNT1
                .Item("TABLE_NAME") = TABLE_NAME
                .Item("TABLE_KEY") = TABLE_KEY
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("EVENT_TYPE") = EVENT_TYPE
                .Item("EVENT_DESC") = EVENT_DESC
                .Item("EVENT_KEY") = ""
                .Item("FORM_NAME") = Me.Name
                .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                .Item("SELECTION_NO") = SELECTION_NO
                .Item("XNO") = XNO

            End With

            tblTATEVNT1.Rows.Add(rowTATEVNT1)

            .Update(tblTATEVNT1)
            .Dispose()
        End With
    End Sub

    Sub Write_Event_Log_Batch(
    ByRef TABLE_NAME As String,
    ByRef sql As String)

        Dim tblTATEVNT1 As New DataTable
        With ASCDATA1.GetDataAdapter(tblTATEVNT1, "TATEVNT1", "*", True, -1, False, 0)

            For Each row As DataRow In ASCDATA1.GetDataTable(sql).Rows
                Dim rowTATEVNT1 As DataRow = tblTATEVNT1.NewRow
                With rowTATEVNT1
                    .Item("TABLE_NAME") = TABLE_NAME
                    .Item("TABLE_KEY") = row(0)
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("EVENT_TYPE") = ""
                    .Item("EVENT_DESC") = row(1)
                    .Item("EVENT_KEY") = ""
                    .Item("FORM_NAME") = Me.Name
                    .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                    .Item("SELECTION_NO") = SELECTION_NO
                    .Item("XNO") = XNO

                End With

                tblTATEVNT1.Rows.Add(rowTATEVNT1)
            Next

            .Update(tblTATEVNT1)
            .Dispose()
        End With
    End Sub

    Sub Set_grd1stRow(
    ByRef grd As UltraWinGrid.UltraGrid,
    Optional ByVal grd2 As UltraWinGrid.UltraGrid = Nothing)
        If grd.Rows.Count > 0 Then
            grd.ActiveRow = grd.Rows(0)
            If grd2 IsNot Nothing Then
                grd2.Visible = True
            End If
        Else
            If grd2 IsNot Nothing Then
                grd2.Visible = False
            End If
        End If
    End Sub

    Sub Show_Filter(
    ByRef grd As UltraWinGrid.UltraGrid,
    Optional ByVal show_filter As Boolean = True,
    Optional ByVal enhanced As Boolean = False)
        With grd.DisplayLayout.Override
            If enhanced Then
                .FilterUIProvider = Me.enhancedGrdFilter
            End If
            .FilterUIType = UltraWinGrid.FilterUIType.FilterRow
            .FilterRowAppearance.BackColor = System.Drawing.Color.AliceBlue
            .FilterClearButtonLocation = UltraWinGrid.FilterClearButtonLocation.Row
            .FilterOperatorLocation = UltraWinGrid.FilterOperatorLocation.AboveOperand

            'Assign the FilterUIProvider to use the new Excel style filtering from UltraGridFilterUIProvider
            '.FilterUIProvider = New SupportDialogs.FilterUIProvider.UltraGridFilterUIProvider '  Me.UltraGridFilterUIProvider1

            If show_filter Then
                .AllowRowFiltering = DefaultableBoolean.True
                .FilterUIType = UltraWinGrid.FilterUIType.FilterRow
            Else
                .AllowRowFiltering = DefaultableBoolean.False
            End If
        End With
    End Sub

    Sub Clear_All_Filters(ByRef grd As UltraWinGrid.UltraGrid)
        For Each band As Infragistics.Win.UltraWinGrid.UltraGridBand In grd.DisplayLayout.Bands
            band.ColumnFilters.ClearAllFilters()
        Next
    End Sub

    Sub Sort_cmbColumns(
    ByVal cmb As UltraWinGrid.UltraCombo,
    Optional ByVal COLUMN_NAMEs As String = "")

        Dim BAND As Int32 = 0

        If COLUMN_NAMEs = "" Then
            For Each c As UltraWinGrid.UltraGridColumn _
            In cmb.DisplayLayout.Bands(BAND).Columns
                If Not c.Hidden Then
                    COLUMN_NAMEs = c.Key
                    Exit For
                End If
            Next
        End If

        With cmb
            cmb.PerformAction(UltraWinGrid.UltraComboAction.Dropdown)

            If COLUMN_NAMEs <> "" Then
                If .DisplayLayout Is Nothing Then
                    Exit Sub
                Else
                    .DisplayLayout.Bands(BAND).SortedColumns.Clear()
                    For Each COLUMN_NAME As String In Split(COLUMN_NAMEs, ",")
                        COLUMN_NAME = Trim(COLUMN_NAME)
                        .DisplayLayout.Bands(BAND).SortedColumns.Add(COLUMN_NAME, (COLUMN_NAME = COLUMN_NAME.ToLower))
                    Next
                    If .Rows.Count <> 0 Then .ActiveRow = .Rows(0) ' .Rows.GetRowAtVisibleIndex(0)
                End If
            End If
        End With
    End Sub

    Sub Sort_grdColumns(
    ByRef grd As UltraWinGrid.UltraGrid,
    Optional ByVal COLUMN_NAMEs As String = "",
    Optional ByVal LockSort As Boolean = False,
    Optional ByVal BAND As Int32 = 0,
    Optional ByVal set_active_row_to_top_row As Boolean = True)
        If COLUMN_NAMEs = "" Then
            For Each c As UltraWinGrid.UltraGridColumn _
            In grd.DisplayLayout.Bands(BAND).Columns
                If Not c.Hidden Then
                    COLUMN_NAMEs = c.Key
                    Exit For
                End If
            Next
        End If

        With grd
            .Selected.Rows.Clear()
            If COLUMN_NAMEs <> "" Then
                If .DisplayLayout Is Nothing Then
                    Exit Sub
                Else

                    .DisplayLayout.Bands(BAND).SortedColumns.Clear()
                    For Each COLUMN_NAME As String In Split(COLUMN_NAMEs, ",")
                        COLUMN_NAME = Trim(COLUMN_NAME)
                        If .DisplayLayout.Bands(BAND).Columns.Exists(COLUMN_NAME) Then
                            .DisplayLayout.Bands(BAND).SortedColumns.Add(COLUMN_NAME, (COLUMN_NAME = COLUMN_NAME.ToLower))
                        Else
                            MsgBox(COLUMN_NAME & " does not exist in grid " & grd.Name,
                                   MsgBoxStyle.OkOnly, "Please call ABS")
                            ' to trap CP abnd GF errors in SOI key does not exist
                        End If
                    Next
                    If .Rows.Count <> 0 Then .ActiveRow = .Rows(0) ' .Rows.GetRowAtVisibleIndex(0)
                    'Debug.Print(.Rows(0).Cells(0).Text & ":" & .Rows(0).Cells(1).Text)
                    'If ROWs.Count <> 0 Then .ActiveRow = .Rows.GetRowAtVisibleIndex(0)
                    'Stop
                    ' this is not working - need to move to top of grid

                End If
            End If
        End With

        If LockSort Then
            grd.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.Select
            If BAND = 0 Then
                grd.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.Select
            Else
                grd.DisplayLayout.Bands(BAND).Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.Select
            End If
        Else
            If grd.DisplayLayout.Override.HeaderClickAction <> UltraWinGrid.HeaderClickAction.SortMulti And
            grd.DisplayLayout.Override.HeaderClickAction <> UltraWinGrid.HeaderClickAction.Default Then
                If BAND = 0 Then
                    grd.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.SortMulti
                Else
                    grd.DisplayLayout.Bands(BAND).Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.SortMulti
                End If
            Else
                If grd.DisplayLayout.Bands(BAND).Override.HeaderClickAction <> UltraWinGrid.HeaderClickAction.SortMulti Then
                    grd.DisplayLayout.Bands(BAND).Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.SortMulti
                End If
            End If
        End If

        'If grd.Selected.Rows.Count <> 0 Then
        grd.Selected.Rows.Clear()
        'End If
    End Sub

    Public Overridable Sub cal_ValidationError(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinSchedule.DateValidationErrorEventArgs)
        If e.ErrorCode = UltraWinSchedule.DateValidationError.UnableToParseValue Then
            If (Len(e.ErrorValue & "") = 6 Or Len(e.ErrorValue & "") = 8) And Not e.ErrorValue.ToString.Contains("/") Then
                e.NewValue = Mid(e.ErrorValue, 1, 2) & "/" & Mid(e.ErrorValue, 3, 2) & "/" & Mid(e.ErrorValue, 5, 4)
            End If
        End If
    End Sub

    Private Sub cal_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Return Then
            Dim ctl As UltraWinSchedule.UltraCalendarCombo = DirectCast(sender, UltraWinSchedule.UltraCalendarCombo)
            If (Len(ctl.Text & "") = 6 Or Len(ctl.Text & "") = 8) And Not ctl.Text.ToString.Contains("/") Then
                ctl.Value = CDate(Mid(ctl.Text, 1, 2) & "/" & Mid(ctl.Text, 3, 2) & "/" & Mid(ctl.Text, 5, 4))
            End If
        End If
    End Sub

    Public Overridable Sub dte_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            SendKeys.Send(Chr(9))
        End If
    End Sub

    Public Overridable Sub dte_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub med_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            SendKeys.Send(Chr(9))
        End If
    End Sub

    Private Sub opt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            SendKeys.Send(Chr(9))
        End If
    End Sub

    Sub Summary_Table(
    ByVal TABLE_NAME As String,
    ByVal TABLE_NAME_child As String,
    ByVal GROUP_BY As String,
    ByVal AGGREGATES As String)

        Dim group_by_COLUMN_NAMEs As String() = Split(GROUP_BY, ",")
        Dim group_by_COLs As Integer = UBound(group_by_COLUMN_NAMEs)

        dst.Tables.Add(
        ASCMAIN1.Distinct_Values(TABLE_NAME,
        dst.Tables(TABLE_NAME_child),
        group_by_COLUMN_NAMEs))


        Dim DC1() As DataColumn
        Dim DC2() As DataColumn
        ReDim DC1(group_by_COLs)
        ReDim DC2(group_by_COLs)

        For i As Integer = 0 To group_by_COLs
            DC1(i) = dst.Tables(TABLE_NAME).Columns(group_by_COLUMN_NAMEs(i))
            DC2(i) = dst.Tables(TABLE_NAME_child).Columns(group_by_COLUMN_NAMEs(i))
        Next
        dst.Relations.Add(TABLE_NAME, DC1, DC2)

        Dim aggregate_COLUMN_NAMEs As String() = Split(AGGREGATES, ",")
        Dim aggregate_COLs As Integer = UBound(aggregate_COLUMN_NAMEs)

        For i As Integer = 0 To aggregate_COLs
            dst.Tables(TABLE_NAME).Columns.Add(aggregate_COLUMN_NAMEs(i),
            dst.Tables(TABLE_NAME_child).Columns(aggregate_COLUMN_NAMEs(i)).DataType,
            "SUM(CHILD." & aggregate_COLUMN_NAMEs(i) & ")")
        Next

    End Sub

    Sub Range_Events(ByVal c As Control)
        For Each cc As Control In c.Controls
            If TypeOf cc Is UltraWinEditors.UltraCheckEditor Then
                Dim chkctl As UltraWinEditors.UltraCheckEditor = DirectCast(cc, UltraWinEditors.UltraCheckEditor)
                AddHandler chkctl.CheckedChanged, AddressOf Range_chk_CheckedChanged
                Call Range_dte(cc)
            End If
        Next
    End Sub

    Private Sub Range_chk_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call Range_dte(sender)
    End Sub

    Sub Range_dte(ByVal sender As Control)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        COLUMN_NAME = Mid(COLUMN_NAME, 4)
        If DirectCast(sender, UltraWinEditors.UltraCheckEditor).Checked Then
            Absx1.dteFor(COLUMN_NAME).Enabled = False
            Absx1.dteFor(COLUMN_NAME).Value = Null
        Else
            Absx1.dteFor(COLUMN_NAME).Enabled = True
            Absx1.dteFor(COLUMN_NAME).Value = Format(Now, "MM/dd/yyyy")
        End If

    End Sub

    Sub Create_Update_Command(ByVal COMMAND_NAME As String, ByVal PARAMETERS As String)
        Dim cmd As OracleCommand = ASCMAIN1.oraCon.CreateCommand
        cmd.CommandText = ASCMAIN1.sql
        If PARAMETERS <> "" Then
            Call ASCDATA1.Create_Parameters(cmd, PARAMETERS)
        End If
        Update_CMDs.Add(COMMAND_NAME, cmd)
    End Sub

    Function Update_Command(
    ByVal COMMAND_NAME As String,
    ByVal ParamArray PARAMETERS() As String)

        If PARAMETERS.Length <> 0 Then
            For i As Integer = 0 To Update_CMDs(COMMAND_NAME).Parameters.Count - 1
                Dim z As String = "PARM" & CStr(i + 1)
                Update_CMDs(COMMAND_NAME).Parameters(z).Value = PARAMETERS(i)
            Next
        End If

        Return Update_CMDs(COMMAND_NAME).ExecuteNonQuery
    End Function

    Overridable Function Generate_Report(
    ByVal RPT As String,
    Optional ByVal RPT_TITLE As String = "",
    Optional ByVal SUBT As String = "",
    Optional ByVal RecordSelectionFormula As String = "",
    Optional ByVal ExportFormat As String = "",
    Optional ByVal TempExportFilenameBody As String = "",
    Optional ByVal archive_this_report As Boolean = True)

        Call ASCMAIN1.Progress("Now Printing " & IIf(RPT_TITLE <> "", RPT_TITLE, Me.Text))
        Return clsASCBASE1.Generate_Report(RPT, RPT_TITLE, SUBT,
            False, False,
            RecordSelectionFormula,
            ExportFormat, TempExportFilenameBody, archive_this_report)

    End Function

    Sub Print_Report_Begin()
        Call ASCMAIN1.Progress("Now Printing Reports")
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()

        clsASCBASE1.Print_Report_Begin()

        F = clsASCBASE1.F
    End Sub

    Sub Print_Report_to_Printer(
    Optional ByVal PRINTER_NAME As String = "",
    Optional ByVal RecordSelectionFormula As String = "")

        If RecordSelectionFormula <> "" Then
            If ASCMAIN1.CR_RPT.RecordSelectionFormula = "" Then
                ASCMAIN1.CR_RPT.RecordSelectionFormula = RecordSelectionFormula
            Else
                ASCMAIN1.CR_RPT.RecordSelectionFormula.Replace(ASCMAIN1.CR_RPT.RecordSelectionFormula, RecordSelectionFormula)
            End If
        End If
        If PRINTER_NAME <> "" Then
            ASCMAIN1.CR_RPT.PrintOptions.PrinterName = PRINTER_NAME
        End If
        ASCMAIN1.CR_RPT.PrintToPrinter(1, False, 0, 0)
    End Sub

    Sub Print_Report_End(
    Optional ByVal print_without_showing As Boolean = False,
    Optional ByVal close_report_viewer As Boolean = False,
    Optional ByVal PrinterName As String = "",
    Optional ByVal number_of_copies As Int32 = 1,
    Optional ByVal streamIPandPort As String = "")

        clsASCBASE1.Print_Report_End(
        print_without_showing,
        close_report_viewer,
        PrinterName,
        number_of_copies,
        streamIPandPort)

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
    End Sub

    Sub Create_ASTAUDTE()
        With dst
            .Tables.Add("ASTAUDTE") ' Events
            .Tables("ASTAUDTE").Columns.Add("INIT_DATE", GetType(System.DateTime))
            .Tables("ASTAUDTE").Columns.Add("INIT_OPER", GetType(System.String))
            .Tables("ASTAUDTE").Columns.Add("EVENT_DESC", GetType(System.String))
        End With
    End Sub

    Sub INIT_LAST(ByVal TABLE_NAME As String,
    Optional ByVal UseRowState As Boolean = False,
    Optional ByVal sqlWhere As String = "",
    Optional ByVal LASTonINIT As Boolean = False)
        ' THIS ROUTINE WILL EVENTUALLY INCLUDE
        ' 3) COMPLETE AUDIT TRAIL OF COLUMNS BY COMPARING TO ORIGINAL VALUES

        For Each row As DataRow In dst.Tables(TABLE_NAME).Select(sqlWhere, "", DataViewRowState.CurrentRows)
            'FM USES New and Edit, and not N and E - probably should be normalized
            If (Not UseRowState And (EntryMode = "N" Or EntryMode = "New")) Or (UseRowState And row.RowState = DataRowState.Added) Then
                row("INIT_DATE") = DATETIME_STAMP
                row("INIT_OPER") = ASCMAIN1.USER_ID
                If LASTonINIT Then
                    row("LAST_DATE") = DATETIME_STAMP
                    row("LAST_OPER") = ASCMAIN1.USER_ID
                End If
            ElseIf (Not UseRowState And (EntryMode = "E" Or EntryMode = "Edit")) Or (UseRowState And row.RowState = DataRowState.Modified) Then
                If dst.Tables(TABLE_NAME).Columns.Contains("LAST_DATE") Then
                    row("LAST_DATE") = DATETIME_STAMP
                    row("LAST_OPER") = ASCMAIN1.USER_ID
                End If
            End If
        Next

    End Sub

    Sub WriteAuditTrail(ByVal TABLE_NAME As String)

        Dim UPDATE_COLUMNS As New List(Of String)
        For Each OP As OracleParameter In TDAs(TABLE_NAME).UpdateCommand.Parameters
            UPDATE_COLUMNS.Add(OP.SourceColumn)
        Next

        Dim NED As String = "*"
        If AUDIT.ContainsKey(TABLE_NAME) Then
            NED = AUDIT(TABLE_NAME)
        End If
        Dim audit_N As Boolean = (NED = "*" Or InStr(NED, "N") <> 0)
        Dim audit_E As Boolean = (NED = "*" Or InStr(NED, "E") <> 0)
        Dim audit_D As Boolean = (NED = "*" Or InStr(NED, "D") <> 0)

        Dim FM_MODE As String = ""
        For Each row As DataRow In dst.Tables(TABLE_NAME).Rows
            Dim KEY_VALUE As String = ""
            For i As Integer = 0 To dst.Tables(TABLE_NAME).PrimaryKey.Length - 1
                If row.RowState = DataRowState.Deleted Then
                    KEY_VALUE &= ":" & row.Item(i, DataRowVersion.Original)
                Else
                    KEY_VALUE &= ":" & row.Item(i)
                End If
            Next
            KEY_VALUE = Mid(KEY_VALUE, 2)

            If row.RowState = DataRowState.Unchanged Then
            Else
                For Each dc As DataColumn In dst.Tables(TABLE_NAME).Columns
                    If UPDATE_COLUMNS.Contains(dc.ColumnName) Then
                        FM_MODE = ""
                        Dim rowASTAUDT1 As DataRow = dst.Tables("ASTAUDT1").NewRow

                        Dim COLUMN_NAME As String = dc.ColumnName
                        If row.RowState = DataRowState.Deleted And audit_D Then
                            FM_MODE = "D"
                            rowASTAUDT1.Item("OLD_VALUE") = row.Item(COLUMN_NAME, DataRowVersion.Original)
                        Else
                            If row.RowState = DataRowState.Added Then
                                If row.Item(COLUMN_NAME, DataRowVersion.Current) & "" = "" _
                                Or Not audit_N Then
                                Else
                                    FM_MODE = "N"
                                    rowASTAUDT1.Item("NEW_VALUE") = row.Item(COLUMN_NAME, DataRowVersion.Current)
                                End If
                            ElseIf row.RowState = DataRowState.Modified Then
                                If row.Item(COLUMN_NAME, DataRowVersion.Current).Equals(row.Item(COLUMN_NAME, DataRowVersion.Original)) _
                                Or Not audit_E Then
                                Else
                                    FM_MODE = "E"
                                    rowASTAUDT1.Item("OLD_VALUE") = row.Item(COLUMN_NAME, DataRowVersion.Original)
                                    rowASTAUDT1.Item("NEW_VALUE") = row.Item(COLUMN_NAME, DataRowVersion.Current)
                                End If
                            End If
                        End If
                        If FM_MODE <> "" Then
                            Dim DATETIME_STAMPX As DateTime = DATETIME_STAMP
                            If Format(DATETIME_STAMPX, "MM/dd/yyyy") = "01/01/0001" Then
                                DATETIME_STAMPX = Now + ASCMAIN1.NowTSD
                            End If
                            rowASTAUDT1.Item("TABLE_NAME") = TABLE_NAME
                            rowASTAUDT1.Item("KEY_VALUE") = KEY_VALUE
                            rowASTAUDT1.Item("COLUMN_NAME") = COLUMN_NAME
                            rowASTAUDT1.Item("USER_ID") = ASCMAIN1.USER_ID
                            rowASTAUDT1.Item("INIT_DATE") = DATETIME_STAMPX
                            rowASTAUDT1.Item("FM_MODE") = FM_MODE
                            rowASTAUDT1.Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                            rowASTAUDT1.Item("SELECTION_NO") = SELECTION_NO
                            rowASTAUDT1.Item("XNO") = XNO
                            dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)
                        End If
                    End If
                Next
            End If
        Next
        Update_Record_TDA("ASTAUDT1")
        dst.Tables("ASTAUDT1").Rows.Clear()
    End Sub

    Sub Create_BAs(ByVal TABLE_NAME As String)
        clsASCBASE1.Create_BAs(TABLE_NAME)
    End Sub

    Sub Clear_dst()
        clsASCBASE1.dst = New DataSet

        Setup_DataLayer()

        'dst = clsASCBASE1.dst
        'TBLs = clsASCBASE1.TBLs
        'TDAs = clsASCBASE1.TDAs
        'pROWs = clsASCBASE1.pROWs
        'DVWs = clsASCBASE1.DVWs
        'TBL_SCHEMAs = clsASCBASE1.TBL_SCHEMAs
        'F = clsASCBASE1.F
        'CR_params = clsASCBASE1.CR_params
        'ROWs = clsASCBASE1.ROWs
        'CMDs = clsASCBASE1.CMDs
        'BA_CMDs = clsASCBASE1.BA_CMDs

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
        F = clsASCBASE1.F
        CR_params = clsASCBASE1.CR_params
        ROWs = clsASCBASE1.ROWs
        CMDs = clsASCBASE1.CMDs
        BA_CMDs = clsASCBASE1.BA_CMDs
    End Sub

    Sub Update_BAs(ByVal TABLE_NAME As String)
        clsASCBASE1.Update_BAs(TABLE_NAME)
    End Sub

    Sub EnforceConstraints(Optional ByVal enforce As Boolean = True)
        If clsASCBASE1 Is Nothing Then Exit Sub

        Dim eMessage As String = clsASCBASE1.EnforceConstraints(enforce)
        If eMessage <> "" Then
            MsgBox(eMessage, MsgBoxStyle.OkOnly, "ABSolution will Terminate")
            If ASCMAIN1.Running_in_VS Then
                Stop
            Else
                Throw New Exception(eMessage)
            End If
        End If
    End Sub

    Public Function CodeValues _
    (ByVal CodeValueKey As String) As Dictionary(Of String, String)

        Dim CVL As Dictionary(Of String, String)
        CVL = ASCMAIN1.TACMAIN1.CodeValues(CodeValueKey)
        If CVL Is Nothing OrElse CVL.Count = 0 Then
            Dim SQL As String = ASCMAIN1.TACMAIN1.Get_Code_SQL_X(ASCMAIN1.ActiveForm.Name, COLUMN_NAME, "")
            If SQL <> "" Then
                For Each row As DataRow In ASCDATA1.GetDataTable(SQL).Rows
                    CVL.Add(row.Item(0) & "", row.Item(1) & "")
                Next
            End If
        End If
        Return CVL
    End Function

    Public Function nSoftwareKeys(ByVal key As String) As String

        Return ASCMAIN1.nSoftwareKeys(key)
    End Function

    Function Excel_Import(ByVal grd As UltraWinGrid.UltraGrid, Optional ByVal Band As Integer = 0) As Int32

        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
            openFileDialog1.Filter = "xls files (*.xls)|*.xls" ' |xlsx files (*.xlsx)|*.xlsx"
            'https://social.technet.microsoft.com/Forums/en-US/45ebe780-967e-4748-a146-9ec7a15626dc/unexpected-error-from-external-database-driver-1?forum=itmanager
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
            openFileDialog1.RestoreDirectory = True

            Excel_Import = -1

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        If FILENAME <> "" Then
            Try
                Dim dt As New DataTable
                Dim EXCEL_SHEET As String = ""

                If 1 <> 1 Then
                    Dim strConnection As String = "Provider=Microsoft.Jet.OleDb.4.0;" &
                    "data source=" & FILENAME & ";" &
                    "Extended Properties=Excel 8.0;"
                    Dim objConnection As New System.Data.OleDb.OleDbConnection(strConnection)
                    objConnection.Open()
                    Dim dbSchema As DataTable = objConnection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, Nothing)
                    If dbSchema.Rows.Count = 0 Then
                        MsgBox("No Sheets Found")
                        Exit Function
                    End If

                    EXCEL_SHEET = dbSchema.Rows(0).Item("TABLE_NAME")
                    If dbSchema.Rows.Count > 1 Then
                        Dim dtx As New DataTable
                        dtx.Columns.Add("TABLE_NAME")
                        Dim TABLE_NAME As String = ""
                        For Each row As DataRow In dbSchema.Rows

                            Dim SHEET_NAME As String = row.Item("TABLE_NAME")
                            If SHEET_NAME.EndsWith("$") Or (SHEET_NAME.StartsWith("'") And SHEET_NAME.EndsWith("$'")) Then
                                If SHEET_NAME.StartsWith("'") Then
                                    SHEET_NAME = Mid(SHEET_NAME, 2, Len(SHEET_NAME) - 2)
                                End If
                                TABLE_NAME = Mid(SHEET_NAME, 1, Len(SHEET_NAME) - 1)
                                dtx.Rows.Add(TABLE_NAME)
                            End If
                        Next
                        If dtx.Rows.Count = 1 Then
                            EXCEL_SHEET = TABLE_NAME & "$"
                        Else
                            Dim frmmsg As New ASFMSGBF
                            frmmsg.Show_grd(dtx, Me, "Select Excel Sheet to Load")
                            If frmmsg.grow Is Nothing Then
                                Exit Function
                            End If
                            EXCEL_SHEET = frmmsg.grow.Cells("TABLE_NAME").Text & "$"
                        End If
                    Else

                    End If

                    Dim strSQL As String = "SELECT * FROM [" & EXCEL_SHEET & "]"
                    Dim objCommand As New System.Data.OleDb.OleDbCommand(strSQL, objConnection)
                    Dim objAdapter As New System.Data.OleDb.OleDbDataAdapter(strSQL, objConnection)
                    objAdapter.FillSchema(dt, SchemaType.Source)
                    Excel_Import_DataTable_Intitialization(dt)
                    objAdapter.Fill(dt)
                    objConnection.Close()

                Else


                    Dim oWB As SpreadsheetGear.IWorkbook
                    Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
                    Dim range As SpreadsheetGear.IRange = Nothing


                    oWB = SpreadsheetGear.Factory.GetWorkbook(FILENAME)


                    Dim rangeCopyFrom As SpreadsheetGear.IRange = Nothing
                    Dim rangePaste_To As SpreadsheetGear.IRange = Nothing
                    ' Dim range As SpreadsheetGear.IRange = Nothing

                    Dim XL_ROWS As Integer = 0
                    Dim XL_COLS As Integer = 0



                    If oWB.Sheets.Count = 0 Then
                        MsgBox("No Sheets Found")
                        Exit Function
                    End If

                    EXCEL_SHEET = oWB.Worksheets(0).Name
                    If oWB.Sheets.Count > 1 Then
                        Dim dtx As New DataTable
                        dtx.Columns.Add("TABLE_NAME")
                        Dim TABLE_NAME As String = ""
                        For Each oSheet In oWB.Sheets
                            TABLE_NAME = oSheet.Name
                            dtx.Rows.Add(TABLE_NAME)
                        Next
                        Dim frmmsg As New ASFMSGBF
                        frmmsg.Show_grd(dtx, Me, "Select Excel Sheet to Load")
                        If frmmsg.grow Is Nothing Then
                            Exit Function
                        End If
                        EXCEL_SHEET = frmmsg.grow.Cells("TABLE_NAME").Text ' & "$"
                    End If

                    oSheet = oWB.Sheets(EXCEL_SHEET)
                    With dt
                        For j As Integer = 0 To oSheet.UsedRange.ColumnCount - 1
                            Dim jC As String = Excel_Cell0(-1, j)
                            .Columns.Add(jC)
                        Next
                    End With

                    Excel_Import_DataTable_Intitialization(dt)
                    ' watch out - UsedRange starts at 1st used row, so if blank rows at top, they are not included
                    Dim i As Int64 = 0
                    For Each xlRow As SpreadsheetGear.IRange In oSheet.UsedRange.Rows ' For i As Int64 = 0 To oSheet.UsedRange.RowCount - 1
                        Dim row As DataRow = dt.NewRow
                        Dim row_has_data As Boolean = False
                        For j As Integer = 0 To oSheet.UsedRange.ColumnCount - 1
                            'If oSheet.Cells(i, j).Value & "" <> "" Then
                            '    row_has_data = True
                            '    row.Item(j) = oSheet.Cells(i, j).Value
                            'End If
                            If xlRow.Cells(0, j).Value & "" <> "" Then
                                row_has_data = True
                                row.Item(j) = xlRow.Cells(0, j).Value
                            End If
                        Next
                        i += 1
                        If row_has_data Then dt.Rows.Add(row)
                    Next

                End If


                Dim F As New ASFEXCL1(Me)
                F.dt = dt
                F.grd = grd
                F.Band = Band
                F.ShowDialog()

                If F.STATUS = "OK" Then
                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Now Loading from Workbook")
                    Application.DoEvents()

                    Dim load_by_table As Boolean = False
                    Dim load_handled As Boolean = False
                    Excel_Import_Pre_Process(grd, load_by_table, load_handled, F)

                    Dim dtx As DataTable
                    If F.grdExcel.Tag & "" = "Excel" Then ' set the tag to Excel in the form overrides Excel_Import_Pre_Process to force this routine to read excel and place data into a new datatable if oledb does not work out - like at AHA for the CC Trans
                        Dim excel As New Microsoft.Office.Interop.Excel.Application
                        Dim wb As Microsoft.Office.Interop.Excel.Workbook = excel.Workbooks.Open(FILENAME)

                        Dim ws As Microsoft.Office.Interop.Excel.Worksheet
                        If wb.Worksheets.Count = 1 Then
                            ws = wb.Worksheets(1)
                        Else
                            ws = wb.Worksheets(Mid(EXCEL_SHEET, 1, Len(EXCEL_SHEET) - 1))
                        End If

                        dtx = F.dt.Clone

                        For rowx As Int64 = 1 To F.grdExcel.Rows.Count + 1
                            If rowx > F.xRi + 1 Then
                                Dim rx2 As DataRow = dtx.NewRow
                                For colx As Integer = F.xCi + 1 To F.grdExcel.DisplayLayout.Bands(0).Columns.Count
                                    rx2.Item(colx - 1) = ws.Cells(rowx, colx).value
                                Next
                                dtx.Rows.Add(rx2)
                            End If
                        Next
                        F.grdExcel.DataSource = dtx

                        ws = Nothing
                        wb.Close()
                        wb = Nothing
                        excel.Quit()
                        excel = Nothing

                    End If


                    If Not load_handled Then
                        Dim row_count As Int32 = 0

                        Dim grdds As DataTable = DirectCast(grd.DataSource, DataTable)
                        Dim PK As New List(Of String)
                        If grdds.PrimaryKey.Length <> 0 Then
                            For Each dc As DataColumn In grdds.PrimaryKey
                                PK.Add(dc.ColumnName)
                            Next
                        End If

                        Dim tbldups As New DataTable '  = grdds.Clone
                        For Each dc As DataColumn In grdds.Columns
                            tbldups.Columns.Add(dc.ColumnName, dc.DataType)
                        Next


                        Dim COLsPK As New List(Of String)
                        Dim COLs() As String = Nothing
                        ReDim COLs(0)
                        Dim ic As Integer = -1
                        With grd.DisplayLayout.Bands(Band)
                            For I As Integer = 0 To .Columns.Count - 1
                                If Not .Columns(I).Hidden Then
                                    ic += 1
                                    ReDim Preserve COLs(ic)
                                    COLs(COLs.Length - 1) = .Columns(I).Key
                                    If PK.Contains(.Columns(I).Key) Then
                                        COLsPK.Add(.Columns(I).Key)
                                    End If
                                End If
                            Next
                        End With

                        Dim c As Integer = 0
                        Dim COLx As New Dictionary(Of String, Integer)
                        With F.grdExcel.DisplayLayout.Bands(Band)
                            For I As Integer = 0 To .Columns.Count - 1
                                If Not .Columns(I).Hidden Then
                                    'COLx.Add(COLs(c), I)
                                    COLx.Add(.Columns(I).Tag, I)
                                    c = c + 1
                                    If c = COLs.Length Then ' If c + 1 = COLs.Length Then ' +1 DIDNT WORK FOR JH RETAILS IMPORT FINLAY
                                        Exit For
                                    End If
                                End If
                            Next
                        End With

                        Dim CONSECUTIVE_BLANK_ROWS As Int32 = 0

                        grd.BeginUpdate()
                        'grd.SuspendRowSynchronization()
                        'grd.Visible = False

                        dt.Columns.Add("EXCEL_UPLOAD_STATUS")
                        Dim rows_failed As Int64 = 0
                        loading_grd_from_Excel = True
                        ' see GLFBUDM1 grdGLTACCT2_BeforeRowUpdate for example on how to use EMsg2
                        kill_switch = False

                        ASCMAIN1.Progress("Now Loading Rows")

                        For Each gr As UltraWinGrid.UltraGridRow In F.grdExcel.Rows

                            If Not gr.Hidden Then ' why would we test for gr.hidden? this is preventing the loading of the spreadsheets - answer - we hide the rows which are prior to the starting row, as selected by the user
                                row_count += 1
                                If row_count = 1 Then
                                    ASCMAIN1.Progress("Now Loading Rows", CStr(row_count))
                                Else
                                    ASCMAIN1.Progress("-", CStr(row_count))
                                End If
                                Dim GRX As String = ""
                                For ICOL As Integer = 0 To gr.Cells.Count - 1
                                    GRX &= ":" & gr.Cells(ICOL).Value
                                Next
                                GRX = Mid(GRX, 2)

                                Try

                                    ' see if row exists
                                    Dim exists As Boolean = False
                                    Dim grdrow As UltraWinGrid.UltraGridRow = Nothing

                                    If COLsPK.Count <> 0 Then
                                        For Each grdrow In grd.Rows
                                            Dim this_row_matches As Boolean = True
                                            For Each col As String In COLsPK
                                                If grdrow.Cells(col).Value & "" <> gr.Cells(COLx(col)).Value & "" Then
                                                    this_row_matches = False
                                                    Exit For
                                                End If
                                            Next
                                            If this_row_matches Then
                                                exists = True
                                                Dim rowdups As DataRow = tbldups.NewRow
                                                For i As Int32 = 0 To tbldups.Columns.Count - 1
                                                    rowdups.Item(i) = grdrow.Cells(i).Value
                                                Next
                                                tbldups.Rows.Add(rowdups)
                                                Exit For
                                            End If
                                        Next
                                    End If

                                    If exists Then

                                        If grd.DisplayLayout.Override.AllowUpdate <> DefaultableBoolean.False Then

                                            For Each col As String In COLx.Keys
                                                If grd.DisplayLayout.Bands(Band).Columns(col).CellActivation = UltraWinGrid.Activation.AllowEdit Then
                                                    If grdrow.Cells(col).Column.DataType.ToString = "System.DateTime" Then
                                                        If (gr.Cells(COLx(col)).Value & "") <> "" Then
                                                            If IsNumeric(gr.Cells(COLx(col)).Value) Then
                                                                grdrow.Cells(col).Value = CDate(DateTime.FromOADate(gr.Cells(COLx(col)).Value))
                                                            Else
                                                                grdrow.Cells(col).Value = gr.Cells(COLx(col)).Value
                                                            End If
                                                        End If
                                                    Else
                                                        grdrow.Cells(col).Value = gr.Cells(COLx(col)).Value
                                                    End If

                                                End If
                                            Next
                                            grdrow.Update()

                                        End If

                                    Else

                                        If grd.DisplayLayout.Override.AllowAddNew <> UltraWinGrid.AllowAddNew.No Then
                                            If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.DataChanged Then
                                                grd.ActiveRow.CancelUpdate()
                                            End If

                                            If Not load_by_table Then
                                                grd.DisplayLayout.Bands(Band).AddNew.Activate()
                                                For Each col As String In COLx.Keys
                                                    If grd.DisplayLayout.Bands(Band).Columns(col).CellActivation = UltraWinGrid.Activation.AllowEdit Then
                                                        grd.ActiveCell = grd.ActiveRow.Cells(col)
                                                        grd.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
                                                        grd.ActiveRow.Cells(col).Value = gr.Cells(COLx(col)).Value
                                                        grd.PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)
                                                        If grd.ActiveCell.IsInEditMode Then
                                                            grd.ActiveRow.CancelUpdate()
                                                            gr.Cells("EXCEL_UPLOAD_STATUS").Value = "FAILED"
                                                            rows_failed += 1
                                                            Exit For
                                                        End If
                                                    End If
                                                Next
                                                If grd.ActiveRow IsNot Nothing Then
                                                    If grd.ActiveRow.IsAddRow Then
                                                        '  Debug.Print(grd.ActiveRow.Cells("CUST_STORE_NO").Value)
                                                        EMsg2 = ""
                                                        Dim WORKED As Boolean = grd.ActiveRow.Update

                                                        If kill_switch Then
                                                            grd.ActiveRow.CancelUpdate()
                                                            Throw New Exception("Upload Terminated")
                                                        End If

                                                        If Not WORKED Then
                                                            '  If grd.DisplayLayout.Bands(0).Columns.Contains("EXCEL_UPLOAD_STATUS") Then
                                                            If EMsg2 <> "" Then
                                                                gr.Cells("EXCEL_UPLOAD_STATUS").Value = EMsg2
                                                            Else
                                                                If grdError IsNot Nothing Then gr.Cells("EXCEL_UPLOAD_STATUS").Value = grdError.ErrorText ' "FAILED - ERROR"
                                                            End If
                                                            'End If
                                                            rows_failed += 1
                                                            If grd.ActiveRow IsNot Nothing Then
                                                                grd.ActiveRow.CancelUpdate()
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                Dim DL As Int32 = 0
                                                Dim row As DataRow = grdds.NewRow
                                                For Each col As String In COLx.Keys
                                                    If grd.DisplayLayout.Bands(Band).Columns(col).CellActivation = UltraWinGrid.Activation.AllowEdit Then
                                                        If grdds.Columns(col).DataType.ToString = "System.Decimal" _
                                                        Or grdds.Columns(col).DataType.ToString = "System.Double" _
                                                        Or grdds.Columns(col).DataType.ToString = "System.Int16" _
                                                        Or grdds.Columns(col).DataType.ToString = "System.Int32" _
                                                        Or grdds.Columns(col).DataType.ToString = "System.Int64" _
                                                        Or grdds.Columns(col).DataType.ToString = "System.Integer" Then
                                                            row.Item(col) = Val(gr.Cells(COLx(col)).Value & "")
                                                        Else
                                                            row.Item(col) = gr.Cells(COLx(col)).Value
                                                        End If
                                                        'row.Item(col) = gr.Cells(COLx(col)).Value
                                                        DL += Len(row.Item(col) & "")
                                                    End If
                                                Next

                                                Excel_Import_Custom_Processing_row(row, gr, grdds)

                                                If DL = 0 Then
                                                    CONSECUTIVE_BLANK_ROWS += 1
                                                Else
                                                    CONSECUTIVE_BLANK_ROWS = 0
                                                End If
                                                If CONSECUTIVE_BLANK_ROWS >= 10 Then
                                                    Exit For
                                                End If
                                                grdds.Rows.Add(row)
                                            End If
                                        End If
                                    End If
                                Catch ex As Exception
                                    If kill_switch OrElse MsgBox("Exception Occurred:" & vbCrLf & ex.Message & vbCrLf & vbCrLf & GRX, MsgBoxStyle.OkCancel, "Error Working with Row") = MsgBoxResult.Cancel Then
                                        Me.Cursor = Cursors.Default
                                        ASCMAIN1.Progress("")
                                        Application.DoEvents()
                                        grd.EndUpdate()
                                        loading_grd_from_Excel = False
                                        Exit Function
                                    End If
                                End Try

                            End If
                        Next

                        loading_grd_from_Excel = False
                        'grd.Visible = True
                        'grd.ResumeRowSynchronization()
                        grd.EndUpdate()

                        If rows_failed Then
                            Using fr As New ASFMSGBF
                                fr.Show_grd(dt, Me, "Some Rows Failed to Update - Please Check Last Column for Messages")
                            End Using
                        End If

                        If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.DataChanged Then
                            grd.ActiveRow.CancelUpdate()
                        End If
                        grd.DisplayLayout.Bands(Band).SortedColumns.RefreshSort(False)
                        If grd.Rows.Count > 0 Then
                            grd.ActiveRow = grd.Rows(0)
                        End If

                        Excel_Import = row_count

                        If tbldups.Rows.Count <> 0 Then
                            Dim Fmsg As New ASFMSGBF
                            Fmsg.Show_grd(tbldups, ASCMAIN1.ActiveForm, "Duplicate Rows from Spreadsheet")
                        End If
                    End If

                    Excel_Import_Post_Process(grd, F)

                End If
                F.Dispose()

            Catch ex As Exception
                MsgBox("Exception Occurred:" & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Error Opening Excel Workbook")
            Finally

            End Try
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
        Application.DoEvents()
    End Function


    Function Excel_Import_SG(ByVal grd As UltraWinGrid.UltraGrid, Optional ByVal Band As Integer = 0) As Int32

        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
            ' openFileDialog1.Filter = "xls files (*.xls)|*.xls" ' |xlsx files (*.xlsx)|*.xlsx"
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
            openFileDialog1.RestoreDirectory = True

            Excel_Import_SG = -1

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        Dim oWB As SpreadsheetGear.IWorkbook
        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing

        'Dim rangeCopyFrom As SpreadsheetGear.IRange
        'Dim rangePaste_To As SpreadsheetGear.IRange

        'Dim XL_ROWS As Integer
        'Dim XL_COLS As Integer

        If FILENAME <> "" Then
            Try
                oWB = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
                If oWB.Worksheets.Count = 0 Then
                    MsgBox("No Sheets Found")
                    Exit Function
                End If

                Dim EXCEL_SHEET As String = oWB.Worksheets(0).Name
                If oWB.Worksheets.Count > 1 Then
                    Dim dtx As New DataTable
                    dtx.Columns.Add("TABLE_NAME")
                    Dim TABLE_NAME As String = ""
                    For Each oSheet In oWB.Worksheets
                        Dim SHEET_NAME As String = oSheet.Name
                        TABLE_NAME = SHEET_NAME ' Mid(SHEET_NAME, 1, Len(SHEET_NAME) - 1)
                        dtx.Rows.Add(TABLE_NAME)
                    Next
                    Using frmmsg As New ASFMSGBF
                        frmmsg.Show_grd(dtx, Me, "Select Excel Sheet to Load")
                        If frmmsg.grow Is Nothing Then
                            Exit Function
                        End If
                        EXCEL_SHEET = frmmsg.grow.Cells("TABLE_NAME").Text ' & "$"
                    End Using
                End If
                oSheet = oWB.Sheets(EXCEL_SHEET)

                ' need to get the data out of SG sheet and into a datatable
                Dim dt As New DataTable
                Excel_Import_DataTable_Intitialization(dt)

                Dim vMax As Integer = -1
                Dim grdcols As New Dictionary(Of Integer, DataColumn)
                For Each grdcol As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(Band).Columns
                    If Not grdcol.Hidden And grdcol.CellActivation = UltraWinGrid.Activation.AllowEdit Then
                        Dim v As Integer = grdcol.Header.VisiblePosition
                        Dim d As New DataColumn(grdcol.Key, grdcol.DataType)
                        d.Caption = grdcol.Header.Caption
                        grdcols.Add(v, d)
                        If v > vMax Then vMax = v
                    End If
                Next

                For i As Integer = 0 To vMax
                    If grdcols.ContainsKey(i) Then
                        dt.Columns.Add(grdcols(i))
                    End If
                Next


                For r As Integer = 1 To oSheet.Cells.RowCount
                    If oSheet.Cells(r, 0).Value & "" = "" Then Exit For
                    Dim row As DataRow = dt.NewRow
                    For c As Integer = 0 To dt.Columns.Count - 1
                        If oSheet.Cells(r, c).Value & "" <> "" Then
                            row.Item(c) = oSheet.Cells(r, c).Value
                        End If
                    Next
                    dt.Rows.Add(row)
                Next


                Dim F As New ASFEXCL1(Me)
                F.dt = dt
                F.grd = grd
                F.Band = Band
                F.ShowDialog()

                If F.STATUS = "OK" Then
                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Now Loading from Workbook")

                    '    Dim F As ASFEXCL1 = Nothing

                    Dim load_by_table As Boolean = False
                    Dim load_handled As Boolean = False
                    Excel_Import_SG = Excel_Import_Pre_Process_SG(grd, dt, load_by_table, load_handled, F)

                    Dim dtx As DataTable
                    If F.grdExcel.Tag & "" = "Excel" Then ' set the tag to Excel in the form overrides Excel_Import_Pre_Process to force this routine to read excel and place data into a new datatable if oledb does not work out - like at AHA for the CC Trans
                        Dim excel As New Microsoft.Office.Interop.Excel.Application
                        Dim wb As Microsoft.Office.Interop.Excel.Workbook = excel.Workbooks.Open(FILENAME)

                        Dim ws As Microsoft.Office.Interop.Excel.Worksheet
                        If wb.Worksheets.Count = 1 Then
                            ws = wb.Worksheets(1)
                        Else
                            ws = wb.Worksheets(Mid(EXCEL_SHEET, 1, Len(EXCEL_SHEET) - 1))
                        End If

                        dtx = F.dt.Clone

                        For rowx As Int64 = 1 To F.grdExcel.Rows.Count + 1
                            If rowx > F.xRi + 1 Then
                                Dim rx2 As DataRow = dtx.NewRow
                                For colx As Integer = F.xCi + 1 To F.grdExcel.DisplayLayout.Bands(0).Columns.Count
                                    rx2.Item(colx - 1) = ws.Cells(rowx, colx).value
                                Next
                                dtx.Rows.Add(rx2)
                            End If
                        Next
                        F.grdExcel.DataSource = dtx

                        ws = Nothing
                        wb.Close()
                        wb = Nothing
                        excel.Quit()
                        excel = Nothing

                    End If


                    If Not load_handled Then
                        Dim row_count As Int32 = 0

                        Dim grdds As DataTable = DirectCast(grd.DataSource, DataTable)
                        Dim PK As New List(Of String)
                        If grdds.PrimaryKey.Length <> 0 Then
                            For Each dc As DataColumn In grdds.PrimaryKey
                                PK.Add(dc.ColumnName)
                            Next
                        End If

                        Dim tbldups As New DataTable '  = grdds.Clone
                        For Each dc As DataColumn In grdds.Columns
                            tbldups.Columns.Add(dc.ColumnName, dc.DataType)
                        Next


                        Dim COLsPK As New List(Of String)
                        Dim COLs() As String = Nothing
                        ReDim COLs(0)
                        Dim ic As Integer = -1
                        With grd.DisplayLayout.Bands(Band)
                            For I As Integer = 0 To .Columns.Count - 1
                                If Not .Columns(I).Hidden Then
                                    ic += 1
                                    ReDim Preserve COLs(ic)
                                    COLs(COLs.Length - 1) = .Columns(I).Key
                                    If PK.Contains(.Columns(I).Key) Then
                                        COLsPK.Add(.Columns(I).Key)
                                    End If
                                End If
                            Next
                        End With

                        Dim c As Integer = 0
                        Dim COLx As New Dictionary(Of String, Integer)
                        With F.grdExcel.DisplayLayout.Bands(Band)
                            For I As Integer = 0 To .Columns.Count - 1
                                If Not .Columns(I).Hidden Then
                                    'COLx.Add(COLs(c), I)
                                    COLx.Add(.Columns(I).Key, I) ' .Tag
                                    c = c + 1
                                    If c = COLs.Length Then ' If c + 1 = COLs.Length Then ' +1 DIDNT WORK FOR JH RETAILS IMPORT FINLAY
                                        Exit For
                                    End If
                                End If
                            Next
                        End With

                        Dim CONSECUTIVE_BLANK_ROWS As Int32 = 0

                        grd.BeginUpdate()

                        dt.Columns.Add("EXCEL_UPLOAD_STATUS")
                        Dim rows_failed As Int64 = 0

                        For Each gr As UltraWinGrid.UltraGridRow In F.grdExcel.Rows

                            If Not gr.Hidden Then ' why would we test for gr.hidden? this is preventing the loading of the spreadsheets - answer - we hide the rows which are prior to the starting row, as selected by the user
                                row_count += 1

                                Dim GRX As String = ""
                                For ICOL As Integer = 0 To gr.Cells.Count - 1
                                    GRX &= ":" & gr.Cells(ICOL).Value
                                Next
                                GRX = Mid(GRX, 2)

                                Try

                                    ' see if row exists
                                    Dim exists As Boolean = False
                                    Dim grdrow As UltraWinGrid.UltraGridRow = Nothing

                                    If COLsPK.Count <> 0 Then
                                        For Each grdrow In grd.Rows
                                            Dim this_row_matches As Boolean = True
                                            For Each col As String In COLsPK
                                                If grdrow.Cells(col).Value & "" <> gr.Cells(COLx(col)).Value & "" Then
                                                    this_row_matches = False
                                                    Exit For
                                                End If
                                            Next
                                            If this_row_matches Then
                                                exists = True
                                                Dim rowdups As DataRow = tbldups.NewRow
                                                For i As Int32 = 0 To tbldups.Columns.Count - 1
                                                    rowdups.Item(i) = grdrow.Cells(i).Value
                                                Next
                                                tbldups.Rows.Add(rowdups)
                                                Exit For
                                            End If
                                        Next
                                    End If

                                    If exists Then

                                        If grd.DisplayLayout.Override.AllowUpdate <> DefaultableBoolean.False Then

                                            For Each col As String In COLx.Keys
                                                If grd.DisplayLayout.Bands(Band).Columns(col).CellActivation = UltraWinGrid.Activation.AllowEdit Then
                                                    grdrow.Cells(col).Value = gr.Cells(COLx(col)).Value
                                                End If
                                            Next
                                            grdrow.Update()

                                        End If

                                    Else

                                        If grd.DisplayLayout.Override.AllowAddNew <> UltraWinGrid.AllowAddNew.No Then
                                            If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.DataChanged Then
                                                grd.ActiveRow.CancelUpdate()
                                            End If

                                            If Not load_by_table Then
                                                grd.DisplayLayout.Bands(Band).AddNew.Activate()
                                                For Each col As String In COLx.Keys
                                                    If grd.DisplayLayout.Bands(Band).Columns(col).CellActivation = UltraWinGrid.Activation.AllowEdit Then
                                                        grd.ActiveCell = grd.ActiveRow.Cells(col)
                                                        grd.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
                                                        grd.ActiveRow.Cells(col).Value = gr.Cells(COLx(col)).Value
                                                        grd.PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)
                                                        If grd.ActiveCell.IsInEditMode Then
                                                            grd.ActiveRow.CancelUpdate()
                                                            gr.Cells("EXCEL_UPLOAD_STATUS").Value = "FAILED"
                                                            rows_failed += 1
                                                            Exit For
                                                        End If
                                                    End If
                                                Next
                                                If grd.ActiveRow IsNot Nothing Then
                                                    If grd.ActiveRow.IsAddRow Then
                                                        '  Debug.Print(grd.ActiveRow.Cells("CUST_STORE_NO").Value)
                                                        Dim WORKED As Boolean = grd.ActiveRow.Update
                                                        If Not WORKED Then
                                                            If grd.DisplayLayout.Bands(0).Columns.Contains("EXCEL_UPLOAD_STATUS") Then
                                                                gr.Cells("EXCEL_UPLOAD_STATUS").Value = grdError.ErrorText ' "FAILED - ERROR"
                                                            End If
                                                            rows_failed += 1
                                                            If grd.ActiveRow IsNot Nothing Then
                                                                grd.ActiveRow.CancelUpdate()
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                Dim DL As Int32 = 0
                                                Dim row As DataRow = grdds.NewRow
                                                For Each col As String In COLx.Keys
                                                    If grd.DisplayLayout.Bands(Band).Columns(col).CellActivation = UltraWinGrid.Activation.AllowEdit Then
                                                        If grdds.Columns(col).DataType.ToString = "System.Decimal" _
                                                        Or grdds.Columns(col).DataType.ToString = "System.Double" _
                                                        Or grdds.Columns(col).DataType.ToString = "System.Int16" _
                                                        Or grdds.Columns(col).DataType.ToString = "System.Int32" _
                                                        Or grdds.Columns(col).DataType.ToString = "System.Int64" _
                                                        Or grdds.Columns(col).DataType.ToString = "System.Integer" Then
                                                            row.Item(col) = Val(gr.Cells(COLx(col)).Value & "")
                                                        Else
                                                            row.Item(col) = gr.Cells(COLx(col)).Value
                                                        End If
                                                        'row.Item(col) = gr.Cells(COLx(col)).Value
                                                        DL += Len(row.Item(col) & "")
                                                    End If
                                                Next

                                                Excel_Import_Custom_Processing_row(row, gr, grdds)

                                                If DL = 0 Then
                                                    CONSECUTIVE_BLANK_ROWS += 1
                                                Else
                                                    CONSECUTIVE_BLANK_ROWS = 0
                                                End If
                                                If CONSECUTIVE_BLANK_ROWS >= 10 Then
                                                    Exit For
                                                End If
                                                grdds.Rows.Add(row)
                                            End If
                                        End If
                                    End If
                                Catch ex As Exception
                                    If MsgBox("Exception Occurred:" & vbCrLf & ex.Message & vbCrLf & vbCrLf & GRX, MsgBoxStyle.OkCancel, "Error Working with Row") = MsgBoxResult.Cancel Then
                                        Me.Cursor = Cursors.Default
                                        ASCMAIN1.Progress("")
                                        Application.DoEvents()
                                        grd.EndUpdate()
                                        Exit Function
                                    End If
                                End Try

                            End If
                        Next

                        'grd.Visible = True
                        'grd.ResumeRowSynchronization()
                        grd.EndUpdate()

                        If rows_failed Then
                            Using fr As New ASFMSGBF
                                fr.Show_grd(dt, Me, "Some Rows Failed to Update - Please Check Last Column for Status")
                            End Using
                        End If

                        If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.DataChanged Then
                            grd.ActiveRow.CancelUpdate()
                        End If
                        grd.DisplayLayout.Bands(Band).SortedColumns.RefreshSort(False)
                        If grd.Rows.Count > 0 Then
                            grd.ActiveRow = grd.Rows(0)
                        End If

                        Excel_Import_SG = row_count

                        If tbldups.Rows.Count <> 0 Then
                            Dim Fmsg As New ASFMSGBF
                            Fmsg.Show_grd(tbldups, ASCMAIN1.ActiveForm, "Duplicate Rows from Spreadsheet")
                        End If
                    End If

                    Excel_Import_Post_Process(grd, F)

                End If
                F.Dispose()

            Catch ex As Exception
                MsgBox("Exception Occurred:" & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Error Opening Excel Workbook")
            Finally

            End Try
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
        Application.DoEvents()
    End Function

    Public Overridable Sub Excel_Import_Custom_Processing_row _
    (ByVal row As DataRow, ByVal grow As UltraWinGrid.UltraGridRow,
     Optional ByVal tbl As DataTable = Nothing)

    End Sub

    Public Overridable Sub Excel_Import_DataTable_Intitialization _
    (ByRef dt As DataTable)

    End Sub

    Public Overridable Sub Excel_Import_Pre_Process _
    (ByVal grd As UltraWinGrid.UltraGrid,
     Optional ByRef load_by_table As Boolean = False,
     Optional ByRef load_handled As Boolean = False,
     Optional ByRef F As ASFEXCL1 = Nothing)

    End Sub

    Public Overridable Function Excel_Import_Pre_Process_SG _
    (ByVal grd As UltraWinGrid.UltraGrid, dt As DataTable,
     Optional ByRef load_by_table As Boolean = False,
     Optional ByRef load_handled As Boolean = False,
     Optional ByRef F As ASFEXCL1 = Nothing) As Int64
        Return 0 ' number of records that were successfully processed. return -1 to cancel
    End Function

    Public Overridable Sub Excel_Import_Post_Process _
    (ByVal grd As UltraWinGrid.UltraGrid, F As ASFEXCL1)

    End Sub

    Sub Load_Tool_Tips()

        ASCMAIN1.sql = "Select * from ASTTTIP1 where FORM_NAME = '" & MENU_ITEM_OBJECT & "'"
        Create_TDA(dst.Tables.Add, "ASTTTIP1", "**", 0)
        Fill_Records("ASTTTIP1")
        'dst.Tables.Add(ASCDATA1.GetDataTable)

        ' set this value to however many milliseconds the tooltip delay should be
        timer.Interval = 500
        ' when the timer ticks we want our method to be called
        AddHandler timer.Tick, AddressOf OnTimerTick
        'tooltip_msg = String.Empty

    End Sub

    Sub Set_Security_Context()
        If Absx1.htbABSSecurityCodes.Count <> 0 Then
            For Each ContainerControl As DictionaryEntry In Absx1.htbABSSecurityCodes
                Dim SECURITY_CODE As String = ""
                Dim enabled As Boolean = False
                Dim C As Control = DirectCast(ContainerControl.Key, Control)
                For i As Integer = 1 To ContainerControl.Value.ToString.Length / 2
                    SECURITY_CODE = Mid(ContainerControl.Value, (i - 1) * 2 + 1, 2)
                    If InStr(ASCMAIN1.USER_SECURITY_CODEs, SECURITY_CODE) <> 0 Then
                        enabled = True
                        Exit For
                    End If
                Next
                If Not enabled Then
                    Set_Read_Only(C, True)
                End If
            Next
        End If
    End Sub

    Function Get_SEL_CODE_VALUEs(
    ByVal grd As UltraWinGrid.UltraGrid,
    Optional ByVal return_empty_if_all_selected As Boolean = True
    ) As String
        Dim CODE_VALUEs As String = ""
        Dim all_selected As Boolean = True
        For Each row As DataRow In DirectCast(grd.DataSource, DataTable).Rows
            If row.Item("SEL") = "1" Then
                CODE_VALUEs &= ",'" & row.Item("CODE_VALUE") & "'"
            Else
                all_selected = False
            End If
        Next
        If all_selected Then
            If return_empty_if_all_selected Then
                CODE_VALUEs = ""
            End If
        Else
            CODE_VALUEs = Mid(CODE_VALUEs, 2)
        End If

        Return CODE_VALUEs
    End Function

    Sub Setup_CODE_VALUEs(ByVal grd As UltraWinGrid.UltraGrid,
                          ByVal TABLE_NAME As String,
                          ByVal COLUMN_NAME_CODE As String,
                          ByVal COLUMN_NAME_DESC As String)
        Dim DT As New DataTable
        DT = ASCDATA1.GetDataTable("Select " & COLUMN_NAME_CODE & " CODE_VALUE, " & COLUMN_NAME_DESC & " DESC_VALUE, '1' SEL from " & TABLE_NAME & "", TABLE_NAME, 1)
        DT.Columns("SEL").ReadOnly = False
        grd.DataSource = DT
    End Sub

    Public Overridable Sub CustomSummary_DataRows(
    ByVal summarySettings As UltraWinGrid.SummarySettings,
    ByVal row As UltraWinGrid.UltraGridRow,
    ByRef CustomValue As Double,
    ByVal grd As UltraWinGrid.UltraGrid)

    End Sub

    Public Overridable Function CustomSummary_End(
    ByVal summarySettings As UltraWinGrid.SummarySettings,
    ByVal rows As UltraWinGrid.RowsCollection,
    ByVal CustomValue As Double,
    ByVal grd As UltraWinGrid.UltraGrid) As Double

        Return CustomValue
    End Function

    Public Overridable Function CustomStringSummary_End(
    ByVal summarySettings As UltraWinGrid.SummarySettings,
    ByVal rows As UltraWinGrid.RowsCollection,
    ByVal CustomValue As String,
    ByVal grd As UltraWinGrid.UltraGrid) As String

        Return CustomValue
    End Function

    Sub Show_Document(ByVal FILENAME As String)
        Dim p As Process = Nothing
        Try
            If My.Computer.FileSystem.FileExists(FILENAME) Then
                p = Process.Start(FILENAME)
                If p IsNot Nothing Then
                    p.Dispose()
                End If
            End If

        Catch ex As Exception

        Finally

        End Try

    End Sub

    Function Prepare_Report_dst(
    ByVal RPT As String,
    ByVal perform_fill As Boolean,
    ByVal ParamArray parms() As Object) As ASCBASE1

        Dim rptClass As ASFSRPTM = Load_rptClass(RPT)
        Return rptClass.Prepare_dst(perform_fill, parms)
    End Function

    Function Initialize_Report(ByVal rptFORMNAME As String,
                               Optional ByVal ExportFilenameDefault As String = "",
                               Optional ByVal ExportFormatDefault As String = "RPT",
                               Optional ByVal ArchiveReportsDefault As Boolean = False)

        If Not REPORTS.ContainsKey(rptFORMNAME) Then
            REPORTS.Add(rptFORMNAME, Load_rptClass(rptFORMNAME))
            With REPORTS(rptFORMNAME)
                .Prepare_dst(False, "")
            End With
        End If

        With REPORTS(rptFORMNAME)
            .ExportFilenameDefault = ExportFilenameDefault
            .ExportFormatDefault = ExportFormatDefault
            .ArchiveReportsDefault = ArchiveReportsDefault
        End With

        Return "" ' LATER
    End Function

    Sub Process_Report(ByVal rptFILENAME As String, ByVal ParamArray parms() As Object)
        REPORTS(rptFILENAME).Fill_Records_RPT(parms)
        With REPORTS(rptFILENAME)
            .Print_Report_Begin()
            .Print_Report()
            .Print_Report_End()
        End With
    End Sub
    Function Load_rptClass(ByVal MENU_ITEM_OBJECT As String) As ASFSRPTM
        Dim MODULE_ID As String = Mid(MENU_ITEM_OBJECT, 1, 2)
        Dim sLocation As String = ""

        Dim buildType As String
#If DEBUG Then
        buildType = "x64\Debug"
#Else
        buildType = "x64\Release"
#End If

        If ASCMAIN1.Running_in_VS Then
            sLocation = ASCMAIN1.Folders("root") & MODULE_ID & "\bin\" & buildType & "\" & MODULE_ID & ".dll"
        Else
            sLocation = ASCMAIN1.Folders("bin") & MODULE_ID & ".dll"
        End If

        Dim sType As String = MODULE_ID & "." & MENU_ITEM_OBJECT
        Dim formAsm As System.Reflection.Assembly = System.Reflection.Assembly.LoadFrom(sLocation)

        Dim ClassType As Type = formAsm.GetType(sType)
        Dim Classobj As New Object
        Classobj = Activator.CreateInstance(ClassType)

        Dim rptClass As ASFSRPTM = CType(Classobj, ASFSRPTM)
        rptClass.remotely_controlled = True

        rptClass.MENU_ITEM_TYPE = "R"
        rptClass.MENU_ITEM_OBJECT = MENU_ITEM_OBJECT
        rptClass.XNO = ASCMAIN1.Next_Control_No(MENU_ITEM_OBJECT & ".XNO")

        Return rptClass ' CType(Classobj, ASFSRPTM)
    End Function

    Private Class CustomSummary
        Implements UltraWinGrid.ICustomSummaryCalculator

        Private CustomValue As Double = 0
        Private CallingForm As ASFBASE0
        Private grd As UltraWinGrid.UltraGrid
        Private CustomType As String = ""

        Public Sub New(ByVal f As ASFBASE0, ByVal g As UltraWinGrid.UltraGrid, Optional c As String = "")
            CallingForm = f
            grd = g
            CustomType = c
        End Sub

        Private Sub BeginCustomSummary(ByVal summarySettings As UltraWinGrid.SummarySettings, ByVal rows As UltraWinGrid.RowsCollection) Implements UltraWinGrid.ICustomSummaryCalculator.BeginCustomSummary
            ' Begins the summary for the SummarySettings object passed in. Implementation of
            ' this method should reset any state variables used for calculating the summary.
            CustomValue = 0
        End Sub

        Private Sub AggregateCustomSummary(ByVal summarySettings As UltraWinGrid.SummarySettings, ByVal row As UltraWinGrid.UltraGridRow) Implements UltraWinGrid.ICustomSummaryCalculator.AggregateCustomSummary
            ' Here is where we process each row that gets passed in.
            ' Each row belongs to the current summary being processed
            CallingForm.CustomSummary_DataRows _
            (summarySettings, row, CustomValue, grd)
        End Sub

        Private Function EndCustomSummary(ByVal summarySettings As UltraWinGrid.SummarySettings, ByVal rows As UltraWinGrid.RowsCollection) Implements UltraWinGrid.ICustomSummaryCalculator.EndCustomSummary
            ' This gets called when the every row has been processed so here is where we
            ' would return the calculated summary value.
            If CustomType = "String" Then
                Return CallingForm.CustomStringSummary_End(summarySettings, rows, CustomValue, grd)
            Else
                Return CallingForm.CustomSummary_End(summarySettings, rows, CustomValue, grd)
            End If
        End Function
    End Class

    Private Sub ASFBASE0_BindingContextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.BindingContextChanged
        'Stop
    End Sub

    Public Sub Dropped_On(ByVal e As System.Windows.Forms.DragEventArgs)
        Dim ENTITY As Dropped_On_Entity = Dropped_On_Context()

        If ENTITY.TABLE_NAME <> "" Then
            Show_ASFATTA1(ENTITY, e)
        End If
    End Sub

    Public Sub Show_ASFATTA1(
    ByVal ENTITY As Dropped_On_Entity,
    ByVal e As System.Windows.Forms.DragEventArgs)

        ASCMAIN1.Progress("Now Attaching Files ...")
        Dim F As New ASFATTA1
        F.ENTITY = ENTITY
        F.eDND = e
        F.ShowDialog()
        F.Dispose()
        ASCMAIN1.Progress("")
    End Sub

    Public Overridable Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        E.TABLE_NAME = ""
        E.TABLE_KEY = ""
        E.read_only = True
        E.enabled = False

        Return E
    End Function

    Public Overridable Function Data_Export_Context() As Data_Export_Entity

        If ASCMAIN1.USER_SECURITY_CODEs.Contains("EX") Then
            Dim E As New Data_Export_Entity
            E.enabled = True
            ASTDATA1s.Clear()
            For Each T As DataTable In dst.Tables
                If Not T.TableName Like "AST*" Then
                    ASTDATA1s.Add(T.TableName, T.TableName)
                End If
            Next
            Return E
        End If

    End Function

    Public Overridable Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity

        E.TABLE_NAME = ""
        E.COLUMN_NAME = ""
        E.CODE_VALUE = ""
        E.DESC_VALUE = ""
        E.ATTACHMENT_NOTES = ""
        Return E
    End Function

    Public Overridable Function Events_Context() As Events_Entity

        Dim E As New Events_Entity

        E.TABLE_NAME = ""
        E.TABLE_KEY = ""
        E.read_only = True
        E.enabled = False

        Return E
    End Function

    Public Sub Show_Attachments()

        Dim E As Dropped_On_Entity = Dropped_On_Context()

        If E.TABLE_NAME <> "" Then
            Dim F As New ASFATTA1
            F.ENTITY = E
            F.ShowDialog()
            F.Dispose()

            Try
                Attachments_Changed()
            Catch ex As Exception

            End Try

        End If

    End Sub

    Public Sub Show_Log()

        If ASCMAIN1.ActiveForm.OwnedForms.Count <> 0 Then
            For Each frm As Form In ASCMAIN1.ActiveForm.OwnedForms
                If TypeOf (frm) Is ASFCONV1 Then
                    Exit Sub
                End If
            Next
        End If

        Dim E As Log_Entity = Log_Context()

        If E.TABLE_NAME <> "" And E.enabled Then
            Dim F As New ASFCONV1(ASCMAIN1.ActiveForm, E)
            ' NON-MODAL - NEED TO RESOLVE DONE BUTTON, AND FORMS DISAPPEARING BEHIND EACHOTHER
            'F.Owner = ASCMAIN1.ActiveForm
            'F.Show()

            F.ShowDialog()
            F.Show()
            F.Dispose()
        End If

    End Sub

    Public Sub Show_Data()

        If ASCMAIN1.ActiveForm.OwnedForms.Count <> 0 Then
            For Each frm As Form In ASCMAIN1.ActiveForm.OwnedForms
                If TypeOf (frm) Is ASFCONV1 Then
                    Exit Sub
                End If
            Next
        End If

        Dim E As Data_Export_Entity = Data_Export_Context()

        If E.enabled Then
            Dim F As New ASFDATA1(ASCMAIN1.ActiveForm) '(ASCMAIN1.ActiveForm, E)
            ' NON-MODAL - NEED TO RESOLVE DONE BUTTON, AND FORMS DISAPPEARING BEHIND EACHOTHER
            'F.Owner = ASCMAIN1.ActiveForm
            'F.Show()

            F.ShowDialog()
            F.Show()
            F.Dispose()
        End If

    End Sub

    Public Sub Show_Events()

        If ASCMAIN1.ActiveForm.OwnedForms.Count <> 0 Then
            For Each frm As Form In ASCMAIN1.ActiveForm.OwnedForms
                If TypeOf (frm) Is ASFCONV1 Then
                    Exit Sub
                End If
            Next
        End If

        Dim E As Events_Entity = Events_Context()

        If E.TABLE_NAME <> "" And E.enabled Then
            Dim F As New ASFEVNT1(ASCMAIN1.ActiveForm, E)
            ' NON-MODAL - NEED TO RESOLVE DONE BUTTON, AND FORMS DISAPPEARING BEHIND EACHOTHER
            'F.Owner = ASCMAIN1.ActiveForm
            'F.Show()

            F.ShowDialog()
            F.Show()
            F.Dispose()
        End If

    End Sub

    Public Overridable Function Audit_Context() As Audit_Entity

        Dim E As New Audit_Entity

        E.TABLE_NAME = ""
        E.TABLE_DESC = ""
        E.KEY_VALUE = ""
        E.KEY_DESC = ""

        Return E
    End Function

    Public Sub Show_Audit()

        Dim E As Audit_Entity = Audit_Context()

        If E.TABLE_NAME <> "" Then

            ' This code was lifted from ASFODEM Case "Audit Trail"
            Dim TABLE_NAMEs As New List(Of String)
            TABLE_NAMEs.Add(E.TABLE_NAME)

            ASCMAIN1.sql = "Select TABLE_NAME from ASTAUDT0 where FORM_NAME = '" & E.TABLE_NAME & "'"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                If Not TABLE_NAMEs.Contains(row.Item("TABLE_NAME") & "") Then
                    TABLE_NAMEs.Add(row.Item("TABLE_NAME") & "")
                End If
            Next

            Dim sqlTABLE_NAME As String = ""
            If TABLE_NAMEs.Count > 1 Then
                sqlTABLE_NAME = "TABLE_NAME, "
            End If

            Dim F As New ASFMSGBF
            F.grdGroupBy = True
            F.grdFilter = True
            '= '" & E.TABLE_NAME & "'"
            ASCMAIN1.sql = "Select " & sqlTABLE_NAME & IIf(E.KEY_VALUE = "", "KEY_VALUE, ", "") & "COLUMN_NAME, INIT_DATE, USER_ID, OLD_VALUE, NEW_VALUE, XNO, FM_MODE from ASTAUDT1 where TABLE_NAME in ('" & Join(TABLE_NAMEs.ToArray, "','") & "')" & IIf(E.KEY_VALUE = "", " and (1=1", " and (KEY_VALUE = '" & E.KEY_VALUE & "'") & " or KEY_VALUE like '" & E.KEY_VALUE & ":%') and COLUMN_NAME NOT IN ('INIT_DATE','LAST_DATE','INIT_OPER','LAST_OPER')"
            F.Show_grd(ASCDATA1.GetDataTable, ASCMAIN1.ActiveForm, "Audit Trail for Table: " & E.TABLE_NAME & " " & E.TABLE_DESC & IIf(E.KEY_VALUE = "", "", "; Record: " & E.KEY_VALUE & " " & E.KEY_DESC), "ASTAUDT1")
            F.Dispose()
            F = Nothing
        End If
    End Sub

    Private Sub tlb_AfterToolCloseup(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolDropdownEventArgs) Handles tlb.AfterToolCloseup

    End Sub

    Public Overridable Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)

        'If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
        '    e.Cancel = True
        '    Exit Sub
        'End If


        If Not GRDs.ContainsKey(Mid(e.Tool.Key, 4)) Then
            Exit Sub
        End If
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.Key, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Tag = "X"
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
            tlb_sbt.Tag = ""
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Tag = "X"
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
            tlb_sbt.Tag = ""
        End If
        If tlb_pop.Tools.Exists("Show Pins") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Pins"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not (grd.DisplayLayout.Override.FixedHeaderIndicator = UltraWinGrid.FixedHeaderIndicator.None)
        End If

    End Sub

    Public Overridable Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If e.Tool.OwningMenu Is Nothing Then Exit Sub

        If GRDs.ContainsKey(Mid(e.Tool.OwningMenu.Key, 4)) Then
            grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        End If

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag <> "X" Then
                    grd.Rows.ColumnFilters.ClearAllFilters()
                    Show_Filter(grd, tlb_sbt.Checked)
                End If

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag <> "X" Then
                    grd.DisplayLayout.Bands(0).SortedColumns.Clear()
                    grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
                End If

            Case "Show Pins"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    grd.DisplayLayout.Override.FixedHeaderIndicator = UltraWinGrid.FixedHeaderIndicator.Button
                Else
                    grd.DisplayLayout.Override.FixedHeaderIndicator = UltraWinGrid.FixedHeaderIndicator.None
                End If

            Case "Export to JSON"
                Dim dirJSON As String = "C:\JSON"
                dirJSON = ASCMAIN1.Folders("Temp") & "JSON\"
                If Not System.IO.Directory.Exists(dirJSON) Then
                    System.IO.Directory.CreateDirectory(dirJSON)
                End If
                Dim jsonFile As String = System.IO.Path.Combine(dirJSON, $"{grd.Name}.json")
                If System.IO.File.Exists(jsonFile) Then
                    System.IO.File.Delete(jsonFile)
                End If
                Dim totalColumnWidth As Double = 0
                For Each col As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
                    If Not col.Hidden Then
                        totalColumnWidth += col.Width
                    End If
                Next

                Using sw As New System.IO.StreamWriter(jsonFile)
                    sw.WriteLine($"case '{grd.Name}': {"{"}")
                    sw.WriteLine("return grdColumns = [" & vbCrLf)
                    For Each col As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
                        If Not col.Hidden Then
                            Dim objPfx As String = vbTab & "{"
                            Dim objSfx As String = vbTab & "}"
                            Dim columnClasses As New List(Of String)
                            Dim COLUMN_KEY As String = $"COLUMN_KEY: '{col.Key}',"
                            Dim COLUMN_HEADER As String = $"COLUMN_HEADER: '{col.Header.Caption}',"
                            Dim COLUMN_WIDTH_calc As Double = (col.Width / totalColumnWidth) * 100
                            Dim COLUMN_WIDTH As String = $"COLUMN_WIDTH: '{COLUMN_WIDTH_calc}%',"
                            Dim COLUMN_MASK As String = $"COLUMN_MASK: '',"
                            Dim DATA_TYPE As String = ""

                            columnClasses.Add("grd-column")

                            Select Case col.DataType.ToString
                                Case "System.DateTime"
                                    DATA_TYPE = $"DATA_TYPE: 'D',"
                                    columnClasses.Add("col-datetime")
                                Case "System.Int16", "System.Int32", "System.Int64"
                                    DATA_TYPE = $"DATA_TYPE: 'N',"
                                    columnClasses.Add("col-number")
                                    columnClasses.Add("col-integer")
                                    COLUMN_MASK = $"COLUMN_MASK: '###,##0',"
                                Case "System.Double", "System.Single", "System.Decimal"
                                    DATA_TYPE = $"DATA_TYPE: 'N',"
                                    columnClasses.Add("col-number")
                                    columnClasses.Add("col-non-integer")
                                    COLUMN_MASK = $"COLUMN_MASK: '###,##0.00',"
                                Case "System.String"
                                    DATA_TYPE = $"DATA_TYPE: 'V',"
                                    columnClasses.Add("col-string")
                                    If col.Format = "" Then
                                        If col.Key Like "*_PHONE_NO" _
                                        OrElse col.Key Like "*_TEL_NO" _
                                        OrElse col.Key Like "*_FAX_NO" _
                                        OrElse col.Key Like "*_PHONE" _
                                        OrElse col.Key Like "*_TEL" _
                                        OrElse col.Key Like "*_FAX" _
                                        OrElse col.Key Like "*_CELL" _
                                        OrElse col.Key = "PHONE" _
                                        OrElse col.Key = "CUST_SMS_NO" _
                                        OrElse col.Key = "SMS_TO" _
                                        OrElse col.Key = "SMS_FROM" Then
                                            COLUMN_MASK = "COLUMN_MASK: '(###) ###-####',"
                                            columnClasses.Add("col-phone")
                                        End If

                                    End If
                            End Select
                            Dim columnClass_str As String = String.Join(" ", columnClasses.ToArray())
                            Dim COLUMN_CLASSES As String = $"COLUMN_CLASSES: '{columnClass_str}',"
                            Dim colDef As String = $"{objPfx} " & vbCrLf _
                                & vbTab & vbTab & $"{COLUMN_KEY}" & vbCrLf _
                                & vbTab & vbTab & $"{COLUMN_HEADER}" & vbCrLf _
                                & vbTab & vbTab & $"{COLUMN_WIDTH}" & vbCrLf _
                                & vbTab & vbTab & $"{COLUMN_MASK}" & vbCrLf _
                                & vbTab & vbTab & $"{DATA_TYPE}" & vbCrLf _
                                & vbTab & vbTab & $"{COLUMN_CLASSES}" & vbCrLf _
                                & $"{objSfx},"
                            sw.WriteLine(colDef)
                        End If
                    Next
                    sw.WriteLine("]")
                    sw.WriteLine("}")
                    sw.Close()
                    sw.Dispose()
                End Using
                Process.Start("notepad.exe", jsonFile)



        End Select

        If grd Is Nothing Then
            Exit Sub
        Else
            If grd.ActiveRow Is Nothing Then
                Exit Sub
            Else
                If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
                    Exit Sub
                End If
            End If

        End If

        Select Case e.Tool.Key
            'Case "Customer Inquiry"
            '    Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
            '    Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")

        End Select
    End Sub

    Public Overridable Sub tlb_ToolValueChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolEventArgs)

    End Sub

#Region "grdSetup"
    Sub Clear_grdSetup(ByVal grdSetup As UltraWinGrid.UltraGrid, ByVal Clear_All As Boolean)
        grdSetup.UpdateData()
        grdSetup.ActiveRow = Nothing
        For Each dr As DataRow In DirectCast(grdSetup.DataSource, DataTable).Rows
            dr.Item("SEQUENCE") = DBNull.Value
            dr.Item("PAGE_BREAK") = "0"
            If Clear_All Then
                dr.Item("EXCLUDE") = "0"
                dr.Item("GROUP_ALL_OTHERS") = "0"
                dr.Item("CODE_VALUES") = ""
            End If
        Next

        Dim gsp As grdSetupParms = DirectCast(grdSetup.Tag, grdSetupParms)
        gsp.SEQs = 0
        Re_SEQ(grdSetup)

    End Sub

    Sub Setup_grdSetup(ByVal grdSetup As UltraWinGrid.UltraGrid,
                       Optional ByVal grdSetup_FORM_NAME As String = "")

        Dim gsp As New grdSetupParms
        gsp.SEQs = 0
        gsp.COLUMN_NAME_last = ""
        gsp.FORM_NAME = grdSetup_FORM_NAME
        If gsp.FORM_NAME = "" Then
            gsp.FORM_NAME = Me.Name
        End If
        grdSetup.Tag = gsp

        Dim tblASTDSQLA As DataTable = Create_tblASTDSQLA()
        grdSetup.DataSource = tblASTDSQLA

        With grdSetup
            With .DisplayLayout.Bands(0)
                .Columns("COLUMN_NAME").Hidden = True
                With .Columns("COLUMN_CAPTION")
                    .ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always

                    .CellAppearance.BackColor = Color.PaleGreen
                    .CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                    .CellButtonAppearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "Sort_Ascending")
                    .CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.CellSelect
                    .Header.Caption = "Code"
                    .Style = Infragistics.Win.UltraWinGrid.ColumnStyle.EditButton
                    .Width = 120
                End With
                With .Columns("CODE_VALUES")
                    .ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always
                    .CellButtonAppearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "arrow_up_blue")
                    .Header.Caption = "Values"
                    .Style = Infragistics.Win.UltraWinGrid.ColumnStyle.EditButton
                    .Width = 250
                End With
                With .Columns("EXCLUDE")
                    .CellAppearance.TextHAlignAsString = "Center"
                    .Header.Appearance.TextHAlignAsString = "Center"
                    .Header.Caption = "Excl"
                    .CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
                    .Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
                    .Width = 40
                End With
                With .Columns("SEQUENCE")
                    .AutoCompleteMode = Infragistics.Win.AutoCompleteMode.None
                    .CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly
                    .CellAppearance.TextHAlignAsString = "Center"
                    .Header.Appearance.TextHAlignAsString = "Center"
                    .Header.Caption = "Seq"
                    .Width = 40
                End With
                With .Columns("PAGE_BREAK")
                    .CellAppearance.TextHAlignAsString = "Center"
                    .Header.Appearance.TextHAlignAsString = "Center"
                    .Header.Caption = "Page"
                    .Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
                    .Width = 40
                End With
                .Columns("PAGE_BREAK").Hidden = True
                .Columns("SORTABLE").Hidden = True
                With .Columns("GROUP_ALL_OTHERS")
                    .CellAppearance.TextHAlignAsString = "Center"
                    .Header.Appearance.TextHAlignAsString = "Center"
                    .Header.Caption = "Grp"
                    .Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
                    .Width = 40
                End With
                .Columns("GROUP_ALL_OTHERS").Hidden = True
                .Columns("COLUMN_LAST").Hidden = True
            End With

        End With

        ASCMAIN1.grdInitializeLayout(grdSetup)

        grdSetup.DisplayLayout.GroupByBox.Hidden = True
        grdSetup.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns
        grdSetup.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.Select
        grdSetup.DisplayLayout.Override.RowSelectors = DefaultableBoolean.Default
        grdSetup.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        grdSetup.Text = "Click Sort Icon to Sort, Double-Click Seq Heading to Clear Sort"

        AddHandler grdSetup.AfterCellUpdate, AddressOf grdSetup_AfterCellUpdate
        AddHandler grdSetup.BeforeRowUpdate, AddressOf grdSetup_BeforeRowUpdate
        AddHandler grdSetup.ClickCellButton, AddressOf grdSetup_ClickCellButton
        AddHandler grdSetup.DoubleClickHeader, AddressOf grdSetup_DoubleClickHeader
        AddHandler grdSetup.InitializeRow, AddressOf grdSetup_InitializeRow
        AddHandler grdSetup.KeyDown, AddressOf grdSetup_KeyDown
        AddHandler grdSetup.KeyPress, AddressOf grdSetup_KeyPress
        AddHandler grdSetup.Leave, AddressOf grdSetup_Leave

        'Load_Popup_Menu(grdSetup, "BB", "Save As", "Load Settings")

        If Not ROWs.ContainsKey("GLTPARM1") Then
            Get_PARM("GLTPARM1")
        End If

        Dim COLUMN_CAPTION As String = ""
        For Each dr As DataRow In ASCDATA1.GetDataTable("Select ASTDSQLA.COLUMN_NAME, NVL(ASTDSQLA.COLUMN_CAPTION,ASTDSQLK.COLUMN_CAPTION) COLUMN_CAPTION, ASTDSQLA.SORTABLE, ASTDSQLA.COLUMN_LAST from ASTDSQLA,ASTDSQLK WHERE ASTDSQLK.COLUMN_NAME (+) = ASTDSQLA.COLUMN_NAME and ASTDSQLA.FORM_NAME = '" & gsp.FORM_NAME & "' ORDER BY NVL(ASTDSQLA.COLUMN_CAPTION,ASTDSQLK.COLUMN_CAPTION)").Rows
            If dr.Item("COLUMN_NAME") = "SEG2_CODE" And ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & "" = "" _
            Or dr.Item("COLUMN_NAME") = "SEG3_CODE" And ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & "" = "" _
            Or dr.Item("COLUMN_NAME") = "SEG4_CODE" And ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & "" = "" _
            Then
                ' SKIP IT
            Else
                COLUMN_CAPTION = dr.Item("COLUMN_CAPTION") & ""
                If dr.Item("COLUMN_NAME") = "SEG2_CODE" Then
                    COLUMN_CAPTION = ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC")
                ElseIf dr.Item("COLUMN_NAME") = "SEG3_CODE" Then
                    COLUMN_CAPTION = ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC")
                ElseIf dr.Item("COLUMN_NAME") = "SEG4_CODE" Then
                    COLUMN_CAPTION = ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC")
                End If
                If dr.Item("SORTABLE") & "" = "1" Or dr.Item("COLUMN_LAST") & "" = "1" Then
                    gsp.PB_Report = True
                End If
                If dr.Item("COLUMN_LAST") & "" = "1" Then
                    gsp.COLUMN_NAME_last = dr.Item("COLUMN_NAME")
                    dr.Item("SORTABLE") = "0"
                    'PB_Report = True ?
                End If
                Call Add_Row(tblASTDSQLA, COLUMN_CAPTION, dr.Item("COLUMN_NAME") & "", dr.Item("SORTABLE") & "")
            End If

        Next dr

        grdSetup.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChangeOrLostFocus

        gsp.SEQs = 0
        Re_SEQ(grdSetup)

        If grdSetup.Rows.Count > 0 Then
            grdSetup.ActiveRow = grdSetup.Rows(0)
        End If

        grdSetup.UpdateData()

        grdSetup.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdSetup.DisplayLayout.Bands(0).SortedColumns.Add("COLUMN_CAPTION", False)
    End Sub

    Function Create_tblASTDSQLA() As DataTable
        Dim tblASTDSQLA As New DataTable
        With tblASTDSQLA
            .Columns.Add("COLUMN_NAME")
            .Columns.Add("COLUMN_CAPTION")
            .Columns.Add("CODE_VALUES")
            .Columns.Add("EXCLUDE")
            .Columns("EXCLUDE").DefaultValue = "0"
            .Columns.Add("SEQUENCE", GetType(System.Int16))
            .Columns.Add("PAGE_BREAK")
            .Columns("PAGE_BREAK").DefaultValue = "0"
            .Columns.Add("SORTABLE")
            .Columns.Add("GROUP_ALL_OTHERS")
            .Columns.Add("COLUMN_LAST")
            .PrimaryKey = New DataColumn() {tblASTDSQLA.Columns("COLUMN_NAME")}
        End With
        Return tblASTDSQLA
    End Function

    Sub Add_Row(ByVal tblASTDSQLA As DataTable,
    ByVal COLUMN_CAPTION As String,
    ByVal COLUMN_NAME As String,
    ByVal SORTABLE As String)
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

    Sub Re_SEQ(ByVal grdSetup As UltraWinGrid.UltraGrid,
    Optional ByVal COLUMN_NAME As String = "",
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

        Dim gsp As grdSetupParms = DirectCast(grdSetup.Tag, grdSetupParms)

        If gsp.COLUMN_NAME_last <> "" Then
            row = tbl.Rows.Find(gsp.COLUMN_NAME_last)
            row.Item("SEQUENCE") = System.DBNull.Value
            row.Item("PAGE_BREAK") = "0"
        End If

        gsp.SEQs = 0
        For Each dr As DataRow In tbl.Select _
            ("SEQUENCE IS NOT NULL OR SEQUENCE <> ''", "SEQUENCE")
            gsp.SEQs += 1
            dr.Item("SEQUENCE") = gsp.SEQs
        Next

        If gsp.COLUMN_NAME_last <> "" Then
            row = tbl.Rows.Find(gsp.COLUMN_NAME_last)
            gsp.SEQs += 1
            row.Item("SEQUENCE") = gsp.SEQs
        End If

    End Sub

    Private Sub grdSetup_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs)
        If e.Cell.Column.Key <> "CODE_VALUES" Then
            Dim grdSetup As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
            grdSetup.UpdateData()
        End If
    End Sub

    Private Sub grdSetup_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs)
        Dim grdSetup As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)

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

    Private Sub grdSetup_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs)
        Dim grdSetup As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        If e.Cell.Column.Key = "COLUMN_CAPTION" Then
            If e.Cell.Row.Cells("SORTABLE").Text = "1" Then
                If e.Cell.Row.Cells("SEQUENCE").Text <> "" Then
                    Call Re_SEQ(grdSetup, e.Cell.Row.Cells("COLUMN_NAME").Text, False)
                Else
                    Call Re_SEQ(grdSetup, e.Cell.Row.Cells("COLUMN_NAME").Text, True)
                End If
            End If
        ElseIf e.Cell.Column.Key = "CODE_VALUES" Then
            ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL(grdSetup.ActiveRow.Cells("COLUMN_NAME").Text)
            If ASCMAIN1.CodeSelector.SQL <> "" Then
                ASCMAIN1.CodeSelector.MultipleSelections = True
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = Replace(grdSetup.ActiveRow.Cells("CODE_VALUES").Text & "", ",", Chr(0))
                Dim F As New ASFCODE1
                F.ShowDialog()
                F.Dispose()
                If ASCMAIN1.CodeSelector.Selections <> 0 Then
                    grdSetup.ActiveRow.Cells("CODE_VALUES").Value = Mid$(Replace(ASCMAIN1.CodeSelector.SelectedCodes0, Chr(0), ","), 2)
                    grdSetup.UpdateData()
                End If
            End If
        End If
    End Sub

    Private Sub grdSetup_DoubleClickHeader(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickHeaderEventArgs)
        Call Clear_grdSetup(sender, False)
    End Sub

    Private Sub grdSetup_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs)
        If e.Row.Cells("SORTABLE").Text <> "1" Then
            e.Row.Cells("COLUMN_CAPTION").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Edit
        End If
    End Sub

    Private Sub grdSetup_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim grdSetup As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)

        If e.KeyValue = Windows.Forms.Keys.Delete Then
            If grdSetup.ActiveCell.Column.Key = "SEQUENCE" Then
                If grdSetup.ActiveCell.Text <> "" Then
                    'grdSetup.ActiveCell.Value = DBNull.Value
                    'grdSetup.UpdateData()
                    Call Re_SEQ(grdSetup, grdSetup.ActiveRow.Cells("COLUMN_NAME").Text, False)
                End If
            End If
        End If

        If e.KeyValue = Windows.Forms.Keys.Enter Then
            If grdSetup.ActiveCell.Column.Key = "CODE_VALUES" Then
                grdSetup.Update()
            End If
        End If
    End Sub

    Private Sub grdSetup_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim grdSetup As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        Dim gsp As grdSetupParms = DirectCast(grdSetup.Tag, grdSetupParms)

        If grdSetup.ActiveCell IsNot Nothing Then
            If grdSetup.ActiveCell.Column.Key = "SEQUENCE" And grdSetup.ActiveRow.Cells("SORTABLE").Text = "1" Then
                Dim COLUMN_NAME As String = grdSetup.ActiveRow.Cells("COLUMN_NAME").Text
                Dim SEQcur As Integer = Val(grdSetup.ActiveCell.Text)
                Dim SEQnew As Integer = Val(e.KeyChar)
                If SEQnew < 1 Or SEQnew = SEQcur Or (SEQcur = 0 And SEQnew > gsp.SEQs + 1) Or (SEQcur <> 0 And SEQnew > gsp.SEQs) Then
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
                    gsp.SEQs += 1
                End If
            End If
        End If
    End Sub

    Private Sub grdSetup_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim grdSetup As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        grdSetup.UpdateData()
    End Sub
#End Region

    Overridable Sub Format_grd_ASFMSGBF(ByVal grd As UltraWinGrid.UltraGrid, ByVal grdCode As String)
        Select Case grdCode
            Case "ASTAUDT1"
                With grd.DisplayLayout.Override
                    .CellClickAction = UltraWinGrid.CellClickAction.CellSelect
                    .ActiveRowAppearance.BackColor = Drawing.Color.Empty
                    .ActiveRowAppearance.ForeColor = Drawing.Color.Empty
                End With
                With grd.DisplayLayout.Bands(0)
                    .Override.CellClickAction = UltraWinGrid.CellClickAction.CellSelect
                    ' .Columns("COLUMN_NAME").Hidden = True
                    With .Columns("COLUMN_NAME")
                        .Header.Caption = "Data Field"
                        .Width = 150
                    End With
                    With .Columns("USER_ID")
                        .Header.Caption = "By"
                        .Width = 80
                        '.Style = UltraWinGrid.ColumnStyle.CheckBox
                    End With
                    With .Columns("INIT_DATE")
                        .Format = "MM/dd/yy HH:mm"
                        .Width = 120
                        .Header.Caption = "Changed"
                    End With
                    With .Columns("OLD_VALUE")
                        .Header.Caption = "Old Value"
                        .Width = 100
                    End With
                    With .Columns("NEW_VALUE")
                        .Header.Caption = "New Value"
                        .Width = 100
                    End With
                    With .Columns("XNO")
                        .Header.Caption = "Session"
                        .Width = 100
                    End With
                    With .Columns("FM_MODE")
                        .Header.Caption = "Mode"
                        .Width = 50
                        .Header.Appearance.TextHAlign = HAlign.Center
                        .CellAppearance.TextHAlign = HAlign.Center
                    End With

                End With

        End Select
    End Sub


    Sub Process_DragDrop()

        If ENTITY.READ_ONLY Then
            Exit Sub
        End If

        Application.DoEvents()

        Dim files() As String = eDND.Data.GetData(DataFormats.FileDrop)

        If files IsNot Nothing Then
            For Each FILENAME As String In files
                Dim Msg As String = Attach_File(FILENAME)
                If Msg <> "" Then
                    MsgBox(Msg, MsgBoxStyle.OkOnly, "Error Attempting to Attach File ")
                End If
            Next
        Else
            Try
                Dim outlook As Microsoft.Office.Interop.Outlook.Application = CType(Microsoft.VisualBasic.Interaction.GetObject("", "Outlook.Application"), Microsoft.Office.Interop.Outlook.Application)
                Dim explorer As Microsoft.Office.Interop.Outlook.Explorer = outlook.ActiveExplorer

                For i As Int32 = 0 To explorer.Selection.Count - 1
                    Dim mail As Microsoft.Office.Interop.Outlook.MailItem = CType(explorer.Selection.Item(i + 1), Microsoft.Office.Interop.Outlook.MailItem)
                    mail.SaveAs(ASCMAIN1.Folders("Temp") & "mailitem.msg")

                    Dim FILENAME As String = ASCMAIN1.Folders("Temp") & "mailitem.msg"
                    Dim Msg As String = Attach_File(FILENAME, mail.Subject, mail.SenderName, mail.SentOn)
                    If Msg <> "" Then
                        MsgBox(Msg, MsgBoxStyle.OkOnly, "Error Attempting to Attach File ")
                    End If
                    mail = Nothing
                Next

                outlook = Nothing
                explorer = Nothing

            Catch ex As System.Exception

                MsgBox(ex, "Error - Outlook request not found")

            End Try

        End If

        Application.DoEvents()

    End Sub

    Function Attach_File(
    ByVal FILENAME As String,
    Optional ByVal ATTACHMENT_DESC As String = "",
    Optional ByVal ATTACHMENT_ORIGINATOR As String = "",
    Optional ByVal ATTACHMENT_DATETIME As Date = Nothing,
    Optional ByVal Update_Immediately As Boolean = True)

        Dim LAST_DATE As Date = Now
        If EXT_allowed Is Nothing Then
            ASCMAIN1.sql = "Select * from ASTATTA1"
            EXT_allowed = New List(Of String)
            For Each rowASTATTA1 As DataRow In ASCDATA1.GetDataTable.Rows
                EXT_allowed.Add(rowASTATTA1.Item("ATTACHMENT_EXT"))
            Next
        End If

        Dim Msg As String = ""

        Try
            Dim FILENAME_SEGMENTS() As String = Split(FILENAME, ".")
            Dim FILENAME_EXT As String = FILENAME_SEGMENTS(UBound(FILENAME_SEGMENTS)).ToUpper
            If Not EXT_allowed.Contains(FILENAME_EXT) Then
                Msg = "Unsupported File Type (" & FILENAME_EXT & ") for file " & FILENAME
            Else
                Dim rowASTATTA2 As DataRow = dst.Tables("ASTATTA2").NewRow
                rowASTATTA2.Item("TABLE_NAME") = ENTITY.TABLE_NAME
                rowASTATTA2.Item("COLUMN_NAME") = ENTITY.COLUMN_NAME
                rowASTATTA2.Item("CODE_VALUE") = ENTITY.CODE_VALUE
                rowASTATTA2.Item("ATTACHMENT_DESC") = ATTACHMENT_DESC
                rowASTATTA2.Item("ATTACHMENT_FILENAME") = FILENAME
                rowASTATTA2.Item("ATTACHMENT_EXT") = FILENAME_EXT.ToUpper
                rowASTATTA2.Item("COMPUTER_NAME") = ASCMAIN1.COMPUTER_NAME
                rowASTATTA2.Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                rowASTATTA2.Item("INIT_DATE") = LAST_DATE
                rowASTATTA2.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowASTATTA2.Item("LAST_DATE") = LAST_DATE
                rowASTATTA2.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowASTATTA2.Item("ATTACHMENT_TYPE") = ""

                Dim FF As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)

                If ATTACHMENT_DESC = "" Then
                    rowASTATTA2.Item("ATTACHMENT_DESC") = FF.Name
                End If

                rowASTATTA2.Item("ATTACHMENT_ORIGINATOR") = ATTACHMENT_ORIGINATOR

                If ATTACHMENT_DATETIME.ToString = "1/1/0001 12:00:00 AM" Then
                    rowASTATTA2.Item("ATTACHMENT_DATETIME") = FF.LastWriteTime
                Else
                    rowASTATTA2.Item("ATTACHMENT_DATETIME") = ATTACHMENT_DATETIME
                End If


                Dim FolderName As String = ASCMAIN1.Folders("Temp")
                If Update_Immediately Then
                    FolderName = ASCMAIN1.Folders("Attach")
                End If

                Dim ATTACHMENT_NO As String = ""
                Do
                    ATTACHMENT_NO = ASCMAIN1.Next_Control_No("ASTATTA2.ATTACHMENT_NO")
                Loop While My.Computer.FileSystem.FileExists(FolderName & ATTACHMENT_NO)

                My.Computer.FileSystem.CopyFile(FILENAME, FolderName & ATTACHMENT_NO)

                rowASTATTA2.Item("ATTACHMENT_NO") = ATTACHMENT_NO

                dst.Tables("ASTATTA2").Rows.Add(rowASTATTA2)

                ' Me.Name = "ASFATTA1" Then ' If ASCMAIN1.ActiveForm.Name = "ASFATTA1" Then
                If Update_Immediately Then
                    TDAs("ASTATTA2").Update(dst.Tables("ASTATTA2"))
                End If
                ENTITY.ATTACHMENT_NO = ATTACHMENT_NO

            End If
        Catch ex As Exception
            Msg = ex.Message
        End Try

        Return Msg
    End Function

    Function Create_Relation(
    ByVal TABLE_NAME_parent As String,
    ByVal TABLE_NAME_child As String,
    ByVal COLUMN_NAMEsParent As String,
    Optional ByVal COLUMN_NAMEsChild As String = "")

        Dim R As DataRelation = ASCDATA1.GetRelation _
        (dst, TABLE_NAME_parent, TABLE_NAME_child, COLUMN_NAMEsParent, COLUMN_NAMEsChild)

        dst.Relations.Add(R)

        Return R
    End Function

    Sub Show_RowState(ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs, ByVal TABLE_NAME As String)
        If e.Row.IsAddRow Then
            e.Row.RowSelectorAppearance.BackColor = System.Drawing.Color.LightGreen
            e.Row.RowSelectorAppearance.BackColor2 = System.Drawing.Color.Green
            Exit Sub
        End If

        Dim i As Int32 = e.Row.ListIndex
        Select Case dst.Tables(TABLE_NAME).Rows(i).RowState
            Case DataRowState.Added
                e.Row.RowSelectorAppearance.BackColor = System.Drawing.Color.LightGreen
                e.Row.RowSelectorAppearance.BackColor2 = System.Drawing.Color.Green
            Case DataRowState.Modified
                e.Row.RowSelectorAppearance.BackColor = System.Drawing.Color.LightSkyBlue
                e.Row.RowSelectorAppearance.BackColor2 = System.Drawing.Color.Blue
        End Select
    End Sub

    Sub Enlist_Transaction()
        'TDAs(TABLE_NAME).InsertCommand.Transaction = ASCMAIN1.T
        'TDAs(TABLE_NAME).UpdateCommand.Transaction = ASCMAIN1.T
        'TDAs(TABLE_NAME).DeleteCommand.Transaction = ASCMAIN1.T
    End Sub

    Sub ReParent_Tabs(ByVal tab As UltraWinTabControl.UltraTabControl)
        With tab
            For Each ctl As Control In .Controls
                If TypeOf (ctl) Is UltraWinTabControl.UltraTabSharedControlsPage Then
                Else
                    For Each ctl2 As Control In ctl.Controls
                        ctl2.Parent = .Parent
                    Next
                End If
            Next
            .Visible = False
        End With
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="grd"></param>
    ''' <param name="ColPosition"></param>
    ''' <param name="allow_update">G = Follow Grid, I = Use InquiryMode, Y = Yes always, N = No always</param>
    ''' <param name="TABLE_NAME"></param>
    ''' <param name="COLUMN_NAME"></param>
    ''' <remarks></remarks>
    Sub Add_Attachment_Column(
    ByVal grd As UltraWinGrid.UltraGrid,
    ByVal ColPosition As Int16,
    Optional ByVal allow_update As String = "G",
    Optional ByVal TABLE_NAME As String = "",
    Optional ByVal COLUMN_NAME As String = "")

        Dim imgA As System.Drawing.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "paperclip")

        With grd.DisplayLayout.Bands(0).Columns.Add("ATTACHMENTS")
            .Width = grd.DisplayLayout.Bands(0).RowSelectorWidthResolved
            .Style = UltraWinGrid.ColumnStyle.Button
            .Header.Appearance.Image = imgA
            .Header.Appearance.ImageHAlign = HAlign.Center
            .CellAppearance.TextHAlign = HAlign.Center
            .Header.VisiblePosition = ColPosition
            .Header.Caption = ""
            .ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
        End With

        If TABLE_NAME = "" Then
            TABLE_NAME = grd.DisplayLayout.Bands(0).Key
        End If
        If COLUMN_NAME = "" Then
            COLUMN_NAME = dst.Tables(TABLE_NAME).Columns(0).ColumnName
        End If

        Dim AB As New Attachment_Button
        AB.TABLE_NAME = TABLE_NAME
        AB.COLUMN_NAME = COLUMN_NAME
        AB.allow_update = allow_update
        AB.grd = grd
        grds_with_Attachments.Add(grd.DisplayLayout.Bands(0).Key, AB)

    End Sub

    Private Sub UltraGridExcelExporter1_CellExported(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ExcelExport.CellExportedEventArgs) Handles UltraGridExcelExporter1.CellExported
        If e.GridColumn.Style = UltraWinGrid.ColumnStyle.Image Then
            'e.CurrentWorksheet.Rows(e.CurrentRowIndex).Cells(e.CurrentColumnIndex).Value = "x"
            If e.Value Is Nothing Then
                Exit Sub
            End If
            'Dim image As Image = image.FromFile("C:\Images\Image1.bmp")
            Dim image As Image = e.Value

            Dim imageShape As Infragistics.Documents.Excel.WorksheetImage =
              New Infragistics.Documents.Excel.WorksheetImage(image)

            Dim cellA1 As Infragistics.Documents.Excel.WorksheetCell =
               e.CurrentWorksheet.Rows.Item(e.CurrentRowIndex).Cells.Item(e.CurrentColumnIndex)
            Dim cellA2 As Infragistics.Documents.Excel.WorksheetCell =
               e.CurrentWorksheet.Rows.Item(e.CurrentRowIndex + 1).Cells.Item(e.CurrentColumnIndex)
            'Dim cellA1 As Infragistics.Documents.Excel.WorksheetCell = _
            '   e.CurrentWorksheet.Rows.Item(0).Cells.Item(0)

            ' The top-left corner of the image should be at the 
            ' top-left corner of cell A1
            imageShape.PositioningMode = Infragistics.Documents.Excel.ShapePositioningMode.MoveAndSizeWithCells

            imageShape.TopLeftCornerCell = cellA1
            imageShape.TopLeftCornerPosition = New PointF(0.0F, 0.0F)

            ' The bottom-right corner of the image should be at 
            ' the bottom-right corner of cell A1
            'imageShape.BottomRightCornerCell = cellA1
            'imageShape.BottomRightCornerPosition = New PointF(100.0F, 100.0F)
            imageShape.BottomRightCornerCell = cellA2
            imageShape.BottomRightCornerPosition = New PointF(100.0F, 0.0F)

            e.CurrentWorksheet.Shapes.Add(imageShape)
        End If
    End Sub

    Private Sub UltraGridExcelExporter1_CellExporting(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ExcelExport.CellExportingEventArgs) Handles UltraGridExcelExporter1.CellExporting
        'Dim sCellType As String
        'Dim sCellContents As String
        'sCellType = e.Value.GetType().FullName

        'If sCellType = "System.String" Then
        '    sCellContents = e.Value
        '    If sCellContents.StartsWith("Accounting Manager") = True Then
        '        e.Workbook.WindowOptions.SelectedWorksheet.Rows(e.CurrentRowIndex).Cells(e.CurrentColumnIndex).Value = "AM"
        '    End If
        'End If

    End Sub

    Private Sub UltraGridExcelExporter1_InitializeSummary(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ExcelExport.InitializeSummaryEventArgs) Handles UltraGridExcelExporter1.InitializeSummary
        If e.Summary.DisplayFormat <> "" Then
            e.ExcelFormatStr = Replace(Replace(e.Summary.DisplayFormat, "{0:", ""), "}", "")
        End If
    End Sub

    Private Sub UltraGridExcelExporter1_SummaryCellExporting(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ExcelExport.SummaryCellExportingEventArgs) Handles UltraGridExcelExporter1.SummaryCellExporting

    End Sub

    Private Sub UltraGridExcelExporter1_SummaryCellExported(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ExcelExport.SummaryCellExportedEventArgs) Handles UltraGridExcelExporter1.SummaryCellExported

    End Sub

    Private Sub UltraGridExcelExporter1_InitializeColumn(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ExcelExport.InitializeColumnEventArgs) Handles UltraGridExcelExporter1.InitializeColumn
        If e.Column.Format <> "" Then
            e.ExcelFormatStr = e.Column.Format
        ElseIf e.Column.DataType.Name = "DateTime" Then
            e.ExcelFormatStr = "mm/dd/yyyy"
        End If

    End Sub

    Private Sub UltraGridExcelExporter1_FormulaExported(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ExcelExport.FormulaExportedEventArgs) Handles UltraGridExcelExporter1.FormulaExported

    End Sub

    Private Sub UltraGridExcelExporter1_FormulaExporting(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ExcelExport.FormulaExportingEventArgs) Handles UltraGridExcelExporter1.FormulaExporting

    End Sub

    Private Sub UltraGridExcelExporter1_HeaderRowExported(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ExcelExport.HeaderRowExportedEventArgs) Handles UltraGridExcelExporter1.HeaderRowExported

    End Sub

    Private Sub UltraGridExcelExporter1_HeaderRowExporting(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ExcelExport.HeaderRowExportingEventArgs) Handles UltraGridExcelExporter1.HeaderRowExporting
        If e.Band.Tag = "" Then
            e.Band.Tag = CStr(e.CurrentRowIndex)
        Else
            e.Cancel = True
        End If
    End Sub

    Public Sub Spell_Check(ByVal COLUMN_NAME As String)
        Absx1.txtFor(COLUMN_NAME).SpellChecker = ASFMAIN1.UltraSpellChecker1
    End Sub

    Public Sub Spell_Check(ByVal COLUMN_NAMEs() As String)
        For Each COLUMN_NAME As String In COLUMN_NAMEs
            Spell_Check(COLUMN_NAME)
        Next
    End Sub


    Sub Create_Outlook_mailitem(
    ByVal MAIL_TO As String,
    ByVal MAIL_CC As String,
    ByVal MAIL_SUBJECT As String,
    ByVal MAIL_BODY As String,
    Optional ByVal MAIL_ATTACHMENTS() As String = Nothing)

        Dim outlook As Microsoft.Office.Interop.Outlook.Application = CType(Microsoft.VisualBasic.Interaction.GetObject("", "Outlook.Application"), Microsoft.Office.Interop.Outlook.Application)
        'Dim explorer As Microsoft.Office.Interop.Outlook.Explorer = outlook.ActiveExplorer
        Dim mail As Microsoft.Office.Interop.Outlook.MailItem ' = CType(explorer.Selection.Item(i + 1), Microsoft.Office.Interop.Outlook.MailItem)

        mail = outlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)

        mail.To = MAIL_TO
        mail.CC = MAIL_CC
        mail.Subject = MAIL_SUBJECT
        mail.Body = MAIL_BODY

        If MAIL_ATTACHMENTS IsNot Nothing Then
            For Each MAIL_ATTACHMENT As String In MAIL_ATTACHMENTS
                mail.Attachments.Add(MAIL_ATTACHMENT)
            Next
        End If

        mail.Save()

    End Sub

    Sub Synch_TABLE_NAME(ByVal TABLE_NAME As String)
        If dst.Tables.Contains(TABLE_NAME) Then
            If Me.BindingContext.Contains(dst.Tables(TABLE_NAME)) Then
                Dim X As CurrencyManager = Me.BindingContext(dst.Tables(TABLE_NAME))
                Try
                    X.EndCurrentEdit()
                Catch ex As Exception

                End Try
            End If
        End If
    End Sub

    Function Gembox_Import_Sheet_to_DataTable(
                                             ByVal COLs As Integer,
                                             Optional ByRef FILENAME As String = "") As DataTable

        Dim dataTable As DataTable = Nothing

        ' Dim FILENAME As String = ""
        FILENAME = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"

            openFileDialog1.RestoreDirectory = True

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        If FILENAME <> "" Then

            SpreadsheetInfo.SetLicense(ASCMAIN1.GemboxKey)

            Dim ef As ExcelFile = New ExcelFile
            If FILENAME.ToLower.EndsWith("xls") Then
                ef.LoadXls(FILENAME, XlsOptions.PreserveAll)
                ' ExcelFile.Load(FILENAME, New GemBox.Spreadsheet.XlsLoadOptions With {.PreserveOptions = XlsOptions.PreserveAll})
            Else
                ef.LoadXlsx(FILENAME, XlsxOptions.PreserveMakeCopy)
                ' ExcelFile.Load(FILENAME, New GemBox.Spreadsheet.XlsLoadOptions)
            End If

            dataTable = New DataTable

            ' Depending on the format of the input file, you need to change this:

            If COLs = 0 Then
                COLs = ef.Worksheets(0).Columns.Count
            End If

            For I As Integer = 1 To COLs
                dataTable.Columns.Add("COL" & Format(I, "000"), GetType(String))
            Next

            ' Select the first worksheet from the file.
            Dim ws As ExcelWorksheet = ef.Worksheets(0)

            ' Manage ExtractDataError.WrongType error.
            ' GemBox.Spreadsheet doesn't automatically convert numbers to strings in ExtractToDataTable() method because of culture issues; 
            ' someone would expect the number 12.4 as "12.4" and someone else as "12,4".
            ' In this case we'll skip such row.
            AddHandler ws.ExtractDataEvent,
            Function(sender As Object, e As ExtractDataDelegateEventArgs) e.Action = ExtractDataEventAction.SkipRow

            ' Extract the data from the worksheet to the DataTable.
            ' Data is extracted starting at first row and first column for 10 rows or until the first empty row appears.
            For r As Integer = 1 To ef.Worksheets(0).Rows.Count
                Dim row As DataRow = dataTable.NewRow
                For c As Integer = 1 To COLs
                    row.Item(c - 1) = ef.Worksheets(0).Rows(r - 1).Cells(c - 1).Value
                Next
                dataTable.Rows.Add(row)
            Next
            'ws.ExtractToDataTable(dataTable, ef.Worksheets(0).Rows.Count, ExtractDataOptions.SkipEmptyRows, ws.Rows(0), ws.Columns(0))
        End If

        Return dataTable
    End Function

    Function Create_Temporary_Table(TABLE_NAME As String, keyCOLUMN_NAMEs As String) As String

        ' code to demonstrate that DDL within a BeginTans actually commits to a savepoint and restarts the transaction
        'If ASCMAIN1.Running_in_VS Then
        '    Stop
        '    ASCMAIN1.sql = "UPDATE ICTSTAT2 SET WHSE_QTY_ON_HAND = WHSE_QTY_ON_HAND - 1000  WHERE ITEM_CODE = 'JLB306010SF'"
        '    ASCDATA1.ExecuteSQL()
        'End If

        ASCMAIN1.sql = "Select * from " & TABLE_NAME & " where ROWNUM < 1"
        Dim TABLE_NAME_temp As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & TABLE_NAME_temp & " Add Primary Key (" & keyCOLUMN_NAMEs & ")")

        ' code to demonstrate that DDL within a BeginTans actually commits to a savepoint and restarts the transaction
        'If ASCMAIN1.Running_in_VS Then
        '    Stop
        '    ASCMAIN1.sql = "UPDATE ICTSTAT2 SET WHSE_QTY_ON_HAND = WHSE_QTY_ON_HAND - 1000  WHERE ITEM_CODE = 'JLB306010SF'"
        '    ASCDATA1.ExecuteSQL()
        'End If


        ASCMAIN1.sql = "Select * from " & TABLE_NAME_temp & " where ROWNUM < 1"
        Create_TDA(dst.Tables.Add(TABLE_NAME), TABLE_NAME_temp, "**", 0)

        Return TABLE_NAME_temp
    End Function

    Sub Add_Codes(
        grd As UltraWinGrid.UltraGrid,
        TABLE_NAME As String,
        COLUMN_NAME As String,
        Codes_Caption As String)

        Dim TABLE_NAME_grid As String = DirectCast(grd.DataSource, DataTable).TableName
        Dim sql_where As String = Get_List_of_Codes(TABLE_NAME & "." & COLUMN_NAME & " not in", TABLE_NAME_grid, COLUMN_NAME)
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL(COLUMN_NAME, , sql_where)

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading " & Codes_Caption)

                grd.Visible = False
                If grd.ActiveRow IsNot Nothing Then grd.ActiveRow.CancelUpdate()
                Try
                    For Each CODE As String In ASCMAIN1.CodeSelector.SelectedCodes
                        grd.ActiveRow = grd.DisplayLayout.Bands(0).AddNew
                        grd.ActiveRow.Cells(COLUMN_NAME).Value = CODE
                        grd.ActiveRow.Update()
                    Next
                Catch ex As Exception

                End Try

                grd.Visible = True
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
            End If
        End If
    End Sub

    Function Get_List_of_Codes(
        ByVal sql_where_clause As String,
        TABLE_NAME As String,
        COLUMN_NAME As String,
        Optional filter As String = "") As String

        Dim sql_where As String = ""
        Dim CODEs As String = ""
        For Each row As DataRow In dst.Tables(TABLE_NAME).Select(filter)
            CODEs &= ",'" & row.Item(COLUMN_NAME) & "'"
        Next
        If CODEs <> "" Then
            sql_where = sql_where_clause & " (" & Mid(CODEs, 2) & ")"
        End If
        Return sql_where
    End Function

    Sub Gembox_Excel_Export(grds() As UltraWinGrid.UltraGrid)
        Me.Cursor = Cursors.WaitCursor
        SpreadsheetInfo.SetLicense(ASCMAIN1.GemboxKey)
        Dim myWorkbook As New GemBox.Spreadsheet.ExcelFile

        If dst.Tables.Contains("ASTGRIDC") Then
            dst.Tables("ASTGRIDC").Rows.Clear()
        Else
            With dst.Tables.Add("ASTGRIDC")
                .Columns.Add("SHEET", GetType(System.Int32))
                .Columns.Add("ROW", GetType(System.Int32))
                .Columns.Add("COL", GetType(System.Int32))
                .Columns.Add("COLOR1", GetType(System.Int64))
                .Columns.Add("COLOR2", GetType(System.Int64))
                .Columns.Add("GRADIENT", GetType(System.Int32))
            End With
        End If

        For Each grd As UltraWinGrid.UltraGrid In grds
            Gembox_Export_to_Excel_Add_grd(myWorkbook, grd, False, grd.Text)
        Next
        Gembox_Export_to_Excel_Show(myWorkbook, Me.Text)

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Function Get_SelectCommand(TABLE_NAME As String) As String
        Return TDAs(TABLE_NAME).SelectCommand.CommandText
    End Function

    Sub Set_SelectCommand(TABLE_NAME As String, sql As String)
        TDAs(TABLE_NAME).SelectCommand.CommandText = sql
    End Sub

    'Sub Update_Record_TDA_Rows(TABLE_NAME As String)
    '    TDAs(TABLE_NAME).Update(dst.Tables(TABLE_NAME))
    'End Sub

    Function sqlCleanParameter(sql As String) As String
        Return Replace(sql, "'", "")
    End Function

    Function Get_Code(VIEW_NAME As String, Optional TABLE_NAME As String = "",
                      Optional sql_where As String = "",
                      Optional key_value As String = "",
                      Optional TABLE_NAME_temp As String = "") As String

        Dim CODE_VALUE As String = ""
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL(VIEW_NAME, TABLE_NAME,
                                                                  sql_where, key_value, TABLE_NAME_temp)
        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = False
            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
            Using F As New ASFCODE1
                F.ShowDialog()
            End Using
            CODE_VALUE = ASCMAIN1.CodeSelector.SelectedCode
        End If

        Return CODE_VALUE
    End Function

    Function Get_Codes(VIEW_NAME As String, Optional TABLE_NAME As String = "",
                      Optional sql_where As String = "",
                      Optional key_value As String = "",
                      Optional TABLE_NAME_temp As String = "") As List(Of String)

        Dim CODE_VALUEs As New List(Of String)
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL(VIEW_NAME, TABLE_NAME,
                                                                  sql_where, key_value, TABLE_NAME_temp)
        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = False
            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
            Using F As New ASFCODE1
                F.ShowDialog()
            End Using
            For Each CODE_VALUE As String In ASCMAIN1.CodeSelector.SelectedCodes
                CODE_VALUEs.Add(CODE_VALUE)
            Next
        End If

        Return CODE_VALUEs
    End Function

    Sub grd_Appearance_LightGray(grd As UltraWinGrid.UltraGrid)
        For Each gcol As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
            With gcol.Header.Appearance
                .BackColor = Color.White
                .BackColor2 = Color.LightGray
                .BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            End With
        Next
    End Sub
    ''' <summary>
    ''' Returns an expression representing an Excel Cell 
    ''' </summary>
    ''' <param name="Row">1-based Row, 1 = Row 1, 0 = Entire Column</param>
    ''' <param name="Col">1-based Column, 1 = Column A, 0 = Entire Row</param>
    ''' <param name="ABSOLUTE">0 = Nothing, 1 = Absolute Column, 2 = Absolute Row, 3 = Absolute Cell</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Excel_Cell(Row As Integer, Col As Integer, Optional ABSOLUTE As Integer = 0) As String

        Dim c As String = Chr(((Col - 1) Mod 26) + 65)
        If Int((Col - 1) / 26) > 0 Then
            c = Chr(Int((Col - 1) / 26) + 64) & c
        End If

        Dim z1 As String = ""
        Dim z2 As String = ""

        If Row > 0 Then
            If ABSOLUTE = 1 Then
                z1 = "$"
            End If
            If ABSOLUTE = 2 Then
                z2 = "$"
            End If
            If ABSOLUTE = 3 Then
                z1 = "$"
                z2 = "$"
            End If
        End If

        Excel_Cell = z1 & c & z2 & IIf(Row > 0, CStr(Row), "")
    End Function
    ''' <summary>
    ''' Returns an expression representing an Excel Cell
    ''' This is a 0-based version of the Excel_Cell function
    ''' </summary>
    ''' <param name="Row">0-based Row, 0 = Row 1, -1 = Entire Column</param>
    ''' <param name="Col">0-based Column, 0 = Column A, -1 = Entire Row</param>
    ''' <param name="ABSOLUTE">0 = Nothing, 1 = Absolute Row, 2 = Absolute Column, 3 = Absolute Cell</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Excel_Cell0(Row As Integer, Col As Integer, Optional ABSOLUTE As Integer = 0) As String
        Return Excel_Cell(Row + 1, Col + 1, ABSOLUTE)
    End Function


    ''' <summary>
    ''' Loads a DataTable into a SpreadsheetGear Range
    ''' You can send in a DataTable (tbl), or send in Nothing as tbl to use ASCMAIN1.sql
    ''' You can send in a grd to use for column formatting, or send in a grdfmt collection
    ''' Use grd or grdfmt (never both, grd trumps grdfmt) to format columns, and to pull a subset of the columns from the table
    ''' Note that grd.DataSource is not used if tbl is Nothing - send in DirectCast(grd.DataSource,DataTable) as tbl to use grd.DataSource
    ''' 
    ''' To set up grdfmt:
    ''' Dim grdfmt as New Dictionary(Of String, grdcolfmt)
    ''' grdfmt.Add(gcol.Key, New grdcolfmt With {.Key = gcol.Key, .Caption = gcol.Header.Caption, .Format = gcol.Format, .BackColor2 = gcol.Header.Appearance.BackColor2, .Index = 0, .Width = gcol.Width, .TextHAlign = gcol.CellAppearance.TextHAlign})
    ''' or
    ''' grdfmt_add(grdfmt, "CUST_STORE_NO", grdSOTORDF1)
    ''' </summary>
    ''' <param name="Rx">1-based Row where the DataTable should be placed</param>
    ''' <param name="Cx">1-based Column where the DataTable should be placed</param>
    ''' <param name="tbl">The DataTable to be loaded, or Nothing if using ASCMAIN1.sql</param>
    ''' <param name="oSheet">The SpreadsheetGear Worksheet to load the DataTable into</param>
    ''' <param name="grd">grd containing formatting information for columns</param>
    ''' <param name="grdfmt">grdfmt collection containing formatting information for columns</param>
    ''' <param name="OrderBy">The Sort to use to sequence the rows in tbl</param>
    ''' <param name="WhereClause">The clause to use to filter the rows in tbl, no need for the keyword "where": "CUST_STATE = '' and CUST_STATUS = 'A'"</param>
    ''' <param name="clear_headings">True to clear out the headings after loading DataTable</param>
    ''' <returns>This Function Returns the DataTable after Loading it into the SpreadsheetGear Worksheet</returns>
    ''' <remarks></remarks>
    Function Load_DataTable_into_SGXLS(
        Rx As Int64,
        Cx As Int64,
        tbl As DataTable,
        oSheet As SpreadsheetGear.IWorksheet,
        Optional grd As UltraWinGrid.UltraGrid = Nothing,
        Optional grdfmt As Dictionary(Of String, grdcolfmt) = Nothing,
        Optional OrderBy As String = "",
        Optional WhereClause As String = "",
        Optional clear_headings As Boolean = False) As DataTable

        Dim range As SpreadsheetGear.IRange = Nothing

        If tbl Is Nothing Then tbl = ASCDATA1.GetDataTable

        Dim tbl3 As DataTable = Nothing

        If tbl.ParentRelations.Count > 0 Then
            tbl3 = tbl
        Else
            tbl3 = tbl.Copy
        End If

        Dim dvw As DataView = tbl.DefaultView
        Dim dvw3 As DataView = tbl3.DefaultView

        If OrderBy <> "" Then
            dvw3.Sort = OrderBy
        Else
            dvw3.Sort = dvw.Sort
        End If
        If WhereClause <> "" Then
            dvw3.RowFilter = WhereClause
        Else
            dvw3.RowFilter = dvw.RowFilter
        End If
        tbl3 = dvw3.ToTable

        'Dim COLS As New Dictionary(Of String, Integer)


        For Each dc As DataColumn In tbl3.Columns
            If dc.ColumnName.StartsWith("'") And dc.ColumnName.EndsWith("'") Then
                dc.ColumnName = Mid(dc.ColumnName, 2, dc.ColumnName.Length - 2)
            End If
        Next

        If grd IsNot Nothing Then
            Dim visible_positions As New Dictionary(Of Integer, String)
            Dim visible_position_max As Integer = 0

            For i As Integer = 0 To grd.DisplayLayout.Bands(0).Columns.Count - 1
                Dim gcol As UltraWinGrid.UltraGridColumn = grd.DisplayLayout.Bands(0).Columns(i)
                Dim c As String = gcol.Key
                'COLS.Add(c, i)
                Dim visible_position = gcol.Header.VisiblePosition
                visible_positions.Add(visible_position, c)
                If visible_position > visible_position_max Then
                    visible_position_max = visible_position
                End If
            Next

            grdfmt = New Dictionary(Of String, grdcolfmt)

            For i As Integer = 0 To visible_position_max
                Dim c As String = visible_positions(i)
                Dim gcol As UltraWinGrid.UltraGridColumn = grd.DisplayLayout.Bands(0).Columns(c)
                If Not gcol.Hidden Then
                    grdfmt.Add(gcol.Key, New grdcolfmt _
                        With {.Key = gcol.Key,
                              .Caption = gcol.Header.Caption,
                              .Format = gcol.Format,
                              .BackColor2 = gcol.Header.Appearance.BackColor2,
                              .Index = 0,
                              .Width = gcol.Width,
                              .TextHAlign = gcol.CellAppearance.TextHAlign})
                End If
            Next
        End If

        If grdfmt IsNot Nothing Then
            Dim tbl2 As New DataTable
            Dim ci As Integer = -1
            For Each c As String In grdfmt.Keys
                tbl2.Columns.Add(c, tbl3.Columns(c).DataType)
                tbl2.Columns(c).Caption = grdfmt(c).Caption
                Dim g As grdcolfmt = grdfmt(c)
                ci += 1 : g.Index = ci
            Next

            For Each row As DataRow In tbl3.Select("")
                Dim row2 As DataRow = tbl2.NewRow
                For Each c As DataColumn In tbl2.Columns
                    row2.Item(c.ColumnName) = row.Item(c.ColumnName)
                Next
                tbl2.Rows.Add(row2)
            Next
            tbl3 = tbl2
        End If

        For c As Integer = 0 To tbl3.Columns.Count - 1
            range = oSheet.Range(Excel_Cell(Rx + 1, Cx + c) & ":" & Excel_Cell(Rx + tbl3.Rows.Count, Cx + c))
            Dim COLUMN_NAME As String = tbl3.Columns(c).ColumnName
            If grdfmt IsNot Nothing AndAlso grdfmt.ContainsKey(COLUMN_NAME) Then
                'Dim i As Integer = COLS(COLUMN_NAME)
                Dim i As Integer = grdfmt(COLUMN_NAME).Index
                Dim g As grdcolfmt = grdfmt(COLUMN_NAME)
                If g.Format IsNot Nothing Then
                    range.NumberFormat = g.Format
                Else
                    'range.NumberFormat = "Text"
                    range.NumberFormat = "@"
                End If
                range.ColumnWidth = g.Width / 10
                range.HorizontalAlignment = IIf(g.TextHAlign = HAlign.Center, SpreadsheetGear.HAlign.Center,
                                                         IIf(g.TextHAlign = HAlign.Right, SpreadsheetGear.HAlign.Right, SpreadsheetGear.HAlign.Left))

            End If
        Next

        range = oSheet.Range(Excel_Cell(Rx, Cx))
        range.CopyFromDataTable(tbl3, SpreadsheetGear.Data.SetDataFlags.None)
        If clear_headings Then
            range = oSheet.Range(Excel_Cell(Rx + 0, Cx + 0) & ":" & Excel_Cell(Rx + 0, Cx + tbl3.Columns.Count - 1))
            oSheet.Range(Excel_Cell(Rx, Cx)).EntireRow.Clear()
        Else
            For c As Integer = 0 To tbl3.Columns.Count - 1

                Dim COLUMN_NAME As String = tbl3.Columns(c).ColumnName
                If grdfmt IsNot Nothing AndAlso grdfmt.ContainsKey(COLUMN_NAME) Then
                    Dim i As Integer = grdfmt(COLUMN_NAME).Index
                    Dim g As grdcolfmt = grdfmt(COLUMN_NAME)
                    ' With oSheet.Cells(Rx + 0, Cx + c)
                    ' sg is 0 based, and excel automation is 1 based
                    With oSheet.Cells(Excel_Cell(Rx + 0, Cx + c))
                        .Value = g.Caption
                        .Interior.Color = SpreadsheetGear.Drawing.Color.GetSpreadsheetGearColor(g.BackColor2)
                        .HorizontalAlignment = IIf(g.TextHAlign = HAlign.Center, SpreadsheetGear.HAlign.Center,
                                               IIf(g.TextHAlign = HAlign.Right, SpreadsheetGear.HAlign.Right, SpreadsheetGear.HAlign.Left))

                    End With
                End If
            Next
        End If

        Return tbl3
    End Function

    Public Sub grdfmt_add(grdfmt As Dictionary(Of String, grdcolfmt), COLUMN_NAME As String, GRD As UltraWinGrid.UltraGrid)
        Dim gcol As UltraWinGrid.UltraGridColumn = GRD.DisplayLayout.Bands(0).Columns(COLUMN_NAME)

        grdfmt.Add(gcol.Key, New grdcolfmt _
            With {.Key = gcol.Key,
                  .Caption = gcol.Header.Caption,
                  .Format = gcol.Format,
                  .BackColor2 = gcol.Header.Appearance.BackColor2,
                  .Index = 0,
                  .Width = gcol.Width,
                  .TextHAlign = gcol.CellAppearance.TextHAlign})
    End Sub
    Private Sub grd_Error(sender As Object, e As UltraWinGrid.ErrorEventArgs)
        grdError = e
    End Sub

    Sub ReleaseCOMObject(reference As Object)
        Exit Sub
        Try
            Do While (System.Runtime.InteropServices.Marshal.ReleaseComObject(reference) <= 0)
                reference = Nothing
            Loop

        Catch ex As Exception
            reference = Nothing
        Finally
            GC.Collect()
        End Try

    End Sub

    Structure grdcolfmt
        Dim Key As String
        Dim Caption As String
        Dim Format As String
        Dim BackColor2 As System.Drawing.Color
        Dim Index As Integer
        Dim Width As Integer
        Dim TextHAlign As Infragistics.Win.HAlign
    End Structure

    Function Add_Document_to_ASTSPRF1(FILENAME_ORIG As String,
                                      Optional Show_Document_after_Archiving As Boolean = True,
                                      Optional Add_Document_to_Files_to_Publish As Boolean = True) As DataRow

        Dim FILETYPE As String = ""

        Dim fi As System.IO.FileInfo

        Try
            fi = My.Computer.FileSystem.GetFileInfo(FILENAME_ORIG)
            FILETYPE = Mid(fi.Extension, 2).ToUpper

            If LookUp("ASTATTA1", FILETYPE) Is Nothing Then
                MsgBox("Cannot Include " & FILENAME_ORIG & " as a Publishable File", MsgBoxStyle.OkOnly, "Unsupported File Type (" & fi.Extension & ")")
                Return Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.InnerException.Message, MsgBoxStyle.OkOnly, "Error gettting File Information")
            Return Nothing
        End Try


        If Not dst.Tables.Contains("ASTSPRF1") Then
            Create_TDA(dst.Tables.Add, "ASTSPRF1", "*")
        End If

        Dim REPORT_NO As String = ASCMAIN1.Next_Control_No("ASTSPRF1.REPORT_NO")

        Dim rowASTSPRF1 As DataRow = dst.Tables("ASTSPRF1").NewRow
        With rowASTSPRF1
            .Item("REPORT_NO") = REPORT_NO
            .Item("FORM_NAME") = MENU_ITEM_OBJECT ' FORM_NAME
            .Item("XNO") = XNO
            .Item("USER_ID") = ASCMAIN1.USER_ID
            .Item("YYYYPP") = ASCMAIN1.CYP
            .Item("YP_LEGEND") = ASCMAIN1.Get_Legend(ASCMAIN1.CYP)
            .Item("RPT_TITLE") = MENU_ITEM_DESC
            .Item("RPT") = ""
            .Item("COMPUTER_NAME") = ASCMAIN1.COMPUTER_NAME
            .Item("REPORT_DATE") = Now + ASCMAIN1.NowTSD
            .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
            .Item("MENU_ITEM_OBJECT") = MENU_ITEM_OBJECT
            .Item("MENU_ITEM_TYPE") = MENU_ITEM_TYPE
            .Item("MENU_ID") = MENU_ID
            .Item("MENU_ITEM_SECURITY") = MENU_ITEM_SECURITY
            .Item("VERSION_NO") = ASCMAIN1.VERSION_NO

            Dim ok_to_archive_report As String = ""

            If ASCMAIN1.Running_in_VS Then
                If ASCMAIN1.DBS_SERVER = "" Then
                    ok_to_archive_report = "N"
                End If
                If ok_to_archive_report = "" Then
                    If ASCMAIN1.JOB_STREAM_CODE.Length = 0 AndAlso MsgBoxResult.No = MsgBox("Copy File to Archive?", MsgBoxStyle.YesNo, "You are running a Report on a Development Machine") Then
                        ok_to_archive_report = "N"
                    Else
                        ok_to_archive_report = "Y"
                    End If
                End If
            Else
                ok_to_archive_report = "Y"
            End If

            Dim FILENAME As String = ASCMAIN1.DBS_COMPANY & "_" & REPORT_NO & "." & FILETYPE

            If ASCMAIN1.JOB_STREAM_CODE.Length <> 0 Then
                ASCMAIN1.JOB_STREAM_REPORT_FILES.Add(FILENAME, FILENAME_ORIG)
            End If


            If ok_to_archive_report = "Y" Then
                My.Computer.FileSystem.CopyFile(
                    FILENAME_ORIG,
                    ASCMAIN1.Folders("Archive") & "Reports\" & Mid(FILENAME, 1, 3) & "\" & Mid(FILENAME, 5, 5) & "\" & FILENAME)
            End If

            .Item("FILENAME") = FILENAME
            .Item("FILETYPE") = FILETYPE
            .Item("FILENAME_ORIG") = FILENAME_ORIG
        End With

        dst.Tables("ASTSPRF1").Rows.Add(rowASTSPRF1)

        Update_Record_TDA("ASTSPRF1")

        If Add_Document_to_Files_to_Publish Then
            FILENAMEs_to_Publish.Add(FILENAME_ORIG)
        End If

        If Show_Document_after_Archiving Then
            Show_Document(FILENAME_ORIG)
        End If

        Return rowASTSPRF1

    End Function

    Sub Register_Dependency(TABLE_NAME As String)
        'GRANT CHANGE NOTIFICATION TO INT;
        'SELECT * FROM USER_CQ_NOTIFICATION_QUERIES;
        'update sptcwrxx set cwrx_date = cwrx_date + 1;
        'COMMIT;

        Dim ada As OracleDataAdapter = TDAs(TABLE_NAME)

        Dim dep As OracleDependency = New OracleDependency(ada.SelectCommand)
        With ada.SelectCommand
            .AddRowid = True
            .Notification.IsNotifiedOnce = False
            AddHandler dep.OnChange, AddressOf dep_OnChange
        End With

        oraDeps.Add(TABLE_NAME, dep)

    End Sub
    Sub dep_OnChange(sender As Object, eventArgs As OracleNotificationEventArgs)
        Stop
    End Sub

    Sub Set_DX_Column(grd As UltraWinGrid.UltraGrid,
                      COLUMN_NAME As String,
                      Optional COLUMN_CAPTION As String = "",
                      Optional WIDTH As Decimal = -1,
                      Optional FORMAT As String = "",
                      Optional summary_type As String = "",
                      Optional color As System.Drawing.Color = Nothing,
                      Optional band As Integer = 0)

        ' passing in grd with "" for COLUMN_NAME initializes the grid - ie, it makes all columns hidden

        If COLUMN_NAME = "" Then
            ASCMAIN1.grdInitializeLayout(grd, Me)
            grd.DisplayLayout.Bands(band).SortedColumns.Clear()

            For Each gcol As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(band).Columns
                gcol.Hidden = True
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackColor2 = Color.LightGray

                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

            Next
            grd.DisplayLayout.Bands(band).Summaries.Clear()
            grd.DisplayLayout.CaptionVisible = DefaultableBoolean.Default
            grd.DisplayLayout.CaptionAppearance.TextHAlign = HAlign.Left

            If Not tlb.Tools.Exists(grd.Name) Then
                Load_Popup_Menu(grd, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
            End If

            grd.DisplayLayout.Bands(band).ColumnFilters.ClearAllFilters()

            'Try
            '    Load_Popup_Menu(grd, "SSS", "Show Filter", "Show GroupBox", "Show Pins")

            'Catch ex As Exception

            'End Try
            Exit Sub
        End If

        Dim max_VisiblePosition As Integer = -1
        For Each gcol As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(band).Columns
            If Not gcol.Hidden Then
                If gcol.Header.VisiblePosition > max_VisiblePosition Then
                    max_VisiblePosition = gcol.Header.VisiblePosition
                End If
            End If
        Next

        With grd.DisplayLayout.Bands(band).Columns(COLUMN_NAME)

            If COLUMN_CAPTION <> "" Then .Header.Caption = COLUMN_CAPTION

            .Header.VisiblePosition = max_VisiblePosition + 1
            .Hidden = False
            If WIDTH = 0 Then
                .Hidden = True
            ElseIf WIDTH <> -1 Then
                .Width = WIDTH
            End If

            If FORMAT <> "" Then
                .Format = FORMAT
            End If

            If summary_type <> "" Then
                Create_Summary(grd, COLUMN_NAME, summary_type, , FORMAT)
            End If

            If color <> System.Drawing.Color.Empty Then
                .Header.Appearance.BackColor2 = color
            End If
        End With
    End Sub

    Function Add_DX_Grid(dxTabs As Infragistics.Win.UltraWinTabControl.UltraTabControl, Optional dxCaption As String = "Data Export", Optional grdName As String = "") As UltraWinGrid.UltraGrid
        Dim utp As New Infragistics.Win.UltraWinTabControl.UltraTabPageControl
        Dim grd As New Infragistics.Win.UltraWinGrid.UltraGrid
        Dim dxKey As String = IIf(grdName <> "", grdName, "grd_DX")
        Dim removeExisting As Boolean = False

        For t As Integer = 0 To dxTabs.Tabs.Count - 1
            If dxTabs.Tabs(t).Key = dxKey Then
                removeExisting = True
            End If
        Next
        If removeExisting Then
            dxTabs.Tabs.Remove(dxTabs.Tabs(dxKey))
        End If

        Dim i As Integer = dxTabs.Tabs.Count
        dxTabs.Controls.Add(utp)
        dxTabs.Tabs(i).Text = dxCaption
        dxTabs.Tabs(i).Key = dxKey
        utp.Controls.Add(grd)

        grd.Text = dxCaption
        grd.Dock = System.Windows.Forms.DockStyle.Fill
        grd.Name = dxKey

        grd.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.SortMulti
        grd.DisplayLayout.Override.RowSelectors = DefaultableBoolean.Default
        grd.DisplayLayout.Override.RowSelectorHeaderStyle = UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        grd.DisplayLayout.GroupByBox.Hidden = False
        grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True
        grd.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
        grd.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
        grd.DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.OutlookGroupBy
        grd.DisplayLayout.Override.FilterClearButtonLocation = UltraWinGrid.FilterClearButtonLocation.Row
        grd.DisplayLayout.Override.FilterUIType = UltraWinGrid.FilterUIType.FilterRow
        grd.DisplayLayout.Override.FilterRowAppearance.BackColor = Color.AliceBlue
        grd.DisplayLayout.Override.FilterOperatorLocation = UltraWinGrid.FilterOperatorLocation.AboveOperand
        grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        grd.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.Default

        ASCMAIN1.grdInitializeLayout(grd)

        'Load_Popup_Menu(grd, "SSBB", "Show Filter", "Show GroupBox", "Toggle Mask", "Add Total")

        Initialize_Controls_for_a_Container(utp)

        Return grd
    End Function
    Function Validate_Accounts_and_Segments_EMsg _
        (dt As DataTable, Optional Automated_JE As Boolean = False) As String

        Dim JE_ERRORS As List(Of String) = Validate_Accounts_and_Segments(dt, Automated_JE)

        If JE_ERRORS.Count = 0 Then
            Return ""
        Else
            Dim EMsg As String = ""
            For Each JE_ERROR_X As String In JE_ERRORS
                'ACCT_CODE & vbTab & ACCT_SEG_ID & vbTab & ACCT_SEG_CODE & vbTab & ERROR_TEXT
                Dim JE_ERROR() As String = Split(JE_ERROR_X, vbTab)
                EMsg &= vbCr & JE_ERROR(3) & ": " & JE_ERROR(0) & IIf(JE_ERROR(1) = "", "", "-" & JE_ERROR(1) & " " & JE_ERROR(2))
            Next
            Return EMsg

        End If
    End Function

    Function Validate_Accounts_and_Segments _
        (dt As DataTable, Optional Automated_JE As Boolean = False) As List(Of String)

        Dim ERROR_TEXT As String = ""
        Dim JE_ERRORS As New List(Of String)

        For Each rowACCT_CODE As DataRow In ASCMAIN1.Distinct_Values("", dt, "ACCT_CODE").Rows

            Dim ACCT_CODE As String = rowACCT_CODE.Item("ACCT_CODE") & ""
            Dim ACCT_SEG_ID As String = ""
            Dim ACCT_SEG_CODE As String = ""
            Dim rowGLTACCT1 As DataRow = LookUp("GLTACCT1", ACCT_CODE)
            If rowGLTACCT1 Is Nothing Then
                ERROR_TEXT = "Invalid Account Code"
                JE_ERRORS.Add(ACCT_CODE & vbTab & ACCT_SEG_ID & vbTab & ACCT_SEG_CODE & vbTab & ERROR_TEXT)
            Else
                If rowGLTACCT1.Item("ACCT_STATUS") & "" <> "A" Then
                    ERROR_TEXT = "Account Status is not Active"
                    JE_ERRORS.Add(ACCT_CODE & vbTab & ACCT_SEG_ID & vbTab & ACCT_SEG_CODE & vbTab & ERROR_TEXT)
                ElseIf rowGLTACCT1.Item("ACCT_SUB_CTL") & "" = "1" And Not Automated_JE Then
                    ERROR_TEXT = "Acct Code is a Control Account"
                    JE_ERRORS.Add(ACCT_CODE & vbTab & ACCT_SEG_ID & vbTab & ACCT_SEG_CODE & vbTab & ERROR_TEXT)
                Else
                    For i As Integer = 2 To 4
                        ACCT_SEG_ID = CStr(i)
                        Dim COLUMN_NAME As String = "SEG" & CStr(i) & "_CODE"
                        Dim ACCT_SEG_DEFAULT As String = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i))
                        For Each rowACCT_SEG_CODE As DataRow In ASCMAIN1.Distinct_Values("", "ACCT_CODE = '" & ACCT_CODE & "'", dt, COLUMN_NAME).Rows
                            ACCT_SEG_CODE = rowACCT_SEG_CODE.Item(COLUMN_NAME) & ""
                            Dim ACCT_SEG_TYPE As String = ROWs("GLTPARM1").Item("GL_PARM_SEG" & CStr(i) & "_DESC") & ""
                            If ACCT_SEG_TYPE = "" Then
                                ACCT_SEG_TYPE = "Segment " & ACCT_SEG_ID
                            End If
                            cdr = LookUp("GLTSEGM1", New String() {ACCT_SEG_ID, ACCT_SEG_CODE})
                            If cdr Is Nothing Then
                                ERROR_TEXT = "Invalid " & ACCT_SEG_TYPE & " Code"
                                JE_ERRORS.Add(ACCT_CODE & vbTab & ACCT_SEG_ID & vbTab & ACCT_SEG_CODE & vbTab & ERROR_TEXT)
                            Else
                                If cdr.Item("ACCT_SEG_NO_GL") & "" = "1" Then
                                    ERROR_TEXT = ACCT_SEG_TYPE & " Code not Permitted for J/E"
                                    JE_ERRORS.Add(ACCT_CODE & vbTab & ACCT_SEG_ID & vbTab & ACCT_SEG_CODE & vbTab & ERROR_TEXT)
                                End If
                                If cdr.Item("ACCT_SEG_STATUS") & "" <> "A" Then
                                    ERROR_TEXT = ACCT_SEG_TYPE & " Code not Active"
                                    JE_ERRORS.Add(ACCT_CODE & vbTab & ACCT_SEG_ID & vbTab & ACCT_SEG_CODE & vbTab & ERROR_TEXT)
                                End If

                                If ROWs("GLTPARM1").Item("GL_PARM_MAND_SEG_CTL") & "" = "1" Then
                                    ' 0 or D = Requires Default
                                    ' 1 or N = Requires Non-Default
                                    ' 2 OR A = Any Value
                                    ' ELIMINATING OLD 0/1/2 HAVING CHANGED GLTACCT1 AND ASTCODE1, AND TO MAKE WAY FOR NEW VALUES

                                    If rowGLTACCT1.Item("ACCT_SEG" & CStr(i) & "_MAND") & "" = "D" Then
                                        If ACCT_SEG_CODE <> ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i)) Then
                                            ERROR_TEXT = "Acct " & ACCT_CODE & " requires Default Value (" & ACCT_SEG_DEFAULT & ") for " & ACCT_SEG_TYPE
                                            JE_ERRORS.Add(ACCT_CODE & vbTab & ACCT_SEG_ID & vbTab & ACCT_SEG_CODE & vbTab & ERROR_TEXT)
                                        End If
                                    End If
                                    If rowGLTACCT1.Item("ACCT_SEG" & CStr(i) & "_MAND") & "" = "N" Then
                                        If ACCT_SEG_CODE = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i)) Then
                                            ERROR_TEXT = "Acct " & ACCT_CODE & " requires non-Default Value (" & ACCT_SEG_DEFAULT & ") for " & ACCT_SEG_TYPE
                                            JE_ERRORS.Add(ACCT_CODE & vbTab & ACCT_SEG_ID & vbTab & ACCT_SEG_CODE & vbTab & ERROR_TEXT)
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    Next
                End If
            End If
        Next

        Return JE_ERRORS

    End Function

    Public Sub MakeTransparent(chk As ABSCS.ABSCheckBox)
        With chk
            .Appearance.ForeColor = System.Drawing.Color.White
            .Appearance.BackColor = System.Drawing.Color.FromArgb(98, 160, 232)
            .Appearance.BackColor2 = System.Drawing.Color.FromArgb(83, 115, 191)
            .Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        End With
    End Sub

    Public Sub MakeTransparent(chk As Infragistics.Win.UltraWinEditors.UltraCheckEditor)
        With chk
            .Appearance.ForeColor = System.Drawing.Color.White
            .Appearance.BackColor = System.Drawing.Color.FromArgb(98, 160, 232)
            .Appearance.BackColor2 = System.Drawing.Color.FromArgb(83, 115, 191)
            .Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        End With
    End Sub

    Public Sub MakeTransparent(lbl As Infragistics.Win.Misc.UltraLabel)
        With lbl
            .Appearance.ForeColor = System.Drawing.Color.White
            .Appearance.BackColor = System.Drawing.Color.FromArgb(98, 160, 232)
            .Appearance.BackColor2 = System.Drawing.Color.FromArgb(83, 115, 191)
            .Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        End With
    End Sub
    Overridable Sub Export_rows_to_SQL_custom()
        ' this method is called from the Developer Options menu
        ' it is overridden in ASFCODEM to provide standard support in SRM
        ' for other forms, if not overriden, nothing happens - you must do an override such as the following
        'Export_rows_to_SQL("ASTVIEW1,ASTVIEW2,ASTVIEW3,ASTVIEW5", "where VIEW_NAME = '" & VIEW_NAME & "' and TABLE_NAME = '" & TABLE_NAME & "'")
    End Sub

    Overridable Sub Export_rows_to_SQL(
        TABLE_NAMEs As String,
        KEY_COLUMNs As String,
        sql_where As String,
        VIEW_NAME As String,
        TABLE_NAME_view As String)

        If ScreenMode Then
            Export_row_to_SQL(TABLE_NAMEs, sql_where)
        Else
            ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL(VIEW_NAME, TABLE_NAME_view)

            If ASCMAIN1.CodeSelector.SQL <> "" Then
                ASCMAIN1.CodeSelector.MultipleSelections = True
                Dim F As New ASFCODE1
                F.ShowDialog()
                F.Dispose()
                If ASCMAIN1.CodeSelector.Selections <> 0 Then

                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Now Exporting to SQL")

                    Dim S As String = ""

                    For Each rowCode As DataRow In ASCMAIN1.CodeSelector.SelectedRows
                        Dim sqlw As String = ""
                        For Each COLUMN_NAME_key As String In Split(KEY_COLUMNs, ",")
                            sqlw &= " and " & COLUMN_NAME_key & " = '" & rowCode.Item(COLUMN_NAME_key) & "'"
                        Next
                        sqlw = ASCMAIN1.SQL_Add_WHERE(sqlw)
                        '"where VIEW_NAME = '" & VIEW_NAME & "' and TABLE_NAME = '" & TABLE_NAME & "'"

                        If sql_where <> "" Then
                            If Trim(sql_where).ToLower.StartsWith("where") Then
                                sqlw = sql_where
                            Else
                                sqlw &= sql_where
                            End If
                        End If

                        S &= Export_row_to_SQL(TABLE_NAMEs, sqlw)

                    Next

                    Using frm As New ASFMSGBF
                        frm.Width = Me.Width * 0.8
                        frm.Height = Me.Height * 0.8
                        frm.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
                        frm.Get_txtblock_from_User("Insert Statements", "Selected Rows", S, True)
                    End Using

                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("")

                End If
            End If
        End If
    End Sub

    Overridable Function Export_row_to_SQL(TABLE_NAMEs As String, sql_where As String) As String
        Dim S As String = ""
        For Each TABLE_NAME As String In Split(TABLE_NAMEs, ",")
            ASCMAIN1.sql = "Select * from " & TABLE_NAME & " " & sql_where
            For Each row2 As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, TABLE_NAME).Select("")
                S &= Create_Insert_SQL_from_row(row2)
            Next
        Next
        Return S
    End Function

    Overridable Function Create_Insert_SQL_from_row(row As DataRow) As String
        ' don't know if we should include the column names in the insert statement
        Dim sql As String = ""
        Dim tbl As DataTable = row.Table
        Dim COLUMN_NAMEs As String = ""
        For Each dc As DataColumn In tbl.Columns
            Dim COLUMN_NAME As String = dc.ColumnName
            COLUMN_NAMEs &= "," & COLUMN_NAME
            If row.Item(COLUMN_NAME) & "" = "" Then
                sql &= ", NULL"
            Else
                Select Case dc.DataType.ToString
                    Case "System.String"
                        sql &= ",'" & Replace(row.Item(COLUMN_NAME), "'", "''") & "'"
                    Case "System.Int16", "System.Int32", "System.Int64", "System.Decimal", "System.Double"
                        sql &= "," & row.Item(COLUMN_NAME)
                    Case "System.Date"
                        sql &= ",'" & Format(row.Item(COLUMN_NAME), "dd-MMM-yyyy HH:mm") & "'"
                    Case Else
                        Stop ' unhandled datatype
                End Select
            End If
        Next

        sql = "Insert into " & tbl.TableName & " (" & Mid(COLUMN_NAMEs, 2) & ")" & " Values (" & Mid(sql, 2) & ");" & vbCrLf
        Return sql
    End Function
End Class

''' <summary>
''' This CreationFilter class will create a CheckBoxUIElement in each Header
''' in the grid that has a Style = Checkbox. It will fire the 
''' HeaderCheckBoxClicked event whenever the CheckBox is clicked.
''' Note that in order to maintain the CheckState, this CreationFilter uses the 
''' Tag proprty of the Header. So if the program uses the tag for something else
''' this will not work. 
''' </summary>
''' <remarks></remarks>
Public Class CheckBoxOnHeader_CreationFilter
    ' Implements the CreationFilter interface
    Implements IUIElementCreationFilter

    Public ColumnNames As List(Of String) = New List(Of String)

    ' This event will fire when the CheckBox is clicked. 
    Public Event HeaderCheckBoxClicked(ByVal sender As Object, ByVal e As HeaderCheckBoxEventArgs)

    Public Sub AfterCreateChildElements(ByVal parent As Infragistics.Win.UIElement) Implements Infragistics.Win.IUIElementCreationFilter.AfterCreateChildElements
        ' Check for the HeaderUIElement
        If TypeOf parent Is UltraWinGrid.HeaderUIElement Then
            ' Get the actual ColumnHeader that the HeaderUIElement is attached to
            Dim aColHeader As Infragistics.Win.UltraWinGrid.ColumnHeader
            aColHeader = CType(parent, UltraWinGrid.HeaderUIElement).Header

            ' Only put the Checkbox in the header of the ComboBox, AllowEdit
            If aColHeader.Column.Style = UltraWinGrid.ColumnStyle.CheckBox _
                    AndAlso (ColumnNames.Count = 0 OrElse ColumnNames.Contains(aColHeader.Column.Key)) _
                    AndAlso aColHeader.Column.CellActivation = UltraWinGrid.Activation.AllowEdit Then
                Dim aTextUIElement As TextUIElement
                Dim aCheckBoxUIElement As CheckBoxUIElement

                ' Since the grid sometimes re-uses UIElements, we need to check to make sure 
                ' the header does not already have a CheckBoxUIElement attached to it.
                ' If it does, we just get a reference to the existing CheckBoxUIElement,
                ' and reset its properties.
                aCheckBoxUIElement = parent.GetDescendant(GetType(CheckBoxUIElement))

                If aCheckBoxUIElement Is Nothing Then
                    ' Create a New CheckBoxUIElement
                    aCheckBoxUIElement = New CheckBoxUIElement(parent)
                End If

                ' Get the TextUIElement - this is where the text for the 
                ' Header is displayed. We need this so we can push it to the right
                ' in order to make room for the CheckBox
                aTextUIElement = CType(parent.GetDescendant(GetType(TextUIElement)), TextUIElement)

                ' Sanity check
                If aTextUIElement Is Nothing Then Exit Sub

                ' Get the Header and see if the Tag has been set. I the Tag is 
                ' set, we will assume it's the stored CheckState. This has to be
                ' done in order to maintain the CheckState when the grid repaints and
                ' UIElement are destroyed and recreated. 
                Dim aHeader As Infragistics.Win.UltraWinGrid.ColumnHeader = CType(aCheckBoxUIElement.GetAncestor(GetType(UltraWinGrid.HeaderUIElement)).GetContext(GetType(Infragistics.Win.UltraWinGrid.ColumnHeader)), Infragistics.Win.UltraWinGrid.ColumnHeader)

                If aHeader.Tag Is Nothing Then
                    ' If the tag was nothing, this is probably the first time this 
                    ' HeaderRow is being displayed, so default to Unchecked
                    aHeader.Tag = CheckState.Unchecked
                Else
                    aCheckBoxUIElement.CheckState = CType(aHeader.Tag, CheckState)
                End If

                ' Hook the ElementClick of the CheckBoxUIElement
                AddHandler aCheckBoxUIElement.ElementClick, AddressOf aCheckBoxUIElement_ElementClick

                ' Add the CheckBoxUIElement to the HeaderUIElement
                parent.ChildElements.Add(aCheckBoxUIElement)

                ' Position the CheckBoxUIElement. The number 3 here is used for 3
                ' pixels of padding between the CheckBox and the side of the header
                ' The CheckBox is shifted down slightly so it is centered in the header
                aCheckBoxUIElement.Rect = New Rectangle(parent.Rect.X + 3, parent.Rect.Y + ((parent.Rect.Height - aCheckBoxUIElement.CheckSize.Height) / 2), aCheckBoxUIElement.CheckSize.Width, aCheckBoxUIElement.CheckSize.Height)

                ' Push the TextUIElement to the right a little to make 
                ' room for the CheckBox. 3 pixels of padding are used again. 
                aTextUIElement.Rect = New Rectangle(aCheckBoxUIElement.Rect.Right + 3, aTextUIElement.Rect.Y, parent.Rect.Width - (aCheckBoxUIElement.Rect.Right - parent.Rect.X), aTextUIElement.Rect.Height)
            Else
                ' If the column is not a boolean column, we do not want to have a checkbox in it
                ' Since UIElements can be reused by the grid, there is a chance that one of the
                ' HeaderUIElements that we added a checkbox to for a boolean column header
                ' will be reused in a column that is not boolean.  In this case, we must remove
                ' the checkbox so that it will not appear in an inappropriate column header.
                Dim aCheckBoxUIElement As CheckBoxUIElement
                aCheckBoxUIElement = parent.GetDescendant(GetType(CheckBoxUIElement))

                If Not aCheckBoxUIElement Is Nothing Then
                    parent.ChildElements.Remove(aCheckBoxUIElement)
                    aCheckBoxUIElement.Dispose()
                End If
            End If
        End If
    End Sub

    Public Function BeforeCreateChildElements(ByVal parent As Infragistics.Win.UIElement) As Boolean Implements Infragistics.Win.IUIElementCreationFilter.BeforeCreateChildElements
        ' Don't need to do anything here.
        Return False
    End Function

    Private Sub aCheckBoxUIElement_ElementClick(ByVal sender As Object, ByVal e As Infragistics.Win.UIElementEventArgs)
        ' Get the CheckBoxUIElement that was clicked
        Dim aCheckBoxUIElement As CheckBoxUIElement = CType(e.Element, CheckBoxUIElement)

        ' Get the Header associated with this particular element
        Dim aHeaderUIElement As UltraWinGrid.HeaderUIElement = CType(aCheckBoxUIElement.GetAncestor(GetType(UltraWinGrid.HeaderUIElement)), UltraWinGrid.HeaderUIElement)
        Dim aHeader As Infragistics.Win.UltraWinGrid.ColumnHeader = CType(aHeaderUIElement.GetContext(GetType(Infragistics.Win.UltraWinGrid.ColumnHeader)), Infragistics.Win.UltraWinGrid.ColumnHeader)

        ' Set the Tag on the Header to the new CheckState
        aHeader.Tag = aCheckBoxUIElement.CheckState

        ' So that we can apply various changes only to the relevant Rows collection that the header belongs to
        Dim hRows As UltraWinGrid.RowsCollection = CType(aHeaderUIElement.GetContext(GetType(UltraWinGrid.RowsCollection)), UltraWinGrid.RowsCollection)

        ' Raise an event so the programmer can do something when the CheckState changes
        RaiseEvent HeaderCheckBoxClicked(Me, New HeaderCheckBoxEventArgs(aHeader, aCheckBoxUIElement.CheckState, hRows))
    End Sub

    ' EventArgs used for the HeaderCheckBoxClicked event. This event has to pass in the CheckState and the Header
    Public Class HeaderCheckBoxEventArgs
        Inherits EventArgs

        Public Sub New(ByVal Header As Infragistics.Win.UltraWinGrid.ColumnHeader, ByVal CheckState As CheckState, ByRef Rows As UltraWinGrid.RowsCollection)
            mvarHeader = Header
            mvarCheckState = CheckState
            mvarRowsCollection = Rows
        End Sub

        Private mvarRowsCollection As UltraWinGrid.RowsCollection
        Private mvarHeader As Infragistics.Win.UltraWinGrid.ColumnHeader
        Private mvarCheckState As CheckState

        ' Expose the rows collection for the specific row island that the header belongs to
        Public ReadOnly Property Rows() As UltraWinGrid.RowsCollection
            Get
                Return mvarRowsCollection
            End Get
        End Property

        Public ReadOnly Property Header() As Infragistics.Win.UltraWinGrid.ColumnHeader
            Get
                Return mvarHeader
            End Get
        End Property

        Public Property CheckState() As CheckState
            Get
                Return mvarCheckState
            End Get
            Set(ByVal Value As CheckState)
                mvarCheckState = Value
            End Set
        End Property
    End Class
End Class